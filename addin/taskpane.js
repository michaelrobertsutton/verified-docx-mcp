// Pane logic (https://github.com/michaelrobertsutton/JennyStack/issues/106).
// Plain script, no framework, no bundler. Runs once Office.onReady fires.
//
// Two independent things happen on load:
//   1. WP-1's read-only report (runAndRender/postReport below, unchanged):
//      reads WordApi support, document URL/hash, and comments, and POSTs
//      the result to the bridge's /report route for the lead's manual
//      runbook (docs/live-mode.md).
//   2. WP-2's WSS ops channel (connectOpsSocket onward): connects to the
//      bridge's live/bridge.py /ops endpoint, sends `hello`, heartbeats
//      every 5s, and applies every op the MCP server sends via
//      dispatchOp -- this is what actually writes to the document, always
//      reading back afterward (pre/post body hashes) so the caller never
//      has to trust an unconfirmed write. See src/verified_docx_mcp/live/
//      protocol.py's module docstring for the wire shape and the exact
//      Word JS calls each op is documented against.

/* global Office, Word, WebSocket, console */

let lastReport = null;

// issue #84: pane lifecycle diagnostics. A pane that Word reloads or kills
// leaves no trace on its own, so every load keeps a small ring buffer of
// errors and lifecycle events in sessionStorage; the NEXT load reads it back
// as `previous_events` in its boot report, which is how a reload cause
// becomes visible. Nothing in this block may ever throw into the pane.
const PANE_EVENTS_KEY = "vdm_pane_events";
const BOOT_SEQ_KEY = "vdm_boot_seq";
const PANE_EVENTS_MAX = 50;
let REPORT_TIMEOUT_MS = 15000;
const SCRIPT_START_MS = Date.now();
const bootTimings = { office_ready: null, report_build: null, ws_connect: null };
let paneEvents = [];
let wsConnectStartMs = null;

function safeStorage(name) {
  try {
    return typeof globalThis[name] === "undefined" ? null : globalThis[name];
  } catch (err) {
    return null; // some hosts throw on mere access
  }
}

function persistPaneEvents() {
  try {
    const store = safeStorage("sessionStorage");
    if (store) store.setItem(PANE_EVENTS_KEY, JSON.stringify(paneEvents));
  } catch (err) {
    // sessionStorage unavailable or full: diagnostics are best-effort
  }
}

function recordPaneEvent(kind, message, stack) {
  try {
    paneEvents.push({
      t: new Date().toISOString(),
      kind,
      message: String(message === undefined || message === null ? "" : message).slice(0, 400),
      stack: String(stack || "").slice(0, 400),
    });
    if (paneEvents.length > PANE_EVENTS_MAX) paneEvents = paneEvents.slice(-PANE_EVENTS_MAX);
    persistPaneEvents();
  } catch (err) {
    // never let diagnostics break the pane
  }
}

// Read (and clear) the previous load's events BEFORE any listener can
// overwrite the key, and bump the per-profile boot counter.
function takePreviousEvents() {
  try {
    const store = safeStorage("sessionStorage");
    if (!store) return [];
    const raw = store.getItem(PANE_EVENTS_KEY);
    store.removeItem(PANE_EVENTS_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed.slice(-PANE_EVENTS_MAX) : [];
  } catch (err) {
    return [];
  }
}

function nextBootSeq() {
  try {
    const store = safeStorage("localStorage");
    if (!store) return null;
    const n = (parseInt(store.getItem(BOOT_SEQ_KEY), 10) || 0) + 1;
    store.setItem(BOOT_SEQ_KEY, String(n));
    return n;
  } catch (err) {
    return null;
  }
}

const PREVIOUS_PANE_EVENTS = takePreviousEvents();
const BOOT_SEQ = nextBootSeq();

function installPaneDiagnostics() {
  try {
    if (typeof window !== "undefined" && window && typeof window.addEventListener === "function") {
      window.addEventListener("error", (e) =>
        recordPaneEvent("error", (e && e.message) || "error", e && e.error && e.error.stack)
      );
      window.addEventListener("unhandledrejection", (e) => {
        const reason = e && e.reason;
        recordPaneEvent("unhandledrejection", (reason && reason.message) || String(reason), reason && reason.stack);
      });
      window.addEventListener("pagehide", () => recordPaneEvent("pagehide", ""));
      window.addEventListener("beforeunload", () => recordPaneEvent("beforeunload", ""));
    }
    if (typeof document !== "undefined" && document && typeof document.addEventListener === "function") {
      document.addEventListener("visibilitychange", () =>
        recordPaneEvent("visibilitychange", document.visibilityState || "")
      );
    }
    recordPaneEvent("boot", `boot_seq=${BOOT_SEQ}`);
  } catch (err) {
    // diagnostics are best-effort
  }
}
installPaneDiagnostics();

function bootDiagnostics() {
  return {
    boot_seq: BOOT_SEQ,
    instance_id: PANE_INSTANCE_ID,
    previous_events: PREVIOUS_PANE_EVENTS,
    boot_ms: { ...bootTimings },
  };
}

function withTimeout(promise, ms, label) {
  let timer;
  const timeout = new Promise((_, reject) => {
    timer = setTimeout(() => reject(new Error(`${label} timed out after ${ms} ms`)), ms);
  });
  return Promise.race([promise, timeout]).finally(() => clearTimeout(timer));
}

function setStatus(text) {
  document.getElementById("status").textContent = text;
}

function setOutput(obj) {
  const text = JSON.stringify(obj, null, 2);
  document.getElementById("output").textContent = text;
  return text;
}

async function sha256Hex(text) {
  const encoder = new TextEncoder();
  const data = encoder.encode(text);
  const digest = await crypto.subtle.digest("SHA-256", data);
  return Array.from(new Uint8Array(digest))
    .map((b) => b.toString(16).padStart(2, "0"))
    .join("");
}

// Requirement-set support: 1.4 is the WP-1 gate (docs/live-mode.md
// "What WP-1 must record"); 1.5/1.6 are reported for information only --
// nothing in this spike or the plan requires them.
function requirementSets() {
  const sets = {};
  for (const version of ["1.4", "1.5", "1.6"]) {
    try {
      sets[version] = Office.context.requirements.isSetSupported("WordApi", version);
    } catch (err) {
      sets[version] = `error: ${err && err.message ? err.message : String(err)}`;
    }
  }
  return sets;
}

// Two syncs regardless of comment count: (1) the collection with each
// comment's scalar fields, (2) the per-comment extras the caller asked for.
// options.ids       -- only these comment ids (others get no extras loaded)
// options.anchors   -- also load each comment's anchor text (getRange().text);
//                      default true, the expensive part on big docs
// options.replies   -- also load each comment's replies; default false
async function collectComments(context, options) {
  const opts = options || {};
  const wantAnchors = opts.anchors !== false;
  const wantReplies = !!opts.replies;
  const idFilter = Array.isArray(opts.ids) ? new Set(opts.ids) : null;

  const comments = context.document.body.getComments();
  comments.load("items/id,items/content,items/authorName,items/creationDate,items/resolved");
  await context.sync();

  // issue #39: counts describe the WHOLE collection, never the ids-filtered
  // subset, so a caller can compare them with what it received.
  const counts = {
    total: comments.items.length,
    open: comments.items.filter((c) => !c.resolved).length,
  };
  const picked = idFilter ? comments.items.filter((c) => idFilter.has(c.id)) : comments.items;
  const extras = picked.map((comment) => {
    let range = null;
    if (wantAnchors) {
      range = comment.getRange();
      range.load("text");
      range.paragraphs.load("items/text");
    }
    let replies = null;
    if (wantReplies) {
      replies = comment.replies; // CommentReplyCollection is a property, not a method
      replies.load("items/id,items/content,items/authorName,items/creationDate");
    }
    return { comment, range, replies };
  });
  if (wantAnchors || wantReplies) await context.sync();

  const items = extras.map(({ comment, range, replies }) => {
    const out = {
      id: comment.id,
      content: comment.content,
      authorName: comment.authorName,
      creationDate: comment.creationDate,
      resolved: comment.resolved,
    };
    if (wantAnchors) {
      out.anchorText = range.text;
      out.anchorParagraphText = range.paragraphs.items.map(p => p.text).join("\n");
    }
    if (wantReplies) {
      out.replies = replies.items.map((r) => ({
        id: r.id,
        content: r.content,
        authorName: r.authorName,
        creationDate: r.creationDate,
      }));
    }
    return out;
  });
  return { items, counts, observedAt: new Date().toISOString() };
}

async function buildReport() {
  return measuredWordRun(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();

    const bodyText = body.text || "";
    const bodyHash = await sha256Hex(bodyText);
    const { items: comments } = await collectComments(context);

    return {
      wp: "issue-106-wp1",
      generated_at: new Date().toISOString(),
      host: Office.context.host,
      platform: Office.context.platform,
      requirementSets: requirementSets(),
      documentUrl: Office.context.document.url,
      bodyTextLength: bodyText.length,
      bodyTextSha256: bodyHash,
      comments,
    };
  });
}

async function runAndRender() {
  setStatus("Reading document…");
  const buildStart = Date.now();
  try {
    // issue #84: bounded, so a hung comments/body read still ends in a report
    // (the ops socket starts first and independently of this).
    const report = await withTimeout(buildReport(), REPORT_TIMEOUT_MS, "report build");
    bootTimings.report_build = Date.now() - buildStart;
    Object.assign(report, bootDiagnostics());
    lastReport = report;
    setOutput(report);
    renderSummary(report);
    document.getElementById("copy-json").disabled = false;
    await postReport(report);
    setStatus(
      `OK — host=${report.host} platform=${report.platform} ` +
        `WordApi1.4=${report.requirementSets["1.4"]} comments=${report.comments.length}`
    );
  } catch (err) {
    bootTimings.report_build = Date.now() - buildStart;
    lastReport = null;
    const message = err && err.message ? err.message : String(err);
    recordPaneEvent("load_error", message, err && err.stack);
    try {
      document.getElementById("copy-json").disabled = true;
      setOutput({ error: message, debugInfo: err && err.debugInfo ? err.debugInfo : undefined });
      setStatus(`ERROR: ${message}`);
    } catch (uiErr) {
      // the minimal report below matters more than the pane's own text
    }
    // issue #84: still tell the bridge this load happened, and why it failed.
    let documentUrl = null;
    try {
      documentUrl = Office.context.document.url || null;
    } catch (urlErr) {
      // leave null
    }
    await postReport({
      generated_at: new Date().toISOString(),
      documentUrl,
      load_error: message,
      ...bootDiagnostics(),
    });
  }
}

async function postReport(report) {
  // WP-1 convenience: hand the report to the bridge so nothing has to be
  // copied out of the pane by hand. Failure is reported, never fatal --
  // the JSON stays visible below and "Copy JSON" still works.
  const resultEl = document.getElementById("ping-result");
  try {
    const resp = await fetch("https://localhost:53135/report", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(report),
    });
    const body = await resp.json();
    resultEl.textContent = `report posted to bridge (HTTP ${resp.status}): ${JSON.stringify(body)}`;
  } catch (err) {
    resultEl.textContent = `report POST failed: ${err && err.message ? err.message : String(err)}`;
  }
}

async function copyJson() {
  if (!lastReport) return;
  const text = JSON.stringify(lastReport, null, 2);
  try {
    await navigator.clipboard.writeText(text);
    setStatus("Copied JSON to clipboard. Paste it into a file for scripts/compare_comment_ids.py.");
  } catch (err) {
    // Clipboard API can be blocked in some sideload contexts; fall back to
    // a visible textarea-free path: the JSON is already in the <pre> block,
    // so tell the lead to select-copy it manually.
    setStatus(
      `Clipboard write failed (${err && err.message ? err.message : err}); ` +
        "select the JSON below and copy it manually."
    );
  }
}

async function pingBridge() {
  const resultEl = document.getElementById("ping-result");
  resultEl.textContent = "Pinging https://localhost:53135/ping ...";
  try {
    const resp = await fetch("https://localhost:53135/ping");
    const body = await resp.json();
    resultEl.textContent = `ping OK (HTTP ${resp.status}): ${JSON.stringify(body)}`;
  } catch (err) {
    resultEl.textContent = `ping FAILED: ${err && err.message ? err.message : String(err)}`;
  }
}

// ---------------------------------------------------------------------------
// WP-2: WSS ops channel (https://github.com/michaelrobertsutton/JennyStack/
// issues/106). Connects to live/bridge.py's /ops endpoint, sends `hello`,
// heartbeats every 5s, and dispatches every op live/protocol.py defines,
// always reading the document back after a mutation (pre/post body
// SHA-256) so a caller never has to trust an unconfirmed write.
// ---------------------------------------------------------------------------

// live/bridge.py's DEFAULT_OPS_PORT = DEFAULT_PORT + 1: the WSS channel is
// a SEPARATE port from the static pane server (53135) -- see that module's
// docstring for why (a synchronous http.server listener and an asyncio
// websockets listener cannot share one accept() loop cleanly).
const OPS_PORT = 53136;
const HEARTBEAT_INTERVAL_MS = 5000;
const RECONNECT_MIN_MS = 1000;
const RECONNECT_MAX_MS = 30000;
const OP_LOG_LIMIT = 20;

// issue #22 B2: sent in `hello`; see the capabilities comment at the send
// site (connectOpsSocket's onopen) for what this defends against.
// issue #27: "cell_edit" = the cell_get/cell_set ops below (live table-cell
// edits). Reported only by builds that implement them, so a server never
// sends cell_set to a pane that would answer "unknown op".
// issue #31: "comments_by_id" = comments_list accepts `ids` / `include_anchor`.
// issue #33: "body_ooxml" = the body_ooxml op (live reads for read_document /
// find_sections / list_tables).
// issue #75: "cell_multiline" = cell_get returns per-paragraph text, cell_set
// writes multi-paragraph/line-break content correctly and self-checks/restores.
// issue #34: "table_edit" = the table_get/table_insert/cells_set ops.
// issue #39: "comment_counts" = comments_list reports counts/scope/observed_at.
// issue #77: "paragraph_insert" = the paragraph_insert op (styled, optionally tracked paragraphs).
// issue #80: "comment_delete" = the comment_delete op (removes a whole comment thread).
// issue #82: "comment_loss_ids" = replace honors payload.allow_comment_loss_ids (per-thread consent).
// issue #39: "comments_ooxml_crosscheck" = the unfiltered comments_list adds an `ooxml` open-thread count.
const PANE_CAPABILITIES = ["paragraph_positions_v1", "cell_paragraph_merge_v1", "cell_font_inheritance_v1", "heartbeat_liveness_v1", "approved_cell_stream_v1", "anchored_text_v1", "formatted_replacements_v1", "approved_changes_v1", "section_locks", "shared_queue", "row_scope", "cell_edit", "cell_multiline", "comments_by_id", "table_edit", "body_ooxml", "live_revisions", "comment_loss_guard", "replacement_formatting", "format_readback", "textboxes", "delete_paragraph", "shape_guard", "comment_counts", "autoopen", "paragraph_insert", "comment_delete", "comment_loss_ids", "table_delete", "paragraphs_unlist", "revisions_scoped", "comments_ooxml_crosscheck", "body_hash_v1", "compact_evidence_v1"];

// issue #39: per-load id, sent in `hello`, so the server can say which pane
// instance answered. Falls back when crypto.randomUUID is unavailable.
const PANE_INSTANCE_ID = (() => {
  try {
    if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") return crypto.randomUUID();
  } catch (err) {
    // fall through to the non-crypto id below
  }
  return `pane-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
})();

// issue #33: the largest body_ooxml reply (UTF-8 bytes of its JSON) the pane
// will send. Kept below the bridge's 64 MiB websocket max_size so an
// oversized document gets an explicit `too_large` refusal instead of a frame
// the bridge rejects by dropping the whole socket.
const MAX_BODY_OOXML_BYTES = 48 * 1024 * 1024;
const FLAT_OPC_NS = "http://schemas.microsoft.com/office/2006/xmlPackage";

let opsSocket = null;
let heartbeatTimer = null;
let reconnectTimer = null;
let reconnectDelayMs = RECONNECT_MIN_MS;
let opLog = [];

function setWsStatus(text) {
  const el = document.getElementById("ws-status");
  if (!el) return;
  el.textContent = `Ops channel: ${text}`;
  // Color the line by state so it reads at a glance in Word's dark theme.
  el.dataset.state = text === "connected" ? "connected" : /disconnected|failed/.test(text) ? "down" : "pending";
}

function renderSummary(report) {
  // Plain-language card for the writer; the raw JSON stays under "Details".
  const nameEl = document.getElementById("doc-name");
  const capEl = document.getElementById("cap-line");
  const commentsEl = document.getElementById("comments-line");
  const warnEl = document.getElementById("doc-warning");
  if (!nameEl || !capEl || !commentsEl || !warnEl) return;
  const url = report.documentUrl || "";
  const name = url ? url.split(/[\\/]/).pop() : "";
  nameEl.textContent = name || "(unsaved document)";
  warnEl.classList.toggle("show", !url);
  const ok14 = !!(report.requirementSets && report.requirementSets["1.4"]);
  capEl.textContent = ok14 ? "supported" : "NOT supported";
  capEl.className = "v " + (ok14 ? "ok" : "bad");
  const open = (report.comments || []).filter((c) => !c.resolved).length;
  updateCommentsCounter(open, report.generated_at);
}

// issue #39: the counter used to be set once at pane load and never again, so
// it could disagree with a later list_open_items with no way to tell why. It
// now refreshes on every full comments_list and shows when it was read.
function updateCommentsCounter(open, observedAt) {
  const el = document.getElementById("comments-line");
  if (!el) return;
  const at = observedAt ? ` (as of ${String(observedAt).slice(11, 19)})` : "";
  el.textContent = `${open}${at}`;
}

function renderActivity() {
  const ul = document.getElementById("activity");
  if (!ul) return;
  if (!opLog.length) {
    ul.innerHTML = '<li class="t">No edits received yet.</li>';
    return;
  }
  ul.innerHTML = "";
  opLog.slice(0, 10).forEach((e) => {
    const li = document.createElement("li");
    const t = document.createElement("span");
    t.className = "t";
    t.textContent = e.time.slice(11, 19);
    const label = document.createElement("span");
    label.className = e.ok ? "ok" : "bad";
    label.textContent = (e.ok ? "applied " : "refused ") + e.op;
    li.appendChild(t);
    li.appendChild(label);
    if (e.detail) {
      const d = document.createElement("span");
      d.className = "t";
      d.textContent = "  " + String(e.detail).slice(0, 80);
      li.appendChild(d);
    }
    ul.appendChild(li);
  });
}

function logOp(op, ok, detail) {
  opLog.unshift({ time: new Date().toISOString(), op, ok, detail: detail || "" });
  opLog = opLog.slice(0, OP_LOG_LIMIT);
  const el = document.getElementById("op-log");
  if (!el) return;
  el.textContent = opLog
    .map((e) => `${e.time}  ${e.ok ? "OK  " : "FAIL"}  ${e.op}${e.detail ? "  " + e.detail : ""}`)
    .join("\n");
  renderActivity();
}

async function currentBodyHash() {
  return Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();
    return sha256Hex(body.text || "");
  });
}

function refusalError(message, code) {
  // Thrown by an op handler below to make the WSS reply ok:false with a
  // typed code, matching live/protocol.py's documented refusal shape --
  // never an uncaught exception reaching the pane's onmessage handler.
  // Defaults to LIVE_OP_FAILED; a search-and-count-gate refusal (issue
  // #106 WP-4) passes "zero_match"/"match_count_mismatch" instead, the
  // same two names live/protocol.py's OP_ERROR_* constants use, so
  // server.py's live comment tools can map them onto ZERO_MATCH/
  // MATCH_COUNT_MISMATCH the same way file mode's locate() ladder does.
  const err = new Error(message);
  err.code = code || "LIVE_OP_FAILED";
  return err;
}

const paneNow = () => typeof performance !== "undefined" ? performance.now() : Date.now();
let heartbeatLiveness = false;
let lastIdleHashMs = -Infinity;
let idleHashPending = false;
const IDLE_HASH_INTERVAL_MS = 30000;

function connectOpsSocket() {
  if (reconnectTimer) {
    clearTimeout(reconnectTimer);
    reconnectTimer = null;
  }
  setWsStatus("connecting…");
  if (wsConnectStartMs === null) wsConnectStartMs = Date.now();
  let socket;
  try {
    socket = new WebSocket(`wss://localhost:${OPS_PORT}/ops`);
  } catch (err) {
    setWsStatus(`failed to open (${err && err.message ? err.message : err}); retrying`);
    scheduleReconnect();
    return;
  }
  opsSocket = socket;
  heartbeatLiveness = false;
  lastIdleHashMs = -Infinity;

  socket.onopen = async () => {
    reconnectDelayMs = RECONNECT_MIN_MS;
    try {
      const hash = await enqueuePaneOperation(async () => {
        const sampled = await currentBodyHash();
        if (socket === opsSocket) {
          paneEpoch = crypto.randomUUID();
          revisionHandles.clear();
        }
        return sampled;
      });
      if (socket !== opsSocket || socket.readyState !== WebSocket.OPEN) return;
      socket.send(
        JSON.stringify({
          type: "hello",
          documentUrl: Office.context.document.url,
          host: Office.context.host,
          platform: Office.context.platform,
          requirementSets: requirementSets(),
          bodySha256: hash,
          // issue #22 B2: op-level feature names THIS pane build
          // implements, independent of WordApi version support
          // (requirementSets above) -- a pane can run on a WordApi
          // version new enough for row_scope's own Office JS calls while
          // still running an OLDER BUILD of this exact file that never
          // learned the rowAnchor wire field. The server's
          // live/write_mode.py require_capability refuses BEFORE sending
          // an op that depends on a capability not listed here, rather
          // than risk this pane silently ignoring an unknown payload key
          // and running an unscoped op instead.
          capabilities: PANE_CAPABILITIES,
          instanceId: PANE_INSTANCE_ID,
        })
      );
      setWsStatus("connected");
      if (bootTimings.ws_connect === null) bootTimings.ws_connect = Date.now() - wsConnectStartMs;
      startHeartbeat();
    } catch (err) {
      setWsStatus(`hello failed (${err && err.message ? err.message : err})`);
    }
  };

  socket.onmessage = (event) => {
    let message;
    try {
      message = JSON.parse(event.data);
    } catch (err) {
      return; // malformed message from the bridge; ignore rather than crash the pane
    }
    if (socket !== opsSocket) return;
    if (message.type === "hello_ack") {
      heartbeatLiveness = (message.capabilities || []).includes("heartbeat_liveness_v1");
      return;
    }
    if (message.type !== "request") return;
    const requestId = message.request_id;
    const op = message.op;
    const payload = message.payload || {};
    enqueuePaneOperation(async () => {
      if (socket.readyState !== WebSocket.OPEN) return;
      try {
        const result = await measuredDispatch(op, payload);
        logOp(payload.clientId ? `${op} [${payload.clientId.slice(0, 8)}]` : op, true, summarizeResult(op, result));
        socket.send(JSON.stringify({ type: "reply", request_id: requestId, ok: true, result }));
      } catch (err) {
        const code = err && err.code ? err.code : "LIVE_OP_FAILED";
        const msg = err && err.message ? err.message : String(err);
        logOp(payload.clientId ? `${op} [${payload.clientId.slice(0, 8)}]` : op, false, msg);
        socket.send(
          JSON.stringify({ type: "reply", request_id: requestId, ok: false, error: { code, message: msg } })
        );
      }
    }).catch(() => {}); // closed socket: result is unknown; never replay
  };

  socket.onclose = () => {
    if (socket !== opsSocket) return;
    stopHeartbeat();
    setWsStatus(`disconnected — reconnecting in ${Math.round(reconnectDelayMs / 1000)}s`);
    scheduleReconnect();
  };

  socket.onerror = () => {
    // The WebSocket spec fires onclose right after onerror; the reconnect
    // loop lives there, not here.
  };
}

function scheduleReconnect() {
  if (reconnectTimer) return;
  reconnectTimer = setTimeout(() => {
    reconnectTimer = null;
    connectOpsSocket();
  }, reconnectDelayMs);
  reconnectDelayMs = Math.min(reconnectDelayMs * 2, RECONNECT_MAX_MS);
}

// Liveness never waits for Word. Hash work shares the operation queue and is
// coalesced while busy; no hash is represented as a new document read.
function heartbeatTick(socket) {
  if (socket !== opsSocket || socket.readyState !== WebSocket.OPEN) return;
  if (heartbeatLiveness) {
    socket.send(JSON.stringify({type:"heartbeat", documentUrl:Office.context.document.url,
      digestState:"unavailable"}));
    if (paneQueueDepth || idleHashPending || paneNow()-lastIdleHashMs < IDLE_HASH_INTERVAL_MS) return;
  } else if (idleHashPending) return;
  idleHashPending = true;
  enqueuePaneOperation(async () => {
    if (socket !== opsSocket || socket.readyState !== WebSocket.OPEN) return;
    const started = paneNow();
    try {
      const hash = await currentBodyHash();
      if (socket !== opsSocket || socket.readyState !== WebSocket.OPEN) return;
      lastIdleHashMs = paneNow();
      socket.send(JSON.stringify({type:"heartbeat", documentUrl:Office.context.document.url,
        bodySha256:hash, ...(heartbeatLiveness ? {digestState:"sampled",
          metrics:{hash_requests:1, sync_count:1, elapsed_ms:paneNow()-started}} : {})}));
    } catch (err) {
      // The independent liveness timer continues even if Word is blocked.
    }
  }).catch(() => {}).finally(() => {idleHashPending = false;});
}
function startHeartbeat() {
  stopHeartbeat();
  const socket = opsSocket;
  heartbeatTimer = setInterval(() => heartbeatTick(socket), HEARTBEAT_INTERVAL_MS);
}

function stopHeartbeat() {
  if (heartbeatTimer) {
    clearInterval(heartbeatTimer);
    heartbeatTimer = null;
  }
}

function summarizeResult(op, result) {
  if (op === "search") return `${(result.matches || []).length} match(es)`;
  if (op === "replace" || op === "format") return `match_count=${result.match_count}`;
  if (op === "comment_add") return `comment_id=${result.comment_id}`;
  if (op === "cell_get") return `${(result.text || "").length} char(s)`;
  if (op === "cell_set") return `applied=${result.applied}`;
  if (op === "table_get") return `${(result.rows || []).length} row(s)`;
  if (op === "table_insert") return `table_index=${result.table_index}`;
  if (op === "cells_set") return `applied=${result.applied}`;
  if (op === "paragraph_insert") return `inserted=${(result.inserted || []).length}`;
  if (op === "comment_delete") return `deleted=${result.deleted} still_present=${result.still_present}`;
  return "";
}

const SHAPE_GUARDED_OPS = new Set(["approved_cell_set", "replace", "format", "cell_set", "cells_set",
  "table_insert", "cell_paragraph_merge", "paragraph_delete", "paragraph_insert", "revisions_accept", "revisions_reject",
  "comment_add", "comment_reply", "comment_resolve", "comment_delete", "save", "table_delete", "paragraphs_unlist"]);

// One queue across reconnects. A timed-out write can still be running in Word;
// later commands must not overlap it, including commands on a new socket.
let paneOperationQueue = Promise.resolve();
let paneQueueDepth = 0;
function enqueuePaneOperation(operation) {
  paneQueueDepth++;
  const result = paneOperationQueue.then(operation).finally(() => {paneQueueDepth--;});
  paneOperationQueue = result.catch(() => {});
  return result;
}
const READ_ONLY_OPS = new Set(["ping", "describe", "body_ooxml", "body_hash", "search", "comments_list",
  "cell_get", "table_get", "shapes_list", "textboxes_list", "textboxes_read", "scope_describe",
  "revisions_list", "autoopen_get", "sections_list", "execution_snapshot"]);

async function dispatchOp(op, payload) {
  if (op === "save" && payload.executionSections)
    await measuredWordRun(context => sectionGuard(context, payload));
  const sectionScoped = !!((payload.ownSections && payload.ownSections.length) ||
    (payload.executionSections && payload.executionSections.length));
  if (!READ_ONLY_OPS.has(op) && payload.clientId && !sectionScoped) {
    // issue #107: save persists, it neither reads nor changes the body, so it needs no baseline and must
    // not be refused "because someone is typing" (the broker no longer sends one). A hash that IS
    // supplied is still honoured, like any other write.
    if (!payload.expectedBodySha256 && op !== "save")
      throw refusalError("Shared write requires a body baseline", "LIVE_STALE");
    if (payload.expectedBodySha256)
      await measuredWordRun(async context => requireFreshBody(context, payload.expectedBodySha256));
  }
  if (op === "approved_cell_set") return dispatchOpUnchecked(op, payload); // verifies scoped shapes inside its boundary snapshots
  if (!SHAPE_GUARDED_OPS.has(op)) return dispatchOpUnchecked(op, payload);
  const before = await (payload.executionSections ? opExecutionShapes(payload) : opShapesList());
  // Accepting/rejecting tracked deletions can remove whole anchor paragraphs.
  if (op.startsWith("revisions_") && before.shapes.length)
    throw refusalError("Revision changes with anchored shapes require manual review", "ANCHORED_SHAPES");
  const result = await dispatchOpUnchecked(op, payload);
  let after;
  try { after = await (payload.executionSections ? opExecutionShapes(payload) : opShapesList());}
  catch (err) {
    throw refusalError(`Shape read-back failed after ${op}; the edit may have applied; ` +
      `no rollback attempted; shapes_before=${before.shapes.length}; ${err.message || err}`, "VERIFICATION_FAILED");
  }
  const identity = shapes => shapes.map(({anchor_paragraph, ...shape}) => JSON.stringify(shape)).sort();
  if (JSON.stringify(identity(before.shapes)) !== JSON.stringify(identity(after.shapes)))
    throw refusalError(`Shape verification failed after ${op}; no rollback attempted; ` +
      JSON.stringify({shapes_before:before.shapes.length, shapes_after:after.shapes.length}), "VERIFICATION_FAILED");
  return {...result, shapes_before:before.shapes.length, shapes_after:after.shapes.length};
}

async function dispatchOpUnchecked(op, payload) {
  switch (op) {
    case "shapes_list":
      return opShapesList();
    case "execution_snapshot":
      return measuredWordRun(context => executionSnapshot(context, payload.section_slugs));
    case "sections_list":
      return opSectionsList();
    case "ping":
      return opPing();
    case "describe":
      return opDescribe();
    case "autoopen_get":
      return opAutoopenGet();
    case "autoopen_set":
      return opAutoopenSet(payload);
    case "body_ooxml":
      return opBodyOoxml();
    case "body_hash":
      return opBodyHash();
    case "search":
      return opSearch(payload);
    case "replace":
      return opReplace(payload);
    case "format":
      return opFormat(payload);
    case "comments_list":
      return opCommentsList(payload);
    case "comment_add":
      return opCommentAdd(payload);
    case "comment_reply":
      return opCommentReply(payload);
    case "comment_resolve":
      return opCommentResolve(payload);
    case "comment_delete":
      return opCommentDelete(payload);
    case "cell_get":
      return opCellGet(payload);
    case "approved_cell_set":
      return opApprovedCellSet(payload);
    case "cell_set":
      return opCellSet(payload);
    case "table_get":
      return opTableGet(payload);
    case "table_insert":
      return opTableInsert(payload);
    case "cells_set":
      return opCellsSet(payload);
    case "cell_paragraph_merge":
      return opCellParagraphMerge(payload);
    case "paragraph_delete":
      return opParagraphDelete(payload);
    case "paragraph_insert":
      return opParagraphInsert(payload);
    case "paragraphs_unlist":
      return opParagraphsUnlist(payload);
    case "table_delete":
      return opTableDelete(payload);
    case "textboxes_list":
      return opTextboxesList();
    case "textboxes_read":
      return opTextboxesRead(payload);
    case "scope_describe":
      return opScopeDescribe(payload);
    case "revisions_list":
      return opRevisionsList();
    case "revisions_accept":
      return opRevisionsMutate(payload, "accept");
    case "revisions_reject":
      return opRevisionsMutate(payload, "reject");
    case "save":
      return opSave();
    default:
      throw refusalError(`unknown op ${op}`);
  }
}

// -- ping / describe --------------------------------------------------

async function opPing() {
  // Round-trips the JS engine through Word.run -- proves the pane is
  // alive, not just the socket (live/protocol.py's `ping` docstring).
  return measuredWordRun(async (context) => {
    context.document.body.load("text");
    await context.sync();
    return {};
  });
}

// issue #66: Office opens the add-in's task pane with a document that carries this
// document setting (manifest TaskpaneId Office.AutoShowTaskpaneWithDocument). The
// setting travels in the .docx; the pane must already be open once to write it.
const AUTOOPEN_SETTING = "Office.AutoShowTaskpaneWithDocument";
function autoopenSupported() {
  try {
    return Office.context.requirements.isSetSupported("AddinCommands", "1.1");
  } catch (error) {
    return false;
  }
}
function opAutoopenGet() {
  return {
    enabled: Office.context.document.settings.get(AUTOOPEN_SETTING) === true,
    supported: autoopenSupported(),
    session_epoch: paneEpoch,
  };
}
async function opAutoopenSet(payload) {
  if (typeof payload.enabled !== "boolean") throw refusalError("enabled must be true or false", "INVALID_INPUT");
  if (!autoopenSupported()) throw refusalError("This Office host does not support auto-opening a task pane (AddinCommands 1.1)", "LIVE_CAPABILITY_MISSING");
  const settings = Office.context.document.settings;
  const before = settings.get(AUTOOPEN_SETTING) === true;
  settings.set(AUTOOPEN_SETTING, payload.enabled);
  await new Promise((resolve, reject) => settings.saveAsync(result =>
    result.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error(`settings.saveAsync failed: ${result.error && result.error.message}`))));
  const after = settings.get(AUTOOPEN_SETTING) === true;
  if (after !== payload.enabled) throw refusalError(`autoopen setting did not take (wanted ${payload.enabled}, read back ${after})`, "VERIFICATION_FAILED");
  return {enabled: after, was_enabled: before, supported: true, session_epoch: paneEpoch};
}

async function opDescribe() {
  return measuredWordRun(async (context) => {
    const body = context.document.body;
    body.load("text");
    context.document.load(["changeTrackingMode", "saved"]);
    await context.sync();
    const hash = await sha256Hex(body.text || "");
    return {
      session_epoch: paneEpoch,
      documentUrl: Office.context.document.url,
      bodySha256: hash,
      changeTrackingMode: String(context.document.changeTrackingMode),
      saved: context.document.saved,
    };
  });
}

// -- body_ooxml (issue #33) ---------------------------------------------

// Empty every pkg:binaryData payload (images and other binary parts) and
// return {ooxml, strippedParts}. The server's projection needs a part's
// name and relationships, never its bytes. Namespace-aware (the Flat OPC
// namespace, not a literal "pkg:" prefix), so an alternate serialization
// cannot slip a payload past it. Throws if the XML does not parse, so a
// malformed export is an error rather than an unstripped multi-MiB reply.
function stripBinaryParts(ooxml) {
  const doc = new DOMParser().parseFromString(ooxml, "application/xml");
  if (doc.getElementsByTagName("parsererror").length > 0) {
    throw refusalError("getOoxml returned XML that did not parse");
  }
  const strippedParts = [];
  const nodes = doc.getElementsByTagNameNS(FLAT_OPC_NS, "binaryData");
  for (let i = 0; i < nodes.length; i += 1) {
    const node = nodes[i];
    const part = node.parentNode;
    strippedParts.push(part && part.getAttributeNS(FLAT_OPC_NS, "name"));
    node.textContent = "";
  }
  return { ooxml: new XMLSerializer().serializeToString(doc), strippedParts };
}

async function opBodyOoxml() {
  return measuredWordRun(async (context) => {
    const body = context.document.body;
    const ooxmlResult = body.getOoxml();
    body.load("text");
    await context.sync();
    const hash = await sha256Hex(body.text || "");
    const { ooxml, strippedParts } = stripBinaryParts(ooxmlResult.value);
    const result = {
      ooxml,
      bodySha256: hash,
      documentUrl: Office.context.document.url,
      strippedParts,
    };
    const bytes = new TextEncoder().encode(JSON.stringify(result)).length;
    if (bytes > MAX_BODY_OOXML_BYTES) {
      throw refusalError(
        `body OOXML is ${bytes} bytes after stripping binaries, over the ${MAX_BODY_OOXML_BYTES}-byte limit`,
        "too_large"
      );
    }
    return result;
  });
}


// -- section locks (#47) ------------------------------------------------
// A section is what find_sections reports: a heading paragraph (outline level
// 1-9, outside tables) up to the next heading at any level. The broker sends
// `forbiddenSections` (slugs locked by OTHER clients) and `ownSections`
// ({slug, expectedSha256} locked by this client) with every shared write; the
// pane re-resolves them from the live headings and checks the real mutation
// target before anything is written. A slug that is missing or ambiguous
// fails closed.
function sectionSlug(text) {
  return (text || "").trim().toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "") || "section";
}
async function loadSectionMap(context, headingsOnly = false) {
  // A separate collection from body.paragraphs: ops (paragraph_delete, table_insert)
  // keep paragraph objects from their own load, and reloading that collection here
  // would invalidate them.
  const paragraphs = context.document.body.getRange("Whole").paragraphs;
  paragraphs.load(headingsOnly ? "items/outlineLevel,items/tableNestingLevel" : "items/text,items/outlineLevel,items/tableNestingLevel");
  await context.sync();
  const items = paragraphs.items;
  if (headingsOnly) {
    items.filter(p => p.tableNestingLevel === 0 && p.outlineLevel >= 1 && p.outlineLevel <= 9).forEach(p => p.load("text"));
    await context.sync();
  }
  const heads = [];
  items.forEach((p, i) => {
    const level = p.outlineLevel;
    if (p.tableNestingLevel === 0 && Number.isInteger(level) && level >= 1 && level <= 9)
      heads.push({ index: i, level, text: (p.text || "").trim() });
  });
  const seen = new Map();
  const sections = heads.map((h, n) => {
    const slug = sectionSlug(h.text);
    const ordinal = (seen.get(slug) || 0) + 1;
    seen.set(slug, ordinal);
    const end = n + 1 < heads.length ? heads[n + 1].index : items.length;
    return { slug, section_key: `${slug}-${ordinal}`, heading_text: h.text, level: h.level, start: h.index, end };
  });
  return { items, sections };
}
// Top-level body paragraphs omit table cells. End at the next heading's
// start (or body end), so intervening tables remain inside the section.
function sectionRange(context, items, section) {
  const end = section.end < items.length ? items[section.end].getRange("Start") : context.document.body.getRange("End");
  return items[section.start].getRange("Whole").expandTo(end);
}
function sectionSha(items, section) {
  return sha256Hex(items.slice(section.start, section.end).map(p => p.text || "").join("\n"));
}
async function opSectionsList() {
  return measuredWordRun(async (context) => {
    const { items, sections } = await loadSectionMap(context);
    const bodySha256 = await requireFreshBody(context);
    const counts = new Map();
    sections.forEach(s => counts.set(s.slug, (counts.get(s.slug) || 0) + 1));
    const listed = [];
    for (const s of sections) {
      listed.push({
        section_key: s.section_key, slug: s.slug, heading_text: s.heading_text, level: s.level,
        paragraph_count: s.end - s.start, slug_unique: counts.get(s.slug) === 1,
        sectionSha256: await sectionSha(items, s),
      });
    }
    return { sections: listed, bodySha256 };
  });
}
// #110 live_revision: the fresh body hash and a 12-hex digest per heading
// section in one pass, with no body transfer. sha12 is the first 12 hex of the
// same sectionSha256 that sections_list reports, so file and live digests agree.
async function opBodyHash() {
  return measuredWordRun(async (context) => {
    const { items, sections } = await loadSectionMap(context);
    const bodySha256 = await requireFreshBody(context);
    const digests = [];
    for (const s of sections) digests.push([s.section_key, (await sectionSha(items, s)).slice(0, 12)]);
    return { bodySha256, sections: digests };
  });
}
function refuseDocumentWide(payload, what) {
  if ((payload.forbiddenSections || []).length)
    throw refusalError(`${what} are document-wide and another client holds a section lock`, "LOCKED_BY_OTHER_CLIENT");
  if ((payload.ownSections || []).length)
    throw refusalError(`${what} are document-wide; release your section locks first`, "OUTSIDE_LOCKED_SECTION");
}
// Word.LocationRelation values. Fail closed: only a range wholly before/after (or merely
// adjacent to) a section is outside it; anything else, including "Unrelated", touches it.
const SECTION_OUTSIDE = new Set(["Before", "AdjacentBefore", "After", "AdjacentAfter"]);
const SECTION_WITHIN = new Set(["Inside", "InsideStart", "InsideEnd", "Equal"]);
async function sectionGuard(context, payload) {
  const forbiddenSlugs = payload.forbiddenSections || [];
  const execution = payload.executionSections || [];
  const own = payload.ownSections || [];
  let executionMap=null, scopeShapes=null;
  if (execution.length) {
    const current = await (payload.approvedCellStream ? executionXmlSnapshot : executionSnapshot)(context, execution.map(s => s.slug));
    executionMap=current.section_map;
    scopeShapes=current.shapes || null;
    if (JSON.stringify(current.sections) !== JSON.stringify(execution))
      throw refusalError("Protected section changed; nothing was written", "LIVE_STALE");
  }
  if (!forbiddenSlugs.length && !own.length && !execution.length) return { active: false, async check() {} };
  if (payload.scope && payload.scope !== "body")
    throw refusalError("Section locks do not cover text-box scopes; release the locks first", "OUTSIDE_LOCKED_SECTION");
  const { items, sections } = executionMap || await loadSectionMap(context, execution.length > 0);
  const resolve = (slug) => {
    const hits = sections.filter(s => s.slug === slug);
    if (hits.length !== 1)
      throw refusalError(`Locked section ${JSON.stringify(slug)} ${hits.length ? "is ambiguous" : "was not found"}; ` +
        "its owner must release and re-lock", "LOCK_SCOPE_UNRESOLVED");
    return hits[0];
  };
  // Whole paragraphs, paragraph marks included: a paragraph's "End" point sits before its
  // mark, so a whole-paragraph target would not count as inside a section built from it.
  const rangeOf = s => sectionRange(context, items, s);
  const forbidden = forbiddenSlugs.map(slug => ({ slug, range: rangeOf(resolve(slug)) }));
  const mine = [];
  for (const o of [...own, ...execution]) {
    const section = resolve(o.slug);
    if (o.expectedSha256 && execution.length) {
      items.slice(section.start,section.end).forEach(p => p.load("text"));
      await context.sync();
    }
    if (o.expectedSha256 && (await sectionSha(items, section)) !== o.expectedSha256)
      throw refusalError(`Section ${JSON.stringify(o.slug)} changed since it was read; nothing was written. Re-read and retry.`, "stale");
    mine.push({ slug: o.slug, range: rangeOf(section) });
  }
  return {
    active: true, scopeShapes,
    async check(targetsOrThunk) {
      const targets = typeof targetsOrThunk === "function" ? targetsOrThunk() : targetsOrThunk;
      const forbiddenRels = targets.map(t => forbidden.map(f => t.compareLocationWith(f.range)));
      const ownRels = targets.map(t => mine.map(m => t.compareLocationWith(m.range)));
      try {
        await context.sync();
      } catch (err) {
        // Fail closed: an unverifiable target is never written.
        throw refusalError(`Could not verify the write target against section locks; nothing was written: ` +
          `${err && err.message ? err.message : err} ${JSON.stringify((err && err.debugInfo) || {})}`, "LOCK_SCOPE_UNRESOLVED");
      }
      targets.forEach((_, i) => {
        const hit = forbidden.findIndex((f, j) => !SECTION_OUTSIDE.has(forbiddenRels[i][j].value));
        if (hit >= 0)
          throw refusalError(`Section ${JSON.stringify(forbidden[hit].slug)} is locked by another client; nothing was written`,
            "LOCKED_BY_OTHER_CLIENT");
        if (mine.length && !mine.some((m, j) => SECTION_WITHIN.has(ownRels[i][j].value)))
          throw refusalError("The target is outside the sections you have locked; nothing was written", "OUTSIDE_LOCKED_SECTION");
      });
    },
  };
}


// A structured, scoped snapshot. Cell identities are coordinates within an exact
// table/section structure, never a global search for possibly repeated text.
async function executionSnapshot(context, slugs) {
  if (!Array.isArray(slugs) || !slugs.length || new Set(slugs).size !== slugs.length)
    throw refusalError("Unique protected section slugs are required", "INVALID_INPUT");
  const {items, sections} = await loadSectionMap(context, true);
  const chosen = slugs.map(slug => {
    const found = sections.filter(s => s.slug === slug);
    if (found.length !== 1) throw refusalError("Missing or ambiguous protected section", "LOCK_SCOPE_UNRESOLVED");
    return found[0];
  });
  chosen.forEach(s => items.slice(s.start,s.end).filter(p => p.tableNestingLevel === 0).forEach(p => p.load("text")));
  await context.sync();
  const table_addresses = [];
  const cell_paragraphs = {};
  const ranges = chosen.map(s => sectionRange(context, items, s));
  const tables = await loadTopLevelTables(context);
  const relations = tables.map(t => ranges.map(range => t.getRange().compareLocationWith(range)));
  await context.sync();
  const result = chosen.map(s => ({slug:s.slug, paragraphs:items.slice(s.start,s.end)
    .filter(p => p.tableNestingLevel === 0).map(p => p.text || ""), tables:[]}));
  for (let i=0; i<tables.length; i++) {
    const index = relations[i].findIndex(r => SECTION_WITHIN.has(r.value));
    if (index < 0) continue;
    const grid = await loadRowsAndCells(context, tables[i]);
    grid.forEach(row => row.forEach(cell => cell.body.paragraphs.load("items/text")));
    await context.sync();
    grid.forEach((row,r) => row.forEach((cell,c) => {cell_paragraphs[`${chosen[index].slug}|${result[index].tables.length+1}|${r+1}|${c+1}`]=cell.body.paragraphs.items.map(p=>canonicalCellText(p.text || ""));}));
    const section = chosen[index];
    const preceding = items.slice(section.start, section.end).filter(p => p.tableNestingLevel === 0)
      .map(p => p.getRange("Whole").compareLocationWith(tables[i].getRange()));
    await context.sync();
    const after_paragraph = preceding.filter(r => r.value === "Before" || r.value === "AdjacentBefore").length - 1;
    table_addresses.push({slug:chosen[index].slug, global_index:i+1, local_index:result[index].tables.length+1});
    result[index].tables.push({table_index:result[index].tables.length+1, after_paragraph, rows:grid.map(row => row.map(cell => cell.body.paragraphs.items.map(p => canonicalCellText(p.text || "")).join("\n"))),
      paragraph_counts:grid.map(row => row.map(cell => cell.body.paragraphs.items.length))});
  }
  context.document.load("saved");
  await context.sync();
  const captured={sections:result, table_addresses, saved:context.document.saved};
  Object.defineProperty(captured,"section_map",{value:{items,sections}});
  Object.defineProperty(captured,"cell_paragraphs",{value:cell_paragraphs});
  return captured;
}


async function executionTargetRanges(context, payload, targets) {
  if (!payload.executionTargetSection) return targets;
  const {items, sections} = await loadSectionMap(context, true);
  const found = sections.filter(s => s.slug === payload.executionTargetSection);
  if (found.length !== 1) throw refusalError("Target section is missing or ambiguous", "LOCK_SCOPE_UNRESOLVED");
  const s = found[0];
  const range = sectionRange(context, items, s);
  const relations = targets.map(t => t.compareLocationWith(range));
  await context.sync();
  return targets.filter((_, i) => SECTION_WITHIN.has(relations[i].value));
}

// -- search -------------------------------------------------------------

async function opSearch(payload) {
  return measuredWordRun(async (context) => {
    const results = context.document.body.search(payload.find, {
      matchCase: !!payload.matchCase,
      matchWholeWord: !!payload.matchWholeWord,
    });
    results.load("text");
    await context.sync();

    const matches = [];
    for (let i = 0; i < results.items.length; i++) {
      const range = results.items[i];
      const paragraph = range.paragraphs.getFirstOrNullObject();
      paragraph.load("text");
      // eslint-disable-next-line no-await-in-loop -- each match's paragraph
      // must be loaded and synced before the next one is read; Word.run
      // does not support batching getFirstOrNullObject() across matches.
      await context.sync();
      let contextBefore = "";
      let contextAfter = "";
      // Context is scoped to the match's own paragraph, not the full body
      // -- Word's Range API has no direct "N characters before/after in
      // the whole document" primitive short of expanding the range across
      // paragraph boundaries, which is unnecessary for this op's purpose
      // (a human-readable anchor hint, not an exact-position API).
      if (!paragraph.isNullObject && paragraph.text) {
        const idx = paragraph.text.indexOf(range.text);
        if (idx >= 0) {
          contextBefore = paragraph.text.slice(Math.max(0, idx - 20), idx);
          contextAfter = paragraph.text.slice(idx + range.text.length, idx + range.text.length + 20);
        }
      }
      matches.push({ index: i, text: range.text, contextBefore, contextAfter });
    }
    return { matches };
  });
}

// -- replace / format ---------------------------------------------------

// issue #22 (Codex review): both ops toggle changeTrackingMode for the
// duration of the mutation and must restore it afterward even if the
// mutation's own context.sync() throws -- a plain sequential restore (the
// pre-issue-#22 shape) left the document stuck on trackAll on that path.
// Both now restore in a finally block.

// issue #22 B2: resolves payload.rowAnchor's own table row and returns
// its cells, for `find` to be searched within (Body.search per cell,
// merged) instead of the whole document body. Uses only well-documented,
// stable WordApi 1.3 members (Range.parentTableCellOrNullObject,
// TableCell.parentRow, TableRow.cells, TableCell.body) rather than a
// single combined "row range" -- there is no directly analogous
// TableRow.getRange() call in the Word JS API to get one Range spanning
// every cell in a row, so this searches cell-by-cell and merges the
// results instead. Manual sideload verification against a real
// duplicate-cell table (docs/live-mode.md) still needed -- fake_pane.py's
// own row model (a text delimiter) cannot exercise this actual object
// graph, only the server's own rowAnchor plumbing around it.
async function resolveRowCells(context, rowAnchorText) {
  const anchorResults = context.document.body.search(rowAnchorText, { matchCase: true, matchWholeWord: false });
  anchorResults.load("items");
  await context.sync();
  if (anchorResults.items.length !== 1) {
    throw refusalError(
      `rowAnchor ${JSON.stringify(rowAnchorText)} must be unique in the document, found ${anchorResults.items.length}`
    );
  }
  const cell = anchorResults.items[0].parentTableCellOrNullObject;
  cell.load("isNullObject");
  await context.sync();
  if (cell.isNullObject) {
    throw refusalError(`rowAnchor ${JSON.stringify(rowAnchorText)} does not resolve to a table cell`);
  }
  const row = cell.parentRow;
  const rowCells = row.cells;
  rowCells.load("items");
  await context.sync();
  return rowCells.items;
}

async function searchScoped(context, payload) {
  if (payload.scope && payload.scope !== "body") {
    if (payload.rowAnchor) throw refusalError("Row scope cannot be combined with text-box scopes", "INVALID_INPUT");
    const resolved = await resolveScopes(context, payload.scope, true);
    const merged = [];
    for (const item of resolved.bodies) {
      if (item.scope === "body") {
        const found = item.body.search(payload.find, {matchCase: true, matchWholeWord: false});
        found.load("text");
        await context.sync();
        merged.push(...found.items);
        continue;
      }
      // Range.search inside shape text HANGS Word for Mac 16.113.3 when it has a hit (observed: the
      // sync after search froze Word at 100% CPU; a no-match search returned normally), and other
      // hosts return nothing. So shape text is never searched: only exact paragraph content is
      // addressed, through paragraph Content ranges (no paragraph mark, no shape anchor).
      item.body.load("text");
      const paragraphs = item.body.paragraphs;
      paragraphs.load("items/text");
      await context.sync();
      if (!payload.find || !String(item.body.text || "").includes(payload.find)) continue;
      const exact = paragraphs.items.filter(p => p.text === payload.find);
      // Every literal occurrence must be addressed: don't silently skip a
      // substring in another paragraph when a whole paragraph also matches.
      const occurrences = String(item.body.text || "").split(payload.find).length - 1;
      if (!payload.find || exact.length !== occurrences) throw refusalError(
        `Word cannot search inside ${item.shape.name}; this host supports only exact whole-paragraph shape edits`,
        "LIVE_CAPABILITY_MISSING");
      const ranges = exact.map(p => p.getRange("Content"));
      ranges.forEach(r => r.load("text"));
      await context.sync();
      if (ranges.some(r => r.text !== payload.find)) throw refusalError(
        "Shape paragraph content could not be addressed exactly", "LIVE_CAPABILITY_MISSING");
      merged.push(...ranges);
    }
    return merged;
  }
  if (payload.rowAnchor === null || payload.rowAnchor === undefined) {
    const source = payload.executionTargetSection ? await targetSectionRange(context, payload) : context.document.body;
    const results = source.search(payload.find, { matchCase: true, matchWholeWord: false });
    results.load("text");
    await context.sync();
    return executionTargetRanges(context, payload, results.items);
  }
  const rowCells = await resolveRowCells(context, payload.rowAnchor);
  const perCellResults = rowCells.map((cell) =>
    cell.body.search(payload.find, { matchCase: true, matchWholeWord: false })
  );
  perCellResults.forEach((r) => r.load("text"));
  await context.sync();
  const merged = [];
  perCellResults.forEach((r) => merged.push(...r.items));
  return merged;
}

// Explicit segments never derive formatting from a neighbouring run.
function explicitSegments(payload) {
  if (payload.replacement_segments === undefined || payload.replacement_segments === null) return null;
  const segments = payload.replacement_segments;
  if (!Array.isArray(segments) || segments.length > 500 || segments.map(s => s.text).join('') !== payload.replace)
    throw refusalError('Segments must concatenate exactly to replace', 'INVALID_INPUT');
  for (const s of segments) {
    if (!s || typeof s.text !== 'string' || !s.text || /[\x00-\x08\x0a\x0c-\x1f]/.test(s.text) ||
        Object.keys(s).some(k => !['text','format','hyperlink'].includes(k)) ||
        !s.format || Object.keys(s.format).sort().join(',') !== [...FONT_FIELDS].sort().join(',') ||
        ['bold','italic','strikeThrough','doubleStrikeThrough','subscript','superscript'].some(k => typeof s.format[k] !== 'boolean') ||
        (s.format.subscript && s.format.superscript) || !['None','Single','Double'].includes(s.format.underline) ||
        !/^#[0-9a-f]{6}$/i.test(s.format.color) || typeof s.format.name !== 'string' || !s.format.name.trim() ||
        typeof s.format.size !== 'number' || !Number.isFinite(s.format.size) || s.format.size < 1 || s.format.size > 1638 ||
        s.format.size * 2 !== Math.floor(s.format.size * 2) ||
        (s.hyperlink != null && (typeof s.hyperlink !== 'string' || !/^(https?:|mailto:)/.test(s.hyperlink) || /[\x00-\x1f]/.test(s.hyperlink))))
      throw refusalError('Explicit segments require complete valid font properties and an optional URL', 'INVALID_INPUT');
  }
  return segments;
}
function xmlSignature(node) {
  if (node.nodeType === 3) return node.nodeValue;
  const attrs = [...(node.attributes || [])].filter(a => !a.name.startsWith('xmlns') && !a.localName.startsWith('rsid'))
    .map(a => [a.namespaceURI || '',a.localName,a.value]).sort();
  return [node.namespaceURI || '',node.localName,attrs,[...(node.childNodes || [])].map(xmlSignature)];
}
function inlineLedger(xml) {
  const root = shapeXml(xml), W = SHAPE_W;
  const packageRoot = new DOMParser().parseFromString(xml, 'application/xml');
  const documentNode = root.getElementsByTagNameNS(W,'document')[0] || root;
  let paragraphs = [...documentNode.getElementsByTagNameNS(W,'p')].filter(p => {
    for (let n=p.parentNode;n && n!==documentNode;n=n.parentNode)
      if (n.namespaceURI===W && n.localName==='txbxContent') return false;
    return true;
  });
  // Word exports an extra empty final paragraph for some anchored ranges.
  // The range's paragraph collection was independently checked to contain one.
  if (paragraphs.length > 1) {
    const substantive=paragraphs.filter(p => shapeNodes(p).length || [...p.getElementsByTagNameNS(W,'t')].some(t=>t.textContent.length));
    paragraphs=substantive.length ? substantive : paragraphs.slice(0,1);
  }
  if (paragraphs.length !== 1) throw refusalError('Formatted edits require one paragraph: '+JSON.stringify(paragraphs.map(p=>p.textContent.slice(0,150))), 'STRUCTURAL_BOUNDARY');
  const rels = new Map([...packageRoot.getElementsByTagNameNS('http://schemas.openxmlformats.org/package/2006/relationships','Relationship')]
    .filter(r => /\/hyperlink$/.test(r.getAttribute('Type'))).map(r => [r.getAttribute('Id'),r.getAttribute('Target')]));
  const chars = [];
  for (const run of paragraphs[0].getElementsByTagNameNS(W,'r')) {
    let link=null, inShape=false;
    for (let n=run.parentNode;n && n!==paragraphs[0];n=n.parentNode) {
      if (n.namespaceURI===W && n.localName==='txbxContent') inShape=true;
      if (n.namespaceURI===W && n.localName==='hyperlink') {
        const id=n.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id');
        link=rels.get(id) || n.getAttributeNS(W,'anchor') || null;
        if (id && !rels.has(id)) throw refusalError('Cannot resolve hyperlink formatting', 'LIVE_CAPABILITY_MISSING');
      }
    }
    if (inShape) continue;
    const props = [...run.childNodes].find(n => n.namespaceURI===W && n.localName==='rPr');
    const format=JSON.stringify(props ? xmlSignature(props) : null);
    for (const node of run.childNodes) {
      if (node.namespaceURI!==W) continue;
      const text=node.localName==='t' ? node.textContent : node.localName==='tab' ? '\t' : node.localName==='br' ? '\v' : '';
      for (const ch of text) chars.push({ch,format,link});
    }
  }
  return chars;
}
async function formattedBoundary(context, ranges, payload) {
  const collections=ranges.map(r => r.paragraphs);
  collections.forEach(p => p.load('items'));
  await context.sync();
  if (collections.some(p => p.items.length!==1)) throw refusalError('Text replacement crosses paragraphs', 'STRUCTURAL_BOUNDARY');
  const paragraphs=collections.map(p => p.items[0]);
  const xml=paragraphs.map(p => p.getOoxml());
  await context.sync();
  return xml.map((x,i) => {
    const chars=inlineLedger(x.value), text=chars.map(c => c.ch).join('');
    const start=text.indexOf(payload.find);
    if (start<0 || text.indexOf(payload.find,start+1)>=0) throw refusalError('Paragraph target is not unique', 'MATCH_COUNT_MISMATCH');
    const offset=Array.from(text.slice(0,start)).length;
    return {paragraph:paragraphs[i], before:chars.slice(0,offset), after:chars.slice(offset+Array.from(payload.find).length)};
  });
}
async function verifyFormattedBoundary(context, boundaries, payload) {
  const xml=boundaries.map(b => b.paragraph.getOoxml());
  await context.sync();
  boundaries.forEach((b,i) => {
    const chars=inlineLedger(xml[i].value), n=Array.from(payload.replace).length;
    if (JSON.stringify(chars.slice(0,b.before.length))!==JSON.stringify(b.before) ||
        JSON.stringify(chars.slice(b.before.length+n))!==JSON.stringify(b.after) ||
        chars.slice(b.before.length,b.before.length+n).map(c => c.ch).join('')!==payload.replace)
      throw refusalError('Text or formatting outside the replacement changed; reconcile before continuing', 'VERIFICATION_FAILED');
  });
}

// Text edits may touch an anchored paragraph only when the actual replacement
// range excludes every drawing/anchor and the paragraph mark. Never rewrite OOXML.
async function anchoredTextBoundary(context, ranges, payload) {
  const collections=ranges.map(r => r.paragraphs);
  collections.forEach(c => c.load('items'));
  await context.sync();
  if (collections.some(c => c.items.length!==1) || /[\r\n]/.test(payload.find + payload.replace))
    throw refusalError('Anchored text edits cannot replace paragraph boundaries', 'ANCHORED_SHAPES');
  const paragraphs=collections.map(c => c.items[0]);
  const xml=paragraphs.map(p => p.getOoxml());
  await context.sync();
  const boundaries=[];
  for (let i=0;i<xml.length;i++) {
    const root=shapeXml(xml[i].value), nodes=shapeNodes(root);
    if (!nodes.length) continue;
    const targetXml=ranges[i].getOoxml();
    await context.sync();
    if (shapeNodes(shapeXml(targetXml.value)).length)
      throw refusalError('Replacement includes a callout anchor; edit text before or after it', 'ANCHORED_SHAPES');
    const ids=inventoryShapes(xml[i].value).map(s => s.id);
    if (ids.some(id => !id) || new Set(ids).size!==ids.length)
      throw refusalError('Cannot establish stable shape identity for this paragraph', 'ANCHORED_SHAPES');
    boundaries.push({paragraph:paragraphs[i], signatures:nodes.map(n => JSON.stringify(xmlSignature(n))).sort(), ids});
  }
  return boundaries;
}
async function verifyAnchoredText(context, boundaries) {
  if (!boundaries.length) return;
  const xml=boundaries.map(b => b.paragraph.getOoxml());
  await context.sync();
  boundaries.forEach((b,i) => {
    const signatures=shapeNodes(shapeXml(xml[i].value)).map(n => JSON.stringify(xmlSignature(n))).sort();
    if (JSON.stringify(signatures)!==JSON.stringify(b.signatures))
      throw refusalError('Callout shape, wrapping, position, content or anchor paragraph changed; reconcile the write', 'VERIFICATION_FAILED');
  });
}

async function opReplace(payload) {
  const expectedMatches = payload.expected_matches;
  const segments = explicitSegments(payload);
  return measuredWordRun(async (context) => {
    if (!payload.executionSections) await requireFreshBody(context, payload.expectedBodySha256);
    await requireFreshScope(context, payload);
    const preHash = await mutationHash(context, payload);

    const guard = await sectionGuard(context, payload);
    const matchItems = await searchScoped(context, payload);

    if (matchItems.length !== expectedMatches) {
      throw refusalError(
        `expected ${expectedMatches} match(es) for ${JSON.stringify(payload.find)}, found ${matchItems.length}`
      );
    }

    await guard.check(matchItems);
    const anchorBoundaries = await anchoredTextBoundary(context, matchItems, payload);
    await guardRevisions(context, matchItems);
    const shapeBodies = await scopeShapeBodies(context, payload.scope);
    const commentsBefore = await guardComments(context, matchItems, payload.allow_comment_loss, shapeBodies, payload.allow_comment_loss_ids);
    const policy = payload.inherit_format || "replaced";
    if (!["replaced", "previous", "none"].includes(policy)) throw refusalError("Invalid inherit_format", "INVALID_INPUT");
    if (policy === "none" && !Office.context.requirements.isSetSupported("WordApiDesktop", "1.3"))
      throw refusalError("inherit_format=none requires WordApiDesktop 1.3", "LIVE_CAPABILITY_MISSING");
    matchItems.forEach(r => r.font.load(FONT_FIELDS));
    await context.sync();
    const fonts = matchItems.map(r => fontSnapshot(r.font));
    if (segments === null && policy === "replaced" && fonts.some(f => Object.values(f).some(v => v === null || v === undefined || v === "")))
      throw refusalError("Mixed formatting; choose previous or none explicitly", "MIXED_FORMATTING");
    const boundaries = segments !== null ? await formattedBoundary(context, matchItems, payload) : null;
    let segmentRanges = [];
    let inserted = [];
    let previousMode = null;
    if (payload.track_changes) {
      context.document.load("changeTrackingMode");
      await context.sync();
      previousMode = context.document.changeTrackingMode;
    }

    await (await sectionGuard(context, payload)).check(matchItems); // Fresh protected content at mutation boundary.
    matchItems.forEach(range => range.load("text"));
    await context.sync();
    if (matchItems.some(range => range.text !== payload.find))
      throw refusalError("Replacement target changed during preflight", "LIVE_STALE");
    const matches = matchItems.map((range) => ({ before: range.text, after: payload.replace }));
    try {
      if (previousMode !== null) context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
      inserted = matchItems.map((range, i) => {
        if (segments !== null && segments.length) {
          let first=null, last=null;
          const entries=[];
          for (const segment of segments) {
            const target=last ? last.getRange('End').insertText(segment.text, 'After') : range.insertText(segment.text, 'Replace');
            FONT_FIELDS.forEach(k => {target.font[k]=segment.format[k];});
            if (!first) first=target;
            last=target; entries.push({range:target,segment});
          }
          segmentRanges.push(entries);
          return first.expandTo(last);
        }
        const target = range.insertText(payload.replace, Word.InsertLocation.replace);
        if (segments === null && payload.replace && policy === "replaced") FONT_FIELDS.forEach(k => {target.font[k] = fonts[i][k];});
        if (segments === null && payload.replace && policy === "none") target.font.reset();
        return target;
      });
      await context.sync();
      if (segments !== null && segmentRanges.flat().length) {
        // Set links only after all text exists. Clearing a new adjacent range
        // while it still inherits the preceding link can remove that whole link.
        segmentRanges.flat().forEach(e => {e.range.hyperlink='';});
        await context.sync();
        segmentRanges.flat().filter(e => e.segment.hyperlink).forEach(e => {e.range.hyperlink=e.segment.hyperlink;});
        await context.sync();
        segmentRanges.flat().forEach(e => FONT_FIELDS.forEach(k => {e.range.font[k]=e.segment.format[k];}));
        await context.sync();
      }
    } finally {
      if (previousMode !== null) {
        context.document.changeTrackingMode = previousMode;
        await context.sync();
      }
    }

    try {
    segmentRanges.flat().forEach(e => {e.range.load("text,hyperlink");e.range.font.load(FONT_FIELDS);});
    inserted.forEach(r => {r.load("text"); if (payload.replace) r.font.load(FONT_FIELDS);});
    await context.sync();
    inserted.forEach((r, i) => {
      matches[i].after = r.text;
      matches[i].fontBefore = fonts[i];
      matches[i].fontAfter = payload.replace ? fontSnapshot(r.font) : null;
      if (r.text !== payload.replace || (segments === null && payload.replace && policy === "replaced" &&
          FONT_FIELDS.some(k => matches[i].fontAfter[k] !== fonts[i][k])))
        throw refusalError(`Replacement read-back failed: ${JSON.stringify(matches)}`, "VERIFICATION_FAILED");
    });
    for (const entry of segmentRanges.flat()) {
      const {range,segment}=entry;
      if (range.text!==segment.text || (range.hyperlink || null)!==(segment.hyperlink || null) ||
          FONT_FIELDS.some(k => String(range.font[k]).toLowerCase()!==String(segment.format[k]).toLowerCase()))
        throw refusalError('Explicit segment readback failed: '+JSON.stringify({expected:segment,actual:{text:range.text,format:fontSnapshot(range.font),hyperlink:range.hyperlink || null}}), 'VERIFICATION_FAILED');
    }
    if (boundaries) await verifyFormattedBoundary(context, boundaries, payload);
    await verifyAnchoredText(context, anchorBoundaries);
    if (segments !== null) matches.forEach((m,i) => {m.segments=segmentRanges[i]?.map(e => ({text:e.range.text,format:fontSnapshot(e.range.font),hyperlink:e.range.hyperlink || null})) || [];});
    const postHash = await mutationHash(context, {...payload, expectedBodySha256:undefined});

    let commentsRemoved = [];
    if (!payload.executionSections) {
    const remainingScope = await resolveScopes(context, payload.scope || "body", true);
    const remainingComments = [];
    for (const item of remainingScope.bodies) {
      if (item.scope !== "body") {
        remainingComments.push(...(await shapeCommentIds(context, item.body)).map(id => ({id})));
        continue;
      }
      const collection = item.body.getComments();
      collection.load("items/id");
      await context.sync();
      remainingComments.push(...collection.items);
    }
    const remainingIds = new Set(remainingComments.map(c => c.id));
    commentsRemoved = commentsBefore.filter(c => !remainingIds.has(c.id));
    }
    const scopePost = payload.executionSections ? postHash : await scopeHash(context, payload.scope || "body");
    return { scope_post: scopePost, comments_removed: commentsRemoved, applied: true, match_count: matchItems.length, matches, pre: preHash, post: postHash };
    } catch (error) {
      throw refusalError('Replacement may have applied; independent verification failed: '+(error.message || error), 'VERIFICATION_FAILED');
    }
  });
}

async function opFormat(payload) {
  const expectedMatches = payload.expected_matches;
  return measuredWordRun(async (context) => {
    await requireFreshBody(context, payload.expectedBodySha256);
    await requireFreshScope(context, payload);
    const body = context.document.body;
    body.load("text");
    await context.sync();
    const preHash = await sha256Hex(body.text || "");

    const guard = await sectionGuard(context, payload);
    const matchItems = await searchScoped(context, payload);

    if (matchItems.length !== expectedMatches) {
      throw refusalError(
        `expected ${expectedMatches} match(es) for ${JSON.stringify(payload.find)}, found ${matchItems.length}`
      );
    }

    await guard.check(matchItems);
    if (payload.scope && payload.scope !== "body") {
      await guardRevisions(context, matchItems);
      await guardComments(context, matchItems, false, await scopeShapeBodies(context, payload.scope));
    }

    let previousMode = null;
    if (payload.track_changes) {
      context.document.load("changeTrackingMode");
      await context.sync();
      previousMode = context.document.changeTrackingMode;
      context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    }

    const matches = matchItems.map((range) => ({ before: range.text, after: range.text }));
    try {
      matchItems.forEach((range) => {
        if (payload.bold !== null && payload.bold !== undefined) range.font.bold = payload.bold;
        if (payload.italic !== null && payload.italic !== undefined) range.font.italic = payload.italic;
        if (payload.underline !== null && payload.underline !== undefined) {
          range.font.underline = payload.underline ? Word.UnderlineType.single : Word.UnderlineType.none;
        }
        // issue #22: strike/color, the two format_text gained for the
        // proposal-lead incident (marking edits in a font color; strike
        // was silently dropped in live mode before this).
        if (payload.strike !== null && payload.strike !== undefined) range.font.strikeThrough = payload.strike;
        // issue #22: send an explicit "#RRGGBB" -- the documented Office
        // JS convention for font.color, and what its own GETTER always
        // returns (verified against a real Word sideload: font.color
        // read back as "#3B3838" even when the SETTER was given a bare
        // "3B3838" with no "#" -- Word tolerated it, but relying on that
        // leniency instead of matching the getter's own format is not
        // worth it now that it's been checked).
        if (payload.color !== null && payload.color !== undefined) {
          range.font.color = payload.color.startsWith("#") ? payload.color : `#${payload.color}`;
        }
      });
      await context.sync();
    } finally {
      if (previousMode !== null) {
        context.document.changeTrackingMode = previousMode;
        await context.sync();
      }
    }

    // issue #22: re-load font.color/strikeThrough per match AFTER sync,
    // rather than echoing the request back -- lets the server detect a
    // write that didn't actually take (a protected range, a stale
    // object reference) instead of trusting an unconfirmed "applied".
    matchItems.forEach(range => range.load("text"));
    matchItems.forEach((range) => range.font.load(["bold", "italic", "underline", "color", "strikeThrough"]));
    await context.sync();
    matchItems.forEach((range, i) => {
      matches[i].after = range.text;
      matches[i].boldAfter = range.font.bold;
      matches[i].italicAfter = range.font.italic;
      matches[i].underlineAfter = range.font.underline === Word.UnderlineType.none ? false :
        range.font.underline === Word.UnderlineType.single ? true : null;
      matches[i].colorAfter = range.font.color;
      matches[i].strikeAfter = range.font.strikeThrough;
    });

    const postBody = context.document.body;
    postBody.load("text");
    await context.sync();
    const postHash = await sha256Hex(postBody.text || "");

    const scopePost = payload.executionSections ? postHash : await scopeHash(context, payload.scope || "body");
    return { scope_post: scopePost, applied: true, match_count: matchItems.length, matches, pre: preHash, post: postHash };
  });
}

// -- table cells (issue #27) ----------------------------------------------

// Resolves payload.{table_index,row_index,cell_index} (all 1-based; the
// table numbering is list_tables' table_id) to a Word.TableCell.
//
// A document with a NESTED table is refused outright: the server numbers
// tables in document order INCLUDING nested ones, and body.tables' handling
// of nested tables is not something this pane can map back onto that
// numbering with confidence. A refusal is honest; a wrong-cell edit into a
// co-author's open document is not recoverable. (cell_set's own
// compare-and-set on the cell's text is a second line of defense.)
//
// Uses only stable members: Body.tables, Table.nestingLevel, Table.rows,
// TableRow.cells (WordApi 1.3). Rows/cells are addressed by their position
// in rows.items / row.cells.items, matching the server's own <w:tr>/<w:tc>
// counting (a horizontally merged cell counts once; a vertically merged
// continuation cell is still a cell of its row). Manual sideload
// verification against a real table (docs/live-mode.md) is still needed --
// fake_pane.py's table model cannot exercise this object graph.
async function loadTopLevelTables(context, scope = null, batchNesting = false) {
  const tables = (scope || context.document.body).tables;
  tables.load(batchNesting ? "items/nestingLevel" : "items");
  await context.sync();
  if (!batchNesting && tables.items.length) {
    tables.items.forEach((t) => t.load("nestingLevel"));
    await context.sync();
  }
  if (tables.items.some((t) => t.nestingLevel > 1)) {
    throw refusalError(
      "the document contains a nested table; table numbering cannot be mapped reliably, so live cell edits are refused"
    );
  }
  return tables.items;
}

async function targetSectionRange(context, payload) {
  const {items,sections} = await loadSectionMap(context, true);
  const matches = sections.filter(s => s.slug === payload.executionTargetSection);
  if (matches.length !== 1) throw refusalError("Target section missing or ambiguous", "LOCK_SCOPE_UNRESOLVED");
  const s = matches[0];
  return sectionRange(context, items, s);
}

async function targetTables(context, payload) {
  if (!payload.executionTargetSection) return loadTopLevelTables(context);
  const {items,sections} = await loadSectionMap(context, true);
  const matches = sections.filter(s => s.slug === payload.executionTargetSection);
  if (matches.length !== 1) throw refusalError("Target section missing or ambiguous", "LOCK_SCOPE_UNRESOLVED");
  const s = matches[0];
  return loadTopLevelTables(context, sectionRange(context, items, s), !!payload.approvedCellStream);
}

async function mutationHash(context, payload) {
  if (payload.executionSections) {
    const state = await executionSnapshot(context, payload.executionSections.map(s=>s.slug));
    return sha256Hex(JSON.stringify(state.sections));
  }
  return requireFreshBody(context, payload.expectedBodySha256);
}

async function resolveTableCell(context, payload) {
  const tableItems = await targetTables(context, payload);
  const table = tableItems[payload.table_index - 1];
  if (!table) {
    throw refusalError(`no table ${payload.table_index} (the document has ${tableItems.length})`,
      payload.approvedCellStream ? "LIVE_STALE" : undefined);
  }
  if (payload.approvedCellStream && Office.context.requirements.isSetSupported('WordApi', '1.3')) {
    if (!Number.isInteger(payload.row_index) || payload.row_index < 1 ||
        !Number.isInteger(payload.cell_index) || payload.cell_index < 1)
      throw refusalError('Invalid approved cell coordinates', 'INVALID_INPUT');
    return table.getCell(payload.row_index - 1, payload.cell_index - 1);
  }
  const rows = table.rows;
  rows.load("items");
  await context.sync();
  const row = rows.items[payload.row_index - 1];
  if (!row) {
    throw refusalError(
      `row_index ${payload.row_index} is out of range for table ${payload.table_index} (${rows.items.length} row(s))`
    );
  }
  row.cells.load("items");
  await context.sync();
  const cell = row.cells.items[payload.cell_index - 1];
  if (!cell) {
    throw refusalError(
      `cell_index ${payload.cell_index} is out of range for row ${payload.row_index} (${row.cells.items.length} cell(s))`
    );
  }
  return cell;
}

// Writes one cell's paragraphs. Shared by cell_set, cells_set and table_insert
// so a cell is always written one way. A single-paragraph cell is only QUEUED
// (the caller syncs); a multi-paragraph cell needs syncs in the middle, so
// callers must await this.
//
// issue #75, what real Word (Mac) does to a table cell:
//  * Paragraph.insertBreak / Range.insertBreak accept only "Before"/"After";
//    "End" is InvalidArgument. Break after the last inserted range.
//  * Writing text through the proxy that body.insertParagraph("", "End")
//    returns (or through paragraphs.getLast() queued in the same batch) put
//    paragraph 2's text into paragraph 1 and left an empty trailing
//    paragraph. So no write ever depends on how a not-yet-synced proxy
//    resolves. Strategies (opCellSet tries them in order and verifies each):
//      "load-items": insert the empty paragraphs, sync, load the REAL
//                    paragraph list, then write into items[i] by position.
//      "after-chain": paragraph.insertParagraph("", "After") from the
//                    previous paragraph's own proxy.
const CELL_WRITE_STRATEGIES = ["load-items", "after-chain"];

function writeRunsInto(paragraph, runs) {
  let lastRange = null;
  (runs || []).forEach((run) => {
    if (run.hard_break) {
      const anchor = lastRange || paragraph.getRange("Start");
      anchor.insertBreak(Word.BreakType.line, Word.InsertLocation.after);
      return;
    }
    const range = paragraph.insertText(run.text, Word.InsertLocation.end);
    range.font.bold = !!run.bold;
    range.font.italic = !!run.italic;
    if (run.link) {
      range.hyperlink = run.link;
    }
    lastRange = range;
  });
}

async function writeCellParagraphs(context, cell, paragraphs, strategy = CELL_WRITE_STRATEGIES[0]) {
  cell.body.clear();
  // After clear() the cell holds one empty paragraph.
  if (paragraphs.length <= 1) {
    writeRunsInto(cell.body.paragraphs.getFirst(), paragraphs[0]);
    return;
  }
  if (strategy === "after-chain") {
    let paragraph = cell.body.paragraphs.getFirst();
    paragraphs.forEach((runs, i) => {
      if (i > 0) paragraph = paragraph.insertParagraph("", Word.InsertLocation.after);
      writeRunsInto(paragraph, runs);
    });
    return;
  }
  for (let i = 1; i < paragraphs.length; i++) cell.body.insertParagraph("", Word.InsertLocation.end);
  await context.sync();
  cell.body.paragraphs.load("items");
  await context.sync();
  const items = cell.body.paragraphs.items;
  paragraphs.forEach((runs, i) => {
    if (items[i]) writeRunsInto(items[i], runs);
  });
}

const canonicalCellText = s => String(s).replace(/\r\n?|\v/g, "\n");
const normCellText = (s) => String(s).replace(/\s+/g, " ").trim();

// One string per intended paragraph; a hard break reads back as whitespace
// (Word reports a line break as \v), so it is rendered as "\n" here and the
// comparison normalizes whitespace. An empty write leaves one empty paragraph.
function intendedCellParagraphs(paragraphs) {
  const out = (paragraphs || []).map((runs) =>
    (runs || []).map((run) => (run.hard_break ? "\n" : run.text)).join("")
  );
  return out.length ? out : [""];
}

async function readCellParagraphs(context, cell) {
  cell.body.paragraphs.load("items/text");
  await context.sync();
  return cell.body.paragraphs.items.map((p) => p.text || "");
}

const sameParagraphs = (want, got) =>
  want.length === got.length && want.every((w, i) => normCellText(w) === normCellText(got[i]));

// Rewrites a cell as plain paragraphs of `texts` (character formatting is not
// preserved) and returns true only if it reads back as those texts. Used when
// a write failed its self-check; it never uses OOXML (a getOoxml/insertOoxml
// round trip put the neighbouring cell's text into this one on real Word).
async function restoreCellText(context, cell, texts) {
  const specs = texts.map((t) => [{ text: String(t).replace(/[\r\n\v]+$/, ""), bold: false, italic: false }]);
  for (const strategy of CELL_WRITE_STRATEGIES) {
    try {
      await writeCellParagraphs(context, cell, specs, strategy);
      await context.sync();
      const got = await readCellParagraphs(context, cell);
      if (sameParagraphs(texts, got)) return true;
    } catch (err) {
      // try the next strategy
    }
  }
  return false;
}

async function opCellGet(payload) {
  return measuredWordRun(async (context) => {
    const cell = await resolveTableCell(context, payload);
    cell.body.load("text");
    cell.body.paragraphs.load("items/text");
    // issue #107: the server derives the cell's first-run font from this OOXML, so
    // file mode and live mode share one definition of "first run".
    const ooxml = payload.include_ooxml ? cell.body.getOoxml() : null;
    await context.sync();
    const result = { text: cell.body.text || "", paragraphs: cell.body.paragraphs.items.map((p) => p.text || "") };
    if (ooxml) result.ooxml = ooxml.value;
    return result;
  });
}

// issue #107: the font the server read from the cell's first run, applied to the whole
// cell after a write and read back. Name, size and colour only: the Word API cannot
// set the complex-script size.
function applyInheritedFont(cell, font) {
  for (const key of ["name", "size", "color"]) {
    const value = font[key];
    if (value !== null && value !== undefined && value !== "") cell.body.font[key] = key === "color" && !String(value).startsWith("#") ? `#${value}` : value;
  }
}

async function verifyInheritedFont(context, cell, font) {
  cell.body.font.load("name,size,color");
  await context.sync();
  for (const key of ["name", "size", "color"]) {
    const want = font[key];
    if (want === null || want === undefined || want === "") continue;
    const got = cell.body.font[key];
    const same = key === "color" ? String(got).replace("#", "").toUpperCase() === String(want).replace("#", "").toUpperCase() : got === want;
    if (!same) throw refusalError(`Inherited cell font ${key} failed readback (wanted ${want}, found ${got})`, "VERIFICATION_FAILED");
  }
}

async function opCellSet(payload) {
  return measuredWordRun(async (context) => {
    const preHash = await mutationHash(context, payload);

    const cell = await resolveTableCell(context, payload);
    cell.body.load("text");
    await context.sync();
    const before = payload.executionSections ? (await readCellParagraphs(context, cell)).map(canonicalCellText).join("\n") : (cell.body.text || "");
    // Compare-and-set: refuse if the cell changed since the server read it
    // (a co-author typing in this same cell), rather than overwrite them.
    if (before !== payload.expected_before_text) {
      throw refusalError(
        "the cell's text changed since it was read (another editor may be working in it); nothing was written", "CELL_TEXT_MISMATCH"
      );
    }

    await (await sectionGuard(context, payload)).check(() => [cell.body.getRange("Whole")]);
    if (payload.executionSections) {
      await guardComments(context, [cell.body.getRange("Whole")], false);
      await guardRevisions(context, [cell.body.getRange("Whole")]);
    }
    await guardAnchoredBodies(context, [cell.body]);
    let previousMode = null;
    if (payload.track_changes) {
      context.document.load("changeTrackingMode");
      await context.sync();
      previousMode = context.document.changeTrackingMode;
    }

    // The cell's paragraphs before the write, for the restore below (issue #75).
    const beforeParagraphs = await readCellParagraphs(context, cell);
    if (canonicalCellText(beforeParagraphs.join("\n")) !== canonicalCellText(payload.expected_before_text) ||
        (payload.expected_before_paragraph_count != null && beforeParagraphs.length !== payload.expected_before_paragraph_count))
      throw refusalError("Cell changed during preflight; nothing was written", "LIVE_STALE");
    const want = intendedCellParagraphs(payload.paragraphs);

    // Write, then self-check: the cell must hold exactly the intended
    // paragraphs. An untracked write that does not is retried with the next
    // strategy (each starts with clear()); if none works the old TEXT is
    // written back. A tracked write is tried once and never rewritten (the
    // user can reject it in the Review pane).
    let usedStrategy = null;
    let got = [];
    try {
      if (previousMode !== null) context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
      const strategies = payload.track_changes ? CELL_WRITE_STRATEGIES.slice(0, 1) : CELL_WRITE_STRATEGIES;
      for (const strategy of strategies) {
        await writeCellParagraphs(context, cell, payload.paragraphs || [], strategy);
        if (payload.inherit_font) applyInheritedFont(cell, payload.inherit_font);
        await context.sync();
        got = await readCellParagraphs(context, cell);
        if (sameParagraphs(want, got)) {
          usedStrategy = strategy;
          break;
        }
      }
    } finally {
      if (previousMode !== null) {
        context.document.changeTrackingMode = previousMode;
        await context.sync();
      }
    }
    if (usedStrategy !== null && payload.inherit_font) await verifyInheritedFont(context, cell, payload.inherit_font);
    if (usedStrategy === null) {
      const detail = `expected ${JSON.stringify(want.map(normCellText))}, found ${JSON.stringify(got.map(normCellText))}`;
      if (payload.track_changes) {
        throw refusalError(
          `the cell did not read back as written (${detail}); the write is a tracked change -- reject it in Word's Review pane`,
          "cell_write_not_rolled_back"
        );
      }
      const rolledBack = await restoreCellText(context, cell, beforeParagraphs);
      throw refusalError(
        `the cell did not read back as written (${detail}); ` +
          (rolledBack
            ? "the previous text was written back (character formatting may differ)"
            : "the previous content could NOT be restored -- fix the cell by hand"),
        rolledBack ? "cell_write_rolled_back" : "cell_write_not_rolled_back"
      );
    }

    cell.body.load("text");
    const postHash = await mutationHash(context, {...payload, expectedBodySha256:undefined});

    return { applied: true, before, after: cell.body.text || "", pre: preHash, post: postHash, strategy: usedStrategy };
  });
}

// -- live tables (issue #34) ------------------------------------------------

// Server vocabulary <-> Word enum names. Word's TableCell.horizontalAlignment
// takes "Centered"/"Justified", not the "center"/"both" the tool exposes.
const H_ALIGN_TO_WORD = { left: "Left", center: "Centered", right: "Right", both: "Justified" };
const H_ALIGN_FROM_WORD = { Left: "left", Centered: "center", Right: "right", Justified: "both" };
const V_ALIGN_TO_WORD = { top: "Top", center: "Center", bottom: "Bottom" };
const V_ALIGN_FROM_WORD = { Top: "top", Center: "center", Bottom: "bottom" };

function withHash(hex) {
  return hex.startsWith("#") ? hex : `#${hex}`;
}

// Same normalization as the server's tables._texts_match_ladder: straight
// quotes, no soft hyphens, collapsed whitespace, trimmed.
// Built from code points (not literals) so no invisible character lives in the source.
const charClass = (codes) => new RegExp(`[${codes.map((c) => String.fromCharCode(c)).join("")}]`, "g");
const SMART_SINGLE_QUOTES = charClass([0x2018, 0x2019, 0x201a, 0x201b]);
const SMART_DOUBLE_QUOTES = charClass([0x201c, 0x201d, 0x201e, 0x201f]);
const SOFT_HYPHEN = charClass([0x00ad]);

function normalizeAnchorText(text) {
  return String(text || "")
    .replace(SMART_SINGLE_QUOTES, "'")
    .replace(SMART_DOUBLE_QUOTES, '"')
    .replace(SOFT_HYPHEN, "")
    .replace(/\s+/g, " ")
    .trim();
}

// Runs fn with change tracking forced to trackAll when `on`, and restores the
// previous mode in a finally (a throwing sync must not leave the document
// stuck on trackAll -- same rule as replace/format).
async function withTracking(context, on, fn) {
  let previousMode = null;
  if (on) {
    context.document.load("changeTrackingMode");
    await context.sync();
    previousMode = context.document.changeTrackingMode;
    context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
  }
  try {
    return await fn();
  } finally {
    if (previousMode !== null) {
      context.document.changeTrackingMode = previousMode;
      await context.sync();
    }
  }
}

// Reads the body hash and refuses (code "stale") BEFORE any write if the
// caller's expectedBodySha256 no longer matches. The session's own check
// compares result.pre only after the op has already run, which cannot undo a
// write that landed.
async function requireFreshBody(context, expectedHash) {
  const body = context.document.body;
  body.load("text");
  await context.sync();
  const preHash = await sha256Hex(body.text || "");
  if (expectedHash && expectedHash !== preHash) {
    throw refusalError(
      "the document changed since it was read (body hash mismatch); nothing was written. Re-read and retry.",
      "stale"
    );
  }
  return preHash;
}

async function loadRowsAndCells(context, table) {
  const rows = table.rows;
  rows.load("items");
  await context.sync();
  rows.items.forEach((row) => row.cells.load("items"));
  await context.sync();
  return rows.items.map((row) => row.cells.items);
}

// The table's own part of its flat-OPC XML, so a style definition elsewhere in
// the package cannot be mistaken for a merged cell.
function documentPartXml(flatOpc) {
  const start = flatOpc.indexOf('pkg:name="/word/document.xml"');
  if (start < 0) return flatOpc;
  const end = flatOpc.indexOf("</pkg:part>", start);
  return end < 0 ? flatOpc.slice(start) : flatOpc.slice(start, end);
}

async function opTableGet(payload) {
  return measuredWordRun(async (context) => {
    const tableItems = await targetTables(context, payload);
    const table = tableItems[payload.table_index - 1];
    if (!table) {
      throw refusalError(
        `no table ${payload.table_index} (the document has ${tableItems.length})`,
        "table_not_found"
      );
    }
    table.load(["style", "headerRowCount"]);
    const ooxml = table.getRange().getOoxml();
    await context.sync();
    const grid = await loadRowsAndCells(context, table);
    grid.forEach((cells) =>
      cells.forEach((cell) => {
        cell.load(["shadingColor", "horizontalAlignment", "verticalAlignment", "columnWidth"]);
        cell.body.load("text");
        if (payload.executionTargetSection) cell.body.paragraphs.load("items/text");
        cell.body.font.load(["bold", "color", "size"]);
      })
    );
    await context.sync();
    const part = documentPartXml(ooxml.value || "");
    return {
      style: table.style || "",
      headerRowCount: table.headerRowCount,
      merged: /<w:gridSpan[\s/>]/.test(part) || /<w:vMerge[\s/>]/.test(part),
      rows: grid.map((cells) =>
        cells.map((cell) => ({
          text: payload.executionTargetSection ? cell.body.paragraphs.items.map(p => canonicalCellText(p.text || "")).join("\n") : (cell.body.text || ""),
          fill: cell.shadingColor || "",
          align: H_ALIGN_FROM_WORD[cell.horizontalAlignment] || String(cell.horizontalAlignment),
          valign: V_ALIGN_FROM_WORD[cell.verticalAlignment] || String(cell.verticalAlignment),
          width: cell.columnWidth,
          bold: cell.body.font.bold,
          color: cell.body.font.color || "",
          size: cell.body.font.size,
        }))
      ),
    };
  });
}

// Resolves the style to apply BEFORE anything is inserted, so a bad style never
// leaves a stray table behind (an insert-then-delete "rollback" is not one).
// Returns {style} or {styleBuiltIn}.
async function resolveInsertStyle(context, payload, tableItems) {
  if (payload.style_from_table_index) {
    const source = tableItems[payload.style_from_table_index - 1];
    if (!source) {
      throw refusalError(
        `style_from_table_index ${payload.style_from_table_index} is out of range (the document has ${tableItems.length})`,
        "style_not_found"
      );
    }
    source.load("style");
    await context.sync();
    if (!source.style) {
      throw refusalError(`table ${payload.style_from_table_index} has no table style to copy`, "style_not_found");
    }
    return { style: source.style };
  }
  if (payload.style_builtin) {
    const enumValues = Word.BuiltInStyleName || {};
    const value = enumValues[payload.style_builtin] || Object.values(enumValues).find((v) => v === payload.style_builtin);
    if (!value) {
      throw refusalError(`${JSON.stringify(payload.style_builtin)} is not a Word built-in style name`, "style_not_found");
    }
    return { styleBuiltIn: value };
  }
  if (payload.style) {
    tableItems.forEach((t) => t.load("style"));
    await context.sync();
    if (tableItems.some((t) => t.style === payload.style)) return { style: payload.style };
    if (Office.context.requirements.isSetSupported("WordApi", "1.5")) {
      const found = context.document.getStyles().getByNameOrNullObject(payload.style);
      found.load("isNullObject");
      await context.sync();
      if (!found.isNullObject) return { style: payload.style };
    }
    throw refusalError(
      `table style ${JSON.stringify(payload.style)} was not found (it is not used by an existing table and the ` +
        "document's style list does not name it); use style_from_table_id or a built-in style",
      "style_not_found"
    );
  }
  return {};
}

function findAnchorParagraph(paragraphItems, anchor) {
  const wanted = normalizeAnchorText(anchor.paragraph_text);
  const matches = paragraphItems.filter((p) => p.tableNestingLevel === 0 && normalizeAnchorText(p.text) === wanted);
  if (matches.length !== 1) {
    throw refusalError(
      `expected 1 top-level paragraph matching ${JSON.stringify(anchor.paragraph_text)}, found ${matches.length}`,
      matches.length === 0 ? "zero_match" : "match_count_mismatch"
    );
  }
  return matches[0];
}

// issue #79: Paragraph.detachFromList (WordApi 1.3). `getParagraphs` returns a FRESH collection
// each call, so a recount after the detach never reads a stale proxy.
async function detachListItems(context, getParagraphs) {
  const paragraphs = getParagraphs();
  paragraphs.load("items/isListItem");
  await context.sync();
  const listed = paragraphs.items.filter((p) => p.isListItem);
  listed.forEach((p) => p.detachFromList());
  if (listed.length) await context.sync();
  return listed.length;
}

async function countListItems(context, getParagraphs) {
  const paragraphs = getParagraphs();
  paragraphs.load("items/isListItem");
  await context.sync();
  return paragraphs.items.filter((p) => p.isListItem).length;
}

async function opTableInsert(payload) {
  const rowCount = payload.rows.length;
  const colCount = payload.rows[0].length;
  return measuredWordRun(async (context) => {
    const preHash = await requireFreshBody(context, payload.expectedBodySha256);
    const tableItems = await loadTopLevelTables(context);
    const styleChoice = await resolveInsertStyle(context, payload, tableItems);

    const anchor = payload.anchor || null;
    let paragraphItems = null;
    if (anchor && anchor.paragraph_text !== undefined) {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items/text,items/tableNestingLevel");
      await context.sync();
      paragraphItems = paragraphs.items;
    }
    let target = null;
    if (anchor && anchor.paragraph_text !== undefined) {
      target = findAnchorParagraph(paragraphItems, anchor);
    } else if (anchor && anchor.after_table_index !== undefined) {
      target = tableItems[anchor.after_table_index - 1];
      if (!target) {
        throw refusalError(
          `after_table_index ${anchor.after_table_index} is out of range (the document has ${tableItems.length})`,
          "table_not_found"
        );
      }
    }

    const guard = await sectionGuard(context, payload);
    if (guard.active) {
      // "before" a heading would land outside the section the anchor is in.
      if (anchor && anchor.paragraph_text !== undefined && anchor.position === "before")
        throw refusalError("Under section locks insert tables after an anchor, not before it", "OUTSIDE_LOCKED_SECTION");
      const anchorRange = target
        ? target.getRange()
        : context.document.body.paragraphs.getLast().getRange();
      await guard.check([anchorRange]);
    }
    const emptyValues = payload.rows.map((row) => row.map(() => ""));
    let newTable = null;
    let listItemsInTable = 0;
    await withTracking(context, payload.track_changes, async () => {
      if (target && anchor.paragraph_text !== undefined) {
        newTable = target.insertTable(rowCount, colCount, anchor.position === "before" ? "Before" : "After", emptyValues);
      } else if (target) {
        newTable = target.insertTable(rowCount, colCount, "After", emptyValues);
      } else {
        newTable = context.document.body.insertTable(rowCount, colCount, "End", emptyValues);
      }
      if (styleChoice.style) newTable.style = styleChoice.style;
      if (styleChoice.styleBuiltIn) newTable.styleBuiltIn = styleChoice.styleBuiltIn;
      newTable.headerRowCount = payload.header_rows || 0;
      const grid = await loadRowsAndCells(context, newTable);
      for (const [r, row] of payload.rows.entries()) {
        for (const [c, spec] of row.entries()) {
          const cell = grid[r][c];
          await writeCellParagraphs(context, cell, spec.paragraphs || []);
          if (spec.bold) cell.body.font.bold = true;
          if (spec.color) cell.body.font.color = withHash(spec.color);
          if (spec.fill) cell.shadingColor = withHash(spec.fill);
          if (spec.align) cell.horizontalAlignment = H_ALIGN_TO_WORD[spec.align];
          if (spec.valign) cell.verticalAlignment = V_ALIGN_TO_WORD[spec.valign];
          if (payload.column_widths_pt) cell.columnWidth = payload.column_widths_pt[c];
        }
      }
      if (payload.font_size_pt) newTable.font.size = payload.font_size_pt;
      await context.sync();
      // #79: a table inserted next to a list item can inherit its numbering in every cell paragraph.
      await detachListItems(context, () => newTable.getRange().paragraphs);
      listItemsInTable = await countListItems(context, () => newTable.getRange().paragraphs);
    });

    // The new table's own 1-based index (list_tables numbering): compare each
    // top-level table's range with the new one. compareLocationWith returns a
    // ClientResult, so .value is only readable after the sync.
    const allTables = context.document.body.tables;
    allTables.load("items");
    await context.sync();
    const relations = allTables.items.map((t) => t.getRange().compareLocationWith(newTable.getRange()));
    await context.sync();
    const position = relations.findIndex((rel) => rel.value === "Equal");
    if (position < 0) {
      throw refusalError(
        "the table was inserted but its position could not be determined; call list_tables to find it. Nothing was rolled back."
      );
    }

    const postBody = context.document.body;
    postBody.load("text");
    await context.sync();
    const postHash = await sha256Hex(postBody.text || "");
    return { applied: true, table_index: position + 1, list_items_in_table: listItemsInTable, pre: preHash, post: postHash };
  });
}

async function opCellsSet(payload) {
  return measuredWordRun(async (context) => {
    const preHash = await mutationHash(context, payload);
    const tableItems = await targetTables(context, payload);

    const gridByTable = new Map();
    for (const spec of payload.cells) {
      const table = tableItems[spec.table_index - 1];
      if (!table) {
        throw refusalError(
          `no table ${spec.table_index} (the document has ${tableItems.length})`,
          "table_not_found"
        );
      }
      if (!gridByTable.has(spec.table_index)) {
        // eslint-disable-next-line no-await-in-loop -- one table at a time keeps the batch small
        gridByTable.set(spec.table_index, await loadRowsAndCells(context, table));
      }
    }
    const cells = payload.cells.map((spec) => {
      const grid = gridByTable.get(spec.table_index);
      const row = grid[spec.row_index - 1];
      if (!row) throw refusalError(`row_index ${spec.row_index} is out of range for table ${spec.table_index}`);
      const cell = row[spec.cell_index - 1];
      if (!cell) throw refusalError(`cell_index ${spec.cell_index} is out of range for row ${spec.row_index}`);
      return cell;
    });

    // Word batches are not transactions. Narrow the window as far as the API
    // allows: compare every cell, then read them ALL again immediately before
    // queuing the writes, and refuse (writing nothing) on any difference.
    let writingFinished = false;
    const readAll = async () => {
      if (payload.executionSections) {
        const texts = [];
        for (let i=0;i<cells.length;i++) {
          const paragraphs = await readCellParagraphs(context, cells[i]);
          if (!writingFinished && payload.cells[i].expected_before_paragraph_count != null && paragraphs.length !== payload.cells[i].expected_before_paragraph_count)
            throw refusalError("Row paragraph structure changed", "LIVE_STALE");
          texts.push(paragraphs.map(canonicalCellText).join("\n"));
        }
        return texts;
      }
      cells.forEach((cell) => cell.body.load("text"));
      await context.sync();
      return cells.map((cell) => payload.executionSections ? canonicalCellText(cell.body.text || "") : (cell.body.text || ""));
    };
    const mismatch = (texts) => payload.cells.findIndex((spec, i) => texts[i] !== spec.expected_before_text);
    const before = await readAll();
    let bad = mismatch(before);
    if (bad < 0) bad = mismatch(await readAll());
    if (bad >= 0) {
      throw refusalError(
        `cell ${bad + 1} of ${payload.cells.length} changed since it was read (another editor may be working in it); nothing was written`, "CELL_TEXT_MISMATCH"
      );
    }

    await (await sectionGuard(context, payload)).check(() => cells.map(c => c.body.getRange("Whole")));
    if (payload.executionSections) {
      await guardComments(context, cells.map(c => c.body.getRange("Whole")), false);
      await guardRevisions(context, cells.map(c => c.body.getRange("Whole")));
    }
    await guardAnchoredBodies(context, cells.map(c => c.body));
    if (mismatch(await readAll()) >= 0)
      throw refusalError("Row changed during preflight; nothing was written", "LIVE_STALE");
    await withTracking(context, payload.track_changes, async () => {
      if (mismatch(await readAll()) >= 0)
        throw refusalError("Row changed immediately before mutation", "LIVE_STALE");
      for (const [i, spec] of payload.cells.entries()) await writeCellParagraphs(context, cells[i], spec.paragraphs || []);
      await context.sync();
    });

    writingFinished = true;
    const after = await readAll();
    const postHash = await mutationHash(context, {...payload, expectedBodySha256:undefined});
    return { applied: true, before, after, pre: preHash, post: postHash };
  });
}

// -- comments -------------------------------------------------------------

// payload (all optional, added in issue #31): `ids` restricts the result to
// those comment ids; `include_anchor: false` skips the per-comment
// getRange() (used by reply/resolve verification, which never reads it).
// `timing_ms` lets a caller see where a slow list went.
const COMMENTS_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const COMMENTS_W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";
const COMMENTS_W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml";

function commentsXmlAttr(el, ns, local, qname) {
  const direct = el.getAttribute(qname);
  if (direct !== null && direct !== undefined) return direct;
  const viaNs = typeof el.getAttributeNS === "function" ? el.getAttributeNS(ns, local) : null;
  return viaNs === "" ? null : viaNs;
}

// issue #39: count open comment THREADS from the comments parts of a flat OPC
// package (body.getOoxml()), independently of body.getComments(). A thread
// root is a w:comment whose last paragraph's w14:paraId has no paraIdParent in
// commentsExtended.xml; replies carry a paraIdParent and are not threads. A root
// is resolved when its commentEx done is "1"/"true". `available` is false when
// either part is missing or the package does not parse, so the caller never
// treats an absent cross-check as "complete". `ids` are w:id values (not
// necessarily the Office.js comment ids).
function crosscheckCommentsOoxml(xml) {
  const unavailable = (reason) => ({ available: false, total_threads: null, open_threads: null, ids: [], reason });
  try {
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    if (doc.getElementsByTagName("parsererror").length) return unavailable("unparseable");
    const partNames = [...doc.getElementsByTagNameNS(FLAT_OPC_NS, "part")]
      .map((part) => part.getAttribute("pkg:name"));
    if (!partNames.includes("/word/comments.xml")) return unavailable("comments part missing");
    if (!partNames.includes("/word/commentsExtended.xml")) return unavailable("commentsExtended part missing");
    const extended = new Map();
    for (const ex of doc.getElementsByTagNameNS(COMMENTS_W15_NS, "commentEx")) {
      const paraId = commentsXmlAttr(ex, COMMENTS_W15_NS, "paraId", "w15:paraId");
      if (!paraId) continue;
      const done = commentsXmlAttr(ex, COMMENTS_W15_NS, "done", "w15:done");
      extended.set(paraId, {
        parent: commentsXmlAttr(ex, COMMENTS_W15_NS, "paraIdParent", "w15:paraIdParent"),
        done: done === "1" || done === "true",
      });
    }
    let total = 0;
    const ids = [];
    for (const comment of doc.getElementsByTagNameNS(COMMENTS_W_NS, "comment")) {
      const paragraphs = comment.getElementsByTagNameNS(COMMENTS_W_NS, "p");
      const last = paragraphs.length ? paragraphs[paragraphs.length - 1] : null;
      const paraId = last ? commentsXmlAttr(last, COMMENTS_W14_NS, "paraId", "w14:paraId") : null;
      const ex = paraId ? extended.get(paraId) : undefined;
      if (ex && ex.parent) continue; // a reply, not a thread
      total += 1;
      if (!(ex && ex.done)) ids.push(commentsXmlAttr(comment, COMMENTS_W_NS, "id", "w:id"));
    }
    return { available: true, total_threads: total, open_threads: ids.length, ids };
  } catch (err) {
    return unavailable(`parse error: ${err && err.message ? err.message : err}`);
  }
}

async function opCommentsList(payload) {
  const p = payload || {};
  const started = paneNow();
  return measuredWordRun(async (context) => {
    const { items, counts, observedAt } = await collectComments(context, {
      ids: p.ids,
      anchors: p.include_anchor !== false,
      replies: true,
    });
    // issue #39: the unfiltered listing also reads the body's OOXML so a gap between
    // getComments() and the comments parts is visible. A failure here must not fail the list.
    let ooxml;
    if (p.ids === undefined || p.ids === null) {
      try {
        const packaged = context.document.body.getOoxml();
        await context.sync();
        ooxml = crosscheckCommentsOoxml(packaged.value);
      } catch (err) {
        ooxml = { available: false, total_threads: null, open_threads: null, ids: [], reason: `read failed: ${err && err.message ? err.message : err}` };
      }
    }
    // issue #39: only the body collection is read (body.getComments()), so
    // comments anchored outside the main body (headers, footers, text boxes;
    // issue #35) are not covered. `scope` says so instead of implying more.
    // `counts`/`observed_at` describe the whole collection at read time.
    if (p.ids === undefined || p.ids === null) updateCommentsCounter(counts.open, observedAt);
    return {
      comments: items,
      counts,
      scope: "body",
      observed_at: observedAt,
      session_epoch: paneEpoch,
      ...(ooxml ? { ooxml } : {}),
      timing_ms: { total: paneNow() - started },
    };
  });
}

async function opCommentAdd(payload) {
  const expectedMatches = payload.expected_matches;
  return measuredWordRun(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();
    const preHash = await sha256Hex(body.text || "");

    const results = body.search(payload.find, { matchCase: true, matchWholeWord: false });
    results.load("text");
    await context.sync();

    if (results.items.length !== expectedMatches) {
      // issue #106 WP-4: zero_match vs match_count_mismatch, the same
      // split live/protocol.py's OP_ERROR_* constants name (see
      // refusalError above) -- server.py maps these onto file mode's own
      // ZERO_MATCH/MATCH_COUNT_MISMATCH.
      const code = results.items.length === 0 ? "zero_match" : "match_count_mismatch";
      throw refusalError(
        `expected ${expectedMatches} match(es) for ${JSON.stringify(payload.find)}, found ${results.items.length}`,
        code
      );
    }

    const guard = await sectionGuard(context, payload);
    await guard.check(() => [results.items[0]]);
    const comment = results.items[0].insertComment(payload.text);
    comment.load("id");
    await context.sync();

    const postBody = context.document.body;
    postBody.load("text");
    await context.sync();
    const postHash = await sha256Hex(postBody.text || "");

    return { comment_id: comment.id, pre: preHash, post: postHash };
  });
}

async function opCommentReply(payload) {
  return measuredWordRun(async (context) => {
    checkCommentEpoch(payload);
    const comments = context.document.body.getComments();
    comments.load("items/id");
    await context.sync();
    const comment = comments.items.find((c) => c.id === payload.comment_id);
    if (!comment) throw refusalError(`no comment with id ${JSON.stringify(payload.comment_id)}`);
    await verifyCommentIdentity(context, comment, payload.identity);
    await (await sectionGuard(context, payload)).check(() => [comment.getRange()]);
    const reply = comment.reply(payload.text);
    reply.load("id");
    await context.sync();
    return { reply_id: reply.id };
  });
}

async function opCommentResolve(payload) {
  return measuredWordRun(async (context) => {
    checkCommentEpoch(payload);
    const comments = context.document.body.getComments();
    comments.load("items/id");
    await context.sync();
    const comment = comments.items.find((c) => c.id === payload.comment_id);
    if (!comment) throw refusalError(`no comment with id ${JSON.stringify(payload.comment_id)}`);
    await verifyCommentIdentity(context, comment, payload.identity);
    await (await sectionGuard(context, payload)).check(() => [comment.getRange()]);
    comment.resolved = !!payload.resolved;
    await context.sync();
    return { resolved: !!payload.resolved };
  });
}

// issue #80: delete a whole thread (Comment.delete() removes its replies).
// The pane only reports what it re-read afterwards; the server decides
// whether `still_present` is a failure.
async function opCommentDelete(payload) {
  return measuredWordRun(async (context) => {
    checkCommentEpoch(payload);
    const comments = context.document.body.getComments();
    comments.load("items/id");
    await context.sync();
    const comment = comments.items.find((c) => c.id === payload.comment_id);
    if (!comment) throw refusalError(`no comment with id ${JSON.stringify(payload.comment_id)}`);
    await verifyCommentIdentity(context, comment, payload.identity);
    await (await sectionGuard(context, payload)).check(() => [comment.getRange()]);
    comment.load("resolved");
    const replies = comment.replies; // CommentReplyCollection is a property, not a method
    replies.load("items/id");
    await context.sync();
    if (payload.require_resolved && !comment.resolved)
      throw refusalError(`comment ${JSON.stringify(payload.comment_id)} is not resolved`);
    const replyCount = replies.items.length;
    comment.delete();
    await context.sync();
    const after = context.document.body.getComments();
    after.load("items/id");
    await context.sync();
    return {
      deleted: true,
      reply_count: replyCount,
      still_present: after.items.some((c) => c.id === payload.comment_id),
    };
  });
}

async function opSave() {
  return measuredWordRun(async (context) => {
    context.document.save();
    await context.sync();
    return { saved: true };
  });
}

Office.onReady((info) => {
  bootTimings.office_ready = Date.now() - SCRIPT_START_MS;
  if (info.host !== Office.HostType.Word) {
    setStatus(`Loaded outside Word (host=${info.host}); this pane is Word-only.`);
    return;
  }
  document.getElementById("copy-json").addEventListener("click", copyJson);
  document.getElementById("ping-bridge").addEventListener("click", pingBridge);
  document.getElementById("ping-bridge").disabled = false;
  // issue #84: the ops socket starts first and independently; the report
  // build is bounded and non-fatal, so neither can delay or break the other.
  connectOpsSocket();
  runAndRender();
});

// #48: handles are pane-epoch identities, not invented Office.js revision ids.
let paneEpoch = crypto.randomUUID();
const revisionHandles = new Map();
function revisionsSupported() {
  return Office.context.requirements.isSetSupported("WordApi", "1.6");
}
async function revisionObjects(context, body) {
  const changes = body.getTrackedChanges();
  changes.load("items/author,items/date,items/text,items/type");
  await context.sync();
  return changes;
}
// Word for Mac's getTrackedChanges() can merge an adjacent insertion and
// deletion into one item, so the API list undercounts. The document part's own
// revision markup is the cross-check; null means "could not be read".
const REVISION_MARKUP = /<w:(?:ins|del|moveFrom|moveTo|rPrChange|pPrChange|sectPrChange|tblPrChange|trPrChange|tcPrChange|tblGridChange|cellIns|cellDel|cellMerge|numberingChange)[\s/>]/g;
// One guarded read of the body OOXML: null means it could not be read. Shared by the
// markup count and the revision listing so a failed read never fails the whole op.
async function bodyOoxmlOrNull(context) {
  try {
    const ooxml = context.document.body.getOoxml();
    await context.sync();
    return String(ooxml.value || "");
  } catch (error) {
    return null;
  }
}
function revisionMarkupCount(xml) {
  const start = xml.indexOf('pkg:name="/word/document.xml"');
  let part = xml;
  if (start >= 0) {
    const end = xml.indexOf("</pkg:part>", start);
    part = end >= 0 ? xml.slice(start, end) : xml.slice(start);
  }
  return (part.match(REVISION_MARKUP) || []).length;
}
async function bodyRevisionMarkupCount(context) {
  const xml = await bodyOoxmlOrNull(context);
  return xml === null ? null : revisionMarkupCount(xml);
}
// WordApiDesktop 1.4 Word.Revision: Word for Mac's TrackedChange list omits deletions and merges
// neighbours, but this collection lists every content revision (deletions too) in document order, each
// with its own accept()/reject(). Handles are registered like the TrackedChange ones; `desktop` marks
// them so the mutation path loads the right fields. null = not available or unreadable (the server then
// keeps the old behaviour).
function desktopRevisionsSupported() {
  try {
    return Office.context.requirements.isSetSupported("WordApiDesktop", "1.4");
  } catch (error) {
    return false;
  }
}
async function desktopRevisionList(context) {
  if (!desktopRevisionsSupported()) return null;
  try {
    const collection = context.document.body.getRange("Whole").revisions;
    collection.load("items/type,items/author,items/date");
    await context.sync();
    const ranges = collection.items.map(r => {const g = r.range; g.load("text"); return g;});
    await context.sync();
    collection.track();
    return collection.items.map((revision, i) => {
      const id = `revision:${paneEpoch}:${crypto.randomUUID()}`;
      revision.track();
      revisionHandles.set(id, {change: revision, collection, epoch: paneEpoch, desktop: true});
      return {revision_id: id, type: revision.type, author: revision.author, date: revision.date, text: ranges[i].text};
    });
  } catch (error) {
    return null;
  }
}
async function opRevisionsList() {
  if (!revisionsSupported()) return {revisions: null, coverage: "unavailable", reason: "WordApi 1.6 required"};
  return measuredWordRun(async context => {
    const changes = await revisionObjects(context, context.document.body);
    const entries = changes.items.map(change => {
      const id = `revision:${paneEpoch}:${crypto.randomUUID()}`;
      change.track();
      changes.track();
      const range = change.getRange();
      range.paragraphs.load("items/text");
      revisionHandles.set(id, {change, collection: changes, epoch: paneEpoch});
      return {id, change, range};
    });
    await context.sync();
    const desktop = await desktopRevisionList(context);
    const xml = await bodyOoxmlOrNull(context);
    const markup = xml === null ? null : revisionMarkupCount(xml);
    // An unreadable body OOXML must not fail the listing: omit revision_ooxml and the
    // server keeps the Office.js inventory (ooxml_revision_count null = unverified).
    return {...(desktop ? {desktop_revisions: desktop} : {}), ...(xml === null ? {} : {revision_ooxml: xml}), coverage: "body", session_epoch: paneEpoch, ooxml_revision_count: markup, revisions: entries.map(({id, change, range}) => ({
      revision_id: id, type: change.type, author: change.author, date: change.date,
      text: change.text, paragraph_context: range.paragraphs.items.map(p => p.text), scope: "body"
    }))};
  });
}
async function guardRevisions(context, ranges) {
  if (!revisionsSupported()) throw refusalError("Cannot inspect revisions: WordApi 1.6 required", "LIVE_CAPABILITY_MISSING");
  const refuse = () => refusalError("Accept or reject intersecting revisions before replacing text", "TRACKED_CHANGES_PRESENT");
  const collections = ranges.map(r => r.getTrackedChanges());
  collections.forEach(c => c.load("items"));
  await context.sync();
  if (collections.some(c => c.items.length)) throw refuse();
  // Word for Mac's range.getTrackedChanges() only returns changes fully inside
  // the range, so a match that sits inside (or straddles) a larger revision
  // reports nothing. Compare against every revision in the match's paragraphs.
  const scoped = ranges.filter(r => r.paragraphs && typeof r.paragraphs.getFirst === "function" && typeof r.compareLocationWith === "function");
  const nearby = scoped.map(r => {
    const first = r.paragraphs.getFirst(), last = r.paragraphs.getLast();
    const changes = first.getRange("Start").expandTo(last.getRange("End")).getTrackedChanges();
    changes.load("items");
    return changes;
  });
  await context.sync();
  const comparisons = [];
  nearby.forEach((changes, i) => changes.items.forEach(change => comparisons.push(scoped[i].compareLocationWith(change.getRange()))));
  if (comparisons.length) await context.sync();
  const apart = new Set(["Before", "After", "AdjacentBefore", "AdjacentAfter"]);
  if (comparisons.some(c => !apart.has(c.value))) throw refuse();
}
// issue #79: revision markup split into "inside the Nth top-level table" and "everywhere else", read
// from the body OOXML (the API list undercounts). Outside revisions are compared as a multiset of
// kind|author|text, so ids and dates Word renumbers do not matter.
const REVISION_KINDS = ["ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange", "sectPrChange", "tblPrChange",
  "trPrChange", "tcPrChange", "tblGridChange", "cellIns", "cellDel", "cellMerge", "numberingChange"];
function tableRevisionInventory(xml, tableIndex) {
  const body = shapeXml(xml);
  const tables = [...body.getElementsByTagNameNS(SHAPE_W, "tbl")].filter(t => {
    for (let p = t.parentNode; p && p !== body; p = p.parentNode)
      if (p.namespaceURI === SHAPE_W && p.localName === "tbl") return false;
    return true;
  });
  const target = tables[tableIndex - 1] || null;
  const insideOf = node => {
    for (let p = node.parentNode; p; p = p.parentNode) if (p === target) return true;
    return false;
  };
  let inside = 0;
  const outside = [];
  REVISION_KINDS.forEach(kind => [...body.getElementsByTagNameNS(SHAPE_W, kind)].forEach(node => {
    if (target && insideOf(node)) inside += 1;
    else outside.push(`${kind}|${node.getAttribute("w:author") || ""}|${node.textContent || ""}`);
  }));
  return {tableCount: tables.length, inside, outside: outside.sort()};
}
async function opRevisionsMutateScoped(payload, action) {
  return measuredWordRun(async context => {
    await requireFreshBody(context, payload.expectedBodySha256);
    const tableItems = await loadTopLevelTables(context);
    const table = tableItems[payload.table_index - 1];
    if (!table) throw refusalError(`no table ${payload.table_index} (the document has ${tableItems.length})`, "table_not_found");
    const range = table.getRange();
    // Fail closed BEFORE any write: a pane that cannot scope or cannot verify must never run it.
    if (typeof range.getTrackedChanges !== "function")
      throw refusalError("This host cannot scope tracked changes to a table; nothing was changed", "LIVE_CAPABILITY_MISSING");
    const xmlBefore = await bodyOoxmlOrNull(context);
    if (xmlBefore === null)
      throw refusalError("Scoped revision changes need the body OOXML, which could not be read; nothing was changed", "LIVE_CAPABILITY_MISSING");
    const before = tableRevisionInventory(xmlBefore, payload.table_index);
    const scoped = range.getTrackedChanges();
    scoped[`${action}All`]();
    await context.sync();
    const xmlAfter = await bodyOoxmlOrNull(context);
    context.document.body.load("text");
    await context.sync();
    const post = await sha256Hex(context.document.body.text || "");
    if (xmlAfter === null)
      throw refusalError(`Scoped ${action} was sent but the body could not be re-read; it may have applied. pre=${before.tableCount} table(s)`, "VERIFICATION_FAILED");
    // Rejecting a tracked table insertion removes the table, which is the expected outcome.
    const tableRemoved = tableRevisionInventory(xmlAfter, 0).tableCount === before.tableCount - 1;
    // Index 0 matches no table, so every remaining revision counts as "outside".
    const after = tableRevisionInventory(xmlAfter, tableRemoved ? 0 : payload.table_index);
    const result = {applied: true, scoped: true, table_index: payload.table_index, table_removed: tableRemoved,
      before_count: before.inside, after_count: tableRemoved ? 0 : after.inside, changed_count: before.inside,
      outside_count: before.outside.length, tables_before: before.tableCount, tables_after: after.tableCount, post};
    const verified = (tableRemoved || after.tableCount === before.tableCount) && result.after_count === 0 &&
      JSON.stringify(after.outside) === JSON.stringify(before.outside);
    if (!verified) throw refusalError(JSON.stringify(result), "VERIFICATION_FAILED");
    return result;
  });
}
async function opRevisionsMutate(payload, action) {
  const scoped = payload.table_index !== undefined && payload.table_index !== null;
  if (scoped && payload.revision_ids != null)
    throw refusalError("table_index and revision_ids are mutually exclusive", "INVALID_INPUT");
  if (scoped) {
    if (!revisionsSupported()) throw refusalError("WordApi 1.6 required", "LIVE_CAPABILITY_MISSING");
    refuseDocumentWide(payload, "Revision changes");
    return opRevisionsMutateScoped(payload, action);
  }
  const retained = payload.revision_ids == null ? [] : payload.revision_ids.map(id => {
    const entry = revisionHandles.get(id);
    if (!entry || entry.epoch !== paneEpoch) throw refusalError("Stale revision handle; list revisions again", "REVISION_ID_NOT_FOUND");
    return entry.change;
  });
  if (!revisionsSupported()) throw refusalError("WordApi 1.6 required", "LIVE_CAPABILITY_MISSING");
  refuseDocumentWide(payload, "Revision changes");
  const execute = async context => {
    await requireFreshBody(context, payload.expectedBodySha256);
    const changes = await revisionObjects(context, context.document.body);
    const before = changes.items.length;
    const markupBefore = await bodyRevisionMarkupCount(context);
    const all = payload.revision_ids == null;
    const selected = all ? changes.items : payload.revision_ids.map(id => {
      const entry = revisionHandles.get(id);
      if (!entry || entry.epoch !== paneEpoch) throw refusalError("Stale revision handle; list revisions again", "REVISION_ID_NOT_FOUND");
      return entry.change;
    });
    // Validate every retained object before queuing any mutation (a desktop Revision has no `text`).
    const desktopHandles = new Set((payload.revision_ids || []).filter(id => (revisionHandles.get(id) || {}).desktop)
      .map(id => revisionHandles.get(id).change));
    selected.forEach(c => c.load(desktopHandles.has(c) ? "author,date,type" : "author,date,text,type"));
    await context.sync();
    // One collection call covers items Word's list merged; per-item calls do not.
    if (all && typeof changes[`${action}All`] === "function") changes[`${action}All`]();
    else selected.forEach(c => c[action]());
    await context.sync();
    const remaining = await revisionObjects(context, context.document.body);
    const markupAfter = await bodyRevisionMarkupCount(context);
    context.document.body.load("text");
    await context.sync();
    const result = {applied: true, before_count: before, after_count: remaining.items.length,
      changed_count: selected.length, ooxml_before: markupBefore, ooxml_after: markupAfter,
      post: await sha256Hex(context.document.body.text || "")};
    // The API list can undercount, so progress is judged on the markup itself.
    const verified = all ? remaining.items.length === 0 && !markupAfter
      : markupBefore != null && markupAfter != null ? markupAfter < markupBefore
        : result.after_count === before - selected.length;
    if (!verified) throw refusalError(JSON.stringify(result), "VERIFICATION_FAILED");
    revisionHandles.forEach(e => {e.change.untrack(); e.collection.untrack();});
    revisionHandles.clear();
    await context.sync();
    return result;
  };
  return retained.length ? measuredWordRun(retained, execute) : measuredWordRun(execute);
}

function checkCommentEpoch(payload) {
  if (payload.session_epoch && payload.session_epoch !== paneEpoch)
    throw refusalError("Comment handle is from another pane epoch", "COMMENT_ID_STALE");
}
async function verifyCommentIdentity(context, comment, identity) {
  if (!identity) return;
  comment.load("content,authorName");
  const range = comment.getRange();
  range.load("text");
  range.paragraphs.load("items/text");
  await context.sync();
  const actual = {content: comment.content, authorName: comment.authorName,
    anchorText: range.text, anchorParagraphText: range.paragraphs.items.map(p => p.text).join("\n")};
  const norm = x => String(x || "").replace(/\s+/g, " ").trim();
  if (Object.keys(identity).some(k => norm(actual[k]) !== norm(identity[k])))
    throw refusalError("Comment identity changed; re-list before editing", "COMMENT_ID_STALE");
}

// Word for Mac 16.113.3 throws GeneralException for every comment lookup on a range or body
// inside shape text (range.getComments, parentBody.getComments), and Word does not keep
// comments in text boxes at all (a comment anchor patched into a text box was gone from the
// live OOXML after load). So for shape scopes the comment check reads the shape body's own
// OOXML for comment marks instead, and fails closed when that cannot be read.
async function shapeCommentIds(context, body) {
  let xml;
  try {
    const ooxml = body.getOoxml();
    await context.sync();
    xml = String(ooxml.value || "");
  } catch (error) {
    throw refusalError("Cannot verify comments in shape text on this host (body OOXML unreadable)", "LIVE_CAPABILITY_MISSING");
  }
  return [...new Set([...xml.matchAll(/<w:comment(?:RangeStart|RangeEnd|Reference)\b[^>]*\bw:id="([^"]*)"/g)].map(m => m[1]))];
}
async function guardShapeComments(context, shapeBodies, allowLoss, allowIds) {
  const found = [];
  for (const body of shapeBodies) found.push(...await shapeCommentIds(context, body));
  // Shape scopes only know comment ids (no resolved state or reply counts), so per-thread consent cannot be checked.
  if (found.length && Array.isArray(allowIds))
    throw refusalError(`allow_comment_loss_ids is not supported for text-box scopes; comments in shape text: ${JSON.stringify(found)}`, "WOULD_DELETE_COMMENTS");
  if (found.length && allowLoss !== true)
    throw refusalError(`Replacement overlaps comments in shape text: ${JSON.stringify(found)}`, "WOULD_DELETE_COMMENTS");
  return found.map(id => ({id}));
}
// The text-box bodies a scoped op addresses (null for the main body).
async function scopeShapeBodies(context, scope) {
  if (!scope || scope === "body") return null;
  return (await resolveScopes(context, scope, true)).bodies.filter(b => b.scope !== "body").map(b => b.body);
}

// issue #82: per-thread consent. Every overlapping thread must be listed, resolved, and have the listed reply count.
function enforceCommentLossIds(affected, allowIds) {
  for (const entry of allowIds)
    if (!entry || typeof entry.comment_id !== "string" || !Number.isInteger(entry.reply_count) || entry.reply_count < 0)
      throw refusalError("allow_comment_loss_ids entries must be {comment_id, reply_count}", "INVALID_INPUT");
  const byId = new Map(allowIds.map(e => [e.comment_id, e]));
  for (const c of affected) {
    const entry = byId.get(c.id);
    if (!entry)
      throw refusalError(`Replacement overlaps comment ${c.id} which is not in allow_comment_loss_ids`, "WOULD_DELETE_COMMENTS");
    if (c.resolved !== true)
      throw refusalError(`Comment ${c.id} is listed in allow_comment_loss_ids but is not resolved`, "WOULD_DELETE_COMMENTS");
    if (c.replies.length !== entry.reply_count)
      throw refusalError(`Comment ${c.id} has ${c.replies.length} reply(ies), allow_comment_loss_ids expected ${entry.reply_count}`, "WOULD_DELETE_COMMENTS");
  }
}

async function guardComments(context, ranges, allowLoss, shapeBodies, allowIds) {
  if (Array.isArray(allowIds) && allowLoss === true)
    throw refusalError("allow_comment_loss and allow_comment_loss_ids are mutually exclusive", "INVALID_INPUT");
  if (shapeBodies && shapeBodies.length) return guardShapeComments(context, shapeBodies, allowLoss, allowIds);
  const affected = new Map();
  for (const range of ranges) {
    const comments = range.getComments();
    comments.load("items/id,items/content,items/authorName,items/resolved");
    await context.sync();
    const entries = comments.items.map(comment => {
      const anchor = comment.getRange();
      anchor.load("text");
      comment.replies.load("items/id,items/content,items/authorName");
      return {comment, anchor, relation: anchor.compareLocationWith(range)};
    });
    if (entries.length) await context.sync();
    const disjoint = new Set(["Before", "After", "AdjacentBefore", "AdjacentAfter"]);
    entries.filter(e => !disjoint.has(e.relation.value)).forEach(e => affected.set(e.comment.id, {
      id: e.comment.id, content: e.comment.content, authorName: e.comment.authorName,
      resolved: e.comment.resolved, anchorText: e.anchor.text,
      replies: e.comment.replies.items.map(r => ({id: r.id, content: r.content, authorName: r.authorName}))
    }));
  }
  const result = [...affected.values()];
  if (Array.isArray(allowIds)) {
    enforceCommentLossIds(result, allowIds);
    return result;
  }
  if (result.length && allowLoss !== true)
    throw refusalError(`Replacement overlaps comments: ${JSON.stringify(result)}`, "WOULD_DELETE_COMMENTS");
  return result;
}

const FONT_FIELDS = ["bold", "italic", "underline", "strikeThrough", "color", "name", "size",
  "doubleStrikeThrough", "subscript", "superscript"];
function fontSnapshot(font) {
  return Object.fromEntries(FONT_FIELDS.map(k => [k, font[k]]));
}

async function resolveScopes(context, scope, writing=false) {
  if (scope === 'body') return {bodies: [{scope: 'body', body: context.document.body}], warnings: []};
  if (!Office.context.requirements.isSetSupported('WordApiDesktop', '1.2'))
    throw refusalError('Text boxes require WordApiDesktop 1.2', 'LIVE_CAPABILITY_MISSING');
  if (scope !== 'all' && !scope.startsWith(`textbox:${paneEpoch}:`))
    throw refusalError('Stale or invalid text-box handle; list again', 'LIVE_STALE');
  try {
    const shapes = context.document.body.shapes;
    shapes.load('items/id,items/name,items/type');
    await context.sync();
    const warnings = shapes.items.filter(s => s.type === Word.ShapeType.group || s.type === Word.ShapeType.canvas)
      .map(s => ({shape_id: s.id, reason: 'group/canvas text coverage unavailable'}));
    if (scope === 'all' && writing && warnings.length)
      throw refusalError('Cannot safely write all scopes with incomplete shape coverage', 'LIVE_CAPABILITY_MISSING');
    const selected = shapes.items.filter(s => s.type === Word.ShapeType.textBox || s.type === Word.ShapeType.geometricShape)
      .filter(s => scope === 'all' || `textbox:${paneEpoch}:${s.id}` === scope);
    if (scope !== 'all' && selected.length !== 1) throw refusalError('Text box no longer exists', 'LIVE_STALE');
    selected.forEach(s => s.body.load('text'));
    await context.sync();
    const bodies = selected.map(s => {
      if (typeof s.body.text !== 'string') throw new Error('host returned missing text');
      return {scope: `textbox:${paneEpoch}:${s.id}`, body: s.body, shape: {id:s.id,name:s.name,type:s.type}};
    });
    if (scope === 'all') bodies.unshift({scope:'body',body:context.document.body});
    return {bodies,warnings};
  } catch (err) {
    if (err.code === 'LIVE_STALE' || err.code === 'LIVE_CAPABILITY_MISSING') throw err;
    throw refusalError(`Shape read failed (retryable): ${err.message || err}`, 'HOST_SHAPE_READ_FAILED');
  }
}
async function scopeHash(context, scope) {
  const resolved = await resolveScopes(context, scope);
  resolved.bodies.forEach(b => b.body.load('text'));
  await context.sync();
  return sha256Hex(JSON.stringify(resolved.bodies.map(b => [b.scope,b.body.text])));
}
async function requireFreshScope(context, payload) {
  if (payload.expectedScopeSha256 && await scopeHash(context,payload.scope || 'body') !== payload.expectedScopeSha256)
    throw refusalError('Scope changed since it was read; nothing written', 'LIVE_STALE');
}
async function opScopeDescribe(payload) {
  return measuredWordRun(async context => ({scopeSha256:await scopeHash(context,payload.scope || 'body'),
    bodySha256:await requireFreshBody(context),session_epoch:paneEpoch}));
}
async function opTextboxesList() {
  return measuredWordRun(async context => {
    const resolved = await resolveScopes(context,'all');
    const textboxes = [];
    for (const item of resolved.bodies.filter(b => b.shape)) {
      textboxes.push({textbox_id:item.scope,text:item.body.text,shape:item.shape,
        revision:`live:scope:${await scopeHash(context,item.scope)}`});
    }
    return {textboxes,session_epoch:paneEpoch,coverage:resolved.warnings.length ? 'partial':'body_shapes',warnings:resolved.warnings};
  });
}
async function opTextboxesRead(payload) {
  return measuredWordRun(async context => {
    const resolved = await resolveScopes(context,payload.scope);
    const item = resolved.bodies[0];
    return {textbox_id:item.scope,text:item.body.text,shape:item.shape,session_epoch:paneEpoch,
      revision:`live:scope:${await scopeHash(context,payload.scope)}`};
  });
}

function inspectDeleteParagraph(xml) {
  const doc = new DOMParser().parseFromString(xml,'application/xml');
  const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  if (doc.getElementsByTagName('parsererror').length)
    throw refusalError('Paragraph OOXML inspection failed', 'STRUCTURAL_BOUNDARY');
  const bodies = [...doc.getElementsByTagNameNS(w,'body')];
  if (bodies.length !== 1) throw refusalError('Incomplete paragraph body inspection', 'STRUCTURAL_BOUNDARY');
  if (shapeNodes(bodies[0]).length)
    throw refusalError('Paragraph anchors a shape; re-anchoring is not supported', 'ANCHORED_SHAPES');
  const paragraphs = [...bodies[0].getElementsByTagNameNS(w,'p')];
  // Word's getOoxml appends one empty paragraph after the range (issue #33's
  // trailing empty paragraph); anything else extra means the inspection saw
  // more than the target paragraph.
  const emptyTrailer = q => !(q.textContent || '').trim() && !shapeNodes(q).length &&
    !q.getElementsByTagNameNS(w,'drawing').length && !q.getElementsByTagNameNS(w,'pict').length &&
    !q.getElementsByTagNameNS(w,'sectPr').length;
  if (paragraphs.length < 1 || paragraphs.length > 2 || (paragraphs.length === 2 && !emptyTrailer(paragraphs[1])))
    throw refusalError('Incomplete paragraph inspection', 'STRUCTURAL_BOUNDARY');
  const p = paragraphs[0];
  if (p.getElementsByTagNameNS(w,'sectPr').length)
    throw refusalError('Paragraph is a section boundary', 'STRUCTURAL_BOUNDARY');
}
async function opParagraphDelete(payload) {
  return measuredWordRun(async context => {
    const pre = await mutationHash(context, payload);
    const source = payload.executionTargetSection ? await targetSectionRange(context, payload) : context.document.body;
    const paragraphs = source.paragraphs;
    paragraphs.load('items/text,items/tableNestingLevel');
    await context.sync();
    const before = paragraphs.items.map(p => p.text);
    const indexes = before.map((text,i) => text === payload.anchor_text ? i : -1).filter(i => i >= 0);
    if (indexes.length !== 1) throw refusalError(`Expected unique whole paragraph, found ${indexes.length}`, 'MATCH_COUNT_MISMATCH');
    const index = indexes[0];
    const target = paragraphs.items[index];
    // issue #107: a paragraph inside a (non-nested) table cell may go, but never the cell's last one.
    if (target.tableNestingLevel > 1)
      throw refusalError('Cannot delete a paragraph inside a nested table', 'STRUCTURAL_BOUNDARY');
    if (target.tableNestingLevel === 1) {
      const cellParagraphs = target.parentTableCell.body.paragraphs;
      cellParagraphs.load('items/text');
      await context.sync();
      // Real Word (16.114.3): a cell's FINAL paragraph cannot be removed -- delete() blanks its text and
      // leaves an empty paragraph, so the write would land and then fail read-back. Refuse before writing.
      // `anchor_text` is unique in the whole body (checked above), so equal text identifies the paragraph.
      const last = cellParagraphs.items[cellParagraphs.items.length-1];
      if (cellParagraphs.items.length <= 1 || (last && last.text === payload.anchor_text))
        throw refusalError("Cannot delete the last paragraph of a cell (Word keeps a cell's final paragraph); delete an earlier one or use merge_paragraph_with_next", 'STRUCTURAL_BOUNDARY');
    } else if (index === before.length-1) {
      throw refusalError('Cannot delete the final body paragraph', 'STRUCTURAL_BOUNDARY');
    }
    const range = target.getRange(Word.RangeLocation.whole);
    await (await sectionGuard(context, payload)).check(() => [range]);
    const xml = target.getOoxml();
    await context.sync();
    inspectDeleteParagraph(xml.value);
    await guardComments(context,[range],false);
    await guardRevisions(context,[range]);
    await withTracking(context,payload.track_changes,async () => {
      if (payload.executionSections) {
        target.load('text');
        await context.sync();
        if (target.text !== payload.anchor_text) throw refusalError('Paragraph changed during preflight', 'LIVE_STALE');
      }
      target.delete();
      await context.sync();
    });
    const afterSource = payload.executionTargetSection ? await targetSectionRange(context, payload) : context.document.body;
    const remaining = afterSource.paragraphs;
    remaining.load('items/text');
    await context.sync();
    const after = remaining.items.map(p => p.text);
    const expected = before.filter((_,i) => i !== index);
    // compact_evidence_v1 (#112): counts and the two neighbouring paragraphs, not every
    // paragraph of the section twice. The read-back below still compares the full lists here.
    const clip = text => text == null ? null : (text.length <= 200 ? text : text.slice(0, 200) + "…");
    const result = {applied:true,before_count:before.length,after_count:after.length,
      neighbours:{previous:clip(index > 0 ? before[index-1] : null), next:clip(index+1 < before.length ? before[index+1] : null)},
      pre,post:await mutationHash(context, {...payload, expectedBodySha256:undefined})};
    if (JSON.stringify(after) !== JSON.stringify(expected))
      // The uncertain-outcome path needs the full lists to reconcile, so they stay in this message only.
      throw refusalError(`Paragraph deletion read-back failed; observed effects: ${JSON.stringify({...result, before_paragraphs:before, after_paragraphs:after})}`, 'VERIFICATION_FAILED');
    return result;
  });
}

// -- clear_numbering / delete_table (issue #79) ------------------------------

// Detaches every list-item paragraph in one table (table_index) or in the single paragraph
// whose whole text equals paragraph_text. Detaching changes no text, so the body hash must not move.
async function opParagraphsUnlist(payload) {
  return measuredWordRun(async context => {
    const pre = await requireFreshBody(context, payload.expectedBodySha256);
    const byTable = payload.table_index !== undefined && payload.table_index !== null;
    let range, loadScope;
    if (byTable) {
      const tableItems = await loadTopLevelTables(context);
      const table = tableItems[payload.table_index - 1];
      if (!table) throw refusalError(`no table ${payload.table_index} (the document has ${tableItems.length})`, 'table_not_found');
      range = table.getRange();
      loadScope = async () => {
        const paragraphs = table.getRange().paragraphs;
        paragraphs.load('items/isListItem');
        await context.sync();
        return paragraphs.items;
      };
    } else {
      const all = context.document.body.paragraphs;
      all.load('items/text');
      await context.sync();
      const hits = all.items.filter(p => p.text === payload.paragraph_text);
      if (hits.length !== 1) throw refusalError(`Expected unique whole paragraph, found ${hits.length}`, 'MATCH_COUNT_MISMATCH');
      const target = hits[0];
      range = target.getRange(Word.RangeLocation.whole);
      loadScope = async () => {
        target.load('isListItem');
        await context.sync();
        return [target];
      };
    }
    await (await sectionGuard(context, payload)).check(() => [range]);
    await guardRevisions(context, [range]);
    const listed = (await loadScope()).filter(p => p.isListItem);
    listed.forEach(p => p.detachFromList());
    if (listed.length) await context.sync();
    const remaining = (await loadScope()).filter(p => p.isListItem).length;
    context.document.body.load('text');
    await context.sync();
    const post = await sha256Hex(context.document.body.text || '');
    const result = {applied: true, unlisted: listed.length, remaining_list_items: remaining, pre, post};
    if (remaining || post !== pre)
      throw refusalError(`Numbering removal read-back failed; observed effects: ${JSON.stringify(result)}`, 'VERIFICATION_FAILED');
    return result;
  });
}

// Whole-table text as get_table reports it (cells joined by tab, rows by newline); the caller's
// expected_text is compared whitespace-insensitively through normalizeAnchorText.
function tableTextOf(grid) {
  return grid.map(cells => cells.map(c => c.body.text || '').join('\t')).join('\n');
}

async function opTableDelete(payload) {
  return measuredWordRun(async context => {
    const pre = await requireFreshBody(context, payload.expectedBodySha256);
    const tableItems = await loadTopLevelTables(context);
    const table = tableItems[payload.table_index - 1];
    if (!table) throw refusalError(`no table ${payload.table_index} (the document has ${tableItems.length})`, 'table_not_found');
    const grid = await loadRowsAndCells(context, table);
    grid.forEach(cells => cells.forEach(cell => cell.body.load('text')));
    await context.sync();
    const actual = tableTextOf(grid);
    if (normalizeAnchorText(actual) !== normalizeAnchorText(payload.expected_text))
      throw refusalError(`Table ${payload.table_index} text does not match expected_text; nothing was deleted. Table reads: ${JSON.stringify(actual)}`, 'CELL_TEXT_MISMATCH');
    const range = table.getRange();
    await (await sectionGuard(context, payload)).check(() => [range]);
    const xml = range.getOoxml();
    await context.sync();
    if (shapeNodes(shapeXml(xml.value)).length)
      throw refusalError('Table anchors a shape; re-anchoring is not supported', 'ANCHORED_SHAPES');
    await guardComments(context, [range], false);
    await guardRevisions(context, [range]);
    const outside = async () => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load(payload.executionTargetSection ? 'items/tableNestingLevel' : 'items/text,items/tableNestingLevel');
    if (payload.executionTargetSection) {
      const scoped = await loadSectionMap(context, true);
      const target = scoped.sections.find(s => s.slug === payload.executionTargetSection);
      if (!target) throw refusalError('Missing target section', 'LOCK_SCOPE_UNRESOLVED');
      await context.sync();
      paragraphs.items.slice(target.start,target.end).forEach(p => p.load('text'));
    }
      await context.sync();
      return paragraphs.items.filter(p => p.tableNestingLevel === 0).map(p => p.text);
    };
    const outsideBefore = await outside();
    await withTracking(context, payload.track_changes, async () => {
      table.delete();
      await context.sync();
    });
    const tablesAfter = (await loadTopLevelTables(context)).length;
    const outsideAfter = await outside();
    context.document.body.load('text');
    await context.sync();
    const result = {applied: true, table_index: payload.table_index, tables_before: tableItems.length, tables_after: tablesAfter,
      deleted_text: actual, tracked: !!payload.track_changes, pre, post: await sha256Hex(context.document.body.text || '')};
    // A tracked table deletion leaves the (struck-through) table in the document until it is accepted.
    const countOk = payload.track_changes ? tablesAfter === tableItems.length || tablesAfter === tableItems.length - 1
      : tablesAfter === tableItems.length - 1;
    if (!countOk || JSON.stringify(outsideBefore) !== JSON.stringify(outsideAfter))
      throw refusalError(`Table deletion read-back failed; observed effects: ${JSON.stringify(result)}`, 'VERIFICATION_FAILED');
    return result;
  });
}

// -- live paragraph insertion (issue #77) ------------------------------------

// Resolves one paragraph spec's style BEFORE any write, so an unknown style
// never leaves a half-inserted paragraph behind. Returns {kind, value}: kind
// "name" is assigned to Paragraph.style, kind "builtin" to styleBuiltIn.
async function resolveParagraphStyle(context, spec) {
  const builtins = Word.BuiltInStyleName || {};
  if (!spec.style) return { kind: "builtin", value: builtins.normal || "Normal" };
  if (Office.context.requirements.isSetSupported("WordApi", "1.5")) {
    const found = context.document.getStyles().getByNameOrNullObject(spec.style);
    found.load("isNullObject,type");
    await context.sync();
    if (!found.isNullObject) {
      if (found.type !== "Paragraph")
        throw refusalError(`style ${JSON.stringify(spec.style)} is a ${found.type} style; paragraph styles only`, "UNSUPPORTED_STYLE_TYPE");
      return { kind: "name", value: spec.style };
    }
  }
  const squashed = spec.style.replace(/\s+/g, "").toLowerCase();
  const key = Object.keys(builtins).find(k => k.toLowerCase() === squashed || String(builtins[k]).toLowerCase() === squashed);
  if (key) return { kind: "builtin", value: builtins[key] };
  throw refusalError(`paragraph style ${JSON.stringify(spec.style)} was not found in the document`, "STYLE_NOT_FOUND");
}

async function opParagraphInsert(payload) {
  if (payload.anchor && typeof payload.anchor === "object") return opParagraphInsertPositional(payload);
  return measuredWordRun(async context => {
    const pre = await requireFreshBody(context, payload.expectedBodySha256);
    const specs = payload.paragraphs || [];
    if (!specs.length) throw refusalError("No paragraphs to insert", "INVALID_INPUT");
    const styles = [];
    for (const spec of specs) styles.push(await resolveParagraphStyle(context, spec));

    const found = context.document.body.search(payload.anchor, { matchCase: true, matchWholeWord: false });
    found.load("items");
    await context.sync();
    if (found.items.length !== payload.expected_matches)
      throw refusalError(`expected ${payload.expected_matches} match(es) for ${JSON.stringify(payload.anchor)}, found ${found.items.length}`,
        found.items.length === 0 ? "ZERO_MATCH" : "MATCH_COUNT_MISMATCH");

    const anchorPara = found.items[0].paragraphs.getFirst();
    anchorPara.load("text,tableNestingLevel");
    const all = context.document.body.paragraphs;
    all.load("items/text,items/tableNestingLevel");
    await context.sync();
    if (anchorPara.tableNestingLevel !== 0)
      throw refusalError("The anchor is inside a table cell; anchor on a body paragraph", "STRUCTURAL_BOUNDARY");
    const before = all.items.map(q => q.text);
    let candidates = before.map((t, i) => (t === anchorPara.text ? i : -1)).filter(i => i >= 0);
    if (candidates.length > 1) {
      const rels = candidates.map(i => all.items[i].getRange().compareLocationWith(anchorPara.getRange()));
      await context.sync();
      candidates = candidates.filter((_, j) => rels[j].value === "Equal");
    }
    if (candidates.length !== 1)
      throw refusalError("Could not resolve the anchor to exactly one body paragraph", "STRUCTURAL_BOUNDARY");
    const index = candidates[0];
    const position = payload.position === "before" ? "before" : "after";

    const guard = await sectionGuard(context, payload);
    if (guard.active) {
      if (position === "before")
        throw refusalError("Under section locks insert paragraphs after an anchor, not before it", "OUTSIDE_LOCKED_SECTION");
      await guard.check([anchorPara.getRange()]);
    }

    const inserted = [];
    await withTracking(context, payload.track_changes, async () => {
      let previous = null;
      specs.forEach((spec, i) => {
        const para = previous
          ? previous.insertParagraph(spec.text, Word.InsertLocation.after)
          : anchorPara.insertParagraph(spec.text, position === "before" ? Word.InsertLocation.before : Word.InsertLocation.after);
        if (styles[i].kind === "name") para.style = styles[i].value;
        else para.styleBuiltIn = styles[i].value;
        if (spec.color) para.font.color = withHash(spec.color);
        inserted.push(para);
        previous = para;
      });
      await context.sync();
    });

    inserted.forEach(para => { para.load("text,style,styleBuiltIn"); para.font.load("color"); });
    const fresh = context.document.body.paragraphs;
    fresh.load("items/text");
    context.document.body.load("text");
    await context.sync();

    const observed = inserted.map((para, i) => ({
      text: String(para.text || "").replace(/[\r\n]+$/, ""),
      style: para.style,
      style_builtin: para.styleBuiltIn,
      color: para.font.color,
      requested_style: styles[i],
    }));
    const at = position === "before" ? index : index + 1;
    const expected = [...before.slice(0, at), ...specs.map(spec => spec.text), ...before.slice(at)];
    const after = fresh.items.map(q => q.text);
    const problems = [];
    observed.forEach((o, i) => {
      if (o.text !== specs[i].text) problems.push(`paragraph ${i + 1} text`);
      const styleOk = styles[i].kind === "name" ? o.style === styles[i].value : o.style_builtin === styles[i].value;
      if (!styleOk) problems.push(`paragraph ${i + 1} style`);
      if (specs[i].color && String(o.color || "").toLowerCase() !== withHash(specs[i].color).toLowerCase())
        problems.push(`paragraph ${i + 1} color`);
    });
    if (JSON.stringify(after) !== JSON.stringify(expected)) problems.push("paragraph list");
    if (problems.length) {
      const detail = `${problems.join(", ")}; observed: ${JSON.stringify(observed)}`;
      if (payload.track_changes)
        throw refusalError(`Paragraph insertion did not read back as written (${detail}); the insertions are tracked changes -- reject them in Word's Review pane`, "VERIFICATION_FAILED");
      let rolledBack = false;
      try {
        inserted.forEach(para => para.delete());
        await context.sync();
        const restored = context.document.body.paragraphs;
        restored.load("items/text");
        await context.sync();
        rolledBack = JSON.stringify(restored.items.map(q => q.text)) === JSON.stringify(before);
      } catch (err) {
        rolledBack = false;
      }
      throw refusalError(`Paragraph insertion did not read back as written (${detail}); ` +
        (rolledBack ? "the inserted paragraphs were removed" : "the inserted paragraphs could NOT be removed -- fix the document by hand"),
        "VERIFICATION_FAILED");
    }
    return { applied: true, position, anchor_index: index + 1, inserted: observed,
      before_count: before.length, after_count: after.length, pre, post: await sha256Hex(context.document.body.text || "") };
  });
}

// Shape inventory reads OOXML, so it covers DrawingML and legacy VML without
// requiring the newer desktop-only Shape API. Indices include table paragraphs.
const SHAPE_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const SHAPE_WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const SHAPE_V = 'urn:schemas-microsoft-com:vml';
function shapeXml(xml) {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length)
    throw refusalError('Cannot inspect shape anchors', 'HOST_SHAPE_READ_FAILED');
  const bodies = [...doc.getElementsByTagNameNS(SHAPE_W, 'body')];
  if (bodies.length !== 1)
    throw refusalError('Incomplete shape body inspection', 'HOST_SHAPE_READ_FAILED');
  return bodies[0];
}
function shapeNodes(root) {
  // Ignore the fallback representation of an AlternateContent choice.
  const mc = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
  return [...root.getElementsByTagNameNS(SHAPE_WP, 'anchor'),
    ...['shape','rect','roundrect','oval','line','polyline','arc','curve','image','group']
      .flatMap(tag => [...root.getElementsByTagNameNS(SHAPE_V, tag)])].filter(node => {
      for (let p = node.parentNode; p && p !== root; p = p.parentNode)
        if (p.namespaceURI === mc && p.localName === 'Fallback' &&
            p.parentNode.getElementsByTagNameNS(mc, 'Choice').length) return false;
      return true;
    });
}
function inventoryShapes(xml) {
  const root = shapeXml(xml);
  const paragraphs = [...root.getElementsByTagNameNS(SHAPE_W, 'p')];
  return shapeNodes(root).map(node => {
    let p = node.parentNode;
    while (p && !(p.namespaceURI === SHAPE_W && p.localName === 'p')) p = p.parentNode;
    const first = tag => node.getElementsByTagNameNS(SHAPE_WP, tag)[0];
    const props = first('docPr');
    const extent = first('extent');
    const position = tag => {
      const el = first(tag);
      return el ? {relative_to:el.getAttribute('relativeFrom'), value:el.textContent} : null;
    };
    return {id: props ? props.getAttribute('id') : node.getAttribute('id'),
      name: props ? props.getAttribute('name') : null,
      kind:node.namespaceURI === SHAPE_V ? 'vml' : 'drawingml',
      anchor_paragraph:{index:paragraphs.indexOf(p)+1,
        text:p ? [...p.getElementsByTagNameNS(SHAPE_W,'t')].filter(t => {
          let parent=t.parentNode;
          while(parent && parent !== p) {
            if(parent.namespaceURI === SHAPE_W && parent.localName === 'txbxContent') return false;
            parent=parent.parentNode;
          }
          return true;
        }).map(t => t.textContent).join('') : null},
      size_emu:extent ? {width:extent.getAttribute('cx'),height:extent.getAttribute('cy')} : null,
      horizontal_position:position('positionH'),vertical_position:position('positionV'),
      legacy_style:node.namespaceURI === SHAPE_V ? node.getAttribute('style') : null,
      definition:JSON.stringify(xmlSignature(node))};
  });
}
async function opExecutionShapes(payload) {
  return measuredWordRun(async context => {
    const shapes = [];
    for (const s of payload.executionSections) {
      const range = await targetSectionRange(context, {executionTargetSection:s.slug});
      const xml = range.getOoxml();
      await context.sync();
      shapes.push(...inventoryShapes(xml.value));
    }
    return {shapes, coverage:'protected_sections'};
  });
}
async function opShapesList() {
  return measuredWordRun(async context => {
    const xml = context.document.body.getOoxml();
    await context.sync();
    return {shapes:inventoryShapes(xml.value), coverage:'document_body',
      paragraph_indexing:'all body paragraphs including tables and textboxes; 1-based'};
  });
}
async function guardAnchoredBodies(context, bodies) {
  const xml = bodies.map(body => body.getOoxml());
  await context.sync();
  if (xml.some(x => shapeNodes(shapeXml(x.value)).length))
    throw refusalError('Edit touches a paragraph anchoring a shape; re-anchoring is unsupported', 'ANCHORED_SHAPES');
}
async function guardAnchoredParagraphs(context, ranges) {
  const collections = ranges.map(r => r.paragraphs);
  collections.forEach(p => p.load('items'));
  await context.sync();
  // Conservative: any text replacement in an anchored paragraph is refused.
  await guardAnchoredBodies(context, collections.flatMap(p => p.items));
}

// Compatible preflight reads are queued together, including paragraph-wide
// revision coverage (Mac omits enclosing revisions from a smaller range).
async function guardApprovedCell(context, cell, trackChanges) {
  if (!revisionsSupported()) throw refusalError('Cannot inspect revisions: WordApi 1.6 required', 'LIVE_CAPABILITY_MISSING');
  const target = cell.body.getRange('Whole');
  const comments = target.getComments();
  const changes = target.getTrackedChanges();
  const nearby = target.paragraphs.getFirst().getRange('Start')
    .expandTo(target.paragraphs.getLast().getRange('End')).getTrackedChanges();
  comments.load('items'); changes.load('items'); nearby.load('items');
  const xml = cell.body.getOoxml();
  if (trackChanges) context.document.load('changeTrackingMode');
  await context.sync();
  if (comments.items.length) await guardComments(context, [target], false);
  if (changes.items.length) throw refusalError('Accept or reject intersecting revisions before replacing text', 'TRACKED_CHANGES_PRESENT');
  const comparisons = nearby.items.map(change => target.compareLocationWith(change.getRange()));
  if (comparisons.length) await context.sync();
  const apart = new Set(['Before','After','AdjacentBefore','AdjacentAfter']);
  if (comparisons.some(c => !apart.has(c.value)))
    throw refusalError('Accept or reject intersecting revisions before replacing text', 'TRACKED_CHANGES_PRESENT');
  if (shapeNodes(shapeXml(xml.value)).length)
    throw refusalError('Edit touches a paragraph anchoring a shape', 'ANCHORED_SHAPES');
  return trackChanges ? context.document.changeTrackingMode : null;
}

// A sequential execution stream with one durable server dispatch/ack per cell.
// No section state survives a request; every mutation checks a fresh snapshot.
async function opApprovedCellSet(payload) {
  if (!payload.executionSections?.length || !payload.executionTargetSection)
    throw refusalError('Approved cell stream requires protected section preconditions', 'INVALID_INPUT');
  return measuredWordRun(async context => {
    payload={...payload,approvedCellStream:true};
    setExecutionPhase('preflight');
    let cell, previousMode, shapesBefore, before;
    try {
      cell=await resolveTableCell(context,payload);
      const target=cell.body.getRange('Whole');
      previousMode = await guardApprovedCell(context, cell, payload.track_changes);
      // This fresh protected scope check is deliberately after asynchronous guards.
      const boundary=await sectionGuard(context,payload);
      // Load exact target text in the same batch as the final location checks.
      cell.body.paragraphs.load("items/text");
      await boundary.check([target]);
      shapesBefore=boundary.scopeShapes;
      if (!Array.isArray(shapesBefore)) throw refusalError("Scoped shape baseline unavailable", "LIVE_CAPABILITY_MISSING");
      before=cell.body.paragraphs.items.map(p => p.text || "");
      if (before.map(canonicalCellText).join('\n')!==payload.expected_before_text ||
          (payload.expected_before_paragraph_count!=null && before.length!==payload.expected_before_paragraph_count))
        throw refusalError('Cell changed at mutation boundary; nothing written', 'CELL_TEXT_MISMATCH');
    } catch (error) {
      // A stale native proxy/coordinate can fail before the final scoped read.
      // This catch covers reads only: it must never wrap any queued mutation.
      if (['InvalidArgument','ItemNotFound','InvalidObjectPath'].includes(error.code))
        throw refusalError('Approved target changed during preflight; nothing was written: '+error.message, 'LIVE_STALE');
      throw error;
    }
    setExecutionPhase('mutation');
    try {
      if (previousMode!==null) context.document.changeTrackingMode=Word.ChangeTrackingMode.trackAll;
      // Exactly one write. An uncertain or failed write never switches strategies.
      await writeCellParagraphs(context,cell,payload.paragraphs,'load-items');
      // The font rides in the same batch as the write (issue #107), so it is never a second write.
      if (payload.inherit_font) applyInheritedFont(cell,payload.inherit_font);
      await context.sync();
    } finally {
      if (previousMode!==null) {context.document.changeTrackingMode=previousMode;await context.sync();}
    }
    setExecutionPhase('verification');
    // Freshly resolve the section/table/grid; no read from the mutated proxy.
    try {
    if (payload.inherit_font) await verifyInheritedFont(context,cell,payload.inherit_font);
    const post=await executionXmlSnapshot(context,payload.executionSections.map(s=>s.slug));
    const shapeIdentity=shapes=>shapes.map(({anchor_paragraph,...shape})=>JSON.stringify(shape)).sort();
    if (JSON.stringify(shapeIdentity(shapesBefore))!==JSON.stringify(shapeIdentity(post.shapes)))
      throw refusalError('Scoped shapes changed after mutation','VERIFICATION_FAILED');
    const section=post.sections.find(s=>s.slug===payload.executionTargetSection);
    const table=section?.tables.find(t=>t.table_index===payload.table_index);
    const text=table?.rows[payload.row_index-1]?.[payload.cell_index-1];
    const count=table?.paragraph_counts[payload.row_index-1]?.[payload.cell_index-1];
    const want=intendedCellParagraphs(payload.paragraphs).map(canonicalCellText);
    if (text!==want.join('\n') || count!==want.length)
      throw refusalError('Independent cell readback differs; reconcile the write', 'VERIFICATION_FAILED');
    return {applied:true,shapes_before:shapesBefore.length,shapes_after:post.shapes.length,before:before.join('\n'),after:text,
      readback:{text,paragraphs:post.cell_paragraphs[`${payload.executionTargetSection}|${payload.table_index}|${payload.row_index}|${payload.cell_index}`]},
      pre:await sha256Hex(JSON.stringify(payload.executionSections)),post:await sha256Hex(JSON.stringify(post.sections)),strategy:'approved-single-write'};
    } catch(error) {throw refusalError('Approved write may have applied; verification failed: '+(error.message || error),'VERIFICATION_FAILED');}
  });
}

let executionMetrics=null;
function setExecutionPhase(phase) {if(executionMetrics) executionMetrics.phase=phase;}
async function measuredWordRun(...args) {
  const callback=args.pop();
  return Word.run(...args,async context => {
    if (!executionMetrics) return callback(context);
    const metrics=executionMetrics, originalSync=context.sync, native=originalSync.bind(context);
    context.sync=async (...syncArgs) => {
      const start=paneNow(), phase=metrics.phase || 'operation';
      try {return await native(...syncArgs);}
      finally {
        metrics.sync_count++;
        const item=metrics.phases[phase] ||= {sync_count:0,sync_ms:0};
        item.sync_count++;item.sync_ms+=paneNow()-start;
      }
    };
    try {return await callback(context);}
    finally {context.sync=originalSync;}
  });
}
async function measuredDispatch(op,payload) {
  const start=paneNow();
  executionMetrics={sync_count:0,phases:{},phase:'preflight'};
  try {
    const result=await dispatchOp(op,payload);
    return {...result,metrics:{elapsed_ms:paneNow()-start,sync_count:executionMetrics.sync_count,phases:executionMetrics.phases}};
  } finally {executionMetrics=null;}
}

function directWordChildren(node, name) {
  return [...node.childNodes].filter(n=>n.namespaceURI===SHAPE_W && n.localName===name);
}
function visibleWordText(node) {
  if (node.namespaceURI===SHAPE_W) {
    if (['txbxContent','del','moveFrom','drawing','pict','object'].includes(node.localName)) return '';
    if (node.localName==='t') return node.textContent;
    if (node.localName==='tab') return '\t';
    if (node.localName==='br' || node.localName==='cr') return '\n';
  }
  return [...(node.childNodes || [])].filter(n=>n.nodeType===1).map(visibleWordText).join('');
}
async function executionXmlSnapshot(context, slugs) {
  const {items,sections}=await loadSectionMap(context,true);
  const chosen=slugs.map(slug=>{
    const found=sections.filter(s=>s.slug===slug);
    if(found.length!==1) throw refusalError('Missing or ambiguous protected section','LOCK_SCOPE_UNRESOLVED');
    return found[0];
  });
  const xml=chosen.map(s=>sectionRange(context,items,s).getOoxml());
  chosen.forEach(s=>items.slice(s.start,s.end).filter(p=>p.tableNestingLevel===0).forEach(p=>p.load('text')));
  await context.sync();
  const cell_paragraphs={};
  const shapes=[];
  const result=chosen.map((s,index)=>{
    const body=shapeXml(xml[index].value),tables=[];
    shapes.push(...inventoryShapes(xml[index].value).map(shape=>({...shape,section:s.slug})));
    let preceding=-1;
    for (const node of body.childNodes) {
      if (node.namespaceURI!==SHAPE_W) continue;
      if (node.localName==='p') preceding++;
      if (node.localName!=='tbl') continue;
      if (node.getElementsByTagNameNS(SHAPE_W,'tbl').length)
        throw refusalError('Nested tables are unsupported','LIVE_CAPABILITY_MISSING');
      const grid=directWordChildren(node,'tr').map(row=>directWordChildren(row,'tc').map(cell=>{
        if (cell.getElementsByTagNameNS(SHAPE_W,'vMerge').length || cell.getElementsByTagNameNS(SHAPE_W,'gridSpan').length)
          throw refusalError('Merged tables use the conservative executor','LIVE_CAPABILITY_MISSING');
        return directWordChildren(cell,'p').map(visibleWordText);
      }));
      const id=tables.length+1;
      grid.forEach((row,r)=>row.forEach((paras,c)=>{cell_paragraphs[`${s.slug}|${id}|${r+1}|${c+1}`]=paras;}));
      tables.push({table_index:id,after_paragraph:preceding,rows:grid.map(row=>row.map(paras=>paras.join('\n'))),paragraph_counts:grid.map(row=>row.map(paras=>paras.length))});
    }
    return {slug:s.slug,paragraphs:items.slice(s.start,s.end).filter(p=>p.tableNestingLevel===0).map(p=>p.text || ''),tables};
  });
  const captured={sections:result};
  Object.defineProperty(captured,'shapes',{value:shapes});
  Object.defineProperty(captured,'section_map',{value:{items,sections}});
  Object.defineProperty(captured,'cell_paragraphs',{value:cell_paragraphs});
  return captured;
}

// -- structural paragraph anchors (issue #107) -------------------------------

// {after_table_id: N} inserts after top-level table N; {empty_paragraph_after_table_id: N} fills the
// empty body paragraph directly after it. The "empty" policy is the server's list
// (payload.forbidden_empty_tags), applied to the OOXML read here, so both modes refuse the same paragraphs.
async function opParagraphInsertPositional(payload) {
  return measuredWordRun(async context => {
    const pre = await requireFreshBody(context, payload.expectedBodySha256);
    const specs = payload.paragraphs || [];
    if (!specs.length) throw refusalError("No paragraphs to insert", "INVALID_INPUT");
    const fill = payload.anchor.empty_paragraph_after_table_id !== undefined;
    const tableId = fill ? payload.anchor.empty_paragraph_after_table_id : payload.anchor.after_table_id;
    const styles = [];
    for (const spec of specs) styles.push(await resolveParagraphStyle(context, spec));
    const tables = await loadTopLevelTables(context);
    const table = tables[tableId - 1];
    if (!table) throw refusalError(`no table ${tableId} (the document has ${tables.length})`, "TABLE_NOT_FOUND");

    let empty = null;
    if (fill) {
      const ooxml = context.document.body.getOoxml();
      await context.sync();
      const doc = new DOMParser().parseFromString(ooxml.value, "application/xml");
      const ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
      const children = [...doc.getElementsByTagNameNS(ns, "body")[0].children];
      const xmlTables = children.filter(n => n.localName === "tbl");
      // Word's table numbering (body.tables) and the body's direct children must agree, or the paragraph
      // checked here could belong to a different table (a table inside a content control).
      if (xmlTables.length !== tables.length)
        throw refusalError("Table numbering differs between Word and the body XML (a table inside a content control?); refusing to fill a paragraph it cannot identify", "STRUCTURAL_BOUNDARY");
      const xmlTable = xmlTables[tableId - 1];
      const next = xmlTable ? children[children.indexOf(xmlTable) + 1] : null;
      if (!next || next.localName !== "p")
        throw refusalError("There is no paragraph immediately after the table to fill", "STRUCTURAL_BOUNDARY");
      const forbidden = new Set(payload.forbidden_empty_tags || []);
      const hit = [...next.getElementsByTagName("*")].find(n =>
        (n.namespaceURI === ns && forbidden.has(n.localName)) || (n.namespaceURI === ns && n.localName === "t" && n.textContent));
      if (hit) throw refusalError(`The paragraph after the table is not empty (${hit.localName})`, "STRUCTURAL_BOUNDARY");
      if (shapeNodes(next).length) throw refusalError("The paragraph after the table anchors a shape", "ANCHORED_SHAPES");
      empty = table.getRange("After").paragraphs.getFirst();
      empty.load("text,tableNestingLevel"); await context.sync();
      if (empty.text || empty.tableNestingLevel)
        throw refusalError("The paragraph after the table is not empty", "STRUCTURAL_BOUNDARY");
      await guardComments(context, [empty.getRange()], false);
      await guardRevisions(context, [empty.getRange()]);
    }
    // The destination is what gets written: the empty paragraph itself, else the gap after the table.
    const destination = fill ? empty.getRange() : table.getRange("After");
    await (await sectionGuard(context, payload)).check(() => [destination]);

    const body = context.document.body.paragraphs;
    body.load("items/text"); await context.sync();
    const before = body.items.map(p => p.text);
    const inserted = [];
    await withTracking(context, payload.track_changes, async () => {
      let previous = null;
      specs.forEach((spec, i) => {
        let para;
        if (i === 0 && fill) { para = empty; para.insertText(spec.text, "Replace"); }
        else para = (previous || table).insertParagraph(spec.text, "After");
        if (!(i === 0 && fill && !spec.style)) {  // a fill keeps the paragraph's own style unless one is named
          if (styles[i].kind === "name") para.style = styles[i].value;
          else para.styleBuiltIn = styles[i].value;
        }
        if (spec.color) para.font.color = withHash(spec.color);
        inserted.push(para); previous = para;
      });
      await context.sync();
    });
    inserted.forEach(p => { p.load("text,style,styleBuiltIn"); p.font.load("color"); });
    const fresh = context.document.body.paragraphs;
    fresh.load("items/text"); await context.sync();
    const observed = inserted.map((p, i) => ({text: String(p.text || "").replace(/[\r\n]+$/, ""), style: p.style,
      style_builtin: p.styleBuiltIn, color: p.font.color, requested_style: styles[i]}));
    const after = fresh.items.map(p => p.text);
    const problems = [];
    observed.forEach((o, i) => { if (o.text !== specs[i].text) problems.push(`paragraph ${i + 1} text`); });
    // Paragraph count and neighbours: the filled paragraph replaces one, an insert adds none.
    if (after.length !== before.length + specs.length - (fill ? 1 : 0)) problems.push("paragraph count");
    if (problems.length)
      throw refusalError(`Positional paragraph insertion did not read back as written (${problems.join(", ")}); observed: ${JSON.stringify(observed)}`, "VERIFICATION_FAILED");
    return {applied: true, inserted: observed, position: "after", before_count: before.length, after_count: after.length,
      pre, post: await mutationHash(context, {...payload, expectedBodySha256: undefined})};
  });
}

// -- merge two adjacent paragraphs of one cell (issue #107) ------------------

// The server built `ooxml` (one merged paragraph, hyperlink relationships intact) and sends
// `expected_text` (the pair as it read them). The pair must still read that way, or nothing is written.
// Tracked merges are not offered: how Word records them through insertOoxml is unproven.
async function opCellParagraphMerge(payload) {
  if (payload.track_changes) throw refusalError('Tracked paragraph merges are not supported', 'INVALID_INPUT');
  return measuredWordRun(async context=>{
    const pre=await mutationHash(context,payload);
    const cell=await resolveTableCell(context,payload);
    const paragraphs=cell.body.paragraphs;paragraphs.load('items/text');await context.sync();
    const i=payload.paragraph_index-1;
    if(!Number.isInteger(i) || i<0 || i+1>=paragraphs.items.length)
      throw refusalError('No next paragraph in this cell','STRUCTURAL_BOUNDARY');
    const before=paragraphs.items.map(p=>p.text);
    const pair=paragraphs.items.slice(i,i+2);
    // The pair must still read as the server read it. Text, not XML: real Word's per-paragraph getOoxml does
    // not canonicalize like the body export (every merge was refused LIVE_STALE), and the body-hash check
    // below (requireFreshBody) already guards everything else that could have changed.
    const flat=t=>String(t||'').replace(/[\r\n\u000b\u0007]/g,'');
    if(!Array.isArray(payload.expected_text) || payload.expected_text.length!==2)
      throw refusalError('Merge needs the two paragraph texts','INVALID_INPUT');
    pair.forEach((p,j)=>{
      if(flat(p.text)!==flat(payload.expected_text[j]))
        throw refusalError('Paragraph changed during preparation','LIVE_STALE');
    });
    const range=pair[0].getRange('Content').getRange('Start').expandTo(pair[1].getRange('Content').getRange('End'));
    await (await sectionGuard(context,payload)).check(()=>[range]);
    await guardComments(context,[range],false);await guardRevisions(context,[range]);
    await guardAnchoredParagraphs(context,[range]);await requireFreshBody(context,pre);
    // Word's own tracking mode applies to insertOoxml even when track_changes is false, which would
    // record the merge as a tracked change: a mode this op declares unsupported. Refuse, don't switch.
    context.document.load('changeTrackingMode');await context.sync();
    if (context.document.changeTrackingMode !== Word.ChangeTrackingMode.off)
      throw refusalError('Track Changes is on in Word; turn it off to merge paragraphs (tracked merges are not supported)','TRACKED_CHANGES_PRESENT');
    range.insertOoxml(payload.ooxml,'Replace');await context.sync();
    const after=await readCellParagraphs(context,cell);
    if(after.length !== before.length-1)
      throw refusalError('Paragraph merge count differs','VERIFICATION_FAILED');
    return {applied:true,pre,post:await mutationHash(context,{...payload,expectedBodySha256:undefined})};
  });
}
