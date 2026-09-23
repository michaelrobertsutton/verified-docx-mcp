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

async function collectComments(context) {
  const body = context.document.body;
  const comments = body.getComments();
  comments.load("items");
  await context.sync();

  const results = [];
  for (const comment of comments.items) {
    comment.load(["id", "content", "authorName", "creationDate", "resolved"]);
    const range = comment.getRange();
    range.load("text");
    results.push({ comment, range });
  }
  await context.sync();

  return results.map(({ comment, range }) => ({
    id: comment.id,
    content: comment.content,
    authorName: comment.authorName,
    creationDate: comment.creationDate,
    resolved: comment.resolved,
    anchorText: range.text,
  }));
}

async function buildReport() {
  return Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();

    const bodyText = body.text || "";
    const bodyHash = await sha256Hex(bodyText);
    const comments = await collectComments(context);

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
  try {
    const report = await buildReport();
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
    lastReport = null;
    document.getElementById("copy-json").disabled = true;
    const message = err && err.message ? err.message : String(err);
    setOutput({ error: message, debugInfo: err && err.debugInfo ? err.debugInfo : undefined });
    setStatus(`ERROR: ${message}`);
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
const PANE_CAPABILITIES = ["row_scope", "cell_edit"];

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
  commentsEl.textContent = String(open);
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

function connectOpsSocket() {
  if (reconnectTimer) {
    clearTimeout(reconnectTimer);
    reconnectTimer = null;
  }
  setWsStatus("connecting…");
  let socket;
  try {
    socket = new WebSocket(`wss://localhost:${OPS_PORT}/ops`);
  } catch (err) {
    setWsStatus(`failed to open (${err && err.message ? err.message : err}); retrying`);
    scheduleReconnect();
    return;
  }
  opsSocket = socket;

  socket.onopen = async () => {
    reconnectDelayMs = RECONNECT_MIN_MS;
    try {
      const hash = await currentBodyHash();
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
        })
      );
      setWsStatus("connected");
      startHeartbeat();
    } catch (err) {
      setWsStatus(`hello failed (${err && err.message ? err.message : err})`);
    }
  };

  socket.onmessage = async (event) => {
    let message;
    try {
      message = JSON.parse(event.data);
    } catch (err) {
      return; // malformed message from the bridge; ignore rather than crash the pane
    }
    if (message.type !== "request") return;
    const requestId = message.request_id;
    const op = message.op;
    const payload = message.payload || {};
    try {
      const result = await dispatchOp(op, payload);
      logOp(op, true, summarizeResult(op, result));
      socket.send(JSON.stringify({ type: "reply", request_id: requestId, ok: true, result }));
    } catch (err) {
      const code = err && err.code ? err.code : "LIVE_OP_FAILED";
      const msg = err && err.message ? err.message : String(err);
      logOp(op, false, msg);
      socket.send(
        JSON.stringify({ type: "reply", request_id: requestId, ok: false, error: { code, message: msg } })
      );
    }
  };

  socket.onclose = () => {
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

function startHeartbeat() {
  stopHeartbeat();
  heartbeatTimer = setInterval(async () => {
    if (!opsSocket || opsSocket.readyState !== WebSocket.OPEN) return;
    try {
      const hash = await currentBodyHash();
      opsSocket.send(
        JSON.stringify({ type: "heartbeat", documentUrl: Office.context.document.url, bodySha256: hash })
      );
    } catch (err) {
      // Best-effort: a single failed heartbeat just means the bridge's
      // SessionRegistry evicts this session after missed_heartbeats and
      // the reconnect loop (onclose -> scheduleReconnect) takes over.
    }
  }, HEARTBEAT_INTERVAL_MS);
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
  return "";
}

async function dispatchOp(op, payload) {
  switch (op) {
    case "ping":
      return opPing();
    case "describe":
      return opDescribe();
    case "search":
      return opSearch(payload);
    case "replace":
      return opReplace(payload);
    case "format":
      return opFormat(payload);
    case "comments_list":
      return opCommentsList();
    case "comment_add":
      return opCommentAdd(payload);
    case "comment_reply":
      return opCommentReply(payload);
    case "comment_resolve":
      return opCommentResolve(payload);
    case "cell_get":
      return opCellGet(payload);
    case "cell_set":
      return opCellSet(payload);
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
  return Word.run(async (context) => {
    context.document.body.load("text");
    await context.sync();
    return {};
  });
}

async function opDescribe() {
  return Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    context.document.load(["changeTrackingMode", "saved"]);
    await context.sync();
    const hash = await sha256Hex(body.text || "");
    return {
      documentUrl: Office.context.document.url,
      bodySha256: hash,
      changeTrackingMode: String(context.document.changeTrackingMode),
      saved: context.document.saved,
    };
  });
}

// -- search -------------------------------------------------------------

async function opSearch(payload) {
  return Word.run(async (context) => {
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
  if (payload.rowAnchor === null || payload.rowAnchor === undefined) {
    const results = context.document.body.search(payload.find, { matchCase: true, matchWholeWord: false });
    results.load("text");
    await context.sync();
    return results.items;
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

async function opReplace(payload) {
  const expectedMatches = payload.expected_matches;
  return Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();
    const preHash = await sha256Hex(body.text || "");

    const matchItems = await searchScoped(context, payload);

    if (matchItems.length !== expectedMatches) {
      throw refusalError(
        `expected ${expectedMatches} match(es) for ${JSON.stringify(payload.find)}, found ${matchItems.length}`
      );
    }

    let previousMode = null;
    if (payload.track_changes) {
      context.document.load("changeTrackingMode");
      await context.sync();
      previousMode = context.document.changeTrackingMode;
      context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    }

    const matches = matchItems.map((range) => ({ before: range.text, after: payload.replace }));
    try {
      matchItems.forEach((range) => range.insertText(payload.replace, Word.InsertLocation.replace));
      await context.sync();
    } finally {
      if (previousMode !== null) {
        context.document.changeTrackingMode = previousMode;
        await context.sync();
      }
    }

    const postBody = context.document.body;
    postBody.load("text");
    await context.sync();
    const postHash = await sha256Hex(postBody.text || "");

    return { applied: true, match_count: matchItems.length, matches, pre: preHash, post: postHash };
  });
}

async function opFormat(payload) {
  const expectedMatches = payload.expected_matches;
  return Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();
    const preHash = await sha256Hex(body.text || "");

    const matchItems = await searchScoped(context, payload);

    if (matchItems.length !== expectedMatches) {
      throw refusalError(
        `expected ${expectedMatches} match(es) for ${JSON.stringify(payload.find)}, found ${matchItems.length}`
      );
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
    matchItems.forEach((range) => range.font.load(["color", "strikeThrough"]));
    await context.sync();
    matchItems.forEach((range, i) => {
      matches[i].colorAfter = range.font.color;
      matches[i].strikeAfter = range.font.strikeThrough;
    });

    const postBody = context.document.body;
    postBody.load("text");
    await context.sync();
    const postHash = await sha256Hex(postBody.text || "");

    return { applied: true, match_count: matchItems.length, matches, pre: preHash, post: postHash };
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
async function resolveTableCell(context, payload) {
  const tables = context.document.body.tables;
  tables.load("items");
  await context.sync();
  tables.items.forEach((t) => t.load("nestingLevel"));
  await context.sync();
  if (tables.items.some((t) => t.nestingLevel > 1)) {
    throw refusalError(
      "the document contains a nested table; table numbering cannot be mapped reliably, so live cell edits are refused"
    );
  }
  const table = tables.items[payload.table_index - 1];
  if (!table) {
    throw refusalError(`no table ${payload.table_index} (the document has ${tables.items.length})`);
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

async function opCellGet(payload) {
  return Word.run(async (context) => {
    const cell = await resolveTableCell(context, payload);
    cell.body.load("text");
    await context.sync();
    return { text: cell.body.text || "" };
  });
}

async function opCellSet(payload) {
  return Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();
    const preHash = await sha256Hex(body.text || "");

    const cell = await resolveTableCell(context, payload);
    cell.body.load("text");
    await context.sync();
    const before = cell.body.text || "";
    // Compare-and-set: refuse if the cell changed since the server read it
    // (a co-author typing in this same cell), rather than overwrite them.
    if (before !== payload.expected_before_text) {
      throw refusalError(
        "the cell's text changed since it was read (another editor may be working in it); nothing was written"
      );
    }

    let previousMode = null;
    if (payload.track_changes) {
      context.document.load("changeTrackingMode");
      await context.sync();
      previousMode = context.document.changeTrackingMode;
      context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    }

    try {
      cell.body.clear();
      const paragraphs = payload.paragraphs || [];
      paragraphs.forEach((runs, i) => {
        // After clear() the cell holds one empty paragraph: fill it first,
        // append the rest.
        const paragraph =
          i === 0 ? cell.body.paragraphs.getFirst() : cell.body.insertParagraph("", Word.InsertLocation.end);
        (runs || []).forEach((run) => {
          if (run.hard_break) {
            paragraph.insertBreak(Word.BreakType.line, Word.InsertLocation.end);
            return;
          }
          const range = paragraph.insertText(run.text, Word.InsertLocation.end);
          range.font.bold = !!run.bold;
          range.font.italic = !!run.italic;
          if (run.link) {
            range.hyperlink = run.link;
          }
        });
      });
      await context.sync();
    } finally {
      if (previousMode !== null) {
        context.document.changeTrackingMode = previousMode;
        await context.sync();
      }
    }

    cell.body.load("text");
    const postBody = context.document.body;
    postBody.load("text");
    await context.sync();
    const postHash = await sha256Hex(postBody.text || "");

    return { applied: true, before, after: cell.body.text || "", pre: preHash, post: postHash };
  });
}

// -- comments -------------------------------------------------------------

async function opCommentsList() {
  return Word.run(async (context) => {
    const summaries = await collectComments(context); // WP-1 helper: id, content, authorName, creationDate, resolved, anchorText
    const comments = [];
    for (const summary of summaries) {
      // eslint-disable-next-line no-await-in-loop -- each comment's replies
      // collection must be loaded and synced before the next comment's.
      comments.push(await withReplies(context, summary));
    }
    return { comments };
  });
}

async function withReplies(context, commentSummary) {
  const comments = context.document.body.getComments();
  comments.load("items/id");
  await context.sync();
  const match = comments.items.find((c) => c.id === commentSummary.id);
  if (!match) return Object.assign({}, commentSummary, { replies: [] });
  const replies = match.replies; // CommentReplyCollection is a property, not a method
  replies.load("items");
  await context.sync();
  replies.items.forEach((r) => r.load(["id", "content", "authorName", "creationDate"]));
  await context.sync();
  return Object.assign({}, commentSummary, {
    replies: replies.items.map((r) => ({
      id: r.id,
      content: r.content,
      authorName: r.authorName,
      creationDate: r.creationDate,
    })),
  });
}

async function opCommentAdd(payload) {
  const expectedMatches = payload.expected_matches;
  return Word.run(async (context) => {
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
  return Word.run(async (context) => {
    const comments = context.document.body.getComments();
    comments.load("items/id");
    await context.sync();
    const comment = comments.items.find((c) => c.id === payload.comment_id);
    if (!comment) throw refusalError(`no comment with id ${JSON.stringify(payload.comment_id)}`);
    const reply = comment.reply(payload.text);
    reply.load("id");
    await context.sync();
    return { reply_id: reply.id };
  });
}

async function opCommentResolve(payload) {
  return Word.run(async (context) => {
    const comments = context.document.body.getComments();
    comments.load("items/id");
    await context.sync();
    const comment = comments.items.find((c) => c.id === payload.comment_id);
    if (!comment) throw refusalError(`no comment with id ${JSON.stringify(payload.comment_id)}`);
    comment.resolved = !!payload.resolved;
    await context.sync();
    return { resolved: !!payload.resolved };
  });
}

async function opSave() {
  return Word.run(async (context) => {
    context.document.save();
    await context.sync();
    return { saved: true };
  });
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    setStatus(`Loaded outside Word (host=${info.host}); this pane is Word-only.`);
    return;
  }
  document.getElementById("copy-json").addEventListener("click", copyJson);
  document.getElementById("ping-bridge").addEventListener("click", pingBridge);
  document.getElementById("ping-bridge").disabled = false;
  runAndRender();
  connectOpsSocket();
});
