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

async function opReplace(payload) {
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
      throw refusalError(
        `expected ${expectedMatches} match(es) for ${JSON.stringify(payload.find)}, found ${results.items.length}`
      );
    }

    let previousMode = null;
    if (payload.track_changes) {
      context.document.load("changeTrackingMode");
      await context.sync();
      previousMode = context.document.changeTrackingMode;
      context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    }

    const matches = results.items.map((range) => ({ before: range.text, after: payload.replace }));
    results.items.forEach((range) => range.insertText(payload.replace, Word.InsertLocation.replace));
    await context.sync();

    if (previousMode !== null) {
      context.document.changeTrackingMode = previousMode;
      await context.sync();
    }

    const postBody = context.document.body;
    postBody.load("text");
    await context.sync();
    const postHash = await sha256Hex(postBody.text || "");

    return { applied: true, match_count: results.items.length, matches, pre: preHash, post: postHash };
  });
}

async function opFormat(payload) {
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
      throw refusalError(
        `expected ${expectedMatches} match(es) for ${JSON.stringify(payload.find)}, found ${results.items.length}`
      );
    }

    let previousMode = null;
    if (payload.track_changes) {
      context.document.load("changeTrackingMode");
      await context.sync();
      previousMode = context.document.changeTrackingMode;
      context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    }

    const matches = results.items.map((range) => ({ before: range.text, after: range.text }));
    results.items.forEach((range) => {
      if (payload.bold !== null && payload.bold !== undefined) range.font.bold = payload.bold;
      if (payload.italic !== null && payload.italic !== undefined) range.font.italic = payload.italic;
      if (payload.underline !== null && payload.underline !== undefined) {
        range.font.underline = payload.underline ? Word.UnderlineType.single : Word.UnderlineType.none;
      }
    });
    await context.sync();

    if (previousMode !== null) {
      context.document.changeTrackingMode = previousMode;
      await context.sync();
    }

    const postBody = context.document.body;
    postBody.load("text");
    await context.sync();
    const postHash = await sha256Hex(postBody.text || "");

    return { applied: true, match_count: results.items.length, matches, pre: preHash, post: postHash };
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
