// WP-1 spike pane logic (https://github.com/michaelrobertsutton/JennyStack/issues/106).
// Plain script, no framework, no bundler. Runs once Office.onReady fires,
// then lets the lead re-run via the buttons. Everything it reads is
// read-only -- this pane makes no document edits (see docs/live-mode.md).

/* global Office, Word, console */

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

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    setStatus(`Loaded outside Word (host=${info.host}); this pane is Word-only.`);
    return;
  }
  document.getElementById("copy-json").addEventListener("click", copyJson);
  document.getElementById("ping-bridge").addEventListener("click", pingBridge);
  document.getElementById("ping-bridge").disabled = false;
  runAndRender();
});
