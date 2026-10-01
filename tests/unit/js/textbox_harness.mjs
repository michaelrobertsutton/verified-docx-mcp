// Loads addin/taskpane.js into a vm with a stubbed Office/Word object model
// and exercises the text-box ops (issue #35): the shared shape traversal
// (groups, canvases, geometric shapes, duplicate ids, unreadable bodies),
// the WordApiDesktop 1.2 capability gate, the compare-and-set that must
// refuse BEFORE any insertText, and that the body path never touches shapes.
// The stubs prove the pane's control flow only; what a real Word returns for
// body.shapes needs the manual sideload check in docs/live-mode.md.
// Prints one JSON line of results.
import crypto from "node:crypto";
import fs from "node:fs";
import vm from "node:vm";

const src = fs.readFileSync(process.argv[2], "utf8");
const sha = (t) => crypto.createHash("sha256").update(t, "utf8").digest("hex");

let desktopSupported = true;
const calls = { insertText: [], bodySearch: 0 };

const collection = (items) => ({ items, load() {} });

function makeRange(shape, start, end) {
  return {
    text: shape._text.slice(start, end),
    load() {},
    font: { load() {} },
    insertText(t) {
      calls.insertText.push([shape.id, t]);
      shape._text = shape._text.slice(0, start) + t + shape._text.slice(end);
      return { text: t, load() {} };
    },
  };
}

function makeShape(id, type, opts = {}) {
  const shape = {
    id,
    type,
    _text: opts.text ?? "",
    shapeGroup: { shapes: collection(opts.children || []) },
    canvas: { shapes: collection(opts.children || []) },
  };
  shape.body = {
    load() {
      if (opts.unreadable) throw new Error("no text frame");
    },
    get text() {
      return shape._text;
    },
    get paragraphs() {
      return { items: shape._text.split("\r").map(() => ({})), load() {} };
    },
    search(find) {
      const ranges = [];
      let from = 0;
      for (;;) {
        const i = shape._text.indexOf(find, from);
        if (i < 0) break;
        ranges.push(makeRange(shape, i, i + find.length));
        from = i + find.length;
      }
      return { items: ranges, load() {} };
    },
  };
  return shape;
}

function makeContext(shapes, bodyText = "body text") {
  const body = {
    text: bodyText,
    load() {},
    shapes: collection(shapes),
    search() {
      calls.bodySearch += 1;
      return { items: [], load() {} };
    },
  };
  return { document: { body, load() {}, changeTrackingMode: "Off" }, sync: async () => {} };
}

const sandbox = {
  console,
  Date,
  JSON,
  Map,
  Set,
  Promise,
  Object,
  String,
  Error,
  document: { getElementById: () => ({ addEventListener() {}, textContent: "", style: {} }) },
  window: {},
  crypto: globalThis.crypto,
  TextEncoder,
  Office: {
    onReady() {},
    HostType: { Word: "Word" },
    context: {
      document: { url: "https://x.example/Doc.docx" },
      requirements: { isSetSupported: (set) => (set === "WordApiDesktop" ? desktopSupported : true) },
    },
  },
  Word: {
    run: null,
    ChangeTrackingMode: { trackAll: "TrackAll" },
    InsertLocation: { replace: "Replace" },
    UnderlineType: { single: "Single", none: "None" },
  },
};
vm.createContext(sandbox);
vm.runInContext(
  src +
    "\n;globalThis.__list = opTextboxList; globalThis.__replace = opReplace; globalThis.__caps = paneCapabilities;",
  sandbox
);

async function attempt(fn) {
  try {
    return { ok: true, result: await fn() };
  } catch (err) {
    return { ok: false, code: err.code, message: err.message };
  }
}

function withShapes(shapes) {
  sandbox.Word.run = (fn) => fn(makeContext(shapes));
}

const out = {};

// -- capability gate ---------------------------------------------------------
desktopSupported = true;
out.capsWithDesktop = Array.from(sandbox.__caps());
desktopSupported = false;
out.capsWithoutDesktop = Array.from(sandbox.__caps());
withShapes([makeShape("1", "TextBox", { text: "x" })]);
const unsupported = await attempt(() => sandbox.__list());
out.unsupportedCode = unsupported.code;
desktopSupported = true;

// -- traversal: groups, canvases, geometric shapes, pictures -------------------
const tree = [
  makeShape("1", "TextBox", { text: "Why Team Skyward\rsecond" }),
  makeShape("2", "Group", {
    children: [
      makeShape("3", "TextBox", { text: "inside group" }),
      makeShape("4", "Canvas", { children: [makeShape("5", "GeometricShape", { text: "deep callout" })] }),
    ],
  }),
  makeShape("6", "Picture"),
];
withShapes(tree);
const listed = await attempt(() => sandbox.__list());
out.listedIds = listed.result.textboxes.map((t) => t.shape_id);
out.listedTypes = listed.result.textboxes.map((t) => t.type);
out.deepGroupPath = listed.result.textboxes.find((t) => t.shape_id === "5").group_path;
out.firstParagraphCount = listed.result.textboxes[0].paragraph_count;
out.firstShaMatches = listed.result.textboxes[0].text_sha256 === sha("Why Team Skyward\rsecond");
out.skipped = listed.result.skipped;
out.incompleteEmpty = listed.result.incomplete.length === 0;

// -- duplicate ids and unreadable bodies are reported, never silently skipped --
withShapes([
  makeShape("7", "TextBox", { text: "a" }),
  makeShape("7", "TextBox", { text: "b" }),
  makeShape("8", "TextBox", { text: "ok" }),
  makeShape("9", "TextBox", { unreadable: true }),
]);
const messy = await attempt(() => sandbox.__list());
out.messyIds = messy.result.textboxes.map((t) => t.shape_id);
out.messyIncomplete = messy.result.incomplete.map((i) => [i.shape_id, i.reason.split(":")[0]]);

// -- exhaustive scope refuses on incomplete coverage; a single shape still works
const expectOk = { 8: sha("ok") };
const refuseIncomplete = await attempt(() =>
  sandbox.__replace({ find: "ok", expected_matches: 1, replace: "OK", scope: "textboxes", expect: expectOk })
);
out.incompleteCode = refuseIncomplete.code;
const single = await attempt(() =>
  sandbox.__replace({ find: "ok", expected_matches: 1, replace: "OK", scope: "shape:8", expect: expectOk })
);
out.singleOk = single.ok;
out.singlePostText = single.ok ? single.result.shapes[0].post_text : null;
out.singleAfter = single.ok ? single.result.matches[0].after : null;

// -- compare-and-set: refuse BEFORE any insertText ---------------------------
calls.insertText.length = 0;
withShapes([makeShape("1", "TextBox", { text: "alpha beta" })]);
const stale = await attempt(() =>
  sandbox.__replace({
    find: "beta",
    expected_matches: 1,
    replace: "B",
    scope: "textboxes",
    expect: { 1: sha("alpha beta EDITED BY A CO-AUTHOR") },
  })
);
out.staleCode = stale.code;
out.insertsAfterStale = calls.insertText.length;

const vanished = await attempt(() =>
  sandbox.__replace({
    find: "beta",
    expected_matches: 1,
    replace: "B",
    scope: "textboxes",
    expect: { 1: sha("alpha beta"), 99: sha("gone") },
  })
);
out.vanishedCode = vanished.code;

const noExpect = await attempt(() =>
  sandbox.__replace({ find: "beta", expected_matches: 1, replace: "B", scope: "textboxes" })
);
out.noExpectOk = noExpect.ok;

const wrongCount = await attempt(() =>
  sandbox.__replace({
    find: "beta",
    expected_matches: 2,
    replace: "B",
    scope: "textboxes",
    expect: { 1: sha("alpha beta") },
  })
);
out.wrongCountMessage = wrongCount.message;
out.insertsAfterWrongCount = calls.insertText.length;

// -- the happy path reads back what Word reports ---------------------------------
const good = await attempt(() =>
  sandbox.__replace({
    find: "beta",
    expected_matches: 1,
    replace: "B",
    scope: "textboxes",
    expect: { 1: sha("alpha beta") },
  })
);
out.goodOk = good.ok;
out.goodShapes = good.result.shapes;
out.goodMatch = good.result.matches[0];
out.goodBodyMatchCount = good.result.body_match_count;
out.goodPreEqualsPost = good.result.pre === good.result.post;

// -- invalid scope, and body scope never touches shapes -------------------------
const invalid = await attempt(() => sandbox.__replace({ find: "a", expected_matches: 1, replace: "b", scope: "nope" }));
out.invalidScopeOk = invalid.ok;
calls.bodySearch = 0;
calls.insertText.length = 0;
withShapes([makeShape("1", "TextBox", { text: "beta" })]);
const body = await attempt(() => sandbox.__replace({ find: "beta", expected_matches: 1, replace: "B" }));
out.bodyPathRefusal = body.ok ? null : body.message;
out.bodySearchCalls = calls.bodySearch;
out.insertsOnBodyPath = calls.insertText.length;

console.log(JSON.stringify(out));
