import { webcrypto } from "node:crypto";
// Loads addin/taskpane.js into a vm with a stub Word context and checks
// opCommentDelete (issue #80). Properties that Office.js only exposes after
// load() throw when read unloaded, like the real thing. Prints one JSON line.
import fs from "node:fs";
import vm from "node:vm";

const src = fs.readFileSync(process.argv[2], "utf8");

function makeDoc({ resolved = true, deleteSticks = true } = {}) {
  const state = { deleted: [], syncs: 0 };
  const items = ["c1", "c2"].map((id, i) => {
    const loaded = new Set();
    const real = { id, content: `body ${i}`, authorName: "A", resolved: id === "c1" ? resolved : false };
    const c = {
      load(names) { String(names).split(",").forEach((n) => loaded.add(n.trim())); },
      getRange: () => ({ text: `anchor ${i}`, load() {}, paragraphs: { items: [{ text: `p ${i}` }], load() {} } }),
      replies: { items: [{ id: `r${i}` }], load() {} },
      delete() { state.deleted.push(id); },
    };
    for (const k of Object.keys(real)) {
      Object.defineProperty(c, k, {
        get() {
          if (k !== "id" && !loaded.has(k)) throw new Error(`property ${k} not loaded`);
          return real[k];
        },
      });
    }
    return c;
  });
  const context = {
    document: {
      body: {
        getComments: () => ({
          get items() { return deleteSticks ? items.filter((c) => !state.deleted.includes(c.id)) : items; },
          load() {},
        }),
      },
    },
    sync: async () => { state.syncs += 1; },
  };
  return { state, context };
}

const sandbox = {
  console, Date, JSON, Map, Set, Promise,
  document: { getElementById: () => ({ addEventListener() {}, textContent: "", style: {} }) },
  window: {}, crypto: webcrypto, TextEncoder,
  Office: { onReady() {}, HostType: { Word: "Word" }, context: {} },
  Word: { run: null },
};
vm.createContext(sandbox);
vm.runInContext(src + "\n;globalThis.__del = opCommentDelete;", sandbox);

async function run(opts, payload) {
  const { state, context } = makeDoc(opts);
  sandbox.Word.run = (fn) => fn(context);
  try {
    const result = await sandbox.__del(payload);
    return { result, deleted: state.deleted };
  } catch (e) {
    return { error: String(e.message), code: e.code, deleted: state.deleted };
  }
}

const out = {};
out.ok = await run({}, { comment_id: "c1", require_resolved: true });
out.openRefused = await run({ resolved: false }, { comment_id: "c1", require_resolved: true });
out.openAllowedWhenNotRequired = await run({ resolved: false }, { comment_id: "c1", require_resolved: false });
out.missing = await run({}, { comment_id: "nope" });
out.staleEpoch = await run({}, { comment_id: "c1", session_epoch: "old" });
out.identityMismatch = await run({}, { comment_id: "c1", identity: { content: "different" } });
out.notDeleted = await run({ deleteSticks: false }, { comment_id: "c1" });
console.log(JSON.stringify(out));
