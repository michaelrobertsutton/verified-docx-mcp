// Loads addin/taskpane.js into a vm with stub Office/Word/DOM globals and a
// fake Word context that counts sync() calls, then checks opCommentsList.
// Prints one JSON line of results; exit code 1 on a thrown error.
import fs from "node:fs";
import vm from "node:vm";

const src = fs.readFileSync(process.argv[2], "utf8");

function makeDoc(n) {
  const state = { syncs: 0, getRange: 0 };
  const comments = Array.from({ length: n }, (_, i) => {
    const c = {
      id: `c${i}`,
      content: `body ${i}`,
      authorName: "A",
      creationDate: "2026-09-29T00:00:00Z",
      resolved: i % 2 === 0,
      _anchor: `anchor ${i}`,
      load() {},
      getRange() {
        state.getRange += 1;
        return { text: c._anchor, load() {} };
      },
      replies: {
        items: [{ id: `r${i}`, content: `reply ${i}`, authorName: "B", creationDate: "d" }],
        load() {},
      },
    };
    return c;
  });
  const context = {
    document: { body: { getComments: () => ({ items: comments, load() {} }) } },
    sync: async () => {
      state.syncs += 1;
    },
  };
  return { state, context };
}

const sandbox = {
  console,
  Date,
  JSON,
  Map,
  Set,
  Promise,
  document: { getElementById: () => ({ addEventListener() {}, textContent: "", style: {} }) },
  window: {},
  crypto: {},
  TextEncoder,
  Office: { onReady() {}, HostType: { Word: "Word" }, context: {} },
  Word: { run: null },
};
vm.createContext(sandbox);
vm.runInContext(src + "\n;globalThis.__opCommentsList = opCommentsList;", sandbox);

async function run(n, payload) {
  const { state, context } = makeDoc(n);
  sandbox.Word.run = (fn) => fn(context);
  const result = await sandbox.__opCommentsList(payload);
  return { result, syncs: state.syncs, getRange: state.getRange };
}

const out = {};
const small = await run(3);
const big = await run(150);
out.syncsSmall = small.syncs;
out.syncsBig = big.syncs;
out.bigCount = big.result.comments.length;
out.firstKeys = Object.keys(big.result.comments[0]);
out.firstReplyCount = big.result.comments[0].replies.length;
out.hasTiming = typeof big.result.timing_ms.total === "number";
const filtered = await run(150, { ids: ["c7", "c9", "nope"] });
out.filteredIds = filtered.result.comments.map((c) => c.id);
out.filteredGetRange = filtered.getRange;
const noAnchor = await run(150, { ids: ["c7"], include_anchor: false });
out.noAnchorGetRange = noAnchor.getRange;
out.noAnchorHasAnchorText = "anchorText" in noAnchor.result.comments[0];
// issue #39: counts describe the whole collection, even under an ids filter.
out.bigCounts = big.result.counts;
out.filteredCounts = filtered.result.counts;
out.scope = big.result.scope;
out.hasObservedAt = typeof big.result.observed_at === "string" && big.result.observed_at.length > 0;
console.log(JSON.stringify(out));
