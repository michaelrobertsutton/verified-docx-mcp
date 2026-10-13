import { webcrypto } from "node:crypto";
// Loads addin/taskpane.js into a vm with stub Office/Word/DOM globals and a
// fake Word context that counts sync() calls, then checks opCommentsList.
// Prints one JSON line of results; exit code 1 on a thrown error.
import fs from "node:fs";
import vm from "node:vm";
import { DOMParser } from "./shape_dom.mjs";

const src = fs.readFileSync(process.argv[2], "utf8");

function makeDoc(n, ooxml) {
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
        return { text: c._anchor, load() {}, paragraphs: {items: [{text: `paragraph ${i}`}], load() {}} };
      },
      replies: {
        items: [{ id: `r${i}`, content: `reply ${i}`, authorName: "B", creationDate: "d" }],
        load() {},
      },
    };
    return c;
  });
  const context = {
    document: {
      body: {
        getComments: () => ({ items: comments, load() {} }),
        ...(ooxml === undefined ? {} : { getOoxml: () => ({ value: ooxml }) }),
      },
    },
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
  crypto: webcrypto,
  TextEncoder,
  Office: { onReady() {}, HostType: { Word: "Word" }, context: {} },
  Word: { run: null },
  DOMParser,
};
vm.createContext(sandbox);
vm.runInContext(src + "\n;globalThis.__opCommentsList = opCommentsList; globalThis.__checkEpoch = checkCommentEpoch;", sandbox);

async function run(n, payload, ooxml) {
  const { state, context } = makeDoc(n, ooxml);
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

// issue #39: OOXML cross-check. Three roots (one resolved by commentsExtended) plus one reply.
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const W15 = "http://schemas.microsoft.com/office/word/2012/wordml";
const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage";
const comment = (id, para) => `<w:comment w:id="${id}"><w:p w14:paraId="X${id}"/><w:p w14:paraId="${para}"><w:r><w:t>c${id}</w:t></w:r></w:p></w:comment>`;
const pkg = (parts) => `<pkg:package xmlns:pkg="${PKG}" xmlns:w="${W}" xmlns:w14="${W14}" xmlns:w15="${W15}">${parts}</pkg:package>`;
const part = (name, body) => `<pkg:part pkg:name="${name}"><pkg:xmlData>${body}</pkg:xmlData></pkg:part>`;
const commentsPart = part("/word/comments.xml", `<w:comments>${comment(1, "A1")}${comment(2, "A2")}${comment(3, "A3")}${comment(4, "A4")}</w:comments>`);
const extendedPart = part("/word/commentsExtended.xml",
  `<w15:commentsEx><w15:commentEx w15:paraId="A1" w15:done="0"/><w15:commentEx w15:paraId="A2" w15:done="1"/>` +
  `<w15:commentEx w15:paraId="A3" w15:done="false"/><w15:commentEx w15:paraId="A4" w15:paraIdParent="A1" w15:done="0"/></w15:commentsEx>`);
const full = await run(1, undefined, pkg(commentsPart + extendedPart));
out.ooxml = full.result.ooxml;
out.ooxmlFiltered = (await run(1, { ids: ["c0"] }, pkg(commentsPart + extendedPart))).result.ooxml ?? null;
out.ooxmlNoExtended = (await run(1, undefined, pkg(commentsPart))).result.ooxml;
out.ooxmlNoComments = (await run(1, undefined, pkg(extendedPart))).result.ooxml;
out.ooxmlGarbage = (await run(1, undefined, "<pkg:package")).result.ooxml;
out.ooxmlNoReader = big.result.ooxml;
out.ooxmlSyncs = full.syncs;
out.capability = vm.runInContext("PANE_CAPABILITIES.includes('comments_ooxml_crosscheck')", sandbox);
console.log(JSON.stringify(out));

try { sandbox.__checkEpoch({session_epoch: "stale"}); throw new Error("stale epoch accepted"); } catch (e) { if (e.code !== "COMMENT_ID_STALE") throw e; }
