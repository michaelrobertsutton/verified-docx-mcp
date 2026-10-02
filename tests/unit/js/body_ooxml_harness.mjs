// Loads addin/taskpane.js into a vm with stub Office/Word/DOM globals and
// checks opBodyOoxml (issue #33): the reply shape, that binary payloads are
// looked up by the Flat OPC NAMESPACE (not a "pkg:" prefix) and emptied, the
// parse-error refusal, and the too_large refusal. Node has no DOMParser, so a
// minimal stub stands in; real-DOM behavior is covered by the live Word spike.
// Prints one JSON line of results.
import fs from "node:fs";
import vm from "node:vm";

const src = fs.readFileSync(process.argv[2], "utf8");
const FLAT_OPC_NS = "http://schemas.microsoft.com/office/2006/xmlPackage";

const seen = { ns: null, local: null, emptied: [], parsed: [] };

class StubDOMParser {
  parseFromString(xml, mime) {
    seen.parsed.push(mime);
    const bad = xml === "BAD";
    // The "document" is just { xml, parts }: three parts, two with binary payloads.
    const parts = [
      { name: "/word/media/a.png", binary: "AAAA" },
      { name: "/word/document.xml", binary: null },
      { name: "/word/media/b.png", binary: "BBBB" },
    ];
    const nodes = parts
      .filter((p) => p.binary !== null)
      .map((p) => ({
        parentNode: { getAttributeNS: (ns, local) => (ns === FLAT_OPC_NS && local === "name" ? p.name : null) },
        set textContent(v) {
          seen.emptied.push([p.name, v]);
        },
      }));
    return {
      xml,
      getElementsByTagName: (tag) => (bad && tag === "parsererror" ? [{}] : []),
      getElementsByTagNameNS: (ns, local) => {
        seen.ns = ns;
        seen.local = local;
        return nodes;
      },
    };
  }
}
class StubXMLSerializer {
  serializeToString(doc) {
    return `<serialized of="${doc.xml.length}"/>`;
  }
}

function makeContext(ooxml, text) {
  return {
    document: {
      body: {
        getOoxml: () => ({ value: ooxml }),
        load() {},
        text,
      },
    },
    sync: async () => {},
  };
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
  crypto: globalThis.crypto,
  TextEncoder,
  DOMParser: StubDOMParser,
  XMLSerializer: StubXMLSerializer,
  Office: { onReady() {}, HostType: { Word: "Word" }, context: { document: { url: "https://x.example/Doc.docx" } } },
  Word: { run: null },
};
vm.createContext(sandbox);
vm.runInContext(
  src + "\n;globalThis.__opBodyOoxml = opBodyOoxml; globalThis.__caps = PANE_CAPABILITIES; globalThis.__max = MAX_BODY_OOXML_BYTES;",
  sandbox
);

async function run(ooxml, text) {
  sandbox.Word.run = (fn) => fn(makeContext(ooxml, text));
  try {
    return { ok: true, result: await sandbox.__opBodyOoxml() };
  } catch (err) {
    return { ok: false, code: err.code, message: err.message };
  }
}

const out = {};
out.capabilities = Array.from(sandbox.__caps);

const good = await run("<pkg:package/>", "hello");
out.ok = good.ok;
out.keys = Object.keys(good.result || {});
out.ooxml = good.result.ooxml;
out.strippedParts = good.result.strippedParts;
out.documentUrl = good.result.documentUrl;
out.bodySha256 = good.result.bodySha256;
out.seenNs = seen.ns;
out.seenLocal = seen.local;
out.emptied = seen.emptied.slice(); // copy: later runs keep appending to `seen`
out.mime = seen.parsed[0];

const bad = await run("BAD", "hello");
out.badOk = bad.ok;
out.badCode = bad.code;

// A body whose reply would exceed the limit: shrink the limit instead of
// building 48 MiB. MAX_BODY_OOXML_BYTES is a top-level const, so re-evaluate
// the script with a small one.
const small = src.replace("const MAX_BODY_OOXML_BYTES = 48 * 1024 * 1024;", "const MAX_BODY_OOXML_BYTES = 64;");
const sandbox2 = { ...sandbox };
vm.createContext(sandbox2);
vm.runInContext(small + "\n;globalThis.__opBodyOoxml = opBodyOoxml;", sandbox2);
sandbox2.Word = { run: (fn) => fn(makeContext("<pkg:package/>", "hello")) };
try {
  await sandbox2.__opBodyOoxml();
  out.tooLargeCode = null;
} catch (err) {
  out.tooLargeCode = err.code;
}

console.log(JSON.stringify(out));
