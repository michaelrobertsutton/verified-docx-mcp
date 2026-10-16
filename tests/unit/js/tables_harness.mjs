import {DOMParser,EMPTY} from './shape_dom.mjs';
// Loads addin/taskpane.js into a vm with a MOCK Word object model and drives the
// issue #34 ops (table_get / table_insert / cells_set) plus the refactored
// cell_set. Prints one JSON line: {results: [{name, pass, detail}]}.
//
// The mock enforces the Office.js rules the pane code has to respect, so a
// pass means more than "the setters were called":
//   * a proxy property can only be read after load() AND a following sync()
//   * ClientResult.value (compareLocationWith, getOoxml) is only readable
//     after the sync that resolves it
//   * writes are QUEUED and only applied at sync(), in order; a call log is
//     recorded at queue time, so "nothing was written" means an empty log
//   * Word's enum strings ("Centered", "Justified", ...) are what the setters
//     receive, not the server's vocabulary
// It is still a model, not Word: real-Word behaviour stays UNVERIFIED
// (docs/live-mode.md runbook).
import fs from "node:fs";
import vm from "node:vm";

const src = fs.readFileSync(process.argv[2], "utf8");
const results = [];
const check = (name, pass, detail) => results.push({ name, pass: !!pass, detail: pass ? "" : String(detail ?? "") });

// ---------------------------------------------------------------- mock model

function makeModel(opts = {}) {
  const model = {
    paragraphs: (opts.paragraphs || ["Intro", "Proof", "Tail"]).map((p) => (typeof p === "string" ? { text: p } : { ...p })),
    tables: [],
    order: [],
    mode: opts.mode || "Off",
    styles: new Set(opts.styles || ["Table Grid"]),
    supports15: opts.supports15 !== false,
    calls: [], // every queued call/assignment, in queue order
    failOnInsert: !!opts.failOnInsert,
    onSyncCount: null,
    // issue #79 knobs
    ooxmlMode: opts.ooxmlMode || "empty", // "revisions": body.getOoxml reflects model revisions
    ooxmlUnreadable: false,
    dropDetach: false, // detachFromList is accepted but does not take
    dropTableDelete: false, // table.delete() is accepted but does not take
    scopedRejectTouchesOutside: false,
    onTableDelete: null,
  };
  model.paragraphs.forEach((p) => model.order.push({ kind: "p", ref: p }));
  (opts.tables || []).forEach((spec) => addTable(model, spec, null));
  return model;
}

function newCell(text = "") {
  return { paras: [text], fill: "", hAlign: "Left", vAlign: "Top", width: 72, bold: false, color: "", size: 11 };
}

function addTable(model, spec, atOrderIndex) {
  const t = {
    style: spec.style ?? "Table Grid",
    headerRowCount: spec.headerRowCount ?? 0,
    nesting: spec.nesting ?? 1,
    rows: (spec.rows || [["x"]]).map((r) => r.map((text) => newCell(text))),
    revs: spec.revs ? spec.revs.map((r) => ({ ...r })) : [],
    ooxml: spec.ooxml ?? '<pkg:package><pkg:part pkg:name="/word/document.xml"><w:tbl/></pkg:part></pkg:package>',
  };
  if (spec.hAlign) t.rows.forEach((r) => r.forEach((c) => (c.hAlign = spec.hAlign)));
  if (spec.listed) t.rows.forEach((r) => r.forEach((c) => (c.listed = true)));
  model.tables.push(t);
  const item = { kind: "t", ref: t };
  if (atOrderIndex === null) model.order.push(item);
  else model.order.splice(atOrderIndex, 0, item);
  return t;
}

function cellText(c) {
  return c.paras.join("\n");
}

function bodyText(model) {
  return model.order
    .map((it) => (it.kind === "p" ? it.ref.text : it.ref.rows.map((r) => r.map(cellText).join("\t")).join("\n")))
    .join("\n");
}

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const revXml = (r) => `<w:${r.kind} w:author="${r.author}"><w:r><w:t>${r.text}</w:t></w:r></w:${r.kind}>`;
function revisionOoxml(model) {
  const items = model.order.map((it) =>
    it.kind === "p"
      ? `<w:p>${(it.ref.revs || []).map(revXml).join("")}<w:r><w:t>${it.ref.text}</w:t></w:r></w:p>`
      : `<w:tbl>${it.ref.rows
          .map((r, ri) => `<w:tr>${ri === 0 ? `<w:trPr>${it.ref.revs.map(revXml).join("")}</w:trPr>` : ""}${r
            .map((c) => `<w:tc><w:p><w:r><w:t>${cellText(c)}</w:t></w:r></w:p></w:tc>`)
            .join("")}</w:tr>`)
          .join("")}</w:tbl>`
  );
  return `<w:document ${W_NS}><w:body>${items.join("")}</w:body></w:document>`;
}

const topLevelTables = (model) => model.order.filter((i) => i.kind === "t").map((i) => i.ref);

// ------------------------------------------------------------- mock Office.js

function makeContext(model) {
  const ctx = { queue: [], syncs: 0 };
  const log = (s) => model.calls.push(s);

  function node(label, readable, impl, setters = {}) {
    const loaded = new Set();
    impl.load = (spec) => {
      (model.loads ||= []).push([label, String(spec)]);
      const names = (Array.isArray(spec) ? spec : String(spec).split(",")).map((s) => s.trim());
      ctx.queue.push(() => names.forEach((n) => loaded.add(n.split("/").pop())));
      return proxy;
    };
    impl.__seed = (names) => names.forEach((n) => loaded.add(n));
    const proxy = new Proxy(impl, {
      get(t, k) {
        if (typeof k === "string" && readable.includes(k) && !loaded.has(k)) {
          throw new Error(`${label}.${k} read before load()+sync()`);
        }
        return t[k];
      },
      set(t, k, v) {
        log(`${label}.${k}=${JSON.stringify(v)}`);
        ctx.queue.push(() => (setters[k] ? setters[k](v) : (t[k] = v)));
        return true;
      },
    });
    return proxy;
  }

  function clientResult(compute) {
    let resolved = false;
    let v;
    ctx.queue.push(() => {
      v = compute();
      resolved = true;
    });
    return {
      get value() {
        if (!resolved) throw new Error("ClientResult.value read before sync()");
        return v;
      },
    };
  }

  // A collection whose `items` appear at the sync after load("items[/child,...]").
  function collection(label, build, childSeed) {
    const impl = {};
    const loadFn = (spec) => {
      (model.loads ||= []).push([label, String(spec)]);
      const names = String(spec).split(",").map((s) => s.trim());
      const children = names.filter((n) => n.startsWith("items/")).map((n) => n.slice(6));
      ctx.queue.push(() => {
        impl.__items = build();
        children.forEach((c) => impl.__items.forEach((it) => it.__seed([c])));
        if (childSeed) impl.__items.forEach((it) => it.__seed(childSeed));
        impl.__itemsLoaded = true;
      });
    };
    return {
      load: loadFn,
      get items() {
        if (!impl.__itemsLoaded) throw new Error(`${label}.items read before load()+sync()`);
        return impl.__items;
      },
    };
  }

  function removeTable(t) {
    model.tables.splice(model.tables.indexOf(t), 1);
    model.order.splice(model.order.findIndex((i) => i.ref === t), 1);
  }

  function emptyCollection(label) {
    return collection(label, () => []);
  }

  // A table's own tracked changes. Like Word, rejecting a tracked table insertion removes the table.
  function tableChanges(t) {
    const c = collection("table.trackedChanges", () => t.revs.map((r) => ({ ...r, load() {}, __seed() {} })));
    c.acceptAll = () => {
      log("table.trackedChanges.acceptAll()");
      ctx.queue.push(() => (t.revs = []));
    };
    c.rejectAll = () => {
      log("table.trackedChanges.rejectAll()");
      ctx.queue.push(() => {
        if (t.revs.some((r) => r.tableInsert)) removeTable(t);
        else t.revs = [];
        if (model.scopedRejectTouchesOutside) model.paragraphs.forEach((p) => (p.revs = []));
      });
    };
    return c;
  }

  function range(table) {
    return {
      compareLocationWith: (other) => clientResult(() => (other.__table === table ? "Equal" : "Before")),
      getOoxml: () => clientResult(() => table.ooxml),
      paragraphs: collection("table.range.paragraphs", () =>
        table.rows.flatMap((r) => r.flatMap((c) => c.paras.map((text) => makePara2({ text, __cell: c }, table.nesting))))
      ),
      getComments: () => emptyCollection("range.comments"),
      getTrackedChanges: () => tableChanges(table),
      __table: table,
    };
  }

  function makeCellBody(cell) {
    const fontImpl = {
      get bold() { return cell.bold; },
      get color() { return cell.color; },
      get size() { return cell.size; },
    };
    const font = node("cell.body.font", ["bold", "color", "size"], fontImpl, {
      bold: (v) => (cell.bold = v),
      color: (v) => (cell.color = v),
    });
    // Word's real rules that issue #75 tripped over (model.macCellQuirk, on by
    // default): insertBreak accepts only Before/After, and insertParagraph's
    // returned proxy in a table cell resolves to the PREVIOUS paragraph.
    const quirk = () => model.macCellQuirk !== false;
    const invalidArgument = (what) => {
      const e = new Error(`${what}: InvalidArgument`);
      e.code = "InvalidArgument";
      return e;
    };
    const makeRange = (idxRef) => ({
      font: { set bold(v) {}, set italic(v) {} },
      set hyperlink(v) {},
      insertBreak: (type, where) => {
        log(`range.insertBreak(${type},${where})`);
        ctx.queue.push(() => {
          if (where !== "Before" && where !== "After") throw invalidArgument("range.insertBreak");
          cell.paras[idxRef.i] += "\v";
        });
      },
    });
    const makePara = (idxRef) => {
      const pImpl = {
        get text() { return cell.paras[idxRef.i]; },
        insertText: (text, where) => {
          log(`paragraph.insertText(${JSON.stringify(text)},${where})`);
          ctx.queue.push(() => {
            if (model.dropParagraphText === idxRef.i) return; // simulate a write Word loses
            cell.paras[idxRef.i] += text;
          });
          return makeRange(idxRef);
        },
        insertBreak: (type, where) => {
          log(`paragraph.insertBreak(${type},${where})`);
          ctx.queue.push(() => {
            if (where !== "Before" && where !== "After") throw invalidArgument("paragraph.insertBreak");
            cell.paras[idxRef.i] += "\v";
          });
        },
        insertParagraph: (text, where) => {
          log(`paragraph.insertParagraph(${JSON.stringify(text)},${where})`);
          const ref2 = { i: null };
          ctx.queue.push(() => {
            cell.paras.splice(idxRef.i + 1, 0, text);
            ref2.i = idxRef.i + 1;
          });
          return makePara(ref2);
        },
        getRange: () => makeRange(idxRef),
        delete: () => {
          log("paragraph.delete()");
          ctx.queue.push(() => cell.paras.splice(idxRef.i, 1));
        },
      };
      return node("paragraph", ["text"], pImpl);
    };
    const paragraphs = collection("cell.body.paragraphs", () => cell.paras.map((_, i) => makePara({ i })));
    paragraphs.getFirst = () => makePara({ i: 0 });
    // Real Word (observed, issue #75): a getLast() queued in the same batch as an insert resolves
    // against the paragraphs as of the last sync, so writes through it land in an earlier paragraph.
    paragraphs.getLast = () => makePara({ get i() { return quirk() ? Math.min(cell.paras.length, cell.__synced ?? cell.paras.length) - 1 : cell.paras.length - 1; } });
    const impl = {
      getRange: () => ({}),
      get text() {
        const value = cellText(cell);
        // fires only for reads the PANE makes (the body-hash computation goes
        // through bodyText(), not here), so a test can model "a co-author
        // edits this cell right after the pane's Nth read of it".
        if (model.onCellTextRead) model.onCellTextRead(cell);
        return value;
      },
      font,
      clear: () => {
        log("cell.body.clear()");
        ctx.queue.push(() => (cell.paras = [""]));
      },
      paragraphs,
      insertParagraph: (text, where) => {
        log(`cell.body.insertParagraph(${JSON.stringify(text)},${where})`);
        const ref = { i: null };
        ctx.queue.push(() => {
          cell.paras.push(text);
          if (model.insertParagraphDoublesEmpty) cell.paras.push(""); // a Word build whose insert adds an extra mark
          ref.i = quirk() ? cell.paras.length - 2 : cell.paras.length - 1;
        });
        return makePara(ref);
      },
      insertOoxml: (_xml, where) => {
        log(`cell.body.insertOoxml(${where})`);
        // Word's getOoxml round-trip leaves one extra empty paragraph.
        ctx.queue.push(() => (cell.paras = [...(cell.__snap || [""]), ""]));
      },
    };
    impl.getOoxml = () =>
      clientResult(() => {
        cell.__snap = cell.paras.slice();
        return EMPTY;
      });
    return node("cell.body", ["text"], impl);
  }

  function makeCell(cell) {
    const impl = {
      get shadingColor() { return cell.fill; },
      get horizontalAlignment() { return cell.hAlign; },
      get verticalAlignment() { return cell.vAlign; },
      get columnWidth() { return cell.width; },
      body: makeCellBody(cell),
      __cell: cell,
    };
    return node("cell", ["shadingColor", "horizontalAlignment", "verticalAlignment", "columnWidth"], impl, {
      shadingColor: (v) => (cell.fill = v),
      horizontalAlignment: (v) => (cell.hAlign = v),
      verticalAlignment: (v) => (cell.vAlign = v),
      columnWidth: (v) => (cell.width = v),
    });
  }

  function makeRow(row) {
    return { cells: collection("row.cells", () => row.map(makeCell)) };
  }

  function makeTable(t) {
    const impl = {
      get style() { return t.style; },
      get headerRowCount() { return t.headerRowCount; },
      get nestingLevel() { return t.nesting; },
      rows: collection("table.rows", () => t.rows.map(makeRow)),
      font: { set size(v) { log(`table.font.size=${v}`); ctx.queue.push(() => t.rows.forEach((r) => r.forEach((c) => (c.size = v)))); } },
      getRange: () => range(t),
      insertTable: (r, c, where, values) => insertTable(`table.insertTable`, where, r, c, () => model.order.findIndex((i) => i.ref === t) + 1, null),
      delete: () => {
        log("table.delete()");
        ctx.queue.push(() => {
          if (model.dropTableDelete) return;
          if (model.onTableDelete) model.onTableDelete();
          if (model.mode === "TrackAll") t.revs.push({ kind: "del", author: "Me", text: "" });
          else removeTable(t);
        });
      },
      __table: t,
    };
    return node("table", ["style", "headerRowCount", "nestingLevel"], impl, {
      style: (v) => (t.style = v),
      styleBuiltIn: (v) => (t.style = v),
      headerRowCount: (v) => (t.headerRowCount = v),
    });
  }

  function insertTable(label, where, rowCount, colCount, indexFn, anchorPara) {
    log(`${label}(${rowCount},${colCount},${where})`);
    const t = { style: "Table Grid", headerRowCount: 0, nesting: 1, rows: Array.from({ length: rowCount }, () => Array.from({ length: colCount }, () => newCell(""))), ooxml: EMPTY, revs: [] };
    ctx.queue.push(() => {
      if (model.failOnInsert) throw new Error("simulated Word failure while inserting");
      // issue #79: cells written next to a list item inherit its numbering; a tracked insert is one revision
      if (anchorPara && anchorPara.listed) t.rows.forEach((r) => r.forEach((c) => (c.listed = true)));
      if (model.mode === "TrackAll") t.revs.push({ kind: "ins", author: "Me", text: "", tableInsert: true });
      model.tables.push(t);
      model.order.splice(indexFn(), 0, { kind: "t", ref: t });
    });
    return makeTable(t);
  }

  const makePara2 = (p, level) => makeParagraph(p, level);
  const holder = (p) => p.__cell || p;
  function makeParagraph(p, level = 0) {
    const who = p.watch ? "anchor" : "paragraph";
    const impl = {
      get text() { return p.text; },
      get tableNestingLevel() { return level; },
      get isListItem() { return !!holder(p).listed; },
      detachFromList: () => {
        log(`${who}.detachFromList()`);
        ctx.queue.push(() => { if (!model.dropDetach) holder(p).listed = false; });
      },
      getRange: () => ({
        compareLocationWith: () => clientResult(() => "Before"),
        getOoxml: () => clientResult(() => EMPTY),
        paragraphs: emptyCollection("paragraph.range.paragraphs"),
        getComments: () => emptyCollection("range.comments"),
        getTrackedChanges: () => emptyCollection("range.trackedChanges"),
        font: node(`${who}.font`, [], {}),
      }),
      insertTable: (r, c, where) =>
        insertTable("paragraph.insertTable", where, r, c, () => {
          const at = model.order.findIndex((i) => i.ref === p);
          return where === "Before" ? at : at + 1;
        }, p),
    };
    return node(who, ["text", "tableNestingLevel", "isListItem"], impl);
  }

  const body = node(
    "body",
    ["text"],
    {
      getOoxml: () => clientResult(() => {
        if (model.ooxmlUnreadable) throw new Error("simulated unreadable body OOXML");
        return model.ooxmlMode === "revisions" ? revisionOoxml(model) : EMPTY;
      }),
      get text() { return bodyText(model); },
      // Like Word, body.paragraphs includes the paragraphs INSIDE table cells
      // (tableNestingLevel 1), in document order.
      paragraphs: collection("body.paragraphs", () =>
        model.order.flatMap((item) =>
          item.kind === "p"
            ? [makeParagraph(item.ref, 0)]
            : item.ref.rows.flatMap((r) => r.flatMap((c) => c.paras.map((text) => makeParagraph({ text, __cell: c }, item.ref.nesting))))
        )
      ),
      tables: collection("body.tables", () => topLevelTables(model).map(makeTable)),
      insertTable: (r, c, where) =>
        insertTable("body.insertTable", where, r, c, () => (where === "Start" ? 0 : model.order.length)),
    }
  );

  const document = node(
    "document",
    ["changeTrackingMode"],
    {
      get changeTrackingMode() { return model.mode; },
      body,
      getStyles: () => ({
        getByNameOrNullObject: (name) =>
          node("style", ["isNullObject"], { get isNullObject() { return !model.styles.has(name); } }),
      }),
      save: () => {},
    },
    { changeTrackingMode: (v) => (model.mode = v) }
  );

  ctx.document = document;
  ctx.sync = async () => {
    ctx.syncs += 1;
    const q = ctx.queue;
    ctx.queue = [];
    for (const fn of q) fn();
    // what a later batch's not-yet-synced getLast() resolves against (issue #75)
    model.tables.forEach((t) => t.rows.forEach((r) => r.forEach((c) => (c.__synced = c.paras.length))));
  };
  return ctx;
}

// ------------------------------------------------------------------- sandbox

function loadPane(model) {
  const sandbox = {
    DOMParser,
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
    Office: {
      onReady() {},
      HostType: { Word: "Word" },
      context: { requirements: { isSetSupported: (_n, v) => (v === "1.5" ? model.supports15 : true) } },
    },
    Word: {
      run: (fn) => fn(makeContext(model)),
      InsertLocation: { end: "End", replace: "Replace", before: "Before", after: "After" },
      BreakType: { line: "Line" },
      RangeLocation: { whole: "Whole" },
      ChangeTrackingMode: { trackAll: "TrackAll", off: "Off" },
      BuiltInStyleName: { gridTable4_Accent1: "GridTable4_Accent1", tableGrid: "TableGrid" },
    },
  };
  vm.createContext(sandbox);
  vm.runInContext(
    src +
      "\n;globalThis.__ops = { opTableInsert, opTableGet, opCellsSet, opCellSet, opCellGet, opParagraphsUnlist, opTableDelete, opRevisionsMutate, dispatchOp, PANE_CAPABILITIES };",
    sandbox
  );
  sandbox.__ops.scopedTest = (onGuard) => {
    sandbox.onGuard = onGuard;
    vm.runInContext(`executionSnapshot = async () => ({sections:[{slug:'target'}]});
      sectionGuard = async () => ({active:true, check:async()=>{}});
      targetTables = async context => loadTopLevelTables(context);
      guardComments = async () => { if (globalThis.onGuard) globalThis.onGuard(); };
      guardRevisions = async () => {};`, sandbox);
  };
  return sandbox.__ops;
}

async function sha(text) {
  const d = await globalThis.crypto.subtle.digest("SHA-256", new TextEncoder().encode(text));
  return Array.from(new Uint8Array(d)).map((b) => b.toString(16).padStart(2, "0")).join("");
}

async function run(fn) {
  try {
    return { ok: true, value: await fn() };
  } catch (err) {
    return { ok: false, code: err.code, message: err.message };
  }
}

const para = (text, bold = false) => [{ text, bold, italic: false, link: null, hard_break: false }];
const spec = (text, extra = {}) => ({ paragraphs: text === "" ? [] : [para(text)], fill: null, color: null, bold: null, align: null, valign: null, ...extra });
const writes = (model) => model.calls.filter((c) => /insertTable|insertText|clear\(\)|insertParagraph|\.style=|\.styleBuiltIn=|\.shadingColor=|Alignment=|columnWidth=|font\./.test(c));

// ----------------------------------------------------------------- scenarios

// 1. happy path: every property reaches Word with Word's enum names
{
  const model = makeModel({ styles: ["Table Grid", "Grid Table 4 - Accent 1"] });
  const ops = loadPane(model);
  const pre = await sha(bodyText(model));
  const r = await run(() =>
    ops.opTableInsert({
      rows: [
        [spec("Head", { fill: "1F3864", color: "FFFFFF", bold: true, align: "center", valign: "center" }), spec("Value", { align: "both", valign: "bottom" })],
        [spec("a", { align: "right", valign: "top" }), spec("b", { align: "left" })],
      ],
      anchor: { paragraph_text: "Proof", position: "after" },
      style: "Grid Table 4 - Accent 1",
      header_rows: 1,
      column_widths_pt: [100, 200],
      font_size_pt: 10,
      track_changes: false,
      expectedBodySha256: pre,
    })
  );
  check("insert: succeeds", r.ok, r.message);
  const t = model.tables[0];
  check("insert: table placed after the anchor paragraph", model.order[2].ref === t && model.order[1].ref.text === "Proof", JSON.stringify(model.order.map((i) => i.kind)));
  check("insert: returns the new table's 1-based index", r.ok && r.value.table_index === 1, JSON.stringify(r.value));
  check("insert: reports pre/post body hashes", r.ok && r.value.pre === pre && typeof r.value.post === "string" && r.value.post !== pre, JSON.stringify(r.value));
  check("insert: used paragraph.insertTable(...,After)", model.calls.includes("paragraph.insertTable(2,2,After)"), model.calls.join(" | "));
  check("insert: style applied", t.style === "Grid Table 4 - Accent 1" && t.headerRowCount === 1, `${t.style}/${t.headerRowCount}`);
  check("insert: cell text written", cellText(t.rows[0][0]) === "Head" && cellText(t.rows[1][1]) === "b", JSON.stringify(t.rows));
  const h = t.rows[0][0];
  check("insert: fill/color/bold applied", h.fill === "#1F3864" && h.color === "#FFFFFF" && h.bold === true, JSON.stringify(h));
  check("insert: alignment enums are Word's (Centered/Justified/Right/Left)", h.hAlign === "Centered" && t.rows[0][1].hAlign === "Justified" && t.rows[1][0].hAlign === "Right" && t.rows[1][1].hAlign === "Left", JSON.stringify(t.rows.map((r) => r.map((c) => c.hAlign))));
  check("insert: vertical alignment enums (Center/Bottom/Top)", h.vAlign === "Center" && t.rows[0][1].vAlign === "Bottom" && t.rows[1][0].vAlign === "Top", JSON.stringify(t.rows.map((r) => r.map((c) => c.vAlign))));
  check("insert: column widths per column", t.rows[0][0].width === 100 && t.rows[1][1].width === 200, JSON.stringify(t.rows.map((r) => r.map((c) => c.width))));
  check("insert: font size on the table", t.rows.every((r) => r.every((c) => c.size === 10)), JSON.stringify(t.rows.map((r) => r.map((c) => c.size))));
}

// 2. position before / append / after_table_index; index via compareLocationWith(.value)
{
  const model = makeModel({ tables: [{ rows: [["T1"]] }, { rows: [["T2"]] }] });
  const ops = loadPane(model);
  const r = await run(() => ops.opTableInsert({ rows: [[spec("new")]], anchor: { paragraph_text: "Tail", position: "before" }, header_rows: 0 }));
  // existing tables sit at the end of the order, so "before Tail" lands before both
  check("insert: position before -> index 1, existing tables shift", r.ok && r.value.table_index === 1 && cellText(model.tables[2].rows[0][0]) === "new", JSON.stringify(r));
  const r2 = await run(() => ops.opTableInsert({ rows: [[spec("end")]], anchor: null, header_rows: 0 }));
  check("insert: no anchor appends via body.insertTable(End)", r2.ok && model.calls.includes("body.insertTable(1,1,End)") && r2.value.table_index === 4, JSON.stringify(r2));
  const r3 = await run(() => ops.opTableInsert({ rows: [[spec("mid")]], anchor: { after_table_index: 2 }, header_rows: 0 }));
  check("insert: after_table_index uses table.insertTable(After)", r3.ok && model.calls.includes("table.insertTable(1,1,After)") && r3.value.table_index === 3, JSON.stringify(r3));
  const r4 = await run(() => ops.opTableInsert({ rows: [[spec("x")]], anchor: { after_table_index: 99 }, header_rows: 0 }));
  check("insert: after_table_index out of range -> table_not_found", !r4.ok && r4.code === "table_not_found", JSON.stringify(r4));
}

// 3. stale body hash refuses BEFORE any write
{
  const model = makeModel();
  const ops = loadPane(model);
  const r = await run(() => ops.opTableInsert({ rows: [[spec("x")]], anchor: null, header_rows: 0, expectedBodySha256: "0".repeat(64) }));
  check("stale: refused with code 'stale'", !r.ok && r.code === "stale", JSON.stringify(r));
  check("stale: nothing was queued for writing", writes(model).length === 0 && model.tables.length === 0, model.calls.join(" | "));
}

// 4. styles are validated BEFORE inserting
{
  const model = makeModel({ supports15: false });
  const ops = loadPane(model);
  const r = await run(() => ops.opTableInsert({ rows: [[spec("x")]], anchor: null, style: "Nope Style", header_rows: 0 }));
  check("style: unknown name without WordApi 1.5 -> style_not_found", !r.ok && r.code === "style_not_found", JSON.stringify(r));
  check("style: nothing inserted on a bad style", writes(model).length === 0 && model.tables.length === 0, model.calls.join(" | "));

  const model2 = makeModel({ supports15: true, styles: ["Table Grid"] });
  const r2 = await run(() => loadPane(model2).opTableInsert({ rows: [[spec("x")]], anchor: null, style: "Nope Style", header_rows: 0 }));
  check("style: name absent from the style list (1.5) -> style_not_found, nothing inserted", !r2.ok && r2.code === "style_not_found" && model2.tables.length === 0, JSON.stringify(r2));

  const model3 = makeModel({ supports15: false, tables: [{ style: "Fancy", rows: [["T"]] }] });
  const r3 = await run(() => loadPane(model3).opTableInsert({ rows: [[spec("x")]], anchor: null, style: "Fancy", header_rows: 0 }));
  check("style: a name already used by a table is accepted without 1.5", r3.ok && model3.tables[1].style === "Fancy", JSON.stringify(r3));

  const r4 = await run(() => loadPane(model3).opTableInsert({ rows: [[spec("x")]], anchor: null, style_from_table_index: 1, header_rows: 0 }));
  check("style: style_from_table_index copies the source style", r4.ok && model3.tables[model3.tables.length - 1].style === "Fancy", JSON.stringify(r4));
  const r5 = await run(() => loadPane(model3).opTableInsert({ rows: [[spec("x")]], anchor: null, style_from_table_index: 9, header_rows: 0 }));
  check("style: style_from_table_index out of range -> style_not_found", !r5.ok && r5.code === "style_not_found", JSON.stringify(r5));

  const model6 = makeModel();
  const r6 = await run(() => loadPane(model6).opTableInsert({ rows: [[spec("x")]], anchor: null, style_builtin: "gridTable4_Accent1", header_rows: 0 }));
  check("style: built-in enum key maps to Word's BuiltInStyleName value", r6.ok && model6.tables[0].style === "GridTable4_Accent1", JSON.stringify(r6));
  const r7 = await run(() => loadPane(model6).opTableInsert({ rows: [[spec("x")]], anchor: null, style_builtin: "bogus", header_rows: 0 }));
  check("style: unknown built-in -> style_not_found", !r7.ok && r7.code === "style_not_found" && model6.tables.length === 1, JSON.stringify(r7));
}

// 5. anchors
{
  const RSQUO = String.fromCharCode(0x2019);
  const SHY = String.fromCharCode(0x00ad);
  const model = makeModel({ paragraphs: ["Dup", "Dup", "Once", `Smart ${RSQUO}quote${RSQUO}  here`, `soft${SHY}hyphen`] });
  const ops = loadPane(model);
  const r1 = await run(() => ops.opTableInsert({ rows: [[spec("x")]], anchor: { paragraph_text: "Dup", position: "after" }, header_rows: 0 }));
  check("anchor: >1 match -> match_count_mismatch", !r1.ok && r1.code === "match_count_mismatch", JSON.stringify(r1));
  const r2 = await run(() => ops.opTableInsert({ rows: [[spec("x")]], anchor: { paragraph_text: "Nope", position: "after" }, header_rows: 0 }));
  check("anchor: 0 matches -> zero_match", !r2.ok && r2.code === "zero_match", JSON.stringify(r2));
  check("anchor: refusals insert nothing", model.tables.length === 0 && writes(model).length === 0, model.calls.join(" | "));
  const r3 = await run(() => ops.opTableInsert({ rows: [[spec("x")]], anchor: { paragraph_text: "  Smart 'quote' here ", position: "after" }, header_rows: 0 }));
  check("anchor: matched modulo quotes/whitespace", r3.ok, JSON.stringify(r3));
  const r4 = await run(() => ops.opTableInsert({ rows: [[spec("x")]], anchor: { paragraph_text: "softhyphen", position: "after" }, header_rows: 0 }));
  check("anchor: soft hyphens are ignored when matching", r4.ok, JSON.stringify(r4));

  // a paragraph INSIDE a table cell is never an anchor candidate
  const model2 = makeModel({ paragraphs: ["Proof", "Tail"], tables: [{ rows: [["Proof", "other"]] }] });
  const r5 = await run(() => loadPane(model2).opTableInsert({ rows: [[spec("x")]], anchor: { paragraph_text: "Proof", position: "after" }, header_rows: 0 }));
  check("anchor: in-table paragraphs are not candidates (still exactly 1 match)", r5.ok, JSON.stringify(r5));
}

// 6. change tracking: forced on, restored after success AND after a throwing sync
{
  const model = makeModel({ mode: "Off" });
  const ops = loadPane(model);
  const r = await run(() => ops.opTableInsert({ rows: [[spec("x")]], anchor: null, header_rows: 0, track_changes: true }));
  check("tracking: previous mode restored after a tracked insert", r.ok && model.mode === "Off", `${model.mode} ${JSON.stringify(r)}`);
  check("tracking: trackAll was set during the insert", model.calls.indexOf('document.changeTrackingMode="TrackAll"') >= 0 && model.calls.indexOf('document.changeTrackingMode="TrackAll"') < model.calls.indexOf("body.insertTable(1,1,End)"), model.calls.join(" | "));

  const model2 = makeModel({ mode: "Off", failOnInsert: true });
  const r2 = await run(() => loadPane(model2).opTableInsert({ rows: [[spec("x")]], anchor: null, header_rows: 0, track_changes: true }));
  check("tracking: restored even when the sync throws", !r2.ok && model2.mode === "Off", `${model2.mode} ${JSON.stringify(r2)}`);

  const model3 = makeModel({ mode: "TrackAll" });
  const r3 = await run(() => loadPane(model3).opTableInsert({ rows: [[spec("x")]], anchor: null, header_rows: 0, track_changes: true }));
  check("tracking: a document already on TrackAll stays on TrackAll", r3.ok && model3.mode === "TrackAll", model3.mode);

  const model4 = makeModel({ mode: "Off" });
  const r4 = await run(() => loadPane(model4).opTableInsert({ rows: [[spec("x")]], anchor: null, header_rows: 0, track_changes: false }));
  check("tracking: not touched when track_changes is false", r4.ok && !model4.calls.some((c) => c.includes("changeTrackingMode")), model4.calls.join(" | "));
}

// 7. nested table documents are refused
{
  const model = makeModel({ tables: [{ rows: [["x"]], nesting: 2 }] });
  const ops = loadPane(model);
  const r = await run(() => ops.opTableInsert({ rows: [[spec("x")]], anchor: null, header_rows: 0 }));
  check("nested: insert refused", !r.ok && /nested/.test(r.message), JSON.stringify(r));
  const g = await run(() => ops.opTableGet({ table_index: 1 }));
  check("nested: table_get refused", !g.ok && /nested/.test(g.message), JSON.stringify(g));
}

// 8. table_get
{
  const model = makeModel({
    tables: [
      { style: "Fancy", headerRowCount: 1, rows: [["a", "b"], ["c", "d"]], hAlign: "Centered", ooxml: '<pkg:package><pkg:part pkg:name="/word/styles.xml"><w:style><w:tcPr><w:gridSpan w:val="2"/></w:tcPr></w:style></pkg:part><pkg:part pkg:name="/word/document.xml"><w:tbl><w:tr/></w:tbl></pkg:part></pkg:package>' },
      { rows: [["m"]], ooxml: '<pkg:package><pkg:part pkg:name="/word/document.xml"><w:tbl><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr></w:tc></w:tbl></pkg:part></pkg:package>' },
      { rows: [["v"]], ooxml: '<pkg:package><pkg:part pkg:name="/word/document.xml"><w:tbl><w:tcPr><w:vMerge w:val="restart"/></w:tcPr></w:tbl></pkg:part></pkg:package>' },
    ],
  });
  model.tables[0].rows[0][0].bold = true;
  model.tables[0].rows[0][0].color = "#FF0000";
  model.tables[0].rows[0][0].fill = "#112233";
  const ops = loadPane(model);
  const g = await run(() => ops.opTableGet({ table_index: 1 }));
  check("get: ok and style/header reported", g.ok && g.value.style === "Fancy" && g.value.headerRowCount === 1, JSON.stringify(g));
  check("get: alignment reported in the SERVER's vocabulary", g.ok && g.value.rows[0][0].align === "center" && g.value.rows[0][0].valign === "top", JSON.stringify(g.value && g.value.rows[0][0]));
  check("get: cell fields", g.ok && g.value.rows[0][0].text === "a" && g.value.rows[0][0].bold === true && g.value.rows[0][0].color === "#FF0000" && g.value.rows[0][0].fill === "#112233" && g.value.rows[1][1].text === "d", JSON.stringify(g.value));
  check("get: gridSpan in the STYLES part is not a merged cell", g.ok && g.value.merged === false, JSON.stringify(g.value && g.value.merged));
  const gm = await run(() => ops.opTableGet({ table_index: 2 }));
  check("get: gridSpan in the table's own part -> merged", gm.ok && gm.value.merged === true, JSON.stringify(gm));
  const gv = await run(() => ops.opTableGet({ table_index: 3 }));
  check("get: vMerge -> merged", gv.ok && gv.value.merged === true, JSON.stringify(gv));
  const gn = await run(() => ops.opTableGet({ table_index: 9 }));
  check("get: missing table -> table_not_found", !gn.ok && gn.code === "table_not_found", JSON.stringify(gn));
}

// 9. cells_set: CAS with re-read, stale precondition, atomic refusal
{
  const mk = () => makeModel({ tables: [{ rows: [["a", "b"], ["c", "d"]] }] });
  const cells = (before = ["a", "b"]) => [
    { table_index: 1, row_index: 1, cell_index: 1, paragraphs: [para("A2")], expected_before_text: before[0] },
    { table_index: 1, row_index: 1, cell_index: 2, paragraphs: [para("B2")], expected_before_text: before[1] },
  ];
  const model = mk();
  const ops = loadPane(model);
  const pre = await sha(bodyText(model));
  const r = await run(() => ops.opCellsSet({ cells: cells(), track_changes: false, expectedBodySha256: pre }));
  check("cells_set: writes every cell and reports before/after", r.ok && cellText(model.tables[0].rows[0][0]) === "A2" && cellText(model.tables[0].rows[0][1]) === "B2" && r.value.before.join() === "a,b" && r.value.after.join() === "A2,B2", JSON.stringify(r));
  check("cells_set: other cells untouched", cellText(model.tables[0].rows[1][0]) === "c", JSON.stringify(model.tables[0].rows));

  const model2 = mk();
  const r2 = await run(() => loadPane(model2).opCellsSet({ cells: cells(["a", "WRONG"]), track_changes: false }));
  check("cells_set: a changed cell refuses the WHOLE batch, nothing written", !r2.ok && /changed since it was read/.test(r2.message) && writes(model2).length === 0 && cellText(model2.tables[0].rows[0][0]) === "a", JSON.stringify(r2));

  const model3 = mk();
  const r3 = await run(() => loadPane(model3).opCellsSet({ cells: cells(), track_changes: false, expectedBodySha256: "f".repeat(64) }));
  check("cells_set: stale body hash refused before any write", !r3.ok && r3.code === "stale" && writes(model3).length === 0, JSON.stringify(r3));

  // a co-author edits a cell BETWEEN the first read and the pre-write re-read
  // The co-author's edit lands right AFTER the pane's first read of that cell,
  // i.e. exactly between the CAS read and the pre-write re-read.
  const model4 = mk();
  const ops4 = loadPane(model4);
  const target = model4.tables[0].rows[0][1];
  let paneReads = 0;
  model4.onCellTextRead = (cell) => {
    if (cell !== target) return;
    paneReads += 1;
    if (paneReads === 1) cell.paras = ["someone typed"]; // lands after the pane's first read
  };
  const r4 = await run(() => ops4.opCellsSet({ cells: cells(), track_changes: false }));
  check("cells_set: an edit between the first read and the pre-write re-read is refused", !r4.ok && /changed since it was read/.test(r4.message), JSON.stringify(r4));
  check("cells_set: ... with nothing written (not even the unchanged cell)", writes(model4).length === 0 && cellText(model4.tables[0].rows[0][0]) === "a", model4.calls.join(" | "));
}

// 10. cell_set still works after the writeCellParagraphs refactor
{
  const model = makeModel({ tables: [{ rows: [["old", "keep"]] }] });
  const ops = loadPane(model);
  const r = await run(() =>
    ops.opCellSet({ table_index: 1, row_index: 1, cell_index: 1, paragraphs: [para("Hi", true), para("there")], expected_before_text: "old", track_changes: false })
  );
  check("cell_set: multi-paragraph write via the shared helper", r.ok && model.tables[0].rows[0][0].paras.join("|") === "Hi|there" && r.value.after === "Hi\nthere", JSON.stringify(r) + JSON.stringify(model.tables[0].rows));
  const r2 = await run(() =>
    ops.opCellSet({ table_index: 1, row_index: 1, cell_index: 2, paragraphs: [para("x")], expected_before_text: "stale", track_changes: false })
  );
  check("cell_set: compare-and-set still refuses a changed cell", !r2.ok && /changed since it was read/.test(r2.message) && cellText(model.tables[0].rows[0][1]) === "keep", JSON.stringify(r2));
}

// 11. wiring
{
  const model = makeModel({ tables: [{ rows: [["a"]] }] });
  const ops = loadPane(model);
  check("capabilities: table_edit is reported", ops.PANE_CAPABILITIES.includes("table_edit"), JSON.stringify(ops.PANE_CAPABILITIES));
  const g = await run(() => ops.dispatchOp("table_get", { table_index: 1 }));
  check("dispatch: table_get routed", g.ok && g.value.rows[0][0].text === "a", JSON.stringify(g));
  const i = await run(() => ops.dispatchOp("table_insert", { rows: [[spec("n")]], anchor: null, header_rows: 0 }));
  check("dispatch: table_insert routed", i.ok && i.value.applied === true, JSON.stringify(i));
  const c = await run(() => ops.dispatchOp("cells_set", { cells: [{ table_index: 1, row_index: 1, cell_index: 1, paragraphs: [para("z")], expected_before_text: "a" }], track_changes: false }));
  check("dispatch: cells_set routed", c.ok && c.value.applied === true, JSON.stringify(c));
}

// 12. issue #75: multi-paragraph / hard-break writes under real Word's quirks
{
  const hb = { text: "", bold: false, italic: false, link: null, hard_break: true };
  const set = (ops, paragraphs, extra = {}) =>
    run(() => ops.opCellSet({ table_index: 1, row_index: 1, cell_index: 1, paragraphs, expected_before_text: "old", track_changes: false, ...extra }));

  // (a) two paragraphs: no merge, no trailing empty paragraph
  const m1 = makeModel({ tables: [{ rows: [["old"]] }] });
  const r1 = await set(loadPane(m1), [para("Session 2. Build a long list."), para("Pre-read: SOO 4.3", false)]);
  check("cell_set #75: two paragraphs stay two paragraphs", r1.ok && m1.tables[0].rows[0][0].paras.join("|") === "Session 2. Build a long list.|Pre-read: SOO 4.3", JSON.stringify(r1) + JSON.stringify(m1.tables[0].rows));

  // (b) hard break: Before/After only, lands inside one paragraph
  const m2 = makeModel({ tables: [{ rows: [["old"]] }] });
  const r2 = await set(loadPane(m2), [[...para("line one"), hb, ...para("line two")]]);
  check("cell_set #75: hard break inserts a line break in one paragraph", r2.ok && m2.tables[0].rows[0][0].paras.length === 1 && m2.tables[0].rows[0][0].paras[0] === "line one\vline two", JSON.stringify(r2) + JSON.stringify(m2.tables[0].rows));
  check("cell_set #75: insertBreak uses After, never End", m2.calls.some((c) => c === "range.insertBreak(Line,After)") && !m2.calls.some((c) => /insertBreak\(Line,End\)/.test(c)), m2.calls.join(" | "));

  // 3 paragraphs
  const m2b = makeModel({ tables: [{ rows: [["old"]] }] });
  const r2b = await set(loadPane(m2b), [para("a"), para("b"), para("c")]);
  check("cell_set #75: three paragraphs", r2b.ok && m2b.tables[0].rows[0][0].paras.join("|") === "a|b|c", JSON.stringify(r2b) + JSON.stringify(m2b.tables[0].rows));

  // (c) a write Word loses on every strategy -> old text written back, typed code
  const m3 = makeModel({ tables: [{ rows: [["old"]] }] });
  m3.dropParagraphText = 1;
  const r3 = await set(loadPane(m3), [para("first"), para("second")]);
  check("cell_set #75: a failed self-check is rolled back with a typed code", !r3.ok && r3.code === "cell_write_rolled_back" && m3.tables[0].rows[0][0].paras.join("|") === "old", JSON.stringify(r3) + JSON.stringify(m3.tables[0].rows));
  check("cell_set #75: the restore never uses OOXML", !m3.calls.some((c) => /insertOoxml/.test(c)), m3.calls.join(" | "));

  // (d) the same under track_changes -> tried once, never rewritten
  const m4 = makeModel({ tables: [{ rows: [["old"]] }] });
  m4.dropParagraphText = 1;
  const r4 = await set(loadPane(m4), [para("first"), para("second")], { track_changes: true });
  check("cell_set #75: tracked write is not restored", !r4.ok && r4.code === "cell_write_not_rolled_back" && m4.calls.filter((c) => c === "cell.body.clear()").length === 1, JSON.stringify(r4) + m4.calls.join(" | "));

  // (c2) first strategy yields a stray empty paragraph -> the next strategy succeeds
  const m6 = makeModel({ tables: [{ rows: [["old"]] }] });
  m6.insertParagraphDoublesEmpty = true;
  const r6 = await set(loadPane(m6), [para("first"), para("second")]);
  check("cell_set #75: falls back to after-chain when load-items misreads", r6.ok && r6.value.strategy === "after-chain" && m6.tables[0].rows[0][0].paras.join("|") === "first|second", JSON.stringify(r6) + JSON.stringify(m6.tables[0].rows));
  const r6b = await set(loadPane(makeModel({ tables: [{ rows: [["old"]] }] })), [para("first"), para("second")]);
  check("cell_set #75: load-items is the first strategy", r6b.ok && r6b.value.strategy === "load-items", JSON.stringify(r6b));

  // (e) cell_get returns paragraphs
  const m5 = makeModel({ tables: [{ rows: [["x"]] }] });
  m5.tables[0].rows[0][0].paras = ["one", "two"];
  const r5 = await run(() => loadPane(m5).opCellGet({ table_index: 1, row_index: 1, cell_index: 1 }));
  check("cell_get #75: returns per-paragraph text", r5.ok && JSON.stringify(r5.value.paragraphs) === '["one","two"]' && r5.value.text === "one\ntwo", JSON.stringify(r5));

  // (f) capability
  check("capabilities #75: cell_multiline is reported", loadPane(m5).PANE_CAPABILITIES.includes("cell_multiline"), "");
}

// ------------------------------------------------- issue #79: table clean-up

const anchorWrites = (model) => model.calls.filter((c) => c.startsWith("anchor."));
const trackedCalls = (model) => model.calls.filter((c) => /trackedChanges\.(accept|reject)All|table\.delete/.test(c));

// 13. a table inserted after a list item must not keep the numbering
{
  const model = makeModel({ paragraphs: ["Intro", { text: "Step one", listed: true, watch: true }, "Tail"] });
  const ops = loadPane(model);
  const r = await run(() => ops.opTableInsert({ rows: [[spec("a"), spec("b")], [spec("c"), spec("d")]], anchor: { paragraph_text: "Step one", position: "after" }, header_rows: 0 }));
  check("numbering #79: insert after a list item succeeds", r.ok, r.message);
  check("numbering #79: every cell paragraph was detached from the list", model.calls.filter((c) => c === "paragraph.detachFromList()").length === 4 && model.tables[0].rows.every((row) => row.every((c) => !c.listed)), model.calls.join(" | "));
  check("numbering #79: list_items_in_table is reported as 0 after detaching", r.ok && r.value.list_items_in_table === 0, JSON.stringify(r.value));
  check("numbering #79: the anchor paragraph is never written", anchorWrites(model).length === 0, model.calls.join(" | "));

  const plain = makeModel();
  const r2 = await run(() => loadPane(plain).opTableInsert({ rows: [[spec("a")]], anchor: { paragraph_text: "Proof", position: "after" }, header_rows: 0 }));
  check("numbering #79: nothing to detach after a plain paragraph", r2.ok && r2.value.list_items_in_table === 0 && !plain.calls.includes("paragraph.detachFromList()"), JSON.stringify(r2) + plain.calls.join(" | "));

  const stuck = makeModel({ paragraphs: ["Intro", { text: "Step one", listed: true }] });
  stuck.dropDetach = true;
  const r3 = await run(() => loadPane(stuck).opTableInsert({ rows: [[spec("a"), spec("b")]], anchor: { paragraph_text: "Step one", position: "after" }, header_rows: 0 }));
  check("numbering #79: a detach that did not take is reported, not hidden", r3.ok && r3.value.list_items_in_table === 2, JSON.stringify(r3));
}

// 14. paragraphs_unlist
{
  const model = makeModel({ tables: [{ rows: [["a", "b"], ["c", "d"]], listed: true, ooxml: EMPTY }, { rows: [["x"]], ooxml: EMPTY }] });
  const ops = loadPane(model);
  const pre = await sha(bodyText(model));
  const r = await run(() => ops.dispatchOp("paragraphs_unlist", { table_index: 1, expectedBodySha256: pre }));
  check("unlist #79: table scope detaches every list-item cell paragraph", r.ok && r.value.unlisted === 4 && r.value.remaining_list_items === 0 && r.value.applied === true, JSON.stringify(r));
  check("unlist #79: reports pre == post (no text changed)", r.ok && r.value.pre === pre && r.value.post === pre, JSON.stringify(r));
  check("unlist #79: other tables untouched", model.calls.filter((c) => c === "paragraph.detachFromList()").length === 4, model.calls.join(" | "));
  const again = await run(() => ops.opParagraphsUnlist({ table_index: 1 }));
  check("unlist #79: a second run is a verified no-op", again.ok && again.value.unlisted === 0, JSON.stringify(again));
  const missing = await run(() => ops.opParagraphsUnlist({ table_index: 7 }));
  check("unlist #79: unknown table is table_not_found", !missing.ok && missing.code === "table_not_found", JSON.stringify(missing));

  const pm = makeModel({ paragraphs: ["Intro", { text: "Item", listed: true }, "Tail"] });
  const pr = await run(() => loadPane(pm).opParagraphsUnlist({ paragraph_text: "Item" }));
  check("unlist #79: single paragraph scope", pr.ok && pr.value.unlisted === 1 && pm.paragraphs[1].listed === false, JSON.stringify(pr));
  const dup = makeModel({ paragraphs: [{ text: "Item", listed: true }, { text: "Item", listed: true }, "Tail"] });
  const dr = await run(() => loadPane(dup).opParagraphsUnlist({ paragraph_text: "Item" }));
  check("unlist #79: ambiguous paragraph is MATCH_COUNT_MISMATCH with nothing written", !dr.ok && dr.code === "MATCH_COUNT_MISMATCH" && !dup.calls.includes("paragraph.detachFromList()"), JSON.stringify(dr));
  const none = await run(() => loadPane(dup).opParagraphsUnlist({ paragraph_text: "Nope" }));
  check("unlist #79: absent paragraph is MATCH_COUNT_MISMATCH", !none.ok && none.code === "MATCH_COUNT_MISMATCH", JSON.stringify(none));

  const tracked = makeModel({ tables: [{ rows: [["a"]], listed: true, ooxml: EMPTY, revs: [{ kind: "ins", author: "Colleague", text: "a" }] }] });
  const tr = await run(() => loadPane(tracked).opParagraphsUnlist({ table_index: 1 }));
  check("unlist #79: tracked changes in the table refuse before any write", !tr.ok && tr.code === "TRACKED_CHANGES_PRESENT" && !tracked.calls.includes("paragraph.detachFromList()"), JSON.stringify(tr));

  const stuck = makeModel({ tables: [{ rows: [["a"]], listed: true, ooxml: EMPTY }] });
  stuck.dropDetach = true;
  const sr = await run(() => loadPane(stuck).opParagraphsUnlist({ table_index: 1 }));
  check("unlist #79: a detach that did not take is VERIFICATION_FAILED", !sr.ok && sr.code === "VERIFICATION_FAILED", JSON.stringify(sr));

  const locked = makeModel({ tables: [{ rows: [["a"]], listed: true, ooxml: EMPTY }] });
  const lr = await run(() => loadPane(locked).dispatchOp("paragraphs_unlist", { table_index: 1, clientId: "c", expectedBodySha256: "wrong" }));
  check("unlist #79: a stale baseline refuses before any write", !lr.ok && lr.code === "stale" && !locked.calls.includes("paragraph.detachFromList()"), JSON.stringify(lr));
}

// 15. table_delete
{
  const mk = (extra = {}) => makeModel({ tables: [{ rows: [["R1C1", "R1C2"], ["R2C1", "R2C2"]], ooxml: EMPTY, ...extra }, { rows: [["keep"]], ooxml: EMPTY }] });
  const model = mk();
  const ops = loadPane(model);
  const before = model.paragraphs.map((p) => p.text);
  const r = await run(() => ops.dispatchOp("table_delete", { table_index: 1, expected_text: "R1C1\tR1C2\nR2C1\tR2C2" }));
  check("table_delete #79: deletes the matching table", r.ok && r.value.applied === true && r.value.tables_before === 2 && r.value.tables_after === 1 && model.tables.length === 1, JSON.stringify(r));
  check("table_delete #79: the other table and neighbour paragraphs are intact", cellText(model.tables[0].rows[0][0]) === "keep" && JSON.stringify(model.paragraphs.map((p) => p.text)) === JSON.stringify(before), JSON.stringify(model.tables));
  check("table_delete #79: reports the deleted text for the evidence", r.ok && r.value.deleted_text === "R1C1\tR1C2\nR2C1\tR2C2", JSON.stringify(r));

  const ws = mk();
  const rw = await run(() => loadPane(ws).opTableDelete({ table_index: 1, expected_text: "  R1C1   R1C2\n R2C1\tR2C2 " }));
  check("table_delete #79: expected_text is whitespace-insensitive", rw.ok, rw.message);

  const mm = mk();
  const rm = await run(() => loadPane(mm).opTableDelete({ table_index: 1, expected_text: "something else" }));
  check("table_delete #79: text mismatch refuses with CELL_TEXT_MISMATCH and nothing is written", !rm.ok && rm.code === "CELL_TEXT_MISMATCH" && trackedCalls(mm).length === 0 && mm.tables.length === 2, JSON.stringify(rm));

  const shaped = mk({ ooxml: '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:body><w:tbl><w:tr><w:tc><w:p><w:r><w:drawing><wp:anchor/></w:drawing></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>' });
  const rs = await run(() => loadPane(shaped).opTableDelete({ table_index: 1, expected_text: "R1C1\tR1C2\nR2C1\tR2C2" }));
  check("table_delete #79: a table anchoring a shape is refused (ANCHORED_SHAPES)", !rs.ok && rs.code === "ANCHORED_SHAPES" && trackedCalls(shaped).length === 0, JSON.stringify(rs));

  const rev = mk({ revs: [{ kind: "ins", author: "Colleague", text: "x" }] });
  const rr = await run(() => loadPane(rev).opTableDelete({ table_index: 1, expected_text: "R1C1\tR1C2\nR2C1\tR2C2" }));
  check("table_delete #79: tracked changes inside the table are refused", !rr.ok && rr.code === "TRACKED_CHANGES_PRESENT" && trackedCalls(rev).length === 0, JSON.stringify(rr));

  const gone = mk();
  const rg = await run(() => loadPane(gone).opTableDelete({ table_index: 9, expected_text: "x" }));
  check("table_delete #79: unknown table is table_not_found", !rg.ok && rg.code === "table_not_found", JSON.stringify(rg));

  const nested = makeModel({ tables: [{ rows: [["a"]], nesting: 2, ooxml: EMPTY }] });
  const rn = await run(() => loadPane(nested).opTableDelete({ table_index: 1, expected_text: "a" }));
  check("table_delete #79: nested tables are refused", !rn.ok && /nested table/.test(rn.message) && trackedCalls(nested).length === 0, JSON.stringify(rn));

  const dropped = mk();
  dropped.dropTableDelete = true;
  const rd = await run(() => loadPane(dropped).opTableDelete({ table_index: 1, expected_text: "R1C1\tR1C2\nR2C1\tR2C2" }));
  check("table_delete #79: a delete that did not take is VERIFICATION_FAILED", !rd.ok && rd.code === "VERIFICATION_FAILED", JSON.stringify(rd));

  const collateral = mk();
  collateral.onTableDelete = () => (collateral.paragraphs[0].text = "Changed");
  const rc = await run(() => loadPane(collateral).opTableDelete({ table_index: 1, expected_text: "R1C1\tR1C2\nR2C1\tR2C2" }));
  check("table_delete #79: a changed neighbour paragraph is VERIFICATION_FAILED", !rc.ok && rc.code === "VERIFICATION_FAILED", JSON.stringify(rc));

  const tracked = mk();
  const rt = await run(() => loadPane(tracked).opTableDelete({ table_index: 1, expected_text: "R1C1\tR1C2\nR2C1\tR2C2", track_changes: true }));
  check("table_delete #79: tracked delete leaves the table marked and restores the mode", rt.ok && rt.value.tracked === true && tracked.mode === "Off" && tracked.tables.length === 2, JSON.stringify(rt) + tracked.mode);
}

// 16. table-scoped accept / reject
{
  const colleague = [{ kind: "ins", author: "Colleague", text: "elsewhere" }];
  const mk = (extra = {}) => makeModel({
    ooxmlMode: "revisions",
    paragraphs: ["Intro", { text: "Proof", revs: colleague }, "Tail"],
    tables: [{ rows: [["T1"]], ooxml: EMPTY, revs: [{ kind: "ins", author: "Colleague", text: "in" }, { kind: "del", author: "Colleague", text: "gone" }] }, { rows: [["T2"]], ooxml: EMPTY, revs: [{ kind: "ins", author: "Other", text: "t2" }] }],
    ...extra,
  });
  check("scoped #79: revisions_scoped, table_delete and paragraphs_unlist are reported", ["revisions_scoped", "table_delete", "paragraphs_unlist"].every((c) => loadPane(mk()).PANE_CAPABILITIES.includes(c)), "");

  const model = mk();
  const ops = loadPane(model);
  const r = await run(() => ops.dispatchOp("revisions_accept", { table_index: 1 }));
  check("scoped #79: accept only mutates the named table's revisions", r.ok && r.value.scoped === true && r.value.changed_count === 2 && model.tables[0].revs.length === 0 && model.tables[1].revs.length === 1 && model.paragraphs[1].revs.length === 1, JSON.stringify(r));
  check("scoped #79: acceptAll is called on the table's own collection only", model.calls.filter((c) => /acceptAll/.test(c)).length === 1 && !model.calls.some((c) => /body/.test(c)), model.calls.join(" | "));
  check("scoped #79: the table survives an accept (table_removed false)", r.ok && r.value.table_removed === false && model.tables.length === 2, JSON.stringify(r));

  const rej = mk();
  const rr = await run(() => loadPane(rej).opRevisionsMutate({ table_index: 2 }, "reject"));
  check("scoped #79: reject only mutates the named table's revisions", rr.ok && rr.value.changed_count === 1 && rej.tables[0].revs.length === 2 && rej.paragraphs[1].revs.length === 1, JSON.stringify(rr));

  const nocoll = mk();
  nocoll.ooxmlUnreadable = true;
  const rn = await run(() => loadPane(nocoll).opRevisionsMutate({ table_index: 1 }, "reject"));
  check("scoped #79: a pane that cannot verify refuses before any write", !rn.ok && rn.code === "LIVE_CAPABILITY_MISSING" && trackedCalls(nocoll).length === 0, JSON.stringify(rn));

  const both = mk();
  const rb = await run(() => loadPane(both).opRevisionsMutate({ table_index: 1, revision_ids: ["x"] }, "reject"));
  check("scoped #79: table_index with revision_ids is INVALID_INPUT", !rb.ok && rb.code === "INVALID_INPUT" && trackedCalls(both).length === 0, JSON.stringify(rb));

  const locked = mk();
  const rl = await run(() => loadPane(locked).opRevisionsMutate({ table_index: 1, forbiddenSections: ["intro"] }, "reject"));
  check("scoped #79: refused while another client holds a section lock", !rl.ok && rl.code === "LOCKED_BY_OTHER_CLIENT" && trackedCalls(locked).length === 0, JSON.stringify(rl));

  const missing = mk();
  const rm = await run(() => loadPane(missing).opRevisionsMutate({ table_index: 9 }, "reject"));
  check("scoped #79: unknown table is table_not_found", !rm.ok && rm.code === "table_not_found", JSON.stringify(rm));

  const leaky = mk({});
  leaky.scopedRejectTouchesOutside = true;
  const rk = await run(() => loadPane(leaky).opRevisionsMutate({ table_index: 1 }, "reject"));
  check("scoped #79: a change to revisions OUTSIDE the table is VERIFICATION_FAILED", !rk.ok && rk.code === "VERIFICATION_FAILED", JSON.stringify(rk));

  const none = mk({ tables: [{ rows: [["T1"]], ooxml: EMPTY }] });
  const rz = await run(() => loadPane(none).opRevisionsMutate({ table_index: 1 }, "accept"));
  check("scoped #79: a table with no revisions is a verified no-op", rz.ok && rz.value.changed_count === 0, JSON.stringify(rz));
}

// 17. #79 item 4: tracked insert after a formatted anchor, then scoped reject, leaves the anchor alone
{
  const model = makeModel({
    ooxmlMode: "revisions",
    paragraphs: ["Intro", { text: "Purple italic anchor", watch: true, listed: false, revs: [] }, { text: "Tail", revs: [{ kind: "ins", author: "Colleague", text: "colleague edit" }] }],
  });
  const ops = loadPane(model);
  const anchor = model.paragraphs[1];
  anchor.font = { color: "#800080", italic: true };
  const fontBefore = JSON.stringify(anchor.font);
  const ins = await run(() => ops.opTableInsert({ rows: [[spec("a")]], anchor: { paragraph_text: "Purple italic anchor", position: "after" }, header_rows: 0, track_changes: true }));
  check("anchor #79: tracked insert after the anchor succeeds", ins.ok && ins.value.table_index === 1 && model.tables[0].revs.some((r) => r.tableInsert), JSON.stringify(ins) + model.calls.join(" | "));
  const rej = await run(() => ops.dispatchOp("revisions_reject", { table_index: ins.ok ? ins.value.table_index : 1 }));
  check("anchor #79: scoped reject removes the tracked table", rej.ok && rej.value.table_removed === true && model.tables.length === 0, JSON.stringify(rej));
  check("anchor #79: the colleague's tracked change survives", model.paragraphs[2].revs.length === 1, JSON.stringify(model.paragraphs));
  check("anchor #79: no write ever touches the anchor paragraph", anchorWrites(model).length === 0, model.calls.join(" | "));
  check("anchor #79: the anchor's text and formatting are unchanged", anchor.text === "Purple italic anchor" && JSON.stringify(anchor.font) === fontBefore, JSON.stringify(anchor));
  check("anchor #79: tracking mode is back to what it was", model.mode === "Off", model.mode);
}

// Approved execution must not load body text; target CAS runs after async guards.
{
  const model = makeModel({tables:[{rows:[['old']]}]});
  const ops = loadPane(model); ops.scopedTest();
  const r = await run(() => ops.opCellSet({table_index:1,row_index:1,cell_index:1,
    paragraphs:[para('new')], expected_before_text:'old', executionSections:[{slug:'target'}], executionTargetSection:'target'}));
  check('approved cell: mutation verifies', r.ok, JSON.stringify(r));
  check('approved cell: no full-body text load', !(model.loads || []).some(([label,spec]) => label === 'body' && spec.includes('text')), JSON.stringify(model.loads));
}
{
  const model = makeModel({tables:[{rows:[['old']]}]});
  const ops = loadPane(model); ops.scopedTest(() => { model.tables[0].rows[0][0].paras=['coauthor']; });
  const r = await run(() => ops.opCellSet({table_index:1,row_index:1,cell_index:1,
    paragraphs:[para('new')], expected_before_text:'old', track_changes:true, executionSections:[{slug:'target'}], executionTargetSection:'target'}));
  check('approved cell: refusal preserves tracking mode', model.mode === 'Off', model.mode);
  check('approved cell: edit during comment guard refuses', !r.ok && r.code === 'LIVE_STALE', JSON.stringify(r));
  check('approved cell: coauthor text survives', cellText(model.tables[0].rows[0][0]) === 'coauthor', JSON.stringify(model.calls));
}
console.log(JSON.stringify({ results }));
