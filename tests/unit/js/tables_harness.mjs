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
    paragraphs: (opts.paragraphs || ["Intro", "Proof", "Tail"]).map((text) => ({ text })),
    tables: [],
    order: [],
    mode: opts.mode || "Off",
    styles: new Set(opts.styles || ["Table Grid"]),
    supports15: opts.supports15 !== false,
    calls: [], // every queued call/assignment, in queue order
    failOnInsert: !!opts.failOnInsert,
    onSyncCount: null,
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
    ooxml: spec.ooxml ?? '<pkg:package><pkg:part pkg:name="/word/document.xml"><w:tbl/></pkg:part></pkg:package>',
  };
  if (spec.hAlign) t.rows.forEach((r) => r.forEach((c) => (c.hAlign = spec.hAlign)));
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

const topLevelTables = (model) => model.order.filter((i) => i.kind === "t").map((i) => i.ref);

// ------------------------------------------------------------- mock Office.js

function makeContext(model) {
  const ctx = { queue: [], syncs: 0 };
  const log = (s) => model.calls.push(s);

  function node(label, readable, impl, setters = {}) {
    const loaded = new Set();
    impl.load = (spec) => {
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

  function range(table) {
    return {
      compareLocationWith: (other) => clientResult(() => (other.__table === table ? "Equal" : "Before")),
      getOoxml: () => clientResult(() => table.ooxml),
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
    const makePara = (idxRef) => {
      const pImpl = {
        insertText: (text, where) => {
          log(`paragraph.insertText(${JSON.stringify(text)},${where})`);
          const r = { font: { set bold(v) {}, set italic(v) {} }, set hyperlink(v) {} };
          ctx.queue.push(() => (cell.paras[idxRef.i] += text));
          return r;
        },
        insertBreak: (type, where) => {
          log(`paragraph.insertBreak(${type},${where})`);
          ctx.queue.push(() => (cell.paras[idxRef.i] += "\n"));
        },
      };
      return pImpl;
    };
    const impl = {
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
      paragraphs: {
        getFirst: () => makePara({ i: 0 }),
      },
      insertParagraph: (text, where) => {
        log(`cell.body.insertParagraph(${JSON.stringify(text)},${where})`);
        const ref = { i: null };
        ctx.queue.push(() => {
          cell.paras.push(text);
          ref.i = cell.paras.length - 1;
        });
        return makePara(ref);
      },
    };
    impl.getOoxml = () => clientResult(() => EMPTY);
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
      insertTable: (r, c, where, values) => insertTable(`table.insertTable`, where, r, c, () => model.order.findIndex((i) => i.ref === t) + 1),
      delete: () => log("table.delete()"),
      __table: t,
    };
    return node("table", ["style", "headerRowCount", "nestingLevel"], impl, {
      style: (v) => (t.style = v),
      styleBuiltIn: (v) => (t.style = v),
      headerRowCount: (v) => (t.headerRowCount = v),
    });
  }

  function insertTable(label, where, rowCount, colCount, indexFn) {
    log(`${label}(${rowCount},${colCount},${where})`);
    const t = { style: "Table Grid", headerRowCount: 0, nesting: 1, rows: Array.from({ length: rowCount }, () => Array.from({ length: colCount }, () => newCell(""))), ooxml: "" };
    ctx.queue.push(() => {
      if (model.failOnInsert) throw new Error("simulated Word failure while inserting");
      model.tables.push(t);
      model.order.splice(indexFn(), 0, { kind: "t", ref: t });
    });
    return makeTable(t);
  }

  function makeParagraph(p, level = 0) {
    const impl = {
      get text() { return p.text; },
      get tableNestingLevel() { return level; },
      insertTable: (r, c, where) =>
        insertTable("paragraph.insertTable", where, r, c, () => {
          const at = model.order.findIndex((i) => i.ref === p);
          return where === "Before" ? at : at + 1;
        }),
    };
    return node("paragraph", ["text", "tableNestingLevel"], impl);
  }

  const body = node(
    "body",
    ["text"],
    {
      getOoxml: () => clientResult(() => EMPTY),
      get text() { return bodyText(model); },
      // Like Word, body.paragraphs includes the paragraphs INSIDE table cells
      // (tableNestingLevel 1), in document order.
      paragraphs: collection("body.paragraphs", () =>
        model.order.flatMap((item) =>
          item.kind === "p"
            ? [makeParagraph(item.ref, 0)]
            : item.ref.rows.flatMap((r) => r.flatMap((c) => c.paras.map((text) => makeParagraph({ text }, item.ref.nesting))))
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
      InsertLocation: { end: "End", replace: "Replace" },
      BreakType: { line: "Line" },
      ChangeTrackingMode: { trackAll: "TrackAll", off: "Off" },
      BuiltInStyleName: { gridTable4_Accent1: "GridTable4_Accent1", tableGrid: "TableGrid" },
    },
  };
  vm.createContext(sandbox);
  vm.runInContext(
    src +
      "\n;globalThis.__ops = { opTableInsert, opTableGet, opCellsSet, opCellSet, dispatchOp, PANE_CAPABILITIES };",
    sandbox
  );
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

console.log(JSON.stringify({ results }));
