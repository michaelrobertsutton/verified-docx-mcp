import fs from 'node:fs';
import vm from 'node:vm';
import assert from 'node:assert/strict';
import {webcrypto} from 'node:crypto';
// Mocked-Word checks for the issue #77 paragraph_insert op. A model, not Word:
// real-Word behaviour is covered by docs/acceptance/77-insert-paragraphs.md.
let paras = [], mode = 'Off', dropColor = false, modeAtInsert = [];
const knownStyles = new Set(['Heading 1', 'Heading 2', 'Body Text', 'Emphasis']);
const styleTypes = {Emphasis: 'Character'};

function makePara(text, tableNestingLevel = 0) {
  const p = {text, style: 'Normal', styleBuiltIn: 'Normal', tableNestingLevel, _color: '',
    load() {},
    getRange: () => ({ref: p, compareLocationWith: other => ({value: other.ref === p ? 'Equal' : 'Before'})}),
    insertParagraph(t, loc) {
      const n = makePara(t);
      // Word's insertParagraph inherits the neighbour's formatting.
      n.style = p.style; n.styleBuiltIn = p.styleBuiltIn; n._color = p._color;
      modeAtInsert.push(mode);
      paras.splice(paras.indexOf(p) + (loc === 'Before' ? 0 : 1), 0, n);
      return n;
    },
    delete() { paras.splice(paras.indexOf(p), 1); },
  };
  p.font = {load() {}, get color() { return p._color; }, set color(v) { if (!dropColor) p._color = v; }};
  return p;
}
function reset(texts, opts = {}) {
  paras = texts.map((t, i) => makePara(t, opts.tableIndex === i ? 1 : 0));
  if (opts.anchorStyle) paras.forEach(p => { p.style = opts.anchorStyle; p.styleBuiltIn = 'Heading1'; });
  mode = 'Off'; dropColor = !!opts.dropColor; modeAtInsert = [];
}
const texts = () => paras.map(p => p.text);

const context = {
  document: {
    body: {
      load() {}, get text() { return paras.map(p => p.text).join('\n'); },
      paragraphs: {load() {}, get items() { return paras.slice(); }},
      search: find => ({load() {}, items: paras.filter(p => p.text.includes(find)).map(p => ({paragraphs: {getFirst: () => p}}))}),
    },
    load() {},
    get changeTrackingMode() { return mode; }, set changeTrackingMode(v) { mode = v; },
    getStyles: () => ({getByNameOrNullObject: name => ({load() {}, isNullObject: !knownStyles.has(name), type: styleTypes[name] || 'Paragraph'})}),
  },
  sync: async () => {},
};
const sandbox = {crypto: webcrypto, TextEncoder, console, document: {getElementById: () => ({addEventListener() {}})},
  Office: {onReady() {}, context: {requirements: {isSetSupported: () => true}}},
  Word: {run: fn => fn(context), InsertLocation: {before: 'Before', after: 'After'},
    BuiltInStyleName: {normal: 'Normal', heading2: 'Heading2'}, ChangeTrackingMode: {trackAll: 'TrackAll'}}};
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(process.argv[2], 'utf8') + '\nglobalThis.op=opParagraphInsert;', sandbox);
const op = payload => sandbox.op({expected_matches: 1, position: 'after', ...payload});
const specs = [{text: 'New heading', style: 'Heading 2'}, {text: 'Purple text', color: '7030A0'}, {text: 'Green text', color: '00B050'}];

// Tracked happy path, after the anchor; the anchor's style must not leak (explicit Normal).
reset(['Intro', 'Proof', 'Tail'], {anchorStyle: 'Heading 1'});
let r = await op({anchor: 'Proof', paragraphs: specs, track_changes: true});
assert.deepEqual(texts(), ['Intro', 'Proof', 'New heading', 'Purple text', 'Green text', 'Tail']);
assert.equal(r.inserted[0].style, 'Heading 2');
assert.equal(r.inserted[1].style_builtin, 'Normal');
assert.equal(r.inserted[1].color, '#7030A0');
assert.equal(r.inserted[2].color, '#00B050');
assert.equal(r.anchor_index, 2); assert.equal(r.before_count, 3); assert.equal(r.after_count, 6);
assert.ok(modeAtInsert.length === 3 && modeAtInsert.every(m => m === 'TrackAll'));
assert.equal(mode, 'Off');

// Untracked, before the anchor.
reset(['Intro', 'Proof', 'Tail']);
await op({anchor: 'Proof', paragraphs: specs.slice(0, 2), position: 'before'});
assert.deepEqual(texts(), ['Intro', 'New heading', 'Purple text', 'Proof', 'Tail']);
assert.ok(modeAtInsert.every(m => m === 'Off'));

// Refusals write nothing and leave tracking off.
const refuse = async (setup, payload, code) => {
  reset(...setup);
  const before = texts();
  await assert.rejects(() => op(payload), e => e.code === code);
  assert.deepEqual(texts(), before); assert.equal(mode, 'Off');
};
await refuse([['Intro', 'Proof']], {anchor: 'Nope', paragraphs: specs, track_changes: true}, 'ZERO_MATCH');
await refuse([['a x', 'b x']], {anchor: 'x', paragraphs: specs}, 'MATCH_COUNT_MISMATCH');
await refuse([['Intro', 'Proof']], {anchor: 'Proof', paragraphs: [{text: 't', style: 'Nope'}]}, 'STYLE_NOT_FOUND');
await refuse([['Intro', 'Proof']], {anchor: 'Proof', paragraphs: [{text: 't', style: 'Emphasis'}]}, 'UNSUPPORTED_STYLE_TYPE');
await refuse([['Intro', 'Proof'], {tableIndex: 1}], {anchor: 'Proof', paragraphs: specs}, 'STRUCTURAL_BOUNDARY');
await refuse([['Intro', 'Proof']], {anchor: 'Proof', paragraphs: specs, expectedBodySha256: 'bad'}, 'stale');

// A built-in style name resolves when the document does not list it.
reset(['Intro', 'Proof']);
knownStyles.delete('Heading 2');
r = await op({anchor: 'Proof', paragraphs: [{text: 'Built in', style: 'Heading 2'}]});
assert.equal(r.inserted[0].requested_style.kind, 'builtin'); assert.equal(r.inserted[0].style_builtin, 'Heading2');
knownStyles.add('Heading 2');

// A dropped color fails read-back: untracked is removed again, tracked is left for the Review pane.
reset(['Intro', 'Proof', 'Tail'], {dropColor: true});
await assert.rejects(() => op({anchor: 'Proof', paragraphs: specs}),
  e => e.code === 'VERIFICATION_FAILED' && /were removed/.test(e.message));
assert.deepEqual(texts(), ['Intro', 'Proof', 'Tail']); assert.equal(mode, 'Off');
reset(['Intro', 'Proof', 'Tail'], {dropColor: true});
await assert.rejects(() => op({anchor: 'Proof', paragraphs: specs, track_changes: true}),
  e => e.code === 'VERIFICATION_FAILED' && /Review pane/.test(e.message));
assert.equal(texts().length, 6); assert.equal(mode, 'Off');
console.log('paragraph insertion passed');
