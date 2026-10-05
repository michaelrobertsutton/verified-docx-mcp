import fs from 'node:fs';
import vm from 'node:vm';
import assert from 'node:assert/strict';

const sandbox = {
  console, Date, JSON, Map, Set, Promise, TextEncoder, crypto: globalThis.crypto,
  document: {getElementById: () => ({addEventListener() {}, style: {}})},
  window: {}, Office: {onReady() {}}, Word: {run: fn => fn(globalThis.__ctx)},
};
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(process.argv[2], 'utf8') + `
 globalThis.api = {sectionSlug, loadSectionMap, sectionGuard, refuseDocumentWide, dispatchOp, opSectionsList};
`, sandbox);
const api = sandbox.api;
const plain = v => JSON.parse(JSON.stringify(v));

// A fake body: paragraph i occupies [10*i, 10*i+10). Relation names are Word.LocationRelation's.
function relation(a, b) {
  if (a.start === b.start && a.end === b.end) return 'Equal';
  if (a.end < b.start) return 'Before';
  if (a.end === b.start) return 'AdjacentBefore';
  if (a.start > b.end) return 'After';
  if (a.start === b.end) return 'AdjacentAfter';
  if (a.start >= b.start && a.end <= b.end) return a.start === b.start ? 'InsideStart' : a.end === b.end ? 'InsideEnd' : 'Inside';
  if (a.start <= b.start && a.end >= b.end) return 'Contains';
  return a.start < b.start ? 'OverlapsBefore' : 'OverlapsAfter';
}
const makeRange = (start, end) => ({
  start, end,
  getRange(where) { return makeRange(where === 'Start' ? start : end, where === 'Start' ? start : end); },
  expandTo(other) { return makeRange(Math.min(start, other.start), Math.max(end, other.end)); },
  compareLocationWith(other) { return {value: this.unrelated ? 'Unrelated' : relation(this, other)}; },
});
function fakeContext(spec) {
  const items = spec.map(([text, outlineLevel, tableNestingLevel], i) => ({
    text, outlineLevel: outlineLevel ?? 10, tableNestingLevel: tableNestingLevel ?? 0,
    getRange(where) { return makeRange(10 * i + (where === 'End' ? 10 : 0), 10 * i + (where === 'Start' ? 0 : 10)); },
  }));
  // The fake range model gives a paragraph's Start/End getRange the right collapsed points.
  items.forEach((p, i) => { p.getRange = where => where === 'Start' ? makeRange(10 * i, 10 * i) : where === 'End' ? makeRange(10 * i + 10, 10 * i + 10) : makeRange(10 * i, 10 * i + 10); });
  return {
    document: {body: {getRange() { return {paragraphs: {items, load() {}}}; }, paragraphs: {items, load() {}}, text: spec.map(p => p[0]).join('\n'), load() {}}},
    sync: async () => {},
  };
}
const doc = [
  ['Preamble'],                       // 0 (no section)
  ['Intro', 1], ['intro body'],       // 1-2  intro
  ['Scope', 1], ['scope body'], ['cell', 10, 1], // 3-5 scope (a table cell paragraph)
  ['Scope detail', 2], ['detail body'],  // 6-7
  ['Intro', 1], ['second intro'],     // 8-9 duplicate heading text
];
globalThis.__ctx = fakeContext(doc);
sandbox.__ctx = globalThis.__ctx;

// slug and ordinal follow find_sections; table paragraphs are never headings
assert.equal(api.sectionSlug('  Scope & Detail! '), 'scope-detail');
assert.equal(api.sectionSlug('???'), 'section');
const map = await api.loadSectionMap(globalThis.__ctx);
assert.deepEqual(plain(map.sections.map(s => [s.section_key, s.start, s.end])),
  [['intro-1', 1, 3], ['scope-1', 3, 6], ['scope-detail-1', 6, 8], ['intro-2', 8, 10]]);

const para = i => globalThis.__ctx.document.body.paragraphs.items[i].getRange('Whole');
const expectCode = async (promise, code) => assert.rejects(promise, e => e.code === code, code);

// no scopes -> no guard, nothing checked
const none = await api.sectionGuard(globalThis.__ctx, {});
assert.equal(none.active, false);
await none.check([para(4)]);

// forbidden: a target inside, equal to, or overlapping a section is refused; outside is fine
const forbid = await api.sectionGuard(globalThis.__ctx, {forbiddenSections: ['scope-detail']});
await forbid.check([para(4), para(1)]);
await expectCode(forbid.check([para(7)]), 'LOCKED_BY_OTHER_CLIENT');
await expectCode(forbid.check([makeRange(55, 65)]), 'LOCKED_BY_OTHER_CLIENT');   // overlaps the boundary
await expectCode(forbid.check([makeRange(0, 200)]), 'LOCKED_BY_OTHER_CLIENT');   // contains it
const unrelated = makeRange(0, 5); unrelated.unrelated = true;
await expectCode(forbid.check([unrelated]), 'LOCKED_BY_OTHER_CLIENT');           // fail closed

// a heading whose text is not unique (or missing) cannot be resolved: fail closed
await expectCode(api.sectionGuard(globalThis.__ctx, {forbiddenSections: ['intro']}), 'LOCK_SCOPE_UNRESOLVED');
await expectCode(api.sectionGuard(globalThis.__ctx, {forbiddenSections: ['gone']}), 'LOCK_SCOPE_UNRESOLVED');

// own sections: confined to them, with a per-section freshness hash
const sha = async text => Buffer.from(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text)))
  .toString('hex');
const scopeSha = await sha('Scope detail\ndetail body');
const mine = await api.sectionGuard(globalThis.__ctx, {ownSections: [{slug: 'scope-detail', expectedSha256: scopeSha}]});
await mine.check([para(7)]);
await expectCode(mine.check([para(4)]), 'OUTSIDE_LOCKED_SECTION');
await expectCode(mine.check([para(7), para(1)]), 'OUTSIDE_LOCKED_SECTION');
await expectCode(api.sectionGuard(globalThis.__ctx, {ownSections: [{slug: 'scope-detail', expectedSha256: 'old'}]}), 'stale');
await expectCode(api.sectionGuard(globalThis.__ctx, {ownSections: [{slug: 'scope-detail'}], scope: 'textbox-1'}), 'OUTSIDE_LOCKED_SECTION');

// document-wide ops refuse under any section lock
api.refuseDocumentWide({}, 'Revision changes');
assert.throws(() => api.refuseDocumentWide({forbiddenSections: ['x']}, 'Revision changes'), e => e.code === 'LOCKED_BY_OTHER_CLIENT');
assert.throws(() => api.refuseDocumentWide({ownSections: [{slug: 'x'}]}, 'Revision changes'), e => e.code === 'OUTSIDE_LOCKED_SECTION');

// sections_list reports unique-ness and a hash per section
const listed = await api.opSectionsList();
assert.equal(listed.sections.find(s => s.section_key === 'scope-detail-1').sectionSha256, scopeSha);
assert.deepEqual(plain(listed.sections.filter(s => !s.slug_unique).map(s => s.section_key)), ['intro-1', 'intro-2']);
console.log('section guard: resolution, fail-closed relations, own/forbidden scopes, document-wide refusal passed');
