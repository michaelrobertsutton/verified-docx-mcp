// Issue #113 (anchored_find_v1): replace/format located by a literal prefix and suffix plus the
// sha256 of the span's normalized text, paired inside one paragraph. Prints one JSON line.
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {fixture} from './edit_harness.mjs';

const path = process.argv[2];
const LONG = 'Skyward will modernize the claims platform ' + 'with measured, incremental releases '.repeat(15) + 'by 2027.';
const OTHER = 'Skyward will modernize the claims platform, but this paragraph has no suffix.';
const make = (text) => {
  const f = fixture(path, text, false, true);
  // The fixture has no tracked-change model; the revision guard is covered by revisions_harness.
  vm.runInContext('guardRevisions = async () => {};', f.sandbox);
  return f;
};
const hash = async (f, text) => vm.runInContext(`spanHash(${JSON.stringify(text)})`, f.sandbox);
const anchored = (prefix, suffix, sha) => ({prefix, suffix, expected_text_sha256: sha, expected_matches: 1});
const out = {};

// Parity with locate.normalize_span_text: quotes, NBSP-class spaces, whitespace runs, soft hyphen.
const probe = make('x');
out.parityText = 'A\u201cquoted\u201d \u2018x\u2019\u00a0\u00a0and\u0009\u0009tabs\u000b\u2003em\u00adsoft \u00abg\u00bb\u2032 \u001c end ';
out.parityNormalized = await vm.runInContext(`normalizeSpanText(${JSON.stringify(out.parityText)})`, probe.sandbox);
out.parityHash = await hash(probe, out.parityText);

// A ~600-character paragraph whose prefix also opens another paragraph (which has no suffix).
const text = `Heading\n${OTHER}\n${LONG}\nTail paragraph.`;
const f = make(text);
const sha = await hash(f, LONG);
const result = await f.sandbox.api.opReplace({...anchored(LONG.slice(0, 60), LONG.slice(-40), sha), replace: 'Replaced.'});
assert.equal(result.match_count, 1);
assert.equal(result.matches[0].before, LONG);
assert.equal(f.state.text, `Heading\n${OTHER}\nReplaced.\nTail paragraph.`);
out.longReplaced = true;

// Two identical paragraphs: refused, nothing written.
const twice = make(`${LONG}\n${LONG}`);
await assert.rejects(() => twice.sandbox.api.opReplace({...anchored(LONG.slice(0, 60), LONG.slice(-40), sha), replace: 'x'}),
  e => e.code === 'AMBIGUOUS_SPAN');
assert.equal(twice.state.text, `${LONG}\n${LONG}`);
out.ambiguousRefused = true;

// A wrong hash, and a suffix that only exists in the next paragraph: zero_match, nothing written.
const wrong = make(text);
await assert.rejects(() => wrong.sandbox.api.opReplace({...anchored(LONG.slice(0, 60), LONG.slice(-40), 'a'.repeat(64)), replace: 'x'}),
  e => e.code === 'zero_match');
const across = make(`${LONG.slice(0, 300)}\n${LONG.slice(300)}`);
await assert.rejects(() => across.sandbox.api.opReplace({...anchored(LONG.slice(0, 60), LONG.slice(-40), sha), replace: 'x'}),
  e => e.code === 'zero_match');
assert.equal(wrong.state.text, text);
out.zeroMatchRefused = true;

// Malformed: find together with prefix, an over-long anchor, a bad hash.
const bad = make(text);
for (const payload of [
  {...anchored(LONG.slice(0, 60), LONG.slice(-40), sha), find: LONG.slice(0, 60), replace: 'x'},
  {...anchored(LONG.slice(0, 231), LONG.slice(-40), sha), replace: 'x'},
  {...anchored(LONG.slice(0, 60), LONG.slice(-40), 'nothex'), replace: 'x'},
]) await assert.rejects(() => bad.sandbox.api.opReplace(payload), e => e.code === 'INVALID_INPUT');
assert.equal(bad.state.text, text);
out.invalidRefused = true;

// The hash is re-checked at the mutation boundary: a same-length edit after the search is refused.
const raced = make(text);
vm.runInContext('const realGuard = guardComments; guardComments = async (...a) => { mutate(); return realGuard(...a); };',
  Object.assign(raced.sandbox, {mutate: () => { raced.state.text = raced.state.text.replace('incremental', 'INCREMENTAL'); }}));
await assert.rejects(() => raced.sandbox.api.opReplace({...anchored(LONG.slice(0, 60), LONG.slice(-40), sha), replace: 'x'}),
  e => e.code === 'LIVE_STALE');
assert.ok(raced.state.text.includes('INCREMENTAL'));
out.racedRefused = true;

// requireAnchoredText alone: payload.find is rewritten to the moved text, so the plain text
// compare passes and only the hash re-check at the mutation boundary can refuse.
const moveAfterSearch = (f) => vm.runInContext(
  'const realSearch = searchScoped; searchScoped = async (ctx, p) => { const hits = await realSearch(ctx, p); ' +
  'mutate(); if (p.find !== undefined) p.find = hits[0].text; return hits; };',
  Object.assign(f.sandbox, {mutate: () => { f.state.text = f.state.text.replace('incremental', 'INCREMENTAL'); }}));
const hashOnly = make(text);
moveAfterSearch(hashOnly);
await assert.rejects(() => hashOnly.sandbox.api.opReplace({...anchored(LONG.slice(0, 60), LONG.slice(-40), sha), replace: 'x'}),
  e => e.code === 'LIVE_STALE' && /Anchored target changed/.test(e.message));
assert.ok(hashOnly.state.text.includes('INCREMENTAL') && !hashOnly.state.text.includes('\nx\n'));
out.hashOnlyRaceRefused = true;

// opFormat has no text compare: its hash re-check refuses a moved span, and nothing is formatted.
const fmtRace = make(text);
moveAfterSearch(fmtRace);
await assert.rejects(() => fmtRace.sandbox.api.opFormat({...anchored(LONG.slice(0, 60), LONG.slice(-40), sha), bold: true}),
  e => e.code === 'LIVE_STALE');
const fmtStart = fmtRace.state.text.indexOf('Skyward will modernize the claims platform with');
assert.ok(fmtRace.state.chars.slice(fmtStart, fmtStart + LONG.length).every(c => !c.bold));
out.formatRaceRefused = true;

// format_text over the same anchored span.
const fmt = make(text);
const formatted = await fmt.sandbox.api.opFormat({...anchored(LONG.slice(0, 60), LONG.slice(-40), sha), bold: true});
assert.equal(formatted.match_count, 1);
assert.equal(formatted.matches[0].boldAfter, true);
const start = fmt.state.text.indexOf(LONG);
assert.ok(fmt.state.chars.slice(start, start + LONG.length).every(c => c.bold));
out.formatApplied = true;

console.log(JSON.stringify(out));
