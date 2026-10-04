import fs from 'node:fs';
import vm from 'node:vm';
import assert from 'node:assert/strict';
import {webcrypto} from 'node:crypto';
let items = [];
let hidden = 0; // revisions Word's API merged into neighbours but the markup still holds
const collection = {get items() {return items;}, load() {}, track() {}, untrack() {},
  acceptAll() {items = []; hidden = 0;}, rejectAll() {items = []; hidden = 0;}};
const markupOf = () => `<pkg:part pkg:name="/word/document.xml">${'<w:ins w:id="1"/>'.repeat(items.length + hidden)}<w:delText>x</w:delText></pkg:part>`;
const body = {text: 'text', load() {}, getTrackedChanges: () => collection, getOoxml: () => ({value: markupOf()})};
const context = {document: {body}, sync: async () => {}};
function revision(text) {
  return {author: 'A', date: 'd', text, type: 'Added', load() {}, track() {}, untrack() {},
    getRange: () => ({paragraphs: {items: [{text: 'paragraph'}], load() {}}}),
    accept() {items = items.filter(x => x !== this);}, reject() {this.accept();}};
}
const sandbox = {crypto: webcrypto, TextEncoder, console, document: {getElementById: () => ({addEventListener() {}})},
  Office: {onReady() {}, context: {requirements: {isSetSupported: () => true}}},
  Word: {run: (objects, fn) => (fn || objects)(context)}};
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(process.argv[2], 'utf8') + '\nglobalThis.api={opRevisionsList,opRevisionsMutate,guardRevisions};',sandbox);
items = [revision('duplicate'),revision('duplicate')];
const listing = await sandbox.api.opRevisionsList();
assert.equal(listing.revisions.length, 2);
assert.notEqual(listing.revisions[0].revision_id,listing.revisions[1].revision_id);
await assert.rejects(() => sandbox.api.guardRevisions(context,[{getTrackedChanges: () => collection}]),e => e.code === 'TRACKED_CHANGES_PRESENT');
const result = await sandbox.api.opRevisionsMutate({revision_ids:[listing.revisions[0].revision_id]},'accept');
assert.equal(result.before_count,2);assert.equal(result.after_count,1);
await assert.rejects(() => sandbox.api.opRevisionsMutate({revision_ids:[listing.revisions[0].revision_id]},'reject'));
const all = await sandbox.api.opRevisionsMutate({},'reject');assert.equal(all.after_count,0);
// Coverage cross-check: the API list omits revisions the markup still holds.
items = [revision('a')]; hidden = 2;
const partial = await sandbox.api.opRevisionsList();
assert.equal(partial.revisions.length, 1);assert.equal(partial.ooxml_revision_count, 3);
// Accept-all goes through the collection so merged revisions are covered too.
const everything = await sandbox.api.opRevisionsMutate({},'accept');
assert.equal(everything.ooxml_after, 0);assert.equal(items.length, 0);
// An unreadable body OOXML must not fail the listing (it once aborted live_status too):
// the Office.js inventory is kept, revision_ooxml is omitted, the markup count is unknown.
items = [revision('a')]; hidden = 0;
const realOoxml = body.getOoxml;
body.getOoxml = () => {throw new Error('getOoxml failed');};
const unreadable = await sandbox.api.opRevisionsList();
assert.equal(unreadable.revisions.length, 1);assert.equal(unreadable.ooxml_revision_count, null);
assert.equal('revision_ooxml' in unreadable, false);
body.getOoxml = realOoxml;
// A readable body ships its OOXML, read once, alongside the inventory.
let reads = 0;
body.getOoxml = () => {reads++; return realOoxml();};
const readable = await sandbox.api.opRevisionsList();
assert.match(readable.revision_ooxml, /<w:ins /);assert.equal(readable.ooxml_revision_count, 1);assert.equal(reads, 1);
body.getOoxml = realOoxml; items = [];
// An accept that leaves markup behind must not claim success.
items = [revision('a')]; hidden = 1;
collection.acceptAll = () => {};
await assert.rejects(() => sandbox.api.opRevisionsMutate({},'accept'), e => e.code === 'VERIFICATION_FAILED');
// Guard: a match inside a larger revision is refused; an apart one is allowed.
const guardRange = loc => ({getTrackedChanges: () => ({items: [], load() {}}),
  paragraphs: {getFirst: () => ({getRange: () => ({expandTo: () => ({getTrackedChanges: () => ({items: [{getRange: () => 'r'}], load() {}})})})}),
               getLast: () => ({getRange: () => 'end'})},
  compareLocationWith: () => ({value: loc})});
await assert.rejects(() => sandbox.api.guardRevisions(context,[guardRange('Inside')]), e => e.code === 'TRACKED_CHANGES_PRESENT');
await assert.rejects(() => sandbox.api.guardRevisions(context,[guardRange('OverlapsEnd')]), e => e.code === 'TRACKED_CHANGES_PRESENT');
await sandbox.api.guardRevisions(context,[guardRange('AdjacentBefore')]);
await sandbox.api.guardRevisions(context,[guardRange('After')]);
sandbox.Office.context.requirements.isSetSupported = () => false;
assert.equal((await sandbox.api.opRevisionsList()).revisions,null);
console.log('revision behavior passed');
