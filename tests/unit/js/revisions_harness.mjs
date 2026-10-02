import fs from 'node:fs';
import vm from 'node:vm';
import assert from 'node:assert/strict';
import {webcrypto} from 'node:crypto';
let items = [];
const collection = {get items() {return items;}, load() {}, track() {}, untrack() {}};
const body = {text: 'text', load() {}, getTrackedChanges: () => collection};
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
sandbox.Office.context.requirements.isSetSupported = () => false;
assert.equal((await sandbox.api.opRevisionsList()).revisions,null);
console.log('revision behavior passed');
