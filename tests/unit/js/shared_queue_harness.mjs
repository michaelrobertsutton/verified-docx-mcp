import fs from 'node:fs';
import vm from 'node:vm';
import assert from 'node:assert/strict';
const sandbox = {
  console, Date, JSON, Map, Set, Promise, TextEncoder, crypto: globalThis.crypto,
  document: {getElementById: () => ({addEventListener() {}, style: {}})},
  window: {}, Office: {onReady() {}}, Word: {run: fn => fn({})},
};
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(process.argv[2], 'utf8') + `
 globalThis.queue = enqueuePaneOperation;
 globalThis.dispatch = dispatchOp;
 globalThis.mutations = 0;
 requireFreshBody = async (context, hash) => {
   if (hash !== 'current') throw refusalError('stale', 'LIVE_STALE');
 };
 dispatchOpUnchecked = async () => { globalThis.mutations++; return {}; };
 opShapesList = async () => ({shapes: []});
`, sandbox);
let release;
const gate = new Promise(resolve => { release = resolve; });
const events = [];
const first = sandbox.queue(async () => { events.push('start'); await gate; events.push('end'); });
const second = sandbox.queue(async () => {events.push('second'); throw Error('expected');});
const handled = second.catch(() => {});
const third = sandbox.queue(async () => events.push('third'));
await new Promise(resolve => setTimeout(resolve, 5));
assert.deepEqual(events, ['start']);
release();
await Promise.all([first, handled, third]);
assert.deepEqual(events, ['start', 'end', 'second', 'third']);
for (const op of ['replace', 'format', 'comment_reply', 'comment_resolve', 'autoopen_set', 'save']) {
  await assert.rejects(sandbox.dispatch(op, {clientId:'a', expectedBodySha256:'old'}), {code:'LIVE_STALE'});
}
assert.equal(sandbox.mutations, 0);
await sandbox.dispatch('comment_reply', {clientId:'a', expectedBodySha256:'current'});
assert.equal(sandbox.mutations, 1);
console.log('queue order, rejection recovery, and all-family stale preflight passed');
