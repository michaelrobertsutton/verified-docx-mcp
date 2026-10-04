import fs from 'node:fs';
import vm from 'node:vm';
import assert from 'node:assert/strict';
import {webcrypto} from 'node:crypto';
const store = new Map();
let saveFails = false, supported = true, ignoreSet = false;
const settings = {
  get: key => store.get(key),
  set: (key, value) => {if (!ignoreSet) store.set(key, value);},
  saveAsync: cb => cb(saveFails ? {status: 'failed', error: {message: 'boom'}} : {status: 'succeeded'}),
};
const sandbox = {crypto: webcrypto, TextEncoder, console, document: {getElementById: () => ({addEventListener() {}})},
  Office: {onReady() {}, AsyncResultStatus: {Succeeded: 'succeeded'},
    context: {document: {settings}, requirements: {isSetSupported: () => supported}}},
  Word: {run: () => {}}};
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(process.argv[2], 'utf8') + '\nglobalThis.api={opAutoopenGet,opAutoopenSet};', sandbox);
const KEY = 'Office.AutoShowTaskpaneWithDocument';
assert.deepEqual({e: sandbox.api.opAutoopenGet().enabled, s: sandbox.api.opAutoopenGet().supported}, {e: false, s: true});
const on = await sandbox.api.opAutoopenSet({enabled: true});
assert.equal(on.enabled, true);assert.equal(on.was_enabled, false);assert.equal(store.get(KEY), true);
assert.equal(sandbox.api.opAutoopenGet().enabled, true);
const off = await sandbox.api.opAutoopenSet({enabled: false});
assert.equal(off.enabled, false);assert.equal(off.was_enabled, true);assert.equal(store.get(KEY), false);
// Only booleans: never coerce "false" or 1 into a document change.
await assert.rejects(() => sandbox.api.opAutoopenSet({enabled: 'yes'}), e => e.code === 'INVALID_INPUT');
await assert.rejects(() => sandbox.api.opAutoopenSet({}), e => e.code === 'INVALID_INPUT');
// A host without AddinCommands 1.1 refuses before touching the document.
supported = false;
await assert.rejects(() => sandbox.api.opAutoopenSet({enabled: true}), e => e.code === 'LIVE_CAPABILITY_MISSING');
assert.equal(store.get(KEY), false);
assert.equal(sandbox.api.opAutoopenGet().supported, false);
supported = true;
// A failing save surfaces as an error, and a setting that does not read back never claims success.
saveFails = true;
await assert.rejects(() => sandbox.api.opAutoopenSet({enabled: true}), e => /boom/.test(e.message));
saveFails = false; store.set(KEY, false); ignoreSet = true;
await assert.rejects(() => sandbox.api.opAutoopenSet({enabled: true}), e => e.code === 'VERIFICATION_FAILED');
console.log('autoopen behavior passed');
