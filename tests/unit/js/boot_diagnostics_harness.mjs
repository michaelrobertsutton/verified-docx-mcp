// issue #84: pane boot diagnostics (ring buffer, previous_events, boot_seq,
// storage failures, decoupled connect, minimal failure report). Each load()
// runs addin/taskpane.js fresh in a vm context, sharing only the storages the
// caller passes, the way a reloaded pane shares sessionStorage/localStorage.
import fs from 'node:fs';
import vm from 'node:vm';
import {webcrypto} from 'node:crypto';

const source = fs.readFileSync(process.argv[2], 'utf8');

function memoryStorage() {
  const m = new Map();
  return {
    getItem: k => (m.has(k) ? m.get(k) : null),
    setItem: (k, v) => { m.set(k, String(v)); },
    removeItem: k => { m.delete(k); },
    dump: () => Object.fromEntries(m),
  };
}
function throwingStorage() {
  const boom = () => { throw new Error('storage disabled'); };
  return {getItem: boom, setItem: boom, removeItem: boom};
}

// wordMode: 'ok' | 'reject' | 'hang'
function load({session, local, wordMode = 'ok', throwOnAccess = false}) {
  const listeners = {};
  const docListeners = {};
  const posts = [];
  const sockets = [];
  let readyCb = null;
  const el = () => ({
    textContent: '', className: '', dataset: {}, disabled: false,
    classList: {toggle() {}}, addEventListener() {},
  });
  const sandbox = {
    crypto: webcrypto, TextEncoder, console, setTimeout, clearTimeout, setInterval, clearInterval,
    window: {addEventListener: (type, fn) => { (listeners[type] ||= []).push(fn); }},
    document: {
      getElementById: () => el(),
      addEventListener: (type, fn) => { (docListeners[type] ||= []).push(fn); },
      visibilityState: 'hidden',
    },
    Office: {
      HostType: {Word: 'Word'},
      onReady: cb => { readyCb = cb; },
      context: {host: 'Word', platform: 'Mac', document: {url: 'https://x.example/Row01.docx'},
        requirements: {isSetSupported: () => true}},
    },
    Word: {
      run: () => {
        if (wordMode === 'reject') return Promise.reject(new Error('getComments blew up'));
        if (wordMode === 'hang') return new Promise(() => {});
        return Promise.resolve({
          wp: 'issue-106-wp1', generated_at: new Date().toISOString(), host: 'Word', platform: 'Mac',
          requirementSets: {'1.4': true}, documentUrl: 'https://x.example/Row01.docx',
          bodyTextLength: 3, bodyTextSha256: 'abc', comments: [{id: 'c1', resolved: false}],
        });
      },
    },
    WebSocket: class { constructor(url) { this.url = url; sockets.push(this); } },
    fetch: async (url, init) => {
      posts.push({url, body: JSON.parse(init.body)});
      return {status: 200, json: async () => ({ok: true})};
    },
  };
  if (throwOnAccess) {
    for (const name of ['sessionStorage', 'localStorage']) {
      Object.defineProperty(sandbox, name, {get() { throw new Error('SecurityError'); }, configurable: true});
    }
  } else {
    sandbox.sessionStorage = session;
    sandbox.localStorage = local;
  }
  vm.createContext(sandbox);
  vm.runInContext(
    source + '\nglobalThis.api={recordPaneEvent,setReportTimeout:ms=>{REPORT_TIMEOUT_MS=ms;}};', sandbox);
  return {
    api: sandbox.api, listeners, docListeners, posts, sockets,
    boot: () => readyCb({host: 'Word'}),
  };
}
const tick = ms => new Promise(r => setTimeout(r, ms));
const out = {};

// 1. ring buffer is capped at 50 and persisted to sessionStorage.
{
  const session = memoryStorage(), local = memoryStorage();
  const pane = load({session, local});
  for (let i = 0; i < 120; i++) pane.api.recordPaneEvent('test', `e${i}`, 'stack-'.repeat(200));
  const stored = JSON.parse(session.getItem('vdm_pane_events'));
  out.ringLength = stored.length;
  out.ringFirst = stored[0].message;
  out.ringLast = stored[stored.length - 1].message;
  out.ringStackLen = stored[0].stack.length;
  out.ringKeys = Object.keys(stored[0]).sort();
}

// 2. lifecycle + error handlers record; the next load reports and clears them.
{
  const session = memoryStorage(), local = memoryStorage();
  const first = load({session, local});
  first.listeners.error[0]({message: 'kaboom', error: {stack: 'Error: kaboom\n at x'}});
  first.listeners.unhandledrejection[0]({reason: new Error('rejected')});
  first.listeners.pagehide[0]({});
  first.listeners.beforeunload[0]({});
  first.docListeners.visibilitychange[0]({});
  const second = load({session, local});
  out.afterSecondLoadKeys = Object.keys(session.dump());
  out.afterSecondLoadStored = JSON.parse(session.getItem('vdm_pane_events')).map(e => e.kind);
  second.boot();
  await tick(20);
  const report = second.posts[0].body;
  out.previousKinds = report.previous_events.map(e => e.kind);
  out.previousErrorMessage = report.previous_events.find(e => e.kind === 'error').message;
  out.reportHasInstance = typeof report.instance_id === 'string' && report.instance_id.length > 0;
  out.reportBootMsKeys = Object.keys(report.boot_ms).sort();
  out.reportKeepsFields = [report.wp, report.documentUrl, report.comments.length, report.load_error];
  const third = load({session, local});
  third.boot();
  await tick(20);
  out.thirdPreviousKinds = third.posts[0].body.previous_events.map(e => e.kind);
}

// 3. boot_seq increments on every boot (localStorage counter).
{
  const session = memoryStorage(), local = memoryStorage();
  const seqs = [];
  for (let i = 0; i < 3; i++) {
    const pane = load({session, local});
    pane.boot();
    await tick(20);
    seqs.push(pane.posts[0].body.boot_seq);
  }
  out.bootSeqs = seqs;
}

// 4. storages that throw (on use, or on mere access) never break boot.
for (const [label, opts] of [
  ['throwingMethods', {session: throwingStorage(), local: throwingStorage()}],
  ['throwingAccess', {throwOnAccess: true}],
]) {
  const pane = load(opts);
  pane.api.recordPaneEvent('x', 'still fine');
  pane.boot();
  await tick(20);
  const report = pane.posts[0].body;
  out[label] = {
    posted: pane.posts.length, bootSeq: report.boot_seq, previous: report.previous_events,
    comments: report.comments.length, sockets: pane.sockets.length,
  };
}

// 5. connect starts independently of the report build; failures post a minimal report.
for (const mode of ['reject', 'hang']) {
  const session = memoryStorage(), local = memoryStorage();
  const pane = load({session, local, wordMode: mode});
  pane.api.setReportTimeout(40);
  pane.boot();
  out[`${mode}SocketsImmediately`] = pane.sockets.length; // before any await
  out[`${mode}PostsImmediately`] = pane.posts.length;
  await tick(150);
  const report = pane.posts[0].body;
  out[mode] = {
    posts: pane.posts.length,
    keys: Object.keys(report).sort(),
    loadError: report.load_error,
    documentUrl: report.documentUrl,
    bootSeq: report.boot_seq,
    previousIsArray: Array.isArray(report.previous_events),
    storedKinds: JSON.parse(session.getItem('vdm_pane_events')).map(e => e.kind),
  };
}

console.log(JSON.stringify(out));
