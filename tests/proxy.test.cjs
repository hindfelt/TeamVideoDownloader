const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '../proxy.js'), 'utf8');

function harness() {
  const listeners = {}, scripts = new Map(), requests = [], reloads = [], errors = [];
  let access = false, grantRequest = false, pending;
  const event = name => ({ addListener(cb) { listeners[name] = cb; } });
  const chrome = {
    runtime: { onInstalled: event('install'), onStartup: event('startup') },
    permissions: {
      contains: async () => access,
      request(options, cb) { requests.push(options); access = grantRequest; pending = cb(grantRequest); },
      onAdded: event('added'), onRemoved: event('removed')
    },
    scripting: {
      getRegisteredContentScripts: async () => [...scripts.values()],
      registerContentScripts: async entries => { for (const entry of entries) scripts.set(entry.id, entry); },
      unregisterContentScripts: async ({ ids }) => { for (const id of ids) scripts.delete(id); }
    },
    action: { onClicked: event('click') },
    tabs: { reload: async id => { reloads.push(id); } }
  };
  vm.runInNewContext(source, { chrome, console: { error(...args) { errors.push(args); } } });
  return { scripts, requests, reloads, errors,
    set access(value) { access = value; },
    set grant(value) { grantRequest = value; },
    async emit(name, ...args) {
      listeners[name](...args);
      await pending;
      await new Promise(resolve => setImmediate(resolve));
    }
  };
}

test('non-proxy pages and lookalike hosts cannot request optional host access', async () => {
  const h = harness();
  for (const url of ['https://tenant.sharepoint.com/', 'https://attacker.invalid/',
    'https://tenant.sharepoint.com.mcas.ms.attacker.invalid/', 'http://tenant.sharepoint.com.mcas.ms/']) {
    await h.emit('click', { id: 1, url });
  }
  assert.equal(h.requests.length, 0);
  assert.equal(h.scripts.size, 0);
});

test('denying proxy permission leaves scripts unregistered and the tab untouched', async () => {
  const h = harness();
  await h.emit('install');
  await h.emit('click', { id: 7, url: 'https://tenant.sharepoint.com.mcas.ms/' });
  assert.equal(h.requests.length, 1);
  assert.equal(h.scripts.size, 0);
  assert.deepEqual(h.reloads, []);
  assert.deepEqual(h.errors, []);
});

test('granting and revoking proxy access registers and removes only bundled scripts', async () => {
  const h = harness();
  h.grant = true;
  await h.emit('click', { id: 7, url: 'https://tenant.sharepoint.com.mcas.ms/' });
  assert.deepEqual(h.reloads, [7]);
  assert.equal(h.scripts.size, 2);
  for (const script of h.scripts.values()) {
    assert.deepEqual(Array.from(script.matches), ['https://*.sharepoint.com.mcas.ms/*']);
    for (const file of script.js) assert(fs.existsSync(path.join(__dirname, '..', file)));
  }
  h.access = false;
  await h.emit('removed');
  assert.equal(h.scripts.size, 0);
  assert.deepEqual(h.errors, []);
});
