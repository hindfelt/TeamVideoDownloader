const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { webcrypto } = require('node:crypto');

const root = path.resolve(__dirname, '..');
const source = fs.readFileSync(path.join(root, 'content.js'), 'utf8');
const workerSource = fs.readFileSync(path.join(root, 'mux-worker.js'), 'utf8');
const manifestUrl = 'https://media.svc.ms/transform/videomanifest?part=index&format=dash';

// Synthetic DOM nodes exercise manifest processing without adding a runtime
// dependency. XML parsing itself remains the browser's native DOMParser.
class Element {
  constructor(tag, attrs = {}, children = [], textContent = '') {
    Object.assign(this, { tag, attrs, children, textContent });
    for (const child of children) child.parent = this;
  }
  getAttribute(name) { return this.attrs[name] ?? null; }
  querySelectorAll(tag) {
    return this.children.flatMap(c => [...(c.tag === tag ? [c] : []), ...c.querySelectorAll(tag)]);
  }
  querySelector(tag) { return this.querySelectorAll(tag)[0] || null; }
  closest(tag) { return this.tag === tag ? this : this.parent?.closest(tag) || null; }
}
function dash({ repeat = '1', duration = '10', media = 'seg-$Number%03d$.m4s', base, key, iv = '00'.repeat(16), timeline = true, period = 'PT20S', timescale = '1' } = {}) {
  const seg = new Element('SegmentTemplate', { initialization: 'init.mp4', media, duration, timescale },
    timeline ? [new Element('SegmentTimeline', {}, [new Element('S', { d: duration, r: repeat })])] : []);
  const protection = key ? [new Element('ContentProtection', { schemeIdUri: 'urn:mpeg:dash:sea:2012' }, [
    new Element('SegmentEncryption', { schemeIdUri: 'urn:mpeg:dash:sea:aes128-cbc:2013' }),
    new Element('CryptoPeriod', { keyUriTemplate: key, IV: iv })
  ])] : [];
  const as = new Element('AdaptationSet', { contentType: 'video' }, [
    new Element('Representation', { id: 'video', bandwidth: '1000' }, [seg]), ...protection
  ]);
  return new Element('MPD', {}, [
    ...(base ? [new Element('BaseURL', {}, [], base)] : []),
    new Element('Period', { duration: period }, [as])
  ]);
}
function harness(options = {}) {
  const handlers = {}, elements = [], requests = [], logs = [];
  const window = { location: { origin: 'https://tenant.sharepoint.com' }, addEventListener(type, cb) { handlers[type] = cb; } };
  window.top = window.self = window;
  const document = { readyState: 'complete', title: 'Recording',
    createElement(tag) {
      const e = { tag, style: {}, appendChild() {}, click() {}, addEventListener(type, cb) { this[type] = cb; } };
      elements.push(e); return e;
    }, body: { appendChild() {}, removeChild() {} }
  };
  const context = vm.createContext({ window, document, location: { href: 'https://tenant.sharepoint.com/' },
    chrome: { runtime: { onMessage: { addListener() {} }, getURL: name => 'chrome-extension://test/' + name } },
    console: { error(...args) { logs.push(args); }, info() {}, warn() {} },
    URL, AbortController, ArrayBuffer, Uint8Array, Blob, TextDecoder, crypto: webcrypto,
    setTimeout: options.setTimeout || setTimeout, clearTimeout, Worker: options.Worker,
    DOMParser: class { parseFromString() { return options.dom || dash(); } },
    fetch: async (url, init) => { requests.push({ url, init }); return options.fetch ? options.fetch(url, init) : new Response('', { status: 418 }); }
  });
  // Expose closure internals only in this in-memory test copy.
  vm.runInContext(source.replace(/\}\)\(\);\s*$/, `
    globalThis.check = { MAX_DOWNLOAD_BYTES, mediaUrl, mediaFetchInit, readLimitedResponse, parseDashManifest,
      downloadDashSegments, muxTracks, getCapture: () => ({ manifestUrl: videoManifestUrl, spopactoken: videoSpopActoken }) };
  })();`), context);
  return { ...context.check, requests, logs, elements, context,
    capture(url, token, overrides = {}) {
      handlers.message({ source: window, origin: window.location.origin,
        data: { type: 'VIDEO_MANIFEST_URL', manifestUrl: url, spopactoken: token }, ...overrides });
    },
    click(trusted = true) { return elements.find(e => e.tag === 'button').click({ isTrusted: trusted }); }
  };
}

test('destinations allow Microsoft media and proxy URLs, rejecting lookalikes and unsafe URLs', () => {
  const h = harness();
  for (const url of [manifestUrl, 'https://tenant.sharepoint.com/file', 'https://media.svc.ms.mcas.ms/key', 'https://tenant.sharepoint.com.mcas.ms/file']) {
    assert.equal(h.mediaUrl(url), url);
  }
  for (const url of ['https://svc.ms.attacker.invalid/key', 'https://evilsharepoint.com/file', 'https://attacker.invalid/svc.ms',
    'http://media.svc.ms/key', 'data:text/plain,test', 'https://user:pass@media.svc.ms/key', 'https://media.svc.ms:8443/key', {}, '']) {
    assert.throws(() => h.mediaUrl(url));
  }
});

test('capture validation blocks cross-frame messages, malformed tokens, and stale token reuse', () => {
  const h = harness();
  h.capture(manifestUrl, 'TOKEN_A');
  h.capture('https://attacker.invalid/videomanifest');
  assert.equal(h.getCapture().manifestUrl, manifestUrl);
  h.capture(manifestUrl, 'FORGED', { origin: 'https://attacker.invalid' });
  h.capture(manifestUrl, 'FORGED', { source: {} });
  h.capture(manifestUrl, 'bad\r\nheader');
  assert.equal(h.getCapture().spopactoken, 'TOKEN_A');
  h.capture(manifestUrl); // a fallback for this exact recording may retain its token
  assert.equal(h.getCapture().spopactoken, 'TOKEN_A');
  h.capture(manifestUrl + '&recording=other');
  assert.equal(h.getCapture().spopactoken, null);
});

test('auth is restricted to captured origin and authenticated requests reject redirects', () => {
  const h = harness(), capture = { manifestUrl, spopactoken: 'TOKEN_A' };
  const own = h.mediaFetchInit('https://media.svc.ms/key', capture);
  assert.equal(own.headers['x-spopactoken'], 'TOKEN_A');
  assert.equal(own.redirect, 'error');
  const other = h.mediaFetchInit('https://other.svc.ms/key', capture);
  assert.equal(other.headers, undefined);
  assert.equal(other.redirect, 'error');
  assert.throws(() => h.mediaFetchInit('https://attacker.invalid/key', capture));
});

test('scripted button clicks cannot initiate downloads', async () => {
  const h = harness();
  h.capture(manifestUrl, 'TOKEN_A');
  await h.click(false);
  assert.equal(h.requests.length, 0);
});

test('a normal encrypted download retains its capture even when a new message arrives', async () => {
  const key = new Uint8Array(16), iv = new Uint8Array(16);
  const cryptoKey = await webcrypto.subtle.importKey('raw', key, 'AES-CBC', false, ['encrypt']);
  const encrypted = await webcrypto.subtle.encrypt({ name: 'AES-CBC', iv }, cryptoKey, new Uint8Array([1, 2, 3]));
  let h;
  h = harness({ dom: dash({ key: 'key', repeat: '0' }), fetch: async url => {
    if (url === manifestUrl) {
      h.capture('https://other.svc.ms/videomanifest', 'TOKEN_B');
      return new Response('<MPD/>');
    }
    return new Response(url.endsWith('/key') ? key : encrypted);
  } });
  h.capture(manifestUrl, 'TOKEN_A');
  await h.click();
  assert.equal(h.logs.length, 0);
  assert.equal(h.requests[0].init.headers['x-spopactoken'], 'TOKEN_A');
  const keyRequest = h.requests.find(r => r.url.endsWith('/key'));
  assert.equal(keyRequest.url, 'https://media.svc.ms/transform/key');
  assert.equal(keyRequest.init.headers['x-spopactoken'], 'TOKEN_A');
  assert(h.requests.every(r => r.init.redirect === 'error'));
  assert.equal(h.getCapture().spopactoken, 'TOKEN_B');
});

test('valid DASH templates resolve relative media and key URLs', () => {
  const h = harness({ dom: dash({ base: '../recording/', key: 'key' }) });
  const [track] = h.parseDashManifest('<MPD/>', manifestUrl);
  assert.deepEqual(Array.from(track.segments), ['https://media.svc.ms/recording/seg-001.m4s', 'https://media.svc.ms/recording/seg-002.m4s']);
  assert.equal(track.encryption.keyUri, 'https://media.svc.ms/recording/key');
});

test('manifest limits reject unbounded repetitions, huge formatting widths, and unsafe media', () => {
  for (const config of [{ repeat: '1000000000' }, { repeat: '-1' }, { duration: '0' },
    { media: '$Number%01000000000d$.m4s' }, { base: 'http://media.svc.ms/' },
    { media: '//attacker.invalid/segment' }, { key: 'https://svc.ms.attacker.invalid/key' },
    { key: 'key', iv: 'zz'.repeat(16) }, { timeline: false, timescale: '0' },
    { timeline: false, period: 'PT999999999999H' }]) {
    const h = harness({ dom: dash(config) });
    assert.throws(() => h.parseDashManifest('<MPD/>', manifestUrl), undefined, JSON.stringify(config));
  }
  assert.throws(() => harness().parseDashManifest(' '.repeat(4 * 1024 * 1024 + 1), manifestUrl), /too large/);
});

test('response limits cancel oversized streams even without Content-Length', async () => {
  const h = harness();
  let cancelled = false;
  const response = new Response(new ReadableStream({
    start(controller) { controller.enqueue(new Uint8Array(17)); },
    cancel() { cancelled = true; }
  }));
  await assert.rejects(h.readLimitedResponse(response, 16), /size limit/);
  assert(cancelled);
  await assert.rejects(h.readLimitedResponse(new Response(new Uint8Array(17), { headers: { 'Content-Length': '1' } }), 16), /size limit/);
  const budget = { used: 3584 * 1024 * 1024 - 2 };
  await assert.rejects(h.readLimitedResponse(new Response(new Uint8Array(3)), 16, budget), /size limit/);
  assert.equal((await h.readLimitedResponse(new Response(new Uint8Array(16), { headers: { 'Content-Length': '36' } }), 16)).byteLength, 16);
});

test('downloader accepts bytes through the 3.5 GiB boundary and rejects the next byte', async () => {
  const h = harness();
  const limit = 3.5 * 1024 ** 3;
  assert.equal(h.MAX_DOWNLOAD_BYTES, limit);
  const budget = { used: limit - 2 };
  assert.equal((await h.readLimitedResponse(new Response(new Uint8Array(2)), limit, budget)).byteLength, 2);
  assert.equal(budget.used, limit);
  await assert.rejects(h.readLimitedResponse(new Response(new Uint8Array(1)), limit, budget), /size limit/);
  assert.equal(budget.used, limit);
});

const video = fs.readFileSync(path.join(__dirname, 'fixtures/video.mp4'));
const audio = fs.readFileSync(path.join(__dirname, 'fixtures/audio.mp4'));
async function mux(videoBytes = video, audioBytes = audio) {
  let handler, result;
  vm.runInNewContext(workerSource, { Uint8Array, ArrayBuffer,
    self: { addEventListener(_type, cb) { handler = cb; }, postMessage(data) { if (!data.progress) result = data; } }
  });
  await handler({ data: { video: [videoBytes], audio: [audioBytes] } });
  return result;
}

test('muxer accepts real fragmented H.264/AAC fixtures and returns a flat MP4', async () => {
  const output = await mux();
  assert.equal(output.error, undefined);
  const bytes = Buffer.from(output.result);
  assert(bytes.includes(Buffer.from('moov')));
  assert(bytes.includes(Buffer.from('mdat')));
  assert(!bytes.includes(Buffer.from('moof')));
  assert.equal(bytes.indexOf('trak') === bytes.lastIndexOf('trak'), false);
});

test('muxer rejects huge sample counts, truncated tables, and invalid box boundaries', async () => {
  const hugeCount = Buffer.from(video);
  hugeCount.writeUInt32BE(1000000000, hugeCount.indexOf('trun') + 8);
  assert.match((await mux(hugeCount)).error, /Too many MP4 samples/);
  const truncated = Buffer.from(video);
  const countOffset = truncated.indexOf('trun') + 8;
  truncated.writeUInt32BE(truncated.readUInt32BE(countOffset) + 1, countOffset);
  assert.match((await mux(truncated)).error, /Truncated MP4 sample table/);
  const oversized = Buffer.from(video);
  oversized.writeUInt32BE(0xffffffff, oversized.indexOf('moov') - 4);
  assert.match((await mux(oversized)).error, /box size/);
  const child = Buffer.from(video);
  child.writeUInt32BE(0xffffffff, child.indexOf('traf') - 4);
  assert.match((await mux(child)).error, /box size/);
  const excessData = Buffer.from(video);
  const trun = excessData.indexOf('trun') - 4;
  const flags = excessData.readUInt32BE(trun + 8) & 0xffffff;
  assert(flags & 0x200); // fixture has explicit sample sizes
  const firstSample = trun + 16 + ((flags & 1) ? 4 : 0) + ((flags & 4) ? 4 : 0);
  excessData.writeUInt32BE(0xffffffff, firstSample + ((flags & 0x100) ? 4 : 0));
  assert.match((await mux(excessData)).error, /samples exceed media data/);
});

test('mux worker rejects invalid chunk types and excessive chunk counts', async () => {
  assert.match((await mux({ byteLength: 1 })).error, /Invalid mux chunk/);
  let handler, result;
  vm.runInNewContext(workerSource, { Uint8Array, ArrayBuffer,
    self: { addEventListener(_type, cb) { handler = cb; }, postMessage(data) { result = data; } }
  });
  await handler({ data: { video: Array(100003).fill(new Uint8Array(0)), audio: [] } });
  assert.match(result.error, /Invalid mux input/);
});

test('mux input limit matches the downloader and rejects oversized input before concatenation', async () => {
  let handler, result;
  const context = vm.createContext({ Uint8Array, ArrayBuffer,
    self: { addEventListener(_type, cb) { handler = cb; }, postMessage(data) { result = data; } }
  });
  vm.runInContext(workerSource, context);
  assert.equal(vm.runInContext('MAX_INPUT_BYTES', context), harness().MAX_DOWNLOAD_BYTES);
  // Repeated references need only one MiB of actual memory; the logical input
  // crosses the limit and must fail before concatChunks allocates anything.
  const chunk = new Uint8Array(1024 * 1024);
  await handler({ data: { video: Array(3585).fill(chunk), audio: [] } });
  assert.match(result.error, /mux size limit/);
});

test('stalled mux workers are terminated on timeout and abort', async () => {
  let terminated = 0, timeout;
  const h = harness({ Worker: class { postMessage() {} terminate() { terminated++; } },
    setTimeout(cb, delay) { assert.equal(delay, 120000); timeout = cb; return 1; },
    fetch: () => new Response('// bundled worker') });
  const controller = new AbortController();
  const pending = h.muxTracks([], [], () => {}, controller.signal);
  while (!timeout) await new Promise(resolve => setImmediate(resolve));
  timeout();
  await assert.rejects(pending, /timed out/);
  assert.equal(terminated, 1);
  const aborted = h.muxTracks([], [], () => {}, controller.signal);
  await new Promise(resolve => setImmediate(resolve));
  controller.abort();
  await assert.rejects(aborted, /Cancelled/);
  assert.equal(terminated, 2);
});
