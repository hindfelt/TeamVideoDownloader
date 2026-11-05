// content.js — isolated-world content script.
//
// Receives the manifest URL + x-spopactoken from intercept.js (page world),
// then on click downloads the DASH segments, decrypts the DASH-SEA
// AES-128-CBC "clear-key" encryption in-browser (Web Crypto), and muxes the
// audio+video fMP4 tracks into a flat MP4 via mux-worker.js.
//
// Download/decrypt/mux pipeline adapted from
// brendangooden/ms-teams-sharepoint-downloader (MIT License,
// Copyright (c) 2025 Brendan Gooden).
(function () {
  'use strict';

  let videoManifestUrl = null;
  let videoSpopActoken = null;
  let downloading = false;
  const concurrency = 4; // global in-flight segment budget; SharePoint throttles in the low teens
  const MAX_MANIFEST_BYTES = 4 * 1024 * 1024;
  const MAX_DOWNLOAD_BYTES = 3584 * 1024 * 1024; // 3.5 GiB: under the 32-bit MP4 offset ceiling
  const MAX_SEGMENTS = 100000;
  const TRAILING_SEGMENT_SKIP_STATUSES = [404, 409, 416];
  const isTopFrame = window.top === window.self;

  // --- receive captures from the page-world interceptor ---
  window.addEventListener('message', (event) => {
    if (event.source !== window || event.origin !== window.location.origin || !event.data) return;
    if (event.data.type === 'VIDEO_MANIFEST_URL') {
      // Same-origin messages are still page-controlled. Validate every capture,
      // and never carry a token over to a different recording URL.
      let url;
      try { url = mediaUrl(event.data.manifestUrl); } catch (_) { return; }
      const token = event.data.spopactoken;
      if (token != null && (typeof token !== 'string' || token.length > 16384 || !/^[\x20-\x7e]*$/.test(token))) return;
      videoSpopActoken = token || (videoManifestUrl === url ? videoSpopActoken : null);
      videoManifestUrl = url;
      showButton();
      // If the recording lives in a child frame, this frame renders its own
      // (working) button — tell the top frame to drop its placeholder so the
      // user doesn't see two buttons.
      if (!isTopFrame) {
        try { chrome.runtime.sendMessage({ type: 'recordingFrameReady' }); } catch (_) {}
      }
    }
  });

  // The top frame's placeholder hides itself once a child frame reports it owns
  // the real recording (and only while the top frame hasn't captured its own).
  try {
    chrome.runtime.onMessage.addListener((request) => {
      if (request && request.type === 'hidePlaceholderButton' && isTopFrame && !videoManifestUrl) {
        hideButton();
      }
    });
  } catch (_) { /* chrome.runtime may be unavailable in some sandboxed frames */ }

  function mediaFetchInit(url, capture, extra) {
    const target = new URL(mediaUrl(url));
    const init = Object.assign({}, extra || {}, { redirect: 'error' });
    // Bind the token to the exact origin on which it was captured. Host
    // permissions and postMessage origin checks do not validate destinations.
    if (capture.spopactoken && target.origin === new URL(capture.manifestUrl).origin) {
      init.headers = Object.assign({}, init.headers || {}, { 'x-spopactoken': capture.spopactoken });
    }
    return init;
  }

  function mediaUrl(value, base) {
    if (typeof value !== 'string' || !value || value.length > 32768) throw new Error('Invalid media URL');
    const u = new URL(value, base);
    const h = u.hostname.toLowerCase();
    if (u.protocol !== 'https:' || u.username || u.password || u.port || u.href.length > 32768 ||
        !['sharepoint.com', 'svc.ms', 'sharepoint.com.mcas.ms', 'svc.ms.mcas.ms'].some(s => h === s || h.endsWith('.' + s))) {
      throw new Error('Untrusted media destination');
    }
    return u.href;
  }

  async function readLimitedResponse(response, maxBytes, budget) {
    // Count decoded bytes as they arrive. Content-Length can be absent or
    // describe compressed transport bytes rather than the decoded body.
    if (!response.body) return new ArrayBuffer(0);
    const reader = response.body.getReader();
    const chunks = [];
    let size = 0;
    try {
      for (;;) {
        const { done, value } = await reader.read();
        if (done) break;
        size += value.byteLength;
        if (size > maxBytes || (budget && value.byteLength > MAX_DOWNLOAD_BYTES - budget.used)) {
          await reader.cancel();
          throw new Error('Recording exceeds the download size limit');
        }
        if (budget) budget.used += value.byteLength;
        chunks.push(value);
      }
    } finally { reader.releaseLock(); }
    const out = new Uint8Array(size);
    let offset = 0;
    for (const chunk of chunks) { out.set(chunk, offset); offset += chunk.byteLength; }
    return out.buffer;
  }

  // ===========================================================================
  // DASH manifest parsing
  // ===========================================================================
  function hexToBytes(hex) {
    if (!/^[a-f0-9]{32}$/i.test(hex)) throw new Error('Invalid AES-128 IV');
    const out = new Uint8Array(hex.length / 2);
    for (let i = 0; i < out.length; i++) out[i] = parseInt(hex.substr(i * 2, 2), 16);
    return out;
  }

  function parseDashManifest(xmlText, manifestUrl) {
    if (xmlText.length > MAX_MANIFEST_BYTES) throw new Error('Manifest is too large');
    const doc = new DOMParser().parseFromString(xmlText, 'application/xml');
    if (doc.querySelector('parsererror')) throw new Error('Failed to parse DASH manifest XML');

    const baseUrlEl = doc.querySelector('BaseURL');
    const baseUrl = mediaUrl((baseUrlEl && baseUrlEl.textContent.trim()) || '.', manifestUrl);
    let urlCharacters = 0;

    function toAbsolute(url) {
      if (!url) return '';
      const absolute = mediaUrl(url, baseUrl);
      urlCharacters += absolute.length;
      if (urlCharacters > 64 * 1024 * 1024) throw new Error('Manifest expands to too many URLs');
      return absolute;
    }
    function expandTemplate(tpl, repId, bandwidth, number, time) {
      if (tpl.length > 32768 || repId.length > 1024 || bandwidth.length > 1024) throw new Error('Invalid segment template');
      return tpl
        .replace(/\$RepresentationID\$/g, () => repId)
        .replace(/\$Bandwidth\$/g, () => bandwidth)
        .replace(/\$Number%0(\d+)d\$/g, (_, w) => {
          const width = Number(w);
          if (!Number.isSafeInteger(width) || width > 32) throw new Error('Invalid segment number format');
          return String(number).padStart(width, '0');
        })
        .replace(/\$Number\$/g, String(number))
        .replace(/\$Time\$/g, String(time));
    }

    const adaptationSets = Array.from(doc.querySelectorAll('AdaptationSet'));
    if (adaptationSets.length > 16) throw new Error('Too many manifest tracks');
    const isMuxed = adaptationSets.length === 1;
    const tracks = [];
    let segmentCount = 0;
    function integer(value, min = 0) {
      const n = Number(value);
      if (!/^\d+$/.test(value) || !Number.isSafeInteger(n) || n < min) throw new Error('Invalid manifest segment count or timing');
      return n;
    }
    function reserveSegments(count) {
      if (!Number.isSafeInteger(count) || count < 0 || count > MAX_SEGMENTS - segmentCount) throw new Error('Too many recording segments');
      segmentCount += count;
    }

    for (const as of adaptationSets) {
      let type = as.getAttribute('contentType') || '';
      if (!type) {
        const mime = as.getAttribute('mimeType') || '';
        type = mime.startsWith('video') ? 'video' : mime.startsWith('audio') ? 'audio' : '';
      }
      if (isMuxed) type = 'muxed';

      const reps = Array.from(as.querySelectorAll('Representation'))
        .sort((a, b) => parseInt(b.getAttribute('bandwidth') || '0', 10) - parseInt(a.getAttribute('bandwidth') || '0', 10));
      const rep = reps[0];
      if (!rep) continue;

      const repId = rep.getAttribute('id') || '';
      const bandwidth = rep.getAttribute('bandwidth') || '';
      const mimeType = rep.getAttribute('mimeType') || as.getAttribute('mimeType') || '';
      const segTpl = rep.querySelector('SegmentTemplate') || as.querySelector('SegmentTemplate');
      if (!segTpl) continue;

      const startNumber = integer(segTpl.getAttribute('startNumber') || '1');
      const initUrl = toAbsolute(expandTemplate(segTpl.getAttribute('initialization') || '', repId, bandwidth, startNumber, 0));
      const mediaTpl = segTpl.getAttribute('media') || '';
      const segments = [];

      const timeline = segTpl.querySelector('SegmentTimeline');
      if (timeline) {
        let t = 0, segNum = startNumber;
        for (const s of timeline.querySelectorAll('S')) {
          const sT = s.getAttribute('t');
          if (sT !== null) t = integer(sT);
          const d = integer(s.getAttribute('d') || '0', 1);
          const r = integer(s.getAttribute('r') || '0');
          reserveSegments(r + 1);
          if (!Number.isSafeInteger(t + d * (r + 1)) || !Number.isSafeInteger(segNum + r + 1)) throw new Error('Invalid manifest timing');
          for (let i = 0; i <= r; i++) {
            segments.push(toAbsolute(expandTemplate(mediaTpl, repId, bandwidth, segNum, t)));
            t += d; segNum++;
          }
        }
      } else {
        const duration = integer(segTpl.getAttribute('duration') || '0');
        const timescale = integer(segTpl.getAttribute('timescale') || '1', 1);
        const period = as.closest('Period');
        const periodDur = period ? (() => {
          const m = (period.getAttribute('duration') || '').match(/PT(?:(\d+)H)?(?:(\d+)M)?(?:([\d.]+)S)?/);
          return m ? parseInt(m[1] || '0') * 3600 + parseInt(m[2] || '0') * 60 + parseFloat(m[3] || '0') : 0;
        })() : 0;
        if (duration > 0 && periodDur > 0) {
          const count = Math.ceil(periodDur / (duration / timescale));
          reserveSegments(count);
          if (!Number.isSafeInteger(startNumber + count) || !Number.isSafeInteger(count * duration)) throw new Error('Invalid manifest timing');
          for (let i = 0; i < count; i++) {
            segments.push(toAbsolute(expandTemplate(mediaTpl, repId, bandwidth, startNumber + i, i * duration)));
          }
        }
      }

      // DASH-SEA AES-128-CBC encryption (HTTP-fetchable key, not hard DRM).
      let encryption = null;
      const seaCp = [...as.querySelectorAll('ContentProtection')].find(cp =>
        cp.getAttribute('schemeIdUri') === 'urn:mpeg:dash:sea:2012');
      if (seaCp) {
        const segEnc = seaCp.querySelector('SegmentEncryption');
        const scheme = segEnc ? segEnc.getAttribute('schemeIdUri') : '';
        const period = seaCp.querySelector('CryptoPeriod');
        const keyUri = period ? period.getAttribute('keyUriTemplate') : null;
        const ivAttr = period ? (period.getAttribute('IV') || '') : '';
        if (/aes128-cbc/i.test(scheme) && keyUri && ivAttr) {
          encryption = { scheme: 'aes-128-cbc', keyUri: toAbsolute(keyUri), iv: hexToBytes(ivAttr.replace(/^0x/i, '')) };
        }
      }

      tracks.push({ type, mimeType, initUrl, segments, encryption });
    }
    return tracks;
  }

  // ===========================================================================
  // Segment download + decrypt
  // ===========================================================================
  function abortableSleep(ms, signal) {
    return new Promise((resolve, reject) => {
      if (signal && signal.aborted) return reject(Object.assign(new Error('Cancelled'), { name: 'AbortError' }));
      const t = setTimeout(() => { if (signal) signal.removeEventListener('abort', onAbort); resolve(); }, ms);
      function onAbort() { clearTimeout(t); reject(Object.assign(new Error('Cancelled'), { name: 'AbortError' })); }
      if (signal) signal.addEventListener('abort', onAbort, { once: true });
    });
  }

  async function fetchWithRetry(url, init, signal, onThrottle, maxAttempts = 6) {
    url = mediaUrl(url);
    init = Object.assign({}, init, { redirect: 'error' });
    let attempt = 0;
    for (;;) {
      attempt++;
      if (signal && signal.aborted) throw Object.assign(new Error('Cancelled'), { name: 'AbortError' });
      let resp;
      try {
        resp = await fetch(url, init);
      } catch (e) {
        if (e.name === 'AbortError') throw e;
        if (attempt >= maxAttempts) throw e;
        const delayMs = Math.min(1000 * Math.pow(2, attempt - 1), 30000);
        if (onThrottle) onThrottle({ attempt, delayMs, status: 0 });
        await abortableSleep(delayMs, signal);
        continue;
      }
      if ((resp.status === 429 || resp.status === 503) && attempt < maxAttempts) {
        const headerSecs = parseInt(resp.headers.get('Retry-After'), 10);
        const delayMs = Number.isFinite(headerSecs) && headerSecs > 0
          ? Math.min(headerSecs * 1000, 30000)
          : Math.min(1000 * Math.pow(2, attempt - 1), 30000);
        if (onThrottle) onThrottle({ attempt, delayMs, status: resp.status });
        await abortableSleep(delayMs, signal);
        continue;
      }
      return resp;
    }
  }

  async function downloadDashSegments(tracks, capture, onProgress, signal) {
    const budget = { used: 0 };
    const totalSegs = tracks.reduce((s, t) => s + (t.initUrl ? 1 : 0) + t.segments.length, 0);
    let done = 0;
    function reportProgress(text) { onProgress(done, totalSegs, text); }
    function noteThrottle({ attempt, delayMs, status }) {
      reportProgress(`HTTP ${status || 'network'} — backing off ${Math.round(delayMs / 1000)}s (attempt ${attempt})...`);
    }

    const trackStates = await Promise.all(tracks.map(async (track) => {
      const label = tracks.length > 1 ? ` (${track.type} track)` : '';
      let cryptoKey = null;
      if (track.encryption) {
        reportProgress(`Fetching encryption key${label}...`);
        const init = mediaFetchInit(track.encryption.keyUri, capture, { signal });
        const keyResp = await fetchWithRetry(track.encryption.keyUri, init, signal, noteThrottle);
        if (!keyResp.ok) throw new Error(`Encryption key fetch failed: HTTP ${keyResp.status}`);
        const keyBuf = await readLimitedResponse(keyResp, 16);
        if (keyBuf.byteLength !== 16) throw new Error('Invalid AES-128 key size');
        cryptoKey = await crypto.subtle.importKey('raw', keyBuf, { name: 'AES-CBC' }, false, ['decrypt']);
      }
      async function decryptIfNeeded(buf) {
        if (!cryptoKey) return buf;
        return await crypto.subtle.decrypt({ name: 'AES-CBC', iv: track.encryption.iv }, cryptoKey, buf);
      }

      const orderedBufs = new Array((track.initUrl ? 1 : 0) + track.segments.length);
      let segStart = 0;
      if (track.initUrl) {
        reportProgress(`Fetching init segment${label}...`);
        const r = await fetchWithRetry(track.initUrl, { signal }, signal, noteThrottle);
        if (!r.ok) throw new Error(`Init segment failed: HTTP ${r.status}`);
        orderedBufs[0] = await decryptIfNeeded(await readLimitedResponse(r, MAX_DOWNLOAD_BYTES, budget));
        done++; segStart = 1;
      }
      return { track, orderedBufs, segStart, decryptIfNeeded };
    }));

    const queue = [];
    for (const st of trackStates) for (let si = 0; si < st.track.segments.length; si++) queue.push({ st, si });
    reportProgress(`Downloading ${queue.length} segments (${concurrency} parallel)...`);

    await new Promise((resolve, reject) => {
      if (queue.length === 0) return resolve();
      let qIdx = 0, inFlight = 0, rejected = false;
      function launch() {
        while (!rejected && inFlight < concurrency && qIdx < queue.length) {
          if (signal && signal.aborted) { rejected = true; return reject(Object.assign(new Error('Cancelled'), { name: 'AbortError' })); }
          const job = queue[qIdx++];
          inFlight++;
          const isLastSegment = job.si === job.st.track.segments.length - 1;
          fetchWithRetry(job.st.track.segments[job.si], { signal }, signal, noteThrottle)
            .then(r => {
              // SharePoint's transcoder can refuse the final, partial segment
              // (seen as HTTP 409 when under a second of media remains). Drop
              // only that trailing segment; any other failure still aborts.
              if (!r.ok && isLastSegment && TRAILING_SEGMENT_SKIP_STATUSES.includes(r.status)) {
                console.warn(`[Teams DL] Skipping final ${job.st.track.type} segment: HTTP ${r.status}`);
                if (r.body) r.body.cancel().catch(() => {});
                return null;
              }
              if (!r.ok) throw new Error(`Segment failed: HTTP ${r.status}`);
              return readLimitedResponse(r, MAX_DOWNLOAD_BYTES, budget);
            })
            .then(buf => buf === null ? new ArrayBuffer(0) : job.st.decryptIfNeeded(buf))
            .then(buf => {
              if (rejected) return;
              job.st.orderedBufs[job.st.segStart + job.si] = buf;
              done++;
              reportProgress(`Downloading segments... (${done}/${totalSegs})`);
              inFlight--;
              if (inFlight === 0 && qIdx >= queue.length) resolve(); else launch();
            })
            .catch(err => { if (!rejected) { rejected = true; reject(err); } });
        }
      }
      launch();
    });

    return trackStates.map(s => s.orderedBufs);
  }

  // ===========================================================================
  // Mux (Web Worker)
  // ===========================================================================
  let _muxWorkerBlobUrl = null;
  async function getMuxWorkerUrl() {
    if (_muxWorkerBlobUrl) return _muxWorkerBlobUrl;
    const resp = await fetch(chrome.runtime.getURL('mux-worker.js'));
    if (!resp.ok) throw new Error(`mux-worker fetch failed: HTTP ${resp.status}`);
    const src = await resp.text();
    _muxWorkerBlobUrl = URL.createObjectURL(new Blob([src], { type: 'application/javascript' }));
    return _muxWorkerBlobUrl;
  }

  function toArrayBuffers(chunks) {
    return chunks.map(b => {
      if (b instanceof ArrayBuffer) return b;
      if (ArrayBuffer.isView(b)) {
        return b.byteOffset === 0 && b.byteLength === b.buffer.byteLength ? b.buffer : b.slice().buffer;
      }
      return b;
    });
  }

  async function muxTracks(videoChunks, audioChunks, onProgress, signal) {
    const workerUrl = await getMuxWorkerUrl();
    if (signal.aborted) throw Object.assign(new Error('Cancelled'), { name: 'AbortError' });
    return new Promise((resolve, reject) => {
      const worker = new Worker(workerUrl);
      const finish = (error, result) => {
        clearTimeout(timer);
        signal.removeEventListener('abort', onAbort);
        worker.terminate();
        if (error) reject(error); else resolve(result);
      };
      const onAbort = () => finish(Object.assign(new Error('Cancelled'), { name: 'AbortError' }));
      const timer = setTimeout(() => finish(new Error('Muxing timed out')), 120000);
      signal.addEventListener('abort', onAbort, { once: true });
      worker.onmessage = (e) => {
        if (e.data.progress) { const p = e.data.progress; onProgress(p.done, p.total, p.text); return; }
        if (e.data.error) { finish(new Error(e.data.error)); return; }
        if (e.data.result) { finish(null, e.data.result); }
      };
      worker.onerror = (e) => finish(new Error(e.message || 'mux-worker crashed'));
      const video = toArrayBuffers(videoChunks);
      const audio = toArrayBuffers(audioChunks);
      try {
        worker.postMessage({ video, audio }, [...video, ...audio]);
      } catch (_) {
        try { worker.postMessage({ video, audio }); } catch (error) { finish(error); }
      }
    });
  }

  function downloadFile(data, filename) {
    const blob = new Blob(Array.isArray(data) ? data : [data], { type: 'video/mp4' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  // ===========================================================================
  // Orchestration
  // ===========================================================================
  const HARD_DRM_SCHEMES = [
    'edef8ba9-79d6-4ace-a3c8-27dcd51d21ed', // Widevine
    '9a04f079-9840-4286-ab92-e65be0885f95', // PlayReady
    '94ce86fb-07ff-4f43-adb8-93d2fa968ca2'  // FairPlay
  ];

  async function runDownload(capture, filename, onProgress, signal) {
    onProgress(0, 1, 'Fetching manifest...');
    const resp = await fetch(capture.manifestUrl, mediaFetchInit(capture.manifestUrl, capture, { signal }));
    if (!resp.ok) throw new Error(`Manifest fetch failed: HTTP ${resp.status}`);
    const xmlText = new TextDecoder().decode(await readLimitedResponse(resp, MAX_MANIFEST_BYTES));

    const cpSchemes = [...xmlText.matchAll(/<ContentProtection\b[^>]*schemeIdUri="([^"]+)"/gi)].map(m => m[1].toLowerCase());
    if (cpSchemes.some(s => HARD_DRM_SCHEMES.some(uuid => s.includes(uuid)))) {
      throw Object.assign(new Error('DRM_PROTECTED'), { isDrm: true });
    }

    onProgress(0, 1, 'Parsing manifest...');
    const allTracks = parseDashManifest(xmlText, capture.manifestUrl);
    if (!allTracks.length) throw new Error('No tracks found in manifest');
    // Diagnostic: a recording cut short usually means multiple Periods or a
    // segment count that does not match the advertised duration.
    const periodCount = (xmlText.match(/<Period\b/g) || []).length;
    const mpdDuration = (xmlText.match(/mediaPresentationDuration="([^"]+)"/) || [])[1] || '?';
    console.info('[Teams DL] manifest:', { periods: periodCount, mediaPresentationDuration: mpdDuration,
      tracks: allTracks.map(t => ({ type: t.type, segments: t.segments.length, encrypted: !!t.encryption })) });

    const videoTrack = allTracks.find(t => t.type === 'video' || t.type === 'muxed');
    const audioTrack = allTracks.find(t => t.type === 'audio');
    const safe = filename.replace(/[^a-z0-9\s_-]/gi, '_');

    if (videoTrack && audioTrack) {
      const trackData = await downloadDashSegments([videoTrack, audioTrack], capture, onProgress, signal);
      const muxed = await muxTracks(trackData[0], trackData[1], onProgress, signal);
      downloadFile(muxed, safe + '.mp4');
    } else {
      const only = videoTrack || audioTrack || allTracks[0];
      const trackData = await downloadDashSegments([only], capture, onProgress, signal);
      downloadFile(trackData[0], safe + (only.type === 'audio' ? '.m4a' : '.mp4'));
    }
    onProgress(1, 1, 'Download complete!');
  }

  // ===========================================================================
  // UI
  // ===========================================================================
  let btn, statusEl;
  function injectButton() {
    if (btn) return;
    const wrap = document.createElement('div');
    wrap.style.cssText = 'position:fixed;bottom:20px;right:20px;z-index:2147483647;font-family:Segoe UI,system-ui,sans-serif;display:none;flex-direction:column;align-items:flex-end;gap:6px;';
    statusEl = document.createElement('div');
    statusEl.style.cssText = 'background:#222;color:#fff;padding:4px 10px;border-radius:6px;font-size:12px;max-width:280px;display:none;';
    btn = document.createElement('button');
    btn.textContent = '⬇ Download recording';
    btn.style.cssText = 'background:#6264a7;color:#fff;border:0;padding:10px 16px;border-radius:8px;font-size:14px;font-weight:600;cursor:pointer;box-shadow:0 2px 8px rgba(0,0,0,.3);';
    btn.addEventListener('click', onClick);
    wrap.appendChild(statusEl); wrap.appendChild(btn);
    (document.body || document.documentElement).appendChild(wrap);
    btn._wrap = wrap;
  }
  function showButton() {
    injectButton();
    if (btn) btn._wrap.style.display = 'flex';
  }
  function hideButton() {
    if (btn) btn._wrap.style.display = 'none';
  }
  function setStatus(text) {
    if (!statusEl) return;
    statusEl.textContent = text;
    statusEl.style.display = text ? 'block' : 'none';
  }

  async function onClick(event) {
    if (!event.isTrusted) return;
    if (downloading) return;
    if (!videoManifestUrl) { setStatus('Play the recording for a few seconds first.'); return; }
    downloading = true;
    btn.disabled = true;
    const controller = new AbortController();
    const capture = Object.freeze({ manifestUrl: videoManifestUrl, spopactoken: videoSpopActoken });
    const filename = (document.title || 'teams_recording').replace(/[^a-z0-9\s_-]/gi, '_').trim() || 'teams_recording';
    const onProgress = (done, total, text) => setStatus(text || `${done}/${total}`);
    try {
      await runDownload(capture, filename, onProgress, controller.signal);
      setStatus('Done! Saved to your Downloads.');
    } catch (err) {
      console.error('[Teams DL] download failed', err);
      if (err.isDrm) {
        setStatus('This recording uses hard DRM (Widevine/PlayReady) — cannot be decrypted client-side.');
      } else {
        setStatus('Failed: ' + (err.message || err));
      }
    } finally {
      controller.abort(); // stop outstanding segment requests after any failure
      downloading = false;
      btn.disabled = false;
    }
  }

  // Only the SharePoint Stream player page (…/_layouts/15/stream.aspx, also on
  // sharepoint-df.com) gets the instant placeholder — we don't want a floating
  // button on every SharePoint doc/list. Everywhere else the button still
  // appears the moment a recording is actually detected (capture-based reveal).
  function isStreamPage() {
    try {
      return /https:\/\/[^/]*sharepoint(?:-df)?\.com(?:\.mcas\.ms)?\/.*\/_layouts\/15\/stream\.aspx/i.test(location.href);
    } catch (_) {
      return false;
    }
  }

  // Surface the overlay button in the top frame immediately on Stream pages,
  // before any manifest is captured. Clicking it before playback prompts the
  // user to play the recording first; once a manifest is seen (here or in a
  // child frame) the button is wired to the real download.
  function ready() {
    injectButton();
    if (isTopFrame && isStreamPage()) showButton();
  }
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', ready);
  } else {
    ready();
  }
})();
