// proxy.js
// Opt-in support for SharePoint behind the Microsoft Defender for Cloud Apps
// session proxy (hosts like contoso-my.sharepoint.com.mcas.ms). The proxy
// hosts are optional permissions so existing users get no new install
// warning. Clicking the toolbar icon on a proxied page asks for access once,
// then the same content scripts the manifest declares for SharePoint are
// registered for the proxy host.

const PROXY_ORIGINS = [
  'https://*.sharepoint.com.mcas.ms/*',
  'https://*.svc.ms.mcas.ms/*'
];
const PAGE_MATCHES = ['https://*.sharepoint.com.mcas.ms/*'];
const PROXY_PAGE_RE = /^https:\/\/[^/]+\.sharepoint\.com\.mcas\.ms\//i;
const SCRIPT_IDS = ['proxy-intercept', 'proxy-content'];

const SCRIPTS = [
  {
    id: 'proxy-intercept',
    matches: PAGE_MATCHES,
    js: ['intercept.js'],
    runAt: 'document_start',
    allFrames: true,
    world: 'MAIN'
  },
  {
    id: 'proxy-content',
    matches: PAGE_MATCHES,
    js: ['content.js'],
    runAt: 'document_idle',
    allFrames: true
  }
];

async function hasProxyAccess() {
  return chrome.permissions.contains({ origins: PROXY_ORIGINS });
}

// Make the registered scripts match the granted permission. Idempotent, so it
// is safe to run on install, update, browser start and permission changes.
async function syncScripts() {
  const granted = await hasProxyAccess();
  const existing = await chrome.scripting.getRegisteredContentScripts({ ids: SCRIPT_IDS });
  const existingIds = new Set(existing.map((s) => s.id));

  if (granted) {
    const missing = SCRIPTS.filter((s) => !existingIds.has(s.id));
    if (missing.length) await chrome.scripting.registerContentScripts(missing);
  } else if (existingIds.size) {
    await chrome.scripting.unregisterContentScripts({ ids: [...existingIds] });
  }
}

// Until proxy access is granted the extension cannot run on proxy pages, so
// it cannot show anything there itself. declarativeContent can still swap the
// toolbar icon on matching URLs without host access: show a dotted icon on
// proxied SharePoint pages as a hint to click it. Removed once access exists.
const HINT_RULE_ID = 'proxy-hint';
let hintImagePromise = null;

function hintImageData() {
  if (!hintImagePromise) {
    hintImagePromise = (async () => {
      const resp = await fetch(chrome.runtime.getURL('icons/team_lit.png'));
      const bitmap = await createImageBitmap(await resp.blob());
      const out = {};
      for (const size of [16, 32]) {
        const canvas = new OffscreenCanvas(size, size);
        const ctx = canvas.getContext('2d');
        ctx.drawImage(bitmap, 0, 0, size, size);
        const r = size * 0.22;
        ctx.beginPath();
        ctx.arc(size - r, r, r, 0, Math.PI * 2);
        ctx.fillStyle = '#f97316';
        ctx.fill();
        ctx.lineWidth = Math.max(1, size / 16);
        ctx.strokeStyle = '#ffffff';
        ctx.stroke();
        out[size] = ctx.getImageData(0, 0, size, size);
      }
      return out;
    })();
    hintImagePromise.catch(() => { hintImagePromise = null; });
  }
  return hintImagePromise;
}

function removeHintRules() {
  return new Promise((resolve) => chrome.declarativeContent.onPageChanged.removeRules(undefined, () => resolve()));
}

async function syncHint() {
  if (!chrome.declarativeContent) return;
  const granted = await hasProxyAccess();
  await removeHintRules();
  if (granted) return;
  const imageData = await hintImageData();
  await new Promise((resolve) => chrome.declarativeContent.onPageChanged.addRules([{
    id: HINT_RULE_ID,
    conditions: [new chrome.declarativeContent.PageStateMatcher({
      pageUrl: { schemes: ['https'], hostSuffix: '.sharepoint.com.mcas.ms' }
    })],
    actions: [new chrome.declarativeContent.SetIcon({ imageData })]
  }], () => resolve()));
}

function syncAll() {
  syncScripts().catch(console.error);
  syncHint().catch(console.error);
}

chrome.runtime.onInstalled.addListener(syncAll);
chrome.runtime.onStartup.addListener(syncAll);
chrome.permissions.onAdded.addListener(syncAll);
chrome.permissions.onRemoved.addListener(syncAll);

chrome.action.onClicked.addListener((tab) => {
  // activeTab exposes tab.url only for the tab the user clicked on.
  if (!tab || !tab.url || !PROXY_PAGE_RE.test(tab.url)) return;

  // permissions.request must run synchronously inside the click gesture,
  // before any await.
  chrome.permissions.request({ origins: PROXY_ORIGINS }, async (granted) => {
    if (chrome.runtime.lastError || !granted) return;
    try {
      await syncScripts();
      await chrome.tabs.reload(tab.id);
    } catch (error) {
      console.error('[Teams DL] Proxy setup failed:', error);
    }
  });
});
