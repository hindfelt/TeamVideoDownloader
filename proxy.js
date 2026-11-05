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

chrome.runtime.onInstalled.addListener(() => { syncScripts().catch(console.error); });
chrome.runtime.onStartup.addListener(() => { syncScripts().catch(console.error); });
chrome.permissions.onAdded.addListener(() => { syncScripts().catch(console.error); });
chrome.permissions.onRemoved.addListener(() => { syncScripts().catch(console.error); });

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
