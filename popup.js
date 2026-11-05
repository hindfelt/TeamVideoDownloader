// popup.js

document.addEventListener('DOMContentLoaded', function () {
    console.log('Popup loaded');
    const ffmpegButton = document.getElementById('copyCommand');
    const ytDlpButton = document.getElementById('copyYtDlpCommand');
    const status = document.getElementById('status');

    ffmpegButton.addEventListener('click', async () => {
        status.textContent = 'Processing...';

        try {
            const tab = await getActiveTab();
            const [capture, pageInfo] = await Promise.all([
                getManifestCapture(tab.id),
                getPageInfo(tab.id)
            ]);

            const manifestUrl = capture.manifestUrl || pageInfo.manifestUrl;
            const headers = capture.headers || null;
            const meetingName = sanitizeFileName(pageInfo.title || tab.title || 'teams_recording');

            if (!manifestUrl) {
                status.textContent = 'No manifest URL found. Play the recording for a few seconds, then click again.';
                return;
            }

            const command = buildFfmpegCommand(removeEnableCdn(manifestUrl), headers, meetingName);
            await navigator.clipboard.writeText(command);
            status.textContent = hasTransportAuth(headers)
                ? 'ffmpeg command with auth copied.'
                : 'ffmpeg command copied. Use yt-dlp if it returns 401.';
        } catch (error) {
            console.error('Error:', error);
            status.textContent = `Error: ${error.message}`;
        }
    });

    ytDlpButton.addEventListener('click', async () => {
        status.textContent = 'Processing...';

        try {
            const tab = await getActiveTab();
            const [capture, pageInfo] = await Promise.all([
                getManifestCapture(tab.id),
                getPageInfo(tab.id)
            ]);
            const pageUrl = pageInfo.streamUrl || pageInfo.pageUrl || tab.url;

            if (!pageUrl) {
                status.textContent = 'Could not read the recording page URL.';
                return;
            }

            const command = buildYtDlpCommand(pageUrl);
            await navigator.clipboard.writeText(command);
            status.textContent = pageInfo.streamUrl || capture.manifestUrl
                ? 'yt-dlp command copied.'
                : 'yt-dlp command copied for this page.';
        } catch (error) {
            console.error('Error:', error);
            status.textContent = `Error: ${error.message}`;
        }
    });
});

async function getActiveTab() {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab?.id) throw new Error('No active tab found');
    return tab;
}

async function getManifestCapture(tabId) {
    try {
        return await chrome.runtime.sendMessage({ type: 'getManifest', tabId });
    } catch (error) {
        console.log('Background lookup failed, falling back to page scan', error);
        return {};
    }
}

async function getPageInfo(tabId) {
    try {
        const [result] = await chrome.scripting.executeScript({
            target: { tabId },
            func: () => {
                function findManifestUrl() {
                    const storedUrl = document.body?.getAttribute('data-manifest-url');
                    if (storedUrl) return storedUrl;

                    for (const entry of performance.getEntriesByType('resource')) {
                        if (entry.name.includes('videomanifest')) return entry.name;
                    }

                    return null;
                }

                function findStreamUrl() {
                    const links = Array.from(document.querySelectorAll('a[href]'), (link) => link.href);
                    return links.find((href) => /https:\/\/[^/]*sharepoint(?:-df)?\.com\/.*\/_layouts\/15\/stream\.aspx/i.test(href)) || null;
                }

                return {
                    manifestUrl: findManifestUrl(),
                    pageUrl: location.href,
                    streamUrl: findStreamUrl(),
                    title: document.title
                };
            }
        });

        return result?.result || {};
    } catch (error) {
        console.log('Page scan failed', error);
        return {};
    }
}

function buildFfmpegCommand(url, headers, meetingName) {
    const args = ['ffmpeg'];

    if (headers?.['user-agent']) {
        args.push('-user_agent', shellQuote(headers['user-agent']));
    }

    const headerLines = getHeaderLines(headers);
    if (headerLines.length) {
        args.push('-headers', ansiCStringQuote(headerLines.join('\r\n') + '\r\n'));
    }

    args.push('-i', shellQuote(url), '-codec', 'copy', shellQuote(`${meetingName}.mp4`));
    return args.join(' ');
}

function buildYtDlpCommand(pageUrl) {
    return `yt-dlp --cookies-from-browser chrome --merge-output-format mp4 ${shellQuote(pageUrl)}`;
}

function getHeaderLines(headers) {
    if (!headers) return [];

    return ['cookie', 'authorization', 'referer', 'origin']
        .filter((name) => headers[name])
        .map((name) => `${titleCaseHeader(name)}: ${headers[name]}`);
}

function hasTransportAuth(headers) {
    return Boolean(headers?.cookie || headers?.authorization);
}

function titleCaseHeader(name) {
    return name.replace(/(^|-)([a-z])/g, (_, sep, ch) => sep + ch.toUpperCase());
}

function removeEnableCdn(url) {
    return url.replace(/([?&])enableCdn=\d+(&?)/i, (_match, separator, trailingAmp) => {
        return trailingAmp ? separator : '';
    });
}

function sanitizeFileName(value) {
    return String(value || 'teams_recording')
        .replace('Microsoft Teams', '')
        .trim()
        .replace(/[^a-zA-Z0-9._-]+/g, '_')
        .replace(/^_+|_+$/g, '') || 'teams_recording';
}

function shellQuote(value) {
    return "'" + String(value).replace(/'/g, "'\\''") + "'";
}

function ansiCStringQuote(value) {
    return "$'" + String(value)
        .replace(/\\/g, '\\\\')
        .replace(/'/g, "\\'")
        .replace(/\r/g, '\\r')
        .replace(/\n/g, '\\n') + "'";
}
