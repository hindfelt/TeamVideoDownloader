# Chrome Web Store listing (v2.2.2)

Paste these into the Developer Dashboard. The title and summary come from
`manifest.json`; the description, icon and screenshots are set on the
**Store listing** tab.

## Title (from manifest.json)

Recording Downloader for Microsoft Teams

## Summary (from manifest.json, max 132 characters)

Download Microsoft Teams / SharePoint Stream meeting recordings (in-browser, decrypts AES-128 segments).

## Store icon

icons/team_lit.png (128 x 128)

## Screenshot

promo.png (1280 x 800)

## Description

Save the meeting recordings you already have access to as a regular MP4 file, straight from your browser.

Recording Downloader for Microsoft Teams works on SharePoint Stream recording pages, where recordings of Microsoft Teams meetings are stored. Open a recording, play it for a few seconds, and click the "Download recording" button that appears in the bottom-right corner of the page. The extension does the rest and saves the finished .mp4 to your Downloads folder.

How it works
• Captures the recording's video manifest and short-lived access token as the player loads them.
• Downloads the audio and video segments in parallel, with automatic retry when SharePoint throttles requests.
• Decrypts AES-128 encrypted segments locally using your browser's Web Crypto API.
• Combines audio and video into a single, seekable MP4 that plays in VLC, QuickTime and other players.

Private by design
• Everything happens in your browser. No data is sent to the developer or any third party.
• The only network requests go to the same SharePoint/Stream servers the recording is streamed from.
• No account, no command line, no ffmpeg needed.

Optional support for the Defender for Cloud Apps session proxy
If your organization routes SharePoint through the Defender session proxy (*.sharepoint.com.mcas.ms), click the extension's toolbar icon on the recording page and allow access. The page reloads with downloading enabled. Check with your IT or security team first: the proxy enforces your organization's download policies.

Limitations
• Recordings protected with hard DRM (Widevine, PlayReady, FairPlay) cannot be downloaded.
• Downloads are limited to 3.5 GiB of audio/video per recording.
• If a download fails with HTTP 401 (the access token has expired), refresh the page, play the recording for a few seconds and try again.

Only download recordings you are authorized to access.

Disclaimer: This extension is an independent tool and is not affiliated with, endorsed by, or sponsored by Microsoft Corporation.
