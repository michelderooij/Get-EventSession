# Get-EventSession

## Getting Started

Script to assist in downloading Microsoft Ignite, Inspire, Build, MEC or Custom event contents, or return
session information for easier digesting.

Video downloads leverage yt-dlp, which the script will download automatically when not present.
ffmpeg is also required to merge separate video and audio streams; the script will download it automatically as well.

To prevent retrieving session information for every run, the script caches session information locally.

### Prerequisites

* PowerShell 5.1 or later (PowerShell 7 recommended)
* yt-dlp.exe (automatic download from [github.com/yt-dlp/yt-dlp](https://github.com/yt-dlp/yt-dlp/releases/latest))
* ffmpeg.exe (automatic download from [github.com/BtbN/FFmpeg-Builds](https://github.com/BtbN/FFmpeg-Builds/releases))

### Usage

Download all available contents of Build 2026 sessions containing the word 'AI' in the title to D:\Build:
```
.\Get-EventSession.ps1 -Event Build2026 -DownloadFolder D:\Build -Keyword 'AI'
```

Download only 200- and 300-level Python sessions from Build 2026:
```
.\Get-EventSession.ps1 -Event Build2026 -DownloadFolder D:\Build -ProgrammingLanguage Python -SessionLevel 200,300
```

Get archived Ignite 2024 sessions from the Ignite archive portal and show their info:
```
.\Get-EventSession.ps1 -Event Ignite2024 -InfoOnly
```

Get information of all sessions, and output only location and time information for sessions (co-)presented by Tony Redmond:
```
.\Get-EventSession.ps1 -InfoOnly | Where {$_.Speakers -contains 'Tony Redmond'} | Select Title, location, startDateTime
```

Download all available contents of sessions BRK3248 and BRK3186 to D:\Ignite:
```
.\Get-EventSession.ps1 -DownloadFolder D:\Ignite -ScheduleCode BRK3248,BRK3186
```

## FAQ

### Microsoft authentication for protected content

Custom event URLs are not restricted to a particular hostname. When no valid catalog cache is available, the
script inspects the catalog anonymously first. Public catalogs are retrieved without signing in. Authentication
responses trigger the existing shared Microsoft authentication support. OAuth redirects and challenges identify
personal (MSA) or work/school (Microsoft 365) authorities where possible; ambiguous portals retain their own
browser account selection. Cookie names alone cannot reliably identify the account type.

For Custom events, specify the base URL with `-EventUrl` and use `-Event Custom`:

```powershell
.\Get-EventSession.ps1 -Event Custom -EventUrl https://example.com/sessions/search -DownloadFolder D:\Event
```

GET catalogs with `items` and paged catalogs with `sessions` are supported. Retrieval first tries anonymous GET.
If that fails, it tries POST using matching-origin credentials from `MSAAuth.cache`. Without a matching cache,
POST is tried anonymously before sign-in. If these requests fail with a recoverable HTTP response, the existing
provider-aware authentication flow is used. No catalog-method parameter is needed. Transport failures and
server errors do not trigger sign-in. Portals needing browser-generated request context
can open a dedicated Edge window; sign in only if the portal requires it, open the event sessions page, and
click Continue. Keep Edge open until capture completes. Edge is required for browser request discovery;
yt-dlp is not required in InfoOnly mode.

```powershell
.\Get-EventSession.ps1 -Event Custom -EventUrl https://event.contoso.com/api/schedule-builder/sessions/search -Verbose -InfoOnly -Refresh
```

Relevant authentication cookies are stored in `MSAAuth.cache`, with an origin marker to prevent reuse for a
different host or port. The file holds the current portal's authentication; switching portals replaces that
entry after successful validation. Custom catalog caches remain endpoint-specific. `-Refresh` refreshes the
catalog, not necessarily your sign-in. Captured request context is encrypted for the current Windows user and
validated on reuse; rejected or expired context can require browser sign-in again. Cookies remain in Netscape
format. Treat authentication caches and their backups as credentials; do not share or commit them.

Session-search catalogs can contain summaries without media links. For compatible schedule-builder portals,
download mode fetches the selected session's `eventInfo` details and uses its published video and slide links.
If fresh detail context is needed, the existing Edge flow follows the selected session's published detail link
from the sessions page. Detail request context is encrypted in the auth cache for reuse. InfoOnly does not
perform these additional media lookups. A slide deck is downloaded only when the details publish one.
Compatible same-origin detail-page context is reused for other selected session IDs, avoiding a sign-in
prompt per session. Successful cached and in-memory authentication validation persists updated cookies and context.
An explicit false on-demand availability flag skips session content processing; true or missing flags
retain normal processing. Detail records with `virtual.videoEnabled = false` also skip content processing.

```powershell
.\Get-EventSession.ps1 -Event Custom -EventUrl https://myevent.contoso.com/api/schedule-builder/sessions/search -DownloadFolder D:\MyEvent -ScheduleCode KEY01
```

### Tests

Run the focused regression suite with Pester 5:

```powershell
Invoke-Pester -Path .\Get-EventSession.Tests.ps1
```

### Archive-backed event metadata
Archive-backed events such as Ignite2024 are derived from HTML session cards instead of a structured catalog API.
For those events, the script stamps properties when the card exposes enough information.
Some metadata remains source-dependent, so property availability can vary between event types.

### YouTube authentication
YouTube requires authentication to prevent automated downloads. You will see yt-dlp errors like
"Sign in to confirm you're not a bot. Use --cookies-from-browser or --cookies for the authentication".
To work around this, you can either:
* Use direct downloads when available with `-PreferDirect`. You may end up with larger files, but you can
  compress and downscale them afterwards using the `Compress-MP4.ps1` script in this repository.
* Pass cookies to yt-dlp from a cookies file (Netscape format) or directly from a browser. See
  [yt-dlp cookie export docs](https://github.com/yt-dlp/yt-dlp/wiki/Extractors#exporting-youtube-cookies) for details.
  Using a separate account is recommended when downloading large numbers of videos.

Syntax:
```
.\Get-EventSession.ps1 .. -CookiesFile <File>
.\Get-EventSession.ps1 .. -CookiesFromBrowser <brave|chrome|chromium|edge|firefox|opera|safari|vivaldi|whale>
```

### Why do downloads happen twice?
Depending on the source, the video and audio streams are seperate. First the video stream is fetched, then the audio stream.
After fetching, the two are merged.

### How to specify format?
Depending on availability and source, the default format is worstvideo+bestaudio/best. This means the worst quality video and best audio
stream are fetched and merged. 'Best' is attempted if no selection could be made. You can also specify bestvideo+bestaudio to get the best
quality video, but these files can be substantially larger. You can also perform more complex filter, e.g.
bestvideo[height=540][filesize<384MB]+bestaudio,bestvideo[height=720][filesize<512MB]+bestaudio,bestvideo[height=360]+bestaudio,bestvideo+bestaudio
1) This would first attempt to download the video of 540p if it is less than 384MB, and best audio.
2) When not present, then attempt to downlod video of 720p less than 512MB.
3) Thirdly, attempt to download video of 360p with best audio.
4) If none of the previous filters found matches, just pick the best video and best audio streams

## Credits

Special thanks to [Mattias Fors](http://deploywindows.info), [Scott Ladewig](http://ladewig.com), [Tim Pringle](http://www.powershell.amsterdam), and [Andy Race](https://github.com/AndyRace).

## License

This project is licensed under the MIT License - see the LICENSE.md for details.


