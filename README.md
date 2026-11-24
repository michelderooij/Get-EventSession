# Get-EventSession

## Getting Started

Download videos and presentations from Microsoft Ignite, Inspire, Build, MEC, or Custom events.
You can also retrieve session information without downloading content.

The script automatically downloads yt-dlp and ffmpeg when needed for video downloads.

Session information is saved locally to speed up later runs. Use `-Refresh` to retrieve the latest session list.

### Prerequisites

* Windows PowerShell 5.1 or PowerShell 7 or later
* yt-dlp.exe (automatic download from [github.com/yt-dlp/yt-dlp](https://github.com/yt-dlp/yt-dlp/releases/latest))
* ffmpeg.exe (automatic download from [github.com/BtbN/FFmpeg-Builds](https://github.com/BtbN/FFmpeg-Builds/releases))

### Usage

Download available videos and presentations for Build 2026 sessions with 'AI' in the title:
```powershell
.\Get-EventSession.ps1 -Event Build2026 -DownloadFolder D:\Build -Keyword 'AI'
```

Download only 200- and 300-level Python sessions from Build 2026:
```powershell
.\Get-EventSession.ps1 -Event Build2026 -DownloadFolder D:\Build -ProgrammingLanguage Python -SessionLevel 200,300
```

Show session information for Ignite 2024:
```powershell
.\Get-EventSession.ps1 -Event Ignite2024 -InfoOnly
```

Show the title, location, and start time of Ignite sessions presented by Tony Redmond:
```powershell
.\Get-EventSession.ps1 -Event Ignite -InfoOnly | Where-Object {$_.SpeakerNames -contains 'Satya Nadella'} | Select-Object Title, location, startDateTime
```

Download available content for two Ignite sessions:
```powershell
.\Get-EventSession.ps1 -Event Ignite -DownloadFolder D:\Ignite -ScheduleCode BRK3248,BRK3186
```

## FAQ

### Microsoft authentication for protected content

Public session information usually needs no sign-in. For protected content, sign in with a personal
Microsoft account or a work/school account that has access to the event.
For Custom events, you can either follow the script's sign-in prompts or use cookies exported from your browser.

#### Signing in when prompted

For Custom events, use `-Event Custom` and provide the event's session catalog URL with `-EventUrl`:

```powershell
.\Get-EventSession.ps1 -Event Custom -EventUrl https://customevent.microsoft.com/api/schedule-builder/sessions -DownloadFolder D:\Event
```

- Follow any Microsoft sign-in prompt. If an Edge window opens, sign in if required, open the event's
  sessions page, and click **Continue**. Leave Edge open until the script finishes collecting the information.
- Saved sign-in details may be reused on later runs. You may need to sign in again when they expire or
  when switching events.
- `-Refresh` retrieves a fresh session list; it does not force a new sign-in.

If retrieval fails, check the event URL, your access to the portal, and your network or proxy connection.
Add `-Verbose` for more information.

Windows PowerShell 5.1 and PowerShell 7 are supported. Requests that inspect authentication
redirects do not automatically follow them; on Windows PowerShell 5.1 the script uses a
compatible request path to preserve the original HTTP status and authentication headers.
Sign-in options still depend on the portal and compatible identity libraries available on the device.

#### Using exported cookies

Use `-CookiesFile` to access protected Custom sessions and their videos, presentations, and captions
with cookies exported from your browser. When authentication is needed, the script uses this file
instead of prompting you to sign in.

1. Install **Get cookies.txt LOCALLY** in the browser you use to access the portal:
   - **Edge or Chrome:** install from the
     [Chrome Web Store](https://chromewebstore.google.com/detail/get-cookiestxt-locally/cclelndahbckbenkjhflpdbgdldlbecc).
     In Edge, allow extensions from other stores if prompted, then confirm installation.
   - **Firefox:** install from
     [Firefox Add-ons](https://addons.mozilla.org/en-US/firefox/addon/get-cookies-txt-locally/).
2. Open the event's session portal, sign in, and confirm that you can view its sessions and available media.
3. While on the signed-in portal, open the extension and export the site's cookies in **Netscape
   cookies.txt format**. Save the export outside this repository, for example
   `.\event_cookies.txt`. Export only the sites needed for the event.
4. Pass the saved file to the script along with `-Event Custom` and the catalog's `-EventUrl`.

Fetch fresh session information:

```powershell
.\Get-EventSession.ps1 -Event Custom -EventUrl https://customevent.microsoft.com/api/schedule-builder/sessions -CookiesFile '.\custom_cookies.txt' -InfoOnly -Refresh
```

Download published videos and decks:

```powershell
.\Get-EventSession.ps1 -Event Custom -EventUrl https://customevent.microsoft.com/api/schedule-builder/sessions -CookiesFile '.\custom_cookies.txt' -DownloadFolder D:\Custom
```

`-CookiesFile` also works with `-PreferDirect`. For Custom events it takes precedence over
`-CookiesFromBrowser`; if both are supplied, the browser option is ignored with a warning.
For other events, `-CookiesFile` is used for video downloads, not for signing in to the session portal.

If the file is invalid, expired, or rejected, the script reports an error without opening a sign-in window.
Sign in again in your browser and export fresh cookies. `-Refresh` updates the session list, not the cookies.
Some portals cannot be accessed with exported cookies alone; in that case, omit `-CookiesFile` to use
the normal sign-in flow. Media hosted on a separate protected site may need cookies from that site too.

Your export is not modified or saved into the script's authentication cache.
**Treat cookie exports, saved authentication files, and their backups as credentials:** keep them private,
restrict file access, and never share or commit them.
Review the browser extension's permissions before installing it.

### What content is available for Custom events?

Only content published by the event and accessible to your account can be downloaded. Some sessions
have presentations but no video. The script checks selected sessions for available videos and decks
when downloading; `-InfoOnly` returns session information without these additional checks.

- Use `-NoVideos` to download presentations without videos.
- Use `-NoSlidedecks` to skip presentations.
- Presentations can be PowerPoint or PDF files. If a session has multiple decks, the first is named
  `SessionCode-Title.pptx`, followed by `SessionCode-Title-1.pptx`, `SessionCode-Title-2.pptx`, and so on,
  using each file's actual extension.
- Existing files are skipped unless you specify `-Overwrite`. Older filenames are not renamed automatically.

### Why is some session information missing?

The information available varies by event. Older events, such as Ignite2024, may not include every
speaker, location, or timing detail.

### YouTube authentication
YouTube may require you to sign in before downloading. You may see an error such as
"Sign in to confirm you're not a bot. Use --cookies-from-browser or --cookies for the authentication".
To work around this, you can either:
* Use direct downloads when available with `-PreferDirect`. Files may be larger; use
  [Compress-MP4.ps1](./Compress-MP4.ps1) to compress or downscale them afterwards.
* Pass cookies to yt-dlp from a cookies file (Netscape format) or directly from a browser. See
  [yt-dlp cookie export docs](https://github.com/yt-dlp/yt-dlp/wiki/Extractors#exporting-youtube-cookies) for details.
  Using a separate account is recommended when downloading large numbers of videos.

Options:
```powershell
.\Get-EventSession.ps1 .. -CookiesFile <File>
.\Get-EventSession.ps1 .. -CookiesFromBrowser <brave|chrome|chromium|edge|firefox|opera|safari|vivaldi|whale>
```

### Why do downloads happen twice?
Some sources provide video and audio separately. The script downloads both, then merges them into one file.

### How do I choose video quality?

Use `-Format` to choose video quality. The default, `worstvideo+bestaudio/best`, selects the lowest-quality
video and best audio available, with a combined format as a fallback.

For higher quality, use `-Format 'bestvideo+bestaudio/best'`. Downloads may be substantially larger.
`-PreferDirect` ignores `-Format`.

For more control, provide a list of preferred formats:

```powershell
.\Get-EventSession.ps1 -Event Build2026 -DownloadFolder D:\Build -Format 'bestvideo[height=540][filesize<384MB]+bestaudio,bestvideo[height=720][filesize<512MB]+bestaudio,bestvideo[height=360]+bestaudio,bestvideo+bestaudio'
```

This tries, in order:

1. A 540p video under 384 MB, with the best audio.
2. A 720p video under 512 MB, with the best audio.
3. A 360p video, with the best audio.
4. The best available video and audio if none of those options is available.

## Credits

Special thanks to [Mattias Fors](http://deploywindows.info), [Scott Ladewig](http://ladewig.com), [Tim Pringle](http://www.powershell.amsterdam), and [Andy Race](https://github.com/AndyRace).

## License

This project is licensed under the MIT License - see [LICENSE.md](./LICENSE.md) for details.


