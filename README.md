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

Get information of all sessions, and output only location and time information for sessions (co-)presented by Tony Redmond:
```
.\Get-EventSession.ps1 -InfoOnly | Where {$_.Speakers -contains 'Tony Redmond'} | Select Title, location, startDateTime
```

Download all available contents of sessions BRK3248 and BRK3186 to D:\Ignite:
```
.\Get-EventSession.ps1 -DownloadFolder D:\Ignite -ScheduleCode BRK3248,BRK3186
```

## FAQ

### MSA authentication for protected content
Some events (Custom events or sessions with protected slidedecks) require you to sign in with a Microsoft account.
When the script detects this, it will open an embedded sign-in dialog automatically. After signing in once, the
session is cached for the remainder of the run so you are not prompted again.

For Custom events, specify the base URL with `-EventUrl` and use `-Event Custom`:
```
.\Get-EventSession.ps1 -Event Custom -EventUrl https://example.com/sessions -DownloadFolder D:\Event
```

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

 
