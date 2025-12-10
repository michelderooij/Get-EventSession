# Get-EventSession

## Getting Started

Script to assist in downloading Microsoft Ignite, Inspire, Build or MEC contents or return 
session information for easier digesting. 

Video downloads will leverage a utility which can be downloaded
from https://yt-dl.org/latest/youtube-dl.exe, and needs to reside in the same folder
as the script. The script itself will try to download the utility when the utility is not present.

To prevent retrieving session information for every run, the script will cache session information.

### Prerequisites

* PowerShell 3.0
* YouTube-dl.exe (automatic download from [here](https://yt-dl.org))
* ffmpeg, (automatic download from [here](https://ffmpeg.zeranoe.com/builds/win32/static/ffmpeg-latest-win32-static.zip)) 

### Usage

Download all available contents of sessions containing the word 'Exchange' in the title to D:\Ignite:
```
.\Get-EventSession.ps1 -DownloadFolder D:\Ignite -Format 18 -Keyword 'Exchange'
```

Get information of all sessions, and output only location and time information for sessions (co-)presented by Tony Redmond:
```
.\Get-EventSession.ps1 -InfoOnly | Where {$_.Speakers -contains 'Tony Redmond'} | Select Title, location, startDateTime
```

Download all available contents of sessions BRK3248 and BRK3186 to D:\Ignite
```
.\Get-EventSession.ps1 -DownloadFolder D:\Ignite -ScheduleCode BRK3248,BRK3186
```

## FAQ

### YouTube authentication
Recently, YouTube started requiring authentication to prevent automated downloads. Symptom is that you will see yt-dlp operations
resulting in "ERROR [youtube] XXXXXXXXXX : Sign in to confirm you're not a bot. Use --cookies-from-browser or --cookies for the authentication" errors.
To support this, you can either:
* Use direct downloads when available, by using PreferDirect. Disadvantage is that you cannot specify a format, and you may end up with large files. However, you can
compress and descale those in bulk afterwards using the Compress-MP4.ps1 script from this same repository.
* Use yt-dlp.exe with downloaded cookies (NetScape format) or cookies from a browser. See https://github.com/yt-dlp/yt-dlp/wiki/Extractors#exporting-youtube-cookies
on how to export cookies. Be advised that your account might get flagged, so using a burner account is recommended when downloading lots of videos.
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

 
