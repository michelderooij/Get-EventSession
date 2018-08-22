# Get-IgniteSession

## Getting Started

This script can download all the Microsoft Ignite slidedecks and videos that are
available from the Ignite portal. Video downloads will leverage a utility which can be downloaded
from https://yt-dl.org/latest/youtube-dl.exe, and needs to reside in the same folder
as the script. The script itself will try to download the utility when the utility is not present.

When you are interested in retrieving session information only, you can use the InfoOnly switch.

To prevent retrieving session information for every run, the script will cache session information

### Prerequisites

* PowerShell 3.0
* YouTube-dl.exe (automatic download from [here](https://yt-dl.org))

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

## Credits

* Mattias Fors ([blog](http://deploywindows.info))
* Scott Ladewig ([blog](http://ladewig.com))
* Tim Pringle ([blog](http://www.powershell.amsterdam))
* Andy Race ([GitHub](https://github.com/AndyRace))

## License

This project is licensed under the MIT License - see the LICENSE.md for details.

 