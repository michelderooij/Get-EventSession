<#
    .SYNOPSIS
    Script to assist in downloading Microsoft Ignite, Inspire, Build or MEC contents, or return 
    session information for easier digesting. Video downloads will leverage external utilities, 
    depending on the used video format. To prevent retrieving session information for every run,
    the script will cache session information.

    Be advised that downloading of OnDemand contents from Azure Media Services is throttled to real-time
    speed. To lessen the pain, the script performs simultaneous downloads of multiple videos streams. Those
    downloads will each open in their own (minimized) window so you can track progress. Finally, CTRL-C
    is catched by the script because we need to stop download jobs when aborting the script.

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Michel de Rooij 	        http://eightwone.com
    Version 3.96, May 25th, 2023

    Special thanks to:
    Mattias Fors 	        http://deploywindows.info
    Scott Ladewig 	        http://ladewig.com
    Tim Pringle                 http://www.powershell.amsterdam
    Andy Race                   https://github.com/AndyRace
    Richard van Nieuwenhuizen

    .DESCRIPTION
    This script can download Microsoft Ignite, Inspire, Build and MEC session information and available 
    slidedecks and videos using MyIgnite/MyInspire/MyBuild techcommunity portal.

    Video downloads will leverage one or more utilities:
    - YouTube-dl, which can be downloaded from https://yt-dl.org/latest/youtube-dl.exe. This utility
      needs to reside in the same folder as the script. The script itself will try to download this 
      utility when the utility is not present.
    - ffmpeg, which can be downloaded from https://ffmpeg.zeranoe.com/builds/win32/static/ffmpeg-latest-win32-static.zip. 
      This utility needs to reside in the same folder as the script, or you need to specify its location using -FFMPEG. 
      The utility is used to bind the seperate video and audio streams of Azure Media Services files 
      in single files.

    When you are interested in retrieving session information only, you can use the InfoOnly switch.

    Note: MEC sessions are not published through the usual API, so I worked around it by digesting its playlist as
    if it were a catalog. Consequence is that filtering might be limited, eg. no Category or Product etc.

    .PARAMETER DownloadFolder
    Specifies location to download sessions to. When omitted, will use 'systemdrive'\'Event'.

    .PARAMETER Format
    The Format specified depends on the media hosting the source videos:
    - Direct Downloads
    - Azure Media Services
    - YouTube

    Azure Media Services
    ====================
    For Azure Media Services, default option is worstvideo+bestaudio/best. Alternatively, you can 
    select other formats (when present), e.g. bestvideo+bestaudio. Note that the format requested 
    needs to be present in the stream package. Storage required for bestvideo is significantly 
    more than worstvideo. Note that you can also provide complex filter and preference, e.g. 
    bestvideo[height=540][filesize<384MB]+bestaudio,bestvideo[height=720][filesize<512MB]+bestaudio,bestvideo[height=360]+bestaudio,bestvideo+bestaudio
    1) This would first attempt to download the video of 540p if it is less than 384MB, and best audio.
    2) When not present, then attempt to downlod video of 720p less than 512MB.
    3) Thirdly, attempt to download video of 360p with best audio.
    4) If none of the previous filters found matches, just pick the best video and best audio streams.

    For Azure Media Services, you could also use format tags, such as 1_V_video_1 or 1_V_video_3.
    Note that these formats might not be consistent for different streams, e.g. 1_V_video_1
    might represent 1280x720 in one stream, while corresponding to 960x540 in another. To 
    prevent this, usage of filters is recommended.

    YouTube
    =======
    For YouTube videos, you can use the following formats:
    160          mp4        256x144    DASH video  108k , avc1.4d400b, 30fps, video only
    133          mp4        426x240    DASH video  242k , avc1.4d400c, 30fps, video only
    134          mp4        640x360    DASH video  305k , avc1.4d401e, 30fps, video only
    135          mp4        854x480    DASH video 1155k , avc1.4d4014, 30fps, video only
    136          mp4        1280x720   DASH video 2310k , avc1.4d4016, 30fps, video only
    137          mp4        1920x1080  DASH video 2495k , avc1.640028, 30fps, video only
    18           mp4        640x360    medium , avc1.42001E,  mp4a.40.2@ 96k
    22           mp4        1280x720   hd720 , avc1.64001F,  mp4a.40.2@192k (best, default)

    You can use filters or priority when selecting the media:
    - Filters allow you to put criteria on the media you select to download, e.g. 
      "bestvideo[height<=540]+bestaudio" will download the video stream where video is 540p at 
      most plus the audio stream (and ffmpeg will combine the two to a single MP4 file). It 
      allows you also to do cool things like "bestvideo[filesize<200M]+bestaudio". 
    - Priority allows you to provide additional criteria if the previous one fails, such as 
      when a desired quality is not available, e.g. "bestvideo+bestaudio/worstvideo+bestaudio" 
      will download worst video and best audio stream when the best video and audio streams 
      are not present.

    Format selection filter courtesey of Youtube-DL; for more examples, see 
    https://github.com/ytdl-org/youtube-dl/blob/master/README.md#format-selection-examples

    Direct Downloads
    ================
    Direct Downloads are downloaded directly from the provided downloadVideoLink source.

    .PARAMETER Captions
    When specified, for Azure Media Services contents, downloads caption files where available. 
    Files are usually in VTT format, and playable by VLC Player a.o. Note that captions might not always 
    be accurate due to machine translation, but at least will help in following the story :)

    .PARAMETER Subs
    When specified, for YouTube and Azure Media Services, downloads subtitles in provided languages by 
    specifying one or more 2-letter language codes seperated by a comma, e.g. en,fr,de,nl. Downloaded 
    subtitles may be in VTT or SRT format. Again, the subtitles might not always be accurate due to machine 
    translation. Note: For Azure Media Services, will only download caption in first language specified

    .PARAMETER Language
    When specified, for Azure Media hosted contents, downloads videos with specified audio stream where
    available. Note that if you mix this with specifying your own Format parameter, you need to
    add the language in the filter yourself, e.g. bestaudio[format_id*=German]. Default value is English, 
    as otherwise YouTube will download the last audio stream from the manifest (which often is Spanish).

    .PARAMETER Keyword
    Only retrieve sessions with this keyword in their session description.

    .PARAMETER Title
    Only retrieve sessions with this keyword in their session title.

    .PARAMETER Speaker
    Only retrieve sessions with this speaker.

    .PARAMETER Product
    Only retrieve sessions for this product. You need to specify the full product, subproducts seperated
    by '/', e.g. 'Microsoft 365/Office 365/Office 365 Management'. Wildcards are allowed.

    .PARAMETER Category
    Only retrieve sessions for this category. You need to specify the full category, subcategories seperated
    by '/', e.g. 'M365/Admin, Identity & Mgmt'. Wildcards are allowed.

    .PARAMETER SolutionArea
    Only retrieve sessions for this solution area. You need to specify the full 
    name, e.g. 'Modern Workplace'. Wildcards are allowed.

    .PARAMETER LearningPath
    Only retrieve sessions part of this this learningPath. You need to specify 
    the full name, e.g. 'Data Analyst'. Wildcards are allowed.

    .PARAMETER Topic
    Only retrieve sessions for this topic area. Wildcards are allowed.

    .PARAMETER ScheduleCode
    Only retrieve sessions with this session code. You can use one or more codes.

    .PARAMETER NoVideos
    Switch to indicate you don't want to download videos.

    .PARAMETER NoSlidedecks
    Switch to indicate you don't want to download slidedecks.

    .PARAMETER NoGuessing
    Switch to indicate you don't want the script to try to guess the URLs to retrieve media from MS Studios.

    .PARAMETER NoRepeats
    Switch to indicate you don't want the script to download repeated sessions.

    .PARAMETER FFMPEG
    Specifies full location of ffmpeg.exe utility. When omitted, it is searched for and
    when required extracted to the current folder.

    .PARAMETER MaxDownloadJobs
    Specifies the maximum number of concurrent downloads.

    .PARAMETER Proxy
    Specify the URI of the proxy to use, e.g. http://proxy:8080. When omitted, the current 
    system settings will be used.

    .PARAMETER Start
    Item number to start crawling with - useful for restarts.

    .PARAMETER Event
    Specify what event to download sessions for. 
    Options are:
    - Ignite                  : Ignite events (current)
    - Ignite2022,Ignite2021   : Ignite contents from that year/time
    - Inspire                 : Inspire contents (current)
    - Inspire2022,Inspire2021 : Inspire contents from that year
    - Build                   : Build contents (current)
    - Build2023,Build2022     : Build contents from that year
    - MEC                     : MEC contents

    .PARAMETER OGVPicker
    Specify that you want to pick sessions to download using Out-GridView.

    .PARAMETER InfoOnly
    Tells the script to return session information only.
    Note that by default, only session code and title will be displayed.

    .PARAMETER Overwrite
    Skips detecting existing files, overwriting them if they exist.

    .PARAMETER PreferDirect
    Instructs script to prefer direct video downloads over Azure Media Services, when both are 
    available. Note that direct downloads may be faster, but offer only single quality downloads, 
    where AMS may offer multiple video qualities.

    .PARAMETER Timestamp
    Tells script to change the timestamp of the downloaded media files to match the original 
    session timestamp, when available.

    .PARAMETER Locale
    When supported by the event, filters sessions on localization. 
    Currently supported: de-DE, zh-CN, en-US, ja-JP, es-CO, fr-FR. 
    When omitted, defaults to en-US.

    .PARAMETER Refresh
    When specified, this switch will try fetch current catalog information from online, ignoring
    any cached information which might be present.

    .NOTES
    The youtube-dl.exe utility requires Visual C++ 2010 redist package
    https://www.microsoft.com/en-US/download/details.aspx?id=5555

    Changelog
    =========
    2.0   Initial (Mattias Fors)
    2.1   Added video downloading, reformatting code (Michel de Rooij)
    2.11  Fixed titles with apostrophes
          Added Keyword and Title parameter
    2.12  Replaced pptx download Invoke-WebRequest with .NET webclient request (=faster)
          Fixed titles with backslashes (who does that?)
    2.13  Adjusts pptx timestamp to publishing timestamp
    2.14  Made filtering case-insensitive
          Added NoVideos to download slidedecks only
    2.15  Fixed downloading of differently embedded youtube videos
          Added timestamping of downloaded pptx files
          Minor output changes
    2.16  More illegal character fixups
    2.17  Bumped max post to check to 1750
    2.18  Added option to download for sessions listed in a schedule shared from MyIgnite
          Added lookup of video YouTube URl from MyIgnite if not found in TechCommunity
          Added check to make sure conversation titles begin with session code
          Added check to make sure we skip conversations we've already checked since some RSS IDs are duplicates
    2.19  Added trimming of filenames
    2.20  Incorporated Tim Pringle's code to use JSON to acess MyIgnite catalog
          Added option to select speaker
          Added caching of session information (expires in 1 day, or remove .cache file)
          Removed Start parameter (we're now pre-reading the catalog)
    2.21  Added proxy support, using system configured setting
          Fixed downloading of slidedecks
    2.22  Added URL parameter
          Renamed script to IgniteDownloader.ps1
    2.5   Added InfoOnly switch
          Added Product parameter
          Renamed script to Get-IgniteSession.ps1
    2.6   Fixed slide deck downloading
          Added Overwrite switch
    2.61  Added placeholder slide deck removal
    2.62  Fixed Overwrite logic bug
          Renamed to singular Get-IgniteSession to keep in line with PoSH standards
    2.63  Fixed bug reporting failed pptx download
          Added reporting of placeholder decks and videos
    2.64  Added processing of direct download links for videos
    2.65  Added option to specify multiple sessionCode codes
          Added note in source that format only works for YouTube video downloads.
          Added youtube-dl returncode check in case it won't run (e.g. missing VC library).
    2.66  Added proper downloading of session info using UTF-8 (no more '???')
          Additional trimming of spaces and CRLF's in property values
    2.7   Added Event parameter to switch between Ignite and Inspire catalog
          Renamed script to Get-EventSession
          Changed cached session info name to include event
          Removed obsolete URL parameter
          Added code to download slidedecks in PDF (Inspire)
          Cleanup of script synopsis/description/etc.
    2.8   Added downloading of Azure Media Services hosted streaming media
          Added simultaneous downloading of AMS hosted OnDemand streams
          Added NoSlidedecks switch
    2.9   Added Category parameter
          Fixed searching on Product
          Increased itemsPerPage when retrieving catalog
    2.91  Update to video downloading routine due to changes in published session info
    2.92  Fix 'Could not create SSL/TLS secure channel' issues with Invoke-WebRequest
    2.93  Update to slidedeck downloading routine due to changes in published session info
    2.94  Fixed cleanup of finished jobs
    2.95  Fixed encoding of filenames
    2.96  Fixed terminating cleanup when no slidedecks are being downloaded
          Added testing for contents to show contents is not available rather than generic 'problem'
    2.97  Update to change in video downloading location (YouTube)
          Changed default Format due to switch in video hosting - see YouTube format table
    2.971 Changed regex for YouTube matching to skip 'Coming Soon'
          Made verbose mode less noisy
    2.98  Converted background downloads to single background job queue
          Cosmetics
    2.981 Added cleanup of occasional leftovers (eg *.mp4.f5_A_aac_UND_2_192_1.ytdl, *.f1_V_video_3.mp4)
    2.982 Minor tweaks
    2.983 Added OGVPicker switch
    2.984 Changed keyword search to description, not abstract
          Fixed searching for Products and Category
          Added searching for SolutionArea
          Added searching for LearningPath
    2.985 Added Proxy support
    2.986 Minor update to accomodate publishing of slideDecks links
    3.0   Added Build support
    3.01  Added CTRL-Break notice to 'waiting for downloads' message
          Fixed 'No video located for' message
    3.1   Updated to work with the Inspire 2019 catalog
          Cosmetics
    3.11  Some more Cosmetics
    3.12  Updated to work with current Ignite & Build catalogs
          Bumped the download retry limits for YouTube-dl a bit
    3.13  Updated Ignite catalog endpoints
    3.14  Removed superfluous testing loading of main event page
          Fixed LearningPath option verbose output
          Some code cosmetics
    3.15  Added Topic parameter
    3.16  Corrected prefixes for Ignite 2019
    3.17  Added NoGuess switch
          Added NoRepeats switch
    3.18  Added Ignite2018 event
    3.19  Fixed video downloading
    3.20  Fixed background job cleanup
    3.21  Added Timestamp switch
          Updated file naming to strip embbeded name of format, e.g. f1_V_video_3
          Added stopping of Youtube-DL helper app spawned processes
    3.22  Added skipping of processing future sessions
    3.23  Added Captions switch and Subs parameter
          Added skipping of additional repeats (schedule code ending in R2/R3)
          Fixed filename construction containing '%'
          Added filtering options to description of Format parameter
          Decreased probing/retrieving video URLs from Azure Media Services (speed benefit)
    3.24  Added PreferDirect switch
          Enhanced Format parameter description
    3.25  Updated Youtube-DL download URL
    3.26  Updated mutual exclusion for PreferDirect & other parameters/switches
          Added workaround for long file names (NT Style name syntax)
          Added PowerShell ISE detection
          Added Garbage Collection
    3.27  Reworked jobs for downloading videos
          Added status bars for downloading of videos
          Failed video downloads will show last line of error output
          Added replacement of square brackets in file names
          Removed obsolete Clean-VideoLeftOvers call
    3.28  Uncommented line to cleanup output files after downloading video
          Changed 'Error' lines to single line outputs or throws (where appropriate)
    3.29  Added 'Stopped downloading ..' messages when terminating
    3.30  Increased wait cycle during progress refresh
          Added schedule code to progress status
          Revised detection successful video downloads
    3.31  Corrected video cleanup logic
    3.32  Do not assume Slidedeck exists when size is 0
    3.33  Fixed typo when specifying format for direct YouTube downloads
    3.34  Updated for Build 2020
          Added NoRepeat filtering for Build 2020
          Made Event parameter mandatory, and not defaulting to Ignite
          Added filtering example to Format parameter spec
    3.35  Updated for Inspire 2020
    3.36  Small fix for Inspire repeat session naming
    3.37  Added ExcludecommunityTopic parameter (so you can skip 'Fun and Wellness' Animal Cam contents)
          Modified Keyword and Title parameters (can be multiple values now)
    3.38  Added detection of filetype for presentations (PPTX/PDF)
    3.39  Added code to deal with specifying <Event><Year>
    3.40  Modified API endpoint for Ignite 2020
          Changed yearless Event specification to add year suffix, eg Ignite->Ignite2020, etc.
          Fixed Azure Media Services video scraping for Ignite2020
    3.41  Fixed: Error message for timeless sessions after downloading caption file
          Fixed: Downloading of caption files when video file is already downloaded
    3.42  Changed source location of ffmpeg. Download will now fetch current static x64 release.
    3.43  Fixed Ignite 2020 slidedeck 'trial & error' URL
    3.44  Fixed downloading of non-PDF slidedecks
    3.45  Help updated for -Event
    3.46  Changed downloading of caption files in background jobs as well
          Optimized caption downloading preventing unnecessary page downloads
    3.47  Added Captions to PreferDirect command set
    3.50  Updated for Ignite 2021
          Small cleanup
    3.51  Updated for Build 2021
    3.52  Updated NoRepeats maximum repeat check
          Added Language parameter to support Azure Media Services hosted videos containing multiple audio tracks
    3.53  Updated for Inspire 2021 
    3.54  Fixed adding Language filter when complex Format is specified
    3.55  Fixed audio stream selection when requested language is not available or only single audio stream is present
    3.60  Added support for Ignite 2021; specify individual event using Ignite2021H1 (Spring) or Ignite2021H2 (Fall)
    3.61  Added support for (direct) downloading of Ignite Fall 2021 videos
    3.62  Added Cleanup video leftover files if video file exists (to remove clutter)
          Changed lifetime of cached session information to 8 hours
          Fixed post-download counts
    3.63  Fixed keyword filtering
    3.64  Changed filter so that default language is picked when specified language is not available
    3.65  Updated for Build 2022
          Added Locale parameter to filter local content
          Fixed applying timestamp due to DateTime formatting changes
    3.66  Fixed filtering on langLocale
          Default Locale set to en-US
    3.67  Added removal of placeholder deck/video/vtt files
    3.68  Fixed caching when specifying Event without year tag, eg. Build vs Build2022
          Removed default Locale as that would mess things up for Events where data does not contain that information (yet).
    3.69  Updated for Inspire 2022
    3.70  Added MEC support
    3.71  Fixed MEC description & speaker parsing
    3.72  Fixed usage of format & subs arguments for direct YouTube downloads
    3.73  Added MEC slide deck support
          Fixed MEC parsing of description
    3.74  Fixed MEC processing of multi-line descriptions
    3.75  Added Ignite 2022 support
    3.76  Removed session code uniqueness when storing session data, as session data now can contain multiple entries per language using the same code
    3.77  Corrected API endpoints for some of the older events
    3.78  Fixed content-based help
    3.79  Fixed issue with placeholder detection
          Fixed path handling, fixes file detection and timestamping a.o.
          Added PowerShell 5.1 requirement (tested with)
    3.80  Fixed redundant passing of Format to YouTube-dl
    3.81  Moved to using ytl-dl, a fork of Youtube-DL (not maintained any longer)
    3.82  Fixed new folder creation
    3.83  Updated for Build 2023
          Removed Ignite 2018 and Ignite 2019
    3.9   Fixed retrieval of API-based catalogs for events
          Switched to using REST calls for those API-based catalogs
          Added Refresh Switch
          Removed archived events (<2021) as MS archives sessions selectively from previous years
          Merged Ignite2021H1 and Ignite2021H2 to Ignite2021
    3.91  Fixed output mentioning youtube-dl instead of actual tool (yt-dlp)
    3.92  Added .docx caption support for Build2023
    3.93  Fixed scraping streams from Azure Media Services for Build2023+
          Reinstated caption downloading with VTT instead of docx (can use Sub to download alt. language)
    3.94  Added ytp-dl's --concurrent-fragments support (default 4)
    3.95  Fixed localized VTT downloading for Build 2023+ from Azure Media Services
    3.96  Removed hidden character causing "Â : The term 'Â' is not recognized .." messages.

    .EXAMPLE
    Download all available contents of Ignite sessions containing the word 'Teams' in the title to D:\Ignite, and skip sessions from the CommunityTopic 'Fun and Wellness'
    .\Get-EventSession.ps1 -DownloadFolder D:\Ignite-Format 22 -Keyword 'Teams' -Event Ignite -ExcludecommunityTopic 'Fun and Wellness'

    .EXAMPLE
    Get information of all sessions, and output only location and time information for sessions (co-)presented by Tony Redmond:
    .\Get-EventSession.ps1 -InfoOnly | Where {$_.Speakers -contains 'Tony Redmond'} | Select Title, location, startDateTime

    .EXAMPLE
    Download all available contents of sessions BRK3248 and BRK3186 to D:\Ignite
    .\Get-EventSession.ps1 -DownloadFolder D:\Ignite -ScheduleCode BRK3248,BRK3186

    .EXAMPLE
    View all Exchange Server related sessions as Ignite including speakers(s), and sort them by date/time
    Get-EventSession.ps1 -Event Ignite -InfoOnly -Product '*Exchange Server*' | Sort-Object startDateTime | Select-Object @{n='Session'; e={$_.sessionCode}}, @{n='When';e={([datetime]$_.startDateTime).ToString('g')}}, title, @{n='Speakers'; e={$_.speakerNames -join ','}}

    .EXAMPLE
    Get all available sessions, display them in a GridView to select multiple at once, and download them to D:\Ignite
    .\Get-EventSession.ps1 -ScheduleCode (.\Get-EventSession.ps1 -InfoOnly | Out-GridView -Title 'Select Videos to Download, or Cancel for all Videos' -PassThru).SessionCode -MaxDownloadJobs 10 -DownloadFolder 'D:\Ignite'
#>
[cmdletbinding( DefaultParameterSetName = 'Default' )]
param(
    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'DownloadDirect')]
    [string]$DownloadFolder,

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [string]$Format,

    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [parameter( Mandatory = $false, ParameterSetName = 'DownloadDirect')]
    [string[]]$Keyword,

    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [parameter( Mandatory = $false, ParameterSetName = 'DownloadDirect')]
    [string[]]$Title,

    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [parameter( Mandatory = $false, ParameterSetName = 'DownloadDirect')]
    [string]$Speaker,

    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [parameter( Mandatory = $false, ParameterSetName = 'DownloadDirect')]
    [string]$Product,

    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [parameter( Mandatory = $false, ParameterSetName = 'DownloadDirect')]
    [string]$Category = '',

    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [parameter( Mandatory = $false, ParameterSetName = 'DownloadDirect')]
    [string]$SolutionArea,

    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [parameter( Mandatory = $false, ParameterSetName = 'DownloadDirect')]
    [string]$LearningPath,

    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [parameter( Mandatory = $false, ParameterSetName = 'DownloadDirect')]
    [string]$Topic,

    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [parameter( Mandatory = $false, ParameterSetName = 'DownloadDirect')]
    [string[]]$ScheduleCode,

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [parameter( Mandatory = $false, ParameterSetName = 'DownloadDirect')]
    [string[]]$ExcludecommunityTopic,

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [switch]$NoVideos,

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'DownloadDirect')]
    [switch]$NoSlidedecks,

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'DownloadDirect')]
    [string]$FFMPEG,

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'DownloadDirect')]
    [ValidateRange(1,128)] 
    [int]$MaxDownloadJobs=4,

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [parameter( Mandatory = $false, ParameterSetName = 'DownloadDirect')]
    [uri]$Proxy=$null,

    [parameter( Mandatory = $true, ParameterSetName = 'Download')]
    [parameter( Mandatory = $true, ParameterSetName = 'Default')]
    [parameter( Mandatory = $true, ParameterSetName = 'Info')]
    [parameter( Mandatory = $true, ParameterSetName = 'DownloadDirect')]
    [ValidateSet('MEC','MEC2022','Ignite', 'Ignite2022', 'Ignite2021', 'Inspire', 'Inspire2022', 'Inspire2021', 'Build', 'Build2023','Build2022', 'Build2021')]
    [string]$Event='',

    [parameter( Mandatory = $true, ParameterSetName = 'Info')]
    [switch]$InfoOnly,

    [parameter( Mandatory = $true, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'DownloadDirect')]
    [switch]$OGVPicker,

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'DownloadDirect')]
    [switch]$Overwrite,

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [parameter( Mandatory = $false, ParameterSetName = 'DownloadDirect')]
    [switch]$NoRepeats,

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'DownloadDirect')]
    [switch]$NoGuessing,

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'DownloadDirect')]
    [switch]$Timestamp,

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [string[]]$Subs,

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [string]$Language='English',

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [ValidateSet('de-DE','zh-CN','en-US','ja-JP','es-CO','fr-FR')]
    [string[]]$Locale='en-US',

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [parameter( Mandatory = $false, ParameterSetName = 'DownloadDirect')]
    [switch]$Refresh,

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'DownloadDirect')]
    [switch]$Captions,

    [parameter( Mandatory = $true, ParameterSetName = 'DownloadDirect')]
    [switch]$PreferDirect,

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'DownloadDirect')]
    $ConcurrentFragments= 4
)

    # Max age for cache, older than this # hours will force info refresh
    $MaxCacheAge = 24

    $YouTubeEXE = 'yt-dlp.exe'
    $YouTubeDL = Join-Path $PSScriptRoot $YouTubeEXE
    $FFMPEG= Join-Path $PSScriptRoot 'ffmpeg.exe'

    $YTlink = 'https://www.videohelp.com/download/yt-dlp.exe'
    $FFMPEGlink = 'https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-essentials.zip'

    # Fix 'Could not create SSL/TLS secure channel' issues with Invoke-WebRequest
    [Net.ServicePointManager]::SecurityProtocol = "tls12, tls11, tls" 

    $script:BackgroundDownloadJobs= @()

    Function Iif($Cond, $IfTrue, $IfFalse) {
        If( $Cond) { $IfTrue } Else { $IfFalse }
    }

    Function Fix-FileName ($title) {
        return (((((((($title -replace '\]', ')') -replace '\[', '(') -replace [char]0x202f, ' ') -replace '["\\/\?\*]', ' ') -replace ':', '-') -replace '  ', ' ') -replace '\?\?\?', '') -replace '\<|\>|:|"|/|\\|\||\?|\*', '').Trim()
    }

    Function Get-IEProxy {
        If ( (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings').ProxyEnable -ne 0) {
            $proxies = (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings').proxyServer
            if ($proxies) {
                if ($proxies -ilike "*=*") {
                    return $proxies -replace "=", "://" -split (';') | Select-Object -First 1
                }
                Else {
                    return ('http://{0}' -f $proxies)
                }
            }
            Else {
                return $null
            }
        }
        Else {
            return $null
        }
    }

    Function Test-ResolvedPath( $Path) {
        $null -ne (Get-ChildItem -LiteralPath $Path -ErrorAction SilentlyContinue)
    }

    Function Clean-VideoLeftovers ( $videofile) {
        $masks= '.*.mp4.part', '.*.mp4.ytdl'
	ForEach( $mask in $masks) {
            $FileMask= $videofile -replace '.mp4', $mask
            Get-Item -LiteralPath $FileMask -ErrorAction SilentlyContinue | ForEach-Object {
                Write-Verbose ('Removing leftover file {0}' -f $_.fullname)
	        Remove-Item -LiteralPath $_.fullname -Force -ErrorAction SilentlyContinue
            }
        }
    }

    Function Get-BackgroundDownloadJobs {
        $Temp= @()
        ForEach( $job in $script:BackgroundDownloadJobs) {

            switch( $job.Type) {
                1 {
                    $isJobRunning= $job.job.State -eq 'Running'
                }
                2 {
                    $isJobRunning= -not $job.job.hasExited
                }
                3 {
                    $isJobRunning= $job.job.State -eq 'Running'
                }
                default {
                    $isJobRunning= $false
                }
            }
            if( $isJobRunning) {
                $Temp+= $job
            }
            Else {
                # Job finished, process result
                switch( $job.Type) {
                    1 {
                        $isJobSuccess= $job.job.State -eq 'Completed'
                        $DeckInfo[ $InfoDownload]++
                    }
                    2 {
                        $isJobSuccess= Test-ResolvedPath -Path $job.file
                        $VideoInfo[ $InfoDownload]++
                        Write-Progress -Id $job.job.Id -Activity ('Video {0} {1}' -f $Job.scheduleCode, $Job.title) -Completed
                    }
                    3 {
                        $isJobSuccess= $job.job.State -eq 'Completed'
                    }
                    default {
                        $isJobSuccess= $false
                    }
                }

                # Test if file is placeholder
                $isPlaceholder= $false
		If( Test-ResolvedPath -Path $job.file) {
                    $FileObj= Get-ChildItem -LiteralPath $job.file
                    If( $FileObj.Length -eq 42) {

                        If( (Get-Content -LiteralPath $job.File) -eq 'No resource file is available for download') {
                            Write-Warning ('Removing {0} placeholder file {1}' -f $job.scheduleCode, $job.file)
                            Remove-Item -LiteralPath $job.file -Force
                            $isPlaceholder= $true

                            Switch( $job.Type) {
                                1 {
                                    # Placeholder Deck file downloaded
                                    $DeckInfo[ $InfoDownload]--
                                    $DeckInfo[ $InfoPlaceholder]++
                                }
                                2 {
                                    # Placeholder Video file downloaded
                                    $VideoInfo[ $InfoDownload]--
                                    $VideoInfo[ $InfoPlaceholder]++
                                }
                                3 {
                                    # Placeholder VTT file downloaded
                                }
                            }
                        }
                        Else {
                            # Placeholder different text?
                        }
                    }
                }

		If( $isJobSuccess -and -not $isPlaceholder) {

                    Write-Host ('Downloaded {0}' -f $job.file) -ForegroundColor Green

                    # Do we need to adjust timestamp
                    If( $job.Timestamp) {
                        #Set timestamp
                        Write-Verbose ('Applying timestamp {0} to {1}' -f $job.Timestamp, $job.file)
                        $FileObj= Get-ChildItem -LiteralPath $job.file
                        $FileObj.CreationTime= Get-Date -Date $job.Timestamp
                        $FileObj.LastWriteTime= Get-Date -Date $job.Timestamp
                    }

                    If( $job.Type -eq 2) {
                        # Clean video leftovers
                        Clean-VideoLeftovers $job.file
                    }
                }
                Else {
                    switch( $job.Type) {
                        1 {
                            Write-Host ('Problem downloading slidedeck {0} {1}' -f $job.scheduleCode, $job.title) -ForegroundColor Red
                            $job.job.ChildJobs | Stop-Job | Out-Null
                            $job.job | Stop-Job -PassThru | Remove-Job -Force | Out-Null
                        }
                        2 {
                            $LastLine= (Get-Content -LiteralPath $job.stdErrTempFile -ErrorAction SilentlyContinue) | Select-Object -Last 1
                            Write-Host ('Problem downloading video {0} {1}: {2}' -f $job.scheduleCode, $job.title, $LastLine) -ForegroundColor Red
                            Remove-Item -LiteralPath $job.stdOutTempFile, $job.stdErrTempFile -Force -ErrorAction Ignore
                        }
                        3 {
                            Write-Host ('Problem downloading captions {0} {1}' -f $job.scheduleCode, $job.title) -ForegroundColor Red
                            $job.job.ChildJobs | Stop-Job | Out-Null
                            $job.job | Stop-Job -PassThru | Remove-Job -Force | Out-Null
                        }
                        default {
                        }
                    }
                }
            }
        }
        $Num= ($Temp| Measure-Object).Count
        $script:BackgroundDownloadJobs= $Temp
        Show-BackgroundDownloadJobs 
        return $Num
    }

    Function Show-BackgroundDownloadJobs {
        $Num=0
        $NumDeck= 0
        $NumVid= 0
        $NumCaption= 0
        ForEach( $BGJob in $script:BackgroundDownloadJobs) {
            $Num++
            Switch( $BGJob.Type) {
                1 {
                     $NumDeck++
                }
                2 {
                     $NumVid++
                }
                3 {
                     $NumCaption++
                }
            }
        }
        Write-Progress -Id 2 -Activity 'Background Download Jobs' -Status ('Total {0} in progress ({1} slidedeck, {2} video and {3} caption files)' -f $Num, $NumDeck, $NumVid, $NumCaption)

        ForEach( $job in $script:BackgroundDownloadJobs) {
            If( $Job.Type -eq 2) {

                # Get last line of YT log to display for video downloads
                $LastLine= (Get-Content -LiteralPath $job.stdOutTempFile -ErrorAction SilentlyContinue) | Select-Object -Last 1
                If(!( $LastLine)) {
                    $LastLine= 'Evaluating..'
                }
                Write-Progress -Id $job.job.id -Activity ('Video {0} {1}' -f $job.scheduleCode, $Job.title) -Status $LastLine -ParentId 2
                $progressId++
            }
        }
    }

    Function Stop-BackgroundDownloadJobs {
        # Trigger update jobs running data
        $null= Get-BackgroundDownloadJobs
        # Stop all slidedeck background jobs
        ForEach( $BGJob in $script:BackgroundDownloadJobs ) { 
            Switch( $BGJob.Type) {
                1 {
                    $BGJob.Job.ChildJobs | Stop-Job -PassThru 
	            $BGJob.Job | Stop-Job -PassThru | Remove-Job -Force -ErrorAction SilentlyContinue
                }
                2 {
                    Stop-Process -Id $BGJob.job.id -Force -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 5
                    Remove-Item -LiteralPath $BGJob.stdOutTempFile, $BGJob.stdErrTempFile -Force -ErrorAction Ignore
                }
                3 {
                    $BGJob.Job.ChildJobs | Stop-Job -PassThru 
	            $BGJob.Job | Stop-Job -PassThru | Remove-Job -Force -ErrorAction SilentlyContinue
                }
            }
            Write-Warning ('Stopped downloading {0} {1}' -f $BGJob.scheduleCode, $BGJob.title) 
	    }
    }

    Function Add-BackgroundDownloadJob {
        param(
            $Type, 
            $FilePath,
            $DownloadUrl,
            $ArgumentList,
            $File,
            $Timestamp= $null,
            $Title='',
            $ScheduleCode=''
        )
        $JobsRunning= Get-BackgroundDownloadJobs
        If ( $JobsRunning -ge $MaxDownloadJobs) {
            Write-Host ('Maximum background download jobs reached ({0}), waiting for free slot - press Ctrl-C once to abort..' -f $JobsRunning)
            While ( $JobsRunning -ge $MaxDownloadJobs) {
                if ([system.console]::KeyAvailable) { 
                    Start-Sleep 1
                    $key = [system.console]::readkey($true)
                    if (($key.modifiers -band [consolemodifiers]"control") -and ($key.key -eq "C")) {
                        Write-Host "TERMINATING" -ForegroundColor Red
                        Stop-BackgroundDownloadJobs
                        Exit -1
                    }
                }
                $JobsRunning= Get-BackgroundDownloadJobs
            }
        }
        Switch( $Type) {
            1 {
                # Slidedeck
                $job= Start-Job -ScriptBlock { 
                    param( $url, $file) 
                    $wc = New-Object System.Net.WebClient
                    $wc.Encoding = [System.Text.Encoding]::UTF8
                    $wc.DownloadFile( $url, $file) 
                } -ArgumentList $DownloadUrl, $FilePath
                $stdOutTempFile = $null
                $stdErrTempFile = $null
            }
            2 {
                # Video
                $TempFile= Join-Path ($env:TEMP) (New-Guid).Guid
                $stdOutTempFile = '{0}-Out.log' -f $TempFile
                $stdErrTempFile = '{0}-Err.log' -f $TempFile
                $ProcessParam= @{
                    FilePath= $FilePath
                    ArgumentList= $ArgumentList
                    RedirectStandardError= $stdErrTempFile 
                    RedirectStandardOutput= $stdOutTempFile 
                    Wait= $false
                    Passthru= $true
                    NoNewWindow= $true
                    #WindowStyle= [System.Diagnostics.ProcessWindowStyle]::Normal
                }
                $job= Start-Process @ProcessParam
            }
            3 {
                # Caption
                $job= Start-Job -ScriptBlock { 
                    param( $url, $file) 
                    $wc = New-Object System.Net.WebClient
                    $wc.Encoding = [System.Text.Encoding]::UTF8
                    $wc.DownloadFile( $url, $file) 
                } -ArgumentList $DownloadUrl, $FilePath
                $stdOutTempFile = $null
                $stdErrTempFile = $null
            }
        }
        $object= New-Object -TypeName PSObject -Property @{
            Type= $Type
            job= $job
            file= $file
            title= $Title
            url= $DownloadUrl
            scheduleCode= $ScheduleCode
            timestamp= $timestamp
            stdOutTempFile= $stdOutTempFile
            stdErrTempFile= $stdErrTempFile
        }

        $script:BackgroundDownloadJobs+= $object
        Show-BackgroundDownloadJobs
    }

##########
# MAIN
##########
#Requires -Version 5.1

    If( $psISE) {
        Throw( 'Running from PowerShell ISE is not supported due to requirement to capture console input for proper termination of the script. Please run from a regular PowerShell session.')
    }

    If( $Proxy) {
        $ProxyURL= $Proxy
    }
    Else {
        $ProxyURL = Get-IEProxy
    }
    If ( $ProxyURL) {
        Write-Host "Using proxy address $ProxyURL"
    }
    Else {
        Write-Host "No proxy setting detected, using direct connection"
    }

    # Determine what event URLs to use. 
    # Use {0} for session code (eg BRK123), {1} for session id (guid)
    Switch( $Event) {
        {'MEC','MEC2022' -contains $_} {
            $EventName= 'MEC2022'
            $EventType='YT'
            $EventYTUrl= 'https://www.youtube.com/playlist?list=PLxdTT6-7g--2POisC5XcDQxUXHhWsoZc9'
            $EventLocale= 'en-us'
            $CaptionExt= 'vtt'
        }
        {'Ignite','Ignite2022' -contains $_} {
            $EventName= 'Ignite2022'
            $EventType='API'
            $EventAPIUrl= 'https://api.ignite.microsoft.com'
            $EventSearchURI= 'api/session/search'
            $SessionUrl= 'https://medius.studios.ms/Embed/video-nc/IG22-{0}'
            $CaptionURL= 'https://medius.studios.ms/video/asset/CAPTION/IG22-{0}'
            $SlidedeckUrl= 'https://medius.microsoft.com/video/asset/PPT/{0}'
            $Method= 'Post'
            # Note: to have literal accolades and not string formatter evaluate interior, use a pair:
            $EventSearchBody= '{{"itemsPerPage":{0},"searchFacets":{{"dateFacet":[{{"startDateTime":"2022-10-12T12:00:00.000Z","endDateTime":"2022-10-12T21:59:00.000Z"}}]}},"searchPage":{1},"searchText":"*","sortOption":"Chronological"}}'
            $CaptionExt= 'vtt'
        }
        {'Ignite2021' -contains $_} {
            $EventName= 'Ignite2021'
            $EventType='API'
            $EventAPIUrl= 'https://api.ignite.microsoft.com'
            $EventSearchURI= 'api/session/search'
            $SessionUrl= 'https://medius.studios.ms/Embed/video-nc/IG21-{0}'
            $CaptionURL= 'https://medius.studios.ms/video/asset/CAPTION/IG21-{0}'
            $SlidedeckUrl= 'https://medius.microsoft.com/video/asset/PPT/{0}'
            $Method= 'Post'
            $EventSearchBody= '{{"itemsPerPage":{0},"searchFacets":{{"dateFacet":[{{"startDateTime":"2021-11-01T08:00:00-05:00","endDateTime":"2021-11-30T19:00:00-05:00"}}]}},"searchPage":{1},"searchText":"*","sortOption":"Chronological"}}'
            $CaptionExt= 'vtt'
        }
        {'Inspire', 'Inspire2022' -contains $_} {
            $EventName= 'Inspire2022'
            $EventType='API'
            $EventAPIUrl= 'https://api.inspire.microsoft.com'
            $EventSearchURI= 'api/session/search'
            $SessionUrl= 'https://medius.studios.ms/video/asset/HIGHMP4/INSP22-{0}'
            $CaptionURL= 'https://medius.studios.ms/video/asset/CAPTION/INSP22-{0}'
            $SlidedeckUrl= 'https://medius.studios.ms/video/asset/PPT/INSP22-{0}'
            $Method= 'Post'
            $EventSearchBody= '{{"itemsPerPage":{0},"searchFacets":{{"dateFacet":[{{"startDateTime":"2022-07-19T08:00:00-05:00","endDateTime":"2022-07-20T19:00:00-05:00"}}]}},"searchPage":{1},"searchText":"*","sortOption":"Chronological"}}'
            $CaptionExt= 'vtt'
        }
        {'Inspire2021' -contains $_} {
            $EventName= 'Inspire2021'
            $EventType='API'
            $EventAPIUrl= 'https://api.inspire.microsoft.com'
            $EventSearchURI= 'api/session/search'
            $SessionUrl= 'https://medius.studios.ms/video/asset/HIGHMP4/INSP21-{0}'
            $CaptionURL= 'https://medius.studios.ms/video/asset/CAPTION/INSP21-{0}'
            $SlidedeckUrl= 'https://medius.studios.ms/video/asset/PPT/INSP21-{0}'
            $Method= 'Post'
            $EventSearchBody= '{{"itemsPerPage":{0},"searchFacets":{{"dateFacet":[{{"startDateTime":"2021-01-01T08:00:00-05:00","endDateTime":"2021-12-31T19:00:00-05:00"}}]}},"searchPage":{1},"searchText":"*","sortOption":"Chronological"}}'
            $CaptionExt= 'vtt'
        }
        {'Build', 'Build2023' -contains $_} {
            $EventName= 'Build2023'
            $EventType='API'
            $EventAPIUrl= 'https://api.build.microsoft.com'
            $EventSearchURI= 'api/session/search'
            $SessionUrl= 'https://medius.studios.ms/video/asset/HIGHMP4/B23-{0}'
            $CaptionURL= 'https://medius.studios.ms/video/asset/CAPTION/B23-{0}'
            $SlidedeckUrl= 'https://medius.studios.ms/video/asset/PPT/B23-{0}'
            $Method= 'Post'
            $EventSearchBody= '{{"itemsPerPage":{0},"searchFacets":{{"dateFacet":[{{"startDateTime":"2023-01-01T08:00:00-05:00","endDateTime":"2023-12-31T19:00:00-05:00"}}]}},"searchPage":{1},"searchText":"*","sortOption":"Chronological"}}'
            $CaptionExt= 'vtt'
        }
        {'Build2022' -contains $_} {
            $EventName= 'Build2022'
            $EventType='API'
            $EventAPIUrl= 'https://api.build.microsoft.com'
            $EventSearchURI= 'api/session/search'
            $SessionUrl= 'https://medius.studios.ms/video/asset/HIGHMP4/B22-{0}'
            $CaptionURL= 'https://medius.studios.ms/video/asset/CAPTION/B22-{0}'
            $SlidedeckUrl= 'https://medius.studios.ms/video/asset/PPT/B22-{0}'
            $Method= 'Post'
            $EventSearchBody= '{{"itemsPerPage":{0},"searchFacets":{{"dateFacet":[{{"startDateTime":"2022-01-01T08:00:00-05:00","endDateTime":"2022-12-31T19:00:00-05:00"}}]}},"searchPage":{1},"searchText":"*","sortOption":"Chronological"}}'
            $CaptionExt= 'vtt'
        }
        default {
            Write-Host ('Unknown event: {0}' -f $Event) -ForegroundColor Red
            Exit -1
        }
    }

    If (-not ($InfoOnly)) {

        # If no download folder set, use system drive with event subfolder
        If( -not( $DownloadFolder)) {
            $DownloadFolder= '{0}\{1}' -f $ENV:SystemDrive, $EventName
        }

        Add-Type -AssemblyName System.Web
        Write-Host "Using download path: $DownloadFolder"
        # Create the local content path if not exists
        if ( (Test-Path $DownloadFolder) -eq $false ) {
            New-Item -Path $DownloadFolder -ItemType Directory | Out-Null
        }

        If ( $NoVideos) {
            Write-Host 'Will skip downloading videos'
            $DownloadVideos = $false
        }
        Else {
            If (-not( Test-Path $YouTubeDL)) {
                Write-Host ('{0} not found, will try to download from {1}' -f $YouTubeEXE, $YTLink)
                Invoke-WebRequest -Uri $YTLink -OutFile $YouTubeDL -Proxy $ProxyURL
            }
            If ( Test-Path $YouTubeDL) {
                Write-Host ('Running self-update of {0}' -f $YouTubeEXE)

                $Arg = @('-U')
                If ( $ProxyURL) { $Arg += "--proxy $ProxyURL" }

                $pinfo = New-Object System.Diagnostics.ProcessStartInfo
                $pinfo.FileName = $YouTubeDL
                $pinfo.RedirectStandardError = $true
                $pinfo.RedirectStandardOutput = $true
                $pinfo.UseShellExecute = $false
                $pinfo.Arguments = $Arg
                Write-Verbose ('Running {0} using {1}' -f $pinfo.FileName, ($pinfo.Arguments -join ' '))
                $p = New-Object System.Diagnostics.Process
                $p.StartInfo = $pinfo
                $p.Start() | Out-Null
                $stdout = $p.StandardOutput.ReadToEnd()
                $stderr = $p.StandardError.ReadToEnd()
                $p.WaitForExit()

                If ($p.ExitCode -ne 0) {
                    If ( $stderr -contains 'Error launching') {
                        Throw ('Problem running {0}. Make sure this is an x86 system, and the required Visual C++ 2010 redistribution package is installed (available from https://www.microsoft.com/en-US/download/details.aspx?id=5555).' -f $YouTubeEXE)
                    }
                    Else {
                        Write-Host $stderr
                    }
                }
                Else {
                    Write-Host $stdout
                }
                $DownloadVideos = $true
            }
            Else {
                Write-Warning ('Unable to locate or download {0}, will skip downloading YouTube videos' -f $YouTubeEXE)
                $DownloadVideos = $false
            }

            If (-not( Test-Path $FFMPEG)) {

                Write-Host ('ffmpeg.exe not found, will try to download from {0}' -f $FFMPEGlink)
                $TempFile= Join-Path $PSScriptRoot 'ffmpeg-latest-win32-static.zip'
                Invoke-WebRequest -Uri $FFMPEGlink -OutFile $TempFile -Proxy $ProxyURL

                If( Test-Path $TempFile) {
                    Add-Type -AssemblyName System.IO.Compression.FileSystem
                    Write-Host ('{0} downloaded, trying to extract ffmpeg.exe' -f $TempFile)
                    $FFMPEGZip= [System.IO.Compression.ZipFile]::OpenRead( $TempFile)
                    $FFMPEGEntry= $FFMPEGZip.Entries | Where-Object {$_.FullName -like '*/ffmpeg.exe'}
                    If( $FFMPEGEntry) {
                        Try {
                            [System.IO.Compression.ZipFileExtensions]::ExtractToFile( $FFMPEGEntry, $FFMPEG)
                            $FFMPEGZip.Dispose()
                            Remove-Item -LiteralPath $TempFile -Force
                        }
                        Catch {
                            Throw ('Problem extracting ffmpeg.exe from {0}' -f $FFMPEGZip)
                        }
                    }
                    Else {
                        Throw 'ffmpeg.exe missing in downloaded archive'
                    }
                }
            }
            If ( Test-Path $FFMPEG) {
                Write-Host ('ffmpeg.exe located at {0}' -f $FFMPEG)
                $DownloadAMSVideos= $true
            }
            Else {
                Write-Warning 'Unable to locate or download ffmpeg.exe, will skip downloading Azure Media Services videos'
                $DownloadAMSVideos = $false
            }
        }
    }

    $SessionCache = Join-Path $PSScriptRoot ('{0}-Sessions.cache' -f $EventName)
    $SessionCacheValid = $false

    If( $Refresh) {
        Write-Host 'Refresh specified, will read session information from the online catalog'
    }
    Else {
        If ( Test-Path $SessionCache) {
            Try {
                If ( (Get-childItem -LiteralPath $SessionCache).LastWriteTime -ge (Get-Date).AddHours( - $MaxCacheAge)) {
                    Write-Host 'Session cache file found, reading session information'
                    $data = Import-CliXml -LiteralPath $SessionCache -ErrorAction SilentlyContinue
                    $SessionCacheValid = $true
                }
                Else {
                   Write-Warning 'Cache information expired, will re-read information from catalog'
                }
            }
            Catch {
                Write-Host 'Error reading cache file or cache file invalid - will read from online catalog' -ForegroundColor Red
            }
        }
    }

    If ( -not( $SessionCacheValid)) {

      Switch($EventType) {
        'API' {

            Write-Host ('Reading {0} session catalog' -f $EventName)
            $web = @{
                userAgent   = 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.103 Safari/537.36'
                requestUri  = [uri]('{0}/{1}' -f $EventAPIUrl, $EventSearchURI)
                headers     = @{'Content-Type'='application/json'}
                itemsPerPage= 100
            }

            Try {
                $SearchBody= $EventSearchBody -f '1', '1'
                Write-Verbose ('Using URI {0}' -f $web.requestUri)
                $searchResultsResponse = Invoke-RestMethod -Uri $web.requestUri -Body $searchbody -Method $Method -Headers $web.headers -UserAgent $web.userAgent -WebSession $session -Proxy $ProxyURL
                $searchResults= $searchResultsResponse.data
            }
            Catch {
                Throw ('Problem retrieving session catalog: {0}' -f $error[0])
            }
            [int32]$sessionCount = $searchResultsResponse.total
            [int32]$remainder = 0

            $PageCount = [System.Math]::DivRem($sessionCount, $web.itemsPerPage, [ref]$remainder)
            If ($remainder -gt 0) {
                $PageCount++
            }

            Write-Host ('Reading information for {0} sessions' -f $sessionCount)
            $data = [System.Collections.ArrayList]@()
            $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet', [string[]]('sessionCode', 'title'))
            $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
            For ($page = 1; $page -le $PageCount; $page++) {
                Write-Progress -Id 1 -Activity "Retrieving Session Catalog" -Status "Processing page $page of $PageCount" -PercentComplete ($page / $PageCount * 100)
                $SearchBody= $EventSearchBody -f $web.itemsPerPage, $page
                $searchResultsResponse = Invoke-RestMethod -Uri $web.requestUri -Body $searchbody -Method $Method -Headers $web.headers -UserAgent $web.userAgent -WebSession $session  -Proxy $ProxyURL
                ForEach ( $Item in $searchResultsResponse.data) {
                    $Item.PSObject.Properties | ForEach-Object {
                        if ($_.Name -eq 'speakerNames') { $Item.($_.Name) = @($_.Value) }
                        if ($_.Name -eq 'products') { $Item.($_.Name) = @($_.Value -replace [char]9, '/') }
                        if ($_.Name -eq 'contentCategory') { $Item.($_.Name) = @(($_.Value -replace [char]9, '/') -replace ' / ', '/') }
                    }
                    Write-Verbose ('Adding info for session {0}' -f $Item.sessionCode)
                    $Item.PSObject.TypeNames.Insert(0, 'Session.Information')
                    $Item | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                    $data.Add( $Item) | Out-Null
                }
            }
            Write-Progress -Id 1 -Completed -Activity "Finished retrieval of catalog" 
        }
            
        'YT' {
            # YouTube published - Use yt-dlp to download the playlist as JSON so we can parse it to 'expected format'
            Write-Host ('Reading {0} playlist information (might take a while) ..' -f $EventName)
            $data = [System.Collections.ArrayList]@()
            $Arg= [System.Collections.ArrayList]@()
            If ( $ProxyURL) { 
                $Arg.Add( '--proxy "{0}"' -f $ProxyURL) | Out-Null
            }
            $Arg.Add( '--socket-timeout 90') | Out-Null
            $Arg.Add( '--retries 15') | Out-Null
            $Arg.Add( '--dump-json') | Out-Null
            $Arg.Add( ('"{0}"' -f $EventYTUrl)) | Out-Null

            $pinfo = New-Object System.Diagnostics.ProcessStartInfo
            $pinfo.FileName = $YouTubeDL
            $pinfo.RedirectStandardError = $true
            $pinfo.RedirectStandardOutput = $true
            $pinfo.UseShellExecute = $false
            $pinfo.Arguments = $Arg
            Write-Verbose ('Running {0} using {1}' -f $pinfo.FileName, ($pinfo.Arguments -join ' '))

            $p = New-Object System.Diagnostics.Process
            $p.StartInfo = $pinfo
            $p.Start() | Out-Null
            $stdout = $p.StandardOutput.ReadToEnd()
            $stderr = $p.StandardError.ReadToEnd()
            $p.WaitForExit()

            If ($p.ExitCode -ne 0) {
                Throw ('Problem running {0}: {1}' -f $YouTubeEXE, $stderr)
            }

            Try {
                Write-Verbose ('Converting from Json ..')
                # Trim any trailing empty lines, convert single string with line-breaks to array for JSON conversion
                $JsonData= ($stdout.Trim([System.Environment]::Newline) -Split "`n") | ConvertFrom-Json
            }
            Catch {
                Throw( 'Output does not seem to be proper JSON format, see {0}' -f $TempJsonFile)
            }

            ForEach( $Item in $JsonData) {

                $SpeakerNames= [System.Collections.ArrayList]@()

                # Description match pattern? Set Desc+Speakers, otherwise Desc=Description, assume no Speakers defined

                If($Item.Description -match '^(?<Description>[\s\S]*?)(\s)*(Download the slide deck from (?<Slidedeck>https:\/\/.*?)[\.]?)?(\s)*(Speaker(s)?:(\s)?(?<Speakers>.*))?(\s)*$') {
                    $Description= $Matches.Description
                    $Matches.Speakers -Split ';' | ForEach-Object { $SpeakerNames.Add( $_.Trim() ) |Out-Null }
                    $SlidedeckUrl= $Matches.Slidedeck
                }
                Else {
                    $Description= $Item.Description
                    $SlidedeckUrl= $null
                }

                # Slidedeck url, construct real link:
                If( $SlidedeckUrl) {
                    # https://www.microsoft.com/en-us/download/details.aspx?id=104608 -> https://www.microsoft.com/en-us/download/confirmation.aspx?id=104608

                    If( $SlidedeckUrl -match '^(?<host>https:\/\/www\.microsoft\.com).*id=(?<id>[\d]+)$') {
                        $SlideDeck= '{0}/en-us/download/confirmation.aspx?id={1}' -f $Matches.host, $Matches.id
                    }
                    Else {
                        Write-Warning ('Unexpected slide deck URL format: {0}' -f $SlidedeckUrl)
                        $Slidedeck= $null
                    }
                }
                Else {
                    $SlideDeck= $null
                }

                $object = [PSCustomObject]@{
                    sessionCode= [string]('{0:d2}' -f $Item.playlist_index)
                    SessionType= 'On-Demand'
                    Title= $Item.Title
                    Description= $Description
                    onDemand= $Item.webpage_url
                    Views= $Item.view_count
                    Likes= $Item.like_count
                    Duration= [timespan]::FromSeconds( $Item.duration).ToString()
                    langLocale= $EventLocale
                    SolutionArea= $Item.Tags
                    contentCategory= $Item.categories
                    SpeakerNames= $SpeakerNames
                    Slidedeck= $Slidedeck
                    startDateTime= [Datetime]::ParseExact( $Item.upload_date, 'yyyyMMdd', $null)
                    onDemandThumbnail= ($Item.thumbnails | Sort-Object -Property Id | Select-Object -First 1).Url
                }
                Write-Verbose ('Adding info for session {0}' -f $Item.Title)
                $data.Add( $object) | Out-Null
            }
        }

        default {
          Throw( 'Unknown event catalog type {0}' -f $EventType)
        }
      }

      Write-Host 'Storing session information'
      $data | Export-CliXml -Encoding Unicode -Force -LiteralPath $SessionCache

    }

    $SessionsToGet = $data
    $TotalNumberOfSessions= ($SessionsToGet | Measure-Object).Count

    If ($scheduleCode) {
        Write-Verbose ('Session code(s) specified: {0}' -f ($ScheduleCode -join ','))
        $SessionsToGet = $SessionsToGet | Where-Object { $scheduleCode -contains $_.sessioncode }
    }

    If ($ExcludeCommunityTopic) {
        Write-Verbose ('Excluding community topic: {0}' -f $ExcludeCommunityTopic)
        $SessionsToGet = $SessionsToGet | Where-Object { $ExcludeCommunityTopic -inotcontains $_.CommunityTopic  }
    }

    If ($Speaker) {
        Write-Verbose ('Speaker keyword specified: {0}' -f $Speaker)
        $SessionsToGet = $SessionsToGet | Where-Object { $_.speakerNames | Where-Object {$_ -ilike $Speaker} }
    }

    If ($Product) {
        Write-Verbose ('Product specified: {0}' -f $Product)
        $SessionsToGet = $SessionsToGet | Where-Object { $_.products | Where-Object {$_ -ilike $Product }}
    }

    If ($Category) {
        Write-Verbose ('Category specified: {0}' -f $Category)
        $SessionsToGet = $SessionsToGet | Where-Object { $_.contentCategory | Where-Object {$_ -ilike $Category }}
    }

    If ($SolutionArea) {
        Write-Verbose ('SolutionArea specified: {0}' -f $SolutionArea)
        $SessionsToGet = $SessionsToGet | Where-Object { $_.solutionArea | Where-Object {$_ -ilike $SolutionArea }}
    }

    If ($LearningPath) {
        Write-Verbose ('LearningPath specified: {0}' -f $LearningPath)
        $SessionsToGet = $SessionsToGet | Where-Object { $_.learningPath | Where-Object {$_ -ilike $LearningPath }}
    }

    If ($Topic) {
        Write-Verbose ('Topic specified: {0}' -f $Topic)
        $SessionsToGet = $SessionsToGet | Where-Object { $_.topic | Where-Object {$_ -ilike $Topic }}
    }

    If ($Locale) {
        Write-Verbose ('Locale(s) specified: {0}' -f ($Locale -join ','))
        Write-Verbose ('Sessions Pre: {0}'  -f ($SessionsToGet.Count))
        $SessionsToGetTemp= [System.Collections.ArrayList]@()
        ForEach( $item in $Locale) {
            $SessionsToGet | Where-Object {$item -ieq $_.langLocale} | ForEach-Object { $null= $SessionsToGetTemp.Add(  $_ ) }
        }
        $SessionsToGet= $SessionsToGetTemp | Sort-Object -Unique -Property sessionCode
        Write-Verbose ('Sessions Post: {0}'  -f ($SessionsToGet.Count))
    }

    If ($Title) {
        Write-Verbose ('Title keyword(s) specified: {0}' -f ( $Title -join ','))
        $SessionsToGetTemp= [System.Collections.ArrayList]@()
        ForEach( $item in $Title) {
            $SessionsToGet | Where-Object {$_.title -ilike ('*{0}*' -f $item) } | ForEach-Object { $null= $SessionsToGetTemp.Add(  $_ ) }
        }
        $SessionsToGet= $SessionsToGetTemp | Sort-Object -Unique -Property sessionCode
    }

    If ($Keyword) {
        Write-Verbose ('Description keyword(s) specified: {0}' -f ( $Keyword -join ','))
        $SessionsToGetTemp= [System.Collections.ArrayList]@()
        ForEach( $item in $Keyword) {
            $SessionsToGet | Where-Object {$_.description -ilike ('*{0}*' -f $item) } | ForEach-Object { $null= $SessionsToGetTemp.Add(  $_ ) }
        }
        $SessionsToGet= $SessionsToGetTemp | Sort-Object -Unique -Property sessionCode
    }

    If ($NoRepeats) {
        Write-Verbose ('Skipping repeated sessions')
        $SessionsToGet = $SessionsToGet | Where-Object {$_.sessionCode -inotmatch '^*R[1-9]?$' -and $_.sessionCode -inotmatch '^[A-Z]+[0-9]+[B-C]+$'}
    }

    If ( $InfoOnly) {
        Write-Verbose ('There are {0} sessions matching your criteria.' -f (($SessionsToGet | Measure-Object).Count))
        Write-Output $SessionsToGet
    }
    Else {

        If( $OGVPicker) {
            $SessionsToGet= $SessionsToGet | Out-GridView -Title 'Select Videos to Download, or Cancel for all Videos' -PassThru
        }

        $i = 0
        $DeckInfo = @(0, 0, 0)
        $VideoInfo = @(0, 0, 0)
        $InfoDownload = 0
        $InfoPlaceholder = 1
        $InfoExist = 2

        $myTimeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById( 'US Eastern Standard Time')

        [console]::TreatControlCAsInput = $true

        $SessionsSelected = ($SessionsToGet | Measure-Object).Count
        Write-Host ('There are {0} sessions matching your criteria.' -f $SessionsSelected)
        Foreach ($SessionToGet in $SessionsToGet) {

            $i++
            Write-Progress -Id 1 -Activity 'Inspecting session information' -Status "Processing session $i of $SessionsSelected" -PercentComplete ($i / $SessionsSelected * 100)
            If( $SessionToGet.sessionCode) {
                $FileName = Fix-FileName ('{0}-{1}' -f $SessionToGet.sessionCode.Trim(), $SessionToGet.title.Trim())
            }
            Else {
                $FileName = Fix-FileName ('{0}' -f $SessionToGet.title.Trim())
            }
            If(! ([string]::IsNullOrEmpty( $SessionToGet.startDateTime))) {
                # Get session localized timestamp, undoing TZ adjustments
                $SessionTime= [System.TimeZoneInfo]::ConvertTime((Get-Date -Date $SessionToGet.startDateTime).ToUniversalTime(), $myTimeZone ).toString('g')
            }
            Else {
                $SessionTime= $null
            }
            Write-Host ('Processing info session {0} from {1} [{2}]' -f $FileName, (Iif -Cond $SessionTime -IfTrue $SessionTime -IfFalse 'No Timestamp'), $SessionToGet.langLocale)
            If(!([string]::IsNullOrEmpty( $SessionToGet.startDateTime)) -and (Get-Date -Date $SessionToGet.startDateTime) -ge (Get-Date)) {
                Write-Verbose ('Skipping session {0}: Didn''t take place yet' -f $SessionToGet.sessioncode)
            }
            Else {

              If( ! $NoVideos) {
                If ( $DownloadVideos -or $DownloadAMSVideos) {

                    $vidfileName = ("$FileName.mp4")
                    $vidFullFile = '\\?\{0}' -f (Join-Path $DownloadFolder $vidfileName)

                    if ((Test-ResolvedPath -Path $vidFullFile) -and -not $Overwrite) {
                        Write-Host ('Video exists {0}' -f $vidfileName) -ForegroundColor Gray
                        $VideoInfo[ $InfoExist]++
                        # Clean video leftovers
                        Clean-VideoLeftovers $vidFullFile
                    }
                    else {
                        $downloadLink= $null
                        If ( !( [string]::IsNullOrEmpty( $SessionToGet.onDemand)) ) {
                            If( $PreferDirect -and !( [string]::IsNullOrEmpty( $SessionToGet.downloadVideoLink))) {
                                $downloadLink = $SessionToGet.downloadVideoLink
                            }
                            Else {
                                $downloadLink = $SessionToGet.onDemand
                            }
                        }
                        Else {
                            If (!( [string]::IsNullOrEmpty( $SessionToGet.downloadVideoLink)) ) {
                                $downloadLink = $SessionToGet.downloadVideoLink
                            }
                            Else {
                                If( $NoGuessing) {
                                    $downloadLink= $null
                                }
                                Else { 
                                    # Try session page, eg https://medius.studios.ms/Embed/Video/IG18-BRK2094
                                    $downloadLink = $SessionUrl -f $SessionToGet.SessionCode
                                }
                            }
                        }
                        If( $downloadLink -match '(medius\.studios\.ms\/Embed\/Video|medius\.microsoft\.com|mediastream\.microsoft\.com)' ) {
                            Write-Verbose ('Checking hosted video link {0}' -f $downloadLink)
                            Try {
                                $DownloadedPage= Invoke-WebRequest -Uri $downloadLink -Proxy $ProxyURL -DisableKeepAlive -ErrorAction SilentlyContinue
                            }
                            Catch {
                                Write-Warning ('Problem downloading from {0}' -f $downloadLink)
                            }
                            If( $DownloadedPage) {                        
                                $OnDemandPage= $DownloadedPage.RawContent 
                                
                                If( $OnDemandPage -match 'StreamUrl = "(?<AzureStreamURL>https://mediusprod\.streaming\.mediaservices\.windows\.net/.*manifest)";') {

                                    Write-Verbose ('Using Azure Media Services URL {0}' -f $matches.AzureStreamURL)
                                    $Endpoint= '{0}(format=mpd-time-csf)' -f $matches.AzureStreamURL
                                    $Arg = @( ('-o "{0}"' -f ($vidFullFile -replace '%', '%%')), $Endpoint)

                                    # Construct Format for this specific video, language and audio languages available
                                    If ( $Format) {
                                        $ThisFormat= $Format
                                    } 
                                    Else { 
                                        $ThisFormat= 'worstvideo+bestaudio'
                                    }
 
                                    If( $SessionToGet.audioLanguage) {

                                        If( $SessionToGet.audioLanguage.Count -gt 1) {
                                            # Session has multiple audio tracks
                                            If( $SessionToGet.audioLanguage -icontains $Language) {
                                                Write-Warning ('Multiple audio languages available; will try downloading {0} audio stream' -f $Language)
                                                $ThisLanguage= $Language
                                            }
                                            Else {
                                                $ThisLanguage= $SessionToGet.audioLanguage | Select -First 1
                                                Write-Warning ('Requested language {0} not available; will use default stream ({1})' -f $Language, $ThisLanguage)
                                            }

                                            # Take specified Format apart so we can insert the language filter per specification
                                            $ThisFormatElem= $ThisFormat -Split ','
                                            $NewFormat= [System.Collections.ArrayList]@()
                                            ForEach( $Elem in $ThisFormatElem) {
                                                If( $Elem -match '^(?<pre>.*audio)(\[(?<audioparam>.*)\])?(?<post>(.*)?)$' ) {
                                                    If( $matches.audioparam) {
                                                        $NewFormatElem= '{0}[format_id*={1},{2}]{3}' -f $matches.Pre, $ThisLanguage, $matches.audioparam, $matches.post
                                                    }
                                                    Else {
                                                        $NewFormatElem= '{0}[format_id*={1}]{2}' -f $matches.Pre, $ThisLanguage, $matches.post
                                                    }
                                                }
                                                Else {
                                                    $NewFormatElem= $Elem
                                                    Write-Warning ('Problem determining where to add language criteria in {0}, leaving criteria as-is' -f $NewFormat)
                                                }
                                                $null= $NewFormat.Add( $NewFormatElem)
                                            }

                                            # With language filters determined, recreate filter and add whole non-language specific qualifiers as next best 
                                            $ThisFormat= ($NewFormat -Join ','), $ThisFormat -Join ','

                                        }
                                        Else {
                                            # Only 1 Language available, so use default audio stream
                                            Write-Warning ('Only single audio stream available, will use default audio stream')
                                        }
                                    }
                                    Else {
                                        # No multiple audio languages, use default audio stream
                                        Write-Warning ('Multiple audio streams not available, will use default audio stream')
                                    }
                                    $Arg += ('--format {0}' -f $ThisFormat)
                                }
                                Else {
                                    # Check for embedded YouTube 
                                    If( $OnDemandPage -match '"https:\/\/www\.youtube-nocookie\.com\/embed\/(?<YouTubeID>.+?)\?enablejsapi=1&"') {
                                        $Endpoint= 'https://www.youtube.com/watch?v={0}' -f $matches.YouTubeID
                                        Write-Verbose ('Using YouTube URL {0}' -f $Endpoint)
                                        $Arg = @( ('-o "{0}"' -f ($vidFullFile -replace '%', '%%')), $Endpoint)
                                        $Arg += ('--concurrent-fragments {0}' -f $ConcurrentFragments)
                                        If ( $Format) { $Arg += ('--format {0}' -f $Format) } Else { $Arg += ('--format 22') }
                                        If ( $Subs) { $Arg += ('--sub-lang {0}' -f ($Subs -Join ',')), ('--write-sub'), ('--write-auto-sub'), ('--convert-subs srt') }
                                    }
                                    Else {
                                        Write-Warning "Skipping: Embedded AMS or YouTube URL not found"
                                        $EndPoint= $null
                                    }
                                }

                            }
                            Else {
                                 Write-Warning ('Skipping: {0} unavailable' -f $downloadLink)
                            }
                        }
                        Else {
                            # Direct
                            Write-Verbose ('Using direct video link {0}' -f $downloadLink)
                            If( $downloadLink) {
                                $Endpoint= $downloadLink
                                $Arg = @( ('-o "{0}"' -f $vidFullFile), $downloadLink)
                            }
                            Else {
                                Write-Warning ('No video link for {0}' -f ($SessionToGet.Title))
                                $Endpoint= $null
                            }
                        }
                        If( $Endpoint) {
                            # Direct, AMS or YT video found, attempt download but first define common parameters

                            If ( $ProxyURL) { 
                                $Arg += ('--proxy "{0}"' -f $ProxyURL)
                            }
                            $Arg+= '--socket-timeout 90'
                            $Arg+= '--no-check-certificate'                            
                            $Arg+= '--retries 15'
                            $Arg+= '--concurrent-fragments {0}' -f $ConcurrentFragments

                            If ( $Subs) { $Arg += ('--sub-lang {0}' -f ($Subs -Join ',')), ('--write-sub'), ('--write-auto-sub'), ('--convert-subs srt') }

                            Write-Verbose ('Running: {0} {1}' -f $YouTubeEXE, ($Arg -join ' '))
                            Add-BackgroundDownloadJob -Type 2 -FilePath $YouTubeDL -ArgumentList $Arg -File $vidFullFile -Timestamp $SessionTime -scheduleCode ($SessionToGet.sessioncode) -Title ($SessionToGet.Title)
                        }
                        Else {
                            # Video not available or no link found
                            $VideoInfo[ $InfoPlaceholder]++
                        }
                    }

                    If( $Captions) {
                        $captionExtFile= $vidFullFile -replace '.mp4', ('.{0}' -f $CaptionExt)

                        If ((Test-ResolvedPath -Path $captionExtFile) -and -not $Overwrite) {
                            Write-Host ('Caption file exists {0}' -f $captionExtFile) -ForegroundColor Gray
                        }
                        Else {
                            # Caption file in AMS needs seperate download, fetch onDemand page if not already downloaded for video
                            If(! $OnDemandPage) {
                                If( $SessionToGet.onDemand) {
                                    Try {
                            		Write-Host ('Fetching video page to retrieve transcript information from {0}' -f $SessionToGet.onDemand) 
                                        $DownloadedPage= Invoke-WebRequest -Uri $SessionToGet.onDemand -Proxy $ProxyURL -DisableKeepAlive -ErrorAction SilentlyContinue
                                        If( $DownloadedPage) {                        
                                            $OnDemandPage= $DownloadedPage.RawContent 
                                        }
                                    }
                                    Catch {
                                        #Problem retrieving file, look for alternative options
                                    }
                                }
                              
                            }
                            # Check for vtt files before we check any direct caption file (likely docx now)
                            If( $OnDemandPage -match 'captionsConfiguration = (?<CaptionsJSON>{.*});') {
                                $CaptionConfig= ($matches.CaptionsJSON | ConvertFrom-Json).languageList
                                If( $Subs) {
                                    $captionFileLink= ($CaptionConfig | Where-Object {$_.srclang -eq $Subs}).src
                                }
                                If(! $captionFileLink) {
                                    $captionFileLink= ($CaptionConfig | Where-Object {$_.srclang -eq 'en'}).src
                                }
                            }
                            If( ! $CaptionFileLink) {
                                $captionFileLink= $SessionToGet.captionFileLink
                            }
                            If( ! $captionFileLink) {

                                If(! $OnDemandPage) {
                                    # Try if there is caption file reference on page
                                    Try {
                                        $DownloadedPage= Invoke-WebRequest -Uri $downloadLink -Proxy $ProxyURL -DisableKeepAlive -ErrorAction SilentlyContinue
                                        $OnDemandPage= $DownloadedPage.RawContent
                                    }
                                    Catch {
                                        $DownloadedPage= $null
                                        $onDemandPage= $null
                                    } 
                                }
                                Else {
                                    # Reuse one from video download
                                }

                                If( $OnDemandPage -match '"(?<AzureCaptionURL>https:\/\/mediusprodstatic\.studios\.ms\/asset-[a-z0-9\-]+\/transcript\{0}\?.*?)"' -f $CaptionExt) {
                                    $captionFileLink= $matches.AzureCaptionURL
                                }
                                If( ! $captionFileLink) {
                                    $captionFileLink= $captionURL -f $SessionToGet.SessionCode
                                }
                            }
                            If( $captionFileLink) {
                                Write-Verbose ('Retrieving caption file from URL {0}' -f $captionFileLink)

                                 $captionFullFile= $captionExtFile
                                 Write-Verbose ('Downloading {0} to {1}' -f $captionFileLink,  $captionFullFile)
                                 Add-BackgroundDownloadJob -Type 3 -FilePath $captionExtFile -DownloadUrl $captionFileLink -File $captionFullFile -Timestamp $SessionTime -scheduleCode ($SessionToGet.sessioncode) -Title ($SessionToGet.Title)

                             }
                             Else {
                                 Write-Warning "Subtitles requested, but no Caption URL found"
                             }
                        }
                    }
                    $OnDemandPage= $null
                }
              }

              If(! $NoSlidedecks) {
                If ( !( [string]::IsNullOrEmpty( $SessionToGet.slideDeck)) ) {
                    $downloadLink = $SessionToGet.slideDeck
                }
                Else {
                    If( $NoGuessing) {
                        $downloadLink= $null
                    }
                    Else {
                        # Try alternative construction
                        $downloadLink = $SlidedeckUrl -f $SessionToGet.SessionCode
                    }
                }

                If ($downloadLink -match "view.officeapps.live.com.*PPTX" -or $downloadLink -match 'downloaddocument' -or $downloadLink -match 'medius' -or $downloadLink -match 'confirmation\.aspx') {

                    $DownloadURL = [System.Web.HttpUtility]::UrlDecode( $downloadLink )
                    Try {
                       If( $downloadLink -notmatch 'confirmation\.aspx') {
                           $ValidUrl= Invoke-WebRequest -Uri $DownloadURL -Method HEAD -UseBasicParsing -DisableKeepAlive -MaximumRedirection 10 -ErrorAction SilentlyContinue
                       }
                       Else {
                           $ValidUrl= Invoke-WebRequest -Uri $DownloadURL -Method GET -UseBasicParsing -DisableKeepAlive -MaximumRedirection 10 -ErrorAction SilentlyContinue
                       }
                    }
                    Catch {
                        $ValidUrl= $false
                    }

                    If( $downloadLink -match 'confirmation\.aspx' -and $ValidURL.Headers.'Content-Type' -ilike 'text/html') {
                        # Extra parsing for MS downloads
                        If( $ValidUrl.RawContent -match 'href="(?<Url>https:\/\/download\.microsoft\.com\/download[\/0-9\-]*\/.*(pdf|pptx))".*click here to download manually') {
                            $DownloadURL= [System.Web.HttpUtility]::UrlDecode( $Matches.Url)
                            $ValidUrl= Invoke-WebRequest -Uri $DownloadURL -Method HEAD -UseBasicParsing -DisableKeepAlive -MaximumRedirection 10 -ErrorAction SilentlyContinue
                        }
                    }

                    If( $ValidUrl ) {
                        If( $DownloadURL -like '*.pdf' -or $ValidURL.Headers.'Content-Type' -ieq 'application/pdf') {
                            # Slidedeck offered is PDF format
                            $slidedeckFile = '{0}.pdf' -f $FileName
                        }
                        Else {
                            $slidedeckFile = '{0}.pptx' -f $FileName
                        }
                        $slidedeckFullFile =  '\\?\{0}' -f (Join-Path $DownloadFolder $slidedeckFile)
                        if ((Test-ResolvedPath -Path $slidedeckFullFile) -and ((Get-ChildItem -LiteralPath $slidedeckFullFile -ErrorAction SilentlyContinue).Length -gt 0) -and -not $Overwrite) {
                            Write-Host ('Slidedeck exists {0}' -f $slidedeckFile) -ForegroundColor Gray 
                            $DeckInfo[ $InfoExist]++
                        }
                        Else {
                            Write-Verbose ('Downloading {0} to {1}' -f $DownloadURL,  $slidedeckFullFile)
                            Add-BackgroundDownloadJob -Type 1 -FilePath $slidedeckFullFile -DownloadUrl $DownloadURL -File $slidedeckFullFile -Timestamp $SessionTime -scheduleCode ($SessionToGet.sessioncode) -Title ($SessionToGet.Title)
                        }
                    }
                    Else {
                        Write-Warning ('Skipping: Slidedeck unavailable {0}' -f $DownloadURL)
                        $DeckInfo[ $InfoPlaceholder]++
                    }
                }
                Else {
                    Write-Warning ('No slidedeck link for {0}' -f ($SessionToGet.Title))
                }
              }
            }

            $JobsRunning= Get-BackgroundDownloadJobs

            if ([system.console]::KeyAvailable) { 
                $key = [system.console]::readkey($true)
                if (($key.modifiers -band [consolemodifiers]"control") -and ($key.key -eq "C")) {
                    Write-Host "TERMINATING" -ForegroundColor Red
                    Stop-BackgroundDownloadJobs
                    Exit -1
                }
            }
                   
        }

        $ProcessedSessions= $i

        Write-Progress -Id 1 -Completed -Activity "Finished processing session information"

        $JobsRunning= Get-BackgroundDownloadJobs
        If ( $JobsRunning -gt 0) {
            Write-Host ('Waiting for download jobs to finish - press Ctrl-C once to abort)' -f $JobsRunning)
            While  ( $JobsRunning -gt 0) {
                if ([system.console]::KeyAvailable) { 
                    Start-Sleep 1
                    $key = [system.console]::readkey($true)
                    if (($key.modifiers -band [consolemodifiers]"control") -and ($key.key -eq "C")) {
                        Write-Host "TERMINATING" -ForegroundColor Red
                        Stop-BackgroundDownloadJobs
                        Exit -1
                    }
                }
                Start-Sleep 5
                $JobsRunning= Get-BackgroundDownloadJobs
            }
        }
        Else {
            Write-Host ('Background download jobs have finished' -f $JobsRunning)
        }

        Write-Progress -Id 2 -Completed -Activity "Download jobs finished"  

        Write-Host ('Selected {0} sessions out of a total of {1}' -f $ProcessedSessions, $TotalNumberOfSessions)
        Write-Host ('Downloaded {0} slide decks and {1} videos.' -f $DeckInfo[ $InfoDownload], $VideoInfo[ $InfoDownload])
        Write-Host ('Not (yet) available: {0} slide decks and {1} videos' -f $DeckInfo[ $InfoPlaceholder], $VideoInfo[ $InfoPlaceholder])
        Write-Host ('Skipped {0} slide decks and {1} videos as they were already downloaded.' -f $DeckInfo[ $InfoExist], $VideoInfo[ $InfoExist])
    }
