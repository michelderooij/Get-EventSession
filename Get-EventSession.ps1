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

    Michel de Rooij
    http://eightwone.com
    Version 4.37, December 30, 2025

    Special thanks to: Mattias Fors, Scott Ladewig, Tim Pringle, Andy Race, Richard van Nieuwenhuizen

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
    prevent this, usage of filters is recommended. For Azure Streams, you can use the following format:
    478        mp4 320x180     30 │ ~237.16MiB  478k m3u8  │ avc3.4d4016  478k video only
    628        mp4 384x216     30 │ ~311.58MiB  628k m3u8  │ avc3.4d401e  628k video only
    928        mp4 512x288     30 │ ~460.43MiB  928k m3u8  │ avc3.4d4020  928k video only
    1428       mp4 640x360     30 │ ~708.50MiB 1428k m3u8  │ avc3.4d4020 1428k video only
    2128       mp4 960x540     30 │ ~  1.03GiB 2128k m3u8  │ avc3.4d4020 2128k video only
    3128       mp4 1280x720    30 │ ~  1.52GiB 3128k m3u8  │ avc3.640029 3128k video only
    6128       mp4 1920x1080   30 │ ~  2.97GiB 6128k m3u8  │ avc3.64002a 6128k video only

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
    - Ignite                                       : Ignite contents (current)
    - Ignite2025,Ignite2024,Ignite2023             : Ignite contents from that year/time
    - Inspire                                      : Inspire contents (current)
    - Inspire2023                                  : Inspire contents from that year
    - Build                                        : Build contents (current)
    - Build2025,Build2024,Build2023                : Build contents from that year
    - MEC                                          : MEC contents (current)
    - MEC2022                                      : MEC contents from that year

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
    where AMS may offer multiple video qualities. Direct downloads ignores Format parameter.

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

    .PARAMETER ConcurrentFragments
    Specifies the number of fragments for yt-dlp to download simultaneously when downloading videos.
    Default is 4.

    .PARAMETER TempPath
    This will allow you to specify a folder for yt-dlp to store temporary files in, eg fragments.
    When omitted, the folder where the videos are saved to will be used.

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
          Fixed path handling, fixes file detection and timestamping a.o.7
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
    3.97  Added Inspire 2023
    3.98  Fixed retrieval of Inspire 2023 catalog
    3.99  Fixed reporting of element when we cannot add language filter
    4.00  Updated yt-dlp download location
          Changed checking yt-dlp.exe presence & validity
    4.01  Updated Event parameter help
    4.02  Added Ignite 2023
    4.10  Added Build 2024
    4.11  Fixed bug in downloading captions
    4.20  Added Ignite 2024
    4.21  Fixed date-range for Ignite 2024 ao
    4.22  Fixed download locations for Ignite 2024 content
          Added Azure Stream format guidance
    4.23  Added TempPath parameter to specifying yt-dlp temporary files location
          Fixed overwrite mode when calling yt-dlp
          Added parameter description for ConcurrentFragments
          Fixed reporting of failed downloads
          Some minor code cleanup
    4.3   Added Build 2025
          Rewrite for new catalog API endpoint and session hosting
    4.31  Fixed downloading Captions for direct video links
    4.32  Fixed default format when downloading from YouTube
    4.33  Added Ignite 2025
          Removed 2021 and 2022 event options
    4.34  Added removal of Ignite2025 placeholder files
    4.35  Setting video output preset to mp4 to make sure merging results in mp4 file, not an mkv
          Added header to output
    4.36  Fixed downloading of direct video links for MP2T type (YouTube)
          Added Cookies and CookiesFromBrowser support for yt-dlp (YouTube)

    TODO:
    - Add processing of archived events through new API endpoint (starting with Build)

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
    [parameter( Mandatory = $false, ParameterSetName = 'DownloadDirect')]
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
    [ValidateRange(1, 128)]
    [int]$MaxDownloadJobs = 4,

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [parameter( Mandatory = $false, ParameterSetName = 'DownloadDirect')]
    [uri]$Proxy = $null,

    [parameter( Mandatory = $true, ParameterSetName = 'Download')]
    [parameter( Mandatory = $true, ParameterSetName = 'Default')]
    [parameter( Mandatory = $true, ParameterSetName = 'Info')]
    [parameter( Mandatory = $true, ParameterSetName = 'DownloadDirect')]
    [ValidateSet('MEC', 'MEC2022', 'Ignite', 'Ignite2025', 'Ignite2024', 'Ignite2023', 'Inspire', 'Inspire2023', 'Build', 'Build2025')]
    [string]$Event = '',

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
    [string]$Language = 'English',

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [ValidateSet('de-DE', 'zh-CN', 'en-US', 'ja-JP', 'es-CO', 'fr-FR')]
    [string[]]$Locale = 'en-US',

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
    $ConcurrentFragments = 4,

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'DownloadDirect')]
    [ValidateScript({ Test-Path -Path $_ -PathType Container })]
    [string]$TempPath,

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
    [string]$CookieFile,

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [ValidateSet( 'brave', 'chrome', 'chromium', 'edge', 'firefox', 'opera', 'safari', 'vivaldi', 'whale' )]
    [string]$CookiesFromBrowser
)

# Max age for cache, older than this # hours will force info refresh
$MaxCacheAge = 8

$YouTubeEXE = 'yt-dlp.exe'
$YouTubeDL = Join-Path $PSScriptRoot $YouTubeEXE
$FFMPEG = Join-Path $PSScriptRoot 'ffmpeg.exe'

$YTlink = 'https://github.com/yt-dlp/yt-dlp/releases/download/2023.07.06/yt-dlp.exe'
$FFMPEGlink = 'https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-essentials.zip'

# Fix 'Could not create SSL/TLS secure channel' issues with Invoke-WebRequest
[Net.ServicePointManager]::SecurityProtocol = "tls12, tls11, tls"

$script:BackgroundDownloadJobs = @()

function Iif($Cond, $IfTrue, $IfFalse) {
    if ( $Cond) { $IfTrue } else { $IfFalse }
}

function Fix-FileName ($title) {
    return (((((((($title -replace '\]', ')') -replace '\[', '(') -replace [char]0x202f, ' ') -replace '["\\/\?\*]', ' ') -replace ':', '-') -replace '  ', ' ') -replace '\?\?\?', '') -replace '\<|\>|:|"|/|\\|\||\?|\*', '').Trim()
}

function Get-IEProxy {
    if ( (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings').ProxyEnable -ne 0) {
        $proxies = (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings').proxyServer
        if ($proxies) {
            if ($proxies -ilike "*=*") {
                return $proxies -replace "=", "://" -split (';') | Select-Object -First 1
            }
            else {
                return ('http://{0}' -f $proxies)
            }
        }
        else {
            return $null
        }
    }
    else {
        return $null
    }
}

function Clean-VideoLeftovers ( $videofile) {
    $masks = '.*.mp4.part', '.*.mp4.ytdl'
    foreach ( $mask in $masks) {
        if ( $TempPath) {
            $FileMask = (Join-Path -Path $TempPath -ChildPath (Split-Path -Path $videofile -Leaf)) -replace '.mp4', $mask
        }
        else {
            $FileMask = $videofile -replace '.mp4', $mask
        }
        Get-Item -LiteralPath $FileMask -ErrorAction SilentlyContinue | ForEach-Object {
            Write-Verbose ('Removing leftover file {0}' -f $_.fullname)
            Remove-Item -LiteralPath $_.fullname -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-BackgroundDownloadJobs {
    $Temp = @()
    foreach ( $job in $script:BackgroundDownloadJobs) {

        switch ( $job.Type) {
            1 {
                $isJobRunning = $job.job.State -eq 'Running'
            }
            2 {
                $isJobRunning = -not $job.job.hasExited
            }
            3 {
                $isJobRunning = $job.job.State -eq 'Running'
            }
            default {
                $isJobRunning = $false
            }
        }
        if ( $isJobRunning) {
            $Temp += $job
        }
        else {
            # Job finished, process result
            switch ( $job.Type) {
                1 {
                    $isJobSuccess = $job.job.State -eq 'Completed'
                    $DeckInfo[ $InfoDownload]++
                    Write-Progress -Id $job.job.Id -Activity ('Slidedeck {0} {1}' -f $Job.scheduleCode, $Job.title) -Completed
                }
                2 {
                    $isJobSuccess = Test-Path -LiteralPath $job.file
                    $VideoInfo[ $InfoDownload]++
                    Write-Progress -Id $job.job.Id -Activity ('Video {0} {1}' -f $Job.scheduleCode, $Job.title) -Completed
                }
                3 {
                    $isJobSuccess = Test-Path -LiteralPath $job.file
                    Write-Progress -Id $job.job.Id -Activity ('Captions {0} {1}' -f $Job.scheduleCode, $Job.title) -Completed
                }
                default {
                    $isJobSuccess = $false
                }
            }

            # Test if file is placeholder
            $isPlaceholder = $false
            if ( Test-Path -LiteralPath $job.file) {
                $FileObj = Get-ChildItem -LiteralPath $job.file
                if ( $FileObj.Length -lt 1kb) {

                    if ( @('No resource file is available for download', 'No resource file is available for download for the given id') -contains (Get-Content -LiteralPath $job.File) ) {
                        Write-Warning ('Removing {0} placeholder file {1}' -f $job.scheduleCode, $job.file)
                        Remove-Item -LiteralPath $job.file -Force
                        $isPlaceholder = $true

                        switch ( $job.Type) {
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
                    else {
                        # Placeholder different text?
                    }
                }
            }

            if ( $isJobSuccess -and -not $isPlaceholder) {

                Write-Host ('Downloaded {0}' -f $job.file) -ForegroundColor Green

                # Do we need to adjust timestamp
                if ( $job.Timestamp) {
                    if ( Test-Path -LiteralPath $job.file) {
                        Write-Verbose ('Applying timestamp {0} to {1}' -f $job.Timestamp, $job.file)
                        $FileObj = Get-ChildItem -LiteralPath $job.file
                        $FileObj.CreationTime = Get-Date -Date $job.Timestamp
                        $FileObj.LastWriteTime = Get-Date -Date $job.Timestamp
                    }
                    else {
                        Write-Warning ('File {0} not found for timestamp adjustment' -f $job.file)
                    }
                }

                if ( $job.Type -eq 2) {
                    # Clean video leftovers
                    Clean-VideoLeftovers $job.file
                }
            }
            else {
                switch ( $job.Type) {
                    1 {
                        Write-Host ('Problem downloading or missing slidedeck of {0} {1}' -f $job.scheduleCode, $job.title) -ForegroundColor Red
                        $job.job.ChildJobs | Stop-Job | Out-Null
                        $job.job | Stop-Job -PassThru | Remove-Job -Force | Out-Null
                    }
                    2 {
                        $LastLine = (Get-Content -LiteralPath $job.stdErrTempFile -ErrorAction SilentlyContinue) | Select-Object -Last 1
                        Write-Host ('Problem downloading or missing video of {0} {1}: {2}' -f $job.scheduleCode, $job.title, $LastLine) -ForegroundColor Red
                        Remove-Item -LiteralPath $job.stdOutTempFile, $job.stdErrTempFile -Force -ErrorAction Ignore
                    }
                    3 {
                        Write-Host ('Problem downloading or missing captions of {0} {1}' -f $job.scheduleCode, $job.title) -ForegroundColor Red
                        $job.job.ChildJobs | Stop-Job | Out-Null
                        $job.job | Stop-Job -PassThru | Remove-Job -Force | Out-Null
                    }
                    default {
                    }
                }
            }
        }
    }
    $Num = ($Temp | Measure-Object).Count
    $script:BackgroundDownloadJobs = $Temp
    Show-BackgroundDownloadJobs
    return $Num
}

function Show-BackgroundDownloadJobs {
    $Num = 0
    $NumDeck = 0
    $NumVid = 0
    $NumCaption = 0
    foreach ( $BGJob in $script:BackgroundDownloadJobs) {
        $Num++
        switch ( $BGJob.Type) {
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

    foreach ( $job in $script:BackgroundDownloadJobs) {
        if ( $Job.Type -eq 2) {

            # Get last line of YT log to display for video downloads
            $LastLine = (Get-Content -LiteralPath $job.stdOutTempFile -ErrorAction SilentlyContinue) | Select-Object -Last 1
            if (!( $LastLine)) {
                $LastLine = 'Evaluating..'
            }
            Write-Progress -Id $job.job.id -Activity ('Video {0} {1}' -f $job.scheduleCode, $Job.title) -Status $LastLine -ParentId 2
            $progressId++
        }
    }
}

function Stop-BackgroundDownloadJobs {
    # Trigger update jobs running data
    $null = Get-BackgroundDownloadJobs
    # Stop all slidedeck background jobs
    foreach ( $BGJob in $script:BackgroundDownloadJobs ) {
        switch ( $BGJob.Type) {
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

function Add-BackgroundDownloadJob {
    param(
        $Type,
        $FilePath,
        $DownloadUrl,
        $ArgumentList,
        $File,
        $Timestamp = $null,
        $Title = '',
        $ScheduleCode = ''
    )
    $JobsRunning = Get-BackgroundDownloadJobs
    if ( $JobsRunning -ge $MaxDownloadJobs) {
        Write-Host ('Maximum background download jobs reached ({0}), waiting for free slot - press Ctrl-C once to abort..' -f $JobsRunning)
        while ( $JobsRunning -ge $MaxDownloadJobs) {
            if ([system.console]::KeyAvailable) {
                Start-Sleep 1
                $key = [system.console]::readkey($true)
                if (($key.modifiers -band [consolemodifiers]"control") -and ($key.key -eq "C")) {
                    Write-Host "TERMINATING" -ForegroundColor Red
                    Stop-BackgroundDownloadJobs
                    exit -1
                }
            }
            $JobsRunning = Get-BackgroundDownloadJobs
        }
    }
    switch ( $Type) {
        1 {
            # Slidedeck
            $job = Start-Job -ScriptBlock {
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
            $TempFile = Join-Path ($env:TEMP) (New-Guid).Guid
            $stdOutTempFile = '{0}-Out.log' -f $TempFile
            $stdErrTempFile = '{0}-Err.log' -f $TempFile
            $ProcessParam = @{
                FilePath               = $FilePath
                ArgumentList           = $ArgumentList
                RedirectStandardError  = $stdErrTempFile
                RedirectStandardOutput = $stdOutTempFile
                Wait                   = $false
                Passthru               = $true
                NoNewWindow            = $true
                #WindowStyle= [System.Diagnostics.ProcessWindowStyle]::Normal
            }
            $job = Start-Process @ProcessParam
        }
        3 {
            # Caption
            $job = Start-Job -ScriptBlock {
                param( $url, $file)
                $wc = New-Object System.Net.WebClient
                $wc.Encoding = [System.Text.Encoding]::UTF8
                $wc.DownloadFile( $url, $file)
            } -ArgumentList $DownloadUrl, $FilePath
            $stdOutTempFile = $null
            $stdErrTempFile = $null
        }
    }
    $object = New-Object -TypeName PSObject -Property @{
        Type           = $Type
        job            = $job
        file           = $file
        title          = $Title
        url            = $DownloadUrl
        scheduleCode   = $ScheduleCode
        timestamp      = $timestamp
        stdOutTempFile = $stdOutTempFile
        stdErrTempFile = $stdErrTempFile
    }

    $script:BackgroundDownloadJobs += $object
    Show-BackgroundDownloadJobs
}

##########
# MAIN
##########

Write-Host( '*' * 78)
Write-Host( 'Get-EventSession v4.37')
Write-Host( 'Microsoft event video and slidedeck downloading script')
Write-Host( 'Source: https://github.com/michelderooij/Get-EventSession')
Write-Host( '*' * 78)

if ( $psISE) {
    throw( 'Running from PowerShell ISE is not supported due to requirement to capture console input for proper termination of the script. Please run from a regular PowerShell session.')
}

if ( $Proxy) {
    $ProxyURL = $Proxy
}
else {
    $ProxyURL = Get-IEProxy
}
if ( $ProxyURL) {
    Write-Host "Using proxy address $ProxyURL"
}
else {
    Write-Host "No proxy setting detected, using direct connection"
}

# Determine what event URLs to use.
# Use {0} for session code (eg BRK123), {1} for session id (guid)
switch ( $Event) {
    { 'MEC', 'MEC2022' -contains $_ } {
        $EventName = 'MEC2022'
        $EventType = 'YT'
        $EventYTUrl = 'https://www.youtube.com/playlist?list=PLxdTT6-7g--2POisC5XcDQxUXHhWsoZc9'
        $EventLocale = 'en-us'
        $CaptionExt = 'vtt'
    }
    { 'Ignite', 'Ignite2025' -contains $_ } {
        $EventName = 'Ignite2025'
        $EventType = 'API2'
        $EventAPIUrl = 'https://api-v2.ignite.microsoft.com/api/session/all/en-US'
        $SessionUrl = 'https://medius.microsoft.com/video/asset/HIGHMP4/{0}'
        $CaptionURL = 'https://medius.studios.ms/video/asset/CAPTION/IG25-{0}'
        $SlidedeckUrl = 'https://medius.microsoft.com/video/asset/PPT/{0}'
        $Method = 'GET'
        # Note: to have literal accolades and not string formatter evaluate interior, use a pair:
        $EventSearchBody = '{{"itemsPerPage":{0},"searchFacets":{{"dateFacet":[{{"startDateTime":"2025-11-01T12:00:00.000Z","endDateTime":"2025-11-30T21:59:00.000Z"}}]}},"searchPage":{1},"searchText":"*","sortOption":"Chronological"}}'
        $CaptionExt = 'vtt'
    }
    { 'Ignite2024' -contains $_ } {
        $EventName = 'Ignite2024'
        $EventType = 'API'
        $EventAPIUrl = 'https://api-v2.ignite.microsoft.com'
        $EventSearchURI = 'api/session/search'
        $SessionUrl = 'https://medius.microsoft.com/video/asset/HIGHMP4/{0}'
        $CaptionURL = 'https://medius.studios.ms/video/asset/CAPTION/IG24-{0}'
        $SlidedeckUrl = 'https://medius.microsoft.com/video/asset/PPT/{0}'
        $Method = 'Post'
        # Note: to have literal accolades and not string formatter evaluate interior, use a pair:
        $EventSearchBody = '{{"itemsPerPage":{0},"searchFacets":{{"dateFacet":[{{"startDateTime":"2024-11-19T12:00:00.000Z","endDateTime":"2024-11-22T21:59:00.000Z"}}]}},"searchPage":{1},"searchText":"*","sortOption":"Chronological"}}'
        $CaptionExt = 'vtt'
    }
    { 'Ignite2023' -contains $_ } {
        $EventName = 'Ignite2023'
        $EventType = 'API'
        $EventAPIUrl = 'https://api.ignite.microsoft.com'
        $EventSearchURI = 'api/session/search'
        $SessionUrl = 'https://medius.studios.ms/Embed/video-nc/IG23-{0}'
        $CaptionURL = 'https://medius.studios.ms/video/asset/CAPTION/IG23-{0}'
        $SlidedeckUrl = 'https://medius.microsoft.com/video/asset/PPT/{0}'
        $Method = 'Post'
        # Note: to have literal accolades and not string formatter evaluate interior, use a pair:
        $EventSearchBody = '{{"itemsPerPage":{0},"searchFacets":{{"dateFacet":[{{"startDateTime":"2023-11-13T12:00:00.000Z","endDateTime":"2023-11-18T21:59:00.000Z"}}]}},"searchPage":{1},"searchText":"*","sortOption":"Chronological"}}'
        $CaptionExt = 'vtt'
    }
    { 'Inspire', 'Inspire2023' -contains $_ } {
        $EventName = 'Inspire2023'
        $EventType = 'API'
        $EventAPIUrl = 'https://api.inspire.microsoft.com'
        $EventSearchURI = 'api/session/search'
        $SessionUrl = 'https://medius.studios.ms/video/asset/HIGHMP4/INSP23-{0}'
        $CaptionURL = 'https://medius.studios.ms/video/asset/CAPTION/INSP23-{0}'
        $SlidedeckUrl = 'https://medius.studios.ms/video/asset/PPT/INSP23-{0}'
        $Method = 'Post'
        $EventSearchBody = '{{"itemsPerPage":{0},"searchFacets":{{"dateFacet":[{{"startDateTime":"2023-01-01T08:00:00-05:00","endDateTime":"2023-08-01T19:00:00-05:00"}}]}},"searchPage":{1},"searchText":"*","sortOption":"Chronological"}}'
        $CaptionExt = 'vtt'
    }
    { 'Build', 'Build2025' -contains $_ } {
        $EventName = 'Build2025'
        $EventType = 'API2'
        $EventAPIUrl = 'https://eventtools.event.microsoft.com/build2025-prod/fallback/session-all-en-us.json'
        $SessionUrl = 'https://medius.microsoft.com/video/asset/HIGHMP4/{0}'
        $CaptionURL = 'https://medius.microsoft.com/video/asset/CAPTION/{0}'
        $SlidedeckUrl = 'https://medius.microsoft.com/video/asset/PPT/{0}'
        $Method = 'GET'
        $CaptionExt = 'vtt'
        $PreferDirect = $True
    }
    { 'Build2024' -contains $_ } {
        $EventName = 'Build2024'
        $EventType = 'API'
        $EventAPIUrl = 'https://api-v2.build.microsoft.com'
        $EventSearchURI = 'api/session/search'
        $SessionUrl = 'https://medius.studios.ms/video/asset/HIGHMP4/B24-{0}'
        $CaptionURL = 'https://medius.studios.ms/video/asset/CAPTION/B24-{0}'
        $SlidedeckUrl = 'https://medius.studios.ms/video/asset/PPT/B24-{0}'
        $Method = 'Post'
        $EventSearchBody = '{{"itemsPerPage":{0},"searchFacets":{{"dateFacet":[{{"startDateTime":"2024-05-21T08:00:00-05:00","endDateTime":"2024-05-24T19:00:00-05:00"}}]}},"searchPage":{1},"searchText":"*","sortOption":"Chronological"}}'
        $CaptionExt = 'vtt'
    }
    { 'Build2023' -contains $_ } {
        $EventName = 'Build2023'
        $EventType = 'API'
        $EventAPIUrl = 'https://api-v2.build.microsoft.com'
        $EventSearchURI = 'api/session/search'
        $SessionUrl = 'https://medius.studios.ms/video/asset/HIGHMP4/B23-{0}'
        $CaptionURL = 'https://medius.studios.ms/video/asset/CAPTION/B23-{0}'
        $SlidedeckUrl = 'https://medius.studios.ms/video/asset/PPT/B23-{0}'
        $Method = 'Post'
        $EventSearchBody = '{{"itemsPerPage":{0},"searchFacets":{{"dateFacet":[{{"startDateTime":"2023-01-01T08:00:00-05:00","endDateTime":"2023-12-31T19:00:00-05:00"}}]}},"searchPage":{1},"searchText":"*","sortOption":"Chronological"}}'
        $CaptionExt = 'vtt'
    }
    default {
        Write-Host ('Unknown event: {0}' -f $Event) -ForegroundColor Red
        exit -1
    }
}

if (-not ($InfoOnly)) {

    # If no download folder set, use system drive with event subfolder
    if ( -not( $DownloadFolder)) {
        $DownloadFolder = '{0}\{1}' -f $ENV:SystemDrive, $EventName
    }

    Add-Type -AssemblyName System.Web
    Write-Host "Using download path: $DownloadFolder"
    # Create the local content path if not exists
    if ( (Test-Path $DownloadFolder) -eq $false ) {
        New-Item -Path $DownloadFolder -ItemType Directory | Out-Null
    }

    if ( $NoVideos) {
        Write-Host 'Will skip downloading videos'
        $DownloadVideos = $false
    }
    else {
        if (-not( Test-Path $YouTubeDL)) {
            Write-Host ('{0} not found, will try to download from {1}' -f $YouTubeEXE, $YTLink)
            Invoke-WebRequest -Uri $YTLink -OutFile $YouTubeDL -Proxy $ProxyURL
        }
        if ( Test-Path $YouTubeDL) {
            Write-Host ('Running self-update of {0}' -f $YouTubeEXE)

            $Arg = @('-U')
            if ( $ProxyURL) { $Arg += "--proxy $ProxyURL" }

            $pinfo = New-Object System.Diagnostics.ProcessStartInfo
            $pinfo.FileName = $YouTubeDL
            $pinfo.RedirectStandardError = $true
            $pinfo.RedirectStandardOutput = $true
            $pinfo.UseShellExecute = $false
            $pinfo.Arguments = $Arg

            Write-Verbose ('Running {0} using {1}' -f $pinfo.FileName, ($pinfo.Arguments -join ' '))
            try {
                $p = New-Object System.Diagnostics.Process
                $p.StartInfo = $pinfo
                $p.Start() | Out-Null
                $stdout = $p.StandardOutput.ReadToEnd()
                $stderr = $p.StandardError.ReadToEnd()
                $p.WaitForExit()
            }
            catch {
                throw ('Problem running {0}. Make sure this is an x86 system, and the required Visual C++ 2010 redistribution package is installed (available from https://www.microsoft.com/en-US/download/details.aspx?id=5555).' -f $YouTubeEXE)
            }
            $DownloadVideos = $true
        }
        else {
            Write-Warning ('Unable to locate or download {0}, will skip downloading YouTube videos' -f $YouTubeEXE)
            $DownloadVideos = $false
        }

        if (-not( Test-Path $FFMPEG)) {

            Write-Host ('ffmpeg.exe not found, will try to download from {0}' -f $FFMPEGlink)
            $TempFile = Join-Path $PSScriptRoot 'ffmpeg-latest-win32-static.zip'
            Invoke-WebRequest -Uri $FFMPEGlink -OutFile $TempFile -Proxy $ProxyURL

            if ( Test-Path $TempFile) {
                Add-Type -AssemblyName System.IO.Compression.FileSystem
                Write-Host ('{0} downloaded, trying to extract ffmpeg.exe' -f $TempFile)
                $FFMPEGZip = [System.IO.Compression.ZipFile]::OpenRead( $TempFile)
                $FFMPEGEntry = $FFMPEGZip.Entries | Where-Object { $_.FullName -like '*/ffmpeg.exe' }
                if ( $FFMPEGEntry) {
                    try {
                        [System.IO.Compression.ZipFileExtensions]::ExtractToFile( $FFMPEGEntry, $FFMPEG)
                        $FFMPEGZip.Dispose()
                        Remove-Item -LiteralPath $TempFile -Force
                    }
                    catch {
                        throw ('Problem extracting ffmpeg.exe from {0}' -f $FFMPEGZip)
                    }
                }
                else {
                    throw 'ffmpeg.exe missing in downloaded archive'
                }
            }
        }
        if ( Test-Path $FFMPEG) {
            Write-Host ('ffmpeg.exe located at {0}' -f $FFMPEG)
            $DownloadAMSVideos = $true
        }
        else {
            Write-Warning 'Unable to locate or download ffmpeg.exe, will skip downloading Azure Media Services videos'
            $DownloadAMSVideos = $false
        }
    }
}

$SessionCache = Join-Path $PSScriptRoot ('{0}-Sessions.cache' -f $EventName)
$SessionCacheValid = $false

if ( $Refresh) {
    Write-Host 'Refresh specified, will read session information from the online catalog'
}
else {
    if ( Test-Path $SessionCache) {
        try {
            if ( (Get-ChildItem -LiteralPath $SessionCache).LastWriteTime -ge (Get-Date).AddHours( - $MaxCacheAge)) {
                Write-Host 'Session cache file found, reading session information'
                $data = Import-Clixml -LiteralPath $SessionCache -ErrorAction SilentlyContinue
                $SessionCacheValid = $true
            }
            else {
                Write-Warning 'Cache information expired, will re-read information from catalog'
            }
        }
        catch {
            Write-Host 'Error reading cache file or cache file invalid - will read from online catalog' -ForegroundColor Red
        }
    }
}

if ( -not( $SessionCacheValid)) {

    switch ($EventType) {
        'API2' {
            Write-Host ('Reading {0} session catalog' -f $EventName)
            $web = @{
                userAgent  = 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.103 Safari/537.36'
                requestUri = [uri]('{0}' -f $EventAPIUrl)
                headers    = @{'Content-Type' = 'application/json; charset=utf-8'; 'Accept-Encoding' = 'deflate, gzip' }
                Timeout    = 300
            }
            try {
                Write-Verbose ('Using URI {0}' -f $web.requestUri)
                $ResultsResponse = Invoke-RestMethod -Uri $web.requestUri -Method $Method -Headers $web.headers -UserAgent $web.userAgent -WebSession $session -Proxy $ProxyURL -Timeout $web.Timeout
            }
            catch {
                throw ('Problem retrieving session catalog: {0}' -f $error[0])
            }
            [int32]$sessionCount = ($ResultsResponse | Measure-Object ).Count
            Write-Host ('Processing information for {0} sessions' -f $sessionCount)
            $data = [System.Collections.ArrayList]@()
            $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet', [string[]]('sessionCode', 'title'))
            $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
            foreach ( $item in $ResultsResponse) {
                $Item.PSObject.Properties | ForEach-Object {
                    if ( @('speakerNames') -icontains $_.Name ) {
                        $Item.($_.Name) = @($_.Value)
                    }
                    if ( @('products', 'contentCategory') -icontains $_.Name ) {
                        $Item.($_.Name) = @($_.Value -replace [char]9, '/')
                    }
                    if ( @('topic', 'sessionType', 'sessionLevel', 'audienceTypes', 'deliveryTypes', 'viewingOptions', 'event', 'programmingLanguages') -icontains $_.Name ) {
                        $Item.($_.Name) = $_.Value.displayValue -join ','
                    }
                }
                Write-Verbose ('Adding info for session {0}' -f $Item.sessionCode)
                $Item.PSObject.TypeNames.Insert(0, 'Session.Information')
                $Item | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                $data.Add( $Item) | Out-Null
            }
        }
        'API' {
            Write-Host ('Reading {0} session catalog' -f $EventName)
            $web = @{
                userAgent    = 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.103 Safari/537.36'
                requestUri   = [uri]('{0}/{1}' -f $EventAPIUrl, $EventSearchURI)
                headers      = @{'Content-Type' = 'application/json; charset=utf-8'; 'Accept-Encoding' = 'deflate, gzip' }
                itemsPerPage = 1000
            }
            try {
                $SearchBody = $EventSearchBody -f '1', '1'
                Write-Verbose ('Using URI {0}' -f $web.requestUri)
                $searchResultsResponse = Invoke-RestMethod -Uri $web.requestUri -Body $searchbody -Method $Method -Headers $web.headers -UserAgent $web.userAgent -WebSession $session -Proxy $ProxyURL
            }
            catch {
                throw ('Problem retrieving session catalog: {0}' -f $error[0])
            }
            [int32]$sessionCount = $searchResultsResponse.total
            [int32]$remainder = 0
            $PageCount = [System.Math]::DivRem($sessionCount, $web.itemsPerPage, [ref]$remainder)
            if ($remainder -gt 0) {
                $PageCount++
            }
            Write-Host ('Reading information for {0} sessions' -f $sessionCount)
            $data = [System.Collections.ArrayList]@()
            $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet', [string[]]('sessionCode', 'title'))
            $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
            for ($page = 1; $page -le $PageCount; $page++) {
                Write-Progress -Id 1 -Activity "Retrieving Session Catalog" -Status "Processing page $page of $PageCount" -PercentComplete ($page / $PageCount * 100)
                $SearchBody = $EventSearchBody -f $web.itemsPerPage, $page
                $searchResultsResponse = Invoke-RestMethod -Uri $web.requestUri -Body $searchbody -Method $Method -Headers $web.headers -UserAgent $web.userAgent -WebSession $session -Proxy $ProxyURL
                foreach ( $Item in $searchResultsResponse.data) {
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
            $Arg = [System.Collections.ArrayList]@()
            if ( $ProxyURL) {
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

            if ($p.ExitCode -ne 0) {
                throw ('Problem running {0}: {1}' -f $YouTubeEXE, $stderr)
            }

            try {
                Write-Verbose ('Converting from Json ..')
                # Trim any trailing empty lines, convert single string with line-breaks to array for JSON conversion
                $JsonData = ($stdout.Trim([System.Environment]::Newline) -split "`n") | ConvertFrom-Json
            }
            catch {
                throw( 'Output does not seem to be proper JSON format, see {0}' -f $TempJsonFile)
            }

            foreach ( $Item in $JsonData) {

                $SpeakerNames = [System.Collections.ArrayList]@()

                # Description match pattern? Set Desc+Speakers, otherwise Desc=Description, assume no Speakers defined

                if ($Item.Description -match '^(?<Description>[\s\S]*?)(\s)*(Download the slide deck from (?<Slidedeck>https:\/\/.*?)[\.]?)?(\s)*(Speaker(s)?:(\s)?(?<Speakers>.*))?(\s)*$') {
                    $Description = $Matches.Description
                    $Matches.Speakers -split ';' | ForEach-Object { $SpeakerNames.Add( $_.Trim() ) | Out-Null }
                    $SlidedeckUrl = $Matches.Slidedeck
                }
                else {
                    $Description = $Item.Description
                    $SlidedeckUrl = $null
                }

                # Slidedeck url, construct real link:
                if ( $SlidedeckUrl) {
                    # https://www.microsoft.com/en-us/download/details.aspx?id=104608 -> https://www.microsoft.com/en-us/download/confirmation.aspx?id=104608

                    if ( $SlidedeckUrl -match '^(?<host>https:\/\/www\.microsoft\.com).*id=(?<id>[\d]+)$') {
                        $SlideDeck = '{0}/en-us/download/confirmation.aspx?id={1}' -f $Matches.host, $Matches.id
                    }
                    else {
                        Write-Warning ('Unexpected slide deck URL format: {0}' -f $SlidedeckUrl)
                        $Slidedeck = $null
                    }
                }
                else {
                    $SlideDeck = $null
                }

                $object = [PSCustomObject]@{
                    sessionCode       = [string]('{0:d2}' -f $Item.playlist_index)
                    SessionType       = 'On-Demand'
                    Title             = $Item.Title
                    Description       = $Description
                    onDemand          = $Item.webpage_url
                    Views             = $Item.view_count
                    Likes             = $Item.like_count
                    Duration          = [timespan]::FromSeconds( $Item.duration).ToString()
                    langLocale        = $EventLocale
                    SolutionArea      = $Item.Tags
                    contentCategory   = $Item.categories
                    SpeakerNames      = $SpeakerNames
                    Slidedeck         = $Slidedeck
                    startDateTime     = [Datetime]::ParseExact( $Item.upload_date, 'yyyyMMdd', $null)
                    onDemandThumbnail = ($Item.thumbnails | Sort-Object -Property Id | Select-Object -First 1).Url
                }
                Write-Verbose ('Adding info for session {0}' -f $Item.Title)
                $data.Add( $object) | Out-Null
            }
        }

        default {
            throw( 'Unknown event catalog type {0}' -f $EventType)
        }
    }

    Write-Host 'Storing session information'
    $data | Export-Clixml -Encoding Unicode -Force -LiteralPath $SessionCache

}

$SessionsToGet = $data
$TotalNumberOfSessions = ($SessionsToGet | Measure-Object).Count

if ($scheduleCode) {
    Write-Verbose ('Session code(s) specified: {0}' -f ($ScheduleCode -join ','))
    $SessionsToGet = $SessionsToGet | Where-Object { $scheduleCode -contains $_.sessioncode }
}

if ($ExcludeCommunityTopic) {
    Write-Verbose ('Excluding community topic: {0}' -f $ExcludeCommunityTopic)
    $SessionsToGet = $SessionsToGet | Where-Object { $ExcludeCommunityTopic -inotcontains $_.CommunityTopic }
}

if ($Speaker) {
    Write-Verbose ('Speaker keyword specified: {0}' -f $Speaker)
    $SessionsToGet = $SessionsToGet | Where-Object { $_.speakerNames | Where-Object { $_ -ilike $Speaker } }
}

if ($Product) {
    Write-Verbose ('Product specified: {0}' -f $Product)
    $SessionsToGet = $SessionsToGet | Where-Object { $_.products | Where-Object { $_ -ilike $Product } }
}

if ($Category) {
    Write-Verbose ('Category specified: {0}' -f $Category)
    $SessionsToGet = $SessionsToGet | Where-Object { $_.contentCategory | Where-Object { $_ -ilike $Category } }
}

if ($SolutionArea) {
    Write-Verbose ('SolutionArea specified: {0}' -f $SolutionArea)
    $SessionsToGet = $SessionsToGet | Where-Object { $_.solutionArea | Where-Object { $_ -ilike $SolutionArea } }
}

if ($LearningPath) {
    Write-Verbose ('LearningPath specified: {0}' -f $LearningPath)
    $SessionsToGet = $SessionsToGet | Where-Object { $_.learningPath | Where-Object { $_ -ilike $LearningPath } }
}

if ($Topic) {
    Write-Verbose ('Topic specified: {0}' -f $Topic)
    $SessionsToGet = $SessionsToGet | Where-Object { $_.topic | Where-Object { $_ -ilike $Topic } }
}

if ($Locale) {
    Write-Verbose ('Locale(s) specified: {0}' -f ($Locale -join ','))
    $SessionsToGetTemp = [System.Collections.ArrayList]@()
    foreach ( $item in $Locale) {
        $SessionsToGet | Where-Object { $item -ieq $_.langLocale } | ForEach-Object { $null = $SessionsToGetTemp.Add(  $_ ) }
    }
    $SessionsToGet = $SessionsToGetTemp | Sort-Object -Unique -Property sessionCode
}

if ($Title) {
    Write-Verbose ('Title keyword(s) specified: {0}' -f ( $Title -join ','))
    $SessionsToGetTemp = [System.Collections.ArrayList]@()
    foreach ( $item in $Title) {
        $SessionsToGet | Where-Object { $_.title -ilike ('*{0}*' -f $item) } | ForEach-Object { $null = $SessionsToGetTemp.Add(  $_ ) }
    }
    $SessionsToGet = $SessionsToGetTemp | Sort-Object -Unique -Property sessionCode
}

if ($Keyword) {
    Write-Verbose ('Description keyword(s) specified: {0}' -f ( $Keyword -join ','))
    $SessionsToGetTemp = [System.Collections.ArrayList]@()
    foreach ( $item in $Keyword) {
        $SessionsToGet | Where-Object { $_.description -ilike ('*{0}*' -f $item) } | ForEach-Object { $null = $SessionsToGetTemp.Add(  $_ ) }
    }
    $SessionsToGet = $SessionsToGetTemp | Sort-Object -Unique -Property sessionCode
}

if ($NoRepeats) {
    Write-Verbose ('Skipping repeated sessions')
    $SessionsToGet = $SessionsToGet | Where-Object { $_.sessionCode -inotmatch '^*R[1-9]?$' -and $_.sessionCode -inotmatch '^[A-Z]+[0-9]+[B-C]+$' }
}

if ( $InfoOnly) {
    Write-Verbose ('There are {0} sessions matching your criteria.' -f (($SessionsToGet | Measure-Object).Count))
    Write-Output $SessionsToGet
}
else {

    if ( $OGVPicker) {
        $SessionsToGet = $SessionsToGet | Out-GridView -Title 'Select Videos to Download, or Cancel for all Videos' -PassThru
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

    if ( $DownloadFolder -inotlike '\\?\*') {
        # Apply extended-length path prefix to support long paths
        if ( $DownloadFolder -ilike '\\*') {
            $DownloadFolder = '\\?\UNC\{0}' -f $DownloadFolder.Substring(2)
        }
        else {
            $DownloadFolder = '\\?\{0}' -f $DownloadFolder
        }
    }
}

foreach ($SessionToGet in $SessionsToGet) {

    $i++
    Write-Progress -Id 1 -Activity 'Inspecting session information' -Status "Processing session $i of $SessionsSelected" -PercentComplete ($i / $SessionsSelected * 100)
    if ( $SessionToGet.sessionCode) {
        $FileName = Fix-FileName ('{0}-{1}' -f $SessionToGet.sessionCode.Trim(), $SessionToGet.title.Trim())
    }
    else {
        $FileName = Fix-FileName ('{0}' -f $SessionToGet.title.Trim())
    }
    if (! ([string]::IsNullOrEmpty( $SessionToGet.startDateTime))) {
        # Get session localized timestamp, undoing TZ adjustments
        $SessionTime = [System.TimeZoneInfo]::ConvertTime((Get-Date -Date $SessionToGet.startDateTime).ToUniversalTime(), $myTimeZone ).toString('g')
    }
    else {
        $SessionTime = $null
    }
    Write-Host ('Processing info session {0} from {1} [{2}]' -f $FileName, (Iif -Cond $SessionTime -IfTrue $SessionTime -IfFalse 'No Timestamp'), $SessionToGet.langLocale)
    if (!([string]::IsNullOrEmpty( $SessionToGet.startDateTime)) -and (Get-Date -Date $SessionToGet.startDateTime) -ge (Get-Date)) {
        Write-Verbose ('Skipping session {0}: Didn''t take place yet' -f $SessionToGet.sessioncode)
    }
    else {

        if ( ! $NoVideos) {

            $onDemandPage = $null

            if ( $DownloadVideos -or $DownloadAMSVideos) {

                $vidfileName = '{0}.mp4' -f $FileName
                $vidFullFile = Join-Path -Path $DownloadFolder -ChildPath $vidfileName

                if ((Test-Path -LiteralPath $vidFullFile) -and -not $Overwrite) {
                    Write-Host ('Video exists {0}' -f $vidfileName) -ForegroundColor Gray
                    $VideoInfo[ $InfoExist]++
                    # Clean video leftovers
                    Clean-VideoLeftovers $vidFullFile
                }
                else {
                    $downloadLink = $null
                    if ( [string]::IsNullOrEmpty( $SessionToGet.downloadVideoLink)) {
                        if ( [string]::IsNullOrEmpty( $SessionToGet.onDemand)) {
                            if ( $NoGuessing) {
                                $downloadLink = $null
                            }
                            else {
                                # Try session page, eg https://medius.studios.ms/Embed/Video/IG18-BRK2094
                                $downloadLink = $SessionUrl -f $SessionToGet.SessionCode
                            }
                        }
                        else {
                            $downloadLink = $SessionToGet.onDemand
                        }
                    }
                    else {
                        if ( [string]::IsNullOrEmpty( $SessionToGet.onDemand)) {
                            $downloadLink = $SessionToGet.downloadVideoLink
                        }
                        else {
                            if ( $PreferDirect) {
                                $downloadLink = $SessionToGet.downloadVideoLink
                            }
                            else {
                                $downloadLink = $SessionToGet.onDemand
                            }
                        }
                    }

                    Write-Verbose ('Checking download link {0}' -f $downloadLink)
                    try {
                        $Response = Invoke-WebRequest -Method HEAD -Uri $downloadLink -Proxy $ProxyURL -DisableKeepAlive -ErrorAction SilentlyContinue
                        $DirectLink = @( 'video/mp4', 'video/MP2T') -contains $Response.Headers.'Content-Type'
                    }
                    catch {
                        $DirectLink = $False
                    }

                    if ( ! ( $DirectLink) -and $downloadLink -match '(medius\.studios\.ms\/Embed\/Video|medius\.microsoft\.com|mediastream\.microsoft\.com)' ) {
                        $DownloadedPage = Invoke-WebRequest -Method Get -Uri $downloadLink -Proxy $ProxyURL -DisableKeepAlive -ErrorAction SilentlyContinue
                        $OnDemandPage = $DownloadedPage.RawContent
                        $Endpoint = $null

                        if ( $OnDemandPage -match 'StreamUrl = "(?<AzureStreamURL>https://mediusprod\.streaming\.mediaservices\.windows\.net/.*manifest)";') {
                            Write-Verbose ('Using Azure Media Services URL {0}' -f $matches.AzureStreamURL)
                            $Endpoint = '{0}(format=mpd-time-csf)' -f $matches.AzureStreamURL
                        }
                        if ( $OnDemandPage -match 'StreamUrl = "(?<AzureStreamURL>https://stream\.event\.microsoft\.com/.*master\.m3u8)";') {
                            Write-Verbose ('Using Azure Media Stream URL {0}' -f $matches.AzureStreamURL)
                            $Endpoint = '{0}?(format=mpd-time-csf)' -f $matches.AzureStreamURL
                        }

                        if ($Endpoint) {
                            $Arg = @( ('-o "{0}"' -f ($vidFullFile -replace '%', '%%')), $Endpoint)

                            # Construct Format for this specific video, language and audio languages available
                            if ( $Format) {
                                $ThisFormat = $Format
                            }
                            else {
                                $ThisFormat = 'worstvideo+bestaudio'
                            }

                            if ( $SessionToGet.audioLanguage) {

                                if ( $SessionToGet.audioLanguage.Count -gt 1) {
                                    # Session has multiple audio tracks
                                    if ( $SessionToGet.audioLanguage -icontains $Language) {
                                        Write-Warning ('Multiple audio languages available; will try downloading {0} audio stream' -f $Language)
                                        $ThisLanguage = $Language
                                    }
                                    else {
                                        $ThisLanguage = $SessionToGet.audioLanguage | Select-Object -First 1
                                        Write-Warning ('Requested language {0} not available; will use default stream ({1})' -f $Language, $ThisLanguage)
                                    }

                                    # Take specified Format apart so we can insert the language filter per specification
                                    $ThisFormatElem = $ThisFormat -split ','
                                    $NewFormat = [System.Collections.ArrayList]@()
                                    foreach ( $Elem in $ThisFormatElem) {
                                        if ( $Elem -match '^(?<pre>.*audio)(\[(?<audioparam>.*)\])?(?<post>(.*)?)$' ) {
                                            if ( $matches.audioparam) {
                                                $NewFormatElem = '{0}[format_id*={1},{2}]{3}' -f $matches.Pre, $ThisLanguage, $matches.audioparam, $matches.post
                                            }
                                            else {
                                                $NewFormatElem = '{0}[format_id*={1}]{2}' -f $matches.Pre, $ThisLanguage, $matches.post
                                            }
                                        }
                                        else {
                                            $NewFormatElem = $Elem
                                            Write-Warning ('Problem determining where to add language criteria in {0}, leaving criteria as-is' -f $Elem)
                                        }
                                        $null = $NewFormat.Add( $NewFormatElem)
                                    }

                                    # With language filters determined, recreate filter and add whole non-language specific qualifiers as next best
                                    $ThisFormat = ($NewFormat -join ','), $ThisFormat -join ','

                                }
                                else {
                                    # Only 1 Language available, so use default audio stream
                                    Write-Warning ('Only single audio stream available, will use default audio stream')
                                }
                            }
                            else {
                                # No multiple audio languages, use default audio stream
                                Write-Warning ('Multiple audio streams not available, will use default audio stream')
                            }
                            $Arg += ('--format {0}' -f $ThisFormat)
                        }
                        else {
                            # Check for embedded YouTube
                            if ( $OnDemandPage -match '"https:\/\/www\.youtube-nocookie\.com\/embed\/(?<YouTubeID>.+?)\?enablejsapi=1&"') {
                                $Endpoint = 'https://www.youtube.com/watch?v={0}' -f $matches.YouTubeID
                                Write-Verbose ('Using YouTube URL {0}' -f $Endpoint)
                                $Arg = @( ('-o "{0}"' -f ($vidFullFile -replace '%', '%%')), $Endpoint)
                                if ( $Format) { $Arg += ('--format {0}' -f $Format) } else { $Arg += ('--format "bestvideo[ext=mp4]+bestaudio[ext=m4a]/best[ext=mp4]/best"') }
                                if ( $Subs) { $Arg += ('--sub-lang {0}' -f ($Subs -join ',')), ('--write-sub'), ('--write-auto-sub'), ('--convert-subs srt') }

                                if ( $CookiesFile) { $Arg += ('--cookies {0}' -f $CookiesFile) }
                                if ( $CookiesFromBrowser) { $Arg += ('--cookies-from-browser {0}' -f $CookiesFromBrowser) }
                            }
                            else {
                                Write-Warning "Skipping: Embedded AMS or YouTube URL not found"
                                $EndPoint = $null
                            }
                        }
                    }
                    else {
                        # Direct
                        Write-Verbose ('Using direct video link {0}' -f $downloadLink)
                        if ( $downloadLink) {
                            $Endpoint = $downloadLink
                            $Arg = @( ('-o "{0}"' -f $vidFullFile), $downloadLink)
                        }
                        else {
                            Write-Warning ('No video link for {0}' -f ($SessionToGet.Title))
                            $Endpoint = $null
                        }
                    }
                    if ( $Endpoint) {
                        # Direct, AMS or YT video found, attempt download but first define common parameters
                        if ( $ProxyURL) {
                            $Arg += ('--proxy "{0}"' -f $ProxyURL)
                        }
                        $Arg += '-t mp4'
                        $Arg += '--socket-timeout 90'
                        $Arg += '--no-check-certificate'
                        $Arg += '--retries 15'
                        $Arg += '--concurrent-fragments {0}' -f $ConcurrentFragments
                        if ( $Subs) { $Arg += ('--sub-lang {0}' -f ($Subs -join ',')), ('--write-sub'), ('--write-auto-sub'), ('--convert-subs srt') }

                        if ( $TempPath) {
                            # When using temp path, we need to use relative path for file and use home for location
                            $OutputTemp = ($Arg | Where-Object { $_ -like '-o *' })
                            $OutputTemp = $OutputTemp.substring(4, $OutputTemp.Length - 4 - 1)
                            $Arg = $Arg | Where-Object { $_ -inotlike '-o *' }
                            $Arg += '-o "{0}"' -f (Split-Path -Path $OutputTemp -Leaf)
                            $Arg += '-P home:"{0}"' -f (Split-Path -Path $OutputTemp -Parent).TrimEnd('\')
                            $Arg += '-P temp:"{0}"' -f $TempPath.TrimEnd('\')
                        }
                        if ( $Overwrite) {
                            $Arg += '--force-overwrites'
                        }

                        Write-Verbose ('Running: {0} {1}' -f $YouTubeEXE, ($Arg -join ' '))
                        Add-BackgroundDownloadJob -Type 2 -FilePath $YouTubeDL -ArgumentList $Arg -File $vidFullFile -Timestamp $SessionTime -scheduleCode ($SessionToGet.sessioncode) -Title ($SessionToGet.Title)
                    }
                    else {
                        # Video not available or no link found
                        $VideoInfo[ $InfoPlaceholder]++
                    }
                }
                if ( $Captions) {
                    $captionExtFile = $vidFullFile -replace '.mp4', ('.{0}' -f $CaptionExt)

                    if ((Test-Path -LiteralPath $captionExtFile) -and -not $Overwrite) {
                        Write-Host ('Caption file exists {0}' -f $captionExtFile) -ForegroundColor Gray
                    }
                    else {

                        # Caption file in AMS needs seperate download, fetch onDemand page if not already downloaded for video
                        if (! $OnDemandPage) {
                            if ( $SessionToGet.onDemand) {
                                try {
                                    Write-Host ('Fetching video page to retrieve transcript information from {0}' -f $SessionToGet.onDemand)
                                    $DownloadedPage = Invoke-WebRequest -Uri $SessionToGet.onDemand -Proxy $ProxyURL -DisableKeepAlive -ErrorAction SilentlyContinue
                                    if ( $DownloadedPage) {
                                        $OnDemandPage = $DownloadedPage.RawContent
                                    }
                                }
                                catch {
                                    #Problem retrieving file, look for alternative options
                                }
                            }

                        }
                        # Check for vtt files before we check any direct caption file (likely docx now)
                        $captionFileLink = $Null
                        if ( $OnDemandPage -match 'captionsConfiguration = (?<CaptionsJSON>{.*});') {
                            $CaptionConfig = ($matches.CaptionsJSON | ConvertFrom-Json).languageList
                            if ( $Subs) {
                                $captionFileLink = ($CaptionConfig | Where-Object { $_.srclang -eq $Subs }).src
                            }
                            if (! $captionFileLink) {
                                $captionFileLink = ($CaptionConfig | Where-Object { $_.srclang -eq 'en' }).src
                            }
                        }
                        if ( ! $CaptionFileLink) {
                            $captionFileLink = $SessionToGet.captionFileLink
                        }
                        if ( ! $captionFileLink) {

                            if (! $OnDemandPage) {
                                # Try if there is caption file reference on page
                                try {
                                    $DownloadedPage = Invoke-WebRequest -Uri $downloadLink -Proxy $ProxyURL -DisableKeepAlive -ErrorAction SilentlyContinue
                                    $OnDemandPage = $DownloadedPage.RawContent
                                }
                                catch {
                                    $DownloadedPage = $null
                                    $onDemandPage = $null
                                }
                            }
                            else {
                                # Reuse one from video download
                            }

                            if ( $OnDemandPage -match '"(?<AzureCaptionURL>https:\/\/mediusprodstatic\.studios\.ms\/asset-[a-z0-9\-]+\/transcript\{0}\?.*?)"' -f $CaptionExt) {
                                $captionFileLink = $matches.AzureCaptionURL
                            }
                            if ( ! $captionFileLink) {
                                $captionFileLink = $captionURL -f $SessionToGet.SessionCode
                            }
                        }
                        if ( $captionFileLink) {
                            Write-Verbose ('Retrieving caption file from URL {0}' -f $captionFileLink)

                            $captionFullFile = $captionExtFile
                            Write-Verbose ('Attempting download {0} to {1}' -f $captionFileLink, $captionFullFile)
                            Add-BackgroundDownloadJob -Type 3 -FilePath $captionExtFile -DownloadUrl $captionFileLink -File $captionFullFile -Timestamp $SessionTime -scheduleCode ($SessionToGet.sessioncode) -Title ($SessionToGet.Title)

                        }
                        else {
                            Write-Warning "Subtitles requested, but no Caption URL found"
                        }
                    }
                }
                $captionFileLink = $null
                $OnDemandPage = $null
            }
        }

        if (! $NoSlidedecks) {
            if ( !( [string]::IsNullOrEmpty( $SessionToGet.slideDeck)) ) {
                $downloadLink = $SessionToGet.slideDeck
            }
            else {
                if ( $NoGuessing) {
                    $downloadLink = $null
                }
                else {
                    # Try alternative construction
                    $downloadLink = $SlidedeckUrl -f $SessionToGet.SessionCode
                }
            }

            if ($downloadLink -match "view.officeapps.live.com.*PPTX" -or $downloadLink -match 'downloaddocument' -or $downloadLink -match 'medius' -or $downloadLink -match 'confirmation\.aspx') {

                $DownloadURL = [System.Web.HttpUtility]::UrlDecode( $downloadLink )
                try {
                    if ( $downloadLink -notmatch 'confirmation\.aspx') {
                        $ValidUrl = Invoke-WebRequest -Uri $DownloadURL -Method HEAD -UseBasicParsing -DisableKeepAlive -MaximumRedirection 10 -ErrorAction SilentlyContinue
                    }
                    else {
                        $ValidUrl = Invoke-WebRequest -Uri $DownloadURL -Method GET -UseBasicParsing -DisableKeepAlive -MaximumRedirection 10 -ErrorAction SilentlyContinue
                    }
                }
                catch {
                    $ValidUrl = $false
                }

                if ( $downloadLink -match 'confirmation\.aspx' -and $ValidURL.Headers.'Content-Type' -ilike 'text/html') {
                    # Extra parsing for MS downloads
                    if ( $ValidUrl.RawContent -match 'href="(?<Url>https:\/\/download\.microsoft\.com\/download[\/0-9\-]*\/.*(pdf|pptx))".*click here to download manually') {
                        $DownloadURL = [System.Web.HttpUtility]::UrlDecode( $Matches.Url)
                        $ValidUrl = Invoke-WebRequest -Uri $DownloadURL -Method HEAD -UseBasicParsing -DisableKeepAlive -MaximumRedirection 10 -ErrorAction SilentlyContinue
                    }
                }

                if ( $ValidUrl ) {
                    if ( $DownloadURL -like '*.pdf' -or $ValidURL.Headers.'Content-Type' -ieq 'application/pdf') {
                        # Slidedeck offered is PDF format
                        $slidedeckFile = '{0}.pdf' -f $FileName
                    }
                    else {
                        $slidedeckFile = '{0}.pptx' -f $FileName
                    }
                    $slidedeckFullFile = Join-Path -Path $DownloadFolder -ChildPath $slidedeckFile
                    if ((Test-Path -LiteralPath $slidedeckFullFile) -and ((Get-ChildItem -LiteralPath $slidedeckFullFile -ErrorAction SilentlyContinue).Length -gt 0) -and -not $Overwrite) {
                        Write-Host ('Slidedeck exists {0}' -f $slidedeckFile) -ForegroundColor Gray
                        $DeckInfo[ $InfoExist]++
                    }
                    else {
                        Write-Verbose ('Downloading {0} to {1}' -f $DownloadURL, $slidedeckFullFile)
                        Add-BackgroundDownloadJob -Type 1 -FilePath $slidedeckFullFile -DownloadUrl $DownloadURL -File $slidedeckFullFile -Timestamp $SessionTime -scheduleCode ($SessionToGet.sessioncode) -Title ($SessionToGet.Title)
                    }
                }
                else {
                    Write-Warning ('Skipping: Slidedeck unavailable {0}' -f $DownloadURL)
                    $DeckInfo[ $InfoPlaceholder]++
                }
            }
            else {
                Write-Warning ('No slidedeck link for {0}' -f ($SessionToGet.Title))
            }
        }
    }

    $JobsRunning = Get-BackgroundDownloadJobs

    if ([system.console]::KeyAvailable) {
        $key = [system.console]::readkey($true)
        if (($key.modifiers -band [consolemodifiers]"control") -and ($key.key -eq "C")) {
            Write-Host "TERMINATING" -ForegroundColor Red
            Stop-BackgroundDownloadJobs
            exit -1
        }
    }

}

$ProcessedSessions = $i

Write-Progress -Id 1 -Completed -Activity "Finished processing session information"

$JobsRunning = Get-BackgroundDownloadJobs
if ( $JobsRunning -gt 0) {
    Write-Host ('Waiting for download jobs to finish - press Ctrl-C once to abort)' -f $JobsRunning)
    while ( $JobsRunning -gt 0) {
        if ([system.console]::KeyAvailable) {
            Start-Sleep 1
            $key = [system.console]::readkey($true)
            if (($key.modifiers -band [consolemodifiers]"control") -and ($key.key -eq "C")) {
                Write-Host "TERMINATING" -ForegroundColor Red
                Stop-BackgroundDownloadJobs
                exit -1
            }
        }
        Start-Sleep 5
        $JobsRunning = Get-BackgroundDownloadJobs
    }
}
else {
    Write-Host ('Background download jobs have finished' -f $JobsRunning)
}

Write-Progress -Id 2 -Completed -Activity "Download jobs finished"

Write-Host ('Selected {0} sessions out of a total of {1}' -f $ProcessedSessions, $TotalNumberOfSessions)
Write-Host ('Downloaded {0} slide decks and {1} videos.' -f $DeckInfo[ $InfoDownload], $VideoInfo[ $InfoDownload])
Write-Host ('Not (yet) available: {0} slide decks and {1} videos' -f $DeckInfo[ $InfoPlaceholder], $VideoInfo[ $InfoPlaceholder])
Write-Host ('Skipped {0} slide decks and {1} videos as they were already downloaded.' -f $DeckInfo[ $InfoExist], $VideoInfo[ $InfoExist])

