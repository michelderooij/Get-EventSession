<#
    .SYNOPSIS
    Script to assist in downloading Microsoft Ignite, Inspire, Build, MEC or Custom event contents, or return
    session information for easier digesting. Video downloads will leverage external utilities,
    depending on the used video format. To prevent retrieving session information for every run,
    the script will cache session information.

    Note: -Language controls preferred audio track selection for videos. Caption language selection
    is controlled by -Subs when using -Captions.

    Be advised that downloading of OnDemand contents from Azure Media Services is throttled to real-time
    speed. To lessen the pain, the script performs simultaneous downloads of multiple videos streams. Those
    downloads will each open in their own (minimized) window so you can track progress. Finally, CTRL-C
    is catched by the script because we need to stop download jobs when aborting the script.

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Michel de Rooij
    http://eightwone.com
    Version 4.41, April 9, 2026

    Special thanks to: Mattias Fors, Scott Ladewig, Tim Pringle, Andy Race, Richard van Nieuwenhuizen

    .DESCRIPTION
    This script can download Microsoft Ignite, Inspire, Build, MEC and Custom event session information
    and available slidedecks and videos using event-specific catalog endpoints.

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
    translation. Note: For Azure Media Services and Custom Medius sessions when using -Captions, this
    controls caption language preference. When omitted, fallback preference is en-us/en.

    .PARAMETER Language
    When specified, for Azure Media hosted contents, downloads videos with specified audio stream where
    available. Note that if you mix this with specifying your own Format parameter, you need to
    add the language in the filter yourself, e.g. bestaudio[format_id*=German]. Default value is English,
    as otherwise YouTube will download the last audio stream from the manifest (which often is Spanish).
    This parameter controls audio stream preference only; it does not control caption language.

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
    - Custom                                       : Custom event endpoint requiring MSA login and page-based paging
    - Ignite                                       : Ignite contents (current)
    - Ignite2025                                    : Ignite contents from that year/time
    - Inspire                                      : Inspire contents (current)
    - Build                                        : Build contents (current)
    - Build2025                                    : Build contents from that year
    - MEC                                          : MEC contents (current)
    - MEC2022                                      : MEC contents from that year

    .PARAMETER OGVPicker
    Specify that you want to pick sessions to download using Out-GridView.

    .PARAMETER EventUrl
    URL to use for Custom events. The URL is automatically expanded with a page parameter.

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
    4.37  Fixed downloading to UNC paths
    4.38  Fixed CookiesFile parameter name
    4.39  Added UseSessionFolders switch to store content of each session in its own folder
    4.40  Added Custom event support with configurable EventUrl
          Added MSA authentication support to Custom Event for when needed
          Added MSA token caching mechanism
          Fixed terminate cleanup to stop yt-dlp child processes
          Some minor fixes for InfoOnly reliability and runtime summary handling
    4.41  Fixed compatibility issue with PowerShell v5.1
          Fixed processing non-Custom events

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
    [ValidateSet('MEC', 'MEC2022', 'Ignite', 'Ignite2025', 'Inspire', 'Build', 'Build2025', 'Custom')]
    [string]$Event = '',

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [parameter( Mandatory = $false, ParameterSetName = 'DownloadDirect')]
    [string]$EventUrl = $null,

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
    [switch]$UseSessionFolders,

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
    [string]$CookiesFile,

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [ValidateSet( 'brave', 'chrome', 'chromium', 'edge', 'firefox', 'opera', 'safari', 'vivaldi', 'whale' )]
    [string]$CookiesFromBrowser
)

# Max age for cache, older than this # hours will force info refresh
$script:MaxCacheAge = 24
# Max age for MSA auth token, expiring in less than these minutes will force re-authentication
$script:MinTokenValidityMinutes= 15

$script:YouTubeEXE = 'yt-dlp.exe'
$script:YouTubeDL = Join-Path $PSScriptRoot $YouTubeEXE
$script:FFMPEG = Join-Path $PSScriptRoot 'ffmpeg.exe'

$script:YTlink = 'https://github.com/yt-dlp/yt-dlp/releases/download/2023.07.06/yt-dlp.exe'
$script:FFMPEGlink = 'https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-essentials.zip'

# Fix 'Could not create SSL/TLS secure channel' issues with Invoke-WebRequest
[Net.ServicePointManager]::SecurityProtocol = "tls12, tls11, tls"

$script:BackgroundDownloadJobs = @()
$script:MSAAuthWebSession = $null
$script:MSAContentAuthRequired = $false
$script:CustomSignedOnDemandUrlCache = @{}
$script:CustomSignedOnDemandAuthRequiredBySession = @{}
$script:MSAAuthInteractiveAttemptedByHost = @{}
$script:SuppressWebRequestDebugDetails = $true

if ($script:SuppressWebRequestDebugDetails) {
    # Keep script-level debug traces while suppressing native HTTP request/response debug dumps.
    $PSDefaultParameterValues['Invoke-WebRequest:Debug'] = $false
    $PSDefaultParameterValues['Invoke-RestMethod:Debug'] = $false
}

# In Windows PowerShell 5.1, this avoids interactive HTML/script parsing prompts.
if ((Get-Command Invoke-WebRequest).Parameters.ContainsKey('UseBasicParsing')) {
    $PSDefaultParameterValues['Invoke-WebRequest:UseBasicParsing'] = $true
}

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

function Set-ObjectPropertyValue {
    param(
        [parameter(Mandatory = $true)]$Object,
        [parameter(Mandatory = $true)][string]$Name,
        $Value
    )

    if ($Object.PSObject.Properties.Match($Name).Count -gt 0) {
        $Object.$Name = $Value
    }
    else {
        $Object | Add-Member -MemberType NoteProperty -Name $Name -Value $Value
    }
}

function Get-CustomEventPageUri {
    param(
        [parameter(Mandatory = $true)][string]$BaseUrl,
        [parameter(Mandatory = $true)][ValidateRange(1, [int]::MaxValue)][int]$Page
    )

    $trimmedUrl = $BaseUrl.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmedUrl)) {
        throw 'EventUrl cannot be empty for Custom events.'
    }

    if ($trimmedUrl -match '\{0\}') {
        return ($trimmedUrl -f $Page)
    }

    if ($trimmedUrl -match '(?i)[\?&]page=') {
        return [regex]::Replace($trimmedUrl, '(?i)([\?&]page=)\d+', ('$1{0}' -f $Page))
    }

    if ($trimmedUrl.Contains('?')) {
        return ('{0}&page={1}' -f $trimmedUrl, $Page)
    }

    return ('{0}?page={1}' -f $trimmedUrl, $Page)
}

function Ensure-WinInetCookieInterop {
    if (-not ('WinInet.NativeMethods' -as [type])) {
        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace WinInet {
    public static class NativeMethods {
        [DllImport("wininet.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern bool InternetGetCookieEx(
            string lpszUrl,
            string lpszCookieName,
            StringBuilder lpszCookieData,
            ref int lpdwSize,
            int dwFlags,
            IntPtr lpReserved);
    }
}
"@
    }
}

function Get-UriCookieHeader {
    param(
        [parameter(Mandatory = $true)][uri]$Uri
    )

    Ensure-WinInetCookieInterop

    $cookieDataSize = 4096
    $cookieData = New-Object System.Text.StringBuilder $cookieDataSize
    $httpOnlyFlag = 0x00002000

    $isSuccess = [WinInet.NativeMethods]::InternetGetCookieEx(
        $Uri.AbsoluteUri,
        $null,
        $cookieData,
        [ref]$cookieDataSize,
        $httpOnlyFlag,
        [IntPtr]::Zero
    )

    if (-not $isSuccess -and $cookieDataSize -gt 0) {
        $cookieData = New-Object System.Text.StringBuilder $cookieDataSize
        $isSuccess = [WinInet.NativeMethods]::InternetGetCookieEx(
            $Uri.AbsoluteUri,
            $null,
            $cookieData,
            [ref]$cookieDataSize,
            $httpOnlyFlag,
            [IntPtr]::Zero
        )
    }

    if ($isSuccess) {
        return $cookieData.ToString()
    }

    return $null
}

function Merge-CookieHeaders {
    param(
        [string[]]$Headers
    )

    $entries = New-Object System.Collections.Generic.List[string]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($header in @($Headers)) {
        if ([string]::IsNullOrWhiteSpace($header)) {
            continue
        }

        foreach ($cookieEntry in ($header -split ';')) {
            $nameValue = $cookieEntry.Trim()
            if ([string]::IsNullOrWhiteSpace($nameValue)) {
                continue
            }

            $cookieParts = $nameValue -split '=', 2
            if ($cookieParts.Count -ne 2 -or [string]::IsNullOrWhiteSpace($cookieParts[0])) {
                continue
            }

            $cookieName = $cookieParts[0].Trim()
            if ($seen.Add($cookieName)) {
                $entries.Add(('{0}={1}' -f $cookieName, $cookieParts[1].Trim())) | Out-Null
            }
        }
    }

    if ($entries.Count -eq 0) {
        return $null
    }

    return ($entries -join '; ')
}

function Test-IsNetscapeCookieCacheContent {
    param(
        [string]$CacheContent
    )

    if ([string]::IsNullOrWhiteSpace($CacheContent)) {
        return $false
    }

    $lines = $CacheContent -split "`r?`n"
    if ($lines -match '^\s*#\s*Netscape HTTP Cookie File\s*$') {
        return $true
    }

    $firstDataLine = $lines |
    Where-Object { $_ -match '\S' -and $_ -notmatch '^\s*#' } |
    Select-Object -First 1

    if ([string]::IsNullOrWhiteSpace($firstDataLine)) {
        return $false
    }

    return [bool]($firstDataLine -match '^\S+\t(TRUE|FALSE)\t\S+\t(TRUE|FALSE)\t\d+\t\S+\t.*$')
}

function Convert-CacheContentToCookieHeader {
    param(
        [string]$CacheContent,
        [parameter(Mandatory = $true)][uri]$CookieUri
    )

    if ([string]::IsNullOrWhiteSpace($CacheContent)) {
        return $null
    }

    if (-not (Test-IsNetscapeCookieCacheContent -CacheContent $CacheContent)) {
        return ($CacheContent -replace "`r?`n", '').Trim()
    }

    $entries = New-Object System.Collections.Generic.List[string]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $targetHost = $CookieUri.Host.TrimStart('.')

    foreach ($rawLine in ($CacheContent -split "`r?`n")) {
        $line = $rawLine.Trim()
        if ([string]::IsNullOrWhiteSpace($line) -or $line.StartsWith('#')) {
            continue
        }

        $parts = $line -split "`t", 7
        if ($parts.Count -lt 7) {
            continue
        }

        $domain = [string]$parts[0]
        if ($domain.StartsWith('#HttpOnly_')) {
            $domain = $domain.Substring(10)
        }

        $normalizedDomain = $domain.Trim().TrimStart('.')
        if ([string]::IsNullOrWhiteSpace($normalizedDomain)) {
            continue
        }

        if (
            $normalizedDomain -ine $targetHost -and
            -not $targetHost.EndsWith(('.{0}' -f $normalizedDomain), [System.StringComparison]::OrdinalIgnoreCase) -and
            -not $normalizedDomain.EndsWith(('.{0}' -f $targetHost), [System.StringComparison]::OrdinalIgnoreCase)
        ) {
            continue
        }

        $cookieName = ([string]$parts[5]).Trim()
        if ([string]::IsNullOrWhiteSpace($cookieName)) {
            continue
        }

        $cookieValue = [string]$parts[6]
        if ($seen.Add($cookieName)) {
            $entries.Add(('{0}={1}' -f $cookieName, $cookieValue)) | Out-Null
        }
    }

    if ($entries.Count -eq 0) {
        return $null
    }

    return ($entries -join '; ')
}

function Convert-CookieHeaderToNetscapeCookieFileContent {
    param(
        [string]$CookieHeader,
        [parameter(Mandatory = $true)][uri]$CookieUri
    )

    if ([string]::IsNullOrWhiteSpace($CookieHeader)) {
        return $null
    }

    $cookieEntries = New-Object System.Collections.Generic.List[psobject]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($cookieEntry in ($CookieHeader -split ';')) {
        $nameValue = $cookieEntry.Trim()
        if ([string]::IsNullOrWhiteSpace($nameValue) -or -not $nameValue.Contains('=')) {
            continue
        }

        $parts = $nameValue -split '=', 2
        $cookieName = $parts[0].Trim()
        if ([string]::IsNullOrWhiteSpace($cookieName) -or -not $seen.Add($cookieName)) {
            continue
        }

        $cookieEntries.Add([pscustomobject]@{
                Name  = $cookieName
                Value = $parts[1]
            }) | Out-Null
    }

    if ($cookieEntries.Count -eq 0) {
        return $null
    }

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('# Netscape HTTP Cookie File') | Out-Null
    $lines.Add('# This file is generated by Get-EventSession for MSA authentication cache.') | Out-Null

    foreach ($cookie in $cookieEntries) {
        $lines.Add(("{0}`tFALSE`t/`tTRUE`t0`t{1}`t{2}" -f $CookieUri.Host, $cookie.Name, $cookie.Value)) | Out-Null
    }

    return ($lines -join [Environment]::NewLine)
}

function Add-CookieHeaderToWebSession {
    param(
        [parameter(Mandatory = $true)][Microsoft.PowerShell.Commands.WebRequestSession]$WebSession,
        [parameter(Mandatory = $true)][uri]$Uri,
        [parameter(Mandatory = $false)][string]$CookieHeader
    )

    if ([string]::IsNullOrWhiteSpace($CookieHeader)) {
        return
    }

    foreach ($cookieEntry in ($CookieHeader -split ';')) {
        $nameValue = $cookieEntry.Trim()
        if ([string]::IsNullOrWhiteSpace($nameValue)) {
            continue
        }

        $cookieParts = $nameValue -split '=', 2
        if ($cookieParts.Count -ne 2 -or [string]::IsNullOrWhiteSpace($cookieParts[0])) {
            continue
        }

        $cookie = New-Object System.Net.Cookie
        $cookie.Name = $cookieParts[0].Trim()
        $cookie.Value = $cookieParts[1].Trim()
        $cookie.Path = '/'
        $cookie.Domain = $Uri.Host

        try {
            $WebSession.Cookies.Add($Uri, $cookie)
        }
        catch {
            Write-Warning ('Skipping cookie {0}: {1}' -f $cookie.Name, $_.Exception.Message)
        }
    }
}

function New-WebSessionFromCookieHeader {
    param(
        [parameter(Mandatory = $true)][string]$CookieHeader,
        [parameter(Mandatory = $true)][uri]$CookieUri
    )

    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    Add-CookieHeaderToWebSession -WebSession $session -Uri $CookieUri -CookieHeader $CookieHeader
    Add-CookieHeaderToWebSession -WebSession $session -Uri ([uri]('{0}://{1}/' -f $CookieUri.Scheme, $CookieUri.Host)) -CookieHeader $CookieHeader
    return $session
}

function Get-WebSessionCookieHeader {
    param(
        [parameter(Mandatory = $true)][Microsoft.PowerShell.Commands.WebRequestSession]$WebSession,
        [parameter(Mandatory = $true)][uri]$CookieUri,
        [switch]$IncludeAllDomains
    )

    $entries = New-Object System.Collections.Generic.List[string]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    try {
        $cookieContainer = $WebSession.Cookies
        $domainTable = $cookieContainer.GetType().InvokeMember('m_domainTable', [System.Reflection.BindingFlags]'NonPublic,Instance,GetField', $null, $cookieContainer, @())
        $targetHost = $CookieUri.Host.TrimStart('.')

        foreach ($domainKey in @($domainTable.Keys)) {
            $domainName = [string]$domainKey
            if ([string]::IsNullOrWhiteSpace($domainName)) {
                continue
            }

            $normalizedDomain = $domainName.TrimStart('.')
            if (-not $IncludeAllDomains) {
                if ($normalizedDomain -ine $targetHost -and -not $targetHost.EndsWith(('.{0}' -f $normalizedDomain), [System.StringComparison]::OrdinalIgnoreCase) -and -not $normalizedDomain.EndsWith(('.{0}' -f $targetHost), [System.StringComparison]::OrdinalIgnoreCase)) {
                    continue
                }
            }

            $domainEntry = $domainTable[$domainKey]
            if (-not $domainEntry) {
                continue
            }

            $pathList = $domainEntry.GetType().InvokeMember('m_list', [System.Reflection.BindingFlags]'NonPublic,Instance,GetField', $null, $domainEntry, @())
            foreach ($pathKey in @($pathList.Keys)) {
                foreach ($cookie in @($pathList[$pathKey])) {
                    if (-not $cookie -or [string]::IsNullOrWhiteSpace([string]$cookie.Name)) {
                        continue
                    }

                    if ($seen.Add([string]$cookie.Name)) {
                        $entries.Add(('{0}={1}' -f [string]$cookie.Name, [string]$cookie.Value)) | Out-Null
                    }
                }
            }
        }
    }
    catch {
        Write-Verbose ('Unable to enumerate full cookie container for {0}: {1}' -f $CookieUri.Host, $_.Exception.Message)
    }

    if ($entries.Count -eq 0) {
        try {
            foreach ($cookie in @($WebSession.Cookies.GetCookies($CookieUri))) {
                if ($cookie -and -not [string]::IsNullOrWhiteSpace([string]$cookie.Name) -and $seen.Add([string]$cookie.Name)) {
                    $entries.Add(('{0}={1}' -f [string]$cookie.Name, [string]$cookie.Value)) | Out-Null
                }
            }
        }
        catch {
            Write-Verbose ('Unable to read host-scoped cookies from web session for {0}: {1}' -f $CookieUri.Host, $_.Exception.Message)
        }
    }

    if ($entries.Count -eq 0) {
        return $null
    }

    return ($entries -join '; ')
}

function Get-CachedMSAWebSession {
    param(
        [parameter(Mandatory = $true)][string]$CachePath,
        [parameter(Mandatory = $true)][uri]$CookieUri
    )

    if (-not (Test-Path -LiteralPath $CachePath)) {
        Write-Verbose ('Auth cache file not found: {0}' -f $CachePath)
        return $null
    }

    try {
        $cacheContent = Get-Content -LiteralPath $CachePath -Raw -ErrorAction Stop
    }
    catch {
        Write-Warning ('Unable to read auth cache file {0}: {1}' -f $CachePath, $_.Exception.Message)
        return $null
    }

    if ([string]::IsNullOrWhiteSpace($cacheContent)) {
        Write-Verbose ('Auth cache file is empty: {0}' -f $CachePath)
        return $null
    }

    Write-Verbose ('Loaded auth cache data from {0} (length {1})' -f $CachePath, $cacheContent.Length)

    $cookieHeader = Convert-CacheContentToCookieHeader -CacheContent $cacheContent -CookieUri $CookieUri
    if ([string]::IsNullOrWhiteSpace($cookieHeader)) {
        Write-Warning ('Auth cache file {0} did not contain usable cookies for {1}.' -f $CachePath, $CookieUri.Host)
        return $null
    }

    $session = New-WebSessionFromCookieHeader -CookieHeader $cookieHeader -CookieUri $CookieUri
    $cookieCount = ($session.Cookies.GetCookies($CookieUri) | Measure-Object).Count
    Write-Verbose ('Constructed cached web session with {0} cookies for {1}' -f $cookieCount, $CookieUri.Host)
    return $session
}

function Save-MSAAuthCookieHeader {
    param(
        [parameter(Mandatory = $true)][string]$CachePath,
        [parameter(Mandatory = $true)][uri]$CookieUri,
        [parameter(Mandatory = $false)][Microsoft.PowerShell.Commands.WebRequestSession]$WebSession,
        [parameter(Mandatory = $false)][string]$CookieHeader
    )

    $cookieHeaderCandidates = New-Object System.Collections.Generic.List[string]

    if (-not [string]::IsNullOrWhiteSpace($CookieHeader)) {
        $cookieHeaderCandidates.Add($CookieHeader) | Out-Null
    }

    if ($WebSession) {
        try {
            $sessionCookieHeader = Get-WebSessionCookieHeader -WebSession $WebSession -CookieUri $CookieUri
            if (-not [string]::IsNullOrWhiteSpace($sessionCookieHeader)) {
                $cookieHeaderCandidates.Add($sessionCookieHeader) | Out-Null
            }
        }
        catch {
            Write-Warning ('Unable to read cookies from in-memory web session for {0}: {1}' -f $CookieUri.Host, $_.Exception.Message)
        }
    }

    $uriCookieHeader = Get-UriCookieHeader -Uri $CookieUri
    if (-not [string]::IsNullOrWhiteSpace($uriCookieHeader)) {
        $cookieHeaderCandidates.Add($uriCookieHeader) | Out-Null
    }

    $cookieHeader = Merge-CookieHeaders -Headers @($cookieHeaderCandidates.ToArray())

    if (-not [string]::IsNullOrWhiteSpace($cookieHeader)) {
        $existingCookieHeader = $null
        if (Test-Path -LiteralPath $CachePath) {
            try {
                $existingCacheContent = Get-Content -LiteralPath $CachePath -Raw -ErrorAction Stop
                $existingCookieHeader = Convert-CacheContentToCookieHeader -CacheContent $existingCacheContent -CookieUri $CookieUri
            }
            catch {
                $existingCookieHeader = $null
            }
        }

        $cookieHeaderToPersist = if ([string]::IsNullOrWhiteSpace($existingCookieHeader)) {
            $cookieHeader
        }
        else {
            Merge-CookieHeaders -Headers @($cookieHeader, $existingCookieHeader)
        }

        if ([string]::IsNullOrWhiteSpace($cookieHeaderToPersist)) {
            $cookieHeaderToPersist = $cookieHeader
        }

        if (Test-Path -LiteralPath $CachePath) {
            try {
                Copy-Item -LiteralPath $CachePath -Destination ('{0}.bak' -f $CachePath) -Force -ErrorAction SilentlyContinue
            }
            catch {
            }
        }

        $cacheContentToPersist = Convert-CookieHeaderToNetscapeCookieFileContent -CookieHeader $cookieHeaderToPersist -CookieUri $CookieUri
        if ([string]::IsNullOrWhiteSpace($cacheContentToPersist)) {
            $cacheContentToPersist = $cookieHeaderToPersist
        }

        Set-Content -LiteralPath $CachePath -Value $cacheContentToPersist -Encoding UTF8
        Write-Verbose ('Saved MSA auth cache to {0} in Netscape cookie format (length {1})' -f $CachePath, $cacheContentToPersist.Length)
    }
    else {
        Write-Warning ('Unable to save MSA auth cache for {0}; cookie header was empty.' -f $CookieUri.Host)
    }
}

function Get-MSAAuthenticationLockName {
    param(
        [parameter(Mandatory = $true)][string]$HostName
    )

    $normalizedHost = ($HostName.ToLowerInvariant() -replace '[^a-z0-9]', '_')
    return ('Global\GetEventSession_MSAAuth_{0}' -f $normalizedHost)
}

function New-MSAAuthenticationMutex {
    param(
        [parameter(Mandatory = $true)][string]$HostName
    )

    $mutexName = Get-MSAAuthenticationLockName -HostName $HostName
    try {
        return New-Object System.Threading.Mutex($false, $mutexName)
    }
    catch {
        Write-Warning ('Unable to create global MSA auth mutex {0}: {1}. Falling back to local mutex.' -f $mutexName, $_.Exception.Message)
        return New-Object System.Threading.Mutex($false, ('Local\{0}' -f ($mutexName -replace '^Global\\', '')))
    }
}

function Get-CustomSessionPageUri {
    param(
        [parameter(Mandatory = $true)][string]$CatalogUrl,
        [parameter(Mandatory = $true)][string]$SessionCode
    )

    if ([string]::IsNullOrWhiteSpace($SessionCode)) {
        return $null
    }

    try {
        $catalogPageUri = [uri](Get-CustomEventPageUri -BaseUrl $CatalogUrl -Page 1)
    }
    catch {
        return $null
    }

    $rootUri = [uri]('{0}://{1}/' -f $catalogPageUri.Scheme, $catalogPageUri.Host)

    $candidates = [System.Collections.ArrayList]@()
    $null = $candidates.Add(([uri]::new($rootUri, ('sessions/{0}' -f $SessionCode))).AbsoluteUri)

    $catalogPath = $catalogPageUri.AbsolutePath.Trim('/')
    if (-not [string]::IsNullOrWhiteSpace($catalogPath) -and $catalogPath -match '(?i)^(?<prefix>.*?)/api(?:/.*)?$') {
        $pathPrefix = $Matches.prefix.Trim('/')
        if (-not [string]::IsNullOrWhiteSpace($pathPrefix)) {
            $null = $candidates.Add(([uri]::new($rootUri, ('{0}/sessions/{1}' -f $pathPrefix, $SessionCode))).AbsoluteUri)
        }
    }

    foreach ($candidate in ($candidates | Select-Object -Unique)) {
        try {
            return [uri]$candidate
        }
        catch {
        }
    }

    return $null
}

function Test-MSAAuthenticationRequired {
    param(
        $Response = $null,
        [System.Exception]$Exception = $null
    )

    $statusCode = $null
    $location = $null
    $responseUri = $null
    $responseContent = $null

    if ($Response) {
        if ($Response.PSObject.Properties.Match('StatusCode').Count -gt 0) {
            try {
                $statusCode = [int]$Response.StatusCode
            }
            catch {
                $statusCode = $null
            }
        }
        if ($Response.PSObject.Properties.Match('Headers').Count -gt 0 -and $Response.Headers) {
            try {
                $location = [string]($Response.Headers.Location | Select-Object -First 1)
            }
            catch {
                $location = $null
            }
        }

        if ($Response.PSObject.Properties.Match('BaseResponse').Count -gt 0 -and $Response.BaseResponse -and $Response.BaseResponse.ResponseUri) {
            try {
                $responseUri = [string]$Response.BaseResponse.ResponseUri.AbsoluteUri
            }
            catch {
                $responseUri = $null
            }
        }
        elseif ($Response.PSObject.Properties.Match('ResponseUri').Count -gt 0 -and $Response.ResponseUri) {
            try {
                $responseUri = [string]$Response.ResponseUri.AbsoluteUri
            }
            catch {
                $responseUri = $null
            }
        }

        if ($Response.PSObject.Properties.Match('Content').Count -gt 0 -and -not [string]::IsNullOrWhiteSpace([string]$Response.Content)) {
            try {
                $responseContent = [string]$Response.Content
            }
            catch {
                $responseContent = $null
            }
        }
    }

    if (-not $statusCode -and $Exception -and $Exception.Response) {
        try {
            $statusCode = [int]$Exception.Response.StatusCode
        }
        catch {
            $statusCode = $null
        }

        try {
            $location = [string]($Exception.Response.Headers.Location | Select-Object -First 1)
        }
        catch {
            $location = $null
        }
    }

    if ($statusCode -in @(401, 403)) {
        return $true
    }

    if ($statusCode -in @(301, 302, 307, 308) -and $location -match '(?i)(login|signin|oauth|microsoftonline\.com|live\.com)') {
        return $true
    }

    if ($responseUri -match '(?i)(login\.microsoftonline\.com|login\.live\.com|msauth)') {
        return $true
    }

    if ($responseContent -match '(?i)<title>\s*sign in to your account\s*</title>|PageID"\s+content="ConvergedSignIn"') {
        return $true
    }

    if ($Exception -and $Exception.Message -match '(?i)(401|403|unauthori[sz]ed|forbidden|authentication required|sign.?in)') {
        return $true
    }

    return $false
}

function Get-MSAAuthenticatedWebSession {
    param(
        [parameter(Mandatory = $true)][uri]$TargetUri,
        [uri]$Proxy,
        [switch]$ValidateCachedSession,
        [switch]$ForceInteractive
    )

    $authCachePath = Join-Path $PSScriptRoot 'MSAAuth.cache'
    $cookieUri = [uri]('{0}://{1}/' -f $TargetUri.Scheme, $TargetUri.Host)
    $hostKey = $cookieUri.Host.ToLowerInvariant()
    $authMutex = $null
    $lockAcquired = $false

    if ($ForceInteractive) {
        $script:MSAAuthWebSession = $null
    }

    if ($script:MSAAuthWebSession) {
        if (-not $ValidateCachedSession) {
            return $script:MSAAuthWebSession
        }

        try {
            $inMemoryValidationResponse = Invoke-WebRequest -Uri $TargetUri.AbsoluteUri -Method Get -WebSession $script:MSAAuthWebSession -Proxy $Proxy -Headers @{ 'Accept' = 'application/json'; 'Content-Type' = 'application/json; charset=utf-8' } -ErrorAction Stop -Verbose:$false
            if (Test-MSAAuthenticationRequired -Response $inMemoryValidationResponse) {
                throw 'In-memory session redirected to a sign-in page.'
            }
            Write-Verbose ('Using in-memory MSA authentication for {0}' -f $cookieUri.Host)
            return $script:MSAAuthWebSession
        }
        catch {
            Write-Verbose ('In-memory MSA authentication for {0} appears invalid; reacquiring session.' -f $cookieUri.Host)
            $script:MSAAuthWebSession = $null
        }
    }

    try {
        $authMutex = New-MSAAuthenticationMutex -HostName $cookieUri.Host

        if ($authMutex) {
            try {
                $lockAcquired = $authMutex.WaitOne(0)
            }
            catch [System.Threading.AbandonedMutexException] {
                $lockAcquired = $true
                Write-Warning 'Detected abandoned MSA authentication lock; continuing with lock ownership.'
            }

            if (-not $lockAcquired) {
                Write-Host ('MSA authentication for {0} is in progress by another run; waiting for completion ...' -f $cookieUri.Host)
                try {
                    $authMutex.WaitOne() | Out-Null
                    $lockAcquired = $true
                }
                catch [System.Threading.AbandonedMutexException] {
                    $lockAcquired = $true
                    Write-Warning 'Detected abandoned MSA authentication lock while waiting; continuing with lock ownership.'
                }
            }
        }

        # Re-check in-memory session after waiting for lock owner to finish.
        if ($script:MSAAuthWebSession) {
            return $script:MSAAuthWebSession
        }

        $session = Get-CachedMSAWebSession -CachePath $authCachePath -CookieUri $cookieUri

        if ($session -and $ValidateCachedSession) {
            try {
                $validationResponse = Invoke-WebRequest -Uri $TargetUri.AbsoluteUri -Method Get -WebSession $session -Proxy $Proxy -Headers @{ 'Accept' = 'application/json'; 'Content-Type' = 'application/json; charset=utf-8' } -ErrorAction Stop -Verbose:$false
                if (Test-MSAAuthenticationRequired -Response $validationResponse) {
                    throw 'Cached session redirected to a sign-in page.'
                }
                Write-Host 'Using cached MSA authentication from MSAAuth.cache'
            }
            catch {
                Write-Warning 'Cached MSA authentication appears invalid; interactive sign-in is required.'
                $session = $null
            }
        }

        if (-not $session) {
            Write-Verbose ('Attempting silent MSA sign-in bootstrap for {0}' -f $cookieUri.Host)
            $session = Get-MSAAuthWebSessionSilent -StartUri $TargetUri -CookieUri $cookieUri -Proxy $Proxy
            if ($session) {
                $silentSessionValid = $true
                $silentCookieCount = 0
                try {
                    $silentValidationResponse = Invoke-WebRequest -Uri $TargetUri.AbsoluteUri -Method Get -WebSession $session -Proxy $Proxy -Headers @{ 'Accept' = 'application/json'; 'Content-Type' = 'application/json; charset=utf-8' } -ErrorAction Stop -Verbose:$false
                    if (Test-MSAAuthenticationRequired -Response $silentValidationResponse) {
                        throw 'Silent session redirected to a sign-in page.'
                    }
                }
                catch {
                    $silentSessionValid = $false
                    Write-Verbose ('Silent MSA sign-in bootstrap for {0} did not produce an authenticated session; falling back to interactive sign-in.' -f $cookieUri.Host)
                    $silentCookieHeader = Get-WebSessionCookieHeader -WebSession $session -CookieUri $cookieUri -IncludeAllDomains
                    if (-not [string]::IsNullOrWhiteSpace($silentCookieHeader)) {
                        $silentCookieCount = ($silentCookieHeader -split ';').Count
                    }
                }

                if ($silentSessionValid) {
                    Write-Verbose ('Acquired MSA authentication silently for {0}' -f $cookieUri.Host)
                    Save-MSAAuthCookieHeader -CachePath $authCachePath -CookieUri $cookieUri -WebSession $session
                }
                elseif ($silentCookieCount -gt 0) {
                    Write-Verbose ('Captured {0} silent-session cookie(s) for {1}, but validation still redirected; continuing to interactive sign-in.' -f $silentCookieCount, $cookieUri.Host)
                    $session = $null
                }
                else {
                    $session = $null
                }
            }

            if (-not $session) {
                Write-Verbose ('Silent MSA sign-in bootstrap did not yield an authenticated session for {0}' -f $cookieUri.Host)
                if ($script:MSAAuthInteractiveAttemptedByHost.ContainsKey($hostKey) -and $script:MSAAuthInteractiveAttemptedByHost[$hostKey]) {
                    Write-Warning ('MSA interactive sign-in for {0} was already attempted in this run and no reusable cache/session is available; skipping repeated prompt.' -f $cookieUri.Host)
                    return $null
                }

                $script:MSAAuthInteractiveAttemptedByHost[$hostKey] = $true
                Write-Host 'Opening Microsoft sign-in dialog (embedded browser) for event authentication'
                $session = Get-MSAAuthWebSession -StartUri $TargetUri -CookieUri $cookieUri -Proxy $Proxy

                $interactiveSessionValid = $true
                $interactiveCookieCount = 0
                try {
                    $interactiveValidationResponse = Invoke-WebRequest -Uri $TargetUri.AbsoluteUri -Method Get -WebSession $session -Proxy $Proxy -Headers @{ 'Accept' = 'application/json'; 'Content-Type' = 'application/json; charset=utf-8' } -ErrorAction Stop -Verbose:$false
                    if (Test-MSAAuthenticationRequired -Response $interactiveValidationResponse) {
                        throw 'Interactive session redirected to a sign-in page.'
                    }
                }
                catch {
                    $interactiveSessionValid = $false
                    Write-Warning ('Interactive MSA sign-in for {0} did not produce an authenticated session.' -f $cookieUri.Host)
                    $interactiveCookieHeader = Get-WebSessionCookieHeader -WebSession $session -CookieUri $cookieUri -IncludeAllDomains
                    if (-not [string]::IsNullOrWhiteSpace($interactiveCookieHeader)) {
                        $interactiveCookieCount = ($interactiveCookieHeader -split ';').Count
                    }
                }

                if (-not $interactiveSessionValid) {
                    if ($interactiveCookieCount -gt 0) {
                        Write-Verbose ('Proceeding with {0} captured interactive-session cookie(s) for {1}; validation endpoint still redirected.' -f $interactiveCookieCount, $cookieUri.Host)
                    }
                    else {
                        return $null
                    }
                }

                Save-MSAAuthCookieHeader -CachePath $authCachePath -CookieUri $cookieUri -WebSession $session
            }
        }

        $script:MSAAuthWebSession = $session
        return $session
    }
    finally {
        if ($lockAcquired -and $authMutex) {
            try {
                $authMutex.ReleaseMutex() | Out-Null
            }
            catch {
            }
        }
        if ($authMutex) {
            $authMutex.Dispose()
        }
    }
}

function Invoke-WebRequestWithMSAAuthSupport {
    param(
        [parameter(Mandatory = $true)][string]$Uri,
        [parameter(Mandatory = $true)][ValidateSet('Get', 'Head', 'Post')][string]$Method,
        [uri]$Proxy,
        [hashtable]$Headers = $null,
        [string]$OutFile = $null,
        [int]$MaximumRedirection = 10,
        [switch]$DisableKeepAlive
    )

    if ([string]::IsNullOrWhiteSpace($Uri)) {
        return $null
    }

    try {
        [uri]$requestUri = $Uri
    }
    catch {
        Write-Warning ('Skipping invalid URI: {0}' -f $Uri)
        return $null
    }

    $invokeParams = @{
        Uri         = $Uri
        Method      = $Method
        ErrorAction = 'Stop'
        Verbose     = $false
    }
    if ($Proxy) { $invokeParams.Proxy = $Proxy }
    if ($Headers) { $invokeParams.Headers = $Headers }
    if (-not [string]::IsNullOrWhiteSpace($OutFile)) { $invokeParams.OutFile = $OutFile }
    if ($DisableKeepAlive) { $invokeParams.DisableKeepAlive = $true }
    if ($MaximumRedirection -ge 0) { $invokeParams.MaximumRedirection = $MaximumRedirection }

    if ($script:MSAContentAuthRequired -and $script:MSAAuthWebSession) {
        $invokeParams.WebSession = $script:MSAAuthWebSession
    }

    if (-not $script:SuppressWebRequestDebugDetails) {
        $requestUriText = $requestUri.AbsoluteUri
        $requestQueryLength = $requestUri.Query.TrimStart('?').Length
        if (-not [string]::IsNullOrWhiteSpace($OutFile)) {
            Write-Debug ('WebRequest: {0} {1} with query length {2} output to {3}' -f $Method.ToUpperInvariant(), $requestUri.GetLeftPart([System.UriPartial]::Path), $requestQueryLength, $OutFile)
        }
        elseif ($requestQueryLength -gt 0) {
            Write-Debug ('WebRequest: {0} {1} with query length {2}' -f $Method.ToUpperInvariant(), $requestUri.GetLeftPart([System.UriPartial]::Path), $requestQueryLength)
        }
        else {
            Write-Debug ('WebRequest: {0} {1}' -f $Method.ToUpperInvariant(), $requestUriText)
        }
    }

    try {
        $response = Invoke-WebRequest @invokeParams

        if (Test-MSAAuthenticationRequired -Response $response) {
            throw [System.Exception]::new('Authentication required - sign-in response detected.')
        }

        $responseContentType = $null
        if ($response -and $response.PSObject.Properties.Match('Headers').Count -gt 0 -and $response.Headers) {
            $responseContentType = [string]$response.Headers['Content-Type']
        }
        if (-not $script:SuppressWebRequestDebugDetails -and $response -and $response.PSObject.Properties.Match('StatusCode').Count -gt 0) {
            Write-Debug ('WebResponse: {0} {1}{2}' -f [int]$response.StatusCode, [string]$response.StatusDescription, (Iif -Cond ([string]::IsNullOrWhiteSpace($responseContentType)) -IfTrue '' -IfFalse (' with {0} payload' -f $responseContentType)))
        }
        return $response
    }
    catch {
        $webException = $_.Exception
        if (-not (Test-MSAAuthenticationRequired -Exception $webException)) {
            return $null
        }

        if (-not $script:MSAContentAuthRequired) {
            Write-Verbose ('Detected authentication requirement for {0}; acquiring reusable MSA session.' -f $requestUri.Host)
        }
        $script:MSAContentAuthRequired = $true

        $needsInteractive = [bool]$script:MSAAuthWebSession
        $session = Get-MSAAuthenticatedWebSession -TargetUri $requestUri -Proxy $Proxy -ForceInteractive:$needsInteractive
        if (-not $session) {
            Write-Warning ('Unable to acquire authenticated MSA session for {0}; skipping request.' -f $requestUri.Host)
            return $null
        }
        $invokeParams.WebSession = $session

        try {
            $response = Invoke-WebRequest @invokeParams
            $responseContentType = $null
            if ($response -and $response.PSObject.Properties.Match('Headers').Count -gt 0 -and $response.Headers) {
                $responseContentType = [string]$response.Headers['Content-Type']
            }
            if (-not $script:SuppressWebRequestDebugDetails -and $response -and $response.PSObject.Properties.Match('StatusCode').Count -gt 0) {
                Write-Debug ('WebResponse: {0} {1}{2}' -f [int]$response.StatusCode, [string]$response.StatusDescription, (Iif -Cond ([string]::IsNullOrWhiteSpace($responseContentType)) -IfTrue '' -IfFalse (' with {0} payload' -f $responseContentType)))
            }
            return $response
        }
        catch {
            $script:MSAAuthWebSession = $null
            Write-Warning ('Authenticated request failed for {0}: {1}' -f $Uri, $_.Exception.Message)
            return $null
        }
    }
}

function New-MSAWebSessionFromBrowserContext {
    param(
        [parameter(Mandatory = $true)][uri]$CookieUri,
        [parameter(Mandatory = $true)][uri]$StartUri,
        [parameter(Mandatory = $true)]$Browser,
        $VisitedUrls,
        [uri]$Proxy
    )

    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $cookieHeaderCandidates = New-Object System.Collections.Generic.List[string]

    try {
        if ($Browser.Document -and -not [string]::IsNullOrWhiteSpace($Browser.Document.Cookie)) {
            $documentCookieHeader = $Browser.Document.Cookie
            $cookieHeaderCandidates.Add($documentCookieHeader) | Out-Null
            Add-CookieHeaderToWebSession -WebSession $session -Uri $CookieUri -CookieHeader $documentCookieHeader
            Add-CookieHeaderToWebSession -WebSession $session -Uri ([uri]('{0}://{1}/' -f $CookieUri.Scheme, $CookieUri.Host)) -CookieHeader $documentCookieHeader
            Write-Verbose ('Captured browser document cookies for {0} (length {1})' -f $CookieUri.Host, $documentCookieHeader.Length)
        }
    }
    catch {
        Write-Warning ('Unable to read browser document cookies for {0}: {1}' -f $CookieUri.Host, $_.Exception.Message)
    }

    $cookieUris = @(
        $CookieUri,
        [uri]('{0}://{1}/' -f $CookieUri.Scheme, $CookieUri.Host),
        $StartUri
    ) | Select-Object -Unique

    if ($Browser.Url -and -not [string]::IsNullOrWhiteSpace($Browser.Url.AbsoluteUri)) {
        try {
            $cookieUris += [uri]$Browser.Url.AbsoluteUri
        }
        catch {
        }
    }

    foreach ($visitedUrl in @($VisitedUrls)) {
        try {
            $visitedUri = [uri]$visitedUrl
            $cookieUris += $visitedUri
        }
        catch {
        }
    }

    $cookieUris += @(
        [uri]'https://login.live.com/',
        [uri]'https://login.microsoftonline.com/',
        [uri]'https://account.microsoft.com/'
    )

    $cookieUris = @($cookieUris | Select-Object -Unique)
    Write-Verbose ('Collecting authentication cookies from {0} URL(s) for {1}' -f $cookieUris.Count, $CookieUri.Host)

    foreach ($cookieTarget in $cookieUris) {
        $cookieHeader = Get-UriCookieHeader -Uri $cookieTarget
        if (-not [string]::IsNullOrWhiteSpace($cookieHeader)) {
            $cookieHeaderCandidates.Add($cookieHeader) | Out-Null
        }
        Add-CookieHeaderToWebSession -WebSession $session -Uri $cookieTarget -CookieHeader $cookieHeader
    }

    $mergedCookieHeader = Merge-CookieHeaders -Headers @($cookieHeaderCandidates.ToArray())
    if (-not [string]::IsNullOrWhiteSpace($mergedCookieHeader)) {
        Add-CookieHeaderToWebSession -WebSession $session -Uri $CookieUri -CookieHeader $mergedCookieHeader
        Add-CookieHeaderToWebSession -WebSession $session -Uri ([uri]('{0}://{1}/' -f $CookieUri.Scheme, $CookieUri.Host)) -CookieHeader $mergedCookieHeader
    }

    $hostCookieCount = ($session.Cookies.GetCookies($CookieUri) | Measure-Object).Count
    $allCookieHeader = Get-WebSessionCookieHeader -WebSession $session -CookieUri $CookieUri -IncludeAllDomains
    $allCookieCount = 0
    if (-not [string]::IsNullOrWhiteSpace($allCookieHeader)) {
        $allCookieCount = ($allCookieHeader -split ';').Count
    }

    if ($hostCookieCount -eq 0 -and $allCookieCount -eq 0) {
        $finalBrowserUrl = if ($Browser.Url) { $Browser.Url.AbsoluteUri } else { 'unknown' }
        throw ('No cookies were captured for {0} (final browser URL: {1}). Authentication may have failed.' -f $CookieUri.Host, $finalBrowserUrl)
    }

    Write-Verbose ('Constructed browser-backed session with {0} host cookie(s) and {1} total cookie(s).' -f $hostCookieCount, $allCookieCount)

    try {
        Invoke-WebRequest -Method Head -Uri $CookieUri -WebSession $session -Proxy $Proxy -ErrorAction SilentlyContinue -Verbose:$false | Out-Null
    }
    catch {
        Write-Warning ('Session warmup request returned: {0}' -f $_.Exception.Message)
    }

    return $session
}

function Get-MSAAuthWebSessionSilent {
    param(
        [parameter(Mandatory = $true)][uri]$StartUri,
        [parameter(Mandatory = $true)][uri]$CookieUri,
        [uri]$Proxy,
        [int]$TimeoutSeconds = 8
    )

    if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
        return $null
    }

    if (-not [Environment]::UserInteractive) {
        return $null
    }

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $state = [hashtable]::Synchronized(@{ Completed = $false })
    $visitedUrls = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $form = $null
    $browser = $null

    try {
        $form = New-Object System.Windows.Forms.Form
        $form.ShowInTaskbar = $false
        $form.StartPosition = 'Manual'
        $form.Location = New-Object System.Drawing.Point(-32000, -32000)
        $form.Width = 1
        $form.Height = 1
        $form.Opacity = 0

        $browser = New-Object System.Windows.Forms.WebBrowser
        $browser.Dock = 'Fill'
        $browser.ScriptErrorsSuppressed = $true
        $browser.add_DocumentCompleted({
                param($browserSender, $e)
                if ($browserSender.Url) {
                    $null = $visitedUrls.Add($browserSender.Url.AbsoluteUri)
                    if ($browserSender.Url.Host -ieq $CookieUri.Host) {
                        $state.Completed = $true
                    }
                }
            })

        $form.Controls.Add($browser)
        $form.Show()
        $browser.Navigate($StartUri.AbsoluteUri)

        $deadline = (Get-Date).AddSeconds([Math]::Max(3, $TimeoutSeconds))
        while ((Get-Date) -lt $deadline -and -not $state.Completed) {
            [System.Windows.Forms.Application]::DoEvents()
            [System.Threading.Thread]::Sleep(100)
        }

        if (-not $state.Completed) {
            return $null
        }

        return New-MSAWebSessionFromBrowserContext -CookieUri $CookieUri -StartUri $StartUri -Browser $browser -VisitedUrls $visitedUrls -Proxy $Proxy
    }
    catch {
        Write-Verbose ('Silent MSA sign-in bootstrap failed: {0}' -f $_.Exception.Message)
        return $null
    }
    finally {
        if ($form) {
            try {
                $form.Close()
            }
            catch {
            }
            try {
                $form.Dispose()
            }
            catch {
            }
        }
    }
}

function Get-MSAAuthWebSession {
    param(
        [parameter(Mandatory = $true)][uri]$StartUri,
        [parameter(Mandatory = $true)][uri]$CookieUri,
        [uri]$Proxy
    )

    if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
        throw 'MSA sign-in UI requires an STA runspace. Start PowerShell with -STA and retry.'
    }

    if (-not [Environment]::UserInteractive) {
        throw 'MSA sign-in UI requires an interactive desktop session. Run the script in an interactive PowerShell terminal.'
    }

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Microsoft Account Sign-In'
    $form.Width = 1200
    $form.Height = 900
    $form.StartPosition = 'CenterScreen'
    $form.ShowInTaskbar = $true
    $form.TopMost = $true

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Dock = 'Top'
    $statusLabel.Height = 30
    $statusLabel.Text = 'Sign in with your Microsoft account and click Continue when the event page is loaded.'

    $buttonPanel = New-Object System.Windows.Forms.Panel
    $buttonPanel.Dock = 'Bottom'
    $buttonPanel.Height = 44

    $continueButton = New-Object System.Windows.Forms.Button
    $continueButton.Text = 'Continue'
    $continueButton.Width = 100
    $continueButton.Height = 28
    $continueButton.Left = 980
    $continueButton.Top = 8
    $continueButton.Enabled = $false

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = 'Cancel'
    $cancelButton.Width = 100
    $cancelButton.Height = 28
    $cancelButton.Left = 1090
    $cancelButton.Top = 8

    $browser = New-Object System.Windows.Forms.WebBrowser
    $browser.Dock = 'Fill'
    $browser.ScriptErrorsSuppressed = $true
    $visitedUrls = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    $continueButton.Add_Click({
            $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $form.Close()
        })

    $cancelButton.Add_Click({
            $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $form.Close()
        })

    $browser.add_DocumentCompleted({
            param($browserSender, $e)
            if ($browserSender.Url) {
                $null = $visitedUrls.Add($browserSender.Url.AbsoluteUri)
                $statusLabel.Text = ('Current page: {0}' -f $browserSender.Url.AbsoluteUri)
                if ($browserSender.Url.Host -ieq $CookieUri.Host) {
                    $continueButton.Enabled = $true
                }
            }
        })

    $buttonPanel.Controls.Add($continueButton)
    $buttonPanel.Controls.Add($cancelButton)
    $form.Controls.Add($browser)
    $form.Controls.Add($buttonPanel)
    $form.Controls.Add($statusLabel)

    $form.Add_Shown({
            $form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
            $form.Activate()
            $form.BringToFront()
        })

    Write-Host ('Launching Microsoft sign-in dialog in process {0}. If it is not visible, use Alt+Tab to switch to "Microsoft Account Sign-In".' -f $PID)
    $browser.Navigate($StartUri.AbsoluteUri)
    $dialogResult = $form.ShowDialog()

    if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
        throw 'MSA authentication cancelled.'
    }

    return New-MSAWebSessionFromBrowserContext -CookieUri $CookieUri -StartUri $StartUri -Browser $browser -VisitedUrls $visitedUrls -Proxy $Proxy
}

function Get-CustomEventCatalog {
    param(
        [parameter(Mandatory = $true)][string]$CatalogUrl,
        [uri]$Proxy
    )

    $catalogUri = [uri](Get-CustomEventPageUri -BaseUrl $CatalogUrl -Page 1)
    $session = Get-MSAAuthenticatedWebSession -TargetUri $catalogUri -Proxy $Proxy -ValidateCachedSession

    $data = [System.Collections.ArrayList]@()
    $seenSessionKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet', [string[]]('sessionCode', 'title'))
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)

    $page = 1
    $totalPages = $null

    while ($true) {
        $pageUri = Get-CustomEventPageUri -BaseUrl $CatalogUrl -Page $page

        $status = 'Processing page {0}' -f $page
        $percent = 0
        if ($totalPages) {
            $status = 'Processing page {0} of {1}' -f $page, $totalPages
            $percent = [Math]::Min(100, [int](($page / $totalPages) * 100))
        }
        Write-Progress -Id 1 -Activity 'Retrieving Session Catalog' -Status $status -PercentComplete $percent

        try {
            $response = Invoke-RestMethod -Uri $pageUri -Method Get -WebSession $session -Proxy $Proxy -Headers @{ 'Accept' = 'application/json'; 'Content-Type' = 'application/json; charset=utf-8' }
        }
        catch {
            throw ('Problem retrieving custom session catalog page {0}: {1}' -f $page, $error[0])
        }

        if (-not $totalPages -and $response.PSObject.Properties.Match('totalPages').Count -gt 0) {
            $totalPages = [int]$response.totalPages
        }

        $items = @($response.items)
        if ($items.Count -eq 0) {
            break
        }

        foreach ($item in $items) {
            $item.PSObject.Properties | ForEach-Object {
                if (@('speakerNames') -icontains $_.Name) {
                    $item.($_.Name) = @($_.Value)
                }
                if (@('products', 'contentCategory') -icontains $_.Name) {
                    $item.($_.Name) = @($_.Value -replace [char]9, '/')
                }
                if (@('topic', 'sessionType', 'sessionLevel', 'audienceTypes', 'deliveryTypes', 'viewingOptions', 'event', 'programmingLanguages') -icontains $_.Name) {
                    $item.($_.Name) = $_.Value.displayValue -join ','
                }
            }

            $canonicalCode = $null
            if ($item.PSObject.Properties.Match('scheduleCode').Count -gt 0 -and -not [string]::IsNullOrWhiteSpace([string]$item.scheduleCode)) {
                $canonicalCode = [string]$item.scheduleCode
            }
            elseif ($item.PSObject.Properties.Match('code').Count -gt 0 -and -not [string]::IsNullOrWhiteSpace([string]$item.code)) {
                $canonicalCode = [string]$item.code
            }

            if ($canonicalCode) {
                Set-ObjectPropertyValue -Object $item -Name 'sessionCode' -Value $canonicalCode
                Set-ObjectPropertyValue -Object $item -Name 'scheduleCode' -Value $canonicalCode
                Set-ObjectPropertyValue -Object $item -Name 'code' -Value $canonicalCode
            }

            $canonicalCategory = $null
            if ($item.PSObject.Properties.Match('contentCategory').Count -gt 0 -and -not [string]::IsNullOrWhiteSpace([string]$item.contentCategory)) {
                $canonicalCategory = [string]$item.contentCategory
            }
            elseif ($item.PSObject.Properties.Match('subcategoriesString').Count -gt 0 -and -not [string]::IsNullOrWhiteSpace([string]$item.subcategoriesString)) {
                $canonicalCategory = [string]$item.subcategoriesString
            }

            if ($canonicalCategory) {
                Set-ObjectPropertyValue -Object $item -Name 'contentCategory' -Value $canonicalCategory
            }

            if ($item.PSObject.Properties.Match('langLocale').Count -eq 0 -or [string]::IsNullOrWhiteSpace([string]$item.langLocale)) {
                Set-ObjectPropertyValue -Object $item -Name 'langLocale' -Value 'en-US'
            }

            $displayCode = if ($item.PSObject.Properties.Match('sessionCode').Count -gt 0) { [string]$item.sessionCode } else { [string]$item.id }
            Write-Verbose ('Adding info for session {0}' -f $displayCode)

            $item.PSObject.TypeNames.Insert(0, 'Session.Information')
            $item | Add-Member MemberSet PSStandardMembers $PSStandardMembers -Force

            $dedupeKey = if (-not [string]::IsNullOrWhiteSpace($canonicalCode)) {
                $canonicalCode
            }
            elseif ($item.PSObject.Properties.Match('id').Count -gt 0 -and -not [string]::IsNullOrWhiteSpace([string]$item.id)) {
                [string]$item.id
            }
            else {
                '{0}-{1}' -f $page, [string]$item.title
            }

            if ($seenSessionKeys.Add($dedupeKey)) {
                $data.Add($item) | Out-Null
            }
        }

        if ($totalPages -and $page -ge $totalPages) {
            break
        }

        $page++
    }

    Write-Progress -Id 1 -Completed -Activity 'Finished retrieval of catalog'
    return $data
}

function Resolve-CustomSignedOnDemandUrl {
    param(
        [parameter(Mandatory = $true)]$Session,
        [parameter(Mandatory = $true)][AllowNull()][AllowEmptyString()][string]$OnDemandUrl,
        [parameter(Mandatory = $true)][string]$CatalogUrl,
        [uri]$Proxy
    )

    if ([string]::IsNullOrWhiteSpace($OnDemandUrl)) {
        return $OnDemandUrl
    }

    if ($OnDemandUrl -notmatch '(?i)^https://medius\.microsoft\.com/Embed/video-aes/') {
        return $OnDemandUrl
    }

    if ($OnDemandUrl -match '(?i)([?&]uid=|[?&]at=)') {
        return $OnDemandUrl
    }

    $sessionCode = $null
    foreach ($candidate in @('scheduleCode', 'sessionCode', 'code')) {
        if ($Session.PSObject.Properties.Match($candidate).Count -gt 0 -and -not [string]::IsNullOrWhiteSpace([string]$Session.$candidate)) {
            $sessionCode = [string]$Session.$candidate
            break
        }
    }

    if ([string]::IsNullOrWhiteSpace($sessionCode)) {
        return $OnDemandUrl
    }

    if ($script:CustomSignedOnDemandAuthRequiredBySession.ContainsKey($sessionCode)) {
        $null = $script:CustomSignedOnDemandAuthRequiredBySession.Remove($sessionCode)
    }

    if ($script:CustomSignedOnDemandUrlCache.ContainsKey($sessionCode)) {
        return $script:CustomSignedOnDemandUrlCache[$sessionCode]
    }

    try {
        $sessionPageUri = Get-CustomSessionPageUri -CatalogUrl $CatalogUrl -SessionCode $sessionCode
        if (-not $sessionPageUri) {
            return $OnDemandUrl
        }
        $webSession = Get-MSAAuthenticatedWebSession -TargetUri $sessionPageUri -Proxy $Proxy -ValidateCachedSession
        if (-not $webSession) {
            $script:CustomSignedOnDemandAuthRequiredBySession[$sessionCode] = $true
            Write-Warning ('Unable to acquire authenticated web session for {0}; cannot resolve signed media URL.' -f $sessionCode)
            return $OnDemandUrl
        }
        $sessionPage = Invoke-WebRequest -Uri $sessionPageUri.AbsoluteUri -Method Get -WebSession $webSession -Proxy $Proxy -DisableKeepAlive -ErrorAction Stop -Verbose:$false
    }
    catch {
        Write-Warning ('Unable to retrieve session page for {0}: {1}' -f $sessionCode, $_.Exception.Message)
        return $OnDemandUrl
    }

    if (Test-MSAAuthenticationRequired -Response $sessionPage) {
        Write-Verbose ('Session page for {0} returned a sign-in response; forcing interactive authentication and retrying.' -f $sessionCode)
        $script:MSAAuthWebSession = $null

        try {
            $webSession = Get-MSAAuthenticatedWebSession -TargetUri $sessionPageUri -Proxy $Proxy -ValidateCachedSession -ForceInteractive
            if (-not $webSession) {
                $script:CustomSignedOnDemandAuthRequiredBySession[$sessionCode] = $true
                Write-Warning ('Unable to acquire authenticated web session for {0} after sign-in response; cannot resolve signed media URL.' -f $sessionCode)
                return $OnDemandUrl
            }

            $sessionPage = Invoke-WebRequest -Uri $sessionPageUri.AbsoluteUri -Method Get -WebSession $webSession -Proxy $Proxy -DisableKeepAlive -ErrorAction Stop -Verbose:$false
        }
        catch {
            Write-Warning ('Unable to retrieve authenticated session page for {0}: {1}' -f $sessionCode, $_.Exception.Message)
            return $OnDemandUrl
        }

        if (Test-MSAAuthenticationRequired -Response $sessionPage) {
            $script:CustomSignedOnDemandAuthRequiredBySession[$sessionCode] = $true
            Write-Warning ('Session page for {0} still resolves to sign-in after re-authentication; signed media URL cannot be resolved.' -f $sessionCode)
            return $OnDemandUrl
        }
    }

    $pageContent = $sessionPage.Content
    if ([string]::IsNullOrWhiteSpace($pageContent)) {
        return $OnDemandUrl
    }

    $videoId = $null
    $videoMatch = [regex]::Match($OnDemandUrl, '(?i)/video-aes/(?<id>[0-9a-f\-]+)')
    if ($videoMatch.Success) {
        $videoId = [regex]::Escape($videoMatch.Groups['id'].Value)
    }

    $signedUrlPattern = if ($videoId) {
        '(?i)https://medius\.microsoft\.com/Embed/video-aes/{0}\?[^"''\s<]+' -f $videoId
    }
    else {
        '(?i)https://medius\.microsoft\.com/Embed/video-aes/[^"''\s<]+\?[^"''\s<]+'
    }

    $signedUrlMatch = [regex]::Match($pageContent, $signedUrlPattern)
    if (-not $signedUrlMatch.Success) {
        $signedUrlEscapedPattern = if ($videoId) {
            '(?i)https:\\/\\/medius\.microsoft\.com\\/Embed\\/video-aes\\/{0}\\?[^"''\s<]+' -f $videoId
        }
        else {
            '(?i)https:\\/\\/medius\.microsoft\.com\\/Embed\\/video-aes\\/[^"''\s<]+\\?[^"''\s<]+'
        }
        $signedUrlMatch = [regex]::Match($pageContent, $signedUrlEscapedPattern)
    }

    if (-not $signedUrlMatch.Success) {
        return $OnDemandUrl
    }

    $resolvedUrlRaw = $signedUrlMatch.Value
    if ($resolvedUrlRaw -match '\\/|\\u[0-9a-fA-F]{4}') {
        try {
            $resolvedUrlRaw = [regex]::Unescape($resolvedUrlRaw)
        }
        catch {
        }
    }

    $resolvedUrl = [System.Net.WebUtility]::HtmlDecode($resolvedUrlRaw)
    if ($resolvedUrl -match '(?i)[?&]uid=' -and $resolvedUrl -match '(?i)[?&]at=') {
        $script:CustomSignedOnDemandUrlCache[$sessionCode] = $resolvedUrl
        Write-Verbose ('Resolved signed Custom on-demand URL for session {0}' -f $sessionCode)
        return $resolvedUrl
    }

    return $OnDemandUrl
}

function Resolve-OnDemandManifestUrlFromCoreConfiguration {
    param(
        [string]$OnDemandPage
    )

    if ([string]::IsNullOrWhiteSpace($OnDemandPage)) {
        return $null
    }

    $coreConfigurationMatch = [regex]::Match($OnDemandPage, '(?s)let\s+coreConfiguration\s*=\s*(?<json>\{.*?\})\s*;')
    if (-not $coreConfigurationMatch.Success) {
        return $null
    }

    try {
        $coreConfiguration = ($coreConfigurationMatch.Groups['json'].Value | ConvertFrom-Json -Depth 100)
    }
    catch {
        Write-Warning ('Unable to parse coreConfiguration JSON from embed page: {0}' -f $_.Exception.Message)
        return $null
    }

    $mainManifests = @($coreConfiguration.manifests.main)
    if ($mainManifests.Count -eq 0) {
        return $null
    }

    $manifestCandidates = $mainManifests |
    Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.manifest) } |
    Sort-Object -Property @{ Expression = { [int](Iif -Cond ($_.PSObject.Properties.Match('weight').Count -gt 0) -IfTrue $_.weight -IfFalse 0) }; Descending = $true }

    if ($manifestCandidates.Count -eq 0) {
        return $null
    }

    $selectedManifest = [string]($manifestCandidates | Select-Object -First 1).manifest
    if ($selectedManifest -match '(?i)^https://') {
        Write-Debug ('Resolved coreConfiguration manifest URL {0}' -f $selectedManifest)
        return $selectedManifest
    }

    return $null
}

function Resolve-OnDemandCaptionsConfiguration {
    param(
        [parameter(Mandatory = $true)][string]$OnDemandPage
    )

    if ([string]::IsNullOrWhiteSpace($OnDemandPage)) {
        return $null
    }

    $captionsConfigurationPatterns = @(
        '(?s)(?:let|const|var)\s+captionsConfiguration\s*=\s*(?<json>\{.*?\})\s*;',
        '(?s)captionsConfiguration\s*=\s*(?<json>\{.*?\})\s*;'
    )

    foreach ($pattern in $captionsConfigurationPatterns) {
        $captionsConfigurationMatch = [regex]::Match($OnDemandPage, $pattern)
        if (-not $captionsConfigurationMatch.Success) {
            continue
        }

        try {
            $captionsConfiguration = ($captionsConfigurationMatch.Groups['json'].Value | ConvertFrom-Json -Depth 100)
        }
        catch {
            Write-Warning ('Unable to parse captionsConfiguration JSON from embed page: {0}' -f $_.Exception.Message)
            continue
        }

        $languageList = @($captionsConfiguration.languageList)
        if ($languageList.Count -gt 0) {
            Write-Debug ('Resolved captionsConfiguration with {0} language entries' -f $languageList.Count)
            return $languageList
        }
    }

    return $null
}

function Get-PreferredOnDemandYtDlpFormat {
    param(
        [string]$OnDemandPage,
        [string]$Endpoint,
        [string]$FallbackFormat = 'worstvideo+bestaudio'
    )

    if ([string]::IsNullOrWhiteSpace($OnDemandPage) -or [string]::IsNullOrWhiteSpace($Endpoint)) {
        return $FallbackFormat
    }

    $hasAdaptiveFormats = $false

    $coreConfigurationMatch = [regex]::Match($OnDemandPage, '(?s)let\s+coreConfiguration\s*=\s*(?<json>\{.*?\})\s*;')
    if ($coreConfigurationMatch.Success) {
        try {
            $coreConfiguration = ($coreConfigurationMatch.Groups['json'].Value | ConvertFrom-Json -Depth 100)
            $mainManifests = @($coreConfiguration.manifests.main) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.manifest) }
            if ($mainManifests.Count -gt 0) {
                $hasAdaptiveFormats = $true
            }
        }
        catch {
            $hasAdaptiveFormats = $false
        }
    }

    if (-not $hasAdaptiveFormats) {
        $hasAdaptiveFormats = [bool]($OnDemandPage -match '(?i)(manifests"\s*:\s*\{\s*"main"|master\.m3u8|EXT-X-STREAM-INF)')
    }

    if (-not $hasAdaptiveFormats) {
        return $FallbackFormat
    }

    if ($Endpoint -match '(?i)(\.m3u8($|\?)|stream\.event\.microsoft\.com)') {
        return 'bestvideo+bestaudio/best'
    }

    return $FallbackFormat
}

function Get-OnDemandKmsBearerToken {
    param(
        [parameter(Mandatory = $true)][string]$OnDemandPage,
        [parameter(Mandatory = $true)][string]$OnDemandUrl,
        [parameter(Mandatory = $true)][string]$ManifestUrl,
        [uri]$Proxy
    )

    if ([string]::IsNullOrWhiteSpace($OnDemandPage) -or [string]::IsNullOrWhiteSpace($OnDemandUrl) -or [string]::IsNullOrWhiteSpace($ManifestUrl)) {
        return $null
    }

    if ($ManifestUrl -notmatch '(?i)^https://stream\.event\.microsoft\.com/') {
        return $null
    }

    $videoId = [regex]::Match($OnDemandPage, 'id="hdnVideoId"\s+value="(?<v>[^"]+)"').Groups['v'].Value
    if ([string]::IsNullOrWhiteSpace($videoId)) {
        $videoId = [regex]::Match($OnDemandPage, 'const\s+id\s*=\s*''(?<v>[^'']+)'';').Groups['v'].Value
    }

    $refreshToken = [regex]::Match($OnDemandPage, 'var\s+refreshToken\s*=\s*"(?<v>[^"]+)";').Groups['v'].Value
    $channelKey = [regex]::Match($OnDemandPage, 'let\s+channelKey\s*=\s*"(?<v>[^"]+)";').Groups['v'].Value
    $channelGuid = [regex]::Match($OnDemandPage, 'let\s+channelGuid\s*=\s*"(?<v>[^"]+)";').Groups['v'].Value

    if ([string]::IsNullOrWhiteSpace($videoId) -or [string]::IsNullOrWhiteSpace($refreshToken) -or [string]::IsNullOrWhiteSpace($channelKey) -or [string]::IsNullOrWhiteSpace($channelGuid)) {
        return $null
    }

    $assetId = $null
    $arrAssetsMatch = [regex]::Match($OnDemandPage, '(?s)const\s+arrAssets\s*=\s*(?<json>\[.*?\]);')
    if ($arrAssetsMatch.Success) {
        try {
            $assets = $arrAssetsMatch.Groups['json'].Value | ConvertFrom-Json -Depth 50
            $videoAsset = @($assets | Where-Object { $_.AssetType -eq 'Video' }) | Select-Object -First 1
            if ($videoAsset -and -not [string]::IsNullOrWhiteSpace([string]$videoAsset.AssetId)) {
                $assetId = [string]$videoAsset.AssetId
            }
            elseif ($videoAsset -and -not [string]::IsNullOrWhiteSpace([string]$videoAsset.BlobName)) {
                $assetId = [string]$videoAsset.BlobName
            }
        }
        catch {
            Write-Warning ('Unable to parse arrAssets JSON for KMS token request: {0}' -f $_.Exception.Message)
        }
    }

    if ([string]::IsNullOrWhiteSpace($assetId)) {
        $assetId = [regex]::Match($OnDemandPage, '"AssetType":"Video"[^\}]*?"AssetId":"(?<v>[^"]+)"').Groups['v'].Value
    }

    if ([string]::IsNullOrWhiteSpace($assetId)) {
        return $null
    }

    $origin = if ($ManifestUrl -match '(?i)https://stream\.event\.microsoft\.com/prodnc') { 'origin2' } else { 'origin1' }

    $kmsUrl = 'https://medius.microsoft.com/Embed/GetKMSToken/{0}?rt={1}&origin={2}&channelKey={3}&assetId={4}&channelGuid={5}' -f $videoId, [uri]::EscapeDataString($refreshToken), [uri]::EscapeDataString($origin), [uri]::EscapeDataString($channelKey), [uri]::EscapeDataString($assetId), [uri]::EscapeDataString($channelGuid)
    Write-Debug ('Attempting KMS token retrieval for video {0} (origin={1})' -f $videoId, $origin)

    try {
        $mediusSession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        $embedRequestParams = @{
            Uri         = $OnDemandUrl
            Method      = 'Get'
            WebSession  = $mediusSession
            ErrorAction = 'Stop'
            Verbose     = $false
        }
        if ($Proxy) { $embedRequestParams.Proxy = $Proxy }
        Invoke-WebRequest @embedRequestParams | Out-Null

        $kmsRequestParams = @{
            Uri         = $kmsUrl
            Method      = 'Get'
            WebSession  = $mediusSession
            Headers     = @{ Referer = $OnDemandUrl; 'User-Agent' = 'Mozilla/5.0' }
            ErrorAction = 'Stop'
            Verbose     = $false
        }
        if ($Proxy) { $kmsRequestParams.Proxy = $Proxy }

        $kmsResponse = Invoke-RestMethod @kmsRequestParams
        if ($kmsResponse -and -not [string]::IsNullOrWhiteSpace([string]$kmsResponse.KMSToken)) {
            Write-Debug 'KMS token retrieval succeeded'
            return [string]$kmsResponse.KMSToken
        }
    }
    catch {
        Write-Warning ('Unable to retrieve KMS token: {0}' -f $_.Exception.Message)
    }

    return $null
}

function Get-KmsBearerTokenDetails {
    param(
        [string]$KmsBearerTokenRaw
    )

    if ([string]::IsNullOrWhiteSpace($KmsBearerTokenRaw)) {
        return [pscustomobject]@{
            RawToken            = $null
            AuthorizationToken  = $null
            ExpiryUtc           = $null
            HasParsedJson       = $false
        }
    }

    $authorizationToken = $KmsBearerTokenRaw
    $expiryUtc = $null
    $hasParsedJson = $false

    if ($KmsBearerTokenRaw.TrimStart().StartsWith('{')) {
        try {
            $tokenJson = $KmsBearerTokenRaw | ConvertFrom-Json -Depth 20
            $hasParsedJson = $true

            if ($tokenJson.PSObject.Properties.Match('token').Count -gt 0 -and -not [string]::IsNullOrWhiteSpace([string]$tokenJson.token)) {
                $authorizationToken = [string]$tokenJson.token
            }

            if ($tokenJson.PSObject.Properties.Match('exp').Count -gt 0 -and $null -ne $tokenJson.exp) {
                $expValue = $tokenJson.exp
                if ($expValue -is [int] -or $expValue -is [long] -or $expValue -is [double] -or $expValue -is [decimal]) {
                    $expiryUtc = [DateTimeOffset]::FromUnixTimeSeconds([int64]$expValue).UtcDateTime
                }
                else {
                    $expString = [string]$expValue
                    if ($expString -match '^\d+$') {
                        $expiryUtc = [DateTimeOffset]::FromUnixTimeSeconds([int64]$expString).UtcDateTime
                    }
                    else {
                        try {
                            $parsedExp = [DateTimeOffset]::Parse($expString)
                            $expiryUtc = $parsedExp.UtcDateTime
                        }
                        catch {
                            $expiryUtc = $null
                        }
                    }
                }
            }
        }
        catch {
            $hasParsedJson = $false
            $authorizationToken = $KmsBearerTokenRaw
            $expiryUtc = $null
        }
    }

    return [pscustomobject]@{
        RawToken            = $KmsBearerTokenRaw
        AuthorizationToken  = $authorizationToken
        ExpiryUtc           = $expiryUtc
        HasParsedJson       = $hasParsedJson
    }
}

function Test-IsMicrosoftContentUrl {
    param(
        [string]$Url
    )

    if ([string]::IsNullOrWhiteSpace($Url)) {
        return $false
    }

    try {
        [uri]$parsedUri = $Url
    }
    catch {
        return $false
    }

    return [bool]($parsedUri.Host -match '(?i)(^|\.)microsoft\.com$')
}

function Test-IsNetscapeCookieFile {
    param(
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return $false
    }

    try {
        $sampleLines = Get-Content -LiteralPath $Path -TotalCount 25 -ErrorAction Stop
    }
    catch {
        return $false
    }

    $firstDataLine = $sampleLines |
    Where-Object { $_ -match '\S' -and $_ -notmatch '^\s*#' } |
    Select-Object -First 1

    if ($sampleLines -match '^\s*#\s*Netscape HTTP Cookie File\s*$') {
        return $true
    }

    if ([string]::IsNullOrWhiteSpace($firstDataLine)) {
        return $false
    }

    return [bool]($firstDataLine -match '^\S+\s+(TRUE|FALSE)\s+\S+\s+(TRUE|FALSE)\s+\d+\s+\S+\s+.*$')
}

function Resolve-YtDlpCookieFile {
    param(
        [string]$PreferredCookieFile
    )

    $candidateFiles = [System.Collections.ArrayList]@()
    if (-not [string]::IsNullOrWhiteSpace($PreferredCookieFile)) {
        $null = $candidateFiles.Add($PreferredCookieFile)
    }

    $defaultCookieCandidates = @(
        Join-Path $PSScriptRoot 'cookie.txt',
        Join-Path $PSScriptRoot 'cookies.txt'
    )

    foreach ($candidate in $defaultCookieCandidates) {
        if ($candidateFiles -notcontains $candidate) {
            $null = $candidateFiles.Add($candidate)
        }
    }

    foreach ($candidate in $candidateFiles) {
        if (Test-IsNetscapeCookieFile -Path $candidate) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }

    return $null
}

function Resolve-CaptionSourceByPreferredLanguage {
    param(
        $LanguageList,
        [string[]]$PreferredLanguages
    )

    if (-not $LanguageList) {
        return $null
    }

    $captions = @($LanguageList | Where-Object {
            -not [string]::IsNullOrWhiteSpace($_.src) -and -not [string]::IsNullOrWhiteSpace($_.srclang)
        })

    if (-not $captions) {
        return $null
    }

    $languageCandidates = [System.Collections.ArrayList]@()

    foreach ($language in $PreferredLanguages) {
        if ([string]::IsNullOrWhiteSpace($language)) {
            continue
        }

        $normalizedLanguage = $language.Trim().ToLowerInvariant().Replace('_', '-')
        if (-not [string]::IsNullOrWhiteSpace($normalizedLanguage) -and $languageCandidates -notcontains $normalizedLanguage) {
            $null = $languageCandidates.Add($normalizedLanguage)
        }

        if ($normalizedLanguage -match '-') {
            $baseLanguage = $normalizedLanguage.Split('-')[0]
            if (-not [string]::IsNullOrWhiteSpace($baseLanguage) -and $languageCandidates -notcontains $baseLanguage) {
                $null = $languageCandidates.Add($baseLanguage)
            }
        }
    }

    foreach ($candidate in $languageCandidates) {
        $matchingCaption = $captions | Where-Object {
            $captionLanguage = ([string]$_.srclang).ToLowerInvariant().Replace('_', '-')
            ($captionLanguage -eq $candidate) -or ($captionLanguage.Split('-')[0] -eq $candidate)
        } | Select-Object -First 1

        if ($matchingCaption) {
            $selectedCaptionLanguage = ([string]$matchingCaption.srclang).ToLowerInvariant().Replace('_', '-')
            return [PSCustomObject]@{
                Src      = $matchingCaption.src
                Language = $selectedCaptionLanguage
            }
        }
    }

    return $null
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

function Stop-ProcessTree {
    param(
        [int]$ProcessId
    )

    if ($ProcessId -le 0) {
        return
    }

    if (-not (Get-Process -Id $ProcessId -ErrorAction SilentlyContinue)) {
        return
    }

    if ($IsWindows) {
        try {
            & taskkill.exe /PID $ProcessId /T /F | Out-Null
            return
        }
        catch {
            Write-Warning ('taskkill failed for PID {0}: {1}' -f $ProcessId, $_.Exception.Message)
        }
    }

    try {
        Stop-Process -Id $ProcessId -Force -ErrorAction SilentlyContinue
    }
    catch {
        Write-Warning ('Stop-Process failed for PID {0}: {1}' -f $ProcessId, $_.Exception.Message)
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

                if ( $job.Type -eq 2) {
                    $VideoInfo[ $InfoDownload]++
                }

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
                Stop-ProcessTree -ProcessId ([int]$BGJob.job.id)
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
Write-Host( 'Get-EventSession v4.40')
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
    { 'Custom' -contains $_ } {
        $EventName = 'Custom'
        $EventType = 'CUSTOM'
        $EventAPIUrl = $EventUrl
        $CaptionExt = 'vtt'
        $PreferDirect = $True
    }
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
    { 'Inspire' -contains $_ } {
        $EventName = 'Inspire'
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
    default {
        Write-Host ('Unknown event: {0}' -f $Event) -ForegroundColor Red
        exit -1
    }
}

if (-not ($InfoOnly)) {

    # If no download folder set, use system drive root for Custom events,
    # and system drive with event subfolder for other events.
    if ( -not( $DownloadFolder)) {
        if ($EventType -eq 'CUSTOM') {
            $DownloadFolder = '{0}\' -f $ENV:SystemDrive
        }
        else {
            $DownloadFolder = '{0}\{1}' -f $ENV:SystemDrive, $EventName
        }
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
        'CUSTOM' {
            Write-Host ('Reading {0} session catalog' -f $EventName)
            $data = Get-CustomEventCatalog -CatalogUrl $EventAPIUrl -Proxy $ProxyURL

            [int32]$sessionCount = ($data | Measure-Object).Count
            Write-Host ('Processing information for {0} sessions' -f $sessionCount)
        }
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

if ($Captions -and -not $Subs -and $PSBoundParameters.ContainsKey('Language') -and -not [string]::IsNullOrWhiteSpace([string]$Language)) {
    Write-Warning ('Language parameter controls audio stream selection only; use Subs to select caption language preferences.')
}

# Initialize counters used by both download and info-only flows.
$i = 0
$DeckInfo = @(0, 0, 0)
$VideoInfo = @(0, 0, 0)
$InfoDownload = 0
$InfoPlaceholder = 1
$InfoExist = 2
$SessionsSelected = ($SessionsToGet | Measure-Object).Count
$cachedKmsBearerTokenRaw = $null
$cachedKmsBearerTokenDetails = $null

$myTimeZone = $null
try {
    $myTimeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById('US Eastern Standard Time')
}
catch {
    $myTimeZone = [System.TimeZoneInfo]::Local
    Write-Warning ('Unable to resolve US Eastern Standard Time, using local timezone {0}' -f $myTimeZone.Id)
}

if ( $InfoOnly) {
    # Info-only mode should never queue downloads.
    $NoVideos = $true
    $NoSlidedecks = $true
    Write-Verbose ('There are {0} sessions matching your criteria.' -f $SessionsSelected)
    Write-Output $SessionsToGet
}
else {

    if ( $OGVPicker) {
        $SessionsToGet = $SessionsToGet | Out-GridView -Title 'Select Videos to Download, or Cancel for all Videos' -PassThru
    }

    [console]::TreatControlCAsInput = $true

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
    $ProgressPercent = if ( $SessionsSelected -gt 0 ) { ($i / $SessionsSelected * 100) } else { 0 }
    Write-Progress -Id 1 -Activity 'Inspecting session information' -Status "Processing session $i of $SessionsSelected" -PercentComplete $ProgressPercent
    if ( $SessionToGet.sessionCode) {
        $FileName = Fix-FileName ('{0}-{1}' -f $SessionToGet.sessionCode.Trim(), $SessionToGet.title.Trim())
    }
    else {
        $FileName = Fix-FileName ('{0}' -f $SessionToGet.title.Trim())
    }
    if (! ([string]::IsNullOrEmpty( $SessionToGet.startDateTime) -and [string]::IsNullOrWhiteSpace( $SessionToGet.startDateTime)) ) {
        try {
            # Get session localized timestamp, undoing TZ adjustments
            $SessionTime = [System.TimeZoneInfo]::ConvertTime((Get-Date -Date $SessionToGet.startDateTime).ToUniversalTime(), $myTimeZone ).toString('g')
        }
        catch {
            $SessionTime = $null
            Write-Warning ('Unable to convert startDateTime for {0}: {1}' -f $SessionToGet.sessionCode, $_.Exception.Message)
        }
    }
    else {
        $SessionTime = $null
    }
    Write-Host ('Processing info session {0} from {1} [{2}]' -f $FileName, (Iif -Cond $SessionTime -IfTrue $SessionTime -IfFalse '[No Timestamp]'), $SessionToGet.langLocale)
    if (!([string]::IsNullOrEmpty( $SessionToGet.startDateTime)) -and (Get-Date -Date $SessionToGet.startDateTime) -ge (Get-Date)) {
        Write-Verbose ('Skipping session {0}: Future session' -f $SessionToGet.sessioncode)
    }
    else {

        # When storing session content in subfolders per session, override the content target folder to be the session subfolder
        if ( $UseSessionFolders) {
            $SessionFolder = Join-Path -Path $DownloadFolder -ChildPath $FileName
            if ( (Test-Path $SessionFolder) -eq $false ) {
                New-Item -Path $SessionFolder -ItemType Directory | Out-Null
            }
            $ContentTargetFolder = $SessionFolder
        }
        else {
            $ContentTargetFolder = $DownloadFolder
        }

        if ( ! $NoVideos) {

            $onDemandPage = $null

            if ( $DownloadVideos -or $DownloadAMSVideos) {

                $vidfileName = '{0}.mp4' -f $FileName
                $vidFullFile = [System.IO.Path]::Combine( $ContentTargetFolder, $vidfileName)

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

                    if ($EventType -eq 'CUSTOM') {
                        $downloadLink = Resolve-CustomSignedOnDemandUrl -Session $SessionToGet -OnDemandUrl $downloadLink -CatalogUrl $EventAPIUrl -Proxy $ProxyURL
                    }

                    $customSessionCode = $null
                    foreach ($codeCandidate in @('scheduleCode', 'sessionCode', 'code')) {
                        if ($SessionToGet.PSObject.Properties.Match($codeCandidate).Count -gt 0 -and -not [string]::IsNullOrWhiteSpace([string]$SessionToGet.$codeCandidate)) {
                            $customSessionCode = [string]$SessionToGet.$codeCandidate
                            break
                        }
                    }

                    $customAuthRequired = $false
                    if ($EventType -eq 'CUSTOM' -and -not [string]::IsNullOrWhiteSpace($customSessionCode)) {
                        $customAuthRequired = [bool]($script:CustomSignedOnDemandAuthRequiredBySession.ContainsKey($customSessionCode) -and $script:CustomSignedOnDemandAuthRequiredBySession[$customSessionCode])
                    }

                    if ($customAuthRequired) {
                        Write-Warning ('Skipping: Authentication required for Custom session {0}. Complete Microsoft sign-in and retry.' -f $customSessionCode)
                        $Endpoint = $null
                        $Arg = $null
                        $downloadLink = $null
                    }

                    try {
                        $Response = Invoke-WebRequestWithMSAAuthSupport -Method Head -Uri $downloadLink -Proxy $ProxyURL -DisableKeepAlive
                        $DirectLink = @( 'video/mp4', 'video/MP2T') -contains $Response.Headers.'Content-Type'
                    }
                    catch {
                        $DirectLink = $False
                    }

                    $kmsBearerToken = $null
                    $kmsBearerAuthorizationToken = $null

                    if ( ! ( $DirectLink) -and $downloadLink -match '(medius\.studios\.ms\/Embed\/Video|medius\.microsoft\.com|mediastream\.microsoft\.com)' ) {
                        $DownloadedPage = Invoke-WebRequestWithMSAAuthSupport -Method Get -Uri $downloadLink -Proxy $ProxyURL -DisableKeepAlive
                        $OnDemandPage = $DownloadedPage.Content
                        $Endpoint = $null

                        if ( $OnDemandPage -match 'StreamUrl = "(?<AzureStreamURL>https://mediusprod\.streaming\.mediaservices\.windows\.net/.*manifest)";') {
                            Write-Verbose ('Using Azure Media Services URL {0}' -f $matches.AzureStreamURL)
                            $Endpoint = '{0}(format=mpd-time-csf)' -f $matches.AzureStreamURL
                        }
                        if ( $OnDemandPage -match 'StreamUrl = "(?<AzureStreamURL>https://stream\.event\.microsoft\.com/.*master\.m3u8)";') {
                            Write-Verbose ('Using Azure Media Stream URL {0}' -f $matches.AzureStreamURL)
                            $Endpoint = '{0}?(format=mpd-time-csf)' -f $matches.AzureStreamURL
                        }

                        if (-not $Endpoint) {
                            $CoreConfigurationManifestUrl = Resolve-OnDemandManifestUrlFromCoreConfiguration -OnDemandPage $OnDemandPage
                            if ($CoreConfigurationManifestUrl) {
                                Write-Debug ('Using coreConfiguration manifest URL {0}' -f $CoreConfigurationManifestUrl)
                                $Endpoint = $CoreConfigurationManifestUrl
                            }
                        }

                        if ($Endpoint) {
                            $kmsBearerToken = $null
                            if ( $Endpoint -match '(?i)^https://stream\.event\.microsoft\.com/') {
                                $hasActiveMsaSession = ($null -ne $script:MSAAuthWebSession)
                                $forceFreshKmsTokenPerCustomDownload = ($EventType -eq 'CUSTOM' -and ($script:MSAContentAuthRequired -or $hasActiveMsaSession))
                                $shouldFetchKmsToken = $true
                                if ($forceFreshKmsTokenPerCustomDownload) {
                                    $customSessionCodeForLog = if ([string]::IsNullOrWhiteSpace($customSessionCode)) { '<unknown>' } else { $customSessionCode }
                                    Write-Debug ('Requesting fresh KMS bearer token for Custom session {0} because MSA authentication is required.' -f $customSessionCodeForLog)
                                }
                                elseif (-not [string]::IsNullOrWhiteSpace($cachedKmsBearerTokenRaw)) {
                                    if ($null -eq $cachedKmsBearerTokenDetails) {
                                        $cachedKmsBearerTokenDetails = Get-KmsBearerTokenDetails -KmsBearerTokenRaw $cachedKmsBearerTokenRaw
                                    }

                                    if ($cachedKmsBearerTokenDetails -and $cachedKmsBearerTokenDetails.ExpiryUtc) {
                                        $refreshThresholdUtc = (Get-Date).ToUniversalTime().AddMinutes( $MinTokenValidityMinutes)
                                        if ($cachedKmsBearerTokenDetails.ExpiryUtc -gt $refreshThresholdUtc) {
                                            $shouldFetchKmsToken = $false
                                            Write-Debug ('Reusing cached KMS bearer token (expires {0:u})' -f $cachedKmsBearerTokenDetails.ExpiryUtc)
                                        }
                                        else {
                                            Write-Debug ('Cached KMS bearer token expires soon ({0:u}); requesting a new token.' -f $cachedKmsBearerTokenDetails.ExpiryUtc)
                                        }
                                    }
                                    else {
                                        # Token was previously fetched, but expiry is unavailable; keep using it unless retrieval fails.
                                        $shouldFetchKmsToken = $false
                                        Write-Debug 'Reusing previously fetched KMS bearer token (no parseable expiry found).'
                                    }
                                }

                                if ($shouldFetchKmsToken) {
                                    $kmsBearerToken = Get-OnDemandKmsBearerToken -OnDemandPage $OnDemandPage -OnDemandUrl $downloadLink -ManifestUrl $Endpoint -Proxy $ProxyURL
                                    if (-not [string]::IsNullOrWhiteSpace($kmsBearerToken)) {
                                        $cachedKmsBearerTokenRaw = $kmsBearerToken
                                        $cachedKmsBearerTokenDetails = Get-KmsBearerTokenDetails -KmsBearerTokenRaw $kmsBearerToken
                                    }
                                }
                                else {
                                    $kmsBearerToken = $cachedKmsBearerTokenRaw
                                }

                                if ($kmsBearerToken) {
                                    if ($null -eq $cachedKmsBearerTokenDetails -or $cachedKmsBearerTokenDetails.RawToken -ne $kmsBearerToken) {
                                        $cachedKmsBearerTokenDetails = Get-KmsBearerTokenDetails -KmsBearerTokenRaw $kmsBearerToken
                                    }
                                    $kmsBearerAuthorizationToken = $cachedKmsBearerTokenDetails.AuthorizationToken
                                    Write-Debug 'Resolved KMS bearer token for encrypted stream.event manifest'
                                }
                                else {
                                    Write-Warning 'KMS bearer token could not be resolved; stream download may fail with 401 on key retrieval'
                                }
                            }

                            $Arg = @( ('-o "{0}"' -f ($vidFullFile -replace '%', '%%')), $Endpoint)

                            # Construct Format for this specific video, language and audio languages available
                            $hasUserDefinedFormat = -not [string]::IsNullOrWhiteSpace([string]$Format)
                            if ( $hasUserDefinedFormat) {
                                $ThisFormat = ([string]$Format).Trim()
                            }
                            else {
                                $ThisFormat = Get-PreferredOnDemandYtDlpFormat -OnDemandPage $OnDemandPage -Endpoint $Endpoint -FallbackFormat 'worstvideo+bestaudio'
                                if ($ThisFormat -ne 'worstvideo+bestaudio') {
                                    Write-Verbose ('Detected adaptive format metadata on medius page; using preferred yt-dlp format {0}' -f $ThisFormat)
                                }
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

                            if ($kmsBearerToken) {
                                if ([string]::IsNullOrWhiteSpace($kmsBearerAuthorizationToken)) {
                                    $kmsBearerAuthorizationToken = $kmsBearerToken
                                }
                                $Arg += ('--add-headers "Authorization: Bearer {0}"' -f $kmsBearerAuthorizationToken)
                            }
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
                        if ( $downloadLink) {
                            $Endpoint = $downloadLink
                            $Arg = @( ('-o "{0}"' -f $vidFullFile), $downloadLink)
                        }
                        else {
                            Write-Host ('No video link for {0}' -f ($SessionToGet.Title))
                            $Endpoint = $null
                        }
                    }
                    if ( $Endpoint) {
                        # Direct, AMS or YT video found, attempt download but first define common parameters
                        $requiresMicrosoftCookieAuth = (($EventType -eq 'CUSTOM') -or $script:MSAContentAuthRequired -or -not [string]::IsNullOrWhiteSpace($kmsBearerToken))
                        $targetIsMicrosoftUrl = (Test-IsMicrosoftContentUrl -Url $Endpoint) -or (Test-IsMicrosoftContentUrl -Url $downloadLink)
                        $ytDlpCookieFile = $null

                        if ($requiresMicrosoftCookieAuth -and $targetIsMicrosoftUrl -and -not $CookiesFromBrowser) {
                            $ytDlpCookieFile = Resolve-YtDlpCookieFile -PreferredCookieFile $CookiesFile
                            if ($ytDlpCookieFile) {
                                Write-Verbose ('Using Netscape cookie file for authenticated Microsoft download: {0}' -f $ytDlpCookieFile)
                                $Arg += ('--cookies "{0}"' -f $ytDlpCookieFile)
                            }
                        }

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
                    if ([string]::IsNullOrWhiteSpace($CaptionExt)) {
                        $CaptionExt = 'vtt'
                        Write-Verbose 'Caption extension was not set; defaulting to vtt'
                    }
                    $captionExtFile = $vidFullFile -replace '.mp4', ('.{0}' -f $CaptionExt)

                    if ((Test-Path -LiteralPath $captionExtFile) -and -not $Overwrite) {
                        Write-Host ('Caption file exists {0}' -f $captionExtFile) -ForegroundColor Gray
                    }
                    else {
                        $captionInfoSourceUrl = $SessionToGet.onDemand
                        if ([string]::IsNullOrWhiteSpace($captionInfoSourceUrl) -and -not [string]::IsNullOrWhiteSpace($downloadLink)) {
                            $captionInfoSourceUrl = $downloadLink
                        }
                        if (($EventType -eq 'CUSTOM') -and -not [string]::IsNullOrWhiteSpace($captionInfoSourceUrl)) {
                            $captionInfoSourceUrl = Resolve-CustomSignedOnDemandUrl -Session $SessionToGet -OnDemandUrl $captionInfoSourceUrl -CatalogUrl $EventAPIUrl -Proxy $ProxyURL
                        }

                        # Caption file in AMS needs seperate download, fetch onDemand page if not already downloaded for video
                        if (! $OnDemandPage) {
                            if ( $captionInfoSourceUrl) {
                                try {
                                    Write-Host ('Fetching video page to retrieve transcript information from {0}' -f $captionInfoSourceUrl)
                                    $DownloadedPage = Invoke-WebRequestWithMSAAuthSupport -Method Get -Uri $captionInfoSourceUrl -Proxy $ProxyURL -DisableKeepAlive
                                    if ( $DownloadedPage) {
                                        $OnDemandPage = $DownloadedPage.Content
                                    }
                                }
                                catch {
                                    #Problem retrieving file, look for alternative options
                                }
                            }

                        }
                        # Check for vtt files before we check any direct caption file (likely docx now)
                        $captionFileLink = $Null
                        $captionLanguageSelected = $null
                        If( -not [string]::IsNullOrEmpty($OnDemandPage)) {
                            $CaptionConfig = Resolve-OnDemandCaptionsConfiguration -OnDemandPage $OnDemandPage
                        }
                        if ( $CaptionConfig) {
                            $preferredCaptionLanguages = @()
                            if ( $Subs) {
                                $preferredCaptionLanguages += $Subs
                            }
                            $preferredCaptionLanguages += 'en-us', 'en'

                            $captionSelection = Resolve-CaptionSourceByPreferredLanguage -LanguageList $CaptionConfig -PreferredLanguages $preferredCaptionLanguages
                            if ($captionSelection) {
                                $captionFileLink = $captionSelection.Src
                                $captionLanguageSelected = $captionSelection.Language
                            }
                        }
                        if ( ! $CaptionFileLink) {
                            $captionFileLink = $SessionToGet.captionFileLink
                        }
                        if ( ! $captionFileLink) {

                            if (! $OnDemandPage) {
                                # Try if there is caption file reference on page
                                try {
                                    $DownloadedPage = Invoke-WebRequestWithMSAAuthSupport -Method Get -Uri $captionInfoSourceUrl -Proxy $ProxyURL -DisableKeepAlive
                                    $OnDemandPage = $DownloadedPage.Content
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
                            if ($captionLanguageSelected) {
                                Write-Verbose ('Selected caption language {0} for session {1}' -f $captionLanguageSelected, $SessionToGet.sessioncode)
                            }
                            else {
                                Write-Verbose ('Selected caption language unknown for session {0} (caption source did not expose language metadata)' -f $SessionToGet.sessioncode)
                            }
                            Write-Verbose ('Retrieving caption file from URL {0}' -f $captionFileLink)

                            $captionFullFile = $captionExtFile
                            Write-Verbose ('Attempting download {0} to {1}' -f $captionFileLink, $captionFullFile)
                            $captionNeedsAuthDownload = (Test-IsMicrosoftContentUrl -Url $captionFileLink)
                            if ($captionNeedsAuthDownload) {
                                Write-Verbose ('Downloading caption file with authenticated request for session {0}' -f $SessionToGet.sessioncode)
                                $captionResponse = Invoke-WebRequestWithMSAAuthSupport -Method Get -Uri $captionFileLink -Proxy $ProxyURL -DisableKeepAlive -OutFile $captionFullFile
                                if (-not $captionResponse -and -not ((Test-Path -LiteralPath $captionFullFile) -and ((Get-ChildItem -LiteralPath $captionFullFile -ErrorAction SilentlyContinue).Length -gt 0))) {
                                    Write-Warning ('Problem downloading or missing captions of {0} {1}' -f $SessionToGet.scheduleCode, $SessionToGet.Title)
                                }
                                elseif ( $SessionTime -and $Timestamp) {
                                    Set-ItemProperty -LiteralPath $captionFullFile -Name LastWriteTime -Value ([System.DateTime]$SessionTime).ToUniversalTime()
                                }
                            }
                            else {
                                Add-BackgroundDownloadJob -Type 3 -FilePath $captionExtFile -DownloadUrl $captionFileLink -File $captionFullFile -Timestamp $SessionTime -scheduleCode ($SessionToGet.sessioncode) -Title ($SessionToGet.Title)
                            }

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
                        $ValidUrl = Invoke-WebRequestWithMSAAuthSupport -Uri $DownloadURL -Method Head -DisableKeepAlive -MaximumRedirection 10 -Proxy $ProxyURL
                    }
                    else {
                        $ValidUrl = Invoke-WebRequestWithMSAAuthSupport -Uri $DownloadURL -Method Get -DisableKeepAlive -MaximumRedirection 10 -Proxy $ProxyURL
                    }
                }
                catch {
                    $ValidUrl = $false
                }

                if ( $downloadLink -match 'confirmation\.aspx' -and $ValidURL.Headers.'Content-Type' -ilike 'text/html') {
                    # Extra parsing for MS downloads
                    if ( $ValidUrl.RawContent -match 'href="(?<Url>https:\/\/download\.microsoft\.com\/download[\/0-9\-]*\/.*(pdf|pptx))".*click here to download manually') {
                        $DownloadURL = [System.Web.HttpUtility]::UrlDecode( $Matches.Url)
                        $ValidUrl = Invoke-WebRequestWithMSAAuthSupport -Uri $DownloadURL -Method Head -DisableKeepAlive -MaximumRedirection 10 -Proxy $ProxyURL
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
                    $slidedeckFullFile = [System.IO.Path]::Combine( $ContentTargetFolder, $slidedeckFile)
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
                Write-Host ('No slidedeck link for {0}' -f ($SessionToGet.Title)) 
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

    # Clear empty per-session folder
    if ($UseSessionFolders -and -not (Get-ChildItem -Path $ContentTargetFolder)) {
        Remove-Item -Path $ContentTargetFolder -Force
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