#requires -version 5.1
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
    Version 4.55, November 24, 2026

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

    .PARAMETER ProgrammingLanguage
    Only retrieve sessions tagged with one or more of the specified programming languages.
    When multiple languages are specified, sessions matching any of the specified languages are included.

    .PARAMETER SessionLevel
    Only retrieve sessions at the specified level(s). Valid values are 100 (Foundational), 200 (Intermediate),
    300 (Advanced), and 400 (Expert). When not specified, sessions at all levels are included, as well as
    sessions without a level designation.

    .PARAMETER ScheduleCode
    Only retrieve sessions with this session code. You can use one or more codes.

    .PARAMETER NoVideos
    Switch to indicate you don't want to download videos.
    Published slidedecks can still be downloaded, including for sessions without on-demand video.

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
    - Custom                                       : Custom event endpoint supporting personal or Microsoft 365 work/school sign-in
    - Ignite                                       : Ignite contents (current)
    - Ignite2026                                   : Ignite contents from that year
    - Ignite2025                                   : Ignite contents from that year
    - Inspire                                      : Inspire contents (current)
    - Build                                        : Build contents (current)
    - Build2026                                    : Build contents from that year
    - Build2025                                    : Build contents from that year
    - MEC                                          : MEC contents (current)
    - MEC2022                                      : MEC contents from that year

    .PARAMETER OGVPicker
    Specify that you want to pick sessions to download using Out-GridView.

    .PARAMETER EventUrl
    HTTPS URL to use for Custom events, without a hostname restriction. The catalog is inspected
    anonymously first. Authentication is acquired only when required. GET catalogs use a page
    parameter; supported POST catalogs use JSON paging and browser request-context discovery when needed.
    Schedule-builder /api/schedule-builder/sessions endpoints can discover the /sessions/search POST endpoint.
    Selected session details can publish multiple slidedecks. The first deck uses the usual session-based
    filename; subsequent decks receive -1, -2, etc. before their actual file extension.

    .PARAMETER CookiesFile
    Netscape-format cookie export for yt-dlp downloads. For Custom events, also supplies cookies
    when catalog or content authentication is needed, instead of cached authentication or sign-in.
    Required portal request context is discovered through authenticated HTML requests only.
    Invalid or rejected cookies and unavailable context cause an error without interactive fallback.
    The export is not modified and its credentials/context are not saved to authentication caches.
    For Custom events, this parameter takes precedence over CookiesFromBrowser.

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

    .PARAMETER UseSessionFolders
    Stores the content of each session in its own subfolder of the download folder, instead of
    saving everything directly in the download folder. The subfolder is named after the session,
    matching the name used for the video file, and receives all artifacts of that session such
    as the video, slidedeck and caption files. Session folders that end up empty are removed.

    .PARAMETER NormalizeFileNames
    Makes generated file and folder names safe for OneDrive/SharePoint sync. Applies Unicode
    compatibility normalization (NFKC), converts special spaces to regular spaces, removes control
    and invisible characters, replaces characters OneDrive rejects (<>:"/\|?*) with an underscore,
    trims leading/trailing spaces and periods, and avoids reserved Windows device names.

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
    [ValidateSet('.NET', 'ASP.NET Core', 'C', 'C#', 'C++', 'CUDA', 'Java', 'JavaScript', 'Node.js', 'Python', 'Spark', 'TypeScript', 'sgs')]
    [string[]]$ProgrammingLanguage,

    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [parameter( Mandatory = $false, ParameterSetName = 'DownloadDirect')]
    [ValidateSet(100, 200, 300, 400)]
    [int[]]$SessionLevel,

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
    [ValidateSet('MEC', 'MEC2022', 'Ignite', 'Ignite2025', 'Ignite2026', 'Inspire', 'Build', 'Build2026', 'Build2025', 'Custom', 'Ignite2024', 'Ignite2023', 'Ignite2022', 'Ignite2021H1', 'Ignite2021H2', 'Ignite2020')]
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
    [switch]$NormalizeFileNames,

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
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [string]$TempPath,

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'DownloadDirect')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$CookiesFile,

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [ValidateSet( 'brave', 'chrome', 'chromium', 'edge', 'firefox', 'opera', 'safari', 'vivaldi', 'whale' )]
    [string]$CookiesFromBrowser
)

# Max age for cache, older than this # hours will force info refresh
$script:MaxCacheAge = 24
# Max age for MSA auth token, expiring in less than these minutes will force re-authentication
$script:MinTokenValidityMinutes = 15

$script:YouTubeEXE = 'yt-dlp.exe'
$script:YouTubeDL = Join-Path $PSScriptRoot $YouTubeEXE
$script:FFMPEG = Join-Path $PSScriptRoot 'ffmpeg.exe'

$script:YTlink = 'https://github.com/yt-dlp/yt-dlp/releases/download/2023.07.06/yt-dlp.exe'
$script:FFMPEGlink = 'https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-essentials.zip'

# Fix 'Could not create SSL/TLS secure channel' issues with Invoke-WebRequest
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$script:BackgroundDownloadJobs = @()
$script:JobProgressCache = @{}
$script:JobFileLastWrite = @{}
$script:MSAAuthWebSession = $null
$script:MSAAuthOrigin = $null
$script:MSABearerToken = $null
$script:MSAContentAuthRequired = $false
$script:CustomSignedOnDemandUrlCache = @{}
$script:CustomSignedOnDemandAuthRequiredBySession = @{}
$script:MSAAuthInteractiveAttemptedByHost = @{}
$script:SuppressWebRequestDebugDetails = $true
$script:CustomCookieFile = $null
$script:CustomCookieWebSession = $null

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
    $cleaned = (((((((($title -replace '\]', ')') -replace '\[', '(') -replace [char]0x202f, ' ') -replace '["\\/\?\*]', ' ') -replace ':', '-') -replace '  ', ' ') -replace '\?\?\?', '') -replace '\<|\>|:|"|/|\\|\||\?|\*', '').Trim()
    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars() | ForEach-Object { [Regex]::Escape($_) }
    $cleaned = $cleaned -replace ($invalidChars -join '|'), ''
    if ($NormalizeFileNames) {
        $cleaned = ConvertTo-SyncSafeFileName -Name $cleaned
    }
    return $cleaned
}

function ConvertTo-SyncSafeFileName {
    # Makes a file name safe for OneDrive/SharePoint sync (and Windows in general)
    param([string]$Name)

    # Trademark/copyright signs would otherwise be expanded by NFKC (e.g. ™ -> TM)
    $Name = $Name -replace '[\u2122\u00AE\u00A9]', ''
    # Compatibility normalization folds ligatures, full-width forms and typographic variants
    $Name = $Name.Normalize([Text.NormalizationForm]::FormKC)
    # En/em dashes are common in titles; use a plain hyphen
    $Name = $Name -replace '[\u2013\u2014]', '-'
    # Unicode space separators (NBSP, thin space, ...) become a regular space
    $Name = $Name -replace '\p{Zs}', ' '
    # Drop control and invisible format characters (zero-width space, BOM, soft hyphen, ...)
    $Name = $Name -replace '[\p{Cc}\p{Cf}]', ''
    # Characters rejected by OneDrive/Windows
    $Name = $Name -replace '[<>:"/\\|?*]', '_'
    $Name = $Name -replace '\s{2,}', ' '
    # No leading/trailing spaces or periods, no leading double periods
    $Name = $Name.Trim().Trim('.').Trim()
    if ([string]::IsNullOrWhiteSpace($Name)) { $Name = '_' }
    # Windows reserved device names
    if ($Name -match '^(?i:CON|PRN|AUX|NUL|COM[0-9]|LPT[0-9])$') { $Name = '_{0}' -f $Name }
    return $Name
}

function ConvertTo-ArgumentString {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) {
        return '""'
    }

    $stringValue = [string]$Value
    $escapedValue = $stringValue -replace '(\\*)"', '$1$1\"'
    $escapedValue = $escapedValue -replace '(\\+)$', '$1$1'
    return '"{0}"' -f $escapedValue
}

function Get-ObjectPropertyValue {
    param(
        [parameter(Mandatory = $true)][object]$Object,
        [parameter(Mandatory = $true)][string[]]$Name
    )

    foreach ($propertyName in $Name) {
        if ($Object.PSObject.Properties.Match($propertyName).Count -gt 0) {
            $value = $Object.$propertyName
            if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) {
                return $value
            }
        }
    }

    return $null
}

function ConvertTo-MetadataValueList {
    param([AllowNull()][object]$Value)

    $items = [System.Collections.ArrayList]@()
    if ($null -eq $Value) {
        return @()
    }

    foreach ($entry in @($Value)) {
        if ($null -eq $entry) {
            continue
        }

        if ($entry -is [string]) {
            foreach ($part in ($entry -split [char]9)) {
                $candidate = $part.Trim()
                if (-not [string]::IsNullOrWhiteSpace($candidate)) {
                    $items.Add($candidate) | Out-Null
                }
            }
            continue
        }

        $displayValue = Get-ObjectPropertyValue -Object $entry -Name @('displayValue', 'name', 'title', 'logicalValue')
        if ($displayValue) {
            $items.Add([string]$displayValue) | Out-Null
        }
        else {
            $candidate = ([string]$entry).Trim()
            if (-not [string]::IsNullOrWhiteSpace($candidate)) {
                $items.Add($candidate) | Out-Null
            }
        }
    }

    return @($items | Select-Object -Unique)
}

function Format-MetadataValueList {
    param([AllowNull()][object]$Value)

    return (ConvertTo-MetadataValueList -Value $Value) -join '; '
}

function ConvertTo-SessionDateTime {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [datetime]) {
        return $Value
    }

    $stringValue = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($stringValue)) {
        return $null
    }

    # Catalogs publish US-style timestamps, which fail to parse under other cultures
    $parsed = [datetime]::MinValue
    $formats = [string[]]@('MM/dd/yyyy HH:mm:ss', 'M/d/yyyy H:mm:ss', 'MM/dd/yyyy', 'yyyy-MM-ddTHH:mm:ss', 'yyyy-MM-dd HH:mm:ss', 'yyyy-MM-dd', 'yyyyMMdd')
    if ([datetime]::TryParseExact($stringValue, $formats, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
        return $parsed
    }

    if ([datetime]::TryParse($stringValue, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
        return $parsed
    }

    if ([datetime]::TryParse($stringValue, [System.Globalization.CultureInfo]::CurrentCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
        return $parsed
    }

    return $null
}

function Get-SessionPresentationUrl {
    param(
        [parameter(Mandatory = $true)][object]$Session,
        [AllowNull()][string]$FallbackSlidedeckUrl
    )

    $presentationUrl = Get-ObjectPropertyValue -Object $Session -Name @('slideDeck', 'slidedeck', 'presentationUrl')
    if ($presentationUrl) {
        return [string]$presentationUrl
    }

    $sessionCode = Get-ObjectPropertyValue -Object $Session -Name @('sessionCode', 'scheduleCode', 'code')
    if ($FallbackSlidedeckUrl -and $sessionCode) {
        # The fallback is a template for most events but is a scraped URL for custom events, so
        # substitute the placeholder literally: -f would throw on a URL containing other braces.
        return $FallbackSlidedeckUrl.Replace('{0}', [string]$sessionCode)
    }

    return $null
}

function Get-SessionOriginalVideoUrl {
    param(
        [parameter(Mandatory = $true)][object]$Session,
        [AllowNull()][string]$EventName,
        [AllowNull()][string]$FallbackVideoUrl
    )

    if ($EventName -like 'Build*') {
        $localizedId = Get-ObjectPropertyValue -Object $Session -Name @('localizedId')
        if ($localizedId) {
            $locale = Get-ObjectPropertyValue -Object $Session -Name @('langLocale')
            if (-not $locale) {
                $locale = 'en-US'
            }
            return 'https://build.microsoft.com/{0}/sessions/{1}' -f $locale, $localizedId
        }
    }

    $sessionUrl = Get-ObjectPropertyValue -Object $Session -Name @('onDemand', 'webpage_url', 'registrationLink')
    if ($sessionUrl) {
        return [string]$sessionUrl
    }

    return $FallbackVideoUrl
}

function Get-SessionRepositoryUrl {
    param(
        [parameter(Mandatory = $true)][object]$Session
    )

    # GitHub paths that look like owner/repo but are product or marketing pages
    $nonRepositoryOwners = @('about', 'apps', 'collections', 'customer-stories', 'enterprise', 'events', 'explore', 'features', 'login', 'marketplace', 'organizations', 'orgs', 'pricing', 'readme', 'resources', 'security', 'settings', 'solutions', 'sponsors', 'team', 'topics', 'trending')

    $candidates = [System.Collections.ArrayList]@()
    foreach ($propertyName in @('relatedResources', 'nextStep', 'sessionLinks')) {
        if ($Session.PSObject.Properties.Match($propertyName).Count -eq 0) {
            continue
        }

        foreach ($resource in @($Session.$propertyName)) {
            if ($null -eq $resource) {
                continue
            }

            if ($resource -is [string]) {
                $link = $resource
                $category = ''
            }
            else {
                $link = Get-ObjectPropertyValue -Object $resource -Name @('link', 'url', 'linkUrl')
                $category = [string](Get-ObjectPropertyValue -Object $resource -Name @('category'))
            }

            if ($link) {
                $candidates.Add([PSCustomObject]@{ Url = [string]$link; IsCodeSample = ($category -match '(?i)code') }) | Out-Null
            }
        }
    }

    $description = Get-ObjectPropertyValue -Object $Session -Name @('description', 'aiDescription')
    if ($description) {
        foreach ($match in [regex]::Matches([string]$description, '(?i)https://github\.com/[^\s<>"'')\]]+')) {
            $candidates.Add([PSCustomObject]@{ Url = $match.Value; IsCodeSample = $false }) | Out-Null
        }
    }

    # Prefer links explicitly published as code samples, keeping catalog order otherwise
    $orderedCandidates = @($candidates | Where-Object { $_.IsCodeSample }) + @($candidates | Where-Object { -not $_.IsCodeSample })

    foreach ($candidate in $orderedCandidates) {
        $url = (($candidate.Url -split '[?#]')[0]).TrimEnd('.', ',', ')', '/')
        if ($url -notmatch '(?i)^https://github\.com/(?<owner>[A-Za-z0-9][A-Za-z0-9-]*)/(?<repo>[A-Za-z0-9_.-]+)$') {
            continue
        }

        if ($nonRepositoryOwners -contains $Matches.owner.ToLowerInvariant()) {
            continue
        }

        return 'https://github.com/{0}/{1}' -f $Matches.owner, $Matches.repo
    }

    return $null
}

function New-EventSessionVideoMetadata {
    param(
        [parameter(Mandatory = $true)][object]$Session,
        [AllowNull()][string]$EventName,
        [AllowNull()][string]$OriginalVideoUrl,
        [AllowNull()][string]$PresentationUrl,
        [AllowNull()][string]$RepositoryUrl
    )

    $description = Get-ObjectPropertyValue -Object $Session -Name @('description', 'aiDescription')
    $sessionCode = Get-ObjectPropertyValue -Object $Session -Name @('sessionCode', 'scheduleCode', 'code')
    $speakerNames = Format-MetadataValueList -Value (Get-ObjectPropertyValue -Object $Session -Name @('speakerNames', 'speakers'))
    $tags = @()
    foreach ($tagProperty in @('tags', 'products', 'contentCategory', 'solutionArea', 'topic', 'programmingLanguages', 'sessionType')) {
        if ($Session.PSObject.Properties.Match($tagProperty).Count -gt 0) {
            $tags += ConvertTo-MetadataValueList -Value $Session.$tagProperty
        }
    }

    $year = $null
    $startDateTime = Get-ObjectPropertyValue -Object $Session -Name @('startDateTime')
    if ($startDateTime) {
        $sessionDate = ConvertTo-SessionDateTime -Value $startDateTime
        if ($sessionDate) {
            $year = $sessionDate.Year
        }
        else {
            Write-Verbose ('Unable to determine metadata year for {0} from value {1}' -f $sessionCode, $startDateTime)
        }
    }

    $commentLines = [System.Collections.ArrayList]@()
    if ($description) {
        $commentLines.Add(([string]$description).Trim()) | Out-Null
        $commentLines.Add('') | Out-Null
    }
    if ($PresentationUrl) {
        $commentLines.Add(('Presentation: {0}' -f $PresentationUrl)) | Out-Null
    }
    if ($RepositoryUrl) {
        $commentLines.Add(('GitHub repository: {0}' -f $RepositoryUrl)) | Out-Null
    }
    if ($OriginalVideoUrl) {
        $commentLines.Add(('Original video: {0}' -f $OriginalVideoUrl)) | Out-Null
    }

    return [PSCustomObject]@{
        Title        = [string](Get-ObjectPropertyValue -Object $Session -Name @('title'))
        Subtitle     = [string]$sessionCode
        Tags         = (($tags | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Unique) -join '; ')
        Comments     = ($commentLines -join [System.Environment]::NewLine).Trim()
        Artists      = $speakerNames
        Year         = $year
        Genre        = 'Documentary; Lecture'
        Producer     = 'Microsoft'
        PromotionUrl = $OriginalVideoUrl
    }
}

function Set-VideoMetadata {
    param(
        [parameter(Mandatory = $true)][string]$File,
        [parameter(Mandatory = $true)][object]$Metadata
    )

    if (-not (Test-Path -LiteralPath $script:FFMPEG)) {
        Write-Warning ('Unable to write video metadata for {0}: ffmpeg.exe was not found at {1}' -f $File, $script:FFMPEG)
        return
    }

    if (-not (Test-Path -LiteralPath $File)) {
        Write-Warning ('Unable to write video metadata because file was not found: {0}' -f $File)
        return
    }

    if (-not $script:MetadataTimeoutSeconds) {
        $script:MetadataTimeoutSeconds = 900
    }

    $originalItem = Get-Item -LiteralPath $File
    $creationTime = $originalItem.CreationTime
    $lastWriteTime = $originalItem.LastWriteTime
    $outputFile = '{0}.{1}.metadata{2}' -f [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($File), [System.IO.Path]::GetFileNameWithoutExtension($File)), (New-Guid).Guid, [System.IO.Path]::GetExtension($File)
    # -nostdin stops ffmpeg from consuming the console input handle, and the quiet logging
    # options keep the redirected stderr pipe from filling up while remuxing large files.
    $arguments = @('-nostdin', '-hide_banner', '-loglevel', 'error', '-nostats', '-y', '-i', $File, '-map', '0', '-c', 'copy')

    # Only tags the MP4 muxer maps to standard atoms are stored; custom keys are silently
    # dropped by ffmpeg, and 'use_metadata_tags' would make every tag unreadable to Windows.
    $metadataMap = [ordered]@{
        title        = $Metadata.Title
        show         = $Metadata.Subtitle
        episode_id   = $Metadata.Subtitle
        artist       = $Metadata.Artists
        album_artist = $Metadata.Artists
        date         = $Metadata.Year
        genre        = $Metadata.Genre
        comment      = $Metadata.Comments
        description  = $Metadata.Comments
        synopsis     = $Metadata.Comments
        keywords     = $Metadata.Tags
        copyright    = $Metadata.Producer
        network      = $Metadata.Producer
    }

    foreach ($metadataItem in $metadataMap.GetEnumerator()) {
        if ($null -ne $metadataItem.Value -and -not [string]::IsNullOrWhiteSpace([string]$metadataItem.Value)) {
            $arguments += '-metadata'
            $arguments += ('{0}={1}' -f $metadataItem.Key, $metadataItem.Value)
        }
    }

    $arguments += $outputFile

    $pinfo = New-Object System.Diagnostics.ProcessStartInfo
    $pinfo.FileName = $script:FFMPEG
    $pinfo.RedirectStandardError = $true
    $pinfo.RedirectStandardOutput = $true
    $pinfo.UseShellExecute = $false
    $pinfo.CreateNoWindow = $true
    $pinfo.Arguments = ($arguments | ForEach-Object { ConvertTo-ArgumentString $_ }) -join ' '

    Write-Verbose ('Writing video metadata to {0}' -f $File)
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $pinfo
    $process.Start() | Out-Null

    # Drain both pipes concurrently; reading one to the end before the other deadlocks as soon
    # as ffmpeg fills the buffer of the pipe that is not being read.
    $stdoutTask = $process.StandardOutput.ReadToEndAsync()
    $stderrTask = $process.StandardError.ReadToEndAsync()

    if (-not $process.WaitForExit( $script:MetadataTimeoutSeconds * 1000)) {
        Write-Warning ('Timed out writing video metadata for {0} after {1} seconds' -f $File, $script:MetadataTimeoutSeconds)
        try { $process.Kill() } catch {}
        try { $process.WaitForExit( 5000) | Out-Null } catch {}
        Remove-Item -LiteralPath $outputFile -Force -ErrorAction SilentlyContinue
        return
    }

    $stdout = $stdoutTask.GetAwaiter().GetResult()
    $stderr = $stderrTask.GetAwaiter().GetResult()

    if ($process.ExitCode -eq 0 -and (Test-Path -LiteralPath $outputFile) -and (Get-Item -LiteralPath $outputFile).Length -gt 0) {
        Move-Item -LiteralPath $outputFile -Destination $File -Force
        [System.IO.File]::SetCreationTime($File, $creationTime)
        [System.IO.File]::SetLastWriteTime($File, $lastWriteTime)
        Write-Verbose ('Video metadata written to {0}' -f $File)
    }
    else {
        Remove-Item -LiteralPath $outputFile -Force -ErrorAction SilentlyContinue
        Write-Warning ('Unable to write video metadata for {0}: {1}' -f $File, (($stderr, $stdout | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join [System.Environment]::NewLine))
    }
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

function ConvertFrom-EventJson {
    param(
        [parameter(Mandatory = $true, ValueFromPipeline = $true)][string]$InputObject,
        [int]$Depth = 1024
    )
    process {
        $parameters = @{ InputObject = $InputObject; ErrorAction = 'Stop' }
        if ($PSVersionTable.PSVersion.Major -ge 6) { $parameters.Depth = $Depth }
        ConvertFrom-Json @parameters
    }
}

function Invoke-WebRequestNoRedirect {
    param(
        [parameter(Mandatory = $true)][hashtable]$Parameters,
        [switch]$Rest
    )

    if ($Parameters.MaximumRedirection -ne 0) { throw 'This request helper requires redirects to be disabled.' }
    if ($PSVersionTable.PSVersion.Major -ge 7) {
        if ($Rest) { return Invoke-RestMethod @Parameters }
        return Invoke-WebRequest @Parameters
    }

    # Windows PowerShell can lose the response to a redirect with MaximumRedirection = 0.
    $request = [System.Net.HttpWebRequest]::Create([uri]$Parameters.Uri)
    $request.AllowAutoRedirect = $false
    $request.Method = if ($Parameters.Method) { [string]$Parameters.Method } else { 'GET' }
    $timeout = if ($Parameters.TimeoutSec) { [int]$Parameters.TimeoutSec * 1000 } else { 30000 }
    $request.Timeout = $timeout
    $request.ReadWriteTimeout = $timeout
    $request.AutomaticDecompression = [System.Net.DecompressionMethods]'GZip, Deflate'
    if ($Parameters.Proxy) { $request.Proxy = [System.Net.WebProxy]::new([uri]$Parameters.Proxy) }
    if ($Parameters.WebSession) {
        $request.CookieContainer = $Parameters.WebSession.Cookies
        $request.UserAgent = $Parameters.WebSession.UserAgent
        if ($Parameters.WebSession.Credentials) { $request.Credentials = $Parameters.WebSession.Credentials }
        if ($Parameters.WebSession.UseDefaultCredentials) { $request.UseDefaultCredentials = $true }
        if ($Parameters.WebSession.Proxy -and -not $Parameters.Proxy) { $request.Proxy = $Parameters.WebSession.Proxy }
        foreach ($certificate in $Parameters.WebSession.Certificates) { $null = $request.ClientCertificates.Add($certificate) }
    }
    $headers = @{}
    if ($Parameters.WebSession) {
        foreach ($key in $Parameters.WebSession.Headers.Keys) { $headers[$key] = $Parameters.WebSession.Headers[$key] }
    }
    if ($Parameters.Headers) {
        foreach ($key in $Parameters.Headers.Keys) { $headers[$key] = $Parameters.Headers[$key] }
    }
    foreach ($key in $headers.Keys) {
        switch ($key) {
            'Accept' { $request.Accept = $headers[$key] }
            'Content-Type' { $request.ContentType = $headers[$key] }
            'User-Agent' { $request.UserAgent = $headers[$key] }
            'Referer' { $request.Referer = $headers[$key] }
            default { $request.Headers[$key] = $headers[$key] }
        }
    }
    if ($Parameters.UserAgent) { $request.UserAgent = $Parameters.UserAgent }
    if ($Parameters.ContentType) { $request.ContentType = $Parameters.ContentType }
    if ($Parameters.ContainsKey('Body')) {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes([string]$Parameters.Body)
        $request.ContentLength = $bytes.Length
        $stream = $request.GetRequestStream()
        try { $stream.Write($bytes, 0, $bytes.Length) }
        finally { $stream.Dispose() }
    }

    $response = $null
    try {
        try { $response = [System.Net.HttpWebResponse]$request.GetResponse() }
        catch [System.Net.WebException] {
            $failure = $_.Exception
            while ($failure.InnerException -and $failure -isnot [System.Net.WebException]) { $failure = $failure.InnerException }
            if (-not $failure.Response) { throw $failure }
            $response = $failure.Response
        }
        $responseHeaders = @{}
        foreach ($key in $response.Headers.AllKeys) { $responseHeaders[$key] = $response.Headers[$key] }
        if ([int]$response.StatusCode -ge 300) {
            # Snapshot metadata before closing the response; callers inspect it after catching.
            $exception = [System.Exception]::new(('HTTP request returned status {0}.' -f [int]$response.StatusCode))
            $exception | Add-Member -NotePropertyName Response -NotePropertyValue ([pscustomobject]@{
                StatusCode = [int]$response.StatusCode; Headers = $responseHeaders; ResponseUri = $response.ResponseUri
            })
            throw $exception
        }
        $encoding = if ($response.ContentType -match '(?i);\s*charset=') { [System.Text.Encoding]::GetEncoding($response.CharacterSet) } else { [System.Text.Encoding]::UTF8 }
        $reader = [System.IO.StreamReader]::new($response.GetResponseStream(), $encoding, $true)
        try { $content = $reader.ReadToEnd() }
        finally { $reader.Dispose() }
        if ($Rest) {
            if ($response.ContentType -match '(?i)(?:application|text)/(?:[\w.+-]+\+)?json\b') {
                return ConvertFrom-Json -InputObject $content -ErrorAction Stop
            }
            return $content
        }
        return [pscustomobject]@{
            StatusCode = [int]$response.StatusCode
            Headers = $responseHeaders
            Content = $content
            BaseResponse = $response
        }
    }
    finally {
        if ($response) { $response.Close() }
    }
}

function Invoke-WebWithRetry {
    param(
        [scriptblock]$ScriptBlock,
        [hashtable]$Variables = @{},
        [int]$MaxAttempts = 3,
        [int]$InitialDelaySeconds = 2,
        [switch]$SuppressErrorDetails
    )
    $attempt = 0
    $delay = $InitialDelaySeconds
    do {
        $attempt++
        try {
            foreach ($kv in $Variables.GetEnumerator()) { Set-Variable -Name $kv.Key -Value $kv.Value -Scope 0 }
            return (& $ScriptBlock)
        }
        catch {
            if ($_.Exception -is [System.Management.Automation.PipelineStoppedException]) { throw }
            if ($attempt -ge $MaxAttempts) { throw }
            $statusCode = $null
            if ($_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }
            # Redirects and client errors need caller handling, not a transient-error retry.
            if ($statusCode -and $statusCode -ge 300 -and $statusCode -lt 500 -and $statusCode -ne 429) { throw }
            $errorMessage = if ($SuppressErrorDetails) { 'HTTP request failed; credential-bearing details withheld' } else { $_.Exception.Message }
            Write-Warning ('Request failed (attempt {0}/{1}): {2}. Retrying in {3}s...' -f $attempt, $MaxAttempts, $errorMessage, $delay)
            Start-Sleep -Seconds $delay
            $delay *= 2
        }
    } while ($attempt -lt $MaxAttempts)
}

function Invoke-JsonRestMethod {
    # Windows PowerShell 5.1 decodes JSON responses without a charset as ISO-8859-1, which corrupts
    # non-ASCII titles (e.g. ’ becomes â + control characters). Decode explicitly as UTF-8 there.
    param(
        [parameter(Mandatory = $true)][uri]$Uri,
        [string]$Method = 'Get',
        [hashtable]$Headers,
        [string]$UserAgent,
        $WebSession,
        [string]$Proxy,
        $Body,
        [int]$TimeoutSec
    )

    $parameters = @{ Uri = $Uri; Method = $Method }
    if ($Headers) { $parameters.Headers = $Headers }
    if ($UserAgent) { $parameters.UserAgent = $UserAgent }
    if ($WebSession) { $parameters.WebSession = $WebSession }
    if ($Proxy) { $parameters.Proxy = $Proxy }
    if ($null -ne $Body) { $parameters.Body = $Body }
    if ($TimeoutSec) { $parameters.TimeoutSec = $TimeoutSec }

    if ($PSVersionTable.PSEdition -eq 'Core') {
        return (Invoke-RestMethod @parameters)
    }
    $response = Invoke-WebRequest @parameters -UseBasicParsing
    $text = [System.Text.Encoding]::UTF8.GetString($response.RawContentStream.ToArray())
    return ($text | ConvertFrom-Json)
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

function New-CustomCookieFileException {
    param([string]$Message)

    $exception = [System.InvalidOperationException]::new($Message)
    $exception.Data['CookieFileFailure'] = $true
    return $exception
}

function ConvertFrom-NetscapeCookieContent {
    param([parameter(Mandatory = $true)][AllowEmptyString()][string]$Content)

    $session = [Microsoft.PowerShell.Commands.WebRequestSession]::new()
    $lines = $Content -split "`r?`n"
    $session.Cookies.Capacity = [Math]::Max(300, $lines.Count)
    $session.Cookies.PerDomainCapacity = $session.Cookies.Capacity
    $lineNumber = 0
    foreach ($line in $lines) {
        $lineNumber++
        if ([string]::IsNullOrWhiteSpace($line) -or ($line.StartsWith('#') -and -not $line.StartsWith('#HttpOnly_'))) { continue }
        $parts = $line -split "`t"
        $httpOnly = $parts[0].StartsWith('#HttpOnly_')
        $domain = if ($httpOnly) { $parts[0].Substring(10) } else { $parts[0] }
        $expires = 0L
        if ($parts.Count -ne 7 -or $parts[1] -cnotin @('TRUE', 'FALSE') -or $parts[3] -cnotin @('TRUE', 'FALSE') -or
            $domain -notmatch '^\.?[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*$' -or
            -not $parts[2].StartsWith('/') -or $parts[2] -match '[\x00-\x20;]' -or
            $parts[5] -notmatch '^[^\x00-\x20()<>@,;:\\"/\[\]?={}]+$' -or
            $parts[6] -match '[\x00-\x1f\x7f;]' -or
            -not [long]::TryParse($parts[4], [ref]$expires) -or $expires -lt 0) {
            throw [System.FormatException]::new(('Invalid Netscape cookie record at line {0}; cookie values withheld.' -f $lineNumber))
        }
        try {
            $expiry = if ($expires -gt 0) { [DateTimeOffset]::FromUnixTimeSeconds($expires).UtcDateTime } else { [datetime]::MinValue }
            if ($expires -gt 0 -and $expiry -le [datetime]::UtcNow) { continue }
            $cookie = [System.Net.Cookie]::new($parts[5], $parts[6], $parts[2])
            $cookie.Secure = $parts[3] -ceq 'TRUE'
            $cookie.HttpOnly = $httpOnly
            if ($expires -gt 0) { $cookie.Expires = $expiry }
            $hostName = $domain.TrimStart('.')
            if ($parts[1] -ceq 'TRUE') {
                $cookie.Domain = '.' + $hostName
                $session.Cookies.Add($cookie)
            }
            else {
                $session.Cookies.Add([uri]('https://{0}{1}' -f $hostName, $parts[2]), $cookie)
            }
        }
        catch {
            throw [System.FormatException]::new(('Invalid Netscape cookie record at line {0}; cookie values withheld.' -f $lineNumber))
        }
    }
    return $session
}

function Get-WebSessionCookies {
    param([parameter(Mandatory = $true)][Microsoft.PowerShell.Commands.WebRequestSession]$WebSession)

    $container = $WebSession.Cookies
    if ($container.PSObject.Methods['GetAllCookies']) { return $container.GetAllCookies() }
    # .NET Framework has no public all-domain enumerator. Keep the original Cookie objects.
    $flags = [System.Reflection.BindingFlags]'NonPublic,Instance,GetField'
    $domains = $container.GetType().InvokeMember('m_domainTable', $flags, $null, $container, @())
    foreach ($domain in $domains.Values) {
        $paths = $domain.GetType().InvokeMember('m_list', $flags, $null, $domain, @())
        foreach ($collection in $paths.Values) {
            foreach ($cookie in $collection) { $cookie }
        }
    }
}

function Get-CustomCookieWebSession {
    param([parameter(Mandatory = $true)][uri]$TargetUri)

    if (-not $script:CustomCookieWebSession) {
        try {
            $content = Get-Content -LiteralPath $script:CustomCookieFile -Raw -ErrorAction Stop
            $script:CustomCookieWebSession = ConvertFrom-NetscapeCookieContent -Content ([string]$content)
        }
        catch {
            if ($_.Exception -is [System.Management.Automation.PipelineStoppedException]) { throw }
            $reason = if ($_.Exception -is [System.FormatException]) { $_.Exception.Message } else { 'Unable to read the cookie export; check that it is an accessible file.' }
            throw (New-CustomCookieFileException -Message $reason)
        }
        $portalUri = [uri]$EventAPIUrl
        $hasPortalCookie = @(Get-WebSessionCookies -WebSession $script:CustomCookieWebSession | Where-Object {
            $probeUri = [uri]::new($portalUri, $_.Path)
            $script:CustomCookieWebSession.Cookies.GetCookies($probeUri).Count -gt 0
        }).Count -gt 0
        if (-not $hasPortalCookie) {
            $script:CustomCookieWebSession = $null
            throw (New-CustomCookieFileException -Message 'The cookie export contains no unexpired cookies for the Custom portal. Sign in and export its cookies again.')
        }
    }
    $script:MSABearerToken = $null
    $script:MSAAuthOrigin = $TargetUri.GetLeftPart([System.UriPartial]::Authority)
    $script:MSAAuthWebSession = $script:CustomCookieWebSession
    $script:MSAContentAuthRequired = $true
    return $script:CustomCookieWebSession
}

function ConvertTo-NetscapeCookieContent {
    param([parameter(Mandatory = $true)][Microsoft.PowerShell.Commands.WebRequestSession]$WebSession)

    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add('# Netscape HTTP Cookie File')
    $domainImplicit = [System.Net.Cookie].GetProperty('DomainImplicit', [System.Reflection.BindingFlags]'Instance,NonPublic')
    foreach ($cookie in (Get-WebSessionCookies -WebSession $WebSession)) {
        if ($cookie.Expired) { continue }
        $includeSubdomains = -not $domainImplicit.GetValue($cookie)
        $domain = if ($includeSubdomains) { '.' + $cookie.Domain.TrimStart('.') } else { $cookie.Domain.TrimStart('.') }
        if ($cookie.HttpOnly) { $domain = '#HttpOnly_' + $domain }
        $expiry = if ($cookie.Expires -eq [datetime]::MinValue) { 0 } else { ([DateTimeOffset]$cookie.Expires.ToUniversalTime()).ToUnixTimeSeconds() }
        $lines.Add(("{0}`t{1}`t{2}`t{3}`t{4}`t{5}`t{6}" -f $domain, $includeSubdomains.ToString().ToUpperInvariant(),
            $cookie.Path, $cookie.Secure.ToString().ToUpperInvariant(), $expiry, $cookie.Name, $cookie.Value))
    }
    return $lines -join [Environment]::NewLine
}

function New-CustomYtDlpCookieFile {
    param([parameter(Mandatory = $true)][Microsoft.PowerShell.Commands.WebRequestSession]$WebSession)

    $path = $null
    try {
        $path = [System.IO.Path]::GetTempFileName()
        [System.IO.File]::WriteAllText($path, (ConvertTo-NetscapeCookieContent -WebSession $WebSession), [System.Text.UTF8Encoding]::new($false))
    }
    catch {
        if ($path) { Remove-Item -LiteralPath $path -Force -ErrorAction Stop }
        if ($_.Exception -is [System.Management.Automation.PipelineStoppedException]) { throw }
        throw (New-CustomCookieFileException -Message 'Unable to prepare the temporary download cookie file; credentials withheld.')
    }
    return $path
}

function Start-CustomCookieDownloadJob {
    param(
        [parameter(Mandatory = $true)][string]$Url,
        [parameter(Mandatory = $true)][string]$File,
        [uri]$Proxy
    )

    $session = Get-CustomCookieWebSession -TargetUri ([uri]$Url)
    $content = ConvertTo-NetscapeCookieContent -WebSession $session
    $definition = ${function:ConvertFrom-NetscapeCookieContent}.ToString()
    Start-Job -ScriptBlock {
        param($url, $file, $proxy, $content, $definition)
        $ProgressPreference = 'SilentlyContinue'
        . ([scriptblock]::Create("function ConvertFrom-NetscapeCookieContent { $definition }"))
        $session = ConvertFrom-NetscapeCookieContent -Content $content
        $temporaryFile = $file + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
        try {
            $parameters = @{
                Uri = $url; Method = 'Get'; OutFile = $temporaryFile; PassThru = $true
                WebSession = $session; ErrorAction = 'Stop'; Verbose = $false; Debug = $false
            }
            if ($proxy) { $parameters.Proxy = $proxy }
            try { $response = Invoke-WebRequest @parameters }
            catch {
                if ($_.Exception -is [System.Management.Automation.PipelineStoppedException]) { throw }
                throw 'Cookie-backed download failed; credentials and response details withheld.'
            }
            if ([string]$response.Headers['Content-Type'] -match '(?i)^text/html\b') {
                throw 'Cookie-backed download returned HTML instead of the requested file. Sign in and export cookies again.'
            }
            Move-Item -LiteralPath $temporaryFile -Destination $file -Force -ErrorAction Stop
        }
        finally {
            if (Test-Path -LiteralPath $temporaryFile) { Remove-Item -LiteralPath $temporaryFile -Force -ErrorAction Stop }
        }
    } -ArgumentList $Url, $File, $Proxy, $content, $definition
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
        if ([string]::IsNullOrWhiteSpace($line) -or ($line.StartsWith('#') -and -not $line.StartsWith('#HttpOnly_'))) {
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
            -not $targetHost.EndsWith(('.{0}' -f $normalizedDomain), [System.StringComparison]::OrdinalIgnoreCase)
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
        [switch]$IncludeAllDomains,
        [switch]$AsCookies
    )

    $entries = New-Object System.Collections.Generic.List[string]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $cookies = [System.Collections.ArrayList]@()

    try {
        $targetHost = $CookieUri.Host.TrimStart('.')
        foreach ($cookie in (Get-WebSessionCookies -WebSession $WebSession)) {
            $normalizedDomain = $cookie.Domain.TrimStart('.')
            if (-not $IncludeAllDomains) {
                if ($normalizedDomain -ine $targetHost -and -not $targetHost.EndsWith(('.{0}' -f $normalizedDomain), [System.StringComparison]::OrdinalIgnoreCase) -and -not $normalizedDomain.EndsWith(('.{0}' -f $targetHost), [System.StringComparison]::OrdinalIgnoreCase)) {
                    continue
                }
            }

            $null = $cookies.Add($cookie)
            if ($seen.Add([string]$cookie.Name)) {
                $entries.Add(('{0}={1}' -f [string]$cookie.Name, [string]$cookie.Value)) | Out-Null
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
                    $null = $cookies.Add($cookie)
                    $entries.Add(('{0}={1}' -f [string]$cookie.Name, [string]$cookie.Value)) | Out-Null
                }
            }
        }
        catch {
            Write-Verbose ('Unable to read host-scoped cookies from web session for {0}: {1}' -f $CookieUri.Host, $_.Exception.Message)
        }
    }

    if ($AsCookies) { return $cookies.ToArray() }
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

    $originMatch = [regex]::Match($cacheContent, '(?m)^#\s*auth_origin:\s*(\S+)')
    if ($originMatch.Success -and $originMatch.Groups[1].Value -ne $CookieUri.GetLeftPart([System.UriPartial]::Authority)) {
        Write-Verbose 'Cached authentication belongs to a different portal origin.'
        return $null
    }

    # Restore cached bearer token if present and still valid.
    $bearerMatch = [regex]::Match($cacheContent, '(?m)^#\s*msa_bearer:\s*(\S+)')
    if ($bearerMatch.Success -and $originMatch.Success) {
        $cachedToken = $bearerMatch.Groups[1].Value.Trim()
        if (Test-JwtTokenValid -Token $cachedToken -MinValidityMinutes $script:MinTokenValidityMinutes) {
            $script:MSABearerToken = $cachedToken
            Write-Verbose ('Restored MSA bearer token from auth cache (still valid).')
        }
        else {
            Write-Verbose ('Cached MSA bearer token has expired; will require fresh sign-in for token.')
        }
    }

    $cookieHeader = Convert-CacheContentToCookieHeader -CacheContent $cacheContent -CookieUri $CookieUri
    if ([string]::IsNullOrWhiteSpace($cookieHeader)) {
        if ($script:MSABearerToken) {
            return New-Object Microsoft.PowerShell.Commands.WebRequestSession
        }
        Write-Warning ('Auth cache file {0} did not contain usable cookies for {1}.' -f $CachePath, $CookieUri.Host)
        return $null
    }

    if (Test-IsNetscapeCookieCacheContent -CacheContent $cacheContent) {
        $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        foreach ($line in ($cacheContent -split "`r?`n")) {
            if ([string]::IsNullOrWhiteSpace($line) -or ($line.StartsWith('#') -and -not $line.StartsWith('#HttpOnly_'))) { continue }
            $parts = $line -split "`t", 7
            if ($parts.Count -ne 7) { continue }
            $domain = $parts[0] -replace '^#HttpOnly_', ''
            $normalizedDomain = $domain.TrimStart('.')
            if ($normalizedDomain -ine $CookieUri.Host -and -not $CookieUri.Host.EndsWith('.' + $normalizedDomain, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
            try {
                $cookie = New-Object System.Net.Cookie($parts[5], $parts[6], $parts[2], $domain)
                $cookie.Secure = $parts[3] -eq 'TRUE'
                $cookie.HttpOnly = $parts[0].StartsWith('#HttpOnly_')
                if ([long]$parts[4] -gt 0) { $cookie.Expires = [DateTimeOffset]::FromUnixTimeSeconds([long]$parts[4]).UtcDateTime }
                if (-not $cookie.Expired) { $session.Cookies.Add($cookie) }
            }
            catch { Write-Verbose 'Skipping an invalid cached cookie.' }
        }
    }
    else {
        $session = New-WebSessionFromCookieHeader -CookieHeader $cookieHeader -CookieUri $CookieUri
    }
    $portalMatch = [regex]::Match($cacheContent, '(?m)^#\s*portal_page:\s*(\S+)')
    if ($portalMatch.Success) {
        $portalUri = $null
        if ([uri]::TryCreate($portalMatch.Groups[1].Value, [System.UriKind]::Absolute, [ref]$portalUri) -and $portalUri.GetLeftPart([System.UriPartial]::Authority) -eq $CookieUri.GetLeftPart([System.UriPartial]::Authority)) {
            Set-ObjectPropertyValue -Object $session -Name 'EventPortalUri' -Value $portalUri.GetLeftPart([System.UriPartial]::Path)
        }
    }
    $requestMatch = [regex]::Match($cacheContent, '(?m)^#\s*protected_request:\s*(\S+)')
    if ($requestMatch.Success) {
        try {
            $protectedRequest = ConvertTo-SecureString -String $requestMatch.Groups[1].Value -ErrorAction Stop
            $requestText = [System.Management.Automation.PSCredential]::new('event', $protectedRequest).GetNetworkCredential().Password
            $detailRequests = @{}
            if ($requestText.TrimStart().StartsWith('{')) {
                $requestState = $requestText | ConvertFrom-Json
                $requestText = [string]$requestState.RequestUri
                foreach ($property in $requestState.DetailRequests.PSObject.Properties) {
                    $detailUri = $null
                    if ([uri]::TryCreate([string]$property.Value, [System.UriKind]::Absolute, [ref]$detailUri) -and $detailUri.GetLeftPart([System.UriPartial]::Authority) -eq $CookieUri.GetLeftPart([System.UriPartial]::Authority)) {
                        $detailRequests[$property.Name] = $detailUri.AbsoluteUri
                    }
                }
                Set-ObjectPropertyValue -Object $session -Name 'EventDetailRequests' -Value $detailRequests
            }
            $requestUri = $null
            if ([uri]::TryCreate($requestText, [System.UriKind]::Absolute, [ref]$requestUri) -and $requestUri.GetLeftPart([System.UriPartial]::Authority) -eq $CookieUri.GetLeftPart([System.UriPartial]::Authority)) {
                Set-ObjectPropertyValue -Object $session -Name 'EventRequestUri' -Value $requestUri.AbsoluteUri
            }
        }
        catch { Write-Verbose 'Cached event request context could not be decrypted; browser discovery may be required.' }
    }
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

    if ($script:CustomCookieFile) { return }

    $cookieHeaderCandidates = New-Object System.Collections.Generic.List[string]
    $hasExplicitCookieHeader = -not [string]::IsNullOrWhiteSpace($CookieHeader)
    $fingerprintText = $null

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

        $cacheContentToPersist = Convert-CookieHeaderToNetscapeCookieFileContent -CookieHeader $cookieHeaderToPersist -CookieUri $CookieUri
        if ($WebSession -and -not $hasExplicitCookieHeader) {
            $sessionCookies = @(Get-WebSessionCookieHeader -WebSession $WebSession -CookieUri $CookieUri -AsCookies)
            if ($sessionCookies.Count -gt 0) {
                $cookieLines = @('# Netscape HTTP Cookie File')
                foreach ($cookie in $sessionCookies) {
                    $expiry = if ($cookie.Expires -eq [datetime]::MinValue) { 0 } else { ([DateTimeOffset]$cookie.Expires.ToUniversalTime()).ToUnixTimeSeconds() }
                    $domain = if ($cookie.HttpOnly) { '#HttpOnly_{0}' -f $cookie.Domain } else { $cookie.Domain }
                    $cookieLines += "{0}`t{1}`t{2}`t{3}`t{4}`t{5}`t{6}" -f $domain, $(if ($cookie.Domain.StartsWith('.')) { 'TRUE' } else { 'FALSE' }), $cookie.Path, $(if ($cookie.Secure) { 'TRUE' } else { 'FALSE' }), $expiry, $cookie.Name, $cookie.Value
                }
                $cacheContentToPersist = $cookieLines -join [Environment]::NewLine
            }
        }
        if ([string]::IsNullOrWhiteSpace($cacheContentToPersist)) {
            $cacheContentToPersist = $cookieHeaderToPersist
        }
        else {
            $cacheContentToPersist += '{0}# auth_origin: {1}' -f [Environment]::NewLine, $CookieUri.GetLeftPart([System.UriPartial]::Authority)
            if ($WebSession -and $WebSession.EventPortalUri) {
                $portalUri = [uri]$WebSession.EventPortalUri
                if ($portalUri.GetLeftPart([System.UriPartial]::Authority) -eq $CookieUri.GetLeftPart([System.UriPartial]::Authority)) {
                    $cacheContentToPersist += '{0}# portal_page: {1}' -f [Environment]::NewLine, $portalUri.GetLeftPart([System.UriPartial]::Path)
                }
            }
            $fingerprintText = $cacheContentToPersist
            if ($WebSession -and ($WebSession.EventRequestUri -or $WebSession.EventDetailRequests)) {
                $requestText = [string]$WebSession.EventRequestUri
                $fingerprintText += "`n" + $requestText
                if ($WebSession.EventDetailRequests) {
                    $requestText = @{ RequestUri = $requestText; DetailRequests = $WebSession.EventDetailRequests } | ConvertTo-Json -Depth 4 -Compress
                    # Detail templates differ only by session ID; the query holds the reusable context.
                    $fingerprintText += "`n" + ((@($WebSession.EventDetailRequests.Values | ForEach-Object { ([uri][string]$_).Query }) | Sort-Object -Unique) -join "`n")
                }
                $protectedRequest = ConvertFrom-SecureString -SecureString (ConvertTo-SecureString -String $requestText -AsPlainText -Force)
                $cacheContentToPersist += '{0}# protected_request: {1}' -f [Environment]::NewLine, $protectedRequest
            }
            # Persist a valid bearer token so the next session can reuse it without re-authenticating.
            if (-not [string]::IsNullOrWhiteSpace($script:MSABearerToken) -and
                (Test-JwtTokenValid -Token $script:MSABearerToken -MinValidityMinutes $script:MinTokenValidityMinutes)) {
                $cacheContentToPersist = ('{0}{1}# msa_bearer: {2}' -f $cacheContentToPersist, [Environment]::NewLine, $script:MSABearerToken)
                $fingerprintText += "`n" + $script:MSABearerToken
            }
        }
        if (-not $fingerprintText) { $fingerprintText = $cacheContentToPersist }
        $hash = [System.Security.Cryptography.SHA256]::Create()
        try { $fingerprint = [BitConverter]::ToString($hash.ComputeHash([System.Text.Encoding]::UTF8.GetBytes(('{0}|{1}' -f $CachePath, $fingerprintText)))) }
        finally { $hash.Dispose() }
        if ($fingerprint -eq $script:MSAAuthCacheFingerprint -and (Test-Path -LiteralPath $CachePath)) {
            Write-Debug 'MSA auth cache unchanged; skipping save.'
            return
        }

        if (Test-Path -LiteralPath $CachePath) {
            try { Copy-Item -LiteralPath $CachePath -Destination ('{0}.bak' -f $CachePath) -Force -ErrorAction SilentlyContinue } catch {}
        }
        Set-Content -LiteralPath $CachePath -Value $cacheContentToPersist -Encoding UTF8
        $script:MSAAuthCacheFingerprint = $fingerprint
        Write-Verbose ('Saved MSA auth cache to {0} in Netscape cookie format (length {1})' -f $CachePath, $cacheContentToPersist.Length)
    }
    else {
        # WAM auth produces no cookies — persist the Bearer token alone so the next run can skip re-auth.
        if (-not [string]::IsNullOrWhiteSpace($script:MSABearerToken) -and
            (Test-JwtTokenValid -Token $script:MSABearerToken -MinValidityMinutes $script:MinTokenValidityMinutes)) {
            if (Test-Path -LiteralPath $CachePath) {
                try { Copy-Item -LiteralPath $CachePath -Destination ('{0}.bak' -f $CachePath) -Force -ErrorAction SilentlyContinue } catch {}
            }
            $tokenOnlyCache = ('# Netscape HTTP Cookie File{0}# auth_origin: {1}{0}# msa_bearer: {2}' -f [Environment]::NewLine, $CookieUri.GetLeftPart([System.UriPartial]::Authority), $script:MSABearerToken)
            Set-Content -LiteralPath $CachePath -Value $tokenOnlyCache -Encoding UTF8
            Write-Verbose ('Saved WAM Bearer token to MSA auth cache {0}' -f $CachePath)
        }
        else {
            Write-Warning ('Unable to save MSA auth cache for {0}; cookie header was empty.' -f $CookieUri.Host)
        }
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
        [System.Exception]$Exception = $null,
        [string]$RequestUri = $null
    )

    $statusCode = $null
    $location = $null
    $responseUri = $null
    $responseContent = $null
    $responseContentType = $null

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
            try {
                $responseContentType = [string]($Response.Headers['Content-Type'])
            }
            catch {
                $responseContentType = $null
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

    # SPA portals (e.g. summit.microsoft.com) serve their frontend shell (text/html, HTTP 200) for
    # all unauthenticated requests, including API paths. Detect this as an auth failure so the caller
    # falls through to interactive sign-in instead of treating the HTML shell as a valid API response.
    if ($statusCode -eq 200 -and $responseContentType -imatch 'text/html') {
        $effectiveUri = if (-not [string]::IsNullOrWhiteSpace($responseUri)) { $responseUri }
        elseif (-not [string]::IsNullOrWhiteSpace($RequestUri)) { $RequestUri }
        else { $null }
        if ($effectiveUri -match '/api/') {
            return $true
        }
    }

    return $false
}

function Test-JwtTokenValid {
    param(
        [string]$Token,
        [int]$MinValidityMinutes = 5
    )

    if ([string]::IsNullOrWhiteSpace($Token)) { return $false }

    $parts = $Token.Split('.')
    if ($parts.Count -lt 2) { return $false }

    try {
        $base64 = $parts[1] -replace '-', '+' -replace '_', '/'
        $pad = 4 - ($base64.Length % 4)
        if ($pad -ne 4) { $base64 += ('=' * $pad) }
        $payloadBytes = [Convert]::FromBase64String($base64)
        $payloadJson = [System.Text.Encoding]::UTF8.GetString($payloadBytes)
        $payload = $payloadJson | ConvertFrom-Json
        if ($payload.PSObject.Properties.Match('exp').Count -gt 0) {
            $expiry = [DateTimeOffset]::FromUnixTimeSeconds([long]$payload.exp)
            return ($expiry -gt [DateTimeOffset]::UtcNow.AddMinutes($MinValidityMinutes))
        }
    }
    catch {}

    return $true
}

function Get-MSAAuthApiHeaders {
    $headers = @{ 'Accept' = 'application/json' }
    if (-not [string]::IsNullOrWhiteSpace($script:MSABearerToken)) {
        if (Test-JwtTokenValid -Token $script:MSABearerToken -MinValidityMinutes $script:MinTokenValidityMinutes) {
            $headers['Authorization'] = ('Bearer {0}' -f $script:MSABearerToken)
        }
        else {
            Write-Verbose 'Cached MSA bearer token has expired; excluding from request.'
            $script:MSABearerToken = $null
        }
    }
    return $headers
}

function Test-MSAWebSession {
    param(
        [Microsoft.PowerShell.Commands.WebRequestSession]$WebSession,
        [uri]$TargetUri,
        [uri]$Proxy,
        [scriptblock]$SessionValidator
    )

    if (-not $WebSession) { return $false }
    if ($SessionValidator) { return [bool](& $SessionValidator $WebSession) }
    $response = Invoke-WebRequest -Uri $TargetUri.AbsoluteUri -Method Head -WebSession $WebSession -Proxy $Proxy -Headers (Get-MSAAuthApiHeaders) -ErrorAction Stop -Verbose:$false
    return -not (Test-MSAAuthenticationRequired -Response $response -RequestUri $TargetUri.AbsoluteUri)
}

function Get-MSAAuthenticatedWebSession {
    param(
        [parameter(Mandatory = $true)][uri]$TargetUri,
        [uri]$Proxy,
        [switch]$ValidateCachedSession,
        [switch]$ForceInteractive,
        [switch]$PersistCache,
        [scriptblock]$SessionValidator,
        [scriptblock]$OnAuthenticatedBrowser
    )

    if ($script:CustomCookieFile) {
        $session = Get-CustomCookieWebSession -TargetUri $TargetUri
        try {
            if (($ValidateCachedSession -or $SessionValidator) -and
                -not (Test-MSAWebSession -WebSession $session -TargetUri $TargetUri -Proxy $Proxy -SessionValidator $SessionValidator)) {
                throw 'Cookie validation failed.'
            }
        }
        catch {
            if ($_.Exception -is [System.Management.Automation.PipelineStoppedException] -or $_.Exception.Data['CookieFileFailure']) { throw }
            if ($_.Exception.Data['FailureKind'] -in @('Http', 'Transport', 'Context', 'Response')) {
                $status = 0
                $hasStatus = [int]::TryParse([string]$_.Exception.Data['StatusCode'], [ref]$status) -and $status -ge 100 -and $status -le 599
                $cause = if ($hasStatus) { 'HTTP {0}' -f $status } else { [string]$_.Exception.Data['FailureKind'] }
                throw (New-CustomCookieFileException -Message ('Cookie-backed Custom request failed ({0}). Check the catalog endpoint and required portal context; no interactive authentication was attempted. Credentials and response details withheld.' -f $cause))
            }
            throw (New-CustomCookieFileException -Message 'Exported cookies could not authenticate the request or supply required portal context. Sign in and export cookies again; no interactive authentication was attempted.')
        }
        return $session
    }

    # Suppress IWR progress for all validation requests in this function.
    # The assignment is function-scoped and does not affect the caller's $ProgressPreference.
    $ProgressPreference = 'SilentlyContinue'

    $cookieUri = [uri]($TargetUri.GetLeftPart([System.UriPartial]::Authority) + '/')
    $origin = $TargetUri.GetLeftPart([System.UriPartial]::Authority)
    $authCachePath = Join-Path $PSScriptRoot 'MSAAuth.cache'
    $hostKey = $cookieUri.Host.ToLowerInvariant()
    $authMutex = $null
    $lockAcquired = $false
    $portalPageUri = $null

    if ($ForceInteractive -or $script:MSAAuthOrigin -ne $origin) {
        $script:MSAAuthWebSession = $null
        $script:MSABearerToken = $null
    }
    $script:MSAAuthOrigin = $origin

    if ($script:MSAAuthWebSession) {
        if (-not $ValidateCachedSession) {
            return $script:MSAAuthWebSession
        }

        try {
            if (-not (Test-MSAWebSession -WebSession $script:MSAAuthWebSession -TargetUri $TargetUri -Proxy $Proxy -SessionValidator $SessionValidator)) {
                throw 'In-memory session redirected to a sign-in page.'
            }
            Write-Verbose ('Using in-memory MSA authentication for {0}' -f $cookieUri.Host)
            if ($PersistCache) { Save-MSAAuthCookieHeader -CachePath $authCachePath -CookieUri $cookieUri -WebSession $script:MSAAuthWebSession }
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

        $session = if ($ForceInteractive) { $null } else { Get-CachedMSAWebSession -CachePath $authCachePath -CookieUri $cookieUri }
        if ($session -and $session.EventPortalUri) { $portalPageUri = [uri]$session.EventPortalUri }

        if ($session -and $ValidateCachedSession) {
            try {
                if (-not (Test-MSAWebSession -WebSession $session -TargetUri $TargetUri -Proxy $Proxy -SessionValidator $SessionValidator)) {
                    throw 'Cached session redirected to a sign-in page.'
                }
                Write-Host ('Using cached Microsoft authentication for {0}' -f $cookieUri.Host)
            }
            catch {
                Write-Warning 'Cached Microsoft authentication or event request context could not be validated; browser sign-in may be required.'
                $session = $null
            }
        }

        # If -CookiesFromBrowser was specified, try importing cookies from that browser via yt-dlp
        # before falling back to the silent/interactive sign-in flows.
        if (-not $session -and -not [string]::IsNullOrWhiteSpace($CookiesFromBrowser)) {
            $browserImportSession = Get-MSAAuthFromBrowserCookies -Browser $CookiesFromBrowser -CookieUri $cookieUri -YtDlpPath $script:YouTubeDL -CachePath $authCachePath -Proxy $Proxy
            if ($browserImportSession) {
                try {
                    if (Test-MSAWebSession -WebSession $browserImportSession -TargetUri $TargetUri -Proxy $Proxy -SessionValidator $SessionValidator) {
                        Write-Host ('Using cookies imported from {0} for {1}.' -f $CookiesFromBrowser, $CookieUri.Host)
                        $session = $browserImportSession
                    }
                    else {
                        Write-Warning ('Cookies from {0} did not authenticate with {1}; falling back to sign-in dialog.' -f $CookiesFromBrowser, $CookieUri.Host)
                    }
                }
                catch {
                    Write-Warning ('Could not validate {0} cookies for {1}: {2}' -f $CookiesFromBrowser, $CookieUri.Host, $_.Exception.Message)
                }
            }
        }

        if (-not $session) {
            Write-Verbose ('Attempting silent MSA sign-in bootstrap for {0}' -f $cookieUri.Host)
            $session = if ($OnAuthenticatedBrowser -or $ForceInteractive) { $null } else { Get-MSAAuthWebSessionSilent -StartUri $TargetUri -CookieUri $cookieUri -Proxy $Proxy }
            if ($session) {
                $silentSessionValid = $true
                $silentCookieCount = 0
                try {
                    if (-not (Test-MSAWebSession -WebSession $session -TargetUri $TargetUri -Proxy $Proxy -SessionValidator $SessionValidator)) {
                        throw 'Silent session validation indicates authentication is required.'
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
                }
                elseif ($silentCookieCount -gt 0) {
                    Write-Verbose ('Captured {0} silent-session cookie(s) for {1}; trying OAuth implicit flow with collected live.com cookies before interactive sign-in.' -f $silentCookieCount, $cookieUri.Host)

                    # The bootstrap WebBrowser visited login.live.com — if the machine has a live.com
                    # WININET session (Windows MSA sign-in), those cookies are in $session now.
                    # Follow portal → AAD → live.com redirect chain; live.com silently issues id_token.
                    $implicitSession = Invoke-MSAImplicitFlowSignIn -CookieUri $cookieUri -LiveComSession $session -Proxy $Proxy
                    if ($implicitSession) {
                        try {
                            if (Test-MSAWebSession -WebSession $implicitSession -TargetUri $TargetUri -Proxy $Proxy -SessionValidator $SessionValidator) {
                                Write-Verbose ('OAuth implicit flow sign-in succeeded for {0}.' -f $cookieUri.Host)
                                $session = $implicitSession
                                $silentSessionValid = $true
                            }
                            else {
                                Write-Verbose 'OAuth implicit flow: API still requires sign-in; falling back to interactive.'
                                $script:MSABearerToken = $null
                                $session = $null
                            }
                        }
                        catch {
                            Write-Verbose ('OAuth implicit flow validation failed: {0}; falling back to interactive.' -f $_.Exception.Message)
                            $script:MSABearerToken = $null
                            $session = $null
                        }
                    }
                    else {
                        $session = $null
                    }
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
                if ($OnAuthenticatedBrowser) {
                    Write-Host 'Opening event portal browser. Sign in only if the portal requires it, then open the sessions page.'
                }
                else {
                    Write-Host 'Opening Microsoft sign-in dialog for event authentication (personal or work/school account).'
                }
                $session = Get-MSAAuthWebSession -StartUri $TargetUri -CookieUri $cookieUri -Proxy $Proxy -OnAuthenticatedBrowser $OnAuthenticatedBrowser -PortalUri $portalPageUri

                $interactiveSessionValid = $true
                $interactiveCookieCount = 0
                try {
                    if (-not (Test-MSAWebSession -WebSession $session -TargetUri $TargetUri -Proxy $Proxy -SessionValidator $SessionValidator)) {
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
                    if ($SessionValidator) { return $null }
                    if ($interactiveCookieCount -gt 0) {
                        Write-Verbose ('Proceeding with {0} captured cookie(s) for {1}; API access may be limited.' -f $interactiveCookieCount, $cookieUri.Host)
                    }
                    else {
                        return $null
                    }
                }

            }
        }

        if ($PersistCache -and $session) { Save-MSAAuthCookieHeader -CachePath $authCachePath -CookieUri $cookieUri -WebSession $session }
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

    if ($script:MSAContentAuthRequired) {
        if ($script:CustomCookieFile) {
            $invokeParams.WebSession = Get-CustomCookieWebSession -TargetUri $requestUri
        }
        if ($script:MSAAuthWebSession) {
            $invokeParams.WebSession = $script:MSAAuthWebSession
        }
        if (-not [string]::IsNullOrWhiteSpace($script:MSABearerToken) -and
            (Test-JwtTokenValid -Token $script:MSABearerToken -MinValidityMinutes $script:MinTokenValidityMinutes)) {
            if (-not $invokeParams.ContainsKey('Headers')) { $invokeParams.Headers = @{} }
            if (-not $invokeParams.Headers.ContainsKey('Authorization')) {
                $invokeParams.Headers['Authorization'] = ('Bearer {0}' -f $script:MSABearerToken)
            }
        }
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
        $savedProgressPref = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        try { $response = Invoke-WebRequest @invokeParams } finally { $ProgressPreference = $savedProgressPref }

        if (Test-MSAAuthenticationRequired -Response $response -RequestUri $Uri) {
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
        if ($webException -is [System.Management.Automation.PipelineStoppedException]) { throw }
        if (-not (Test-MSAAuthenticationRequired -Exception $webException)) {
            if ($script:CustomCookieFile) {
                throw (New-CustomCookieFileException -Message 'Custom content request failed; credentials and response details withheld.')
            }
            return $null
        }

        if ($script:CustomCookieFile -and $invokeParams.WebSession) {
            throw (New-CustomCookieFileException -Message 'Custom content rejected the exported cookies. Sign in and export cookies again; no interactive authentication was attempted.')
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
            $ProgressPreference = 'SilentlyContinue'
            $response = Invoke-WebRequest @invokeParams
            if ($script:CustomCookieFile -and (Test-MSAAuthenticationRequired -Response $response -RequestUri $Uri)) {
                throw (New-CustomCookieFileException -Message 'Custom content rejected the exported cookies. Sign in and export cookies again.')
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
            if ($script:CustomCookieFile) {
                if ($_.Exception -is [System.Management.Automation.PipelineStoppedException] -or $_.Exception.Data['CookieFileFailure']) { throw }
                throw (New-CustomCookieFileException -Message 'Custom content request failed with exported cookies; credentials and response details withheld.')
            }
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

    # Wrap with @() so a single unique result stays an array and += works on all PS versions.
    $cookieUris = @(@(
            $CookieUri,
            [uri]('{0}://{1}/' -f $CookieUri.Scheme, $CookieUri.Host),
            $StartUri
        ) | Select-Object -Unique)

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

        # Navigate to the portal root, not the API endpoint, so the SPA login flow is triggered.
        $portalRootUri = [uri]('{0}://{1}/' -f $CookieUri.Scheme, $CookieUri.Host)
        $browser.Navigate($portalRootUri.AbsoluteUri)

        $deadline = (Get-Date).AddSeconds([Math]::Max(3, $TimeoutSeconds))
        while ((Get-Date) -lt $deadline -and -not $state.Completed) {
            [System.Windows.Forms.Application]::DoEvents()
            [System.Threading.Thread]::Sleep(100)
        }

        if (-not $state.Completed) {
            return $null
        }

        return New-MSAWebSessionFromBrowserContext -CookieUri $CookieUri -StartUri $portalRootUri -Browser $browser -VisitedUrls $visitedUrls -Proxy $Proxy
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

function Invoke-MSAImplicitFlowSignIn {
    # Follow the portal's own AAD→live.com redirect chain with live.com cookies already in the session.
    # login.live.com silently issues an id_token form_post when the MSA session is valid.
    # We parse the form, POST it to /signin-oidc, and optionally use the id_token as Bearer.
    param(
        [parameter(Mandatory = $true)][uri]$CookieUri,
        [parameter(Mandatory = $true)][Microsoft.PowerShell.Commands.WebRequestSession]$LiveComSession,
        [uri]$Proxy
    )

    $portalRoot = [uri]('{0}://{1}/' -f $CookieUri.Scheme, $CookieUri.Host)
    $liveUri = [uri]'https://login.live.com/'
    $aadUri = [uri]'https://login.microsoftonline.com/'

    # Seed a fresh session with only the live.com (and AAD) cookies from the bootstrap.
    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $session.UserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36'
    foreach ($srcUri in @($liveUri, $aadUri)) {
        foreach ($c in @($LiveComSession.Cookies.GetCookies($srcUri))) {
            try { $session.Cookies.Add($srcUri, $c) } catch {}
        }
    }

    $liveCookieCount = ($session.Cookies.GetCookies($liveUri) | Measure-Object).Count
    if ($liveCookieCount -eq 0) {
        Write-Verbose 'OAuth implicit flow: no login.live.com cookies available — skipping.'
        return $null
    }
    Write-Verbose ('OAuth implicit flow: starting portal redirect chain with {0} live.com cookie(s).' -f $liveCookieCount)

    $invokeParams = @{
        Uri                = $portalRoot
        WebSession         = $session
        MaximumRedirection = 15
        UseBasicParsing    = $true
        UserAgent          = $session.UserAgent
        ErrorAction        = 'Stop'
        Verbose            = $false
    }
    if ($Proxy) { $invokeParams.Proxy = $Proxy }

    $html = $null
    try {
        $resp = Invoke-WebRequest @invokeParams
        $html = [string]$resp.Content
    }
    catch {
        # AAD or live.com sometimes returns a non-2xx status on the form_post page itself.
        try {
            $errorResponse = $_.Exception.Response
            if ($errorResponse -is [System.Net.HttpWebResponse]) {
                $reader = [System.IO.StreamReader]::new($errorResponse.GetResponseStream())
                try { $html = $reader.ReadToEnd() }
                finally { $reader.Dispose(); $errorResponse.Close() }
            }
            else { $html = [string]$errorResponse.Content.ReadAsStringAsync().GetAwaiter().GetResult() }
        } catch {}
        if ([string]::IsNullOrWhiteSpace($html)) {
            Write-Verbose ('OAuth implicit flow: redirect chain failed: {0}' -f $_.Exception.Message)
            return $null
        }
    }

    # The form_post page from live.com has action="…/signin-oidc" with hidden id_token field.
    if ($html -notmatch '(?i)action=[''"]([^''"]*/signin-oidc[^''"]*)[''"]') {
        Write-Verbose 'OAuth implicit flow: no signin-oidc form_post in response — live.com cookies expired or sign-in needed.'
        return $null
    }
    $formAction = $Matches[1]

    # Parse all hidden input fields.
    $formBody = [ordered]@{}
    foreach ($m in [regex]::Matches($html, '(?i)<input\b[^>]*\btype=[''"]hidden[''"][^>]*/?>')) {
        $n = ([regex]::Match($m.Value, '(?i)\bname=[''"]([^''"]+)[''"]')).Groups[1].Value
        $v = ([regex]::Match($m.Value, '(?i)\bvalue=[''"]([^''"]*)')).Groups[1].Value
        if ($n) { $formBody[$n] = [System.Net.WebUtility]::HtmlDecode($v) }
    }
    # Also catch name= before type= ordering variant
    foreach ($m in [regex]::Matches($html, '(?i)<input\b[^>]*\bname=[''"]([^''"]+)[''"][^>]*\bvalue=[''"]([^''"]*)')) {
        $n = $m.Groups[1].Value
        if ($n -and -not $formBody.Contains($n)) { $formBody[$n] = [System.Net.WebUtility]::HtmlDecode($m.Groups[2].Value) }
    }

    if ($formBody.Count -eq 0) {
        Write-Verbose 'OAuth implicit flow: form_post matched but no fields parsed.'
        return $null
    }

    # The id_token from the form POST is a JWT — try it as a Bearer token candidate.
    # Some portal APIs validate the MSA id_token directly in the Authorization header.
    $idToken = [string]$formBody['id_token']
    if ($idToken -match '^eyJ') {
        Write-Verbose 'OAuth implicit flow: id_token present — storing as Bearer token candidate.'
        $script:MSABearerToken = $idToken
    }

    Write-Verbose ('OAuth implicit flow: POSTing {0} field(s) to {1}' -f $formBody.Count, $formAction)
    $postParams = @{
        Uri                = $formAction
        Method             = 'POST'
        Body               = $formBody
        WebSession         = $session
        MaximumRedirection = 10
        UseBasicParsing    = $true
        UserAgent          = $session.UserAgent
        ErrorAction        = 'SilentlyContinue'
        Verbose            = $false
    }
    if ($Proxy) { $postParams.Proxy = $Proxy }

    try { Invoke-WebRequest @postParams | Out-Null } catch {
        Write-Verbose ('OAuth implicit flow: POST to signin-oidc: {0}' -f $_.Exception.Message)
    }

    $portalCookieCount = ($session.Cookies.GetCookies($portalRoot) | Measure-Object).Count
    Write-Verbose ('OAuth implicit flow: session now has {0} portal cookie(s).' -f $portalCookieCount)
    return $session
}

function Get-EdgeExecutablePath {
    $candidates = @(
        (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path $env:LOCALAPPDATA 'Microsoft\Edge\Application\msedge.exe')
    )
    foreach ($c in $candidates) {
        if (Test-Path -LiteralPath $c) { return $c }
    }
    $fromPath = Get-Command msedge.exe -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source
    if ($fromPath -and (Test-Path -LiteralPath $fromPath)) { return $fromPath }
    return $null
}

function Invoke-EdgeCDPCommand {
    param(
        [parameter(Mandatory = $true)][string]$WebSocketUrl,
        [parameter(Mandatory = $true)][string]$Method,
        [hashtable]$Params = @{},
        [int]$TimeoutSeconds = 15
    )

    $ws = New-Object System.Net.WebSockets.ClientWebSocket
    $cts = New-Object System.Threading.CancellationTokenSource
    $cts.CancelAfter([Math]::Max(1, $TimeoutSeconds) * 1000)
    try {
        $ws.ConnectAsync([uri]$WebSocketUrl, $cts.Token).GetAwaiter().GetResult()
        $payload = @{ id = 1; method = $Method; params = $Params } | ConvertTo-Json -Compress -Depth 5
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($payload)
        $ws.SendAsync([ArraySegment[byte]]::new($bytes), [System.Net.WebSockets.WebSocketMessageType]::Text, $true, $cts.Token).GetAwaiter().GetResult()
        $receiveBuffer = New-Object byte[] 65536
        while (-not $cts.IsCancellationRequested) {
            $message = New-Object System.IO.MemoryStream
            try {
                do {
                    $result = $ws.ReceiveAsync([ArraySegment[byte]]::new($receiveBuffer), $cts.Token).GetAwaiter().GetResult()
                    if ($result.MessageType -eq [System.Net.WebSockets.WebSocketMessageType]::Close) { throw 'CDP connection closed.' }
                    $message.Write($receiveBuffer, 0, $result.Count)
                } while (-not $result.EndOfMessage)
                $response = [System.Text.Encoding]::UTF8.GetString($message.ToArray()) | ConvertFrom-Json
            }
            finally { $message.Dispose() }
            if ($response.id -eq 1) {
                if ($response.error) { throw 'CDP command was rejected.' }
                return $response
            }
        }
        throw 'CDP command timed out.'
    }
    catch { throw ('CDP command {0} failed or timed out; response details withheld.' -f $Method) }
    finally {
        $cts.Cancel()
        $ws.Dispose()
        $cts.Dispose()
    }
}

function New-MSAWebSessionFromCDPCookies {
    param(
        [parameter(Mandatory = $true)][uri]$CookieUri,
        [parameter(Mandatory = $true)][array]$Cookies,
        [uri]$Proxy
    )

    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $targetHost = $CookieUri.Host.ToLowerInvariant()
    $addedCount = 0

    foreach ($c in $Cookies) {
        $cookieDomain = ([string]$c.domain).TrimStart('.').ToLowerInvariant()
        if ($cookieDomain -ine $targetHost -and -not $targetHost.EndsWith(('.{0}' -f $cookieDomain), [System.StringComparison]::OrdinalIgnoreCase)) {
            continue
        }

        $netCookie = New-Object System.Net.Cookie
        $netCookie.Name = [string]$c.name
        $netCookie.Value = [string]$c.value
        $netCookie.Path = if ([string]::IsNullOrWhiteSpace([string]$c.path)) { '/' } else { [string]$c.path }
        $netCookie.Domain = [string]$c.domain
        $netCookie.Secure = [bool]$c.secure
        $netCookie.HttpOnly = [bool]$c.httpOnly
        if ($c.expires -gt 0) { $netCookie.Expires = [DateTimeOffset]::FromUnixTimeSeconds([long]$c.expires).UtcDateTime }

        try {
            $session.Cookies.Add($CookieUri, $netCookie)
            $addedCount++
        }
        catch {
            Write-Verbose ('Skipping CDP cookie {0}: {1}' -f $netCookie.Name, $_.Exception.Message)
        }
    }

    Write-Verbose ('Constructed CDP-backed session with {0} cookies for {1}.' -f $addedCount, $targetHost)

    try {
        Invoke-WebRequest -Method Head -Uri $CookieUri -WebSession $session -Proxy $Proxy -ErrorAction SilentlyContinue -Verbose:$false | Out-Null
    }
    catch {
        Write-Verbose ('Session warmup request returned: {0}' -f $_.Exception.Message)
    }

    return $session
}

function Get-MSALAccessTokenFromStorageJson {
    param(
        [string]$StorageJson
    )

    if ([string]::IsNullOrWhiteSpace($StorageJson) -or $StorageJson -eq 'null' -or $StorageJson -eq '{}') {
        return $null
    }

    try {
        $storage = $StorageJson | ConvertFrom-Json
    }
    catch { return $null }

    $bestToken = $null
    $bestExpiry = $null

    foreach ($key in $storage.PSObject.Properties.Name) {
        $rawValue = [string]$storage.$key
        if ([string]::IsNullOrWhiteSpace($rawValue)) { continue }

        $token = $null
        $expiry = $null

        if ($rawValue.TrimStart().StartsWith('{')) {
            try {
                $tokenObj = $rawValue | ConvertFrom-Json
                if ($tokenObj.PSObject.Properties.Match('secret').Count -gt 0) {
                    $s = [string]$tokenObj.secret
                    if (-not [string]::IsNullOrWhiteSpace($s) -and $s.StartsWith('eyJ')) {
                        $token = $s
                    }
                }
                if ($tokenObj.PSObject.Properties.Match('expiresOn').Count -gt 0) {
                    try { $expiry = [DateTimeOffset]::FromUnixTimeSeconds([long]$tokenObj.expiresOn) } catch {}
                }
            }
            catch {}
        }
        elseif ($rawValue.StartsWith('eyJ')) {
            $token = $rawValue
        }

        if ($token -and (-not $bestToken -or ($expiry -and (-not $bestExpiry -or $expiry -gt $bestExpiry)))) {
            $bestToken = $token
            $bestExpiry = $expiry
        }
    }

    return $bestToken
}

function Get-MSALAccessTokenFromCdpPage {
    param(
        [parameter(Mandatory = $true)][int]$DebugPort,
        [parameter(Mandatory = $true)][uri]$CookieUri,
        [int]$TimeoutSeconds = 10
    )

    try {
        $cdpPagesJson = Invoke-RestMethod -Uri ('http://127.0.0.1:{0}/json' -f $DebugPort) -Method Get -ErrorAction Stop -Verbose:$false
    }
    catch {
        Write-Verbose ('Unable to list CDP pages on port {0}: {1}' -f $DebugPort, $_.Exception.Message)
        return $null
    }

    $targetPages = New-Object System.Collections.Generic.List[psobject]
    foreach ($cdpPage in @($cdpPagesJson)) {
        if ($cdpPage.type -ne 'page' -or -not $cdpPage.webSocketDebuggerUrl -or -not $cdpPage.url) { continue }
        try {
            if (([uri]$cdpPage.url).Host -ieq $CookieUri.Host) {
                $targetPages.Add($cdpPage) | Out-Null
            }
        }
        catch {}
    }

    if ($targetPages.Count -eq 0) {
        Write-Verbose ('No CDP pages found for host {0} to extract MSAL token from.' -f $CookieUri.Host)
        return $null
    }

    # Extract only accesstoken-keyed entries to minimise data transfer
    $jsExpression = @'
(function(){var r={},s=[window.sessionStorage,window.localStorage];for(var i=0;i<s.length;i++){try{var t=s[i];for(var j=0;j<t.length;j++){var k=t.key(j);if(k&&k.toLowerCase().indexOf('accesstoken')!==-1){r[k]=t.getItem(k);}}}catch(e){}}return JSON.stringify(r);})()
'@

    foreach ($page in $targetPages) {
        $pageWsUrl = [string]$page.webSocketDebuggerUrl
        Write-Verbose ('Querying page storage for MSAL token at {0}' -f $page.url)

        try {
            $evalResponse = Invoke-EdgeCDPCommand -WebSocketUrl $pageWsUrl -Method 'Runtime.evaluate' -Params @{
                expression    = $jsExpression
                returnByValue = $true
            } -TimeoutSeconds $TimeoutSeconds
        }
        catch {
            Write-Verbose ('CDP Runtime.evaluate failed for {0}: {1}' -f $page.url, $_.Exception.Message)
            continue
        }

        $storageJson = $null
        if ($evalResponse -and $evalResponse.result -and $evalResponse.result.result) {
            $storageJson = [string]$evalResponse.result.result.value
        }

        $token = Get-MSALAccessTokenFromStorageJson -StorageJson $storageJson
        if ($token) {
            Write-Verbose ('Extracted MSAL access token from page {0}' -f $page.url)
            return $token
        }
    }

    Write-Verbose ('No MSAL access token found in page storage for {0}.' -f $CookieUri.Host)
    return $null
}

function Get-MSAAuthFromBrowserCookies {
    param(
        [parameter(Mandatory = $true)][string]$Browser,
        [parameter(Mandatory = $true)][uri]$CookieUri,
        [parameter(Mandatory = $true)][string]$YtDlpPath,
        [parameter(Mandatory = $true)][string]$CachePath,
        [uri]$Proxy
    )

    if (-not (Test-Path -LiteralPath $YtDlpPath)) {
        Write-Verbose 'yt-dlp not found; cannot import browser cookies.'
        return $null
    }

    Write-Host ('Importing {0} cookies for {1} via yt-dlp...' -f $Browser, $CookieUri.Host)

    $tempCookieFile = Join-Path $env:TEMP ('GetEventSession_cookies_{0}.txt' -f [System.IO.Path]::GetRandomFileName())

    try {
        # --cookies FILE acts as both source and sink: yt-dlp writes the full cookie jar
        # (including the --cookies-from-browser import) back to the file on exit.
        $ytArgs = @(
            '--cookies-from-browser', $Browser,
            '--cookies', $tempCookieFile,
            '--simulate',
            '--ignore-errors',
            '--no-warnings',
            '--quiet',
            $CookieUri.AbsoluteUri
        )
        if ($Proxy) { $ytArgs += '--proxy', $Proxy.AbsoluteUri }

        $proc = Start-Process -FilePath $YtDlpPath -ArgumentList $ytArgs -Wait -PassThru -WindowStyle Hidden -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 500

        if (-not (Test-Path -LiteralPath $tempCookieFile)) {
            Write-Warning ('yt-dlp did not produce a cookie file for {0}; {1} browser cookies may not be accessible.' -f $CookieUri.Host, $Browser)
            return $null
        }

        $cacheContent = Get-Content -LiteralPath $tempCookieFile -Raw -ErrorAction SilentlyContinue
        if ([string]::IsNullOrWhiteSpace($cacheContent)) {
            Write-Warning ('Browser cookie file is empty; no cookies were exported from {0}.' -f $Browser)
            return $null
        }

        $cookieHeader = Convert-CacheContentToCookieHeader -CacheContent $cacheContent -CookieUri $CookieUri
        if ([string]::IsNullOrWhiteSpace($cookieHeader)) {
            Write-Warning ('No cookies found for {0} in the {1} browser profile.' -f $CookieUri.Host, $Browser)
            return $null
        }

        $session = New-WebSessionFromCookieHeader -CookieHeader $cookieHeader -CookieUri $CookieUri
        $cookieCount = ($session.Cookies.GetCookies($CookieUri) | Measure-Object).Count
        Write-Verbose ('Imported {0} cookie(s) for {1} from {2}.' -f $cookieCount, $CookieUri.Host, $Browser)

        Save-MSAAuthCookieHeader -CachePath $CachePath -CookieUri $CookieUri -CookieHeader $cookieHeader
        return $session
    }
    catch {
        Write-Warning ('Failed to import {0} browser cookies: {1}' -f $Browser, $_.Exception.Message)
        return $null
    }
    finally {
        if (Test-Path -LiteralPath $tempCookieFile) {
            Remove-Item -LiteralPath $tempCookieFile -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-MSAAuthFromCookiePaste {
    param(
        [parameter(Mandatory = $true)][uri]$CookieUri,
        [parameter(Mandatory = $true)][string]$CachePath,
        [uri]$Proxy
    )

    if (-not [Environment]::UserInteractive) { return $null }
    if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) { return $null }

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Manual Cookie Import'
    $form.Width = 680
    $form.Height = 500
    $form.StartPosition = 'CenterScreen'
    $form.ShowInTaskbar = $true
    $form.TopMost = $true

    $label = New-Object System.Windows.Forms.Label
    $label.Left = 12
    $label.Top = 12
    $label.Width = 644
    $label.Height = 130
    $label.Text = (
        "The embedded browser could not authenticate with {0}.`n`n" +
        "To sign in manually:`n" +
        "  1. Open Firefox / Chrome / Edge and sign in to {0}`n" +
        "  2. Press F12 to open Developer Tools`n" +
        "  3. Go to the Network tab and reload the page`n" +
        "  4. Click any request to {0} in the list`n" +
        "  5. In the request headers panel, find the 'Cookie:' line`n" +
        "  6. Copy the entire value (the long text after 'Cookie: ')`n" +
        "  7. Paste it into the box below and click Import"
    ) -f $CookieUri.Host

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Left = 12
    $textBox.Top = 148
    $textBox.Width = 644
    $textBox.Height = 256
    $textBox.Multiline = $true
    $textBox.ScrollBars = 'Vertical'
    $textBox.WordWrap = $true
    $textBox.Font = New-Object System.Drawing.Font('Consolas', 8)

    $importButton = New-Object System.Windows.Forms.Button
    $importButton.Text = 'Import'
    $importButton.Width = 100
    $importButton.Height = 28
    $importButton.Left = 452
    $importButton.Top = 416

    $skipButton = New-Object System.Windows.Forms.Button
    $skipButton.Text = 'Skip'
    $skipButton.Width = 100
    $skipButton.Height = 28
    $skipButton.Left = 560
    $skipButton.Top = 416

    $importButton.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.Close() })
    $skipButton.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $form.Close() })

    $form.Controls.AddRange(@($label, $textBox, $importButton, $skipButton))
    $form.Add_Shown({ $form.Activate(); $form.BringToFront(); $textBox.Focus() })

    $dialogResult = $form.ShowDialog()

    if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK) { return $null }

    $cookieString = $textBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($cookieString)) { return $null }

    $session = New-WebSessionFromCookieHeader -CookieHeader $cookieString -CookieUri $CookieUri
    if (-not $session) { return $null }

    Save-MSAAuthCookieHeader -CachePath $CachePath -CookieUri $CookieUri -CookieHeader $cookieString
    Write-Host ('Imported browser cookies for {0} and saved to auth cache.' -f $CookieUri.Host)
    return $session
}

function Resolve-OAuthChallenge {
    param(
        [parameter(Mandatory = $true)][uri]$TargetUri,
        [uri]$Proxy
    )

    $result = @{ ClientId = $null; TenantId = 'common'; Scope = $null; AccountType = 'Unknown' }

    $reqParams = @{ Uri = $TargetUri.AbsoluteUri; MaximumRedirection = 0; ErrorAction = 'Stop'; Verbose = $false }
    if ($Proxy) { $reqParams.Proxy = $Proxy }

    $location = $null
    $wwwAuth = $null

    try {
        $null = Invoke-WebRequestNoRedirect -Parameters $reqParams
    }
    catch {
        if ($_.Exception.Response) {
            $sc = [int]$_.Exception.Response.StatusCode
            if ($sc -in @(301, 302, 307, 308)) {
                $location = [string]($_.Exception.Response.Headers.Location | Select-Object -First 1)
            }
            elseif ($sc -eq 401) {
                $wwwAuth = [string]($_.Exception.Response.Headers['WWW-Authenticate'] | Select-Object -First 1)
            }
        }
    }

    # Follow up to 5 redirects if a 302 was not captured above (server may return 200 after redirect).
    if (-not $location -and -not $wwwAuth) {
        try {
            $followParams = @{ Uri = $TargetUri.AbsoluteUri; MaximumRedirection = 5; ErrorAction = 'Stop'; Verbose = $false }
            if ($Proxy) { $followParams.Proxy = $Proxy }
            $followed = Invoke-WebRequest @followParams
            $finalUri = $null
            if ($followed.PSObject.Properties.Match('BaseResponse').Count -gt 0 -and $followed.BaseResponse.ResponseUri) {
                $finalUri = $followed.BaseResponse.ResponseUri.AbsoluteUri
            }
            elseif ($followed.PSObject.Properties.Match('ResponseUri').Count -gt 0 -and $followed.ResponseUri) {
                $finalUri = $followed.ResponseUri.AbsoluteUri
            }
            if ($finalUri -and $finalUri -match '(?i)(microsoftonline\.com|login\.live\.com|login\.microsoft\.com)') {
                $location = $finalUri
            }
        }
        catch {}
    }

    if ($location) {
        if ($location -match '(?i)client_id=([^&#]+)') { $result.ClientId = [Uri]::UnescapeDataString($Matches[1].Trim()) }
        if ($location -match '(?i)/([a-f0-9\-]{36}|common|consumers|organizations)/') { $result.TenantId = $Matches[1] }
        if ($location -match '(?i)[?&]scope=([^&#]+)') { $result.Scope = [Uri]::UnescapeDataString($Matches[1].Trim()) }
    }

    if ($wwwAuth) {
        if ($wwwAuth -match 'client_id="([^"]+)"') { $result.ClientId = $Matches[1] }
        if ($wwwAuth -match 'authorization_uri="[^"]*/([a-f0-9\-]{36}|common|consumers|organizations)/') { $result.TenantId = $Matches[1] }
        if ($wwwAuth -match 'resource="([^"]+)"') { $result.Scope = $Matches[1] + '/.default' }
    }

    # For MSAL.js SPAs that return 200+HTML for all URLs (no redirects), scan the HTML for inline
    # MSAL configuration.  Also probe the portal root if the target was an API sub-path.
    if (-not $result.ClientId) {
        $scanUris = [System.Collections.Generic.List[string]]::new()
        $scanUris.Add($TargetUri.AbsoluteUri)
        $portalRoot = ('{0}://{1}/' -f $TargetUri.Scheme, $TargetUri.Host)
        if ($portalRoot -ne $TargetUri.AbsoluteUri) { $scanUris.Add($portalRoot) }

        foreach ($scanUri in $scanUris) {
            if ($result.ClientId) { break }
            try {
                $scanParams = @{ Uri = $scanUri; MaximumRedirection = 5; ErrorAction = 'Stop'; Verbose = $false }
                if ($Proxy) { $scanParams.Proxy = $Proxy }
                $scanResp = Invoke-WebRequest @scanParams
                $body = [string]$scanResp.Content
                # Match MSAL.js config patterns: clientId: "..." or "clientId":"..."
                if ($body -match '(?:clientId|client_id)[''"]?\s*[:=]\s*[''"]([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})[''"]') {
                    $result.ClientId = $Matches[1]
                    if ($body -match '"authority"\s*[=:]\s*[''"]https?://[^''"/]+/([a-f0-9\-]{36}|common|consumers|organizations)/?[''"]') {
                        $result.TenantId = $Matches[1]
                    }
                }
            }
            catch {}
        }
    }

    if ($location -match '(?i)^https://login\.live\.com/') {
        $result.TenantId = 'consumers'
    }
    if ($result.TenantId -in @('consumers', '9188040d-6c67-4c5b-b112-36a304b66dad')) {
        $result.AccountType = 'MSA'
    }
    elseif ($result.TenantId -and $result.TenantId -notin @('common', '')) {
        $result.AccountType = 'M365'
    }
    return $result
}

function Invoke-WAMViaPS5Subprocess {
    param(
        [parameter(Mandatory = $true)][string]$Authority,
        [parameter(Mandatory = $true)][string]$ClientId,
        [parameter(Mandatory = $true)][string]$Scope,
        [int]$TimeoutSeconds = 120
    )

    $ps5Path = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $ps5Path)) {
        Write-Verbose 'WAM subprocess: powershell.exe (v5.1) not found — skipping.'
        return $null
    }

    # This script runs inside powershell.exe (v5.1) where WinRT loads natively.
    # WAM's RequestTokenForWindowAsync requires an HWND owned by the calling process.
    # GetDesktopWindow() returns a system window that this process does not own, causing E_ACCESSDENIED.
    # Fix: create a tiny invisible WinForms form to obtain a real owned HWND.
    $ps5Script = @'
# Parameters are passed via environment variables to avoid command-line argument quoting issues.
trap { [Console]::Error.WriteLine('TRAP: ' + $_); exit 1 }
$Authority = $env:WAM_PS5_AUTHORITY
$ClientId  = $env:WAM_PS5_CLIENTID
$Scope     = $env:WAM_PS5_SCOPE
try {
    $null = [Windows.Security.Authentication.Web.Core.WebAuthenticationCoreManager, Windows.Security.Authentication.Web.Core, ContentType=WindowsRuntime]
    $null = [Windows.Security.Authentication.Web.Core.WebTokenRequest, Windows.Security.Authentication.Web.Core, ContentType=WindowsRuntime]
    $null = [Windows.Security.Authentication.Web.Core.WebTokenRequestStatus, Windows.Security.Authentication.Web.Core, ContentType=WindowsRuntime]
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop

    # AsTask() is a .NET extension method from System.Runtime.WindowsRuntime.dll; it is not
    # available as an instance method on the COM object PS5.1 returns for WinRT IAsyncOperation.
    # Load the DLL and call it as a static method instead.
    $rtDir    = [System.Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory()
    $winRtDll = Join-Path $rtDir 'System.Runtime.WindowsRuntime.dll'
    if (Test-Path $winRtDll) { Add-Type -Path $winRtDll -ErrorAction SilentlyContinue }

    # Resolve provider before the message loop starts (blocking .GetAwaiter().GetResult() is fine here).
    $provider = $null
    foreach ($provUri in @($Authority)) {
        $pt   = [Windows.Security.Authentication.Web.Core.WebAuthenticationCoreManager]::FindAccountProviderAsync($provUri)
        $prov = [System.WindowsRuntimeSystemExtensions]::AsTask($pt).GetAwaiter().GetResult()
        if ($prov) { $provider = $prov; break }
    }
    if (-not $provider) { [Console]::Error.WriteLine('WAM: no account provider found for authority ' + $Authority); exit 1 }

    # This process is launched with -STA so Application.Run is safe on the main thread.
    # WAM's async completions require a real Win32 message loop — DoEvents()-in-a-loop is not enough.
    $shared = @{ Token = $null; Err = $null }

    $form = New-Object System.Windows.Forms.Form
    $form.ShowInTaskbar   = $false
    $form.FormBorderStyle = 'None'
    $form.Opacity         = 0
    $form.Width           = 1
    $form.Height          = 1

    $form.add_Shown({
        # Message loop is now running — start the WAM request with our form's HWND as parent.
        $req         = [Windows.Security.Authentication.Web.Core.WebTokenRequest]::new($provider, $Scope, $ClientId)
        $requestTask = [System.WindowsRuntimeSystemExtensions]::AsTask(
            [Windows.Security.Authentication.Web.Core.WebAuthenticationCoreManager]::RequestTokenForWindowAsync($form.Handle, $req)
        )
        $deadline    = (Get-Date).AddSeconds(62)

        # WinForms Timer fires on the STA UI thread via WM_TIMER — keeps the pump free for WAM callbacks.
        $timer          = New-Object System.Windows.Forms.Timer
        $timer.Interval = 100
        $timer.add_Tick({
            $timedOut = (Get-Date) -gt $deadline
            if ($requestTask.IsCompleted -or $timedOut) {
                $timer.Stop()
                if ($requestTask.IsCompleted) {
                    try {
                        $r = $requestTask.Result
                        if ($r.ResponseStatus -eq [Windows.Security.Authentication.Web.Core.WebTokenRequestStatus]::Success) {
                            $shared.Token = [string]$r.ResponseData[0].Token
                        } elseif ($r.ResponseStatus -eq [Windows.Security.Authentication.Web.Core.WebTokenRequestStatus]::UserCancel) {
                            $shared.Err = 'UserCancel'
                        } else {
                            $errCode = if ($r.ResponseError) { $r.ResponseError.ErrorCode } else { '?' }
                            $errMsg  = if ($r.ResponseError) { $r.ResponseError.ErrorMessage } else { '' }
                            $shared.Err = 'WAM failed (' + $r.ResponseStatus + '): code=' + $errCode + ' ' + $errMsg
                        }
                    } catch { $shared.Err = $_.Exception.Message }
                } else {
                    $shared.Err = 'WAM request timed out.'
                }
                $form.Close()
            }
        })
        $timer.Start()
    })

    [System.Windows.Forms.Application]::Run($form)

    if ($shared.Err) { [Console]::Error.WriteLine($shared.Err); exit 1 }
    if ([string]::IsNullOrWhiteSpace($shared.Token)) { [Console]::Error.WriteLine('WAM returned no token.'); exit 1 }
    [Console]::Out.WriteLine($shared.Token)
    exit 0
}
catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
}
'@

    $tempScript = $null
    $proc = $null
    try {
        $tempScript = [System.IO.Path]::GetTempFileName() -replace '\.tmp$', '.ps1'
        [System.IO.File]::WriteAllText($tempScript, $ps5Script, [System.Text.Encoding]::UTF8)

        $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.FileName = $ps5Path
        $startInfo.Arguments = '-NoProfile -STA -ExecutionPolicy Bypass -File "{0}"' -f $tempScript
        $startInfo.UseShellExecute = $false
        $startInfo.RedirectStandardOutput = $true
        $startInfo.RedirectStandardError = $true
        $startInfo.CreateNoWindow = $true
        # Pass WAM parameters via environment variables — avoids quoting/parsing issues for
        # values like scopes that contain spaces when passed as command-line arguments.
        $startInfo.EnvironmentVariables['WAM_PS5_AUTHORITY'] = $Authority
        $startInfo.EnvironmentVariables['WAM_PS5_CLIENTID'] = $ClientId
        $startInfo.EnvironmentVariables['WAM_PS5_SCOPE'] = $Scope

        $proc = [System.Diagnostics.Process]::Start($startInfo)
        $stdoutTask = $proc.StandardOutput.ReadToEndAsync()
        $stderrTask = $proc.StandardError.ReadToEndAsync()

        if (-not $proc.WaitForExit($TimeoutSeconds * 1000)) {
            try { $proc.Kill() } catch {}
            Write-Verbose ('WAM subprocess timed out after {0}s.' -f $TimeoutSeconds)
            return $null
        }

        $stdout = $stdoutTask.GetAwaiter().GetResult().Trim()
        $stderr = $stderrTask.GetAwaiter().GetResult().Trim()

        if ($proc.ExitCode -ne 0) {
            if ($stderr -match '(?i)UserCancel') { throw 'MSA authentication cancelled.' }
            Write-Verbose ('WAM subprocess failed (exit {0}): {1}' -f $proc.ExitCode, $stderr)
            return $null
        }

        if ([string]::IsNullOrWhiteSpace($stdout)) {
            Write-Verbose 'WAM subprocess succeeded but returned an empty token.'
            return $null
        }

        return $stdout
    }
    catch {
        if ($_.Exception.Message -match '(?i)cancel') { throw }
        Write-Verbose ('WAM subprocess error: {0}' -f $_.Exception.Message)
        return $null
    }
    finally {
        if ($proc -and -not $proc.HasExited) { try { $proc.Kill() } catch {} }
        if ($proc) { $proc.Dispose() }
        if ($tempScript -and (Test-Path -LiteralPath $tempScript)) {
            Remove-Item -LiteralPath $tempScript -Force -ErrorAction SilentlyContinue
        }
    }
}

function Find-MSALAssemblies {
    # Returns a pscustomobject { MsalPath; BrokerPath; HasBroker; Version } for the best available
    # Microsoft.Identity.Client.dll, preferring copies that include the WAM broker DLL alongside them.
    # Returns $null when no copy is found.
    $candidates = [System.Collections.Generic.List[pscustomobject]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    $searchRoots = @($env:PSModulePath -split [System.IO.Path]::PathSeparator) + @(
        (Join-Path $env:ProgramFiles 'WindowsPowerShell\Modules'),
        (Join-Path $env:ProgramFiles 'PowerShell\Modules'),
        (Join-Path ([System.Environment]::GetFolderPath('MyDocuments')) 'WindowsPowerShell\Modules'),
        (Join-Path ([System.Environment]::GetFolderPath('MyDocuments')) 'PowerShell\Modules')
    ) | Where-Object { $_ -and (Test-Path -LiteralPath $_) -and $seen.Add($_) }

    foreach ($root in $searchRoots) {
        Get-ChildItem -Path $root -Filter 'Microsoft.Identity.Client.dll' -Recurse -Depth 6 -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -notmatch '(?i)\\(runtimes|ref|native|resources)\\' } |
        ForEach-Object {
            try {
                if ($PSVersionTable.PSVersion.Major -le 5 -and $_.FullName -match '(?i)\\net(?:[5-9]|[1-9][0-9]+)\.[0-9]+[^\\]*\\') { return }
                $ver = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($_.FullName).FileVersion -as [version]
                if (-not $ver) { return }
                $brokerPath = Join-Path $_.DirectoryName 'Microsoft.Identity.Client.Broker.dll'
                $hasBroker = Test-Path -LiteralPath $brokerPath
                $candidates.Add([pscustomobject]@{
                        MsalPath      = $_.FullName
                        BrokerPath    = if ($hasBroker) { $brokerPath } else { $null }
                        HasBroker     = $hasBroker
                        IsNetStandard = $_.FullName -match '(?i)\\netstandard2\.0\\'
                        IsFramework   = $_.FullName -match '(?i)\\net4[^\\]*\\'
                        Version       = $ver
                    })
            }
            catch {}
        }
    }

    if ($candidates.Count -eq 0) { return $null }
    return $candidates |
    Sort-Object @{ e = 'HasBroker'; Descending = $true }, @{ e = { $PSVersionTable.PSVersion.Major -le 5 -and $_.IsFramework }; Descending = $true }, @{ e = 'IsNetStandard'; Descending = $true }, @{ e = 'Version'; Descending = $true } |
    Select-Object -First 1
}

function Get-WAMAccessTokenViaMSAL {
    param(
        [parameter(Mandatory = $true)][string]$Authority,
        [parameter(Mandatory = $true)][string]$ClientId,
        [parameter(Mandatory = $true)][string]$Scope,
        [int]$TimeoutSeconds = 120
    )

    $assemblies = Find-MSALAssemblies
    if (-not $assemblies) {
        Write-Verbose 'MSAL WAM: Microsoft.Identity.Client.dll not found in any module path.'
        return $null
    }
    if (-not $assemblies.HasBroker) {
        Write-Verbose ('MSAL WAM: MSAL {0} found but no broker DLL alongside it; WAM requires the broker.' -f $assemblies.Version)
        return $null
    }

    Write-Verbose ('MSAL WAM: MSAL {0} + broker at {1}' -f $assemblies.Version, [System.IO.Path]::GetDirectoryName($assemblies.MsalPath))

    # Pre-load the DLLs so they are in the AppDomain before the compiled helper references them.
    # PS7's New-Object / [Type] lookup fails for types in dynamically loaded assemblies; compiling
    # a thin C# wrapper via Add-Type lets Roslyn resolve extension methods and nested types properly.
    foreach ($dllPath in @($assemblies.MsalPath, $assemblies.BrokerPath)) {
        try { Add-Type -Path $dllPath -ErrorAction Stop } catch {
            if ($_.Exception.Message -notmatch '(?i)already') {
                Write-Verbose ('MSAL WAM: failed to load {0}: {1}' -f ([System.IO.Path]::GetFileName($dllPath)), $_.Exception.Message)
                return $null
            }
        }
    }

    # Compile once per session; reuse on subsequent calls.
    $helperTypeName = 'GetEventSession.Internal.MSALWAMHelper'
    $helperType = $helperTypeName -as [type]
    if (-not $helperType) {
        $csSrc = @'
using System;
using System.Runtime.InteropServices;
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Broker;

namespace GetEventSession.Internal {
    public static class MSALWAMHelper {
        [DllImport("kernel32.dll")]
        private static extern IntPtr GetConsoleWindow();
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hwnd);

        public static string RequestToken(string authority, string clientId, string[] scopes, int timeoutMs) {
            var brokerOpts = new BrokerOptions(BrokerOptions.OperatingSystems.Windows);
            var app = PublicClientApplicationBuilder.Create(clientId)
                .WithAuthority(authority)
                .WithBroker(brokerOpts)
                .Build();

            // Try silent — no UI if the account is already cached by WAM.
            var accounts = app.GetAccountsAsync().GetAwaiter().GetResult();
            foreach (var acct in accounts) {
                try {
                    var sr = app.AcquireTokenSilent(scopes, acct).ExecuteAsync().GetAwaiter().GetResult();
                    if (sr != null && !string.IsNullOrWhiteSpace(sr.AccessToken)) return sr.AccessToken;
                } catch {}
            }

            // WAM broker requires a non-zero HWND; use the console window as parent.
            var hwnd = GetConsoleWindow();
            if (hwnd == IntPtr.Zero || !IsWindowVisible(hwnd)) hwnd = GetForegroundWindow();
            if (hwnd == IntPtr.Zero) throw new InvalidOperationException("No window is available to parent the WAM dialog.");
            var ir = app.AcquireTokenInteractive(scopes)
                .WithParentActivityOrWindow(hwnd)
                .ExecuteAsync().GetAwaiter().GetResult();
            return ir == null ? null : ir.AccessToken;
        }
    }
}
'@
        $runtimeDir = Split-Path ([System.Object].Assembly.Location)
        $msalRefs = [System.Collections.Generic.List[string]]::new()
        $msalRefs.Add($assemblies.MsalPath)
        $msalRefs.Add($assemblies.BrokerPath)
        foreach ($stdDll in @('netstandard.dll')) {
            $p = Join-Path $runtimeDir $stdDll
            if (-not (Test-Path -LiteralPath $p)) { $p = Join-Path (Join-Path $runtimeDir 'Facades') $stdDll }
            if (Test-Path -LiteralPath $p) { $msalRefs.Add($p) }
        }

        try {
            Add-Type -TypeDefinition $csSrc -ReferencedAssemblies $msalRefs.ToArray() -ErrorAction Stop
            $helperType = $helperTypeName -as [type]
        }
        catch {
            Write-Verbose ('MSAL WAM: C# wrapper compilation failed: {0}' -f $_.Exception.Message)
            return $null
        }
    }

    if (-not $helperType) {
        Write-Verbose 'MSAL WAM: helper type unavailable after compilation.'
        return $null
    }

    $scopes = [string[]]@($Scope -split '\s+' | Where-Object { $_ })
    try {
        $token = $helperType::RequestToken($Authority, $ClientId, $scopes, $TimeoutSeconds * 1000)
        return $token
    }
    catch {
        if ($_.Exception.Message -match '(?i)(cancel|AADSTS65004)') { throw 'MSA authentication cancelled.' }
        Write-Verbose ('MSAL WAM token request failed: {0}' -f $_.Exception.Message)
        throw ('MSAL WAM token request failed: {0}' -f $_.Exception.Message)
    }
}

function Get-OAuthTokenViaPKCE {
    param(
        [parameter(Mandatory = $true)][string]$ClientId,
        [string]$TenantId = 'consumers',
        [string]$Scope = 'openid profile email offline_access',
        [int]$TimeoutSeconds = 120
    )

    if (-not [System.Net.HttpListener]::IsSupported) {
        throw 'HttpListener not supported on this platform.'
    }

    # Build PKCE pair.
    $verifierBytes = New-Object byte[] 32
    ([System.Security.Cryptography.RandomNumberGenerator]::Create()).GetBytes($verifierBytes)
    $codeVerifier = [Convert]::ToBase64String($verifierBytes) -replace '\+', '-' -replace '/', '_' -replace '=', ''
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    $challengeBytes = $sha256.ComputeHash([System.Text.Encoding]::ASCII.GetBytes($codeVerifier))
    $codeChallenge = [Convert]::ToBase64String($challengeBytes) -replace '\+', '-' -replace '/', '_' -replace '=', ''

    # Find a free port.
    $port = $null
    $listener = $null
    foreach ($candidate in (Get-Random -Minimum 49152 -Maximum 65000 -Count 10)) {
        try {
            $l = New-Object System.Net.HttpListener
            $l.Prefixes.Add("http://localhost:$candidate/")
            $l.Start()
            $port = $candidate
            $listener = $l
            break
        }
        catch { if ($l) { try { $l.Stop() } catch {} } }
    }
    if (-not $listener) { throw 'Could not bind a local HTTP listener port for OAuth redirect.' }

    $redirectUri = "http://localhost:$port/"
    $state = [Guid]::NewGuid().ToString('N')
    $effectiveScope = if ([string]::IsNullOrWhiteSpace($Scope)) { 'openid profile email offline_access' } else { $Scope }

    $authUrl = ('https://login.microsoftonline.com/{0}/oauth2/v2.0/authorize?client_id={1}' +
        '&response_type=code&redirect_uri={2}&scope={3}' +
        '&code_challenge={4}&code_challenge_method=S256&state={5}' +
        '&prompt=select_account') -f
    $TenantId, $ClientId,
    [Uri]::EscapeDataString($redirectUri),
    [Uri]::EscapeDataString($effectiveScope),
    $codeChallenge, $state

    try {
        Write-Host ('Opening browser for OAuth sign-in ({0})...' -f $TenantId)
        Start-Process $authUrl

        $contextTask = $listener.GetContextAsync()
        $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
        while (-not $contextTask.IsCompleted -and (Get-Date) -lt $deadline) {
            [System.Threading.Thread]::Sleep(200)
        }
        if (-not $contextTask.IsCompleted) {
            throw ('OAuth browser sign-in timed out after {0}s.' -f $TimeoutSeconds)
        }

        $ctx = $contextTask.GetAwaiter().GetResult()
        $query = $ctx.Request.Url.Query.TrimStart('?')

        # Send a close-the-tab page back so the browser window looks finished.
        $doneHtml = '<html><head><title>Sign-in complete</title></head><body>' +
        '<h3>Sign-in complete — you can close this tab.</h3></body></html>'
        $doneBytes = [System.Text.Encoding]::UTF8.GetBytes($doneHtml)
        $ctx.Response.ContentType = 'text/html; charset=utf-8'
        $ctx.Response.OutputStream.Write($doneBytes, 0, $doneBytes.Length)
        $ctx.Response.Close()

        $params = @{}
        foreach ($pair in ($query -split '&')) {
            $kv = $pair -split '=', 2
            if ($kv.Count -eq 2) { $params[$kv[0]] = [Uri]::UnescapeDataString($kv[1]) }
        }

        if ($params['error']) {
            throw ('OAuth error from AAD: {0} — {1}' -f $params['error'], $params['error_description'])
        }
        if ($params['state'] -ne $state) {
            throw 'OAuth state mismatch — possible CSRF.'
        }
        $code = $params['code']
        if ([string]::IsNullOrWhiteSpace($code)) { throw 'No authorization code in OAuth callback.' }

        $tokenBody = @{
            client_id     = $ClientId
            grant_type    = 'authorization_code'
            code          = $code
            redirect_uri  = $redirectUri
            code_verifier = $codeVerifier
        }
        $tokenParams = @{
            Uri         = 'https://login.microsoftonline.com/{0}/oauth2/v2.0/token' -f $TenantId
            Method      = 'Post'
            Body        = $tokenBody
            ErrorAction = 'Stop'
        }
        $tokenResponse = Invoke-RestMethod @tokenParams
        if ([string]::IsNullOrWhiteSpace($tokenResponse.access_token)) {
            throw 'OAuth token endpoint returned no access_token.'
        }
        Write-Verbose ('OAuth PKCE sign-in succeeded (tenant={0}).' -f $TenantId)
        return $tokenResponse.access_token
    }
    finally {
        try { $listener.Stop(); $listener.Close() } catch {}
    }
}

function Get-OAuthTokenViaDeviceCode {
    param(
        [parameter(Mandatory = $true)][string]$ClientId,
        [string]$TenantId = 'consumers',
        [string]$Scope = 'openid profile email offline_access',
        [int]$TimeoutSeconds = 300,
        [uri]$Proxy
    )

    $baseUrl = ('https://login.microsoftonline.com/{0}/oauth2/v2.0' -f $TenantId)
    $effectiveScope = if ([string]::IsNullOrWhiteSpace($Scope)) { 'openid profile email offline_access' } else { $Scope }

    $dcParams = @{
        Uri         = "$baseUrl/devicecode"
        Method      = 'POST'
        Body        = @{ client_id = $ClientId; scope = $effectiveScope }
        ErrorAction = 'Stop'
        Verbose     = $false
    }
    if ($Proxy) { $dcParams.Proxy = $Proxy }

    $dcResp = Invoke-RestMethod @dcParams

    # AAD returns a ready-made user message; display it and open the verification URL.
    Write-Host ''
    Write-Host $dcResp.message
    Write-Host ''
    try { Start-Process $dcResp.verification_uri } catch {}

    $interval = [Math]::Max(5, [int]$dcResp.interval)
    $expiresIn = [Math]::Max(60, [int]$dcResp.expires_in)
    $deadline = (Get-Date).AddSeconds([Math]::Min($TimeoutSeconds, $expiresIn))

    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds $interval
        $tokenParams = @{
            Uri         = "$baseUrl/token"
            Method      = 'POST'
            Body        = @{
                client_id   = $ClientId
                grant_type  = 'urn:ietf:params:oauth:grant-type:device_code'
                device_code = $dcResp.device_code
            }
            ErrorAction = 'Stop'
            Verbose     = $false
        }
        if ($Proxy) { $tokenParams.Proxy = $Proxy }

        try {
            $tok = Invoke-RestMethod @tokenParams
            if (-not [string]::IsNullOrWhiteSpace($tok.access_token)) { return $tok.access_token }
        }
        catch {
            $errJson = $null
            try { $errJson = ($_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction Stop) } catch {}
            if ($errJson) {
                switch ($errJson.error) {
                    'authorization_pending' { }
                    'slow_down' { $interval += 5 }
                    'authorization_declined' { throw 'MSA device code authentication was declined.' }
                    'expired_token' { throw 'MSA device code has expired.' }
                    default { throw ('Device code token error: {0} — {1}' -f $errJson.error, $errJson.error_description) }
                }
            }
            else { throw }
        }
    }
    throw ('MSA device code sign-in timed out after {0}s.' -f $TimeoutSeconds)
}

function Get-WAMAccessToken {
    param(
        [parameter(Mandatory = $true)][string]$ClientId,
        [parameter(Mandatory = $true)][string]$TenantId,
        [string]$Scope
    )

    $authority = ('https://login.microsoft.com/{0}' -f $TenantId)
    $effectiveScope = if ([string]::IsNullOrWhiteSpace($Scope)) { 'openid profile email offline_access' } else { $Scope }

    # ── Try native PS WinRT type loading (PS5.1, and PS7 where the WinRT projection is present) ──
    $winRTLoadError = $null
    try {
        $null = [Windows.Security.Authentication.Web.Core.WebAuthenticationCoreManager, Windows.Security.Authentication.Web.Core, ContentType = WindowsRuntime]
        $null = [Windows.Security.Authentication.Web.Core.WebTokenRequest, Windows.Security.Authentication.Web.Core, ContentType = WindowsRuntime]
        $null = [Windows.Security.Authentication.Web.Core.WebTokenRequestStatus, Windows.Security.Authentication.Web.Core, ContentType = WindowsRuntime]
    }
    catch {
        $winRTLoadError = $_
    }

    if ($winRTLoadError) {
        # ContentType=WindowsRuntime is unavailable in PS7/.NET 6+.
        # Path 2: MSAL.NET + WAM broker DLL — ships with Az, Graph, ExchangeOnline PS modules.
        Write-Verbose 'WinRT unavailable in this PS version; trying MSAL.NET WAM broker.'
        $msalToken = Get-WAMAccessTokenViaMSAL -Authority $authority -ClientId $ClientId -Scope $effectiveScope
        if (-not [string]::IsNullOrWhiteSpace($msalToken)) {
            return $msalToken
        }
        throw 'WAM broker did not return an access token. Check verbose output for MSAL discovery or compilation failures.'
    }

    # ── Native PS WinRT path (PS5.1) ─────────────────────────────────────────
    if (-not ('WinAPI.ConsoleWindow' -as [type])) {
        $addTypeParams = @{
            Namespace        = 'WinAPI'
            Name             = 'ConsoleWindow'
            MemberDefinition = '[DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();'
            ErrorAction      = 'SilentlyContinue'
        }
        Add-Type @addTypeParams
    }

    $shared = [hashtable]::Synchronized(@{ Token = $null; Error = $null })

    $staThread = [System.Threading.Thread]::new([System.Threading.ThreadStart] {
            try {
                $hwnd = [IntPtr]::Zero
                if ('WinAPI.ConsoleWindow' -as [type]) {
                    $hwnd = [WinAPI.ConsoleWindow]::GetConsoleWindow()
                }

                $findTask = [Windows.Security.Authentication.Web.Core.WebAuthenticationCoreManager]::FindAccountProviderAsync($authority)
                $provider = $findTask.AsTask().GetAwaiter().GetResult()

                if (-not $provider) {
                    $shared.Error = ('WAM: no account provider found for {0}.' -f $authority)
                    return
                }

                $tokenReq = [Windows.Security.Authentication.Web.Core.WebTokenRequest]::new($provider, $effectiveScope, $ClientId)
                $requestTask = [Windows.Security.Authentication.Web.Core.WebAuthenticationCoreManager]::RequestTokenForWindowAsync($hwnd, $tokenReq)
                $tokenResult = $requestTask.AsTask().GetAwaiter().GetResult()

                $status = $tokenResult.ResponseStatus
                if ($status -eq [Windows.Security.Authentication.Web.Core.WebTokenRequestStatus]::Success) {
                    $shared.Token = [string]$tokenResult.ResponseData[0].Token
                }
                elseif ($status -eq [Windows.Security.Authentication.Web.Core.WebTokenRequestStatus]::UserCancel) {
                    $shared.Error = 'UserCancel'
                }
                else {
                    $errCode = if ($tokenResult.ResponseError) { $tokenResult.ResponseError.ErrorCode } else { '?' }
                    $errMsg = if ($tokenResult.ResponseError) { [string]$tokenResult.ResponseError.ErrorMessage } else { '' }
                    $shared.Error = ('WAM request failed ({0}): code={1} {2}' -f $status, $errCode, $errMsg)
                }
            }
            catch {
                $shared.Error = $_.Exception.Message
            }
        })

    $staThread.SetApartmentState([System.Threading.ApartmentState]::STA)
    $staThread.IsBackground = $true
    $staThread.Start()
    $staThread.Join(60000) | Out-Null

    if ($shared.Error -eq 'UserCancel') { throw 'MSA authentication cancelled.' }
    if (-not [string]::IsNullOrWhiteSpace($shared.Error)) { throw $shared.Error }
    if ([string]::IsNullOrWhiteSpace($shared.Token)) {
        if ($staThread.IsAlive) {
            Write-Warning 'WAM sign-in dialog did not complete within 60 seconds; falling back to Edge CDP.'
        }
        throw 'WAM returned no token.'
    }

    return $shared.Token
}

function Invoke-EdgeCDPAuthSession {
    param(
        [parameter(Mandatory = $true)][uri]$CookieUri,
        [parameter(Mandatory = $true)][uri]$PortalRootUri,
        [uri]$Proxy,
        [scriptblock]$OnAuthenticatedBrowser
    )

    $edgePath = Get-EdgeExecutablePath
    if (-not $edgePath) {
        throw 'Microsoft Edge not found on this machine.'
    }

    # Pick a random unprivileged debug port and guarantee uniqueness for this PID.
    $debugPort = Get-Random -Minimum 19200 -Maximum 19999
    $tempProfile = Join-Path $env:TEMP ('EdgeAuth_{0}_{1}' -f $PID, [System.IO.Path]::GetRandomFileName().Replace('.', ''))

    $edgeArgs = @(
        ('--remote-debugging-port={0}' -f $debugPort),
        '--remote-debugging-address=127.0.0.1',
        ('--user-data-dir={0}' -f $tempProfile),
        '--no-first-run',
        '--no-default-browser-check',
        '--disable-sync',
        '--disable-extensions',
        '--disable-background-mode',
        # Prevent Edge from using Windows SSO to silently authenticate and then close the process
        # before the user has a chance to click Continue and before CDP cookies can be collected.
        '--disable-features=msSingleSignOn',
        '--start-maximized',
        $PortalRootUri.AbsoluteUri
    )

    $edgeProcess = $null
    try {
        $edgeProcess = Start-Process -FilePath $edgePath -ArgumentList $edgeArgs -PassThru -ErrorAction Stop

        # Show the instruction dialog immediately; Edge opens in parallel.
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        $form = New-Object System.Windows.Forms.Form
        $form.Text = 'Microsoft Account Sign-In'
        $form.Width = 560
        $form.Height = 230
        $form.StartPosition = 'CenterScreen'
        $form.ShowInTaskbar = $true
        $form.TopMost = $true
        $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $form.MaximizeBox = $false
        $form.MinimizeBox = $false

        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Left = 16; $lbl.Top = 16; $lbl.Width = 516; $lbl.Height = 80
        $lbl.Text = ('Sign in at {0} in the Edge window that opened, then click Continue here.{1}IMPORTANT: Keep Edge open until you have clicked Continue — closing Edge early prevents cookie collection.' -f $CookieUri.Host, [Environment]::NewLine)
        if ($OnAuthenticatedBrowser) {
            $form.Text = 'Event Portal'
            $lbl.Text = ('Open the event sessions page at {0}, signing in only if required, then click Continue here. Keep Edge open until the catalog request has been captured.' -f $CookieUri.Host)
        }

        $statusLbl = New-Object System.Windows.Forms.Label
        $statusLbl.Left = 16; $statusLbl.Top = 104; $statusLbl.Width = 516; $statusLbl.Height = 20
        $statusLbl.ForeColor = [System.Drawing.SystemColors]::GrayText

        $btnOK = New-Object System.Windows.Forms.Button
        $btnOK.Text = 'Continue'; $btnOK.Width = 100; $btnOK.Height = 28
        $btnOK.Left = 336; $btnOK.Top = 138

        $btnCxl = New-Object System.Windows.Forms.Button
        $btnCxl.Text = 'Cancel'; $btnCxl.Width = 100; $btnCxl.Height = 28
        $btnCxl.Left = 444; $btnCxl.Top = 138

        # Validate process liveness and show when Edge's CDP endpoint is ready.
        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 2000
        $timer.Add_Tick({
                if ($edgeProcess.HasExited) {
                    $timer.Stop()
                    $statusLbl.ForeColor = [System.Drawing.Color]::DarkOrange
                    $statusLbl.Text = 'Edge opened. When sign-in completed, click Continue — otherwise Cancel. Do not forget to click Accept cookies if prompted in Edge.'
                    return
                }
                if ($statusLbl.Tag -ne 'ready') {
                    try {
                        $null = Invoke-RestMethod -Uri ('http://127.0.0.1:{0}/json/version' -f $debugPort) -ErrorAction Stop -Verbose:$false
                        $statusLbl.Text = 'Edge is ready. Sign in, then click Continue.'
                        $statusLbl.ForeColor = [System.Drawing.Color]::DarkGreen
                        $statusLbl.Tag = 'ready'
                    }
                    catch {}
                }
            })

        $btnOK.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.Close() })
        $btnCxl.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $form.Close() })

        $form.Controls.AddRange(@($lbl, $statusLbl, $btnOK, $btnCxl))
        $form.Add_Shown({ $timer.Start(); $form.Activate(); $form.BringToFront() })

        Write-Host ('Edge sign-in window opened for {0}. Complete sign-in, then click Continue.' -f $CookieUri.Host)

        $allCookies = @()
        $targetCookies = @()
        $cdpVersionUrl = ('http://127.0.0.1:{0}/json/version' -f $debugPort)
        $maxCDPAttempts = 3

        for ($cdpAttempt = 1; $cdpAttempt -le $maxCDPAttempts; $cdpAttempt++) {
            if ($cdpAttempt -gt 1) {
                $statusLbl.Text = ('No sign-in detected — sign in to Edge and click Continue ({0}/{1}).' -f $cdpAttempt, $maxCDPAttempts)
                if ($OnAuthenticatedBrowser) {
                    $statusLbl.Text = ('Catalog request not ready. Reload the sessions page, then Continue ({0}/{1}).' -f $cdpAttempt, $maxCDPAttempts)
                }
                $statusLbl.ForeColor = [System.Drawing.Color]::DarkOrange
                $statusLbl.Tag = $null
                $form.DialogResult = [System.Windows.Forms.DialogResult]::None
                if (-not $edgeProcess.HasExited) { $timer.Start() }
            }

            $dialogResult = $form.ShowDialog()
            $timer.Stop()

            if ($dialogResult -eq [System.Windows.Forms.DialogResult]::Cancel) {
                throw 'MSA authentication cancelled.'
            }

            if ($edgeProcess.HasExited) {
                $exitCode = try { $edgeProcess.ExitCode } catch { '?' }

                # Clean exit (code 0): Edge closed after auth redirect (Windows SSO auto-close or
                # user closed the window). The session cookies are still on disk in the temp profile.
                # Relaunch Edge on the same profile with about:blank so CDP can read them back.
                if ($exitCode -eq 0 -and (Test-Path -LiteralPath $tempProfile)) {
                    Write-Verbose 'Edge exited cleanly; relaunching on same profile to recover cookies via CDP...'
                    $recoveryProcess = $null
                    try {
                        $recoveryArgs = @(
                            ('--remote-debugging-port={0}' -f $debugPort),
                            '--remote-debugging-address=127.0.0.1',
                            ('--user-data-dir={0}' -f $tempProfile),
                            '--no-first-run', '--no-default-browser-check',
                            '--disable-sync', '--disable-extensions',
                            '--disable-background-mode',
                            '--disable-features=msSingleSignOn',
                            'about:blank'
                        )
                        $recoveryProcess = Start-Process -FilePath $edgePath -ArgumentList $recoveryArgs -PassThru -ErrorAction Stop
                        $recoveryCdp = $null
                        $recoveryDeadline = (Get-Date).AddSeconds(12)
                        while ((Get-Date) -lt $recoveryDeadline -and -not $recoveryProcess.HasExited) {
                            try {
                                $recoveryCdp = Invoke-RestMethod -Uri $cdpVersionUrl -Method Get -ErrorAction Stop -Verbose:$false
                                break
                            }
                            catch { [System.Threading.Thread]::Sleep(500) }
                        }
                        if ($recoveryCdp -and -not [string]::IsNullOrWhiteSpace([string]$recoveryCdp.webSocketDebuggerUrl)) {
                            $recoveryCdpCookies = Invoke-EdgeCDPCommand -WebSocketUrl ([string]$recoveryCdp.webSocketDebuggerUrl) -Method 'Storage.getCookies'
                            if ($recoveryCdpCookies -and $recoveryCdpCookies.result -and $recoveryCdpCookies.result.cookies) {
                                $recoveredAll = @($recoveryCdpCookies.result.cookies)
                                $recoveredTarget = @($recoveredAll | Where-Object {
                                        $d = ([string]$_.domain).TrimStart('.').ToLowerInvariant()
                                        $d -ieq $CookieUri.Host -or $CookieUri.Host.EndsWith(('.{0}' -f $d), [System.StringComparison]::OrdinalIgnoreCase)
                                    })
                                if ($recoveredTarget.Count -gt 0) {
                                    Write-Verbose ('CDP recovery: collected {0} host cookie(s) from relaunched Edge profile.' -f $recoveredTarget.Count)
                                    try { Invoke-EdgeCDPCommand -WebSocketUrl ([string]$recoveryCdp.webSocketDebuggerUrl) -Method 'Browser.close' -TimeoutSeconds 3 | Out-Null } catch {}
                                    $recoveryProcess.WaitForExit(2000) | Out-Null
                                    return New-MSAWebSessionFromCDPCookies -CookieUri $CookieUri -Cookies $recoveredAll -Proxy $Proxy
                                }
                            }
                        }
                    }
                    catch {
                        Write-Verbose ('Edge cookie recovery relaunch failed: {0}' -f $_.Exception.Message)
                    }
                    finally {
                        if ($recoveryProcess -and -not $recoveryProcess.HasExited) { try { $recoveryProcess.Kill() } catch {} }
                        if ($recoveryProcess) { $recoveryProcess.Dispose() }
                    }
                }

                # Last resort: system WinInet store (populated if Windows SSO completed auth there)
                $winInetHeader = Get-UriCookieHeader -Uri $CookieUri
                if (-not [string]::IsNullOrWhiteSpace($winInetHeader)) {
                    Write-Verbose ('Edge exited; recovered {0} cookie(s) from system cookie store.' -f ($winInetHeader -split ';').Count)
                    return New-WebSessionFromCookieHeader -CookieHeader $winInetHeader -CookieUri $CookieUri
                }

                throw ('Edge closed (exit code {0}) before sign-in cookies could be collected. Keep Edge open until after clicking Continue.' -f $exitCode)
            }

            # Wait for the debug endpoint to become available (Edge may still be finishing startup).
            $cdpVersion = $null
            $cdpDeadline = (Get-Date).AddSeconds(15)
            while ((Get-Date) -lt $cdpDeadline -and -not $edgeProcess.HasExited) {
                try {
                    $cdpVersion = Invoke-RestMethod -Uri $cdpVersionUrl -Method Get -ErrorAction Stop -Verbose:$false
                    break
                }
                catch { [System.Threading.Thread]::Sleep(500) }
            }

            if (-not $cdpVersion -or [string]::IsNullOrWhiteSpace([string]$cdpVersion.webSocketDebuggerUrl)) {
                throw ('CDP endpoint on port {0} did not respond. Edge may not have started with remote debugging.' -f $debugPort)
            }

            $browserWsUrl = [string]$cdpVersion.webSocketDebuggerUrl
            Write-Verbose ('CDP: collecting cookies via browser endpoint {0}' -f $browserWsUrl)

            # Storage.getCookies is the browser-level command; Network.getAllCookies only works on page-level endpoints.
            $cdpCookies = Invoke-EdgeCDPCommand -WebSocketUrl $browserWsUrl -Method 'Storage.getCookies'
            if (-not $cdpCookies -or -not $cdpCookies.result -or -not $cdpCookies.result.cookies) {
                if ($cdpAttempt -lt $maxCDPAttempts) {
                    Write-Host ('No cookies returned by CDP (attempt {0}/{1}); please sign in and click Continue again.' -f $cdpAttempt, $maxCDPAttempts)
                    continue
                }
                throw 'CDP Storage.getCookies returned no cookies; sign-in may not have completed.'
            }
            $allCookies = @($cdpCookies.result.cookies)

            $targetCookies = @($allCookies | Where-Object {
                    $d = ([string]$_.domain).TrimStart('.').ToLowerInvariant()
                    $d -ieq $CookieUri.Host -or $CookieUri.Host.EndsWith(('.{0}' -f $d), [System.StringComparison]::OrdinalIgnoreCase)
                })
            Write-Verbose ('CDP: {0} total / {1} host-matched cookies for {2}.' -f $allCookies.Count, $targetCookies.Count, $CookieUri.Host)

            if ($targetCookies.Count -gt 0) {
                if ($OnAuthenticatedBrowser) {
                    try {
                        $browserSession = New-MSAWebSessionFromCDPCookies -CookieUri $CookieUri -Cookies $allCookies -Proxy $Proxy
                        $null = & $OnAuthenticatedBrowser $debugPort $browserSession
                        return $browserSession
                    }
                    catch {
                        if ($cdpAttempt -ge $maxCDPAttempts) { throw }
                        if ($_.Exception.Data['StatusCode']) {
                            Write-Warning ('Catalog request validation failed (HTTP {0}). Reload the sessions page, then click Continue.' -f $_.Exception.Data['StatusCode'])
                        }
                        else {
                            Write-Warning 'Catalog context is not ready. Open or reload the event sessions page in the same Edge window, then click Continue.'
                        }
                        continue
                    }
                }
                break
            }

            if ($cdpAttempt -lt $maxCDPAttempts) {
                Write-Host ('No cookies found for {0} (attempt {1}/{2}); please complete sign-in and click Continue again.' -f $CookieUri.Host, $cdpAttempt, $maxCDPAttempts)
            }
        }

        if ($targetCookies.Count -eq 0) {
            throw ('No cookies collected for {0} via CDP after {1} attempts. Was sign-in completed?' -f $CookieUri.Host, $maxCDPAttempts)
        }

        # Wait for the portal SPA to load and MSAL.js to store the access token in sessionStorage.
        # The redirect from the login page back to the portal may still be in progress when the
        # user clicks Continue, so we poll until the token appears or a 20-second window closes.
        $msalToken = $null
        $msalCutoff = (Get-Date).AddSeconds(20)
        Write-Verbose ('CDP: waiting for portal page and MSAL token on {0}...' -f $CookieUri.Host)
        while (-not $msalToken -and (Get-Date) -lt $msalCutoff -and -not $edgeProcess.HasExited) {
            $msalToken = Get-MSALAccessTokenFromCdpPage -DebugPort $debugPort -CookieUri $CookieUri
            if (-not $msalToken) { [System.Threading.Thread]::Sleep(1500) }
        }
        if ($msalToken) {
            $script:MSABearerToken = $msalToken
            Write-Verbose ('CDP: captured MSAL access token for {0}.' -f $CookieUri.Host)
        }
        else {
            Write-Verbose ('CDP: MSAL token not found for {0} after waiting; proceeding with cookies only.' -f $CookieUri.Host)
        }

        return New-MSAWebSessionFromCDPCookies -CookieUri $CookieUri -Cookies $allCookies -Proxy $Proxy
    }
    finally {
        # Graceful close first, hard kill if Edge is still running after 2 s.
        if ($timer) { $timer.Stop(); $timer.Dispose() }
        if ($form) { $form.Dispose() }
        if ($edgeProcess -and -not $edgeProcess.HasExited) {
            try {
                $browserWsUrl = (Invoke-RestMethod -Uri ('http://127.0.0.1:{0}/json/version' -f $debugPort) -ErrorAction SilentlyContinue -Verbose:$false).webSocketDebuggerUrl
                if ($browserWsUrl) {
                    Invoke-EdgeCDPCommand -WebSocketUrl $browserWsUrl -Method 'Browser.close' -TimeoutSeconds 3 | Out-Null
                }
            }
            catch {}
            $edgeProcess.WaitForExit(2000) | Out-Null
            if (-not $edgeProcess.HasExited) {
                try { $edgeProcess.Kill() } catch {}
            }
        }
        if ($edgeProcess) { $edgeProcess.Dispose() }
        if ($tempProfile -and (Test-Path -LiteralPath $tempProfile)) {
            try { Remove-Item -LiteralPath $tempProfile -Recurse -Force -ErrorAction SilentlyContinue } catch {}
        }
    }
}

function Get-MSAAuthWebSession {
    param(
        [parameter(Mandatory = $true)][uri]$StartUri,
        [parameter(Mandatory = $true)][uri]$CookieUri,
        [uri]$Proxy,
        [scriptblock]$OnAuthenticatedBrowser,
        [uri]$PortalUri
    )

    if (-not [Environment]::UserInteractive) {
        throw 'MSA sign-in UI requires an interactive desktop session. Run the script in an interactive PowerShell terminal.'
    }

    # Navigate to the portal root so the SPA login flow fires instead of returning raw JSON.
    $portalRootUri = [uri]('{0}://{1}/' -f $CookieUri.Scheme, $CookieUri.Host)
    if ($PortalUri -and $PortalUri.GetLeftPart([System.UriPartial]::Authority) -eq $CookieUri.GetLeftPart([System.UriPartial]::Authority)) {
        $portalRootUri = $PortalUri
    }

    # ── PATH 1: WAM (Windows Authentication Manager) ─────────────────────────
    # Auto-discovers the AAD client_id from the 302/401 challenge, then shows
    # the native Windows account picker — no browser window required.
    Write-Verbose ('Probing {0} for OAuth2 challenge parameters...' -f $StartUri.Host)
    $challengeUri = if ($PortalUri) { $portalRootUri } else { $StartUri }
    $challenge = Resolve-OAuthChallenge -TargetUri $challengeUri -Proxy $Proxy
    Write-Verbose ('Microsoft account requirement: {0}; authority tenant: {1}' -f $challenge.AccountType, $challenge.TenantId)

    if (-not $OnAuthenticatedBrowser -and -not [string]::IsNullOrWhiteSpace($challenge.ClientId)) {
        Write-Verbose ('OAuth2 challenge: client_id={0}  tenant={1}' -f $challenge.ClientId, $challenge.TenantId)
        try {
            Write-Host ('Requesting sign-in via Windows Authentication Manager for {0}...' -f $CookieUri.Host)
            $bearerToken = Get-WAMAccessToken -ClientId $challenge.ClientId -TenantId $challenge.TenantId -Scope $challenge.Scope
            if (-not [string]::IsNullOrWhiteSpace($bearerToken)) {
                $script:MSABearerToken = $bearerToken
                Write-Verbose 'WAM sign-in succeeded; Bearer token stored.'
                $wamSession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
                $wamSession.UserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36'
                return $wamSession
            }
        }
        catch {
            if ($_.Exception.Message -match '(?i)cancelled|UserCancel') { throw 'MSA authentication cancelled.' }
            Write-Warning ('WAM sign-in failed: {0}' -f $_.Exception.Message)
        }

        # ── PATH 1b: OAuth device code flow ──────────────────────────────────
        # WAM failed; device code flow needs no redirect URI — AAD returns a user_code
        # and polls until the user enters it at https://microsoft.com/devicelogin.
        # Falls through if the client does not have device-code flow enabled.
        if ([string]::IsNullOrWhiteSpace($script:MSABearerToken)) {
            try {
                $dcTenant = if ([string]::IsNullOrWhiteSpace($challenge.TenantId)) { 'common' } else { $challenge.TenantId }
                $dcScope = if ([string]::IsNullOrWhiteSpace($challenge.Scope)) { 'openid profile email offline_access' } else { $challenge.Scope }
                Write-Verbose ('Trying OAuth device code flow for {0} (tenant={1})...' -f $CookieUri.Host, $dcTenant)
                $dcToken = Get-OAuthTokenViaDeviceCode -ClientId $challenge.ClientId -TenantId $dcTenant -Scope $dcScope -Proxy $Proxy
                if (-not [string]::IsNullOrWhiteSpace($dcToken)) {
                    $script:MSABearerToken = $dcToken
                    Write-Verbose 'OAuth device code sign-in succeeded; Bearer token stored.'
                    $dcSession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
                    $dcSession.UserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36'
                    return $dcSession
                }
            }
            catch {
                if ($_.Exception.Message -match '(?i)cancelled|declined') { throw 'MSA authentication cancelled.' }
                Write-Verbose ('OAuth device code failed: {0}' -f $_.Exception.Message)
                Write-Host 'Falling back to Edge CDP sign-in...'
            }
        }
    }
    else {
        Write-Verbose 'Using portal-managed browser sign-in because browser request context is needed or no usable OAuth challenge was found.'
    }

    # ── PATH 2: Edge CDP ──────────────────────────────────────────────────────
    # Opens a dedicated Edge instance with remote debugging. User completes sign-in
    # in the browser; cookies are extracted via Chrome DevTools Protocol.
    $edgePath = Get-EdgeExecutablePath
    if ($edgePath) {
        return Invoke-EdgeCDPAuthSession -CookieUri $CookieUri -PortalRootUri $portalRootUri -Proxy $Proxy -OnAuthenticatedBrowser $OnAuthenticatedBrowser
    }

    if ($OnAuthenticatedBrowser) {
        throw 'Microsoft Edge is required to discover this event catalog request.'
    }

    # ── PATH 3: Embedded IE/Trident with paste-box fallback ───────────────────
    # Last resort for machines without Edge. Requires STA apartment; only works on sites that render in IE11.
    if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -eq [System.Threading.ApartmentState]::STA) {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        # Synchronized state shared between the form's event handlers and the outer scope.
        $formState = [hashtable]::Synchronized(@{ PastedCookies = $null })

        $form = New-Object System.Windows.Forms.Form
        $form.Text = 'Microsoft Account Sign-In'
        $form.Width = 1200
        $form.Height = 1020
        $form.StartPosition = 'CenterScreen'
        $form.ShowInTaskbar = $true
        $form.TopMost = $true

        $statusLabel = New-Object System.Windows.Forms.Label
        $statusLabel.Dock = 'Top'
        $statusLabel.Height = 30
        $statusLabel.Text = ('Sign in at {0} in the browser below, then click Continue.' -f $CookieUri.Host)

        # Bottom panel: cookie-paste fallback + Continue/Cancel buttons
        $bottomPanel = New-Object System.Windows.Forms.Panel
        $bottomPanel.Dock = 'Bottom'
        $bottomPanel.Height = 148

        $pasteLabel = New-Object System.Windows.Forms.Label
        $pasteLabel.Left = 8
        $pasteLabel.Top = 6
        $pasteLabel.Width = 1168
        $pasteLabel.Height = 36
        $pasteLabel.Text = (
            'If the embedded browser cannot sign in: open {0} in Firefox/Chrome/Edge, press F12, ' +
            'go to the Network tab, click any request to the site, copy the Cookie: header value, and paste it below.'
        ) -f $CookieUri.Host

        $pasteBox = New-Object System.Windows.Forms.TextBox
        $pasteBox.Left = 8
        $pasteBox.Top = 46
        $pasteBox.Width = 1168
        $pasteBox.Height = 54
        $pasteBox.Multiline = $true
        $pasteBox.ScrollBars = 'Horizontal'
        $pasteBox.WordWrap = $false
        $pasteBox.Font = New-Object System.Drawing.Font('Consolas', 8)

        $continueButton = New-Object System.Windows.Forms.Button
        $continueButton.Text = 'Continue'
        $continueButton.Width = 100
        $continueButton.Height = 28
        $continueButton.Left = 980
        $continueButton.Top = 112
        $continueButton.Enabled = $false

        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Text = 'Cancel'
        $cancelButton.Width = 100
        $cancelButton.Height = 28
        $cancelButton.Left = 1090
        $cancelButton.Top = 112

        $browser = New-Object System.Windows.Forms.WebBrowser
        $browser.Dock = 'Fill'
        $browser.ScriptErrorsSuppressed = $true
        $visitedUrls = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

        $continueButton.Add_Click({
                $formState.PastedCookies = $pasteBox.Text.Trim()
                $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $form.Close()
            })

        $cancelButton.Add_Click({
                $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                $form.Close()
            })

        # Enable Continue when the browser lands on the portal host.
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

        # Also enable Continue as soon as the user pastes anything into the cookie box.
        $pasteBox.add_TextChanged({
                $continueButton.Enabled = (-not [string]::IsNullOrWhiteSpace($pasteBox.Text))
            })

        $bottomPanel.Controls.AddRange(@($pasteLabel, $pasteBox, $continueButton, $cancelButton))
        $form.Controls.Add($browser)
        $form.Controls.Add($bottomPanel)
        $form.Controls.Add($statusLabel)

        $form.Add_Shown({
                $form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
                $form.Activate()
                $form.BringToFront()
            })

        Write-Host ('Launching Microsoft sign-in dialog. If it is not visible, use Alt+Tab to switch to "Microsoft Account Sign-In".')
        $browser.Navigate($portalRootUri.AbsoluteUri)
        $dialogResult = $form.ShowDialog()

        if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
            throw 'MSA authentication cancelled.'
        }

        # If the user pasted cookies, build and return a session from those (they take priority
        # over WinInet because the embedded IE engine cannot complete modern MSAL flows).
        $pastedCookieString = $formState.PastedCookies
        if (-not [string]::IsNullOrWhiteSpace($pastedCookieString)) {
            $pastedSession = New-WebSessionFromCookieHeader -CookieHeader $pastedCookieString -CookieUri $CookieUri
            if ($pastedSession) {
                Write-Verbose ('Using pasted cookies for {0} ({1} cookie(s)).' -f $CookieUri.Host, ($pastedSession.Cookies.GetCookies($CookieUri) | Measure-Object).Count)
                return $pastedSession
            }
        }

        # Fall back to WinInet cookies captured from the embedded browser session.
        return New-MSAWebSessionFromBrowserContext -CookieUri $CookieUri -StartUri $portalRootUri -Browser $browser -VisitedUrls $visitedUrls -Proxy $Proxy
    }

    throw 'MSA sign-in requires WAM (no AAD challenge detected), Microsoft Edge (not found), or an STA runspace (use powershell.exe -STA). None of these are available in this session.'
}

function Get-CustomEventBrowserRequestContext {
    param(
        [parameter(Mandatory = $true)][int]$DebugPort,
        [parameter(Mandatory = $true)][uri]$CatalogUri,
        [string]$SessionId
    )

    $endpointJson = ConvertTo-Json -InputObject @(Get-CustomCatalogEndpointPaths -Uri $CatalogUri) -Compress
    $expression = @'
new Promise(function(resolve){var endpoints=ENDPOINT,deadline=Date.now()+10000;function inspect(){var entries=performance.getEntriesByType('resource');for(var endpoint of endpoints){for(var index=entries.length-1;index>=0;index--){var resource=new URL(entries[index].name,location.href);if(resource.origin+resource.pathname===endpoint){resolve({url:resource.href,referer:location.origin+location.pathname});return;}}}if(Date.now()>=deadline){resolve(null);return;}setTimeout(inspect,250);}inspect();})
'@
    $expression = $expression.Replace('ENDPOINT', $endpointJson)
    $pages = @(Invoke-RestMethod -Uri ('http://127.0.0.1:{0}/json' -f $DebugPort) -Verbose:$false -Debug:$false -ErrorAction Stop)
    foreach ($page in $pages) {
        if ($page.type -ne 'page' -or -not $page.webSocketDebuggerUrl) { continue }
        $pageUri = $null
        if (-not [uri]::TryCreate([string]$page.url, [System.UriKind]::Absolute, [ref]$pageUri) -or $pageUri.Authority -ine $CatalogUri.Authority -or $pageUri.Scheme -ine $CatalogUri.Scheme) { continue }
        if ($SessionId) {
            $sessionIdJson = $SessionId | ConvertTo-Json -Compress
            $linkExpression = '(function(){var suffix="/session-details/"+SESSIONID;var link=Array.from(document.links).find(function(link){return new URL(link.href,location.href).pathname.endsWith(suffix);});return link?link.href:null;})()'.Replace('SESSIONID', $sessionIdJson)
            $linkResult = Invoke-EdgeCDPCommand -WebSocketUrl $page.webSocketDebuggerUrl -Method 'Runtime.evaluate' -Params @{ expression = $linkExpression; returnByValue = $true }
            $detailPageUri = $null
            if ([uri]::TryCreate([string]$linkResult.result.result.value, [System.UriKind]::Absolute, [ref]$detailPageUri) -and $detailPageUri.GetLeftPart([System.UriPartial]::Authority) -eq $CatalogUri.GetLeftPart([System.UriPartial]::Authority)) {
                Write-Verbose ('Opening session detail page for {0}' -f $SessionId)
                $null = Invoke-EdgeCDPCommand -WebSocketUrl $page.webSocketDebuggerUrl -Method 'Page.navigate' -Params @{ url = $detailPageUri.AbsoluteUri }
                [System.Threading.Tasks.Task]::Delay(1000).GetAwaiter().GetResult()
            }
        }
        Write-Verbose ('Inspecting catalog request context on {0}' -f $pageUri.GetLeftPart([System.UriPartial]::Path))
        $evaluation = Invoke-EdgeCDPCommand -WebSocketUrl $page.webSocketDebuggerUrl -Method 'Runtime.evaluate' -Params @{ expression = $expression; returnByValue = $true; awaitPromise = $true }
        $context = $evaluation.result.result.value
        if ($context -and $context.url -and $context.referer) {
            $requestUri = $null
            $refererUri = $null
            if ([uri]::TryCreate([string]$context.url, [System.UriKind]::Absolute, [ref]$requestUri) -and
                [uri]::TryCreate([string]$context.referer, [System.UriKind]::Absolute, [ref]$refererUri) -and
                (Test-CustomCatalogEndpoint -Uri $requestUri -CatalogUri $CatalogUri) -and
                (Test-IsCustomCatalogUri -Uri $refererUri) -and
                $refererUri.GetLeftPart([System.UriPartial]::Authority) -eq $CatalogUri.GetLeftPart([System.UriPartial]::Authority)) {
                Write-Verbose ('Captured event catalog request context for {0}; query values are not logged.' -f $CatalogUri.Host)
                return @{ Uri = $requestUri.AbsoluteUri; Referer = [string]$context.referer }
            }
        }
    }
    throw 'No event catalog request was found. Open or reload the event sessions page in the sign-in window and retry.'
}

function Test-IsCustomCatalogUri {
    param([uri]$Uri)

    return ($Uri -and $Uri.IsAbsoluteUri -and $Uri.Scheme -eq 'https' -and [string]::IsNullOrWhiteSpace($Uri.UserInfo))
}

function Get-CustomCatalogEndpointPaths {
    param([parameter(Mandatory = $true)][uri]$Uri)

    $origin = $Uri.GetLeftPart([System.UriPartial]::Authority)
    $path = $Uri.AbsolutePath
    $paths = @($path)
    if ($path -cmatch '^/api/schedule-builder/sessions(?:/search)?/?$') {
        $paths += '/api/schedule-builder/sessions', '/api/schedule-builder/sessions/', '/api/schedule-builder/sessions/search', '/api/schedule-builder/sessions/search/'
    }
    foreach ($candidate in @($paths | Select-Object -Unique)) { '{0}{1}' -f $origin, $candidate }
}

function Test-CustomCatalogEndpoint {
    param([uri]$Uri, [uri]$CatalogUri)

    if (-not (Test-IsCustomCatalogUri -Uri $Uri) -or -not (Test-IsCustomCatalogUri -Uri $CatalogUri)) { return $false }
    return @(Get-CustomCatalogEndpointPaths -Uri $CatalogUri) -ccontains $Uri.GetLeftPart([System.UriPartial]::Path)
}

function Get-CustomEventCacheKey {
    param([parameter(Mandatory = $true)][uri]$Uri)

    Add-Type -AssemblyName System.Web
    $query = [System.Web.HttpUtility]::ParseQueryString($Uri.Query)
    foreach ($name in @('pcId', 'page', 'access_token', 'id_token', 'at', 'uid', 'sig')) { $query.Remove($name) }
    $identity = '{0}?{1}|catalog-v2' -f $Uri.GetLeftPart([System.UriPartial]::Path), $query.ToString()
    $hash = [System.Security.Cryptography.SHA256]::Create()
    try {
        return ([BitConverter]::ToString($hash.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($identity))).Replace('-', '').ToLowerInvariant())
    }
    finally { $hash.Dispose() }
}

function Initialize-CustomCookieRequestContext {
    param(
        [parameter(Mandatory = $true)][Microsoft.PowerShell.Commands.WebRequestSession]$WebSession,
        [parameter(Mandatory = $true)][uri]$CatalogUri,
        [string]$SessionId,
        [uri]$Proxy
    )

    if ($WebSession.EventPortalUri -and (Update-CustomPortalRequestContext -WebSession $WebSession -CatalogUri $CatalogUri -SessionId $SessionId -Proxy $Proxy)) { return }
    $origin = $CatalogUri.GetLeftPart([System.UriPartial]::Authority)
    $parameters = @{ Uri = $origin + '/'; WebSession = $WebSession; MaximumRedirection = 5; TimeoutSec = 30; ErrorAction = 'Stop'; Verbose = $false; Debug = $false }
    if ($Proxy) { $parameters.Proxy = $Proxy }
    try { $page = Invoke-WebRequest @parameters }
    catch {
        if ($_.Exception -is [System.Management.Automation.PipelineStoppedException]) { throw }
        throw (New-CustomCookieFileException -Message 'Unable to inspect the Custom portal with exported cookies; credentials and response details withheld.')
    }
    if (Test-MSAAuthenticationRequired -Response $page -RequestUri $parameters.Uri) {
        throw (New-CustomCookieFileException -Message 'The Custom portal rejected the exported cookies. Sign in and export cookies again.')
    }
    $candidates = [System.Collections.Generic.List[uri]]::new()
    $pageUri = [uri]($origin + '/')
    if ($page.BaseResponse.ResponseUri) { $pageUri = $page.BaseResponse.ResponseUri }
    elseif ($page.BaseResponse.RequestMessage.RequestUri) { $pageUri = $page.BaseResponse.RequestMessage.RequestUri }
    if ($pageUri.GetLeftPart([System.UriPartial]::Authority) -eq $origin) {
        if ([string]$page.Content -match '\spc-id=["'']') { $candidates.Add($pageUri) }
        foreach ($link in @($page.Links)) {
            $candidate = $null
            if ([uri]::TryCreate($pageUri, [string]$link.href, [ref]$candidate) -and (Test-IsCustomCatalogUri -Uri $candidate) -and
                $candidate.GetLeftPart([System.UriPartial]::Authority) -eq $origin -and $candidate.AbsolutePath -match '/sessions/?$' -and
                $candidate.AbsolutePath -notmatch '^/api/') { $candidates.Add($candidate) }
        }
    }
    if ($CatalogUri.AbsolutePath -match '^/api/schedule-builder/sessions(?:/search)?/?$') {
        $candidates.Add([uri]($origin + '/sessions'))
    }
    foreach ($candidate in @($candidates | Select-Object -Unique | Select-Object -First 10)) {
        Set-ObjectPropertyValue -Object $WebSession -Name 'EventPortalUri' -Value $candidate.AbsoluteUri
        if (Update-CustomPortalRequestContext -WebSession $WebSession -CatalogUri $CatalogUri -SessionId $SessionId -Proxy $Proxy) { return }
    }
    throw (New-CustomCookieFileException -Message 'Exported cookies did not supply the required Custom portal page context. No browser discovery or interactive authentication was attempted.')
}

function Initialize-CustomRequestContext {
    param(
        [parameter(Mandatory = $true)][hashtable]$Context,
        [parameter(Mandatory = $true)][Microsoft.PowerShell.Commands.WebRequestSession]$WebSession,
        [uri]$Proxy
    )

    if ($Context.Uri -and $Context.Referer) { return }
    if ((-not $WebSession.EventRequestUri -or -not $WebSession.EventPortalUri) -and $script:CustomCookieFile) {
        Initialize-CustomCookieRequestContext -WebSession $WebSession -CatalogUri $Context.CatalogUri -Proxy $Proxy
    }
    if (-not $WebSession.EventRequestUri -or -not $WebSession.EventPortalUri) { throw 'Custom catalog context requires browser discovery.' }
    $requestUri = [uri]$WebSession.EventRequestUri
    if (-not (Test-CustomCatalogEndpoint -Uri $requestUri -CatalogUri $Context.CatalogUri)) { throw 'Custom request context has an unexpected endpoint.' }
    $Context.Uri = $requestUri.AbsoluteUri
    $Context.Referer = $WebSession.EventPortalUri
}

function Update-CustomPortalRequestContext {
    param(
        [parameter(Mandatory = $true)][Microsoft.PowerShell.Commands.WebRequestSession]$WebSession,
        [parameter(Mandatory = $true)][uri]$CatalogUri,
        [string]$SessionId,
        [uri]$Proxy
    )

    # Page-context IDs (pcId) are short-lived and rendered server-side into the signed-in sessions page.
    $portalUri = $null
    if (-not [uri]::TryCreate([string]$WebSession.EventPortalUri, [System.UriKind]::Absolute, [ref]$portalUri) -or
        $portalUri.GetLeftPart([System.UriPartial]::Authority) -ne $CatalogUri.GetLeftPart([System.UriPartial]::Authority)) { return $false }
    $parameters = @{ Uri = $portalUri.AbsoluteUri; WebSession = $WebSession; MaximumRedirection = 0; TimeoutSec = 30; UseBasicParsing = $true; ErrorAction = 'Stop'; Verbose = $false; Debug = $false }
    if ($Proxy) { $parameters.Proxy = $Proxy }
    try { $html = [string](Invoke-WebRequestNoRedirect -Parameters $parameters).Content }
    catch {
        if ($_.Exception -is [System.Management.Automation.PipelineStoppedException]) { throw }
        Write-Verbose 'Unable to refresh event page context from the sessions page.'
        return $false
    }
    $tag = [regex]::Match($html, '<[\w-]+\b[^>]*\spc-id=["''][^>]*>').Value
    if (-not $tag) { return $false }
    Add-Type -AssemblyName System.Web
    $attributes = @{}
    foreach ($match in [regex]::Matches($tag, '\s([\w-]*pc-id)=["'']([^"'']*)["'']')) {
        $attributes[$match.Groups[1].Value.ToLowerInvariant()] = [System.Web.HttpUtility]::HtmlDecode($match.Groups[2].Value)
    }
    if (-not $attributes['pc-id']) { return $false }
    $catalogPath = $CatalogUri.GetLeftPart([System.UriPartial]::Path)
    $priorUri = $null
    if ([uri]::TryCreate([string]$WebSession.EventRequestUri, [System.UriKind]::Absolute, [ref]$priorUri) -and
        (Test-CustomCatalogEndpoint -Uri $priorUri -CatalogUri $CatalogUri)) {
        $catalogPath = $priorUri.GetLeftPart([System.UriPartial]::Path)
    }
    Set-ObjectPropertyValue -Object $WebSession -Name 'EventRequestUri' -Value ('{0}?pcId={1}' -f $catalogPath, [uri]::EscapeDataString($attributes['pc-id']))
    if ($SessionId) {
        $detailPcId = if ($attributes['event-details-pc-id']) { $attributes['event-details-pc-id'] } else { $attributes['pc-id'] }
        $detailUri = [uri]::new($CatalogUri, ('/api/event/details/{0}?pcId={1}' -f $SessionId, [uri]::EscapeDataString($detailPcId)))
        Set-ObjectPropertyValue -Object $WebSession -Name 'EventDetailRequests' -Value @{ $SessionId = $detailUri.AbsoluteUri }
    }
    Write-Verbose ('Refreshed event page context for {0} from the signed-in sessions page; values are not logged.' -f $CatalogUri.Host)
    return $true
}

function Get-CustomCatalogFailureSummary {
    param(
        [parameter(Mandatory = $true)][System.Exception]$Exception,
        [int]$StatusCode = 0
    )

    $cause = 'unclassified request failure'
    $hint = 'Check EventUrl, proxy settings, network connectivity, and the PowerShell runtime.'
    $underlying = $Exception
    for ($inner = $Exception; $inner; $inner = $inner.InnerException) {
        $underlying = $inner
        if ($inner -is [System.TimeoutException] -or $inner -is [System.OperationCanceledException]) {
            $cause = 'request timeout'
            $hint = 'Check whether the catalog host or proxy is reachable and responding within the request timeout.'
            break
        }
        if ($inner -is [System.Security.Authentication.AuthenticationException]) {
            $cause = 'TLS authentication failure'
            $hint = 'Check certificate trust, system time, and HTTPS inspection by a proxy or firewall.'
            break
        }
        if ($inner -is [System.Net.Sockets.SocketException]) {
            $cause = 'DNS or socket connection failure'
            $hint = 'Check DNS resolution, host connectivity, firewall rules, and proxy settings.'
            break
        }
        if ($inner -is [System.Management.Automation.ParameterBindingException]) {
            $cause = 'request parameter binding failure'
            $hint = 'Check EventUrl and Proxy values, and run the script under PowerShell 7 or later.'
            break
        }
        if ($inner -is [System.Net.WebException]) {
            switch ($inner.Status) {
                'Timeout' { $cause = 'request timeout'; $hint = 'Check catalog host and proxy responsiveness.' }
                'NameResolutionFailure' { $cause = 'DNS resolution failure'; $hint = 'Check DNS resolution for the catalog host.' }
                'ProxyNameResolutionFailure' { $cause = 'proxy DNS resolution failure'; $hint = 'Check the configured proxy address and its DNS resolution.' }
                'TrustFailure' { $cause = 'certificate trust failure'; $hint = 'Check certificate trust, system time, and HTTPS inspection.' }
                'SecureChannelFailure' { $cause = 'TLS secure-channel failure'; $hint = 'Check TLS support and HTTPS inspection by a proxy or firewall.' }
                default { $cause = 'connection failure'; $hint = 'Check catalog host connectivity and proxy settings.' }
            }
            break
        }
        if ($inner.GetType().FullName -ceq 'System.Net.Http.HttpRequestException') {
            $cause = 'HTTP connection failure'
            if ($inner.PSObject.Properties.Match('HttpRequestError').Count -gt 0) {
                switch ([string]$inner.HttpRequestError) {
                    'NameResolutionError' { $cause = 'DNS resolution failure'; $hint = 'Check DNS resolution for the catalog host and proxy.' }
                    'SecureConnectionError' { $cause = 'TLS connection failure'; $hint = 'Check certificate trust, system time, and HTTPS inspection.' }
                    'ConnectionError' { $hint = 'Check catalog host connectivity, firewall rules, and proxy settings.' }
                    'ProxyTunnelError' { $cause = 'proxy tunnel failure'; $hint = 'Check proxy authentication and HTTPS tunneling permissions.' }
                }
            }
        }
    }
    if ($StatusCode -gt 0) {
        $cause = 'HTTP {0}' -f $StatusCode
        $hint = if ($StatusCode -ge 500) { 'The server or proxy returned an error; check service availability and retry later.' } else { 'Check the catalog endpoint and its supported request method.' }
    }
    return 'Cause: {0}. Exception: {1} (HRESULT 0x{2:X8}). {3}' -f $cause, $underlying.GetType().FullName, $underlying.HResult, $hint
}

function Invoke-CustomCatalogPage {
    param(
        [parameter(Mandatory = $true)][hashtable]$Context,
        [parameter(Mandatory = $true)][Microsoft.PowerShell.Commands.WebRequestSession]$WebSession,
        [ValidateRange(1, 1000)][int]$Page = 1,
        [uri]$Proxy
    )

    $body = @{
        keyword = $null; filters = @{}; date = $null; startTime = $null; endTime = $null
        sort = @{ column = 'start'; direction = 0 }; page = $Page; pageSize = 20
        aiSearch = $false; upcoming = $false; onDemand = $false; hasAssets = $false; onSchedule = $false
    }
    $requestStarted = $false
    try {
        if ($Context.Method -ne 'Get') { Initialize-CustomRequestContext -Context $Context -WebSession $WebSession -Proxy $Proxy }
        else { $Context.Uri = Get-CustomEventPageUri -BaseUrl $Context.CatalogUri.AbsoluteUri -Page $Page }
        $requestUri = [uri]$Context.Uri
        if (-not (Test-CustomCatalogEndpoint -Uri $requestUri -CatalogUri $Context.CatalogUri)) { throw 'Invalid Custom request endpoint.' }
        $method = if ($Context.Method) { $Context.Method } else { 'Post' }
        $parameters = @{
            Uri = $requestUri.AbsoluteUri; Method = $method; WebSession = $WebSession
            Headers = @{ Accept = 'application/json' }
            MaximumRedirection = 0; ErrorAction = 'Stop'; Verbose = $false; Debug = $false; TimeoutSec = 30
        }
        if ($method -eq 'Post') {
            $parameters.ContentType = 'application/json'
            $parameters.Body = $body | ConvertTo-Json -Depth 5 -Compress
            $parameters.Headers.Origin = $requestUri.GetLeftPart([System.UriPartial]::Authority)
        }
        if ($Context.Referer) {
            $refererUri = [uri]$Context.Referer
            if ($refererUri.GetLeftPart([System.UriPartial]::Authority) -ne $requestUri.GetLeftPart([System.UriPartial]::Authority)) { throw 'Invalid Custom referer origin.' }
            $parameters.Headers.Referer = $Context.Referer
        }
        if ($Context.AuthenticationRequired -and $script:MSAAuthOrigin -eq $requestUri.GetLeftPart([System.UriPartial]::Authority)) {
            $authHeaders = Get-MSAAuthApiHeaders
            if ($authHeaders.Authorization) { $parameters.Headers.Authorization = $authHeaders.Authorization }
        }
        if ($Proxy) { $parameters.Proxy = $Proxy }
        $requestStarted = $true
        $response = Invoke-WebWithRetry -ScriptBlock { Invoke-WebRequestNoRedirect -Parameters $parameters -Rest } -Variables @{ parameters = $parameters } -SuppressErrorDetails
    }
    catch {
        if ($_.Exception -is [System.Management.Automation.PipelineStoppedException]) { throw }
        if ($_.Exception.Data['CookieFileFailure']) { throw }
        $status = $null
        if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }
        if ($Context.AllowSearchFallback -and $Context.Method -eq 'Post' -and
            $requestUri.AbsolutePath -cmatch '^/api/schedule-builder/sessions/?$' -and
            $status -in @(301, 302, 303, 307, 308, 400, 401, 403, 404, 405)) {
            $Context.AllowSearchFallback = $false
            $builder = [System.UriBuilder]::new($requestUri)
            $builder.Path = '/api/schedule-builder/sessions/search'
            $builder.Fragment = ''
            $Context.Uri = $builder.Uri.AbsoluteUri
            Set-ObjectPropertyValue -Object $WebSession -Name 'EventRequestUri' -Value $Context.Uri
            Write-Verbose 'Trying the supported schedule-builder /sessions/search POST endpoint.'
            return Invoke-CustomCatalogPage -Context $Context -WebSession $WebSession -Page $Page -Proxy $Proxy
        }
        if ($Context.Method -eq 'Get' -and $Context.AllowMethodSwitch -ne $false -and $status -eq 405 -and [string]$_.Exception.Response.Headers.Allow -match '(?i)POST') {
            $Context.Method = 'Post'
            $Context.Uri = $Context.CatalogUri.AbsoluteUri
            $Context.Referer = $Context.CatalogUri.GetLeftPart([System.UriPartial]::Authority) + '/'
            return Invoke-CustomCatalogPage -Context $Context -WebSession $WebSession -Page $Page -Proxy $Proxy
        }
        if ($Context.Method -eq 'Post' -and $Context.AllowMethodSwitch -eq $true -and $status -eq 405 -and [string]$_.Exception.Response.Headers.Allow -match '(?i)GET') {
            $Context.Method = 'Get'
            $Context.AllowMethodSwitch = $false
            return Invoke-CustomCatalogPage -Context $Context -WebSession $WebSession -Page $Page -Proxy $Proxy
        }
        $cause = if ($status) { 'HTTP {0}' -f $status } elseif ($requestStarted) { 'request failure' } else { 'request context unavailable or invalid' }
        $summary = Get-CustomCatalogFailureSummary -Exception $_.Exception -StatusCode ([int]$status)
        $exception = [System.InvalidOperationException]::new(('Custom catalog page {0} failed ({1}). {2} Request credentials and response details are withheld.' -f $Page, $cause, $summary))
        $exception.Data['FailureKind'] = if ($status) { 'Http' } elseif ($requestStarted) { 'Transport' } else { 'Context' }
        if ($status) { $exception.Data['StatusCode'] = $status }
        if (Test-MSAAuthenticationRequired -Exception $_.Exception) {
            $Context.AuthenticationRequired = $true
            $exception.Data['AuthenticationRequired'] = $true
        }
        throw $exception
    }
    if ($response -is [string] -and (Test-MSAAuthenticationRequired -Response ([pscustomobject]@{ StatusCode = 200; Headers = @{ 'Content-Type' = 'text/html' }; Content = $response }) -RequestUri $requestUri.AbsoluteUri)) {
        $exception = [System.InvalidOperationException]::new('Custom catalog requires authentication.')
        $Context.AuthenticationRequired = $true
        $exception.Data['AuthenticationRequired'] = $true
        throw $exception
    }
    if ($response -and $response -isnot [string] -and $response.PSObject.Properties.Match('items').Count -gt 0) {
        $Context.AllowSearchFallback = $false
        return $response
    }
    if (-not $response -or $response -is [string] -or $response.PSObject.Properties.Match('sessions').Count -eq 0 -or $response.PSObject.Properties.Match('totalCount').Count -eq 0 -or $response.PSObject.Properties.Match('pageSize').Count -eq 0 -or $response.PSObject.Properties.Match('page').Count -eq 0) {
        $exception = [System.InvalidOperationException]::new('Custom portal returned an unexpected catalog response; no session cache was written.')
        $exception.Data['FailureKind'] = 'Response'
        throw $exception
    }
    $responsePage = 0
    $pageSize = 0
    $totalCount = 0
    if (-not [int]::TryParse([string]$response.page, [ref]$responsePage) -or
        -not [int]::TryParse([string]$response.pageSize, [ref]$pageSize) -or
        -not [int]::TryParse([string]$response.totalCount, [ref]$totalCount) -or
        $responsePage -ne $Page -or $pageSize -le 0 -or $totalCount -lt 0) {
        $exception = [System.InvalidOperationException]::new('Custom portal returned invalid pagination information; no session cache was written.')
        $exception.Data['FailureKind'] = 'Response'
        throw $exception
    }
    $Context.AllowSearchFallback = $false
    return $response
}

function Invoke-CustomSessionDetailRequest {
    param(
        [parameter(Mandatory = $true)][uri]$Uri,
        [parameter(Mandatory = $true)][Microsoft.PowerShell.Commands.WebRequestSession]$WebSession,
        [uri]$Proxy
    )

    if (-not (Test-IsCustomCatalogUri -Uri $Uri)) { throw 'Custom session details require an HTTPS URL without embedded credentials.' }
    $parameters = @{ Uri = $Uri.AbsoluteUri; Method = 'Get'; WebSession = $WebSession; Headers = @{ Accept = 'application/json' }; MaximumRedirection = 0; TimeoutSec = 30; ErrorAction = 'Stop'; Verbose = $false; Debug = $false }
    if ($Proxy) { $parameters.Proxy = $Proxy }
    if ($script:MSAAuthOrigin -eq $Uri.GetLeftPart([System.UriPartial]::Authority)) {
        $headers = Get-MSAAuthApiHeaders
        if ($headers.Authorization) { $parameters.Headers.Authorization = $headers.Authorization }
    }
    try {
        $details = Invoke-WebWithRetry -ScriptBlock { Invoke-WebRequestNoRedirect -Parameters $parameters -Rest } -Variables @{ parameters = $parameters } -SuppressErrorDetails
    }
    catch {
        if ($_.Exception -is [System.Management.Automation.PipelineStoppedException]) { throw }
        $status = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { $null }
        $exception = [System.InvalidOperationException]::new(('Custom session detail request failed{0}; credentials and response details withheld.' -f $(if ($status) { ' (HTTP {0})' -f $status } else { '' })))
        if ($status) { $exception.Data['StatusCode'] = $status }
        if (Test-MSAAuthenticationRequired -Exception $_.Exception) { $exception.Data['AuthenticationRequired'] = $true }
        throw $exception
    }
    if (-not $details -or $details -is [string]) { throw 'Custom session details did not return a JSON object.' }
    return $details
}

function Test-SessionHasOnDemandContent {
    param([parameter(Mandatory = $true)][object]$Session)

    foreach ($name in @('hasOnDemandContent', 'isOnDemand', 'onDemand')) {
        $property = $Session.PSObject.Properties[$name]
        if ($null -eq $property -or $null -eq $property.Value) { continue }
        if ($property.Value -is [bool]) { return $property.Value }
        $available = $false
        if ($property.Value -is [string] -and [bool]::TryParse($property.Value.Trim(), [ref]$available)) { return $available }
    }
    return $true
}

function Set-CustomSessionDetailMetadata {
    param(
        [parameter(Mandatory = $true)][object]$Session,
        [parameter(Mandatory = $true)][object]$Details,
        [uri]$CatalogUri
    )

    $eventInfo = $Details.eventInfo
    if (-not $eventInfo -or [string]$eventInfo.id -ne [string]$Session.id) { throw 'Custom session details do not match the selected session ID.' }
    if ($null -ne $eventInfo.virtual -and $eventInfo.virtual.PSObject.Properties.Match('videoEnabled').Count -gt 0) {
        Set-ObjectPropertyValue -Object $Session -Name 'hasOnDemandContent' -Value $eventInfo.virtual.videoEnabled
    }
    $videoUrl = [string]$eventInfo.virtual.videoUrl
    if ($videoUrl -and (Test-SessionHasOnDemandContent -Session $Session)) {
        $videoUri = $null
        if ([uri]::TryCreate($videoUrl, [System.UriKind]::Absolute, [ref]$videoUri) -and $videoUri.Scheme -eq 'https') {
            Set-ObjectPropertyValue -Object $Session -Name 'onDemand' -Value $videoUri.AbsoluteUri
        }
    }
    if ($eventInfo.eventData.Description) { Set-ObjectPropertyValue -Object $Session -Name 'description' -Value $eventInfo.eventData.Description }
    $speakers = @()
    foreach ($group in $eventInfo.facilitators.PSObject.Properties) {
        foreach ($speaker in @($group.Value)) {
            if ($speaker.formattedName) { $speakers += [string]$speaker.formattedName }
        }
    }
    if ($speakers.Count -gt 0) { Set-ObjectPropertyValue -Object $Session -Name 'speakerNames' -Value @($speakers | Select-Object -Unique) }
    $presentations = [System.Collections.Generic.List[string]]::new()
    foreach ($asset in @($eventInfo.assets) + @($eventInfo.externalResources)) {
        if (-not $asset) { continue }
        $url = Get-ObjectPropertyValue -Object $asset -Name @('url', 'linkUrl', 'downloadUrl')
        $assetUri = $null
        if ($url -and [uri]::TryCreate([string]$url, [System.UriKind]::Absolute, [ref]$assetUri) -and
            (Test-IsCustomCatalogUri -Uri $assetUri) -and $assetUri.AbsolutePath -match '(?i)\.(pptx|pdf)$') {
            $presentations.Add($assetUri.AbsoluteUri)
        }
        else {
            $assetId = [guid]::Empty
            $assetName = [string](Get-ObjectPropertyValue -Object $asset -Name @('name', 'fileName'))
            if ((Test-IsCustomCatalogUri -Uri $CatalogUri) -and $assetName -match '(?i)\.(pptx|pdf)$' -and
                [guid]::TryParse([string]$asset.id, [ref]$assetId)) {
                $presentations.Add(([uri]::new($CatalogUri, ('/api/asset/download-url?id={0}' -f $assetId))).AbsoluteUri)
            }
            elseif ($assetUri -and $assetName -match '(?i)\.(pptx|pdf)$' -and
                (Test-CustomAssetDownloadUri -Uri $assetUri -CatalogUri $CatalogUri)) {
                $presentations.Add($assetUri.AbsoluteUri)
            }
        }
    }
    $deckUrls = @(Get-SessionPresentationUrls -Session ([pscustomobject]@{ slideDecks = $presentations.ToArray() }))
    Set-ObjectPropertyValue -Object $Session -Name 'slideDecks' -Value $deckUrls
    if ($deckUrls.Count -gt 0) { Set-ObjectPropertyValue -Object $Session -Name 'slideDeck' -Value $deckUrls[0] }
}

function Test-CustomAssetDownloadUri {
    param([uri]$Uri, [uri]$CatalogUri)

    if (-not (Test-IsCustomCatalogUri -Uri $Uri) -or -not (Test-IsCustomCatalogUri -Uri $CatalogUri) -or
        $Uri.GetLeftPart([System.UriPartial]::Authority) -ne $CatalogUri.GetLeftPart([System.UriPartial]::Authority) -or
        $Uri.AbsolutePath -cne '/api/asset/download-url') { return $false }
    $query = [System.Web.HttpUtility]::ParseQueryString($Uri.Query)
    $assetId = [guid]::Empty
    return [guid]::TryParse([string]$query['id'], [ref]$assetId)
}

function Get-SessionPresentationUrls {
    param([parameter(Mandatory = $true)][object]$Session)

    $urls = [System.Collections.Generic.SortedDictionary[string, string]]::new([System.StringComparer]::Ordinal)
    $sources = @($Session.slideDecks)
    if (@($sources | Where-Object { $_ }).Count -eq 0) { $sources = @($Session.slideDeck) }
    foreach ($source in $sources) {
        $uri = $null
        if (-not [uri]::TryCreate([string]$source, [System.UriKind]::Absolute, [ref]$uri) -or
            -not (Test-IsCustomCatalogUri -Uri $uri)) { continue }
        $key = $uri.AbsoluteUri
        if ($uri.AbsolutePath -ceq '/api/asset/download-url') {
            $query = [System.Web.HttpUtility]::ParseQueryString($uri.Query)
            $assetId = [guid]::Empty
            if ([guid]::TryParse([string]$query['id'], [ref]$assetId)) {
                $key = '{0}?id={1}' -f $uri.GetLeftPart([System.UriPartial]::Path), $assetId
            }
        }
        if (-not $urls.ContainsKey($key)) { $urls.Add($key, $uri.AbsoluteUri) }
    }
    return @($urls.Values)
}

function Invoke-CustomAssetDownloadRequest {
    param(
        [parameter(Mandatory = $true)][uri]$Uri,
        [parameter(Mandatory = $true)][Microsoft.PowerShell.Commands.WebRequestSession]$WebSession,
        [uri]$Proxy
    )

    $parameters = @{ Uri = $Uri.AbsoluteUri; Method = 'Get'; WebSession = $WebSession; MaximumRedirection = 0; TimeoutSec = 30; ErrorAction = 'Stop'; Verbose = $false; Debug = $false; Headers = @{ Accept = 'application/json' } }
    if ($Proxy) { $parameters.Proxy = $Proxy }
    if ($script:MSAAuthOrigin -eq $Uri.GetLeftPart([System.UriPartial]::Authority)) { $parameters.Headers = Get-MSAAuthApiHeaders }
    try { $response = Invoke-WebWithRetry -ScriptBlock { Invoke-WebRequestNoRedirect -Parameters $parameters -Rest } -Variables @{ parameters = $parameters } -SuppressErrorDetails }
    catch {
        if ($_.Exception -is [System.Management.Automation.PipelineStoppedException]) { throw }
        $exception = [System.InvalidOperationException]::new('Presentation asset request failed; credentials and response details withheld.')
        if (Test-MSAAuthenticationRequired -Exception $_.Exception) { $exception.Data['AuthenticationRequired'] = $true }
        throw $exception
    }
    if ($response -is [string] -and (Test-MSAAuthenticationRequired -Response ([pscustomobject]@{ StatusCode = 200; Headers = @{ 'Content-Type' = 'text/html' }; Content = $response }) -RequestUri $Uri.AbsoluteUri)) {
        $exception = [System.InvalidOperationException]::new('Presentation asset requires authentication.')
        $exception.Data['AuthenticationRequired'] = $true
        throw $exception
    }
    $downloadUri = $null
    if ($response.success -isnot [bool] -or -not $response.success -or -not [uri]::TryCreate([string]$response.url, [System.UriKind]::Absolute, [ref]$downloadUri) -or
        -not (Test-IsCustomCatalogUri -Uri $downloadUri)) { throw 'Presentation asset returned an invalid download response; response details withheld.' }
    return $downloadUri.AbsoluteUri
}

function Resolve-CustomPresentationUrl {
    param(
        [parameter(Mandatory = $true)][uri]$Uri,
        [parameter(Mandatory = $true)][uri]$CatalogUri,
        [uri]$Proxy
    )

    if (-not (Test-CustomAssetDownloadUri -Uri $Uri -CatalogUri $CatalogUri)) { throw 'Invalid presentation asset endpoint.' }
    $origin = $CatalogUri.GetLeftPart([System.UriPartial]::Authority)
    if ($script:MSAAuthOrigin -ne $origin) { $script:MSAAuthWebSession = $null; $script:MSABearerToken = $null }
    $script:MSAAuthOrigin = $origin
    $webSession = $script:MSAAuthWebSession
    if ($script:CustomCookieFile) { $webSession = Get-CustomCookieWebSession -TargetUri $Uri }
    elseif (-not $webSession) { $webSession = Get-CachedMSAWebSession -CachePath (Join-Path $PSScriptRoot 'MSAAuth.cache') -CookieUri ([uri]($origin + '/')) }
    if (-not $webSession) { $webSession = New-Object Microsoft.PowerShell.Commands.WebRequestSession }
    try {
        $url = Invoke-CustomAssetDownloadRequest -Uri $Uri -WebSession $webSession -Proxy $Proxy
        $script:MSAAuthWebSession = $webSession
        return $url
    }
    catch {
        if (-not $_.Exception.Data['AuthenticationRequired']) { throw }
    }
    $state = @{}
    $validator = {
        param($session)
        $state.Url = Invoke-CustomAssetDownloadRequest -Uri $Uri -WebSession $session -Proxy $Proxy
        return $true
    }
    $webSession = Get-MSAAuthenticatedWebSession -TargetUri $Uri -Proxy $Proxy -ValidateCachedSession -PersistCache -SessionValidator $validator
    if (-not $webSession -or -not $state.Url) { throw 'Unable to authenticate for the presentation asset.' }
    return $state.Url
}

function Add-SessionPresentationDownloadJobs {
    param(
        [parameter(Mandatory = $true)][object]$Session,
        [parameter(Mandatory = $true)][uri]$CatalogUri,
        [parameter(Mandatory = $true)][string]$Folder,
        [parameter(Mandatory = $true)][string]$BaseName,
        [string]$Timestamp,
        [switch]$Overwrite,
        [uri]$Proxy
    )

    $sources = @(Get-SessionPresentationUrls -Session $Session)
    $result = [pscustomobject]@{ Existing = 0; Missing = 0 }
    for ($index = 0; $index -lt $sources.Count; $index++) {
        $name = if ($index -gt 0) { '{0}-{1}' -f $BaseName, $index } else { $BaseName }
        try {
            $sourceUri = [uri]$sources[$index]
            $isResolver = Test-CustomAssetDownloadUri -Uri $sourceUri -CatalogUri $CatalogUri
            if ($sourceUri.AbsolutePath -ceq '/api/asset/download-url' -and -not $isResolver) { throw 'Invalid asset origin.' }
            $downloadUrl = if ($isResolver) { Resolve-CustomPresentationUrl -Uri $sourceUri -CatalogUri $CatalogUri -Proxy $Proxy } else { $sourceUri.AbsoluteUri }
            $downloadUri = [uri]$downloadUrl
            if (-not (Test-IsCustomCatalogUri -Uri $downloadUri)) { throw 'Invalid presentation URL.' }
            $extension = [System.IO.Path]::GetExtension($downloadUri.AbsolutePath).ToLowerInvariant()
            if ($extension -in @('.pptx', '.pdf')) {
                $file = [System.IO.Path]::Combine($Folder, ($name + $extension))
                if (-not $Overwrite -and (Test-Path -LiteralPath $file) -and (Get-Item -LiteralPath $file).Length -gt 0) {
                    Write-Host ('Slidedeck exists {0}' -f ($name + $extension)) -ForegroundColor Gray
                    $result.Existing++
                    continue
                }
            }
            $headers = $null
            $parameters = @{ Uri = $downloadUrl; Method = 'Head'; MaximumRedirection = 0; TimeoutSec = 30; ErrorAction = 'Stop'; Verbose = $false; Debug = $false }
            if ($Proxy) { $parameters.Proxy = $Proxy }
            if ($script:CustomCookieFile) {
                $parameters.WebSession = Get-CustomCookieWebSession -TargetUri $downloadUri
            }
            elseif ($downloadUri.GetLeftPart([System.UriPartial]::Authority) -eq $CatalogUri.GetLeftPart([System.UriPartial]::Authority) -and
                $script:MSAAuthOrigin -eq $CatalogUri.GetLeftPart([System.UriPartial]::Authority) -and $script:MSAAuthWebSession) {
                $parameters.WebSession = $script:MSAAuthWebSession
                $headers = Get-MSAAuthApiHeaders
                $cookieHeader = $script:MSAAuthWebSession.Cookies.GetCookieHeader($downloadUri)
                if ($cookieHeader) { $headers.Cookie = $cookieHeader }
                $parameters.Headers = $headers
            }
            $response = Invoke-WebRequestNoRedirect -Parameters $parameters
            if ($script:CustomCookieFile -and (Test-MSAAuthenticationRequired -Response $response -RequestUri $downloadUrl)) {
                throw (New-CustomCookieFileException -Message 'Presentation download rejected the exported cookies. Sign in and export cookies again.')
            }
            $contentType = ([string]($response.Headers['Content-Type'] | Select-Object -First 1) -split ';')[0].Trim()
            if ($contentType -eq 'application/pdf') { $extension = '.pdf' }
            elseif ($contentType -eq 'application/vnd.openxmlformats-officedocument.presentationml.presentation') { $extension = '.pptx' }
            elseif ($contentType -ne 'application/octet-stream' -or $extension -notin @('.pptx', '.pdf')) { throw 'Presentation response has an unsupported file type.' }
            $file = [System.IO.Path]::Combine($Folder, ($name + $extension))
            if (-not $Overwrite -and (Test-Path -LiteralPath $file) -and (Get-Item -LiteralPath $file).Length -gt 0) {
                Write-Host ('Slidedeck exists {0}' -f ($name + $extension)) -ForegroundColor Gray
                $result.Existing++
                continue
            }
            Write-Verbose ('Downloading presentation {0} for session {1}; source credentials withheld.' -f ($index + 1), $Session.sessionCode)
            Add-BackgroundDownloadJob -Type 1 -FilePath $file -DownloadUrl $downloadUrl -File $file -Timestamp $Timestamp -ScheduleCode $Session.sessionCode -Title $Session.title -Headers $headers -Proxy $Proxy
        }
        catch {
            if ($_.Exception -is [System.Management.Automation.PipelineStoppedException] -or $_.Exception.Message -match '(?i)cancelled') { throw }
            if ($script:CustomCookieFile -and ($_.Exception.Data['CookieFileFailure'] -or (Test-MSAAuthenticationRequired -Exception $_.Exception))) {
                throw (New-CustomCookieFileException -Message 'Presentation authentication failed with exported cookies; credentials and response details withheld.')
            }
            $result.Missing++
            Write-Warning ('Unable to download presentation {0} for session {1}; credentials and response details withheld.' -f ($index + 1), $Session.sessionCode)
        }
    }
    return $result
}

function Get-CustomSessionPageMetadata {
    param(
        [parameter(Mandatory = $true)][object]$Details,
        [parameter(Mandatory = $true)][uri]$CatalogUri,
        [parameter(Mandatory = $true)][Microsoft.PowerShell.Commands.WebRequestSession]$WebSession,
        [uri]$Proxy
    )

    $listUri = $null
    if (-not [uri]::TryCreate([string]$Details.eventListUrl, [System.UriKind]::Absolute, [ref]$listUri) -or
        -not (Test-IsCustomCatalogUri -Uri $listUri) -or
        $listUri.GetLeftPart([System.UriPartial]::Authority) -ne $CatalogUri.GetLeftPart([System.UriPartial]::Authority) -or
        $listUri.AbsolutePath -notmatch '/sessions/?$') { return $Details }
    $sessionId = [guid]::Empty
    if (-not [guid]::TryParse([string]$Details.eventInfo.id, [ref]$sessionId)) { return $Details }
    $builder = [System.UriBuilder]::new($listUri)
    $builder.Path = $listUri.AbsolutePath.TrimEnd('/').Substring(0, $listUri.AbsolutePath.TrimEnd('/').LastIndexOf('/')) + '/session-details/' + $sessionId
    $builder.Query = ''
    $builder.Fragment = ''
    $pageUri = $builder.Uri
    $parameters = @{ Uri = $pageUri.AbsoluteUri; Method = 'Get'; WebSession = $WebSession; MaximumRedirection = 0; TimeoutSec = 30; ErrorAction = 'Stop'; Verbose = $false; Debug = $false }
    if ($Proxy) { $parameters.Proxy = $Proxy }
    Write-Verbose ('Inspecting published session page for assets: {0}' -f $pageUri.GetLeftPart([System.UriPartial]::Path))
    try { $page = Invoke-WebRequestNoRedirect -Parameters $parameters }
    catch {
        if ($_.Exception -is [System.Management.Automation.PipelineStoppedException]) { throw }
        $exception = [System.InvalidOperationException]::new('Unable to inspect the selected session page; credentials and response details withheld.')
        if (Test-MSAAuthenticationRequired -Exception $_.Exception) { $exception.Data['AuthenticationRequired'] = $true }
        throw $exception
    }
    if (Test-MSAAuthenticationRequired -Response $page -RequestUri $pageUri.AbsoluteUri) {
        $exception = [System.InvalidOperationException]::new('Selected session page requires authentication.')
        $exception.Data['AuthenticationRequired'] = $true
        throw $exception
    }
    $tag = [regex]::Match([string]$page.Content, '(?is)<hio-event-details\b[^>]*>').Value
    $attributes = @{}
    foreach ($match in [regex]::Matches($tag, '\s([\w-]+)=["'']([^"'']*)["'']')) {
        $attributes[$match.Groups[1].Value.ToLowerInvariant()] = [System.Net.WebUtility]::HtmlDecode($match.Groups[2].Value)
    }
    if ($attributes['pc-id'] -and $attributes['event-id'] -eq [string]$sessionId) {
        $detailUri = [uri]::new($CatalogUri, ('/api/event/details/{0}?pcId={1}' -f $sessionId, [uri]::EscapeDataString($attributes['pc-id'])))
        if ($attributes['visibility-pc-id']) {
            $detailUri = [uri]($detailUri.AbsoluteUri + '&visibilityPcId=' + [uri]::EscapeDataString($attributes['visibility-pc-id']))
        }
        $pageDetails = Invoke-CustomSessionDetailRequest -Uri $detailUri -WebSession $WebSession -Proxy $Proxy
        if (-not $pageDetails.eventInfo -or [string]$pageDetails.eventInfo.id -ne [string]$sessionId) { throw 'Selected session page details have an unexpected ID.' }
        $Details = $pageDetails
        Set-ObjectPropertyValue -Object $WebSession -Name 'EventDetailRequests' -Value @{ ([string]$sessionId) = $detailUri.AbsoluteUri }
    }
    $assets = [System.Collections.Generic.List[object]]::new()
    foreach ($asset in @($Details.eventInfo.assets)) { if ($asset) { $assets.Add($asset) } }
    foreach ($link in @($page.Links)) {
        $assetUri = $null
        if (-not $link.href -or -not [uri]::TryCreate($pageUri, [System.Net.WebUtility]::HtmlDecode([string]$link.href), [ref]$assetUri) -or
            -not (Test-IsCustomCatalogUri -Uri $assetUri)) { continue }
        if ($assetUri.AbsolutePath -match '(?i)\.(pptx|pdf)$') {
            $assets.Add([pscustomobject]@{ url = $assetUri.AbsoluteUri })
        }
        elseif (Test-CustomAssetDownloadUri -Uri $assetUri -CatalogUri $CatalogUri) {
            $assetId = [System.Web.HttpUtility]::ParseQueryString($assetUri.Query)['id']
            $assets.Add([pscustomobject]@{ id = $assetId; url = $assetUri.AbsoluteUri; name = [string]$link.title })
        }
    }
    Set-ObjectPropertyValue -Object $Details.eventInfo -Name 'assets' -Value $assets.ToArray()
    Write-Verbose ('Inspected selected session page; {0} asset record(s) available for presentation extraction.' -f $assets.Count)
    return $Details
}

function Update-CustomSessionDetails {
    param(
        [parameter(Mandatory = $true)][object]$Session,
        [parameter(Mandatory = $true)][uri]$CatalogUri,
        [uri]$Proxy
    )

    if (-not $Session.id -or $CatalogUri.AbsolutePath -notmatch '(?i)^/api/schedule-builder/sessions(?:/search)?/?$') { return }
    $sessionId = [guid]::Empty
    if (-not [guid]::TryParse([string]$Session.id, [ref]$sessionId)) { return }
    $detailUri = [uri]::new($CatalogUri, ('/api/event/details/{0}' -f $sessionId))
    $cachedSession = if ($script:CustomCookieFile) {
        Get-CustomCookieWebSession -TargetUri $detailUri
    }
    elseif ($script:MSAAuthOrigin -eq $CatalogUri.GetLeftPart([System.UriPartial]::Authority) -and $script:MSAAuthWebSession) {
        $script:MSAAuthWebSession
    }
    else {
        Get-CachedMSAWebSession -CachePath (Join-Path $PSScriptRoot 'MSAAuth.cache') -CookieUri ([uri]($CatalogUri.GetLeftPart([System.UriPartial]::Authority) + '/'))
    }
    $state = @{ Uri = $detailUri; DetailRequests = @{}; RequestUri = $null; PortalUri = $null }
    if ($cachedSession) {
        $state.RequestUri = $cachedSession.EventRequestUri
        $state.PortalUri = $cachedSession.EventPortalUri
        if ($cachedSession.EventDetailRequests) {
            foreach ($key in $cachedSession.EventDetailRequests.Keys) { $state.DetailRequests[$key] = $cachedSession.EventDetailRequests[$key] }
        }
        $requestCandidates = @($state.DetailRequests[[string]$sessionId]) + @($state.DetailRequests.Values)
        foreach ($requestCandidate in $requestCandidates) {
            $candidateUri = $null
            if (-not [uri]::TryCreate([string]$requestCandidate, [System.UriKind]::Absolute, [ref]$candidateUri)) { continue }
            if ($candidateUri.GetLeftPart([System.UriPartial]::Authority) -ne $detailUri.GetLeftPart([System.UriPartial]::Authority) -or
                $candidateUri.UserInfo -or $candidateUri.AbsolutePath -notmatch '(?i)^/api/event/details/[0-9a-f-]{36}/?$') { continue }
            $requestBuilder = [System.UriBuilder]::new($candidateUri)
            $requestBuilder.Path = $detailUri.AbsolutePath
            $requestBuilder.Fragment = ''
            $state.Uri = $requestBuilder.Uri
            break
        }
    }
    $validator = {
        param($webSession)
        try {
            $state.Details = Invoke-CustomSessionDetailRequest -Uri $state.Uri -WebSession $webSession -Proxy $Proxy
            if (-not $state.Details.eventInfo -or [string]$state.Details.eventInfo.id -ne [string]$sessionId) { throw 'Custom session detail response has an unexpected ID.' }
        }
        catch {
            if ($_.Exception -is [System.Management.Automation.PipelineStoppedException]) { throw }
            if ($script:CustomCookieFile) {
                Initialize-CustomCookieRequestContext -WebSession $webSession -CatalogUri $CatalogUri -SessionId ([string]$sessionId) -Proxy $Proxy
            }
            else {
                if (-not $webSession.EventPortalUri -and $state.PortalUri) { Set-ObjectPropertyValue -Object $webSession -Name 'EventPortalUri' -Value $state.PortalUri }
                if (-not (Update-CustomPortalRequestContext -WebSession $webSession -CatalogUri $CatalogUri -SessionId ([string]$sessionId) -Proxy $Proxy)) { throw }
            }
            $state.Uri = [uri]$webSession.EventDetailRequests[[string]$sessionId]
            $state.RequestUri = $webSession.EventRequestUri
            $state.Details = Invoke-CustomSessionDetailRequest -Uri $state.Uri -WebSession $webSession -Proxy $Proxy
            if (-not $state.Details.eventInfo -or [string]$state.Details.eventInfo.id -ne [string]$sessionId) { throw 'Custom session detail response has an unexpected ID.' }
        }
        $state.Details = Get-CustomSessionPageMetadata -Details $state.Details -CatalogUri $CatalogUri -WebSession $webSession -Proxy $Proxy
        if ($webSession.EventDetailRequests -and $webSession.EventDetailRequests[[string]$sessionId]) {
            $state.Uri = [uri]$webSession.EventDetailRequests[[string]$sessionId]
        }
        $state.DetailRequests = @{ ([string]$sessionId) = $state.Uri.AbsoluteUri }
        Set-ObjectPropertyValue -Object $webSession -Name 'EventDetailRequests' -Value $state.DetailRequests
        if ($state.RequestUri) { Set-ObjectPropertyValue -Object $webSession -Name 'EventRequestUri' -Value $state.RequestUri }
        $portalUri = $null
        if ([uri]::TryCreate([string]$state.Details.eventListUrl, [System.UriKind]::Absolute, [ref]$portalUri) -and $portalUri.GetLeftPart([System.UriPartial]::Authority) -eq $CatalogUri.GetLeftPart([System.UriPartial]::Authority)) {
            Set-ObjectPropertyValue -Object $webSession -Name 'EventPortalUri' -Value $portalUri.GetLeftPart([System.UriPartial]::Path)
        }
        return $true
    }
    $capture = {
        param($port, $webSession)
        try {
            $null = & $validator $webSession
            return
        }
        catch {
            if ($_.Exception -is [System.Management.Automation.PipelineStoppedException]) { throw }
        }
        $context = Get-CustomEventBrowserRequestContext -DebugPort $port -CatalogUri $detailUri -SessionId ([string]$sessionId)
        $state.Uri = [uri]$context.Uri
        $null = & $validator $webSession
    }
    try {
        Write-Verbose ('Retrieving published media information for Custom session {0}' -f $Session.sessionCode)
        $webSession = Get-MSAAuthenticatedWebSession -TargetUri $detailUri -Proxy $Proxy -ValidateCachedSession -PersistCache -SessionValidator $validator -OnAuthenticatedBrowser $capture
        if (-not $webSession) { throw 'Unable to obtain selected session details.' }
        Set-CustomSessionDetailMetadata -Session $Session -Details $state.Details -CatalogUri $CatalogUri
    }
    catch {
        if ($_.Exception -is [System.Management.Automation.PipelineStoppedException] -or $_.Exception.Message -match '(?i)cancelled') { throw }
        if ($script:CustomCookieFile) { throw }
        Write-Warning ('Unable to resolve media for Custom session {0}: {1}' -f $Session.sessionCode, $_.Exception.Message)
    }
}

function Get-CustomEventCatalog {
    param(
        [parameter(Mandatory = $true)][string]$CatalogUrl,
        [uri]$Proxy
    )

    $catalogUri = [uri]$CatalogUrl
    if (-not (Test-IsCustomCatalogUri -Uri $catalogUri)) { throw 'Custom event catalogs require an absolute HTTPS URL without embedded credentials.' }
    $catalogContext = @{ CatalogUri = $catalogUri; Method = 'Get'; AllowMethodSwitch = $false; AllowSearchFallback = $true }
    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $publicResponse = $null
    $needsAuthFlow = $false
    try {
        Write-Verbose 'Trying anonymous GET for the Custom session catalog.'
        $publicResponse = Invoke-CustomCatalogPage -Context $catalogContext -WebSession $session -Proxy $Proxy
    }
    catch {
        if ($_.Exception -is [System.Management.Automation.PipelineStoppedException]) { throw }
        $statusCode = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { $_.Exception.Data['StatusCode'] }
        if (-not $_.Exception.Data['AuthenticationRequired'] -and (-not $statusCode -or $statusCode -ge 500)) {
            if ($_.Exception.Data['FailureKind']) { throw }
            $summary = Get-CustomCatalogFailureSummary -Exception $_.Exception -StatusCode ([int]$statusCode)
            throw ('Unable to inspect the Custom catalog; no authentication was attempted because the request failed without a recoverable HTTP response. {0} Request credentials and response details are withheld.' -f $summary)
        }
        $catalogContext.AuthenticationRequired = [bool]$_.Exception.Data['AuthenticationRequired']
        if ($script:CustomCookieFile) { Write-Verbose ('Anonymous catalog GET failed (HTTP {0}); trying anonymous POST before exported cookies.' -f $statusCode) }
        else { Write-Verbose ('Anonymous catalog GET failed (HTTP {0}); trying cached POST context.' -f $statusCode) }
    }
    if (-not $publicResponse) {
        $catalogContext.Method = 'Post'
        $catalogContext.AllowMethodSwitch = $true
        $catalogContext.Uri = $catalogUri.AbsoluteUri
        $catalogContext.Referer = $catalogUri.GetLeftPart([System.UriPartial]::Authority) + '/'
        $origin = $catalogUri.GetLeftPart([System.UriPartial]::Authority)
        if ($script:MSAAuthOrigin -ne $origin) {
            $script:MSAAuthWebSession = $null
            $script:MSABearerToken = $null
        }
        $script:MSAAuthOrigin = $origin
        $cachedSession = if ($script:CustomCookieFile) { $null } else { Get-CachedMSAWebSession -CachePath (Join-Path $PSScriptRoot 'MSAAuth.cache') -CookieUri ([uri]($origin + '/')) }
        if ($cachedSession) {
            $session = $cachedSession
            $catalogContext.AuthenticationRequired = $true
            $cachedRequestUri = $null
            if ($cachedSession.EventRequestUri -and [uri]::TryCreate([string]$cachedSession.EventRequestUri, [System.UriKind]::Absolute, [ref]$cachedRequestUri) -and (Test-CustomCatalogEndpoint -Uri $cachedRequestUri -CatalogUri $catalogUri)) {
                $catalogContext.Uri = $cachedRequestUri.AbsoluteUri
                if ($cachedSession.EventPortalUri) { $catalogContext.Referer = $cachedSession.EventPortalUri }
            }
            Write-Verbose 'Trying Custom catalog POST with cached Microsoft authentication.'
        }
        elseif ($script:CustomCookieFile) { Write-Verbose 'Trying anonymous POST before exported-cookie authentication.' }
        else { Write-Verbose 'No matching authentication cache; trying anonymous POST before browser sign-in.' }
        try {
            try {
                $publicResponse = Invoke-CustomCatalogPage -Context $catalogContext -WebSession $session -Proxy $Proxy
            }
            catch {
                if ($_.Exception -is [System.Management.Automation.PipelineStoppedException] -or -not $cachedSession -or
                    -not (Update-CustomPortalRequestContext -WebSession $cachedSession -CatalogUri $catalogUri -Proxy $Proxy)) { throw }
                $catalogContext.Uri = $cachedSession.EventRequestUri
                $catalogContext.Referer = $cachedSession.EventPortalUri
                $publicResponse = Invoke-CustomCatalogPage -Context $catalogContext -WebSession $session -Proxy $Proxy
                Save-MSAAuthCookieHeader -CachePath (Join-Path $PSScriptRoot 'MSAAuth.cache') -CookieUri ([uri]($origin + '/')) -WebSession $cachedSession
            }
            if ($cachedSession) { $script:MSAAuthWebSession = $cachedSession }
        }
        catch {
            if ($_.Exception -is [System.Management.Automation.PipelineStoppedException]) { throw }
            $statusCode = $_.Exception.Data['StatusCode']
            if (-not $_.Exception.Data['AuthenticationRequired'] -and (-not $statusCode -or $statusCode -ge 500)) { throw }
            $needsAuthFlow = $true
            $catalogContext.AuthenticationRequired = $catalogContext.AuthenticationRequired -or [bool]$_.Exception.Data['AuthenticationRequired']
            $discoveredUri = $null
            if ($script:CustomCookieFile -and [uri]::TryCreate([string]$catalogContext.Uri, [System.UriKind]::Absolute, [ref]$discoveredUri) -and
                (Test-CustomCatalogEndpoint -Uri $discoveredUri -CatalogUri $catalogUri)) {
                $catalogContext.DiscoveredRequestUri = $discoveredUri.AbsoluteUri
            }
            $catalogContext.Uri = $null
            $catalogContext.Referer = $null
            if ($script:CustomCookieFile) { Write-Verbose 'Anonymous POST failed; using exported cookies without interactive authentication.' }
            else { Write-Verbose 'Cached or anonymous POST failed; using the shared provider-aware authentication flow.' }
        }
    }
    $authParameters = @{ Proxy = $Proxy; ValidateCachedSession = $true; PersistCache = $true }
    $authParameters.SessionValidator = {
        param($webSession)
        if ($script:CustomCookieFile -and $catalogContext.DiscoveredRequestUri -and -not $webSession.EventRequestUri) {
            Set-ObjectPropertyValue -Object $webSession -Name 'EventRequestUri' -Value $catalogContext.DiscoveredRequestUri
        }
        try { $catalogContext.FirstPage = Invoke-CustomCatalogPage -Context $catalogContext -WebSession $webSession -Proxy $Proxy }
        catch {
            if ($_.Exception -is [System.Management.Automation.PipelineStoppedException]) { throw }
            if (-not $script:CustomCookieFile -or $_.Exception.Data['FailureKind'] -in @('Transport', 'Response')) { throw }
            Initialize-CustomCookieRequestContext -WebSession $webSession -CatalogUri $catalogUri -Proxy $Proxy
            $catalogContext.Uri = $webSession.EventRequestUri
            $catalogContext.Referer = $webSession.EventPortalUri
            $catalogContext.FirstPage = Invoke-CustomCatalogPage -Context $catalogContext -WebSession $webSession -Proxy $Proxy
        }
        if ($catalogContext.Referer) { Set-ObjectPropertyValue -Object $webSession -Name 'EventPortalUri' -Value $catalogContext.Referer }
        Set-ObjectPropertyValue -Object $webSession -Name 'EventRequestUri' -Value $catalogContext.Uri
        return $true
    }
    $authParameters.OnAuthenticatedBrowser = {
        param($debugPort, $browserSession)
        try {
            $null = & $authParameters.SessionValidator $browserSession
            return
        }
        catch {
            if ($_.Exception.Data['AuthenticationRequired']) {
                $token = Get-MSALAccessTokenFromCdpPage -DebugPort $debugPort -CookieUri $catalogUri
                if ($token) {
                    $script:MSABearerToken = $token
                    $null = & $authParameters.SessionValidator $browserSession
                    return
                }
            }
            if ($catalogContext.Method -ne 'Post') { throw }
            $captured = Get-CustomEventBrowserRequestContext -DebugPort $debugPort -CatalogUri $catalogUri
            $catalogContext.Uri = $captured.Uri
            $catalogContext.Referer = $captured.Referer
            $null = & $authParameters.SessionValidator $browserSession
        }
    }
    if ($needsAuthFlow) {
        $session = Get-MSAAuthenticatedWebSession -TargetUri $catalogUri @authParameters
    }
    else { $catalogContext.FirstPage = $publicResponse }
    $isCustom = $catalogContext.FirstPage -and $catalogContext.FirstPage.PSObject.Properties.Match('sessions').Count -gt 0
    if (-not $session) { throw 'Unable to authenticate with the custom event catalog.' }

    # Use a browser-like User-Agent so the API returns full session data rather than null placeholders.
    if ($session) {
        $session.UserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36'
    }

    $data = [System.Collections.ArrayList]@()
    $seenSessionKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet', [string[]]('sessionCode', 'title'))
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)

    $page = 1
    $totalPages = $null
    $authRefreshed = $false

    try {
        while ($true) {

            $status = 'Processing page {0}' -f $page
            $percent = 0
            if ($totalPages) {
                $status = 'Processing page {0} of {1}' -f $page, $totalPages
                $percent = [Math]::Min(100, [int](($page / $totalPages) * 100))
            }
            Write-Progress -Id 1 -Activity 'Retrieving Session Catalog' -Status $status -PercentComplete $percent

            try {
                $response = if ($page -eq 1 -and $catalogContext.FirstPage) { $catalogContext.FirstPage } else { Invoke-CustomCatalogPage -Context $catalogContext -WebSession $session -Page $page -Proxy $Proxy }
            }
            catch {
                if ($_.Exception -is [System.Management.Automation.PipelineStoppedException]) { throw }
                if ($_.Exception.Data['AuthenticationRequired'] -or ($catalogContext.Method -eq 'Post' -and $_.Exception.Data['StatusCode'] -eq 400)) {
                    if (-not $authRefreshed) {
                        $authRefreshed = $true
                        if ($script:CustomCookieFile) {
                            $session = Get-CustomCookieWebSession -TargetUri $catalogUri
                            if ($catalogContext.Method -eq 'Post') {
                                Initialize-CustomCookieRequestContext -WebSession $session -CatalogUri $catalogUri -Proxy $Proxy
                                $catalogContext.Uri = $session.EventRequestUri
                                $catalogContext.Referer = $session.EventPortalUri
                            }
                            $catalogContext.FirstPage = $null
                            continue
                        }
                        $catalogContext.Uri = $null
                        $catalogContext.FirstPage = $null
                        $script:MSAAuthInteractiveAttemptedByHost.Remove($catalogUri.Host.ToLowerInvariant())
                        $session = Get-MSAAuthenticatedWebSession -TargetUri $catalogUri @authParameters -ForceInteractive
                        if (-not $session) { throw 'Custom authentication refresh failed.' }
                        continue
                    }
                    if ($script:CustomCookieFile) {
                        throw (New-CustomCookieFileException -Message 'Custom catalog pagination rejected the exported cookies or page context. No interactive authentication was attempted.')
                    }
                    throw
                }
                throw ('Problem retrieving custom session catalog page {0}: {1}' -f $page, $error[0])
            }

            if ($isCustom) {
                $totalPages = [int][Math]::Ceiling([int]$response.totalCount / [double][int]$response.pageSize)
                if ($totalPages -gt 1000) { throw 'Custom catalog exceeds the pagination safety limit.' }
            }
            if (-not $totalPages -and $response.PSObject.Properties.Match('totalPages').Count -gt 0) {
                $totalPages = [int]$response.totalPages
            }

            $items = @(if ($isCustom) { $response.sessions } else { $response.items }) | Where-Object { $null -ne $_ }
            $items = @($items)
            if ($items.Count -eq 0) {
                if ($isCustom -and $data.Count -ne [int]$response.totalCount) { throw 'Custom portal returned an incomplete catalog; no session cache was written.' }
                break
            }

            $addedThisPage = 0
            foreach ($item in $items) {
                if ($isCustom) {
                    Set-ObjectPropertyValue -Object $item -Name 'startDateTime' -Value (ConvertTo-SessionDateTime -Value $item.start)
                    Set-ObjectPropertyValue -Object $item -Name 'endDateTime' -Value (ConvertTo-SessionDateTime -Value $item.end)
                    Set-ObjectPropertyValue -Object $item -Name 'speakerNames' -Value @(ConvertTo-MetadataValueList -Value $item.speakers)
                }
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

                $dedupeKey = if ($isCustom -and $item.id) {
                    [string]$item.id
                }
                elseif (-not [string]::IsNullOrWhiteSpace($canonicalCode)) {
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
                    $addedThisPage++
                }
            }

            if ($isCustom -and $addedThisPage -eq 0) { throw 'Custom pagination did not advance; no session cache was written.' }
            if ($totalPages -and $page -ge $totalPages) {
                if ($isCustom -and $data.Count -ne [int]$response.totalCount) { throw 'Custom portal returned an incomplete catalog; no session cache was written.' }
                break
            }

            $page++
        }

    }
    finally { Write-Progress -Id 1 -Completed -Activity 'Finished retrieval of catalog' }

    if ($data.Count -eq 0) {
        Write-Warning ('No sessions were retrieved from the catalog at {0}.' -f $CatalogUrl)
        Write-Warning ('If the portal requires a culture or locale parameter (e.g. ?culture=en-us), include it in -EventUrl.')
        Write-Warning ('Example: -EventUrl ''https://summit.microsoft.com/api/sessions?culture=en-us''')
    }
    else {
        Write-Verbose ('Retrieved {0} session(s) from catalog.' -f $data.Count)
    }

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
        $webSession = Get-MSAAuthenticatedWebSession -TargetUri $sessionPageUri -Proxy $Proxy -ValidateCachedSession -PersistCache
        if (-not $webSession) {
            $script:CustomSignedOnDemandAuthRequiredBySession[$sessionCode] = $true
            Write-Warning ('Unable to acquire authenticated web session for {0}; cannot resolve signed media URL.' -f $sessionCode)
            return $OnDemandUrl
        }
        $sessionPage = Invoke-WebRequest -Uri $sessionPageUri.AbsoluteUri -Method Get -WebSession $webSession -Proxy $Proxy -DisableKeepAlive -ErrorAction Stop -Verbose:$false
    }
    catch {
        if ($script:CustomCookieFile) {
            if ($_.Exception -is [System.Management.Automation.PipelineStoppedException] -or $_.Exception.Data['CookieFileFailure']) { throw }
            throw (New-CustomCookieFileException -Message 'Unable to resolve the signed media URL with exported cookies; credentials and response details withheld.')
        }
        Write-Warning ('Unable to retrieve session page for {0}: {1}' -f $sessionCode, $_.Exception.Message)
        return $OnDemandUrl
    }

    if (Test-MSAAuthenticationRequired -Response $sessionPage) {
        if ($script:CustomCookieFile) {
            throw (New-CustomCookieFileException -Message 'The session page rejected the exported cookies. Sign in and export cookies again.')
        }
        Write-Verbose ('Session page for {0} returned a sign-in response; forcing interactive authentication and retrying.' -f $sessionCode)
        $script:MSAAuthWebSession = $null

        try {
            $webSession = Get-MSAAuthenticatedWebSession -TargetUri $sessionPageUri -Proxy $Proxy -ValidateCachedSession -ForceInteractive -PersistCache
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
        $coreConfiguration = ($coreConfigurationMatch.Groups['json'].Value | ConvertFrom-EventJson -Depth 100)
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
            $captionsConfiguration = ($captionsConfigurationMatch.Groups['json'].Value | ConvertFrom-EventJson -Depth 100)
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
            $coreConfiguration = ($coreConfigurationMatch.Groups['json'].Value | ConvertFrom-EventJson -Depth 100)
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
            $assets = $arrAssetsMatch.Groups['json'].Value | ConvertFrom-EventJson -Depth 50
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
        $mediusSession = if ($script:CustomCookieFile) { Get-CustomCookieWebSession -TargetUri ([uri]$OnDemandUrl) } else { New-Object Microsoft.PowerShell.Commands.WebRequestSession }
        $embedRequestParams = @{
            Uri         = $OnDemandUrl
            Method      = 'Get'
            WebSession  = $mediusSession
            ErrorAction = 'Stop'
            Verbose     = $false
        }
        if ($Proxy) { $embedRequestParams.Proxy = $Proxy }
        $embedResponse = Invoke-WebRequest @embedRequestParams
        if ($script:CustomCookieFile -and (Test-MSAAuthenticationRequired -Response $embedResponse -RequestUri $OnDemandUrl)) {
            throw (New-CustomCookieFileException -Message 'The video page rejected the exported cookies. Sign in and export cookies again.')
        }

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
        if ($script:CustomCookieFile) {
            throw (New-CustomCookieFileException -Message 'The video service did not return a media token with exported cookies; credentials and response details withheld.')
        }
    }
    catch {
        if ($script:CustomCookieFile) {
            if ($_.Exception -is [System.Management.Automation.PipelineStoppedException] -or $_.Exception.Data['CookieFileFailure']) { throw }
            throw (New-CustomCookieFileException -Message 'Unable to retrieve a media token with exported cookies; credentials and response details withheld.')
        }
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
            RawToken           = $null
            AuthorizationToken = $null
            ExpiryUtc          = $null
            HasParsedJson      = $false
        }
    }

    $authorizationToken = $KmsBearerTokenRaw
    $expiryUtc = $null
    $hasParsedJson = $false

    if ($KmsBearerTokenRaw.TrimStart().StartsWith('{')) {
        try {
            $tokenJson = $KmsBearerTokenRaw | ConvertFrom-EventJson -Depth 20
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
        RawToken           = $KmsBearerTokenRaw
        AuthorizationToken = $authorizationToken
        ExpiryUtc          = $expiryUtc
        HasParsedJson      = $hasParsedJson
    }
}

function Test-IsProtectedContentUrl {
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

    # URLs with Azure SAS token parameters are already self-authenticated; MSA auth is not needed.
    if ($parsedUri.Query -match '(?i)[?&]sig=') {
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

function ConvertTo-CaptionLanguageTag {
    # Normalizes a caption language such as 'bg-BG' or 'bg_bg' to the ISO 639 subtag 'bg'.
    # Media players (Plex, Jellyfin, Kodi) match subtitle languages on that subtag, so the
    # region part is dropped to keep the resulting file name recognizable to them.
    param(
        [AllowNull()][string]$Language
    )

    if ([string]::IsNullOrWhiteSpace($Language)) {
        return $null
    }

    $tag = $Language.Trim().Replace('_', '-').ToLowerInvariant().Split('-')[0]

    if ($tag -notmatch '^[a-z]{2,3}$') {
        return $null
    }

    return $tag
}

function Add-CaptionLanguageToFileName {
    # Inserts the language subtag before the extension, eg 'Session.vtt' becomes 'Session.bg.vtt'.
    param(
        [parameter(Mandatory = $true)][string]$Path,
        [AllowNull()][string]$LanguageTag
    )

    if ([string]::IsNullOrWhiteSpace($LanguageTag)) {
        return $Path
    }

    $extension = [System.IO.Path]::GetExtension($Path)
    $withoutExtension = $Path.Substring(0, $Path.Length - $extension.Length)

    if ($withoutExtension -match ('(?i)\.{0}$' -f [regex]::Escape($LanguageTag))) {
        return $Path
    }

    return '{0}.{1}{2}' -f $withoutExtension, $LanguageTag, $extension
}

function Get-ExistingCaptionFile {
    # Returns a previously downloaded caption for this session, taking into account that it
    # may carry a language subtag which is unknown until the caption source is resolved.
    param(
        [parameter(Mandatory = $true)][string]$Path
    )

    if (Test-Path -LiteralPath $Path) {
        return $Path
    }

    $directory = [System.IO.Path]::GetDirectoryName($Path)
    if ([string]::IsNullOrWhiteSpace($directory) -or -not (Test-Path -LiteralPath $directory)) {
        return $null
    }

    $extension = [System.IO.Path]::GetExtension($Path)
    $fileName = [System.IO.Path]::GetFileName($Path)
    $baseName = $fileName.Substring(0, $fileName.Length - $extension.Length)

    # -Filter is used rather than -Path so that characters like [ ] in session titles are
    # not interpreted as PowerShell wildcards.
    $existing = @(Get-ChildItem -LiteralPath $directory -Filter ('{0}.*{1}' -f $baseName, $extension) -File -ErrorAction SilentlyContinue)
    if ($existing.Count -gt 0) {
        return $existing[0].FullName
    }

    return $null
}

function Remove-EmptySessionFolder {
    # Removes a per-session folder that ended up without any content. Downloads are queued as
    # background jobs, so a folder can still look empty while a job is about to write into it;
    # those folders are kept so the download does not lose its target directory.
    param(
        [parameter(Mandatory = $true)][string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return
    }

    if (@(Get-ChildItem -LiteralPath $Path -Force -ErrorAction SilentlyContinue).Count -gt 0) {
        return
    }

    foreach ($pendingJob in $script:BackgroundDownloadJobs) {
        $pendingFile = [string]$pendingJob.file
        if ([string]::IsNullOrWhiteSpace($pendingFile)) {
            continue
        }
        if ([System.IO.Path]::GetDirectoryName($pendingFile) -eq $Path) {
            Write-Verbose ('Keeping session folder {0}: a download job still targets it' -f $Path)
            return
        }
    }

    Write-Verbose ('Removing empty session folder {0}' -f $Path)
    Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
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

    if ([System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT) {
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
    param([switch]$SuppressShow)
    $Temp = @()
    foreach ( $job in $script:BackgroundDownloadJobs) {

        switch ( $job.Type) {
            1 {
                $isJobRunning = $job.job.State -in 'NotStarted', 'Running', 'Blocked'
            }
            2 {
                $isJobRunning = -not $job.job.hasExited
            }
            3 {
                $isJobRunning = $job.job.State -in 'NotStarted', 'Running', 'Blocked'
            }
            default {
                $isJobRunning = $false
            }
        }
        if ( $isJobRunning) {
            $Temp += $job
        }
        else {
            if ($job.cookieTempFile -and (Test-Path -LiteralPath $job.cookieTempFile)) {
                Remove-Item -LiteralPath $job.cookieTempFile -Force -ErrorAction Stop
            }
            # Job finished, process result
            switch ( $job.Type) {
                1 {
                    $isJobSuccess = $job.job.State -eq 'Completed' -and (Test-Path -LiteralPath $job.file) -and (Get-Item -LiteralPath $job.file).Length -gt 0
                    Write-Progress -Id ($job.job.Id + 1000) -Activity ('Slidedeck {0} {1}' -f $Job.scheduleCode, $Job.title) -Completed
                }
                2 {
                    $isJobSuccess = Test-Path -LiteralPath $job.file
                    if ($job.cookieTempFile) { $isJobSuccess = $isJobSuccess -and $job.job.ExitCode -eq 0 }
                    Write-Progress -Id $job.job.Id -Activity ('Video {0} {1}' -f $Job.scheduleCode, $Job.title) -Completed
                }
                3 {
                    $isJobSuccess = Test-Path -LiteralPath $job.file
                    if ($script:CustomCookieFile) { $isJobSuccess = $isJobSuccess -and $job.job.State -eq 'Completed' }
                    Write-Progress -Id ($job.job.Id + 2000) -Activity ('Captions {0} {1}' -f $Job.scheduleCode, $Job.title) -Completed
                }
                default {
                    $isJobSuccess = $false
                }
            }

            Write-Host ('Testing placeholders against {0} {1}' -f $job.scheduleCode, $job.file)
            # Test if file is placeholder
            $isPlaceholder = $false
            if ( Test-Path -LiteralPath $job.file) {
                $FileObj = Get-ChildItem -LiteralPath $job.file
                if ( $FileObj.Length -lt 1kb) {

                    if ( @('No resource file is available for download', 'No resource file is available for download.', 'No resource file is available for download for the given id') -contains (Get-Content -LiteralPath $job.File) ) {
                        Write-Warning ('Removing {0} placeholder file {1}' -f $job.scheduleCode, $job.file)
                        Remove-Item -LiteralPath $job.file -Force
                        $isPlaceholder = $true

                        switch ( $job.Type) {
                            1 {
                                # Placeholder Deck file downloaded
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
                        Write-Warning ('File {0} for {1} {2} is smaller than 1KB but does not contain expected placeholder text; leaving in place for manual review' -f $job.file, $job.scheduleCode, $job.title)
                    }
                }
            }

            if ( $isJobSuccess -and -not $isPlaceholder) {

                if ($job.Type -eq 1) {
                    $DeckInfo[$InfoDownload]++
                }
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
                    if ($job.metadata) {
                        Set-VideoMetadata -File $job.file -Metadata $job.metadata
                    }

                    # Clean video leftovers
                    Clean-VideoLeftovers $job.file
                }
            }
            else {
                if ($script:CustomCookieFile) {
                    Write-Warning ('Cookie-backed download failed for session {0}. Check media access or cookie expiry; credentials and response details withheld.' -f $job.scheduleCode)
                }
                switch ( $job.Type) {
                    1 {
                        if (-not $isPlaceholder) { $DeckInfo[$InfoPlaceholder]++ }
                        Write-Host ('Problem downloading or missing slidedeck of {0} {1}' -f $job.scheduleCode, $job.title) -ForegroundColor Red
                        $job.job.ChildJobs | Stop-Job | Out-Null
                        $job.job | Stop-Job -PassThru | Remove-Job -Force | Out-Null
                    }
                    2 {
                        $LastLine = if ($job.cookieTempFile) { 'Cookie-backed downloader details withheld.' } else { (Get-Content -LiteralPath $job.stdErrTempFile -ErrorAction SilentlyContinue) | Select-Object -Last 1 }
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
    if (-not $SuppressShow) { Show-BackgroundDownloadJobs }
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
        if ( $Job.Type -eq 1) {
            Write-Progress -Id ($job.job.id + 1000) -Activity ('Slidedeck {0} {1}' -f $job.scheduleCode, $Job.title) -Status 'Downloading...' -ParentId 2
        }
        if ( $Job.Type -eq 3) {
            $captionFileInfo = Get-Item -LiteralPath $job.file -ErrorAction SilentlyContinue
            $captionDownloaded = if ($captionFileInfo) { $captionFileInfo.Length } else { 0 }
            if ($job.totalBytes -gt 0) {
                $captionPct = [int][Math]::Min(99, [Math]::Round($captionDownloaded / $job.totalBytes * 100))
                $captionStatus = 'Downloaded: {0:F1} KB of {1:F1} KB' -f ($captionDownloaded / 1KB), ($job.totalBytes / 1KB)
                Write-Progress -Id ($job.job.id + 2000) -Activity ('Caption {0} {1}' -f $job.scheduleCode, $Job.title) -Status $captionStatus -PercentComplete $captionPct -ParentId 2
            }
            else {
                $captionStatus = if ($captionDownloaded -gt 0) { 'Downloaded: {0:F1} KB' -f ($captionDownloaded / 1KB) } else { 'Starting...' }
                Write-Progress -Id ($job.job.id + 2000) -Activity ('Caption {0} {1}' -f $job.scheduleCode, $Job.title) -Status $captionStatus -ParentId 2
            }
        }
        if ( $Job.Type -eq 2) {

            # Get last line of yt-dlp log; cache by file mtime to avoid re-reading unchanged files
            $jobKey = [string]$job.job.id
            $fileInfo = Get-Item -LiteralPath $job.stdOutTempFile -ErrorAction SilentlyContinue
            if ($fileInfo -and $script:JobFileLastWrite[$jobKey] -ne $fileInfo.LastWriteTime) {
                $script:JobFileLastWrite[$jobKey] = $fileInfo.LastWriteTime
                $line = (Get-Content -LiteralPath $job.stdOutTempFile -ErrorAction SilentlyContinue) | Select-Object -Last 1
                if ($line) { $script:JobProgressCache[$jobKey] = $line }
            }
            $LastLine = if ($script:JobProgressCache[$jobKey]) { $script:JobProgressCache[$jobKey] } else { 'Evaluating..' }
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
                if ($BGJob.cookieTempFile) {
                    if (-not $BGJob.job.HasExited) {
                        throw (New-CustomCookieFileException -Message 'The video downloader did not stop. Its temporary cookie file is retained until the process can be stopped safely.')
                    }
                    if (Test-Path -LiteralPath $BGJob.cookieTempFile) { Remove-Item -LiteralPath $BGJob.cookieTempFile -Force -ErrorAction Stop }
                }
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

function Get-MSADownloadAuthHeaders {
    # Returns a headers hashtable (Authorization and/or Cookie) suitable for passing to
    # Add-BackgroundDownloadJob for authenticated downloads from Microsoft content URLs.
    # Only acquires an MSA session (interactively when needed) after content requests have
    # encountered an authentication challenge; anonymously accessible events stay anonymous.
    param(
        [parameter(Mandatory = $true)][string]$Url,
        [uri]$Proxy
    )

    $headers = @{}

    if ($script:CustomCookieFile) {
        $session = Get-CustomCookieWebSession -TargetUri ([uri]$Url)
        $cookieHeader = $session.Cookies.GetCookieHeader([uri]$Url)
        if ($cookieHeader) { return @{ Cookie = $cookieHeader } }
        return $null
    }

    # No authentication challenge encountered so far; content is anonymously accessible.
    if (-not $script:MSAContentAuthRequired) {
        return $null
    }

    # Authentication is required but we have no credentials yet; acquire them now (may prompt interactively).
    if (-not $script:MSAAuthWebSession -and [string]::IsNullOrWhiteSpace($script:MSABearerToken)) {
        try {
            [uri]$contentUri = $Url
            $acquiredSession = Get-MSAAuthenticatedWebSession -TargetUri $contentUri -Proxy $Proxy -ValidateCachedSession
            if ($acquiredSession) {
                $script:MSAAuthWebSession = $acquiredSession
            }
        }
        catch {
            Write-Warning ('Unable to acquire MSA authentication for background download: {0}' -f $_.Exception.Message)
        }
    }

    # Add Bearer token as Authorization header when valid.
    if (-not [string]::IsNullOrWhiteSpace($script:MSABearerToken) -and
        (Test-JwtTokenValid -Token $script:MSABearerToken -MinValidityMinutes $script:MinTokenValidityMinutes)) {
        $headers['Authorization'] = ('Bearer {0}' -f $script:MSABearerToken)
    }

    # Extract cookies from the web session as a plain string so the value survives job serialization.
    if ($script:MSAAuthWebSession) {
        try {
            [uri]$contentUri = $Url
            $cookieHeader = Get-WebSessionCookieHeader -WebSession $script:MSAAuthWebSession -CookieUri $contentUri -IncludeAllDomains
            if (-not [string]::IsNullOrWhiteSpace($cookieHeader)) {
                $headers['Cookie'] = $cookieHeader
            }
        }
        catch {
            Write-Verbose ('Unable to extract cookies for background download auth headers: {0}' -f $_.Exception.Message)
        }
    }

    if ($headers.Count -eq 0) {
        return $null
    }

    return $headers
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
        $ScheduleCode = '',
        $Metadata = $null,
        [hashtable]$Headers = $null,
        [uri]$Proxy = $null
    )
    $JobsRunning = Get-BackgroundDownloadJobs -SuppressShow
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
    $cookieTempFile = $null
    switch ( $Type) {
        1 {
            # Slidedeck
            if ($script:CustomCookieFile) {
                $job = Start-CustomCookieDownloadJob -Url $DownloadUrl -File $FilePath -Proxy $Proxy
            }
            elseif ($Headers) {
                $job = Start-Job -ScriptBlock {
                    param($url, $file, $headers, $proxy)
                    $ProgressPreference = 'SilentlyContinue'
                    $invokeParams = @{
                        Uri         = $url
                        Method      = 'Get'
                        OutFile     = $file
                        ErrorAction = 'Stop'
                        Verbose     = $false
                    }
                    if ($headers) { $invokeParams.Headers = $headers }
                    if ($proxy) { $invokeParams.Proxy = $proxy }
                    Invoke-WebRequest @invokeParams | Out-Null
                } -ArgumentList $DownloadUrl, $FilePath, $Headers, $Proxy
            }
            else {
                $job = Start-Job -ScriptBlock {
                    param( $url, $file)
                    $wc = New-Object System.Net.WebClient
                    $wc.Encoding = [System.Text.Encoding]::UTF8
                    $wc.DownloadFile( $url, $file)
                } -ArgumentList $DownloadUrl, $FilePath
            }
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
            if ($script:CustomCookieFile) {
                $session = Get-CustomCookieWebSession -TargetUri ([uri]$EventAPIUrl)
                $cookieTempFile = New-CustomYtDlpCookieFile -WebSession $session
                $ProcessParam.ArgumentList = @($ArgumentList) + ('--cookies {0}' -f (ConvertTo-ArgumentString -Value $cookieTempFile))
            }
            try {
                $job = if ($cookieTempFile) { Start-Process @ProcessParam -ErrorAction Stop } else { Start-Process @ProcessParam }
            }
            catch {
                if ($cookieTempFile) {
                    Remove-Item -LiteralPath $cookieTempFile -Force -ErrorAction Stop
                    throw (New-CustomCookieFileException -Message 'Unable to start the video downloader; credentials and process arguments withheld.')
                }
                throw
            }
        }
        3 {
            # Caption
            $totalBytes = $null
            try {
                $headParams = @{ Uri = $DownloadUrl; Method = 'Head'; ErrorAction = 'Stop'; Verbose = $false }
                if ($script:CustomCookieFile) { $headParams.WebSession = Get-CustomCookieWebSession -TargetUri ([uri]$DownloadUrl) }
                elseif ($Headers) { $headParams.Headers = $Headers }
                if ($Proxy) { $headParams.Proxy = $Proxy }
                $savedPP = $ProgressPreference ; $ProgressPreference = 'SilentlyContinue'
                try { $headResponse = Invoke-WebRequest @headParams } finally { $ProgressPreference = $savedPP }
                if ($headResponse -and $headResponse.Headers['Content-Length']) {
                    $totalBytes = [long]$headResponse.Headers['Content-Length']
                }
            }
            catch { }
            if ($script:CustomCookieFile) {
                $job = Start-CustomCookieDownloadJob -Url $DownloadUrl -File $FilePath -Proxy $Proxy
            }
            elseif ($Headers) {
                $job = Start-Job -ScriptBlock {
                    param($url, $file, $headers, $proxy)
                    $ProgressPreference = 'SilentlyContinue'
                    $invokeParams = @{
                        Uri         = $url
                        Method      = 'Get'
                        OutFile     = $file
                        ErrorAction = 'Stop'
                        Verbose     = $false
                    }
                    if ($headers) { $invokeParams.Headers = $headers }
                    if ($proxy) { $invokeParams.Proxy = $proxy }
                    Invoke-WebRequest @invokeParams | Out-Null
                } -ArgumentList $DownloadUrl, $FilePath, $Headers, $Proxy
            }
            else {
                $job = Start-Job -ScriptBlock {
                    param( $url, $file)
                    $wc = New-Object System.Net.WebClient
                    $wc.Encoding = [System.Text.Encoding]::UTF8
                    $wc.DownloadFile( $url, $file)
                } -ArgumentList $DownloadUrl, $FilePath
            }
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
        metadata       = $Metadata
        stdOutTempFile = $stdOutTempFile
        stdErrTempFile = $stdErrTempFile
        totalBytes     = $totalBytes
        cookieTempFile = $cookieTempFile
    }

    $script:BackgroundDownloadJobs += $object
    Show-BackgroundDownloadJobs
}

##########
# MAIN
##########

if ($Event -eq 'Custom' -and $CookiesFile) {
    $script:CustomCookieFile = (Resolve-Path -LiteralPath $CookiesFile -ErrorAction Stop).ProviderPath
    if ($CookiesFromBrowser) {
        Write-Warning 'CookiesFile takes precedence for Custom events; CookiesFromBrowser is ignored.'
        $CookiesFromBrowser = $null
    }
}

try {
Write-Host( '*' * 78)
Write-Host( 'Get-EventSession v4.55')
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
    { 'Ignite', 'Ignite2026' -contains $_ } {
        $EventName = 'Ignite2026'
        $EventType = 'API2'
        $EventAPIUrl = 'https://eventtools.event.microsoft.com/ignite2026-prod/fallback/session-all-en-us.json'
        $SessionUrl = 'https://medius.microsoft.com/video/asset/HIGHMP4/{0}'
        $CaptionURL = 'https://medius.microsoft.com/video/asset/CAPTION/{0}'
        $SlidedeckUrl = 'https://medius.microsoft.com/video/asset/PPT/{0}'
        $Method = 'Get'
        $CaptionExt = 'vtt'
        $PreferDirect = $True
    }
    { 'Ignite2025' -contains $_ } {
        $EventName = 'Ignite2025'
        $EventType = 'API2'
        $EventAPIUrl = 'https://eventtools.event.microsoft.com/ignite2025-prod/fallback/session-all-en-us.json'
        $SessionUrl = 'https://medius.microsoft.com/video/asset/HIGHMP4/{0}'
        $CaptionURL = 'https://medius.microsoft.com/video/asset/CAPTION/{0}'
        $SlidedeckUrl = 'https://medius.microsoft.com/video/asset/PPT/{0}'
        $Method = 'Get'
        $CaptionExt = 'vtt'
        $PreferDirect = $True
    }
    { 'Ignite2024', 'Ignite2023', 'Ignite2022', 'Ignite2021H1', 'Ignite2021H2', 'Ignite2020' -contains $_ } {
        $EventName = $Event
        $EventType = 'Archive'
        $EventAPIUrl = 'https://medius.microsoft.com'
        switch ( $Event) {
            'Ignite2024' {
                $EventSearchURI = '/Home/VideoSearchSessionArchive?searchString=&page={{0}}&embededSearchFilterExpression=((Channels+eq+%27Microsoft+Ignite+2024%27))&orderBy=VideoTitle&pageSize=100&pageSizes=100%2C24%2C48%2C96&searchgridId=ee8a4eea-d554-4186-859f-8c733020023d&loc=en'
            }
            'Ignite2023' {
                $EventSearchURI = '/Home/VideoSearchSessionArchive?searchString=&page={{0}}&embededSearchFilterExpression=((Channels+eq+%27Microsoft+Ignite+2023%27))&orderBy=VideoTitle&pageSize=100&pageSizes=100%2C24%2C48%2C96&searchgridId=ee8a4eea-d554-4186-859f-8c733020023d&loc=en'
            }
            'Ignite2022' {
                $EventSearchURI = '/Home/VideoSearchSessionArchive?searchString=&page={{0}}&embededSearchFilterExpression=((Channels+eq+%27Microsoft+Ignite+2022%27))&orderBy=VideoTitle&pageSize=100&pageSizes=100%2C24%2C48%2C96&searchgridId=ee8a4eea-d554-4186-859f-8c733020023d&loc=en'
            }
            'Ignite2021H1' {
                $EventSearchURI = '/Home/VideoSearchSessionArchive?searchString=&page={{0}}&embededSearchFilterExpression=((Channels+eq+%27Microsoft+Ignite+March+2021%27))&orderBy=VideoTitle&pageSize=100&pageSizes=100%2C24%2C48%2C96&searchgridId=ee8a4eea-d554-4186-859f-8c733020023d&loc=en'
            }
            'Ignite2021H2' {
                $EventSearchURI = '/Home/VideoSearchSessionArchive?searchString=&page={{0}}&embededSearchFilterExpression=((Channels+eq+%27Microsoft+Ignite+Falls+-+November+2021%27))&orderBy=VideoTitle&pageSize=100&pageSizes=100%2C24%2C48%2C96&searchgridId=ee8a4eea-d554-4186-859f-8c733020023d&loc=en'
            }
            'Ignite2020' {
                $EventSearchURI = '/Home/VideoSearchSessionArchive?searchString=&page={{0}}&embededSearchFilterExpression=((Channels+eq+%27Microsoft+Ignite+2020%27))&orderBy=VideoTitle&pageSize=100&pageSizes=100%2C24%2C48%2C96&searchgridId=ee8a4eea-d554-4186-859f-8c733020023d&loc=en'
            }
            default {
                throw( 'Error (Unsupported Ignite event: {0})' -f $Event)
                exit -1
            }
        }
        $SessionUrl = 'https://medius.microsoft.com/video/asset/HIGHMP4/{0}'
        $CaptionURL = 'https://medius.microsoft.com/video/asset/CAPTION/{0}'
        $SlidedeckUrl = 'https://medius.microsoft.com/video/asset/PPT/{0}'
        $Method = 'Get'
        $CaptionExt = 'vtt'
        $PreferDirect = $True
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
    { 'Build', 'Build2026' -contains $_ } {
        $EventName = 'Build2026'
        $EventType = 'API2'
        $EventAPIUrl = 'https://eventtools.event.microsoft.com/build2026-prod/fallback/session-all-en-us.json'
        $SessionUrl = 'https://medius.microsoft.com/video/asset/HIGHMP4/{0}'
        $CaptionURL = 'https://medius.microsoft.com/video/asset/CAPTION/{0}'
        $SlidedeckUrl = 'https://medius.microsoft.com/video/asset/PPT/{0}'
        $Method = 'Get'
        $CaptionExt = 'vtt'
        $PreferDirect = $True
    }
    { 'Build2025' -contains $_ } {
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

if ($script:CustomCookieFile -and -not (Test-IsCustomCatalogUri -Uri ([uri]$EventAPIUrl))) {
    throw 'Custom event catalogs require an absolute HTTPS URL without embedded credentials.'
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
            try {
                Invoke-WebRequest -Uri $YTLink -OutFile $YouTubeDL -Proxy $ProxyURL -ErrorAction Stop
            }
            catch {
                Write-Warning ('Failed to download {0}: {1}' -f $YouTubeEXE, $_.Exception.Message)
            }
        }
        if ( Test-Path $YouTubeDL) {
            Write-Host ('Running self-update of {0}' -f $YouTubeEXE)

            $Arg = @('-U')
            if ( $ProxyURL) { $Arg += ('--proxy "{0}"' -f $ProxyURL) }

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
                if (-not $p.WaitForExit(300000)) { $p.Kill(); Write-Warning ('{0} self-update timed out after 5 minutes and was terminated.' -f $YouTubeEXE) }
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
            try {
                Invoke-WebRequest -Uri $FFMPEGlink -OutFile $TempFile -Proxy $ProxyURL -ErrorAction Stop
            }
            catch {
                Write-Warning ('Failed to download ffmpeg: {0}' -f $_.Exception.Message)
            }

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
                $data = Import-Clixml -LiteralPath $SessionCache -ErrorAction Stop
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
        'Archive' {
            Write-Host ('Reading {0} archived session catalog' -f $EventName)
            $web = @{
                userAgent  = 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.103 Safari/537.36'
                requestUri = $null
            }
            $data = [System.Collections.ArrayList]@()
            $seenSessionKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
            $searchUriTemplate = ($EventSearchURI -replace '\{\{0\}\}', '{0}')
            $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet', [string[]]('sessionCode', 'title'))
            $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)

            $page = 1
            while ($true) {
                Write-Progress -Id 1 -Activity 'Retrieving Session Catalog' -Status ('Processing page {0}' -f $page)

                $web.requestUri = [uri]('{0}/{1}' -f $EventAPIUrl, ($searchUriTemplate -f $page))
                try {
                    Write-Verbose ('Using URI {0}' -f $web.requestUri)
                    $searchResultsResponse = Invoke-WebWithRetry -ScriptBlock { Invoke-WebRequest -Uri $web.requestUri -Method Get -UserAgent $web.userAgent -WebSession $session -Proxy $ProxyURL } -Variables @{ web = $web; session = $session; ProxyURL = $ProxyURL }
                }
                catch {
                    if ($_.Exception -is [System.Management.Automation.PipelineStoppedException]) { throw }
                    throw ('Problem retrieving archived session catalog page {0}: {1}' -f $page, $error[0])
                }

                $html = [string]$searchResultsResponse.Content
                if ([string]::IsNullOrWhiteSpace($html)) {
                    break
                }

                $cardMatches = [regex]::Matches($html, '(?is)<li class="col mb-4 cards-session">(?<card>.*?)</li>')
                if ($cardMatches.Count -eq 0) {
                    break
                }

                $addedThisPage = 0
                foreach ($cardMatch in $cardMatches) {
                    $cardHtml = [string]$cardMatch.Groups['card'].Value

                    $shortCode = [regex]::Match($cardHtml, 'data-shortcode="(?<v>[^"]+)"', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase).Groups['v'].Value
                    if ([string]::IsNullOrWhiteSpace($shortCode)) {
                        $shortCode = [regex]::Match($cardHtml, 'data-shortCode="(?<v>[^"]+)"', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase).Groups['v'].Value
                    }
                    if ([string]::IsNullOrWhiteSpace($shortCode)) {
                        continue
                    }
                    if (-not $seenSessionKeys.Add($shortCode)) {
                        continue
                    }

                    $sessionCode = $shortCode
                    if ($shortCode -match '^[^-]+-(?<code>.+)$') {
                        $sessionCode = $Matches.code
                    }

                    # Archive cards also contain non-session videos where the shortcode tail is an internal media id
                    # (for example he*/ct* hex-like values). Keep only schedule-like session codes.
                    $isScheduleStyleCode = ($sessionCode -cmatch '^[A-Z][A-Za-z0-9-]*$') -or ($sessionCode -cmatch '^ct\d{9}$')
                    if (-not $isScheduleStyleCode) {
                        Write-Debug ('Skipping non-session archive card with code token {0}' -f $sessionCode)
                        continue
                    }

                    $sessionTitleRaw = [regex]::Match($cardHtml, '(?is)<a class="text-black text-dec-default openvideo"[^>]*>(?<v>.*?)</a>').Groups['v'].Value
                    if ([string]::IsNullOrWhiteSpace($sessionTitleRaw)) {
                        $sessionTitleRaw = [regex]::Match($cardHtml, 'aria-label="(?<v>[^"]+)"').Groups['v'].Value
                    }
                    $sessionTitle = [string][System.Net.WebUtility]::HtmlDecode(($sessionTitleRaw -replace '<[^>]+>', '').Trim())

                    $channelRaw = [regex]::Match($cardHtml, '(?is)<p class="keynote-text[^"]*">(?<v>.*?)</p>').Groups['v'].Value
                    $channel = [string][System.Net.WebUtility]::HtmlDecode(($channelRaw -replace '<[^>]+>', '').Trim())

                    $duration = [System.Net.WebUtility]::HtmlDecode(([regex]::Match($cardHtml, '(?is)<p class="glyph-prepend glyph-prepend-circle-fill durationmin[^"]*">(?<v>.*?)</p>').Groups['v'].Value -replace '<[^>]+>', '').Trim())
                    $thumbnail = [System.Net.WebUtility]::HtmlDecode([regex]::Match($cardHtml, '(?is)<img\s+[^>]*class="[^"]*thumbnail-grid[^"]*"[^>]*\ssrc="(?<v>[^"]+)"').Groups['v'].Value)

                    $startDateTime = $null
                    $endDateTime = $null
                    $publishDateRaw = [regex]::Match($cardHtml, 'localPublishDateTime\s*=\s*new\s+Date\("(?<v>[^"]+)"\)', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase).Groups['v'].Value
                    if (-not [string]::IsNullOrWhiteSpace($publishDateRaw) -and $publishDateRaw -notmatch '^0001-01-01\s') {
                        try {
                            $startDateTime = [datetime]::Parse($publishDateRaw, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeUniversal)
                        }
                        catch {
                            Write-Verbose ('Could not parse archive publish date "{0}" for session {1}' -f $publishDateRaw, $sessionCode)
                        }
                    }

                    $durationTimeSpan = $null
                    $durationInMinutes = $null
                    $durationHourMinuteMatch = [regex]::Match($duration, '^(?<h>\d{1,2}):(?<m>\d{2})(?::(?<s>\d{2}))?\s*min$', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                    if ($durationHourMinuteMatch.Success) {
                        $hours = [int]$durationHourMinuteMatch.Groups['h'].Value
                        $minutes = [int]$durationHourMinuteMatch.Groups['m'].Value
                        $seconds = if ($durationHourMinuteMatch.Groups['s'].Success) { [int]$durationHourMinuteMatch.Groups['s'].Value } else { 0 }
                        try {
                            $durationTimeSpan = New-TimeSpan -Hours $hours -Minutes $minutes -Seconds $seconds
                        }
                        catch {
                            $durationTimeSpan = $null
                        }
                        if ($durationTimeSpan) {
                            $durationInMinutes = [int][math]::Round($durationTimeSpan.TotalMinutes, 0)
                        }
                    }
                    else {
                        $durationMinuteMatch = [regex]::Match($duration, '^(?<m>\d{1,4})\s*min$', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                        if ($durationMinuteMatch.Success) {
                            try {
                                $durationInMinutes = [int]$durationMinuteMatch.Groups['m'].Value
                                $durationTimeSpan = New-TimeSpan -Minutes $durationInMinutes
                            }
                            catch {
                                $durationTimeSpan = $null
                                $durationInMinutes = $null
                            }
                        }
                    }

                    if (-not $startDateTime) {
                        # Keep a deterministic non-future placeholder datetime when the card has no publish date.
                        $startDateTime = [datetime]'1900-01-01 09:00:00'
                    }

                    if ($startDateTime -and $durationTimeSpan) {
                        $endDateTime = $startDateTime.Add($durationTimeSpan)
                    }

                    $speakers = [System.Collections.ArrayList]@()
                    foreach ($speakerMatch in [regex]::Matches($cardHtml, '(?is)<(p|span|a)[^>]*class="[^"]*speaker[^"]*"[^>]*>(?<v>.*?)</(p|span|a)>')) {
                        $speakerText = [System.Net.WebUtility]::HtmlDecode(($speakerMatch.Groups['v'].Value -replace '<[^>]+>', '').Trim())
                        if ($speakerText -and $speakerText -notmatch '^(Speaker|Speakers)$') {
                            foreach ($speakerName in ($speakerText -split ',|;')) {
                                $candidate = $speakerName.Trim()
                                if ($candidate) {
                                    $speakers.Add($candidate) | Out-Null
                                }
                            }
                        }
                    }
                    $speakerNames = @($speakers | Select-Object -Unique)

                    $otherAssets = @()
                    $otherLinksRaw = [regex]::Match($cardHtml, '(?is)let\s+BoxViewOtherLinks\s*=\s*"(?<v>.*?)";').Groups['v'].Value
                    if (-not [string]::IsNullOrWhiteSpace($otherLinksRaw)) {
                        try {
                            $otherAssets = @(([regex]::Unescape($otherLinksRaw) | ConvertFrom-Json))
                        }
                        catch {
                            Write-Verbose ('Could not parse BoxViewOtherLinks for session {0}: {1}' -f $sessionCode, $_.Exception.Message)
                        }
                    }

                    $captionAssets = @()
                    $captionLinksRaw = [regex]::Match($cardHtml, '(?is)let\s+BoxViewCaptionUrl\s*=\s*"(?<v>.*?)";').Groups['v'].Value
                    if (-not [string]::IsNullOrWhiteSpace($captionLinksRaw)) {
                        try {
                            $captionAssets = @(([regex]::Unescape($captionLinksRaw) | ConvertFrom-Json))
                        }
                        catch {
                            Write-Verbose ('Could not parse BoxViewCaptionUrl for session {0}: {1}' -f $sessionCode, $_.Exception.Message)
                        }
                    }

                    $videoLink = ($otherAssets | Where-Object { $_.AssetType -eq 'HighQualityUrl' } | Select-Object -First 1).LinkUrl
                    $slideDeckLink = ($otherAssets | Where-Object { $_.AssetType -eq 'PPTUrl' } | Select-Object -First 1).LinkUrl
                    $captionLink = ($otherAssets | Where-Object { $_.AssetType -eq 'CaptionUrl' } | Select-Object -First 1).LinkUrl
                    $captionLanguage = $null
                    if (-not $captionLink) {
                        $preferredCaptionAsset = ($captionAssets | Where-Object { $_.FileName -match '(?i)Caption_(?<lang>[A-Za-z]{2}[-_][A-Za-z]{2})\.vtt$' -and $Matches.lang -match '^(?i)(en-us|en-gb)$' } | Select-Object -First 1)
                        if ($preferredCaptionAsset) {
                            $captionLink = $preferredCaptionAsset.StreamUrl
                            if ($preferredCaptionAsset.FileName -match '(?i)Caption_(?<lang>[A-Za-z]{2}[-_][A-Za-z]{2})\.vtt$') {
                                $captionLanguage = ($Matches.lang -replace '_', '-').ToLowerInvariant()
                            }
                        }
                    }
                    if (-not $captionLink) {
                        $fallbackCaptionAsset = ($captionAssets | Select-Object -First 1)
                        if ($fallbackCaptionAsset) {
                            $captionLink = $fallbackCaptionAsset.StreamUrl
                            if ($fallbackCaptionAsset.FileName -match '(?i)Caption_(?<lang>[A-Za-z]{2}[-_][A-Za-z]{2})\.vtt$') {
                                $captionLanguage = ($Matches.lang -replace '_', '-').ToLowerInvariant()
                            }
                        }
                    }
                    if (-not $captionLanguage -and $captionLink) {
                        $captionAssetByUrl = ($captionAssets | Where-Object { $_.StreamUrl -eq $captionLink } | Select-Object -First 1)
                        if ($captionAssetByUrl -and $captionAssetByUrl.FileName -match '(?i)Caption_(?<lang>[A-Za-z]{2}[-_][A-Za-z]{2})\.vtt$') {
                            $captionLanguage = ($Matches.lang -replace '_', '-').ToLowerInvariant()
                        }
                    }
                    if (-not $captionLanguage -and $captionLink -match '(?i)Caption_(?<lang>[A-Za-z]{2}[-_][A-Za-z]{2})\.vtt') {
                        $captionLanguage = ($Matches.lang -replace '_', '-').ToLowerInvariant()
                    }

                    $item = [PSCustomObject]@{
                        sessionCode        = $sessionCode
                        scheduleCode       = $sessionCode
                        code               = $sessionCode
                        title              = $sessionTitle
                        description        = $null
                        speakerNames       = $speakerNames
                        products           = @($channel)
                        contentCategory    = @($channel)
                        langLocale         = 'en-US'
                        startDateTime      = $startDateTime
                        endDateTime        = $endDateTime
                        durationInMinutes  = $durationInMinutes
                        onDemand           = $null
                        downloadVideoLink  = $videoLink
                        onDemandThumbnail  = $thumbnail
                        slideDeck          = $slideDeckLink
                        captionFileLink    = $captionLink
                        transcriptFileLink = $captionLink
                        captionLanguage    = $captionLanguage
                    }

                    Write-Verbose ('Adding info for session {0}' -f $item.sessionCode)
                    $item.PSObject.TypeNames.Insert(0, 'Session.Information')
                    $item | Add-Member MemberSet PSStandardMembers $PSStandardMembers -Force
                    $data.Add($item) | Out-Null
                    $addedThisPage++
                }

                if ($addedThisPage -eq 0) {
                    break
                }

                $page++
                if ($page -gt 1000) {
                    Write-Warning 'Reached pagination safety limit while scraping archive catalog; stopping at page 1000.'
                    break
                }
            }

            [int32]$sessionCount = ($data | Measure-Object).Count
            Write-Host ('Processing information for {0} sessions' -f $sessionCount)
            Write-Progress -Id 1 -Completed -Activity "Finished retrieval of catalog"
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
                $ResultsResponse = Invoke-WebWithRetry -ScriptBlock { Invoke-JsonRestMethod -Uri $web.requestUri -Method $Method -Headers $web.headers -UserAgent $web.userAgent -WebSession $session -Proxy $ProxyURL -TimeoutSec $web.Timeout } -Variables @{ web = $web; Method = $Method; session = $session; ProxyURL = $ProxyURL }
            }
            catch {
                if ($_.Exception -is [System.Management.Automation.PipelineStoppedException]) { throw }
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
                $searchResultsResponse = Invoke-WebWithRetry -ScriptBlock { Invoke-JsonRestMethod -Uri $web.requestUri -Body $searchbody -Method $Method -Headers $web.headers -UserAgent $web.userAgent -WebSession $session -Proxy $ProxyURL } -Variables @{ web = $web; searchbody = $searchbody; Method = $Method; session = $session; ProxyURL = $ProxyURL }
            }
            catch {
                if ($_.Exception -is [System.Management.Automation.PipelineStoppedException]) { throw }
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
                $searchResultsResponse = Invoke-WebWithRetry -ScriptBlock { Invoke-JsonRestMethod -Uri $web.requestUri -Body $searchbody -Method $Method -Headers $web.headers -UserAgent $web.userAgent -WebSession $session -Proxy $ProxyURL } -Variables @{ web = $web; searchbody = $searchbody; Method = $Method; session = $session; ProxyURL = $ProxyURL }
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
            try {
                $p.Start() | Out-Null
                $stdout = $p.StandardOutput.ReadToEnd()
                $stderr = $p.StandardError.ReadToEnd()
                if (-not $p.WaitForExit(300000)) { $p.Kill(); throw ('Timed out waiting for {0} to retrieve the YouTube playlist.' -f $YouTubeEXE) }
            }
            catch {
                throw ('Problem starting {0}: {1}' -f $YouTubeEXE, $_.Exception.Message)
            }

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

if ($Speaker -or $Product -or $Category -or $SolutionArea -or $LearningPath -or $Topic -or $ProgrammingLanguage -or $SessionLevel) {
    if ($Speaker) { Write-Verbose ('Speaker keyword specified: {0}' -f $Speaker) }
    if ($Product) { Write-Verbose ('Product specified: {0}' -f $Product) }
    if ($Category) { Write-Verbose ('Category specified: {0}' -f $Category) }
    if ($SolutionArea) { Write-Verbose ('SolutionArea specified: {0}' -f $SolutionArea) }
    if ($LearningPath) { Write-Verbose ('LearningPath specified: {0}' -f $LearningPath) }
    if ($Topic) { Write-Verbose ('Topic specified: {0}' -f $Topic) }
    if ($ProgrammingLanguage) { Write-Verbose ('Programming language(s) specified: {0}' -f ($ProgrammingLanguage -join ',')) }
    if ($SessionLevel) { Write-Verbose ('Session level(s) specified: {0}' -f ($SessionLevel -join ',')) }
    $SessionsToGet = $SessionsToGet | Where-Object {
        $s = $_
        (-not $Speaker -or ($s.speakerNames | Where-Object { $_ -ilike $Speaker })) -and
        (-not $Product -or ($s.products | Where-Object { $_ -ilike $Product })) -and
        (-not $Category -or ($s.contentCategory | Where-Object { $_ -ilike $Category })) -and
        (-not $SolutionArea -or ($s.solutionArea | Where-Object { $_ -ilike $SolutionArea })) -and
        (-not $LearningPath -or ($s.learningPath | Where-Object { $_ -ilike $LearningPath })) -and
        (-not $Topic -or ($s.topic | Where-Object { $_ -ilike $Topic })) -and
        (-not $ProgrammingLanguage -or (
            -not [string]::IsNullOrWhiteSpace($s.programmingLanguages) -and
            ($s.programmingLanguages -split ',' | Where-Object { $ProgrammingLanguage -icontains $_ })
        )) -and
        (-not $SessionLevel -or (
            -not [string]::IsNullOrWhiteSpace($s.sessionLevel) -and
            $s.sessionLevel -match '^\((\d+)\)' -and
            $SessionLevel -contains [int]$Matches[1]
        ))
    }
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
    $SessionsToGet = $SessionsToGet | Where-Object { $_.sessionCode -inotmatch 'R[1-9]?$' -and $_.sessionCode -inotmatch '^[A-Z]+[0-9]+[B-C]+$' }
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
    # Reset per-session state so a skipped session cannot act on the previous session's folder
    $ContentTargetFolder = $null
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
    elseif (-not (Test-SessionHasOnDemandContent -Session $SessionToGet) -and ($InfoOnly -or $NoSlidedecks)) {
        Write-Verbose ('Skipping session {0}: on-demand content is explicitly unavailable' -f $SessionToGet.sessionCode)
    }
    else {

        # When storing session content in subfolders per session, override the content target folder to be the session subfolder
        if (-not $InfoOnly -and $EventType -eq 'CUSTOM' -and ((-not $NoVideos -and (Test-SessionHasOnDemandContent -Session $SessionToGet) -and -not $SessionToGet.onDemand -and -not $SessionToGet.downloadVideoLink) -or -not $NoSlidedecks)) {
            Update-CustomSessionDetails -Session $SessionToGet -CatalogUri ([uri]$EventAPIUrl) -Proxy $ProxyURL
            if (-not (Test-SessionHasOnDemandContent -Session $SessionToGet) -and ($NoSlidedecks -or @(Get-SessionPresentationUrls -Session $SessionToGet).Count -eq 0)) {
                Write-Verbose ('Skipping session {0}: details indicate no on-demand content' -f $SessionToGet.sessionCode)
                continue
            }
        }
        if ( $UseSessionFolders -and -not $InfoOnly) {
            $SessionFolder = Join-Path -Path $DownloadFolder -ChildPath $FileName
            if ( (Test-Path -LiteralPath $SessionFolder) -eq $false ) {
                New-Item -Path $SessionFolder -ItemType Directory | Out-Null
            }
            $ContentTargetFolder = $SessionFolder
        }
        else {
            $ContentTargetFolder = $DownloadFolder
        }

        if ( ! $NoVideos -and (Test-SessionHasOnDemandContent -Session $SessionToGet)) {

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
                        if ($script:CustomCookieFile -and $_.Exception.Data['CookieFileFailure']) { throw }
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
                                    Write-Debug 'KMS bearer token could not be resolved; stream download may fail with 401 on key retrieval'
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

                                if ( $CookiesFile -and -not $script:CustomCookieFile) { $Arg += ('--cookies {0}' -f $CookiesFile) }
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
                        $targetIsMicrosoftUrl = (Test-IsProtectedContentUrl -Url $Endpoint) -or (Test-IsProtectedContentUrl -Url $downloadLink)
                        $ytDlpCookieFile = $null

                        if ($requiresMicrosoftCookieAuth -and $targetIsMicrosoftUrl -and -not $CookiesFromBrowser -and -not $script:CustomCookieFile) {
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
                        $videoMetadataOriginalUrl = Get-SessionOriginalVideoUrl -Session $SessionToGet -EventName $EventName -FallbackVideoUrl $downloadLink
                        $videoMetadataPresentationUrl = Get-SessionPresentationUrl -Session $SessionToGet -FallbackSlidedeckUrl $SlidedeckUrl
                        $videoMetadataRepositoryUrl = Get-SessionRepositoryUrl -Session $SessionToGet
                        $videoMetadata = New-EventSessionVideoMetadata -Session $SessionToGet -EventName $EventName -OriginalVideoUrl $videoMetadataOriginalUrl -PresentationUrl $videoMetadataPresentationUrl -RepositoryUrl $videoMetadataRepositoryUrl
                        Add-BackgroundDownloadJob -Type 2 -FilePath $YouTubeDL -ArgumentList $Arg -File $vidFullFile -Timestamp $SessionTime -scheduleCode ($SessionToGet.sessioncode) -Title ($SessionToGet.Title) -Metadata $videoMetadata
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
                    $existingCaptionFile = Get-ExistingCaptionFile -Path $captionExtFile

                    if ($existingCaptionFile -and -not $Overwrite) {
                        Write-Host ('Caption file exists {0}' -f $existingCaptionFile) -ForegroundColor Gray
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
                                    if ($script:CustomCookieFile -and $_.Exception.Data['CookieFileFailure']) { throw }
                                    #Problem retrieving file, look for alternative options
                                }
                            }

                        }
                        # Check for vtt files before we check any direct caption file (likely docx now)
                        $captionFileLink = $Null
                        $captionLanguageSelected = $null
                        if ( -not [string]::IsNullOrEmpty($OnDemandPage)) {
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
                                    if ($script:CustomCookieFile -and $_.Exception.Data['CookieFileFailure']) { throw }
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
                            if (-not $captionLanguageSelected) {
                                $captionLanguageSelected = [string](Get-ObjectPropertyValue -Object $SessionToGet -Name @('captionLanguage'))
                            }
                            if (-not $captionLanguageSelected -and $captionFileLink -match '(?i)Caption_(?<lang>[a-z]{2,3}([-_][a-z0-9]{2,4})?)\.') {
                                $captionLanguageSelected = $Matches.lang
                            }

                            if ($captionLanguageSelected) {
                                Write-Verbose ('Selected caption language {0} for session {1}' -f $captionLanguageSelected, $SessionToGet.sessioncode)
                            }
                            else {
                                Write-Verbose ('Selected caption language unknown for session {0} (caption source did not expose language metadata)' -f $SessionToGet.sessioncode)
                            }
                            Write-Verbose ('Retrieving caption file from URL {0}' -f $captionFileLink)

                            # Suffix the file with the language subtag, eg Session.bg.vtt, so media
                            # players can determine the subtitle language from the file name.
                            $captionLanguageTag = ConvertTo-CaptionLanguageTag -Language $captionLanguageSelected
                            if (-not $captionLanguageTag) {
                                Write-Verbose ('No usable caption language subtag for session {0}; saving captions without language suffix' -f $SessionToGet.sessioncode)
                            }
                            $captionFullFile = Add-CaptionLanguageToFileName -Path $captionExtFile -LanguageTag $captionLanguageTag

                            if ((Test-Path -LiteralPath $captionFullFile) -and -not $Overwrite) {
                                Write-Host ('Caption file exists {0}' -f $captionFullFile) -ForegroundColor Gray
                            }
                            else {
                                Write-Verbose ('Attempting download {0} to {1}' -f $captionFileLink, $captionFullFile)
                                $captionAuthHeaders = $null
                                if (Test-IsProtectedContentUrl -Url $captionFileLink) {
                                    $captionAuthHeaders = Get-MSADownloadAuthHeaders -Url $captionFileLink -Proxy $ProxyURL
                                    if ($captionAuthHeaders) {
                                        Write-Verbose ('Using authenticated download for caption file of session {0}' -f $SessionToGet.sessioncode)
                                    }
                                }
                                Add-BackgroundDownloadJob -Type 3 -FilePath $captionFullFile -DownloadUrl $captionFileLink -File $captionFullFile -Timestamp $SessionTime -scheduleCode ($SessionToGet.sessioncode) -Title ($SessionToGet.Title) -Headers $captionAuthHeaders -Proxy $ProxyURL
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
            if ($EventType -eq 'CUSTOM' -and @(Get-SessionPresentationUrls -Session $SessionToGet).Count -gt 0) {
                $presentationResult = Add-SessionPresentationDownloadJobs -Session $SessionToGet -CatalogUri ([uri]$EventAPIUrl) -Folder $ContentTargetFolder -BaseName $FileName -Timestamp $SessionTime -Overwrite:$Overwrite -Proxy $ProxyURL
                $DeckInfo[$InfoExist] += $presentationResult.Existing
                $DeckInfo[$InfoPlaceholder] += $presentationResult.Missing
            }
            else {
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
                        if ($script:CustomCookieFile -and $_.Exception.Data['CookieFileFailure']) { throw }
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
                            $slidedeckAuthHeaders = $null
                            if (Test-IsProtectedContentUrl -Url $DownloadURL) {
                                $slidedeckAuthHeaders = Get-MSADownloadAuthHeaders -Url $DownloadURL -Proxy $ProxyURL
                                if ($slidedeckAuthHeaders) {
                                    Write-Verbose ('Using authenticated download for slidedeck of session {0}' -f $SessionToGet.sessioncode)
                                }
                            }
                            Add-BackgroundDownloadJob -Type 1 -FilePath $slidedeckFullFile -DownloadUrl $DownloadURL -File $slidedeckFullFile -Timestamp $SessionTime -scheduleCode ($SessionToGet.sessioncode) -Title ($SessionToGet.Title) -Headers $slidedeckAuthHeaders -Proxy $ProxyURL
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
    if ($UseSessionFolders -and -not [string]::IsNullOrWhiteSpace([string]$ContentTargetFolder)) {
        Remove-EmptySessionFolder -Path $ContentTargetFolder
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
}
finally {
    if ($script:CustomCookieFile -and $script:BackgroundDownloadJobs.Count -gt 0) { Stop-BackgroundDownloadJobs }
}