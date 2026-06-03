# Changelog

## 4.45
- Routed caption file downloads through the background download engine instead of synchronously
- Extended Add-BackgroundDownloadJob slidedeck (Type 1) and caption (Type 3) jobs to use Invoke-WebRequest with auth headers when available, falling back to WebClient otherwise
- Added Get-MSADownloadAuthHeaders helper to acquire and serialize MSA Bearer token and session cookies for background job use
- Fixed GetNewClosure compatibility issue on PowerShell 5.x
- Added -ProgrammingLanguage filter parameter to select sessions by programming language (OR logic across specified languages)
- Added -SessionLevel filter parameter to select sessions by level (100/200/300/400)

## 4.44
- Updated authentication for Custom events that need it due to changes in IE/Trident

## 4.43
- Enforced TLS 1.2 only; removed TLS 1.0/1.1 fallback
- Fixed invalid regex in -NoRepeats filter (was silently matching nothing)
- Combined multiple session filter passes into a single pipeline pass
- Added automatic retry with exponential back-off for catalog fetch calls
- Added timeout to yt-dlp and ffmpeg subprocesses to prevent indefinite hangs
- Wrapped yt-dlp/ffmpeg downloads and process execution in error handlers
- Fixed cache file load to properly catch and report corrupt cache files
- Hardened file name sanitisation against platform-invalid characters
- Quoted proxy URL in yt-dlp argument string
- Cached progress log reads in background job polling to reduce disk I/O

## 4.42
- Added Build 2026 support

## 4.41
- Fixed compatibility issue with PowerShell v5.1
- Fixed processing non-Custom events

## 4.40
- Added Custom event support with configurable EventUrl
- Added MSA authentication support to Custom Event for when needed
- Added MSA token caching mechanism
- Fixed terminate cleanup to stop yt-dlp child processes
- Some minor fixes for InfoOnly reliability and runtime summary handling

## 4.39
- Added UseSessionFolders switch to store content of each session in its own folder

## 4.38
- Fixed CookiesFile parameter name

## 4.37
- Fixed downloading to UNC paths

## 4.36
- Fixed downloading of direct video links for MP2T type (YouTube)
- Added Cookies and CookiesFromBrowser support for yt-dlp (YouTube)

## 4.35
- Setting video output preset to mp4 to make sure merging results in mp4 file, not an mkv
- Added header to output

## 4.34
- Added removal of Ignite2025 placeholder files

## 4.33
- Added Ignite 2025
- Removed 2021 and 2022 event options

## 4.32
- Fixed default format when downloading from YouTube

## 4.31
- Fixed downloading Captions for direct video links

## 4.3
- Added Build 2025
- Rewrite for new catalog API endpoint and session hosting

## 4.23
- Added TempPath parameter to specifying yt-dlp temporary files location
- Fixed overwrite mode when calling yt-dlp
- Added parameter description for ConcurrentFragments
- Fixed reporting of failed downloads
- Some minor code cleanup

## 4.22
- Fixed download locations for Ignite 2024 content
- Added Azure Stream format guidance

## 4.21
- Fixed date-range for Ignite 2024 ao

## 4.20
- Added Ignite 2024

## 4.11
- Fixed bug in downloading captions

## 4.10
- Added Build 2024

## 4.02
- Added Ignite 2023

## 4.01
- Updated Event parameter help

## 4.00
- Updated yt-dlp download location
- Changed checking yt-dlp.exe presence & validity

## 3.99
- Fixed reporting of element when we cannot add language filter

## 3.98
- Fixed retrieval of Inspire 2023 catalog

## 3.97
- Added Inspire 2023

## 3.96
- Removed hidden character causing "Â : The term 'Â' is not recognized .." messages

## 3.95
- Fixed localized VTT downloading for Build 2023+ from Azure Media Services

## 3.94
- Added yt-dlp's --concurrent-fragments support (default 4)

## 3.93
- Fixed scraping streams from Azure Media Services for Build2023+
- Reinstated caption downloading with VTT instead of docx (can use Sub to download alt. language)

## 3.92
- Added .docx caption support for Build2023

## 3.91
- Fixed output mentioning youtube-dl instead of actual tool (yt-dlp)

## 3.9
- Fixed retrieval of API-based catalogs for events
- Switched to using REST calls for those API-based catalogs
- Added Refresh switch
- Removed archived events (<2021) as MS archives sessions selectively from previous years
- Merged Ignite2021H1 and Ignite2021H2 to Ignite2021

## 3.83
- Updated for Build 2023
- Removed Ignite 2018 and Ignite 2019

## 3.82
- Fixed new folder creation

## 3.81
- Moved to using yt-dlp, a fork of Youtube-DL (not maintained any longer)

## 3.80
- Fixed redundant passing of Format to YouTube-dl

## 3.79
- Fixed issue with placeholder detection
- Fixed path handling, fixes file detection and timestamping a.o.
- Added PowerShell 5.1 requirement (tested with)

## 3.78
- Fixed content-based help

## 3.77
- Corrected API endpoints for some of the older events

## 3.76
- Removed session code uniqueness when storing session data, as session data now can contain multiple entries per language using the same code

## 3.75
- Added Ignite 2022 support

## 3.74
- Fixed MEC processing of multi-line descriptions

## 3.73
- Added MEC slide deck support
- Fixed MEC parsing of description

## 3.72
- Fixed usage of format & subs arguments for direct YouTube downloads

## 3.71
- Fixed MEC description & speaker parsing

## 3.70
- Added MEC support

## 3.69
- Updated for Inspire 2022

## 3.68
- Fixed caching when specifying Event without year tag, eg. Build vs Build2022
- Removed default Locale as that would mess things up for Events where data does not contain that information (yet)

## 3.67
- Added removal of placeholder deck/video/vtt files

## 3.66
- Fixed filtering on langLocale
- Default Locale set to en-US

## 3.65
- Updated for Build 2022
- Added Locale parameter to filter local content
- Fixed applying timestamp due to DateTime formatting changes

## 3.64
- Changed filter so that default language is picked when specified language is not available

## 3.63
- Fixed keyword filtering

## 3.62
- Added Cleanup video leftover files if video file exists (to remove clutter)
- Changed lifetime of cached session information to 8 hours
- Fixed post-download counts

## 3.61
- Added support for (direct) downloading of Ignite Fall 2021 videos

## 3.60
- Added support for Ignite 2021; specify individual event using Ignite2021H1 (Spring) or Ignite2021H2 (Fall)

## 3.55
- Fixed audio stream selection when requested language is not available or only single audio stream is present

## 3.54
- Fixed adding Language filter when complex Format is specified

## 3.53
- Updated for Inspire 2021

## 3.52
- Updated NoRepeats maximum repeat check
- Added Language parameter to support Azure Media Services hosted videos containing multiple audio tracks

## 3.51
- Updated for Build 2021

## 3.50
- Updated for Ignite 2021
- Small cleanup

## 3.47
- Added Captions to PreferDirect command set

## 3.46
- Changed downloading of caption files in background jobs as well
- Optimized caption downloading preventing unnecessary page downloads

## 3.45
- Help updated for -Event

## 3.44
- Fixed downloading of non-PDF slidedecks

## 3.43
- Fixed Ignite 2020 slidedeck 'trial & error' URL

## 3.42
- Changed source location of ffmpeg. Download will now fetch current static x64 release.

## 3.41
- Fixed: Error message for timeless sessions after downloading caption file
- Fixed: Downloading of caption files when video file is already downloaded

## 3.40
- Modified API endpoint for Ignite 2020
- Changed yearless Event specification to add year suffix, eg Ignite->Ignite2020, etc.
- Fixed Azure Media Services video scraping for Ignite2020

## 3.39
- Added code to deal with specifying \<Event>\<Year>

## 3.38
- Added detection of filetype for presentations (PPTX/PDF)

## 3.37
- Added ExcludeCommunityTopic parameter (so you can skip 'Fun and Wellness' Animal Cam contents)
- Modified Keyword and Title parameters (can be multiple values now)

## 3.36
- Small fix for Inspire repeat session naming

## 3.35
- Updated for Inspire 2020

## 3.34
- Updated for Build 2020
- Added NoRepeat filtering for Build 2020
- Made Event parameter mandatory, and not defaulting to Ignite
- Added filtering example to Format parameter spec

## 3.33
- Fixed typo when specifying format for direct YouTube downloads

## 3.32
- Do not assume Slidedeck exists when size is 0

## 3.31
- Corrected video cleanup logic

## 3.30
- Increased wait cycle during progress refresh
- Added schedule code to progress status
- Revised detection successful video downloads

## 3.29
- Added 'Stopped downloading ..' messages when terminating

## 3.28
- Uncommented line to cleanup output files after downloading video
- Changed 'Error' lines to single line outputs or throws (where appropriate)

## 3.27
- Reworked jobs for downloading videos
- Added status bars for downloading of videos
- Failed video downloads will show last line of error output
- Added replacement of square brackets in file names
- Removed obsolete Clean-VideoLeftOvers call

## 3.26
- Updated mutual exclusion for PreferDirect & other parameters/switches
- Added workaround for long file names (NT Style name syntax)
- Added PowerShell ISE detection
- Added Garbage Collection

## 3.25
- Updated Youtube-DL download URL

## 3.24
- Added PreferDirect switch
- Enhanced Format parameter description

## 3.23
- Added Captions switch and Subs parameter
- Added skipping of additional repeats (schedule code ending in R2/R3)
- Fixed filename construction containing '%'
- Added filtering options to description of Format parameter
- Decreased probing/retrieving video URLs from Azure Media Services (speed benefit)

## 3.22
- Added skipping of processing future sessions

## 3.21
- Added Timestamp switch
- Updated file naming to strip embedded name of format, e.g. f1_V_video_3
- Added stopping of Youtube-DL helper app spawned processes

## 3.20
- Fixed background job cleanup

## 3.19
- Fixed video downloading

## 3.18
- Added Ignite2018 event

## 3.17
- Added NoGuess switch
- Added NoRepeats switch

## 3.16
- Corrected prefixes for Ignite 2019

## 3.15
- Added Topic parameter

## 3.14
- Removed superfluous testing loading of main event page
- Fixed LearningPath option verbose output
- Some code cosmetics

## 3.13
- Updated Ignite catalog endpoints

## 3.12
- Updated to work with current Ignite & Build catalogs
- Bumped the download retry limits for YouTube-dl a bit

## 3.11
- Some more cosmetics

## 3.1
- Updated to work with the Inspire 2019 catalog
- Cosmetics

## 3.01
- Added CTRL-Break notice to 'waiting for downloads' message
- Fixed 'No video located for' message

## 3.0
- Added Build support

## 2.986
- Minor update to accommodate publishing of slideDecks links

## 2.985
- Added Proxy support

## 2.984
- Changed keyword search to description, not abstract
- Fixed searching for Products and Category
- Added searching for SolutionArea
- Added searching for LearningPath

## 2.983
- Added OGVPicker switch

## 2.982
- Minor tweaks

## 2.981
- Added cleanup of occasional leftovers (eg *.mp4.f5_A_aac_UND_2_192_1.ytdl, *.f1_V_video_3.mp4)

## 2.98
- Converted background downloads to single background job queue
- Cosmetics

## 2.971
- Changed regex for YouTube matching to skip 'Coming Soon'
- Made verbose mode less noisy

## 2.97
- Update to change in video downloading location (YouTube)
- Changed default Format due to switch in video hosting - see YouTube format table

## 2.96
- Fixed terminating cleanup when no slidedecks are being downloaded
- Added testing for contents to show contents is not available rather than generic 'problem'

## 2.95
- Fixed encoding of filenames

## 2.94
- Fixed cleanup of finished jobs

## 2.93
- Update to slidedeck downloading routine due to changes in published session info

## 2.92
- Fix 'Could not create SSL/TLS secure channel' issues with Invoke-WebRequest

## 2.91
- Update to video downloading routine due to changes in published session info

## 2.9
- Added Category parameter
- Fixed searching on Product
- Increased itemsPerPage when retrieving catalog

## 2.8
- Added downloading of Azure Media Services hosted streaming media
- Added simultaneous downloading of AMS hosted OnDemand streams
- Added NoSlidedecks switch

## 2.7
- Added Event parameter to switch between Ignite and Inspire catalog
- Renamed script to Get-EventSession
- Changed cached session info name to include event
- Removed obsolete URL parameter
- Added code to download slidedecks in PDF (Inspire)
- Cleanup of script synopsis/description/etc.

## 2.6
- Fixed slide deck downloading
- Added Overwrite switch

## 2.61
- Added placeholder slide deck removal

## 2.62
- Fixed Overwrite logic bug
- Renamed to singular Get-IgniteSession to keep in line with PoSH standards

## 2.63
- Fixed bug reporting failed pptx download
- Added reporting of placeholder decks and videos

## 2.64
- Added processing of direct download links for videos

## 2.65
- Added option to specify multiple sessionCode codes
- Added note in source that format only works for YouTube video downloads
- Added youtube-dl returncode check in case it won't run (e.g. missing VC library)

## 2.66
- Added proper downloading of session info using UTF-8 (no more '???')
- Additional trimming of spaces and CRLF's in property values

## 2.5
- Added InfoOnly switch
- Added Product parameter
- Renamed script to Get-IgniteSession.ps1

## 2.22
- Added URL parameter
- Renamed script to IgniteDownloader.ps1

## 2.21
- Added proxy support, using system configured setting
- Fixed downloading of slidedecks

## 2.20
- Incorporated Tim Pringle's code to use JSON to access MyIgnite catalog
- Added option to select speaker
- Added caching of session information (expires in 1 day, or remove .cache file)
- Removed Start parameter (we're now pre-reading the catalog)

## 2.19
- Added trimming of filenames

## 2.18
- Added option to download for sessions listed in a schedule shared from MyIgnite
- Added lookup of video YouTube URL from MyIgnite if not found in TechCommunity
- Added check to make sure conversation titles begin with session code
- Added check to make sure we skip conversations we've already checked since some RSS IDs are duplicates

## 2.17
- Bumped max post to check to 1750

## 2.16
- More illegal character fixups

## 2.15
- Fixed downloading of differently embedded youtube videos
- Added timestamping of downloaded pptx files
- Minor output changes

## 2.14
- Made filtering case-insensitive
- Added NoVideos to download slidedecks only

## 2.13
- Adjusts pptx timestamp to publishing timestamp

## 2.12
- Replaced pptx download Invoke-WebRequest with .NET webclient request (=faster)
- Fixed titles with backslashes (who does that?)

## 2.11
- Fixed titles with apostrophes
- Added Keyword and Title parameter

## 2.1
- Added video downloading, reformatting code (Michel de Rooij)

## 2.0
- Initial release (Mattias Fors)
