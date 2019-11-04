﻿<#
    .SYNOPSIS
    Script to assist in downloading Microsoft Ignite, Inspire or Build contents or return 
    session information for easier digesting. Video downloads will leverage external utilities, 
    depending on the used video format. To prevent retrieving session information for every run,
    the script will cache session information.

    Be advised that downloading of OnDemand contents from Azure Media Services is throttled to real-time
    speed. To lessen the pain, the script performs simultaneous downloads of multiple videos streams. Those
    downloads will each open in their own (minimized) window so you can track progress. Finally, CTRL-C
    is catched by the script because we need to stop download jobs when aborting the script.

    .AUTHOR
    Michel de Rooij 	        http://eightwone.com

    Special thanks to:
    Mattias Fors 	            http://deploywindows.info
    Scott Ladewig 	            http://ladewig.com
    Tim Pringle                 http://www.powershell.amsterdam
    Andy Race                   https://github.com/AndyRace
    Richard van Nieuwenhuizen

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Version 3.17, November 4th, 2019

    .DESCRIPTION
    This script can download Microsoft Ignite, Inspire and Build session information and available 
    slidedecks and videos using MyIgnite/MyInspire/MyBuild techcommunity portal.

    Video downloads will leverage one or more utilities:
    - YouTube-dl, which can be downloaded from https://yt-dl.org/latest/youtube-dl.exe. This utility
      needs to reside in the same folder as the script. The script itself will try to download this 
      utility when the utility is not present.
    - ffmpeg, which can be downloaded from https://ffmpeg.zeranoe.com/builds/win32/static/ffmpeg-latest-win32-static.zip. 
      This utility needs to reside in the same folder as the script, or you need to specify its location using -FFMPEG. 
      The utility is used to bind the seperate video and audio streams of Azure Media Services files 
      in single files.

    When you are interested in retrieving session information only, you can use
    the InfoOnly switch.

    .REQUIREMENTS
    The youtube-dl.exe utility requires Visual C++ 2010 redist package
    https://www.microsoft.com/en-US/download/details.aspx?id=5555

    .PARAMETER DownloadFolder
    Specifies location to download sessions to. When omitted, will use 'systemdrive'\'Event'.

    .PARAMETER Format
    For Azure media, default option is worstvideo+bestaudio/best. Alternatively, you can select 
    other formats (when present), e.g. bestvideo+bestaudio. Note that 
    1) The format requested needs to be present in the stream package
    2) Storage required for bestvideo is significantly more than worstvideo

    For YouTube videos, you can use the following formats:
    160          mp4        256x144    DASH video  108k , avc1.4d400b, 30fps, video only
    133          mp4        426x240    DASH video  242k , avc1.4d400c, 30fps, video only
    134          mp4        640x360    DASH video  305k , avc1.4d401e, 30fps, video only
    135          mp4        854x480    DASH video 1155k , avc1.4d4014, 30fps, video only
    136          mp4        1280x720   DASH video 2310k , avc1.4d4016, 30fps, video only
    137          mp4        1920x1080  DASH video 2495k , avc1.640028, 30fps, video only
    18           mp4        640x360    medium , avc1.42001E,  mp4a.40.2@ 96k
    22           mp4        1280x720   hd720 , avc1.64001F,  mp4a.40.2@192k (best, default)
    
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
    Specify what event to download sessions for. Valid values are Ignite (Default), Inspire and Build.

    .PARAMETER OGVPicker
    Specify that you want to pick sessions to download using Out-GridView.

    .PARAMETER InfoOnly
    Tells the script to return session information only.
    Note that by default, only session code and title will be displayed.

    .PARAMETER Overwrite
    Skips detecting existing files, overwriting them if they exist.

    .REVISION
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
          Added Ignite2018 event

    .EXAMPLE
    Download all available contents of Ignite sessions containing the word 'Teams' in the title to D:\Ignite:
    .\Get-EventSession.ps1 -DownloadFolder D:\Ignite-Format 22 -Keyword 'Teams' -Event Ignite

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
#Requires -Version 3.0

[cmdletbinding( DefaultParameterSetName = 'Default' )]
param(
    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [string]$DownloadFolder,

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [string]$Format= $null,

    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [string]$Keyword = '',

    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [string]$Title = '',

    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [string]$Speaker = '',

    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [string]$Product = '',

    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [string]$Category = '',

    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [string]$SolutionArea = '',

    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [string]$LearningPath= '',

    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [string]$Topic= '',

    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [string[]]$ScheduleCode = "",

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [switch]$NoVideos,

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [switch]$NoSlidedecks,

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [string]$FFMPEG,

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [ValidateRange(1,128)] 
    [int]$MaxDownloadJobs=4,

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [uri]$Proxy=$null,

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [ValidateSet('Ignite', 'Inspire','Build','Ignite2018')]
    [string]$Event='Ignite',

    [parameter( Mandatory = $true, ParameterSetName = 'Info')]
    [switch]$InfoOnly,

    [parameter( Mandatory = $true, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [switch]$OGVPicker,

    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [switch]$Overwrite,

    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [switch]$NoRepeats,

    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [switch]$NoGuessing
)

    # Max age for cache, older than this # days will force info refresh
    $MaxCacheAge = 1

    $YouTubeDL = Join-Path $PSScriptRoot 'youtube-dl.exe'
    $FFMPEG= Join-Path $PSScriptRoot 'ffmpeg.exe'
    $SessionCache = Join-Path $PSScriptRoot ('{0}-Sessions.cache' -f $Event)

    $YTlink = 'https://github.com/rg3/youtube-dl/releases/download/2016.09.27/youtube-dl.exe'
    $FFMPEGlink = 'https://ffmpeg.zeranoe.com/builds/win32/static/ffmpeg-latest-win32-static.zip'

    # Fix 'Could not create SSL/TLS secure channel' issues with Invoke-WebRequest
    [Net.ServicePointManager]::SecurityProtocol = "tls12, tls11, tls" 

    $script:BackgroundDownloadJobs= @()

    Function Fix-FileName ($title) {
        return ((((($title -replace '["\\/\?\*]', ' ') -replace ':', '-') -replace '  ', ' ') -replace '\?\?\?', '') -replace '\<|\>|:|"|/|\\|\||\?|\*', '').Trim()
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

    Function Clean-VideoLeftovers ($videofile) {
        $masks= '.mp4.f5_*.part', '.mp4.f5_*.ytdl', '.f1_*.mp4'
	ForEach( $mask in $masks) {
            $FileMask= $videofile -replace '.mp4', $mask
            $files= Get-Item -Path $FileMask -ErrorAction SilentlyContinue | ForEach {
                Write-Verbose ('Removing leftover file {0}' -f $_.fullname)
		$_ | Remove-Item -ErrorAction SilentlyContinue
            }
        }
    }

    Function Get-BackgroundDownloadJobs {
        $Temp= @()
        ForEach( $job in $script:BackgroundDownloadJobs) {
            if($job.job.State -ieq 'Running') {
                $Temp+= $job
            }
            Else {
                # Job finished, add to total
		If( $job.job.State -ieq 'Completed' -and (Test-Path -Path $job.file)) {
                    Write-Host ('Downloaded {0}' -f $job.file) -ForegroundColor Green
                    If( $job.Type -eq 1) {
                        # Slidedeck
                        $DeckInfo[ $InfoDownload]++
                    }
                    Else {
                        # Video
                        $VideoInfo[ $InfoDownload]++
                        Clean-VideoLeftovers $job.file
                    }
                }
                Else {
		    If( $job.Type -eq 1) {
                        Write-Error ('Problem downloading {0}: {1}' -f $job.file, (Receive-Job -Job $job.job))
                    }
                    Else {
                        Write-Error ('Problem downloading {0}: {1}' -f $job.file, (Receive-Job -Job $job.job.ChildJobs[0]))
                    }
                }
                $job.job.ChildJobs | Stop-Job 
                $job.job | Stop-Job -PassThru | Remove-Job -Force
            }
        }
        $script:BackgroundDownloadJobs= $Temp
        $Num= ($script:BackgroundDownloadJobs | Measure-Object).Count
        $NumDeck= ($script:BackgroundDownloadJobs | Where-Object {$_.Type -eq 1} | Measure-Object).Count
        $NumVid= ($script:BackgroundDownloadJobs | Where-Object {$_.Type -eq 2} | Measure-Object).Count
        Write-Progress -Id 2 -Activity 'Background Download Jobs' -Status ('{0} downloads in progress ({1} slidedecks / {2} videos)' -f $Num, $NumDeck, $NumVid) -PercentComplete ($Num / $MaxDownloadJobs * 100)
        return $Num
    }

    Function Stop-BackgroundDownloadJobs {
        ForEach( $BGJob in $script:BackgroundDownloadJobs ) { 
		$BGJob.Job.ChildJobs | Stop-Job 
		$BGJob.Job | Stop-Job -PassThru | Remove-Job -Force
	}
    }

    Function Add-BackgroundDownloadJob {
        param(
            $Type, 
            $FilePath,
            $DownloadUrl,
            $ArgumentList,
            $File
        )
        $JobsRunning= Get-BackgroundDownloadJobs
        If ( $JobsRunning -ge $MaxDownloadJobs) {
            Write-Host ('Maximum background download jobs reached ({0}), waiting for free slot - press Ctrl-C once to abort..' -f $JobsRunning)
            While ( $JobsRunning -ge $MaxDownloadJobs) {
                if ([system.console]::KeyAvailable) { 
                    Start-Sleep 1
                    $key = [system.console]::readkey($true)
                    if (($key.modifiers -band [consolemodifiers]"control") -and ($key.key -eq "C")) {
                        Write-Host "TERMINATING" 
                        Stop-BackgroundDownloadJobs
                        Exit -1
                    }
                }
                Start-Sleep 1
                $JobsRunning= Get-BackgroundDownloadJobs
            }
        }
        Write-Host ('Initiating download of {0}' -f $File)
        If( $Type -eq 1) {
            # Slidedeck
            $job= Start-Job -ScriptBlock { 
                param( $url, $file) 
                $wc = New-Object System.Net.WebClient
                $wc.Encoding = [System.Text.Encoding]::UTF8
                $wc.DownloadFile( $url, $file) 
            } -ArgumentList $DownloadUrl, $FilePath
        }
        Else {
            # Video
            $job= Start-Job -ScriptBlock {
                param( $arglist, $file)
                $myProcess= Start-Process -FilePath $file -ArgumentList $arglist -Passthru -NoNewWindow
                While( Get-Process -Id $myProcess.Id -ErrorAction SilentlyContinue) {
                    Start-Sleep 1
                }
            } -ArgumentList $ArgumentList, $FilePath
        }
	$object= New-Object -TypeName PSObject -Property @{
            Type= $Type
            job= $job
            file= $file
        }

        $script:BackgroundDownloadJobs+= $object
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

    # Determine what event URLs to use
    Switch( $Event) {
        'Ignite' {
            $EventAPIUrl= 'https://api-myignite.techcommunity.microsoft.com'
            $EventSearchURI= 'api/session/search'
            $SessionUrl= 'https://medius.studios.ms/Embed/Video/IG19-{0}'
            $SlidedeckUrl= 'https://mediusprodstatic.studios.ms/presentations/Ignite2019/{0}.pptx'
            $Method= 'Post'
            # Note: to have literal accolades and not string formatter evaluate interior, use a pair:
            $EventSearchBody = '{{"itemsPerPage":{0},"searchText":"*","searchPage":{1},"sortOption":"None","searchFacets":{{"facets":[],"personalizationFacets":[]}}}}'
        }
        'Ignite2018' {
            $EventAPIUrl= 'https://api-myignite.techcommunity.microsoft.com'
            $EventSearchURI= 'api/session/search'
            $SessionUrl= 'https://medius.studios.ms/Embed/Video/IG18-{0}'
            $SlidedeckUrl= 'https://mediusprodstatic.studios.ms/presentations/Ignite2018/{0}.pptx'
            $Method= 'Post'
            # Note: to have literal accolades and not string formatter evaluate interior, use a pair:
            $EventSearchBody = '{{"itemsPerPage":{0},"searchText":"*","searchPage":{1},"sortOption":"None","searchFacets":{{"facets":[],"personalizationFacets":[]}},"from":"2018-09-24T08:00:00-05:00","to":"2018-09-28T19:00:00-05:00"}}'
        }
        'Inspire' {
            $EventAPIUrl= 'https://api.myinspire.microsoft.com'
            $EventSearchURI= 'api/session/search'
            $SessionUrl= ''
            $SlidedeckUrl= ''
            $Method= 'Post'
            $EventSearchBody = '{{"itemsPerPage":{0},"searchText":"*","searchPage":{1},"sortOption":"None","searchFacets":{{"facets":[],"personalizationFacets":[]}}}}'
        }
        'Build' {
            $EventAPIUrl= 'https://api.mybuild.techcommunity.microsoft.com'
            $EventSearchURI= 'api/session/search'
            $SessionUrl= ''
            $SlidedeckUrl= ''
            $Method= 'Post'
            $EventSearchBody = '{{"itemsPerPage":{0},"searchText":"*","searchPage":{1},"sortOption":"None","searchFacets":{{"facets":[],"personalizationFacets":[]}}}}'
        }
        default {
            Write-Error ('Unknown event: {0}' -f $Event)
            Exit -1
        }
    }

    If (-not ($InfoOnly)) {

        # If no download folder set, use system drive with event subfolder
        If( -not( $DownloadFolder)) {
            $DownloadFolder= '{0}\{1}' -f $ENV:SystemDrive, $Event
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
                Write-Host ('youtube-dl.exe not found, will try to download from {0}' -f $YTLink)
                Invoke-WebRequest -Uri $YTLink -OutFile $YouTubeDL -Proxy $ProxyURL
            }
            If ( Test-Path $YouTubeDL) {
                Write-Host ('Running self-update of youtube-dl.exe')

                $Arg = @('-U')
                If ( $ProxyURL) { $Arg += "--proxy $ProxyURL" }

                $pinfo = New-Object System.Diagnostics.ProcessStartInfo
                $pinfo.FileName = $YouTubeDL
                $pinfo.RedirectStandardError = $true
                $pinfo.RedirectStandardOutput = $true
                $pinfo.UseShellExecute = $false
                $pinfo.Arguments = $Arg
                $p = New-Object System.Diagnostics.Process
                $p.StartInfo = $pinfo
                $p.Start() | Out-Null
                $stdout = $p.StandardOutput.ReadToEnd()
                $stderr = $p.StandardError.ReadToEnd()
                $p.WaitForExit()

                If ($p.ExitCode -ne 0) {
                    If ( $stderr -contains 'Error launching') {
                        Write-Error 'Problem running youtube-dl.exe. Make sure this is an x86 system, and the required Visual C++ 2010 redistribution package is installed (available from https://www.microsoft.com/en-US/download/details.aspx?id=5555).'
                        Exit -1
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
                Write-Warning 'Unable to locate or download youtube-dl.exe, will skip downloading YouTube videos'
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
                            Remove-Item -Path $TempFile -Force
                        }
                        Catch {
                            Write-Error ('Problem extracting ffmpeg.exe from {0}' -f $FFMPEGZip)
                        }
                    }
                    Else {
                        Write-Warning 'ffmpeg.exe missing in downloaded archive'
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

    $SessionCacheValid = $false
    If ( Test-Path $SessionCache) {
        Try {
            If ( (Get-childItem -Path $SessionCache).LastWriteTime -ge (Get-Date).AddDays( - $MaxCacheAge)) {
                Write-Host 'Session cache file found, reading session information'
                $data = Import-CliXml -Path $SessionCache -ErrorAction SilentlyContinue
                $SessionCacheValid = $true
            }
            Else {
                Write-Warning 'Cache information expired, will re-read information from catalog'
            }
        }
        Catch {
            Write-Error 'Error reading cache file or cache file invalid - will read from online catalog'
        }
    }

    If ( -not( $SessionCacheValid)) {

        Write-Host 'Reading session catalog'
        $web = @{
            contentType = 'application/json; charset=utf-8'
            userAgent   = 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.103 Safari/537.36'
            requestUri  = [uri]('{0}/{1}' -f $EventAPIUrl, $EventSearchURI)
            itemsPerPage= 100
        }

        Try {
            $SearchBody= $EventSearchBody -f '12', '1'
            Write-Verbose ('Using URI {0}' -f $web.requestUri)
            $searchResultsResponse = Invoke-WebRequest -Uri $web.requestUri -Body $searchbody -Method $Method -ContentType $web.contentType -UserAgent $web.userAgent -WebSession $session -Proxy $ProxyURL
            $searchResults = [system.Text.Encoding]::UTF8.GetString($searchResultsResponse.RawContentStream.ToArray());
        }
        Catch {
            Write-Error ('Problem retrieving session catalog: {0}' -f $error[0])
            Exit 1
        }
        $sessiondata = ConvertFrom-Json -InputObject $searchResults
        [int32] $sessionCount = $sessiondata.total
        [int32] $remainder = 0
 
        $PageCount = [System.Math]::DivRem($sessionCount, $web.itemsPerPage, [ref]$remainder)
        If ($remainder -gt 0) {
            $PageCount ++
        }

        Write-Host ('Reading information for {0} sessions' -f $sessionCount)
        $data = @()
        $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet', [string[]]('sessionCode', 'title'))
        $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
        For ($page = 1; $page -le $PageCount; $page++) {
            Write-Progress -Id 1 -Activity "Retrieving Session Catalog" -Status "Processing page $page of $PageCount" -PercentComplete ($page / $PageCount * 100)
            $SearchBody= $EventSearchBody -f $web.itemsPerPage, $page
            $searchResultsResponse = Invoke-WebRequest -Uri $web.requestUri -Body $searchbody -Method $Method -ContentType $web.contentType -UserAgent $web.userAgent -WebSession $session  -Proxy $ProxyURL
            $searchResults = [system.Text.Encoding]::UTF8.GetString($searchResultsResponse.RawContentStream.ToArray());
            $sessiondata = ConvertFrom-Json -InputObject $searchResults
            ForEach ( $Item in $sessiondata.data) {
                $object = $Item -as [PSCustomObject]
                $object.PSObject.Properties | % {
                    if ($_.Name -eq 'speakerNames') { $object.($_.Name) = @($_.Value) }
                    if ($_.Name -eq 'products') { $object.($_.Name) = @($_.Value -replace [char]9, '/') }
                    if ($_.Name -eq 'contentCategory') { $object.($_.Name) = @(($_.Value -replace [char]9, '/') -replace ' / ', '/') }
                }
                Write-Verbose ('Adding info for session {0}' -f $Object.sessionCode)
                $object.PSObject.TypeNames.Insert(0, 'Session.Information')
                $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                [array]$data += $object
            }
        }
        Write-Progress -Id 1 -Completed -Activity "Finished retrieval of Ignite Session Catalog" 
        Write-Host 'Storing session information'
        $data | Sort-Object -Property sessionCode -Unique | Export-CliXml -Encoding Unicode -Force -Path $SessionCache
    }

    $SessionsToGet = $data

    If ($scheduleCode) {
        Write-Verbose ('Session code(s) specified: {0}' -f ($ScheduleCode -join ','))
        $SessionsToGet = $SessionsToGet | Where-Object { $scheduleCode -contains $_.sessioncode }
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

    If ($Title) {
        Write-Verbose ('Title keyword specified: {0}' -f $Title)
        $SessionsToGet = $SessionsToGet | Where-Object {$_.title -ilike ('*{0}*' -f $Title) }
    }

    If ($NoRepeats) {
        Write-Verbose ('Skipping repeat sessions')
        $SessionsToGet = $SessionsToGet | Where-Object {$_.sessionCode -inotlike '*R'}
    }

    If ($Keyword) {
        Write-Verbose ('Description keyword specified: {0}' -f $Keyword)
        $SessionsToGet = $SessionsToGet | Where-Object {$_.description -ilike ('*{0}*' -f $Keyword) }
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

        [console]::TreatControlCAsInput = $true

        $SessionsSelected = ($SessionsToGet | Measure-Object).Count
        Write-Host ('There are {0} sessions matching your criteria.' -f $SessionsSelected)
        Foreach ($SessionToGet in $SessionsToGet) {
            $i++
            Write-Progress -Id 1 -Activity 'Inspecting session information' -Status "Processing session $i of $SessionsSelected" -PercentComplete ($i / $SessionsSelected * 100)
            $FileName = Fix-FileName ('{0} - {1}' -f $SessionToGet.sessionCode.Trim(), $SessionToGet.title.Trim())

            Write-Host ('Processing info session {0}' -f $FileName)

            If( ! $NoVideos) {
                If ( $DownloadVideos -or $DownloadAMSVideos) {

                    $vidfileName = ("$FileName.mp4")
                    $vidFullFile = Join-Path $DownloadFolder $vidfileName
                    if ((Test-Path -Path $vidFullFile) -and -not $Overwrite) {
                        Write-Host ('Video exists {0}' -f $vidfileName) -ForegroundColor Gray
                        $VideoInfo[ $InfoExist]++
                        Clean-VideoLeftovers $vidFullFile
                    }
                    else {
                        If ( !( [string]::IsNullOrEmpty( $SessionToGet.onDemand)) ) {
                            $downloadLink = $SessionToGet.onDemand
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
                        If( $downloadLink -match 'medius\.studios\.ms\/Embed\/Video' ) {
                            Write-Verbose ('Checking hosted video link {0}' -f $downloadLink)
                            Try {
                                $ValidUrl= Invoke-WebRequest -Uri $downloadLink -Method HEAD -UseBasicParsing -DisableKeepAlive -ErrorAction SilentlyContinue
                            }
                            Catch {
                                $ValidUrl= $false
                            }
                            If( $ValidUrl) {                        
                                $OnDemandPage= (Invoke-WebRequest -Uri $downloadLink -Proxy $ProxyURL).RawContent 
                                
                                # Check for embedded AMS 
                                If( $OnDemandPage -match '<video id="azuremediaplayer" class=".*?" data-id="(?<AzureStreamURL>.*?)"><\/video>') {
                                    Write-Verbose ('Using Azure Media Services URL {0}' -f $matches.AzureStreamURL)
                                    $Endpoint= '{0}(format=mpd-time-csf)' -f $matches.AzureStreamURL
                                    $Arg = @( ('-o "{0}"' -f $vidFullFile), $Endpoint, '--retries 50', '--fragment-retries 50')
                                    If ( $ProxyURL) { $Arg += ('--proxy {0}' -f $ProxyURL) }
                                    If ( $Format) { $Arg += ('--format {0}' -f $Format) } Else { $Arg += ('--format worstvideo+bestaudio/best') }
                                }
                                Else {
                                    # Check for embedded YouTube 
                                    If( $OnDemandPage -match '"https:\/\/www\.youtube-nocookie\.com\/embed\/(?<YouTubeID>.+?)\?enablejsapi=1&"') {
                                        $Endpoint= 'https://www.youtube.com/watch?v={0}' -f $matches.YouTubeID
                                        Write-Verbose ('Using YouTube URL {0}' -f $Endpoint)
                                        $Arg = @( ('-o "{0}"' -f $vidFullFile), $Endpoint, '--retries 50', '--fragment-retries 50')
                                        If ( $ProxyURL) { $Arg += ('--proxy {0}' -f $ProxyURL) }
                                        If ( $Format) { $Arg += ('--format {0}' -f $Format) } Else { $Arg += ('--format 22') }
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
                            If( $downloadLink) {
                                $Endpoint= $downloadLink
                                $Arg = @( ('-o "{0}"' -f $vidFullFile), $downloadLink, '--no-check-certificate', '--retries 50', '--fragment-retries 50')
                                If ( $ProxyURL) { 
                                    $Arg += ('--proxy "{0}"' -f $ProxyURL)
                                }
                                If ( $Format) { 
                                    $Arg += ('--format {0}' -f $Format) 
                                }
                            }
                            Else {
                                Write-Warning ('No video link for {0}' -f ($SessionToGet.Title))
                                $Endpoint= $null
                            }
                        }
                        If( $Endpoint) {
                            # Direct, AMS or YT video found, attempt download
                            Write-Verbose ('Running: youtube-dl.exe {0}' -f ($Arg -join ' '))
                            Add-BackgroundDownloadJob -Type 2 -FilePath $YouTubeDL -ArgumentList $Arg -File $vidFullFile
                        }
                        Else {
                            # Video not available or no link found
                            $DeckInfo[ $InfoPlaceholder]++
                        }
                    }
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

                If ($downloadLink -match "view.officeapps.live.com.*PPTX" -or $downloadLink -match 'downloaddocument' -or $downloadLink -match 'medius') {
                    If( $SessionToGet.slidedeck -match 'downloaddocument') {
                        # Slidedeck offered is PDF format
                        $slidedeckFile = '{0}.pdf' -f $FileName
                        $DeckType= 1
                    }
                    Else {
                        $slidedeckFile = '{0}.pptx' -f $FileName
                        $DeckType= 0
                    }
                    $slidedeckFullFile = Join-Path $DownloadFolder $slidedeckFile
                    if ((Test-Path -Path  $slidedeckFullFile) -and -not $Overwrite) {
                        Write-Host ('Slidedeck exists {0}' -f $slidedeckFile) -ForegroundColor Gray 
                        $DeckInfo[ $InfoExist]++
                    }
                    else {
                        If( $DeckType= 0) {
                            $encodedURL = ($downloadLink -split 'src=')[1]
                        }
                        Else {
                            $encodedURL = $downloadLink
                        }
                        $DownloadURL = [System.Web.HttpUtility]::UrlDecode( $encodedURL)
                        Try {
                            $ValidUrl= Invoke-WebRequest -Uri $DownloadUrl -Method HEAD -UseBasicParsing -DisableKeepAlive -ErrorAction SilentlyContinue
                        }
                        Catch {
                            $ValidUrl= $false
                        }
                        If( $ValidUrl) {                        
                            Write-Verbose ('Downloading {0} to {1}' -f $DownloadURL,  $slidedeckFullFile)
                            Add-BackgroundDownloadJob -Type 1 -FilePath $slidedeckFullFile -DownloadUrl $DownloadURL -File $slidedeckFullFile
                        }
                        Else {
                            Write-Warning ('Skipping: Unavailable {0}' -f $DownloadURL)
                            $DeckInfo[ $InfoPlaceholder]++
                        }
                    }
                }
                Else {
                    Write-Warning ('No slidedeck link for {0}' -f ($SessionToGet.Title))
                }
            }

            if ([system.console]::KeyAvailable) { 
                $key = [system.console]::readkey($true)
                if (($key.modifiers -band [consolemodifiers]"control") -and ($key.key -eq "C")) {
                    Write-Host "TERMINATING"
                    Stop-BackgroundDownloadJobs
                    Exit -1
                }
            }
                   
        }

        Write-Progress -Id 1 -Completed -Activity "Finished processing session information"

        $JobsRunning= Get-BackgroundDownloadJobs
        If ( $JobsRunning -gt 0) {
            Write-Host ('Waiting for download jobs to finish - press Ctrl-C once to abort)' -f $JobsRunning)
            While  ( $JobsRunning -gt 0) {
                if ([system.console]::KeyAvailable) { 
                    Start-Sleep 1
                    $key = [system.console]::readkey($true)
                    if (($key.modifiers -band [consolemodifiers]"control") -and ($key.key -eq "C")) {
                        Write-Host "TERMINATING"
                        Stop-BackgroundDownloadJobs
                        Exit -1
                    }
                }
                $JobsRunning= Get-BackgroundDownloadJobs
            }
        }
        Else {
            Write-Host ('Background download jobs have finished' -f $JobsRunning)
        }

        Write-Progress -Id 2 -Completed -Activity "Download jobs finished"  

        Write-Host ('Downloaded {0} slide decks and {1} videos.' -f $DeckInfo[ $InfoDownload], $VideoInfo[ $InfoDownload])
        Write-Host ('{0} slide decks and {1} videos are not yet available.' -f $DeckInfo[ $InfoPlaceholder], $VideoInfo[ $InfoPlaceholder])
        Write-Host ('{0} slide decks and {1} videos were skipped as they are already present.' -f $DeckInfo[ $InfoExist], $VideoInfo[ $InfoExist])
    }


