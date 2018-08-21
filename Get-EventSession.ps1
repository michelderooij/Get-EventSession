<#
    .SYNOPSIS
    Script to assist in downloading Microsoft Ignite or Inspire contents or return 
    session information for easier digesting. Be advised that downloading videos will 
    make the video non-interactive, by removing options provided by Stream such as
    subtitles, transcripts, etc.

    .AUTHOR
    Michel de Rooij 	http://eightwone.com

    Special thanks to:
    Mattias Fors 	    http://deploywindows.info
    Scott Ladewig 	    http://ladewig.com
    Tim Pringle         http://www.powershell.amsterdam
    Andy Race           https://github.com/AndyRace

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Version 2.7, July 23rd, 2018

    .DESCRIPTION
    This script can download Microsoft Ignite & Inspire session information and available 
    slidedecks and videos using MyIgnite/MyInspire portal.

    Video downloads will leverage a utility which can be downloaded from
    https://yt-dl.org/latest/youtube-dl.exe, and needs to reside in the same folder
    as the script. The script will try to download the utility when it is not present.

    When you are interested in retrieving session information only, you can use
    the InfoOnly switch.

    To prevent retrieving session information for every run, the script will cache
    session information

    .REQUIREMENTS
    The youtube-dl.exe utility requires Visual C++ 2010 redist package
    https://www.microsoft.com/en-US/download/details.aspx?id=5555

    .PARAMETER DownloadFolder
    Specifies location to download sessions to. When omitted, will use 'systemdrive'\'Event'.

    .PARAMETER Format
    Specify mp4 video format to download using youtube-dl.exe. This only works
    for videos retrieved from YouTube; direct downloads are in the format provided.

    Possible values:
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
    Only retrieve sessions for this product.

    .PARAMETER ScheduleCode
    Only retrieve sessions with this session code. You can use one or more codes.

    .PARAMETER NoVideos
    Switch to indicate you only want to download the slidedecks

    .PARAMETER Start
    Item number to start crawling with - useful for restarts

    .PARAMETER Event
    Specify what event to download sessions for. Valid values are Ignite (Default), 
    Inspire.

    .PARAMETER InfoOnly
    Tells the script to return session information only.
    Note that by default, only session code and title will be displayed.

    .PARAMETER Overwrite
    Skips detecting existing files, overwriting them if they exist.

    .REVISION
    2.0  Initial (Mattias Fors)
    2.1  Added video downloading, reformatting code (Michel de Rooij)
    2.11 Fixed titles with apostrophes
         Added Keyword and Title parameter
    2.12 Replaced pptx download Invoke-WebRequest with .NET webclient request (=faster)
         Fixed titles with backslashes (who does that?)
    2.13 Adjusts pptx timestamp to publishing timestamp
    2.14 Made filtering case-insensitive
         Added NoVideos to download slidedecks only
    2.15 Fixed downloading of differently embedded youtube videos
         Added timestamping of downloaded pptx files
         Minor output changes
    2.16 More illegal character fixups
    2.17 Bumped max post to check to 1750
    2.18 Added option to download for sessions listed in a schedule shared from MyIgnite
         Added lookup of video YouTube URl from MyIgnite if not found in TechCommunity
         Added check to make sure conversation titles begin with session code
         Added check to make sure we skip conversations we've already checked since some RSS IDs are duplicates
    2.19 Added trimming of filenames
    2.20 Incorporated Tim Pringle's code to use JSON to acess MyIgnite catalog
         Added option to select speaker
         Added caching of session information (expires in 1 day, or remove .cache file)
         Removed Start parameter (we're now pre-reading the catalog)
    2.21 Added proxy support, using system configured setting
         Fixed downloading of slidedecks
    2.22 Added URL parameter
         Renamed script to IgniteDownloader.ps1
    2.5  Added InfoOnly switch
         Added Product parameter
         Renamed script to Get-IgniteSession.ps1
    2.6  Fixed slide deck downloading
         Added Overwrite switch
    2.61 Added placeholder slide deck removal
    2.62 Fixed Overwrite logic bug
         Renamed to singular Get-IgniteSession to keep in line with PoSH standards
    2.63 Fixed bug reporting failed pptx download
         Added reporting of placeholder decks and videos
    2.64 Added processing of direct download links for videos
    2.65 Added option to specify multiple sessionCode codes
         Added note in source that format only works for YouTube video downloads.
         Added youtube-dl returncode check in case it won't run (e.g. missing VC library).
    2.66 Added proper downloading of session info using UTF-8 (no more '???')
         Additional trimming of spaces and CRLF's in property values
    2.7  Added Event parameter to switch between Ignite and Inspire catalog
         Renamed script to Get-EventSession
         Changed cached session info name to include event
         Removed obsolete URL parameter
         Added code to download slidedecks in PDF (Inspire)
         Cleanup of script synopsis/description/etc.

    .EXAMPLE
    Download all available contents of Inspire sessions containing the word 'Teams' in the title to D:\Inspire:
    .\Get-EventSession.ps1 -DownloadFolder D:\Inspire -Format 18 -Keyword 'Teams' -Event Inspire

    .EXAMPLE
    Get information of all sessions, and output only location and time information for sessions (co-)presented by Tony Redmond:
    .\Get-EventSession.ps1 -InfoOnly | Where {$_.Speakers -contains 'Tony Redmond'} | Select Title, location, startDateTime

    .EXAMPLE
    Download all available contents of sessions BRK3248 and BRK3186 to D:\Ignite
    .\Get-EventSession.ps1 -DownloadFolder D:\Ignite -ScheduleCode BRK3248,BRK3186
    #>
#Requires -Version 3.0

[cmdletbinding( DefaultParameterSetName = 'Default' )]
param(
    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [string]$DownloadFolder,

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [ValidateSet(160, 133, 134, 135, 136, 137, 18, 22)]
    [int]$Format = 22,

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
    [string[]]$ScheduleCode = "",

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [switch]$NoVideos,

    [parameter( Mandatory = $false, ParameterSetName = 'Download')]
    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [ValidateSet('Ignite', 'Inspire')]
    [string]$Event='Ignite',

    [parameter( Mandatory = $true, ParameterSetName = 'Info')]
    [switch]$InfoOnly,

    [parameter( Mandatory = $false, ParameterSetName = 'Default')]
    [parameter( Mandatory = $false, ParameterSetName = 'Info')]
    [switch]$Overwrite
)
begin {

    Function Fix-FileName ($title) {
        return ((((($title -replace '["\\/\?\*]', ' ') -replace ':', '-') -replace '  ', ' ') -replace '\?\?\?', '') -replace "'","").Trim()
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

}

process {

    # Max age for cache, older than this # days will force info refresh
    $MaxCacheAge = 1

    $YouTubeDL = Join-Path $PSScriptRoot 'youtube-dl.exe'
    $SessionCache = Join-Path $PSScriptRoot ('{0}-Sessions.cache' -f $Event)

    $ProxyURL = Get-IEProxy
    If ( $ProxyURL) {
        Write-Host "Using proxy address $ProxyURL"
    }
    Else {
        Write-Host "No proxy setting detected, using direct connection"
    }

    # Determine what event URLs to use
    Switch( $Event) {
        'Ignite' {
            $EventAPIUrl= 'https://api.myignite.microsoft.com/api'
            $EventWebUrl= 'https://myignite.microsoft.com/'
        }
        'Inspire' {
            $EventAPIUrl= 'https://api.myinspire.microsoft.com/api'
            $EventWebUrl= 'https://myinspire.microsoft.com/'
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
                $YTlink = 'https://github.com/rg3/youtube-dl/releases/download/2016.09.27/youtube-dl.exe'
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
                        Throw 'Problem running youtube-dl.exe. Make sure this is an x86 system, and the required Visual C++ 2010 redistribution package is installed (available from https://www.microsoft.com/en-US/download/details.aspx?id=5555).'
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
                Write-Host 'Unable to locate or download youtube-dl.exe, will skip downloading YouTube videos'
                $DownloadVideos = $false
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
        # Get session info using code from Tim Pringle site http://www.powershell.amsterdam/2016/08/05/using-powershell-to-get-data-for-microsoft-ignite/
        $web = @{
            contentType = 'application/json; charset=utf-8'
            userAgent   = 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.103 Safari/537.36'
            baseURL     = $EventAPIUrl
            searchURL   = 'session/anon/search'
        }
 
        $searchbody = '{"searchText":"*","sortOption":"None","searchFacets":{"facets":[],"personalizationFacets":[]}}'
        Try {
            $request = Invoke-WebRequest -Uri $EventWebUrl -Method Get -ContentType $web.contentType -UserAgent $web.userAgent -SessionVariable session -Proxy $ProxyURL
            $searchResultsResponse = Invoke-WebRequest -Uri "$($web.baseURL)/$($web.searchURL)" -Body $searchbody -Method Post -ContentType $web.contentType -UserAgent $web.userAgent -WebSession $session  -Proxy $ProxyURL
            $searchResults = [system.Text.Encoding]::UTF8.GetString($searchResultsResponse.RawContentStream.ToArray());
        }
        Catch {
            Throw ('Problem retrieving session catalog: {0}' -f $error[0])
            Exit 1
        }
        $sessiondata = ConvertFrom-Json -InputObject $searchResults
        [int32] $sessionCount = $sessiondata.total
        [int32] $remainder = 0
 
        $PageCount = [System.Math]::DivRem($sessionCount, 10, [ref]$remainder)
        If ($remainder -gt 0) {
            $PageCount ++
        }

        Write-Host ('Reading information for {0} sessions' -f $sessionCount)
        $data = @()
        $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet', [string[]]('sessionCode', 'title'))
        $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
        For ($page = 1; $page -le $PageCount; $page++) {
            Write-Progress -Activity "Retrieving MyIgnite Session Catalog" -Status "Processing page $page of $PageCount" -PercentComplete ($page / $PageCount * 100)
            $searchbody = "{`"searchText`":`"*`",`"searchPage`":$($page),`"sortOption`":`"None`",`"searchFacets`":{`"facets`":[],`"personalizationFacets`":[]}}"
            $searchResultsResponse = Invoke-WebRequest -Uri "$($web.baseURL)/$($web.searchURL)" -Body $searchbody -Method Post -ContentType $web.contentType -UserAgent $web.userAgent -WebSession $session  -Proxy $ProxyURL
            $searchResults = [system.Text.Encoding]::UTF8.GetString($searchResultsResponse.RawContentStream.ToArray());
            $sessiondata = ConvertFrom-Json -InputObject $searchResults
            ForEach ( $Item in $sessiondata.data) {
                $object = $Item -as [PSCustomObject]
                $object.PSObject.Properties | % {
                    if ($_.Value.Trim) { $object.($_.Name) = $_.Value.Trim() }
                    if ($_.Name -eq 'speakerNames') { $object.($_.Name) = @($_.Value) }
                }
                Write-Verbose ('Adding info for session {0}' -f $Object.sessionCode)
                $object.PSObject.TypeNames.Insert(0, 'Session.Information')
                $object | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                [array]$data += $object
            }
        }
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
        $SessionsToGet = $SessionsToGet | Where-Object { $Speaker -in $_.speakerNames }
    }

    If ($Product) {
        Write-Verbose ('Product specified: {0}' -f $Product)
        $SessionsToGet = $SessionsToGet | Where-Object { $Product -in $_.products }
    }

    If ($Title) {
        Write-Verbose ('Title keyword specified: {0}' -f $Title)
        $SessionsToGet = $SessionsToGet | Where-Object {$_.title -ilike "*$Title*" }
    }

    If ($Keyword) {
        Write-Verbose ('Abstract keyword specified: {0}' -f $Keyword)
        $SessionsToGet = $SessionsToGet | Where-Object {$_.abstract -ilike "*$Keyword*" }
    }

    If ( $InfoOnly) {
        Write-Verbose ('There are {0} sessions matching your criteria.' -f (($SessionsToGet | Measure-Object).Count))
        Write-Output $SessionsToGet
    }
    Else {

        $i = 0
        $DeckInfo = @(0, 0, 0)
        $VideoInfo = @(0, 0, 0)
        $InfoDownload = 0
        $InfoPlaceholder = 1
        $InfoExist = 2

        $SessionsSelected = ($SessionsToGet | Measure-Object).Count
        Write-Host ('There are {0} sessions matching your criteria.' -f $SessionsSelected)
        Foreach ($SessionToGet in $SessionsToGet) {
            $i++
            Write-Progress -Activity 'Downloading session content' -Status "Downloading $i of $SessionsSelected" -PercentComplete ($i / $SessionsSelected * 100)
            $FileName = Fix-FileName "$($SessionToGet.sessionCode.Trim()) - $($SessionToGet.title.Trim())"

            Write-Host ('Inspecting contents for session {0}' -f $FileName)

            If ( $DownloadVideos -and ($SessionToGet.onDemand -or $SessionToGet.downloadVideoLink) ) {
                Write-Verbose 'Video is available.'
                $vidfileName = ("$FileName.mp4")
                $vidFullFile = Join-Path $DownloadFolder $vidfileName
                if ((Test-Path -Path $vidFullFile) -and -not $Overwrite) {
                    Write-Host "Video file exists, skipping. $($vidfileName)" -ForegroundColor Yellow
                    $VideoInfo[ $InfoExist]++
                }
                else {
                    If ( [string]::IsNullOrEmpty( $SessionToGet.downloadVideoLink) ) {
                        $downloadLink = $SessionToGet.onDemand
                    }
                    Else {
                        $downloadLink = $SessionToGet.downloadVideoLink
                    }
                    If( $downloadLink -match 'medius\.studios\.ms\/Embed\/Video' ) {
                        Write-Verbose ('Video hosted on Azure Media Services, extracting link from {0}' -f $downloadLink)
                        $OnDemandPage= (Invoke-WebRequest -Uri $downloadLink).RawContent 
                        
                        If( $OnDemandPage -match '<video id="azuremediaplayer" class=".*?" data-id="(?<AzureStreamURL>.*?)"><\/video>') {
                            Write-Verbose ('Using Azure Media Services URL {0}' -f $matches.AzureStreamURL)
                        }
                        #<video id="azuremediaplayer" class="azuremediaplayer amp-default-skin fullscreen-bg-video amp-big-play-centered" data-id="https://mediusprod.streaming.mediaservices.windows.net/a7835ea2-3e9d-436b-81bb-4c4c2e3d1d15/3f417748-0329-4f26-90cf-2e817fbb.ism/manifest"></video>
                    }
                    Else {
                        Write-Verbose ('Running: youtube-dl.exe -o "{0}" {1}' -f $vidFullFile, $downloadLink)
                        $Arg = "-o ""$vidFullFile""", $downloadLink, "--no-check-certificate"
                        If ( $ProxyURL) { $Arg += "--proxy $ProxyURL" }
                        Start-Process -FilePath $YouTubeDL -ArgumentList $Arg -NoNewWindow -Wait
                    }
                    If ( Test-Path $vidFullFile) {
                        Write-Host "Downloaded $vidFullFile" -ForegroundColor Green
                        $VideoInfo[ $InfoDownload]++
                    }
                    Else {
                        Write-Host "Problem downloading $vidFullFile from $downloadLink" -ForegroundColor Red
                    }
                }
            }
            Else {
                Write-Host "Skip downloading video: $($SessionToGet.Title)"
            }

            If ($SessionToGet.slideDeck -match "view.officeapps.live.com.*PPTX" -or $SessionToGet.slidedeck -match 'downloaddocument' ) {
                Write-Verbose 'Slide deck is available for download.'
                $slidedeckFile = '{0}.pptx' -f $FileName
                $DeckType= 0
                If( $SessionToGet.slidedeck -match 'downloaddocument') {
                    # Slidedeck offered is PDF format
                    $slidedeckFile = '{0}.pdf' -f $FileName
                    $DeckType= 1
                }
                 $slidedeckFullFile = Join-Path $DownloadFolder $slidedeckFile
                if ((Test-Path -Path  $slidedeckFullFile) -and -not $Overwrite) {
                    Write-Host "Slide deck file exists, skipping. $($slidedeckFile)" -ForegroundColor Yellow
                    $DeckInfo[ $InfoExist]++
                }
                else {
                    If( $DeckType= 0) {
                        $encodedURL = ($sessionToGet.slideDeck -split 'src=')[1]
                    }
                    Else {
                        $encodedURL = $sessionToGet.slideDeck
                    }
                    $slidedeckURL = [System.Web.HttpUtility]::UrlDecode( $encodedURL)
                    Write-Verbose ('Downloading {0} to {1}' -f $slidedeckURL,  $slidedeckFullFile)
                    $wc = New-Object net.webclient
                    $wc.DownloadFile( $slidedeckURL,  $slidedeckFullFile)

                    If (Test-Path  $slidedeckFullFile) {
                        If ((Get-Item -Path  $slidedeckFullFile).Length -eq 0) {
                            Write-Host "File  $slidedeckFullFile is zero length, removing" -ForegroundColor Yellow
                            Remove-Item -Path " $slidedeckFullFile"
                        }
                        If ((Get-Item -Path  $slidedeckFullFile).Length -eq 631596) {
                            Write-Host "File  $slidedeckFullFile is placeholder slide deck, removing" -ForegroundColor Yellow
                            Remove-Item -Path " $slidedeckFullFile"
                            $DeckInfo[ $InfoPlaceholder]++
                        }
                        If (Test-Path  $slidedeckFullFile) {
                            Write-Host "Downloaded  $slidedeckFullFile" -ForegroundColor Green
                            $DeckInfo[ $InfoDownload]++
                        }
                    }
                    Else {
                        Write-Host "Problem downloading  $slidedeckFullFile" -ForegroundColor Red
                    }
                }
            }
        }
        Write-Host ('Downloaded {0} slide decks and {1} videos.' -f $DeckInfo[ $InfoDownload], $VideoInfo[ $InfoDownload])
        Write-Host ('Skipped {0} placeholder slide decks, and {1} videos are not yet available.' -f $DeckInfo[ $InfoPlaceholder], $VideoInfo[ $InfoPlaceholder])
        Write-Host ('{0} slide decks and {1} videos were skipped as they are already present.' -f $DeckInfo[ $InfoExist], $VideoInfo[ $InfoExist])
    }
}
