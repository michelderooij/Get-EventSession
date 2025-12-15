<#
    .SYNOPSIS
    Converts and compresses video files using ffmpeg with configurable quality and dimension settings.

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Michel de Rooij
    https://github.com/michelderooij/Get-EventSession
    Version 1.03, December 15th, 2025

    .DESCRIPTION
    This script processes video files (MP4, WMV, AVI) in a specified directory, converting them to MP4 format
    using H.264 video encoding and AAC audio encoding. It can scale videos to target dimensions, apply quality
    settings via CRF (Constant Rate Factor), and only process files meeting minimum dimension requirements.
    The script will not upscale videos.
    
    The script automatically downloads ffmpeg if not present and preserves file timestamps after conversion.
    Only files that result in smaller sizes after conversion are kept (except for WMV and AVI which are always converted).

    .PARAMETER SourcePath
    The source directory path to search for video files. The script will recursively search this directory
    for *.mp4, *.wmv, and *.avi files to process.

    .PARAMETER FFMPEG
    Optional path to ffmpeg.exe. If not specified, the script looks for ffmpeg.exe in the script directory.
    If ffmpeg.exe is not found, it will be automatically downloaded from https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-essentials.zip

    .PARAMETER MinimumHeight
    Optional minimum height in pixels. Only videos with height greater than or equal to this value will be processed.
    If omitted, files are processed regardless of height.

    .PARAMETER MinimumWidth
    Optional minimum width in pixels. Only videos with width greater than or equal to this value will be processed.
    If omitted, files are processed regardless of width.

    .PARAMETER TargetHeight
    Target height in pixels for the output video. The script uses min(TargetHeight, ih) to avoid upscaling.
    Default is 720 pixels. Set to 0 or omit to preserve original height.

    .PARAMETER TargetWidth
    Target width in pixels for the output video. The script uses min(TargetWidth, iw) to avoid upscaling.
    Default is 1024 pixels. Set to 0 or omit to preserve original width.

    .PARAMETER CRF
    Constant Rate Factor for video quality. Range is 0-51 where lower values mean better quality.
    - 0 is lossless
    - 18-23 is visually lossless (23 is default)
    - 28 is acceptable quality
    - 51 is worst quality
    Default is 23, which provides good quality at reasonable file sizes.

    .PARAMETER Preset
    ffmpeg encoding preset that controls encoding speed vs compression efficiency.
    Valid values: ultrafast, superfast, veryfast, faster, fast, medium, slow, slower, veryslow
    - Faster presets = quicker encoding but larger files
    - Slower presets = longer encoding but better compression
    Default is 'medium' which balances speed and compression.

    .PARAMETER Priority
    Process priority for ffmpeg execution.
    Valid values: Idle, BelowNormal, Normal, AboveNormal, High, RealTime
    Default is 'Normal'. Use 'Idle' or 'BelowNormal' to reduce system impact during encoding.

    .EXAMPLE
    .\Convert-MP4.ps1 -SourcePath "D:\Videos"
    
    Processes all video files in D:\Videos using default settings (720p target height, 1024 target width, CRF 23, medium preset).

    .EXAMPLE
    .\Convert-MP4.ps1 -SourcePath "D:\Videos" -MinimumHeight 1080 -TargetHeight 720 -Preset slower
    
    Processes only videos with height >= 1080p, downscaling them to 720p using the 'slower' preset for better compression.

    .EXAMPLE
    .\Convert-MP4.ps1 -SourcePath "D:\Videos" -TargetHeight 0 -TargetWidth 0 -CRF 28 -Preset fast
    
    Recompresses videos at their original dimensions using CRF 28 (lower quality) and fast preset for quick processing.

    .EXAMPLE
    .\Convert-MP4.ps1 -SourcePath "D:\Videos" -FFMPEG "C:\Tools\ffmpeg.exe" -MinimumWidth 1920 -MinimumHeight 1080
    
    Uses ffmpeg from a custom location and only processes Full HD (1920x1080) or larger videos.

    .EXAMPLE
    .\Convert-MP4.ps1 -SourcePath "D:\Videos" -Priority BelowNormal
    
    Processes all videos with lower priority to reduce system impact during encoding.

    .NOTES
    - The script preserves creation and modification timestamps
    - Original files are only replaced if the new file is smaller (or for WMV/AVI conversions)
    - Temporary files are created in %TEMP% during processing
    - The script uses H.264 video codec and AAC audio codec for maximum compatibility
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -Path $_ -PathType Container })]
    [string]$SourcePath,

    [Parameter(Mandatory = $false)]
    [string]$FFMPEG,

    [Parameter(Mandatory = $false)]
    [int]$MinimumHeight,

    [Parameter(Mandatory = $false)]
    [int]$MinimumWidth,

    [Parameter(Mandatory = $false)]
    [int]$TargetHeight = 720,

    [Parameter(Mandatory = $false)]
    [int]$TargetWidth = 1024,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 51)]
    [int]$CRF = 23,

    [Parameter(Mandatory = $false)]
    [ValidateSet('ultrafast', 'superfast', 'veryfast', 'faster', 'fast', 'medium', 'slow', 'slower', 'veryslow')]
    [string]$Preset = 'medium',

    [Parameter(Mandatory = $false)]
    [ValidateSet('Idle', 'BelowNormal', 'Normal', 'AboveNormal', 'High', 'RealTime')]
    [string]$Priority = 'Normal'
)

# Check for ffmpeg.exe
if (-not $FFMPEG) {
    $FFMPEG = Join-Path $PSScriptRoot 'ffmpeg.exe'
}

if (-not (Test-Path $FFMPEG)) {
    Write-Host ('ffmpeg.exe not found at {0}, attempting to download...' -f $FFMPEG) -ForegroundColor Yellow
    
    $FFMPEGlink = 'https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-essentials.zip'
    $tempFile = Join-Path $env:TEMP 'ffmpeg-release-essentials.zip'
    
    try {
        Invoke-WebRequest -Uri $FFMPEGlink -OutFile $tempFile
        
        if( Test-Path $tempFile) {
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            Write-Host ('Downloaded {0}, extracting ffmpeg.exe...' -f $tempFile) -ForegroundColor Yellow
            
            $FFMPEGZip = [System.IO.Compression.ZipFile]::OpenRead( $tempFile)
            $FFMPEGEntry = $FFMPEGZip.Entries | Where-Object { $_.FullName -like '*/bin/ffmpeg.exe' }
            
            if( $FFMPEGEntry) {
                try {
                    [System.IO.Compression.ZipFileExtensions]::ExtractToFile( $FFMPEGEntry, $FFMPEG, $true)
                    $FFMPEGZip.Dispose()
                    Remove-Item -LiteralPath $tempFile -Force
                    Write-Host ('ffmpeg.exe extracted to {0}' -f $FFMPEG) -ForegroundColor Green
                }
                catch {
                    throw ('Problem extracting ffmpeg.exe from {0}: {1}' -f $FFMPEGZip, $_.Exception.Message)
                }
            }
            else {
                throw 'ffmpeg.exe not found in downloaded archive'
            }
        }
    }
    catch {
        throw ('Unable to download or extract ffmpeg.exe: {0}' -f $_.Exception.Message)
    }
}

if (-not (Test-Path $FFMPEG)) {
    throw ('ffmpeg.exe not found at {0}' -f $FFMPEG)
}

Write-Host ('Using ffmpeg.exe at {0}' -f $FFMPEG) -ForegroundColor Green

$shell = New-Object -ComObject Shell.Application

# Get all video files from source path
$videoFiles = Get-ChildItem -Path $SourcePath -Recurse -Include *.mp4, *.wmv, *.avi -ErrorAction SilentlyContinue

foreach( $inputVid in $videoFiles) {
    Write-Host ('Processing {0}' -f $inputVid.FullName) -ForegroundColor White
    
    if( $inputVid -and (Test-Path $inputVid.FullName)) {
        $folder = $shell.Namespace( $inputVid.DirectoryName)
        $file = $folder.ParseName( $inputVid.Name)
        $title = if( $file.ExtendedProperty( 'System.Title')) { $file.ExtendedProperty( 'System.Title') } else { $file.ExtendedProperty( 'System.Video.StreamName') }
        $durationMin = [math]::Round( $file.ExtendedProperty( 'System.Media.Duration') / 10000000 / 60, 2)
        $height = $file.ExtendedProperty( 'System.Video.FrameHeight')
        $width = $file.ExtendedProperty( 'System.Video.FrameWidth')
        $bitrate = $file.ExtendedProperty( 'System.Video.EncodingBitrate')
        $totalBitrate = $file.ExtendedProperty( 'System.Video.TotalBitrate')

        Write-Host ('Properties of {3}: Duration {0}m, Width: {1}, Height: {2} Bitrate: {4} (Total {5})' -f $durationMin, $width, $height, $title, $bitrate, $totalBitrate) -ForegroundColor White

        # Check minimum dimensions
        $shouldProcess = $true
        
        if( [uint]$MinimumHeight -gt 0 -and [uint]$height -gt 0 -and $height -lt $MinimumHeight) {
            Write-Host ('{0} height ({1}) is below minimum ({2}), skipping' -f $inputVid.FullName, $height, $MinimumHeight) -ForegroundColor Yellow
            $shouldProcess = $false
        }
        
        if( [uint]$MinimumWidth -gt 0 -and [uint]$width -gt 0 -and $width -lt $MinimumWidth) {
            Write-Host ('{0} width ({1}) is below minimum ({2}), skipping' -f $inputVid.FullName, $width, $MinimumWidth) -ForegroundColor Yellow
            $shouldProcess = $false
        }

        if( $shouldProcess) {
            $orgFile = $inputVid.FullName
            $tempFile = Join-Path $env:TEMP ([io.path]::GetFileName( $orgFile))

            # Construct scale filter
            $scaleHeight = if( $TargetHeight) { 'min({0},ih)' -f $TargetHeight } else { 'ih' }
            $scaleWidth = if( $TargetWidth) { 'min({0},iw)' -f $TargetWidth } else { 'iw' }
            $filt = 'scale=''{0}'':''{1}''' -f $scaleWidth, $scaleHeight

            Write-Host ('Running ffmpeg with preset: {0}, CRF: {1}, filter: {2}, priority: {3}' -f $Preset, $CRF, $filt, $Priority) -ForegroundColor Cyan

            $ffmpegArgs = @(
                '-nostdin'
                '-y'
                '-i', ('"{0}"' -f $orgFile)
                '-c:v', 'libx264'
                '-pix_fmt', 'yuv420p'
                '-c:a', 'aac'
                '-b:a', '128k'
                '-threads', '0'
                '-preset', $Preset
                '-crf', $CRF
                '-vf', $filt
                '-map', '0'
                '-map_metadata', '0'
                ('"{0}"' -f $tempFile)
            )

            $process = Start-Process -FilePath $FFMPEG -ArgumentList $ffmpegArgs -PassThru -NoNewWindow

            # Set process priority if not default
            if( $Priority -ne 'Normal') {
                try {
                    $process.PriorityClass = $Priority
                }
                catch {
                    Write-Warning ('Failed to set process priority to {0}: {1}' -f $Priority, $_)
                }
            }

            # Wait for process to complete
            $process.WaitForExit()

            if( ($process.ExitCode -eq 0) -and (Test-Path $tempFile)) {
                $newFile = Get-ChildItem -Path $tempFile
                
                if( $inputVid.Length -gt $newFile.Length -or $orgFile -like '*.wmv' -or $orgFile -like '*.avi') {
                    $ct = $inputVid.CreationTime
                    $lwt = $inputVid.LastWriteTime

                    Write-Host ('Setting LastWriteTime to {0}' -f $lwt) -ForegroundColor Green
                    $newFile.LastWriteTime = $lwt

                    Write-Host ('Before: {0:N2}MB, After: {1:N2}MB - Savings: {2:N2}MB' -f ($inputVid.Length / 1MB), ($newFile.Length / 1MB), (($inputVid.Length - $newFile.Length) / 1MB)) -ForegroundColor Green

                    $newOrgFile = ($orgFile -replace '\.wmv$', '.mp4') -replace '\.avi$', '.mp4'

                    Write-Host ('Moving {0} to {1}' -f $tempFile, $newOrgFile) -ForegroundColor White
                    Move-Item -Path $tempFile -Destination $newOrgFile -Force

                    Write-Host ('Setting CreationTime to {0}' -f $ct) -ForegroundColor Green
                    (Get-ChildItem $newOrgFile).CreationTime = $ct

                    if( $newOrgFile -ne $orgFile) {
                        Remove-Item -LiteralPath $orgFile -Force
                    }
                }
                else {
                    Write-Host ('New size of {0} is not smaller (change {1:N2}MB) - not replacing' -f $orgFile, (($inputVid.Length - $newFile.Length) / 1MB)) -ForegroundColor Yellow
                    Remove-Item -Path $tempFile -Force
                }
            }
            else {
                Write-Host ('ffmpeg failed to process {0}' -f $orgFile) -ForegroundColor Red
            }
        }
    }
    else {
        Write-Host ('{0} does not exist (any longer), skipping' -f $inputVid.FullName) -ForegroundColor Yellow
    }
}
