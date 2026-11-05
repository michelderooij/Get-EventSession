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
    Optional minimum height in pixels. Videos qualify when their height strictly exceeds this value.
    Matching any configured height, width, or size filter qualifies. If no minimum is set, all files qualify.

    .PARAMETER MinimumWidth
    Optional minimum width in pixels. Videos qualify when their width strictly exceeds this value.
    Matching any configured height, width, or size filter qualifies. If no minimum is set, all files qualify.

    .PARAMETER MinimumSize
    Optional minimum source file size in bytes, inclusive. Accepts PowerShell values such as 500MB or 1GB.
    Matching this size or either configured dimension qualifies. Default is 0, which disables the size filter.

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
    NVIDIA encoding uses this value as CQ instead; CQ and CRF quality are not identical, and CQ 0 is automatic, not lossless.

    .PARAMETER Preset
    ffmpeg encoding preset that controls encoding speed vs compression efficiency.
    Valid values: ultrafast, superfast, veryfast, faster, fast, medium, slow, slower, veryslow
    - Faster presets = quicker encoding but larger files
    - Slower presets = longer encoding but better compression
    Default is 'medium' which balances speed and compression.
    NVIDIA encoding maps these presets to p1 (fastest) through p7 (slowest).

    .PARAMETER Priority
    Process priority for ffmpeg execution.
    Valid values: Idle, BelowNormal, Normal, AboveNormal, High, RealTime
    Default is 'Normal'. Use 'Idle' or 'BelowNormal' to reduce system impact during encoding.

    .PARAMETER KeepOriginal
    Rename each replaced source to its full filename plus .bak, overwriting any existing backup.
    Backups are only created for accepted encodes. Retaining a backup does not free disk space.

    .PARAMETER NoNvidia
    Skip NVIDIA encoder detection and always use CPU encoding with libx264.

    .EXAMPLE
    .\Convert-MP4.ps1 -SourcePath "D:\Videos"

    Processes all video files in D:\Videos using default settings (720p target height, 1024 target width, CRF 23, medium preset).

    .EXAMPLE
    .\Convert-MP4.ps1 -SourcePath "D:\Videos" -MinimumHeight 1080 -TargetHeight 720 -Preset slower

    Processes only videos with height > 1080 pixels, downscaling them to 720p using the 'slower' preset for better compression.

    .EXAMPLE
    .\Convert-MP4.ps1 -SourcePath "D:\Videos" -TargetHeight 0 -TargetWidth 0 -CRF 28 -Preset fast

    Recompresses videos at their original dimensions using CRF 28 (lower quality) and fast preset for quick processing.

    .EXAMPLE
    .\Convert-MP4.ps1 -SourcePath "D:\Videos" -FFMPEG "C:\Tools\ffmpeg.exe" -MinimumWidth 1920 -MinimumHeight 1080

    Uses ffmpeg from a custom location and only processes videos wider than 1920 pixels or taller than 1080 pixels.

    .EXAMPLE
    .\Convert-MP4.ps1 -SourcePath "D:\Videos" -Priority BelowNormal

    Processes all videos with lower priority to reduce system impact during encoding.

    .EXAMPLE
    .\Compress-MP4.ps1 -SourcePath "D:\Videos" -MinimumSize 500MB -KeepOriginal

    Processes videos at least 500 MiB in size and retains replaced originals as .bak files.

    .NOTES
    - NVIDIA h264_nvenc is preferred when available; a failed attempt retries with libx264 for this and all remaining files
    - Throughput is original size divided by ffmpeg runtime, including any failed NVIDIA attempt
    - Pressing q stops ffmpeg processing for the current file
    - The script preserves creation and modification timestamps
    - Original files are only replaced if the new file is smaller (or for WMV/AVI conversions)
    - Temporary files are created in %TEMP% during processing
    - The script uses H.264 video codec and AAC audio codec for maximum compatibility
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [string]$SourcePath,

    [Parameter(Mandatory = $false)]
    [string]$FFMPEG,

    [Parameter(Mandatory = $false)]
    [int]$MinimumHeight,

    [Parameter(Mandatory = $false)]
    [int]$MinimumWidth,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 9223372036854775807)]
    [long]$MinimumSize = 0,

    [Parameter(Mandatory = $false)]
    [int]$TargetHeight = 1080,

    [Parameter(Mandatory = $false)]
    [int]$TargetWidth = 1920,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 51)]
    [int]$CRF = 23,

    [Parameter(Mandatory = $false)]
    [ValidateSet('ultrafast', 'superfast', 'veryfast', 'faster', 'fast', 'medium', 'slow', 'slower', 'veryslow')]
    [string]$Preset = 'medium',

    [Parameter(Mandatory = $false)]
    [ValidateSet('Idle', 'BelowNormal', 'Normal', 'AboveNormal', 'High', 'RealTime')]
    [string]$Priority = 'Normal',

    [Parameter(Mandatory = $false)]
    [switch]$KeepOriginal,

    [Parameter(Mandatory = $false)]
    [switch]$NoNvidia
)

function Test-NvidiaEncoder {
    param([string]$FFMPEG)

    try {
        $encoders = & $FFMPEG -hide_banner -loglevel error -encoders 2>$null
        return ($LASTEXITCODE -eq 0 -and [bool]($encoders -match '^\s*V\S*\s+h264_nvenc\s'))
    }
    catch {
        return $false
    }
}

function Get-VideoEncoderArguments {
    param([bool]$NvidiaMode, [string]$Preset, [int]$CRF)

    if ($NvidiaMode) {
        $nvencPresets = @{
            ultrafast = 'p1'; superfast = 'p1'; veryfast = 'p2'; faster = 'p3'
            fast = 'p4'; medium = 'p4'; slow = 'p5'; slower = 'p6'; veryslow = 'p7'
        }
        return @('-hide_banner', '-loglevel', 'info', '-c:v', 'h264_nvenc', '-preset', $nvencPresets[$Preset], '-rc', 'vbr', '-cq', $CRF, '-b:v', '0')
    }

    return @('-hide_banner', '-loglevel', 'info', '-c:v', 'libx264', '-preset', $Preset, '-crf', $CRF)
}

function Read-VideoStopRequest {
    if (-not [Console]::IsInputRedirected -and [Console]::KeyAvailable) {
        return [Console]::ReadKey($true).KeyChar -eq 'q'
    }
    return $false
}

function Invoke-VideoEncodeAttempt {
    param([string]$FFMPEG, [string[]]$Arguments, [string]$Priority)

    $process = $null
    $timer = New-Object System.Diagnostics.Stopwatch
    try {
        $process = Start-Process -FilePath $FFMPEG -ArgumentList $Arguments -PassThru -NoNewWindow -ErrorAction Stop
        $null = $process.Handle
        $timer.Start()
        if ($Priority -ne 'Normal') {
            try { $process.PriorityClass = $Priority }
            catch { Write-Warning ('Failed to set process priority to {0}: {1}' -f $Priority, $_) }
        }

        $cancelled = $false
        while (-not $process.WaitForExit(200)) {
            if (Read-VideoStopRequest) {
                $cancelled = $true
                break
            }
        }
        if (-not $process.HasExited) {
            $process.Kill()
            $process.WaitForExit()
        }
        $timer.Stop()
        $exitCode = $process.ExitCode
        $cancelled = $cancelled -or $exitCode -eq -1073741510
        return [pscustomobject]@{ ExitCode = $exitCode; Cancelled = $cancelled; Seconds = $timer.Elapsed.TotalSeconds }
    }
    finally {
        try {
            if ($process) {
                try {
                    if (-not $process.HasExited) {
                        $process.Kill()
                        $process.WaitForExit()
                    }
                }
                finally { $process.Dispose() }
            }
        }
        finally {
            $timer.Stop()
        }
    }
}

function Invoke-VideoEncoding {
    param(
        [string]$FFMPEG, [string[]]$Arguments, [string]$OutputPath,
        [ref]$NvidiaMode, [string]$Preset, [int]$CRF, [string]$Priority
    )

    $seconds = 0.0
    $success = $false
    try {
        do {
            $useNvidia = $NvidiaMode.Value
            $encoderArgs = Get-VideoEncoderArguments -NvidiaMode $useNvidia -Preset $Preset -CRF $CRF
            $attemptArgs = @($Arguments) + $encoderArgs + @(('"{0}"' -f $OutputPath))
            $result = Invoke-VideoEncodeAttempt -FFMPEG $FFMPEG -Arguments $attemptArgs -Priority $Priority
            $seconds += $result.Seconds
            if ($result.Cancelled) {
                return [pscustomobject]@{ Success = $false; Cancelled = $true; Seconds = $seconds }
            }
            $success = $result.ExitCode -eq 0 -and (Test-Path -LiteralPath $OutputPath) -and (Get-Item -LiteralPath $OutputPath).Length -gt 0
            if ($success) {
                return [pscustomobject]@{ Success = $true; Cancelled = $false; Seconds = $seconds }
            }
            if (Test-Path -LiteralPath $OutputPath) {
                Remove-Item -LiteralPath $OutputPath -Force -ErrorAction Stop
            }
            if ($useNvidia) {
                $NvidiaMode.Value = $false
                Write-Warning 'NVIDIA encoding failed; retrying with libx264 and using CPU encoding for all remaining files.'
            }
        } while ($useNvidia)

        return [pscustomobject]@{ Success = $false; Cancelled = $false; Seconds = $seconds }
    }
    finally {
        if (-not $success -and (Test-Path -LiteralPath $OutputPath)) {
            Remove-Item -LiteralPath $OutputPath -Force -ErrorAction Stop
        }
    }
}

function Install-CompressedVideo {
    param([System.IO.FileInfo]$Original, [string]$OutputPath, [switch]$KeepOriginal)

    $orgFile = $Original.FullName
    $destination = [io.path]::ChangeExtension($orgFile, '.mp4')
    $creationTime = $Original.CreationTime
    $lastWriteTime = $Original.LastWriteTime
    $backupPath = $orgFile + '.bak'
    $backedUp = $false
    (Get-Item -LiteralPath $OutputPath).LastWriteTime = $lastWriteTime
    if ($KeepOriginal) {
        Move-Item -LiteralPath $orgFile -Destination $backupPath -Force -ErrorAction Stop
        $backedUp = $true
    }
    try {
        Move-Item -LiteralPath $OutputPath -Destination $destination -Force -ErrorAction Stop
    }
    catch {
        if ($backedUp) {
            Move-Item -LiteralPath $backupPath -Destination $orgFile -Force -ErrorAction Stop
        }
        throw
    }
    $installed = Get-Item -LiteralPath $destination
    $installed.CreationTime = $creationTime
    $installed.LastWriteTime = $lastWriteTime
    if (-not $KeepOriginal -and $destination -ne $orgFile) {
        Remove-Item -LiteralPath $orgFile -Force -ErrorAction Stop
    }
}

function Get-VideoThroughput {
    param([long]$OriginalBytes, [double]$Seconds)

    if ($Seconds -le 0) { return 0.0 }
    return $OriginalBytes / $Seconds / 1MB
}

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

        if ( Test-Path $tempFile) {
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            Write-Host ('Downloaded {0}, extracting ffmpeg.exe...' -f $tempFile) -ForegroundColor Yellow

            $FFMPEGZip = [System.IO.Compression.ZipFile]::OpenRead( $tempFile)
            $FFMPEGEntry = $FFMPEGZip.Entries | Where-Object { $_.FullName -like '*/bin/ffmpeg.exe' }

            if ( $FFMPEGEntry) {
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

$nvidiaMode = $false
if (-not $NoNvidia) {
    $nvidiaMode = Test-NvidiaEncoder -FFMPEG $FFMPEG
}
Write-Host ('NVIDIA encoding available: {0}' -f $nvidiaMode) -ForegroundColor Cyan

$shell = New-Object -ComObject Shell.Application

# Get all video files from source path
$videoFiles = Get-ChildItem -Path $SourcePath -Recurse -Include *.mp4, *.wmv, *.avi -ErrorAction SilentlyContinue

foreach ( $inputVid in $videoFiles) {
    Write-Host ('Processing {0}' -f $inputVid.FullName) -ForegroundColor White

    if ( $inputVid -and (Test-Path -LiteralPath $inputVid.FullName)) {
        $folder = $shell.Namespace( $inputVid.DirectoryName)
        $file = $folder.ParseName( $inputVid.Name)
        $title = if ( $file.ExtendedProperty( 'System.Title')) { $file.ExtendedProperty( 'System.Title') } else { $file.ExtendedProperty( 'System.Video.StreamName') }
        $durationMin = [math]::Round( $file.ExtendedProperty( 'System.Media.Duration') / 10000000 / 60, 2)
        $height = $file.ExtendedProperty( 'System.Video.FrameHeight')
        $width = $file.ExtendedProperty( 'System.Video.FrameWidth')
        $bitrate = $file.ExtendedProperty( 'System.Video.EncodingBitrate')
        $totalBitrate = $file.ExtendedProperty( 'System.Video.TotalBitrate')

        Write-Host ('Properties of {3}: Duration {0}m, Width: {1}, Height: {2} Bitrate: {4} (Total {5})' -f $durationMin, $width, $height, $title, $bitrate, $totalBitrate) -ForegroundColor White

        # Check minimum dimensions and file size
        $shouldProcess = (
            ($MinimumHeight -le 0 -and $MinimumWidth -le 0 -and $MinimumSize -le 0) -or
            ($MinimumHeight -gt 0 -and $height -gt $MinimumHeight) -or
            ($MinimumWidth -gt 0 -and $width -gt $MinimumWidth) -or
            ($MinimumSize -gt 0 -and $inputVid.Length -ge $MinimumSize)
        )

        if (-not $shouldProcess) {
            Write-Host ('{0}: no minimum filter matched (width: {1}, height: {2}, size: {3} bytes; minimum width: {4}, height: {5}, size: {6} bytes), skipping' -f $inputVid.FullName, $width, $height, $inputVid.Length, $MinimumWidth, $MinimumHeight, $MinimumSize) -ForegroundColor Yellow
        }

        if ( $shouldProcess) {
            $orgFile = $inputVid.FullName
            $originalBytes = $inputVid.Length
            $tempFile = Join-Path $env:TEMP ('Compress-MP4-{0}.mp4' -f [guid]::NewGuid())

            # Construct scale filter
            $scaleHeight = if ( $TargetHeight) { 'min({0},ih)' -f $TargetHeight } else { 'ih' }
            $scaleWidth = if ( $TargetWidth) { 'min({0},iw)' -f $TargetWidth } else { 'iw' }
            $filt = 'scale=''{0}'':''{1}''' -f $scaleWidth, $scaleHeight

            Write-Host ('Running ffmpeg with preset: {0}, CRF: {1}, filter: {2}, priority: {3}' -f $Preset, $CRF, $filt, $Priority) -ForegroundColor Cyan
            Write-Host "Press 'q' to stop ffmpeg encoding and cancel the batch; the original will be preserved." -ForegroundColor Yellow

            $ffmpegArgs = @(
                '-nostdin'
                '-y'
                '-i', ('"{0}"' -f $orgFile)
                '-pix_fmt', 'yuv420p'
                '-c:a', 'aac'
                '-b:a', '128k'
                '-threads', '0'
                '-vf', $filt
                '-map', '0'
                '-map_metadata', '0'
            )

            try {
                $result = Invoke-VideoEncoding -FFMPEG $FFMPEG -Arguments $ffmpegArgs -OutputPath $tempFile -NvidiaMode ([ref]$nvidiaMode) -Preset $Preset -CRF $CRF -Priority $Priority
                if ($result.Cancelled) {
                    Write-Host 'Encoding cancelled; stopping the batch. Original file preserved.' -ForegroundColor Yellow
                    break
                }
                if (-not $result.Success) {
                    Write-Host ('ffmpeg failed to process {0}' -f $orgFile) -ForegroundColor Red
                    continue
                }
                $newFile = Get-Item -LiteralPath $tempFile
                $throughput = Get-VideoThroughput -OriginalBytes $originalBytes -Seconds $result.Seconds
                Write-Host ('Before: {0:N2}MiB, After: {1:N2}MiB - Size reduction: {2:N2}MiB; ffmpeg: {3:N2}s; Throughput: {4:N2}MiB/s' -f ($originalBytes / 1MB), ($newFile.Length / 1MB), (($originalBytes - $newFile.Length) / 1MB), $result.Seconds, $throughput) -ForegroundColor Green

                if ($originalBytes -gt $newFile.Length -or $orgFile -like '*.wmv' -or $orgFile -like '*.avi') {
                    Install-CompressedVideo -Original $inputVid -OutputPath $tempFile -KeepOriginal:$KeepOriginal
                    if ($KeepOriginal) {
                        Write-Host ('Original retained as {0}.bak; backup retains disk usage.' -f $orgFile) -ForegroundColor Green
                    }
                }
                else {
                    Write-Host ('New size of {0} is not smaller - not replacing' -f $orgFile) -ForegroundColor Yellow
                }
            }
            finally {
                if (Test-Path -LiteralPath $tempFile) {
                    Remove-Item -LiteralPath $tempFile -Force -ErrorAction Stop
                }
            }
        }
    }
    else {
        Write-Host ('{0} does not exist (any longer), skipping' -f $inputVid.FullName) -ForegroundColor Yellow
    }
}
