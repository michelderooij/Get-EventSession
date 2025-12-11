<#
    .SYNOPSIS
    Compresses media files embedded in PowerPoint presentations (.pptx) to reduce file size.

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Michel de Rooij
    https://github.com/michelderooij/Get-EventSession
    Version 1.01, December 12th, 2025

    .DESCRIPTION
    This script processes PowerPoint (.pptx) files in a specified directory, compressing embedded images
    (JPEG, PNG) and videos (MP4, WMV, AVI, MOV) to reduce overall file size while maintaining acceptable quality.
    
    The script automatically downloads ffmpeg and pngquant if not present, extracts media from the PPTX,
    recompresses it based on specified quality settings, and updates the presentation with the optimized media.
    Original file timestamps are preserved, and only files that result in size reduction are replaced.

    .PARAMETER SourcePath
    The source directory path to search for .pptx files. The script will recursively search this directory
    for PowerPoint presentations to process.

    .PARAMETER FFMPEG
    Optional path to ffmpeg.exe. If not specified, the script looks for ffmpeg.exe in the script directory.
    If ffmpeg.exe is not found, it will be automatically downloaded from https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-essentials.zip

    .PARAMETER PNGQUANT
    Optional path to pngquant.exe. If not specified, the script looks for pngquant.exe in the script directory.
    If pngquant.exe is not found, it will be automatically downloaded from https://pngquant.org/pngquant-windows.zip

    .PARAMETER ImageQuality
    JPEG compression quality (1-100). Higher values mean better quality but larger files.
    For PNG files, this is used to calculate a quality range (ImageQuality±10).
    Default is 70, which provides good quality at reasonable file sizes.

    .PARAMETER MinimumHeight
    Optional minimum height in pixels for videos. Only videos with height greater than or equal to this value will be processed.
    If omitted, videos are processed regardless of height.

    .PARAMETER MinimumWidth
    Optional minimum width in pixels for videos. Only videos with width greater than or equal to this value will be processed.
    If omitted, videos are processed regardless of width.

    .PARAMETER TargetHeight
    Target height in pixels for video output. The script uses min(TargetHeight, ih) to avoid upscaling.
    Default is 720 pixels. Set to 0 to preserve original height.

    .PARAMETER TargetWidth
    Target width in pixels for video output. The script uses min(TargetWidth, iw) to avoid upscaling.
    Default is 1024 pixels. Set to 0 to preserve original width.

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

    .PARAMETER Backup
    If specified, creates a backup of the original .pptx file with a .bak extension before replacing it with the compressed version.

    .EXAMPLE
    .\Convert-PPTX.ps1 -SourcePath "D:\Presentations"
    
    Processes all .pptx files in D:\Presentations using default settings (70% image quality, 720p video target height).

    .EXAMPLE
    .\Convert-PPTX.ps1 -SourcePath "D:\Presentations" -ImageQuality 85 -TargetHeight 1080 -Preset slower -Backup
    
    Processes presentations with higher image quality (85%), 1080p video target, slower preset for better compression, and creates backups.

    .EXAMPLE
    .\Convert-PPTX.ps1 -SourcePath "D:\Presentations" -MinimumHeight 720 -MinimumWidth 1280 -CRF 28
    
    Only processes videos that are at least 720p (1280x720), using CRF 28 for more aggressive compression.

    .EXAMPLE
    .\Convert-PPTX.ps1 -SourcePath "D:\Presentations" -FFMPEG "C:\Tools\ffmpeg.exe" -PNGQUANT "C:\Tools\pngquant.exe"
    
    Uses custom paths for ffmpeg and pngquant executables.

    .NOTES
    - Requires ffmpeg for video processing (auto-downloaded if not present)
    - Requires pngquant for PNG optimization (auto-downloaded if not present)
    - The script preserves creation and modification timestamps
    - Only replaces files if the compressed version is smaller
    - Temporary files are created in %TEMP% during processing
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true, Position=0)]
    [ValidateScript({ Test-Path -Path $_ -PathType Container })]
    [string]$SourcePath,

    [Parameter(Mandatory=$false)]
    [string]$FFMPEG,

    [Parameter(Mandatory=$false)]
    [string]$PNGQUANT,

    [Parameter(Mandatory=$false)]
    [ValidateRange(1,100)]
    [int]$ImageQuality = 70,

    [Parameter(Mandatory=$false)]
    [int]$MinimumHeight,

    [Parameter(Mandatory=$false)]
    [int]$MinimumWidth,

    [Parameter(Mandatory=$false)]
    [int]$TargetHeight = 720,

    [Parameter(Mandatory=$false)]
    [int]$TargetWidth = 1024,

    [Parameter(Mandatory=$false)]
    [ValidateRange(0, 51)]
    [int]$CRF = 23,

    [Parameter(Mandatory=$false)]
    [ValidateSet('ultrafast', 'superfast', 'veryfast', 'faster', 'fast', 'medium', 'slow', 'slower', 'veryslow')]
    [string]$Preset = 'medium',

    [switch]$Backup
)

function Test-CommandExists {
    param([string]$Name)
    return $null -ne (Get-Command $Name -ErrorAction SilentlyContinue)
}

function Get-VideoProperties {
    param($FilePath, $FFMPEGPath)
    
    try {
        $probe = & $FFMPEGPath -i $FilePath 2>&1
        $probeText = $probe | Out-String
        
        $width = 0
        $height = 0
        
        if( $probeText -match 'Stream.*Video.*?(\d+)x(\d+)') {
            $width = [int]$matches[1]
            $height = [int]$matches[2]
        }
        
        return @{
            Width = $width
            Height = $height
        }
    }
    catch {
        Write-Warning ('Failed to get video properties for {0}' -f $FilePath)
        return @{
            Width = 0
            Height = 0
        }
    }
}

function Recompress-JpegWithBitmap {
    param($SrcPath, $DstPath, [int]$Quality)
    try {
        Add-Type -AssemblyName System.Drawing

        # Load the image as a Bitmap
        $bitmap = New-Object System.Drawing.Bitmap( $SrcPath)

        # Get JPEG encoder
        $encoders = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders()
        $jpegCodec = $encoders | Where-Object { $_.MimeType -eq 'image/jpeg' } | Select-Object -First 1

        if( -not $jpegCodec) {
            $bitmap.Dispose()
            return $false
        }

        # Set up encoder parameters for quality
        $encParams = New-Object System.Drawing.Imaging.EncoderParameters( 1)
        $encParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter( [System.Drawing.Imaging.Encoder]::Quality, [long]$Quality)

        # Save with compression
        $bitmap.Save( $DstPath, $jpegCodec, $encParams)
        $bitmap.Dispose()

        return $true
    } catch {
        Write-Warning ('Error recompressing {0} : {1}' -f $SrcPath, $_)
        return $false
    }
}

function Recompress-PngWithPngquant {
    param($SrcPath, $DstPath, [string]$QualityRange='60-90', $PNGQUANTPath)
    try {
        # Use pngquant for lossy PNG compression with good quality
        $process = Start-Process -FilePath $PNGQUANTPath -ArgumentList @( ('--quality={0}' -f $QualityRange), '--output', ('"""""' + $DstPath + '"""""'), ('"""""' + $SrcPath + '"""""')) -Wait -PassThru -NoNewWindow -RedirectStandardError 'NUL'
        return $process.ExitCode -eq 0
    } catch {
        Write-Warning ('Error optimizing PNG with pngquant {0} : {1}' -f $SrcPath, $_)
        return $false
    }
}

function Reencode-MediaWithFFmpeg {
    param($SrcPath, $DstPath, [int]$TargetHeight, [int]$TargetWidth, [int]$CRF, [string]$Preset, $Extension= '.mp4', $FFMPEGPath)
    try {
        # Construct scale filter
        $scaleHeight = if( $TargetHeight) { 'min({0},ih)' -f $TargetHeight } else { 'ih' }
        $scaleWidth = if( $TargetWidth) { 'min({0},iw)' -f $TargetWidth } else { 'iw' }
        $scaleFilter = 'scale=''{0}'':''{1}''' -f $scaleWidth, $scaleHeight
        
        $arguments = @(
            '-i', $SrcPath
            '-vf', $scaleFilter
            '-y'  # Overwrite output file
        )
        
        switch( $Extension.ToLower()) {
            '.wmv' {
                $arguments+= '-f'
                $arguments+= 'asf'
                $arguments+= '-c:v'
                $arguments+= 'wmv2'
                $arguments+= '-b:v'
                $arguments+= '2500k'
                $arguments+= '-c:a'
                $arguments+= 'wmav2'
                $arguments+= '-b:a'
                $arguments+= '128k'
            }
            '.mov' {
                $arguments+= '-f'
                $arguments+= 'mov'
                $arguments+= '-crf'
                $arguments+= $CRF.ToString()
                $arguments+= '-preset'
                $arguments+= $Preset
                $arguments+= '-c:v'
                $arguments+= 'libx264'
                $arguments+= '-c:a'
                $arguments+= 'aac'
                $arguments+= '-b:a'
                $arguments+= '128k'
            }
            '.avi' {
                $arguments+= '-f'
                $arguments+= 'avi'
                $arguments+= '-c:v'
                $arguments+= 'mpeg4'
                $arguments+= '-q:v'
                $arguments+= '3'
                $arguments+= '-c:a'
                $arguments+= 'mp3'
                $arguments+= '-b:a'
                $arguments+= '128k'
           }
            default {
                # mp4
                $arguments+= '-f'
                $arguments+= 'mp4'
                $arguments+= '-crf'
                $arguments+= $CRF.ToString()
                $arguments+= '-preset'
                $arguments+= $Preset
                $arguments+= '-c:v'
                $arguments+= 'libx264'
                $arguments+= '-c:a'
                $arguments+= 'aac'
                $arguments+= '-b:a'
                $arguments+= '128k'
            }
        }
        $arguments+= $DstPath

        Write-Output ('FFmpeg arguments: {0}' -f ($arguments -join ' '))

        $process = Start-Process -FilePath $FFMPEGPath -ArgumentList $arguments -Wait -PassThru -NoNewWindow -RedirectStandardError 'NUL'
        return $process.ExitCode -eq 0
    } catch {
        Write-Warning ('Error reencoding video {0} : {1}' -f $SrcPath, $_)
        return $false
    }
}

function Update-PptxMedia {
    param($PptxPath, $MediaDir)

    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem

        # Get all media files that were processed
        $mediaFiles = Get-ChildItem -Path $MediaDir -File

        # Open the PPTX for update
        $zip = [System.IO.Compression.ZipFile]::Open( $PptxPath, [System.IO.Compression.ZipArchiveMode]::Update)

        foreach( $file in $mediaFiles) {
            # Use forward slashes for ZIP entry paths
            $entryPath = 'ppt/media/{0}' -f $file.Name

            # Find the existing entry
            $existingEntry = $zip.Entries | Where-Object { $_.FullName -eq $entryPath }

            if( $existingEntry) {
                Write-Output ('Updating: {0}' -f $entryPath)

                # Open the entry and overwrite its content
                $entryStream = $existingEntry.Open()
                $entryStream.SetLength( 0)  # Clear existing content

                # Copy new file content
                $fileStream = [System.IO.File]::OpenRead( $file.FullName)
                $fileStream.CopyTo( $entryStream)

                # Close streams
                $fileStream.Close()
                $fileStream.Dispose()
                $entryStream.Close()
                $entryStream.Dispose()
            }
            else {
                Write-Warning ('Entry not found in archive: {0}' -f $entryPath)
            }
        }

        $zip.Dispose()
        return $true

    } catch {
        Write-Warning ('Error updating PPTX media: {0}' -f $_)
        if( $zip) {
            try { $zip.Dispose() } catch { }
        }
        return $false
    }
}

# -- Begin

# Check for ffmpeg.exe
if( -not $FFMPEG) {
    $FFMPEG = Join-Path $PSScriptRoot 'ffmpeg.exe'
}

if( -not (Test-Path $FFMPEG)) {
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

if( -not (Test-Path $FFMPEG)) {
    throw ('ffmpeg.exe not found at {0}' -f $FFMPEG)
}

Write-Host ('Using ffmpeg.exe at {0}' -f $FFMPEG) -ForegroundColor Green

# Check for pngquant.exe
if( -not $PNGQUANT) {
    $PNGQUANT = Join-Path $PSScriptRoot 'pngquant.exe'
}

if( -not (Test-Path $PNGQUANT)) {
    Write-Host ('pngquant.exe not found at {0}, attempting to download...' -f $PNGQUANT) -ForegroundColor Yellow
    
    $PNGQUANTlink = 'https://pngquant.org/pngquant-windows.zip'
    $tempFile = Join-Path $env:TEMP 'pngquant-windows.zip'
    
    try {
        Invoke-WebRequest -Uri $PNGQUANTlink -OutFile $tempFile
        
        if( Test-Path $tempFile) {
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            Write-Host ('Downloaded {0}, extracting pngquant.exe...' -f $tempFile) -ForegroundColor Yellow
            
            $PNGQUANTZip = [System.IO.Compression.ZipFile]::OpenRead( $tempFile)
            $PNGQUANTEntry = $PNGQUANTZip.Entries | Where-Object { $_.Name -eq 'pngquant.exe' }
            
            if( $PNGQUANTEntry) {
                try {
                    [System.IO.Compression.ZipFileExtensions]::ExtractToFile( $PNGQUANTEntry, $PNGQUANT, $true)
                    $PNGQUANTZip.Dispose()
                    Remove-Item -LiteralPath $tempFile -Force
                    Write-Host ('pngquant.exe extracted to {0}' -f $PNGQUANT) -ForegroundColor Green
                }
                catch {
                    throw ('Problem extracting pngquant.exe from {0}: {1}' -f $PNGQUANTZip, $_.Exception.Message)
                }
            }
            else {
                throw 'pngquant.exe not found in downloaded archive'
            }
        }
    }
    catch {
        throw ('Unable to download or extract pngquant.exe: {0}' -f $_.Exception.Message)
    }
}

if( -not (Test-Path $PNGQUANT)) {
    throw ('pngquant.exe not found at {0}' -f $PNGQUANT)
}

Write-Host ('Using pngquant.exe at {0}' -f $PNGQUANT) -ForegroundColor Green

$ProgressPreference = 'SilentlyContinue'

# Get all PPTX files from source path
$pptxFiles = Get-ChildItem -Path $SourcePath -Recurse -Filter *.pptx -ErrorAction SilentlyContinue

foreach( $pptxFile in $pptxFiles) {
    Write-Output ''
    Write-Output '=========================================='
    Write-Output ('Processing PPTX: {0}' -f $pptxFile.FullName)
    Write-Output ('Using Image Quality: {0}%, Video Target: {1}x{2}px, CRF: {3}, Preset: {4}' -f $ImageQuality, $TargetWidth, $TargetHeight, $CRF, $Preset)
    Write-Output '=========================================='
    Write-Output ''

    $absPath = $pptxFile.FullName
    $tempDir = Join-Path $env:TEMP ('pptx_recompress_{0}' -f [guid]::NewGuid().Guid)
    New-Item -ItemType Directory -Path $tempDir | Out-Null

    $ffmpegAvailable = Test-Path $FFMPEG
    $pngquantAvailable = Test-Path $PNGQUANT

$success= $false
try {
    # Create a working copy
    $workingCopy = '{0}.working' -f $absPath
    Copy-Item -LiteralPath $absPath -Destination $workingCopy -Force

    # Expand PPTX
    Expand-Archive -Path $workingCopy -DestinationPath $tempDir -Force

    $mediaDir = Join-Path $tempDir 'ppt\media'
    if( -not (Test-Path $mediaDir)) {
        Write-Output 'No embedded media found in PPTX.'
    } else {
        $files = Get-ChildItem -Path $mediaDir -File
        foreach( $f in $files) {
            $orig = $f.FullName
            $ext = $f.Extension.ToLower()
            $tmp = "$orig.tmp"

            $processed = $false

            # Process JPEG files
            if( $ext -eq '.jpg' -or $ext -eq '.jpeg') {
                Write-Output ('Processing JPEG: {0} (Quality: {1}%)' -f $f.Name, $ImageQuality)
                if( $ffmpegAvailable) {
                    $ok = Recompress-JpegWithBitmap -SrcPath $orig -DstPath $tmp -Quality $ImageQuality
                    $processed = $true
                } else {
                    Write-Warning ('FFmpeg not found. Skipping JPEG file: {0}' -f $f.Name)
                    continue
                }
            }
            # Process PNG files
            elseif( $ext -eq '.png') {
                $QualityRange= '{0}-{1}' -f (-10 + $ImageQuality), (10 + $ImageQuality)
                Write-Output ('Processing PNG: {0} (Range quality: {1}%)' -f $f.Name, $QualityRange)
                if( $pngquantAvailable) {
                    $ok = Recompress-PngWithPngquant -SrcPath $orig -DstPath $tmp -QualityRange $QualityRange -PNGQUANTPath $PNGQUANT
                    $processed = $true
                } else {
                    Write-Warning ('pngquant not found. Skipping PNG file: {0}' -f $f.Name)
                    continue
                }
            }
            # Process video files
            elseif( $ext -eq '.mp4' -or $ext -eq '.wmv' -or $ext -eq '.avi' -or $ext -eq '.mov') {
                if( $ffmpegAvailable) {
                    # Get video properties to check dimensions
                    $videoProps = Get-VideoProperties -FilePath $orig -FFMPEGPath $FFMPEG
                    
                    # Check minimum dimensions
                    $shouldProcess = $true
                    
                    if( $MinimumHeight -and $videoProps.Height -lt $MinimumHeight) {
                        Write-Output ('Skipping {0}: height ({1}) below minimum ({2})' -f $f.Name, $videoProps.Height, $MinimumHeight)
                        $shouldProcess = $false
                    }
                    
                    if( $MinimumWidth -and $videoProps.Width -lt $MinimumWidth) {
                        Write-Output ('Skipping {0}: width ({1}) below minimum ({2})' -f $f.Name, $videoProps.Width, $MinimumWidth)
                        $shouldProcess = $false
                    }
                    
                    if( $shouldProcess) {
                        Write-Output ('Processing video: {0} (Current: {1}x{2}, Target: {3}x{4}, CRF: {5}, Preset: {6})' -f $f.Name, $videoProps.Width, $videoProps.Height, $TargetWidth, $TargetHeight, $CRF, $Preset)
                        $ok = Reencode-MediaWithFFmpeg -SrcPath $orig -DstPath $tmp -TargetHeight $TargetHeight -TargetWidth $TargetWidth -CRF $CRF -Preset $Preset -Extension $ext -FFMPEGPath $FFMPEG
                        $processed = $true
                    }
                } else {
                    Write-Warning ('FFmpeg not found. Skipping {0} file: {1}' -f $ext, $f.Name)
                    continue
                }
            }
            else {
                Write-Output ('Skipping unsupported file: {0}' -f $f.Name)
                continue
            }

            # Replace original file if processing was successful
            if( $processed) {
                if( $ok -and (Test-Path $tmp)) {
                    $source= Get-ChildItem -Path $orig
                    $target= Get-ChildItem -Path $tmp
                    $diffsize= ($source.Length - $target.Length) / 1kb
                    If( $diffSize -gt 0) {
                        try {
                            $percentSaved = [math]::Round( ($diffsize * 1kb / $source.Length) * 100, 1)
                            Write-Output ('Replacing {0}, saved {1} KB ({2}%)' -f $f.Name, [int]$diffsize, $percentSaved)
                            Move-Item -LiteralPath $tmp -Destination $orig -Force
                        } catch {
                            Write-Warning ('Failed to replace {0}: {1}' -f $f.Name, $_)
                            Remove-Item -LiteralPath $tmp -ErrorAction SilentlyContinue
                        }
                    } else {
                        Write-Warning ('{0} size would increase ({1} KB), not replacing' -f $f.Name, [int]$diffsize)
                        Remove-Item -LiteralPath $orig -ErrorAction SilentlyContinue
                    }
                } else {
                    Write-Warning ('Recompression failed for {0} - keeping original' -f $f.Name)
                    Remove-Item -LiteralPath $tmp -ErrorAction SilentlyContinue
                }
            }
        }

        # Now update the working copy with the compressed media
        Write-Output 'Updating PPTX with compressed media...'
        $updateSuccess = Update-PptxMedia -PptxPath $workingCopy -MediaDir $mediaDir

        if( -not $updateSuccess) {
            Write-Error 'Failed to update PPTX with compressed media'
            exit 6
        }
    }

    # Replace original file only if compressed version is smaller
    try {
        $oldFile= Get-ChildItem -Path $absPath -erroraction silentlycontinue
        $newFile= Get-ChildItem -Path $workingCopy -ErrorAction silentlycontinue

        $diffsize= [math]::round( ($oldFile.Length - $newFile.Length) / 1MB, 2)

        if( $diffsize -gt 0) {
            $ct= $oldfile.creationTime
            $lwt= $oldfile.lastWriteTime

             $percentSaved = [math]::Round( ($diffsize * 1kb / $oldFile.Length) * 100, 1)
            Write-Host ('Saved {0} MB ({1}%)' -f $diffsize, $percentSaved) -ForegroundColor Green

            Write-Output ('Setting CreationTime to {0}, LastWriteTime to {1}' -f  $ct, $lwt)
            $newFile.creationTime = $ct
            $newFile.lastWriteTime = $lwt

            If( $Backup) {
                Write-Output ('Backing up original {0} to {1}.bak' -f $absPath, $absPath)
                Move-Item -LiteralPath $absPath -Destination ('{0}.bak' -f $absPath) -Force
            }

            Write-Output ('Moving compressed {0} to {1}' -f $workingCopy, $absPath)
            Move-Item -LiteralPath $workingCopy -Destination $absPath -Force

            Write-Output ('Saved recompressed PPTX to: {0}' -f $absPath)
            $success= $true
        }
        else {
            Write-Host ('Recompressed file is not smaller (difference: {0} MB), keeping original' -f $diffsize) -ForegroundColor Yellow
            Remove-Item -LiteralPath $workingCopy -Force -ErrorAction SilentlyContinue
            $success= $false
        }

    } catch {
        Write-Error ('Failed to overwrite original PPTX: {0}' -f $_)
        if( Test-Path $workingCopy) { Remove-Item -LiteralPath $workingCopy -Force }
    }
}
Catch {
    Write-Error ('An error occurred: {0}' -f $_)
}
finally {
    # Cleanup
    if( Test-Path $tempDir) {
        Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}
}

