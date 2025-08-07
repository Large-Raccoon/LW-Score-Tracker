#####################################################################################
## ------ Configuration ------
#####################################################################################

# --- Module Directory ---
$script:modules = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'modules'

# --- Import Module(s) ---
Import-Module "$script:modules\AdbTools.psm1" -Force
Import-Module "$script:modules\WindowsTools.psm1" -Force

# --- Tessdata Source ---
$script:TessdataSource = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'config\tessdata'
    
# Get configuration settings from config.json
$ConfigPath = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'config\config.json'
$script:Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json

# Get resolution configuration settings from resolution-config.json
$ResConfigPath = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'config\resolution-config.json'
$script:ResConfig = Get-Content $ResConfigPath -Raw | ConvertFrom-Json

# Generate full paths to prerequisite files
$magick = Get-ProductExecutable -ProductName 'ImageMagick' -FileName 'magick.exe'
$tesseract = Get-ProductExecutable -ProductName 'Tesseract-OCR' -FileName 'tesseract.exe'
$python = Get-ProductExecutable -ProductName 'Python' -FileName 'Python.exe'
$cv2 = Join-Path -Path $modules -ChildPath $script:Config.Core.Cv2Script

# Ensure traineddata has been added to tessdata
$TessPath = (Join-Path (Get-ProductExecutable -ProductName 'Tesseract-OCR' -PathOnly) 'tessdata')
$TrainedData = (Get-ChildItem -Path $TessdataSource -Filter '*.traineddata').Name

ForEach ($File in $TrainedData) {
    $TessdataDestination = Join-Path $TessPath $File -EA SilentlyContinue

    if (Test-Path $TessdataDestination -EA SilentlyContinue) {
        $DestinationFileHash = Get-FileHash $TessdataDestination -Algorithm SHA256
        $SourceFileHash = Get-FileHash (Join-Path $TessdataSource $File) -Algorithm SHA256
        if ($SourceFileHash -ne $DestinationFileHash) {
            Copy-Item -Path (Join-Path $TessdataSource $File) -Destination $TessdataDestination -Force
        }
    }
}

#####################################################################################
## ------ Function: Multithread Helper ------
#####################################################################################

Function Invoke-Multithreading {
param (
    [Parameter(Mandatory = $true)]
    [ScriptBlock]$ScriptBlock,
    [Parameter(Mandatory = $true)]
    [Array]$ArgumentList,  # must be a jagged array
    [int]$MaxThreads = 6,
    [switch]$Quiet,
    [string]$JobName
)
    
    Write-Host "JOB EXECUTION: $JobName ($MaxThreads threads)" -ForegroundColor Cyan

    $totalJobs = $ArgumentList.Count
    $completedJobs = 0

    # Create and open a runspace pool
    $runspacePool = [runspacefactory]::CreateRunspacePool(1, $MaxThreads)
    $runspacePool.Open()

    # Jobs collection
    [System.Collections.ArrayList]$jobs = @()

    foreach ($params in $ArgumentList) {
        $ps = [powershell]::Create()
        $ps.RunspacePool = $runspacePool

        # Wrap the script block to allow argument injection
        $wrappedScript = {
            param ($innerScript, $argValues)
            try {
                & ([scriptblock]::Create($innerScript)).Invoke($argValues) *>$null
            } catch {
                Write-Error "Error in parallel task: $_"
                }
            }

        $ps.AddScript($wrappedScript).
            AddArgument($ScriptBlock.ToString()).
            AddArgument($params) | Out-Null

        $null = $jobs.Add([PSCustomObject]@{
            Pipe  = $ps
            Async = $ps.BeginInvoke()
            })
        }

    # Wait for all jobs to finish and update the progress bar
    foreach ($job in $jobs) {
        try {
            $null = $job.Pipe.EndInvoke($job.Async)
        } catch {
            Write-Error "Error in runspace: $_"
        } finally {
            $job.Pipe.Dispose()
        }

        $completedJobs++
        if (-not $Quiet) {
            $percentComplete = [math]::Round(($completedJobs / $totalJobs) * 100, 0)
            Write-Progress -Activity "Job: $JobName" `
                            -Status "Status: $completedJobs of $totalJobs" `
                            -PercentComplete $percentComplete
        }
    }

    $runspacePool.Close()
    $runspacePool.Dispose()

    if (-not $Quiet) {
        Write-Progress -Activity "Job: $JobName" -Completed
        Write-Host "Status: Completed`n" -ForegroundColor Green
    }
}


#####################################################################################
## ------ Functions: Image Processing ------
#####################################################################################

Function Import-Image {
param (
    [Parameter(Mandatory = $true)]
    [string]$Source,
    [string]$Destination,
    [ValidateSet("VS", "TD", "KS")]
    [string]$Type
    )

    # Create destination if it doesn't exist
    if (-not (Test-Path $Destination)) {
        New-Item -ItemType Directory -Path $Destination | Out-Null
    }

    # Get all PNG files, sort by LastWriteTime (oldest first)
    $pngFiles = Get-ChildItem -Path $Source -Filter *.png | Sort-Object LastWriteTime

    # Rename and move in order
    $counter = 1
    Try {
        foreach ($file in $pngFiles) {
            if ($Type -eq 'VS') { $newFileName = "VS_Day1_{0:D3}.png" -f $counter }
            if ($Type -eq 'TD') { $newFileName = "TD_{0:D3}.png" -f $counter }
            if ($Type -eq 'KS') { $newFileName = "KS_{0:D3}.png" -f $counter }

            $destinationPath = Join-Path $Destination $newFileName
            Move-Item -Path $file.FullName -Destination $destinationPath -Force
            $counter++
            }
    } Catch {
        Write-Host "ERROR: Failed to import image(s)." -ForegroundColor Red
        exit 1
        }
    Write-Host "Images moved and renamed successfully!" -ForegroundColor Green
}


Function Get-Screenshots {
param (
    [Parameter(Mandatory=$true)]  [string]$SavePath,
    [Parameter(Mandatory=$true)]  [string]$SaveName,
    [Parameter(Mandatory=$false)]  [int]$LoopCount,
    [Parameter(Mandatory=$true)]  [string]$Type,
    [Parameter(Mandatory=$false)] [string]$ScrollSetting, # Setting to weekly uses weekly scrolling parameters
    [Parameter(Mandatory=$false)] [switch]$Manual,
    [Parameter(Mandatory=$false)]  [int]$PPS,
    [Parameter(Mandatory=$false)]  [int]$PlayerCount,
    [Parameter(Mandatory=$false)]  [string]$Day
    )

    # Initialize counter
    [int]$counter = 1

    if (!($Manual)) {
        do {
            # Format filename
            $TripleDigit = $counter.ToString("D3")
            $FileName    = "$($SaveName)_$TripleDigit.png"

            # Capture screenshot
            Get-AdbScreenshot -SavePath $SavePath -SaveName $FileName

            # Scroll only if this isn’t the last screenshot
            if ($counter -lt $loopCount) {
                if ($Type -eq 'VS' -and $ScrollSetting -ne 'Weekly') {
                    Invoke-AdbSwipe `
                        -StartX 600 -StartY 1400 `
                        -EndX 600   -EndY $Config.ADB.VsSwipeDistance `
                        -Duration 3250
                }
                elseif ($Type -eq 'VS' -and $ScrollSetting -eq 'Weekly') {
                    Invoke-AdbSwipe `
                        -StartX 600 -StartY 1400 `
                        -EndX 600   -EndY $Config.ADB.VsWklySwipeDistance `
                        -Duration 3250
                }
                elseif ($Type -in 'TD','KS') {
                    Invoke-AdbSwipe `
                        -StartX 600 -StartY 1400 `
                        -EndX 600   -EndY $Config.ADB.TdSwipeDistance `
                        -Duration 3250
                }
                # Delay to avoid tap animation
                Start-Sleep -Seconds 1
            }

            # Update progress bar
            $percentComplete = [math]::Round(($counter / $loopCount) * 100, 0)
            Write-Progress `
                -Activity "Job: Capture $Type Scoreboard" `
                -Status   "Status: $counter of $loopCount" `
                -PercentComplete $percentComplete

            # Increment for next iteration
            $counter++

        } while ($counter -le $loopCount)
    }

if ($Manual) {
    explorer.exe $SavePath
    :mainLoop while ($true) {
        $counter = 1  # reset counter at start of day
        for ($i = 1; $i -le $PlayerCount; $i += $PPS) {
            $startRank = $i
            $endRank   = [math]::Min($i + $PPS - 1, $PlayerCount)

            $TripleDigit = $counter.ToString('D3')
            $FileName    = "$SaveName`_$TripleDigit.png"

            if ($lastStartRank -notmatch '\d') { $lastStartRank = '' }
            if ($lastEndRank -notmatch '\d') { $lastEndRank = 'N/A' }

            $stopAll = $false

            while ($true) {
                Clear-Host
                if ($Type -eq 'VS' -and $Day) { $ScreenTitleType = "$Type Day $Day" }
                else { $ScreenTitleType = "$Type" }
                Write-Host "$ScreenTitleType Manual Screenshot Capture`n"
                Write-Host "Last screenshot taken: $LastFileName"
                Write-Host "Next screenshot to take: $FileName`n"
                Write-Host "=================================================="
                Write-Host "[1] Take screenshot of ranks " -NoNewLine; Write-Host "$startRank–$endRank" -ForegroundColor Cyan
                Write-Host "[2] Retake screenshot of previous ranks " -NoNewLine; Write-Host "$lastStartRank-$lastEndRank" -ForegroundColor Yellow
                Write-Host "[3] Stop taking screenshots"
                Write-Host "[Q] Quit the script`n"

                $choice = (Read-Host 'Enter an option from above').Trim().ToUpper()

                if ($choice -notin '1','2','3','Q') {
                    Write-Host "`nERROR: Invalid input`n" -ForegroundColor Red
                    continue
                }

                if ($choice -eq '2') {
                    Get-AdbScreenshot -SavePath $SavePath -SaveName $LastFileName
                    continue
                }

                if ($choice -eq '1') {
                    Get-AdbScreenshot -SavePath $SavePath -SaveName $FileName
                    
                    $lastStartRank = $startRank
                    $lastEndRank = $endRank
                    $LastFileName = $FileName
                    $counter++

                    if ($endRank -eq $PlayerCount) {
                        while ($true) {
                            Clear-Host
                            if ($Type -eq 'VS' -and $Day) { $ScreenTitleType = "$Type Day $Day" }
                            else { $ScreenTitleType = "$Type" }
                            Write-Host "$ScreenTitleType Manual Screenshot Capture`n"
                            Write-Host "Last screenshot taken: $LastFileName`n"
                            Write-Host "Next screenshot to take: $FileName"
                            Write-Host "=================================================="
                            Write-Host "[1] Done taking screenshots"
                            Write-Host "[2] Retake screenshot of previous $PPS ranks"
                            Write-Host "[3] Retake ALL screenshots for $ScreenTitleType"
                            Write-Host "[Q] Quit script`n"

                            $finalChoice = (Read-Host 'Enter an option from above').Trim().ToUpper()

                            switch ($finalChoice) {
                                '1' {
                                    $stopAll = $true
                                    break
                                }
                                '2' {
                                    Get-AdbScreenshot -SavePath $SavePath -SaveName $FileName
                                }
                                '3' {
                                    Write-Host "`nRestarting all screenshots for Day $Day...`n"
                                    continue mainLoop  # properly restarts the whole loop
                                }
                                'Q' {
                                    exit
                                }
                                default {
                                    Write-Host "`nERROR: Invalid input`n" -ForegroundColor Red
                                }
                            }
                        }
                    }
                }
                elseif ($choice -eq '3') {
                    $stopAll = $true
                }
                elseif ($choice -eq 'Q') {
                    exit
                }

                break
            }

            if ($stopAll) {
                break
            }
        }

        break  # exits the outer `while ($true)` after full successful run
    }
}
}

Function Invoke-ResizeTemplate {
param (
    [string]$InputImg,
    [string]$OutputImg,
    [string]$SourceImg,
    [int]$TargetWidth = 0,
    [int]$TargetHeight = 0
    )

    $objName = $inputImg.Split('\')[-1]

    #If the output directory does not exist, create it
    $OutputPath  = Split-Path $OutputImg -Parent
    If (!(Test-Path -Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath }

    # Get template dimensions
    $img = [System.Drawing.Image]::FromFile($InputImg)
    [int]$TemplateWidth = $img.Width
    [int]$TemplateHeight = $img.Height
    $img.Dispose()

    # Get source image dimensions
    $img = [System.Drawing.Image]::FromFile($SourceImg)
    [int]$SourceWidth = $img.Width
    #[int]$SourceHeight = $img.Height
    $img.Dispose()

    # Get scaled  template dimensions if exists
    if (Test-Path -Path $OutputImg) {
        $img = [System.Drawing.Image]::FromFile($OutputImg)
        [int]$ScaledTemplateWidth = $img.Width
        #[int]$ScaledTemplateHeight = $img.Height
        $img.Dispose()
        }

    # Dynamically determine the reference base width
    $baseWidth = switch ($SourceWidth) {
        #{ $_ -lt 800 }     { 1065; break }
        { $_ -lt 1300 }    { 1070; break }
        { $_ -lt 1600 }    { 1075; break }
        { $_ -lt 1900 }    { 1300; break }
        { $_ -lt 2200 }    { 1400; break }
        { $_ -lt 2500 }    { 3350; break }
        default            { 1100 } # fallback
    }

    # Allow override if explicit dimensions are passed
    if (($TargetWidth -eq 0) -and ($TargetHeight -eq 0)) {
        $scaleFactor = $SourceWidth / $baseWidth
        $TargetWidth = [math]::Round($TemplateWidth * $scaleFactor)
        $TargetHeight = [math]::Round($TemplateHeight * $scaleFactor)
    }

    # Only scale the templates if a suitable scaled template does not exist
    if ($TargetWidth -ne $ScaledTemplateWidth) {
        Write-Host "Scaling template $objName to ${TargetWidth}x${TargetHeight}"
        & $magick "$InputImg" -resize "${TargetWidth}x${TargetHeight}!" "$OutputImg"
    }
}


Function Get-ResolutionConfig {
param (
    [int]$ImgWidth,
    [int]$ImgHeight,
    [hashtable]$defaults = @{
            OffsetX = 210
            OffsetY = -65
            OffsetCW = 675
            OffsetCH = 125
            }
    )

    # Try to find an exact resolution match
    $exactMatch = $script:ResConfig | Where-Object { $_.Width -eq $ImgWidth -and $_.Height -eq $ImgHeight }

    if ($exactMatch) { return $exactMatch }

    # No exact match, find the closest based on Euclidean distance
    Write-Warning "No exact match for resolution ${ImgWidth}x${ImgHeight}. Finding closest fallback..."

    $closest = $null
    $minDistance = [double]::MaxValue

    foreach ($item in $script:ResConfig) {
        $dx = $ImgWidth - $Item.Width
        $dy = $ImgHeight - $Item.Height
        $distance = [math]::Sqrt(($dx * $dx) + ($dy * $dy))

        if ($distance -lt $minDistance) {
            $minDistance = $distance
            $closest = $item
        }
    }

    if ($closest) {
        Write-Warning "Using closest match: ${($closest.Width)}x${($closest.Height)} (distance: $([math]::Round($minDistance, 2)))"
        return $closest
    }

    # Fallback to defaults if all else fails.
    Write-Warning "No suitable fallback found. Using defaults."
    return $defaults
}


Function Get-Coords {
    param (
        [Parameter(Mandatory = $true)]
        [string]$refimg,
        [string]$srcimg
        )
    # Craft variable that will contain --threshold parameter
    [double]$Cv2Thresh = $Config.Core.Cv2Thresh
    $thresholdParam = "--threshold=$Cv2Thresh"

    # Form command line based on whether -thresh param was used or not.
    if ($Config.Core.Cv2Debug -eq '1') {
        $params = @(
            $cv2,
            $refimg,
            $srcimg,
            $thresholdParam
            "--debug"
        )
    } else {
        $params = @(
            $cv2,
            $refimg,
            $srcimg,
            $thresholdParam
        )
    }
    
    # Run python cv2 script with crafted parameters 
    $result = & $python @params

    # If the result was not $null, calculate center of match and return the coordinates.
    if ($null -ne $result) {
        $xylist = $result -split ','
        if ($xylist.Length -lt 2) { return $null }

        # Get image size
        $img = [System.Drawing.Image]::FromFile($refImg)
        [int]$refImgWidth = $img.Width
        [int]$refImgHeight = $img.Height
        $img.Dispose()

        # Calculate the center of match
        [int]$script:xMatchStart = $xylist[0]
        [int]$script:yMatchStart = $xylist[1]
        [int]$script:xMatchEnd = ($xMatchStart + $refImgWidth)
        [int]$script:yMatchEnd = ($yMatchStart + $refImgHeight)
        [int]$script:xCenterMatch = (($refImgWidth / 2) + $xMatchStart)
        [int]$script:yCenterMatch = (($refImgHeight / 2) + $yMatchStart)

        return [PSCustomObject]@{
            x = $xCenterMatch
            y = $yCenterMatch
        }
    } else {
        return $null
    }
}

Function Get-PPS {
Param (
    [array]$ReferenceImages,
    [string]$ScreenPath,
    [string]$Type
    )

    $srcImg = (Get-ChildItem "$ScreenPath\$Type" -Filter '*.png' | Sort-Object Name).FullName | Where-Object {$_ -notlike '*_debug.png'} | Select-Object -First 1
    $PPS = 0

    if ($null -ne $srcImg) {
        foreach ($ref in $ReferenceImages) {
            # Reset coordinates to avoid false positives
            $match = $null
            $xCenter = 0
            $yCenter = 0
            
            # Check for matching coordinates. -thresh 0.90 should avoid false positives
            $match = Get-Coords -refimg $ref -srcimg $srcImg

            # If a match was found, assign x and y to individual variables
            if ($null -ne $match) {
                $xCenter = $match.x
                $yCenter = $match.y

                # Make absolutely sure we have a match and calculate PPS size accordingly.
                if ((($xCenter -ne 0) -and ($yCenter -ne 0)) -or ($null -ne $xCenter) -and ($null -ne $yCenter)) {
                    $PPS++
                }
            }
        }
    }
    return $PPS
}

Function Use-ProcessImage {
    param (
        [Parameter(Mandatory = $true)]
        [string]$imgPath,
        [string]$Type
        )

    $objName = Split-Path $imgPath -Leaf

    $arguments = @(
        "`"$imgPath`""
        "stdout" # Output to standard output so we can store this in memory instead of a text file.
        "-l `"$($Config.ImageProcessing.Languages)`""
        "--oem 1"
        "--psm 6"
        )

    # Define process info. This allows us to write output into memory via stdout instead of a text file.
    $processInfo = New-Object System.Diagnostics.ProcessStartInfo
    $processInfo.FileName = "$tesseract"
    $processInfo.Arguments = $arguments -join " "
    $processInfo.RedirectStandardOutput = $true
    $processInfo.UseShellExecute = $false
    $processInfo.CreateNoWindow = $true
    $processInfo.StandardOutputEncoding = [System.Text.Encoding]::UTF8

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $processInfo
    $process.Start() | Out-Null

    $ocrText = $process.StandardOutput.ReadToEnd()
    $process.WaitForExit()

    <# For debugging only
    Write-Host "OCR Text:"
    $ocrText | Out-Host
    Write-Host ""
    #>

    # Split and trim lines
    $lines = $ocrText -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }

    If ($Type -eq 'VS') {
        # Define score pattern (OCR friendly by allowing periods and long commas)
        $scorePattern = '^\d{1,3}(?:[,\.\uFF0C\uFF0E\u00A0\u202F\s]\d{3})+$'

        # Find the first line matching the score pattern
        $scoreLine = $lines |
            Where-Object { $_ -match $scorePattern } |
            Select-Object -First 1

        # Replace periods and long commas with a comma, then strip everything but digits & commas
        if ($scoreLine) {
            $score = $scoreLine `
                -replace '[\.\uFF0C\uFF0E\u00A0\u202F\s]', ',' `
                -replace '[^\d,]', ''
        }
        else {
            $score = ''
        }
    }

    If ($Type -in @('TD', 'KS')) {
        # Score is the first line that matches score format
        $scorePattern = '^\d+$'

        # Score is the first line that matches score format
        $scoreLine = $lines | Where-Object { $_ -match $scorePattern } | Select-Object -First 1
        $score = if ($scoreLine) { $scoreLine -replace '[^\d]', '' } else { "" }
    }

    # Player is first line that does NOT match the score format
    $playerLine = $lines | Where-Object { $_ -notmatch $scorePattern } | Select-Object -First 1
    $player = if ($playerLine) { ($playerLine -replace '[^\p{L}\p{N} ]', '').Trim() } else { "" }

    # Define the rank
    $rank = if ($objName -match '(\d{3})\.png$') { [int]$matches[1] } else { 'N/A' }

    return [PSCustomObject]@{
        'Rank'   = $rank
        'Player' = $player
        'Score'  = $score
    }
}

Function Invoke-MTCropping {
param (
    [array]$Jobs
    )

    # Now run jobs after the batch
    if ($Jobs.Count -gt 0) {
        Invoke-Multithreading -ScriptBlock {
            Param ($SrcImg, $CroppedImg, $CropX, $CropY, $CropW, $CropH, $magick)

            Function Use-CropImage {
            param (
                [string]$SrcImg,
                [string]$CroppedImg,
                [int]$CropX,
                [int]$CropY,
                [int]$CropW,
                [int]$CropH,
                [string]$magick
                )

                # Optional debug output
                Write-Output "Cropping $SrcImg X:$cropX Y:$cropY Size:${cropW}x${cropH}"

                # Build and run ImageMagick crop command
                $cropArgs = "${cropW}x${cropH}+$cropX+$cropY"
                & $magick "$SrcImg" -crop "$cropArgs" +repage "$CroppedImg"
            }
            Use-CropImage -SrcImg "$SrcImg" -CroppedImg "$CroppedImg" -CropX $CropX -CropY $CropY -CropW $CropW -CropH $CropH -magick $magick
        } -ArgumentList $Jobs -MaxThreads $Config.Core.Threads -JobName 'Crop Score Images'
    }
}


Function Invoke-MTPreprocessing {
param (
    [array]$Jobs
    )

    # Now run jobs after the batch
    if ($Jobs.Count -gt 0) {
        Invoke-Multithreading -ScriptBlock {
            Param ($imgPath, $imgDest, $magick, $Type)

            function Use-PreprocessImage {
                Param (
                    [string]$imgPath,
                    [string]$imgDest,
                    [string]$magick,
                    [string]$Type
                    )

                # Assign a unique guid to the namespace in memory
                $scoreName = "mpr:score_$([guid]::NewGuid().ToString("N"))"
                $playerName  = "mpr:name_$([guid]::NewGuid().ToString("N"))"

                # VS preprocessing parameters
                if ($Type -eq 'VS') {
                    $params = @(
                        "$imgPath",
                        
                        # Resize, convert to grayscale, remove the pfp from image, artifically set DPI to 300, ease up the bold font
                        '-resize', '945x175!',
                        '-monochrome',
                        '-fill', 'white',
                        '-draw', 'rectangle 0,80 550,550',
                        '-units', 'PixelsPerInch',
                        '-density', '300',
                        '-morphology', 'Dilate', 'Square:1',
                        '-bordercolor', 'White', '-border', '10x10',

                        # Move score to the left
                        '(', '-clone', '0', '-crop', '600x100+550+60', '+repage', '-write', $scoreName, '+delete', ')',
                        '-fill', 'white', '-draw', 'rectangle 550,60 1150,160',
                        $scoreName, '-geometry', '+0+100', '-composite',

                        # Move player name to the right
                        '(', '-clone', '0', '-crop', '600x100+0+0', '+repage', '-write', $playerName, '+delete', ')',
                        '-fill', 'white', '-draw', 'rectangle 0,0 600,100',
                        $playerName, '-geometry', '+100+0', '-composite',

                        # Remove transparency layer
                        '-background', 'white',
                        '-alpha', 'remove',

                        # Full path to destination of edited image
                        "$imgDest"
                        )
                    }

                # TD and KS preprocessing parameters
                if ($Type -in @('TD', 'KS')) {
                    $params = @(
                        "$imgPath",

                        # Resize, convert to grayscale, remove the pfp from image, artifically set DPI to 300, ease up the bold font
                        '-resize', '945x190!',
                        '-monochrome',
                        '-fill', 'white',
                        '-draw', 'rectangle 0,0 80,190',
                        '-units', 'PixelsPerInch',
                        '-density', '300',
                        '-morphology', 'Dilate', 'Square:1',
                        '-bordercolor', 'White', '-border', '10x10',
                        
                        # Move player name up and to the left
                        '(', '-clone', '0', '-crop', '585x90+70+30', '+repage', '-write', $playerName, '+delete', ')',
                        '-fill', 'white', '-draw', 'rectangle 0,0 645,190',
                        $playerName, '-geometry', '+40-10', '-composite',

                        # Move score down to the left
                        '(', '-clone', '0', '-crop', '260x65+675+40', '+repage', '-write', $scoreName, '+delete', ')',
                        '-fill', 'white', '-draw', 'rectangle 675,40 940,95',
                        $scoreName, '-geometry', '+20+80', '-composite',
                        
                        # Remove transparency layer
                        '-background', 'white',
                        '-alpha', 'remove',

                        # Full path to destination of edited image
                        "$imgDest"
                        )
                    }
                & $magick @params
                }
            Use-PreprocessImage -imgPath $imgPath -imgDest $imgDest -magick $magick -Type $Type
        } -ArgumentList $Jobs -MaxThreads $Config.Core.Threads -JobName 'Preprocess Images'
    }
}

Export-ModuleMember -Function *