#####################################################################################
## ------ Configuration ------
#####################################################################################

# --- Module Directory ---
$script:modules = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'modules'

# --- Import Module(s) ---
Import-Module "$script:modules\WindowsTools.psm1" -Force
Import-Module "$script:modules\ImageProcessing.psm1" -Force

# --- Configuration ---
# Generate full paths to prerequisite files
$magick = Get-ProductExecutable -ProductName ImageMagick -FileName magick.exe

# Get configuration settings from config.json
$script:ConfigPath = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'config\config.json'
$script:Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json

#####################################################################################
## ------ Functions: Job Collection ------
#####################################################################################
Function Get-Jobs {
param (
    [string]$Templates,
    [string]$scaledtemplates,
    [string]$screens,
    [string]$Preproc,
    [array]$day,
    [int]$PPS,
    [string]$Type
    )

    # Generate list of VS or Str (for TD and KS) ranking image templates that will be referenced
    If ($Type -eq 'VS') { $referenceImgs = (Get-ChildItem "$Templates\ranking" -Filter "VS_Rank*.png" | Sort-Object Name).FullName }
    If ($Type -in @('TD', 'KS')) { $referenceImgs = (Get-ChildItem "$Templates\ranking" -Filter "Str_Rank*.png" | Sort-Object Name).FullName }

    # Load a sample image from the screenshots
    $imgSample = (Get-ChildItem "$screens\$Type" -Filter "${Type}_*.png" | Sort-Object Name | Select-Object -First 1).FullName

    If (!($imgSample)) {
        Write-Host "ERROR: You don't have any $Type screenshots to process." -ForegroundColor Red
        exit
    }

    # Check image dimensions of sample image
    $img = [System.Drawing.Image]::FromFile($imgSample)
    $sampleWidth = $img.Width
    #$sampleHeight = $img.Height
    $img.Dispose()

    # Resize the templates proportionately to the screenshots if necessary
    If ($sampleWidth -notin '1200', '1080') {
        $referenceImgs | ForEach-Object {
            $imgName = Split-Path $_ -Leaf
            Invoke-ResizeTemplate -InputImg $_ -OutputImg "$scaledtemplates\ranking\$imgName" -SourceImg $imgSample
        }
        If ($Type -eq 'VS') { $referenceImgs = (Get-ChildItem "$scaledtemplates\ranking" -Filter "VS_Rank*.png" | Sort-Object Name).FullName }
        If ($Type -in @('TD', 'KS')) { $referenceImgs = (Get-ChildItem "$scaledtemplates\ranking" -Filter "Str_Rank*.png" | Sort-Object Name).FullName }
    }

    # Configure players per screenshot (set batch size)
    If ($PPS -eq '0' -or $null -eq $PPS) {
        $RefImgFirstTen = $referenceImgs | Select-Object -First 10
        $PPS = Get-PPS -ReferenceImages $refImgFirstTen -ScreenPath $Screens -Type $Type
        [int]$batchsize = $PPS
        Write-Host "PPS: $batchsize`n"
        } Else {
            [int]$batchsize = $PPS
            }
    
    # Generate list of VS screenshots for number of days specified in -Day parameter if -Type VS was specified.
    If ($Type -eq 'VS') {
        $ScreenshotLists = @{}
        
        # VS job collection for daily score requirements.
        if ($Config.Alliance.VS_DailyMin -ne '0') {
            foreach ($i in 1..6) {
                if ($day -contains "$i" -or $day -eq 'All') {
                    $ScreenshotLists["VS_Day$i"] = Get-ChildItem "$screens\VS" -Filter "VS_Day${i}_*" |
                        Sort-Object Name |
                        Where-Object { $_.FullName -notlike '*_debug.png' } |
                        Select-Object -ExpandProperty FullName
                }
            }
        }
        # VS job collection for weekly score requirements.
        if ($Config.Alliance.VS_WeeklyMin -ne '0') {
            $ScreenshotLists["$Type"] = Get-ChildItem "$screens\$Type" -Filter "$Type_*" |
                Sort-Object Name |
                Where-Object { $_.FullName -notlike '*_debug.png' } |
                Select-Object -ExpandProperty FullName
        }
    }

    # Generate list of Tech Donation or Kill Score screenshots if -Type TD or KS was specified.
    If ($Type -in @('TD', 'KS')) {
        $ScreenshotLists = @{}
        $ScreenshotLists["$Type"] = Get-ChildItem "$screens\$Type" -Filter "$Type_*" |
            Sort-Object Name |
            Where-Object { $_.FullName -notlike '*_debug.png' } |
            Select-Object -ExpandProperty FullName
        }


    # These hashtables will be universally used. Initialize them here.
    $CropJobs = @()
    $PreprocJobs = @()
    
    # GetEnumerator allows us to list key value pairs from a hash table
    foreach ($kvp in $ScreenshotLists.GetEnumerator() | Sort-Object Key) {
        $ListName = $kvp.Key
        $TargetList = $kvp.Value

        #Write-Host "Processing $ListName -> $($TargetList.Count) items"

        $refIndex = 0  # Tracks global reference index across all screenshots
        
        #Write-Host "Getting coordinates from:"
        [int]$ScreenshotNumber = 0
        foreach ($src in $TargetList) {
            #$StartRank = $refIndex+1
            #$EndRank = $PPS+$refIndex

            # Progress bar
            [int]$ScreenshotNumber++
            $percentComplete = [math]::Round(($ScreenshotNumber / $($TargetList.Count)) * 100, 0)
            Write-Progress -Activity "Job: Look for rank matches in $ListName" `
                            -Status "Status: $ScreenshotNumber of $($TargetList.Count)" `
                            -PercentComplete $percentComplete

            # End job collection once all applicable rank templates (referenceImgs) have finished image matching
            if ($refIndex -ge $referenceImgs.Count) { break }

            # Calculate remaining rank templates to image match and which will be selected for the batch
            $remaining = $referenceImgs.Count - $refIndex
            $batchSize = [Math]::Min($PPS, $remaining)
            $refBatch = $referenceImgs | Select-Object -Skip $refIndex -First $batchSize

            # Take current batch of rank templates and check for matches against scoreboard screenshot
            foreach ($ref in $refBatch) {
                $refname = Split-Path $ref -Leaf -Resolve
                $refname = $refname -replace '^VS', $ListName
                $srcname = Split-Path $src -Leaf -Resolve
                
                # Perform image matching
                Clear-Variable -Name match -Force -EA SilentlyContinue
                $match = Get-Coords -refimg "$ref" -srcimg "$src"
                
                # If $match did not return $null, consider the match found and add a job
                if ($null -ne $match) {
                    $xCenter = $($match.x)
                    $yCenter = $($match.y)

                    If ($Type -eq 'VS') {
                        $CropDestinationPath = Join-Path -Path $preproc -ChildPath $refname
                    }

                    If ($Type -eq 'TD') {
                        $modifiedName = $refname -replace '^[^_]+', 'TD'
                        $CropDestinationPath = Join-Path -Path $preproc -ChildPath $modifiedName
                    }

                    If ($Type -eq 'KS') {
                        $modifiedName = $refname -replace '^[^_]+', 'KS'
                        $CropDestinationPath = Join-Path -Path $preproc -ChildPath $modifiedName
                    }

                    $CropSourcePath = Join-Path -Path $preproc -ChildPath $refname

                    # Load image to get actual dimensions
                    $img = [System.Drawing.Image]::FromFile($src)
                    $actualWidth = $img.Width
                    $actualHeight = $img.Height
                    $img.Dispose()

                    # Get offsets from config
                    $offsets = Get-ResolutionConfig -imgWidth $actualWidth -imgHeight $actualHeight

                    # Apply scaled offsets relative to center of match
                    # Apply to crop location
                    $cropX = $Match.x + $offsets.OffsetX
                    $cropY = $Match.y + $offsets.OffsetY
                    $cropW = $offsets.RankRow_W
                    $cropH = $offsets.RankRow_H

                    # Add crop job arguments
                    $Cropjobs += ,@($src, $CropDestinationPath, $cropX, $cropY, $cropW, $cropH, $magick)

                    # Add preprocessing job arguments
                    $PreprocJobs += ,@($CropDestinationPath, $CropDestinationPath, $magick, $Type)
                    }
                }
                $refIndex += $PPS
            }
        }
    Write-Progress -Activity "Running Parallel Jobs" -Completed
    Write-Host "Status: Completed`n" -ForegroundColor Green
    return @{
    CropJobs  = $CropJobs
    PreprocJobs = $PreprocJobs
    ScreenshotLists = $ScreenshotLists
    }
}