<#
.SYNOPSIS
Gets scoreboard data from screenshots and writes to text for use with Excel or Google Sheets.

.DESCRIPTION
Allows users the ability to have screenshots of VS (Alliance Duel), TD (Alliance Tech Donations), 
and KS (Kill Score) processed into text through an OCR. Screenshots may be taken automatically or 
imported from manually taken screenshots.

.PARAMETER -Mode
Determine what scoreboard will be processed.

-Mode VS = Alliance Duel
-Mode TD = Tech Donations
-Mode KS - Kill Score

.PARAMETER -Day
Specifies which days of VS scores are to be captured.
Required when processing VS scores with daily score requirements. Not required for weekly score requirements.

-Day All = Processes all VS days 1-6
-Day 1,2,3,4 = Processes VS days 1-4
-Day 6 = Processes VS day 6 only

.PARAMETER -PPS
Specifies how many players are clearly depicted (with rank number fully intact) in each scoreboard screenshot.
Required when automatically collecting screenshots. Will attempt to automatically set when -NoBot is used.
For best results, always specify the PPS.

-PPS 5 = Specifies that 5 players and their ranks are clearly depicted in each screenshot

.PARAMETER -NoBot
Script will proceed without automatic screenshot capture. Must manually add screenshot to the import folder.

.EXAMPLE
Get VS scores from days 1-4 with automatic screenshot capture:
.\LW-ScoreTracker.ps1 -Mode VS -Day 1,2,3,4 -PPS 5

Get VS scores from days 1-4 with manually taken screenshots:
.\LW-ScoreTracker.ps1 -Mode VS -Day 1,2,3,4 -PPS 5 -NoBot

Get VS scores when your alliance has weekly score requirements with automatic screenshot capture:
.\LW-ScoreTracker.ps1 -Mode VS -PPS 5

Get weekly tech donation scores with automatic screenshot capture:
.\LW-ScoreTracker.ps1 -Mode TD -PPS 5

Get kill scores with automatic screenshot capture:
.\LW-ScoreTracker.ps1 -Mode KS -PPS 5

.NOTES
Version 1.0 - August 4th, 2025

########################
##---Prerequisites--- ##
########################
The following components are required to use this tool.

1. Microsoft Windows
Development and testing was all performed on Windows 11. Windows 10 should work as well.
- Minimum Version: Windows 10

2. Powershell
The bulk of this tool uses Powershell where version 5.1 is installed by default on Windws 10 and 11.
However, Powershell 7 provides a superior experience and can be acquired from the link below if desired.
- Source: https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.4
- Minimum Version: 5.1 (default in Windows 10 and 11)
- Recommended Version: 7.5.2 (or greater)

3. Python
A small portion of this tool requires Python in order to leverage OpenCV (cv2). This is crucial for getting 
coordinates of image matches.
- Source: https://www.python.org/downloads/
- Minimum Version: 3.6.0

4. OpenCV (cv2)
Image match recognition tool that will be used by Python for getting coordinates of image matches.
- Source: Installed from terminal. (Command: pip install opencv-contrib-python)
- Minimum Version: 4.12.0.88

5. ImageMagick
A Powershell/Windows friendly image processing tool. It is required for the preparation of images
which will be read by the OCR.
- Source: https://imagemagick.org/script/download.php
- Minimum Version: 7.1.1-47 Q16-HDRI (64-bit)

6. Android Debug Bridge (adb)
Copy your ADB files into C:\ADB, C:\Program Files\ADB, or add to PATH if saving elsewhere
- Source: https://developer.android.com/tools/releases/platform-tools
- Minimum Version: 36.0.0

7. Tesseract OCR
Reads text from images and converts to actual text format for processing. Required for reading player names
and scores from screenshots.
- Source: https://github.com/UB-Mannheim/tesseract/wiki
- Minimum Version: 5.0.4.2

##################################
## ---Optional Prerequisites--- ##
##################################

8. Google Sheets
- Google Sheets export on Powershell 5.1 requires a .p12 certificate and service account configured in 
config.json. Powershell 7 will support both .p12 certificate and json certificate options.

#>

#####################################################################################
## ------ Configuration ------
#####################################################################################

# Cmdlet Bindings
[CmdletBinding()]
Param(
    [Parameter(Mandatory=$false)]
    [ValidateScript( {($_ -match '^(all|([1-6](,[1-6])*)?)$')}) ]
    [array]$Day,
    [switch]$NoBot,
    [int]$PPS,
    [ValidateSet("VS", "TD", "KS")]
    [string]$Mode
    )

# --- Import Config File ---
$ConfigPath = Join-Path -Path $PSScriptRoot -ChildPath 'config\config.json'
$Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json

# --- Configuration ---
$RosterData = Import-Csv -Path "$PSScriptRoot\config\$($Config.Alliance.RosterName)" -Encoding UTF8

# --- Script Directories ---
$modules = "$PSScriptRoot\modules"
$ConfigDir = "$PSScriptRoot\config"
$screens = "$PSScriptRoot\imgproc\screens"
$templates = "$PSScriptRoot\config\templates"
$Reports = "$PSScriptRoot\reports"
$preproc = "$PSScriptRoot\imgproc\preproc"
$scaledtemplates = "$PSScriptRoot\imgproc\templates_scaled"
$Logs = "$PSScriptRoot\logs"
$import = "$PSScriptRoot\import"

# --- Dynamic Configurations ---
[int]$VsMinScore = if ($Config.Alliance.VS_DailyMin -ne '0') { $Config.Alliance.VS_DailyMin }
                elseif ($Config.Alliance.VS_WeeklyMin -ne '0') { $Config.Alliance.VS_WeeklyMin }
                else { '0' }

[int]$TdMinScore = if ($Config.Alliance.TD_WeeklyMin -ne '0') { $Config.Alliance.TD_WeeklyMin }

# Status info for the user
Clear-Host
Write-Host "Loading modules. Please wait..."

# --- Modules to import
Import-Module "$modules\WindowsTools.psm1" -Force
Import-Module "$modules\AdbTools.psm1" -Force
Import-Module "$modules\ImageProcessing.psm1" -Force
Import-Module "$modules\TextProcessing.psm1" -Force
Import-Module "$modules\JobCollection.psm1" -Force

# --- FUNCTION DEBUGGING ---
# (add lines here to quickly test module functions)

#####################################################################################
## ------ 1: Prepare the local environment ------
#####################################################################################

# Check if prerequisite folders exist. Create if not.
if (!(Test-Path $screens -EA SilentlyContinue)) { New-Item -ItemType Directory -Path $screens | Out-Null }
if (!(Test-Path "$screens\VS" -EA SilentlyContinue)) { New-Item -ItemType Directory -Path "$screens\VS" | Out-Null }
if (!(Test-Path "$screens\TD" -EA SilentlyContinue)) { New-Item -ItemType Directory -Path "$screens\TD" | Out-Null }
if (!(Test-Path "$screens\KS" -EA SilentlyContinue)) { New-Item -ItemType Directory -Path "$screens\KS" | Out-Null }
if (!(Test-Path "$screens\Nav" -EA SilentlyContinue)) { New-Item -ItemType Directory -Path "$screens\Nav" | Out-Null }
if (!(Test-Path $preproc -EA SilentlyContinue)) { New-Item -ItemType Directory -Path $preproc | Out-Null }
if (!(Test-Path $logs -EA SilentlyContinue)) { New-Item -ItemType Directory -Path $logs | Out-Null }
if (!(Test-Path $Reports -EA SilentlyContinue)) { New-Item -ItemType Directory -Path $Reports | Out-Null }
if (!(Test-Path "$import\KS" -EA SilentlyContinue)) { New-Item -ItemType Directory -Path "$import\KS" | Out-Null }
if (!(Test-Path "$import\TD" -EA SilentlyContinue)) { New-Item -ItemType Directory -Path "$import\TD" | Out-Null }
if (!(Test-Path "$import\VS" -EA SilentlyContinue)) { New-Item -ItemType Directory -Path "$import\VS" | Out-Null }

# Informational variables for the header
if ($NoBot) { $ScreenshotSource = 'Manual\Existing' } else { $ScreenshotSource = 'Auto' }
if ($null -eq $Day) { $HeaderDay = 'N/A' } else { $HeaderDay = $Day }
if ($Mode -contains 'VS') { [int]$HeaderScore = $VsMinScore }
if ($Mode -contains 'TD') { [int]$HeaderScore = $TdMinScore }
if ($Mode -contains 'VS' -and $Config.Alliance.VS_DailyMin -ne '0') { $ScoreFrequency = 'Daily' } else { $ScoreFrequency = 'Weekly' }
if ($Mode -contains 'TD' -and $Config.Alliance.TD_WeeklyMin -ne '0') { $ScoreFrequency = 'Weekly' }
if ($Mode -contains 'KS' -and $Config.Alliance.TD_WeeklyMin -ne '0') { $ScoreFrequency = 'N/A' }

# Format the score into an easy to read format.
$HeaderScoreFormatted = "{0:N0}" -f $HeaderScore

# Display the informational header
Clear-Host
Write-Host @(
"
======================================
    --- Last War Score Tracker ---    
======================================
Image Source: $ScreenshotSource
Roster File: $($Config.Alliance.RosterName)
Player Count: $($RosterData.Player.Count)

Scoreboard: $Mode
Day: $HeaderDay
Minimum Score: $HeaderScoreFormatted
Score Frequency: $ScoreFrequency

Script will begin in 10 seconds...
Press CTRL+C to cancel
======================================"
)

# Validate parameters and config.json
if (($Mode -eq 'VS') -and ($Config.Alliance.VS_DailyMin -ne '0') -and ($Config.Alliance.VS_WeeklyMin -ne '0')) {
    Write-Host "ERROR: You cannot set both daily and weekly requirements in config.json." -ForegroundColor Red
    Write-Host "Set VS_DailyMin or VS_WeeklyMin to 0 in config.json then try again."-ForegroundColor Magenta
    exit
}

# Enforce $Day parameter if $Mode is set to VS with VS_DailyMin greater than 0 in config.json
if ($Mode -contains 'VS' -and $Config.Alliance.VS_DailyMin -ne '0'-and -not $PSBoundParameters.ContainsKey('Day')) {
    Write-Host "ERROR: Parameter -Day is required when using -Mode VS with daily score requirements.`n" -ForegroundColor Red
    exit
}

# Enforce -PPS parameter if -NoBot was not supplied
if (($PPS -le '0' -or -not $PSBoundParameters.ContainsKey('PPS')) -and -not $PSBoundParameters.ContainsKey('NoBot')) {
    Write-Host "ERROR: Parameter -PPS with a valid number is required if -NoBot is not supplied.`n" -ForegroundColor Red
    Write-Host "To resolve, check how many ranks are clearly displayed without scrolling on your scoreboard." -ForegroundColor Red
    Write-Host "Add the number after -PPS. Example: -PPS 5" -ForegroundColor Red
    exit
}

# Enforce using a populated roster
if ($RosterData.Player.Count -eq 0) {
    Write-Host "ERROR: Your $($Config.Alliance.RosterName) file needs to be populated with your players.`n" -ForegroundColor Red
    exit
}

# Wait before proceeding in case user decides to cancel
Start-Sleep -Seconds 10

# Delete old screenshots
if (!($NoBot)) { 
    Get-ChildItem -Path "$screens\*" -File -Recurse -EA SilentlyContinue | Remove-Item -Force -EA SilentlyContinue
}

# Preproc should always be cleared before generating new images.
Remove-Item -Path "$preproc\*" -Recurse -Force -EA SilentlyContinue

# Remove logs and debug screenshots
Remove-Item -Path "$logs\*" -Recurse -Force -EA SilentlyContinue

#####################################################################################
## ------ 2: Import user's screenshots if running with -NoBot ------
#####################################################################################

if ($NoBot) {
    if ($Mode -contains 'VS') {
        # Set Type to match Mode for this loop
        $Type = @()
        [string]$Type = 'VS'

        Write-Host "IMAGE IMPORT: Check for new $Type screenshots to import" -ForegroundColor Cyan

        # Get a list of VS screenshots to import
        $ImportList = @()
        $ImportList = (Get-ChildItem "$import\$Type" -Filter "*.png" -EA SilentlyContinue | Sort-Object Name).FullName

        if (($ImportList).Count -gt '0') {
            $ExistingScreenshots = (Get-ChildItem "$screens\$Type" -Filter "$Type*.png" -EA SilentlyContinue).FullName
            $SampleScreen = (Get-ChildItem "$screens\$Type" -Filter "$Type*.png" -EA SilentlyContinue | Sort-Object Name).FullName | Select-Object -First 1
            $ImportHash = Get-FileHash -Path ($ImportList | Select-Object -First 1) -Algorithm SHA256 -EA SilentlyContinue
            if ($ExistingScreenshots.Count -ge 1) { $SampleHash = Get-FileHash -Path $SampleScreen -Algorithm SHA256 -EA SilentlyContinue }
            if ($ImportHash -ne $SampleHash) {
                Remove-Item -Path "$screens\$Type" -Recurse -Force -EA SilentlyContinue
                Import-Image -Source "$import\$Type" -Destination "$screens\$Type" -Mode $Type
                Write-Host "Status: Completed - Imported $($ImportList.Count) image(s).`n" -ForegroundColor Green
                }
        } Else {
            Write-Host "Status: Completed - No new $Type screenshot(s) detected.`n" -ForegroundColor Yellow
            #if 
        }
        }
    if ($Mode -contains 'TD') {
        # Set Type to match Mode for this loop
        $Type = @()
        [string]$Type = 'TD'

        Write-Host "IMAGE IMPORT: Check for new $Type screenshots to import" -ForegroundColor Cyan

        # Get a list of VS screenshots to import
        $ImportList = @()
        $ImportList = (Get-ChildItem "$import\$Type" -Filter "*.png" -EA SilentlyContinue | Sort-Object Name).FullName

        if (($ImportList).Count -gt '0') {
            $SampleScreen = (Get-ChildItem "$screens\$Type" -Filter "$Type*.png" -EA SilentlyContinue | Sort-Object Name).FullName | Select-Object -First 1
            $ImportHash = Get-FileHash -Path ($ImportList | Select-Object -First 1) -Algorithm SHA256
            $SampleHash = Get-FileHash -Path $SampleScreen -Algorithm SHA256
            if ($ImportHash -ne $SampleHash) {
                Write-Host "IMAGE IMPORT: Tech Donation ($Type) screenshot(s) detected. Importing $($ImportList.Count) image(s).`n" -ForegroundColor Green
                Remove-Item -Path "$screens\$Type" -Recurse -Force -EA SilentlyContinue
                Import-Image -Source "$import\$Type" -Destination "$screens\$Type" -Mode $Type
                }
            } Else {
                Write-Host "IMAGE IMPORT: No new $Type screenshot(s) detected.`n"
                }
        }

    if ($Mode -contains 'KS') {
        # Set Type to match Mode for this loop
        $Type = @()
        [string]$Type = 'KS'

        Write-Host "IMAGE IMPORT: Check for new $Type screenshots to import" -ForegroundColor Cyan

        # Get a list of VS screenshots to import
        $ImportList = @()
        $ImportList = (Get-ChildItem "$import\$Type" -Filter "*.png" -EA SilentlyContinue | Sort-Object Name).FullName

        if (($ImportList).Count -gt '0') {
            $SampleScreen = (Get-ChildItem "$screens\$Type" -Filter "$Type*.png" -EA SilentlyContinue | Sort-Object Name).FullName | Select-Object -First 1
            $ImportHash = Get-FileHash -Path ($ImportList | Select-Object -First 1) -Algorithm SHA256
            $SampleHash = Get-FileHash -Path $SampleScreen -Algorithm SHA256
            if ($ImportHash -ne $SampleHash) {
                Write-Host "IMAGE IMPORT: Kill Score ($Type) screenshot(s) detected. Importing $($ImportList.Count) image(s).`n" -ForegroundColor Green
                Remove-Item -Path "$screens\$Type" -Recurse -Force -EA SilentlyContinue
                Import-Image -Source "$import\$Type" -Destination "$screens\$Type" -Mode $Type
                }
            } Else {
                Write-Host "IMAGE IMPORT: No new $Type screenshot(s) detected.`n"
                }
        }
    }# end -NoBot image import loop

#####################################################################################
## ------ 3: BOT OPERATION: Collect Screenshots ------
#####################################################################################

# Only run this loop if -NoBot was NOT specified in params.
if (!($NoBot)) {

    Write-Host "INITIALIZATION: Prepare for navigation" -ForegroundColor Cyan

    # If Emulator is set to Bluestacks in config.json, confirm ADB is enabled in it. If not, exit.
    if ($Config.ADB.Emulator -eq 'Bluestacks') {
        $AdbSetting = Get-BluestacksAdbSetting -ConfigPath "$env:ProgramData\BlueStacks_nxt\bluestacks.conf"
        if (!($AdbSetting)) {
            Write-Host "ERROR: ADB is not enabled in Bluestacks. Enable then try again." -ForegroundColor Red
            exit
        }
    }

    # If using Bluestacks, confirm it's running. Launch it if not.
    if ($($Config.ADB.Emulator) -eq 'Bluestacks') {
        # Remove the file extension to get the process name
        $ProcessName = $Config.ADB.EmulatorExe -replace '\.[^.]+$', ''
        
        # Check if process is running. Launch if not.
        if (!(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue)) {
            Write-Host "$($Config.ADB.Emulator) is not running. Launching..."
            $Instance = Get-BlueStacksInstance
            $BlueStacks = Get-ProductExecutable -ProductName $Config.ADB.Emulator -FileName $Config.ADB.EmulatorExe

            # Build parameters and execute
            $params = @(
                '--instance'
                $Instance
            )
            & $BlueStacks $params

            # Wait 30 seconds to give BlueStacks some time to load.
            [int]$Delay = 30
            Write-Host "Waiting $Delay seconds before continuing."
            Start-Sleep -Seconds $Delay
        }
    }

    # Start ADB if not running and connect to device
    Start-AdbServer
    Connect-AdbDevice

    # Check the current app on screen and launch package if needed.
    Start-AdbPackage -Package $Config.ADB.Package

    # If we made it here, we can safely import the navigation module
    Import-Module "$modules\Navigation.psm1" -Force

    # Make sure we're not on the screen that appears when active from another device.
    Try {
        Confirm-LoggedInGame -TemplateName 'DifferentDevice.png'
        Write-Host "Status: Completed`n" -ForegroundColor Green
    }
    Catch {
        Write-Host "ERROR: Could not confirm logged into the game." -ForegroundColor Red
        exit
    }

    # Find the home screen where we can start the navigation from. Exit if not found.
    Find-Home

    # Determine number of loops needed to capture screenshots of all players in roster.
    $loopCount = [math]::Ceiling($rosterData.Player.Count / $PPS)

    # Begin VS screenshot capture routine
    if ($Mode -contains 'VS') {
        $Type = @()
        [string]$Type = 'VS'

        # Remove existing screenshots
        Get-ChildItem -Path "$screens\$Type" -File -Recurse | Remove-Item

        # Begin navigation
        Select-Button -TemplateName 'VS.png'
        Select-Button -TemplateName 'VsPointsRanking.png'
        Select-Button -TemplateName 'YourAllianceCheckbox.png'
                
        # if VS_DailyMin in config.json is set to anything but 0, capture daily scores.
        if ($Config.Alliance.VS_DailyMin -ne '0') {
            # Day 1 VS scores
            if (($day -contains '1') -or ($day -eq 'All')) {
                Write-Host "NAVIGATION: Get $Type Day 1 screenshots" -ForegroundColor Cyan
                Select-Button -TemplateName 'VS-Day1.png'
                Try {
                    Get-Screenshots -SavePath "$screens\$Type" `
                    -SaveName "${Type}_Day1" `
                    -LoopCount $loopCount `
                    -Type $Type
                    Write-Host "Status: Completed`n" -ForegroundColor Green
                } Catch {
                    Write-Host "ERROR: Failed to run Get-Screenshot function" -ForegroundColor Red
                }
            }
        
            # Day 2 VS scores
            if (($day -contains '2') -or ($day -eq 'All')) {
                Write-Host "NAVIGATION: Get $Type Day 2 screenshots" -ForegroundColor Cyan
                Select-Button -TemplateName 'VS-Day2.png'
                Try {
                    Get-Screenshots -SavePath "$screens\$Type" `
                    -SaveName "${Type}_Day2" `
                    -LoopCount $loopCount `
                    -Type $Type
                    Write-Host "Status: Completed`n" -ForegroundColor Green
                } Catch {
                    Write-Host "ERROR: Failed to run Get-Screenshot function" -ForegroundColor Red
                }
            }

            # Day 3 VS scores
            if (($day -contains '3') -or ($day -eq 'All')) {
                Write-Host "NAVIGATION: Get $Type Day 3 screenshots" -ForegroundColor Cyan
                Select-Button -TemplateName 'VS-Day3.png'
                Try {
                    Get-Screenshots -SavePath "$screens\$Type" `
                    -SaveName "${Type}_Day3" `
                    -LoopCount $loopCount `
                    -Type $Type
                    Write-Host "Status: Completed`n" -ForegroundColor Green
                } Catch {
                    Write-Host "ERROR: Failed to run Get-Screenshot function" -ForegroundColor Red
                }
            }

            # Day 4 VS scores
            if (($day -contains '4') -or ($day -eq 'All')) {
                Write-Host "NAVIGATION: Get $Type Day 4 screenshots" -ForegroundColor Cyan
                Select-Button -TemplateName 'VS-Day4.png'
                Try {
                    Get-Screenshots -SavePath "$screens\$Type" `
                    -SaveName "${Type}_Day4" `
                    -LoopCount $loopCount `
                    -Type $Type
                    Write-Host "Status: Completed`n" -ForegroundColor Green
                } Catch {
                    Write-Host "ERROR: Failed to run Get-Screenshot function" -ForegroundColor Red
                }
            }

            # Day 5 VS scores
            if (($day -contains '5') -or ($day -eq 'All')) {
                Write-Host "NAVIGATION: Get $Type Day 5 screenshots" -ForegroundColor Cyan
                Select-Button -TemplateName 'VS-Day5.png'
                Try {
                    Get-Screenshots -SavePath "$screens\$Type" `
                    -SaveName "${Type}_Day5" `
                    -LoopCount $loopCount `
                    -Type $Type
                    Write-Host "Status: Completed`n" -ForegroundColor Green
                } Catch {
                    Write-Host "ERROR: Failed to run Get-Screenshot function" -ForegroundColor Red
                }
            }

            # Day 6 VS scores
            if (($day -contains '6') -or ($day -eq 'All')) {
                Write-Host "NAVIGATION: Get $Type Day 6 screenshots" -ForegroundColor Cyan
                Select-Button -TemplateName 'VS-Day6.png'
                Try {
                    Get-Screenshots -SavePath "$screens\$Type" `
                    -SaveName "${Type}_Day6" `
                    -LoopCount $loopCount `
                    -Type $Type
                    Write-Host "Status: Completed`n" -ForegroundColor Green
                } Catch {
                    Write-Host "ERROR: Failed to run Get-Screenshot function" -ForegroundColor Red
                }
            }
        }# end VS Day screenshot capture loop

        # if VS_WeeklyMin in config.json is set to anything but 0, capture weekly scores.
        if ($Config.Alliance.VS_WeeklyMin -ne '0') {
            Select-Button -TemplateName 'VsWeekly.png' -BackupTemplate 'VsWeekly-EndOfVs.png'
            Get-Screenshots -SavePath "$screens\$Type" -SaveName "${Type}_Weekly" -LoopCount $loopCount -Type $Type -ScrollSetting 'Weekly'
        }
    }

    # Begin TD (Tech Donations) screen capture loop
    if ($Mode -contains 'TD') {
        $Type = @()
        [string]$Type = 'TD'
        
        Select-Button -TemplateName 'Alliance.png'
        Select-Button -TemplateName 'StrengthRanking.png'
        Select-Button -TemplateName 'StrRnkDonate.png'
        Select-Button -TemplateName 'StrRnkDonateWeekly.png'

        # Capture screenshots
        Get-Screenshots -SavePath "$screens\$Type" -SaveName $Type -LoopCount $loopCount -Type $Type
        }# end TD (Tech Donations) screenshot capture loop

    # Begin KS (Kill Score) screen capture loop
    if ($Mode -contains 'KS') {
        $Type = @()
        [string]$Type = 'KS'
        
        Select-Button -TemplateName 'Alliance.png'
        Select-Button -TemplateName 'StrengthRanking.png'
        Select-Button -TemplateName 'StrRnkKills.png'

        # Capture screenshots
        Get-Screenshots -SavePath "$screens\$Type" -SaveName $Type -LoopCount $loopCount -Type $Type
        }# end KS (Kill Score) screenshot capture loop
    

    }# end navigation loop


#####################################################################################
## ------ 4: Prepare Screenshots for OCR (preprocessing) ------
#####################################################################################
if ($Mode -contains 'VS') {
    # Set Type to match Mode for this loop
    $Type = @()
    [string]$Type = 'VS'
    
    # Generate VS jobs
    Write-Host "JOB COLLECTION: Collecting jobs to process..." -ForegroundColor Cyan
    $Jobs = Get-Jobs -Templates $Templates -ScaledTemplates $ScaledTemplates -Screens $Screens -Preproc $Preproc -Day $Day -PPS $PPS -Type $Type
    
    # Assign
    $CropJobs = $Jobs.CropJobs
    $PreprocJobs = $Jobs.PreprocJobs
    $ScreenshotLists = $Jobs.ScreenshotLists
    
    # Execute the jobs we collected multithreaded
    Invoke-MTCropping -Jobs $CropJobs
    Invoke-MTPreprocessing -Jobs $PreprocJobs
}

if ($Mode -contains 'TD') {
    # Set Type to match Mode for this loop
    $Type = @()
    [string]$Type = 'TD'
    

    Write-Host "JOB COLLECTION: Collecting jobs to process..." -ForegroundColor Cyan
    
    # Generate VS jobs
    $Jobs = Get-Jobs -Templates $Templates -ScaledTemplates $ScaledTemplates -Screens $Screens -Preproc $Preproc -Day $Day -PPS $PPS -Type $Type
    
    # Assign
    $CropJobs = $Jobs.CropJobs
    $PreprocJobs = $Jobs.PreprocJobs
    $ScreenshotLists = $Jobs.ScreenshotLists
    
    # Execute the jobs we collected multithreaded
    Invoke-MTCropping -Jobs $CropJobs
    Invoke-MTPreprocessing -Jobs $PreprocJobs
}

if ($Mode -contains 'KS') {
    # Set Type to match Mode for this loop
    $Type = @()
    [string]$Type = 'KS'
    

    Write-Host "JOB COLLECTION: Collecting jobs to process..." -ForegroundColor Cyan
    
    # Identify jobs
    $Jobs = Get-Jobs -Templates $Templates -ScaledTemplates $ScaledTemplates -Screens $Screens -Preproc $Preproc -Day $Day -PPS $PPS -Type $Type
    
    # Assign jobs
    $CropJobs = $Jobs.CropJobs
    $PreprocJobs = $Jobs.PreprocJobs
    $ScreenshotLists = $Jobs.ScreenshotLists

    # Execute the jobs we collected multithreaded
    Invoke-MTCropping -Jobs $CropJobs
    Invoke-MTPreprocessing -Jobs $PreprocJobs
}

#####################################################################################
## ------ 5: Process the OCR results and export to CSV
#####################################################################################
# GetEnumerator allows us to list key value pairs from a hash table
foreach ($kvp in $ScreenshotLists.GetEnumerator() | Sort-Object Key) {
    
    # Use the key name from the KVP to determine $NamePrefix
    $NamePrefix = $kvp.Key

    # Define our report type
    if ($Mode -eq 'VS') { $Type = 'VS' }
    if ($Mode -eq 'TD') { $Type = 'TD' }
    if ($Mode -eq 'KS') { $Type = 'KS' }

    # Get list of preprocessed images for the OCR
    #Write-Host "Getting preprocessed images for $NamePrefix now."
    $ImageFiles = Get-ChildItem -Path $Preproc -Filter "$NamePrefix*.png" | Sort-Object Name
    if ($ImageFiles.Count -eq 0) {
        Write-Host "No image files found for $NamePrefix. Skipping."
        continue
    }

    # Get initial report (process through OCR)
    Write-Host "REPORT GENERATION: Get $NamePrefix report" -ForegroundColor Cyan
    $Report = Get-Report -Images $ImageFiles -Type $Type

    # Clean up names that were interpreted incorrectly by the OCR
    Write-Host "REPORT GENERATION: Correct player names" -ForegroundColor Cyan
    $Report = Update-PlayerNames -RosterData $RosterData -Report $Report -Origin $NamePrefix

    # Set the minimum score based on the $Type (determined by $Mode)
    $MinScore = (Get-Variable -Name "$($Type)MinScore" -ErrorAction SilentlyContinue).Value

    # Calculate if player passed minimum requirements
    Write-Host "REPORT GENERATION: Calculate pass/fail status" -ForegroundColor Cyan
    $Report = Update-PassFail -Report $Report -MinScore $MinScore -RosterData $RosterData

    # Export the report to a CSV file
    Write-Host "REPORT GENERATION: Export $NamePrefix report to csv" -ForegroundColor Cyan
    Write-Host "Your report will be saved to the following path:"
    Write-Host "$reports"
    Export-Report -Path $reports -Name $NamePrefix -Report $Report

    # Output the result to the console
    #$Report | Out-Host

#####################################################################################
## ------ 6: OPTIONAL: Export CSV to Google Sheets
#####################################################################################
    # Define the starting cell of CSV data that will be exported to sheet
    if ($Config.GoogleSheets.Enabled -eq '1') {
        Write-Host "REPORT GENERATION: Export $NamePrefix report to Google Sheet" -ForegroundColor Cyan
        
        # Import the Google Sheets module
        Import-Module "$modules\GoogleSheets.psm1" -Force
        
        # Define the starting cells for each report type
        if ($NamePrefix -eq 'VS_Day1') { $StartCell = $Config.GoogleSheets.VS_Day1Cell }
        if ($NamePrefix -eq 'VS_Day2') { $StartCell = $Config.GoogleSheets.VS_Day2Cell }
        if ($NamePrefix -eq 'VS_Day3') { $StartCell = $Config.GoogleSheets.VS_Day3Cell }
        if ($NamePrefix -eq 'VS_Day4') { $StartCell = $Config.GoogleSheets.VS_Day4Cell }
        if ($NamePrefix -eq 'VS_Day5') { $StartCell = $Config.GoogleSheets.VS_Day5Cell }
        if ($NamePrefix -eq 'VS_Day6') { $StartCell = $Config.GoogleSheets.VS_Day6Cell }
        if ($NamePrefix -eq 'VS') { $StartCell = $Config.GoogleSheets.VS_WeeklyCell }
        if ($NamePrefix -eq 'TD') { $StartCell = $Config.GoogleSheets.TD_Cell }
        if ($NamePrefix -eq 'KS') { $StartCell = $Config.GoogleSheets.KS_Cell }

        # Get te most recently created csv file to export
        $CsvPath = Get-RecentCsv -Reports $reports -NamePrefix $NamePrefix

        # Build path the certificate for getting access token.
        $CertPath = Join-Path -Path $ConfigDir -ChildPath $Config.GoogleSheets.CertFile

        # Build arguments for clearing any existing data
        $params = @{
            SpreadsheetID = $Config.GoogleSheets.SpreadsheetID
            SheetName     = $Config.GoogleSheets.SheetName
            StartCell     = $StartCell
            EndCell       = (Get-EndCellFromStartCell -StartCell $StartCell -ColumnsToRight 3 -RowsToInclude 101)
            CertPath      = $CertPath
            }

        # Clear cells defined in arguments
        Clear-GoogleSheetRange @params

        # Build arguments for export function
        $params = @{
            SpreadsheetID = $Config.GoogleSheets.SpreadsheetID
            SheetName     = $Config.GoogleSheets.SheetName
            StartCell     = $StartCell
            CsvPath       = $CsvPath
            CertPath      = $CertPath
            }
        
        # Export the CSV with the defined arguments
        Export-CsvToGoogleSheet @params
    }
}
