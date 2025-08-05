#####################################################################################
## ------ Configuration ------
#####################################################################################

# --- Module Directory ---
$script:modules = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'modules'

# --- Templates Directory ---
$script:OriginalTemplates = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'config\templates'
$script:ScaledTemplates = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'imgproc\templates_scaled'

# --- Screenshot Directory ---
$script:Nav = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'imgproc\screens\Nav'

# --- Logs Directory ---
$script:Logs = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'Logs'

# --- Screenshot Default Name ---
$script:GameState = 'GameState.png'

# --- Import Module(s) ---
Import-Module "$script:modules\ImageProcessing.psm1" -Force
Import-Module "$script:modules\AdbTools.psm1" -Force

# --- Configuration ---
# Get configuration settings from config.json
$script:ConfigPath = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'config\config.json'
$script:Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json

# --- Dynamic Configuration ---
# Take a screenshot to determine if scaling is needed later
Get-AdbScreenshot -SavePath $Nav -SaveName $GameState # Check if we need to add a delay during testing

# Load a sample image from the screenshots
$imgSample = (Get-ChildItem "$Nav" -Filter '*.png' | Sort-Object Name | Select-Object -First 1).FullName

# Check image dimensions of sample image
$img = [System.Drawing.Image]::FromFile($imgSample)
$sampleWidth = $img.Width
$sampleHeight = $img.Height
$img.Dispose()

# Generate list of UI image templates that will be referenced
$referenceImgs = (Get-ChildItem "$OriginalTemplates\ui" -Filter '*.png' | Sort-Object Name).FullName

# Resize the templates proportionately to the screenshots if necessary
If ($sampleWidth -notin '1200', '1080') {
    $referenceImgs | ForEach-Object {
        $imgName = Split-Path $_ -Leaf

        Write-Host "The result of imgName is:"
        $imgName | Out-Host

        Invoke-ResizeTemplate -InputImg $_ -OutputImg "$ScaledTemplates\ui\$imgName" -SourceImg $imgSample
        $script:TemplatePath = Join-Path $ScaledTemplates 'ui'
    }
} Else {
        $script:TemplatePath = Join-Path $OriginalTemplates 'ui'
    }

#####################################################################################
## ------ Functions: Android Navigation ------
#####################################################################################

Function Find-Home {

    Write-Host "NAVIGATION: Search for home screen" -ForegroundColor Cyan

    $ReferenceImage1 = Join-Path $script:TemplatePath 'Quit.png'
    $ReferenceImage2 = Join-Path $script:TemplatePath 'home.png'
    $SourceImage = Join-Path $script:Nav $script:GameState
    
    # Define looping
    [int]$maxLoops = '18'
    [int]$loopCount = '0'

    do {
        Invoke-AdbBack
        Start-Sleep -Milliseconds 1000

        # Take a screenshot to determine where we currently are
        Get-AdbScreenshot -SavePath $Nav -SaveName $GameState # Check if we need to add a delay during testing
        Start-Sleep -Seconds 1

        # Check for a template match
        Clear-Variable -Name match -Force -EA SilentlyContinue
        $match = Get-Coords -refimg $ReferenceImage1 -srcimg $SourceImage

        # If "Quit the game?" found, hit back once more. Else, hit it 3 more times.
        if ($null -ne $match) {
            Invoke-AdbBack
            Get-AdbScreenshot -SavePath $Nav -SaveName $GameState # Check if we need to add a delay during testing
            
            # Check for a template match
            Clear-Variable -Name match -Force -EA SilentlyContinue
            $match = Get-Coords -refimg $ReferenceImage2 -srcimg $SourceImage

            if ($null -ne $match) {
                Write-Host "Status: Completed`n" -ForegroundColor Green
                return $null
            }
        }
        #Increment the loop count
        $loopCount++
    }
    while ($loopCount -lt $maxLoops)
    Write-Host "Status: ERROR - Could not find the home screen.`n" -ForegroundColor Red
    exit
}

Function Confirm-LoggedInGame {
    param (
        [Parameter(Mandatory = $true)]    
        [string]$TemplateName
    )
    $CurrentApp = Get-AdbForegroundPackage

    if ($CurrentApp.Contains($Config.ADB.Package)) {
        # Take a screenshot to determine where we currently are
        Get-AdbScreenshot -SavePath $Nav -SaveName $GameState # Check if we need to add a delay during testing

        # Check for a template match
        Clear-Variable -Name match -Force -EA SilentlyContinue
        $match = Get-Coords -refimg "$TemplatePath\$TemplateName" -srcimg "$($Nav)\$($GameState)"
        if ($null -ne $match) {
            # code to close app
            Stop-AdbPackage -PackageName $($Config.ADB.Package)
            #Start-Sleep -Seconds 10
            Start-AdbPackage -PackageName $($Config.ADB.Package)
        }
    }
}

Function Select-Button {
    param (
        [Parameter(Mandatory = $true)]    
        [string]$TemplateName,
        [string]$BackupTemplate
    )

    # Take a screenshot to determine where we currently are
    Get-AdbScreenshot -SavePath $Nav -SaveName $GameState # Check if we need to add a delay during testing

    # Check for a template match
    Clear-Variable -Name match -Force -EA SilentlyContinue
    $match = Get-Coords -refimg "$TemplatePath\$TemplateName" -srcimg "$($Nav)\$($GameState)"
    if ($null -ne $match) {
        # Tap the matching coordinates
        Invoke-AdbTap -X $($match.x) -Y $($match.y)
    } Else {
        $match = Get-Coords -refimg "$TemplatePath\$BackupTemplate" -srcimg "$($Nav)\$($GameState)"
        if ($null -ne $match) {
            # Tap the matching coordinates
            Invoke-AdbTap -X $($match.x) -Y $($match.y)
        } Else {
            Write-Host "Failed to find template match for $TemplateName. I don't know what to tap!" -ForegroundColor Red
            exit
        }
    }

}

Export-ModuleMember -Function *