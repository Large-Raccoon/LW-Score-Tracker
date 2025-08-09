#####################################################################################
## ------ Configuration ------
#####################################################################################

# --- Module Directory ---
$script:modules = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'modules'

# --- Import Module(s) ---
Import-Module "$script:modules\WindowsTools.psm1" -Force

# Get configuration settings from config.json
$script:ConfigPath = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'config\config.json'
$script:Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json

# Define adb executable path
$script:adb = Get-ProductExecutable -ProductName 'ADB' -FileName 'adb.exe'

#####################################################################################
## ------ Functions: ADB ------
#####################################################################################

Function Start-AdbServer{
    $adbProcess = Get-Process -Name 'adb' -ErrorAction SilentlyContinue

    if ($adbProcess) {
        return
    } else {
        & $adb start-server
    }
}

function Connect-AdbDevice {

    # Get list of currently connected devices
    $adbDevicesOutput = & $adb devices
    $deviceList = @()

    # Parse output, skipping the first line (header)
    foreach ($line in $adbDevicesOutput[1..($adbDevicesOutput.Count - 1)]) {
        if ($line -match '^(\S+)\s+device$') {
            $deviceList += $matches[1]
        }
    }

    # If no devices, try to connect
    if ($Config.ADB.Device -notcontains $deviceList) {
        Write-Host "No ADB device found. Attempting to connect to $($Config.ADB.Device)..."

        $connectResult = & $adb connect $($Config.ADB.Device)

        if ($connectResult -match 'connected to') {
            Write-Host "Connected to $($Config.ADB.Device)"
        }
        else {
            Write-Host "ERROR: Failed to connect to $($Config.ADB.Device)" -ForegroundColor Red
            Write-Host "Check if your ADB device\emulator is running then try again." -ForegroundColor Red
            exit
        }
    }
    return
}

function Get-AdbForegroundPackage {
    $cmd = 'dumpsys window | grep mCurrentFocus'
    $params = @("-s", $Config.ADB.Device, "shell", $cmd)
    $CurrentApp = & $adb $params

    return $CurrentApp
}

Function Start-AdbPackage {
    param (
        [Parameter(Mandatory = $true)]
        [string]$PackageName
    )

    $maxAttempts = 6
    $sleepInterval = 30

    for ($i = 1; $i -le $maxAttempts; $i++) {
        Write-Host "Attempt ${i} of ${maxAttempts}: Checking if $PackageName is running..."

        $CurrentApp = Get-AdbForegroundPackage

        if ($CurrentApp -match $PackageName) {
            Write-Host "Detected $PackageName"
            #Write-Host "Status: Completed`n" -ForegroundColor Green
            break
        }

        Write-Host "$PackageName not detected. Launching app..."
        $cmd = "monkey -p $PackageName -c android.intent.category.LAUNCHER 1"
        $params = @("-s", $Config.ADB.Device, "shell", $cmd)
        & $adb @params > $null 2>&1

        $totalWait = $sleepInterval * $i
        Write-Host "Waiting $totalWait seconds before next check..."
        Start-Sleep -Seconds $totalWait
    }

    if ($CurrentApp -notmatch $PackageName) {
        Write-Host "ERROR: $PackageName failed to launch after $maxAttempts attempts." -ForegroundColor Red
        exit
    }
}

function Stop-AdbPackage {
    param (
        [Parameter(Mandatory=$true)]
        [string]$PackageName
    )

    # Run the force-stop command
    $cmd = "am force-stop $PackageName"
    $params = @("-s", $Config.ADB.Device, "shell", $cmd)
    & $adb @params
    Start-Sleep -Seconds 5
}

Function Invoke-AdbCurrentFocus {
    $cmd    = "dumpsys window windows | grep mCurrentFocus"
    $params   = @("-s", $Config.ADB.Device, "shell", $cmd)
    $output = & $adb @params

    if ($output -notmatch 'com\.fun\.lastwar\.gp') {
        Start-AdbPackage -PackageName $Config.ADB.Package
    }
}

Function Invoke-AdbTap {
param (
    [Parameter(Mandatory = $true)]
    [string]$X,
    [string]$Y
)

    $cmd = "input tap $X $Y"
    $params = @('-s', $Config.ADB.Device, "shell", $cmd)
    & $adb $params
    Start-Sleep -Seconds 1
}

Function Invoke-AdbSwipe {
param (
    [Parameter(Mandatory=$true)]
    [string]$StartX,
    [string]$StartY,
    [string]$EndX,
    [string]$EndY,
    [string]$Duration
)

    $cmd = "input swipe $StartX $StartY $EndX $EndY $Duration && input tap $StartX $StartY"
    $params = @("-s", $Config.ADB.Device, "shell", $cmd)
    & $adb @params
}

Function Invoke-AdbBack {
    $cmd = "input keyevent 4"
    $params = @("-s", $Config.ADB.Device, "shell", $cmd)
    & $adb @params
}

Function Get-AdbScreenshot {
Param (
    [Parameter(Mandatory=$true)]
    [string]$SavePath,
    [string]$SaveName
    )

    $FullName = Join-Path $SavePath $SaveName
    & $adb -s $Config.ADB.Device shell screencap -p /sdcard/screen.png
    & $adb -s $Config.ADB.Device pull -q /sdcard/screen.png $FullName
    & $adb -s $Config.ADB.Device shell rm /sdcard/screen.png
}

Export-ModuleMember -Function *