#####################################################################################
## ------ Configuration ------
#####################################################################################

# Get configuration settings from config.json
$ConfigPath = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'config\config.json'
$script:Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json

# --- .NET ASSEMBLIES ---
# System.Drawing: For getting image dimensions
Add-Type -AssemblyName System.Drawing

if ($Config.PC.Enabled -eq '1') {
    Add-Type @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;

public class GameCapture {
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
}
"@
}
#####################################################################################
## ------ Functions: Windows Tools ------
#####################################################################################

# Locate the product code associated with defined $ARPName and retrieve the installation folder.
Function Get-ProductExecutable {
Param (
    [Parameter(Mandatory=$true)] [string]$ProductName,
    [Parameter(Mandatory=$false)] [string]$FileName,
    [Parameter(Mandatory=$false)] [switch]$PathOnly
    )

    # Define registry keys to search through.
    $ProductsX64 = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall'
    $ProductsX86 = 'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    $Software = 'HKLM:\Software'
    
    $RegistryKeys =@(
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\Software'
        )

    $InstallDirectories =@(
        $env:ProgramFiles,
        ${env:ProgramFiles(x86)}
        "$env:SystemDrive\"
        "$env:LOCALAPPDATA\"
        )

    # Check through the registry
    ForEach ($Key in $RegistryKeys) {
        $InstallLocation = Get-ItemProperty $Key\* |
            Where-Object {($_.DisplayName -like $ProductName + '*')} |
            Select-Object -First 1 |
            Select-Object -ExpandProperty InstallLocation -EA SilentlyContinue

        if (![string]::IsNullOrEmpty($InstallLocation)) {
            if (Test-Path -Path $InstallLocation -EA SilentlyContinue) {
                if ($PathOnly) {
                    return $InstallLocation
                }

                $FullPath = Join-Path -Path $InstallLocation -ChildPath $FileName -EA SilentlyContinue
                if (Test-Path -Path $FullPath -EA SilentlyContinue) {
                    return $FullPath
                }
            }
        }
    }

    # Check if we can find the executable in standard directories on the file system
    ForEach ($Directory in $InstallDirectories) {
        $InstallLocation = (Get-ChildItem $Directory | Where-Object {($_.Name -like '*' + $ProductName + '*')} | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
        if (![string]::IsNullOrEmpty($InstallLocation)) {
            if ($PathOnly) {
                return $InstallLocation
            }

            $FullPath = (Get-ChildItem -Path "$InstallLocation" -Recurse -File -Filter $FileName | Select-Object -First 1).FullName
            if (![string]::IsNullOrEmpty($FullPath)) {
                return $FullPath
            }
        }
    }

    # Absolute fallback will rely on executable being in PATH
    if (!($PathOnly)) {
        $PathAlias = ($FileName -replace "\.[^.]+$","")
        return $PathAlias
    }
}

#####################################################################################
## ------ Functions: BlueStacks for Windows ------
#####################################################################################

Function Get-BlueStacksInstance {
    param (
        [string]$ConfigPath = "$env:ProgramData\BlueStacks_nxt\bluestacks.conf"
    )

    if (-not (Test-Path $ConfigPath)) {
        Write-Warning "Config file not found at: $ConfigPath"
        return $null
    }

    $content = Get-Content $ConfigPath -Raw

    # Regex to get value between "bst.instance." and ".status.adb_port="5555""
    if ($content -match 'bst\.instance\.([^.]+)\.status\.adb_port="5555"') {
        $instance = $matches[1]
        return $instance
    } else {
        return $false
    }
}

function Get-BluestacksAdbSetting {
    param (
        [string]$ConfigPath = "$env:ProgramData\BlueStacks_nxt\bluestacks.conf"
    )

    if (-not (Test-Path $ConfigPath)) {
        Write-Warning "Config file not found at: $ConfigPath"
        return $null
    }

    $content = Get-Content $ConfigPath -Raw

    # Regex to match bst.enable_adb_access="0" or "1"
    if ($content -match 'bst\.enable_adb_access\s*=\s*"(?<value>[01])"') {
        $value = $matches['value']
        switch ($value) {
            "0" { return $false }
            "1" { return $true }
        }
    } else {
        return $false
    }
}

#####################################################################################
## ------ Functions: Last War PC Client ------
#####################################################################################

function Get-PcScreenshot {
    param(
        [string]$ProcessName,
        [string]$SavePath,
        [string]$SaveName
    )
    
    $FullName = Join-Path $SavePath $SaveName
    $proc = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 }

    if (-not $proc) {
        Write-Error "Could not find window for process '$ProcessName'"
        return
    }

    $hWnd = $proc.MainWindowHandle

    $rect = New-Object GameCapture+RECT
    [GameCapture]::GetWindowRect($hWnd, [ref]$rect) | Out-Null

    $width = $rect.Right - $rect.Left
    $height = $rect.Bottom - $rect.Top

    $bitmap = New-Object Drawing.Bitmap $width, $height
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    $graphics.CopyFromScreen($rect.Left, $rect.Top, 0, 0, $bitmap.Size)

    $bitmap.Save($FullName)
    $graphics.Dispose()
    $bitmap.Dispose()
}

function Set-WindowSize {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $ProcessName,
        [Parameter(Mandatory)][int]    $Width,
        [Parameter(Mandatory)][int]    $Height
    )

    # Load P/Invoke helper
    if (-not ('Win32WindowHelperUserRect' -as [type])) {
        try {
            Add-Type @"
using System;
using System.Runtime.InteropServices;

public class Win32WindowHelperUserRect {
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    public static extern bool SetWindowPos(
        IntPtr hWnd,
        IntPtr hWndInsertAfter,
        int X,
        int Y,
        int cx,
        int cy,
        uint uFlags
    );

    public static readonly IntPtr HWND_TOP     = new IntPtr(0);
    public const          uint    SWP_NOZORDER = 0x0004;
}
"@ -ErrorAction Stop
        } catch {
            Write-Warning "Add-Type failed: $_"
        }
    }

    # Ensure System.Windows.Forms is loaded
    if (-not ('System.Windows.Forms.Screen' -as [type])) {
        try {
            Add-Type -AssemblyName 'System.Windows.Forms' -ErrorAction Stop
        } catch {
            Write-Warning "Could not load System.Windows.Forms: $_"
        }
    }

    # Find the window
    $proc = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowHandle -ne 0 } |
            Select-Object -First 1

    if (-not $proc) {
        Write-Error "No visible window found for process '$ProcessName'."
        return
    }

    # Read current RECT and compute clamp
    $rect = New-Object Win32WindowHelperUserRect+RECT
    [Win32WindowHelperUserRect]::GetWindowRect($proc.MainWindowHandle, [ref]$rect) | Out-Null

    $screen   = [System.Windows.Forms.Screen]::FromHandle($proc.MainWindowHandle)
    if (-not $screen) { $screen = [System.Windows.Forms.Screen]::PrimaryScreen }

    $wa   = $screen.WorkingArea
    $minX = $wa.Left
    $minY = $wa.Top
    $maxX = $wa.Right  - $Width
    $maxY = $wa.Bottom - $Height

    $newX = [Math]::Min([Math]::Max($rect.Left, $minX), $maxX)
    $newY = [Math]::Min([Math]::Max($rect.Top,  $minY), $maxY)

    # Resize & reposition
    [Win32WindowHelperUserRect]::SetWindowPos(
        $proc.MainWindowHandle,
        [Win32WindowHelperUserRect]::HWND_TOP,
        $newX,
        $newY,
        $Width,
        $Height,
        [Win32WindowHelperUserRect]::SWP_NOZORDER
    ) | Out-Null
}

Export-ModuleMember -Function *