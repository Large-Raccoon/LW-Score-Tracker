#####################################################################################
## ------ Configuration ------
#####################################################################################

# --- .NET ASSEMBLIES ---
# System.Drawing: For getting image dimensions
Add-Type -AssemblyName System.Drawing

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

            $FullPath = (Get-ChildItem -Path "$InstallLocation" -Recurse -File -Filter $FileName).FullName
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
## ------ Functions: Tesseract for Windows ------
#####################################################################################

Export-ModuleMember -Function *