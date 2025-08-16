#####################################################################################
## ------ Configuration ------
#####################################################################################

# --- Module Directory ---
$script:modules = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'modules'

# --- Logs Directory ---
$script:Logs = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'Logs'

# --- Import Module(s) ---
Import-Module "$script:modules\ImageProcessing.psm1" -Force

# --- Configuration ---
    
# Get configuration settings from config.json
$script:ConfigPath = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'config\config.json'
$script:Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json

# Logging function
Function Write-Log {
    param ($Message)
    Out-File -FilePath "$Logs\NameCorrection.log" -InputObject "$Message" -Append -Force
    }

#####################################################################################
## ------ Functions: Text Processing ------
#####################################################################################

Function Get-LevenshteinDistance {
    param (
        [string]$s,
        [string]$t
        )

    $n = $s.Length
    $m = $t.Length

    # Initialize a jagged array
    $d = @()
    for ($i = 0; $i -le $n; $i++) {
        $d += ,(0..$m)
        }

    # Fill base cases
    for ($i = 0; $i -le $n; $i++) { $d[$i][0] = $i }
    for ($j = 0; $j -le $m; $j++) { $d[0][$j] = $j }

    for ($i = 1; $i -le $n; $i++) {
        for ($j = 1; $j -le $m; $j++) {
            $cost = if ($s[$i - 1] -eq $t[$j - 1]) { 0 } else { 1 }

            $d[$i][$j] = [Math]::Min(
                [Math]::Min($d[$i - 1][$j] + 1, $d[$i][$j - 1] + 1),
                $d[$i - 1][$j - 1] + $cost
            )
        }
    }
    return $d[$n][$m]
}

Function Get-Report {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Images,
        [string]$Type
    )

    $report = @()
    [int]$RankNumber = 0
    foreach ($img in $Images) {
        $ocrResult = Use-ProcessImage -imgPath $img.FullName -Type $Type

        # Debug only
        #Write-Host ">> OCR: Rank=$($ocrResult.Rank), Player=$($ocrResult.Player), Score=$($ocrResult.Score)"
        
        # Progress bar
        $RankNumber++
        $percentComplete = [math]::Round(($RankNumber / $($Images.Count)) * 100, 0)
        Write-Progress -Activity "Job: Process rank screenshots through OCR" `
                        -Status "Rank: $RankNumber of $($Images.Count)" `
                        -PercentComplete $percentComplete
        

        $reportObj = [PSCustomObject]@{
            'Rank'   = $ocrResult.Rank
            'Player' = $ocrResult.Player
            'Score'  = $ocrResult.Score
        }
        $report += $reportObj
    }
    Write-Progress -Activity "Job: Process rank screenshots through OCR" -Completed
    Write-Host "Status: Completed`n" -ForegroundColor Green
    return $report
}

function Update-PlayerNames {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Report,

        [Parameter(Mandatory = $true)]
        [array]$RosterData,

        [Parameter(Mandatory = $true)]
        [string]$Origin
    )

#Write-Log "`nCorrection log for: $Origin`n"
Write-Log @("Correction Log - $Origin

Only player names that were incorrectly read by the OCR will be listed here.

Legend:
✅ = An alias match was performed. No issue here!
⚠️ = A fuzzy match was performed. Difference shows how many characters off the name was.
❌ = Failed to guess the correct name.

Instructions:
For ⚠️, you may leave them as is if no issues. Otherwise, add invalid name as an alias to roster.csv.
For ❌, confirm player is in your roster.csv. Add invalid name as an alias if necessary to roster.csv.
")

    $updatedReport = @()
    $usedNames     = @()  # Track player names already used

    foreach ($entry in $Report) {
        $originalName = $entry.Player
        $updatedName  = $originalName

        # SKIP if the OCR name is blank
        if ([string]::IsNullOrWhiteSpace($originalName)) {
            # leave $updatedName as-is (blank) and go straight to updating $updatedReport
        }
        elseif ($RosterData.Player -notcontains $originalName) {
            # If no roster match, check for a matching alias to correct the OCR name.
            $aliasMatch = $RosterData |
            Where-Object {
                (
                (![string]::IsNullOrWhiteSpace($_.Alias1) -and $_.Alias1 -eq $originalName) -or
                (![string]::IsNullOrWhiteSpace($_.Alias2) -and $_.Alias2 -eq $originalName) -or
                (![string]::IsNullOrWhiteSpace($_.Alias3) -and $_.Alias3 -eq $originalName) -or
                (![string]::IsNullOrWhiteSpace($_.Alias4) -and $_.Alias4 -eq $originalName) -or
                (![string]::IsNullOrWhiteSpace($_.Alias5) -and $_.Alias5 -eq $originalName) -or
                (![string]::IsNullOrWhiteSpace($_.Alias6) -and $_.Alias6 -eq $originalName)
                )
            } | Select-Object -First 1

            if ($aliasMatch) {
                $updatedName = $aliasMatch.Player
                #Write-Host "Rank $($Entry.Rank) - `"$originalName`" corrected to `"$updatedName`"" -ForegroundColor Cyan
                Write-Log  "✅ Rank $($Entry.Rank)"
                Write-Log "Invalid Name: `"$originalName`""
                Write-Log "Corrected Name: `"$updatedName`""
                Write-Log "Character Difference: N/A`n"
            }
            else {
                # If no alias match, perform a fuzzy match using Levenshtein formula.
                $bestMatch = ''
                $bestScore = [int]::MaxValue

                foreach ($trueName in $RosterData.Player) {
                    # Skip this loop if player name has already been used in a previous correction.
                    if ($usedNames -contains $trueName) { continue } 

                    $dist = Get-LevenshteinDistance -s $originalName -t $trueName
                    if ($dist -lt $bestScore) {
                        $bestScore = $dist
                        $bestMatch = $trueName
                    }
                }

                if ($bestScore -le $Config.NameCorrection.MaxDistance) {
                    $updatedName = $bestMatch
                    #Write-Host "Rank $($Entry.Rank) - `"$originalName`" corrected to `"$updatedName`" (Distance: $bestScore)" -ForegroundColor Yellow
                    Write-Log  "⚠️ Rank $($Entry.Rank)"
                    Write-Log "Invalid Name: `"$originalName`""
                    Write-Log "Corrected Name: `"$updatedName`""
                    Write-Log "Character Difference: $bestScore`n"
                }
                else {
                    #Write-Host "MATCH FAILURE: Rank $($Entry.Rank) - `"$originalName`" (closest match: `"$bestMatch`" - distance: $bestScore)" -ForegroundColor Red
                    Write-Log "❌: Rank $($Entry.Rank)"
                    Write-Log "Invalid Name: `"$originalName`""
                    Write-Log "Closest Name: `"$bestMatch`""
                    Write-Log "Character Difference: $bestScore`n"
                }
            }
        }

        # If name already used and somehow managed to get through, remove it.
        if ($updatedName -and ($usedNames -contains $updatedName) -and ($originalName -ne $updatedName)) {
            Write-Warning "Duplicate resolved name '$updatedName' detected. Clearing for player '$originalName'"
            Write-Log     "DUPLICATE DETECTED: $updatedName."
            $updatedName = ""
        }
        elseif ($updatedName) {
            $usedNames += $updatedName
        }

        # Final pass: remove any names not in roster after corrections
        for ($i = 0; $i -lt $updatedReport.Count; $i++) {
            $name = $updatedReport[$i].Player
            if (![string]::IsNullOrWhiteSpace($name) -and ($RosterData.Player -notcontains $name)) {
                $updatedReport[$i].Player = ''
            }
        }

        # Update the report
        $updatedReport += [PSCustomObject]@{
            Rank   = $entry.Rank
            Player = $updatedName
            Score  = $entry.Score
        }
    }
    Write-Host "Status: Completed`n" -ForegroundColor Green
    return $updatedReport
}

function Update-PassFail {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Report,

        [Parameter(Mandatory = $true)]
        [int]$MinScore,  # Format: "7,200,000"

        [Parameter(Mandatory = $true)]
        [array]$RosterData  # Expecting objects with a 'Player' property
    )

    # Dictionary for quick lookup of existing players
    $existingPlayers = @{}
    foreach ($entry in $Report) {
        $existingPlayers[$entry.Player] = $true
    }

    # Start with updated report entries (evaluating Pass)
    $updatedReport = foreach ($entry in $Report) {
        $scoreRaw = $entry.Score
        $passStatus = "?"

        if (-not [string]::IsNullOrWhiteSpace($scoreRaw)) {
            try {
                $scoreInt = [int]($scoreRaw -replace '[^\d]')
                $passStatus = if ($scoreInt -ge $MinScore) { "Y" } else { "N" }
            }
            catch {
                $passStatus = "?"
            }
        }

        [PSCustomObject]@{
            Rank   = $entry.Rank
            Player = $entry.Player
            Score  = $entry.Score
            Pass   = $passStatus
        }
    }

    # Add missing players from roster
    foreach ($player in $RosterData.Player) {
        if (-not $existingPlayers.ContainsKey($player)) {
            $updatedReport += [PSCustomObject]@{
                Rank   = "N/A"
                Player = $player
                Score  = "0"
                Pass   = "N"
            }
        }
    }
    Write-Host "Status: Completed`n" -ForegroundColor Green
    return $updatedReport
}

Function Export-Report {
param (
    [Parameter(Mandatory=$true)]
    [string]$Path,
    [string]$Name,
    [array]$Report
)
   
    # Convert the PS Object into CSV format
    $csvContent = $Report | ConvertTo-Csv -NoTypeInformation
    
    # Create UTF-8 with BOM encoding object
    $utf8BomEncoding = New-Object System.Text.UTF8Encoding $true
    
    # Get date and append to ReportName
    $csvName = "$(Get-Date -Format "MMddyy")_$($Name)"

    # Full path to output CSV (adjust if needed)
    $csvPath = Join-Path $Path "$csvName.csv"

    # Use .Net System.IO file class to write to a csv file. Workaround for PowerShell 5.1 to write UTF-8 with BOM
    Try {
        [System.IO.File]::WriteAllLines("$csvPath", $csvContent, $utf8BomEncoding)
        } Catch {
            Write-Warning "Report $csvName needs to be closed before continuing."
            pause
            [System.IO.File]::WriteAllLines("$csvPath", $csvContent, $utf8BomEncoding)
        }
    Write-Host "Status: Completed`n" -ForegroundColor Green
}

Function Get-RecentCsv {
param (
    [string]$reports,       # Path to the folder containing CSV files
    [string]$NamePrefix     # Name pattern to match
)
    # Identify csv files with name like $Type
    $matchingFiles = Get-ChildItem -Path $reports -Filter "*_${NamePrefix}*.csv" -File

    # If no match, exit or handle accordingly
    if (-not $matchingFiles) {
        Write-Warning "No CSV files found matching '*_${NamePrefix}*.csv' in '$reports'"
        return $null
    }

    # Select the most recently created matching file
    $latestFile = ($matchingFiles | Sort-Object CreationTime -Descending | Select-Object -First 1).FullName

    # Return the result to main script
    return $latestFile
}

Export-ModuleMember -Function *