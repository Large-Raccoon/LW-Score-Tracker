#####################################################################################
## ------ Configuration ------
#####################################################################################

# --- Module Directory ---
$script:modules = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'modules'

# --- Import Module(s) ---
Import-Module "$script:modules\AdbTools.psm1" -Force
Import-Module "$script:modules\WindowsTools.psm1" -Force

# --- Configuration ---
# Get configuration settings from config.json
$ConfigPath = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'config\config.json'
$script:Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json


#####################################################################################
## ------ Functions: Helper ------
#####################################################################################

function Get-ColumnIndexFromLetter {
    [CmdletBinding()]
    Param(
        [string]$ColumnLetters
    )
    # Converts a column letter sequence (e.g. "AA") into a 1-based column index (e.g. 27).
    # Algorithm: treat letters as base-26 digits (A=1, B=2, ..., Z=26).
    $ColumnLetters = $ColumnLetters.ToUpper()        # Ensure input is uppercase for consistency
    $index = 0
    foreach ($char in $ColumnLetters.ToCharArray()) {
        $index = $index * 26 + ([int][char]$char - [int][char]'A' + 1)
    }
    return $index
}
function Get-ColumnLetterFromIndex {
    [CmdletBinding()]
    Param(
        [int]$ColumnIndex
    )
    if ($ColumnIndex -lt 1) { return "" }
    $letters = ""
    while ($ColumnIndex -gt 0) {
        $offset = ($ColumnIndex - 1) % 26
        $letters = [char]([byte][char]'A' + $offset) + $letters
        $ColumnIndex = [math]::Floor(($ColumnIndex - 1) / 26)
    }
    return $letters
}
function Get-EndCellFromStartCell {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $StartCell,
        [int] $ColumnsToRight = 3,
        [int] $RowsToInclude  = 100
    )

    # Parse StartCell into column letters + row number
    if ($StartCell -notmatch '^([A-Z]+)(\d+)$') {
        throw "StartCell '$StartCell' is not valid A1 notation."
    }
    $startColLetters = $Matches[1]
    $startRowNumber  = [int]$Matches[2]

    # Compute numeric indices
    $startColIndex = Get-ColumnIndexFromLetter -ColumnLetters $startColLetters
    $endColIndex   = $startColIndex + $ColumnsToRight

    # Convert back to letters and calculate end row
    $endColLetters = Get-ColumnLetterFromIndex -ColumnIndex $endColIndex
    $endRowNumber  = $startRowNumber + $RowsToInclude - 1

    # Return the A1 notation for the EndCell
    return "$endColLetters$endRowNumber"
}
 # Base64URL encoding helper
    function Set-EncodedBase64Url($bytes) {
        [Convert]::ToBase64String($bytes) `
            -replace '\+', '-' -replace '/', '_' -replace '=', ''
    }

#####################################################################################
## ------ Functions: Google Sheets ------
#####################################################################################

function Get-GoogleAccessToken {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateScript({ Test-Path $_ })]
        [string]$CertPath,

        [string]$Account = $Config.GoogleSheets.Account,

        [string]$P12Password = "notasecret"
    )

    # Constants
    $Scope = "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive https://www.googleapis.com/auth/drive.file"
    $ExpirationSeconds = 3600
    $TokenUri = "https://oauth2.googleapis.com/token"

    # Load key material
    if ($CertPath -like "*.json") {
        $json = Get-Content $CertPath -Raw | ConvertFrom-Json
        $ServiceAccount = $json.client_email
        $privateKeyPem = $json.private_key

        if ($PSVersionTable.PSVersion.Major -ge 7) {
            $rsa = [System.Security.Cryptography.RSA]::Create()
            $rsa.ImportFromPem($privateKeyPem)
        }
        else {
            Write-Host "ERROR: json certificate is only supported in PowerShell 7+." -ForegroundColor Red
            Write-Host "Use a .p12 certificate for Powershell 5.1." -ForegroundColor Red
            exit
        }
    }
    elseif ($CertPath -like '*.p12') {

    # Load the certificate
    $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
        $CertPath,                                  # path
        $P12Password,                               # password
        [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable `
        -bor [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::MachineKeySet
    )

    if ($PSVersionTable.PSVersion.Major -ge 7) {
        # In PS 7+ just grab the RSA key
        $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
    }
    else {
        # PS 5.1: wrap the cert’s RSA into an AES-enabled CSP so SHA-256 works
        # Extract the parameters from the default PrivateKey
        $oldRsa     = $cert.PrivateKey
        $rsaParams  = $oldRsa.ExportParameters($true)

        # Create a PROV_RSA_AES (provider type 24) CSP
        $cspParams              = New-Object System.Security.Cryptography.CspParameters
        $cspParams.ProviderType = 24                                        # PROV_RSA_AES
        $cspParams.ProviderName = 'Microsoft Enhanced RSA and AES Cryptographic Provider'
        $cspParams.Flags        = [System.Security.Cryptography.CspProviderFlags]::UseMachineKeyStore

        # New RSACryptoServiceProvider with that CSP, then import parameters
        $rsa = New-Object System.Security.Cryptography.RSACryptoServiceProvider($cspParams)
        $rsa.ImportParameters($rsaParams)
    }

    if (-not $rsa) {
        throw 'Unable to retrieve or initialize an RSA key for signing.'
    }

    $ServiceAccount = $Account
}
    else {
        throw "Unsupported certificate type: $CertPath"
    }

    # Prepare JWT
    $now   = [DateTime]::UtcNow
    $iat   = [Math]::Floor((New-TimeSpan -Start ([datetime]'1970-01-01') -End $now).TotalSeconds)
    $exp   = $iat + $ExpirationSeconds

    $header = @{ alg = "RS256"; typ = "JWT" } | ConvertTo-Json -Compress
    $claims = @{
        iss   = $ServiceAccount
        scope = $Scope
        aud   = $TokenUri
        iat   = $iat
        exp   = $exp
    } | ConvertTo-Json -Compress


    $headerB64 = Set-EncodedBase64Url ([Text.Encoding]::UTF8.GetBytes($header))
    $claimsB64 = Set-EncodedBase64Url ([Text.Encoding]::UTF8.GetBytes($claims))
    $toSign    = "$headerB64.$claimsB64"

    # Sign JWT
    if ($PSVersionTable.PSVersion.Major -ge 7) {
    $sigBytes = $rsa.SignData(
        [Text.Encoding]::UTF8.GetBytes($toSign),
        [System.Security.Cryptography.HashAlgorithmName]::SHA256,
        [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
        )
    }
    else {
        # .NET Framework RSACryptoServiceProvider
        # PS 5.1: pass the SHA-256 OID, not the literal string "SHA256"
        $sigBytes = $rsa.SignData(
        [Text.Encoding]::UTF8.GetBytes($toSign),
        [System.Security.Cryptography.CryptoConfig]::MapNameToOID('SHA256')
        )
    }
    $sigB64 = Set-EncodedBase64Url $sigBytes
    $jwt = "$toSign.$sigB64"

    # Request access token
    $body = @{
        grant_type = "urn:ietf:params:oauth:grant-type:jwt-bearer"
        assertion  = $jwt
    }

    $formParts = $body.GetEnumerator() |
    ForEach-Object { "$($_.Key)=$($_.Value)" }

    $formBody = $formParts -join "&"

    $response = Invoke-RestMethod -Method Post -Uri $TokenUri `
        -ContentType "application/x-www-form-urlencoded" `
        -Body $formBody

    # Outoput the access token
    return $response.access_token
}


function Export-CsvToGoogleSheet {
param (
    [Parameter(Mandatory = $true)]
    [string]$SpreadsheetID,     # The unique ID of the Google Spreadsheet (found in its URL)
    [string]$SheetName,         # Name of the target sheet/tab within the spreadsheet
    [string]$StartCell,         # Starting cell for the data (upper-left corner for placing the CSV data)
    [string]$CsvPath,           # Path to the CSV file containing data to upload
    [string]$Account,           # Email address of service account created in Google Cloud (https://console.cloud.google.com/)
    [string]$CertPath,          # Name of certificate file downloaded from Google Cloud. Must be saved in .\config directory
    [string]$CertPswd           # Certificate password
    )

    # Step 1: Provide the access token for Google Sheets API.
    #         This token should be obtained via OAuth2 (e.g., service account or OAuth client flow) and must include the Sheets scope.
    $accessToken = Get-GoogleAccessToken -CertPath $CertPath

    # Step 2: Read the CSV file content.
    #         Use Import-Csv for robust CSV parsing (handles quoted fields).
    #         Specify UTF8 encoding to preserve non-English characters (e.g., accents, symbols).
    if (!(Test-Path -LiteralPath $CsvPath)) {
        Write-Error "CSV file not found at path: $CsvPath"
        return
    }
    $csvData = Import-Csv -Path $CsvPath -Encoding UTF8

    # Step 3: Prepare the data in the format required by Google Sheets API.
    #         The API expects a ValueRange object with 'values' as an array of rows (each row itself an array).
    #         We'll construct a 2D array ($values) where each sub-array is one row of cell values.
    $values = @()

    # If the CSV has a header row, include it as the first row of values.
    # Import-Csv uses the first line as header by default. We can retrieve header names from the first object.
    if ($csvData.Count -gt 0) {
        # Get the column order from the CSV header (property names of the first object).
        $headerColumns = $csvData[0].PSObject.Properties | Select-Object -ExpandProperty Name
        # Add the header row values (as an array of column names) to $values.
        $values += ,($headerColumns)  # use unary comma to add as single array element
    }

    # Add each data row from the CSV.
    foreach ($row in $csvData) {
        # For each row object, extract cell values in the same column order as the header.
        $rowValues = @()
        foreach ($col in $headerColumns) {
            $rowValues += $row.$col  # preserve original data as string (Import-Csv outputs strings for each field)
        }
        $values += ,($rowValues)
    }

    # Now $values is a 2D array ([[col1, col2, ...], [val1, val2, ...], [val1, val2, ...], ...]) ready to send.

    # Step 4: Construct the target A1 notation range for the update.
    #         We have the starting cell ($StartCell, e.g. "B2"), and we know how many rows and columns the data has.
    #         Calculate the ending cell coordinates to form "SheetName!<start>:<end>".
    #         Also, handle quoting the sheet name if it contains spaces or special characters.
    if ($StartCell -notmatch "^[A-Za-z]+[0-9]+$") {
        Write-Error "Start cell '$StartCell' is not in valid A1 notation (e.g., \"A1\" or \"C5\")."
        return
    }
    # Split the start cell into column letters and row number
    $startColumnLetters = ($StartCell -replace "\d", "")      # e.g. "B" from "B2"
    $startRowNumber     = [int]($StartCell -replace "\D", "") # e.g. 2 from "B2"
    # Convert start column letters to numeric index
    $startColIndex = Get-ColumnIndexFromLetter -ColumnLetters $startColumnLetters
    # Determine size of the data
    $totalRows    = $values.Count            # total rows to write (including header)
    $totalColumns = ($values[0]).Count       # number of columns in each row
    # Calculate ending cell coordinates
    $endColIndex  = $startColIndex + $totalColumns - 1
    $endRowNumber = $startRowNumber + $totalRows - 1
    $endColumnLetters = Get-ColumnLetterFromIndex -ColumnIndex $endColIndex
    # Build the full range string.
    # If sheet name contains spaces or special chars, wrap it in single quotes per A1 notation rules.
    $SheetNameEscaped = $SheetName
    if ($SheetNameEscaped -match "[^A-Za-z0-9_]") {
        $SheetNameEscaped = "'" + $SheetNameEscaped + "'"
    }
    $rangeA1 = "$SheetNameEscaped!$startColumnLetters$startRowNumber`:$endColumnLetters$endRowNumber"
    # (The backtick before : is to ensure the variable expansion is clear; it joins the start and end cell addresses.)

    # Step 5: Create the JSON payload for the Sheets API.
    #         It should include the range, majorDimension, and values as defined by the API.
    $bodyObject = @{
        range          = $rangeA1
        majorDimension = "ROWS"   # we're sending data row by row
        values         = $values  # the 2D array of values to write
    }
    # Convert the body to a JSON string. Use a sufficient depth to include all nested arrays.
    # (Default -Depth is 2, which is not enough for our nested array structure.)
    $jsonBody = $bodyObject | ConvertTo-Json -Depth 5

    # Step 6: Send the HTTP request to Google Sheets API to update the values.
    #         We'll use Invoke-RestMethod with the proper URI, method, headers, and body.
    #         Attach the access token as a Bearer token in the Authorization header.
    #         Set Content-Type to application/json with UTF-8 to avoid 400 errors due to encoding.
    # Construct the request URI. We include the valueInputOption as a query parameter.
    # Using USER_ENTERED so Google Sheets will interpret data (e.g., "10" as number 10, dates as dates). 
    $updateUri = "https://sheets.googleapis.com/v4/spreadsheets/$($SpreadsheetID)/values/$([uri]::EscapeDataString($rangeA1))?valueInputOption=USER_ENTERED"

    # Prepare headers with authorization.
    $authHeader = @{ "Authorization" = "Bearer $accessToken" }

    # Perform the API call (HTTP PUT request to set values in the specified range).
    try {
        $response = Invoke-RestMethod -Uri $updateUri -Method Put -Headers $authHeader `
                -ContentType "application/json; charset=utf-8" -Body $jsonBody
        # Note: The ContentType header with charset UTF-8 ensures special characters are handled correctly.
        # If ContentType is missing or incorrect, the API may return (400) Bad Request.
    }
    catch {
        Write-Error "Failed to update Google Sheet. Error: $($_.Exception.Message)"
        return
    }

    # Step 7: Check the response and output the result.
    if ($null -ne $response) {
        # The response should contain an UpdateValuesResponse with details of the update.
        # $updatedRange = $response.updatedRange
        # $updatedRows  = $response.updatedRows
        # $updatedCols  = $response.updatedColumns
        # $updatedCells = $response.updatedCells
        Write-Host "Status: Completed`n" -ForegroundColor Green
    }
    else {
        Write-Host "WARNING: No response received from the Google API." -ForegroundColor Yellow
    }
}

function Clear-GoogleSheetRange {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $SpreadsheetId,
        [Parameter(Mandatory)][string] $SheetName,
        [Parameter(Mandatory)][string] $StartCell,
        [Parameter(Mandatory)][string] $EndCell,
        [Parameter(Mandatory)][string] $CertPath
    )

    $accessToken = Get-GoogleAccessToken -CertPath $CertPath

    # Parse StartCell
    if ($StartCell -notmatch '^([A-Z]+)(\d+)$') {
        throw "StartCell '$StartCell' is not valid A1 notation."
    }
    $startColLetters = $Matches[1]
    $startRow        = [int]$Matches[2]

    # Parse EndCell
    if ($EndCell -notmatch '^([A-Z]+)(\d+)$') {
        throw "EndCell '$EndCell' is not valid A1 notation."
    }
    $endColLetters = $Matches[1]
    $endRow        = [int]$Matches[2]

    # Convert to numeric indices & validate ordering
    $startColIndex = Get-ColumnIndexFromLetter -ColumnLetters $startColLetters
    $endColIndex   = Get-ColumnIndexFromLetter -ColumnLetters $endColLetters

    if ($endColIndex -lt $startColIndex) {
        throw "End column '$endColLetters' precedes start column '$startColLetters'."
    }
    if ($endRow -lt $startRow) {
        throw "End row '$endRow' is above start row '$startRow'."
    }

    # Normalize letters (in case users passed lowercase, etc.)
    $normStartCol = Get-ColumnLetterFromIndex -ColumnIndex $startColIndex
    $normEndCol   = Get-ColumnLetterFromIndex -ColumnIndex $endColIndex

    # Build A1‐range, quoting the sheet name only if it contains spaces/special chars
    $escapedSheet = if ($SheetName -match '\W') { "'$SheetName'" } else { $SheetName }
    $rawRange     = "${escapedSheet}!${normStartCol}${startRow}:${normEndCol}${endRow}"

    # URL‐encode the range for the REST call
    $encodedRange = [uri]::EscapeDataString($rawRange)

    # Invoke the Sheets API clear method
    $url     = "https://sheets.googleapis.com/v4/spreadsheets/$SpreadsheetId/values/$encodedRange`:clear"
    $headers = @{ Authorization = "Bearer $AccessToken" }

    try {
        $Response = Invoke-RestMethod -Uri $url -Method Post -Headers $headers | Out-Null
        #Write-Host "Successfully cleared $rawRange"
    }
    catch {
        Write-Error "Failed to clear range: $($_.Exception.Message)"
    }
}

Export-ModuleMember -Function *