Set-StrictMode -Version Latest

function ConvertTo-Base64Url {
    param([byte[]]$Bytes)
    return [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
}

function Get-GoogleServiceAccountToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$CredentialsPath,
        [switch]$AsSession
    )

    if (-not (Test-Path -LiteralPath $CredentialsPath -PathType Leaf)) {
        throw "Soubor Google credentials neexistuje: $CredentialsPath"
    }
    $credentials = Get-Content -LiteralPath $CredentialsPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($credentials.type -ne 'service_account' -or -not $credentials.client_email -or -not $credentials.private_key) {
        throw 'Google credentials musí být JSON klíč účtu typu service_account.'
    }

    $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $headerJson = @{ alg = 'RS256'; typ = 'JWT' } | ConvertTo-Json -Compress
    $claimJson = @{
        iss = [string]$credentials.client_email
        scope = 'https://www.googleapis.com/auth/spreadsheets'
        aud = 'https://oauth2.googleapis.com/token'
        iat = $now
        exp = $now + 3600
    } | ConvertTo-Json -Compress

    $header = ConvertTo-Base64Url ([Text.Encoding]::UTF8.GetBytes($headerJson))
    $claim = ConvertTo-Base64Url ([Text.Encoding]::UTF8.GetBytes($claimJson))
    $unsigned = "$header.$claim"

    $pem = ([string]$credentials.private_key) -replace '-----BEGIN PRIVATE KEY-----', '' `
        -replace '-----END PRIVATE KEY-----', '' -replace '\s', ''
    $keyBytes = [Convert]::FromBase64String($pem)
    $rsa = [Security.Cryptography.RSA]::Create()
    try {
        [int]$bytesRead = 0
        $rsa.ImportPkcs8PrivateKey($keyBytes, [ref]$bytesRead)
        if ($bytesRead -ne $keyBytes.Length) { throw 'Soukromý klíč nebyl načten celý.' }
        $signatureBytes = $rsa.SignData(
            [Text.Encoding]::UTF8.GetBytes($unsigned),
            [Security.Cryptography.HashAlgorithmName]::SHA256,
            [Security.Cryptography.RSASignaturePadding]::Pkcs1
        )
    }
    finally { $rsa.Dispose() }

    $assertion = "$unsigned.$(ConvertTo-Base64Url $signatureBytes)"
    $tokenResponse = Invoke-RestMethod -Method Post -Uri 'https://oauth2.googleapis.com/token' `
        -ContentType 'application/x-www-form-urlencoded' `
        -Body @{ grant_type = 'urn:ietf:params:oauth:grant-type:jwt-bearer'; assertion = $assertion } `
        -TimeoutSec 60
    if (-not $tokenResponse.access_token) { throw 'Google OAuth nevrátil access token.' }
    if ($AsSession) {
        $lifetime = if ($tokenResponse.PSObject.Properties['expires_in']) { [int]$tokenResponse.expires_in } else { 3600 }
        return [pscustomobject]@{
            AccessToken = [string]$tokenResponse.access_token
            ExpiresUtc = [DateTimeOffset]::UtcNow.AddSeconds($lifetime)
        }
    }
    return [string]$tokenResponse.access_token
}

function Get-SiolaFreshAccessToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)][string]$CredentialsPath,
        [int]$MinimumRemainingSeconds = 300
    )
    if (([DateTimeOffset]$Session.ExpiresUtc) -le [DateTimeOffset]::UtcNow.AddSeconds($MinimumRemainingSeconds)) {
        $fresh = Get-GoogleServiceAccountToken -CredentialsPath $CredentialsPath -AsSession
        $Session.AccessToken = $fresh.AccessToken
        $Session.ExpiresUtc = $fresh.ExpiresUtc
    }
    return [string]$Session.AccessToken
}

function Invoke-GoogleSheetsRequest {
    param(
        [Parameter(Mandatory)][string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$AccessToken,
        [AllowNull()]$Body = $null,
        [int]$MaxAttempts = 4
    )
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            $arguments = @{
                Method = $Method
                Uri = $Uri
                Headers = @{ Authorization = "Bearer $AccessToken" }
                ErrorAction = 'Stop'
                TimeoutSec = 60
            }
            if ($null -ne $Body) {
                $arguments.ContentType = 'application/json; charset=utf-8'
                $arguments.Body = ($Body | ConvertTo-Json -Depth 20 -Compress)
            }
            return Invoke-RestMethod @arguments
        }
        catch {
            $errorRecord = $_
            $status = $null
            if ($errorRecord.Exception.Response) { $status = [int]$errorRecord.Exception.Response.StatusCode }
            $retryable = $status -in @(408, 429, 500, 502, 503, 504)
            if (-not $retryable -or $attempt -eq $MaxAttempts) { throw }
            [int]$delaySeconds = if ($status -eq 429) { 65 } else { [math]::Pow(2, $attempt - 1) }
            if ($status -eq 429) {
                try {
                    $retryAfter = $errorRecord.Exception.Response.Headers.RetryAfter
                    if ($null -ne $retryAfter -and $null -ne $retryAfter.Delta) {
                        $delaySeconds = [math]::Max($delaySeconds, [math]::Ceiling($retryAfter.Delta.TotalSeconds))
                    }
                }
                catch {}
            }
            Start-Sleep -Seconds ([math]::Min($delaySeconds, 180))
        }
    }
}

function Get-SiolaSheetValues {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SpreadsheetId,
        [Parameter(Mandatory)][string]$WorksheetName,
        [Parameter(Mandatory)][string]$AccessToken
    )
    $escapedSheet = $WorksheetName.Replace("'", "''")
    $range = [uri]::EscapeDataString("'$escapedSheet'!A:AZ")
    $uri = "https://sheets.googleapis.com/v4/spreadsheets/$SpreadsheetId/values/$range" +
        '?majorDimension=ROWS&valueRenderOption=UNFORMATTED_VALUE&dateTimeRenderOption=FORMATTED_STRING'
    $response = Invoke-GoogleSheetsRequest -Method Get -Uri $uri -AccessToken $AccessToken
    if ($null -eq $response.values -or $response.values.Count -lt 1) {
        throw "List $WorksheetName je prázdný nebo jej nelze načíst."
    }
    return [pscustomobject]@{ Rows = [object[]]@($response.values) }
}

function Set-SiolaSheetCells {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SpreadsheetId,
        [Parameter(Mandatory)][string]$WorksheetName,
        [Parameter(Mandatory)][string]$AccessToken,
        [Parameter(Mandatory)][object[]]$Updates
    )
    if ($Updates.Count -eq 0) { return }
    $escapedSheet = $WorksheetName.Replace("'", "''")
    $data = [Collections.Generic.List[object]]::new()
    foreach ($update in $Updates) {
        $nestedValues = [object[][]]@([object[]]@($update.Value))
        $data.Add([pscustomobject]@{
            range = "'$escapedSheet'!$($update.Cell)"
            majorDimension = 'ROWS'
            values = $nestedValues
        })
    }
    $body = [pscustomobject]@{
        valueInputOption = 'RAW'
        includeValuesInResponse = $false
        data = [object[]]@($data)
    }
    $uri = "https://sheets.googleapis.com/v4/spreadsheets/$SpreadsheetId/values:batchUpdate"
    $null = Invoke-GoogleSheetsRequest -Method Post -Uri $uri -AccessToken $AccessToken -Body $body
}

function Get-SiolaAutomationOwner {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SpreadsheetId,
        [Parameter(Mandatory)][string]$AccessToken
    )
    $uri = "https://sheets.googleapis.com/v4/spreadsheets/$SpreadsheetId/developerMetadata:search"
    $body = [pscustomobject]@{
        dataFilters = [object[]]@([pscustomobject]@{
            developerMetadataLookup = [pscustomobject]@{
                metadataKey = 'siola_automation_owner'
                visibility = 'DOCUMENT'
            }
        })
    }
    $response = Invoke-GoogleSheetsRequest -Method Post -Uri $uri -AccessToken $AccessToken -Body $body
    if ($null -eq $response) { return '' }
    $metadataMatches = if ($null -ne $response.PSObject.Properties['matchedDeveloperMetadata']) {
        @($response.matchedDeveloperMetadata)
    }
    else { @() }
    if ($metadataMatches.Count -eq 0) { return '' }
    $owners = @($metadataMatches | ForEach-Object { [string]$_.developerMetadata.metadataValue } |
        Where-Object { $_ } | Select-Object -Unique)
    if ($owners.Count -ne 1 -or $metadataMatches.Count -ne 1) {
        throw 'Google tabulka obsahuje více zámků SIOLA. LIVE byl bezpečně zastaven.'
    }
    return [string]$owners[0]
}

function Register-SiolaAutomationOwner {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SpreadsheetId,
        [Parameter(Mandatory)][string]$AccessToken,
        [Parameter(Mandatory)][string]$InstallationId
    )
    $existing = Get-SiolaAutomationOwner -SpreadsheetId $SpreadsheetId -AccessToken $AccessToken
    if ($existing) {
        if ($existing -cne $InstallationId) {
            throw "Tabulka je už přiřazena jiné instalaci SIOLA ($existing)."
        }
        return
    }

    $uri = "https://sheets.googleapis.com/v4/spreadsheets/${SpreadsheetId}:batchUpdate"
    $body = [pscustomobject]@{
        requests = [object[]]@([pscustomobject]@{
            createDeveloperMetadata = [pscustomobject]@{
                developerMetadata = [pscustomobject]@{
                    metadataKey = 'siola_automation_owner'
                    metadataValue = $InstallationId
                    location = [pscustomobject]@{ spreadsheet = $true }
                    visibility = 'DOCUMENT'
                }
            }
        })
    }
    $null = Invoke-GoogleSheetsRequest -Method Post -Uri $uri -AccessToken $AccessToken -Body $body
    $verified = Get-SiolaAutomationOwner -SpreadsheetId $SpreadsheetId -AccessToken $AccessToken
    if ($verified -cne $InstallationId) {
        throw 'Nepodařilo se výhradně přiřadit tabulku této instalaci. LIVE byl zastaven.'
    }
}

function Assert-SiolaAutomationOwner {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SpreadsheetId,
        [Parameter(Mandatory)][string]$AccessToken,
        [Parameter(Mandatory)][string]$InstallationId
    )
    $owner = Get-SiolaAutomationOwner -SpreadsheetId $SpreadsheetId -AccessToken $AccessToken
    if (-not $owner) { throw 'Tabulka není přiřazena žádné instalaci. Spusťte ENABLE_LIVE.cmd.' }
    if ($owner -cne $InstallationId) {
        throw "LIVE je povolený pouze pro jinou instalaci SIOLA ($owner)."
    }
}

function Transfer-SiolaAutomationOwner {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SpreadsheetId,
        [Parameter(Mandatory)][string]$AccessToken,
        [Parameter(Mandatory)][string]$InstallationId
    )
    $existing = Get-SiolaAutomationOwner -SpreadsheetId $SpreadsheetId -AccessToken $AccessToken
    if (-not $existing) {
        Register-SiolaAutomationOwner -SpreadsheetId $SpreadsheetId -AccessToken $AccessToken `
            -InstallationId $InstallationId
        return
    }
    if ($existing -ceq $InstallationId) { return }

    $uri = "https://sheets.googleapis.com/v4/spreadsheets/${SpreadsheetId}:batchUpdate"
    $body = [pscustomobject]@{
        requests = [object[]]@([pscustomobject]@{
            updateDeveloperMetadata = [pscustomobject]@{
                dataFilters = [object[]]@([pscustomobject]@{
                    developerMetadataLookup = [pscustomobject]@{
                        metadataKey = 'siola_automation_owner'
                        visibility = 'DOCUMENT'
                    }
                })
                developerMetadata = [pscustomobject]@{ metadataValue = $InstallationId }
                fields = 'metadataValue'
            }
        })
    }
    $null = Invoke-GoogleSheetsRequest -Method Post -Uri $uri -AccessToken $AccessToken -Body $body
    $verified = Get-SiolaAutomationOwner -SpreadsheetId $SpreadsheetId -AccessToken $AccessToken
    if ($verified -cne $InstallationId) { throw 'Převzetí tabulky novou instalací se nepodařilo.' }
}

Export-ModuleMember -Function Get-GoogleServiceAccountToken, Get-SiolaSheetValues, Set-SiolaSheetCells, `
    Get-SiolaAutomationOwner, Register-SiolaAutomationOwner, Assert-SiolaAutomationOwner, `
    Transfer-SiolaAutomationOwner, Get-SiolaFreshAccessToken
