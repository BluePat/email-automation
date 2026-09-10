Set-StrictMode -Version Latest
$script:RequestDeadlineUtc = [DateTimeOffset]::MaxValue
$script:RequestMaxAttempts = 4

function Set-SiolaGoogleRequestDeadline {
    [CmdletBinding()]
    param([Parameter(Mandatory)][DateTimeOffset]$DeadlineUtc)
    $script:RequestDeadlineUtc = $DeadlineUtc
}

function Set-SiolaGoogleRequestMaxAttempts {
    [CmdletBinding()]
    param([ValidateRange(1, 4)][int]$MaxAttempts)
    $script:RequestMaxAttempts = $MaxAttempts
}

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
    $MaxAttempts = [math]::Min($MaxAttempts, $script:RequestMaxAttempts)
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        $remainingSeconds = [math]::Floor(($script:RequestDeadlineUtc - [DateTimeOffset]::UtcNow).TotalSeconds)
        if ($remainingSeconds -le 0) { throw 'Časový limit běhu vypršel před požadavkem Google Sheets.' }
        [int]$requestTimeoutSeconds = if ($remainingSeconds -gt 60) { 60 } else { [math]::Max(1, [int]$remainingSeconds) }
        try {
            $arguments = @{
                Method = $Method
                Uri = $Uri
                Headers = @{ Authorization = "Bearer $AccessToken" }
                ErrorAction = 'Stop'
                TimeoutSec = $requestTimeoutSeconds
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
            $httpException = $errorRecord.Exception
            if (-not $httpException.PSObject.Properties['Response'] -and
                $null -ne $httpException.InnerException) {
                $httpException = $httpException.InnerException
            }
            if ($httpException.PSObject.Properties['Response'] -and $null -ne $httpException.Response) {
                $status = [int]$httpException.Response.StatusCode
            }
            elseif ($httpException.PSObject.Properties['StatusCode'] -and $null -ne $httpException.StatusCode) {
                $status = [int]$httpException.StatusCode
            }
            $retryable = $status -in @(408, 429, 500, 502, 503, 504)
            if (-not $retryable -or $attempt -eq $MaxAttempts) { throw }
            [int]$delaySeconds = if ($status -eq 429) { 65 } else { [math]::Pow(2, $attempt - 1) }
            if ($status -eq 429) {
                try {
                    $retryAfter = $httpException.Response.Headers.RetryAfter
                    if ($null -ne $retryAfter -and $null -ne $retryAfter.Delta) {
                        $delaySeconds = [math]::Max($delaySeconds, [math]::Ceiling($retryAfter.Delta.TotalSeconds))
                    }
                }
                catch {}
            }
            $sleepSeconds = [math]::Min($delaySeconds, 180)
            if ([DateTimeOffset]::UtcNow.AddSeconds($sleepSeconds + 1) -ge $script:RequestDeadlineUtc) {
                throw 'Časový limit běhu vypršel během opakování požadavku Google Sheets.'
            }
            Start-Sleep -Seconds $sleepSeconds
        }
    }
}

function Get-SiolaWorksheetId {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SpreadsheetId,
        [Parameter(Mandatory)][string]$WorksheetName,
        [Parameter(Mandatory)][string]$AccessToken
    )
    $fields = [uri]::EscapeDataString('sheets.properties(sheetId,title)')
    $uri = "https://sheets.googleapis.com/v4/spreadsheets/$SpreadsheetId`?fields=$fields"
    $response = Invoke-GoogleSheetsRequest -Method Get -Uri $uri -AccessToken $AccessToken
    $matches = @($response.sheets | Where-Object { [string]$_.properties.title -ceq $WorksheetName })
    if ($matches.Count -ne 1) { throw "List '$WorksheetName' nebyl jednoznačně nalezen." }
    return [int]$matches[0].properties.sheetId
}

function New-SiolaRowTargets {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SpreadsheetId,
        [Parameter(Mandatory)][int]$SheetId,
        [Parameter(Mandatory)][string]$AccessToken,
        [Parameter(Mandatory)][int[]]$RowNumbers,
        [Parameter(Mandatory)][string]$TargetPrefix
    )
    if ($RowNumbers.Count -eq 0) { return @() }
    $requests = [Collections.Generic.List[object]]::new()
    $targets = [Collections.Generic.List[object]]::new()
    foreach ($rowNumber in $RowNumbers) {
        $value = "$TargetPrefix|$([guid]::NewGuid().ToString('N'))"
        $targets.Add([pscustomobject]@{ Value = $value; OriginalRowNumber = $rowNumber })
        $requests.Add([pscustomobject]@{
            createDeveloperMetadata = [pscustomobject]@{
                developerMetadata = [pscustomobject]@{
                    metadataKey = 'siola_row_target'
                    metadataValue = $value
                    location = [pscustomobject]@{
                        dimensionRange = [pscustomobject]@{
                            sheetId = $SheetId
                            dimension = 'ROWS'
                            startIndex = $rowNumber - 1
                            endIndex = $rowNumber
                        }
                    }
                    visibility = 'DOCUMENT'
                }
            }
        })
    }
    $uri = "https://sheets.googleapis.com/v4/spreadsheets/${SpreadsheetId}:batchUpdate"
    $body = [pscustomobject]@{ requests = [object[]]@($requests) }
    # Creating metadata is not idempotent. Never replay an ambiguous POST.
    $null = Invoke-GoogleSheetsRequest -Method Post -Uri $uri -AccessToken $AccessToken -Body $body -MaxAttempts 1
    return @($targets)
}

function Get-SiolaRowTargets {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SpreadsheetId,
        [Parameter(Mandatory)][string]$AccessToken,
        [Parameter(Mandatory)][string]$TargetPrefix
    )
    $uri = "https://sheets.googleapis.com/v4/spreadsheets/$SpreadsheetId/developerMetadata:search"
    $body = [pscustomobject]@{
        dataFilters = [object[]]@([pscustomobject]@{
            developerMetadataLookup = [pscustomobject]@{
                metadataKey = 'siola_row_target'
                visibility = 'DOCUMENT'
            }
        })
    }
    $response = Invoke-GoogleSheetsRequest -Method Post -Uri $uri -AccessToken $AccessToken -Body $body
    $matches = if ($null -ne $response -and $response.PSObject.Properties['matchedDeveloperMetadata']) {
        @($response.matchedDeveloperMetadata)
    } else { @() }
    $targets = @($matches | ForEach-Object {
        $metadata = $_.developerMetadata
        $value = [string]$metadata.metadataValue
        $range = $metadata.location.dimensionRange
        if ($value.StartsWith("$TargetPrefix|", [StringComparison]::Ordinal) -and
            [string]$range.dimension -ceq 'ROWS' -and ([int]$range.endIndex - [int]$range.startIndex) -eq 1) {
            [pscustomobject]@{
                MetadataId = [int]$metadata.metadataId
                Value = $value
                RowNumber = [int]$range.startIndex + 1
            }
        }
    })
    $duplicates = @($targets | Group-Object Value | Where-Object Count -ne 1)
    if ($duplicates.Count) { throw 'Dočasné značky řádků nejsou jednoznačné.' }
    return $targets
}

function Set-SiolaRowTargetCells {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SpreadsheetId,
        [Parameter(Mandatory)][string]$AccessToken,
        [Parameter(Mandatory)][object[]]$TargetUpdates
    )
    if ($TargetUpdates.Count -eq 0) { return }
    $data = [Collections.Generic.List[object]]::new()
    foreach ($targetUpdate in $TargetUpdates) {
        $updates = @($targetUpdate.Updates)
        $maxColumn = [int](($updates | Measure-Object ColumnIndex -Maximum).Maximum)
        $rowValues = [object[]]::new($maxColumn + 1)
        foreach ($update in $updates) { $rowValues[[int]$update.ColumnIndex] = $update.Value }
        $data.Add([pscustomobject]@{
            dataFilter = [pscustomobject]@{
                developerMetadataLookup = [pscustomobject]@{
                    metadataKey = 'siola_row_target'
                    metadataValue = [string]$targetUpdate.TargetValue
                    visibility = 'DOCUMENT'
                    locationType = 'ROW'
                }
            }
            majorDimension = 'ROWS'
            values = [object[][]]@([object[]]$rowValues)
        })
    }
    $uri = "https://sheets.googleapis.com/v4/spreadsheets/$SpreadsheetId/values:batchUpdateByDataFilter"
    $body = [pscustomobject]@{ valueInputOption = 'RAW'; includeValuesInResponse = $false; data = [object[]]@($data) }
    $response = Invoke-GoogleSheetsRequest -Method Post -Uri $uri -AccessToken $AccessToken -Body $body
    if ([int]$response.totalUpdatedRows -ne $TargetUpdates.Count) {
        throw "Google Sheets upravil $($response.totalUpdatedRows) řádků místo očekávaných $($TargetUpdates.Count)."
    }
}

function Remove-SiolaRowTargets {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SpreadsheetId,
        [Parameter(Mandatory)][string]$AccessToken,
        [Parameter(Mandatory)][object[]]$Targets
    )
    $ids = @($Targets | ForEach-Object { [int]$_.MetadataId } | Select-Object -Unique)
    if ($ids.Count -eq 0) { return }
    $requests = @($ids | ForEach-Object {
        [pscustomobject]@{
            deleteDeveloperMetadata = [pscustomobject]@{
                dataFilter = [pscustomobject]@{ developerMetadataLookup = [pscustomobject]@{ metadataId = $_ } }
            }
        }
    })
    $uri = "https://sheets.googleapis.com/v4/spreadsheets/${SpreadsheetId}:batchUpdate"
    $null = Invoke-GoogleSheetsRequest -Method Post -Uri $uri -AccessToken $AccessToken `
        -Body ([pscustomobject]@{ requests = [object[]]$requests })
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
    # Creating owner metadata is not idempotent. A rerun reconciles an ambiguous result.
    $null = Invoke-GoogleSheetsRequest -Method Post -Uri $uri -AccessToken $AccessToken -Body $body -MaxAttempts 1
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

function Get-SiolaAutomationLeases {
    param(
        [Parameter(Mandatory)][string]$SpreadsheetId,
        [Parameter(Mandatory)][string]$AccessToken
    )
    $uri = "https://sheets.googleapis.com/v4/spreadsheets/$SpreadsheetId/developerMetadata:search"
    $body = [pscustomobject]@{
        dataFilters = [object[]]@([pscustomobject]@{
            developerMetadataLookup = [pscustomobject]@{
                metadataKey = 'siola_automation_lease'
                visibility = 'DOCUMENT'
            }
        })
    }
    $response = Invoke-GoogleSheetsRequest -Method Post -Uri $uri -AccessToken $AccessToken -Body $body
    $matches = if ($null -ne $response -and $response.PSObject.Properties['matchedDeveloperMetadata']) {
        @($response.matchedDeveloperMetadata)
    }
    else { @() }
    return @($matches | ForEach-Object {
        $metadata = $_.developerMetadata
        $parts = ([string]$metadata.metadataValue).Split('|')
        [long]$expiresUnix = 0
        if ($parts.Count -ne 3 -or -not [long]::TryParse($parts[2], [ref]$expiresUnix)) {
            throw 'Běhový zámek SIOLA má neplatný formát.'
        }
        [pscustomobject]@{
            MetadataId = [int]$metadata.metadataId
            InstallationId = [string]$parts[0]
            RunId = [string]$parts[1]
            ExpiresUtc = [DateTimeOffset]::FromUnixTimeSeconds($expiresUnix)
        }
    })
}

function Get-SiolaAutomationLease {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SpreadsheetId,
        [Parameter(Mandatory)][string]$AccessToken
    )
    $leases = @(Get-SiolaAutomationLeases -SpreadsheetId $SpreadsheetId -AccessToken $AccessToken)
    if ($leases.Count -eq 0) { return $null }
    if ($leases.Count -ne 1) { throw 'Google tabulka obsahuje více běhových zámků SIOLA.' }
    return $leases[0]
}

function Set-SiolaAutomationLeaseValue {
    param(
        [Parameter(Mandatory)][string]$SpreadsheetId,
        [Parameter(Mandatory)][string]$AccessToken,
        [Parameter(Mandatory)][string]$InstallationId,
        [Parameter(Mandatory)][string]$RunId,
        [Parameter(Mandatory)][DateTimeOffset]$ExpiresUtc,
        [AllowNull()]$ExistingLease
    )
    $value = "$InstallationId|$RunId|$($ExpiresUtc.ToUnixTimeSeconds())"
    $request = if ($null -eq $ExistingLease) {
        [pscustomobject]@{
            createDeveloperMetadata = [pscustomobject]@{
                developerMetadata = [pscustomobject]@{
                    metadataKey = 'siola_automation_lease'
                    metadataValue = $value
                    location = [pscustomobject]@{ spreadsheet = $true }
                    visibility = 'DOCUMENT'
                }
            }
        }
    }
    else {
        [pscustomobject]@{
            updateDeveloperMetadata = [pscustomobject]@{
                dataFilters = [object[]]@([pscustomobject]@{
                    developerMetadataLookup = [pscustomobject]@{ metadataId = [int]$ExistingLease.MetadataId }
                })
                developerMetadata = [pscustomobject]@{ metadataValue = $value }
                fields = 'metadataValue'
            }
        }
    }
    $uri = "https://sheets.googleapis.com/v4/spreadsheets/${SpreadsheetId}:batchUpdate"
    $body = [pscustomobject]@{ requests = [object[]]@($request) }
    $null = Invoke-GoogleSheetsRequest -Method Post -Uri $uri -AccessToken $AccessToken -Body $body -MaxAttempts 1
}

function Enter-SiolaAutomationLease {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SpreadsheetId,
        [Parameter(Mandatory)][string]$AccessToken,
        [Parameter(Mandatory)][string]$InstallationId,
        [Parameter(Mandatory)][string]$RunId,
        [int]$LeaseSeconds = 1800
    )
    $now = [DateTimeOffset]::UtcNow
    $existing = @(Get-SiolaAutomationLeases -SpreadsheetId $SpreadsheetId -AccessToken $AccessToken)
    $active = @($existing | Where-Object { $_.ExpiresUtc -gt $now })
    if ($active.Count -gt 0) {
        if ($active.Count -eq 1 -and $active[0].InstallationId -ceq $InstallationId -and
            $active[0].RunId -ceq $RunId) { return $active[0] }
        $until = ($active | Sort-Object ExpiresUtc -Descending | Select-Object -First 1).ExpiresUtc
        throw "Tabulku právě zpracovává jiný nebo nejednoznačný běh SIOLA do $($until.ToString('o'))."
    }
    $expires = $now.AddSeconds($LeaseSeconds)
    $value = "$InstallationId|$RunId|$($expires.ToUnixTimeSeconds())"
    $requests = [Collections.Generic.List[object]]::new()
    foreach ($expired in $existing) {
        $requests.Add([pscustomobject]@{
            deleteDeveloperMetadata = [pscustomobject]@{
                dataFilter = [pscustomobject]@{
                    developerMetadataLookup = [pscustomobject]@{ metadataId = [int]$expired.MetadataId }
                }
            }
        })
    }
    $requests.Add([pscustomobject]@{
        createDeveloperMetadata = [pscustomobject]@{
            developerMetadata = [pscustomobject]@{
                metadataKey = 'siola_automation_lease'
                metadataValue = $value
                location = [pscustomobject]@{ spreadsheet = $true }
                visibility = 'DOCUMENT'
            }
        }
    })
    $uri = "https://sheets.googleapis.com/v4/spreadsheets/${SpreadsheetId}:batchUpdate"
    # Delete+create is an atomic generation transition but is not safe to replay.
    $null = Invoke-GoogleSheetsRequest -Method Post -Uri $uri -AccessToken $AccessToken `
        -Body ([pscustomobject]@{ requests = [object[]]@($requests) }) -MaxAttempts 1
    $verified = Get-SiolaAutomationLease -SpreadsheetId $SpreadsheetId -AccessToken $AccessToken
    if ($null -eq $verified -or $verified.InstallationId -cne $InstallationId -or $verified.RunId -cne $RunId) {
        throw 'Nepodařilo se výhradně získat běhový zámek SIOLA.'
    }
    return $verified
}

function Update-SiolaAutomationLease {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SpreadsheetId,
        [Parameter(Mandatory)][string]$AccessToken,
        [Parameter(Mandatory)][string]$InstallationId,
        [Parameter(Mandatory)][string]$RunId,
        [int]$LeaseSeconds = 1800
    )
    $existing = Get-SiolaAutomationLease -SpreadsheetId $SpreadsheetId -AccessToken $AccessToken
    if ($null -eq $existing -or $existing.InstallationId -cne $InstallationId -or $existing.RunId -cne $RunId -or
        $existing.ExpiresUtc -le [DateTimeOffset]::UtcNow) {
        throw 'Tento běh již nevlastní platný běhový zámek SIOLA.'
    }
    # The non-retried update plus verification can take two 60-second requests.
    # Refuse a late renewal so takeover cannot observe expiry during this transition.
    if ($existing.ExpiresUtc -le [DateTimeOffset]::UtcNow.AddSeconds(180)) {
        throw 'Běhový zámek je příliš blízko vypršení pro bezpečné obnovení.'
    }
    $expires = [DateTimeOffset]::UtcNow.AddSeconds($LeaseSeconds)
    Set-SiolaAutomationLeaseValue -SpreadsheetId $SpreadsheetId -AccessToken $AccessToken `
        -InstallationId $InstallationId -RunId $RunId -ExpiresUtc $expires -ExistingLease $existing
    $verified = Get-SiolaAutomationLease -SpreadsheetId $SpreadsheetId -AccessToken $AccessToken
    if ($null -eq $verified -or $verified.InstallationId -cne $InstallationId -or $verified.RunId -cne $RunId) {
        throw 'Obnovení běhového zámku SIOLA se nepodařilo.'
    }
    return $verified
}

function Exit-SiolaAutomationLease {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SpreadsheetId,
        [Parameter(Mandatory)][string]$AccessToken,
        [Parameter(Mandatory)][string]$InstallationId,
        [Parameter(Mandatory)][string]$RunId
    )
    $matching = @(Get-SiolaAutomationLeases -SpreadsheetId $SpreadsheetId -AccessToken $AccessToken |
        Where-Object { $_.InstallationId -ceq $InstallationId -and $_.RunId -ceq $RunId })
    if ($matching.Count -ne 1) { return }
    $uri = "https://sheets.googleapis.com/v4/spreadsheets/${SpreadsheetId}:batchUpdate"
    $body = [pscustomobject]@{
        requests = [object[]]@([pscustomobject]@{
            deleteDeveloperMetadata = [pscustomobject]@{
                dataFilter = [pscustomobject]@{
                    developerMetadataLookup = [pscustomobject]@{ metadataId = [int]$matching[0].MetadataId }
                }
            }
        })
    }
    $null = Invoke-GoogleSheetsRequest -Method Post -Uri $uri -AccessToken $AccessToken -Body $body
}

function Transfer-SiolaAutomationOwner {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SpreadsheetId,
        [Parameter(Mandatory)][string]$AccessToken,
        [Parameter(Mandatory)][string]$InstallationId
    )
    $takeoverRunId = "takeover-$([guid]::NewGuid().ToString('N'))"
    $null = Enter-SiolaAutomationLease -SpreadsheetId $SpreadsheetId -AccessToken $AccessToken `
        -InstallationId $InstallationId -RunId $takeoverRunId
    try {
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
    finally {
        Exit-SiolaAutomationLease -SpreadsheetId $SpreadsheetId -AccessToken $AccessToken `
            -InstallationId $InstallationId -RunId $takeoverRunId
    }
}

Export-ModuleMember -Function Get-GoogleServiceAccountToken, Get-SiolaSheetValues, `
    Get-SiolaAutomationOwner, Register-SiolaAutomationOwner, Assert-SiolaAutomationOwner, `
    Transfer-SiolaAutomationOwner, Get-SiolaFreshAccessToken, Get-SiolaAutomationLease, `
    Enter-SiolaAutomationLease, Update-SiolaAutomationLease, Exit-SiolaAutomationLease, `
    Set-SiolaGoogleRequestDeadline, Set-SiolaGoogleRequestMaxAttempts, Get-SiolaWorksheetId, `
    New-SiolaRowTargets, Get-SiolaRowTargets, `
    Set-SiolaRowTargetCells, Remove-SiolaRowTargets
