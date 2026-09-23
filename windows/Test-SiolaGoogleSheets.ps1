Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Assert-SiolaGoogleTest {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "Google adapter self-test selhal: $Message" }
}

$global:SiolaMockCalls = [Collections.Generic.List[object]]::new()
$global:SiolaMockSleeps = [Collections.Generic.List[int]]::new()
$global:SiolaMockMode = 'normal'
$global:SiolaMockLeaseExpiry = 0L
$global:SiolaMockOAuthAttempts = 0
function global:Invoke-RestMethod {
    param($Method, $Uri, $Headers, $ErrorAction, $TimeoutSec, $ContentType, $Body)
    $parsedBody = if ($Body -is [string]) { $Body | ConvertFrom-Json } else { $Body }
    $global:SiolaMockCalls.Add([pscustomobject]@{ Method = $Method; Uri = $Uri; Body = $parsedBody })
    if ($Uri -like '*oauth2.googleapis.com/token*') {
        $global:SiolaMockOAuthAttempts++
        if ($global:SiolaMockMode -eq 'oauth-transient-once' -and $global:SiolaMockOAuthAttempts -eq 1) {
            throw [Net.Http.HttpRequestException]::new('simulated DNS failure')
        }
        if ($global:SiolaMockMode -eq 'oauth-transient-always') {
            throw [Net.Http.HttpRequestException]::new('simulated timeout')
        }
        if ($global:SiolaMockMode -eq 'oauth-auth-error') {
            $response = [Net.Http.HttpResponseMessage]::new([Net.HttpStatusCode]::BadRequest)
            throw [Microsoft.PowerShell.Commands.HttpResponseException]::new(
                'simulated invalid grant', $response
            )
        }
        return [pscustomobject]@{ access_token = 'mock-token'; expires_in = 3600 }
    }
    if ($Uri -like '*developerMetadata:search*') {
        if ($global:SiolaMockMode -eq 'one-owner') {
            return [pscustomobject]@{ matchedDeveloperMetadata = @([pscustomobject]@{
                developerMetadata = [pscustomobject]@{
                    metadataId = 31
                    metadataValue = 'installation-one'
                }
            }) }
        }
        if ($global:SiolaMockMode -eq 'duplicate-owner') {
            return [pscustomobject]@{ matchedDeveloperMetadata = @(
                [pscustomobject]@{ developerMetadata = [pscustomobject]@{
                    metadataId = 31; metadataValue = 'installation-one'
                } },
                [pscustomobject]@{ developerMetadata = [pscustomobject]@{
                    metadataId = 32; metadataValue = 'installation-two'
                } }
            ) }
        }
        if ($global:SiolaMockMode -eq 'late-lease') {
            return [pscustomobject]@{ matchedDeveloperMetadata = @([pscustomobject]@{
                developerMetadata = [pscustomobject]@{
                    metadataId = 41
                    metadataValue = "installation|run|$global:SiolaMockLeaseExpiry"
                }
            }) }
        }
        return [pscustomobject]@{ matchedDeveloperMetadata = @() }
    }
    if ($global:SiolaMockMode -eq 'ambiguous-create') {
        $response = [Net.Http.HttpResponseMessage]::new([Net.HttpStatusCode]::ServiceUnavailable)
        throw [Microsoft.PowerShell.Commands.HttpResponseException]::new(
            'simulated ambiguous HTTP failure', $response
        )
    }
    if ($Uri -like '*batchUpdateByDataFilter*') { return [pscustomobject]@{ totalUpdatedRows = 1 } }
    return [pscustomobject]@{}
}
function global:Start-Sleep {
    param([int]$Seconds)
    $global:SiolaMockSleeps.Add($Seconds)
}

try {
    Import-Module (Join-Path $PSScriptRoot 'Siola.GoogleSheets.psm1') -Force
    $googleModule = Get-Module Siola.GoogleSheets

    Set-SiolaGoogleRequestDeadline -DeadlineUtc ([DateTimeOffset]::UtcNow.AddMinutes(10))
    $global:SiolaMockMode = 'oauth-transient-once'
    $global:SiolaMockOAuthAttempts = 0
    $global:SiolaMockSleeps.Clear()
    $tokenResponse = & $googleModule { Invoke-GoogleOAuthTokenRequest -Assertion 'test-assertion' }
    Assert-SiolaGoogleTest ($tokenResponse.access_token -ceq 'mock-token' -and
        $global:SiolaMockOAuthAttempts -eq 2 -and $global:SiolaMockSleeps.Count -eq 1) `
        'přechodná síťová chyba OAuth se musí omezeně zopakovat a poté uspět'

    $global:SiolaMockMode = 'oauth-auth-error'
    $global:SiolaMockOAuthAttempts = 0
    $global:SiolaMockSleeps.Clear()
    $authFailed = $false
    try { $null = & $googleModule { Invoke-GoogleOAuthTokenRequest -Assertion 'test-assertion' } }
    catch { $authFailed = $true }
    Assert-SiolaGoogleTest ($authFailed -and $global:SiolaMockOAuthAttempts -eq 1 -and
        $global:SiolaMockSleeps.Count -eq 0) 'OAuth 400 se nesmí opakovat'

    $global:SiolaMockMode = 'oauth-transient-always'
    $global:SiolaMockOAuthAttempts = 0
    $global:SiolaMockSleeps.Clear()
    $transientFailed = $false
    try { $null = & $googleModule { Invoke-GoogleOAuthTokenRequest -Assertion 'test-assertion' } }
    catch { $transientFailed = $true }
    Assert-SiolaGoogleTest ($transientFailed -and $global:SiolaMockOAuthAttempts -eq 4 -and
        $global:SiolaMockSleeps.Count -eq 3) 'trvalá síťová chyba OAuth musí skončit po čtyřech pokusech'

    $global:SiolaMockMode = 'normal'
    $owner = Get-SiolaAutomationOwner -SpreadsheetId example -AccessToken token
    Assert-SiolaGoogleTest ($owner -ceq '') 'nulový počet zámků vlastníka se musí vrátit jako prázdná hodnota'

    $global:SiolaMockMode = 'one-owner'
    $owner = Get-SiolaAutomationOwner -SpreadsheetId example -AccessToken token
    Assert-SiolaGoogleTest ($owner -ceq 'installation-one') 'jeden zámek vlastníka se musí načíst jako skalární ID'

    $global:SiolaMockMode = 'duplicate-owner'
    $duplicateOwnerFailed = $false
    try { $null = Get-SiolaAutomationOwner -SpreadsheetId example -AccessToken token }
    catch { $duplicateOwnerFailed = $true }
    Assert-SiolaGoogleTest $duplicateOwnerFailed 'více zámků vlastníka musí bezpečně zastavit LIVE'

    $global:SiolaMockMode = 'normal'

    $targets = @(New-SiolaRowTargets -SpreadsheetId example -SheetId 7 -AccessToken token `
        -RowNumbers @(9) -TargetPrefix test)
    $createCall = $global:SiolaMockCalls[-1]
    $created = $createCall.Body.requests[0].createDeveloperMetadata.developerMetadata
    Assert-SiolaGoogleTest ($targets.Count -eq 1) 'musí vzniknout právě jedna lokální značka'
    Assert-SiolaGoogleTest ($created.location.dimensionRange.startIndex -eq 8 -and
        $created.location.dimensionRange.endIndex -eq 9) 'řádková značka musí používat nulový GridRange'

    $global:SiolaMockCalls.Clear()
    Set-SiolaRowTargetCells -SpreadsheetId example -AccessToken token -TargetUpdates @(
        [pscustomobject]@{
            TargetValue = 'test|one'
            Updates = @([pscustomobject]@{ ColumnIndex = 3; Value = 'ODESLÁNO' })
        }
    )
    $write = $global:SiolaMockCalls[-1].Body.data[0]
    Assert-SiolaGoogleTest ($write.dataFilter.developerMetadataLookup.metadataValue -ceq 'test|one') `
        'zápis musí filtrovat přesnou hodnotu značky'
    Assert-SiolaGoogleTest ($null -eq $write.values[0][0] -and $write.values[0][3] -ceq 'ODESLÁNO') `
        'řídký zápis musí přeskočit předchozí sloupce'

    $global:SiolaMockCalls.Clear()
    $global:SiolaMockMode = 'ambiguous-create'
    $ambiguousFailed = $false
    try {
        $null = New-SiolaRowTargets -SpreadsheetId example -SheetId 7 -AccessToken token `
            -RowNumbers @(9) -TargetPrefix ambiguous
    }
    catch { $ambiguousFailed = $true }
    Assert-SiolaGoogleTest $ambiguousFailed 'nejednoznačné vytvoření značky musí selhat bezpečně'
    Assert-SiolaGoogleTest ($global:SiolaMockCalls.Count -eq 1) `
        'neidempotentní vytvoření značky se nesmí automaticky opakovat'

    $global:SiolaMockCalls.Clear()
    $global:SiolaMockMode = 'late-lease'
    $global:SiolaMockLeaseExpiry = [DateTimeOffset]::UtcNow.AddSeconds(100).ToUnixTimeSeconds()
    $lateLeaseFailed = $false
    try {
        $null = Update-SiolaAutomationLease -SpreadsheetId example -AccessToken token `
            -InstallationId installation -RunId run
    }
    catch { $lateLeaseFailed = $true }
    $batchCalls = @($global:SiolaMockCalls | Where-Object Uri -like '*:batchUpdate')
    Assert-SiolaGoogleTest ($lateLeaseFailed -and $batchCalls.Count -eq 0) `
        'zámek blízko vypršení se nesmí obnovovat v rizikovém okně'

    $global:SiolaMockCalls.Clear()
    Set-SiolaGoogleRequestDeadline -DeadlineUtc ([DateTimeOffset]::UtcNow.AddSeconds(-1))
    $deadlineFailed = $false
    try {
        $null = Get-SiolaWorksheetId -SpreadsheetId example -WorksheetName sheet -AccessToken token
    }
    catch { $deadlineFailed = $true }
    Assert-SiolaGoogleTest ($deadlineFailed -and $global:SiolaMockCalls.Count -eq 0) `
        'po časovém limitu se nesmí zahájit HTTP požadavek'

    Write-Host 'SIOLA Google adapter self-test: OK'
}
finally {
    Remove-Item Function:\global:Invoke-RestMethod -ErrorAction SilentlyContinue
    Remove-Item Function:\global:Start-Sleep -ErrorAction SilentlyContinue
    Remove-Variable SiolaMockCalls, SiolaMockSleeps, SiolaMockMode, SiolaMockLeaseExpiry, `
        SiolaMockOAuthAttempts -Scope Global -ErrorAction SilentlyContinue
}
