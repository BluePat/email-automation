Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Assert-SiolaGoogleTest {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "Google adapter self-test selhal: $Message" }
}

$global:SiolaMockCalls = [Collections.Generic.List[object]]::new()
$global:SiolaMockMode = 'normal'
$global:SiolaMockLeaseExpiry = 0L
function global:Invoke-RestMethod {
    param($Method, $Uri, $Headers, $ErrorAction, $TimeoutSec, $ContentType, $Body)
    $parsedBody = if ($Body) { $Body | ConvertFrom-Json } else { $null }
    $global:SiolaMockCalls.Add([pscustomobject]@{ Method = $Method; Uri = $Uri; Body = $parsedBody })
    if ($Uri -like '*developerMetadata:search*') {
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

try {
    Import-Module (Join-Path $PSScriptRoot 'Siola.GoogleSheets.psm1') -Force

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
    Remove-Variable SiolaMockCalls, SiolaMockMode, SiolaMockLeaseExpiry -Scope Global -ErrorAction SilentlyContinue
}
