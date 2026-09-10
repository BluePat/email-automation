Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'Siola.Core.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Siola.GoogleSheets.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Siola.Outlook.psm1') -Force

function Assert-Siola {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "Self-test selhal: $Message" }
}

function New-FixtureRow {
    param([int]$RowNumber = 2)
    return [pscustomobject]@{
        RowNumber = $RowNumber
        Call = 'RES+3/2022'
        RmNumber = '12345'
        Applicant = 'Obec Příkladov'
        ProjectName = 'Instalace FVE v obci Příkladov'
        Grant = 1057223
        SecretarySalutation = 'Vážená paní tajemnice Vzorová,'
        SecretaryEmail = 'secretary@example.com'
        MayorSalutation = 'Vážený pane starosto Vzorový,'
        MayorEmail = 'mayor@example.com'
        Status = 'K ODESLÁNÍ'
        MayorStatus = ''
        SecretaryStatus = ''
    }
}

$signature = [pscustomobject]@{
    Name = 'Jan Vzorový'
    Phone = '+420 000 000 000'
    Email = 'sender@example.com'
    Company = 'Ukázková firma s.r.o.'
    Address = 'Vzorová 123, 100 00 Praha'
    CompanyEmail = 'info@example.com'
    CompanyId = '00000000'
}

$base = New-FixtureRow
$result = Get-SiolaPreparedBatch -Rows @($base) -Mode TEST -RunId selftest -BatchSize 50 `
    -TestRecipient test@example.com -Signature $signature
Assert-Siola ($result.JobCount -eq 2) 'musí vzniknout dva samostatné e-maily'
$mayor = @($result.Groups[0].Jobs | Where-Object Role -eq 'STAROSTA')[0]
$secretary = @($result.Groups[0].Jobs | Where-Object Role -eq 'TAJEMNIK')[0]
Assert-Siola ($mayor.IntendedTo -eq 'mayor@example.com') 'starosta musí použít Email - STAROSTA'
Assert-Siola ($mayor.BodyHtml.Contains('Vážený pane starosto Vzorový,')) 'starosta musí použít Oslovení - STAROSTA'
Assert-Siola ($secretary.IntendedTo -eq 'secretary@example.com') 'tajemník musí použít Email - TAJEMNÍK'
Assert-Siola ($secretary.BodyHtml.Contains('Vážená paní tajemnice Vzorová,')) 'tajemník musí použít Oslovení - TAJEMNÍK'
Assert-Siola ($mayor.To -eq 'test@example.com' -and $secretary.To -eq 'test@example.com') 'TEST musí přesměrovat příjemce'

$validateResult = Get-SiolaPreparedBatch -Rows @($base) -Mode VALIDATE -RunId other-run -BatchSize 50 `
    -Signature $signature
Assert-Siola ((Get-SiolaBatchApprovalFingerprint $result) -ceq
    (Get-SiolaBatchApprovalFingerprint $validateResult)) 'TEST a VALIDATE musí mít stejný schvalovací otisk'
$changedContent = New-FixtureRow
$changedContent.ProjectName = 'Jiný projekt v obci Příkladov'
$changedResult = Get-SiolaPreparedBatch -Rows @($changedContent) -Mode VALIDATE -RunId other-run -BatchSize 50 `
    -Signature $signature
Assert-Siola ((Get-SiolaBatchApprovalFingerprint $result) -cne
    (Get-SiolaBatchApprovalFingerprint $changedResult)) 'změna obsahu musí změnit schvalovací otisk'
$movedRow = New-FixtureRow -RowNumber 99
Assert-Siola ((Get-SiolaRowApprovalFingerprint $base) -ceq
    (Get-SiolaRowApprovalFingerprint $movedRow)) 'přesun řádku nesmí změnit obsahový otisk'

$notExact = New-FixtureRow
$notExact.Status = ' K ODESLÁNÍ '
$exactResult = Get-SiolaPreparedBatch -Rows @($notExact) -Mode VALIDATE -RunId selftest -BatchSize 50 -Signature $signature
Assert-Siola ($exactResult.JobCount -eq 0) 'stav musí být přesná hodnota K ODESLÁNÍ'

$duplicate = New-FixtureRow -RowNumber 3
$duplicateResult = Get-SiolaPreparedBatch -Rows @($base, $duplicate) -Mode VALIDATE -RunId selftest -BatchSize 50 -Signature $signature
Assert-Siola ($duplicateResult.JobCount -eq 2) 'duplicitní Číslo RM se nesmí odeslat dvakrát'
Assert-Siola ($duplicateResult.Groups[0].Jobs[0].BodyHtml.Contains('1.057.223,- Kč')) 'duplicitní dotace se nesmí sečíst dvakrát'

$conflict = New-FixtureRow -RowNumber 3
$conflict.Grant = 42
$conflictResult = Get-SiolaPreparedBatch -Rows @($base, $conflict) -Mode VALIDATE -RunId selftest -BatchSize 50 -Signature $signature
Assert-Siola ($conflictResult.ValidationErrorCount -eq 1 -and $conflictResult.JobCount -eq 0) 'rozporná duplicita musí zastavit skupinu'

$historical = New-FixtureRow -RowNumber 2
$historical.Status = 'ODESLÁNO'
$historical.MayorStatus = 'ODESLÁNO | old-job'
$historical.SecretaryStatus = 'ODESLÁNO | old-job-2'
$newDuplicate = New-FixtureRow -RowNumber 3
$historicalResult = Get-SiolaPreparedBatch -Rows @($historical, $newDuplicate) -Mode VALIDATE -RunId selftest -BatchSize 50 -Signature $signature
Assert-Siola ($historicalResult.ValidationErrorCount -eq 1 -and $historicalResult.JobCount -eq 0) 'historicky odeslané RM se nesmí znovu připravit'

$blankDuplicate = New-FixtureRow -RowNumber 3
$blankDuplicate.MayorEmail = ''
$blankResult = Get-SiolaPreparedBatch -Rows @($base, $blankDuplicate) -Mode VALIDATE -RunId selftest -BatchSize 50 -Signature $signature
Assert-Siola ($blankResult.ValidationErrorCount -eq 1) 'povinný údaj nesmí být převzat z jiného řádku skupiny'

Assert-Siola (Test-SiolaSentStatus 'ODESLÁNO') 'přesný stav ODESLÁNO musí být rozpoznán'
Assert-Siola (Test-SiolaSentStatus 'ODESLÁNO | job-123') 'stav ODESLÁNO s identifikátorem musí být rozpoznán'
Assert-Siola (-not (Test-SiolaSentStatus 'ODESLÁNO NE')) 'náhodný prefix ODESLÁNO nesmí být považován za odeslaný'

$usGrant = New-FixtureRow
$usGrant.Grant = '1,057.22'
$usGrantResult = Get-SiolaPreparedBatch -Rows @($usGrant) -Mode VALIDATE -RunId selftest -BatchSize 50 -Signature $signature
Assert-Siola ($usGrantResult.ValidationErrorCount -eq 1) 'US textový formát dotace se nesmí tiše chybně převést'

$czechGrant = New-FixtureRow
$czechGrant.Grant = '1.057.223,50 Kč'
$czechGrantResult = Get-SiolaPreparedBatch -Rows @($czechGrant) -Mode VALIDATE -RunId selftest -BatchSize 50 -Signature $signature
Assert-Siola ($czechGrantResult.ValidationErrorCount -eq 0 -and $czechGrantResult.JobCount -eq 2) 'jednoznačný český formát dotace musí projít'

$orphanSalutation = New-FixtureRow
$orphanSalutation.SecretaryEmail = ''
$orphanResult = Get-SiolaPreparedBatch -Rows @($orphanSalutation) -Mode VALIDATE -RunId selftest -BatchSize 50 -Signature $signature
Assert-Siola ($orphanResult.ValidationErrorCount -eq 1) 'oslovení tajemníka bez e-mailu musí být chyba'

Write-Host 'SIOLA self-test: OK'
