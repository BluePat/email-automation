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

$base = New-FixtureRow
$result = Get-SiolaPreparedBatch -Rows @($base) -Mode TEST -RunId selftest -BatchSize 50 `
    -TestRecipient test@example.com
Assert-Siola ($result.JobCount -eq 2) 'musí vzniknout dva samostatné e-maily'
$mayor = @($result.Groups[0].Jobs | Where-Object Role -eq 'STAROSTA')[0]
$secretary = @($result.Groups[0].Jobs | Where-Object Role -eq 'TAJEMNIK')[0]
Assert-Siola ($mayor.IntendedTo -eq 'mayor@example.com') 'starosta musí použít Email - STAROSTA'
Assert-Siola ([Net.WebUtility]::HtmlDecode($mayor.BodyHtml).Contains('Vážený pane starosto Vzorový,')) 'starosta musí použít Oslovení - STAROSTA'
Assert-Siola ($secretary.IntendedTo -eq 'secretary@example.com') 'tajemník musí použít Email - TAJEMNÍK'
Assert-Siola ([Net.WebUtility]::HtmlDecode($secretary.BodyHtml).Contains('Vážená paní tajemnice Vzorová,')) 'tajemník musí použít Oslovení - TAJEMNÍK'
Assert-Siola ($mayor.To -eq 'test@example.com' -and $secretary.To -eq 'test@example.com') 'TEST musí přesměrovat příjemce'

$validateResult = Get-SiolaPreparedBatch -Rows @($base) -Mode VALIDATE -RunId other-run -BatchSize 50
Assert-Siola ((Get-SiolaBatchApprovalFingerprint $result) -ceq
    (Get-SiolaBatchApprovalFingerprint $validateResult)) 'TEST a VALIDATE musí mít stejný schvalovací otisk'
$changedContent = New-FixtureRow
$changedContent.ProjectName = 'Jiný projekt v obci Příkladov'
$changedResult = Get-SiolaPreparedBatch -Rows @($changedContent) -Mode VALIDATE -RunId other-run -BatchSize 50
Assert-Siola ((Get-SiolaBatchApprovalFingerprint $result) -cne
    (Get-SiolaBatchApprovalFingerprint $changedResult)) 'změna obsahu musí změnit schvalovací otisk'
$movedRow = New-FixtureRow -RowNumber 99
Assert-Siola ((Get-SiolaRowApprovalFingerprint $base) -ceq
    (Get-SiolaRowApprovalFingerprint $movedRow)) 'přesun řádku nesmí změnit obsahový otisk'

$notExact = New-FixtureRow
$notExact.Status = ' K ODESLÁNÍ '
$exactResult = Get-SiolaPreparedBatch -Rows @($notExact) -Mode VALIDATE -RunId selftest -BatchSize 50
Assert-Siola ($exactResult.JobCount -eq 0) 'stav musí být přesná hodnota K ODESLÁNÍ'

$duplicate = New-FixtureRow -RowNumber 3
$duplicateResult = Get-SiolaPreparedBatch -Rows @($base, $duplicate) -Mode VALIDATE -RunId selftest -BatchSize 50
Assert-Siola ($duplicateResult.JobCount -eq 2) 'duplicitní Číslo RM se nesmí odeslat dvakrát'
Assert-Siola ($duplicateResult.Groups[0].Jobs[0].BodyHtml.Contains('1.057.223,- Kč')) 'duplicitní dotace se nesmí sečíst dvakrát'

$conflict = New-FixtureRow -RowNumber 3
$conflict.Grant = 42
$conflictResult = Get-SiolaPreparedBatch -Rows @($base, $conflict) -Mode VALIDATE -RunId selftest -BatchSize 50
Assert-Siola ($conflictResult.ValidationErrorCount -eq 1 -and $conflictResult.JobCount -eq 0) 'rozporná duplicita musí zastavit skupinu'

$historical = New-FixtureRow -RowNumber 2
$historical.Status = 'ODESLÁNO'
$historical.MayorStatus = 'ODESLÁNO | old-job'
$historical.SecretaryStatus = 'ODESLÁNO | old-job-2'
$newDuplicate = New-FixtureRow -RowNumber 3
$historicalResult = Get-SiolaPreparedBatch -Rows @($historical, $newDuplicate) -Mode VALIDATE -RunId selftest -BatchSize 50
Assert-Siola ($historicalResult.ValidationErrorCount -eq 1 -and $historicalResult.JobCount -eq 0) 'historicky odeslané RM se nesmí znovu připravit'

$blankDuplicate = New-FixtureRow -RowNumber 3
$blankDuplicate.MayorEmail = ''
$blankResult = Get-SiolaPreparedBatch -Rows @($base, $blankDuplicate) -Mode VALIDATE -RunId selftest -BatchSize 50
Assert-Siola ($blankResult.ValidationErrorCount -eq 1) 'povinný údaj nesmí být převzat z jiného řádku skupiny'

Assert-Siola (Test-SiolaSentStatus 'ODESLÁNO') 'přesný stav ODESLÁNO musí být rozpoznán'
Assert-Siola (Test-SiolaSentStatus 'ODESLÁNO | job-123') 'stav ODESLÁNO s identifikátorem musí být rozpoznán'
Assert-Siola (-not (Test-SiolaSentStatus 'ODESLÁNO NE')) 'náhodný prefix ODESLÁNO nesmí být považován za odeslaný'

$usGrant = New-FixtureRow
$usGrant.Grant = '1,057.22'
$usGrantResult = Get-SiolaPreparedBatch -Rows @($usGrant) -Mode VALIDATE -RunId selftest -BatchSize 50
Assert-Siola ($usGrantResult.ValidationErrorCount -eq 1) 'US textový formát dotace se nesmí tiše chybně převést'

$czechGrant = New-FixtureRow
$czechGrant.Grant = '1.057.223,50 Kč'
$czechGrantResult = Get-SiolaPreparedBatch -Rows @($czechGrant) -Mode VALIDATE -RunId selftest -BatchSize 50
Assert-Siola ($czechGrantResult.ValidationErrorCount -eq 0 -and $czechGrantResult.JobCount -eq 2) 'jednoznačný český formát dotace musí projít'

$orphanSalutation = New-FixtureRow
$orphanSalutation.SecretaryEmail = ''
$orphanResult = Get-SiolaPreparedBatch -Rows @($orphanSalutation) -Mode VALIDATE -RunId selftest -BatchSize 50
Assert-Siola ($orphanResult.ValidationErrorCount -eq 1) 'oslovení tajemníka bez e-mailu musí být chyba'

foreach ($missingMarker in @('není', 'neni', ' NENÍ ')) {
    $missingSecretary = New-FixtureRow
    $missingSecretary.SecretaryEmail = $missingMarker
    $missingSecretary.SecretarySalutation = ''
    $missingSecretaryResult = Get-SiolaPreparedBatch -Rows @($missingSecretary) -Mode VALIDATE `
        -RunId selftest -BatchSize 50
    Assert-Siola ($missingSecretaryResult.ValidationErrorCount -eq 0) `
        "hodnota '$missingMarker' musí znamenat chybějícího tajemníka"
    Assert-Siola ($missingSecretaryResult.JobCount -eq 1 -and
        $missingSecretaryResult.Groups[0].Jobs[0].Role -eq 'STAROSTA') `
        "pro chybějícího tajemníka '$missingMarker' smí vzniknout jen e-mail starostovi"
}
$missingSecretaryInBothFields = New-FixtureRow
$missingSecretaryInBothFields.SecretaryEmail = 'není'
$missingSecretaryInBothFields.SecretarySalutation = 'neni'
$missingSecretaryInBothResult = Get-SiolaPreparedBatch -Rows @($missingSecretaryInBothFields) `
    -Mode VALIDATE -RunId selftest -BatchSize 50
Assert-Siola ($missingSecretaryInBothResult.ValidationErrorCount -eq 0 -and
    $missingSecretaryInBothResult.JobCount -eq 1) `
    'značka chybějící hodnoty musí být povolena v obou volitelných polích tajemníka'

$missingMayor = New-FixtureRow
$missingMayor.MayorEmail = 'není'
$missingMayor.MayorSalutation = 'neni'
$missingMayorResult = Get-SiolaPreparedBatch -Rows @($missingMayor) -Mode VALIDATE `
    -RunId selftest -BatchSize 50
Assert-Siola ($missingMayorResult.ValidationErrorCount -eq 0) `
    'hodnota neni musí znamenat chybějícího starostu'
Assert-Siola ($missingMayorResult.JobCount -eq 1 -and
    $missingMayorResult.Groups[0].Jobs[0].Role -eq 'TAJEMNIK') `
    'pro chybějícího starostu smí vzniknout jen e-mail tajemníkovi'

$missingBothRecipients = New-FixtureRow
$missingBothRecipients.MayorEmail = 'není'
$missingBothRecipients.MayorSalutation = ''
$missingBothRecipients.SecretaryEmail = 'neni'
$missingBothRecipients.SecretarySalutation = ''
$missingBothResult = Get-SiolaPreparedBatch -Rows @($missingBothRecipients) -Mode VALIDATE `
    -RunId selftest -BatchSize 50
Assert-Siola ($missingBothResult.ValidationErrorCount -eq 1 -and $missingBothResult.JobCount -eq 0) `
    'alespoň jeden příjemce musí být k dispozici'

$orphanMayorSalutation = New-FixtureRow
$orphanMayorSalutation.MayorEmail = 'není'
$orphanMayorResult = Get-SiolaPreparedBatch -Rows @($orphanMayorSalutation) -Mode VALIDATE `
    -RunId selftest -BatchSize 50
Assert-Siola ($orphanMayorResult.ValidationErrorCount -eq 1) `
    'skutečné oslovení starosty bez jeho e-mailu musí být chyba'

$mergedHtml = Merge-SiolaOutlookSignature `
    -MessageHtml '<div id="message">Text zprávy</div>' `
    -SignatureDocumentHtml '<html><body><div id="signature">Výchozí podpis</div></body></html>'
Assert-Siola ($mergedHtml -match '<body><div id="message">Text zprávy</div><div id="signature">') `
    'text zprávy musí být vložen před výchozí podpis Outlooku'
$signatureFingerprintA = Get-SiolaOutlookSignatureFingerprint `
    '<html><body><p>Jan Vzorový</p><img src="cid:image001.png@run-a" alt="Logo"></body></html>'
$signatureFingerprintB = Get-SiolaOutlookSignatureFingerprint `
    '<html><body><p>Jan Vzorový</p><img src="cid:image001.png@run-b" alt="Logo"></body></html>'
$changedSignatureFingerprint = Get-SiolaOutlookSignatureFingerprint `
    '<html><body><p>Jana Vzorová</p><img src="cid:image001.png@run-c" alt="Logo"></body></html>'
Assert-Siola ($signatureFingerprintA -ceq $signatureFingerprintB) `
    'proměnlivé Outlook CID nesmí měnit otisk stejného podpisu'
Assert-Siola ($signatureFingerprintA -cne $changedSignatureFingerprint) `
    'změna textu podpisu musí změnit jeho otisk'

Write-Host 'SIOLA self-test: OK'
