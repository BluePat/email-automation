[CmdletBinding()]
param(
    [string]$Mode = '',
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'config.json'),
    [string]$ApprovalDigestPath = '',
    [switch]$ForceUnlock
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ($PSVersionTable.PSVersion.Major -lt 7) { throw 'Automatizace vyžaduje PowerShell 7 nebo novější.' }
if (-not $IsWindows) { throw 'Ostré spuštění je podporované pouze ve Windows.' }

Import-Module (Join-Path $PSScriptRoot 'Siola.Core.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Siola.GoogleSheets.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Siola.Outlook.psm1') -Force
$script:LogDegraded = $false

function Write-SiolaLog {
    param([string]$Level, [string]$Event, [hashtable]$Data = @{})
    $entry = [ordered]@{
        timestamp = [DateTimeOffset]::Now.ToString('o')
        level = $Level
        event = $Event
        runId = $script:RunId
        data = $Data
    }
    $line = $entry | ConvertTo-Json -Depth 8 -Compress
    try { Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8 }
    catch {
        $script:LogDegraded = $true
        Write-Warning "Provozní log nelze zapsat: $($_.Exception.Message)"
    }
    Write-Host "[$Level] $Event"
}

function Get-HeaderMap {
    param([object[]]$HeaderRow)
    $map = [Collections.Generic.Dictionary[string, int]]::new([StringComparer]::Ordinal)
    for ($index = 0; $index -lt $HeaderRow.Count; $index++) {
        $name = ([string]$HeaderRow[$index]).Trim()
        if (-not $name) { continue }
        if ($map.ContainsKey($name)) { throw "Duplicitní hlavička: $name." }
        $map.Add($name, $index)
    }
    $required = @(
        'Výzva', 'Číslo RM', 'Žadatel', 'Název akce', 'Dotace (Kč)',
        'Oslovení - TAJEMNÍK', 'Email - TAJEMNÍK',
        'Oslovení - STAROSTA', 'Email - STAROSTA', 'Stav',
        'Stav STAROSTA', 'Datum e-mailu STAROSTA',
        'Stav TAJEMNÍK', 'Datum e-mailu TAJEMNÍK'
    )
    $missing = @($required | Where-Object { -not $map.ContainsKey($_) })
    if ($missing.Count) { throw "Chybí povinné sloupce: $($missing -join ', ')." }
    return ,$map
}

function Get-RowValue {
    param([object[]]$Row, [Collections.Generic.Dictionary[string, int]]$Headers, [string]$Name)
    $index = $Headers[$Name]
    if ($index -ge $Row.Count -or $null -eq $Row[$index]) { return '' }
    return $Row[$index]
}

function ConvertTo-SourceRows {
    param([object[]]$Values, [Collections.Generic.Dictionary[string, int]]$Headers)
    $rows = [Collections.Generic.List[object]]::new()
    for ($index = 1; $index -lt $Values.Count; $index++) {
        $row = [object[]]$Values[$index]
        $rows.Add([pscustomobject]@{
            RowNumber = $index + 1
            Call = Get-RowValue $row $Headers 'Výzva'
            RmNumber = Get-RowValue $row $Headers 'Číslo RM'
            Applicant = Get-RowValue $row $Headers 'Žadatel'
            ProjectName = Get-RowValue $row $Headers 'Název akce'
            Grant = Get-RowValue $row $Headers 'Dotace (Kč)'
            SecretarySalutation = Get-RowValue $row $Headers 'Oslovení - TAJEMNÍK'
            SecretaryEmail = Get-RowValue $row $Headers 'Email - TAJEMNÍK'
            MayorSalutation = Get-RowValue $row $Headers 'Oslovení - STAROSTA'
            MayorEmail = Get-RowValue $row $Headers 'Email - STAROSTA'
            Status = Get-RowValue $row $Headers 'Stav'
            MayorStatus = Get-RowValue $row $Headers 'Stav STAROSTA'
            SecretaryStatus = Get-RowValue $row $Headers 'Stav TAJEMNÍK'
        })
    }
    return @($rows)
}

function New-TargetCellUpdate {
    param([int]$ColumnIndex, [AllowNull()]$Value)
    return [pscustomobject]@{ ColumnIndex = $ColumnIndex; Value = $(if ($null -eq $Value) { '' } else { $Value }) }
}

function Get-SiolaVerifiedRowTargets {
    param(
        [Parameter(Mandatory)][string]$TargetPrefix,
        [Parameter(Mandatory)][object[]]$ExpectedTargets,
        [Parameter(Mandatory)][object[]]$FreshRows,
        [Parameter(Mandatory)][string]$SpreadsheetId,
        [Parameter(Mandatory)][string]$AccessToken
    )
    $located = @(Get-SiolaRowTargets -SpreadsheetId $SpreadsheetId -AccessToken $AccessToken `
        -TargetPrefix $TargetPrefix)
    if ($located.Count -ne $ExpectedTargets.Count) {
        throw 'Některé dočasné značky řádků chybí nebo přebývají.'
    }
    $expectedValues = @($ExpectedTargets.Value | Sort-Object)
    $actualValues = @($located.Value | Sort-Object)
    if (($expectedValues -join "`n") -cne ($actualValues -join "`n")) {
        throw 'Dočasné značky řádků se změnily.'
    }
    $freshByNumber = @{}
    foreach ($row in $FreshRows) { $freshByNumber[[int]$row.RowNumber] = $row }
    foreach ($target in $located) {
        if (-not $freshByNumber.ContainsKey([int]$target.RowNumber)) {
            throw 'Označený řádek už v tabulce neexistuje.'
        }
        $target | Add-Member -NotePropertyName Row -NotePropertyValue $freshByNumber[[int]$target.RowNumber] -Force
    }
    return $located
}

function Set-SiolaTargetUpdates {
    param(
        [Parameter(Mandatory)][object[]]$Targets,
        [Parameter(Mandatory)][scriptblock]$BuildUpdates,
        [Parameter(Mandatory)][string]$SpreadsheetId,
        [Parameter(Mandatory)][string]$AccessToken
    )
    $payload = @($Targets | ForEach-Object {
        [pscustomobject]@{ TargetValue = $_.Value; Updates = [object[]]@(& $BuildUpdates $_) }
    })
    Set-SiolaRowTargetCells -SpreadsheetId $SpreadsheetId -AccessToken $AccessToken -TargetUpdates $payload
}

function Assert-SiolaTargetApproval {
    param([Parameter(Mandatory)]$Group, [Parameter(Mandatory)][object[]]$Targets)
    $expected = @($Group.ApprovalRowFingerprints | Sort-Object)
    $actual = @($Targets | ForEach-Object { Get-SiolaRowApprovalFingerprint $_.Row } | Sort-Object)
    if ($expected.Count -ne $actual.Count -or ($expected -join "`n") -cne ($actual -join "`n")) {
        throw "Označené řádky pro '$($Group.Applicant)' neodpovídají schválenému obsahu. Nic nebylo odesláno."
    }
}

function Get-ShortFailure {
    param([string]$Message)
    $clean = ([regex]::Replace($Message, '\s+', ' ')).Trim()
    if (-not $clean) { $clean = 'Neznámá chyba odeslání' }
    return $clean.Substring(0, [math]::Min(160, $clean.Length))
}

function Assert-SiolaHeaderLayout {
    param(
        [Parameter(Mandatory)]$Expected,
        [Parameter(Mandatory)]$Actual
    )
    foreach ($headerName in $Expected.Keys) {
        if (-not $Actual.ContainsKey($headerName) -or $Actual[$headerName] -ne $Expected[$headerName]) {
            throw "Pořadí sloupců se během běhu změnilo ($headerName). LIVE byl zastaven."
        }
    }
}

function Resolve-SiolaFreshEligibleGroup {
    param([Parameter(Mandatory)]$Group, [Parameter(Mandatory)][object[]]$FreshRows)
    $expected = [string[]]@($Group.ApprovalRowFingerprints | Sort-Object)
    $matchApplicant = if ($Group.PSObject.Properties['MatchApplicant']) { [string]$Group.MatchApplicant } else { [string]$Group.Applicant }
    $applicantKey = ([regex]::Replace($matchApplicant, '\s+', ' ')).Trim().ToLowerInvariant()
    $sameApplicant = @($FreshRows | Where-Object {
        ([string]$_.Status) -ceq 'K ODESLÁNÍ' -and
        ([regex]::Replace(([string]$_.Applicant), '\s+', ' ')).Trim().ToLowerInvariant() -ceq $applicantKey
    })
    $actual = [string[]]@($sameApplicant | ForEach-Object { Get-SiolaRowApprovalFingerprint $_ } | Sort-Object)
    if ($expected.Count -ne $actual.Count -or ($expected -join "`n") -cne ($actual -join "`n")) {
        throw "Připravené řádky žadatele '$($Group.Applicant)' se od načtení změnily. Nic nebylo odesláno."
    }
    return [pscustomobject]@{ RowNumbers = [int[]]@($sameApplicant.RowNumber); Rows = [object[]]$sameApplicant }
}

function Get-SiolaExpectedClaimFingerprints {
    param(
        [Parameter(Mandatory)]$Group,
        [Parameter(Mandatory)][object[]]$ClaimSourceRows,
        [Parameter(Mandatory)][string]$RunId
    )
    $fingerprints = [Collections.Generic.List[string]]::new()
    foreach ($row in $ClaimSourceRows) {
        $expected = [pscustomobject]@{
            Call = $row.Call; RmNumber = $row.RmNumber; Applicant = $row.Applicant
            ProjectName = $row.ProjectName; Grant = $row.Grant
            SecretarySalutation = $row.SecretarySalutation; SecretaryEmail = $row.SecretaryEmail
            MayorSalutation = $row.MayorSalutation; MayorEmail = $row.MayorEmail
            Status = "ZPRACOVÁVÁ SE | $RunId"
            MayorStatus = $row.MayorStatus; SecretaryStatus = $row.SecretaryStatus
        }
        foreach ($job in $Group.Jobs) {
            if ($job.Role -eq 'STAROSTA') { $expected.MayorStatus = "ZPRACOVÁVÁ SE | $($job.JobId)" }
            else { $expected.SecretaryStatus = "ZPRACOVÁVÁ SE | $($job.JobId)" }
        }
        $fingerprints.Add((Get-SiolaRowApprovalFingerprint $expected))
    }
    return [string[]]@($fingerprints | Sort-Object)
}

function Resolve-SiolaClaimedGroup {
    param(
        [Parameter(Mandatory)]$Group,
        [Parameter(Mandatory)][object[]]$ClaimSourceRows,
        [Parameter(Mandatory)][object[]]$FreshRows,
        [Parameter(Mandatory)][string]$RunId
    )
    $expectedStatus = "ZPRACOVÁVÁ SE | $RunId"
    $matchApplicant = if ($Group.PSObject.Properties['MatchApplicant']) { [string]$Group.MatchApplicant } else { [string]$Group.Applicant }
    $applicantKey = ([regex]::Replace($matchApplicant, '\s+', ' ')).Trim().ToLowerInvariant()
    $newReadyRows = @($FreshRows | Where-Object {
        ([string]$_.Status) -ceq 'K ODESLÁNÍ' -and
        ([regex]::Replace(([string]$_.Applicant), '\s+', ' ')).Trim().ToLowerInvariant() -ceq $applicantKey
    })
    if ($newReadyRows.Count -gt 0) {
        throw "Po rezervaci přibyly připravené řádky pro '$($Group.Applicant)'. Před pokračováním je nutná nová kontrola."
    }
    $claimed = @($FreshRows | Where-Object { ([string]$_.Status) -ceq $expectedStatus })
    $expected = Get-SiolaExpectedClaimFingerprints -Group $Group -ClaimSourceRows $ClaimSourceRows -RunId $RunId
    $expectedSet = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($fingerprint in $expected) { $null = $expectedSet.Add($fingerprint) }
    $matching = @($claimed | Where-Object { $expectedSet.Contains((Get-SiolaRowApprovalFingerprint $_)) })
    $actual = [string[]]@($matching | ForEach-Object { Get-SiolaRowApprovalFingerprint $_ } | Sort-Object)
    if ($expected.Count -ne $actual.Count -or ($expected -join "`n") -cne ($actual -join "`n")) {
        throw "Rezervace řádků pro '$($Group.Applicant)' se změnila. Před dalším krokem zkontrolujte tabulku a Odeslanou poštu."
    }
    return [pscustomobject]@{ RowNumbers = [int[]]@($matching.RowNumber); Rows = [object[]]$matching }
}

function Get-SiolaVerifiedSheetRows {
    param(
        [Parameter(Mandatory)][string]$SpreadsheetId,
        [Parameter(Mandatory)][string]$WorksheetName,
        [Parameter(Mandatory)][string]$AccessToken,
        [Parameter(Mandatory)]$ExpectedHeaders
    )
    $result = Get-SiolaSheetValues -SpreadsheetId $SpreadsheetId -WorksheetName $WorksheetName -AccessToken $AccessToken
    $values = [object[]]$result.Rows
    $actualHeaders = Get-HeaderMap ([object[]]$values[0])
    Assert-SiolaHeaderLayout -Expected $ExpectedHeaders -Actual $actualHeaders
    return @(ConvertTo-SourceRows -Values $values -Headers $actualHeaders)
}

function Write-SiolaPreviewReport {
    param([Parameter(Mandatory)]$Batch, [Parameter(Mandatory)][string]$Path)
    $sections = [Collections.Generic.List[string]]::new()
    foreach ($group in $Batch.Groups) {
        foreach ($job in $group.Jobs) {
            $applicant = [Net.WebUtility]::HtmlEncode([string]$job.Applicant)
            $role = [Net.WebUtility]::HtmlEncode([string]$job.Role)
            $recipient = [Net.WebUtility]::HtmlEncode([string]$job.IntendedTo)
            $subject = [Net.WebUtility]::HtmlEncode([string]$job.Subject)
            $rows = [Net.WebUtility]::HtmlEncode(($job.RowNumbers -join ', '))
            $sections.Add("<section><h2>$applicant – $role</h2><p><strong>Příjemce:</strong> $recipient<br><strong>Řádky:</strong> $rows<br><strong>Předmět:</strong> $subject</p><div class=`"email`">$($job.BodyHtml)</div></section>")
        }
    }
    $html = @"
<!doctype html><html lang="cs"><head><meta charset="utf-8"><title>SIOLA – kontrola všech e-mailů</title>
<style>body{font-family:Arial,sans-serif;max-width:1100px;margin:24px auto;padding:0 16px}section{border-top:4px solid #333399;margin:28px 0;padding-top:12px}.email{border:1px solid #bbb;padding:16px;background:#fff}h1{color:#333399}</style></head>
<body><h1>SIOLA – kontrola všech připravených e-mailů</h1><p>Běh: $($script:RunId). Počet zpráv: $($Batch.JobCount). Tento soubor nic neodesílá.</p>
$($sections -join "`n")
</body></html>
"@
    Set-Content -LiteralPath $Path -Value $html -Encoding UTF8
}

function Show-SiolaFailureNotification {
    param([string]$Message)
    try {
        $notificationDirectory = if ([string]::IsNullOrWhiteSpace([string]$dataDirectory)) {
            Join-Path $env:LOCALAPPDATA 'SIOLA Email Automation\data'
        }
        else { $dataDirectory }
        New-Item -ItemType Directory -Path $notificationDirectory -Force | Out-Null
        $noticePath = Join-Path $notificationDirectory 'ACTION_REQUIRED.txt'
        $notice = "SIOLA automatizace vyžaduje kontrolu.`r`nČas: $([DateTimeOffset]::Now.ToString('yyyy-MM-dd HH:mm:ss zzz'))`r`n$Message`r`nLog: $($script:LogPath)"
        Set-Content -LiteralPath $noticePath -Value $notice -Encoding UTF8
    }
    catch {}
    if ($IsWindows) {
        try { & "$env:SystemRoot\System32\msg.exe" $env:USERNAME "SIOLA automatizace selhala. Otevřete ACTION_REQUIRED.txt v datové složce." 2>$null } catch {}
    }
}

$script:RunId = "$(Get-Date -Format 'yyyyMMddTHHmmss')-$([guid]::NewGuid().ToString('N').Substring(0, 8))"
$dataDirectory = Join-Path $env:LOCALAPPDATA 'SIOLA Email Automation\data'
$script:LogPath = Join-Path $dataDirectory "bootstrap-$($script:RunId).jsonl"
try {
if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
    throw "Konfigurace neexistuje: $ConfigPath. Nejdříve spusťte Install-SiolaAutomation.ps1."
}
$config = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
$effectiveMode = $(if ($Mode) { $Mode.ToUpperInvariant() } else { ([string]$config.mode).ToUpperInvariant() })
if ($effectiveMode -notin @('VALIDATE', 'TEST', 'LIVE')) { throw 'Mode musí být VALIDATE, TEST nebo LIVE.' }
if ($ApprovalDigestPath -and $effectiveMode -ne 'VALIDATE') {
    throw 'ApprovalDigestPath lze použít pouze v režimu VALIDATE.'
}
foreach ($name in @('spreadsheetId', 'worksheetName', 'credentialsPath', 'outlookSenderSmtpAddress')) {
    if (-not ([string]$config.$name).Trim()) { throw "V konfiguraci chybí $name." }
}
if (-not $config.PSObject.Properties['signature'] -or $null -eq $config.signature) {
    throw 'V konfiguraci chybí signature. Spusťte znovu instalaci.'
}
foreach ($name in @('name', 'phone', 'email', 'company', 'address', 'companyEmail', 'companyId')) {
    if (-not $config.signature.PSObject.Properties[$name] -or -not ([string]$config.signature.$name).Trim()) {
        throw "V konfiguraci podpisu chybí $name."
    }
}
foreach ($name in @('email', 'companyEmail')) {
    $signatureEmail = ([string]$config.signature.$name).Trim()
    try {
        $parsedSignatureEmail = [Net.Mail.MailAddress]::new($signatureEmail)
        if ($parsedSignatureEmail.Address -ine $signatureEmail) { throw 'Neplatná adresa' }
    }
    catch { throw "E-mail podpisu $name nemá platný formát." }
}
if ($effectiveMode -eq 'LIVE' -and (-not $config.PSObject.Properties['installationId'] -or
        -not ([string]$config.installationId).Trim())) {
    throw 'V konfiguraci chybí installationId. Spusťte znovu instalaci.'
}
if ($effectiveMode -eq 'LIVE') {
    $currentBinding = "$env:COMPUTERNAME|$([Security.Principal.WindowsIdentity]::GetCurrent().Name)"
    if (-not $config.PSObject.Properties['machineBinding'] -or [string]$config.machineBinding -cne $currentBinding) {
        throw 'Konfigurace patří jinému počítači nebo uživateli. Spusťte instalaci na tomto počítači.'
    }
}
if ($effectiveMode -eq 'TEST' -and (-not $config.PSObject.Properties['testBatchSize'] -or
        [int]$config.testBatchSize -lt 1)) {
    throw 'testBatchSize musí být alespoň 1.'
}
$configuredBatchSize = if ($config.PSObject.Properties['batchSize']) { [int]$config.batchSize } else { 50 }
$delaySeconds = if ($config.PSObject.Properties['delaySeconds']) { [int]$config.delaySeconds } else { 3 }
$confirmationTimeout = if ($config.PSObject.Properties['sendConfirmationTimeoutSeconds']) {
    [int]$config.sendConfirmationTimeoutSeconds
}
else { 120 }
$runDeadlineMinutes = if ($config.PSObject.Properties['runDeadlineMinutes']) { [int]$config.runDeadlineMinutes } else { 300 }
if ($configuredBatchSize -lt 1 -or $configuredBatchSize -gt 100) { throw 'batchSize musí být mezi 1 a 100.' }
if ($delaySeconds -lt 0 -or $delaySeconds -gt 300) { throw 'delaySeconds musí být mezi 0 a 300.' }
if ($confirmationTimeout -lt 30 -or $confirmationTimeout -gt 600) {
    throw 'sendConfirmationTimeoutSeconds musí být mezi 30 a 600.'
}
if ($runDeadlineMinutes -lt 30 -or $runDeadlineMinutes -gt 330) {
    throw 'runDeadlineMinutes musí být mezi 30 a 330.'
}
$runDeadlineUtc = [DateTimeOffset]::UtcNow.AddMinutes($runDeadlineMinutes)
Set-SiolaGoogleRequestDeadline -DeadlineUtc $runDeadlineUtc

$dataDirectory = [Environment]::ExpandEnvironmentVariables([string]$config.dataDirectory)
if (-not $dataDirectory) { throw 'V konfiguraci chybí dataDirectory.' }
$logDirectory = Join-Path $dataDirectory 'logs'
New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
$retentionDays = if ($config.PSObject.Properties['logRetentionDays']) { [int]$config.logRetentionDays } else { 90 }
if ($retentionDays -lt 7) { throw 'logRetentionDays musí být alespoň 7.' }
$previewRetentionDays = if ($config.PSObject.Properties['previewRetentionDays']) {
    [int]$config.previewRetentionDays
}
else { 2 }
if ($previewRetentionDays -lt 1 -or $previewRetentionDays -gt 7) {
    throw 'previewRetentionDays musí být mezi 1 a 7.'
}
Get-ChildItem -LiteralPath $logDirectory -Filter 'run-*.jsonl' -File -ErrorAction SilentlyContinue |
    Where-Object LastWriteTimeUtc -lt ([DateTime]::UtcNow.AddDays(-$retentionDays)) |
    Remove-Item -Force -ErrorAction SilentlyContinue
$previewDirectory = Join-Path $dataDirectory 'previews'
if (Test-Path -LiteralPath $previewDirectory -PathType Container) {
    Get-ChildItem -LiteralPath $previewDirectory -Filter 'preview-*.html' -File -ErrorAction SilentlyContinue |
        Where-Object LastWriteTimeUtc -lt ([DateTime]::UtcNow.AddDays(-$previewRetentionDays)) |
        Remove-Item -Force -ErrorAction SilentlyContinue
}
$script:LogPath = Join-Path $logDirectory "run-$($script:RunId).jsonl"
$lockPath = Join-Path $dataDirectory 'automation.lock'

if ($ForceUnlock) {
    if (Test-Path -LiteralPath $lockPath) {
        $probe = $null
        try { $probe = [IO.File]::Open($lockPath, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None) }
        catch { throw 'Zámek právě používá aktivní proces. Nebyl odstraněn.' }
        finally { if ($null -ne $probe) { $probe.Dispose() } }
        Remove-Item -LiteralPath $lockPath -Force
    }
    Write-Host 'Zámek byl odstraněn. Před LIVE spuštěním ověřte Odeslanou poštu a stavy ZPRACOVÁVÁ SE.'
    return
}
}
catch {
    if ([string]::IsNullOrWhiteSpace([string]$dataDirectory)) {
        $dataDirectory = Join-Path $env:LOCALAPPDATA 'SIOLA Email Automation\data'
        $script:LogPath = Join-Path $dataDirectory "bootstrap-$($script:RunId).jsonl"
    }
    try { New-Item -ItemType Directory -Path $dataDirectory -Force | Out-Null } catch {}
    $bootstrapFailure = Get-ShortFailure $_.Exception.Message
    Write-SiolaLog CRITICAL BOOTSTRAP_FAILED @{ detail = $bootstrapFailure }
    Show-SiolaFailureNotification $bootstrapFailure
    throw
}

$lockStream = $null
$outlook = $null
$tokenSession = $null
$credentialsPath = ''
$leaseAcquired = $false
try {
    try {
        $lockStream = [IO.File]::Open($lockPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        $lockText = [Text.Encoding]::UTF8.GetBytes("runId=$($script:RunId); started=$([DateTimeOffset]::Now.ToString('o'))")
        $lockStream.Write($lockText, 0, $lockText.Length)
        $lockStream.Flush()
    }
    catch { throw "Automatizace je již spuštěná nebo zůstal bezpečnostní zámek: $lockPath" }

    Write-SiolaLog INFO START @{ mode = $effectiveMode }
    $credentialsPath = [Environment]::ExpandEnvironmentVariables([string]$config.credentialsPath)
    $tokenSession = Get-GoogleServiceAccountToken -CredentialsPath $credentialsPath -AsSession
    $accessToken = [string]$tokenSession.AccessToken
    $sheetId = Get-SiolaWorksheetId -SpreadsheetId $config.spreadsheetId `
        -WorksheetName $config.worksheetName -AccessToken $accessToken
    if ($effectiveMode -eq 'LIVE') {
        Assert-SiolaAutomationOwner -SpreadsheetId $config.spreadsheetId -AccessToken $accessToken `
            -InstallationId ([string]$config.installationId)
        $null = Enter-SiolaAutomationLease -SpreadsheetId $config.spreadsheetId -AccessToken $accessToken `
            -InstallationId ([string]$config.installationId) -RunId $script:RunId
        $leaseAcquired = $true
    }
    $sheetResult = Get-SiolaSheetValues -SpreadsheetId $config.spreadsheetId `
        -WorksheetName $config.worksheetName -AccessToken $accessToken
    $values = [object[]]$sheetResult.Rows
    $headers = Get-HeaderMap ([object[]]$values[0])
    $sourceRows = @(ConvertTo-SourceRows -Values $values -Headers $headers)

    # VALIDATE and TEST inspect every eligible applicant. TEST limits only actual test sends.
    $batchSize = if ($effectiveMode -eq 'LIVE') { $configuredBatchSize } else { [int]::MaxValue }
    $batch = Get-SiolaPreparedBatch -Rows $sourceRows -Mode $effectiveMode -RunId $script:RunId `
        -BatchSize $batchSize -TestRecipient ([string]$config.testRecipient) -Signature $config.signature

    Write-SiolaLog INFO PREPARED @{
        eligibleRows = $batch.EligibleRowCount
        applicants = $batch.SelectedApplicantCount
        jobs = $batch.JobCount
        validationErrors = $batch.ValidationErrorCount
    }
    foreach ($invalid in $batch.InvalidGroups) {
        Write-SiolaLog ERROR VALIDATION_ERROR @{
            applicant = $invalid.Applicant
            rows = $invalid.RowNumbers
            errors = $invalid.Errors
        }
    }

    if ($effectiveMode -eq 'VALIDATE') {
        $batch | Select-Object Mode, EligibleRowCount, SelectedApplicantCount, JobCount, ValidationErrorCount |
            Format-List | Out-Host
        if ($batch.ValidationErrorCount -gt 0) { throw 'Kontrola našla chyby. Tabulka nebyla změněna.' }
        if ($ApprovalDigestPath) {
            Get-SiolaBatchApprovalFingerprint -Batch $batch | Set-Content -LiteralPath $ApprovalDigestPath -Encoding UTF8
        }
        Write-SiolaLog INFO VALIDATION_OK @{}
        return
    }

    if ($effectiveMode -eq 'TEST') {
        $receiptPath = Join-Path $dataDirectory 'test-success.json'
        if (Test-Path -LiteralPath $receiptPath -PathType Leaf) {
            try {
                $previousReceipt = Get-Content -LiteralPath $receiptPath -Raw -Encoding UTF8 | ConvertFrom-Json
                if ([string]$previousReceipt.previewPath) {
                    Remove-Item -LiteralPath ([string]$previousReceipt.previewPath) -Force -ErrorAction SilentlyContinue
                }
            }
            catch { Write-SiolaLog WARNING PREVIOUS_PREVIEW_CLEANUP_FAILED @{} }
        }
        Remove-Item -LiteralPath $receiptPath -Force -ErrorAction SilentlyContinue
        if ($batch.ValidationErrorCount -gt 0) { throw 'TEST byl zastaven kvůli chybám validace. Tabulka nebyla změněna.' }
        if ($batch.JobCount -eq 0) { Write-SiolaLog INFO NOTHING_TO_TEST @{}; return }

        New-Item -ItemType Directory -Path $previewDirectory -Force | Out-Null
        $previewPath = Join-Path $previewDirectory "preview-$($script:RunId).html"
        Write-SiolaPreviewReport -Batch $batch -Path $previewPath

        $outlook = Connect-SiolaOutlook -SenderSmtpAddress ([string]$config.outlookSenderSmtpAddress)
        $testGroups = @($batch.Groups | Select-Object -First ([int]$config.testBatchSize))
        [int]$sentJobs = 0
        foreach ($group in $testGroups) {
            foreach ($job in $group.Jobs) {
                Send-SiolaOutlookJob -OutlookContext $outlook -Job $job `
                    -ConfirmationTimeoutSeconds $confirmationTimeout
                $sentJobs++
                Write-SiolaLog INFO TEST_SENT_CONFIRMED @{ jobId = $job.JobId; role = $job.Role; intendedTo = $job.IntendedTo }
                Start-Sleep -Seconds $delaySeconds
            }
        }
        $receipt = [ordered]@{
            completedUtc = [DateTimeOffset]::UtcNow.ToString('o')
            spreadsheetId = [string]$config.spreadsheetId
            worksheetName = [string]$config.worksheetName
            sender = [string]$config.outlookSenderSmtpAddress
            testRecipient = [string]$config.testRecipient
            installationId = [string]$config.installationId
            machineBinding = [string]$config.machineBinding
            previewPath = $previewPath
            previewJobs = [int]$batch.JobCount
            sentTestJobs = $sentJobs
            approvalDigest = Get-SiolaBatchApprovalFingerprint -Batch $batch
            runtimeFingerprint = Get-SiolaRuntimeFingerprint -RootPath $PSScriptRoot
        }
        $receipt | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $receiptPath -Encoding UTF8
        try { Start-Process -FilePath $previewPath | Out-Null } catch {}
        Write-SiolaLog INFO TEST_COMPLETE @{ sentJobs = $sentJobs; previewJobs = $batch.JobCount; sheetChanged = $false }
        return
    }

    # LIVE records invalid/completed groups, then preflights Outlook before claiming any sendable group.
    $nonSendGroups = @($batch.InvalidGroups) + @($batch.CompletedGroups)
    foreach ($nonSendGroup in $nonSendGroups) {
        $targetPrefix = "$($script:RunId)|status|$([guid]::NewGuid().ToString('N'))"
        $createdTargets = @()
        $locatedTargets = @()
        try {
            $accessToken = Get-SiolaFreshAccessToken -Session $tokenSession -CredentialsPath $credentialsPath
            $statusFreshRows = @(Get-SiolaVerifiedSheetRows -SpreadsheetId $config.spreadsheetId `
                -WorksheetName $config.worksheetName -AccessToken $accessToken -ExpectedHeaders $headers)
            $resolved = Resolve-SiolaFreshEligibleGroup -Group $nonSendGroup -FreshRows $statusFreshRows
            $createdTargets = @(New-SiolaRowTargets -SpreadsheetId $config.spreadsheetId -SheetId $sheetId `
                -AccessToken $accessToken -RowNumbers ([int[]]$resolved.RowNumbers) -TargetPrefix $targetPrefix)

            $statusFreshRows = @(Get-SiolaVerifiedSheetRows -SpreadsheetId $config.spreadsheetId `
                -WorksheetName $config.worksheetName -AccessToken $accessToken -ExpectedHeaders $headers)
            $locatedTargets = @(Get-SiolaVerifiedRowTargets -TargetPrefix $targetPrefix `
                -ExpectedTargets $createdTargets -FreshRows $statusFreshRows `
                -SpreadsheetId $config.spreadsheetId -AccessToken $accessToken)
            Assert-SiolaTargetApproval -Group $nonSendGroup -Targets $locatedTargets

            $isInvalid = $batch.InvalidGroups -contains $nonSendGroup
            $detail = if ($isInvalid) { Get-ShortFailure ($nonSendGroup.Errors -join '; ') } else { '' }
            Set-SiolaTargetUpdates -Targets $locatedTargets -SpreadsheetId $config.spreadsheetId `
                -AccessToken $accessToken -BuildUpdates {
                    param($target)
                    if (-not $isInvalid) {
                        return ,(New-TargetCellUpdate $headers['Stav'] 'ODESLÁNO')
                    }
                    $cellUpdates = [Collections.Generic.List[object]]::new()
                    $cellUpdates.Add((New-TargetCellUpdate $headers['Stav'] 'CHYBA VALIDACE'))
                    if (-not (Test-SiolaSentStatus $target.Row.MayorStatus)) {
                        $cellUpdates.Add((New-TargetCellUpdate $headers['Stav STAROSTA'] "CHYBA VALIDACE | $detail"))
                    }
                    if ([string]$target.Row.SecretaryEmail -and -not (Test-SiolaSentStatus $target.Row.SecretaryStatus)) {
                        $cellUpdates.Add((New-TargetCellUpdate $headers['Stav TAJEMNÍK'] "CHYBA VALIDACE | $detail"))
                    }
                    return @($cellUpdates)
                }
        }
        finally {
            if ($targetPrefix) {
                try {
                    $accessToken = Get-SiolaFreshAccessToken -Session $tokenSession -CredentialsPath $credentialsPath
                    $cleanupTargets = @(Get-SiolaRowTargets -SpreadsheetId $config.spreadsheetId `
                        -AccessToken $accessToken -TargetPrefix $targetPrefix)
                    Remove-SiolaRowTargets -SpreadsheetId $config.spreadsheetId -AccessToken $accessToken `
                        -Targets $cleanupTargets
                }
                catch { Write-SiolaLog WARNING ROW_TARGET_CLEANUP_FAILED @{ prefix = $targetPrefix } }
            }
        }
    }

    if ($batch.JobCount -eq 0) {
        Write-SiolaLog INFO NOTHING_TO_SEND @{}
        if ($batch.ValidationErrorCount -gt 0) {
            throw "Tabulka obsahuje $($batch.ValidationErrorCount) chyb validace."
        }
        return
    }

    $outlook = Connect-SiolaOutlook -SenderSmtpAddress ([string]$config.outlookSenderSmtpAddress)
    [int]$processedJobs = 0
    [bool]$deferredForDeadline = $false
    foreach ($group in $batch.Groups) {
        # After a row is marked, Google writes are not retried. Budget every remaining
        # 60-second request, Outlook confirmation, and five minutes for local cleanup.
        $googleCallBudgetSeconds = (([int]$group.Jobs.Count * 6) + 21) * 60
        $worstCaseGroupSeconds = ([int]$group.Jobs.Count * ($confirmationTimeout + $delaySeconds)) + `
            $googleCallBudgetSeconds + 300
        if ([DateTimeOffset]::UtcNow.AddSeconds($worstCaseGroupSeconds) -gt $runDeadlineUtc) {
            $deferredForDeadline = $true
            Write-SiolaLog WARNING RUN_DEADLINE_REACHED @{
                applicant = $group.Applicant
                remainingJobs = $batch.JobCount - $processedJobs
            }
            break
        }
        $targetPrefix = ''
        $createdTargets = @()
        Set-SiolaGoogleRequestMaxAttempts -MaxAttempts 1
        try {
        $accessToken = Get-SiolaFreshAccessToken -Session $tokenSession -CredentialsPath $credentialsPath
        Assert-SiolaAutomationOwner -SpreadsheetId $config.spreadsheetId -AccessToken $accessToken `
            -InstallationId ([string]$config.installationId)
        $null = Update-SiolaAutomationLease -SpreadsheetId $config.spreadsheetId -AccessToken $accessToken `
            -InstallationId ([string]$config.installationId) -RunId $script:RunId

        # Re-read and compare every email-driving value immediately before this applicant is claimed.
        $freshSourceRows = @(Get-SiolaVerifiedSheetRows -SpreadsheetId $config.spreadsheetId `
            -WorksheetName $config.worksheetName -AccessToken $accessToken -ExpectedHeaders $headers)
        $freshGroup = Resolve-SiolaFreshEligibleGroup -Group $group -FreshRows $freshSourceRows
        $targetPrefix = "$($script:RunId)|send|$([guid]::NewGuid().ToString('N'))"
        $createdTargets = @(New-SiolaRowTargets -SpreadsheetId $config.spreadsheetId -SheetId $sheetId `
            -AccessToken $accessToken -RowNumbers ([int[]]$freshGroup.RowNumbers) -TargetPrefix $targetPrefix)
        $markedFreshRows = @(Get-SiolaVerifiedSheetRows -SpreadsheetId $config.spreadsheetId `
            -WorksheetName $config.worksheetName -AccessToken $accessToken -ExpectedHeaders $headers)
        $locatedTargets = @(Get-SiolaVerifiedRowTargets -TargetPrefix $targetPrefix `
            -ExpectedTargets $createdTargets -FreshRows $markedFreshRows `
            -SpreadsheetId $config.spreadsheetId -AccessToken $accessToken)
        Assert-SiolaTargetApproval -Group $group -Targets $locatedTargets
        $claimSourceRows = [object[]]@($locatedTargets.Row)
        Set-SiolaTargetUpdates -Targets $locatedTargets -SpreadsheetId $config.spreadsheetId `
            -AccessToken $accessToken -BuildUpdates {
                param($target)
                $claimUpdates = [Collections.Generic.List[object]]::new()
                $claimUpdates.Add((New-TargetCellUpdate $headers['Stav'] "ZPRACOVÁVÁ SE | $($script:RunId)"))
                foreach ($claimJob in $group.Jobs) {
                    $statusColumn = if ($claimJob.Role -eq 'STAROSTA') { 'Stav STAROSTA' } else { 'Stav TAJEMNÍK' }
                    $claimUpdates.Add((New-TargetCellUpdate $headers[$statusColumn] "ZPRACOVÁVÁ SE | $($claimJob.JobId)"))
                }
                return @($claimUpdates)
            }
        Write-SiolaLog INFO CLAIMED @{ applicant = $group.Applicant; rows = $locatedTargets.RowNumber; jobs = $group.Jobs.Count }

        $outcomes = @{}
        foreach ($job in $group.Jobs) {
            $accessToken = Get-SiolaFreshAccessToken -Session $tokenSession -CredentialsPath $credentialsPath
            Assert-SiolaAutomationOwner -SpreadsheetId $config.spreadsheetId -AccessToken $accessToken `
                -InstallationId ([string]$config.installationId)
            $null = Update-SiolaAutomationLease -SpreadsheetId $config.spreadsheetId -AccessToken $accessToken `
                -InstallationId ([string]$config.installationId) -RunId $script:RunId
            $claimRows = @(Get-SiolaVerifiedSheetRows -SpreadsheetId $config.spreadsheetId `
                -WorksheetName $config.worksheetName -AccessToken $accessToken -ExpectedHeaders $headers)
            $currentTargets = @(Get-SiolaVerifiedRowTargets -TargetPrefix $targetPrefix `
                -ExpectedTargets $createdTargets -FreshRows $claimRows `
                -SpreadsheetId $config.spreadsheetId -AccessToken $accessToken)
            $verifiedClaim = Resolve-SiolaClaimedGroup -Group $group -ClaimSourceRows $claimSourceRows `
                -FreshRows $claimRows -RunId $script:RunId
            $targetRows = @($currentTargets.RowNumber | Sort-Object)
            if (($targetRows -join ',') -cne (@($verifiedClaim.RowNumbers | Sort-Object) -join ',')) {
                throw 'Rezervace se přesunula na jiné řádky než dočasné značky. Odesílání bylo zastaveno.'
            }
            try {
                Send-SiolaOutlookJob -OutlookContext $outlook -Job $job `
                    -ConfirmationTimeoutSeconds $confirmationTimeout
                $outcomes[$job.Role] = [pscustomobject]@{ Success = $true; Detail = ''; Job = $job }
            }
            catch {
                $detail = Get-ShortFailure $_.Exception.Message
                $outcomes[$job.Role] = [pscustomobject]@{ Success = $false; Detail = $detail; Job = $job }
            }
            if ($outcomes[$job.Role].Success) {
                Write-SiolaLog INFO OUTLOOK_SENT_CONFIRMED @{
                    jobId = $job.JobId; role = $job.Role; applicant = $job.Applicant; to = $job.IntendedTo
                }
            }
            else {
                Write-SiolaLog ERROR OUTLOOK_FAILED @{
                    jobId = $job.JobId; role = $job.Role; applicant = $job.Applicant; detail = $outcomes[$job.Role].Detail
                }
            }
            $processedJobs++
            Start-Sleep -Seconds $delaySeconds
        }

        $accessToken = Get-SiolaFreshAccessToken -Session $tokenSession -CredentialsPath $credentialsPath
        Assert-SiolaAutomationOwner -SpreadsheetId $config.spreadsheetId -AccessToken $accessToken `
            -InstallationId ([string]$config.installationId)
        $null = Update-SiolaAutomationLease -SpreadsheetId $config.spreadsheetId -AccessToken $accessToken `
            -InstallationId ([string]$config.installationId) -RunId $script:RunId
        $finalRows = @(Get-SiolaVerifiedSheetRows -SpreadsheetId $config.spreadsheetId `
            -WorksheetName $config.worksheetName -AccessToken $accessToken -ExpectedHeaders $headers)
        $finalTargets = @(Get-SiolaVerifiedRowTargets -TargetPrefix $targetPrefix `
            -ExpectedTargets $createdTargets -FreshRows $finalRows `
            -SpreadsheetId $config.spreadsheetId -AccessToken $accessToken)
        $finalGroup = Resolve-SiolaClaimedGroup -Group $group -ClaimSourceRows $claimSourceRows `
            -FreshRows $finalRows -RunId $script:RunId
        if ((@($finalTargets.RowNumber | Sort-Object) -join ',') -cne (@($finalGroup.RowNumbers | Sort-Object) -join ',')) {
            throw 'Výsledné řádky neodpovídají dočasným značkám. Zkontrolujte Odeslanou poštu.'
        }

        $sentAt = [DateTime]::Now.ToOADate()
        $mayorSent = $group.MayorAlreadySent -or ($outcomes.ContainsKey('STAROSTA') -and $outcomes['STAROSTA'].Success)
        $secretarySent = (-not $group.SecretaryRequired) -or $group.SecretaryAlreadySent -or
            ($outcomes.ContainsKey('TAJEMNIK') -and $outcomes['TAJEMNIK'].Success)
        $anySent = $mayorSent -or ($group.SecretaryRequired -and $secretarySent)
        $anyFailed = @($outcomes.Values | Where-Object { -not $_.Success }).Count -gt 0
        $overall = if ($mayorSent -and $secretarySent) { 'ODESLÁNO' } `
            elseif ($anySent) { 'ČÁSTEČNĚ ODESLÁNO' } `
            elseif ($anyFailed) { 'CHYBA' } else { 'ZPRACOVÁVÁ SE' }
        try {
            $accessToken = Get-SiolaFreshAccessToken -Session $tokenSession -CredentialsPath $credentialsPath
            Set-SiolaTargetUpdates -Targets $finalTargets -SpreadsheetId $config.spreadsheetId `
                -AccessToken $accessToken -BuildUpdates {
                    param($target)
                    $finalUpdates = [Collections.Generic.List[object]]::new()
                    foreach ($outcome in $outcomes.Values) {
                        $statusColumn = if ($outcome.Job.Role -eq 'STAROSTA') { 'Stav STAROSTA' } else { 'Stav TAJEMNÍK' }
                        $dateColumn = if ($outcome.Job.Role -eq 'STAROSTA') { 'Datum e-mailu STAROSTA' } else { 'Datum e-mailu TAJEMNÍK' }
                        if ($outcome.Success) {
                            $finalUpdates.Add((New-TargetCellUpdate $headers[$statusColumn] "ODESLÁNO | $($outcome.Job.JobId)"))
                            $finalUpdates.Add((New-TargetCellUpdate $headers[$dateColumn] $sentAt))
                        }
                        else {
                            $finalUpdates.Add((New-TargetCellUpdate $headers[$statusColumn] "CHYBA | $($outcome.Detail)"))
                        }
                    }
                    $finalUpdates.Add((New-TargetCellUpdate $headers['Stav'] $overall))
                    return @($finalUpdates)
                }
        }
        catch {
            Write-SiolaLog CRITICAL RESULT_WRITE_FAILED @{
                applicant = $group.Applicant
                rows = $group.RowNumbers
                instruction = 'Zkontrolujte Odeslanou poštu. Řádky ponechte jako ZPRACOVÁVÁ SE.'
            }
            throw
        }
        Write-SiolaLog INFO GROUP_COMPLETE @{ applicant = $group.Applicant; status = $overall }
        }
        finally {
            if ($targetPrefix) {
                try {
                    $accessToken = Get-SiolaFreshAccessToken -Session $tokenSession -CredentialsPath $credentialsPath
                    $cleanupTargets = @(Get-SiolaRowTargets -SpreadsheetId $config.spreadsheetId `
                        -AccessToken $accessToken -TargetPrefix $targetPrefix)
                    Remove-SiolaRowTargets -SpreadsheetId $config.spreadsheetId -AccessToken $accessToken `
                        -Targets $cleanupTargets
                }
                catch { Write-SiolaLog WARNING ROW_TARGET_CLEANUP_FAILED @{ prefix = $targetPrefix } }
            }
            Set-SiolaGoogleRequestMaxAttempts -MaxAttempts 4
        }
    }
    Write-SiolaLog INFO COMPLETE @{
        jobs = $processedJobs
        deferredJobs = $(if ($deferredForDeadline) { $batch.JobCount - $processedJobs } else { 0 })
        validationErrors = $batch.ValidationErrorCount
    }
    $resultLabel = if ($batch.ValidationErrorCount -gt 0) { 'CHYBA VALIDACE' } else { 'OK' }
    Set-Content -LiteralPath (Join-Path $dataDirectory 'LAST_RESULT.txt') `
        -Value "$resultLabel | $([DateTimeOffset]::Now.ToString('yyyy-MM-dd HH:mm:ss zzz')) | zpráv: $processedJobs" -Encoding UTF8
    if ($batch.ValidationErrorCount -gt 0) {
        throw "Běh dokončil odesílání, ale našel $($batch.ValidationErrorCount) chyb validace."
    }
    elseif ($script:LogDegraded) {
        Show-SiolaFailureNotification 'Běh dokončil práci, ale provozní log nebylo možné úplně zapsat.'
    }
    else {
        Remove-Item -LiteralPath (Join-Path $dataDirectory 'ACTION_REQUIRED.txt') -Force -ErrorAction SilentlyContinue
    }
}
catch {
    $failureDetail = Get-ShortFailure $_.Exception.Message
    try { Write-SiolaLog CRITICAL STOPPED @{ detail = $failureDetail } } catch {}
    Show-SiolaFailureNotification $failureDetail
    throw
}
finally {
    if ($leaseAcquired -and $null -ne $tokenSession -and $credentialsPath) {
        try {
            $accessToken = Get-SiolaFreshAccessToken -Session $tokenSession -CredentialsPath $credentialsPath
            Exit-SiolaAutomationLease -SpreadsheetId $config.spreadsheetId -AccessToken $accessToken `
                -InstallationId ([string]$config.installationId) -RunId $script:RunId
        }
        catch { Write-Warning 'Běhový zámek se nepodařilo odstranit; automaticky vyprší.' }
    }
    Disconnect-SiolaOutlook $outlook
    if ($null -ne $lockStream) {
        $lockStream.Dispose()
        Remove-Item -LiteralPath $lockPath -Force -ErrorAction SilentlyContinue
    }
}
