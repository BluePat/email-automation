[CmdletBinding()]
param(
    [string]$Mode = '',
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'config.json'),
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

function New-CellUpdate {
    param([int]$RowNumber, [int]$ColumnIndex, [AllowNull()]$Value)
    return [pscustomobject]@{
        Cell = "$(ConvertTo-A1Column ($ColumnIndex + 1))$RowNumber"
        Value = $(if ($null -eq $Value) { '' } else { $Value })
    }
}

function Get-ShortFailure {
    param([string]$Message)
    $clean = ([regex]::Replace($Message, '\s+', ' ')).Trim()
    if (-not $clean) { $clean = 'Neznámá chyba odeslání' }
    return $clean.Substring(0, [math]::Min(160, $clean.Length))
}

function Get-SiolaRowFingerprint {
    param([Parameter(Mandatory)]$Row)
    $payload = [ordered]@{
        RowNumber = [int]$Row.RowNumber
        Call = [string]$Row.Call
        RmNumber = [string]$Row.RmNumber
        Applicant = [string]$Row.Applicant
        ProjectName = [string]$Row.ProjectName
        Grant = $Row.Grant
        SecretarySalutation = [string]$Row.SecretarySalutation
        SecretaryEmail = [string]$Row.SecretaryEmail
        MayorSalutation = [string]$Row.MayorSalutation
        MayorEmail = [string]$Row.MayorEmail
        Status = [string]$Row.Status
        MayorStatus = [string]$Row.MayorStatus
        SecretaryStatus = [string]$Row.SecretaryStatus
    }
    $bytes = [Text.Encoding]::UTF8.GetBytes(($payload | ConvertTo-Json -Depth 4 -Compress))
    return [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))
}

function Assert-SiolaGroupUnchanged {
    param(
        [Parameter(Mandatory)]$Group,
        [Parameter(Mandatory)][hashtable]$OriginalRows,
        [Parameter(Mandatory)][hashtable]$FreshRows
    )
    foreach ($rowNumber in $Group.RowNumbers) {
        if (-not $FreshRows.ContainsKey($rowNumber)) {
            throw "Řádek $rowNumber už v tabulce neexistuje. LIVE byl zastaven."
        }
        $before = Get-SiolaRowFingerprint $OriginalRows[$rowNumber]
        $after = Get-SiolaRowFingerprint $FreshRows[$rowNumber]
        if ($before -cne $after) {
            throw "Řádek $rowNumber se po validaci změnil. Nebyl odeslán žádný e-mail pro $($Group.Applicant)."
        }
    }
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

$dataDirectory = [Environment]::ExpandEnvironmentVariables([string]$config.dataDirectory)
if (-not $dataDirectory) { throw 'V konfiguraci chybí dataDirectory.' }
$logDirectory = Join-Path $dataDirectory 'logs'
New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
$retentionDays = if ($config.PSObject.Properties['logRetentionDays']) { [int]$config.logRetentionDays } else { 90 }
if ($retentionDays -lt 7) { throw 'logRetentionDays musí být alespoň 7.' }
Get-ChildItem -LiteralPath $logDirectory -Filter 'run-*.jsonl' -File -ErrorAction SilentlyContinue |
    Where-Object LastWriteTimeUtc -lt ([DateTime]::UtcNow.AddDays(-$retentionDays)) |
    Remove-Item -Force -ErrorAction SilentlyContinue
$previewDirectory = Join-Path $dataDirectory 'previews'
if (Test-Path -LiteralPath $previewDirectory -PathType Container) {
    Get-ChildItem -LiteralPath $previewDirectory -Filter 'preview-*.html' -File -ErrorAction SilentlyContinue |
        Where-Object LastWriteTimeUtc -lt ([DateTime]::UtcNow.AddDays(-$retentionDays)) |
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
    if ($effectiveMode -eq 'LIVE') {
        Assert-SiolaAutomationOwner -SpreadsheetId $config.spreadsheetId -AccessToken $accessToken `
            -InstallationId ([string]$config.installationId)
    }
    $sheetResult = Get-SiolaSheetValues -SpreadsheetId $config.spreadsheetId `
        -WorksheetName $config.worksheetName -AccessToken $accessToken
    $values = [object[]]$sheetResult.Rows
    $headers = Get-HeaderMap ([object[]]$values[0])
    $sourceRows = @(ConvertTo-SourceRows -Values $values -Headers $headers)

    # VALIDATE and TEST inspect every eligible applicant. TEST limits only actual test sends.
    $batchSize = if ($effectiveMode -eq 'LIVE') { [int]$config.batchSize } else { [int]::MaxValue }
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
        Write-SiolaLog INFO VALIDATION_OK @{}
        return
    }

    $confirmationTimeout = if ($config.PSObject.Properties['sendConfirmationTimeoutSeconds']) {
        [int]$config.sendConfirmationTimeoutSeconds
    }
    else { 120 }

    if ($effectiveMode -eq 'TEST') {
        $receiptPath = Join-Path $dataDirectory 'test-success.json'
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
                Start-Sleep -Seconds ([int]$config.delaySeconds)
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
        }
        $receipt | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $receiptPath -Encoding UTF8
        try { Start-Process -FilePath $previewPath | Out-Null } catch {}
        Write-SiolaLog INFO TEST_COMPLETE @{ sentJobs = $sentJobs; previewJobs = $batch.JobCount; sheetChanged = $false }
        return
    }

    # LIVE records invalid/completed groups, then preflights Outlook before claiming any sendable group.
    $updates = [Collections.Generic.List[object]]::new()
    $rowByNumber = @{}
    foreach ($row in $sourceRows) { $rowByNumber[[int]$row.RowNumber] = $row }
    $nonSendGroups = @($batch.InvalidGroups) + @($batch.CompletedGroups)
    if ($nonSendGroups.Count -gt 0) {
        $accessToken = Get-SiolaFreshAccessToken -Session $tokenSession -CredentialsPath $credentialsPath
        $statusFreshResult = Get-SiolaSheetValues -SpreadsheetId $config.spreadsheetId `
            -WorksheetName $config.worksheetName -AccessToken $accessToken
        $statusFreshValues = [object[]]$statusFreshResult.Rows
        $statusFreshHeaders = Get-HeaderMap ([object[]]$statusFreshValues[0])
        foreach ($headerName in $headers.Keys) {
            if (-not $statusFreshHeaders.ContainsKey($headerName) -or
                $statusFreshHeaders[$headerName] -ne $headers[$headerName]) {
                throw "Pořadí sloupců se během běhu změnilo ($headerName). LIVE byl zastaven bez zápisu stavů."
            }
        }
        $statusFreshRows = @(ConvertTo-SourceRows -Values $statusFreshValues -Headers $statusFreshHeaders)
        $statusFreshByNumber = @{}
        foreach ($statusFreshRow in $statusFreshRows) {
            $statusFreshByNumber[[int]$statusFreshRow.RowNumber] = $statusFreshRow
        }
        foreach ($nonSendGroup in $nonSendGroups) {
            Assert-SiolaGroupUnchanged -Group $nonSendGroup -OriginalRows $rowByNumber -FreshRows $statusFreshByNumber
        }
    }
    foreach ($invalid in $batch.InvalidGroups) {
        $detail = Get-ShortFailure ($invalid.Errors -join '; ')
        foreach ($rowNumber in $invalid.RowNumbers) {
            $row = $rowByNumber[$rowNumber]
            $updates.Add((New-CellUpdate $rowNumber $headers['Stav'] 'CHYBA VALIDACE'))
            if (-not (Test-SiolaSentStatus $row.MayorStatus)) {
                $updates.Add((New-CellUpdate $rowNumber $headers['Stav STAROSTA'] "CHYBA VALIDACE | $detail"))
            }
            if ([string]$row.SecretaryEmail -and -not (Test-SiolaSentStatus $row.SecretaryStatus)) {
                $updates.Add((New-CellUpdate $rowNumber $headers['Stav TAJEMNÍK'] "CHYBA VALIDACE | $detail"))
            }
        }
    }
    foreach ($group in $batch.CompletedGroups) {
        foreach ($rowNumber in $group.RowNumbers) {
            $updates.Add((New-CellUpdate $rowNumber $headers['Stav'] 'ODESLÁNO'))
        }
    }
    if ($updates.Count) {
        $accessToken = Get-SiolaFreshAccessToken -Session $tokenSession -CredentialsPath $credentialsPath
        Set-SiolaSheetCells -SpreadsheetId $config.spreadsheetId -WorksheetName $config.worksheetName `
            -AccessToken $accessToken -Updates @($updates)
    }

    if ($batch.JobCount -eq 0) {
        Write-SiolaLog INFO NOTHING_TO_SEND @{}
        if ($batch.ValidationErrorCount -gt 0) { Show-SiolaFailureNotification 'Tabulka obsahuje chyby validace.' }
        return
    }

    $outlook = Connect-SiolaOutlook -SenderSmtpAddress ([string]$config.outlookSenderSmtpAddress)
    foreach ($group in $batch.Groups) {
        $accessToken = Get-SiolaFreshAccessToken -Session $tokenSession -CredentialsPath $credentialsPath
        Assert-SiolaAutomationOwner -SpreadsheetId $config.spreadsheetId -AccessToken $accessToken `
            -InstallationId ([string]$config.installationId)

        # Re-read and compare every email-driving value immediately before this applicant is claimed.
        $freshResult = Get-SiolaSheetValues -SpreadsheetId $config.spreadsheetId `
            -WorksheetName $config.worksheetName -AccessToken $accessToken
        $freshValues = [object[]]$freshResult.Rows
        $freshHeaders = Get-HeaderMap ([object[]]$freshValues[0])
        foreach ($headerName in $headers.Keys) {
            if (-not $freshHeaders.ContainsKey($headerName) -or $freshHeaders[$headerName] -ne $headers[$headerName]) {
                throw "Pořadí sloupců se během běhu změnilo ($headerName). LIVE byl zastaven."
            }
        }
        $freshSourceRows = @(ConvertTo-SourceRows -Values $freshValues -Headers $freshHeaders)
        $freshByNumber = @{}
        foreach ($freshRow in $freshSourceRows) { $freshByNumber[[int]$freshRow.RowNumber] = $freshRow }
        Assert-SiolaGroupUnchanged -Group $group -OriginalRows $rowByNumber -FreshRows $freshByNumber

        $claims = [Collections.Generic.List[object]]::new()
        foreach ($rowNumber in $group.RowNumbers) {
            $claims.Add((New-CellUpdate $rowNumber $headers['Stav'] "ZPRACOVÁVÁ SE | $($script:RunId)"))
            foreach ($job in $group.Jobs) {
                $statusColumn = if ($job.Role -eq 'STAROSTA') { 'Stav STAROSTA' } else { 'Stav TAJEMNÍK' }
                $claims.Add((New-CellUpdate $rowNumber $headers[$statusColumn] "ZPRACOVÁVÁ SE | $($job.JobId)"))
            }
        }
        Set-SiolaSheetCells -SpreadsheetId $config.spreadsheetId -WorksheetName $config.worksheetName `
            -AccessToken $accessToken -Updates @($claims)
        Write-SiolaLog INFO CLAIMED @{ applicant = $group.Applicant; rows = $group.RowNumbers; jobs = $group.Jobs.Count }

        $outcomes = @{}
        foreach ($job in $group.Jobs) {
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
            Start-Sleep -Seconds ([int]$config.delaySeconds)
        }

        $final = [Collections.Generic.List[object]]::new()
        $sentAt = [DateTime]::Now.ToOADate()
        foreach ($outcome in $outcomes.Values) {
            $statusColumn = if ($outcome.Job.Role -eq 'STAROSTA') { 'Stav STAROSTA' } else { 'Stav TAJEMNÍK' }
            $dateColumn = if ($outcome.Job.Role -eq 'STAROSTA') { 'Datum e-mailu STAROSTA' } else { 'Datum e-mailu TAJEMNÍK' }
            foreach ($rowNumber in $group.RowNumbers) {
                if ($outcome.Success) {
                    $final.Add((New-CellUpdate $rowNumber $headers[$statusColumn] "ODESLÁNO | $($outcome.Job.JobId)"))
                    $final.Add((New-CellUpdate $rowNumber $headers[$dateColumn] $sentAt))
                }
                else {
                    $final.Add((New-CellUpdate $rowNumber $headers[$statusColumn] "CHYBA | $($outcome.Detail)"))
                }
            }
        }

        $mayorSent = $group.MayorAlreadySent -or ($outcomes.ContainsKey('STAROSTA') -and $outcomes['STAROSTA'].Success)
        $secretarySent = (-not $group.SecretaryRequired) -or $group.SecretaryAlreadySent -or
            ($outcomes.ContainsKey('TAJEMNIK') -and $outcomes['TAJEMNIK'].Success)
        $anySent = $mayorSent -or ($group.SecretaryRequired -and $secretarySent)
        $anyFailed = @($outcomes.Values | Where-Object { -not $_.Success }).Count -gt 0
        $overall = if ($mayorSent -and $secretarySent) { 'ODESLÁNO' } `
            elseif ($anySent) { 'ČÁSTEČNĚ ODESLÁNO' } `
            elseif ($anyFailed) { 'CHYBA' } else { 'ZPRACOVÁVÁ SE' }
        foreach ($rowNumber in $group.RowNumbers) {
            $final.Add((New-CellUpdate $rowNumber $headers['Stav'] $overall))
        }

        try {
            $accessToken = Get-SiolaFreshAccessToken -Session $tokenSession -CredentialsPath $credentialsPath
            Set-SiolaSheetCells -SpreadsheetId $config.spreadsheetId -WorksheetName $config.worksheetName `
                -AccessToken $accessToken -Updates @($final)
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
    Write-SiolaLog INFO COMPLETE @{ jobs = $batch.JobCount; validationErrors = $batch.ValidationErrorCount }
    Set-Content -LiteralPath (Join-Path $dataDirectory 'LAST_RESULT.txt') `
        -Value "OK | $([DateTimeOffset]::Now.ToString('yyyy-MM-dd HH:mm:ss zzz')) | zpráv: $($batch.JobCount)" -Encoding UTF8
    if ($script:LogDegraded) {
        Show-SiolaFailureNotification 'Běh dokončil práci, ale provozní log nebylo možné úplně zapsat.'
    }
    elseif ($batch.ValidationErrorCount -gt 0) {
        Show-SiolaFailureNotification "Běh dokončil odesílání, ale našel $($batch.ValidationErrorCount) chyb validace."
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
    Disconnect-SiolaOutlook $outlook
    if ($null -ne $lockStream) {
        $lockStream.Dispose()
        Remove-Item -LiteralPath $lockPath -Force -ErrorAction SilentlyContinue
    }
}
