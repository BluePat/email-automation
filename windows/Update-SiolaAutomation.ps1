[CmdletBinding()]
param(
    [string]$InstallDirectory = (Join-Path $env:LOCALAPPDATA 'SIOLA Email Automation'),
    [string]$TaskName = 'SIOLA Email Automation'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Assert-RequiredConfigText {
    param(
        [Parameter(Mandatory)][object]$Config,
        [Parameter(Mandatory)][string]$Name
    )
    if (-not $Config.PSObject.Properties[$Name] -or
        [string]::IsNullOrWhiteSpace([string]$Config.$Name)) {
        throw "V existující konfiguraci chybí '$Name'. Spusťte znovu úplnou instalaci přes INSTALL.cmd."
    }
}

if ($PSVersionTable.PSVersion.Major -lt 7 -or -not $IsWindows) {
    throw 'Aktualizaci spusťte ve Windows v PowerShellu 7.'
}

$configPath = Join-Path $InstallDirectory 'config.json'
$task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($null -eq $task) {
    throw "Naplánovaná úloha '$TaskName' nebyla nalezena. Použijte úplnou instalaci přes INSTALL.cmd."
}
Disable-ScheduledTask -TaskName $TaskName | Out-Null
$task = Get-ScheduledTask -TaskName $TaskName
if ([string]$task.State -ceq 'Running') {
    throw "Úloha '$TaskName' právě běží. Byla vypnuta; aktualizaci opakujte až po jejím skončení."
}

if (-not (Test-Path -LiteralPath $configPath -PathType Leaf)) {
    throw "Instalace nebyla nalezena v '$InstallDirectory'. Nejdříve spusťte INSTALL.cmd."
}

try { $config = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json }
catch { throw "Existující config.json nelze načíst. Programové soubory ani konfigurace nebyly změněny: $($_.Exception.Message)" }

foreach ($requiredName in @(
    'spreadsheetId', 'worksheetName', 'credentialsPath', 'dataDirectory',
    'outlookSenderSmtpAddress', 'testRecipient', 'installationId', 'machineBinding'
)) {
    Assert-RequiredConfigText -Config $config -Name $requiredName
}

$identity = [Security.Principal.WindowsIdentity]::GetCurrent().Name
$currentBinding = "$env:COMPUTERNAME|$identity"
if ([string]$config.machineBinding -cne $currentBinding) {
    throw 'Konfigurace patří jinému počítači nebo uživateli. Použijte úplnou instalaci a případně TAKE_OVER.cmd.'
}
if ([string]$config.spreadsheetId -notmatch '^[A-Za-z0-9_-]+$') {
    throw 'Uložené Google Sheet ID nemá platný formát. Spusťte znovu úplnou instalaci.'
}
foreach ($emailSetting in @('outlookSenderSmtpAddress', 'testRecipient')) {
    $emailValue = [string]$config.$emailSetting
    try {
        $parsedEmail = [Net.Mail.MailAddress]::new($emailValue)
        if ($parsedEmail.Address -ine $emailValue) { throw 'Neplatná adresa' }
    }
    catch { throw "Uložená hodnota '$emailSetting' nemá platný e-mailový formát." }
}

$credentialsPath = [Environment]::ExpandEnvironmentVariables([string]$config.credentialsPath)
if (-not (Test-Path -LiteralPath $credentialsPath -PathType Leaf)) {
    throw "Uložený Google klíč nebyl nalezen: $credentialsPath"
}
try { $credentialCheck = Get-Content -LiteralPath $credentialsPath -Raw -Encoding UTF8 | ConvertFrom-Json }
catch { throw "Uložený Google klíč nelze načíst: $($_.Exception.Message)" }
if ($credentialCheck.type -ne 'service_account' -or -not $credentialCheck.client_email) {
    throw 'Uložený Google klíč není platný service-account JSON soubor.'
}

$dataDirectory = [Environment]::ExpandEnvironmentVariables([string]$config.dataDirectory)
if ([string]::IsNullOrWhiteSpace($dataDirectory)) { throw 'Uložená datová složka je prázdná.' }
if (-not (Test-Path -LiteralPath $dataDirectory -PathType Container)) {
    throw "Uložená datová složka neexistuje: $dataDirectory"
}

$runtimeFiles = @(
    'Invoke-SiolaAutomation.ps1', 'Siola.Core.psm1', 'Siola.GoogleSheets.psm1', 'Siola.Outlook.psm1',
    'Run-Validation.ps1', 'Run-Test.ps1', 'Enable-SiolaLive.ps1', 'Test-SiolaCore.ps1',
    'Test-SiolaOutlookDiagnostic.ps1', 'VALIDATE.cmd', 'TEST.cmd', 'ENABLE_LIVE.cmd',
    'TAKE_OVER.cmd', 'OUTLOOK_DIAGNOSTIC.cmd'
)
foreach ($file in $runtimeFiles) {
    if (-not (Test-Path -LiteralPath (Join-Path $PSScriptRoot $file) -PathType Leaf)) {
        throw "Aktualizační balíček není úplný; chybí soubor '$file'. Nic nebylo nahrazeno."
    }
}

$updateLockStream = $null
$stagingDirectory = $null
$backupDirectory = $null
$temporaryConfigPath = $null
$configBackupPath = $null
$automationLock = Join-Path $dataDirectory 'automation.lock'
try {
    try {
        $updateLockStream = [IO.File]::Open($automationLock, [IO.FileMode]::CreateNew,
            [IO.FileAccess]::Write, [IO.FileShare]::None)
    }
    catch {
        throw 'Automatizace běží nebo zůstal její bezpečnostní zámek. Nejprve stav ověřte podle provozního manuálu.'
    }

    # Fail closed before replacing code. All existing settings remain unchanged except for safe mode.
    if ($config.PSObject.Properties['mode']) { $config.mode = 'VALIDATE' }
    else { $config | Add-Member -NotePropertyName mode -NotePropertyValue 'VALIDATE' }
    $temporaryConfigPath = Join-Path $InstallDirectory ".config-update-$([guid]::NewGuid().ToString('N')).tmp"
    $configBackupPath = Join-Path $InstallDirectory ".config-update-backup-$([guid]::NewGuid().ToString('N')).json"
    $config | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath $temporaryConfigPath -Encoding UTF8
    [IO.File]::Replace($temporaryConfigPath, $configPath, $configBackupPath)
    $temporaryConfigPath = $null
    Remove-Item -LiteralPath $configBackupPath -Force -ErrorAction SilentlyContinue

    # Starting a real update invalidates any previous code-bound approval, even if replacement later fails.
    Remove-Item -LiteralPath (Join-Path $dataDirectory 'test-success.json') -Force -ErrorAction SilentlyContinue
    $previewDirectory = Join-Path $dataDirectory 'previews'
    if (Test-Path -LiteralPath $previewDirectory -PathType Container) {
        Get-ChildItem -LiteralPath $previewDirectory -Filter 'preview-*.html' -File |
            Remove-Item -Force -ErrorAction SilentlyContinue
    }

    $stagingDirectory = Join-Path ([IO.Path]::GetTempPath()) "siola-update-$([guid]::NewGuid().ToString('N'))"
    New-Item -ItemType Directory -Path $stagingDirectory | Out-Null
    foreach ($file in $runtimeFiles) {
        Copy-Item -LiteralPath (Join-Path $PSScriptRoot $file) -Destination (Join-Path $stagingDirectory $file)
        Unblock-File -LiteralPath (Join-Path $stagingDirectory $file)
    }

    # Test the staged version, not the currently installed copy.
    & (Join-Path $stagingDirectory 'Test-SiolaCore.ps1')

    $stagedHashes = @{}
    $previouslyInstalled = @{}
    $backupDirectory = Join-Path ([IO.Path]::GetTempPath()) "siola-backup-$([guid]::NewGuid().ToString('N'))"
    New-Item -ItemType Directory -Path $backupDirectory | Out-Null
    foreach ($file in $runtimeFiles) {
        $stagedHashes[$file] = (Get-FileHash -LiteralPath (Join-Path $stagingDirectory $file) -Algorithm SHA256).Hash
        $installedPath = Join-Path $InstallDirectory $file
        $previouslyInstalled[$file] = Test-Path -LiteralPath $installedPath -PathType Leaf
        if ($previouslyInstalled[$file]) {
            Copy-Item -LiteralPath $installedPath -Destination (Join-Path $backupDirectory $file)
        }
    }

    try {
        foreach ($file in $runtimeFiles) {
            $installedPath = Join-Path $InstallDirectory $file
            Move-Item -LiteralPath (Join-Path $stagingDirectory $file) -Destination $installedPath -Force
            $installedHash = (Get-FileHash -LiteralPath $installedPath -Algorithm SHA256).Hash
            if ($installedHash -cne $stagedHashes[$file]) {
                throw "Kontrolní součet nainstalovaného souboru '$file' nesouhlasí."
            }
        }
    }
    catch {
        $updateFailure = $_.Exception.Message
        $rollbackFailures = [Collections.Generic.List[string]]::new()
        foreach ($file in $runtimeFiles) {
            $installedPath = Join-Path $InstallDirectory $file
            try {
                if ($previouslyInstalled[$file]) {
                    Copy-Item -LiteralPath (Join-Path $backupDirectory $file) -Destination $installedPath -Force
                }
                elseif (Test-Path -LiteralPath $installedPath -PathType Leaf) {
                    Remove-Item -LiteralPath $installedPath -Force
                }
            }
            catch { $rollbackFailures.Add("${file}: $($_.Exception.Message)") }
        }
        if ($rollbackFailures.Count -gt 0) {
            throw "Aktualizace selhala ($updateFailure) a návrat některých souborů také selhal: $($rollbackFailures -join '; '). Úloha zůstává vypnutá."
        }
        throw "Aktualizace selhala a původní programové soubory byly obnoveny: $updateFailure"
    }

    Write-Host ''
    Write-Host "Program byl bezpečně aktualizován v: $InstallDirectory"
    Write-Host 'Konfigurace, Google klíč a čas naplánované úlohy zůstaly zachované.'
    Write-Host "Naplánovaná úloha '$TaskName' je VYPNUTÁ a režim je VALIDATE."
    Write-Host 'Nyní spusťte VALIDATE.cmd, TEST.cmd a po kontrole ENABLE_LIVE.cmd.'
}
finally {
    if ($temporaryConfigPath -and (Test-Path -LiteralPath $temporaryConfigPath -PathType Leaf)) {
        Remove-Item -LiteralPath $temporaryConfigPath -Force -ErrorAction SilentlyContinue
    }
    if ($configBackupPath -and (Test-Path -LiteralPath $configBackupPath -PathType Leaf)) {
        Remove-Item -LiteralPath $configBackupPath -Force -ErrorAction SilentlyContinue
    }
    if ($stagingDirectory -and (Test-Path -LiteralPath $stagingDirectory -PathType Container)) {
        Remove-Item -LiteralPath $stagingDirectory -Recurse -Force -ErrorAction SilentlyContinue
    }
    if ($backupDirectory -and (Test-Path -LiteralPath $backupDirectory -PathType Container)) {
        Remove-Item -LiteralPath $backupDirectory -Recurse -Force -ErrorAction SilentlyContinue
    }
    if ($null -ne $updateLockStream) {
        $lockName = $updateLockStream.Name
        $updateLockStream.Dispose()
        Remove-Item -LiteralPath $lockName -Force -ErrorAction SilentlyContinue
    }
}
