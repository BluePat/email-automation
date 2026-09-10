[CmdletBinding()]
param(
    [string]$InstallDirectory = (Join-Path $env:LOCALAPPDATA 'SIOLA Email Automation'),
    [string]$TaskName = 'SIOLA Email Automation',
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Read-RequiredText {
    param([Parameter(Mandatory)][string]$Prompt)
    $value = (Read-Host $Prompt).Trim()
    if (-not $value) { throw "Hodnota '$Prompt' nesmí být prázdná." }
    return $value
}

function Assert-EmailAddress {
    param([Parameter(Mandatory)][string]$Value, [Parameter(Mandatory)][string]$Label)
    try {
        $parsed = [Net.Mail.MailAddress]::new($Value)
        if ($parsed.Address -ine $Value) { throw 'Neplatná adresa' }
    }
    catch { throw "$Label nemá platný e-mailový formát." }
}

if ($PSVersionTable.PSVersion.Major -lt 7 -or -not $IsWindows) {
    throw 'Instalaci spusťte ve Windows v PowerShellu 7.'
}

$existingConfig = Join-Path $InstallDirectory 'config.json'
$previousConfig = $null
if (Test-Path -LiteralPath $existingConfig -PathType Leaf) {
    try { $previousConfig = Get-Content -LiteralPath $existingConfig -Raw -Encoding UTF8 | ConvertFrom-Json }
    catch { throw 'Existující config.json nelze načíst. Opravte jej nebo zazálohujte instalaci.' }
}
if ((Test-Path -LiteralPath $existingConfig) -and -not $Force) {
    throw "Instalace už existuje v $InstallDirectory. Bez parametru -Force nebude přepsána."
}

$installLockStream = $null
$stagingDirectory = $null
if ($Force -and $null -ne $previousConfig) {
    $existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($null -ne $existingTask) {
        Disable-ScheduledTask -TaskName $TaskName | Out-Null
        $existingTask = Get-ScheduledTask -TaskName $TaskName
        if ([string]$existingTask.State -ceq 'Running') {
            throw "Úloha '$TaskName' právě běží. Byla vypnuta; instalaci opakujte až po jejím skončení."
        }
    }
    $existingDataDirectory = if ($previousConfig.PSObject.Properties['dataDirectory']) {
        [Environment]::ExpandEnvironmentVariables([string]$previousConfig.dataDirectory)
    } else { Join-Path $InstallDirectory 'data' }
    New-Item -ItemType Directory -Path $existingDataDirectory -Force | Out-Null
    $existingAutomationLock = Join-Path $existingDataDirectory 'automation.lock'
    try {
        $installLockStream = [IO.File]::Open($existingAutomationLock, [IO.FileMode]::CreateNew,
            [IO.FileAccess]::Write, [IO.FileShare]::None)
    }
    catch {
        throw 'Automatizace běží nebo zůstal její bezpečnostní zámek. Nejprve stav ověřte a použijte ForceUnlock.'
    }
}

try {

$credentialSource = (Read-Host 'Úplná cesta ke Google service-account JSON klíči').Trim().Trim('"')
if (-not (Test-Path -LiteralPath $credentialSource -PathType Leaf)) { throw 'Zadaný JSON soubor neexistuje.' }
$credentialCheck = Get-Content -LiteralPath $credentialSource -Raw -Encoding UTF8 | ConvertFrom-Json
if ($credentialCheck.type -ne 'service_account' -or -not $credentialCheck.client_email) {
    throw 'Soubor není platný service-account JSON klíč.'
}

$spreadsheetInput = Read-RequiredText 'Google Sheet URL nebo ID'
$spreadsheetId = if ($spreadsheetInput -match '/spreadsheets/d/([^/]+)') { $Matches[1] } else { $spreadsheetInput.Trim() }
if ($spreadsheetId -notmatch '^[A-Za-z0-9_-]+$') { throw 'Google Sheet ID nemá platný formát.' }

$sender = Read-RequiredText 'Odesílající účet v Classic Outlook'
Assert-EmailAddress -Value $sender -Label 'Odesílající účet'
$testRecipient = Read-RequiredText 'E-mail, na který se mají poslat TEST zprávy'
Assert-EmailAddress -Value $testRecipient -Label 'Testovací e-mail'
$signatureName = Read-RequiredText 'Jméno do podpisu e-mailu'
$signaturePhone = Read-RequiredText 'Telefon do podpisu e-mailu'
$signatureEmail = (Read-Host "E-mail do osobního podpisu [$sender]").Trim()
if (-not $signatureEmail) { $signatureEmail = $sender }
Assert-EmailAddress -Value $signatureEmail -Label 'E-mail v osobním podpisu'
$signatureCompany = Read-RequiredText 'Název společnosti do podpisu'
$signatureAddress = Read-RequiredText 'Adresa společnosti do podpisu'
$signatureCompanyEmail = Read-RequiredText 'Obecný e-mail společnosti do podpisu'
Assert-EmailAddress -Value $signatureCompanyEmail -Label 'E-mail společnosti'
$signatureCompanyId = Read-RequiredText 'IČ společnosti do podpisu'
$dailyTimeText = Read-Host 'Denní čas spuštění ve formátu HH:mm [08:00]'
if (-not $dailyTimeText) { $dailyTimeText = '08:00' }
[datetime]$dailyTime = [datetime]::MinValue
if (-not [datetime]::TryParseExact($dailyTimeText, 'HH:mm', [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.DateTimeStyles]::None, [ref]$dailyTime)) {
    throw 'Čas musí mít formát HH:mm.'
}

$dataDirectory = Join-Path $InstallDirectory 'data'
$secretDirectory = Join-Path $dataDirectory 'secrets'
New-Item -ItemType Directory -Path $InstallDirectory, $dataDirectory, $secretDirectory -Force | Out-Null
$identity = [Security.Principal.WindowsIdentity]::GetCurrent().Name
& icacls.exe $InstallDirectory '/inheritance:r' '/grant:r' `
    "${identity}:(OI)(CI)F" '*S-1-5-18:(OI)(CI)F' '/T' '/C' | Out-Null
if ($LASTEXITCODE -ne 0) { throw 'Nepodařilo se zabezpečit instalační a datový adresář.' }

$runtimeFiles = @(
    'Invoke-SiolaAutomation.ps1', 'Siola.Core.psm1', 'Siola.GoogleSheets.psm1', 'Siola.Outlook.psm1',
    'Run-Validation.ps1', 'Run-Test.ps1', 'Enable-SiolaLive.ps1', 'Test-SiolaCore.ps1',
    'VALIDATE.cmd', 'TEST.cmd', 'ENABLE_LIVE.cmd', 'TAKE_OVER.cmd'
)
$stagingDirectory = Join-Path ([IO.Path]::GetTempPath()) "siola-install-$([guid]::NewGuid().ToString('N'))"
New-Item -ItemType Directory -Path $stagingDirectory | Out-Null
foreach ($file in $runtimeFiles) {
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot $file) -Destination (Join-Path $stagingDirectory $file)
    Unblock-File -LiteralPath (Join-Path $stagingDirectory $file)
}
& (Join-Path $stagingDirectory 'Test-SiolaCore.ps1')
foreach ($file in $runtimeFiles) {
    Move-Item -LiteralPath (Join-Path $stagingDirectory $file) -Destination (Join-Path $InstallDirectory $file) -Force
}

$credentialDestination = Join-Path $secretDirectory 'google-service-account.json'
$sourceFullPath = [IO.Path]::GetFullPath($credentialSource)
$destinationFullPath = [IO.Path]::GetFullPath($credentialDestination)
if ($sourceFullPath -ine $destinationFullPath) {
    Copy-Item -LiteralPath $credentialSource -Destination $credentialDestination -Force
}
if ((Get-FileHash -LiteralPath $credentialSource -Algorithm SHA256).Hash -cne
    (Get-FileHash -LiteralPath $credentialDestination -Algorithm SHA256).Hash) {
    throw 'Kopie Google klíče neodpovídá originálu.'
}

$config = [ordered]@{
    mode = 'VALIDATE'
    spreadsheetId = $spreadsheetId
    worksheetName = 'Obce a města'
    credentialsPath = $credentialDestination
    dataDirectory = $dataDirectory
    outlookSenderSmtpAddress = $sender
    testRecipient = $testRecipient
    signature = [ordered]@{
        name = $signatureName
        phone = $signaturePhone
        email = $signatureEmail
        company = $signatureCompany
        address = $signatureAddress
        companyEmail = $signatureCompanyEmail
        companyId = $signatureCompanyId
    }
    batchSize = 50
    testBatchSize = 3
    delaySeconds = 3
    sendConfirmationTimeoutSeconds = 120
    runDeadlineMinutes = 300
    logRetentionDays = 90
    previewRetentionDays = 2
    installationId = $(if ($null -ne $previousConfig -and $previousConfig.PSObject.Properties['installationId']) {
        [string]$previousConfig.installationId
    } else { [guid]::NewGuid().ToString('D') })
    machineBinding = "$env:COMPUTERNAME|$identity"
}
$config | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $existingConfig -Encoding UTF8
Remove-Item -LiteralPath (Join-Path $dataDirectory 'test-success.json') -Force -ErrorAction SilentlyContinue

$pwsh = (Get-Process -Id $PID).Path
$actionArguments = "-NoLogo -NoProfile -NonInteractive -File `"$(Join-Path $InstallDirectory 'Invoke-SiolaAutomation.ps1')`""
$action = New-ScheduledTaskAction -Execute $pwsh -Argument $actionArguments
$trigger = New-ScheduledTaskTrigger -Daily -At $dailyTime
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Hours 6)
$principal = New-ScheduledTaskPrincipal -UserId $identity -LogonType Interactive -RunLevel Limited
Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings `
    -Principal $principal -Description 'SIOLA: Google Sheets -> Classic Outlook; pouze při přihlášeném uživateli.' `
    -Force | Out-Null
Disable-ScheduledTask -TaskName $TaskName | Out-Null

if ($sourceFullPath -ine $destinationFullPath) {
    $deleteOriginal = Read-Host 'Instalace je hotová. Pro TRVALÉ smazání původního Google klíče napište přesně SMAZAT; jinak zůstane zachován'
    if ($deleteOriginal -ceq 'SMAZAT') {
        try {
            [IO.File]::Delete($sourceFullPath)
            if (Test-Path -LiteralPath $sourceFullPath) { throw 'Soubor stále existuje.' }
            Write-Host 'Původní Google klíč byl trvale smazán.'
        }
        catch { Write-Warning "Instalace je hotová, ale původní klíč se nepodařilo smazat: $sourceFullPath" }
    }
    else { Write-Warning "Původní citlivý klíč zůstal v: $sourceFullPath" }
}

Write-Host ''
Write-Host "Instalace je připravená v: $InstallDirectory"
Write-Host "Google tabulku sdílejte jako Editor s účtem: $($credentialCheck.client_email)"
Write-Host "Naplánovaná úloha '$TaskName' je z bezpečnostních důvodů VYPNUTÁ."
Write-Host 'Po nasdílení spusťte VALIDATE.cmd a poté TEST.cmd.'
Write-Host 'Až po ruční kontrole testů spusťte ENABLE_LIVE.cmd.'
}
finally {
    if ($stagingDirectory -and (Test-Path -LiteralPath $stagingDirectory -PathType Container)) {
        Remove-Item -LiteralPath $stagingDirectory -Recurse -Force -ErrorAction SilentlyContinue
    }
    if ($null -ne $installLockStream) {
        $lockName = $installLockStream.Name
        $installLockStream.Dispose()
        Remove-Item -LiteralPath $lockName -Force -ErrorAction SilentlyContinue
    }
}
