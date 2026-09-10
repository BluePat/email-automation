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

$credentialSource = Read-Host 'Úplná cesta ke Google service-account JSON klíči'
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

$runtimeFiles = @(
    'Invoke-SiolaAutomation.ps1', 'Siola.Core.psm1', 'Siola.GoogleSheets.psm1', 'Siola.Outlook.psm1',
    'Run-Validation.ps1', 'Run-Test.ps1', 'Enable-SiolaLive.ps1', 'Test-SiolaCore.ps1',
    'VALIDATE.cmd', 'TEST.cmd', 'ENABLE_LIVE.cmd', 'TAKE_OVER.cmd'
)
foreach ($file in $runtimeFiles) {
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot $file) -Destination (Join-Path $InstallDirectory $file) -Force
    Unblock-File -LiteralPath (Join-Path $InstallDirectory $file)
}
& (Join-Path $InstallDirectory 'Test-SiolaCore.ps1')

$credentialDestination = Join-Path $secretDirectory 'google-service-account.json'
Copy-Item -LiteralPath $credentialSource -Destination $credentialDestination -Force
if ((Get-FileHash -LiteralPath $credentialSource -Algorithm SHA256).Hash -cne
    (Get-FileHash -LiteralPath $credentialDestination -Algorithm SHA256).Hash) {
    throw 'Kopie Google klíče neodpovídá originálu.'
}
$identity = [Security.Principal.WindowsIdentity]::GetCurrent().Name
& icacls.exe $secretDirectory '/inheritance:r' '/grant:r' "${identity}:(OI)(CI)F" | Out-Null
if ($LASTEXITCODE -ne 0) { throw 'Nepodařilo se zabezpečit adresář s Google klíčem.' }

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
    logRetentionDays = 90
    installationId = $(if ($null -ne $previousConfig -and $previousConfig.PSObject.Properties['installationId']) {
        [string]$previousConfig.installationId
    } else { [guid]::NewGuid().ToString('D') })
    machineBinding = "$env:COMPUTERNAME|$identity"
}
$config | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $existingConfig -Encoding UTF8
Remove-Item -LiteralPath (Join-Path $dataDirectory 'test-success.json') -Force -ErrorAction SilentlyContinue

$sourceFullPath = [IO.Path]::GetFullPath($credentialSource)
$destinationFullPath = [IO.Path]::GetFullPath($credentialDestination)
if ($sourceFullPath -ine $destinationFullPath) {
    $keepOriginal = Read-Host 'Zabezpečená kopie klíče je hotová. Pro ponechání původního souboru napište PONECHAT; Enter jej přesune do Koše'
    if ($keepOriginal -cne 'PONECHAT') {
        try {
            Add-Type -AssemblyName Microsoft.VisualBasic
            [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(
                $sourceFullPath,
                [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs,
                [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin
            )
        }
        catch { throw "Klíč byl bezpečně zkopírován, ale originál se nepodařilo přesunout do Koše: $sourceFullPath" }
    }
    else { Write-Warning "Původní citlivý klíč zůstal v: $sourceFullPath" }
}

$pwsh = (Get-Process -Id $PID).Path
$actionArguments = "-NoLogo -NoProfile -NonInteractive -File `"$(Join-Path $InstallDirectory 'Invoke-SiolaAutomation.ps1')`""
$action = New-ScheduledTaskAction -Execute $pwsh -Argument $actionArguments
$trigger = New-ScheduledTaskTrigger -Daily -At $dailyTime
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Hours 2)
$principal = New-ScheduledTaskPrincipal -UserId $identity -LogonType Interactive -RunLevel Limited
Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings `
    -Principal $principal -Description 'SIOLA: Google Sheets -> Classic Outlook; pouze při přihlášeném uživateli.' `
    -Force | Out-Null
Disable-ScheduledTask -TaskName $TaskName | Out-Null

Write-Host ''
Write-Host "Instalace je připravená v: $InstallDirectory"
Write-Host "Google tabulku sdílejte jako Editor s účtem: $($credentialCheck.client_email)"
Write-Host "Naplánovaná úloha '$TaskName' je z bezpečnostních důvodů VYPNUTÁ."
Write-Host 'Po nasdílení spusťte VALIDATE.cmd a poté TEST.cmd.'
Write-Host 'Až po ruční kontrole testů spusťte ENABLE_LIVE.cmd.'
