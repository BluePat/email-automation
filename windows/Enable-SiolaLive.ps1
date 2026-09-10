[CmdletBinding()]
param(
    [string]$TaskName = 'SIOLA Email Automation',
    [switch]$TakeOver
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSVersion.Major -lt 7 -or -not $IsWindows) {
    throw 'Tento krok spusťte ve Windows v PowerShellu 7.'
}

$configPath = Join-Path $PSScriptRoot 'config.json'
if (-not (Test-Path -LiteralPath $configPath)) { throw 'Chybí config.json. Nejdříve spusťte instalaci.' }
$config = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
$dataDirectory = [Environment]::ExpandEnvironmentVariables([string]$config.dataDirectory)
if (-not $dataDirectory) { throw 'V konfiguraci chybí dataDirectory.' }
Import-Module (Join-Path $PSScriptRoot 'Siola.Core.psm1') -Force

Write-Host 'Probíhá poslední kontrola tabulky. Nic se nebude měnit ani odesílat.'
$digestPath = Join-Path $dataDirectory "activation-digest-$([guid]::NewGuid().ToString('N')).txt"
try {
    & (Join-Path $PSScriptRoot 'Invoke-SiolaAutomation.ps1') -Mode VALIDATE -ConfigPath $configPath `
        -ApprovalDigestPath $digestPath
    if (-not (Test-Path -LiteralPath $digestPath -PathType Leaf)) { throw 'Kontrola nevytvořila otisk obsahu.' }
    $currentApprovalDigest = (Get-Content -LiteralPath $digestPath -Raw -Encoding UTF8).Trim()
}
finally { Remove-Item -LiteralPath $digestPath -Force -ErrorAction SilentlyContinue }

$receiptPath = Join-Path $dataDirectory 'test-success.json'
if (-not (Test-Path -LiteralPath $receiptPath -PathType Leaf)) {
    throw 'Chybí potvrzení úspěšného TESTU. Nejdříve spusťte TEST.cmd.'
}
$receipt = Get-Content -LiteralPath $receiptPath -Raw -Encoding UTF8 | ConvertFrom-Json
$completed = [DateTimeOffset]::Parse([string]$receipt.completedUtc, [Globalization.CultureInfo]::InvariantCulture)
if ($completed -lt [DateTimeOffset]::UtcNow.AddHours(-24)) {
    throw 'Poslední úspěšný TEST je starší než 24 hodin. Spusťte TEST.cmd znovu.'
}
foreach ($pair in @(
    [pscustomobject]@{ Receipt = 'spreadsheetId'; Config = 'spreadsheetId' },
    [pscustomobject]@{ Receipt = 'worksheetName'; Config = 'worksheetName' },
    [pscustomobject]@{ Receipt = 'sender'; Config = 'outlookSenderSmtpAddress' },
    [pscustomobject]@{ Receipt = 'testRecipient'; Config = 'testRecipient' },
    [pscustomobject]@{ Receipt = 'installationId'; Config = 'installationId' },
    [pscustomobject]@{ Receipt = 'machineBinding'; Config = 'machineBinding' }
)) {
    if ([string]$receipt.($pair.Receipt) -cne [string]$config.($pair.Config)) {
        throw 'Konfigurace se od posledního TESTU změnila. Spusťte TEST.cmd znovu.'
    }
}
if (-not (Test-Path -LiteralPath ([string]$receipt.previewPath) -PathType Leaf)) {
    throw 'Chybí soubor s náhledem všech zpráv. Spusťte TEST.cmd znovu.'
}
if ([int]$receipt.sentTestJobs -lt 1) {
    throw 'Poslední TEST neposlal žádnou testovací zprávu. Spusťte TEST.cmd znovu.'
}
if (-not $receipt.PSObject.Properties['approvalDigest'] -or
    [string]$receipt.approvalDigest -cne $currentApprovalDigest) {
    throw 'Obsah připravených e-mailů se od posledního TESTU změnil. Spusťte TEST.cmd znovu.'
}
$currentRuntimeFingerprint = Get-SiolaRuntimeFingerprint -RootPath $PSScriptRoot
if (-not $receipt.PSObject.Properties['runtimeFingerprint'] -or
    [string]$receipt.runtimeFingerprint -cne $currentRuntimeFingerprint) {
    throw 'Provozní soubory se od posledního TESTU změnily. Spusťte TEST.cmd znovu.'
}

$expectedConfirmation = if ($TakeOver) { 'PREVZIT' } else { 'LIVE' }
$prompt = if ($TakeOver) {
    'Nejdříve vypněte starou instalaci. Pro převzetí tabulky touto instalací napište přesně PREVZIT'
}
else { 'Potvrďte kontrolu TEST e-mailů i celého HTML náhledu. Pro zapnutí napište přesně LIVE' }
$confirmation = Read-Host $prompt
if ($confirmation -cne $expectedConfirmation) { throw 'Ostrý režim nebyl zapnut.' }

Import-Module (Join-Path $PSScriptRoot 'Siola.GoogleSheets.psm1') -Force
$credentialsPath = [Environment]::ExpandEnvironmentVariables([string]$config.credentialsPath)
$accessToken = Get-GoogleServiceAccountToken -CredentialsPath $credentialsPath
if ($TakeOver) {
    Transfer-SiolaAutomationOwner -SpreadsheetId ([string]$config.spreadsheetId) -AccessToken $accessToken `
        -InstallationId ([string]$config.installationId)
}
else {
    Register-SiolaAutomationOwner -SpreadsheetId ([string]$config.spreadsheetId) -AccessToken $accessToken `
        -InstallationId ([string]$config.installationId)
}
$config.mode = 'LIVE'
$config | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $configPath -Encoding UTF8
Enable-ScheduledTask -TaskName $TaskName | Out-Null
foreach ($sensitivePath in @([string]$receipt.previewPath, $receiptPath)) {
    if (-not $sensitivePath) { continue }
    try { Remove-Item -LiteralPath $sensitivePath -Force -ErrorAction Stop }
    catch { Write-Warning "Citlivý soubor po zapnutí LIVE nešlo odstranit: $sensitivePath" }
}
$previewDirectory = Join-Path ([Environment]::ExpandEnvironmentVariables([string]$config.dataDirectory)) 'previews'
if (Test-Path -LiteralPath $previewDirectory -PathType Container) {
    Get-ChildItem -LiteralPath $previewDirectory -Filter 'preview-*.html' -File -ErrorAction SilentlyContinue |
        ForEach-Object {
            try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop }
            catch { Write-Warning "Starší citlivý náhled nešlo odstranit: $($_.FullName)" }
        }
}
Write-Host "Naplánovaná úloha '$TaskName' je zapnutá v režimu LIVE."
