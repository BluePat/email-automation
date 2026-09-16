[CmdletBinding()]
param([string]$ConfigPath = (Join-Path $PSScriptRoot 'config.json'))

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSVersion.Major -lt 7 -or -not $IsWindows) {
    throw 'Diagnostiku spusťte ve Windows v PowerShellu 7.'
}
if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
    throw "Chybí konfigurace: $ConfigPath"
}

$config = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
$expected = ([string]$config.outlookSenderSmtpAddress).Trim()
if (-not $expected) { throw 'V konfiguraci chybí outlookSenderSmtpAddress.' }
$dataDirectory = [Environment]::ExpandEnvironmentVariables([string]$config.dataDirectory)
if (-not $dataDirectory) { $dataDirectory = $PSScriptRoot }
New-Item -ItemType Directory -Path $dataDirectory -Force | Out-Null
$reportPath = Join-Path $dataDirectory "outlook-diagnostic-$(Get-Date -Format 'yyyyMMddTHHmmss').txt"
$outlookModulePath = Join-Path $PSScriptRoot 'Siola.Outlook.psm1'
Import-Module $outlookModulePath -Force

$report = [Collections.Generic.List[string]]::new()
function Add-ReportLine {
    param([AllowEmptyString()][string]$Text = '')
    $report.Add($Text)
    Write-Host $Text
}

function Format-ExactText {
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return '<NULL>' }
    $text = [string]$Value
    $codes = @($text.ToCharArray() | ForEach-Object { 'U+{0:X4}' -f [int]$_ }) -join ' '
    return "[$text] length=$($text.Length) chars=$codes"
}

function Release-DiagnosticComObject {
    param([AllowNull()]$Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [Runtime.InteropServices.Marshal]::ReleaseComObject($Value) | Out-Null
    }
}

function Get-DiagnosticAddressEntrySmtp {
    param([AllowNull()]$AddressEntry)
    if ($null -eq $AddressEntry) { return '' }
    $accessor = $null
    $exchangeUser = $null
    try {
        $smtp = ''
        try {
            $accessor = $AddressEntry.PropertyAccessor
            $smtp = [string]$accessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x39FE001E')
        }
        catch {}
        if (-not $smtp -and [string]$AddressEntry.Type -ieq 'EX') {
            try {
                $exchangeUser = $AddressEntry.GetExchangeUser()
                if ($null -ne $exchangeUser) { $smtp = [string]$exchangeUser.PrimarySmtpAddress }
            }
            catch {}
        }
        if (-not $smtp -and [string]$AddressEntry.Type -ieq 'SMTP') {
            $smtp = [string]$AddressEntry.Address
        }
        return $smtp.Trim()
    }
    finally {
        Release-DiagnosticComObject $exchangeUser
        Release-DiagnosticComObject $accessor
    }
}

function Get-DiagnosticMailIdentity {
    param([Parameter(Mandatory)]$Mail, [Parameter(Mandatory)][string]$Stage)
    $account = $null
    $sender = $null
    $accountSmtp = ''
    $senderSmtp = ''
    $accountError = ''
    $senderError = ''
    try {
        try {
            $account = $Mail.SendUsingAccount
            if ($null -ne $account) { $accountSmtp = [string]$account.SmtpAddress }
        }
        catch { $accountError = $_.Exception.Message }
        try {
            $sender = $Mail.Sender
            $senderSmtp = Get-DiagnosticAddressEntrySmtp $sender
        }
        catch { $senderError = $_.Exception.Message }
        Add-ReportLine "--- $Stage ---"
        Add-ReportLine "SendUsingAccount: $(Format-ExactText $accountSmtp)"
        Add-ReportLine "Sender SMTP:      $(Format-ExactText $senderSmtp)"
        Add-ReportLine "Account equals expected: $([string]::Equals($accountSmtp.Trim(), $expected, [StringComparison]::OrdinalIgnoreCase))"
        Add-ReportLine "Sender equals expected:  $([string]::Equals($senderSmtp.Trim(), $expected, [StringComparison]::OrdinalIgnoreCase))"
        if ($accountError) { Add-ReportLine "Account read error: $accountError" }
        if ($senderError) { Add-ReportLine "Sender read error: $senderError" }
    }
    finally {
        Release-DiagnosticComObject $sender
        Release-DiagnosticComObject $account
    }
}

$outlook = $null
$mail = $null
$currentUser = $null
$senderEntry = $null
$inspector = $null
try {
    Add-ReportLine 'SIOLA OUTLOOK DIAGNOSTIKA - NIC SE NEODESÍLÁ'
    Add-ReportLine "Čas: $([DateTimeOffset]::Now.ToString('o'))"
    Add-ReportLine "PowerShell: $($PSVersionTable.PSVersion)"
    Add-ReportLine "Outlook modul: $outlookModulePath"
    Add-ReportLine "SHA256 modulu: $((Get-FileHash -LiteralPath $outlookModulePath -Algorithm SHA256).Hash)"
    Add-ReportLine "Očekávaný účet: $(Format-ExactText $expected)"

    $outlook = Connect-SiolaOutlook -SenderSmtpAddress $expected
    Add-ReportLine "Počet Outlook účtů: $($outlook.Session.Accounts.Count)"
    for ($index = 1; $index -le $outlook.Session.Accounts.Count; $index++) {
        $listedAccount = $null
        try {
            $listedAccount = $outlook.Session.Accounts.Item($index)
            $listedSmtp = [string]$listedAccount.SmtpAddress
            Add-ReportLine "Účet $index: SMTP=$(Format-ExactText $listedSmtp); typ=$($listedAccount.AccountType); shoda=$([string]::Equals($listedSmtp.Trim(), $expected, [StringComparison]::OrdinalIgnoreCase))"
        }
        finally { Release-DiagnosticComObject $listedAccount }
    }

    $mail = $outlook.Application.CreateItem(0)
    $mail.Subject = '[SIOLA DIAGNOSTIKA – NEODESÍLAT] kontrola účtu'
    $mail.BodyFormat = 2
    $mail.HTMLBody = '<p><strong>DIAGNOSTIKA – NEODESÍLAT</strong></p><p>Zpráva nemá příjemce a bude zahozena.</p>'
    Get-DiagnosticMailIdentity -Mail $mail -Stage 'nová zpráva před nastavením'

    $currentUser = $outlook.Account.CurrentUser
    $senderEntry = $currentUser.AddressEntry
    Add-ReportLine "CurrentUser AddressEntry: $(Format-ExactText (Get-DiagnosticAddressEntrySmtp $senderEntry))"
    try { $mail.Sender = $senderEntry }
    catch { Add-ReportLine "Běžný Sender setter error: $($_.Exception.Message)" }
    try { $mail.SendUsingAccount = $outlook.Account }
    catch { Add-ReportLine "Běžný SendUsingAccount setter error: $($_.Exception.Message)" }
    Get-DiagnosticMailIdentity -Mail $mail -Stage 'po běžném PowerShell nastavení'

    try {
        $inspector = $mail.GetInspector
        Get-DiagnosticMailIdentity -Mail $mail -Stage 'po vytvoření Inspectoru'
    }
    catch { Add-ReportLine "Inspector error: $($_.Exception.Message)" }

    try {
        $mail.GetType().InvokeMember('SendUsingAccount', [Reflection.BindingFlags]::SetProperty, `
            $null, $mail, @($outlook.Account)) | Out-Null
        Get-DiagnosticMailIdentity -Mail $mail -Stage 'po reflection SetProperty'
    }
    catch { Add-ReportLine "Reflection setter error: $($_.Exception.Message)" }

    $mail.Display($false)
    Start-Sleep -Milliseconds 500
    try { $mail.Sender = $senderEntry }
    catch { Add-ReportLine "Sender setter po Display error: $($_.Exception.Message)" }
    try { $mail.SendUsingAccount = $outlook.Account }
    catch { Add-ReportLine "SendUsingAccount setter po Display error: $($_.Exception.Message)" }
    $mail.HTMLBody = '<p><strong>DIAGNOSTIKA – NEODESÍLAT</strong></p><p>Zpráva nemá příjemce a bude zahozena.</p>'
    Get-DiagnosticMailIdentity -Mail $mail -Stage 'po zobrazení konceptu a opakovaném nastavení'
    if ($mail.Recipients.Count -ne 0) {
        throw 'Bezpečnostní kontrola selhala: diagnostický koncept nesmí mít příjemce.'
    }
    Add-ReportLine 'Počet příjemců diagnostického konceptu: 0'
    Add-ReportLine ''
    Add-ReportLine 'V otevřeném konceptu zkontrolujte pole Od. Koncept nemá příjemce a nelze jej odeslat.'
    $null = Read-Host 'Po kontrole pole Od stiskněte Enter; koncept bude zahozen'
    Add-ReportLine 'Uživatel dokončil vizuální kontrolu pole Od.'
}
catch {
    Add-ReportLine "CHYBA: $($_.Exception.Message)"
    throw
}
finally {
    if ($null -ne $mail) {
        try { $mail.Close(1) } catch {}
    }
    Release-DiagnosticComObject $inspector
    Release-DiagnosticComObject $senderEntry
    Release-DiagnosticComObject $currentUser
    Release-DiagnosticComObject $mail
    Disconnect-SiolaOutlook $outlook
    $report | Set-Content -LiteralPath $reportPath -Encoding UTF8
    Write-Host "`nDiagnostický protokol: $reportPath"
}
