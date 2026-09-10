Set-StrictMode -Version Latest

function Release-SiolaComObject {
    param([AllowNull()]$Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) | Out-Null
    }
}

function Release-SiolaComReference {
    param([AllowNull()]$Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [Runtime.InteropServices.Marshal]::ReleaseComObject($Value) | Out-Null
    }
}

function Connect-SiolaOutlook {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SenderSmtpAddress)

    if (-not $IsWindows) { throw 'Classic Outlook lze automatizovat pouze ve Windows.' }
    $SenderSmtpAddress = $SenderSmtpAddress.Trim()
    try { $outlook = New-Object -ComObject Outlook.Application }
    catch { throw "Classic Outlook nelze spustit. Ověřte, že je nainstalovaný a nakonfigurovaný. $($_.Exception.Message)" }

    $session = $outlook.GetNamespace('MAPI')
    if ([bool]$session.Offline) {
        Release-SiolaComObject $session
        Release-SiolaComObject $outlook
        throw 'Classic Outlook je v režimu Pracovat offline nebo není připojený.'
    }

    $account = $null
    for ($index = 1; $index -le $session.Accounts.Count; $index++) {
        $candidate = $session.Accounts.Item($index)
        if ([string]$candidate.SmtpAddress -ieq $SenderSmtpAddress) {
            $account = $candidate
            break
        }
        Release-SiolaComObject $candidate
    }
    if ($null -eq $account) {
        Release-SiolaComObject $session
        Release-SiolaComObject $outlook
        throw "V Classic Outlook nebyl nalezen odesílající účet $SenderSmtpAddress."
    }

    return [pscustomobject]@{
        Application = $outlook
        Session = $session
        Account = $account
        SenderSmtpAddress = $SenderSmtpAddress.Trim()
    }
}

function Get-SiolaRecipientSmtpAddress {
    param([Parameter(Mandatory)]$Recipient)
    $entry = $null
    $accessor = $null
    $exchangeUser = $null
    $exchangeList = $null
    try {
        $entry = $Recipient.AddressEntry
        $smtp = ''
        try {
            $accessor = $entry.PropertyAccessor
            $smtp = [string]$accessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x39FE001E')
        }
        catch {}
        if (-not $smtp -and [string]$entry.Type -ieq 'EX') {
            try {
                $exchangeUser = $entry.GetExchangeUser()
                if ($null -ne $exchangeUser) { $smtp = [string]$exchangeUser.PrimarySmtpAddress }
            }
            catch {}
            if (-not $smtp) {
                try {
                    $exchangeList = $entry.GetExchangeDistributionList()
                    if ($null -ne $exchangeList) { $smtp = [string]$exchangeList.PrimarySmtpAddress }
                }
                catch {}
            }
        }
        if (-not $smtp -and [string]$entry.Type -ieq 'SMTP') { $smtp = [string]$entry.Address }
        return $smtp.Trim()
    }
    finally {
        Release-SiolaComObject $exchangeList
        Release-SiolaComObject $exchangeUser
        Release-SiolaComObject $accessor
        Release-SiolaComObject $entry
    }
}

function Test-SiolaFolderForJob {
    param([Parameter(Mandatory)]$Folder, [Parameter(Mandatory)][string]$JobId)
    $items = $null
    $match = $null
    try {
        $items = $Folder.Items
        $escapedJobId = $JobId.Replace("'", "''")
        $match = $items.Find("[BillingInformation] = '$escapedJobId'")
        return $null -ne $match
    }
    finally {
        Release-SiolaComObject $match
        Release-SiolaComObject $items
    }
}

function Get-SiolaOutlookJobLocation {
    param([Parameter(Mandatory)]$OutlookContext, [Parameter(Mandatory)][string]$JobId)
    $stores = $null
    $store = $null
    $folder = $null
    try {
        $stores = $OutlookContext.Session.Stores
        for ($index = 1; $index -le $stores.Count; $index++) {
            $store = $stores.Item($index)
            foreach ($candidate in @(
                [pscustomobject]@{ Name = 'SENT'; FolderType = 5 },
                [pscustomobject]@{ Name = 'OUTBOX'; FolderType = 4 }
            )) {
                try {
                    $folder = $store.GetDefaultFolder($candidate.FolderType)
                    if (Test-SiolaFolderForJob -Folder $folder -JobId $JobId) { return $candidate.Name }
                }
                catch {}
                finally {
                    Release-SiolaComObject $folder
                    $folder = $null
                }
            }
            Release-SiolaComObject $store
            $store = $null
        }
        return 'UNKNOWN'
    }
    finally {
        Release-SiolaComObject $folder
        Release-SiolaComObject $store
        Release-SiolaComObject $stores
    }
}

function Wait-SiolaOutlookJobSent {
    param(
        [Parameter(Mandatory)]$OutlookContext,
        [Parameter(Mandatory)][string]$JobId,
        [int]$TimeoutSeconds = 120
    )
    $deadline = [DateTimeOffset]::UtcNow.AddSeconds($TimeoutSeconds)
    $lastLocation = 'UNKNOWN'
    do {
        if ([bool]$OutlookContext.Session.Offline) {
            throw 'Outlook se během odesílání odpojil nebo přešel do režimu Pracovat offline.'
        }
        $lastLocation = Get-SiolaOutlookJobLocation -OutlookContext $OutlookContext -JobId $JobId
        if ($lastLocation -eq 'SENT') { return }
        Start-Sleep -Seconds 2
    } while ([DateTimeOffset]::UtcNow -lt $deadline)

    if ($lastLocation -eq 'OUTBOX') {
        throw "Zpráva $JobId zůstala po $TimeoutSeconds sekundách v Poště k odeslání."
    }
    throw "U zprávy $JobId nebylo možné potvrdit uložení v Odeslané poště."
}

function Send-SiolaOutlookJob {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$OutlookContext,
        [Parameter(Mandatory)]$Job,
        [int]$ConfirmationTimeoutSeconds = 120
    )
    if ([bool]$OutlookContext.Session.Offline) { throw 'Classic Outlook je offline.' }
    $mail = $null
    $recipient = $null
    $assignedAccount = $null
    try {
        $mail = $OutlookContext.Application.CreateItem(0)
        $mail.SendUsingAccount = $OutlookContext.Account
        $assignedAccount = $mail.SendUsingAccount
        if ($null -eq $assignedAccount -or
            [string]$assignedAccount.SmtpAddress -ine [string]$OutlookContext.SenderSmtpAddress) {
            throw "Outlook nepotvrdil odesílající účet $($OutlookContext.SenderSmtpAddress)."
        }
        Release-SiolaComReference $assignedAccount
        $assignedAccount = $null
        $recipient = $mail.Recipients.Add([string]$Job.To)
        if (-not $recipient.Resolve()) { throw "Outlook nedokázal ověřit příjemce $($Job.To)." }
        $resolvedSmtp = Get-SiolaRecipientSmtpAddress $recipient
        if (-not $resolvedSmtp -or $resolvedSmtp -ine [string]$Job.To) {
            throw "Outlook přeložil příjemce '$($Job.To)' na jinou adresu '$resolvedSmtp'."
        }
        $mail.Subject = [string]$Job.Subject
        $mail.BodyFormat = 2
        $mail.HTMLBody = [string]$Job.BodyHtml
        $mail.BillingInformation = [string]$Job.JobId
        $mail.DeleteAfterSubmit = $false
        $mail.Send()
        Release-SiolaComObject $recipient
        $recipient = $null
        Release-SiolaComObject $mail
        $mail = $null
        Wait-SiolaOutlookJobSent -OutlookContext $OutlookContext -JobId ([string]$Job.JobId) `
            -TimeoutSeconds $ConfirmationTimeoutSeconds
    }
    finally {
        Release-SiolaComReference $assignedAccount
        Release-SiolaComObject $recipient
        Release-SiolaComObject $mail
    }
}

function Disconnect-SiolaOutlook {
    param([AllowNull()]$OutlookContext)
    if ($null -eq $OutlookContext) { return }
    foreach ($name in @('Account', 'Session', 'Application')) {
        Release-SiolaComObject $OutlookContext.$name
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Export-ModuleMember -Function Connect-SiolaOutlook, Send-SiolaOutlookJob, Disconnect-SiolaOutlook
