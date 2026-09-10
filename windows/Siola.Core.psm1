Set-StrictMode -Version Latest

$script:ReadyStatus = 'K ODESLÁNÍ'
$script:ProcessingStatus = 'ZPRACOVÁVÁ SE'
$script:SentStatus = 'ODESLÁNO'

function Get-CellText {
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return '' }
    return [string]$Value
}

function Get-CleanText {
    param([AllowNull()]$Value)
    return ([regex]::Replace((Get-CellText $Value), '\s+', ' ')).Trim()
}

function Get-DisplayText {
    param([AllowNull()]$Value)
    return (Get-CellText $Value).Trim()
}

function Get-CanonicalText {
    param([AllowNull()]$Value)
    return (Get-CleanText $Value).ToLowerInvariant()
}

function Test-SiolaSentStatus {
    param([AllowNull()]$Value)
    $text = Get-CleanText $Value
    return [regex]::IsMatch($text, '^ODESLÁNO(?:\s*\|\s*.+)?$', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
}

function ConvertTo-HtmlText {
    param([AllowNull()]$Value)
    return [System.Net.WebUtility]::HtmlEncode((Get-DisplayText $Value))
}

function Test-EmailAddress {
    param([string]$Value)
    try {
        $address = [System.Net.Mail.MailAddress]::new((Get-CleanText $Value))
        return $address.Address -ieq (Get-CleanText $Value)
    }
    catch { return $false }
}

function ConvertFrom-GrantValue {
    param([AllowNull()]$Value)

    if ($Value -is [byte] -or $Value -is [int16] -or $Value -is [int32] -or
        $Value -is [int64] -or $Value -is [single] -or $Value -is [double] -or
        $Value -is [decimal]) {
        return [double]$Value
    }

    $valueText = (Get-CleanText $Value) -replace [char]0x00A0, ''
    $valueText = $valueText -replace '\s', '' -replace '(?i)Kč', '' -replace ',-$', ''
    if ([string]::IsNullOrWhiteSpace($valueText)) { return [double]::NaN }

    if ($valueText.Contains(',') -and $valueText.Contains('.')) {
        # Accept only unambiguous Czech grouping such as 1.057.223,50.
        if ($valueText -notmatch '^[-+]?\d{1,3}(?:\.\d{3})+(?:,\d{1,2})?$') { return [double]::NaN }
        $valueText = $valueText.Replace('.', '').Replace(',', '.')
    }
    elseif ($valueText.Contains(',')) {
        # A single Czech decimal comma may have at most two monetary decimals.
        if ($valueText -notmatch '^[-+]?\d+(?:,\d{1,2})?$') { return [double]::NaN }
        $valueText = $valueText.Replace(',', '.')
    }
    elseif ($valueText.Contains('.')) {
        if ($valueText -match '^[-+]?\d{1,3}(?:\.\d{3})+$') {
            $valueText = $valueText.Replace('.', '')
        }
        elseif ($valueText -notmatch '^[-+]?\d+\.\d{1,2}$') { return [double]::NaN }
    }
    elseif ($valueText -notmatch '^[-+]?\d+$') { return [double]::NaN }

    [double]$parsed = 0
    $styles = [Globalization.NumberStyles]::Float -bor [Globalization.NumberStyles]::AllowLeadingSign
    if ([double]::TryParse($valueText, $styles, [Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) {
        return $parsed
    }
    return [double]::NaN
}

function Format-GrantCzk {
    param([double]$Value)
    if ([double]::IsNaN($Value) -or [double]::IsInfinity($Value)) {
        throw 'Dotace není platné číslo.'
    }
    $rounded = [math]::Round($Value, 0, [MidpointRounding]::AwayFromZero)
    $formatted = $rounded.ToString('#,0', [Globalization.CultureInfo]::InvariantCulture).Replace(',', '.')
    return "$formatted,- Kč"
}

function Get-ShortApplicant {
    param([string]$Applicant)
    $display = Get-DisplayText $Applicant
    $prefixes = @('Statutární město', 'Hlavní město', 'Městská část', 'Městský obvod', 'Městys', 'Město', 'Obec')
    foreach ($prefix in $prefixes) {
        if ($display.StartsWith("$prefix ", [StringComparison]::OrdinalIgnoreCase)) {
            $short = $display.Substring($prefix.Length).Trim()
            if ($short) { return $short }
        }
    }
    return $display
}

function Join-CzechList {
    param([string[]]$Items)
    if ($Items.Count -eq 0) { return '' }
    if ($Items.Count -eq 1) { return $Items[0] }
    if ($Items.Count -eq 2) { return "$($Items[0]) a $($Items[1])" }
    return "$(($Items[0..($Items.Count - 2)] -join ', ')) a $($Items[-1])"
}

function Get-UniqueValues {
    param(
        [object[]]$Rows,
        [string]$Property,
        [switch]$CaseInsensitive
    )
    $result = [Collections.Generic.List[string]]::new()
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($row in $Rows) {
        $value = Get-DisplayText $row.$Property
        if (-not $value) { continue }
        $key = if ($CaseInsensitive) { Get-CanonicalText $value } else { $value }
        if ($seen.Add($key)) { $result.Add($value) }
    }
    return @($result)
}

function Resolve-SingleValue {
    param(
        [object[]]$Rows,
        [string]$Property,
        [string]$Label,
        [bool]$Required,
        [bool]$CaseInsensitive = $false
    )
    $values = @(Get-UniqueValues -Rows $Rows -Property $Property -CaseInsensitive:$CaseInsensitive)
    $blankRows = @($Rows | Where-Object { -not (Get-CleanText $_.$Property) })
    if (($Required -or $values.Count -gt 0) -and $blankRows.Count -gt 0) {
        return [pscustomobject]@{
            Value = ''
            Error = "Chybí $Label na řádcích: $($blankRows.RowNumber -join ', ')."
        }
    }
    if ($Required -and $values.Count -eq 0) {
        return [pscustomobject]@{ Value = ''; Error = "Chybí $Label." }
    }
    if ($values.Count -gt 1) {
        return [pscustomobject]@{ Value = ''; Error = "Rozpor v poli ${Label}: $($values -join ' | ')." }
    }
    return [pscustomobject]@{ Value = $(if ($values.Count) { $values[0] } else { '' }); Error = '' }
}

function Get-RoleSentState {
    param([object[]]$Rows, [string]$StatusProperty, [string]$Label)
    $sent = @($Rows | Where-Object { Test-SiolaSentStatus $_.$StatusProperty })
    if ($sent.Count -gt 0 -and $sent.Count -lt $Rows.Count) {
        return [pscustomobject]@{
            AlreadySent = $false
            Error = "Nekonzistentní ${Label}: jen část řádků je označena ODESLÁNO."
        }
    }
    return [pscustomobject]@{ AlreadySent = ($sent.Count -eq $Rows.Count); Error = '' }
}

function Get-Projects {
    param([object[]]$Rows, [Collections.Generic.List[string]]$Errors)
    $groups = @($Rows | Group-Object { Get-CanonicalText $_.RmNumber })
    $projects = [Collections.Generic.List[object]]::new()

    foreach ($group in $groups) {
        $projectRows = @($group.Group)
        $rm = Get-DisplayText $projectRows[0].RmNumber
        if (-not $rm) {
            foreach ($row in $projectRows) { $Errors.Add("Řádek $($row.RowNumber): chybí Číslo RM.") }
            continue
        }

        $names = @(Get-UniqueValues -Rows $projectRows -Property ProjectName)
        $calls = @(Get-UniqueValues -Rows $projectRows -Property Call)
        $amounts = @($projectRows | ForEach-Object { ConvertFrom-GrantValue $_.Grant })

        if ($names.Count -ne 1) { $Errors.Add("Číslo RM ${rm}: chybí nebo se liší Název akce.") }
        if ($calls.Count -ne 1) { $Errors.Add("Číslo RM ${rm}: chybí nebo se liší Výzva.") }
        if (@($amounts | Where-Object { [double]::IsNaN($_) -or $_ -le 0 }).Count -gt 0) {
            $Errors.Add("Číslo RM ${rm}: Dotace (Kč) musí být kladné číslo.")
        }

        $validAmounts = @($amounts | Where-Object { -not [double]::IsNaN($_) })
        if ($validAmounts.Count -gt 1) {
            $first = $validAmounts[0]
            if (@($validAmounts | Where-Object { [math]::Abs($_ - $first) -gt 0.01 }).Count -gt 0) {
                $Errors.Add("Číslo RM ${rm}: duplicitní řádky mají rozdílnou Dotaci (Kč).")
            }
        }

        if ($names.Count -eq 1 -and $calls.Count -eq 1 -and $validAmounts.Count -gt 0) {
            $projects.Add([pscustomobject]@{
                Rm = $rm
                Name = $names[0]
                Call = $calls[0]
                Grant = [double]$validAmounts[0]
            })
        }
    }
    return @($projects)
}

function Get-SiolaRowApprovalFingerprint {
    param([Parameter(Mandatory)]$Row)
    $payload = [ordered]@{
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

function New-EmailHtml {
    param(
        [string]$Salutation,
        [object[]]$Projects,
        [string]$Call,
        [string]$FormattedGrant,
        [Parameter(Mandatory)][object]$Signature,
        [string]$TestOriginalRecipient = ''
    )
    $quotedNames = @($Projects | ForEach-Object { "&bdquo;$(ConvertTo-HtmlText $_.Name)&ldquo;" })
    if ($Projects.Count -eq 1) {
        $projectPhrase = "projektem <strong>$($quotedNames[0])</strong>. Na jeho realizaci byla"
        $grantPhrase = 'dotace ve výši'
    }
    else {
        $projectPhrase = "projekty <strong>$(Join-CzechList $quotedNames)</strong>. Na jejich realizaci byla"
        $grantPhrase = 'dotace v celkové výši'
    }
    $banner = ''
    if ($TestOriginalRecipient) {
        $banner = "<div style=`"background:#fff2cc;border:1px solid #d6b656;padding:8px;margin-bottom:14px`"><strong>TEST:</strong> Původní příjemce: $(ConvertTo-HtmlText $TestOriginalRecipient)</div>"
    }

    return @"
<div style="font-family:Calibri,Arial,sans-serif;font-size:12pt;line-height:1.25;color:#000">
$banner<p>$(ConvertTo-HtmlText $Salutation)</p>
<p>obracím se na Vás konkrétně v souvislosti s $projectPhrase v programu Modernizačního fondu <strong>$(ConvertTo-HtmlText $Call)</strong> schválena $grantPhrase <strong>$(ConvertTo-HtmlText $FormattedGrant)</strong>.</p>
<p>Nevím, v jaké fázi se nyní projekt nachází, proto si dovoluji krátký dotaz: <strong>máte již zajištěnou administraci dokončení projektu a následného doložení realizace k vyplacení dotace vůči SFŽP?</strong></p>
<p>Obcím v této fázi projektů zajišťujeme zejména:</p>
<ul>
<li><strong>doložení realizace a vyplacení dotace</strong> – kontrolu podkladů a způsobilosti výdajů, žádost o platbu, závěrečné vyhodnocení akce a komunikaci se SFŽP až do uzavření projektu;</li>
<li><strong>následnou správu dotace v době udržitelnosti</strong> – hlídání dotačních povinností, potřebná hlášení a výkazy, změny projektu a součinnost při případných kontrolách.</li>
</ul>
<p>Pokud projekt teprve realizujete, můžeme jeho dotační administraci převzít již nyní. Pokud máte doložení realizace zajištěné, můžeme navázat až správou pětileté udržitelnosti.</p>
<p>V uplynulém roce jsme zajišťovali dotační administraci pro 27 klientů a spravovali projekty s dotacemi přesahujícími 76 mil. Kč.</p>
<p><strong>Stačí mi prosím krátká informace, zda tuto agendu již máte zajištěnou, nebo zda má smysl se o projektu krátce pobavit.</strong> Pokud ji má na starosti někdo jiný, budu Vám vděčný za přesměrování.</p>
<p>Se zdvořilým pozdravem,</p>
<p>$(ConvertTo-HtmlText $Signature.Name)</p>
<p style="color:#333399">tel: $(ConvertTo-HtmlText $Signature.Phone)<br>
e-mail: <a href="mailto:$(ConvertTo-HtmlText $Signature.Email)">$(ConvertTo-HtmlText $Signature.Email)</a></p>
<p style="color:#333399">$(ConvertTo-HtmlText $Signature.Company)<br>
$(ConvertTo-HtmlText $Signature.Address)<br>
Email: <a href="mailto:$(ConvertTo-HtmlText $Signature.CompanyEmail)">$(ConvertTo-HtmlText $Signature.CompanyEmail)</a><br>
IČ: $(ConvertTo-HtmlText $Signature.CompanyId)</p>
<p style="font-size:8pt;color:#333399">Informace obsažené v této zprávě mohou být důvěrného charakteru a mohou požívat zvláštní ochrany. Jsou určeny výhradně uvedeným adresátům. Pokud nejste adresátem, obratem nás, prosím, informujte (zasláním zprávy zpět odesílateli) a zprávu vymažte ze systému. Bez řádně vydaného souhlasu je zakázáno informace obsažené ve zprávě jakýmkoliv způsobem používat či je dále šířit.</p>
<p style="font-size:8pt;color:#333399">This e-mail may contain privileged and confidential information. It is intended for the named recipients only. If you are not an intended recipient, please notify us immediately (by reply e-mail) and delete this e-mail from your system. Any use or retransmission without proper authorization is prohibited.</p>
<p style="font-size:8pt;color:#70ad47">Before you print it, think about the ENVIRONMENT.</p>
</div>
"@
}

function New-EmailJob {
    param(
        [string]$Role,
        [string]$Email,
        [string]$Salutation,
        [string]$Applicant,
        [object[]]$Projects,
        [string]$Call,
        [double]$TotalGrant,
        [int[]]$RowNumbers,
        [string]$RunId,
        [int]$Ordinal,
        [string]$Mode,
        [string]$TestRecipient,
        [Parameter(Mandatory)][object]$Signature
    )
    $jobId = "$RunId-$Ordinal-$Role"
    $baseSubject = "$(Get-ShortApplicant $Applicant) – projekt FVE / dotace $Call"
    $isTest = $Mode -eq 'TEST'
    $approvalBodyHtml = New-EmailHtml -Salutation $Salutation -Projects $Projects -Call $Call `
        -FormattedGrant (Format-GrantCzk $TotalGrant) -Signature $Signature
    return [pscustomobject]@{
        JobId = $jobId
        Role = $Role
        Applicant = $Applicant
        IntendedTo = $Email
        To = $(if ($isTest) { $TestRecipient } else { $Email })
        Subject = $(if ($isTest) { "[TEST – původně $Email] $baseSubject" } else { $baseSubject })
        ApprovalSubject = $baseSubject
        ApprovalBodyHtml = $approvalBodyHtml
        BodyHtml = (New-EmailHtml -Salutation $Salutation -Projects $Projects -Call $Call `
            -FormattedGrant (Format-GrantCzk $TotalGrant) -Signature $Signature `
            -TestOriginalRecipient $(if ($isTest) { $Email } else { '' }))
        RowNumbers = $RowNumbers
    }
}

function Get-SiolaPreparedBatch {
    [CmdletBinding()]
    param(
        [object[]]$Rows,
        [ValidateSet('VALIDATE', 'TEST', 'LIVE')][string]$Mode,
        [string]$RunId,
        [int]$BatchSize = 50,
        [string]$TestRecipient = '',
        [Parameter(Mandatory)][object]$Signature
    )
    if (-not (Get-CleanText $RunId)) { throw 'Chybí runId.' }
    if ($Mode -eq 'TEST' -and -not (Test-EmailAddress $TestRecipient)) {
        throw 'V režimu TEST musí být platný testovací příjemce.'
    }
    if ($BatchSize -lt 1) { throw 'BatchSize musí být alespoň 1.' }
    $requiredSignatureFields = @('Name', 'Phone', 'Email', 'Company', 'Address', 'CompanyEmail', 'CompanyId')
    foreach ($field in $requiredSignatureFields) {
        if (-not $Signature.PSObject.Properties[$field] -or -not (Get-CleanText $Signature.$field)) {
            throw "V podpisu chybí $field."
        }
    }
    if (-not (Test-EmailAddress $Signature.Email) -or -not (Test-EmailAddress $Signature.CompanyEmail)) {
        throw 'Podpis obsahuje neplatnou e-mailovou adresu.'
    }

    $eligible = @($Rows | Where-Object { (Get-CellText $_.Status) -ceq $script:ReadyStatus })
    $applicantGroups = @($eligible | Group-Object { Get-CanonicalText $_.Applicant } |
        Sort-Object { ($_.Group | Measure-Object RowNumber -Minimum).Minimum })
    $validGroups = [Collections.Generic.List[object]]::new()
    $invalidGroups = [Collections.Generic.List[object]]::new()
    $completedGroups = [Collections.Generic.List[object]]::new()
    [int]$ordinal = 0
    [int]$selected = 0

    foreach ($applicantGroup in $applicantGroups) {
        if ($selected -ge $BatchSize) { break }
        $selected++
        $groupRows = @($applicantGroup.Group)
        $errors = [Collections.Generic.List[string]]::new()
        $applicant = Resolve-SingleValue $groupRows Applicant 'Žadatel' $true
        $mayorEmail = Resolve-SingleValue $groupRows MayorEmail 'Email - STAROSTA' $true $true
        $mayorSalutation = Resolve-SingleValue $groupRows MayorSalutation 'Oslovení - STAROSTA' $true
        $secretaryEmail = Resolve-SingleValue $groupRows SecretaryEmail 'Email - TAJEMNÍK' $false $true
        $secretarySalutation = Resolve-SingleValue $groupRows SecretarySalutation 'Oslovení - TAJEMNÍK' ([bool]$secretaryEmail.Value)

        foreach ($resolved in @($applicant, $mayorEmail, $mayorSalutation, $secretaryEmail, $secretarySalutation)) {
            if ($resolved.Error) { $errors.Add($resolved.Error) }
        }
        if (-not $secretaryEmail.Value -and $secretarySalutation.Value) {
            $errors.Add('Je vyplněno Oslovení - TAJEMNÍK, ale chybí Email - TAJEMNÍK.')
        }

        if ($applicant.Value) {
            $currentRowNumbers = [Collections.Generic.HashSet[int]]::new()
            foreach ($groupRow in $groupRows) { $null = $currentRowNumbers.Add([int]$groupRow.RowNumber) }
            foreach ($groupRow in $groupRows) {
                $rmKey = Get-CanonicalText $groupRow.RmNumber
                if (-not $rmKey) { continue }
                $historicalMatches = @($Rows | Where-Object {
                    -not $currentRowNumbers.Contains([int]$_.RowNumber) -and
                    (Get-CanonicalText $_.Applicant) -eq (Get-CanonicalText $applicant.Value) -and
                    (Get-CanonicalText $_.RmNumber) -eq $rmKey -and
                    ((Test-SiolaSentStatus $_.Status) -or (Test-SiolaSentStatus $_.MayorStatus) -or
                        (Test-SiolaSentStatus $_.SecretaryStatus))
                })
                if ($historicalMatches.Count -gt 0) {
                    $errors.Add("Číslo RM $($groupRow.RmNumber) už existuje jako odeslané na řádcích: $($historicalMatches.RowNumber -join ', ').")
                }
            }
        }
        if ($mayorEmail.Value -and -not (Test-EmailAddress $mayorEmail.Value)) {
            $errors.Add('Email - STAROSTA nemá platný formát.')
        }
        if ($secretaryEmail.Value -and -not (Test-EmailAddress $secretaryEmail.Value)) {
            $errors.Add('Email - TAJEMNÍK nemá platný formát.')
        }

        $projects = @(Get-Projects -Rows $groupRows -Errors $errors)
        $calls = @($projects | ForEach-Object { $_.Call } | Select-Object -Unique)
        if ($calls.Count -gt 1) {
            $errors.Add("Žadatel má projekty v různých výzvách: $($calls -join ' | ').")
        }

        $mayorState = Get-RoleSentState $groupRows MayorStatus 'stav STAROSTA'
        if ($mayorState.Error) { $errors.Add($mayorState.Error) }
        $secretaryState = [pscustomobject]@{ AlreadySent = $true; Error = '' }
        if ($secretaryEmail.Value) {
            $secretaryState = Get-RoleSentState $groupRows SecretaryStatus 'stav TAJEMNÍK'
            if ($secretaryState.Error) { $errors.Add($secretaryState.Error) }
        }

        if ($errors.Count -gt 0) {
            $invalidGroups.Add([pscustomobject]@{
                Applicant = $(if ($applicant.Value) { $applicant.Value } else { "(řádek $($groupRows[0].RowNumber))" })
                MatchApplicant = [string]$groupRows[0].Applicant
                RowNumbers = [int[]]@($groupRows.RowNumber)
                ApprovalRowFingerprints = [string[]]@($groupRows | ForEach-Object { Get-SiolaRowApprovalFingerprint $_ } | Sort-Object)
                Errors = [string[]]@($errors)
            })
            continue
        }

        $jobs = [Collections.Generic.List[object]]::new()
        $totalGrant = [double](($projects | Measure-Object Grant -Sum).Sum)
        if (-not $mayorState.AlreadySent) {
            $ordinal++
            $jobs.Add((New-EmailJob -Role STAROSTA -Email $mayorEmail.Value `
                -Salutation $mayorSalutation.Value -Applicant $applicant.Value -Projects $projects `
                -Call $calls[0] -TotalGrant $totalGrant -RowNumbers ([int[]]@($groupRows.RowNumber)) `
                -RunId $RunId -Ordinal $ordinal -Mode $Mode -TestRecipient $TestRecipient -Signature $Signature))
        }
        if ($secretaryEmail.Value -and -not $secretaryState.AlreadySent) {
            $ordinal++
            $jobs.Add((New-EmailJob -Role TAJEMNIK -Email $secretaryEmail.Value `
                -Salutation $secretarySalutation.Value -Applicant $applicant.Value -Projects $projects `
                -Call $calls[0] -TotalGrant $totalGrant -RowNumbers ([int[]]@($groupRows.RowNumber)) `
                -RunId $RunId -Ordinal $ordinal -Mode $Mode -TestRecipient $TestRecipient -Signature $Signature))
        }

        $prepared = [pscustomobject]@{
            Applicant = $applicant.Value
            MatchApplicant = $applicant.Value
            RowNumbers = [int[]]@($groupRows.RowNumber)
            ApprovalRowFingerprints = [string[]]@($groupRows | ForEach-Object { Get-SiolaRowApprovalFingerprint $_ } | Sort-Object)
            Jobs = [object[]]@($jobs)
            MayorRequired = $true
            MayorAlreadySent = [bool]$mayorState.AlreadySent
            SecretaryRequired = [bool]$secretaryEmail.Value
            SecretaryAlreadySent = [bool]$secretaryState.AlreadySent
        }
        if ($jobs.Count -eq 0) { $completedGroups.Add($prepared) } else { $validGroups.Add($prepared) }
    }

    [int]$jobCount = 0
    foreach ($validGroup in $validGroups) { $jobCount += $validGroup.Jobs.Count }
    return [pscustomobject]@{
        Mode = $Mode
        RunId = $RunId
        EligibleRowCount = $eligible.Count
        SelectedApplicantCount = $selected
        JobCount = $jobCount
        ValidationErrorCount = $invalidGroups.Count
        Groups = [object[]]@($validGroups)
        InvalidGroups = [object[]]@($invalidGroups)
        CompletedGroups = [object[]]@($completedGroups)
    }
}

function Get-SiolaBatchApprovalFingerprint {
    param([Parameter(Mandatory)]$Batch)
    $groupRecords = @($Batch.Groups | ForEach-Object {
        $group = $_
        [ordered]@{
            applicant = [string]$group.Applicant
            rows = [string[]]@($group.ApprovalRowFingerprints | Sort-Object)
            jobs = [object[]]@($group.Jobs | Sort-Object Role | ForEach-Object {
                [ordered]@{
                    role = [string]$_.Role
                    intendedTo = [string]$_.IntendedTo
                    subject = [string]$_.ApprovalSubject
                    bodyHtml = [string]$_.ApprovalBodyHtml
                }
            })
        }
    } | Sort-Object applicant)
    $completedRecords = @($Batch.CompletedGroups | ForEach-Object {
        [ordered]@{
            applicant = [string]$_.Applicant
            rows = [string[]]@($_.ApprovalRowFingerprints | Sort-Object)
        }
    } | Sort-Object applicant)
    $payload = [ordered]@{
        eligibleRowCount = [int]$Batch.EligibleRowCount
        groups = [object[]]$groupRecords
        completedGroups = [object[]]$completedRecords
    }
    $bytes = [Text.Encoding]::UTF8.GetBytes(($payload | ConvertTo-Json -Depth 10 -Compress))
    return [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))
}

function Get-SiolaRuntimeFingerprint {
    param([Parameter(Mandatory)][string]$RootPath)
    $files = @(
        'Invoke-SiolaAutomation.ps1',
        'Siola.Core.psm1',
        'Siola.GoogleSheets.psm1',
        'Siola.Outlook.psm1',
        'Enable-SiolaLive.ps1'
    )
    $records = [Collections.Generic.List[string]]::new()
    foreach ($file in $files) {
        $path = Join-Path $RootPath $file
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Chybí provozní soubor $file." }
        $records.Add("$file=$((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash)")
    }
    $bytes = [Text.Encoding]::UTF8.GetBytes(($records -join "`n"))
    return [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))
}

function ConvertTo-A1Column {
    param([int]$OneBasedColumn)
    if ($OneBasedColumn -lt 1) { throw 'Číslo sloupce musí být kladné.' }
    $result = ''
    $number = $OneBasedColumn
    while ($number -gt 0) {
        $number--
        $result = ([char](65 + ($number % 26))).ToString() + $result
        $number = [math]::Floor($number / 26)
    }
    return $result
}

Export-ModuleMember -Function Get-SiolaPreparedBatch, ConvertTo-A1Column, Test-SiolaSentStatus, `
    Get-SiolaRowApprovalFingerprint, Get-SiolaBatchApprovalFingerprint, Get-SiolaRuntimeFingerprint
