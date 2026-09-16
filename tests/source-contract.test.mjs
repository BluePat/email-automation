import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

const core = await readFile(new URL("../windows/Siola.Core.psm1", import.meta.url), "utf8");
const outlook = await readFile(new URL("../windows/Siola.Outlook.psm1", import.meta.url), "utf8");
const runner = await readFile(new URL("../windows/Invoke-SiolaAutomation.ps1", import.meta.url), "utf8");
const installer = await readFile(new URL("../windows/Install-SiolaAutomation.ps1", import.meta.url), "utf8");
const installCmd = await readFile(new URL("../windows/INSTALL.cmd", import.meta.url), "utf8");
const google = await readFile(new URL("../windows/Siola.GoogleSheets.psm1", import.meta.url), "utf8");
const enableLive = await readFile(new URL("../windows/Enable-SiolaLive.ps1", import.meta.url), "utf8");
const operatingManual = await readFile(new URL("../docs/OBSLUHA_CZ.md", import.meta.url), "utf8");
const configExampleText = await readFile(new URL("../windows/config.example.json", import.meta.url), "utf8");
const configExample = JSON.parse(configExampleText);

test("eligibility uses an exact case-sensitive status comparison", () => {
  assert.match(core, /Get-CellText \$_\.Status\) -ceq \$script:ReadyStatus/);
  assert.match(core, /\$script:ReadyStatus = 'K ODESLÁNÍ'/);
});

test("mayor job binds mayor email and mayor salutation", () => {
  assert.match(
    core,
    /New-EmailJob -Role STAROSTA -Email \$mayorEmail\.Value[\s\S]*?-Salutation \$mayorSalutation\.Value/,
  );
});

test("secretary job binds secretary email and secretary salutation", () => {
  assert.match(
    core,
    /New-EmailJob -Role TAJEMNIK -Email \$secretaryEmail\.Value[\s\S]*?-Salutation \$secretarySalutation\.Value/,
  );
});

test("Outlook sends only to the job recipient and does not populate CC or BCC", () => {
  assert.match(outlook, /Recipients\.Add\(\[string\]\$Job\.To\)/);
  assert.doesNotMatch(outlook, /\.CC\s*=/i);
  assert.doesNotMatch(outlook, /\.BCC\s*=/i);
});

test("HTML contract retains Calibri 12 and the required bold question", () => {
  assert.match(core, /font-family:Calibri,Arial,sans-serif;font-size:12pt/);
  assert.match(core, /<strong>máte již zajištěnou administraci/);
});

test("LIVE connects Outlook before any claim is written", () => {
  const sendGroup = runner.indexOf("$freshGroup = Resolve-SiolaFreshEligibleGroup");
  const claimPosition = runner.indexOf("Set-SiolaTargetUpdates -Targets $locatedTargets", sendGroup);
  const outlookPosition = runner.indexOf("Connect-SiolaOutlook", runner.indexOf("# LIVE"));
  assert.ok(claimPosition >= 0);
  assert.ok(outlookPosition >= 0);
  assert.ok(outlookPosition < claimPosition);
});

test("installer leaves the scheduled task disabled", () => {
  assert.match(installer, /Disable-ScheduledTask -TaskName \$TaskName/);
});

test("operating manual covers safe disable and missed-run behavior", () => {
  assert.match(operatingManual, /## Dočasné vypnutí automatu/);
  assert.match(operatingManual, /Zakázání zabrání budoucím spuštěním, ale \*\*nezastaví právě probíhající běh\*\*/);
  assert.match(operatingManual, /Nezapínejte úlohu přímo v Plánovači/);
  assert.match(operatingManual, /## Když počítač není v naplánovaný čas připravený/);
  assert.match(operatingManual, /Automat počítač neprobudí/);
  assert.match(operatingManual, /automatické opakování po chybě/);
  assert.match(installer, /New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew/);
  assert.match(installer, /New-ScheduledTaskPrincipal -UserId \$identity -LogonType Interactive/);
});

test("LIVE verifies every email-driving row using a complete fingerprint", () => {
  for (const field of ["Call", "Applicant", "ProjectName", "Grant", "SecretarySalutation", "SecretaryEmail", "MayorSalutation", "MayorEmail", "Status", "MayorStatus", "SecretaryStatus"]) {
    assert.match(runner, new RegExp(`${field} =`));
  }
  assert.match(runner, /Resolve-SiolaFreshEligibleGroup/);
  assert.match(runner, /Resolve-SiolaClaimedGroup/);
});

test("RM-related columns do not drive validation, grouping, or approval", () => {
  assert.doesNotMatch(core, /RmNumber|Číslo RM/);
  assert.doesNotMatch(runner, /RmNumber|Číslo RM/);
});

test("the Google Sheet is bound to one installation", () => {
  assert.match(google, /siola_automation_owner/);
  assert.match(enableLive, /Register-SiolaAutomationOwner/);
  assert.match(runner, /Assert-SiolaAutomationOwner/);
  assert.match(runner, /machineBinding/);
});

test("Outlook completion is confirmed in Sent Items", () => {
  assert.match(outlook, /BillingInformation/);
  assert.match(outlook, /Wait-SiolaOutlookJobSent/);
  assert.match(outlook, /lastLocation -eq 'SENT'/);
});

test("resolved Outlook SMTP address must equal the requested recipient", () => {
  assert.match(outlook, /Get-SiolaRecipientSmtpAddress/);
  assert.match(outlook, /\$resolvedSmtp -ine \[string\]\$Job\.To/);
});

test("TEST validates and previews every applicant but sends only the configured sample", () => {
  assert.match(runner, /Write-SiolaPreviewReport/);
  assert.match(runner, /\$batchSize = if \(\$effectiveMode -eq 'LIVE'\)/);
  assert.match(runner, /Select-Object -First \(\[int\]\$config\.testBatchSize\)/);
});

test("LIVE requires a recent successful TEST receipt", () => {
  assert.match(enableLive, /test-success\.json/);
  assert.match(enableLive, /AddHours\(-24\)/);
});

test("sent state is matched exactly instead of by prefix", () => {
  assert.match(core, /\^ODESLÁNO\(\?:/);
  assert.doesNotMatch(core, /StartsWith\(\$script:SentStatus/);
});

test("sent timestamps are written as native serial values", () => {
  assert.match(runner, /\[DateTime\]::Now\.ToOADate\(\)/);
});

test("fresh-machine installer avoids the parenthesized expansion trap", () => {
  assert.match(installCmd, /if not errorlevel 1 goto :run_installer/);
  assert.doesNotMatch(installCmd, /if errorlevel 1 \([\s\S]*?set "PWSH=/);
});

test("a success log failure cannot flip a confirmed outcome", () => {
  const send = runner.indexOf("Send-SiolaOutlookJob", runner.indexOf("$outcomes = @{}"));
  const sendCatch = runner.indexOf("catch {", send);
  const confirmedLog = runner.indexOf("OUTLOOK_SENT_CONFIRMED", send);
  assert.ok(send >= 0 && sendCatch > send);
  assert.ok(confirmedLog > sendCatch);
  assert.match(runner, /catch \{\s*\$script:LogDegraded = \$true/);
});

test("Google access tokens are refreshed during LIVE processing", () => {
  assert.match(google, /Get-SiolaFreshAccessToken/);
  assert.match(runner, /Get-GoogleServiceAccountToken[^\n]+-AsSession/);
  assert.ok((runner.match(/Get-SiolaFreshAccessToken/g) ?? []).length >= 3);
});

test("Google HTTP calls have bounded timeouts and quota-aware 429 delay", () => {
  assert.match(google, /-TimeoutSec 60/);
  assert.match(google, /TimeoutSec = \$requestTimeoutSeconds/);
  assert.match(google, /\$status -eq 429\) \{ 65 \}/);
  assert.match(google, /RetryAfter/);
  assert.match(google, /Časový limit běhu vypršel během opakování/);
});

test("non-send status writes receive the same freshness protection", () => {
  assert.match(runner, /\$nonSendGroups/);
  assert.match(runner, /Resolve-SiolaFreshEligibleGroup -Group \$nonSendGroup/);
});

test("Outlook verifies that SendUsingAccount stuck", () => {
  assert.match(outlook, /function Assert-SiolaMailSendingAccount/);
  assert.match(outlook, /\$assignedAccount = \$Mail\.SendUsingAccount/);
  assert.match(outlook, /\$actualSmtpAddress -ine \$ExpectedSmtpAddress/);
  assert.equal((outlook.match(/Assert-SiolaMailSendingAccount -Mail \$mail/g) ?? []).length, 1);
  const body = outlook.indexOf("$mail.HTMLBody = [string]$Job.BodyHtml");
  const account = outlook.indexOf("$mail.SendUsingAccount = $OutlookContext.Account", body);
  const verify = outlook.indexOf("Assert-SiolaMailSendingAccount -Mail $mail", account);
  const send = outlook.indexOf("$mail.Send()", verify);
  assert.ok(body >= 0 && account > body && verify > account && send > verify);
});

test("ambiguous US-formatted grant text fails closed", () => {
  assert.match(core, /Accept only unambiguous Czech grouping/);
  assert.match(core, /notmatch '\^\[-\+\]\?\\d\+\(\?:,\\d\{1,2\}\)\?\$'/);
});

test("secretary salutation without secretary email is a validation error", () => {
  assert.match(core, /Je vyplněno Oslovení - TAJEMNÍK, ale chybí Email - TAJEMNÍK/);
});

test("explicit missing role markers suppress only that recipient", () => {
  assert.match(core, /Test-SiolaExplicitMissingValue/);
  assert.match(core, /@\('není', 'neni'\)/);
  assert.match(core, /SecretaryEmail 'Email - TAJEMNÍK' \$false \$true `[\s\S]*?-AllowMissingMarker/);
  assert.match(core, /MayorEmail 'Email - STAROSTA' \$false \$true `[\s\S]*?-AllowMissingMarker/);
  assert.match(core, /Chybí příjemce: není k dispozici e-mail starosty ani tajemníka/);
  assert.match(runner, /\$mayorSent = \(-not \$group\.MayorRequired\)/);
});

test("only the first call per applicant is emailed and later calls receive a distinct state", () => {
  assert.match(core, /\$applicantRows = @\(\$applicantGroup\.Group \| Sort-Object RowNumber\)/);
  assert.match(core, /\$primaryCallKey = Get-CanonicalText \$applicantRows\[0\]\.Call/);
  assert.match(core, /PrimaryRowNumbers = \[int\[\]\]@\(\$groupRows\.RowNumber\)/);
  assert.match(core, /SuppressedRowNumbers = \[int\[\]\]@\(\$suppressedRows \| ForEach-Object/);
  assert.match(core, /SelectedCall = \[string\]\$applicantRows\[0\]\.Call/);
  assert.match(core, /Chybí Výzva na řádcích/);
  assert.ok((core.match(/selectedCall = \[string\]/g) ?? []).length >= 2);
  assert.doesNotMatch(core, /Žadatel má projekty v různých výzvách/);
  assert.match(runner, /function Test-SiolaSuppressedProjectRow/);
  assert.match(runner, /\$rowCall -cne \$selectedCall/);
  assert.doesNotMatch(runner, /SuppressedRowNumbers\) -contains \$RowNumber/);
  assert.match(runner, /KONTAKTOVÁNO JINÝM PROJEKTEM/);
  assert.match(runner, /\$suppressedStatus = if \(\$anySent\)/);
  assert.match(runner, /if \(-not \(Test-SiolaSuppressedProjectRow[\s\S]*?foreach \(\$claimJob in \$group\.Jobs\)/);
});

test("Sent Items confirmation uses exact Items.Find lookup", () => {
  assert.match(outlook, /\$items\.Find\("\[BillingInformation\]/);
  assert.doesNotMatch(outlook, /\$limit = \[math\]::Min/);
});

test("installer trims and validates both sender and test recipient", () => {
  assert.match(installer, /\$sender = Read-RequiredText 'Odesílající účet/);
  assert.match(installer, /Assert-EmailAddress -Value \$sender/);
  assert.match(installer, /Assert-EmailAddress -Value \$testRecipient -Label 'Testovací e-mail'/);
});

test("public examples contain no live spreadsheet or contact defaults", () => {
  assert.equal(configExample.spreadsheetId, "REPLACE_WITH_GOOGLE_SHEET_ID");
  assert.match(configExample.outlookSenderSmtpAddress, /@example\.com$/);
  assert.match(configExample.testRecipient, /@example\.com$/);
  assert.equal("signature" in configExample, false);
  assert.equal("approvedOutlookSignatureFingerprint" in configExample, false);
  assert.doesNotMatch(installer, /výchozí: dodaná databáze/);
  assert.match(installer, /\$spreadsheetInput = Read-RequiredText 'Google Sheet URL nebo ID'/);
});

test("email template contains the fixed approved signature", () => {
  assert.match(core, /data-siola-signature="fixed-v1"/);
  assert.match(core, /<strong>Jan Burian<\/strong>/);
  assert.match(core, /href="tel:\+420608229916"/);
  assert.match(core, /href="mailto:jan\.burian@siolagroup\.cz"/);
  assert.match(core, /href="mailto:siola@siolagroup\.cz"/);
  assert.match(core, /Informace obsažené v této zprávě mohou být důvěrného charakteru/);
  assert.match(core, /This e-mail may contain privileged and confidential information\./);
  assert.match(core, /Before you print it, think about the ENVIRONMENT\./);
  assert.match(core, /color:#70ad47/);
  assert.match(core, /ApprovalBodyHtml = \$approvalBodyHtml/);
  assert.match(outlook, /\$mail\.HTMLBody = \[string\]\$Job\.BodyHtml/);
  assert.doesNotMatch(outlook, /DefaultSignature|SignatureFingerprint|Wait-SiolaMailDefaultSignature|\.Display\(/);
  assert.doesNotMatch(enableLive, /Connect-SiolaOutlook|outlookSignatureFingerprint|Ověřuji výchozí podpis/);
  assert.doesNotMatch(installer, /(?:Read-Host|Read-RequiredText)[^\n]*(?:podpis|signature)/i);
  assert.doesNotMatch(installer, /approvedOutlookSignatureFingerprint/);
  assert.match(runner, /@\('signature', 'approvedOutlookSignatureFingerprint'\)/);
  assert.match(runner, /PSObject\.Properties\.Remove\(\$legacyProperty\)/);
  assert.match(runner, /\[IO\.File\]::Replace\(\$migrationPath, \$configFullPath, \$null\)/);
  assert.doesNotMatch(runner, /SignatureHtml|Merge-SiolaOutlookSignature|outlookSignatureFingerprint\s*=|config\.approvedOutlookSignatureFingerprint/);
});

test("bootstrap failures create an operator notification", () => {
  const bootstrapTry = runner.indexOf("try {", runner.indexOf("bootstrap-"));
  const configCheck = runner.indexOf("Konfigurace neexistuje", bootstrapTry);
  const bootstrapNotify = runner.indexOf("Show-SiolaFailureNotification", configCheck);
  assert.ok(bootstrapTry >= 0 && configCheck > bootstrapTry && bootstrapNotify > configCheck);
});

test("LIVE gate rejects a zero-message TEST receipt", () => {
  assert.match(enableLive, /receipt\.sentTestJobs -lt 1/);
});

test("LIVE approval is bound to reviewed content and runtime files", () => {
  assert.match(runner, /approvalDigest = Get-SiolaBatchApprovalFingerprint/);
  assert.match(runner, /runtimeFingerprint = Get-SiolaRuntimeFingerprint/);
  assert.match(enableLive, /receipt\.approvalDigest -cne \$currentApprovalDigest/);
  assert.match(enableLive, /receipt\.runtimeFingerprint -cne \$currentRuntimeFingerprint/);
});

test("LIVE owns and renews a distributed lease before sending", () => {
  assert.match(google, /siola_automation_lease/);
  assert.match(runner, /Enter-SiolaAutomationLease/);
  const send = runner.indexOf("Send-SiolaOutlookJob", runner.indexOf("$outcomes = @{}"));
  const renew = runner.lastIndexOf("Update-SiolaAutomationLease", send);
  const claimCheck = runner.lastIndexOf("Resolve-SiolaClaimedGroup", send);
  assert.ok(renew >= 0 && renew < send);
  assert.ok(claimCheck >= 0 && claimCheck < send);
  const transfer = google.indexOf("function Transfer-SiolaAutomationOwner");
  assert.ok(google.indexOf("Enter-SiolaAutomationLease", transfer) > transfer);
});

test("claims and final writes follow unique row metadata instead of A1 coordinates", () => {
  assert.match(google, /metadataKey = 'siola_row_target'/);
  assert.match(google, /batchUpdateByDataFilter/);
  assert.match(google, /locationType = 'ROW'/);
  assert.match(runner, /Get-SiolaVerifiedRowTargets/);
  assert.match(runner, /Assert-SiolaTargetApproval/);
  const finalResolve = runner.indexOf("$finalTargets = @(");
  const finalWrite = runner.indexOf("Set-SiolaTargetUpdates -Targets $finalTargets", finalResolve);
  assert.ok(finalResolve >= 0 && finalWrite > finalResolve);
  assert.doesNotMatch(runner, /Set-SiolaSheetCells/);
});

test("expired leases are replaced with a new immutable metadata record", () => {
  const enter = google.slice(google.indexOf("function Enter-SiolaAutomationLease"), google.indexOf("function Update-SiolaAutomationLease"));
  assert.match(enter, /deleteDeveloperMetadata/);
  assert.match(enter, /createDeveloperMetadata/);
  assert.doesNotMatch(enter, /Set-SiolaAutomationLeaseValue/);
  assert.match(google, /Get-SiolaAutomationLeases/);
  assert.match(google, /AddSeconds\(180\)/);
  assert.match(google, /Set-SiolaAutomationLeaseValue[\s\S]*?-MaxAttempts 1/);
});

test("configuration bounds and scheduler deadline fail safely", () => {
  assert.match(runner, /runDeadlineMinutes musí být mezi 30 a 330/);
  assert.match(runner, /RUN_DEADLINE_REACHED/);
  assert.match(installer, /ExecutionTimeLimit \(New-TimeSpan -Hours 6\)/);
  assert.match(runner, /googleCallBudgetSeconds/);
  assert.match(runner, /Set-SiolaGoogleRequestMaxAttempts -MaxAttempts 1/);
});

test("installation ACLs protect operational data and key cleanup is explicit", () => {
  assert.match(installer, /icacls\.exe \$InstallDirectory/);
  assert.match(installer, /\$sourceFullPath -ine \$destinationFullPath/);
  assert.match(installer, /napište přesně SMAZAT/);
  assert.doesNotMatch(installer, /SendToRecycleBin/);
});

test("LIVE removes the reviewed preview and receipt", () => {
  assert.match(enableLive, /receipt\.previewPath, \$receiptPath/);
  assert.match(enableLive, /Remove-Item -LiteralPath \$sensitivePath/);
  assert.match(enableLive, /Get-ChildItem -LiteralPath \$previewDirectory -Filter 'preview-\*\.html'/);
  assert.match(runner, /previousReceipt\.previewPath/);
});

test("forced reinstall disables the task and acquires the runtime lock before copying", () => {
  const disable = installer.indexOf("Disable-ScheduledTask", installer.indexOf("if ($Force"));
  const lock = installer.indexOf("[IO.File]::Open($existingAutomationLock", disable);
  const stage = installer.indexOf("$stagingDirectory =", lock);
  assert.ok(disable >= 0 && lock > disable && stage > lock);
  assert.match(installer, /State -ceq 'Running'/);
  assert.match(installer, /Move-Item -LiteralPath \(Join-Path \$stagingDirectory/);
});

test("send-group row targets are cleaned in finally without masking final status success", () => {
  assert.match(runner, /finally \{[\s\S]*?ROW_TARGET_CLEANUP_FAILED/);
  const finalWrite = runner.indexOf("Set-SiolaTargetUpdates -Targets $finalTargets");
  const resultCatch = runner.indexOf("RESULT_WRITE_FAILED", finalWrite);
  const cleanup = runner.indexOf("Remove-SiolaRowTargets", resultCatch);
  assert.ok(finalWrite >= 0 && resultCatch > finalWrite && cleanup > resultCatch);
});

test("validation errors surface as a failed scheduled run", () => {
  assert.match(runner, /throw "Běh dokončil odesílání, ale našel/);
  assert.match(runner, /throw "Tabulka obsahuje \$\(\$batch\.ValidationErrorCount\) chyb validace/);
});
