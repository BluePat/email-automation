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
  const claimPosition = runner.indexOf("Set-SiolaSheetCells -SpreadsheetId", runner.indexOf("$claims"));
  const outlookPosition = runner.indexOf("Connect-SiolaOutlook", runner.indexOf("# LIVE"));
  assert.ok(claimPosition >= 0);
  assert.ok(outlookPosition >= 0);
  assert.ok(outlookPosition < claimPosition);
});

test("installer leaves the scheduled task disabled", () => {
  assert.match(installer, /Disable-ScheduledTask -TaskName \$TaskName/);
});

test("LIVE verifies every email-driving row using a complete fingerprint", () => {
  for (const field of ["Call", "RmNumber", "Applicant", "ProjectName", "Grant", "SecretarySalutation", "SecretaryEmail", "MayorSalutation", "MayorEmail", "Status", "MayorStatus", "SecretaryStatus"]) {
    assert.match(runner, new RegExp(`${field} =`));
  }
  assert.match(runner, /Assert-SiolaGroupUnchanged/);
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
  assert.match(google, /TimeoutSec = 60/);
  assert.match(google, /\$status -eq 429\) \{ 65 \}/);
  assert.match(google, /RetryAfter/);
});

test("non-send status writes receive the same freshness protection", () => {
  assert.match(runner, /\$nonSendGroups/);
  assert.match(runner, /Assert-SiolaGroupUnchanged -Group \$nonSendGroup/);
});

test("Outlook verifies that SendUsingAccount stuck", () => {
  assert.match(outlook, /\$assignedAccount = \$mail\.SendUsingAccount/);
  assert.match(outlook, /assignedAccount\.SmtpAddress -ine/);
});

test("ambiguous US-formatted grant text fails closed", () => {
  assert.match(core, /Accept only unambiguous Czech grouping/);
  assert.match(core, /notmatch '\^\[-\+\]\?\\d\+\(\?:,\\d\{1,2\}\)\?\$'/);
});

test("secretary salutation without secretary email is a validation error", () => {
  assert.match(core, /Je vyplněno Oslovení - TAJEMNÍK, ale chybí Email - TAJEMNÍK/);
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
  assert.match(configExample.signature.email, /@example\.com$/);
  assert.match(configExample.signature.companyEmail, /@example\.com$/);
  assert.doesNotMatch(installer, /výchozí: dodaná databáze/);
  assert.match(installer, /\$spreadsheetInput = Read-RequiredText 'Google Sheet URL nebo ID'/);
});

test("email signature comes from validated local configuration", () => {
  assert.match(core, /\$Signature\.Name/);
  assert.match(core, /\$Signature\.CompanyEmail/);
  assert.match(runner, /V konfiguraci chybí signature/);
  assert.match(runner, /-Signature \$config\.signature/);
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
