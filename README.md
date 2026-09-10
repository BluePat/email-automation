# SIOLA Google Sheets to Classic Outlook automation

This is a Windows-first email automation for a non-technical operator. It reads
the existing Google Sheet through the Google Sheets API, sends through a named
account in Classic Outlook, and writes delivery state back to the same sheet.

Development and tests never connect to or modify the real Google Sheet. The first
real connection is made during installation on the colleague's Windows computer.

## Production files

The deployable implementation is in `windows/`:

- `INSTALL.cmd` starts the guarded installer.
- `Invoke-SiolaAutomation.ps1` runs VALIDATE, TEST, or LIVE.
- `Siola.Core.psm1` validates, groups, deduplicates, and renders Czech HTML.
- `Siola.GoogleSheets.psm1` reads and updates exact Google Sheet cells.
- `Siola.Outlook.psm1` sends through Classic Outlook's COM interface.
- `Run-Validation.ps1`, `Run-Test.ps1`, and `Enable-SiolaLive.ps1` expose the
  operator-safe lifecycle.
- `Test-SiolaCore.ps1` is an executable PowerShell behavioral test and is run by
  the installer before credentials or the scheduled task are configured.

## Recipient and salutation contract

The two recipient types are generated independently:

| Job | Recipient | Salutation |
| --- | --- | --- |
| Mayor | `Email - STAROSTA` | `Oslovení - STAROSTA` |
| Secretary | `Email - TAJEMNÍK` | `Oslovení - TAJEMNÍK` |

The secretary is optional and is never placed in CC. A missing or inconsistent
required salutation is a validation error; the code does not guess.

## Safety properties

- Only the exact, case-sensitive cell value `K ODESLÁNÍ` is eligible.
- `K ODESLÁNÍ` is the operator's approval: all email-driving cells on those rows
  are considered validated and frozen. To correct them, remove the status first,
  edit and revalidate, then restore it.
- Only the sheet named exactly `Obce a města` is accessed.
- Columns are located by exact header text, not hard-coded letters.
- Applicants are grouped and projects are deduplicated by `Žadatel + Číslo RM`,
  including already-sent historical rows.
- Conflicting duplicates, mixed calls, bad grants, invalid addresses, or partially
  sent duplicate rows fail validation.
- LIVE verifies that the sheet belongs to this one installation, preflights Outlook,
  and then rereads, fingerprints, and claims only one applicant at a time.
- Non-send status updates receive the same fresh-row and header verification.
- TEST validates and creates a reviewable HTML preview for every eligible applicant,
  sends only the first three applicants to the configured test inbox, and never
  changes the sheet.
- Outlook sending is never automatically retried.
- A message is marked `ODESLÁNO` only after its hidden job identifier is found in
  Classic Outlook's Sent Items. Offline or stuck-Outbox messages fail closed.
- Outlook must confirm the configured sending account after `SendUsingAccount` is
  assigned, and the resolved recipient must match the requested SMTP address.
- A failure to write an Outlook result stops the run. The current applicant remains visibly in
  progress and must be reconciled against Sent Items before any reset.
- A filesystem lock prevents overlapping local runs.
- The scheduled task is installed disabled and runs only in the interactive,
  logged-in Windows session.
- Successful TEST completion is recorded for 24 hours and is required before LIVE
  can be enabled. Runtime logs are retained for 90 days.
- Google tokens refresh during long runs; HTTP calls have bounded timeouts and
  quota-aware retry delays.

`ODESLÁNO` means that Classic Outlook moved the marked message into Sent Items. It
does not guarantee acceptance by the recipient's remote mail server.

## Requirements

- Windows 10 or 11 with PowerShell 7.
- Classic Outlook installed and configured with the sending account.
- The Windows user must remain signed in at the scheduled time.
- A Google service account with the Google Sheets API enabled.
- The source Google Sheet shared with that service account as Editor.

Classic Outlook automation is intentional: Microsoft documents that its desktop
Outlook actions and Outlook Object Model do not support New Outlook.

## Setup and operation

Follow `docs/INSTALACE_CZ.md` on the Windows computer. Routine recovery rules are
in `docs/OBSLUHA_CZ.md`.

The setup has three gates:

1. VALIDATE reads only and sends nothing.
2. TEST validates and previews all messages, sends a sample only to the configured
   test inbox, records successful completion, and writes nothing to the sheet.
3. `Enable-SiolaLive.ps1` repeats validation and requires the operator to type
   `LIVE` before enabling the scheduled task.

To move production to another Windows computer, disable the old scheduled task and
use `TAKE_OVER.cmd` on the fully tested replacement. The old installation then fails
its ownership check before it can claim or send anything.
