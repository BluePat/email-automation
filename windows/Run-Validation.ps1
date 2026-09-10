Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
& (Join-Path $PSScriptRoot 'Invoke-SiolaAutomation.ps1') -Mode VALIDATE
