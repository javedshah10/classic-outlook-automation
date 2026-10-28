<#
.SYNOPSIS
    Replaces the current Outlook email address with your own across the solution files.
.EXAMPLE
    .\set_email.ps1 -NewEmail "student@outlook.com"
    .\set_email.ps1            # prompts for the address
    Works on Windows PowerShell 5.1 and PowerShell 7. Only the exact address text is replaced.
#>
param([string]$NewEmail)

$ErrorActionPreference = "Stop"
$BaseDir = $PSScriptRoot
$Sync    = Join-Path $BaseDir "sync_inbox.ps1"

# Current address = the value of $TargetEmail in sync_inbox.ps1
$m = Select-String -Path $Sync -Pattern '^\s*\$TargetEmail\s*=\s*"([^"]+)"' | Select-Object -First 1
if (-not $m) { Write-Host "Could not find the `$TargetEmail line in sync_inbox.ps1." -ForegroundColor Red; exit 1 }
$OldEmail = $m.Matches[0].Groups[1].Value
Write-Host "Current address: $OldEmail" -ForegroundColor Cyan

# Show the accounts Classic Outlook itself reports, so you can copy the right address (skipped quietly if Outlook is not open)
$found = @()
try { foreach ($a in (New-Object -ComObject Outlook.Application).GetNamespace("MAPI").Accounts) { $found += [string]$a.SmtpAddress } } catch { }
if ($found.Count) { Write-Host "Accounts found in your Outlook (use one of these exactly):" -ForegroundColor Cyan; $found | ForEach-Object { Write-Host "  $_" } }
else { Write-Host "Could not read the account list (is Classic Outlook open?). Use the address shown in Outlook under File > Account Settings." -ForegroundColor Yellow }

if (-not $NewEmail) { $NewEmail = Read-Host "Enter your Outlook email address" }
$NewEmail = $NewEmail.Trim()
if ($NewEmail -notmatch '^[^@\s"]+@[^@\s"]+\.[^@\s"]+$') { Write-Host "'$NewEmail' doesn't look like an email address. Nothing changed." -ForegroundColor Red; exit 1 }
if ($found.Count -and ($found -notcontains $NewEmail)) { Write-Host "Warning: '$NewEmail' is not one of the Outlook accounts listed above. The scripts will say 'Account ... not found' unless it matches one exactly." -ForegroundColor Yellow }
if ($NewEmail -ieq $OldEmail) { Write-Host "That is already the current address. Nothing changed." -ForegroundColor Yellow; exit 0 }

# Solution files in this folder only (not the data/emails/drafts output folders)
$files = Get-ChildItem -Path (Join-Path $BaseDir "*") -File -Include *.ps1,*.md,*.json,*.txt |
         Where-Object { $_.FullName -ne $PSCommandPath -and $_.Name -ne "active_context.md" }   # active_context.md is generated email data, rebuilt on each sync

$changed = 0
foreach ($f in $files) {
    $hits = @(Select-String -Path $f.FullName -Pattern ([regex]::Escape($OldEmail)))
    if ($hits.Count -eq 0) { continue }
    foreach ($h in $hits) { Write-Host ("  {0}:{1}  {2}" -f $f.Name, $h.LineNumber, $h.Line.Trim()) }
    $bytes  = [IO.File]::ReadAllBytes($f.FullName)
    $hasBom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
    $text   = [IO.File]::ReadAllText($f.FullName)
    $text   = $text.Replace($OldEmail, $NewEmail)          # exact text only
    [IO.File]::WriteAllText($f.FullName, $text, (New-Object Text.UTF8Encoding($hasBom)))
    $changed += $hits.Count
}

Write-Host "Replaced $changed occurrence(s) of $OldEmail with $NewEmail." -ForegroundColor Green
