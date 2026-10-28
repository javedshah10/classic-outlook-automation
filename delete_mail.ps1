<#
.SYNOPSIS
    Deletes emails from the Outlook Inbox or Junk folder by sender and/or subject.
    SAFE BY DESIGN: dry run by default (add -Apply to delete), needs at least one filter,
    refuses to delete more than -Max emails at once, and only does a normal delete
    (emails go to Deleted Items; Deleted Items is never emptied and nothing is erased permanently).
    Emails that the sync recorded in the master CSV are set to Closed and their body file is
    moved to emails\deleted\, so the record stays consistent.
.EXAMPLE
    .\delete_mail.ps1 -Folder Junk -Sender fiestaslotspromo.com            # dry run
    .\delete_mail.ps1 -Folder Junk -Sender fiestaslotspromo.com -Max 300 -Apply
    .\delete_mail.ps1 -Folder Inbox -Sender noreply@example.com -Subject "weekly offer" -Apply
    .\delete_mail.ps1 -Folder Inbox -Subject "webinar" -Days 30
  -Sender rules:  "name@domain.com" = that exact address | "@domain.com" = that domain | "text" = address contains text
#>
param(
    [ValidateSet("Inbox", "Junk")][string]$Folder = "Inbox",
    [string]$Sender,
    [string]$Subject,      # subject must contain this text
    [int]$Days = 0,        # 0 = no date limit; otherwise only mail from the last N days (1 = today)
    [string]$Date,         # or one calendar day, yyyy-MM-dd
    [int]$Max = 50,        # safety stop: refuse to delete more than this in one run
    [switch]$Apply
)

$ErrorActionPreference = "Stop"
# Must match $TargetEmail in sync_inbox.ps1 (set_email.ps1 updates all scripts)
$TargetEmail = "your.name@example.com"
$PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
$inv = [System.Globalization.CultureInfo]::InvariantCulture
$Utf8Enc = if ($PSVersionTable.PSVersion.Major -ge 6) { "utf8BOM" } else { "UTF8" }

if (-not $Sender -and -not $Subject) { Write-Host "Give -Sender and/or -Subject so nothing is deleted by accident." -ForegroundColor Red; exit 1 }

$start = $null; $end = $null
if ($Date) { $start = [datetime]::ParseExact($Date, "yyyy-MM-dd", $inv); $end = $start.AddDays(1) }
elseif ($Days -gt 0) { $start = (Get-Date).Date.AddDays(-($Days - 1)); $end = (Get-Date).Date.AddDays(1) }

$ns = (New-Object -ComObject Outlook.Application).GetNamespace("MAPI")
$acc = $null; foreach ($a in $ns.Accounts) { if ($a.SmtpAddress -ieq $TargetEmail) { $acc = $a; break } }
if (-not $acc) { Write-Host "Account $TargetEmail not found." -ForegroundColor Red; exit 1 }
$fid = if ($Folder -eq "Junk") { 23 } else { 6 }
$fld = $acc.DeliveryStore.GetDefaultFolder($fid)
$items = $fld.Items; $items.Sort("[ReceivedTime]", $true)   # dates are compared in code, never in an Outlook filter

function Get-Smtp ($m) {
    try { if ($m.SenderEmailType -eq "EX") { $x = $m.Sender.GetExchangeUser(); if ($x) { return $x.PrimarySmtpAddress } } } catch { }
    return [string]$m.SenderEmailAddress
}
function Test-Sender ($addr, $rule) {
    $addr = $addr.ToLower(); $rule = $rule.ToLower().Trim()
    if ($rule.StartsWith("@")) { return $addr.EndsWith($rule) }
    if ($rule.Contains("@")) { return $addr -eq $rule }
    return $addr.Contains($rule)
}

$hits = @()
for ($i = 1; $i -le $items.Count; $i++) {
    $m = $items.Item($i); if ($m.Class -ne 43) { continue }
    $rt = [datetime]$m.ReceivedTime
    if ($end -and $rt -ge $end) { continue }
    if ($start -and $rt -lt $start) { break }
    $addr = Get-Smtp $m
    if ($Sender -and -not (Test-Sender $addr $Sender)) { continue }
    $subj = ($m.Subject -replace '[\r\n]+', ' ').Trim()
    if ($Subject -and $subj.IndexOf($Subject, [StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
    $hits += [PSCustomObject]@{ Mail = $m; From = $addr.ToLower(); Subject = $subj; Received = $rt }
}

$desc = @(); if ($Sender) { $desc += "sender '$Sender'" }; if ($Subject) { $desc += "subject contains '$Subject'" }
if ($start) { $desc += "window since " + $start.ToString("yyyy-MM-dd") + $(if ($Date) { " (that day)" }) }
Write-Host "$Folder folder ($($fld.Items.Count) items): $($desc -join ' + ') -> $($hits.Count) match(es)" -ForegroundColor Cyan
if ($hits.Count -eq 0) { Write-Host "Nothing matched. Nothing deleted." -ForegroundColor Yellow; exit 0 }

Write-Host "By sender (top 5):"
$hits | Group-Object From | Sort-Object Count -Descending | Select-Object -First 5 | ForEach-Object { "  {0,4}  {1}" -f $_.Count, $_.Name }
Write-Host "Sample (newest first, 8 of $($hits.Count)):"
$hits | Select-Object -First 8 | ForEach-Object { "  {0:yyyy-MM-dd HH:mm} | {1} | {2}" -f $_.Received, $_.From, $(if ($_.Subject.Length -gt 55) { $_.Subject.Substring(0, 55) + "..." } else { $_.Subject }) }

if ($hits.Count -gt $Max) { Write-Host "$($hits.Count) matches is more than -Max $Max. Narrow the filter, or raise -Max on purpose. Nothing deleted." -ForegroundColor Red; exit 1 }
if (-not $Apply) { Write-Host "DRY RUN: nothing deleted. Re-run with -Apply to move these $($hits.Count) to Deleted Items." -ForegroundColor Yellow; exit 0 }

# Message IDs first (needed to find CSV rows), then delete
$keys = @{}
foreach ($h in $hits) {
    $k = ""; try { $k = $h.Mail.PropertyAccessor.GetProperty($PR_INTERNET_MESSAGE_ID) } catch { }
    if (-not $k) { $k = $h.Mail.EntryID }
    $keys[$k] = $true
}
foreach ($h in $hits) { $h.Mail.Delete() }          # normal delete: goes to Deleted Items
Write-Host "Moved $($hits.Count) email(s) to Deleted Items." -ForegroundColor Green

$closed = 0
$DataDir = Join-Path $PSScriptRoot "data"; $DeletedDir = Join-Path $PSScriptRoot "emails\deleted"
foreach ($f in @(Get-ChildItem -Path $DataDir -Filter "master_*.csv" -ErrorAction SilentlyContinue)) {
    $rows = @(Import-Csv -Path $f.FullName -Encoding $Utf8Enc); $changed = $false
    foreach ($r in $rows) {
        if ($r.InternetMessageID -and $keys.ContainsKey($r.InternetMessageID) -and $r.Triage_Status -ne "Closed") {
            if ($r.BodyFile) {
                $src = Join-Path $PSScriptRoot $r.BodyFile
                if (Test-Path $src) {
                    if (-not (Test-Path $DeletedDir)) { New-Item -ItemType Directory -Path $DeletedDir | Out-Null }
                    Move-Item -Path $src -Destination (Join-Path $DeletedDir (Split-Path $src -Leaf)) -Force
                    $r.BodyFile = "emails/deleted/" + (Split-Path $src -Leaf)
                }
            }
            $r.Triage_Status = "Closed"; $closed++; $changed = $true
        }
    }
    if ($changed) { $rows | Export-Csv -Path $f.FullName -NoTypeInformation -Encoding $Utf8Enc }
}
Write-Host "CSV rows closed: $closed$(if ($closed) { '. Run sync_inbox.ps1 to refresh active_context.md.' })" -ForegroundColor Green
