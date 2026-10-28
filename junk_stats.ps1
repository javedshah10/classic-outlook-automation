<#
.SYNOPSIS
    Read-only. Shows compact stats for the Outlook Junk folder: counts per day and per sender,
    plus a short list with a stable ID per email so you can ask for one to be moved to the Inbox.
    Nothing is saved to disk and nothing is changed.
.EXAMPLE
    .\junk_stats.ps1                       # today
    .\junk_stats.ps1 -Days 3               # today and the 2 days before
    .\junk_stats.ps1 -Date 2026-10-25      # that calendar day
    .\junk_stats.ps1 -Days 7 -Sender example.com -Max 40
#>
param(
    [int]$Days = 1,
    [string]$Date,
    [string]$Sender,     # text to match inside the sender address, e.g. example.com
    [int]$Max = 25       # rows in the list (counts always cover everything in the window)
)

$ErrorActionPreference = "Stop"
# Must match $TargetEmail in sync_inbox.ps1 (set_email.ps1 updates all scripts)
$TargetEmail = "your.name@example.com"
$PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
$inv = [System.Globalization.CultureInfo]::InvariantCulture

if ($Date) {
    $start = [datetime]::ParseExact($Date, "yyyy-MM-dd", $inv); $end = $start.AddDays(1)
    $label = $start.ToString("yyyy-MM-dd")
} else {
    $start = (Get-Date).Date.AddDays(-($Days - 1)); $end = (Get-Date).Date.AddDays(1)
    $label = "last $Days day(s), since " + $start.ToString("yyyy-MM-dd")
}

$ns = (New-Object -ComObject Outlook.Application).GetNamespace("MAPI")
$acc = $null; foreach ($a in $ns.Accounts) { if ($a.SmtpAddress -ieq $TargetEmail) { $acc = $a; break } }
if (-not $acc) { Write-Host "Account $TargetEmail not found." -ForegroundColor Red; exit 1 }
$junk = $acc.DeliveryStore.GetDefaultFolder(23)   # 23 = olFolderJunk

# No date filter in Outlook: Restrict reads 10/11/2026 as 10 November on some PCs. Sort newest first and compare dates here.
$items = $junk.Items; $items.Sort("[ReceivedTime]", $true)

function Get-Smtp ($m) {
    try { if ($m.SenderEmailType -eq "EX") { $x = $m.Sender.GetExchangeUser(); if ($x) { return $x.PrimarySmtpAddress } } } catch { }
    return [string]$m.SenderEmailAddress
}
function Get-JunkId ($m) {
    $k = ""; try { $k = $m.PropertyAccessor.GetProperty($PR_INTERNET_MESSAGE_ID) } catch { }
    if (-not $k) { $k = $m.EntryID }
    $sha = [System.Security.Cryptography.SHA1]::Create()
    "J" + (($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($k)) | Select-Object -First 4 | ForEach-Object { $_.ToString("X2") }) -join "")
}

$rows = @()
for ($i = 1; $i -le $items.Count; $i++) {
    $m = $items.Item($i); if ($m.Class -ne 43) { continue }
    $rt = [datetime]$m.ReceivedTime; if ($rt -ge $end) { continue }; if ($rt -lt $start) { break }   # newest first; dates compared here, not in Outlook
    $addr = (Get-Smtp $m).ToLower()
    if ($Sender -and $addr -notlike "*$($Sender.ToLower())*") { continue }
    $rows += [PSCustomObject]@{ Id = Get-JunkId $m; Received = [datetime]$m.ReceivedTime; From = $addr; Subject = ($m.Subject -replace '[\r\n]+', ' ').Trim() }
}

Write-Host "Junk folder: $($junk.Items.Count) items in total | window: $label$(if ($Sender) { " | sender contains '$Sender'" }) -> $($rows.Count) item(s)" -ForegroundColor Cyan
if ($rows.Count -eq 0) { exit 0 }

Write-Host "By day:"
$rows | Group-Object { $_.Received.ToString("yyyy-MM-dd") } | Sort-Object Name -Descending | ForEach-Object { "  {0}  {1}" -f $_.Name, $_.Count }
Write-Host "By sender (top 10):"
$rows | Group-Object From | Sort-Object Count -Descending | Select-Object -First 10 | ForEach-Object { "  {0,4}  {1}" -f $_.Count, $_.Name }
Write-Host "List (newest first, $([Math]::Min($Max, $rows.Count)) of $($rows.Count); use -Max / -Sender to narrow):"
$rows | Select-Object -First $Max | ForEach-Object {
    $s = if ($_.Subject.Length -gt 55) { $_.Subject.Substring(0, 55) + "..." } else { $_.Subject }
    "  {0} | {1:MM-dd HH:mm} | {2} | {3}" -f $_.Id, $_.Received, $_.From, $s
}
Write-Host "To rescue one: .\move_from_junk.ps1 -Id <ID> -Days <same window> [-Apply]"
