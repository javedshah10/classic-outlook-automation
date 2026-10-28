<#
.SYNOPSIS
    Read-only. Compact stats for UNREAD mail in the Outlook Inbox: how many, by day or month, by sender,
    how many are not yet in the CSV (never synced), plus a short list. Nothing is saved or changed.
.EXAMPLE
    .\inbox_stats.ps1                       # all unread
    .\inbox_stats.ps1 -Days 2               # unread from today and yesterday
    .\inbox_stats.ps1 -Date 2026-10-25      # unread from that calendar day
    .\inbox_stats.ps1 -Sender newsletter -Max 15
#>
param(
    [int]$Days = 0,        # 0 = no date limit
    [string]$Date,
    [string]$Sender,       # text to match inside the sender address
    [int]$Max = 15         # rows in the list (counts always cover everything in the window)
)

$ErrorActionPreference = "Stop"
# Must match $TargetEmail in sync_inbox.ps1 (set_email.ps1 updates all scripts)
$TargetEmail = "your.name@example.com"
$PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
$PR_ATTACH_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
$PR_ATTACH_CID    = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
$inv = [System.Globalization.CultureInfo]::InvariantCulture
$Utf8Enc = if ($PSVersionTable.PSVersion.Major -ge 6) { "utf8BOM" } else { "UTF8" }

$f = "[UnRead] = True"; $label = "all unread"; $winStart = $null; $winEnd = $null   # dates are compared in code (see below)
if ($Date) {
    $s = [datetime]::ParseExact($Date, "yyyy-MM-dd", $inv)
    $winStart = $s; $winEnd = $s.AddDays(1)
    $label = "unread on " + $s.ToString("yyyy-MM-dd")
} elseif ($Days -gt 0) {
    $s = (Get-Date).Date.AddDays(-($Days - 1))
    $winStart = $s
    $label = "unread since " + $s.ToString("yyyy-MM-dd")
}

$ns = (New-Object -ComObject Outlook.Application).GetNamespace("MAPI")
$acc = $null; foreach ($a in $ns.Accounts) { if ($a.SmtpAddress -ieq $TargetEmail) { $acc = $a; break } }
if (-not $acc) { Write-Host "Account $TargetEmail not found." -ForegroundColor Red; exit 1 }
$inbox = $acc.DeliveryStore.GetDefaultFolder(6)
$items = $inbox.Items.Restrict($f); $items.Sort("[ReceivedTime]", $true)

# Message IDs already recorded in any master CSV (to spot mail the sync has never seen)
$known = @{}
foreach ($file in @(Get-ChildItem -Path (Join-Path $PSScriptRoot "data") -Filter "master_*.csv" -ErrorAction SilentlyContinue)) {
    foreach ($r in @(Import-Csv -Path $file.FullName -Encoding $Utf8Enc)) { if ($r.InternetMessageID) { $known[$r.InternetMessageID] = $true } }
}

function Get-Smtp ($m) {
    try { if ($m.SenderEmailType -eq "EX") { $x = $m.Sender.GetExchangeUser(); if ($x) { return $x.PrimarySmtpAddress } } } catch { }
    return [string]$m.SenderEmailAddress
}
function Has-RealAttachment ($m) {
    for ($a = 1; $a -le $m.Attachments.Count; $a++) {
        $att = $m.Attachments.Item($a); if ($att.Type -ne 1) { continue }
        $inline = $false
        try { $inline = [bool]$att.PropertyAccessor.GetProperty($PR_ATTACH_HIDDEN) } catch { }
        if (-not $inline -and [IO.Path]::GetExtension($att.FileName).ToLower() -match '^\.(png|jpe?g|gif)$') { try { if ($att.PropertyAccessor.GetProperty($PR_ATTACH_CID)) { $inline = $true } } catch { } }
        if (-not $inline) { return $true }
    }
    return $false
}

$rows = @()
for ($i = 1; $i -le $items.Count; $i++) {
    $m = $items.Item($i); if ($m.Class -ne 43) { continue }
    $rt = [datetime]$m.ReceivedTime; if ($winEnd -and $rt -ge $winEnd) { continue }; if ($winStart -and $rt -lt $winStart) { break }
    $addr = (Get-Smtp $m).ToLower()
    if ($Sender -and $addr -notlike "*$($Sender.ToLower())*") { continue }
    $mid = ""; try { $mid = $m.PropertyAccessor.GetProperty($PR_INTERNET_MESSAGE_ID) } catch { }
    $rows += [PSCustomObject]@{ Mail = $m; Received = [datetime]$m.ReceivedTime; From = $addr; Subject = ($m.Subject -replace '[\r\n]+', ' ').Trim(); InCsv = ($mid -and $known.ContainsKey($mid)) }
}

$notIn = @($rows | Where-Object { -not $_.InCsv }).Count
Write-Host "Inbox: $label$(if ($Sender) { " | sender contains '$Sender'" }) -> $($rows.Count) item(s) | not yet in CSV: $notIn" -ForegroundColor Cyan
if ($rows.Count -eq 0) { exit 0 }

$span = ($rows | Measure-Object Received -Minimum -Maximum); $monthly = (($span.Maximum - $span.Minimum).TotalDays -gt 45)
Write-Host ("By {0}:" -f $(if ($monthly) { "month" } else { "day" }))
$rows | Group-Object { $_.Received.ToString($(if ($monthly) { "yyyy-MM" } else { "yyyy-MM-dd" })) } | Sort-Object Name -Descending | ForEach-Object { "  {0}  {1}" -f $_.Name, $_.Count }
Write-Host "By sender (top 10):"
$rows | Group-Object From | Sort-Object Count -Descending | Select-Object -First 10 | ForEach-Object { "  {0,4}  {1}" -f $_.Count, $_.Name }
Write-Host "List (newest first, $([Math]::Min($Max, $rows.Count)) of $($rows.Count); * = not yet in CSV, A = has an attachment):"
$rows | Select-Object -First $Max | ForEach-Object {
    $subj = if ($_.Subject.Length -gt 50) { $_.Subject.Substring(0, 50) + "..." } else { $_.Subject }
    $att = if (Has-RealAttachment $_.Mail) { "A" } else { " " }
    "  {0}{1} {2:MM-dd HH:mm} | {3} | {4}" -f $(if ($_.InCsv) { " " } else { "*" }), $att, $_.Received, $_.From, $subj
}
