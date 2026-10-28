<#
.SYNOPSIS
    Moves chosen emails from the Outlook Junk folder back to the Inbox. Nothing is saved to disk.
    Default is a DRY RUN that only lists what it would move; add -Apply to move.
    Pick emails by -Id (from junk_stats.ps1) or by -Sender text; use the same -Days / -Date window.
.EXAMPLE
    .\move_from_junk.ps1 -Id JA1B2C3D4 -Days 3            # dry run
    .\move_from_junk.ps1 -Id JA1B2C3D4,J0F9E8D7C -Days 3 -Apply
    .\move_from_junk.ps1 -Sender sender@example.com -Date 2026-10-25 -Apply
#>
param(
    [string[]]$Id,
    [string]$Sender,
    [int]$Days = 7,
    [string]$Date,
    [int]$Max = 20,        # safety stop: refuse to move more than this many in one run
    [switch]$Apply
)

$ErrorActionPreference = "Stop"
# Must match $TargetEmail in sync_inbox.ps1 (set_email.ps1 updates all scripts)
$TargetEmail = "your.name@example.com"
$PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
$inv = [System.Globalization.CultureInfo]::InvariantCulture

# Attachment flag only: names are listed, the files stay in Outlook and are not saved
$PR_ATTACH_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
$PR_ATTACH_CID    = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
function Get-AttachmentNames ($mail) {
    $names = @()
    for ($a = 1; $a -le $mail.Attachments.Count; $a++) {
        $att = $mail.Attachments.Item($a)
        if ($att.Type -ne 1) { continue }
        $ext = [IO.Path]::GetExtension($att.FileName).ToLower()
        $inline = $false
        try { $inline = [bool]$att.PropertyAccessor.GetProperty($PR_ATTACH_HIDDEN) } catch { }
        if (-not $inline -and $ext -match '^\.(png|jpe?g|gif)$') { try { if ($att.PropertyAccessor.GetProperty($PR_ATTACH_CID)) { $inline = $true } } catch { } }
        if (-not $inline) { $names += $att.FileName }
    }
    return $names
}

if (-not $Id -and -not $Sender) { Write-Host "Give -Id and/or -Sender so nothing is moved by accident." -ForegroundColor Red; exit 1 }

if ($Date) { $start = [datetime]::ParseExact($Date, "yyyy-MM-dd", $inv); $end = $start.AddDays(1) }
else { $start = (Get-Date).Date.AddDays(-($Days - 1)); $end = (Get-Date).Date.AddDays(1) }

$ns = (New-Object -ComObject Outlook.Application).GetNamespace("MAPI")
$acc = $null; foreach ($a in $ns.Accounts) { if ($a.SmtpAddress -ieq $TargetEmail) { $acc = $a; break } }
if (-not $acc) { Write-Host "Account $TargetEmail not found." -ForegroundColor Red; exit 1 }
$junk  = $acc.DeliveryStore.GetDefaultFolder(23)
$inbox = $acc.DeliveryStore.GetDefaultFolder(6)

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

$hits = @(); $idSeen = @{}
for ($i = 1; $i -le $items.Count; $i++) {
    $m = $items.Item($i); if ($m.Class -ne 43) { continue }
    $rt = [datetime]$m.ReceivedTime; if ($rt -ge $end) { continue }; if ($rt -lt $start) { break }   # newest first; dates compared here, not in Outlook
    $jid = Get-JunkId $m
    $addr = (Get-Smtp $m).ToLower()
    $okId = (-not $Id) -or ($Id -contains $jid)
    $okSender = (-not $Sender) -or ($addr -like "*$($Sender.ToLower())*")
    if ($okId -and $okSender) { $hits += [PSCustomObject]@{ Id = $jid; Mail = $m; From = $addr; Received = [datetime]$m.ReceivedTime; Subject = ($m.Subject -replace '[\r\n]+', ' ').Trim() } }
}

# An ID must point to exactly one email; two mails sharing one ID would be ambiguous.
foreach ($g in ($hits | Group-Object Id | Where-Object { $_.Count -gt 1 })) { Write-Host "ID $($g.Name) matches $($g.Count) emails; narrow the window or use -Sender." -ForegroundColor Red; exit 1 }
foreach ($want in @($Id)) { if ($want -and -not ($hits | Where-Object { $_.Id -eq $want })) { Write-Host "ID $want not found in this window (try a wider -Days or the right -Date)." -ForegroundColor Yellow } }

if ($hits.Count -eq 0) { Write-Host "Nothing matched. Nothing moved." -ForegroundColor Yellow; exit 0 }
if ($hits.Count -gt $Max) { Write-Host "$($hits.Count) emails match, more than -Max $Max. Narrow the selection or raise -Max on purpose. Nothing moved." -ForegroundColor Red; exit 1 }

Write-Host "$($hits.Count) email(s) selected:" -ForegroundColor Cyan
$hits | ForEach-Object { "  {0} | {1:yyyy-MM-dd HH:mm} | {2} | {3}" -f $_.Id, $_.Received, $_.From, $_.Subject }
if (-not $Apply) { Write-Host "DRY RUN: nothing moved. Re-run with -Apply to move these to the Inbox." -ForegroundColor Yellow; exit 0 }

$Utf8Enc = if ($PSVersionTable.PSVersion.Major -ge 6) { "utf8BOM" } else { "UTF8" }
$DataDir = Join-Path $PSScriptRoot "data"; $EmailsDir = Join-Path $PSScriptRoot "emails"
$MasterCsv = Join-Path $DataDir ("master_" + (Get-Date).ToString("yyyy_MM") + ".csv")
$DefaultCols = "ShortID,EntryID,InternetMessageID,ThreadID,ReceivedTime,SenderName,SenderEmail,Subject,IsRead_Outlook,IsReplied_Outlook,BodyFile,Triage_Status,Category,Summary,Attachment_Count,Attachment_Files,Has_Attachment".Split(",")
if (-not (Test-Path $EmailsDir)) { New-Item -ItemType Directory -Path $EmailsDir | Out-Null }

# Every month's master CSV, to skip emails the sync already recorded (matched by the stable message ID)
$known = @{}
foreach ($f in @(Get-ChildItem -Path $DataDir -Filter "master_*.csv" -ErrorAction SilentlyContinue)) {
    foreach ($r in @(Import-Csv -Path $f.FullName -Encoding $Utf8Enc)) { if ($r.InternetMessageID) { $known[$r.InternetMessageID] = $r.ShortID } }
}
$Master = @(); if (Test-Path $MasterCsv) { $Master = @(Import-Csv -Path $MasterCsv -Encoding $Utf8Enc) }
$cols = if ($Master.Count) { $Master[0].PSObject.Properties.Name } else { $DefaultCols }

$recorded = 0
foreach ($h in $hits) {
    $m = $h.Mail
    $mid = ""; try { $mid = $m.PropertyAccessor.GetProperty($PR_INTERNET_MESSAGE_ID) } catch { }
    $sname = ($m.SenderName -replace '[\r\n]+', ' ').Trim(); $tid = $m.ConversationID
    $rt0 = [datetime]$m.ReceivedTime; $rt = New-Object DateTime ($rt0.Ticks - ($rt0.Ticks % [TimeSpan]::TicksPerSecond))
    $body = $m.Body
    $moved = $m.Move($inbox)                         # the moved item has a new EntryID in the Inbox
    Write-Host "  [MOVED TO INBOX] $($h.Id)" -ForegroundColor Green
    if (-not $mid) { $mid = $moved.EntryID }
    if ($known.ContainsKey($mid)) { Write-Host "    already recorded as $($known[$mid]); no new row" -ForegroundColor Yellow; continue }

    $sid = "MSG_" + $rt.ToString("yyyyMMdd_HHmmss") + "_" + [guid]::NewGuid().ToString("N").Substring(0, 4)
    $rel = "emails/$sid.md"
    $attNames = @(Get-AttachmentNames $moved)
    $attBlock = if ($attNames.Count) { "`r`n## Attachments (not saved; open the email in Outlook)`r`n" + (($attNames | ForEach-Object { "- $_" }) -join "`r`n") + "`r`n" } else { "" }
    $md = @"
# Email: $($h.Subject)
- **ShortID:** $sid
- **ThreadID:** $tid
- **From:** $sname <$($h.From)>
- **Received:** $($rt.ToString("yyyy-MM-dd HH:mm:ss"))
- **Outlook Replied Flag:** False
- **Rescued from Junk:** yes (moved to Inbox by move_from_junk.ps1)
$attBlock

---

$body
"@
    Set-Content -Path (Join-Path $PSScriptRoot $rel) -Value $md -Encoding $Utf8Enc

    $vals = @{ ShortID = $sid; EntryID = $moved.EntryID; InternetMessageID = $mid; ThreadID = $tid
               ReceivedTime = $rt.ToString("yyyy-MM-dd HH:mm:ss"); SenderName = $sname; SenderEmail = $h.From; Subject = $h.Subject
               IsRead_Outlook = [string](-not $moved.UnRead); IsReplied_Outlook = "False"; BodyFile = $rel
               Triage_Status = "New"; Category = "Unclassified"; Summary = ""
               Attachment_Count = [string]$attNames.Count; Attachment_Files = ($attNames -join ";")
               Has_Attachment = if ($attNames.Count) { "Yes" } else { "No" } }
    $row = [ordered]@{}; foreach ($c in $cols) { $row[$c] = if ($vals.ContainsKey($c)) { $vals[$c] } else { "" } }
    $Master += [PSCustomObject]$row; $known[$mid] = $sid; $recorded++
    $Master | Export-Csv -Path $MasterCsv -NoTypeInformation -Encoding $Utf8Enc   # saved per email, so a later failure cannot lose this record
    Write-Host "    recorded in CSV as $sid (Triage_Status = New)" -ForegroundColor Green
}
Write-Host "Done. $($hits.Count) moved, $recorded recorded in the CSV. Run sync_inbox.ps1 to refresh active_context.md, then triage them." -ForegroundColor Green
