<#
.SYNOPSIS
    Finds an email anywhere in your Outlook account, however old: Inbox (with its subfolders), Sent, Junk, Deleted, Archive and any
    other mail folder. Searching is read-only: Outlook is never changed. Add -Save with the -Id of an email to copy it into the
    local files (emails\<ShortID>.md + a CSV row, status New) so the agent can read, triage and draft a reply to it.
    The same filters must be used for -Save as for the search that showed the ID.
.EXAMPLE
    .\search_mail.ps1 -Sender supplier.com -Since 2026-01-01                 # everything from that sender since January
    .\search_mail.ps1 -Subject "invoice 4471" -Until 2026-03-31              # by subject text, up to a date
    .\search_mail.ps1 -Text "purchase order" -Folder deleted                 # words inside the body, Deleted Items only
    .\search_mail.ps1 -Sender supplier.com -Subject "invoice" -Id S1A2B3C4D -Save
  -Folder: All (default) | Inbox | Sent | Junk | Deleted | any part of a folder name (e.g. Archive, Clients)
  -Sender matches the address or the display name. -Subject and -Text are "contains" matches; -Text reads the body (slower).
#>
param(
    [string]$Sender,
    [string]$Subject,
    [string]$Text,
    [string]$Since,          # yyyy-MM-dd, first day to include
    [string]$Until,          # yyyy-MM-dd, last day to include
    [int]$Days = 0,          # or: only the last N days (1 = today)
    [string]$Folder = "All",
    [string[]]$Id,           # pick email(s) from a previous search result
    [switch]$Save,           # copy the picked email(s) into emails\ and the CSV (needs -Id); Outlook itself is not changed
    [int]$Max = 25,          # rows shown (the total match count is always reported)
    [int]$Scan = 5000        # newest N items examined per folder; shortcut: raise it, or add -Since, for very large folders
)

$ErrorActionPreference = "Stop"
# Must match $TargetEmail in sync_inbox.ps1 (set_email.ps1 updates all scripts)
$TargetEmail = "your.name@example.com"
$PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
$inv = [System.Globalization.CultureInfo]::InvariantCulture
$Utf8Enc = if ($PSVersionTable.PSVersion.Major -ge 6) { "utf8BOM" } else { "UTF8" }

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

if ($Save -and -not $Id) { Write-Host "-Save needs -Id <ID> (run the search first, then pick the ID from the list)." -ForegroundColor Red; exit 1 }

$start = $null; $end = $null
if ($Since) { $start = [datetime]::ParseExact($Since, "yyyy-MM-dd", $inv) } elseif ($Days -gt 0) { $start = (Get-Date).Date.AddDays(-($Days - 1)) }
if ($Until) { $end = [datetime]::ParseExact($Until, "yyyy-MM-dd", $inv).AddDays(1) }
$s_from = if ($Sender) { $Sender.ToLower().Trim() } else { "" }

$ns = (New-Object -ComObject Outlook.Application).GetNamespace("MAPI")
$acc = $null; foreach ($a in $ns.Accounts) { if ($a.SmtpAddress -ieq $TargetEmail) { $acc = $a; break } }
if (-not $acc) { Write-Host "Account $TargetEmail not found." -ForegroundColor Red; exit 1 }
$store = $acc.DeliveryStore   # only this account's store is touched, never the other profiles' data files

function Get-MailFolders ($f, $depth) {
    $out = @()
    try { if ($f.DefaultItemType -eq 0) { $out += $f } } catch { }
    if ($depth -lt 5) { foreach ($c in $f.Folders) { if ($c.Name -ne "Search Folders") { $out += Get-MailFolders $c ($depth + 1) } } }
    return $out
}
$defaults = [ordered]@{ inbox = 6; sent = 5; junk = 23; deleted = 3 }
$skipIds = @(); foreach ($n in 16, 4) { try { $skipIds += $store.GetDefaultFolder($n).EntryID } catch { } }   # Drafts and Outbox are not "received mail"
$sentPath = $store.GetDefaultFolder(5).FolderPath

$targets = @(); $want = $Folder.ToLower().Trim()
if ($defaults.Contains($want)) { $targets = @(Get-MailFolders $store.GetDefaultFolder($defaults[$want]) 0) }
else {
    $all = @(); foreach ($c in $store.GetRootFolder().Folders) { if ($c.Name -ne "Search Folders") { $all += Get-MailFolders $c 0 } }
    if ($want -eq "all") {
        foreach ($k in $defaults.Keys) { $targets += @(Get-MailFolders $store.GetDefaultFolder($defaults[$k]) 0) }   # the main folders first
        foreach ($f in $all) { if ($targets.EntryID -notcontains $f.EntryID) { $targets += $f } }
    } else { $targets = @($all | Where-Object { $_.FolderPath.ToLower().Contains($want) }) }
}
$targets = @($targets | Where-Object { $skipIds -notcontains $_.EntryID })
if ($targets.Count -eq 0) { Write-Host "No mail folder matches '$Folder'." -ForegroundColor Red; exit 1 }

function Get-Smtp ($m) {
    try { if ($m.SenderEmailType -eq "EX") { $x = $m.Sender.GetExchangeUser(); if ($x) { return $x.PrimarySmtpAddress } } } catch { }
    $a = [string]$m.SenderEmailAddress
    if (-not $a -or $a.StartsWith("/")) { $n = [string]$m.SenderName; if ($n) { return $n } else { return "(no sender, a draft?)" } }   # internal Exchange addresses look like /o=...; show the name instead
    return $a
}
function Get-MailId ($m) {
    $k = ""; try { $k = $m.PropertyAccessor.GetProperty($PR_INTERNET_MESSAGE_ID) } catch { }
    if (-not $k) { $k = $m.EntryID }
    $sha = [System.Security.Cryptography.SHA1]::Create()
    "S" + (($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($k)) | Select-Object -First 4 | ForEach-Object { $_.ToString("X2") }) -join "")
}
# Subject and body words are searched inside Outlook (fast); sender and dates are checked here (Restrict misreads day-first dates)
function Esc ($s) { $s.Replace("'", "''") }
$dasl = @()
if ($Subject) { $dasl += '"urn:schemas:httpmail:subject" like ''%' + (Esc $Subject.Trim()) + '%''' }
if ($Text)    { $dasl += '"urn:schemas:httpmail:textdescription" like ''%' + (Esc $Text.Trim()) + '%''' }

$rows = @(); $total = 0; $notes = @(); $seen = @{}
foreach ($f in $targets) {
    $isSent = $f.FolderPath.StartsWith($sentPath, [StringComparison]::OrdinalIgnoreCase)
    $items = $f.Items
    if ($dasl.Count) { $items = $items.Restrict('@SQL=' + ($dasl -join ' AND ')) }
    $items.Sort($(if ($isSent) { "[SentOn]" } else { "[ReceivedTime]" }), $true)
    $n = $items.Count; $limit = [Math]::Min($n, $Scan)
    if ($n -gt $Scan) { $notes += "$($f.Name): only the newest $Scan of $n item(s) were examined (raise -Scan or add -Since)" }
    for ($i = 1; $i -le $limit; $i++) {
        $m = $items.Item($i); if ($m.Class -ne 43) { continue }
        $rt = if ($isSent) { [datetime]$m.SentOn } else { [datetime]$m.ReceivedTime }
        if ($end -and $rt -ge $end) { continue }
        if ($start -and $rt -lt $start) { break }                      # newest first, so everything after this is older
        $addr = (Get-Smtp $m).ToLower()
        if ($s_from -and -not ($addr.Contains($s_from) -or ([string]$m.SenderName).ToLower().Contains($s_from))) { continue }
        $mid = Get-MailId $m
        if ($seen.ContainsKey($mid)) { continue }                      # the same message in two folders counts once
        $seen[$mid] = $true; $total++
        if ($Id -and ($Id -notcontains $mid)) { continue }
        if (-not $Id -and $rows.Count -ge $Max * 3) { continue }       # enough candidates; keep counting only
        $rows += [PSCustomObject]@{ Id = $mid; Mail = $m; Folder = $f.Name; From = $addr; Received = $rt; Subject = ($m.Subject -replace '[\r\n]+', ' ').Trim() }
    }
}
$rows = @($rows | Sort-Object Received -Descending)

$desc = @(); if ($Sender) { $desc += "sender '$Sender'" }; if ($Subject) { $desc += "subject '$Subject'" }; if ($Text) { $desc += "body '$Text'" }
if ($start) { $desc += "from " + $start.ToString("yyyy-MM-dd") }; if ($end) { $desc += "until " + $end.AddDays(-1).ToString("yyyy-MM-dd") }
Write-Host ("Searched {0} folder(s) [{1}]: {2} -> {3} email(s)" -f $targets.Count, $Folder, $(if ($desc.Count) { $desc -join ' + ' } else { 'no filter' }), $total) -ForegroundColor Cyan
foreach ($t in $notes) { Write-Host "  note: $t" -ForegroundColor Yellow }
if ($rows.Count -eq 0) {
    if ($Id) { Write-Host "ID $($Id -join ', ') not found with these filters (use the same filters as the search that showed it)." -ForegroundColor Yellow }
    else { Write-Host "Nothing found." -ForegroundColor Yellow }
    exit 0
}

if (-not $Save) {
    $show = @($rows | Select-Object -First $Max)
    Write-Host "Newest first ($($show.Count) of $total):"
    $show | ForEach-Object {
        $s = if ($_.Subject.Length -gt 55) { $_.Subject.Substring(0, 55) + "..." } else { $_.Subject }
        "  {0} | {1:yyyy-MM-dd HH:mm} | {2} | {3} | {4}" -f $_.Id, $_.Received, $_.Folder, $_.From, $s
    }
    if ($total -gt $show.Count) { Write-Host "  ... narrow it with -Sender / -Subject / -Since / -Until to see the rest." }
    Write-Host "To bring one into the local files: same filters + -Id <ID> -Save"
    exit 0
}

foreach ($want in $Id) { if (-not ($rows | Where-Object { $_.Id -eq $want })) { Write-Host "ID $want not found with these filters." -ForegroundColor Yellow } }

$DataDir = Join-Path $PSScriptRoot "data"; $EmailsDir = Join-Path $PSScriptRoot "emails"
if (-not (Test-Path $DataDir)) { New-Item -ItemType Directory -Path $DataDir | Out-Null }
if (-not (Test-Path $EmailsDir)) { New-Item -ItemType Directory -Path $EmailsDir | Out-Null }
$MasterCsv = Join-Path $DataDir ("master_" + (Get-Date).ToString("yyyy_MM") + ".csv")
$DefaultCols = "ShortID,EntryID,InternetMessageID,ThreadID,ReceivedTime,SenderName,SenderEmail,Subject,IsRead_Outlook,IsReplied_Outlook,BodyFile,Triage_Status,Category,Summary,Attachment_Count,Attachment_Files,Has_Attachment".Split(",")

# Every month's master CSV, to skip emails already recorded (matched by the stable message ID)
$known = @{}
foreach ($f in @(Get-ChildItem -Path $DataDir -Filter "master_*.csv" -ErrorAction SilentlyContinue)) {
    foreach ($r in @(Import-Csv -Path $f.FullName -Encoding $Utf8Enc)) { if ($r.InternetMessageID) { $known[$r.InternetMessageID] = $r } }
}
$Master = @(); if (Test-Path $MasterCsv) { $Master = @(Import-Csv -Path $MasterCsv -Encoding $Utf8Enc) }
$cols = if ($Master.Count) { $Master[0].PSObject.Properties.Name } else { $DefaultCols }

$saved = 0
foreach ($h in $rows) {
    $m = $h.Mail
    $mid = ""; try { $mid = $m.PropertyAccessor.GetProperty($PR_INTERNET_MESSAGE_ID) } catch { }
    if (-not $mid) { $mid = $m.EntryID }
    if ($known.ContainsKey($mid)) {
        $k = $known[$mid]
        Write-Host "  $($h.Id): already recorded as $($k.ShortID) (status $($k.Triage_Status), file $($k.BodyFile)). Change its Triage_Status in the CSV if you want it triaged again." -ForegroundColor Yellow
        continue
    }
    $sname = ($m.SenderName -replace '[\r\n]+', ' ').Trim(); $tid = $m.ConversationID
    $rt = New-Object DateTime ($h.Received.Ticks - ($h.Received.Ticks % [TimeSpan]::TicksPerSecond))
    $sid = "MSG_" + $rt.ToString("yyyyMMdd_HHmmss") + "_" + [guid]::NewGuid().ToString("N").Substring(0, 4)
    $rel = "emails/$sid.md"
    $attNames = @(Get-AttachmentNames $m)
    $attBlock = if ($attNames.Count) { "`r`n## Attachments (not saved; open the email in Outlook)`r`n" + (($attNames | ForEach-Object { "- $_" }) -join "`r`n") + "`r`n" } else { "" }
    $md = @"
# Email: $($h.Subject)
- **ShortID:** $sid
- **ThreadID:** $tid
- **From:** $sname <$($h.From)>
- **Received:** $($rt.ToString("yyyy-MM-dd HH:mm:ss"))
- **Found by search:** yes (Outlook folder: $($h.Folder); the email was not moved)
$attBlock

---

$($m.Body)
"@
    Set-Content -Path (Join-Path $PSScriptRoot $rel) -Value $md -Encoding $Utf8Enc

    $vals = @{ ShortID = $sid; EntryID = $m.EntryID; InternetMessageID = $mid; ThreadID = $tid
               ReceivedTime = $rt.ToString("yyyy-MM-dd HH:mm:ss"); SenderName = $sname; SenderEmail = $h.From; Subject = $h.Subject
               IsRead_Outlook = [string](-not $m.UnRead); IsReplied_Outlook = "False"; BodyFile = $rel
               Triage_Status = "New"; Category = "Unclassified"; Summary = ""
               Attachment_Count = [string]$attNames.Count; Attachment_Files = ($attNames -join ";")
               Has_Attachment = if ($attNames.Count) { "Yes" } else { "No" } }
    $row = [ordered]@{}; foreach ($c in $cols) { $row[$c] = if ($vals.ContainsKey($c)) { $vals[$c] } else { "" } }
    $Master += [PSCustomObject]$row; $known[$mid] = [PSCustomObject]$row; $saved++
    $Master | Export-Csv -Path $MasterCsv -NoTypeInformation -Encoding $Utf8Enc   # saved per email, so a later failure cannot lose this record
    Write-Host "  [SAVED] $($h.Id) -> $sid (Triage_Status = New, file $rel)" -ForegroundColor Green
}
Write-Host "Done. $saved saved locally; Outlook was not changed. Run sync_inbox.ps1 to refresh active_context.md, then triage." -ForegroundColor Green
