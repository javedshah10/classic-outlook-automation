<#
.SYNOPSIS
    Loop 3: Reads data/reply_queue.csv for rows where Reply_Status == "Not_Replied".
    Fetches the original email in Classic Outlook via EntryID, creates a native
    thread reply (.Reply()), injects the LLM's Markdown draft, and either saves
    to Outlook Drafts or sends immediately.
#>

$ErrorActionPreference = "Stop"

# --- YOUR SETTINGS ---
# $false (default) = every reply is saved to Outlook Drafts (Save_Draft), even rows marked Auto_Send.
# Set to $true only after testing, to let rows marked Auto_Send be sent immediately.
$AllowAutoSend = $false

# Must match $TargetEmail in sync_inbox.ps1 (set_email.ps1 updates all scripts)
$TargetEmail = "your.name@example.com"

$BaseDir    = $PSScriptRoot
$DataDir    = Join-Path $BaseDir "data"
$MonthStamp = (Get-Date).ToString("yyyy_MM")
$MasterCsv  = Join-Path $DataDir "master_$MonthStamp.csv"
$ReplyCsv   = Join-Path $DataDir "reply_queue.csv"

if (-not (Test-Path $ReplyCsv)) {
    Write-Host "No reply_queue.csv found. Nothing to dispatch." -ForegroundColor Yellow
    exit
}

$Queue = @(Import-Csv -Path $ReplyCsv -Encoding UTF8)
$PendingReplies = @($Queue | Where-Object { $_.Reply_Status -eq "Not_Replied" })

if ($PendingReplies.Count -eq 0) {
    Write-Host "No rows with Reply_Status == 'Not_Replied' in reply_queue.csv." -ForegroundColor Green
    exit
}

Write-Host "Connecting to Classic Outlook to process $($PendingReplies.Count) pending reply(s)..." -ForegroundColor Cyan
$Outlook   = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")

$MasterData = @()
if (Test-Path $MasterCsv) {
    $MasterData = @(Import-Csv -Path $MasterCsv -Encoding UTF8)
}

# ShortID -> message ID from every month's master CSV. The stored EntryID goes stale when Outlook moves an email
# (Inbox <-> Junk, a rescue); the message ID never changes, so it is the fallback for finding the original.
$PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
$MyAccount = $null; foreach ($a in $Namespace.Accounts) { if ($a.SmtpAddress -ieq $TargetEmail) { $MyAccount = $a; break } }
if (-not $MyAccount) { Write-Host "Account $TargetEmail not found in Outlook." -ForegroundColor Red; exit 1 }
$MyStore = $MyAccount.DeliveryStore
$MessageIdByShort = @{}
foreach ($f in @(Get-ChildItem -Path $DataDir -Filter "master_*.csv" -ErrorAction SilentlyContinue)) {
    foreach ($r in @(Import-Csv -Path $f.FullName -Encoding UTF8)) { if ($r.InternetMessageID) { $MessageIdByShort[$r.ShortID] = $r.InternetMessageID } }
}

# Only this account's own mailbox is searched. Walking every store in the profile touches old or missing data files,
# and Outlook then pops up an error for each one.
function Test-InJunk ($mail) {
    try { return ($mail.Parent.EntryID -eq $MyStore.GetDefaultFolder(23).EntryID) } catch { return $false }
}

function Find-OriginalMail ($row) {
    try { $m = $Namespace.GetItemFromID($row.EntryID); if ($m) { return $m } } catch { }
    $mid = $MessageIdByShort[$row.ShortID]
    if (-not $mid) { return $null }
    $q = '@SQL="' + $PR_INTERNET_MESSAGE_ID + '" = ''' + ($mid -replace "'", "''") + ''''
    foreach ($fid in 6, 23) {                           # Inbox, then Junk
        try {
            $hit = $MyStore.GetDefaultFolder($fid).Items.Restrict($q)
            if ($hit.Count -ge 1) { return $hit.Item(1) }
        } catch { }
    }
    return $null
}

foreach ($row in $Queue) {
    if ($row.Reply_Status -ne "Not_Replied") { continue }

    $draftPath = Join-Path $BaseDir $row.DraftFile
    if (-not (Test-Path $draftPath)) {
        Write-Warning "Draft file missing for $($row.ShortID): $draftPath"
        $row.Reply_Status = "Error_Missing_Draft"
        continue
    }

    try {
        # 1. Locate the original email (by EntryID, or by its message ID when the EntryID is stale)
        $originalMail = Find-OriginalMail $row
        if (-not $originalMail) { throw "original email not found (deleted, or the ShortID has no message ID in any master CSV)" }
        if (Test-InJunk $originalMail) { throw "the original email is in the Junk folder, and Outlook cannot reply from there. Move it to the Inbox first (move_from_junk.ps1), then run this again" }
        if ($originalMail.EntryID -ne $row.EntryID) {
            Write-Host " [ID refreshed] $($row.ShortID): stored EntryID was stale" -ForegroundColor DarkGray
            $row.EntryID = $originalMail.EntryID
            foreach ($mRow in $MasterData) { if ($mRow.ShortID -eq $row.ShortID) { $mRow.EntryID = $originalMail.EntryID } }
        }

        # 2. Create native Outlook reply (preserves ConversationID, recipients, & thread history)
        $replyMail = $originalMail.Reply()

        # 3. Read LLM draft text and prepend above the original email history
        $draftText = (Get-Content -Path $draftPath -Raw -Encoding UTF8).Trim()
        if ($replyMail.BodyFormat -eq 2) {
            # HTML thread: insert the reply at the top of the HTML so the signature, images and quoted thread keep their formatting.
            # (Assigning .Body would flatten the whole message to plain text.)
            $enc = [System.Net.WebUtility]::HtmlEncode($draftText) -replace "`r?`n", "<br>"
            $block = "<div dir=`"auto`" style=`"font-family:Calibri,Arial,sans-serif;font-size:11pt`">$enc</div><br>"
            $html = $replyMail.HTMLBody
            $mt = [regex]::Match($html, '(?is)<body[^>]*>')
            if ($mt.Success) { $replyMail.HTMLBody = $html.Insert($mt.Index + $mt.Length, $block) } else { $replyMail.HTMLBody = $block + $html }
        } else {
            $replyMail.Body = $draftText + "`r`n`r`n" + $replyMail.Body
        }

        # 4. Dispatch according to Dispatch_Mode
        $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        if ($AllowAutoSend -and $row.Dispatch_Mode -eq "Auto_Send") {
            $replyMail.Send()
            $row.Reply_Status  = "Sent_In_Outlook"
            $newMasterStatus   = "Replied"
            $newOutlookReplied = "True"
            Write-Host " [SENT] $($row.ShortID) -> $($row.ToEmail) ($($row.Subject))" -ForegroundColor Green
        } else {
            # Default safe behavior: Save into Outlook Drafts folder for quick human review
            $replyMail.Save()
            $row.Reply_Status  = "Drafted_In_Outlook"
            $newMasterStatus   = "Drafted"
            $newOutlookReplied = $null
            Write-Host " [SAVED TO DRAFTS] $($row.ShortID) -> $($row.ToEmail) ($($row.Subject))" -ForegroundColor Cyan
        }

        $row.Dispatched_At = $timestamp

        # 5. Sync status back to master_YYYY_MM.csv
        foreach ($mRow in $MasterData) {
            if ($mRow.ShortID -eq $row.ShortID -or $mRow.EntryID -eq $row.EntryID) {
                $mRow.Triage_Status = $newMasterStatus
                if ($newOutlookReplied) { $mRow.IsReplied_Outlook = $newOutlookReplied }
            }
        }

    } catch {
        Write-Warning "Failed to process reply for $($row.ShortID): $($_.Exception.Message)"
        $row.Reply_Status = "Error_COM_Lookup"
    }
}

# Save updated CSVs
$Queue | Export-Csv -Path $ReplyCsv -NoTypeInformation -Encoding UTF8
if ($MasterData.Count -gt 0) {
    $MasterData | Export-Csv -Path $MasterCsv -NoTypeInformation -Encoding UTF8
}

Write-Host "Dispatch loop complete!" -ForegroundColor Green