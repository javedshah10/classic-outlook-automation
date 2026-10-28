<#
.SYNOPSIS
    Loop 1: Connects to Classic Outlook, syncs Read/Replied status for existing rows,
    fetches unread emails, writes clean .md bodies, appends metadata to the monthly
    master CSV, and generates active_context.md for the LLM.
#>

$ErrorActionPreference = "Stop"
# UTF-8 with BOM on both PowerShell 5.1 ("UTF8") and 7 ("utf8BOM")
$Utf8Enc = if ($PSVersionTable.PSVersion.Major -ge 6) { "utf8BOM" } else { "UTF8" }

# --- 0. YOUR SETTINGS: enter your own Outlook email address here ---
$TargetEmail = "your.name@example.com"

# Body content: $true = keep only the newest message (strip quoted replies / thread history)
#               $false = keep the complete email thread
$StripQuotedReplies = $false

# Rescan window: re-check mail this many days before the watermark (read or unread);
# already-seen IDs are skipped, so overlap is safe. shortcut: mail later than this is missed, raise it if needed.
$LookbackDays = 2


function Get-FreshBody([string]$text) {
    if (-not $text) { return "" }
    $markers = @('^\s*-{2,}\s*Original Message\s*-{2,}','^\s*_{5,}\s*$','^\s*From:\s.+','^\s*On\s.+wrote:\s*$','^\s*>')
    $fresh = @(); $seen = $false
    foreach ($line in ($text -split "`r?`n")) {
        if ($seen) { $hit=$false; foreach ($m in $markers) { if ($line -match $m) { $hit=$true; break } }; if ($hit) { break } }
        $fresh += $line; if ($line.Trim()) { $seen = $true }
    }
    ($fresh -join "`r`n").Trim()
}

# --- 1. Setup Paths & Directories ---
$BaseDir    = $PSScriptRoot
$DataDir    = Join-Path $BaseDir "data"
$EmailsDir  = Join-Path $BaseDir "emails"
$DraftsDir  = Join-Path $BaseDir "drafts"
$MonthStamp = (Get-Date).ToString("yyyy_MM")
$MasterCsv  = Join-Path $DataDir "master_$MonthStamp.csv"
$ReplyCsv   = Join-Path $DataDir "reply_queue.csv"
$ContextMd  = Join-Path $BaseDir "active_context.md"

foreach ($dir in @($DataDir, $EmailsDir, $DraftsDir)) {
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
}

# Initialize reply_queue.csv if it doesn't exist
if (-not (Test-Path $ReplyCsv)) {
    "ShortID,EntryID,ThreadID,ToEmail,Subject,DraftFile,Dispatch_Mode,Reply_Status,Dispatched_At" |
        Set-Content -Path $ReplyCsv -Encoding $Utf8Enc
}

# --- 2. Connect to Classic Outlook COM Object ---
Write-Host "[1/5] Connecting to Classic Outlook..." -ForegroundColor Cyan
$Outlook   = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")
# Find YOUR account by its email address (never the default account)
$Account = $null
foreach ($acc in $Namespace.Accounts) {
    if ($acc.SmtpAddress -ieq $TargetEmail) { $Account = $acc; break }
}
if (-not $Account) {
    Write-Host "Account $TargetEmail not found. Accounts in Outlook:" -ForegroundColor Red
    foreach ($acc in $Namespace.Accounts) { Write-Host " - $($acc.SmtpAddress)" }
    exit 1
}
$Inbox     = $Account.DeliveryStore.GetDefaultFolder(6) # 6 = olFolderInbox of that account

# MAPI Property Tags
$PR_LAST_VERB_EXECUTED  = "http://schemas.microsoft.com/mapi/proptag/0x10810003"
$PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
$PR_SMTP_ADDRESS        = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
$PR_ATTACH_HIDDEN       = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
$PR_ATTACH_CID          = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"

function Get-OutlookFlags ($MailItem) {
    $isRead = -not $MailItem.UnRead
    $isReplied = $false
    try {
        $lastVerb = $MailItem.PropertyAccessor.GetProperty($PR_LAST_VERB_EXECUTED)
        # 102 = Replied to Sender, 103 = Replied to All
        if ($lastVerb -eq 102 -or $lastVerb -eq 103) { $isReplied = $true }
    } catch { }
    return @{ IsRead = $isRead; IsReplied = $isReplied }
}

function Get-SenderSmtp ($MailItem) {
    try {
        if ($MailItem.SenderEmailType -eq "EX") {
            $exUser = $MailItem.Sender.GetExchangeUser()
            if ($exUser) { return $exUser.PrimarySmtpAddress }
            return $MailItem.PropertyAccessor.GetProperty($PR_SMTP_ADDRESS)
        }
        return $MailItem.SenderEmailAddress
    } catch {
        return $MailItem.SenderEmailAddress
    }
}

# --- 3. Load Master CSV & Sync Existing Pending Rows ---
$MasterData = @()
$ExistingIds = @{}

if (Test-Path $MasterCsv) {
    Write-Host "[2/5] Loading existing master CSV and syncing Outlook flags..." -ForegroundColor Cyan
    $MasterData = @(Import-Csv -Path $MasterCsv -Encoding $Utf8Enc)
    foreach ($row in $MasterData) {
        # Older CSVs have no attachment columns; add them so every row has the same shape
        if ($null -eq $row.PSObject.Properties['Attachment_Count']) { $row | Add-Member -NotePropertyName Attachment_Count -NotePropertyValue "" }
        if ($null -eq $row.PSObject.Properties['Attachment_Files']) { $row | Add-Member -NotePropertyName Attachment_Files -NotePropertyValue "" }
        if ($null -eq $row.PSObject.Properties['Has_Attachment']) { $row | Add-Member -NotePropertyName Has_Attachment -NotePropertyValue $(if ([int]("0" + $row.Attachment_Count) -gt 0) { "Yes" } else { "" }) }
        if ($row.EntryID) { $ExistingIds[$row.EntryID] = $true }
        if ($row.InternetMessageID) { $ExistingIds[$row.InternetMessageID] = $true }

        # Sync live Outlook flags for items that aren't closed yet
        if ($row.Triage_Status -ne "Closed" -and $row.Triage_Status -ne "Ignored") {
            try {
                $item = $Namespace.GetItemFromID($row.EntryID)
                if ($item) {
                    $flags = Get-OutlookFlags -MailItem $item
                    $row.IsRead_Outlook    = [string]$flags.IsRead
                    $row.IsReplied_Outlook = [string]$flags.IsReplied
                    if ($flags.IsReplied -and $row.Triage_Status -eq "Pending_Action") {
                        $row.Triage_Status = "Replied"
                    }
                }
            } catch {
                # Item may have been deleted or moved to another PST
            }
        }
    }
} else {
    Write-Host "[2/5] Creating new monthly master CSV: $MasterCsv" -ForegroundColor Cyan
}

# Earlier months' CSVs: the rescan window can reach back across a month change, so their IDs count as already recorded too
foreach ($oldCsv in @(Get-ChildItem -Path $DataDir -Filter "master_*.csv" -ErrorAction SilentlyContinue | Where-Object { $_.FullName -ne $MasterCsv })) {
    foreach ($oldRow in @(Import-Csv -Path $oldCsv.FullName -Encoding $Utf8Enc)) {
        if ($oldRow.EntryID) { $ExistingIds[$oldRow.EntryID] = $true }
        if ($oldRow.InternetMessageID) { $ExistingIds[$oldRow.InternetMessageID] = $true }
    }
}

# --- 4. Fetch Unread Emails from Inbox ---
Write-Host "[3/5] Scanning Inbox for unread emails..." -ForegroundColor Cyan

# Watermark: load first so we can filter in Outlook AND break in the loop
$WatermarkFile = Join-Path $DataDir "last_scan.json"
$Watermark = $null
if (Test-Path $WatermarkFile) {
    try {
        $rawWm = (Get-Content $WatermarkFile -Raw | ConvertFrom-Json).LastReceivedTime
        $Watermark = [datetime]::ParseExact($rawWm, "yyyy-MM-dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture)
    } catch { $Watermark = $null }
}

# Sort base collection first, then Restrict
$AllItems = $Inbox.Items
$AllItems.Sort("[ReceivedTime]", $true)   # newest first

if ($Watermark) {
    # No date filter in Outlook: Restrict reads 10/11/2026 as 10 November on some PCs, which would hide mail.
    # Items are sorted newest first and the loop below stops at the first item older than the window.
    $WindowStart = $Watermark.AddDays(-$LookbackDays)
    $sFilter = $null
    Write-Host "      Watermark $($Watermark.ToString('yyyy-MM-dd HH:mm:ss')); rescanning read + unread since $($WindowStart.ToString('yyyy-MM-dd HH:mm'))" -ForegroundColor DarkGray
} else {
    $WindowStart = $null
    $sFilter = "[UnRead] = True"
    Write-Host "      No watermark found: full unread scan" -ForegroundColor DarkGray
}

# Diagnostics: a cold or half-loaded Outlook can return nothing; log the inbox state every run
$InboxCount = $AllItems.Count
$InboxNewest = $null
if ($InboxCount -gt 0) { $InboxNewest = [datetime]$AllItems.Item(1).ReceivedTime }
Write-Host "      Inbox holds $InboxCount item(s); newest received: $InboxNewest" -ForegroundColor DarkGray

if ($sFilter) { $UnreadItems = $AllItems.Restrict($sFilter) } else { $UnreadItems = $AllItems }   # not an if-expression: that would unpack the COM collection
$UnreadItems.Sort("[ReceivedTime]", $true)
if ($Watermark -and $InboxCount -eq 0) {   # empty Inbox = Outlook probably still loading (a deleted newest email must not trigger this)
    Write-Warning "Inbox shows 0 items (Outlook may still be loading). Waiting 15s and retrying once."
    Start-Sleep -Seconds 15
    $AllItems = $Inbox.Items
    $AllItems.Sort("[ReceivedTime]", $true)
    if ($sFilter) { $UnreadItems = $AllItems.Restrict($sFilter) } else { $UnreadItems = $AllItems }   # not an if-expression: that would unpack the COM collection
    $UnreadItems.Sort("[ReceivedTime]", $true)
    Write-Host "      After retry: inbox holds $($AllItems.Count) item(s), scanning $($UnreadItems.Count) item(s), newest first" -ForegroundColor DarkGray
}

$NewCount = 0
$NewestSeen = $null
$NewestSubject = ""
$Examined = 0
$StoppedAtWatermark = $false

# 1-based index loop avoids PowerShell pre-enumerating the COM collection
$ItemCount = $UnreadItems.Count
Write-Host "      Scanning up to $ItemCount item(s), newest first" -ForegroundColor DarkGray
for ($i = 1; $i -le $ItemCount; $i++) {
    $mail = $UnreadItems.Item($i)
    if ($null -eq $mail -or $mail.Class -ne 43) { continue }   # olMail only

    $Examined++
    $rt0 = [datetime]$mail.ReceivedTime
    $rt = New-Object DateTime ($rt0.Ticks - ($rt0.Ticks % [TimeSpan]::TicksPerSecond))   # whole seconds

    if (-not $NewestSeen -or $rt -gt $NewestSeen) {
        $NewestSeen = $rt
        $NewestSubject = ($mail.Subject -replace '[\r\n]+', ' ').Trim()
    }

    if ($WindowStart -and $rt -lt $WindowStart) { $StoppedAtWatermark = $true; break }

    $entryId = $mail.EntryID
    $internetId = ""
    try { $internetId = $mail.PropertyAccessor.GetProperty($PR_INTERNET_MESSAGE_ID) } catch { }
    if (-not $internetId) { $internetId = $entryId }

    if ($ExistingIds.ContainsKey($entryId) -or $ExistingIds.ContainsKey($internetId)) { continue }

    $receivedTime = $rt
    $guidSuffix   = [guid]::NewGuid().ToString("N").Substring(0, 4)
    $shortId      = "MSG_" + $receivedTime.ToString("yyyyMMdd_HHmmss") + "_" + $guidSuffix

    $senderName   = ($mail.SenderName -replace '[\r\n]+', ' ').Trim()
    $senderEmail  = (Get-SenderSmtp -MailItem $mail)
    $subject      = ($mail.Subject -replace '[\r\n]+', ' ').Trim()
    $threadId     = $mail.ConversationID
    $flags        = Get-OutlookFlags -MailItem $mail
    $relBodyPath  = "emails/$shortId.md"
    $fullBodyPath = Join-Path $BaseDir $relBodyPath

    # Attachment flag only: the files stay in Outlook and are not saved here
    $attNames = @()
    for ($a = 1; $a -le $mail.Attachments.Count; $a++) {
        $att = $mail.Attachments.Item($a)
        if ($att.Type -ne 1) { continue }   # real files only (not embedded mails or links)
        $ext = [IO.Path]::GetExtension($att.FileName).ToLower()
        $inline = $false
        try { $inline = [bool]$att.PropertyAccessor.GetProperty($PR_ATTACH_HIDDEN) } catch { }
        if (-not $inline -and $ext -match '^\.(png|jpe?g|gif)$') {
            try { if ($att.PropertyAccessor.GetProperty($PR_ATTACH_CID)) { $inline = $true } } catch { }
        }
        if ($inline) { continue }           # signature logos and inline pictures
        $attNames += $att.FileName
    }
    $attBlock = if ($attNames.Count) { "`r`n## Attachments (not saved; open the email in Outlook)`r`n" + (($attNames | ForEach-Object { "- $_" }) -join "`r`n") + "`r`n" } else { "" }

    $mdContent = @"
# Email: $subject
- **ShortID:** $shortId
- **ThreadID:** $threadId
- **From:** $senderName <$senderEmail>
- **Received:** $($receivedTime.ToString("yyyy-MM-dd HH:mm:ss"))
- **Outlook Replied Flag:** $($flags.IsReplied)
$attBlock
---

$(if ($StripQuotedReplies) { Get-FreshBody $mail.Body } else { $mail.Body })
"@
    Set-Content -Path $fullBodyPath -Value $mdContent -Encoding $Utf8Enc

    $newRow = [PSCustomObject]@{
        ShortID           = $shortId
        EntryID           = $entryId
        InternetMessageID = $internetId
        ThreadID          = $threadId
        ReceivedTime      = $receivedTime.ToString("yyyy-MM-dd HH:mm:ss")
        SenderName        = $senderName
        SenderEmail       = $senderEmail
        Subject           = $subject
        IsRead_Outlook    = [string]$flags.IsRead
        IsReplied_Outlook = [string]$flags.IsReplied
        BodyFile          = $relBodyPath
        Triage_Status     = if ($flags.IsRead) { "Ignored" } else { "New" }
        Category          = "Unclassified"
        Summary           = if ($flags.IsRead) { "Already read in Outlook when first synced; auto-ignored." } else { "" }
        Attachment_Count  = [string]$attNames.Count
        Attachment_Files  = ($attNames -join ";")   # names only
        Has_Attachment    = if ($attNames.Count) { "Yes" } else { "No" }
    }

    $MasterData += $newRow
    $ExistingIds[$entryId]    = $true
    $ExistingIds[$internetId] = $true
    $NewCount++
}

Write-Host "      Examined $Examined email(s); stopped at watermark: $StoppedAtWatermark" -ForegroundColor DarkGray

# Save watermark (newest email seen this scan; keep old one if nothing newer)
if ($NewestSeen -and (-not $Watermark -or $NewestSeen -gt $Watermark)) {
    [PSCustomObject]@{
        LastReceivedTime = $NewestSeen.ToString("yyyy-MM-dd HH:mm:ss")
        LastSubject      = $NewestSubject
        UpdatedAt        = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    } | ConvertTo-Json | Set-Content -Path $WatermarkFile -Encoding $Utf8Enc
}

# --- 5. Save Updated Master CSV ---
Write-Host "[4/5] Saving $NewCount new email(s) to $MasterCsv..." -ForegroundColor Cyan
if ($MasterData.Count -gt 0) {
    $MasterData | Export-Csv -Path $MasterCsv -NoTypeInformation -Encoding $Utf8Enc
} else {
    "ShortID,EntryID,InternetMessageID,ThreadID,ReceivedTime,SenderName,SenderEmail,Subject,IsRead_Outlook,IsReplied_Outlook,BodyFile,Triage_Status,Category,Summary,Attachment_Count,Attachment_Files,Has_Attachment" |
        Set-Content -Path $MasterCsv -Encoding $Utf8Enc
}

# --- 6. Rebuild Rolling LLM Memory (active_context.md) ---
Write-Host "[5/5] Rebuilding active_context.md for LLM quick reference..." -ForegroundColor Cyan
$NewEmails     = @($MasterData | Where-Object { $_.Triage_Status -eq "New" })
$PendingEmails = @($MasterData | Where-Object { $_.Triage_Status -eq "Pending_Action" -or $_.Triage_Status -eq "Drafted" })
$RecentReplied = @($MasterData | Where-Object { $_.IsReplied_Outlook -eq "True" -or $_.Triage_Status -eq "Replied" } | Select-Object -Last 10)

$contextLines = @()
$contextLines += "# Active Email Context & Memory"
$contextLines += "_Auto-generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Active File: ``data/master_$MonthStamp.csv``_"
$contextLines += ""
$contextLines += "## 1. Unprocessed Emails Waiting for LLM Triage (``Triage_Status == New``): $($NewEmails.Count)"
if ($NewEmails.Count -eq 0) {
    $contextLines += "- None."
} else {
    foreach ($e in $NewEmails) {
        $attFlag = if ($e.Has_Attachment -eq "Yes") { " | Attachment: YES" } else { "" }
        $contextLines += "- **$($e.ShortID)** | From: ``$($e.SenderEmail)`` | Subject: **$($e.Subject)** | Body: ``$($e.BodyFile)`` | ThreadID: ``$($e.ThreadID)``$attFlag"
    }
}

$contextLines += ""
$contextLines += "## 2. Open Pending Action / Drafted Items: $($PendingEmails.Count)"
if ($PendingEmails.Count -eq 0) {
    $contextLines += "- None."
} else {
    foreach ($p in $PendingEmails) {
        $attFlag = if ($p.Has_Attachment -eq "Yes") { " | Attachment: YES" } else { "" }
        $contextLines += "- **$($p.ShortID)** [$($p.Triage_Status)] | ``$($p.SenderEmail)`` | **$($p.Subject)** | Summary: $($p.Summary) | ThreadID: ``$($p.ThreadID)``$attFlag"
    }
}

$contextLines += ""
$contextLines += "## 3. Recently Replied Threads (Last 10 Reference)"
if ($RecentReplied.Count -eq 0) {
    $contextLines += "- None."
} else {
    foreach ($r in $RecentReplied) {
        $contextLines += "- **$($r.ShortID)** | ``$($r.SenderEmail)`` | **$($r.Subject)** | ThreadID: ``$($r.ThreadID)``"
    }
}

Set-Content -Path $ContextMd -Value ($contextLines -join "`r`n") -Encoding $Utf8Enc
Write-Host "Done! Ready for Claude/LLM triage." -ForegroundColor Green
