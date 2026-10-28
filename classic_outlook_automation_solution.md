# Classic Outlook + LLM Email Automation Solution
**Zero-Admin, Flat-File Architecture (PowerShell COM + CSV/Markdown + Claude)**

> **Read this first - what changed since this document was written.** The `.ps1` files in the folder are the source of truth; the script listings below are an older design snapshot.
> - `sync_inbox.ps1` keeps a watermark (`data/last_scan.json`) and rescans **read and unread** mail from `$LookbackDays` (2) before it; IDs already in the CSV are skipped, so repeat runs add no duplicates. Mail already read in Outlook is saved as `Ignored`. It logs the inbox size and retries once if Outlook looks half-loaded.
> - `delete_marked.ps1` deletes mail you mark `Triage_Status = Delete` (dry run first, then `-Apply`): the email goes to Outlook Deleted Items, the body file moves to `emails/deleted/`, and the row stays as `Closed`.
> - `compose_mail.ps1` opens a new email in an Outlook compose window (To/Cc/subject/text filled in, your signature added by Outlook); you press Send. It never sends.
> - `delete_mail.ps1` deletes Inbox or Junk mail by sender and/or subject (dry run first, `-Max` guard, Deleted Items only) and closes matching CSV rows.
> - `inbox_stats.ps1` gives read-only stats for unread Inbox mail (counts, senders, and how many are not yet in the CSV).
> - `junk_stats.ps1` and `move_from_junk.ps1` give read-only stats for the Outlook Junk folder and move chosen emails back to the Inbox (dry run first, then `-Apply`); a moved email is recorded in the CSV as `New` with its `.md`. Other junk mail is never synced or saved.
> - The CSV has `Has_Attachment` (`Yes`/`No`), `Attachment_Count` and `Attachment_Files`; `active_context.md` marks emails with attachments as `Attachment: YES` so they can be checked.
> - `dispatch_replies.ps1` puts the reply text at the top of the HTML reply, so Outlook's default reply signature, images and the quoted thread keep their formatting (plain-text threads get plain text).
> - `dispatch_replies.ps1` finds the original email by its permanent message ID when the stored `EntryID` has gone stale (it searches your own Inbox and Junk only), and tells you clearly if the original is in Junk or deleted.
> - `CLAUDE_PROMPT.md` holds the current rules for the Claude agent (late-arrival re-sync, FAQ check, housekeeping, known limits).

---

## 1. Executive Summary & Why This Works
Most students and corporate professionals cannot connect AI agents directly to Microsoft 365 because Microsoft Entra ID (formerly Azure AD) and Graph API require **IT Administrator consent**. 

However, if the user has **Classic Outlook** installed and logged in on their Windows machine, Windows exposes a local COM interface (`Outlook.Application`). By separating the workflow into **Deterministic PowerShell Loops** (reading and dispatching via local Outlook) and an **Agentic LLM Loop** (reading, triaging, and drafting via local CSV and `.md` files), we completely bypass API blocks without needing a database.

### Core Architectural Principles
1. **No Email Bodies in CSVs:** Multi-line email bodies, commas, and quotes will corrupt CSV files. CSVs are strictly for **single-line metadata and state flags**. Actual email bodies and drafted replies live in clean, individual `.md` files referenced by ID.
2. **Bi-Directional State Sync:** Before fetching new unread emails, PowerShell checks existing "Pending" rows against Outlook to see if the user manually read or replied to them inside Outlook.
3. **Token-Light Rolling Memory (`active_context.md`):** Instead of forcing the LLM to read a 1,000-row monthly CSV on every run, PowerShell generates a compact `active_context.md` summary so the LLM has instant memory at minimal token cost.
4. **Safe Outbound Dispatch:** The LLM writes a draft `.md` file and sets a flag in `reply_queue.csv`. A second deterministic PowerShell loop picks up un-replied rows, attaches them to the exact original Outlook thread via `.Reply()`, and either saves them to **Outlook Drafts** (`Save_Draft`) or sends them immediately (`Auto_Send`).

---

## 2. System Architecture Diagram

```text
================================================================================
LOOP 1: DETERMINISTIC INGESTION & SYNC (sync_inbox.ps1)
================================================================================
  ( Start )
      |
      v
[ 1. Sync Existing Pending Rows ]
  - Check master_YYYY_MM.csv rows where Triage_Status != "Closed"
  - Lookup email in Outlook via EntryID -> Check MAPI Replied/Read flags
  - If user manually replied in Outlook -> Mark IsReplied_Outlook = True
      |
      v
[ 2. Fetch Unread Emails from Classic Outlook Inbox ]
  - Compare EntryID & InternetMessageID against master_YYYY_MM.csv
  - Already in CSV? -> Skip
      |
      v
[ 3. Store Cleanly (Split Metadata & Body) ]
  - Append metadata row to data/master_YYYY_MM.csv (Triage_Status = "New")
  - Save plain-text email body to emails/<ShortID>.md
      |
      v
[ 4. Rebuild Rolling Memory ]
  - Auto-generate active_context.md (Open threads + Pending actions)

================================================================================
LOOP 2: AGENTIC TRIAGE & DRAFTING (Claude / LLM Agent)
================================================================================
      |
      v
[ 5. Read Context & New Arrivals ]
  - Read: knowledge_base.md + active_context.md
  - Filter: master_YYYY_MM.csv where Triage_Status == "New"
  - Read: Corresponding emails/<ShortID>.md
      |
      v
[ 6. Classify & Draft Response ]
  - Update master_YYYY_MM.csv (Category, Priority, Triage_Status, Summary)
  - If reply needed -> Create drafts/<ShortID>_reply.md
  - Append row to data/reply_queue.csv (Reply_Status = "Not_Replied")

================================================================================
LOOP 3: DETERMINISTIC DISPATCH (dispatch_replies.ps1)
================================================================================
      |
      v
[ 7. Poll Outbound Queue ]
  - Read data/reply_queue.csv where Reply_Status == "Not_Replied"
      |
      v
[ 8. Execute Native Outlook Thread Reply ]
  - Fetch original email via Namespace.GetItemFromID(EntryID)
  - Call $mail.Reply() (Preserves ConversationID, Subject, & Thread History)
  - Prepend text from drafts/<ShortID>_reply.md
  - If Dispatch_Mode == "Auto_Send" -> $reply.Send()
  - Else -> $reply.Save() (Appears in Outlook Drafts folder)
      |
      v
[ 9. Update Records & Memory ]
  - Update data/reply_queue.csv (Reply_Status = "Dispatched", Dispatched_At)
  - Update data/master_YYYY_MM.csv (Triage_Status = "Replied" or "Drafted")
      |
      v
   ( End )
```

---

## 3. Local Folder & File Structure
Create one folder on the local machine (e.g., `C:\OutlookAgent\`). No SQLite or external database is required.

```text
C:\OutlookAgent\
├── sync_inbox.ps1               # Loop 1: Syncs Outlook -> CSV & .md files
├── dispatch_replies.ps1         # Loop 3: Pushes LLM replies -> Outlook
├── knowledge_base.md            # User rules, tone, FAQs, & classification guide
├── active_context.md            # Auto-maintained rolling memory for the LLM
├── data\
│   ├── master_2026_10.csv       # Monthly master tracker (Metadata & Flags ONLY)
│   └── reply_queue.csv          # Outbound reply queue linking LLM to Outlook
├── emails\
│   └── MSG_20261027_001.md      # Extracted plain-text email bodies
└── drafts\
    └── MSG_20261027_001_reply.md # LLM-generated reply bodies
```

---

## 4. Data Schemas & File Templates

### 4.1 `data/master_YYYY_MM.csv`
Created and updated automatically by `sync_inbox.ps1` and the LLM.

| Column Name | Owner | Description |
| :--- | :--- | :--- |
| `ShortID` | PowerShell | Human- and filename-friendly ID (e.g., `MSG_20261027_143022_a1b2`) |
| `EntryID` | PowerShell | Outlook's internal COM ID used for instant `.GetItemFromID()` lookup |
| `InternetMessageID` | PowerShell | Global RFC 2822 Message-ID for strict deduplication |
| `ThreadID` | PowerShell | Outlook's native `ConversationID` linking all replies to the original email |
| `ReceivedTime` | PowerShell | Timestamp (`YYYY-MM-DD HH:mm:ss`) |
| `SenderName` | PowerShell | Display name of sender |
| `SenderEmail` | PowerShell | SMTP email address of sender |
| `Subject` | PowerShell | Email subject line (sanitized of newlines) |
| `IsRead_Outlook` | PowerShell | `True` or `False` synced directly from Outlook |
| `IsReplied_Outlook`| PowerShell | `True` or `False` via Outlook MAPI property `0x10810003` |
| `BodyFile` | PowerShell | Relative path to body (e.g., `emails/MSG_20261027_143022_a1b2.md`) |
| `Triage_Status` | Both | `New` (set by PS1) -> `Pending_Action`, `Drafted`, `Replied`, `Ignored`, `Closed` |
| `Category` | LLM | E.g., `Client_Inquiry`, `Billing`, `Internal`, `Newsletter` |
| `Summary` | LLM | 1-sentence summary of what the email wants |

### 4.2 `data/reply_queue.csv`
The LLM appends rows here whenever an email needs a reply. `dispatch_replies.ps1` monitors this file for `Reply_Status == "Not_Replied"`.

| Column Name | Owner | Description |
| :--- | :--- | :--- |
| `ShortID` | LLM | Matches `ShortID` from `master_YYYY_MM.csv` |
| `EntryID` | LLM | Copied from `master_YYYY_MM.csv` so PowerShell can locate the original email |
| `ThreadID` | LLM | Copied from `master_YYYY_MM.csv` |
| `ToEmail` | LLM | Recipient email address (for logging/verification) |
| `Subject` | LLM | Original subject line |
| `DraftFile` | LLM | Relative path to reply markdown (e.g., `drafts/MSG_..._reply.md`) |
| `Dispatch_Mode` | LLM | `Save_Draft` (safe default: saves in Outlook Drafts) or `Auto_Send` (sends immediately) |
| `Reply_Status` | Both | LLM sets to `Not_Replied`. PowerShell updates to `Drafted_In_Outlook` or `Sent_In_Outlook` |
| `Dispatched_At` | PowerShell | Timestamp when PowerShell pushed the reply to Outlook |

### 4.3 `knowledge_base.md` (Starter Template)
Place this in the root directory so the student can customize their own rules.

```markdown
# Email Triage & Reply Knowledge Base

## 1. My Role & Tone
- **Name:** [Student Name]
- **Role:** [Job Title / Department]
- **Tone:** Concise, warm, professional, and direct. Never use robotic phrases like "I hope this email finds you well."

## 2. Classification Rules (Categories)
- **VIP_Urgent:** Emails from `@clientdomain.com` or my manager (`boss@company.com`).
- **Pending_Action:** Any email asking a direct question, requesting a file, or scheduling a meeting.
- **FYI_Only:** Status updates where I am CC'd and no question is directed at me.
- **Noise:** Newsletters, automated notifications, marketing.

## 3. Reply & Dispatch Policy
- **When to use `Dispatch_Mode = Auto_Send`:** Only for simple scheduling confirmations or standard FAQ requests defined below.
- **When to use `Dispatch_Mode = Save_Draft`:** All client pricing questions, nuanced project updates, or escalations.

## 4. Standard FAQ Answers
- **Pricing / Rate Card Request:** Let them know our standard onboarding package starts at $2,500 and ask if they would like to book a 15-minute discovery call on Thursday.
- **Support / Bug Reports:** Acknowledge receipt, ask for a screenshot of the error, and let them know the technical team reviews tickets within 24 hours.
```

---

## 5. Script 1: `sync_inbox.ps1` (Deterministic Ingestion & State Sync)
Save this file as `sync_inbox.ps1` in your root folder. It connects to Classic Outlook, syncs existing items, fetches unread emails, writes the `.md` body files, updates the monthly CSV, and rebuilds `active_context.md`.

```powershell
<#
.SYNOPSIS
    Loop 1: Connects to Classic Outlook, syncs Read/Replied status for existing rows,
    fetches unread emails, writes clean .md bodies, appends metadata to the monthly
    master CSV, and generates active_context.md for the LLM.
#>

$ErrorActionPreference = "Stop"

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
        Set-Content -Path $ReplyCsv -Encoding UTF8
}

# --- 2. Connect to Classic Outlook COM Object ---
Write-Host "[1/5] Connecting to Classic Outlook..." -ForegroundColor Cyan
$Outlook   = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")
$Inbox     = $Namespace.GetDefaultFolder(6) # 6 = olFolderInbox

# MAPI Property Tags
$PR_LAST_VERB_EXECUTED  = "http://schemas.microsoft.com/mapi/proptag/0x10810003"
$PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
$PR_SMTP_ADDRESS        = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"

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
    $MasterData = @(Import-Csv -Path $MasterCsv -Encoding UTF8)
    foreach ($row in $MasterData) {
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

# --- 4. Fetch Unread Emails from Inbox ---
Write-Host "[3/5] Scanning Inbox for unread emails..." -ForegroundColor Cyan
$UnreadItems = $Inbox.Items.Restrict("[UnRead] = True")
$NewCount = 0

foreach ($mail in $UnreadItems) {
    # Process only MailItems (Class 43 = olMail)
    if ($mail.Class -ne 43) { continue }

    $entryId = $mail.EntryID
    $internetId = ""
    try { $internetId = $mail.PropertyAccessor.GetProperty($PR_INTERNET_MESSAGE_ID) } catch { }
    if (-not $internetId) { $internetId = $entryId }

    # Deduplication Check
    if ($ExistingIds.ContainsKey($entryId) -or $ExistingIds.ContainsKey($internetId)) {
        continue
    }

    # Generate clean ShortID for filenames
    $receivedTime = [datetime]$mail.ReceivedTime
    $guidSuffix   = [guid]::NewGuid().ToString("N").Substring(0, 4)
    $shortId      = "MSG_" + $receivedTime.ToString("yyyyMMdd_HHmmss") + "_" + $guidSuffix

    $senderName   = ($mail.SenderName -replace '[\r\n]+', ' ').Trim()
    $senderEmail  = (Get-SenderSmtp -MailItem $mail)
    $subject      = ($mail.Subject -replace '[\r\n]+', ' ').Trim()
    $threadId     = $mail.ConversationID
    $flags        = Get-OutlookFlags -MailItem $mail
    $relBodyPath  = "emails/$shortId.md"
    $fullBodyPath = Join-Path $BaseDir $relBodyPath

    # Save clean Markdown file for the email body
    $mdContent = @"
# Email: $subject
- **ShortID:** $shortId
- **ThreadID:** $threadId
- **From:** $senderName <$senderEmail>
- **Received:** $($receivedTime.ToString("yyyy-MM-dd HH:mm:ss"))
- **Outlook Replied Flag:** $($flags.IsReplied)

---

$($mail.Body)
"@
    Set-Content -Path $fullBodyPath -Value $mdContent -Encoding UTF8

    # Create Metadata Row
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
        Triage_Status     = "New"
        Category          = "Unclassified"
        Summary           = ""
    }

    $MasterData += $newRow
    $ExistingIds[$entryId]    = $true
    $ExistingIds[$internetId] = $true
    $NewCount++
}

# --- 5. Save Updated Master CSV ---
Write-Host "[4/5] Saving $NewCount new email(s) to $MasterCsv..." -ForegroundColor Cyan
if ($MasterData.Count -gt 0) {
    $MasterData | Export-Csv -Path $MasterCsv -NoTypeInformation -Encoding UTF8
} else {
    "ShortID,EntryID,InternetMessageID,ThreadID,ReceivedTime,SenderName,SenderEmail,Subject,IsRead_Outlook,IsReplied_Outlook,BodyFile,Triage_Status,Category,Summary" |
        Set-Content -Path $MasterCsv -Encoding UTF8
}

# --- 6. Rebuild Rolling LLM Memory (active_context.md) ---
Write-Host "[5/5] Rebuilding active_context.md for LLM quick reference..." -ForegroundColor Cyan
$NewEmails     = @($MasterData | Where-Object { $_.Triage_Status -eq "New" })
$PendingEmails = @($MasterData | Where-Object { $_.Triage_Status -eq "Pending_Action" -or $_.Triage_Status -eq "Drafted" })
$RecentReplied = @($MasterData | Where-Object { $_.IsReplied_Outlook -eq "True" -or $_.Triage_Status -eq "Replied" } | Select-Object -Last 10)

$contextLines = @()
$contextLines += "# Active Email Context & Memory"
$contextLines += "_Auto-generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Active File: `data/master_$MonthStamp.csv`_"
$contextLines += ""
$contextLines += "## 1. Unprocessed Emails Waiting for LLM Triage (`Triage_Status == New`): $($NewEmails.Count)"
if ($NewEmails.Count -eq 0) {
    $contextLines += "- None."
} else {
    foreach ($e in $NewEmails) {
        $contextLines += "- **$($e.ShortID)** | From: `$($e.SenderEmail)` | Subject: **$($e.Subject)** | Body: `$($e.BodyFile)` | ThreadID: `$($e.ThreadID)`"
    }
}

$contextLines += ""
$contextLines += "## 2. Open Pending Action / Drafted Items: $($PendingEmails.Count)"
if ($PendingEmails.Count -eq 0) {
    $contextLines += "- None."
} else {
    foreach ($p in $PendingEmails) {
        $contextLines += "- **$($p.ShortID)** [$($p.Triage_Status)] | `$($p.SenderEmail)` | **$($p.Subject)** | Summary: $($p.Summary) | ThreadID: `$($p.ThreadID)`"
    }
}

$contextLines += ""
$contextLines += "## 3. Recently Replied Threads (Last 10 Reference)"
if ($RecentReplied.Count -eq 0) {
    $contextLines += "- None."
} else {
    foreach ($r in $RecentReplied) {
        $contextLines += "- **$($r.ShortID)** | `$($r.SenderEmail)` | **$($r.Subject)** | ThreadID: `$($r.ThreadID)`"
    }
}

Set-Content -Path $ContextMd -Value ($contextLines -join "`r`n") -Encoding UTF8
Write-Host "Done! Ready for Claude/LLM triage." -ForegroundColor Green
```

---

## 6. Script 2: `dispatch_replies.ps1` (Deterministic Reply Loop)
Save this file as `dispatch_replies.ps1` in your root folder. After the LLM writes drafts and flags rows in `data/reply_queue.csv` as `Not_Replied`, run this script to push those replies directly into Classic Outlook threads.

```powershell
<#
.SYNOPSIS
    Loop 3: Reads data/reply_queue.csv for rows where Reply_Status == "Not_Replied".
    Fetches the original email in Classic Outlook via EntryID, creates a native
    thread reply (.Reply()), injects the LLM's Markdown draft, and either saves
    to Outlook Drafts or sends immediately.
#>

$ErrorActionPreference = "Stop"

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

foreach ($row in $Queue) {
    if ($row.Reply_Status -ne "Not_Replied") { continue }

    $draftPath = Join-Path $BaseDir $row.DraftFile
    if (-not (Test-Path $draftPath)) {
        Write-Warning "Draft file missing for $($row.ShortID): $draftPath"
        $row.Reply_Status = "Error_Missing_Draft"
        continue
    }

    try {
        # 1. Locate exact original email by EntryID
        $originalMail = $Namespace.GetItemFromID($row.EntryID)

        # 2. Create native Outlook reply (preserves ConversationID, recipients, & thread history)
        $replyMail = $originalMail.Reply()

        # 3. Read LLM draft text and prepend above the original email history
        $draftText = (Get-Content -Path $draftPath -Raw -Encoding UTF8).Trim()
        $replyMail.Body = $draftText + "`r`n`r`n" + $replyMail.Body

        # 4. Dispatch according to Dispatch_Mode
        $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        if ($row.Dispatch_Mode -eq "Auto_Send") {
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
```

---

## 7. Prompt Instructions for Claude / LLM Agent (Loop 2)
Whether your students use **Claude Code**, **Claude Desktop (with local filesystem MCP)**, or **Cursor/Windsurf**, give the LLM this exact system prompt to run Loop 2:

```markdown
You are my Email Triage & Reply Agent operating inside this folder. Follow these deterministic steps in order:

1. **Load Context & Rules:**
   - Read `knowledge_base.md` and `active_context.md`.
   - Identify the current month's master CSV listed in `active_context.md` (e.g., `data/master_YYYY_MM.csv`).

2. **Triage New Emails:**
   - For every item listed under "Unprocessed Emails Waiting for LLM Triage (`Triage_Status == New`)" in `active_context.md`, read its corresponding markdown file in `emails/<ShortID>.md`.
   - Check if its `ThreadID` matches any open or recently replied threads in `active_context.md` so you have full conversation context.
   - Update that row in `data/master_YYYY_MM.csv`:
     - Set `Category` based on `knowledge_base.md`.
     - Write a 1-sentence plain-text `Summary` (do NOT use newlines or unescaped quotes in CSV fields).
     - Set `Triage_Status` to either `Pending_Action` (if it needs a reply or task) or `Ignored` (if it is noise/FYI).

3. **Draft Replies & Queue Outbound:**
   - For every row you marked as `Pending_Action` that requires an email reply:
     - Write a clean plain-text reply in `drafts/<ShortID>_reply.md` (write ONLY the reply message text; PowerShell will automatically attach the thread history).
     - Append a new row to `data/reply_queue.csv` with:
       - `ShortID`, `EntryID`, `ThreadID`, `ToEmail` (from `SenderEmail`), `Subject`
       - `DraftFile`: `drafts/<ShortID>_reply.md`
       - `Dispatch_Mode`: `Save_Draft` (or `Auto_Send` if explicitly permitted by `knowledge_base.md`)
       - `Reply_Status`: `Not_Replied`
       - `Dispatched_At`: leave empty (`""`)

4. **Update `active_context.md`:**
   - Refresh `active_context.md` so that the newly triaged items and drafted replies are reflected under "Open Pending Action / Drafted Items".
```

---

## 8. Daily Student Workflow (3 Simple Steps)
1. **Pull New Emails (Deterministic):** Run `.\sync_inbox.ps1` in PowerShell.
2. **Triage & Draft (Agentic):** Tell Claude: *"Run the email triage and reply workflow."*
3. **Push Replies to Outlook (Deterministic):** Run `.\dispatch_replies.ps1` in PowerShell, then check the **Outlook Drafts** folder to review and send.