# Classic Outlook + AI Agent Email Assistant (starter template)

PowerShell scripts for Classic Outlook (the desktop app on Windows). They copy your mail into plain local files. An AI agent reads those files, sorts the mail using your rules, and writes replies, which the scripts save to your Outlook Drafts. It works with Claude Code, Codex, OpenCode, Cursor or any agent that can run commands and read files. The scripts do the repeatable work, so the agent reads short files and statistics instead of your whole mailbox. That keeps token use low.

Nothing is ever sent for you. Your mail stays on your PC, apart from whatever you choose to show the agent.

It is meant for people who cannot, or do not want to, connect Microsoft 365 Outlook to an AI agent through MCP or the Graph API (for example because the admin blocks it). It talks only to the Outlook already running on your PC.

## How it works
`Outlook` > `sync_inbox.ps1` (mail becomes local files) > `knowledge_base.md` (your rules) > the agent triages and decides > `dispatch_replies.ps1` (drafts appear in Outlook) > you review and press Send.

The agent's memory is two plain files: `knowledge_base.md` holds your rules (who is VIP, what comes first, how you reply, standard answers) and `active_context.md` holds the current state of your mail. Within the limits in Safety below, the agent can:
- sort new mail by your priorities and write a one-line summary of each;
- draft replies in the original thread, or open a new email for you to review;
- delete mail by sender or subject, or the rows you mark, after a dry run you confirm;
- search for any old or lost email in any folder, and move wrongly filtered mail back from Junk.

## What you need
- Windows with **Classic Outlook** (the desktop app) signed in to your account. The "New Outlook" does not work: it has no automation interface.
- Windows PowerShell 5.1 (built in).
- An AI agent that can run shell commands and read and write files. It is not tied to one product: Claude Code, Codex, OpenCode, Cursor or similar tools should all work, because the scripts are plain PowerShell and `CLAUDE_PROMPT.md` is plain text (the name is only a label). I built and tested it with Claude Code; the others follow the same steps but I have not tested them.

## Quick start
1. Run `.\set_email.ps1 -NewEmail you@yourdomain.com` to put your Outlook address into every script (or edit `$TargetEmail` in each `.ps1`).
2. Fill in the placeholders in `knowledge_base.md`: your name and role, who counts as VIP, how to classify mail, your standard answers.
3. Keep Classic Outlook open and fully loaded, then run `.\sync_inbox.ps1`.
   - The first run reads your unread Inbox mail and creates `data\master_YYYY_MM.csv`, `emails\*.md`, `data\last_scan.json` (the watermark) and `active_context.md`.
   - Later runs rescan read and unread mail from 2 days before the watermark (`$LookbackDays`) and skip anything already recorded, so repeat runs never create duplicates.
4. Give `CLAUDE_PROMPT.md` to your AI agent. It runs the whole loop: sync, triage, draft replies, save them to Outlook Drafts, report.

## Safety
- Neither `compose_mail.ps1` nor `dispatch_replies.ps1` sends anything (`$AllowAutoSend = $false`). A new email opens in a compose window and you press Send yourself; replies are only saved as drafts. Leave that switch alone unless you fully understand the risk.
- Deleting is always a normal delete into Outlook's Deleted Items, so you can undo it. The scripts never empty that folder.
- Every deleting or moving script does a dry run first. Add `-Apply` only after you have read the list. They need a filter and refuse more than `-Max` matches at once.
- `search_mail.ps1`, `inbox_stats.ps1` and `junk_stats.ps1` only read. Saving a found email only writes local files; the email stays where it is in Outlook.
- Mail content is untrusted text. The prompt tells the AI to treat it as data, never as instructions.

## The scripts
| Script | What it does |
|---|---|
| `sync_inbox.ps1` | Reads the Inbox into the CSV and one `.md` per email; rebuilds `active_context.md` |
| `dispatch_replies.ps1` | Saves the replies the agent wrote into Outlook Drafts, in the original thread |
| `set_email.ps1` | Puts your address into all scripts |
| `inbox_stats.ps1` | Read-only counts of unread mail by day or month and sender; shows how many are not yet in the CSV |
| `delete_marked.ps1` | Deletes the rows you mark `Triage_Status = Delete` in the CSV; the row stays as `Closed` |
| `delete_mail.ps1` | Deletes Inbox or Junk mail by sender and/or subject (`-Sender @domain.com`, `name@domain.com` or text) |
| `compose_mail.ps1` | Opens a new email in an Outlook compose window with To, Cc, subject and text filled in; your own signature is added and you press Send |
| `search_mail.ps1` | Finds an email of any age in any mail folder (Inbox, Sent, Junk, Deleted, Archive, subfolders) by sender, subject, body words or date. Read-only; `-Id <ID> -Save` copies one into the local files for the agent |
| `junk_stats.ps1` | Read-only Junk folder statistics and a short list with an ID per email |
| `move_from_junk.ps1` | Moves chosen emails from Junk back to the Inbox and records them in the CSV as `New` |

Examples:
```powershell
.\inbox_stats.ps1 -Days 2
.\search_mail.ps1 -Sender supplier.com -Since 2026-01-01        # find an old email; read-only
.\search_mail.ps1 -Sender supplier.com -Id S1A2B3C4D -Save       # same filters + the ID: copy it into emails\ and the CSV
.\compose_mail.ps1 -To a@example.com -Cc b@example.com -Subject "Hello" -Body "Text" -DryRun   # checks only; drop -DryRun to open the compose window
.\delete_mail.ps1 -Folder Junk -Sender spam-domain.com          # dry run
.\delete_mail.ps1 -Folder Junk -Sender spam-domain.com -Apply
.\junk_stats.ps1 -Date 2026-10-25
.\move_from_junk.ps1 -Id J1A2B3C4D -Date 2026-10-25 -Apply
```

## Folders
`data\` CSV files and watermark, `emails\` one file per email (including emails you bring in with `search_mail.ps1 -Save`), `drafts\` the reply texts, `emails\deleted\` bodies of deleted mail. They are created on the first run. Keep copies of your CSV in a `backups\` folder before big clean-ups.

## Attachments
The CSV has `Has_Attachment` (`Yes` / `No`), `Attachment_Count` and `Attachment_Files` (file **names** only), and `active_context.md` marks such emails with `Attachment: YES` so the agent flags them instead of guessing what they contain. Attachments are not downloaded or read. Saving and reading them is not included in this version.

## Known limits
- Only the Inbox is synced. Junk is never synced; use the stats and rescue scripts. Older mail, mail in other folders and deleted mail is found with `search_mail.ps1`.
- Mail that arrives more than 2 days older than the watermark is not seen by the sync.
- Outlook must be open and loaded. A freshly started Outlook can return an empty first scan; the sync warns and retries once.
- Outlook's own junk filter can move a rescued email back to Junk. Add the sender to Safe Senders in Outlook.
- If `dispatch_replies.ps1` says "The operation failed", the stored `EntryID` is stale; `CLAUDE_PROMPT.md` explains the fix.
- Scripts compare dates in code because Outlook can misread date strings such as `10/11/2026` as 10 November.

## Privacy
Your mail lives in the folders above. Do not publish them: the included `.gitignore` keeps them out of Git.

## License
MIT. See [LICENSE](LICENSE). You may use, change and share this code, including commercially, as long as the copyright notice stays with it. It comes with no warranty: you run it on your own mailbox at your own risk.
