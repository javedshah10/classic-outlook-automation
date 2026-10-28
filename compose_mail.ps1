<#
.SYNOPSIS
    Prepares a NEW email (not a reply) and opens it in an Outlook compose window, exactly as if you had clicked New:
    Outlook adds your own signature. You read it and press Send yourself.
    This script never sends. (Only if you set $AllowAutoSend = $true below yourself does -Send send it.)
    Fills in To, Cc, Bcc, subject, text and.
    -DryRun checks everything and prints a summary without opening anything.
.EXAMPLE
    .\compose_mail.ps1 -To "a@x.com" -Cc "b@y.com; c@z.com" -Subject "Report" -BodyFile drafts\note.md
    .\compose_mail.ps1 -To "a@x.com" -Subject "Hello" -Body "Short text" -DryRun
#>
param(
    [string]$To,                 # one or more addresses separated by ; or ,
    [string]$Cc,
    [string]$Bcc,
    [string]$Subject,
    [string]$BodyFile,           # plain-text file (UTF-8) with the message text
    [string]$Body,               # or the text itself
    [switch]$Send,               # only sends if $AllowAutoSend below is $true; otherwise the window just opens
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"
# Must match $TargetEmail in sync_inbox.ps1 (set_email.ps1 updates all scripts)
$TargetEmail = "your.name@example.com"
# Same switch idea as dispatch_replies.ps1: keep $false unless you fully accept that this script can send mail.
$AllowAutoSend = $false
$MaxRecipients = 20
$Utf8Enc = if ($PSVersionTable.PSVersion.Major -ge 6) { "utf8BOM" } else { "UTF8" }

function Fail ($msg) { Write-Host "STOPPED: $msg" -ForegroundColor Red; exit 1 }
function Split-Addresses ($text) { @(($text -split '[;,]') | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
$rx = '^[^@\s";<>]+@[^@\s";<>]+\.[^@\s";<>]+$'

$toList = @(Split-Addresses $To); $ccList = @(Split-Addresses $Cc); $bccList = @(Split-Addresses $Bcc)   # @() keeps a single address a list
if ($toList.Count -eq 0) { Fail "give at least one -To address." }
foreach ($a in ($toList + $ccList + $bccList)) { if ($a -notmatch $rx) { Fail "'$a' is not a valid email address." } }
if (($toList.Count + $ccList.Count + $bccList.Count) -gt $MaxRecipients) { Fail "more than $MaxRecipients recipients; split the mail or raise `$MaxRecipients on purpose." }
if (-not $Subject -or -not $Subject.Trim()) { Fail "give a -Subject." }
if ($Body -and $BodyFile) { Fail "use either -Body or -BodyFile, not both." }
if ($BodyFile) {
    if (-not (Test-Path -LiteralPath $BodyFile -PathType Leaf)) { Fail "body file not found: $BodyFile" }
    $Body = Get-Content -LiteralPath $BodyFile -Raw -Encoding UTF8
}
if (-not $Body -or -not $Body.Trim()) { Fail "give the message text with -Body or -BodyFile." }
$Body = $Body.Trim()
if ($Body.Length -gt 20000) { Fail "message text is over 20000 characters." }

$attText = ""

$sending = ($Send -and $AllowAutoSend)
Write-Host "To:      $($toList -join '; ')"
if ($ccList.Count)  { Write-Host "Cc:      $($ccList -join '; ')" }
if ($bccList.Count) { Write-Host "Bcc:     $($bccList -join '; ')" }
Write-Host "Subject: $Subject"
if ($attText) { Write-Host "Attach:  $attText" }
Write-Host ("Text:    {0} characters | action: {1}" -f $Body.Length, $(if ($sending) { "SEND" } else { "open in a compose window for you to review" }))
if ($Send -and -not $AllowAutoSend) { Write-Host "Note: -Send was given but `$AllowAutoSend is `$false, so the window will just open." -ForegroundColor Yellow }
if ($DryRun) { Write-Host "DRY RUN: nothing opened." -ForegroundColor Yellow; exit 0 }

$ol = New-Object -ComObject Outlook.Application
$acc = $null; foreach ($a in $ol.GetNamespace("MAPI").Accounts) { if ($a.SmtpAddress -ieq $TargetEmail) { $acc = $a; break } }
if (-not $acc) { Fail "account $TargetEmail not found in Outlook." }

# Outlook can answer "Operation aborted" while it is still starting up or shutting down after the previous script; wait and retry.
# (CreateItem is refused on some setups, so the item is created in the Drafts folder.)
$mail = $null
for ($try = 1; $try -le 5 -and -not $mail; $try++) {
    try { $mail = $acc.DeliveryStore.GetDefaultFolder(16).Items.Add("IPM.Note") } catch { if ($try -eq 5) { throw }; Start-Sleep -Seconds 4 }
}
$mail.SendUsingAccount = $acc
$mail.Display()                              # opens the compose window; Outlook adds your signature, just like clicking New
$mail.SendUsingAccount = $acc                  # set again after the window opens, or Outlook falls back to the default account
$enc = [System.Net.WebUtility]::HtmlEncode($Body) -replace "`r?`n", "<br>"
$block = "<div dir=`"auto`" style=`"font-family:Calibri,Arial,sans-serif;font-size:11pt`">$enc</div><br>"
$html = $mail.HTMLBody
$mt = [regex]::Match($html, '(?is)<body[^>]*>')
if ($mt.Success) { $mail.HTMLBody = $html.Insert($mt.Index + $mt.Length, $block) } else { $mail.HTMLBody = $block + $html }

$mail.To = ($toList -join "; ")
if ($ccList.Count)  { $mail.CC = ($ccList -join "; ") }
if ($bccList.Count) { $mail.BCC = ($bccList -join "; ") }
$mail.Subject = $Subject
if ($sending) { $mail.Send(); $result = "Sent" }
else {
    $result = "Opened"
    # Also keep a copy in Drafts. Outlook sometimes refuses this ('Operation aborted'); then just press Ctrl+S in the window.
    try { $mail.Save(); $result = "Opened+Draft" } catch { Write-Host "Note: could not save to Drafts from the script; press Ctrl+S in the compose window if you want a draft." -ForegroundColor Yellow }
}
Write-Host "$(if ($sending) { '[SENT]' } else { '[OPENED FOR REVIEW - press Send in Outlook when it looks right]' }) $Subject -> $($toList -join '; ')" -ForegroundColor Green

# Small audit log (stays local; ignored by .gitignore)
$DataDir = Join-Path $PSScriptRoot "data"; if (-not (Test-Path $DataDir)) { New-Item -ItemType Directory -Path $DataDir | Out-Null }
[PSCustomObject]@{ Time = (Get-Date -Format "yyyy-MM-dd HH:mm:ss"); Result = $result; To = ($toList -join "; "); Cc = ($ccList -join "; "); Bcc = ($bccList -join "; "); Subject = $Subject; Attachments = $attText } |
    Export-Csv -Path (Join-Path $DataDir "compose_log.csv") -NoTypeInformation -Append -Encoding $Utf8Enc
