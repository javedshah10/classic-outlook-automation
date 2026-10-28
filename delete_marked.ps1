<#
.SYNOPSIS
    Deletes emails you marked with Triage_Status = Delete in data\master_*.csv.
    For each marked row: moves the Outlook email to Deleted Items (never permanent),
    moves its body file to emails\deleted\, keeps the row and sets Triage_Status = Closed.
    Default is a DRY RUN that only prints what it would do. Add -Apply to do it.
#>
param([switch]$Apply)

$ErrorActionPreference = "Stop"
$Utf8Enc = if ($PSVersionTable.PSVersion.Major -ge 6) { "utf8BOM" } else { "UTF8" }

# Must match $TargetEmail in sync_inbox.ps1 (set_email.ps1 updates both)
$TargetEmail = "your.name@example.com"

$BaseDir    = $PSScriptRoot
$DataDir    = Join-Path $BaseDir "data"
$DeletedDir = Join-Path $BaseDir "emails\deleted"
$PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"

$files = @(Get-ChildItem -Path $DataDir -Filter "master_*.csv" -ErrorAction SilentlyContinue)
$marked = @()
$tables = @{}
foreach ($f in $files) {
    $rows = @(Import-Csv -Path $f.FullName -Encoding $Utf8Enc)
    $tables[$f.FullName] = $rows
    $marked += @($rows | Where-Object { $_.Triage_Status -eq "Delete" } | ForEach-Object { [PSCustomObject]@{ File = $f.FullName; Row = $_ } })
}

if ($marked.Count -eq 0) {
    Write-Host "No rows with Triage_Status = Delete. Nothing to do." -ForegroundColor Green
    exit 0
}

Write-Host "$($marked.Count) email(s) marked Delete:" -ForegroundColor Cyan
foreach ($m in $marked) { Write-Host "  $($m.Row.ShortID) | $($m.Row.SenderEmail) | $($m.Row.Subject)" }
if (-not $Apply) {
    Write-Host "DRY RUN: nothing changed. Re-run with -Apply to delete these." -ForegroundColor Yellow
    exit 0
}

$Outlook   = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")
$Account = $null
foreach ($acc in $Namespace.Accounts) { if ($acc.SmtpAddress -ieq $TargetEmail) { $Account = $acc; break } }
if (-not $Account) { Write-Host "Account $TargetEmail not found." -ForegroundColor Red; exit 1 }
$Inbox = $Account.DeliveryStore.GetDefaultFolder(6)

function Find-Mail ($row) {
    # 1) by EntryID  2) fallback by the stable InternetMessageID (EntryID can change)
    try { $m = $Namespace.GetItemFromID($row.EntryID); if ($m) { return $m } } catch { }
    if ($row.InternetMessageID) {
        try {
            $q = '@SQL="' + $PR_INTERNET_MESSAGE_ID + '" = ''' + ($row.InternetMessageID -replace "'", "''") + ''''
            $hit = $Inbox.Items.Restrict($q)
            if ($hit.Count -eq 1) { return $hit.Item(1) }
        } catch { }
    }
    return $null
}

if (-not (Test-Path $DeletedDir)) { New-Item -ItemType Directory -Path $DeletedDir | Out-Null }

$done = 0; $notFound = 0
foreach ($m in $marked) {
    $row = $m.Row
    $mail = Find-Mail $row
    if ($mail) {
        $mail.Delete()                      # normal delete: goes to Deleted Items
        $status = "moved to Deleted Items"
        $done++
    } else {
        $status = "not found in Outlook (already gone); row closed anyway"
        $notFound++
    }
    if ($row.BodyFile) {
        $src = Join-Path $BaseDir $row.BodyFile
        if (Test-Path $src) {
            Move-Item -Path $src -Destination (Join-Path $DeletedDir (Split-Path $src -Leaf)) -Force
            $row.BodyFile = "emails/deleted/" + (Split-Path $src -Leaf)
        }
    }
    $row.Triage_Status = "Closed"
    Write-Host "  [DELETED] $($row.ShortID): $status" -ForegroundColor Green
}

foreach ($f in $tables.Keys) { $tables[$f] | Export-Csv -Path $f -NoTypeInformation -Encoding $Utf8Enc }
Write-Host "Done. $done deleted, $notFound not found. Run sync_inbox.ps1 to refresh active_context.md." -ForegroundColor Green
