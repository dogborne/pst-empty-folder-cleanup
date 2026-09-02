# PST Empty Folder Cleanup

PowerShell utility for pre-import cleanup of empty folders in PST files. It uses Outlook COM automation to scan a PST, report empty leaf folders, and optionally delete them before a Microsoft 365 or Exchange archive import.

## Safety model

- Report-only by default.
- Creates a working copy by default; the source PST is not modified unless `-InPlace` is explicitly used.
- Deletion requires `-Delete`.
- Supports `-WhatIf` and `-Confirm`.
- Deletes deepest empty leaf folders first and repeats passes so parent folders that become empty can also be removed.
- Protects the PST root, default folders, search folders, and common system folders.
- Writes CSV, JSON, text summary, and log files.

## Requirements

- Windows PowerShell 5.1 or later.
- Microsoft Outlook desktop installed.
- A usable Outlook profile on the processing workstation.
- A copied PST or verified backup before running deletion.

## Usage

Report only:

```powershell
.\Remove-EmptyPstFolders.ps1 -PstPath "D:\PST\RawArchive.pst"
```

Preview deletion:

```powershell
.\Remove-EmptyPstFolders.ps1 -PstPath "D:\PST\RawArchive.pst" -Delete -WhatIf
```

Delete from a working copy and attempt compaction:

```powershell
.\Remove-EmptyPstFolders.ps1 -PstPath "D:\PST\RawArchive.pst" -Delete -Confirm:$false -Compact
```

Process the original PST in place:

```powershell
.\Remove-EmptyPstFolders.ps1 -PstPath "D:\PST\RawArchive.pst" -InPlace -Delete -Confirm:$false
```

Use `-InPlace` only when you have validated backups.

## Output

The script writes a timestamped output folder by default:

- `before-folders.csv`
- `after-folders.csv`
- `deletion-actions.csv`
- `pass-summary.csv`
- `summary.json`
- `summary.txt`
- `PstEmptyFolderCleanup.log`

## Notes

Outlook COM can be slow or fragile with very large PSTs or tens of thousands of folders. Start with report-only mode, then test deletion on a copy before using the cleaned PST for import.
