#requires -Version 5.1
<#
.SYNOPSIS
Pre-import PST cleanup for archive migrations.

.DESCRIPTION
Uses Outlook COM automation to scan a PST and identify/delete empty leaf folders.
By default, the script is report-only and creates a working copy of the PST (unless -InPlace is explicitly set).
Deletion only occurs when -Delete is supplied. Deletion runs deepest-first and repeats passes until no more empty
leaf folders are found or -MaxPasses is reached.

WARNING:
- Requires Outlook desktop (Windows) installed and a usable Outlook profile.
- Always run on a copy or with verified backups.
- Test on a small/sample PST before running against production-sized files.

.PARAMETER PstPath
Required path to the source PST file.

.PARAMETER OutputDirectory
Output folder for CSV reports, summary files, and logs.
Default: .\PstEmptyFolderCleanup-<timestamp>

.PARAMETER WorkingCopyPath
Optional explicit path for the working PST copy (used when -InPlace is not set).

.PARAMETER InPlace
If set, processes the source PST directly (no working copy). Use only with backups.

.PARAMETER Delete
If set, performs deletion of empty leaf folders. Without this switch the script is report-only.

.PARAMETER Compact
If set, attempts best-effort PST compaction via exposed Outlook COM methods (if available in this Outlook build).

.PARAMETER MaxPasses
Maximum deletion passes. Default: 20.

.PARAMETER KeepWorkingCopy
When report-only mode creates a working copy, keep it instead of removing it at the end.

.PARAMETER LogPath
Optional explicit path to the log file.

.EXAMPLE
.\Remove-EmptyPstFolders.ps1 -PstPath "D:\PST\RawArchive.pst"
Creates a working copy, scans it, and outputs before/after reports without deleting folders.

.EXAMPLE
.\Remove-EmptyPstFolders.ps1 -PstPath "D:\PST\RawArchive.pst" -Delete -WhatIf
Shows what deletions would occur (deepest-first passes) without deleting folders.

.EXAMPLE
.\Remove-EmptyPstFolders.ps1 -PstPath "D:\PST\RawArchive.pst" -Delete -Confirm:$false -Compact
Creates a working copy, deletes empty leaf folders in passes, then attempts compaction.

.EXAMPLE
.\Remove-EmptyPstFolders.ps1 -PstPath "D:\PST\RawArchive.pst" -InPlace -Delete -Confirm:$false
Processes the original PST in place. High risk: use only after backup validation.

.NOTES
WhatIf/Confirm support is implemented via SupportsShouldProcess for deletion/compaction actions.
Working-copy creation and report generation still occur when those operations are needed to run the script.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$PstPath,

    [Parameter()]
    [string]$OutputDirectory,

    [Parameter()]
    [string]$WorkingCopyPath,

    [Parameter()]
    [switch]$InPlace,

    [Parameter()]
    [switch]$Delete,

    [Parameter()]
    [switch]$Compact,

    [Parameter()]
    [ValidateRange(1, 500)]
    [int]$MaxPasses = 20,

    [Parameter()]
    [switch]$KeepWorkingCopy,

    [Parameter()]
    [string]$LogPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:CmdletContext = $PSCmdlet
$script:LogFilePath = $null

function Resolve-AbsolutePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter()]
        [switch]$MustExist
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw 'Path cannot be blank.'
    }

    $candidate = $Path
    if (-not [System.IO.Path]::IsPathRooted($candidate)) {
        $candidate = Join-Path -Path (Get-Location).Path -ChildPath $candidate
    }

    $fullPath = [System.IO.Path]::GetFullPath($candidate)
    if ($MustExist -and -not (Test-Path -LiteralPath $fullPath)) {
        throw "Path does not exist: $fullPath"
    }

    return $fullPath
}

function Ensure-Directory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DirectoryPath
    )

    if (Test-Path -LiteralPath $DirectoryPath) {
        $item = Get-Item -LiteralPath $DirectoryPath
        if (-not $item.PSIsContainer) {
            throw "Path exists but is not a directory: $DirectoryPath"
        }
        return
    }

    $null = New-Item -Path $DirectoryPath -ItemType Directory -Force
}

function Initialize-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $logParent = Split-Path -Path $Path -Parent
    if (-not [string]::IsNullOrWhiteSpace($logParent)) {
        Ensure-Directory -DirectoryPath $logParent
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType File -Force | Out-Null
    }

    $script:LogFilePath = $Path
}

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('INFO', 'WARN', 'ERROR', 'DEBUG')]
        [string]$Level,
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
    $line = "[{0}][{1}] {2}" -f $timestamp, $Level, $Message

    if ($script:LogFilePath) {
        Add-Content -LiteralPath $script:LogFilePath -Value $line -Encoding UTF8
    }

    switch ($Level) {
        'WARN' { Write-Warning $Message }
        'DEBUG' { Write-Verbose $Message }
        default { Write-Verbose $Message }
    }
}

function Release-ComObject {
    [CmdletBinding()]
    param(
        [Parameter()]
        $ComObject
    )

    if ($null -ne $ComObject -and [System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
    }
}

function New-OutlookSession {
    [CmdletBinding()]
    param()

    try {
        $application = New-Object -ComObject Outlook.Application
    }
    catch {
        throw "Unable to create Outlook COM object. Ensure Outlook desktop is installed. $($_.Exception.Message)"
    }

    try {
        $namespace = $application.GetNamespace('MAPI')
    }
    catch {
        Release-ComObject -ComObject $application
        throw "Unable to access Outlook MAPI namespace. $($_.Exception.Message)"
    }

    return [pscustomobject]@{
        Application = $application
        Namespace   = $namespace
    }
}

function Get-PstStoreByPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Namespace,
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $target = [System.IO.Path]::GetFullPath($Path)
    foreach ($store in @($Namespace.Stores)) {
        $storePath = $null
        try {
            $storePath = [string]$store.FilePath
        }
        catch {
            continue
        }

        if ([string]::IsNullOrWhiteSpace($storePath)) {
            continue
        }

        $normalizedStorePath = [System.IO.Path]::GetFullPath($storePath)
        if ([System.StringComparer]::OrdinalIgnoreCase.Equals($normalizedStorePath, $target)) {
            return $store
        }
    }

    return $null
}

function Mount-PstStore {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Namespace,
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $existingStore = Get-PstStoreByPath -Namespace $Namespace -Path $Path
    if ($existingStore) {
        return [pscustomobject]@{
            Store            = $existingStore
            AttachedByScript = $false
        }
    }

    if ($Namespace.PSObject.Methods.Name -contains 'AddStoreEx') {
        # 3 = olStoreUnicode
        $Namespace.AddStoreEx($Path, 3)
    }
    else {
        $Namespace.AddStore($Path)
    }

    $mountedStore = Get-PstStoreByPath -Namespace $Namespace -Path $Path
    if (-not $mountedStore) {
        throw "PST was added but could not be resolved in Outlook stores: $Path"
    }

    return [pscustomobject]@{
        Store            = $mountedStore
        AttachedByScript = $true
    }
}

function Get-ProtectedNameSet {
    [CmdletBinding()]
    param()

    $set = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $names = @(
        'Search Folders', 'Search Folder', 'Finder', 'Dossiers de recherche', 'Carpetas de búsqueda', 'Pastas de Pesquisa', 'Suchordner', 'Cartelle di ricerca',
        'Deleted Items', 'Deleted Messages', 'Trash', 'Bin', 'Corbeille', 'Elementos eliminados', 'Itens Excluídos', 'Gelöschte Elemente', 'Eliminata', 'Удаленные',
        'Inbox', 'Outbox', 'Sent Items', 'Sent Mail', 'Drafts', 'Calendar', 'Contacts', 'Tasks', 'Journal', 'Notes', 'Junk E-mail', 'Junk Email',
        'RSS Feeds', 'Sync Issues', 'Conflicts', 'Local Failures', 'Server Failures', 'Suggested Contacts',
        'Top of Personal Folders', 'Top of Outlook data file', 'Top of Information Store'
    )

    foreach ($name in $names) {
        $null = $set.Add($name)
    }

    return $set
}

function Get-ProtectedEntryIdSet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Store
    )

    $set = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $root = $Store.GetRootFolder()
    if ($root -and $root.EntryID) {
        $null = $set.Add([string]$root.EntryID)
    }

    $defaultFolderIds = @(3, 4, 5, 6, 9, 10, 11, 12, 13, 16, 19, 20, 21, 22, 23, 25, 28, 29, 30)
    foreach ($folderId in $defaultFolderIds) {
        try {
            $defaultFolder = $Store.GetDefaultFolder($folderId)
            if ($defaultFolder -and $defaultFolder.EntryID) {
                $null = $set.Add([string]$defaultFolder.EntryID)
            }
        }
        catch {
            continue
        }
    }

    return $set
}

function Get-FolderChildrenSnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Folder
    )

    $children = New-Object System.Collections.Generic.List[object]
    $count = [int]$Folder.Folders.Count
    for ($index = 1; $index -le $count; $index++) {
        $children.Add($Folder.Folders.Item($index))
    }
    return @($children)
}

function Test-IsProtectedFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Record,
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.HashSet[string]]$ProtectedEntryIds,
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.HashSet[string]]$ProtectedNames
    )

    if ($Record.Depth -eq 0) {
        return $true
    }

    if ($Record.EntryID -and $ProtectedEntryIds.Contains($Record.EntryID)) {
        return $true
    }

    if ($Record.Name -and $ProtectedNames.Contains($Record.Name)) {
        return $true
    }

    if ($Record.FolderPath -match '\\Search Folders(\\|$)') {
        return $true
    }

    if ($Record.FolderPath -match '\\Finder(\\|$)') {
        return $true
    }

    return $false
}

function Get-FolderInventory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $RootFolder,
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.HashSet[string]]$ProtectedEntryIds,
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.HashSet[string]]$ProtectedNames
    )

    $records = New-Object System.Collections.Generic.List[object]
    $folderMap = @{}

    $stack = New-Object System.Collections.Stack
    $stack.Push([pscustomobject]@{
            Folder = $RootFolder
            Depth  = 0
        })

    while ($stack.Count -gt 0) {
        $node = $stack.Pop()
        $folder = $node.Folder
        $depth = [int]$node.Depth

        $children = Get-FolderChildrenSnapshot -Folder $folder
        $childCount = $children.Count
        $itemCount = [int]$folder.Items.Count
        $entryId = [string]$folder.EntryID
        $folderPath = [string]$folder.FolderPath
        $name = [string]$folder.Name
        $defaultItemType = $null
        try {
            $defaultItemType = [int]$folder.DefaultItemType
        }
        catch {
            $defaultItemType = $null
        }

        $baseRecord = [pscustomobject]@{
            Depth            = $depth
            FolderPath       = $folderPath
            Name             = $name
            EntryID          = $entryId
            ItemCount        = $itemCount
            ChildFolderCount = $childCount
            IsLeaf           = ($childCount -eq 0)
            IsEmptyLeaf      = ($childCount -eq 0 -and $itemCount -eq 0)
            DefaultItemType  = $defaultItemType
        }

        $isProtected = Test-IsProtectedFolder -Record $baseRecord -ProtectedEntryIds $ProtectedEntryIds -ProtectedNames $ProtectedNames
        $record = [pscustomobject]@{
            Depth            = $baseRecord.Depth
            FolderPath       = $baseRecord.FolderPath
            Name             = $baseRecord.Name
            EntryID          = $baseRecord.EntryID
            ItemCount        = $baseRecord.ItemCount
            ChildFolderCount = $baseRecord.ChildFolderCount
            IsLeaf           = $baseRecord.IsLeaf
            IsEmptyLeaf      = $baseRecord.IsEmptyLeaf
            IsProtected      = $isProtected
            DeleteCandidate  = ($baseRecord.IsEmptyLeaf -and -not $isProtected)
            DefaultItemType  = $baseRecord.DefaultItemType
        }

        $records.Add($record)
        if ($entryId) {
            $folderMap[$entryId] = $folder
        }

        for ($childIndex = $childCount - 1; $childIndex -ge 0; $childIndex--) {
            $stack.Push([pscustomobject]@{
                    Folder = $children[$childIndex]
                    Depth  = $depth + 1
                })
        }
    }

    return [pscustomobject]@{
        Inventory = @($records)
        FolderMap = $folderMap
    }
}

function Remove-EmptyLeafCandidates {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Candidates,
        [Parameter(Mandatory = $true)]
        [hashtable]$FolderMap,
        [Parameter(Mandatory = $true)]
        [int]$PassNumber
    )

    $deletedCount = 0
    $actionRows = New-Object System.Collections.Generic.List[object]

    foreach ($candidate in $Candidates) {
        $result = 'Skipped'
        $detail = ''

        if (-not $Delete) {
            $result = 'ReportOnly'
            $detail = 'Delete switch was not supplied.'
        }
        elseif (-not $FolderMap.ContainsKey($candidate.EntryID)) {
            $result = 'MissingFolderObject'
            $detail = 'Folder object not found in this pass map.'
        }
        elseif ($script:CmdletContext.ShouldProcess($candidate.FolderPath, 'Delete empty leaf folder')) {
            try {
                $FolderMap[$candidate.EntryID].Delete()
                $deletedCount++
                $result = 'Deleted'
                $detail = 'Deleted successfully.'
            }
            catch {
                $result = 'Error'
                $detail = $_.Exception.Message
                Write-Log -Level 'ERROR' -Message ("Failed to delete '{0}': {1}" -f $candidate.FolderPath, $_.Exception.Message)
                throw
            }
        }
        else {
            $result = 'WhatIfOrDeclined'
            $detail = 'Deletion not executed due WhatIf/Confirm decision.'
        }

        $actionRows.Add([pscustomobject]@{
                Pass      = $PassNumber
                Depth     = $candidate.Depth
                FolderPath = $candidate.FolderPath
                EntryID   = $candidate.EntryID
                Result    = $result
                Detail    = $detail
            })
    }

    return [pscustomobject]@{
        DeletedCount = $deletedCount
        Actions      = @($actionRows)
    }
}

function Invoke-PstCompaction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Store,
        [Parameter(Mandatory = $true)]
        [string]$PstFilePath
    )

    $result = [ordered]@{
        Requested = [bool]$Compact
        Attempted = $false
        Succeeded = $false
        Method    = $null
        Message   = 'Compaction not requested.'
    }

    if (-not $Compact) {
        return [pscustomobject]$result
    }

    $result.Message = 'No COM compaction method is exposed by this Outlook build.'
    $bindingFlags = [System.Reflection.BindingFlags]::InvokeMethod
    foreach ($methodName in @('CompactNow', 'Compact')) {
        try {
            if ($script:CmdletContext.ShouldProcess($PstFilePath, "Invoke Outlook COM method '$methodName'")) {
                $result.Attempted = $true
                [void]$Store.GetType().InvokeMember($methodName, $bindingFlags, $null, $Store, $null)
                $result.Succeeded = $true
                $result.Method = $methodName
                $result.Message = "Compaction method '$methodName' invoked."
                return [pscustomobject]$result
            }
            else {
                $result.Attempted = $true
                $result.Method = $methodName
                $result.Message = "Compaction method '$methodName' skipped due WhatIf/Confirm decision."
                return [pscustomobject]$result
            }
        }
        catch {
            continue
        }
    }

    return [pscustomobject]$result
}

function Write-TextSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Summary,
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $lines = @(
        'PST Empty Folder Cleanup Summary',
        ('RunTimestamp: {0}' -f $Summary.RunTimestamp),
        ('InputPstPath: {0}' -f $Summary.InputPstPath),
        ('ProcessedPstPath: {0}' -f $Summary.ProcessedPstPath),
        ('InPlace: {0}' -f $Summary.InPlace),
        ('DeleteMode: {0}' -f $Summary.DeleteMode),
        ('MaxPasses: {0}' -f $Summary.MaxPasses),
        ('PassesExecuted: {0}' -f $Summary.PassesExecuted),
        ('BeforeFolderCount: {0}' -f $Summary.BeforeFolderCount),
        ('BeforeDeleteCandidates: {0}' -f $Summary.BeforeDeleteCandidates),
        ('DeletedFolders: {0}' -f $Summary.DeletedFolders),
        ('AfterFolderCount: {0}' -f $Summary.AfterFolderCount),
        ('AfterDeleteCandidates: {0}' -f $Summary.AfterDeleteCandidates),
        ('CompactionRequested: {0}' -f $Summary.Compaction.Requested),
        ('CompactionAttempted: {0}' -f $Summary.Compaction.Attempted),
        ('CompactionSucceeded: {0}' -f $Summary.Compaction.Succeeded),
        ('CompactionMethod: {0}' -f $Summary.Compaction.Method),
        ('CompactionMessage: {0}' -f $Summary.Compaction.Message),
        ('OutputDirectory: {0}' -f $Summary.Reports.OutputDirectory),
        ('LogPath: {0}' -f $Summary.Reports.LogPath)
    )

    Set-Content -LiteralPath $Path -Value $lines -Encoding UTF8
}

$outlookApplication = $null
$mapiNamespace = $null
$pstStore = $null
$attachedByScript = $false
$workingCopyCreated = $false
$targetPstPath = $null
$removeWorkingCopyOnExit = $false
$summary = $null

try {
    $resolvedSourcePst = Resolve-AbsolutePath -Path $PstPath -MustExist
    if ([System.IO.Path]::GetExtension($resolvedSourcePst) -ne '.pst') {
        throw "PstPath must point to a .pst file: $resolvedSourcePst"
    }

    if ($InPlace -and $WorkingCopyPath) {
        throw 'WorkingCopyPath cannot be used together with -InPlace.'
    }

    if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
        $OutputDirectory = Join-Path -Path (Get-Location).Path -ChildPath ("PstEmptyFolderCleanup-{0}" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
    }
    $resolvedOutputDirectory = Resolve-AbsolutePath -Path $OutputDirectory
    Ensure-Directory -DirectoryPath $resolvedOutputDirectory

    if ([string]::IsNullOrWhiteSpace($LogPath)) {
        $LogPath = Join-Path -Path $resolvedOutputDirectory -ChildPath 'PstEmptyFolderCleanup.log'
    }
    else {
        $LogPath = Resolve-AbsolutePath -Path $LogPath
    }
    Initialize-Log -Path $LogPath

    Write-Log -Level 'INFO' -Message "Starting run for PST: $resolvedSourcePst"
    Write-Log -Level 'INFO' -Message ("Mode: {0}; Delete: {1}; MaxPasses: {2}" -f ($(if ($InPlace) { 'InPlace' } else { 'WorkingCopy' }), [bool]$Delete, $MaxPasses))

    if ($InPlace) {
        $targetPstPath = $resolvedSourcePst
        Write-Log -Level 'WARN' -Message 'InPlace mode enabled. Original PST may be modified.'
    }
    else {
        if ([string]::IsNullOrWhiteSpace($WorkingCopyPath)) {
            $base = [System.IO.Path]::GetFileNameWithoutExtension($resolvedSourcePst)
            $ext = [System.IO.Path]::GetExtension($resolvedSourcePst)
            $WorkingCopyPath = Join-Path -Path $resolvedOutputDirectory -ChildPath ("{0}.working{1}" -f $base, $ext)
        }
        $resolvedWorkingCopyPath = Resolve-AbsolutePath -Path $WorkingCopyPath

        if (Test-Path -LiteralPath $resolvedWorkingCopyPath) {
            throw "Working copy path already exists. Choose a different WorkingCopyPath: $resolvedWorkingCopyPath"
        }

        Write-Log -Level 'INFO' -Message "Creating working copy: $resolvedWorkingCopyPath"
        Copy-Item -LiteralPath $resolvedSourcePst -Destination $resolvedWorkingCopyPath -Force
        $workingCopyCreated = $true
        $targetPstPath = $resolvedWorkingCopyPath
    }

    $session = New-OutlookSession
    $outlookApplication = $session.Application
    $mapiNamespace = $session.Namespace
    Write-Log -Level 'INFO' -Message 'Outlook COM session initialized.'

    $mountResult = Mount-PstStore -Namespace $mapiNamespace -Path $targetPstPath
    $pstStore = $mountResult.Store
    $attachedByScript = [bool]$mountResult.AttachedByScript
    Write-Log -Level 'INFO' -Message ("PST mounted. AttachedByScript={0}" -f $attachedByScript)

    $rootFolder = $pstStore.GetRootFolder()
    $protectedNames = Get-ProtectedNameSet
    $protectedEntryIds = Get-ProtectedEntryIdSet -Store $pstStore

    $beforeScan = Get-FolderInventory -RootFolder $rootFolder -ProtectedEntryIds $protectedEntryIds -ProtectedNames $protectedNames
    $beforeInventory = $beforeScan.Inventory
    $beforeCandidates = @($beforeInventory | Where-Object { $_.DeleteCandidate })
    Write-Log -Level 'INFO' -Message ("Before scan complete. Folders={0}; EmptyLeafCandidates={1}" -f $beforeInventory.Count, $beforeCandidates.Count)

    $allActions = New-Object System.Collections.Generic.List[object]
    $passSummaries = New-Object System.Collections.Generic.List[object]
    $passesExecuted = 0
    $totalDeleted = 0
    $currentScan = $beforeScan

    if ($Delete) {
        while ($passesExecuted -lt $MaxPasses) {
            $passesExecuted++
            $inventory = $currentScan.Inventory
            $folderMap = $currentScan.FolderMap
            $candidates = @($inventory | Where-Object { $_.DeleteCandidate } | Sort-Object -Property @{ Expression = 'Depth'; Descending = $true }, @{ Expression = 'FolderPath'; Descending = $true })
            $candidateCount = $candidates.Count

            if ($candidateCount -eq 0) {
                $passSummaries.Add([pscustomobject]@{
                        Pass               = $passesExecuted
                        CandidateCount     = 0
                        DeletedCount       = 0
                        RemainingAfterPass = 0
                        Note               = 'No candidates found. Stopping.'
                    })
                break
            }

            Write-Log -Level 'INFO' -Message ("Pass {0}: candidates={1}" -f $passesExecuted, $candidateCount)
            $passResult = Remove-EmptyLeafCandidates -Candidates $candidates -FolderMap $folderMap -PassNumber $passesExecuted
            foreach ($action in $passResult.Actions) {
                $allActions.Add($action)
            }

            $totalDeleted += [int]$passResult.DeletedCount
            $currentScan = Get-FolderInventory -RootFolder $rootFolder -ProtectedEntryIds $protectedEntryIds -ProtectedNames $protectedNames
            $remaining = @($currentScan.Inventory | Where-Object { $_.DeleteCandidate }).Count

            $passSummaries.Add([pscustomobject]@{
                    Pass               = $passesExecuted
                    CandidateCount     = $candidateCount
                    DeletedCount       = [int]$passResult.DeletedCount
                    RemainingAfterPass = $remaining
                    Note               = ''
                })

            if ($passResult.DeletedCount -le 0) {
                Write-Log -Level 'WARN' -Message ("Pass {0}: no folders deleted. Stopping to avoid non-progress loop." -f $passesExecuted)
                break
            }

            if ($remaining -eq 0) {
                Write-Log -Level 'INFO' -Message ("Pass {0}: no remaining candidates." -f $passesExecuted)
                break
            }
        }
    }
    else {
        $passesExecuted = 1
        $passSummaries.Add([pscustomobject]@{
                Pass               = 1
                CandidateCount     = $beforeCandidates.Count
                DeletedCount       = 0
                RemainingAfterPass = $beforeCandidates.Count
                Note               = 'Report-only mode. No deletions performed.'
            })
    }

    $compaction = Invoke-PstCompaction -Store $pstStore -PstFilePath $targetPstPath
    if ($compaction.Requested -and -not $compaction.Succeeded) {
        Write-Log -Level 'WARN' -Message $compaction.Message
    }
    elseif ($compaction.Succeeded) {
        Write-Log -Level 'INFO' -Message $compaction.Message
    }

    $afterScan = Get-FolderInventory -RootFolder $rootFolder -ProtectedEntryIds $protectedEntryIds -ProtectedNames $protectedNames
    $afterInventory = $afterScan.Inventory
    $afterCandidates = @($afterInventory | Where-Object { $_.DeleteCandidate })

    $beforeCsvPath = Join-Path -Path $resolvedOutputDirectory -ChildPath 'before-folders.csv'
    $afterCsvPath = Join-Path -Path $resolvedOutputDirectory -ChildPath 'after-folders.csv'
    $actionsCsvPath = Join-Path -Path $resolvedOutputDirectory -ChildPath 'deletion-actions.csv'
    $passSummaryCsvPath = Join-Path -Path $resolvedOutputDirectory -ChildPath 'pass-summary.csv'
    $summaryJsonPath = Join-Path -Path $resolvedOutputDirectory -ChildPath 'summary.json'
    $summaryTextPath = Join-Path -Path $resolvedOutputDirectory -ChildPath 'summary.txt'

    $reportFields = @('Depth', 'FolderPath', 'Name', 'EntryID', 'ItemCount', 'ChildFolderCount', 'IsLeaf', 'IsEmptyLeaf', 'IsProtected', 'DeleteCandidate', 'DefaultItemType')
    $beforeInventory | Select-Object $reportFields | Export-Csv -LiteralPath $beforeCsvPath -NoTypeInformation -Encoding UTF8
    $afterInventory | Select-Object $reportFields | Export-Csv -LiteralPath $afterCsvPath -NoTypeInformation -Encoding UTF8
    $passSummaries | Export-Csv -LiteralPath $passSummaryCsvPath -NoTypeInformation -Encoding UTF8

    if ($allActions.Count -gt 0) {
        $allActions | Export-Csv -LiteralPath $actionsCsvPath -NoTypeInformation -Encoding UTF8
    }
    else {
        Set-Content -LiteralPath $actionsCsvPath -Value 'No deletion actions recorded.' -Encoding UTF8
    }

    $summary = [pscustomobject]@{
        ScriptName             = 'Remove-EmptyPstFolders.ps1'
        RunTimestamp           = (Get-Date).ToString('o')
        InputPstPath           = $resolvedSourcePst
        ProcessedPstPath       = $targetPstPath
        InPlace                = [bool]$InPlace
        WorkingCopyCreated     = $workingCopyCreated
        DeleteMode             = [bool]$Delete
        MaxPasses              = $MaxPasses
        PassesExecuted         = $passesExecuted
        BeforeFolderCount      = $beforeInventory.Count
        BeforeDeleteCandidates = $beforeCandidates.Count
        DeletedFolders         = $totalDeleted
        AfterFolderCount       = $afterInventory.Count
        AfterDeleteCandidates  = $afterCandidates.Count
        Compaction             = $compaction
        Reports                = [pscustomobject]@{
            OutputDirectory = $resolvedOutputDirectory
            BeforeCsv       = $beforeCsvPath
            AfterCsv        = $afterCsvPath
            ActionsCsv      = $actionsCsvPath
            PassSummaryCsv  = $passSummaryCsvPath
            SummaryJson     = $summaryJsonPath
            SummaryText     = $summaryTextPath
            LogPath         = $LogPath
        }
    }

    $summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $summaryJsonPath -Encoding UTF8
    Write-TextSummary -Summary $summary -Path $summaryTextPath
    Write-Log -Level 'INFO' -Message "Summary written to: $summaryJsonPath"

    if ($workingCopyCreated -and -not $Delete -and -not $KeepWorkingCopy) {
        $removeWorkingCopyOnExit = $true
        Write-Log -Level 'INFO' -Message 'Working copy scheduled for removal at end of report-only run. Use -KeepWorkingCopy to retain it.'
    }
}
catch {
    if ($script:LogFilePath) {
        Write-Log -Level 'ERROR' -Message ("Run failed: {0}" -f $_.Exception.Message)
    }
    throw
}
finally {
    if ($attachedByScript -and $mapiNamespace -and $pstStore) {
        try {
            $rootToRemove = $pstStore.GetRootFolder()
            $mapiNamespace.RemoveStore($rootToRemove)
            Write-Log -Level 'INFO' -Message 'Detached PST store from Outlook profile.'
            Release-ComObject -ComObject $rootToRemove
        }
        catch {
            Write-Warning ("Unable to detach PST store cleanly: {0}" -f $_.Exception.Message)
        }
    }

    Release-ComObject -ComObject $pstStore
    Release-ComObject -ComObject $mapiNamespace
    Release-ComObject -ComObject $outlookApplication
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    if ($removeWorkingCopyOnExit -and $targetPstPath -and (Test-Path -LiteralPath $targetPstPath)) {
        try {
            Remove-Item -LiteralPath $targetPstPath -Force
            Write-Log -Level 'INFO' -Message "Removed working copy: $targetPstPath"
        }
        catch {
            Write-Warning ("Unable to remove working copy '{0}': {1}" -f $targetPstPath, $_.Exception.Message)
        }
    }
}

if ($summary) {
    $summary
}
