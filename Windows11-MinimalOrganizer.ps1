[CmdletBinding()]
param(
    [bool]$DryRun = $true,
    [ValidateRange(1, 86400)]
    [int]$MinFileAgeSeconds = 120,
    [ValidateSet('Initial', 'Maintenance')]
    [string]$Mode = 'Initial',
    [string]$TargetRootOverride,
    [string[]]$SourceFoldersOverride,
    [switch]$Undo
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:RunId = [guid]::NewGuid().ToString()
$script:RunStart = Get-Date
$script:CurrentLogFile = $null

function Get-UserPaths {
    param(
        [string]$TargetRootOverride,
        [string[]]$SourceFoldersOverride
    )

    $profile = [Environment]::GetFolderPath('UserProfile')
    $desktop = [Environment]::GetFolderPath('Desktop')
    $downloads = Join-Path $profile 'Downloads'
    $documents = [Environment]::GetFolderPath('MyDocuments')
    $pictures = [Environment]::GetFolderPath('MyPictures')
    $videos = [Environment]::GetFolderPath('MyVideos')
    $music = [Environment]::GetFolderPath('MyMusic')

    $targetRoot = if ($TargetRootOverride) { $TargetRootOverride } else { Join-Path $profile 'Ablage' }

    $defaultSources = @($desktop, $downloads, $documents, $pictures, $videos, $music)
    $sources = if ($SourceFoldersOverride -and $SourceFoldersOverride.Count -gt 0) { $SourceFoldersOverride } else { $defaultSources }

    [PSCustomObject]@{
        UserProfile = $profile
        Desktop = $desktop
        Downloads = $downloads
        Documents = $documents
        Pictures = $pictures
        Videos = $videos
        Music = $music
        TargetRoot = $targetRoot
        Sources = $sources
    }
}

function Ensure-Folder {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        [void](New-Item -ItemType Directory -Path $Path -Force)
    }
}

function Initialize-FolderStructure {
    param([Parameter(Mandatory = $true)][string]$TargetRoot)

    $paths = @{
        Root = $TargetRoot
        Inbox = Join-Path $TargetRoot '01_Inbox'
        Documents = Join-Path $TargetRoot '02_Dokumente'
        Pictures = Join-Path $TargetRoot '03_Bilder'
        Videos = Join-Path $TargetRoot '04_Videos'
        Audio = Join-Path $TargetRoot '05_Audio'
        Archive = Join-Path $TargetRoot '06_Archive'
        Installer = Join-Path $TargetRoot '07_Setup_Installer'
        OldUnsorted = Join-Path $TargetRoot '99_Alt\Unsortiert'
        Logs = Join-Path $TargetRoot '_Logs'
    }

    foreach ($k in $paths.Keys) {
        Ensure-Folder -Path $paths[$k]
    }

    return [PSCustomObject]$paths
}

function Get-CategoryByExtension {
    param([Parameter(Mandatory = $true)][string]$Extension)

    $ext = $Extension.TrimStart('.').ToLowerInvariant()
    switch ($ext) {
        { $_ -in @('pdf', 'doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'txt', 'rtf', 'csv', 'md') } { return 'Documents' }
        { $_ -in @('jpg', 'jpeg', 'png', 'gif', 'webp', 'heic', 'tiff') } { return 'Pictures' }
        { $_ -in @('mp4', 'mov', 'mkv', 'avi') } { return 'Videos' }
        { $_ -in @('mp3', 'wav', 'flac', 'm4a') } { return 'Audio' }
        { $_ -in @('zip', '7z', 'rar', 'tar', 'gz') } { return 'Archive' }
        { $_ -in @('exe', 'msi', 'msp', 'iso') } { return 'Installer' }
        default { return 'Inbox' }
    }
}

function Get-DateBuckets {
    param(
        [Parameter(Mandatory = $true)][System.IO.FileInfo]$File,
        [Parameter(Mandatory = $true)][string]$Category
    )

    $fileDate = $File.LastWriteTime

    if ($Category -eq 'Pictures') {
        Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue
        $exifDatePropId = 36867
        try {
            $image = [System.Drawing.Image]::FromFile($File.FullName)
            $prop = $image.PropertyItems | Where-Object { $_.Id -eq $exifDatePropId } | Select-Object -First 1
            if ($prop) {
                $raw = [System.Text.Encoding]::ASCII.GetString($prop.Value).Trim([char]0)
                $parsed = [datetime]::MinValue
                if ([datetime]::TryParseExact($raw, 'yyyy:MM:dd HH:mm:ss', $null, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
                    $fileDate = $parsed
                }
            }
            $image.Dispose()
        }
        catch {
            # Optional EXIF read failed -> fallback to LastWriteTime
        }
    }

    [PSCustomObject]@{
        Year = $fileDate.ToString('yyyy')
        Month = $fileDate.ToString('MM')
        Date = $fileDate
    }
}

function Get-TargetPath {
    param(
        [Parameter(Mandatory = $true)][System.IO.FileInfo]$File,
        [Parameter(Mandatory = $true)][string]$Category,
        [Parameter(Mandatory = $true)]$Folders
    )

    $dateBuckets = Get-DateBuckets -File $File -Category $Category

    switch ($Category) {
        'Documents' { return (Join-Path $Folders.Documents $dateBuckets.Year) }
        'Pictures' { return (Join-Path (Join-Path $Folders.Pictures $dateBuckets.Year) $dateBuckets.Month) }
        'Videos' { return (Join-Path $Folders.Videos $dateBuckets.Year) }
        'Audio' { return $Folders.Audio }
        'Archive' { return $Folders.Archive }
        'Installer' { return $Folders.Installer }
        default { return $Folders.Inbox }
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$LogFile,
        [Parameter(Mandatory = $true)][hashtable]$Entry
    )

    $line = [PSCustomObject]$Entry | ConvertTo-Json -Compress -Depth 5
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
}

function Move-FileSafely {
    param(
        [Parameter(Mandatory = $true)][System.IO.FileInfo]$File,
        [Parameter(Mandatory = $true)][string]$DestinationFolder,
        [Parameter(Mandatory = $true)][bool]$DryRun,
        [Parameter(Mandatory = $true)][string]$LogFile,
        [Parameter(Mandatory = $true)][string]$Category,
        [Parameter(Mandatory = $true)][string]$Rule
    )

    Ensure-Folder -Path $DestinationFolder

    $targetPath = Join-Path $DestinationFolder $File.Name
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($File.Name)
    $ext = $File.Extension
    $counter = 1

    while (Test-Path -LiteralPath $targetPath) {
        $suffix = '_{0:d3}' -f $counter
        $targetPath = Join-Path $DestinationFolder ($baseName + $suffix + $ext)
        $counter++
    }

    $status = 'Planned'
    $errorMessage = $null

    try {
        if (-not $DryRun) {
            Move-Item -LiteralPath $File.FullName -Destination $targetPath
            $status = 'Moved'
        }
    }
    catch {
        $status = 'Error'
        $errorMessage = $_.Exception.Message
    }

    Write-Log -LogFile $LogFile -Entry @{
        timestamp = (Get-Date).ToString('o')
        runId = $script:RunId
        mode = $Mode
        dryRun = $DryRun
        source = $File.FullName
        target = $targetPath
        category = $Category
        rule = $Rule
        status = $status
        error = $errorMessage
    }
}

function Test-IsExcludedPath {
    param([Parameter(Mandatory = $true)][string]$Path)

    $excludeRegex = '(\\|/)(AppData|Windows|Program Files( \(x86\))?|ProgramData|OneDriveTemp|OneDrive\\.*\\Cache)(\\|/|$)'
    return [regex]::IsMatch($Path, $excludeRegex, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
}

function Get-FilesFromSources {
    param(
        [Parameter(Mandatory = $true)][string[]]$Sources,
        [Parameter(Mandatory = $true)][string]$TargetRoot
    )

    $allFiles = New-Object System.Collections.Generic.List[System.IO.FileInfo]

    foreach ($src in $Sources) {
        if (-not (Test-Path -LiteralPath $src)) { continue }
        if (Test-IsExcludedPath -Path $src) { continue }

        $items = Get-ChildItem -LiteralPath $src -File -Recurse -Force -ErrorAction SilentlyContinue
        foreach ($item in $items) {
            if (Test-IsExcludedPath -Path $item.FullName) { continue }
            if ($item.FullName.StartsWith($TargetRoot, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
            $allFiles.Add($item)
        }
    }

    return $allFiles
}

function Invoke-Sorting {
    param(
        [Parameter(Mandatory = $true)][string[]]$Sources,
        [Parameter(Mandatory = $true)]$UserPaths,
        [Parameter(Mandatory = $true)]$Folders,
        [Parameter(Mandatory = $true)][bool]$DryRun,
        [Parameter(Mandatory = $true)][int]$MinFileAgeSeconds,
        [Parameter(Mandatory = $true)][string]$LogFile,
        [Parameter(Mandatory = $true)][string]$ModeName
    )

    $files = Get-FilesFromSources -Sources $Sources -TargetRoot $UserPaths.TargetRoot
    $now = Get-Date

    foreach ($file in $files) {
        $ext = $file.Extension.TrimStart('.').ToLowerInvariant()

        if ($ext -in @('tmp', 'crdownload', 'part')) {
            Write-Log -LogFile $LogFile -Entry @{
                timestamp = (Get-Date).ToString('o')
                runId = $script:RunId
                mode = $ModeName
                dryRun = $DryRun
                source = $file.FullName
                target = $null
                category = 'Ignored'
                rule = 'Ignore temp/incomplete extensions'
                status = 'Skipped'
                error = $null
            }
            continue
        }

        if ($file.DirectoryName.StartsWith($UserPaths.Downloads, [System.StringComparison]::OrdinalIgnoreCase)) {
            $ageSeconds = ($now - $file.LastWriteTime).TotalSeconds
            if ($ageSeconds -lt $MinFileAgeSeconds) {
                Write-Log -LogFile $LogFile -Entry @{
                    timestamp = (Get-Date).ToString('o')
                    runId = $script:RunId
                    mode = $ModeName
                    dryRun = $DryRun
                    source = $file.FullName
                    target = $null
                    category = 'Ignored'
                    rule = "Downloads protection (< $MinFileAgeSeconds seconds)"
                    status = 'Skipped'
                    error = $null
                }
                continue
            }
        }

        $category = Get-CategoryByExtension -Extension $file.Extension
        $targetFolder = Get-TargetPath -File $file -Category $category -Folders $Folders

        if ($file.DirectoryName.Equals($targetFolder, [System.StringComparison]::OrdinalIgnoreCase)) {
            Write-Log -LogFile $LogFile -Entry @{
                timestamp = (Get-Date).ToString('o')
                runId = $script:RunId
                mode = $ModeName
                dryRun = $DryRun
                source = $file.FullName
                target = $null
                category = $category
                rule = 'Already in target folder'
                status = 'Skipped'
                error = $null
            }
            continue
        }

        Move-FileSafely -File $file -DestinationFolder $targetFolder -DryRun $DryRun -LogFile $LogFile -Category $category -Rule 'Extension mapping + date bucket'
    }
}

function Run-InitialScan {
    param(
        [Parameter(Mandatory = $true)]$UserPaths,
        [Parameter(Mandatory = $true)]$Folders,
        [Parameter(Mandatory = $true)][bool]$DryRun,
        [Parameter(Mandatory = $true)][int]$MinFileAgeSeconds,
        [Parameter(Mandatory = $true)][string]$LogFile
    )

    Invoke-Sorting -Sources $UserPaths.Sources -UserPaths $UserPaths -Folders $Folders -DryRun $DryRun -MinFileAgeSeconds $MinFileAgeSeconds -LogFile $LogFile -ModeName 'Initial'
}

function Run-MaintenanceScan {
    param(
        [Parameter(Mandatory = $true)]$UserPaths,
        [Parameter(Mandatory = $true)]$Folders,
        [Parameter(Mandatory = $true)][bool]$DryRun,
        [Parameter(Mandatory = $true)][int]$MinFileAgeSeconds,
        [Parameter(Mandatory = $true)][string]$LogFile
    )

    $maintenanceSources = @($UserPaths.Desktop, $UserPaths.Downloads, $Folders.Inbox)
    Invoke-Sorting -Sources $maintenanceSources -UserPaths $UserPaths -Folders $Folders -DryRun $DryRun -MinFileAgeSeconds $MinFileAgeSeconds -LogFile $LogFile -ModeName 'Maintenance'
}

function Undo-LastRun {
    param(
        [Parameter(Mandatory = $true)][string]$LogsFolder,
        [Parameter(Mandatory = $true)][bool]$DryRun
    )

    $pointer = Join-Path $LogsFolder 'LastRun.txt'
    if (-not (Test-Path -LiteralPath $pointer)) {
        throw 'Kein LastRun-Zeiger gefunden. Undo nicht möglich.'
    }

    $lastLog = Get-Content -LiteralPath $pointer -Encoding UTF8 | Select-Object -First 1
    if (-not $lastLog -or -not (Test-Path -LiteralPath $lastLog)) {
        throw 'Letzte Logdatei nicht gefunden. Undo nicht möglich.'
    }

    $undoLog = Join-Path $LogsFolder ("undo_{0:yyyyMMdd_HHmmss}.jsonl" -f (Get-Date))

    $entries = Get-Content -LiteralPath $lastLog -Encoding UTF8 | ForEach-Object { $_ | ConvertFrom-Json }
    $movements = $entries | Where-Object { $_.status -eq 'Moved' }

    foreach ($entry in ($movements | Sort-Object timestamp -Descending)) {
        $status = 'Planned'
        $errorMessage = $null
        try {
            if (-not $DryRun) {
                $originalFolder = Split-Path -Path $entry.source -Parent
                Ensure-Folder -Path $originalFolder

                $undoTarget = $entry.source
                $counter = 1
                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($undoTarget)
                $ext = [System.IO.Path]::GetExtension($undoTarget)
                while (Test-Path -LiteralPath $undoTarget) {
                    $suffix = '_UNDO_{0:d3}' -f $counter
                    $undoTarget = Join-Path $originalFolder ($baseName + $suffix + $ext)
                    $counter++
                }

                Move-Item -LiteralPath $entry.target -Destination $undoTarget
                $status = 'MovedBack'
            }
        }
        catch {
            $status = 'Error'
            $errorMessage = $_.Exception.Message
        }

        Write-Log -LogFile $undoLog -Entry @{
            timestamp = (Get-Date).ToString('o')
            runId = $script:RunId
            mode = 'Undo'
            dryRun = $DryRun
            source = $entry.target
            target = $entry.source
            category = 'Undo'
            rule = 'Undo-LastRun'
            status = $status
            error = $errorMessage
        }
    }

    $undoPointer = Join-Path $LogsFolder 'LastUndo.txt'
    Set-Content -Path $undoPointer -Value $undoLog -Encoding UTF8
}

$userPaths = Get-UserPaths -TargetRootOverride $TargetRootOverride -SourceFoldersOverride $SourceFoldersOverride
$folders = Initialize-FolderStructure -TargetRoot $userPaths.TargetRoot
$script:CurrentLogFile = Join-Path $folders.Logs ("run_{0:yyyyMMdd_HHmmss}_{1}.jsonl" -f $script:RunStart, $Mode.ToLowerInvariant())

if ($Undo) {
    Undo-LastRun -LogsFolder $folders.Logs -DryRun $DryRun
    exit 0
}

switch ($Mode) {
    'Initial' {
        Run-InitialScan -UserPaths $userPaths -Folders $folders -DryRun $DryRun -MinFileAgeSeconds $MinFileAgeSeconds -LogFile $script:CurrentLogFile
    }
    'Maintenance' {
        Run-MaintenanceScan -UserPaths $userPaths -Folders $folders -DryRun $DryRun -MinFileAgeSeconds $MinFileAgeSeconds -LogFile $script:CurrentLogFile
    }
}

$lastRunPointer = Join-Path $folders.Logs 'LastRun.txt'
Set-Content -Path $lastRunPointer -Value $script:CurrentLogFile -Encoding UTF8

Write-Output "Completed. Mode=$Mode DryRun=$DryRun Log=$script:CurrentLogFile"
