# Directory/Subdirectory tree size reporter
# generated mostly by Copilot, tweaked by robryan
# (c)2026, GPL3 licensed, no guarantees whatsoever, run at your own risk
# USAGE: dirtreereport.ps1 -Path <"\\server\share" or disk path> -Threads <number of cpu threads> -OutputFormat <HTML or CSV>
# made to run in ps 5.1 on a basic win2022 server install

param(
    [Parameter(Mandatory=$true)]
    [string]$Path,

    [ValidateSet("CSV","HTML")]
    [string]$OutputFormat = "CSV",

    [int]$Threads = 8,

    [string]$OutputFile = "$(Get-Date -Format 'yyyyMMdd_HHmmss')_DirectoryReport.$($OutputFormat.ToLower())"
)

Add-Type -AssemblyName System.Collections

Write-Host "Indexing files under $Path ..."

# Single-pass file enumeration
$allFiles = Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue

# Unique folder list
$allFolders = $allFiles.Directory.FullName | Sort-Object -Unique
$total = $allFolders.Count
$counter = 0

Write-Host "Found $total folders. Starting $Threads threads..."

# Create runspace pool
$pool = [runspacefactory]::CreateRunspacePool(1, $Threads)
$pool.Open()

$jobs = New-Object System.Collections.ArrayList

foreach ($folder in $allFolders) {
    $counter++
    Write-Host ("[{0}/{1}] Queuing: {2}" -f $counter, $total, $folder)

    $ps = [powershell]::Create()
    $ps.RunspacePool = $pool

    $null = $ps.AddScript({
        param($folderPath)

        Write-Host "Processing: $folderPath"

        $files = Get-ChildItem -Path $folderPath -File -Recurse -ErrorAction SilentlyContinue
        $size  = ($files | Measure-Object Length -Sum).Sum

        [PSCustomObject]@{
            FolderPath = $folderPath
            SizeBytes  = $size
        }
    }).AddArgument($folder)

    $handle = $ps.BeginInvoke()
    $jobs.Add([pscustomobject]@{
        PowerShell = $ps
        Handle     = $handle
    }) | Out-Null
}

Write-Host "All tasks queued. Waiting for threads to finish..."

# Collect results
$results = foreach ($job in $jobs) {
    $output = $job.PowerShell.EndInvoke($job.Handle)
    $job.PowerShell.Dispose()
    $output
}

# Add root folder, but strip provider qualifier
$rootPath = (Resolve-Path $Path).ProviderPath
$rootSize = ($allFiles | Measure-Object Length -Sum).Sum
$results += [PSCustomObject]@{
    FolderPath = $rootPath
    SizeBytes  = $rootSize
}

# Normalize all FolderPath values to plain filesystem paths (no provider prefix)
$results = $results | ForEach-Object {
    $fp = $_.FolderPath -replace '^[^:]+::',''
    [PSCustomObject]@{
        FolderPath = $fp
        SizeBytes  = $_.SizeBytes
    }
}

# Correct depth calculation: per-path, not flattened
$maxDepth = ($results | ForEach-Object {
    ($_.FolderPath -split '[\\/]').Count
} | Measure-Object -Maximum).Maximum

Write-Host "Max folder depth detected: $maxDepth"

# Build final output with separate columns
$final = foreach ($item in $results) {
    $parts = $item.FolderPath -split '[\\/]'

    while ($parts.Count -lt $maxDepth) {
        $parts += ""
    }

    $obj = [ordered]@{}

    for ($i = 0; $i -lt $maxDepth; $i++) {
        $obj["Level$($i+1)"] = $parts[$i]
    }

    $obj["FolderPath"] = $item.FolderPath
    $obj["SizeBytes"]  = $item.SizeBytes

    $size = $item.SizeBytes
    $obj["SizeHuman"] = switch ($size) {
        {$_ -ge 1TB} { "{0:N2} TB" -f ($size/1TB); break }
        {$_ -ge 1GB} { "{0:N2} GB" -f ($size/1GB); break }
        {$_ -ge 1MB} { "{0:N2} MB" -f ($size/1MB); break }
        {$_ -ge 1KB} { "{0:N2} KB" -f ($size/1KB); break }
        default      { "$size bytes" }
    }

    [PSCustomObject]$obj
}

# Output
switch ($OutputFormat) {
    "CSV"  { $final | Export-Csv -Path $OutputFile -NoTypeInformation }
    "HTML" { $final | ConvertTo-Html | Out-File $OutputFile }
}

Write-Host "Report written to $OutputFile"
