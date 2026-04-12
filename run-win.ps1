param(
    [string]$Configuration = "Debug",
    [string]$Framework = "net9.0-windows10.0.19041.0",
    [switch]$Clean,
    [switch]$Incremental,
    [switch]$NoLaunch
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = $PSScriptRoot
$projectPath = Join-Path $repoRoot "PDFTemplateGenerator\PDFTemplateGenerator.csproj"
$projectDir = Split-Path $projectPath -Parent
$outputExe = Join-Path $repoRoot "PDFTemplateGenerator\bin\$Configuration\$Framework\win10-x64\PDFTemplateGenerator.exe"

if (-not (Test-Path $projectPath)) {
    throw "Project file not found: $projectPath"
}

function Get-ChildProcessIds {
    param([int[]]$ParentIds)

    $queue = [System.Collections.Generic.Queue[int]]::new()
    $visited = [System.Collections.Generic.HashSet[int]]::new()

    foreach ($id in $ParentIds) {
        $queue.Enqueue($id)
        [void]$visited.Add($id)
    }

    $allIds = [System.Collections.Generic.List[int]]::new()
    while ($queue.Count -gt 0) {
        $currentId = $queue.Dequeue()
        $children = Get-CimInstance Win32_Process -Filter "ParentProcessId = $currentId" -ErrorAction SilentlyContinue
        foreach ($child in $children) {
            $childId = [int]$child.ProcessId
            if (-not $visited.Contains($childId)) {
                [void]$visited.Add($childId)
                $queue.Enqueue($childId)
                $allIds.Add($childId)
            }
        }
    }

    return $allIds
}

Write-Host "Stopping running PDFTemplateGenerator process tree..."
$appProcesses = Get-Process PDFTemplateGenerator -ErrorAction SilentlyContinue
if ($appProcesses) {
    $rootIds = @($appProcesses | Select-Object -ExpandProperty Id)
    $childIds = Get-ChildProcessIds -ParentIds $rootIds
    $idsToStop = @($rootIds + $childIds | Sort-Object -Unique)

    foreach ($processId in ($idsToStop | Sort-Object -Descending)) {
        $proc = Get-Process -Id $processId -ErrorAction SilentlyContinue
        if ($proc) {
            Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
        }
    }
}

if ($Clean) {
    Write-Host "Cleaning bin/obj artifacts..."
    Remove-Item -Recurse -Force (Join-Path $repoRoot "PDFTemplateGenerator\bin") -ErrorAction SilentlyContinue
    Remove-Item -Recurse -Force (Join-Path $repoRoot "PDFTemplateGenerator\obj") -ErrorAction SilentlyContinue
}

$buildArgs = @(
    "build",
    $projectPath,
    "-f", $Framework,
    "-c", $Configuration,
    "-p:WindowsPackageType=None",
    "-p:GenerateAppInstallerFile=false",
    "-p:AppxPackageSigningEnabled=false"
)

if (-not $Incremental) {
    $buildArgs += "--no-incremental"
}

Write-Host "Running: dotnet $($buildArgs -join ' ')"
$maxBuildAttempts = 3
$buildSucceeded = $false

for ($attempt = 1; $attempt -le $maxBuildAttempts; $attempt++) {
    if ($attempt -gt 1) {
        Write-Host "Retrying build (attempt $attempt of $maxBuildAttempts)..."
    }

    $buildOutput = & dotnet @buildArgs 2>&1
    $exitCode = $LASTEXITCODE

    foreach ($line in $buildOutput) {
        Write-Host $line
    }

    if ($exitCode -eq 0) {
        $buildSucceeded = $true
        break
    }

    $combinedOutput = ($buildOutput | Out-String)
    $isFileLockError = $combinedOutput -match "being used by another process"

    if ($isFileLockError -and $attempt -lt $maxBuildAttempts) {
        Write-Host "Detected transient file lock during static web asset compression. Cleaning compressed outputs before retry..."
        $compressedDir = Join-Path $projectDir "obj\$Configuration\$Framework\win10-x64\compressed"
        Remove-Item -Recurse -Force $compressedDir -ErrorAction SilentlyContinue
        continue
    }

    throw "Build failed with exit code $exitCode"
}

if (-not $buildSucceeded) {
    throw "Build failed after $maxBuildAttempts attempts"
}

if ($NoLaunch) {
    Write-Host "Build completed. Skipping launch because -NoLaunch was provided."
    exit 0
}

if (-not (Test-Path $outputExe)) {
    throw "Expected executable not found: $outputExe"
}

Write-Host "Launching app..."
Start-Process -FilePath $outputExe -WorkingDirectory (Split-Path $outputExe -Parent)
Write-Host "App started: $outputExe"
