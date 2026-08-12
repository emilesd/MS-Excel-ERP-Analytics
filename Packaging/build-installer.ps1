<#
.SYNOPSIS
    Builds MyOlap in Release and packages it into a single per-user installer EXE.

.DESCRIPTION
    Produces Packaging\output\MyOlap-Setup-<version>.exe containing:
      - the packed Excel-DNA XLL (64-bit and 32-bit Office variants)
      - the matching native SQLite library
      - the Excel-DNA host metadata files
      - the .NET 8 Desktop Runtime (x64) as a bundled prerequisite

    The .NET runtime is downloaded once into Packaging\prereq and reused.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File Packaging\build-installer.ps1
#>
[CmdletBinding()]
param(
    [switch]$SkipBuild,
    [string]$Iscc = "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
)

$ErrorActionPreference = 'Stop'

$packaging = $PSScriptRoot
$repo      = Split-Path $packaging -Parent
$project   = Join-Path $repo 'MyOlap\MyOlap.csproj'
$outDir    = Join-Path $repo 'MyOlap\bin\Release\net8.0-windows'
$payload   = Join-Path $packaging 'payload'
$prereq    = Join-Path $packaging 'prereq'
$runtime   = Join-Path $prereq 'windowsdesktop-runtime-8-win-x64.exe'

if (-not (Test-Path $Iscc)) {
    throw "Inno Setup compiler not found at '$Iscc'. Install it with: winget install JRSoftware.InnoSetup"
}

if (-not $SkipBuild) {
    Write-Host '[1/4] Building MyOlap (Release)...' -ForegroundColor Cyan
    & dotnet build $project -c Release --nologo -v minimal
    if ($LASTEXITCODE -ne 0) { throw 'Build failed.' }
} else {
    Write-Host '[1/4] Skipping build (-SkipBuild).' -ForegroundColor DarkGray
}

Write-Host '[2/4] Staging payload...' -ForegroundColor Cyan
New-Item -ItemType Directory -Force -Path $payload, $prereq, (Join-Path $packaging 'output') | Out-Null

$copies = @{
    'publish\MyOlap-AddIn64-packed.xll'        = 'MyOlap-AddIn64-packed.xll'
    'publish\MyOlap-AddIn-packed.xll'          = 'MyOlap-AddIn-packed.xll'
    'runtimes\win-x64\native\e_sqlite3.dll'    = 'e_sqlite3-x64.dll'
    'runtimes\win-x86\native\e_sqlite3.dll'    = 'e_sqlite3-x86.dll'
    'MyOlap.deps.json'                         = 'MyOlap.deps.json'
    'MyOlap-AddIn.deps.json'                   = 'MyOlap-AddIn.deps.json'
    'MyOlap-AddIn64.deps.json'                 = 'MyOlap-AddIn64.deps.json'
    'MyOlap.runtimeconfig.json'                = 'MyOlap.runtimeconfig.json'
}

foreach ($src in $copies.Keys) {
    $from = Join-Path $outDir $src
    if (-not (Test-Path $from)) { throw "Missing build artifact: $from" }
    Copy-Item $from (Join-Path $payload $copies[$src]) -Force
}

Copy-Item (Join-Path $repo 'MyOlap\UserGuide.md') (Join-Path $payload 'UserGuide.md') -Force
Get-ChildItem $payload -File | ForEach-Object { Unblock-File $_.FullName -ErrorAction SilentlyContinue }

Write-Host '[3/4] Ensuring .NET 8 Desktop Runtime prerequisite...' -ForegroundColor Cyan
if (Test-Path $runtime) {
    Write-Host "      Reusing $runtime" -ForegroundColor DarkGray
} else {
    Invoke-WebRequest -Uri 'https://aka.ms/dotnet/8.0/windowsdesktop-runtime-win-x64.exe' `
                      -OutFile $runtime -UseBasicParsing
}

Write-Host '[4/4] Compiling installer...' -ForegroundColor Cyan
Push-Location $packaging
try {
    & $Iscc 'MyOlap.iss'
    if ($LASTEXITCODE -ne 0) { throw 'Inno Setup compilation failed.' }
} finally {
    Pop-Location
}

Get-ChildItem (Join-Path $packaging 'output') -Filter '*.exe' |
    Select-Object Name, @{ n = 'SizeMB'; e = { [math]::Round($_.Length / 1MB, 1) } }, LastWriteTime |
    Format-Table -AutoSize
