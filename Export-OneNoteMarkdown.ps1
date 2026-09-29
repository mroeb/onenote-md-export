<#
.SYNOPSIS
  Export local OneNote notebooks to a Markdown folder tree.

.DESCRIPTION
  Thin wrapper around onenote-md.exe so you can call it from PowerShell.
  Builds the executable on first use if it is missing.

.EXAMPLE
  .\Export-OneNoteMarkdown.ps1 -List
.EXAMPLE
  .\Export-OneNoteMarkdown.ps1 -OutputDir D:\notes -Notebook Reports
#>
[CmdletBinding()]
param(
    # Defaults to a path outside the repository so a stray run cannot leave
    # private notebook content inside the project tree.
    [string]   $OutputDir   = (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'onenote-export'),
    [string]   $Notebook,
    [string]   $Section,
    [switch]   $List,
    [switch]   $NoImages,
    [switch]   $SkipExisting,
    [switch]   $DryRun
)

$ErrorActionPreference = 'Stop'

$exe = Join-Path $PSScriptRoot 'bin\onenote-md.exe'
if (-not (Test-Path $exe)) {
    Write-Host "Building onenote-md.exe ..." -ForegroundColor Cyan
    & (Join-Path $PSScriptRoot 'build.cmd')
    if ($LASTEXITCODE -ne 0) { throw "build.cmd failed" }
}

$argList = @($OutputDir)
if ($Notebook)      { $argList += @('--notebook', $Notebook) }
if ($Section)       { $argList += @('--section',  $Section)  }
if ($List)          { $argList += '--list' }
if ($NoImages)      { $argList += '--no-images' }
if ($SkipExisting)  { $argList += '--skip-existing' }
if ($DryRun)        { $argList += '--dry-run' }

& $exe @argList
exit $LASTEXITCODE
