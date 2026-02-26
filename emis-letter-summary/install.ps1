Param(
  [string]$PythonVersion = "3.11"
)

$ErrorActionPreference = "Stop"
$toolDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $toolDir

py -$PythonVersion -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt

Write-Host "Setup complete."
Write-Host "Run with: .\\run.bat --help"
