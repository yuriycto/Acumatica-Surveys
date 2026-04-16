# Publishes the Survey26R1.zip customization to the local Acumatica instance
# at http://localhost/Survey using AcuPower.CustomizationTools.
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File .\publish-local.ps1

[CmdletBinding()]
param(
    [string]$BaseUrl       = "http://localhost/Survey",
    [string]$User          = "admin",
    [string]$Password      = "123",
    [string]$Company       = "Company",
    [string]$ProjectName   = "Survey26R1",
    [string]$ZipPath       = "D:\SourceCode\SurveyYZ\Survey26R1.zip",
    [string]$ToolDll       = "D:\SourceCode\AcuPower.CustomizationTools\bin\Debug\netstandard2.0\AcuPower.CustomizationTools.dll",
    [string]$NewtonsoftDll = "D:\SourceCode\AcuPower.CustomizationTools\bin\Release\netstandard2.0\Newtonsoft.Json.dll"
)

$ErrorActionPreference = "Stop"

Write-Host "Loading AcuPower.CustomizationTools ..." -ForegroundColor Cyan
Add-Type -Path $NewtonsoftDll
Add-Type -Path $ToolDll

Write-Host "Publishing '$ProjectName' from $ZipPath to $BaseUrl (tenant=$Company) ..." -ForegroundColor Cyan

$r = [AcuPower.CustomizationTools.PowerShell.AcumaticaCst]::ImportAndPublish(
    $BaseUrl, $User, $Password, $ProjectName, $ZipPath, $true, $Company)

Write-Host ""
Write-Host "IsCompleted: $($r.IsCompleted)  IsFailed: $($r.IsFailed)" -ForegroundColor Cyan

if ($r.Log) {
    Write-Host ""
    Write-Host "Log (last 40 entries):"
    $r.Log | Select-Object -Last 40 | ForEach-Object {
        "  [{0}] {1}" -f $_.Level, $_.Message
    }
}

if ($r.IsFailed) {
    Write-Host ""
    Write-Host "PUBLISH FAILED" -ForegroundColor Red
    exit 1
} else {
    Write-Host ""
    Write-Host "Publish OK" -ForegroundColor Green
}
