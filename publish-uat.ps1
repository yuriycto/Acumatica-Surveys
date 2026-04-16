# End-to-end publish of both customer packages against a vanilla sales demo.
#
# Publishes:
#   1) Survey26R1            (functionality + SQL schemas + SUContent + MUI)
#   2) Survey26R1SalesDemo   (starter data: Numbering, SurveySetup, demo survey)
#
# Against:
#   http://localhost/SurveyUAT  (admin / 123, tenant Company)

[CmdletBinding()]
param(
    [string]$BaseUrl      = "http://localhost/SurveyUAT",
    [string]$User         = "admin",
    [string]$Password     = "123",
    [string]$Company      = "Company",
    [string]$RepoRoot     = "E:\SourceCode\AcumaticaSurvey\Acumatica-Surveys",
    [string]$ToolDll      = "E:\SourceCode\AcuPower.CustomizationTools\bin\Debug\netstandard2.0\AcuPower.CustomizationTools.dll",
    [string]$NewtonsoftDll = "E:\SourceCode\AcuPower.CustomizationTools\bin\Release\netstandard2.0\Newtonsoft.Json.dll"
)

$ErrorActionPreference = "Stop"
Add-Type -Path $NewtonsoftDll
Add-Type -Path $ToolDll

function Publish($projectName, $zipPath) {
    Write-Host ""
    Write-Host "===== Publishing $projectName =====" -ForegroundColor Cyan
    Write-Host "Zip: $zipPath"
    $r = [AcuPower.CustomizationTools.PowerShell.AcumaticaCst]::ImportAndPublish(
        $BaseUrl, $User, $Password, $projectName, $zipPath, $true, $Company)
    Write-Host "IsCompleted: $($r.IsCompleted)  IsFailed: $($r.IsFailed)"
    if ($r.Log) {
        $r.Log | Select-Object -Last 20 | ForEach-Object {
            "  [{0}] {1}" -f $_.Level, $_.Message
        }
    }
    if ($r.IsFailed) { Write-Host "FAILED" -ForegroundColor Red } else { Write-Host "OK" -ForegroundColor Green }
    return $r
}

$r1 = Publish "Survey26R1"            (Join-Path $RepoRoot "Survey26R1.zip")
if ($r1.IsFailed) { exit 1 }
$r2 = Publish "Survey26R1SalesDemo"   (Join-Path $RepoRoot "Survey26R1.SalesDemo.zip")
if ($r2.IsFailed) { exit 1 }

Write-Host ""
Write-Host "Both packages published successfully." -ForegroundColor Green
