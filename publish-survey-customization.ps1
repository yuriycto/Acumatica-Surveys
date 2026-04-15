# Publishes the Acumatica Survey customization to the local instance
# using the AcuPower.CustomizationTools library.
#
# Requirements:
#   - Acumatica site running at $BaseUrl
#   - Admin credentials valid for the tenant
#   - Built DLL at $DllPath

[CmdletBinding()]
param(
    [string]$BaseUrl      = "http://localhost/SurveyProject",
    [string]$User         = "admin",
    [string]$Password     = "123",
    [string]$Company      = "Company",
    [string]$ProjectName  = "Survey26R1",
    [string]$RepoRoot     = "E:\SourceCode\AcumaticaSurvey\Acumatica-Surveys",
    [string]$ToolDllPath  = "E:\SourceCode\AcuPower.CustomizationTools\bin\Debug\netstandard2.0\AcuPower.CustomizationTools.dll",
    [string]$DllPath      = "E:\SourceCode\AcumaticaSurvey\Acumatica-Surveys\PX.Survey.Ext\bin\net48\PX.Survey.Ext.dll"
)

$ErrorActionPreference = "Stop"

# Load the AcuPower tool
Add-Type -Path "E:\SourceCode\AcuPower.CustomizationTools\bin\Release\netstandard2.0\Newtonsoft.Json.dll"
Add-Type -Path $ToolDllPath

# Screens we are registering
$screens = @(
    @{ Id = "SU101000"; Aspx = "Pages\SU101000.aspx"; Graph = "PX.Survey.Ext.SurveySetupMaint";     Title = "Survey Preferences"  },
    @{ Id = "SU201000"; Aspx = "Pages\SU201000.aspx"; Graph = "PX.Survey.Ext.SurveyMaint";          Title = "Survey"              },
    @{ Id = "SU204003"; Aspx = "Pages\SU204003.aspx"; Graph = "PX.Survey.Ext.SurveyComponentMaint"; Title = "Survey Components"   },
    @{ Id = "SU301000"; Aspx = "Pages\SU301000.aspx"; Graph = "PX.Survey.Ext.SurveyCollectorMaint"; Title = "Survey Collector"    },
    @{ Id = "SU501000"; Aspx = "Pages\SU501000.aspx"; Graph = "PX.Survey.Ext.SurveyProcess";        Title = "Process Survey"      }
)

# Build the project with the fluent builder
$builder = New-Object AcuPower.CustomizationTools.Builder.CustomizationProjectBuilder(
    "Acumatica Surveys 26R1",
    "26.100")

# --- DLLs under Bin (main + deps) ---
$binDir = Split-Path $DllPath -Parent
$binFiles = @(
    $DllPath,
    (Join-Path $binDir "Scriban.dll")
    # System.Runtime.CompilerServices.Unsafe / System.Threading.Tasks.Extensions
    # already exist in the Acumatica Bin with newer versions - do not override.
)
foreach ($bin in $binFiles) {
    if (Test-Path $bin) {
        $name = [System.IO.Path]::GetFileName($bin)
        $builder.AddFile("Bin\$name", $bin) | Out-Null
        Write-Host "Added: Bin\$name"
    }
}

# --- Pages + Screen registrations + Modern UI files ---
foreach ($s in $screens) {
    $aspxDisk     = Join-Path $RepoRoot $s.Aspx
    $aspxContent  = [System.IO.File]::ReadAllText($aspxDisk)
    $aspxVirtPath = "~/Pages/SU/$($s.Id).aspx"

    # Page entry
    $builder.AddPage($aspxVirtPath, $aspxContent) | Out-Null
    # Screen entry (this is what puts it in the Screens section of the editor)
    $builder.AddScreen($s.Id) | Out-Null

    # Modern UI stubs
    foreach ($ext in @("ts", "html")) {
        $muiRel  = "screens\SU\$($s.Id)\$($s.Id).$ext"
        $muiDisk = Join-Path $RepoRoot "ModernUI\$muiRel"
        if (Test-Path $muiDisk) {
            $bytes = [System.IO.File]::ReadAllBytes($muiDisk)
            $builder.AddModernUiFile($muiRel, $s.Id, $bytes) | Out-Null
        }
    }

    # Sitemap node so the screen is reachable
    $builder.AddSiteMapNode(
        $s.Id, $s.Title, $aspxVirtPath, $s.Graph, "D", $null, $null) | Out-Null

    Write-Host "Added screen: $($s.Id) - $($s.Title)"
}

# --- Webhook ---
$builder.AddWebhook(
    "Survey",
    "PX.Survey.Ext.WebHook.SurveyWebhookServerHandler",
    "9d2fd3bc-aae4-4b34-ab66-f40879be1be2",
    $true, $false, 2, 100, $null) | Out-Null

# Build ZIP
$zipBytes = $builder.BuildPackageZip()
$zipPath  = Join-Path $env:TEMP "$ProjectName.zip"
[System.IO.File]::WriteAllBytes($zipPath, $zipBytes)
Write-Host ""
Write-Host "Built package: $zipPath ($($zipBytes.Length) bytes)"

# --- Publish ---
Write-Host ""
Write-Host "Importing and publishing to $BaseUrl ..."
$result = [AcuPower.CustomizationTools.PowerShell.AcumaticaCst]::ImportAndPublish(
    $BaseUrl, $User, $Password, $ProjectName, $zipPath, $true, $Company)

Write-Host ""
Write-Host "=== Publish result ==="
Write-Host "IsCompleted: $($result.IsCompleted)"
Write-Host "IsFailed:    $($result.IsFailed)"
Write-Host "--- Log ---"
$result.Log | ForEach-Object {
    "{0}  {1}  {2}" -f $_.Timestamp, $_.Level, $_.Message
}
