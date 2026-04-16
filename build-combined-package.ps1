# Builds a self-contained Survey26R1.zip that merges Base (SQL schemas,
# SUContent, sitemap, GIs, XportScenarios, ScreenWithRights) with
# my additions (5 Screen entries + 10 Modern UI per-tenant files) and
# the latest DLL.
#
# Output: e:\SourceCode\AcumaticaSurvey\Acumatica-Surveys\Survey26R1.zip

[CmdletBinding()]
param(
    [string]$RepoRoot    = "E:\SourceCode\AcumaticaSurvey\Acumatica-Surveys",
    [string]$ToolDll     = "E:\SourceCode\AcuPower.CustomizationTools\bin\Debug\netstandard2.0\AcuPower.CustomizationTools.dll",
    [string]$NewtonsoftDll = "E:\SourceCode\AcuPower.CustomizationTools\bin\Release\netstandard2.0\Newtonsoft.Json.dll",
    [string]$DllPath     = "E:\SourceCode\AcumaticaSurvey\Acumatica-Surveys\PX.Survey.Ext\bin\net48\PX.Survey.Ext.dll",
    [string]$OutZip      = "E:\SourceCode\AcumaticaSurvey\Acumatica-Surveys\Survey26R1.zip",
    [string]$ProductVersion = "26.100"
)

$ErrorActionPreference = "Stop"

$pkgRoot  = Join-Path $RepoRoot "Customizations\Package"
$projDir  = Join-Path $pkgRoot "_project"
$muiRoot  = Join-Path $pkgRoot "screens"

# --- 1. Copy MUI stubs from ModernUI\ into the package source ---
Remove-Item $muiRoot -Recurse -Force -ErrorAction SilentlyContinue
$muiSrc = Join-Path $RepoRoot "ModernUI\screens"
Copy-Item -Recurse -Force $muiSrc $muiRoot
Write-Host "Copied MUI files into $muiRoot"

# --- 2. Generate Screen and PerTenantFile XML items into _project\ ---
$screenIds = @("SU101000","SU201000","SU204003","SU301000","SU501000")

# Clean any old generated items
Get-ChildItem $projDir -Filter "Screen_SU*.xml" | Remove-Item -Force
Get-ChildItem $projDir -Filter "PerTenantFile_*.xml" | Remove-Item -Force

foreach ($id in $screenIds) {
    # Screen entry (registers the screen in the Customization Project Editor's Screens section).
    # Acumatica's CstScreen deserializer only reads the "ID" attribute; Type/Url are ignored.
    $screenXml = "<Screen ID=""$id"" />"
    Set-Content -Path (Join-Path $projDir "Screen_$id.xml") -Value $screenXml -Encoding UTF8

    # PerTenantFile XMLs are NOT written to _project/ - they come from
    # the imperative AddModernUiFile calls below to avoid duplication.
}
Write-Host "Generated Screen_*.xml (5) and PerTenantFile_*.xml (10) in _project\"

# --- 3. Build the ZIP using AcuPower.CustomizationTools ---
Add-Type -Path $NewtonsoftDll
Add-Type -Path $ToolDll

# The tool duplicates MUI files as <File> in addition to <PerTenantFile>.
# To avoid duplicates we hand-roll the ZIP: read ProjectMetadata for description,
# walk pkgRoot, and honor the file type.
$metadataPath = Join-Path $projDir "ProjectMetadata.xml"
[xml]$meta = Get-Content $metadataPath
$description = $meta.project.description

$builder = New-Object AcuPower.CustomizationTools.Builder.CustomizationProjectBuilder($description, $ProductVersion)

# Replace Bin DLL with the freshly built one
$replaceDll = @{
    (Join-Path $pkgRoot "Bin\PX.Survey.Ext.dll") = $DllPath
}

# Register each aspx as a <Page> so it shows up in the Customization Project
# Editor's Screens section. AddPage deflate/base64-encodes pageSource.
foreach ($id in $screenIds) {
    $aspxPath = Join-Path $pkgRoot "Pages\SU\$id.aspx"
    if (Test-Path $aspxPath) {
        $aspxText = Get-Content $aspxPath -Raw
        $builder.AddPage("~/Pages/SU/$id.aspx", $aspxText) | Out-Null
    } else {
        Write-Warning "Missing aspx for Page entry: $aspxPath"
    }
}

# Collect MUI files that should register as PerTenantFile, not File
$perTenantPaths = @{}
foreach ($id in $screenIds) {
    foreach ($ext in @("ts","html")) {
        $rel = "screens\SU\$id\$id.$ext"
        $perTenantPaths[$rel] = $id
    }
}

# Walk all files under pkgRoot except _project
$pkgRootInfo = Get-Item $pkgRoot
$projRootAbs = (Get-Item $projDir).FullName
foreach ($file in Get-ChildItem $pkgRoot -File -Recurse) {
    if ($file.FullName.StartsWith($projRootAbs, [StringComparison]::OrdinalIgnoreCase)) { continue }

    $rel = $file.FullName.Substring($pkgRootInfo.FullName.Length).TrimStart('\','/')
    # Possibly replace contents
    if ($replaceDll.ContainsKey($file.FullName)) {
        $bytes = [System.IO.File]::ReadAllBytes($replaceDll[$file.FullName])
    } else {
        $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
    }

    if ($perTenantPaths.ContainsKey($rel)) {
        $builder.AddModernUiFile($rel, $perTenantPaths[$rel], $bytes) | Out-Null
    } else {
        $builder.AddFile($rel, $bytes) | Out-Null
    }
}

# Inject every XML from _project/ into project.xml verbatim via a helper trick:
# builder has no raw-XML API. We'll build its ZIP, then re-open and patch project.xml.
$zipBytes = $builder.BuildPackageZip()

# --- 4. Patch project.xml to append _project XMLs ---
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Gather _project item XMLs in priority order (replicating tool's own ordering)
$priorityMap = @{
    "File"=0; "Page"=1; "Screen"=2; "PerTenantFile"=3; "Sql"=4;
    "SiteMapNode"=5; "ScreenWithRights"=6; "Webhook"=7; "Report"=8;
    "Dashboard"=9; "Wiki"=10; "GenericInquiry"=11; "GenericInquiryScreen"=11;
    "BpEvent"=12; "MobileSiteMap"=13; "XportScenario"=14
}
$items = Get-ChildItem $projDir -Filter "*.xml" |
    Where-Object { $_.Name -ne "ProjectMetadata.xml" } |
    ForEach-Object {
        $x = [xml](Get-Content $_.FullName -Raw)
        $root = $x.DocumentElement.LocalName
        $pri = if ($priorityMap.ContainsKey($root)) { $priorityMap[$root] } else { 100 }
        [pscustomobject]@{ Path=$_.FullName; Name=$_.Name; Priority=$pri; Xml=(Get-Content $_.FullName -Raw).Trim() }
    } |
    Sort-Object Priority, Name

# Open the built zip, read current project.xml, append items, rewrite project.xml
$ms = New-Object System.IO.MemoryStream
$ms.Write($zipBytes, 0, $zipBytes.Length)
$ms.Position = 0
$archive = New-Object System.IO.Compression.ZipArchive($ms, [System.IO.Compression.ZipArchiveMode]::Update, $true)
$projEntry = $archive.GetEntry("project.xml")
$reader = New-Object System.IO.StreamReader($projEntry.Open())
$existingXml = $reader.ReadToEnd()
$reader.Dispose()

# Insert additional items before the closing </project>
$additional = ($items | ForEach-Object { "  " + $_.Xml }) -join "`r`n"
$updatedXml = $existingXml -replace "</project>\s*$", "$additional`r`n</project>`r`n"

# Replace project.xml
$projEntry.Delete()
$newEntry = $archive.CreateEntry("project.xml", [System.IO.Compression.CompressionLevel]::Optimal)
$writer = New-Object System.IO.StreamWriter($newEntry.Open())
$writer.Write($updatedXml); $writer.Dispose()

$archive.Dispose()
[System.IO.File]::WriteAllBytes($OutZip, $ms.ToArray())
$ms.Dispose()

Write-Host ""
Write-Host "Built: $OutZip ($(([System.IO.FileInfo]$OutZip).Length) bytes)"
Write-Host "Contains:"
$zip = [System.IO.Compression.ZipFile]::OpenRead($OutZip)
[xml]$px = [string]($zip.GetEntry("project.xml").Open() | ForEach-Object { $sr = New-Object System.IO.StreamReader($_); $sr.ReadToEnd() })
$zip.Dispose()
$px.SelectNodes("/project/*") | Group-Object LocalName | Sort-Object Name | ForEach-Object {
    "  {0,-22} {1}" -f $_.Name, $_.Count
}
