param (
    [Parameter(Mandatory = $true)]
    [Alias("SOW", "sow", "s")]
    [string]$SOWParam,

    [Parameter(Mandatory = $true)]
    [Alias("Year", "y")]
    [string]$YearParam
)

# Hardcoded paths
$LocalProjectRoot = "C:\Users\YourName\Projects"
$ShareDriveRootBase = "\\path\to\root"
$ReportTemplatePath = "\\sharedrive\templates\StatusReportTemplate.docx"

# Sanitize input
$CleanSOW = ($SOWParam -replace '\D', '')
$Year = ($YearParam -replace '\D', '')
$ShareDriveRoot = Join-Path $ShareDriveRootBase $Year

# Find the project folder
$ProjectFolder = Get-ChildItem -Path $ShareDriveRoot -Directory | Where-Object {
    $_.Name -match $CleanSOW
} | Select-Object -First 1

if (-not $ProjectFolder) {
    Write-Error "[!] No project folder found containing number $CleanSOW in $ShareDriveRoot"
    exit 1
}

# Destination path
$DestinationPath = Join-Path $LocalProjectRoot $ProjectFolder.Name

# Get files for progress tracking
$Files = Get-ChildItem -Path $ProjectFolder.FullName -Recurse -File
$Total = $Files.Count
$Counter = 0

# Ensure local directory exists
if (-not (Test-Path $DestinationPath)) {
    New-Item -Path $DestinationPath -ItemType Directory | Out-Null
}

# Copy with progress
foreach ($File in $Files) {
    $RelativePath = $File.FullName.Substring($ProjectFolder.FullName.Length)
    $TargetPath = Join-Path $DestinationPath $RelativePath
    $TargetDir = Split-Path $TargetPath -Parent

    if (-not (Test-Path $TargetDir)) {
        New-Item -Path $TargetDir -ItemType Directory -Force | Out-Null
    }

    Copy-Item -Path $File.FullName -Destination $TargetPath -Force

    $Counter++
    Write-Progress -Activity "[i] Copying project files..." -Status "$Counter of $Total files copied" -PercentComplete (($Counter / $Total) * 100)
}

# Shortcut
$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut("$DestinationPath\ShareDriveShortcut.lnk")
$Shortcut.TargetPath = $ProjectFolder.FullName
$Shortcut.Save()

# Copy and rename report template
$ReportTargetPath = Join-Path $DestinationPath "$CleanSOW`StatusReport.docx"
Copy-Item -Path $ReportTemplatePath -Destination $ReportTargetPath -Force

Write-Host "`n[*] Project '$($ProjectFolder.Name)' copied to '$DestinationPath'"
Write-Host "[*] Shortcut and '$CleanSOW`StatusReport.docx' added."
