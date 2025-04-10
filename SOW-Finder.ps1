param (
    [Parameter(Mandatory = $true)]
    [string]$SOWNumberRaw
)

# Hardcoded paths
$LocalProjectRoot = "C:\Users\YourName\Projects"
$ShareDriveRoot = "\\sharedrive\projects"
$ReportTemplatePath = "\\sharedrive\templates\StatusReportTemplate.docx"  # Update if needed

# Sanitize input: keep only digits
$SOWNumber = ($SOWNumberRaw -replace '\D', '')

# Find matching project folder
$ProjectFolder = Get-ChildItem -Path $ShareDriveRoot -Directory | Where-Object {
    $_.Name -match $SOWNumber
} | Select-Object -First 1

if (-not $ProjectFolder) {
    Write-Error "❌ No project folder found containing number $SOWNumber"
    exit 1
}

# Full destination path
$DestinationPath = Join-Path -Path $LocalProjectRoot -ChildPath $ProjectFolder.Name

# Get all files for progress tracking
$Files = Get-ChildItem -Path $ProjectFolder.FullName -Recurse -File
$Total = $Files.Count
$Counter = 0

# Ensure destination folder exists
if (-not (Test-Path $DestinationPath)) {
    New-Item -Path $DestinationPath -ItemType Directory | Out-Null
}

# Copy files with progress
foreach ($File in $Files) {
    $RelativePath = $File.FullName.Substring($ProjectFolder.FullName.Length)
    $TargetPath = Join-Path -Path $DestinationPath -ChildPath $RelativePath
    $TargetDir = Split-Path -Path $TargetPath -Parent

    if (-not (Test-Path $TargetDir)) {
        New-Item -Path $TargetDir -ItemType Directory -Force | Out-Null
    }

    Copy-Item -Path $File.FullName -Destination $TargetPath -Force

    $Counter++
    Write-Progress -Activity "Copying project files..." -Status "$Counter of $Total files copied" -PercentComplete (($Counter / $Total) * 100)
}

# Create shortcut to shared project folder
$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut("$DestinationPath\ShareDriveShortcut.lnk")
$Shortcut.TargetPath = $ProjectFolder.FullName
$Shortcut.Save()

# Copy and rename status report template
$ReportTargetPath = Join-Path $DestinationPath "$SOWNumber`StatusReport.docx"
Copy-Item -Path $ReportTemplatePath -Destination $ReportTargetPath -Force

Write-Host "Project '$($ProjectFolder.Name)' copied successfully to '$DestinationPath'"
Write-Host "Shortcut and status report added."
