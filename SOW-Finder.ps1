param (
    [Parameter(Mandatory = $true)]
    [Alias("sow", "s")]
    [string]$SOW,

    [Parameter(Mandatory = $true)]
    [Alias("year", "y")]
    [string]$Year
)

# Configurable base paths
$LocalProjectRoot = "C:\Users\YourName\Projects"
$ShareDriveRootBase = "\\path\to\root"
$ReportTemplatePath = "\\sharedrive\templates\StatusReportTemplate.docx"

# Sanitize input: digits only
$CleanSOW = ($SOW -replace '\D', '')

# Share drive path for given year
$ShareDriveRoot = Join-Path $ShareDriveRootBase $Year

# Find matching project folder
$ProjectFolder = Get-ChildItem -Path $ShareDriveRoot -Directory | Where-Object {
    $_.Name -match $CleanSOW
} | Select-Object -First 1

if (-not $ProjectFolder) {
    Write-Error "[!] No project folder found containing number $CleanSOW in $ShareDriveRoot"
    exit 1
}

# Full local destination path
$DestinationPath = Join-Path $LocalProjectRoot $ProjectFolder.Name

# Get all files to track progress
$Files = Get-ChildItem -Path $ProjectFolder.FullName -Recurse -File
$Total = $Files.Count
$Counter = 0

# Create destination folder if needed
if (-not (Test-Path $DestinationPath)) {
    New-Item -Path $DestinationPath -ItemType Directory | Out-Null
}

# Copy files with progress bar
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

# Create shortcut to original folder
$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut("$DestinationPath\ShareDriveShortcut.lnk")
$Shortcut.TargetPath = $ProjectFolder.FullName
$Shortcut.Save()

# Copy and rename report template
$ReportTargetPath = Join-Path $DestinationPath "$CleanSOW`StatusReport.docx"
Copy-Item -Path $ReportTemplatePath -Destination $ReportTargetPath -Force

Write-Host "`n [*] Project '$($ProjectFolder.Name)' copied to '$DestinationPath'"
Write-Host "[*] Shortcut and '$CleanSOW`StatusReport.docx' added."
