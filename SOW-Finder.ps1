param (
    [Parameter(Mandatory = $true)]
    [Alias("s", "sw")]
    [string]$SOWParam,

    [Parameter(Mandatory = $false)]
    [Alias("y", "yr")]
    [string]$YearParam = (Get-Date).Year.ToString()
)

# Hardcoded paths
$LocalProjectRoot = "C:\Users\YourName\Projects"
$ShareDriveRootBase = "\\path\to\root"
$ReportTemplatePath = "\\sharedrive\templates\StatusReportTemplate.docx"

# Sanitize inputs
$CleanSOW = ($SOWParam -replace '\D', '')
$Year = ($YearParam -replace '\D', '')
$ShareDriveRoot = Join-Path $ShareDriveRootBase $Year

# Confirm share drive year folder exists
if (-not (Test-Path $ShareDriveRoot)) {
    Write-Error "[!] The share year folder '$ShareDriveRoot' does not exist."
    exit 1
}

# Look for project folder containing the SOW
$ProjectFolder = Get-ChildItem -Path $ShareDriveRoot -Directory | Where-Object {
    $_.Name -match $CleanSOW
} | Select-Object -First 1

if (-not $ProjectFolder) {
    Write-Error "[!] No project folder found containing number $CleanSOW in $ShareDriveRoot"
    exit 1
}

# Create destination path with "SOW_<SOW>" format
$DestinationFolderName = "SOW_$CleanSOW"
$DestinationPath = Join-Path $LocalProjectRoot $DestinationFolderName

# Get all files from share folder for progress
$Files = Get-ChildItem -Path $ProjectFolder.FullName -Recurse -File
$Total = $Files.Count
$Counter = 0

# Create local folder if needed
if (-not (Test-Path $DestinationPath)) {
    New-Item -Path $DestinationPath -ItemType Directory | Out-Null
}

# Copy files with progress
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

# Create shortcut to share drive project folder
$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut("$DestinationPath\ShareDriveShortcut.lnk")
$Shortcut.TargetPath = $ProjectFolder.FullName
$Shortcut.Save()

# Copy status report template and rename
$ReportTargetPath = Join-Path $DestinationPath "$CleanSOW`StatusReport.docx"
Copy-Item -Path $ReportTemplatePath -Destination $ReportTargetPath -Force

Write-Host "`n[*] Project '$($ProjectFolder.Name)' copied to '$DestinationPath'"
Write-Host "[*] Shortcut and status report '$($CleanSOW)StatusReport.docx' created."
