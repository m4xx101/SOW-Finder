param (
    [Parameter(Mandatory=$true)]
    [string]$SOWNumberRaw
)

# Hardcoded local destination base path
$LocalProjectRoot = "C:\Users\YourName\Projects"

# Sanitize SOW input: keep only digits
$SOWNumber = ($SOWNumberRaw -replace '\D', '')

# Define share drive root path
$ShareDriveRoot = "\\sharedrive\projects"

# Search for folder with matching SOW number in name
$ProjectFolder = Get-ChildItem -Path $ShareDriveRoot -Directory | Where-Object {
    $_.Name -match $SOWNumber
} | Select-Object -First 1

if (-not $ProjectFolder) {
    Write-Error "❌ No project folder found containing number $SOWNumber"
    exit 1
}

# Full local path
$DestinationPath = Join-Path -Path $LocalProjectRoot -ChildPath $ProjectFolder.Name

# Copy entire folder to local path
Copy-Item -Path $ProjectFolder.FullName -Destination $DestinationPath -Recurse

# Create shortcut to original folder
$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut("$DestinationPath\ShareDriveShortcut.lnk")
$Shortcut.TargetPath = $ProjectFolder.FullName
$Shortcut.Save()

Write-Host "✅ Project '$($ProjectFolder.Name)' copied to '$DestinationPath'"
Write-Host "🔗 Shortcut to shared drive folder created."
