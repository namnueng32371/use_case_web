# Auto-setup script for MySQL (run once when installing on a new machine).
# Must be run from PowerShell as Administrator.

Write-Host "=== Looking for installed MySQL Server ===" -ForegroundColor Cyan

$found = Get-ChildItem "C:\Program Files\MySQL" -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -like "MySQL Server*" } |
    Select-Object -First 1

if (-not $found) {
    Write-Host "MySQL Server was not found on this machine." -ForegroundColor Red
    Write-Host "Please install MySQL Server (choose 'Server only') from https://dev.mysql.com/downloads/installer/ first, then run this script again." -ForegroundColor Yellow
    exit 1
}

$mysqlDir = $found.FullName
$versionName = $found.Name
Write-Host "Found: $mysqlDir" -ForegroundColor Green

$dataDir = "C:\ProgramData\MySQL\$versionName\Data"
$iniPath = "C:\ProgramData\MySQL\$versionName\my.ini"
$serviceName = "MySQL84"

if (Test-Path $dataDir) {
    Write-Host "Data directory already exists at $dataDir - skipping initialization." -ForegroundColor Yellow
} else {
    Write-Host "=== Creating data directory (first time only) ===" -ForegroundColor Cyan
    New-Item -ItemType Directory -Force -Path $dataDir | Out-Null
    & "$mysqlDir\bin\mysqld.exe" --initialize-insecure --datadir="$dataDir" --basedir="$mysqlDir"
    Write-Host "Done." -ForegroundColor Green
}

Write-Host "=== Writing config file ===" -ForegroundColor Cyan
@"
[mysqld]
datadir=$($dataDir -replace '\\','/')
basedir=$($mysqlDir -replace '\\','/')
port=3306
"@ | Set-Content -Path $iniPath -Encoding ASCII
Write-Host "Config written to: $iniPath" -ForegroundColor Green

Write-Host "=== Installing MySQL as a Windows service (auto-starts on boot) ===" -ForegroundColor Cyan
$existingService = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
if ($existingService) {
    Write-Host "Service already installed. Starting it..." -ForegroundColor Yellow
    Start-Service -Name $serviceName -ErrorAction SilentlyContinue
} else {
    & "$mysqlDir\bin\mysqld.exe" --install $serviceName --defaults-file="$iniPath"
    net start $serviceName
}

Start-Sleep -Seconds 2
Write-Host "=== Checking result ===" -ForegroundColor Cyan
& "$mysqlDir\bin\mysqladmin.exe" -u root ping

Write-Host ""
Write-Host "Done! MySQL is ready and will start automatically every time this computer boots." -ForegroundColor Green
Write-Host "Next step: run schema.sql (or restore a backup .sql file) to create the database, then run server.py." -ForegroundColor Green
