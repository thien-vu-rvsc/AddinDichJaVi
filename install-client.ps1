# Install Client script for JP-VI Translator Add-in
# Copies the production manifest and registers it in Microsoft Office Trusted Catalogs Registry

$ErrorActionPreference = "Stop"

# 1. Define paths
$targetDir = "$env:LOCALAPPDATA\JP-VI-Translator"
$manifestName = "manifest-prod.xml"
$sourcePath = Join-Path $PSScriptRoot $manifestName

if (-not (Test-Path $sourcePath)) {
    Write-Error "Khong tim thay file $manifestName tai thu muc chay script."
}

# 2. Create target directory and copy manifest
if (-not (Test-Path $targetDir)) {
    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
}
Copy-Item -Path $sourcePath -Destination (Join-Path $targetDir $manifestName) -Force

Write-Host "Da sao chep manifest den: $targetDir" -ForegroundColor Green

# 3. Write to Windows Registry for Office Trust Center
$registryPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"
$guid = "{d0efe418-8186-4b70-aafc-6c8265846579}" # Matches the Add-in ID
$catalogPath = Join-Path $registryPath $guid

# Ensure parent key exists
if (-not (Test-Path $registryPath)) {
    $officePath = "HKCU:\Software\Microsoft\Office"
    if (-not (Test-Path "$officePath\16.0")) {
        New-Item -Path $officePath -Name "16.0" -Force | Out-Null
    }
    if (-not (Test-Path "$officePath\16.0\WEF")) {
        New-Item -Path "$officePath\16.0" -Name "WEF" -Force | Out-Null
    }
    New-Item -Path "$officePath\16.0\WEF" -Name "TrustedCatalogs" -Force | Out-Null
}

# Create/Overwrite catalog key
if (-not (Test-Path $catalogPath)) {
    New-Item -Path $registryPath -Name $guid -Force | Out-Null
}

# Set registry values
Set-ItemProperty -Path $catalogPath -Name "Url" -Value $targetDir
Set-ItemProperty -Path $catalogPath -Name "Flags" -Value 1 -Type DWord
Set-ItemProperty -Path $catalogPath -Name "Type" -Value 1 -Type DWord

Write-Host "Da dang ky thu muc tin cay vao Windows Registry cho Office." -ForegroundColor Green
Write-Host ""
Write-Host "Cai dat thanh cong! Vui long lam theo huong dan sau de kich hoat Add-in:" -ForegroundColor Cyan
Write-Host "1. Mo ung dung Word hoac Excel."
Write-Host "2. Vao the Home (Trang chu) -> Chon Add-ins (hoac My Add-ins)."
Write-Host "3. Chon muc Shared Folder (hoac Thu muc chia se)."
Write-Host "4. Chon 'JP-VI Translator' va bam Add (hoac Them)."
Write-Host ""
Write-Host "Nhan phim bat ky de thoat..."
