# Script cài đặt cục bộ cho Office Add-in dịch thuật JP-VI
# Vui lòng chạy script này với quyền Administrator

$utf8 = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = $utf8

# 1. Kiểm tra và yêu cầu quyền Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "==================================================================" -ForegroundColor Yellow
    Write-Host "Yêu cầu quyền Administrator để cài đặt chứng chỉ và Registry..." -ForegroundColor Yellow
    Write-Host "Đang tự động khởi chạy lại dưới quyền Administrator..." -ForegroundColor Yellow
    Write-Host "==================================================================" -ForegroundColor Yellow
    Start-Sleep -Seconds 1
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

Write-Host "==================================================================" -ForegroundColor Green
Write-Host "   BẮT ĐẦU CÀI ĐẶT BỘ DỊCH THUẬT JP-VI CỦA USER                   " -ForegroundColor Green
Write-Host "==================================================================" -ForegroundColor Green

# Xác định thư mục cài đặt AppData
$installDir = "$env:APPDATA\JP_VI_Translator"
Write-Host "Thư mục cài đặt mục tiêu: $installDir" -ForegroundColor Cyan

# Tạo thư mục cài đặt nếu chưa có
if (-not (Test-Path $installDir)) {
    New-Item -Path $installDir -ItemType Directory -Force | Out-Null
}

# Tắt tiến trình cũ để tránh lỗi khóa tệp tin
Write-Host "Đang dừng tiến trình dịch thuật cũ (nếu có)..." -ForegroundColor Yellow
Stop-Process -Name "agent_backend" -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

# Sao chép các tệp tin sang thư mục AppData
Write-Host "Đang sao chép tệp tin cài đặt sang AppData..." -ForegroundColor Yellow

# Copy dist
if (Test-Path "dist") {
    Copy-Item -Path "dist" -Destination $installDir -Recurse -Force
}
# Copy agent_backend.exe
if (Test-Path "agent_backend.exe") {
    Copy-Item -Path "agent_backend.exe" -Destination $installDir -Force
}
# Copy manifest.xml hoặc manifest-local.xml
if (Test-Path "manifest-local.xml") {
    Copy-Item -Path "manifest-local.xml" -Destination "$installDir\manifest.xml" -Force
} elseif (Test-Path "manifest.xml") {
    Copy-Item -Path "manifest.xml" -Destination $installDir -Force
}
# Copy run-server.bat
if (Test-Path "run-server.bat") {
    Copy-Item -Path "run-server.bat" -Destination $installDir -Force
}

# Di chuyển ngữ cảnh chạy vào thư mục AppData
Push-Location $installDir

# 2. Khởi chạy Server để tự động sinh chứng chỉ SSL (nếu chưa có)
if (-not (Test-Path "localhost.crt")) {
    Write-Host "`n[1/4] Đang khởi tạo chứng chỉ SSL (HTTPS) cho localhost..." -ForegroundColor Cyan
    if (Test-Path "agent_backend.exe") {
        Write-Host "Chạy agent_backend.exe..."
        $process = Start-Process .\agent_backend.exe -ArgumentList "--port 3000" -PassThru -WindowStyle Hidden
    } else {
        Write-Host "Chạy python agent_backend.py..."
        $process = Start-Process python -ArgumentList "agent_backend.py --port 3000" -PassThru -WindowStyle Hidden
    }
    # Chờ 3 giây để chương trình chạy và sinh file cert/key
    Start-Sleep -Seconds 3
    # Tắt tiến trình
    Stop-Process -Id $process.Id -Force
} else {
    Write-Host "`n[1/4] Đã tìm thấy chứng chỉ SSL localhost.crt có sẵn." -ForegroundColor Green
}

# 3. Cài đặt và tin cậy chứng chỉ
if (Test-Path "localhost.crt") {
    Write-Host "`n[2/4] Đang tin cậy chứng chỉ SSL cục bộ..." -ForegroundColor Cyan
    try {
        Import-Certificate -FilePath "localhost.crt" -CertStoreLocation "Cert:\LocalMachine\Root" -ErrorAction Stop
        Write-Host "Chứng chỉ SSL cục bộ đã được thêm vào Trusted Root Store thành công!" -ForegroundColor Green
    } catch {
        Write-Host "Lỗi khi nhập chứng chỉ bảo mật: $_" -ForegroundColor Red
        Write-Host "Bạn có thể tự nhập file 'localhost.crt' vào Trusted Root Certificate Authorities thủ công." -ForegroundColor Yellow
    }
} else {
    Write-Host "`n[2/4] Không thể tự động tạo chứng chỉ SSL. Có thể thiếu Python hoặc file thực thi backend." -ForegroundColor Red
}

# 4. Đăng ký Add-in với Microsoft Office qua Registry
Write-Host "`n[3/4] Đang đăng ký Add-in với Microsoft Office..." -ForegroundColor Cyan
$regPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
try {
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }
    
    # Tạo giá trị Registry trỏ tới thư mục chứa manifest.xml trong AppData
    $valueName = "JP_VI_Translator_Local"
    New-ItemProperty -Path $regPath -Name $valueName -Value $installDir -PropertyType String -Force | Out-Null
    
    Write-Host "Đăng ký Registry thành công!" -ForegroundColor Green
    Write-Host "Đã đăng ký thư mục: $installDir" -ForegroundColor Gray
} catch {
    Write-Host "Lỗi khi đăng ký Registry: $_" -ForegroundColor Red
}

# 5. Tạo các Shortcut chạy nhanh và khởi động cùng Windows
Write-Host "`n[4/4] Đang tạo các Shortcut..." -ForegroundColor Cyan
try {
    $WshShell = New-Object -ComObject WScript.Shell
    
    # Shortcut ngoài Desktop
    $DesktopPath = [System.IO.Path]::Combine([System.Environment]::GetFolderPath('Desktop'), "JP-VI Translator Server.lnk")
    $Shortcut = $WshShell.CreateShortcut($DesktopPath)
    $Shortcut.TargetPath = "$installDir\run-server.bat"
    $Shortcut.WorkingDirectory = $installDir
    $Shortcut.Description = "Start JP-VI Translator Backend Server"
    $Shortcut.Save()
    
    # Shortcut khởi động cùng Windows (Startup folder)
    $StartupPath = [System.IO.Path]::Combine([System.Environment]::GetFolderPath('Startup'), "JP-VI Translator Server.lnk")
    $StartupShortcut = $WshShell.CreateShortcut($StartupPath)
    $StartupShortcut.TargetPath = "$installDir\run-server.bat"
    $StartupShortcut.WorkingDirectory = $installDir
    $StartupShortcut.WindowStyle = 7 # Khởi chạy ở dạng thu nhỏ (minimized)
    $StartupShortcut.Description = "Auto Start JP-VI Translator Backend Server"
    $StartupShortcut.Save()
    
    Write-Host "Tạo các shortcut thành công!" -ForegroundColor Green
    Write-Host "  - Shortcut màn hình Desktop: $DesktopPath" -ForegroundColor Gray
    Write-Host "  - Shortcut khởi động cùng Windows (Startup): $StartupPath" -ForegroundColor Gray
} catch {
    Write-Host "Lỗi khi tạo shortcut: $_" -ForegroundColor Red
}

Pop-Location

Write-Host "`n==================================================================" -ForegroundColor Green
Write-Host "   CÀI ĐẶT HOÀN TẤT THÀNH CÔNG!                                  " -ForegroundColor Green
Write-Host "==================================================================" -ForegroundColor Green
Write-Host "Thông tin cài đặt:" -ForegroundColor Yellow
Write-Host "1. Toàn bộ chương trình được cài đặt tại: $installDir"
Write-Host "2. Server dịch thuật sẽ tự động khởi chạy (minimized) mỗi khi bạn mở máy tính."
Write-Host "3. Mở Excel, Word hoặc PowerPoint."
Write-Host "4. Vào tab 'Home' -> chọn 'Add-ins' (hoặc 'My Add-ins')."
Write-Host "5. Bạn sẽ thấy Add-in 'JP-VI Translator' xuất hiện trong nhóm 'Developer Add-ins'."
Write-Host "=================================================================="
Write-Host "Nhấn phím bất kỳ để thoát..."
Read-Host
