$chrome = "$env:ProgramFiles\Google\Chrome\Application\chrome.exe"
$profile = "C:\AIworkspace\autocretop\chrome_profile"
$port = 9222

if (-not (Test-Path $chrome)) {
    Write-Host "Chrome을 찾을 수 없습니다." -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $profile)) {
    New-Item -ItemType Directory -Path $profile | Out-Null
}

Write-Host "CRETOP 전용 Chrome을 실행합니다..." -ForegroundColor Cyan
Write-Host "Profile : $profile"
Write-Host "Port    : $port"

Start-Process $chrome -ArgumentList @(
    "--remote-debugging-port=$port",
    "--user-data-dir=$profile"
)

Write-Host ""
Write-Host "CRETOP 전용 Chrome 실행 완료." -ForegroundColor Green
Write-Host "Playwright CDP: http://127.0.0.1:$port"