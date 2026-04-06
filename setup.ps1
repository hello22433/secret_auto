# =============================================
# 신입 PC 자동 설정 스크립트
# GitHub에서 설정을 자동으로 받아옵니다
# =============================================

# === 파라미터 (run.bat에서 전달됨) ===
param(
    [string]$GITHUB_OWNER = "hello22433",
    [string]$GITHUB_REPO  = "secret_auto",
    [string]$GITHUB_TOKEN = ""
)

if (-not $GITHUB_TOKEN) {
    Write-Host "[ERROR] GitHub 토큰이 설정되지 않았습니다." -ForegroundColor Red
    Write-Host "run.bat 파일의 TOKEN 값을 확인하세요." -ForegroundColor Red
    Read-Host "아무 키나 누르면 종료합니다"
    exit 1
}

# UTF-8 출력 설정
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# --- 관리자 권한 확인 및 재실행 ---
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "[!] 관리자 권한이 필요합니다. 재실행합니다..." -ForegroundColor Yellow
    Start-Process powershell.exe -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`" -GITHUB_TOKEN `"$GITHUB_TOKEN`""
    exit
}

# --- GitHub에서 config.json 다운로드 ---
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "   신입 PC 자동 설정 도구" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "최신 설정을 다운로드 중..." -ForegroundColor Yellow -NoNewline

try {
    $headers = @{
        "Authorization" = "token $GITHUB_TOKEN"
        "Accept"        = "application/vnd.github.v3.raw"
    }
    $configUrl = "https://api.github.com/repos/$GITHUB_OWNER/$GITHUB_REPO/contents/config.json"
    $configRaw = Invoke-RestMethod -Uri $configUrl -Headers $headers
    $config = $configRaw | ConvertFrom-Json
    Write-Host " OK" -ForegroundColor Green
}
catch {
    Write-Host " FAIL" -ForegroundColor Red
    Write-Host "  설정 다운로드 실패: $($_.Exception.Message)" -ForegroundColor DarkGray
    Read-Host "아무 키나 누르면 종료합니다"
    exit 1
}

# --- 사용자 정보 입력 ---
Write-Host ""
$userName = Read-Host "이름을 입력하세요"
$department = Read-Host "부서를 입력하세요"
$pcName = $env:COMPUTERNAME

Write-Host ""
Write-Host "이름: $userName / 부서: $department / PC: $pcName" -ForegroundColor Green
Write-Host "설정을 시작합니다..." -ForegroundColor Green
Write-Host ""

# --- 결과 저장용 ---
$results = @()
$allSuccess = $true
$startTime = Get-Date

# --- 명령어 순차 실행 ---
$total = $config.commands.Count
$current = 0

foreach ($cmd in $config.commands) {
    $current++
    $stepResult = @{
        name    = $cmd.name
        status  = "pending"
        output  = ""
        error   = ""
    }

    Write-Host "[$current/$total] $($cmd.name)..." -ForegroundColor Yellow -NoNewline

    try {
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo.FileName = "cmd.exe"
        $process.StartInfo.Arguments = "/c $($cmd.command)"
        $process.StartInfo.RedirectStandardOutput = $true
        $process.StartInfo.RedirectStandardError = $true
        $process.StartInfo.UseShellExecute = $false
        $process.StartInfo.CreateNoWindow = $true
        $process.StartInfo.StandardOutputEncoding = [System.Text.Encoding]::GetEncoding(949)
        $process.StartInfo.StandardErrorEncoding = [System.Text.Encoding]::GetEncoding(949)

        $process.Start() | Out-Null

        $timeoutMs = $cmd.timeout * 1000
        $exited = $process.WaitForExit($timeoutMs)

        $output = $process.StandardOutput.ReadToEnd()
        $errorOutput = $process.StandardError.ReadToEnd()

        if (-not $exited) {
            $process.Kill()
            $stepResult.status = "timeout"
            $stepResult.error = "시간 초과 ($($cmd.timeout)초)"
            $allSuccess = $false
            Write-Host " TIMEOUT" -ForegroundColor Red
        }
        elseif ($cmd.expect -ne "" -and $output -notmatch [regex]::Escape($cmd.expect)) {
            $stepResult.status = "fail"
            $stepResult.output = $output.Trim()
            $stepResult.error = "기대 문구 불일치: '$($cmd.expect)'"
            $allSuccess = $false
            Write-Host " FAIL" -ForegroundColor Red
            Write-Host "  출력: $($output.Trim())" -ForegroundColor DarkGray
        }
        else {
            $stepResult.status = "success"
            $stepResult.output = $output.Trim()
            Write-Host " OK" -ForegroundColor Green
        }
    }
    catch {
        $stepResult.status = "error"
        $stepResult.error = $_.Exception.Message
        $allSuccess = $false
        Write-Host " ERROR" -ForegroundColor Red
        Write-Host "  $($_.Exception.Message)" -ForegroundColor DarkGray
    }

    $results += $stepResult
}

# --- 잠금화면 이미지 설정 ---
Write-Host ""
Write-Host "잠금화면 이미지 설정 중..." -ForegroundColor Yellow -NoNewline

$lockscreenResult = @{
    name   = "잠금화면 이미지 설정"
    status = "pending"
    output = ""
    error  = ""
}

try {
    # GitHub에서 잠금화면 이미지 다운로드
    $imageTemp = "$env:TEMP\lockscreen_company.png"
    $imageUrl = "https://api.github.com/repos/$GITHUB_OWNER/$GITHUB_REPO/contents/assets/lockscreen.png"
    $imageHeaders = @{
        "Authorization" = "token $GITHUB_TOKEN"
        "Accept"        = "application/vnd.github.v3.raw"
    }
    Invoke-WebRequest -Uri $imageUrl -Headers $imageHeaders -OutFile $imageTemp -UseBasicParsing

    # 잠금화면 이미지 복사
    $destPath = $config.lockscreen.local_path
    $destDir = Split-Path $destPath -Parent
    if (-not (Test-Path $destDir)) {
        New-Item -ItemType Directory -Path $destDir -Force | Out-Null
    }
    Copy-Item -Path $imageTemp -Destination $destPath -Force

    # 레지스트리로 잠금화면 이미지 강제 설정
    $regPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization"
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }
    Set-ItemProperty -Path $regPath -Name "LockScreenImage" -Value $destPath
    Set-ItemProperty -Path $regPath -Name "NoChangingLockScreen" -Value 1 -Type DWord

    $lockscreenResult.status = "success"
    $lockscreenResult.output = "잠금화면 이미지 설정 완료: $destPath"
    Write-Host " OK" -ForegroundColor Green
}
catch {
    $lockscreenResult.status = "error"
    $lockscreenResult.error = $_.Exception.Message
    $allSuccess = $false
    Write-Host " FAIL" -ForegroundColor Red
    Write-Host "  $($_.Exception.Message)" -ForegroundColor DarkGray
}

$results += $lockscreenResult

# --- 결과 요약 ---
$endTime = Get-Date
$duration = ($endTime - $startTime).TotalSeconds

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "   설정 결과 요약" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

foreach ($r in $results) {
    $icon = switch ($r.status) {
        "success" { "[O]" }
        "fail"    { "[X]" }
        "timeout" { "[T]" }
        "error"   { "[!]" }
        default   { "[?]" }
    }
    $color = if ($r.status -eq "success") { "Green" } else { "Red" }
    Write-Host "  $icon $($r.name)" -ForegroundColor $color
}

$successCount = ($results | Where-Object { $_.status -eq "success" }).Count
Write-Host ""
Write-Host "총 $($results.Count)개 중 ${successCount}개 성공 (소요시간: $([math]::Round($duration, 1))초)" -ForegroundColor Cyan

# --- GitHub Issue로 결과 보고 ---
Write-Host ""
Write-Host "결과를 관리자 페이지에 보고 중..." -ForegroundColor Yellow -NoNewline

try {
    $overallStatus = if ($allSuccess) { "SUCCESS" } else { "PARTIAL" }

    $bodyLines = @()
    $bodyLines += "## PC 설정 결과"
    $bodyLines += ""
    $bodyLines += "| 항목 | 값 |"
    $bodyLines += "|---|---|"
    $bodyLines += "| 이름 | $userName |"
    $bodyLines += "| 부서 | $department |"
    $bodyLines += "| PC명 | $pcName |"
    $bodyLines += "| 실행일시 | $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') |"
    $bodyLines += "| 결과 | **$overallStatus** ($successCount/$($results.Count)) |"
    $bodyLines += ""
    $bodyLines += "## 상세 결과"
    $bodyLines += ""
    $bodyLines += "| 단계 | 상태 | 비고 |"
    $bodyLines += "|---|---|---|"

    foreach ($r in $results) {
        $statusEmoji = if ($r.status -eq "success") { "pass" } else { "FAIL" }
        $note = if ($r.error) { $r.error } else { $r.output.Substring(0, [Math]::Min(50, $r.output.Length)) }
        $bodyLines += "| $($r.name) | $statusEmoji | $note |"
    }

    $body = $bodyLines -join "`n"

    $issueTitle = "[$overallStatus] $userName ($department) - $(Get-Date -Format 'yyyy-MM-dd')"

    $labels = @()
    if ($allSuccess) { $labels += "success" } else { $labels += "partial" }
    $labels += "onboarding"

    $issueData = @{
        title  = $issueTitle
        body   = $body
        labels = $labels
    } | ConvertTo-Json -Depth 3

    $reportHeaders = @{
        "Authorization" = "token $GITHUB_TOKEN"
        "Accept"        = "application/vnd.github.v3+json"
        "Content-Type"  = "application/json; charset=utf-8"
    }

    $apiUrl = "https://api.github.com/repos/$GITHUB_OWNER/$GITHUB_REPO/issues"

    $response = Invoke-RestMethod -Uri $apiUrl -Method Post -Headers $reportHeaders -Body ([System.Text.Encoding]::UTF8.GetBytes($issueData))

    Write-Host " OK" -ForegroundColor Green
    Write-Host "  보고 완료: $($response.html_url)" -ForegroundColor DarkGray
}
catch {
    Write-Host " FAIL" -ForegroundColor Red
    Write-Host "  GitHub 보고 실패: $($_.Exception.Message)" -ForegroundColor DarkGray
    Write-Host "  (관리자에게 수동 보고해주세요)" -ForegroundColor DarkGray
}

# --- 완료 ---
Write-Host ""
if ($allSuccess) {
    Write-Host "모든 설정이 완료되었습니다!" -ForegroundColor Green
} else {
    Write-Host "일부 설정이 실패했습니다. 관리자에게 문의하세요." -ForegroundColor Yellow
}
Write-Host ""
Read-Host "아무 키나 누르면 종료합니다"
