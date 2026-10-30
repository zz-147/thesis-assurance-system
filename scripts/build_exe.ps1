# 把 tas 打包成免安装的 Windows 目录版（老师电脑没 Python 也能用）
# 用法：右键「使用 PowerShell 运行」，或命令行：
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts\build_exe.ps1
# 产物：dist\tas-web\ 目录（双击里面的「启动老师端.bat」即可）
$ErrorActionPreference = "Continue"
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root

function Find-TasPython {
    $candidates = @()
    if ($env:TAS_PYTHON) { $candidates += $env:TAS_PYTHON }
    $candidates += "D:\dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe"
    $pythonCmd = Get-Command python -ErrorAction SilentlyContinue
    if ($pythonCmd) { $candidates += $pythonCmd.Source }
    foreach ($candidate in $candidates) {
        if (-not $candidate -or -not (Test-Path $candidate)) { continue }
        try {
            & $candidate -c "import docx, lxml, openpyxl" 2>$null
            if ($LASTEXITCODE -eq 0) { return $candidate }
        } catch { }
    }
    return $null
}

$python = Find-TasPython
if (-not $python) {
    Write-Host "未找到可用的 Python（需要 python-docx / lxml / openpyxl）。" -ForegroundColor Red
    exit 1
}
Write-Host "使用 Python: $python"

# 1) 安装/检查 PyInstaller
& $python -c "import PyInstaller" 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "安装 PyInstaller（首次打包需要，约 1 分钟）……"
    & $python -m pip install --upgrade pyinstaller --disable-pip-version-check
    if ($LASTEXITCODE -ne 0) {
        Write-Host "PyInstaller 不可用（网络受限），改用「免安装绿色版」打包……" -ForegroundColor Yellow
        & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "build_portable.ps1")
        exit $LASTEXITCODE
    }
}

# 2) PyInstaller 打包（目录版，启动比单文件快）
$dist = Join-Path $Root "dist"
$build = Join-Path $Root "build"
$entry = Join-Path $Root "scripts\tas_entry.py"
$srcPath = Join-Path $Root "src"
Write-Host "开始打包（约 1-3 分钟）……"
& $python -m PyInstaller `
    --noconfirm --clean --onedir --console `
    --name tas-web `
    --paths $srcPath `
    --collect-all docx `
    --collect-all openpyxl `
    --hidden-import lxml `
    --hidden-import lxml.etree `
    --hidden-import lxml._elementpath `
    --distpath $dist `
    --workpath (Join-Path $build "work") `
    --specpath $build `
    $entry
if ($LASTEXITCODE -ne 0) {
    Write-Host "PyInstaller 打包失败，改用「免安装绿色版」打包……" -ForegroundColor Yellow
    & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "build_portable.ps1")
    exit $LASTEXITCODE
}

$appDir = Join-Path $dist "tas-web"
if (-not (Test-Path $appDir)) {
    Write-Host "未找到产物目录：$appDir" -ForegroundColor Red
    exit 1
}

# 3) 把规则/规范/文档/启动器放到 exe 旁边
foreach ($item in @("rules", "profiles", "docs")) {
    $source = Join-Path $Root $item
    if (Test-Path $source) {
        Copy-Item $source (Join-Path $appDir $item) -Recurse -Force
    }
}
Copy-Item (Join-Path $Root "README.md") (Join-Path $appDir "README.md") -Force -ErrorAction SilentlyContinue

$launcher = @'
@echo off
chcp 65001 >nul
cd /d "%~dp0"
set "TAS_PROJECT_ROOT=%~dp0"
if not exist "data" mkdir "data"
tas-web.exe web --host 127.0.0.1 --port 8765 --open-browser
pause
'@
Set-Content -Path (Join-Path $appDir "启动老师端.bat") -Value $launcher -Encoding Default

$stopper = @'
@echo off
chcp 65001 >nul
powershell -NoProfile -ExecutionPolicy Bypass -Command "$c=(Get-NetTCPConnection -LocalPort 8765 -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1); if ($c) { Stop-Process -Id $c.OwningProcess -Force; Write-Host '已停止老师端' } else { Write-Host '老师端未在运行' }; pause"
'@
Set-Content -Path (Join-Path $appDir "停止老师端.bat") -Value $stopper -Encoding Default

$guide = @'
tas 老师端（免安装版）
========================

1. 双击「启动老师端.bat」→ 自动打开浏览器 http://127.0.0.1:8765
2. 教师端 → 选择规范 → 「选择整个文件夹」/ 拖入多份 .docx → 批量检查
3. 下载 batch.xlsx、学生反馈（feedback.zip）等
4. 关闭窗口即退出；端口被占用时双击「停止老师端.bat」

数据在本目录 data/ 下；不要删除该目录。
详细说明见 docs\TEACHER_GUIDE.md。
'@
Set-Content -Path (Join-Path $appDir "使用说明.txt") -Value $guide -Encoding Default

# 4) 压缩成 zip（可选，方便发给老师）
$zipPath = Join-Path $dist "tas-web-免安装.zip"
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
Compress-Archive -Path $appDir -DestinationPath $zipPath -CompressionLevel Optimal
Write-Host ""
Write-Host "打包完成：" -ForegroundColor Green
Write-Host "  目录：$appDir"
Write-Host "  zip ：$zipPath"
Write-Host "双击 $appDir\启动老师端.bat 即可试用（无需安装 Python）。"
