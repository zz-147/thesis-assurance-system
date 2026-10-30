# 构建「免安装绿色版」：把 Python 运行时 + 项目代码 + 依赖打包到 dist\tas-web-portable
# 不需要联网、不需要 PyInstaller；老师电脑解压后双击「启动老师端.bat」即可用。
# 用法：powershell -NoProfile -ExecutionPolicy Bypass -File scripts\build_portable.ps1
$ErrorActionPreference = "Continue"
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root

function Find-RuntimePython {
    $candidates = @()
    if ($env:TAS_PYTHON) { $candidates += $env:TAS_PYTHON }
    $candidates += "D:\dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe"
    $candidates += (Get-Command python -ErrorAction SilentlyContinue | ForEach-Object { $_.Source })
    foreach ($exe in $candidates) {
        if (-not $exe -or -not (Test-Path $exe)) { continue }
        & $exe -c "import docx, lxml, openpyxl" 2>$null
        if ($LASTEXITCODE -eq 0) { return (Split-Path -Parent $exe) }
    }
    return $null
}

$runtime = Find-RuntimePython
if (-not $runtime) {
    Write-Host "未找到带 python-docx/lxml/openpyxl 的 Python 运行时。" -ForegroundColor Red
    exit 1
}
Write-Host "Python 运行时：$runtime"

$out = Join-Path $Root "dist\tas-web-portable"
if (Test-Path $out) { Remove-Item $out -Recurse -Force }
New-Item -ItemType Directory -Force -Path $out | Out-Null
$dstPython = Join-Path $out "python"
New-Item -ItemType Directory -Force -Path $dstPython | Out-Null

Write-Host "复制 Python 运行时（约 30-60 秒）……"
$runtimeFiles = @("python.exe", "pythonw.exe", "python3.dll", "python312.dll",
                  "vcruntime140.dll", "vcruntime140_1.dll", "LICENSE.txt")
foreach ($name in $runtimeFiles) {
    $source = Join-Path $runtime $name
    if (Test-Path $source) { Copy-Item $source $dstPython -Force }
}
foreach ($dir in @("DLLs", "Lib")) {
    $source = Join-Path $runtime $dir
    if (Test-Path $source) { Copy-Item $source $dstPython -Recurse -Force }
}

Write-Host "精简标准库/第三方包……"
foreach ($dir in @("Lib\test", "Lib\idlelib", "Lib\tkinter", "Lib\turtledemo",
                   "Lib\lib2to3", "Lib\ensurepip", "Lib\site-packages\pip",
                   "Lib\site-packages\pkg_resources", "Lib\site-packages\setuptools",
                   "Lib\site-packages\wheel", "Lib\site-packages\pandas*",
                   "Lib\site-packages\numpy*", "Lib\site-packages\PIL",
                   "Lib\site-packages\pptx", "Lib\site-packages\xlsxwriter",
                   "Lib\site-packages\dateutil", "Lib\site-packages\tzdata*",
                   "Lib\site-packages\win32*", "Lib\site-packages\pywin32*")) {
    $path = Join-Path $dstPython $dir
    if (Test-Path $path) { Remove-Item $path -Recurse -Force -ErrorAction SilentlyContinue }
}
# site-packages 只保留本项目需要的包
$keep = @("docx", "lxml", "openpyxl", "et_xmlfile", "typing_extensions.py", "__pycache__")
$sitePackages = Join-Path $dstPython "Lib\site-packages"
if (Test-Path $sitePackages) {
    Get-ChildItem $sitePackages -Force | Where-Object {
        $keep -notcontains $_.Name -and $_.Name -notlike "et_xmlfile*" -and
        $_.Name -notlike "typing_extensions*" -and $_.Name -notlike "docx*" -and
        $_.Name -notlike "lxml*" -and $_.Name -notlike "openpyxl*"
    } | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
}
Get-ChildItem $dstPython -Recurse -Directory -Filter "__pycache__" -ErrorAction SilentlyContinue |
    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "复制项目代码与资源……"
foreach ($item in @("src", "rules", "profiles", "docs")) {
    $source = Join-Path $Root $item
    if (Test-Path $source) { Copy-Item $source (Join-Path $out $item) -Recurse -Force }
}
Copy-Item (Join-Path $Root "README.md") (Join-Path $out "README.md") -Force -ErrorAction SilentlyContinue

$launcher = @'
@echo off
chcp 65001 >nul
cd /d "%~dp0"
set "TAS_PROJECT_ROOT=%~dp0"
set "PYTHONPATH=%~dp0src"
set "PYTHONIOENCODING=utf-8"
if not exist "data" mkdir "data"
"%~dp0python\python.exe" -m tas web --host 127.0.0.1 --port 8765 --open-browser
pause
'@
Set-Content -Path (Join-Path $out "启动老师端.bat") -Value $launcher -Encoding Default

$stopper = @'
@echo off
chcp 65001 >nul
powershell -NoProfile -ExecutionPolicy Bypass -Command "$c=(Get-NetTCPConnection -LocalPort 8765 -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1); if ($c) { Stop-Process -Id $c.OwningProcess -Force; Write-Host '已停止老师端' } else { Write-Host '老师端未在运行' }; pause"
'@
Set-Content -Path (Join-Path $out "停止老师端.bat") -Value $stopper -Encoding Default

$guide = @'
tas 老师端（免安装绿色版）
==========================

1. 双击「启动老师端.bat」→ 自动打开浏览器 http://127.0.0.1:8765
2. 教师端 → 选择规范 → 「选择整个文件夹」/ 拖入多份 .docx → 批量检查
3. 下载 batch.xlsx、学生反馈（feedback.zip）等
4. 关闭窗口即退出；端口被占用时双击「停止老师端.bat」

整个目录可整体复制到别的电脑（Windows 64 位），不需要安装 Python。
数据保存在本目录 data/ 下；不要删除该目录。
详细说明见 docs\TEACHER_GUIDE.md。
如页面预览/缩略图提示失败：确认电脑装了 Microsoft Word，并把本目录放在普通位置（不要放在受限/沙箱目录）。
'@
Set-Content -Path (Join-Path $out "使用说明.txt") -Value $guide -Encoding Default

Write-Host "自检：内置 Python 能否导入依赖……"
& (Join-Path $dstPython "python.exe") -c "import docx, lxml, openpyxl; print('deps ok')" 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "内置运行时导入依赖失败。" -ForegroundColor Red
    exit 1
}

$size = (Get-ChildItem $out -Recurse -File | Measure-Object Length -Sum).Sum
Write-Host ("打包完成：{0}（{1:N1} MB）" -f $out, ($size / 1MB)) -ForegroundColor Green
$zipPath = Join-Path $Root "dist\tas-web-portable.zip"
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
Write-Host "压缩 zip（约 20-40 MB，可能需要 1 分钟）……"
Compress-Archive -Path $out -DestinationPath $zipPath -CompressionLevel Optimal
Write-Host "zip：$zipPath"
Write-Host "解压后双击「启动老师端.bat」即可，无需安装 Python。"
