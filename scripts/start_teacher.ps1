# tas 老师端一键启动（双击项目根目录的「启动老师端.bat」）
# 作用：找到可用的 Python → 启动本地 Web 服务 → 自动打开浏览器。
# 关闭这个窗口即停止服务；数据只在本机。
$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root

function Test-TasPython([string]$Exe, [string[]]$Prefix) {
    if (-not $Exe -or -not (Test-Path $Exe)) { return $false }
    $checkArgs = @($Prefix) + @("-c", "import docx, lxml, openpyxl")
    try {
        & $Exe @checkArgs *> $null
        return ($LASTEXITCODE -eq 0)
    } catch {
        return $false
    }
}

$candidates = @()
if ($env:TAS_PYTHON) { $candidates += ,@($env:TAS_PYTHON, @()) }
$candidates += ,@("D:\dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe", @())
$pythonCmd = Get-Command python -ErrorAction SilentlyContinue
if ($pythonCmd) { $candidates += ,@($pythonCmd.Source, @()) }
$pyCmd = Get-Command py -ErrorAction SilentlyContinue
if ($pyCmd) { $candidates += ,@($pyCmd.Source, @("-3")) }

$python = $null
$prefixArgs = @()
foreach ($candidate in $candidates) {
    $exe = $candidate[0]
    $prefix = $candidate[1]
    if (Test-TasPython $exe $prefix) {
        $python = $exe
        $prefixArgs = $prefix
        break
    }
}

if (-not $python) {
    Write-Host ""
    Write-Host "未找到可用的 Python（需要 python-docx / lxml / openpyxl）。" -ForegroundColor Red
    Write-Host "解决办法："
    Write-Host "  1) 安装 Python 3.10+ 后运行：python -m pip install python-docx lxml openpyxl"
    Write-Host "  2) 或设置环境变量 TAS_PYTHON 指向已有的 python.exe"
    Write-Host ""
    Read-Host "按回车键退出"
    exit 1
}

$env:PYTHONPATH = Join-Path $Root "src"
$env:PYTHONIOENCODING = "utf-8"
$port = 8765

# 已在运行：直接打开浏览器
try {
    $health = Invoke-WebRequest -Uri "http://127.0.0.1:$port/health" -TimeoutSec 2 -UseBasicParsing
    if ($health.StatusCode -eq 200) {
        Write-Host "老师端已经在运行，直接打开浏览器：http://127.0.0.1:$port/"
        Start-Process "http://127.0.0.1:$port/"
        Read-Host "按回车键退出"
        exit 0
    }
} catch {
    # 没有在运行，继续启动
}

Write-Host "TAS 老师端启动中……（关闭本窗口即停止服务）"
$serverArgs = @($prefixArgs) + @("-m", "tas", "web", "--host", "127.0.0.1", "--port", "$port")
$process = Start-Process -FilePath $python `
    -ArgumentList $serverArgs `
    -PassThru `
    -WorkingDirectory $Root

$started = $false
for ($i = 0; $i -lt 30; $i++) {
    Start-Sleep -Milliseconds 500
    if ($process.HasExited) { break }
    try {
        $health = Invoke-WebRequest -Uri "http://127.0.0.1:$port/health" -TimeoutSec 2 -UseBasicParsing
        if ($health.StatusCode -eq 200) {
            $started = $true
            break
        }
    } catch {
        # 等待服务就绪
    }
}

if (-not $started) {
    Write-Host ""
    Write-Host "启动失败。常见原因：端口 8765 被占用，或 Python 依赖缺失。" -ForegroundColor Red
    Write-Host "可在项目目录手动运行："
    Write-Host ('  $env:PYTHONPATH = "' + $Root + '\src"')
    Write-Host ('  & "' + $python + '" -m tas web --port 8765')
    Write-Host ""
    Read-Host "按回车键退出"
    exit 1
}

Start-Process "http://127.0.0.1:$port/"
Write-Host ""
Write-Host "已启动：http://127.0.0.1:$port/"
Write-Host "老师操作：教师端 → 选择文件夹 / 多选文档 → 批量检查 → 下载 xlsx 和学生反馈包"
Write-Host "关闭本窗口会退出启动器；若下次提示 8765 被占用，双击「停止老师端.bat」。"
Write-Host "数据保存在本机 data/ 目录。"
Wait-Process -Id $process.Id
