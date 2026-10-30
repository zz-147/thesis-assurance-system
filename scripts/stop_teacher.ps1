# 停止 tas 老师端（双击项目根目录的「停止老师端.bat」）
# 作用：找到占用 8765 端口的服务进程并结束；找不到就提示未运行。
$conn = Get-NetTCPConnection -LocalPort 8765 -State Listen -ErrorAction SilentlyContinue |
    Select-Object -First 1
if ($conn) {
    $serverPid = $conn.OwningProcess
    try {
        Stop-Process -Id $serverPid -Force
        Write-Host "已停止老师端（PID $serverPid），端口 8765 已释放。" -ForegroundColor Green
    } catch {
        Write-Host "停止失败：$($_.Exception.Message)" -ForegroundColor Red
        Write-Host "可以在任务管理器里结束 python.exe 进程。"
    }
} else {
    Write-Host "老师端当前没有在运行（8765 端口空闲）。"
}
Write-Host ""
Read-Host "按回车键关闭窗口" | Out-Null
