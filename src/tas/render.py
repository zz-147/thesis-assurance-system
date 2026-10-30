"""文档渲染预览：docx → PDF（Word COM）→ PNG（Windows.Data.Pdf WinRT）。

只依赖 Windows 自带能力（Word + PowerShell 5.1 + WinRT），不引入第三方库。
用于 Web 检查/修复结果页的"所见即所得"预览，以及 CLI `tas preview`。

安全/稳定性：
- 只读打开 docx（ReadOnly=True），不修改原文件；导出到临时目录；
- Word 宏强制禁用（AutomationSecurity=3），DisplayAlerts=0，Visible=False；
- 每一步都有超时，失败抛 RenderError（附 PowerShell stderr 摘要），不影响检查/修复主流程。
"""

from __future__ import annotations

import os
import platform
import shutil
import subprocess
from pathlib import Path
from typing import Optional

PAGE_GLOB = "page-*.png"
WORD_TO_PDF_TIMEOUT = 120.0
PDF_TO_PNG_TIMEOUT = 120.0


class RenderError(RuntimeError):
    """渲染失败（环境缺失/Word 异常/超时）。"""


def render_available() -> bool:
    return platform.system() == "Windows" and shutil.which("powershell") is not None


# ---------------------------------------------------------------------------
# PowerShell 调用
# ---------------------------------------------------------------------------
def _run_powershell(script: str, env_extra: dict[str, str], timeout: float) -> str:
    if not render_available():
        raise RenderError("渲染预览需要 Windows + PowerShell（当前环境不支持）")
    env = os.environ.copy()
    env.update(env_extra)
    try:
        completed = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-Command",
                script,
            ],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            env=env,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except subprocess.TimeoutExpired as exc:
        raise RenderError(f"渲染超时（>{timeout:g}s）") from exc
    except OSError as exc:
        raise RenderError(f"无法启动 PowerShell：{exc}") from exc
    if completed.returncode != 0:
        detail = (completed.stderr or completed.stdout or "").strip()
        detail = detail[-800:] if detail else "（无输出）"
        raise RenderError(f"渲染失败（exit {completed.returncode}）：{detail}")
    return completed.stdout or ""


# ---------------------------------------------------------------------------
# 1) docx → PDF（Word COM）
# ---------------------------------------------------------------------------
_WORD_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
$word = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    try { $word.AutomationSecurity = 3 } catch {}
    $doc = $word.Documents.Open($env:TAS_RENDER_SRC, $false, $true)
    try {
        $doc.ExportAsFixedFormat($env:TAS_RENDER_PDF, 17)
    } finally {
        $doc.Close(0)
    }
} finally {
    if ($word -ne $null) { $word.Quit() }
}
Write-Output 'ok'
"""


def word_to_pdf(docx: Path | str, pdf: Path | str, *, timeout: float = WORD_TO_PDF_TIMEOUT) -> Path:
    docx, pdf = Path(docx), Path(pdf)
    if not docx.is_file():
        raise RenderError(f"文件不存在：{docx}")
    pdf.parent.mkdir(parents=True, exist_ok=True)
    if pdf.exists():
        pdf.unlink()
    _run_powershell(
        _WORD_SCRIPT,
        {"TAS_RENDER_SRC": str(docx.resolve()), "TAS_RENDER_PDF": str(pdf.resolve())},
        timeout,
    )
    if not pdf.is_file() or pdf.stat().st_size < 1000:
        raise RenderError("Word 导出 PDF 失败（输出文件不存在或过小）")
    return pdf


# ---------------------------------------------------------------------------
# 2) PDF → PNG（Windows.Data.Pdf / WinRT）
# ---------------------------------------------------------------------------
_PDF_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType=WindowsRuntime]
$null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType=WindowsRuntime]
$null = [Windows.Storage.Streams.DataWriter, Windows.Storage.Streams, ContentType=WindowsRuntime]

function Await-Operation($WinRtTask, $ResultType) {
    $asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
        $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and
        $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
    })[0]
    $netTask = $asTask.MakeGenericMethod($ResultType).Invoke($null, @($WinRtTask))
    $netTask.Wait(-1) | Out-Null
    $netTask.Result
}
function Await-Action($WinRtAction) {
    $asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
        $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and
        $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction'
    })[0]
    $netTask = $asTask.Invoke($null, @($WinRtAction))
    $netTask.Wait(-1) | Out-Null
}

$pdfPath = $env:TAS_RENDER_PDF
$outDir = $env:TAS_RENDER_OUT
$maxPages = [int]$env:TAS_RENDER_MAX
$width = [int]$env:TAS_RENDER_WIDTH
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
Get-ChildItem $outDir -Filter 'page-*.png' -ErrorAction SilentlyContinue | Remove-Item -Force

$bytes = [System.IO.File]::ReadAllBytes($pdfPath)
$stream = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
$writer = New-Object Windows.Storage.Streams.DataWriter($stream.GetOutputStreamAt(0))
$writer.WriteBytes($bytes)
Await-Operation ($writer.StoreAsync()) ([uint32]) | Out-Null
Await-Operation ($writer.FlushAsync()) ([bool]) | Out-Null
$writer.DetachStream() | Out-Null

$pdf = Await-Operation ([Windows.Data.Pdf.PdfDocument]::LoadFromStreamAsync($stream)) ([Windows.Data.Pdf.PdfDocument])
$count = [Math]::Min($pdf.PageCount, $maxPages)
for ($i = 0; $i -lt $count; $i++) {
    $page = $pdf.GetPage($i)
    $imgStream = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
    $opts = New-Object Windows.Data.Pdf.PdfPageRenderOptions
    $opts.DestinationWidth = $width
    Await-Action ($page.RenderToStreamAsync($imgStream, $opts))
    $netStream = [System.IO.WindowsRuntimeStreamExtensions]::AsStreamForRead($imgStream)
    $outPath = Join-Path $outDir ("page-{0:D2}.png" -f ($i + 1))
    $fs = [System.IO.File]::Create($outPath)
    $netStream.CopyTo($fs)
    $fs.Close()
    $netStream.Close()
    $imgStream.Dispose()
    $page.Dispose()
}
Write-Output ("pages=" + $count + "/" + $pdf.PageCount)
"""


def pdf_to_png(
    pdf: Path | str,
    out_dir: Path | str,
    *,
    max_pages: int = 2,
    width: int = 1400,
    timeout: float = PDF_TO_PNG_TIMEOUT,
) -> list[Path]:
    pdf, out_dir = Path(pdf), Path(out_dir)
    if not pdf.is_file():
        raise RenderError(f"PDF 不存在：{pdf}")
    out_dir.mkdir(parents=True, exist_ok=True)
    _run_powershell(
        _PDF_SCRIPT,
        {
            "TAS_RENDER_PDF": str(pdf.resolve()),
            "TAS_RENDER_OUT": str(out_dir.resolve()),
            "TAS_RENDER_MAX": str(max(1, int(max_pages))),
            "TAS_RENDER_WIDTH": str(max(400, int(width))),
        },
        timeout,
    )
    pages = sorted(out_dir.glob(PAGE_GLOB))
    if not pages:
        raise RenderError("PDF 渲染未生成页面图片")
    return pages


# ---------------------------------------------------------------------------
# 对外入口
# ---------------------------------------------------------------------------
def render_preview(
    docx: Path | str,
    out_dir: Optional[Path | str] = None,
    *,
    max_pages: int = 2,
    width: int = 1400,
    timeout: float = WORD_TO_PDF_TIMEOUT,
) -> list[Path]:
    """把 docx 渲染成前 N 页 PNG（供 Web/CLI 预览）。返回 PNG 路径列表。"""
    docx = Path(docx)
    out_dir = Path(out_dir) if out_dir is not None else docx.parent / "preview"
    out_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = out_dir / "preview.pdf"
    word_to_pdf(docx, pdf_path, timeout=timeout)
    return pdf_to_png(pdf_path, out_dir, max_pages=max_pages, width=width, timeout=timeout)
