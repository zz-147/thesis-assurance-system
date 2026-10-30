"""tas Web 双入口（教师端 / 学生端），纯标准库实现，本地运行、无账号。

学生端：上传 .docx → 选择规范 → 检查报告 → 一键修复 → 下载 原名_fixed.docx + 修复报告
教师端：选择规范 → 上传多份 .docx → 批量检查 → 班级汇总（通过率 / Top 问题 / 逐份）→ 下载 xlsx/csv/html

启动：
    python -m tas web --host 127.0.0.1 --port 8765
安全约定：默认只绑定 127.0.0.1；上传文件只在本机临时目录处理；下载只允许访问本次 token 目录。
"""

from __future__ import annotations

import html
import json
import os
import re
import sys
import tempfile
import uuid
from datetime import datetime
from email.parser import BytesParser
from email.policy import default as email_policy
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Optional
from urllib.parse import unquote, urlparse

from ..checkers import IMPLEMENTED_RULES, run_checks
from ..config import RuleSet, list_profiles, resolve_profile
from ..docx_io import DocModel
from ..report import Report
from ..roles import classify

MAX_UPLOAD_BYTES = 30 * 1024 * 1024
STORE_ROOT = Path(tempfile.mkdtemp(prefix="tas-web-"))
_TOKEN_RE = re.compile(r"^[0-9a-f]{12,40}$")

# 拖拽上传：CSS / JS 单独存放，避免 _page 的 f-string 转义问题
_DROP_CSS = """
.dropzone { position:relative; border:2px dashed #cbd5e1; border-radius:12px; background:#f8fafc;
            padding:18px 16px; text-align:center; cursor:pointer; transition:.15s; outline:none; }
.dropzone:hover, .dropzone:focus, .dropzone.drag { border-color:var(--accent); background:#eff6ff; }
.dropzone.drag { transform:scale(1.01); box-shadow:0 0 0 4px rgba(37,99,235,.08); }
.dropzone.has-file { border-style:solid; border-color:#86efac; background:#f0fdf4; }
.dropzone .dz-icon { font-size:26px; line-height:1.2; }
.dropzone .dz-main { font-size:14px; font-weight:600; margin:4px 0 2px; }
.dropzone .dz-sub { font-size:12px; color:var(--muted); }
.chips { display:flex; flex-wrap:wrap; gap:8px; margin-top:8px; }
.chip { font-size:12px; background:#eef2ff; color:#1e3a8a; border-radius:999px; padding:4px 10px; }
.preview-grid { display:grid; grid-template-columns:repeat(auto-fill,minmax(320px,1fr)); gap:14px; margin:10px 0 18px; }
.preview-grid figure { margin:0; border:1px solid var(--line); border-radius:10px; overflow:hidden; background:#fff; }
.preview-grid img { display:block; width:100%; height:auto; }
.preview-grid figcaption { font-size:12px; color:var(--muted); padding:6px 10px; border-top:1px solid var(--line); }
label.check { display:flex; align-items:center; gap:8px; font-size:13px; color:var(--ink); margin:12px 0 0; }
label.check input { width:auto; }
textarea { width:100%; padding:8px 10px; border:1px solid var(--line); border-radius:8px;
           font:inherit; background:#fff; resize:vertical; box-sizing:border-box; }
.roster-list { margin:6px 0 0 18px; font-size:13px; }
.folder-row { display:flex; align-items:center; gap:10px; flex-wrap:wrap; margin-top:10px; }
.folder-btn { width:auto; padding:7px 12px; border:1px dashed var(--accent); border-radius:8px;
              background:#eff6ff; color:var(--accent); font-size:13px; cursor:pointer; }
.folder-btn:hover { background:#dbeafe; }
.folder-hint { font-size:12px; color:var(--muted); }
.thumb-grid { display:flex; flex-wrap:wrap; gap:10px; margin:10px 0 18px; }
.thumb-grid figure { margin:0; width:150px; border:1px solid var(--line); border-radius:8px;
                     padding:6px; background:#fff; }
.thumb-grid img { display:block; width:100%; height:auto; border-radius:4px; }
.thumb-grid figcaption { font-size:11px; color:var(--muted); margin-top:4px;
                         white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
"""

_DROP_JS = r"""
<script>
(function () {
  function fmtSize(bytes) {
    if (bytes === undefined || bytes === null) return '';
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / 1048576).toFixed(1) + ' MB';
  }
  function setupZone(zone) {
    var input = document.getElementById(zone.getAttribute('data-input'));
    var listId = zone.getAttribute('data-list');
    var list = listId ? document.getElementById(listId) : null;
    if (!input) return;
    var multiple = !!input.multiple;
    function render() {
      if (list) {
        list.innerHTML = '';
        var files = input.files || [];
        for (var i = 0; i < files.length; i++) {
          var chip = document.createElement('span');
          chip.className = 'chip';
          chip.textContent = files[i].name + ' · ' + fmtSize(files[i].size);
          list.appendChild(chip);
        }
      }
      zone.classList.toggle('has-file', !!(input.files && input.files.length));
    }
    function assign(files) {
      if (!files || !files.length) return;
      if (!multiple && files.length > 1) {
        alert('学生端一次只能处理一份文档，已保留第一份：' + files[0].name);
      }
      try {
        var dt = new DataTransfer();
        for (var i = 0; i < files.length; i++) {
          if (multiple || i === 0) dt.items.add(files[i]);
        }
        input.files = dt.files;
      } catch (err) {
        if (!multiple) input.files = files;
      }
      render();
    }
    zone.addEventListener('click', function (event) {
      if (event.target === input) return;
      input.click();
    });
    zone.addEventListener('keydown', function (event) {
      if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); input.click(); }
    });
    ['dragenter', 'dragover'].forEach(function (name) {
      zone.addEventListener(name, function (event) {
        event.preventDefault(); event.stopPropagation(); zone.classList.add('drag');
      });
    });
    ['dragleave', 'drop'].forEach(function (name) {
      zone.addEventListener(name, function (event) {
        event.preventDefault(); event.stopPropagation(); zone.classList.remove('drag');
      });
    });
    zone.addEventListener('drop', function (event) {
      var dt = event.dataTransfer;
      if (dt && dt.files && dt.files.length) assign(dt.files);
    });
    input.addEventListener('change', render);
    render();
  }
  // 教师端：选择整个文件夹（webkitdirectory），自动过滤 .docx 并批量提交
  var folderInput = document.getElementById('teacher-dir');
  var teacherInput = document.getElementById('teacher-files');
  if (folderInput && teacherInput) {
    folderInput.addEventListener('change', function () {
      var picked = [];
      var all = folderInput.files || [];
      for (var i = 0; i < all.length; i++) {
        var file = all[i];
        if (!/\.docx$/i.test(file.name)) continue;
        if (file.name.indexOf('~$') === 0) continue;
        picked.push(file);
      }
      if (!picked.length) {
        alert('这个文件夹里没有找到 .docx 文件。');
        return;
      }
      try {
        var dt = new DataTransfer();
        for (var j = 0; j < picked.length; j++) dt.items.add(picked[j]);
        teacherInput.files = dt.files;
        teacherInput.dispatchEvent(new Event('change'));
        var preview = picked.slice(0, 3).map(function (f) { return f.name; }).join('、');
        alert('已选择 ' + picked.length + ' 份 .docx：' + preview + (picked.length > 3 ? ' …' : ''));
      } catch (err) {
        alert('当前浏览器不支持选择文件夹，请把文件拖进上传框，或点「选择文件」多选。');
      }
    });
  }
  document.querySelectorAll('.dropzone').forEach(setupZone);
  document.addEventListener('dragover', function (e) { e.preventDefault(); });
  document.addEventListener('drop', function (e) { e.preventDefault(); });
})();
</script>
"""


def _dropzone(
    input_id: str,
    name: str,
    *,
    multiple: bool = False,
    hint: str = "",
    accept: str = ".docx",
    title: str = ".docx",
) -> str:
    """拖拽上传区：支持从桌面拖入，也支持点击选择。"""
    list_id = input_id + "-list"
    main = f"把 {title} 文件拖到这里" + ("（可多份）" if multiple else "")
    sub = hint or "或点击选择文件 · 支持从桌面直接拖入"
    return (
        f'<div class="dropzone" tabindex="0" role="button" '
        f'data-input="{input_id}" data-list="{list_id}">'
        f'<div class="dz-icon">📄</div>'
        f'<div class="dz-main">{_esc(main)}</div>'
        f'<div class="dz-sub">{_esc(sub)}</div>'
        f'<input type="file" id="{input_id}" name="{name}" accept="{_esc(accept)}"'
        f'{" multiple" if multiple else ""} required hidden>'
        "</div>"
        f'<div class="chips" id="{list_id}"></div>'
    )


# ---------------------------------------------------------------------------
# 工具
# ---------------------------------------------------------------------------
def _esc(value) -> str:
    return html.escape(str(value), quote=True)


def _store_dir(token: str) -> Path:
    if not _TOKEN_RE.match(token or ""):
        raise ValueError("非法 token")
    directory = STORE_ROOT / token
    if not directory.is_dir():
        raise FileNotFoundError("任务不存在或已过期")
    return directory


def _new_token() -> str:
    token = uuid.uuid4().hex
    (STORE_ROOT / token).mkdir(parents=True, exist_ok=True)
    return token


def _safe_name(name: str) -> str:
    name = Path(name or "document.docx").name
    name = re.sub(r"[^\w.\-\u4e00-\u9fff]+", "_", name)
    return name[:120] or "document.docx"


def _resolve(ruleset_id: Optional[str]):
    from ..config import DEFAULT_PROFILE_ID

    return resolve_profile(ruleset_id or DEFAULT_PROFILE_ID)


def _profiles_select(name: str = "ruleset") -> str:
    try:
        profiles = list_profiles()
    except Exception:
        profiles = []
    if not profiles:
        return f'<select name="{name}"><option value="">默认规则（rules/rules.json）</option></select>'
    options = []
    for profile in profiles:
        options.append(
            f'<option value="{_esc(profile.id)}">{_esc(profile.name)}（{_esc(profile.discipline)}）</option>'
        )
    return f'<select name="{name}">' + "".join(options) + "</select>"


def _page(title: str, body: str) -> bytes:
    return f"""<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{_esc(title)} · TAS</title>
<style>
:root {{ --bg:#f6f7fb; --card:#fff; --ink:#1f2937; --muted:#6b7280; --line:#e5e7eb; --accent:#2563eb; --ok:#059669; --warn:#d97706; --err:#dc2626; }}
* {{ box-sizing:border-box; }}
body {{ margin:0; font-family:-apple-system,"Segoe UI","Microsoft YaHei",sans-serif; background:var(--bg); color:var(--ink); }}
header {{ background:#0f172a; color:#fff; padding:18px 28px; }}
header h1 {{ margin:0 0 4px; font-size:20px; }}
header p {{ margin:0; color:#cbd5e1; font-size:13px; }}
main {{ max-width:1180px; margin:24px auto 60px; padding:0 20px; }}
.grid {{ display:grid; grid-template-columns:1fr 1fr; gap:20px; }}
@media (max-width: 860px) {{ .grid {{ grid-template-columns:1fr; }} }}
.card {{ background:var(--card); border:1px solid var(--line); border-radius:12px; padding:20px; box-shadow:0 1px 2px rgba(15,23,42,.04); }}
.card h2 {{ margin:0 0 4px; font-size:17px; }}
.card .sub {{ color:var(--muted); font-size:13px; margin:0 0 14px; }}
{_DROP_CSS}
label {{ display:block; font-size:13px; color:var(--muted); margin:12px 0 6px; }}
select, input[type=file], input[type=text] {{ width:100%; padding:8px 10px; border:1px solid var(--line); border-radius:8px; background:#fff; }}
button {{ margin-top:14px; padding:9px 18px; border:0; border-radius:8px; background:var(--accent); color:#fff; font-size:14px; cursor:pointer; }}
button.secondary {{ background:#eef2ff; color:var(--accent); }}
button:hover {{ opacity:.92; }}
table {{ width:100%; border-collapse:collapse; font-size:13px; margin:10px 0; }}
th, td {{ border-bottom:1px solid var(--line); padding:7px 8px; text-align:left; vertical-align:top; }}
th {{ background:#f8fafc; color:#475569; font-weight:600; }}
.badge {{ display:inline-block; padding:1px 8px; border-radius:999px; font-size:12px; font-weight:600; }}
.err {{ background:#fef2f2; color:var(--err); }} .warn {{ background:#fffbeb; color:var(--warn); }} .ok {{ background:#ecfdf5; color:var(--ok); }} .info {{ background:#eff6ff; color:var(--accent); }}
.summary {{ display:flex; gap:22px; flex-wrap:wrap; margin:8px 0 18px; }}
.summary div {{ background:#f8fafc; border:1px solid var(--line); border-radius:10px; padding:10px 16px; }}
.summary b {{ display:block; font-size:22px; }}
.muted {{ color:var(--muted); }} .small {{ font-size:12px; }}
a {{ color:var(--accent); }}
pre {{ background:#0f172a; color:#e2e8f0; padding:12px; border-radius:8px; overflow:auto; font-size:12px; }}
</style></head><body>
<header><h1>TAS — 论文格式检查 / 修复</h1>
<p>本地运行 · 无账号 · 文档只在本机处理 · 只改格式不改内容 · 参考文献不自动修改</p></header>
<main>{body}</main>{_DROP_JS}</body></html>""".encode("utf-8")


def _result_page(title: str, sections: list[str]) -> bytes:
    return _page(title, "".join(sections))


def _issues_table(report: Report, limit: int = 300) -> str:
    rows = []
    for issue in report.issues[:limit]:
        badge = {"error": "err", "warning": "warn", "info": "info"}.get(issue.severity, "info")
        rows.append(
            f"<tr><td><span class='badge {badge}'>{_esc(issue.severity)}</span></td>"
            f"<td>{_esc(issue.rule_id)}<br><span class='small muted'>{_esc(issue.title)}</span></td>"
            f"<td>{_esc(issue.locator)}</td><td>{_esc(issue.message)}</td>"
            f"<td class='small'>{_esc(issue.actual)}<br>→ {_esc(issue.expected)}</td></tr>"
        )
    if not rows:
        return "<p class='muted'>未发现问题（当前已实现规则范围内）。</p>"
    extra = ""
    if len(report.issues) > limit:
        extra = f"<p class='muted'>仅显示前 {limit} 条，完整报告见下载文件。</p>"
    return (
        "<table><thead><tr><th>级别</th><th>规则</th><th>位置</th><th>说明</th><th>实际 → 期望</th></tr></thead><tbody>"
        + "".join(rows)
        + "</tbody></table>"
        + extra
    )


def _preview_html(token: str, pages: list, note: str = "") -> str:
    """页面预览（PNG）HTML；渲染失败时给提示但不影响主流程。"""
    if not pages:
        return f"<p class='small muted'>{_esc(note)}</p>" if note else ""
    figures = "".join(
        f'<figure><img src="/download/{_esc(token)}/preview/{_esc(page.name)}" alt="第 {index} 页预览">'
        f"<figcaption>第 {index} 页</figcaption></figure>"
        for index, page in enumerate(pages, 1)
    )
    tip = f"<p class='small muted'>{_esc(note)}</p>" if note else ""
    return f"<h3>页面预览（前 {len(pages)} 页）</h3>{tip}<div class='preview-grid'>{figures}</div>"


def _roster_text(report: dict) -> str:
    """把名单核对结果渲染成文本（老师可下载/转发）。"""
    lines = [
        "【名单核对】",
        (
            f"名单人数：{report.get('total_roster', 0)} ｜ "
            f"收到文件：{report.get('total_files', 0)} ｜ "
            f"已匹配：{report.get('matched', 0)}"
        ),
        "",
        f"缺交（{len(report.get('missing') or [])} 人）：",
    ]
    for item in report.get("missing") or []:
        lines.append(f"  - {item.get('student_id', '')} {item.get('name', '')}".rstrip())
    if not report.get("missing"):
        lines.append("  （无，全部交齐）")
    lines.append("")
    lines.append(f"重复提交（{len(report.get('duplicates') or [])} 人）：")
    for item in report.get("duplicates") or []:
        files = "、".join(Path(path).name for path in item.get("files", []))
        lines.append(
            f"  - {item.get('student_id', '')} {item.get('name', '')}：{files}".rstrip()
        )
    if not report.get("duplicates"):
        lines.append("  （无）")
    lines.append("")
    lines.append(f"匹配不到名单的文件（{len(report.get('unmatched_files') or [])} 份）：")
    for item in report.get("unmatched_files") or []:
        lines.append(f"  - {Path(item.get('file', '')).name}：{item.get('reason', '')}")
    if not report.get("unmatched_files"):
        lines.append("  （无）")
    lines.append("")
    lines.append(f"文件名提醒（{len(report.get('naming_warnings') or [])} 条）：")
    for item in report.get("naming_warnings") or []:
        lines.append(f"  - {Path(item.get('file', '')).name}：{item.get('reason', '')}")
    if not report.get("naming_warnings"):
        lines.append("  （无）")
    return "\n".join(lines) + "\n"


def _summary_cards(report: Report) -> str:
    counts = report.counts()
    stats = report.stats or {}
    return (
        "<div class='summary'>"
        f"<div><b class='err'>{counts['error']}</b><span class='muted'>错误</span></div>"
        f"<div><b class='warn'>{counts['warning']}</b><span class='muted'>警告</span></div>"
        f"<div><b>{counts['info']}</b><span class='muted'>提示</span></div>"
        f"<div><b>{stats.get('paragraphs', 0)}</b><span class='muted'>段落</span></div>"
        f"<div><b>{stats.get('tables', 0)}</b><span class='muted'>表格</span></div>"
        f"<div><b>{stats.get('sections', 0)}</b><span class='muted'>节</span></div>"
        f"<div><b>{stats.get('images', 0)}</b><span class='muted'>图片</span></div>"
        "</div>"
    )


# ---------------------------------------------------------------------------
# 请求处理
# ---------------------------------------------------------------------------
_MOBILE_CSS = """
#m-wrap { max-width: 720px; margin: 0 auto; }
#m-wrap input:not([type=file]), #m-wrap select { width:100%; padding:9px 10px; border:1px solid var(--line);
  border-radius:8px; margin:6px 0; background:#fff; font-size:15px; }
.m-list .m-item { border:1px solid var(--line); border-radius:10px; padding:10px 12px; margin:8px 0; background:#fff; cursor:pointer; }
.m-list .m-item:active { background:#f1f5f9; }
.m-item .m-title { font-weight:700; font-size:15px; }
.m-item .m-meta { color:var(--muted); font-size:12px; margin-top:3px; }
.m-item .m-code { font-family:ui-monospace,Consolas,monospace; color:var(--accent); font-weight:700; }
.m-actions { display:flex; gap:10px; flex-wrap:wrap; margin-top:10px; }
.m-sub { border:1px solid var(--line); border-radius:10px; padding:10px 12px; margin:8px 0; background:#fff; }
.m-sub .m-row { display:flex; justify-content:space-between; gap:8px; font-size:14px; }
.m-sub .m-small { color:var(--muted); font-size:12px; margin-top:3px; }
.m-pass { color:var(--ok); font-weight:700; } .m-fail { color:var(--err); font-weight:700; }
.m-summary { display:flex; gap:10px; flex-wrap:wrap; margin:10px 0; }
.m-summary span { background:#f8fafc; border:1px solid var(--line); border-radius:8px; padding:6px 10px; font-size:13px; }
.m-toast { position:fixed; left:50%; bottom:24px; transform:translateX(-50%); background:#0f172a; color:#fff;
  padding:10px 16px; border-radius:999px; font-size:13px; opacity:0; transition:.2s; pointer-events:none; z-index:99; }
.m-toast.show { opacity:.95; }
"""

_MOBILE_JS = """
(function () {
  var state = { code: null, profiles: [] };
  var NICK_KEY = 'tas_nickname';
  function $(id) { return document.getElementById(id); }
  function esc(text) {
    return String(text === undefined || text === null ? '' : text)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }
  function toast(message) {
    var el = document.createElement('div');
    el.className = 'm-toast'; el.textContent = message;
    document.body.appendChild(el);
    setTimeout(function () { el.classList.add('show'); }, 10);
    setTimeout(function () { el.classList.remove('show'); }, 2400);
    setTimeout(function () { el.remove(); }, 2800);
  }
  async function api(path, options) {
    options = options || {};
    var response = await fetch(path, options);
    var data = null;
    try { data = await response.json(); } catch (err) { data = null; }
    if (!response.ok || !data || data.ok === false) {
      throw new Error((data && data.error) || ('请求失败：HTTP ' + response.status));
    }
    return data;
  }
  function nick() {
    var value = localStorage.getItem(NICK_KEY) || '';
    if (!value) {
      value = (prompt('请输入你的昵称（仅用于展示）') || '').trim() || '匿名同学';
      localStorage.setItem(NICK_KEY, value);
    }
    return value;
  }
  async function loadProfiles() {
    try {
      var data = await api('/api/profiles');
      state.profiles = data.profiles || [];
      var select = $('c-ruleset');
      select.innerHTML = state.profiles.map(function (item) {
        return '<option value="' + esc(item.id) + '">' + esc(item.name) + '</option>';
      }).join('');
    } catch (err) { toast(err.message); }
  }
  async function loadCollections() {
    try {
      var data = await api('/api/collections');
      var box = $('collections');
      if (!data.collections || !data.collections.length) {
        box.innerHTML = '<p class="muted">还没有接龙，先在下面发起一个吧。</p>';
        return;
      }
      box.innerHTML = data.collections.map(function (item) {
        return '<div class="m-item" onclick="openCollection(\\'' + esc(item.code) + '\\')">'
          + '<div class="m-title">' + esc(item.title) + '</div>'
          + '<div class="m-meta">发起人 ' + esc(item.owner) + ' ｜ 规范 ' + esc(item.ruleset)
          + ' ｜ 已提交 ' + (item.submission_count || 0) + ' 份 ｜ code '
          + '<span class="m-code">' + esc(item.code) + '</span></div></div>';
      }).join('');
    } catch (err) { toast(err.message); }
  }
  async function createCollection() {
    var title = ($('c-title').value || '').trim();
    if (!title) { toast('请填写接龙标题'); return; }
    try {
      var data = await api('/api/collections', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          title: title, owner: nick(), ruleset: $('c-ruleset').value,
          note: ($('c-note').value || '').trim()
        })
      });
      toast('接龙已创建：' + data.collection.code);
      $('c-title').value = ''; $('c-note').value = '';
      await loadCollections();
      openCollection(data.collection.code);
    } catch (err) { toast(err.message); }
  }
  async function openCollection(code) {
    state.code = code;
    $('create-card').style.display = 'none';
    $('detail-card').style.display = 'block';
    try {
      var data = await api('/api/collections/' + encodeURIComponent(code));
      var collection = data.collection;
      $('d-title').textContent = collection.title;
      $('d-meta').textContent = 'code ' + collection.code + ' ｜ 发起人 ' + collection.owner
        + ' ｜ 规范 ' + collection.ruleset + (collection.due ? ' ｜ 截止 ' + collection.due : '')
        + (collection.note ? ' ｜ ' + collection.note : '');
      $('d-student').value = localStorage.getItem(NICK_KEY) || '';
      renderSubmissions(data.submissions || []);
      $('d-summary').innerHTML = '';
    } catch (err) { toast(err.message); }
  }
  function renderSubmissions(items) {
    var box = $('d-submissions');
    if (!items.length) { box.innerHTML = '<p class="muted">还没有人提交。</p>'; return; }
    box.innerHTML = items.map(function (item) {
      var badge = item.status === '通过' ? 'm-pass' : 'm-fail';
      var counts = item.counts || {};
      return '<div class="m-sub"><div class="m-row"><b>' + esc(item.student) + '</b>'
        + '<span class="' + badge + '">' + esc(item.status) + '</span></div>'
        + '<div class="m-small">' + esc(item.filename) + ' ｜ 错误 ' + (counts.error || 0)
        + ' ｜ 警告 ' + (counts.warning || 0) + ' ｜ 提交 ' + esc(item.created_at)
        + (item.error ? ' ｜ ' + esc(item.error) : '') + '</div>'
        + '<div class="m-actions"><button class="secondary" onclick="showReport(\\''
        + esc(item.id) + '\\')">查看报告</button></div></div>';
    }).join('');
  }
  async function submitDocument() {
    var fileInput = $('d-file');
    if (!fileInput.files || !fileInput.files.length) { toast('请先选择一个 .docx 文件'); return; }
    var student = ($('d-student').value || '').trim() || nick();
    localStorage.setItem(NICK_KEY, student);
    var form = new FormData();
    form.append('student', student);
    form.append('file', fileInput.files[0]);
    try {
      var data = await api('/api/collections/' + encodeURIComponent(state.code) + '/submit', {
        method: 'POST', body: form
      });
      toast('提交成功：' + (data.submission.counts || {}).error + ' 个错误');
      fileInput.value = '';
      openCollection(state.code);
    } catch (err) { toast(err.message); }
  }
  async function reviewCollection() {
    try {
      var data = await api('/api/collections/' + encodeURIComponent(state.code) + '/review', { method: 'POST' });
      var summary = data.summary || {};
      var top = (summary.top_rules || []).slice(0, 5).map(function (pair) {
        return '<li>' + esc(pair[0]) + '：' + pair[1] + ' 次</li>';
      }).join('');
      $('d-summary').innerHTML = '<h3>一键审核结果</h3>'
        + '<div class="m-summary"><span>文档 ' + (summary.total || 0) + '</span>'
        + '<span>检查成功 ' + (summary.checked || 0) + '</span>'
        + '<span>失败 ' + (summary.failed || 0) + '</span>'
        + '<span><b>' + Math.round((summary.pass_rate || 0) * 100) + '%</b> 通过率</span>'
        + '<span class="m-fail">错误 ' + ((summary.counts || {}).error || 0) + '</span>'
        + '<span class="m-warn">警告 ' + ((summary.counts || {}).warning || 0) + '</span></div>'
        + (top ? '<b>共性问题 Top</b><ul>' + top + '</ul>' : '');
      renderSubmissions(data.submissions || []);
      toast('审核完成：通过率 ' + Math.round((summary.pass_rate || 0) * 100) + '%');
    } catch (err) { toast(err.message); }
  }
  async function showReport(submissionId) {
    try {
      var data = await api('/api/collections/' + encodeURIComponent(state.code)
        + '/submissions/' + encodeURIComponent(submissionId) + '/report');
      var report = data.report || {};
      var issues = (report.issues || []).slice(0, 12).map(function (issue) {
        return '<li><span class="' + (issue.severity === 'error' ? 'm-fail' : '') + '">'
          + esc(issue.severity) + '</span> ' + esc(issue.rule_id) + ' '
          + esc(issue.locator) + '：' + esc(issue.message) + '</li>';
      }).join('');
      var counts = report.counts || {};
      $('d-summary').innerHTML = '<h3>检查报告</h3>'
        + '<div class="m-summary"><span>错误 ' + (counts.error || 0) + '</span>'
        + '<span>警告 ' + (counts.warning || 0) + '</span>'
        + '<span>提示 ' + (counts.info || 0) + '</span></div>'
        + (issues ? '<ul>' + issues + '</ul>' : '<p class="muted">无问题。</p>');
      window.scrollTo({ top: 0, behavior: 'smooth' });
    } catch (err) { toast(err.message); }
  }
  function backToList() {
    state.code = null;
    $('detail-card').style.display = 'none';
    $('create-card').style.display = 'block';
    loadCollections();
  }
  window.createCollection = createCollection;
  window.openCollection = openCollection;
  window.submitDocument = submitDocument;
  window.reviewCollection = reviewCollection;
  window.showReport = showReport;
  window.backToList = backToList;
  loadProfiles().then(loadCollections);
})();
"""


class TasWebHandler(BaseHTTPRequestHandler):
    server_version = "tas-web/0.1"

    # -- 输出 --------------------------------------------------------------
    def _send(self, status: int, body: bytes, content_type: str = "text/html; charset=utf-8") -> None:
        if os.environ.get("TAS_WEB_VERBOSE") == "1":
            try:
                client = self.client_address[0]
            except Exception:  # noqa: BLE001
                client = "?"
            print(
                f"[tas] {client} {getattr(self, 'command', '?')} {getattr(self, 'path', '?')}"
                f" -> {int(status)} ({len(body)} bytes)",
                file=sys.stderr,
                flush=True,
            )
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def _send_json(self, payload: dict, status: int = 200) -> None:
        body = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
        self._send(status, body, "application/json; charset=utf-8")

    def _read_json(self) -> dict:
        length = int(self.headers.get("Content-Length", "0") or 0)
        raw = self.rfile.read(length) if length > 0 else b""
        if not raw:
            return {}
        try:
            data = json.loads(raw.decode("utf-8"))
        except json.JSONDecodeError as exc:
            raise ValueError(f"JSON 解析失败：{exc}") from exc
        return data if isinstance(data, dict) else {}

    def log_message(self, fmt, *args):  # 安静一点，避免刷屏
        return

    def handle_one_request(self) -> None:
        # 手机浏览器/小程序关闭 keep-alive 连接时，避免刷一堆 ConnectionResetError 堆栈
        try:
            super().handle_one_request()
        except (ConnectionResetError, BrokenPipeError):
            self.close_connection = True

    # -- 解析 multipart ----------------------------------------------------
    def _read_multipart(self) -> tuple[dict[str, str], list[dict]]:
        content_type = self.headers.get("Content-Type", "")
        length = int(self.headers.get("Content-Length", "0"))
        if length <= 0 or length > MAX_UPLOAD_BYTES * 10:
            raise ValueError("请求为空或过大")
        body = self.rfile.read(length)
        raw = (
            f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8") + body
        )
        message = BytesParser(policy=email_policy).parsebytes(raw)
        fields: dict[str, str] = {}
        files: list[dict] = []
        parts = message.iter_parts() if message.is_multipart() else []
        for part in parts:
            name = part.get_param("name", header="content-disposition")
            filename = part.get_filename()
            payload = part.get_payload(decode=True) or b""
            if filename:
                files.append({"field": name, "filename": filename, "data": payload})
            else:
                fields[str(name)] = payload.decode("utf-8", errors="replace")
        return fields, files

    def _save_upload(
        self,
        token: str,
        uploaded: dict,
        index: int = 0,
        *,
        allowed_suffixes: tuple[str, ...] = (".docx",),
    ) -> Path:
        directory = STORE_ROOT / token / "uploads"
        directory.mkdir(parents=True, exist_ok=True)
        name = _safe_name(uploaded["filename"])
        suffix = Path(name).suffix.lower()
        if suffix not in allowed_suffixes:
            allowed = " / ".join(allowed_suffixes)
            raise ValueError(f"只支持 {allowed}（.doc 请先另存为 .docx）")
        if len(uploaded["data"]) > MAX_UPLOAD_BYTES:
            raise ValueError(f"文件过大（> {MAX_UPLOAD_BYTES // 1024 // 1024}MB）：{name}")
        target = directory / f"{index:02d}_{name}"
        target.write_bytes(uploaded["data"])
        return target

    # -- 路由 --------------------------------------------------------------
    def do_GET(self):  # noqa: N802
        path = urlparse(self.path).path
        try:
            if path.startswith("/api/"):
                return self._api_get(path)
            if path in ("/", "/index.html"):
                return self._home()
            if path == "/m":
                return self._mobile()
            if path == "/health":
                return self._send_json({"ok": True})
            if path.startswith("/download/"):
                return self._download(path)
            self._send(HTTPStatus.NOT_FOUND, _page("404", "<div class='card'><h2>404</h2><p>页面不存在。</p></div>"))
        except Exception as exc:  # noqa: BLE001
            if path.startswith("/api/"):
                return self._send_json({"ok": False, "error": str(exc)}, HTTPStatus.BAD_REQUEST)
            self._send(HTTPStatus.BAD_REQUEST, _page("出错", f"<div class='card'><h2>请求出错</h2><p>{_esc(exc)}</p></div>"))

    def do_POST(self):  # noqa: N802
        path = urlparse(self.path).path
        try:
            if path.startswith("/api/"):
                return self._api_post(path)
            fields, files = self._read_multipart()
            if path == "/student/check":
                return self._student_check(fields, files)
            if path == "/student/fix":
                return self._student_fix(fields)
            if path == "/spec/parse":
                return self._spec_parse(fields, files)
            if path == "/spec/apply":
                return self._spec_apply(fields)
            if path == "/teacher/batch":
                return self._teacher_batch(fields, files)
            self._send(HTTPStatus.NOT_FOUND, _page("404", "<div class='card'><h2>404</h2></div>"))
        except Exception as exc:  # noqa: BLE001
            if path.startswith("/api/"):
                return self._send_json({"ok": False, "error": str(exc)}, HTTPStatus.BAD_REQUEST)
            self._send(HTTPStatus.BAD_REQUEST, _page("出错", f"<div class='card'><h2>请求出错</h2><p>{_esc(exc)}</p></div>"))

    # -- 接龙 API（供小程序 / H5 / 第三方调用）------------------------------
    def _api_get(self, path: str) -> None:
        from .. import collections

        if path == "/api/diag":
            # 真机网络排查：手机浏览器打开 http://<电脑IP>:8765/api/diag
            # 能看到 client 就是手机 IP → 说明“手机到电脑”这一层是通的。
            address = self.server.server_address
            host = str(address[0])
            port = int(address[1])
            return self._send_json(
                {
                    "ok": True,
                    "service": "tas-web",
                    "time": datetime.now().astimezone().isoformat(timespec="seconds"),
                    "client": self.client_address[0],
                    "server_host": host,
                    "server_port": port,
                    "base_url": f"http://{host}:{port}",
                    "verbose": os.environ.get("TAS_WEB_VERBOSE") == "1",
                }
            )
        if path == "/api/profiles":
            from ..config import DEFAULT_PROFILE_ID

            return self._send_json(
                {
                    "default": DEFAULT_PROFILE_ID,
                    "profiles": [
                        {"id": p.id, "name": p.name, "discipline": p.discipline}
                        for p in list_profiles()
                    ],
                }
            )
        if path == "/api/collections":
            return self._send_json({"ok": True, "collections": collections.list_collections()})
        match = re.match(r"^/api/collections/([^/]+)/submissions/([^/]+)/report$", path)
        if match:
            report = collections.get_submission_report(match.group(1), match.group(2))
            return self._send_json({"ok": True, "report": report})
        match = re.match(r"^/api/collections/([^/]+)$", path)
        if match:
            payload = collections.get_collection(match.group(1))
            return self._send_json({"ok": True, **payload})
        return self._send_json({"ok": False, "error": "接口不存在"}, HTTPStatus.NOT_FOUND)

    def _api_post(self, path: str) -> None:
        from .. import collections

        if path == "/api/collections":
            body = self._read_json()
            item = collections.create_collection(
                title=body.get("title", ""),
                owner=body.get("owner", ""),
                ruleset=body.get("ruleset") or "",
                note=body.get("note", ""),
                due=body.get("due", ""),
            )
            return self._send_json({"ok": True, "collection": item})
        match = re.match(r"^/api/collections/([^/]+)/submit$", path)
        if match:
            fields, files = self._read_multipart()
            if not files:
                raise ValueError("请选择要提交的 .docx 文件")
            uploaded = files[0]
            if os.environ.get("TAS_WEB_VERBOSE") == "1":
                print(
                    f"[tas] upload from {self.client_address[0]}:"
                    f" student={fields.get('student', '')!r}"
                    f" filename={uploaded['filename']!r}"
                    f" size={len(uploaded['data'])}",
                    file=sys.stderr,
                    flush=True,
                )
            submission = collections.add_submission(
                match.group(1),
                fields.get("student", ""),
                uploaded["filename"],
                uploaded["data"],
            )
            return self._send_json({"ok": True, "submission": submission})
        match = re.match(r"^/api/collections/([^/]+)/review$", path)
        if match:
            result = collections.review_collection(match.group(1))
            return self._send_json({"ok": True, **result})
        return self._send_json({"ok": False, "error": "接口不存在"}, HTTPStatus.NOT_FOUND)

    # -- 页面 --------------------------------------------------------------
    def _home(self) -> None:
        body = f"""
<div class="grid">
  <section class="card">
    <h2>学生端：检查 + 一键修复</h2>
    <p class="sub">上传一篇 .docx，按所选规范检查；确认后可一键修复格式并下载新文件（原文件不覆盖，内容不改）。</p>
    <form method="post" action="/student/check" enctype="multipart/form-data">
      <label>选择规范</label>{_profiles_select()}
      <label>论文 / 文档（.docx）</label>
      {_dropzone("student-file", "file", hint="把论文拖进来，或点击选择 · 只支持 .docx")}
      <label class="check"><input type="checkbox" name="preview" value="1" checked>
        生成页面预览（首次约 10–15 秒，只多跑一次 Word 渲染）</label>
      <button type="submit">检查格式</button>
    </form>
  </section>
  <section class="card">
    <h2>教师端：批量检查 + 班级汇总</h2>
    <p class="sub">选择规范，一次拖入/选择多份学生文档（或整个文件夹）；输出逐份问题清单、班级通过率、共性问题 Top，可下载 xlsx/csv/html + <b>学生反馈包</b>。</p>
    <form method="post" action="/teacher/batch" enctype="multipart/form-data">
      <label>选择规范</label>{_profiles_select()}
      <label>学生文档（可多份 .docx）</label>
      {_dropzone("teacher-files", "files", multiple=True, hint="把多份论文一起拖进来，或点击选择 · 支持批量")}
      <div class="folder-row">
        <button type="button" class="folder-btn" onclick="document.getElementById('teacher-dir').click()">选择整个文件夹（只收 .docx）</button>
        <span class="folder-hint">微信收到的文件先存到一个文件夹，可整批选择</span>
        <input type="file" id="teacher-dir" webkitdirectory multiple hidden>
      </div>
      <label>名单核对（可选，每行「学号 姓名」或「学号,姓名」）</label>
      <textarea name="roster" rows="3" placeholder="例如：
23050301001 张三
23050301002 李四"></textarea>
      <label class="check"><input type="checkbox" name="fix" value="1">
        同时生成安全修复版（只改格式，不改内容/文献，原文件不覆盖）</label>
      <label class="check"><input type="checkbox" name="preview" value="1">
        生成首页缩略图（Word 渲染，每份约 2–10 秒）</label>
      <label class="check"><input type="checkbox" name="refs" value="1">
        文献真实性核验（联网，每份最多 5 条，较慢）</label>
      <button type="submit">批量检查</button>
    </form>
    <p class="small muted">说明：只做检查与汇总，不修改学生文档；「学生反馈包」里每份文档一个 .txt，可直接发给学生；文献真实性核验在 CLI（tas refs）完成。</p>
  </section>
  <section class="card" style="grid-column:1/-1">
    <h2>规范管理（教师端）：规范文档 → 规则集</h2>
    <p class="sub">把规范文档（.docx / .txt / .md）拖进来，DeepSeek 逐条抽取规则候选；审阅后生成新的 profile，
    可直接用于学生检查/修复。需要启动 Web 的终端里有环境变量 TAS_LLM_API_KEY。</p>
    <form method="post" action="/spec/parse" enctype="multipart/form-data">
      <label>基准规范（用于逐条 diff，通常选学院默认）</label>{_profiles_select("base_profile")}
      <label>规范文档</label>
      {_dropzone("spec-file", "file", accept=".docx,.txt,.md", title="规范文档（.docx/.txt/.md）", hint="把规范文档拖进来，或点击选择 · 支持 .docx/.txt/.md")}
      <button type="submit">抽取规则候选（LLM，约 5–30 秒）</button>
    </form>
    <p class="small muted">流程：抽取候选 → 审阅 diff → 确认生成 profile。只有规范文档文本会发给 DeepSeek，学生文档不上传。</p>
  </section>
</div>"""
        self._send(HTTPStatus.OK, _page("首页", body))

    # -- 手机 H5：文档接龙（不拆教师/学生端，便于真机快速验证）---------------
    def _mobile(self) -> None:
        body = f"""
<style>{_MOBILE_CSS}</style>
<div id="m-wrap">
  <section class="card" id="create-card">
    <h2>文档接龙（内测 H5）</h2>
    <p class="sub">发起接龙 → 同学提交 .docx → 一键审核格式是否规范。手机浏览器可直接用；
    微信小程序源码见仓库 <code>miniprogram/</code>。</p>
    <h3>进行中的接龙</h3>
    <div id="collections" class="m-list"><p class="muted">加载中…</p></div>
    <h3>发起接龙</h3>
    <input id="c-title" placeholder="标题，如：第 3 次课程设计报告">
    <input id="c-owner" placeholder="你的昵称（自动记住）">
    <select id="c-ruleset"></select>
    <input id="c-note" placeholder="说明（可选）">
    <button onclick="createCollection()">发起接龙</button>
  </section>
  <section class="card" id="detail-card" style="display:none">
    <button class="secondary" onclick="backToList()">← 返回列表</button>
    <h2 id="d-title"></h2>
    <p class="sub" id="d-meta"></p>
    <div id="d-summary"></div>
    <h3>提交的文档</h3>
    <div id="d-submissions"></div>
    <h3>提交我的文档</h3>
    <input id="d-student" placeholder="你的昵称">
    <input type="file" id="d-file" accept=".docx">
    <div class="m-actions">
      <button onclick="submitDocument()">提交</button>
      <button class="secondary" onclick="reviewCollection()">一键审核格式</button>
    </div>
  </section>
  <p class="small muted">无账号，昵称仅用于展示；数据保存在本机 <code>data/</code> 目录。
  微信小程序见 <code>miniprogram/</code>，真机联调把 <code>utils/api.js</code> 的 BASE_URL 改成电脑局域网 IP。</p>
</div>
<script>{_MOBILE_JS}</script>"""
        self._send(HTTPStatus.OK, _page("文档接龙", body))

    # -- 学生端 ------------------------------------------------------------
    def _student_check(self, fields: dict, files: list[dict]) -> None:
        if not files:
            raise ValueError("请上传一个 .docx 文件")
        token = _new_token()
        path = self._save_upload(token, files[0])
        profile, ruleset = _resolve(fields.get("ruleset"))
        model = DocModel.load(path)
        roles = classify(model, profile.discipline)
        issues = run_checks(model, roles, ruleset, profile=profile.discipline)
        report = Report(
            path=path,
            profile=profile.discipline,
            issues=issues,
            stats=model.stats(),
            rules_run=len([rid for rid in IMPLEMENTED_RULES if rid in ruleset.rules]),
            rules_total=len(ruleset),
        )
        directory = STORE_ROOT / token
        (directory / "report.json").write_text(
            json.dumps(report.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8"
        )
        (directory / "report.md").write_text(report.render_markdown(), encoding="utf-8")

        preview_requested = fields.get("preview") == "1"
        preview_pages: list = []
        preview_note = ""
        if preview_requested:
            (directory / ".preview").write_text("1", encoding="utf-8")
            try:
                from ..render import render_preview

                preview_pages = render_preview(path, directory / "preview", max_pages=2)
            except Exception as exc:  # noqa: BLE001 - 预览失败不影响检查主流程
                preview_note = f"页面预览生成失败（不影响检查）：{exc}"

        body = f"""
<div class="card">
  <h2>检查结果：{_esc(files[0]['filename'])}</h2>
  <p class="sub">规范：{_esc(profile.name)}（{_esc(profile.discipline)}）｜ token：{_esc(token)}</p>
  {_summary_cards(report)}
  {_preview_html(token, preview_pages, preview_note)}
  {_issues_table(report)}
  <form method="post" action="/student/fix">
    <input type="hidden" name="token" value="{_esc(token)}">
    <input type="hidden" name="ruleset" value="{_esc(profile.id)}">
    {'<input type="hidden" name="preview" value="1">' if preview_requested else ''}
    <button type="submit">一键修复格式（只改格式，不改内容）</button>
  </form>
  <p class="small"><a href="/download/{_esc(token)}/report.md">下载 Markdown 报告</a> ｜
  <a href="/download/{_esc(token)}/report.json">下载 JSON 报告</a> ｜ <a href="/">返回首页</a></p>
</div>"""
        self._send(HTTPStatus.OK, _page("检查结果", body))

    def _student_fix(self, fields: dict) -> None:
        from ..fixer import fix_document

        token = fields.get("token", "")
        directory = _store_dir(token)
        source = next((directory / "uploads").glob("*.docx"), None)
        if source is None:
            raise FileNotFoundError("找不到原始文件，请重新上传")
        profile, ruleset = _resolve(fields.get("ruleset"))
        out = directory / "fixed.docx"
        result = fix_document(source, out, profile=profile.discipline, ruleset=ruleset)
        (directory / "fix.json").write_text(
            json.dumps(result.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8"
        )
        preview_pages: list = []
        preview_note = ""
        if (directory / ".preview").exists():
            try:
                from ..render import render_preview

                preview_pages = render_preview(out, directory / "fixed_preview", max_pages=2)
            except Exception as exc:  # noqa: BLE001 - 预览失败不影响修复结果
                preview_note = f"修复后页面预览生成失败：{exc}"
        actions = [
            f"<tr><td>{_esc(a.rule_id)}</td><td>{_esc(a.locator)}</td><td>{_esc(a.description)}</td></tr>"
            for a in result.actions[:200]
        ]
        actions_html = (
            "<table><thead><tr><th>规则</th><th>位置</th><th>修改说明</th></tr></thead><tbody>"
            + "".join(actions)
            + "</tbody></table>"
            if actions
            else "<p class='muted'>没有需要修复的格式问题。</p>"
        )
        remaining = result.recheck_issues
        body = f"""
<div class="card">
  <h2>修复完成</h2>
  <p class="sub">修复动作 {len(result.actions)} 项 ｜ 剩余问题 {len(remaining)} 项 ｜
  内容校验：{'✅ 内容 hash 一致（文字未改）' if result.content_hash_before == result.content_hash_after else '❌ 异常'}</p>
  {actions_html}
  {_preview_html(token, preview_pages, preview_note)}
  <h3>剩余问题（需人工确认，涉及文字/编号/题注/文献的不会自动改）</h3>
  {_issues_table(type('R', (), {'issues': remaining})(), limit=80) if remaining else "<p class='muted'>无。</p>"}
  <p><b>下载修复后的文档：</b> <a href="/download/{_esc(token)}/fixed.docx">fixed.docx</a> ｜
  <a href="/download/{_esc(token)}/fix.json">修复报告 JSON</a> ｜
  <a href="/download/{_esc(token)}/report.md">原始检查报告</a> ｜ <a href="/">返回首页</a></p>
  <p class="small muted">原文件未被覆盖；参考文献、编号、题注、字数类问题只给建议。</p>
</div>"""
        self._send(HTTPStatus.OK, _page("修复完成", body))

    # -- 规范管理（教师端：规范文档 → 规则候选 → profile）-------------------
    def _render_error(self, title: str, message: str) -> None:
        body = (
            f"<div class='card'><h2>{_esc(title)}</h2>"
            f"<pre>{_esc(message)}</pre>"
            "<p><a href='/'>返回首页</a></p></div>"
        )
        self._send(HTTPStatus.OK, _page(title, body))

    def _spec_parse(self, fields: dict, files: list[dict]) -> None:
        from ..llm import LLMConfigError, LLMError, chat_json, load_config
        from ..spec_io import extract_spec_text
        from ..spec_profile import diff_candidates, extract_candidates, render_report_markdown

        if not files:
            raise ValueError("请上传一个规范文档（.docx/.txt/.md）")
        token = _new_token()
        path = self._save_upload(
            token, files[0], allowed_suffixes=(".docx", ".txt", ".md", ".markdown")
        )
        profile, ruleset = _resolve(fields.get("base_profile"))
        try:
            spec_text = extract_spec_text(path)
        except ValueError as exc:
            return self._render_error("规范文档读取失败", str(exc))

        try:
            config = load_config()

            def _chat(messages):
                return chat_json(messages, config=config)

            report = extract_candidates(
                spec_text,
                ruleset,
                chat=_chat,
                model=config.model,
                base_profile=profile.id,
            )
        except LLMConfigError as exc:
            return self._render_error(
                "LLM 配置错误",
                f"{exc}\n\n设置方式（在启动 Web 的终端里）：\n"
                '  setx TAS_LLM_API_KEY "sk-..."\n'
                "设置后重启 Web 服务再试。",
            )
        except LLMError as exc:
            return self._render_error("LLM 调用失败", str(exc))

        diff = diff_candidates(report, ruleset, base_profile=profile.id)
        directory = STORE_ROOT / token
        payload = {
            "base_profile": profile.id,
            "source_file": files[0]["filename"],
            "model": report.model,
            "extraction": report.to_dict(),
            "diff": diff.to_dict(),
        }
        (directory / "candidates.json").write_text(
            json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        (directory / "diff.md").write_text(
            render_report_markdown(report, diff), encoding="utf-8"
        )

        counts = diff.counts()
        rows = []
        for item in diff.diffs:
            if item.action == "keep":
                continue
            candidate = item.candidate
            expected_text = json.dumps(candidate.expected, ensure_ascii=False)
            changes = "；".join(item.changes[:6])
            rows.append(
                "<tr>"
                f"<td><span class='badge {'warn' if item.action == 'modify' else 'ok'}'>{_esc(item.action)}</span></td>"
                f"<td>{_esc(candidate.id)}<br><span class='small muted'>{_esc(candidate.name)}</span></td>"
                f"<td class='small'>{_esc(candidate.source)}</td>"
                f"<td class='small'>{_esc(candidate.spec)[:140]}<br>"
                f"<span class='muted'>expected: {_esc(expected_text)[:180]}</span>"
                + (f"<br><span class='muted'>变化：{_esc(changes)}</span>" if changes else "")
                + "</td></tr>"
            )
        rows_html = (
            "<table><thead><tr><th>动作</th><th>规则</th><th>出处</th><th>规范/期望</th></tr></thead><tbody>"
            + "".join(rows)
            + "</tbody></table>"
            if rows
            else "<p class='muted'>全部候选与现有规则一致（keep），无需修改；可直接返回。</p>"
        )
        unmapped_html = "".join(
            f"<li>{_esc(item.get('source', ''))} {_esc(item.get('text', ''))[:120]}</li>"
            for item in diff.unmapped[:15]
        ) or "<li class='muted'>无</li>"

        body = f"""
<div class="card">
  <h2>规范抽取结果：{_esc(files[0]['filename'])}</h2>
  <p class="sub">模型：{_esc(report.model)} ｜ 基准：{_esc(profile.name)} ｜ token：{_esc(token)}</p>
  <div class="summary">
    <div><b>{counts.get('keep', 0)}</b><span class="muted">keep</span></div>
    <div><b class="warn">{counts.get('modify', 0)}</b><span class="muted">modify</span></div>
    <div><b class="ok">{counts.get('new', 0)}</b><span class="muted">new</span></div>
    <div><b>{len(diff.unmapped)}</b><span class="muted">unmapped</span></div>
    <div><b>{len(diff.uncovered_existing)}</b><span class="muted">未覆盖现有规则</span></div>
  </div>
  {rows_html}
  <h3>unmapped（无法程序化检查，需人工/流程处理）</h3>
  <ul>{unmapped_html}</ul>
  <p class="small"><a href="/download/{_esc(token)}/diff.md">下载完整 diff.md</a> ｜
  <a href="/download/{_esc(token)}/candidates.json">下载 candidates.json</a></p>
  <h3>确认生成新 profile</h3>
  <form method="post" action="/spec/apply">
    <input type="hidden" name="token" value="{_esc(token)}">
    <input type="hidden" name="source_file" value="{_esc(files[0]['filename'])}">
    <label>新 profile id（英文/数字/下划线，用于 --ruleset）</label>
    <input type="text" name="profile_id" placeholder="course_xxx_2025" required
           pattern="[A-Za-z0-9][A-Za-z0-9_-]{{1,59}}">
    <label>显示名称</label>
    <input type="text" name="name" placeholder="XX 课程 2025 论文格式要求" required>
    <label>学科体系</label>
    <select name="discipline">
      <option value="sci_eng">理工农（第N章 / N.N）</option>
      <option value="humanities">文经管（一、/（一）/1.）</option>
    </select>
    <label class="check"><input type="checkbox" name="register" value="1" checked>
      注册到 profile 列表（可直接用于检查/修复）</label>
    <button type="submit">生成新 profile</button>
  </form>
  <p class="small"><a href="/">返回首页</a></p>
</div>"""
        self._send(HTTPStatus.OK, _page("规范抽取结果", body))

    def _spec_apply(self, fields: dict) -> None:
        from ..config import DEFAULT_PROFILE, PROFILES, profiles_index_path, project_root
        from ..spec_profile import apply_candidates_file

        token = fields.get("token", "")
        directory = _store_dir(token)
        candidates = directory / "candidates.json"
        if not candidates.is_file():
            raise FileNotFoundError("找不到候选文件，请重新上传规范并抽取")
        profile_id = (fields.get("profile_id") or "").strip()
        if not re.match(r"^[A-Za-z0-9][A-Za-z0-9_-]{1,59}$", profile_id):
            raise ValueError("profile id 不合法（英文/数字/下划线，2–60 字符）")
        name = (fields.get("name") or "").strip()
        if not name:
            raise ValueError("请填写 profile 名称")
        discipline = fields.get("discipline") or DEFAULT_PROFILE
        if discipline not in PROFILES:
            raise ValueError("学科体系不合法")
        register = fields.get("register") == "1"
        out_path = (
            project_root() / "profiles" / f"{profile_id}.rules.json"
            if register
            else directory / "profile.rules.json"
        )
        result = apply_candidates_file(
            candidates,
            profile_id=profile_id,
            name=name,
            discipline=discipline,
            out_path=out_path,
            index_path=profiles_index_path() if register else None,
            source_spec=fields.get("source_file", ""),
        )
        if register:
            download_html = (
                "<p>规则文件已写入：<code>profiles/"
                f"{_esc(profile_id)}.rules.json</code>，并注册到 profile 列表（刷新首页下拉可见）。</p>"
            )
        else:
            download_html = (
                f"<p><a href='/download/{_esc(token)}/profile.rules.json'>下载规则文件</a>"
                "（未注册；如需使用请手动加入 profiles/index.json）</p>"
            )
        body = f"""
<div class="card">
  <h2>新 profile 已生成：{_esc(profile_id)}</h2>
  <p class="sub">名称：{_esc(name)} ｜ 学科：{_esc(discipline)} ｜ 规则总数：{result['rules_total']}
  （keep {result['keep']} ｜ modify {result['modify']} ｜ new {result['new']}）｜
  unmapped {result['unmapped']} 条</p>
  {download_html}
  <h3>下一步</h3>
  <ul>
    <li>回到首页，在学生端/教师端「选择规范」下拉中选择 <code>{_esc(profile_id)}</code>，重新检查文档；</li>
    <li>CLI：<code>tas check 论文.docx --ruleset {_esc(profile_id)}</code> / <code>tas fix ...</code>；</li>
    <li>注意：new 规则需要绑定 checker 才会参与检查（当前参数化对 keep/modify 的页面/正文/标题/前置规则生效）。</li>
  </ul>
  <p><a href="/">返回首页</a></p>
</div>"""
        self._send(HTTPStatus.OK, _page("profile 已生成", body))

    # -- 教师端 ------------------------------------------------------------
    def _teacher_batch(self, fields: dict, files: list[dict]) -> None:
        from ..batch import check_files, summarize
        from ..batch_fix import fix_files, summarize_fix, write_fix_zip
        from ..export import export_csv, export_html, export_xlsx
        from ..feedback import export_feedback_zip
        from ..roster import match_submissions, parse_roster

        if not files:
            raise ValueError("请至少上传一份 .docx")
        token = _new_token()
        saved = [self._save_upload(token, uploaded, index=index) for index, uploaded in enumerate(files)]
        profile, ruleset = _resolve(fields.get("ruleset"))
        results = check_files(saved, ruleset=ruleset, profile=profile.discipline, ruleset_id=profile.id)
        summary = summarize(results, ruleset_id=profile.id)
        directory = STORE_ROOT / token

        thumbnails = None
        thumb_summary = None
        thumb_zip = None
        if fields.get("preview") == "1":
            from ..thumbnails import (
                export_thumbnail_zip,
                render_thumbnails,
                summarize_thumbs,
                thumbnails_map,
            )

            thumb_results = render_thumbnails(
                saved, directory / "thumbnails", max_pages=1, width=900
            )
            thumbnails = thumbnails_map(thumb_results)
            thumb_summary = summarize_thumbs(thumb_results)
            if any(item.ok for item in thumb_results):
                thumb_zip = export_thumbnail_zip(
                    thumb_results, directory / "thumbnails.zip"
                )

        refs_summary = None
        refs_json = None
        refs_md = None
        if fields.get("refs") == "1":
            from ..batch_refs import verify_batch, write_refs_report

            try:
                refs_limit = max(1, min(int(fields.get("refs_limit") or 5), 20))
            except (TypeError, ValueError):
                refs_limit = 5
            refs_docs, refs_obj = verify_batch(
                saved,
                profile=profile.discipline,
                limit_per_doc=refs_limit,
                pause=0.0,
            )
            refs_json, refs_md = write_refs_report(
                refs_obj,
                refs_docs,
                directory / "refs-report.json",
                directory / "refs-report.md",
            )
            refs_summary = refs_obj.to_dict()

        export_xlsx(results, summary, directory / "batch.xlsx")
        export_csv(results, summary, directory / "batch.csv")
        (directory / "batch.html").write_text(
            export_html(summary, results, thumbnails=thumbnails), encoding="utf-8"
        )
        export_feedback_zip(
            results,
            directory / "feedback.zip",
            summary=summary,
            profile_label=profile.name,
        )

        roster_report = None
        roster_text = (fields.get("roster") or "").strip()
        if roster_text:
            roster_report = match_submissions(saved, parse_roster(roster_text))
            (directory / "roster-report.txt").write_text(
                _roster_text(roster_report), encoding="utf-8"
            )

        fix_summary = None
        if fields.get("fix") == "1":
            fix_results = fix_files(
                saved,
                directory / "fixed",
                ruleset=ruleset,
                profile=profile.discipline,
            )
            fix_summary = summarize_fix(fix_results)
            write_fix_zip(fix_results, directory / "fixed.zip")

        per_doc_rows = []
        for item in summary.per_doc:
            name = Path(item.get("path", "")).name
            per_doc_rows.append(
                "<tr>"
                f"<td>{_esc(name)}</td>"
                f"<td class='err'>{item.get('error', 0)}</td>"
                f"<td class='warn'>{item.get('warning', 0)}</td>"
                f"<td>{item.get('info', 0)}</td>"
                f"<td>{'✅' if item.get('passed') else '❌'} {_esc(item.get('status', ''))}</td>"
                f"<td>{float(item.get('duration_s', 0)):.2f}s</td>"
                "</tr>"
            )
        top_rows = "".join(
            f"<tr><td>{_esc(rule)}</td><td>{count}</td></tr>" for rule, count in summary.top_rules
        )

        roster_block = ""
        if roster_report:
            missing_items = roster_report.get("missing") or []
            duplicate_items = roster_report.get("duplicates") or []
            unmatched_items = roster_report.get("unmatched_files") or []
            warning_items = roster_report.get("naming_warnings") or []
            missing_html = "".join(
                f"<li>{_esc(item.get('student_id', ''))} {_esc(item.get('name', ''))}</li>"
                for item in missing_items
            ) or "<li class='muted'>无，全部交齐</li>"
            duplicate_html = "".join(
                f"<li>{_esc(item.get('student_id', ''))} {_esc(item.get('name', ''))}："
                + "、".join(_esc(Path(p).name) for p in item.get("files", []))
                + "</li>"
                for item in duplicate_items
            ) or "<li class='muted'>无</li>"
            unmatched_html = "".join(
                f"<li>{_esc(Path(item.get('file', '')).name)}：{_esc(item.get('reason', ''))}</li>"
                for item in unmatched_items
            ) or "<li class='muted'>无</li>"
            warning_html = "".join(
                f"<li>{_esc(Path(item.get('file', '')).name)}：{_esc(item.get('reason', ''))}</li>"
                for item in warning_items
            ) or "<li class='muted'>无</li>"
            roster_block = f"""<h3>名单核对</h3>
  <p class="sub">应交 {roster_report.get('total_roster', 0)} 人 ｜ 收到 {roster_report.get('total_files', 0)} 份 ｜ 已匹配 {roster_report.get('matched', 0)} 份</p>
  <p><b>缺交（{len(missing_items)}）</b></p><ul class="roster-list">{missing_html}</ul>
  <p><b>重复提交（{len(duplicate_items)}）</b></p><ul class="roster-list">{duplicate_html}</ul>
  <p><b>匹配不到名单（{len(unmatched_items)}）</b></p><ul class="roster-list">{unmatched_html}</ul>
  <p><b>文件名提醒（{len(warning_items)}）</b></p><ul class="roster-list">{warning_html}</ul>
  <p class="small"><a href="/download/{_esc(token)}/roster-report.txt">下载名单报告</a></p>"""

        fix_block = ""
        if fix_summary:
            before_error = int(fix_summary.get("before_counts", {}).get("error", 0))
            after_error = int(fix_summary.get("after_counts", {}).get("error", 0))
            fix_block = f"""<h3>批量安全修复</h3>
  <p class="sub">成功 {fix_summary.get('fixed', 0)} 份 ｜ 失败 {fix_summary.get('failed', 0)} 份 ｜
  错误 {before_error} → {after_error} ｜ 修改动作 {fix_summary.get('actions', 0)} 次</p>
  <p><a href="/download/{_esc(token)}/fixed.zip">下载修复包（*_fixed.docx + 修复报告）</a></p>
  <p class="small muted">只改格式，内容 hash 校验不变；参考文献/正文不动；原文件不覆盖。</p>"""

        thumb_block = ""
        if thumbnails:
            figures = "".join(
                f"<figure><a href='/download/{_esc(token)}/{_esc(relative)}' target='_blank'>"
                f"<img src='/download/{_esc(token)}/{_esc(relative)}' loading='lazy' alt='缩略图'></a>"
                f"<figcaption>{_esc(Path(name).name)}</figcaption></figure>"
                for name, relative in list(thumbnails.items())[:12]
            )
            thumb_block = (
                "<h3>首页缩略图（成功 "
                f"{thumb_summary.get('ok', 0)}/{thumb_summary.get('total', 0)}）</h3>"
                f"<div class='thumb-grid'>{figures}</div>"
            )

        refs_block = ""
        if refs_summary:
            refs_block = f"""<h3>文献真实性核验</h3>
  <p class="sub">文档 {refs_summary.get('total_docs', 0)} ｜ 文献 {refs_summary.get('total_entries', 0)} ｜
  verified {refs_summary.get('verified', 0)} ｜ likely {refs_summary.get('likely', 0)} ｜
  not_found {refs_summary.get('not_found', 0)} ｜ error {refs_summary.get('error', 0)}</p>
  <p class="small muted">not_found 仅表示未在当前索引检索到，中文期刊可能未被收录，需人工核实（未检索到 ≠ 造假）。</p>
  <p><a href="/download/{_esc(token)}/refs-report.md">下载文献核验报告（md）</a> ｜
  <a href="/download/{_esc(token)}/refs-report.json">JSON</a></p>"""

        extra_links = ""
        if fix_summary:
            extra_links += f'<a href="/download/{_esc(token)}/fixed.zip">修复包（zip）</a> ｜ '
        if roster_report:
            extra_links += f'<a href="/download/{_esc(token)}/roster-report.txt">名单报告</a> ｜ '
        if thumb_zip:
            extra_links += f'<a href="/download/{_esc(token)}/thumbnails.zip">缩略图包（zip）</a> ｜ '
        if refs_summary:
            extra_links += f'<a href="/download/{_esc(token)}/refs-report.md">文献核验报告</a> ｜ '
        body = f"""
<div class="card">
  <h2>班级批量检查结果</h2>
  <p class="sub">规范：{_esc(profile.name)} ｜ 文件 {summary.total} 份 ｜ 检查成功 {summary.checked} ｜
  失败 {summary.failed} ｜ <b>通过率 {summary.pass_rate:.1%}</b></p>
  <div class="summary">
    <div><b class="err">{summary.counts.get('error',0)}</b><span class="muted">错误总数</span></div>
    <div><b class="warn">{summary.counts.get('warning',0)}</b><span class="muted">警告总数</span></div>
    <div><b>{summary.counts.get('info',0)}</b><span class="muted">提示总数</span></div>
  </div>
  <h3>共性问题 Top {len(summary.top_rules)}</h3>
  <table><thead><tr><th>规则</th><th>命中次数</th></tr></thead><tbody>{top_rows or "<tr><td colspan='2' class='muted'>无</td></tr>"}</tbody></table>
  <h3>逐份结果</h3>
  <table><thead><tr><th>文件</th><th>错误</th><th>警告</th><th>提示</th><th>结论</th><th>耗时</th></tr></thead>
  <tbody>{''.join(per_doc_rows)}</tbody></table>
  {roster_block}
  {fix_block}
  {thumb_block}
  {refs_block}
  <p><b>下载：</b> <a href="/download/{_esc(token)}/batch.xlsx">汇总 Excel（xlsx）</a> ｜
  <a href="/download/{_esc(token)}/batch.csv">明细 CSV</a> ｜
  <a href="/download/{_esc(token)}/batch.html">HTML 报告</a> ｜
  <a href="/download/{_esc(token)}/feedback.zip">学生反馈（zip）</a> ｜
  {extra_links}<a href="/">返回首页</a></p>
  <p class="small muted">「学生反馈」zip 里每份文档一个 .txt（结论 + 主要问题 + 修改建议），可直接发给对应学生；00_班级汇总.txt 含通过率与未通过名单。</p>
</div>"""
        self._send(HTTPStatus.OK, _page("批量检查结果", body))

    # -- 下载 --------------------------------------------------------------
    def _download(self, path: str) -> None:
        parts = path.split("/", 3)
        if len(parts) != 4:
            raise ValueError("下载路径不合法")
        _, _, token, filename = parts
        directory = _store_dir(token).resolve()
        target = (directory / unquote(filename)).resolve()
        if not target.is_relative_to(directory) or target == directory:
            raise ValueError("下载路径不合法")
        if not target.is_file():
            raise FileNotFoundError("文件不存在")
        content_type = {
            ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ".png": "image/png",
            ".pdf": "application/pdf",
            ".json": "application/json; charset=utf-8",
            ".csv": "text/csv; charset=utf-8",
        }.get(target.suffix.lower(), "text/plain; charset=utf-8")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(target.stat().st_size))
        self.send_header("Content-Disposition", f'attachment; filename="{target.name}"')
        self.end_headers()
        self.wfile.write(target.read_bytes())


# ---------------------------------------------------------------------------
# 启动
# ---------------------------------------------------------------------------
def create_server(host: str = "127.0.0.1", port: int = 0) -> ThreadingHTTPServer:
    """创建（未启动的）HTTP 服务；port=0 时由系统分配。"""
    server = ThreadingHTTPServer((host, port), TasWebHandler)
    server.daemon_threads = True
    return server


def serve(host: str = "127.0.0.1", port: int = 8765) -> None:
    server = create_server(host, port)
    actual_host, actual_port = server.server_address[:2]
    print(f"tas web 已启动：http://{actual_host}:{actual_port}")
    print("按 Ctrl+C 停止。所有文档只在本机处理。")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止。")
    finally:
        server.server_close()


def main(argv: Optional[list[str]] = None) -> int:
    import argparse

    parser = argparse.ArgumentParser(prog="tas web", description="tas Web 双入口（本地）")
    parser.add_argument("--host", default="127.0.0.1", help="监听地址（默认 127.0.0.1，仅本机）")
    parser.add_argument("--port", type=int, default=8765, help="端口（默认 8765）")
    args = parser.parse_args(argv)
    serve(args.host, args.port)
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
