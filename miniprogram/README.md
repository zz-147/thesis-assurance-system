# 文档接龙小程序（tas · 微信小程序前端）

tas 后端（`D:\zdsh\work\tas`）的微信小程序前端：**单一界面「文档接龙」**，不分学生端/教师端。

- 任何人可**发起接龙**（选检查规范、填标题/昵称/说明/截止时间），后端生成 6 位接龙码；
- 同学从**分享卡片/群聊**进入同一接龙，从微信聊天记录里选 `.docx` 提交（无需注册登录，昵称存本机）；
- 任何人可点**一键审核**，后端批量检查全部文档并返回通过率、错误/警告总数、共性问题 Top、每份状态；
- 点某份提交记录可查看该份文档的**报告摘要**（问题最多的规则 + 明细）。

原生 WXML / WXSS / JS，无第三方依赖、无 npm 构建。

---

## 1. 目录结构

```
miniprogram/
├─ app.js                     # 入口；昵称本地缓存 + ensureNickname（wx.showModal 输入）
├─ app.json                    # 页面注册、导航栏、window 配置
├─ app.wxss                    # 全局基础样式（卡片/按钮/标签/状态块）
├─ project.config.json         # 开发者工具项目配置（appid 占位 touristappid，需替换）
├─ sitemap.json                # 页面被微信索引规则
├─ utils/api.js                # 后端接口封装：BASE_URL + request/uploadFile + 全部 API
└─ pages/
   ├─ index/                   # 首页：接龙列表 + 发起接龙表单
   │  ├─ index.json / index.wxml / index.wxss / index.js
   └─ collection/              # 接龙详情：提交文档 + 一键审核 + 报告摘要 + 分享
      ├─ collection.json / collection.wxml / collection.wxss / collection.js
```

## 2. 先启动后端

在 tas 项目根目录（PowerShell，Python 路径按你本机实际环境）：

```powershell
$env:PYTHONPATH = "D:\zdsh\work\tas\src"
Set-Location D:\zdsh\work\tas
python -m tas web --port 8765
```

浏览器打开 <http://127.0.0.1:8765/api/profiles> 能看到 JSON（规范列表）就说明后端正常。
文档接龙接口为 `/api/collections` 系列（契约见 `utils/api.js` 注释与本文第 5 节）。

> 真机预览时后端需要监听所有网卡：`python -m tas web --host 0.0.0.0 --port 8765`，
> 并允许 Windows 防火墙放行 8765 端口。

## 3. 用微信开发者工具导入

1. 打开 **微信开发者工具** → 「导入项目」；
2. 目录选择 **`D:\zdsh\work\tas\miniprogram`**（注意：选的是 miniprogram 子目录，不是 tas 根目录）；
3. **AppID**：
   - 项目 `project.config.json` 当前 AppID 为 **`wx9e3d7bdd9b6ae328`**（已用该账号在微信开发者工具 2.02.0 中导入并跑通模拟器）；
   - 换账号/发布时改成自己的 AppID；没有正式 AppID 可在导入弹窗里点「测试号」获取（测试号不支持云开发）。
   - `touristappid` 是占位值，新版开发者工具会报「不存在此 AppID」。
4. 打开「详情 → 本地设置」，勾选：
   **「不校验合法域名、web-view（业务域名）、TLS 版本以及 HTTPS 证书」**
   （`project.config.json` 里已预设 `"setting": { "urlCheck": false }`，若仍报域名错误，手动再勾一次）。
5. 模拟器里默认请求 `http://127.0.0.1:8765`（见 `utils/api.js` 的 `BASE_URL`），后端在本机启动即可直接联调。
6. 建议把基础库设为 **2.17.1 及以上**（昵称输入用了 `wx.showModal` 的 `editable`；工程默认 `libVersion: 3.0.0`）。

### 模拟器里怎么走一遍

1. 首页「＋ 填写」→ 填标题/昵称 → 选规范 → 「创建接龙」→ 弹窗显示 6 位接龙码；
2. 点「复制并查看」进入详情页 → 「从聊天里选 .docx 提交」（开发者工具对 `wx.chooseMessageFile` 支持有限，选不到聊天文件时用真机验证）；
3. 提交成功后点「一键审核」→ 看通过率/错误警告/共性问题 Top/每份状态；
4. 点提交记录某一行 → 页面底部出现「报告摘要」；
5. 右上「···」或详情页「分享给同学」→ 卡片路径带 `?code=XXXXXX`，别人点开直接进同一接龙。

## 4. 真机预览（手机与电脑同一局域网）

1. 查电脑局域网 IPv4：`ipconfig`，例如 `192.168.1.8`（项目里默认已填 `http://10.127.173.183:8765`，换成你的电脑 IP）；
2. **换网络/IP 变了不用改代码**：小程序首页底部「网络自检」→「修改地址」→ 输入 `http://电脑IP:8765` 保存到本机；
   也可以直接改 `utils/api.js` 顶部的 `DEFAULT_BASE_URL`（真机调试必须用电脑局域网 IP，不能用 `127.0.0.1`）；
3. 后端用 `--host 0.0.0.0 --port 8765` 启动，Windows 防火墙放行 TCP 8765；
4. 手机和电脑连**同一个 Wi-Fi**（不要用手机热点隔离网络），手机浏览器先打开
   `http://192.168.1.8:8765/api/profiles` 验证能通；
5. 开发者工具「预览」扫码 → 真机：在**微信聊天里**打开一个 `.docx` 文件消息 → 小程序里点「从聊天里选 .docx 提交」即可选到；
6. 真机仍报「网络连接失败」：依次检查 IP 是否写对、后端是否 `--host 0.0.0.0`、防火墙、手机是否同一网段（部分校园网/访客 Wi-Fi 会隔离设备）。**最快的判断方法**：
   - 手机浏览器打开 `http://电脑IP:8765/api/diag`：能看到 JSON 且 `client` 是你手机 IP → 网络通；打不开 → 网络层问题；
   - 小程序首页底部点 **「网络自检」**：会显示 BASE_URL、耗时、手机网络类型、后端看到的手机 IP；失败会给检查清单；
   - 后端用 `python -m tas web --host 0.0.0.0 --port 8765 --verbose` 启动，手机点提交后看终端有没有请求日志——没有日志 = 请求没到电脑。
   完整排查步骤见 `docs/MINIPROGRAM.md` §11。

## 5. 页面 / 接口对应关系

`utils/api.js` 里 `BASE_URL` 为后端根地址（默认 `http://127.0.0.1:8765`，末尾不要带 `/`）。

| 页面 | 交互 | 接口 | 关键字段 |
|---|---|---|---|
| 首页 | 加载接龙列表 | `GET /api/collections` | `collections[].code/title/owner/ruleset/note/due/created_at/submission_count` |
| 首页 | 规范 picker | `GET /api/profiles` | `default`、`profiles[].id/name/discipline` |
| 首页 | 发起接龙 | `POST /api/collections`（JSON `{title, owner, ruleset, note, due}`） | 返回 `collection.code` |
| 详情页 | 接龙信息 + 提交列表 | `GET /api/collections/{code}` | `collection`、`submissions[].id/student/filename/created_at/counts{error,warning,info}/status` |
| 详情页 | 提交文档 | `POST /api/collections/{code}/submit`（multipart：`student` 字段 + `file` 文件） | 返回 `submission`（含 `counts`） |
| 详情页 | 一键审核 | `POST /api/collections/{code}/review` | `summary.total/checked/failed/pass_rate/counts/top_rules`、`submissions[]` |
| 详情页 | 查看某份报告 | `GET /api/collections/{code}/submissions/{sid}/report` | `Report.to_dict`：`counts`、`rules{total,checked}`、`issues_by_rule`、`issues[]`、`stats`、`file`、`profile_label` |
| 详情页 | 分享 | `onShareAppMessage` | `path: /pages/collection/collection?code={code}` |
| 全部页面 | 失败提示 | 后端 4xx/5xx `{"ok":false,"error":"中文"}` | 直接展示 `error` 文案 |

- 所有请求失败都会弹**中文提示**（`utils/api.js` 把后端 `error`、HTTP 状态码、网络/超时分别翻译成中文）；
- 提交/审核/创建按钮都有 `loading` + `disabled` + 进行中标记，**防重复提交**；
- 列表/提交记录/报告明细都有**空状态**与滚动区域；
- 昵称：首次发起/提交时用 `wx.showModal`（可编辑输入框）获取，`wx.setStorageSync('nickname')` 存本机，不涉及账号登录。

## 6. 文件限制与使用边界（重要）

- **单文件上限 10MB**：这是小程序 `wx.uploadFile` 的限制，代码在选文件后按 `api.MAX_UPLOAD_BYTES` 先校验，超限直接提示压缩/删减；
- **只支持 `.docx`**：`.doc`/`.pdf` 等先另存为 `.docx`（后端也只收 `.docx`）；
- **只能从聊天记录/聊天文件里选**：小程序拿不到手机任意本地目录的读取权限，`wx.chooseMessageFile` 就从微信聊天文件里选 —— 正好符合「群里接龙传文档」的使用习惯（先把文档发到群里或文件传输助手，再进小程序选）；
- 文档只上传到你自己的 tas 后端（本机/自建服务器），小程序端不做任何内容解析；
- 无账号体系：谁都能发起、谁都能审核（符合当前需求）；昵称仅本机缓存，可随时在表单里改名。

## 7. 上线注意

1. **必须 HTTPS**：把 `utils/api.js` 的 `BASE_URL` 改成 `https://你的域名`（小程序不允许请求 `http://` 或裸 IP）；
2. 在**小程序后台 → 开发管理 → 开发设置 → 服务器域名**里，把该域名同时加到：
   - **request 合法域名**
   - **uploadFile 合法域名**
3. 域名需已备案且证书有效（TLS 1.2+）；后端部署到公网服务器并开放对应端口；
4. `project.config.json` 的 `appid` 换成正式 AppID，上传体验版 → 提交审核 → 发布。

## 8. 常见问题

| 现象 | 原因 / 处理 |
|---|---|
| 打开首页就提示「网络连接失败」 | 后端没启动 / `BASE_URL` 不对；模拟器用 `127.0.0.1:8765`，真机用电脑局域网 IP |
| 提示「接口不存在（HTTP 404）」 | 后端版本旧、没有 `/api/collections` 接口；确认启动的是新版 tas 后端 |
| 提示「不校验合法域名」相关错误 | 「详情 → 本地设置」勾选「不校验合法域名…」；或改用 HTTPS 域名 |
| 选文件时没有聊天文件 | 先把 `.docx` 发到微信聊天/文件传输助手；开发者工具能力有限，推荐真机验证 |
| 提交报「只支持 .docx」 | 文件是 `.doc` 或扩展名被改过；用 Word 另存为 `.docx` |
| 提交报「文件太大」 | 超过 10MB；压缩图片/删减内容，或让后端放宽限制后同步调整 `api.MAX_UPLOAD_BYTES` |
| 真机分享打开后空白/缺 code | 分享路径必须带 `?code=`；从详情页右上角分享或「分享给同学」按钮分享，不要复制页面链接 |
| 一键审核没结果 | 还没有人提交文档；先让同学提交再审核 |

## 9. 验收自查（静态）

本目录不需要 npm 安装，用 Node.js 校验所有 `.json` 可解析：

```powershell
Set-Location D:\zdsh\work\tas\miniprogram
Get-ChildItem -Recurse -Filter *.json | ForEach-Object {
  node -e "JSON.parse(require('fs').readFileSync(process.argv[1],'utf8'));console.log('OK',process.argv[1])" $_.FullName
}
```

- [x] 所有 `.json` 可被 `JSON.parse` 解析（`project.config.json` / `app.json` / `sitemap.json` / 两个页面 `.json`）；
- [x] 页面 JS 无第三方依赖、无 npm 包；`node --check` 通过；
- [x] WXML 标签闭合、`bind*` 事件在对应页面 JS 中均有实现；
- [x] 请求路径/字段与第 5 节接口契约一一对应（字段只来自契约，未自创）；
- [x] 加载中 / 失败中文提示 / 空状态 / 按钮加载与防重复均有实现；
- [x] 已用本地真实 tas 后端（临时端口 + 临时数据目录）联调验证全链路：
  创建接龙（JSON）→ 提交 .docx（multipart `student` + `file`）→ 一键审核（`summary.pass_rate` 0.5 → 页面 50%、
  `top_rules` 显示）→ 查看某份报告（`{ok, report}` 包装已兼容，`Report.to_dict` 字段解析正常）→ 错误场景（非法
  code/非 .docx/空标题）均返回 HTTP 4xx + `{ok:false,error:"中文"}`，页面按中文提示展示。
