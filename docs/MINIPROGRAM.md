# 文档接龙（微信小程序）设计方案

> 目标：把 tas 的格式检查能力做成**班里真正能用的"接龙"**——发起人建接龙，同学从微信聊天里选文档提交，一键审核全班文档是否规范。
> 状态：后端 API + 手机 H5（`/m`）已完成；微信小程序源码在 `miniprogram/`（需用微信开发者工具导入）。

## 1. 为什么是"接龙 + 小程序"

班级现状：老师/班委在群里发起接龙收文档，同学把 `.docx` 发到群里或私聊；文件命名乱、版本乱，**格式是否规范往往要人工一个个打开看**。

小程序的天然优势：

- 分享卡片一键进同一接龙（`onShareAppMessage`），不用加好友/建群机器人；
- `wx.chooseMessageFile` 可以直接**从微信聊天记录里选文档**，正贴合"群里交文件"的习惯；
- 提交后立刻检查、老师一键审核全班，问题清单当场可见；
- 无账号：昵称本地缓存，零注册成本（符合当前"先不做账号"）。

## 2. 端到端流程

```
发起人（任何同学/老师）
  └─ 小程序首页：填标题 / 选规范（profile）/ 说明 → 「发起接龙」→ 得到 6 位 code
       └─ 分享卡片到微信群
同学
  └─ 点卡片 → 进入接龙 → 「提交我的文档」→ 从聊天记录选 .docx → 提交
       └─ 后端立刻用当前规范检查 → 返回错误/警告数 + 状态
任何人（发起人/老师/同学）
  └─ 「一键审核」→ 后端逐份复查 → 通过率 + 错误/警告总数 + 共性问题 Top + 逐份状态
       └─ 点某份「查看报告」→ 每条问题（规则/位置/说明）
```

## 3. 架构（当前 = 单机/内网版）

```
微信小程序 / 手机 H5(/m) / curl
        │  JSON + multipart
        ▼
tas web 服务（Python 标准库 http.server，本机 127.0.0.1:8765）
        │
        ├─ /api/collections…        接龙业务（collections.py）
        │     ├─ data/collections.json         接龙与提交记录
        │     └─ data/collections/<code>/      提交的 .docx
        └─ 检查引擎（DocModel + roles + run_checks）
              └─ 与 CLI / Web 完全同一套规则（profile 可选）
```

- **不拆教师端/学生端**：任何人可以发起、提交、审核；页面只有"接龙"一个概念。
- **无账号**：昵称仅展示；接龙 code 即入口。
- **数据本地**：JSON + 文件；上线可替换为云开发数据库/对象存储（见 §7）。

## 4. HTTP API（已实现，测试覆盖）

| 方法 | 路径 | 说明 |
|---|---|---|
| GET | `/api/profiles` | 规范列表（小程序发起接龙时选择） |
| GET | `/api/collections` | 接龙列表（按创建时间倒序） |
| POST | `/api/collections` | 发起接龙：JSON `{title, owner, ruleset, note, due}` → `{ok, collection:{code,...}}` |
| GET | `/api/collections/{code}` | 接龙详情 + 提交列表 |
| POST | `/api/collections/{code}/submit` | multipart：`student` + `file`（.docx）→ 立即检查并返回计数 |
| POST | `/api/collections/{code}/review` | 一键审核：逐份复查 → `summary`（总数/通过率/错误警告汇总/共性问题 Top）+ `submissions` |
| GET | `/api/collections/{code}/submissions/{sid}/report` | 单份完整检查报告（`Report.to_dict` 结构） |

错误统一：HTTP 4xx + `{"ok": false, "error": "可读中文"}`。文件限制 `.docx`、单文件 ≤30MB（小程序端上传工具本身上限 10MB，见 §6）。

## 5. 微信小程序目录（`miniprogram/`）

```
miniprogram/
├─ project.config.json      # appid 用 touristappid 占位，导入时替换
├─ app.js / app.json / app.wxss / sitemap.json
├─ utils/api.js             # BASE_URL + request/uploadFile 封装
└─ pages/
   ├─ index/                # 接龙列表 + 发起接龙
   └─ collection/           # 接龙详情：提交 / 一键审核 / 报告
```

详见 `miniprogram/README.md`（导入步骤、真机联调、上线域名配置）。

## 6. 本地调试（微信开发者工具）

1. 启动后端：`python -m tas web --port 8765`（Windows 下先设 `PYTHONPATH`）。
2. 微信开发者工具 → 导入项目 → 选择 `D:\zdsh\work\tas\miniprogram`。
3. `project.config.json` 里 AppID 填自己的（没有就选"测试号"）。
4. 详情 → 本地设置 → 勾选「不校验合法域名、web-view（业务域名）、TLS 版本以及 HTTPS 证书」。
5. 模拟器：`utils/api.js` 的 `BASE_URL = "http://127.0.0.1:8765"` 即可。
6. 真机预览：手机与电脑同一 Wi-Fi，把 `BASE_URL` 改成电脑局域网 IP（`ipconfig` 查看），并放行 Windows 防火墙 8765 端口。
7. 文件来源：小程序不能读手机任意本地文件，只能 `wx.chooseMessageFile` 从**聊天记录/相册**选；建议同学先把文档发到"文件传输助手"或群里，再进小程序选。

## 7. 上线到微信需要补的（现实约束，面试可讲）

| 项 | 说明 |
|---|---|
| 小程序主体 | 个人/企业注册小程序，拿到 AppID；类目需与"教育/工具"匹配，涉及用户生成内容需隐私协议 |
| 域名 | 后端必须 **HTTPS + 已备案域名**；小程序后台配置 `request` 与 `uploadFile` 合法域名 |
| 文件大小 | `wx.uploadFile` 单文件上限 10MB；大文档需压缩或分片（当前 30MB 服务端限制仅本地） |
| 数据存储 | 从本地 JSON 换成云开发数据库/对象存储或自建 DB（PostgreSQL/SQLite）+ 定期备份 |
| 身份 | 当前无账号；上线建议 `wx.login` → openid 做提交人绑定（防冒名）、发起人管理员权限 |
| 通知 | 截止提醒/审核完成可用订阅消息（`wx.requestSubscribeMessage`） |
| 文档预览 | 报告里可 `wx.downloadFile` + `wx.openDocument` 直接打开学生文档（只读） |
| 并发/安全 | 多用户同时提交要加锁/队列；上传文件做类型/大小/内容校验，禁止可执行文件 |

## 8. 当前限制（如实说明）

- 单机/内网可用；**未做微信登录与权限模型**（任何人可审核，适合小班互评场景）。
- 数据存在 `data/` 本地目录（`.gitignore` 已忽略），多设备不共享；上线需换共享存储。
- 检查报告目前为 JSON/页面清单；小程序端只展示摘要，完整报告后续可导出 PDF/图片。
- 页面预览（Word 渲染）仅桌面 Web 有；小程序暂未接（可后续用后端渲染图 + `image` 展示）。
- 格式"修复"仍走桌面 Web/CLI；小程序当前聚焦"检查+审核"（符合接龙场景）。

## 10. 模拟器实测记录（2026-10-26）

环境：微信开发者工具 **2.02.0**（Windows）、已登录、项目已导入（AppID `wx9e3d7bdd9b6ae328`）、后端 `tas web` 本机 `127.0.0.1:8765`、自动化工具 `miniprogram-automator 0.13.0`（通过 `cli auto --auto-port 9420` 连接）。

已跑通并截图：

| 场景 | 结果 |
|---|---|
| 首页加载 | 4 个规范从 `/api/profiles` 加载成功；接龙列表显示 code/发起人/规范/提交数/截止 |
| 发起接龙 | 表单填写 → 创建成功（code `43CR9L`），列表刷新 |
| 接龙详情 | 真实 2 份提交（学生 庄子剑）显示状态「未通过」、错误/警告计数 |
| 一键审核（真实文档） | 2 份、通过率 **0%**、错误 620 / 警告 85 / 提示 15；共性问题 Top：BODY-001 235、BODY-002 192、HEAD-002 52… |
| 单份报告 | 文件名/规范/37/38 条规则/段落 134·表格 7·图片 1；按规则统计 + 问题明细（含位置、说明） |
| 提交展示 | 通过真实 API 提交一份合规文档 → 列表显示「通过」、错误 0 |
| 一键审核（合规文档） | 通过率 **100%**、错误/警告/提示全 0、未发现共性问题 |

截图（本机生成，PNG 已 gitignore）：`eval/mp_screenshots/`
`index.png`、`create_form.png`、`collection.png`、`review.png`、`report.png`、`submit_result.png`、`review_pass.png`。

**未能自动化验证**：`wx.chooseMessageFile` 是原生聊天文件选择器，模拟器里无法真正点开/mock 不可靠；因此「从聊天里选文件」按钮到上传成功这一步需**真机**验证（把 .docx 先发到「文件传输助手」再进小程序选择）。提交成功后到后端、列表刷新、审核、报告已在模拟器验证。

## 11. 真机网络问题排查（2026-10-27 新增）

真机提交报「网络连接失败 / 超时」时，按**从外到内**四层定位，不要盲改代码：

### 第 1 层：后端是否监听所有网卡

```powershell
# 必须是 0.0.0.0:8765；如果显示 127.0.0.1:8765，手机一定连不上
Get-NetTCPConnection -LocalPort 8765 -State Listen
# 本机健康检查
Invoke-WebRequest http://127.0.0.1:8765/health
```

启动命令：`python -m tas web --host 0.0.0.0 --port 8765 --verbose`
（`--verbose` 会打印每个请求：客户端 IP、方法、路径、状态码）

### 第 2 层：手机能否到电脑（最关键的一步）

**手机浏览器**打开：`http://<电脑局域网IP>:8765/api/diag`

| 结果 | 结论 |
|---|---|
| 返回 JSON，且 `"client"` 是手机 IP（192.168.x.x） | **网络层通了**，问题在小程序配置/上传环节 |
| 打不开/超时 | 网络层问题：不同 Wi-Fi、IP 错、防火墙、AP 隔离、iOS 本地网络权限 |
| 返回 404/JSON 报错 | 网络通但后端版本旧（重启后端） |

> `/api/diag` 返回服务端看到的客户端 IP、监听地址、时间、是否 verbose——一眼判断“请求到底有没有到电脑”。

### 第 3 层：小程序内「网络自检」

首页底部新增 **「网络自检 · 手机连不上后端时点这里」**：

- 成功弹窗：后端地址、耗时、手机网络类型、电脑看到的手机 IP、服务端监听地址（点「修改地址」可直接改后端 IP，存本机，不用改代码）
- 失败弹窗：当前 BASE_URL + 具体错误 + 检查清单 + 「修改地址」入口

> 项目里 `DEFAULT_BASE_URL` 默认填的是当前电脑局域网 IP；换网络（IP 变了）时在小程序里点「修改地址」即可，无需重新改代码/编译。

同时可用手机浏览器打开 `http://<电脑IP>:8765/m`（H5 版）做对照：
- **H5 能提交、小程序不能** → 小程序侧问题（BASE_URL 没改成电脑 IP / 域名校验 / 真机预览模式限制）
- **H5 也不能** → 后端或网络问题

### 第 4 层：后端请求日志（判断请求有没有到达）

```powershell
python -m tas web --host 0.0.0.0 --port 8765 --verbose
# 手机点提交后，看终端是否出现：
# [tas] 192.168.x.x POST /api/collections/XXXXXX/submit -> 200 (...)
# [tas] upload from 192.168.x.x: student='张三' filename='xx.docx' size=...
```

- **完全没有日志** → 请求没到服务器：网络/防火墙/域名校验/BASE_URL 问题
- **有日志但 4xx/5xx** → 网络 OK，看后端返回的错误（文件格式、大小、接龙 code）

### 常见错误对照表

| 现象/报错 | 原因 | 处理 |
|---|---|---|
| `request:fail timeout` | 请求超时，包到不了电脑 | 检查同 Wi-Fi、防火墙、电脑 IP；手机浏览器先测 `/api/diag` |
| `ERR_CONNECTION_REFUSED`（-102） | 端口没监听/后端没启动/绑了 127.0.0.1 | `--host 0.0.0.0` 重启；`Get-NetTCPConnection` 确认 |
| `ERR_CONNECTION_TIMED_OUT` | 防火墙丢包 / 不同网段 / 路由隔离 | 放行 8765、检查 IP 网段；校园网常开启设备隔离 |
| `ERR_NAME_NOT_RESOLVED` | BASE_URL 写了域名但解析不了 | 改回电脑局域网 IP |
| `url not in domain list` / 域名校验 | 真机预览仍在校验合法域名 | 用**真机调试**（开发版）、或后端上 HTTPS 域名并在小程序后台配置 |
| `uploadFile:fail` / 上传中途失败 | 文件 >10MB、网络抖动、后端 30MB 限制 | 压缩文档；重试；看后端日志 |
| iOS 手机所有局域网请求都失败 | 微信没有「本地网络」权限 | iPhone 设置 → 微信 → 本地网络 → 打开，重启微信 |
| 手机浏览器也打不开 | 最常见：手机用 4G/校园访客网/设备隔离 | 手机连电脑同一 Wi-Fi；或电脑连手机热点；或内网穿透/公网部署 |

### 其它排查命令

```powershell
# 电脑局域网 IP（找 WLAN/以太网 IPv4，192.168.x.x 或 10.x.x.x）
ipconfig
# 从另一台设备测端口（PowerShell）
Test-NetConnection -ComputerName <电脑IP> -Port 8765
# 防火墙放行（需管理员）
New-NetFirewallRule -DisplayName "tas-8765" -Direction Inbound -Protocol TCP -LocalPort 8765 -Action Allow
```

## 12. 体验入口（不用微信开发者工具）

后端已内置手机 H5 内测页：启动服务后浏览器/手机访问 **`http://127.0.0.1:8765/m`**，即可完整走"发起接龙 → 提交 .docx → 一键审核 → 查看报告"。接口与小程序完全一致，可先用它验证流程，再进微信开发者工具。
