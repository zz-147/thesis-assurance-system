# TAS 老师端演示资料（2026-10-30）

## 先看这个
- **演示报告_一页总结.pdf / .docx**：一页看懂价值、指标、界面
- **演示报告_完整版.pdf / .docx**：完整流程 + 截图 + 产物说明（4 页）
- **TAS_30秒演示脚本.md**：面试/发老师时可直接照念的 30 秒流程
- **简历摘要.md**：可直接改成简历 bullet 的版本

## 演示数据（生成论文，非真实学生）
- `名单.txt`：8 人名单（张三/李四/王五/赵六/钱七/孙八/周九/吴十）
- `提交/`：6 份 docx（含李四重复提交、赵六文件名无学号；王五/周九/吴十缺交）
- `报告/`：完整批处理产物
  - `batch.xlsx` / `batch.csv` / `batch.html`：班级报告（含缩略图列、条形图）
  - `feedback.zip`：每份学生反馈（可直接转发）+ 班级汇总
  - `roster-report.txt`：缺交/重复/命名核对
  - `thumbnails.zip`：6 份首页缩略图
  - `refs-report.md` / `.json`：文献真实性核验（含证据 URL 与"人工核实"提示）
  - `反馈示例.html`：反馈效果预览
- `截图/`：首页、批量报告、反馈示例

## 复现命令（项目根目录）
```powershell
& $PY -m tas web --port 8765   # 或双击 启动老师端.bat
& $PY -m tas batch demo/提交 --out-dir demo/报告 --fix --roster demo/名单.txt --preview --refs --refs-limit 2
```

## 说明
- 演示论文来自生成器（`thesis-format/out`），不含真实学生隐私。
- 缩略图/页面预览需要 Windows + Microsoft Word。
- 文献核验：not_found 仅表示未在当前索引检索到，中文期刊可能未被收录，不等于造假。
