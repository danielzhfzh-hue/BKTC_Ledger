# BKTC Ledger 工程交接（v1.2.0）

## 当前架构

```text
pywebview UI
  ├─ 六表编辑 / 新建订单 / 筛选 / 跨表查询
  ├─ 保存 ───────────────→ SQLite .db（权威源）
  ├─ 生成台账 ───────────→ xlsx（生成产物）
  └─ 导出编辑副本 → Excel/WPS → 差异预览 → 二次确认 → SQLite
```

- `app.py`：桌面壳、文件选择、本地 API、更新下载、两阶段导入令牌。
- `database.py`：关系表、JSON 迁移、事务、备份、修订、Excel 往返。
- `core.py`：六表 schema、规范化、派生、业务校验、查询导出与台账入口。
- `build_ledger_main.py`：导航页、未回收管理表、JOB 页生成。
- `ui/`：纯 HTML/CSS/JS；大表行虚拟化，跨表查询包含「未付款订单」和「未回款明细」派生数据源。
- `tests/`：SQLite、回导冲突和生成引擎边界回归。

真实业务数据仍位于仓库外：

- 旧源：`订单整理/BKTC上海POU营业管理表.records.json`（迁移后只作保留/兼容）。
- 新权威源：`订单整理/BKTC上海POU营业管理表.db`。
- 生成产物：`订单整理/BKTC上海POU营业管理表.xlsx`。

`.db*`、JSON、xlsx 与备份均被 `.gitignore` 排除。

## SQLite 约束与保存

- 合同 JOB 唯一；五个子表通过 JOB 外键关联合同。
- 发货批次以 `(JOB No, 发货批次)` 唯一；设备通过同一组合外键关联批次。
- 保存前运行 `core.derive` 与 `core.validate`；error 会拒绝写入。
- 写入采用 `BEGIN IMMEDIATE` 单事务；提交前执行 `foreign_key_check`。
- `PRAGMA journal_mode=WAL / synchronous=FULL / foreign_keys=ON`。
- 每次写入前用 SQLite backup API 在 `备份/` 留最近 10 个数据库快照。
- `_meta.revision` 每次成功写入 +1；Excel 编辑副本绑定 `database_id + base_revision`。
- 每行业务 schema 外的字段保存在 `_extra_json`，用于无损保留飞书/钉钉关联与公式字段。

## Excel 回导协议

`database.export_editable_workbook` 导出六个业务 sheet：

- 隐藏 `_记录ID` 列；新增行允许 ID 为空。
- 自动派生字段不导出，导入后统一重算。
- `_BKTC_META` veryHidden sheet 保存格式版本、数据库 ID、修订和 hash。
- 普通文本以 `= + - @` 开头时转义；导入拒绝公式单元格。

`prepare_editable_import` 生成新增/修改/删除逐字段差异；`Api` 只保留最后一个随机 token。`apply_editable_import` 在提交前再次校验数据库修订、database_id、内容 SHA-256 和 Excel size/mtime，成功后走正常事务保存与备份。

## 关键交互不变量

- 切表必须清空 `selection`/`anchor`，避免索引跨表复用。
- 合同删除由 `tryCascadeDelete` 独占；取消任一次确认不得退回普通删除。
- 所有新行/复制行/拆分批次必须生成新 `记录ID`；后端仍会补齐缺失或重复 ID。
- 修改合同 JOB No 时同步五个子表；修改发货批次名时同步设备和覆盖批次。
- 切换数据库前处理未保存状态；窗口关闭使用 `beforeunload` 保护。

## 验证

```bash
cd /Users/danielzhu/projects/订单整理/BKTC_Ledger
.venv/bin/python -m unittest discover -s tests -v
.venv/bin/python -m py_compile app.py core.py database.py build_ledger_main.py
node --check ui/app.js
```

真实数据安全演练应以 JSON 为只读源，在 `/tmp` 创建数据库与 xlsx；禁止把业务数据写入仓库。已验证 40/87/615/134/102/112 六表数量、全部记录 ID、未知字段、外键、完整性、0 差异往返和 42 sheet 生成。

## 打包与发布

`.github/workflows/build.yml` 在 `windows-latest` / `macos-latest` 上安装依赖、运行 unittest，再用 PyInstaller onedir 打包。push 到主分支触发构建；tag `v*` 还会创建 Release。

发布流程使用 `github:yeet`：从 `agent/*` 分支提交并推送，创建 draft PR。版本升级时再由维护者决定是否打 `v1.2.0` tag；本任务不自动创建 Release tag。

## 外部同步

父目录的 `feishu_upsert.py` / `dingtalk_push.py` 仍以旧 JSON 路径为默认值，不属于本仓库。若后续继续使用，应把参数指向数据库并通过 `core.load_store(.db)` 读取，或先显式导出兼容 JSON；不要把数据库内容提交到 GitHub。
