# 上海康肯销售订单管理系统工程交接（v1.6.0）

## 当前架构

```text
pywebview UI
  ├─ 六表编辑 / 新建订单 / 筛选 / 跨表查询 / 报价单
  ├─ 保存 ───────────────→ SQLite .db（权威源）
  ├─ 导出台账 ───────────→ 同目录同名 xlsx（只读生成产物）
  └─ 导出报价 ───────────→ Downloads/报价单/*.xlsx（来自已保存报价）
```

- `app.py`：桌面壳、数据库文件选择、本地 API 与更新下载。
- `database.py`：订单及报价关系表、JSON 迁移、事务、备份、修订与字段级审计。
- `quotation.py`：依据两份业务示范抽取的通用 9 列报价版式与公式输出。
- `core.py`：六表 schema、规范化、派生、业务校验、查询导出与台账入口。
- `build_ledger_main.py`：导航页、未回收管理表、JOB 页生成。
- `ui/`：纯 HTML/CSS/JS；六个业务标签页均有独立录入窗口，大表行虚拟化，跨表查询包含「未付款订单」和「未回款明细」派生数据源。
- `tests/`：SQLite、桌面 API 和生成引擎边界回归。

真实业务数据统一位于项目的非 Git 跟踪目录：

- 权威源：`BKTC_Ledger/data/BKTC_Ledger.db`。
- 生成产物：`BKTC_Ledger/data/BKTC_Ledger.xlsx`。
- 旧 JSON：只在 `data/legacy/` 保留兼容副本，不参与日常运行。

`.db*`、JSON、xlsx 与备份均被 `.gitignore` 排除。

## SQLite 约束与保存

- 合同 JOB 唯一；五个子表通过 JOB 外键关联合同。
- 发货批次以 `(JOB No, 发货批次)` 唯一；设备通过同一组合外键关联批次。
- 保存前运行 `core.derive` 与 `core.validate`；error 会拒绝写入。
- 写入采用 `BEGIN IMMEDIATE` 单事务；提交前执行 `foreign_key_check`。
- `PRAGMA journal_mode=WAL / synchronous=FULL / foreign_keys=ON`。
- 每次写入前用 SQLite backup API 在 `备份/` 留最近 10 个数据库快照。
- `_meta.revision` 每次成功写入 +1；桌面 API 使用修订号阻止旧窗口覆盖新版。
- 每行业务 schema 外的字段保存在 `_extra_json`，用于无损保留飞书/钉钉关联与公式字段。
- `quotation` 与 `quotation_item` 一对多保存报价头和明细；明细删除/更新随报价事务处理，并进入同一审计哈希链。

## Excel 输出约束

桌面应用只持久化数据库路径。台账路径由数据库路径确定：`example.db` 对应同目录下的 `example.xlsx`。Excel 只由数据库生成，不提供外部编辑回导入口。

## 审计模型

- `_audit_event`：一次成功保存一个事件，记录修订、时间、操作人、来源、动作、应用/平台/设备和哈希。
- `_audit_change`：按稳定 `记录ID` 保存新增/删除快照或修改字段的旧值/新值；自动派生字段过滤，关联同步与系统重算单独标记。
- 旧库首次打开时在 `_meta.audit_start_revision/audit_started_at` 建立基线，不为既有记录补写新增事件；`json_migration`、`database_created`、`portable_starter` 同样只建立基线。
- 审计写入与六表替换、外键检查、修订更新处于同一个 `BEGIN IMMEDIATE` 事务；失败整体回滚。
- 四个 SQLite trigger 阻止普通 UPDATE/DELETE 审计表，`audit_chain_head` 与连续 SHA-256 哈希用于检测直接篡改或截断。

## 关键交互不变量

- 切表必须清空 `selection`/`anchor`，避免索引跨表复用。
- 合同删除由 `tryCascadeDelete` 独占；取消任一次确认不得退回普通删除。
- 所有新行/复制行/拆分批次必须生成新 `记录ID`；后端仍会补齐缺失或重复 ID。
- 合同客户自动映射到五个子表；客户与 JOB 筛选在六表间保持，并可交叉筛选。
- 修改合同 JOB No、发货批次名、设备番号或付款条款款类时按稳定记录 ID 同步关联表。
- 开票/回款覆盖使用按批次分组的设备多选；同设备范围的多笔开票/回款金额累计计算。
- 付款条件按 JOB 整组维护且比例必须为 100%；开票/回款款类只能来自同一 JOB 的付款条件。
- 设备录入只能引用同一 JOB 的既有批次；开票/回款批次和制造番号选择器只展示同一 JOB 数据，逐台选择会回写批次并按实际设备行计算覆盖台数。
- 设置页只保存数据库路径，启动参数与环境变量仍可覆盖；XLSX 输出路径自动派生。
- 切换数据库前处理未保存状态；窗口关闭使用 `beforeunload` 保护。
- 报价编辑使用独立未保存状态；导出前必须先保存。来源 JOB 只从已保存数据库带入，避免把订单草稿误当权威数据。
- 历史价格只按去空格、忽略大小写后的客户与型号精确匹配；当前报价不进入自己的历史结果。

## 验证

```bash
cd /Users/danielzhu/projects/订单整理/BKTC_Ledger
.venv/bin/python -m unittest discover -s tests -v  # 当前 65 项
.venv/bin/python -m py_compile app.py core.py database.py quotation.py build_ledger_main.py create_portable_starter.py
node --check ui/app.js
```

真实数据安全演练应先用 SQLite backup API 复制到 `/tmp`，再在副本上保存；禁止把业务数据提交到 Git。当前权威库已验证 `integrity_check = ok`、外键检查无异常、schema v4，并可生成 42 sheet 台账。

## 打包与发布

`.github/workflows/build.yml` 在 `windows-latest` / `macos-latest` 上安装依赖、运行 unittest，再用 PyInstaller onedir 打包。PR 与 push 到主分支触发构建；产物内含空白启动 DB/XLSX，tag `v*` 还会创建 Release（Windows zip、macOS tar.gz）。

推送 `main` 会触发双平台构建；打 `v1.6.0` tag 会由工作流创建 GitHub Release，并上传 Windows zip 与 macOS tar.gz。公开资产只含空白启动数据；本机交付包另行注入 `data/` 当前副本。

## 外部同步

父目录的 `feishu_upsert.py` / `dingtalk_push.py` 仍以旧 JSON 路径为默认值，不属于本仓库。若后续继续使用，应把参数指向数据库并通过 `core.load_store(.db)` 读取，或先显式导出兼容 JSON；不要把数据库内容提交到 GitHub。
