# BKTC 台账维护工具 — 交接文档

更新：2026-08-04　·　维护者入口：先读本文，再看 [README.md](README.md)（用户向）

## 0. 一句话现状

`BKTC_Ledger`（pywebview 桌面应用，**v1.1.9**）是**当前唯一的编辑入口**；`BKTC上海POU营业管理表.records.json` 是**权威数据源**；`build_ledger_main.py` 是**生成引擎**（已并入工具,自包含）；`BKTC上海POU营业管理表.xlsx` 是**生成产物**（42 sheet）。**打包走 GitHub Actions**(win+mac CI);飞书/钉钉已同步(见 §6)。

> v1.1.5–v1.1.9 新增：任意六表**筛选+排序**、**单表导出 Excel**(📥)、**跨表查询导出**(🔍 report builder，按客户等任意关联字段筛+选字段+导出)、发货批次**拆分**(⇲)、**3 处派生字段自动同步**(付款条件文本/应收回款日/验收状态，见 §8)。

> 旧文档 `订单整理/HANDOFF.md` 描述「飞书 pipeline 为权威源」的旧架构,已过时——以本文为准(其 §三数据口径修复/脏数据清单仍可查)。

## 1. 架构与数据流

```
┌─ BKTC_Ledger (pywebview) ─┐   六标签页编辑 → state.data
│  core.py  数据层           │
│  app.py   桌面壳 + Api     │   APP_DIR 支持 sys._MEIPASS(打包冻结态)
│  build_ledger_main.py 引擎 │   (已并入,自包含)
│  ui/      HTML/JS/CSS      │
└─────────────┬──────────────┘
              │ 保存数据         │ 生成台账
              ▼                 ▼
   records.json (权威源)   build_ledger_main.build_from_data
   六表 + 预警规则           ↓
   覆盖前 → .bak           备份 xlsx → 重写 42 sheet
              ↑                 ↓
              └── (仅 records.json ↔ xlsx;「从 Excel 同步」已于 v1.1 移除,records.json 为唯一权威源)
```

- **保存** `save_store` → `derive`(normalize+派生) → 覆盖 records.json(覆盖前 `shutil.copy2` → `.bak`)
- **生成** `generate_xlsx` → `derive`+`validate` → `build_from_data` → 备份 xlsx → 重写 导航页/未回收管理表/40 JOB 页
- **派生同步(显示)**：JS `recomputeDerived()` 镜像 `core.derive` 的派生字段(发货批次合计/合同总台数·型号·付款条件文本/验收状态/应收回款日/覆盖台数),编辑后实时刷新,无需保存重启。

## 2. 关键文件与位置

> 工具根目录 = **git 仓库** `~/projects/订单整理/BKTC_Ledger/`(2026-08-04 从 `~/Documents/销售订单管理多维表格构筑/台账维护工具/` 迁入并自包含)。**GitHub 私有仓**:https://github.com/danielzhfzh-hue/BKTC_Ledger(只有源码,无数据)。

| 文件 | 位置 | 角色 |
|---|---|---|
| `app.py` | `BKTC_Ledger/` | pywebview 壳 + `Api`;`APP_DIR` 已支持 `sys._MEIPASS` |
| `core.py` | `BKTC_Ledger/` | 数据层:`SCHEMA`/`DERIVED`/`normalize`/`derive`/`validate`/`save_store`/`generate_xlsx`/`backup_xlsx`/`export_filtered`(「从 Excel 同步」相关 load_from_workbook/merge_workbook/reconcile 已于 v1.1 移除) |
| `build_ledger_main.py` | `BKTC_Ledger/`(已并入,自包含) | **生成引擎**:`build_from_data`/`write_job_sheet`/`compute_unpaid_rows`/`build_unpaid_sheet`/`verify`;core.py 直接 `from build_ledger_main import`。⚠️ 父目录 `~/Documents/销售订单管理多维表格构筑/` 另有 **legacy 副本**(供旧脚本),**改引擎只改工具内这份** |
| `ui/{index.html,app.js,style.css}` | `BKTC_Ledger/ui/` | 前端(行虚拟化、多行粘贴、批次重命名) |
| `.github/workflows/build.yml` | `BKTC_Ledger/` | **GitHub Actions 打包**(win+mac 矩阵,见 §4) |
| `.gitignore` | `BKTC_Ledger/` | 排除 `.venv/ dist/ build/ __pycache__/ *.spec` |
| `build_win.bat` / `README_win.md` | `BKTC_Ledger/` | Windows 本地构建(纯 ASCII)+ 说明;**备选**,主路径是 GitHub CI |
| `BKTC上海POU营业管理表.records.json` | `订单整理/`(工具外) | **权威源**(六表+预警规则;含 Excel 不展示的付款条件/覆盖串/应收日) |
| `BKTC上海POU营业管理表.xlsx` | `订单整理/` | 生成产物(导航页+未回收管理表+40 JOB 页) |
| `*.records.json.bak` | `订单整理/` | records.json 上次保存快照(自动滚动备份) |
| `feishu_upsert.py` / `dingtalk_push.py` | `订单整理/` | records.json → 飞书 upsert(dry-run/--apply)/ 钉钉 aitable 全量镜像;sys.path 指向 `BKTC_Ledger/` |
| `dist/BKTC_Ledger.app` | `BKTC_Ledger/dist/` | Mac 本地打包产物(ad-hoc 签名);CI 另有 macOS 产物 |

## 3. 改动历史

### 2026-08-03(工具修复,已验证)
| # | 文件:行 | 改动 | 为什么 |
|---|---|---|---|
| 1 容差口径 | `build_ledger_main.py:204,556,912` | `write_job_sheet` 加 `rules`;行预警 `+0.01`→`+{tol:g}`;`build_from_data` 透传 | 金额容差原本只对未回收表生效,JOB 页/导航页写死 0.01 |
| 2 records 备份 | `core.py:268` | `save_store` 覆盖前 `.bak` | 权威源被覆盖无 undo |
| 3 日期解析 | `core.py:97`+`validate` | `_norm` DATE 放宽 `[-/.年]`+单数字;失败保留原值;validate 加格式 warning | 粘贴非标准日期被静默清空 |
| 4 数字清空 | `core.py:88`+`app.js:270` | NUMBER/INT 保留 `null`;UI 清空存 null | 清空单价→0→设备变免费 |
| 5 粘贴带 JOB | `app.js:449` | `pasteGrid` 新行带 `state.job` | 与 add 按钮不一致 |

### 2026-08-04(迁移 + 打包)
- 工具从 `~/Documents/销售订单管理多维表格构筑/台账维护工具/` 迁到 `~/projects/订单整理/BKTC_Ledger/`;`build_ledger_main.py` 并入(自包含);`core.py` 去父目录 sys.path,`app.py` 加 `sys._MEIPASS`。
- **改名 BKTC_Ledger(ASCII)**:中文名「台账维护工具」拷到 Windows 会乱码——文件夹/程序名/.bat 全改 ASCII 根治。
- **GitHub Actions 打包**(主路径,见 §4):win+mac CI 全绿;Windows 产物 247 文件/30M,macOS 51M。
- `build_win.bat` 改纯 ASCII(中文 echo + `chcp 65001` 在中文 Windows cmd 解析崩 → 闪退)。
- 数据迁 Windows 只需 `ledger.records.json` 一个文件(ASCII 名,见 §4)。

### 2026-08-04 v1.1（功能升级）
- **预警灯跨平台**:导航页「预警」列原用 emoji 🔴🟡🟢(Mac 彩色、Windows 单色/方框),改纯文本「有逾期/部分未回·关注/已回款/无开票」+ 条件格式底色(红 FBE2E2 / 黄 FFF2CC / 绿 E2EFDA);JOB 页隐藏「行预警」列同步去 emoji。Mac/Win 一致。
- **按钮重排**:顶栏 7 键 → 标题 + 数据/台账徽标 + 「⚙ 设置」;日常三键(保存/生成/打开)独立 actionbar。「从飞书导出导入」「从当前 Excel 同步」**删除**(后者删后 core 的 `load_from_workbook`/`merge_workbook`/`reconcile_store_with_excel` 一并移除)。「设置」面板放文件路径更换 + 检查更新。
- **手动升级**:`app.py` 加 `__version__`、`check_update()`/`download_update()`(标准库 urllib/tarfile);tag `v*` 触发 CI 出 GitHub Release(`BKTC_Ledger-macOS.tar.gz` / `-Windows.tar.gz`);应用内「设置→检查更新→下载并打开」解压到 `~/Downloads/BKTC_Ledger_update/`,用户手动替换(运行中的 exe/app 不自替换)。
- **备份整理**:生成台账不再在同目录堆 `_备份_TS.xlsx`,改 `备份/BKTC_TS.xlsx`(`core.backup_xlsx`,留近 10 份)。
- **仓库公开**:`.gitignore` 加 `*.records.json`/`*.xlsx`/`备份/`/`*.bak`;HANDOFF §6 的飞书/钉钉 base id、corp_id 已 scrub。⚠️ `app.py` 的 DEFAULT 路径仍含本机用户名(仅路径字符串,非数据),公开可接受。

### 2026-08-04 v1.1.1–v1.1.4（定稿补丁，最新 = v1.1.4）
- **v1.1.1** 检查更新限流修复:`check_update` 改走网页 `releases/latest` 302 重定向读 tag(避 GitHub API 匿名 60/h 限流),资产 URL 直接构造 `releases/download/vX.Y.Z/BKTC_Ledger-<平台>.tar.gz`。
- **v1.1.2** 预警规则 tab 可见性(原 `switchTab` 漏把 panel-表格 在预警规则下隐藏)+ 空日期单元格显示「—」(macOS WKWebView `<input type=date value="">` 会把空值显示成今天,造成"脏数据"错觉;改点击才挂 picker)。
- **v1.1.3** 六表 Excel 式筛选+排序:点表头升/降/取消;每列表头「▾」弹该列去重值勾选(多列 AND)+搜索+全选;移除客户/JOB 下拉(被列筛选取代);保留关键字全表搜。
- **v1.1.4 拆分批次 + 派生实时重算**:发货批次页「⇲ 拆分批次」选中源批次→按**设备型号分组**勾选移出(整型号全选/挑个别机)→目标批次+出荷日→移动+自动建/合并目标行(只改设备 `发货批次`,不动开票/回款覆盖);`recomputeDerived()`(JS 镜像 `core.derive`)挂到 updateRec/增删/粘贴/重命名/拆分,批次合计/合同总台数·型号/开票回款覆盖台数 实时刷新。列筛选按钮▾放大。

### 2026-08-05/06 v1.1.5–v1.1.9（导出 / 跨表查询 / 派生同步，最新 = v1.1.9）
- **v1.1.5 单表导出 Excel**:任意六表筛选+排序后,工具栏「📥 导出 Excel」→勾字段→导出到 `~/Downloads/<表>_导出_<ts>.xlsx`;日期/数字按类型转真 Excel 值(`core.export_filtered`,支持 `{name,type}` 入参)。
- **v1.1.6 虚拟关联列(已撤)**:曾给表格加 客户/出荷日 虚拟列;用户嫌不通用,v1.1.7 撤掉、表格回单表。
- **v1.1.7 跨表查询导出(report builder)**:顶栏「🔍 跨表查询」面板——选基础表→勾任意字段(基础表+关联父表字段树)→加筛选条件→预览前10行→导出。父表 join(多对一不炸行):设备台账→合同/批次;开票/回款→合同/付款条件;批次/条款→合同。lookup:`contractByJob`/`batchByJobBatch`/`termByJobKind`(`buildLookups`)。
- **v1.1.8 筛选值改下拉**:跨表查询筛选「值」从手填改为该字段去重值下拉(`distinctValues`+`populateValueSelect`,选字段后填);操作符 等于/不等于/大于/小于/为空/不为空;未选值不生效。
- **v1.1.9 三处派生字段自动同步**:`合同订单.付款条件`(=付款条件.说明按；拼接)、`开票记录.应收回款日`(=开票日+账期天数)、`设备台账.验收状态`(=质保开始日 有无) 加入 `DERIVED`+`derive()`+`recomputeDerived()`,表格内变只读、编辑源字段即更新。数据验证 0 不一致。

## 4. 打包与分发(GitHub CI = 主路径)

**仓库**:https://github.com/danielzhfzh-hue/BKTC_Ledger(私有,`gh` 已登录 `danielzhfzh-hue`)。**源码 15 个文件,无数据**。

**工作流** `.github/workflows/build.yml`:矩阵 `[windows-latest, macos-latest]`;`push` 到 main/master 或手动 `workflow_dispatch` 触发;setup-python 3.11 → `pip install pywebview openpyxl pyinstaller` → `pyinstaller --windowed --name BKTC_Ledger --add-data "ui{;|:}ui" --collect-all webview --hidden-import webview.platforms.edgechromium app.py` → 上传 `dist/` 为 `BKTC_Ledger-Windows` / `BKTC_Ledger-macOS`。

**改了源码后怎么出包**:push 到 main(自动构建),或网页 Actions →「Run workflow」。约 5–10 分钟。

**拿产物**:
```bash
gh run list --workflow=build.yml --limit 1            # 找最新 run id
gh run download <run-id> -n BKTC_Ledger-Windows -D ./out   # 下载解压到 ./out
# 或网页:Actions → 点 run → Artifacts → 下 BKTC_Ledger-Windows
```

**发版与升级(v1.1+,主路径)**:改 `app.py` 的 `__version__` → commit → `git tag vX.Y.Z && git push --tags` → CI 自动构建 win+mac 并发布到 GitHub Releases(资产 `BKTC_Ledger-macOS.tar.gz` / `BKTC_Ledger-Windows.tar.gz`)。用户在应用「⚙ 设置 → 检查更新」见新版,点「下载并打开」自动下载解压到 `~/Downloads/BKTC_Ledger_update/`,退出本程序后用新版本替换旧文件夹(未签名,首开警告照旧)。仓库公开 → 下载免 token。

**Mac 本地 .app(备选)**:`BKTC_Ledger/.venv`(3.11)→ `.venv/bin/pyinstaller --noconfirm --windowed --clean --name BKTC_Ledger --add-data "ui:ui" --collect-all webview --hidden-import webview.platforms.edgechromium app.py` → `dist/BKTC_Ledger.app`(ad-hoc 签名,首次右键打开)。

**Windows 本地构建(备选)**:Win10/11 + Python 3.11 → 双击 `build_win.bat`(纯 ASCII)→ `dist\BKTC_Ledger\`。

**首次运行警告**(未签名,内部用,不可避免):Mac 右键→打开;Windows SmartScreen「更多信息→仍要运行」+ 杀软加白名单。

**Windows 用到的数据(只拷一个文件)**:`ledger.records.json`(= `BKTC上海POU营业管理表.records.json` 改 ASCII 名,733K,六表+规则)。Windows 上 exe →「打开数据文件…」选它 →「选择台账文件…」选 xlsx 路径 →「生成台账」。xlsx 不用拷(工具重新生成)。

## 5. 运行与验证(本机 Mac)

```bash
# 启动(首次自动建 .venv)
cd ~/projects/订单整理/BKTC_Ledger && .venv/bin/python app.py
# 或双击 启动.command(已改 python3.11)

# 健康检查(0 error 为健康)
.venv/bin/python -c "import core; d=core.derive(core.load_store('/Users/danielzhu/projects/订单整理/BKTC上海POU营业管理表.records.json')); print('未回收', core.summary(d)['未回收行数'], 'err', len([i for i in core.validate(d) if i['severity']=='error']))"

# 不碰真实数据,验证生成(输出 /tmp)
.venv/bin/python -c "import core; core.generate_xlsx(core.load_store('/Users/danielzhu/projects/订单整理/BKTC上海POU营业管理表.records.json'), '/tmp/_check.xlsx')"
```

## 6. 飞书 / 钉钉 同步现状

### 飞书(base `<飞书 base app_token>`,6 表)
- 614 设备/133 批次/40 合同/87 条款/102 开票/112 回款;双向 link + 4 运营视图 + 公式字段。
- **勿跑** `feishu_rebuild.py`(删表丢视图/公式);旧 `feishu_relational_import.py` 只 append 且读 stale 的 per-JOB CRM xlsx,不用。
- **已用 `feishu_upsert.py` 同步(2026-08-04)**:全量 upsert(本地为准),12 改 + 2 增,重跑 dry-run 0 差。脚本默认 dry-run,`--apply` 才写;只同步 SAFE 白名单主字段(**不碰**链接/公式/开票状态/回款状态/覆盖派生——records.json 覆盖製造番号为空会让旧 compute_status 把 93 条已回款误判超期,故状态交飞书侧维护);按 record_id 匹配只增改不删;日期按 calendar day 比(消 8h 时差)、数字容忍 0.01、选择字段 list 归一化。

### 钉钉(base `<钉钉 base id>`,6 表)
- **已用 `dingtalk_push.py` 镜像(2026-08-04)**:6 表 + 全量 40/87/615/134/102/112;字段 text/number/dateTime(选择类用 text);每批 100。
- **位置 = 个人空间**(非 BKTC上海营业部)。dws OAuth 缺团队空间 aitable 写权限(create/copy `--folder-id` 均拒团队 folder-id,drive mkdir 却能用)。**需钉钉 UI 手动移动**到 BKTC上海营业部,或给 dws 应用授「团队空间 aitable 编辑」权限后 CLI copy。
- `dws` 已认证(corp_id `<corp_id>`)。

## 7. 遗留低优先(未修)
| # | 位置 | 问题 | 修法 |
|---|---|---|---|
| ② | `build_ledger_main.py:789` | 未回收表 static 行(26BS009 式)写 Python 预警字面量、公式行写公式,填色按 Python 值 → 跨午夜文本/底色不符(纯显示) | static 行也用公式 |
| ③ | `build_ledger_main.py:643` | 无设备且条款无比率的 JOB 被 continue 静默跳过 | total=None 落一条 ④待确认 |
| ④ | `build_ledger_main.py:276` | 无设备合成行每款类只取首笔,多笔少算 | 该款类汇总 |
| ⑦ | `build_ledger_main.py:154` | openpyxl 把 `=` 开头文本写成公式(变 #NAME?),0 命中 | set_cell 前置 `'` |

## 8. 坑 / 注意

- **records.json 是权威源**,Excel 没有付款条件/覆盖串/应收日结构化字段——别丢;丢了只能从 `.bak` 恢复。
- 生成台账的备份在 `备份/` 子目录(`core.backup_xlsx`,留近 10 份);不再在根目录堆 `_备份_*.xlsx`。
- 「从当前 Excel 同步」功能已于 v1.1 **移除**(records.json 是唯一权威源;旧 `merge_list` 会丢 parse 不回的记录)。如需从旧 Excel 救数据,临时回退到 v1.0。
- **Windows 一切用 ASCII**:文件夹/程序名/.bat/数据文件名都不能有中文(乱码);`build_win.bat` 已纯 ASCII;数据文件用 `ledger.records.json`。
- 勿跑 `feishu_rebuild.py`;`transform_relational.py` 缺近期修复,重跑回退。
- `build_ledger_main.py` 末尾 `verify()` 仍硬编码 0.01(工具不调,仅 standalone 跑用)。
- 脏数据(26BS009 空单/22BS006 人工状态/23BS004-058/-061 双记录/源表 L 列 10 倍笔误)见 `订单整理/HANDOFF.md` §三。
- 设备业务键 = `(JOB No, 製造番号, 機番)`;款类 5 规范名(预付/发货/到货/验收/质保 + 全额);含税 = 未税 × 1.13。

### 派生字段：已自动 / 决定不做 / 可做未做（2026-08-06 梳理）
- **已自动派生（表格内只读，编辑源字段）**：发货批次(台数/未税合计/含税合计/覆盖製造番号)、合同订单(总台数/设备型号/**付款条件文本**←付款条件.说明拼接)、设备台账(**验收状态**←质保开始日有无)、开票记录(**应收回款日**←开票日+账期天数)、开票/回款(覆盖台数)。逻辑在 `core.derive()` + JS `recomputeDerived()`。
- **决定不做（维持手录）**：① 开票/回款 **含税金额**——两表无"未税金额"字段，是实际开票/收款金额，直接录入（不是 未税×1.13 的派生；要那样得新增未税字段+回填，用户选择不改）。② 回款 **对应应收回款日**——51 个(JOB,款类)组里 15 组多发票/多期、27 条回款对应组无开票，不够干净。
- **可做未做（已验证干净，待用户确认）**：设备台账 **质保期** = `round((质保结束日 − 质保开始日).days / 365)年`（499/502 一致；3 个不一致是数据标错，派生会自动纠正为 1年）。需时加进 `DERIVED`+`derive`+`recomputeDerived`。

## 9. 相关文档

- `BKTC_Ledger/README.md` — 工具使用指南(用户向,操作对照表)
- `BKTC_Ledger/README_win.md` — Windows 构建/分发/首次运行说明
- GitHub:https://github.com/danielzhfzh-hue/BKTC_Ledger — 源码 + Actions 打包
- `订单整理/HANDOFF.md` — 旧飞书 pipeline + 数据口径修复记录 + 脏数据清单
- `订单整理/CLAUDE.md` — `transform_relational.py` 转换规则(旧路径)
