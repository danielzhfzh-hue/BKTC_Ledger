# BKTC 台账维护工具 — 交接文档

更新：2026-08-04　·　维护者入口：先读本文，再看 [README.md](README.md)（用户向）

## 0. 一句话现状

`BKTC_Ledger`（pywebview 桌面应用）是**当前唯一的编辑入口**；`BKTC上海POU营业管理表.records.json` 是**权威数据源**；`build_ledger_main.py` 是**生成引擎**（已并入工具,自包含）；`BKTC上海POU营业管理表.xlsx` 是**生成产物**（42 sheet）。**打包走 GitHub Actions**(win+mac CI);飞书/钉钉已同步(见 §6)。

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
              └── 从当前 Excel 同步 ◄── BKTC上海POU营业管理表.xlsx
                  (load_from_workbook + merge_workbook)
```

- **保存** `save_store` → `derive`(normalize+派生) → 覆盖 records.json(覆盖前 `shutil.copy2` → `.bak`)
- **生成** `generate_xlsx` → `derive`+`validate` → `build_from_data` → 备份 xlsx → 重写 导航页/未回收管理表/40 JOB 页
- **同步回** `sync_from_excel` → `load_from_workbook`(解析可见单元格+隐藏公式锚点) + `merge_workbook`(并入 store,丢 parse 不回的记录) → 覆盖 records.json(有 `.bak`)

## 2. 关键文件与位置

> 工具根目录 = **git 仓库** `~/projects/订单整理/BKTC_Ledger/`(2026-08-04 从 `~/Documents/销售订单管理多维表格构筑/台账维护工具/` 迁入并自包含)。**GitHub 私有仓**:https://github.com/danielzhfzh-hue/BKTC_Ledger(只有源码,无数据)。

| 文件 | 位置 | 角色 |
|---|---|---|
| `app.py` | `BKTC_Ledger/` | pywebview 壳 + `Api`;`APP_DIR` 已支持 `sys._MEIPASS` |
| `core.py` | `BKTC_Ledger/` | 数据层:`SCHEMA`/`normalize`/`derive`/`validate`/`save_store`/`generate_xlsx`/`load_from_workbook`/`merge_workbook` |
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
- 「从当前 Excel 同步」用 `merge_list` 会丢 parse 不回的记录(隐藏公式锚点脆弱);有 `.bak` 兜底但**务必看 toast 里的 `-Z`**。
- **Windows 一切用 ASCII**:文件夹/程序名/.bat/数据文件名都不能有中文(乱码);`build_win.bat` 已纯 ASCII;数据文件用 `ledger.records.json`。
- 勿跑 `feishu_rebuild.py`;`transform_relational.py` 缺近期修复,重跑回退。
- `build_ledger_main.py` 末尾 `verify()` 仍硬编码 0.01(工具不调,仅 standalone 跑用)。
- 脏数据(26BS009 空单/22BS006 人工状态/23BS004-058/-061 双记录/源表 L 列 10 倍笔误)见 `订单整理/HANDOFF.md` §三。
- 设备业务键 = `(JOB No, 製造番号, 機番)`;款类 5 规范名(预付/发货/到货/验收/质保 + 全额);含税 = 未税 × 1.13。

## 9. 相关文档

- `BKTC_Ledger/README.md` — 工具使用指南(用户向,操作对照表)
- `BKTC_Ledger/README_win.md` — Windows 构建/分发/首次运行说明
- GitHub:https://github.com/danielzhfzh-hue/BKTC_Ledger — 源码 + Actions 打包
- `订单整理/HANDOFF.md` — 旧飞书 pipeline + 数据口径修复记录 + 脏数据清单
- `订单整理/CLAUDE.md` — `transform_relational.py` 转换规则(旧路径)
