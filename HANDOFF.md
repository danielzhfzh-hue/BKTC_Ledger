# BKTC 台账维护工具 — 交接文档

更新：2026-08-03　·　维护者入口：先读本文，再看 [README.md](README.md)（用户向使用指南）

## 0. 一句话现状

`台账维护工具`（pywebview 桌面应用）是**当前唯一的编辑入口**；`BKTC上海POU营业管理表.records.json` 是**权威数据源**；`build_ledger_main.py` 是**生成引擎**；`BKTC上海POU营业管理表.xlsx` 是**生成产物**（42 sheet）。飞书已同步、钉钉已镜像（在个人空间，待 UI 移入团队空间，见 §6）。

> 旧文档 `订单整理/HANDOFF.md` 描述的是「飞书 pipeline 为权威源」的旧架构，已被本工具取代——以本文为准。旧文档仍可用于查数据口径修复记录（24BS017/25BS031 等）和脏数据清单。

## 1. 架构与数据流

```
┌─ 台账维护工具 (pywebview) ─┐    六标签页编辑 → state.data
│  core.py  数据层            │
│  app.py   桌面壳 + Api      │
│  ui/      HTML/JS/CSS       │
└─────────────┬───────────────┘
              │ 保存数据        │ 生成台账
              ▼                ▼
   records.json (权威源)   build_ledger_main.build_from_data
   六表 + 预警规则          ↓
   覆盖前 → .bak           备份 xlsx → 重写 42 sheet
              ↑                ↓
              └── 从当前 Excel 同步 ◄── BKTC上海POU营业管理表.xlsx
                  (load_from_workbook + merge_workbook)
```

- **保存** `save_store` → `derive`(normalize+派生) → 覆盖 records.json（覆盖前 `shutil.copy2` → `.bak`）
- **生成** `generate_xlsx` → `derive`+`validate` → `build_from_data` → 备份 xlsx → 重写 导航页/未回收管理表/40 JOB 页
- **同步回** `sync_from_excel` → `load_from_workbook`(解析可见单元格+隐藏公式锚点) + `merge_workbook`(并入 store，丢 parse 不回的记录) → 覆盖 records.json（有 `.bak`）

## 2. 关键文件

> 工具根目录:`~/projects/订单整理/BKTC_Ledger/`(2026-08-04 从 `~/Documents/销售订单管理多维表格构筑/台账维护工具/` 迁入,挨着数据;**已自包含**)。

| 文件 | 位置 | 角色 |
|---|---|---|
| `core.py` | `BKTC_Ledger/` | 数据层：`SCHEMA`/`normalize`/`derive`/`validate`/`save_store`/`generate_xlsx`/`load_from_workbook`/`merge_workbook` |
| `app.py` | `BKTC_Ledger/` | pywebview 壳 + `Api`;`APP_DIR` 已支持 `sys._MEIPASS`(打包冻结态) |
| `ui/{index.html,app.js,style.css}` | `台账维护工具/ui/` | 前端（行虚拟化、多行粘贴、批次重命名） |
| `build_ledger_main.py` | `BKTC_Ledger/`(已并入,**工具自包含**) | **生成引擎**:`build_from_data`/`write_job_sheet`/`compute_unpaid_rows`/`build_unpaid_sheet`/`verify`;core.py 直接 `from build_ledger_main import`。⚠️ 父目录 `~/Documents/销售订单管理多维表格构筑/` 另有 **legacy 副本**(供旧脚本 transform/validate);**改引擎只改工具内这份** |
| `BKTC上海POU营业管理表.records.json` | `订单整理/` | **权威源**（六表 + 预警规则；含 Excel 不展示的付款条件/覆盖串/应收日） |
| `BKTC上海POU营业管理表.xlsx` | `订单整理/` | 生成产物（导航页 + 未回收管理表 + 40 JOB 页） |
| `*.records.json.bak` | `订单整理/` | records.json 的上次保存快照（2026-08-03 起自动滚动备份） |
| `feishu_upsert.py` / `dingtalk_push.py` | `订单整理/` | records.json → 飞书 upsert（dry-run/--apply）/ 钉钉 aitable 全量镜像 |
| `build_win.bat` / `README_win.md` | `BKTC_Ledger/` | Windows 构建脚本 + 说明(用户在 Win 机器跑,出 onedir 文件夹) |
| `dist/BKTC台账维护工具.app` | `台账维护工具/dist/` | Mac 打包产物(ad-hoc 签名,首次右键→打开);重建见 §5 |

## 3. 2026-08-03 修复（本轮，已验证）

| # | 文件:行 | 改动 | 为什么 |
|---|---|---|---|
| 1 容差口径 | `build_ledger_main.py:204,556,912` | `write_job_sheet` 加 `rules` 参数；行预警公式 `+0.01` → `+{tol:g}`；`build_from_data` 透传 `rules` | 金额容差原本只对未回收表生效，JOB 页行预警 + 导航页 🔴/🟡 写死 0.01，改容差后三处口径冲突 |
| 2 records 备份 | `core.py:268` | `save_store` 覆盖前 `shutil.copy2(path, path+".bak")` | 权威源被保存/同步直接覆盖，无 undo |
| 3 日期解析 | `core.py:97` + `validate` | `_norm` DATE 放宽 `[-/.年]` + 单数字；失败**保留原值**（不再清空）；`validate` 加日期格式 warning | 粘贴 `2025/1/5`/`2025年1月5日` 等非标准日期被静默清空 |
| 4 数字清空 | `core.py:88` + `app.js:270` | `_norm` NUMBER/INT 保留 `null`；UI 清空存 `null`（不再变 0） | 清空未税单价→0→设备变免费，拉低合同总额 |
| 5 粘贴带 JOB | `app.js:449` | `pasteGrid` 新建行带 `state.job`（合同订单还带 `state.customer`） | 与「＋添加行」按钮不一致，粘贴新行无 JOB 触发校验错 |

验证：真实 records.json → /tmp 生成，42 sheet / 79 未回收行 / ¥29,725,308.67 / 0 error；`write_job_sheet` 行预警随 `金额容差` 变（None/0.01→`+0.01`、5.0→`+5`）。

## 4. 遗留低优先（未修，按优先级）

| # | 位置 | 问题 | 修法提要 |
|---|---|---|---|
| ② | `build_ledger_main.py:789` | 未回收表 static 行(26BS009 式)写 Python 预警字面量、公式行写 Excel 公式，填色按 Python 值 → 跨午夜文本与底色不符（纯显示） | static 行也用公式驱动，消除双口径 |
| ③ | `build_ledger_main.py:643` | 无设备且条款无比率的 JOB 被 `continue` 静默跳过，有钱款却不在未回收表 | total=None 时落一条 ④待确认 |
| ④ | `build_ledger_main.py:276` | 无设备合成行对每款类只取首笔，多笔同款类少算进导航合计 | 该款类汇总所有笔 |
| ⑦ | `build_ledger_main.py:154` | openpyxl 把 `=` 开头自由文本写成公式（变 `#NAME?`），当前 0 命中 | `set_cell` 对 `=+-@` 开头字符串前置 `'` |

> ①⑤⑥ 已修（见 §3）。②③④ 是边角单隐患，⑦ 是防御性。

## 5. 运行与验证

```bash
# 启动工具（首次自动建 .venv 装依赖）
双击 台账维护工具/启动.command
# 或
cd ~/projects/订单整理/BKTC_Ledger && .venv/bin/python app.py

# 健康检查：当前数据校验（0 error 为健康）
cd ~/projects/订单整理/BKTC_Ledger
.venv/bin/python -c "import core; d=core.derive(core.load_store('/Users/danielzhu/projects/订单整理/BKTC上海POU营业管理表.records.json')); print('未回收', core.summary(d)['未回收行数'], 'err', len([i for i in core.validate(d) if i['severity']=='error']))"

# 不碰真实数据，验证生成（输出到 /tmp）
.venv/bin/python -c "import core; core.generate_xlsx(core.load_store('/Users/danielzhu/projects/订单整理/BKTC上海POU营业管理表.records.json'), '/tmp/_check.xlsx')"

# 审计 Excel 结构（sheet/公式/隐藏列/数据瑕疵）
# openpyxl data_only=False 读公式(见历次会话 /tmp/inspect_ledger.py)

# 打包 Mac .app(本机,Python 3.11,onedir,ad-hoc 签名)
cd ~/projects/订单整理/BKTC_Ledger
.venv/bin/pyinstaller --noconfirm --windowed --clean --name "BKTC台账维护工具" \
  --add-data "ui:ui" --collect-all webview --hidden-import "webview.platforms.edgechromium" app.py
# → dist/BKTC台账维护工具.app ; 首次右键打开(Gatekeeper 未识别开发者)

# Windows:在 Win10/11 机器跑 build_win.bat(见 README_win.md)→ dist/BKTC台账维护工具/(onedir 文件夹)
```

## 6. 飞书 / 钉钉 同步现状（重要）

### 飞书（base `<飞书 base app_token>`，6 表）
- 现有 614 设备 / 133 批次 / 40 合同 / 87 条款 / 102 开票 / 112 回款；有双向 link + 4 运营视图 + 公式字段（应收监控/距应收天数等）。
- **现有脚本都不能直接用**：
  - `feishu_relational_import.py` 只 `batch_create`（append），从 per-JOB CRM xlsx 读；CRM xlsx 已 stale（36/42 缺 `覆盖製造番号`），且不带 records.json。
  - `feishu_rebuild.py` 删表重建 → **丢视图+公式，勿跑**。
- **没有 records.json → 飞书 的路径**。要同步需新写 upsert 脚本。
- **好消息**：records.json 里所有存量记录都带飞书 `record_id`（`recvr1Z...`），可按 ID 增量更新；只有今天加的 HOT N2 设备是本地 ID（`rec-xlsx-设备台账-614`）需新增。
- **同步状态（2026-08-04 完成）**：用 `feishu_upsert.py` 全量 upsert（本地为准），12 改 + 2 增（HOT N2 设备+批次入库、25BS011 总台数 67→68、4 機番、5 含税金额修正、26BS009 总台数 4→0 等），重跑 dry-run 0 差异。新 record_id 已回写 records.json。
- **upsert 脚本要点**（`订单整理/feishu_upsert.py`）：默认 dry-run，`--apply` 才写；只同步 SAFE 白名单主字段（不碰 链接/公式/开票状态/回款状态/覆盖派生——records.json 的覆盖製造番号为空会让 compute_status 误判 93 条已回款→超期，故状态字段交飞书侧维护）；按 record_id 匹配（存量带 `recvr1Z...`），无 ID 的 insert，**只增改不删**；日期按 calendar day 比（消 8h 时差）、数字容忍 0.01、选择字段 list 归一化。

### 钉钉 ai 表格
- **同步状态（2026-08-04 完成）**：base `<钉钉 base id>`「BKTC POU 订单台账」已建，6 表 + 全量记录 40/87/615/134/102/112。脚本 `订单整理/dingtalk_push.py`（复用 FIELD_ORDER/conv/NAME_MAP，与飞书同值；字段类型 text/number/dateTime，选择类用 text；每批 100，钉钉上限）。
- **位置 = 个人空间**（非 BKTC上海营业部）。dws 的 OAuth 缺团队空间 aitable 写权限：`aitable create --folder-id` 与 `base copy --target-folder-id` 均拒团队空间 folder-id（drive mkdir 却能用）。**需钉钉 UI 手动移动**到 BKTC上海营业部（…菜单 → 移动）；或给 dws 应用授「团队空间 aitable 编辑」权限后 CLI copy 即可。
- `dws` CLI 已认证（token 有效，corp_id `<corp_id>`）。

## 7. 坑 / 注意

- **records.json 是权威源**，Excel 没有 付款条件/覆盖串/应收日 结构化字段——别丢 records.json；丢了只能从 `.bak` 恢复。
- 「从当前 Excel 同步」用 `merge_list` 会丢 parse 不回的记录（依赖隐藏公式锚点，脆弱）；有 `.bak` 兜底但**务必看 toast 里的 `-Z`**。
- 勿跑 `feishu_rebuild.py`（丢飞书视图/公式）；`transform_relational.py` 缺近期修复（25BS031/25BS011/24BS017/24BS009 等），重跑会回退。
- `build_ledger_main.py` 末尾 `verify()` 仍硬编码 0.01（1145 行附近）——工具不调它（仅 `python3 build_ledger_main.py` 独立跑用），改容差后 standalone verify 会误报；需要时一并改。
- 脏数据（26BS009 空单 / 22BS006 人工状态 / 23BS004-058/-061 双记录 / 源表 L 列 10 倍笔误）见 `订单整理/HANDOFF.md` §三。
- 设备业务键 = `(JOB No, 製造番号, 機番)`；款类 5 规范名（预付/发货/到货/验收/质保 + 全额）；含税 = 未税 × 1.13。

## 8. 相关文档

- `台账维护工具/README.md` — 工具使用指南（用户向，含操作对照表）
- `订单整理/HANDOFF.md` — 旧飞书 pipeline + 数据口径修复记录 + 脏数据清单
- `订单整理/CLAUDE.md` — `transform_relational.py` 转换规则（旧路径）
- `销售订单管理多维表格构筑/飞书与本地数据一致性校验报告_2026-08-03.md` — 上次全量一致性校验
