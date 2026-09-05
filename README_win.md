# Windows 构建与使用

## 运行要求

- Windows 10/11 与 Edge WebView2（通常已内置）。
- 源码运行/本地构建需要 Python 3.11 或 3.12；安装时勾选 `Add Python to PATH`。

## 直接启动

- Release 压缩包：完整解压后双击 `BKTC_Ledger\BKTC_Ledger.exe`，不要只复制 exe。
- 源码压缩包：双击 `启动.bat`；脚本会检查 Python、按需安装依赖，并在缺失时创建 `data\BKTC_Ledger.db` 及其初始 Excel 导出文件。

## 构建

在项目目录双击 `build_win.bat`，或运行：

```bat
build_win.bat
```

脚本会建立 `.venv`、安装 `pywebview openpyxl pyinstaller`，并生成 `dist\BKTC_Ledger\`。分发时必须复制整个文件夹，而不是只复制 exe；数据库及其 Excel 导出文件位于同级 `data\` 子目录。

GitHub Actions 是推荐构建路径：Windows 与 macOS 均会先运行可靠性测试，再打包产物。
标签发布时 Windows 产物使用 `.zip`，可直接用 Win11 资源管理器解压。

设置页的“检查更新”检查的是 GitHub Release，不是普通的 `main` 分支 push。发布新版本时需要先提交代码，再创建版本标签（本次为 `v1.7.1`），GitHub Actions 才会生成 Windows zip。Windows 端点击“下载并打开”会下载并解压到 `下载\BKTC_Ledger_update`，然后请退出正在运行的旧程序，用新目录替换旧程序目录；当前不会在运行中的 exe 上自动覆盖更新。

## 首次选择数据

推荐给数据库使用 ASCII 文件名，例如 `ledger.db`：

- 界面：设置 →「打开 / 迁移…」选择 `ledger.db`；若只有旧 `ledger.records.json`，直接选择它，应用会生成同目录 `ledger.db` 并保留原 JSON。
- 文件选择器不可用时：把完整路径粘贴到设置页输入框，再点「应用路径」。
- 环境变量：`BKTC_DATABASE=D:\data\ledger.db`。
- 命令行：`BKTC_Ledger.exe --database "D:\data\ledger.db"`。

界面选择或粘贴的数据库路径会保存到当前 Windows 用户的 `%APPDATA%\BKTC_Ledger\config.json`，下次双击启动会自动恢复；命令行参数和环境变量优先于该配置。Excel 不单独关联：`ledger.db` 始终导出为同目录下的 `ledger.xlsx`。

设置页的“当前操作人”也保存在该配置中。审计从 v1.5.0 启用时点开始，不回填旧数据；之后每次成功保存的新增、修改和删除可在“审计记录”页面查询和导出。

报价单在应用内保存到同一个 SQLite 数据库。报价主体可在北京康肯与 KANKEN TECHNO（日本）之间选择；付款条件支持客户优先的历史预设和中英文显示。历史报价既可按客户+型号，也可只按型号查询；“从模板创建”会重置编号、日期、有效期和交货期。新建订单可直接导入已保存报价的设备和结构化付款条款。导出文件位于当前用户的 `下载\报价单\`，不需要也不能把报价 XLSX 关联为数据源。

旧版 `BKTC_STORE=<records.json>` / `--store <records.json>` 仍可用于一次性迁移。

本机交付包的 `data\BKTC_Ledger.db` / `data\BKTC_Ledger.xlsx` 是当前项目数据；公开 GitHub Release 仍不上传客户数据。复制数据库前先退出应用，或复制 `备份/` 中的完整快照，避免遗漏 WAL 中尚未归档的事务。

如果程序目录位于 `Program Files`、只读同步目录或受控文件夹，Windows 版会自动把缺少写权限的内置数据库复制到 `%LOCALAPPDATA%\BKTC_Ledger\data`，并在设置中显示实际使用路径，请以设置页显示的路径为准。导出台账时如果 `BKTC_Ledger.xlsx` 正被 Excel 打开或标记为只读，应用会先保存数据库，再在同一目录生成带时间戳的备用 XLSX，并显示其完整路径；关闭 Excel 后即可重新导出到标准文件。

## Windows 首次运行警告

当前内部版本未购买代码签名证书，SmartScreen 可能提示未知发布者：点「更多信息」→「仍要运行」。杀毒软件若误报，应对整个 `BKTC_Ledger` 文件夹添加排除，而不是只保留 exe。

彻底消除信誉警告需要 OV/EV 代码签名证书；这不影响 SQLite 数据完整性或应用功能。
