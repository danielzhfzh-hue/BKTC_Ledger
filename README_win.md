# Windows 构建与使用

## 运行要求

- Windows 10/11 与 Edge WebView2（通常已内置）。
- 源码运行/本地构建需要 Python 3.11 或 3.12；安装时勾选 `Add Python to PATH`。

## 直接启动

- Release 压缩包：完整解压后双击 `BKTC_Ledger\BKTC_Ledger.exe`，不要只复制 exe。
- 源码压缩包：双击 `启动.bat`；脚本会检查 Python、按需安装依赖，并在缺失时创建 `data\BKTC_Ledger.db` / `data\BKTC_Ledger.xlsx`。

## 构建

在项目目录双击 `build_win.bat`，或运行：

```bat
build_win.bat
```

脚本会建立 `.venv`、安装 `pywebview openpyxl pyinstaller`，并生成 `dist\BKTC_Ledger\`。分发时必须复制整个文件夹，而不是只复制 exe；数据库和 Excel 位于同级 `data\` 子目录。

GitHub Actions 是推荐构建路径：Windows 与 macOS 均会先运行可靠性测试，再打包产物。
标签发布时 Windows 产物使用 `.zip`，可直接用 Win11 资源管理器解压。

## 首次选择数据

推荐给数据库使用 ASCII 文件名，例如 `ledger.db`：

- 界面：设置 →「打开 / 迁移…」选择 `ledger.db`；若只有旧 `ledger.records.json`，直接选择它，应用会生成同目录 `ledger.db` 并保留原 JSON。
- 文件选择器不可用时：把完整路径粘贴到设置页输入框，再点「应用路径」。
- 环境变量：`BKTC_DATABASE=D:\data\ledger.db`、`BKTC_XLSX=D:\data\ledger.xlsx`。
- 命令行：`BKTC_Ledger.exe --database "D:\data\ledger.db" --xlsx "D:\data\ledger.xlsx"`。

界面选择或粘贴的 DB/XLSX 路径会保存到当前 Windows 用户的 `%APPDATA%\BKTC_Ledger\config.json`，下次双击启动会自动恢复；命令行参数和环境变量优先于该配置。

旧版 `BKTC_STORE=<records.json>` / `--store <records.json>` 仍可用于一次性迁移。

本机交付包的 `data\BKTC_Ledger.db` / `data\BKTC_Ledger.xlsx` 是当前项目数据；公开 GitHub Release 仍不上传客户数据。复制数据库前先退出应用，或复制 `备份/` 中的完整快照，避免遗漏 WAL 中尚未归档的事务。

## Windows 首次运行警告

当前内部版本未购买代码签名证书，SmartScreen 可能提示未知发布者：点「更多信息」→「仍要运行」。杀毒软件若误报，应对整个 `BKTC_Ledger` 文件夹添加排除，而不是只保留 exe。

彻底消除信誉警告需要 OV/EV 代码签名证书；这不影响 SQLite 数据完整性或应用功能。
