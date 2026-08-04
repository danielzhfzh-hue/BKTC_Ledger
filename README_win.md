# Windows 构建与使用说明

> 本程序无法在 Mac 上跨平台编译出 Windows exe。需在 **Windows 10/11** 机器上用本目录的 `build_win.bat` 一键构建。

## 一、前置要求

- **Python 3.11**(或 3.12):到 https://www.python.org/downloads/ 下载,安装时勾选 **Add Python to PATH** 与 **py launcher**。
- **Windows 10/11**:WebView2 运行时(Win10/11 一般自带;没有就装 [Microsoft Edge WebView2 Evergreen](https://developer.microsoft.com/microsoft-edge/webview2/))。
- 本目录文件(已随工具迁移过来):`app.py` / `core.py` / `build_ledger_main.py` / `ui/` / `build_win.bat`。

## 二、构建(一条命令)

在本目录(台账维护工具)下双击 `build_win.bat`,或命令行:

```bat
cd /d 台账维护工具所在目录
build_win.bat
```

脚本会:建 `.venv` → 装 `pywebview openpyxl pyinstaller` → PyInstaller 打包。产物在 **`dist\BKTC_Ledger\`**(onedir 整文件夹)。

> 文件夹/exe 名用 ASCII(`BKTC_Ledger`)是为了避开 Windows 中文乱码。若想换名,改 `build_win.bat` 里 `--name`。

## 三、分发

把 **整个 `dist\BKTC_Ledger\` 文件夹** 拷给使用者(不是只拷 exe——onedir 依赖同目录的 DLL/资源)。使用者双击 `BKTC_Ledger.exe` 启动。

## 四、首次运行警告(未签名,不可避免)

PyInstaller 打包且**未购买代码签名证书**的程序,Windows 首次运行必然报警——这是行业通病,不是病毒:

1. **SmartScreen「Windows 保护了你的电脑」** → 点「**更多信息**」→ 点「**仍要运行**」。
2. **杀毒软件(Defender / 360 / 火绒等)误报**:
   - 把 `dist\BKTC台账维护工具\` 整个文件夹加入杀软**白名单/排除项**;
   - 或提交 [Microsoft 信誉库](https://www.microsoft.com/en-us/wdsi/filesubmission)做信誉(降低长期误报,需数天)。

> 彻底消除警告需购买 **Windows 代码签名证书(OV/EV)**,内部用可不必。

## 五、数据文件(首次启动)

程序默认的数据路径是 macOS 的,Windows 上首次启动会**找不到数据**(显示空)。任选一种指给它:

- **界面**「打开数据文件…」选 `BKTC上海POU营业管理表.records.json`,「选择台账文件…」选 `BKTC上海POU营业管理表.xlsx`(两个文件都从 Mac 拷过来);
- **环境变量**:`BKTC_STORE=<...records.json 的 Win 路径>`、`BKTC_XLSX=<...xlsx 的 Win 路径>`;
- **命令行**:`BKTC台账维护工具.exe --xlsx "D:\...\台账.xlsx" --store "D:\...\台账.records.json"`。

## 六、以后数据更新后

records.json 在 Windows 这边改/同步后,直接在程序里「保存数据 → 生成台账」即可,不用重新构建 exe。只有改了 `app.py`/`core.py`/`build_ledger_main.py`/`ui/` 才需重跑 `build_win.bat`。
