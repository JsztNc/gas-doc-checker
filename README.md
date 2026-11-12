# 线上资料检查工具

给一个工程编号，自动用 Everything 在本机/共享盘找到 `<编号>-b.线上资料.docx`，
拆开整份资料对照《线上燃气竣工资料填写规范与校验要点》逐项比对，列出会被退回的硬伤。

Windows 10/11 64 位免安装，双击 `线上资料检查.exe` 即用。

## 界面

- 输入工程编号 → 回车或点「开始检查」
- 结果按 错误 / 警告 / 提示 三级着色
- 底部按钮：打开文档、自动修正、打开报告、报告文件夹
- 右上角「检查更新」：连 GitHub Releases 检查新版本，有新版则一键下载、自动替换并重启

详细用法见 [使用说明.txt](使用说明.txt)。

## 目录结构

| 路径 | 说明 |
| --- | --- |
| `线上资料检查.exe` | 打包好的入口（PyInstaller one-folder） |
| `_internal/` | 运行时依赖，不能删 |
| `ui.py` | WinForms 窗口界面（含检查更新按钮） |
| `check.py` | 检查主程序，命令行入口 |
| `updater.py` | 自动更新：GitHub API 查版本 → 下载 zip → 生成替换脚本 → 重启 |
| `version.txt` | 当前版本号，发新版时改这里 |
| `checker/` | 解析与规则（docx_reader / rules / finder / fixer） |
| `bin/es.exe` | Everything 命令行搜索工具 |
| `检查报告/` | 运行生成的报告（gitignore，不入库） |

## 从源码运行

需要 Python 3.9+ 和 `pythonnet`：

```bat
pip install pythonnet pywin32
python ui.py
python check.py NC23A005
```

## 发布新版本

1. 修改 `version.txt` 的版本号，例如 `1.1.0`
2. 把整个程序目录（含 `_internal`、`bin`、`checker`、各 `.py`、`version.txt`、
   使用说明、bat）打包成 zip，**顶层必须是程序文件本身**，不要多套一层文件夹。
   例如 zip 里直接是 `线上资料检查.exe`、`_internal/`、`check.py`……
3. 在 GitHub 上建 Release，tag 填 `v1.1.0`，把这个 zip 作为附件上传
4. 客户端点「检查更新」即可自动下载替换

自动更新只替换程序自带文件（`_internal/`、`checker/`、`bin/`、各 `.py/.bat/.exe`、
`version.txt`、使用说明），**不会动 `检查报告/` 和用户自己放进去的文件**。

## 从网盘下载后闪退 / 报 Python.Runtime.Loader.Initialize

Windows 会给从网盘或邮件下载的文件打上"来自网络"标记（Zone.Identifier），
.NET 会拒绝加载带这个标记的 DLL，导致程序一启动就崩溃。

**解决**：双击目录下的 `解除网络标记.bat`，然后重新打开程序（一次即可）。
或手动：右键程序文件夹 → 属性 → 勾选"解除锁定" → 确定。

## 自动更新原理

`updater.py` 只用 Python 标准库：

1. `GET https://api.github.com/repos/lonyii/gas-doc-checker/releases/latest`
2. 比较 `tag_name` 与本地 `version.txt`
3. 有新版则下载 zip 到 `_tmp/update/release.zip` 并解压
4. 在系统临时目录生成一个 bat：等主程序退出 → 删旧文件 → 拷新文件 → 重启程序
5. 主程序关闭，由 bat 完成替换（Windows 下运行中的 exe 不能覆盖自己）
