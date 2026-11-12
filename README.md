# 线上资料检查工具

给一个工程编号，自动用 Everything 在本机/共享盘找到 `<编号>-b.线上资料.docx`，
拆开整份资料对照《线上燃气竣工资料填写规范与校验要点》逐项比对，列出会被退回的硬伤。

Windows 10/11 64 位免安装，**单文件** `线上资料检查.exe`，双击即用。

## 界面

- 输入工程编号 → 回车或点「开始检查」
- 结果按 错误 / 警告 / 提示 三级着色
- 底部按钮：打开文档、自动修正、打开报告、报告文件夹
- 右上角「使用说明」：弹出可滚动窗口显示完整说明
- 右上角「检查更新」：连 GitHub Releases 检查新版本，有新版则一键下载新 exe、自动替换并重启
- 右上角「反馈」：弹出窗口填写意见（类型/内容/选填联系方式），一键匿名提交到 GitHub Issues，无需登录账号

详细用法见 [使用说明.txt](使用说明.txt)（界面「使用说明」按钮里也是这份内容）。

## 单文件版（v1.1.0 起）

整个程序就是**一个 exe 文件**：Python 运行库、lxml、python-docx、.NET 桥、
es.exe、使用说明、version.txt 全部打在里面。拷到别的电脑双击就能跑，
不用带 `_internal` 或任何附件。第一次启动会解压几秒，属正常。

exe 旁边会自动生成 `检查报告/`（历次报告）和 `_tmp/`（临时文件，可删）。

## 源码结构（开发者用，普通用户不需要）

| 路径 | 说明 |
| --- | --- |
| `ui.py` | WinForms 窗口界面（含使用说明、检查更新按钮） |
| `check.py` | 检查主程序，命令行入口 |
| `updater.py` | 自动更新：GitHub API 查版本 → 下载新 exe → 生成替换脚本 → 重启 |
| `version.txt` | 当前版本号，发新版时改这里 |
| `checker/` | 解析与规则（docx_reader / rules / finder / fixer） |
| `bin/es.exe` | Everything 命令行搜索工具（打包时内嵌） |

## 从源码运行

需要 Python 3.9+ 和 `pythonnet`：

```bat
pip install pythonnet pywin32
python ui.py
python check.py NC23A005
```

## 版本记录

- **v1.2.0**：新增「反馈」按钮，一键匿名提交意见到 GitHub Issues。
- **v1.1.0**：改为**单文件 exe**；更新只替换 exe 本身；界面增加「使用说明」按钮；使用说明内嵌进 exe。
- **v1.0.1**：没搜到文件时显示干净提示，不再弹技术堆栈。
- **v1.0.0**：首版。

## 从网盘下载后闪退 / 报 Python.Runtime.Loader.Initialize

Windows 会给从网盘或邮件下载的文件打上"来自网络"标记（Zone.Identifier），
.NET 会拒绝加载带这个标记的 DLL，导致程序一启动就崩溃。

**解决**：双击目录下的 `解除网络标记.bat`，然后重新打开程序（一次即可）。
或手动：右键 exe → 属性 → 勾选"解除锁定" → 确定。

## 发布新版本

1. 修改 `version.txt` 的版本号，例如 `1.2.0`
2. 用 PyInstaller **onefile** 模式重新打包（需把 `bin/es.exe`、`version.txt`、
   `使用说明.txt` 作为 `--add-data` 打进去），得到单个 `线上资料检查.exe`
3. 在 GitHub 上建 Release，tag 填 `v1.2.0`，**把这个 exe 作为附件上传**
   （文件名建议 `gas-doc-checker-v1.2.0-win64.exe`）
4. 客户端点「检查更新」即可自动下载新 exe 并替换自己

自动更新**只替换 exe 自己**；`检查报告/` 和用户放在目录里的其他文件一律不动。
为兼容 v1.0.x 老用户，Release 可同时附一份旧结构的 zip 整包（老版更新逻辑找 zip）。

## 用户反馈

程序右上角「反馈」按钮可匿名提交意见到本仓库 Issues（带 `user-feedback` 标签），
**自动附带当前工程编号和正在进行的操作**（检查/修正/更新、结果统计、界面状态、操作时间），
不用手动描述现场，便于复现问题。
提交用一枚**仅本仓库 Issues: write 权限**的 fine-grained token（`feedback.py` 里的 `TOKEN`），
请勿使用有 repo/workflow 权限的 token。反馈 Issue 标题统一 `[用户反馈] <类型> - <简述>`，
正文按「反馈类型/反馈内容/联系方式/环境信息」四段结构化，便于自动提取。

**AI 处理反馈**：接到「处理反馈」指令时，运行工作区的 `fetch_feedback.py`
（从 GitHub 拉取未处理的 `user-feedback` Issue 并结构化输出），逐条呈现给用户，
等用户决策是否修正、如何修正，未经确认不得擅自改代码。处理完的 Issue 加
`feedback-handled` 标签（脚本默认不再列出）。

## 自动更新原理

`updater.py` 只用 Python 标准库：

1. `GET https://api.github.com/repos/JsztNc/gas-doc-checker/releases/latest`
2. 比较 `tag_name` 与内嵌的 `version.txt`
3. 有新版则下载 Release 里的 `.exe` 资产到 `_tmp/update/`，校验 MZ 头
4. 在系统临时目录生成一个 bat：等主程序退出 → 用新 exe 覆盖旧 exe → 重启
5. 主程序关闭，由 bat 完成替换（Windows 下运行中的 exe 不能覆盖自己）
