# -*- coding: utf-8 -*-
"""检查更新模块 —— 从 GitHub Releases 拉取最新版并替换文件。

设计要点：
1. 只用标准库（urllib / zipfile / json），不额外装依赖，打包后也能跑。
2. 检查、下载在后台线程做，界面通过 BeginInvoke 回调，不卡窗口。
3. 运行中的 exe 无法覆盖自己，所以真正的替换交给一个放在系统临时目录的
   bat 脚本：先等主程序退出，再删旧文件、解压新文件、重新启动。
4. 只替换"程序自带"的文件；检查报告、_tmp、用户放在目录里的东西一律不动。

GitHub 上每个版本请发一个 Release，并上传一个 zip 包，zip 内容就是整个
程序文件夹（线上资料检查.exe、_internal、checker、*.py、版本文件等），
可以带一层顶层目录，程序会自动识别。tag 用 v1.1.0 这种格式。
"""
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import traceback
import urllib.error
import urllib.request
import zipfile

BASE = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False)
                                      else os.path.abspath(__file__)))

GITHUB_REPO = "lonyii/gas-doc-checker"
API_LATEST = "https://api.github.com/repos/%s/releases/latest" % GITHUB_REPO
RELEASES_PAGE = "https://github.com/%s/releases" % GITHUB_REPO
VERSION_FILE = os.path.join(BASE, "version.txt")
EXE_NAME = "线上资料检查.exe"

# 更新时会先删除这些"程序自带"项，再从新包整体解压回来。
# 故意不列：检查报告、_tmp、用户自己放进去的文件。
MANAGED_FILES = (
    "ui.py", "check.py", "updater.py", "version.txt", "使用说明.txt",
    "检查资料.bat", "命令行检查.bat", "_runtime.bat", EXE_NAME,
)
MANAGED_DIRS = ("_internal", "checker", "bin")


def base_dir():
    return BASE


def local_version():
    # 打包后 version.txt 可能只随资源打进 _internal（exe 同级那份用于更新时覆盖），
    # 两处都找一下，避免显示 0.0.0。
    candidates = [os.path.join(BASE, "version.txt")]
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(os.path.join(meipass, "version.txt"))
    for path in candidates:
        try:
            with open(path, "r", encoding="utf-8") as fh:
                text = fh.read().strip()
                if text:
                    return text
        except Exception:
            pass
    return "0.0.0"


def parse_version(text):
    nums = [int(x) for x in re.findall(r"\d+", text or "")]
    return tuple(nums) if nums else (0,)


def compare_version(a, b):
    """a > b 返回 1，相等 0，小于 -1。按数字段逐位比，长度不等补 0。"""
    va, vb = parse_version(a), parse_version(b)
    n = max(len(va), len(vb))
    va = va + (0,) * (n - len(va))
    vb = vb + (0,) * (n - len(vb))
    return (va > vb) - (va < vb)


def check_latest(timeout=12):
    """访问 GitHub API，返回 dict：
    ok, local, latest, has_update, notes, published_at, page,
    asset_name, asset_url, error
    """
    local = local_version()
    out = {"ok": False, "local": local, "latest": "", "has_update": False,
           "notes": "", "published_at": "", "page": RELEASES_PAGE,
           "asset_name": "", "asset_url": "", "error": ""}
    try:
        req = urllib.request.Request(API_LATEST, headers={
            "User-Agent": "gas-doc-checker/%s" % local,
            "Accept": "application/vnd.github+json",
        })
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8", "replace"))
        latest = (data.get("tag_name") or "").strip()
        out["latest"] = latest.lstrip("vV")
        out["notes"] = data.get("body") or ""
        out["published_at"] = (data.get("published_at") or "")[:10]
        out["page"] = data.get("html_url") or RELEASES_PAGE
        assets = [a for a in (data.get("assets") or [])
                  if (a.get("name") or "").lower().endswith(".zip")]
        # 优先带 win/windows 字样的整包，其次第一个 zip
        assets.sort(key=lambda a: (0 if re.search(r"win", a["name"], re.I) else 1))
        if assets:
            out["asset_name"] = assets[0].get("name") or ""
            out["asset_url"] = assets[0].get("browser_download_url") or ""
        out["has_update"] = bool(out["latest"]) and compare_version(out["latest"], local) > 0
        out["ok"] = True
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            out["error"] = "仓库还没有发布任何版本"
        elif exc.code == 403:
            out["error"] = "GitHub 接口访问次数已达上限，请稍后再试"
        else:
            out["error"] = "GitHub 返回错误：HTTP %s" % exc.code
    except Exception as exc:
        out["error"] = "连不上 GitHub：%s" % exc
    return out


def download_release(url, on_progress=None, timeout=180):
    """把 Release 的 zip 下载到 _tmp/update/，返回本地路径。"""
    if not url:
        raise RuntimeError("这个版本没有提供 zip 整包，只能去网页手动下载")
    dst_dir = os.path.join(BASE, "_tmp", "update")
    os.makedirs(dst_dir, exist_ok=True)
    dst = os.path.join(dst_dir, "release.zip")
    req = urllib.request.Request(url, headers={"User-Agent": "gas-doc-checker-updater"})
    with urllib.request.urlopen(req, timeout=timeout) as resp, open(dst, "wb") as fh:
        total = 0
        try:
            total = int(resp.headers.get("Content-Length") or 0)
        except Exception:
            pass
        done = 0
        last_report = 0.0
        while True:
            chunk = resp.read(256 * 1024)
            if not chunk:
                break
            fh.write(chunk)
            done += len(chunk)
            now = time.time()
            if on_progress and now - last_report > 0.2:
                last_report = now
                on_progress(done, total)
        if on_progress:
            on_progress(done, total)
    return dst


def _zip_root(names):
    """如果 zip 里所有内容都在同一个顶层目录下，返回该目录名，否则返回 ''。"""
    firsts = set()
    for name in names:
        name = name.replace("\\", "/").lstrip("/")
        if name:
            firsts.add(name.split("/")[0])
    if len(firsts) == 1:
        only = next(iter(firsts))
        for name in names:
            if name.replace("\\", "/").lstrip("/").split("/")[0] != only:
                return ""
        return only if any("/" in n for n in names if n) else ""
    return ""


def prepare_update(zip_path, tag):
    """解压 zip 并生成替换脚本，返回 (bat路径, 说明字符串)。
    校验包里必须有可执行文件或源码入口，避免拿错包把程序弄坏。
    """
    work = os.path.join(BASE, "_tmp", "update")
    stage = os.path.join(work, "new")
    if os.path.isdir(stage):
        shutil.rmtree(stage, ignore_errors=True)
    os.makedirs(stage, exist_ok=True)
    with zipfile.ZipFile(zip_path) as zf:
        names = zf.namelist()
        root = _zip_root(names)
        zf.extractall(stage)
    src = os.path.join(stage, root) if root else stage
    if not (os.path.isfile(os.path.join(src, EXE_NAME))
            or (os.path.isfile(os.path.join(src, "ui.py"))
                and os.path.isfile(os.path.join(src, "check.py")))):
        raise RuntimeError("这个 zip 不像程序整包（里面没有 %s 或 ui.py+check.py），已取消更新" % EXE_NAME)

    pid = os.getpid()
    frozen = bool(getattr(sys, "frozen", False))
    pyexe = sys.executable if not frozen else ""

    # bat 放到系统临时目录，避免它自己所在的文件夹被改名/删除
    fd, bat = tempfile.mkstemp(prefix="gasdoc_update_", suffix=".bat")
    os.close(fd)
    q = lambda s: '"%s"' % s.replace('"', '""')
    lines = [
        "@echo off",
        "chcp 65001 >nul",
        "title 正在更新程序…",
        "echo 正在更新线上资料检查工具，请勿关闭此窗口。",
        ":wait",
        'tasklist /FI "PID eq %d" 2>NUL | findstr /I /C:"%d" >NUL' % (pid, pid),
        "if not errorlevel 1 (ping -n 2 127.0.0.1 >NUL & goto wait)",
        "ping -n 2 127.0.0.1 >NUL",
    ]
    for name in MANAGED_DIRS:
        lines.append('if exist %s rmdir /S /Q %s' % (q(os.path.join(BASE, name)),
                                                     q(os.path.join(BASE, name))))
    for name in MANAGED_FILES:
        lines.append('if exist %s del /F /Q %s' % (q(os.path.join(BASE, name)),
                                                   q(os.path.join(BASE, name))))
    lines += [
        'xcopy /E /Y /R /H /I %s %s >NUL' % (q(os.path.join(src, "*")), q(BASE + os.sep)),
    ]
    if frozen:
        lines.append('if exist %s start "" %s' % (q(os.path.join(BASE, EXE_NAME)),
                                                  q(os.path.join(BASE, EXE_NAME))))
    elif pyexe:
        lines.append('start "" %s %s' % (q(pyexe), q(os.path.join(BASE, "ui.py"))))
    lines += [
        "echo 更新完成，窗口将在几秒后关闭。",
        "ping -n 4 127.0.0.1 >NUL",
        'del /F /Q "%~f0" >NUL 2>&1',
    ]
    with open(bat, "w", encoding="utf-8") as fh:
        fh.write("\r\n".join(lines) + "\r\n")
    return bat, "已下载 v%s，程序即将关闭并完成替换" % tag


def start_update(bat_path):
    """启动替换脚本（分离进程），随后主程序应自行退出。"""
    subprocess.Popen(["cmd", "/c", "start", "", bat_path], shell=False,
                     cwd=tempfile.gettempdir(),
                     creationflags=getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0))


def open_releases():
    try:
        os.startfile(RELEASES_PAGE)          # noqa: S606
    except Exception:
        pass
