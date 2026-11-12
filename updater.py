# -*- coding: utf-8 -*-
"""检查更新模块 —— 单文件版：从 GitHub Releases 下载最新 exe 并替换自己。

设计要点：
1. 只用标准库（urllib / json），不额外装依赖，打包后也能跑。
2. 程序是 PyInstaller --onefile 打出的单个 exe，运行时解压到临时目录；
   version.txt、使用说明.txt、bin/es.exe 都作为资源打在 exe 内部。
3. Release 只上传一个 .exe 资产（如 gas-doc-checker-v1.1.0-win64.exe）。
   检查更新时下载这个 exe，校验是有效 PE 文件后，用系统临时目录的 bat
   等主程序退出 → 用新 exe 覆盖旧 exe → 重新打开。
4. 只替换 exe 自己；检查报告、_tmp、用户放在目录里的东西一律不动。

发版：tag 用 v1.1.0，把编好的 exe 作为 Release 资产上传即可。
老版本 v1.0.x 是 zip 整包结构，本模块也兼容读取 zip（见 _asset_from_release）。
"""
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.request
import zipfile

BASE = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False)
                                      else os.path.abspath(__file__)))

GITHUB_REPO = "JsztNc/gas-doc-checker"
API_LATEST = "https://api.github.com/repos/%s/releases/latest" % GITHUB_REPO
RELEASES_PAGE = "https://github.com/%s/releases" % GITHUB_REPO
EXE_NAME = "线上资料检查.exe"
HELP_FILE = "使用说明.txt"


def base_dir():
    return BASE


def _resource_path(name):
    """优先用打包资源目录（onefile 解压出的 _MEIPASS，反映 exe 内嵌真实版本），
    再回退到 exe 同级目录。这样版本号不会被源码目录里那份 version.txt 掩盖。"""
    for d in (getattr(sys, "_MEIPASS", None), BASE):
        if not d:
            continue
        p = os.path.join(d, name)
        if os.path.isfile(p):
            return p
    return os.path.join(BASE, name)


def local_version():
    try:
        with open(_resource_path("version.txt"), "r", encoding="utf-8") as fh:
            return fh.read().strip() or "0.0.0"
    except Exception:
        return "0.0.0"


def help_text():
    """返回内嵌的使用说明全文（供界面「使用说明」窗口显示）。"""
    try:
        with open(_resource_path(HELP_FILE), "r", encoding="utf-8") as fh:
            return fh.read()
    except Exception as exc:
        return "（使用说明读取失败：%s）" % exc


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


def _asset_from_release(data):
    """从 release 数据里挑可更新的资产：优先 .exe，退回 .zip（兼容旧版）。"""
    assets = data.get("assets") or []
    for name, url in _iter_assets(assets, ".exe"):
        return name, url, "exe"
    for name, url in _iter_assets(assets, ".zip"):
        return name, url, "zip"
    return "", "", ""


def _iter_assets(assets, suffix):
    hits = [(a.get("name") or "", a.get("browser_download_url") or "")
            for a in assets
            if (a.get("name") or "").lower().endswith(suffix)]
    hits.sort(key=lambda x: (0 if re.search(r"win64|win", x[0], re.I) else 1))
    return hits


def check_latest(timeout=12):
    """访问 GitHub API，返回 dict：
    ok, local, latest, has_update, notes, published_at, page,
    asset_name, asset_url, asset_kind, error
    """
    local = local_version()
    out = {"ok": False, "local": local, "latest": "", "has_update": False,
           "notes": "", "published_at": "", "page": RELEASES_PAGE,
           "asset_name": "", "asset_url": "", "asset_kind": "", "error": ""}
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
        name, url, kind = _asset_from_release(data)
        out["asset_name"], out["asset_url"], out["asset_kind"] = name, url, kind
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


def download_release(url, on_progress=None, timeout=300):
    """下载资产到 _tmp/update/，文件名按扩展名命名，返回本地路径。"""
    if not url:
        raise RuntimeError("这个版本没有提供可下载的更新文件，只能去网页手动下载")
    ext = ".exe" if url.lower().split("?")[0].endswith(".exe") else ".zip"
    dst_dir = os.path.join(BASE, "_tmp", "update")
    os.makedirs(dst_dir, exist_ok=True)
    dst = os.path.join(dst_dir, "release" + ext)
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


def _find_exe_in_zip(zip_path, stage):
    """兼容旧版 zip 整包：解压并找出里面的 exe（或 ui.py+check.py）。"""
    with zipfile.ZipFile(zip_path) as zf:
        names = zf.namelist()
        firsts = set(n.replace("\\", "/").lstrip("/").split("/")[0] for n in names if n)
        root = ""
        if len(firsts) == 1:
            only = next(iter(firsts))
            if all(n.replace("\\", "/").lstrip("/").split("/")[0] == only for n in names):
                root = only if any("/" in n for n in names if n) else ""
        zf.extractall(stage)
    src = os.path.join(stage, root) if root else stage
    exe = os.path.join(src, EXE_NAME)
    if os.path.isfile(exe):
        return exe
    return None


def prepare_update(asset_path, tag):
    """校验下载的文件，生成替换脚本，返回 (bat路径, 说明字符串)。

    单文件版只做一件事：用新 exe 覆盖旧 exe。
    兼容旧 zip：从 zip 里取出 exe 再替换。
    """
    if asset_path.lower().endswith(".zip"):
        work = os.path.join(BASE, "_tmp", "update")
        stage = os.path.join(work, "new")
        if os.path.isdir(stage):
            shutil.rmtree(stage, ignore_errors=True)
        os.makedirs(stage, exist_ok=True)
        new_exe = _find_exe_in_zip(asset_path, stage)
        if not new_exe:
            raise RuntimeError("这个 zip 里没有 %s，已取消更新" % EXE_NAME)
    else:
        new_exe = asset_path
        with open(new_exe, "rb") as fh:
            if fh.read(2) != b"MZ":
                raise RuntimeError("下载的文件不是有效的 exe，已取消更新")

    target = os.path.join(BASE, EXE_NAME) if getattr(sys, "frozen", False) else sys.executable
    if not os.path.isfile(target):
        raise RuntimeError("找不到当前程序 %s，已取消更新" % target)

    pid = os.getpid()
    # bat 放到系统临时目录，避免被更新流程影响
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
        "set RETRY=0",
        ":copy",
        'copy /Y /B %s %s >NUL' % (q(new_exe), q(target)),
        "if not errorlevel 1 goto copied",
        "set /a RETRY+=1",
        "if %RETRY% lss 15 (ping -n 1 127.0.0.1 >NUL & goto copy)",
        "goto fail",
        ":copied",
        'start "" %s' % q(target),
        "echo 更新完成，窗口将在几秒后关闭。",
        "ping -n 4 127.0.0.1 >NUL",
        'del /F /Q "%~f0" >NUL 2>&1',
        "goto :eof",
        ":fail",
        "echo 替换失败：旧程序可能被占用，请手动把新 exe 拷到程序目录。",
        "pause",
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
