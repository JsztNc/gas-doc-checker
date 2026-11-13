# -*- coding: utf-8 -*-
"""检查更新模块 —— V1.4.2 起为多文件增量更新。

设计要点：
1. 只用标准库（urllib / json / zipfile / hashlib），打包后也能跑。
2. 程序是多文件版（onedir）：exe + _internal + checker/（审核规则外挂）等，
   整个目录打成 zip 发布到 GitHub Release。
3. 更新流程 = 下载 zip → 解压 → 逐文件与本地比对 md5 → 只覆盖有变动的文件：
   - 未被占用的文件（checker/*.py、使用说明等）当场替换，改前备份到
     `_改前备份/更新-<版本>/`；
   - 被占用的文件（exe 自身、_internal 里的 dll）记入待补拷清单，由替换脚本
     在主程序退出后补拷，然后重新打开程序。
4. 兼容旧版：老版本（v1.4.1 及之前）是 onefile 单 exe，其更新器只认 .exe 资产；
   因此 Release 同时附一个过渡用单文件 exe，老用户自更新拿到它后，即具备
   增量更新能力，之后每次只换有变动的文件。

发版：tag 用 v1.4.2，上传两个资产：
    gas-doc-checker-v1.4.2.zip   多文件整包（新格式，增量更新的数据源）
    gas-doc-checker-v1.4.2.exe   过渡用单文件版（checker 编译在内）
"""
import hashlib
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
    """优先用打包资源目录（onefile 解压出的 _MEIPASS / onedir 的 _internal，
    反映程序内嵌真实版本），再回退到 exe 同级目录。"""
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
    """返回使用说明全文（供界面「使用说明」窗口显示）。"""
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
    """从 release 数据里挑可更新的资产：V1.4.2 起优先 zip（多文件增量包），
    没有再退回 exe（过渡单文件 / 更老的发版）。"""
    assets = data.get("assets") or []
    for name, url in _iter_assets(assets, ".zip"):
        return name, url, "zip"
    for name, url in _iter_assets(assets, ".exe"):
        return name, url, "exe"
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


# ---------------------------------------------------------------- 增量更新
def _md5(path, chunk=1 << 20):
    h = hashlib.md5()
    with open(path, "rb") as fh:
        for b in iter(lambda: fh.read(chunk), b""):
            h.update(b)
    return h.hexdigest()


def _extract_zip(zip_path, stage):
    """解压 zip 到 stage，返回内容根目录（兼容 zip 里带一层同名目录）。"""
    if os.path.isdir(stage):
        shutil.rmtree(stage, ignore_errors=True)
    os.makedirs(stage, exist_ok=True)
    with zipfile.ZipFile(zip_path) as zf:
        names = [n.replace("\\", "/").lstrip("/") for n in zf.namelist() if n]
        firsts = set(n.split("/")[0] for n in names)
        root = ""
        if len(firsts) == 1 and any("/" in n for n in names):
            root = next(iter(firsts))
        zf.extractall(stage)
    return os.path.join(stage, root) if root else stage


def _changed_files(src):
    """src 里每个文件与 BASE 同路径比对，返回有变动（或本地没有）的相对路径列表。"""
    out = []
    for root, dirs, files in os.walk(src):
        dirs[:] = [d for d in dirs if d not in ("_tmp", "检查报告", "_改前备份")]
        for name in files:
            sp = os.path.join(root, name)
            rel = os.path.relpath(sp, src)
            tp = os.path.join(BASE, rel)
            try:
                if os.path.isfile(tp) and _md5(sp) == _md5(tp):
                    continue
            except Exception:
                pass
            out.append(rel)
    return out


def _backup(rel, tag):
    """把要被覆盖的现有文件备份到 _改前备份/更新-<tag>/<rel>。"""
    tp = os.path.join(BASE, rel)
    if not os.path.isfile(tp):
        return
    bdir = os.path.join(BASE, "_改前备份", "更新-" + tag)
    dst = os.path.join(bdir, rel)
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    if not os.path.exists(dst):
        shutil.copy2(tp, dst)


def _apply_now(src, changed, tag):
    """当场复制未锁定的变动文件，返回 (已拷, 待补拷) 两个相对路径列表。"""
    copied, deferred = [], []
    for rel in changed:
        sp = os.path.join(src, rel)
        tp = os.path.join(BASE, rel)
        if os.path.isfile(tp):
            _backup(rel, tag)
        try:
            os.makedirs(os.path.dirname(tp), exist_ok=True)
            shutil.copy2(sp, tp)
            copied.append(rel)
        except PermissionError:
            deferred.append(rel)
        except OSError:
            deferred.append(rel)
    return copied, deferred


def _write_pending_bat(deferred, src, tag, pid):
    """生成替换脚本：等主程序退出 → 按清单补拷被锁定的文件 → 重启程序。"""
    fd, bat = tempfile.mkstemp(prefix="gasdoc_update_", suffix=".bat")
    os.close(fd)
    lst = os.path.join(BASE, "_tmp", "update", "pending_files.txt")
    os.makedirs(os.path.dirname(lst), exist_ok=True)
    with open(lst, "w", encoding="gbk", errors="replace") as fh:
        for rel in deferred:
            fh.write("%s\t%s\r\n" % (os.path.join(src, rel),
                                     os.path.join(BASE, rel)))
    q = lambda s: '"%s"' % s.replace('"', '""')
    lines = [
        "@echo off",
        "chcp 936 >nul",
        "title 正在更新程序…",
        "echo 正在更新线上资料检查工具（仅替换有变动的文件），请勿关闭此窗口。",
        ":wait",
        'tasklist /FI "PID eq %d" 2>NUL | findstr /I /C:"%d" >NUL' % (pid, pid),
        "if not errorlevel 1 (ping -n 2 127.0.0.1 >NUL & goto wait)",
        "ping -n 2 127.0.0.1 >NUL",
        "set RETRY=0",
        ":copy",
        'for /f "usebackq tokens=1,2 delims=	" %%A in (%s) do (' % q(lst),
        '  if not exist "%%~dpB" mkdir "%%~dpB"',
        '  copy /Y /B "%%A" "%%B" >NUL',
        ")",
        'copy /Y /B %s %s >NUL' % (q(os.path.join(src, EXE_NAME)),
                                   q(os.path.join(BASE, EXE_NAME))),
        "if not errorlevel 1 goto copied",
        "set /a RETRY+=1",
        "if %RETRY% lss 15 (ping -n 1 127.0.0.1 >NUL & goto copy)",
        "goto fail",
        ":copied",
        'start "" %s' % q(os.path.join(BASE, EXE_NAME)),
        'rd /s /q %s >NUL 2>&1' % q(src),
        "echo 更新完成，窗口将在几秒后关闭。",
        "ping -n 4 127.0.0.1 >NUL",
        'del /F /Q "%~f0" >NUL 2>&1',
        "goto :eof",
        ":fail",
        "echo 替换失败：部分文件仍被占用，请稍后重试更新或手动覆盖。",
        "pause",
    ]
    with open(bat, "w", encoding="gbk", errors="replace") as fh:
        fh.write("\r\n".join(lines) + "\r\n")
    return bat


def prepare_update(asset_path, tag):
    """校验下载的更新包，执行增量替换，返回 (bat路径, 说明字符串)。

    zip（V1.4.2 起的标准形态）：解压后逐文件比对，只覆盖有变动的文件；
      未锁定的当场替换（先备份），被锁定的（exe、加载中的 dll）写进
      待补拷清单，由替换脚本在程序退出后补拷并重启程序。
    exe（过渡单文件版自更新）：与旧版行为一致，只替换 exe 自己。
    """
    work = os.path.join(BASE, "_tmp", "update")
    if asset_path.lower().endswith(".zip"):
        src = _extract_zip(asset_path, os.path.join(work, "new"))
        if not os.path.isfile(os.path.join(src, EXE_NAME)):
            raise RuntimeError("这个 zip 里没有 %s，已取消更新" % EXE_NAME)
        changed = _changed_files(src)
        if not changed:
            raise RuntimeError("与当前版本内容一致，没有需要更新的文件")
        copied, deferred = _apply_now(src, changed, tag)
        bat = _write_pending_bat(deferred, src, tag, os.getpid())
        msg = ("已下载 v%s：共 %d 个文件有变动，其中 %d 个已当场替换%s"
               % (tag, len(changed), len(copied),
                  "，%d 个（程序本体等）将在退出后替换" % len(deferred) if deferred else ""))
        return bat, msg

    # ---- 兼容：exe 资产（过渡单文件版 / 旧版） ----
    new_exe = asset_path
    with open(new_exe, "rb") as fh:
        if fh.read(2) != b"MZ":
            raise RuntimeError("下载的文件不是有效的 exe，已取消更新")
    work2 = os.path.join(BASE, "_tmp", "update")
    os.makedirs(work2, exist_ok=True)
    src = os.path.join(work2, "exe_stage")
    if os.path.isdir(src):
        shutil.rmtree(src, ignore_errors=True)
    os.makedirs(src, exist_ok=True)
    staged_exe = os.path.join(src, EXE_NAME)
    shutil.copy2(new_exe, staged_exe)
    target = os.path.join(BASE, EXE_NAME)
    if os.path.isfile(target):
        _backup(EXE_NAME, tag)
    bat = _write_pending_bat([EXE_NAME], src, tag, os.getpid())
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
