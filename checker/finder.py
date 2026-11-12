# -*- coding: utf-8 -*-
"""用 Everything 的命令行版（es.exe）按工程编号找「<编号>-b.线上资料.docx」。"""
import csv
import os
import subprocess
import sys

if getattr(sys, "frozen", False):
    # 打包成 exe 后，es.exe 放在 exe 同级的 bin 目录
    BASE = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ES_CANDIDATES = [
    os.path.join(BASE, "bin", "es.exe"),
    # 打包后 --add-data 会把 bin 放进 _internal ，一并找
    os.path.join(getattr(sys, "_MEIPASS", BASE), "bin", "es.exe"),
    r"C:\Program Files\Everything\es.exe",
    r"C:\Program Files (x86)\Everything\es.exe",
    "es.exe",
]


def es_path():
    for p in ES_CANDIDATES:
        if p == "es.exe" or os.path.isfile(p):
            return p
    return None


def _run_export(query, out_csv):
    exe = es_path()
    if not exe:
        return None, "没找到 es.exe（Everything 命令行工具），请放到程序目录 bin\\ 下"
    if os.path.exists(out_csv):
        try:
            os.remove(out_csv)
        except OSError:
            pass
    cmd = [exe, "-csv", "-utf8-bom", "-dm", "-n", "200",
           "-export-csv", out_csv, query]
    try:
        subprocess.run(cmd, capture_output=True, timeout=40)
    except subprocess.TimeoutExpired:
        return None, "Everything 查询超时"
    except OSError as exc:
        return None, "调用 es.exe 失败：%s" % exc
    if not os.path.isfile(out_csv):
        return None, "Everything 没有返回结果文件（服务端是否在运行？）"
    rows = []
    with open(out_csv, "r", encoding="utf-8-sig", newline="") as fh:
        for row in csv.DictReader(fh):
            path = (row.get("Filename") or "").strip()
            if path:
                rows.append({"path": path, "mtime": (row.get("Date Modified") or "").strip()})
    return rows, None


def find(code, tmpdir=None):
    """返回 (候选列表, 错误信息)。候选按「修改时间新→路径短」排序。

    同时接受 .docx 与 .pdf 两种资料格式，文件名兼容
    「<编号>-b.线上资料」与「<编号>-线上资料」两种写法。
    """
    code = (code or "").strip().upper()
    if not code:
        return [], "工程编号为空"
    tmpdir = tmpdir or os.path.join(BASE, "_tmp")
    os.makedirs(tmpdir, exist_ok=True)
    out = os.path.join(tmpdir, "es_%s.csv" % code.replace("/", "_"))

    def _keep(r):
        name = os.path.basename(r["path"]).lower()
        ext = os.path.splitext(name)[1]
        return (ext in (".docx", ".pdf")
                and code.lower() in name
                and "线上资料" in name)

    # 主搜索：<编号> 线上资料（不限扩展名），再按文件名过滤
    rows, err = _run_export("%s 线上资料" % code, out)
    if err:
        return [], err
    hits = [r for r in rows if _keep(r)]
    if not hits:
        # 兜底：精确名（含 -b. 写法），兼容老命名 / 共享盘索引差异
        rows2, err2 = _run_export("%s-b.线上资料" % code, out)
        if not err2:
            hits = [r for r in rows2 if _keep(r)]

    if not hits:
        return [], ("没搜到「%s」相关的 线上资料.docx / .pdf。\n"
                    "可能原因：编号写错、Everything 没开着、或文件所在的共享盘没连上。"
                    % code)

    def rank(r):
        return (r["mtime"] or "", -len(r["path"]))

    hits.sort(key=rank, reverse=True)
    try:
        os.remove(out)
    except OSError:
        pass
    return hits, None


if __name__ == "__main__":
    hits, err = find(sys.argv[1])
    print("err:", err)
    for h in hits:
        print(h["mtime"], h["path"])
