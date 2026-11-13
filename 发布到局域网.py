# -*- coding: utf-8 -*-
"""发布到局域网共享目录（客户端「检查更新」优先走这条路）。

用法（在仓库根目录）：
    python 发布到局域网.py --src dist_v144\\stage --version 1.4.4
    python 发布到局域网.py --src dist_v144\\stage --version 1.4.4 --manifest-only

做三件事：
1. 生成 files.json（相对路径 → 大小/md5，客户端只读它做比对，避免逐文件读共享）
2. 把整个程序目录拷到共享（--manifest-only 时跳过拷贝，只补清单/版本标记）
3. 写 version.txt（无 BOM）与可选的 更新说明.txt

前提：当前 Windows 会话已能访问共享（资源管理器打开过或 net use 过）。
"""
import argparse
import json
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import updater  # noqa: E402


def write_text(path, text):
    """无 BOM 写 UTF-8，换行统一 CRLF。"""
    with open(path, "w", encoding="utf-8", newline="\r\n") as fh:
        fh.write(text)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", required=True, help="待发布的程序目录（dist_vXXX\\stage）")
    ap.add_argument("--version", required=True, help="版本号，如 1.4.4")
    ap.add_argument("--dest", default=updater.LAN_ROOT, help="目标共享目录")
    ap.add_argument("--note", default="", help="更新说明文字（可不填）")
    ap.add_argument("--manifest-only", action="store_true",
                    help="只写清单/版本标记，不拷贝程序文件")
    args = ap.parse_args()

    src = os.path.abspath(args.src)
    if not os.path.isdir(src):
        raise SystemExit("源目录不存在：%s" % src)
    if not os.path.isfile(os.path.join(src, updater.EXE_NAME)):
        raise SystemExit("源目录里没有 %s，确认路径是否指向 stage 目录" % updater.EXE_NAME)
    if not os.path.isdir(args.dest):
        raise SystemExit("目标共享目录不可访问：%s" % args.dest)

    if not args.manifest_only:
        print("拷贝 %s -> %s …" % (src, args.dest))
        shutil.copytree(src, args.dest, dirs_exist_ok=True)
        print("拷贝完成")

    man = updater.make_manifest(src)
    man["version"] = args.version
    with open(os.path.join(args.dest, updater.LAN_MANIFEST), "w",
              encoding="utf-8", newline="\n") as fh:
        json.dump(man, fh, ensure_ascii=False, indent=1)
    write_text(os.path.join(args.dest, updater.LAN_VERSION_FILE), args.version + "\n")
    if args.note:
        write_text(os.path.join(args.dest, "更新说明.txt"), args.note)
    print("已发布 v%s：%d 个文件，清单已更新" % (args.version, len(man["files"])))


if __name__ == "__main__":
    main()
