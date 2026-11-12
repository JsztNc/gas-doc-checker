# -*- coding: utf-8 -*-
"""按工程编号自动找「<编号>-线上资料.docx / .pdf」并对照规范检查。

用法：
    python check.py NC23A005
    python check.py ID25E021 --path "D:\\某项目\\ID25E021-线上资料.docx"
    python check.py IE25E001 --path "D:\\某项目\\IE25E001-线上资料.pdf"
    python check.py IE25E001 --html          # 另存一份 html 报告
"""
import argparse
import datetime
import io
import os
import sys

if getattr(sys, "frozen", False):
    BASE = os.path.dirname(os.path.abspath(sys.executable))   # 打包后：报告写在 exe 旁边
else:
    BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)

from checker.docx_reader import Doc                     # noqa: E402
from checker.finder import find                         # noqa: E402
from checker.rules import ERR, INFO, WARN, check        # noqa: E402

REPORT_DIR = os.path.join(BASE, "检查报告")


def summarize(f, findings):
    n = {ERR: 0, WARN: 0, INFO: 0}
    for x in findings:
        n[x.level] = n.get(x.level, 0) + 1
    return n


def to_markdown(code, path, f, findings):
    n = summarize(f, findings)
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    lines = []
    lines.append("# %s 线上资料检查报告" % code)
    lines.append("")
    lines.append("- 文件：`%s`" % path)
    lines.append("- 项目：%s" % (f.meta.get("项目名称") or "（未识别）"))
    lines.append("- 编号：%s　施工单位：%s" % (f.meta.get("项目编号") or "-",
                                      f.meta.get("施工单位") or "-"))
    lines.append("- 检查时间：%s" % now)
    lines.append("- 结论：**错误 %d 条，警告 %d 条，提示 %d 条**" % (n[ERR], n[WARN], n[INFO]))
    lines.append("")
    if not findings:
        lines.append("没有发现明显问题。")
        return "\n".join(lines)
    cur = None
    for lv in (ERR, WARN, INFO):
        group = [x for x in findings if x.level == lv]
        if not group:
            continue
        lines.append("## %s（%d 条）" % (lv, len(group)))
        lines.append("")
        for x in group:
            lines.append("- **%s**：%s" % (x.rule, x.msg))
        lines.append("")
    lines.append("---")
    lines.append("说明：本报告由《线上燃气竣工资料填写规范与校验要点》自动比对生成，"
                 "只核能机器判断的项目（编号、时间、次数、米数、证件有效期等）；"
                 "照片、签字、印章等需人工复核。")
    return "\n".join(lines)


def to_html(md_text, code):
    import html
    body = []
    for line in md_text.split("\n"):
        t = line.rstrip()
        if t.startswith("## "):
            body.append("<h2>%s</h2>" % html.escape(t[3:]))
        elif t.startswith("# "):
            body.append("<h1>%s</h1>" % html.escape(t[2:]))
        elif t.startswith("- "):
            txt = html.escape(t[2:]).replace("**", "")
            body.append("<li>%s</li>" % txt)
        elif not t:
            body.append("")
        else:
            body.append("<p>%s</p>" % html.escape(t).replace("**", ""))
    css = ("body{font-family:'Microsoft YaHei',sans-serif;max-width:900px;margin:32px auto;"
           "line-height:1.75;color:#222}h1{font-size:22px}h2{font-size:17px;margin-top:22px;"
           "border-left:4px solid #1777ff;padding-left:8px}li{margin:4px 0}")
    return ("<!doctype html><meta charset='utf-8'><title>%s 检查报告</title>"
            "<style>%s</style>%s" % (html.escape(code), css, "\n".join(body)))


def run(code, path=None, prefer=0, html_out=False, quiet=False, json_out=None,
        do_fix=False):
    """返回 (doc, findings, 报告路径)；出错抛 RuntimeError。"""
    tried = []
    if not path:
        hits, err = find(code)
        if err:
            raise RuntimeError(err)
        if not hits:
            raise RuntimeError(
                "没搜到「%s」相关的 线上资料.docx / .pdf。\n"
                "可能原因：编号写错、Everything 没开着、或文件所在的共享盘没连上。"
                % code)
        order = list(range(prefer, len(hits))) + list(range(0, prefer))
        path = None
        for i in order:                       # 同名副本可能有打不开的，逐个试
            cand = hits[i]["path"]
            tried.append(cand)
            if os.path.isfile(cand):
                path = cand
                break
        if not path:
            raise RuntimeError("搜到的 %d 份副本都打不开：\n%s"
                               % (len(hits), "\n".join(tried[:4])))
    if not os.path.isfile(path):
        raise RuntimeError("文件打不开：%s" % path)
    fix_info = None
    if do_fix:                                  # 先把能自动改的改掉，再按改后结果出报告
        from checker.fixer import fix_document
        try:
            fix_info = fix_document(path)
        except PermissionError:
            raise RuntimeError("文件正被 Word/WPS 或预览占着，关掉后再试：%s" % path)
    doc = Doc(path)
    f, findings = check(doc)
    n = summarize(f, findings)

    os.makedirs(REPORT_DIR, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%m%d-%H%M")
    base = os.path.join(REPORT_DIR, "%s-%s" % (code, stamp))
    md = to_markdown(code, path, f, findings)
    with io.open(base + ".md", "w", encoding="utf-8") as fh:
        fh.write(md)
    out = base + ".md"
    if html_out:
        with io.open(base + ".html", "w", encoding="utf-8") as fh:
            fh.write(to_html(md, code))
        out = base + ".html"

    if json_out:
        import json
        payload = {"ok": True, "code": code, "path": path, "meta": f.meta,
                   "counts": n, "report": out, "fix": fix_info,
                   "findings": [{"level": x.level, "rule": x.rule, "msg": x.msg} for x in findings]}
        with io.open(json_out, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, ensure_ascii=False, indent=1)

    if fix_info and not quiet:
        print("已自动修正：编号 %d 处、补人员 %d 人（%s）"
              % (fix_info["codes"], fix_info["members"],
                 "、".join(p[0] for p in fix_info["people"]) or "无"))
        print("改前备份：%s" % fix_info.get("backup"))
    if not quiet:
        print("文件：%s" % path)
        print("项目：%s" % (f.meta.get("项目名称") or "-"))
        print("结果：错误 %d 条，警告 %d 条，提示 %d 条" % (n[ERR], n[WARN], n[INFO]))
        print("")
        for lv in (ERR, WARN, INFO):
            group = [x for x in findings if x.level == lv]
            if not group:
                continue
            print("—— %s ——" % lv)
            for x in group:
                print("  · [%s] %s" % (x.rule, x.msg))
            print("")
        print("报告已存：%s" % out)
    return doc, findings, out


def main():
    ap = argparse.ArgumentParser(description="线上燃气竣工资料自动检查")
    ap.add_argument("code", help="工程编号，如 NC23A005")
    ap.add_argument("--path", help="直接指定 docx/pdf 路径，跳过 Everything 搜索")
    ap.add_argument("--pick", type=int, default=0, help="多个同名文件时取第几个（从 0 开始，按修改时间新→旧）")
    ap.add_argument("--html", action="store_true", help="额外输出 html 版报告")
    ap.add_argument("--json", dest="json_out", help="把结果写成 json（界面调用用）")
    ap.add_argument("--fix", dest="do_fix", action="store_true",
                    help="按规范自动修正（去空格、重排编号、补人员），改前自动备份")
    args = ap.parse_args()
    try:
        run(args.code, args.path, args.pick, args.html,
            quiet=bool(args.json_out), json_out=args.json_out, do_fix=args.do_fix)
    except RuntimeError as exc:
        if args.json_out:
            import json
            with io.open(args.json_out, "w", encoding="utf-8") as fh:
                json.dump({"ok": False, "error": str(exc)}, fh, ensure_ascii=False)
        else:
            print("出错：%s" % exc)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
