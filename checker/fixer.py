# -*- coding: utf-8 -*-
"""按规范重排工序单据编号。

背景（NC26E365 实测）：钢制管道焊接记录里每行是一张单据，单据编号占了
GX0006~GX0035；其余工序单据的编号混在这段号里，造成重复、顺序颠倒。
处理方式：
  · 焊接记录的 21 个单据号保持不变，只去掉编号里的空格（系统会把
    「NC26E365G X0019」当成另一个号）；
  · 其余工序单据按工序先后排到 0036 之后连续编号：
    螺纹连接 → 管道防腐 → 隐蔽工程 → 管道吹扫 → 强度试验 → 严密性试验 → 附件安装
  · 所有编号统一成「项目编号GX0000」格式。

只改「单据编号」格子里匹配编号格式的文本，其余内容一概不动；改前必须备份。
"""
import re

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table

from .docx_reader import classify

# 表格里像编号的文本：允许任意空格混在中间（线上系统导出常见）
CODE_CELL = re.compile(r"^([A-Z]{2}\d{2}[A-Z0-9]{1,8}?)GX0*(\d+)$", re.I)
CODE_FLAT = re.compile(r"^[A-Z]{2}\d{2}[A-Z0-9]{1,8}?GX0*\d+$", re.I)

# 工序顺序 → 起始流水号（排在焊接号段之后）
ORDER = [("luowen", 36), ("fangfu", 37), ("yinbi", 38), ("chuixao", 39),
         ("qiangdu", 40), ("mishi", 41), ("fujian", 42)]


def _norm(s):
    return re.sub(r"\s+", "", s or "")


def _table_kind(tbl):
    """按表头签名判单据类型；强度/严密按压力单位区分。"""
    rows = []
    for row in tbl.rows[:6]:
        rows.append([c.text for c in row.cells])
    kind = classify(rows)
    if kind == "shiyan":
        # 强度用 Mpa，严密用 Kpa，靠单位区分
        blob = _norm("".join(c.text for row in tbl.rows[:4] for c in row.cells)).upper()
        return "qiangdu" if ("MPA" in blob and "KPA" not in blob) else "mishi"
    return kind


def _code_in_cell(cell):
    """返回格子里编号的 (pid, 流水号)；不是编号格返回 None。"""
    for p in cell.paragraphs:
        m = CODE_CELL.match(_norm(p.text))
        if m:
            return m.group(1).upper(), int(m.group(2))
    return None


def _set_cell_code(cell, new_text):
    done = False
    for p in cell.paragraphs:
        if not CODE_CELL.match(_norm(p.text)):
            continue
        _set_text(cell, new_text)
        done = True
        break
    return done


def analyze(path):
    """dry-run：返回 [(位置说明, 旧号, 新号)]，不写文件。"""
    doc = Document(path)
    plan = []
    start = {k: n for k, n in ORDER}
    for ti, tbl in enumerate(doc.tables):
        kind = _table_kind(tbl)
        fixed = None
        if kind == "ganghan":                    # 焊接：号不变，只去空格
            fixed = "keep"
        elif kind in start:
            fixed = start[kind]
        if fixed is None:
            continue
        for ri, row in enumerate(tbl.rows):
            for ci, cell in enumerate(row.cells):
                hit = _code_in_cell(cell)
                if not hit:
                    continue
                pid, seq = hit
                if fixed == "keep":
                    new = "%sGX%04d" % (pid, seq)
                else:
                    new = "%sGX%04d" % (pid, fixed)
                old = cell.text.strip()
                if old != new:
                    plan.append(("表%d %s R%dC%d" % (ti + 1, kind, ri + 1, ci + 1),
                                 old, new))
                break                                 # 一行只改第一个编号格
    return plan


def apply(path, out_path=None):
    """按 analyze 的方案改写。out_path 缺省就地保存。"""
    doc = Document(path)
    start = {k: n for k, n in ORDER}
    changed = 0
    for tbl in doc.tables:
        kind = _table_kind(tbl)
        fixed = "keep" if kind == "ganghan" else start.get(kind)
        if fixed is None:
            continue
        for row in tbl.rows:
            for cell in row.cells:
                hit = _code_in_cell(cell)
                if not hit:
                    continue
                pid, seq = hit
                new = "%sGX%04d" % (pid, seq if fixed == "keep" else fixed)
                if _set_cell_code(cell, new):
                    changed += 1
                break
    doc.save(out_path or path)
    return changed


# --------------------------------------------------------------------------
# 补人员：特种工登记表里有的人，项目人员信息表里必须有
# --------------------------------------------------------------------------
from copy import deepcopy

from docx.table import _Row


def _set_text(cell, text):
    """整体重写这一格：清掉所有 run（含超链接/域里的），再按原格式写一遍。

    线上导出的单元格常把「…有限公司」「负责人」拆成好几个 run，
    只改 runs[0] 会残留半截文字，所以必须整段重建。
    """
    p = cell.paragraphs[0]
    rpr = None
    for r in p._p.iter(qn("w:r")):
        rpr = r.find(qn("w:rPr"))
        break
    for child in list(p._p):
        if child.tag in (qn("w:r"), qn("w:hyperlink"), qn("w:smartTag"), qn("w:fldSimple")):
            p._p.remove(child)
    for extra in cell.paragraphs[1:]:
        extra._element.getparent().remove(extra._element)
    run = p.add_run(text)
    if rpr is not None:
        run._r.insert(0, deepcopy(rpr))


def add_members(path, people, out_path=None):
    """people = [(姓名, 项目角色)]，插到人员表最后一个数据行之后。"""
    doc = Document(path)
    target = None
    for tbl in doc.tables:
        if _table_kind(tbl) == "renyuan":
            target = tbl
            break
    if target is None:
        return 0

    hdr = [c.text.strip() for c in target.rows[target.header_idx if hasattr(target, "header_idx") else 2].cells]
    def col(name):
        for i, h in enumerate(hdr):
            if name in h:
                return i
        return -1

    rows = list(target.rows)
    data_rows = [i for i, r in enumerate(rows) if r.cells[0].text.strip().isdigit()]
    if not data_rows:
        return 0
    last = data_rows[-1]
    seq = int(rows[last].cells[0].text.strip())

    prev = rows[last]._tr
    added = 0
    for k, (name, role) in enumerate(people):
        tr = deepcopy(rows[last]._tr)
        prev.addnext(tr)
        prev = tr
        row = _Row(tr, target)
        _set_text(row.cells[col("序号")], str(seq + k + 1))
        _set_text(row.cells[col("单位类型")], "施工单位")
        if col("单位名称") >= 0:
            _set_text(row.cells[col("单位名称")], "江苏中天市政工程有限公司")
        _set_text(row.cells[col("姓名")], name)
        _set_text(row.cells[col("项目角色")], role)
        for key in ("联系方式", "备注"):
            if col(key) >= 0:
                _set_text(row.cells[col(key)], "")
        added += 1
    doc.save(out_path or path)
    return added


def suggest_members(path):
    """特种工登记表里有、项目人员信息表里没有的人 → [(姓名, 岗位)]"""
    from .docx_reader import Doc
    d = Doc(path)
    ry, tz = d.one("renyuan"), d.one("tezhong")
    if not ry or not tz:
        return []
    t1 = ry.first_table()
    i = t1.col("姓名")
    if i < 0:
        return []
    have = {r[i].strip() for r in t1.records() if i < len(r)}
    t2 = tz.first_table()
    jn, jr = t2.col("姓名"), t2.col("作业岗位")
    if jn < 0:
        return []
    out = []
    for r in t2.records():
        name = r[jn].strip() if jn < len(r) else ""
        role = (r[jr].strip() if 0 <= jr < len(r) else "") or "施工作业人员"
        if name and name not in have:
            out.append((name, role))
    return out


def fix_document(path, out_path=None, backup_dir=None):
    """备份 → 重排编号 → 补人员，返回 {'backup':..., 'codes':n, 'members':n}"""
    import os
    import shutil
    import datetime

    bak = None
    if backup_dir is not False:
        bdir = backup_dir or os.path.join(os.path.dirname(os.path.abspath(path)), "_改前备份")
        os.makedirs(bdir, exist_ok=True)
        stamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        base = os.path.splitext(os.path.basename(path))[0]
        bak = os.path.join(bdir, "%s.%s.docx" % (base, stamp))
        shutil.copy2(path, bak)

    target = out_path or path
    people = suggest_members(path)
    n_codes = apply(path, target)
    n_mem = add_members(target, people) if people else 0
    return {"backup": bak, "codes": n_codes, "members": n_mem, "people": people}
