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
        fixed = "keep" if kind == "ganghan" else start.get(kind)
        if fixed is None:
            continue
        seq = fixed if fixed != "keep" else 0
        for ri, row in enumerate(tbl.rows):
            for ci, cell in enumerate(row.cells):
                hit = _code_in_cell(cell)
                if not hit:
                    continue
                pid, oldseq = hit
                # 焊接：保留本行原号只去空格；其余工序：按出现顺序连续编号，不把多行压成同一个号
                new = "%sGX%04d" % (pid, oldseq if fixed == "keep" else seq)
                old = cell.text.strip()
                if old != new:
                    plan.append(("表%d %s R%dC%d" % (ti + 1, kind, ri + 1, ci + 1),
                                 old, new))
                if fixed != "keep":
                    seq += 1
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
        seq = fixed if fixed != "keep" else 0
        for row in tbl.rows:
            for cell in row.cells:
                hit = _code_in_cell(cell)
                if not hit:
                    continue
                pid, oldseq = hit
                new = "%sGX%04d" % (pid, oldseq if fixed == "keep" else seq)
                if _set_cell_code(cell, new):
                    changed += 1
                if fixed != "keep":
                    seq += 1
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

    # 人员表表头：在前 3 行里找含「项目角色 / 姓名」那一行（python-docx 表无 header_idx）
    hri = 0
    for ri, row in enumerate(target.rows[:3]):
        if any(("项目角色" in (c.text or "")) or ("姓名" in (c.text or "")) for c in row.cells):
            hri = ri
            break
    hdr = [c.text.strip() for c in target.rows[hri].cells]
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


# --------------------------------------------------------------------------
# 新增自动修正（均有唯一正确答案的机械错误，可安全代改）：
#   1) 项目编号统一   2) 汇总表检查次数修正   3) 发泡剂检漏备注补全
# --------------------------------------------------------------------------
import re as _re
from collections import Counter
from copy import deepcopy as _dc

from docx import Document as _PyDoc
from .docx_reader import Doc as _ParsedDoc
from .rules import Facts as _Facts, bill_codes as _bill_codes

STD_NOTE = "已采用发泡剂检查所有焊口、接口进行检查无渗漏均为合格"


def _qn(tag):
    from docx.oxml.ns import qn as _q
    return _q(tag)


def _find_col(tbl, *names):
    """在表格前若干行里找含某字样的表头格，返回 (行号, 列号)，找不到 (-1,-1)。"""
    for ri, row in enumerate(tbl.rows[:6]):
        for ci, c in enumerate(row.cells):
            t = _re.sub(r"\s+", "", c.text or "")
            if any(n in t for n in names):
                return ri, ci
    return -1, -1


def _append_text(cell, text):
    """在单元格末尾追加一段文字（保留原有内容 / 格式）。"""
    p = cell.paragraphs[-1] if cell.paragraphs else cell.add_paragraph()
    run = p.add_run(text)
    src = None
    for pp in cell.paragraphs:
        for r in pp.runs:
            src = r
            break
        if src:
            break
    if src is not None:
        rpr = src._r.find(_qn("w:rPr"))
        if rpr is not None:
            run._r.insert(0, _dc(rpr))
    return run


def _canonical_pid(parsed):
    """取全套资料里出现最多的 项目编号/工程编号 作为规范值。"""
    cnt = Counter()
    for b in parsed.bills:
        t = b.first_table()
        if not t:
            continue
        kv = t.kv()
        for k in ("项目编号", "工程编号"):
            v = (kv.get(k) or "").replace(" ", "").upper()
            if v:
                cnt[v] += 1
    return cnt.most_common(1)[0][0] if cnt else None


def analyze_pid(path):
    """dry-run：返回 [(位置, 旧值, 新值)]，不写文件。"""
    parsed = _ParsedDoc(path)
    pid = _canonical_pid(parsed)
    if not pid:
        return []
    plan = []
    for ti, tbl in enumerate(_PyDoc(path).tables):
        for ri, row in enumerate(tbl.rows):
            for ci, c in enumerate(row.cells):
                if _re.sub(r"\s+", "", c.text or "") in ("项目编号", "工程编号"):
                    if ci + 1 >= len(row.cells):
                        continue
                    nxt = row.cells[ci + 1]
                    cur = _re.sub(r"\s+", "", nxt.text or "").upper()
                    if cur and cur != pid:
                        plan.append(("表%d R%dC%d" % (ti + 1, ri + 1, ci + 2),
                                     nxt.text.strip(), pid))
    return plan


def apply_pid(path, out_path=None):
    pid = _canonical_pid(_ParsedDoc(path))
    if not pid:
        return 0
    doc = _PyDoc(path)
    changed = 0
    for tbl in doc.tables:
        for row in tbl.rows:
            for ci, c in enumerate(row.cells):
                if _re.sub(r"\s+", "", c.text or "") in ("项目编号", "工程编号"):
                    if ci + 1 >= len(row.cells):
                        continue
                    nxt = row.cells[ci + 1]
                    cur = _re.sub(r"\s+", "", nxt.text or "").upper()
                    if cur and cur != pid:
                        _set_text(nxt, pid)
                        changed += 1
    doc.save(out_path or path)
    return changed


def _summary_count(parsed, name):
    """汇总表某工序名称应写的检查次数（单据张数），无匹配返回 None。"""
    facts = _Facts(parsed)
    name = _re.sub(r"\s+", "", name or "")
    if "强度" in name or "严密" in name:
        indoors = True if "室内" in name else (False if "室外" in name else None)
        strength = "强度" in name
        sel = [b for b in facts.all("shiyan")
               if (indoors is None or ("室内" in b.title) == indoors)
               and (("强度" in b.title) == strength)]
        if not sel:
            return None
        return sum(len(set(_bill_codes(b))) for b in sel)
    FIXED = [("PE管电熔", "dianrong_jc"), ("PE管热熔", "rean_jc"),
             ("钢质管道焊接", "ganghan"), ("螺纹连接", "luowen"), ("防腐", "fangfu"),
             ("附属设施安装", "fujian"), ("沟槽回填", "huitian"), ("沟槽开挖", "wajue"),
             ("隐蔽", "yinbi"), ("吹扫", "chuixao")]
    for kw, kind in FIXED:
        if kw in name:
            b = facts.get(kind)
            return len(set(_bill_codes(b))) if b else None
    return None


def analyze_summary(path):
    parsed = _ParsedDoc(path)
    plan = []
    for ti, tbl in enumerate(_PyDoc(path).tables):
        if _table_kind(tbl) != "huizong":
            continue
        hri, nci = _find_col(tbl, "工序名称")
        _, cci = _find_col(tbl, "检查次数", "次数")
        if nci < 0 or cci < 0:
            continue
        for ri in range(hri + 1, len(tbl.rows)):
            row = tbl.rows[ri]
            if cci >= len(row.cells):
                continue
            name = row.cells[nci].text.strip()
            cur = _re.sub(r"\s+", "", row.cells[cci].text or "")
            tgt = _summary_count(parsed, name)
            if tgt is None:
                continue
            if cur != str(tgt):
                plan.append(("表%d 汇总「%s」" % (ti + 1, name),
                             row.cells[cci].text.strip() or "（空）", str(tgt)))
    return plan


def apply_summary(path, out_path=None):
    doc = _PyDoc(path)
    parsed = _ParsedDoc(path)
    changed = 0
    for tbl in doc.tables:
        if _table_kind(tbl) != "huizong":
            continue
        hri, nci = _find_col(tbl, "工序名称")
        _, cci = _find_col(tbl, "检查次数", "次数")
        if nci < 0 or cci < 0:
            continue
        for ri in range(hri + 1, len(tbl.rows)):
            row = tbl.rows[ri]
            if cci >= len(row.cells):
                continue
            name = row.cells[nci].text.strip()
            cur = _re.sub(r"\s+", "", row.cells[cci].text or "")
            tgt = _summary_count(parsed, name)
            if tgt is None:
                continue
            if cur != str(tgt):
                _set_text(row.cells[cci], str(tgt))
                changed += 1
    doc.save(out_path or path)
    return changed


def analyze_note(path):
    plan = []
    for ti, tbl in enumerate(_PyDoc(path).tables):
        k = _table_kind(tbl)
        if k not in ("qiangdu", "mishi", "shiyan"):
            continue
        hri, bri = _find_col(tbl, "备注", "备注栏")
        if bri < 0:
            continue
        has = any("发泡剂" in (tbl.rows[r].cells[bri].text or "")
                  for r in range(hri + 1, len(tbl.rows)))
        if has:
            continue
        target = None
        for r in range(hri + 1, len(tbl.rows)):
            if (tbl.rows[r].cells[bri].text or "").strip():
                target = tbl.rows[r].cells[bri]
                break
        if target is None and hri + 1 < len(tbl.rows):
            target = tbl.rows[hri + 1].cells[bri]
        if target is not None:
            plan.append(("表%d 试验备注" % (ti + 1),
                         target.text.strip() or "（空）", STD_NOTE))
    return plan


def apply_note(path, out_path=None):
    doc = _PyDoc(path)
    changed = 0
    for tbl in doc.tables:
        k = _table_kind(tbl)
        if k not in ("qiangdu", "mishi", "shiyan"):
            continue
        hri, bri = _find_col(tbl, "备注", "备注栏")
        if bri < 0:
            continue
        has = any("发泡剂" in (tbl.rows[r].cells[bri].text or "")
                  for r in range(hri + 1, len(tbl.rows)))
        if has:
            continue
        target = None
        for r in range(hri + 1, len(tbl.rows)):
            if (tbl.rows[r].cells[bri].text or "").strip():
                target = tbl.rows[r].cells[bri]
                break
        if target is None and hri + 1 < len(tbl.rows):
            target = tbl.rows[hri + 1].cells[bri]
        if target is None:
            continue
        prefix = "；" if (target.text or "").strip() else ""
        _append_text(target, prefix + STD_NOTE)
        changed += 1
    doc.save(out_path or path)
    return changed


def fix_document(path, out_path=None, backup_dir=None):
    """备份 → 项目编号统一 → 汇总次数修正 → 发泡剂备注补全 → 重排编号 → 补人员。

    返回 {'backup','pid','summary','note','codes','members','people'} 各组改动数。
    PDF 暂不支持写回，显式报错提示先转 Word。
    """
    import os
    import shutil
    import datetime

    if path.lower().endswith(".pdf"):
        raise RuntimeError("自动修正目前只支持 Word(.docx)；这份是 PDF，请先用 WPS/Adobe 的"
                           "「PDF 转 Word」转成可编辑 Word 后再修正。")

    bak = None
    if backup_dir is not False:
        bdir = backup_dir or os.path.join(os.path.dirname(os.path.abspath(path)), "_改前备份")
        os.makedirs(bdir, exist_ok=True)
        stamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        base = os.path.splitext(os.path.basename(path))[0]
        bak = os.path.join(bdir, "%s.%s.docx" % (base, stamp))
        shutil.copy2(path, bak)

    target = out_path or path
    n_pid = apply_pid(path, target)
    n_sum = apply_summary(target, target) if n_pid >= 0 else 0
    n_note = apply_note(target, target)
    n_codes = apply(target, target)
    people = suggest_members(target)
    n_mem = add_members(target, people) if people else 0
    return {"backup": bak, "pid": n_pid, "summary": n_sum, "note": n_note,
            "codes": n_codes, "members": n_mem, "people": people}
