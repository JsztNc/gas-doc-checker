# -*- coding: utf-8 -*-
"""按规范重排工序单据编号。

背景（NC26E365 实测）：钢制管道焊接记录里每行是一张单据，单据编号占了
GX0006~GX0035；其余工序单据的编号混在这段号里，造成重复、顺序颠倒。
处理方式：
  · 焊接记录的 21 个单据号保持不变，只去掉编号里的空格（系统会把
    「NC26E365G X0019」当成另一个号）；
  · 其余工序单据从 0036 起，按出现顺序「全局连续」编号（同一工序多行、
    跨工序都不撞号）：螺纹连接 → 管道防腐 → 隐蔽工程 → 管道吹扫 →
    强度试验 → 严密性试验 → 附件安装
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

# 需要重排流水号的工序（焊接记录号段保持不变）。
# 从 RESEQ_START 起，按表格出现顺序「全局连续」编号——同一工序多行、跨工序都不会撞号。
ORDER = [("luowen", 36), ("fangfu", 37), ("yinbi", 38), ("chuixao", 39),
         ("qiangdu", 40), ("mishi", 41), ("fujian", 42)]
RESEQ_START = 36
_RESEQ_KINDS = {k for k, _ in ORDER}

# 单据类型中文名（用于对比弹窗里的「位置」说明）
_KIND_CN = {"luowen": "螺纹连接", "fangfu": "防腐", "yinbi": "隐蔽工程",
            "chuixao": "管道吹扫", "qiangdu": "强度试验", "mishi": "严密性试验",
            "fujian": "附件安装", "ganghan": "钢质管道焊接", "renyuan": "项目人员信息表",
            "tezhong": "特种工登记表", "huizong": "工序质量检查汇总表",
            "shiyan": "试验记录", "wajue": "沟槽开挖", "huitian": "沟槽回填",
            "kaigong": "开工报告", "kaitong": "阀门/附件"}


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


def _plan_codes(doc):
    """生成重排方案 [(表号0基, 行号0基, 工序kind, 编号格, 新号)]。

    · 焊接记录：保留本行原号，只去掉夹在编号里的空格；
    · 其余工序：从 RESEQ_START 起按出现顺序全局连续编号，
      同一工序多行、跨工序都保证不撞号。
    """
    plan = []
    seq = RESEQ_START
    for ti, tbl in enumerate(doc.tables):
        kind = _table_kind(tbl)
        if kind != "ganghan" and kind not in _RESEQ_KINDS:
            continue
        for ri, row in enumerate(tbl.rows):
            for cell in row.cells:
                hit = _code_in_cell(cell)
                if not hit:
                    continue
                pid, oldseq = hit
                if kind == "ganghan":
                    new = "%sGX%04d" % (pid, oldseq)
                else:
                    new = "%sGX%04d" % (pid, seq)
                    seq += 1
                plan.append((ti, ri, kind, cell, new))
                break                                 # 一行只认第一个编号格
    return plan


def _code_loc(ti, ri, kind):
    return "表%d %s 第%d行" % (ti + 1, _KIND_CN.get(kind, kind), ri + 1)


def analyze(path):
    """dry-run：返回 [(位置说明, 旧号, 新号)]，不写文件。"""
    doc = Document(path)
    plan = []
    for ti, ri, kind, cell, new in _plan_codes(doc):
        old = cell.text.strip()
        if old != new:
            plan.append((_code_loc(ti, ri, kind), old, new))
    return plan


def apply(path, out_path=None):
    """按 analyze 的方案改写。返回 diffs = [(位置, 旧值, 新值)]。"""
    doc = Document(path)
    diffs = []
    for ti, ri, kind, cell, new in _plan_codes(doc):
        old = cell.text.strip()
        if old != new and _set_cell_code(cell, new):
            diffs.append((_code_loc(ti, ri, kind), old, new))
    doc.save(out_path or path)
    return diffs


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
    added = []
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
        added.append((name, role))
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


# 合法工程编号形如 NC25E571 / IE25E001 / NC23E622GX0002（允许中间夹空格）
_PID_VALUE = _re.compile(r"^[A-Z]{2}\d{2}[A-Z0-9]{0,20}$", _re.I)


def _logical_cells(row):
    """python-docx 对合并单元格会重复返回同一个 <w:tc>；按底层 tc 去重，
    返回 [(列号, 单元格)]，这样「标签格 → 值格」才不会指回标签自己。"""
    out, seen = [], set()
    for ci, c in enumerate(row.cells):
        key = id(c._tc)
        if key in seen:
            continue
        seen.add(key)
        out.append((ci, c))
    return out


def _iter_pid_cells(tbl):
    """产出 (行号, 标签文本, 值格)：每行最多认一个 项目编号/工程编号 标签。"""
    for ri, row in enumerate(tbl.rows):
        logical = _logical_cells(row)
        for k, (_ci, c) in enumerate(logical):
            label = _re.sub(r"\s+", "", c.text or "")
            if label in ("项目编号", "工程编号"):
                if k + 1 < len(logical):
                    yield ri, label, logical[k + 1][1]
                break


def analyze_pid(path):
    """dry-run：返回 [(位置, 旧值, 新值)]，不写文件。"""
    parsed = _ParsedDoc(path)
    pid = _canonical_pid(parsed)
    if not pid:
        return []
    plan = []
    for ti, tbl in enumerate(_PyDoc(path).tables):
        for ri, label, vc in _iter_pid_cells(tbl):
            cur = _re.sub(r"\s+", "", vc.text or "").upper()
            if _PID_VALUE.match(cur) and cur != pid:
                plan.append(("表%d 第%d行「%s」" % (ti + 1, ri + 1, label),
                             vc.text.strip(), pid))
    return plan


def apply_pid(path, out_path=None):
    pid = _canonical_pid(_ParsedDoc(path))
    if not pid:
        return []
    doc = _PyDoc(path)
    diffs, seen = [], set()
    for ti, tbl in enumerate(doc.tables):
        for ri, label, vc in _iter_pid_cells(tbl):
            cur = _re.sub(r"\s+", "", vc.text or "").upper()
            # 只改「本来就是编号、但和主导编号不一致」的值格，绝不碰标签/名称
            if not _PID_VALUE.match(cur) or cur == pid:
                continue
            key = id(vc._tc)
            if key in seen:                 # 同一合并格可能出现在多行视图里
                continue
            seen.add(key)
            old = vc.text.strip()
            _set_text(vc, pid)
            diffs.append(("表%d 第%d行「%s」" % (ti + 1, ri + 1, label), old, pid))
    doc.save(out_path or path)
    return diffs


def _summary_count(parsed, name):
    """汇总表某工序名称应写的检查次数（单据张数），无匹配返回 None。

    直接复用 rules.summary_target —— 检查（r_summary）和修正用同一口径，
    避免「检查说 3 次、修正却没改」这类漂移。
    """
    from .rules import summary_target
    _level, cnt = summary_target(_Facts(parsed), name)
    return cnt


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
    diffs = []
    for ti, tbl in enumerate(doc.tables):
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
                old = row.cells[cci].text.strip() or "（空）"
                _set_text(row.cells[cci], str(tgt))
                diffs.append(("表%d 汇总「%s」" % (ti + 1, name), old, str(tgt)))
    doc.save(out_path or path)
    return diffs


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
    diffs = []
    for ti, tbl in enumerate(doc.tables):
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
        old = target.text.strip() or "（空）"
        prefix = "；" if (target.text or "").strip() else ""
        _append_text(target, prefix + STD_NOTE)
        diffs.append(("表%d 试验记录备注栏" % (ti + 1), old, prefix + STD_NOTE))
    doc.save(out_path or path)
    return diffs


def fix_document(path, out_path=None, backup_dir=None, save_as=True):
    """按规范自动修正，返回修改前后对比所需的全部信息。

    默认 save_as=True：把修正结果另存为「<原名>_修正后.docx」，原文件不动；
    仅在 save_as=False（就地覆盖）时才写一份 _改前备份。

    返回 dict：
        saved_as : 修正后文件（save_as 时即另存文件，否则等于 path）
        original : 原始文件（未改动）
        backup   : 就地覆盖时的备份路径（另存模式为 None）
        diffs    : [(类别, 位置, 修改前, 修改后), ...] 用于前后对比展示
        pid/summary/note/codes/members : 各类改动计数；people : 补充人员列表
    PDF 暂不支持写回，显式报错提示先转 Word。
    """
    import os
    import shutil
    import datetime

    if path.lower().endswith(".pdf"):
        raise RuntimeError("自动修正目前只支持 Word(.docx)；这份是 PDF，请先用 WPS/Adobe 的"
                           "「PDF 转 Word」转成可编辑 Word 后再修正。")

    if out_path:
        target = out_path
    elif save_as:
        d = os.path.dirname(os.path.abspath(path))
        base = os.path.splitext(os.path.basename(path))[0]
        target = os.path.join(d, "%s_修正后.docx" % base)
    else:
        target = path

    # 只有在「就地覆盖」时才需要备份；另存模式下原文件本身就是「改前」
    bak = None
    if target == path and backup_dir is not False:
        bdir = backup_dir or os.path.join(os.path.dirname(os.path.abspath(path)), "_改前备份")
        os.makedirs(bdir, exist_ok=True)
        stamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        bbase = os.path.splitext(os.path.basename(path))[0]
        bak = os.path.join(bdir, "%s.%s.docx" % (bbase, stamp))
        shutil.copy2(path, bak)

    # 从原文件复制出工作副本，再在工作副本上逐步套用各项修正
    # （就地覆盖时 target == path，不能自己拷自己）
    if os.path.abspath(target) != os.path.abspath(path):
        shutil.copy2(path, target)

    diffs = []
    for loc, old, new in apply_pid(target, target):
        diffs.append(("项目编号", loc, old, new))
    # 先规范单据编号（去空格 / 消除重复），再按改后实际张数回写汇总次数，
    # 否则汇总目标值会按「旧的重复编号」算，与复检结果对不上。
    for loc, old, new in apply(target, target):
        diffs.append(("单据编号", loc, old, new))
    for loc, old, new in apply_summary(target, target):
        diffs.append(("汇总次数", loc, old, new))
    for loc, old, new in apply_note(target, target):
        diffs.append(("试验备注", loc, old, new))

    people = suggest_members(target)
    added = add_members(target, people) if people else []
    for nm, role in added:
        diffs.append(("补人员", "项目人员信息表", "（原表缺）", "%s / %s" % (nm, role)))

    cnt = lambda c: sum(1 for d in diffs if d[0] == c)
    return {"saved_as": target, "original": os.path.abspath(path), "backup": bak,
            "diffs": diffs, "pid": cnt("项目编号"), "summary": cnt("汇总次数"),
            "note": cnt("试验备注"), "codes": cnt("单据编号"),
            "members": len(added), "people": people}
