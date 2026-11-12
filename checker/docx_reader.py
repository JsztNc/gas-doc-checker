# -*- coding: utf-8 -*-
"""读取「线上资料.docx」，按单据拆成结构化数据。

不依赖 python-docx：docx 就是 zip + XML，用标准库解析 document.xml，
按文档顺序取出段落和表格，再按表头特征识别单据类型。
"""
import re
import zipfile
import xml.etree.ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# 单据编号，如 NC23A005GX0008 / NW26-21GX0014 / NC23A005CL01
CODE_RE = re.compile(r"[A-Z]{2}\d{2}[-A-Z0-9]{2,8}?(?:GX|HS|JD|BS|KGBG|AQ|CL|JG|WT)\d*", re.I)


def _cell_text(el):
    parts = []
    for node in el.iter():
        if node.tag == W + "t":
            parts.append(node.text or "")
        elif node.tag == W + "tab":
            parts.append(" ")
        elif node.tag == W + "br":
            parts.append("\n")
    return "".join(parts).strip()


def read_items(path):
    """返回 [("p", text) | ("t", [[cell, ...], ...]), ...]，保持文档顺序。"""
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    body = ET.fromstring(xml).find(W + "body")
    items = []
    for child in body:
        if child.tag == W + "p":
            t = _cell_text(child)
            if t:
                items.append(("p", t))
        elif child.tag == W + "tbl":
            rows = []
            for tr in child.findall(W + "tr"):
                cells = [_cell_text(tc) for tc in tr.findall(W + "tc")]
                if any(c for c in cells):
                    rows.append(cells)
            if rows:
                items.append(("t", rows))
        elif child.tag == W + "sdt":
            for sub in child.iter(W + "tbl"):
                rows = []
                for tr in sub.findall(W + "tr"):
                    cells = [_cell_text(tc) for tc in tr.findall(W + "tc")]
                    if any(c for c in cells):
                        rows.append(cells)
                if rows:
                    items.append(("t", rows))
    return items


# ---------------------------------------------------------------- 单据识别
# 顺序有意义：越靠前越优先
SIGS = [
    ("mulu",       ["资料名称"]),
    ("huisen",     ["会审通过日期"]),
    ("jiaodi",     ["交底通过日期"]),
    ("fangan",     ["申报内容系第"]),
    ("kaigong",    ["计划开工日期"]),
    ("renyuan",    ["项目角色"]),
    ("tezhong",    ["作业岗位"]),
    ("jiedian",    ["关键节点"]),
    ("wajue",      ["管沟实际挖深"]),
    ("huitian",    ["回填长度"]),
    ("yinbi",      [("隐藏部位", "隐蔽部位")]),
    ("luowen",     ["螺纹质量"]),
    ("fangfu",     [("刷漆防腐质量检查", "检查管道长度")]),
    ("rean_cs",    ["加热时间", "吸热时间"]),
    ("rean_jc",    ["翻边宽度"]),
    ("dianrong_cs", [("氧化时间", "焊接电压", "电熔时间")]),
    ("dianrong_jc", ["承插深度"]),
    ("ganghan",    ["焊口编号", "焊工姓名"]),
    ("chuixao",    ["吹扫压力"]),
    ("fujian",     ["检查部位", "压力试验"]),
    ("huizong",    ["工序名称", "检查次数"]),
    ("anquan",     ["检查类型", "检查项"]),
    ("hexiao",     ["物料编号", "领料量"]),
    ("jungong",    ["验收项"]),
    ("fujian_list", [("单据名称", "附件名称")]),
    ("shiyan",     ["压力表读数"]),
    ("mishi",      ["试验日期", "时间hrs"]),
    ("qiangdu",    ["试验日期", "试验压力"]),
]

HEADER_WORDS = {"序号", "公司名称", "项目编号", "单据编号", "检查时间", "审批时间", "备注",
                "签字栏", "施工单位", "年", "月", "日", "合计", "审批人", "检查人员"}

KIND_LABEL = {
    "mulu": "卷内目录", "huisen": "图纸会审记录", "jiaodi": "设计交底会议纪要",
    "fangan": "施工组织设计（方案）报审表", "kaigong": "开工报告",
    "renyuan": "项目人员信息表", "tezhong": "特种工登记表", "jiedian": "关键时间节点清单",
    "wajue": "沟槽开挖检查记录", "huitian": "沟槽回填检查记录", "yinbi": "隐蔽工程检查记录",
    "luowen": "螺纹连接检查记录", "fangfu": "管道防腐检查记录",
    "rean_cs": "PE管热熔连接工艺参数", "rean_jc": "PE管热熔连接检查记录",
    "dianrong_cs": "PE管电熔连接工艺参数", "dianrong_jc": "PE管电熔连接检查记录",
    "ganghan": "钢制管道焊接检查记录", "chuixao": "管道吹扫检查记录",
    "fujian": "管道附件及设备安装检查记录", "huizong": "工序质量检查汇总表",
    "anquan": "安全及特种作业检查表", "hexiao": "服务商材料使用量单",
    "jungong": "项目竣工验收单", "qiangdu": "强度试验记录", "mishi": "气密（严密性）试验记录",
    "shenpi": "单据审批记录", "shiyan": "试压（强度/严密性）记录",
    "fujian_list": "其他格式附件",
}


class Table(object):
    def __init__(self, rows):
        self.rows = rows
        self.header_idx = self._guess_header()

    def _guess_header(self):
        """表头行 = 前 8 行里字段最多、且不像签字栏/数据行的那一行。"""
        best, best_score = 0, -1
        for i, row in enumerate(self.rows[:8]):
            texts = [c for c in row if c]
            if len(texts) < 3:
                continue
            if sum(1 for c in texts if c in ("年", "月", "日")) >= 2:
                continue                      # 签字栏
            if all(re.fullmatch(r"[\d\.\-\s/:]+", c) for c in texts):
                continue                      # 纯数字行
            score = len(texts)
            if score > best_score:
                best, best_score = i, score
        return best

    @property
    def header(self):
        return self.rows[self.header_idx] if self.rows else []

    def col(self, *names):
        """返回列号：表头包含任一名（模糊）。找不到返回 -1。"""
        for i, h in enumerate(self.header):
            hh = re.sub(r"\s+", "", h)
            for n in names:
                if n in hh:
                    return i
        return -1

    def records(self):
        """表头之后的有效数据行，跳过每页重复的表头/签字栏/影像页。"""
        out = []
        for row in self.rows[self.header_idx + 1:]:
            if not row or not row[0].strip():
                continue
            texts = [re.sub(r"\s+", "", c) for c in row if c]
            if sum(1 for c in texts if c in HEADER_WORDS) >= 2:
                continue
            if any(t.startswith(("备注：", "备注:", "签字栏", "影像照片")) for t in texts):
                continue
            out.append(row)
        return out

    def kv(self):
        """两列一组的关键值（"公司名称 | 南昌燃气 | 单据编号 | xxx"）。"""
        d = {}
        for row in self.rows[: self.header_idx + 1]:
            cells = [c.strip() for c in row]
            for i in range(0, len(cells) - 1, 2):
                k, v = cells[i], cells[i + 1]
                if k and v and len(k) <= 16:
                    d.setdefault(re.sub(r"\s+", "", k), v)
        return d

    def text(self):
        return "\n".join(" | ".join(c for c in r if c) for r in self.rows)


def classify(rows):
    blob = "\n".join(" | ".join(c for c in r if c) for r in rows[:6])
    flat = re.sub(r"\s+", "", blob)
    for kind, groups in SIGS:
        ok = True
        for g in groups:
            alts = g if isinstance(g, tuple) else (g,)
            if not any(re.sub(r"\s+", "", a) in flat for a in alts):
                ok = False
                break
        if ok:
            return kind
    return None


class Bill(object):
    def __init__(self, kind, title):
        self.kind = kind
        self.title = title or KIND_LABEL.get(kind, kind)
        self.tables = []

    @property
    def label(self):
        return KIND_LABEL.get(self.kind, self.kind)

    def first_table(self):
        return self.tables[0] if self.tables else None

    def all_text(self):
        return "\n".join(t.text() for t in self.tables)

    def codes(self):
        """单据编号。

        有「单据编号」这一列时只认这一列——焊接记录里每行的「焊口编号」是
        另一套号（规范里焊口编号沿用工程缩写），混进来会被误判成编号重复。
        """
        found = []
        for t in self.tables:
            col = t.col("单据编号")
            if col >= 0:
                for row in t.records():
                    if col < len(row):
                        for m in CODE_RE.finditer(row[col].upper().replace(" ", "")):
                            found.append(m.group(0))
        if found:
            return found
        for t in self.tables:
            for row in t.rows:
                for c in row:
                    for m in CODE_RE.finditer(c.upper().replace(" ", "")):
                        found.append(m.group(0))
        return found

    def weld_codes(self):
        """焊口编号（焊接/连接记录里每行一个），与单据编号分开统计。"""
        out = []
        for t in self.tables:
            col = t.col("焊口编号")
            if col < 0:
                continue
            for row in t.records():
                if col < len(row):
                    out.append(row[col])
        return out

    def approvals(self):
        """表内可能夹着「单据审批记录」子表，逐行找表头再往下读。"""
        out = []
        for ti, t in enumerate(self.tables):
            for i, row in enumerate(t.rows):
                norm = [re.sub(r"\s+", "", c) for c in row]
                if "审批节点" not in norm or "审批人" not in norm:
                    continue
                idx = {}
                for j, c in enumerate(norm):
                    for key in ("审批节点", "审批人", "审批时间", "审批状态"):
                        if c.startswith(key) and key not in idx:
                            idx[key] = j
                for r2 in t.rows[i + 1:]:
                    if not r2 or not r2[0].strip().isdigit():
                        break

                    def g(key):
                        k = idx.get(key, -1)
                        return r2[k].strip() if 0 <= k < len(r2) else ""

                    tm = parse_dt(g("审批时间"))
                    if not tm:                       # 合并单元格错位时退回整行找时间
                        for c in r2:
                            tm = parse_dt(c)
                            if tm:
                                break
                    if tm:
                        out.append({"node": g("审批节点"), "person": g("审批人"),
                                    "status": g("审批状态"), "time": tm,
                                    "cells": len(r2), "tbl": ti,
                                    "filled": sum(1 for c in r2 if c.strip())})
        return out


DT_RE = re.compile(r"(\d{4})[-./年](\d{1,2})[-./月](\d{1,2})日?\s*(\d{1,2})[:：](\d{2})(?:[:：](\d{2}))?")


def parse_dt(s):
    """宽容解析日期时间：2024-08-0115:08:03 / 2024.08.30 / 2024-12-29 11:24:01。"""
    if not s:
        return None
    s = str(s).strip()
    m = DT_RE.search(s)
    if m:
        y, mo, d, hh, mi, ss = m.groups()
        return "%04d-%02d-%02d %02d:%02d:%02d" % (int(y), int(mo), int(d), int(hh), int(mi), int(ss or 0))
    m = re.search(r"(\d{4})[-./年](\d{1,2})[-./月](\d{1,2})日?", s)
    if m:
        y, mo, d = m.groups()
        return "%04d-%02d-%02d 00:00:00" % (int(y), int(mo), int(d))
    return None


def parse_date_only(s):
    dt = parse_dt(s)
    return dt[:10] if dt else None


def build_bills(items):
    """把 items 串成 Bill 列表。无标题的连续表格按表头判定类型并接续。"""
    bills = []
    pending = None
    for kind, val in items:
        if kind == "p":
            t = val.strip()
            if re.search(r"\(\d+/\d+\)$", t) or CODE_RE.fullmatch(t.replace(" ", "")):
                continue                      # 页码/单据号，不是标题
            if len(t) <= 40 and not t.startswith(("备注", "签字栏", "（", "(")):
                pending = t
            continue
        rows = val
        k = classify(rows)
        if k is None:
            # 无表头的续页表格：列数一致就接到上一张表
            if bills and bills[-1].tables and len(rows[0]) == len(bills[-1].tables[-1].rows[0]):
                bills[-1].tables.append(Table(rows))
            continue
        # 同类型连续表格且列数一致 = 同一张单据的续页
        if pending is None and bills and bills[-1].kind == k and bills[-1].tables \
                and len(rows[0]) == len(bills[-1].tables[-1].rows[0]):
            bills[-1].tables.append(Table(rows))
            continue
        b = Bill(k, pending)
        b.tables.append(Table(rows))
        bills.append(b)
        pending = None
    return bills


class Doc(object):
    def __init__(self, path):
        self.path = path
        self.items = read_items(path)
        self.bills = build_bills(self.items)
        self.meta = self._meta()

    def _meta(self):
        meta = {}
        for b in self.bills:
            t = b.first_table()
            if not t:
                continue
            kv = t.kv()
            for src, dst in (("项目编号", "项目编号"), ("工程编号", "项目编号"),
                             ("项目名称", "项目名称"), ("工程名称", "项目名称"),
                             ("施工单位", "施工单位"), ("服务商", "施工单位")):
                if src in kv and dst not in meta:
                    meta[dst] = kv[src]
        return meta

    def by_kind(self, kind):
        return [b for b in self.bills if b.kind == kind]

    def one(self, kind):
        found = self.by_kind(kind)
        return found[0] if found else None
