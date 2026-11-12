# -*- coding: utf-8 -*-
"""按《线上燃气竣工资料填写规范与校验要点》逐项检查。

每条规则返回 Finding 列表；数据不足时返回 info 说明「无法判断」，不硬报错。
"""
import datetime
import re

from .docx_reader import CODE_RE, KIND_LABEL, parse_date_only, parse_dt

ERR, WARN, INFO = "错误", "警告", "提示"


class Finding(object):
    def __init__(self, level, rule, msg, where=""):
        self.level = level
        self.rule = rule
        self.msg = msg
        self.where = where

    def __repr__(self):
        return "[%s] %s %s" % (self.level, self.rule, self.msg)


# ---------------------------------------------------------------- 取值工具
def rows_of(bill):
    """单据下的明细行（跨续页表格、跳过表头）。"""
    out = []
    for t in bill.tables:
        if t.col("审批节点") >= 0:
            continue
        out.extend(t.records())
    return out


def cell(row, idx):
    return row[idx].strip() if 0 <= idx < len(row) else ""


def first_time(bill, *names):
    """取单据里某列的最早时间。"""
    if bill is None:
        return None
    times = []
    for t in bill.tables:
        for name in names:
            i = t.col(name)
            if i < 0:
                continue
            for row in t.records():
                tm = parse_dt(cell(row, i))
                if tm:
                    times.append(tm)
    return min(times) if times else None


def last_time(bill, *names):
    if bill is None:
        return None
    times = []
    for t in bill.tables:
        for name in names:
            i = t.col(name)
            if i < 0:
                continue
            for row in t.records():
                tm = parse_dt(cell(row, i))
                if tm:
                    times.append(tm)
    return max(times) if times else None


def approval_last(bill):
    if bill is None:                     # 缺这类单据时不要炸
        return None
    ap = bill.approvals()
    return max(a["time"] for a in ap) if ap else None


def approval_first(bill):
    if bill is None:
        return None
    ap = bill.approvals()
    return min(a["time"] for a in ap) if ap else None


def bill_codes(bill):
    return [c for c in bill.codes() if "GX" in c or "AQ" in c]


def num_of(code):
    m = re.search(r"(GX|AQ|CL|HS|JD|BS)(\d+)$", code.upper())
    return int(m.group(2)) if m else None


def kind_first(doc, kind):
    b = doc.one(kind)
    return b


class Facts(object):
    """从文档里抽取一次，供各规则复用。"""

    def __init__(self, doc):
        self.doc = doc
        self.meta = doc.meta
        self.pid = (self.meta.get("项目编号") or "").replace(" ", "").upper()
        self.bills = doc.bills
        self.by = {}
        for b in doc.bills:
            self.by.setdefault(b.kind, []).append(b)

    def get(self, kind):
        lst = self.by.get(kind) or []
        return lst[0] if lst else None

    def all(self, kind):
        return self.by.get(kind) or []

    def has(self, kind):
        return bool(self.by.get(kind))

    # 常用时间
    def t_shencha(self, kind):
        b = self.get(kind)
        return first_time(b, "检查时间", "检查日期") if b else None

    def t_pizhun(self, kind):
        b = self.get(kind)
        return approval_last(b) if b else None

    def codes_gx(self):
        out = []
        for b in self.bills:
            for c in bill_codes(b):
                n = num_of(c)
                if n is not None and "GX" in c:
                    out.append((n, c, b))
        return out

    def hexiao_items(self):
        """材料核销单：[(物料名称, 单位, 实际用量)]，跨页去重。"""
        out = []
        seen = set()
        for b in self.all("hexiao"):
            for t in b.tables:
                i_n = t.col("物料名称")
                i_u = t.col("单位")
                i_a = t.col("实际用量")
                i_c = t.col("物料编号")
                if i_n < 0 or i_a < 0:
                    continue
                for row in t.records():
                    cid = cell(row, i_c)
                    if cid and cid in seen:
                        continue
                    seen.add(cid)
                    name = cell(row, i_n)
                    try:
                        qty = float(cell(row, i_a) or 0)
                    except ValueError:
                        continue
                    out.append((name, cell(row, i_u), qty))
        return out


# ---------------------------------------------------------------- 规则
DIR_MAP = [
    ("图纸会审", ["huisen"], "图纸会审"),
    ("设计交底", ["jiaodi"], "设计交底"),
    ("施工组织设计", ["fangan"], "施工组织设计"),
    ("开工报告", ["kaigong"], "开工报告"),
    ("项目成员信息", ["renyuan"], "项目人员信息"),
    ("项目人员", ["renyuan"], "项目人员信息"),
    ("特种工", ["tezhong"], "特种工"),
    ("关键时间节点", ["jiedian"], "关键时间节点"),
    ("沟槽开挖", ["wajue"], "沟槽开挖"),
    ("阀门", ["wajue"], ""),
    ("回填", ["huitian"], "沟槽回填"),
    ("隐蔽", ["yinbi"], "隐蔽工程"),
    ("螺纹连接", ["luowen"], "螺纹连接"),
    ("防腐", ["fangfu"], "管道防腐"),
    ("热熔", ["rean_cs", "rean_jc"], "热熔"),
    ("电熔", ["dianrong_cs", "dianrong_jc"], "电熔"),
    ("焊接", ["ganghan"], "焊接"),
    ("吹扫", ["chuixao"], "吹扫"),
    ("强度", ["qiangdu", "shiyan"], "强度试验"),
    ("严密", ["mishi", "shiyan"], "气密"),
    ("气密", ["mishi", "shiyan"], "气密"),
    ("附属设施安装", ["fujian"], "管道附件"),
    ("管道、管件", ["fujian"], "管道附件"),
    ("汇总表", ["huizong"], "汇总表"),
    ("安全及特种作业", ["anquan"], "安全"),
    ("材料核销", ["hexiao"], "材料使用量"),
    ("竣工验收", ["jungong"], "竣工验收"),
]


def r_structure(f):
    out = []
    mulu = f.get("mulu")
    if not mulu:
        return [Finding(INFO, "无法核对目录", "文档里没有卷内目录")]
    t = mulu.first_table()
    i = t.col("资料名称")
    listed = []
    for row in t.records():
        name = cell(row, i)
        if not name or len(name) > 24:
            continue
        if not re.match(r"^\d+$", row[0].strip()):      # 跳过目录后面的审批记录行
            continue
        listed.append(name)
    full = " ".join(b.title + b.all_text() for b in f.bills)
    for name in listed:
        kinds = None
        for kw, ks, _ in DIR_MAP:
            if kw in name:
                kinds = ks
                break
        if kinds:
            if not any(f.has(k) for k in kinds):
                out.append(Finding(ERR, "目录有正文无",
                                   "目录第 %d 项「%s」，正文里没有这张单据" % (listed.index(name) + 1, name)))
        else:
            core = re.sub(r"[（(].*?[)）]", "", name)[:4]
            if core and core not in re.sub(r"\s+", "", full):
                out.append(Finding(INFO, "目录有正文无",
                                   "目录第 %d 项「%s」正文里没匹配到，确认是否确实缺" % (listed.index(name) + 1, name)))
    # 项目编号一致性
    if f.pid:
        bad = set()
        for b in f.bills:
            tt = b.first_table()
            if not tt:
                continue
            kv = tt.kv()
            p = (kv.get("项目编号") or kv.get("工程编号") or "").replace(" ", "").upper()
            if p and p != f.pid:
                bad.add((p, b.label))
        for p, lab in sorted(bad):
            out.append(Finding(WARN, "项目编号不一致",
                               "「%s」里写的是 %s，与整套的 %s 不同" % (lab, p, f.pid)))
    # 审批行单元格数异常（合并单元格错位）
    for b in f.bills:
        ap = b.approvals()
        if not ap:
            continue
        cells_cnt = {}
        filled_cnt = {}
        for a in ap:
            cells_cnt[a["cells"]] = cells_cnt.get(a["cells"], 0) + 1
            filled_cnt[a["filled"]] = filled_cnt.get(a["filled"], 0) + 1
        mc = max(cells_cnt, key=cells_cnt.get)
        mf = max(filled_cnt, key=filled_cnt.get)
        done = set()
        for a in ap:
            # 只是某格没填（如 OA 审批中状态留空）不算错位，必须格数本身也不一样
            if a["cells"] == mc or a["filled"] == mf:
                continue
            key = (b.label, a["node"], a["time"])
            if key in done:
                continue
            done.add(key)
            out.append(Finding(INFO, "审批格子被拆开",
                               "「%s」审批节点 %s（%s）这一行是 %d 格、填了 %d 项，其余行是 %d 格 %d 项 —— "
                               "线上导出时把审批人拆成了两格，内容没丢，打印前看一眼要不要在 Word 里合并"
                               % (b.title[:14], a["node"], a["time"][:16],
                                  a["cells"], a["filled"], mc, mf)))
    return out


def r_codes(f):
    out = []
    # 重复
    seen = {}
    for b in f.bills:
        for c in set(bill_codes(b)):        # 同一张单跨多页会重复读到同一号，只算一次
            seen.setdefault(c, []).append(b.label)
    dups = [(c, w) for c, w in sorted(seen.items()) if len(w) > 1 and "AQ" not in c]
    if len(dups) == 1:
        c, w = dups[0]
        out.append(Finding(ERR, "单据编号重复",
                           "%s 出现了 %d 次（%s）" % (c, len(w), "、".join(sorted(set(w))))))
    elif dups:
        sample = "；".join("%s（%s）" % (c, "/".join(sorted(set(w))[:2])) for c, w in dups[:5])
        out.append(Finding(ERR, "单据编号重复",
                           "共 %d 个编号被两张以上单据占用：%s%s"
                           % (len(dups), sample, " 等" if len(dups) > 5 else "")))
    # 编号里带空格
    space_hit = 0
    for b in f.bills:
        for t in b.tables:
            for row in t.rows:
                for cellv in row:
                    if re.search(r"[A-Z]{2}\d{2}[^\s]*\s+\d{3,4}", cellv) and CODE_RE.search(cellv.replace(" ", "")):
                        space_hit += 1
                        if space_hit <= 6:
                            out.append(Finding(ERR, "单据编号带空格",
                                               "「%s」里编号写成「%s」，系统会当成两个号"
                                               % (b.label, cellv.strip())))
    if space_hit > 6:
        out.append(Finding(ERR, "单据编号带空格", "共 %d 处编号中间夹了空格" % space_hit))
    # GX 断号
    nums = sorted({n for n, _, _ in f.codes_gx()})
    if len(nums) > 3:
        gaps = []
        for a, b2 in zip(nums, nums[1:]):
            if b2 - a > 2:
                gaps.append("GX%04d~GX%04d" % (a + 1, b2 - 1))
        if gaps:
            out.append(Finding(INFO, "GX 流水断号",
                               "中间缺号：%s（可能是漏单，也可能是别的编号段）" % "、".join(gaps[:6])))
    # 螺纹连接 vs 吹扫
    cx = f.get("chuixao")
    lw = f.get("luowen")
    if cx and lw:
        cmin = min([num_of(c) for c in bill_codes(cx) if num_of(c)], default=None)
        lmax = max([num_of(c) for c in bill_codes(lw) if num_of(c)], default=None)
        if cmin and lmax and lmax >= cmin:
            out.append(Finding(ERR, "编号先后颠倒",
                               "螺纹连接编号最大到 GX%04d，吹扫最小是 GX%04d —— 螺纹连接应排在吹扫之前"
                               % (lmax, cmin)))
    return out


def r_time_chain(f):
    out = []
    hs = f.get("huisen")
    jd = f.get("jiaodi")
    fa = f.get("fangan")
    kg = f.get("kaigong")
    d_hs = None
    if hs:
        t = hs.first_table()
        d_hs = parse_date_only(t.kv().get("会审通过日期", ""))
    d_jd = None
    if jd:
        d_jd = parse_date_only(jd.first_table().kv().get("交底通过日期", ""))
    a_fa = approval_last(fa) if fa else None
    a_kg = approval_last(kg) if kg else None
    plan_kg = None
    if kg:
        plan_kg = parse_date_only(kg.first_table().kv().get("计划开工日期", ""))
    chain = [("图纸会审", d_hs), ("设计交底", d_jd), ("方案报审审批", a_fa), ("开工报告审批", a_kg)]
    known = [(n, v) for n, v in chain if v]
    for (n1, v1), (n2, v2) in zip(known, known[1:]):
        if v2 < v1:
            out.append(Finding(ERR, "报审时间倒流",
                               "%s（%s）早于 %s（%s）" % (n2, v2[:16], n1, v1[:16])))
    if plan_kg and a_kg and plan_kg < a_kg[:10]:
        out.append(Finding(WARN, "计划开工早于审批",
                           "计划开工 %s 早于开工报告审批 %s" % (plan_kg, a_kg[:16])))
    # 每张单据：审批时间递增 + 首审不早于填报
    seen_msg = set()
    for b in f.bills:
        pages = {}
        for a in b.approvals():
            pages.setdefault(a["tbl"], []).append(a)
        for _, ap in pages.items():
            for a1, a2 in zip(ap, ap[1:]):
                if a2["time"] < a1["time"]:
                    key = (b.label, a1["node"], a1["time"], a2["node"], a2["time"])
                    if key in seen_msg:
                        break
                    seen_msg.add(key)
                    out.append(Finding(ERR, "审批时间倒流",
                                       "「%s」%s（%s）早于 %s（%s）"
                                       % (b.label, a2["node"], a2["time"][:16], a1["node"], a1["time"][:16])))
                    break
        fill = first_time(b, "创建时间", "检查时间", "检查日期")
        ap2 = b.approvals()
        if fill and ap2:
            first_ap = min(a["time"] for a in ap2)
            if first_ap < fill:
                out.append(Finding(WARN, "审批早于填报",
                                   "「%s」首次审批 %s 早于单据填写时间 %s"
                                   % (b.label, first_ap[:16], fill[:16])))
    return out


def r_node_list(f):
    """关键时间节点清单 vs 实际单据时间。"""
    out = []
    jd = f.get("jiedian")
    if not jd:
        return [Finding(INFO, "无法核对节点清单", "文档里没有关键时间节点清单")]
    t = jd.first_table()
    i_n = t.col("关键节点")
    i_t = t.col("批准时间")
    actual = []

    def add(kw, kind, getter):
        v = getter(kind)
        if v:
            actual.append((kw, v))

    add("设计交底", "jiaodi", lambda k: f.t_pizhun(k))
    add("开工报告", "kaigong", lambda k: f.t_pizhun(k))
    add("沟槽开挖", "wajue", lambda k: f.t_shencha(k))
    add("沟槽回填", "huitian", lambda k: f.t_shencha(k))
    add("PE管电熔", "dianrong_jc", lambda k: last_time(f.get(k), "检查时间"))
    add("PE管热熔", "rean_jc", lambda k: last_time(f.get(k), "检查时间"))
    add("钢质管道焊接", "ganghan", lambda k: last_time(f.get(k), "焊接日期"))
    add("隐蔽工程", "yinbi", lambda k: f.t_shencha(k))
    add("管道吹扫", "chuixao", lambda k: last_time(f.get(k), "检查时间"))
    add("竣工验收", "jungong", lambda k: approval_last(f.get(k)))
    add("材料核销", "hexiao", lambda k: approval_last(f.get(k)))

    for row in t.records():
        name = re.sub(r"\s+", "", cell(row, i_n))
        node_t = parse_dt(cell(row, i_t))
        if not name or not node_t:
            continue
        for kw, real in actual:
            if kw and kw in name:
                d1 = node_t[:10]
                d2 = real[:10]
                if d1 != d2:
                    out.append(Finding(WARN, "节点时间与单据不符",
                                       "清单里「%s」写 %s，实际单据是 %s"
                                       % (cell(row, i_n), node_t[:16], real[:16])))
                break
    return out


def r_work_order(f):
    """工序先后：开挖→(连接焊接)→回填→隐蔽→吹扫→强度→严密→附件。"""
    out = []
    T = {}
    T["开挖"] = f.t_shencha("wajue")
    T["回填"] = f.t_shencha("huitian")
    conns = []
    for kind, col in (("rean_jc", "检查时间"), ("dianrong_jc", "检查时间"),
                      ("ganghan", "焊接日期"), ("luowen", "检查时间")):
        v = last_time(f.get(kind), col)
        if v:
            conns.append(v)
    T["连接/焊接完成"] = max(conns) if conns else None
    T["隐蔽"] = f.t_shencha("yinbi")
    T["吹扫"] = last_time(f.get("chuixao"), "检查时间")
    qd = [first_time(b, "试验日期") for b in f.all("qiangdu")]
    T["强度试验"] = min([x for x in qd if x]) if any(qd) else None
    ms = [first_time(b, "试验日期") for b in f.all("mishi")]
    T["严密性试验"] = min([x for x in ms if x]) if any(ms) else None
    T["附件安装"] = f.t_shencha("fujian")

    pairs = [("开挖", "回填", True), ("开挖", "连接/焊接完成", False),
             ("连接/焊接完成", "隐蔽", False), ("回填", "隐蔽", False),
             ("隐蔽", "吹扫", False), ("吹扫", "强度试验", False),
             ("强度试验", "严密性试验", False), ("严密性试验", "附件安装", False)]
    for a, b, same_day_ok in pairs:
        va, vb = T.get(a), T.get(b)
        if not va or not vb:
            continue
        if vb < va:
            if same_day_ok and va[:10] == vb[:10]:
                continue
            out.append(Finding(ERR, "工序时间倒流",
                               "%s（%s）早于 %s（%s）" % (b, vb[:16], a, va[:16])))
    # 材料领用在开工之后
    kg = f.get("kaigong")
    jd = f.get("jiedian")
    if kg and jd:
        plan = parse_dt(kg.first_table().kv().get("计划开工日期", ""))
        t = jd.first_table()
        i_n, i_t = t.col("关键节点"), t.col("批准时间")
        for row in t.records():
            if "材料领用" in cell(row, i_n):
                v = parse_dt(cell(row, i_t))
                if plan and v and v < plan:
                    out.append(Finding(ERR, "材料领用早于开工",
                                       "材料领用 %s 早于计划开工 %s" % (v[:16], plan[:10])))
    return out


def r_finish_order(f):
    """收尾环节：核销在竣工验收前、安全检查不晚于竣工、竣工审批是最后一道。"""
    out = []
    jg = f.get("jungong")
    if not jg:
        return []
    t = jg.first_table()
    yanshou = parse_date_only(t.kv().get("验收时间", ""))
    end = approval_last(jg)
    hx = f.get("hexiao")
    if hx and yanshou:
        conf = last_time(hx, "确认日期")
        if conf and conf[:10] > yanshou:
            out.append(Finding(WARN, "核销晚于竣工验收",
                               "材料核销确认日期 %s 晚于竣工验收时间 %s；规范要求核销单"
                               "填写在竣工验收前一日、确认日期取最后一个审批时间"
                               % (conf[:16], yanshou)))
    aq = f.get("anquan")
    if aq and yanshou:
        last_aq = last_time(aq, "检查时间")
        if last_aq and last_aq[:10] > yanshou:
            out.append(Finding(WARN, "安全检查晚于竣工验收",
                               "安全及特种作业检查最后一次 %s，晚于竣工验收 %s"
                               % (last_aq[:16], yanshou)))
    if end:
        later = []
        for b in f.bills:
            if b is jg or b.kind in ("mulu", "fujian_list"):
                continue
            v = approval_last(b)
            if v and v > end:
                later.append((b.label, v))
        if later:
            lab, v = max(later, key=lambda x: x[1])
            out.append(Finding(INFO, "竣工审批不是最后",
                               "「%s」审批 %s 晚于竣工验收审批 %s" % (lab, v[:16], end[:16])))
    return out


def r_summary(f):
    """汇总表检查次数 vs 前面单张工序记录的实际条数。"""
    out = []
    hz = f.get("huizong")
    if not hz:
        return [Finding(INFO, "无法核对汇总表", "文档里没有工序质量检查汇总表")]
    t = hz.first_table()
    i_n, i_c = t.col("工序名称"), t.col("检查次数")

    def counts(kind):
        b = f.get(kind)
        if not b:
            return None
        return len(rows_of(b)), len(set(bill_codes(b)))

    def shiyan(indoors=None, strength=None):
        sel = [b for b in f.all("shiyan")
               if (indoors is None or (("室内" in b.title) == indoors))
               and (strength is None or (("强度" in b.title) == strength))]
        if not sel:
            return None
        rows = sum(len(rows_of(b)) for b in sel)
        codes = sum(len(set(bill_codes(b))) for b in sel)
        return rows, codes

    FIXED = [("PE管电熔", "dianrong_jc", "warn"), ("PE管热熔", "rean_jc", "warn"),
             ("钢质管道焊接", "ganghan", "warn"), ("螺纹连接", "luowen", "err"),
             ("防腐", "fangfu", "err"), ("附属设施安装", "fujian", "err"),
             ("沟槽回填", "huitian", "err"), ("回填", "huitian", "err"),
             ("沟槽开挖", "wajue", "err"), ("开挖", "wajue", "err"),
             ("隐蔽", "yinbi", "err"), ("吹扫", "chuixao", "err")]
    for row in t.records():
        name = cell(row, i_n)
        try:
            cnt = int(float(cell(row, i_c) or 0))
        except ValueError:
            continue
        target = None
        level = ERR
        for kw, kind, lv in FIXED:
            if kw in name:
                target = counts(kind)
                level = WARN if lv == "warn" else ERR
                break
        if target is None and ("强度" in name or "严密" in name):
            level = WARN
            if "室内" in name:
                target = shiyan(True, "强度" in name)
            elif "室外" in name:
                target = shiyan(False, "强度" in name)
            else:
                target = shiyan(None, "强度" in name)
        if not target:
            continue
        rows, codes = target
        if cnt not in (rows, codes):
            out.append(Finding(level, "汇总次数对不上",
                               "汇总表「%s」写 %d 次；前面明细 %d 条、单据 %d 个（两者都对不上，"
                               "按规范应以单据张数计）" % (name, cnt, rows, codes)))
    return out


SPEC_METER = re.compile(r"([A-Za-z]*\s?\d+[\d\.\*]*)[-—]?\s*(\d+(?:\.\d+)?)\s*m", re.I)


def r_material(f):
    """材料核销单米数 vs 竣工验收单验收内容。"""
    out = []
    jg = f.get("jungong")
    items = f.hexiao_items()
    if not jg or not items:
        return [Finding(INFO, "无法核对材料米数", "缺核销单或竣工验收单")]
    content = ""
    t = jg.first_table()
    for row in t.rows:
        for c in row:
            if "m、" in c or "m；" in c or re.search(r"\d+m", c):
                if len(c) > len(content):
                    content = c
    if not content:
        return [Finding(INFO, "无法核对材料米数", "竣工验收单里没找到工程量描述")]

    def norm_material(name):
        n = name.upper()
        if "PE管" in name:
            m = re.search(r"DE?(\d+)", n)
            return ("PE", m.group(1) if m else "")
        if "无缝" in name:
            m = re.search(r"D(\d+)", n.replace("DN", "D"))
            return ("无缝", m.group(1) if m else "")
        if "镀锌" in name:
            m = re.search(r"DN(\d+)", n)
            return ("镀锌", m.group(1) if m else "")
        return (None, "")

    hexiao = {}
    for name, unit, qty in items:
        if unit not in ("米", "M", "m"):
            continue
        cat, spec = norm_material(name)
        if not cat:
            continue
        hexiao[(cat, spec)] = hexiao.get((cat, spec), 0) + qty

    jg_items = {}
    for m in re.finditer(r"(PE管|无缝钢管|热镀锌钢管|镀锌管)[：:]([^；;。]+)", content):
        cat_name, body = m.group(1), m.group(2)
        cat = "PE" if "PE" in cat_name else ("无缝" if "无缝" in cat_name else "镀锌")
        for piece in re.split(r"[、,，；;]", body):
            ms = re.search(r"(?:dn|DN|D)\s*(\d+)", piece)
            mq = re.search(r"(\d+(?:\.\d+)?)\s*m(?![a-zA-Z])", piece, re.I)
            if ms and mq:
                key = (cat, ms.group(1))
                jg_items[key] = jg_items.get(key, 0) + float(mq.group(1))

    for key in sorted(set(list(hexiao) + list(jg_items))):
        a = hexiao.get(key, 0)
        b = jg_items.get(key, 0)
        if abs(a - b) > 0.5:
            out.append(Finding(ERR, "米数对不上",
                               "%s d%s：核销单 %g 米，竣工验收单写 %g 米" % (key[0], key[1], a, b)))
    return out


def _mval(s):
    m = re.search(r"(\d+(?:\.\d+)?)", s.replace(" ", ""))
    return float(m.group(1)) if m else None


def r_tests(f):
    out = []
    short = []
    today = datetime.date.today().isoformat()
    for kind, label in (("qiangdu", "强度试验"), ("mishi", "气密试验")):
        bills = f.all(kind)
        if not bills:
            continue
        for b in bills:
            t = b.first_table()
            if not t:
                continue
            i_p = t.col("试验压力", "压力MPa")
            i_d = t.col("试验日期")
            i_v = t.col("有效期")
            recs = t.records()
            # 读数只升不降
            vals = []
            for row in recs:
                v = _mval(cell(row, i_p)) if i_p >= 0 else None
                if v is not None:
                    vals.append((v, cell(row, i_p)))
            for (v1, s1), (v2, s2) in zip(vals, vals[1:]):
                if v2 < v1 - 1e-9:
                    out.append(Finding(ERR, "试验压力下降",
                                       "「%s」读数从 %s 降到 %s" % (b.title, s1, s2)))
                    break
            if len(vals) < 3:
                short.append("%s(%d条)" % (b.title.replace("管道", "").replace("试验记录", ""), len(vals)))
            # 备注句
            if "发泡剂" not in b.all_text():
                out.append(Finding(ERR, "缺发泡剂检漏备注",
                                   "「%s」没有「已采用发泡剂检查所有焊口、接口进行检查无渗漏均为合格」"
                                   % b.title))
            # 压力表有效期
            if i_v >= 0:
                for row in recs:
                    d1 = parse_date_only(cell(row, i_d))
                    d2 = cell(row, i_v).strip()
                    if not d1 or not d2:
                        continue
                    m = re.match(r"(\d{2})[-./](\d{2})[-./](\d{2})", d2)
                    if m:
                        y = 2000 + int(m.group(1))
                        exp = "%04d-%s-%s" % (y, m.group(2), m.group(3))
                        if exp < d1:
                            out.append(Finding(ERR, "压力表过期",
                                               "「%s」试验日 %s，压力表有效期只到 %s" % (b.title, d1, exp)))
                            break
    if short:
        out.append(Finding(WARN, "试验节点不足三条",
                           "这些记录不足 3 条读数（试验前/中/结果）：" + "、".join(short[:8])
                           + ("等 %d 张" % len(short) if len(short) > 8 else "")))
    return out


def r_special(f):
    out = []
    b = f.get("tezhong")
    if not b:
        return [Finding(ERR, "缺特种工登记表", "全套资料里没有特种工登记表")]
    t = b.first_table()
    i_n, i_j, i_c, i_d = t.col("姓名"), t.col("作业岗位"), t.col("证件编号"), t.col("有效日期")
    rows = t.records()
    posts = [(cell(r, i_n), cell(r, i_j), cell(r, i_c), cell(r, i_d)) for r in rows]
    has_pe = bool(f.all("rean_jc") or f.all("dianrong_jc") or f.all("rean_cs"))
    has_galv = bool(f.get("luowen"))
    hexiao_pe = any("PE管" in n for n, u, q in f.hexiao_items())
    has_pe = has_pe or hexiao_pe
    if has_pe and not any("PE" in j.upper() for _, j, _, _ in posts):
        out.append(Finding(ERR, "特种工缺 PE 焊工", "项目有 PE 管，但特种工表里没有 PE 焊工"))
    if has_galv and not any("电工" in j or "焊工" in j and "PE" not in j.upper() for _, j, _, _ in posts):
        out.append(Finding(WARN, "特种工缺电工", "项目有镀锌管螺纹连接，规范要求登记电工"))
    if f.get("ganghan") and not any(("电焊" in j) or ("焊工" in j and "PE" not in j.upper()) for _, j, _, _ in posts):
        out.append(Finding(WARN, "特种工缺电焊工", "项目有钢制管道焊接，应登记电焊工"))
    # 证件有效期
    end = None
    jg = f.get("jungong")
    if jg:
        tt = jg.first_table()
        end = parse_date_only(tt.kv().get("验收时间", "")) or approval_last(jg)
    today = datetime.date.today().isoformat()
    for name, post, cid, exp in posts:
        d = parse_date_only(exp)
        if not d:
            continue
        if end and d < end[:10]:
            out.append(Finding(ERR, "证件有效期不足",
                               "%s（%s）证件 %s 有效期到 %s，早于工程验收 %s"
                               % (name, post, cid, d, end[:10])))
        elif d < today:
            out.append(Finding(WARN, "证件已过期",
                               "%s（%s）证件 %s 已于 %s 过期，出资料前需换证或换人"
                               % (name, post, cid, d)))
    # 人员表一致性
    ry = f.get("renyuan")
    if ry:
        tt = ry.first_table()
        i2 = tt.col("姓名")
        names = {cell(r, i2) for r in tt.records() if cell(r, i2)}
        miss = sorted({name for name, _, _, _ in posts if name and name not in names})
        if miss:
            out.append(Finding(WARN, "人员对不上",
                               "特种工登记表里的人没出现在项目人员信息表：%s" % "、".join(miss)))
    return out


def r_empty(f):
    """无对应管材却保留单据。"""
    out = []
    items = f.hexiao_items()
    if not items:                       # 没有核销单就无从判断材料有无，交给「无法核对米数」那条
        return out
    has_pe = any("PE管" in n or "PE" in n or "电熔" in n or "热熔" in n for n, u, q in items)
    has_galv = any("镀锌" in n for n, u, q in items)
    has_steel = any("无缝" in n for n, u, q in items)
    if not has_pe:
        for kind in ("rean_cs", "rean_jc", "dianrong_cs", "dianrong_jc"):
            b = f.get(kind)
            if b and len(rows_of(b)) > 0:
                out.append(Finding(ERR, "多余单据", "材料里没有 PE 管，却还有「%s」" % b.title))
    if not has_galv and f.get("luowen"):
        out.append(Finding(WARN, "多余单据", "材料里没有镀锌管，却还有螺纹连接检查记录"))
    return out


RULES = [
    ("结构完整性", r_structure),
    ("单据编号", r_codes),
    ("时间链", r_time_chain),
    ("节点清单", r_node_list),
    ("工序顺序", r_work_order),
    ("收尾顺序", r_finish_order),
    ("汇总表次数", r_summary),
    ("材料米数", r_material),
    ("试验记录", r_tests),
    ("特种工", r_special),
    ("空单据", r_empty),
]


def check(doc):
    f = Facts(doc)
    findings = []
    for name, fn in RULES:
        try:
            findings.extend(fn(f) or [])
        except Exception as exc:  # 单条规则崩了不影响其他
            findings.append(Finding(WARN, "检查中断", "「%s」这条规则执行出错：%s" % (name, exc)))
    order = {ERR: 0, WARN: 1, INFO: 2}
    findings.sort(key=lambda x: order.get(x.level, 3))
    return f, findings
