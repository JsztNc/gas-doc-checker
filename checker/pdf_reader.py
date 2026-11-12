# -*- coding: utf-8 -*-
"""读取文字型「线上资料.pdf」，按单据拆成与 docx_reader 相同的 items 结构。

PDF 不是 zip+XML，没有显式的单元格结构；这里用 PyMuPDF 做两件互补的事：
1. 表格：page.find_tables() 按表格线（lattice）或文本流（stream）识别表格，
   抽成 [行][列] 的单元格文本——直接对应 docx_reader 的 ("t", rows)。
2. 段落：文本块里“不在任何表格内”的部分，作为 ("p", text) 候选标题。

产出的 items 顺序按页面 y 坐标排列，下游 build_bills / classify / Table
完全复用 docx_reader 的逻辑，无需改动。
"""
import warnings

# PyMuPDF 2.x 把 `fitz` 标记为弃用，但仍是稳定别名；屏蔽该 FutureWarning 避免刷屏。
warnings.filterwarnings("ignore", category=FutureWarning)
import fitz  # PyMuPDF


def _in_table_rect(x0, y0, x1, y1, table_rects):
    """文本块中心是否落在某个表格矩形内（用于剔除表格单元格文字，避免重复）。"""
    cx, cy = (x0 + x1) / 2.0, (y0 + y1) / 2.0
    for tr in table_rects:
        if tr[0] <= cx <= tr[2] and tr[1] <= cy <= tr[3]:
            return True
    return False


def read_items(path):
    """返回 [("p", text) | ("t", [[cell, ...], ...]), ...]，保持文档顺序。"""
    doc = fitz.open(path)
    items = []
    try:
        for pg in doc:
            tf = pg.find_tables()
            tables = getattr(tf, "tables", []) or []
            table_rects = [t.bbox for t in tables]

            page_items = []  # (y0, kind, value)
            # 1) 文本块 → 段落（剔除落在表格内的）
            for b in pg.get_text("blocks"):
                x0, y0, x1, y1, txt = b[0], b[1], b[2], b[3], b[4]
                txt = (txt or "").strip()
                if not txt:
                    continue
                if _in_table_rect(x0, y0, x1, y1, table_rects):
                    continue
                page_items.append((y0, "p", txt))

            # 2) 表格 → 单元格矩阵
            for t in tables:
                try:
                    rows = t.extract()
                except Exception:
                    continue
                if not rows:
                    continue
                ncol = max((len(r) for r in rows), default=0)
                if ncol == 0:
                    continue
                # fitz 空单元格可能是 None，且单元格内换行会把编号切开，
                # 统一成 "" 并去掉单元格内换行（与 docx 解析一致，避免误判“编号带空格”）
                norm = [[(c or "").replace("\n", "").replace("\r", "") for c in r] +
                        [""] * (ncol - len(r)) for r in rows]
                # 跳过全是空行的表
                if not any(any(c for c in row) for row in norm):
                    continue
                page_items.append((t.bbox[1], "t", norm))

            page_items.sort(key=lambda x: x[0])
            for _, kind, val in page_items:
                items.append((kind, val))
    finally:
        doc.close()
    return items
