# -*- coding: utf-8 -*-
"""问题定位 + 文档截图预览 + 单条内容改写。

做法（不需要额外装任何东西，只要本机有 Word）：
  1. 用 Word 后台打开那份 docx（只读、不显示窗口）
  2. 按问题描述里挑出的关键字（单据编号 / 日期）在全文里找
  3. 把命中位置所在的整张表格「复制为图片」，再从剪贴板取 EMF 渲染成 PNG
  4. 改写时用 Word 的查找替换，格式、字体、表格线都不动；改前先备份

对外三个函数：
  needle_of(msg)                  → 从问题描述里挑出可在文档中搜索的文字
  render(path, needle, out_png)   → 出截图，返回 dict(ok/found/page/note/png)
  replace_all(path, old, new)     → 查找替换并保存，返回 dict(ok/count/backup)
"""
import datetime
import os
import queue
import re
import shutil
import threading
import time

# 注意：不能在这里 import checker —— 多文件版把 checker 放在 exe 外（可改），
# 打包时被排除，import 阶段会炸；改成用的时候再取，取不到就用内置的同款正则。
_CODE_RE_FALLBACK = re.compile(
    r"[A-Z]{2}\d{2}[-A-Z0-9]{2,8}?(?:GX|HS|JD|BS|KGBG|AQ|CL|JG|WT)\d*", re.I)


def _code_re():
    try:
        from checker.docx_reader import CODE_RE
        return CODE_RE
    except Exception:
        return _CODE_RE_FALLBACK

_PNG_W = 1080          # 预览图目标宽度（像素）
_APP = None            # 复用同一个 Word 实例，避免每次预览都重启 Word
_JOBS = queue.Queue()
_THREAD = None


# ------------------------------------------------------------------ 工具
def needle_of(msg):
    """从问题描述里挑一段最适合在文档里搜的文字。
    优先级：单据编号 > 完整日期时间 > 年月日 > 纯数字串。找不到返回空串。"""
    if not msg:
        return ""
    # 「编号中间夹了空格」那种：把整串（含空格）取出来，用户才好直接改掉空格
    m = re.search(r"[A-Z]{2}\d{2}[A-Z0-9-]{2,}\s+[A-Z0-9]{2,}", msg)
    if m:
        return m.group(0)
    m = _code_re().search(msg)
    if m:
        return m.group(0)
    m = re.search(r"\d{4}-\d{2}-\d{2}(?:\s+\d{1,2}:\d{2})?", msg)
    if m:
        return m.group(0)
    m = re.search(r"\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日", msg)
    if m:
        return m.group(0)
    m = re.search(r"\bGX\d{3,4}\b", msg)
    if m:
        return m.group(0)
    return ""


def _get_app():
    global _APP
    if _APP is None:
        import win32com.client as win32
        _APP = win32.DispatchEx("Word.Application")
        _APP.Visible = False
        _APP.DisplayAlerts = 0
    return _APP


def shutdown():
    """关掉后台 Word（程序退出、或连着几分钟没用时调用）。"""
    global _APP
    try:
        if _APP is not None:
            _APP.Quit()
    except Exception:
        pass
    _APP = None


# ------------------------------------------------- 后台单线程（Word 只能在一个线程里用）
def _loop():
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pass
    try:
        while True:
            job = _JOBS.get()
            if job is None:
                return
            box, fn, args, kwargs = job
            try:
                box["r"] = fn(*args, **kwargs)
            except BaseException as exc:            # noqa: BLE001
                box["r"] = {"ok": False, "error": str(exc).split("\n")[0][:200]}
    finally:
        shutdown()
        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _ensure_worker():
    global _THREAD
    if _THREAD is None or not _THREAD.is_alive():
        _THREAD = threading.Thread(target=_loop, name="doc-preview", daemon=True)
        _THREAD.start()
    return _THREAD


def submit(fn, *args, **kwargs):
    """把一件 Word 活丢给后台线程，立刻返回一个 box；界面用 Timer 轮询 box['r']。

    Word 的 COM 对象不能跨线程用，所以所有 Word 操作都排在这一个线程里做；
    进程里只起一次 Word，后续每次预览 0.5~1 秒。"""
    _ensure_worker()
    box = {}
    _JOBS.put((box, fn, args, kwargs))
    return box


def stop_worker():
    """退出程序前调用：让后台线程收工并关掉 Word。"""
    global _THREAD
    try:
        _JOBS.put(None)
    except Exception:
        pass
    _THREAD = None


def _open(path, read_only=True):
    app = _get_app()
    doc = app.Documents.Open(FileName=os.path.abspath(path), ReadOnly=read_only,
                             AddToRecentFiles=False, Visible=False)
    return doc


def _friendly(exc, path):
    txt = str(exc)
    if "0x800A1436" in txt or "被占用" in txt or "locked" in txt.lower():
        return "这份文件正被 Word/WPS 打开着，先关掉它再试。\n（文件：%s）" % path
    if "0x800A1524" in txt:
        return "Word 弹了个对话框拦住了操作，请先在 Word 里正常打开一次这份文件再试。"
    return "调用 Word 失败：%s" % txt.split("\n")[0][:160]


# ------------------------------------------------------------------ 出图
def render(path, needle, out_png, width=_PNG_W):
    """把 needle 命中位置所在的表格截成图片。返回 dict。"""
    import clr  # noqa: F401
    import System
    clr.AddReference("System.Drawing")
    import System.Drawing as SD
    import win32clipboard
    import win32con

    out = {"ok": False, "png": "", "found": "", "page": "", "note": "", "error": ""}
    if not path or not os.path.isfile(path):
        out["error"] = "找不到文件：%s" % path
        return out
    if not path.lower().endswith((".docx", ".doc", ".docm")):
        out["error"] = "这份是 PDF，没法截图预览也没法就地修改；PDF 只能看检查报告。"
        return out

    doc = None
    emf_path = os.path.splitext(out_png)[0] + ".emf"
    try:
        doc = _open(path, read_only=True)
        target = None
        if needle:
            rng = doc.Content
            # 同样用位置参数（FindText, MatchCase, MatchWholeWord, MatchWildcards,
            # MatchSoundsLike, MatchAllWordForms, Forward, Wrap, Format, ReplaceWith, Replace）
            ok = rng.Find.Execute(needle, False, False, False, False, False, True, 0, False, "", 0)
            if ok:
                out["found"] = (rng.Text or "").strip()
                try:
                    rng.HighlightColorIndex = 7          # wdYellow，只改内存不存盘
                except Exception:
                    pass
                try:
                    out["page"] = int(doc.Range(0, rng.End).Information(3))   # 页码
                except Exception:
                    pass
                try:
                    if rng.Tables.Count > 0:
                        target = rng.Tables(1).Range
                    else:
                        target = rng.Paragraphs(1).Range
                except Exception:
                    target = rng
        if target is None:
            # 没定位到：给文档开头的表格/段落兜底，让用户至少看到上下文
            try:
                target = doc.Tables(1).Range if doc.Tables.Count > 0 \
                    else doc.Paragraphs(1).Range
                out["note"] = ("没在文档里搜到「%s」，先显示文档开头" % needle
                               if needle else "这条没有可定位的文字，先显示文档开头")
            except Exception:
                out["error"] = "这份文档里没有可显示的表格或段落。"
                return out

        target.CopyAsPicture()
        time.sleep(0.45)

        win32clipboard.OpenClipboard()
        try:
            data = win32clipboard.GetClipboardData(win32con.CF_ENHMETAFILE)
        finally:
            win32clipboard.CloseClipboard()
        if not data:
            out["error"] = "没取到截图数据（剪贴板被别的程序占了？）"
            return out
        with open(emf_path, "wb") as fh:
            fh.write(data)

        img = SD.Image.FromFile(emf_path)
        w = max(320, min(width, img.Width))
        h = max(80, int(img.Height * (float(w) / img.Width)))
        if h > 6000:                                  # 超长表格压一压，免得内存爆
            w = max(320, int(w * 6000.0 / h))
            h = 6000
        bmp = SD.Bitmap(w, h)
        g = SD.Graphics.FromImage(bmp)
        g.Clear(SD.Color.White)
        try:
            g.InterpolationMode = SD.Drawing2D.InterpolationMode.HighQualityBicubic
            g.SmoothingMode = SD.Drawing2D.SmoothingMode.HighQuality
        except Exception:
            pass
        g.DrawImage(img, SD.Rectangle(0, 0, w, h))
        g.Dispose()
        bmp.Save(out_png, SD.Imaging.ImageFormat.Png)
        bmp.Dispose()
        img.Dispose()
        out["png"] = out_png
        out["ok"] = True
        return out
    except Exception as exc:
        out["error"] = _friendly(exc, path)
        return out
    finally:
        try:
            if doc is not None:
                doc.Close(False)
        except Exception:
            pass
        try:
            if os.path.isfile(emf_path):
                os.remove(emf_path)
        except Exception:
            pass


# ------------------------------------------------------------------ 改内容
def replace_all(path, old, new, backup=True):
    """把文档里所有 old 替换成 new 并保存。返回 dict(ok/count/backup/error)。"""
    out = {"ok": False, "count": 0, "backup": "", "error": ""}
    old, new = (old or "").strip(), (new or "")
    if not old:
        out["error"] = "没有要替换的原文。"
        return out
    if old == new.strip():
        out["error"] = "改后内容和原来一样，不用保存。"
        return out
    if not path or not os.path.isfile(path):
        out["error"] = "找不到文件：%s" % path
        return out

    bak = ""
    doc = None
    try:
        if backup:
            bdir = os.path.join(os.path.dirname(os.path.abspath(path)), "_改前备份")
            os.makedirs(bdir, exist_ok=True)
            stamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
            bbase = os.path.splitext(os.path.basename(path))[0]
            bak = os.path.join(bdir, "%s.%s.docx" % (bbase, stamp))
            shutil.copy2(path, bak)

        doc = _open(path, read_only=False)
        # 数一遍原文出现几次（用全文文本算，不动文档结构）
        try:
            cnt = (doc.Content.Text or "").count(old)
        except Exception:
            cnt = 0
        if cnt == 0:
            # 全文文本里没有 → 再用 Word 找一次（防跨段落被拼在一起的情况）
            probe = doc.Content
            if not probe.Find.Execute(old, False, False, False, False, False, True, 0, False, "", 0):
                out["error"] = ("文档里找不到「%s」，可能这段字在图片、页眉或文本框里，"
                                "只能手工在 Word 里改。" % old)
                return out
        # 位置参数调用：pywin32 对 Find.Execute 的关键字参数映射不可靠，
        # 用命名参数会「返回 True 但内容没变」（这个坑踩过）
        # FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike,
        # MatchAllWordForms, Forward, Wrap, Format, ReplaceWith, Replace(wdReplaceAll)
        f = doc.Content.Find
        f.ClearFormatting()
        f.Replacement.ClearFormatting()
        f.Execute(old, False, False, False, False, False, True, 1, False, new, 2)
        try:
            left = (doc.Content.Text or "").count(old)
        except Exception:
            left = 0
        doc.Save()
        if left >= cnt and cnt:
            out["ok"] = False
            out["error"] = ("Word 执行了替换但文档内容没变化（原文可能被拆成了多段），"
                            "建议点「在 Word 里打开」手工核对。")
            return out
        out["ok"] = True
        out["count"] = max(1, cnt - left)
        out["backup"] = bak
        if left:
            out["warn"] = "还有 %d 处没改到（原文在文档里被拆成了多段），请核对。" % left
        return out
    except Exception as exc:
        out["error"] = _friendly(exc, path)
        if bak and os.path.isfile(bak):          # 改失败就把备份留着，至少原件还在
            out["backup"] = bak
        return out
    finally:
        try:
            if doc is not None:
                doc.Close(False)
        except Exception:
            pass
