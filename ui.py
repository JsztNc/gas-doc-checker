# -*- coding: utf-8 -*-
"""线上资料检查 —— 原生窗口版（WinForms，无网页内核）。

输入工程编号 → 自动用 Everything 找「<编号>-线上资料.docx / .pdf」→ 对照规范检查 → 列问题。
"""
import json
import os
import subprocess
import sys
import threading
import time
import traceback

BASE = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False)
                                      else os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)

import updater                                          # noqa: E402
import feedback                                         # noqa: E402

if "--version" in sys.argv:
    # 便于离线核验打包进 exe 的真实版本号（不影响正常 GUI 运行）
    try:
        print("version", updater.local_version())
    except Exception:
        print("version unknown")
    sys.exit(0)

# ---------------------------------------------------------------- 无界面模式
if "--worker-check" in sys.argv:
    _i = sys.argv.index("--worker-check")
    _code, _out = sys.argv[_i + 1], sys.argv[_i + 2]
    _rest = sys.argv[_i + 3:]
    _path = _rest[_rest.index("--path") + 1] if "--path" in _rest else None
    _fix = "--fix" in _rest
    try:
        import check as _check
        _check.run(_code, path=_path, json_out=_out, quiet=True, do_fix=_fix)
    except RuntimeError as exc:
        # 预期内的业务错误（没搜到文件、文件打不开等）只回传消息本身，
        # 不把 traceback 弹给用户。
        with open(_out, "w", encoding="utf-8") as fh:
            json.dump({"ok": False, "error": str(exc), "kind": "business"},
                      fh, ensure_ascii=False)
    except BaseException:
        with open(_out, "w", encoding="utf-8") as fh:
            json.dump({"ok": False, "error": traceback.format_exc(limit=3), "kind": "crash"},
                      fh, ensure_ascii=False)
    sys.exit(0)

import clr                                            # noqa: E402
import System                                         # noqa: E402
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
from System.Windows.Forms import (                    # noqa: E402
    Form, Button, TextBox, Label, ListView, ColumnHeader, View, Application,
    FormBorderStyle, FormStartPosition, MessageBox, DialogResult, Panel,
    CharacterCasing, FlatStyle, ColumnHeaderAutoResizeStyle,
    MessageBoxButtons, MessageBoxIcon, ScrollBars, DockStyle, ComboBox, ComboBoxStyle)
from System.Drawing import Font, FontStyle, Color, Size, Point   # noqa: E402

FONT = Font("Microsoft YaHei UI", 10)
FONT_B = Font("Microsoft YaHei UI", 10, FontStyle.Bold)
FONT_T = Font("Microsoft YaHei UI", 15, FontStyle.Bold)
BLUE = Color.FromArgb(23, 119, 255)
LEVEL_COLOR = {"错误": Color.FromArgb(214, 48, 48),
               "警告": Color.FromArgb(214, 132, 0),
               "提示": Color.FromArgb(120, 126, 136)}


def mk_button(text, x, y, w, h, color=None, bold=True):
    b = Button()
    b.Text = text
    b.Location = Point(x, y)
    b.Size = Size(w, h)
    b.Font = FONT_B if bold else FONT
    b.FlatStyle = FlatStyle.Flat
    if color is not None:
        b.BackColor = color
        b.ForeColor = Color.White
    return b


def open_path(path):
    try:
        os.startfile(path)          # noqa: S606
    except Exception:
        pass


class App(object):
    def __init__(self):
        self.proc = None
        self.result = None
        self.fixing = False
        self.updating = False
        # 供「反馈」自动附带：最近一次操作的上下文
        self.context = {"code": "", "action": "（还没进行操作）", "detail": "",
                        "status": "", "time": ""}
        self.build()

    # ------------------------------------------------------------ 界面
    def build(self):
        f = Form()
        f.Text = "线上资料检查"
        f.ClientSize = Size(1000, 640)
        f.StartPosition = FormStartPosition.CenterScreen
        f.FormBorderStyle = FormBorderStyle.FixedSingle
        f.MaximizeBox = False
        f.Font = FONT
        self.form = f

        title = Label()
        title.Text = "线上燃气竣工资料检查"
        title.Font = FONT_T
        title.Location = Point(16, 14)
        title.Size = Size(420, 34)
        f.Controls.Add(title)

        sub = Label()
        sub.Text = "按《线上燃气竣工资料填写规范与校验要点》自动比对"
        sub.ForeColor = Color.Gray
        sub.Location = Point(20, 50)
        sub.Size = Size(460, 20)
        f.Controls.Add(sub)

        lbl = Label()
        lbl.Text = "工程编号"
        lbl.Location = Point(20, 84)
        lbl.Size = Size(72, 26)
        f.Controls.Add(lbl)

        self.txt = TextBox()
        self.txt.Location = Point(94, 80)
        self.txt.Size = Size(220, 28)
        self.txt.Font = Font("Consolas", 11)
        self.txt.CharacterCasing = CharacterCasing.Upper
        f.Controls.Add(self.txt)

        self.btn = mk_button("开始检查", 326, 78, 110, 32, BLUE)
        self.btn.Click += self.on_check
        f.Controls.Add(self.btn)

        self.btn_file = mk_button("选文件…", 444, 78, 88, 32, None, bold=False)
        self.btn_file.Click += self.on_pick_file
        f.Controls.Add(self.btn_file)

        self.btn_update = mk_button("检查更新", 856, 18, 118, 30, None, bold=False)
        self.btn_update.Click += self.on_update
        f.Controls.Add(self.btn_update)

        self.btn_help = mk_button("使用说明", 736, 18, 112, 30, None, bold=False)
        self.btn_help.Click += self.on_help
        f.Controls.Add(self.btn_help)

        self.btn_feedback = mk_button("反馈", 676, 18, 54, 30, None, bold=False)
        self.btn_feedback.Click += self.on_feedback
        f.Controls.Add(self.btn_feedback)

        self.lbl_version = Label()
        self.lbl_version.Text = "v%s" % updater.local_version()
        self.lbl_version.ForeColor = Color.Gray
        self.lbl_version.Location = Point(606, 24)
        self.lbl_version.Size = Size(66, 20)
        f.Controls.Add(self.lbl_version)

        self.status = Label()
        self.status.Text = "输入工程编号（如 NC23A005），回车即开始"
        self.status.ForeColor = Color.Gray
        self.status.Location = Point(544, 84)
        self.status.Size = Size(430, 24)
        f.Controls.Add(self.status)

        self.list = ListView()
        self.list.View = View.Details
        self.list.FullRowSelect = True
        self.list.GridLines = False
        self.list.HideSelection = False
        self.list.Location = Point(16, 124)
        self.list.Size = Size(966, 452)
        self.list.Font = FONT
        self.list.Columns.Add("级别", 60)
        self.list.Columns.Add("检查项", 180)
        self.list.Columns.Add("问题说明", 700)
        f.Controls.Add(self.list)

        self.lbl_file = Label()
        self.lbl_file.Text = ""
        self.lbl_file.ForeColor = Color.Gray
        self.lbl_file.Location = Point(16, 588)
        self.lbl_file.Size = Size(516, 22)
        f.Controls.Add(self.lbl_file)

        self.btn_fix = mk_button("自动修正", 636, 584, 104, 30, Color.FromArgb(23, 161, 90))
        self.btn_fix.Click += self.on_fix
        f.Controls.Add(self.btn_fix)

        self.btn_open = mk_button("打开文档", 536, 584, 92, 30, None, bold=False)
        self.btn_open.Click += lambda s, a: open_path(self.result["path"]) if self.result else None
        f.Controls.Add(self.btn_open)

        self.btn_report = mk_button("打开报告", 748, 584, 92, 30, None, bold=False)
        self.btn_report.Click += lambda s, a: open_path(self.result["report"]) if self.result else None
        f.Controls.Add(self.btn_report)

        self.btn_folder = mk_button("报告文件夹", 846, 584, 92, 30, None, bold=False)
        self.btn_folder.Click += lambda s, a: open_path(os.path.join(BASE, "检查报告"))
        f.Controls.Add(self.btn_folder)

        self.txt.KeyDown += self.on_key
        f.AcceptButton = self.btn
        f.Icon = self.load_icon()
        return f

    def load_icon(self):
        try:
            from System.Drawing import Icon as DIcon
            p = os.path.join(BASE, "icon.ico")
            if os.path.isfile(p):
                return DIcon(p)
        except Exception:
            pass
        return None

    # ------------------------------------------------------------ 事件
    def on_key(self, sender, args):
        if args.KeyCode == 13:
            self.on_check(None, None)

    def say(self, text, color=None):
        self.status.Text = text
        self.status.ForeColor = color or Color.Gray
        self.context["status"] = text

    def set_context(self, code=None, action=None, detail=None):
        """记录最近操作，供「反馈」自动附带。"""
        if code is not None:
            self.context["code"] = code
        elif hasattr(self, "txt") and self.txt is not None:
            self.context["code"] = (self.txt.Text or "").strip().upper()
        if action is not None:
            self.context["action"] = action
            self.context["time"] = time.strftime("%Y-%m-%d %H:%M:%S")
        if detail is not None:
            self.context["detail"] = detail

    def on_pick_file(self, sender, args):
        try:
            from System.Windows.Forms import OpenFileDialog
            d = OpenFileDialog()
            d.Filter = "线上资料 (*.docx;*.pdf)|*.docx;*.pdf"
            if d.ShowDialog(self.form) == DialogResult.OK:
                code = os.path.basename(d.FileName).split("-")[0]
                self.txt.Text = code
                self.set_context(code=code, action="选文件", detail=d.FileName)
                self.start(code, d.FileName)
        except Exception:
            self.say("选文件出错：%s" % traceback.format_exc(limit=1), LEVEL_COLOR["错误"])

    # ------------------------------------------------------------ 使用说明
    def on_help(self, sender, args):
        dlg = Form()
        dlg.Text = "使用说明"
        dlg.ClientSize = Size(760, 620)
        dlg.StartPosition = FormStartPosition.CenterScreen
        dlg.FormBorderStyle = FormBorderStyle.FixedSingle
        dlg.MaximizeBox = False
        dlg.Font = FONT
        txt = TextBox()
        txt.Multiline = True
        txt.ReadOnly = True
        txt.ScrollBars = ScrollBars.Both
        txt.WordWrap = False
        txt.Dock = DockStyle.Fill
        txt.Font = Font("Microsoft YaHei UI", 10)
        txt.Text = updater.help_text()
        txt.SelectionStart = 0
        dlg.Controls.Add(txt)
        dlg.ShowDialog(self.form)

    # ------------------------------------------------------------ 反馈
    def on_feedback(self, sender, args):
        dlg = Form()
        dlg.Text = "意见反馈 (v%s)" % updater.local_version()
        dlg.ClientSize = Size(560, 460)
        dlg.StartPosition = FormStartPosition.CenterScreen
        dlg.FormBorderStyle = FormBorderStyle.FixedSingle
        dlg.MaximizeBox = False
        dlg.Font = FONT

        lbl1 = Label()
        lbl1.Text = "反馈类型"
        lbl1.Location = Point(20, 18)
        lbl1.Size = Size(80, 24)
        dlg.Controls.Add(lbl1)

        combo = ComboBox()
        combo.Location = Point(100, 14)
        combo.Size = Size(180, 28)
        combo.DropDownStyle = ComboBoxStyle.DropDownList
        for item in ("问题", "建议", "其他"):
            combo.Items.Add(item)
        combo.SelectedIndex = 0
        dlg.Controls.Add(combo)

        lbl2 = Label()
        lbl2.Text = "反馈内容（必填）"
        lbl2.Location = Point(20, 52)
        lbl2.Size = Size(200, 24)
        dlg.Controls.Add(lbl2)

        body = TextBox()
        body.Multiline = True
        body.Location = Point(20, 78)
        body.Size = Size(520, 210)
        body.ScrollBars = ScrollBars.Vertical
        body.Font = Font("Microsoft YaHei UI", 10)
        dlg.Controls.Add(body)

        lbl3 = Label()
        lbl3.Text = "联系方式（选填，方便回复你）"
        lbl3.Location = Point(20, 298)
        lbl3.Size = Size(300, 24)
        dlg.Controls.Add(lbl3)

        contact = TextBox()
        contact.Location = Point(20, 324)
        contact.Size = Size(520, 28)
        dlg.Controls.Add(contact)

        tip = Label()
        tip.Text = "匿名提交到 GitHub，无需登录；提交后可在仓库 Issues 里看到。"
        tip.ForeColor = Color.Gray
        tip.Location = Point(20, 360)
        tip.Size = Size(520, 22)
        dlg.Controls.Add(tip)

        btn_send = mk_button("提交反馈", 20, 392, 110, 34, BLUE)
        btn_cancel = mk_button("取消", 142, 392, 80, 34, None, bold=False)
        dlg.Controls.Add(btn_send)
        dlg.Controls.Add(btn_cancel)
        btn_cancel.Click += lambda s, a: dlg.Close()

        def do_send(s, a):
            feedback._dbg("do_send enter (UI thread)")
            kind = str(combo.Text or "其他")
            text = body.Text or ""
            # 关键：所有控件值都在【UI 线程】这里读好，作为普通字符串传入后台线程。
            # 绝不能在后台线程里再访问 WinForms 控件（如 contact.Text）——那会触发
            # WinForms 跨线程控件访问，Release 模式下把读取同步 marshal 回 UI 线程，
            # 而此刻 UI 线程往往还卡在 do_send 内没回到消息泵，造成双方死锁 + 窗口「未响应」。
            contact_text = contact.Text or ""
            ctx = dict(self.context)
            if not text.strip():
                MessageBox.Show(dlg, "请先写点反馈内容", "意见反馈",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)
                return
            btn_send.Enabled = False
            btn_send.Text = "提交中…"

            # 结果放在线程间共享的普通容器里，由 UI 线程的 Timer 轮询取回，
            # 彻底避免「后台线程回调 -> 主窗体 BeginInvoke -> 模态对话框不派发」
            # 导致的窗口卡死（Windows 标记为「未响应」）。
            result = {}

            def work():
                feedback._dbg("thread started")
                try:
                    result["ok"], result["msg"] = feedback.submit(
                        kind, text, contact_text, context=ctx)
                except Exception as exc:
                    result["ok"], result["msg"] = False, "提交失败：%s" % exc
                feedback._dbg("thread done ok=%s" % result.get("ok"))
            # 关键：必须用 .NET 原生线程（System.Threading.Thread），不能用 Python 的
            # threading.Thread。从 WinForms 回调里起 threading.Thread 会触发 pythonnet 的
            # GIL 死锁——主线程卡死（界面「未响应」），新线程也永远拿不到 GIL，work() 根本不执行。
            # 本项目 main 入口（见文件末尾）同样用 System.Threading.Thread 规避此问题。
            _th = System.Threading.Thread(System.Threading.ThreadStart(work))
            _th.IsBackground = True
            _th.Start()
            feedback._dbg("thread launched, UI thread returned")

            def tick(sender, ev):
                if not result:                 # 后台还没返回
                    return
                timer.Stop()
                feedback._dbg("tick fired, showing result")
                if dlg.IsDisposed or not dlg.IsHandleCreated:
                    return                     # 用户已关掉对话框，不再弹窗
                btn_send.Enabled = True
                btn_send.Text = "提交反馈"
                MessageBox.Show(
                    dlg, result["msg"], "意见反馈", MessageBoxButtons.OK,
                    MessageBoxIcon.Information if result["ok"] else MessageBoxIcon.Error)
                if result["ok"]:
                    dlg.Close()

            timer = System.Windows.Forms.Timer()
            timer.Interval = 200
            timer.Tick += tick
            timer.Start()
        btn_send.Click += do_send

        dlg.ShowDialog(self.form)

    # ------------------------------------------------------------ 检查更新
    def _ui(self, fn):
        try:
            self.form.BeginInvoke(System.Action(fn))
        except Exception:
            pass

    def on_update(self, sender, args):
        if self.updating:
            return
        self.updating = True
        self.btn_update.Enabled = False
        self.set_context(action="检查更新", detail="查询 GitHub Releases 最新版本")
        self.say("正在检查 GitHub 上的新版本…", BLUE)
        threading.Thread(target=self._check_update_worker, daemon=True).start()

    def _check_update_worker(self):
        try:
            info = updater.check_latest()
        except BaseException as exc:
            info = {"ok": False, "error": str(exc)}
        self._ui(lambda: self._update_checked(info))

    def _update_checked(self, info):
        self.btn_update.Enabled = True
        self.updating = False
        if not info.get("ok"):
            msg = info.get("error") or "未知错误"
            self.say("检查更新失败：%s" % msg.split("\n")[0][:64], LEVEL_COLOR["错误"])
            ask = "检查更新失败：%s\n\n是否打开 releases 页面自己看看？" % msg
            if MessageBox.Show(self.form, ask, "检查更新",
                               MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes:
                updater.open_releases()
            return
        if not info.get("has_update"):
            self.say("已是最新版本 v%s" % info["local"], Color.FromArgb(23, 161, 90))
            MessageBox.Show(self.form, "当前已是最新版本 v%s。" % info["local"],
                            "检查更新", MessageBoxButtons.OK, MessageBoxIcon.Information)
            return
        notes = (info.get("notes") or "").strip()
        if len(notes) > 480:
            notes = notes[:480] + " …"
        msg = "发现新版本 v%s（当前 v%s，发布于 %s）\n\n" % (
            info["latest"], info["local"], info.get("published_at") or "-")
        if notes:
            msg += "更新说明：\n%s\n\n" % notes
        if info.get("asset_url"):
            msg += "是否现在下载并自动更新？\n（程序会关闭几秒，然后重新打开；检查报告不会动）"
        else:
            msg += "这个版本没有提供 zip 整包，只能去 releases 页面手动下载。"
        if MessageBox.Show(self.form, msg, "发现新版本",
                           MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes:
            return
        if info.get("asset_url"):
            self._download_and_apply(info)
        else:
            updater.open_releases()

    def _download_and_apply(self, info):
        self.updating = True
        self.btn_update.Enabled = False
        self.say("正在下载 v%s…" % info["latest"], BLUE)

        def progress(done, total):
            txt = ("%d%%" % (done * 100 // total)) if total else ("%d KB" % (done // 1024))
            self._ui(lambda: self.say("正在下载 v%s… %s" % (info["latest"], txt), BLUE))

        def work():
            try:
                path = updater.download_release(info["asset_url"], on_progress=progress)
                bat, tip = updater.prepare_update(path, info["latest"])
                updater.start_update(bat)
                self._ui(lambda: self.say(tip, Color.FromArgb(23, 161, 90)))
                self._ui(lambda: Application.Exit())
            except BaseException as exc:
                err = "更新失败：%s" % exc
                self._ui(lambda: self.say(err.split("\n")[0][:80], LEVEL_COLOR["错误"]))
                self._ui(lambda: MessageBox.Show(self.form, err, "更新失败",
                                                 MessageBoxButtons.OK, MessageBoxIcon.Error))
                self._ui(lambda: setattr(self.btn_update, "Enabled", True))
                self.updating = False

        threading.Thread(target=work, daemon=True).start()

    def on_check(self, sender, args):
        code = (self.txt.Text or "").strip().upper()
        if not code:
            self.say("先输入工程编号", LEVEL_COLOR["错误"])
            return
        self.start(code, None)

    def on_fix(self, sender, args):
        if self.proc is not None and self.proc.poll() is None:
            self.say("上一次还在跑，稍等…", LEVEL_COLOR["警告"])
            return
        path = (self.result or {}).get("path")
        code = (self.txt.Text or "").strip().upper()
        if not path and not code:
            MessageBox.Show(self.form, "先输入工程编号并检查一次，再点自动修正", "自动修正")
            return
        ask = ("将按规范自动修正这份资料（原文件先备份到同目录「_改前备份」）：\n\n"
               "· 单据编号：去掉空格、按工序重排流水、消除重复\n"
               "· 项目编号：整套统一为同一个编号\n"
               "· 工序质量汇总表：检查次数按实际单据张数回写\n"
               "· 试验记录：备注栏补「发泡剂检漏」标准句\n"
               "· 项目人员信息表：把特种工名单补齐\n\n"
               "注：只改格式 / 编号 / 次数 / 统一文本这类机械错误；时间先后、米数取舍、"
               "缺哪张单、照片与签字盖章等需判断的，仍要人工。要开始吗？")
        if MessageBox.Show(self.form, ask, "自动修正",
                           MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes:
            return
        if code:
            self.start(code, path, do_fix=True)
        else:
            self.say("先输入工程编号", LEVEL_COLOR["错误"])

    def start(self, code, path, do_fix=False):
        if self.proc is not None and self.proc.poll() is None:
            self.say("上一次还在跑，稍等…", LEVEL_COLOR["警告"])
            return
        self.set_context(code=code,
                         action="自动修正" if do_fix else "开始检查",
                         detail=("指定文件：%s" % path) if path else "按编号用 Everything 搜索")
        self.list.Items.Clear()
        self.result = None
        self.lbl_file.Text = ""
        self.btn.Enabled = False
        self.fixing = do_fix
        self.say("正在按规范修正…（改完会自动重新检查）" if do_fix
                 else "正在用 Everything 搜索并解析文档…（大文件要十几秒）", BLUE)
        out = os.path.join(BASE, "_tmp")
        os.makedirs(out, exist_ok=True)
        self.json_path = os.path.join(out, "check_%s.json" % code)
        if os.path.exists(self.json_path):
            try:
                os.remove(self.json_path)
            except OSError:
                pass
        if getattr(sys, "frozen", False):
            cmd = [sys.executable, "--worker-check", code, self.json_path]
        else:
            cmd = [sys.executable, os.path.join(BASE, "check.py"), code, "--json", self.json_path]
        if path:
            cmd += ["--path", path]
        if do_fix:
            cmd.append("--fix")
        try:
            self.proc = subprocess.Popen(cmd, cwd=BASE, stdout=subprocess.DEVNULL,
                                         stderr=subprocess.DEVNULL)
        except Exception as exc:
            self.say("启动检查进程失败：%s" % exc, LEVEL_COLOR["错误"])
            self.btn.Enabled = True
            return
        threading.Thread(target=self.wait, daemon=True).start()

    def wait(self):
        t0 = time.time()
        while self.proc.poll() is None:
            time.sleep(0.3)
            if time.time() - t0 > 300:
                try:
                    self.proc.kill()
                except Exception:
                    pass
                break
        time.sleep(0.2)
        data = None
        try:
            with open(self.json_path, "r", encoding="utf-8") as fh:
                data = json.load(fh)
        except Exception:
            data = {"ok": False, "error": "检查进程没有返回结果（可能崩溃了）"}
        try:
            self.form.BeginInvoke(System.Action(lambda: self.done(data)))
        except Exception:
            pass

    def done(self, data):
        self.btn.Enabled = True
        if not data or not data.get("ok"):
            msg = (data or {}).get("error") or "未知错误"
            self.say(msg.split("\n")[0][:80], LEVEL_COLOR["错误"])
            self.set_context(detail="操作失败：" + msg.split("\n")[0][:120])
            if (data or {}).get("kind") == "business":
                tip = (msg + "\n\n"
                       "可以点「选文件…」直接指定那份 docx / pdf；\n"
                       "或确认：编号没写错、Everything 在运行、共享盘已连上。")
                MessageBox.Show(self.form, tip, "检查没成功",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)
            else:
                MessageBox.Show(self.form, msg, "检查出错")
            return
        self.result = data
        counts = data.get("counts") or {}
        err, warn, info = counts.get("错误", 0), counts.get("警告", 0), counts.get("提示", 0)
        self.set_context(
            detail="检查完成：错误 %d / 警告 %d / 提示 %d；文件：%s"
                   % (err, warn, info, data.get("path") or "-"))
        self.lbl_file.Text = data.get("path", "")
        self.list.BeginUpdate()
        for item in data.get("findings") or []:
            lv = item.get("level", "")
            row = self.list.Items.Add(lv)
            row.SubItems.Add(item.get("rule", ""))
            row.SubItems.Add(item.get("msg", ""))
            row.ForeColor = LEVEL_COLOR.get(lv, Color.Black)
        self.list.EndUpdate()
        if err:
            self.say("错误 %d 条，警告 %d 条 —— 先改错误项" % (err, warn), LEVEL_COLOR["错误"])
        elif warn:
            self.say("没有错误，警告 %d 条" % warn, LEVEL_COLOR["警告"])
        else:
            self.say("没发现明显问题（提示 %d 条）" % info, Color.FromArgb(23, 161, 90))
        self.list.Columns[2].Width = 700
        if getattr(self, "fixing", False) and data.get("fix"):
            fx = data["fix"]
            self.fixing = False
            who = "、".join(x[0] for x in fx.get("people") or []) or "无"
            parts = []
            if fx.get("codes"):
                parts.append("· 单据编号 %d 处（去空格、重排流水、消除重复）" % fx["codes"])
            if fx.get("pid"):
                parts.append("· 统一项目编号 %d 处" % fx["pid"])
            if fx.get("summary"):
                parts.append("· 修正汇总表检查次数 %d 处" % fx["summary"])
            if fx.get("note"):
                parts.append("· 补全发泡剂检漏备注 %d 张" % fx["note"])
            if fx.get("members"):
                parts.append("· 补充人员 %d 人（%s）" % (fx["members"], who))
            if not parts:
                parts.append("· 没有可自动修正的问题")
            MessageBox.Show(self.form,
                            "已自动修正：\n\n" + "\n".join(parts)
                            + "\n\n改前备份：%s" % (fx.get("backup") or "-"),
                            "自动修正完成", MessageBoxButtons.OK, MessageBoxIcon.Information)


def main():
    Application.EnableVisualStyles()
    app = App()
    if "--demo" in sys.argv:                      # 截图/演示用假数据
        app.txt.Text = "NC23A005"
        app.done({"ok": True, "path": r"Z://工作文件//2025//竣工资料更改//第二批//NC23A005-b.线上资料.docx",
                  "counts": {"错误": 3, "警告": 2, "提示": 1},
                  "findings": [
                      {"level": "错误", "rule": "工序时间倒流", "msg": "隐蔽（2024-08-20 15:41）早于 连接/焊接完成（2024-08-25 14:29）"},
                      {"level": "错误", "rule": "米数对不上", "msg": "镀锌 d50：核销单 90 米，竣工验收单写 54 米"},
                      {"level": "错误", "rule": "单据编号带空格", "msg": "「沟槽回填检查记录」里编号写成「NC23A005GX0 009」"},
                      {"level": "警告", "rule": "汇总次数对不上", "msg": "汇总表「PE管电熔检查记录」写 1 次；前面明细 35 条、单据 35 个"},
                      {"level": "警告", "rule": "证件已过期", "msg": "沈克成（电工）证件 赣W012023279011 已于 2025-12-04 过期"},
                      {"level": "提示", "rule": "竣工审批不是最后", "msg": "「服务商材料使用量单」审批晚于竣工验收审批"}]})
    if "--code" in sys.argv:
        _code = sys.argv[sys.argv.index("--code") + 1]
        app.txt.Text = _code
        app.form.Shown += lambda s, a: app.start(_code, None)
    if "--shot" in sys.argv:
        _png = sys.argv[sys.argv.index("--shot") + 1]
        _delay = 2.5
        if "--shot-delay" in sys.argv:
            _delay = float(sys.argv[sys.argv.index("--shot-delay") + 1])

        def _snap():
            from System.Drawing import Bitmap, Rectangle
            time.sleep(_delay)
            bmp = Bitmap(app.form.Width, app.form.Height)
            app.form.DrawToBitmap(bmp, Rectangle(0, 0, app.form.Width, app.form.Height))
            bmp.Save(_png)
            for f in list(Application.OpenForms):
                f.Close()

        _th = threading.Thread(target=_snap, daemon=True)
        app.form.Shown += lambda s, a: _th.start()
    Application.Run(app.form)


if __name__ == "__main__":
    try:
        import System.Threading as _ST
        _t = _ST.Thread(_ST.ThreadStart(main))
        _t.SetApartmentState(_ST.ApartmentState.STA)    # 文件夹/文件对话框需要 STA
        _t.Start()
        _t.Join()
    except Exception:
        traceback.print_exc()
        main()
