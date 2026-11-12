# -*- coding: utf-8 -*-
"""用户反馈模块 —— 一键匿名提交到 GitHub Issues。

使用一枚最小权限的 fine-grained token（仅本仓库 Issues: write），
随程序打包分发。使用者无需 GitHub 账号，点了就提交。

反馈以 Issue 形式存到仓库，标题前缀统一为 [用户反馈]，便于 AI 后续提取：
    [用户反馈] <类型> - <简述>

正文用固定结构，便于机器解析和人工阅读：
    ### 反馈类型
    ...
    ### 反馈内容
    ...
    ### 联系方式（可选）
    ...
    ### 环境信息
    ...（自动附带：版本、系统、时间）
"""
import datetime
import json
import os
import sys
import urllib.error
import urllib.request

BASE = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False)
                                      else os.path.abspath(__file__)))

GITHUB_REPO = "JsztNc/gas-doc-checker"
API_ISSUES = "https://api.github.com/repos/%s/issues" % GITHUB_REPO
ISSUES_PAGE = "https://github.com/%s/issues?q=label%%3Auser-feedback" % GITHUB_REPO

# ------------------------------------------------------------------
# 最小权限 token（fine-grained PAT：仅 gas-doc-checker 仓库，Issues: write）
# 明文凭证绝不能写进源码/仓库（GitHub secret scanning 会拦截推送）。
# 改为从同目录的 feedback_token.txt 读取；该文件已 gitignore，打包时作为
# 资源嵌进 exe。只在开发者本机的源码目录里才有这个 txt。
TOKEN_FILE = "feedback_token.txt"


def _load_token():
    for d in (getattr(sys, "_MEIPASS", None), BASE):
        if not d:
            continue
        try:
            with open(os.path.join(d, TOKEN_FILE), "r", encoding="utf-8") as fh:
                tok = fh.read().strip()
                if tok:
                    return tok
        except Exception:
            pass
    return ""
# ------------------------------------------------------------------

ISSUE_LABEL = "user-feedback"
TITLE_PREFIX = "[用户反馈]"


def version():
    try:
        path = os.path.join(BASE, "version.txt")
        if not os.path.isfile(path):
            path = os.path.join(getattr(sys, "_MEIPASS", BASE), "version.txt")
        with open(path, "r", encoding="utf-8") as fh:
            return fh.read().strip()
    except Exception:
        return "未知"


def submit(kind, content, contact="", context=None):
    """提交反馈，返回 (成功与否, 提示文本)。

    context: dict，由界面传入的当前上下文，至少含：
        code      当前工程编号
        action    当前/最近进行的操作（如「开始检查」「自动修正」「检查更新」）
        detail    操作细节（如检查结果统计、文档路径）
        status    界面底部状态栏文字
        time      操作发生时间
    """
    content = (content or "").strip()
    if not content:
        return False, "反馈内容不能为空"

    title = "%s %s - %s" % (
        TITLE_PREFIX,
        (kind or "其他").strip(),
        content.replace("\n", " ")[:30],
    )
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    env = "版本 v%s；Windows %s；提交时间 %s" % (
        version(), os.environ.get("OS_NAME", "Windows"), now)
    ctx = context or {}
    context_text = (
        "工程编号：%s\n"
        "当前操作：%s\n"
        "操作细节：%s\n"
        "界面状态：%s\n"
        "操作时间：%s"
        % (ctx.get("code") or "（未输入）",
           ctx.get("action") or "（无）",
           ctx.get("detail") or "（无）",
           ctx.get("status") or "（无）",
           ctx.get("time") or "（无）")
    )
    body = (
        "### 反馈类型\n%s\n\n"
        "### 反馈内容\n%s\n\n"
        "### 当前操作与工程编号\n%s\n\n"
        "### 联系方式（可选）\n%s\n\n"
        "### 环境信息\n%s\n\n"
        "---\n（本反馈由程序内「反馈」按钮自动提交）\n"
        % (kind or "其他", content, context_text,
           (contact or "").strip() or "（未填写）", env)
    )

    payload = {"title": title, "body": body, "labels": [ISSUE_LABEL]}
    token = _load_token()
    if not token:
        return False, "反馈功能尚未配置提交凭证，请联系程序维护者"

    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(API_ISSUES, data=data, method="POST", headers={
        "Authorization": "Bearer %s" % token,
        "Accept": "application/vnd.github+json",
        "Content-Type": "application/json; charset=utf-8",
        "User-Agent": "gas-doc-checker-feedback/%s" % version(),
        "X-GitHub-Api-Version": "2022-11-28",
    })
    try:
        with urllib.request.urlopen(req, timeout=20) as resp:
            result = json.loads(resp.read().decode("utf-8", "replace"))
        url = result.get("html_url") or ISSUES_PAGE
        return True, "反馈已提交，谢谢！可在 GitHub 查看：%s" % url
    except urllib.error.HTTPError as exc:
        detail = ""
        try:
            err = json.loads(exc.read().decode("utf-8", "replace"))
            detail = err.get("message") or ""
        except Exception:
            pass
        return False, "提交失败（HTTP %s）：%s" % (exc.code, detail or "请稍后重试")
    except Exception as exc:
        return False, "提交失败：%s" % exc


def open_issues_page():
    try:
        os.startfile(ISSUES_PAGE)          # noqa: S606
    except Exception:
        pass
