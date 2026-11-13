# -*- coding: utf-8 -*-
"""设备登记 —— 读取本机设备名并上报信息平台，首次使用自动授权。

设计要点（V1.4.5）：
  · 设备名：计算机名（COMPUTERNAME）
  · 设备ID：计算机名 + 网卡 MAC 后 6 位；计算机改名后仍能对应同一台机器
  · 程序启动时自动上报，平台首次收到即自动登记并放行（无需人工建授权）
  · 管理员在信息平台「设备管理」里点停用后，本机下次启动会被拦下
  · 上报失败（断网 / 云端故障 / 公司网络限制）一律放行（fail-open），
    只有平台明确回复「已停用 / 已过期」才拦——绝不让网络问题把人挡在门外
"""
import getpass
import json
import os
import platform
import socket
import urllib.parse
import urllib.request
import uuid

PLATFORM_URL = "https://info-platform.app.workbuddy.host"
SHARECODE = "KC4MIHDOVT"
APP_NAME = "线上资料检查"


def device_name():
    """本机设备名（计算机名）。"""
    n = (os.environ.get("COMPUTERNAME") or "").strip()
    if not n:
        try:
            n = (socket.gethostname() or "").strip()
        except Exception:
            n = ""
    return n or "未知设备"


def _mac_tail():
    try:
        return ("%012x" % uuid.getnode())[-6:]
    except Exception:
        return "000000"


def device_id():
    """设备唯一标识：计算机名 + MAC 后 6 位。"""
    return "%s-%s" % (device_name(), _mac_tail())


def user_name():
    try:
        return (os.environ.get("USERNAME") or getpass.getuser() or "").strip()
    except Exception:
        return ""


def os_name():
    try:
        return ("%s %s" % (platform.system(), platform.release())).strip()
    except Exception:
        return ""


def register(version, timeout=8):
    """上报本机，返回 dict：

    ok       平台是否放行
    blocked  是否被明确拒绝（停用/过期）——只有这种才应该拦用户
    first    是否首次登记
    msg      平台返回的提示
    net_error 连不上时的异常摘要（网络问题，应放行）
    """
    query = {
        "app": APP_NAME,
        "device": device_id(),
        "name": device_name(),
        "user": user_name(),
        "os": os_name(),
        "ver": str(version),
        "sharecode": SHARECODE,
    }
    url = "%s/api/check?%s" % (PLATFORM_URL, urllib.parse.urlencode(query))
    try:
        req = urllib.request.Request(
            url, headers={"User-Agent": "gas-doc-checker/%s" % version})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            info = json.loads(resp.read().decode("utf-8", "replace"))
    except Exception as exc:                       # 断网/超时/云端异常 → 放行
        return {"ok": False, "blocked": False, "first": False,
                "msg": "连不上信息平台", "net_error": str(exc)[:200],
                "name": device_name(), "id": device_id()}

    msg = info.get("msg") or ""
    info["blocked"] = (not info.get("ok")) and ("停用" in msg or "过期" in msg)
    info["name"] = info.get("device") or device_name()
    info["id"] = info.get("device_id") or device_id()
    return info
