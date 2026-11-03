# -*- coding: utf-8 -*-
"""紫鸟浏览器本地接口封装（导数自动化，2026-10-27）

为什么走紫鸟而不是直接开浏览器：
  账号登录、短信/图形验证码、IP 环境隔离全在紫鸟里做好了。紫鸟客户端切到
  「web_driver 模式」后会在本地开一个 HTTP 服务，startBrowser 返回
  `debuggingPort`，Playwright 用 connect_over_cdp 直接接管那个**已登录**的
  店铺窗口——我们只管点页面、点下载，不碰登录。

官方资料：
  开通 WebDriver 权限  https://open.ziniao.com/docSupport?docId=99
  接口对接文档          https://open.ziniao.com/docSupport?docId=98
  官方示例             https://github.com/ziniao-open/ziniao_webdriver_demo

用法：
    from ziniao_api import Ziniao
    z = Ziniao.load()                       # 读同目录 config.json
    z.start_client()                        # 以 web_driver 模式拉起客户端
    stores = z.browser_list()               # 店铺列表
    ret = z.start_browser(stores[0]["browserOauth"])
    with z.page(ret["debuggingPort"]) as page:
        page.goto("https://sellercentral.amazon.com/")
    z.stop_browser(ret["browserOauth"])
"""
import json
import os
import subprocess
import sys
import time
import uuid
from contextlib import contextmanager

import requests

_HERE = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(_HERE, "config.json")

DEFAULT_PORT = 16851

# 紫鸟客户端在 Windows 下的默认安装位置（config.json 里没填时按这个顺序找）
CLIENT_CANDIDATES = [
    r"D:\Program Files\ziniao\ziniao.exe",
    r"C:\Program Files\ziniao\ziniao.exe",
    r"C:\Program Files (x86)\ziniao\ziniao.exe",
]


class ZiniaoError(RuntimeError):
    pass


class Ziniao:
    def __init__(self, company, username, password, port=DEFAULT_PORT,
                 client_path=None, log=print):
        self.user_info = {"company": company, "username": username, "password": password}
        self.port = int(port)
        self.client_path = client_path or self._guess_client()
        self.log = log

    # ---------------- 配置 ----------------
    @classmethod
    def load(cls, path=CONFIG_PATH, log=print):
        if not os.path.exists(path):
            raise ZiniaoError(
                "没找到配置文件：%s\n"
                "复制 config.example.json 成 config.json，填 company / username / password。"
                % path)
        with open(path, encoding="utf-8") as f:
            cfg = json.load(f)
        z = cfg.get("ziniao", cfg)
        missing = [k for k in ("company", "username", "password") if not z.get(k)]
        if missing:
            raise ZiniaoError("config.json 里这几个字段还是空的：%s" % ", ".join(missing))
        return cls(z["company"], z["username"], z["password"],
                   port=z.get("port", DEFAULT_PORT),
                   client_path=z.get("client_path"), log=log)

    @staticmethod
    def _guess_client():
        for p in CLIENT_CANDIDATES:
            if os.path.exists(p):
                return p
        return CLIENT_CANDIDATES[0]

    # ---------------- 底层通讯 ----------------
    def _post(self, payload, timeout=120):
        body = {"requestId": str(uuid.uuid4())}
        body.update(payload)
        body.update(self.user_info)
        url = "http://127.0.0.1:%d" % self.port
        try:
            r = requests.post(url, json.dumps(body).encode("utf-8"), timeout=timeout)
        except requests.exceptions.ConnectionError:
            raise ZiniaoError(
                "连不上紫鸟本地接口 127.0.0.1:%d。\n"
                "  1) 紫鸟客户端要开着，并且是 **web_driver 模式**（不是平时那种普通模式）；\n"
                "  2) 账号要开通 WebDriver 权限：https://open.ziniao.com/docSupport?docId=99\n"
                "  3) 端口不是默认 16851 的话，改 config.json 里的 port。"
                % self.port)
        try:
            return json.loads(r.text)
        except ValueError:
            raise ZiniaoError("紫鸟返回的不是 JSON：%r" % r.text[:200])

    def _check(self, r, action):
        code = str(r.get("statusCode", r.get("code", "")))
        if code == "0":
            return r
        if code == "-10003":
            raise ZiniaoError("紫鸟登录失败（-10003）：检查 config.json 的 company/username/password。\n%r"
                              % r)
        raise ZiniaoError("%s 失败，statusCode=%s\n%r" % (action, code, r))

    # ---------------- 客户端进程 ----------------
    def start_client(self, wait=8, kill_existing=False):
        """以 web_driver 模式拉起紫鸟客户端。

        kill_existing=True 会先杀掉已经在跑的 ziniao.exe —— 普通模式的客户端
        不会响应接口调用，必须重启成 web_driver 模式。默认不动，避免打断你正在做的事。
        """
        if kill_existing:
            self.log("[紫鸟] 关闭已有客户端…")
            subprocess.run(["taskkill", "/F", "/IM", "ziniao.exe"],
                           capture_output=True, text=True)
            time.sleep(2)
        if not os.path.exists(self.client_path):
            raise ZiniaoError("紫鸟客户端不存在：%s（改 config.json 的 client_path）"
                              % self.client_path)
        self.log("[紫鸟] 启动 web_driver 模式：%s --port=%d" % (self.client_path, self.port))
        subprocess.Popen([self.client_path,
                          "--run_type=web_driver",
                          "--ipc_type=http",
                          "--port=%d" % self.port])
        time.sleep(wait)
        return self

    def update_core(self, retries=3):
        """更新内核到 webdriver 需要的版本（紫鸟要求的预热步骤）。"""
        for i in range(retries):
            r = self._post({"action": "updateCore"})
            if str(r.get("statusCode")) == "0":
                return r
            if str(r.get("statusCode")) == "-10003":
                self._check(r, "updateCore")
            self.log("[紫鸟] updateCore 第 %d 次未成功：%s" % (i + 1, r.get("statusCode")))
            time.sleep(2)
        return None

    def exit_client(self):
        try:
            return self._post({"action": "exit"}, timeout=15)
        except Exception:
            return None

    # ---------------- 店铺 ----------------
    def browser_list(self):
        """→ [{'browserName','siteId','browserOauth','browserId', ...}, ...]"""
        r = self._post({"action": "getBrowserList"})
        self._check(r, "getBrowserList")
        return r.get("browserList") or []

    def start_browser(self, store, headless=0, extra=None):
        """启动一个店铺环境 → 返回含 debuggingPort / downloadPath / launcherPage 的 dict。

        store 可以是 browserOauth（字符串）或 browserId（纯数字）。
        """
        payload = {
            "action": "startBrowser",
            "isWaitPluginUpdate": 0,
            "isHeadless": headless,
            "isWebDriverReadOnlyMode": 0,
            "cookieTypeLoad": 0,
            "cookieTypeSave": 0,
            "runMode": "1",
            "isLoadUserPlugin": False,
            "pluginIdType": 1,
            "privacyMode": 0,
            # 1 = 下载文件不弹「选择保存路径」，直接落到 downloadPath
            "notPromptForDownload": 1,
        }
        s = str(store)
        if s.isdigit():
            payload["browserId"] = s
        else:
            payload["browserOauth"] = s
        if extra:
            payload.update(extra)
        r = self._post(payload)
        self._check(r, "startBrowser")
        return r

    def stop_browser(self, store):
        payload = {"action": "stopBrowser"}
        s = str(store)
        if s.isdigit():
            payload["browserId"] = s
        else:
            payload["browserOauth"] = s
        try:
            return self._check(self._post(payload, timeout=30), "stopBrowser")
        except ZiniaoError:
            return None

    # ---------------- Playwright 接管 ----------------
    @contextmanager
    def page(self, debugging_port, timeout_ms=60000):
        """连上店铺窗口，交出一个可用的 page。

        紫鸟店铺窗口是 Chromium，CDP 端口就在 127.0.0.1:debuggingPort。
        不需要 chromedriver，Playwright 直接连。
        """
        from playwright.sync_api import sync_playwright
        pw = sync_playwright().start()
        browser = None
        try:
            browser = pw.chromium.connect_over_cdp("http://127.0.0.1:%s" % debugging_port,
                                                   timeout=timeout_ms)
            ctx = browser.contexts[0] if browser.contexts else browser.new_context()
            pg = ctx.pages[0] if ctx.pages else ctx.new_page()
            pg.set_default_timeout(timeout_ms)
            yield pg
        finally:
            # 只断开连接，不 close —— 店铺窗口要留着给下一个任务用，
            # 真正的关闭交给 stop_browser。
            try:
                if browser is not None:
                    browser.close()
            except Exception:
                pass
            pw.stop()


def _main():
    """命令行自检：python ziniao_api.py [--kill] [--open <browserOauth>]"""
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--kill", action="store_true", help="先杀掉已在跑的紫鸟客户端（重启成 web_driver 模式）")
    ap.add_argument("--open", default=None, help="启动指定店铺并打印当前页面标题/URL")
    a = ap.parse_args()

    try:
        z = Ziniao.load()
        z.start_client(kill_existing=a.kill)
        z.update_core()
        stores = z.browser_list()
    except ZiniaoError as e:
        print("\n[失败] %s" % e)
        return 2

    print("\n共 %d 个店铺：" % len(stores))
    for s in stores:
        print("  %-28s siteId=%-4s oauth=%s"
              % (s.get("browserName"), s.get("siteId"), s.get("browserOauth")))
    if a.open:
        try:
            ret = z.start_browser(a.open)
            port = ret.get("debuggingPort")
            print("\n店铺已启动  debuggingPort=%s  downloadPath=%s"
                  % (port, ret.get("downloadPath")))
            with z.page(port) as pg:
                print("  当前页：", pg.title(), "|", pg.url)
        except ZiniaoError as e:
            print("\n[失败] %s" % e)
            return 2
        finally:
            z.stop_browser(a.open)
    return 0


if __name__ == "__main__":
    sys.exit(_main())
