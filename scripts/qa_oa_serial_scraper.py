#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
质量协议 OA 流水号抓取脚本
=========================
参考 oa-scraper 的 oa_scraper.py 框架，针对质量协议审批流程抓取流水号。
通过 Playwright + Chrome Persistent Context 登录 OA，
导航到质量协议相关审批表单，提取流水号字段。

用法:
  python qa_oa_serial_scraper.py                    # 正常执行
  python qa_oa_serial_scraper.py --dump-fields      # 仅转储表单字段ID（调试用）
"""

import json
import logging
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path

sys.path.insert(0, r"../oa-scraper")
import yaml
from playwright.sync_api import sync_playwright

# =============================================================================
# 日志配置
# =============================================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(name)s] %(levelname)s %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger("qa-oa-scraper")

# =============================================================================
# 配置
# =============================================================================
OA_WEAVER_DIR = r"../oa-scraper"
QA_PROJ_DIR = r".\agreement-automation"
OUTPUT_DIR = r".\scripts"

# OA 账号配置
ACCOUNTS = [
    {
        "id": "account_A",
        "name": os.environ.get("OA_ACCOUNT_NAME", "OA账号"),
        "user_data_dir": os.path.join(OA_WEAVER_DIR, "data", "chrome_profiles", "zhangsan"),
        "home_url": os.environ.get("OA_HOME_URL", ""),
        "username": os.environ.get("OA_ACCOUNT_USERNAME", ""),
        "password": os.environ.get("OA_ACCOUNT_PASSWORD", ""),
    },
]

# 质量协议相关的 OA 菜单路径
QA_MENU_TEXTS = [
    "质量保证协议",
    "质量协议",
    "合同管理",
    "采购合同",
]

# 质量协议表单中可能包含流水号的 CSS 选择器（需先通过 dump-fields 模式确认）
SERIAL_SELECTORS = [
    "#field104800span",
    "#field104900span",
    "#field105000span",
    "span[id*='field'][id*='span']",
]


# =============================================================================
# OA 连接管理器
# =============================================================================
class OABrowserManager:
    def __init__(self):
        self._playwright = None
        self._context = None
        self._current_account = None

    def start_for(self, account: dict):
        """为指定账号启动 Persistent Context"""
        user_dir = account["user_data_dir"]
        logger.info(f"  Starting Chrome Persistent Context: {account['name']}")
        logger.info(f"     Profile: {user_dir}")

        self._clean_lock_files(user_dir)

        self._playwright = sync_playwright().start()
        self._context = self._playwright.chromium.launch_persistent_context(
            user_data_dir=user_dir,
            headless=True,
            channel="chrome",
            args=[
                "--start-maximized",
                "--disable-blink-features=AutomationControlled",
            ],
            no_viewport=True,
        )
        self._current_account = account
        page = self._context.pages[0] if self._context.pages else self._context.new_page()
        return self._context, page

    def stop(self):
        if self._context:
            try:
                self._context.close()
            except Exception:
                pass
        if self._playwright:
            try:
                self._playwright.stop()
            except Exception:
                pass

    @staticmethod
    def _clean_lock_files(user_dir: str):
        import glob
        patterns = ["lockfile", "SingletonLock", "SingletonCookie", "SingletonSocket"]
        for pat in patterns:
            for fp in glob.glob(os.path.join(user_dir, "**", pat), recursive=True):
                try:
                    os.remove(fp)
                except Exception:
                    pass

    @staticmethod
    def check_login(page, selector: str = "#e8rightContentDiv", timeout_ms: int = 10000) -> bool:
        if not selector:
            return True
        try:
            el = page.wait_for_selector(selector, timeout=timeout_ms)
            return el is not None and el.is_visible()
        except Exception:
            return False

    @staticmethod
    def auto_login(page, username: str, password: str, timeout_ms: int = 15000) -> bool:
        current_url = page.url
        if "Login.jsp" not in current_url and "/login/" not in current_url:
            logger.info("  Current page is not login page, no login needed")
            return True

        logger.info("  Detected login page, auto-filling...")
        try:
            loginid_input = page.locator("#loginid")
            if loginid_input.count() == 0:
                logger.warning("  Cannot find #loginid")
                return False
            loginid_input.fill(username)

            pwd_input = page.locator("#userpassword")
            if pwd_input.count() == 0:
                logger.warning("  Cannot find #userpassword")
                return False
            pwd_input.fill(password)

            login_btn = page.locator("#login")
            if login_btn.count() > 0:
                login_btn.click()
            else:
                alt_btn = page.get_by_text("Login", exact=True).first
                if alt_btn.count() > 0:
                    alt_btn.click()
                else:
                    logger.warning("  Cannot find login button")
                    return False

            try:
                page.wait_for_load_state("networkidle", timeout=timeout_ms)
            except Exception:
                pass
            time.sleep(3)

            if "Login.jsp" in page.url or "/login/" in page.url:
                logger.warning("  Still on login page after login")
                return False
            logger.info(f"  Auto-login success: {page.url[:80]}")
            return True
        except Exception as e:
            logger.warning(f"  Auto-login exception: {e}")
            return False


# =============================================================================
# OA 表单字段嗅探器
# =============================================================================
class OAFieldSniffer:
    @staticmethod
    def dump_all_page_fields(main_page):
        """从主页面找到所有 frame 并 dump 字段"""
        all_fields = []
        for pi, page in enumerate([main_page] + [f for f in main_page.frames]):
            try:
                elements = page.locator("[id*='field']")
                count = elements.count()
                if count > 0:
                    logger.info(f"  Page/Frame {pi}: {count} field elements")
                    for ei in range(min(count, 100)):
                        try:
                            el = elements.nth(ei)
                            el_id = el.get_attribute("id")
                            tag = el.evaluate("el => el.tagName")
                            text = el.inner_text()[:60] if el.inner_text() else ""
                            val = el.get_attribute("value") or ""
                            if el_id and ("field" in el_id.lower()):
                                all_fields.append({
                                    "page_idx": pi,
                                    "tag": tag,
                                    "id": el_id,
                                    "text": text.strip(),
                                    "value": val.strip(),
                                })
                        except Exception:
                            pass
            except Exception:
                pass
        return all_fields


# =============================================================================
# 质量协议 OA 流水号抓取器
# =============================================================================
class QAOASerialScraper:
    def __init__(self):
        self.browser_mgr = OABrowserManager()
        self.serials = []
        self.all_fields = []

    def run(self, dump_fields_only=False):
        logger.info("=" * 60)
        logger.info("  Quality Agreement OA Serial Scraper")
        logger.info("=" * 60)

        for account in ACCOUNTS:
            self._run_single_account(account, dump_fields_only)

        return self.serials

    def _run_single_account(self, account: dict, dump_fields_only=False):
        logger.info(f"\nAccount: {account['name']} ({account['id']})")
        if not account.get("password"):
            logger.error("缺少 OA 账号密码，请设置环境变量 OA_ACCOUNT_PASSWORD 后重试")
            return
        if not account.get("home_url"):
            logger.error("缺少 OA 地址，请设置环境变量 OA_HOME_URL 后重试")
            return
        context, page = None, None
        try:
            context, page = self.browser_mgr.start_for(account)

            # 1. Navigate to OA
            logger.info("Navigating to OA...")
            page.goto(account["home_url"], timeout=30000)
            time.sleep(3)

            # 2. Check login
            if not self.browser_mgr.check_login(page):
                logger.info("Session may be expired, trying auto-login...")
                if not self.browser_mgr.auto_login(page, account["username"], account["password"]):
                    logger.error("Login failed")
                    return

            logger.info("OA session valid")

            # 3. Navigate to workflow menu
            self._navigate_to_workflow(page)

            # 4. Dump fields if needed
            if dump_fields_only:
                logger.info("Dumping OA form fields...")
                self.all_fields = OAFieldSniffer.dump_all_page_fields(page)
                self._print_fields()
                return

            # 5. Extract serials from list
            logger.info("Extracting serials from approval list...")
            self._extract_serials(page)

        except Exception as e:
            logger.error(f"Account processing error: {e}")
            import traceback
            traceback.print_exc()
        finally:
            if self.browser_mgr:
                self.browser_mgr.stop()

    def _navigate_to_workflow(self, page):
        """Navigate to workflow menu"""
        # Click top menu
        for menu_text in ["Process", "流程"]:
            try:
                m = page.get_by_text(menu_text, exact=True).first
                if m.count() > 0:
                    m.click()
                    time.sleep(2)
                    break
            except:
                pass

        # Click left menu items
        for menu_text in ["Pending", "待办事宜", "Completed", "已办事宜"]:
            try:
                item = page.get_by_text(menu_text, exact=True).first
                if item.count() > 0:
                    item.click()
                    time.sleep(2)
                    break
            except:
                pass

        # Search for QA-related items in all frames
        for qa_text in QA_MENU_TEXTS:
            for f in page.frames:
                try:
                    item = f.get_by_text(qa_text, exact=False).first
                    if item.count() > 0:
                        item.click()
                        logger.info(f"Clicked '{qa_text}'")
                        time.sleep(3)
                        return True
                except:
                    pass

        logger.info("QA menu item not found. Will try to search list directly.")
        return False

    def _extract_serials(self, page):
        """Extract serial numbers from list"""
        for fi, f in enumerate(page.frames):
            try:
                rows = f.locator("tr")
                rc = rows.count()
                logger.info(f"  Frame {fi}: {rc} rows")
                for ri in range(rc):
                    try:
                        text = rows.nth(ri).inner_text()
                        # Look for serial patterns
                        for pat in [r'(HT[-_]?MB[-_]?QM[\d-]+)', r'(QA[\d-]+)', r'(YCSW\d+)']:
                            m = re.search(pat, text)
                            if m:
                                serial = m.group(1)
                                self.serials.append({
                                    "serial": serial,
                                    "frame": fi,
                                    "row": ri,
                                    "text": text[:200],
                                })
                                logger.info(f"  Found serial: {serial}")
                                break
                    except:
                        pass
            except:
                pass

        if not self.serials:
            logger.info("No QA workflow records found in OA")

    def _print_fields(self):
        print(f"\n{'='*60}")
        print(f"  OA Form Fields: {len(self.all_fields)} total")
        print(f"{'='*60}")
        self.all_fields.sort(key=lambda x: (x.get("page_idx", 0), x.get("id", "")))
        for f in self.all_fields:
            print(f"  [{f['page_idx']}] <{f['tag']}> id={f['id']}")
            if f["text"]:
                print(f"      text: {f['text'][:60]}")
            if f["value"]:
                print(f"      value: {f['value'][:60]}")


# =============================================================================
# Main
# =============================================================================
if __name__ == "__main__":
    scraper = QAOASerialScraper()
    dump_only = "--dump-fields" in sys.argv
    serials = scraper.run(dump_fields_only=dump_only)

    if serials:
        print(f"\nFound {len(serials)} serials:")
        for s in serials:
            print(f"  - {s['serial']}")
    else:
        print(f"\nNo QA OA serials found.")
        print(f"Note: Quality agreements need to be submitted to OA workflow first.")

    print(f"\nDone")
