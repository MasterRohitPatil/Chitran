import asyncio
import os
import subprocess
import sys
import time
from pathlib import Path
from typing import Optional, Tuple, Dict, Any

import httpx
from playwright.async_api import async_playwright, Browser, BrowserContext, Page, CDPSession

PROJECT_ROOT = Path(__file__).resolve().parent.parent
LAUNCH_BAT = PROJECT_ROOT / "scripts" / "launch_chrome_cdp.bat"
CDP_URL = "http://127.0.0.1:9222"


async def is_cdp_alive(timeout_sec: float = 1.0) -> bool:
    """Checks if Chrome DevTools Protocol is reachable on port 9222."""
    try:
        async with httpx.AsyncClient(timeout=timeout_sec) as client:
            resp = await client.get(f"{CDP_URL}/json/version")
            return resp.status_code == 200
    except Exception:
        return False


async def ensure_chrome_cdp() -> bool:
    """Ensures Chrome is running with CDP enabled. Spawns launch_chrome_cdp.bat if needed."""
    if await is_cdp_alive():
        return True

    print("[Browser Bridge] Chrome CDP port 9222 not reachable. Launching Chrome with CDP...")
    try:
        subprocess.Popen(
            [str(LAUNCH_BAT)],
            shell=True,
            creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
        )
    except Exception as e:
        print(f"[Browser Bridge] Failed to execute launch script: {e}")
        return False

    # Poll for up to 6 seconds for CDP to come online
    for _ in range(12):
        await asyncio.sleep(0.5)
        if await is_cdp_alive():
            print("[Browser Bridge] Chrome CDP is now online and listening.")
            return True

    print("[Browser Bridge] Timed out waiting for Chrome CDP.")
    return False


class BrowserManager:
    """Manages CDP connection, focus emulation, and background tabs."""

    def __init__(self):
        self._playwright = None
        self._browser: Optional[Browser] = None

    async def connect(self) -> Tuple[Browser, BrowserContext]:
        """Connects to active Chrome instance via CDP."""
        alive = await ensure_chrome_cdp()
        if not alive:
            raise ConnectionError("Could not connect to Chrome on port 9222.")

        if self._playwright is None:
            self._playwright = await async_playwright().start()

        if self._browser is not None:
            try:
                # Verify existing CDP connection is genuinely responsive
                _ = self._browser.contexts
            except Exception:
                self._browser = None

        if self._browser is None or not self._browser.is_connected():
            try:
                self._browser = await self._playwright.chromium.connect_over_cdp(CDP_URL)
            except Exception:
                # Reset playwright instance once and retry fresh
                try:
                    await self._playwright.stop()
                except Exception:
                    pass
                self._playwright = await async_playwright().start()
                self._browser = await self._playwright.chromium.connect_over_cdp(CDP_URL)

        contexts = self._browser.contexts
        context = contexts[0] if contexts else await self._browser.new_context()
        return self._browser, context

    async def get_target_page(self, prefer_active: bool = True) -> Tuple[Page, CDPSession]:
        """
        Retrieves the target tab and enables virtual focus emulation.
        Virtual focus makes DOM timers & inputs think the tab is focused without moving physical mouse.
        """
        _, context = await self.connect()
        pages = context.pages

        if prefer_active and pages:
            # Pick the most recent open page
            page = pages[-1]
        else:
            page = await context.new_page()

        # Apply Chrome DevTools Protocol Focus Emulation
        cdp_session = await context.new_cdp_session(page)
        await cdp_session.send("Emulation.setFocusEmulationEnabled", {"enabled": True})

        return page, cdp_session

    async def create_background_tab(self, url: str) -> Tuple[Page, CDPSession]:
        """
        Opens a URL in a silent background tab using CDP Target.createTarget(background=True)
        so it never steals mouse/keyboard focus or interrupts active video/work.
        """
        browser, context = await self.connect()
        existing_pages = set(context.pages)

        try:
            browser_cdp = await browser.new_browser_cdp_session()
            await browser_cdp.send("Target.createTarget", {"url": url, "background": True})
            await browser_cdp.detach()
        except Exception:
            # Fallback to standard Playwright new_page if browser CDP Target.createTarget fails
            page = await context.new_page()
            await page.goto(url, wait_until="domcontentloaded")
            cdp_session = await context.new_cdp_session(page)
            await cdp_session.send("Emulation.setFocusEmulationEnabled", {"enabled": True})
            return page, cdp_session

        # Wait briefly for Playwright context to register the new target
        target_page = None
        for _ in range(20):
            await asyncio.sleep(0.2)
            new_pages = [p for p in context.pages if p not in existing_pages]
            if new_pages:
                target_page = new_pages[-1]
                break

        if target_page is None:
            target_page = context.pages[-1] if context.pages else await context.new_page()

        await target_page.wait_for_load_state("domcontentloaded")
        cdp_session = await context.new_cdp_session(target_page)
        await cdp_session.send("Emulation.setFocusEmulationEnabled", {"enabled": True})
        return target_page, cdp_session

    async def get_active_page_info(self) -> Dict[str, Any]:
        """Returns details about currently active tab in Sidekick Chrome."""
        try:
            page, _ = await self.get_target_page(prefer_active=True)
            return {
                "status": "connected",
                "title": await page.title(),
                "url": page.url,
            }
        except Exception as e:
            return {
                "status": "error",
                "message": str(e)
            }

    async def close(self):
        """Disconnects Playwright from CDP without closing the actual browser."""
        if self._browser and self._browser.is_connected():
            await self._browser.close()
            self._browser = None
        if self._playwright:
            await self._playwright.stop()
            self._playwright = None


# Singleton instance
browser_manager = BrowserManager()


async def _self_test():
    print("[Browser Bridge] Running self-test...")
    alive = await is_cdp_alive()
    print(f"  CDP alive: {alive}")
    if not alive:
        print("  Triggering ensure_chrome_cdp()...")
        alive = await ensure_chrome_cdp()
        print(f"  CDP alive after launch: {alive}")

    bm = BrowserManager()
    try:
        page, cdp = await bm.get_target_page(prefer_active=True)
        title = await page.title()
        url = page.url
        print(f"  Connected to page!")
        print(f"  Title: {title}")
        print(f"  URL: {url}")
        print(f"  Virtual focus emulation enabled successfully: True")
    except Exception as e:
        print(f"  Self-test failed: {e}")
    finally:
        await bm.close()


if __name__ == "__main__":
    asyncio.run(_self_test())

