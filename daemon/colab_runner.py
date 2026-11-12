import asyncio
from typing import Dict, Any
from playwright.async_api import Page

from browser import browser_manager
from notify import notify_windows


async def _is_login_required(page: Page) -> bool:
    """Checks if Colab redirected to Google Sign-In or shows a blocking login prompt."""
    url = page.url.lower()
    if "accounts.google.com" in url:
        return True
    try:
        sign_in_btn = await page.query_selector("a:has-text('Sign in'), paper-button:has-text('Sign in')")
        if sign_in_btn and await sign_in_btn.is_visible():
            # Only treat as blocking if no cells are editable/present
            cells = await page.query_selector_all("div.cell")
            if not cells:
                return True
    except Exception:
        pass
    return False


async def _dismiss_trust_warning_if_present(page: Page) -> None:
    """Handles Colab's 'This notebook was not authored by Google' dialog."""
    try:
        for selector in ("mwc-button:has-text('Run anyway')", "button:has-text('Run anyway')", "#ok"):
            btn = await page.query_selector(selector)
            if btn and await btn.is_visible():
                await btn.click()
                await asyncio.sleep(0.8)
                break
    except Exception:
        pass


async def _trigger_run_all(page: Page) -> bool:
    """Attempts to trigger 'Run all' via Colab UI menu or Ctrl+F9 shortcut."""
    # 1. Try explicit Runtime -> Run all menu
    try:
        runtime_menu = await page.query_selector("#runtime-menu-button")
        if runtime_menu and await runtime_menu.is_visible():
            await runtime_menu.click()
            await asyncio.sleep(0.5)
            run_all_item = await page.query_selector(
                "div[role='menuitem']:has-text('Run all'), command-menu-item:has-text('Run all')"
            )
            if run_all_item:
                await run_all_item.click()
                await asyncio.sleep(0.6)
                await _dismiss_trust_warning_if_present(page)
                return True
    except Exception:
        pass

    # 2. Try mock/test runner button if testing locally
    try:
        test_btn = await page.query_selector("#run-all-btn, button:has-text('Run all')")
        if test_btn and await test_btn.is_visible():
            await test_btn.click()
            return True
    except Exception:
        pass

    # 3. Fallback: Colab global keyboard shortcut (Ctrl + F9)
    try:
        await page.keyboard.press("Control+F9")
        await asyncio.sleep(0.6)
        await _dismiss_trust_warning_if_present(page)
        return True
    except Exception:
        return False


async def _check_colab_state(page: Page) -> Dict[str, Any]:
    """Inspects Colab DOM to determine if cells are still running or hit an error."""
    return await page.evaluate("""() => {
        // Real Colab running indicators & mock test indicators
        const runningCells = document.querySelectorAll(
            'div.cell.running, colab-run-button[running], [data-running="true"]'
        ).length;
        const queuedCells = document.querySelectorAll(
            'div.cell.pending, colab-run-button[queued], [data-queued="true"]'
        ).length;
        const errorOutputs = document.querySelectorAll(
            'div.output_error, .colab-error-output, [data-error="true"]'
        ).length;
        const completedBadge = document.querySelector('[data-colab-done="true"]') !== null;
        return {
            running: runningCells,
            queued: queuedCells,
            errors: errorOutputs,
            completedBadge: completedBadge
        };
    }""")


async def run_colab_notebook(
    notebook_url: str = "",
    poll_interval: float = 5.0,
    max_wait_sec: float = 7200.0
) -> Dict[str, Any]:
    """
    Opens (or attaches to) a Google Colab notebook in a silent unthrottled background tab,
    triggers 'Run all', and monitors execution until completion or cancellation.
    """
    notebook_url = notebook_url.strip()

    try:
        if notebook_url and notebook_url.startswith(("http://", "https://", "file:///")):
            page, _ = await browser_manager.create_background_tab(notebook_url)
            # Allow Colab SPA scripts to hydrate
            await asyncio.sleep(2.0)
        else:
            page, _ = await browser_manager.get_target_page(prefer_active=True)

        current_url = page.url.lower()
        if "chrome://" in current_url:
            msg = "No notebook URL provided and active tab is a Chrome internal page."
            notify_windows("Chitran — Colab Error", msg)
            return {"status": "error", "message": msg}

        if await _is_login_required(page):
            await page.bring_to_front()
            msg = "Google Sign-In required on Colab. Complete login in Chrome and retry."
            notify_windows("Chitran — Auth Required", msg)
            return {"status": "auth_required", "message": msg}

        triggered = await _trigger_run_all(page)
        if not triggered:
            msg = "Could not trigger 'Run all' on the target tab."
            notify_windows("Chitran — Colab Error", msg)
            return {"status": "error", "message": msg}

        page_title = await page.title()
        notify_windows(
            "Chitran — Colab Started",
            f"Running '{page_title}' in background. Safe to press Win + L."
        )

        # Wait briefly for first cell spinner to engage
        await asyncio.sleep(2.0)

        elapsed = 0.0
        idle_streak = 0
        seen_active = False

        while elapsed < max_wait_sec:
            state = await _check_colab_state(page)
            active_count = state["running"] + state["queued"]

            if active_count > 0:
                seen_active = True
                idle_streak = 0
            else:
                idle_streak += 1

            if state["errors"] > 0:
                msg = f"Notebook '{page_title}' stopped with {state['errors']} cell error(s)."
                notify_windows("Chitran — Colab Error", msg)
                return {
                    "status": "error",
                    "page_title": page_title,
                    "errors": state["errors"],
                    "message": msg
                }

            # Consider finished when cells ran & returned to idle for 2 polls, or explicit done badge
            if state["completedBadge"] or (seen_active and idle_streak >= 2) or (not seen_active and idle_streak >= 3):
                msg = f"Colab execution finished for '{page_title}'."
                notify_windows("Chitran — Colab Complete", msg)
                return {
                    "status": "success",
                    "page_title": page_title,
                    "elapsed_sec": round(elapsed, 1),
                    "message": msg
                }

            await asyncio.sleep(poll_interval)
            elapsed += poll_interval

        msg = f"Colab monitor reached max timeout ({int(max_wait_sec)}s) for '{page_title}'."
        notify_windows("Chitran — Colab Timeout", msg)
        return {"status": "timeout", "page_title": page_title, "message": msg}

    except asyncio.CancelledError:
        notify_windows("Chitran — Colab Aborted", "Background Colab monitor cancelled via cn kill.")
        raise
    except Exception as e:
        msg = f"Colab runner failed: {e}"
        notify_windows("Chitran — Colab Error", msg)
        return {"status": "error", "message": msg}

