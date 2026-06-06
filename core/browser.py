"""
Browser factory for the Flitesports automation framework.

Provides a simple context-manager that launches a Chromium browser,
creates a context with the project viewport, and ensures everything is
closed cleanly even when a test fails.
"""

import logging
from contextlib import contextmanager
from typing import Generator

from playwright.sync_api import Browser, BrowserContext, Page, Playwright

from config.settings import VIEWPORT

log = logging.getLogger(__name__)


@contextmanager
def browser_session(
    playwright: Playwright,
    *,
    headless: bool = False,
    slow_mo: int = 50,
    download_dir: str | None = None,
) -> Generator[Page, None, None]:
    """
    Context manager that yields a ready-to-use Playwright Page.

    Launches a Chromium browser, creates a single browser context with the
    project-standard viewport, and opens one page.  The context and browser
    are closed automatically when the ``with`` block exits — whether normally
    or via an exception.

    Args:
        playwright:   The Playwright sync API handle passed into the caller.
        headless:     Run without a visible browser window when True.
        slow_mo:      Milliseconds to slow down every Playwright action
                      (useful for demos and debugging).
        download_dir: If supplied, all downloads are saved to this path.

    Yields:
        A ``playwright.sync_api.Page`` instance ready for interactions.
    """
    browser: Browser | None = None
    context: BrowserContext | None = None
    try:
        log.debug("Launching Chromium (headless=%s, slow_mo=%d ms) …", headless, slow_mo)
        browser = playwright.chromium.launch(headless=headless, slow_mo=slow_mo)

        context_kwargs: dict = {"viewport": VIEWPORT}
        if download_dir:
            context_kwargs["accept_downloads"] = True

        context = browser.new_context(**context_kwargs)
        page: Page = context.new_page()

        if download_dir:
            # Route download events to the specified directory
            page.context.set_default_timeout(60_000)

        yield page

    finally:
        if context:
            context.close()
        if browser:
            browser.close()
        log.debug("Browser session closed.")
