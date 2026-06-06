"""
Page object for the Falcon staging application launcher.

Covers navigation to the Falcon App Launcher, launching the Order Exporter
and Order Generator tools, configuring order ranges, and handling file downloads.
"""

import logging
from pathlib import Path

from playwright.sync_api import Download, Page

from config.settings import (
    DOWNLOADS_DIR,
    FALCON_APP_LAUNCHER_URL,
    FALCON_EXPORT_CHANNEL,
    FALCON_ORDER_RANGE,
    FALCON_PRODUCTION_ORDER,
    VIEWPORT,
)

log = logging.getLogger(__name__)


class FalconPage:
    """
    Encapsulates all browser interactions with the Falcon staging environment.

    Supported actions
    -----------------
    * ``navigate_to_launcher()``              — open the Falcon App Launcher
    * ``download_order_exporter_report()``    — configure and download the
                                                Order Exporter workbook
    * ``download_order_generator_report()``  — configure and download the
                                                Order Generator workbook
    """

    def __init__(self, page: Page) -> None:
        self._page = page

    # ── Public interface ───────────────────────────────────────────────────────

    def navigate_to_launcher(self) -> None:
        """Navigate to the Falcon App Launcher page."""
        log.info("Navigating to Falcon App Launcher: %s", FALCON_APP_LAUNCHER_URL)
        self._page.goto(FALCON_APP_LAUNCHER_URL)
        self._page.wait_for_load_state("networkidle")

    def download_order_exporter_report(self) -> Path:
        """
        Launch the Order Exporter, configure it for the project order range and
        channel, trigger a download, and return the local path of the saved file.

        Returns:
            ``Path`` to the downloaded Excel workbook.
        """
        log.info("Launching Order Exporter …")
        self._page.locator("a", has_text="Order Exporter").click()
        self._page.wait_for_load_state("networkidle")

        log.info("Configuring Order Exporter: range=%s, channel=%s",
                 FALCON_ORDER_RANGE, FALCON_EXPORT_CHANNEL)
        self._page.locator("input[name='order_number_range']").fill(FALCON_ORDER_RANGE)

        channel_select = self._page.locator("select[name='channel']")
        if channel_select.count() > 0:
            channel_select.select_option(FALCON_EXPORT_CHANNEL)

        with self._page.expect_download() as download_info:
            self._page.locator("button[type='submit'], button:has-text('Export')").first.click()

        download: Download = download_info.value
        dest = DOWNLOADS_DIR / download.suggested_filename
        download.save_as(str(dest))
        log.info("Order Exporter workbook saved → %s", dest)
        return dest

    def download_order_generator_report(self) -> Path:
        """
        Launch the Order Generator (Production Order tool), configure it for
        the project order range, trigger a download, and return the local path.

        Returns:
            ``Path`` to the downloaded Excel workbook.
        """
        log.info("Navigating back to App Launcher …")
        self._page.goto(FALCON_APP_LAUNCHER_URL)
        self._page.wait_for_load_state("networkidle")

        log.info("Launching Order Generator …")
        self._page.locator("a", has_text="Production Order").click()
        self._page.wait_for_load_state("networkidle")

        log.info("Configuring Order Generator: range=%s", FALCON_ORDER_RANGE)
        self._page.locator("input[name='order_number_range']").fill(FALCON_ORDER_RANGE)
        self._page.locator("input[name='production_order_name']").fill(FALCON_PRODUCTION_ORDER)

        with self._page.expect_download() as download_info:
            self._page.locator("button[type='submit'], button:has-text('Generate')").first.click()

        download: Download = download_info.value
        dest = DOWNLOADS_DIR / download.suggested_filename
        download.save_as(str(dest))
        log.info("Order Generator workbook saved → %s", dest)
        return dest
