"""
Page object for the Flitesports CRM login page and sign-out flow.
"""

import logging

from playwright.sync_api import Page

from config.settings import CRM_BASE_URL

log = logging.getLogger(__name__)


class LoginPage:
    """
    Encapsulates all interactions with the CRM authentication layer.

    Supported actions
    -----------------
    * ``navigate()``             — go to /login
    * ``login(email, password)`` — fill credentials and submit the form
    * ``sign_out()``             — open the user menu and click Sign out
    """

    def __init__(self, page: Page) -> None:
        self._page = page

    # ── Selectors ──────────────────────────────────────────────────────────────
    _SEL_EMAIL       = "form > div:nth-of-type(1) input"
    _SEL_PASSWORD    = "form > div:nth-of-type(2) input"
    _SEL_SUBMIT_SPAN = "span.lp-submit__label"
    _SEL_SUBMIT_BTN  = "form > button"
    _SEL_USER_ROLE   = "span.user-role"
    _SEL_HEADER_ICON = "header i"
    _SEL_MENU_ITEMS  = "ul button"

    # ── Public interface ───────────────────────────────────────────────────────

    def navigate(self) -> None:
        """Navigate to the CRM login page and wait for it to be ready."""
        log.info("Navigating to login page: %s/login", CRM_BASE_URL)
        self._page.goto(f"{CRM_BASE_URL}/login")
        self._page.wait_for_load_state("networkidle")

    def login(self, email: str, password: str) -> None:
        """
        Fill the login form with *email* and *password* and submit it.

        The method handles two submit-button variants observed across CRM
        builds: a ``<span>`` with class ``lp-submit__label`` (standard build)
        and a plain ``<button>`` element (super-admin build).
        """
        log.info("Entering credentials for %s …", email)
        self._page.locator(self._SEL_EMAIL).fill(email)
        self._page.locator(self._SEL_PASSWORD).fill(password)

        # Prefer the span-based submit; fall back to the button variant
        span_submit = self._page.locator(self._SEL_SUBMIT_SPAN)
        if span_submit.count() > 0 and span_submit.is_visible():
            span_submit.click()
        else:
            self._page.locator(self._SEL_SUBMIT_BTN).click()

        self._page.wait_for_load_state("networkidle")
        log.info("Login successful.")

    def sign_out(self) -> None:
        """
        Open the user-account menu and click the Sign out option.

        Handles two menu-trigger variants:
          * ``span.user-role``  (standard admin / sales-rep sessions)
          * ``header i``        (super-admin sessions)

        Before attempting to open the menu, any Toastify notification banners
        are given up to 8 seconds to clear.  These overlays intercept pointer
        events and will cause the click to time out if not dismissed first.
        """
        log.info("Signing out …")

        # Wait for any post-submission toast notifications to fully disappear
        # before interacting with the header, as they sit in a higher z-index
        # layer and block pointer events on elements beneath them.
        try:
            self._page.wait_for_selector(
                ".Toastify__toast-container", state="hidden", timeout=8_000
            )
        except Exception:
            pass  # No lingering toast — proceed immediately

        # Also ensure no modal overlay is still mounted on top of the page.
        try:
            self._page.wait_for_selector(
                "[role='dialog']", state="hidden", timeout=5_000
            )
        except Exception:
            pass

        user_role = self._page.locator(self._SEL_USER_ROLE)
        if user_role.count() > 0 and user_role.is_visible():
            user_role.click()
        else:
            self._page.locator(self._SEL_HEADER_ICON).click()

        self._page.wait_for_selector(self._SEL_MENU_ITEMS, state="visible", timeout=8_000)
        self._page.get_by_text("Sign out").click()
        self._page.wait_for_load_state("networkidle")
        log.info("Signed out successfully.")
