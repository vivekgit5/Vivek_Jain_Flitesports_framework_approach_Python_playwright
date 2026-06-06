"""
Page object for the Flitesports CRM Users module.

Covers navigation, the add-user modal, role-based filtering, user search,
and opening the edit modal for the first result in the table.
"""

import logging

from playwright.sync_api import Page

log = logging.getLogger(__name__)


class UsersPage:
    """
    Encapsulates all interactions with the CRM Users section.

    Supported actions
    -----------------
    * ``navigate()``             — click the USERS nav link
    * ``open_add_user_modal()``  — click the top-bar ADD USER button
    * ``filter_by_role(value)``  — select a role from the filter dropdown
    * ``search_user(query)``     — type into the search input
    * ``click_first_row_edit()`` — click the pencil icon on the first table row
    """

    def __init__(self, page: Page) -> None:
        self._page = page

    # ── Selectors ──────────────────────────────────────────────────────────────
    _SEL_NAV_USERS    = "USERS"
    _SEL_ADD_BUTTON   = "div.crm-topbar__actions > button"
    _SEL_ROLE_FILTER  = "#app select:nth-of-type(1)"
    _SEL_SEARCH_INPUT = (
        "xpath=//*[@id='app']/div[1]/div/main/section/div[1]/div[2]/div/input"
    )
    _SEL_EDIT_BUTTON  = "tr:nth-of-type(1) button.ico-btn--edit"

    # ── Public interface ───────────────────────────────────────────────────────

    def navigate(self) -> None:
        """Click the USERS navigation link and wait for the page to load."""
        log.info("Navigating to the Users module …")
        self._page.get_by_text(self._SEL_NAV_USERS).first.click()
        self._page.wait_for_load_state("networkidle")

    def open_add_user_modal(self) -> None:
        """
        Click the ADD USER button in the top-bar action area.
        Waits until the modal dialog is visible before returning.
        """
        log.info("Opening the Add User modal …")
        self._page.locator(self._SEL_ADD_BUTTON).click()
        self._page.wait_for_selector("role=dialog", state="visible", timeout=10_000)

    def filter_by_role(self, role_value: str) -> None:
        """
        Select a role from the Users page filter dropdown.

        Args:
            role_value: The HTML ``<option>`` value string.
                        Use the constants defined in ``config.settings``
                        (``ROLE_FILTER_ADMIN``, ``ROLE_FILTER_SALES_REP``, etc.).
        """
        log.info("Filtering users by role value: %s", role_value)
        self._page.locator(self._SEL_ROLE_FILTER).select_option(role_value)
        self._page.wait_for_load_state("networkidle")

    def search_user(self, query: str) -> None:
        """
        Type *query* into the user search input and allow the table to update.

        Args:
            query: Search string (e.g. ``"vivek"``).
        """
        log.info("Searching for user: '%s' …", query)
        search = self._page.locator(self._SEL_SEARCH_INPUT)
        search.click()
        search.fill(query)
        self._page.wait_for_timeout(800)

    def click_first_row_edit(self) -> None:
        """
        Click the edit (pencil) icon on the first row of the users table
        and wait for the edit modal to become visible.
        """
        log.info("Opening edit modal for the first result …")
        self._page.locator(self._SEL_EDIT_BUTTON).first.click()
        self._page.wait_for_selector("role=dialog", state="visible", timeout=15_000)
        self._page.wait_for_timeout(400)
