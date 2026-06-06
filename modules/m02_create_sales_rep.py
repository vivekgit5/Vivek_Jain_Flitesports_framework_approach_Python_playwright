"""
Module M02 – Create Sales Representative
==========================================
Logs in as CRM admin, opens the Add User modal, selects the Sales Rep
role, fills in Faker-generated identity and contact details, assigns
a partner and an account manager, submits the form, and signs out.

An XLSX execution report is written to the ``reports/`` directory on
completion (pass or fail).

Usage
-----
    python -m modules.m02_create_sales_rep
"""

import logging
import time
from dataclasses import dataclass

from faker import Faker
from playwright.sync_api import Playwright, expect, sync_playwright

from config.settings import CRM_ADMIN_EMAIL, CRM_ADMIN_PASS
from core.browser import browser_session
from core.reporter import XlsxReporter
from pages.login_page import LoginPage
from pages.users_page import UsersPage

log = logging.getLogger(__name__)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)

MODULE_NAME = "M02 – Create Sales Representative"

# Hard-coded reference data used for the partner and account-manager searches
_PARTNER_SEARCH_KEYWORD  = "VIVEK KRISTIN 995"
_ACCT_MGR_SEARCH_KEYWORD = "VIVEK DYLAN"
_ACCT_MGR_TARGET_NAME    = "VIVEK DYLAN RODRIGUEZ ACCOUNT MANAGER QA"


# ── Test-data model ────────────────────────────────────────────────────────────

@dataclass
class SalesRepData:
    first_name: str
    last_name:  str
    phone:      str
    email:      str


def _generate_sales_rep_data() -> SalesRepData:
    """Return a ``SalesRepData`` instance with millisecond-precise unique values.

    Naming convention
    -----------------
    * First name : ``Vivek <Faker first name>``
    * Last name  : ``<Faker last name> QA``
    * Email      : ``vivek+sr<tag>@bitcot.com``  (5-digit ms-based tag)
    * Phone      : US-formatted  ``+1 (###) ###-####``
    """
    fake = Faker("en_US")
    tag  = int(time.time() * 1_000) % 90_000 + 10_000

    data = SalesRepData(
        first_name=f"Vivek {fake.first_name()}",
        last_name=f"{fake.last_name()} QA",
        phone=fake.numerify("+1 (###) ###-####"),
        email=f"vivek+sr{tag}@bitcot.com",
    )
    log.info("Generated Sales Rep data → %s", data)
    return data


# ── Private helpers ────────────────────────────────────────────────────────────

def _close_dropdown(page, step_label: str) -> None:
    """Reliably close an open search dropdown by pressing Escape then clicking outside."""
    log.info("[%s] Closing dropdown …", step_label)
    page.keyboard.press("Escape")
    page.wait_for_timeout(600)
    page.mouse.click(760, 180)
    page.wait_for_timeout(700)
    page.mouse.click(760, 180)
    page.wait_for_timeout(500)
    page.mouse.wheel(0, 120)
    page.wait_for_timeout(500)


def _click_done_button(page, section, step_label: str) -> None:
    """Click the 'Done' button inside the active dropdown panel, or fall back to close."""
    modal    = page.locator("role=dialog")
    done_btn = modal.locator("button:has-text('Done')")
    try:
        done_btn.wait_for(state="visible", timeout=5_000)
        done_btn.click()
        page.wait_for_timeout(400)
    except Exception:
        log.warning("[%s] 'Done' button not found — falling back to close.", step_label)
        _close_dropdown(page, step_label)


def _search_and_select(
    page,
    placeholder: str,
    search_keyword: str,
    target_name: str | None,
    step_label: str,
) -> str:
    """
    Type *search_keyword* into a modal typeahead field, wait for results,
    click the first item whose name contains *target_name* (or the first
    result when *target_name* is None), close the dropdown, and return
    the selected name.

    Args:
        page:           Active Playwright page.
        placeholder:    Placeholder text of the search input.
        search_keyword: Text to type to trigger the API search.
        target_name:    Optional exact or partial name to match.  When None
                        the very first result is selected.
        step_label:     Label written to the execution report.

    Returns:
        The text of the selected item.
    """
    modal = page.locator("role=dialog")

    search_input = modal.get_by_placeholder(placeholder)
    search_input.wait_for(state="visible", timeout=8_000)
    search_input.scroll_into_view_if_needed()
    search_input.click()
    page.wait_for_timeout(500)

    search_input.press("Control+a")
    search_input.press("Delete")
    page.wait_for_timeout(200)
    search_input.press_sequentially(search_keyword, delay=80)
    log.info("[%s] Typed '%s' — waiting for API results …", step_label, search_keyword)
    page.wait_for_timeout(4_000)

    section      = search_input.locator("xpath=ancestor::div[4]")
    result_items = section.locator("span.srf-item-name")
    result_items.first.wait_for(state="visible", timeout=15_000)

    selected_name = ""
    if target_name:
        for item in result_items.all():
            text = (item.text_content() or "").strip().upper()
            if target_name.upper() in text:
                selected_name = text
                item.click()
                break

    if not selected_name:
        first_item    = result_items.first
        selected_name = (first_item.text_content() or "").strip()
        first_item.click()

    page.wait_for_timeout(7_000)
    _click_done_button(page, section, step_label)
    log.info("[%s] Selected: '%s'", step_label, selected_name)
    return selected_name


# ── Module orchestrator ────────────────────────────────────────────────────────

def run(playwright: Playwright) -> None:
    """
    Execute the full Create Sales Representative flow.

    Steps
    -----
    1.  Generate test data
    2.  Login as CRM admin
    3.  Navigate to Users → open Add User modal
    4.  Select Sales Rep role
    5.  Fill identity fields (first name, last name, title)
    6.  Fill contact fields (phone, email)
    7.  Select additional dropdown value
    8.  Assign partner (VIVEK KRISTIN 995)
    9.  Assign account manager (VIVEK DYLAN RODRIGUEZ)
    10. Submit form
    11. Sign out
    """
    reporter = XlsxReporter(MODULE_NAME)
    data     = _generate_sales_rep_data()
    reporter.add_step(
        "Generated Sales Rep test data",
        "INFO",
        f"{data.first_name} {data.last_name}  |  {data.email}  |  {data.phone}",
    )

    with browser_session(playwright, slow_mo=80) as page:
        try:
            login = LoginPage(page)
            users = UsersPage(page)
            modal = lambda: page.locator("role=dialog")

            # ── Login ─────────────────────────────────────────────────────────
            login.navigate()
            reporter.add_step("Navigate to CRM login page", "PASS", page.url)
            login.login(CRM_ADMIN_EMAIL, CRM_ADMIN_PASS)
            reporter.add_step("Authenticate as CRM admin", "PASS", CRM_ADMIN_EMAIL)

            # ── Users module ──────────────────────────────────────────────────
            users.navigate()
            reporter.add_step("Navigate to Users module", "PASS")
            users.open_add_user_modal()
            reporter.add_step("Open Add User modal", "PASS")

            # ── Role ──────────────────────────────────────────────────────────
            log.info("Selecting Sales Rep role …")
            role_select = modal().locator("select").first
            role_select.wait_for(state="visible", timeout=8_000)

            matched = False
            for opt in role_select.locator("option").all():
                if "sales" in (opt.text_content() or "").lower():
                    role_select.select_option(label=(opt.text_content() or "").strip())
                    matched = True
                    break
            if not matched:
                role_select.select_option(index=1)

            page.wait_for_timeout(800)
            backdrop = page.locator("div.gl-backdrop")
            if backdrop.count() > 0:
                try:
                    backdrop.wait_for(state="detached", timeout=5_000)
                except Exception:
                    pass
            page.mouse.click(760, 350)
            page.wait_for_timeout(800)
            reporter.add_step("Select Sales Rep role", "PASS")

            # ── Identity fields ───────────────────────────────────────────────
            page.wait_for_selector("div.srf-identity-fields", state="visible", timeout=15_000)

            first_input = modal().get_by_placeholder("JOHN")
            if first_input.count() == 0:
                first_input = modal().locator("div.srf-identity-fields > div:nth-of-type(1) input")
            first_input.first.fill(data.first_name)

            last_input = modal().get_by_placeholder("DOE")
            if last_input.count() == 0:
                last_input = modal().locator("div.srf-identity-fields > div:nth-of-type(2) input")
            last_input.first.fill(data.last_name)
            page.keyboard.press("Tab")
            reporter.add_step(
                "Fill First Name and Last Name",
                "PASS",
                f"{data.first_name} {data.last_name}",
            )

            title_sel = modal().locator("div.srf-identity-fields select")
            if title_sel.count() > 0:
                title_sel.first.select_option(value="2")
                reporter.add_step("Select title / prefix", "PASS", "Value: 2")

            # ── Contact fields ────────────────────────────────────────────────
            phone_input = modal().locator(
                "div:nth-of-type(2) > div.srf-section-body > div:nth-of-type(1) input"
            )
            phone_input.click()
            page.wait_for_timeout(400)
            page.keyboard.press("Escape")
            phone_input.click(click_count=3)
            phone_input.fill(data.phone)
            page.keyboard.press("Tab")
            reporter.add_step("Fill Phone number", "PASS", data.phone)

            email_input = modal().locator(
                "div:nth-of-type(2) > div.srf-section-body > div:nth-of-type(2) input"
            )
            email_input.click()
            email_input.fill(data.email)
            page.keyboard.press("Tab")
            reporter.add_step("Fill Email address", "PASS", data.email)

            # ── Additional dropdown ───────────────────────────────────────────
            extra_sel = modal().locator("div:nth-of-type(4) select")
            if extra_sel.count() > 0:
                extra_sel.first.select_option(value="5")
                reporter.add_step("Select additional dropdown value", "PASS", "Value: 5")

            # ── Assign partner ────────────────────────────────────────────────
            partner_name = _search_and_select(
                page,
                placeholder="Search partners...",
                search_keyword=_PARTNER_SEARCH_KEYWORD,
                target_name=_PARTNER_SEARCH_KEYWORD,
                step_label="Assign Partner",
            )
            reporter.add_step("Assign Partner", "PASS", partner_name)

            # ── Assign account manager ────────────────────────────────────────
            acct_mgr_name = _search_and_select(
                page,
                placeholder="Search account managers...",
                search_keyword=_ACCT_MGR_SEARCH_KEYWORD,
                target_name=_ACCT_MGR_TARGET_NAME,
                step_label="Assign Account Manager",
            )
            reporter.add_step("Assign Account Manager", "PASS", acct_mgr_name)

            # ── Submit ────────────────────────────────────────────────────────
            log.info("Submitting the Create Sales Rep form …")
            create_btn = page.locator("button.srf-btn--submit")
            create_btn.scroll_into_view_if_needed()
            expect(create_btn).to_be_visible(timeout=10_000)
            expect(create_btn).to_be_enabled(timeout=10_000)
            create_btn.click()
            page.wait_for_load_state("networkidle")
            reporter.add_step("Submit Create Sales Rep form", "PASS")

            # ── Sign out ──────────────────────────────────────────────────────
            login.sign_out()
            reporter.add_step("Sign out", "PASS")

            reporter.add_step("Module completed successfully", "PASS")
            log.info("✓  %s  —  all steps passed.", MODULE_NAME)

        except Exception as exc:
            reporter.add_step("Module encountered an unhandled error", "FAIL", str(exc))
            log.exception("Module failed: %s", exc)
            raise

        finally:
            report_path = reporter.save()
            log.info("Report → %s", report_path)


if __name__ == "__main__":
    with sync_playwright() as pw:
        run(pw)
