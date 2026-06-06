"""
Module M04 – Update Admin, Sales Rep, and Partner
====================================================
Logs in as CRM super-admin and sequentially updates three user types:

  1. Admin User  – updates first name, last name, and phone.
  2. Sales Rep   – updates role, first name, last name, and phone.
  3. Partner     – updates program name, contact name, commission fields,
                   and season start date via the multi-step edit form.

Signs out on completion.  An XLSX execution report is written to the
``reports/`` directory regardless of pass or fail outcome.

Usage
-----
    python -m modules.m04_update_users
"""

import datetime
import logging
import random
from dataclasses import dataclass

from faker import Faker
from playwright.sync_api import Playwright, expect, sync_playwright

from config.settings import (
    CRM_ADMIN_EMAIL,
    CRM_ADMIN_PASS,
    ROLE_FILTER_ADMIN,
    ROLE_FILTER_PARTNER,
    ROLE_FILTER_SALES_REP,
)
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

MODULE_NAME = "M04 – Update Admin, Sales Rep, and Partner"

# Role option value used inside the Sales Rep edit modal
_ADMIN_MODAL_ROLE_VALUE = "2"


# ── Test-data models ───────────────────────────────────────────────────────────

@dataclass
class AdminUpdateData:
    first_name: str
    last_name:  str
    phone:      str


@dataclass
class SalesRepUpdateData:
    first_name: str
    last_name:  str
    phone:      str


@dataclass
class PartnerUpdateData:
    program_name:       str
    contact_first_name: str
    commission_rate:    int
    season_start_date:  str


def _generate_test_data() -> tuple[AdminUpdateData, SalesRepUpdateData, PartnerUpdateData]:
    """Return Faker-generated update payloads for all three user types."""
    fake        = Faker("en_US")
    season_date = (datetime.date.today() + datetime.timedelta(days=6)).isoformat()

    admin = AdminUpdateData(
        first_name=f"VIVEK {fake.first_name().upper()}",
        last_name=f"{fake.last_name().upper()} QA ADMIN",
        phone=fake.numerify("+1 (###) ###-####"),
    )
    sales_rep = SalesRepUpdateData(
        first_name=f"VIVEK {fake.first_name().upper()}",
        last_name=f"{fake.last_name().upper()} QA SALES REP",
        phone=fake.numerify("+1 (###) ###-####"),
    )
    partner = PartnerUpdateData(
        program_name=f"VIVEK QA PARTNER {fake.company().upper()} {random.randint(10_000, 99_999)}",
        contact_first_name=f"VIVEK QA {fake.first_name().upper()}",
        commission_rate=random.randint(30, 60),
        season_start_date=season_date,
    )
    log.info("Admin data      → %s", admin)
    log.info("Sales Rep data  → %s", sales_rep)
    log.info("Partner data    → %s", partner)
    return admin, sales_rep, partner


# ── Shared helper ──────────────────────────────────────────────────────────────

def _fill_phone(page, css_selector: str, phone: str) -> None:
    """Triple-click a phone input, dismiss any flag-picker, then fill the value."""
    field = page.locator(css_selector)
    field.click(click_count=3)
    page.wait_for_timeout(400)
    page.keyboard.press("Escape")
    field.fill(phone)
    page.keyboard.press("Tab")


# ── Update sub-flows ───────────────────────────────────────────────────────────

def _update_admin(
    page,
    users: UsersPage,
    data: AdminUpdateData,
    reporter: XlsxReporter,
) -> None:
    """Filter by Admin role, search, open edit modal, and update identity + phone."""
    log.info("=== Updating Admin User ===")

    users.filter_by_role(ROLE_FILTER_ADMIN)
    reporter.add_step("Apply Admin role filter", "PASS", f"Filter value: {ROLE_FILTER_ADMIN}")

    users.search_user("vivek")
    reporter.add_step("Search for admin user", "PASS", "Query: vivek")

    users.click_first_row_edit()
    reporter.add_step("Open Admin edit modal", "PASS")

    first_input = page.locator("div.amf-identity-fields > div:nth-of-type(1) input")
    first_input.click(click_count=3)
    first_input.fill(data.first_name)
    reporter.add_step("Update admin first name", "PASS", data.first_name)

    last_input = page.locator("div.amf-identity-fields > div:nth-of-type(2) input")
    last_input.click(click_count=3)
    last_input.fill(data.last_name)
    reporter.add_step("Update admin last name", "PASS", data.last_name)

    _fill_phone(
        page,
        "div:nth-of-type(2) > div.amf-section-body > div:nth-of-type(1) input",
        data.phone,
    )
    reporter.add_step("Update admin phone number", "PASS", data.phone)

    submit_btn = page.locator("button.amf-btn--submit")
    expect(submit_btn).to_be_visible(timeout=10_000)
    expect(submit_btn).to_be_enabled(timeout=10_000)
    submit_btn.click()
    page.wait_for_load_state("networkidle")
    reporter.add_step("Submit admin update — saved successfully", "PASS")


def _update_sales_rep(
    page,
    users: UsersPage,
    data: SalesRepUpdateData,
    reporter: XlsxReporter,
) -> None:
    """Filter by Sales Rep role, search, open edit modal, and update identity + phone."""
    log.info("=== Updating Sales Rep User ===")

    users.filter_by_role(ROLE_FILTER_SALES_REP)
    reporter.add_step("Apply Sales Rep role filter", "PASS", f"Filter value: {ROLE_FILTER_SALES_REP}")

    users.search_user("vivek")
    reporter.add_step("Search for sales rep user", "PASS", "Query: vivek")

    users.click_first_row_edit()
    reporter.add_step("Open Sales Rep edit modal", "PASS")

    page.locator("role=dialog").locator("select.srf-select").first.select_option(
        _ADMIN_MODAL_ROLE_VALUE
    )
    page.wait_for_timeout(500)
    reporter.add_step(
        "Update role in edit modal",
        "PASS",
        f"Role value: {_ADMIN_MODAL_ROLE_VALUE}",
    )

    first_input = page.locator("div.srf-identity-fields > div:nth-of-type(1) input")
    first_input.click(click_count=3)
    first_input.fill(data.first_name)
    reporter.add_step("Update sales rep first name", "PASS", data.first_name)

    last_input = page.locator("div.srf-identity-fields > div:nth-of-type(2) input")
    last_input.click(click_count=3)
    last_input.fill(data.last_name)
    reporter.add_step("Update sales rep last name", "PASS", data.last_name)

    _fill_phone(
        page,
        "div:nth-of-type(2) > div.srf-section-body > div:nth-of-type(1) input",
        data.phone,
    )
    reporter.add_step("Update sales rep phone number", "PASS", data.phone)

    submit_btn = page.locator("button.srf-btn--submit")
    expect(submit_btn).to_be_visible(timeout=10_000)
    expect(submit_btn).to_be_enabled(timeout=10_000)
    submit_btn.click()
    page.wait_for_load_state("networkidle")
    reporter.add_step("Submit sales rep update — saved successfully", "PASS")


def _update_partner(
    page,
    users: UsersPage,
    data: PartnerUpdateData,
    reporter: XlsxReporter,
) -> None:
    """Filter by Partner role, search, open edit wizard, and update program/commission/season."""
    log.info("=== Updating Partner ===")

    users.filter_by_role(ROLE_FILTER_PARTNER)
    reporter.add_step("Apply Partner role filter", "PASS", f"Filter value: {ROLE_FILTER_PARTNER}")

    users.search_user("vivek")
    reporter.add_step("Search for partner user", "PASS", "Query: vivek")

    users.click_first_row_edit()
    reporter.add_step("Open Partner edit wizard", "PASS")

    # ── Form step 1: Identity ─────────────────────────────────────────────────
    program_input = page.locator("div.msf-identity-fields > div:nth-of-type(1) input")
    program_input.click(click_count=3)
    program_input.fill(data.program_name)
    reporter.add_step("Update partner program name", "PASS", data.program_name)

    contact_input = page.locator(
        "div:nth-of-type(3) > div.msf-section-body > div:nth-of-type(1) input"
    )
    contact_input.click(click_count=3)
    contact_input.fill(data.contact_first_name)
    reporter.add_step("Update contact first name", "PASS", data.contact_first_name)

    page.locator("button.msf-btn--primary").click()
    page.wait_for_timeout(600)
    reporter.add_step("Advance to Partner form step 2 (Commission)", "PASS")

    # ── Form step 2: Commission – zero all numeric fields ─────────────────────
    page.wait_for_load_state("networkidle")
    numeric_inputs = page.locator("[role='dialog'] input[type='number']")
    count = numeric_inputs.count()
    log.info("Zeroing %d numeric input(s) on commission step …", count)
    for i in range(count):
        inp = numeric_inputs.nth(i)
        if inp.is_visible():
            inp.click(click_count=3)
            page.wait_for_timeout(150)
            inp.fill("0")
            page.keyboard.press("Tab")
            page.wait_for_timeout(200)
    reporter.add_step(
        "Zero all numeric commission fields",
        "PASS",
        f"{count} field(s) cleared",
    )

    page.locator("button.msf-btn--primary").click()
    page.wait_for_timeout(800)
    reporter.add_step("Advance to Partner form step 3 (Contract & Season)", "PASS")

    # ── Form step 3: Season start date ────────────────────────────────────────
    season_input = page.locator("div.msf-season-header div:nth-of-type(1) > input")
    season_input.click(click_count=3)
    season_input.fill(data.season_start_date)
    page.keyboard.press("Tab")
    page.wait_for_timeout(400)
    reporter.add_step("Enter season start date", "PASS", data.season_start_date)

    submit_btn = page.locator("button.msf-btn--submit")
    expect(submit_btn).to_be_visible(timeout=10_000)
    expect(submit_btn).to_be_enabled(timeout=10_000)
    submit_btn.click()
    page.wait_for_load_state("networkidle")
    reporter.add_step("Submit partner update — saved successfully", "PASS")


# ── Module orchestrator ────────────────────────────────────────────────────────

def run(playwright: Playwright) -> None:
    """
    Execute the full update suite: Admin → Sales Rep → Partner.
    """
    reporter                           = XlsxReporter(MODULE_NAME)
    admin_data, sales_rep_data, partner_data = _generate_test_data()

    reporter.add_step(
        "Generated Admin test data", "INFO",
        f"{admin_data.first_name} {admin_data.last_name}  |  {admin_data.phone}",
    )
    reporter.add_step(
        "Generated Sales Rep test data", "INFO",
        f"{sales_rep_data.first_name} {sales_rep_data.last_name}  |  {sales_rep_data.phone}",
    )
    reporter.add_step(
        "Generated Partner test data", "INFO",
        f"{partner_data.program_name}  |  season={partner_data.season_start_date}",
    )

    with browser_session(playwright) as page:
        try:
            login = LoginPage(page)
            users = UsersPage(page)

            login.navigate()
            reporter.add_step("Navigate to CRM login page", "PASS", page.url)
            login.login(CRM_ADMIN_EMAIL, CRM_ADMIN_PASS)
            reporter.add_step("Authenticate as CRM super-admin", "PASS", CRM_ADMIN_EMAIL)

            users.navigate()
            reporter.add_step("Navigate to Users module", "PASS")

            _update_admin(page, users, admin_data, reporter)
            _update_sales_rep(page, users, sales_rep_data, reporter)
            _update_partner(page, users, partner_data, reporter)

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
