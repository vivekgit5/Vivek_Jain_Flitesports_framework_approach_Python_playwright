"""
Module M03 – Create New Partner (Super-Admin)
==============================================
Logs in as super-admin, opens the Add User wizard, and completes all five
wizard steps to create a new Partner/Club account.

Wizard steps
------------
  Step 1 – Program Identity  : name, partner type, sport, contact, products,
                                contract details
  Step 2 – Commission Rates  : wholesale / retail / suggested-retail margins
  Step 3 – Rep Assignment    : search & select a sales rep, set commission
  Step 4 – Partner Commission: per-rep and partner commission percentages
  Step 5 – Season Dates      : season start date → SUBMIT

An XLSX execution report is written to the ``reports/`` directory on
completion (pass or fail).

Usage
-----
    python -m modules.m03_create_new_partner
"""

import datetime
import logging
import random
from dataclasses import dataclass

from faker import Faker
from playwright.sync_api import Playwright, expect, sync_playwright

from config.settings import (
    COMMISSION_PARTNER,
    COMMISSION_REP,
    COMMISSION_RETAIL,
    COMMISSION_SUGGESTED_RETAIL,
    COMMISSION_WHOLESALE,
    CONTRACT_ENABLED,
    CRM_ADMIN_EMAIL,
    CRM_ADMIN_PASS,
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

MODULE_NAME = "M03 – Create New Partner (Super-Admin)"


# ── Test-data model ────────────────────────────────────────────────────────────

@dataclass
class PartnerData:
    program_name:      str
    first_name:        str
    last_name:         str
    phone:             str
    email:             str
    contract_end_date: str
    season_start_date: str


def _generate_partner_data() -> PartnerData:
    """Return a ``PartnerData`` instance with Faker-generated values.

    Dates
    -----
    * Contract end date : today + 365 days
    * Season start date : today
    """
    fake       = Faker("en_US")
    unique_tag = random.randint(1_000, 99_999)
    today      = datetime.date.today()

    data = PartnerData(
        program_name      = f"QA Partner {fake.company()} {unique_tag}",
        first_name        = f"QA {fake.first_name()}",
        last_name         = f"{fake.last_name()} Partner",
        phone             = fake.numerify("+1 (###) ###-####"),
        email             = f"vivek+p{unique_tag}@bitcot.com",
        contract_end_date = (today + datetime.timedelta(days=365)).strftime("%Y-%m-%d"),
        season_start_date = today.strftime("%Y-%m-%d"),
    )
    log.info("Generated partner data → %s", data)
    return data


# ── Shared wizard helper ───────────────────────────────────────────────────────

def _click_continue(page) -> None:
    """Click the primary CONTINUE button inside the wizard modal."""
    page.locator("button.msf-btn--primary", has_text="CONTINUE").click()
    page.wait_for_timeout(600)
    log.info("Clicked CONTINUE.")


# ── Wizard step implementations ────────────────────────────────────────────────

def _step1_program_identity(page, partner: PartnerData, reporter: XlsxReporter) -> None:
    """Wizard Step 1 – Program Identity."""
    log.info("── Wizard Step 1: Program Identity ──")
    modal = page.locator("role=dialog")

    # Select PARTNER/CLUB role to activate the wizard form
    modal.locator("select").first.select_option(label="PARTNER/CLUB")
    page.wait_for_selector("div.msf-identity-fields input", state="visible", timeout=15_000)
    reporter.add_step("Select PARTNER/CLUB user role", "PASS")

    # Program name
    modal.locator("div.msf-identity-fields > div:nth-of-type(1) input").fill(partner.program_name)
    reporter.add_step("Enter Program Name", "PASS", partner.program_name)

    # Partner type (PrimeVue combobox)
    modal.locator("div.msf-identity-fields").get_by_role("combobox").first.click()
    page.wait_for_timeout(400)
    page.get_by_role("option", name="partner-CAMP").click()
    reporter.add_step("Select partner type", "PASS", "partner-CAMP")

    # Sport multiselect – BASEBALL
    modal.locator(
        "div.msf-panel > div:nth-of-type(1) div.p-multiselect-label-container > div"
    ).click()
    page.wait_for_selector("[role='option'][aria-label='BASEBALL']", state="visible", timeout=8_000)
    page.locator("[role='option'][aria-label='BASEBALL']").click()
    modal.locator("div.msf-panel").click(position={"x": 533, "y": 287})
    page.wait_for_timeout(300)
    reporter.add_step("Select sport", "PASS", "BASEBALL")

    # Admin contact – first name
    modal.locator(
        "div:nth-of-type(3) > div.msf-section-body > div:nth-of-type(1) input"
    ).fill(partner.first_name)
    reporter.add_step("Enter admin contact first name", "PASS", partner.first_name)

    # Admin contact – last name
    modal.locator(
        "div:nth-of-type(3) > div.msf-section-body > div:nth-of-type(2) input"
    ).fill(partner.last_name)
    reporter.add_step("Enter admin contact last name", "PASS", partner.last_name)

    # Admin contact – phone
    phone_input = modal.locator(
        "div:nth-of-type(3) > div.msf-section-body > div:nth-of-type(3) input"
    )
    phone_input.click()
    page.wait_for_timeout(300)
    page.keyboard.press("Escape")
    phone_input.click(click_count=3)
    phone_input.fill(partner.phone)
    page.keyboard.press("Tab")
    reporter.add_step("Enter admin contact phone", "PASS", partner.phone)

    # Admin contact – email
    email_input = modal.locator("div:nth-of-type(3) div:nth-of-type(4) input")
    email_input.click(click_count=3)
    email_input.fill(partner.email)
    page.keyboard.press("Tab")
    reporter.add_step("Enter admin contact email", "PASS", partner.email)

    # Dismiss floating suggestion pickers
    page.mouse.click(613, 324)
    page.wait_for_timeout(300)

    # Catalogue-size dropdown
    modal.locator("div.msf-section-body > div:nth-of-type(4) select").select_option("5")
    reporter.add_step("Select catalogue size", "PASS", "5")

    # Level dropdown
    modal.locator("div:nth-of-type(5) select").select_option("2")
    reporter.add_step("Select level", "PASS", "2")

    # League multiselect – select first available option
    modal.locator(".p-multiselect").nth(1).locator(".p-multiselect-label-container").click()
    page.wait_for_selector(
        "li.p-multiselect-item, li.p-multiselect-option",
        state="visible",
        timeout=10_000,
    )
    items = page.locator("li.p-multiselect-item, li.p-multiselect-option")
    first_label = (items.first.text_content() or "").strip()
    items.first.click()
    page.keyboard.press("Escape")
    page.wait_for_timeout(400)
    reporter.add_step("Select League", "PASS", first_label)

    # Contract toggle + end date
    modal.locator("div:nth-of-type(8) select").select_option(CONTRACT_ENABLED)
    reporter.add_step("Enable contract (Yes)", "PASS")

    modal.locator(
        "div:nth-of-type(8) div.msf-grid-2 > div:nth-of-type(2) input"
    ).fill(partner.contract_end_date)
    reporter.add_step("Enter contract end date", "PASS", partner.contract_end_date)

    _click_continue(page)
    reporter.add_step("Complete Wizard Step 1 – Program Identity", "PASS")


def _step2_commission_rates(page, reporter: XlsxReporter) -> None:
    """Wizard Step 2 – Commission Rates."""
    log.info("── Wizard Step 2: Commission Rates ──")
    modal = page.locator("role=dialog")

    modal.locator(
        "div:nth-of-type(1) > div.msf-section-body > div > div:nth-of-type(1) input"
    ).fill(COMMISSION_WHOLESALE)
    reporter.add_step("Enter wholesale commission", "PASS", f"{COMMISSION_WHOLESALE}%")

    modal.locator(
        "div:nth-of-type(2) > div.msf-section-body > div > div:nth-of-type(1) input"
    ).fill(COMMISSION_RETAIL)
    reporter.add_step("Enter retail commission", "PASS", f"{COMMISSION_RETAIL}%")

    modal.locator("div:nth-of-type(3) div.msf-grid-2 input").fill(COMMISSION_SUGGESTED_RETAIL)
    reporter.add_step("Enter suggested-retail commission", "PASS", f"{COMMISSION_SUGGESTED_RETAIL}%")

    _click_continue(page)
    reporter.add_step("Complete Wizard Step 2 – Commission Rates", "PASS")


def _step3_rep_assignment(page, reporter: XlsxReporter) -> None:
    """Wizard Step 3 – Sales Rep Assignment."""
    log.info("── Wizard Step 3: Rep Assignment ──")
    modal = page.locator("role=dialog")

    modal.locator("div.msf-rep-dropdown-container input").click()
    page.wait_for_timeout(600)

    first_rep_label = modal.locator("div.msf-rep-dropdown-container label").first
    rep_name        = first_rep_label.locator("span.msf-rep-name").inner_text()
    first_rep_label.click()
    reporter.add_step("Select first available sales rep", "PASS", rep_name)

    modal.locator("div.msf-rep-dropdown-list button", has_text="Done").click()
    page.wait_for_timeout(400)
    reporter.add_step("Confirm rep selection", "PASS")

    _click_continue(page)

    modal.locator("div.msf-rep-selected-configs input").fill(COMMISSION_REP)
    reporter.add_step("Enter rep commission", "PASS", f"{COMMISSION_REP}%")

    modal.locator("div.pc-section input").fill(COMMISSION_PARTNER)
    reporter.add_step("Enter partner commission", "PASS", f"{COMMISSION_PARTNER}%")

    modal.locator(
        "div.msf-grid-2 > div:nth-of-type(2) span", has_text="Search & select products"
    ).click()
    page.wait_for_timeout(400)
    modal.locator("div.pc-section button", has_text="Done").click()
    page.wait_for_timeout(300)
    reporter.add_step("Confirm product selection (defaults retained)", "PASS")

    _click_continue(page)
    reporter.add_step("Complete Wizard Step 3 – Rep Assignment", "PASS")


def _step4_partner_commissions(page, reporter: XlsxReporter) -> None:
    """Wizard Step 4 – Partner Commission Details."""
    log.info("── Wizard Step 4: Partner Commission Details ──")
    modal = page.locator("role=dialog")

    pc_input = modal.locator("div.pc-section input")
    pc_input.click(click_count=3)
    pc_input.fill(COMMISSION_PARTNER)
    reporter.add_step("Update partner commission", "PASS", f"{COMMISSION_PARTNER}%")

    rep_input = modal.locator("div.msf-rep-selected-configs input")
    rep_input.click(click_count=3)
    rep_input.fill(COMMISSION_SUGGESTED_RETAIL)
    reporter.add_step("Update rep commission", "PASS", f"{COMMISSION_SUGGESTED_RETAIL}%")

    _click_continue(page)
    reporter.add_step("Complete Wizard Step 4 – Partner Commission Details", "PASS")


def _step5_season_and_submit(page, partner: PartnerData, reporter: XlsxReporter) -> None:
    """Wizard Step 5 – Season Dates and final submission."""
    log.info("── Wizard Step 5: Season Dates ──")
    modal = page.locator("role=dialog")

    season_input = modal.locator("div.msf-season-header div:nth-of-type(1) > input")
    season_input.click(click_count=3)
    season_input.fill(partner.season_start_date)
    reporter.add_step("Enter season start date", "PASS", partner.season_start_date)

    submit_btn = page.locator("button.msf-btn--submit")
    submit_btn.scroll_into_view_if_needed()
    expect(submit_btn).to_be_visible(timeout=10_000)
    expect(submit_btn).to_be_enabled(timeout=10_000)
    submit_btn.click()
    page.wait_for_load_state("networkidle")
    reporter.add_step("Submit partner creation wizard — Partner created", "PASS")


# ── Module orchestrator ────────────────────────────────────────────────────────

def run(playwright: Playwright) -> None:
    """
    Execute the full Create New Partner flow.

    Steps
    -----
    1.  Generate test data
    2.  Login as CRM super-admin
    3.  Navigate to Users module → open Add User wizard
    4.  Wizard Step 1  – Program Identity
    5.  Wizard Step 2  – Commission Rates
    6.  Wizard Step 3  – Rep Assignment
    7.  Wizard Step 4  – Partner Commission Details
    8.  Wizard Step 5  – Season Dates → Submit
    9.  Sign out
    """
    reporter = XlsxReporter(MODULE_NAME)
    partner  = _generate_partner_data()
    reporter.add_step(
        "Generated test partner data",
        "INFO",
        (
            f"Program: {partner.program_name}  |  "
            f"Contact: {partner.first_name} {partner.last_name}  |  "
            f"Email: {partner.email}"
        ),
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
            users.open_add_user_modal()
            reporter.add_step("Open Add User wizard", "PASS")

            _step1_program_identity(page, partner, reporter)
            _step2_commission_rates(page, reporter)
            _step3_rep_assignment(page, reporter)
            _step4_partner_commissions(page, reporter)
            _step5_season_and_submit(page, partner, reporter)

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
