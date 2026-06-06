"""
Module M01 – Create Admin User
================================
Logs in as CRM admin, opens the Add User modal, selects the Admin role,
fills in Faker-generated contact details, submits the form, and signs out.

An XLSX execution report is written to the ``reports/`` directory on
completion (pass or fail).

Usage
-----
    python -m modules.m01_create_admin_user
"""

import logging
import random
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

MODULE_NAME = "M01 – Create Admin User"


# ── Test-data model ────────────────────────────────────────────────────────────

@dataclass
class AdminUserData:
    first_name: str
    last_name:  str
    phone:      str
    email:      str


def _generate_admin_user_data() -> AdminUserData:
    """Return an ``AdminUserData`` instance with Faker-generated values.

    Naming convention
    -----------------
    * First name : ``Vivek <Faker first name>``
    * Last name  : ``<Faker last name> QA Admin``
    * Email      : ``vivek+<5-digit-tag>@bitcot.com``  (always unique)
    * Phone      : US-formatted  ``+1 (###) ###-####``
    """
    fake = Faker("en_US")
    tag  = random.randint(10_000, 99_999)

    data = AdminUserData(
        first_name=f"Vivek {fake.first_name()}",
        last_name=f"{fake.last_name()} QA Admin",
        phone=fake.numerify("+1 (###) ###-####"),
        email=f"vivek+{tag}@bitcot.com",
    )
    log.info("Generated admin user data → %s", data)
    return data


# ── Module orchestrator ────────────────────────────────────────────────────────

def run(playwright: Playwright) -> None:
    """
    Execute the full Create Admin User flow.

    Steps
    -----
    1.  Generate test data
    2.  Navigate to CRM login and authenticate
    3.  Navigate to Users module
    4.  Open the Add User modal
    5.  Select the Admin role
    6.  Fill First Name, Last Name, Phone, Email
    7.  Select an additional dropdown value
    8.  Submit the form
    9.  Sign out
    """
    reporter = XlsxReporter(MODULE_NAME)
    user     = _generate_admin_user_data()
    reporter.add_step(
        "Generated test user data",
        "INFO",
        f"{user.first_name} {user.last_name}  |  {user.email}  |  {user.phone}",
    )

    with browser_session(playwright) as page:
        try:
            login = LoginPage(page)
            users = UsersPage(page)
            modal = lambda: page.locator("role=dialog")

            # ── Login ─────────────────────────────────────────────────────────
            login.navigate()
            reporter.add_step("Navigate to CRM login page", "PASS", f"{page.url}")

            login.login(CRM_ADMIN_EMAIL, CRM_ADMIN_PASS)
            reporter.add_step("Authenticate as CRM admin", "PASS", CRM_ADMIN_EMAIL)

            # ── Users module ──────────────────────────────────────────────────
            users.navigate()
            reporter.add_step("Navigate to Users module", "PASS")

            users.open_add_user_modal()
            reporter.add_step("Open Add User modal", "PASS")

            # ── Role selection ────────────────────────────────────────────────
            log.info("Selecting user role (index 1) …")
            modal().locator("select").first.select_option(index=1)
            reporter.add_step("Select Admin user role", "PASS")

            # ── Identity fields ───────────────────────────────────────────────
            log.info("Filling First Name: %s", user.first_name)
            modal().locator(
                "div.amf-identity-fields > div:nth-of-type(1) input"
            ).fill(user.first_name)

            log.info("Filling Last Name: %s", user.last_name)
            modal().locator(
                "div.amf-identity-fields > div:nth-of-type(2) input"
            ).fill(user.last_name)
            page.keyboard.press("Tab")
            reporter.add_step(
                "Fill First Name and Last Name",
                "PASS",
                f"{user.first_name} {user.last_name}",
            )

            # ── Phone ─────────────────────────────────────────────────────────
            log.info("Filling Phone: %s", user.phone)
            phone_input = modal().locator(
                "div:nth-of-type(2) > div.amf-section-body > div:nth-of-type(1) input"
            )
            phone_input.click()
            page.wait_for_timeout(400)
            page.keyboard.press("Escape")
            phone_input.click(click_count=3)
            phone_input.fill(user.phone)
            page.keyboard.press("Tab")
            reporter.add_step("Fill Phone number", "PASS", user.phone)

            # ── Email ─────────────────────────────────────────────────────────
            log.info("Filling Email: %s", user.email)
            email_input = modal().locator(
                "div:nth-of-type(2) > div.amf-section-body > div:nth-of-type(2) input"
            )
            email_input.click()
            email_input.fill(user.email)
            page.keyboard.press("Tab")
            reporter.add_step("Fill Email address", "PASS", user.email)

            # ── Additional dropdown ───────────────────────────────────────────
            log.info("Selecting additional dropdown value (5) …")
            modal().locator("div:nth-of-type(4) select").select_option("5")
            reporter.add_step("Select additional dropdown value", "PASS", "Value: 5")

            # ── Submit ────────────────────────────────────────────────────────
            log.info("Submitting the Add User form …")
            submit_btn = page.locator("button.amf-btn--submit")
            submit_btn.scroll_into_view_if_needed()
            expect(submit_btn).to_be_visible(timeout=10_000)
            expect(submit_btn).to_be_enabled(timeout=10_000)
            submit_btn.click()
            page.wait_for_load_state("networkidle")
            reporter.add_step("Submit Add User form — Admin user created", "PASS")

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
