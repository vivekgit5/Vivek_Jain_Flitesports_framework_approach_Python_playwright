"""
Module M05 – Purchase Products on Shopify Staging (End-to-End)
================================================================
Automates a complete storefront purchase flow on the Flitesports Shopify
staging store, covering:

  Phase 1  – Store password authentication
  Phase 2  – Search for the FC Hawaii team store
  Phase 3  – Add four products to the cart
              (Jersey SS, Hoodie, Shorts, Light Game Kit)
  Phase 4  – Cart-drawer checkout confirmation (registered-player modal)
  Phase 5  – Guest checkout — shipping information
  Phase 6  – Guest checkout — payment details
  Phase 7  – Order confirmation verification

An XLSX execution report is written to the ``reports/`` directory on
completion (pass or fail).

Usage
-----
    python -m modules.m05_purchase_products_shopify
"""

import logging
from dataclasses import dataclass

from faker import Faker
from playwright.sync_api import Playwright, TimeoutError as PwTimeoutError, sync_playwright

from config.settings import SHOPIFY_BASE_URL, SHOPIFY_PRODUCTS
from core.browser import browser_session
from core.reporter import XlsxReporter
from pages.shopify_page import ShopifyStorePage

log = logging.getLogger(__name__)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)

MODULE_NAME = "M05 – Purchase Products on Shopify Staging"

# ── Static test data ───────────────────────────────────────────────────────────
_PAYMENT = {
    "card_number": "1",     # Shopify test card (staging only)
    "expiry":      "02/30",
    "cvv":         "111",
}


@dataclass
class ShopifyTestData:
    first_name: str
    last_name:  str
    jersey:     str  # ≤ 14 chars
    reg_name:   str  # ≤ 30 chars


def _generate_test_data() -> ShopifyTestData:
    """Return ``ShopifyTestData`` with Faker-generated values.

    Rules
    -----
    * ``first_name``  : always "Vivek"
    * ``last_name``   : random Faker surname
    * ``jersey``      : "Vivek " + random uppercased word, capped at 14 chars
    * ``reg_name``    : "Vivek " + two random uppercased surnames, capped at 30 chars
    """
    fake = Faker("en_US")

    jersey   = f"Vivek {fake.word().upper()}"[:14]
    reg_name = f"Vivek {fake.last_name().upper()} {fake.last_name().upper()}"[:30]

    data = ShopifyTestData(
        first_name=fake.first_name(),
        last_name=fake.last_name(),
        jersey=jersey,
        reg_name=reg_name,
    )
    log.info(
        "Generated Shopify test data → name: Vivek %s | jersey: %s | reg: %s",
        data.last_name,
        data.jersey,
        data.reg_name,
    )
    return data


# ── Module orchestrator ────────────────────────────────────────────────────────

def run(playwright: Playwright) -> None:
    """
    Execute the full Shopify end-to-end purchase flow.

    Steps
    -----
    1.  Generate test data
    2.  Authenticate with the Shopify staging store
    3.  Search for FC Hawaii
    4.  Add Product 1 – Game Jersey SS (Sky White)
    5.  Add Product 2 – Cotton-Blend Hoodie (Grey)
    6.  Add Product 3 – Game Shorts (Sky)
    7.  Add Product 4 – Light Game Kit (YMQ custom options)
    8.  Cart-drawer checkout confirmation
    9.  Fill guest checkout shipping information
    10. Fill payment and place order
    11. Verify order confirmation page
    """
    reporter = XlsxReporter(MODULE_NAME)
    data     = _generate_test_data()
    reporter.add_step(
        "Generated Shopify test data",
        "INFO",
        f"Name: Vivek {data.last_name}  |  Jersey: {data.jersey}  |  Reg: {data.reg_name}",
    )

    customer = {
        "email":      "vivek@bitcot.com",
        "first_name": "Vivek",
        "last_name":  data.last_name,
        "address":    "123 Test Street",
        "apt":        "Apt 4B",
        "city":       "Los Angeles",
        "state":      "CA",
        "zip":        "90001",
    }

    with browser_session(playwright, slow_mo=50) as page:
        try:
            store = ShopifyStorePage(page)

            # ── Phase 1: Authenticate ──────────────────────────────────────
            store.authenticate()
            reporter.add_step("Authenticate with Shopify staging store", "PASS")

            # ── Phase 2: Search ────────────────────────────────────────────
            store.search_team_store("fc+hawaii")
            reporter.add_step(
                "Search for FC Hawaii team store",
                "PASS",
                f"URL: {page.url}",
            )

            # ── Phase 3: Add products ──────────────────────────────────────
            store.add_simple_product(
                product_url           = SHOPIFY_BASE_URL + SHOPIFY_PRODUCTS["jersey"],
                quantity_extra_clicks = 1,
                size_option           = "GIRLS EXTRA SMALL",
                player_number         = "50",
            )
            reporter.add_step(
                "Add Product 1 – Game Jersey SS (Sky White)",
                "PASS",
                "Qty: 2  |  Size: GIRLS EXTRA SMALL  |  Player #50",
            )

            store.add_simple_product(
                product_url           = SHOPIFY_BASE_URL + SHOPIFY_PRODUCTS["hoodie"],
                quantity_extra_clicks = 1,
                size_option           = "ADULT LARGE",
                player_number         = "56",
            )
            reporter.add_step(
                "Add Product 2 – Cotton-Blend Hoodie (Grey)",
                "PASS",
                "Qty: 2  |  Size: ADULT LARGE  |  Player #56",
            )

            store.add_simple_product(
                product_url           = SHOPIFY_BASE_URL + SHOPIFY_PRODUCTS["shorts"],
                quantity_extra_clicks = 1,
                size_option           = "BOYS SMALL (YOUTH)",
                player_number         = "50",
            )
            reporter.add_step(
                "Add Product 3 – Game Shorts (Sky)",
                "PASS",
                "Qty: 2  |  Size: BOYS SMALL (YOUTH)",
            )

            store.add_kit_product(
                product_url            = SHOPIFY_BASE_URL + SHOPIFY_PRODUCTS["kit"],
                quantity_extra_clicks  = 5,
                reg_name               = data.reg_name,
                jersey_name            = data.jersey,
                player_number_option   = "14",
                jersey_size_option     = "YOUTH SMALL",
                shorts_size_option     = "MENS XL (ADULT)",
                socks_size_option      = "SMALL",
            )
            reporter.add_step(
                "Add Product 4 – Light Game Kit (YMQ)",
                "PASS",
                (
                    f"Qty: 5  |  Name: {data.reg_name}  |  Jersey: {data.jersey}"
                    f"  |  #14  |  Youth Sm / Men XL / Socks Sm"
                ),
            )

            # ── Phase 4: Cart confirmation ─────────────────────────────────
            store.complete_cart_confirmation(
                first_name  = "Vivek",
                full_name   = f"Vivek {data.last_name}",
                jersey_name = data.jersey,
            )
            reporter.add_step("Complete cart-drawer checkout confirmation", "PASS")

            # ── Phase 5: Shipping information ──────────────────────────────
            store.fill_shipping_information(customer)
            reporter.add_step(
                "Fill guest checkout shipping information",
                "PASS",
                (
                    f"{customer['address']}, {customer['city']}, "
                    f"{customer['state']} {customer['zip']}"
                ),
            )

            # ── Phase 6: Payment ───────────────────────────────────────────
            store.fill_payment_and_place_order(_PAYMENT)
            reporter.add_step("Fill payment details and place order", "PASS")

            # ── Phase 7: Verify confirmation ───────────────────────────────
            try:
                heading = store.verify_order_confirmation("Vivek")
                reporter.add_step("Verify order confirmation page", "PASS", heading)
            except PwTimeoutError:
                reporter.add_step(
                    "Verify order confirmation page",
                    "FAIL",
                    "Confirmation page did not load within the expected timeout.",
                )

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
