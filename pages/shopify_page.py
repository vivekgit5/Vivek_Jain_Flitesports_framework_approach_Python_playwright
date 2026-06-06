"""
Page object for the Flitesports Shopify staging storefront.

Covers store password authentication, product page interactions, cart
drawer management, and the guest checkout flow (shipping + payment).
"""

import logging

from playwright.sync_api import Page, TimeoutError as PwTimeoutError

from config.settings import SHOPIFY_BASE_URL, SHOPIFY_STORE_PASSWORD, DEFAULT_TIMEOUT

log = logging.getLogger(__name__)


class ShopifyStorePage:
    """
    Encapsulates all interactions with the Flitesports Shopify staging store.

    Public methods map 1-to-1 with the logical phases of the purchase flow:

    Phase 1 – ``authenticate()``
    Phase 2 – ``search_team_store(query)``
    Phase 3 – ``add_product_to_cart(url, options)``
    Phase 4 – ``complete_cart_confirmation(player_data)``
    Phase 5 – ``fill_shipping_information(customer)``
    Phase 6 – ``fill_payment_and_place_order(payment)``
    Phase 7 – ``verify_order_confirmation(first_name)``
    """

    def __init__(self, page: Page) -> None:
        self._page = page

    # ── Internal helpers ───────────────────────────────────────────────────────

    def _wait_for_cloudflare(self, max_retries: int = 12) -> None:
        """
        Poll for the presence of a Cloudflare Turnstile challenge and attempt
        to resolve it automatically.  Gives up gracefully after *max_retries*
        attempts to avoid blocking the test indefinitely.
        """
        try:
            for _ in range(max_retries):
                if not self._page.locator("body[class*='challenge']").is_visible(timeout=1_500):
                    return

                try:
                    checkbox = self._page.frame_locator(
                        "iframe[src*='challenges.cloudflare.com']"
                    ).locator("[id*='challenge-stage']").first
                    if checkbox.is_visible(timeout=2_000):
                        checkbox.click(timeout=3_000)
                        log.info("Clicked Cloudflare Turnstile checkbox — waiting for verification …")
                        self._page.wait_for_timeout(4_000)
                        continue
                except Exception:
                    pass

                self._page.wait_for_timeout(5_000)
        except Exception:
            pass

    def _close_cart_drawer(self) -> None:
        """Close the slide-out cart drawer when it is visibly open."""
        try:
            drawer = self._page.locator("#CartDrawer")
            if drawer.is_visible(timeout=3_000):
                close_btn = drawer.locator("div.drawer__fixed-header button").first
                if close_btn.is_visible(timeout=2_000):
                    close_btn.click(timeout=5_000)
                    self._page.wait_for_timeout(500)
        except Exception:
            pass

    # ── Phase 1 ────────────────────────────────────────────────────────────────

    def authenticate(self) -> None:
        """
        Navigate to the password-protected staging store and submit the
        store access password.
        """
        log.info("Authenticating with the Shopify staging store …")
        self._page.goto(f"{SHOPIFY_BASE_URL}/password")
        self._page.wait_for_load_state("domcontentloaded")
        self._page.locator("#password").fill(SHOPIFY_STORE_PASSWORD)
        self._page.locator("button[type='submit']").click()
        self._page.wait_for_load_state("load")
        log.info("Store authentication successful.")

    # ── Phase 2 ────────────────────────────────────────────────────────────────

    def search_team_store(self, query: str) -> None:
        """
        Navigate directly to the store search results URL for *query*.

        Args:
            query: URL-encoded search query (e.g. ``"fc+hawaii"``).
        """
        log.info("Searching the store for: %s", query)
        self._page.goto(f"{SHOPIFY_BASE_URL}/search?q={query}", wait_until="load")

    # ── Phase 3 ────────────────────────────────────────────────────────────────

    def add_simple_product(
        self,
        product_url: str,
        quantity_extra_clicks: int,
        size_option: str,
        player_number: str,
    ) -> None:
        """
        Navigate to a simple (non-customised) product page, set the quantity,
        size, and player number, then add to cart.

        Args:
            product_url:           Full URL of the product page.
            quantity_extra_clicks: Number of additional clicks on the "+" button
                                   (starting quantity is 1).
            size_option:           Exact text of the size option to select.
            player_number:         Player number to select from the dropdown.
        """
        log.info("Adding simple product: %s", product_url)
        self._page.goto(product_url)
        self._page.wait_for_load_state("load")
        self._wait_for_cloudflare()
        self._close_cart_drawer()

        section = self._page.locator("[id^='ProductSection-']").first
        section.wait_for(timeout=DEFAULT_TIMEOUT)

        for _ in range(quantity_extra_clicks):
            section.locator("button.js-qty__adjust--plus").click()
            self._page.wait_for_timeout(300)

        section.locator("select[id*='-option-0']").select_option(size_option)
        self._page.wait_for_timeout(400)

        section.locator("#player-number").select_option(player_number)
        self._page.wait_for_timeout(300)

        section.locator("div.payment-buttons button, button[name='add']").first.click()
        self._page.wait_for_timeout(1_500)
        self._close_cart_drawer()
        log.info("Simple product added to cart.")

    def add_kit_product(
        self,
        product_url: str,
        quantity_extra_clicks: int,
        reg_name: str,
        jersey_name: str,
        player_number_option: str,
        jersey_size_option: str,
        shorts_size_option: str,
        socks_size_option: str,
    ) -> None:
        """
        Navigate to a YMQ-customised kit product page, fill all custom option
        fields, and add the item to the cart.

        Args:
            product_url:            Full URL of the kit product page.
            quantity_extra_clicks:  Extra "+" button clicks (starting from 1).
            reg_name:               Registered player name (≤ 30 chars).
            jersey_name:            Jersey print name (≤ 14 chars).
            player_number_option:   Player number value for the dropdown.
            jersey_size_option:     Jersey size value for the dropdown.
            shorts_size_option:     Shorts size value for the dropdown.
            socks_size_option:      Socks size value for the dropdown.
        """
        log.info("Adding kit product (YMQ): %s", product_url)
        self._page.goto(product_url)
        self._page.wait_for_load_state("load")
        self._wait_for_cloudflare()
        self._page.wait_for_timeout(2_000)
        self._close_cart_drawer()

        section = self._page.locator("[id^='ProductSection-']").first
        section.wait_for(timeout=DEFAULT_TIMEOUT)

        plus_btn = section.locator("button.js-qty__adjust--plus")
        plus_btn.wait_for(timeout=DEFAULT_TIMEOUT)
        plus_btn.dblclick()
        self._page.wait_for_timeout(200)
        for _ in range(quantity_extra_clicks - 2):
            plus_btn.click()
            self._page.wait_for_timeout(200)

        ymq_inputs  = self._page.locator("[id^='ymq-option-value-']")
        ymq_selects = self._page.locator("[id^='ymq-option-select-']")

        ymq_inputs.nth(0).click()
        ymq_inputs.nth(0).fill(reg_name)
        self._page.wait_for_timeout(200)

        ymq_inputs.nth(1).click()
        ymq_inputs.nth(1).fill(jersey_name)
        self._page.wait_for_timeout(200)

        ymq_selects.nth(0).select_option(player_number_option)
        self._page.wait_for_timeout(300)
        ymq_selects.nth(1).select_option(jersey_size_option)
        self._page.wait_for_timeout(300)
        ymq_selects.nth(2).select_option(shorts_size_option)
        self._page.wait_for_timeout(300)
        ymq_selects.nth(3).select_option(socks_size_option)
        self._page.wait_for_timeout(300)

        self._page.locator("div.payment-buttons button, button[name='add']").first.click()
        self._page.wait_for_timeout(2_000)
        log.info("Kit product added to cart.")

    # ── Phase 4 ────────────────────────────────────────────────────────────────

    def complete_cart_confirmation(
        self,
        first_name: str,
        full_name: str,
        jersey_name: str,
    ) -> None:
        """
        Proceed through the cart-drawer checkout confirmation flow:

        1. Click Checkout in the cart drawer footer.
        2. Check the registered-player checkbox.
        3. Fill in the player-details modal.
        4. Confirm the order details and navigate to Shopify checkout.

        Args:
            first_name:  Customer first name (used in modal field 1).
            full_name:   Customer full name (used in modal field 2).
            jersey_name: Jersey print name (used in modal field 3).
        """
        log.info("Proceeding through cart confirmation modal …")
        self._page.locator("div.drawer__footer span", has_text="Check out").click()
        self._page.wait_for_load_state("load", timeout=180_000)
        self._page.wait_for_timeout(2_000)

        self._page.locator(
            "#checkout-customization-summary label > input"
        ).check(timeout=180_000)
        self._page.wait_for_timeout(500)

        modal = "#checkout-customization-modal"
        self._page.locator(f"{modal} div:nth-of-type(1) > input").fill(first_name)
        self._page.locator(f"{modal} div:nth-of-type(2) > input").fill(full_name)
        self._page.locator(f"{modal} div:nth-of-type(3) > input").fill(jersey_name)

        confirmed_cb = self._page.locator("#checkout-customization-confirmed")
        if not confirmed_cb.is_checked():
            confirmed_cb.click()
        self._page.wait_for_timeout(500)
        self._page.wait_for_function(
            "() => !document.querySelector('#confirm-checkout').disabled",
            timeout=10_000,
        )
        self._page.locator("#confirm-checkout").click()
        self._page.wait_for_load_state("load")
        log.info("Cart confirmation complete — on Shopify checkout.")

    # ── Phase 5 ────────────────────────────────────────────────────────────────

    def fill_shipping_information(self, customer: dict) -> None:
        """
        Fill the guest checkout shipping form.

        Args:
            customer: Dict with keys ``email``, ``first_name``, ``last_name``,
                      ``address``, ``apt``, ``city``, ``state``, ``zip``.
        """
        log.info("Filling guest checkout shipping information …")
        self._page.wait_for_load_state("load")
        self._page.wait_for_timeout(1_500)

        self._page.locator(
            "#email, [autocomplete='email'], [name='email']"
        ).first.fill(customer["email"])
        self._page.wait_for_timeout(400)

        self._page.locator(
            "[autocomplete='given-name'], [name='firstName']"
        ).first.fill(customer["first_name"])
        self._page.locator(
            "[autocomplete='family-name'], [name='lastName']"
        ).first.fill(customer["last_name"])

        self._page.locator(
            "[autocomplete='address-line1'], [name='address1'], #shipping-address1"
        ).first.fill(customer["address"])
        self._page.wait_for_timeout(300)

        self._page.locator(
            "[autocomplete='address-line2'], [name='address2']"
        ).first.fill(customer.get("apt", ""))

        self._page.locator(
            "[autocomplete='address-level2'], [name='city']"
        ).first.fill(customer["city"])

        self._page.locator(
            "[autocomplete='address-level1'], select[name='zone'], select[name='province']"
        ).first.select_option(customer["state"])

        self._page.locator(
            "[autocomplete='postal-code'], [name='zip'], [name='postalCode']"
        ).first.fill(customer["zip"])

        log.info("Shipping information filled.")

    # ── Phase 6 ────────────────────────────────────────────────────────────────

    def fill_payment_and_place_order(self, payment: dict) -> None:
        """
        Fill the credit-card fields (Shopify PCI iframes) and submit the order.

        Args:
            payment: Dict with keys ``card_number``, ``expiry``, ``cvv``.
        """
        log.info("Filling payment details …")

        card_frame = self._page.frame_locator(
            "iframe[title*='Card number'], iframe[id^='card-fields-number-']"
        ).first
        card_frame.locator("#number").fill(payment["card_number"])

        expiry_frame = self._page.frame_locator(
            "iframe[title*='Expiry'], iframe[title*='expiry'], iframe[src*='expiry-ltr.html']"
        ).first
        expiry_frame.locator("#expiry").fill(payment["expiry"])

        cvv_frame = self._page.frame_locator(
            "iframe[title*='Security code'], iframe[title*='CVV'], "
            "iframe[src*='verification_value-ltr.html']"
        ).first
        cvv_frame.locator("#verification_value").fill(payment["cvv"])

        self._page.locator("#checkout-pay-button").click()
        self._page.wait_for_load_state("networkidle", timeout=30_000)
        log.info("Order submitted.")

    # ── Phase 7 ────────────────────────────────────────────────────────────────

    def verify_order_confirmation(self, first_name: str) -> str:
        """
        Verify the order confirmation (thank-you) page is displayed.

        Returns:
            The heading text when found, or raises ``PwTimeoutError`` on failure.
        """
        log.info("Verifying order confirmation page …")
        heading = self._page.locator(
            "header h2", has_text=f"Thank you, {first_name}!"
        )
        heading.wait_for(timeout=20_000)
        heading_text = (heading.text_content() or "").strip()

        self._page.locator("header p", has_text="ORDER CONFIRMATION").wait_for(timeout=5_000)
        log.info("Order confirmation verified: %s", heading_text)
        return heading_text
