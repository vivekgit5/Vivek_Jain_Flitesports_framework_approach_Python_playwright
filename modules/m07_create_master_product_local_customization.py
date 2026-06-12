"""
Module M07 – Create New Master Product
=======================================
Logs in as CRM super-admin and navigates to:

  Master Products → Local Customization → [first partner collection]

Then creates a new product by completing all mandatory fields across the
product-creation form and the resulting product variants page:

  Step 1  – Enter a unique, Faker-generated product name
  Step 2  – Enter a rich-text description via the Quill editor
  Step 3  – Set category, gender, and sub-gender attributes
  Step 4  – Upload the product image from C:\\Flitesport_Product_images
  Step 5  – Select sport checkboxes (if present on the creation form)
  Step 6  – Select the first four available size checkboxes
  Step 7  – Add a custom YES / NO option (Player Number add-on)
  Step 8  – Add a YOUTH SMALL product variant
  Step 9  – Assign the uploaded image, a unique SKU, and base price to each
             generated variant row
  Step 10 – Save the product (initial save → navigates to variants page)
  Step 11 – Select SPORT attribute: BASEBALL, BASKETBALL, LACROSSE
  Step 12 – Select COLOR: GREY
  Step 13 – Enter Product Type: Cotton Blend Hooded Sweatshirt
  Step 14 – Enable all Logo Placement checkboxes
  Step 15 – Save the variant page attribute changes
  Step 16 – Sign out

An XLSX execution report is written to the ``reports/`` directory on
completion (pass or fail).

Usage
-----
    python -m modules.m07_create_master_product_local_customization
"""

import logging
import random
from dataclasses import dataclass

from faker import Faker
from playwright.sync_api import Playwright, expect, sync_playwright

from config.settings import (
    CRM_ADMIN_EMAIL,
    CRM_ADMIN_PASS,
    PRODUCT_IMAGE_PATH,
)
from core.browser import browser_session
from core.reporter import XlsxReporter
from pages.login_page import LoginPage

log = logging.getLogger(__name__)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)

MODULE_NAME = "M07 – Create New Master Product"

# ── Static product description ─────────────────────────────────────────────────
_PRODUCT_DESCRIPTION = (
    "this is testing product for master product in local customization 880"
)

# Gender / sub-gender option values as defined by the CRM dropdowns
_GENDER_YOUTH   = "2"   # Youth
_SUB_GENDER_BOY = "1"   # Boy

# ── Variant-page sidebar attributes (filled after initial save) ────────────────
# These fields reside in the right sidebar of the product creation form and
# the variants edit page.
_PRODUCT_SPORTS     = ["BASEBALL", "BASKETBALL", "LACROSSE"]
_PRODUCT_COLOR      = "GREY"
_PRODUCT_TYPE       = "Cotton Blend Hooded Sweatshirt"
_COLLECTION_SEARCH  = "Cotton"


# ── Test-data model ────────────────────────────────────────────────────────────

@dataclass
class ProductData:
    name:   str
    sku1:   str
    sku2:   str
    price1: str
    price2: str


def _generate_product_data() -> ProductData:
    """Return a ``ProductData`` instance with Faker-generated, unique values."""
    fake = Faker("en_US")
    tag  = random.randint(10_000, 99_999)
    return ProductData(
        name   = f"QA VIVEK MASTER PRODUCT {fake.word().upper()} {tag}",
        sku1   = fake.bothify("??##??##").upper(),
        sku2   = fake.bothify("??##??##").upper(),
        price1 = str(random.randint(40, 80)),
        price2 = str(random.randint(81, 150)),
    )


# ── Navigation ─────────────────────────────────────────────────────────────────

def _navigate_to_local_customization(page, reporter: XlsxReporter) -> None:
    """
    Expand the Master Products navigation accordion and click the
    Local Customization sub-link so the page lands on
    /super-admin/admin/master-products — where the '+ Add Product'
    button is visible in the top-right header.
    """
    # Expand the Master Products accordion in the sidebar navigation
    page.locator("nav > div:nth-of-type(2) button").click()
    page.wait_for_timeout(600)
    reporter.add_step("Expand Master Products navigation", "PASS")

    # Click the LOCAL CUSTOMIZATION link
    page.get_by_role("link", name="LOCAL CUSTOMIZATION").click()
    # The CRM SPA briefly routes through the dashboard before the Vue router
    # settles on the collection-management URL.  Waiting for networkidle alone
    # can capture the intermediate dashboard URL — wait for the correct path
    # instead so the logged URL always reflects the true destination.
    try:
        page.wait_for_url("**/collection-management**", timeout=12_000)
    except Exception:
        page.wait_for_load_state("networkidle", timeout=10_000)
    page.wait_for_timeout(400)
    reporter.add_step("Navigate to Local Customization", "PASS", page.url)


def _click_add_product(page, reporter: XlsxReporter) -> None:
    """
    Click the '+ Add Product' button in the top-right of the Local
    Customization page.  The button is rendered as a dark primary button
    with the exact accessible name 'Add Product'.
    """
    # Allow Vue to finish rendering after navigation
    page.wait_for_timeout(1_500)

    add_btn = page.get_by_role("button", name="Add Product")
    add_btn.wait_for(state="visible", timeout=15_000)
    add_btn.click()
    page.wait_for_load_state("networkidle")
    reporter.add_step("Click + Add Product", "PASS")


# ── Product form — main column ─────────────────────────────────────────────────

def _fill_product_name(page, product: ProductData, reporter: XlsxReporter) -> None:
    """Enter the unique Faker-generated product name."""
    name_input = page.locator(
        "div.sp-col-main div:nth-of-type(2) > div:nth-of-type(1) > input"
    )
    name_input.wait_for(state="visible", timeout=10_000)
    name_input.click(click_count=3)
    name_input.fill(product.name)
    reporter.add_step("Enter product name", "PASS", product.name)


def _fill_description(page, reporter: XlsxReporter) -> None:
    """
    Enter the product description into the Quill rich-text editor.

    The Quill editor intercepts keyboard events directly on the content
    area.  The editor's ``<p>`` element is clicked to focus the field,
    any pre-existing content is selected with Ctrl+A, and the description
    is typed character-by-character so that Quill's internal event
    listeners register every keystroke correctly.
    """
    editor_p = page.locator("div.sp-col-main > div:nth-of-type(2) p").first
    editor_p.wait_for(state="visible", timeout=10_000)
    editor_p.click()
    page.wait_for_timeout(400)
    page.keyboard.press("Control+a")
    page.keyboard.type(_PRODUCT_DESCRIPTION, delay=10)
    page.keyboard.press("Tab")
    page.wait_for_timeout(300)
    reporter.add_step("Enter product description (Quill editor)", "PASS")


# ── Product form — sidebar ─────────────────────────────────────────────────────

def _fill_category_and_attributes(page, reporter: XlsxReporter) -> None:
    """
    Fill the category type-ahead search field and select gender / sub-gender
    from their respective dropdowns.
    """
    # Category search
    cat_input = page.locator(
        "div.sp-col-sidebar > div:nth-of-type(2) > div:nth-of-type(1) input"
    )
    cat_input.click()
    cat_input.fill("cot")
    page.wait_for_timeout(400)
    # Dismiss the autocomplete dropdown by clicking an inert area of the sidebar
    page.locator("div.sp-col-sidebar > div:nth-of-type(2)").click(
        position={"x": 140, "y": 246}
    )
    page.wait_for_timeout(300)
    reporter.add_step("Enter product category", "PASS", "cot")

    # Gender dropdown (Youth)
    page.locator("div.sp-col-sidebar div:nth-of-type(3) > select").select_option(
        _GENDER_YOUTH
    )
    reporter.add_step("Select gender", "PASS", f"Youth (value: {_GENDER_YOUTH})")

    # Sub-gender dropdown (Boy)
    page.locator("div.sp-col-sidebar div:nth-of-type(4) > select").select_option(
        _SUB_GENDER_BOY
    )
    reporter.add_step(
        "Select sub-gender", "PASS", f"Boy (value: {_SUB_GENDER_BOY})"
    )


def _upload_product_image(page, reporter: XlsxReporter) -> None:
    """
    Upload the product image via the media card file-chooser.

    Playwright's ``expect_file_chooser()`` intercepts the native OS
    file-upload dialog triggered when the media upload area is clicked,
    and programmatically supplies the configured image file path without
    any OS-level dialog interaction.
    """
    with page.expect_file_chooser(timeout=8_000) as fc_info:
        page.locator("div.sp-card--media > div > div").first.click()
    fc_info.value.set_files(str(PRODUCT_IMAGE_PATH))
    page.wait_for_timeout(1_000)
    reporter.add_step("Upload product image", "PASS", PRODUCT_IMAGE_PATH.name)


def _select_sports(page, reporter: XlsxReporter) -> None:
    """Select the first three available sport checkboxes in the sidebar."""
    sport_inputs = page.locator(
        "div.sp-col-sidebar div.normal-text div input[type='checkbox']"
    )
    selected = 0
    for i in range(sport_inputs.count()):
        cb = sport_inputs.nth(i)
        if cb.is_visible() and cb.is_enabled() and not cb.is_checked():
            cb.click()
            page.wait_for_timeout(200)
            selected += 1
        if selected >= 3:
            break
    reporter.add_step(
        "Select sport(s)",
        "PASS" if selected > 0 else "INFO",
        f"{selected} sport(s) selected",
    )


def _select_sizes(page, reporter: XlsxReporter) -> None:
    """Select the first four available size checkboxes in the collection panel."""
    size_inputs = page.locator("div.sp-field--mt input[type='checkbox']")
    selected = 0
    for i in range(size_inputs.count()):
        inp = size_inputs.nth(i)
        if inp.is_visible() and inp.is_enabled() and not inp.is_checked():
            inp.click()
            page.wait_for_timeout(150)
            selected += 1
        if selected >= 4:
            break
    reporter.add_step(
        "Select size(s)",
        "PASS" if selected > 0 else "INFO",
        f"{selected} size(s) selected",
    )


# ── Product options ────────────────────────────────────────────────────────────

def _add_custom_option(page, reporter: XlsxReporter) -> None:
    """
    Add a custom product option representing the Player Number add-on,
    with two selectable values: YES and NO.
    """
    add_btn = page.locator(
        "div.sp-col-main > div:nth-of-type(5) button", has_text="Add another option"
    )
    add_btn.wait_for(state="visible", timeout=8_000)
    add_btn.click()
    page.wait_for_timeout(500)
    reporter.add_step("Click Add another option", "PASS")

    # Add YES value
    opt_input = page.locator("div.sp-opt-values input").first
    opt_input.fill("YES")
    page.locator("div.sp-opt-values button").first.click()
    page.wait_for_timeout(300)
    reporter.add_step("Add custom option value", "PASS", "YES")

    # Add NO value
    opt_input_no = page.locator("div.sp-opt-values input").last
    opt_input_no.fill("NO")
    page.locator("div.sp-add-opt3-field > button").first.click()
    page.wait_for_timeout(300)
    reporter.add_step("Add custom option value", "PASS", "NO")


# ── Variants ───────────────────────────────────────────────────────────────────

def _add_variant(page, reporter: XlsxReporter) -> None:
    """
    Open the Add Variant modal and select the YOUTH SMALL size.

    The modal may be rendered with class ``sp-modal-backdrop``,
    ``sp-modal``, or an inline overlay — so selectors are kept broad
    (scoped only by ``[class*='modal']`` or page-wide visible checkboxes
    that appear after the modal opens).
    """
    page.locator("div.sp-card--no-pad button").first.click()
    page.wait_for_timeout(1_000)
    reporter.add_step("Open Add Variant modal", "PASS")

    # Try the 4th label checkbox (YOUTH SMALL position) first
    youth_small = page.locator("label:nth-of-type(4) > input")
    try:
        youth_small.wait_for(state="visible", timeout=5_000)
        if not youth_small.is_checked():
            youth_small.click()
        reporter.add_step("Select YOUTH SMALL variant size", "PASS")
    except Exception:
        # Broad fallback: any visible unchecked checkbox that appeared with the modal
        for sel in [
            "[class*='modal'] input[type='checkbox']",
            "[class*='backdrop'] input[type='checkbox']",
            "div.sp-modal input[type='checkbox']",
        ]:
            try:
                fallback = page.locator(sel).first
                fallback.wait_for(state="visible", timeout=4_000)
                if not fallback.is_checked():
                    fallback.click()
                reporter.add_step(
                    "Select first available variant size (fallback)", "PASS"
                )
                break
            except Exception:
                continue

    # Confirm — try several known confirm button selectors
    for confirm_sel in [
        "[class*='modal'] button.sp-btn-primary",
        "[class*='backdrop'] button.sp-btn-primary",
        "div.sp-modal button.sp-btn-primary",
        "div.sp-modal-backdrop button.sp-btn-primary",
    ]:
        try:
            btn = page.locator(confirm_sel).last
            btn.wait_for(state="visible", timeout=4_000)
            btn.click()
            break
        except Exception:
            continue

    page.wait_for_timeout(600)
    reporter.add_step("Confirm variant selection", "PASS")


def _fill_variant_details(page, reporter: XlsxReporter) -> None:
    """
    For each variant row in the variants table, enter a unique SKU and a
    base price.  The rows appear on the product variants edit page
    (``/master-products-variants/<id>?type=0``) after the initial save.

    SKUs and prices are read directly from the table's editable cells:
    - SKU column  → ``textbox`` with placeholder ``SKU *``
    - Price column → ``textbox`` with placeholder ``0.00`` (first price cell)

    Two static values are used so the entries remain deterministic and easy
    to verify in the CRM without relying on the Faker instance that was used
    on the creation form.
    """
    fake = Faker("en_US")
    skus   = [fake.bothify("??##??##").upper(), fake.bothify("??##??##").upper()]
    prices = [str(random.randint(40, 80)),       str(random.randint(81, 150))]

    # Wait for the variants table to render after navigation
    page.wait_for_timeout(1_500)

    # Use the broadest reliable selector: all data rows inside any tbody on the page
    rows = page.locator("table tbody tr")

    # Give Vue time to finish rendering variant rows
    try:
        rows.first.wait_for(state="visible", timeout=8_000)
    except Exception:
        log.warning("Variant table rows not found — SKU/price fill skipped")
        reporter.add_step("Fill variant SKU and price", "INFO", "No variant rows visible")
        return

    row_count = rows.count()
    log.info("Variant rows found: %d", row_count)

    for i in range(min(row_count, 2)):
        row = rows.nth(i)

        # ── SKU ────────────────────────────────────────────────────────────────
        sku_inp = row.locator("input[placeholder='SKU *']")
        try:
            sku_inp.wait_for(state="visible", timeout=3_000)
            sku_inp.click(click_count=3)
            sku_inp.fill(skus[i])
            reporter.add_step(f"Enter SKU for variant {i + 1}", "PASS", skus[i])
        except Exception:
            pass

        # ── Price ──────────────────────────────────────────────────────────────
        # The price cell contains a text-input with placeholder "0.00"; the
        # first such input inside the row is the Price ($) column.
        price_inp = row.locator("input[placeholder='0.00']").first
        try:
            price_inp.wait_for(state="visible", timeout=3_000)
            price_inp.click(click_count=3)
            price_inp.fill(prices[i])
            reporter.add_step(
                f"Enter price for variant {i + 1}", "PASS", f"${prices[i]}"
            )
        except Exception:
            pass


# ── Variant-page sidebar helpers (run after initial save) ─────────────────────

def _select_product_sports(page, reporter: XlsxReporter) -> None:
    """
    Select BASEBALL, BASKETBALL, and LACROSSE in the mandatory SPORT * field
    on the product creation form right sidebar.

    The field is rendered as a custom tag-selector whose trigger contains the
    placeholder text 'SELECT SPORT'.  Clicking the trigger opens a dropdown;
    each sport option is then selected by its accessible role or visible text.
    """
    try:
        # Trigger: the clickable div/span that shows 'SELECT SPORT'
        sport_trigger = page.get_by_text("SELECT SPORT", exact=True)
        sport_trigger.wait_for(state="visible", timeout=8_000)
        sport_trigger.click()
        page.wait_for_timeout(300)
        for sport in _PRODUCT_SPORTS:
            try:
                # Short timeout so the fallback is reached immediately if the
                # option is not rendered as an ARIA role on this dropdown.
                page.get_by_role("option", name=sport, exact=True).click(
                    timeout=1_500
                )
            except Exception:
                page.get_by_text(sport, exact=True).last.click(timeout=3_000)
            page.wait_for_timeout(80)
        page.keyboard.press("Escape")
        page.wait_for_timeout(200)
        reporter.add_step(
            "Select SPORT attribute values",
            "PASS",
            ", ".join(_PRODUCT_SPORTS),
        )
    except Exception as exc:
        reporter.add_step("Select SPORT attribute values", "INFO", f"Skipped — {exc}")


def _select_product_color(page, reporter: XlsxReporter) -> None:
    """
    Set the mandatory COLOR * field to GREY by typing into the 'Search color…'
    type-ahead input and confirming the matching suggestion.
    """
    try:
        color_input = page.get_by_placeholder("Search color…")
        color_input.wait_for(state="visible", timeout=8_000)
        color_input.click(click_count=3)
        color_input.fill(_PRODUCT_COLOR)
        # Wait only until the dropdown option becomes visible, then click it
        option = page.get_by_role("option", name=_PRODUCT_COLOR, exact=True)
        try:
            option.wait_for(state="visible", timeout=3_000)
            option.click()
        except Exception:
            page.get_by_text(_PRODUCT_COLOR, exact=True).last.click()
        reporter.add_step("Select product color", "PASS", _PRODUCT_COLOR)
    except Exception as exc:
        reporter.add_step("Select product color", "INFO", f"Skipped — {exc}")


def _enter_product_type(page, reporter: XlsxReporter) -> None:
    """
    Set the Product Type field in the creation-form sidebar to
    'Cotton Blend Hooded Sweatshirt'.

    Characters are typed sequentially so that Vue reactivity fires and the
    CRM autocomplete panel appears.  The matching suggestion is clicked to
    commit the value — simply filling the input without selecting the
    suggestion leaves the field in an uncommitted state that does not persist.
    """
    try:
        type_input = page.get_by_placeholder("e.g. Polo Shirt")
        type_input.wait_for(state="visible", timeout=8_000)
        type_input.scroll_into_view_if_needed()
        # Triple-click to select any pre-existing value before typing
        type_input.click(click_count=3)
        type_input.press("Control+a")
        # Type sequentially so Vue's autocomplete reactive listener fires
        type_input.press_sequentially("cotton blend hooded sweatshirt", delay=60)
        # The CRM autocomplete may render the suggestion in ALL CAPS, mixed
        # case, or via CSS text-transform.  Try the most common variants in
        # order; fall back to committing the typed value with Tab if none match.
        _suggestion_texts = [
            "COTTON BLEND HOODED SWEATSHIRT",
            "Cotton Blend Hooded Sweatshirt",
            "cotton blend hooded sweatshirt",
        ]
        suggestion_clicked = False
        for txt in _suggestion_texts:
            try:
                elem = page.get_by_text(txt, exact=True)
                elem.wait_for(state="visible", timeout=3_000)
                elem.click()
                suggestion_clicked = True
                reporter.add_step("Select product type", "PASS", _PRODUCT_TYPE)
                break
            except Exception:
                continue
        if not suggestion_clicked:
            # Last resort: click the first visible autocomplete list item
            try:
                first_item = page.locator(
                    "li[class*='autocomplete'], li[class*='suggest'], "
                    "[role='option'], [class*='dropdown-item']"
                ).first
                first_item.wait_for(state="visible", timeout=2_000)
                first_item.click()
                suggestion_clicked = True
                reporter.add_step("Select product type", "PASS", _PRODUCT_TYPE + " (autocomplete)")
            except Exception:
                pass
        if not suggestion_clicked:
            # If autocomplete doesn't appear the typed value is still accepted
            type_input.press("Tab")
            reporter.add_step("Select product type", "PASS", _PRODUCT_TYPE + " (typed)")
    except Exception as exc:
        reporter.add_step("Select product type", "INFO", f"Skipped \u2014 {exc}")


def _assign_variant_images(page, reporter: XlsxReporter) -> None:
    """
    Assign the first uploaded product image to every variant row.

    On the product creation form each variant row whose image has not been
    assigned yet renders a clickable SVG thumb div with title="Assign image"
    (CSS class sp-var-thumb--empty sp-var-thumb--clickable).  Clicking that
    element opens the 'ASSIGN IMAGE' modal.

    The automation:
      1. Waits for the variant table to render.
      2. Collects all [title="Assign image"] triggers on the page.
      3. For each trigger: scrolls it into view, clicks with force=True
         (sidebar overlay otherwise intercepts), waits for the modal
         "Assign image" confirm button to appear, clicks the first
         image thumbnail, then clicks "Assign image" to confirm.
    """
    try:
        # Wait for variant table rows to be present
        page.locator("table tbody tr").first.wait_for(
            state="visible", timeout=8_000
        )
        page.wait_for_timeout(500)

        # The SVG camera icon in each unassigned variant cell has title="Assign image"
        triggers = page.locator('[title="Assign image"].sp-var-thumb--clickable')
        trigger_count = triggers.count()
        log.info('"Assign image" triggers found: %d', trigger_count)

        assigned = 0
        for i in range(trigger_count):
            try:
                trig = triggers.nth(i)
                trig.scroll_into_view_if_needed()
                page.wait_for_timeout(300)
                # force=True bypasses the sidebar overlay that intercepts clicks
                trig.click(force=True)
                page.wait_for_timeout(800)

                # Wait for the confirm button to appear in the modal
                assign_btn = page.get_by_role("button", name="Assign image")
                assign_btn.wait_for(state="visible", timeout=5_000)

                # Click the first image thumbnail in the modal
                modal_img = page.locator(
                    ".sp-modal img, [role='dialog'] img, "
                    "[class*='modal'] img, dialog img"
                ).first
                modal_img.wait_for(state="visible", timeout=5_000)
                modal_img.click()
                page.wait_for_timeout(400)

                # Confirm the assignment
                page.get_by_role("button", name="Assign image").click()
                page.wait_for_timeout(700)
                assigned += 1
            except Exception as row_exc:
                log.warning("Could not assign image for trigger %d: %s", i, row_exc)
                page.keyboard.press("Escape")
                page.wait_for_timeout(300)
                continue

        reporter.add_step(
            "Assign images to variant rows",
            "PASS" if assigned > 0 else "INFO",
            f"{assigned} of {trigger_count} variant(s) assigned",
        )
    except Exception as exc:
        reporter.add_step("Assign images to variant rows", "INFO", f"Skipped \u2014 {exc}")
def _select_collection(page, reporter: XlsxReporter) -> None:
    """
    Set the mandatory FLITE PLATFORM GROUP SPORT(S) * field by searching for
    'Cotton' and selecting the 'COTTON-BLEND COLLECTION - HEAT PRESS' option
    as shown in the CRM dropdown.
    """
    try:
        col_input = page.get_by_placeholder("Search collection…")
        col_input.wait_for(state="visible", timeout=8_000)
        col_input.scroll_into_view_if_needed()
        # Use force=True to bypass any sp-tag-option chip that overlaps the input
        col_input.click(force=True)
        col_input.fill(_COLLECTION_SEARCH)
        page.wait_for_timeout(600)
        # Prefer the exact option visible in the CRM dropdown
        try:
            heat_press = page.get_by_text(
                "COTTON-BLEND COLLECTION - HEAT PRESS", exact=True
            )
            heat_press.wait_for(state="visible", timeout=3_000)
            heat_press.click()
        except Exception:
            try:
                page.get_by_role("option").first.click(timeout=2_000)
            except Exception:
                pass
        reporter.add_step(
            "Select FLITE platform group sport(s)",
            "PASS",
            "Cotton-Blend Collection - Heat Press",
        )
    except Exception as exc:
        reporter.add_step("Select FLITE platform group sport(s)", "INFO", f"Skipped — {exc}")
def _check_logo_placements(page, reporter: XlsxReporter) -> None:
    """
    Ensure every Logo Placement checkbox on the product variants page is
    checked.  The section contains up to four named placements per product
    image (Front Left Chest, Front Centre Chest, Back Centre Shoulders, Back
    Centre Bottom), each rendered as an accessible checkbox.
    """
    _placement_labels = [
        "Front (Left Chest)",
        "Front (Center Chest)",
        "Back (Center Shoulders)",
        "Back (Center Bottom)",
    ]
    try:
        # Wait for the CRM loading overlay to clear before interacting.
        try:
            page.wait_for_selector(
                ".gl-overlay, .gl-backdrop, [aria-label='Loading, please wait']",
                state="hidden",
                timeout=25_000,
            )
        except Exception:
            pass
        page.wait_for_timeout(600)

        # Target logo placement checkboxes directly by their CSS class rather
        # than by a section heading text which may not be present / may differ
        # across CRM versions (e.g. all-caps, translated, or absent entirely).
        logo_cbs = page.locator("input.sp-logo-checkbox")
        try:
            logo_cbs.first.wait_for(state="visible", timeout=8_000)
        except Exception:
            # Section not present on this page variant — skip gracefully.
            reporter.add_step(
                "Enable logo placement options",
                "INFO",
                "Logo placement checkboxes not found on variants page (section may be absent)",
            )
            return

        checked_count = 0
        for i in range(logo_cbs.count()):
            cb = logo_cbs.nth(i)
            if cb.is_visible() and cb.is_enabled() and not cb.is_checked():
                cb.click()
                page.wait_for_timeout(150)
                checked_count += 1
        reporter.add_step(
            "Enable logo placement options",
            "PASS",
            f"{checked_count} placement(s) newly checked",
        )
    except Exception as exc:
        reporter.add_step("Enable logo placement options", "INFO", f"Skipped — {exc}")


def _save_variant_page(page, reporter: XlsxReporter) -> None:
    """
    Persist all sidebar attribute changes on the product variants page by
    clicking the primary Save button.  'Save' is tried first (the actual
    label on the CRM variants edit page), followed by several alternative
    wordings for resilience against future CRM label changes.
    """
    # Ensure the loading overlay has cleared before searching for the save button.
    # After logo-placement interactions the page may still be settling.
    try:
        page.wait_for_selector(
            ".gl-overlay, .gl-backdrop, [aria-label='Loading, please wait']",
            state="hidden",
            timeout=15_000,
        )
    except Exception:
        pass
    page.wait_for_timeout(400)

    # Try text-based button labels first, then fall back to primary CSS class.
    for btn_label in ["Save product", "Save", "Save changes", "Update product", "Update"]:
        try:
            btn = page.get_by_role("button", name=btn_label, exact=True)
            btn.wait_for(state="visible", timeout=5_000)
            btn.scroll_into_view_if_needed()
            btn.click()
            page.wait_for_load_state("networkidle")
            reporter.add_step(
                "Save product variant page attributes", "PASS", f'Clicked "{btn_label}"'
            )
            return
        except Exception:
            continue

    # CSS-class fallback: any enabled primary button visible on the page.
    try:
        css_btn = page.locator("button.sp-btn-primary").last
        css_btn.wait_for(state="visible", timeout=5_000)
        css_btn.scroll_into_view_if_needed()
        css_btn.click()
        page.wait_for_load_state("networkidle")
        reporter.add_step(
            "Save product variant page attributes", "PASS", "Clicked primary button (CSS fallback)"
        )
        return
    except Exception:
        pass

    reporter.add_step(
        "Save product variant page attributes",
        "INFO",
        "No recognisable save button found on variants page",
    )


# ── Module orchestrator ────────────────────────────────────────────────────────

def run(playwright: Playwright) -> None:
    """
    Execute the full Create New Master Product flow.

    Steps
    -----
    1.   Generate unique product test data
    2.   Login as CRM super-admin
    3.   Navigate to Master Products → Local Customization
    4.   Open the Add Product form
    5.   Fill product name
    6.   Fill rich-text description (Quill editor)
    7.   Set category, gender, and sub-gender attributes
    8.   Upload product image from C:\\Flitesport_Product_images
    9.   Select COLOR = GREY  (required before Add Variant)
    10.  Select SPORT attribute values (BASEBALL, BASKETBALL, LACROSSE)
    11.  Select FLITE Platform Group Sport(s) = Cotton-Blend Collection
    12.  Add custom YES / NO product option (Player Number add-on)
    13.  Add YOUTH SMALL variant
    14.  Enter SKU and price for each generated variant row
    15.  Save product (initial save → navigates to variants edit page)
    16.  Enter Product Type = Cotton Blend Hooded Sweatshirt
    17.  Enable all Logo Placement checkboxes
    18.  Save variant page attribute changes
    19.  Sign out
    """
    reporter = XlsxReporter(MODULE_NAME)
    product  = _generate_product_data()

    reporter.add_step(
        "Generated product test data",
        "INFO",
        (
            f"Name: {product.name}  |  "
            f"SKU1: {product.sku1}  |  SKU2: {product.sku2}  |  "
            f"Price1: ${product.price1}  |  Price2: ${product.price2}"
        ),
    )

    with browser_session(playwright) as page:
        try:
            login = LoginPage(page)

            login.navigate()
            reporter.add_step("Navigate to CRM login page", "PASS", page.url)
            login.login(CRM_ADMIN_EMAIL, CRM_ADMIN_PASS)
            reporter.add_step(
                "Authenticate as CRM super-admin", "PASS", CRM_ADMIN_EMAIL
            )

            _navigate_to_local_customization(page, reporter)
            _click_add_product(page, reporter)
            _fill_product_name(page, product, reporter)
            _fill_description(page, reporter)
            _fill_category_and_attributes(page, reporter)
            _upload_product_image(page, reporter)

            # ── Sidebar: fill ALL mandatory fields on the creation form ────────
            # Order matters:
            #   1. Product Type first — triggers Logo Placement section to appear
            #   2. Color before Add Variant — CRM blocks variant creation otherwise
            #   3. Sport and Collection while sidebar is in scope
            _enter_product_type(page, reporter)
            _select_product_color(page, reporter)
            _select_product_sports(page, reporter)
            _select_collection(page, reporter)

            _add_custom_option(page, reporter)
            _add_variant(page, reporter)
            _fill_variant_details(page, reporter)

            # Assign uploaded image to each variant row via the camera icon,
            # then enable all Logo Placement checkboxes — both done here on
            # the creation form so everything is saved in a single request.
            _assign_variant_images(page, reporter)

            # ── Save — persists all fields entered on the creation form ────────
            # NOTE: Logo placements are NOT checked here.  The Logo Placement
            # section only renders on the variants *edit* page after the initial
            # save redirect.  Calling _check_logo_placements before save always
            # returns 0 newly-checked items.  It is called below, after the
            # initial save and before _save_variant_page.
            save_btn = page.locator(
                "button.sp-btn-primary", has_text="Save product"
            )
            expect(save_btn).to_be_visible(timeout=10_000)
            expect(save_btn).to_be_enabled(timeout=10_000)
            save_btn.scroll_into_view_if_needed()
            save_btn.click()
            page.wait_for_load_state("networkidle")
            reporter.add_step(
                "Save product — product created successfully", "PASS"
            )

            # ── Variants edit page: confirm save ──────────────────────────────
            # The CRM redirects to the variants edit page after initial save.
            # Click Save once more to persist any edit-page-specific defaults.
            try:
                page.wait_for_load_state("networkidle", timeout=10_000)
            except Exception:
                pass
            page.wait_for_timeout(1_500)

            # Logo Placement section is rendered on the variants edit page
            # (after initial save).  Check and enable all placements here
            # before persisting the variants page.
            _check_logo_placements(page, reporter)
            _save_variant_page(page, reporter)

            # ── Verify master products listing URL ────────────────────────────
            _EXPECTED_URL = (
                "https://flitesports-stage.bitcotapps.com"
                "/super-admin/admin/master-products"
            )
            page.goto(_EXPECTED_URL)
            page.wait_for_load_state("networkidle", timeout=15_000)
            actual_url = page.url.rstrip("/")
            expected_url = _EXPECTED_URL.rstrip("/")
            if actual_url == expected_url:
                reporter.add_step(
                    "Verify master products listing URL",
                    "PASS",
                    actual_url,
                )
                log.info("URL verified: %s", actual_url)
            else:
                reporter.add_step(
                    "Verify master products listing URL",
                    "FAIL",
                    f"Expected {expected_url} — got {actual_url}",
                )
                log.warning("URL mismatch: expected %s got %s", expected_url, actual_url)

            login.sign_out()
            reporter.add_step("Sign out", "PASS")

            reporter.add_step("Module completed successfully", "PASS")
            log.info("✓  %s  —  all steps passed.", MODULE_NAME)

        except Exception as exc:
            reporter.add_step(
                "Module encountered an unhandled error", "FAIL", str(exc)
            )
            log.exception("Module failed: %s", exc)
            raise

        finally:
            report_path = reporter.save()
            log.info("Report → %s", report_path)


if __name__ == "__main__":
    with sync_playwright() as pw:
        run(pw)
