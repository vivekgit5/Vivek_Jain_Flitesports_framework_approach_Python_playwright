"""
Central configuration for the Flitesports automation framework.

All environment-specific constants live here.  Update this file when
credentials, URLs, or test-data defaults need to change — no module
code should need to be edited for environment switches.
"""

from pathlib import Path

# ── Directory layout ───────────────────────────────────────────────────────────
ROOT_DIR      = Path(__file__).resolve().parent.parent
REPORTS_DIR   = ROOT_DIR / "reports"
DATA_DIR      = ROOT_DIR / "data"
DOWNLOADS_DIR = ROOT_DIR / "downloads"

# Ensure output directories exist at import time
REPORTS_DIR.mkdir(exist_ok=True)
DOWNLOADS_DIR.mkdir(exist_ok=True)

# ── CRM application (Flitesports staging) ─────────────────────────────────────
CRM_BASE_URL    = "https://flitesports-stage.bitcotapps.com"
CRM_ADMIN_EMAIL = "admin@bitcot.com"
CRM_ADMIN_PASS  = "Bitcot@123"

# ── Shopify staging store ──────────────────────────────────────────────────────
SHOPIFY_BASE_URL       = "https://flite-sports-staging.myshopify.com"
SHOPIFY_STORE_PASSWORD = "123456"

# Product slugs used in the purchase-products module
SHOPIFY_PRODUCTS = {
    "jersey": "/products/fc-hawaii-game-jersey-ss-sky-white",
    "hoodie": "/products/fc-hawaii-cotton-blend-hoodie-grey",
    "shorts": "/products/fc-hawaii-game-shorts-sky",
    "kit":    "/products/fc-hawaii-light-game-kit",
}

# ── Falcon staging ─────────────────────────────────────────────────────────────
FALCON_APP_LAUNCHER_URL  = "https://falcon-stage.flitesports.com/apps/"
FALCON_ORDER_RANGE       = "1050-1070"
FALCON_EXPORT_CHANNEL    = "all"
FALCON_PRODUCTION_ORDER  = "Falcon_Stage_Comparison_Test"

# ── Browser defaults ───────────────────────────────────────────────────────────
VIEWPORT         = {"width": 1520, "height": 695}
DEFAULT_TIMEOUT  = 15_000   # milliseconds

# ── CRM role-filter values (Users page dropdown) ──────────────────────────────
ROLE_FILTER_ADMIN     = "6"
ROLE_FILTER_SALES_REP = "4"
ROLE_FILTER_PARTNER   = "3"

# ── Static test assets — Partner logos ────────────────────────────────────────
# All four official Flitesports logo variants are stored under data/logos/.
# PARTNER_LOGO_PATH is the default logo uploaded to the Partner Program Identity
# form in both the creation wizard (M03) and the update wizard (M04).
# The CRM blocks form submission if the primary logo field is left empty.
_LOGOS_DIR             = DATA_DIR / "logos"
LOGO_BLACK             = _LOGOS_DIR / "flite_logo_black.png"
LOGO_GRAY              = _LOGOS_DIR / "flite_logo_gray.png"
LOGO_RED               = _LOGOS_DIR / "flite_logo_red.png"
LOGO_OLIVE_GREEN       = _LOGOS_DIR / "flite_logo_olive_green.png"

# Default logo used by M03 / M04 automation runs
PARTNER_LOGO_PATH      = LOGO_OLIVE_GREEN

# ── Static test assets — Product image ────────────────────────────────────────
# Image file uploaded to the Master Product media card during product creation
# (M07).  Uses the first official Flitesports product image from the shared
# product images folder on the local machine.
PRODUCT_IMAGE_PATH     = Path(r"C:\Flitesport_Product_images\product_img1.png")

# ── Commission defaults used in partner creation / update ─────────────────────
COMMISSION_WHOLESALE        = "45"
COMMISSION_RETAIL           = "67"
COMMISSION_SUGGESTED_RETAIL = "89"
COMMISSION_REP              = "78"
COMMISSION_PARTNER          = "77"
CONTRACT_ENABLED            = "yes"
