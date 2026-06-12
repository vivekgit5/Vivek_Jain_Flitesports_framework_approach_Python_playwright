# Vivek Jain — Flitesports Framework Automation (Python Playwright)

A professional, modular end-to-end automation framework for the **Flitesports**
Shopify eCommerce platform and its associated CRM / Falcon backend.  
Built with **Python + Playwright** following a **Page Object Model (POM)** architecture.

---

## Project Structure

```
Vivek_Jain_Custom_changes_Flitesports_Framework_Automation_Playwriter_Python/
│
├── config/
│   └── settings.py                    # All URLs, credentials, and environment constants
│
├── core/
│   ├── browser.py                     # Browser session factory (context manager)
│   └── reporter.py                    # XLSX report generator (step timings, pass/fail)
│
├── pages/                             # Page Object Model — one class per application area
│   ├── login_page.py                  # CRM login + sign-out
│   ├── users_page.py                  # CRM Users module (navigate, filter, search, edit)
│   ├── shopify_page.py                # Shopify storefront (auth, products, checkout)
│   └── falcon_page.py                 # Falcon App Launcher (Order Exporter / Generator)
│
├── modules/
│   │
│   │   # ── Group 1: Vivek Custom Staging Site — Flitesports (M01–M04, M07) ──
│   ├── m01_create_admin_user.py
│   ├── m02_create_sales_rep.py
│   ├── m03_create_new_partner.py
│   ├── m04_update_users.py
│   ├── m07_create_master_product_local_customization.py
│   │
│   └── Vivek_Existing_Site_Flitesports_Including_Falcon/
│       │   # ── Group 2: Vivek Existing Site — Flitesports + Falcon (M05–M06) ──
│       ├── m05_purchase_products_shopify.py
│       └── m06_falcon_order_comparison.py
│
├── data/
│   └── logos/                         # Official Flitesports partner logo variants
│       ├── flite_logo_black.png
│       ├── flite_logo_gray.png
│       ├── flite_logo_red.png
│       └── flite_logo_olive_green.png  # Default logo used in automation
│
├── reports/                           # Auto-generated XLSX execution reports (git-ignored)
├── downloads/                         # Downloaded files from Falcon (git-ignored)
│
├── run_all.py                         # Master runner — executes both groups in sequence
├── run_custom_staging_site.py            # Runner — Group 1 only (M01–M04, CRM staging)
├── run_existing_site.py               # Runner — Group 2 only (M05–M06, Shopify + Falcon)
├── requirements.txt                   # Python package dependencies
└── README.md                          # This file
```

---

## Tech Stack

| Tool | Version | Purpose |
|------|---------|---------|
| Python | 3.11+ | Runtime |
| Playwright | 1.44+ | Browser automation |
| openpyxl | 3.1+ | XLSX report generation |
| pandas | 2.2+ | Excel parsing & comparison (M06) |
| Faker | 24+ | Dynamic test-data generation |

---

## Setup

### 1. Clone the repository

```bash
git clone https://github.com/vivekgit5/Vivek_Jain_Flitesports_framework_approach_Python_playwright.git
cd Vivek_Jain_Flitesports_framework_approach_Python_playwright
```

### 2. Create and activate a virtual environment

```bash
python -m venv .venv

# Windows
.venv\Scripts\activate

# macOS / Linux
source .venv/bin/activate
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Install Playwright browsers

```bash
python -m playwright install chromium
```

---

## Module Groups

The automation suite is divided into two groups that correspond to the two
distinct application environments under test.

---

### Group 1 — Vivek Custom Staging Site: Flitesports

These modules exercise the Flitesports **CRM staging environment** and are
executed first in every full-suite run.

| Module | Description |
|--------|-------------|
| M01 | Create Admin User |
| M02 | Create Sales Representative |
| M03 | Create New Partner (Super-Admin) |
| M04 | Update Admin, Sales Rep & Partner |
| M07 | Create New Master Product (Local Customization) |

#### M01 — Create Admin User
Creates a new CRM admin user with Faker-generated identity and contact details.

```bash
python -m modules.m01_create_admin_user
```

#### M02 — Create Sales Representative
Creates a new Sales Rep user, assigns a partner and account manager.

```bash
python -m modules.m02_create_sales_rep
```

#### M03 — Create New Partner (Super-Admin)
Runs the 5-step partner creation wizard: program identity (including logo
upload), commissions, rep assignment, partner commissions, and season dates.

The **Program Identity** step includes the mandatory **Partner Type** field
(a PrimeVue combobox).  The automation dynamically selects the **first
available option** from the dropdown using a CSS attribute selector
(`[role='option']`) so the step is resilient to future option-order changes.

```bash
python -m modules.m03_create_new_partner
```

#### M04 — Update Admin, Sales Rep, and Partner
Sequentially updates an existing admin user, a sales rep, and a partner record
across three form steps:

- **Admin** — updates first name, last name, and phone number.
- **Sales Rep** — updates role assignment, first name, last name, and phone number.
- **Partner** — updates program name, Partner Type, contact name, commission fields,
  and season start date via the multi-step edit wizard.

**Partner Type field** — The edit wizard resets the Partner Type combobox to an
empty state on open.  The automation dynamically opens the PrimeVue overlay and
selects the **first available option** (e.g. `PARTNER-CLUB`) using a CSS
attribute selector (`[role='option']`) so the step is resilient to any future
option-order or naming changes, without relying on a hardcoded label.

**Commission balancing** — Step 2 reads the read-only *Total FLITE Rep
Commission* percentage displayed on the panel and mirrors it into every editable
numeric input so that the "Representatives commission total" validation passes.

```bash
python -m modules.m04_update_users
```

#### M07 — Create New Master Product (Local Customization)
Logs in as CRM super-admin, opens the first partner collection under
Master Products → Local Customization, and creates a new product end-to-end:

- **Product name** — Faker-generated (`QA VIVEK MASTER PRODUCT <word> <random>`)
- **Description** — short test text entered via Quill rich-text editor
- **Category / gender / sub-gender** — Cotton category, Youth / Boy
- **Product image** — uploads `product_img1.png` from the configured local path
- **Product type** — types "cotton blend hooded sweatshirt" into the autocomplete;
  tries ALL-CAPS, mixed-case, and lowercase suggestion variants for resilience
  against CRM CSS `text-transform` differences, falling back to Tab commit
- **Color** — GREY
- **Sport attributes** — BASEBALL, BASKETBALL, LACROSSE
- **FLITE platform group sport(s)** — selects "COTTON-BLEND COLLECTION - HEAT PRESS"
- **Custom option** — adds Player Number add-on with YES / NO values
- **Variant** — adds YOUTH SMALL size; fills Faker-generated SKU and price per row
- **Assign variant images** — clicks the image-assign trigger on each row (2 of 2)
- **Save** — submits the create form; waits for `wait_for_url` to confirm the SPA
  has navigated to the correct collection-management URL (not an intermediate route)
- **Logo placements** — checks `input.sp-logo-checkbox` elements on the variants
  edit page (post-save); waits for the `gl-overlay` loading spinner to clear
  before interacting so pointer events are never intercepted
- **Save variant page** — persists sidebar attributes; waits for overlay clearance
  and falls back to `button.sp-btn-primary` CSS class if labelled buttons are absent
- **URL verification** — asserts `/super-admin/admin/master-products` matches exactly
- **Sign out**

```bash
python -m modules.m07_create_master_product_local_customization
```

**To run Group 1 in isolation** (without executing the Shopify / Falcon modules):

```bash
python run_custom_staging_site.py
```

---

### Group 2 — Vivek Existing Site: Flitesports (Including Falcon)

These modules exercise the **live Flitesports Shopify storefront** and the
**Falcon** order-management application.  They are housed under the dedicated
sub-package `modules/Vivek_Existing_Site_Flitesports_Including_Falcon/` and
run sequentially after Group 1 completes in every full-suite run.

| Module | Description |
|--------|-------------|
| M05 | Purchase Products on Shopify Staging (End-to-End) |
| M06 | Falcon Stage: Order Exporter vs Order Generator Comparison |

#### M05 — Purchase Products on Shopify Staging (End-to-End)
> **Group 2 — Vivek Existing Site: Flitesports** &nbsp;|&nbsp; `modules/Vivek_Existing_Site_Flitesports_Including_Falcon/m05_purchase_products_shopify.py`

Full storefront purchase flow: store auth → product search → add 4 products →
cart confirmation → guest checkout (shipping + payment) → order confirmation.

```bash
python -m modules.Vivek_Existing_Site_Flitesports_Including_Falcon.m05_purchase_products_shopify
```

#### M06 — Falcon Stage: Order Exporter vs Order Generator Comparison
> **Group 2 — Vivek Existing Site: Flitesports** &nbsp;|&nbsp; `modules/Vivek_Existing_Site_Flitesports_Including_Falcon/m06_falcon_order_comparison.py`

Downloads both Excel workbooks from the Falcon staging environment, performs a
field-by-field comparison by Order ID, and produces two detailed XLSX reports.

```bash
python -m modules.Vivek_Existing_Site_Flitesports_Including_Falcon.m06_falcon_order_comparison
```

**To run Group 2 in isolation** (without executing the CRM staging modules):

```bash
python run_existing_site.py
```

---

## Running the Full Suite

```bash
python run_all.py
```

`run_all.py` executes both groups **in sequence** and prints a consolidated
summary table grouped by environment, showing the outcome and total duration
(in seconds) for each module.

## Running a Specific Group

| Command | Scope |
|---------|-------|
| `python run_custom_staging_site.py` | Group 1 only — Vivek Custom Staging Site: Flitesports (M01–M04, M07) |
| `python run_existing_site.py` | Group 2 only — Vivek Existing Site: Flitesports Including Falcon (M05–M06) |
| `python run_all.py` | Both groups in sequence (full suite) |

**Sample output (`run_all.py`):**

```
[ Vivek Custom Staging Site — Flitesports ]
  Code    Module                                          Status    Duration (s)
  ──────────────────────────────────────────────────────────────────────────
  ✓ M01   Create Admin User                               PASSED    12.30s
  ✓ M02   Create Sales Representative                     PASSED    14.10s
  ✓ M03   Create New Partner (Super-Admin)                PASSED    47.80s
  ✓ M04   Update Admin, Sales Rep & Partner               PASSED    62.50s
  ✓ M07   Create New Master Product                       PASSED    55.20s

[ Vivek Existing Site — Flitesports (Including Falcon) ]
  ✓ M05   Purchase Products on Shopify Staging            PASSED    38.20s
  ✓ M06   Falcon Order Exporter vs Generator              PASSED    55.90s
```

---

## Reports

Every module writes a timestamped XLSX report to the `reports/` directory
automatically on completion — whether the run passes or fails.

### Report filename pattern

```
reports/<ModuleName>_<YYYYMMDD_HHMMSS>.xlsx
```

**Example:**
```
reports/M01_–_Create_Admin_User_20260606_160512.xlsx
```

### Report structure

Each XLSX file contains two sheets:

| Sheet | Contents |
|-------|----------|
| **Summary** | Module name, run date, start/end time, total duration in seconds, step counts (passed / failed / info), overall PASSED / FAILED result, project and author details |
| **Steps** | Sequential step log — step number, name, status (PASS / FAIL / INFO), detail notes, timestamp, step duration (seconds), cumulative time (seconds) |

The **Step Duration** and **Cumulative Time** columns make it easy to identify
slow operations and optimise the automation flow.

---

## Environment Configuration

All environment-specific values live in **`config/settings.py`**.  
Update this file when URLs, credentials, or test-data defaults need to change —
no module code needs editing for environment switches.

| Constant | Description |
|----------|-------------|
| `CRM_BASE_URL` | Flitesports CRM staging URL |
| `CRM_ADMIN_EMAIL` | Admin login email |
| `CRM_ADMIN_PASS` | Admin login password |
| `SHOPIFY_BASE_URL` | Shopify staging store URL |
| `SHOPIFY_STORE_PASSWORD` | Store access password |
| `FALCON_APP_LAUNCHER_URL` | Falcon staging launcher URL |
| `FALCON_ORDER_RANGE` | Order number range for M06 comparison |

---

## Data-Generation Constraints

### Partner Program Name — Alphanumeric-Only Enforcement

The **Program Name** field in the Flitesports CRM does not accept special
characters (e.g. `&`, `-`, `.`, `'`, `,`).  
To ensure reliable form submission across **M03** and **M04**, all
Faker-generated company names are sanitised at data-generation time using a
regular-expression filter that retains only alphanumeric characters and spaces:

```python
re.sub(r'[^A-Za-z0-9 ]', '', fake.company()).strip()
```

This constraint is enforced in:

| Module | Field affected |
|--------|---------------|
| `m03_create_new_partner.py` | `program_name` (Partner creation wizard — Step 1) |
| `m04_update_users.py` | `program_name` (Partner update flow) |

---

## Changelog

### 2026-06-12
- **fix(m07):** Resolved four independent issues in the Create Master Product flow:
  - **Navigation race** — replaced `wait_for_load_state('networkidle')` with
    `wait_for_url('**/collection-management**')` so the logged URL always reflects
    the true destination instead of the intermediate `/dashboard` SPA hop.
  - **Product type autocomplete** — tries ALL-CAPS, mixed-case, and lowercase
    suggestion variants before falling back to the first autocomplete list item,
    then Tab commit; handles CRM CSS `text-transform` differences robustly.
  - **Logo placements** — moved `_check_logo_placements()` to after the initial
    save redirect where the section actually renders on the variants edit page;
    targets checkboxes by `input.sp-logo-checkbox` CSS class, removing the
    brittle section-heading text dependency.
  - **`gl-overlay` blocking** — added `wait_for_selector('.gl-overlay', state='hidden')`
    before both logo-placement clicks and the save button so the CRM loading
    spinner never intercepts pointer events.
  - **Save button fallback** — added `button.sp-btn-primary` CSS class fallback
    after the labelled-button loop so the step succeeds even when button text or
    text-transform differs from expected values.
- **fix(m04):** Corrected Partner Type dropdown selection in the partner update
  wizard — replaced `get_by_role("option")` (which matched hidden native
  `<option>` elements and caused a 5 s timeout) with the CSS attribute selector
  `[role='option']` to target only visible PrimeVue overlay items.  The first
  available option is now selected dynamically, making the step label-agnostic
  and resilient to future CRM option changes.

### 2026-06-11
- **feat(m07):** Add `m07_create_master_product_local_customization` — full
  end-to-end master product creation under Local Customization, including
  Quill rich-text description, variant/SKU/price management, image assignment,
  logo-placement selection, and URL verification.

### 2026-06-10
- **fix(m03):** Handle mandatory Partner Type field in the New Partner wizard —
  dynamically select the first PrimeVue overlay option using `[role='option']`
  (consistent pattern now also applied to M04).
- **refactor(runners):** Rename `run_staging.py` → `run_custom_staging_site.py`
  for naming consistency with `run_existing_site.py`.

---

## Author

**Vivek Jain**  
QA Automation Engineer — Bitcot  
Project: Flitesports (Shopify eCommerce + CRM Automation)
