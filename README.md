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

The **Program Identity** step now includes the mandatory **Partner Type** field
(a PrimeVue combobox added to the Edit Program form).  The automation
dynamically selects the **first available option** (`PARTNER-CLUB`) from the
dropdown so the step is resilient to future option-order changes.

```bash
python -m modules.m03_create_new_partner
```

#### M04 — Update Admin, Sales Rep, and Partner
Sequentially updates an existing admin user, a sales rep, and a partner record
(including partner type selection and commission equality balancing).

```bash
python -m modules.m04_update_users
```

#### M07 — Create New Master Product (Local Customization)
Logs in as CRM super-admin, opens the first partner collection under
Master Products → Local Customization, and creates a new product end-to-end:

- **Product name** — Faker-generated (`QA VIVEK MASTER PRODUCT MAN <random>`)
- **Description** — short hardcoded test text entered via Quill rich-text editor
- **Category / gender / sub-gender** — Cotton category, Youth / Boy
- **Product image** — uploads `product_img1.png` from the configured local path
- **Product type** — types "cotton blend hooded sweatshirt" into the autocomplete and selects the suggestion
- **Color** — GREY
- **Sport attributes** — BASEBALL, BASKETBALL, LACROSSE
- **FLITE platform group sport(s)** — selects "COTTON-BLEND COLLECTION - HEAT PRESS"
- **Custom option** — adds Player Number add-on with YES / NO values
- **Variant** — adds YOUTH SMALL size; fills unique Faker-generated SKU and price for both rows
- **Assign variant images** — clicks the image-assign trigger on each variant row and confirms via the modal (2 of 2)
- **Logo placements** — enables available placement checkboxes
- **Save** — submits the create form and waits for network idle
- **URL verification** — navigates to `/super-admin/admin/master-products` and asserts the URL matches exactly
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

## Author

**Vivek Jain**  
QA Automation Engineer — Bitcot  
Project: Flitesports (Shopify eCommerce + CRM Automation)
