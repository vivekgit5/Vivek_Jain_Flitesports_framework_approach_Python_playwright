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
│   └── settings.py          # All URLs, credentials, and environment constants
│
├── core/
│   ├── browser.py           # Browser session factory (context manager)
│   └── reporter.py          # XLSX report generator (step timings, pass/fail)
│
├── pages/                   # Page Object Model — one class per application area
│   ├── login_page.py        # CRM login + sign-out
│   ├── users_page.py        # CRM Users module (navigate, filter, search, edit)
│   ├── shopify_page.py      # Shopify storefront (auth, products, checkout)
│   └── falcon_page.py       # Falcon App Launcher (Order Exporter / Generator)
│
├── modules/                 # Executable automation modules (M01–M06)
│   ├── m01_create_admin_user.py
│   ├── m02_create_sales_rep.py
│   ├── m03_create_new_partner.py
│   ├── m04_update_users.py
│   ├── m05_purchase_products_shopify.py
│   └── m06_falcon_order_comparison.py
│
├── data/                    # Static test data files (JSON recordings, etc.)
├── reports/                 # Auto-generated XLSX execution reports (git-ignored)
├── downloads/               # Downloaded files from Falcon (git-ignored)
│
├── run_all.py               # Master runner — executes all modules in sequence
├── requirements.txt         # Python package dependencies
└── README.md                # This file
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

## Running Individual Modules

All modules are run from the **project root** directory.

### M01 — Create Admin User
Creates a new CRM admin user with Faker-generated identity and contact details.

```bash
python -m modules.m01_create_admin_user
```

### M02 — Create Sales Representative
Creates a new Sales Rep user, assigns a partner and account manager.

```bash
python -m modules.m02_create_sales_rep
```

### M03 — Create New Partner (Super-Admin)
Runs the 5-step partner creation wizard: program identity, commissions, rep
assignment, partner commissions, and season dates.

```bash
python -m modules.m03_create_new_partner
```

### M04 — Update Admin, Sales Rep, and Partner
Sequentially updates an existing admin user, a sales rep, and a partner record.

```bash
python -m modules.m04_update_users
```

### M05 — Purchase Products on Shopify Staging (End-to-End)
Full storefront purchase flow: store auth → product search → add 4 products →
cart confirmation → guest checkout (shipping + payment) → order confirmation.

```bash
python -m modules.m05_purchase_products_shopify
```

### M06 — Falcon Stage: Order Exporter vs Order Generator Comparison
Downloads both Excel workbooks from the Falcon staging environment, performs a
field-by-field comparison by Order ID, and produces two detailed XLSX reports.

```bash
python -m modules.m06_falcon_order_comparison
```

---

## Running All Modules in Sequence

```bash
python run_all.py
```

This executes M01 through M06 in order and prints a consolidated summary table
showing the outcome and total duration (in seconds) for each module.

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
