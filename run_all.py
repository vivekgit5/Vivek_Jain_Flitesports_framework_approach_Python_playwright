"""
run_all.py – Master test runner for the Flitesports automation framework.

Executes all modules in two logically separated groups that reflect the
two distinct application environments under test:

  Group 1 – Vivek Custom Staging Site — Flitesports
  -------------------------------------------------------
  M01  Create Admin User
  M02  Create Sales Representative
  M03  Create New Partner (Super-Admin)
  M04  Update Admin, Sales Rep & Partner

  Group 2 – Vivek Existing Site — Flitesports (Including Falcon)
  ---------------------------------------------------------------
  M05  Purchase Products on Shopify Staging (End-to-End)
  M06  Falcon Order Exporter vs Order Generator Comparison

Each group runs sequentially.  A consolidated summary table is printed at
the end showing the outcome and duration of every module across both groups.

Usage
-----
    python run_all.py

To run a single module independently, see the individual commands in README.md.
"""

import importlib
import logging
import sys
import time
from dataclasses import dataclass

log = logging.getLogger(__name__)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)

# ── Module groups ──────────────────────────────────────────────────────────────
#
# Each entry is a tuple of:
#   (module_code, importable_path, friendly_name)
#
# Group 1: CRM / custom staging-site modules (M01–M04)
GROUP_STAGING = {
    "label": "Vivek Custom Staging Site — Flitesports",
    "modules": [
        ("M01", "modules.m01_create_admin_user",        "Create Admin User"),
        ("M02", "modules.m02_create_sales_rep",         "Create Sales Representative"),
        ("M03", "modules.m03_create_new_partner",       "Create New Partner (Super-Admin)"),
        ("M04", "modules.m04_update_users",             "Update Admin, Sales Rep & Partner"),
    ],
}

# Group 2: Existing Flitesports site + Falcon modules (M05–M06)
GROUP_EXISTING = {
    "label": "Vivek Existing Site — Flitesports (Including Falcon)",
    "modules": [
        (
            "M05",
            "modules.Vivek_Existing_Site_Flitesports_Including_Falcon"
            ".m05_purchase_products_shopify",
            "Purchase Products on Shopify Staging",
        ),
        (
            "M06",
            "modules.Vivek_Existing_Site_Flitesports_Including_Falcon"
            ".m06_falcon_order_comparison",
            "Falcon Order Exporter vs Generator",
        ),
    ],
}

ALL_GROUPS = [GROUP_STAGING, GROUP_EXISTING]


@dataclass
class ModuleResult:
    group:    str
    code:     str
    name:     str
    status:   str   # PASSED | FAILED
    duration: float


def _run_module(module_path: str) -> None:
    """Import and execute the module's ``run()`` entry point."""
    mod   = importlib.import_module(module_path)
    entry = getattr(mod, "run", None)
    if entry is None:
        raise AttributeError(
            f"Module '{module_path}' does not expose a 'run' function."
        )

    import inspect
    from playwright.sync_api import sync_playwright

    sig = inspect.signature(entry)
    if sig.parameters:
        # Module expects a Playwright handle (M01–M05)
        with sync_playwright() as pw:
            entry(pw)
    else:
        # Module manages Playwright internally (M06)
        entry()


def _run_group(group: dict, results: list) -> None:
    """Execute all modules in *group* sequentially and append to *results*."""
    label   = group["label"]
    modules = group["modules"]

    log.info("")
    log.info("=" * 70)
    log.info("  GROUP: %s", label)
    log.info("  Modules to execute: %d", len(modules))
    log.info("=" * 70)

    for code, module_path, friendly_name in modules:
        log.info("")
        log.info("──  Running %s: %s  ──", code, friendly_name)
        start = time.time()
        try:
            _run_module(module_path)
            duration = round(time.time() - start, 2)
            results.append(ModuleResult(label, code, friendly_name, "PASSED", duration))
            log.info("✓  %s completed in %.2fs", code, duration)
        except Exception as exc:
            duration = round(time.time() - start, 2)
            results.append(ModuleResult(label, code, friendly_name, "FAILED", duration))
            log.error("✗  %s FAILED after %.2fs — %s", code, duration, exc)


def main() -> None:
    results: list[ModuleResult] = []
    overall_start = time.time()

    log.info("=" * 70)
    log.info("  Flitesports Automation Framework — Full Suite Run")
    log.info("  Groups: %d  |  Total modules: %d",
             len(ALL_GROUPS),
             sum(len(g["modules"]) for g in ALL_GROUPS))
    log.info("=" * 70)

    for group in ALL_GROUPS:
        _run_group(group, results)

    total_duration = round(time.time() - overall_start, 2)

    # ── Consolidated summary ───────────────────────────────────────────────────
    log.info("")
    log.info("=" * 70)
    log.info("  SUITE SUMMARY  —  Total duration: %.2fs", total_duration)
    log.info("=" * 70)

    current_group = None
    for r in results:
        if r.group != current_group:
            log.info("")
            log.info("  [ %s ]", r.group)
            log.info("  %-6s  %-46s  %-8s  %s",
                     "Code", "Module", "Status", "Duration (s)")
            log.info("  %s", "-" * 66)
            current_group = r.group
        symbol = "✓" if r.status == "PASSED" else "✗"
        log.info(
            "  %s %-4s  %-46s  %-8s  %.2fs",
            symbol, r.code, r.name, r.status, r.duration,
        )

    log.info("")
    log.info("=" * 70)
    passed = sum(1 for r in results if r.status == "PASSED")
    failed = sum(1 for r in results if r.status == "FAILED")
    log.info("  PASSED: %d  |  FAILED: %d  |  TOTAL: %d",
             passed, failed, len(results))
    log.info("=" * 70)

    sys.exit(0 if failed == 0 else 1)


if __name__ == "__main__":
    main()
