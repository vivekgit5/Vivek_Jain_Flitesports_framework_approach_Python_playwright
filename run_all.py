"""
run_all.py – Master test runner for the Flitesports automation framework.

Executes every module in sequence and prints a final consolidated summary
showing the outcome and duration of each module.

Usage
-----
    python run_all.py

To run a single module instead, see the individual commands in README.md.
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

# ── Ordered list of modules to execute ────────────────────────────────────────
MODULES = [
    ("M01", "modules.m01_create_admin_user",        "Create Admin User"),
    ("M02", "modules.m02_create_sales_rep",         "Create Sales Representative"),
    ("M03", "modules.m03_create_new_partner",       "Create New Partner (Super-Admin)"),
    ("M04", "modules.m04_update_users",             "Update Admin, Sales Rep & Partner"),
    ("M05", "modules.m05_purchase_products_shopify","Purchase Products on Shopify Staging"),
    ("M06", "modules.m06_falcon_order_comparison",  "Falcon Order Exporter vs Generator"),
]


@dataclass
class ModuleResult:
    code:     str
    name:     str
    status:   str   # PASSED | FAILED | SKIPPED
    duration: float


def _run_module(module_path: str) -> None:
    """Import and execute the module's ``run()`` entry point."""
    mod = importlib.import_module(module_path)
    # All modules expose either run(playwright) or run() depending on whether
    # they need a Playwright handle or manage it internally.
    from playwright.sync_api import sync_playwright

    entry = getattr(mod, "run", None)
    if entry is None:
        raise AttributeError(f"Module '{module_path}' does not expose a 'run' function.")

    import inspect
    sig = inspect.signature(entry)
    if sig.parameters:
        # Module accepts a playwright argument (M01–M05)
        with sync_playwright() as pw:
            entry(pw)
    else:
        # Module manages playwright internally (M06)
        entry()


def main() -> None:
    results: list[ModuleResult] = []
    overall_start = time.time()

    log.info("=" * 70)
    log.info("  Flitesports Automation Framework – Full Suite Run")
    log.info("  Modules to execute: %d", len(MODULES))
    log.info("=" * 70)

    for code, module_path, friendly_name in MODULES:
        log.info("")
        log.info("──  Running %s: %s  ──", code, friendly_name)
        start = time.time()
        try:
            _run_module(module_path)
            duration = round(time.time() - start, 2)
            results.append(ModuleResult(code, friendly_name, "PASSED", duration))
            log.info("✓  %s completed in %.2fs", code, duration)
        except Exception as exc:
            duration = round(time.time() - start, 2)
            results.append(ModuleResult(code, friendly_name, "FAILED", duration))
            log.error("✗  %s FAILED after %.2fs — %s", code, duration, exc)

    total_duration = round(time.time() - overall_start, 2)

    # ── Consolidated summary ───────────────────────────────────────────────────
    log.info("")
    log.info("=" * 70)
    log.info("  SUITE SUMMARY  —  Total duration: %.2fs", total_duration)
    log.info("=" * 70)
    log.info("  %-6s  %-48s  %-8s  %s", "Code", "Module", "Status", "Duration (s)")
    log.info("  %s", "-" * 66)
    for r in results:
        symbol = "✓" if r.status == "PASSED" else "✗"
        log.info(
            "  %s %-4s  %-48s  %-8s  %.2fs",
            symbol, r.code, r.name, r.status, r.duration,
        )
    log.info("=" * 70)

    passed = sum(1 for r in results if r.status == "PASSED")
    failed = sum(1 for r in results if r.status == "FAILED")
    log.info("  PASSED: %d  |  FAILED: %d  |  TOTAL: %d", passed, failed, len(results))
    log.info("=" * 70)

    sys.exit(0 if failed == 0 else 1)


if __name__ == "__main__":
    main()
