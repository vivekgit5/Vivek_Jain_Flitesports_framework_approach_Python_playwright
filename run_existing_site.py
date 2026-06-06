"""
run_existing_site.py – Runner for Group 2: Vivek Existing Site — Flitesports (Including Falcon)

Executes only the existing-site modules (M05–M06) in sequence:

  M05  Purchase Products on Shopify Staging (End-to-End)
  M06  Falcon Order Exporter vs Order Generator Comparison

Use this runner when you want to validate the live Flitesports Shopify
storefront and the Falcon order-management application in isolation,
without executing the CRM staging modules.

Usage
-----
    python run_existing_site.py

To run the full suite (both groups), use:
    python run_all.py

To run the staging group only (M01–M04), use:
    python run_custom_staging_site.py
"""

import logging
import sys
import time

from run_all import GROUP_EXISTING, ModuleResult, _run_group

log = logging.getLogger(__name__)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)


def main() -> None:
    results: list[ModuleResult] = []
    overall_start = time.time()

    log.info("=" * 70)
    log.info("  Flitesports Automation — Group 2 Run")
    log.info("  Vivek Existing Site — Flitesports (Including Falcon)")
    log.info("  Modules to execute: %d", len(GROUP_EXISTING["modules"]))
    log.info("=" * 70)

    _run_group(GROUP_EXISTING, results)

    total_duration = round(time.time() - overall_start, 2)

    log.info("")
    log.info("=" * 70)
    log.info("  SUMMARY  —  Total duration: %.2fs", total_duration)
    log.info("=" * 70)
    log.info("  %-6s  %-46s  %-8s  %s",
             "Code", "Module", "Status", "Duration (s)")
    log.info("  %s", "-" * 66)
    for r in results:
        symbol = "✓" if r.status == "PASSED" else "✗"
        log.info(
            "  %s %-4s  %-46s  %-8s  %.2fs",
            symbol, r.code, r.name, r.status, r.duration,
        )
    log.info("=" * 70)

    passed = sum(1 for r in results if r.status == "PASSED")
    failed = sum(1 for r in results if r.status == "FAILED")
    log.info("  PASSED: %d  |  FAILED: %d  |  TOTAL: %d",
             passed, failed, len(results))
    log.info("=" * 70)

    sys.exit(0 if failed == 0 else 1)


if __name__ == "__main__":
    main()
