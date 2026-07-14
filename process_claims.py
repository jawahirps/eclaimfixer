#!/usr/bin/env python3
"""DHA eClaim batch automation runner with control, validation, and resumable state tracking."""

import os
import sys
import logging
import argparse
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, List, Tuple

from dotenv import load_dotenv
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright, Page, expect


# Logging setup
def setup_logging(log_level=logging.INFO):
    """Configure logging to file and console."""
    log_dir = Path(".")
    timestamp = datetime.now().strftime("%Y%m%d")
    log_file = log_dir / f"run_{timestamp}.log"

    logging.basicConfig(
        level=log_level,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler(sys.stdout),
        ],
    )
    logger = logging.getLogger(__name__)
    return logger


logger = None  # Initialized after arg parsing


class ClaimValidator:
    """Validate input file and claims before processing."""

    @staticmethod
    def validate_file(input_file: str) -> Tuple[bool, str]:
        """Validate Excel file exists and is readable."""
        if not os.path.exists(input_file):
            return False, f"File not found: {input_file}"

        try:
            wb = load_workbook(input_file)
            ws = wb.active
            if ws.max_row < 2:
                return False, "File has no data rows (header only)"
            return True, f"File valid: {ws.max_row - 1} claims"
        except Exception as e:
            return False, f"File read error: {str(e)}"

    @staticmethod
    def validate_structure(input_file: str) -> Tuple[bool, str]:
        """Validate Excel structure and required columns."""
        try:
            wb = load_workbook(input_file)
            ws = wb.active

            # Check headers
            expected_headers = ["#", "SOURCE", "CLAIM_ID", "STATUS", "LOADED_AT", "SAVED_AT", "NOTES"]
            actual_headers = [ws.cell(row=1, column=i).value for i in range(1, 8)]

            if actual_headers != expected_headers:
                return False, f"Invalid headers. Expected: {expected_headers}, Got: {actual_headers}"

            return True, "Structure valid"
        except Exception as e:
            return False, f"Structure validation error: {str(e)}"

    @staticmethod
    def precheck_claims(input_file: str, verbose: bool = False) -> Dict:
        """Precheck all claims in the file."""
        try:
            wb = load_workbook(input_file)
            ws = wb.active

            stats = {
                "total": 0,
                "pending": 0,
                "success": 0,
                "error_count": 0,
                "by_source": {},
                "issues": [],
                "exception": None,
            }

            for row_num in range(2, ws.max_row + 1):
                num = ws.cell(row=row_num, column=1).value
                source = ws.cell(row=row_num, column=2).value
                claim_id = ws.cell(row=row_num, column=3).value
                status = ws.cell(row=row_num, column=4).value

                stats["total"] += 1

                # Count by status
                if status == "PENDING":
                    stats["pending"] += 1
                elif status == "SUCCESS":
                    stats["success"] += 1
                elif status == "ERROR":
                    stats["error_count"] += 1

                # Count by source
                if source not in stats["by_source"]:
                    stats["by_source"][source] = 0
                stats["by_source"][source] += 1

                # Validate claim data
                if not claim_id or not str(claim_id).strip():
                    stats["issues"].append(f"Row {num}: Missing CLAIM_ID")
                if not source or not str(source).strip():
                    stats["issues"].append(f"Row {num}: Missing SOURCE")

            return stats
        except Exception as e:
            return {"exception": str(e)}


class ClaimProcessor:
    """Process DHA eClaim portal claims via browser automation."""

    def __init__(
        self,
        portal_url: str,
        username: str,
        password: str,
        headless: bool = False,
        delay_secs: float = 1.5,
        dry_run: bool = False,
    ):
        self.portal_url = portal_url
        self.username = username
        self.password = password
        self.headless = headless
        self.delay_secs = delay_secs
        self.dry_run = dry_run
        self.page: Optional[Page] = None
        self.browser = None
        self.context = None

    def start_browser(self):
        """Launch browser and login to portal."""
        if self.dry_run:
            logger.info("[DRY RUN] Browser launch skipped")
            return

        playwright = sync_playwright().start()
        self.browser = playwright.chromium.launch(headless=self.headless)
        self.context = self.browser.new_context()
        self.page = self.context.new_page()

        logger.info(f"Navigating to {self.portal_url}")
        self.page.goto(self.portal_url)
        self._login()

    def _login(self):
        """Login to DHA eClaim portal."""
        logger.info(f"Logging in as {self.username}")
        self.page.fill("input[name='username']", self.username)
        self.page.fill("input[name='password']", self.password)
        self.page.click("button[type='submit']")

        # Wait for navigation and page to stabilize
        self.page.wait_for_load_state("networkidle", timeout=30000)
        logger.info("Login successful")

    def process_claim(self, claim_id: str) -> dict:
        """Process a single claim through the portal.

        Returns dict with keys: status, loaded_at, saved_at, notes
        """
        result = {
            "status": "ERROR",
            "loaded_at": None,
            "saved_at": None,
            "notes": "",
        }

        try:
            result["loaded_at"] = datetime.now().isoformat()

            if self.dry_run:
                logger.info(f"[DRY RUN] Would process: {claim_id}")
                result["status"] = "SUCCESS"
                result["notes"] = "[DRY RUN] Simulated successful processing"
                return result

            # 1. FETCH: Search claim by ID
            logger.info(f"[1/5] FETCH: {claim_id}")
            self.page.goto(f"{self.portal_url}/claims/search")
            self.page.wait_for_load_state("networkidle", timeout=30000)

            search_field = self.page.locator("input[placeholder*='Claim ID']")
            search_field.fill(claim_id)
            search_field.press("Enter")
            self.page.wait_for_load_state("networkidle", timeout=30000)

            # 2. LOAD: Open claim detail/edit form
            logger.info(f"[2/5] LOAD: {claim_id}")
            claim_link = self.page.locator(f"text={claim_id}")

            if claim_link.count() == 0:
                result["notes"] = "Claim not found in search results"
                result["status"] = "ERROR"
                logger.error(f"  ✗ Not found")
                return result

            claim_link.first.click()
            self.page.wait_for_load_state("networkidle", timeout=30000)

            # 3. RELOAD: Click Reload button
            logger.info(f"[3/5] RELOAD: {claim_id}")
            reload_button = self.page.locator("button:has-text('Reload')")

            if reload_button.count() > 0:
                reload_button.click()
                self.page.wait_for_load_state("networkidle", timeout=30000)
                logger.info(f"  ✓ Reloaded")
            else:
                logger.warning(f"  ! Reload button not found (continuing)")

            # 4. VERIFY: Confirm claim ID matches
            logger.info(f"[4/5] VERIFY: {claim_id}")
            claim_id_display = self.page.locator(f"text={claim_id}")

            if claim_id_display.count() == 0:
                result["notes"] = "Claim ID mismatch after reload"
                result["status"] = "ERROR"
                logger.error(f"  ✗ Verification failed")
                return result

            logger.info(f"  ✓ Verified")

            # 5. SAVE: Click Save button
            logger.info(f"[5/5] SAVE: {claim_id}")
            save_button = self.page.locator("button:has-text('Save')")

            if save_button.count() == 0:
                result["notes"] = "Save button not found"
                result["status"] = "ERROR"
                logger.error(f"  ✗ Save button not found")
                return result

            save_button.click()

            # Wait for save confirmation
            try:
                self.page.wait_for_load_state("networkidle", timeout=30000)
                success_msg = self.page.locator("text=/Success|Saved|Updated/i")
                if success_msg.count() > 0:
                    result["status"] = "SUCCESS"
                    result["notes"] = success_msg.first.text_content()
                    logger.info(f"  ✓ Saved: {result['notes']}")
                else:
                    result["status"] = "SUCCESS"
                    result["notes"] = "Saved (no confirmation message)"
                    logger.info(f"  ✓ Saved")
            except Exception as e:
                result["status"] = "ERROR"
                result["notes"] = f"Save timeout or error: {str(e)}"
                logger.error(f"  ✗ Save error: {str(e)}")

        except Exception as e:
            result["status"] = "ERROR"
            result["notes"] = f"Processing error: {str(e)}"
            logger.error(f"  ✗ Exception: {str(e)}")

        finally:
            result["saved_at"] = datetime.now().isoformat()

        return result

    def close(self):
        """Close browser."""
        if self.page:
            self.page.close()
        if self.context:
            self.context.close()
        if self.browser:
            self.browser.close()
        logger.info("Browser closed")


def process_claims_file(
    input_file: str,
    skip_precheck: bool = False,
    dry_run: bool = False,
    only_precheck: bool = False,
):
    """Read claims from Excel and process each PENDING claim.

    Args:
        input_file: Path to Excel file with claims
        skip_precheck: Skip validation before processing
        dry_run: Simulate processing without browser automation
        only_precheck: Only run precheck, don't process
    """
    # Validate file
    valid, msg = ClaimValidator.validate_file(input_file)
    if not valid:
        logger.error(msg)
        sys.exit(1)

    logger.info(f"Input file: {input_file}")

    # Validate structure
    valid, msg = ClaimValidator.validate_structure(input_file)
    if not valid:
        logger.error(msg)
        sys.exit(1)

    # Precheck claims
    logger.info("Precheck: analyzing claims...")
    stats = ClaimValidator.precheck_claims(input_file)

    if "exception" in stats and stats["exception"]:
        logger.error(f"Precheck error: {stats['exception']}")
        sys.exit(1)

    logger.info(f"  Total: {stats['total']} claims")
    logger.info(f"  Pending: {stats['pending']} | Success: {stats['success']} | Error: {stats['error_count']}")
    for source, count in stats["by_source"].items():
        logger.info(f"    {source}: {count}")

    if stats["issues"]:
        logger.warning(f"  Issues found ({len(stats['issues'])}):")
        for issue in stats["issues"][:5]:
            logger.warning(f"    - {issue}")
        if len(stats["issues"]) > 5:
            logger.warning(f"    ... and {len(stats['issues']) - 5} more")

    if only_precheck:
        logger.info("Precheck complete. Use --only-precheck=False to process.")
        return

    if stats["pending"] == 0:
        logger.info("No PENDING claims to process. Exiting.")
        return

    if not skip_precheck and stats["issues"]:
        logger.warning("Issues found in precheck. Fix before processing or use --skip-precheck")
        sys.exit(1)

    # Load environment
    load_dotenv()
    portal_url = os.getenv("PORTAL_URL")
    portal_user = os.getenv("PORTAL_USER")
    portal_pass = os.getenv("PORTAL_PASS")
    headless = os.getenv("HEADLESS", "False").lower() == "true"
    delay_secs = float(os.getenv("DELAY_SECS", "1.5"))

    if not dry_run and not all([portal_url, portal_user, portal_pass]):
        logger.error("Missing required environment variables: PORTAL_URL, PORTAL_USER, PORTAL_PASS")
        sys.exit(1)

    mode = "[DRY RUN]" if dry_run else ""
    logger.info(f"{mode} Starting claim processing")
    if not dry_run:
        logger.info(f"  Portal: {portal_url} | Headless: {headless} | Delay: {delay_secs}s")

    # Initialize processor and browser
    processor = ClaimProcessor(portal_url, portal_user, portal_pass, headless, delay_secs, dry_run)
    processor.start_browser()

    try:
        # Load workbook
        wb = load_workbook(input_file)
        ws = wb.active

        processed = 0
        skipped = 0

        logger.info(f"Processing {stats['pending']} PENDING claims...")

        # Process rows
        for row_num in range(2, ws.max_row + 1):
            claim_num_cell = ws.cell(row=row_num, column=1)
            source_cell = ws.cell(row=row_num, column=2)
            claim_id_cell = ws.cell(row=row_num, column=3)
            status_cell = ws.cell(row=row_num, column=4)
            loaded_at_cell = ws.cell(row=row_num, column=5)
            saved_at_cell = ws.cell(row=row_num, column=6)
            notes_cell = ws.cell(row=row_num, column=7)

            current_status = status_cell.value
            claim_id = claim_id_cell.value
            source = source_cell.value
            claim_num = claim_num_cell.value

            # Skip non-PENDING rows
            if current_status != "PENDING":
                skipped += 1
                continue

            progress = f"[{processed + 1}/{stats['pending']}]"
            logger.info(f"{progress} Claim {claim_num} ({source})")

            # Process claim
            result = processor.process_claim(claim_id)

            # Update spreadsheet
            status_cell.value = result["status"]
            loaded_at_cell.value = result["loaded_at"]
            saved_at_cell.value = result["saved_at"]
            notes_cell.value = result["notes"]

            # Save after each row (resumable)
            wb.save(input_file)
            processed += 1

            # Small delay between claims
            if delay_secs > 0 and not dry_run:
                import time

                time.sleep(delay_secs)

        logger.info(f"✓ Processing complete: {processed} processed, {skipped} skipped")

    except Exception as e:
        logger.error(f"Fatal error during processing: {str(e)}")
        sys.exit(1)

    finally:
        processor.close()


def main():
    """CLI entry point with argument parsing."""
    parser = argparse.ArgumentParser(
        description="DHA eClaim batch processing automation with state tracking",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s script_input.xlsx              # Process with precheck
  %(prog)s claims.xlsx --dry-run          # Simulate processing
  %(prog)s claims.xlsx --only-precheck    # Just validate
  %(prog)s claims.xlsx --skip-precheck    # Process without validation
        """,
    )

    parser.add_argument("input_file", nargs="?", default="script_input.xlsx", help="Excel file with claims (default: script_input.xlsx)")
    parser.add_argument("--dry-run", action="store_true", help="Simulate processing without browser automation")
    parser.add_argument("--only-precheck", action="store_true", help="Only validate claims, don't process")
    parser.add_argument("--skip-precheck", action="store_true", help="Skip validation before processing")
    parser.add_argument("-v", "--verbose", action="store_true", help="Verbose logging (DEBUG level)")

    args = parser.parse_args()

    # Initialize logger with verbosity
    global logger
    log_level = logging.DEBUG if args.verbose else logging.INFO
    logger = setup_logging(log_level)

    logger.debug(f"Arguments: {args}")

    process_claims_file(
        input_file=args.input_file,
        skip_precheck=args.skip_precheck,
        dry_run=args.dry_run,
        only_precheck=args.only_precheck,
    )


if __name__ == "__main__":
    main()
