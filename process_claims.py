#!/usr/bin/env python3
"""DHA eClaim batch automation runner with resumable state tracking."""

import os
import sys
import logging
from datetime import datetime
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright, Page, expect


# Logging setup
def setup_logging():
    """Configure logging to file and console."""
    log_dir = Path(".")
    timestamp = datetime.now().strftime("%Y%m%d")
    log_file = log_dir / f"run_{timestamp}.log"

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler(sys.stdout),
        ],
    )
    logger = logging.getLogger(__name__)
    return logger


logger = setup_logging()


class ClaimProcessor:
    """Process DHA eClaim portal claims via browser automation."""

    def __init__(self, portal_url: str, username: str, password: str, headless: bool = False, delay_secs: float = 1.5):
        self.portal_url = portal_url
        self.username = username
        self.password = password
        self.headless = headless
        self.delay_secs = delay_secs
        self.page: Optional[Page] = None
        self.browser = None
        self.context = None

    def start_browser(self):
        """Launch browser and login to portal."""
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
            # 1. FETCH: Search claim by ID
            logger.info(f"Fetching claim {claim_id}")
            result["loaded_at"] = datetime.now().isoformat()

            # Navigate to search/list page
            self.page.goto(f"{self.portal_url}/claims/search")
            self.page.wait_for_load_state("networkidle", timeout=30000)

            # Fill search field and submit
            search_field = self.page.locator("input[placeholder*='Claim ID']")
            search_field.fill(claim_id)
            search_field.press("Enter")

            self.page.wait_for_load_state("networkidle", timeout=30000)

            # 2. LOAD: Open claim detail/edit form
            logger.info(f"Opening claim detail for {claim_id}")
            claim_link = self.page.locator(f"text={claim_id}")

            if claim_link.count() == 0:
                result["notes"] = f"Claim not found in search results"
                result["status"] = "ERROR"
                logger.error(f"Claim {claim_id} not found")
                return result

            claim_link.first.click()
            self.page.wait_for_load_state("networkidle", timeout=30000)

            # 3. RELOAD: Click Reload button
            logger.info(f"Clicking Reload for {claim_id}")
            reload_button = self.page.locator("button:has-text('Reload')")

            if reload_button.count() > 0:
                reload_button.click()
                self.page.wait_for_load_state("networkidle", timeout=30000)
                logger.info(f"Reload completed for {claim_id}")
            else:
                logger.warning(f"Reload button not found for {claim_id}")

            # 4. VERIFY: Confirm claim ID matches
            logger.info(f"Verifying claim ID {claim_id}")
            claim_id_display = self.page.locator(f"text={claim_id}")

            if claim_id_display.count() == 0:
                result["notes"] = "Claim ID mismatch after reload"
                result["status"] = "ERROR"
                logger.error(f"Claim ID verification failed for {claim_id}")
                return result

            logger.info(f"Claim ID verified: {claim_id}")

            # 5. SAVE: Click Save button
            logger.info(f"Saving claim {claim_id}")
            save_button = self.page.locator("button:has-text('Save')")

            if save_button.count() == 0:
                result["notes"] = "Save button not found"
                result["status"] = "ERROR"
                logger.error(f"Save button not found for {claim_id}")
                return result

            save_button.click()

            # Wait for save confirmation (networkidle or success message)
            try:
                self.page.wait_for_load_state("networkidle", timeout=30000)
                success_msg = self.page.locator("text=/Success|Saved|Updated/i")
                if success_msg.count() > 0:
                    result["status"] = "SUCCESS"
                    result["notes"] = success_msg.first.text_content()
                    logger.info(f"Claim {claim_id} saved successfully: {result['notes']}")
                else:
                    result["status"] = "SUCCESS"
                    result["notes"] = "Saved (no confirmation message)"
                    logger.info(f"Claim {claim_id} saved")
            except Exception as e:
                result["status"] = "ERROR"
                result["notes"] = f"Save timeout or error: {str(e)}"
                logger.error(f"Save failed for {claim_id}: {str(e)}")

        except Exception as e:
            result["status"] = "ERROR"
            result["notes"] = f"Processing error: {str(e)}"
            logger.error(f"Exception processing {claim_id}: {str(e)}")

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


def process_claims_file(input_file: str = "script_input.xlsx"):
    """Read claims from Excel and process each PENDING claim."""
    if not os.path.exists(input_file):
        logger.error(f"Input file not found: {input_file}")
        sys.exit(1)

    # Load environment
    load_dotenv()
    portal_url = os.getenv("PORTAL_URL")
    portal_user = os.getenv("PORTAL_USER")
    portal_pass = os.getenv("PORTAL_PASS")
    headless = os.getenv("HEADLESS", "False").lower() == "true"
    delay_secs = float(os.getenv("DELAY_SECS", "1.5"))

    if not all([portal_url, portal_user, portal_pass]):
        logger.error("Missing required environment variables: PORTAL_URL, PORTAL_USER, PORTAL_PASS")
        sys.exit(1)

    logger.info(f"Starting claim processing from {input_file}")
    logger.info(f"Portal: {portal_url} | Headless: {headless} | Delay: {delay_secs}s")

    # Initialize processor and browser
    processor = ClaimProcessor(portal_url, portal_user, portal_pass, headless, delay_secs)
    processor.start_browser()

    try:
        # Load workbook
        wb = load_workbook(input_file)
        ws = wb.active

        # Track processing
        processed = 0
        skipped = 0
        pending_count = 0

        # Count pending claims first
        for row_num in range(2, ws.max_row + 1):
            status_cell = ws.cell(row=row_num, column=4)
            if status_cell.value == "PENDING":
                pending_count += 1

        logger.info(f"Found {pending_count} PENDING claims to process")

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
            claim_num = claim_num_cell.value

            # Skip non-PENDING rows
            if current_status != "PENDING":
                skipped += 1
                continue

            logger.info(f"Processing claim {claim_num}/{processed + pending_count} ({processed + 1}/{pending_count}): {claim_id}")

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
            if delay_secs > 0:
                import time

                time.sleep(delay_secs)

        logger.info(f"Processing complete: {processed} processed, {skipped} skipped")

    except Exception as e:
        logger.error(f"Fatal error during processing: {str(e)}")
        sys.exit(1)

    finally:
        processor.close()


if __name__ == "__main__":
    input_file = sys.argv[1] if len(sys.argv) > 1 else "script_input.xlsx"
    process_claims_file(input_file)
