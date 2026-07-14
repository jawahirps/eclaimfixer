# DHA eClaim Batch Processing Tool

Automated batch processing tool for DHA eClaim portal claim submissions with resumable state tracking.

## Overview

This tool consists of two main scripts:

1. **build_input.py** - Generates the claim input Excel file (`script_input.xlsx`)
2. **process_claims.py** - Automates portal interactions and processes claims

## Setup

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Configure Environment

Copy `.env.example` to `.env` and fill in your portal credentials:

```bash
cp .env.example .env
```

Edit `.env`:
```env
PORTAL_URL=https://portal.dha.gov/eclaim
PORTAL_USER=your_username
PORTAL_PASS=your_password
HEADLESS=False    # Set to True for production/headless mode
DELAY_SECS=1.5    # Delay between claims (seconds)
```

## Usage

### Generate Input File

```bash
python3 build_input.py
```

Creates `script_input.xlsx` with 92 pre-populated claims:
- **INS012**: 19 claims (light blue)
- **TPA002**: 53 claims (light green)  
- **TPA003**: 20 claims (light yellow)

The file has:
- Frozen header row
- Bold column headers
- Auto-fitted columns
- Color-coded sources
- STATUS initialized to PENDING
- LOADED_AT, SAVED_AT, NOTES columns for tracking

### Process Claims

```bash
python3 process_claims.py [input_file.xlsx]
```

Default: `script_input.xlsx`

**Workflow per claim:**

1. **FETCH** - Search claim by ID on portal
2. **LOAD** - Open claim detail/edit form
3. **RELOAD** - Click Reload button; wait for networkidle
4. **VERIFY** - Confirm CLAIM_ID on screen matches expected
5. **SAVE** - Click Save; capture response

**Features:**
- Resumable: Updates Excel after each claim (safe interruption)
- Skips rows with STATUS != PENDING (allows resume)
- Logs all activity to `run_YYYYMMDD.log`
- Timestamps LOADED_AT, SAVED_AT for tracking
- Captures portal responses in NOTES column

## Excel File Format

| Column | Header    | Type       | Description                  |
|--------|-----------|------------|------------------------------|
| A      | #         | Integer    | Row number (1-92)             |
| B      | SOURCE    | String     | INS012 / TPA002 / TPA003       |
| C      | CLAIM_ID  | String     | Raw eClaim ID                 |
| D      | STATUS    | String     | PENDING / SUCCESS / ERROR     |
| E      | LOADED_AT | Timestamp  | When claim was fetched        |
| F      | SAVED_AT  | Timestamp  | When claim was saved          |
| G      | NOTES     | String     | Portal response or error msg  |

## Logging

Activity is logged to `run_YYYYMMDD.log` in the current directory:

```
2026-07-14 10:25:30,123 - INFO - Starting claim processing from script_input.xlsx
2026-07-14 10:25:30,456 - INFO - Found 92 PENDING claims to process
2026-07-14 10:25:35,789 - INFO - Processing claim 1/92: ALNDEIRA_010626_425600.41580
...
```

## Resume Processing

If processing is interrupted:

1. The Excel file is automatically saved after each claim
2. Re-run: `python3 process_claims.py`
3. Rows with STATUS=SUCCESS are skipped
4. Processing resumes from the next PENDING row

## Troubleshooting

### Browser not found

Playwright needs to download Chromium on first run:

```bash
playwright install chromium
```

### Timeout errors

Increase `DELAY_SECS` in `.env` if claims aren't loading:

```env
DELAY_SECS=3.0
```

### Portal login fails

- Verify credentials in `.env`
- Check if portal URL is correct
- Ensure portal is accessible

### Element not found errors

Portal selectors may need updating if UI changes:
- Search field: `input[placeholder*='Claim ID']`
- Reload button: `button:has-text('Reload')`
- Save button: `button:has-text('Save')`

Update selectors in `process_claims.py` as needed.

## Stack

- **Python** 3.10+
- **openpyxl** - Excel file handling
- **playwright** - Browser automation (Chromium)
- **python-dotenv** - Environment configuration

## Files

- `build_input.py` - Input file generator
- `process_claims.py` - Automation runner
- `requirements.txt` - Python dependencies
- `.env.example` - Configuration template
- `script_input.xlsx` - Generated input file (after running build_input.py)
- `run_*.log` - Execution logs
