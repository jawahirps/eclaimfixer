# DHA eClaim Batch Tool — Rebuilt Features

**Version 2.0** - Enhanced control, flexible input, and validation

## What's New

### 1. **Flexible Input Management** (build_input.py)

Generate Excel files with **custom claims, output paths, and initial status**.

```bash
# Default: Uses built-in 92 claims
build_input.py

# Custom output file
build_input.py --output my_batch.xlsx

# Load claims from JSON
build_input.py --claims custom_claims.json --output batch.xlsx

# Initialize with different status
build_input.py --status SUCCESS  # e.g., for recovery scenarios
```

**JSON format for custom claims:**
```json
{
  "INS012": ["CLAIM_ID_1", "CLAIM_ID_2"],
  "TPA002": ["CLAIM_ID_3"],
  "CUSTOM_SOURCE": ["CLAIM_ID_4"]
}
```

### 2. **Precheck Validation** (process_claims.py)

**Analyze and validate claims BEFORE processing** — catches issues early.

```bash
# Just validate, don't process
process_claims.py claims.xlsx --only-precheck

# Output:
# - Total claims, pending/success/error counts
# - Claims by source (color-coded)
# - Issues found (missing IDs, invalid data)
# - No browser automation required
```

**What precheck detects:**
- ✓ File format and structure
- ✓ Missing claim IDs or sources
- ✓ Status distribution (PENDING/SUCCESS/ERROR)
- ✓ Source breakdown with color coding

### 3. **Dry-Run Mode** (process_claims.py)

**Simulate processing without browser automation** — test workflows safely.

```bash
# Simulate all 92 claims in seconds
process_claims.py --dry-run

# Output:
# [DRY RUN] Browser launch skipped
# [1/92] Claim 1 (INS012)
# [DRY RUN] Would process: ALNDEIRA_010626_425600.41580
# [2/92] Claim 2 (INS012)
# ...
```

**Use cases:**
- Verify input file is correct
- Test resume functionality
- Check logging output format
- Validate workflow without portal access

### 4. **Variable Input Files**

Process **any Excel file**, not just `script_input.xlsx`.

```bash
# Process multiple batches sequentially
process_claims.py batch1.xlsx
process_claims.py batch2.xlsx
process_claims.py batch3.xlsx

# Or in production:
for file in batches/*.xlsx; do
  python3 process_claims.py "$file"
done
```

State tracking works independently for each file.

### 5. **Configuration Profiles** (config_manager.py)

**Manage multiple portal environments** — prod, staging, dev.

```bash
# Create profiles
config_manager.py create prod \
  --url "https://portal.dha.gov/eclaim" \
  --user alice \
  --pass secret123 \
  --headless

config_manager.py create staging \
  --url "https://staging.dha.gov" \
  --user bob \
  --pass testpass

# Switch environments
config_manager.py load prod      # Activate prod
config_manager.py load staging   # Switch to staging

# View all profiles
config_manager.py list

# View active config (passwords masked)
config_manager.py show

# Remove a profile
config_manager.py delete staging
```

**Profile storage:**
- Stored in `.env.profiles/` directory
- Loaded into `.env` when activated
- Supports all environment variables:
  - PORTAL_URL (variable per environment!)
  - PORTAL_USER
  - PORTAL_PASS
  - HEADLESS mode
  - DELAY_SECS per claim

### 6. **Enhanced CLI with Argument Parsing**

All scripts now have proper help and arguments.

```bash
# process_claims.py
process_claims.py --help
process_claims.py input.xlsx --dry-run --verbose
process_claims.py claims.xlsx --skip-precheck  # Force processing
process_claims.py batch.xlsx -v                # DEBUG logging

# build_input.py
build_input.py --help
build_input.py -o output.xlsx --status PENDING -v

# config_manager.py
config_manager.py --help
config_manager.py create --help
```

### 7. **Improved Logging**

Better visibility into processing workflow.

```
2026-07-14 14:29:40 - INFO - Precheck: analyzing claims...
2026-07-14 14:29:40 - INFO -   Total: 92 claims
2026-07-14 14:29:40 - INFO -   Pending: 92 | Success: 0 | Error: 0
2026-07-14 14:29:40 - INFO -     INS012: 19
2026-07-14 14:29:40 - INFO -     TPA002: 53
2026-07-14 14:29:40 - INFO -     TPA003: 20
2026-07-14 14:29:40 - INFO - Processing 92 PENDING claims...
2026-07-14 14:29:40 - INFO - [1/92] Claim 1 (INS012)
2026-07-14 14:29:40 - INFO - [1/5] FETCH: ALNDEIRA_010626_425600.41580
2026-07-14 14:29:41 - INFO - [2/5] LOAD: ALNDEIRA_010626_425600.41580
2026-07-14 14:29:42 - INFO - [3/5] RELOAD: ALNDEIRA_010626_425600.41580
2026-07-14 14:29:42 - INFO -   ✓ Reloaded
2026-07-14 14:29:43 - INFO - [4/5] VERIFY: ALNDEIRA_010626_425600.41580
2026-07-14 14:29:43 - INFO -   ✓ Verified
2026-07-14 14:29:44 - INFO - [5/5] SAVE: ALNDEIRA_010626_425600.41580
2026-07-14 14:29:44 - INFO -   ✓ Saved
```

Step numbers [1/5] show workflow progress.
Status indicators (✓, ✗, !) show immediate results.

### 8. **Resumable Processing** (Preserved + Enhanced)

Processing **saves after each claim**, enabling:
- Safe interruption
- Resume from PENDING rows only
- Skip SUCCESS/ERROR rows on resume

```bash
# Start processing
python3 process_claims.py batch.xlsx
# Press Ctrl+C after claim 10...

# Later, resume
python3 process_claims.py batch.xlsx
# Rows 1-10 skipped (already processed)
# Processing resumes at claim 11
```

## Workflow Examples

### Example 1: Validate Before Processing

```bash
# 1. Generate claims
build_input.py --output claims_batch1.xlsx

# 2. Validate before portal access
process_claims.py claims_batch1.xlsx --only-precheck

# Output:
# Precheck: analyzing claims...
#   Total: 92 claims
#   Pending: 92 | Success: 0 | Error: 0
#   INS012: 19, TPA002: 53, TPA003: 20
# No issues found.

# 3. Process
process_claims.py claims_batch1.xlsx
```

### Example 2: Multi-Environment Deployment

```bash
# Setup environments once
config_manager.py create prod --url https://prod.portal.com --user alice --pass secret --headless
config_manager.py create dev --url https://dev.portal.com --user bob --pass testpass

# Deploy to dev (test)
config_manager.py load dev
python3 process_claims.py claims.xlsx

# Deploy to prod (production)
config_manager.py load prod
python3 process_claims.py claims.xlsx
```

### Example 3: Custom Claims Batch

```bash
# Create custom claims file
cat > my_claims.json << 'EOF'
{
  "VENDOR_A": ["ID_001", "ID_002", "ID_003"],
  "VENDOR_B": ["ID_004", "ID_005"]
}
EOF

# Generate with custom claims
build_input.py --claims my_claims.json --output vendor_batch.xlsx

# Process
process_claims.py vendor_batch.xlsx
```

### Example 4: Dry-Run for Testing

```bash
# Test workflow without browser/portal
process_claims.py claims.xlsx --dry-run

# Review logs
tail -f run_20260714.log

# Verify Excel updates (check script_input.xlsx)
# All rows should show STATUS=SUCCESS after dry-run

# Then run real processing
process_claims.py claims.xlsx
```

## Architecture Changes

### Before (v1)
- ❌ Hardcoded `script_input.xlsx`
- ❌ No validation before processing
- ❌ No simulation mode
- ❌ Single portal URL
- ❌ Basic logging

### After (v2)
- ✅ Dynamic input files
- ✅ Precheck validation (mandatory by default)
- ✅ Dry-run simulation
- ✅ Multiple portal profiles
- ✅ Enhanced logging with progress
- ✅ Better error messages
- ✅ Flexible claims sources

## File Structure

```
eclaimfixer/
├── build_input.py           # Generate Excel with custom claims
├── process_claims.py        # Main automation with precheck & dry-run
├── config_manager.py        # Manage portal profiles (NEW)
├── requirements.txt         # Dependencies
├── .env.example            # Config template
├── .gitignore              # Excludes .env, logs, profiles (NEW)
├── script_input.xlsx       # Default input file (generated)
├── .env.profiles/          # Config profiles directory (generated)
│   ├── prod.env
│   ├── staging.env
│   └── dev.env
├── run_YYYYMMDD.log        # Execution logs (generated)
└── README_BATCH_TOOL.md    # Original documentation
```

## Control Flags Summary

| Flag | Script | Purpose |
|------|--------|---------|
| `--output` | build_input.py | Custom output filename |
| `--claims` | build_input.py | Load claims from JSON |
| `--status` | build_input.py | Initial claim status |
| `--dry-run` | process_claims.py | Simulate without browser |
| `--only-precheck` | process_claims.py | Validation only |
| `--skip-precheck` | process_claims.py | Force process (risky) |
| `-v, --verbose` | Both | DEBUG level logging |

## Benefits

✓ **Better Control** — Validate, simulate, and process independently
✓ **Flexible Input** — Support any claim list, any output file
✓ **Safe** — Precheck catches issues before portal access
✓ **Multi-Tenant** — Manage prod/staging/dev with profiles
✓ **Observable** — Clear logging with step-by-step progress
✓ **Resumable** — Pick up where you left off
✓ **Testable** — Dry-run mode requires zero portal access

## Migration from v1

Existing `script_input.xlsx` files work unchanged:

```bash
# v1 workflow (still works)
python3 process_claims.py

# v2 workflow (recommended)
python3 process_claims.py --only-precheck
python3 process_claims.py --dry-run
python3 process_claims.py
```

No breaking changes — all v1 features preserved and enhanced.
