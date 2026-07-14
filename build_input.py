#!/usr/bin/env python3
"""Generate Excel input file with DHA eClaim data - supports flexible sources and claims."""

import argparse
import json
from pathlib import Path
from typing import Dict, List

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# Default claim data organized by source
DEFAULT_CLAIMS_BY_SOURCE = {
    "INS012": [
        "ALNDEIRA_010626_425600.41580", "ALNDEIRA_010626_425626.21883",
        "ALNDEIRA_050626_426206.85135", "ALNDEIRA_060626_426322.16638",
        "ALNDEIRA_100626_426914.33653", "ALNDEIRA_100626_426970.63209",
        "ALNDEIRA_120626_427324.99754", "ALNDEIRA_130626_427484.81226",
        "ALNDEIRA_140626_427523.34474", "ALNDEIRA_140626_427541.7552",
        "ALNDEIRA_140626_427591.15787", "ALNDEIRA_140626_427664.35042",
        "ALNDEIRA_150626_427799.67319", "ALNDEIRA_160626_427868.48334",
        "ALNDEIRA_170626_428172.59254", "ALNDEIRA_200626_428548.37655",
        "ALNDEIRA_210626_428687.58976", "ALNDEIRA_250626_429340.2214",
        "ALNDEIRA_250626_429408.41179",
    ],
    "TPA002": [
        "ALNDEIRA_020626_425726.5781", "ALNDEIRA_020626_425740.74896",
        "ALNDEIRA_030626_425911.13497", "ALNDEIRA_030626_425961.12754",
        "ALNDEIRA_040626_425995.70726", "ALNDEIRA_050626_426164.5482",
        "ALNDEIRA_060626_426302.13300", "ALNDEIRA_060626_426303.29664",
        "ALNDEIRA_060626_426403.90237", "ALNDEIRA_070626_426443.31433",
        "ALNDEIRA_070626_426516.67002", "ALNDEIRA_080626_426571.12605",
        "ALNDEIRA_080626_426575.31277", "ALNDEIRA_080626_426580.8133",
        "ALNDEIRA_110626_427201.19163", "ALNDEIRA_120626_427225.84063",
        "ALNDEIRA_140626_427528.77316", "ALNDEIRA_140626_427532.74797",
        "ALNDEIRA_140626_427534.87116", "ALNDEIRA_140626_427628.21253",
        "ALNDEIRA_140626_427648.68141", "ALNDEIRA_150626_427684.43338",
        "ALNDEIRA_150626_427693.19163", "ALNDEIRA_150626_427710.89741",
        "ALNDEIRA_150626_427712.29071", "ALNDEIRA_150626_427813.89741",
        "ALNDEIRA_150626_427850.7465", "ALNDEIRA_170626_428027.1627",
        "ALNDEIRA_170626_428037.41880", "ALNDEIRA_170626_428038.91292",
        "ALNDEIRA_170626_428046.14813", "ALNDEIRA_180626_428302.18551",
        "ALNDEIRA_190626_428490.45945", "ALNDEIRA_200626_428600.1666",
        "ALNDEIRA_210626_428689.56526", "ALNDEIRA_220626_428824.90431",
        "ALNDEIRA_220626_428846.66771", "ALNDEIRA_220626_428876.1825",
        "ALNDEIRA_220626_428889.70595", "ALNDEIRA_220626_428962.96223",
        "ALNDEIRA_230626_428991.22085", "ALNDEIRA_230626_429036.6791",
        "ALNDEIRA_240626_429246.95770", "ALNDEIRA_250626_429307.88423",
        "ALNDEIRA_260626_429468.80436", "ALNDEIRA_260626_429487.29731",
        "ALNDEIRA_270626_429638.93286", "ALNDEIRA_270626_429692.9265",
        "ALNDEIRA_280626_429754.31194", "ALNDEIRA_280626_429757.63800",
        "ALNDEIRA_280626_429820.67267", "ALNDEIRA_290626_429914.30507",
        "ALNDEIRA_300626_430168.90636",
    ],
    "TPA003": [
        "ALNDEIRA_010626_425591.57780", "ALNDEIRA_020626_425722.28146",
        "ALNDEIRA_030626_425878.89535", "ALNDEIRA_080626_426569.17283",
        "ALNDEIRA_080626_426584.40294", "ALNDEIRA_120626_427251.9285",
        "ALNDEIRA_130626_427379.38272", "ALNDEIRA_150626_427704.27680",
        "ALNDEIRA_160626_427972.56839", "ALNDEIRA_160626_428021.88609",
        "ALNDEIRA_170626_428146.92239", "ALNDEIRA_180626_428198.99260",
        "ALNDEIRA_180626_428287.84821", "ALNDEIRA_190626_428364.82181",
        "ALNDEIRA_200626_428529.21321", "ALNDEIRA_270626_429653.16296",
        "ALNDEIRA_270626_429740.4706", "ALNDEIRA_280626_429761.83596",
        "ALNDEIRA_280626_429762.75225", "ALNDEIRA_290626_430028.100212",
    ],
}

# Color fills for sources
DEFAULT_SOURCE_COLORS = {
    "INS012": "ADD8E6",  # light blue
    "TPA002": "90EE90",  # light green
    "TPA003": "FFFFE0",  # light yellow
}


def load_claims_from_json(json_file: str) -> Dict[str, List[str]]:
    """Load claims from JSON file.

    Expected format:
    {
        "INS012": ["claim1", "claim2", ...],
        "TPA002": ["claim3", ...],
        ...
    }
    """
    with open(json_file, "r") as f:
        return json.load(f)


def create_input_file(
    filename: str = "script_input.xlsx",
    claims_by_source: Dict[str, List[str]] = None,
    source_colors: Dict[str, str] = None,
    status: str = "PENDING",
):
    """Create the input Excel file with claims data.

    Args:
        filename: Output Excel file path
        claims_by_source: Dictionary of source -> [claim_ids] (uses defaults if None)
        source_colors: Dictionary of source -> hex_color (uses defaults if None)
        status: Initial status for all claims (default: PENDING)
    """
    if claims_by_source is None:
        claims_by_source = DEFAULT_CLAIMS_BY_SOURCE

    if source_colors is None:
        source_colors = DEFAULT_SOURCE_COLORS

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Claims"

    # Add headers
    headers = ["#", "SOURCE", "CLAIM_ID", "STATUS", "LOADED_AT", "SAVED_AT", "NOTES"]
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col)
        cell.value = header
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Freeze header row
    ws.freeze_panes = "A2"

    # Add data rows
    row_num = 2
    claim_index = 1
    total_claims = sum(len(claims) for claims in claims_by_source.values())

    for source in sorted(claims_by_source.keys()):
        color_hex = source_colors.get(source, "FFFFFF")
        color_fill = PatternFill(start_color=color_hex, fill_type="solid")

        for claim_id in claims_by_source[source]:
            ws.cell(row=row_num, column=1).value = claim_index
            ws.cell(row=row_num, column=2).value = source
            ws.cell(row=row_num, column=3).value = claim_id
            ws.cell(row=row_num, column=4).value = status

            # Apply source color to SOURCE and CLAIM_ID columns
            for col in [2, 3]:
                ws.cell(row=row_num, column=col).fill = color_fill

            # Center align the # column
            ws.cell(row=row_num, column=1).alignment = Alignment(horizontal="center")

            row_num += 1
            claim_index += 1

    # Auto-fit columns
    column_widths = {
        "A": 5,
        "B": 12,
        "C": 30,
        "D": 12,
        "E": 20,
        "F": 20,
        "G": 40,
    }
    for col_letter, width in column_widths.items():
        ws.column_dimensions[col_letter].width = width

    wb.save(filename)

    print(f"✓ Created {filename} with {total_claims} claims")
    for source in sorted(claims_by_source.keys()):
        count = len(claims_by_source[source])
        print(f"  - {source}: {count} claims")


def main():
    """CLI entry point for generating input files."""
    parser = argparse.ArgumentParser(
        description="Generate Excel input file for DHA eClaim batch processing",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s                                    # Create script_input.xlsx (default)
  %(prog)s --output claims_batch.xlsx         # Custom output file
  %(prog)s --claims claims.json               # Load claims from JSON
  %(prog)s --status SUCCESS                   # Initialize with SUCCESS status
        """,
    )

    parser.add_argument("-o", "--output", default="script_input.xlsx", help="Output Excel file (default: script_input.xlsx)")
    parser.add_argument("-c", "--claims", help="JSON file with claims data (overrides defaults)")
    parser.add_argument("-s", "--status", default="PENDING", help="Initial status for all claims (default: PENDING)")
    parser.add_argument("-v", "--verbose", action="store_true", help="Verbose output")

    args = parser.parse_args()

    # Load claims
    claims_by_source = DEFAULT_CLAIMS_BY_SOURCE
    if args.claims:
        if not Path(args.claims).exists():
            print(f"Error: Claims file not found: {args.claims}")
            return 1
        if args.verbose:
            print(f"Loading claims from {args.claims}...")
        claims_by_source = load_claims_from_json(args.claims)

    if args.verbose:
        print(f"Output file: {args.output}")
        print(f"Initial status: {args.status}")
        print(f"Sources: {list(claims_by_source.keys())}")

    # Create file
    create_input_file(
        filename=args.output,
        claims_by_source=claims_by_source,
        status=args.status,
    )

    return 0


if __name__ == "__main__":
    import sys

    sys.exit(main())
