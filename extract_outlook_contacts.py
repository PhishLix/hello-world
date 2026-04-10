#!/usr/bin/env python3
"""
Extract contact information from Microsoft Outlook .pst files and export to Excel.

Parses emails in a .pst file, extracts structured contact data (names, addresses,
phone numbers, emails, monetary amounts), and writes the results to an Excel
spreadsheet.

Usage:
    python extract_outlook_contacts.py input.pst -o output.xlsx
    python extract_outlook_contacts.py input.pst --demo   # generate demo output without a real .pst
"""

import argparse
import os
import re
import sys
import uuid
from datetime import datetime

try:
    import pypff
except ImportError:
    pypff = None

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
except ImportError:
    sys.exit(
        "ERROR: openpyxl is not installed. Run:\n"
        "  pip install openpyxl\n"
    )

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

EXCEL_COLUMNS = [
    "Source ID",
    "First Name",
    "Last Name",
    "City",
    "State",
    "Phone",
    "Email",
    "Amount of Money Lost",
]

# US state abbreviations for validation
US_STATES = {
    "AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA",
    "HI", "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD",
    "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ",
    "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC",
    "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY",
    "DC",
}

US_STATE_NAMES = {
    "alabama": "AL", "alaska": "AK", "arizona": "AZ", "arkansas": "AR",
    "california": "CA", "colorado": "CO", "connecticut": "CT",
    "delaware": "DE", "florida": "FL", "georgia": "GA", "hawaii": "HI",
    "idaho": "ID", "illinois": "IL", "indiana": "IN", "iowa": "IA",
    "kansas": "KS", "kentucky": "KY", "louisiana": "LA", "maine": "ME",
    "maryland": "MD", "massachusetts": "MA", "michigan": "MI",
    "minnesota": "MN", "mississippi": "MS", "missouri": "MO",
    "montana": "MT", "nebraska": "NE", "nevada": "NV",
    "new hampshire": "NH", "new jersey": "NJ", "new mexico": "NM",
    "new york": "NY", "north carolina": "NC", "north dakota": "ND",
    "ohio": "OH", "oklahoma": "OK", "oregon": "OR", "pennsylvania": "PA",
    "rhode island": "RI", "south carolina": "SC", "south dakota": "SD",
    "tennessee": "TN", "texas": "TX", "utah": "UT", "vermont": "VT",
    "virginia": "VA", "washington": "WA", "west virginia": "WV",
    "wisconsin": "WI", "wyoming": "WY", "district of columbia": "DC",
}

# ---------------------------------------------------------------------------
# Regex helpers
# ---------------------------------------------------------------------------

# Matches common key: value patterns found in structured emails
# e.g. "First Name: John" or "First Name - John" or "First Name = John"
def _kv_pattern(key_alts: str) -> re.Pattern:
    """Build a regex that matches key-value lines for the given key alternatives."""
    return re.compile(
        rf"(?:^|\n)\s*(?:{key_alts})\s*[:=\-]\s*(.+)",
        re.IGNORECASE,
    )


RE_FIRST_NAME = _kv_pattern(
    r"first\s*name|given\s*name|fname|first"
)
RE_LAST_NAME = _kv_pattern(
    r"last\s*name|surname|family\s*name|lname|last"
)
RE_CITY = _kv_pattern(r"city|town|municipality")
RE_STATE = _kv_pattern(r"state|province|st")
RE_PHONE = _kv_pattern(
    r"phone|telephone|tel|cell|mobile|ph|phone\s*number|contact\s*number"
)
RE_EMAIL_FIELD = _kv_pattern(
    r"e-?mail|email\s*address|contact\s*email"
)
RE_AMOUNT = _kv_pattern(
    r"amount|loss|money\s*lost|amount\s*lost|amount\s*of\s*money\s*lost|"
    r"total\s*loss|financial\s*loss|damages|amount\s*of\s*loss"
)
RE_SOURCE_ID = _kv_pattern(
    r"source\s*id|case\s*id|reference|ref|id|case\s*number|case\s*no|ticket"
)
RE_FULL_NAME = _kv_pattern(r"name|full\s*name|contact\s*name|complainant")

# Standalone patterns (not key-value) used as fallbacks
RE_EMAIL_ADDR = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
RE_PHONE_STANDALONE = re.compile(
    r"(?<!\d)"
    r"(?:\+?1[\s.\-]?)?"
    r"\(?\d{3}\)?[\s.\-]?\d{3}[\s.\-]?\d{4}"
    r"(?!\d)"
)
RE_MONEY = re.compile(r"\$\s?[\d,]+(?:\.\d{2})?")


# ---------------------------------------------------------------------------
# Extraction logic
# ---------------------------------------------------------------------------

def _first_match(pattern: re.Pattern, text: str) -> str:
    """Return the first stripped match group, or empty string."""
    m = pattern.search(text)
    return m.group(1).strip() if m else ""


def _clean_phone(raw: str) -> str:
    """Normalize a phone string to (XXX) XXX-XXXX format when possible."""
    digits = re.sub(r"\D", "", raw)
    if digits.startswith("1") and len(digits) == 11:
        digits = digits[1:]
    if len(digits) == 10:
        return f"({digits[:3]}) {digits[3:6]}-{digits[6:]}"
    return raw.strip()


def _clean_money(raw: str) -> str:
    """Normalize a money string to $X,XXX.XX format."""
    raw = raw.strip()
    if not raw:
        return ""
    # Remove $ and spaces, keep digits, commas, dots
    cleaned = raw.replace("$", "").replace(" ", "").replace(",", "")
    try:
        amount = float(cleaned)
        return f"${amount:,.2f}"
    except ValueError:
        return raw


def _normalize_state(raw: str) -> str:
    """Convert state to two-letter abbreviation if recognized."""
    raw = raw.strip()
    upper = raw.upper()
    if upper in US_STATES:
        return upper
    lower = raw.lower()
    if lower in US_STATE_NAMES:
        return US_STATE_NAMES[lower]
    return raw


def _split_full_name(full_name: str) -> tuple:
    """Split a full name into (first, last). Handles 'Last, First' format."""
    full_name = full_name.strip()
    if not full_name:
        return ("", "")
    if "," in full_name:
        parts = [p.strip() for p in full_name.split(",", 1)]
        return (parts[1], parts[0]) if len(parts) == 2 else (parts[0], "")
    parts = full_name.split()
    if len(parts) == 1:
        return (parts[0], "")
    return (parts[0], " ".join(parts[1:]))


def extract_contact_from_text(text: str, message_id: str = "") -> dict:
    """
    Parse an email body and extract contact fields.

    Uses key-value pattern matching first (e.g. 'First Name: John'),
    then falls back to standalone regex patterns for email, phone, and amounts.
    """
    record = {col: "" for col in EXCEL_COLUMNS}

    # --- Source ID ---
    source_id = _first_match(RE_SOURCE_ID, text)
    if source_id:
        record["Source ID"] = source_id
    elif message_id:
        record["Source ID"] = message_id
    else:
        record["Source ID"] = str(uuid.uuid4())[:8].upper()

    # --- Names ---
    first = _first_match(RE_FIRST_NAME, text)
    last = _first_match(RE_LAST_NAME, text)

    # If no explicit first/last, try full name field
    if not first and not last:
        full = _first_match(RE_FULL_NAME, text)
        first, last = _split_full_name(full)

    record["First Name"] = first
    record["Last Name"] = last

    # --- City / State ---
    record["City"] = _first_match(RE_CITY, text)
    state_raw = _first_match(RE_STATE, text)
    record["State"] = _normalize_state(state_raw) if state_raw else ""

    # --- Phone ---
    phone = _first_match(RE_PHONE, text)
    if not phone:
        m = RE_PHONE_STANDALONE.search(text)
        if m:
            phone = m.group(0)
    record["Phone"] = _clean_phone(phone) if phone else ""

    # --- Email ---
    email = _first_match(RE_EMAIL_FIELD, text)
    if not email:
        m = RE_EMAIL_ADDR.search(text)
        if m:
            email = m.group(0)
    record["Email"] = email

    # --- Amount of Money Lost ---
    amount = _first_match(RE_AMOUNT, text)
    if not amount:
        m = RE_MONEY.search(text)
        if m:
            amount = m.group(0)
    record["Amount of Money Lost"] = _clean_money(amount) if amount else ""

    return record


# ---------------------------------------------------------------------------
# PST traversal
# ---------------------------------------------------------------------------

def _get_message_body(message) -> str:
    """Extract plain-text or HTML body from a pypff message object."""
    body = ""
    try:
        body = message.plain_text_body
        if body and isinstance(body, bytes):
            body = body.decode("utf-8", errors="replace")
    except Exception:
        pass

    if not body:
        try:
            body = message.html_body
            if body and isinstance(body, bytes):
                body = body.decode("utf-8", errors="replace")
            if body:
                # Rough HTML tag stripping for content extraction
                body = re.sub(r"<br\s*/?>", "\n", body, flags=re.IGNORECASE)
                body = re.sub(r"<[^>]+>", " ", body)
                body = re.sub(r"&nbsp;", " ", body)
                body = re.sub(r"&amp;", "&", body)
                body = re.sub(r"&lt;", "<", body)
                body = re.sub(r"&gt;", ">", body)
        except Exception:
            pass

    return body or ""


def _get_message_id(message) -> str:
    """Try to get a meaningful identifier from the message."""
    try:
        subject = message.subject
        if subject:
            return subject[:50]
    except Exception:
        pass
    return ""


def _process_folder(folder, records: list, stats: dict):
    """Recursively process a pypff folder and its subfolders."""
    # Process messages in this folder
    try:
        num_messages = folder.number_of_sub_messages
    except Exception:
        num_messages = 0

    for i in range(num_messages):
        try:
            message = folder.get_sub_message(i)
            stats["emails_scanned"] += 1

            body = _get_message_body(message)
            if not body.strip():
                stats["empty_bodies"] += 1
                continue

            msg_id = _get_message_id(message)
            record = extract_contact_from_text(body, msg_id)

            # Only include records that have at least a name or email
            has_data = any(
                record[f]
                for f in ["First Name", "Last Name", "Email", "Phone"]
            )
            if has_data:
                records.append(record)
                stats["contacts_extracted"] += 1
            else:
                stats["no_contact_data"] += 1

        except Exception as exc:
            stats["errors"] += 1
            print(f"  Warning: Could not process message {i}: {exc}", file=sys.stderr)

    # Recurse into subfolders
    try:
        num_folders = folder.number_of_sub_folders
    except Exception:
        num_folders = 0

    for i in range(num_folders):
        try:
            subfolder = folder.get_sub_folder(i)
            folder_name = ""
            try:
                folder_name = subfolder.name
            except Exception:
                pass
            if folder_name:
                print(f"  Scanning folder: {folder_name}")
            _process_folder(subfolder, records, stats)
        except Exception as exc:
            stats["errors"] += 1
            print(f"  Warning: Could not access subfolder {i}: {exc}", file=sys.stderr)


def extract_from_pst(pst_path: str) -> list:
    """
    Open a .pst file and extract contact records from all emails.

    Returns a list of dicts, one per extracted contact.
    """
    if pypff is None:
        sys.exit(
            "ERROR: pypff (libpff-python) is not installed. Run:\n"
            "  pip install libpff-python\n"
        )

    if not os.path.isfile(pst_path):
        sys.exit(f"ERROR: File not found: {pst_path}")

    print(f"Opening PST file: {pst_path}")
    pst = pypff.file()
    pst.open(pst_path)

    records = []
    stats = {
        "emails_scanned": 0,
        "contacts_extracted": 0,
        "empty_bodies": 0,
        "no_contact_data": 0,
        "errors": 0,
    }

    root = pst.get_root_folder()
    _process_folder(root, records, stats)

    pst.close()

    print(f"\nExtraction complete:")
    print(f"  Emails scanned:       {stats['emails_scanned']}")
    print(f"  Contacts extracted:   {stats['contacts_extracted']}")
    print(f"  Empty email bodies:   {stats['empty_bodies']}")
    print(f"  No contact data:      {stats['no_contact_data']}")
    print(f"  Errors:               {stats['errors']}")

    return records


# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------

def write_to_excel(records: list, output_path: str):
    """Write extracted contact records to a formatted Excel workbook."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Extracted Contacts"

    # --- Header styling ---
    header_font = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    # Write header row
    for col_idx, col_name in enumerate(EXCEL_COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border

    # --- Data styling ---
    data_font = Font(name="Calibri", size=11)
    data_alignment = Alignment(vertical="center")
    money_format = '#,##0.00'

    # Write data rows
    for row_idx, record in enumerate(records, start=2):
        for col_idx, col_name in enumerate(EXCEL_COLUMNS, start=1):
            value = record.get(col_name, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = data_font
            cell.alignment = data_alignment
            cell.border = thin_border

            # Format the money column as currency if it's numeric
            if col_name == "Amount of Money Lost" and value:
                cleaned = value.replace("$", "").replace(",", "")
                try:
                    cell.value = float(cleaned)
                    cell.number_format = '"$"#,##0.00'
                except ValueError:
                    pass

    # --- Auto-fit column widths ---
    column_widths = {
        "Source ID": 15,
        "First Name": 18,
        "Last Name": 18,
        "City": 18,
        "State": 10,
        "Phone": 18,
        "Email": 30,
        "Amount of Money Lost": 22,
    }
    for col_idx, col_name in enumerate(EXCEL_COLUMNS, start=1):
        col_letter = ws.cell(row=1, column=col_idx).column_letter
        ws.column_dimensions[col_letter].width = column_widths.get(col_name, 15)

    # Freeze the header row
    ws.freeze_panes = "A2"

    # Auto-filter on header
    ws.auto_filter.ref = ws.dimensions

    wb.save(output_path)
    print(f"\nExcel file saved: {output_path}")
    print(f"  Total rows: {len(records)}")


# ---------------------------------------------------------------------------
# Demo mode — generates sample output without needing a real .pst file
# ---------------------------------------------------------------------------

DEMO_EMAILS = [
    """
    Case ID: RPT-2024-0451
    First Name: Maria
    Last Name: Gonzalez
    City: Newark
    State: New Jersey
    Phone: (973) 555-0142
    Email: maria.gonzalez@example.com
    Amount of Money Lost: $4,500.00

    I was contacted by someone claiming to be from my bank. They convinced me
    to wire money to a fraudulent account.
    """,
    """
    Source ID: RPT-2024-0452
    Name: James O'Brien
    City: Jersey City
    State: NJ
    Telephone: 201-555-0198
    E-mail: jobrien22@example.com
    Total Loss: $12,750.00

    Received a phishing email that appeared to come from my employer's HR
    department requesting updated direct deposit information.
    """,
    """
    Reference: RPT-2024-0453
    First Name: Aisha
    Last Name: Patel
    City: Trenton
    State: NJ
    Cell: 609.555.0173
    Email: aisha.patel@example.com
    Amount Lost: $850.00

    Purchased items from a fraudulent online storefront. Products never arrived
    and the website has since disappeared.
    """,
    """
    Case Number: RPT-2024-0454
    First Name: Robert
    Last Name: Chen
    City: Princeton
    State: New Jersey
    Phone Number: (609) 555-0201
    Contact Email: r.chen@example.com
    Financial Loss: $23,000.00

    Invested in what turned out to be a cryptocurrency scam promoted via
    social media advertisements.
    """,
    """
    Ticket: RPT-2024-0455
    Complainant: Williams, Sandra
    City: Atlantic City
    State: NJ
    Mobile: 609-555-0187
    Email Address: sandra.w@example.com
    Amount of Money Lost: $1,200.00

    Romance scam conducted over a dating application over the course of
    three months.
    """,
    """
    ID: RPT-2024-0456
    Full Name: David Kim
    City: Hoboken
    State: NJ
    Tel: +1 201 555 0164
    Email: dkim@example.com
    Damages: $6,300.00

    Received a phone call from someone claiming to be the IRS demanding
    immediate payment via gift cards.
    """,
]


def generate_demo_records() -> list:
    """Generate sample contact records from built-in demo email text."""
    print("Running in DEMO mode with sample email data...\n")
    records = []
    for i, body in enumerate(DEMO_EMAILS):
        record = extract_contact_from_text(body, f"demo-msg-{i+1}")
        records.append(record)
    print(f"  Demo contacts generated: {len(records)}")
    return records


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Extract contacts from Outlook .pst emails and export to Excel.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s inbox.pst
  %(prog)s inbox.pst -o contacts.xlsx
  %(prog)s --demo
  %(prog)s --demo -o sample_output.xlsx
        """,
    )
    parser.add_argument(
        "pst_file",
        nargs="?",
        help="Path to the Outlook .pst file to process",
    )
    parser.add_argument(
        "-o", "--output",
        default=None,
        help="Output Excel file path (default: extracted_contacts.xlsx)",
    )
    parser.add_argument(
        "--demo",
        action="store_true",
        help="Run with built-in sample data (no .pst file needed)",
    )

    args = parser.parse_args()

    if not args.demo and not args.pst_file:
        parser.error("Please provide a .pst file path, or use --demo for sample output.")

    # Determine output path
    if args.output:
        output_path = args.output
    elif args.pst_file:
        base = os.path.splitext(os.path.basename(args.pst_file))[0]
        output_path = f"{base}_contacts.xlsx"
    else:
        output_path = "extracted_contacts.xlsx"

    # Extract records
    if args.demo:
        records = generate_demo_records()
    else:
        records = extract_from_pst(args.pst_file)

    if not records:
        print("\nNo contact data found. No Excel file created.")
        sys.exit(0)

    # Write to Excel
    write_to_excel(records, output_path)


if __name__ == "__main__":
    main()
