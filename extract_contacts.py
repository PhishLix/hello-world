#!/usr/bin/env python3
"""
extract_contacts.py

Scans the GARY shared directory for files named INFO.doc, info.doc, Info.doc,
INFO.docx, info.docx, or Info.docx, extracts contact information from each,
and writes the results to GMMContacts.csv on the user's Desktop.

The INFO files are expected to contain unstructured legal case data with
contact blocks (attorneys, parties, etc.) separated by blank lines.

Usage:
    py extract_contacts.py --source "S:\\GARY"
    py extract_contacts.py --source "S:\\GARY" --output "C:\\Users\\You\\Desktop\\GMMContacts.csv"

Requirements:
    pip install python-docx olefile
    (LibreOffice must be installed for legacy .doc fallback conversion)
"""

import argparse
import csv
import glob
import os
import re
import subprocess
import sys
import tempfile

# ---------------------------------------------------------------------------
# Configuration defaults
# ---------------------------------------------------------------------------
DEFAULT_GARY_PATHS = [
    r"S:\GARY",
    r"G:\GARY",
    r"\\server\GARY",
    r"/mnt/gary",
    r"/mnt/GARY",
    r"/Volumes/GARY",
]

# ---------------------------------------------------------------------------
# Helpers – reading .docx and .doc files
# ---------------------------------------------------------------------------

def read_docx(filepath):
    """Return the full text of a .docx file."""
    from docx import Document
    doc = Document(filepath)
    return "\n".join(para.text for para in doc.paragraphs)


def read_doc_olefile(filepath):
    """Attempt to extract text from a legacy .doc file using olefile."""
    import olefile
    if not olefile.isOleFile(filepath):
        return None
    ole = olefile.OleFileIO(filepath)
    text_parts = []
    for stream in ole.listdir():
        stream_path = "/".join(stream)
        if stream_path.lower() in ("worddocument", "1table", "0table"):
            continue
        try:
            data = ole.openstream(stream).read()
            decoded = data.decode("utf-8", errors="ignore")
            printable = re.sub(r"[^\x20-\x7E\n\r\t]", " ", decoded)
            printable = re.sub(r" {3,}", " ", printable)
            if len(printable.strip()) > 20:
                text_parts.append(printable.strip())
        except Exception:
            pass
    ole.close()
    return "\n".join(text_parts) if text_parts else None


def read_doc_libreoffice(filepath):
    """Convert a .doc to .docx via LibreOffice, then read the .docx."""
    with tempfile.TemporaryDirectory() as tmpdir:
        try:
            subprocess.run(
                ["libreoffice", "--headless", "--convert-to", "docx",
                 "--outdir", tmpdir, filepath],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=30,
            )
        except (FileNotFoundError, subprocess.TimeoutExpired):
            return None

        base = os.path.splitext(os.path.basename(filepath))[0] + ".docx"
        converted = os.path.join(tmpdir, base)
        if os.path.isfile(converted):
            return read_docx(converted)
    return None


def read_file(filepath):
    """Return the text content of a .doc or .docx file."""
    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".docx":
        return read_docx(filepath)

    text = read_doc_olefile(filepath)
    if text:
        return text
    text = read_doc_libreoffice(filepath)
    if text:
        return text

    with open(filepath, "rb") as f:
        raw = f.read()
    decoded = raw.decode("utf-8", errors="ignore")
    return re.sub(r"[^\x20-\x7E\n\r\t]", " ", decoded)


# ---------------------------------------------------------------------------
# Contact extraction – block-based approach for unstructured legal docs
# ---------------------------------------------------------------------------

PHONE_RE = re.compile(
    r"\(?\d{3}\)?[\s.\-]*\d{3}[\s.\-]*\d{4}"
)
EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+")
# "City, ST ZIP" or "City, ST ZIP-PLUS4"
CITY_STATE_ZIP_RE = re.compile(
    r"^(.*?),\s*([A-Z]{2})\s+(\d{5}(?:-\d{4})?)(.*)$", re.MULTILINE
)
# "County County, ST" pattern for individual plaintiffs/parties
COUNTY_RE = re.compile(
    r"^([\w\s]+?County),\s*([A-Z]{2})\s*$", re.MULTILINE
)
# Street address: starts with a number, contains typical street words
STREET_RE = re.compile(
    r"^\d+[\w\s,.#-]+"
    r"(?:STREET|PKWY|PARKWAY|AVENUE|AVE|BLVD|BOULEVARD|DRIVE|DR|"
    r"ROAD|RD|LANE|LN|COURT|CT|PLACE|PL|WAY|CIRCLE|CIR|"
    r"TERRACE|TER|TRAIL|TRL|PIKE|HWY|HIGHWAY|YARDS|PLAZA)",
    re.IGNORECASE | re.MULTILINE,
)
# Patterns that indicate a block is a case header, not a contact
CASE_HEADER_RE = re.compile(
    r"(?:Civil Action|Docket N|IN RE:|v\.\s|et al\.|Hon\.\s|Assigned to:|"
    r"Transferee|Master Docket|MDL Docket)",
    re.IGNORECASE,
)
# Lines that are clearly NOT a person's name
SKIP_NAME_PATTERNS = [
    re.compile(r"(?:IN RE|Civil Action|Master Docket|MDL Docket|Assigned to|"
               r"Transferee|Hon\.|v\.|et al|ATTORNEY TO BE|LEAD ATTORNEY|"
               r"Attorneys for|^\()", re.IGNORECASE),
    PHONE_RE,
    EMAIL_RE,
    STREET_RE,
    re.compile(r"^\d+[\s.\-]*\d+[\s.\-]*\d+"),  # bare phone number
    re.compile(r"^Fax:", re.IGNORECASE),
    re.compile(r"^Tel:", re.IGNORECASE),
    re.compile(r"^Email:", re.IGNORECASE),
]
# Firm/org indicators
FIRM_RE = re.compile(
    r"\b(?:LLP|LLC|L\.L\.P|L\.L\.C|P\.?C\.?|PLLC|INC\.?|CORP\.?|"
    r"& ASSOCIATES|LAW\s+(?:GROUP|FIRM|OFFICE))\b",
    re.IGNORECASE,
)


def _is_skip_line(line):
    """Return True if the line should not be treated as a person's name."""
    for pat in SKIP_NAME_PATTERNS:
        if pat.search(line):
            return True
    return False


def _is_case_header(block):
    """Return True if the block looks like a case caption / header."""
    lines = [ln.strip() for ln in block.strip().splitlines() if ln.strip()]
    if not lines:
        return False
    # If more than half the lines match case-header patterns, skip it
    header_lines = sum(1 for ln in lines if CASE_HEADER_RE.search(ln))
    return header_lines >= len(lines) / 2


def _block_has_contact_info(block):
    """Return True if the block contains at least a phone or email."""
    return bool(PHONE_RE.search(block) or EMAIL_RE.search(block))


def _extract_one_contact(block, source_path):
    """Extract a single contact dict from a text block."""
    lines = [ln.strip() for ln in block.strip().splitlines() if ln.strip()]
    if not lines:
        return None

    contact = {
        "source_file": source_path,
        "name": "",
        "firm": "",
        "address": "",
        "city": "",
        "state": "",
        "zip": "",
        "county": "",
        "phone": "",
        "fax": "",
        "email": "",
        "role": "",
    }

    # --- Emails (collect all) ---
    emails = []
    for line in lines:
        found = EMAIL_RE.findall(line)
        emails.extend(found)
    contact["email"] = "; ".join(dict.fromkeys(emails))

    # --- Phone / Fax ---
    phones = []
    faxes = []
    for line in lines:
        nums = PHONE_RE.findall(line)
        if nums:
            if re.match(r"^\s*Fax", line, re.IGNORECASE):
                faxes.extend(nums)
            elif re.match(r"^\s*Tel", line, re.IGNORECASE):
                phones.extend(nums)
            else:
                phones.extend(nums)
    contact["phone"] = "; ".join(dict.fromkeys(phones))
    contact["fax"] = "; ".join(dict.fromkeys(faxes))

    # --- City, State, ZIP ---
    csz_match = CITY_STATE_ZIP_RE.search(block)
    if csz_match:
        contact["city"] = csz_match.group(1).strip()
        contact["state"] = csz_match.group(2).strip()
        contact["zip"] = csz_match.group(3).strip()

    # --- County, State (for individuals) ---
    for line in lines:
        county_match = COUNTY_RE.match(line)
        if county_match:
            contact["county"] = county_match.group(1).strip()
            if not contact["state"]:
                contact["state"] = county_match.group(2).strip()
            break

    # --- Street address ---
    addr_match = STREET_RE.search(block)
    if addr_match:
        contact["address"] = addr_match.group(0).strip()

    # --- Firm / organization ---
    for line in lines:
        if FIRM_RE.search(line):
            # Clean off leading/trailing junk
            firm = line.strip()
            # Remove attorney ID if stuck on same line
            firm = re.sub(r"\([\w.\s]+Attorney.*?\)", "", firm).strip()
            contact["firm"] = firm
            break

    # --- Role / description (e.g. "Attorneys for Defendants") ---
    for line in lines:
        if re.match(r"(?:Attorneys? for|Counsel for|LEAD ATTORNEY|"
                    r"ATTORNEY TO BE NOTICED)", line, re.IGNORECASE):
            # Grab this line and the next (which may have the party name)
            idx = lines.index(line)
            role_parts = [line]
            if idx + 1 < len(lines):
                next_line = lines[idx + 1]
                if not PHONE_RE.search(next_line) and not EMAIL_RE.search(next_line):
                    role_parts.append(next_line)
            contact["role"] = " / ".join(role_parts)
            break

    # --- Name: first line that looks like a person's name ---
    for line in lines:
        cleaned = line.strip()
        if not cleaned:
            continue
        if _is_skip_line(cleaned):
            continue
        if FIRM_RE.search(cleaned):
            continue
        if CITY_STATE_ZIP_RE.match(cleaned):
            continue
        if COUNTY_RE.match(cleaned):
            continue
        # Looks like a name
        contact["name"] = cleaned
        break

    # Only keep the record if we have meaningful contact info
    if contact["name"] or contact["email"] or contact["phone"]:
        return contact
    return None


def extract_contacts(text, source_path):
    """
    Split text into blocks (separated by blank lines) and extract
    contacts from blocks that contain phone numbers or emails.
    """
    contacts = []
    # Split on one or more blank lines
    blocks = re.split(r"\n\s*\n", text)

    for block in blocks:
        block = block.strip()
        if len(block) < 5:
            continue
        if not _block_has_contact_info(block):
            continue
        if _is_case_header(block):
            continue

        contact = _extract_one_contact(block, source_path)
        if contact:
            contacts.append(contact)

    return contacts


# ---------------------------------------------------------------------------
# File discovery
# ---------------------------------------------------------------------------

TARGET_FILENAMES = {"info.doc", "info.docx"}


def find_info_files(root_dir):
    """Recursively find all INFO.doc(x) files under *root_dir*."""
    matches = []
    for dirpath, _dirnames, filenames in os.walk(root_dir):
        for fname in filenames:
            if fname.lower() in TARGET_FILENAMES:
                matches.append(os.path.join(dirpath, fname))
    return sorted(matches)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def resolve_gary_path(override=None):
    """Return the GARY shared-directory path, or exit with a message."""
    if override:
        if os.path.isdir(override):
            return override
        print(f"ERROR: Specified source directory does not exist: {override}")
        sys.exit(1)

    for candidate in DEFAULT_GARY_PATHS:
        if os.path.isdir(candidate):
            return candidate

    print(
        "ERROR: Could not locate the GARY shared directory.\n"
        "Tried: " + ", ".join(DEFAULT_GARY_PATHS) + "\n\n"
        "Please specify the path with:\n"
        '    py extract_contacts.py --source "S:\\GARY"'
    )
    sys.exit(1)


def resolve_output_path(override=None):
    """Return the output CSV path, creating the Desktop dir if needed."""
    if override:
        os.makedirs(os.path.dirname(override) or ".", exist_ok=True)
        return override

    home = os.path.expanduser("~")
    candidates = [
        os.path.join(home, "Desktop"),
        os.path.join(home, "OneDrive", "Desktop"),
        os.path.join(home, "OneDrive - *", "Desktop"),
    ]
    for pattern in candidates:
        for path in glob.glob(pattern):
            if os.path.isdir(path):
                return os.path.join(path, "GMMContacts.csv")

    desktop = os.path.join(home, "Desktop")
    os.makedirs(desktop, exist_ok=True)
    return os.path.join(desktop, "GMMContacts.csv")


CSV_COLUMNS = [
    "source_file",
    "name",
    "firm",
    "address",
    "city",
    "state",
    "zip",
    "county",
    "phone",
    "fax",
    "email",
    "role",
]


def main():
    parser = argparse.ArgumentParser(
        description="Extract contacts from INFO.doc(x) files in the GARY shared directory."
    )
    parser.add_argument(
        "--source", default=None,
        help="Path to the GARY shared directory (auto-detected if omitted).",
    )
    parser.add_argument(
        "--output", default=None,
        help="Path for the output CSV file (defaults to ~/Desktop/GMMContacts.csv).",
    )
    args = parser.parse_args()

    gary_dir = resolve_gary_path(args.source)
    output_csv = resolve_output_path(args.output)

    print(f"Scanning: {gary_dir}")
    info_files = find_info_files(gary_dir)
    print(f"Found {len(info_files)} INFO file(s).\n")

    if not info_files:
        print("No INFO.doc or INFO.docx files found. Nothing to do.")
        sys.exit(0)

    all_contacts = []

    for filepath in info_files:
        print(f"  Reading: {filepath}")
        try:
            text = read_file(filepath)
        except Exception as exc:
            print(f"    WARNING: Could not read file – {exc}")
            continue

        contacts = extract_contacts(text, filepath)
        print(f"    Extracted {len(contacts)} contact(s).")
        all_contacts.extend(contacts)

    with open(output_csv, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_COLUMNS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(all_contacts)

    print(f"\nDone. {len(all_contacts)} contact(s) written to:\n  {output_csv}")


if __name__ == "__main__":
    main()
