#!/usr/bin/env python3
"""
extract_contacts.py

Scans the GARY shared directory for files named INFO.doc, info.doc, Info.doc,
INFO.docx, info.docx, or Info.docx, extracts contact information from each,
and writes the results to GMMContacts.csv on the user's Desktop.

Usage:
    python3 extract_contacts.py
    python3 extract_contacts.py --source "G:\\GARY" --output "C:\\Users\\YourName\\Desktop\\GMMContacts.csv"

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
# Configuration defaults – adjust these to match your environment
# ---------------------------------------------------------------------------
# Common paths where the GARY share might be mounted.  The script tries each
# in order and uses the first one that exists, unless --source is supplied.
DEFAULT_GARY_PATHS = [
    r"G:\GARY",
    r"G:",
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
    # The main text stream in a .doc is 'WordDocument', but raw text is
    # often easier to pull from the 'Word Document' or via the
    # '\x01CompObj' or by reading all streams.  A lightweight approach:
    text_parts = []
    for stream in ole.listdir():
        stream_path = "/".join(stream)
        if stream_path.lower() in ("worddocument", "1table", "0table"):
            continue  # binary streams
        try:
            data = ole.openstream(stream).read()
            decoded = data.decode("utf-8", errors="ignore")
            # Keep only printable stretches
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
                [
                    "libreoffice",
                    "--headless",
                    "--convert-to",
                    "docx",
                    "--outdir",
                    tmpdir,
                    filepath,
                ],
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

    # Legacy .doc – try olefile first, then LibreOffice
    text = read_doc_olefile(filepath)
    if text:
        return text
    text = read_doc_libreoffice(filepath)
    if text:
        return text

    # Last resort: read raw bytes and pull printable strings
    with open(filepath, "rb") as f:
        raw = f.read()
    decoded = raw.decode("utf-8", errors="ignore")
    return re.sub(r"[^\x20-\x7E\n\r\t]", " ", decoded)


# ---------------------------------------------------------------------------
# Contact extraction
# ---------------------------------------------------------------------------

# Regex patterns for typical contact-info fields
PHONE_RE = re.compile(
    r"(\+?1?\s*[-.\(]?\s*\d{3}\s*[-.\)]?\s*\d{3}\s*[-.]?\s*\d{4})"
)
EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+")
ZIP_RE = re.compile(r"\b\d{5}(?:-\d{4})?\b")

# Label patterns that appear before a value on the same or next line
LABEL_RE = {
    "name": re.compile(
        r"(?:(?:full\s*)?name|contact(?:\s*name)?|attorney|lawyer|firm\s*name"
        r"|company|manager|owner)\s*[:;\-]?\s*(.*)",
        re.IGNORECASE,
    ),
    "address": re.compile(
        r"(?:address|street|mailing\s*address|location)\s*[:;\-]?\s*(.*)",
        re.IGNORECASE,
    ),
    "city": re.compile(
        r"(?:city)\s*[:;\-]\s*(.*)", re.IGNORECASE
    ),
    "state": re.compile(
        r"(?:state)\s*[:;\-]\s*(.*)", re.IGNORECASE
    ),
    "zip": re.compile(
        r"(?:zip(?:\s*code)?|postal\s*code)\s*[:;\-]?\s*(.*)", re.IGNORECASE
    ),
    "phone": re.compile(
        r"(?:phone|tel(?:ephone)?|cell|mobile|fax|office)\s*[:;\-#]?\s*(.*)",
        re.IGNORECASE,
    ),
    "email": re.compile(
        r"(?:e[\-\s]?mail)\s*[:;\-]?\s*(.*)", re.IGNORECASE
    ),
}


def extract_contacts(text, source_path):
    """
    Parse *text* and return a list of dicts, each representing one contact.
    Each dict has keys: source, name, address, city, state, zip, phone,
    email, and raw_text (for debugging / manual review).
    """
    contacts = []
    # Try to split on obvious record separators (horizontal rules, page
    # breaks, repeated dashes/underscores, or large blank gaps).
    records = re.split(
        r"\n\s*(?:[-_=*]{5,}|={3,}|\f)\s*\n|\n{3,}", text
    )
    if not records:
        records = [text]

    for record in records:
        record = record.strip()
        if len(record) < 10:
            continue

        contact = {
            "source_file": source_path,
            "name": "",
            "address": "",
            "city": "",
            "state": "",
            "zip": "",
            "phone": "",
            "email": "",
        }

        # 1. Try labelled fields
        for field, pattern in LABEL_RE.items():
            match = pattern.search(record)
            if match:
                value = match.group(1).strip().strip(":;- ")
                if value:
                    contact[field] = value

        # 1b. Fallback: look for "City, ST ZIP" pattern on any line
        if not contact["city"] or not contact["state"]:
            csz = re.search(
                r"^([A-Za-z .-]+),\s*([A-Z]{2})\s+(\d{5}(?:-\d{4})?)",
                record,
                re.MULTILINE,
            )
            if csz:
                if not contact["city"]:
                    contact["city"] = csz.group(1).strip()
                if not contact["state"]:
                    contact["state"] = csz.group(2).strip()
                if not contact["zip"]:
                    contact["zip"] = csz.group(3).strip()

        # 2. Fallback: extract emails and phones via global regex
        if not contact["email"]:
            emails = EMAIL_RE.findall(record)
            if emails:
                contact["email"] = "; ".join(dict.fromkeys(emails))

        if not contact["phone"]:
            phones = PHONE_RE.findall(record)
            if phones:
                cleaned = []
                for p in phones:
                    p = p.strip()
                    # Filter out numbers that are too short
                    digits = re.sub(r"\D", "", p)
                    if len(digits) >= 7:
                        cleaned.append(p)
                contact["phone"] = "; ".join(dict.fromkeys(cleaned))

        if not contact["zip"]:
            zips = ZIP_RE.findall(record)
            if zips:
                contact["zip"] = zips[0]

        # 3. Heuristic: if no name found, use the first non-empty line
        if not contact["name"]:
            for line in record.splitlines():
                line = line.strip()
                if line and not re.match(r"^[-_=*]+$", line):
                    # Skip lines that look like labels only
                    if ":" in line and len(line.split(":", 1)[0].split()) <= 3:
                        continue
                    contact["name"] = line
                    break

        # Only keep the record if we found at least a name or email or phone
        if contact["name"] or contact["email"] or contact["phone"]:
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
        '    python3 extract_contacts.py --source "G:\\GARY"'
    )
    sys.exit(1)


def resolve_output_path(override=None):
    """Return the output CSV path, creating the Desktop dir if needed."""
    if override:
        os.makedirs(os.path.dirname(override) or ".", exist_ok=True)
        return override

    # Try common Desktop locations
    home = os.path.expanduser("~")
    candidates = [
        os.path.join(home, "Desktop"),
        os.path.join(home, "OneDrive", "Desktop"),
        os.path.join(home, "OneDrive - *", "Desktop"),  # glob-friendly
    ]
    for pattern in candidates:
        for path in glob.glob(pattern):
            if os.path.isdir(path):
                return os.path.join(path, "GMMContacts.csv")

    # Fallback: create Desktop in home
    desktop = os.path.join(home, "Desktop")
    os.makedirs(desktop, exist_ok=True)
    return os.path.join(desktop, "GMMContacts.csv")


CSV_COLUMNS = [
    "source_file",
    "name",
    "address",
    "city",
    "state",
    "zip",
    "phone",
    "email",
]


def main():
    parser = argparse.ArgumentParser(
        description="Extract contacts from INFO.doc(x) files in the GARY shared directory."
    )
    parser.add_argument(
        "--source",
        default=None,
        help="Path to the GARY shared directory (auto-detected if omitted).",
    )
    parser.add_argument(
        "--output",
        default=None,
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

    # Write CSV
    with open(output_csv, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_COLUMNS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(all_contacts)

    print(f"\nDone. {len(all_contacts)} contact(s) written to:\n  {output_csv}")


if __name__ == "__main__":
    main()
