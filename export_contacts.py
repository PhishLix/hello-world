#!/usr/bin/env python3
"""
Contacts to Excel Exporter

This script searches for a file named "contacts" on the computer,
extracts the data, and exports it to an Excel file.

Supported formats: .txt, .csv, .json, .xml, .vcf (vCard)
"""

import os
import sys
import csv
import json
import re
from pathlib import Path

try:
    from openpyxl import Workbook
except ImportError:
    print("Error: openpyxl is required. Install it with: pip install openpyxl")
    sys.exit(1)


def find_contacts_file(search_paths=None):
    """
    Search for a file named 'contacts' with various extensions.

    Args:
        search_paths: List of paths to search. Defaults to home directory.

    Returns:
        Path to the contacts file if found, None otherwise.
    """
    if search_paths is None:
        search_paths = [Path.home()]

    # Supported extensions
    extensions = ['', '.txt', '.csv', '.json', '.xml', '.vcf']

    found_files = []

    for search_path in search_paths:
        search_path = Path(search_path)
        if not search_path.exists():
            continue

        for root, dirs, files in os.walk(search_path):
            # Skip hidden directories and common system directories
            dirs[:] = [d for d in dirs if not d.startswith('.') and d not in
                      ['node_modules', '__pycache__', 'venv', '.git', 'Library',
                       'AppData', 'Application Data']]

            for file in files:
                file_lower = file.lower()
                for ext in extensions:
                    if file_lower == f"contacts{ext}":
                        found_files.append(Path(root) / file)

    if not found_files:
        return None

    if len(found_files) == 1:
        return found_files[0]

    # If multiple files found, let user choose
    print("Multiple contacts files found:")
    for i, f in enumerate(found_files, 1):
        print(f"  {i}. {f}")

    while True:
        try:
            choice = input("Select a file (enter number): ")
            idx = int(choice) - 1
            if 0 <= idx < len(found_files):
                return found_files[idx]
        except (ValueError, KeyboardInterrupt):
            pass
        print("Invalid selection. Please try again.")


def parse_txt_file(file_path):
    """Parse a plain text contacts file."""
    contacts = []

    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()

    lines = content.strip().split('\n')

    if not lines:
        return [], []

    # Try to detect if first line is a header
    first_line = lines[0]

    # Check for common delimiters
    for delimiter in ['\t', ',', '|', ';']:
        if delimiter in first_line:
            # Treat as delimited file
            reader = csv.reader(lines, delimiter=delimiter)
            rows = list(reader)
            if rows:
                headers = rows[0]
                contacts = rows[1:] if len(rows) > 1 else []
                return headers, contacts

    # Parse as key-value pairs or structured text
    headers = ['Name', 'Email', 'Phone', 'Address', 'Notes']
    current_contact = {}

    for line in lines:
        line = line.strip()
        if not line:
            if current_contact:
                contacts.append(current_contact)
                current_contact = {}
            continue

        # Try to parse key: value format
        if ':' in line:
            key, value = line.split(':', 1)
            key = key.strip().title()
            value = value.strip()
            current_contact[key] = value
        else:
            # Assume it's a name if no key
            if 'Name' not in current_contact:
                current_contact['Name'] = line
            else:
                current_contact['Notes'] = current_contact.get('Notes', '') + ' ' + line

    if current_contact:
        contacts.append(current_contact)

    # Convert dict contacts to list format
    if contacts and isinstance(contacts[0], dict):
        all_keys = set()
        for c in contacts:
            all_keys.update(c.keys())
        headers = sorted(all_keys)
        contact_rows = [[c.get(h, '') for h in headers] for c in contacts]
        return headers, contact_rows

    return headers, contacts


def parse_csv_file(file_path):
    """Parse a CSV contacts file."""
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
        # Try to detect the dialect
        sample = f.read(4096)
        f.seek(0)

        try:
            dialect = csv.Sniffer().sniff(sample)
            reader = csv.reader(f, dialect)
        except csv.Error:
            reader = csv.reader(f)

        rows = list(reader)

    if not rows:
        return [], []

    headers = rows[0]
    contacts = rows[1:] if len(rows) > 1 else []

    return headers, contacts


def parse_json_file(file_path):
    """Parse a JSON contacts file."""
    with open(file_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    # Handle different JSON structures
    if isinstance(data, list):
        contacts_list = data
    elif isinstance(data, dict):
        # Try common keys for contacts arrays
        for key in ['contacts', 'data', 'people', 'entries', 'items']:
            if key in data and isinstance(data[key], list):
                contacts_list = data[key]
                break
        else:
            # Single contact or flat structure
            contacts_list = [data]
    else:
        return [], []

    if not contacts_list:
        return [], []

    # Extract all unique keys as headers
    all_keys = set()
    for contact in contacts_list:
        if isinstance(contact, dict):
            all_keys.update(contact.keys())

    headers = sorted(all_keys)

    # Convert to rows
    contacts = []
    for contact in contacts_list:
        if isinstance(contact, dict):
            row = [str(contact.get(h, '')) for h in headers]
            contacts.append(row)

    return headers, contacts


def parse_xml_file(file_path):
    """Parse an XML contacts file."""
    try:
        import xml.etree.ElementTree as ET
    except ImportError:
        print("XML parsing not available")
        return [], []

    tree = ET.parse(file_path)
    root = tree.getroot()

    contacts = []
    all_keys = set()

    # Find contact elements (try common names)
    contact_elements = []
    for tag in ['contact', 'person', 'entry', 'item', 'record']:
        contact_elements = root.findall(f'.//{tag}')
        if contact_elements:
            break

    if not contact_elements:
        # Use direct children of root
        contact_elements = list(root)

    for elem in contact_elements:
        contact = {}
        for child in elem:
            tag = child.tag.replace('{', '').split('}')[-1]  # Remove namespace
            contact[tag] = child.text or ''
            all_keys.add(tag)

        # Also check attributes
        for attr, value in elem.attrib.items():
            contact[attr] = value
            all_keys.add(attr)

        if contact:
            contacts.append(contact)

    headers = sorted(all_keys)
    contact_rows = [[c.get(h, '') for h in headers] for c in contacts]

    return headers, contact_rows


def parse_vcf_file(file_path):
    """Parse a vCard (.vcf) contacts file."""
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()

    contacts = []
    headers = ['Name', 'Full Name', 'Email', 'Phone', 'Organization', 'Title', 'Address', 'Notes']

    # Split into individual vCards
    vcards = re.split(r'BEGIN:VCARD', content, flags=re.IGNORECASE)

    for vcard in vcards:
        if not vcard.strip():
            continue

        contact = {h: '' for h in headers}

        # Parse vCard fields
        for line in vcard.split('\n'):
            line = line.strip()
            if ':' not in line:
                continue

            field, value = line.split(':', 1)
            field = field.split(';')[0].upper()

            if field == 'FN':
                contact['Full Name'] = value
            elif field == 'N':
                parts = value.split(';')
                name_parts = [p for p in parts if p]
                contact['Name'] = ' '.join(reversed(name_parts))
            elif field == 'EMAIL':
                contact['Email'] = value if not contact['Email'] else f"{contact['Email']}; {value}"
            elif field == 'TEL':
                contact['Phone'] = value if not contact['Phone'] else f"{contact['Phone']}; {value}"
            elif field == 'ORG':
                contact['Organization'] = value.replace(';', ' ')
            elif field == 'TITLE':
                contact['Title'] = value
            elif field == 'ADR':
                contact['Address'] = value.replace(';', ', ')
            elif field == 'NOTE':
                contact['Notes'] = value

        if any(contact.values()):
            contacts.append([contact[h] for h in headers])

    return headers, contacts


def parse_contacts_file(file_path):
    """
    Parse a contacts file based on its extension.

    Returns:
        Tuple of (headers, contacts) where contacts is a list of lists.
    """
    file_path = Path(file_path)
    ext = file_path.suffix.lower()

    parsers = {
        '.csv': parse_csv_file,
        '.json': parse_json_file,
        '.xml': parse_xml_file,
        '.vcf': parse_vcf_file,
        '.txt': parse_txt_file,
        '': parse_txt_file,  # No extension, try as text
    }

    parser = parsers.get(ext, parse_txt_file)

    try:
        return parser(file_path)
    except Exception as e:
        print(f"Error parsing file: {e}")
        # Fallback to text parser
        if parser != parse_txt_file:
            print("Trying text parser as fallback...")
            return parse_txt_file(file_path)
        return [], []


def export_to_excel(headers, contacts, output_path):
    """
    Export contacts to an Excel file.

    Args:
        headers: List of column headers
        contacts: List of contact rows
        output_path: Path for the output Excel file
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Contacts"

    # Write headers
    if headers:
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = cell.font.copy(bold=True)

    # Write data
    for row_idx, contact in enumerate(contacts, 2):
        for col_idx, value in enumerate(contact, 1):
            ws.cell(row=row_idx, column=col_idx, value=value)

    # Auto-adjust column widths
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column_letter].width = adjusted_width

    wb.save(output_path)
    print(f"Successfully exported {len(contacts)} contacts to {output_path}")


def main():
    """Main entry point."""
    print("Contacts to Excel Exporter")
    print("=" * 40)

    # Determine search paths
    search_paths = [Path.home()]

    # Add common locations
    if sys.platform == 'win32':
        search_paths.extend([
            Path.home() / 'Documents',
            Path.home() / 'Desktop',
            Path.home() / 'Downloads',
        ])
    else:
        search_paths.extend([
            Path.home() / 'Documents',
            Path.home() / 'Desktop',
            Path.home() / 'Downloads',
            Path('/tmp'),
        ])

    # Allow custom path as argument
    if len(sys.argv) > 1:
        custom_path = Path(sys.argv[1])
        if custom_path.is_file():
            contacts_file = custom_path
        else:
            search_paths = [custom_path]
            contacts_file = None
    else:
        contacts_file = None

    # Search for contacts file
    if contacts_file is None:
        print("\nSearching for contacts file...")
        contacts_file = find_contacts_file(search_paths)

    if contacts_file is None:
        print("\nNo contacts file found!")
        print("\nPlease create a contacts file with one of these names:")
        print("  - contacts.txt")
        print("  - contacts.csv")
        print("  - contacts.json")
        print("  - contacts.xml")
        print("  - contacts.vcf")
        print("\nOr specify the path as an argument:")
        print(f"  python {sys.argv[0]} /path/to/your/contacts/file")
        sys.exit(1)

    print(f"\nFound contacts file: {contacts_file}")

    # Parse the file
    print("Parsing contacts...")
    headers, contacts = parse_contacts_file(contacts_file)

    if not contacts:
        print("No contacts found in the file!")
        sys.exit(1)

    print(f"Found {len(contacts)} contacts with {len(headers)} fields")

    # Determine output path
    output_dir = Path.cwd()
    output_path = output_dir / "contacts.xlsx"

    # Export to Excel
    print(f"\nExporting to Excel...")
    export_to_excel(headers, contacts, output_path)

    print("\nDone!")
    return 0


if __name__ == '__main__':
    sys.exit(main())
