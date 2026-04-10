# Outlook PST Contact Extractor

Extract contact information from Microsoft Outlook `.pst` files and export the data to a formatted Excel spreadsheet.

## What It Does

This tool scans every email in a `.pst` file, parses structured fields from the email body (key-value pairs such as `First Name: John`), and writes one row per contact to an Excel workbook with the following columns:

| Column | Description |
|---|---|
| Source ID | Case/reference ID found in the email, or auto-generated |
| First Name | Contact's first/given name |
| Last Name | Contact's last/family name |
| City | City of residence |
| State | US state (normalized to two-letter abbreviation) |
| Phone | Phone number (normalized to `(XXX) XXX-XXXX`) |
| Email | Email address |
| Amount of Money Lost | Reported financial loss (formatted as currency) |

## Prerequisites

- Python 3.8+
- pip

## Installation

```bash
pip install -r requirements.txt
```

**Note:** The `libpff-python` package (for reading `.pst` files) may require system-level dependencies on some platforms. On Debian/Ubuntu:

```bash
sudo apt-get install libpff-dev
pip install libpff-python
```

On macOS with Homebrew:

```bash
brew install libpff
pip install libpff-python
```

## Usage

### Extract contacts from a .pst file

```bash
python extract_outlook_contacts.py inbox.pst
```

This creates `inbox_contacts.xlsx` in the current directory.

### Specify a custom output file

```bash
python extract_outlook_contacts.py inbox.pst -o my_contacts.xlsx
```

### Run in demo mode (no .pst file needed)

```bash
python extract_outlook_contacts.py --demo
python extract_outlook_contacts.py --demo -o sample_output.xlsx
```

Demo mode uses built-in sample email data to generate a fully populated Excel file so you can preview the output format.

## How Contact Data Is Extracted

The script recognizes structured key-value patterns commonly found in complaint or report emails:

```
First Name: Maria
Last Name: Gonzalez
City: Newark
State: New Jersey
Phone: (973) 555-0142
Email: maria.gonzalez@example.com
Amount of Money Lost: $4,500.00
```

It also handles variations such as:
- `Name: Williams, Sandra` (Last, First format)
- `Full Name: David Kim`
- `Tel`, `Cell`, `Mobile`, `Telephone` for phone fields
- `Total Loss`, `Financial Loss`, `Damages` for money fields
- `Case ID`, `Reference`, `Ticket` for source ID fields
- Full state names are normalized to two-letter abbreviations (e.g., `New Jersey` to `NJ`)

If no key-value match is found, the script falls back to standalone regex patterns to find email addresses, phone numbers, and dollar amounts in the email body.

## Output Format

The Excel file includes:
- Formatted header row with column filters
- Frozen header pane for easy scrolling
- Auto-sized columns
- Currency formatting on the "Amount of Money Lost" column

## Author

Justin

## License

This project is licensed under the MIT License.
