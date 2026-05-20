# Law Firm Time Tracker

A self-hosted web app for tracking billable time at a small law firm (2–10 users).
Runs as a single Flask server on one machine on the office LAN; everyone else
uses it from a browser.

## Features

- **Per-user start/stop timer** with project selection and a free-text description
  of work performed.
- **Manual time entry** for when a user forgets the timer or logs time after the fact.
- **Edit / delete** time entries (until they have been included on an invoice).
- **Projects** with client name, address, default billing rate, and a team roster.
  Per-project per-user billing-rate overrides are supported.
- **Team leaders** — any project member can be marked as a leader. Leaders can run
  reports and generate invoices for that project.
- **Reports** — filter time on a project by date range, see totals per timekeeper,
  and a full detail listing. Printable.
- **Invoices** — generate a printable invoice for a project covering a date range;
  entries included are marked as invoiced and become read-only. The invoice
  document shows each user's task, description, hours, billing rate, and total
  due, plus a summary by timekeeper.
- **User management** (admin only) with passwords stored as salted hashes.

## Quick start

Requires Python 3.10+.

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# Generate a stable secret key so sessions survive restarts:
export SECRET_KEY="$(python -c 'import secrets; print(secrets.token_hex(32))')"

python app.py
```

The server listens on `0.0.0.0:5000` by default. From any machine on the same
LAN, browse to `http://<server-ip>:5000/`.

A default administrator account is created on first startup:

- username: `admin`
- password: `admin`

**Change this password immediately** under *Users → admin → Edit*, then add
real user accounts and projects.

## Configuration

Environment variables:

| Variable       | Default                               | Purpose                    |
| -------------- | ------------------------------------- | -------------------------- |
| `SECRET_KEY`   | `change-me-in-production`             | Flask session signing key  |
| `DATABASE_URL` | `sqlite:///<instance>/timetracker.db` | SQLAlchemy database URI    |
| `HOST`         | `0.0.0.0`                             | Bind address               |
| `PORT`         | `5000`                                | Listen port                |

The SQLite database file lives in `instance/timetracker.db`. Back it up
periodically (a plain file copy when the server is idle is sufficient).

## Production deployment

For a small office, the simplest robust setup is to run the app under a
process supervisor with a real WSGI server. Example with `gunicorn`:

```bash
pip install gunicorn
gunicorn -w 2 -b 0.0.0.0:5000 'app:app'
```

For a Linux server, drop the following into `/etc/systemd/system/timetracker.service`
(adjust paths and user):

```ini
[Unit]
Description=Law Firm Time Tracker
After=network.target

[Service]
User=timetracker
WorkingDirectory=/opt/timetracker
Environment="SECRET_KEY=<generated-secret>"
ExecStart=/opt/timetracker/.venv/bin/gunicorn -w 2 -b 0.0.0.0:5000 app:app
Restart=on-failure

[Install]
WantedBy=multi-user.target
```

Then `systemctl enable --now timetracker` and the app will start on boot.

If the firm has internet exposure or wants HTTPS, put the app behind nginx
or Caddy with TLS — but for a pure LAN deployment, plain HTTP on a trusted
network is acceptable.

## Roles

- **User** — log in, run their own timer, enter time manually, view and edit
  their own entries.
- **Project leader** (assigned per-project) — additionally see all entries on
  their project(s), run reports, and generate invoices.
- **Administrator** — create/edit users and projects, assign leaders, see and
  edit any time entry.
