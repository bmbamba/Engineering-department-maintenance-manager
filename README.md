
═══════════════════════════════════════════════════════════
  Equipment Maintenance Manager  v1.0.0
  by Clan
═══════════════════════════════════════════════════════════

ABOUT
─────
Equipment Maintenance Manager is a desktop application for
tracking and scheduling equipment maintenance for engineering
and operations teams.


FEATURES
────────
  • Equipment register with categories and locations
  • Automatic next-service date calculation
  • Overdue and due-soon alerts (visual + email)
  • Service history log per equipment item
  • Email alerts via SMTP (overdue + daily digest)
  • PDF and CSV report export
  • Project save / load (.mmp files)
  • Bulk mark-serviced for multiple items
  • Statistics dashboard with category breakdown
  • Dark theme UI


REQUIREMENTS
────────────
  Windows 10 or later (64-bit)
  No Python installation required — the installer includes
  everything needed to run the application.


INSTALLATION
────────────
  1. Run  Maintenance Manager Setup.exe
  2. Follow the on-screen installer steps
  3. Launch from the Start Menu or Desktop shortcut


FIRST RUN
─────────
  On first launch the app creates:
    equipment.db      — your equipment database
    config.json       — your saved preferences
    maintenance_manager.log — application log

  These files are stored in the same folder as the executable.


EMAIL ALERTS
────────────
  Go to Settings > Email Alerts to configure SMTP.
  Tested with Gmail (use an App Password), Outlook, and
  most standard SMTP servers.

  Gmail setup:
    Host:     smtp.gmail.com
    Port:     587
    TLS:      Yes
    Username: your Gmail address
    Password: your App Password (not your Google password)
              Generate at: myaccount.google.com/apppasswords


CSV IMPORT FORMAT
─────────────────
  The CSV must have these column headers:
    name, equip_id, location, category, interval,
    last_service, next_service, notes

  Dates must be in YYYY-MM-DD format.
  Interval is in days.


LICENSE
───────
  See LICENSE.txt for full terms.
  Copyright (c) 2026 Clan. All rights reserved.
  You may not sell this software without written permission from Clan.


SUPPORT
───────
  For issues or feedback, contact the Clan development team.

═══════════════════════════════════════════════════════════
