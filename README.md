# IT Operations Dashboard

A local ticketing system dashboard that syncs with Excel.

## Quick Start

### Windows
1. Install Python 3.10+ from [python.org](https://python.org)
2. Double-click `setup_windows.bat` OR run in Command Prompt:
   ```
   python -m venv venv
   venv\Scripts\activate
   pip install flask pandas openpyxl
   python app.py
   ```
3. Open browser to `http://127.0.0.1:5000`

### Mac
1. Run in Terminal:
   ```
   python3 -m venv venv
   source venv/bin/activate
   pip install flask pandas openpyxl
   python app.py
   ```
2. Open browser to `http://127.0.0.1:5000`

## Default Login
- **Username:** `admin`
- **Password:** `admin123`

## Files to Copy
Copy the entire folder. These are the key files:
- `app.py` - Backend server
- `templates/` - HTML pages
- `tickets.xlsx` - Your ticket data
- `.env/users.json` - User credentials (keep private!)

## Notes
- Don't share `.env/` folder (contains passwords)
- `tickets.xlsx` is your data source - keep it backed up
