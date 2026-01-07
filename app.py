import os
import json
import hashlib
import secrets
from functools import wraps
from flask import Flask, render_template, jsonify, request, session, redirect, url_for, send_file
from datetime import datetime, time
from openpyxl import Workbook
from io import BytesIO
import psycopg2
from psycopg2.extras import RealDictCursor

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'ticketing-dashboard-secret-key-2025')

# Database URL from environment (Render provides this)
DATABASE_URL = os.environ.get('DATABASE_URL')

# Check multiple locations for users file (local dev, Render secrets, Docker)
USERS_FILE_LOCATIONS = [
    '.env/users.json',              # Local development
    '/etc/secrets/users_data.json', # Render secret files
    'users.json',                   # Working directory fallback
]

def get_users_file():
    """Find the users file from multiple possible locations."""
    for path in USERS_FILE_LOCATIONS:
        if os.path.exists(path):
            return path
    return USERS_FILE_LOCATIONS[0]  # Default to first option

USERS_FILE = get_users_file()

# --- DATABASE FUNCTIONS ---
def clean_database_url(url):
    """Clean database URL by removing unsupported parameters for psycopg2."""
    if not url:
        return url

    from urllib.parse import urlparse, parse_qs, urlencode, urlunparse

    try:
        parsed = urlparse(url)

        # Parse query parameters
        params = parse_qs(parsed.query)

        # Remove unsupported parameters (like channel_binding from Neon)
        unsupported = ['channel_binding', 'options']
        for param in unsupported:
            params.pop(param, None)

        # Flatten params (parse_qs returns lists)
        clean_params = {k: v[0] if len(v) == 1 else v for k, v in params.items()}

        # Reconstruct URL
        new_query = urlencode(clean_params)
        clean_url = urlunparse((
            parsed.scheme,
            parsed.netloc,
            parsed.path,
            parsed.params,
            new_query,
            parsed.fragment
        ))

        print(f"[DB] Cleaned URL parameters: removed {[p for p in unsupported if p in parse_qs(parsed.query)]}")
        return clean_url
    except Exception as e:
        print(f"[DB] Warning: Could not parse URL, using as-is: {e}")
        return url

def get_db_connection():
    """Get a database connection."""
    if not DATABASE_URL:
        print("[DB] ERROR: DATABASE_URL not set!")
        return None
    try:
        clean_url = clean_database_url(DATABASE_URL)
        conn = psycopg2.connect(clean_url, cursor_factory=RealDictCursor)
        return conn
    except Exception as e:
        print(f"[DB] ERROR: Failed to connect to database: {e}")
        return None

def init_database():
    """Initialize the database tables."""
    conn = get_db_connection()
    if not conn:
        print("[DB] Cannot initialize database - no connection")
        return False

    try:
        cur = conn.cursor()
        cur.execute('''
            CREATE TABLE IF NOT EXISTS tickets (
                id SERIAL PRIMARY KEY,
                ticket_id VARCHAR(20) UNIQUE NOT NULL,
                title TEXT NOT NULL,
                status VARCHAR(50) DEFAULT 'Open',
                priority VARCHAR(50) DEFAULT 'Low',
                request_type VARCHAR(100),
                staff_assigned VARCHAR(100),
                requester VARCHAR(100),
                date_opened DATE DEFAULT CURRENT_DATE,
                description TEXT,
                resolution_notes TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        conn.commit()
        print("[DB] Database initialized successfully")

        # Check if we need to migrate data from Excel
        cur.execute("SELECT COUNT(*) as count FROM tickets")
        count = cur.fetchone()['count']
        if count == 0:
            print("[DB] No tickets in database, attempting to migrate from Excel...")
            migrate_from_excel(conn)
        else:
            print(f"[DB] Found {count} existing tickets in database")

        cur.close()
        conn.close()
        return True
    except Exception as e:
        print(f"[DB] ERROR initializing database: {e}")
        import traceback
        traceback.print_exc()
        return False

def migrate_from_excel(conn):
    """Migrate existing data from Excel file to database."""
    try:
        from openpyxl import load_workbook
        EXCEL_FILE = 'tickets.xlsx'
        SHEET_NAME = 'IT Service Tickets'

        if not os.path.exists(EXCEL_FILE):
            print(f"[DB] Excel file {EXCEL_FILE} not found, skipping migration")
            return

        wb = load_workbook(EXCEL_FILE)
        ws = wb[SHEET_NAME]

        cur = conn.cursor()
        migrated = 0

        for row in range(2, 1000):
            title = ws[f'B{row}'].value
            if not title:
                break

            ticket_id = f"IT-2025{row-1:04d}"
            status = ws[f'C{row}'].value or 'Open'
            priority = ws[f'D{row}'].value or 'Low'
            request_type = ws[f'E{row}'].value or ''
            staff_assigned = ws[f'F{row}'].value or ''
            requester = ws[f'G{row}'].value or ''
            date_val = ws[f'H{row}'].value
            description = ws[f'J{row}'].value or ''
            resolution = ws[f'K{row}'].value or ''

            # Handle date conversion
            if date_val is None:
                date_opened = datetime.now().date()
            elif hasattr(date_val, 'date'):
                date_opened = date_val.date() if hasattr(date_val, 'date') else date_val
            elif hasattr(date_val, 'strftime'):
                date_opened = date_val
            else:
                try:
                    date_opened = datetime.strptime(str(date_val), '%Y-%m-%d').date()
                except:
                    date_opened = datetime.now().date()

            cur.execute('''
                INSERT INTO tickets (ticket_id, title, status, priority, request_type,
                                    staff_assigned, requester, date_opened, description, resolution_notes)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                ON CONFLICT (ticket_id) DO NOTHING
            ''', (ticket_id, str(title).strip(), str(status).strip(), str(priority).strip(),
                  str(request_type).strip(), str(staff_assigned).strip(), str(requester).strip(),
                  date_opened, str(description).strip(), str(resolution).strip() if resolution else ''))
            migrated += 1

        conn.commit()
        cur.close()
        wb.close()
        print(f"[DB] Migrated {migrated} tickets from Excel to database")
    except Exception as e:
        print(f"[DB] ERROR during Excel migration: {e}")
        import traceback
        traceback.print_exc()

# --- USER MANAGEMENT ---
def load_users():
    """Load users from JSON file."""
    try:
        users_path = get_users_file()
        print(f"[AUTH] Loading users from: {users_path}")
        with open(users_path, 'r') as f:
            data = json.load(f)
            users = data.get('users', [])
            print(f"[AUTH] Loaded {len(users)} users successfully")
            return users
    except FileNotFoundError:
        print(f"[AUTH] ERROR: Users file not found at any location: {USERS_FILE_LOCATIONS}")
        return []
    except json.JSONDecodeError as e:
        print(f"[AUTH] ERROR: Invalid JSON in users file: {e}")
        return []
    except Exception as e:
        print(f"[AUTH] ERROR: Failed to load users: {e}")
        return []

def save_users(users):
    """Save users to JSON file."""
    with open(USERS_FILE, 'w') as f:
        json.dump({'users': users}, f, indent=2)

def hash_password(password):
    """Hash password using SHA256."""
    return hashlib.sha256(password.encode()).hexdigest()

def verify_password(password, password_hash):
    """Verify password against hash."""
    return hash_password(password) == password_hash

def get_user(username):
    """Get user by username."""
    users = load_users()
    for user in users:
        if user['username'] == username:
            return user
    return None

def login_required(f):
    """Decorator to require login."""
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user' not in session:
            return jsonify({'error': 'Not authenticated'}), 401
        return f(*args, **kwargs)
    return decorated

def admin_required(f):
    """Decorator to require admin role."""
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user' not in session:
            return jsonify({'error': 'Not authenticated'}), 401
        if session['user'].get('role') != 'admin':
            return jsonify({'error': 'Admin access required'}), 403
        return f(*args, **kwargs)
    return decorated

# --- DATABASE TICKET FUNCTIONS ---
def read_tickets_from_db():
    """Read all tickets from the database."""
    conn = get_db_connection()
    if not conn:
        return []

    try:
        cur = conn.cursor()
        cur.execute('''
            SELECT ticket_id, title, status, priority, request_type,
                   staff_assigned, requester, date_opened, description, resolution_notes
            FROM tickets
            ORDER BY date_opened DESC, id DESC
        ''')
        rows = cur.fetchall()
        cur.close()
        conn.close()

        tickets = []
        for row in rows:
            date_opened = row['date_opened']
            if date_opened:
                date_str = date_opened.strftime('%Y-%m-%d') if hasattr(date_opened, 'strftime') else str(date_opened)
                # Calculate days open
                if row['status'] in ['Closed', 'Resolved']:
                    days_open = -1
                else:
                    days_open = (datetime.now().date() - date_opened).days
                    if days_open < 0:
                        days_open = 0
            else:
                date_str = ''
                days_open = 0

            tickets.append({
                'ticket_id': row['ticket_id'],
                'title': row['title'] or '',
                'status': row['status'] or 'Open',
                'priority': row['priority'] or 'Low',
                'request_type': row['request_type'] or '',
                'staff_assigned': row['staff_assigned'] or '',
                'requester': row['requester'] or '',
                'date_opened': date_str,
                'days_open': days_open,
                'description': row['description'] or '',
                'resolution_notes': row['resolution_notes'] or ''
            })

        return tickets
    except Exception as e:
        print(f"[DB] ERROR reading tickets: {e}")
        import traceback
        traceback.print_exc()
        return []

def get_next_ticket_id():
    """Generate the next ticket ID."""
    conn = get_db_connection()
    if not conn:
        return 'IT-20250001'

    try:
        cur = conn.cursor()
        cur.execute("SELECT ticket_id FROM tickets ORDER BY id DESC LIMIT 1")
        row = cur.fetchone()
        cur.close()
        conn.close()

        if row:
            # Extract number from ticket_id (e.g., IT-20250001 -> 20250001)
            num = int(row['ticket_id'].replace('IT-', ''))
            return f"IT-{num + 1:08d}"
        return 'IT-20250001'
    except Exception as e:
        print(f"[DB] ERROR getting next ticket ID: {e}")
        return 'IT-20250001'

def create_ticket_in_db(ticket_data):
    """Create a new ticket in the database."""
    conn = get_db_connection()
    if not conn:
        return None

    try:
        cur = conn.cursor()
        ticket_id = get_next_ticket_id()

        cur.execute('''
            INSERT INTO tickets (ticket_id, title, status, priority, request_type,
                                staff_assigned, requester, date_opened, description, resolution_notes)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            RETURNING ticket_id
        ''', (
            ticket_id,
            ticket_data.get('title', ''),
            ticket_data.get('status', 'Open'),
            ticket_data.get('priority', 'Low'),
            ticket_data.get('request_type', ''),
            ticket_data.get('staff_assigned', ''),
            ticket_data.get('requester', ''),
            datetime.now().date(),
            ticket_data.get('description', ''),
            ticket_data.get('resolution_notes', '')
        ))

        result = cur.fetchone()
        conn.commit()
        cur.close()
        conn.close()

        print(f"[DB] Created ticket: {ticket_id}")
        return result['ticket_id']
    except Exception as e:
        print(f"[DB] ERROR creating ticket: {e}")
        import traceback
        traceback.print_exc()
        return None

def update_ticket_in_db(ticket_id, ticket_data):
    """Update a ticket in the database."""
    conn = get_db_connection()
    if not conn:
        return False

    try:
        cur = conn.cursor()

        # Build dynamic update query
        updates = []
        values = []

        field_mapping = {
            'title': 'title',
            'status': 'status',
            'priority': 'priority',
            'request_type': 'request_type',
            'staff_assigned': 'staff_assigned',
            'requester': 'requester',
            'description': 'description',
            'resolution_notes': 'resolution_notes'
        }

        for key, db_field in field_mapping.items():
            if key in ticket_data:
                updates.append(f"{db_field} = %s")
                values.append(ticket_data[key])

        if not updates:
            return True

        updates.append("updated_at = CURRENT_TIMESTAMP")
        values.append(ticket_id)

        query = f"UPDATE tickets SET {', '.join(updates)} WHERE ticket_id = %s"
        cur.execute(query, values)

        updated = cur.rowcount > 0
        conn.commit()
        cur.close()
        conn.close()

        print(f"[DB] Updated ticket: {ticket_id}, affected rows: {cur.rowcount}")
        return updated
    except Exception as e:
        print(f"[DB] ERROR updating ticket: {e}")
        import traceback
        traceback.print_exc()
        return False

def delete_ticket_from_db(ticket_id):
    """Delete a ticket from the database."""
    conn = get_db_connection()
    if not conn:
        return False

    try:
        cur = conn.cursor()
        cur.execute("DELETE FROM tickets WHERE ticket_id = %s", (ticket_id,))
        deleted = cur.rowcount > 0
        conn.commit()
        cur.close()
        conn.close()

        print(f"[DB] Deleted ticket: {ticket_id}")
        return deleted
    except Exception as e:
        print(f"[DB] ERROR deleting ticket: {e}")
        return False

def calculate_stats(tickets):
    """Calculates KPIs from the ticket list."""
    total = len(tickets)
    open_count = sum(1 for t in tickets if t['status'] not in ['Closed', 'Resolved'])
    closed_count = sum(1 for t in tickets if t['status'] in ['Closed', 'Resolved'])
    critical = sum(1 for t in tickets if t['priority'] == 'Critical' and t['status'] not in ['Closed', 'Resolved'])

    statuses = {}
    for t in tickets:
        s = t['status']
        statuses[s] = statuses.get(s, 0) + 1

    priorities = {}
    for t in tickets:
        p = t['priority']
        priorities[p] = priorities.get(p, 0) + 1

    request_types = {}
    for t in tickets:
        rt = t['request_type']
        if rt and rt != 'nan':
            request_types[rt] = request_types.get(rt, 0) + 1

    staff_workload = {}
    for t in tickets:
        staff = t['staff_assigned']
        if staff and staff != 'nan':
            if staff not in staff_workload:
                staff_workload[staff] = {'assigned': 0, 'open': 0}
            staff_workload[staff]['assigned'] += 1
            if t['status'] not in ['Closed', 'Resolved']:
                staff_workload[staff]['open'] += 1

    return {
        'total': total,
        'open': open_count,
        'closed': closed_count,
        'critical': critical,
        'statuses': statuses,
        'priorities': priorities,
        'request_types': request_types,
        'staff_workload': staff_workload
    }

def get_dropdown_options():
    """Get unique values for dropdown fields."""
    tickets = read_tickets_from_db()
    return {
        'request_types': sorted(set(t['request_type'] for t in tickets if t['request_type'] and t['request_type'] != 'nan')),
        'staff': sorted(set(t['staff_assigned'] for t in tickets if t['staff_assigned'] and t['staff_assigned'] != 'nan')),
        'requesters': sorted(set(t['requester'] for t in tickets if t['requester'] and t['requester'] != 'nan'))
    }

# --- EXCEL EXPORT FUNCTION ---
def generate_excel_from_db():
    """Generate an Excel file from database tickets."""
    tickets = read_tickets_from_db()

    wb = Workbook()
    ws = wb.active
    ws.title = "IT Service Tickets"

    # Header row
    headers = ['Ticket ID', 'Title', 'Status', 'Priority', 'Request Type',
               'Staff Assigned', 'Requester', 'Date Opened', 'Days Open',
               'Description', 'Resolution Notes']
    ws.append(headers)

    # Data rows
    for t in tickets:
        days_open = t['days_open'] if t['days_open'] != -1 else '-'
        ws.append([
            t['ticket_id'],
            t['title'],
            t['status'],
            t['priority'],
            t['request_type'],
            t['staff_assigned'],
            t['requester'],
            t['date_opened'],
            days_open,
            t['description'],
            t['resolution_notes']
        ])

    # Auto-adjust column widths
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column].width = adjusted_width

    return wb

# --- ROUTES ---
@app.route('/login')
def login_page():
    if 'user' in session:
        return redirect('/')
    return render_template('login.html')

@app.route('/api/login', methods=['POST'])
def api_login():
    data = request.json
    username = data.get('username', '').strip()
    password = data.get('password', '')

    user = get_user(username)
    if user and verify_password(password, user['password_hash']):
        session['user'] = {
            'username': user['username'],
            'role': user['role'],
            'display_name': user['display_name']
        }
        return jsonify({'status': 'success', 'user': session['user']})
    return jsonify({'error': 'Invalid credentials'}), 401

@app.route('/api/logout', methods=['POST'])
def api_logout():
    session.pop('user', None)
    return jsonify({'status': 'success'})

@app.route('/api/me')
def api_me():
    if 'user' in session:
        return jsonify(session['user'])
    return jsonify({'error': 'Not authenticated'}), 401

@app.route('/api/users', methods=['POST'])
@admin_required
def api_create_user():
    data = request.json
    username = data.get('username', '').strip()
    password = data.get('password', '')
    role = data.get('role', 'viewer')
    display_name = data.get('display_name', username)

    if not username or not password:
        return jsonify({'error': 'Username and password required'}), 400

    if get_user(username):
        return jsonify({'error': 'Username already exists'}), 400

    users = load_users()
    users.append({
        'username': username,
        'password_hash': hash_password(password),
        'role': role,
        'display_name': display_name
    })
    save_users(users)
    return jsonify({'status': 'success'}), 201

@app.route('/api/users', methods=['GET'])
@admin_required
def api_list_users():
    users = load_users()
    return jsonify([{
        'username': u['username'],
        'role': u['role'],
        'display_name': u['display_name']
    } for u in users])

@app.route('/api/users/<username>', methods=['PUT'])
@admin_required
def api_update_user(username):
    data = request.json
    users = load_users()

    for user in users:
        if user['username'] == username:
            if 'role' in data:
                user['role'] = data['role']
            if 'display_name' in data:
                user['display_name'] = data['display_name']
            if 'password' in data and data['password']:
                user['password_hash'] = hash_password(data['password'])
            save_users(users)
            return jsonify({'status': 'success'})
    return jsonify({'error': 'User not found'}), 404

@app.route('/api/users/<username>', methods=['DELETE'])
@admin_required
def api_delete_user(username):
    if username == 'admin':
        return jsonify({'error': 'Cannot delete admin user'}), 400

    users = load_users()
    users = [u for u in users if u['username'] != username]
    save_users(users)
    return jsonify({'status': 'success'})

@app.route('/api/change-password', methods=['POST'])
@login_required
def api_change_password():
    data = request.json
    current_password = data.get('current_password', '')
    new_password = data.get('new_password', '')

    if not new_password:
        return jsonify({'error': 'New password required'}), 400

    users = load_users()
    username = session['user']['username']

    for user in users:
        if user['username'] == username:
            if not verify_password(current_password, user['password_hash']):
                return jsonify({'error': 'Current password incorrect'}), 401
            user['password_hash'] = hash_password(new_password)
            save_users(users)
            return jsonify({'status': 'success'})
    return jsonify({'error': 'User not found'}), 404

@app.route('/')
def index():
    if 'user' not in session:
        return redirect('/login')
    return render_template('dashboard.html')

@app.route('/api/stats')
def api_stats():
    tickets = read_tickets_from_db()
    stats = calculate_stats(tickets)
    return jsonify(stats)

@app.route('/api/options')
def api_options():
    return jsonify(get_dropdown_options())

@app.route('/api/tickets')
def api_tickets():
    tickets = read_tickets_from_db()
    # Convert -1 sentinel to '-' for closed/resolved tickets
    for t in tickets:
        if t['days_open'] == -1:
            t['days_open'] = '-'
    return jsonify(tickets)

@app.route('/api/tickets', methods=['POST'])
def create_ticket():
    try:
        data = request.json
        ticket_id = create_ticket_in_db(data)
        if ticket_id:
            return jsonify({'status': 'success', 'ticket_id': ticket_id}), 201
        return jsonify({'error': 'Failed to save'}), 500
    except Exception as e:
        print(f"[API] Error creating ticket: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/tickets/<ticket_id>', methods=['PUT'])
def update_ticket(ticket_id):
    try:
        data = request.json
        if update_ticket_in_db(ticket_id, data):
            return jsonify({'status': 'success'})
        return jsonify({'error': 'Ticket not found'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/tickets/<ticket_id>', methods=['DELETE'])
def delete_ticket(ticket_id):
    try:
        if delete_ticket_from_db(ticket_id):
            return jsonify({'status': 'success'})
        return jsonify({'error': 'Ticket not found'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/refresh', methods=['POST'])
def api_refresh():
    return jsonify({'status': 'success', 'message': 'Data refreshed from database'})

@app.route('/api/export')
def api_export():
    """Export tickets to Excel file."""
    try:
        wb = generate_excel_from_db()
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        filename = f"tickets_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        print(f"[API] Error exporting: {e}")
        return jsonify({'error': str(e)}), 500

# Startup diagnostics
def print_diagnostics():
    """Print startup diagnostics."""
    print("\n" + "="*50)
    print(" IT OPSCENTER - STARTUP DIAGNOSTICS")
    print("="*50)

    # Database check
    print("\n[DATABASE]")
    if DATABASE_URL:
        print(f"  DATABASE_URL: {'*' * 20}...{DATABASE_URL[-20:]}")
        print("  Initializing database...")
        if init_database():
            print("  Database: OK")
        else:
            print("  Database: FAILED")
    else:
        print("  DATABASE_URL: NOT SET - using Excel fallback")

    # Auth check
    print("\n[AUTHENTICATION]")
    for path in USERS_FILE_LOCATIONS:
        exists = os.path.exists(path)
        status = "FOUND" if exists else "not found"
        print(f"  {path}: {status}")

    users_file = get_users_file()
    print(f"\n  Active users file: {users_file}")

    users = load_users()
    if users:
        print(f"  Users loaded: {len(users)}")
        for u in users:
            print(f"    - {u.get('username')} ({u.get('role')})")
    else:
        print("  WARNING: No users loaded! Login will fail.")

    print("="*50 + "\n")

# Run diagnostics on startup
print_diagnostics()

if __name__ == '__main__':
    print("\n" + "="*50)
    print(" IT OPERATIONS CENTER - DASHBOARD")
    print(" Open browser to: http://127.0.0.1:5000")
    print("="*50 + "\n")
    app.run(debug=True, port=5000)
