import os
import json
import hashlib
import secrets
from functools import wraps
from flask import Flask, render_template, jsonify, request, session, redirect, url_for
from datetime import datetime, time
from openpyxl import load_workbook

app = Flask(__name__)
# Use fixed secret key for production/Docker (allows session persistence across workers)
# In development, you can set SECRET_KEY env var or use the fallback
app.secret_key = os.environ.get('SECRET_KEY', 'ticketing-dashboard-secret-key-2025')
EXCEL_FILE = 'tickets.xlsx'
SHEET_NAME = 'IT Service Tickets'

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

# --- USER MANAGEMENT ---
def load_users():
    """Load users from JSON file."""
    try:
        users_path = get_users_file()
        with open(users_path, 'r') as f:
            return json.load(f).get('users', [])
    except:
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

def read_excel_data():
    """Reads all necessary data directly from the Excel file using openpyxl."""
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb[SHEET_NAME]
        
        tickets = []
        for row in range(2, 1000):  # Start from row 2 (after header)
            title = ws[f'B{row}'].value
            if not title:  # Empty row = end of data
                break
            
            # Compute Ticket ID from row number: IT-2025XXXX where XXXX = row - 1
            ticket_id = f"IT-2025{row-1:04d}"
            
            # Read other fields
            status = ws[f'C{row}'].value or ''
            priority = ws[f'D{row}'].value or ''
            request_type = ws[f'E{row}'].value or ''
            staff_assigned = ws[f'F{row}'].value or ''
            requester = ws[f'G{row}'].value or ''
            date_val = ws[f'H{row}'].value
            # days_open is calculated dynamically, not read from Excel
            description = ws[f'J{row}'].value or ''
            resolution = ws[f'K{row}'].value or ''
            
            # Handle date conversion
            date_obj = None
            if date_val is None:
                date_str = ''
            elif isinstance(date_val, time):
                date_str = datetime.today().strftime('%Y-%m-%d')
                date_obj = datetime.today()
            elif hasattr(date_val, 'strftime'):
                date_str = date_val.strftime('%Y-%m-%d')
                # Convert to datetime if it's a date object
                if hasattr(date_val, 'hour'):
                    date_obj = date_val
                else:
                    date_obj = datetime.combine(date_val, datetime.min.time())
            else:
                date_str = str(date_val)
                try:
                    date_obj = datetime.strptime(date_str, '%Y-%m-%d')
                except:
                    date_obj = None
            
            # Calculate days_open dynamically from date_opened
            # For Closed/Resolved tickets, show "-" (we'll use -1 as sentinel)
            if status.strip() in ['Closed', 'Resolved']:
                days_open_int = -1  # Sentinel for closed tickets
            elif date_obj:
                days_open_int = (datetime.now() - date_obj).days
                if days_open_int < 0:
                    days_open_int = 0
            else:
                days_open_int = 0
            
            tickets.append({
                'ticket_id': ticket_id,
                'title': str(title).strip(),
                'status': str(status).strip(),
                'priority': str(priority).strip(),
                'request_type': str(request_type).strip(),
                'staff_assigned': str(staff_assigned).strip(),
                'requester': str(requester).strip(),
                'date_opened': date_str,
                'days_open': days_open_int,
                'description': str(description).strip(),
                'resolution_notes': str(resolution).strip() if resolution else ''
            })
        
        wb.close()
        return tickets
    except Exception as e:
        print(f"Error reading Excel: {e}")
        import traceback
        traceback.print_exc()
        return []

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

# --- EXCEL WRITE FUNCTIONS ---
def get_next_ticket_id():
    """Generate the next ticket ID."""
    tickets = read_excel_data()
    if not tickets:
        return 'IT-20250001'
    ids = [int(t['ticket_id'].replace('IT-', '')) for t in tickets if t['ticket_id'].startswith('IT-')]
    return f"IT-{max(ids) + 1:08d}"

def find_first_empty_row(ws):
    """Find the first row where column B (Title) is empty."""
    for row in range(2, 1000):  # Start from row 2 (after header)
        if not ws[f'B{row}'].value:
            return row
    return ws.max_row + 1

def append_to_excel(ticket_data):
    """Append a new ticket to the Excel file."""
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb[SHEET_NAME]
        
        # Find first empty row (where Title is empty)
        new_row = find_first_empty_row(ws)
        print(f"Appending to row {new_row}")
        
        # Write data to empty row - DON'T overwrite column A (has formula)
        ws[f'B{new_row}'] = ticket_data.get('title', '')
        ws[f'C{new_row}'] = ticket_data.get('status', 'Open')
        ws[f'D{new_row}'] = ticket_data.get('priority', 'Low')
        ws[f'E{new_row}'] = ticket_data.get('request_type', '')
        ws[f'F{new_row}'] = ticket_data.get('staff_assigned', '')
        ws[f'G{new_row}'] = ticket_data.get('requester', '')
        ws[f'H{new_row}'] = datetime.now()  # Date as datetime object
        ws[f'J{new_row}'] = ticket_data.get('description', '')
        ws[f'K{new_row}'] = ticket_data.get('resolution_notes', '')
        
        wb.save(EXCEL_FILE)
        return True
    except Exception as e:
        print(f"Error appending to Excel: {e}")
        import traceback
        traceback.print_exc()
        return False

def find_ticket_row(ws, ticket_id):
    """Find the row number for a ticket from its ID."""
    # Ticket ID format: IT-2025XXXX where XXXX is row-1
    # So IT-20250001 = row 2, IT-20250002 = row 3, etc.
    try:
        # Extract the number part (last 4 digits)
        num_part = int(ticket_id[-4:])
        row = num_part + 1  # row = XXXX + 1
        # Verify the row has data
        if ws[f'B{row}'].value:
            return row
    except:
        pass
    return None

def update_excel_row(ticket_id, ticket_data):
    """Update a specific ticket row in Excel."""
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb[SHEET_NAME]
        
        row = find_ticket_row(ws, ticket_id)
        if row is None:
            return False
        
        if 'title' in ticket_data: ws[f'B{row}'] = ticket_data['title']
        if 'status' in ticket_data: ws[f'C{row}'] = ticket_data['status']
        if 'priority' in ticket_data: ws[f'D{row}'] = ticket_data['priority']
        if 'request_type' in ticket_data: ws[f'E{row}'] = ticket_data['request_type']
        if 'staff_assigned' in ticket_data: ws[f'F{row}'] = ticket_data['staff_assigned']
        if 'requester' in ticket_data: ws[f'G{row}'] = ticket_data['requester']
        if 'description' in ticket_data: ws[f'J{row}'] = ticket_data['description']
        if 'resolution_notes' in ticket_data: ws[f'K{row}'] = ticket_data['resolution_notes']
        wb.save(EXCEL_FILE)
        return True
    except Exception as e:
        print(f"Error updating Excel: {e}")
        import traceback
        traceback.print_exc()
        return False

def delete_excel_row(ticket_id):
    """Delete a ticket row from Excel by clearing its contents."""
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb[SHEET_NAME]
        
        row = find_ticket_row(ws, ticket_id)
        if row is None:
            return False
        
        # Clear the row contents (don't delete row to preserve formulas)
        for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
            ws[f'{col}{row}'] = None
        wb.save(EXCEL_FILE)
        return True
    except Exception as e:
        print(f"Error deleting from Excel: {e}")
        import traceback
        traceback.print_exc()
        return False

def get_dropdown_options():
    """Get unique values for dropdown fields."""
    tickets = read_excel_data()
    return {
        'request_types': sorted(set(t['request_type'] for t in tickets if t['request_type'] and t['request_type'] != 'nan')),
        'staff': sorted(set(t['staff_assigned'] for t in tickets if t['staff_assigned'] and t['staff_assigned'] != 'nan')),
        'requesters': sorted(set(t['requester'] for t in tickets if t['requester'] and t['requester'] != 'nan'))
    }

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
    # Don't return password hashes
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
    tickets = read_excel_data()
    stats = calculate_stats(tickets)
    return jsonify(stats)

@app.route('/api/options')
def api_options():
    return jsonify(get_dropdown_options())

@app.route('/api/tickets')
def api_tickets():
    tickets = read_excel_data()
    tickets.sort(key=lambda x: x['date_opened'], reverse=True)
    # Convert -1 sentinel to '-' for closed/resolved tickets
    for t in tickets:
        if t['days_open'] == -1:
            t['days_open'] = '-'
    return jsonify(tickets)

@app.route('/api/tickets', methods=['POST'])
def create_ticket():
    try:
        data = request.json
        data['ticket_id'] = get_next_ticket_id()
        data['date_opened'] = datetime.now().strftime('%Y-%m-%d')
        if append_to_excel(data):
            return jsonify({'status': 'success', 'ticket_id': data['ticket_id']}), 201
        return jsonify({'error': 'Failed to save'}), 500
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/tickets/<ticket_id>', methods=['PUT'])
def update_ticket(ticket_id):
    try:
        data = request.json
        if update_excel_row(ticket_id, data):
            return jsonify({'status': 'success'})
        return jsonify({'error': 'Ticket not found'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/tickets/<ticket_id>', methods=['DELETE'])
def delete_ticket(ticket_id):
    try:
        if delete_excel_row(ticket_id):
            return jsonify({'status': 'success'})
        return jsonify({'error': 'Ticket not found'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/refresh', methods=['POST'])
def api_refresh():
    return jsonify({'status': 'success', 'message': 'Data refreshed from Excel'})

if __name__ == '__main__':
    print("\n" + "="*50)
    print(" IT OPERATIONS CENTER - DASHBOARD")
    print(f" Reading from: {EXCEL_FILE}")
    print(" Open browser to: http://127.0.0.1:5000")
    print("="*50 + "\n")
    app.run(debug=True, port=5000)