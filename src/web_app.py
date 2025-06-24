import os
import json
from flask import Flask, render_template_string, request, send_file, redirect, url_for, session, abort
import io
import csv
from datetime import datetime, timedelta
import pandas as pd
import holidays
import bcrypt
from functools import wraps
import re

# Import parser classes
from job_parser_core import JobParser, BC04Parser

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10MB upload limit

uk_holidays = holidays.UK()

HISTORY_FILE = os.path.join(os.path.dirname(__file__), 'job_history.json')
USERS_FILE = os.path.join(os.path.dirname(__file__), 'users.json')
SECRET_KEY = 'REPLACE_THIS_WITH_A_RANDOM_SECRET_KEY'
app.secret_key = SECRET_KEY

def load_job_history():
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return []

def save_job_history(history):
    with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
        json.dump(history, f, ensure_ascii=False, indent=2)

def load_users():
    if os.path.exists(USERS_FILE):
        with open(USERS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return []

def save_users(users):
    with open(USERS_FILE, 'w', encoding='utf-8') as f:
        json.dump(users, f, ensure_ascii=False, indent=2)

def hash_password(password):
    return bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')

def check_password(password, hashed):
    return bcrypt.checkpw(password.encode('utf-8'), hashed.encode('utf-8'))

def get_user(username):
    users = load_users()
    for user in users:
        if user['username'] == username:
            return user
    return None

def add_user(username, password, enabled=True):
    users = load_users()
    if get_user(username):
        return False  # User already exists
    hashed = hash_password(password)
    users.append({'username': username, 'password': hashed, 'enabled': enabled})
    save_users(users)
    return True

def set_user_enabled(username, enabled):
    users = load_users()
    for user in users:
        if user['username'] == username:
            user['enabled'] = enabled
            save_users(users)
            return True
    return False

def set_user_password(username, password):
    users = load_users()
    for user in users:
        if user['username'] == username:
            user['password'] = hash_password(password)
            save_users(users)
            return True
    return False

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'username' not in session or not session.get('username') or not (get_user(session['username']) and get_user(session['username']).get('enabled', False)):
            return redirect(url_for('login', next=request.path))
        return f(*args, **kwargs)
    return decorated

job_history = load_job_history()

# Ensure at least one user exists
if not load_users():
    # Use the bcrypt hash generated earlier for '301103'
    users = [
        {
            "username": "bradlakin1",
            "password": "$2b$12$CZInzpyaBYPEPMR4aGTLW.NNpcCJPX4lYxRKqiabU3AQk6ma/jnZG",
            "enabled": True
        }
    ]
    save_users(users)

def calculate_delivery_date_ac01(collection_date_str):
    collection_date = datetime.strptime(collection_date_str, "%d/%m/%Y")
    current_date = collection_date
    business_days = 0
    while business_days < 3:
        current_date += timedelta(days=1)
        if current_date.weekday() < 5 and current_date not in uk_holidays:
            business_days += 1
    return current_date.strftime("%d/%m/%Y")

def calculate_delivery_date_bc04(collection_date_str):
    collection_date = datetime.strptime(collection_date_str, "%d/%m/%Y")
    delivery = collection_date + timedelta(days=1)
    while delivery.weekday() >= 5 or delivery in uk_holidays:
        delivery += timedelta(days=1)
    return delivery.strftime("%d/%m/%Y")

TEMPLATE = '''
... (keep your HTML template as is) ...
'''

def normalize_line_endings(text):
    return text.replace('\r\n', '\n').replace('\r', '\n')

@app.route('/auto_delivery_date', methods=['POST'])
def auto_delivery_date():
    job_type = request.form.get('job_type')
    collection_date = request.form.get('collection_date')
    if not collection_date:
        return ''
    if job_type == 'AC01':
        return calculate_delivery_date_ac01(collection_date)
    elif job_type == 'BC04':
        return calculate_delivery_date_bc04(collection_date)
    else:
        return collection_date

@app.route('/login', methods=['GET', 'POST'])
def login():
    error = None
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        user = get_user(username)
        if user and user.get('enabled', False) and check_password(password, user['password']):
            session['username'] = username
            return redirect(url_for('index'))
        error = 'Invalid credentials or account disabled.'
    return render_template_string('''
    <html><head><title>Login</title><style>body{background:#e0f2e9;font-family:sans-serif;} .login-box{background:#fff;max-width:400px;margin:80px auto;padding:40px 32px 32px 32px;border-radius:14px;box-shadow:0 4px 24px #1b6e3a22;} h2{color:#1b6e3a;} label{font-weight:600;} input{width:100%;padding:12px;margin:8px 0 18px 0;border-radius:7px;border:1.5px solid #bfc7d1;font-size:1.08em;} button{background:#1b6e3a;color:#fff;border:none;padding:14px 0;border-radius:8px;font-size:1.1em;width:100%;font-weight:700;box-shadow:0 2px 8px #1b6e3a22;} .error{color:#d00;margin-bottom:12px;}</style></head><body><div class="login-box"><h2>Login</h2>{% if error %}<div class="error">{{ error }}</div>{% endif %}<form method="POST"><label>Username:</label><input name="username" required><label>Password:</label><input name="password" type="password" required><button type="submit">Login</button></form></div></body></html>
    ''', error=error)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/history/<path:filename>')
@login_required
def protected_history_file(filename):
    static_history_dir = os.path.join(os.path.dirname(__file__), 'static', 'history')
    file_path = os.path.join(static_history_dir, filename)
    if not os.path.exists(file_path):
        abort(404)
    return send_file(file_path, as_attachment=True)

@app.route('/', methods=['GET', 'POST'])
@login_required
def index():
    global job_history
    error = None
    debug = None
    job_type = request.form.get('job_type', 'AC01')
    job_data = request.form.get('job_data', '')
    collection_date = request.form.get('collection_date', datetime.now().strftime('%d/%m/%Y'))
    delivery_date = request.form.get('delivery_date', '')

    # Auto-set delivery date if not provided
    if not delivery_date:
        if job_type == 'AC01':
            delivery_date = calculate_delivery_date_ac01(collection_date)
        elif job_type == 'BC04':
            delivery_date = calculate_delivery_date_bc04(collection_date)
        else:
            delivery_date = collection_date

    if request.method == 'POST':
        if job_type in ['AC01', 'BC04', 'EU01']:
            job_data_norm = normalize_line_endings(job_data)
            if job_type == 'AC01' or job_type == 'EU01':
                parser = JobParser(collection_date, delivery_date)
            elif job_type == 'BC04':
                parser = BC04Parser(collection_date, delivery_date)
            else:
                parser = None
            jobs = parser.parse_jobs(job_data_norm) if parser else []
            if not jobs:
                debug = f"<b>Debug:</b><br>Input preview (first 500 chars):<br><pre>{job_data_norm[:500]}</pre><br>Jobs found: 0"
                error = "No valid jobs found. Please check your input format."
            else:
                output = io.StringIO()
                writer = csv.DictWriter(output, fieldnames=jobs[0].keys())
                writer.writeheader()
                writer.writerows(jobs)
                output.seek(0)
                # Save CSV to static/history with timestamp
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                csv_filename = f"history_{job_type}_{timestamp}.csv"
                static_history_dir = os.path.join(os.path.dirname(__file__), 'static', 'history')
                os.makedirs(static_history_dir, exist_ok=True)
                csv_path = os.path.join(static_history_dir, csv_filename)
                with open(csv_path, 'w', encoding='utf-8', newline='') as f:
                    f.write(output.getvalue())
                # Add to job history (user is placeholder for now)
                job_history.insert(0, {
                    'timestamp': timestamp,
                    'job_type': job_type,
                    'csv_path': f'history/{csv_filename}',
                    'user': session.get('username')
                })
                save_job_history(job_history)
                return send_file(io.BytesIO(output.getvalue().encode('utf-8')), mimetype='text/csv', as_attachment=True, download_name=f'{job_type}_jobs_{timestamp}.csv')
        elif job_type in ['GR11', 'CW09']:
            file = request.files.get('file')
            if not file:
                error = "Please upload an Excel or CSV file."
            else:
                try:
                    df = pd.read_excel(file) if file.filename.endswith('.xlsx') else pd.read_csv(file)
                    # ... (rest of your GR11/CW09 logic unchanged) ...
                except Exception as e:
                    error = f"Failed to process file: {e}"
    # List all CSVs in static/history for job history
    static_history_dir = os.path.join(os.path.dirname(__file__), 'static', 'history')
    if os.path.exists(static_history_dir):
        files = sorted(os.listdir(static_history_dir), reverse=True)
        job_history = [row for row in job_history if os.path.exists(os.path.join(static_history_dir, row['csv_path'].split('/')[-1]))]
    else:
        files = []
    return render_template_string(TEMPLATE, job_type=job_type, job_data=job_data, collection_date=collection_date, delivery_date=delivery_date, error=error, debug=debug, job_history=job_history, username=session.get('username'))

def is_admin():
    return session.get('username') in ['admin', 'bradlakin1']

@app.route('/admin', methods=['GET', 'POST'])
@login_required
def admin_panel():
    if not is_admin():
        abort(403)
    msg = None
    users = load_users()
    if request.method == 'POST':
        action = request.form.get('action')
        username = request.form.get('username', '').strip()
        if action == 'add':
            password = request.form.get('password', '')
            if get_user(username):
                msg = f'User {username} already exists.'
            else:
                add_user(username, password, enabled=True)
                msg = f'User {username} added and enabled.'
        elif action == 'enable':
            if set_user_enabled(username, True):
                msg = f'User {username} enabled.'
        elif action == 'disable':
            if set_user_enabled(username, False):
                msg = f'User {username} disabled.'
        elif action == 'setpw':
            password = request.form.get('password', '')
            if set_user_password(username, password):
                msg = f'Password updated for {username}.'
    # For template, convert list of users to dict for easier rendering
    users_dict = {u['username']: type('obj', (), u) for u in users}
    return render_template_string('''
    <html><head><title>Admin Panel</title><style>body{background:#e0f2e9;font-family:sans-serif;} .admin-box{background:#fff;max-width:600px;margin:40px auto;padding:40px 32px 32px 32px;border-radius:14px;box-shadow:0 4px 24px #1b6e3a22;} h2{color:#1b6e3a;} table{width:100%;border-collapse:collapse;margin-bottom:24px;} th,td{border:1px solid #bfc7d1;padding:8px 10px;} th{background:#f5f7fa;} tr:nth-child(even){background:#f7f9fc;} .btn{background:#1b6e3a;color:#fff;border:none;padding:6px 16px;border-radius:6px;font-size:1em;font-weight:600;margin:0 2px;} .btn:disabled{background:#bfc7d1;} .msg{color:#1b6e3a;margin-bottom:12px;font-weight:600;} .form-row{margin-bottom:18px;} label{font-weight:600;}</style></head><body><div class="admin-box"><h2>User Management</h2>{% if msg %}<div class="msg">{{ msg }}</div>{% endif %}<table><tr><th>Username</th><th>Status</th><th>Actions</th></tr>{% for u, v in users.items() %}<tr><td>{{ u }}</td><td>{{ 'ENABLED' if v.enabled else 'DISABLED' }}</td><td><form method="post" style="display:inline"><input type="hidden" name="username" value="{{ u }}"><button class="btn" name="action" value="enable" {% if v.enabled %}disabled{% endif %}>Enable</button><button class="btn" name="action" value="disable" {% if not v.enabled %}disabled{% endif %}>Disable</button></form><form method="post" style="display:inline"><input type="hidden" name="username" value="{{ u }}"><input type="text" name="password" placeholder="New password" required style="width:110px;"><button class="btn" name="action" value="setpw">Set Password</button></form></td></tr>{% endfor %}</table><h3>Add New User</h3><form method="post"><div class="form-row"><label>Username:</label><input name="username" required></div><div class="form-row"><label>Password:</label><input name="password" type="password" required></div><button class="btn" name="action" value="add">Add User</button></form><div style="margin-top:24px;"><a href="/">Back to main</a></div></div></body></html>
    ''', users=users_dict, msg=msg)
