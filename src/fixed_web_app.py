from flask import Flask, render_template_string, request, send_file, redirect, url_for, session, abort, flash
import io
import csv
from datetime import datetime, timedelta
import os
import sys
import pandas as pd
import holidays
import json
import bcrypt
from functools import wraps
import re
from werkzeug.security import generate_password_hash, check_password_hash

# Import parser classes
sys.path.append(os.path.dirname(__file__))
from job_parser_core import JobParser, BC04Parser

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10MB upload limit

# Delivery date calculation logic
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
    return {}  # Return an empty dictionary, not a list

def save_users(users):
    with open(USERS_FILE, 'w', encoding='utf-8') as f:
        json.dump(users, f, ensure_ascii=False, indent=2)

def hash_password(password):
    return bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')

def check_password(password, hashed):
    return bcrypt.checkpw(password.encode('utf-8'), hashed.encode('utf-8'))

def get_user(username):
    users = load_users()
    if username in users:
        user_data = users[username]
        user_data['username'] = username
        return user_data
    return None

def add_user(username, password, enabled=True):
    users = load_users()
    if username in users:
        return False  # User already exists
    users[username] = {'password': hash_password(password), 'enabled': enabled}
    save_users(users)
    return True

def set_user_enabled(username, enabled):
    users = load_users()
    if username not in users:
        return False
    users[username]['enabled'] = enabled
    save_users(users)
    return True

def set_user_password(username, password):
    users = load_users()
    if username not in users:
        return False
    users[username]['password'] = hash_password(password)
    save_users(users)
    return True

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
    # Create a new user with bcrypt hash for '301103'
    add_user("bradlakin1", "301103", enabled=True)

# Rest of the file remains the same
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

# Include the rest of your web_app.py code here
# ...

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
        # Process form submission
        pass  # Your existing code here

    return render_template_string("Your template here", job_type=job_type, job_data=job_data, collection_date=collection_date, delivery_date=delivery_date, error=error, debug=debug, job_history=job_history, username=session.get('username'))

def is_admin():
    return session.get('username') in ['admin', 'bradlakin1']

@app.route('/admin', methods=['GET', 'POST'])
@login_required
def admin_panel():
    if not is_admin():
        abort(403)
    # Admin panel code here
    return "Admin panel"

@app.route('/static/<path:filename>')
def static_files(filename):
    return send_file(os.path.join(os.path.dirname(__file__), 'static', filename))

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=int(os.environ.get('PORT', 5000))) 