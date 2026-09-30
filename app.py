from flask import Flask, render_template, request, jsonify, send_file, session, redirect, url_for

from werkzeug.utils import secure_filename

from openpyxl import Workbook, load_workbook

from openpyxl.styles import Font, PatternFill, Alignment

import os

import shutil

from datetime import datetime

import json

from functools import wraps

import hashlib

import hmac

from io import BytesIO

from reportlab.pdfgen import canvas

from reportlab.lib.pagesizes import A4

from reportlab.lib.utils import ImageReader

from dotenv import load_dotenv



load_dotenv()



app = Flask(__name__)

app.secret_key = os.getenv('SECRET_KEY', 'change-this-secret-key-in-production')

app.config['SESSION_PERMANENT'] = False

app.config['SESSION_COOKIE_HTTPONLY'] = True

app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'



EXCEL_DIR = 'excel_sheets'

MONTHS_FILE = 'months_data.json'

BACKUP_DIR = 'backups'

DELETED_MONTHS_DIR = os.path.join(BACKUP_DIR, 'deleted_months')

BACKUP_MONTHS_FILE = os.path.join(BACKUP_DIR, 'backup_months.json')

CHEQUE_DEPOSITS_FILE = 'cheque_deposits.json'

CHEQUE_UPLOAD_DIR = 'cheque_slips'

CHEQUE_SECTION_MARKER = 'CHEQUE DEPOSITS'

ALLOWED_CHEQUE_EXTENSIONS = {'png', 'jpg', 'jpeg', 'webp'}

app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024





def load_cheque_deposits():

    if os.path.exists(CHEQUE_DEPOSITS_FILE):

        try:

            with open(CHEQUE_DEPOSITS_FILE, 'r', encoding='utf-8') as f:

                data = json.load(f)

                return data if isinstance(data, dict) else {}

        except Exception:

            return {}

    return {}





def save_cheque_deposits(data):

    with open(CHEQUE_DEPOSITS_FILE, 'w', encoding='utf-8') as f:

        json.dump(data, f, indent=2, ensure_ascii=False)





def get_month_cheque_deposits(month_name):

    if not month_name:

        return []

    deposits = load_cheque_deposits()

    rows = [item for item in deposits.values() if item.get('month') == month_name]

    rows.sort(key=lambda x: (str(x.get('date', '')), str(x.get('time', '')), str(x.get('created_at', ''))))

    return rows





def next_cheque_deposit_id(month_name):

    prefix = 'CHQ-' + datetime.now().strftime('%Y%m%d')

    deposits = load_cheque_deposits()

    used = []

    for dep_id in deposits.keys():

        if dep_id.startswith(prefix + '-'):

            try:

                used.append(int(dep_id.rsplit('-', 1)[1]))

            except Exception:

                pass

    return f"{prefix}-{max(used, default=0) + 1:03d}"





def cheque_file_allowed(filename):

    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_CHEQUE_EXTENSIONS





def find_cheque_section_row(ws):

    for row_no in range(1, ws.max_row + 1):

        if str(ws.cell(row_no, 1).value or '').strip().upper() == CHEQUE_SECTION_MARKER:

            return row_no

    return None





def remove_cheque_section(ws):

    marker_row = find_cheque_section_row(ws)

    if marker_row:

        ws.delete_rows(marker_row, ws.max_row - marker_row + 1)





def sync_cheque_deposits_to_excel(month_name, include_all=False):

    """Write cheque deposits below normal transactions in the SAME Records worksheet."""

    filename = os.path.join(EXCEL_DIR, f"{month_name}.xlsx")

    if not os.path.exists(filename):

        return False



    wb = load_workbook(filename)

    if 'Cheque Deposits' in wb.sheetnames:

        del wb['Cheque Deposits']

    ws = wb['Records'] if 'Records' in wb.sheetnames else wb.active

    remove_cheque_section(ws)



    deposits = get_month_cheque_deposits(month_name)

    if not include_all:

        deposits = [d for d in deposits if d.get('excel_added')]



    if deposits:

        ws.append([None])

        marker_row = ws.max_row + 1

        ws.append([CHEQUE_SECTION_MARKER])

        ws.merge_cells(start_row=marker_row, start_column=1, end_row=marker_row, end_column=6)

        title_cell = ws.cell(marker_row, 1)

        title_cell.fill = PatternFill(start_color='1A1A2E', end_color='1A1A2E', fill_type='solid')

        title_cell.font = Font(bold=True, color='FFFFFF', size=12)

        title_cell.alignment = Alignment(horizontal='left')



        headers = ['No', 'Date', 'Time', 'Deposit ID', 'Amount', 'Notes']

        ws.append(headers)

        header_row = ws.max_row

        header_fill = PatternFill(start_color='D7193F', end_color='D7193F', fill_type='solid')

        for col in range(1, 7):

            cell = ws.cell(header_row, col)

            cell.fill = header_fill

            cell.font = Font(bold=True, color='FFFFFF')

            cell.alignment = Alignment(horizontal='center')



        for idx, dep in enumerate(deposits, start=1):

            ws.append([

                idx,

                dep.get('date', ''),

                dep.get('time', ''),

                dep.get('deposit_id', ''),

                parse_amount(dep.get('amount', 0)),

                dep.get('notes', '')

            ])

            ws.cell(ws.max_row, 5).number_format = '#,##0.00'



        widths = {'A': 7, 'B': 14, 'C': 12, 'D': 26, 'E': 16, 'F': 42}

        for col, width in widths.items():

            current = ws.column_dimensions[col].width or 0

            ws.column_dimensions[col].width = max(current, width)



    wb.save(filename)

    wb.close()

    return True





def hash_password(password):

    salt = os.getenv('PASSWORD_SALT', 'uta-default-salt-change-me')

    return hashlib.sha256((salt + str(password)).encode('utf-8')).hexdigest()



DEFAULT_USERS = {

    'thanzeel': hash_password(os.getenv('THANZEEL_PASSWORD', 'utams2179')),

    'kaiff': hash_password(os.getenv('KAIFF_PASSWORD', 'utams2179')),

}



def load_users():

    raw = os.getenv('UTA_USERS_JSON')

    if not raw:

        return DEFAULT_USERS

    try:

        plain_users = json.loads(raw)

        return {u.lower(): hash_password(p) for u, p in plain_users.items()}

    except Exception:

        return DEFAULT_USERS



USERS = load_users()



# Candidate Management System access

CANDIDATE_USERS = {'kaiff', 'firnas', 'thanzeel'}

CANDIDATE_PASSWORD = os.getenv('CANDIDATE_PASSWORD', '').strip()



CANDIDATE_DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'candidate_data')

CANDIDATE_INVOICES_FILE = os.path.join(CANDIDATE_DATA_DIR, 'invoices.json')



def load_candidate_invoices():

    os.makedirs(CANDIDATE_DATA_DIR, exist_ok=True)

    if not os.path.exists(CANDIDATE_INVOICES_FILE):

        return []

    try:

        with open(CANDIDATE_INVOICES_FILE, 'r', encoding='utf-8') as f:

            data = json.load(f)

            return data if isinstance(data, list) else []

    except Exception:

        return []



def save_candidate_invoices(invoices):

    os.makedirs(CANDIDATE_DATA_DIR, exist_ok=True)

    temp_file = CANDIDATE_INVOICES_FILE + '.tmp'

    with open(temp_file, 'w', encoding='utf-8') as f:

        json.dump(invoices, f, indent=2, ensure_ascii=False)

    os.replace(temp_file, CANDIDATE_INVOICES_FILE)



def next_candidate_invoice_number():

    year = datetime.now().year

    prefix = f"INV-{year}-"

    max_no = 0

    for invoice in load_candidate_invoices():

        number = str(invoice.get('invoiceNo', ''))

        if number.startswith(prefix):

            try:

                max_no = max(max_no, int(number.rsplit('-', 1)[1]))

            except Exception:

                pass

    return f"{prefix}{max_no + 1:03d}"







# Display information for the two authorised UTA account users.

# The image files should remain in /static as thanzeel.png and kaiff.png.

USER_PROFILES = {

    'thanzeel': {

        'name': 'Thanzeel',

        'role': 'Administrator',

        'photo': 'thanzeel.png',

    },

    'kaiff': {

        'name': 'Kaiff',

        'role': 'Administrator',

        'photo': 'kaiff.png',

    },

}



def get_accounts_user():
    """Return the signed-in Accounts user without affecting Candidate login."""
    username = session.get('accounts_user')
    if username:
        return str(username).lower()

    # Backward compatibility for sessions created before separate logins.
    if session.get('system') == 'accounts' and session.get('user'):
        return str(session.get('user')).lower()

    return None


def get_candidate_user():
    """Return the signed-in Candidate user without affecting Accounts login."""
    username = session.get('candidate_user')
    if username:
        return str(username).lower()

    # Backward compatibility for sessions created before separate logins.
    if session.get('system') == 'candidate' and session.get('user'):
        return str(session.get('user')).lower()

    return None


def get_request_system():
    """Identify which system is serving the current request."""
    path = request.path or ''
    if path.startswith('/candidate/') or path.startswith('/api/candidate/'):
        return 'candidate'
    return 'accounts'


@app.context_processor
def inject_current_user_profile():
    """Make the correct signed-in user's profile available to every template."""
    current_system = get_request_system()

    if current_system == 'candidate':
        username = get_candidate_user() or ''
    else:
        username = get_accounts_user() or ''

    profile = USER_PROFILES.get(username, {
        'name': username.title() if username else 'User',
        'role': 'Candidate Staff' if current_system == 'candidate' else 'Administrator',
        'photo': 'logo.png',
    })

    if current_system == 'candidate':
        if username == 'firnas':
            profile = {
                'name': 'Firnas',
                'role': 'Candidate Staff',
                'photo': 'logo.png',
            }
        elif username in USER_PROFILES:
            profile = dict(USER_PROFILES[username])
            profile['role'] = 'Candidate Staff'

    return {
        'current_user_profile': profile,
        'current_system': current_system,
    }


def login_required(f):
    """Protect Accounts System routes."""
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not get_accounts_user():
            if request.path.startswith('/get_') or request.method != 'GET':
                return jsonify({
                    'success': False,
                    'message': 'Unauthorized. Please login to Accounts again.'
                }), 401
            return redirect(url_for('login_page'))
        return f(*args, **kwargs)
    return wrapper


def candidate_login_required(f):
    """Protect Candidate Management System routes."""
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not get_candidate_user():
            if request.path.startswith('/api/'):
                return jsonify({
                    'success': False,
                    'message': 'Unauthorized. Please login to Candidates again.'
                }), 401
            return redirect(url_for('login_page'))
        return f(*args, **kwargs)
    return wrapper


def create_backup(reason, month_name=None):

    os.makedirs(BACKUP_DIR, exist_ok=True)

    stamp = datetime.now().strftime('%Y%m%d_%H%M%S')

    backup_folder = os.path.join(BACKUP_DIR, f'{stamp}_{reason}')

    os.makedirs(backup_folder, exist_ok=True)

    if os.path.exists(MONTHS_FILE):

        shutil.copy2(MONTHS_FILE, os.path.join(backup_folder, 'months_data.json'))

    if month_name:

        excel_file = os.path.join(EXCEL_DIR, f'{month_name}.xlsx')

        if os.path.exists(excel_file):

            shutil.copy2(excel_file, os.path.join(backup_folder, f'{month_name}.xlsx'))

    return backup_folder



# Create Excel directory if it doesn't exist

if not os.path.exists(EXCEL_DIR):

    os.makedirs(EXCEL_DIR)

if not os.path.exists(BACKUP_DIR):

    os.makedirs(BACKUP_DIR)

if not os.path.exists(DELETED_MONTHS_DIR):

    os.makedirs(DELETED_MONTHS_DIR)

if not os.path.exists(CHEQUE_UPLOAD_DIR):

    os.makedirs(CHEQUE_UPLOAD_DIR)









def load_backup_months():

    if os.path.exists(BACKUP_MONTHS_FILE):

        with open(BACKUP_MONTHS_FILE, 'r') as f:

            return json.load(f)

    return {}



def save_backup_months(data):

    with open(BACKUP_MONTHS_FILE, 'w') as f:

        json.dump(data, f, indent=2)



def move_month_to_backup(month_name):

    """Move a deleted month into the Backup page instead of permanently deleting it."""

    months_data = load_months_data()

    if month_name not in months_data:

        return False

    os.makedirs(DELETED_MONTHS_DIR, exist_ok=True)

    month_data = months_data[month_name].copy()

    source_file = os.path.join(EXCEL_DIR, f"{month_name}.xlsx")

    backup_file = os.path.join(DELETED_MONTHS_DIR, f"{month_name}.xlsx")

    if os.path.exists(source_file):

        shutil.move(source_file, backup_file)

    month_data['deleted_date'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    month_data['backup_filename'] = backup_file

    month_data['status'] = 'backup'

    backup_months = load_backup_months()

    backup_months[month_name] = month_data

    save_backup_months(backup_months)

    del months_data[month_name]

    save_months_data(months_data)

    return True



def load_months_data():

    """Load months data from JSON file"""

    if os.path.exists(MONTHS_FILE):

        with open(MONTHS_FILE, 'r') as f:

            return json.load(f)

    return {}



def save_months_data(data):

    """Save months data to JSON file"""

    with open(MONTHS_FILE, 'w') as f:

        json.dump(data, f, indent=2)



def parse_amount(value, default=0):

    """Parse amounts entered with commas, e.g. 10,000."""

    try:

        return float(str(value).replace(",", "").strip() or default)

    except Exception:

        return float(default)



def format_amount(value):

    """Format amounts with comma grouping across the app."""

    try:

        number = float(value or 0)

    except (TypeError, ValueError):

        number = 0

    if number.is_integer():

        return f"{int(number):,}"

    return f"{number:,.2f}"



app.jinja_env.filters["amount"] = format_amount



def apply_excel_amount_format(filename):

    """Format normal record amounts and cheque-deposit amounts before download."""

    if not os.path.exists(filename):

        return

    wb = load_workbook(filename)

    ws = wb.active

    marker_row = find_cheque_section_row(ws)

    normal_end = (marker_row - 1) if marker_row else ws.max_row

    for row_no in range(2, normal_end + 1):

        for col in [6, 7, 9]:

            cell = ws.cell(row_no, col)

            cell.number_format = "#,##0"

            if isinstance(cell.value, float) and cell.value.is_integer():

                cell.value = int(cell.value)

    if marker_row:

        for row_no in range(marker_row + 2, ws.max_row + 1):

            ws.cell(row_no, 5).number_format = '#,##0.00'

    wb.save(filename)

    wb.close()





def get_current_month():

    """Get current month-year string"""

    return datetime.now().strftime("%B_%Y")



def get_current_month_display():

    """Get current month display name"""

    return datetime.now().strftime("%B %Y")



def create_monthly_sheet(month_name, opening_balance, opening_date=None):

    """Create a new Excel sheet for a month"""

    filename = os.path.join(EXCEL_DIR, f"{month_name}.xlsx")

    wb = Workbook()

    ws = wb.active

    ws.title = "Records"

    # Add headers

    headers = ['No', 'Ref No', 'Date', 'Subject', 'Pass No', 'In Payment', 'Out Payment', 'Sub Agent', 'Balance']

    ws.append(headers)

    # Style headers

    header_fill = PatternFill(start_color="1a1a2e", end_color="1a1a2e", fill_type="solid")

    header_font = Font(bold=True, color="FFFFFF")

    for col in range(1, len(headers) + 1):

        cell = ws.cell(1, col)

        cell.fill = header_fill

        cell.font = header_font

        cell.alignment = Alignment(horizontal="center")

    # Add opening balance row

    if not opening_date:

        opening_date = datetime.now().strftime("%Y-%m-%d")

    ws.append([1, 'OPENING', opening_date, 'OPENING BALANCE', '', 0, 0, '', opening_balance])

    # Style opening balance row

    opening_fill = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")

    opening_font = Font(bold=True, color="2E7D32")

    for col in range(1, 10):

        cell = ws.cell(2, col)

        cell.fill = opening_fill

        cell.font = opening_font

    for col in [6, 7, 9]:

        ws.cell(2, col).number_format = '#,##0'



    # Adjust column widths

    ws.column_dimensions['A'].width = 8

    ws.column_dimensions['B'].width = 12

    ws.column_dimensions['C'].width = 12

    ws.column_dimensions['D'].width = 30

    ws.column_dimensions['E'].width = 12

    ws.column_dimensions['F'].width = 15

    ws.column_dimensions['G'].width = 15

    ws.column_dimensions['H'].width = 15

    ws.column_dimensions['I'].width = 15

    wb.save(filename)

    return filename



def close_monthly_sheet(month_name):

    """Close a monthly sheet (mark as completed)"""

    months_data = load_months_data()

    if month_name in months_data:

        months_data[month_name]['status'] = 'closed'

        months_data[month_name]['closed_date'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        save_months_data(months_data)



def delete_month_sheet(month_name):

    """Delete a monthly sheet completely"""

    try:

        months_data = load_months_data()

        # Delete Excel file

        filename = os.path.join(EXCEL_DIR, f"{month_name}.xlsx")

        if os.path.exists(filename):

            os.remove(filename)

        # Remove from months data

        if month_name in months_data:

            del months_data[month_name]

            save_months_data(months_data)

        return True

    except Exception as e:

        print(f"Error deleting month: {e}")

        return False



def is_month_active(month_name):

    """Check if a month sheet is active"""

    months_data = load_months_data()

    if month_name in months_data:

        return months_data[month_name].get('status') == 'active'

    return False



def get_active_month():

    """Get the currently active month"""

    months_data = load_months_data()

    for month, data in months_data.items():

        if data.get('status') == 'active':

            return month

    return None



def get_next_reference_number(month_name):

    """Get the next reference number for a specific month"""

    filename = os.path.join(EXCEL_DIR, f"{month_name}.xlsx")

    if not os.path.exists(filename):

        return 1

    wb = load_workbook(filename)

    ws = wb.active

    max_num = 0

    for row in ws.iter_rows(min_row=3, values_only=True):

        if row[1] and isinstance(row[1], str) and row[1].startswith('UTA-'):

            try:

                num = int(row[1].split('-')[1])

                max_num = max(max_num, num)

            except:

                continue

    wb.close()

    return max_num + 1



def get_next_row_number(month_name):

    """Get the next row number"""

    filename = os.path.join(EXCEL_DIR, f"{month_name}.xlsx")

    if not os.path.exists(filename):

        return 2

    wb = load_workbook(filename)

    ws = wb.active

    max_no = 0

    for row in ws.iter_rows(min_row=2, values_only=True):

        if row[0] and isinstance(row[0], (int, float)) and row[0] != 'OPENING':

            max_no = max(max_no, int(row[0]))

    wb.close()

    return max_no + 1



def get_current_balance(month_name):

    """Get the current balance from normal transaction rows only."""

    filename = os.path.join(EXCEL_DIR, f"{month_name}.xlsx")

    if not os.path.exists(filename):

        return 0

    wb = load_workbook(filename)

    ws = wb.active

    last_balance = 0

    for row in ws.iter_rows(min_row=2, values_only=True):

        if str(row[0] or '').strip().upper() == CHEQUE_SECTION_MARKER:

            break

        if row[1] == 'OPENING':

            last_balance = float(row[8]) if row[8] else 0

        elif row[8] is not None:

            try:

                last_balance = float(row[8]) if row[8] else 0

            except Exception:

                pass

    wb.close()

    return last_balance



@app.route('/')
def index():
    accounts_user = get_accounts_user()
    candidate_user = get_candidate_user()
    last_system = session.get('last_system')

    if last_system == 'candidate' and candidate_user:
        return redirect(url_for('candidate_dashboard'))

    if last_system == 'accounts' and accounts_user:
        return redirect(url_for('dashboard'))

    if accounts_user:
        return redirect(url_for('dashboard'))

    if candidate_user:
        return redirect(url_for('candidate_dashboard'))

    return redirect(url_for('login_page'))


@app.route('/login', methods=['GET'])
def login_page():
    # Always show the selector page so a user can sign in to the other system
    # without logging out of the system that is already open in another tab.
    return render_template('login.html')


@app.route('/login', methods=['POST'])
def login():
    data = request.get_json(silent=True) or {}
    username = str(data.get('username', '')).strip().lower()
    password = str(data.get('password', ''))
    selected_system = str(data.get('system', 'accounts')).strip().lower()

    if selected_system == 'candidate':
        if not CANDIDATE_PASSWORD:
            return jsonify({
                'success': False,
                'message': 'Candidate password is not configured on the server.'
            }), 503

        if username in CANDIDATE_USERS and hmac.compare_digest(password, CANDIDATE_PASSWORD):
            session.permanent = False

            # Remove only legacy shared-login keys. Do not clear Accounts login.
            session.pop('user', None)
            session.pop('system', None)
            session.pop('login_time', None)

            session['candidate_user'] = username
            session['candidate_login_time'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            session['last_system'] = 'candidate'

            return jsonify({
                'success': True,
                'message': 'Login successful',
                'redirect': url_for('candidate_dashboard')
            })

        return jsonify({
            'success': False,
            'message': 'Invalid username or password'
        }), 401

    stored_hash = USERS.get(username)
    if stored_hash and hmac.compare_digest(stored_hash, hash_password(password)):
        session.permanent = False

        # Remove only legacy shared-login keys. Do not clear Candidate login.
        session.pop('user', None)
        session.pop('system', None)
        session.pop('login_time', None)

        session['accounts_user'] = username
        session['accounts_login_time'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        session['last_system'] = 'accounts'

        return jsonify({
            'success': True,
            'message': 'Login successful',
            'redirect': url_for('dashboard')
        })

    return jsonify({
        'success': False,
        'message': 'Invalid username or password'
    }), 401


@app.route('/logout')
def logout():
    """
    Log out only the system that sent the request.

    Existing templates can continue using /logout. Candidate pages are
    detected from the Referer. You can also use /logout?system=accounts
    or /logout?system=candidate explicitly.
    """
    requested_system = str(request.args.get('system', '')).strip().lower()

    if requested_system not in {'accounts', 'candidate'}:
        referrer = request.referrer or ''
        if '/candidate/' in referrer:
            requested_system = 'candidate'
        elif referrer:
            requested_system = 'accounts'

    if requested_system == 'candidate':
        session.pop('candidate_user', None)
        session.pop('candidate_login_time', None)

        if session.get('last_system') == 'candidate':
            session['last_system'] = 'accounts' if get_accounts_user() else None

    elif requested_system == 'accounts':
        session.pop('accounts_user', None)
        session.pop('accounts_login_time', None)
        session.pop('active_month', None)

        if session.get('last_system') == 'accounts':
            session['last_system'] = 'candidate' if get_candidate_user() else None

    else:
        # Direct /logout with no useful Referer clears both systems.
        session.pop('accounts_user', None)
        session.pop('accounts_login_time', None)
        session.pop('candidate_user', None)
        session.pop('candidate_login_time', None)
        session.pop('active_month', None)
        session.pop('last_system', None)

    # Remove legacy shared-login keys if an old cookie still has them.
    session.pop('user', None)
    session.pop('system', None)
    session.pop('login_time', None)

    if session.get('last_system') is None:
        session.pop('last_system', None)

    return redirect(url_for('login_page'))


@app.route('/candidate/dashboard')
@candidate_login_required
def candidate_dashboard():
    return render_template(
        'candidate/dashboard.html',
        current_user=get_candidate_user(),
        current_date=datetime.now().strftime('%A, %d %B %Y')
    )


@app.route('/candidate/invoice')
@candidate_login_required
def candidate_invoice():
    return render_template(
        'candidate/invoice.html',
        current_user=get_candidate_user(),
        current_date=datetime.now().strftime('%A, %d %B %Y')
    )


@app.route('/api/candidate/invoices', methods=['GET'])

@candidate_login_required

def candidate_invoices_api():

    invoices = load_candidate_invoices()

    invoices.sort(key=lambda x: str(x.get('createdAt', '')), reverse=True)

    return jsonify({'success': True, 'invoices': invoices})



@app.route('/api/candidate/invoices/next-number', methods=['GET'])

@candidate_login_required

def candidate_invoice_next_number_api():

    return jsonify({'success': True, 'invoiceNo': next_candidate_invoice_number()})



@app.route('/api/candidate/invoices', methods=['POST'])

@candidate_login_required

def candidate_invoice_create_api():

    data = request.get_json(silent=True) or {}

    client_name = str(data.get('clientName', '')).strip()

    if not client_name:

        return jsonify({'success': False, 'message': 'Client name is required.'}), 400



    items = data.get('items', [])

    if not isinstance(items, list):

        return jsonify({'success': False, 'message': 'Invalid invoice items.'}), 400



    clean_items = []

    subtotal = 0.0

    for item in items:

        if not isinstance(item, dict):

            continue

        description = str(item.get('description', '')).strip()

        try:

            quantity = float(item.get('quantity', 0) or 0)

            price = float(item.get('price', 0) or 0)

        except (TypeError, ValueError):

            quantity = 0.0

            price = 0.0

        total = quantity * price

        clean_items.append({

            'description': description,

            'quantity': quantity,

            'price': price,

            'total': total

        })

        subtotal += total



    invoices = load_candidate_invoices()

    invoice_id = int(datetime.now().timestamp() * 1000)

    invoice_no = str(data.get('invoiceNo', '')).strip() or next_candidate_invoice_number()

    if any(str(inv.get('invoiceNo', '')).strip().lower() == invoice_no.lower() for inv in invoices):

        invoice_no = next_candidate_invoice_number()



    invoice = {

        'id': invoice_id,

        'invoiceNo': invoice_no,

        'date': str(data.get('date', '')).strip() or datetime.now().strftime('%Y-%m-%d'),

        'clientName': client_name,

        'passportNo': str(data.get('passportNo', '')).strip(),

        'destination': str(data.get('destination', '')).strip(),

        'items': clean_items,

        'subtotal': subtotal,

        'vat': 0,

        'grandTotal': subtotal,

        'createdAt': datetime.now().isoformat(timespec='seconds'),

        'createdBy': get_candidate_user()

    }

    invoices.append(invoice)

    save_candidate_invoices(invoices)

    return jsonify({'success': True, 'message': 'Invoice saved successfully.', 'invoice': invoice})



@app.route('/api/candidate/invoices/<int:invoice_id>', methods=['DELETE'])

@candidate_login_required

def candidate_invoice_delete_api(invoice_id):

    invoices = load_candidate_invoices()

    new_invoices = [inv for inv in invoices if int(inv.get('id', 0) or 0) != invoice_id]

    if len(new_invoices) == len(invoices):

        return jsonify({'success': False, 'message': 'Invoice not found.'}), 404

    save_candidate_invoices(new_invoices)

    return jsonify({'success': True, 'message': 'Invoice deleted successfully.'})



@app.route('/dashboard')

@login_required

def dashboard():

    """Dashboard showing all months"""

    months_data = load_months_data()

    active_month = get_active_month()

    current_month_display = get_current_month_display()

    return render_template('dashboard.html', 

                         months_data=months_data, 

                         active_month=active_month,

                         current_month_display=current_month_display,

                         current_user=get_accounts_user())



@app.route('/create_month', methods=['POST'])

@login_required

def create_month():

    """Create a new month sheet"""

    try:

        data = request.json

        opening_balance = parse_amount(data.get('opening_balance', 0))

        creation_date_str = data.get('creation_date', '').strip()



        # Validate and parse the creation date

        if creation_date_str:

            try:

                creation_date = datetime.strptime(creation_date_str, "%Y-%m-%d")

            except ValueError:

                return jsonify({'success': False, 'message': 'Invalid date format. Please use YYYY-MM-DD.'})

        else:

            creation_date = datetime.now()



        # Derive month key and display name from the chosen date

        current_month = creation_date.strftime("%B_%Y")

        month_display = creation_date.strftime("%B %Y")



        # Check if month already exists

        months_data = load_months_data()

        if current_month in months_data:

            return jsonify({'success': False, 'message': f'Month {month_display} already exists!'})



        # Create Excel file using the chosen date as the opening balance date

        opening_date_str = creation_date.strftime("%Y-%m-%d")

        filename = create_monthly_sheet(current_month, opening_balance, opening_date=opening_date_str)



        # Save month data — store the chosen date (not server now()) as created_date

        months_data[current_month] = {

            'display_name': month_display,

            'created_date': creation_date.strftime("%Y-%m-%d") + " 00:00:00",

            'opening_balance': opening_balance,

            'status': 'active',

            'filename': filename

        }

        save_months_data(months_data)



        session['active_month'] = current_month



        return jsonify({

            'success': True,

            'message': f'Month {month_display} created successfully!',

            'month_name': current_month

        })



    except Exception as e:

        return jsonify({'success': False, 'message': f'Error: {str(e)}'})



@app.route('/delete_month', methods=['POST'])

@login_required

def delete_month():

    """Delete a month sheet completely"""

    try:

        data = request.json

        month_name = data.get('month_name')

        if not month_name:

            return jsonify({'success': False, 'message': 'Month name not provided'})

        # Check if it's the active month

        active_month = get_active_month()

        if active_month == month_name:

            return jsonify({'success': False, 'message': 'Cannot delete active month! Please end the month first.'})

        # Move the month to the Backup page instead of deleting permanently

        create_backup('delete_month', month_name)

        if move_month_to_backup(month_name):

            return jsonify({'success': True, 'message': 'Month moved to Backup successfully!'})

        else:

            return jsonify({'success': False, 'message': 'Error moving month to Backup'})

    except Exception as e:

        return jsonify({'success': False, 'message': f'Error: {str(e)}'})



@app.route('/end_month', methods=['POST'])

@login_required

def end_month():

    """Close the current month"""

    try:

        active_month = get_active_month()

        if not active_month:

            return jsonify({'success': False, 'message': 'No active month found!'})

        deposits = load_cheque_deposits()

        changed = False

        for dep in deposits.values():

            if dep.get('month') == active_month and not dep.get('excel_added'):

                dep['excel_added'] = True

                changed = True

        if changed:

            save_cheque_deposits(deposits)

        sync_cheque_deposits_to_excel(active_month, include_all=True)

        close_monthly_sheet(active_month)

        session.pop('active_month', None)

        return jsonify({'success': True, 'message': f'Month closed successfully!'})

    except Exception as e:

        return jsonify({'success': False, 'message': f'Error: {str(e)}'})



@app.route('/enter')

@login_required

def enter_records():

    """Enter records page"""

    active_month = get_active_month()

    if not active_month:

        return render_template('no_active_month.html')

    months_data = load_months_data()

    return render_template('enter_records.html', 

                         active_month=active_month, 

                         month_display=months_data[active_month]['display_name'])



@app.route('/view')

@login_required

def view_records():

    """View records page"""

    months_data = load_months_data()

    return render_template('view_records.html', months_data=months_data)



@app.route('/view_month/<month_name>')

@login_required

def view_month_records(month_name):

    """View specific month records"""

    months_data = load_months_data()

    if month_name not in months_data:

        return "Month not found", 404

    return render_template('view_month.html', 

                         month_name=month_name, 

                         month_data=months_data[month_name],

                         cheque_deposits=get_month_cheque_deposits(month_name))



@app.route('/delete_record_from_month', methods=['POST'])

@login_required

def delete_record_from_month():

    """Delete a single record from a specific month"""

    try:

        data = request.json

        month_name = data.get('month_name')

        ref_no = data.get('ref_no') or str(data.get('record_no', ''))



        if not month_name or not ref_no:

            return jsonify({'success': False, 'message': 'Missing required data'})

        filename = os.path.join(EXCEL_DIR, f"{month_name}.xlsx")

        if not os.path.exists(filename):

            return jsonify({'success': False, 'message': 'Month file not found'})

        wb = load_workbook(filename)

        ws = wb.active

        remove_cheque_section(ws)

        row_to_delete = None

        for row in range(3, ws.max_row + 1):

            if str(ws.cell(row, 2).value or '').strip() == str(ref_no).strip():

                row_to_delete = row

                break

        if row_to_delete:

            create_backup('delete_record', month_name)

            ws.delete_rows(row_to_delete)



            # Re-sort by date, reassign Ref Nos sequentially, and recalculate balances

            recalculate_and_sort_sheet(ws)



            wb.save(filename)

            wb.close()

            sync_cheque_deposits_to_excel(month_name)

            return jsonify({'success': True, 'message': f'Record deleted and references re-indexed successfully!'})

        else:

            wb.close()

            return jsonify({'success': False, 'message': 'Record not found'})

    except Exception as e:

        print(f"Error deleting record: {e}")

        return jsonify({'success': False, 'message': f'Error deleting: {str(e)}'})



@app.route('/update_record_in_month', methods=['POST'])

@login_required

def update_record_in_month():

    """Update a single record in a specific month"""

    try:

        data = request.json

        month_name = data.get('month_name')

        ref_no = data.get('ref_no')



        if not month_name or not ref_no:

            return jsonify({'success': False, 'message': 'Missing required data'})



        if not data.get('date') or not data.get('subject'):

            return jsonify({'success': False, 'message': 'Date and Subject are required'})



        filename = os.path.join(EXCEL_DIR, f"{month_name}.xlsx")

        if not os.path.exists(filename):

            return jsonify({'success': False, 'message': 'Month file not found'})



        wb = load_workbook(filename)

        ws = wb.active

        remove_cheque_section(ws)

        row_to_update = None

        for row in range(3, ws.max_row + 1):

            if str(ws.cell(row, 2).value or '').strip() == str(ref_no).strip():

                row_to_update = row

                break



        if not row_to_update:

            wb.close()

            return jsonify({'success': False, 'message': 'Record not found'})



        create_backup('update_record', month_name)



        ws.cell(row_to_update, 3).value = data.get('date', '')

        ws.cell(row_to_update, 4).value = data.get('subject', '')

        ws.cell(row_to_update, 5).value = data.get('pass_no', '')

        ws.cell(row_to_update, 6).value = parse_amount(data.get('in_payment', 0))

        ws.cell(row_to_update, 7).value = parse_amount(data.get('out_payment', 0))

        ws.cell(row_to_update, 8).value = data.get('sub_agent', '')



        # Re-sort by date, reassign Ref Nos sequentially, and recalculate balances

        recalculate_and_sort_sheet(ws)



        wb.save(filename)

        wb.close()

        sync_cheque_deposits_to_excel(month_name)

        return jsonify({'success': True, 'message': 'Record updated successfully!'})



    except Exception as e:

        print(f"Error updating record: {e}")

        return jsonify({'success': False, 'message': f'Error updating: {str(e)}'})



@app.route('/get_month_records/<month_name>')

@login_required

def get_month_records(month_name):

    """Get normal transaction records for a specific month."""

    try:

        filename = os.path.join(EXCEL_DIR, f"{month_name}.xlsx")

        if not os.path.exists(filename):

            return jsonify({'records': []})

        wb = load_workbook(filename)

        ws = wb.active

        records = []

        for row in ws.iter_rows(min_row=3, values_only=True):

            if str(row[0] or '').strip().upper() == CHEQUE_SECTION_MARKER:

                break

            if row[0] and row[1] and str(row[1]).startswith('UTA-'):

                records.append({

                    'no': row[0], 'ref_no': row[1] or '', 'date': str(row[2]) if row[2] else '',

                    'subject': row[3] or '', 'pass_no': row[4] or '',

                    'in_payment': float(row[5]) if row[5] else 0,

                    'out_payment': float(row[6]) if row[6] else 0,

                    'sub_agent': row[7] or '', 'balance': float(row[8]) if row[8] else 0

                })

        wb.close()

        return jsonify({'records': records})

    except Exception as e:

        print(f"Error getting records: {e}")

        return jsonify({'records': [], 'error': str(e)})



def recalculate_and_sort_sheet(ws):

    """Sort normal transactions, renumber references and recalculate balances."""

    opening_row_data = None

    transaction_rows = []

    for row in ws.iter_rows(min_row=2, values_only=True):

        row = list(row)

        if str(row[0] or '').strip().upper() == CHEQUE_SECTION_MARKER:

            break

        if row[1] == 'OPENING':

            opening_row_data = row

        elif row[0] is not None and row[1] and str(row[1]).startswith('UTA-'):

            transaction_rows.append(row)

    if opening_row_data is None:

        return

    def parse_date_safe(row):

        try:

            date_val = row[2]

            if isinstance(date_val, str):

                return datetime.strptime(date_val, '%Y-%m-%d')

            if hasattr(date_val, 'year'):

                return datetime(date_val.year, date_val.month, date_val.day)

        except Exception:

            pass

        return datetime(9999, 12, 31)

    transaction_rows.sort(key=parse_date_safe)

    ws.delete_rows(2, ws.max_row)

    ws.append(opening_row_data)

    opening_row_idx = ws.max_row

    opening_fill = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")

    opening_font = Font(bold=True, color="2E7D32")

    for col in range(1, 10):

        ws.cell(opening_row_idx, col).fill = opening_fill

        ws.cell(opening_row_idx, col).font = opening_font

    for col in [6, 7, 9]:

        ws.cell(opening_row_idx, col).number_format = '#,##0'

    balance = float(opening_row_data[8]) if opening_row_data[8] is not None else 0

    for i, row in enumerate(transaction_rows):

        in_payment = float(row[5]) if row[5] else 0

        out_payment = float(row[6]) if row[6] else 0

        balance = balance + in_payment - out_payment

        ws.append([i + 2, f"UTA-{str(i + 1).zfill(2)}", row[2], row[3], row[4], in_payment, out_payment, row[7], balance])

        for col in [6, 7, 9]:

            ws.cell(ws.max_row, col).number_format = '#,##0'





@app.route('/save_records', methods=['POST'])

@login_required

def save_records():

    """Save records to active month sheet"""

    try:

        active_month = get_active_month()

        if not active_month:

            return jsonify({'success': False, 'message': 'No active month! Please create a new month first.'})

        data = request.json

        records = data.get('records', [])

        if not records:

            return jsonify({'success': False, 'message': 'No records to save'})

        filename = os.path.join(EXCEL_DIR, f"{active_month}.xlsx")

        if not os.path.exists(filename):

            return jsonify({'success': False, 'message': 'Month file not found!'})

        wb = load_workbook(filename)

        ws = wb.active

        remove_cheque_section(ws)

        next_ref_num = get_next_reference_number(active_month)

        current_balance = get_current_balance(active_month)

        saved_count = 0

        ref_offset = 0

        for record in records:

            if not record.get('date') or not record.get('subject'):

                continue

            ref_no = f"UTA-{str(next_ref_num + ref_offset).zfill(2)}"

            in_payment = parse_amount(record.get('in_payment', 0))

            out_payment = parse_amount(record.get('out_payment', 0))

            new_balance = current_balance + in_payment - out_payment



            # Row number is a placeholder; recalculate_and_sort_sheet will fix it

            row = [

                0,

                ref_no,

                record.get('date', ''),

                record.get('subject', ''),

                record.get('pass_no', ''),

                in_payment,

                out_payment,

                record.get('sub_agent', ''),

                new_balance

            ]

            ws.append(row)

            current_balance = new_balance

            saved_count += 1

            ref_offset += 1



        # Sort all rows by date and recalculate balances in one pass

        recalculate_and_sort_sheet(ws)



        wb.save(filename)

        wb.close()

        sync_cheque_deposits_to_excel(active_month)

        return jsonify({

            'success': True,

            'message': f'Successfully saved {saved_count} records!',

            'next_reference': f"UTA-{str(next_ref_num + saved_count).zfill(2)}"

        })

    except Exception as e:

        print(f"Error saving records: {e}")

        return jsonify({'success': False, 'message': f'Error saving: {str(e)}'})

@app.route('/get_next_reference')

@login_required

def get_next_reference():

    """Get the next reference number for display"""

    active_month = get_active_month()

    if not active_month:

        return jsonify({'next_reference': 'No Active Month'})

    next_num = get_next_reference_number(active_month)

    return jsonify({'next_reference': f"UTA-{str(next_num).zfill(2)}"})



@app.route('/get_stats')

@login_required

def get_stats():

    """Get statistics for dashboard — shows active month data only."""

    months_data = load_months_data()

    total_months = len(months_data)

    active_month = get_active_month()



    total_records = 0

    total_in_payment = 0

    total_out_payment = 0

    total_references = 0

    net_balance = 0



    if active_month:

        filename = os.path.join(EXCEL_DIR, f"{active_month}.xlsx")

        if os.path.exists(filename):

            wb = load_workbook(filename)

            ws = wb.active

            opening_balance = 0



            for row in ws.iter_rows(min_row=2, values_only=True):

                if str(row[0] or '').strip().upper() == CHEQUE_SECTION_MARKER:

                    break

                if row[1] == 'OPENING':

                    opening_balance = float(row[8]) if row[8] else 0

                elif row[0] is not None:

                    total_records += 1

                    total_in_payment += float(row[5]) if row[5] else 0

                    total_out_payment += float(row[6]) if row[6] else 0

                    if row[1] and row[1] != 'OPENING':

                        total_references += 1



            wb.close()

            net_balance = opening_balance + total_in_payment - total_out_payment



    return jsonify({

        'total_months': total_months,

        'total_records': total_records,

        'total_in_payment': total_in_payment,

        'total_out_payment': total_out_payment,

        'total_references': total_references,

        'net_balance': net_balance,

        'active_month': active_month,

        'active_month_display': months_data.get(active_month, {}).get('display_name', 'None') if active_month else 'None'

    })



@app.route('/get_month_stats/<month_name>')

@login_required

def get_month_stats(month_name):

    """Get statistics for a specific month"""

    try:

        filename = os.path.join(EXCEL_DIR, f"{month_name}.xlsx")

        if not os.path.exists(filename):

            return jsonify({'error': 'Month not found'}), 404

        wb = load_workbook(filename)

        ws = wb.active

        total_records = 0

        total_in_payment = 0

        total_out_payment = 0

        opening_balance = 0

        closing_balance = 0

        total_references = set()

        for row in ws.iter_rows(min_row=2, values_only=True):

            if str(row[0] or '').strip().upper() == CHEQUE_SECTION_MARKER:

                break

            if row[0] == 'OPENING':

                opening_balance = float(row[8]) if row[8] else 0

            elif row[0] and row[0] != 'OPENING' and row[0] is not None:

                total_records += 1

                total_in_payment += float(row[5]) if row[5] else 0

                total_out_payment += float(row[6]) if row[6] else 0

                if row[1]:

                    total_references.add(row[1])

                closing_balance = float(row[8]) if row[8] else 0

        wb.close()

        return jsonify({

            'total_records': total_records,

            'total_in_payment': total_in_payment,

            'total_out_payment': total_out_payment,

            'total_references': len(total_references),

            'opening_balance': opening_balance,

            'closing_balance': closing_balance,

            'net_change': closing_balance - opening_balance

        })

    except Exception as e:

        print(f"Error getting month stats: {e}")

        return jsonify({'error': str(e)}), 500





@app.route('/cheque-deposits')

@login_required

def cheque_deposits_page():

    active_month = get_active_month()

    months_data = load_months_data()

    deposits = get_month_cheque_deposits(active_month)

    return render_template(

        'cheque_deposits.html',

        active_month=active_month,

        month_display=months_data.get(active_month, {}).get('display_name', 'No Active Month') if active_month else 'No Active Month',

        deposits=deposits,

        next_deposit_id=next_cheque_deposit_id(active_month) if active_month else ''

    )





@app.route('/save_cheque_deposit', methods=['POST'])

@login_required

def save_cheque_deposit():

    try:

        active_month = get_active_month()

        if not active_month:

            return jsonify({'success': False, 'message': 'No active month. Create a month first.'}), 400



        deposit_id = str(request.form.get('deposit_id', '')).strip() or next_cheque_deposit_id(active_month)

        deposits = load_cheque_deposits()

        existing = deposits.get(deposit_id, {})



        image_filename = existing.get('image_filename', '')

        image = request.files.get('slip_image')

        if image and image.filename:

            if not cheque_file_allowed(image.filename):

                return jsonify({'success': False, 'message': 'Only PNG, JPG, JPEG or WEBP images are allowed.'}), 400

            ext = secure_filename(image.filename).rsplit('.', 1)[1].lower()

            image_filename = f"{secure_filename(deposit_id)}.{ext}"

            image.save(os.path.join(CHEQUE_UPLOAD_DIR, image_filename))



        entry = {

            'deposit_id': deposit_id,

            'month': active_month,

            'date': str(request.form.get('date', '')).strip(),

            'time': str(request.form.get('time', '')).strip(),

            'transaction_id': str(request.form.get('transaction_id', '')).strip(),

            'location': str(request.form.get('location', '')).strip(),

            'account_number': str(request.form.get('account_number', '')).strip(),

            'account_name': str(request.form.get('account_name', '')).strip(),

            'nic_number': str(request.form.get('nic_number', '')).strip(),

            'contact_number': str(request.form.get('contact_number', '')).strip(),

            'reference_number': str(request.form.get('reference_number', '')).strip(),

            'cheque_no': str(request.form.get('cheque_no', '')).strip(),

            'bank': str(request.form.get('bank', '')).strip(),

            'branch': str(request.form.get('branch', '')).strip(),

            'amount': parse_amount(request.form.get('amount', 0)),

            'notes': str(request.form.get('notes', '')).strip(),

            'raw_text': str(request.form.get('raw_text', '')).strip(),

            'image_filename': image_filename,

            'excel_added': bool(existing.get('excel_added', False)),

            'created_by': existing.get('created_by') or get_accounts_user(),

            'created_at': existing.get('created_at') or datetime.now().strftime('%Y-%m-%d %H:%M:%S'),

            'updated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        }

        deposits[deposit_id] = entry

        save_cheque_deposits(deposits)



        if entry['excel_added']:

            sync_cheque_deposits_to_excel(active_month)



        return jsonify({

            'success': True,

            'message': 'Cheque deposit saved successfully.',

            'deposit_id': deposit_id,

            'deposit': entry

        })

    except Exception as e:

        return jsonify({'success': False, 'message': f'Error saving cheque deposit: {str(e)}'}), 500





@app.route('/add_cheque_deposit_to_excel', methods=['POST'])

@login_required

def add_cheque_deposit_to_excel():

    try:

        data = request.get_json(silent=True) or {}

        deposit_id = str(data.get('deposit_id', '')).strip()

        deposits = load_cheque_deposits()

        if deposit_id not in deposits:

            return jsonify({'success': False, 'message': 'Save the cheque deposit first.'}), 404



        entry = deposits[deposit_id]

        month_name = entry.get('month')

        entry['excel_added'] = True

        entry['updated_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        deposits[deposit_id] = entry

        save_cheque_deposits(deposits)

        sync_cheque_deposits_to_excel(month_name)

        return jsonify({'success': True, 'message': 'Added to the existing Records Excel sheet.'})

    except Exception as e:

        return jsonify({'success': False, 'message': f'Error adding to Excel: {str(e)}'}), 500





@app.route('/delete_cheque_deposit', methods=['POST'])

@login_required

def delete_cheque_deposit():

    try:

        data = request.get_json(silent=True) or {}

        deposit_id = str(data.get('deposit_id', '')).strip()

        deposits = load_cheque_deposits()

        if deposit_id not in deposits:

            return jsonify({'success': False, 'message': 'Cheque deposit not found.'}), 404



        entry = deposits[deposit_id]

        month_name = entry.get('month')

        image_filename = entry.get('image_filename')

        if image_filename:

            image_path = os.path.join(CHEQUE_UPLOAD_DIR, image_filename)

            if os.path.exists(image_path):

                try:

                    os.remove(image_path)

                except OSError:

                    pass

        del deposits[deposit_id]

        save_cheque_deposits(deposits)

        if month_name:

            sync_cheque_deposits_to_excel(month_name)

        return jsonify({'success': True, 'message': 'Cheque deposit deleted.'})

    except Exception as e:

        return jsonify({'success': False, 'message': f'Error deleting cheque deposit: {str(e)}'}), 500





@app.route('/cheque_deposit_image/<deposit_id>')

@login_required

def cheque_deposit_image(deposit_id):

    deposits = load_cheque_deposits()

    entry = deposits.get(deposit_id)

    if not entry or not entry.get('image_filename'):

        return 'Image not found', 404

    path = os.path.join(CHEQUE_UPLOAD_DIR, entry['image_filename'])

    if not os.path.exists(path):

        return 'Image not found', 404

    return send_file(path)





@app.route('/cheque_deposit_pdf/<deposit_id>')

@login_required

def cheque_deposit_pdf(deposit_id):

    deposits = load_cheque_deposits()

    entry = deposits.get(deposit_id)

    if not entry:

        return 'Cheque deposit not found', 404



    buffer = BytesIO()

    pdf = canvas.Canvas(buffer, pagesize=A4)

    width, height = A4

    pdf.setTitle(f"UTA Cheque Deposit - {deposit_id}")



    pdf.setFont('Helvetica-Bold', 18)

    pdf.drawString(48, height - 55, 'UTA Manpower Service')

    pdf.setFont('Helvetica-Bold', 14)

    pdf.drawString(48, height - 82, 'Cheque Deposit Slip Record')

    pdf.setStrokeColorRGB(0.85, 0.08, 0.20)

    pdf.line(48, height - 94, width - 48, height - 94)



    fields = [

        ('Deposit ID', entry.get('deposit_id', '')),

        ('Date', entry.get('date', '')),

        ('Time', entry.get('time', '')),

        ('Transaction ID', entry.get('transaction_id', '')),

        ('Location', entry.get('location', '')),

        ('Account Number', entry.get('account_number', '')),

        ('Account Name', entry.get('account_name', '')),

        ('NIC Number', entry.get('nic_number', '')),

        ('Contact Number', entry.get('contact_number', '')),

        ('Reference Number', entry.get('reference_number', '')),

        ('Cheque No', entry.get('cheque_no', '')),

        ('Bank', entry.get('bank', '')),

        ('Branch', entry.get('branch', '')),

        ('Amount', format_amount(entry.get('amount', 0))),

        ('Notes', entry.get('notes', '')),

    ]



    y = height - 125

    for label, value in fields:

        pdf.setFont('Helvetica-Bold', 10)

        pdf.drawString(48, y, f'{label}:')

        pdf.setFont('Helvetica', 10)

        display = str(value or '-')

        if len(display) > 70:

            display = display[:67] + '...'

        pdf.drawString(155, y, display)

        y -= 18

        if y < 180:

            pdf.showPage()

            y = height - 60



    image_filename = entry.get('image_filename')

    if image_filename:

        image_path = os.path.join(CHEQUE_UPLOAD_DIR, image_filename)

        if os.path.exists(image_path):

            try:

                img = ImageReader(image_path)

                iw, ih = img.getSize()

                max_w, max_h = width - 96, min(280, y - 55)

                if max_h > 80:

                    scale = min(max_w / iw, max_h / ih)

                    draw_w, draw_h = iw * scale, ih * scale

                    pdf.drawImage(img, 48, max(45, y - draw_h - 10), width=draw_w, height=draw_h, preserveAspectRatio=True)

            except Exception:

                pass



    pdf.save()

    buffer.seek(0)

    return send_file(buffer, mimetype='application/pdf', as_attachment=True, download_name=f'{deposit_id}.pdf')





@app.route('/backup')

@login_required

def backup_page():

    backup_months = load_backup_months()

    return render_template('backup.html', backup_months=backup_months)





@app.route('/restore_backup_month', methods=['POST'])

@login_required

def restore_backup_month():

    try:

        data = request.json

        month_name = data.get('month_name')

        if not month_name:

            return jsonify({'success': False, 'message': 'Month name not provided'})



        backup_months = load_backup_months()

        if month_name not in backup_months:

            return jsonify({'success': False, 'message': 'Backup month not found'})



        months_data = load_months_data()

        if month_name in months_data:

            return jsonify({'success': False, 'message': 'This month already exists in records'})



        os.makedirs(EXCEL_DIR, exist_ok=True)

        backup_data = backup_months[month_name].copy()

        backup_file = backup_data.get('backup_filename') or os.path.join(DELETED_MONTHS_DIR, f"{month_name}.xlsx")

        restore_file = os.path.join(EXCEL_DIR, f"{month_name}.xlsx")



        if os.path.exists(backup_file):

            shutil.move(backup_file, restore_file)



        backup_data.pop('deleted_date', None)

        backup_data.pop('backup_filename', None)

        backup_data['status'] = 'closed'

        backup_data['filename'] = restore_file



        months_data[month_name] = backup_data

        save_months_data(months_data)



        del backup_months[month_name]

        save_backup_months(backup_months)



        return jsonify({'success': True, 'message': 'Month restored from Backup successfully!'})

    except Exception as e:

        return jsonify({'success': False, 'message': f'Error: {str(e)}'})



@app.route('/delete_backup_month', methods=['POST'])

@login_required

def delete_backup_month():

    try:

        data = request.json

        month_name = data.get('month_name')

        if not month_name:

            return jsonify({'success': False, 'message': 'Month name not provided'})

        backup_months = load_backup_months()

        if month_name not in backup_months:

            return jsonify({'success': False, 'message': 'Backup month not found'})

        backup_file = backup_months[month_name].get('backup_filename') or os.path.join(DELETED_MONTHS_DIR, f"{month_name}.xlsx")

        if os.path.exists(backup_file):

            os.remove(backup_file)

        del backup_months[month_name]

        save_backup_months(backup_months)

        return jsonify({'success': True, 'message': 'Backup month deleted permanently!'})

    except Exception as e:

        return jsonify({'success': False, 'message': f'Error: {str(e)}'})



@app.route('/download_excel/<month_name>')

@login_required

def download_excel(month_name):

    """Download a specific month's Excel file"""

    try:

        filename = os.path.join(EXCEL_DIR, f"{month_name}.xlsx")

        if os.path.exists(filename):

            month_info = load_months_data().get(month_name, {})

            sync_cheque_deposits_to_excel(month_name, include_all=(month_info.get('status') == 'closed'))

            apply_excel_amount_format(filename)

            return send_file(

                filename, 

                as_attachment=True, 

                download_name=f'UTA_{month_name}.xlsx'

            )

        else:

            return jsonify({'error': 'File not found'}), 404

    except Exception as e:

        return jsonify({'error': str(e)}), 500



if __name__ == '__main__':

    print("\n" + "="*60)

    print("📊 UTA Manpower Service - Application Started")

    print("="*60)

    print(f"📁 Excel Sheets Directory: {EXCEL_DIR}")

    print(f"🌐 Server: http://localhost:5000")

    print("\n✅ Month-based sheet management enabled")

    print("✅ Automatic balance calculation")

    print("✅ Delete month and record functionality enabled")

    print("="*60 + "\n")

    # Try port 5000, if busy use 5001

    port = 5000

    try:

        app.run(debug=True, port=port, host='127.0.0.1')

    except OSError as e:

        if "Address already in use" in str(e):

            print(f"⚠️ Port {port} is busy. Trying port 5001...")

            port = 5001

            app.run(debug=True, port=port, host='127.0.0.1')

        else:

            raise