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

import tempfile

import threading

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

# Serialize Accounts Excel reads/writes so simultaneous requests cannot overwrite each other.
RECORD_IO_LOCK = threading.RLock()





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
CANDIDATE_SALARY_SLIPS_FILE = os.path.join(CANDIDATE_DATA_DIR, 'salary_slips.json')


def _load_candidate_json_list(filename):
    os.makedirs(CANDIDATE_DATA_DIR, exist_ok=True)
    if not os.path.exists(filename):
        return []
    try:
        with open(filename, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data if isinstance(data, list) else []
    except Exception:
        return []


def _save_candidate_json_list(filename, rows):
    os.makedirs(CANDIDATE_DATA_DIR, exist_ok=True)
    temp_file = filename + '.tmp'
    with open(temp_file, 'w', encoding='utf-8') as f:
        json.dump(rows, f, indent=2, ensure_ascii=False)
    os.replace(temp_file, filename)


def load_candidate_invoices():
    return _load_candidate_json_list(CANDIDATE_INVOICES_FILE)


def save_candidate_invoices(invoices):
    _save_candidate_json_list(CANDIDATE_INVOICES_FILE, invoices)


def load_candidate_salary_slips():
    return _load_candidate_json_list(CANDIDATE_SALARY_SLIPS_FILE)


def save_candidate_salary_slips(slips):
    _save_candidate_json_list(CANDIDATE_SALARY_SLIPS_FILE, slips)


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


def next_candidate_salary_slip_number():
    year = datetime.now().year
    prefix = f"SAL-{year}-"
    max_no = 0
    for slip in load_candidate_salary_slips():
        number = str(slip.get('slipNo', ''))
        if number.startswith(prefix):
            try:
                max_no = max(max_no, int(number.rsplit('-', 1)[1]))
            except Exception:
                pass
    return f"{prefix}{max_no + 1:03d}"
# Display information for authorised UTA users.
# Profile image files must remain in /static as thanzeel.png, kaiff.png and firnas.png.

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
    'firnas': {
        'name': 'Firnas',
        'role': 'Candidate Staff',
        'photo': 'firnas.png',
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

    if current_system == 'candidate' and username in USER_PROFILES:
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



def _normalise_record_date(value):
    """Return record dates as YYYY-MM-DD where possible."""
    if value is None:
        return ''
    if hasattr(value, 'strftime'):
        try:
            return value.strftime('%Y-%m-%d')
        except Exception:
            pass
    raw = str(value).strip()
    if not raw:
        return ''
    for fmt in ('%Y-%m-%d', '%Y-%m-%d %H:%M:%S', '%d/%m/%Y', '%d-%m-%Y'):
        try:
            return datetime.strptime(raw, fmt).strftime('%Y-%m-%d')
        except Exception:
            continue
    try:
        return datetime.fromisoformat(raw).strftime('%Y-%m-%d')
    except Exception:
        return raw


def _record_date_sort_key(value):
    normalised = _normalise_record_date(value)
    try:
        return datetime.strptime(normalised, '%Y-%m-%d')
    except Exception:
        return datetime(9999, 12, 31)


def _normalise_import_header(value):
    """Normalise an uploaded Excel column heading for reliable matching."""
    text = str(value or '').strip().lower()
    for ch in ('_', '-', '.', ':', '/', '\\', '(', ')'):
        text = text.replace(ch, ' ')
    return ' '.join(text.split())


EXCEL_IMPORT_HEADER_ALIASES = {
    'date': {
        'date', 'transaction date', 'entry date', 'payment date', 'record date',
    },
    'subject': {
        'subject', 'description', 'details', 'detail', 'particulars', 'particular',
        'narration', 'purpose', 'remarks', 'remark',
    },
    'in_payment': {
        'in payment', 'in', 'credit', 'income', 'received', 'receipt', 'cash in',
        'deposit', 'amount in',
    },
    'out_payment': {
        'out payment', 'out', 'debit', 'expense', 'paid', 'payment', 'cash out',
        'withdrawal', 'amount out',
    },
}


def _match_import_header(value):
    normalised = _normalise_import_header(value)
    if not normalised:
        return None
    for canonical, aliases in EXCEL_IMPORT_HEADER_ALIASES.items():
        normalised_aliases = {_normalise_import_header(alias) for alias in aliases}
        if normalised in normalised_aliases:
            return canonical
    return None


def _parse_excel_import_amount(value):
    """Parse common Excel/currency amount formats without failing the whole import."""
    if value is None or value == '':
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    raw = str(value).strip()
    if not raw:
        return 0.0
    negative = raw.startswith('(') and raw.endswith(')')
    raw = raw.replace(',', '')
    raw = raw.replace('Rs.', '').replace('Rs', '').replace('LKR', '').replace('lkr', '')
    raw = raw.replace('(', '').replace(')', '').strip()
    filtered = ''.join(ch for ch in raw if ch.isdigit() or ch in '.-')
    if filtered in ('', '-', '.', '-.'):
        return 0.0
    try:
        number = float(filtered)
        return -abs(number) if negative else number
    except Exception:
        return 0.0


def _find_excel_import_header(ws):
    """Find the best header row in the first 20 rows and return its field mapping."""
    best = None
    max_scan = min(ws.max_row or 0, 20)
    for row_no in range(1, max_scan + 1):
        values = [cell.value for cell in ws[row_no]]
        mapping = {}
        for col_idx, value in enumerate(values, start=1):
            canonical = _match_import_header(value)
            if canonical and canonical not in mapping:
                mapping[canonical] = col_idx
        score = len(mapping)
        # Date + subject are essential for Enter Records. Other fields are optional.
        if 'date' in mapping and 'subject' in mapping:
            candidate = (score, -row_no, row_no, mapping)
            if best is None or candidate > best:
                best = candidate
    if best is None:
        return None, {}
    return best[2], best[3]


def _excel_cell_value(ws, row_no, mapping, field):
    col = mapping.get(field)
    return ws.cell(row_no, col).value if col else None


def _record_sheet_is_legacy(ws):
    """Return True for the old 9-column Accounts layout with Pass No/Sub Agent."""
    headers = [_normalise_import_header(ws.cell(1, col).value) for col in range(1, min(ws.max_column, 9) + 1)]
    return (
        len(headers) >= 9
        or 'pass no' in headers
        or 'pass number' in headers
        or 'sub agent' in headers
        or (len(headers) >= 9 and headers[8] == 'balance')
    )


def _canonical_account_row(values, legacy=False):
    """Convert either the old 9-column row or new 7-column row to one canonical layout.

    Canonical columns: No, Ref No, Date, Subject, In Payment, Out Payment, Balance.
    Old Pass No and Sub Agent values are intentionally discarded.
    """
    row = list(values)
    needed = 9 if legacy else 7
    if len(row) < needed:
        row.extend([None] * (needed - len(row)))
    if legacy:
        return [row[0], row[1], row[2], row[3], row[5], row[6], row[8]]
    return row[:7]


def _collect_normal_transactions(ws):
    """Read all Accounts transactions using the new 7-column schema.

    Existing 9-column workbooks are read safely and automatically migrated the next
    time the sheet is repaired/saved. Pass No and Sub Agent are dropped permanently.
    """
    opening_row = None
    transactions = []
    legacy = _record_sheet_is_legacy(ws)

    for source_order, values in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        raw = list(values)
        if raw and str(raw[0] or '').strip().upper() == CHEQUE_SECTION_MARKER:
            break

        row = _canonical_account_row(raw, legacy=legacy)
        ref_text = str(row[1] or '').strip()
        if ref_text.upper() == 'OPENING':
            opening_row = row
            continue

        has_transaction_data = bool(
            ref_text
            or str(row[2] or '').strip()
            or str(row[3] or '').strip()
            or parse_amount(row[4], 0)
            or parse_amount(row[5], 0)
        )

        if ref_text.upper().startswith('UTA-') or (str(row[2] or '').strip() and str(row[3] or '').strip()):
            transactions.append({'source_order': source_order, 'row': row})
        elif has_transaction_data and (str(row[2] or '').strip() or str(row[3] or '').strip()):
            transactions.append({'source_order': source_order, 'row': row})

    return opening_row, transactions


def _atomic_save_workbook(wb, filename):
    """Save an Excel workbook via a temporary file, then replace the live file atomically."""
    directory = os.path.dirname(filename) or '.'
    os.makedirs(directory, exist_ok=True)
    fd, temp_path = tempfile.mkstemp(prefix='.uta-save-', suffix='.xlsx', dir=directory)
    os.close(fd)
    try:
        wb.save(temp_path)
        os.replace(temp_path, filename)
    finally:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass


def repair_month_records_file(month_name, sync_cheques=True):
    """Repair ordering, numbering, references and balances for one month in-place."""
    filename = os.path.join(EXCEL_DIR, f"{month_name}.xlsx")
    if not os.path.exists(filename):
        return 0

    with RECORD_IO_LOCK:
        wb = load_workbook(filename)
        ws = wb.active
        remove_cheque_section(ws)
        count = recalculate_and_sort_sheet(ws)
        _atomic_save_workbook(wb, filename)
        wb.close()
        if sync_cheques:
            month_info = load_months_data().get(month_name, {})
            sync_cheque_deposits_to_excel(
                month_name,
                include_all=(month_info.get('status') == 'closed')
            )
        return count


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

        for col in [5, 6, 7]:

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

    headers = ['No', 'Ref No', 'Date', 'Subject', 'In Payment', 'Out Payment', 'Balance']

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

    ws.append([1, 'OPENING', opening_date, 'OPENING BALANCE', 0, 0, opening_balance])

    # Style opening balance row

    opening_fill = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")

    opening_font = Font(bold=True, color="2E7D32")

    for col in range(1, 8):

        cell = ws.cell(2, col)

        cell.fill = opening_fill

        cell.font = opening_font

    for col in [5, 6, 7]:

        ws.cell(2, col).number_format = '#,##0'



    # Adjust column widths

    ws.column_dimensions['A'].width = 8

    ws.column_dimensions['B'].width = 12

    ws.column_dimensions['C'].width = 12

    ws.column_dimensions['D'].width = 30

    ws.column_dimensions['E'].width = 15

    ws.column_dimensions['F'].width = 15

    ws.column_dimensions['G'].width = 15

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
    """Return the next sequential UTA reference from the actual transaction count."""
    filename = os.path.join(EXCEL_DIR, f"{month_name}.xlsx")
    if not os.path.exists(filename):
        return 1
    with RECORD_IO_LOCK:
        wb = load_workbook(filename, data_only=False)
        ws = wb.active
        _, transactions = _collect_normal_transactions(ws)
        wb.close()
    return len(transactions) + 1




def get_next_row_number(month_name):
    """Return the next display row number after the opening-balance row."""
    filename = os.path.join(EXCEL_DIR, f"{month_name}.xlsx")
    if not os.path.exists(filename):
        return 2
    with RECORD_IO_LOCK:
        wb = load_workbook(filename, data_only=False)
        ws = wb.active
        _, transactions = _collect_normal_transactions(ws)
        wb.close()
    return len(transactions) + 2




def get_current_balance(month_name):
    """Calculate the balance from opening balance + all normal transactions."""
    filename = os.path.join(EXCEL_DIR, f"{month_name}.xlsx")
    if not os.path.exists(filename):
        return 0
    with RECORD_IO_LOCK:
        wb = load_workbook(filename, data_only=False)
        ws = wb.active
        opening_row, transactions = _collect_normal_transactions(ws)
        wb.close()
    if opening_row is None:
        return 0
    balance = parse_amount(opening_row[6], 0)
    for item in transactions:
        row = item['row']
        balance += parse_amount(row[4], 0) - parse_amount(row[5], 0)
    return balance




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
    invoices = load_candidate_invoices()
    invoices.sort(key=lambda x: str(x.get('updatedAt') or x.get('createdAt', '')), reverse=True)

    salary_slips = load_candidate_salary_slips()
    salary_slips.sort(key=lambda x: str(x.get('updatedAt') or x.get('createdAt', '')), reverse=True)

    today_key = datetime.now().strftime('%Y-%m-%d')
    total_invoice_value = 0.0
    total_salary_value = 0.0
    today_invoices = 0
    today_salary_slips = 0

    for invoice in invoices:
        try:
            total_invoice_value += float(invoice.get('grandTotal', 0) or 0)
        except (TypeError, ValueError):
            pass
        if str(invoice.get('date', '')).strip() == today_key:
            today_invoices += 1

    for slip in salary_slips:
        try:
            total_salary_value += float(slip.get('totalAmount', 0) or 0)
        except (TypeError, ValueError):
            pass
        if str(slip.get('date', '')).strip() == today_key:
            today_salary_slips += 1

    return render_template(
        'candidate/dashboard.html',
        current_user=get_candidate_user(),
        current_date=datetime.now().strftime('%A, %d %B %Y'),
        total_invoices=len(invoices),
        today_invoices=today_invoices,
        total_invoice_value=total_invoice_value,
        recent_invoices=invoices[:5],
        total_salary_slips=len(salary_slips),
        today_salary_slips=today_salary_slips,
        total_salary_value=total_salary_value,
        recent_salary_slips=salary_slips[:5]
    )


@app.route('/candidate/invoice')
@candidate_login_required
def candidate_invoice():
    return render_template(
        'candidate/invoice.html',
        current_user=get_candidate_user(),
        current_date=datetime.now().strftime('%A, %d %B %Y')
    )


@app.route('/candidate/salary-slip')
@candidate_login_required
def candidate_salary_slip():
    return render_template(
        'candidate/salary_slip.html',
        current_user=get_candidate_user(),
        current_date=datetime.now().strftime('%A, %d %B %Y')
    )


# ---------------- Candidate Invoice API ----------------

@app.route('/api/candidate/invoices', methods=['GET'])
@candidate_login_required
def candidate_invoices_api():
    invoices = load_candidate_invoices()
    invoices.sort(key=lambda x: str(x.get('updatedAt') or x.get('createdAt', '')), reverse=True)
    return jsonify({'success': True, 'invoices': invoices})


@app.route('/api/candidate/invoices/next-number', methods=['GET'])
@candidate_login_required
def candidate_invoice_next_number_api():
    return jsonify({'success': True, 'invoiceNo': next_candidate_invoice_number()})


def _clean_invoice_payload(data):
    client_name = str(data.get('clientName', '')).strip()
    if not client_name:
        return None, 'Client name is required.'

    items = data.get('items', [])
    if not isinstance(items, list):
        return None, 'Invalid invoice items.'

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

    if not clean_items:
        return None, 'At least one invoice item is required.'

    return {
        'date': str(data.get('date', '')).strip() or datetime.now().strftime('%Y-%m-%d'),
        'clientName': client_name,
        'passportNo': str(data.get('passportNo', '')).strip(),
        'destination': str(data.get('destination', '')).strip(),
        'items': clean_items,
        'subtotal': subtotal,
        'vat': 0,
        'grandTotal': subtotal,
    }, None


@app.route('/api/candidate/invoices', methods=['POST'])
@candidate_login_required
def candidate_invoice_create_api():
    data = request.get_json(silent=True) or {}
    clean, error = _clean_invoice_payload(data)
    if error:
        return jsonify({'success': False, 'message': error}), 400

    invoices = load_candidate_invoices()
    invoice_no = str(data.get('invoiceNo', '')).strip() or next_candidate_invoice_number()
    if any(str(inv.get('invoiceNo', '')).strip().lower() == invoice_no.lower() for inv in invoices):
        invoice_no = next_candidate_invoice_number()

    invoice = {
        'id': int(datetime.now().timestamp() * 1000),
        'invoiceNo': invoice_no,
        **clean,
        'createdAt': datetime.now().isoformat(timespec='seconds'),
        'createdBy': get_candidate_user()
    }
    invoices.append(invoice)
    save_candidate_invoices(invoices)
    return jsonify({'success': True, 'message': 'Invoice saved successfully.', 'invoice': invoice})


@app.route('/api/candidate/invoices/<int:invoice_id>', methods=['GET'])
@candidate_login_required
def candidate_invoice_get_api(invoice_id):
    invoice = next((inv for inv in load_candidate_invoices() if int(inv.get('id', 0) or 0) == invoice_id), None)
    if not invoice:
        return jsonify({'success': False, 'message': 'Invoice not found.'}), 404
    return jsonify({'success': True, 'invoice': invoice})


@app.route('/api/candidate/invoices/<int:invoice_id>', methods=['PUT'])
@candidate_login_required
def candidate_invoice_update_api(invoice_id):
    data = request.get_json(silent=True) or {}
    clean, error = _clean_invoice_payload(data)
    if error:
        return jsonify({'success': False, 'message': error}), 400

    invoices = load_candidate_invoices()
    index = next((i for i, inv in enumerate(invoices) if int(inv.get('id', 0) or 0) == invoice_id), None)
    if index is None:
        return jsonify({'success': False, 'message': 'Invoice not found.'}), 404

    existing = invoices[index]
    invoice_no = str(data.get('invoiceNo', '')).strip() or str(existing.get('invoiceNo', '')).strip() or next_candidate_invoice_number()
    if any(i != index and str(inv.get('invoiceNo', '')).strip().lower() == invoice_no.lower() for i, inv in enumerate(invoices)):
        return jsonify({'success': False, 'message': 'Invoice number already exists.'}), 400

    updated = {
        **existing,
        'invoiceNo': invoice_no,
        **clean,
        'updatedAt': datetime.now().isoformat(timespec='seconds'),
        'updatedBy': get_candidate_user()
    }
    invoices[index] = updated
    save_candidate_invoices(invoices)
    return jsonify({'success': True, 'message': 'Invoice updated successfully.', 'invoice': updated})


@app.route('/api/candidate/invoices/<int:invoice_id>', methods=['DELETE'])
@candidate_login_required
def candidate_invoice_delete_api(invoice_id):
    invoices = load_candidate_invoices()
    new_invoices = [inv for inv in invoices if int(inv.get('id', 0) or 0) != invoice_id]
    if len(new_invoices) == len(invoices):
        return jsonify({'success': False, 'message': 'Invoice not found.'}), 404
    save_candidate_invoices(new_invoices)
    return jsonify({'success': True, 'message': 'Invoice deleted successfully.'})


# ---------------- Candidate Salary Slip API ----------------

@app.route('/api/candidate/salary-slips', methods=['GET'])
@candidate_login_required
def candidate_salary_slips_api():
    slips = load_candidate_salary_slips()
    slips.sort(key=lambda x: str(x.get('updatedAt') or x.get('createdAt', '')), reverse=True)
    return jsonify({'success': True, 'salarySlips': slips})


def _clean_salary_slip_payload(data):
    full_name = str(data.get('fullName', '')).strip()
    if not full_name:
        return None, 'Full name is required.'

    salary_month = str(data.get('salaryMonth', '')).strip()
    if not salary_month:
        return None, 'Salary month is required.'

    items = data.get('items', [])
    if not isinstance(items, list):
        return None, 'Invalid salary items.'

    clean_items = []
    total_amount = 0.0
    for item in items:
        if not isinstance(item, dict):
            continue
        description = str(item.get('description', '')).strip()
        try:
            amount = float(item.get('amount', 0) or 0)
        except (TypeError, ValueError):
            amount = 0.0
        clean_items.append({'description': description, 'amount': amount})
        total_amount += amount

    if not clean_items:
        return None, 'At least one salary item is required.'

    return {
        'fullName': full_name,
        'fullAddress': str(data.get('fullAddress', '')).strip(),
        'nicNo': str(data.get('nicNo', '')).strip(),
        'date': str(data.get('date', '')).strip() or datetime.now().strftime('%Y-%m-%d'),
        'salaryMonth': salary_month,
        'items': clean_items,
        'totalAmount': total_amount,
    }, None


@app.route('/api/candidate/salary-slips', methods=['POST'])
@candidate_login_required
def candidate_salary_slip_create_api():
    data = request.get_json(silent=True) or {}
    clean, error = _clean_salary_slip_payload(data)
    if error:
        return jsonify({'success': False, 'message': error}), 400

    slips = load_candidate_salary_slips()
    slip = {
        'id': int(datetime.now().timestamp() * 1000),
        'slipNo': next_candidate_salary_slip_number(),
        **clean,
        'createdAt': datetime.now().isoformat(timespec='seconds'),
        'createdBy': get_candidate_user()
    }
    slips.append(slip)
    save_candidate_salary_slips(slips)
    return jsonify({'success': True, 'message': 'Salary slip saved successfully.', 'salarySlip': slip})


@app.route('/api/candidate/salary-slips/<int:slip_id>', methods=['GET'])
@candidate_login_required
def candidate_salary_slip_get_api(slip_id):
    slip = next((row for row in load_candidate_salary_slips() if int(row.get('id', 0) or 0) == slip_id), None)
    if not slip:
        return jsonify({'success': False, 'message': 'Salary slip not found.'}), 404
    return jsonify({'success': True, 'salarySlip': slip})


@app.route('/api/candidate/salary-slips/<int:slip_id>', methods=['PUT'])
@candidate_login_required
def candidate_salary_slip_update_api(slip_id):
    data = request.get_json(silent=True) or {}
    clean, error = _clean_salary_slip_payload(data)
    if error:
        return jsonify({'success': False, 'message': error}), 400

    slips = load_candidate_salary_slips()
    index = next((i for i, row in enumerate(slips) if int(row.get('id', 0) or 0) == slip_id), None)
    if index is None:
        return jsonify({'success': False, 'message': 'Salary slip not found.'}), 404

    existing = slips[index]
    updated = {
        **existing,
        **clean,
        'updatedAt': datetime.now().isoformat(timespec='seconds'),
        'updatedBy': get_candidate_user()
    }
    slips[index] = updated
    save_candidate_salary_slips(slips)
    return jsonify({'success': True, 'message': 'Salary slip updated successfully.', 'salarySlip': updated})


@app.route('/api/candidate/salary-slips/<int:slip_id>', methods=['DELETE'])
@candidate_login_required
def candidate_salary_slip_delete_api(slip_id):
    slips = load_candidate_salary_slips()
    new_slips = [row for row in slips if int(row.get('id', 0) or 0) != slip_id]
    if len(new_slips) == len(slips):
        return jsonify({'success': False, 'message': 'Salary slip not found.'}), 404
    save_candidate_salary_slips(new_slips)
    return jsonify({'success': True, 'message': 'Salary slip deleted successfully.'})


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



@app.route('/preview_excel_records', methods=['POST'])
@login_required
def preview_excel_records():
    """Read an uploaded Excel file in memory and return a safe import preview.

    Nothing is written to the live monthly workbook here. The browser loads valid rows
    into the existing Enter Records cards, and the normal /save_records route performs
    the actual save so numbering, date ordering and balances use one canonical path.
    """
    try:
        active_month = get_active_month()
        if not active_month:
            return jsonify({'success': False, 'message': 'No active month. Create or activate a month first.'}), 400

        uploaded = request.files.get('excel_file')
        if uploaded is None or not uploaded.filename:
            return jsonify({'success': False, 'message': 'Choose an Excel file first.'}), 400

        extension = os.path.splitext(uploaded.filename)[1].lower()
        if extension not in {'.xlsx', '.xlsm'}:
            return jsonify({
                'success': False,
                'message': 'Please upload an .xlsx or .xlsm Excel file. For old .xls files, open them in Excel and Save As .xlsx first.'
            }), 400

        raw = uploaded.read()
        if not raw:
            return jsonify({'success': False, 'message': 'The uploaded Excel file is empty.'}), 400

        try:
            wb = load_workbook(BytesIO(raw), data_only=True, read_only=True)
        except Exception as exc:
            return jsonify({'success': False, 'message': f'Unable to read this Excel file: {exc}'}), 400

        ws = wb.active
        header_row, mapping = _find_excel_import_header(ws)
        if not header_row:
            wb.close()
            return jsonify({
                'success': False,
                'message': (
                    'Could not identify the Excel headings. The sheet must contain at least Date and Subject/Description columns. '
                    'Supported optional columns include In Payment and Out Payment.'
                )
            }), 400

        rows = []
        skipped_blank = 0
        skipped_special = 0
        max_import_rows = 500

        for row_no in range(header_row + 1, ws.max_row + 1):
            # Stop before a UTA cheque-deposit section when importing one of our exported sheets.
            first_values = [ws.cell(row_no, col).value for col in range(1, min(ws.max_column, 3) + 1)]
            if any(str(v or '').strip().upper() == CHEQUE_SECTION_MARKER for v in first_values):
                break

            date_raw = _excel_cell_value(ws, row_no, mapping, 'date')
            subject_raw = _excel_cell_value(ws, row_no, mapping, 'subject')
            in_raw = _excel_cell_value(ws, row_no, mapping, 'in_payment')
            out_raw = _excel_cell_value(ws, row_no, mapping, 'out_payment')

            raw_values = [date_raw, subject_raw, in_raw, out_raw]
            if all(v is None or str(v).strip() == '' for v in raw_values):
                skipped_blank += 1
                continue

            subject_text = str(subject_raw or '').strip()
            # Skip opening/total-style rows from a UTA workbook instead of importing them as transactions.
            combined = ' '.join(str(v or '').strip().upper() for v in first_values + [subject_raw])
            if 'OPENING BALANCE' in combined or combined.strip() == 'OPENING' or subject_text.upper() == 'OPENING BALANCE':
                skipped_special += 1
                continue
            if subject_text.upper() in {'TOTAL', 'TOTAL AMOUNT', 'GRAND TOTAL'}:
                skipped_special += 1
                continue

            date_value = _normalise_record_date(date_raw)
            in_payment = _parse_excel_import_amount(in_raw)
            out_payment = _parse_excel_import_amount(out_raw)
            errors = []

            if not date_value or _record_date_sort_key(date_value).year == 9999:
                errors.append('Valid date required')
            if not subject_text:
                errors.append('Subject required')

            rows.append({
                'source_row': row_no,
                'date': date_value if _record_date_sort_key(date_value).year != 9999 else str(date_raw or '').strip(),
                'subject': subject_text,
                'in_payment': in_payment,
                'out_payment': out_payment,
                'valid': not errors,
                'errors': errors,
            })

            if len(rows) >= max_import_rows:
                break

        wb.close()

        valid_rows = [row for row in rows if row['valid']]
        invalid_rows = [row for row in rows if not row['valid']]
        total_in = sum(float(row['in_payment'] or 0) for row in valid_rows)
        total_out = sum(float(row['out_payment'] or 0) for row in valid_rows)

        response = jsonify({
            'success': True,
            'filename': secure_filename(uploaded.filename),
            'sheet_name': ws.title,
            'header_row': header_row,
            'rows': rows,
            'total_rows': len(rows),
            'valid_count': len(valid_rows),
            'invalid_count': len(invalid_rows),
            'total_in': total_in,
            'total_out': total_out,
            'net': total_in - total_out,
            'skipped_blank': skipped_blank,
            'skipped_special': skipped_special,
            'limited': len(rows) >= max_import_rows and ws.max_row > header_row + max_import_rows,
        })
        response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
        return response
    except Exception as e:
        print(f'Excel import preview error: {e}')
        return jsonify({'success': False, 'message': f'Excel import error: {str(e)}'}), 500


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
    """Delete one record, then rebuild every No/UTA reference/balance safely."""
    try:
        data = request.get_json(silent=True) or {}
        month_name = data.get('month_name')
        ref_no = str(data.get('ref_no') or data.get('record_no') or '').strip()
        if not month_name or not ref_no:
            return jsonify({'success': False, 'message': 'Missing required data'})

        filename = os.path.join(EXCEL_DIR, f"{month_name}.xlsx")
        if not os.path.exists(filename):
            return jsonify({'success': False, 'message': 'Month file not found'})

        with RECORD_IO_LOCK:
            wb = load_workbook(filename)
            ws = wb.active
            remove_cheque_section(ws)
            before_count = recalculate_and_sort_sheet(ws)

            row_to_delete = None
            for row_no in range(3, ws.max_row + 1):
                if str(ws.cell(row_no, 2).value or '').strip().upper() == ref_no.upper():
                    row_to_delete = row_no
                    break

            if row_to_delete is None:
                wb.close()
                return jsonify({'success': False, 'message': 'Record not found. Refresh the page and try again.'}), 404

            create_backup('delete_record', month_name)
            ws.delete_rows(row_to_delete)
            after_count = recalculate_and_sort_sheet(ws)
            if after_count != before_count - 1:
                wb.close()
                raise RuntimeError('Delete integrity check failed. Nothing was saved.')

            _atomic_save_workbook(wb, filename)
            wb.close()
            sync_cheque_deposits_to_excel(month_name)

        return jsonify({
            'success': True,
            'message': 'Record deleted. Dates, row numbers, references and balances were rebuilt successfully.',
            'total_records': after_count,
            'next_reference': f"UTA-{after_count + 1:02d}",
        })
    except Exception as e:
        print(f"Error deleting record: {e}")
        return jsonify({'success': False, 'message': f'Error deleting: {str(e)}'}), 500




@app.route('/update_record_in_month', methods=['POST'])
@login_required
def update_record_in_month():
    """Edit one record, then re-sort and rebuild references/balances without changing record count."""
    try:
        data = request.get_json(silent=True) or {}
        month_name = data.get('month_name')
        ref_no = str(data.get('ref_no') or '').strip()
        date_value = _normalise_record_date(data.get('date'))
        subject = str(data.get('subject') or '').strip()

        if not month_name or not ref_no:
            return jsonify({'success': False, 'message': 'Missing required data'})
        if not date_value or not subject:
            return jsonify({'success': False, 'message': 'Date and Subject are required'})
        if _record_date_sort_key(date_value).year == 9999:
            return jsonify({'success': False, 'message': 'Invalid date'})

        filename = os.path.join(EXCEL_DIR, f"{month_name}.xlsx")
        if not os.path.exists(filename):
            return jsonify({'success': False, 'message': 'Month file not found'})

        with RECORD_IO_LOCK:
            wb = load_workbook(filename)
            ws = wb.active
            remove_cheque_section(ws)
            before_count = recalculate_and_sort_sheet(ws)

            row_to_update = None
            for row_no in range(3, ws.max_row + 1):
                if str(ws.cell(row_no, 2).value or '').strip().upper() == ref_no.upper():
                    row_to_update = row_no
                    break

            if row_to_update is None:
                wb.close()
                return jsonify({'success': False, 'message': 'Record not found. Refresh the page and try again.'}), 404

            create_backup('update_record', month_name)
            ws.cell(row_to_update, 3).value = date_value
            ws.cell(row_to_update, 4).value = subject
            ws.cell(row_to_update, 5).value = parse_amount(data.get('in_payment', 0), 0)
            ws.cell(row_to_update, 6).value = parse_amount(data.get('out_payment', 0), 0)

            after_count = recalculate_and_sort_sheet(ws)
            if after_count != before_count:
                wb.close()
                raise RuntimeError('Update integrity check failed. Nothing was saved.')

            _atomic_save_workbook(wb, filename)
            wb.close()
            sync_cheque_deposits_to_excel(month_name)

        return jsonify({
            'success': True,
            'message': 'Record updated. Dates, row numbers, references and balances were rebuilt successfully.',
            'total_records': after_count,
            'next_reference': f"UTA-{after_count + 1:02d}",
        })
    except Exception as e:
        print(f"Error updating record: {e}")
        return jsonify({'success': False, 'message': f'Error updating: {str(e)}'}), 500




@app.route('/get_month_records/<month_name>')
@login_required
def get_month_records(month_name):
    """Return a repaired, canonical list of normal transactions for a month."""
    try:
        filename = os.path.join(EXCEL_DIR, f"{month_name}.xlsx")
        if not os.path.exists(filename):
            return jsonify({'records': [], 'next_reference': 'UTA-01'})

        # Repair legacy gaps/misaligned row numbers before displaying or editing them.
        repair_month_records_file(month_name)

        with RECORD_IO_LOCK:
            wb = load_workbook(filename, data_only=False)
            ws = wb.active
            records = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                if str(row[0] or '').strip().upper() == CHEQUE_SECTION_MARKER:
                    break
                ref_text = str(row[1] or '').strip()
                if ref_text.upper() == 'OPENING':
                    continue
                if ref_text.upper().startswith('UTA-'):
                    records.append({
                        'no': int(row[0]) if isinstance(row[0], (int, float)) else len(records) + 2,
                        'ref_no': ref_text,
                        'date': _normalise_record_date(row[2]),
                        'subject': row[3] or '',
                        'in_payment': parse_amount(row[4], 0),
                        'out_payment': parse_amount(row[5], 0),
                        'balance': parse_amount(row[6], 0),
                    })
            wb.close()

        response = jsonify({
            'records': records,
            'next_reference': f"UTA-{len(records) + 1:02d}",
        })
        response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
        return response
    except Exception as e:
        print(f"Error getting records: {e}")
        return jsonify({'records': [], 'error': str(e)}), 500




def recalculate_and_sort_sheet(ws):
    """Canonicalise Accounts records into the 7-column layout.

    Columns are permanently: No, Ref No, Date, Subject, In Payment, Out Payment, Balance.
    Legacy Pass No/Sub Agent columns are removed during this rebuild.
    """
    opening_row_data, transaction_items = _collect_normal_transactions(ws)
    if opening_row_data is None:
        raise ValueError('Opening balance row is missing. Save cancelled to protect existing data.')

    transaction_items.sort(key=lambda item: (_record_date_sort_key(item['row'][2]), item['source_order']))

    if ws.max_row >= 2:
        ws.delete_rows(2, ws.max_row - 1)
    if ws.max_column > 7:
        ws.delete_cols(8, ws.max_column - 7)

    headers = ['No', 'Ref No', 'Date', 'Subject', 'In Payment', 'Out Payment', 'Balance']
    header_fill = PatternFill(start_color='1A1A2E', end_color='1A1A2E', fill_type='solid')
    header_font = Font(bold=True, color='FFFFFF')
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(1, col)
        cell.value = header
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center')

    opening_balance = parse_amount(opening_row_data[6], 0)
    opening_date = _normalise_record_date(opening_row_data[2])
    ws.append([1, 'OPENING', opening_date, opening_row_data[3] or 'OPENING BALANCE', 0, 0, opening_balance])

    opening_row_idx = ws.max_row
    opening_fill = PatternFill(start_color='E8F5E9', end_color='E8F5E9', fill_type='solid')
    opening_font = Font(bold=True, color='2E7D32')
    for col in range(1, 8):
        ws.cell(opening_row_idx, col).fill = opening_fill
        ws.cell(opening_row_idx, col).font = opening_font
    for col in [5, 6, 7]:
        ws.cell(opening_row_idx, col).number_format = '#,##0'

    balance = opening_balance
    for index, item in enumerate(transaction_items, start=1):
        row = item['row']
        in_payment = parse_amount(row[4], 0)
        out_payment = parse_amount(row[5], 0)
        balance += in_payment - out_payment
        ws.append([
            index + 1,
            f"UTA-{index:02d}",
            _normalise_record_date(row[2]),
            str(row[3] or '').strip(),
            in_payment,
            out_payment,
            balance,
        ])
        for col in [5, 6, 7]:
            ws.cell(ws.max_row, col).number_format = '#,##0'

    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 34
    ws.column_dimensions['E'].width = 16
    ws.column_dimensions['F'].width = 16
    ws.column_dimensions['G'].width = 16

    return len(transaction_items)





@app.route('/save_records', methods=['POST'])
@login_required
def save_records():
    """Save one or more records without losing or replacing existing transactions."""
    try:
        active_month = get_active_month()
        if not active_month:
            return jsonify({'success': False, 'message': 'No active month! Please create a new month first.'})

        payload = request.get_json(silent=True) or {}
        records = payload.get('records', [])
        if not isinstance(records, list) or not records:
            return jsonify({'success': False, 'message': 'No records to save'})

        cleaned_records = []
        for index, record in enumerate(records, start=1):
            if not isinstance(record, dict):
                return jsonify({'success': False, 'message': f'Row {index}: invalid record data'})
            date_value = _normalise_record_date(record.get('date'))
            subject = str(record.get('subject') or '').strip()
            if not date_value or not subject:
                return jsonify({'success': False, 'message': f'Row {index}: Date and Subject are required'})
            if _record_date_sort_key(date_value).year == 9999:
                return jsonify({'success': False, 'message': f'Row {index}: invalid date'})
            cleaned_records.append({
                'date': date_value,
                'subject': subject,
                'in_payment': parse_amount(record.get('in_payment', 0), 0),
                'out_payment': parse_amount(record.get('out_payment', 0), 0),
            })

        filename = os.path.join(EXCEL_DIR, f"{active_month}.xlsx")
        if not os.path.exists(filename):
            return jsonify({'success': False, 'message': 'Month file not found!'})

        with RECORD_IO_LOCK:
            wb = load_workbook(filename)
            ws = wb.active
            remove_cheque_section(ws)

            # First repair legacy gaps so the existing record count is trustworthy.
            before_count = recalculate_and_sort_sheet(ws)

            for record in cleaned_records:
                # Reference/row/balance are deliberately temporary; the canonical rebuild below
                # assigns them after date sorting so inserting an older date cannot overwrite data.
                ws.append([
                    0,
                    'PENDING',
                    record['date'],
                    record['subject'],
                    record['in_payment'],
                    record['out_payment'],
                    0,
                ])

            after_count = recalculate_and_sort_sheet(ws)
            expected_count = before_count + len(cleaned_records)
            if after_count != expected_count:
                wb.close()
                raise RuntimeError(
                    f'Data integrity check failed: expected {expected_count} records, found {after_count}. Nothing was saved.'
                )

            _atomic_save_workbook(wb, filename)
            wb.close()
            sync_cheque_deposits_to_excel(active_month)

        response = jsonify({
            'success': True,
            'message': f'Successfully saved {len(cleaned_records)} record(s)!',
            'saved_count': len(cleaned_records),
            'total_records': after_count,
            'next_reference': f"UTA-{after_count + 1:02d}",
        })
        response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
        return response
    except Exception as e:
        print(f"Error saving records: {e}")
        return jsonify({'success': False, 'message': f'Error saving: {str(e)}'}), 500


@app.route('/get_next_reference')
@login_required
def get_next_reference():
    """Repair the active month if needed and return the true next sequential reference."""
    active_month = get_active_month()
    if not active_month:
        return jsonify({'next_reference': 'No Active Month'})
    try:
        total_records = repair_month_records_file(active_month)
        response = jsonify({'next_reference': f"UTA-{total_records + 1:02d}"})
        response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
        return response
    except Exception as e:
        print(f"Error getting next reference: {e}")
        return jsonify({'next_reference': 'Error', 'message': str(e)}), 500




@app.route('/get_stats')
@login_required
def get_stats():
    """Dashboard statistics for the active month using canonical transaction rows only."""
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
            repair_month_records_file(active_month)
            with RECORD_IO_LOCK:
                wb = load_workbook(filename, data_only=False)
                ws = wb.active
                opening_row, transactions = _collect_normal_transactions(ws)
                opening_balance = parse_amount(opening_row[6], 0) if opening_row else 0
                total_records = len(transactions)
                total_references = len(transactions)
                for item in transactions:
                    row = item['row']
                    total_in_payment += parse_amount(row[4], 0)
                    total_out_payment += parse_amount(row[5], 0)
                wb.close()
                net_balance = opening_balance + total_in_payment - total_out_payment

    response = jsonify({
        'total_months': total_months,
        'total_records': total_records,
        'total_in_payment': total_in_payment,
        'total_out_payment': total_out_payment,
        'total_references': total_references,
        'net_balance': net_balance,
        'active_month': active_month,
        'active_month_display': months_data.get(active_month, {}).get('display_name', 'None') if active_month else 'None',
    })
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    return response




@app.route('/get_month_stats/<month_name>')
@login_required
def get_month_stats(month_name):
    """Statistics for one month; opening balance is read from Ref No = OPENING."""
    try:
        filename = os.path.join(EXCEL_DIR, f"{month_name}.xlsx")
        if not os.path.exists(filename):
            return jsonify({'error': 'Month not found'}), 404

        repair_month_records_file(month_name)
        with RECORD_IO_LOCK:
            wb = load_workbook(filename, data_only=False)
            ws = wb.active
            opening_row, transactions = _collect_normal_transactions(ws)
            opening_balance = parse_amount(opening_row[6], 0) if opening_row else 0
            total_in_payment = 0
            total_out_payment = 0
            for item in transactions:
                row = item['row']
                total_in_payment += parse_amount(row[4], 0)
                total_out_payment += parse_amount(row[5], 0)
            wb.close()

        closing_balance = opening_balance + total_in_payment - total_out_payment
        response = jsonify({
            'total_records': len(transactions),
            'total_in_payment': total_in_payment,
            'total_out_payment': total_out_payment,
            'total_references': len(transactions),
            'opening_balance': opening_balance,
            'closing_balance': closing_balance,
            'net_change': closing_balance - opening_balance,
        })
        response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
        return response
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
    """Download a repaired month workbook."""
    try:
        filename = os.path.join(EXCEL_DIR, f"{month_name}.xlsx")
        if not os.path.exists(filename):
            return jsonify({'error': 'File not found'}), 404
        repair_month_records_file(month_name, sync_cheques=False)
        month_info = load_months_data().get(month_name, {})
        sync_cheque_deposits_to_excel(month_name, include_all=(month_info.get('status') == 'closed'))
        apply_excel_amount_format(filename)
        return send_file(filename, as_attachment=True, download_name=f'UTA_{month_name}.xlsx')
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


