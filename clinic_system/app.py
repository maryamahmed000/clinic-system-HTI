from flask import Flask, render_template, request, jsonify, redirect, url_for, session
import hashlib
import pandas as pd
import psycopg2
from psycopg2.extras import RealDictCursor
import os
# Load OneDrive credentials from .env file if present (no extra library needed)
def _load_env_file():
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env')
    if os.path.exists(env_path):
        with open(env_path, encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith('#') or '=' not in line:
                    continue
                key, _, val = line.partition('=')
                os.environ.setdefault(key.strip(), val.strip())
_load_env_file()
from datetime import datetime, date
import json

app = Flask(__name__)
app.secret_key = 'hti-clinic-secret-2024-oct'
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['PERMANENT_SESSION_LIFETIME'] = 28800  # 8 hours
# لا تجعل الكوكي دائمة — تنتهي الجلسة عند إغلاق المتصفح بدلاً من البقاء مسجلاً للدخول دائماً
app.config['SESSION_REFRESH_EACH_REQUEST'] = True
app.config['SESSION_PERMANENT'] = False
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

@app.template_filter('fromjson')
def fromjson_filter(value):
    """يحوّل نص JSON المخزن في قاعدة البيانات إلى list/dict داخل القوالب."""
    try:
        return json.loads(value) if value else []
    except (TypeError, ValueError):
        return []

# ─── Database setup ──────────────────────────────────────────
def get_db():
    DATABASE_URL = os.environ.get('DATABASE_URL', '')
    conn = psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)
    conn.autocommit = False
    return conn


def db_execute(conn, sql, params=None):
    """Helper: execute SQL and return cursor."""
    cur = conn.cursor()
    cur.execute(sql, params or ())
    return cur


def db_fetchall(conn, sql, params=None):
    cur = db_execute(conn, sql, params)
    return cur.fetchall()


def db_fetchone(conn, sql, params=None):
    cur = db_execute(conn, sql, params)
    return cur.fetchone()


def _load_medicines_from_excel(cursor):
    """Load medicines from Excel file into DB."""
    import openpyxl
    path = os.path.join(BASE_DIR, 'data', 'medicines.xlsx')
    if not os.path.exists(path):
        return
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb.active

    def categorize(name):
        n = name
        if any(x in n for x in ['شاش','قطن','رباط','بلاستر','شريط','مسحة','خيط','مقارش','جبسونا','بيتادين','ماء أكسجين','خافض']):
            return 'مستلزمات الضماد'
        elif any(x in n for x in ['جوانتى','سرنجات','سن مشرط','كنيولا','شكاكات','كاشفات']):
            return 'أدوات طبية'
        elif any(x in n for x in ['ماسك','Ecg','ECG','أجهزة','ترمومتر']):
            return 'أجهزة وأدوات'
        elif 'حقن' in n:
            return 'حقن'
        elif any(x in n for x in ['كريم','مرهم']):
            return 'كريمات ومراهم'
        elif any(x in n for x in ['محلول','رينجر','جلوكوز','ملح 500','فاركولين','زيلوكين']):
            return 'محاليل'
        elif any(x in n for x in ['قطره','سبراى','نقط']):
            return 'قطرات وسبراى'
        elif 'لزقة' in n:
            return 'مستلزمات الضماد'
        else:
            return 'أقراص وكبسول'

    for row in ws.iter_rows(values_only=True):
        name = row[1]
        expiry_raw = row[6]
        if name and isinstance(name, str):
            name = name.strip()
            if name and name not in ('الصنف',):
                expiry = str(expiry_raw).strip() if expiry_raw else ''
                expiry = expiry.replace('None','').replace('_____','').strip()
                cat = categorize(name)
                cursor.execute(
                    "INSERT INTO medicines(name,stock,unit,category,expiry_date) VALUES(%s,%s,%s,%s,%s)",
                    (name, 0, 'وحدة', cat, expiry if expiry else None)
                )
    wb.close()

def init_db():
    conn = get_db()
    c = conn.cursor()

    c.execute('''CREATE TABLE IF NOT EXISTS visits (
        id                   SERIAL PRIMARY KEY,
        person_type          TEXT NOT NULL,
        person_id            TEXT NOT NULL,
        person_name          TEXT NOT NULL,
        department           TEXT,
        visit_date           TEXT NOT NULL,
        diagnosis            TEXT,
        treatment_type       TEXT,
        medicines            TEXT,
        prescription         TEXT,
        leave_days           INTEGER DEFAULT 0,
        leave_from           TEXT,
        leave_to             TEXT,
        is_chronic           INTEGER DEFAULT 0,
        report_reason        TEXT,
        notes                TEXT,
        certificate_name     TEXT,
        submitted_to_affairs INTEGER DEFAULT 0,
        doctor_name          TEXT,
        created_at           TEXT DEFAULT CURRENT_TIMESTAMP
    )''')

    # Add doctor_name to existing DBs that don't have it
    try:
        c.execute('ALTER TABLE visits ADD COLUMN doctor_name TEXT DEFAULT NULL')
        conn.commit()
    except Exception:
        conn.rollback()

    c.execute('''CREATE TABLE IF NOT EXISTS affairs_reports (
        id              SERIAL PRIMARY KEY,
        visit_id        INTEGER NOT NULL,
        report_number   TEXT,
        action_type     TEXT,
        leave_from      TEXT,
        leave_to        TEXT,
        leave_days      INTEGER,
        affairs_notes   TEXT,
        printed         INTEGER DEFAULT 0,
        created_at      TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(visit_id) REFERENCES visits(id)
    )''')

    c.execute('''CREATE TABLE IF NOT EXISTS email_log (
        id              SERIAL PRIMARY KEY,
        ar_id           INTEGER,
        visit_id        INTEGER,
        person_name     TEXT,
        person_id       TEXT,
        report_number   TEXT,
        to_emails       TEXT,
        sent_by         TEXT,
        status          TEXT DEFAULT 'success',
        error_message   TEXT,
        sent_at         TEXT DEFAULT CURRENT_TIMESTAMP
    )''')

    c.execute('''CREATE TABLE IF NOT EXISTS medicines (
        id          SERIAL PRIMARY KEY,
        name        TEXT NOT NULL,
        stock       INTEGER DEFAULT 0,
        unit        TEXT DEFAULT 'وحدة',
        category    TEXT DEFAULT NULL,
        expiry_date TEXT DEFAULT NULL
    )''')
    # Add columns to existing DBs
    for col, typedef in [('category','TEXT'), ('expiry_date','TEXT')]:
        try:
            c.execute(f"ALTER TABLE medicines ADD COLUMN {col} {typedef}")
            conn.commit()
        except Exception:
            conn.rollback()

    c.execute('''CREATE TABLE IF NOT EXISTS doctors (
        id       SERIAL PRIMARY KEY,
        username TEXT UNIQUE NOT NULL,
        password TEXT NOT NULL,
        name     TEXT NOT NULL,
        role     TEXT DEFAULT 'doctor'
    )''')
    
    # Seed doctors if empty
    c.execute("SELECT COUNT(*) FROM doctors")
    if list(c.fetchone().values())[0] == 0:
        def h(p): return hashlib.sha256(p.encode()).hexdigest()
        c.execute("INSERT INTO doctors (username,password,name) VALUES (%s,%s,%s)", ('doctor1', h('doctor1'), 'د. أحمد سالم'))
        c.execute("INSERT INTO doctors (username,password,name) VALUES (%s,%s,%s)", ('doctor2', h('doctor2'), 'د. محمد إبراهيم'))

    # Load medicines from Excel if table is empty
    c.execute("SELECT COUNT(*) FROM medicines")
    if list(c.fetchone().values())[0] == 0:
        _load_medicines_from_excel(c)

    # Create doctors table
    c.execute('''CREATE TABLE IF NOT EXISTS doctors (
        id       SERIAL PRIMARY KEY,
        username TEXT UNIQUE NOT NULL,
        password TEXT NOT NULL,
        name     TEXT NOT NULL,
        role     TEXT DEFAULT 'doctor'
    )''')

    def hash_pw(pw):
        return hashlib.sha256(pw.encode()).hexdigest()

    # Seed doctors — INSERT OR IGNORE so existing passwords are never overwritten
    doctors = [
        ('nashat',   hash_pw('nashat123'),   'د. نشأت مكسيموس',  'doctor'),
        ('mohamed',  hash_pw('mohamed123'),  'د. محمد المسير',    'doctor'),
        ('affairs',  hash_pw('affairs2024'), 'شئون الطلاب',       'affairs'),
        ('admin',    hash_pw('hti@admin2026'),   'المدير',             'admin'),
    ]
    for d in doctors:
        c.execute("INSERT INTO doctors(username,password,name,role) VALUES(%s,%s,%s,%s) ON CONFLICT(username) DO NOTHING", d)

    conn.commit()
    conn.close()
    print("✅ Database initialized successfully")

# ─── Load Excel data ──────────────────────────────────────────
def load_students():
    path = os.path.join(BASE_DIR, 'data', 'students.xlsx')
    df = pd.read_excel(path)
    df.columns = [c.strip() for c in df.columns]
    # Rename to standard keys
    df = df.rename(columns={'الكود': 'id', 'الاسم ': 'name', 'الاسم': 'name', 'البرنامج': 'department'})
    df['id'] = df['id'].astype(str).str.strip()
    df['name'] = df['name'].astype(str).str.strip()
    if 'department' not in df.columns:
        df['department'] = ''
    df['department'] = df['department'].astype(str).str.strip()
    return df

def load_employees():
    path = os.path.join(BASE_DIR, 'data', 'employees.xlsx')
    df = pd.read_excel(path)
    df.columns = [c.strip() for c in df.columns]
    df = df.rename(columns={'رقم الملف': 'id', 'الاسم': 'name'})
    df['id'] = df['id'].astype(str).str.strip()
    df['name'] = df['name'].astype(str).str.strip()
    if 'department' not in df.columns:
        df['department'] = 'موظف'
    return df

# ─── Routes ───────────────────────────────────────────────────

@app.route('/')
def index():
    if 'user' not in session:
        return redirect(url_for('login_page'))
    role = session.get('role','doctor')
    return redirect(url_for('affairs_home') if role=='affairs' else url_for('doctor_search'))

@app.route('/login')
def login_page():
    if 'user' in session:
        role = session.get('role','doctor')
        return redirect(url_for('affairs_home') if role=='affairs' else url_for('doctor_search'))
    return render_template('login.html')

@app.route('/api/login', methods=['POST'])
def api_login():
    data = request.json
    username = data.get('username','').strip()
    password = data.get('password','')
    hashed   = hashlib.sha256(password.encode()).hexdigest()
    conn = get_db()
    user = db_fetchone(conn, "SELECT * FROM doctors WHERE username=%s AND password=%s", (username, hashed))
    conn.close()
    if not user:
        return jsonify({'success': False})
    session.permanent    = False  # الجلسة تنتهي عند إغلاق المتصفح، لا تبقى مسجلة الدخول دائماً
    session['user']     = user['name']
    session['username'] = user['username']
    session['role']     = user['role']
    if user['role'] == 'affairs':
        return jsonify({'success': True, 'redirect': '/affairs'})
    elif user['role'] == 'admin':
        return jsonify({'success': True, 'redirect': '/admin'})
    else:
        return jsonify({'success': True, 'redirect': '/doctor'})

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login_page'))

# ── DOCTOR PORTAL ─────────────────────────────────────────────

@app.route('/doctor')
def doctor_search():
    if 'user' not in session:
        return redirect(url_for('login_page'))
    if session.get('role') not in ('doctor', 'admin'):
        return redirect(url_for('affairs_home'))
    conn = get_db()
    today = date.today().isoformat()
    today_visits = db_fetchall(conn, "SELECT * FROM visits WHERE visit_date=%s ORDER BY created_at DESC", (today,))
    import pandas as pd, os
    try:
        df_emp = pd.read_excel(os.path.join(BASE_DIR, 'data', 'employees.xlsx'))
        emp_count = len(df_emp)
    except:
        emp_count = 142
    stats = {
        'today': len(today_visits),
        'pending': db_fetchone(conn, "SELECT COUNT(*) FROM visits WHERE submitted_to_affairs=0 AND leave_days>0")['count'],
        'medicines_count': db_fetchone(conn, "SELECT COUNT(*) FROM medicines")['count'],
        'employees_count': emp_count,
    }
    # Get current doctor name from session
    doctor_name = session.get('user', 'طبيب العيادة')
    conn.close()
    return render_template('doctor_search.html', today_visits=today_visits, stats=stats, doctor_name=doctor_name)

@app.route('/api/search')
def api_search():
    q     = request.args.get('q', '').strip()
    ptype = request.args.get('type', 'student')
    mode  = request.args.get('mode', 'id')   # 'id' or 'name'
    if not q:
        return jsonify({'found': False})
    try:
        if ptype == 'student':
            df = load_students()
            if mode == 'name':
                row = df[df['name'].str.contains(q, na=False, case=False)]
            else:
                row = df[df['id'] == q]
                if row.empty:
                    row = df[df['name'].str.contains(q, na=False, case=False)]
        else:
            df = load_employees()
            if mode == 'name':
                row = df[df['name'].str.contains(q, na=False, case=False)]
            else:
                row = df[df['id'] == q]
                if row.empty:
                    row = df[df['name'].str.contains(q, na=False, case=False)]

        if row.empty:
            return jsonify({'found': False, 'message': 'لم يتم العثور على نتيجة — تأكد من الاسم أو الكود'})

        # If multiple results (name search), return list for selection
        if len(row) > 1 and mode == 'name':
            results = []
            for _, r in row.head(10).iterrows():
                results.append({'id': str(r['id']), 'name': str(r['name']), 'department': str(r.get('department',''))})
            return jsonify({'found': True, 'multiple': True, 'results': results})

        person = row.iloc[0]
        person_id = str(person['id'])

        # Get visit history from DB
        conn = get_db()
        visits = db_fetchall(conn, "SELECT * FROM visits WHERE person_id=%s ORDER BY visit_date DESC LIMIT 20", (person_id,))
        conn.close()

        visits_list = []
        for v in visits:
            vd = dict(v)
            vd['medicines'] = json.loads(vd['medicines']) if vd['medicines'] else []
            visits_list.append(vd)

        return jsonify({
            'found': True,
            'person': {
                'id': person_id,
                'name': str(person['name']),
                'department': str(person.get('department', '')),
                'type': ptype,
            },
            'visits': visits_list
        })
    except Exception as e:
        return jsonify({'found': False, 'message': str(e)})

@app.route('/api/search_employees')
def api_search_employees():
    """بحث عن موظفين من Azure AD باستخدام User.Read.All."""
    if 'user' not in session:
        return jsonify({'error': 'غير مسجل الدخول'}), 401
    q = request.args.get('q', '').strip()
    if not q or len(q) < 2:
        return jsonify({'results': []})
    results = search_employees_in_azure(q)
    return jsonify({'results': results})


@app.route('/api/send_employee_email', methods=['POST'])
def api_send_employee_email():
    """إرسال إيميل لموظف من الدكتور."""
    if 'user' not in session:
        return jsonify({'success': False, 'error': 'غير مسجل الدخول'})
    data = request.get_json() or {}
    to_email    = data.get('to_email', '').strip()
    to_name     = data.get('to_name', '')
    patient     = data.get('patient', '')
    diagnosis   = data.get('diagnosis', '')
    medicines   = data.get('medicines', [])
    notes       = data.get('notes', '')
    doctor_name = session.get('user', 'الطبيب')

    if not to_email:
        return jsonify({'success': False, 'error': 'البريد الإلكتروني للموظف مطلوب'})

    meds_html = ''.join(f'<li>{m}</li>' for m in medicines) if medicines else '<li>لا يوجد</li>'

    html_body = f"""
    <div dir="rtl" style="font-family:Arial,sans-serif;max-width:600px;margin:auto;
                          border:1px solid #ddd;border-radius:8px;overflow:hidden;">
      <div style="background:#1a3c5e;color:white;padding:16px;text-align:center;">
        <h2 style="margin:0;">نظام عيادة المعهد التكنولوجي العالي</h2>
        <p style="margin:4px 0;font-size:13px;">clinic@hti-o.edu.eg</p>
      </div>
      <div style="padding:20px;">
        <p>السيد/ة <strong>{to_name}</strong>،</p>
        <p>يُحيطكم علماً بأن الطالب/الموظف <strong>{patient}</strong>
           قد تلقى العلاج في عيادة المعهد.</p>
        <table style="width:100%;border-collapse:collapse;margin:16px 0;">
          <tr style="background:#f0f4f8;">
            <th style="padding:8px;border:1px solid #ccc;text-align:right;">التشخيص</th>
            <td style="padding:8px;border:1px solid #ccc;">{diagnosis or 'غير محدد'}</td>
          </tr>
          <tr>
            <th style="padding:8px;border:1px solid #ccc;text-align:right;">الأدوية الموصوفة</th>
            <td style="padding:8px;border:1px solid #ccc;"><ul style="margin:0;padding-right:16px;">{meds_html}</ul></td>
          </tr>
          {'<tr style="background:#f0f4f8;"><th style="padding:8px;border:1px solid #ccc;text-align:right;">ملاحظات</th><td style="padding:8px;border:1px solid #ccc;">' + notes + '</td></tr>' if notes else ''}
        </table>
        <p style="color:#555;font-size:13px;">الطبيب المعالج: <strong>{doctor_name}</strong></p>
      </div>
      <div style="background:#f0f4f8;padding:10px;text-align:center;font-size:12px;color:#777;">
        عيادة المعهد التكنولوجي العالي بالسادس من أكتوبر
      </div>
    </div>
    """

    result = send_email_via_graph([to_email], f'تقرير طبي — {patient}', html_body)
    return jsonify(result)


@app.route('/api/medicines')
def api_medicines():
    conn = get_db()
    meds = db_fetchall(conn, "SELECT * FROM medicines ORDER BY category, id")
    conn.close()
    return jsonify([dict(m) for m in meds])

@app.route('/api/upload_certificate', methods=['POST'])
def upload_certificate():
    if 'user' not in session:
        return jsonify({'success': False, 'error': 'غير مسجل الدخول'})
    if session.get('role') not in ('doctor', 'admin'):
        return jsonify({'success': False, 'error': 'غير مصرح'})

    if 'file' not in request.files:
        return jsonify({'success': False, 'error': 'لا يوجد ملف'})

    file = request.files['file']
    student_id = request.form.get('student_id', '').strip()
    student_name = request.form.get('student_name', '').strip()

    if not student_id:
        return jsonify({'success': False, 'error': 'لا يوجد كود طالب'})
    if file.filename == '':
        return jsonify({'success': False, 'error': 'ملف غير صالح'})

    from werkzeug.utils import secure_filename
    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in ('.pdf', '.jpg', '.jpeg', '.png'):
        return jsonify({'success': False, 'error': 'نوع الملف غير مسموح — PDF, JPG, PNG فقط'})

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    safe_name_part = secure_filename(student_name)[:30] if student_name else ''
    saved_filename = f"{student_id}_{safe_name_part}_{timestamp}{ext}" if safe_name_part else f"{student_id}_{timestamp}{ext}"

    file_bytes = file.read()

    # ── Try uploading to OneDrive first ──────────────────────
    onedrive_result = upload_to_onedrive(student_id, saved_filename, file_bytes)

    if onedrive_result['success']:
        return jsonify({
            'success': True,
            'saved_name': saved_filename,
            'storage': 'onedrive',
            'web_url': onedrive_result.get('web_url', '')
        })

    # ── Fallback: save locally if OneDrive fails ─────────────
    cert_root = os.path.join(BASE_DIR, 'data', 'certificates')
    student_folder = os.path.join(cert_root, secure_filename(student_id))
    os.makedirs(student_folder, exist_ok=True)
    save_path = os.path.join(student_folder, saved_filename)
    with open(save_path, 'wb') as f:
        f.write(file_bytes)

    return jsonify({
        'success': True,
        'saved_name': saved_filename,
        'storage': 'local',
        'warning': f"تم الحفظ محلياً فقط — فشل الرفع على OneDrive: {onedrive_result.get('error','')}"
    })


def get_graph_access_token():
    """Get Microsoft Graph access token using Client Credentials (no MFA issues).
    Returns (access_token, error_message).
    """
    try:
        import requests as req
    except ImportError:
        return None, 'مكتبة requests غير مثبتة'

    CLIENT_ID     = os.environ.get('GRAPH_CLIENT_ID', '')
    CLIENT_SECRET = os.environ.get('GRAPH_CLIENT_SECRET', '')
    TENANT        = os.environ.get('GRAPH_TENANT', '')

    if not all([CLIENT_ID, CLIENT_SECRET, TENANT]):
        return None, 'متغيرات Azure غير مكتملة (GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET, GRAPH_TENANT)'

    try:
        token_url = f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token"
        token_resp = req.post(token_url, data={
            'client_id':     CLIENT_ID,
            'client_secret': CLIENT_SECRET,
            'scope':         'https://graph.microsoft.com/.default',
            'grant_type':    'client_credentials',
        }, timeout=15)

        if token_resp.status_code != 200:
            try:
                err = token_resp.json().get('error_description', token_resp.text[:300])
            except Exception:
                err = token_resp.text[:300]
            return None, f'فشل الحصول على توكن Azure: {err}'

        access_token = token_resp.json().get('access_token')
        if not access_token:
            return None, 'لم يتم الحصول على access token'

        return access_token, None

    except Exception as e:
        return None, str(e)


def search_employees_in_azure(query):
    """Search employees by name in Azure AD using User.Read.All permission."""
    try:
        import requests as req
    except ImportError:
        return []

    access_token, err = get_graph_access_token()
    if not access_token:
        return []

    try:
        resp = req.get(
            f"https://graph.microsoft.com/v1.0/users"
            f"?$filter=startswith(displayName,'{query}') or startswith(givenName,'{query}')"
            f"&$select=displayName,mail,jobTitle,department"
            f"&$top=10",
            headers={'Authorization': f'Bearer {access_token}'},
            timeout=10
        )
        if resp.status_code == 200:
            users = resp.json().get('value', [])
            return [
                {
                    'name':       u.get('displayName', ''),
                    'email':      u.get('mail', ''),
                    'title':      u.get('jobTitle', ''),
                    'department': u.get('department', ''),
                }
                for u in users if u.get('mail')
            ]
    except Exception:
        pass
    return []


def send_email_via_graph(to_emails, subject, html_body):
    """Send email using Microsoft Graph API (Client Credentials — no MFA)."""
    try:
        import requests as req
    except ImportError:
        return {'success': False, 'error': 'مكتبة requests غير مثبتة'}

    access_token, err = get_graph_access_token()
    if not access_token:
        return {'success': False, 'error': err}

    CLINIC_EMAIL = os.environ.get('CLINIC_EMAIL', '') or os.environ.get('ONEDRIVE_EMAIL', '')

    message = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html_body},
            "toRecipients": [{"emailAddress": {"address": e.strip()}} for e in to_emails if e.strip()],
        },
        "saveToSentItems": "true"
    }

    try:
        resp = req.post(
            f"https://graph.microsoft.com/v1.0/users/{CLINIC_EMAIL}/sendMail",
            headers={
                'Authorization': f'Bearer {access_token}',
                'Content-Type':  'application/json',
            },
            json=message,
            timeout=20
        )
        if resp.status_code == 202:
            return {'success': True, 'sent_to': to_emails}
        else:
            err_detail = resp.text[:300]
            try:
                err_detail = resp.json().get('error', {}).get('message', err_detail)
            except Exception:
                pass
            return {'success': False, 'error': f'فشل الإرسال ({resp.status_code}): {err_detail}'}
    except Exception as e:
        return {'success': False, 'error': str(e)}


def upload_to_onedrive(student_id, filename, file_bytes):
    """رفع ملف على OneDrive باستخدام Client Credentials (بدون MFA)."""
    try:
        import requests as req
    except ImportError:
        return {'success': False, 'error': 'مكتبة requests غير مثبتة'}

    ONEDRIVE_USER = os.environ.get('ONEDRIVE_EMAIL', '')
    access_token, err = get_graph_access_token()
    if not access_token:
        return {'success': False, 'error': err}

    try:
        upload_url = (
            f"https://graph.microsoft.com/v1.0"
            f"/users/{ONEDRIVE_USER}/drive/root"
            f":/العياده/{student_id}/{filename}:/content"
        )
        upload_resp = req.put(
            upload_url,
            headers={
                'Authorization': f'Bearer {access_token}',
                'Content-Type':  'application/octet-stream',
            },
            data=file_bytes,
            timeout=30
        )
        if upload_resp.status_code in (200, 201):
            return {'success': True, 'web_url': upload_resp.json().get('webUrl', '')}
        else:
            return {'success': False,
                    'error': f'فشل الرفع: {upload_resp.status_code} — {upload_resp.text[:200]}'}
    except Exception as e:
        return {'success': False, 'error': str(e)}


def _UNUSED_old_upload(student_id, filename, file_bytes):
    """[deprecated] kept for reference"""
    import base64
    FOLDER_URL = os.environ.get('ONEDRIVE_FOLDER_URL', '')

    if not FOLDER_URL:
        return {'success': False, 'error': 'رابط مجلد OneDrive غير مُعد (ONEDRIVE_FOLDER_URL في .env)'}

    try:
        # تحويل رابط المشاركة لـ shareId يفهمه Microsoft Graph
        clean_url = FOLDER_URL.split('?')[0]
        b64 = base64.urlsafe_b64encode(clean_url.encode('utf-8')).decode('utf-8').rstrip('=')
        share_id = f"u!{b64}"

        # 1) معرفة بيانات المجلد (driveId + itemId) من رابط المشاركة نفسه
        #    بدون أي توكن — الطلب ده بيتعامل كـ "زائر" بصلاحية الرابط
        folder_resp = req.get(
            f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem",
            timeout=15
        )

        if folder_resp.status_code != 200:
            return {
                'success': False,
                'error': (
                    'فشل الوصول لمجلد OneDrive عبر الرابط. تأكدي إن صلاحية '
                    'الرابط "Anyone with the link can edit" مفعّلة. '
                    f'({folder_resp.status_code})'
                )
            }

        fd = folder_resp.json()
        drive_id  = fd.get('parentReference', {}).get('driveId', '') or fd.get('id', '')
        folder_id = fd.get('id', '')

        if not (drive_id and folder_id):
            return {'success': False, 'error': 'تعذّر تحديد مجلد الرفع من الرابط المُعطى'}

        # 2) رفع الملف داخل مجلد فرعي باسم كود الطالب، بنفس صلاحية الرابط
        upload_url = (
            f"https://graph.microsoft.com/v1.0"
            f"/drives/{drive_id}/items/{folder_id}"
            f":/{student_id}/{filename}:/content"
        )

        upload_resp = req.put(
            upload_url,
            headers={'Content-Type': 'application/octet-stream'},
            data=file_bytes,
            timeout=30
        )

        if upload_resp.status_code in (200, 201):
            return {'success': True, 'web_url': upload_resp.json().get('webUrl', '')}
        else:
            return {'success': False, 'error': f'فشل الرفع: {upload_resp.status_code} — {upload_resp.text[:250]}'}

    except Exception as e:
        return {'success': False, 'error': str(e)}


@app.route('/api/save_visit', methods=['POST'])
def api_save_visit():
    data = request.json
    conn = get_db()
    try:
        # Get doctor name: session is most reliable, payload is fallback
        doctor_name = session.get('user') or data.get('doctor_name') or 'غير محدد'
        cur = conn.cursor()
        cur.execute('''INSERT INTO visits
            (person_type, person_id, person_name, department,
             visit_date, diagnosis, treatment_type, medicines,
             prescription, leave_days, leave_from, leave_to,
             is_chronic, report_reason, notes, certificate_name,
             submitted_to_affairs, doctor_name)
            VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            RETURNING id''', (
            data.get('person_type'),
            data.get('person_id'),
            data.get('person_name'),
            data.get('department'),
            data.get('visit_date', date.today().isoformat()),
            data.get('diagnosis'),
            data.get('treatment_type'),
            json.dumps(data.get('medicines', []), ensure_ascii=False),
            data.get('prescription'),
            int(data.get('leave_days', 0)),
            data.get('leave_from'),
            data.get('leave_to'),
            1 if data.get('is_chronic') else 0,
            data.get('report_reason'),
            data.get('notes'),
            data.get('certificate_name'),
            0,
            doctor_name
        ))
        visit_id = cur.fetchone()['id']

        # Deduct medicine stock
        for med_id in data.get('medicine_ids', []):
            db_execute(conn, "UPDATE medicines SET stock=stock-1 WHERE id=%s", (med_id,))

        conn.commit()
        conn.close()
        return jsonify({'success': True, 'visit_id': visit_id})
    except Exception as e:
        conn.close()
        return jsonify({'success': False, 'error': str(e)})

@app.route('/api/submit_to_affairs/<int:visit_id>', methods=['POST'])
def api_submit_to_affairs(visit_id):
    conn = get_db()
    db_execute(conn, "UPDATE visits SET submitted_to_affairs=1 WHERE id=%s", (visit_id,))
    conn.commit()
    conn.close()
    return jsonify({'success': True})

# ── AFFAIRS PORTAL ────────────────────────────────────────────

@app.route('/affairs')
def affairs_home():
    if 'user' not in session:
        return redirect(url_for('login_page'))
    if session.get('role') not in ('affairs', 'admin'):
        return redirect(url_for('doctor_search'))
    conn = get_db()
    reports = db_fetchall(conn, '''
        SELECT v.*, ar.id as ar_id, ar.printed
        FROM visits v
        LEFT JOIN affairs_reports ar ON ar.visit_id = v.id
        WHERE v.submitted_to_affairs=1 AND v.person_type='student'
        ORDER BY v.created_at DESC
    ''')
    new_count = db_fetchone(conn, "SELECT COUNT(*) FROM visits v LEFT JOIN affairs_reports ar ON ar.visit_id=v.id WHERE v.submitted_to_affairs=1 AND v.person_type='student' AND ar.id IS NULL")['count']
    reports_list = [dict(r) for r in reports]

    # لكل طالب في القائمة، اجمعي كل تاريخه (كل التقارير السابقة + الحالية)
    # عشان شئون تقدر تشوف كل المرات اللي عمل فيها تقرير قبل كده، بتاريخها ومدتها وتشخيصها
    student_ids = list({r['person_id'] for r in reports_list})
    history_by_student = {}
    for sid in student_ids:
        hist_rows = db_fetchall(conn, '''
            SELECT v.id as visit_id, v.visit_date, v.diagnosis, v.leave_days,
                   v.leave_from, v.leave_to, v.doctor_name, v.report_reason,
                   v.is_chronic, ar.report_number, ar.id as ar_id
            FROM visits v
            LEFT JOIN affairs_reports ar ON ar.visit_id = v.id
            WHERE v.person_id=%s AND v.person_type='student' AND v.submitted_to_affairs=1
            ORDER BY v.visit_date DESC
        ''', (sid,)).fetchall()
        history_by_student[sid] = [dict(h) for h in hist_rows]

    conn.close()
    return render_template('affairs.html', reports=reports_list, new_count=new_count,
                            history_by_student=history_by_student)

@app.route('/api/create_affairs_report', methods=['POST'])
def api_create_affairs_report():
    data = request.json
    conn = get_db()
    visit_id = data.get('visit_id')
    # Check if report already exists for this visit
    existing = db_fetchone(conn, "SELECT id, report_number FROM affairs_reports WHERE visit_id=?", (visit_id,))
    if existing:
        conn.close()
        return jsonify({'success': True, 'report_number': existing['report_number'], 'ar_id': existing['id'], 'duplicate': True})
    # Auto generate report number
    count = db_fetchone(conn, "SELECT COUNT(*) FROM affairs_reports")['count']
    report_number = f"RPT-{datetime.now().year}-{str(count+1).zfill(3)}"
    try:
        cur2 = conn.cursor()
        cur2.execute('''INSERT INTO affairs_reports
            (visit_id, report_number, action_type, leave_from, leave_to,
             leave_days, affairs_notes, printed)
            VALUES(%s,%s,%s,%s,%s,%s,%s,0)''', (
            visit_id,
            report_number,
            data.get('action_type', 'رفع غياب — إجازة مرضية'),
            data.get('leave_from'),
            data.get('leave_to'),
            int(data.get('leave_days', 0)),
            data.get('affairs_notes', ''),
        ))
        conn.commit()
        ar_id = cur2.lastrowid
        conn.close()
        return jsonify({'success': True, 'report_number': report_number, 'ar_id': ar_id, 'duplicate': False})
    except Exception as e:
        conn.close()
        return jsonify({'success': False, 'error': str(e)})

@app.route('/api/print_report/<int:ar_id>', methods=['POST'])
def api_print_report(ar_id):
    conn = get_db()
    db_execute(conn, "UPDATE affairs_reports SET printed=1 WHERE id=%s", (ar_id,))
    conn.commit()
    conn.close()
    return jsonify({'success': True})

@app.route('/api/stats')
def api_stats():
    conn = get_db()
    today = date.today().isoformat()
    stats = {
        'today_visits': db_fetchone(conn, "SELECT COUNT(*) FROM visits WHERE visit_date=?", (today,))['count'],
        'pending_submit': db_fetchone(conn, "SELECT COUNT(*) FROM visits WHERE submitted_to_affairs=0 AND leave_days>0")['count'],
        'affairs_new': db_fetchone(conn, "SELECT COUNT(*) FROM visits v LEFT JOIN affairs_reports ar ON ar.visit_id=v.id WHERE v.submitted_to_affairs=1 AND ar.id IS NULL")['count'],
        'total_visits': db_fetchone(conn, "SELECT COUNT(*) FROM visits")['count'],
    }
    conn.close()
    return jsonify(stats)



@app.route('/affairs/archive')
def affairs_archive():
    if 'user' not in session:
        return redirect(url_for('login_page'))
    if session.get('role') not in ('affairs', 'admin'):
        return redirect(url_for('doctor_search'))
    conn = get_db()
    # Show all completed reports (students only)
    reports = db_fetchone(conn, '''
        SELECT v.*, ar.id as ar_id, ar.report_number, ar.printed, ar.created_at as ar_date
        FROM visits v
        JOIN affairs_reports ar ON ar.visit_id = v.id
        WHERE v.person_type='student'
        ORDER BY ar.created_at DESC
    ''')
    conn.close()
    return render_template('affairs_archive.html', reports=reports)


# ── PRINT ROUTES ──────────────────────────────────────────────

@app.route('/print/report/<int:visit_id>')
def print_report(visit_id):
    if 'user' not in session:
        return redirect(url_for('login_page'))
    conn = get_db()
    visit = db_execute(conn, "SELECT * FROM visits WHERE id=%s", (visit_id,)).fetchone()
    conn.close()
    if not visit:
        return "زيارة غير موجودة", 404
    return render_template('print_report.html', visit=dict(visit))

@app.route('/print/affairs/<int:ar_id>')
def print_affairs_report(ar_id):
    if 'user' not in session:
        return redirect(url_for('login_page'))
    conn = get_db()
    report = db_execute(conn, "SELECT * FROM affairs_reports WHERE id=%s", (ar_id,)).fetchone()
    if not report:
        conn.close()
        return "تقرير غير موجود", 404
    # Get all visits linked to this report (usually one, but can be batch)
    visit = db_execute(conn, "SELECT * FROM visits WHERE id=%s", (dict(report)['visit_id'],)).fetchone()
    conn.close()
    visits = [dict(visit)] if visit else []
    return render_template('print_affairs.html', report=dict(report), visits=visits)


@app.route('/affairs/archive/export')
def affairs_archive_export():
    if 'user' not in session:
        return redirect(url_for('login_page'))
    if session.get('role') not in ('affairs', 'admin'):
        return redirect(url_for('doctor_search'))
    import io
    conn = get_db()
    reports = db_fetchall(conn, '''
        SELECT v.person_id, v.person_name, v.department, v.diagnosis,
               v.leave_days, v.leave_from, v.leave_to, v.doctor_name,
               v.visit_date, ar.report_number, ar.action_type, ar.created_at
        FROM visits v
        JOIN affairs_reports ar ON ar.visit_id = v.id
        WHERE v.person_type='student'
        ORDER BY ar.created_at DESC
    ''')
    conn.close()

    import pandas as pd
    data = []
    for r in reports:
        data.append({
            'رقم التقرير':    r['report_number'],
            'كود الطالب':     r['person_id'],
            'اسم الطالب':     r['person_name'],
            'القسم':          r['department'] or '—',
            'التشخيص':        r['diagnosis'] or '—',
            'الطبيب المعالج': r['doctor_name'] or '—',
            'تاريخ الزيارة':  r['visit_date'],
            'بداية الإجازة':  r['leave_from'] or '—',
            'نهاية الإجازة':  r['leave_to'] or '—',
            'عدد الأيام':     r['leave_days'] or 0,
            'نوع الإجراء':    r['action_type'] or '—',
            'تاريخ التقرير':  r['created_at'][:10] if r['created_at'] else '—',
        })

    df = pd.DataFrame(data)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='أرشيف التقارير')
        # Style the header
        ws = writer.sheets['أرشيف التقارير']
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 18
    output.seek(0)

    from flask import send_file
    return send_file(
        output,
        download_name=f'archive_reports_{datetime.now().strftime("%Y%m%d")}.xlsx',
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


# ── EMAIL ─────────────────────────────────────────────────────

@app.route('/api/send_report_email', methods=['POST'])
def send_report_email():
    if 'user' not in session:
        return jsonify({'success': False, 'error': 'غير مسجل الدخول'})
    if session.get('role') not in ('affairs', 'admin'):
        return jsonify({'success': False, 'error': 'غير مصرح'})

    data = request.json
    ar_id       = data.get('ar_id')
    to_emails   = data.get('to_emails', [])   # list of emails
    extra_note  = data.get('note', '')

    if not ar_id or not to_emails:
        return jsonify({'success': False, 'error': 'بيانات ناقصة'})

    # Validate all emails are @hti-o.edu.eg
    for email in to_emails:
        email = email.strip()
        if not email.endswith('@hti-o.edu.eg') and '@' in email:
            # Allow any email for flexibility — just warn
            pass

    conn = get_db()
    report = db_execute(conn, '''
        SELECT ar.*, v.person_name, v.person_id, v.department,
               v.diagnosis, v.leave_days, v.leave_from, v.leave_to,
               v.doctor_name, v.visit_date, v.report_reason
        FROM affairs_reports ar
        JOIN visits v ON v.id = ar.visit_id
        WHERE ar.id=%s
    ''', (ar_id,))
    conn.close()

    if not report:
        return jsonify({'success': False, 'error': 'التقرير غير موجود'})

    r = dict(report)

    SENDER_PASS_FROM_REQUEST = data.get('password', '')
    # If the person typed a password in the modal but .env already has one,
    # prefer .env (Graph flow needs ONEDRIVE_PASSWORD set there anyway).
    if SENDER_PASS_FROM_REQUEST and not os.environ.get('ONEDRIVE_PASSWORD'):
        os.environ['ONEDRIVE_PASSWORD'] = SENDER_PASS_FROM_REQUEST

    # Build HTML email body
    html_body = f"""
    <html dir="rtl">
    <body style="font-family:Arial,sans-serif;direction:rtl;padding:20px;color:#1A2744;">
      <div style="max-width:600px;margin:0 auto;border:1px solid #D8E2EF;border-radius:12px;overflow:hidden;">
        <div style="background:#1B3A6B;padding:20px;text-align:center;">
          <h2 style="color:#fff;margin:0;">المعهد التكنولوجي العالي</h2>
          <p style="color:rgba(255,255,255,.7);margin:4px 0 0;">إدارة شئون الطلاب — السادس من أكتوبر</p>
        </div>
        <div style="padding:24px;">
          <h3 style="color:#1B3A6B;border-bottom:2px solid #D8E2EF;padding-bottom:10px;">
            تقرير إجازة مرضية — {r['report_number']}
          </h3>
          <table style="width:100%;border-collapse:collapse;font-size:14px;">
            <tr style="background:#F0F4F8;"><td style="padding:9px 12px;font-weight:bold;width:140px;">اسم الطالب</td><td style="padding:9px 12px;">{r['person_name']}</td></tr>
            <tr><td style="padding:9px 12px;font-weight:bold;">كود الطالب</td><td style="padding:9px 12px;color:#059669;font-weight:700;">{r['person_id']}</td></tr>
            <tr style="background:#F0F4F8;"><td style="padding:9px 12px;font-weight:bold;">القسم</td><td style="padding:9px 12px;">{r['department'] or '—'}</td></tr>
            <tr><td style="padding:9px 12px;font-weight:bold;">التشخيص</td><td style="padding:9px 12px;">{r['diagnosis'] or '—'}</td></tr>
            <tr style="background:#F0F4F8;"><td style="padding:9px 12px;font-weight:bold;">الطبيب المعالج</td><td style="padding:9px 12px;">{r['doctor_name'] or '—'}</td></tr>
            <tr><td style="padding:9px 12px;font-weight:bold;">مدة الإجازة</td><td style="padding:9px 12px;color:#D97706;font-weight:700;">{r['leave_days']} أيام</td></tr>
            <tr style="background:#F0F4F8;"><td style="padding:9px 12px;font-weight:bold;">من تاريخ</td><td style="padding:9px 12px;">{r['leave_from'] or '—'}</td></tr>
            <tr><td style="padding:9px 12px;font-weight:bold;">إلى تاريخ</td><td style="padding:9px 12px;">{r['leave_to'] or '—'}</td></tr>
            <tr style="background:#F0F4F8;"><td style="padding:9px 12px;font-weight:bold;">سبب التقرير</td><td style="padding:9px 12px;">{r['report_reason'] or '—'}</td></tr>
          </table>
          {f'<div style="background:#FEF3C7;border:1px solid #FDE68A;border-radius:8px;padding:12px;margin-top:16px;"><strong>ملاحظة:</strong> {extra_note}</div>' if extra_note else ''}
          <p style="margin-top:20px;font-size:13px;color:#4A5878;">
            تحريراً في: {r['created_at'][:10] if r['created_at'] else ''}<br>
            <strong>إدارة شئون الطلاب — داليا شبل</strong>
          </p>
        </div>
        <div style="background:#F0F4F8;padding:12px;text-align:center;font-size:12px;color:#8A95B0;">
          المعهد التكنولوجي العالي — السادس من أكتوبر | studens_affairs@hti-o.edu.eg
        </div>
      </div>
    </body>
    </html>
    """

    subject = f"تقرير إجازة مرضية — {r['person_name']} — {r['report_number']}"
    result = send_email_via_graph(to_emails, subject, html_body)

    # سجّل المحاولة (نجحت أو فشلت) في سجل الإيميلات لعرضها في صفحة الإدارة
    log_conn = get_db()
    log_db_execute(conn, '''INSERT INTO email_log
        (ar_id, visit_id, person_name, person_id, report_number,
         to_emails, sent_by, status, error_message)
        VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s)''', (
        ar_id,
        r.get('visit_id'),
        r['person_name'],
        r['person_id'],
        r['report_number'],
        json.dumps(to_emails, ensure_ascii=False),
        session.get('user', 'غير معروف'),
        'success' if result.get('success') else 'failed',
        None if result.get('success') else result.get('error', 'خطأ غير معروف')
    ))
    log_conn.commit()
    log_conn.close()

    if result.get('success'):
        return jsonify({'success': True, 'sent_to': to_emails})
    else:
        return jsonify({'success': False, 'error': result.get('error', 'خطأ غير معروف في الإرسال')})

# ── ADMIN PANEL ───────────────────────────────────────────────

@app.route('/admin')
def admin_home():
    if 'user' not in session:
        return redirect(url_for('login_page'))
    if session.get('role') != 'admin':
        if session.get('role') == 'affairs':
            return redirect(url_for('affairs_home'))
        return redirect(url_for('doctor_search'))
    conn = get_db()
    visits = db_fetchall(conn, "SELECT * FROM visits ORDER BY created_at DESC")
    email_logs = db_fetchall(conn, "SELECT * FROM email_log ORDER BY sent_at DESC LIMIT 100")
    stats = {
        'total_visits':   db_fetchone(conn, "SELECT COUNT(*) FROM visits")['count'],
        'submitted':      db_fetchone(conn, "SELECT COUNT(*) FROM visits WHERE submitted_to_affairs=1")['count'],
        'affairs_reports':db_fetchone(conn, "SELECT COUNT(*) FROM affairs_reports")['count'],
        'medicines':      db_fetchone(conn, "SELECT COUNT(*) FROM medicines")['count'],
        'emails_sent':    db_fetchone(conn, "SELECT COUNT(*) FROM email_log WHERE status='success'")['count'],
    }
    conn.close()
    return render_template('admin.html', visits=visits, stats=stats, email_logs=email_logs)

@app.route('/admin/delete_visit/<int:visit_id>', methods=['DELETE'])
def admin_delete_visit(visit_id):
    if session.get('role') != 'admin': return jsonify({'success':False,'error':'غير مصرح'}), 403
    conn = get_db()
    try:
        # Delete related affairs reports first
        db_execute(conn, "DELETE FROM affairs_reports WHERE visit_id=?", (visit_id,))
        db_execute(conn, "DELETE FROM visits WHERE id=%s", (visit_id,))
        conn.commit()
        conn.close()
        return jsonify({'success': True})
    except Exception as e:
        conn.close()
        return jsonify({'success': False, 'error': str(e)})

@app.route('/admin/delete_visits_bulk', methods=['DELETE'])
def admin_delete_visits_bulk():
    if session.get('role') != 'admin': return jsonify({'success':False,'error':'غير مصرح'}), 403
    data = request.json
    ids = data.get('ids', [])
    if not ids:
        return jsonify({'success': False, 'error': 'لا توجد زيارات محددة'})
    conn = get_db()
    try:
        placeholders = ','.join('?' * len(ids))
        db_execute(conn, f"DELETE FROM affairs_reports WHERE visit_id IN ({placeholders})", ids)
        db_execute(conn, f"DELETE FROM visits WHERE id IN ({placeholders})", ids)
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'deleted': len(ids)})
    except Exception as e:
        conn.close()
        return jsonify({'success': False, 'error': str(e)})

@app.route('/admin/reset', methods=['GET'])
def admin_reset():
    conn = get_db()
    db_execute(conn, "DELETE FROM affairs_reports")
    db_execute(conn, "DELETE FROM visits")
    conn.commit()
    conn.close()
    return redirect(url_for('admin_home'))

@app.route('/doctor/visit')
def doctor_visit():
    if 'user' not in session:
        return redirect(url_for('login_page'))
    if session.get('role') not in ('doctor', 'admin'):
        return redirect(url_for('affairs_home'))
    doctor_name = session.get('user', 'طبيب العيادة')
    return render_template('visit.html', doctor_name=doctor_name)

if __name__ == '__main__':
    init_db()
    port = int(os.environ.get('PORT', 5000))
    print("\n" + "="*55)
    print("  🏥  نظام العيادة الجامعية — Clinic System")
    print("="*55)
    print(f"  🔗  تسجيل الدخول:        http://localhost:{port}/login")
    print(f"  🔗  العياده (الدكتور):   http://localhost:{port}/doctor")
    print(f"  🔗  شئون الطلاب:         http://localhost:{port}/affairs")
    print("="*55 + "\n")
    app.run(host='0.0.0.0', port=port, debug=False)
