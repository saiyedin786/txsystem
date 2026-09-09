import os
import re
import sqlite3
from io import BytesIO
from datetime import datetime
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, flash, send_file, jsonify, session
from werkzeug.security import generate_password_hash, check_password_hash
import pandas as pd
from init_db import DB_PATH, EXCEL_PATH, init_db

app = Flask(__name__)
app.secret_key = 'bts_database_secret_key_antigravity'

app.add_template_global(max, 'max')
app.add_template_global(min, 'min')


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            flash('Please log in using your Username/Staff No and password to access the system.', 'warning')
            return redirect(url_for('login', next=request.url))
        return f(*args, **kwargs)
    return decorated_function


def row_to_dict(row):
    if row is None:
        return None
    d = dict(row)
    d['Location'] = d.get('location', '')
    d['cpan/maan/vsat'] = d.get('cpan_maan_vsat', '')
    d['tx-system-ip'] = d.get('tx_system_ip', '')
    d['tx-system-location'] = d.get('tx_system_location', '')
    d['tx-system-port'] = d.get('tx_system_port', '')
    return d


def derive_ssa_and_location(site_name, existing_ssa="", existing_loc=""):
    if not site_name:
        return existing_ssa, existing_loc
    
    site_str = str(site_name).strip()
    if '_' in site_str:
        parts = site_str.split('_', 1)
        ssa_part = parts[0].strip()
        loc_part = parts[1].strip()
    else:
        ssa_part = site_str
        loc_part = ""
    
    m_full = re.match(r'^([A-Za-z0-9]*?\d+)(\s*[a-zA-Z].*)', ssa_part)
    if m_full:
        clean_ssa = m_full.group(1).strip()
        extra_loc = m_full.group(2).strip()
        if loc_part:
            loc_part = extra_loc + "_" + loc_part
        else:
            loc_part = extra_loc
        ssa_part = clean_ssa
        
    final_ssa = ssa_part if ssa_part else existing_ssa
    final_loc = loc_part if loc_part else existing_loc
    return final_ssa, final_loc


def transform_single_port(p_str):
    p_str = p_str.strip()
    if not p_str:
        return ""
        
    if "/" in p_str:
        tokens = [t.strip() for t in p_str.split("/") if t.strip()]
        if not tokens:
            return ""
            
        first = tokens[0].lower()
        
        if len(tokens) >= 4:
            fourth = tokens[3].lower()
            return f"{first}/{fourth}"
        elif len(tokens) == 2:
            second = tokens[1].lower()
            if second.startswith('j'):
                second = 'p' + second[1:]
            return f"{first}/{second}"
        elif len(tokens) == 3:
            third = tokens[2].lower()
            if third.startswith('j'):
                third = 'p' + third[1:]
            return f"{first}/{third}"
        else:
            return p_str.lower()
    else:
        return p_str.strip().lower()


def transform_tx_port(port_str):
    if not port_str:
        return ""
    
    port_str = str(port_str).strip()
    if not port_str or port_str.lower() in ('nan', 'none', '-'):
        return ""
        
    delimiter = ";" if ";" in port_str else ("," if "," in port_str else None)
    
    if delimiter:
        parts = port_str.split(delimiter)
        transformed_parts = [transform_single_port(p) for p in parts]
        return (delimiter + " ").join(p for p in transformed_parts if p)
    else:
        return transform_single_port(port_str)


VALID_SEARCH_COLUMNS = {
    'all': 'Generic (All Columns)',
    'site_id': 'Site ID',
    'site_name': 'Site Name',
    'enodeb_address': 'eNodeB IP',
    'ssa': 'SSA',
    'location': 'Location',
    'cpan_maan_vsat': 'Type (CPAN/MAAN/VSAT)',
    'tx_system_ip': 'TX System IP',
    'tx_system_location': 'TX System Location',
    'tx_system_port': 'TX System Port',
    'vlan': 'VLAN'
}


def search_db(query="", search_by="all", selected_ssa="", page=1, per_page=25):
    conn = get_db()
    cursor = conn.cursor()
    
    where_clauses = []
    params = []
    
    if search_by not in VALID_SEARCH_COLUMNS:
        search_by = "all"
        
    if query:
        q_like = f"%{query.strip()}%"
        if search_by != "all":
            where_clauses.append(f"LOWER({search_by}) LIKE LOWER(?)")
            params.append(q_like)
        else:
            where_clauses.append('''
                (LOWER(site_id) LIKE LOWER(?) OR
                 LOWER(site_name) LIKE LOWER(?) OR
                 LOWER(enodeb_address) LIKE LOWER(?) OR
                 LOWER(ssa) LIKE LOWER(?) OR
                 LOWER(location) LIKE LOWER(?) OR
                 LOWER(cpan_maan_vsat) LIKE LOWER(?) OR
                 LOWER(tx_system_ip) LIKE LOWER(?) OR
                 LOWER(tx_system_location) LIKE LOWER(?) OR
                 LOWER(tx_system_port) LIKE LOWER(?) OR
                 LOWER(vlan) LIKE LOWER(?))
            ''')
            params.extend([q_like] * 10)
        
    if selected_ssa:
        where_clauses.append("LOWER(ssa) = LOWER(?)")
        params.append(selected_ssa.strip())
        
    where_sql = ""
    if where_clauses:
        where_sql = "WHERE " + " AND ".join(where_clauses)
        
    cursor.execute("SELECT COUNT(*) FROM bts_sites;")
    total_records = cursor.fetchone()[0]
    
    cursor.execute("SELECT COUNT(*) FROM bts_sites WHERE LOWER(cpan_maan_vsat) LIKE '%cpan%';")
    cpan_total_count = cursor.fetchone()[0]
    
    cursor.execute("SELECT COUNT(*) FROM bts_sites WHERE LOWER(cpan_maan_vsat) LIKE '%maan%';")
    maan_total_count = cursor.fetchone()[0]
    
    count_sql = f"SELECT COUNT(*) FROM bts_sites {where_sql};"
    cursor.execute(count_sql, params)
    filtered_count = cursor.fetchone()[0]
    
    cpan_where = list(where_clauses) + ["LOWER(cpan_maan_vsat) LIKE '%cpan%'"]
    cpan_where_sql = "WHERE " + " AND ".join(cpan_where)
    cursor.execute(f"SELECT COUNT(*) FROM bts_sites {cpan_where_sql};", params)
    cpan_filtered_count = cursor.fetchone()[0]
    
    maan_where = list(where_clauses) + ["LOWER(cpan_maan_vsat) LIKE '%maan%'"]
    maan_where_sql = "WHERE " + " AND ".join(maan_where)
    cursor.execute(f"SELECT COUNT(*) FROM bts_sites {maan_where_sql};", params)
    maan_filtered_count = cursor.fetchone()[0]
    
    total_pages = max(1, (filtered_count + per_page - 1) // per_page)
    page = max(1, min(page, total_pages))
    offset = (page - 1) * per_page
    
    data_sql = f"SELECT * FROM bts_sites {where_sql} ORDER BY id ASC LIMIT ? OFFSET ?;"
    cursor.execute(data_sql, params + [per_page, offset])
    rows = cursor.fetchall()
    
    cursor.execute("SELECT DISTINCT ssa FROM bts_sites WHERE ssa IS NOT NULL AND ssa != '' ORDER BY ssa ASC;")
    unique_ssas = [r['ssa'] for r in cursor.fetchall()]
    
    conn.close()
    
    records = [row_to_dict(r) for r in rows]
    return {
        'records': records,
        'query': query,
        'search_by': search_by,
        'selected_ssa': selected_ssa,
        'unique_ssas': unique_ssas,
        'page': page,
        'per_page': per_page,
        'total_pages': total_pages,
        'total_records': total_records,
        'filtered_count': filtered_count,
        'cpan_total_count': cpan_total_count,
        'maan_total_count': maan_total_count,
        'cpan_filtered_count': cpan_filtered_count,
        'maan_filtered_count': maan_filtered_count
    }


# ==========================================
# AUTHENTICATION & USER REGISTRATION ROUTES
# ==========================================

@app.route('/login', methods=['GET', 'POST'])
def login():
    if session.get('user_id'):
        return redirect(url_for('index'))
        
    if request.method == 'POST':
        user_input = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()
        
        if not user_input or not password:
            flash('Please enter both Username/Staff Number and Password.', 'danger')
            return render_template('login.html')
            
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM users WHERE LOWER(username) = LOWER(?) OR LOWER(staff_no) = LOWER(?);", (user_input, user_input))
        user = cursor.fetchone()
        conn.close()
        
        if user and check_password_hash(user['password_hash'], password):
            session['user_id'] = user['id']
            session['username'] = user['username'] or user['staff_no']
            session['staff_no'] = user['staff_no']
            session['full_name'] = user['full_name']
            session['role'] = user['role']
            
            flash(f"Welcome back, {user['full_name']}!", "success")
            next_url = request.args.get('next')
            if next_url and next_url.startswith('/'):
                return redirect(next_url)
            return redirect(url_for('index'))
        else:
            flash('Invalid Username/Staff Number or Password. Please try again.', 'danger')
            
    return render_template('login.html')


@app.route('/register', methods=['GET', 'POST'])
def register():
    if session.get('user_id'):
        return redirect(url_for('index'))
        
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        staff_no = request.form.get('staff_no', '').strip()
        full_name = request.form.get('full_name', '').strip()
        email = request.form.get('email', '').strip()
        department = request.form.get('department', '').strip()
        password = request.form.get('password', '').strip()
        confirm_password = request.form.get('confirm_password', '').strip()
        
        if not username or not staff_no or not full_name or not password:
            flash('Username, Staff Number, Full Name, and Password are required.', 'danger')
            return render_template('register.html', form_data=request.form)
            
        if password != confirm_password:
            flash('Password and Confirm Password do not match.', 'danger')
            return render_template('register.html', form_data=request.form)
            
        conn = get_db()
        cursor = conn.cursor()
        
        cursor.execute("SELECT id FROM users WHERE LOWER(username) = LOWER(?) OR LOWER(staff_no) = LOWER(?);", (username, staff_no))
        if cursor.fetchone() is not None:
            conn.close()
            flash(f'Username "{username}" or Staff Number "{staff_no}" is already registered.', 'danger')
            return render_template('register.html', form_data=request.form)
            
        pass_hash = generate_password_hash(password)
        cursor.execute('''
            INSERT INTO users (username, staff_no, full_name, email, department, password_hash, role)
            VALUES (?, ?, ?, ?, ?, ?, ?);
        ''', (username, staff_no, full_name, email, department, pass_hash, 'staff'))
        
        conn.commit()
        conn.close()
        
        flash(f'Account for "{username}" created successfully! Please log in with your credentials.', 'success')
        return redirect(url_for('login'))
        
    return render_template('register.html', form_data={})


@app.route('/logout')
def logout():
    session.clear()
    flash('You have been logged out successfully.', 'info')
    return redirect(url_for('login'))


# ==========================================
# USER PROFILE ROUTES
# ==========================================

@app.route('/profile')
@login_required
def profile():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM users WHERE id = ?;", (session['user_id'],))
    user_row = cursor.fetchone()
    conn.close()
    
    if not user_row:
        session.clear()
        flash('User account not found.', 'danger')
        return redirect(url_for('login'))
        
    return render_template('profile.html', user=dict(user_row))


@app.route('/profile/update', methods=['POST'])
@login_required
def update_profile():
    username = request.form.get('username', '').strip()
    full_name = request.form.get('full_name', '').strip()
    email = request.form.get('email', '').strip()
    department = request.form.get('department', '').strip()
    
    if not username or not full_name:
        flash('Username and Full Name are required.', 'danger')
        return redirect(url_for('profile'))
        
    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute("SELECT id FROM users WHERE LOWER(username) = LOWER(?) AND id != ?;", (username, session['user_id']))
    if cursor.fetchone() is not None:
        conn.close()
        flash(f'Username "{username}" is already taken by another user.', 'danger')
        return redirect(url_for('profile'))
        
    cursor.execute('''
        UPDATE users SET username = ?, full_name = ?, email = ?, department = ? WHERE id = ?;
    ''', (username, full_name, email, department, session['user_id']))
    conn.commit()
    conn.close()
    
    session['username'] = username
    session['full_name'] = full_name
    flash('Your profile details have been updated successfully!', 'success')
    return redirect(url_for('profile'))


@app.route('/profile/change-password', methods=['POST'])
@login_required
def change_password():
    current_password = request.form.get('current_password', '').strip()
    new_password = request.form.get('new_password', '').strip()
    confirm_password = request.form.get('confirm_password', '').strip()
    
    if not current_password or not new_password or not confirm_password:
        flash('All password fields are required.', 'danger')
        return redirect(url_for('profile'))
        
    if new_password != confirm_password:
        flash('New Password and Confirm Password do not match.', 'danger')
        return redirect(url_for('profile'))
        
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT password_hash FROM users WHERE id = ?;", (session['user_id'],))
    row = cursor.fetchone()
    
    if not row or not check_password_hash(row['password_hash'], current_password):
        conn.close()
        flash('Incorrect Current Password. Password was not changed.', 'danger')
        return redirect(url_for('profile'))
        
    new_hash = generate_password_hash(new_password)
    cursor.execute("UPDATE users SET password_hash = ? WHERE id = ?;", (new_hash, session['user_id']))
    conn.commit()
    conn.close()
    
    flash('Your password has been updated successfully!', 'success')
    return redirect(url_for('profile'))


# ==========================================
# SITE & DATABASE ROUTES (PROTECTED)
# ==========================================

@app.route('/')
@login_required
def index():
    query = request.args.get('q', '').strip()
    search_by = request.args.get('search_by', 'all').strip()
    selected_ssa = request.args.get('ssa', '').strip()
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 25, type=int)
    
    data = search_db(query=query, search_by=search_by, selected_ssa=selected_ssa, page=page, per_page=per_page)
    return render_template('index.html', **data)


@app.route('/api/search')
@login_required
def api_search():
    query = request.args.get('q', '').strip()
    search_by = request.args.get('search_by', 'all').strip()
    selected_ssa = request.args.get('ssa', '').strip()
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 25, type=int)
    
    data = search_db(query=query, search_by=search_by, selected_ssa=selected_ssa, page=page, per_page=per_page)
    return jsonify(data)


@app.route('/site/new', methods=['GET', 'POST'])
@login_required
def create_site():
    if request.method == 'POST':
        site_id = request.form.get('site_id', '').strip()
        enodeb_address = request.form.get('enodeb_address', '').strip()
        site_name = request.form.get('site_name', '').strip()
        ssa = request.form.get('ssa', '').strip()
        location = request.form.get('Location', '').strip()
        cpan_maan = request.form.get('cpan/maan/vsat', '').strip()
        tx_ip = request.form.get('tx-system-ip', '').strip()
        tx_loc = request.form.get('tx-system-location', '').strip()
        tx_port = transform_tx_port(request.form.get('tx-system-port', '').strip())
        vlan = request.form.get('vlan', '').strip()
        
        if not site_id:
            flash('Site ID is required.', 'danger')
            return render_template('site_form.html', mode='create', form_data=request.form)
            
        conn = get_db()
        cursor = conn.cursor()
        
        cursor.execute("SELECT id FROM bts_sites WHERE LOWER(site_id) = LOWER(?);", (site_id,))
        if cursor.fetchone() is not None:
            conn.close()
            flash(f'Site ID "{site_id}" already exists. Please use a unique Site ID.', 'danger')
            return render_template('site_form.html', mode='create', form_data=request.form)
            
        if site_name and (not ssa or not location):
            derived_ssa, derived_loc = derive_ssa_and_location(site_name, ssa, location)
            if not ssa:
                ssa = derived_ssa
            if not location:
                location = derived_loc
                
        cursor.execute('''
            INSERT INTO bts_sites (
                enodeb_address, site_id, site_name, ssa, location,
                cpan_maan_vsat, tx_system_ip, tx_system_location, tx_system_port, vlan
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
        ''', (enodeb_address, site_id, site_name, ssa, location, cpan_maan, tx_ip, tx_loc, tx_port, vlan))
        
        conn.commit()
        conn.close()
        
        flash(f'Site "{site_id}" created successfully!', 'success')
        return redirect(url_for('view_site', site_id=site_id))
        
    return render_template('site_form.html', mode='create', form_data={})


@app.route('/site/<site_id>')
@login_required
def view_site(site_id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM bts_sites WHERE LOWER(site_id) = LOWER(?);", (site_id,))
    row = cursor.fetchone()
    conn.close()
    
    if not row:
        flash(f'Site ID "{site_id}" not found.', 'warning')
        return redirect(url_for('index'))
        
    site = row_to_dict(row)
    return render_template('view_site.html', site=site)


@app.route('/site/<site_id>/edit', methods=['GET', 'POST'])
@login_required
def edit_site(site_id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM bts_sites WHERE LOWER(site_id) = LOWER(?);", (site_id,))
    row = cursor.fetchone()
    
    if not row:
        conn.close()
        flash(f'Site ID "{site_id}" not found.', 'warning')
        return redirect(url_for('index'))
        
    site = row_to_dict(row)
    
    if request.method == 'POST':
        enodeb_address = request.form.get('enodeb_address', '').strip()
        site_name = request.form.get('site_name', '').strip()
        ssa = request.form.get('ssa', '').strip()
        location = request.form.get('Location', '').strip()
        cpan_maan = request.form.get('cpan/maan/vsat', '').strip()
        tx_ip = request.form.get('tx-system-ip', '').strip()
        tx_loc = request.form.get('tx-system-location', '').strip()
        tx_port = transform_tx_port(request.form.get('tx-system-port', '').strip())
        vlan = request.form.get('vlan', '').strip()
        
        if site_name and (not ssa or not location):
            derived_ssa, derived_loc = derive_ssa_and_location(site_name, ssa, location)
            if not ssa:
                ssa = derived_ssa
            if not location:
                location = derived_loc
                
        cursor.execute('''
            UPDATE bts_sites SET
                enodeb_address = ?,
                site_name = ?,
                ssa = ?,
                location = ?,
                cpan_maan_vsat = ?,
                tx_system_ip = ?,
                tx_system_location = ?,
                tx_system_port = ?,
                vlan = ?,
                updated_at = CURRENT_TIMESTAMP
            WHERE LOWER(site_id) = LOWER(?);
        ''', (enodeb_address, site_name, ssa, location, cpan_maan, tx_ip, tx_loc, tx_port, vlan, site_id))
        
        conn.commit()
        conn.close()
        
        flash(f'Site "{site_id}" updated successfully!', 'success')
        return redirect(url_for('view_site', site_id=site_id))
        
    conn.close()
    return render_template('site_form.html', mode='edit', site=site, form_data=site)


@app.route('/site/<site_id>/delete', methods=['POST'])
@login_required
def delete_site(site_id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM bts_sites WHERE LOWER(site_id) = LOWER(?);", (site_id,))
    conn.commit()
    deleted = cursor.rowcount > 0
    conn.close()
    
    if deleted:
        flash(f'Site "{site_id}" deleted successfully!', 'success')
    else:
        flash(f'Site ID "{site_id}" not found.', 'warning')
        
    return redirect(url_for('index'))


ALL_REPORT_COLUMNS = {
    'site_id': 'Site ID',
    'site_name': 'Site Name',
    'enodeb_address': 'eNodeB IP',
    'ssa': 'SSA',
    'location': 'Location',
    'cpan_maan_vsat': 'cpan/maan/vsat',
    'tx_system_ip': 'tx-system-ip',
    'tx_system_location': 'tx-system-location',
    'tx_system_port': 'tx-system-port',
    'vlan': 'VLAN'
}


@app.route('/export')
@login_required
def export_excel():
    conn = get_db()
    df = pd.read_sql_query("SELECT enodeb_address, site_id, site_name, ssa, location AS Location, cpan_maan_vsat AS 'cpan/maan/vsat', tx_system_ip AS 'tx-system-ip', tx_system_location AS 'tx-system-location', tx_system_port AS 'tx-system-port', vlan FROM bts_sites ORDER BY id ASC;", conn)
    conn.close()
    
    df.to_excel(EXCEL_PATH, index=False, sheet_name='Sheet1')
    return send_file(
        EXCEL_PATH,
        as_attachment=True,
        download_name='btsdatabase.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


@app.route('/export/custom', methods=['GET', 'POST'])
@login_required
def export_custom_report():
    if request.method == 'POST':
        export_format = request.form.get('format', 'xlsx').lower()
        scope = request.form.get('scope', 'filtered')
        query = request.form.get('q', '').strip()
        search_by = request.form.get('search_by', 'all').strip()
        selected_ssa = request.form.get('ssa', '').strip()
        selected_cols = request.form.getlist('cols')
    else:
        export_format = request.args.get('format', 'xlsx').lower()
        scope = request.args.get('scope', 'filtered')
        query = request.args.get('q', '').strip()
        search_by = request.args.get('search_by', 'all').strip()
        selected_ssa = request.args.get('ssa', '').strip()
        selected_cols = request.args.getlist('cols')
        
    if not selected_cols:
        selected_cols = list(ALL_REPORT_COLUMNS.keys())
    else:
        selected_cols = [c for c in selected_cols if c in ALL_REPORT_COLUMNS]
        if not selected_cols:
            selected_cols = list(ALL_REPORT_COLUMNS.keys())
            
    select_parts = [f"{col} AS '{ALL_REPORT_COLUMNS[col]}'" for col in selected_cols]
    select_sql = ", ".join(select_parts)
    
    conn = get_db()
    where_clauses = []
    params = []
    
    if scope == 'filtered':
        if search_by not in VALID_SEARCH_COLUMNS:
            search_by = "all"
            
        if query:
            q_like = f"%{query}%"
            if search_by != "all":
                where_clauses.append(f"LOWER({search_by}) LIKE LOWER(?)")
                params.append(q_like)
            else:
                where_clauses.append('''
                    (LOWER(site_id) LIKE LOWER(?) OR
                     LOWER(site_name) LIKE LOWER(?) OR
                     LOWER(enodeb_address) LIKE LOWER(?) OR
                     LOWER(ssa) LIKE LOWER(?) OR
                     LOWER(location) LIKE LOWER(?) OR
                     LOWER(cpan_maan_vsat) LIKE LOWER(?) OR
                     LOWER(tx_system_ip) LIKE LOWER(?) OR
                     LOWER(tx_system_location) LIKE LOWER(?) OR
                     LOWER(tx_system_port) LIKE LOWER(?) OR
                     LOWER(vlan) LIKE LOWER(?))
                ''')
                params.extend([q_like] * 10)
                
        if selected_ssa:
            where_clauses.append("LOWER(ssa) = LOWER(?)")
            params.append(selected_ssa)
            
    where_sql = ""
    if where_clauses:
        where_sql = "WHERE " + " AND ".join(where_clauses)
        
    query_sql = f"SELECT {select_sql} FROM bts_sites {where_sql} ORDER BY id ASC;"
    df = pd.read_sql_query(query_sql, conn, params=params)
    conn.close()
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    if export_format == 'csv':
        filename = f"CPAN_Report_{timestamp}.csv"
        output = BytesIO()
        csv_bytes = df.to_csv(index=False, encoding='utf-8').encode('utf-8')
        output.write(csv_bytes)
        output.seek(0)
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype='text/csv'
        )
    else:
        filename = f"CPAN_Report_{timestamp}.xlsx"
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Custom Report')
        output.seek(0)
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )


if __name__ == '__main__':
    init_db()
    print("Starting CPAN Transmission System on http://127.0.0.1:5000 ...")
    app.run(host='127.0.0.1', port=5000, debug=True)
