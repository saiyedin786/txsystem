import os
import re
import sqlite3
import ipaddress
import subprocess
from io import BytesIO
from datetime import datetime
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, flash, send_file, jsonify, session
from werkzeug.security import generate_password_hash, check_password_hash
import pandas as pd
from init_db import DB_PATH, EXCEL_PATH, NEW_EXCEL_PATH, init_db

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
    d['tx-system-ip'] = d.get('tx_system_ip', '') or d.get('endpoint_ip', '') or d.get('cpan_a_end_ip', '')
    d['tx-system-location'] = d.get('tx_system_location', '') or d.get('endpoint_node_router', '') or d.get('cpan_a_end_node', '')
    
    raw_port = d.get('endpoint_ports', '') or d.get('cpan_a_end_ports', '') or d.get('tx_system_port', '')
    transformed_port = transform_tx_port(raw_port)
    
    if d.get('endpoint_ports'):
        d['endpoint_ports'] = transform_tx_port(d['endpoint_ports'])
    if d.get('cpan_a_end_ports'):
        d['cpan_a_end_ports'] = transform_tx_port(d['cpan_a_end_ports'])
    if d.get('tx_system_port'):
        d['tx_system_port'] = transformed_port
        
    d['tx-system-port'] = transformed_port
    d['vlan'] = d.get('vlan', '') or d.get('s1_c_vlan', '') or d.get('service_vlans', '') or d.get('oam_vlan', '')
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
    if not p_str:
        return ""
    p_str = str(p_str).strip()
    if not p_str or p_str.lower() in ('nan', 'none', '-'):
        return ""
        
    if "/" in p_str:
        tokens = [t.strip() for t in p_str.split("/") if t.strip()]
        if not tokens:
            return ""
            
        s_part = ""
        j_part = ""
        
        for t in tokens:
            if not s_part and re.match(r'^[Ss]\d+$', t):
                s_part = t.upper()
            elif not j_part and re.match(r'^[Jj]\d+$', t):
                j_part = t.upper()
                
        if not s_part and len(tokens) >= 1:
            s_part = tokens[0].upper()
            
        if not j_part:
            for t in tokens[1:]:
                if re.match(r'^[Pp]\d+$', t):
                    j_part = t.upper()
                    break
            if not j_part and len(tokens) >= 3:
                j_part = tokens[2].upper()
            elif not j_part and len(tokens) >= 2:
                j_part = tokens[1].upper()

        if s_part and j_part:
            return f"{s_part}/{j_part}"
        elif s_part:
            return s_part
        else:
            return p_str.upper()
    else:
        return p_str.strip().upper()


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
    'oam_cef_ip_pool': 'OAM CEF IP Pool',
    'ssa': 'SSA',
    'location': 'Location',
    'cpan_maan_vsat': 'Type (CPAN/MAAN/VSAT)',
    'mgmt_ip': 'Mgmt IP',
    'oam_vlan': 'OAM VLAN',
    'mgmt_rac_vlan': 'Mgmt / RAC VLAN',
    's1_c_vlan': 'S1-C VLAN',
    's1_u_vlan': 'S1-U VLAN',
    'endpoint_ip': 'Endpoint IP',
    'endpoint_node_router': 'Endpoint Node/Router',
    'cpan_a_end_node': 'CPAN A End Node',
    'cpan_a_end_ip': 'CPAN A End IP',
    'cpan_z_end_node': 'CPAN Z End Node',
    'cpan_z_end_ip': 'CPAN Z End IP',
    'oam_vlan': 'OAM VLAN',
    'tx_system_ip': 'TX System IP',
    'tx_system_location': 'TX System Location',
    'tx_system_port': 'TX System Port',
    'vlan': 'VLAN'
}


ALL_SEARCHABLE_COLS = [
    'site_id', 'site_name', 'enodeb_address', 'ssa', 'location', 'cpan_maan_vsat',
    'oam_vlan', 'mgmt_rac_vlan', 's1_c_vlan', 's1_u_vlan', 'mgmt_ip', 'mgmt_gateway',
    's1_u_ip', 'mme_ip', 'endpoint_type', 'endpoint_node_router', 'endpoint_ip',
    'l3_gateway_maan', 'endpoint_ports', 'cpan_a_end_node', 'cpan_a_end_ip',
    'cpan_a_end_ports', 'cpan_z_end_node', 'cpan_z_end_ip', 'cpan_service',
    'service_vlans', 'maan_l3_interface', 'maan_vpn', 'oam_cef_ip_pool',
    'oam_hw_gw', 'oam_hw_ip', 'tx_system_ip', 'tx_system_location', 'tx_system_port', 'vlan'
]


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
            or_clauses = [f"LOWER({c}) LIKE LOWER(?)" for c in ALL_SEARCHABLE_COLS]
            where_clauses.append("(" + " OR ".join(or_clauses) + ")")
            params.extend([q_like] * len(ALL_SEARCHABLE_COLS))
        
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


@app.route('/api/site/<site_id>')
@login_required
def api_get_site(site_id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM bts_sites WHERE LOWER(site_id) = LOWER(?);", (site_id,))
    row = cursor.fetchone()
    conn.close()
    
    if not row:
        return jsonify({'success': False, 'message': f'Site ID "{site_id}" not found.'}), 404
        
    return jsonify({'success': True, 'site': row_to_dict(row)})


@app.route('/api/ping', methods=['POST'])
@login_required
def api_ping():
    data = request.get_json(silent=True) or request.form
    raw_ip = str(data.get('ip', '')).strip()
    
    if not raw_ip:
        return jsonify({'success': False, 'message': 'IP address is required.'}), 400
        
    clean_ip = raw_ip.split('/')[0].strip()
    
    try:
        ip_obj = ipaddress.ip_address(clean_ip)
    except ValueError:
        return jsonify({
            'success': False,
            'message': f'"{raw_ip}" is not a valid IPv4/IPv6 address.'
        }), 400
        
    cmd = ['ping', '-n', '4', str(ip_obj)]
    
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=8)
        output = proc.stdout or proc.stderr or "No ping output returned."
        output_upper = output.upper()
        
        is_reachable = (proc.returncode == 0) and (
            'TTL=' in output_upper or
            'BYTES=' in output_upper or
            'TIME=' in output_upper or
            'REPLY FROM' in output_upper
        ) and ('UNREACHABLE' not in output_upper and 'TIMED OUT' not in output_upper and '100% LOSS' not in output_upper)
        
        return jsonify({
            'success': True,
            'ip': str(ip_obj),
            'raw_ip': raw_ip,
            'output': output.strip(),
            'is_reachable': is_reachable,
            'exit_code': proc.returncode
        })
    except subprocess.TimeoutExpired:
        return jsonify({
            'success': False,
            'ip': str(ip_obj),
            'raw_ip': raw_ip,
            'output': f'Ping request to {ip_obj} timed out (exceeded 8s timeout).',
            'is_reachable': False,
            'exit_code': -1
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error executing ping: {str(e)}'
        }), 500


@app.route('/site/new', methods=['GET', 'POST'])
@login_required
def create_site():
    if request.method == 'POST':
        site_id = request.form.get('site_id', '').strip()
        enodeb_address = request.form.get('enodeb_address', '').strip()
        site_name = request.form.get('site_name', '').strip()
        ssa = request.form.get('ssa', '').strip()
        location = request.form.get('Location', request.form.get('location', '')).strip()
        cpan_maan = request.form.get('cpan_maan_vsat', request.form.get('cpan/maan/vsat', '')).strip()
        
        oam_vlan = request.form.get('oam_vlan', '').strip()
        mgmt_rac_vlan = request.form.get('mgmt_rac_vlan', '').strip()
        s1_c_vlan = request.form.get('s1_c_vlan', '').strip()
        s1_u_vlan = request.form.get('s1_u_vlan', '').strip()
        mgmt_ip = request.form.get('mgmt_ip', '').strip()
        mgmt_gateway = request.form.get('mgmt_gateway', '').strip()
        s1_u_ip = request.form.get('s1_u_ip', '').strip()
        mme_ip = request.form.get('mme_ip', '').strip()
        endpoint_type = request.form.get('endpoint_type', '').strip()
        endpoint_node_router = request.form.get('endpoint_node_router', '').strip()
        endpoint_ip = request.form.get('endpoint_ip', '').strip()
        l3_gateway_maan = request.form.get('l3_gateway_maan', '').strip()
        endpoint_ports = transform_tx_port(request.form.get('endpoint_ports', '').strip())
        cpan_a_end_node = request.form.get('cpan_a_end_node', '').strip()
        cpan_a_end_ip = request.form.get('cpan_a_end_ip', '').strip()
        cpan_a_end_ports = transform_tx_port(request.form.get('cpan_a_end_ports', '').strip())
        cpan_z_end_node = request.form.get('cpan_z_end_node', '').strip()
        cpan_z_end_ip = request.form.get('cpan_z_end_ip', '').strip()
        cpan_service = request.form.get('cpan_service', '').strip()
        service_vlans = request.form.get('service_vlans', '').strip()
        maan_l3_interface = request.form.get('maan_l3_interface', '').strip()
        maan_vpn = request.form.get('maan_vpn', '').strip()
        mask = request.form.get('mask', '').strip()
        route_distinguisher = request.form.get('route_distinguisher', '').strip()
        as_num = request.form.get('as_num', '').strip()
        ems = request.form.get('ems', '').strip()
        oam_cef_ip_pool = request.form.get('oam_cef_ip_pool', '').strip()
        oam_hw_gw = request.form.get('oam_hw_gw', '').strip()
        oam_hw_ip = request.form.get('oam_hw_ip', '').strip()
        
        tx_ip = endpoint_ip or cpan_a_end_ip or request.form.get('tx-system-ip', '').strip()
        tx_loc = endpoint_node_router or cpan_a_end_node or request.form.get('tx-system-location', '').strip()
        tx_port = endpoint_ports or cpan_a_end_ports or transform_tx_port(request.form.get('tx-system-port', '').strip())
        vlan = s1_c_vlan or service_vlans or oam_vlan or request.form.get('vlan', '').strip()
        
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
                enodeb_address, site_id, site_name, ssa, location, cpan_maan_vsat,
                oam_vlan, mgmt_rac_vlan, s1_c_vlan, s1_u_vlan, mgmt_ip, mgmt_gateway,
                s1_u_ip, mme_ip, endpoint_type, endpoint_node_router, endpoint_ip,
                l3_gateway_maan, endpoint_ports, cpan_a_end_node, cpan_a_end_ip,
                cpan_a_end_ports, cpan_z_end_node, cpan_z_end_ip, cpan_service,
                service_vlans, maan_l3_interface, maan_vpn, mask, route_distinguisher,
                as_num, ems, oam_cef_ip_pool, oam_hw_gw, oam_hw_ip,
                tx_system_ip, tx_system_location, tx_system_port, vlan
            ) VALUES (
                ?, ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?,
                ?, ?, ?, ?,
                ?, ?, ?, ?,
                ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?,
                ?, ?, ?, ?
            );
        ''', (
            enodeb_address, site_id, site_name, ssa, location, cpan_maan,
            oam_vlan, mgmt_rac_vlan, s1_c_vlan, s1_u_vlan, mgmt_ip, mgmt_gateway,
            s1_u_ip, mme_ip, endpoint_type, endpoint_node_router, endpoint_ip,
            l3_gateway_maan, endpoint_ports, cpan_a_end_node, cpan_a_end_ip,
            cpan_a_end_ports, cpan_z_end_node, cpan_z_end_ip, cpan_service,
            service_vlans, maan_l3_interface, maan_vpn, mask, route_distinguisher,
            as_num, ems, oam_cef_ip_pool, oam_hw_gw, oam_hw_ip,
            tx_ip, tx_loc, tx_port, vlan
        ))
        
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
        location = request.form.get('Location', request.form.get('location', '')).strip()
        cpan_maan = request.form.get('cpan_maan_vsat', request.form.get('cpan/maan/vsat', '')).strip()
        
        oam_vlan = request.form.get('oam_vlan', '').strip()
        mgmt_rac_vlan = request.form.get('mgmt_rac_vlan', '').strip()
        s1_c_vlan = request.form.get('s1_c_vlan', '').strip()
        s1_u_vlan = request.form.get('s1_u_vlan', '').strip()
        mgmt_ip = request.form.get('mgmt_ip', '').strip()
        mgmt_gateway = request.form.get('mgmt_gateway', '').strip()
        s1_u_ip = request.form.get('s1_u_ip', '').strip()
        mme_ip = request.form.get('mme_ip', '').strip()
        endpoint_type = request.form.get('endpoint_type', '').strip()
        endpoint_node_router = request.form.get('endpoint_node_router', '').strip()
        endpoint_ip = request.form.get('endpoint_ip', '').strip()
        l3_gateway_maan = request.form.get('l3_gateway_maan', '').strip()
        endpoint_ports = transform_tx_port(request.form.get('endpoint_ports', '').strip())
        cpan_a_end_node = request.form.get('cpan_a_end_node', '').strip()
        cpan_a_end_ip = request.form.get('cpan_a_end_ip', '').strip()
        cpan_a_end_ports = transform_tx_port(request.form.get('cpan_a_end_ports', '').strip())
        cpan_z_end_node = request.form.get('cpan_z_end_node', '').strip()
        cpan_z_end_ip = request.form.get('cpan_z_end_ip', '').strip()
        cpan_service = request.form.get('cpan_service', '').strip()
        service_vlans = request.form.get('service_vlans', '').strip()
        maan_l3_interface = request.form.get('maan_l3_interface', '').strip()
        maan_vpn = request.form.get('maan_vpn', '').strip()
        mask = request.form.get('mask', '').strip()
        route_distinguisher = request.form.get('route_distinguisher', '').strip()
        as_num = request.form.get('as_num', '').strip()
        ems = request.form.get('ems', '').strip()
        oam_cef_ip_pool = request.form.get('oam_cef_ip_pool', '').strip()
        oam_hw_gw = request.form.get('oam_hw_gw', '').strip()
        oam_hw_ip = request.form.get('oam_hw_ip', '').strip()
        
        tx_ip = endpoint_ip or cpan_a_end_ip or request.form.get('tx-system-ip', '').strip()
        tx_loc = endpoint_node_router or cpan_a_end_node or request.form.get('tx-system-location', '').strip()
        tx_port = endpoint_ports or cpan_a_end_ports or transform_tx_port(request.form.get('tx-system-port', '').strip())
        vlan = s1_c_vlan or service_vlans or oam_vlan or request.form.get('vlan', '').strip()
        
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
                oam_vlan = ?,
                mgmt_rac_vlan = ?,
                s1_c_vlan = ?,
                s1_u_vlan = ?,
                mgmt_ip = ?,
                mgmt_gateway = ?,
                s1_u_ip = ?,
                mme_ip = ?,
                endpoint_type = ?,
                endpoint_node_router = ?,
                endpoint_ip = ?,
                l3_gateway_maan = ?,
                endpoint_ports = ?,
                cpan_a_end_node = ?,
                cpan_a_end_ip = ?,
                cpan_a_end_ports = ?,
                cpan_z_end_node = ?,
                cpan_z_end_ip = ?,
                cpan_service = ?,
                service_vlans = ?,
                maan_l3_interface = ?,
                maan_vpn = ?,
                mask = ?,
                route_distinguisher = ?,
                as_num = ?,
                ems = ?,
                oam_cef_ip_pool = ?,
                oam_hw_gw = ?,
                oam_hw_ip = ?,
                tx_system_ip = ?,
                tx_system_location = ?,
                tx_system_port = ?,
                vlan = ?,
                updated_at = CURRENT_TIMESTAMP
            WHERE LOWER(site_id) = LOWER(?);
        ''', (
            enodeb_address, site_name, ssa, location, cpan_maan,
            oam_vlan, mgmt_rac_vlan, s1_c_vlan, s1_u_vlan, mgmt_ip, mgmt_gateway,
            s1_u_ip, mme_ip, endpoint_type, endpoint_node_router, endpoint_ip,
            l3_gateway_maan, endpoint_ports, cpan_a_end_node, cpan_a_end_ip,
            cpan_a_end_ports, cpan_z_end_node, cpan_z_end_ip, cpan_service,
            service_vlans, maan_l3_interface, maan_vpn, mask, route_distinguisher,
            as_num, ems, oam_cef_ip_pool, oam_hw_gw, oam_hw_ip,
            tx_ip, tx_loc, tx_port, vlan, site_id
        ))
        
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
    'site_id': 'Site Id',
    'site_name': 'Site Name',
    'enodeb_address': 'eNodeB Address',
    'ssa': 'SSA',
    'location': 'Location',
    'cpan_maan_vsat': 'MAAN/CPAN/VSAT',
    'oam_vlan': 'OAM VLAN',
    'mgmt_rac_vlan': 'Mgmt / RAC VLAN',
    's1_c_vlan': 'S1-C VLAN',
    's1_u_vlan': 'S1-U VLAN',
    'mgmt_ip': 'Mgmt IP',
    'mgmt_gateway': 'Mgmt Gateway',
    's1_u_ip': 'S1-U IP',
    'mme_ip': 'MME IP',
    'endpoint_type': 'Endpoint Type',
    'endpoint_node_router': 'Endpoint — Node / Router',
    'endpoint_ip': 'Endpoint IP',
    'l3_gateway_maan': 'L3 Gateway (MAAN)',
    'endpoint_ports': 'Endpoint Port(s)',
    'cpan_a_end_node': 'CPAN A End Node',
    'cpan_a_end_ip': 'CPAN A End IP',
    'cpan_a_end_ports': 'CPAN A End Port(s)',
    'cpan_z_end_node': 'CPAN Z End Node',
    'cpan_z_end_ip': 'CPAN Z End IP',
    'cpan_service': 'CPAN Service',
    'service_vlans': 'Service VLANs',
    'maan_l3_interface': 'MAAN L3 Interface',
    'maan_vpn': 'MAAN VPN',
    'mask': 'Mask',
    'route_distinguisher': 'Route Distinguisher',
    'as_num': 'AS',
    'ems': 'EMS',
    'oam_cef_ip_pool': 'OAM CEF IP pool',
    'oam_hw_gw': 'OAM HW GW',
    'oam_hw_ip': 'OAM HW IP',
    'tx_system_ip': 'tx-system-ip',
    'tx_system_location': 'tx-system-location',
    'tx_system_port': 'tx-system-port',
    'vlan': 'VLAN'
}


@app.route('/export')
@login_required
def export_excel():
    conn = get_db()
    export_sql = '''
        SELECT
            enodeb_address AS 'eNodeB Address',
            site_id AS 'Site Id',
            site_name AS 'Site Name',
            cpan_maan_vsat AS 'MAAN/CPAN/VSAT',
            oam_vlan AS 'OAM VLAN',
            mgmt_rac_vlan AS 'Mgmt / RAC VLAN',
            s1_c_vlan AS 'S1-C VLAN',
            s1_u_vlan AS 'S1-U VLAN',
            mgmt_ip AS 'Mgmt IP',
            mgmt_gateway AS 'Mgmt Gateway',
            s1_u_ip AS 'S1-U IP',
            mme_ip AS 'MME IP',
            endpoint_type AS 'Endpoint Type',
            endpoint_node_router AS 'Endpoint — Node / Router',
            endpoint_ip AS 'Endpoint IP',
            l3_gateway_maan AS 'L3 Gateway (MAAN)',
            endpoint_ports AS 'Endpoint Port(s)',
            cpan_a_end_node AS 'CPAN A End Node',
            cpan_a_end_ip AS 'CPAN A End IP',
            cpan_a_end_ports AS 'CPAN A End Port(s)',
            cpan_z_end_node AS 'CPAN Z End Node',
            cpan_z_end_ip AS 'CPAN Z End IP',
            cpan_service AS 'CPAN Service',
            service_vlans AS 'Service VLANs',
            maan_l3_interface AS 'MAAN L3 Interface',
            maan_vpn AS 'MAAN VPN',
            mask AS 'Mask',
            route_distinguisher AS 'Route Distinguisher',
            as_num AS 'AS',
            ems AS 'EMS',
            oam_cef_ip_pool AS 'OAM CEF IP pool',
            oam_hw_gw AS 'OAM HW GW',
            oam_hw_ip AS 'OAM HW IP'
        FROM bts_sites ORDER BY id ASC;
    '''
    df = pd.read_sql_query(export_sql, conn)
    conn.close()
    
    target_excel = NEW_EXCEL_PATH if os.path.exists(NEW_EXCEL_PATH) else EXCEL_PATH
    try:
        df.to_excel(target_excel, index=False, sheet_name='Sheet1')
    except Exception as e:
        print(f"Could not update excel file directly: {e}")
        
    return send_file(
        target_excel,
        as_attachment=True,
        download_name='btsdatabase_updated.xlsx',
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
                or_clauses = [f"LOWER({c}) LIKE LOWER(?)" for c in ALL_SEARCHABLE_COLS]
                where_clauses.append("(" + " OR ".join(or_clauses) + ")")
                params.extend([q_like] * len(ALL_SEARCHABLE_COLS))
                
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
