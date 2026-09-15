import os
import re
import sqlite3
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'data')
os.makedirs(DATA_DIR, exist_ok=True)

DB_PATH = os.path.join(DATA_DIR, 'btsdatabase.db') if os.path.exists(os.path.join(DATA_DIR, 'btsdatabase.db')) else os.path.join(BASE_DIR, 'btsdatabase.db')

NEW_EXCEL_PATH = os.path.join(DATA_DIR, 'btsdatabase_updated - v1.xlsx') if os.path.exists(os.path.join(DATA_DIR, 'btsdatabase_updated - v1.xlsx')) else os.path.join(BASE_DIR, 'btsdatabase_updated - v1.xlsx')
OLD_EXCEL_PATH = os.path.join(DATA_DIR, 'btsdatabase.xlsx') if os.path.exists(os.path.join(DATA_DIR, 'btsdatabase.xlsx')) else os.path.join(BASE_DIR, 'btsdatabase.xlsx')

EXCEL_PATH = NEW_EXCEL_PATH if os.path.exists(NEW_EXCEL_PATH) else OLD_EXCEL_PATH


def get_db_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA cache_size = -65536;")  # 64 MB in-memory SQL cache
    conn.execute("PRAGMA temp_store = MEMORY;")   # Keep temporary tables and indices in RAM
    return conn


def derive_ssa_and_location(site_id, site_name, old_ssa_map={}, old_loc_map={}, prefix_map={}):
    site_id_str = str(site_id).strip()
    site_name_str = str(site_name).strip()
    
    ssa = old_ssa_map.get(site_id_str, "")
    loc = old_loc_map.get(site_id_str, "")
    
    if not ssa or ssa.lower() == 'nan':
        m = re.match(r'^[A-Z0-9]+?([A-Z]{3})\d+', site_id_str)
        if m and m.group(1) in prefix_map:
            ssa = prefix_map[m.group(1)]
        elif '_' in site_name_str:
            ssa = site_name_str.split('_')[0]
        elif len(site_name_str) >= 6:
            ssa = site_name_str[:6]
        else:
            ssa = site_name_str
            
    if not loc or loc.lower() == 'nan':
        if '_' in site_name_str:
            loc = site_name_str.split('_', 1)[1]
        else:
            loc = site_name_str
            
    if ssa.lower() == 'nan':
        ssa = ""
    if loc.lower() == 'nan':
        loc = ""
        
    return ssa, loc


def clean_val(val):
    if pd.isna(val) or val is None:
        return ""
    v_str = str(val).strip()
    if v_str.lower() in ('nan', 'none', '<na>'):
        return ""
    if v_str.endswith('.0'):
        v_str = v_str[:-2]
    return v_str


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


def init_db(force_reimport=False):
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # Create or update bts_sites table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS bts_sites (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            enodeb_address TEXT,
            site_id TEXT UNIQUE NOT NULL,
            site_name TEXT,
            ssa TEXT,
            location TEXT,
            cpan_maan_vsat TEXT,
            oam_vlan TEXT,
            mgmt_rac_vlan TEXT,
            s1_c_vlan TEXT,
            s1_u_vlan TEXT,
            mgmt_ip TEXT,
            mgmt_gateway TEXT,
            s1_u_ip TEXT,
            mme_ip TEXT,
            endpoint_type TEXT,
            endpoint_node_router TEXT,
            endpoint_ip TEXT,
            l3_gateway_maan TEXT,
            endpoint_ports TEXT,
            cpan_a_end_node TEXT,
            cpan_a_end_ip TEXT,
            cpan_a_end_ports TEXT,
            cpan_z_end_node TEXT,
            cpan_z_end_ip TEXT,
            cpan_service TEXT,
            service_vlans TEXT,
            maan_l3_interface TEXT,
            maan_vpn TEXT,
            mask TEXT,
            route_distinguisher TEXT,
            as_num TEXT,
            ems TEXT,
            oam_cef_ip_pool TEXT,
            oam_hw_gw TEXT,
            oam_hw_ip TEXT,
            tx_system_ip TEXT,
            tx_system_location TEXT,
            tx_system_port TEXT,
            vlan TEXT,
            reason TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    ''')
    
    # Ensure missing columns exist if table was already created earlier
    cursor.execute("PRAGMA table_info(bts_sites);")
    existing_cols = {col[1] for col in cursor.fetchall()}
    
    expected_cols = [
        'enodeb_address', 'site_id', 'site_name', 'ssa', 'location', 'cpan_maan_vsat',
        'oam_vlan', 'mgmt_rac_vlan', 's1_c_vlan', 's1_u_vlan', 'mgmt_ip', 'mgmt_gateway',
        's1_u_ip', 'mme_ip', 'endpoint_type', 'endpoint_node_router', 'endpoint_ip',
        'l3_gateway_maan', 'endpoint_ports', 'cpan_a_end_node', 'cpan_a_end_ip',
        'cpan_a_end_ports', 'cpan_z_end_node', 'cpan_z_end_ip', 'cpan_service',
        'service_vlans', 'maan_l3_interface', 'maan_vpn', 'mask', 'route_distinguisher',
        'as_num', 'ems', 'oam_cef_ip_pool', 'oam_hw_gw', 'oam_hw_ip',
        'tx_system_ip', 'tx_system_location', 'tx_system_port', 'vlan', 'reason'
    ]
    
    for c in expected_cols:
        if c not in existing_cols:
            cursor.execute(f"ALTER TABLE bts_sites ADD COLUMN {c} TEXT;")
            
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_site_id ON bts_sites(site_id);')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_ssa ON bts_sites(ssa);')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_site_name ON bts_sites(site_name);')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_enodeb ON bts_sites(enodeb_address);')
    
    # Users table for authentication
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT,
            staff_no TEXT UNIQUE NOT NULL,
            full_name TEXT NOT NULL,
            email TEXT,
            department TEXT,
            password_hash TEXT NOT NULL,
            role TEXT DEFAULT 'staff',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    ''')
    
    cursor.execute("PRAGMA table_info(users);")
    user_cols = [col[1] for col in cursor.fetchall()]
    if 'username' not in user_cols:
        cursor.execute("ALTER TABLE users ADD COLUMN username TEXT;")
        
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_staff_no ON users(staff_no);')
    cursor.execute('CREATE UNIQUE INDEX IF NOT EXISTS idx_username ON users(username);')
    
    cursor.execute("DELETE FROM users WHERE staff_no = 'STAFF001';")
    conn.commit()

    # Activity Logs table for tracking system edits & audit trail
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS activity_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            username TEXT,
            action TEXT NOT NULL,
            target_type TEXT,
            target_id TEXT,
            details TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    ''')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_activity_created ON activity_logs(created_at);')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_activity_target ON activity_logs(target_id);')

    # Uploaded Log Files table for Log Analyzer
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS uploaded_log_files (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT NOT NULL,
            stored_filename TEXT UNIQUE NOT NULL,
            file_type TEXT,
            file_size INTEGER,
            line_count INTEGER,
            uploaded_by TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    ''')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_log_files_created ON uploaded_log_files(created_at);')

    # Create cpan_nodes table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS cpan_nodes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ne_ip TEXT,
            location TEXT,
            type TEXT,
            ssa TEXT,
            phase TEXT,
            ne_name TEXT,
            ne_id TEXT,
            dcc_ip TEXT,
            software_version TEXT,
            hardware_version TEXT,
            pcb_version TEXT,
            last_upload_time TEXT,
            orig_location TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    ''')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_cpan_ne_ip ON cpan_nodes(ne_ip);')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_cpan_location ON cpan_nodes(location);')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_cpan_type ON cpan_nodes(type);')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_cpan_ssa ON cpan_nodes(ssa);')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_cpan_ne_name ON cpan_nodes(ne_name);')
    
    cursor.execute('SELECT COUNT(*) FROM users;')
    user_count = cursor.fetchone()[0]
    
    cursor.execute('SELECT COUNT(*) FROM bts_sites;')
    count = cursor.fetchone()[0]
    
    # Check if re-import needed
    if count == 0 or force_reimport:
        if os.path.exists(EXCEL_PATH):
            print(f"Migrating records from Excel ({EXCEL_PATH}) into SQLite ({DB_PATH})...")
            
            # Load old btsdatabase.xlsx if available to build SSA lookup map
            old_ssa_map = {}
            old_loc_map = {}
            prefix_map = {}
            if os.path.exists(OLD_EXCEL_PATH):
                try:
                    df_old = pd.read_excel(OLD_EXCEL_PATH, dtype=str).fillna("")
                    for _, row in df_old.iterrows():
                        sid = str(row.get('site_id', '')).strip()
                        s_ssa = str(row.get('ssa', '')).strip()
                        s_loc = str(row.get('Location', '')).strip()
                        if sid:
                            old_ssa_map[sid] = s_ssa
                            old_loc_map[sid] = s_loc
                            if s_ssa and s_ssa.lower() != 'nan':
                                m = re.match(r'^[A-Z0-9]+?([A-Z]{3})\d+', sid)
                                if m:
                                    prefix_map[m.group(1)] = s_ssa
                except Exception as e:
                    print(f"Note: Could not load old excel for lookup: {e}")
                    
            with open(EXCEL_PATH, 'rb') as f:
                df = pd.read_excel(f, sheet_name=0, dtype=str)
                
            df = df.fillna("")
            
            # Identify columns dynamically
            cols_map = {str(col).strip(): col for col in df.columns}
            
            records_to_insert = []
            for _, row in df.iterrows():
                site_id = clean_val(row.get(cols_map.get('Site Id', 'site_id'), ''))
                if not site_id:
                    continue
                    
                site_name = clean_val(row.get(cols_map.get('Site Name', 'site_name'), ''))
                enodeb_address = clean_val(row.get(cols_map.get('eNodeB Address', 'enodeb_address'), ''))
                cpan_maan_vsat = clean_val(row.get(cols_map.get('MAAN/CPAN/VSAT', 'cpan/maan/vsat'), ''))
                
                # Derive SSA & Location
                ssa = clean_val(row.get('ssa', ''))
                location = clean_val(row.get('Location', ''))
                if not ssa or not location:
                    d_ssa, d_loc = derive_ssa_and_location(site_id, site_name, old_ssa_map, old_loc_map, prefix_map)
                    if not ssa:
                        ssa = d_ssa
                    if not location:
                        location = d_loc
                        
                oam_vlan = clean_val(row.get('OAM VLAN', ''))
                mgmt_rac_vlan = clean_val(row.get('Mgmt / RAC VLAN', ''))
                s1_c_vlan = clean_val(row.get('S1-C VLAN', ''))
                s1_u_vlan = clean_val(row.get('S1-U VLAN', ''))
                mgmt_ip = clean_val(row.get('Mgmt IP', ''))
                mgmt_gateway = clean_val(row.get('Mgmt Gateway', ''))
                s1_u_ip = clean_val(row.get('S1-U IP', ''))
                mme_ip = clean_val(row.get('MME IP', ''))
                endpoint_type = clean_val(row.get('Endpoint Type', ''))
                
                # Endpoint Node/Router has special character handling
                ep_node_key = [k for k in cols_map.keys() if 'Endpoint' in k and 'Node' in k]
                endpoint_node_router = clean_val(row.get(cols_map.get(ep_node_key[0]), '')) if ep_node_key else clean_val(row.get('tx-system-location', ''))
                
                endpoint_ip = clean_val(row.get('Endpoint IP', ''))
                l3_gateway_maan = clean_val(row.get('L3 Gateway (MAAN)', ''))
                endpoint_ports = transform_tx_port(clean_val(row.get('Endpoint Port(s)', '')))
                cpan_a_end_node = clean_val(row.get('CPAN A End Node', ''))
                cpan_a_end_ip = clean_val(row.get('CPAN A End IP', ''))
                cpan_a_end_ports = transform_tx_port(clean_val(row.get('CPAN A End Port(s)', '')))
                cpan_z_end_node = clean_val(row.get('CPAN Z End Node', ''))
                cpan_z_end_ip = clean_val(row.get('CPAN Z End IP', ''))
                cpan_service = clean_val(row.get('CPAN Service', ''))
                service_vlans = clean_val(row.get('Service VLANs', ''))
                maan_l3_interface = clean_val(row.get('MAAN L3 Interface', ''))
                maan_vpn = clean_val(row.get('MAAN VPN', ''))
                mask = clean_val(row.get('Mask', ''))
                route_distinguisher = clean_val(row.get('Route Distinguisher', ''))
                as_num = clean_val(row.get('AS', ''))
                ems = clean_val(row.get('EMS', ''))
                oam_cef_ip_pool = clean_val(row.get('OAM CEF IP pool', ''))
                oam_hw_gw = clean_val(row.get('OAM HW GW', ''))
                oam_hw_ip = clean_val(row.get('OAM HW IP', ''))
                
                # Legacy alias mappings
                tx_system_ip = endpoint_ip or cpan_a_end_ip or clean_val(row.get('tx-system-ip', ''))
                tx_system_location = endpoint_node_router or cpan_a_end_node or clean_val(row.get('tx-system-location', ''))
                tx_system_port = endpoint_ports or cpan_a_end_ports or transform_tx_port(clean_val(row.get('tx-system-port', '')))
                vlan = s1_c_vlan or service_vlans or oam_vlan or clean_val(row.get('vlan', ''))
                
                records_to_insert.append((
                    enodeb_address, site_id, site_name, ssa, location, cpan_maan_vsat,
                    oam_vlan, mgmt_rac_vlan, s1_c_vlan, s1_u_vlan, mgmt_ip, mgmt_gateway,
                    s1_u_ip, mme_ip, endpoint_type, endpoint_node_router, endpoint_ip,
                    l3_gateway_maan, endpoint_ports, cpan_a_end_node, cpan_a_end_ip,
                    cpan_a_end_ports, cpan_z_end_node, cpan_z_end_ip, cpan_service,
                    service_vlans, maan_l3_interface, maan_vpn, mask, route_distinguisher,
                    as_num, ems, oam_cef_ip_pool, oam_hw_gw, oam_hw_ip,
                    tx_system_ip, tx_system_location, tx_system_port, vlan
                ))
                
            if force_reimport:
                cursor.execute('DELETE FROM bts_sites;')
                
            cursor.executemany('''
                INSERT OR REPLACE INTO bts_sites (
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
            ''', records_to_insert)
            
            conn.commit()
            cursor.execute('SELECT COUNT(*) FROM bts_sites;')
            new_count = cursor.fetchone()[0]
            print(f"Successfully migrated {new_count} records into SQLite database.")
        else:
            print("No excel file found. Initialized empty database.")
    else:
        print(f"Database initialized with {count} site records and {user_count} registered user(s).")

    # Import CPAN nodes if table is empty or force_reimport
    cursor.execute('SELECT COUNT(*) FROM cpan_nodes;')
    cpan_node_count = cursor.fetchone()[0]
    if cpan_node_count == 0 or force_reimport:
        cpan_csv = os.path.join(DATA_DIR, 'GUJ_CPAN-NODE_LIST.csv') if os.path.exists(os.path.join(DATA_DIR, 'GUJ_CPAN-NODE_LIST.csv')) else os.path.join(BASE_DIR, 'GUJ_CPAN-NODE_LIST.csv')
        cpan_processed_csv = os.path.join(DATA_DIR, 'cpandatabase.csv') if os.path.exists(os.path.join(DATA_DIR, 'cpandatabase.csv')) else os.path.join(BASE_DIR, 'cpandatabase.csv')
        source_path = cpan_csv if os.path.exists(cpan_csv) else (cpan_processed_csv if os.path.exists(cpan_processed_csv) else None)
        if source_path:
            print(f"Importing CPAN Nodes from {source_path}...")
            import_cpan_nodes_csv(source_path, conn)

        service_csv = os.path.join(DATA_DIR, 'CPAN_Service_List.csv') if os.path.exists(os.path.join(DATA_DIR, 'CPAN_Service_List.csv')) else os.path.join(BASE_DIR, 'CPAN_Service_List.csv')
        if os.path.exists(service_csv):
            print(f"Importing CPAN Services from {service_csv}...")
            import_cpan_services_csv(service_csv, conn)

    # Ensure cpan_services table exists and has data
    cursor.execute("SELECT COUNT(*) FROM cpan_services;")
    srv_count = cursor.fetchone()[0]
    if srv_count == 0:
        service_csv = os.path.join(DATA_DIR, 'CPAN_Service_List.csv') if os.path.exists(os.path.join(DATA_DIR, 'CPAN_Service_List.csv')) else os.path.join(BASE_DIR, 'CPAN_Service_List.csv')
        if os.path.exists(service_csv):
            print(f"Importing CPAN Services from {service_csv}...")
            import_cpan_services_csv(service_csv, conn)

    # Ensure cpan_dl_list table exists and has data
    cursor.execute("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='cpan_dl_list';")
    if not cursor.fetchone():
        dl_count = 0
    else:
        cursor.execute("SELECT COUNT(*) FROM cpan_dl_list;")
        dl_count = cursor.fetchone()[0]
    if dl_count == 0:
        dl_csv = os.path.join(DATA_DIR, 'CPAN_DL_LIST.csv') if os.path.exists(os.path.join(DATA_DIR, 'CPAN_DL_LIST.csv')) else os.path.join(BASE_DIR, 'CPAN_DL_LIST.csv')
        if os.path.exists(dl_csv):
            print(f"Importing CPAN DL List from {dl_csv}...")
            import_cpan_dl_list_csv(dl_csv, conn)

        init_maan_tables(conn)

    conn.close()


def import_cpan_nodes_csv(csv_source, conn=None):
    close_conn = False
    if conn is None:
        conn = get_db_connection()
        close_conn = True
    cursor = conn.cursor()
    
    if isinstance(csv_source, str) and os.path.exists(csv_source):
        df = pd.read_csv(csv_source)
    elif isinstance(csv_source, pd.DataFrame):
        df = csv_source
    else:
        if close_conn:
            conn.close()
        return 0

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS cpan_nodes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ne_ip TEXT,
            location TEXT,
            type TEXT,
            ssa TEXT,
            phase TEXT,
            ne_name TEXT,
            ne_id TEXT,
            dcc_ip TEXT,
            software_version TEXT,
            hardware_version TEXT,
            pcb_version TEXT,
            last_upload_time TEXT,
            orig_location TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    ''')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_cpan_ne_ip ON cpan_nodes(ne_ip);')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_cpan_location ON cpan_nodes(location);')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_cpan_type ON cpan_nodes(type);')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_cpan_ssa ON cpan_nodes(ssa);')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_cpan_ne_name ON cpan_nodes(ne_name);')
    
    records = []
    for idx, row in df.iterrows():
        ne_name_raw = row.get('NE Name', row.get('ne_name', ''))
        if pd.isna(ne_name_raw) or not str(ne_name_raw).strip():
            continue
            
        ne_name = str(ne_name_raw).strip()
        obj_fid = str(row.get('Object Fid', '')).strip() if pd.notna(row.get('Object Fid')) else ''
        
        ne_ip = str(row.get('ne_ip', '')).strip() if pd.notna(row.get('ne_ip')) else ''
        location = str(row.get('location', '')).strip() if pd.notna(row.get('location')) else ''
        ne_type = str(row.get('type', '')).strip() if pd.notna(row.get('type')) else ''
        ssa = str(row.get('ssa', '')).strip() if pd.notna(row.get('ssa')) else ''
        phase = str(row.get('phase', '')).strip() if pd.notna(row.get('phase')) else ''
        
        # If pre-parsed fields are not available in CSV, parse from NE Name
        if not (ne_ip and location and ne_type and ssa):
            parts = ne_name.split('_')
            if len(parts) == 6:
                ne_ip = parts[0].strip()
                location = parts[1].strip()
                ne_type = parts[2].strip()
                ssa = f"{parts[3].strip()}_{parts[4].strip()}"
                phase = parts[5].strip()
            else:
                clean_fid = obj_fid.lstrip('\\').strip()
                fid_parts = clean_fid.split('_')
                if len(fid_parts) == 6:
                    ne_ip = fid_parts[0].strip()
                    location = fid_parts[1].strip()
                    ne_type = fid_parts[2].strip()
                    ssa = f"{fid_parts[3].strip()}_{fid_parts[4].strip()}"
                    phase = fid_parts[5].strip()
                else:
                    ip_match = re.match(r'^(\d+\.\d+\.\d+\.\d+)', ne_name)
                    ne_ip = ip_match.group(1) if ip_match else str(row.get('DCC IP', '')).strip()
                    type_match = re.search(r'(TN\d+[A-Z]?)', ne_name)
                    ne_type = type_match.group(1) if type_match else str(row.get('NE Type', '')).split()[0]
                    loc_match = re.search(r'^\d+\.\d+\.\d+\.\d+[-_](.*?)[-_]TN', ne_name)
                    location = loc_match.group(1).strip() if loc_match else ""
                    if "BHAVNAGAR" in ne_name.upper():
                        ssa = "GJ_SSA BV"
                    else:
                        ssa = ""
                    phase_match = re.search(r'(PH\d*)', ne_name)
                    phase = phase_match.group(1) if phase_match else ""

        records.append((
            ne_ip, location, ne_type, ssa, phase, ne_name,
            str(row.get('NE ID', row.get('ne_id', ''))).lstrip('\\').strip() if pd.notna(row.get('NE ID', row.get('ne_id'))) else '',
            str(row.get('DCC IP', row.get('dcc_ip', ''))).strip() if pd.notna(row.get('DCC IP', row.get('dcc_ip'))) else '',
            str(row.get('Software Version', row.get('software_version', ''))).strip() if pd.notna(row.get('Software Version', row.get('software_version'))) else '',
            str(row.get('Hardware Version', row.get('hardware_version', ''))).strip() if pd.notna(row.get('Hardware Version', row.get('hardware_version'))) else '',
            str(row.get('PCB Version', row.get('pcb_version', ''))).strip() if pd.notna(row.get('PCB Version', row.get('pcb_version'))) else '',
            str(row.get('Last Upload Time', row.get('last_upload_time', ''))).strip() if pd.notna(row.get('Last Upload Time', row.get('last_upload_time'))) else '',
            str(row.get('Location', row.get('orig_location', ''))).strip() if pd.notna(row.get('Location', row.get('orig_location'))) else ''
        ))
        
    cursor.execute('DELETE FROM cpan_nodes;')
    cursor.executemany('''
        INSERT INTO cpan_nodes (
            ne_ip, location, type, ssa, phase, ne_name,
            ne_id, dcc_ip, software_version, hardware_version,
            pcb_version, last_upload_time, orig_location
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
    ''', records)
    conn.commit()
    count = len(records)
    print(f"Imported {count} CPAN nodes into cpan_nodes table.")
    if close_conn:
        conn.close()
    return count


def import_cpan_services_csv(csv_source, conn=None):
    close_conn = False
    if conn is None:
        conn = get_db_connection()
        close_conn = True
    cursor = conn.cursor()

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS cpan_services (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            service_index INTEGER,
            name TEXT,
            service_type TEXT,
            bandwidth_kbps REAL,
            traffic_cir_kbps REAL,
            traffic_eir_kbps REAL,
            network_cir_kbps REAL,
            network_eir_kbps REAL,
            cos TEXT,
            trust_ce_qos TEXT,
            order_name TEXT,
            client TEXT,
            a_end TEXT,
            z_end TEXT,
            create_time TEXT,
            update_time TEXT,
            description TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    ''')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_cpan_services_name ON cpan_services(name);')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_cpan_services_type ON cpan_services(service_type);')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_cpan_services_client ON cpan_services(client);')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_cpan_services_order ON cpan_services(order_name);')

    if isinstance(csv_source, str) and os.path.exists(csv_source):
        df = pd.read_csv(csv_source)
    elif isinstance(csv_source, pd.DataFrame):
        df = csv_source
    else:
        if close_conn:
            conn.close()
        return 0

    records = []
    for idx, row in df.iterrows():
        name = str(row.get('Name', row.get('name', ''))).strip() if pd.notna(row.get('Name', row.get('name'))) else ''
        if not name and pd.isna(row.get('Index', row.get('service_index'))):
            continue

        s_idx_val = row.get('Index', row.get('service_index'))
        s_idx = int(s_idx_val) if pd.notna(s_idx_val) and str(s_idx_val).isdigit() else None
        stype = str(row.get('Service Type', row.get('service_type', ''))).strip() if pd.notna(row.get('Service Type', row.get('service_type'))) else ''
        bw_val = row.get('Bandwidth(Kbps)', row.get('bandwidth_kbps'))
        bw = float(bw_val) if pd.notna(bw_val) else 0.0
        t_cir_val = row.get('Traffic CIR(Kbps)', row.get('traffic_cir_kbps'))
        t_cir = float(t_cir_val) if pd.notna(t_cir_val) else 0.0
        t_eir_val = row.get('Traffic EIR(Kbps)', row.get('traffic_eir_kbps'))
        t_eir = float(t_eir_val) if pd.notna(t_eir_val) else 0.0
        n_cir_val = row.get('Network CIR(Kbps)', row.get('network_cir_kbps'))
        n_cir = float(n_cir_val) if pd.notna(n_cir_val) else 0.0
        n_eir_val = row.get('Network EIR(Kbps)', row.get('network_eir_kbps'))
        n_eir = float(n_eir_val) if pd.notna(n_eir_val) else 0.0

        cos = str(row.get('Cos', row.get('cos', ''))).strip() if pd.notna(row.get('Cos', row.get('cos'))) else ''
        qos = str(row.get('Trust CE QoS', row.get('trust_ce_qos', ''))).strip() if pd.notna(row.get('Trust CE QoS', row.get('trust_ce_qos'))) else ''
        order = str(row.get('Order Name', row.get('order_name', ''))).strip() if pd.notna(row.get('Order Name', row.get('order_name'))) else ''
        client = str(row.get('Client', row.get('client', ''))).strip() if pd.notna(row.get('Client', row.get('client'))) else ''
        a_end = str(row.get('A End', row.get('a_end', ''))).strip() if pd.notna(row.get('A End', row.get('a_end'))) else ''
        z_end = str(row.get('Z End', row.get('z_end', ''))).strip() if pd.notna(row.get('Z End', row.get('z_end'))) else ''
        c_time = str(row.get('Create Time', row.get('create_time', ''))).strip() if pd.notna(row.get('Create Time', row.get('create_time'))) else ''
        u_time = str(row.get('Update Time', row.get('update_time', ''))).strip() if pd.notna(row.get('Update Time', row.get('update_time'))) else ''
        desc = str(row.get('Description', row.get('description', ''))).strip() if pd.notna(row.get('Description', row.get('description'))) else ''

        records.append((
            s_idx, name, stype, bw, t_cir, t_eir, n_cir, n_eir,
            cos, qos, order, client, a_end, z_end, c_time, u_time, desc
        ))

    cursor.execute('DELETE FROM cpan_services;')
    cursor.executemany('''
        INSERT INTO cpan_services (
            service_index, name, service_type, bandwidth_kbps,
            traffic_cir_kbps, traffic_eir_kbps, network_cir_kbps, network_eir_kbps,
            cos, trust_ce_qos, order_name, client, a_end, z_end,
            create_time, update_time, description
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
    ''', records)

    conn.commit()
    count = len(records)
    print(f"Imported {count} CPAN services into cpan_services table.")
    if close_conn:
        conn.close()
    return count


def import_cpan_dl_list_csv(csv_source, conn=None):
    close_conn = False
    if conn is None:
        conn = get_db_connection()
        close_conn = True
    cursor = conn.cursor()

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS cpan_dl_list (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT,
            media_type TEXT,
            bandwidth TEXT,
            signal_type TEXT,
            direction TEXT,
            a_end TEXT,
            z_end TEXT,
            alarm_status TEXT,
            cir_utilization TEXT,
            bandwidth_utilization TEXT,
            order_name TEXT,
            creator TEXT,
            client TEXT,
            cost TEXT,
            create_time TEXT,
            update_time TEXT,
            description TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    ''')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_cpan_dl_name ON cpan_dl_list(name);')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_cpan_dl_creator ON cpan_dl_list(creator);')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_cpan_dl_order ON cpan_dl_list(order_name);')

    if isinstance(csv_source, str) and os.path.exists(csv_source):
        df = pd.read_csv(csv_source)
    elif isinstance(csv_source, pd.DataFrame):
        df = csv_source
    else:
        if close_conn:
            conn.close()
        return 0

    records = []
    for idx, row in df.iterrows():
        name = str(row.get('Name', row.get('name', ''))).strip() if pd.notna(row.get('Name', row.get('name'))) else ''
        if not name:
            continue

        media_type = str(row.get('Media Type', row.get('media_type', ''))).strip() if pd.notna(row.get('Media Type', row.get('media_type'))) else ''
        bandwidth = str(row.get('Bandwidth', row.get('bandwidth', ''))).strip() if pd.notna(row.get('Bandwidth', row.get('bandwidth'))) else ''
        signal_type = str(row.get('Signal Type', row.get('signal_type', ''))).strip() if pd.notna(row.get('Signal Type', row.get('signal_type'))) else ''
        direction = str(row.get('Direction', row.get('direction', ''))).strip() if pd.notna(row.get('Direction', row.get('direction'))) else ''
        a_end = str(row.get('A End', row.get('a_end', ''))).strip() if pd.notna(row.get('A End', row.get('a_end'))) else ''
        z_end = str(row.get('Z End', row.get('z_end', ''))).strip() if pd.notna(row.get('Z End', row.get('z_end'))) else ''
        alarm_status = str(row.get('Alarm Status', row.get('alarm_status', ''))).strip() if pd.notna(row.get('Alarm Status', row.get('alarm_status'))) else ''
        cir_util = str(row.get('CIR Utilization Ratio(%)', row.get('cir_utilization', ''))).strip() if pd.notna(row.get('CIR Utilization Ratio(%)', row.get('cir_utilization'))) else ''
        bw_util = str(row.get('Bandwidth Utilization Ratio(%)', row.get('bandwidth_utilization', ''))).strip() if pd.notna(row.get('Bandwidth Utilization Ratio(%)', row.get('bandwidth_utilization'))) else ''
        order_name = str(row.get('Order Name', row.get('order_name', ''))).strip() if pd.notna(row.get('Order Name', row.get('order_name'))) else ''
        creator = str(row.get('Creator', row.get('creator', ''))).strip() if pd.notna(row.get('Creator', row.get('creator'))) else ''
        client = str(row.get('Client', row.get('client', ''))).strip() if pd.notna(row.get('Client', row.get('client'))) else ''
        cost = str(row.get('Cost', row.get('cost', ''))).strip() if pd.notna(row.get('Cost', row.get('cost'))) else ''
        create_time = str(row.get('Create Time', row.get('create_time', ''))).strip() if pd.notna(row.get('Create Time', row.get('create_time'))) else ''
        update_time = str(row.get('Update Time', row.get('update_time', ''))).strip() if pd.notna(row.get('Update Time', row.get('update_time'))) else ''
        description = str(row.get('Description', row.get('description', ''))).strip() if pd.notna(row.get('Description', row.get('description'))) else ''

        records.append((
            name, media_type, bandwidth, signal_type, direction,
            a_end, z_end, alarm_status, cir_util, bw_util,
            order_name, creator, client, cost, create_time, update_time, description
        ))

    cursor.execute('DELETE FROM cpan_dl_list;')
    cursor.executemany('''
        INSERT INTO cpan_dl_list (
            name, media_type, bandwidth, signal_type, direction,
            a_end, z_end, alarm_status, cir_utilization, bandwidth_utilization,
            order_name, creator, client, cost, create_time, update_time, description
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
    ''', records)

    conn.commit()
    count = len(records)
    print(f"Imported {count} CPAN DL List records into cpan_dl_list table.")
    if close_conn:
        conn.close()
    return count


def get_row_val(row, col_names, default=""):
    for c in col_names:
        for key in row.index:
            if str(key).strip().lower() == c.strip().lower():
                val = row[key]
                if pd.notna(val) and val is not None:
                    s_val = str(val).strip()
                    if s_val.lower() not in ('nan', 'none', '<na>', 'null'):
                        return s_val
    return default


def create_maan_tables_schema(conn=None):
    close_conn = False
    if conn is None:
        conn = get_db_connection()
        close_conn = True
    cursor = conn.cursor()

    # Create MAAN Nodes Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS maan_nodes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ne_ip TEXT,
            location TEXT,
            type TEXT,
            ssa TEXT,
            phase TEXT,
            ne_name TEXT,
            dcc_ip TEXT,
            software_version TEXT,
            hardware_version TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    ''')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_maan_ne_ip ON maan_nodes(ne_ip);')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_maan_location ON maan_nodes(location);')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_maan_type ON maan_nodes(type);')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_maan_ssa ON maan_nodes(ssa);')

    # Create MAAN Services Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS maan_services (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            service_index INTEGER,
            name TEXT,
            service_type TEXT,
            bandwidth_kbps REAL DEFAULT 0.0,
            traffic_cir_kbps REAL DEFAULT 0.0,
            a_end TEXT,
            z_end TEXT,
            order_name TEXT,
            client TEXT,
            description TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    ''')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_maan_srv_name ON maan_services(name);')

    # Create MAAN TLS Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS maan_tls (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT,
            media_type TEXT,
            bandwidth TEXT,
            signal_type TEXT,
            direction TEXT,
            a_end TEXT,
            z_end TEXT,
            alarm_status TEXT,
            cir_utilization TEXT,
            bandwidth_utilization TEXT,
            order_name TEXT,
            creator TEXT,
            client TEXT,
            cost TEXT,
            create_time TEXT,
            update_time TEXT,
            description TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    ''')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_maan_tls_name ON maan_tls(name);')

    conn.commit()
    if close_conn:
        conn.close()


def init_maan_tables(conn=None):
    close_conn = False
    if conn is None:
        conn = get_db_connection()
        close_conn = True
    cursor = conn.cursor()

    create_maan_tables_schema(conn)

    # Seed MAAN Nodes from GJAHMPROV1_Node_Report0.csv or bts_sites if empty
    cursor.execute("SELECT COUNT(*) FROM maan_nodes;")
    if cursor.fetchone()[0] == 0:
        maan_csv = os.path.join(DATA_DIR, 'GJAHMPROV1_Node_Report0.csv') if os.path.exists(os.path.join(DATA_DIR, 'GJAHMPROV1_Node_Report0.csv')) else os.path.join(BASE_DIR, 'GJAHMPROV1_Node_Report0.csv')
        if os.path.exists(maan_csv):
            print(f"Importing MAAN Nodes from {maan_csv}...")
            import_maan_nodes_csv(maan_csv, conn)
        else:
            cursor.execute("""
                SELECT DISTINCT
                    COALESCE(NULLIF(mgmt_ip, ''), NULLIF(endpoint_ip, ''), NULLIF(enodeb_address, ''), '10.228.0.1') as ne_ip,
                    COALESCE(NULLIF(location, ''), NULLIF(site_name, ''), 'GUJARAT') as location,
                    COALESCE(NULLIF(endpoint_type, ''), NULLIF(endpoint_node_router, ''), 'MAAN-Router') as type,
                    COALESCE(NULLIF(ssa, ''), 'GJ_SSA') as ssa,
                    'PH1' as phase,
                    (site_id || '_' || COALESCE(site_name, '')) as ne_name,
                    COALESCE(NULLIF(mgmt_ip, ''), '') as dcc_ip
                FROM bts_sites
                WHERE LOWER(cpan_maan_vsat) LIKE '%maan%';
            """)
            m_nodes = cursor.fetchall()
            if m_nodes:
                cursor.executemany("""
                    INSERT INTO maan_nodes (ne_ip, location, type, ssa, phase, ne_name, dcc_ip)
                    VALUES (?, ?, ?, ?, ?, ?, ?);
                """, [tuple(r) for r in m_nodes])
                print(f"Seeded {len(m_nodes)} MAAN nodes into maan_nodes table.")

    # Seed MAAN Services from bts_sites if empty
    cursor.execute("SELECT COUNT(*) FROM maan_services;")
    if cursor.fetchone()[0] == 0:
        cursor.execute("""
            SELECT DISTINCT
                (site_id || ' - MAAN Service') as name,
                COALESCE(NULLIF(maan_vpn, ''), 'L3VPN') as service_type,
                100000.0 as bandwidth_kbps,
                50000.0 as traffic_cir_kbps,
                COALESCE(NULLIF(endpoint_node_router, ''), site_name) as a_end,
                COALESCE(NULLIF(location, ''), ssa) as z_end,
                COALESCE(NULLIF(site_id, ''), 'ORD_MAAN') as order_name,
                COALESCE(NULLIF(ssa, ''), 'BSNL_MAAN') as client
            FROM bts_sites
            WHERE LOWER(cpan_maan_vsat) LIKE '%maan%';
        """)
        m_srvs = cursor.fetchall()
        if m_srvs:
            cursor.executemany("""
                INSERT INTO maan_services (name, service_type, bandwidth_kbps, traffic_cir_kbps, a_end, z_end, order_name, client)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?);
            """, [tuple(r) for r in m_srvs])
            print(f"Seeded {len(m_srvs)} MAAN services into maan_services table.")

    # Seed MAAN TLS from bts_sites if empty
    cursor.execute("SELECT COUNT(*) FROM maan_tls;")
    if cursor.fetchone()[0] == 0:
        cursor.execute("""
            SELECT DISTINCT
                (site_id || ' - MAAN TLS Circuit') as name,
                'Fiber' as media_type,
                '100M' as bandwidth,
                'Eth' as signal_type,
                COALESCE(NULLIF(endpoint_node_router, ''), site_name) as a_end,
                COALESCE(NULLIF(location, ''), ssa) as z_end,
                'Normal' as alarm_status,
                COALESCE(NULLIF(site_id, ''), 'ORD_TLS') as order_name,
                COALESCE(NULLIF(ssa, ''), 'BSNL_TLS') as client
            FROM bts_sites
            WHERE LOWER(cpan_maan_vsat) LIKE '%maan%';
        """)
        m_tls = cursor.fetchall()
        if m_tls:
            cursor.executemany("""
                INSERT INTO maan_tls (name, media_type, bandwidth, signal_type, a_end, z_end, alarm_status, order_name, client)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?);
            """, [tuple(r) for r in m_tls])
            print(f"Seeded {len(m_tls)} MAAN TLS records into maan_tls table.")

    conn.commit()
    if close_conn:
        conn.close()


def import_maan_nodes_csv(csv_source, conn=None):
    close_conn = False
    if conn is None:
        conn = get_db_connection()
        close_conn = True
    cursor = conn.cursor()

    create_maan_tables_schema(conn)

    if isinstance(csv_source, str) and os.path.exists(csv_source):
        df = pd.read_csv(csv_source)
    elif isinstance(csv_source, pd.DataFrame):
        df = csv_source
    else:
        if close_conn:
            conn.close()
        return 0

    records = []
    for idx, row in df.iterrows():
        ne_ip = get_row_val(row, ['Node IP or Name', 'NE IP', 'ne_ip', 'ip'])
        location = get_row_val(row, ['Location', 'location', 'site_name', 'site'])
        ne_type = get_row_val(row, ['Product Name', 'Product Type', 'Type', 'type', 'hardware_type'])
        ssa = get_row_val(row, ['Partition Label', 'SSA', 'ssa', 'circle'])
        phase = get_row_val(row, ['Product code', 'Version', 'Phase', 'phase'])
        ne_name = get_row_val(row, ['Node Label', 'NE Name', 'ne_name', 'node_name'])
        dcc_ip = get_row_val(row, ['Ethernet IP', 'DCC IP', 'dcc_ip', 'dcc'])

        if not ne_ip and not location and not ne_name:
            continue

        records.append((ne_ip, location, ne_type, ssa, phase, ne_name, dcc_ip))

    cursor.execute('DELETE FROM maan_nodes;')
    cursor.executemany('''
        INSERT INTO maan_nodes (ne_ip, location, type, ssa, phase, ne_name, dcc_ip)
        VALUES (?, ?, ?, ?, ?, ?, ?);
    ''', records)
    conn.commit()
    count = len(records)
    print(f"Imported {count} MAAN nodes into maan_nodes table.")
    if close_conn:
        conn.close()
    return count


def import_maan_services_csv(csv_source, conn=None):
    close_conn = False
    if conn is None:
        conn = get_db_connection()
        close_conn = True
    cursor = conn.cursor()

    init_maan_tables(conn)

    if isinstance(csv_source, str) and os.path.exists(csv_source):
        df = pd.read_csv(csv_source)
    elif isinstance(csv_source, pd.DataFrame):
        df = csv_source
    else:
        if close_conn:
            conn.close()
        return 0

    records = []
    for idx, row in df.iterrows():
        name = str(row.get('Name', row.get('name', ''))).strip() if pd.notna(row.get('Name', row.get('name'))) else ''
        if not name:
            continue
        service_type = str(row.get('Service Type', row.get('service_type', ''))).strip() if pd.notna(row.get('Service Type', row.get('service_type'))) else ''
        bw = float(row.get('Bandwidth(Kbps)', row.get('bandwidth_kbps', 0))) if pd.notna(row.get('Bandwidth(Kbps)', row.get('bandwidth_kbps'))) else 0.0
        t_cir = float(row.get('Traffic CIR(Kbps)', row.get('traffic_cir_kbps', 0))) if pd.notna(row.get('Traffic CIR(Kbps)', row.get('traffic_cir_kbps'))) else 0.0
        order = str(row.get('Order Name', row.get('order_name', ''))).strip() if pd.notna(row.get('Order Name', row.get('order_name'))) else ''
        client = str(row.get('Client', row.get('client', ''))).strip() if pd.notna(row.get('Client', row.get('client'))) else ''
        a_end = str(row.get('A End', row.get('a_end', ''))).strip() if pd.notna(row.get('A End', row.get('a_end'))) else ''
        z_end = str(row.get('Z End', row.get('z_end', ''))).strip() if pd.notna(row.get('Z End', row.get('z_end'))) else ''
        desc = str(row.get('Description', row.get('description', ''))).strip() if pd.notna(row.get('Description', row.get('description'))) else ''
        records.append((idx + 1, name, service_type, bw, t_cir, order, client, a_end, z_end, desc))

    cursor.execute('DELETE FROM maan_services;')
    cursor.executemany('''
        INSERT INTO maan_services (service_index, name, service_type, bandwidth_kbps, traffic_cir_kbps, order_name, client, a_end, z_end, description)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
    ''', records)
    conn.commit()
    count = len(records)
    print(f"Imported {count} MAAN services into maan_services table.")
    if close_conn:
        conn.close()
    return count


def import_maan_tls_csv(csv_source, conn=None):
    close_conn = False
    if conn is None:
        conn = get_db_connection()
        close_conn = True
    cursor = conn.cursor()

    init_maan_tables(conn)

    if isinstance(csv_source, str) and os.path.exists(csv_source):
        df = pd.read_csv(csv_source)
    elif isinstance(csv_source, pd.DataFrame):
        df = csv_source
    else:
        if close_conn:
            conn.close()
        return 0

    records = []
    for idx, row in df.iterrows():
        name = str(row.get('Name', row.get('name', ''))).strip() if pd.notna(row.get('Name', row.get('name'))) else ''
        if not name:
            continue
        media_type = str(row.get('Media Type', row.get('media_type', ''))).strip() if pd.notna(row.get('Media Type', row.get('media_type'))) else ''
        bandwidth = str(row.get('Bandwidth', row.get('bandwidth', ''))).strip() if pd.notna(row.get('Bandwidth', row.get('bandwidth'))) else ''
        signal_type = str(row.get('Signal Type', row.get('signal_type', ''))).strip() if pd.notna(row.get('Signal Type', row.get('signal_type'))) else ''
        a_end = str(row.get('A End', row.get('a_end', ''))).strip() if pd.notna(row.get('A End', row.get('a_end'))) else ''
        z_end = str(row.get('Z End', row.get('z_end', ''))).strip() if pd.notna(row.get('Z End', row.get('z_end'))) else ''
        alarm_status = str(row.get('Alarm Status', row.get('alarm_status', ''))).strip() if pd.notna(row.get('Alarm Status', row.get('alarm_status'))) else ''
        order_name = str(row.get('Order Name', row.get('order_name', ''))).strip() if pd.notna(row.get('Order Name', row.get('order_name'))) else ''
        client = str(row.get('Client', row.get('client', ''))).strip() if pd.notna(row.get('Client', row.get('client'))) else ''
        desc = str(row.get('Description', row.get('description', ''))).strip() if pd.notna(row.get('Description', row.get('description'))) else ''
        records.append((name, media_type, bandwidth, signal_type, a_end, z_end, alarm_status, order_name, client, desc))

    cursor.execute('DELETE FROM maan_tls;')
    cursor.executemany('''
        INSERT INTO maan_tls (name, media_type, bandwidth, signal_type, a_end, z_end, alarm_status, order_name, client, description)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
    ''', records)
    conn.commit()
    count = len(records)
    print(f"Imported {count} MAAN TLS records into maan_tls table.")
    if close_conn:
        conn.close()
    return count


if __name__ == '__main__':
    init_db(force_reimport=True)



