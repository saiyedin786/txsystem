import os
import re
import sqlite3
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, 'btsdatabase.db')

NEW_EXCEL_PATH = os.path.join(BASE_DIR, 'btsdatabase_updated - v1.xlsx')
OLD_EXCEL_PATH = os.path.join(BASE_DIR, 'btsdatabase.xlsx')

EXCEL_PATH = NEW_EXCEL_PATH if os.path.exists(NEW_EXCEL_PATH) else OLD_EXCEL_PATH


def get_db_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
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
        'tx_system_ip', 'tx_system_location', 'tx_system_port', 'vlan'
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
        
    conn.close()


if __name__ == '__main__':
    init_db(force_reimport=True)
