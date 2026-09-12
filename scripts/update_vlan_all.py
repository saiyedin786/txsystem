import os
import gc
import sqlite3
import pandas as pd
import openpyxl

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
bts_path = os.path.join(BASE_DIR, 'btsdatabase_updated - v1.xlsx') if os.path.exists(os.path.join(BASE_DIR, 'btsdatabase_updated - v1.xlsx')) else os.path.join(BASE_DIR, 'btsdatabase.xlsx')
master_path = os.path.join(BASE_DIR, 'Master_Record_TCS.xlsx')
db_path = os.path.join(BASE_DIR, 'btsdatabase.db')
temp_path = os.path.join(BASE_DIR, 'btsdatabase_temp.xlsx')

print("Loading Master_Record_TCS.xlsx...")
if os.path.exists(master_path):
    with open(master_path, 'rb') as f_master:
        df_master = pd.read_excel(f_master)

    lookup_by_bts_id = {}
    lookup_by_bts_name = {}

    for idx, row in df_master.iterrows():
        bts_id = str(row['BTS ID']).strip().upper() if pd.notna(row['BTS ID']) else ""
        bts_name = str(row['BTS Name']).strip().upper() if pd.notna(row['BTS Name']) else ""
        
        vlan_val = ""
        for col in ['S1C B28 VLAN', 'S1C B1 VLAN', 'S1C B41 VLAN']:
            if col in row and pd.notna(row[col]) and str(row[col]).strip().lower() != 'nan':
                val_str = str(row[col]).strip()
                if val_str.endswith('.0'):
                    val_str = val_str[:-2]
                vlan_val = val_str
                break
                
        data = {
            'tx-system-ip': str(row.get('Endpoint Node IP', '')).strip() if pd.notna(row.get('Endpoint Node IP')) and str(row.get('Endpoint Node IP')).strip().lower() != 'nan' else "",
            'tx-system-location': str(row.get('Endpoint Node', '')).strip() if pd.notna(row.get('Endpoint Node')) and str(row.get('Endpoint Node')).strip().lower() != 'nan' else "",
            'tx-system-port': str(row.get('Endpoint Port', '')).strip() if pd.notna(row.get('Endpoint Port')) and str(row.get('Endpoint Port')).strip().lower() != 'nan' else "",
            'vlan': vlan_val
        }
        
        if bts_id and bts_id != 'NAN':
            lookup_by_bts_id[bts_id] = data
        if bts_name and bts_name != 'NAN':
            lookup_by_bts_name[bts_name] = data

    print(f"Loaded {len(lookup_by_bts_id)} BTS IDs from Master_Record_TCS.xlsx")

    # Update SQLite Database btsdatabase.db
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    cursor.execute("SELECT id, site_id, site_name, ssa FROM bts_sites;")
    rows = cursor.fetchall()

    updated_count = 0
    updates_list = []

    for r in rows:
        row_id = r['id']
        site_id_val = str(r['site_id'] or '').strip().upper()
        site_name_val = str(r['site_name'] or '').strip().upper()
        ssa_val = str(r['ssa'] or '').strip().upper()
        
        data = None
        if site_id_val in lookup_by_bts_id:
            data = lookup_by_bts_id[site_id_val]
        elif '_' in site_name_val and site_name_val.split('_', 1)[0] in lookup_by_bts_id:
            data = lookup_by_bts_id[site_name_val.split('_', 1)[0]]
        elif site_name_val in lookup_by_bts_name:
            data = lookup_by_bts_name[site_name_val]
        elif ssa_val in lookup_by_bts_id:
            data = lookup_by_bts_id[ssa_val]
        else:
            for b_id, b_data in lookup_by_bts_id.items():
                if b_id and (site_id_val.endswith(b_id) or site_name_val.startswith(b_id)):
                    data = b_data
                    break
                    
        if data:
            updates_list.append((data['tx-system-ip'], data['tx-system-location'], data['tx-system-port'], data['vlan'], row_id))
            updated_count += 1

    cursor.executemany('''
        UPDATE bts_sites SET
            tx_system_ip = ?,
            tx_system_location = ?,
            tx_system_port = ?,
            vlan = ?
        WHERE id = ?;
    ''', updates_list)

    conn.commit()
    conn.close()
    print(f"Successfully updated SQLite database with VLAN and TX details for {updated_count} / {len(rows)} records!")
else:
    print("Master_Record_TCS.xlsx not found. Skipping utility update.")
