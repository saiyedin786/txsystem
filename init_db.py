import os
import sqlite3
import pandas as pd

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'btsdatabase.db')
EXCEL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'btsdatabase.xlsx')


def get_db_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db(force_reimport=False):
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # 1. Create bts_sites table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS bts_sites (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            enodeb_address TEXT,
            site_id TEXT UNIQUE NOT NULL,
            site_name TEXT,
            ssa TEXT,
            location TEXT,
            cpan_maan_vsat TEXT,
            tx_system_ip TEXT,
            tx_system_location TEXT,
            tx_system_port TEXT,
            vlan TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    ''')
    
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_site_id ON bts_sites(site_id);')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_ssa ON bts_sites(ssa);')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_site_name ON bts_sites(site_name);')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_enodeb ON bts_sites(enodeb_address);')
    
    # 2. Create users table for authentication & profiles
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
    columns = [col[1] for col in cursor.fetchall()]
    if 'username' not in columns:
        cursor.execute("ALTER TABLE users ADD COLUMN username TEXT;")
        
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_staff_no ON users(staff_no);')
    cursor.execute('CREATE UNIQUE INDEX IF NOT EXISTS idx_username ON users(username);')
    
    # Remove default STAFF001 user if present
    cursor.execute("DELETE FROM users WHERE staff_no = 'STAFF001';")
    conn.commit()
    
    cursor.execute('SELECT COUNT(*) FROM users;')
    user_count = cursor.fetchone()[0]
    
    # Check bts_sites count
    cursor.execute('SELECT COUNT(*) FROM bts_sites;')
    count = cursor.fetchone()[0]
    
    if count == 0 or force_reimport:
        if os.path.exists(EXCEL_PATH):
            print(f"Migrating records from Excel ({EXCEL_PATH}) into SQLite ({DB_PATH})...")
            
            with open(EXCEL_PATH, 'rb') as f:
                df = pd.read_excel(f, sheet_name='Sheet1', dtype=str)
                
            df = df.fillna("")
            
            records_to_insert = []
            for _, row in df.iterrows():
                site_id = str(row.get('site_id', '')).strip()
                if not site_id or site_id.lower() == 'nan':
                    continue
                records_to_insert.append((
                    str(row.get('enodeb_address', '')).strip().replace('nan', ''),
                    site_id,
                    str(row.get('site_name', '')).strip().replace('nan', ''),
                    str(row.get('ssa', '')).strip().replace('nan', ''),
                    str(row.get('Location', '')).strip().replace('nan', ''),
                    str(row.get('cpan/maan/vsat', '')).strip().replace('nan', ''),
                    str(row.get('tx-system-ip', '')).strip().replace('nan', ''),
                    str(row.get('tx-system-location', '')).strip().replace('nan', ''),
                    str(row.get('tx-system-port', '')).strip().replace('nan', ''),
                    str(row.get('vlan', '')).strip().replace('nan', '')
                ))
                
            if force_reimport:
                cursor.execute('DELETE FROM bts_sites;')
                
            cursor.executemany('''
                INSERT OR REPLACE INTO bts_sites (
                    enodeb_address, site_id, site_name, ssa, location,
                    cpan_maan_vsat, tx_system_ip, tx_system_location, tx_system_port, vlan
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
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
    init_db()
