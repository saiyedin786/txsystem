import os
import re
import pandas as pd

base_dir = os.path.dirname(os.path.abspath(__file__))
input_file = os.path.join(base_dir, 'GUJ_CPAN-NODE_LIST.csv')
output_file = os.path.join(base_dir, 'cpandatabase.csv')

print(f"Reading {input_file}...")
df = pd.read_csv(input_file)

records = []
for idx, row in df.iterrows():
    ne_name_raw = row['NE Name']
    if pd.isna(ne_name_raw) or not str(ne_name_raw).strip():
        continue
        
    ne_name = str(ne_name_raw).strip()
    obj_fid = str(row.get('Object Fid', '')).strip() if pd.notna(row.get('Object Fid')) else ''
    
    ne_ip = ""
    location = ""
    ne_type = ""
    ssa = ""
    phase = ""
    
    parts = ne_name.split('_')
    if len(parts) == 6:
        ne_ip = parts[0].strip()
        location = parts[1].strip()
        ne_type = parts[2].strip()
        ssa = f"{parts[3].strip()}_{parts[4].strip()}"
        phase = parts[5].strip()
    else:
        # Check if Object Fid has 6 parts
        clean_fid = obj_fid.lstrip('\\').strip()
        fid_parts = clean_fid.split('_')
        if len(fid_parts) == 6:
            ne_ip = fid_parts[0].strip()
            location = fid_parts[1].strip()
            ne_type = fid_parts[2].strip()
            ssa = f"{fid_parts[3].strip()}_{fid_parts[4].strip()}"
            phase = fid_parts[5].strip()
        else:
            # Fallback regex parsing
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

    records.append({
        'ne_ip': ne_ip,
        'location': location,
        'type': ne_type,
        'ssa': ssa,
        'phase': phase,
        'ne_name': ne_name,
        'ne_id': str(row.get('NE ID', '')).lstrip('\\').strip() if pd.notna(row.get('NE ID')) else '',
        'dcc_ip': str(row.get('DCC IP', '')).strip() if pd.notna(row.get('DCC IP')) else '',
        'software_version': str(row.get('Software Version', '')).strip() if pd.notna(row.get('Software Version')) else '',
        'hardware_version': str(row.get('Hardware Version', '')).strip() if pd.notna(row.get('Hardware Version')) else '',
        'pcb_version': str(row.get('PCB Version', '')).strip() if pd.notna(row.get('PCB Version')) else '',
        'last_upload_time': str(row.get('Last Upload Time', '')).strip() if pd.notna(row.get('Last Upload Time')) else '',
        'orig_location': str(row.get('Location', '')).strip() if pd.notna(row.get('Location')) else ''
    })

out_df = pd.DataFrame(records)
out_df.to_csv(output_file, index=False)
print(f"Successfully created {output_file} with {len(out_df)} rows.")
