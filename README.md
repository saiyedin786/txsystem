# CPAN NOC

A web-based Network Operations Center (NOC) system built with **Flask**, **SQLite**, **Bootstrap 5**, and **Pandas** for managing, searching, uploading, and diagnosing CPAN, MAAN, and VSAT network transmission nodes and site database records.

---

## 📁 Project Directory Structure

```
txsystem/
├── data/                            # Database & Raw/Parsed Data Files
│   ├── btsdatabase.db               # SQLite Master Database
│   ├── GUJ_CPAN-NODE_LIST.csv       # CPAN Node List Source CSV
│   ├── cpandatabase.csv             # Structured/Parsed CPAN Database CSV
│   ├── CPAN_Service_List.csv        # CPAN Service Records CSV
│   ├── btsdatabase_updated - v1.xlsx# Primary Excel Source File
│   ├── btsdatabase.xlsx             # Legacy Excel Source File
│   └── btsdatabase_updated.xlsx     # Secondary Excel Source File
├── scripts/                         # Maintenance & Data Processing Scripts
│   ├── create_cpandatabase.py       # Standalone CPAN CSV Generator
│   ├── update_sqlite_transmission.py# Transmission DB Update Helper
│   ├── update_transmission_data.py  # Transmission Data Synchronizer
│   ├── update_transmission_data_fix.py
│   └── update_vlan_all.py           # VLAN Bulk Updater
├── templates/                       # Jinja2 HTML View Templates
│   ├── base.html                    # Base layout & global modal diagnostics
│   ├── index.html                   # Site Records Dashboard
│   ├── cpan_nodes.html              # CPAN Node List View & AJAX Search
│   ├── login.html                   # User Authentication Login
│   ├── profile.html                 # User Profile & Staff Settings
│   ├── register.html                # User Registration
│   ├── site_form.html               # Add / Edit Site Form
│   └── view_site.html               # Full Site Detail View
├── app.py                           # Main Flask Application Server & Routes
├── init_db.py                       # SQLite DB Initialization & Excel/CSV Migration
├── requirements.txt                 # Python Dependencies
└── README.md                        # Project Documentation
```

---

## 🚀 Key Features

1. **CPAN Node List Management (`/cpan-nodes`)**:
   - Automated parsing of `GUJ_CPAN-NODE_LIST.csv` into structured fields (`NE IP`, `Location`, `Type`, `SSA`, `Phase`, `NE Name`, `DCC IP`).
   - **Instant AJAX Search**: Live, debouncing multi-term search with zero page reloads.
   - **Full CRUD Support**: Add new nodes, View full node details, Edit node attributes, and Delete nodes.
   - **CSV/Excel Upload & Export**: Upload updated node lists and export filtered node data.

2. **Transmission Dashboard (`/`)**:
   - Site metrics overview (Total Sites, CPAN Sites, MAAN Sites, Filtered Sites, SSAs).
   - Multi-column search, SSA filter dropdown, and per-page controls.
   - Custom report generation and Excel export.

3. **Interactive Network Diagnostics**:
   - Click any IP button across the dashboard or CPAN nodes view to trigger an instant ping diagnostic modal terminal.

---

## ⚙️ Installation & Usage

1. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

2. **Initialize Database**:
   ```bash
   python init_db.py
   ```

3. **Start the Web Application**:
   ```bash
   python app.py
   ```

4. Access the web dashboard in your browser at `http://127.0.0.1:5000`.
