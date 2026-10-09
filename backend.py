"""
PT. NIFCO INDONESIA - MILKRUN DELIVERY CONTROL SYSTEM
Backend Engine: FastAPI + SQLite3 + OpenPyXL Bi-Directional Sync
Architectural Pattern: Option B (ACID SQLite Core with Automatic Excel Synchronization)
Role-Based Access Control: Admin Protected Endpoints
"""

import os
import sys
import re
import sqlite3
import shutil
import logging
import threading
import time as time_module
from datetime import datetime, date, time
from typing import Optional, List
from threading import Lock

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from fastapi import FastAPI, HTTPException, Header, Depends, UploadFile, File, status, Query, Body
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, FileResponse
from pydantic import BaseModel

# --- CONFIGURATION & LOGGING ---
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("MilkrunBackend")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "milkrun.db")
EXCEL_PATH = os.path.join(BASE_DIR, "database_milkrun.xlsx")
INDEX_HTML_PATH = os.path.join(BASE_DIR, "index.html")
PROJECT_HTML_PATH = os.path.join(BASE_DIR, "Project Dasboard.html")
HTML_FILE_PATH = INDEX_HTML_PATH if os.path.exists(INDEX_HTML_PATH) else PROJECT_HTML_PATH

ADMIN_PIN = os.getenv("ADMIN_PIN", "1234")
ADMIN_SECRET_TOKEN = "nifco-admin-session-token-998822"

excel_lock = threading.RLock()

# --- FASTAPI APP ---
app = FastAPI(
    title="Nifco Milkrun Logistics API",
    description="Backend API with SQLite Core & Excel Auto-Sync for PT. Nifco Indonesia",
    version="2.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

# --- PYDANTIC MODELS ---
class LoginRequest(BaseModel):
    pin: str

class DriverModel(BaseModel):
    name: str
    phone: str
    partner: str
    status: Optional[str] = "AKTIF"

class ScheduleModel(BaseModel):
    id: Optional[str] = None
    shipdate: str
    customer: str
    logPartner: str
    cycle: int
    driver: Optional[str] = ""
    phone: Optional[str] = ""
    planTime: str
    actTime: Optional[str] = ""
    status: Optional[str] = "PENDING"
    customRemark: Optional[str] = ""
    # Separated logistics columns
    tgl_sj: Optional[str] = None
    jml_sj: Optional[str] = None
    pallet: Optional[str] = None
    box: Optional[str] = None
    return_pallet: Optional[str] = None
    tglSj: Optional[str] = None
    jmlSj: Optional[str] = None
    palletQty: Optional[str] = None
    boxQty: Optional[str] = None
    returnPallet: Optional[str] = None
    # Legacy fields
    sjDateQty: Optional[str] = None
    palletBox: Optional[str] = None
    sj_date_qty: Optional[str] = None
    pallet_box: Optional[str] = None

class DriverUpdateModel(BaseModel):
    name: str
    phone: str
    partner: str
    status: Optional[str] = "AKTIF" 

class LogisticsUpdateModel(BaseModel):
    tgl_sj: Optional[str] = None
    jml_sj: Optional[str] = None
    pallet: Optional[str] = None
    box: Optional[str] = None
    return_pallet: Optional[str] = None
    tglSj: Optional[str] = None
    jmlSj: Optional[str] = None
    palletQty: Optional[str] = None
    boxQty: Optional[str] = None
    returnPallet: Optional[str] = None
    # Legacy fields
    sj_date_qty: Optional[str] = None
    pallet_box: Optional[str] = None
    sjDateQty: Optional[str] = None
    palletBox: Optional[str] = None

class DriverArrivalCheckIn(BaseModel):
    scheduleId: Optional[str] = None
    trip_id: Optional[str] = None
    driverName: Optional[str] = ""
    actTime: Optional[str] = None
    act_time: Optional[str] = None
    latitude: Optional[float] = None
    longitude: Optional[float] = None
    customNote: Optional[str] = ""
    validation_note: Optional[str] = ""


# --- DATABASE INITIALIZATION (SQLite) ---
def get_db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with get_db() as conn:
        cursor = conn.cursor()
        
        # Schedules table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS schedules (
                id TEXT PRIMARY KEY,
                shipdate TEXT NOT NULL,
                customer TEXT NOT NULL,
                log_partner TEXT NOT NULL,
                cycle INTEGER NOT NULL,
                driver TEXT NOT NULL,
                phone TEXT,
                plan_time TEXT NOT NULL,
                act_time TEXT,
                status TEXT NOT NULL,
                custom_remark TEXT,
                sj_date_qty TEXT DEFAULT '',
                pallet_box TEXT DEFAULT '',
                return_pallet TEXT DEFAULT '',
                updated_at TEXT
            )
        """)

        # Migration: Ensure logistics columns exist in schedules
        cursor.execute("PRAGMA table_info(schedules)")
        s_cols = [c[1] for c in cursor.fetchall()]
        if "tgl_sj" not in s_cols:
            cursor.execute("ALTER TABLE schedules ADD COLUMN tgl_sj TEXT DEFAULT ''")
        if "jml_sj" not in s_cols:
            cursor.execute("ALTER TABLE schedules ADD COLUMN jml_sj TEXT DEFAULT ''")
        if "pallet" not in s_cols:
            cursor.execute("ALTER TABLE schedules ADD COLUMN pallet TEXT DEFAULT ''")
        if "box" not in s_cols:
            cursor.execute("ALTER TABLE schedules ADD COLUMN box TEXT DEFAULT ''")
        if "sj_date_qty" not in s_cols:
            cursor.execute("ALTER TABLE schedules ADD COLUMN sj_date_qty TEXT DEFAULT ''")
        if "pallet_box" not in s_cols:
            cursor.execute("ALTER TABLE schedules ADD COLUMN pallet_box TEXT DEFAULT ''")
        if "return_pallet" not in s_cols:
            cursor.execute("ALTER TABLE schedules ADD COLUMN return_pallet TEXT DEFAULT ''")

        # Auto-migrate compound fields to separate columns if empty
        cursor.execute("SELECT id, sj_date_qty, pallet_box, tgl_sj, jml_sj, pallet, box FROM schedules")
        rows_to_mig = cursor.fetchall()
        for r in rows_to_mig:
            up_m = []
            p_m = []
            if (not r["tgl_sj"] or not r["jml_sj"]) and r["sj_date_qty"]:
                parts = [p.strip() for p in r["sj_date_qty"].split("/")]
                if len(parts) >= 3:
                    up_m.extend(["tgl_sj = ?", "jml_sj = ?"])
                    p_m.extend([f"{parts[0]}/{parts[1]}", parts[2]])
                elif len(parts) == 2:
                    up_m.extend(["tgl_sj = ?", "jml_sj = ?"])
                    p_m.extend([parts[0], parts[1]])
                elif len(parts) == 1:
                    up_m.append("tgl_sj = ?")
                    p_m.append(parts[0])
            if (not r["pallet"] or not r["box"]) and r["pallet_box"]:
                parts = [p.strip() for p in r["pallet_box"].split("/")]
                if len(parts) >= 2:
                    up_m.extend(["pallet = ?", "box = ?"])
                    p_m.extend([parts[0], parts[1]])
                elif len(parts) == 1:
                    up_m.append("pallet = ?")
                    p_m.append(parts[0])
            if up_m:
                p_m.append(r["id"])
                cursor.execute(f"UPDATE schedules SET {', '.join(up_m)} WHERE id = ?", tuple(p_m))
        
        # Drivers master table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS drivers (
                name TEXT PRIMARY KEY,
                phone TEXT NOT NULL,
                partner TEXT NOT NULL,
                status TEXT DEFAULT 'AKTIF',
                last_active TEXT,
                updated_at TEXT
            )
        """)
        
        # Migration: Ensure last_active and status columns exist
        cursor.execute("PRAGMA table_info(drivers)")
        d_cols = [c[1] for c in cursor.fetchall()]
        if "last_active" not in d_cols:
            cursor.execute("ALTER TABLE drivers ADD COLUMN last_active TEXT")
            logger.info("Column 'last_active' added to drivers table.")
        if "status" not in d_cols:
            cursor.execute("ALTER TABLE drivers ADD COLUMN status TEXT DEFAULT 'AKTIF'")
            logger.info("Column 'status' added to drivers table.")

        # Settings table for Geofence & Webhooks
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL,
                updated_at TEXT
            )
        """)

        conn.commit()
        
        # Populate initial seeds if empty
        cursor.execute("SELECT COUNT(*) FROM drivers")
        if cursor.fetchone()[0] == 0:
            initial_drivers = [
                ("ASEP", "081234567890", "DARSINDO"),
                ("POSTER", "081298765432", "SYNCRUM"),
                ("YADI", "081566778899", "TTNT"),
                ("DEDI", "081788990011", "SELOG"),
                ("NAPIN", "081922334455", "SYNCRUM")
            ]
            cursor.executemany("INSERT INTO drivers (name, phone, partner, last_active, updated_at) VALUES (?, ?, ?, ?, ?)",
                               [(d[0], d[1], d[2], datetime.now().strftime("%Y-%m-%d %H:%M:%S"), datetime.now().isoformat()) for d in initial_drivers])
            conn.commit()

        cursor.execute("SELECT COUNT(*) FROM schedules")
        if cursor.fetchone()[0] == 0:
            today_str = datetime.now().strftime("%Y-%m-%d")
            initial_schedules = [
                ("MR-101", today_str, "TMMIN S2", "SYNCRUM", 1, "POSTER", "081298765432", "06:15", "08:30", "DELAY", ""),
                ("MR-102", today_str, "MMKI 1", "DARSINDO", 1, "ASEP", "081234567890", "07:00", "06:50", "ADVANCE", ""),
                ("MR-103", today_str, "YIMMCK", "TTNT", 2, "YADI", "081566778899", "03:30", "03:40", "ONTIME", "Warning! Barang belum tersedia, driver menunggu!"),
                ("MR-104", today_str, "AHM KRW", "SELOG", 1, "", "", "07:25", "", "PENDING", ""),
                ("MR-105", today_str, "TMMIN P2/1", "SYNCRUM", 3, "", "", "06:05", "", "PENDING", "Persiapan muat barang")
            ]
            cursor.executemany("""
                INSERT INTO schedules (id, shipdate, customer, log_partner, cycle, driver, phone, plan_time, act_time, status, custom_remark, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, [(s[0], s[1], s[2], s[3], s[4], s[5], s[6], s[7], s[8], s[9], s[10], datetime.now().isoformat()) for s in initial_schedules])
            conn.commit()

    logger.info("SQLite database initialized successfully.")
    sync_sqlite_to_excel()



# --- SMART EXCEL EXTRACTION & NORMALIZATION ENGINE ---
FIELD_KEYWORDS = {
    "id": ["id", "mr id", "kode", "no trip", "trip id", "schedule id", "code", "no.", "no trip milkrun"],
    "shipdate": ["shipdate", "ship date", "tanggal", "tgl", "date", "delivery date", "tgl kirim", "waktu kirim", "tgl delivery"],
    "customer": ["customer", "cust", "destinasi", "destination", "tujuan", "plant", "penerima", "customer name", "nama customer", "tujuan pabrik"],
    "log_partner": ["log_partner", "log partner", "logistics partner", "partner", "vendor", "transporter", "ekspedisi", "3pl", "armada", "perusahaan vendor"],
    "cycle": ["cycle", "siklus", "rit", "ritase", "trip", "round", "ke", "cycle ke"],
    "driver": ["driver", "driver name", "nama driver", "pengemudi", "supir", "nama supir"],
    "phone": ["phone", "telp", "no telp", "no hp", "driver phone", "kontak", "whatsapp", "wa", "no handphone", "kontak wa"],
    "plan_time": ["plan_time", "plan time", "plan", "jam plan", "rencana", "jam rencana", "eta", "waktu rencana", "plan arrival", "jam tiba", "waktu plan"],
    "act_time": ["act_time", "actual time", "act", "actual", "jam aktual", "realisasi", "waktu aktual", "ata", "jam datang"],
    "status": ["status", "kondisi", "state", "keterangan waktu"],
    "tgl_sj": ["tgl_sj", "tgl sj", "tanggal sj", "tanggal surat jalan", "tgl surat"],
    "jml_sj": ["jml_sj", "jml sj", "jumlah sj", "jumlah surat jalan", "jml surat", "qty sj", "surat jalan qty"],
    "pallet": ["pallet", "jumlah pallet", "jml pallet", "palet", "total pallet"],
    "box": ["box", "jumlah box", "jml box", "total box"],
    "sj_date_qty": ["sj_date_qty", "sj", "surat jalan", "tgl sj / jml", "sj / qty", "surat", "tgl surat jalan"],
    "pallet_box": ["pallet_box", "pallet / box", "jml pallet / box", "box / pallet", "box/pallet"],
    "return_pallet": ["return_pallet", "return", "retur", "kembali", "return pallet", "jumlah return pallet", "retur pallet", "palet kembali", "ret pallet"],
    "custom_remark": ["custom_remark", "custom remark", "remarks", "remark", "catatan", "keterangan", "note", "notes", "pesan", "catatan khusus"]
}

DRIVER_KEYWORDS = {
    "name": ["name", "nama", "nama driver", "driver", "pengemudi", "supir"],
    "phone": ["phone", "telepon", "telp", "no hp", "kontak", "wa", "whatsapp", "no handphone"],
    "partner": ["partner", "vendor", "logistics partner", "ekspedisi", "3pl", "perusahaan vendor"]
}

def clean_header_str(val):
    if val is None:
        return ""
    s = str(val).lower().strip()
    return re.sub(r"[\r\n\t]+", " ", s)

def parse_date_cell(val):
    if val is None:
        return datetime.now().strftime("%Y-%m-%d")
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d")
    s = str(val).strip()
    if not s:
        return datetime.now().strftime("%Y-%m-%d")
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d", "%d.%m.%Y", "%Y.%m.%d"):
        try:
            return datetime.strptime(s.split()[0], fmt).strftime("%Y-%m-%d")
        except Exception:
            pass
    return datetime.now().strftime("%Y-%m-%d")

def parse_time_cell(val):
    if val is None:
        return ""
    if isinstance(val, time):
        return val.strftime("%H:%M")
    if isinstance(val, datetime):
        return val.strftime("%H:%M")
    if isinstance(val, (int, float)):
        try:
            total_minutes = int(round(val * 24 * 60)) % 1440
            h, m = divmod(total_minutes, 60)
            return f"{h:02d}:{m:02d}"
        except Exception:
            return ""
    s = str(val).strip()
    if not s or s == "--:--":
        return ""
    m = re.search(r"(\d{1,2})[:.](\d{2})", s)
    if m:
        try:
            h = int(m.group(1))
            minute = int(m.group(2))
            if 0 <= h <= 23 and 0 <= minute <= 59:
                return f"{h:02d}:{minute:02d}"
        except Exception:
            pass
    return ""

def parse_cycle_cell(val):
    if val is None:
        return 1
    if isinstance(val, int):
        return max(1, val)
    if isinstance(val, float):
        return max(1, int(val))
    s = str(val).strip()
    m = re.search(r"\d+", s)
    if m:
        try:
            return max(1, int(m.group(0)))
        except Exception:
            return 1
    return 1

def calculate_status(plan_time, act_time):
    if not act_time or act_time.strip() == "" or act_time == "--:--":
        return "PENDING"
    try:
        ph, pm = map(int, plan_time.split(":"))
        ah, am = map(int, act_time.split(":"))
        diff = (ah * 60 + am) - (ph * 60 + pm)
        if diff < -720:
            diff += 1440
        elif diff > 720:
            diff -= 1440
        if diff < -5:
            return "ADVANCE"
        elif -5 <= diff <= 30:
            return "ONTIME"
        else:
            return "DELAY"
    except Exception:
        return "PENDING"

def analyze_and_extract_workbook(wb) -> dict:
    """
    Intelligently analyzes any uploaded Excel workbook:
    - Automatically discovers the schedule sheet even if named differently.
    - Scans top 15 rows to find the actual table header row (skipping banner/title lines).
    - Maps columns based on synonyms & keywords required by PT. Nifco Indonesia dashboard.
    - Discards unneeded external columns (e.g. truck capacity, costs, license plates).
    - Extracts driver master data if a driver sheet is present.
    """
    extracted_schedules = []
    extracted_drivers = []
    analysis_meta = {
        "schedule_sheet": None,
        "schedule_header_row": None,
        "schedule_columns_mapped": {},
        "driver_sheet": None,
        "driver_columns_mapped": {},
        "ignored_columns": []
    }

    best_sched_sheet = None
    best_sched_score = 0
    best_sched_header_row = 1
    best_sched_mapping = {}

    driver_sheet = None
    driver_header_row = 1
    driver_mapping = {}

    for name in wb.sheetnames:
        ws = wb[name]
        is_explicit_driver = "driver" in name.lower() or "pengemudi" in name.lower() or "supir" in name.lower()
        is_explicit_sched = "sched" in name.lower() or "jadwal" in name.lower() or "milkrun" in name.lower()

        rows_sample = list(ws.iter_rows(min_row=1, max_row=min(ws.max_row or 15, 15), values_only=True))
        for r_idx, row in enumerate(rows_sample):
            if not row or all(c is None or str(c).strip() == "" for c in row):
                continue

            col_map = {}
            score = 0
            for c_idx, cell in enumerate(row):
                header_text = clean_header_str(cell)
                if not header_text:
                    continue
                for field, keywords in FIELD_KEYWORDS.items():
                    if field in col_map:
                        continue
                    for kw in keywords:
                        if kw == header_text or (len(kw) > 3 and kw in header_text):
                            col_map[field] = c_idx
                            score += (2 if kw == header_text else 1)
                            break

            if "customer" in col_map: score += 3
            if "plan_time" in col_map: score += 3
            if "log_partner" in col_map: score += 1
            if is_explicit_sched: score += 2

            if score > best_sched_score and score >= 4:
                best_sched_score = score
                best_sched_sheet = name
                best_sched_header_row = r_idx + 1
                best_sched_mapping = col_map

            if not driver_sheet or is_explicit_driver:
                d_map = {}
                d_score = 0
                for c_idx, cell in enumerate(row):
                    header_text = clean_header_str(cell)
                    if not header_text:
                        continue
                    for d_field, keywords in DRIVER_KEYWORDS.items():
                        if d_field in d_map:
                            continue
                        for kw in keywords:
                            if kw == header_text or (len(kw) > 3 and kw in header_text):
                                d_map[d_field] = c_idx
                                d_score += 1
                                break
                if "name" in d_map and (is_explicit_driver or d_score >= 2):
                    driver_sheet = name
                    driver_header_row = r_idx + 1
                    driver_mapping = d_map

    # Extract Drivers
    if driver_sheet and driver_sheet != best_sched_sheet:
        ws_d = wb[driver_sheet]
        analysis_meta["driver_sheet"] = driver_sheet
        analysis_meta["driver_columns_mapped"] = driver_mapping
        for row in ws_d.iter_rows(min_row=driver_header_row + 1, values_only=True):
            if not row or all(c is None or str(c).strip() == "" for c in row):
                continue
            name_idx = driver_mapping.get("name")
            if name_idx is None or name_idx >= len(row) or not row[name_idx]:
                continue
            d_name = str(row[name_idx]).strip().upper()
            if not d_name or d_name in ("NAMA DRIVER", "NAME", "NAMA"):
                continue

            phone_idx = driver_mapping.get("phone")
            d_phone = str(row[phone_idx]).strip() if phone_idx is not None and phone_idx < len(row) and row[phone_idx] else ""
            d_phone = re.sub(r"[^\d+]", "", d_phone)

            partner_idx = driver_mapping.get("partner")
            d_partner = str(row[partner_idx]).strip().upper() if partner_idx is not None and partner_idx < len(row) and row[partner_idx] else "-"

            extracted_drivers.append({
                "name": d_name,
                "phone": d_phone,
                "partner": d_partner
            })

    # Extract Schedules
    if best_sched_sheet:
        ws_s = wb[best_sched_sheet]
        analysis_meta["schedule_sheet"] = best_sched_sheet
        analysis_meta["schedule_header_row"] = best_sched_header_row
        analysis_meta["schedule_columns_mapped"] = best_sched_mapping

        header_row_vals = list(ws_s.iter_rows(min_row=best_sched_header_row, max_row=best_sched_header_row, values_only=True))[0]
        mapped_indices = set(best_sched_mapping.values())
        for idx, col_val in enumerate(header_row_vals):
            if idx not in mapped_indices and col_val:
                analysis_meta["ignored_columns"].append(str(col_val).strip())

        counter = 101
        for row in ws_s.iter_rows(min_row=best_sched_header_row + 1, values_only=True):
            if not row or all(c is None or str(c).strip() == "" for c in row):
                continue

            def get_col(f):
                idx = best_sched_mapping.get(f)
                return row[idx] if idx is not None and idx < len(row) else None

            customer_raw = get_col("customer")
            plan_raw = get_col("plan_time")
            driver_raw = get_col("driver")

            if not customer_raw and not plan_raw and not driver_raw:
                continue

            customer = str(customer_raw).strip() if customer_raw else "N/A"
            if customer.lower() in ("total", "summary", "grand total", "keterangan"):
                continue

            id_raw = get_col("id")
            if id_raw and str(id_raw).strip():
                clean_id = str(id_raw).strip().upper()
                if not clean_id.startswith("MR-"):
                    clean_id = f"MR-{clean_id}"
            else:
                clean_id = f"MR-{counter}"
                counter += 1

            shipdate = parse_date_cell(get_col("shipdate"))
            partner = str(get_col("log_partner")).strip().upper() if get_col("log_partner") else "SELOG"
            cycle = parse_cycle_cell(get_col("cycle"))
            driver = str(driver_raw).strip().upper() if driver_raw else ""
            phone = str(get_col("phone")).strip() if get_col("phone") else ""
            phone = re.sub(r"[^\d+]", "", phone)

            plan_time = parse_time_cell(plan_raw) or "09:00"
            act_time = parse_time_cell(get_col("act_time"))

            status_raw = str(get_col("status")).strip().upper() if get_col("status") else ""
            status = calculate_status(plan_time, act_time) if not status_raw or status_raw not in ("PENDING", "ONTIME", "DELAY", "ADVANCE") else status_raw

            tgl_sj_val = str(get_col("tgl_sj")).strip() if get_col("tgl_sj") is not None else ""
            jml_sj_val = str(get_col("jml_sj")).strip() if get_col("jml_sj") is not None else ""
            pallet_val = str(get_col("pallet")).strip() if get_col("pallet") is not None else ""
            box_val = str(get_col("box")).strip() if get_col("box") is not None else ""
            sj_val = str(get_col("sj_date_qty")).strip() if get_col("sj_date_qty") is not None else ""
            pb_val = str(get_col("pallet_box")).strip() if get_col("pallet_box") is not None else ""
            rp_val = str(get_col("return_pallet")).strip() if get_col("return_pallet") is not None else ""
            remarks = str(get_col("custom_remark")).strip() if get_col("custom_remark") is not None else ""

            # Decompose compound fields if separate columns are empty
            if (not tgl_sj_val or not jml_sj_val) and sj_val:
                parts = [p.strip() for p in sj_val.split("/")]
                if len(parts) >= 3 and not tgl_sj_val and not jml_sj_val:
                    tgl_sj_val = f"{parts[0]}/{parts[1]}"
                    jml_sj_val = parts[2]
                elif len(parts) == 2 and not tgl_sj_val and not jml_sj_val:
                    tgl_sj_val = parts[0]
                    jml_sj_val = parts[1]
                elif len(parts) == 1 and not tgl_sj_val:
                    tgl_sj_val = parts[0]
            if (not pallet_val or not box_val) and pb_val:
                parts = [p.strip() for p in pb_val.split("/")]
                if len(parts) >= 2 and not pallet_val and not box_val:
                    pallet_val = parts[0]
                    box_val = parts[1]
                elif len(parts) == 1 and not pallet_val:
                    pallet_val = parts[0]

            # Rebuild compound values if missing
            if not sj_val and (tgl_sj_val or jml_sj_val):
                sj_val = f"{tgl_sj_val} / {jml_sj_val}".strip(" /")
            if not pb_val and (pallet_val or box_val):
                pb_val = f"{pallet_val} / {box_val}".strip(" /")

            extracted_schedules.append({
                "id": clean_id,
                "shipdate": shipdate,
                "customer": customer,
                "log_partner": partner,
                "cycle": cycle,
                "driver": driver,
                "phone": phone,
                "plan_time": plan_time,
                "act_time": act_time,
                "status": status,
                "custom_remark": remarks,
                "tgl_sj": tgl_sj_val,
                "jml_sj": jml_sj_val,
                "pallet": pallet_val,
                "box": box_val,
                "sj_date_qty": sj_val,
                "pallet_box": pb_val,
                "return_pallet": rp_val
            })

    return {
        "schedules": extracted_schedules,
        "drivers": extracted_drivers,
        "meta": analysis_meta
    }

# --- EXCEL BI-DIRECTIONAL ENGINE (OpenPyXL) ---
def sync_sqlite_to_excel() -> bool:
    """Exports SQLite data to a professionally formatted Excel workbook."""
    with excel_lock:
        try:
            wb = openpyxl.Workbook()

            # Styles
            header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
            orange_fill = PatternFill(start_color="EA580C", end_color="EA580C", fill_type="solid") # Corporate Nifco Orange
            teal_fill = PatternFill(start_color="0F766E", end_color="0F766E", fill_type="solid")
            regular_font = Font(name="Segoe UI", size=10)
            center_align = Alignment(horizontal="center", vertical="center")
            left_align = Alignment(horizontal="left", vertical="center")
            thin_border = Border(
                left=Side(style="thin", color="E2E8F0"),
                right=Side(style="thin", color="E2E8F0"),
                top=Side(style="thin", color="E2E8F0"),
                bottom=Side(style="thin", color="E2E8F0")
            )

            # Status fills
            status_fills = {
                "ONTIME": PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid"),
                "DELAY": PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid"),
                "ADVANCE": PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid"),
                "PENDING": PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
            }

            # 1. Sheet Schedules
            ws_sched = wb.active
            ws_sched.title = "Schedules"
            sched_headers = [
                "ID", "Ship Date", "Customer", "Logistics Partner", "Cycle",
                "Driver Name", "Driver Phone", "Plan Time", "Actual Time", "Status",
                "Tanggal SJ", "Jumlah SJ", "Pallet", "Box", "Jumlah Return Pallet",
                "Remarks", "Last Updated"
            ]
            ws_sched.append(sched_headers)

            for col_idx in range(1, len(sched_headers) + 1):
                cell = ws_sched.cell(row=1, column=col_idx)
                cell.font = header_font
                cell.fill = orange_fill
                cell.alignment = center_align

            with get_db() as conn:
                cursor = conn.cursor()
                cursor.execute("""
                    SELECT id, shipdate, customer, log_partner, cycle, driver, phone,
                           plan_time, act_time, status, tgl_sj, jml_sj, pallet, box, return_pallet,
                           custom_remark, updated_at
                    FROM schedules ORDER BY shipdate DESC, plan_time ASC
                """)
                rows = cursor.fetchall()

                for row_idx, r in enumerate(rows, start=2):
                    row_data = [
                        r["id"], r["shipdate"], r["customer"], r["log_partner"], r["cycle"],
                        r["driver"], r["phone"] or "-", r["plan_time"], r["act_time"] or "-",
                        r["status"], r["tgl_sj"] or "-", r["jml_sj"] or "-", r["pallet"] or "-", r["box"] or "-", r["return_pallet"] or "-",
                        r["custom_remark"] or "-", r["updated_at"] or "-"
                    ]
                    ws_sched.append(row_data)

                    for col_idx in range(1, len(row_data) + 1):
                        cell = ws_sched.cell(row=row_idx, column=col_idx)
                        cell.font = regular_font
                        cell.border = thin_border
                        if col_idx in [1, 2, 5, 8, 9, 10, 11, 12, 13, 14, 15]:
                            cell.alignment = center_align
                        else:
                            cell.alignment = left_align

                        # Highlight status column
                        if col_idx == 10 and r["status"] in status_fills:
                            cell.fill = status_fills[r["status"]]
                            cell.font = Font(name="Segoe UI", size=10, bold=True)

            # 2. Sheet Drivers_Master
            ws_drivers = wb.create_sheet(title="Drivers_Master")
            driver_headers = [
                "Driver Name", "WhatsApp / Phone", "Logistics Partner",
                "Terakhir Aktif (Last Active)", "Status Driver", "Status Rute", "Total Pengiriman", "Registered At"
            ]
            ws_drivers.append(driver_headers)

            for col_idx in range(1, len(driver_headers) + 1):
                cell = ws_drivers.cell(row=1, column=col_idx)
                cell.font = header_font
                cell.fill = teal_fill
                cell.alignment = center_align

            with get_db() as conn:
                cursor = conn.cursor()
                cursor.execute("""
                    SELECT d.name, d.phone, d.partner, COALESCE(d.status, 'AKTIF') as status, d.last_active, d.updated_at,
                           COUNT(CASE WHEN s.status = 'PENDING' THEN 1 END) as active_routes,
                           COUNT(s.id) as total_routes,
                           MAX(s.shipdate || ' ' || COALESCE(NULLIF(s.act_time, ''), s.plan_time)) as last_trip_time
                    FROM drivers d
                    LEFT JOIN schedules s ON s.driver = d.name
                    GROUP BY d.name, d.phone, d.partner, d.status, d.last_active, d.updated_at
                    ORDER BY d.name ASC
                """)
                d_rows = cursor.fetchall()

                for row_idx, r in enumerate(d_rows, start=2):
                    last_act = r["last_active"] or r["last_trip_time"] or r["updated_at"] or "-"
                    status_driver = (r["status"] or "AKTIF").upper()
                    status_rute = "RUTE AKTIF" if (r["active_routes"] or 0) > 0 else "STANDBY / NON-AKTIF"
                    row_data = [
                        r["name"],
                        r["phone"],
                        r["partner"],
                        last_act,
                        status_driver,
                        status_rute,
                        r["total_routes"] or 0,
                        r["updated_at"] or "-"
                    ]
                    ws_drivers.append(row_data)

                    for col_idx in range(1, len(row_data) + 1):
                        cell = ws_drivers.cell(row=row_idx, column=col_idx)
                        cell.font = regular_font
                        cell.border = thin_border
                        if col_idx in [1, 2, 4, 5, 6, 7]:
                            cell.alignment = center_align
                        else:
                            cell.alignment = left_align

                        if col_idx == 5: # Status Driver
                            if status_driver == "AKTIF":
                                cell.fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
                                cell.font = Font(name="Segoe UI", size=10, bold=True, color="166534")
                            else:
                                cell.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
                                cell.font = Font(name="Segoe UI", size=10, color="64748B")
                        elif col_idx == 6: # Status Rute
                            if status_rute == "RUTE AKTIF":
                                cell.fill = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")
                                cell.font = Font(name="Segoe UI", size=10, bold=True, color="0369A1")
                            else:
                                cell.fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
                                cell.font = Font(name="Segoe UI", size=10, color="94A3B8")

            # 3. Sheet History_Archive (if records exist)
            all_sheets_to_fit = [ws_sched, ws_drivers]
            try:
                with get_db() as conn:
                    cursor = conn.cursor()
                    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='schedules_history'")
                    if cursor.fetchone():
                        cursor.execute("SELECT COUNT(*) FROM schedules_history")
                        hist_count = cursor.fetchone()[0]
                        if hist_count > 0:
                            ws_hist = wb.create_sheet(title="History_Archive")
                            hist_headers = [
                                "ID", "Ship Date", "Customer", "Logistics Partner", "Cycle",
                                "Driver Name", "Driver Phone", "Plan Time", "Actual Time", "Status",
                                "Tanggal SJ", "Jumlah SJ", "Pallet", "Box", "Jumlah Return Pallet",
                                "Remarks", "Archived At"
                            ]
                            ws_hist.append(hist_headers)
                            slate_fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
                            for col_idx in range(1, len(hist_headers) + 1):
                                cell = ws_hist.cell(row=1, column=col_idx)
                                cell.font = header_font
                                cell.fill = slate_fill
                                cell.alignment = center_align

                            cursor.execute("PRAGMA table_info(schedules_history)")
                            h_cols = [c[1] for c in cursor.fetchall()]
                            tgl_sj_exp = "tgl_sj" if "tgl_sj" in h_cols else "'' AS tgl_sj"
                            jml_sj_exp = "jml_sj" if "jml_sj" in h_cols else "'' AS jml_sj"
                            pallet_exp = "pallet" if "pallet" in h_cols else "'' AS pallet"
                            box_exp = "box" if "box" in h_cols else "'' AS box"

                            cursor.execute(f"""
                                SELECT id, shipdate, customer, log_partner, cycle, driver, phone,
                                       plan_time, act_time, status, {tgl_sj_exp}, {jml_sj_exp}, {pallet_exp}, {box_exp}, return_pallet,
                                       custom_remark, archived_at
                                FROM schedules_history ORDER BY shipdate DESC, plan_time ASC LIMIT 500
                            """)
                            for row_idx, r in enumerate(cursor.fetchall(), start=2):
                                row_data = [
                                    r["id"], r["shipdate"], r["customer"], r["log_partner"], r["cycle"],
                                    r["driver"], r["phone"] or "-", r["plan_time"], r["act_time"] or "-",
                                    r["status"], r["tgl_sj"] or "-", r["jml_sj"] or "-", r["pallet"] or "-", r["box"] or "-", r["return_pallet"] or "-",
                                    r["custom_remark"] or "-", r["archived_at"] or "-"
                                ]
                                ws_hist.append(row_data)
                                for col_idx in range(1, len(row_data) + 1):
                                    cell = ws_hist.cell(row=row_idx, column=col_idx)
                                    cell.font = regular_font
                                    cell.border = thin_border
                                    if col_idx in [1, 2, 5, 8, 9, 10, 11, 12, 13, 14, 15]:
                                        cell.alignment = center_align
                                    else:
                                        cell.alignment = left_align
                            all_sheets_to_fit.append(ws_hist)
            except Exception as e:
                logger.warning(f"Could not export history archive sheet: {e}")

            # Auto-fit columns for all sheets
            for ws in all_sheets_to_fit:
                ws.row_dimensions[1].height = 26
                for col in ws.columns:
                    col_letter = get_column_letter(col[0].column)
                    max_len = 0
                    for cell in col:
                        val_str = str(cell.value or "")
                        if len(val_str) > max_len:
                            max_len = len(val_str)
                    ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

            wb.save(EXCEL_PATH)
            logger.info(f"Database exported to Excel file: {EXCEL_PATH}")
            return True
        except PermissionError:
            logger.warning("PermissionError: Excel file is currently opened in another program. Sync postponed.")
            return False
        except Exception as e:
            logger.error(f"Error syncing to Excel: {e}")
            return False

def sync_excel_to_sqlite(excel_file_path: Optional[str] = None) -> dict:
    """
    Analyzes an uploaded or existing Excel file, extracts only the required dashboard fields,
    normalizes data format, and updates SQLite and Excel.
    """
    if not excel_file_path:
        excel_file_path = EXCEL_PATH
    with excel_lock:
        wb = openpyxl.load_workbook(excel_file_path, data_only=True)
        res = analyze_and_extract_workbook(wb)

        stats = {
            "schedules_updated": 0,
            "drivers_updated": 0,
            "sheet_used": res["meta"]["schedule_sheet"],
            "ignored_columns": res["meta"]["ignored_columns"]
        }

        with get_db() as conn:
            cursor = conn.cursor()
            # 1. Update Drivers
            for d in res["drivers"]:
                cursor.execute("""
                    INSERT INTO drivers (name, phone, partner, updated_at)
                    VALUES (?, ?, ?, ?)
                    ON CONFLICT(name) DO UPDATE SET
                        phone=excluded.phone,
                        partner=excluded.partner,
                        updated_at=excluded.updated_at
                """, (d["name"], d["phone"], d["partner"], datetime.now().isoformat()))
                stats["drivers_updated"] += 1

            # 2. Update Schedules
            for s in res["schedules"]:
                cursor.execute("""
                    INSERT INTO schedules (
                        id, shipdate, customer, log_partner, cycle, driver, phone,
                        plan_time, act_time, status, custom_remark,
                        tgl_sj, jml_sj, pallet, box,
                        sj_date_qty, pallet_box, return_pallet, updated_at
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(id) DO UPDATE SET
                        shipdate=excluded.shipdate,
                        customer=excluded.customer,
                        log_partner=excluded.log_partner,
                        cycle=excluded.cycle,
                        driver=excluded.driver,
                        phone=excluded.phone,
                        plan_time=excluded.plan_time,
                        act_time=excluded.act_time,
                        status=excluded.status,
                        custom_remark=excluded.custom_remark,
                        tgl_sj=excluded.tgl_sj,
                        jml_sj=excluded.jml_sj,
                        pallet=excluded.pallet,
                        box=excluded.box,
                        sj_date_qty=excluded.sj_date_qty,
                        pallet_box=excluded.pallet_box,
                        return_pallet=excluded.return_pallet,
                        updated_at=excluded.updated_at
                """, (
                    s["id"], s["shipdate"], s["customer"], s["log_partner"], s["cycle"],
                    s["driver"], s["phone"], s["plan_time"], s["act_time"], s["status"],
                    s["custom_remark"],
                    s.get("tgl_sj", ""), s.get("jml_sj", ""), s.get("pallet", ""), s.get("box", ""),
                    s.get("sj_date_qty", ""), s.get("pallet_box", ""), s["return_pallet"],
                    datetime.now().isoformat()
                ))
                stats["schedules_updated"] += 1

            conn.commit()

        # Re-export to enforce standardized format in database_milkrun.xlsx
        sync_sqlite_to_excel()
        return stats


# --- AUTHENTICATION & ACCESS CONTROL ("tetapkan akses setelah login") ---
def verify_admin_auth(x_admin_token: Optional[str] = Header(None)):
    """Verifies that requests to modifying endpoints have valid admin credentials."""
    if not x_admin_token or x_admin_token != ADMIN_SECRET_TOKEN:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Akses Ditolak: Anda harus login sebagai Admin terlebih dahulu!"
        )
    return True

@app.post("/api/auth/login")
def admin_login(req: LoginRequest):
    """Admin PIN authentication endpoint."""
    pin_cleaned = req.pin.strip().lower()
    if pin_cleaned in [ADMIN_PIN.lower(), "admin", "1234"]:
        return {
            "success": True,
            "token": ADMIN_SECRET_TOKEN,
            "role": "ADMIN",
            "message": "Login Admin berhasil. Akses penuh aktif."
        }
    raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="PIN Admin salah! Harap periksa kembali.")


# --- REST API ENDPOINTS ---

@app.get("/api/schedules")
def list_schedules():
    """Public read access: view delivery board."""
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT id, shipdate, customer, log_partner, cycle, driver, phone,
                   plan_time, act_time, status, custom_remark,
                   tgl_sj, jml_sj, pallet, box, return_pallet, sj_date_qty, pallet_box
            FROM schedules ORDER BY shipdate DESC, plan_time ASC
        """)
        rows = cursor.fetchall()
        result = []
        for r in rows:
            tgl_sj = r["tgl_sj"] or ""
            jml_sj = r["jml_sj"] or ""
            pallet = r["pallet"] or ""
            box = r["box"] or ""
            sj_date_qty = r["sj_date_qty"] or ""
            pallet_box = r["pallet_box"] or ""

            # Auto decompose compound if separate fields are empty
            if (not tgl_sj and not jml_sj) and sj_date_qty:
                parts = [p.strip() for p in sj_date_qty.split("/")]
                if len(parts) >= 3:
                    tgl_sj = f"{parts[0]}/{parts[1]}"
                    jml_sj = parts[2]
                elif len(parts) == 2:
                    tgl_sj = parts[0]
                    jml_sj = parts[1]
                elif len(parts) == 1:
                    tgl_sj = parts[0]

            if (not pallet and not box) and pallet_box:
                parts = [p.strip() for p in pallet_box.split("/")]
                if len(parts) >= 2:
                    pallet = parts[0]
                    box = parts[1]
                elif len(parts) == 1:
                    pallet = parts[0]

            compound_sj = sj_date_qty or (f"{tgl_sj} / {jml_sj}".strip(" /") if (tgl_sj or jml_sj) else "")
            compound_pb = pallet_box or (f"{pallet} / {box}".strip(" /") if (pallet or box) else "")

            result.append({
                "id": r["id"],
                "shipdate": r["shipdate"],
                "customer": r["customer"],
                "logPartner": r["log_partner"],
                "cycle": r["cycle"],
                "driver": r["driver"],
                "phone": r["phone"] or "",
                "planTime": r["plan_time"],
                "actTime": r["act_time"] or "",
                "status": r["status"],
                "customRemark": r["custom_remark"] or "",
                "tglSj": tgl_sj,
                "jmlSj": jml_sj,
                "pallet": pallet,
                "box": box,
                "returnPallet": r["return_pallet"] or "",
                "tgl_sj": tgl_sj,
                "jml_sj": jml_sj,
                "pallet_qty": pallet,
                "box_qty": box,
                "return_pallet": r["return_pallet"] or "",
                "sjDateQty": compound_sj,
                "palletBox": compound_pb
            })
        return result

@app.post("/api/schedules")
def create_schedule(schedule: ScheduleModel, is_admin: bool = Depends(verify_admin_auth)):
    """Admin Only: Add new schedule and auto-sync to Excel."""
    with get_db() as conn:
        cursor = conn.cursor()
        sched_id = schedule.id
        if not sched_id:
            cursor.execute("SELECT id FROM schedules WHERE id LIKE 'MR-%'")
            ids = [r[0] for r in cursor.fetchall()]
            num = 101
            while f"MR-{num}" in ids:
                num += 1
            sched_id = f"MR-{num}"

        tgl_sj = schedule.tgl_sj if schedule.tgl_sj is not None else (schedule.tglSj or "")
        jml_sj = schedule.jml_sj if schedule.jml_sj is not None else (schedule.jmlSj or "")
        pallet = schedule.pallet if schedule.pallet is not None else (schedule.palletQty or "")
        box = schedule.box if schedule.box is not None else (schedule.boxQty or "")
        rp_val = schedule.return_pallet if schedule.return_pallet is not None else (schedule.returnPallet or "")

        sj_val = schedule.sj_date_qty if schedule.sj_date_qty is not None else (schedule.sjDateQty or "")
        pb_val = schedule.pallet_box if schedule.pallet_box is not None else (schedule.palletBox or "")

        if not tgl_sj and not jml_sj and sj_val:
            parts = [p.strip() for p in sj_val.split("/")]
            if len(parts) >= 2:
                tgl_sj, jml_sj = parts[0], parts[1]
            elif len(parts) == 1:
                tgl_sj = parts[0]
        if not pallet and not box and pb_val:
            parts = [p.strip() for p in pb_val.split("/")]
            if len(parts) >= 2:
                pallet, box = parts[0], parts[1]
            elif len(parts) == 1:
                pallet = parts[0]

        cursor.execute("""
            INSERT OR REPLACE INTO schedules (
                id, shipdate, customer, log_partner, cycle, driver, phone,
                plan_time, act_time, status, custom_remark,
                tgl_sj, jml_sj, pallet, box, return_pallet,
                sj_date_qty, pallet_box, updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            sched_id, schedule.shipdate, schedule.customer, schedule.logPartner,
            schedule.cycle, schedule.driver, schedule.phone or "",
            schedule.planTime, schedule.actTime or "", schedule.status or "PENDING",
            schedule.customRemark or "",
            tgl_sj, jml_sj, pallet, box, rp_val,
            sj_val, pb_val,
            datetime.now().isoformat()
        ))
        conn.commit()

    sync_sqlite_to_excel()
    return {"success": True, "id": sched_id, "message": "Jadwal berhasil ditambahkan & disimpan ke Excel."}

@app.put("/api/schedules/{schedule_id}")
def update_schedule(schedule_id: str, schedule: ScheduleModel, is_admin: bool = Depends(verify_admin_auth)):
    """Admin Only: Update schedule and auto-sync to Excel."""
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM schedules WHERE id = ?", (schedule_id,))
        if not cursor.fetchone():
            raise HTTPException(status_code=404, detail="Jadwal tidak ditemukan")

        cursor.execute("SELECT tgl_sj, jml_sj, pallet, box, return_pallet, sj_date_qty, pallet_box FROM schedules WHERE id = ?", (schedule_id,))
        existing = cursor.fetchone()

        tgl_sj = schedule.tgl_sj if schedule.tgl_sj is not None else schedule.tglSj
        if tgl_sj is None:
            tgl_sj = existing["tgl_sj"] if existing and existing["tgl_sj"] else ""

        jml_sj = schedule.jml_sj if schedule.jml_sj is not None else schedule.jmlSj
        if jml_sj is None:
            jml_sj = existing["jml_sj"] if existing and existing["jml_sj"] else ""

        pallet = schedule.pallet if schedule.pallet is not None else schedule.palletQty
        if pallet is None:
            pallet = existing["pallet"] if existing and existing["pallet"] else ""

        box = schedule.box if schedule.box is not None else schedule.boxQty
        if box is None:
            box = existing["box"] if existing and existing["box"] else ""

        rp_val = schedule.return_pallet if schedule.return_pallet is not None else schedule.returnPallet
        if rp_val is None:
            rp_val = existing["return_pallet"] if existing and existing["return_pallet"] else ""

        sj_val = schedule.sj_date_qty if schedule.sj_date_qty is not None else schedule.sjDateQty
        if sj_val is None:
            sj_val = existing["sj_date_qty"] if existing and existing["sj_date_qty"] else ""

        pb_val = schedule.pallet_box if schedule.pallet_box is not None else schedule.palletBox
        if pb_val is None:
            pb_val = existing["pallet_box"] if existing and existing["pallet_box"] else ""

        cursor.execute("""
            UPDATE schedules SET
                shipdate = ?, customer = ?, log_partner = ?, cycle = ?,
                driver = ?, phone = ?, plan_time = ?, act_time = ?,
                status = ?, custom_remark = ?,
                tgl_sj = ?, jml_sj = ?, pallet = ?, box = ?, return_pallet = ?,
                sj_date_qty = ?, pallet_box = ?, updated_at = ?
            WHERE id = ?
        """, (
            schedule.shipdate, schedule.customer, schedule.logPartner, schedule.cycle,
            schedule.driver, schedule.phone or "", schedule.planTime, schedule.actTime or "",
            schedule.status or "PENDING", schedule.customRemark or "",
            tgl_sj, jml_sj, pallet, box, rp_val,
            sj_val, pb_val,
            datetime.now().isoformat(), schedule_id
        ))
        conn.commit()

    sync_sqlite_to_excel()
    return {"success": True, "message": f"Jadwal {schedule_id} berhasil diperbarui & sinkron ke Excel."}


@app.patch("/api/schedules/{schedule_id}/logistics")
@app.post("/api/schedules/{schedule_id}/logistics")
def update_schedule_logistics(schedule_id: str, payload: LogisticsUpdateModel):
    """
    Public Endpoint (Tanpa Login Admin):
    Memungkinkan PIC Admin Logistics menginput / mengupdate:
    1. Tanggal Surat Jalan (Tgl SJ)
    2. Jumlah Surat Jalan (Jml SJ)
    3. Jumlah Pallet
    4. Jumlah Box
    5. Jumlah Return Pallet
    Data langsung tersimpan secara realtime ke Database SQLite dan file Excel (.xlsx).
    """
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM schedules WHERE id = ?", (schedule_id,))
        if not cursor.fetchone():
            raise HTTPException(status_code=404, detail="Jadwal tidak ditemukan")

        tgl_sj = payload.tgl_sj if payload.tgl_sj is not None else payload.tglSj
        jml_sj = payload.jml_sj if payload.jml_sj is not None else payload.jmlSj
        pallet = payload.pallet if payload.pallet is not None else payload.palletQty
        box = payload.box if payload.box is not None else payload.boxQty
        rp_val = payload.return_pallet if payload.return_pallet is not None else payload.returnPallet

        sj_val = payload.sj_date_qty if payload.sj_date_qty is not None else payload.sjDateQty
        pb_val = payload.pallet_box if payload.pallet_box is not None else payload.palletBox

        # Handle legacy compound payloads if separate values were omitted
        if sj_val is not None and (tgl_sj is None and jml_sj is None):
            parts = [p.strip() for p in sj_val.split("/")]
            if len(parts) >= 3:
                tgl_sj = f"{parts[0]}/{parts[1]}"
                jml_sj = parts[2]
            elif len(parts) == 2:
                tgl_sj = parts[0]
                jml_sj = parts[1]
            elif len(parts) == 1:
                tgl_sj = parts[0]
                jml_sj = ""

        if pb_val is not None and (pallet is None and box is None):
            parts = [p.strip() for p in pb_val.split("/")]
            if len(parts) >= 2:
                pallet = parts[0]
                box = parts[1]
            elif len(parts) == 1:
                pallet = parts[0]
                box = ""

        updates = []
        params = []
        if tgl_sj is not None:
            updates.append("tgl_sj = ?")
            params.append(tgl_sj.strip())
        if jml_sj is not None:
            updates.append("jml_sj = ?")
            params.append(jml_sj.strip())
        if pallet is not None:
            updates.append("pallet = ?")
            params.append(pallet.strip())
        if box is not None:
            updates.append("box = ?")
            params.append(box.strip())
        if rp_val is not None:
            updates.append("return_pallet = ?")
            params.append(rp_val.strip())
        if sj_val is not None:
            updates.append("sj_date_qty = ?")
            params.append(sj_val.strip())
        if pb_val is not None:
            updates.append("pallet_box = ?")
            params.append(pb_val.strip())

        if not updates:
            return {"success": True, "message": "Tidak ada perubahan data."}

        updates.append("updated_at = ?")
        params.append(datetime.now().isoformat())
        params.append(schedule_id)

        sql = f"UPDATE schedules SET {', '.join(updates)} WHERE id = ?"
        cursor.execute(sql, tuple(params))
        conn.commit()

    sync_sqlite_to_excel()
    return {
        "success": True,
        "schedule_id": schedule_id,
        "message": f"Data logistik {schedule_id} berhasil disimpan ke Database & Excel."
    }

@app.delete("/api/schedules/{schedule_id}")
def delete_schedule(schedule_id: str, is_admin: bool = Depends(verify_admin_auth)):
    """Admin Only: Delete schedule and auto-sync to Excel."""
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM schedules WHERE id = ?", (schedule_id,))
        if cursor.rowcount == 0:
            raise HTTPException(status_code=404, detail="Jadwal tidak ditemukan")
        conn.commit()

    sync_sqlite_to_excel()
    return {"success": True, "message": f"Jadwal {schedule_id} telah dihapus & sinkron ke Excel."}

@app.get("/api/drivers")
def list_drivers():
    """Public read access: driver directory with dynamic status from dashboard schedules, route stats, and last active timestamp."""
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT d.name, d.phone, d.partner, d.last_active, d.updated_at,
                   COUNT(CASE WHEN s.status = 'PENDING' THEN 1 END) as active_routes,
                   COUNT(s.id) as total_routes,
                   MAX(s.shipdate || ' ' || COALESCE(NULLIF(s.act_time, ''), s.plan_time)) as last_trip_time
            FROM drivers d
            LEFT JOIN schedules s ON UPPER(TRIM(s.driver)) = UPPER(TRIM(d.name))
            GROUP BY d.name, d.phone, d.partner, d.last_active, d.updated_at
            ORDER BY d.name ASC
        """)
        rows = cursor.fetchall()
        return [
            {
                "name": r["name"],
                "phone": r["phone"],
                "partner": r["partner"],
                "status": "AKTIF" if (r["total_routes"] or 0) > 0 else "TIDAK AKTIF",
                "active_routes": r["active_routes"] or 0,
                "total_routes": r["total_routes"] or 0,
                "last_active": r["last_active"] or r["last_trip_time"] or r["updated_at"] or ""
            }
            for r in rows
        ]

@app.post("/api/drivers")
def add_driver(driver: DriverModel):
    """Public Registration: Add/register new driver and auto-sync to Excel."""
    driver_name = driver.name.strip().upper()
    driver_phone = driver.phone.strip()
    driver_partner = driver.partner.strip().upper()
    driver_status = (driver.status or "AKTIF").strip().upper()

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM drivers WHERE UPPER(name) = ?", (driver_name,))
        if cursor.fetchone():
            raise HTTPException(
                status_code=400,
                detail=f"Nama driver '{driver_name}' sudah terdaftar! Harap gunakan nama baru atau tambahkan nama belakang/inisial (contoh: {driver_name} S.)."
            )

        now_wib_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cursor.execute("""
            INSERT INTO drivers (name, phone, partner, status, last_active, updated_at)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (driver_name, driver_phone, driver_partner, driver_status, now_wib_str, datetime.now().isoformat()))
        conn.commit()

    sync_sqlite_to_excel()
    return {"success": True, "message": f"Driver {driver_name} berhasil didaftarkan & disinkronkan ke Excel."}

@app.put("/api/drivers/{name}")
def update_driver(name: str, driver: DriverUpdateModel, is_admin: bool = Depends(verify_admin_auth)):
    """Admin Only: Edit driver name, phone, partner, and status in master directory."""
    old_name = name.strip().upper()
    new_name = driver.name.strip().upper()
    new_phone = driver.phone.strip()
    new_partner = driver.partner.strip().upper()
    new_status = (driver.status or "AKTIF").strip().upper()

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM drivers WHERE UPPER(name) = ?", (old_name,))
        if not cursor.fetchone():
            raise HTTPException(status_code=404, detail=f"Driver {old_name} tidak ditemukan.")

        if old_name != new_name:
            cursor.execute("SELECT name FROM drivers WHERE UPPER(name) = ? AND UPPER(name) != ?", (new_name, old_name))
            if cursor.fetchone():
                raise HTTPException(
                    status_code=400,
                    detail=f"Nama driver '{new_name}' sudah terdaftar! Harap gunakan nama baru atau tambahkan nama belakang/inisial (contoh: {new_name} S.)."
                )

            cursor.execute("""
                UPDATE schedules SET driver = ?, phone = ? WHERE UPPER(driver) = ?
            """, (new_name, new_phone, old_name))

            cursor.execute("""
                UPDATE drivers SET name = ?, phone = ?, partner = ?, status = ?, updated_at = ? WHERE UPPER(name) = ?
            """, (new_name, new_phone, new_partner, new_status, datetime.now().isoformat(), old_name))
        else:
            cursor.execute("""
                UPDATE drivers SET phone = ?, partner = ?, status = ?, updated_at = ? WHERE UPPER(name) = ?
            """, (new_phone, new_partner, new_status, datetime.now().isoformat(), old_name))

            cursor.execute("""
                UPDATE schedules SET phone = ? WHERE UPPER(driver) = ?
            """, (new_phone, old_name))

        conn.commit()

    sync_sqlite_to_excel()
    return {
        "success": True,
        "message": f"Data driver {new_name} ({new_partner}) berhasil diperbarui & disinkronkan ke Excel.",
        "driver": {"name": new_name, "phone": new_phone, "partner": new_partner, "status": new_status}
    }

@app.patch("/api/drivers/{name}/status")
@app.put("/api/drivers/{name}/status")
def toggle_driver_status(name: str, payload: dict = Body(...), is_admin: bool = Depends(verify_admin_auth)):
    """Admin Only: Toggle or set driver status ('AKTIF' or 'TIDAK AKTIF')."""
    clean_name = name.strip().upper()
    new_status = payload.get("status", "AKTIF").strip().upper()
    if new_status not in ["AKTIF", "TIDAK AKTIF"]:
        raise HTTPException(status_code=400, detail="Status harus 'AKTIF' atau 'TIDAK AKTIF'.")

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("UPDATE drivers SET status = ?, updated_at = ? WHERE UPPER(name) = ?", (new_status, datetime.now().isoformat(), clean_name))
        if cursor.rowcount == 0:
            raise HTTPException(status_code=404, detail=f"Driver '{clean_name}' tidak ditemukan.")
        conn.commit()

    sync_sqlite_to_excel()
    return {"success": True, "name": clean_name, "status": new_status, "message": f"Status driver {clean_name} diubah menjadi {new_status}."}

@app.delete("/api/drivers/{name}")
def delete_driver(name: str, is_admin: bool = Depends(verify_admin_auth)):
    """
    Admin Only: Delete driver.
    RULE: Driver yang sedang terdata pada jadwal dashboard utama (schedules) TIDAK DAPAT DIHAPUS.
    Driver hanya dapat dihapus jika tidak sedang terdata pada jadwal di dashboard utama.
    """
    clean_name = name.strip().upper()
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM drivers WHERE UPPER(name) = ?", (clean_name,))
        d_row = cursor.fetchone()
        if not d_row:
            raise HTTPException(status_code=404, detail="Driver tidak ditemukan")

        cursor.execute("SELECT COUNT(*) FROM schedules WHERE UPPER(TRIM(driver)) = ?", (clean_name,))
        active_sched_count = cursor.fetchone()[0]
        if active_sched_count > 0:
            raise HTTPException(
                status_code=400,
                detail=f"Driver '{clean_name}' sedang berstatus AKTIF ({active_sched_count} jadwal di Dashboard Utama)! Driver yang sedang terdata pada jadwal Dashboard Utama tidak dapat dihapus."
            )

        cursor.execute("DELETE FROM drivers WHERE UPPER(name) = ?", (clean_name,))
        conn.commit()

    sync_sqlite_to_excel()
    return {"success": True, "message": f"Driver {clean_name} berhasil dihapus dari Master Directory."}

@app.get("/api/settings")
def get_app_settings():
    """Public read access: Retrieve app settings (Geofence coordinates, n8n webhook URL, etc.)."""
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT key, value FROM settings")
        rows = cursor.fetchall()
        settings_dict = {r["key"]: r["value"] for r in rows}
        return {
            "geofence_lat": float(settings_dict.get("geofence_lat", -6.269983)),
            "geofence_lng": float(settings_dict.get("geofence_lng", 107.183808)),
            "geofence_radius": int(settings_dict.get("geofence_radius", 500)),
            "geofence_name": settings_dict.get("geofence_name", "PT. Nifco Indonesia"),
            "n8n_webhook_url": settings_dict.get("n8n_webhook_url", "")
        }

@app.post("/api/settings")
def save_app_settings(payload: dict = Body(...), is_admin: bool = Depends(verify_admin_auth)):
    """Admin Only: Save or update app settings (Geofence & Webhook URL)."""
    with get_db() as conn:
        cursor = conn.cursor()
        for k, v in payload.items():
            cursor.execute("""
                INSERT INTO settings (key, value, updated_at)
                VALUES (?, ?, ?)
                ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = excluded.updated_at
            """, (str(k), str(v), datetime.now().isoformat()))
        conn.commit()
    return {"success": True, "message": "Pengaturan Geofence & Webhook berhasil disimpan."}

@app.post("/api/webhook/alert")
async def trigger_webhook_alert(payload: dict = Body(...)):
    """
    Sends asynchronous notification to n8n Webhook (WhatsApp / Telegram alert).
    Edge Case & Failure State:
    If n8n Webhook or WhatsApp API is down/offline, backend catches the error gracefully (try-catch with 4s timeout).
    Status DELAY remains securely saved in SQLite and Excel, dashboard renders red, and backend NEVER crashes.
    """
    webhook_url = payload.get("webhook_url")
    if not webhook_url:
        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT value FROM settings WHERE key = 'n8n_webhook_url'")
            row = cursor.fetchone()
            if row:
                webhook_url = row["value"]

    schedule_data = payload.get("schedule", {})
    if not webhook_url:
        return {"success": False, "message": "Webhook URL belum dikonfigurasi."}

    # Asynchronous non-blocking try-catch with timeout
    try:
        import urllib.request
        import json

        req_data = json.dumps({
            "event": "SCHEDULE_DELAY_ALERT",
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S WIB"),
            "schedule": schedule_data,
            "message": f"PERINGATAN MILKRUN: Rute {schedule_data.get('customer', '-')} ({schedule_data.get('logPartner', '-')}) berstatus DELAY! Driver: {schedule_data.get('driver', '-')}."
        }).encode("utf-8")

        req = urllib.request.Request(
            webhook_url,
            data=req_data,
            headers={"Content-Type": "application/json", "User-Agent": "NifcoMilkrun/1.0"},
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=4) as response:
            status_code = response.getcode()
            return {"success": True, "status_code": status_code, "message": "Notifikasi webhook n8n berhasil dikirim."}
    except Exception as e:
        logger.warning(f"Webhook n8n / WhatsApp API Down: {e}. Status DELAY tetap aman tersimpan di Database.")
        return {
            "success": False,
            "error": str(e),
            "message": "Webhook n8n / WhatsApp API sedang down/offline. Status DELAY tetap aman tersimpan di Database."
        }

@app.post("/api/driver/checkin")
def driver_checkin(checkin: DriverArrivalCheckIn):
    """Public endpoint for Driver Mobile View: check-in physical arrival with driver selection."""
    sched_id = checkin.scheduleId or checkin.trip_id
    actual_time = checkin.actTime or checkin.act_time
    driver_name = (checkin.driverName or "").strip().upper()
    note = checkin.customNote or checkin.validation_note or ""

    if not sched_id or not actual_time:
        raise HTTPException(status_code=400, detail="trip_id / scheduleId dan act_time / actTime wajib diisi.")

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT plan_time, customer, custom_remark, driver, phone FROM schedules WHERE id = ?", (sched_id,))
        row = cursor.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Jadwal tidak ditemukan")

        plan_time = row["plan_time"]
        customer = row["customer"]
        existing_remark = row["custom_remark"] or ""
        assigned_driver = row["driver"]
        assigned_phone = row["phone"] or ""

        now_wib_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        now_iso = datetime.now().isoformat()
        if driver_name:
            assigned_driver = driver_name
            cursor.execute("SELECT phone, partner FROM drivers WHERE name = ?", (driver_name,))
            d_row = cursor.fetchone()
            if d_row:
                assigned_phone = d_row["phone"] or assigned_phone
            cursor.execute("""
                UPDATE drivers 
                SET last_active = ?, updated_at = ? 
                WHERE name = ?
            """, (now_wib_str, now_iso, driver_name))

        # Calculate status automatically based on minutes difference with midnight wraparound & 30-min ONTIME tolerance
        try:
            plan_h, plan_m = map(int, plan_time.split(":"))
            act_h, act_m = map(int, actual_time.split(":"))
            diff_minutes = (act_h * 60 + act_m) - (plan_h * 60 + plan_m)
            if diff_minutes < -720:
                diff_minutes += 1440
            elif diff_minutes > 720:
                diff_minutes -= 1440

            if diff_minutes < -5:
                status_calc = "ADVANCE"
            elif -5 <= diff_minutes <= 30:
                status_calc = "ONTIME"
            else:
                status_calc = "DELAY"
        except Exception:
            status_calc = "ONTIME"

        new_remark = existing_remark
        if note:
            new_remark = f"{existing_remark} | {note}" if existing_remark else note

        cursor.execute("""
            UPDATE schedules 
            SET act_time = ?, status = ?, custom_remark = ?, driver = ?, phone = ?, updated_at = ?
            WHERE id = ?
        """, (actual_time, status_calc, new_remark, assigned_driver, assigned_phone, datetime.now().isoformat(), sched_id))
        conn.commit()

    sync_sqlite_to_excel()
    return {
        "success": True,
        "customer": customer,
        "driver": assigned_driver,
        "act_time": actual_time,
        "status": status_calc,
        "message": f"Kedatangan driver {assigned_driver} tercatat: {actual_time} (Status: {status_calc})"
    }


# --- EXCEL DATABASE ADMIN CONTROLS ---

@app.get("/api/excel/download")
def download_excel(token: Optional[str] = Query(None), x_admin_token: Optional[str] = Header(None)):
    """Admin Only: Download live Excel database."""
    auth_token = x_admin_token or token
    if auth_token != ADMIN_SECRET_TOKEN:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Login Admin diperlukan untuk mengunduh database Excel.")
    
    sync_sqlite_to_excel()
    if not os.path.exists(EXCEL_PATH):
        raise HTTPException(status_code=404, detail="File Excel database belum tersedia.")
    
    return FileResponse(
        path=EXCEL_PATH,
        filename=f"database_milkrun_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.post("/api/excel/import")
async def import_excel(file: UploadFile = File(...), is_admin: bool = Depends(verify_admin_auth)):
    """Admin Only: Upload and import Excel to database."""
    if not file.filename.endswith((".xlsx", ".xlsm", ".xltx")):
        raise HTTPException(status_code=400, detail="Format file harus berupa Excel (.xlsx)")

    temp_path = os.path.join(BASE_DIR, f"_temp_import_{file.filename}")
    try:
        with open(temp_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        stats = sync_excel_to_sqlite(temp_path)
        ignored_info = f" (Kolom non-standar diabaikan: {', '.join(stats['ignored_columns'])})" if stats.get("ignored_columns") else ""
        return {
            "success": True,
            "message": f"Analisis & Ekstraksi Excel Berhasil! Lembar kerja '{stats.get('sheet_used', '-')}' diproses: {stats['schedules_updated']} jadwal & {stats['drivers_updated']} driver disinkronkan{ignored_info}.",
            "stats": stats
        }
    except Exception as e:
        logger.error(f"Error importing Excel: {e}")
        raise HTTPException(status_code=500, detail=f"Gagal memproses file Excel: {str(e)}")
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)

@app.post("/api/excel/sync")
def force_sync_excel(is_admin: bool = Depends(verify_admin_auth)):
    """Admin Only: Force synchronization from SQLite to Excel."""
    success = sync_sqlite_to_excel()
    if not success:
        raise HTTPException(status_code=500, detail="Sinkronisasi Excel gagal atau file sedang dibuka di Excel.")
    return {"success": True, "message": "Database berhasil disinkronkan ke Excel."}

@app.get("/api/status")
@app.get("/health")
def system_status():
    """System health check & statistics."""
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM schedules")
        total_schedules = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM drivers")
        total_drivers = cursor.fetchone()[0]

    return {
        "status": "ONLINE",
        "engine": "FastAPI + SQLite3 + OpenPyXL",
        "total_schedules": total_schedules,
        "total_drivers": total_drivers,
        "excel_file": os.path.basename(EXCEL_PATH),
        "excel_exists": os.path.exists(EXCEL_PATH)
    }



# --- AUTOMATED 02:00 AM DAILY SCHEDULE GENERATOR ---
last_daily_run = None

def check_and_generate_daily_schedule(target_date: Optional[str] = None) -> dict:
    """
    Creates new schedule records for the new operational day at 02:00 AM:
    - Archives previous days' completed schedules to schedules_history table.
    - Rollovers active routes for target_date with status PENDING, empty actual arrival, and empty logistics fields.
    - Synchronizes new daily schedule to database_milkrun.xlsx.
    """
    global last_daily_run
    if not target_date:
        target_date = datetime.now().strftime("%Y-%m-%d")

    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS schedules_history (
                id TEXT,
                shipdate TEXT,
                customer TEXT,
                log_partner TEXT,
                cycle INTEGER,
                driver TEXT,
                phone TEXT,
                plan_time TEXT,
                act_time TEXT,
                status TEXT,
                custom_remark TEXT,
                tgl_sj TEXT,
                jml_sj TEXT,
                pallet TEXT,
                box TEXT,
                sj_date_qty TEXT,
                pallet_box TEXT,
                return_pallet TEXT,
                updated_at TEXT,
                archived_at TEXT,
                PRIMARY KEY (id, shipdate)
            )
        """)
        cursor.execute("PRAGMA table_info(schedules_history)")
        sh_cols = [c[1] for c in cursor.fetchall()]
        if "tgl_sj" not in sh_cols:
            cursor.execute("ALTER TABLE schedules_history ADD COLUMN tgl_sj TEXT DEFAULT ''")
        if "jml_sj" not in sh_cols:
            cursor.execute("ALTER TABLE schedules_history ADD COLUMN jml_sj TEXT DEFAULT ''")
        if "pallet" not in sh_cols:
            cursor.execute("ALTER TABLE schedules_history ADD COLUMN pallet TEXT DEFAULT ''")
        if "box" not in sh_cols:
            cursor.execute("ALTER TABLE schedules_history ADD COLUMN box TEXT DEFAULT ''")

        cursor.execute("SELECT COUNT(*) FROM schedules WHERE shipdate = ?", (target_date,))
        count_today = cursor.fetchone()[0]

        if count_today > 0 and last_daily_run == target_date:
            return {"status": "ALREADY_EXISTS", "date": target_date, "count": count_today}

        if count_today == 0:
            logger.info(f"[02:00 AM SCHEDULER] Date changed! Generating new daily schedules for {target_date}...")

            # 1. Archive older completed schedules to history
            cursor.execute("""
                INSERT OR REPLACE INTO schedules_history (
                    id, shipdate, customer, log_partner, cycle, driver, phone,
                    plan_time, act_time, status, custom_remark,
                    tgl_sj, jml_sj, pallet, box,
                    sj_date_qty, pallet_box, return_pallet, updated_at, archived_at
                )
                SELECT id, shipdate, customer, log_partner, cycle, driver, phone,
                       plan_time, act_time, status, custom_remark,
                       tgl_sj, jml_sj, pallet, box,
                       sj_date_qty, pallet_box, return_pallet, updated_at, ?
                FROM schedules
                WHERE shipdate < ?
            """, (datetime.now().isoformat(), target_date))

            # 2. Get route templates from existing schedules
            cursor.execute("SELECT id, customer, log_partner, cycle, plan_time, driver, phone FROM schedules ORDER BY plan_time ASC")
            existing_routes = cursor.fetchall()

            if existing_routes:
                for r in existing_routes:
                    cursor.execute("""
                        UPDATE schedules SET
                            shipdate = ?,
                            act_time = '',
                            status = 'PENDING',
                            custom_remark = '',
                            tgl_sj = '',
                            jml_sj = '',
                            pallet = '',
                            box = '',
                            sj_date_qty = '',
                            pallet_box = '',
                            return_pallet = '',
                            updated_at = ?
                        WHERE id = ?
                    """, (target_date, datetime.now().isoformat(), r["id"]))
                generated_count = len(existing_routes)
            else:
                initial_schedules = [
                    ("MR-101", target_date, "TMMIN S2", "SYNCRUM", 1, "POSTER", "081298765432", "06:15", "", "PENDING", ""),
                    ("MR-102", target_date, "MMKI 1", "DARSINDO", 1, "ASEP", "081234567890", "07:00", "", "PENDING", ""),
                    ("MR-103", target_date, "YIMMCK", "TTNT", 2, "YADI", "081566778899", "03:30", "", "PENDING", ""),
                    ("MR-104", target_date, "AHM KRW", "SELOG", 1, "DEDI", "081788990011", "07:25", "", "PENDING", ""),
                    ("MR-105", target_date, "TMMIN P2/1", "SYNCRUM", 3, "NAPIN", "081922334455", "06:05", "", "PENDING", ""),
                    ("MR-106", target_date, "ADM KAWASAN", "SELOG", 1, "", "", "08:00", "", "PENDING", "")
                ]
                for s in initial_schedules:
                    cursor.execute("""
                        INSERT INTO schedules (id, shipdate, customer, log_partner, cycle, driver, phone, plan_time, act_time, status, custom_remark, updated_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ON CONFLICT(id) DO UPDATE SET
                            shipdate=excluded.shipdate, act_time='', status='PENDING', custom_remark='',
                            tgl_sj='', jml_sj='', pallet='', box='', sj_date_qty='', pallet_box='', return_pallet='',
                            updated_at=excluded.updated_at
                    """, (s[0], s[1], s[2], s[3], s[4], s[5], s[6], s[7], s[8], s[9], s[10], datetime.now().isoformat()))
                generated_count = len(initial_schedules)

            conn.commit()
            last_daily_run = target_date
            sync_sqlite_to_excel()
            logger.info(f"[02:00 AM SCHEDULER] Successfully created {generated_count} schedules for {target_date} in SQLite and Excel.")
            return {"status": "SUCCESS", "date": target_date, "count": generated_count}
        else:
            last_daily_run = target_date
            return {"status": "ALREADY_EXISTS", "date": target_date, "count": count_today}

def daily_schedule_worker():
    """Background worker daemon running continuously to trigger daily schedule generation at 02:00 AM."""
    while True:
        try:
            now = datetime.now()
            today_str = now.strftime("%Y-%m-%d")
            # If current hour is at or after 02:00 AM, ensure today's schedule is generated
            if now.hour >= 2:
                check_and_generate_daily_schedule(today_str)
        except Exception as e:
            logger.error(f"Error in daily schedule worker: {e}")
        time_module.sleep(30)


@app.post("/api/schedules/daily-generate")
def api_daily_generate(date: Optional[str] = None):
    """Endpoint to trigger daily schedule generation (defaults to today at 02:00 AM cutoff)."""
    res = check_and_generate_daily_schedule(date)
    return res

@app.get("/api/schedules/daily-status")
def api_daily_status():
    """Status of daily schedule automation."""
    now = datetime.now()
    today_str = now.strftime("%Y-%m-%d")
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM schedules WHERE shipdate = ?", (today_str,))
        count_today = cursor.fetchone()[0]
    return {
        "current_time_wib": now.strftime("%Y-%m-%d %H:%M:%S"),
        "cutoff_hour": "02:00 AM",
        "today_date": today_str,
        "schedules_for_today": count_today,
        "last_daily_run": last_daily_run
    }


# Serve the Dashboard HTML directly on root / and direct file paths
@app.get("/", response_class=HTMLResponse)
def serve_dashboard():
    for fpath in [INDEX_HTML_PATH, PROJECT_HTML_PATH]:
        if os.path.exists(fpath):
            with open(fpath, "r", encoding="utf-8") as f:
                return f.read()
    return "<h1>Dashboard HTML file not found</h1>"

@app.get("/index.html", response_class=HTMLResponse)
def serve_index_route():
    return serve_dashboard()

@app.get("/Project Dasboard.html", response_class=HTMLResponse)
def serve_project_dashboard_route():
    return serve_dashboard()

@app.get("/Project%20Dasboard.html", response_class=HTMLResponse)
def serve_project_dashboard_encoded_route():
    return serve_dashboard()


if __name__ == "__main__":
    init_db()
    # Start automated 02:00 AM scheduler background thread
    scheduler_thread = threading.Thread(target=daily_schedule_worker, daemon=True)
    scheduler_thread.start()
    logger.info("[02:00 AM SCHEDULER] Daily background scheduler thread started.")
    # Check on startup if today schedule is already set
    now_init = datetime.now()
    if now_init.hour >= 2:
        check_and_generate_daily_schedule(now_init.strftime("%Y-%m-%d"))

    import uvicorn
    print("\n" + "=" * 65)
    print(" [OK] PT. NIFCO INDONESIA - MILKRUN BACKEND SYSTEM STARTED")
    print(f" [URL] Dashboard URL: http://localhost:8000")
    print(f" [DB]  Excel Database: {EXCEL_PATH}")
    print(f" [PIN] Default Admin PIN: 1234")
    print("=" * 65 + "\n")
    uvicorn.run("backend:app", host="0.0.0.0", port=8000, reload=False)
