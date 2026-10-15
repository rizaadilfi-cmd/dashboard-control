-- ===============================================================
-- PT. NIFCO INDONESIA - MILKRUN DELIVERY CONTROL SYSTEM
-- SUPABASE POSTGRESQL DATABASE SCHEMA & REALTIME SETUP
-- ===============================================================

-- 1. TABEL JADWAL MILKRUN (17 Kolom Standar Nifco)
CREATE TABLE IF NOT EXISTS schedules (
    id TEXT NOT NULL,
    shipdate TEXT NOT NULL,
    customer TEXT NOT NULL,
    log_partner TEXT NOT NULL,
    cycle INTEGER NOT NULL,
    driver TEXT NOT NULL,
    phone TEXT,
    plan_time TEXT NOT NULL,
    act_time TEXT,
    status TEXT NOT NULL DEFAULT 'PENDING',
    custom_remark TEXT,
    sj_date_qty TEXT DEFAULT '',
    tgl_sj TEXT DEFAULT '',
    jml_sj TEXT DEFAULT '',
    pallet TEXT DEFAULT '',
    box TEXT DEFAULT '',
    pallet_box TEXT DEFAULT '',
    return_pallet TEXT DEFAULT '',
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    PRIMARY KEY (id, shipdate)
);


-- 2. TABEL MASTER DRIVER
CREATE TABLE IF NOT EXISTS drivers (
    name TEXT PRIMARY KEY,
    phone TEXT NOT NULL,
    partner TEXT NOT NULL,
    status TEXT DEFAULT 'AKTIF',
    last_active TEXT,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- 3. TABEL PENGATURAN (GEOFENCE & WEBHOOK n8n)
CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- 4. INSERT DATA AWAL (SEED DRIVER & SETTINGS)
INSERT INTO drivers (name, phone, partner, status)
VALUES 
    ('ASEP', '081234567890', 'DARSINDO', 'AKTIF'),
    ('POSTER', '081298765432', 'SYNCRUM', 'AKTIF'),
    ('YADI', '081566778899', 'TTNT', 'AKTIF'),
    ('DEDI', '081788990011', 'SELOG', 'AKTIF'),
    ('NAPIN', '081922334455', 'SYNCRUM', 'AKTIF')
ON CONFLICT (name) DO NOTHING;

INSERT INTO settings (key, value)
VALUES 
    ('geofence', '{"name":"PT. Nifco Indonesia","lat":-6.357800,"lng":107.289000,"radius":500}'),
    ('webhook', '{"n8n_url":"","status":"ENABLED"}')
ON CONFLICT (key) DO NOTHING;

-- 5. AKTIFKAN SUPABASE REALTIME (Sinkronisasi Otomatis Antar Device)
ALTER TABLE schedules REPLICA IDENTITY FULL;
ALTER TABLE drivers REPLICA IDENTITY FULL;
ALTER TABLE settings REPLICA IDENTITY FULL;

ALTER PUBLICATION supabase_realtime ADD TABLE schedules;
ALTER PUBLICATION supabase_realtime ADD TABLE drivers;
ALTER PUBLICATION supabase_realtime ADD TABLE settings;

-- 6. KEBIJAKAN AKSES PUBLIK (ROW LEVEL SECURITY)
-- Mengizinkan dashboard frontend membaca dan mengupdate data secara aman
ALTER TABLE schedules ENABLE ROW LEVEL SECURITY;
ALTER TABLE drivers ENABLE ROW LEVEL SECURITY;
ALTER TABLE settings ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Allow public read-write for schedules" ON schedules FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "Allow public read-write for drivers" ON drivers FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "Allow public read-write for settings" ON settings FOR ALL USING (true) WITH CHECK (true);
