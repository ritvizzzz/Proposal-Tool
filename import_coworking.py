#!/usr/bin/env python3
"""
import_coworking.py — Sync coworking CSV into proposals.db
Updates existing centres with hotdesk_price + has_coworking=1
"""
import csv, re, sqlite3, os

CSV_PATH = os.path.join(os.path.dirname(__file__), '..', '..', 'Downloads', 'hubble_coworking_london 2.csv')
DB_PATH  = os.path.join(os.path.dirname(__file__), 'proposals.db')

def parse_price(s):
    m = re.search(r'[\d.]+', s or '')
    return int(float(m.group())) if m else None

rows = []
with open(CSV_PATH, newline='', encoding='utf-8-sig') as f:
    for row in csv.DictReader(f):
        rows.append(row)
print(f'CSV: {len(rows)} rows')

conn = sqlite3.connect(DB_PATH)
conn.row_factory = sqlite3.Row

# Ensure columns exist
for col in ['hotdesk_price INTEGER', 'has_coworking INTEGER DEFAULT 0']:
    try:
        conn.execute(f"ALTER TABLE centres ADD COLUMN {col.split()[0]} {' '.join(col.split()[1:])}")
        conn.commit()
    except Exception:
        pass

updated = 0
skipped = 0
for r in rows:
    hid = r['building_id'].strip()
    pd  = parse_price(r.get('price_day_gbp', ''))
    centre = conn.execute("SELECT id FROM centres WHERE hubble_id=?", (hid,)).fetchone()
    if centre:
        conn.execute("UPDATE centres SET hotdesk_price=?, has_coworking=1 WHERE id=?",
                     (pd, centre['id']))
        updated += 1
    else:
        skipped += 1

conn.commit()
conn.close()
print(f'Updated: {updated} centres with coworking pricing')
print(f'Skipped (not in DB): {skipped}')
print('Done ✅')
