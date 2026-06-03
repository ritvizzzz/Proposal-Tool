#!/usr/bin/env python3
"""
Import Hubble London inventory into all three tools:
  - Proposal Tool (SQLite DB)
  - Onboarding Tool (SQLite DB)
  - Map (netlify-deploy/index.html)

Usage: python3 import_hubble.py
"""

import csv, json, os, re, sqlite3, urllib.request, urllib.parse, time, ssl

BASE    = os.path.dirname(os.path.abspath(__file__))
REPO    = os.path.dirname(BASE)
INV_CSV = os.path.join(os.path.expanduser('~'), 'Downloads', 'hubble_london_inventory.csv')
COW_CSV = os.path.join(os.path.expanduser('~'), 'Downloads', 'hubble_coworking_london.csv')

PROPOSAL_DB   = os.path.join(BASE, 'proposals.db')
ONBOARD_DB    = os.path.join(REPO, 'onboarding-tool', 'library.db')
MAP_HTML      = os.path.join(REPO, 'netlify-deploy', 'index.html')

_SSL = ssl._create_unverified_context()

# ── Amenity tag normalisation ────────────────────────────────────────────────
_AMENITY_MAP = {
    'coffee-n-tea': 'Coffee & Tea', 'phone-booths': 'Phone Booths',
    'bike-storage': 'Bike Storage', 'showers': 'Showers',
    'pets-allowed': 'Pets Allowed', 'disabled-access': 'Disabled Access',
    'kitchen': 'Kitchen', 'printing': 'Printing',
    'event-space': 'Event Space', 'events-n-talks': 'Events & Talks',
    'secure-access': 'Secure Access', 'wifi': 'Wi-Fi',
    'mailing-address': 'Mailing Address', 'cleaning': 'Cleaning',
    'roof-terrace': 'Roof Terrace', 'reception': 'Reception',
    '24hr-access': '24/7 Access', 'lockers': 'Lockers',
    'furniture': 'Furnished', 'fruit-n-snacks': 'Snacks',
    'trading-address': 'Trading Address', 'breakout-space': 'Breakout Space',
    'utilities': 'Utilities Included', 'meeting-rooms': 'Meeting Rooms',
    'beer-n-wine': 'Beer & Wine', 'gym': 'Gym', 'cafe': 'Café',
    'onsite-cafe': 'Café', 'cafeteria': 'Café',
    # coworking facilities (already human-readable, pass through)
    'WiFi': 'Wi-Fi', 'Furnished': 'Furnished', 'Lockers': 'Lockers',
    'Showers': 'Showers', 'Meeting rooms': 'Meeting Rooms',
    '24/7 access': '24/7 Access', 'Kitchen': 'Kitchen',
    'Breakout space': 'Breakout Space', 'Bike storage': 'Bike Storage',
    'Event Space': 'Event Space', 'Mailing address': 'Mailing Address',
    'Disabled access': 'Disabled Access', 'Pets Allowed': 'Pets Allowed',
    'Beer N Wine': 'Beer & Wine', 'Café': 'Café', 'Cleaning': 'Cleaning',
    'Coffee & tea': 'Coffee & Tea', 'Snacks': 'Snacks',
    'Phone booths': 'Phone Booths', 'Printing': 'Printing',
    'Reception': 'Reception', 'Secure access': 'Secure Access',
    'Utilities included': 'Utilities Included', 'Trading address': 'Trading Address',
    'Events N Talks': 'Events & Talks', 'Roof terrace': 'Roof Terrace',
    'Gym': 'Gym',
}

def _normalise_amenities(raw_str):
    parts = [p.strip() for p in raw_str.split(',') if p.strip()]
    seen, out = set(), []
    for p in parts:
        nice = _AMENITY_MAP.get(p, p.replace('-', ' ').title())
        if nice not in seen:
            seen.add(nice)
            out.append(nice)
    return out


def _parse_price(price_str):
    """Return (amount_int_or_None, unit_str)."""
    if not price_str:
        return None, None
    m = re.search(r'[\£$]?([\d,]+)', price_str.replace(',', ''))
    if m:
        amt = int(m.group(1).replace(',', ''))
        if '/month' in price_str.lower() or 'month' in price_str.lower():
            return amt, 'MONTHLY'
        if '/day' in price_str.lower() or 'day' in price_str.lower():
            return amt, 'DAILY'
        return amt, 'MONTHLY'
    return None, None


def _extract_brand(name):
    if ' - ' in name:
        return name.split(' - ')[0].strip()
    return name.split()[0] if name else ''


def _clean(s):
    """Remove internal newlines/control chars from a string."""
    return re.sub(r'[\r\n\t]+', ' ', (s or '').strip()).strip()


def _parse_images(raw):
    if not raw or raw.strip() in ('', '[]'):
        return []
    try:
        return [u for u in json.loads(raw) if isinstance(u, str) and u.startswith('http')]
    except Exception:
        return []


# ── Geocoding via postcodes.io batch API ────────────────────────────────────
def geocode_postcodes(postcodes):
    """Return dict postcode -> (lat, lng). Missing ones get None."""
    result = {}
    unique = list(dict.fromkeys(pc.strip().upper() for pc in postcodes if pc and pc.strip()))
    print(f'  Geocoding {len(unique)} unique postcodes in batches of 100 …')
    for i in range(0, len(unique), 100):
        batch = unique[i:i+100]
        payload = json.dumps({'postcodes': batch}).encode()
        req = urllib.request.Request(
            'https://api.postcodes.io/postcodes',
            data=payload,
            headers={'Content-Type': 'application/json'},
        )
        try:
            with urllib.request.urlopen(req, timeout=15, context=_SSL) as r:
                data = json.loads(r.read())
            for item in data.get('result', []):
                pc = item.get('query', '').upper()
                res = item.get('result')
                if res and res.get('latitude'):
                    result[pc] = (float(res['latitude']), float(res['longitude']))
        except Exception as e:
            print(f'    Geocode batch {i}-{i+100} failed: {e}')
        time.sleep(0.2)
    found = sum(1 for v in result.values() if v)
    print(f'  Geocoded {found}/{len(unique)} postcodes.')
    return result


# ── Parse rows ───────────────────────────────────────────────────────────────
def parse_inventory(geo):
    rows = []
    with open(INV_CSV, newline='', encoding='utf-8-sig') as f:
        for r in csv.DictReader(f):
            pc = r.get('postcode', '').strip().upper()
            coords = geo.get(pc)
            if not coords:
                continue
            lat, lng = coords
            total_price, unit = _parse_price(r.get('price_per_desk', ''))
            offices_str = _clean(r.get('offices_count', '') or '')
            cap_min_raw = r.get('capacity_min', '').strip()
            cap_max_raw = r.get('capacity_max', '').strip()
            try:
                cap_min_int = int(cap_min_raw) if cap_min_raw else None
                cap_max_int = int(cap_max_raw) if cap_max_raw else None
            except Exception:
                cap_min_int = cap_max_int = None
            # Per-desk price = total price / min capacity
            if total_price and cap_min_int and cap_min_int > 0:
                price = round(total_price / cap_min_int)
            else:
                price = total_price
            rows.append({
                'hubble_id':      r.get('building_id', ''),
                'name':           _clean(r['name']),
                'address':        f"{_clean(r['address'])}, {pc}",
                'postcode':       pc,
                'brand':          _extract_brand(_clean(r['name'])),
                'lat':            lat,
                'lng':            lng,
                'transport':      _clean(r.get('nearest_tube', '')),
                'amenities':      _normalise_amenities(r.get('amenities', '')),
                'price':          price,
                'price_unit':     unit or 'MONTHLY',
                'seat_type':      'PRIVATE_OFFICE',
                'space_type':     _clean(r.get('product_type', 'Serviced')),
                'hubble_url':     _clean(r.get('hubble_url', '')),
                'capacity_min':   cap_min_int,
                'capacity_max':   cap_max_int,
                'offices_count':  offices_str,
                'images':         _parse_images(r.get('image_urls', '')),
            })
    return rows


def parse_coworking(geo, inv_building_ids, inv_names):
    """Only include coworking spaces NOT already in the inventory CSV."""
    rows = []
    skipped = 0
    with open(COW_CSV, newline='', encoding='utf-8-sig') as f:
        for r in csv.DictReader(f):
            # Skip if already in inventory (duplicate centre)
            if r['building_id'] in inv_building_ids:
                skipped += 1
                continue
            if _clean(r['name']).lower() in inv_names:
                skipped += 1
                continue
            pc = r.get('postcode', '').strip().upper()
            coords = geo.get(pc)
            if not coords:
                continue
            lat, lng = coords
            price, unit = _parse_price(r.get('price_day_gbp', ''))
            rows.append({
                'hubble_id':      r.get('building_id', ''),
                'name':           _clean(r['name']),
                'address':        f"{_clean(r['address'])}, {pc}",
                'postcode':       pc,
                'brand':          _extract_brand(_clean(r['name'])),
                'lat':            lat,
                'lng':            lng,
                'transport':      _clean(r.get('nearest_tube', '')),
                'amenities':      _normalise_amenities(r.get('facilities', '')),
                'price':          price,
                'price_unit':     'DAILY',
                'seat_type':      'DEDICATED_DESK',
                'space_type':     'Coworking',
                'hubble_url':     _clean(r.get('hubble_url', '')),
                'capacity_min':   None,
                'capacity_max':   None,
                'offices_count':  '',
                'images':         _parse_images(r.get('image_urls', '')),
            })
    print(f'  Coworking: {len(rows)} unique kept, {skipped} duplicates skipped')
    return rows


# ── Update Proposal Tool DB ──────────────────────────────────────────────────
def update_proposal_db(spaces):
    print(f'\nUpdating Proposal Tool DB ({PROPOSAL_DB}) …')
    conn = sqlite3.connect(PROPOSAL_DB)
    conn.row_factory = sqlite3.Row

    # Add hubble_id column if missing
    cols = [r[1] for r in conn.execute('PRAGMA table_info(centres)').fetchall()]
    if 'hubble_id' not in cols:
        conn.execute('ALTER TABLE centres ADD COLUMN hubble_id TEXT')

    # Remove all existing non-proposal centres (keep proposals referencing old IDs but clear centres)
    conn.execute('DELETE FROM centres')
    conn.execute('DELETE FROM centre_images')

    inserted = 0
    for sp in spaces:
        amenity_json = json.dumps(sp['amenities'])
        cur = conn.execute(
            '''INSERT INTO centres
               (name, address, city, space_type, brand, price_from, price_unit,
                seat_type, open_hours, amenities, transport, website,
                coordinates, source, hubble_id)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (
                sp['name'], sp['address'], 'london',
                sp['space_type'], sp['brand'],
                sp['price'], sp['price_unit'],
                sp['seat_type'],
                '9:00 AM – 6:00 PM',
                amenity_json,
                sp['transport'],
                sp['hubble_url'],
                f"{sp['lng']};{sp['lat']}",
                'hubble',
                sp['hubble_id'],
            )
        )
        cid = cur.lastrowid
        for i, url in enumerate(sp['images'][:6]):
            conn.execute(
                'INSERT INTO centre_images (centre_id, filename, is_primary, sort_order) VALUES (?,?,?,?)',
                (cid, url, 1 if i == 0 else 0, i)
            )
        inserted += 1

    conn.commit()
    conn.close()
    print(f'  Inserted {inserted} centres into proposal DB.')


# ── Update Onboarding Tool DB ────────────────────────────────────────────────
def update_onboarding_db(spaces):
    print(f'\nUpdating Onboarding Tool DB ({ONBOARD_DB}) …')
    conn = sqlite3.connect(ONBOARD_DB)

    # Ensure schema matches
    conn.execute('''CREATE TABLE IF NOT EXISTS library (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT, address TEXT, brand_slug TEXT, building_name TEXT,
        price INTEGER, lat REAL, lng REAL,
        postcode TEXT, space_type TEXT, transport TEXT, amenities TEXT,
        image_urls TEXT, hubble_url TEXT
    )''')
    # Add any missing columns
    cols = [r[1] for r in conn.execute('PRAGMA table_info(library)').fetchall()]
    for col, typ in [('postcode','TEXT'),('space_type','TEXT'),('transport','TEXT'),
                     ('amenities','TEXT'),('image_urls','TEXT'),('hubble_url','TEXT')]:
        if col not in cols:
            conn.execute(f'ALTER TABLE library ADD COLUMN {col} {typ}')

    conn.execute('DELETE FROM library')

    for sp in spaces:
        conn.execute(
            '''INSERT INTO library
               (name, address, brand_slug, building_name, price, lat, lng,
                postcode, space_type, transport, amenities, image_urls, hubble_url)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (
                sp['name'], sp['address'],
                sp['brand'].lower().replace(' ', '-'),
                sp['name'],
                sp['price'] or 0,
                sp['lat'], sp['lng'],
                sp['postcode'],
                sp['space_type'],
                sp['transport'],
                json.dumps(sp['amenities']),
                json.dumps(sp['images'][:4]),
                sp['hubble_url'],
            )
        )

    conn.commit()
    conn.close()
    print(f'  Inserted {len(spaces)} spaces into onboarding library.')


# Partnered brands (exact brand-name match, from myHQ partnership sheet)
# IWG group sub-brands + direct partners
_PARTNERED_EXACT = {
    'regus','spaces','hq','hq by regus','clubhouse','signature','homework',
    'runway east','xandy','x+y','mindspace','uncommon','workspace',
    'techspace','labs',
}
# Prefix match (brand starts with these)
_PARTNERED_PREFIX = ('regus','spaces','hq ','hq by','clubhouse','signature','homework',
                     'runway east','mindspace','uncommon','workspace','techspace','labs')

def _is_partnered(brand):
    b = brand.lower().strip()
    if b in _PARTNERED_EXACT:
        return True
    for p in _PARTNERED_PREFIX:
        if b.startswith(p):
            return True
    return False


# ── Update Map HTML ──────────────────────────────────────────────────────────
def update_map(spaces):
    print(f'\nUpdating Map ({MAP_HTML}) …')
    with open(MAP_HTML, encoding='utf-8') as f:
        html = f.read()

    # Build new WS array
    ws_list = []
    for i, sp in enumerate(spaces, 1):
        # pricing entry
        pricing = []
        if sp['price']:
            pricing.append({
                'type': 'Private Office' if sp['seat_type'] == 'PRIVATE_OFFICE' else 'Dedicated Desk',
                'cycle': sp['price_unit'],
                'amount': sp['price'],
                'capacityMin': sp.get('capacity_min'),
                'capacityMax': sp.get('capacity_max'),
            })

        # metro list from transport string
        metro = []
        tube_station = ''
        walk_time = ''
        if sp['transport']:
            m = re.match(r'^(.+?)\s*[—–-]\s*(.+)$', sp['transport'])
            if m:
                tube_station = m.group(1).strip()
                walk_time = m.group(2).strip()
                metro = [f"{tube_station} ({walk_time})"]
            else:
                metro = [sp['transport']]
                tube_station = sp['transport']

        ws_list.append({
            'id':            str(i),
            'hubble_id':     sp['hubble_id'],
            'name':          sp['name'],
            'address':       sp['address'],
            'postcode':      sp['postcode'],
            'city':          'London',
            'lat':           round(sp['lat'], 6),
            'lng':           round(sp['lng'], 6),
            'mapurl':        f"https://www.google.com/maps/search/{urllib.parse.quote(sp['address'])}",
            'about':         '',
            'brand':         sp['brand'],
            'spaceType':     sp['space_type'],
            'landmark':      '',
            'metro':         metro,
            'tubeStation':   tube_station,
            'walkTime':      walk_time,
            'transport':     sp['transport'],
            'totalSeats':    None,
            'capacityMin':   sp.get('capacity_min'),
            'capacityMax':   sp.get('capacity_max'),
            'officesCount':  sp.get('offices_count', ''),
            'amenities':     sp['amenities'],
            'pricing':       pricing,
            'photos':        sp['images'][:8],
            'hubble_url':    sp['hubble_url'],
            'meetingRooms':  [],
            'partnered':     _is_partnered(sp['brand']),
        })

    ws_json = json.dumps(ws_list, ensure_ascii=False, separators=(',', ':'))

    # Replace the existing const WS = [...]; block
    new_html = re.sub(
        r'const WS\s*=\s*\[.*?\];',
        f'const WS = {ws_json};',
        html,
        count=1,
        flags=re.DOTALL,
    )

    if new_html == html:
        print('  ERROR: could not find WS data block to replace!')
        return

    with open(MAP_HTML, 'w', encoding='utf-8') as f:
        f.write(new_html)
    print(f'  Map updated with {len(ws_list)} workspaces.')


# ── Main ─────────────────────────────────────────────────────────────────────
def main():
    print('=== Hubble London Import ===')
    print(f'Inventory: {INV_CSV}')
    print(f'Coworking: {COW_CSV}')

    # 1. Collect all postcodes
    all_postcodes = []
    with open(INV_CSV, newline='', encoding='utf-8-sig') as f:
        all_postcodes += [r['postcode'] for r in csv.DictReader(f)]
    with open(COW_CSV, newline='', encoding='utf-8-sig') as f:
        all_postcodes += [r['postcode'] for r in csv.DictReader(f)]

    # 2. Geocode
    geo = geocode_postcodes(all_postcodes)

    # 3. Parse — coworking only gets unique centres not in inventory
    print('\nParsing CSVs …')
    inv = parse_inventory(geo)
    # Build lookup sets from inventory to deduplicate coworking
    inv_ids   = {sp['hubble_id'] for sp in inv}
    inv_names = {sp['name'].lower() for sp in inv}
    cow = parse_coworking(geo, inv_ids, inv_names)
    all_spaces = inv + cow
    print(f'  {len(inv)} inventory (monthly) + {len(cow)} unique coworking (daily) = {len(all_spaces)} total')
    no_images = sum(1 for s in all_spaces if not s['images'])
    print(f'  {len(all_spaces) - no_images} have images, {no_images} have none')

    # 4. Update all tools
    update_proposal_db(all_spaces)
    update_onboarding_db(all_spaces)
    update_map(all_spaces)

    print('\n✓ Done. All three tools updated.')
    print('  Proposal Tool: restart Flask to see new centres.')
    print('  Map: open netlify-deploy/index.html to verify, then redeploy.')


if __name__ == '__main__':
    main()
