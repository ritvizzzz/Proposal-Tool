"""Live office availability, pulled weekly from each operator's own source.

Priority (per operator): the operator's availability sheet/portal first; for
operators without one, the platforms (OfficeHub, then Hubble, then Instant
Offices -- not wired yet). An operator listed in OPERATOR_SOURCES relies on
its sheet alone, even when that sheet shows nothing available.

Each fetcher returns raw units: {building, unit, desks, desks_max, sqft,
price_month, status, available_from, product}. `building` is any text that
identifies the building (name, address, tab title) -- units are matched to
centres by postcode first, then by name words, within that operator's
centres only. Units that don't match are kept with centre_id NULL so they
can be reviewed at /admin/availability/unmatched.

Never store operators' internal comments (tenant names, deal notes): only
the fields above are kept.
"""
import io, os, re, csv, json, time, html as _html
from datetime import datetime, timezone
import requests
import openpyxl

UA = {'User-Agent': 'Mozilla/5.0'}

# Source links and passwords live in the AVAILABILITY_CONFIG environment
# variable (JSON), not here -- this repo is public and these are operators'
# broker-only sheets and portals. Keys: <operator>_sheet (Google Sheet id),
# mindspace_url, rivvia_url, workspace_password.
def _cfg(key):
    val = json.loads(os.environ.get('AVAILABILITY_CONFIG') or '{}').get(key)
    if not val:
        raise RuntimeError(f'{key} not set in AVAILABILITY_CONFIG')
    return val
_POSTCODE = re.compile(r'\b([A-Z]{1,2}\d[A-Z\d]?)\s*(\d[A-Z]{2})\b')
_OUTWARD = re.compile(r'\b([A-Z]{1,2}\d[A-Z\d]?)\b')


# ── Helpers ─────────────────────────────────────────────────────────────────

def _num(v):
    """First number in a cell: '£1,234.50' -> 1234.5, '16 - 24 + MTRs' -> 16."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    m = re.search(r'\d[\d,]*\.?\d*', str(v))
    return float(m.group(0).replace(',', '')) if m else None

def _desk_range(v):
    nums = [float(x) for x in re.findall(r'\d+(?:\.\d+)?', str(v or ''))[:2]]
    if not nums:
        return None, None
    lo = int(round(nums[0]))
    hi = int(round(nums[1])) if len(nums) > 1 and nums[1] > nums[0] and ('-' in str(v) or '–' in str(v) or 'to' in str(v)) else None
    return lo, hi

_GONE = ('occupied', 'under offer', 'not available', 'let ', 'reserved', 'taken', 'no -', 'sold', 'unavailable')
_NOW = ('now', 'immediate', 'available now', 'yes', 'vacant', 'available')

def _status(v):
    """('available' | 'soon' | None-for-not-available, available_from text)."""
    t = str(v or '').strip()
    low = t.lower().strip('* ')
    if not low:
        return None, None
    if any(g in low + ' ' for g in _GONE) and 'soon' not in low:
        return None, None
    if 'soon' in low:
        rest = re.sub(r'(?i)available\s*soon', '', t).strip(' -–:')
        return 'soon', rest or None
    if low in _NOW or low.startswith(('now', 'immediate', 'yes', 'available')):
        return 'available', None
    if isinstance(v, datetime):
        return 'soon', v.strftime('%d %b %Y')
    # A date or month ("March 2027", "16.04.27", "Jan 2027") = available from then
    if re.search(r'\d|jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|tbc', low):
        return 'soon', t
    return None, None

def _unit(building, unit, desks, sqft, price, avail, product=None):
    status, avail_from = _status(avail)
    price_v = _num(price)
    if not status or not price_v:
        return None
    lo, hi = _desk_range(desks)
    return {'building': str(building or '').strip(), 'unit': str(unit or '').strip()[:80],
            'desks': lo, 'desks_max': hi, 'sqft': int(_num(sqft)) if _num(sqft) else None,
            'price_month': int(round(price_v)), 'status': status,
            'available_from': avail_from, 'product': (str(product).strip()[:40] if product else None)}

def _gsheet_xlsx(sheet_id):
    r = requests.get(f'https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx', headers=UA, timeout=60)
    r.raise_for_status()
    return openpyxl.load_workbook(io.BytesIO(r.content), data_only=True)

# Column keywords, tried in order; first header cell containing one wins
_COLS = {
    'unit': ['office number', 'room number', 'resource name', 'office*', 'office', 'suite', 'space', 'ref', 'unit', 'name'],
    'desks': ['desk count', 'maximum capacity', 'desks', 'capacity', 'workstation', 'ws'],
    'sqft': ['size (sq ft)', 'sq. ft', 'sq ft', 'sqft', 'size (m2)', 'size'],
    'price': ['price per month', 'price/mo', 'price pm', 'price ex vat', 'monthly fee', 'price'],
    # 'avail' is the yes/no/when column; 'avail_date' a separate date column
    # (Oru has both: "YES - AVAILABLE NOW" and "DATE AVAILABLE")
    'avail_date': ['available from', 'date available'],
    'avail': ['availability', 'available', 'status'],
    'building': ['building', 'location'],
    'address': ['address'],
    'product': ['serviced or managed', 'product'],
    'floor': ['floor'],
}

def _header_map(row):
    cells = [str(c or '').strip().lower() for c in row]
    out = {}
    for key, words in _COLS.items():
        for w in words:
            hit = next((i for i, c in enumerate(cells) if c and w in c and i not in out.values()), None)
            if hit is not None:
                out[key] = hit
                break
    return out

def _parse_table_sheet(ws, context=''):
    """Generic: find the header row (has a price and an availability column),
    read units beneath it. `context` is extra building text (e.g. tab title)."""
    rows = list(ws.iter_rows(values_only=True))
    units, hdr, start = [], None, 0
    top_text = ' '.join(str(c) for r in rows[:12] for c in r if c)
    for i, r in enumerate(rows):
        h = _header_map(r)
        if 'avail' not in h and 'avail_date' in h:
            h['avail'] = h.pop('avail_date')  # the date column is the only availability signal
        if 'price' in h and 'avail' in h:
            hdr, start = h, i + 1
            break
    if not hdr:
        return units
    last_building, last_address, last_floor = '', '', ''
    # Sheets without a building/address column describe the building in the
    # title rows above the table (Spacemade: address + tab title)
    if 'building' in hdr or 'address' in hdr:
        top_text = ''
    for r in rows[start:]:
        get = lambda k: r[hdr[k]] if k in hdr and hdr[k] < len(r) else None
        if get('building'):
            last_building = str(get('building'))
            last_address = ''
        if get('address'):
            last_address = str(get('address'))
        if get('floor'):
            last_floor = str(get('floor'))
        unit = get('unit')
        if unit and 'sum of' in str(unit).lower():
            continue
        sqft = get('sqft')
        if 'sqft' in hdr and 'm2' in str(rows[start - 1][hdr['sqft']] or '').lower() and _num(sqft):
            sqft = _num(sqft) * 10.764
        label = ' '.join(x for x in (str(unit or '').strip(), '' if unit else last_floor) if x) or last_floor
        u = _unit(' '.join(x for x in (last_building, last_address, context, top_text) if x), label,
                  get('desks'), sqft, get('price'), get('avail'), get('product'))
        if u:
            when = get('avail_date')
            if when and not re.search(r'(?i)part', str(when)):
                st, txt = _status(when)
                if st == 'soon':
                    u['status'], u['available_from'] = 'soon', txt
            units.append(u)
    return units

def _parse_gsheet(sheet_id, tab_context=True):
    wb = _gsheet_xlsx(sheet_id)
    units = []
    for ws in wb.worksheets:
        if ws.sheet_state != 'visible':
            continue
        units += _parse_table_sheet(ws, ws.title if tab_context else '')
    return units


# ── Operator fetchers ───────────────────────────────────────────────────────

def fetch_oneder():
    return _parse_gsheet(_cfg('oneder_sheet'))

def fetch_oru():
    return _parse_gsheet(_cfg('oru_sheet'))

def fetch_spacemade():
    return _parse_gsheet(_cfg('spacemade_sheet'))

def fetch_vallist():
    return _parse_gsheet(_cfg('vallist_sheet'))

def fetch_workpad():
    # Only the first visible tab is the live schedule; it's renamed each week
    # ("18th November") and older ones are hidden.
    wb = _gsheet_xlsx(_cfg('workpad_sheet'))
    ws = next(w for w in wb.worksheets if w.sheet_state == 'visible')
    return _parse_table_sheet(ws)

def fetch_workspace():
    pw = _cfg('workspace_password')
    s = requests.Session()
    s.headers.update(UA)
    s.post('https://workspaceavailability.netlify.app/', data={'password': pw}, timeout=30, allow_redirects=False)
    r = s.get('https://workspaceavailability.netlify.app/availability-data.xlsx', timeout=60)
    if r.status_code != 200 or not r.content.startswith(b'PK'):
        raise RuntimeError(f'Workspace login/download failed ({r.status_code})')
    ws = openpyxl.load_workbook(io.BytesIO(r.content), read_only=True, data_only=True).worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    h = [str(c or '') for c in rows[0]]
    col = {name: h.index(name) for name in h}
    units = []
    for r in rows[1:]:
        g = lambda name: r[col[name]] if name in col else None
        utype = str(g('Unit Types') or '').lower()
        if not any(t in utype for t in ('office', 'studio', 'managed')):
            continue  # storage, workshops etc.
        u = _unit(' '.join(str(x) for x in (g('Property'), g('Address 1 (Property) (Property)'),
                                            g('Postal Code (Property) (Property)')) if x),
                  g('Reference'), None, g('Sq. Ft.'), g('ASK NetRSC PCM'),
                  str(g('Available Text') or '').replace('***', '').strip(), g('Unit Types'))
        if u:
            units.append(u)
    return units

def fetch_mindspace():
    s = requests.get(_cfg('mindspace_url'), headers=UA, timeout=30).text
    names = dict(re.findall(r'data-location-id="(\d+)"[^>]*>.*?location-name[^>]*>([^<]+)<', s, re.S))
    if not names:
        names = {str(i + 1): n for i, n in enumerate(re.findall(r'class="location-name[^"]*">([^<]+)<', s))}
    units = []
    for bid, body in re.findall(r'id="branch-(\d+)"(.*?)(?=id="branch-\d+"|$)', s, re.S):
        loc = re.sub(r'[^\w\s&-]', '', names.get(bid, '')).strip()
        for card in re.findall(r'class="product-card"(.*?)(?=class="product-card"|$)', body, re.S):
            clean = lambda x: re.sub(r'\s+', ' ', _html.unescape(re.sub(r'<[^>]+>', ' ', x or ''))).strip()
            name = re.search(r'class="product-name"[^>]*>(.*?)</h3>', card, re.S)
            avail = re.search(r'class="availability-tag[^"]*"[^>]*>(.*?)</div>', card, re.S)
            text = clean(card)
            price = re.search(r'Price\s*£\s*([\d,]+)', text)
            sqft = re.search(r'Size\s*([\d,]+)\s*Sq', text, re.I)
            desks = re.search(r'(\d+)\s*(?:Workstations?|Desks)', clean(name.group(1)) if name else text)
            if name and price:
                u = _unit('Mindspace ' + loc, clean(name.group(1)), desks and desks.group(1),
                          sqft and sqft.group(1), price.group(1), clean(avail.group(1)) if avail else '')
                if u:
                    units.append(u)
    return units

def fetch_rivvia():
    s = requests.get(_cfg('rivvia_url'), headers=UA, timeout=30).text
    units = []
    for m in re.finditer(r'<table.*?</table>', s, re.S):
        table = m.group(0)
        if '£' not in table:  # Amsterdam tables are in euros
            continue
        # The location's address sits in the card header above its table
        before = re.sub(r'\s+', ' ', _html.unescape(re.sub(r'<[^>]+>', ' ', s[max(0, m.start() - 12000):m.start()])))
        pcs = _POSTCODE.findall(before.upper())
        before = ' '.join(pcs[-1]) if pcs else before[-1500:]
        for tr in re.findall(r'<tr.*?</tr>', table, re.S)[1:]:
            cells = [re.sub(r'\s+', ' ', _html.unescape(re.sub(r'<[^>]+>', ' ', c))).strip()
                     for c in re.findall(r'<t[dh].*?</t[dh]>', tr, re.S)]
            if len(cells) < 5:
                continue
            ref, desks, fee, status = cells[0], cells[1], cells[3], cells[-1]
            if 'occupied' in status.lower():
                continue
            u = _unit('Rivvia ' + before, ref.replace('★', '').strip(), desks, None, fee,
                      'Now' if 'available' in status.lower() or 'offer' in status.lower() else status)
            if u:
                units.append(u)
    return units


# Operator key -> (fetcher, brand prefixes of its centres in the tool)
OPERATOR_SOURCES = {
    'workspace': (fetch_workspace, ['workspace -']),
    'oneder': (fetch_oneder, ['oneder']),
    'oru': (fetch_oru, ['oru space']),
    'spacemade': (fetch_spacemade, ['spacemade']),
    'vallist': (fetch_vallist, ['vallist']),
    'workpad': (fetch_workpad, ['workpad']),
    'mindspace': (fetch_mindspace, ['mindspace']),
    'rivvia': (fetch_rivvia, ['rivvia']),
}


# ── Matching units to centres ───────────────────────────────────────────────

_STOP = {'london', 'the', 'and', 'street', 'road', 'st', 'rd', 'house', 'floor', 'of', 'uk', 'office', 'offices',
         'city', 'old', 'place', 'square', 'park', 'lane', 'court', 'building', 'campus', 'east', 'west',
         'north', 'south', 'high', 'new', 'ground', 'first', 'second', 'unit', 'suite'}

def _words(s):
    return {w for w in re.findall(r'[a-z0-9]+', (s or '').lower()) if w not in _STOP and len(w) > 1}

def _postcodes(s):
    return {a + b for a, b in _POSTCODE.findall((s or '').upper())}

def _outwards(s):
    return set(_OUTWARD.findall((s or '').upper()))

# Operator's own name for a building -> our centre name, where neither the
# postcode nor the words line up
ALIASES = {
    'mindspace': {'shoreditch': 'Mindspace - Liverpool Street', 'old street': 'Mindspace - Churchill House'},
    'spacemade': {'hale house 76': 'Spacemade - Hale House 76 Portland Place',
                  'hale house 42': 'Spacemade - Hale House 42 Portland Place'},
    # Our address for Canalot has no street/postcode, so it ties with its W10 neighbours
    'workspace': {'canalot': 'Workspace - Canalot Studios'},
}

def match_centre(building, candidates, brand_words=(), aliases=None):
    """Best centre for a unit's building text among one operator's centres:
    alias, then a shared full postcode, then a postcode district only one
    centre has, then most shared name/address words (>=2, brand words
    ignored). Ties between centres with the same name pick the first."""
    low = (building or '').lower()
    for key, name in (aliases or {}).items():
        if key in low:
            hit = next((c for c in candidates if c['name'] == name), None)
            if hit:
                return hit['id']
    pcs = _postcodes(building)
    if pcs:
        hits = [c for c in candidates if pcs & _postcodes(c['address'])]
        if len(hits) == 1:
            return hits[0]['id']
        if hits:
            candidates = hits
    bw = _words(building) - set(brand_words)
    def score(c):
        return len(bw & (_words(c['name'] + ' ' + (c['address'] or '')) - set(brand_words)))
    scored = sorted(((score(c), c['id'], c['name'].lower()) for c in candidates), key=lambda t: (-t[0], t[1]))
    # One shared word is enough when no other centre of this operator has
    # any ("Sutton", "Finlaison", "Hammersmith")
    if scored and scored[0][0] == 1 and (len(scored) == 1 or scored[1][0] == 0):
        return scored[0][1]
    if scored and scored[0][0] >= 2:
        top = [t for t in scored if t[0] == scored[0][0]]
        if len(top) == 1 or len({t[2] for t in top}) == 1:
            return top[0][1]
        candidates = [c for c in candidates if c['id'] in {t[1] for t in top}]
    outs = _outwards(building)
    if outs:
        hits = [c for c in candidates if outs & _outwards(c['address'])]
        if len(hits) == 1:
            return hits[0]['id']
    return None


# ── Storage & runs ──────────────────────────────────────────────────────────

SCHEMA = """
CREATE TABLE IF NOT EXISTS centre_availability (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    operator TEXT NOT NULL,
    centre_id INTEGER,
    building TEXT,
    unit TEXT,
    desks INTEGER,
    desks_max INTEGER,
    sqft INTEGER,
    price_month INTEGER,
    status TEXT,
    available_from TEXT,
    product TEXT,
    source TEXT,
    fetched_at TEXT
);
CREATE INDEX IF NOT EXISTS idx_avail_centre ON centre_availability(centre_id);
CREATE TABLE IF NOT EXISTS availability_runs (
    operator TEXT PRIMARY KEY,
    last_run TEXT,
    ok INTEGER,
    units INTEGER,
    matched INTEGER,
    error TEXT
);
"""

def init_schema(conn):
    conn.executescript(SCHEMA)

def refresh_operator(get_db, key):
    """Re-fetch one operator; replaces its rows only if the fetch succeeded,
    so a broken source keeps last week's data instead of going blank."""
    fetcher, prefixes = OPERATOR_SOURCES[key]
    now = datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S')
    try:
        units = fetcher()
    except Exception as e:
        with get_db() as conn:
            conn.execute('INSERT OR REPLACE INTO availability_runs (operator, last_run, ok, units, matched, error) '
                         'VALUES (?,?,0,COALESCE((SELECT units FROM availability_runs WHERE operator=?),0),'
                         'COALESCE((SELECT matched FROM availability_runs WHERE operator=?),0),?)',
                         (key, now, key, key, str(e)[:500]))
        return {'operator': key, 'ok': False, 'error': str(e)}
    with get_db() as conn:
        cands = [dict(r) for r in conn.execute('SELECT id, name, address FROM centres').fetchall()
                 if any((r['name'] or '').lower().startswith(p) for p in prefixes)]
        brand_words = {w for p in prefixes for w in _words(p)}
        conn.execute('DELETE FROM centre_availability WHERE operator=?', (key,))
        matched = 0
        for u in units:
            cid = match_centre(u['building'], cands, brand_words, ALIASES.get(key))
            matched += cid is not None
            conn.execute('''INSERT INTO centre_availability (operator, centre_id, building, unit, desks, desks_max,
                            sqft, price_month, status, available_from, product, source, fetched_at)
                            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                         (key, cid, u['building'][:300], u['unit'], u['desks'], u['desks_max'], u['sqft'],
                          u['price_month'], u['status'], u['available_from'], u['product'], 'operator', now))
        conn.execute('INSERT OR REPLACE INTO availability_runs (operator, last_run, ok, units, matched, error) '
                     'VALUES (?,?,1,?,?,NULL)', (key, now, len(units), matched))
    return {'operator': key, 'ok': True, 'units': len(units), 'matched': matched}

def _due_operators(get_db, ok_age_days=7, failed_age_days=1):
    """Operators to refresh now: last success over a week ago, or last
    attempt failed over a day ago (retried daily rather than weekly)."""
    with get_db() as conn:
        rows = {r['operator']: dict(r) for r in conn.execute('SELECT * FROM availability_runs').fetchall()}
    now = datetime.now(timezone.utc).timestamp()
    due = []
    for key in OPERATOR_SOURCES:
        r = rows.get(key)
        if not r:
            due.append(key)
            continue
        age = now - datetime.strptime(r['last_run'], '%Y-%m-%d %H:%M:%S').replace(tzinfo=timezone.utc).timestamp()
        if age > (ok_age_days if r['ok'] else failed_age_days) * 86400:
            due.append(key)
    return due

def refresh_due(get_db):
    out = []
    for key in _due_operators(get_db):
        out.append(refresh_operator(get_db, key))
        time.sleep(1)
    return out

def refresh_all(get_db):
    out = []
    for key in OPERATOR_SOURCES:
        out.append(refresh_operator(get_db, key))
        time.sleep(1)
    return out
