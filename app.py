import os, sqlite3, json, shutil, base64, io, re
from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for
from werkzeug.utils import secure_filename
from PIL import Image
import openpyxl

app = Flask(__name__)
app.secret_key = 'myhq-proposal-tool-secret'
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, 'proposals.db')
UPLOADS = os.path.join(BASE_DIR, 'uploads')
CENTRE_IMAGES = os.path.join(UPLOADS, 'centres')
PROPOSAL_FILES = os.path.join(UPLOADS, 'proposals')
TEMPLATE_FILES = os.path.join(UPLOADS, 'templates')

os.makedirs(CENTRE_IMAGES, exist_ok=True)
os.makedirs(PROPOSAL_FILES, exist_ok=True)
os.makedirs(TEMPLATE_FILES, exist_ok=True)

# ── Database ────────────────────────────────────────────────────────────────

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with get_db() as conn:
        conn.executescript('''
        CREATE TABLE IF NOT EXISTS centres (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            address TEXT,
            city TEXT,
            about TEXT,
            space_type TEXT,
            brand TEXT,
            price_from INTEGER,
            price_unit TEXT DEFAULT 'MONTHLY',
            seat_type TEXT,
            open_hours TEXT DEFAULT '9:00 AM – 6:00 PM',
            amenities TEXT,
            transport TEXT,
            website TEXT,
            map_url TEXT,
            coordinates TEXT,
            why_recommend TEXT,
            source TEXT DEFAULT 'manual',
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS centre_images (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            centre_id INTEGER NOT NULL,
            filename TEXT NOT NULL,
            label TEXT,
            is_primary INTEGER DEFAULT 0,
            sort_order INTEGER DEFAULT 0,
            FOREIGN KEY (centre_id) REFERENCES centres(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS proposals (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT,
            template TEXT DEFAULT 'london',
            client_name TEXT,
            client_company TEXT,
            client_email TEXT,
            client_location TEXT,
            team_size TEXT,
            space_type TEXT,
            area_required TEXT,
            budget TEXT,
            duration TEXT,
            selected_centres TEXT DEFAULT '[]',
            manual_centres TEXT DEFAULT '[]',
            status TEXT DEFAULT 'draft',
            pptx_filename TEXT,
            pdf_filename TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        );
        ''')
    # migrate: add pdf_filename if missing
    with get_db() as conn:
        cols = [r[1] for r in conn.execute("PRAGMA table_info(proposals)").fetchall()]
        if 'pdf_filename' not in cols:
            conn.execute("ALTER TABLE proposals ADD COLUMN pdf_filename TEXT")

init_db()

# ── Helpers ─────────────────────────────────────────────────────────────────

ALLOWED = {'png','jpg','jpeg','gif','webp'}

def allowed(filename):
    return '.' in filename and filename.rsplit('.',1)[1].lower() in ALLOWED

def save_image(file_obj, folder, name_prefix='img'):
    filename = secure_filename(file_obj.filename)
    ext = filename.rsplit('.',1)[-1].lower() if '.' in filename else 'jpg'
    out_name = f"{name_prefix}_{os.urandom(6).hex()}.{ext}"
    out_path = os.path.join(folder, out_name)
    img = Image.open(file_obj)
    img = img.convert('RGB')
    img.save(out_path, quality=90)
    return out_name

def save_image_bytes(data_bytes, folder, name_prefix='img'):
    out_name = f"{name_prefix}_{os.urandom(6).hex()}.jpg"
    out_path = os.path.join(folder, out_name)
    img = Image.open(io.BytesIO(data_bytes)).convert('RGB')
    img.save(out_path, quality=90)
    return out_name

def centre_image_dir(centre_id):
    d = os.path.join(CENTRE_IMAGES, str(centre_id))
    os.makedirs(d, exist_ok=True)
    return d

# ── Excel import ────────────────────────────────────────────────────────────

def import_excel(filepath):
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb['workspace']
    headers = [c for c in next(ws.iter_rows(values_only=True))]
    h = {v: i for i, v in enumerate(headers) if v}

    pricing = {}
    for sheet in ['dedicated_price_plans', 'flexi_price_plans', 'vo_price_plans']:
        if sheet not in wb.sheetnames:
            continue
        pws = wb[sheet]
        ph = None
        for row in pws.iter_rows(values_only=True):
            if ph is None:
                ph = {v: i for i, v in enumerate(row) if v}
                continue
            wid = row[ph.get('workspace_identifier', 0)]
            if not wid:
                continue
            amt = row[ph.get('amount', ph.get('pricePerSeat', 6))]
            unit = row[ph.get('paymentCycle', ph.get('unit', 3))]  # paymentCycle = MONTHLY/DAILY
            stype = row[ph.get('seatType', ph.get('type', 4))]
            if wid not in pricing and amt:
                pricing[wid] = {'amount': amt, 'unit': unit or 'MONTHLY', 'seat_type': stype}

    amenities_map = {}
    for sheet in ['dedicated_amenities', 'flexi_amenities', 'vo_amenities', 'meeting_room_amenities']:
        if sheet not in wb.sheetnames:
            continue
        aws = wb[sheet]
        ah = None
        for row in aws.iter_rows(values_only=True):
            if ah is None:
                ah = {v: i for i, v in enumerate(row) if v}
                continue
            wid = row[ah.get('workspace_identifier', 0)]
            slug = row[ah.get('amenity_slug', 1)]
            if wid and slug:
                amenities_map.setdefault(wid, [])
                if slug not in amenities_map[wid]:
                    amenities_map[wid].append(slug)

    added = 0
    skipped = 0
    with get_db() as conn:
        for row in ws.iter_rows(values_only=True):
            if row[0] == 'workspace_identifier' or row[0] is None:
                continue
            wid = row[h.get('workspace_identifier', 0)]
            name = row[h.get('name', 1)]
            if not name:
                continue
            existing = conn.execute('SELECT id FROM centres WHERE name=?', (name,)).fetchone()
            if existing:
                skipped += 1
                continue
            addr = row[h.get('address', 4)]
            city = row[h.get('city_slug', 5)]
            about = row[h.get('about', 8)]
            stype = row[h.get('spaceType', 9)]
            brand = row[h.get('workspaceBrand_slug', 11)]
            transport_raw = row[h.get('directions_connectivityDetails', 12)]
            bus = row[h.get('directions_nearestBus', 14)]
            coords = row[h.get('loc_coordinates', 6)]
            mapurl = row[h.get('mapurl', 7)]

            transport_lines = []
            if transport_raw:
                for part in str(transport_raw).split(','):
                    part = part.strip()
                    segments = part.split(':')
                    if len(segments) >= 2:
                        transport_lines.append(segments[1].strip())
            if bus:
                transport_lines.append(f'Bus: {bus}')
            transport = ', '.join(transport_lines) if transport_lines else ''

            price_info = pricing.get(wid, {})
            conn.execute('''INSERT INTO centres
                (name,address,city,about,space_type,brand,price_from,price_unit,seat_type,amenities,transport,map_url,coordinates,source)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                (name, addr, city, about, stype, brand,
                 price_info.get('amount'), price_info.get('unit','MONTHLY'),
                 price_info.get('seat_type'),
                 json.dumps(amenities_map.get(wid, [])),
                 transport, mapurl, coords, 'excel'))
            added += 1
    return added, skipped

# ── Routes: Centres ─────────────────────────────────────────────────────────

@app.route('/')
def index():
    with get_db() as conn:
        centres = conn.execute('SELECT id,name,city,space_type,brand,price_from,price_unit FROM centres ORDER BY name').fetchall()
        proposals = conn.execute('SELECT id,title,client_name,client_company,status,created_at FROM proposals ORDER BY created_at DESC LIMIT 5').fetchall()
        centre_count = conn.execute('SELECT COUNT(*) FROM centres').fetchone()[0]
    return render_template('index.html', centres=centres, proposals=proposals, centre_count=centre_count)

@app.route('/centres')
def centres_list():
    q = request.args.get('q','').strip()
    city = request.args.get('city','')
    space_type = request.args.get('space_type','')
    with get_db() as conn:
        query = 'SELECT c.*, (SELECT filename FROM centre_images WHERE centre_id=c.id AND is_primary=1 LIMIT 1) as primary_image FROM centres c WHERE 1=1'
        params = []
        if q:
            query += ' AND (c.name LIKE ? OR c.address LIKE ?)'
            params += [f'%{q}%', f'%{q}%']
        if city:
            query += ' AND c.city=?'
            params.append(city)
        if space_type:
            query += ' AND c.space_type=?'
            params.append(space_type)
        query += ' ORDER BY c.name'
        centres = conn.execute(query, params).fetchall()
        cities = [r[0] for r in conn.execute('SELECT DISTINCT city FROM centres WHERE city IS NOT NULL ORDER BY city').fetchall()]
        space_types = [r[0] for r in conn.execute('SELECT DISTINCT space_type FROM centres WHERE space_type IS NOT NULL ORDER BY space_type').fetchall()]
    return render_template('centres.html', centres=centres, cities=cities, space_types=space_types, q=q, city=city, space_type=space_type)

@app.route('/centres/<int:cid>')
def centre_detail(cid):
    with get_db() as conn:
        centre = conn.execute('SELECT * FROM centres WHERE id=?', (cid,)).fetchone()
        if not centre:
            return 'Not found', 404
        images = conn.execute('SELECT * FROM centre_images WHERE centre_id=? ORDER BY is_primary DESC, sort_order', (cid,)).fetchall()
    return render_template('centre_detail.html', centre=centre, images=images)

@app.route('/centres/add', methods=['POST'])
def centre_add():
    data = request.json or request.form
    amenities = data.get('amenities', '[]')
    if isinstance(amenities, str):
        try:
            json.loads(amenities)
        except:
            amenities = json.dumps([a.strip() for a in amenities.split(',') if a.strip()])
    with get_db() as conn:
        cur = conn.execute('''INSERT INTO centres
            (name,address,city,about,space_type,brand,price_from,price_unit,open_hours,amenities,transport,website,map_url,why_recommend,source)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (data.get('name'), data.get('address'), data.get('city'),
             data.get('about'), data.get('space_type'), data.get('brand'),
             data.get('price_from') or None, data.get('price_unit','MONTHLY'),
             data.get('open_hours','9:00 AM – 6:00 PM'),
             amenities, data.get('transport'), data.get('website'),
             data.get('map_url'), data.get('why_recommend'), 'manual'))
        return jsonify({'id': cur.lastrowid, 'ok': True})

@app.route('/centres/<int:cid>/update', methods=['POST'])
def centre_update(cid):
    data = request.json or request.form
    fields = ['name','address','city','about','space_type','brand','price_from','price_unit',
              'open_hours','amenities','transport','website','map_url','why_recommend']
    sets = ', '.join(f'{f}=?' for f in fields if f in data)
    vals = [data[f] for f in fields if f in data] + [cid]
    if sets:
        with get_db() as conn:
            conn.execute(f'UPDATE centres SET {sets} WHERE id=?', vals)
    return jsonify({'ok': True})

@app.route('/centres/<int:cid>/delete', methods=['POST','DELETE'])
def centre_delete(cid):
    img_dir = os.path.join(CENTRE_IMAGES, str(cid))
    if os.path.isdir(img_dir):
        shutil.rmtree(img_dir)
    with get_db() as conn:
        conn.execute('DELETE FROM centre_images WHERE centre_id=?', (cid,))
        conn.execute('DELETE FROM centres WHERE id=?', (cid,))
    return jsonify({'ok': True})

@app.route('/centres/<int:cid>/upload-image', methods=['POST'])
def centre_upload_image(cid):
    img_dir = centre_image_dir(cid)
    filenames = []

    # Handle file upload
    for f in request.files.getlist('images'):
        if f and allowed(f.filename):
            name = save_image(f, img_dir, f'centre_{cid}')
            filenames.append(name)

    # Handle base64 paste
    for paste_data in request.form.getlist('paste_data'):
        if paste_data.startswith('data:image'):
            header, b64 = paste_data.split(',', 1)
            raw = base64.b64decode(b64)
            name = save_image_bytes(raw, img_dir, f'centre_{cid}_paste')
            filenames.append(name)

    with get_db() as conn:
        existing = conn.execute('SELECT COUNT(*) FROM centre_images WHERE centre_id=?', (cid,)).fetchone()[0]
        for i, fname in enumerate(filenames):
            is_primary = 1 if (existing == 0 and i == 0) else 0
            conn.execute('INSERT INTO centre_images (centre_id, filename, is_primary, sort_order) VALUES (?,?,?,?)',
                         (cid, fname, is_primary, existing + i))

    images = []
    with get_db() as conn:
        rows = conn.execute('SELECT * FROM centre_images WHERE centre_id=? ORDER BY is_primary DESC, sort_order', (cid,)).fetchall()
        for r in rows:
            images.append({'id': r['id'], 'filename': r['filename'], 'is_primary': r['is_primary']})
    return jsonify({'ok': True, 'images': images})

@app.route('/centres/<int:cid>/set-primary-image/<int:img_id>', methods=['POST'])
def set_primary_image(cid, img_id):
    with get_db() as conn:
        conn.execute('UPDATE centre_images SET is_primary=0 WHERE centre_id=?', (cid,))
        conn.execute('UPDATE centre_images SET is_primary=1 WHERE id=? AND centre_id=?', (img_id, cid))
    return jsonify({'ok': True})

@app.route('/centres/<int:cid>/delete-image/<int:img_id>', methods=['POST','DELETE'])
def delete_image(cid, img_id):
    with get_db() as conn:
        row = conn.execute('SELECT filename FROM centre_images WHERE id=? AND centre_id=?', (img_id, cid)).fetchone()
        if row:
            fpath = os.path.join(CENTRE_IMAGES, str(cid), row['filename'])
            if os.path.exists(fpath):
                os.remove(fpath)
            conn.execute('DELETE FROM centre_images WHERE id=?', (img_id,))
    return jsonify({'ok': True})

@app.route('/centres/<int:cid>/fetch-photos', methods=['POST'])
def centre_fetch_photos(cid):
    import urllib.request, urllib.parse
    data = request.json or {}
    api_key = data.get('api_key') or os.environ.get('GOOGLE_MAPS_API_KEY', '')
    if not api_key:
        return jsonify({'error': 'No API key provided'}), 400

    with get_db() as conn:
        centre = conn.execute('SELECT name, address FROM centres WHERE id=?', (cid,)).fetchone()
    if not centre:
        return jsonify({'error': 'Centre not found'}), 404

    name = centre['name'] or ''
    address = centre['address'] or ''
    query = f"{name} {address}".strip()

    # 1. Find place
    search_url = (
        'https://maps.googleapis.com/maps/api/place/findplacefromtext/json?'
        + urllib.parse.urlencode({'input': query, 'inputtype': 'textquery', 'fields': 'photos', 'key': api_key})
    )
    try:
        with urllib.request.urlopen(search_url, timeout=10) as r:
            result = json.loads(r.read().decode())
    except Exception as e:
        return jsonify({'error': f'Places API error: {e}'}), 500

    candidates = result.get('candidates', [])
    if not candidates:
        return jsonify({'ok': True, 'added': 0, 'note': 'No place found'})

    photos = candidates[0].get('photos', [])[:4]
    if not photos:
        return jsonify({'ok': True, 'added': 0, 'note': 'No photos for this place'})

    img_dir = centre_image_dir(cid)
    added = 0
    with get_db() as conn:
        existing_count = conn.execute('SELECT COUNT(*) FROM centre_images WHERE centre_id=?', (cid,)).fetchone()[0]
        for i, photo in enumerate(photos):
            ref = photo.get('photo_reference')
            if not ref:
                continue
            photo_url = (
                'https://maps.googleapis.com/maps/api/place/photo?'
                + urllib.parse.urlencode({'maxwidth': 800, 'photo_reference': ref, 'key': api_key})
            )
            try:
                req = urllib.request.Request(photo_url, headers={'User-Agent': 'myHQ-proposal-tool'})
                with urllib.request.urlopen(req, timeout=15) as r:
                    raw = r.read()
                out_name = save_image_bytes(raw, img_dir, f'centre_{cid}_gp')
                is_primary = 1 if (existing_count == 0 and i == 0) else 0
                conn.execute(
                    'INSERT INTO centre_images (centre_id, filename, is_primary, sort_order) VALUES (?,?,?,?)',
                    (cid, out_name, is_primary, existing_count + i)
                )
                added += 1
            except Exception:
                continue
    return jsonify({'ok': True, 'added': added})

@app.route('/centres/<int:cid>/import-url', methods=['POST'])
def centre_import_url(cid):
    """Fetch an image from a URL and save it as a centre image."""
    import urllib.request, ssl
    data = request.json or {}
    url = (data.get('url') or '').strip()
    if not url:
        return jsonify({'error': 'No URL provided'}), 400

    with get_db() as conn:
        if not conn.execute('SELECT 1 FROM centres WHERE id=?', (cid,)).fetchone():
            return jsonify({'error': 'Centre not found'}), 404

    ctx = ssl._create_unverified_context()
    try:
        req = urllib.request.Request(url, headers={
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36',
            'Referer': url,
        })
        with urllib.request.urlopen(req, timeout=15, context=ctx) as r:
            raw = r.read()
    except Exception as e:
        return jsonify({'error': f'Failed to fetch URL: {e}'}), 400

    if len(raw) < 1000:
        return jsonify({'error': 'Response too small — not a valid image'}), 400

    img_dir = centre_image_dir(cid)
    try:
        fname = save_image_bytes(raw, img_dir, f'centre_{cid}_url')
    except Exception as e:
        return jsonify({'error': f'Could not save image: {e}'}), 500

    with get_db() as conn:
        existing = conn.execute('SELECT COUNT(*) FROM centre_images WHERE centre_id=?', (cid,)).fetchone()[0]
        is_primary = 1 if existing == 0 else 0
        conn.execute(
            'INSERT INTO centre_images (centre_id, filename, is_primary, sort_order) VALUES (?,?,?,?)',
            (cid, fname, is_primary, existing)
        )
        rows = conn.execute(
            'SELECT * FROM centre_images WHERE centre_id=? ORDER BY is_primary DESC, sort_order', (cid,)
        ).fetchall()
        images = [{'id': r['id'], 'filename': r['filename'], 'is_primary': r['is_primary']} for r in rows]

    return jsonify({'ok': True, 'images': images})


@app.route('/centres/import-excel', methods=['POST'])
def import_excel_route():
    f = request.files.get('file')
    if not f:
        return jsonify({'error': 'No file'}), 400
    tmp = os.path.join(UPLOADS, 'tmp_import.xlsx')
    f.save(tmp)
    try:
        added, skipped = import_excel(tmp)
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
    return jsonify({'ok': True, 'added': added, 'skipped': skipped})

@app.route('/centres/add-page')
def centres_add_page():
    return redirect(url_for('centres_list'))

@app.route('/centres/api/search')
def api_centre_search():
    q = request.args.get('q','').strip()
    with get_db() as conn:
        rows = conn.execute(
            '''SELECT c.id, c.name, c.address, c.city, c.space_type, c.price_from, c.price_unit,
               (SELECT filename FROM centre_images WHERE centre_id=c.id AND is_primary=1 LIMIT 1) as primary_image
               FROM centres c
               WHERE c.name LIKE ? OR c.address LIKE ?
               ORDER BY c.name LIMIT 30''',
            (f'%{q}%', f'%{q}%')).fetchall()
    return jsonify([dict(r) for r in rows])

@app.route('/centres/api/<int:cid>')
def api_centre_detail(cid):
    with get_db() as conn:
        c = conn.execute('SELECT * FROM centres WHERE id=?', (cid,)).fetchone()
        if not c:
            return jsonify({'error': 'Not found'}), 404
        imgs = conn.execute('SELECT id, filename, label, is_primary, sort_order FROM centre_images WHERE centre_id=? ORDER BY is_primary DESC, sort_order', (cid,)).fetchall()
    centre = dict(c)
    img_list = [dict(i) for i in imgs]
    # images: list of URLs (for builder display)
    centre['images'] = [f'/centre-image/{cid}/{i["filename"]}' for i in img_list]
    # images_meta: full records (for image editor modal)
    centre['images_meta'] = img_list
    return jsonify(centre)

# ── Routes: Image serving ────────────────────────────────────────────────────

@app.route('/centre-image/<int:cid>/<path:filename>')
def serve_centre_image(cid, filename):
    img_dir = os.path.join(CENTRE_IMAGES, str(cid))
    return send_file(os.path.join(img_dir, filename))

# ── Routes: Image proxy (for manual/non-DB centres) ─────────────────────────

@app.route('/api/fetch-image-b64', methods=['POST'])
def api_fetch_image_b64():
    """Fetch an image from a URL and return it as a base64 data URI."""
    data = request.json or {}
    url = (data.get('url') or '').strip()
    if not url or not url.startswith('http'):
        return jsonify({'error': 'Invalid URL'}), 400
    try:
        import urllib.request, base64 as _b64
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=10) as resp:
            content_type = resp.headers.get('Content-Type', 'image/jpeg').split(';')[0].strip()
            raw = resp.read()
        b64 = f'data:{content_type};base64,{_b64.b64encode(raw).decode()}'
        return jsonify({'b64': b64})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ── Routes: Map data API ─────────────────────────────────────────────────────

@app.route('/api/map-data')
def api_map_data():
    with get_db() as conn:
        centres = conn.execute('SELECT * FROM centres ORDER BY name').fetchall()
        result = []
        for c in centres:
            cid = c['id']
            imgs = conn.execute(
                'SELECT filename FROM centre_images WHERE centre_id=? ORDER BY is_primary DESC, sort_order',
                (cid,)).fetchall()
            photos = [
                f'http://localhost:5001/centre-image/{cid}/{row["filename"]}'
                for row in imgs
            ]
            entry = dict(c)
            entry['photos'] = photos
            result.append(entry)
    resp = jsonify(result)
    resp.headers['Access-Control-Allow-Origin'] = '*'
    return resp

# ── Routes: Proposals ───────────────────────────────────────────────────────

@app.route('/projects')
def projects_list():
    with get_db() as conn:
        proposals = conn.execute('SELECT * FROM proposals ORDER BY created_at DESC').fetchall()
    # Group by client_company (fall back to client_name then 'Unknown')
    groups = {}
    for p in proposals:
        key = p['client_company'] or p['client_name'] or 'Unknown'
        groups.setdefault(key, []).append(p)
    # Sort groups: most recently updated first
    sorted_groups = sorted(groups.items(), key=lambda kv: kv[1][0]['created_at'], reverse=True)
    return render_template('projects.html', groups=sorted_groups)

@app.route('/proposals')
def proposals_list():
    with get_db() as conn:
        proposals = conn.execute('SELECT * FROM proposals ORDER BY created_at DESC').fetchall()
    return render_template('proposals.html', proposals=proposals)

@app.route('/proposals/new', methods=['GET','POST'])
def proposal_new():
    if request.method == 'POST':
        data = request.json
        with get_db() as conn:
            cur = conn.execute('''INSERT INTO proposals
                (title,template,client_name,client_company,client_email,client_location,
                 team_size,space_type,area_required,budget,duration)
                VALUES (?,?,?,?,?,?,?,?,?,?,?)''',
                (data.get('title','New Proposal'),
                 data.get('template','london'),
                 data.get('client_name'), data.get('client_company'),
                 data.get('client_email'), data.get('client_location'),
                 data.get('team_size'), data.get('space_type'),
                 data.get('area_required'), data.get('budget'),
                 data.get('duration')))
            pid = cur.lastrowid
        return jsonify({'ok': True, 'id': pid})
    # Handle map_ws param (base64 JSON workspace array from London map)
    preselected_ids = []
    map_workspaces = []
    map_ws_b64 = request.args.get('map_ws', '').strip()
    if map_ws_b64:
        try:
            import base64
            decoded = base64.b64decode(map_ws_b64).decode('utf-8')
            spaces = json.loads(decoded)
            if isinstance(spaces, list):
                for sp in spaces:
                    if not isinstance(sp, dict) or not sp.get('name'):
                        continue
                    # Check if this workspace exists in the DB
                    with get_db() as conn:
                        row = conn.execute(
                            'SELECT id FROM centres WHERE LOWER(name) LIKE LOWER(?) LIMIT 1',
                            (f'%{sp["name"]}%',)
                        ).fetchone()
                    if row:
                        preselected_ids.append(row['id'])
                    else:
                        # Workspace not in DB — pass as manual entry
                        if isinstance(sp.get('amenities'), list):
                            sp['amenities'] = json.dumps(sp['amenities'])
                        map_workspaces.append(sp)
        except Exception:
            pass
    # Legacy fallback: map_names param
    elif request.args.get('map_names', '').strip():
        map_names = request.args.get('map_names', '').strip()
        names = [n.strip() for n in map_names.split(',') if n.strip()]
        with get_db() as conn:
            for name in names:
                row = conn.execute(
                    'SELECT id FROM centres WHERE LOWER(name) LIKE LOWER(?) LIMIT 1',
                    (f'%{name}%',)
                ).fetchone()
                if row:
                    preselected_ids.append(row['id'])
    return render_template('builder.html', step=1, proposal=None,
                           preselected_ids=preselected_ids,
                           map_workspaces=map_workspaces)

@app.route('/proposals/<int:pid>')
def proposal_builder(pid):
    with get_db() as conn:
        proposal = conn.execute('SELECT * FROM proposals WHERE id=?', (pid,)).fetchone()
        if not proposal:
            return 'Not found', 404
    return render_template('builder.html', step=1, proposal=dict(proposal), preselected_ids=[], map_workspaces=[])

@app.route('/proposals/<int:pid>/update', methods=['POST'])
def proposal_update(pid):
    data = request.json
    allowed_fields = ['title','template','client_name','client_company','client_email',
                      'client_location','team_size','space_type','area_required','budget',
                      'duration','selected_centres','manual_centres','status']
    sets = ', '.join(f'{f}=?' for f in allowed_fields if f in data)
    vals = [data[f] for f in allowed_fields if f in data] + [pid]
    if sets:
        with get_db() as conn:
            conn.execute(f'UPDATE proposals SET {sets} WHERE id=?', vals)
    return jsonify({'ok': True})

@app.route('/proposals/<int:pid>/delete', methods=['POST','DELETE'])
def proposal_delete(pid):
    with get_db() as conn:
        row = conn.execute('SELECT pptx_filename, pdf_filename FROM proposals WHERE id=?', (pid,)).fetchone()
        if row:
            for fname in [row['pptx_filename'], row['pdf_filename']]:
                if fname:
                    fpath = os.path.join(PROPOSAL_FILES, fname)
                    if os.path.exists(fpath): os.remove(fpath)
        conn.execute('DELETE FROM proposals WHERE id=?', (pid,))
    return jsonify({'ok': True})

@app.route('/proposals/<int:pid>/generate', methods=['POST'])
def proposal_generate(pid):
    with get_db() as conn:
        proposal = conn.execute('SELECT * FROM proposals WHERE id=?', (pid,)).fetchone()
        if not proposal:
            return jsonify({'error': 'Not found'}), 404
        p = dict(proposal)

    selected_ids = json.loads(p.get('selected_centres') or '[]')
    raw_manual = json.loads(p.get('manual_centres') or '[]')

    db_centres = []
    if selected_ids:
        with get_db() as conn:
            for cid in selected_ids:
                c = conn.execute('SELECT * FROM centres WHERE id=?', (cid,)).fetchone()
                if c:
                    imgs = conn.execute(
                        'SELECT filename FROM centre_images WHERE centre_id=? ORDER BY is_primary DESC, sort_order LIMIT 4',
                        (cid,)).fetchall()
                    centre_data = dict(c)
                    centre_data['images'] = [
                        r['filename'] if r['filename'].startswith('http')
                        else os.path.join(CENTRE_IMAGES, str(cid), r['filename'])
                        for r in imgs
                    ]
                    db_centres.append(centre_data)

    # Convert base64 images in manual centres to temp files
    tmp_files = []
    manual_centres = []
    tmp_dir = os.path.join(UPLOADS, 'tmp_render')
    os.makedirs(tmp_dir, exist_ok=True)
    for mc in raw_manual:
        b64_list = mc.get('images_b64') or []
        img_paths = []
        for b64 in b64_list[:4]:
            if b64 and b64.startswith('data:image'):
                try:
                    _, data = b64.split(',', 1)
                    tmp_path = os.path.join(tmp_dir, f'tmp_{os.urandom(6).hex()}.jpg')
                    Image.open(io.BytesIO(base64.b64decode(data))).convert('RGB').save(tmp_path)
                    img_paths.append(tmp_path)
                    tmp_files.append(tmp_path)
                except:
                    pass
        mc_copy = {k: v for k, v in mc.items() if k != 'images_b64'}
        mc_copy['images'] = img_paths
        manual_centres.append(mc_copy)

    template = p.get('template','london')
    try:
        if template == 'india':
            slides = build_india_slides(p, db_centres, manual_centres)
        elif template == 'bold':
            slides = build_bold_slides(p, db_centres, manual_centres)
        else:
            slides = build_london_slides(p, db_centres, manual_centres)
        base_name = f'proposal_{pid}_{os.urandom(4).hex()}'
        pptx_path = render_pptx(slides, os.path.join(PROPOSAL_FILES, base_name + '.pptx'))
        pdf_path  = render_pdf(slides,  os.path.join(PROPOSAL_FILES, base_name + '.pdf'))
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500
    finally:
        for f in tmp_files:
            try: os.remove(f)
            except: pass

    pptx_name = os.path.basename(pptx_path)
    pdf_name  = os.path.basename(pdf_path)
    with get_db() as conn:
        conn.execute('UPDATE proposals SET pptx_filename=?, pdf_filename=?, status=? WHERE id=?',
                     (pptx_name, pdf_name, 'generated', pid))
    return jsonify({'ok': True, 'pptx': pptx_name, 'pdf': pdf_name})

@app.route('/api/detect-template', methods=['POST'])
def api_detect_template():
    """Analyse an uploaded PPTX template and return feature metadata."""
    f = request.files.get('file')
    if not f:
        return jsonify({'error': 'No file'}), 400
    tmp = os.path.join(TEMPLATE_FILES, 'detect_tmp.pptx')
    f.save(tmp)
    features = detect_template_features(tmp)
    try:
        os.remove(tmp)
    except Exception:
        pass
    return jsonify(features)


@app.route('/proposals/<int:pid>/download')
@app.route('/proposals/<int:pid>/download/<fmt>')
def proposal_download(pid, fmt='pptx'):
    with get_db() as conn:
        row = conn.execute('SELECT pptx_filename, pdf_filename, title FROM proposals WHERE id=?', (pid,)).fetchone()
    if not row:
        return 'Not found', 404
    fname = row['pdf_filename'] if fmt == 'pdf' else row['pptx_filename']
    if not fname:
        return 'Not generated yet', 404
    fpath = os.path.join(PROPOSAL_FILES, fname)
    if not os.path.exists(fpath):
        return 'File not found', 404
    safe_title = (row['title'] or 'proposal').replace(' ','_')
    ext = 'pdf' if fmt == 'pdf' else 'pptx'
    return send_file(fpath, as_attachment=True, download_name=f'{safe_title}.{ext}')

# ── Slide-spec generation + dual rendering ───────────────────────────────────
# Delegate to generation.py which contains the full implementation.

from generation import (
    build_london_slides,
    build_india_slides,
    build_bold_slides,
    detect_template_features,
    render_pptx,
    render_pdf,
    get_logo_png,
)


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5001))
    app.run(debug=False, host='0.0.0.0', port=port)
