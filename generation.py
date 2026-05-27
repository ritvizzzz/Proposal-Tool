"""
generation.py – slide-spec building and rendering for the myHQ proposal tool.

Exported:
    build_london_slides(proposal, db_centres, manual_centres) -> list[list[dict]]
    build_india_slides(proposal,  db_centres, manual_centres) -> list[list[dict]]
    render_pptx(slides, out_path) -> out_path
    render_pdf(slides,  out_path) -> out_path
    get_logo_png()                -> path str

All coordinates are in EMU (914 400 EMU = 1 inch).
Canvas: W=12 192 000  H=6 858 000.
"""

import os
import json
import subprocess
import datetime

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
IMG_DIR    = os.path.join(BASE_DIR, 'static', 'images')
COVER_IMG  = os.path.join(IMG_DIR, 'cover-city.jpg')      # city-of-london high-res cover
CROP_IMG   = os.path.join(IMG_DIR, 'workspace-helps.jpg')  # portrait photo for How myHQ Helps
MAP_IMG    = os.path.join(IMG_DIR, 'london-map.png')
TUBE_DIR   = os.path.join(IMG_DIR, 'tube')

# Station name → primary tube line badge key
STATION_LINE_MAP = {
    # A
    'aldgate':                  'circle',
    'aldgate east':             'district',
    'angel':                    'northern',
    'archway':                  'northern',
    # B
    'baker street':             'jubilee',
    'bank':                     'central',
    'barbican':                 'circle',
    'bayswater':                'circle',
    'bermondsey':               'jubilee',
    'bethnal green':            'central',
    'blackfriars':              'circle',
    'bond street':              'jubilee',
    'borough':                  'northern',
    'brixton':                  'victoria',
    # C
    'caledonian road':          'piccadilly',
    'camden town':              'northern',
    'canada water':             'jubilee',
    'canary wharf':             'jubilee',
    'cannon street':            'circle',
    'chancery lane':            'central',
    'chiswick':                 'district',
    'city thameslink':          'overground',
    'clapham common':           'northern',
    'clapham junction':         'overground',
    'clapham north':            'northern',
    'clapham south':            'northern',
    'cockfosters':              'piccadilly',
    'covent garden':            'piccadilly',
    'crossharbour':             'dlr',
    # D
    'dalston junction':         'overground',
    'elephant & castle':        'northern',
    'elephant and castle':      'northern',
    'euston':                   'northern',
    'euston square':            'circle',
    # F
    'farringdon':               'elizabeth',
    'finsbury park':            'piccadilly',
    # G
    'goodge street':            'northern',
    'green park':               'jubilee',
    'greenwich':                'dlr',
    # H
    'hackney central':          'overground',
    'hammersmith':              'district',
    'heron quays':              'dlr',
    'highbury & islington':     'victoria',
    'highbury and islington':   'victoria',
    'holborn':                  'central',
    # K
    "king's cross":             'victoria',
    "king's cross st pancras":  'victoria',
    'kings cross':              'victoria',
    'knightsbridge':            'piccadilly',
    # L
    'lambeth north':            'bakerloo',
    'leicester square':         'northern',
    'liverpool street':         'elizabeth',
    'london bridge':            'jubilee',
    # M
    'mansion house':            'circle',
    'marble arch':              'central',
    'mile end':                 'central',
    'monument':                 'circle',
    'moorgate':                 'circle',
    'mornington crescent':      'northern',
    # N
    'new cross':                'overground',
    # O
    'old street':               'northern',
    'oxford circus':            'victoria',
    # P
    'paddington':               'elizabeth',
    'peckham rye':              'overground',
    'pimlico':                  'victoria',
    'piccadilly circus':        'piccadilly',
    # R
    'ravenscourt park':         'district',
    'richmond':                 'district',
    'royal oak':                'hammersmith',
    'russell square':           'piccadilly',
    # S
    'shadwell':                 'dlr',
    'shoreditch high street':   'overground',
    'sloane square':            'circle',
    'south kensington':         'piccadilly',
    'south quay':               'dlr',
    'southwark':                'jubilee',
    'st james\'s park':         'circle',
    'st james park':            'circle',
    'st paul\'s':               'central',
    'st pauls':                 'central',
    'stamford brook':           'district',
    'stepney green':            'district',
    'stockwell':                'victoria',
    'stratford':                'elizabeth',
    # T
    'temple':                   'circle',
    'tottenham court road':     'elizabeth',
    'tower hill':               'circle',
    'tower gateway':            'dlr',
    # V
    'vauxhall':                 'victoria',
    'victoria':                 'victoria',
    # W
    'warren street':            'victoria',
    'waterloo':                 'jubilee',
    'west ham':                 'district',
    'westminster':              'jubilee',
    'whitechapel':              'elizabeth',
    'wapping':                  'overground',
}


def station_badge_path(station_name):
    """Return badge PNG for a station name, or None."""
    s = station_name.lower().strip()
    # Exact match
    key = STATION_LINE_MAP.get(s)
    if not key:
        # Partial match (station name contained in map key)
        for k, v in STATION_LINE_MAP.items():
            if k in s or s in k:
                key = v
                break
    if not key:
        # Fall back to line keyword matching (e.g. "Northern line" in name)
        for keyword, k in TUBE_LINE_KEYS.items():
            if keyword in s:
                key = k
                break
    if key:
        p = os.path.join(TUBE_DIR, f'{key}.png')
        if os.path.isfile(p):
            return p
    return None


# Tube line keyword → badge image key
TUBE_LINE_KEYS = {
    'jubilee':          'jubilee',
    'central':          'central',
    'northern':         'northern',
    'piccadilly':       'piccadilly',
    'victoria':         'victoria',
    'district':         'district',
    'circle':           'circle',
    'bakerloo':         'bakerloo',
    'elizabeth':        'elizabeth',
    'crossrail':        'elizabeth',
    'metropolitan':     'metropolitan',
    'hammersmith & city':'hammersmith',
    'hammersmith':      'hammersmith',
    'h&c':              'hammersmith',
    'waterloo & city':  'waterloo',
    'waterloo':         'waterloo',
    'dlr':              'dlr',
    'docklands':        'dlr',
    'overground':       'overground',
    'national rail':    'overground',
    'thameslink':       'overground',
}

# ── Canvas & palette ────────────────────────────────────────────────────────────
W, H = 12192000, 6858000
M    = 550000          # left / right margin

BLUE  = '#1E22AA'
MINT  = '#ADDFB3'
BLUSH = '#F4BEAA'
AQUA  = '#96E6DC'
WHITE = '#FFFFFF'
DARK  = '#1D1D1B'
GREY  = '#6B7280'
LGREY = '#F3F4F6'
DGREY = '#9CA3AF'


# ── Primitive draw-command helpers ──────────────────────────────────────────────

def R(x, y, w, h, fill):
    return {'cmd': 'rect', 'x': x, 'y': y, 'w': w, 'h': h, 'fill': fill}


def T(text, x, y, w, h, size, color,
      font='Calibri', bold=False, italic=False, align='L', wrap=True):
    return {
        'cmd':    'text',
        'text':   str(text or ''),
        'x': x, 'y': y, 'w': w, 'h': h,
        'font':   font,
        'size':   size,
        'color':  color,
        'bold':   bold,
        'italic': italic,
        'align':  align,
        'wrap':   wrap,
    }


def I(path, x, y, w, h):
    return {'cmd': 'img', 'path': path, 'x': x, 'y': y, 'w': w, 'h': h}


def L(x1, y1, x2, y2, color='#E5E7EB', width=1):
    return {'cmd': 'line', 'x1': x1, 'y1': y1, 'x2': x2, 'y2': y2,
            'color': color, 'width': width}


# ── Helper utilities ────────────────────────────────────────────────────────────

AMENITY_LABELS = {
    'wifi':                    'Wi-Fi',
    'parking':                 'Parking',
    'four_wheeler_parking':    'Car Parking',
    'two_wheeler_parking':     'Bike Parking',
    'cafeteria':               'Cafeteria',
    'gym':                     'Gym',
    'meeting_rooms':           'Meeting Rooms',
    'meeting_room':            'Meeting Room',
    'phone_booths':            'Phone Booths',
    'phone_booth':             'Phone Booth',
    'reception':               'Reception',
    'printing':                'Printing',
    'lounge':                  'Lounge',
    'security':                '24/7 Security',
    'air_conditioning':        'Air Con',
    'power_backup':            'Power Backup',
    'breakout_area':           'Breakout Area',
    'terrace':                 'Terrace',
    'bike_storage':            'Bike Storage',
    'lift':                    'Lift',
    'shower':                  'Showers',
    'kitchen':                 'Kitchen',
    'event_space':             'Event Space',
    'rooftop':                 'Rooftop',
    'onsite_cafe':             'Café',
    'onsite_cafeteria':        'Café',
    'cctv':                    'CCTV',
    'fire_safety':             'Fire Safety',
    'housekeeping':            'Housekeeping',
    'it_support':              'IT Support',
    'mail_handling':           'Mail Handling',
    'ramp_access':             'Ramp Access',
    'disabled_access':         'Disabled Access',
    'generator_backup':        'Generator',
    'storage':                 'Storage',
    'lockers':                 'Lockers',
    'nursing_room':            'Nursing Room',
    'pantry':                  'Pantry',
    'concierge':               'Concierge',
}

# Dot colours cycling for amenity grid
_AMENITY_COLORS = [AQUA, MINT, '#7CB9E8', BLUSH, '#C8E6C9', '#FFF176', '#FFB74D']


def format_amenities(a):
    try:
        ams = json.loads(a) if isinstance(a, str) else (a or [])
    except Exception:
        ams = []
    labels = []
    for x in ams[:8]:
        slug = str(x).lower().strip()
        label = AMENITY_LABELS.get(slug)
        if not label:
            # Clean up unknown slugs
            label = slug.replace('_', ' ').replace('-', ' ').title()
            if len(label) > 14:
                label = label[:13] + '…'
        labels.append(label)
    return ', '.join(labels) if labels else 'On request'


def amenity_pill_cmds(amenities_raw, x, y, max_w, font):
    """3-column grid: coloured square dot + label text. Clean, no overflow."""
    try:
        ams = json.loads(amenities_raw) if isinstance(amenities_raw, str) else (amenities_raw or [])
    except Exception:
        ams = []
    if not ams:
        return [T('On request', x, y, max_w, 240000, 8, GREY, font=font)]

    cmds = []
    cols = 3
    col_w = max_w // cols
    row_h = 290000
    dot_size = 130000
    dot_gap = 90000

    for i, slug in enumerate(ams[:9]):
        slug = str(slug).lower().strip()
        label = AMENITY_LABELS.get(slug)
        if not label:
            label = slug.replace('_', ' ').replace('-', ' ').title()
            if len(label) > 14:
                label = label[:13] + '…'

        col = i % cols
        row = i // cols
        ix = x + col * col_w
        iy = y + row * row_h

        color = _AMENITY_COLORS[i % len(_AMENITY_COLORS)]
        # Small filled square bullet
        cmds.append(R(ix, iy + 50000, dot_size, dot_size, color))
        # Label text beside it
        cmds.append(T(label,
                      ix + dot_size + dot_gap, iy,
                      col_w - dot_size - dot_gap - 40000, 240000,
                      8, DARK, font=font, wrap=False))

    return cmds


def price_str(c):
    p = c.get('price_from') or c.get('price_str')
    if not p:
        return 'On request'
    if isinstance(p, str):
        return p
    unit = (c.get('price_unit') or 'MONTHLY').upper()
    suffix = {'MONTHLY': '/mo', 'DAILY': '/day', 'HOURLY': '/hr'}.get(unit, '/mo')
    city = (c.get('city') or '').lower()
    symbol = '£' if 'london' in city or 'uk' in city else '₹'
    return f'{symbol}{int(p):,}{suffix}/seat'


def extract_transport_parts(transport_str):
    """Return (tube_line, bus_line, other) from a transport string."""
    s = transport_str or ''
    tube = bus = other = ''
    for part in s.split(','):
        part = part.strip()
        pl = part.lower()
        if 'tube' in pl or 'underground' in pl or 'metro' in pl or 'line' in pl:
            tube = part
        elif 'bus' in pl:
            bus = part
        elif not other:
            other = part
    return tube, bus, other


def tube_badge_path(transport_str):
    """Return the PNG path for the tube line badge best matching transport_str, or None."""
    s = (transport_str or '').lower()
    for keyword, key in TUBE_LINE_KEYS.items():
        if keyword in s:
            p = os.path.join(TUBE_DIR, f'{key}.png')
            if os.path.isfile(p):
                return p
    return None


def get_logo_png(white=False):
    """Return a path to a usable PNG of the myHQ logo (transparent background).
    white=True returns the all-white version for use on dark/coloured backgrounds.
    """
    suffix = '-white' if white else ''
    png_path = os.path.join(BASE_DIR, 'static', 'images', f'myhq-logo{suffix}.png')
    if os.path.exists(png_path) and os.path.getsize(png_path) > 5000:
        return png_path

    # If white version requested but missing, generate from the colour version
    if white:
        base = get_logo_png(white=False)
        try:
            import numpy as np
            from PIL import Image as _PI
            img = _PI.open(base).convert('RGBA')
            data = np.array(img)
            visible = data[:, :, 3] > 10
            data[visible, 0] = 255
            data[visible, 1] = 255
            data[visible, 2] = 255
            _PI.fromarray(data, 'RGBA').save(png_path, 'PNG')
            return png_path
        except Exception:
            return base   # fall back to colour logo

    os.makedirs(os.path.dirname(png_path), exist_ok=True)

    # Candidate SVG sources
    svg_candidates = [
        os.path.join(BASE_DIR, 'static', 'images', 'myhq-logo.svg'),
        os.path.join(BASE_DIR, '..', 'onboarding-tool', 'static', 'myhq-logo.svg'),
        os.path.join(BASE_DIR, 'static', 'myhq-logo.svg'),
    ]
    logo_svg = next((p for p in svg_candidates if os.path.exists(p)), None)

    if logo_svg:
        # Try qlmanage (macOS built-in) and strip white background
        try:
            import numpy as np
            from PIL import Image as _PI
            tmp_dir = os.path.join(BASE_DIR, 'static', 'images')
            subprocess.run(['qlmanage', '-t', '-s', '1280', '-o', tmp_dir, logo_svg],
                           timeout=15, capture_output=True)
            ql_candidate = os.path.join(tmp_dir, os.path.basename(logo_svg) + '.png')
            if os.path.exists(ql_candidate):
                img = _PI.open(ql_candidate).convert('RGBA')
                data = np.array(img)
                white_mask = (data[:,:,0] > 240) & (data[:,:,1] > 240) & (data[:,:,2] > 240)
                data[white_mask, 3] = 0
                result = _PI.fromarray(data, 'RGBA')
                bbox = result.getbbox()
                if bbox:
                    result = result.crop(bbox)
                result.save(png_path, 'PNG')
                os.remove(ql_candidate)
                if os.path.getsize(png_path) > 5000:
                    return png_path
        except Exception:
            pass

    # Fallback: draw "myHQ" text with Pillow
    try:
        from PIL import Image as PILImage, ImageDraw, ImageFont
        img = PILImage.new('RGBA', (320, 120), (255, 255, 255, 0))
        draw = ImageDraw.Draw(img)
        try:
            font = ImageFont.truetype('/System/Library/Fonts/Helvetica.ttc', 64)
        except Exception:
            font = ImageFont.load_default()
        draw.text((10, 20), 'myHQ', fill=(30, 34, 170, 255), font=font)
        img.save(png_path, 'PNG')
        return png_path
    except Exception:
        pass

    try:
        from PIL import Image as PILImage
        PILImage.new('RGBA', (1, 1), (255, 255, 255, 0)).save(png_path, 'PNG')
    except Exception:
        pass
    return png_path


# ── Logo compositing helper ─────────────────────────────────────────────────────

def _logo_on(logo_path, bg_hex):
    """Return path to logo composited onto a solid background colour (for transparent-safe rendering)."""
    import hashlib
    key = bg_hex.lstrip('#').upper()
    cache = os.path.join(BASE_DIR, 'uploads', f'.logo_cache_{key}.png')
    os.makedirs(os.path.dirname(cache), exist_ok=True)
    if os.path.exists(cache) and os.path.getsize(cache) > 500:
        return cache
    try:
        from PIL import Image as _PI
        h = bg_hex.lstrip('#')
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        logo = _PI.open(logo_path).convert('RGBA')
        bg = _PI.new('RGB', logo.size, (r, g, b))
        bg.paste(logo, mask=logo.split()[3])
        bg.save(cache, 'PNG')
        return cache
    except Exception:
        return logo_path


# ── Shared slide builders (used by both templates) ─────────────────────────────

def _small_logo(logo_path, bg='#FFFFFF'):
    """Top-right small logo — composited against bg so it renders cleanly."""
    composited = _logo_on(logo_path, bg)
    return I(composited, W - 1500000, 160000, 1100000, 412000)


def _comparison_table_slide(all_centres, header_fill, header_text_color, font, logo):
    """Slide 2: Space Comparison table with photo column. Vertical layout when > 5 centres."""
    sl = [R(0, 0, W, H, WHITE)]
    sl.append(_small_logo(logo))

    sl.append(T('Space Comparison', M, 250000, W - 2 * M, 450000,
                22, BLUE, font=font, bold=True))
    sl.append(T('All shortlisted workspaces at a glance', M, 720000, W - 2 * M, 300000,
                11, GREY, font=font))

    table_x = M
    table_y = 1100000
    table_w = W - 2 * M
    table_h = H - 1500000

    num_centres = max(len(all_centres), 1)
    hdr_h   = 520000
    raw_row_h = (table_h - hdr_h) // num_centres
    row_h   = max(400000, min(1100000, raw_row_h))

    # Use a photo column – 900000 wide if rows are tall enough
    photo_col_w = 900000 if row_h >= 600000 else 0

    # Column definitions: (label, fixed_width_or_None, field_fn)
    def _centre_tube(c):
        t, b, o = extract_transport_parts(c.get('transport', ''))
        return t or o or '—'

    fixed_cols = [
        ('#',           380000,  lambda i, c: str(i + 1)),
        ('NAME',       1900000,  lambda i, c: c.get('name', '—')),
        ('ADDRESS',    2200000,  lambda i, c: c.get('address', '—')),
        ('PRICE/SEAT', 1300000,  lambda i, c: price_str(c)),
        ('TUBE',       1500000,  lambda i, c: _centre_tube(c)),
        ('HOURS',      1300000,  lambda i, c: c.get('open_hours', '9–6')),
        ('TYPE',       1212000,  lambda i, c: (c.get('space_type') or '—').title()),
    ]
    if photo_col_w:
        fixed_cols.append(('PHOTO', photo_col_w, None))

    # Draw header
    cx = table_x
    for label, cw, _ in fixed_cols:
        sl.append(R(cx, table_y, cw, hdr_h, header_fill))
        sl.append(T(label, cx + 35000, table_y + 110000, cw - 70000, hdr_h - 180000,
                    8, header_text_color, font=font, bold=True))
        cx += cw

    # Draw data rows
    for ri, centre in enumerate(all_centres):
        ry = table_y + hdr_h + ri * row_h
        row_fill = WHITE if ri % 2 == 0 else LGREY
        cx = table_x
        for label, cw, val_fn in fixed_cols:
            sl.append(R(cx, ry, cw, row_h, row_fill))
            if val_fn is None:
                # Photo cell
                imgs = centre.get('images', [])
                photo = imgs[0] if imgs else None
                sl.append(I(photo, cx + 20000, ry + 20000,
                            cw - 40000, row_h - 40000))
            else:
                val = val_fn(ri, centre)
                # Tube column: try to show badge image
                if label == 'TUBE':
                    badge = tube_badge_path(centre.get('transport', ''))
                    if badge:
                        badge_h = min(260000, row_h - 80000)
                        badge_w = int(badge_h * 3.33)  # 320:96 roundel aspect ratio
                        sl.append(I(badge, cx + 35000, ry + (row_h - badge_h) // 2,
                                    badge_w, badge_h))
                    else:
                        sl.append(T(val, cx + 35000, ry + 60000,
                                    cw - 70000, row_h - 100000,
                                    8, DARK, font=font, wrap=True))
                else:
                    sl.append(T(val, cx + 35000, ry + 60000,
                                cw - 70000, row_h - 100000,
                                8, DARK, font=font, wrap=True))
            cx += cw
        sl.append(L(table_x, ry + row_h, table_x + table_w, ry + row_h,
                    color='#E5E7EB', width=1))

    return sl


def _client_requirements_slide(proposal, template, db_centres, manual_centres,
                                font, logo,
                                label_fill, label_text, date_label_color):
    """Slide 4: Client Requirements."""
    sl = [R(0, 0, W, H, WHITE)]
    sl.append(_small_logo(logo))
    sl.append(T('Client Requirements', M, 250000, 6000000, 400000,
                22, BLUE, font=font, bold=True))

    today = datetime.date.today().strftime('%d %B %Y')

    reqs = [
        ('CLIENT',         proposal.get('client_name')     or '—'),
        ('COMPANY',        proposal.get('client_company')  or '—'),
        ('LOCATION',       proposal.get('client_location') or '—'),
        ('TEAM SIZE',      proposal.get('team_size')       or '—'),
        ('SPACE TYPE',     proposal.get('space_type')      or '—'),
        ('AREA REQUIRED',  proposal.get('area_required')   or '—'),
        ('BUDGET',         proposal.get('budget')          or '—'),
        ('DURATION',       proposal.get('duration')        or '—'),
    ]

    lbl_col_w  = 2100000
    val_col_w  = 4700000
    row_h      = 660000
    tbl_x      = M
    tbl_y      = 1000000

    for i, (lbl, val) in enumerate(reqs):
        ry   = tbl_y + i * row_h
        vfill = WHITE if i % 2 == 0 else LGREY
        sl.append(R(tbl_x,               ry, lbl_col_w, row_h, label_fill))
        sl.append(R(tbl_x + lbl_col_w,   ry, val_col_w, row_h, vfill))
        sl.append(T(lbl, tbl_x + 60000, ry + 180000, lbl_col_w - 120000, row_h - 200000,
                    9, label_text, font=font, bold=True))
        sl.append(T(val, tbl_x + lbl_col_w + 80000, ry + 150000,
                    val_col_w - 160000, row_h - 200000,
                    11, DARK, font=font, wrap=True))
        # bottom divider
        sl.append(L(tbl_x, ry + row_h,
                    tbl_x + lbl_col_w + val_col_w, ry + row_h))

    # Right summary card – shows proposal meta only, no repeated fields
    card_x = tbl_x + lbl_col_w + val_col_w + 300000
    card_w = W - card_x - M
    card_h = len(reqs) * row_h
    card_y = tbl_y

    sl.append(R(card_x, card_y, card_w, card_h, BLUE))

    inner_x = card_x + 120000
    inner_w  = card_w - 240000

    # Date label
    sl.append(T(today, inner_x, card_y + 120000, inner_w, 250000,
                9, date_label_color, font=font))
    sl.append(L(card_x + 60000, card_y + 400000,
                card_x + card_w - 60000, card_y + 400000,
                color='#3A3E99', width=1))

    # Company name (prominent heading)
    company = proposal.get('client_company') or proposal.get('client_name') or '—'
    sl.append(T(company, inner_x, card_y + 460000, inner_w, 900000,
                28, WHITE, font=font, bold=True, wrap=True))

    sl.append(L(card_x + 60000, card_y + 1480000,
                card_x + card_w - 60000, card_y + 1480000,
                color='#3A3E99', width=1))

    # "Prepared by myHQ" section
    sl.append(T('PREPARED BY', inner_x, card_y + 1560000, inner_w, 240000,
                8, date_label_color, font=font, bold=True))
    sl.append(T('myHQ by Anarock', inner_x, card_y + 1840000, inner_w, 360000,
                16, WHITE, font=font, bold=True))

    sl.append(L(card_x + 60000, card_y + 2280000,
                card_x + card_w - 60000, card_y + 2280000,
                color='#3A3E99', width=1))

    # Contact
    sl.append(T('YOUR CONSULTANT', inner_x, card_y + 2380000, inner_w, 240000,
                8, date_label_color, font=font, bold=True))
    sl.append(T('workspace@myhq.in', inner_x, card_y + 2660000, inner_w, 280000,
                11, WHITE, font=font))
    sl.append(T('+91 98765 43210', inner_x, card_y + 2980000, inner_w, 280000,
                11, WHITE, font=font))

    # Decorative myHQ icon area
    sl.append(R(card_x, card_y + card_h - 600000, card_w, 600000, '#16199A'))
    sl.append(T('myHQ', inner_x, card_y + card_h - 480000, inner_w, 380000,
                22, WHITE, font=font, bold=True))

    return sl


def _extract_station_name(transport_str):
    """Pull the nearest station name from a transport string."""
    s = (transport_str or '').strip()
    if not s:
        return ''
    # Try common patterns: "X mins walk from Y station", "nearest: Y", "Y (5 min)"
    import re
    m = re.search(r'(?:from|to|nearest[:\s]+)\s*([A-Z][^,\n;(]{2,30}?)(?:\s*station|\s*tube|\s*underground|\s*\(|,|$)',
                  s, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    # Fallback: first proper-noun segment
    parts = re.split(r'[,;:\n]', s)
    for p in parts:
        p = p.strip()
        # Remove leading "Bus:", "Tube:" etc.
        p = re.sub(r'^(bus|tube|underground|overground|rail|dlr|metro)\s*:?\s*', '', p, flags=re.IGNORECASE)
        p = p.strip()
        if len(p) > 3:
            return p[:40]
    return s[:40]


def _extract_walk_time(transport_str):
    """Pull walking time from transport string, e.g. '5 min walk'."""
    import re
    m = re.search(r'(\d+)\s*(?:min|minute|mins)', transport_str or '', re.IGNORECASE)
    if m:
        return f"{m.group(1)} min walk"
    return ''


def _centre_slide(idx, centre, template, font, logo):
    """One centre detail slide (Slide 6+)."""
    sl = [R(0, 0, W, H, WHITE)]

    # ── Accent bar at top ────────────────────────────────────────────────────────
    sl.append(R(0, 0, W, 120000, AQUA if template != 'india' else BLUE))

    # ── Header ──────────────────────────────────────────────────────────────────
    sl.append(T(f'{idx:02d}', M, 140000, 300000, 480000,
                22, BLUE, font=font, bold=True))
    sl.append(T(centre.get('name', 'Centre'),
                M + 380000, 160000, 7600000, 420000,
                18, DARK, font=font, bold=True))
    sl.append(_small_logo(logo))
    sl.append(L(M, 680000, W - M, 680000, color=LGREY, width=1))

    # ── Left column ──────────────────────────────────────────────────────────────
    lcol_x = M
    lcol_w  = 4000000
    cy = 760000

    # ADDRESS
    sl.append(T('ADDRESS', lcol_x, cy, lcol_w, 190000, 7, DGREY, font=font, bold=True))
    cy += 190000
    sl.append(T(centre.get('address') or '—', lcol_x, cy, lcol_w, 280000, 9, DARK, font=font, wrap=True))
    cy += 310000
    sl.append(L(lcol_x, cy, lcol_x + lcol_w, cy, color='#EEEEEE', width=1))
    cy += 50000

    # PRICE + HOURS on same row
    half = (lcol_w - 80000) // 2
    sl.append(T('PRICE / SEAT', lcol_x, cy, half, 190000, 7, DGREY, font=font, bold=True))
    sl.append(T('OPEN HOURS', lcol_x + half + 80000, cy, half, 190000, 7, DGREY, font=font, bold=True))
    cy += 190000
    sl.append(T(price_str(centre), lcol_x, cy, half, 280000, 10, BLUE, font=font, bold=True))
    sl.append(T(centre.get('open_hours') or '9:00 AM – 6:00 PM',
                lcol_x + half + 80000, cy, half, 280000, 9, DARK, font=font))
    cy += 340000
    sl.append(L(lcol_x, cy, lcol_x + lcol_w, cy, color='#EEEEEE', width=1))
    cy += 60000

    # AMENITIES as pills
    sl.append(T('AMENITIES', lcol_x, cy, lcol_w, 190000, 7, DGREY, font=font, bold=True))
    cy += 200000
    amenity_cmds = amenity_pill_cmds(centre.get('amenities', '[]'), lcol_x, cy, lcol_w, font)
    sl.extend(amenity_cmds)
    # Grid is 3-column; calculate actual rows used
    try:
        ams = json.loads(centre.get('amenities', '[]') if isinstance(centre.get('amenities', '[]'), str) else '[]')
        n_ams = min(len(ams), 9)
    except Exception:
        n_ams = 0
    grid_rows = max(1, (n_ams + 2) // 3)
    cy += grid_rows * 290000 + 100000

    sl.append(L(lcol_x, cy, lcol_x + lcol_w, cy, color='#EEEEEE', width=1))
    cy += 60000

    # NEAREST TUBE / TRANSPORT
    transport_str = centre.get('transport', '')
    tube, bus, other = extract_transport_parts(transport_str)
    badge = tube_badge_path(transport_str)
    station_name = _extract_station_name(transport_str)
    walk_time = _extract_walk_time(transport_str)

    sl.append(T('NEAREST TUBE / TRANSPORT', lcol_x, cy, lcol_w, 190000, 7, DGREY, font=font, bold=True))
    cy += 200000

    if badge:
        sl.append(I(badge, lcol_x, cy, 700000, 230000))
        txt_x = lcol_x + 760000
        txt_w = lcol_w - 760000
        if station_name:
            sl.append(T(station_name, txt_x, cy, txt_w, 200000, 9, DARK, font=font, bold=True))
        if walk_time:
            sl.append(T(walk_time, txt_x, cy + 190000, txt_w, 180000, 8, GREY, font=font))
        cy += 320000
    elif tube or station_name or other:
        display = station_name or tube or other
        sl.append(R(lcol_x, cy + 70000, 8000, 140000, BLUE))
        sl.append(T(display, lcol_x + 60000, cy, lcol_w - 60000, 200000, 9, DARK, font=font, bold=True))
        if walk_time:
            sl.append(T(walk_time, lcol_x + 60000, cy + 210000, lcol_w - 60000, 180000, 8, GREY, font=font))
        cy += 320000

    if bus:
        sl.append(R(lcol_x, cy + 70000, 8000, 140000, AQUA if template != 'india' else BLUE))
        sl.append(T(bus, lcol_x + 60000, cy, lcol_w - 60000, 200000, 9, DARK, font=font))
        cy += 270000

    cy += 60000
    sl.append(L(lcol_x, cy, lcol_x + lcol_w, cy, color='#EEEEEE', width=1))
    cy += 60000

    # WHY WE RECOMMEND
    why_text = (centre.get('why_recommend') or centre.get('about')
                or 'A premium flexible workspace in a prime location, ideal for growing teams.')
    sl.append(T('WHY WE RECOMMEND', lcol_x, cy, lcol_w, 190000, 7, BLUE, font=font, bold=True))
    cy += 200000
    # Fill remaining left-column height
    remaining = H - cy - 100000
    sl.append(T(why_text, lcol_x, cy, lcol_w, max(remaining, 400000), 9, DARK, font=font, wrap=True))

    # ── Right column: 2×2 equal photo grid ───────────────────────────────────────
    GAP  = 120000
    rx   = M + lcol_w + 200000
    ry   = 720000
    rw   = W - rx - 80000
    rh   = H - ry - 100000

    cw   = (rw - GAP) // 2
    ch   = (rh - GAP) // 2

    images = centre.get('images', [])

    def _cell(ci, gx, gy):
        p = images[ci] if ci < len(images) else None
        return I(p, gx, gy, cw, ch)

    sl.append(_cell(0, rx,            ry))
    sl.append(_cell(1, rx + cw + GAP, ry))
    sl.append(_cell(2, rx,            ry + ch + GAP))
    sl.append(_cell(3, rx + cw + GAP, ry + ch + GAP))

    return sl


def _existing_clients_slide(font, logo):
    # Full-slide image taken directly from the reference PDF
    slide_img = os.path.join(IMG_DIR, 'slide_existing_clients.png')
    return [I(slide_img, 0, 0, W, H)]


def _testimonials_slide(font, logo, bg_color=WHITE, heading_color=DARK):
    # Full-slide image taken directly from the reference PDF
    slide_img = os.path.join(IMG_DIR, 'slide_word_of_mouth.png')
    return [I(slide_img, 0, 0, W, H)]


def generate_proposal_map(centres, out_path):
    """Screenshot a styled Leaflet map showing only the given centres with numbered pins."""
    import tempfile, json as _json, threading, http.server
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return None

    points = []
    for c in centres:
        coords = c.get('coordinates', '')
        if coords and ';' in coords:
            try:
                lng, lat = coords.split(';')
                points.append({'lat': float(lat), 'lng': float(lng), 'name': c.get('name', '')})
            except Exception:
                pass

    if not points:
        return None

    # Build numbered markers with custom blue pins + name labels
    markers_js_parts = []
    for i, p in enumerate(points, 1):
        label = p['name'].replace("'", "\\'")[:30]
        icon_js = (
            f"L.divIcon({{"
            f"className:'',"
            f"html:'<div style=\"background:#1E22AA;color:#fff;min-width:28px;height:28px;"
            f"border-radius:50%;display:flex;align-items:center;justify-content:center;"
            f"font-family:sans-serif;font-size:13px;font-weight:700;border:2px solid #fff;"
            f"box-shadow:0 2px 8px rgba(0,0,0,0.45)\">{i}</div>',"
            f"iconSize:[28,28],iconAnchor:[14,14]"
            f"}})"
        )
        label_icon = (
            f"L.divIcon({{"
            f"className:'',"
            f"html:'<div style=\"background:rgba(255,255,255,0.92);color:#1D1D1B;padding:2px 6px;"
            f"border-radius:3px;font-family:sans-serif;font-size:10px;font-weight:600;"
            f"white-space:nowrap;box-shadow:0 1px 4px rgba(0,0,0,0.2)\">{label}</div>',"
            f"iconSize:[180,20],iconAnchor:[-16,10]"
            f"}})"
        )
        markers_js_parts.append(
            f"L.marker([{p['lat']},{p['lng']}],{{icon:{icon_js}}}).addTo(map);"
            f"L.marker([{p['lat']},{p['lng']}],{{icon:{label_icon}}}).addTo(map);"
        )
    markers_js = '\n'.join(markers_js_parts)
    bounds_arr = ','.join(f'[{p["lat"]},{p["lng"]}]' for p in points)

    html = f"""<!DOCTYPE html><html><head>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<style>
  html,body{{margin:0;padding:0;background:#e8edf2}}
  #map{{width:100vw;height:100vh}}
  .leaflet-control-attribution,.leaflet-control-zoom{{display:none}}
</style>
</head><body>
<div id="map"></div>
<script>
var map = L.map('map',{{zoomControl:false,attributionControl:false}});
L.tileLayer('https://{{s}}.basemaps.cartocdn.com/rastertiles/voyager/{{z}}/{{x}}/{{y}}{{r}}.png',{{maxZoom:19}}).addTo(map);
{markers_js}
var bounds = L.latLngBounds([{bounds_arr}]);
map.fitBounds(bounds,{{padding:[60,60]}});
</script></body></html>"""

    # Serve HTML via local HTTP so tile requests are HTTPS-allowed
    tmp_dir = tempfile.mkdtemp()
    html_file = os.path.join(tmp_dir, 'map.html')
    with open(html_file, 'w') as f:
        f.write(html)

    import functools
    port = 19871
    handler = functools.partial(http.server.SimpleHTTPRequestHandler, directory=tmp_dir)
    httpd = http.server.HTTPServer(('127.0.0.1', port), handler)
    httpd.allow_reuse_address = True
    t = threading.Thread(target=httpd.serve_forever)
    t.daemon = True
    t.start()

    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch(args=['--no-sandbox'])
            page = browser.new_page(viewport={'width': 1200, 'height': 800})
            page.goto(f'http://127.0.0.1:{port}/map.html', wait_until='networkidle', timeout=20000)
            page.wait_for_timeout(3000)
            page.screenshot(path=out_path, full_page=False)
            browser.close()
        return out_path
    except Exception:
        return None
    finally:
        httpd.shutdown()
        import shutil
        shutil.rmtree(tmp_dir, ignore_errors=True)


_LOCATION_DATA = {
    # key: lowercase keywords to match → (about_text, highlights, [popular_stations])
    'city of london': (
        'The City of London is the UK\'s premier financial district — a square mile packed with global banks, law firms, and tech companies. With Crossrail, Tube, and rail all converging here, getting in and out is effortless.',
        ['One of the world\'s leading financial centres', 'Crossrail (Elizabeth line) cuts journey times dramatically', 'Multiple Tube lines: Central, Circle, District, Jubilee', 'World-class dining and networking at every corner'],
        ['Liverpool Street', 'Bank', 'Cannon Street', 'Moorgate', 'Aldgate'],
    ),
    'canary wharf': (
        'Canary Wharf is London\'s second financial hub, home to HSBC, Barclays, and hundreds of leading firms. Superbly connected via the Jubilee line, DLR, and Elizabeth line, it offers modern workspaces in a landmark riverside setting.',
        ['Home to major global banks and financial institutions', 'Elizabeth line direct to Heathrow & Paddington', 'Jubilee line to Westminster & London Bridge in under 10 mins', 'Extensive shopping, dining, and leisure on-site'],
        ['Canary Wharf', 'Heron Quays', 'South Quay', 'Crossharbour'],
    ),
    'shoreditch': (
        'Shoreditch is London\'s creative and tech heartland, attracting startups, agencies, and fast-growing scaleups. The area buzzes with innovation, with a tight-knit business community and superb east London connections.',
        ['London\'s leading tech and creative hub', 'Excellent connections via Overground and Elizabeth line', 'Walking distance to Liverpool Street and Old Street stations', 'Vibrant independent food, coffee, and social scene'],
        ['Shoreditch High Street', 'Old Street', 'Liverpool Street', 'Bethnal Green'],
    ),
    'old street': (
        'Old Street — the heart of Silicon Roundabout — is where London\'s tech ecosystem thrives. Surrounded by co-working pioneers, accelerators, and VC-backed startups, this is the ideal base for forward-thinking businesses.',
        ['Centre of London\'s tech startup ecosystem', 'Northern line direct to King\'s Cross and the West End', 'Huge choice of flexible and serviced offices nearby', 'Bike-friendly streets and excellent bus connections'],
        ['Old Street', 'Moorgate', 'Angel', 'Barbican'],
    ),
    'mayfair': (
        'Mayfair is synonymous with prestige — London\'s most exclusive business address. Home to luxury brands, private equity, hedge funds, and embassies, offices here send an unmistakable signal of quality and ambition.',
        ['London\'s most prestigious business postcode', 'Steps from Bond Street and Green Park stations', 'Surrounded by five-star hotels and fine dining', 'Green Park and Hyde Park nearby for client walks'],
        ['Bond Street', 'Green Park', 'Oxford Circus', 'Marble Arch'],
    ),
    'victoria': (
        'Victoria is one of London\'s best-connected transport hubs, giving easy access to Heathrow, Gatwick, and the whole of the Underground network. It is a cost-effective alternative to Mayfair with a rapidly improving offer of quality workspaces.',
        ['Direct trains to Gatwick Airport in under 30 minutes', 'Victoria, Circle, and District lines converge here', 'Strong presence of government, NGOs, and professional services', 'Major regeneration driving new workspace supply'],
        ['Victoria', 'Sloane Square', 'St James\'s Park', 'Pimlico'],
    ),
    'paddington': (
        'Paddington has been transformed by the Elizabeth line into one of London\'s fastest-growing business districts. With direct links to Heathrow in 15 minutes and the City in under 10, it offers unmatched connectivity for client-facing teams.',
        ['Elizabeth line to Heathrow Airport in 15 minutes', 'Direct Heathrow Express rail link on-site', 'Rapidly growing cluster of tech and media firms', 'Canal-side setting with excellent work-life amenities'],
        ['Paddington', 'Edgware Road', 'Bayswater', 'Royal Oak'],
    ),
    'hammersmith': (
        'Hammersmith is a well-established west London business district, popular with media, broadcasting, and professional services firms. Excellent Tube and Overground connections make it ideal for teams working across the capital.',
        ['Major media and broadcasting hub (BBC, Disney, L\'Oreal)', 'Four Tube lines: Piccadilly, District, Circle, Hammersmith & City', 'Riverside setting with excellent restaurants and bars', 'Competitive pricing versus central London'],
        ['Hammersmith', 'Ravenscourt Park', 'Stamford Brook', 'Barons Court'],
    ),
    'waterloo': (
        'Waterloo sits at the crossroads of south and central London. The largest rail terminus in the UK ensures unrivalled connectivity, while the South Bank\'s cultural energy makes it a dynamic and inspiring place to work.',
        ['Busiest rail terminus in the UK', 'Jubilee and Bakerloo lines plus overground services', 'South Bank cultural quarter: Tate Modern, Southbank Centre', 'Fast river boat services to Canary Wharf and the City'],
        ['Waterloo', 'Lambeth North', 'Southwark', 'London Bridge'],
    ),
    'london bridge': (
        'London Bridge bridges (literally) the City and South Bank, offering a vibrant mix of workspace types right next to one of London\'s busiest transport hubs. The Borough Market area adds to its appeal with world-class food and culture.',
        ['Seconds from London Bridge Station (Jubilee & Northern lines)', 'Walking distance to the City\'s financial core', 'Borough Market and Bermondsey Street for client entertainment', 'Strong legal, financial, and tech cluster'],
        ['London Bridge', 'Borough', 'Bermondsey', 'Elephant & Castle'],
    ),
    'farringdon': (
        'Farringdon is rapidly becoming one of London\'s most sought-after business locations, thanks to Crossrail making it a new interchange supernode. The area blends the creative energy of Clerkenwell with excellent connectivity.',
        ['New Elizabeth line interchange — connections in every direction', 'Heart of London\'s legal and media village', 'Clerkenwell\'s thriving design and architecture scene nearby', 'Excellent cycling infrastructure'],
        ['Farringdon', 'Barbican', 'Chancery Lane', 'Clerkenwell'],
    ),
    'king\'s cross': (
        'King\'s Cross has been completely reimagined over the past decade and is now home to Google, Facebook, and a thriving knowledge economy cluster. Six Tube lines, Eurostar, and intercity rail make it the most connected location in London.',
        ['Six Tube lines converge at King\'s Cross St. Pancras', 'Eurostar terminal for Paris/Brussels in under 2.5 hours', 'Google UK headquarters and growing tech campus', 'Coal Drops Yard — exceptional dining and retail'],
        ['King\'s Cross St Pancras', 'Euston', 'Russell Square', 'Caledonian Road'],
    ),
    'euston': (
        'Euston is a major transport gateway for those travelling to and from the Midlands and the North. With HS2 investment transforming the area, Euston is set to become one of London\'s most important business locations.',
        ['Major terminus for trains to Birmingham, Manchester, and Scotland', 'Victoria and Northern lines on the doorstep', 'HS2 development creating a new business district', 'Minutes from King\'s Cross and Bloomsbury'],
        ['Euston', 'Warren Street', 'Euston Square', 'Mornington Crescent'],
    ),
    'soho': (
        'Soho is London\'s most vibrant creative quarter — home to advertising agencies, film production companies, music labels, and cutting-edge digital businesses. Its central location and buzzing atmosphere make it the address of choice for creative industries.',
        ['London\'s creative and entertainment epicentre', 'Multiple Tube lines: Northern, Central, Bakerloo, Jubilee', 'Unique mix of independent businesses, studios, and agencies', 'Exceptional restaurant and nightlife scene'],
        ['Tottenham Court Road', 'Piccadilly Circus', 'Oxford Circus', 'Leicester Square'],
    ),
    'holborn': (
        'Holborn is a prime legal and professional services location, sitting at the crossroads of the City and the West End. It is home to London\'s four Inns of Court and many of the world\'s top law firms.',
        ['Home to London\'s legal community and major law firms', 'Central line and Piccadilly line connections', 'Between the City and the West End — walk to both', 'Excellent choice of serviced offices and law-focused spaces'],
        ['Holborn', 'Chancery Lane', 'Temple', 'Farringdon'],
    ),
    'north london': (
        'North London offers a compelling mix of vibrant residential areas, strong transport links, and growing business districts. From Islington\'s tech scene to Camden\'s creative energy, north London suits ambitious businesses seeking an alternative to the centre.',
        ['Victoria, Piccadilly, and Northern lines all serve north London', 'Growing tech and media presence in Islington and Camden', 'More affordable than central London with strong talent pools', 'Great access to King\'s Cross and the Elizabeth line'],
        ['Angel', 'King\'s Cross', 'Camden Town', 'Highbury & Islington'],
    ),
    'west london': (
        'West London offers prestigious addresses, excellent transport links, and a diverse business community. From Hammersmith\'s media hub to Chiswick\'s business parks, the area blends professional credibility with quality of life.',
        ['Piccadilly and District lines provide fast central access', 'Close to Heathrow for international travel', 'Strong media, tech, and professional services cluster', 'Green spaces and high quality of life for staff'],
        ['Hammersmith', 'Chiswick', 'Ealing Broadway', 'Acton'],
    ),
    'east london': (
        'East London has transformed into one of the most exciting business destinations in the UK. From Shoreditch\'s tech startups to Canary Wharf\'s finance giants, east London offers an unmatched diversity of workspace and talent.',
        ['Elizabeth line dramatically cuts journey times across London', 'Diverse business ecosystem from tech to finance', 'Competitive pricing with modern, design-led spaces', 'Thriving cultural and social scene'],
        ['Stratford', 'Bethnal Green', 'Whitechapel', 'Canary Wharf'],
    ),
    'south london': (
        'South London is increasingly attractive to businesses priced out of the north or west, offering excellent value, improving transport links, and a growing creative and digital scene centred around areas like Brixton and London Bridge.',
        ['Improving transport links via Thameslink and Overground', 'Vibrant areas: Brixton, Peckham, and the South Bank', 'Significantly lower rents than comparable north London locations', 'Strong community of independent and growing businesses'],
        ['Waterloo', 'London Bridge', 'Brixton', 'Clapham Junction'],
    ),
    'london': (
        'London is the world\'s most important business city — a global hub for finance, technology, law, media, and professional services. With over 500 flexible workspace options across dozens of boroughs, myHQ helps you find the perfect fit wherever in London you need to be.',
        ['World-class transport network: 11 Tube lines, Overground, Elizabeth line', 'Home to 40% of Europe\'s top 500 companies', 'Unrivalled access to talent from world-leading universities', 'Every major business district within 30 minutes by Tube'],
        ['Oxford Circus', 'Liverpool Street', 'King\'s Cross', 'London Bridge', 'Canary Wharf'],
    ),
}

def _get_location_data(loc_str):
    """Return (about_text, highlights, stations) for a given location string."""
    loc_lower = (loc_str or '').lower().strip()

    # 1. Exact match
    if loc_lower in _LOCATION_DATA:
        return _LOCATION_DATA[loc_lower]

    # 2. loc is a substring of a key (e.g. "old street" matches "old street")
    #    Prefer longer key matches (more specific) over shorter ones
    matches = [(key, data) for key, data in _LOCATION_DATA.items() if loc_lower in key]
    if matches:
        matches.sort(key=lambda x: -len(x[0]))  # longest key = most specific
        # But prefer exact matches / close matches over generic "london"
        non_generic = [(k, d) for k, d in matches if k != 'london']
        return (non_generic or matches)[0][1]

    # 3. Key is a substring of loc (e.g. "central london" contains "london")
    #    Again prefer specificity
    matches2 = [(key, data) for key, data in _LOCATION_DATA.items()
                if key in loc_lower and key != 'london']
    if matches2:
        matches2.sort(key=lambda x: -len(x[0]))
        return matches2[0][1]

    # 4. Word overlap
    loc_words = set(loc_lower.split())
    best_score, best_data = 0, None
    for key, data in _LOCATION_DATA.items():
        score = sum(1 for w in key.split() if w in loc_words)
        if score > best_score:
            best_score, best_data = score, data
    if best_data:
        return best_data
    # Generic fallback
    return (
        f'{loc_str} is a well-connected London business district offering a strong mix of flexible workspace options. '
        'The area benefits from excellent transport links and a growing professional community.',
        [
            'Good Tube and bus connectivity throughout',
            'Variety of flexible workspace options available',
            'Close to key London business districts',
            'Growing professional services and tech community',
        ],
        ['Check local TfL maps for nearest stations'],
    )


def _about_location_slide(proposal, all_centres, font, logo, map_img_path=None):
    loc = proposal.get('client_location') or 'London'
    sl = [R(0, 0, W, H, WHITE)]
    # NOTE: _small_logo is added LAST so it renders above the map image

    # Split layout: left info panel, right = map
    left_w  = 5200000
    map_x   = left_w + 300000
    map_w   = W - map_x - M
    map_y   = 200000
    map_h   = H - 400000

    # Map panel (right side) — drawn first so logo sits on top
    _map_src = map_img_path if (map_img_path and os.path.isfile(map_img_path)) else (MAP_IMG if os.path.isfile(MAP_IMG) else None)
    sl.append(I(_map_src, map_x, map_y, map_w, map_h))
    sl.append(R(map_x, map_y, map_w, 8000, BLUE))

    about_text, highlights, stations = _get_location_data(loc)

    # Left: title
    sl.append(T(f'About {loc}', M, 200000, left_w - M, 480000,
                22, BLUE, font=font, bold=True))

    sl.append(T(about_text, M, 730000, left_w - M, 900000,
                9, DARK, font=font, wrap=True))

    sl.append(T('KEY HIGHLIGHTS', M, 1720000, left_w - M, 250000,
                8, BLUE, font=font, bold=True))

    for hi, hl in enumerate(highlights[:4]):
        sl.append(T(f'•  {hl}', M + 80000, 1980000 + hi * 290000,
                    left_w - M - 80000, 260000, 9, DARK, font=font))

    # Popular stations section
    sl.append(T('POPULAR STATIONS', M, 3130000, left_w - M, 250000,
                8, BLUE, font=font, bold=True))

    station_x = M
    station_y = 3400000
    row_gap = 260000
    for si, stn in enumerate(stations[:5]):
        sy = station_y + si * row_gap
        badge_path = station_badge_path(stn)
        if badge_path:
            bw, bh = 580000, 174000   # 320:96 = 3.33:1 ratio
            sl.append(I(badge_path, station_x, sy, bw, bh))
            sl.append(T(stn, station_x + bw + 60000, sy + 20000,
                        left_w - M - bw - 60000, 200000, 9, DARK, font=font))
        else:
            sl.append(R(station_x, sy + 55000, 12000, 120000, BLUE))
            sl.append(T(stn, station_x + 80000, sy,
                        left_w - M - 80000, 220000, 9, DARK, font=font))

    # Shortlisted spaces
    list_y = station_y + len(stations[:5]) * row_gap + 150000
    if list_y < H - 1200000:
        sl.append(T('SHORTLISTED SPACES', M, list_y, left_w - M, 250000,
                    8, BLUE, font=font, bold=True))
        row_y = list_y + 270000
        for ni, c in enumerate(all_centres):
            if row_y > H - 350000:
                break
            name = c.get('name', '—')
            sl.append(T(f'{ni + 1}.  {name}',
                        M + 80000, row_y,
                        left_w - M - 1000000, 250000, 9, DARK, font=font))
            badge = tube_badge_path(c.get('transport', ''))
            if badge:
                sl.append(I(badge, left_w - 900000, row_y - 10000, 700000, 200000))
            row_y += 275000

    # Logo drawn last — renders on top of the map image
    sl.append(_small_logo(logo))
    return sl


def _make_feature_icon(idx, accent_hex):
    """Draw a clean circular feature icon using PIL and return the file path (cached)."""
    key = f'icon_{idx}_{accent_hex.lstrip("#")}'
    cache = os.path.join(BASE_DIR, 'uploads', f'.{key}.png')
    if os.path.exists(cache) and os.path.getsize(cache) > 200:
        return cache
    try:
        from PIL import Image as _PI, ImageDraw as _ID
        size = 120
        img = _PI.new('RGBA', (size, size), (0, 0, 0, 0))
        d = _ID.Draw(img)
        h = accent_hex.lstrip('#')
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        # Filled circle background
        d.ellipse([0, 0, size - 1, size - 1], fill=(r, g, b, 255))
        wc = (255, 255, 255, 255)
        cx, cy = size // 2, size // 2

        if idx == 0:
            # Checkmark
            pts = [(30, 62), (50, 82), (90, 38)]
            d.line(pts, fill=wc, width=10, joint='curve')
        elif idx == 1:
            # Location pin
            r2 = 20
            d.ellipse([cx - r2, cy - 40, cx + r2, cy], fill=wc)
            d.ellipse([cx - 10, cy - 30, cx + 10, cy - 10], fill=(r, g, b, 255))
            d.polygon([(cx - 18, cy - 8), (cx + 18, cy - 8), (cx, cy + 28)], fill=wc)
        elif idx == 2:
            # Zero cost — circle with cross (no-fee symbol)
            d.ellipse([28, 28, 92, 92], outline=wc, width=7)
            # Diagonal slash
            d.line([(42, 42), (78, 78)], fill=wc, width=7)
        elif idx == 3:
            # Handshake (two arcs)
            d.arc([25, 30, 60, 70], start=180, end=0, fill=wc, width=8)
            d.arc([60, 30, 95, 70], start=180, end=0, fill=wc, width=8)
            d.line([(25, 50), (95, 50)], fill=wc, width=3)
        elif idx == 4:
            # People (two circles + bodies)
            d.ellipse([20, 20, 44, 44], fill=wc)
            d.ellipse([76, 20, 100, 44], fill=wc)
            d.ellipse([45, 15, 75, 45], fill=wc)
            d.arc([5, 55, 55, 90], start=200, end=340, fill=wc, width=8)
            d.arc([65, 55, 115, 90], start=200, end=340, fill=wc, width=8)
            d.arc([32, 50, 88, 90], start=200, end=340, fill=wc, width=9)

        img.save(cache, 'PNG')
        return cache
    except Exception:
        return None


def _how_myhq_helps_slide(font, logo, accent_color=BLUE, icon_color=BLUE):
    """How myHQ helps – 3rd from last slide with icons and cropped workspace photo."""
    sl = [R(0, 0, W, H, WHITE)]
    sl.append(_small_logo(logo))

    # Title area
    sl.append(T('How myHQ helps?', M, 280000, 5400000, 520000,
                26, BLUE, font=font, bold=True))
    sl.append(T("Here's what makes us the most preferred brand for workspace solutions",
                M, 850000, 5000000, 380000,
                10, GREY, font=font))

    # Divider under title
    sl.append(L(M, 1280000, 5600000, 1280000, color='#E5E7EB', width=1))

    features = [
        ('Hassle-free process',
         'End-to-end support from search to move-in — we handle everything so you can focus on your business.'),
        ('Largest coverage',
         '500+ spaces across London, all vetted, toured, and ready for you to move in immediately.'),
        ('Zero brokerage fee',
         'No hidden charges, no brokerage. You pay only for the space — full transparency, always.'),
        ('Transparent & ethical',
         'Honest pricing, clear contracts, and a dedicated consultant who works in your best interest.'),
        ('For all team sizes',
         'Whether you are a solo founder or a team of 500, we match you to the perfect space.'),
    ]

    left_panel_w = 5600000
    feature_y_start = 1420000
    feature_row_h   = 1000000
    icon_size = 500000   # 500k EMU square icon

    for i, (title, desc) in enumerate(features):
        fy = feature_y_start + i * feature_row_h
        icon_path = _make_feature_icon(i, accent_color)
        if icon_path and os.path.isfile(icon_path):
            sl.append(I(icon_path, M, fy + 30000, icon_size, icon_size))
        else:
            sl.append(R(M, fy + 30000, icon_size, icon_size, accent_color))

        sl.append(T(title, M + icon_size + 80000, fy + 40000,
                    left_panel_w - M - icon_size - 80000, 260000,
                    11, DARK, font=font, bold=True))
        sl.append(T(desc, M + icon_size + 80000, fy + 300000,
                    left_panel_w - M - icon_size - 80000, 460000,
                    9, GREY, font=font, wrap=True))

        if i < len(features) - 1:
            sl.append(L(M + icon_size + 80000, fy + feature_row_h - 60000,
                        left_panel_w, fy + feature_row_h - 60000,
                        color='#F0F0F0', width=1))

    # Right photo panel – cropped workspace image
    photo_x = left_panel_w + 200000
    photo_w = W - photo_x - 100000
    crop_photo = CROP_IMG if os.path.isfile(CROP_IMG) else None
    sl.append(I(crop_photo, photo_x, 0, photo_w, H))

    return sl


# ── London slide deck ───────────────────────────────────────────────────────────

def build_london_slides(proposal, db_centres, manual_centres):
    F       = 'Calibri'
    logo    = get_logo_png()
    logo_w  = get_logo_png(white=True)
    all_centres = db_centres + manual_centres
    slides = []

    # ── Slide 1: Cover ──────────────────────────────────────────────────────────
    sl = []
    # Left panel (blush)
    sl.append(R(0, 0, 5100000, H, BLUSH))
    # Right photo (city skyline, high-res)
    cover_photo = COVER_IMG if os.path.isfile(COVER_IMG) else None
    sl.append(I(cover_photo, 5100000, 0, W - 5100000, H))
    # myHQ logo top-left composited on blush panel
    sl.append(I(_logo_on(logo, BLUSH), M, 280000, 1700000, 637000))
    # Title
    client = proposal.get('client_company') or proposal.get('client_name') or 'Your Company'
    sl.append(T(f'Workspace Proposal\nfor {client}',
                M, 2100000, 4600000, 1800000,
                30, BLUE, font=F, bold=True))
    # Subtitle
    sl.append(T('Flexible workspaces, transparent pricing, no brokerage',
                M, 3950000, 4600000, 450000,
                11, GREY, font=F))
    # Client name bottom
    sl.append(T(client, M, 5650000, 4600000, 450000,
                16, BLUE, font=F, bold=True))
    slides.append(sl)

    # ── Slide 2: Comparison Table ────────────────────────────────────────────────
    slides.append(
        _comparison_table_slide(all_centres,
                                header_fill=AQUA,
                                header_text_color=DARK,
                                font=F, logo=logo))

    # ── Slide 3: Client Requirements ────────────────────────────────────────────
    slides.append(
        _client_requirements_slide(
            proposal, 'london', db_centres, manual_centres,
            font=F, logo=logo,
            label_fill=AQUA,
            label_text=DARK,
            date_label_color=MINT))

    # ── Slide 4: About Location ──────────────────────────────────────────────────
    _map_tmp = os.path.join(BASE_DIR, 'uploads', 'proposals', f'map_{id(proposal)}.png')
    _map_path = generate_proposal_map(all_centres, _map_tmp) or MAP_IMG
    slides.append(_about_location_slide(proposal, all_centres, F, logo, map_img_path=_map_path))

    # ── Slides 5+: Centre slides ─────────────────────────────────────────────────
    for idx, centre in enumerate(all_centres, 1):
        slides.append(_centre_slide(idx, centre, 'london', F, logo))

    # ── 3rd from last: How myHQ helps ───────────────────────────────────────────
    slides.append(_how_myhq_helps_slide(F, logo, accent_color=AQUA, icon_color=BLUE))

    # ── Existing Clients ────────────────────────────────────────────────────────
    slides.append(_existing_clients_slide(F, logo))

    # ── Testimonials ────────────────────────────────────────────────────────────
    slides.append(_testimonials_slide(F, logo, bg_color=WHITE, heading_color=DARK))

    return slides


# ── India slide deck ────────────────────────────────────────────────────────────

def build_india_slides(proposal, db_centres, manual_centres):
    F       = 'Calibri'
    logo    = get_logo_png()
    logo_w  = get_logo_png(white=True)
    all_centres = db_centres + manual_centres
    slides = []

    # ── Slide 1: Cover ──────────────────────────────────────────────────────────
    sl = []
    sl.append(R(0, 0, 5100000, H, WHITE))
    cover_photo = COVER_IMG if os.path.isfile(COVER_IMG) else None
    sl.append(I(cover_photo, 5100000, 0, W - 5100000, H))
    sl.append(I(_logo_on(logo, WHITE), M, 280000, 1700000, 637000))
    sl.append(T("The search for your\nperfect workspace\nends here",
                M, 2100000, 4600000, 1800000,
                30, BLUE, font=F, bold=True))
    sl.append(T("India's leading marketplace for flexible workspace solutions",
                M, 3950000, 4600000, 450000,
                11, GREY, font=F))
    client = proposal.get('client_company') or proposal.get('client_name') or 'Your Company'
    sl.append(T(client, M, 5650000, 4600000, 450000,
                16, DARK, font=F, bold=True))
    slides.append(sl)

    # ── Slide 2: Comparison Table ────────────────────────────────────────────────
    slides.append(
        _comparison_table_slide(all_centres,
                                header_fill=BLUE,
                                header_text_color=WHITE,
                                font=F, logo=logo))

    # ── Slide 3: Client Requirements ────────────────────────────────────────────
    slides.append(
        _client_requirements_slide(
            proposal, 'india', db_centres, manual_centres,
            font=F, logo=logo,
            label_fill=BLUE,
            label_text=WHITE,
            date_label_color='#AAAAFF'))

    # ── Slide 4: About Location ──────────────────────────────────────────────────
    _map_tmp = os.path.join(BASE_DIR, 'uploads', 'proposals', f'map_{id(proposal)}.png')
    _map_path = generate_proposal_map(all_centres, _map_tmp) or MAP_IMG
    slides.append(_about_location_slide(proposal, all_centres, F, logo, map_img_path=_map_path))

    # ── Slides 5+: Centre slides ─────────────────────────────────────────────────
    for idx, centre in enumerate(all_centres, 1):
        slides.append(_centre_slide(idx, centre, 'india', F, logo))

    # ── 3rd from last: How myHQ helps ───────────────────────────────────────────
    slides.append(_how_myhq_helps_slide(F, logo, accent_color=BLUE, icon_color=BLUE))

    # ── Existing Clients ────────────────────────────────────────────────────────
    slides.append(_existing_clients_slide(F, logo))

    # ── Testimonials ────────────────────────────────────────────────────────────
    slides.append(_testimonials_slide(F, logo, bg_color=WHITE, heading_color=DARK))

    return slides


# ── PPTX renderer ──────────────────────────────────────────────────────────────

def _rgb(hex_str):
    from pptx.dml.color import RGBColor
    h = hex_str.lstrip('#')
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def render_pptx(slides, out_path):
    from pptx import Presentation
    from pptx.util import Emu, Pt
    from pptx.enum.text import PP_ALIGN
    from pptx.util import Emu as EMU

    prs = Presentation()
    prs.slide_width  = Emu(W)
    prs.slide_height = Emu(H)
    blank_layout = prs.slide_layouts[6]  # completely blank

    ALIGN_MAP = {
        'L': PP_ALIGN.LEFT,
        'C': PP_ALIGN.CENTER,
        'R': PP_ALIGN.RIGHT,
    }

    for slide_cmds in slides:
        sl = prs.slides.add_slide(blank_layout)
        for cmd in slide_cmds:
            ct = cmd['cmd']

            if ct == 'rect':
                x = Emu(cmd['x']); y = Emu(cmd['y'])
                w = Emu(cmd['w']); h = Emu(cmd['h'])
                sh = sl.shapes.add_shape(1, x, y, w, h)
                sh.line.fill.background()
                sh.fill.solid()
                sh.fill.fore_color.rgb = _rgb(cmd['fill'])

            elif ct == 'text':
                if not str(cmd.get('text', '')).strip():
                    continue
                x = Emu(cmd['x']); y = Emu(cmd['y'])
                w = Emu(cmd['w']); h = Emu(cmd['h'])
                tb = sl.shapes.add_textbox(x, y, w, h)
                tf = tb.text_frame
                tf.word_wrap = cmd.get('wrap', True)
                align = ALIGN_MAP.get(cmd.get('align', 'L'), PP_ALIGN.LEFT)
                lines = cmd['text'].split('\n')
                first = True
                for line in lines:
                    if first:
                        p = tf.paragraphs[0]
                        first = False
                    else:
                        p = tf.add_paragraph()
                    p.alignment = align
                    run = p.add_run()
                    run.text = line
                    run.font.name    = cmd.get('font', 'Calibri')
                    run.font.size    = Pt(cmd['size'])
                    run.font.color.rgb = _rgb(cmd['color'])
                    run.font.bold    = cmd.get('bold', False)
                    run.font.italic  = cmd.get('italic', False)

            elif ct == 'img':
                x = Emu(cmd['x']); y = Emu(cmd['y'])
                w = Emu(cmd['w']); h = Emu(cmd['h'])
                path = cmd.get('path')
                if path and os.path.isfile(str(path)):
                    try:
                        from PIL import Image as _PILImg
                        import io as _io2
                        _im = _PILImg.open(str(path))
                        # Normalise palette/odd modes to PNG-safe RGB or RGBA
                        if _im.mode == 'P':
                            _im = _im.convert('RGBA')
                        use_path = str(path)
                        # If image was changed, write to a temp buffer and use that
                        if _im.mode not in ('RGB', 'RGBA', 'L'):
                            _im = _im.convert('RGB')
                        if _im.mode == 'RGBA':
                            # Keep for PPTX (transparency is OK), but re-save to fix palette artefacts
                            _buf = _io2.BytesIO()
                            _im.save(_buf, 'PNG')
                            _buf.seek(0)
                            iw, ih = _im.size
                        else:
                            iw, ih = _im.size
                        tw, th = cmd['w'], cmd['h']
                        # Cover-crop fractions
                        scale = max(tw / iw, th / ih)
                        sw_px = iw * scale
                        sh_px = ih * scale
                        cx = max(0.0, (sw_px - tw) / (2 * sw_px))
                        cy = max(0.0, (sh_px - th) / (2 * sh_px))
                        if _im.mode == 'RGBA':
                            pic = sl.shapes.add_picture(_buf, x, y, w, h)
                        else:
                            pic = sl.shapes.add_picture(use_path, x, y, w, h)
                        pic.crop_left   = cx
                        pic.crop_right  = cx
                        pic.crop_top    = cy
                        pic.crop_bottom = cy
                        continue
                    except Exception:
                        pass
                # grey placeholder
                sh = sl.shapes.add_shape(1, x, y, w, h)
                sh.line.fill.background()
                sh.fill.solid()
                sh.fill.fore_color.rgb = _rgb('#DDDDDD')

            elif ct == 'line':
                x1 = Emu(cmd['x1']); y1 = Emu(cmd['y1'])
                x2 = Emu(cmd['x2']); y2 = Emu(cmd['y2'])
                # Represent as a thin rectangle
                lw = max(Emu(cmd.get('width', 1) * 12700), Emu(12700))
                if x1 == x2:
                    sh = sl.shapes.add_shape(1, x1, y1, lw, Emu(abs(cmd['y2'] - cmd['y1'])))
                else:
                    sh = sl.shapes.add_shape(1, x1, y1, Emu(abs(cmd['x2'] - cmd['x1'])), lw)
                sh.line.fill.background()
                sh.fill.solid()
                sh.fill.fore_color.rgb = _rgb(cmd.get('color', '#E5E7EB'))

    prs.save(out_path)
    return out_path


# ── PDF renderer ────────────────────────────────────────────────────────────────

def render_pdf(slides, out_path):
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.lib.utils import ImageReader
    from reportlab.lib.colors import HexColor
    from reportlab.lib.utils import simpleSplit

    PW, PH  = 960.0, 540.0   # points (13.33" × 7.5" @ 72 dpi)
    E2P     = 1.0 / 12700.0  # EMU → pt

    def ep(v):
        return float(v) * E2P

    def hex_color(h):
        return HexColor(f'#{h.lstrip("#")}')

    def wrap_lines(text, fn, fs, max_w):
        lines = []
        for para in text.split('\n'):
            try:
                wrapped = simpleSplit(para, fn, fs, max_w)
                lines.extend(wrapped if wrapped else [''])
            except Exception:
                lines.extend(para.split('\n') or [''])
        return lines

    c = rl_canvas.Canvas(out_path, pagesize=(PW, PH))

    for slide_cmds in slides:
        for cmd in slide_cmds:
            ct = cmd['cmd']
            sx  = ep(cmd.get('x', 0))
            sy  = ep(cmd.get('y', 0))   # top in EMU space (pt)
            sw  = ep(cmd.get('w', 0))
            sh  = ep(cmd.get('h', 0))
            # PDF y=0 is bottom-left; our y is measured from top
            pdf_y = PH - sy - sh        # bottom of box in PDF coords

            if ct == 'rect':
                col = hex_color(cmd['fill'])
                c.setFillColor(col)
                c.setStrokeColor(col)
                c.rect(sx, pdf_y, sw, sh, fill=1, stroke=0)

            elif ct == 'text':
                text = str(cmd.get('text', ''))
                if not text.strip():
                    continue
                fs     = float(cmd['size'])
                bold   = cmd.get('bold', False)
                italic = cmd.get('italic', False)
                if bold and italic:   fn = 'Helvetica-BoldOblique'
                elif bold:            fn = 'Helvetica-Bold'
                elif italic:          fn = 'Helvetica-Oblique'
                else:                 fn = 'Helvetica'

                c.setFont(fn, fs)
                c.setFillColor(hex_color(cmd['color']))
                align  = cmd.get('align', 'L')
                do_wrap = cmd.get('wrap', True)

                if do_wrap:
                    lines = wrap_lines(text, fn, fs, sw - 4)
                else:
                    lines = text.split('\n')

                line_h = fs * 1.35
                # Start at near the top of the box
                text_top = PH - sy - fs * 0.85
                for li, line in enumerate(lines):
                    ty = text_top - li * line_h
                    if ty < pdf_y:          # clipped
                        break
                    if align == 'C':
                        c.drawCentredString(sx + sw / 2, ty, line)
                    elif align == 'R':
                        c.drawRightString(sx + sw, ty, line)
                    else:
                        c.drawString(sx + 2, ty, line)

            elif ct == 'img':
                path = cmd.get('path')
                if path and os.path.isfile(str(path)):
                    try:
                        from PIL import Image as _PILImg
                        import io as _io
                        _raw = _PILImg.open(str(path))
                        # Convert palette/RGBA modes safely to RGB
                        if _raw.mode in ('P', 'RGBA', 'LA'):
                            _raw = _raw.convert('RGBA')
                            bg = _PILImg.new('RGB', _raw.size, (255, 255, 255))
                            bg.paste(_raw, mask=_raw.split()[3])
                            _im = bg
                        else:
                            _im = _raw.convert('RGB')
                        iw, ih = _im.size
                        # Cover-crop: trim to exact target aspect ratio (center)
                        target_ratio = sw / sh if sh > 0 else 1.0
                        src_ratio    = iw / ih if ih > 0 else 1.0
                        if src_ratio > target_ratio:
                            new_w = int(ih * target_ratio)
                            x0 = (iw - new_w) // 2
                            _im = _im.crop((x0, 0, x0 + new_w, ih))
                        elif src_ratio < target_ratio:
                            new_h = int(iw / target_ratio)
                            y0 = (ih - new_h) // 2
                            _im = _im.crop((0, y0, iw, y0 + new_h))
                        buf = _io.BytesIO()
                        _im.save(buf, 'JPEG', quality=90)
                        buf.seek(0)
                        ir = ImageReader(buf)
                        c.drawImage(ir, sx, pdf_y, sw, sh,
                                    preserveAspectRatio=False, mask='auto')
                        continue
                    except Exception:
                        pass
                # grey placeholder
                c.setFillColor(HexColor('#DDDDDD'))
                c.setStrokeColor(HexColor('#CCCCCC'))
                c.rect(sx, pdf_y, sw, sh, fill=1, stroke=1)

            elif ct == 'line':
                x1 = ep(cmd['x1']); y1 = PH - ep(cmd['y1'])
                x2 = ep(cmd['x2']); y2 = PH - ep(cmd['y2'])
                lw = float(cmd.get('width', 1))
                c.setStrokeColor(hex_color(cmd.get('color', '#E5E7EB')))
                c.setLineWidth(lw)
                c.line(x1, y1, x2, y2)

        c.showPage()

    c.save()
    return out_path
