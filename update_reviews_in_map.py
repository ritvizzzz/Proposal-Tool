#!/usr/bin/env python3
"""
After scrape_google_reviews.py finishes, run this to bake ratings into
the netlify map's WS data so locality search ranks by reviews + price.
"""
import json, os, re

BASE     = os.path.dirname(os.path.abspath(__file__))
REVIEWS  = os.path.join(BASE, 'google_reviews.json')
MAP_HTML = os.path.join(BASE, '..', 'netlify-deploy', 'index.html')

def main():
    if not os.path.exists(REVIEWS):
        print('No google_reviews.json found. Run scrape_google_reviews.py first.')
        return

    with open(REVIEWS) as f:
        reviews = json.load(f)   # keyed by hubble_id

    with open(MAP_HTML, encoding='utf-8') as f:
        html = f.read()

    m = re.search(r'const WS = (\[.*?\]);', html, re.DOTALL)
    data = json.loads(m.group(1))

    updated = 0
    for ws in data:
        hid = str(ws.get('hubble_id', ''))
        r = reviews.get(hid)
        if r and r.get('rating'):
            ws['googleRating']  = r['rating']
            ws['reviewCount']   = r.get('review_count') or 0
        else:
            ws['googleRating']  = None
            ws['reviewCount']   = 0
        updated += 1

    ws_json = json.dumps(data, ensure_ascii=False, separators=(',', ':'))
    new_html = re.sub(
        r'const WS\s*=\s*\[.*?\];',
        f'const WS = {ws_json};',
        html, count=1, flags=re.DOTALL,
    )

    with open(MAP_HTML, 'w', encoding='utf-8') as f:
        f.write(new_html)

    with_ratings = sum(1 for ws in data if ws.get('googleRating'))
    print(f'✓ Updated {updated} spaces, {with_ratings} have Google ratings.')
    print(f'  Now update the locality scoring in the map JS to use googleRating + reviewCount.')

if __name__ == '__main__':
    main()
