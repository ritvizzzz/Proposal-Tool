"""
Bulk-fetch Google Maps photos for all centres in the DB.
Run: python3 fetch_photos.py
Skips centres that already have images.
"""
import os, sys, sqlite3, time
from playwright.sync_api import sync_playwright

DB            = os.path.join(os.path.dirname(__file__), 'proposals.db')
CENTRE_IMAGES = os.path.join(os.path.dirname(__file__), 'uploads', 'centres')
MAX_PHOTOS    = 4
DELAY_S       = 1.2   # polite delay between searches

def fetch_centre(ctx, centre_id, name, address):
    out_dir = os.path.join(CENTRE_IMAGES, str(centre_id))
    os.makedirs(out_dir, exist_ok=True)

    page = ctx.new_page()
    try:
        query = f"{name} {address}".replace(' ', '+')
        page.goto(f'https://www.google.com/maps/search/{query}', timeout=20000)
        page.wait_for_timeout(2000)

        urls = page.eval_on_selector_all('img', """els => [...new Set(
            els.map(e=>e.src).filter(s =>
                s && s.includes('googleusercontent') &&
                !s.includes('photo_profile') && !s.includes('=s')
            )
        )]""")

        saved = []
        for i, url in enumerate(urls[:MAX_PHOTOS]):
            sized = url.split('=')[0] + '=w800-h600-k-no'
            try:
                resp = ctx.request.get(sized, timeout=10000)
                if resp.ok:
                    data = resp.body()
                    if len(data) > 5000:
                        fname = f'gmaps_{i+1}_{os.urandom(4).hex()}.jpg'
                        with open(os.path.join(out_dir, fname), 'wb') as f:
                            f.write(data)
                        saved.append(fname)
            except Exception:
                pass
        return saved
    except Exception as e:
        return []
    finally:
        page.close()


def main():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row

    centres = conn.execute(
        'SELECT c.id, c.name, c.address FROM centres c '
        'WHERE NOT EXISTS (SELECT 1 FROM centre_images ci WHERE ci.centre_id=c.id) '
        'ORDER BY c.id'
    ).fetchall()

    print(f'{len(centres)} centres without photos — fetching now...')

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        ctx = browser.new_context(
            viewport={'width': 1280, 'height': 800},
            user_agent='Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36'
        )

        for idx, c in enumerate(centres, 1):
            print(f'[{idx}/{len(centres)}] {c["name"]}', end=' ... ', flush=True)
            saved = fetch_centre(ctx, c['id'], c['name'], c['address'])

            if saved:
                with conn:
                    for i, fname in enumerate(saved):
                        conn.execute(
                            'INSERT INTO centre_images (centre_id, filename, is_primary, sort_order) VALUES (?,?,?,?)',
                            (c['id'], fname, 1 if i == 0 else 0, i)
                        )
                print(f'{len(saved)} photos saved')
            else:
                print('no photos found')

            time.sleep(DELAY_S)

        browser.close()

    conn.close()
    print('Done.')


if __name__ == '__main__':
    main()
