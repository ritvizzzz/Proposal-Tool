#!/usr/bin/env python3
"""
Scrape Google Maps ratings + review counts for all Hubble centres.
Saves progress to google_reviews.json after every 10 centres.
Run: python3 scrape_google_reviews.py
Can be safely interrupted and restarted — skips already-scraped entries.
"""

import json, os, re, time, sqlite3, random

BASE = os.path.dirname(os.path.abspath(__file__))
DB   = os.path.join(BASE, 'proposals.db')
OUT  = os.path.join(BASE, 'google_reviews.json')

DELAY_MIN = 2.0   # seconds between searches
DELAY_MAX = 4.0
SAVE_EVERY = 10   # save progress every N centres

def load_progress():
    if os.path.exists(OUT):
        with open(OUT) as f:
            return json.load(f)
    return {}

def save_progress(data):
    with open(OUT, 'w') as f:
        json.dump(data, f, indent=2)

def load_centres():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    rows = conn.execute(
        'SELECT id, hubble_id, name, address FROM centres ORDER BY id'
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]

def search_google_maps(page, name, address):
    """Search Google Maps and return (rating, review_count) or (None, None)."""
    query = f"{name} {address}"
    url   = f"https://www.google.com/maps/search/{query.replace(' ', '+')}"
    try:
        page.goto(url, timeout=20000, wait_until='domcontentloaded')
        page.wait_for_timeout(3000)

        # Dismiss any consent/cookie dialog
        for sel in ['button[aria-label*="Accept"]', 'button[aria-label*="Agree"]',
                    'form[action*="consent"] button']:
            try:
                el = page.query_selector(sel)
                if el and el.is_visible():
                    el.click(); page.wait_for_timeout(1500); break
            except Exception:
                pass

        # If search returned a list, click the first result
        first_result = page.query_selector('a.hfpxzc')
        if first_result:
            first_result.click()
            page.wait_for_timeout(3000)

        # Try to extract rating from the page
        rating = None
        review_count = None

        # Get full page source + visible text once
        try:
            page_html = page.content()
            visible   = page.inner_text('body')
        except Exception:
            page_html = ''
            visible   = ''

        # Method 1: JSON-LD / structured data (most reliable)
        m = re.search(r'"ratingValue"\s*:\s*"?([\d.]+)', page_html)
        if m:
            rating = float(m.group(1))
        m2 = re.search(r'"reviewCount"\s*:\s*"?(\d+)', page_html)
        if m2:
            review_count = int(m2.group(1))
        if not review_count:
            m3 = re.search(r'"userRatingCount"\s*:\s*(\d+)', page_html)
            if m3:
                review_count = int(m3.group(1))

        # Method 2: aria-label on rating element
        if rating is None:
            for sel in ['[aria-label*="star"]', '[aria-label*="Star"]',
                        'span.fontDisplayLarge', 'div.F7nice span', '.DkEaL']:
                try:
                    el = page.query_selector(sel)
                    if el:
                        txt = el.get_attribute('aria-label') or el.inner_text()
                        mr = re.search(r'([\d.]+)\s*star', txt, re.IGNORECASE)
                        if mr:
                            rating = float(mr.group(1))
                            rc = re.search(r'([\d,]+)\s*review', txt, re.IGNORECASE)
                            if rc:
                                review_count = int(rc.group(1).replace(',', ''))
                            break
                except Exception:
                    pass

        # Method 3: visible text pattern "4.2 (1,234)" or "4.2\n1,234 reviews"
        if rating is None or not review_count:
            # Rating from visible text
            if rating is None:
                mr = re.search(r'\b([1-5]\.[0-9])\s*\n', visible)
                if mr:
                    rating = float(mr.group(1))
            # Review count — look for "(1,234)" or "1,234 reviews"
            if not review_count:
                rc = re.search(r'\(([\d,]+)\s*\)', visible)
                if rc:
                    try:
                        review_count = int(rc.group(1).replace(',', ''))
                    except Exception:
                        pass
            if not review_count:
                rc2 = re.search(r'([\d,]+)\s+review', visible, re.IGNORECASE)
                if rc2:
                    try:
                        review_count = int(rc2.group(1).replace(',', ''))
                    except Exception:
                        pass

        return rating, review_count

    except Exception as e:
        print(f'    Error: {e}')
        return None, None


def main():
    from playwright.sync_api import sync_playwright

    centres = load_centres()
    progress = load_progress()

    todo = [c for c in centres if c['hubble_id'] not in progress]
    done = len(centres) - len(todo)
    print(f'=== Google Maps Review Scraper ===')
    print(f'Total centres: {len(centres)} | Already done: {done} | Remaining: {len(todo)}')

    if not todo:
        print('All done! Run update_reviews_in_map.py to apply.')
        return

    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            headless=True,
            args=['--no-sandbox', '--disable-blink-features=AutomationControlled']
        )
        ctx = browser.new_context(
            viewport={'width': 1280, 'height': 800},
            user_agent=(
                'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) '
                'AppleWebKit/537.36 (KHTML, like Gecko) '
                'Chrome/120.0.0.0 Safari/537.36'
            ),
            locale='en-GB',
        )
        page = ctx.new_page()

        saved_count = 0
        for idx, c in enumerate(todo, 1):
            hid  = c['hubble_id']
            name = c['name']
            addr = c['address']
            print(f'[{done+idx}/{len(centres)}] {name[:55]}...', end=' ', flush=True)

            # Use name + first part of address for search
            short_addr = addr.split(',')[0] if ',' in addr else addr
            rating, reviews = search_google_maps(page, name, short_addr)

            progress[hid] = {
                'hubble_id':    hid,
                'name':         name,
                'rating':       rating,
                'review_count': reviews,
            }

            if rating:
                print(f'★ {rating} ({reviews or "?"} reviews)')
            else:
                print('no rating found')

            saved_count += 1
            if saved_count % SAVE_EVERY == 0:
                save_progress(progress)
                print(f'  → Progress saved ({len(progress)} total)')

            time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))

        browser.close()

    save_progress(progress)
    found = sum(1 for v in progress.values() if v['rating'])
    print(f'\n✓ Done. {found}/{len(progress)} centres have ratings.')
    print(f'  Results saved to: {OUT}')
    print(f'  Run: python3 update_reviews_in_map.py')


if __name__ == '__main__':
    main()
