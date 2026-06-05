"""
Import partner workspace data from Excel files into proposals.db.
Handles IWG/Regus brands and Fora centres.
"""

import sqlite3
import json
import re
import pandas as pd

DB_PATH = "/Users/radhikasharma/Documents/Claude/Proposal Tool/proposals.db"

REGUS_FILE = "/Users/radhikasharma/Documents/London workspaces/regus_london_all_82.xlsx"
FORA_FILES = [
    "/Users/radhikasharma/Documents/London workspaces/Fora/fora-no-links-1.xlsx",
    "/Users/radhikasharma/Documents/London workspaces/Fora/fora-no-links-2.xlsx",
    "/Users/radhikasharma/Documents/London workspaces/Fora/fora-no-links-3.xlsx",
]

IWG_BRAND_SLUGS = {"regus", "spaces", "hq", "signature", "the-clubhouse"}

BRAND_MAP = {
    "regus": "Regus",
    "spaces": "Spaces",
    "hq": "HQ by Regus",
    "signature": "Signature by Regus",
    "the-clubhouse": "The Clubhouse by Regus",
    "fora": "Fora",
}

# Type label mapping for transport
TYPE_LABEL = {
    "METRO": "",          # tube/metro - no prefix needed
    "LOCAL": "",          # overground/rail - no prefix
    "BUS": "Bus ",
    "TRAM": "Tram ",
}

LINE_FRIENDLY = {
    "NORTHERN": "Northern",
    "CENTRAL": "Central",
    "CIRCLE": "Circle",
    "DISTRICT": "District",
    "JUBILEE": "Jubilee",
    "VICTORIA": "Victoria",
    "PICCADILLY": "Piccadilly",
    "BAKERLOO": "Bakerloo",
    "METROPOLITAN": "Metropolitan",
    "HAMMERSMITH_AND_CITY": "Hammersmith & City",
    "ELIZABETH": "Elizabeth",
    "LONDON_OVERGROUND": "Overground",
    "NATIONAL_RAIL": "National Rail",
    "DLR": "DLR",
    "TRAM": "Tram",
    "THAMESLINK": "Thameslink",
    "SOUTHEASTERN": "Southeastern",
    "SOUTHWESTERN": "Southwestern",
    "CHILTERN_RAILWAYS": "Chiltern Railways",
    "GREAT_WESTERN_RAILWAY": "Great Western Railway",
    "CROSSRAIL": "Elizabeth",
}


def meters_to_walk_time(meters):
    """Convert metres to approximate walk time (80m/min)."""
    try:
        m = float(meters)
        mins = round(m / 80)
        if mins < 1:
            mins = 1
        return mins
    except (ValueError, TypeError):
        return None


def miles_to_walk_time(miles):
    """Convert miles to approximate walk time (3mph walking)."""
    try:
        m = float(miles)
        mins = round(m * 60 / 3)
        if mins < 1:
            mins = 1
        return mins
    except (ValueError, TypeError):
        return None


def parse_transport(connectivity_str, separator=","):
    """
    Parse connectivity details string into a human-readable transport string.
    Format: TYPE:STATION:LINE:DISTANCE  (multiple entries separated by separator)
    Returns the top 2-3 unique station entries as a readable string.
    """
    if not connectivity_str or pd.isna(connectivity_str):
        return None

    raw = str(connectivity_str).strip()
    entries = [e.strip() for e in raw.split(separator) if e.strip()]

    seen_stations = {}  # station -> (line, walk_mins)

    for entry in entries:
        parts = entry.split(":")
        if len(parts) < 4:
            continue
        transport_type = parts[0].strip().upper()
        station = parts[1].strip()
        line = parts[2].strip().upper()
        distance = parts[3].strip()

        # Determine walk time based on separator used (semicolon = meters, comma = miles)
        if separator == ";":
            mins = meters_to_walk_time(distance)
        else:
            mins = miles_to_walk_time(distance)

        if mins is None:
            continue

        # Keep the closest (lowest mins) entry per station
        if station not in seen_stations or mins < seen_stations[station][1]:
            friendly_line = LINE_FRIENDLY.get(line, line.replace("_", " ").title())
            seen_stations[station] = (friendly_line, mins)

    if not seen_stations:
        return None

    # Sort by walk time, take top 3
    sorted_stations = sorted(seen_stations.items(), key=lambda x: x[1][1])[:3]

    parts_out = []
    for station, (line, mins) in sorted_stations:
        parts_out.append(f"{station} — {mins} min walk")

    return " · ".join(parts_out)


def load_excel_data(filepath, brand_filter=None):
    """
    Load workspace data from an Excel file.
    Returns list of dicts ready for DB insertion.
    brand_filter: set of brand slugs to include (None = include all)
    """
    xl = pd.ExcelFile(filepath)

    # Load workspace sheet
    ws = pd.read_excel(xl, sheet_name="workspace")

    # Filter by brand if needed
    if brand_filter:
        ws = ws[ws["workspaceBrand_slug"].isin(brand_filter)]

    if ws.empty:
        return []

    # Load price plans - get minimum price per workspace
    prices = pd.read_excel(xl, sheet_name="dedicated_price_plans")
    # Keep only the columns we need
    prices = prices[["workspace_identifier", "amount", "paymentCycle", "seatType"]].copy()
    prices["amount"] = pd.to_numeric(prices["amount"], errors="coerce")
    # Get minimum price per workspace
    min_prices = (
        prices.dropna(subset=["amount"])
        .sort_values("amount")
        .drop_duplicates(subset=["workspace_identifier"], keep="first")
        .set_index("workspace_identifier")[["amount", "paymentCycle", "seatType"]]
    )

    # Load amenities
    amenities_df = pd.read_excel(xl, sheet_name="dedicated_amenities")
    amenities_grouped = (
        amenities_df.groupby("workspace_identifier")["amenity_slug"]
        .apply(list)
        .to_dict()
    )

    # Detect separator for connectivity details
    sample_connectivity = ws["directions_connectivityDetails"].dropna().head(5).tolist()
    separator = ";"
    for s in sample_connectivity:
        if ";" in str(s):
            separator = ";"
            break
        elif "," in str(s):
            separator = ","
            break

    results = []
    for _, row in ws.iterrows():
        wid = row["workspace_identifier"]
        brand_slug = str(row.get("workspaceBrand_slug", "")).strip()
        brand = BRAND_MAP.get(brand_slug, brand_slug)

        # Price info
        price_from = None
        price_unit = "MONTHLY"
        seat_type = None
        if wid in min_prices.index:
            price_from = int(min_prices.loc[wid, "amount"]) if pd.notna(min_prices.loc[wid, "amount"]) else None
            price_unit = str(min_prices.loc[wid, "paymentCycle"]) if pd.notna(min_prices.loc[wid, "paymentCycle"]) else "MONTHLY"
            seat_type = str(min_prices.loc[wid, "seatType"]) if pd.notna(min_prices.loc[wid, "seatType"]) else None

        # Amenities
        amen_list = amenities_grouped.get(wid, [])
        amenities_json = json.dumps(amen_list) if amen_list else None

        # Transport
        connectivity = row.get("directions_connectivityDetails")
        transport = parse_transport(connectivity, separator=separator)

        # Coordinates: "lng;lat" or "lng,lat" → store as-is
        coords = row.get("loc_coordinates")
        coordinates = str(coords).strip() if coords and pd.notna(coords) else None

        results.append({
            "name": str(row["name"]).strip() if pd.notna(row["name"]) else None,
            "address": str(row["address"]).strip() if pd.notna(row.get("address")) else None,
            "city": str(row.get("city_slug", "london")).strip(),
            "about": str(row["about"]).strip() if pd.notna(row.get("about")) else None,
            "space_type": str(row.get("spaceType", "")).strip() if pd.notna(row.get("spaceType")) else None,
            "brand": brand,
            "price_from": price_from,
            "price_unit": price_unit,
            "seat_type": seat_type,
            "amenities": amenities_json,
            "transport": transport,
            "map_url": str(row.get("mapurl", "")).strip() if pd.notna(row.get("mapurl")) else None,
            "coordinates": coordinates,
            "source": "myhq",
        })

    return results


def upsert_centre(cur, centre):
    """
    Insert or update a centre in the DB.
    Returns 'inserted' or 'updated'.
    """
    # Check if exists by name
    cur.execute("SELECT id, price_from FROM centres WHERE name = ?", (centre["name"],))
    existing = cur.fetchone()

    if existing:
        existing_id, existing_price = existing
        # Update fields, but only update price_from if currently NULL
        new_price = centre["price_from"] if existing_price is None else existing_price

        cur.execute("""
            UPDATE centres SET
                address = ?,
                about = ?,
                amenities = ?,
                transport = ?,
                coordinates = ?,
                map_url = ?,
                price_from = ?
            WHERE id = ?
        """, (
            centre["address"],
            centre["about"],
            centre["amenities"],
            centre["transport"],
            centre["coordinates"],
            centre["map_url"],
            new_price,
            existing_id,
        ))
        return "updated"
    else:
        cur.execute("""
            INSERT INTO centres
                (name, address, city, about, space_type, brand, price_from, price_unit,
                 seat_type, amenities, transport, map_url, coordinates, source)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            centre["name"],
            centre["address"],
            centre["city"],
            centre["about"],
            centre["space_type"],
            centre["brand"],
            centre["price_from"],
            centre["price_unit"],
            centre["seat_type"],
            centre["amenities"],
            centre["transport"],
            centre["map_url"],
            centre["coordinates"],
            centre["source"],
        ))
        return "inserted"


def main():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    total_inserted = 0
    total_updated = 0

    # --- IWG / Regus ---
    print("Loading IWG/Regus data...")
    iwg_centres = load_excel_data(REGUS_FILE, brand_filter=IWG_BRAND_SLUGS)
    print(f"  Found {len(iwg_centres)} London IWG centres after filtering")

    iwg_inserted = iwg_updated = 0
    for centre in iwg_centres:
        result = upsert_centre(cur, centre)
        if result == "inserted":
            iwg_inserted += 1
        else:
            iwg_updated += 1
    conn.commit()
    print(f"  IWG: {iwg_inserted} inserted, {iwg_updated} updated")
    total_inserted += iwg_inserted
    total_updated += iwg_updated

    # --- Fora ---
    print("\nLoading Fora data...")
    fora_inserted = fora_updated = 0
    for filepath in FORA_FILES:
        centres = load_excel_data(filepath, brand_filter=None)
        print(f"  {filepath.split('/')[-1]}: {len(centres)} centres")
        for centre in centres:
            result = upsert_centre(cur, centre)
            if result == "inserted":
                fora_inserted += 1
            else:
                fora_updated += 1
    conn.commit()
    print(f"  Fora: {fora_inserted} inserted, {fora_updated} updated")
    total_inserted += fora_inserted
    total_updated += fora_updated

    conn.close()

    print(f"\n{'='*40}")
    print(f"TOTAL: {total_inserted} inserted, {total_updated} updated")
    print(f"{'='*40}")


if __name__ == "__main__":
    main()
