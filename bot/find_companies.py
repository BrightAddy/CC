#!/usr/bin/env python3
"""Collect construction companies for a region from OpenStreetMap (Overpass API).

Companies WITHOUT a website are included on purpose; the `has_website` column
tells you which is which. Output: a CSV file you can open in Excel/Numbers.

Usage examples:
  python3 find_companies.py                       # Rhein-Sieg-Kreis
  python3 find_companies.py --area Bonn
  python3 find_companies.py --area "Köln" --only-no-website
  python3 find_companies.py --area Siegburg --types builder,roofer

Only data that people published on OpenStreetMap is used. No logos or photos
are collected, and nothing is sent to the companies.
"""
import argparse
import csv
import json
import sys
import time
import urllib.parse
import urllib.request

ENDPOINTS = [
    "https://overpass-api.de/api/interpreter",
    "https://overpass.kumi.systems/api/interpreter",
    "https://overpass.private.coffee/api/interpreter",
]

# friendly name -> OSM tag filter
TYPES = {
    "builder": 'craft"="builder',
    "construction_company": 'office"="construction_company',
    "bricklayer": 'craft"="bricklayer',
    "roofer": 'craft"="roofer',
    "carpenter": 'craft"="carpenter',
    "plasterer": 'craft"="plasterer',
    "concrete": 'craft"="concrete',
    "scaffolder": 'craft"="scaffolder',
}
DEFAULT_TYPES = "builder,construction_company,bricklayer"


def geocode(area):
    """Look up the bounding box (south, west, north, east) of a place via Nominatim."""
    import subprocess
    url = ("https://nominatim.openstreetmap.org/search?format=json&limit=1&countrycodes=de&q="
           + urllib.parse.quote(area))
    p = subprocess.run(["curl", "-sS", "-m", "30", "-A", "trhino-company-finder/1.0", url],
                       capture_output=True, timeout=45)
    try:
        s, n, w, e = [float(x) for x in json.loads(p.stdout.decode())[0]["boundingbox"]]
    except Exception:  # noqa: BLE001
        sys.exit("Could not find '%s'. Try another spelling or pass --bbox S,W,N,E." % area)
    return s, w, n, e


def build_query(bbox, types):
    box = "(%s,%s,%s,%s)" % bbox
    parts = []
    for t in types:
        for kind in ("node", "way", "relation"):
            parts.append('%s["%s"]%s;' % (kind, TYPES[t], box))
    return '[out:json][timeout:120];\n(\n%s\n);\nout center tags;' % "\n".join(parts)


def fetch(query, rounds=5):
    """Try every public server; the free ones are often busy, so back off and retry.

    Uses the system `curl` because Python's bundled SSL on older macOS fails the
    TLS handshake with overpass-api.de.
    """
    import subprocess
    last = None
    for rnd in range(rounds):
        for url in ENDPOINTS:
            try:
                p = subprocess.run(
                    ["curl", "-sS", "-m", "110", "-A", "trhino-company-finder/1.0",
                     "--data-urlencode", "data=" + query, url],
                    capture_output=True, timeout=130)
                body = p.stdout.decode("utf-8", "replace")
                if p.returncode != 0 or not body.lstrip().startswith("{"):
                    raise RuntimeError((p.stderr.decode()[:100] or body[:100]).strip().replace("\n", " "))
                res = json.loads(body)
                if not res.get("elements"):
                    # empty answers are usually a busy/half-loaded server; try the next one
                    raise RuntimeError(res.get("remark", "empty result")[:120])
                return res
            except Exception as e:  # noqa: BLE001
                last = e
                print("  %s -> %s" % (url.split("/")[2], e))
        wait = 20 * (rnd + 1)
        print("  all servers busy, waiting %ds ..." % wait)
        time.sleep(wait)
    sys.exit("No usable answer from any Overpass server (%s). The free servers are busy, "
             "or the area name does not match OpenStreetMap (e.g. 'Bonn', 'Köln', "
             "'Rhein-Sieg-Kreis'). Try again in a few minutes." % last)


def pick(tags, *keys):
    for k in keys:
        v = tags.get(k)
        if v:
            return v.strip()
    return ""


def row(el):
    t = el.get("tags", {})
    name = pick(t, "name", "operator", "brand")
    if not name:
        return None
    street = pick(t, "addr:street")
    nr = pick(t, "addr:housenumber")
    website = pick(t, "contact:website", "website", "url")
    if website and not website.lower().startswith("http"):
        website = "https://" + website
    c = el.get("center") or {}
    return {
        "name": name,
        "type": pick(t, "craft", "office"),
        "street": (street + " " + nr).strip(),
        "postcode": pick(t, "addr:postcode"),
        "city": pick(t, "addr:city"),
        "phone": pick(t, "contact:phone", "phone", "contact:mobile"),
        "email": pick(t, "contact:email", "email"),
        "website": website,
        "has_website": "ja" if website else "nein",
        "osm": "https://www.openstreetmap.org/%s/%s" % (el["type"], el["id"]),
        "lat": el.get("lat", c.get("lat", "")),
        "lon": el.get("lon", c.get("lon", "")),
    }


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--area", default="Rhein-Sieg-Kreis",
                    help="administrative area name as in OSM (default: Rhein-Sieg-Kreis)")
    ap.add_argument("--bbox", default=None,
                    help="optional box 'south,west,north,east' instead of looking up --area")
    ap.add_argument("--types", default=DEFAULT_TYPES,
                    help="comma list of: " + ", ".join(TYPES))
    ap.add_argument("--only-no-website", action="store_true",
                    help="keep only companies that have no website")
    ap.add_argument("--out", default=None, help="output CSV (default: out/<area>.csv)")
    a = ap.parse_args()

    types = [x.strip() for x in a.types.split(",") if x.strip()]
    bad = [x for x in types if x not in TYPES]
    if bad:
        sys.exit("Unknown type(s): %s" % ", ".join(bad))

    print("Searching OpenStreetMap for %s in %s ..." % (", ".join(types), a.area))
    if a.bbox:
        bbox = tuple(float(x) for x in a.bbox.split(","))
    else:
        bbox = geocode(a.area)
        time.sleep(1.1)  # Nominatim politeness
    print("Search box (S,W,N,E): %s" % (bbox,))
    res = fetch(build_query(bbox, types))

    seen, rows = set(), []
    for el in res.get("elements", []):
        r = row(el)
        if not r:
            continue
        key = (r["name"].lower(), r["street"].lower(), r["city"].lower())
        if key in seen:
            continue
        seen.add(key)
        if a.only_no_website and r["website"]:
            continue
        rows.append(r)
    rows.sort(key=lambda r: (r["has_website"], r["city"], r["name"].lower()))

    import os
    out = a.out or os.path.join(os.path.dirname(os.path.abspath(__file__)), "out",
                                a.area.replace(" ", "_") + ".csv")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    cols = ["name", "type", "street", "postcode", "city", "phone", "email",
            "website", "has_website", "osm", "lat", "lon"]
    with open(out, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(rows)

    nw = sum(1 for r in rows if not r["website"])
    print("Found %d companies: %d without website, %d with website" %
          (len(rows), nw, len(rows) - nw))
    print("Saved to %s" % out)


if __name__ == "__main__":
    main()
