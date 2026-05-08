#!/usr/bin/env python3
"""
fetch_data.py — Strategic Inputs Policy Monitor
Fetches intervention data from the GTA API v2 and writes data/dashboard.json.
Run daily via GitHub Actions.
"""

import json
import os
import sys
import time
import requests
from datetime import datetime, timedelta, timezone

# ── Config ────────────────────────────────────────────────────────────────────

API_KEY      = os.environ.get("GTA_API_KEY")
DATA_URL     = "https://api.globaltradealert.org/api/v2/gta/data/"
COUNTS_URL   = "https://api.globaltradealert.org/api/v1/gta/data-counts/"
CUTOFF       = "2026-02-28"
REQUEST_DELAY = 0.5

# ── Product definitions ───────────────────────────────────────────────────────

PRODUCTS = [
    {"name": "Fuels",                     "hs_codes": [270900, 271012, 271019],                "description": "Direct exposure to crude supply disruptions, reduced refinery throughput, and constrained petroleum product exports through the Gulf/Hormuz corridor."},
    {"name": "Fertilizers",               "hs_codes": [310210, 310221, 310310, 310520, 310530],"description": "Strong dependence on natural gas-derived ammonia and refinery-derived sulphur/sulphuric acid used in urea, ammonium sulphate, phosphates, DAP/MAP, and NPK production."},
    {"name": "Sulphur",                   "hs_codes": [250300],                                "description": "Sulphur is primarily recovered from oil refining and gas processing; reduced refinery operations directly constrain supply."},
    {"name": "Methanol",                  "hs_codes": [290511],                                "description": "Produced mainly from natural gas feedstocks; Middle East production/export infrastructure highly exposed to Gulf disruptions."},
    {"name": "Graphite Feedstocks",       "hs_codes": [271311, 271312, 380110],                "description": "Petroleum coke-based graphite feedstocks depend on refinery output and delayed coking capacity."},
    {"name": "Alumina",                   "hs_codes": [281820],                                "description": "Energy-intensive refining process indirectly exposed to higher fuel and gas prices caused by regional energy disruptions."},
    {"name": "Helium",                    "hs_codes": [280429],                                "description": "Significant global supply originates from Qatar and Gulf gas-processing infrastructure dependent on Hormuz shipping routes."},
    {"name": "Monoethylene Glycol (MEG)", "hs_codes": [290531],                                "description": "Petrochemical derivative produced from ethylene; exposed to disruptions in naphtha and gas-based cracker feedstocks."},
    {"name": "Iron Ore",                  "hs_codes": [260111, 260112],                        "description": "Primarily indirect exposure via higher bunker/freight costs and weaker steel-sector demand rather than refinery dependence."},
]

# ── Policy group definitions ──────────────────────────────────────────────────

POLICY_GROUPS = {
    "export_controls":    {"label": "Export Controls",    "intervention_types": ["Export ban","Export quota","Export tax","Export tariff quota","Export licensing requirement","Export price benchmark","Local supply requirement"]},
    "import_barriers":    {"label": "Import Barriers",    "intervention_types": ["Import tariff","Import quota","Import ban","Import tariff quota","Import licensing requirement","Import price benchmark","Minimum import price","Other import charges"]},
    "domestic_subsidies": {"label": "Domestic Subsidies", "mast_chapters": ["L"]},
    "export_subsidies":   {"label": "Export Subsidies",   "intervention_types": ["Export subsidy","Trade finance","Financial assistance in a foreign market","Other export incentive"]},
    "sanctions":          {"label": "Sanctions",          "intervention_types": ["Controls on commercial transactions and investment instruments"]},
}

# ── Helpers ───────────────────────────────────────────────────────────────────

def last_saturday(ref):
    days_back = (ref.weekday() - 5) % 7
    return (ref - timedelta(days=days_back)).strftime("%Y-%m-%d")

def get_headers():
    return {
        "Content-Type": "application/json",
        "Authorization": f"APIKey {API_KEY}",
    }

def parse_count(response):
    """Safely extract count from API response regardless of format."""
    data = response.json()
    if isinstance(data, list):
        return len(data)
    if isinstance(data, dict):
        # Try common count field names
        for field in ("count", "total", "n_interventions"):
            if field in data:
                return int(data[field])
        # Fall back to length of results list if present
        if "results" in data:
            return len(data["results"])
    return 0

def data_count(hs_codes, date_end=None, extra=None):
    """Return total intervention count from the v2 data endpoint."""
    request_data = {
        "affected_products": hs_codes,
        "announcement_period": [CUTOFF, date_end or ""],
    }
    if extra:
        request_data.update(extra)

    body = {"limit": 0, "offset": 0, "request_data": request_data}
    time.sleep(REQUEST_DELAY)
    r = requests.post(DATA_URL, headers=get_headers(), json=body, timeout=30)
    r.raise_for_status()
    return parse_count(r)

def counts_by_implementer(hs_codes):
    """Return {jurisdiction_id: count} using the data-counts endpoint."""
    request_data = {
        "affected_products": hs_codes,
        "announcement_period": [CUTOFF, ""],
        "count_by": ["implementer"],
        "count_variable": "intervention_id",
    }
    time.sleep(REQUEST_DELAY)
    r = requests.post(COUNTS_URL, headers=get_headers(), json={"request_data": request_data}, timeout=30)
    r.raise_for_status()
    data = r.json()
    results = data.get("results", data) if isinstance(data, dict) else data

    implementing = {}
    for row in results:
        impl = row.get("implementer", {})
        jid  = impl.get("jurisdiction_id")
        cnt  = row.get("count", 0)
        if jid and cnt:
            implementing[str(jid)] = cnt
    return implementing

# ── Per-product computation ───────────────────────────────────────────────────

def compute_product(product):
    codes = product["hs_codes"]
    today = datetime.now(timezone.utc)
    sat_this = last_saturday(today)
    sat_prev = last_saturday(today - timedelta(days=7))

    print(f"    total ...", end=" ", flush=True)
    total = data_count(codes)
    print(total)

    print(f"    wow ...", end=" ", flush=True)
    n_this = data_count(codes, date_end=sat_this)
    n_prev = data_count(codes, date_end=sat_prev)
    wow = n_this - n_prev
    print(f"{wow:+d}")

    print(f"    harmful ...", end=" ", flush=True)
    harmful = data_count(codes, extra={"gta_evaluation": ["Red", "Amber"]})
    print(harmful)

    print(f"    liberalising ...", end=" ", flush=True)
    liberalising = data_count(codes, extra={"gta_evaluation": ["Green"]})
    print(liberalising)

    policy_groups = {}
    for key, defn in POLICY_GROUPS.items():
        extra = {}
        if "intervention_types" in defn:
            extra["intervention_types"] = defn["intervention_types"]
        if "mast_chapters" in defn:
            extra["mast_chapters"] = defn["mast_chapters"]
        print(f"    {defn['label']} ...", end=" ", flush=True)
        n = data_count(codes, extra=extra)
        policy_groups[key] = n
        print(n)

    print(f"    implementing jurisdictions ...", end=" ", flush=True)
    implementing = counts_by_implementer(codes)
    print(f"{len(implementing)} countries")

    return {
        "name":                product["name"],
        "hs_codes":            codes,
        "description":         product["description"],
        "total_interventions": total,
        "wow_change":          wow,
        "reference_saturdays": {"current": sat_this, "previous": sat_prev},
        "evaluation":          {"harmful": harmful, "liberalising": liberalising},
        "policy_groups":       policy_groups,
        "implementing":        implementing,
    }

# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    if not API_KEY:
        print("ERROR: GTA_API_KEY environment variable is not set.", file=sys.stderr)
        sys.exit(1)

    print(f"=== Strategic Inputs Policy Monitor — Data Fetch ===")
    print(f"Cutoff : {CUTOFF}")
    print(f"Run    : {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')}")
    print()

    output = {
        "last_updated": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
        "cutoff_date":  CUTOFF,
        "overview": (
            "This dashboard tracks trade policy interventions affecting strategic input "
            "commodities with direct or indirect exposure to the 2026 Iran-Hormuz crisis. "
            "All indicators count interventions announced since 28 February 2026, as recorded "
            "in the Global Trade Alert database. Data is refreshed daily."
        ),
        "products": [],
    }

    for i, product in enumerate(PRODUCTS, 1):
        print(f"[{i}/{len(PRODUCTS)}] {product['name']}")
        output["products"].append(compute_product(product))
        print()

    os.makedirs("data", exist_ok=True)
    with open("data/dashboard.json", "w") as f:
        json.dump(output, f, indent=2)
    print("Saved → data/dashboard.json")

if __name__ == "__main__":
    main()
