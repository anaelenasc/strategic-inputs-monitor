#!/usr/bin/env python3
"""
fetch_data.py — Strategic Inputs Policy Monitor
Fetches intervention data from the GTA API and writes data/dashboard.json.
Run daily via GitHub Actions.

NOTE on API parameter names: this script uses the parameter names exposed
by the GTA REST API (api.globaltradealert.org/api/v1/data/). If any request
returns unexpected results, verify the exact field names against your API
documentation, as the beta API may differ from the MCP abstraction layer.
"""

import json
import os
import sys
import time
import requests
from datetime import datetime, timedelta, timezone

# ── Config ────────────────────────────────────────────────────────────────────

API_KEY  = os.environ.get("GTA_API_KEY")
API_URL  = "https://api.globaltradealert.org/api/v1/data/"
CUTOFF   = "2026-02-28"

# Seconds between API requests (be polite to the API)
REQUEST_DELAY = 0.5

# ── Product definitions ───────────────────────────────────────────────────────

PRODUCTS = [
    {
        "name": "Fuels",
        "hs_codes": [270900, 271012, 271019],
        "description": (
            "Direct exposure to crude supply disruptions, reduced refinery throughput, "
            "and constrained petroleum product exports through the Gulf/Hormuz corridor."
        ),
    },
    {
        "name": "Fertilizers",
        "hs_codes": [310210, 310221, 310310, 310520, 310530],
        "description": (
            "Strong dependence on natural gas-derived ammonia and refinery-derived "
            "sulphur/sulphuric acid used in urea, ammonium sulphate, phosphates, "
            "DAP/MAP, and NPK production."
        ),
    },
    {
        "name": "Sulphur",
        "hs_codes": [250300],
        "description": (
            "Sulphur is primarily recovered from oil refining and gas processing; "
            "reduced refinery operations directly constrain supply."
        ),
    },
    {
        "name": "Methanol",
        "hs_codes": [290511],
        "description": (
            "Produced mainly from natural gas feedstocks; Middle East "
            "production/export infrastructure highly exposed to Gulf disruptions."
        ),
    },
    {
        "name": "Graphite Feedstocks",
        "hs_codes": [271311, 271312, 380110],
        "description": (
            "Petroleum coke-based graphite feedstocks depend on refinery output "
            "and delayed coking capacity."
        ),
    },
    {
        "name": "Alumina",
        "hs_codes": [281820],
        "description": (
            "Energy-intensive refining process indirectly exposed to higher fuel "
            "and gas prices caused by regional energy disruptions."
        ),
    },
    {
        "name": "Helium",
        "hs_codes": [280429],
        "description": (
            "Significant global supply originates from Qatar and Gulf "
            "gas-processing infrastructure dependent on Hormuz shipping routes."
        ),
    },
    {
        "name": "Monoethylene Glycol (MEG)",
        "hs_codes": [290531],
        "description": (
            "Petrochemical derivative produced from ethylene; exposed to "
            "disruptions in naphtha and gas-based cracker feedstocks."
        ),
    },
    {
        "name": "Iron Ore",
        "hs_codes": [260111, 260112],
        "description": (
            "Primarily indirect exposure via higher bunker/freight costs and "
            "weaker steel-sector demand rather than refinery dependence."
        ),
    },
]

# ── Policy group definitions ──────────────────────────────────────────────────

POLICY_GROUPS = {
    "export_controls": {
        "label": "Export Controls",
        "intervention_types": [
            "Export ban",
            "Export quota",
            "Export tax",
            "Export tariff quota",
            "Export licensing requirement",
            "Export price benchmark",
            "Local supply requirement",
        ],
    },
    "import_barriers": {
        "label": "Import Barriers",
        "intervention_types": [
            "Import tariff",
            "Import quota",
            "Import ban",
            "Import tariff quota",
            "Import licensing requirement",
            "Import price benchmark",
            "Minimum import price",
            "Other import charges",
        ],
    },
    "domestic_subsidies": {
        "label": "Domestic Subsidies",
        "mast_chapters": ["L"],
    },
    "export_subsidies": {
        "label": "Export Subsidies",
        "intervention_types": [
            "Export subsidy",
            "Trade finance",
            "Financial assistance in a foreign market",
            "Other export incentive",
        ],
    },
    "sanctions": {
        "label": "Sanctions",
        "intervention_types": [
            "Controls on commercial transactions and investment instruments",
        ],
    },
}

# ── Helpers ───────────────────────────────────────────────────────────────────

def last_saturday(reference: datetime) -> datetime:
    """Return the most recent Saturday on or before reference."""
    days_back = (reference.weekday() - 5) % 7
    return reference - timedelta(days=days_back)


def api_count(hs_codes: list, date_lte: str = None, extra: dict = None) -> int:
    """
    POST to GTA /data/ and return the total intervention count.
    Uses limit=1 and reads the 'count' (or 'total') field from the response.
    """
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"APIKey {API_KEY}",
    }

    request_data = {
        "affected_products": hs_codes,
        "date_announced_gte": CUTOFF,
    }
    if date_lte:
        request_data["date_announced_lte"] = date_lte
    if extra:
        request_data.update(extra)

    body = {
        "limit": 1,          # Minimise payload — we only need the count
        "offset": 0,
        "request_data": request_data,
    }

    time.sleep(REQUEST_DELAY)
    resp = requests.post(API_URL, headers=headers, json=body, timeout=30)
    resp.raise_for_status()
    data = resp.json()

    # The GTA API returns total count in a top-level field.
    # Adjust key name if your API version uses a different field.
    return int(data.get("count", data.get("total", data.get("n_interventions", 0))))


# ── Per-product computation ───────────────────────────────────────────────────

def compute_product(product: dict) -> dict:
    codes = product["hs_codes"]
    today = datetime.now(timezone.utc)
    sat_this = last_saturday(today).strftime("%Y-%m-%d")
    sat_prev = (last_saturday(today) - timedelta(days=7)).strftime("%Y-%m-%d")

    print(f"    Total ...", end=" ", flush=True)
    total = api_count(codes)
    print(total)

    print(f"    WoW  ...", end=" ", flush=True)
    count_this_sat = api_count(codes, date_lte=sat_this)
    count_prev_sat = api_count(codes, date_lte=sat_prev)
    wow = count_this_sat - count_prev_sat
    print(f"{wow:+d}")

    print(f"    Evaluation ...", end=" ", flush=True)
    harmful      = api_count(codes, extra={"gta_evaluation": ["Red", "Amber"]})
    liberalising = api_count(codes, extra={"gta_evaluation": ["Green"]})
    murky        = api_count(codes, extra={"gta_evaluation": ["Murky"]})
    print(f"H:{harmful} L:{liberalising} M:{murky}")

    policy_groups = {}
    for key, defn in POLICY_GROUPS.items():
        extra = {}
        if "intervention_types" in defn:
            extra["intervention_types"] = defn["intervention_types"]
        if "mast_chapters" in defn:
            extra["mast_chapters"] = defn["mast_chapters"]
        print(f"    {defn['label']} ...", end=" ", flush=True)
        n = api_count(codes, extra=extra)
        policy_groups[key] = n
        print(n)

    return {
        "name":               product["name"],
        "hs_codes":           codes,
        "description":        product["description"],
        "total_interventions": total,
        "wow_change":         wow,
        "reference_saturdays": {
            "current":  sat_this,
            "previous": sat_prev,
        },
        "evaluation": {
            "harmful":      harmful,
            "liberalising": liberalising,
            "murky":        murky,
        },
        "policy_groups": policy_groups,
    }


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    if not API_KEY:
        print("ERROR: GTA_API_KEY environment variable is not set.", file=sys.stderr)
        sys.exit(1)

    print(f"=== Strategic Inputs Policy Monitor — Data Fetch ===")
    print(f"Cutoff date : {CUTOFF}")
    print(f"Run time    : {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')}")
    print()

    output = {
        "last_updated": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
        "cutoff_date": CUTOFF,
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
        result = compute_product(product)
        output["products"].append(result)
        print()

    os.makedirs("data", exist_ok=True)
    out_path = "data/dashboard.json"
    with open(out_path, "w") as f:
        json.dump(output, f, indent=2)

    print(f"Saved → {out_path}")


if __name__ == "__main__":
    main()
