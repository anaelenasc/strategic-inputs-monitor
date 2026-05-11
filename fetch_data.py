#!/usr/bin/env python3
"""
fetch_data.py — Strategic Inputs Policy Monitor
Fetches intervention data from the GTA API and writes data/dashboard.json.
Run daily via GitHub Actions.
"""

import json
import os
import sys
import time
import requests
from datetime import datetime, timedelta, timezone

# ── Config ────────────────────────────────────────────────────────────────────

API_KEY    = os.environ.get("GTA_API_KEY")
COUNTS_URL = "https://api.globaltradealert.org/api/v1/gta/data-counts/"
CUTOFF     = "2026-02-28"
END_OPEN   = "2099-12-31"
DELAY      = 0.5

HARMFUL_IDS      = [1, 2]   # Red, Amber
LIBERALISING_IDS = [3, 5]   # Green, Liberalising

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
    "export_controls": {
        "label": "Export Controls",
        "intervention_types": ["Export ban","Export quota","Export tax","Export tariff quota","Export licensing requirement","Export price benchmark","Local supply requirement"],
    },
    "import_barriers": {
        "label": "Import Barriers",
        "intervention_types": ["Import tariff","Import quota","Import ban","Import tariff quota","Import licensing requirement","Import price benchmark","Minimum import price","Other import charges","Internal taxation of imports","Selective import channel restriction"],
    },
    "domestic_subsidies": {
        "label": "Domestic Subsidies",
        "mast_chapters": ["L"],
    },
    "export_subsidies": {
        "label": "Export Subsidies",
        "intervention_types": ["Export subsidy","Trade finance","Financial assistance in foreign market","Other export incentive"],
    },
    "sanctions": {
        "label": "Sanctions",
        "intervention_types": ["Controls on commercial transactions and investment instruments"],
    },
}

# ── Helpers ───────────────────────────────────────────────────────────────────

def last_saturday(ref):
    days_back = (ref.weekday() - 5) % 7
    return (ref - timedelta(days=days_back)).strftime("%Y-%m-%d")

def get_headers():
    return {"Content-Type": "application/json", "Authorization": f"APIKey {API_KEY}"}

def counts_request(hs_codes, date_end=None, count_by=None, extra_filters=None):
    request_data = {
        "affected_products": hs_codes,
        "announcement_period": [CUTOFF, date_end or END_OPEN],
        "count_by": count_by or [],
        "count_variable": "intervention_id",
    }
    if extra_filters:
        request_data.update(extra_filters)
    time.sleep(DELAY)
    r = requests.post(COUNTS_URL, headers=get_headers(), json={"request_data": request_data}, timeout=30)
    r.raise_for_status()
    return r.json()

def total_count(hs_codes, date_end=None, extra_filters=None):
    resp = counts_request(hs_codes, date_end=date_end, extra_filters=extra_filters)
    results = resp.get("results", [])
    if results:
        return sum(row.get("value", 0) for row in results)
    return int(resp.get("count", 0))

# ── Per-product computation ───────────────────────────────────────────────────

def compute_product(product):
    codes = product["hs_codes"]
    today = datetime.now(timezone.utc)
    sat_this = last_saturday(today)
    sat_prev = last_saturday(today - timedelta(days=7))

    # Total + evaluation in one call
    print(f"    total + evaluation ...", end=" ", flush=True)
    resp_eval    = counts_request(codes, count_by=["gta_evaluation"])
    results_eval = resp_eval.get("results", [])
    total        = sum(r.get("value", 0) for r in results_eval)
    harmful      = sum(r["value"] for r in results_eval if r.get("gta_evaluation_id") in HARMFUL_IDS)
    liberalising = sum(r["value"] for r in results_eval if r.get("gta_evaluation_id") in LIBERALISING_IDS)
    print(f"total={total}  harmful={harmful}  liberalising={liberalising}")

    # Week-on-week
    print(f"    wow ...", end=" ", flush=True)
    n_this = total_count(codes, date_end=sat_this)
    n_prev = total_count(codes, date_end=sat_prev)
    wow = n_this - n_prev
    print(f"{wow:+d}")

    # Policy groups via intervention type breakdown
    print(f"    policy groups ...", end=" ", flush=True)
    resp_types = counts_request(codes, count_by=["intervention_type"])
    type_map = {}
    for row in resp_types.get("results", []):
        name = row.get("intervention_type_name", "")
        type_map[name] = type_map.get(name, 0) + row.get("value", 0)

    # Domestic subsidies via MAST chapter breakdown (ID 10 = chapter L)
    resp_mast = counts_request(codes, count_by=["mast_chapter"])
    dom_sub = sum(r.get("value", 0) for r in resp_mast.get("results", []) if r.get("mast_chapter_id") == 10)

    policy_groups = {}
    for key, defn in POLICY_GROUPS.items():
        if key == "domestic_subsidies":
            policy_groups[key] = dom_sub
        else:
            policy_groups[key] = sum(type_map.get(t, 0) for t in defn["intervention_types"])
    print("done")
    for key, defn in POLICY_GROUPS.items():
        print(f"      {defn['label']}: {policy_groups[key]}")

    # Implementing jurisdictions — stored as ISO alpha-3 → count
    # (GTA numeric IDs diverge from ISO 3166-1 for some countries,
    #  so we use implementer_iso which is always reliable)
    print(f"    implementing jurisdictions ...", end=" ", flush=True)
    resp_impl = counts_request(codes, count_by=["implementer"])
    implementing = {}
    for row in resp_impl.get("results", []):
        iso = row.get("implementer_iso")   # e.g. "FRA", "USA"
        cnt = row.get("value", 0)
        if iso and cnt:
            implementing[iso] = cnt
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
        print("ERROR: GTA_API_KEY is not set.", file=sys.stderr)
        sys.exit(1)

    print("=== Strategic Inputs Policy Monitor — Data Fetch ===")
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