#!/usr/bin/env python3
"""
fetch_data.py — Strategic Inputs Policy Monitor
Fetches intervention data from the GTA API and writes:
  - data/dashboard.json      (counts for the dashboard)
  - data/interventions.xlsx  (full intervention list for download)
Run daily via GitHub Actions.
"""

import json
import os
import sys
import time
import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime, timedelta, timezone

# ── Config ────────────────────────────────────────────────────────────────────

API_KEY    = os.environ.get("GTA_API_KEY")
COUNTS_URL = "https://api.globaltradealert.org/api/v1/gta/data-counts/"
DATA_URL   = "https://api.globaltradealert.org/api/v2/gta/data/"
CUTOFF     = "2026-02-28"
END_OPEN   = "2099-12-31"
DELAY      = 0.5

HARMFUL_IDS      = [1, 2]
LIBERALISING_IDS = [3, 5]

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

POLICY_GROUPS = {
    "export_controls":    {"label": "Export Controls",    "intervention_types": ["Export ban","Export quota","Export tax","Export tariff quota","Export licensing requirement","Export price benchmark","Local supply requirement"]},
    "import_barriers":    {"label": "Import Barriers",    "intervention_types": ["Import tariff","Import quota","Import ban","Import tariff quota","Import licensing requirement","Import price benchmark","Minimum import price","Other import charges","Internal taxation of imports","Selective import channel restriction"]},
    "domestic_subsidies": {"label": "Domestic Subsidies", "mast_chapters": ["L"]},
    "export_subsidies":   {"label": "Export Subsidies",   "intervention_types": ["Export subsidy","Trade finance","Financial assistance in foreign market","Other export incentive"]},
    "sanctions":          {"label": "Sanctions",          "intervention_types": ["Controls on commercial transactions and investment instruments"]},
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

def fetch_all_interventions(hs_codes):
    """Fetch every intervention record for given HS codes, paginating as needed."""
    all_results = []
    offset = 0
    limit  = 1000

    while True:
        body = {
            "limit":  limit,
            "offset": offset,
            "request_data": {
                "affected_products":   hs_codes,
                "announcement_period": [CUTOFF, END_OPEN],
            },
        }
        time.sleep(DELAY)
        r = requests.post(DATA_URL, headers=get_headers(), json=body, timeout=60)
        r.raise_for_status()
        data    = r.json()
        results = data.get("results", [])
        all_results.extend(results)

        total  = data.get("count", 0)
        offset += limit
        if offset >= total or not results:
            break

    return all_results

# ── Field extraction helpers ──────────────────────────────────────────────────

def extract_str(val):
    if val is None:
        return ""
    if isinstance(val, str):
        return val
    if isinstance(val, dict):
        return val.get("name", val.get("label", str(val)))
    if isinstance(val, list):
        return "; ".join(extract_str(v) for v in val)
    return str(val)

def extract_row(intv):
    date        = intv.get("date_announced", "")
    title       = intv.get("state_act_title") or intv.get("title") or intv.get("name") or ""
    evaluation  = extract_str(intv.get("gta_evaluation", ""))
    implementer = extract_str(intv.get("implementing_jurisdiction", ""))
    url         = intv.get("intervention_url") or intv.get("url") or ""

    types_raw = intv.get("intervention_types", [])
    if isinstance(types_raw, list):
        int_types = "; ".join(extract_str(t) for t in types_raw)
    else:
        int_types = extract_str(types_raw)

    products_raw = intv.get("affected_products", [])
    if isinstance(products_raw, list):
        hs_parts = []
        for p in products_raw:
            if isinstance(p, dict):
                pid  = p.get("product_id", "")
                name = p.get("name", "")
                hs_parts.append(f"{pid} ({name})" if name else str(pid))
            else:
                hs_parts.append(str(p))
        hs_codes = "; ".join(hs_parts)
    else:
        hs_codes = extract_str(products_raw)

    affected_raw = intv.get("affected_jurisdictions", [])
    if isinstance(affected_raw, list):
        aff_juris = "; ".join(extract_str(j) for j in affected_raw)
    else:
        aff_juris = extract_str(affected_raw)

    return [date, title, evaluation, implementer, int_types, hs_codes, aff_juris, url]

XLSX_HEADERS = [
    "Announcement Date",
    "Title",
    "GTA Evaluation",
    "Implementing Jurisdiction",
    "Intervention Type(s)",
    "Affected HS Codes",
    "Affected Jurisdictions",
    "GTA URL",
]

NAVY  = "0E334D"
WHITE = "FFFFFF"
LIGHT = "E8EEF5"

def style_header_row(ws):
    for cell in ws[1]:
        cell.font      = Font(bold=True, color=WHITE, name="Calibri", size=10)
        cell.fill      = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)

def set_col_widths(ws):
    widths = [14, 60, 16, 28, 40, 40, 50, 55]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

def add_sheet(wb, sheet_name, rows, product_label=None):
    ws = wb.create_sheet(sheet_name[:31])
    ws.freeze_panes = "A2"

    if product_label:
        ws.append([f"Product category: {product_label}"])
        ws["A1"].font = Font(bold=True, color=NAVY, name="Calibri", size=10)
        ws.append(XLSX_HEADERS)
        style_header_row(ws)
        data_start = 3
    else:
        ws.append(XLSX_HEADERS)
        style_header_row(ws)
        data_start = 2

    for i, row in enumerate(rows, data_start):
        ws.append(row)
        if (i - data_start) % 2 == 1:
            for cell in ws[i]:
                cell.fill = PatternFill("solid", fgColor=LIGHT)
        url_cell = ws.cell(row=i, column=8)
        if url_cell.value:
            url_cell.hyperlink = url_cell.value
            url_cell.font      = Font(color="1874CD", underline="single", name="Calibri", size=10)

    set_col_widths(ws)
    ws.row_dimensions[1 if not product_label else 2].height = 18
    return ws

# ── Per-product dashboard counts ─────────────────────────────────────────────

def compute_product_counts(product):
    codes = product["hs_codes"]
    today = datetime.now(timezone.utc)
    sat_this = last_saturday(today)
    sat_prev = last_saturday(today - timedelta(days=7))

    print(f"    total + evaluation ...", end=" ", flush=True)
    resp_eval    = counts_request(codes, count_by=["gta_evaluation"])
    results_eval = resp_eval.get("results", [])
    total        = sum(r.get("value", 0) for r in results_eval)
    harmful      = sum(r["value"] for r in results_eval if r.get("gta_evaluation_id") in HARMFUL_IDS)
    liberalising = sum(r["value"] for r in results_eval if r.get("gta_evaluation_id") in LIBERALISING_IDS)
    print(f"total={total}  harmful={harmful}  liberalising={liberalising}")

    print(f"    wow ...", end=" ", flush=True)
    n_this = total_count(codes, date_end=sat_this)
    n_prev = total_count(codes, date_end=sat_prev)
    wow = n_this - n_prev
    print(f"{wow:+d}")

    print(f"    policy groups ...", end=" ", flush=True)
    resp_types = counts_request(codes, count_by=["intervention_type"])
    type_map = {}
    for row in resp_types.get("results", []):
        name = row.get("intervention_type_name", "")
        type_map[name] = type_map.get(name, 0) + row.get("value", 0)

    resp_mast = counts_request(codes, count_by=["mast_chapter"])
    dom_sub = sum(r.get("value", 0) for r in resp_mast.get("results", []) if r.get("mast_chapter_id") == 10)

    policy_groups = {}
    covered_types = set()
    for key, defn in POLICY_GROUPS.items():
        if key == "domestic_subsidies":
            policy_groups[key] = dom_sub
        else:
            policy_groups[key] = sum(type_map.get(t, 0) for t in defn["intervention_types"])
            covered_types.update(defn["intervention_types"])
    policy_groups["other"] = sum(v for k, v in type_map.items() if k not in covered_types)
    print("done")

    print(f"    implementing jurisdictions ...", end=" ", flush=True)
    resp_impl = counts_request(codes, count_by=["implementer"])
    implementing = {}
    for row in resp_impl.get("results", []):
        iso = row.get("implementer_iso")
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

# ── Excel generation ──────────────────────────────────────────────────────────

def build_excel(interventions_by_product):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    seen     = set()
    all_rows = []
    for product_name, records in interventions_by_product.items():
        for rec in records:
            iid = rec.get("intervention_id")
            if iid not in seen:
                seen.add(iid)
                all_rows.append(extract_row(rec))

    add_sheet(wb, "All Products", all_rows)
    print(f"  All Products sheet: {len(all_rows)} unique interventions")

    for product_name, records in interventions_by_product.items():
        rows = [extract_row(r) for r in records]
        add_sheet(wb, product_name[:31], rows, product_label=product_name)
        print(f"  {product_name} sheet: {len(rows)} interventions")

    os.makedirs("data", exist_ok=True)
    wb.save("data/interventions.xlsx")
    print(f"Saved → data/interventions.xlsx")

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

    interventions_by_product = {}

    for i, product in enumerate(PRODUCTS, 1):
        print(f"[{i}/{len(PRODUCTS)}] {product['name']}")
        output["products"].append(compute_product_counts(product))
        print()

    # Write dashboard JSON first — so it is always saved even if Excel fails
    os.makedirs("data", exist_ok=True)
    with open("data/dashboard.json", "w") as f:
        json.dump(output, f, indent=2)
    print("Saved → data/dashboard.json")

    # Fetch full records for Excel
    print("\nFetching full intervention records for Excel...")
    for i, product in enumerate(PRODUCTS, 1):
        print(f"[{i}/{len(PRODUCTS)}] {product['name']} ...", end=" ", flush=True)
        records = fetch_all_interventions(product["hs_codes"])
        interventions_by_product[product["name"]] = records
        print(f"{len(records)} records")

    print("\nBuilding Excel file...")
    build_excel(interventions_by_product)

if __name__ == "__main__":
    main()
