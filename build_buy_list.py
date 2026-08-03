"""
KSI Weekly Buy List Builder
===========================
Rerun weekly after dropping the new "CAP Raw.xlsx" into this folder:

    python build_buy_list.py

Inputs (this folder):
  - CAP Raw.xlsx        location-level YTD sales/inventory export
  - KSI_Item_master.csv item master (patent flag, velocity, vendors)
  - Vendor and Type.csv vendor -> Domestic / Oversea

Outputs (output/ subfolder):
  - Buy List <date>.xlsx   (Buy List, Vendor Summary, Warehouse Summary,
                            Exceptions, Assumptions)
  - CAP.html               self-contained interactive report (GitHub-ready)
  - index.html             redirect to CAP.html for the Pages homepage

Methodology
-----------
Demand (units/selling-day) = YTD ADU x seasonal index.
  - YTD ADU = Volume / 149 selling days (Jan 1 - Jul 31, 2026), as provided.
  - Seasonal index = (LY Aug-Dec daily rate) / (LY Jan-Jul daily rate),
    computed from PY Volume (same period LY) and Volume_Prior Year (full LY),
    capped to [0.6, 1.8]; only applied when full-LY volume >= 6 units.
  - Recency: if "Rolling Avg 35 ADU" is a true daily rate (auto-detected),
    the base becomes 70% recent / 30% YTD. The current export's column is
    average units per day-with-a-sale (75% of values are exactly 1.0), not
    units/day, so it is auto-rejected and YTD ADU is used alone.

Coverage target (selling days):
  - Primary vendor Oversea:  100 days  (+21 safety days if velocity A)
  - Primary vendor Domestic:  14 days  (+7 safety days if velocity A)

Target inventory = ceil(demand x target days)
Position         = Onhand + OnDock + InTransit + OnOrder
Buy qty          = max(0, target - position)

Exclusions: patented items (Patent=1) and P-velocity items from the master.
Rows with no primary vendor default to the Domestic target and are flagged.
"""

import math
import os
import sys
import warnings
from datetime import date

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(HERE, "output")

RAW_XLSX = os.path.join(HERE, "CAP Raw.xlsx")
MASTER_CSV = os.path.join(HERE, "KSI_Item_master.csv")
VENDOR_CSV = os.path.join(HERE, "Vendor and Type.csv")

WAREHOUSES = ["01NJ", "03LF", "05RO", "07BK", "09SJ", "11MH", "13PA", "15MP"]

SELLING_DAYS_YTD = 149          # Jan 1 - Jul 31 2026 selling days (Volume/ADU in export)
CAL_DAYS_SAME = 212             # calendar days Jan 1 - Jul 31
CAL_DAYS_REMAIN = 153           # calendar days Aug 1 - Dec 31

OVERSEA_DAYS = 100
DOMESTIC_DAYS = 14
SAFETY_A_OVERSEA = 21           # +3 weeks for A items, overseas primary
SAFETY_A_DOMESTIC = 7           # +1 week for A items, domestic primary

SEASONAL_CAP = (0.6, 1.8)
SEASONAL_MIN_LY_UNITS = 6       # need >=6 units full LY to trust a seasonal index

# Recency blend, used only when the export's 35-day column is a true daily
# rate (units sold last 35 days / 35). The current export's "Rolling Avg 35
# ADU" is avg units per day-with-a-sale (median 1.0), which fails detection.
RECENT_WEIGHT = 0.7             # 70% last-35-day ADU, 30% YTD ADU
RECENT_VALID_MEDIAN = 0.5       # median of positive values must be below this


def load_inputs():
    raw = pd.read_excel(RAW_XLSX)
    raw = raw[raw["Warehouse"].isin(WAREHOUSES)].copy()
    raw["ItemNo"] = raw["ItemNo"].astype(str).str.strip()

    master = pd.read_csv(MASTER_CSV, encoding="utf-8-sig")
    master["ItemNo"] = master["ItemNo"].astype(str).str.strip()
    master = master.drop_duplicates("ItemNo")

    vend = pd.read_csv(VENDOR_CSV, encoding="utf-8-sig")
    vend["V1"] = vend["V1"].astype(str).str.strip()
    vtype = dict(zip(vend["V1"], vend["V1Type"]))
    return raw, master, vtype


def build(raw, master, vtype):
    df = raw.merge(
        master[["ItemNo", "Patent", "Companywide veloicty",
                "Primary Vendor", "Secondary Vendor"]],
        on="ItemNo", how="left",
    )

    # --- exclusions ---------------------------------------------------------
    df = df[(df["Patent"].fillna(0) != 1)]
    df = df[df["Companywide veloicty"].fillna("") != "P"]

    # --- vendor type --------------------------------------------------------
    df["Primary Vendor"] = df["Primary Vendor"].fillna("").astype(str).str.strip()
    df["Secondary Vendor"] = df["Secondary Vendor"].fillna("").astype(str).str.strip()
    df["Vendor Type"] = df["Primary Vendor"].map(vtype).fillna("")
    df["Vendor Missing"] = df["Primary Vendor"] == ""

    # --- demand -------------------------------------------------------------
    for c in ["Volume", "PY Volume", "Volume_Prior Year", "ADU", "PY ADU",
              "Location_Onhand", "Location_OnOnDock", "Location_InTransit",
              "Location_OnOrder", "Rolling Avg 35 ADU", "Revenue"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    df["ADU"] = df["ADU"].clip(lower=0)

    # Recency detection: a real 35-day ADU has mostly-fractional values like
    # YTD ADU. The broken per-active-day column has a positive-value median
    # of 1.0 and gets rejected here.
    r35 = df["Rolling Avg 35 ADU"].clip(lower=0)
    pos = r35[r35 > 0]
    recent_valid = len(pos) > 0 and pos.median() < RECENT_VALID_MEDIAN
    if recent_valid:
        base_adu = RECENT_WEIGHT * r35 + (1 - RECENT_WEIGHT) * df["ADU"]
        demand_mode = (f"blend of {RECENT_WEIGHT:.0%} last-35-day ADU + "
                       f"{1 - RECENT_WEIGHT:.0%} YTD ADU")
    else:
        base_adu = df["ADU"]
        demand_mode = ("YTD ADU only ('Rolling Avg 35 ADU' column is not a "
                       "daily rate in this export)")
    df.attrs["demand_mode"] = demand_mode

    ly_same = df["PY Volume"].clip(lower=0)
    ly_full = df["Volume_Prior Year"].clip(lower=0)
    ly_remain = (ly_full - ly_same).clip(lower=0)

    rate_same = ly_same / CAL_DAYS_SAME
    rate_remain = ly_remain / CAL_DAYS_REMAIN
    idx = np.where(
        (ly_full >= SEASONAL_MIN_LY_UNITS) & (rate_same > 0),
        (rate_remain / rate_same.replace(0, np.nan)).fillna(1.0),
        1.0,
    )
    df["Seasonal Index"] = np.clip(idx, *SEASONAL_CAP).round(3)
    df["Demand ADU"] = (base_adu * df["Seasonal Index"]).round(5)

    # --- coverage target ----------------------------------------------------
    oversea = df["Vendor Type"] == "Oversea"
    is_a = df["Final Velocity"].fillna("") == "A"

    df["Lead Days"] = np.where(oversea, OVERSEA_DAYS, DOMESTIC_DAYS)
    df["Safety Days"] = np.where(
        is_a, np.where(oversea, SAFETY_A_OVERSEA, SAFETY_A_DOMESTIC), 0
    )
    df["Target Days"] = df["Lead Days"] + df["Safety Days"]

    df["Target Qty"] = np.ceil(df["Demand ADU"] * df["Target Days"]).astype(int)
    df["Position"] = (
        df["Location_Onhand"] + df["Location_OnOnDock"]
        + df["Location_InTransit"] + df["Location_OnOrder"]
    ).astype(int)
    df["Buy Qty"] = (df["Target Qty"] - df["Position"]).clip(lower=0).astype(int)
    return df


def data_quality_warnings(df):
    warns = []
    for wh, g in df.groupby("Warehouse"):
        if g["Volume"].sum() == 0 and g["Position"].sum() == 0:
            warns.append(
                f"Warehouse {wh}: all volume and inventory are zero in this export "
                f"({len(g):,} rows) - likely truncated by the source system "
                f"('Exported data exceeded the allowed volume'). No buys generated for {wh}; re-export needed."
            )
    n_noveil = int((df["Vendor Missing"] & (df["Buy Qty"] > 0)).sum())
    if n_noveil:
        warns.append(
            f"{n_noveil} buy lines have no primary vendor in KSI_Item_master - "
            f"defaulted to the Domestic 14-day target; see Exceptions sheet."
        )
    return warns


def write_excel(df, buys, warns, out_path):
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    cols = ["Warehouse", "ItemNo", "CAP_ItemNum", "Product Desc", "Model",
            "Final Velocity", "Primary Vendor", "Vendor Type", "Secondary Vendor",
            "ADU", "Seasonal Index", "Demand ADU", "Lead Days", "Safety Days",
            "Target Days", "Target Qty", "Location_Onhand", "Location_OnOnDock",
            "Location_InTransit", "Location_OnOrder", "Position", "Buy Qty",
            "Volume", "PY Volume", "Volume_Prior Year", "Rolling Avg 35 ADU"]

    vend_sum = (buys.groupby(["Primary Vendor", "Vendor Type"], dropna=False)
                .agg(SKU_Locations=("ItemNo", "size"),
                     Unique_SKUs=("ItemNo", "nunique"),
                     Buy_Units=("Buy Qty", "sum"))
                .reset_index().sort_values("Buy_Units", ascending=False))
    wh_sum = (buys.groupby("Warehouse")
              .agg(SKU_Locations=("ItemNo", "size"),
                   Buy_Units=("Buy Qty", "sum"))
              .reset_index().sort_values("Buy_Units", ascending=False))
    exceptions = df[df["Vendor Missing"] & (df["Buy Qty"] > 0)][cols]

    assumptions = pd.DataFrame({"Assumption": [
        f"Demand basis: {df.attrs.get('demand_mode', 'YTD ADU')}. YTD ADU = Volume / 149 selling days (Jan 1 - Jul 31 2026) as provided in CAP Raw.",
        "Seasonal index = (LY Aug-Dec daily rate) / (LY Jan-Jul daily rate) from PY Volume and Volume_Prior Year; capped 0.6-1.8; applied only when full-LY volume >= 6 units.",
        "Recency weighting activates automatically once 'Rolling Avg 35 ADU' is exported as a true daily rate (total units last 35 days / 35); the current column (avg units per day-with-sales) is auto-rejected.",
        "Coverage: Oversea primary = 100 days; Domestic = 14 days. A items add 21 safety days (oversea) / 7 (domestic). Days = selling days, same basis as ADU.",
        "Target Qty = ceil(Demand ADU x Target Days). Buy Qty = max(0, Target - (Onhand + OnDock + InTransit + OnOrder)).",
        "Excluded: patented items (Patent=1) and companywide P-velocity items per KSI_Item_master.",
        "Rows with no primary vendor use the Domestic 14-day target and appear on the Exceptions sheet - assign vendors to fix.",
        "Source export warning: 'Exported data exceeded the allowed volume. Some data may have been omitted.' appears in CAP Raw - some rows may be missing from the source.",
        "Negative ADU (net returns) treated as zero demand.",
    ] + ["DATA WARNING: " + w for w in warns]})

    with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
        buys[cols].to_excel(xw, sheet_name="Buy List", index=False)
        vend_sum.to_excel(xw, sheet_name="Vendor Summary", index=False)
        wh_sum.to_excel(xw, sheet_name="Warehouse Summary", index=False)
        exceptions.to_excel(xw, sheet_name="Exceptions", index=False)
        assumptions.to_excel(xw, sheet_name="Assumptions", index=False)

        wb = xw.book
        hdr_fill = PatternFill("solid", fgColor="1F4E78")
        for ws in wb.worksheets:
            for cell in ws[1]:
                cell.font = Font(name="Arial", bold=True, color="FFFFFF")
                cell.fill = hdr_fill
            ws.freeze_panes = "A2"
            for i, col in enumerate(ws.iter_cols(min_row=1, max_row=1), 1):
                width = min(max(len(str(col[0].value or "")) + 2, 10), 55)
                ws.column_dimensions[get_column_letter(i)].width = width
            ws.auto_filter.ref = ws.dimensions
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    cell.font = Font(name="Arial")


def write_html(buys, warns, run_date, out_path):
    import json

    vend_sum = (buys.groupby(["Primary Vendor", "Vendor Type"], dropna=False)
                ["Buy Qty"].sum().reset_index()
                .sort_values("Buy Qty", ascending=False))
    vdata = [[r["Primary Vendor"] or "(none)", r["Vendor Type"] or "",
              int(r["Buy Qty"])] for _, r in vend_sum.iterrows()]

    wh_type = (buys.assign(t=buys["Vendor Type"].where(
                   buys["Vendor Type"].isin(["Oversea", "Domestic"]), "Unassigned"))
               .pivot_table(index="Warehouse", columns="t", values="Buy Qty",
                            aggfunc="sum", fill_value=0))
    wdata = sorted(
        ([wh, int(row.get("Oversea", 0)), int(row.get("Domestic", 0)),
          int(row.get("Unassigned", 0))] for wh, row in wh_type.iterrows()),
        key=lambda r: -(r[1] + r[2] + r[3]))

    rows = buys[["Warehouse", "ItemNo", "CAP_ItemNum", "Product Desc",
                 "Final Velocity", "Primary Vendor", "Vendor Type",
                 "Demand ADU", "Target Days", "Target Qty", "Position",
                 "Buy Qty"]].values.tolist()
    payload = json.dumps(rows, default=str)

    stats = {
        "skuLoc": f"{len(buys):,}",
        "units": f"{int(buys['Buy Qty'].sum()):,}",
        "oversea": f"{int(buys.loc[buys['Vendor Type']=='Oversea','Buy Qty'].sum()):,}",
        "domestic": f"{int(buys.loc[buys['Vendor Type']=='Domestic','Buy Qty'].sum()):,}",
    }
    warn_html = "".join(f'<div class="warn">&#9888;&#65039; {w}</div>' for w in warns)

    html = HTML_TEMPLATE
    for k, v in {"__DATE__": run_date, "__WARNS__": warn_html,
                 "__SKULOC__": stats["skuLoc"],
                 "__UNITS__": stats["units"], "__OVERSEA__": stats["oversea"],
                 "__DOMESTIC__": stats["domestic"],
                 "__VDATA__": json.dumps(vdata), "__WDATA__": json.dumps(wdata),
                 "__DATA__": payload}.items():
        html = html.replace(k, v)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>KSI Buy List - __DATE__</title>
<style>
:root{--bg:#f6f7f9;--card:#fff;--ink:#1a2733;--mut:#5c6b7a;--line:#e3e8ee;--acc:#1f4e78;--s1:#2a78d6;--s2:#eb6834;--sx:#898781}
@media (prefers-color-scheme: dark){:root{--bg:#12181f;--card:#1a232d;--ink:#e8edf2;--mut:#8fa0b0;--line:#2a3642;--acc:#6aa5d8;--s1:#3987e5;--s2:#d95926}}
*{box-sizing:border-box}body{margin:0;font:14px/1.5 -apple-system,Segoe UI,Arial,sans-serif;background:var(--bg);color:var(--ink);padding:24px}
h1{font-size:20px;margin:0 0 4px}.sub{color:var(--mut);margin-bottom:20px}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px;margin-bottom:20px}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:14px 16px}
.card .v{font-size:22px;font-weight:700}.card .l{color:var(--mut);font-size:12px}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px}
@media(max-width:800px){.grid2{grid-template-columns:1fr}}
.panel{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:14px 16px;overflow-x:auto}
.panel h2{font-size:14px;margin:0 0 10px}
table{border-collapse:collapse;width:100%}th,td{padding:6px 10px;border-bottom:1px solid var(--line);text-align:left;white-space:nowrap}
th{color:var(--mut);font-size:12px;cursor:pointer;user-select:none}
.num{text-align:right;font-variant-numeric:tabular-nums}
.controls{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:10px}
input,select{padding:7px 10px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--ink)}
.pager{display:flex;gap:8px;align-items:center;margin-top:10px;color:var(--mut)}
button{padding:6px 12px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--ink);cursor:pointer}
.badge{display:inline-block;padding:1px 8px;border-radius:10px;font-size:11px;background:var(--acc);color:#fff}
.warn{background:#fff3cd;border:1px solid #ffe08a;color:#664d03;border-radius:8px;padding:10px 14px;margin-bottom:12px}
@media (prefers-color-scheme: dark){.warn{background:#3a3020;border-color:#6b5a2a;color:#ffd97a}}
.ctrlbar{display:flex;gap:16px;align-items:center;margin:0 0 12px;flex-wrap:wrap}
.lg{display:inline-flex;align-items:center;gap:6px;color:var(--mut);font-size:12px}
.sw{width:10px;height:10px;border-radius:3px;display:inline-block}
.seg{margin-left:auto;display:inline-flex;border:1px solid var(--line);border-radius:8px;overflow:hidden}
.seg button{border:0;border-radius:0;padding:6px 12px;font-size:12px}
.seg button.on{background:var(--acc);color:#fff}
.crow{display:grid;grid-template-columns:118px 1fr 76px;align-items:center;gap:8px;min-height:24px}
.crow .nm{font-size:12px;color:var(--ink);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.crow .val{font-size:12px;color:var(--ink);text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
.ghead{font-size:11px;letter-spacing:.06em;text-transform:uppercase;color:var(--mut);margin:12px 0 4px;display:flex;justify-content:space-between}
.ghead:first-child{margin-top:2px}
.track{display:flex;gap:2px;height:16px}
.bseg{height:16px;min-width:1px}
.bseg.end{border-radius:0 4px 4px 0}
#tip{position:fixed;z-index:10;background:var(--card);border:1px solid var(--line);border-radius:8px;padding:8px 10px;font-size:12px;line-height:1.45;pointer-events:none;display:none;box-shadow:0 4px 14px rgba(0,0,0,.18)}
</style></head><body>
<h1>KSI Replenishment Buy List</h1>
<div class="sub">Generated __DATE__ &middot; demand = YTD ADU &times; seasonal index &middot; oversea 100d / domestic 14d (+A-item safety)</div>
__WARNS__
<div class="cards">
<div class="card"><div class="v">__SKULOC__</div><div class="l">SKU-locations to buy</div></div>
<div class="card"><div class="v">__UNITS__</div><div class="l">Total buy units</div></div>
<div class="card"><div class="v">__OVERSEA__</div><div class="l">Oversea vendor units</div></div>
<div class="card"><div class="v">__DOMESTIC__</div><div class="l">Domestic vendor units</div></div>
</div>
<div class="ctrlbar">
<span class="lg"><i class="sw" style="background:var(--s1)"></i>Oversea</span>
<span class="lg"><i class="sw" style="background:var(--s2)"></i>Domestic</span>
<span class="lg" id="lgUn" hidden><i class="sw" style="background:var(--sx)"></i>Unassigned vendor</span>
<span class="seg"><button id="mUnits" class="on">Units</button><button id="mCtn">Containers (1,300/ctn)</button></span>
</div>
<div class="grid2">
<div class="panel"><h2 id="vh">Top vendors</h2><div id="vchart"></div></div>
<div class="panel"><h2 id="whh">By warehouse</h2><div id="wchart"></div></div>
</div>
<div id="tip"></div>
<div class="panel">
<h2>Buy list detail <span class="badge" id="count"></span></h2>
<div class="controls">
<input id="q" placeholder="Search SKU / CAP # / description / vendor" size="36">
<select id="fwh"><option value="">All warehouses</option></select>
<select id="fvt"><option value="">All vendor types</option><option>Oversea</option><option>Domestic</option></select>
<select id="fvel"><option value="">All velocities</option><option>A</option><option>B</option><option>C</option><option>D</option></select>
<button id="csv" title="Downloads the rows matching the current filters">&#11015; Export CSV (filtered)</button>
</div>
<table id="tbl"><thead><tr>
<th data-k="0">WH</th><th data-k="1">SKU</th><th data-k="2">CAP Part #</th><th data-k="3">Description</th><th data-k="4">Vel</th>
<th data-k="5">Vendor</th><th data-k="6">Type</th><th data-k="7" class="num">Demand/day</th>
<th data-k="8" class="num">Target days</th><th data-k="9" class="num">Target</th>
<th data-k="10" class="num">Position</th><th data-k="11" class="num">Buy</th>
</tr></thead><tbody></tbody></table>
<div class="pager"><button id="prev">&laquo; Prev</button><span id="pinfo"></span><button id="next">Next &raquo;</button></div>
</div>
<script>
const DATA=__DATA__;let view=DATA.slice(),page=0,PS=100,sortK=11,sortD=-1;
const $=id=>document.getElementById(id);
[...new Set(DATA.map(r=>r[0]))].sort().forEach(w=>{const o=document.createElement('option');o.textContent=w;$('fwh').appendChild(o)});
function apply(){const q=$('q').value.toLowerCase(),wh=$('fwh').value,vt=$('fvt').value,vl=$('fvel').value;
view=DATA.filter(r=>(!wh||r[0]===wh)&&(!vt||r[6]===vt)&&(!vl||r[4]===vl)&&(!q||(r[1]+' '+r[2]+' '+r[3]+' '+r[5]).toLowerCase().includes(q)));
view.sort((a,b)=>{const x=a[sortK],y=b[sortK];return(typeof x==='number'?x-y:String(x).localeCompare(String(y)))*sortD});
page=0;render()}
function render(){const tb=$('tbl').querySelector('tbody');tb.innerHTML='';
view.slice(page*PS,(page+1)*PS).forEach(r=>{const tr=document.createElement('tr');
tr.innerHTML=`<td>${r[0]}</td><td>${r[1]}</td><td>${r[2]||''}</td><td>${r[3]}</td><td>${r[4]||''}</td><td>${r[5]||'(none)'}</td><td>${r[6]||'-'}</td><td class="num">${(+r[7]).toFixed(3)}</td><td class="num">${r[8]}</td><td class="num">${r[9]}</td><td class="num">${r[10]}</td><td class="num"><b>${r[11]}</b></td>`;
tb.appendChild(tr)});
$('count').textContent=view.length.toLocaleString()+' rows';
$('pinfo').textContent=`page ${page+1} / ${Math.max(1,Math.ceil(view.length/PS))}`}
['q','fwh','fvt','fvel'].forEach(id=>$(id).addEventListener('input',apply));
$('prev').onclick=()=>{if(page>0){page--;render()}};
$('next').onclick=()=>{if((page+1)*PS<view.length){page++;render()}};
document.querySelectorAll('th[data-k]').forEach(th=>th.onclick=()=>{const k=+th.dataset.k;sortD=(k===sortK)?-sortD:-1;sortK=k;apply()});
$('csv').onclick=()=>{
const hdr=['Warehouse','ItemNo','CAP_ItemNum','Description','Velocity','Primary Vendor','Vendor Type','Demand ADU','Target Days','Target Qty','Position','Buy Qty'];
const esc=v=>{v=(v==null?'':String(v));return /[",\n]/.test(v)?'"'+v.replace(/"/g,'""')+'"':v};
const csv=[hdr.join(',')].concat(view.map(r=>r.map(esc).join(','))).join('\r\n');
const a=document.createElement('a');
a.href=URL.createObjectURL(new Blob(['\uFEFF'+csv],{type:'text/csv;charset=utf-8'}));
a.download='buy_list_'+document.title.split(' - ')[1]+'_'+(view.length===DATA.length?'all':'filtered')+'.csv';
a.click();URL.revokeObjectURL(a.href)};
apply();

// ---- charts: top vendors & by warehouse, units <-> containers toggle ----
const VDATA=__VDATA__,WDATA=__WDATA__,CTN=1300;let mode='units';
const COLORS={Oversea:'var(--s1)',Domestic:'var(--s2)',Unassigned:'var(--sx)'};
const fmtC=v=>{const c=v/CTN;return c>0&&c<0.05?'<0.1':c.toLocaleString(undefined,{minimumFractionDigits:1,maximumFractionDigits:1})};
const fmt=v=>mode==='units'?v.toLocaleString():fmtC(v);
const both=v=>v.toLocaleString()+' units · '+fmtC(v)+' ctn';
const tip=$('tip');
function tipMove(e){tip.style.left=Math.min(e.clientX+14,innerWidth-tip.offsetWidth-8)+'px';tip.style.top=Math.min(e.clientY+14,innerHeight-tip.offsetHeight-8)+'px'}
function hoverize(el,html){el.addEventListener('mousemove',e=>{tip.innerHTML=html;tip.style.display='block';tipMove(e)});el.addEventListener('mouseleave',()=>tip.style.display='none')}
function renderVend(){
const el=$('vchart');el.innerHTML='';
const max=Math.max(...VDATA.map(r=>r[2]));
for(const g of ['Oversea','Domestic','Unassigned']){
const rows=VDATA.filter(r=>(r[1]||'Unassigned')===g);
if(!rows.length)continue;
const tot=rows.reduce((s,r)=>s+r[2],0);
const top=rows.slice(0,8),rest=rows.slice(8),restSum=rest.reduce((s,r)=>s+r[2],0);
const h=document.createElement('div');h.className='ghead';h.innerHTML=`<span>${g}</span><span>${fmt(tot)}</span>`;el.appendChild(h);
const items=top.map(r=>[r[0],r[2],false]);if(restSum)items.push([`Other (${rest.length} vendors)`,restSum,true]);
for(const [nm,v] of items){
const row=document.createElement('div');row.className='crow';
row.innerHTML=`<span class="nm" title="${nm}">${nm}</span><span class="track"><span class="bseg end" style="width:${Math.max(100*v/max,.4).toFixed(2)}%;background:${COLORS[g]}"></span></span><span class="val">${fmt(v)}</span>`;
hoverize(row.querySelector('.bseg'),`<b>${nm}</b><br>${g}<br>${both(v)}`);
el.appendChild(row);}}}
function renderWh(){
const el=$('wchart');el.innerHTML='';
const max=Math.max(...WDATA.map(r=>r[1]+r[2]+r[3]));
for(const [wh,os,dom,un] of WDATA){
const tot=os+dom+un;
const segs=[['Oversea',os],['Domestic',dom],['Unassigned',un]].filter(s=>s[1]>0);
const row=document.createElement('div');row.className='crow';
row.innerHTML=`<span class="nm">${wh}</span><span class="track">${segs.map((s,i)=>`<span class="bseg${i===segs.length-1?' end':''}" style="width:${(100*s[1]/max).toFixed(2)}%;background:${COLORS[s[0]]}"></span>`).join('')}</span><span class="val">${fmt(tot)}</span>`;
row.querySelectorAll('.bseg').forEach((sg,i)=>hoverize(sg,`<b>${wh}</b><br>${segs[i][0]}: ${both(segs[i][1])}<br>Total: ${both(tot)}`));
el.appendChild(row);}}
function renderCharts(){renderVend();renderWh();
$('vh').textContent='Top vendors — '+(mode==='units'?'buy units':'containers (1,300 units each)');
$('whh').textContent=(mode==='units'?'Buy units':'Containers')+' by warehouse';
$('mUnits').classList.toggle('on',mode==='units');$('mCtn').classList.toggle('on',mode!=='units')}
$('mUnits').onclick=()=>{mode='units';renderCharts()};
$('mCtn').onclick=()=>{mode='ctn';renderCharts()};
if(VDATA.some(r=>!r[1]))$('lgUn').hidden=false;
renderCharts();
</script></body></html>
"""


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    run_date = date.today().isoformat()

    raw, master, vtype = load_inputs()
    df = build(raw, master, vtype)
    buys = df[df["Buy Qty"] > 0].sort_values("Buy Qty", ascending=False)

    print(f"demand mode: {df.attrs['demand_mode']}")
    warns = data_quality_warnings(df)
    for w in warns:
        print(f"WARNING: {w}")

    xlsx_path = os.path.join(OUT_DIR, f"Buy List {run_date}.xlsx")
    write_excel(df, buys, warns, xlsx_path)
    write_html(buys, warns, run_date, os.path.join(OUT_DIR, "CAP.html"))
    # copy to repo root so GitHub Pages can serve it; index.html redirects
    # to CAP.html so the Pages homepage URL keeps working
    import shutil
    shutil.copyfile(os.path.join(OUT_DIR, "CAP.html"), os.path.join(HERE, "CAP.html"))
    with open(os.path.join(HERE, "index.html"), "w", encoding="utf-8") as f:
        f.write('<!DOCTYPE html><html><head><meta charset="utf-8">'
                '<meta http-equiv="refresh" content="0; url=CAP.html">'
                '<title>KSI Buy List</title></head>'
                '<body><a href="CAP.html">Open the CAP buy list</a></body></html>')

    print(f"rows after filters: {len(df):,}")
    print(f"buy lines: {len(buys):,}  buy units: {int(buys['Buy Qty'].sum()):,}")
    print(f"  oversea units: {int(buys.loc[buys['Vendor Type']=='Oversea','Buy Qty'].sum()):,}")
    print(f"  domestic units: {int(buys.loc[buys['Vendor Type']=='Domestic','Buy Qty'].sum()):,}")
    print(f"  missing-vendor lines: {int((buys['Vendor Missing']).sum()):,}")
    print(f"wrote: {xlsx_path}")
    print(f"wrote: {os.path.join(OUT_DIR, 'CAP.html')}")


if __name__ == "__main__":
    main()
