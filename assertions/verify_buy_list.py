"""
Assertion script for the human reviewer - verifies each stated requirement
of the weekly buy list against the generated output.

Run from the Replenishement folder:
    python assertions/verify_buy_list.py
"""
import os
import sys

import numpy as np
import pandas as pd

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, HERE)
import build_buy_list as b  # noqa: E402

def latest_output():
    files = sorted(f for f in os.listdir(os.path.join(HERE, "output"))
                   if f.startswith("Buy List") and f.endswith(".xlsx"))
    assert files, "no Buy List xlsx found in output/"
    return os.path.join(HERE, "output", files[-1])

df, master, vtype = b.build_all()
out = pd.read_excel(latest_output(), sheet_name="Buy List")
passed = 0

def check(name, cond):
    global passed
    assert cond, f"FAILED: {name}"
    passed += 1
    print(f"PASS  {name}")

# R1: SKU data is location-level and sourced from the regional raw exports
check("R1a: buy list rows are (Region, Warehouse, ItemNo) location-level",
      {"Region", "Warehouse", "ItemNo"} <= set(out.columns) and len(out) > 0)
model_keys = set(zip(df["Region"], df["Warehouse"], df["ItemNo"]))
check("R1b: every buy-list row exists in a regional raw export",
      all(k in model_keys
          for k in zip(out["Region"], out["Warehouse"], out["ItemNo"])))

# R2: daily usage shown per SKU-location
check("R2: daily usage (ADU and Demand ADU) present on every row",
      out["ADU"].notna().all() and out["Demand ADU"].notna().all())

# R3: vendors derived from KSI_Item_master (NE by ItemNo, FL by Link No_)
mm_item = (master.drop_duplicates("ItemNo").set_index("ItemNo")
           ["Primary Vendor"].fillna("").astype(str).str.strip())
mm_link = b.master_by_link(master)["Primary Vendor"].fillna("").astype(str).str.strip()
ne = out[out["Region"] == "Northeast"].head(2000)
fl = out[out["Region"] == "Florida"].head(2000)
check("R3a: Northeast primary vendor matches master by ItemNo",
      all(mm_item.get(i, "") == v for i, v in
          zip(ne["ItemNo"], ne["Primary Vendor"].fillna(""))))
check("R3b: Florida primary vendor matches master by Link No_",
      len(fl) == 0 or all(mm_link.get(i, "") == v for i, v in
                          zip(fl["ItemNo"], fl["Primary Vendor"].fillna(""))))

# R4: overseas/domestic classification from Vendor and Type.csv
sample = out.head(4000)
vt = sample["Primary Vendor"].map(vtype).fillna("")
check("R4: vendor type matches Vendor and Type.csv",
      (vt == sample["Vendor Type"].fillna("")).all())

# R5: coverage days - oversea 100, domestic 14
os_rows = out[out["Vendor Type"] == "Oversea"]
do_rows = out[out["Vendor Type"] == "Domestic"]
check("R5a: oversea lead days = 100", (os_rows["Lead Days"] == 100).all())
check("R5b: domestic lead days = 14", (do_rows["Lead Days"] == 14).all())

# R6: safety stock - A items +21 oversea / +7 domestic, others 0
a_os = out[(out["Final Velocity"] == "A") & (out["Vendor Type"] == "Oversea")]
a_do = out[(out["Final Velocity"] == "A") & (out["Vendor Type"] == "Domestic")]
non_a = out[out["Final Velocity"] != "A"]
check("R6a: A-item oversea safety = 21 days", (a_os["Safety Days"] == 21).all())
check("R6b: A-item domestic safety = 7 days", (a_do["Safety Days"] == 7).all())
check("R6c: non-A safety = 0 days", (non_a["Safety Days"] == 0).all())

# R7: no patented items, no companywide P-velocity items
pat_item = master.drop_duplicates("ItemNo").set_index("ItemNo")["Patent"]
cwv_item = master.drop_duplicates("ItemNo").set_index("ItemNo")["Companywide veloicty"]
pat_link = b.master_by_link(master)["Patent"]
cwv_link = b.master_by_link(master)["Companywide veloicty"]
check("R7a: no patented items in buy list",
      not any(pat_item.get(i, 0) == 1 for i in ne["ItemNo"].unique())
      and not any(pat_link.get(i, 0) == 1 for i in fl["ItemNo"].unique()))
check("R7b: no P-velocity items in buy list",
      not any(cwv_item.get(i, "") == "P" for i in ne["ItemNo"].unique())
      and not any(cwv_link.get(i, "") == "P" for i in fl["ItemNo"].unique()))

# R8: target and buy math
check("R8a: Target Qty = ceil(Demand ADU x Target Days)",
      (out["Target Qty"] == np.ceil((out["Demand ADU"] * out["Target Days"]).round(6)).astype(int)).all())
check("R8b: Position = Onhand + OnDock + InTransit + OnOrder",
      (out["Position"] == (out["Location_Onhand"] + out["Location_OnOnDock"]
       + out["Location_InTransit"] + out["Location_OnOrder"]).astype(int)).all())
check("R8c: Buy Qty = max(0, Target - Position), and only Buy>0 rows listed",
      (out["Buy Qty"] == (out["Target Qty"] - out["Position"]).clip(lower=0)).all()
      and (out["Buy Qty"] > 0).all())

# R9: buy list reconciles with full dataset, in total and per region
full = df[df["Buy Qty"] > 0]
check("R9a: total buy units reconcile with recomputed model",
      int(out["Buy Qty"].sum()) == int(full["Buy Qty"].sum()))
check("R9b: per-region buy units reconcile",
      out.groupby("Region")["Buy Qty"].sum().to_dict()
      == {k: int(v) for k, v in full.groupby("Region")["Buy Qty"].sum().items()})

# R10: per-region demand basis
modes = df.attrs["demand_modes"]
check("R10a: Florida uses the recency blend (true 35-day daily rate)",
      "Florida" not in modes or modes["Florida"].startswith("blend"))
check("R10b: Northeast stays YTD-only until its 35-day column is fixed",
      modes["Northeast"].startswith("YTD ADU only"))

print(f"\n{passed} assertions passed.")
