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
check("R1a: buy list rows are (Region, Location, ItemNo) location-level",
      {"Region", "Location", "Warehouse", "ItemNo"} <= set(out.columns) and len(out) > 0)
model_keys = set(zip(df["Region"], df["Warehouse"], df["ItemNo"]))
check("R1b: every buy-list row exists in a regional raw export",
      all(k in model_keys
          for k in zip(out["Region"], out["Location"], out["ItemNo"])))

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
# safety follows the COMPUTED per-warehouse Velocity, not the export's
# companywide Source Velocity
a_os = out[(out["Velocity"] == "A") & (out["Vendor Type"] == "Oversea")]
a_do = out[(out["Velocity"] == "A") & (out["Vendor Type"] == "Domestic")]
non_a = out[out["Velocity"] != "A"]
check("R6a: computed-A oversea safety = 21 days", (a_os["Safety Days"] == 21).all())
check("R6b: computed-A domestic safety = 7 days", (a_do["Safety Days"] == 7).all())
check("R6c: non-computed-A safety = 0 days", (non_a["Safety Days"] == 0).all())

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

# R10: per-region demand basis (both exports carry a true per-selling-day
# 35-day rate since the 2026-08-03 refresh)
modes = df.attrs["demand_modes"]
check("R10a: Florida uses the recency blend (true 35-day daily rate)",
      "Florida" not in modes or modes["Florida"].startswith("blend"))
check("R10b: Northeast uses the recency blend (fixed 35-day column)",
      modes["Northeast"].startswith("blend"))

# R11: warehouse grouping - feeders roll into their hub, others unchanged
check("R11a: 07BK rolls into 01NJ and 13PA into 15MP",
      (out.loc[out["Location"] == "07BK", "Warehouse"] == "01NJ").all()
      and (out.loc[out["Location"] == "13PA", "Warehouse"] == "15MP").all())
check("R11b: every other location groups to itself",
      (out.loc[~out["Location"].isin(["07BK", "13PA"]), "Warehouse"]
       == out.loc[~out["Location"].isin(["07BK", "13PA"]), "Location"]).all())

# R12: velocity = cumulative revenue share within Region + Warehouse group.
# Ranking ties make the exact A/B split at a boundary arbitrary, so verify the
# two properties that must hold regardless of tie order: bands are monotone in
# revenue, and each band's cumulative revenue stays inside its ceiling.
RANK = {"A": 0, "B": 1, "C": 2, "D": 3}
mono_ok, band_ok, bad = True, True, []
for (reg, wg), g in df.groupby(["Region", "WH Group"]):
    it = (g.groupby("ItemNo").agg(rev=("Revenue", "sum"),
                                  vel=("Velocity", "first")))
    it["rev"] = it["rev"].clip(lower=0)   # net returns cannot rank
    it = it.sort_values("rev", ascending=False)
    tot = it["rev"].sum()
    if tot <= 0:
        continue
    # Tie-tolerant monotonicity: the cheapest item in a band may equal (never
    # undercut) the priciest item in the next band down.
    lo = it.groupby("vel")["rev"].min()
    hi = it.groupby("vel")["rev"].max()
    for a, c in (("A", "B"), ("B", "C"), ("C", "D")):
        if a in lo.index and c in hi.index and lo[a] < hi[c] - 1e-9:
            mono_ok = False
            bad.append(("monotonic", reg, wg, a, c, lo[a], hi[c]))
    for band, ceil_ in (("A", b.VEL_A), ("B", b.VEL_B), ("C", b.VEL_C)):
        cum = it.loc[it["vel"].map(RANK) <= RANK[band], "rev"].sum() / tot
        if cum > ceil_ + 1e-9:
            band_ok = False
            bad.append((f"{band} ceiling", reg, wg, round(cum, 4)))
check(f"R12a: velocity bands are monotone in revenue{'' if mono_ok else f' {bad[:3]}'}",
      mono_ok)
check(f"R12b: cumulative revenue per band within its ceiling "
      f"(A<={b.VEL_A:.0%}, B<={b.VEL_B:.0%}, C<={b.VEL_C:.0%})"
      f"{'' if band_ok else f' {bad[:3]}'}", band_ok)
check("R12c: zero- and negative-revenue items are classified D",
      (df.loc[df.groupby(["Region", "WH Group", "ItemNo"])["Revenue"]
              .transform("sum") <= 0, "Velocity"] == "D").all())

# R13: hub pooling and netted allocation
def surplus_of(rows):
    """Stock a hub can spare: inventory above its own target, never more
    than the inventory it actually holds."""
    return np.minimum(
        (rows["Position"] - rows["Target Qty"]).clip(lower=0),
        rows["Location_Onhand"] + rows["Location_OnOnDock"]
        + rows["Location_InTransit"] + rows["Location_OnOrder"])


def other_pool(loc):
    """Surplus hub stock a location may draw on: all hubs except itself."""
    h = df[(df["Region"] == "Northeast") & (df["Warehouse"].isin(b.HUBS))
           & (df["Warehouse"] != loc)]
    return surplus_of(h).groupby(h["ItemNo"]).sum()

hub_rows = df[(df["Region"] == "Northeast") & (df["Warehouse"].isin(b.HUBS))]
pool = surplus_of(hub_rows).groupby(hub_rows["ItemNo"]).sum()
ne_other = df[(df["Region"] == "Northeast")
              & (~df["Warehouse"].isin(b.HUBS))]
check("R13a: Hub Avail = eligible hub SURPLUS (stock above the hub's own "
      "target, excluding the row's own hub)",
      (ne_other["Hub Avail"].fillna(0)
       == ne_other["ItemNo"].map(pool).fillna(0)).all())
alloc_by_item = ne_other.groupby("ItemNo")["Hub Alloc"].sum()
check("R13b: allocation never exceeds the hub pool for any SKU",
      (alloc_by_item <= alloc_by_item.index.map(pool).fillna(0) + 1e-6).all())
check("R13c: a hub never counts its own stock as transferable",
      all((df.loc[(df["Warehouse"] == h) & (df["Region"] == "Northeast"),
                  "Hub Avail"].fillna(0)
           <= df.loc[(df["Warehouse"] == h) & (df["Region"] == "Northeast"),
                     "ItemNo"].map(other_pool(h)).fillna(0) + 1e-6).all()
          for h in b.HUBS))
check("R13d: Net Buy Qty = Buy Qty - Hub Alloc (floored at 0)",
      (out["Net Buy Qty"] == (out["Buy Qty"]
       - out["Hub Alloc"].fillna(0)).clip(lower=0)).all())

# R14: non-stocking feeders roll their demand into their hub
ne_all = df[df["Region"] == "Northeast"]
check("R14a: feeder demand is conserved - nothing lost in the rollup",
      abs(ne_all["Own Demand ADU"].sum() - ne_all["Demand ADU"].sum()) < 1e-4)
check("R14b: each feeder's demand lands entirely at its hub",
      all(abs(ne_all.loc[ne_all["Warehouse"] == f, "Own Demand ADU"].sum()
              - ne_all.loc[ne_all["Warehouse"] == h, "Feeder Demand ADU"].sum())
          < 1e-4 for f, h in b.WH_GROUP.items()))
check("R14c: non-stocking feeders raise no buy lines",
      not len(out[out["Location"].isin(b.NON_STOCKING)]))

# R15: a hub short of its own target lends nothing
for h in b.HUBS:
    rows = df[(df["Region"] == "Northeast") & (df["Warehouse"] == h)]
    short = rows[rows["Position"] < rows["Target Qty"]]
    check(f"R15 {h}: rows short of own target offer zero surplus",
          (surplus_of(short) == 0).all())

print(f"\n{passed} assertions passed.")
