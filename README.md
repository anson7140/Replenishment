# KSI Weekly Replenishment Buy List

Interactive weekly buy list report: **[CAP.html](CAP.html)** (open via GitHub
Pages, or download and open in a browser). `index.html` just redirects there.

## Weekly refresh

1. Drop the new `CAP Raw.xlsx` into this folder (alongside `KSI_Item_master.csv`
   and `Vendor and Type.csv` — these stay local and are never committed).
2. Run:

   ```bash
   python build_buy_list.py
   ```

3. Outputs: `output/Buy List <date>.xlsx` (local only) and `CAP.html`
   (committed here).
4. Verify: `python assertions/verify_buy_list.py` (16 requirement checks).

## Methodology

- **Demand** = YTD ADU x prior-year seasonal index (capped 0.6–1.8).
- **Coverage**: overseas primary vendor 100 days, domestic 14 days;
  A-velocity items add 21 / 7 safety days respectively.
- **Buy qty** = ceil(demand x target days) − (on hand + on dock + in transit + on order).
- Patented and companywide P-velocity items are excluded.

See the Assumptions sheet in the xlsx output for full details and data-quality
warnings.
