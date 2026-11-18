import os
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
from openpyxl.chart import BarChart, LineChart, Reference

BASE_DIR = os.path.dirname(os.path.dirname(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
EXPORT_DIR = os.path.join(DATA_DIR, "sql_exports")
EXCEL_DIR = os.path.join(BASE_DIR, "excel")
os.makedirs(EXCEL_DIR, exist_ok=True)
OUT_XLSX = os.path.join(EXCEL_DIR, "Supply_Chain_Spend_Cost_Model.xlsx")

# Styling Palette (Corporate Navy & Slate - TE Connectivity Aesthetic)
NAVY_HEADER = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
TEAL_SUBHEADER = PatternFill(start_color="006677", end_color="006677", fill_type="solid")
LIGHT_BG = PatternFill(start_color="F4F7FA", end_color="F4F7FA", fill_type="solid")
CARD_BG = PatternFill(start_color="EBF1F6", end_color="EBF1F6", fill_type="solid")
ACCENT_ORANGE = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
ACCENT_GREEN = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

FONT_TITLE = Font(name="Calibri", size=15, bold=True, color="1B365D")
FONT_SUBTITLE = Font(name="Calibri", size=10, italic=True, color="595959")
FONT_HEADER = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
FONT_BOLD = Font(name="Calibri", size=10, bold=True, color="1F2937")
FONT_REGULAR = Font(name="Calibri", size=10, color="1F2937")
FONT_KPI_VAL = Font(name="Calibri", size=14, bold=True, color="1B365D")
FONT_KPI_LBL = Font(name="Calibri", size=9, bold=True, color="4B5563")

THIN_BORDER_SIDE = Side(border_style="thin", color="D1D5DB")
BORDER_DATA = Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=THIN_BORDER_SIDE, bottom=THIN_BORDER_SIDE)
BORDER_CARD = Border(
    left=Side(border_style="medium", color="1B365D"),
    right=THIN_BORDER_SIDE,
    top=THIN_BORDER_SIDE,
    bottom=THIN_BORDER_SIDE,
)

FMT_INR = '₹#,##0'
FMT_INR_DEC = '₹#,##0.00'
FMT_INT = '#,##0'
FMT_PCT = '0.0%'
FMT_DEC = '0.00'


def style_range_header(ws, row_idx, max_col, fill=NAVY_HEADER):
    ws.row_dimensions[row_idx].height = 26
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row_idx, column=col)
        cell.font = FONT_HEADER
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER_DATA


def auto_fit_columns(ws, min_width=12, max_width=38):
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            val = str(cell.value or "")
            if val.startswith("="):
                val = "123,456,789"
            max_len = max(max_len, len(val))
        ws.column_dimensions[col_letter].width = max(min_width, min(max_width, max_len + 4))


def build_workbook():
    suppliers_df = pd.read_csv(os.path.join(DATA_DIR, "suppliers.csv"))
    skus_df = pd.read_csv(os.path.join(DATA_DIR, "sku_master.csv"))
    pos_df = pd.read_csv(os.path.join(DATA_DIR, "purchase_orders.csv"))
    rfq_df = pd.read_csv(os.path.join(DATA_DIR, "rfq_bom_quotes.csv"))

    cons_df = pd.read_csv(os.path.join(EXPORT_DIR, "03_maverick_consolidation_savings.csv"))
    inv_df = pd.read_csv(os.path.join(EXPORT_DIR, "04_inventory_turnover_risk.csv"))
    trend_df = pd.read_csv(os.path.join(EXPORT_DIR, "05_monthly_spend_otd_trend.csv"))

    wb = openpyxl.Workbook()
    # Create tabs in executive presentation order
    ws_exec = wb.active
    ws_exec.title = "Executive_Summary"
    ws_sup = wb.create_sheet("Supplier_Scorecard")
    ws_rfq = wb.create_sheet("RFQ_BOM_Evaluation")
    ws_cons = wb.create_sheet("Consolidation_Simulator")
    ws_inv = wb.create_sheet("Inventory_Risk")
    ws_po = wb.create_sheet("PO_Fact_Data")
    ws_mstr = wb.create_sheet("Master_Lookup")

    n_po = len(pos_df) + 1  # Last row in PO_Fact_Data (header is row 1)

    # =========================================================================
    # TAB 6 & 7: PO_Fact_Data & Master_Lookup (Populate first so formulas map cleanly)
    # =========================================================================
    # Enrich PO data with Supplier Name & Tier for easy Excel Pivot Table slicing
    sup_map = suppliers_df.set_index("supplier_id")[["supplier_name", "supplier_tier"]]
    po_enriched = pos_df.merge(sup_map, on="supplier_id", how="left")

    po_cols = list(po_enriched.columns)
    ws_po.append(po_cols)
    style_range_header(ws_po, 1, len(po_cols))
    for r_idx, row in enumerate(po_enriched.itertuples(index=False), start=2):
        ws_po.append(list(row))
        for c_idx in [8, 9, 10, 12, 13, 14]:
            ws_po.cell(row=r_idx, column=c_idx).number_format = FMT_INR_DEC
    ws_po.freeze_panes = "A2"
    ws_po.auto_filter.ref = f"A1:{get_column_letter(len(po_cols))}{n_po}"
    auto_fit_columns(ws_po, min_width=12, max_width=28)

    # Master_Lookup: Suppliers (Cols A-H) and SKUs (Cols J-Q)
    sup_cols = list(suppliers_df.columns)
    for c_idx, h in enumerate(sup_cols, start=1):
        ws_mstr.cell(row=1, column=c_idx, value=h)
    style_range_header(ws_mstr, 1, len(sup_cols), fill=TEAL_SUBHEADER)
    for r_idx, row in enumerate(suppliers_df.itertuples(index=False), start=2):
        for c_idx, val in enumerate(row, start=1):
            cell = ws_mstr.cell(row=r_idx, column=c_idx, value=val)
            cell.font = FONT_REGULAR
            cell.border = BORDER_DATA

    sku_cols = list(skus_df.columns)
    for c_idx, h in enumerate(sku_cols, start=10):
        ws_mstr.cell(row=1, column=c_idx, value=h)
        cell = ws_mstr.cell(row=1, column=c_idx)
        cell.font = FONT_HEADER
        cell.fill = TEAL_SUBHEADER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r_idx, row in enumerate(skus_df.itertuples(index=False), start=2):
        for c_idx, val in enumerate(row, start=10):
            cell = ws_mstr.cell(row=r_idx, column=c_idx, value=val)
            cell.font = FONT_REGULAR
            cell.border = BORDER_DATA
    auto_fit_columns(ws_mstr, min_width=12, max_width=32)

    # =========================================================================
    # TAB 2: Supplier_Scorecard (Live SUMIFS / COUNTIFS / AVERAGEIFS)
    # =========================================================================
    ws_sup["A1"] = "SUPPLIER PERFORMANCE SCORECARD & VARIANCE BENCHMARKING"
    ws_sup["A1"].font = FONT_TITLE
    ws_sup["A2"] = "Dynamic formula-driven scorecard linked to PO_Fact_Data (2,800 POs | Jan 2025 – Jun 2026)"
    ws_sup["A2"].font = FONT_SUBTITLE

    sup_headers = [
        "Supplier ID", "Supplier Name", "Supplier Tier", "Location",
        "Total POs", "Total Units", "Material Spend (INR)", "Freight Cost (INR)",
        "Total Landed Spend (INR)", "Spend Share %", "Net Price Variance (INR)",
        "Weighted Variance %", "On-Time Delivery (OTD %)", "Avg Promised LT (Days)",
        "Avg Actual LT (Days)", "Avg LT Slippage (Days)", "Governance Action"
    ]
    for c_idx, h in enumerate(sup_headers, start=1):
        ws_sup.cell(row=4, column=c_idx, value=h)
    style_range_header(ws_sup, 4, len(sup_headers))

    for idx, sup in suppliers_df.iterrows():
        r = 5 + idx
        ws_sup.cell(row=r, column=1, value=sup["supplier_id"])
        ws_sup.cell(row=r, column=2, value=sup["supplier_name"])
        ws_sup.cell(row=r, column=3, value=sup["supplier_tier"])
        ws_sup.cell(row=r, column=4, value=sup["location"])

        # Live Excel Formulas referencing PO_Fact_Data (F=supplier_id, G=qty, J=mat_spend, L=freight, M=landed, N=var, O=prom_lt, P=act_lt, S=otd_flag)
        ws_sup.cell(row=r, column=5, value=f'=COUNTIF(PO_Fact_Data!$F$2:$F${n_po}, A{r})')
        ws_sup.cell(row=r, column=6, value=f'=SUMIFS(PO_Fact_Data!$G$2:$G${n_po}, PO_Fact_Data!$F$2:$F${n_po}, A{r})')
        ws_sup.cell(row=r, column=7, value=f'=SUMIFS(PO_Fact_Data!$J$2:$J${n_po}, PO_Fact_Data!$F$2:$F${n_po}, A{r})')
        ws_sup.cell(row=r, column=8, value=f'=SUMIFS(PO_Fact_Data!$L$2:$L${n_po}, PO_Fact_Data!$F$2:$F${n_po}, A{r})')
        ws_sup.cell(row=r, column=9, value=f'=SUMIFS(PO_Fact_Data!$M$2:$M${n_po}, PO_Fact_Data!$F$2:$F${n_po}, A{r})')
        ws_sup.cell(row=r, column=10, value=f'=I{r}/$I$15')
        ws_sup.cell(row=r, column=11, value=f'=SUMIFS(PO_Fact_Data!$N$2:$N${n_po}, PO_Fact_Data!$F$2:$F${n_po}, A{r})')
        ws_sup.cell(row=r, column=12, value=f'=K{r}/(G{r}-K{r})')
        ws_sup.cell(row=r, column=13, value=f'=AVERAGEIFS(PO_Fact_Data!$S$2:$S${n_po}, PO_Fact_Data!$F$2:$F${n_po}, A{r})')
        ws_sup.cell(row=r, column=14, value=f'=AVERAGEIFS(PO_Fact_Data!$O$2:$O${n_po}, PO_Fact_Data!$F$2:$F${n_po}, A{r})')
        ws_sup.cell(row=r, column=15, value=f'=AVERAGEIFS(PO_Fact_Data!$P$2:$P${n_po}, PO_Fact_Data!$F$2:$F${n_po}, A{r})')
        ws_sup.cell(row=r, column=16, value=f'=O{r}-N{r}')
        ws_sup.cell(row=r, column=17, value=f'=IF(AND(M{r}>=0.92, L{r}<=0.02), "Preferred Strategic", IF(AND(M{r}>=0.85, L{r}<=0.05), "Approved Core", "Consolidate / Exit"))')

        for c in range(1, 18):
            cell = ws_sup.cell(row=r, column=c)
            cell.font = FONT_REGULAR
            cell.border = BORDER_DATA
            if c in [5, 6]:
                cell.number_format = FMT_INT
            elif c in [7, 8, 9, 11]:
                cell.number_format = FMT_INR
            elif c in [10, 12, 13]:
                cell.number_format = FMT_PCT
            elif c in [14, 15, 16]:
                cell.number_format = FMT_DEC

    # Total Row (Row 15)
    ws_sup.cell(row=15, column=1, value="TOTAL")
    ws_sup.cell(row=15, column=2, value="All Active Suppliers")
    ws_sup.cell(row=15, column=5, value="=SUM(E5:E14)")
    ws_sup.cell(row=15, column=6, value="=SUM(F5:F14)")
    ws_sup.cell(row=15, column=7, value="=SUM(G5:G14)")
    ws_sup.cell(row=15, column=8, value="=SUM(H5:H14)")
    ws_sup.cell(row=15, column=9, value="=SUM(I5:I14)")
    ws_sup.cell(row=15, column=10, value="=SUM(J5:J14)")
    ws_sup.cell(row=15, column=11, value="=SUM(K5:K14)")
    ws_sup.cell(row=15, column=12, value="=K15/(G15-K15)")
    ws_sup.cell(row=15, column=13, value=f"=AVERAGE(PO_Fact_Data!$S$2:$S${n_po})")
    ws_sup.cell(row=15, column=14, value=f"=AVERAGE(PO_Fact_Data!$O$2:$O${n_po})")
    ws_sup.cell(row=15, column=15, value=f"=AVERAGE(PO_Fact_Data!$P$2:$P${n_po})")
    ws_sup.cell(row=15, column=16, value="=O15-N15")
    for c in range(1, 18):
        cell = ws_sup.cell(row=15, column=c)
        cell.font = FONT_BOLD
        cell.fill = CARD_BG
        cell.border = BORDER_DATA
        if c in [5, 6]:
            cell.number_format = FMT_INT
        elif c in [7, 8, 9, 11]:
            cell.number_format = FMT_INR
        elif c in [10, 12, 13]:
            cell.number_format = FMT_PCT
        elif c in [14, 15, 16]:
            cell.number_format = FMT_DEC

    # Conditional Formatting on Price Variance % & OTD %
    ws_sup.conditional_formatting.add(
        "L5:L14",
        ColorScaleRule(start_type="min", start_color="63BE7B", mid_type="percentile", mid_value=50, mid_color="FFEB84", end_type="max", end_color="F8696B"),
    )
    ws_sup.conditional_formatting.add(
        "M5:M14",
        ColorScaleRule(start_type="min", start_color="F8696B", mid_type="percentile", mid_value=50, mid_color="FFEB84", end_type="max", end_color="63BE7B"),
    )
    ws_sup.freeze_panes = "A5"
    auto_fit_columns(ws_sup, min_width=12, max_width=32)

    # =========================================================================
    # TAB 3: RFQ_BOM_Evaluation (Industrial Robotics Wiring Harness RFQ Matrix)
    # =========================================================================
    ws_rfq["A1"] = "COMMERCIAL RFQ EVALUATION: INDUSTRIAL ROBOTICS WIRING HARNESS BOM"
    ws_rfq["A1"].font = FONT_TITLE
    ws_rfq["A2"] = "Annual Program Volume: 2,500 Harness Assemblies | Compares 4 Bidders vs. Engineering Should-Cost Model"
    ws_rfq["A2"].font = FONT_SUBTITLE

    # Pivot RFQ quotes into side-by-side bidder comparison matrix
    bom_base = (
        rfq_df[["bom_line", "sku_id", "component_description", "qty_per_assembly", "annual_assembly_volume", "should_cost_unit_inr"]]
        .drop_duplicates()
        .reset_index(drop=True)
    )
    bidders = ["SUP-001", "SUP-002", "SUP-003", "SUP-007"]
    bidder_names = {
        "SUP-001": "SUP-001 (Precision Interconnect)",
        "SUP-002": "SUP-002 (Deccan Cable)",
        "SUP-003": "SUP-003 (Kaveri Electro-Mech)",
        "SUP-007": "SUP-007 (Shenzhen SinoConnect)",
    }

    rfq_headers = [
        "BOM Line", "SKU ID", "Component Description", "Qty / Assy",
        "Annual Qty", "Should-Cost Unit (INR)", "Annual Should-Cost (INR)",
        bidder_names["SUP-001"], bidder_names["SUP-002"],
        bidder_names["SUP-003"], bidder_names["SUP-007"],
        "Lowest Quote (INR)", "Best Bidder (XLOOKUP)", "Cherry-Pick Annual Spend (INR)"
    ]
    for c_idx, h in enumerate(rfq_headers, start=1):
        ws_rfq.cell(row=4, column=c_idx, value=h)
    style_range_header(ws_rfq, 4, len(rfq_headers))

    for idx, item in bom_base.iterrows():
        r = 5 + idx
        ws_rfq.cell(row=r, column=1, value=item["bom_line"])
        ws_rfq.cell(row=r, column=2, value=item["sku_id"])
        ws_rfq.cell(row=r, column=3, value=item["component_description"])
        ws_rfq.cell(row=r, column=4, value=item["qty_per_assembly"])
        ws_rfq.cell(row=r, column=5, value=f"=D{r}*{int(item['annual_assembly_volume'])}")
        ws_rfq.cell(row=r, column=6, value=item["should_cost_unit_inr"])
        ws_rfq.cell(row=r, column=7, value=f"=E{r}*F{r}")

        for b_i, sup_id in enumerate(bidders):
            q_val = rfq_df[(rfq_df["bom_line"] == item["bom_line"]) & (rfq_df["supplier_id"] == sup_id)]["quoted_unit_price_inr"].values[0]
            ws_rfq.cell(row=r, column=8 + b_i, value=float(q_val))

        ws_rfq.cell(row=r, column=12, value=f"=MIN(H{r}:K{r})")
        ws_rfq.cell(row=r, column=13, value=f"=XLOOKUP(L{r}, H{r}:K{r}, $H$4:$K$4)")
        ws_rfq.cell(row=r, column=14, value=f"=L{r}*E{r}")

        for c in range(1, 15):
            cell = ws_rfq.cell(row=r, column=c)
            cell.font = FONT_REGULAR
            cell.border = BORDER_DATA
            if c in [4, 5]:
                cell.number_format = FMT_INT
            elif c in [6, 8, 9, 10, 11, 12]:
                cell.number_format = FMT_INR_DEC
            elif c in [7, 14]:
                cell.number_format = FMT_INR

    # Commercial Summary Block (Rows 18 to 25)
    summary_rows = [
        ("Annual Material Quote (INR)", "=SUMPRODUCT($E$5:$E$16, H5:H16)", "=SUMPRODUCT($E$5:$E$16, I5:I16)", "=SUMPRODUCT($E$5:$E$16, J5:J16)", "=SUMPRODUCT($E$5:$E$16, K5:K16)", "=SUM(N5:N16)", FMT_INR),
        ("One-Time NRE / Tooling Cost (INR)", 45000, 30000, 65000, 85000, 75000, FMT_INR),
        ("Freight & Customs Duty %", 0.030, 0.035, 0.032, 0.095, 0.033, FMT_PCT),
        ("Annual Logistics & Duty Cost (INR)", "=H18*H20", "=I18*I20", "=J18*J20", "=K18*K20", "=N18*N20", FMT_INR),
        ("TOTAL YEAR-1 LANDED COST (INR)", "=H18+H19+H21", "=I18+I19+I21", "=J18+J19+J21", "=K18+K19+K21", "=N18+N19+N21", FMT_INR),
        ("Variance vs. Baseline Should-Cost + 4% Freight", "=H22-($G$17*1.04)", "=I22-($G$17*1.04)", "=J22-($G$17*1.04)", "=K22-($G$17*1.04)", "=N22-($G$17*1.04)", FMT_INR),
        ("Quoted Lead Time (Days)", 14, 16, 15, 26, 15, FMT_INT),
        ("Technical & PPAP Quality Score (/100)", 94, 91, 95, 86, 94, FMT_INT),
    ]

    # Put Annual Should-Cost Total in G17
    ws_rfq.cell(row=17, column=6, value="Should-Cost Total:").font = FONT_BOLD
    ws_rfq.cell(row=17, column=7, value="=SUM(G5:G16)").number_format = FMT_INR
    ws_rfq.cell(row=17, column=7).font = FONT_BOLD

    for idx, (label, v1, v2, v3, v4, v_split, num_fmt) in enumerate(summary_rows):
        r = 18 + idx
        ws_rfq.cell(row=r, column=3, value=label).font = FONT_BOLD
        vals = {8: v1, 9: v2, 10: v3, 11: v4, 14: v_split}
        for col_i, val in vals.items():
            cell = ws_rfq.cell(row=r, column=col_i, value=val)
            cell.font = FONT_BOLD if "TOTAL" in label else FONT_REGULAR
            cell.number_format = num_fmt
            cell.border = BORDER_DATA
            if "TOTAL" in label:
                cell.fill = ACCENT_GREEN if col_i in [8, 14] else CARD_BG

    # Recommendation Callout Box
    ws_rfq["C27"] = "RECOMMENDED AWARD STRATEGY:"
    ws_rfq["C27"].font = FONT_BOLD
    ws_rfq["D27"] = "Award Single-Source Turnkey Contract to SUP-001 (Precision Interconnect India) or Dual-Source Split (SUP-001 + SUP-003). Despite SUP-007's lower EXW price, 9.5% sea/customs duty, ₹85K tooling, and 26-day lead time erode savings."
    ws_rfq["D27"].font = FONT_REGULAR
    ws_rfq.freeze_panes = "A5"
    auto_fit_columns(ws_rfq, min_width=13, max_width=36)

    # =========================================================================
    # TAB 4: Consolidation_Simulator (Maverick Spend -> Tier-1 Contract)
    # =========================================================================
    ws_cons["A1"] = "VENDOR CONSOLIDATION & MAVERICK SPEND RECOVERY SIMULATOR"
    ws_cons["A1"].font = FONT_TITLE
    ws_cons["A2"] = "Top 15 SKUs with Off-Contract Spot Purchases: Simulates Shifting Volume to Tier-1 Contract Pricing & Road Freight"
    ws_cons["A2"].font = FONT_SUBTITLE

    cons_headers = [
        "SKU ID", "SKU Description", "Category", "Maverick POs", "Maverick Units",
        "Benchmark Unit Price (INR)", "Avg Spot Unit Price (INR)",
        "Material Price Savings (INR)", "Freight Mode Savings (INR)",
        "Total Recoverable Savings (INR)", "Landed Cost Reduction %"
    ]
    for c_idx, h in enumerate(cons_headers, start=1):
        ws_cons.cell(row=4, column=c_idx, value=h)
    style_range_header(ws_cons, 4, len(cons_headers))

    for idx, row in cons_df.iterrows():
        r = 5 + idx
        ws_cons.cell(row=r, column=1, value=row["sku_id"])
        ws_cons.cell(row=r, column=2, value=row["sku_description"])
        ws_cons.cell(row=r, column=3, value=row["category"])
        ws_cons.cell(row=r, column=4, value=int(row["maverick_po_count"]))
        ws_cons.cell(row=r, column=5, value=int(row["maverick_units"]))
        ws_cons.cell(row=r, column=6, value=float(row["benchmark_unit_price_inr"]))
        ws_cons.cell(row=r, column=7, value=float(row["avg_maverick_unit_price_inr"]))
        ws_cons.cell(row=r, column=8, value=f"=(G{r}-F{r})*E{r}")
        ws_cons.cell(row=r, column=9, value=float(row["freight_mode_savings_inr"]))
        ws_cons.cell(row=r, column=10, value=f"=H{r}+I{r}")
        ws_cons.cell(row=r, column=11, value=float(row["landed_cost_reduction_pct"]) / 100.0)

        for c in range(1, 12):
            cell = ws_cons.cell(row=r, column=c)
            cell.font = FONT_REGULAR
            cell.border = BORDER_DATA
            if c in [4, 5]:
                cell.number_format = FMT_INT
            elif c in [6, 7]:
                cell.number_format = FMT_INR_DEC
            elif c in [8, 9, 10]:
                cell.number_format = FMT_INR
            elif c == 11:
                cell.number_format = FMT_PCT

    last_c_row = 4 + len(cons_df)
    tot_c_row = last_c_row + 1
    ws_cons.cell(row=tot_c_row, column=1, value="TOTAL TOP 15")
    ws_cons.cell(row=tot_c_row, column=4, value=f"=SUM(D5:D{last_c_row})")
    ws_cons.cell(row=tot_c_row, column=5, value=f"=SUM(E5:E{last_c_row})")
    ws_cons.cell(row=tot_c_row, column=8, value=f"=SUM(H5:H{last_c_row})")
    ws_cons.cell(row=tot_c_row, column=9, value=f"=SUM(I5:I{last_c_row})")
    ws_cons.cell(row=tot_c_row, column=10, value=f"=SUM(J5:J{last_c_row})")
    ws_cons.cell(row=tot_c_row, column=11, value=f"=AVERAGE(K5:K{last_c_row})")
    for c in range(1, 12):
        cell = ws_cons.cell(row=tot_c_row, column=c)
        cell.font = FONT_BOLD
        cell.fill = ACCENT_GREEN
        cell.border = BORDER_DATA
        if c in [4, 5]:
            cell.number_format = FMT_INT
        elif c in [8, 9, 10]:
            cell.number_format = FMT_INR
        elif c == 11:
            cell.number_format = FMT_PCT

    ws_cons.conditional_formatting.add(
        f"J5:J{last_c_row}",
        DataBarRule(start_type="min", end_type="max", color="2E7D32"),
    )
    ws_cons.freeze_panes = "A5"
    auto_fit_columns(ws_cons, min_width=12, max_width=36)

    # =========================================================================
    # TAB 5: Inventory_Risk (Turnover, Days of Supply & Excess Working Capital)
    # =========================================================================
    ws_inv["A1"] = "INVENTORY TURNOVER, DAYS OF SUPPLY (DOS) & SLOW-MOVING RISK"
    ws_inv["A1"].font = FONT_TITLE
    ws_inv["A2"] = "Flags SKUs exceeding 75-day maximum inventory policy and quantifies trapped working capital (INR)"
    ws_inv["A2"].font = FONT_SUBTITLE

    inv_headers = [
        "SKU ID", "SKU Description", "Category", "Unit Benchmark (INR)",
        "Annual Forecast Qty", "On-Hand Qty", "Safety Stock Qty",
        "Inventory Value (INR)", "Annual Turns", "Days of Supply (DOS)",
        "Excess Units (>75d)", "Excess Working Capital (INR)", "Inventory Health Status"
    ]
    for c_idx, h in enumerate(inv_headers, start=1):
        ws_inv.cell(row=4, column=c_idx, value=h)
    style_range_header(ws_inv, 4, len(inv_headers))

    for idx, row in inv_df.iterrows():
        r = 5 + idx
        ws_inv.cell(row=r, column=1, value=row["sku_id"])
        ws_inv.cell(row=r, column=2, value=row["sku_description"])
        ws_inv.cell(row=r, column=3, value=row["category"])
        ws_inv.cell(row=r, column=4, value=float(row["rfq_benchmark_unit_price_inr"]))
        ws_inv.cell(row=r, column=5, value=int(row["annual_forecast_qty"]))
        ws_inv.cell(row=r, column=6, value=int(row["on_hand_inventory_qty"]))
        ws_inv.cell(row=r, column=7, value=int(row["safety_stock_qty"]))
        ws_inv.cell(row=r, column=8, value=f"=F{r}*D{r}")
        ws_inv.cell(row=r, column=9, value=f"=E{r}/F{r}")
        ws_inv.cell(row=r, column=10, value=f"=F{r}/(E{r}/365)")
        ws_inv.cell(row=r, column=11, value=f"=MAX(0, ROUND(F{r}-((E{r}/365)*75), 0))")
        ws_inv.cell(row=r, column=12, value=f"=K{r}*D{r}")
        ws_inv.cell(row=r, column=13, value=f'=IF(J{r}>150, "Slow-Moving / Excess Risk", IF(F{r}<=G{r}, "Reorder Alert", "Optimal Stock"))')

        for c in range(1, 14):
            cell = ws_inv.cell(row=r, column=c)
            cell.font = FONT_REGULAR
            cell.border = BORDER_DATA
            if c == 4:
                cell.number_format = FMT_INR_DEC
            elif c in [5, 6, 7, 11]:
                cell.number_format = FMT_INT
            elif c in [8, 12]:
                cell.number_format = FMT_INR
            elif c in [9, 10]:
                cell.number_format = FMT_DEC

    last_i_row = 4 + len(inv_df)
    tot_i_row = last_i_row + 1
    ws_inv.cell(row=tot_i_row, column=1, value="TOTAL PORTFOLIO")
    ws_inv.cell(row=tot_i_row, column=8, value=f"=SUM(H5:H{last_i_row})")
    ws_inv.cell(row=tot_i_row, column=9, value=f"=AVERAGE(I5:I{last_i_row})")
    ws_inv.cell(row=tot_i_row, column=10, value=f"=AVERAGE(J5:J{last_i_row})")
    ws_inv.cell(row=tot_i_row, column=11, value=f"=SUM(K5:K{last_i_row})")
    ws_inv.cell(row=tot_i_row, column=12, value=f"=SUM(L5:L{last_i_row})")
    for c in range(1, 14):
        cell = ws_inv.cell(row=tot_i_row, column=c)
        cell.font = FONT_BOLD
        cell.fill = CARD_BG
        cell.border = BORDER_DATA
        if c in [8, 12]:
            cell.number_format = FMT_INR
        elif c == 11:
            cell.number_format = FMT_INT
        elif c in [9, 10]:
            cell.number_format = FMT_DEC

    ws_inv.conditional_formatting.add(
        f"J5:J{last_i_row}",
        ColorScaleRule(start_type="min", start_color="63BE7B", mid_type="num", mid_value=75, mid_color="FFEB84", end_type="max", end_color="F8696B"),
    )
    ws_inv.freeze_panes = "A5"
    auto_fit_columns(ws_inv, min_width=12, max_width=36)

    # =========================================================================
    # TAB 1: Executive_Summary (KPI Cards + Category Table + Monthly Trend + Charts)
    # =========================================================================
    ws_exec["A1"] = "SUPPLY CHAIN DELIVERY, SPEND & COST OPTIMIZATION DASHBOARD"
    ws_exec["A1"].font = FONT_TITLE
    ws_exec["A2"] = "TE Connectivity Portfolio Project | Industrial Connectivity, Harnesses & Box Builds (Jan 2025 – Jun 2026)"
    ws_exec["A2"].font = FONT_SUBTITLE

    # Executive KPI Cards (Rows 4-5)
    kpis = [
        ("A", "B", "TOTAL LANDED SPEND (18M)", "=Supplier_Scorecard!I15", FMT_INR),
        ("C", "D", "SHOULD-COST LEAKAGE (INR)", "=Supplier_Scorecard!K15", FMT_INR),
        ("E", "F", "PORTFOLIO OTD %", "=Supplier_Scorecard!M15", FMT_PCT),
        ("G", "H", "AIR EXPRESS FREIGHT LEAKAGE", f'=SUMIFS(PO_Fact_Data!L2:L{n_po}, PO_Fact_Data!K2:K{n_po}, "Air Express")', FMT_INR),
        ("I", "J", "EXCESS INVENTORY CAPITAL", f"=Inventory_Risk!L{tot_i_row}", FMT_INR),
    ]
    ws_exec.row_dimensions[4].height = 18
    ws_exec.row_dimensions[5].height = 28
    for c1, c2, lbl, formula, fmt in kpis:
        ws_exec.merge_cells(f"{c1}4:{c2}4")
        ws_exec.merge_cells(f"{c1}5:{c2}5")
        lbl_cell = ws_exec[f"{c1}4"]
        val_cell = ws_exec[f"{c1}5"]
        lbl_cell.value = lbl
        lbl_cell.font = FONT_KPI_LBL
        lbl_cell.alignment = Alignment(horizontal="center", vertical="center")
        lbl_cell.fill = CARD_BG

        val_cell.value = formula
        val_cell.font = FONT_KPI_VAL
        val_cell.number_format = fmt
        val_cell.alignment = Alignment(horizontal="center", vertical="center")
        val_cell.fill = CARD_BG

    # Category Spend & Freight Breakdown Table (Rows 8 to 14)
    ws_exec["A7"] = "1. CATEGORY SPEND, FREIGHT BURDEN & OTD PERFORMANCE"
    ws_exec["A7"].font = FONT_BOLD

    cat_headers = [
        "Commodity Category", "PO Count", "Material Spend (INR)",
        "Freight Cost (INR)", "Total Landed Spend (INR)",
        "Spend Share %", "Freight / Material %", "Category OTD %"
    ]
    for c_idx, h in enumerate(cat_headers, start=1):
        ws_exec.cell(row=8, column=c_idx, value=h)
    style_range_header(ws_exec, 8, len(cat_headers))

    cats = [
        "Box Build Enclosures",
        "Wiring Harnesses",
        "Multi-Pin Connectors",
        "Shielded Cables & Wires",
        "Backshells & Seals",
    ]
    for idx, cat in enumerate(cats):
        r = 9 + idx
        ws_exec.cell(row=r, column=1, value=cat)
        ws_exec.cell(row=r, column=2, value=f'=COUNTIF(PO_Fact_Data!$E$2:$E${n_po}, A{r})')
        ws_exec.cell(row=r, column=3, value=f'=SUMIFS(PO_Fact_Data!$J$2:$J${n_po}, PO_Fact_Data!$E$2:$E${n_po}, A{r})')
        ws_exec.cell(row=r, column=4, value=f'=SUMIFS(PO_Fact_Data!$L$2:$L${n_po}, PO_Fact_Data!$E$2:$E${n_po}, A{r})')
        ws_exec.cell(row=r, column=5, value=f'=SUMIFS(PO_Fact_Data!$M$2:$M${n_po}, PO_Fact_Data!$E$2:$E${n_po}, A{r})')
        ws_exec.cell(row=r, column=6, value=f"=E{r}/$E$14")
        ws_exec.cell(row=r, column=7, value=f"=D{r}/C{r}")
        ws_exec.cell(row=r, column=8, value=f'=AVERAGEIFS(PO_Fact_Data!$S$2:$S${n_po}, PO_Fact_Data!$E$2:$E${n_po}, A{r})')

        for c in range(1, 9):
            cell = ws_exec.cell(row=r, column=c)
            cell.font = FONT_REGULAR
            cell.border = BORDER_DATA
            if c == 2:
                cell.number_format = FMT_INT
            elif c in [3, 4, 5]:
                cell.number_format = FMT_INR
            elif c in [6, 7, 8]:
                cell.number_format = FMT_PCT

    # Category Total Row (Row 14)
    ws_exec.cell(row=14, column=1, value="TOTAL")
    ws_exec.cell(row=14, column=2, value="=SUM(B9:B13)")
    ws_exec.cell(row=14, column=3, value="=SUM(C9:C13)")
    ws_exec.cell(row=14, column=4, value="=SUM(D9:D13)")
    ws_exec.cell(row=14, column=5, value="=SUM(E9:E13)")
    ws_exec.cell(row=14, column=6, value="=SUM(F9:F13)")
    ws_exec.cell(row=14, column=7, value="=D14/C14")
    ws_exec.cell(row=14, column=8, value="=AVERAGE(H9:H13)")
    for c in range(1, 9):
        cell = ws_exec.cell(row=14, column=c)
        cell.font = FONT_BOLD
        cell.fill = CARD_BG
        cell.border = BORDER_DATA
        if c == 2:
            cell.number_format = FMT_INT
        elif c in [3, 4, 5]:
            cell.number_format = FMT_INR
        elif c in [6, 7, 8]:
            cell.number_format = FMT_PCT

    # Monthly Spend & OTD Trend Table (Rows 17 to 36)
    ws_exec["A16"] = "2. 18-MONTH SPEND & ON-TIME DELIVERY (OTD) TREND"
    ws_exec["A16"].font = FONT_BOLD

    m_headers = ["Month", "PO Count", "Material Spend (INR)", "Freight Cost (INR)", "Total Landed Spend (INR)", "Price Variance (INR)", "Monthly OTD %"]
    for c_idx, h in enumerate(m_headers, start=1):
        ws_exec.cell(row=17, column=c_idx, value=h)
    style_range_header(ws_exec, 17, len(m_headers), fill=TEAL_SUBHEADER)

    for idx, row in trend_df.iterrows():
        r = 18 + idx
        ws_exec.cell(row=r, column=1, value=row["po_month"])
        ws_exec.cell(row=r, column=2, value=int(row["monthly_pos"]))
        ws_exec.cell(row=r, column=3, value=float(row["material_spend_inr"]))
        ws_exec.cell(row=r, column=4, value=float(row["freight_cost_inr"]))
        ws_exec.cell(row=r, column=5, value=float(row["total_landed_spend_inr"]))
        ws_exec.cell(row=r, column=6, value=float(row["monthly_price_variance_inr"]))
        ws_exec.cell(row=r, column=7, value=float(row["monthly_otd_pct"]) / 100.0)
        for c in range(1, 8):
            cell = ws_exec.cell(row=r, column=c)
            cell.font = FONT_REGULAR
            cell.border = BORDER_DATA
            if c == 2:
                cell.number_format = FMT_INT
            elif c in [3, 4, 5, 6]:
                cell.number_format = FMT_INR
            elif c == 7:
                cell.number_format = FMT_PCT

    # Embedded Native Chart 1: Category Landed Spend (Bar Chart)
    bar_chart = BarChart()
    bar_chart.type = "col"
    bar_chart.style = 10
    bar_chart.title = "Total Landed Spend by Category (INR)"
    bar_chart.y_axis.title = "Landed Spend (INR)"
    bar_chart.x_axis.title = "Commodity Category"
    data_ref = Reference(ws_exec, min_col=5, min_row=8, max_row=13)
    cats_ref = Reference(ws_exec, min_col=1, min_row=9, max_row=13)
    bar_chart.add_data(data_ref, titles_from_data=True)
    bar_chart.set_categories(cats_ref)
    bar_chart.legend = None
    bar_chart.width = 16
    bar_chart.height = 9
    ws_exec.add_chart(bar_chart, "J7")

    # Embedded Native Chart 2: Monthly Landed Spend Trend (Line Chart)
    line_chart = LineChart()
    line_chart.style = 12
    line_chart.title = "Monthly Landed Spend Trajectory (Jan 2025 – Jun 2026)"
    line_chart.y_axis.title = "Landed Spend (INR)"
    line_chart.x_axis.title = "Month"
    m_data_ref = Reference(ws_exec, min_col=5, min_row=17, max_row=17 + len(trend_df))
    m_cats_ref = Reference(ws_exec, min_col=1, min_row=18, max_row=17 + len(trend_df))
    line_chart.add_data(m_data_ref, titles_from_data=True)
    line_chart.set_categories(m_cats_ref)
    line_chart.legend = None
    line_chart.width = 16
    line_chart.height = 9
    ws_exec.add_chart(line_chart, "J23")

    auto_fit_columns(ws_exec, min_width=13, max_width=28)

    wb.save(OUT_XLSX)
    print(f"Successfully generated multi-tab Excel workbook: {OUT_XLSX}")


if __name__ == "__main__":
    build_workbook()