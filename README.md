# Supply Chain Delivery, Spend & Cost Optimization Analytics (`procurement-spend-optimizer`)

An end-to-end procurement analytics, SQL data modeling, and Excel commercial decision-support portfolio project tailored for industrial connectivity, wiring harness, and box-build manufacturing operations.

---

## 1. Business Problem & Objective
Manufacturing supply chains managing high-mix electromechanical BOMs (Multi-Pin Connectors, Shielded Cables, Wiring Harnesses, Backshells/Seals, and Box Build Enclosures) frequently suffer from margin erosion caused by:
1. **Off-Contract (Maverick) Spot Buying:** Urgent spot orders bypassing Tier-1 rate contracts at 13%–18% price premiums.
2. **Expedited Air Freight Leakage:** Supplier lead-time slippage triggering costly Air Express shipments (14%–25% of material value vs. 3%–5% standard road freight).
3. **Working Capital Trap:** Unbalanced inventory policies resulting in >150 Days of Supply (DOS) on slow-moving SKUs.

This project builds a reproducible **Python + SQLite + Microsoft Excel** pipeline analyzing **2,800 Purchase Orders (₹24.27 Cr landed spend across 18 months)** to pinpoint cost-saving levers, benchmark 10 suppliers, simulate vendor consolidation, and evaluate a 12-item Wiring Harness RFQ BOM.

---

## 2. Architecture & Repository Structure

```text
procurement-spend-optimizer/
├── data/
│   ├── suppliers.csv                  # 10 Tier-1, Tier-2, and Spot vendors with baseline terms
│   ├── sku_master.csv                 # 45 connectivity & box-build SKUs with Should-Cost & inventory data
│   ├── purchase_orders.csv            # 2,800 PO records (Jan 2025 – Jun 2026) in INR
│   ├── rfq_bom_quotes.csv             # 12-item Robotics Wiring Harness BOM quoted across 4 suppliers
│   ├── procurement_analytics.db       # Indexed SQLite relational database & analytical views
│   └── sql_exports/                   # Clean CSV outputs from standalone SQL queries
├── sql/
│   └── queries.sql                    # Production SQL queries (CTEs, Window Functions, Views)
├── excel/
│   └── Supply_Chain_Spend_Cost_Model.xlsx  # 7-tab dynamic Excel workbook with formulas & charts
├── scripts/
│   ├── generate_data.py               # Synthetic procurement & RFQ dataset generator
│   ├── run_sql_analysis.py            # SQLite ETL pipeline & automated query runner
│   └── build_excel_model.py           # Programmatic Excel model & chart builder (OpenPyXL)
├── requirements.txt
└── README.md
