import os
import random
from datetime import datetime, timedelta
import numpy as np
import pandas as pd

# Set reproducible seed for consistent portfolio numbers
np.random.seed(42)
random.seed(42)

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data")
os.makedirs(DATA_DIR, exist_ok=True)


def generate_suppliers() -> pd.DataFrame:
    suppliers = [
        ("SUP-001", "Precision Interconnect India Pvt Ltd", "Pune, MH", "Tier-1 Contract", "Net 60", 0.96, 1.00, "Road - FTL"),
        ("SUP-002", "Deccan Cable & Harness Systems", "Bengaluru, KA", "Tier-1 Contract", "Net 45", 0.94, 1.01, "Road - LTL"),
        ("SUP-003", "Kaveri Electro-Mech Components", "Chennai, TN", "Tier-1 Contract", "Net 60", 0.95, 0.99, "Road - FTL"),
        ("SUP-004", "Apex Industrial Enclosures", "Ahmedabad, GJ", "Tier-2 Approved", "Net 45", 0.88, 1.04, "Road - LTL"),
        ("SUP-005", "Narmada Polymer & Seal Tech", "Vadodara, GJ", "Tier-2 Approved", "Net 30", 0.90, 1.03, "Road - LTL"),
        ("SUP-006", "Vindhya Wire & Extrusion Corp", "Bhiwadi, RJ", "Tier-2 Approved", "Net 45", 0.85, 1.05, "Road - LTL"),
        ("SUP-007", "Shenzhen SinoConnect Electronics", "Shenzhen, CN", "Global Tier-1", "Net 60", 0.87, 0.94, "Sea + Customs"),
        ("SUP-008", "Metro Rapid Spot Supplies", "Mumbai, MH", "Spot / Maverick", "Net 15", 0.74, 1.18, "Air Express"),
        ("SUP-009", "NCR QuickTurn Harness Works", "Gurugram, HR", "Spot / Maverick", "Net 15", 0.78, 1.15, "Air Express"),
        ("SUP-010", "Southern Fast-Track Connectors", "Coimbatore, TN", "Spot / Maverick", "Net 30", 0.80, 1.13, "Air Express"),
    ]
    cols = [
        "supplier_id", "supplier_name", "location", "supplier_tier",
        "payment_terms", "base_otd_prob", "price_index", "default_freight_mode"
    ]
    return pd.DataFrame(suppliers, columns=cols)


def generate_sku_master() -> pd.DataFrame:
    categories = {
        "Multi-Pin Connectors": {
            "prefix": "CON",
            "items": [
                ("Heavy-Duty 24-Pin Rectangular Housing", 185.0, 14),
                ("M12 8-Pin Circular Industrial Plug", 142.0, 12),
                ("Deutsch DT 12-Pin Sealed Receptacle", 168.0, 14),
                ("Board-to-Wire 16-Pin Gold Contact Header", 96.0, 10),
                ("High-Voltage 4-Pin EV Interlock Connector", 420.0, 18),
                ("RF Coaxial SMA Bulkhead Jack", 115.0, 12),
                ("Modular DIN 41612 64-Pin Backplane Conn", 260.0, 15),
                ("Micro-Fit 10-Pin Dual-Row Receptacle", 64.0, 10),
                ("IP68 Waterproof 6-Pin Sensor Connector", 132.0, 12),
                ("Fiber Optic LC Duplex Industrial Adapter", 210.0, 16),
            ],
        },
        "Shielded Cables & Wires": {
            "prefix": "CBL",
            "items": [
                ("18 AWG Braided Shielded Twisted Pair (per m)", 48.0, 10),
                ("14 AWG Cross-Linked XLPE Automotive Wire (per m)", 36.0, 8),
                ("22 AWG 4-Core Foil-Shielded Instrumentation Cable (per m)", 62.0, 12),
                ("10 AWG High-Temp Silicone Power Cable (per m)", 95.0, 12),
                ("CAT6A Industrial Drag-Chain Ethernet Cable (per m)", 78.0, 14),
                ("16 AWG Tinned Copper Grounding Braid (per m)", 54.0, 10),
                ("CAN-Bus J1939 Twinax Shielded Cable (per m)", 68.0, 12),
                ("20 AWG PTFE Fluoropolymer Aerospace Wire (per m)", 112.0, 15),
                ("Multi-Core 12x0.75mm Control Flex Cable (per m)", 135.0, 14),
            ],
        },
        "Wiring Harnesses": {
            "prefix": "HRN",
            "items": [
                ("Main Chassis Power Distribution Harness", 2450.0, 21),
                ("CNC Servo Motor Feedback & Encoder Harness", 1380.0, 18),
                ("Cabinet Door I/O Sensor Loom Assembly", 890.0, 15),
                ("HVAC Compressor Control Sub-Harness", 640.0, 14),
                ("Battery Management System (BMS) Sense Harness", 1720.0, 20),
                ("Dashboard HMI Display Ribbon & Power Assembly", 760.0, 15),
                ("Engine Bay Overmolded Sensor Wiring Loom", 1950.0, 22),
                ("Telematics & GPS Antenna Jumper Harness", 520.0, 14),
                ("Robotics 6-Axis Articulated Arm Flex Harness", 3150.0, 25),
            ],
        },
        "Backshells & Seals": {
            "prefix": "BSL",
            "items": [
                ("EMI/RFI 90-Degree Aluminum Die-Cast Backshell", 155.0, 12),
                ("Heat-Shrinkable Polyolefin Transition Boot", 44.0, 8),
                ("Silicone Multi-Wire Grommet Seal (IP67)", 22.0, 7),
                ("Strain Relief Cable Gland M25 Brass Nickel", 58.0, 9),
                ("Conductive Elastomer Flange Gasket", 39.0, 9),
                ("Convoluted Split Loom Tubing PA6 (per m)", 28.0, 7),
                ("Cavity Sealing Plug Silicone (Pack of 10)", 18.0, 7),
                ("braided Nomex Abrasion Sleeve (per m)", 65.0, 10),
                ("EMI Shielding Copper Tape Roll 25mm", 190.0, 10),
            ],
        },
        "Box Build Enclosures": {
            "prefix": "BOX",
            "items": [
                ("IP66 Powder-Coated Sheet Metal Junction Box", 1850.0, 20),
                ("Die-Cast Aluminum ECU Controller Housing", 1240.0, 18),
                ("DIN-Rail Modular Relay & Terminal Sub-Assembly", 2680.0, 22),
                ("Polycarbonate Clear-Lid Telemetry Enclosure", 980.0, 15),
                ("Stainless Steel 316L Washdown Control Cabinet", 4450.0, 28),
                ("Custom Machined Heat-Sink Power Module Chassis", 1620.0, 20),
                ("Pre-Wired Industrial PDU 8-Channel Box Build", 3890.0, 25),
                ("IP68 Potted Sensor Junction Module", 790.0, 16),
            ],
        },
    }

    records = []
    sku_counter = 101
    for cat_name, cat_info in categories.items():
        for desc, bench_price, std_lt in cat_info["items"]:
            sku_id = f"{cat_info['prefix']}-{sku_counter}"
            sku_counter += 1

            # Simulate annual demand & inventory levels (with ~18% slow-moving/excess stock SKUs)
            annual_demand = int(np.random.randint(1200, 18000))
            if bench_price > 1500:
                annual_demand = int(np.random.randint(400, 3500))

            is_slow_moving = random.random() < 0.18
            if is_slow_moving:
                # Overstocked: 180 to 320 days of inventory on hand
                days_cover = np.random.randint(180, 320)
            else:
                # Healthy stock: 22 to 65 days of inventory on hand
                days_cover = np.random.randint(22, 65)

            on_hand_qty = int(round((annual_demand / 365.0) * days_cover))
            safety_stock = int(round((annual_demand / 365.0) * 15))

            records.append({
                "sku_id": sku_id,
                "sku_description": desc,
                "category": cat_name,
                "rfq_benchmark_unit_price_inr": bench_price,
                "standard_lead_time_days": std_lt,
                "annual_forecast_qty": annual_demand,
                "on_hand_inventory_qty": on_hand_qty,
                "safety_stock_qty": safety_stock,
            })

    return pd.DataFrame(records)


def generate_purchase_orders(suppliers_df: pd.DataFrame, skus_df: pd.DataFrame, n_rows: int = 2800) -> pd.DataFrame:
    start_date = datetime(2025, 1, 1)
    end_date = datetime(2026, 6, 30)
    total_days = (end_date - start_date).days

    # Map category to preferred Tier-1/Tier-2 suppliers + occasional Spot suppliers
    tier12_sups = suppliers_df[suppliers_df["supplier_tier"] != "Spot / Maverick"]["supplier_id"].tolist()
    spot_sups = suppliers_df[suppliers_df["supplier_tier"] == "Spot / Maverick"]["supplier_id"].tolist()

    sup_lookup = suppliers_df.set_index("supplier_id").to_dict("index")
    sku_records = skus_df.to_dict("records")

    po_rows = []
    for i in range(1, n_rows + 1):
        po_id = f"PO-2025-{10000 + i}"
        sku = random.choice(sku_records)

        # 82% contracted/approved suppliers, 18% maverick/spot purchases
        if random.random() < 0.82:
            sup_id = random.choice(tier12_sups)
        else:
            sup_id = random.choice(spot_sups)

        sup = sup_lookup[sup_id]
        po_date = start_date + timedelta(days=random.randint(0, total_days))

        # Order quantity scaled inversely to unit price
        bench_price = sku["rfq_benchmark_unit_price_inr"]
        if bench_price < 100:
            qty = int(np.random.choice([250, 500, 1000, 1500, 2000]))
        elif bench_price < 500:
            qty = int(np.random.choice([100, 200, 300, 500]))
        else:
            qty = int(np.random.choice([20, 40, 50, 80, 120]))

        # Actual unit price with realistic market noise + supplier price index
        noise = np.random.normal(loc=0.0, scale=0.025)
        # Volume discount for large orders from Tier-1 suppliers
        vol_disc = -0.025 if (qty >= 1000 and sup["supplier_tier"] == "Tier-1 Contract") else 0.0
        actual_unit_price = round(bench_price * (sup["price_index"] + noise + vol_disc), 2)

        # Lead times & On-Time Delivery
        promised_lt = sku["standard_lead_time_days"] + (7 if sup_id == "SUP-007" else 0)
        is_on_time = random.random() < sup["base_otd_prob"]

        if is_on_time:
            delay_days = int(np.random.choice([-2, -1, 0, 0, 0]))
        else:
            delay_days = int(np.random.randint(2, 14))

        actual_lt = max(3, promised_lt + delay_days)
        promised_delivery_date = po_date + timedelta(days=promised_lt)
        actual_delivery_date = po_date + timedelta(days=actual_lt)
        otd_flag = 1 if actual_lt <= promised_lt else 0

        # Freight mode & cost (expedited air freight triggered on spot orders or urgent late orders)
        material_spend = round(qty * actual_unit_price, 2)
        if sup["supplier_tier"] == "Spot / Maverick" or (delay_days >= 7 and random.random() < 0.45):
            freight_mode = "Air Express"
            freight_pct = np.random.uniform(0.14, 0.25)
        elif sup_id == "SUP-007":
            freight_mode = "Sea + Customs"
            freight_pct = np.random.uniform(0.07, 0.11)
        else:
            freight_mode = sup["default_freight_mode"]
            freight_pct = np.random.uniform(0.025, 0.055)

        freight_cost = round(material_spend * freight_pct, 2)
        total_landed_cost = round(material_spend + freight_cost, 2)
        price_variance_inr = round((actual_unit_price - bench_price) * qty, 2)

        po_rows.append({
            "po_id": po_id,
            "po_date": po_date.strftime("%Y-%m-%d"),
            "po_month": po_date.strftime("%Y-%m"),
            "sku_id": sku["sku_id"],
            "category": sku["category"],
            "supplier_id": sup_id,
            "order_qty": qty,
            "rfq_benchmark_unit_price_inr": bench_price,
            "actual_unit_price_inr": actual_unit_price,
            "material_spend_inr": material_spend,
            "freight_mode": freight_mode,
            "freight_cost_inr": freight_cost,
            "total_landed_cost_inr": total_landed_cost,
            "price_variance_inr": price_variance_inr,
            "promised_lead_time_days": promised_lt,
            "actual_lead_time_days": actual_lt,
            "promised_delivery_date": promised_delivery_date.strftime("%Y-%m-%d"),
            "actual_delivery_date": actual_delivery_date.strftime("%Y-%m-%d"),
            "on_time_delivery_flag": otd_flag,
        })

    return pd.DataFrame(po_rows)


def generate_rfq_quotes() -> pd.DataFrame:
    """
    Generates a 12-item Industrial Robotics Wiring Harness BOM quoted across 4 suppliers
    for Commercial RFQ Evaluation & Should-Cost Benchmarking.
    """
    bom_items = [
        ("BOM-01", "CON-101", "Heavy-Duty 24-Pin Rectangular Housing", 2, 185.0),
        ("BOM-02", "CON-102", "M12 8-Pin Circular Industrial Plug", 4, 142.0),
        ("BOM-03", "CON-103", "Deutsch DT 12-Pin Sealed Receptacle", 2, 168.0),
        ("BOM-04", "CBL-111", "18 AWG Braided Shielded Twisted Pair (m)", 12, 48.0),
        ("BOM-05", "CBL-115", "CAT6A Industrial Drag-Chain Ethernet Cable (m)", 6, 78.0),
        ("BOM-06", "CBL-117", "CAN-Bus J1939 Twinax Shielded Cable (m)", 8, 68.0),
        ("BOM-07", "BSL-129", "EMI/RFI 90-Degree Aluminum Die-Cast Backshell", 2, 155.0),
        ("BOM-08", "BSL-130", "Heat-Shrinkable Polyolefin Transition Boot", 6, 44.0),
        ("BOM-09", "BSL-131", "Silicone Multi-Wire Grommet Seal (IP67)", 8, 22.0),
        ("BOM-10", "BSL-134", "Convoluted Split Loom Tubing PA6 (m)", 10, 28.0),
        ("BOM-11", "BSL-136", "Braided Nomex Abrasion Sleeve (m)", 5, 65.0),
        ("BOM-12", "HRN-128", "Sub-Assembly Crimping, Overmolding & Hi-Pot Testing", 1, 680.0),
    ]

    # 4 competing bidders for annual contract of 2,500 harness assemblies
    bidders = [
        ("SUP-001", "Precision Interconnect India Pvt Ltd", 0.98, 45000, 14, 0.030, 94),
        ("SUP-002", "Deccan Cable & Harness Systems", 1.01, 30000, 16, 0.035, 91),
        ("SUP-003", "Kaveri Electro-Mech Components", 0.96, 65000, 15, 0.032, 95),
        ("SUP-007", "Shenzhen SinoConnect Electronics", 0.91, 85000, 26, 0.095, 86),
    ]

    rows = []
    annual_assemblies = 2500
    for line_no, sku_id, desc, qty_per_assy, should_cost in bom_items:
        for sup_id, sup_name, mult, tooling_inr, lt_days, freight_pct, tech_score in bidders:
            # Slight item-level variation per bidder
            item_var = np.random.uniform(-0.02, 0.02)
            quoted_unit = round(should_cost * (mult + item_var), 2)
            # Prorate tooling only on BOM-12 (assembly tooling) so totals don't double-count
            line_tooling = tooling_inr if line_no == "BOM-12" else 0
            rows.append({
                "bom_line": line_no,
                "sku_id": sku_id,
                "component_description": desc,
                "qty_per_assembly": qty_per_assy,
                "annual_assembly_volume": annual_assemblies,
                "should_cost_unit_inr": should_cost,
                "supplier_id": sup_id,
                "supplier_name": sup_name,
                "quoted_unit_price_inr": quoted_unit,
                "nre_tooling_cost_inr": line_tooling,
                "freight_duty_pct": freight_pct,
                "quoted_lead_time_days": lt_days,
                "technical_quality_score": tech_score,
            })

    return pd.DataFrame(rows)


def main():
    print("Generating synthetic procurement datasets for TE Connectivity portfolio...")
    suppliers_df = generate_suppliers()
    skus_df = generate_sku_master()
    pos_df = generate_purchase_orders(suppliers_df, skus_df, n_rows=2800)
    rfq_df = generate_rfq_quotes()

    suppliers_df.to_csv(os.path.join(DATA_DIR, "suppliers.csv"), index=False)
    skus_df.to_csv(os.path.join(DATA_DIR, "sku_master.csv"), index=False)
    pos_df.to_csv(os.path.join(DATA_DIR, "purchase_orders.csv"), index=False)
    rfq_df.to_csv(os.path.join(DATA_DIR, "rfq_bom_quotes.csv"), index=False)

    print(f"Saved suppliers.csv       : {len(suppliers_df):,} rows")
    print(f"Saved sku_master.csv      : {len(skus_df):,} rows")
    print(f"Saved purchase_orders.csv : {len(pos_df):,} rows")
    print(f"Saved rfq_bom_quotes.csv  : {len(rfq_df):,} rows")
    print(f"Total Landed Spend (INR)  : ₹{pos_df['total_landed_cost_inr'].sum():,.2f}")
    print(f"Net Price Variance (INR)  : ₹{pos_df['price_variance_inr'].sum():,.2f}")


if __name__ == "__main__":
    main()