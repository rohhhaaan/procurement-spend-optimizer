-- ============================================================================
-- PROJECT: Supply Chain Delivery, Spend & Cost Optimization Analytics
-- DATABASE: SQLite (data/procurement_analytics.db)
-- AUTHOR: Supply Chain Analytics Portfolio (TE Connectivity Alignment)
-- ============================================================================

-- ----------------------------------------------------------------------------
-- VIEW 1: Enriched Purchase Order Master View (Fact + Dimensions)
-- ----------------------------------------------------------------------------
DROP VIEW IF EXISTS vw_po_enriched;
CREATE VIEW vw_po_enriched AS
SELECT 
    po.po_id,
    po.po_date,
    po.po_month,
    po.sku_id,
    sku.sku_description,
    po.category,
    po.supplier_id,
    sup.supplier_name,
    sup.supplier_tier,
    sup.location AS supplier_location,
    sup.payment_terms,
    po.order_qty,
    po.rfq_benchmark_unit_price_inr,
    po.actual_unit_price_inr,
    po.material_spend_inr,
    po.freight_mode,
    po.freight_cost_inr,
    po.total_landed_cost_inr,
    po.price_variance_inr,
    ROUND((po.actual_unit_price_inr - po.rfq_benchmark_unit_price_inr) / po.rfq_benchmark_unit_price_inr * 100.0, 2) AS unit_price_variance_pct,
    po.promised_lead_time_days,
    po.actual_lead_time_days,
    (po.actual_lead_time_days - po.promised_lead_time_days) AS lead_time_variance_days,
    po.on_time_delivery_flag
FROM purchase_orders po
JOIN suppliers sup ON po.supplier_id = sup.supplier_id
JOIN sku_master sku ON po.sku_id = sku.sku_id;


-- ----------------------------------------------------------------------------
-- QUERY 1: Executive Supplier Scorecard (Spend, Variance %, OTD %, Lead Time)
-- Uses CTEs + Window Functions to rank suppliers by total landed spend
-- ----------------------------------------------------------------------------
WITH supplier_metrics AS (
    SELECT 
        supplier_id,
        supplier_name,
        supplier_tier,
        supplier_location,
        COUNT(po_id) AS total_pos,
        SUM(order_qty) AS total_units_procured,
        ROUND(SUM(material_spend_inr), 2) AS total_material_spend_inr,
        ROUND(SUM(freight_cost_inr), 2) AS total_freight_spend_inr,
        ROUND(SUM(total_landed_cost_inr), 2) AS total_landed_spend_inr,
        ROUND(SUM(price_variance_inr), 2) AS net_price_variance_inr,
        ROUND(SUM(price_variance_inr) / SUM(rfq_benchmark_unit_price_inr * order_qty) * 100.0, 2) AS weighted_price_variance_pct,
        ROUND(AVG(on_time_delivery_flag) * 100.0, 2) AS otd_pct,
        ROUND(AVG(promised_lead_time_days), 1) AS avg_promised_lt_days,
        ROUND(AVG(actual_lead_time_days), 1) AS avg_actual_lt_days,
        ROUND(AVG(lead_time_variance_days), 2) AS avg_lt_slippage_days
    FROM vw_po_enriched
    GROUP BY supplier_id, supplier_name, supplier_tier, supplier_location
)
SELECT 
    RANK() OVER (ORDER BY total_landed_spend_inr DESC) AS spend_rank,
    supplier_id,
    supplier_name,
    supplier_tier,
    total_pos,
    total_landed_spend_inr,
    ROUND(total_landed_spend_inr * 100.0 / SUM(total_landed_spend_inr) OVER (), 2) AS spend_share_pct,
    net_price_variance_inr,
    weighted_price_variance_pct,
    otd_pct,
    avg_promised_lt_days,
    avg_actual_lt_days,
    avg_lt_slippage_days,
    CASE 
        WHEN otd_pct >= 92.0 AND weighted_price_variance_pct <= 2.0 THEN 'Preferred Strategic'
        WHEN otd_pct >= 85.0 AND weighted_price_variance_pct <= 5.0 THEN 'Approved Core'
        ELSE 'High Risk / Consolidate'
    END AS supplier_status
FROM supplier_metrics
ORDER BY total_landed_spend_inr DESC;


-- ----------------------------------------------------------------------------
-- QUERY 2: Category Spend, Freight Burden & Air Express Leakage Analysis
-- Identifies logistics cost drivers across connectivity & box-build categories
-- ----------------------------------------------------------------------------
SELECT 
    category,
    COUNT(DISTINCT sku_id) AS active_skus,
    COUNT(po_id) AS po_count,
    ROUND(SUM(material_spend_inr), 2) AS material_spend_inr,
    ROUND(SUM(freight_cost_inr), 2) AS total_freight_inr,
    ROUND(SUM(total_landed_cost_inr), 2) AS total_landed_cost_inr,
    ROUND(SUM(freight_cost_inr) * 100.0 / SUM(material_spend_inr), 2) AS freight_to_material_pct,
    ROUND(SUM(CASE WHEN freight_mode = 'Air Express' THEN freight_cost_inr ELSE 0 END), 2) AS air_express_freight_inr,
    ROUND(SUM(CASE WHEN freight_mode = 'Air Express' THEN 1 ELSE 0 END) * 100.0 / COUNT(po_id), 2) AS air_shipment_share_pct,
    ROUND(AVG(on_time_delivery_flag) * 100.0, 2) AS category_otd_pct
FROM vw_po_enriched
GROUP BY category
ORDER BY total_landed_cost_inr DESC;


-- ----------------------------------------------------------------------------
-- QUERY 3: SKU-Level Vendor Consolidation & Should-Cost Savings Simulation
-- Quantifies exact INR savings from shifting Spot/Maverick POs to Tier-1 Contract
-- ----------------------------------------------------------------------------
WITH maverick_sku_spend AS (
    SELECT 
        sku_id,
        sku_description,
        category,
        COUNT(po_id) AS maverick_po_count,
        SUM(order_qty) AS maverick_units,
        ROUND(AVG(rfq_benchmark_unit_price_inr), 2) AS benchmark_unit_price_inr,
        ROUND(AVG(actual_unit_price_inr), 2) AS avg_maverick_unit_price_inr,
        ROUND(SUM(material_spend_inr), 2) AS actual_maverick_material_spend_inr,
        ROUND(SUM(rfq_benchmark_unit_price_inr * order_qty), 2) AS should_cost_material_spend_inr,
        ROUND(SUM(freight_cost_inr), 2) AS actual_air_freight_inr,
        -- Standard road freight benchmark is ~4.0% of should-cost material spend
        ROUND(SUM(rfq_benchmark_unit_price_inr * order_qty) * 0.04, 2) AS normalized_road_freight_inr
    FROM vw_po_enriched
    WHERE supplier_tier = 'Spot / Maverick'
    GROUP BY sku_id, sku_description, category
)
SELECT 
    sku_id,
    sku_description,
    category,
    maverick_po_count,
    maverick_units,
    benchmark_unit_price_inr,
    avg_maverick_unit_price_inr,
    ROUND(actual_maverick_material_spend_inr - should_cost_material_spend_inr, 2) AS material_price_savings_inr,
    ROUND(actual_air_freight_inr - normalized_road_freight_inr, 2) AS freight_mode_savings_inr,
    ROUND((actual_maverick_material_spend_inr - should_cost_material_spend_inr) + 
          (actual_air_freight_inr - normalized_road_freight_inr), 2) AS total_recoverable_savings_inr,
    ROUND(((actual_maverick_material_spend_inr - should_cost_material_spend_inr) + 
           (actual_air_freight_inr - normalized_road_freight_inr)) * 100.0 / 
          (actual_maverick_material_spend_inr + actual_air_freight_inr), 2) AS landed_cost_reduction_pct
FROM maverick_sku_spend
ORDER BY total_recoverable_savings_inr DESC
LIMIT 15;


-- ----------------------------------------------------------------------------
-- QUERY 4: Inventory Turnover, Days of Supply (DOS) & Slow-Moving Stock Risk
-- Flags excess working capital tied up in slow-moving connectivity SKUs
-- ----------------------------------------------------------------------------
WITH sku_po_activity AS (
    SELECT 
        sku_id,
        COUNT(po_id) AS historical_pos,
        SUM(order_qty) AS total_ordered_18m,
        ROUND(AVG(actual_unit_price_inr), 2) AS avg_paid_unit_price_inr
    FROM purchase_orders
    GROUP BY sku_id
)
SELECT 
    s.sku_id,
    s.sku_description,
    s.category,
    s.rfq_benchmark_unit_price_inr,
    s.annual_forecast_qty,
    s.on_hand_inventory_qty,
    s.safety_stock_qty,
    ROUND(s.on_hand_inventory_qty * s.rfq_benchmark_unit_price_inr, 2) AS inventory_value_inr,
    ROUND(CAST(s.annual_forecast_qty AS FLOAT) / NULLIF(s.on_hand_inventory_qty, 0), 2) AS annual_inventory_turns,
    ROUND(CAST(s.on_hand_inventory_qty AS FLOAT) / (CAST(s.annual_forecast_qty AS FLOAT) / 365.0), 1) AS days_of_supply_dos,
    -- Excess stock beyond 75 days max policy
    MAX(0, s.on_hand_inventory_qty - ROUND((s.annual_forecast_qty / 365.0) * 75)) AS excess_units,
    ROUND(MAX(0, s.on_hand_inventory_qty - ROUND((s.annual_forecast_qty / 365.0) * 75)) * s.rfq_benchmark_unit_price_inr, 2) AS excess_working_capital_inr,
    CASE 
        WHEN (CAST(s.on_hand_inventory_qty AS FLOAT) / (CAST(s.annual_forecast_qty AS FLOAT) / 365.0)) > 150 THEN 'Slow-Moving / Excess Risk'
        WHEN s.on_hand_inventory_qty <= s.safety_stock_qty THEN 'Reorder Alert'
        ELSE 'Optimal Stock'
    END AS inventory_health_status
FROM sku_master s
LEFT JOIN sku_po_activity p ON s.sku_id = p.sku_id
ORDER BY excess_working_capital_inr DESC;


-- ----------------------------------------------------------------------------
-- QUERY 5: Monthly Spend, Price Variance & OTD Trend (18-Month Time Series)
-- ----------------------------------------------------------------------------
SELECT 
    po_month,
    COUNT(po_id) AS monthly_pos,
    ROUND(SUM(material_spend_inr), 2) AS material_spend_inr,
    ROUND(SUM(freight_cost_inr), 2) AS freight_cost_inr,
    ROUND(SUM(total_landed_cost_inr), 2) AS total_landed_spend_inr,
    ROUND(SUM(price_variance_inr), 2) AS monthly_price_variance_inr,
    ROUND(AVG(on_time_delivery_flag) * 100.0, 2) AS monthly_otd_pct
FROM vw_po_enriched
GROUP BY po_month
ORDER BY po_month ASC;