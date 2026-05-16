import pandas as pd


OFFSET_CATEGORY_TO_COST_ELEMENT = {
    "direct_labor": "Direct Labor Cost",
    "indirect_labor": "Indirect Labor Cost",
    "machine": "Machine Cost",
    "energy": "Energy Cost",
    "logistics": "Logistics Cost",
    "supplier_discount": "Purchased Parts Cost",
    "tooling_amortization": "Tooling Cost",
    "sga": "Corporate SG&A Allocation",
    "warranty_reserve": "Warranty Reserve",
    "risk_reserve": "Risk Reserve",
}


def calculate_offset_mitigation(
    cost_breakdown: pd.DataFrame,
    reduction_percentages: dict[str, float],
    commodity_cost_increase: float,
    base_total_cost: float,
    sales_price: float,
    target_gp_percent: float,
) -> dict[str, object]:
    offset_rows = []

    for category, cost_element in OFFSET_CATEGORY_TO_COST_ELEMENT.items():
        reduction_percent = reduction_percentages.get(category, 0.0)
        category_cost = float(
            cost_breakdown.loc[
                cost_breakdown["cost_element"] == cost_element,
                "cost_amount",
            ].sum()
        )
        savings = category_cost * reduction_percent
        offset_rows.append(
            {
                "category": category,
                "cost_element": cost_element,
                "category_cost": category_cost,
                "reduction_percent": reduction_percent,
                "savings": savings,
            }
        )

    offset_savings = sum(row["savings"] for row in offset_rows)
    net_impact = commodity_cost_increase - offset_savings
    new_total_cost_after_offset = base_total_cost + net_impact
    gp_before = sales_price - (base_total_cost + commodity_cost_increase)
    gp_after = sales_price - new_total_cost_after_offset
    gp_after_percent = gp_after / sales_price * 100 if sales_price else 0.0
    target_gp_amount = sales_price * target_gp_percent / 100
    required_additional_savings = max(target_gp_amount - gp_after, 0.0)
    target_maintained = gp_after_percent >= target_gp_percent

    return {
        "commodity_cost_increase": commodity_cost_increase,
        "offset_savings": offset_savings,
        "net_impact": net_impact,
        "new_total_cost_after_offset": new_total_cost_after_offset,
        "gp_before": gp_before,
        "gp_after": gp_after,
        "gp_after_percent": gp_after_percent,
        "target_gp_percent": target_gp_percent,
        "gap_to_target": gp_after_percent - target_gp_percent,
        "required_additional_savings": required_additional_savings,
        "target_maintained": target_maintained,
        "offset_rows": offset_rows,
    }
