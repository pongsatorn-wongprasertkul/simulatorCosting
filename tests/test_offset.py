import pandas as pd

from app.engine.offset import calculate_offset_mitigation


def test_calculate_offset_mitigation_offsets_commodity_increase() -> None:
    breakdown = pd.DataFrame(
        [
            {"cost_element": "Direct Labor Cost", "cost_amount": 1000},
            {"cost_element": "Indirect Labor Cost", "cost_amount": 500},
            {"cost_element": "Machine Cost", "cost_amount": 300},
            {"cost_element": "Energy Cost", "cost_amount": 200},
            {"cost_element": "Logistics Cost", "cost_amount": 400},
            {"cost_element": "Purchased Parts Cost", "cost_amount": 2000},
            {"cost_element": "Tooling Cost", "cost_amount": 600},
            {"cost_element": "Corporate SG&A Allocation", "cost_amount": 700},
            {"cost_element": "Warranty Reserve", "cost_amount": 800},
            {"cost_element": "Risk Reserve", "cost_amount": 900},
        ]
    )

    result = calculate_offset_mitigation(
        cost_breakdown=breakdown,
        reduction_percentages={
            "direct_labor": 0.05,
            "supplier_discount": 0.02,
            "risk_reserve": 0.10,
        },
        commodity_cost_increase=800,
        base_total_cost=10000,
        sales_price=15000,
        target_gp_percent=25,
    )

    assert result["offset_savings"] == 180
    assert result["net_impact"] == 620
    assert result["new_total_cost_after_offset"] == 10620
    assert result["gp_before"] == 4200
    assert result["gp_after"] == 4380
    assert result["gp_after_percent"] == 29.2
    assert result["target_maintained"] is True
    assert result["required_additional_savings"] == 0


def test_calculate_offset_mitigation_reports_required_savings_gap() -> None:
    breakdown = pd.DataFrame(
        [{"cost_element": "Direct Labor Cost", "cost_amount": 1000}]
    )

    result = calculate_offset_mitigation(
        cost_breakdown=breakdown,
        reduction_percentages={"direct_labor": 0.05},
        commodity_cost_increase=2000,
        base_total_cost=10000,
        sales_price=12000,
        target_gp_percent=20,
    )

    assert result["offset_savings"] == 50
    assert result["net_impact"] == 1950
    assert result["gp_after"] == 50
    assert result["target_maintained"] is False
    assert result["required_additional_savings"] == 2350
