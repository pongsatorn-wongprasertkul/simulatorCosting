from pathlib import Path

import pandas as pd
import pytest

from app.engine.cost_model import (
    MODEL_COST_ALLOCATIONS,
    build_supplier_exposure_by_part,
    calculate_element_impacts,
    calculate_target_total_cost,
    classify_financial_risk,
)


ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"


def _normalized_seal_data(sales_price: float = 1_325_000, target_gp: float = 24):
    parts = pd.read_csv(DATA / "vehicle_parts.csv")
    breakdown = pd.read_csv(DATA / "part_cost_breakdown.csv")
    suppliers = pd.read_csv(DATA / "supplier_master.csv")
    supplier_map = pd.read_csv(DATA / "part_supplier_map.csv")

    vehicle_code = "BYD-SEAL-EV-TH"
    parts = parts[parts["vehicle_code"].eq(vehicle_code)].copy()
    breakdown = breakdown[breakdown["vehicle_code"].eq(vehicle_code)].copy()
    supplier_map = supplier_map[supplier_map["vehicle_code"].eq(vehicle_code)].copy()

    target_total_cost = calculate_target_total_cost(sales_price, target_gp)
    scale = target_total_cost / parts["total_base_cost"].sum()
    parts["total_base_cost"] = parts["total_base_cost"] * scale
    breakdown["cost_amount"] = breakdown["cost_amount"] * scale

    part_targets = parts.set_index("part_code")["total_base_cost"].to_dict()
    for part_code, part_cost in part_targets.items():
        indexes = breakdown.index[breakdown["part_code"].eq(part_code)]
        breakdown.loc[indexes[-1], "cost_amount"] += part_cost - breakdown.loc[indexes, "cost_amount"].sum()

    enriched = breakdown.merge(parts, on=["vehicle_code", "part_code"], how="left")
    supplier_exposure = build_supplier_exposure_by_part(supplier_map, suppliers)
    enriched = enriched.merge(supplier_exposure, on="part_code", how="left")
    for column in ["supplier_oil_factor", "supplier_fx_factor", "supplier_shipping_factor"]:
        enriched[column] = enriched[column].fillna(0.0)
    return parts, breakdown, enriched, target_total_cost


def test_byd_seal_base_cost_from_sales_price_and_target_gp() -> None:
    assert calculate_target_total_cost(1_325_000, 24) == 1_007_000


def test_model_cost_allocations_sum_to_100_percent() -> None:
    for allocation in MODEL_COST_ALLOCATIONS.values():
        assert sum(allocation.values()) == pytest.approx(1.0)


def test_byd_seal_lithium_shock_matches_benchmark_formula() -> None:
    _, _, enriched, target_total_cost = _normalized_seal_data()
    changes = {driver: 0.0 for driver in [
        "Oil Price", "FX Rate", "Steel Price", "Copper Price", "Lithium Price",
        "Electricity Cost", "Labor Cost", "Shipping Cost", "Inflation",
    ]}
    changes["Lithium Price"] = (21_750 - 14_500) / 14_500

    impacted = calculate_element_impacts(enriched, changes, lithium_exposure_of_battery=0.25)

    expected = target_total_cost * MODEL_COST_ALLOCATIONS["BYD-SEAL-EV-TH"]["Battery System"] * 0.25 * 0.50
    actual = impacted["Lithium Price impact"].sum()
    assert actual > 30_000
    assert actual == pytest.approx(expected, rel=0.10)


def test_byd_seal_oil_shock_is_material_when_pass_through_enabled() -> None:
    _, _, enriched, _ = _normalized_seal_data()
    changes = {driver: 0.0 for driver in [
        "Oil Price", "FX Rate", "Steel Price", "Copper Price", "Lithium Price",
        "Electricity Cost", "Labor Cost", "Shipping Cost", "Inflation",
    ]}
    changes["Oil Price"] = (100 - 82) / 82

    impacted = calculate_element_impacts(enriched, changes, lithium_exposure_of_battery=0.25)

    assert impacted["Oil Price impact"].sum() > 5_000


def test_low_net_impact_and_slight_gp_gap_is_not_high_or_critical() -> None:
    risk = classify_financial_risk(
        gp_after_percent=23.92,
        target_gp_percent=24.0,
        net_cost_impact=1_107.7,
        base_total_cost=1_007_000,
        material_supplier_impact_percent=0,
    )
    assert risk in {"safe", "warning"}


def test_vehicle_total_equals_sum_of_module_part_and_breakdown_costs() -> None:
    parts, breakdown, _, target_total_cost = _normalized_seal_data()

    assert parts["total_base_cost"].sum() == pytest.approx(target_total_cost, abs=0.5)
    module_sum = parts.groupby("module_name")["total_base_cost"].sum().sum()
    assert module_sum == pytest.approx(parts["total_base_cost"].sum(), abs=0.5)

    part_breakdown = breakdown.groupby("part_code")["cost_amount"].sum()
    part_costs = parts.set_index("part_code")["total_base_cost"]
    assert part_breakdown.sub(part_costs, fill_value=0).abs().max() <= 0.5
