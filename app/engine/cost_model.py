from __future__ import annotations

import numpy as np
import pandas as pd


DRIVER_NAMES = [
    "Oil Price",
    "FX Rate",
    "Steel Price",
    "Copper Price",
    "Lithium Price",
    "Electricity Cost",
    "Labor Cost",
    "Shipping Cost",
    "Inflation",
]

MODEL_COST_ALLOCATIONS = {
    "BYD-DOLPHIN-TH": {
        "Battery System": 0.32,
        "Body & Chassis": 0.16,
        "Electric Drive Unit": 0.12,
        "Electronics & Software": 0.10,
        "Interior & Comfort": 0.09,
        "Manufacturing & Factory Overhead": 0.08,
        "Logistics / Import / Warehouse": 0.05,
        "Warranty / Quality / Risk Reserve": 0.04,
        "R&D / SG&A Allocation": 0.04,
    },
    "BYD-ATTO3-TH": {
        "Battery System": 0.34,
        "Body & Chassis": 0.17,
        "Electric Drive Unit": 0.12,
        "Electronics & Software": 0.11,
        "Interior & Comfort": 0.08,
        "Manufacturing & Factory Overhead": 0.07,
        "Logistics / Import / Warehouse": 0.04,
        "Warranty / Quality / Risk Reserve": 0.03,
        "R&D / SG&A Allocation": 0.04,
    },
    "BYD-SEAL-EV-TH": {
        "Battery System": 0.37,
        "Body & Chassis": 0.16,
        "Electric Drive Unit": 0.13,
        "Electronics & Software": 0.12,
        "Interior & Comfort": 0.07,
        "Manufacturing & Factory Overhead": 0.06,
        "Logistics / Import / Warehouse": 0.03,
        "Warranty / Quality / Risk Reserve": 0.03,
        "R&D / SG&A Allocation": 0.03,
    },
    "BYD-SEALION6-TH": {
        "Battery System": 0.35,
        "Body & Chassis": 0.18,
        "Electric Drive Unit": 0.12,
        "Electronics & Software": 0.11,
        "Interior & Comfort": 0.08,
        "Manufacturing & Factory Overhead": 0.07,
        "Logistics / Import / Warehouse": 0.04,
        "Warranty / Quality / Risk Reserve": 0.03,
        "R&D / SG&A Allocation": 0.02,
    },
    "BYD-M6-TH": {
        "Battery System": 0.33,
        "Body & Chassis": 0.19,
        "Electric Drive Unit": 0.11,
        "Electronics & Software": 0.10,
        "Interior & Comfort": 0.10,
        "Manufacturing & Factory Overhead": 0.08,
        "Logistics / Import / Warehouse": 0.04,
        "Warranty / Quality / Risk Reserve": 0.03,
        "R&D / SG&A Allocation": 0.02,
    },
}

LOCAL_COUNTRIES = {"Thailand"}
REGIONAL_COUNTRIES = {"China", "Malaysia", "Korea", "Japan"}


def calculate_change_percent(current_value: float, scenario_value: float) -> float:
    if current_value <= 0:
        raise ValueError("Current value must be greater than 0.")
    return (scenario_value - current_value) / current_value


def calculate_target_total_cost(vehicle_sales_price: float, target_gp_percent: float) -> float:
    return vehicle_sales_price * (1 - target_gp_percent / 100)


def calculate_vehicle_gp(vehicle_sales_price: float, total_vehicle_cost: float) -> tuple[float, float]:
    gp_amount = vehicle_sales_price - total_vehicle_cost
    gp_percent = gp_amount / vehicle_sales_price * 100 if vehicle_sales_price else 0.0
    return gp_amount, gp_percent


def supplier_region(country: str, currency: str) -> str:
    if country in LOCAL_COUNTRIES or currency == "THB":
        return "local"
    if country in REGIONAL_COUNTRIES:
        return "regional"
    return "overseas"


def build_supplier_exposure_by_part(
    supplier_map: pd.DataFrame,
    suppliers: pd.DataFrame,
) -> pd.DataFrame:
    if supplier_map.empty:
        return pd.DataFrame(columns=["part_code", "supplier_oil_factor", "supplier_fx_factor", "supplier_shipping_factor"])

    supplier_detail = supplier_map.merge(suppliers, on="supplier_name", how="left")
    regions = supplier_detail.apply(
        lambda row: supplier_region(str(row.get("country", "")), str(row.get("currency", ""))),
        axis=1,
    )
    supplier_detail["supplier_oil_factor"] = regions.map(
        {"local": 0.03, "regional": 0.06, "overseas": 0.10}
    )
    supplier_detail["supplier_fx_factor"] = regions.map(
        {"local": 0.00, "regional": 0.35, "overseas": 0.70}
    )
    supplier_detail["supplier_shipping_factor"] = regions.map(
        {"local": 0.01, "regional": 0.04, "overseas": 0.08}
    )
    for column in ["supplier_oil_factor", "supplier_fx_factor", "supplier_shipping_factor"]:
        supplier_detail[column] = supplier_detail[column] * supplier_detail["allocation_percent"].astype(float)
    exposure = (
        supplier_detail.groupby("part_code", as_index=False)[
            ["supplier_oil_factor", "supplier_fx_factor", "supplier_shipping_factor"]
        ]
        .sum()
    )
    factor_columns = ["supplier_oil_factor", "supplier_fx_factor", "supplier_shipping_factor"]
    exposure[factor_columns] = exposure[factor_columns].clip(lower=0)
    return exposure


def _contains(value: object, *needles: str) -> bool:
    text = str(value).lower()
    return any(needle.lower() in text for needle in needles)


def calculate_row_driver_exposures(row: pd.Series) -> dict[str, float]:
    cost_element = str(row.get("cost_element", ""))
    module = str(row.get("module_name", ""))
    part_name = str(row.get("part_name", ""))
    material_group = str(row.get("material_group", ""))

    exposures = {driver: 0.0 for driver in DRIVER_NAMES}

    # Oil exposure follows explicit EV benchmark pass-through assumptions.
    if cost_element == "Logistics Cost":
        exposures["Oil Price"] += 0.25
    if module == "Logistics / Import / Warehouse" or _contains(part_name, "freight", "logistics", "inbound", "outbound"):
        exposures["Oil Price"] = max(exposures["Oil Price"], 0.30)
    if cost_element == "Energy Cost":
        exposures["Oil Price"] += 0.08
    if cost_element == "ESG/Carbon Cost":
        exposures["Oil Price"] += 0.05
    if cost_element in {"Raw Material Cost", "Purchased Parts Cost"} and _contains(material_group, "plastic"):
        exposures["Oil Price"] += 0.08
    if cost_element in {"Raw Material Cost", "Purchased Parts Cost"} and _contains(material_group, "rubber", "tire"):
        exposures["Oil Price"] += 0.10
    if cost_element == "Purchased Parts Cost":
        exposures["Oil Price"] += float(row.get("supplier_oil_factor", 0.0))

    # Steel exposure is concentrated in body/chassis, steel-heavy parts, and tooling.
    if module == "Body & Chassis" and cost_element in {"Raw Material Cost", "Purchased Parts Cost"}:
        exposures["Steel Price"] = max(exposures["Steel Price"], 0.35)
    if _contains(part_name, "steel body", "chassis frame", "stamped"):
        exposures["Steel Price"] = max(exposures["Steel Price"], 0.35)
    if cost_element == "Tooling Cost":
        exposures["Steel Price"] = max(exposures["Steel Price"], 0.20)
    if _contains(part_name, "suspension", "brake"):
        exposures["Steel Price"] = max(exposures["Steel Price"], 0.15)

    if _contains(part_name, "wiring harness", "hv wiring"):
        exposures["Copper Price"] = max(exposures["Copper Price"], 0.35)
    if _contains(part_name, "e-motor", "motor"):
        exposures["Copper Price"] = max(exposures["Copper Price"], 0.18)
    if _contains(part_name, "inverter"):
        exposures["Copper Price"] = max(exposures["Copper Price"], 0.15)
    if module == "Electronics & Software" and cost_element in {
        "Raw Material Cost",
        "Purchased Parts Cost",
        "Semiconductor Cost",
    }:
        exposures["Copper Price"] = max(exposures["Copper Price"], 0.05)

    if cost_element == "Purchased Parts Cost":
        exposures["FX Rate"] = max(exposures["FX Rate"], float(row.get("supplier_fx_factor", 0.0)))
    if cost_element == "Semiconductor Cost":
        exposures["FX Rate"] = max(exposures["FX Rate"], 0.70)
    if cost_element == "Software Cost":
        exposures["FX Rate"] = max(exposures["FX Rate"], 0.35)
    if cost_element == "Import Duty":
        exposures["FX Rate"] = max(exposures["FX Rate"], float(row.get("supplier_fx_factor", 0.0)), 0.35)
    if cost_element == "Logistics Cost":
        exposures["FX Rate"] = max(exposures["FX Rate"], float(row.get("supplier_fx_factor", 0.0)))

    exposures["Labor Cost"] = {
        "Direct Labor Cost": 1.00,
        "Indirect Labor Cost": 1.00,
        "Quality Cost": 0.30,
        "R&D Allocation": 0.40,
        "Software Cost": 0.35,
        "Corporate SG&A Allocation": 0.30,
    }.get(cost_element, 0.0)

    exposures["Electricity Cost"] = {
        "Energy Cost": 1.00,
        "Machine Cost": 0.25,
    }.get(cost_element, 0.0)
    if module == "Manufacturing & Factory Overhead":
        exposures["Electricity Cost"] = max(exposures["Electricity Cost"], 0.20)
    if module == "Battery System" and _contains(part_name, "assembly"):
        exposures["Electricity Cost"] = max(exposures["Electricity Cost"], 0.15)

    if module == "Logistics / Import / Warehouse":
        exposures["Shipping Cost"] = max(exposures["Shipping Cost"], 0.70)
    if cost_element == "Logistics Cost":
        exposures["Shipping Cost"] = max(exposures["Shipping Cost"], 0.70)
    if cost_element == "Purchased Parts Cost":
        exposures["Shipping Cost"] = max(exposures["Shipping Cost"], float(row.get("supplier_shipping_factor", 0.0)))

    if cost_element == "Purchased Parts Cost":
        exposures["Inflation"] = 0.30
    if cost_element in {"Direct Labor Cost", "Indirect Labor Cost"}:
        exposures["Inflation"] = 0.25
    if cost_element == "Logistics Cost":
        exposures["Inflation"] = 0.20
    if cost_element == "Corporate SG&A Allocation":
        exposures["Inflation"] = 0.25
    if cost_element in {"Warranty Reserve", "Risk Reserve"}:
        exposures["Inflation"] = 0.20

    return exposures


def build_driver_weight_matrix(
    breakdown: pd.DataFrame,
    lithium_exposure_of_battery: float,
) -> tuple[np.ndarray, list[str], pd.Series]:
    weights = np.zeros((len(breakdown), len(DRIVER_NAMES)))
    for row_index, (_, row) in enumerate(breakdown.iterrows()):
        exposures = calculate_row_driver_exposures(row)
        for driver_index, driver_name in enumerate(DRIVER_NAMES):
            weights[row_index, driver_index] = exposures[driver_name]

    battery_mask = breakdown["module_name"].eq("Battery System")
    battery_cost = float(breakdown.loc[battery_mask, "cost_amount"].sum())
    if battery_cost:
        lithium_index = DRIVER_NAMES.index("Lithium Price")
        weights[battery_mask.to_numpy(), lithium_index] = (
            battery_cost
            * lithium_exposure_of_battery
            * (breakdown.loc[battery_mask, "cost_amount"] / battery_cost)
            / breakdown.loc[battery_mask, "cost_amount"]
        ).to_numpy()
    return weights, DRIVER_NAMES.copy(), battery_mask


def calculate_element_impacts(
    breakdown: pd.DataFrame,
    changes: dict[str, float],
    lithium_exposure_of_battery: float,
) -> pd.DataFrame:
    impacted = breakdown.copy()
    weights, driver_names, _ = build_driver_weight_matrix(impacted, lithium_exposure_of_battery)
    change_vector = np.array([changes.get(driver_name, 0.0) for driver_name in driver_names])
    driver_impacts = weights * impacted["cost_amount"].to_numpy()[:, None] * change_vector[None, :]
    for driver_index, driver_name in enumerate(driver_names):
        impacted[f"{driver_name} impact"] = driver_impacts[:, driver_index]
    impacted["element_impact"] = driver_impacts.sum(axis=1)
    return impacted


def calculate_driver_impact_matrix(
    breakdown: pd.DataFrame,
    driver_change_matrix: np.ndarray,
    lithium_exposure_of_battery: float,
) -> np.ndarray:
    weights, _, _ = build_driver_weight_matrix(breakdown, lithium_exposure_of_battery)
    row_change_matrix = driver_change_matrix @ weights.T
    return row_change_matrix * breakdown["cost_amount"].to_numpy()


def classify_financial_risk(
    gp_after_percent: float,
    target_gp_percent: float,
    net_cost_impact: float,
    base_total_cost: float,
    worst_case_gp_percent: float | None = None,
    material_supplier_impact_percent: float = 0.0,
) -> str:
    gp_gap = target_gp_percent - gp_after_percent
    net_cost_impact_percent = net_cost_impact / base_total_cost * 100 if base_total_cost else 0.0
    if gp_gap > 5 or net_cost_impact_percent > 8:
        return "critical"
    if 2 < gp_gap <= 5 or 5 <= net_cost_impact_percent <= 8:
        return "high"
    if 0 <= gp_gap <= 2 or 2 <= net_cost_impact_percent < 5:
        return "warning"
    if material_supplier_impact_percent >= 2:
        return "warning"
    return "safe"
