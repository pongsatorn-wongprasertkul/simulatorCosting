import pandas as pd

DEFAULT_DRIVER_WEIGHTS = {
    "oil_factor": 0.18,
    "fx_factor": 0.22,
    "steel_factor": 0.35,
    "copper_factor": 0.10,
    "lithium_factor": 0.12,
    "electricity_factor": 0.08,
    "labor_factor": 0.10,
    "shipping_factor": 0.08,
}

REQUIRED_COLUMNS = {"base_cost", "oil_change", "fx_change", "steel_change"}
OPTIONAL_DRIVER_PAIRS = {
    "copper_change": "copper_factor",
    "lithium_change": "lithium_factor",
    "electricity_change": "electricity_factor",
    "labor_change": "labor_factor",
    "shipping_change": "shipping_factor",
}


def simulate_cost(input_df: pd.DataFrame) -> pd.DataFrame:
    """Simulate cost changes from oil, FX, and steel movements.

    Expected input columns:
    - base_cost
    - oil_change
    - fx_change
    - steel_change

    Optional input columns:
    - oil_factor
    - fx_factor
    - steel_factor
    - copper_change / copper_factor
    - lithium_change / lithium_factor
    - electricity_change / electricity_factor
    - labor_change / labor_factor
    - shipping_change / shipping_factor
    - selling_price
    """
    missing_columns = REQUIRED_COLUMNS - set(input_df.columns)
    if missing_columns:
        missing = ", ".join(sorted(missing_columns))
        raise ValueError(f"Missing required columns: {missing}")

    output_df = input_df.copy()

    for factor_name, default_value in DEFAULT_DRIVER_WEIGHTS.items():
        if factor_name not in output_df.columns:
            output_df[factor_name] = default_value

    output_df["old_cost"] = output_df["base_cost"]
    multiplier = (
        1
        + output_df["oil_change"] * output_df["oil_factor"]
        + output_df["fx_change"] * output_df["fx_factor"]
        + output_df["steel_change"] * output_df["steel_factor"]
    )
    for change_column, factor_column in OPTIONAL_DRIVER_PAIRS.items():
        if change_column in output_df.columns:
            multiplier = multiplier + output_df[change_column] * output_df[factor_column]

    output_df["new_cost"] = output_df["base_cost"] * multiplier
    output_df["impact_amount"] = output_df["new_cost"] - output_df["old_cost"]

    if "selling_price" in output_df.columns:
        output_df["gp_amount"] = output_df["selling_price"] - output_df["new_cost"]
        output_df["gp_percent"] = output_df["gp_amount"] / output_df["selling_price"]
    else:
        output_df["gp_amount"] = -output_df["impact_amount"]
        output_df["gp_percent"] = output_df["gp_amount"] / output_df["old_cost"]

    return output_df[
        [
            "old_cost",
            "new_cost",
            "impact_amount",
            "gp_amount",
            "gp_percent",
        ]
    ]
