import pandas as pd

from app.engine.simulator import simulate_cost


def test_simulate_cost_uses_dataframe_input_and_output() -> None:
    input_df = pd.DataFrame(
        [
            {
                "base_cost": 1000,
                "selling_price": 1500,
                "oil_change": 0.10,
                "fx_change": -0.05,
                "steel_change": 0.20,
                "oil_factor": 0.20,
                "fx_factor": 0.10,
                "steel_factor": 0.30,
            }
        ]
    )

    result = simulate_cost(input_df)

    assert list(result.columns) == [
        "old_cost",
        "new_cost",
        "impact_amount",
        "gp_amount",
        "gp_percent",
    ]
    assert result.loc[0, "old_cost"] == 1000
    assert result.loc[0, "new_cost"] == 1075
    assert result.loc[0, "impact_amount"] == 75
    assert result.loc[0, "gp_amount"] == 425
    assert result.loc[0, "gp_percent"] == 425 / 1500
