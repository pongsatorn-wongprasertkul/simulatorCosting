import pandas as pd

from app.database.session import Base, SessionLocal, engine
from app.engine.simulator import DEFAULT_DRIVER_WEIGHTS
from app.engine.simulator import simulate_cost
from app.models.simulation import SimulationRun
from app.schemas.simulation import SimulationResponse


def main() -> None:
    Base.metadata.create_all(bind=engine)

    product_name = "Standard Industrial Component"
    result_row = simulate_cost(
        pd.DataFrame(
            [
                {
                    "base_cost": 1000,
                    "selling_price": 1500,
                    "oil_change": 0.10,
                    "fx_change": -0.03,
                    "steel_change": 0.08,
                    **DEFAULT_DRIVER_WEIGHTS,
                }
            ]
        )
    ).iloc[0]
    result = SimulationResponse(
        product_name=product_name,
        old_cost=round(float(result_row["old_cost"]), 2),
        new_cost=round(float(result_row["new_cost"]), 2),
        impact_amount=round(float(result_row["impact_amount"]), 2),
        gp_amount=round(float(result_row["gp_amount"]), 2),
        gp_percent=round(float(result_row["gp_percent"]), 6),
    )

    db = SessionLocal()
    try:
        run = SimulationRun(
            product_name=result.product_name,
            base_cost=result.old_cost,
            adjusted_cost=result.new_cost,
            variance_amount=result.impact_amount,
            variance_percent=result.gp_percent,
            drivers_json=result.model_dump_json(),
        )
        db.add(run)
        db.commit()
    finally:
        db.close()

    print("Initialized cost_simulation.db with sample data.")


if __name__ == "__main__":
    main()
