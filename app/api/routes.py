import pandas as pd
from fastapi import APIRouter, Depends
from sqlalchemy.orm import Session

from app.database.session import get_db
from app.engine.simulator import DEFAULT_DRIVER_WEIGHTS, simulate_cost
from app.models.simulation import SimulationRun
from app.schemas.simulation import SimulationRequest, SimulationResponse

router = APIRouter(prefix="/simulations", tags=["simulations"])


@router.post("", response_model=SimulationResponse)
def create_simulation(
    request: SimulationRequest,
    db: Session = Depends(get_db),
) -> SimulationResponse:
    input_row = {
        "base_cost": request.base_cost,
        "oil_change": request.oil_change,
        "fx_change": request.fx_change,
        "steel_change": request.steel_change,
    }
    if request.selling_price is not None:
        input_row["selling_price"] = request.selling_price
    input_row["oil_factor"] = (
        request.oil_factor
        if request.oil_factor is not None
        else DEFAULT_DRIVER_WEIGHTS["oil_factor"]
    )
    input_row["fx_factor"] = (
        request.fx_factor
        if request.fx_factor is not None
        else DEFAULT_DRIVER_WEIGHTS["fx_factor"]
    )
    input_row["steel_factor"] = (
        request.steel_factor
        if request.steel_factor is not None
        else DEFAULT_DRIVER_WEIGHTS["steel_factor"]
    )

    result_row = simulate_cost(pd.DataFrame([input_row])).iloc[0]
    result = SimulationResponse(
        product_name=request.product_name,
        old_cost=round(float(result_row["old_cost"]), 2),
        new_cost=round(float(result_row["new_cost"]), 2),
        impact_amount=round(float(result_row["impact_amount"]), 2),
        gp_amount=round(float(result_row["gp_amount"]), 2),
        gp_percent=round(float(result_row["gp_percent"]), 6),
    )

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

    return result


@router.get("", response_model=list[SimulationResponse])
def list_simulations(db: Session = Depends(get_db)) -> list[SimulationResponse]:
    runs = db.query(SimulationRun).order_by(SimulationRun.created_at.desc()).all()
    return [SimulationResponse.model_validate_json(run.drivers_json) for run in runs]
