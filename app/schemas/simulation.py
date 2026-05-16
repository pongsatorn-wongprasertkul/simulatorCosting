from pydantic import BaseModel, Field


class SimulationRequest(BaseModel):
    product_name: str = Field(..., min_length=1, max_length=255)
    base_cost: float = Field(..., gt=0)
    selling_price: float | None = Field(default=None, gt=0)
    oil_change: float = 0.0
    fx_change: float = 0.0
    steel_change: float = 0.0
    oil_factor: float | None = None
    fx_factor: float | None = None
    steel_factor: float | None = None


class SimulationResponse(BaseModel):
    product_name: str
    old_cost: float
    new_cost: float
    impact_amount: float
    gp_amount: float
    gp_percent: float
