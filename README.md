# Enterprise Cost Simulation System

A low-cost local-first cost simulation system for modeling business driver impact from changes in oil, FX, steel, labor, and logistics costs.

## Tech Stack

- Python
- FastAPI
- Streamlit
- SQLAlchemy
- SQLite
- pandas
- numpy
- pytest

## Project Structure

```text
app/
  api/          FastAPI routes and application factory
  dashboard/    Streamlit dashboard
  database/     SQLAlchemy engine, sessions, and initialization
  engine/       Cost simulation logic
  models/       SQLAlchemy ORM models
  schemas/      Pydantic request and response schemas
data/            BYD EV enterprise BOM and cost breakdown data
scripts/        Utility scripts
tests/          pytest test suite
```

## Local Setup

Create a virtual environment:

```bash
python -m venv .venv
```

Activate it on Windows PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
```

Install dependencies:

```bash
pip install -r requirements.txt
```

Initialize the local SQLite database with sample data:

```bash
python scripts/init_db.py
```

Run the API:

```bash
uvicorn app.api.main:app --reload
```

Open API docs:

```text
http://127.0.0.1:8000/docs
```

Run the dashboard in another terminal:

```bash
streamlit run app/dashboard/main.py
```

The dashboard uses:

- `data/vehicle_parts.csv` for the BYD EV hierarchical BOM, module, risk, and driver sensitivity data
- `data/part_cost_breakdown.csv` for 18 enterprise automotive cost elements per component

Run tests:

```bash
pytest
```

## Example Simulation

The simulation calculates cost impact from oil, FX, and steel changes:

```text
new_cost = base_cost * (1 + oil_change * oil_factor + fx_change * fx_factor + steel_change * steel_factor)
```

Example request:

```json
{
  "product_name": "Standard Industrial Component",
  "base_cost": 1000,
  "selling_price": 1500,
  "oil_change": 0.1,
  "fx_change": -0.03,
  "steel_change": 0.08
}
```

Change values are percentages expressed as decimals. For example, `0.1` means a 10% increase and `-0.03` means a 3% decrease.

The engine in `app/engine/simulator.py` accepts a pandas DataFrame and returns:

- `old_cost`
- `new_cost`
- `impact_amount`
- `gp_amount`
- `gp_percent`
