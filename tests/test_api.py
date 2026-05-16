from fastapi.testclient import TestClient

from app.api.main import app


def test_health_check() -> None:
    with TestClient(app) as client:
        response = client.get("/health")

    assert response.status_code == 200
    assert response.json() == {"status": "ok"}


def test_create_simulation() -> None:
    with TestClient(app) as client:
        response = client.post(
            "/simulations",
            json={
                "product_name": "API Product",
                "base_cost": 1000,
                "selling_price": 1500,
                "oil_change": 0.10,
                "fx_change": 0,
                "steel_change": 0,
                "oil_factor": 0.20,
            },
        )

    assert response.status_code == 200
    data = response.json()
    assert data["new_cost"] == 1020
    assert data["impact_amount"] == 20
