from typing import Any

import pytest
from flask.testing import FlaskClient

from sheet_ai import db


def test_signup(client: FlaskClient) -> None:
    response = client.post("/signup", json={"email": "kamil@equalto.com"})
    assert response.status_code == 200
    assert db.get_email_addresses() == ["kamil@equalto.com"]

    response = client.post("/signup", json={"email": "tomek@equalto.com"})
    assert response.status_code == 200
    assert db.get_email_addresses() == ["kamil@equalto.com", "tomek@equalto.com"]

    # requesting /signup with an already saved email address is a noop
    response = client.post("/signup", json={"email": "kamil@equalto.com"})
    assert response.status_code == 200
    assert db.get_email_addresses() == ["kamil@equalto.com", "tomek@equalto.com"]


@pytest.mark.parametrize(
    "body",
    [
        {},
        {"email": "not an email address"},
        {"email": "ab@c@d"},
        {"email": "@a"},
        {"email": "a@"},
    ],
)
def test_signup_invalid_post_data(client: FlaskClient, body: dict[str, Any]) -> None:
    response = client.post("/signup", json=body)
    assert response.status_code == 400
    assert db.get_email_addresses() == []
