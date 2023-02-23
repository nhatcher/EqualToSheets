import json
from typing import Any

import pytest
from flask.testing import FlaskClient


@pytest.fixture(name="session_client")
def fixture_session_client(client: FlaskClient) -> FlaskClient:
    client.get("/session")
    return client


@pytest.fixture(autouse=True)
def patch_openai_call(monkeypatch: Any) -> None:
    monkeypatch.setattr("sheet_ai.workbook.create_completion", lambda _: "|**Header**|")


def test_converse(session_client: FlaskClient) -> None:
    response = session_client.post(
        "/converse",
        data={"prompt": json.dumps(["Create budget table."])},
    )
    assert response.status_code == 200
    assert json.loads(response.data) == [[{"input": "Header", "style": {"bold": True}}]]


def test_converse_without_session(client: FlaskClient) -> None:
    response = client.post(
        "/converse",
        data={"prompt": json.dumps(["Create budget table."])},
    )
    assert response.status_code == 401


def test_session_endpoint_rate_limit(client: FlaskClient) -> None:
    for _ in range(20):
        assert client.get("/session").status_code == 200

    assert client.get("/session").status_code == 429


def test_rate_limit_per_session(session_client: FlaskClient, monkeypatch: Any) -> None:
    monkeypatch.setattr("wsgi.app.MAX_PROMPTS_PER_SESSION", 2)

    def get_response_status_code(prompt: str) -> int:
        return session_client.post(
            "/converse",
            data={"prompt": json.dumps([prompt])},
        ).status_code

    # repeating the same prompt uses cache and doesn't affect the rate limit
    for _ in range(3):
        assert get_response_status_code("prompt1") == 200

    assert get_response_status_code("prompt2") == 200

    # the 3rd prompt should be declined
    assert get_response_status_code("prompt3") == 429


def test_converse_invalid_post_data(session_client: FlaskClient, monkeypatch: Any) -> None:
    monkeypatch.setattr("sheet_ai.workbook.create_completion", lambda _: "invalid")

    response = session_client.post(
        "/converse",
        data={"prompt": "not a json"},
    )
    assert response.status_code == 400
    assert b"Invalid POST data" in response.data


def test_converse_could_not_generate_workbook(session_client: FlaskClient, monkeypatch: Any) -> None:
    monkeypatch.setattr("sheet_ai.workbook.create_completion", lambda _: "invalid")

    response = session_client.post(
        "/converse",
        data={"prompt": json.dumps(["Create a table."])},
    )
    assert response.status_code == 404
    assert b"Workbook Not Found" in response.data
