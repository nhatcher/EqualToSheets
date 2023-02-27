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


@pytest.fixture(autouse=True)
def override_settings(monkeypatch: Any) -> None:
    monkeypatch.setattr("wsgi.app.MAX_PROMPTS_PER_SESSION", 2)
    monkeypatch.setattr("wsgi.app.SUDO_PASSWORD", "letmein")


def test_converse(session_client: FlaskClient) -> None:
    response = session_client.post(
        "/converse",
        json={"prompt": ["Create budget table."]},
    )
    assert response.status_code == 200
    assert response.json == [[{"input": "Header", "style": {"bold": True}}]]


def test_converse_without_session(client: FlaskClient) -> None:
    response = client.post(
        "/converse",
        json={"prompt": ["Create budget table."]},
    )
    assert response.status_code == 401


def test_session_endpoint_rate_limit(client: FlaskClient) -> None:
    for _ in range(20):
        assert client.get("/session").status_code == 200

    assert client.get("/session").status_code == 429


def test_rate_limit_per_session(session_client: FlaskClient) -> None:
    # repeating the same prompt uses cache and doesn't affect the rate limit
    for _ in range(3):
        assert _get_converse_status_code(session_client, "prompt1") == 200

    assert _get_converse_status_code(session_client, "prompt2") == 200

    # the 3rd prompt should be declined
    assert _get_converse_status_code(session_client, "prompt3") == 429


@pytest.mark.parametrize("body", [{}, {"prompt": "not a list"}, {"prompt": 42}, {"prompt": [" "]}])
def test_converse_invalid_post_data(session_client: FlaskClient, body: dict[str, Any], monkeypatch: Any) -> None:
    monkeypatch.setattr("sheet_ai.workbook.create_completion", lambda _: "invalid")

    response = session_client.post("/converse", json=body)
    assert response.status_code == 400


def test_converse_could_not_generate_workbook(session_client: FlaskClient, monkeypatch: Any) -> None:
    monkeypatch.setattr("sheet_ai.workbook.create_completion", lambda _: "invalid")

    response = session_client.post(
        "/converse",
        json={"prompt": ["Create a table."]},
    )
    assert response.status_code == 404
    assert b"Workbook Not Found" in response.data


def test_sudo_session(client: FlaskClient) -> None:
    response = client.post("/sudo", json={"password": "letmein"})
    assert response.status_code == 200

    with client.session_transaction() as session:
        assert session["sudo"]

    for prompt_number in range(10):
        assert _get_converse_status_code(client, f"prompt{prompt_number}") == 200


@pytest.mark.parametrize("body", [{}, {"password": "incorrect"}])
def test_sudo_invalid_request(client: FlaskClient, body: dict[str, Any]) -> None:
    response = client.post("/sudo", json=body)
    assert response.status_code == 401


def _get_converse_status_code(client: FlaskClient, prompt: str) -> int:
    return client.post("/converse", json={"prompt": [prompt]}).status_code
