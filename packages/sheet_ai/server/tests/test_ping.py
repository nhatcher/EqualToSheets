from flask.testing import FlaskClient


def test_ping(client: FlaskClient) -> None:
    response = client.get("/ping")
    assert response.data == b"OK"
