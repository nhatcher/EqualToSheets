from typing import Any, Generator

import mongomock
import pytest
from flask.testing import FlaskClient

from wsgi.app import app


@pytest.fixture
def client() -> Generator[FlaskClient, None, None]:
    app.config["TESTING"] = True

    with app.test_client() as client:
        yield client


@pytest.fixture(autouse=True)
def patch_mongo(monkeypatch: Any) -> None:
    client: Any = mongomock.MongoClient()
    monkeypatch.setattr("sheet_ai.db._get_mongo_client", lambda: client)
