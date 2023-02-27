from typing import Any, Generator

import mongomock
import pytest
from flask.testing import FlaskClient

from sheet_ai.db import _get_db
from wsgi.app import app, limiter


@pytest.fixture
def client() -> Generator[FlaskClient, None, None]:
    app.config["TESTING"] = True
    app.secret_key = "flask secret key"

    with app.test_client() as client:
        yield client


@pytest.fixture(autouse=True)
def patch_mongo(monkeypatch: Any) -> Generator[None, None, None]:
    client: Any = mongomock.MongoClient()
    monkeypatch.setattr("sheet_ai.db._get_mongo_client", lambda: client)
    yield
    _get_db.cache_clear()


@pytest.fixture(autouse=True)
def reset_rate_limiter_storage() -> Generator[None, None, None]:
    yield
    limiter.reset()


@pytest.fixture(autouse=True)
def set_environ(monkeypatch: Any) -> None:
    monkeypatch.setenv("GIT_COMMIT", "af7b5cbe5fea4c7b90ba4541643cac087c2277cc")
