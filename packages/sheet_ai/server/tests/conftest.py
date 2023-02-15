from typing import Generator

import pytest
from flask.testing import FlaskClient

from wsgi.app import app


@pytest.fixture
def client() -> Generator[FlaskClient, None, None]:
    app.config["TESTING"] = True

    with app.test_client() as client:
        yield client
