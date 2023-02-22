from typing import Any

import pytest
from openai.error import AuthenticationError, InvalidRequestError, OpenAIError

from sheet_ai.completion import create_completion
from sheet_ai.exceptions import CompletionError


def raise_exception() -> None:
    raise OpenAIError()


def raise_invalid_request_exception() -> None:
    raise InvalidRequestError("message", "param")


def raise_breaking_exception() -> None:
    raise AuthenticationError()


class MockCompletion:
    def __init__(self) -> None:
        class MockChoice:
            text = "completion"

        self.choices = [MockChoice()]


def return_completion() -> MockCompletion:
    return MockCompletion()


def test_retry_logic_with_eventual_success(monkeypatch: Any) -> None:
    responses = iter([raise_invalid_request_exception, raise_exception, return_completion])
    monkeypatch.setattr("sheet_ai.completion.openai.Completion.create", lambda **_: next(responses)())

    assert create_completion([]) == "completion"


def test_retry_logic_with_failure(monkeypatch: Any) -> None:
    responses = iter([raise_exception, raise_exception, raise_exception, return_completion])
    monkeypatch.setattr("sheet_ai.completion.openai.Completion.create", lambda **_: next(responses)())

    # with 2 retries, the response with completion is never reached
    with pytest.raises(CompletionError, match="Could not generate completion"):
        assert create_completion([], retries=2)


def test_retry_logic_with_breaking_exceptions(monkeypatch: Any) -> None:
    responses = iter([raise_exception, raise_breaking_exception, return_completion])
    monkeypatch.setattr("sheet_ai.completion.openai.Completion.create", lambda **_: next(responses)())

    # breaking exceptions don't result in a retry and are re-raised instead
    with pytest.raises(AuthenticationError):
        assert create_completion([])
