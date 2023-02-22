import os
from pathlib import Path
from time import sleep

import openai
from openai.error import AuthenticationError, InvalidRequestError, OpenAIError, PermissionError

from sheet_ai.exceptions import CompletionError

openai.api_key = os.getenv("OPENAI_API_KEY")

DEFAULT_RETRIES = 5

MODEL = "text-davinci-003"
TEMPERATURE = 0.5
MAX_TOKENS = 4000
PREAMBLE = (Path(__file__).resolve().parent / "preamble").read_text()


def create_completion(prompt: list[str], retries: int = DEFAULT_RETRIES) -> str:
    query = PREAMBLE + "\n".join(prompt)
    return _create_completion(query, MAX_TOKENS - len(query) // 4, retries)


def _create_completion(query: str, max_tokens: int, retries: int) -> str:
    try:
        return (
            openai.Completion.create(
                model=MODEL,
                prompt=query,
                temperature=TEMPERATURE,
                max_tokens=max_tokens,
            )
            .choices[0]
            .text.strip()
        )
    except InvalidRequestError:
        if retries > 0:
            # we can try lowering `max_tokens` value
            return _create_completion(query, max(200, max_tokens - 500), retries - 1)
        raise
    except (AuthenticationError, PermissionError):
        # there is no point in retrying when one of these exceptions is raised
        raise
    except OpenAIError:
        if retries > 0:
            sleep(0.5)
            return _create_completion(query, max_tokens, retries - 1)
        raise CompletionError("Could not generate completion")
