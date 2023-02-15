import os
from pathlib import Path

import openai

openai.api_key = os.getenv("OPENAI_API_KEY")

TEMPERATURE = 0.5
MAX_TOKENS = 4000
PREAMBLE = (Path(__file__).resolve().parent / "preamble").read_text()


def create_completion(prompt: list[str]) -> str:
    # TODO: Retry logic
    query = PREAMBLE + "\n".join(prompt)
    completion = openai.Completion.create(
        model="text-davinci-003",
        prompt=query,
        temperature=TEMPERATURE,
        max_tokens=MAX_TOKENS - len(query) // 4,
    )
    return completion.choices[0].text.strip()
