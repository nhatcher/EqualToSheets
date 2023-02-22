from __future__ import annotations

import os
from datetime import datetime
from functools import cache
from hashlib import md5
from typing import Any

from pymongo import ASCENDING, IndexModel, MongoClient
from pymongo.database import Database

from sheet_ai.completion import MODEL, PREAMBLE, TEMPERATURE
from sheet_ai.workbook import WorkbookData

MONGODB_URI = os.getenv("MONGODB_URI")

COMPLETION_PARAMETERS = {
    "model": MODEL,
    "temperature": TEMPERATURE,
    "preamble_md5": md5(PREAMBLE.encode("utf-8"), usedforsecurity=False).hexdigest(),
}


@cache
def _get_db() -> Database[dict[str, Any]]:
    client = _get_mongo_client()
    db = client["sheet-ai"]

    # `create_index` is a NOOP if the index already exists
    db.prompt.create_indexes(
        [
            IndexModel([("session_id", ASCENDING)]),
            IndexModel([("prompt", ASCENDING), ("completion_parameters", ASCENDING)]),
        ],
    )

    return db


def _get_mongo_client() -> MongoClient[dict[str, Any]]:
    return MongoClient(MONGODB_URI)


def get_session_prompt_count(session_id: str) -> int:
    return _get_db().prompt.count_documents({"session_id": session_id})


def get_prompt_response(prompt: list[str]) -> WorkbookData | None:
    document = _get_db().prompt.find_one(
        {
            "prompt": prompt,
            # limit responses to completions using the current settings
            "completion_parameters": COMPLETION_PARAMETERS,
        },
    )
    if not document:
        return None
    return document["workbook"]


def save_prompt_response(session_id: str, prompt: list[str], workbook: WorkbookData) -> None:
    _get_db().prompt.update_one(
        filter={
            "session_id": session_id,
            "prompt": prompt,
            "workbook": workbook,
            "completion_parameters": COMPLETION_PARAMETERS,
        },
        update={"$setOnInsert": {"create_date": datetime.utcnow()}},
        upsert=True,
    )
