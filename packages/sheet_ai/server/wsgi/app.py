import os
from typing import Any
from uuid import uuid4

from flask import Flask, abort, request, session
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address

from sheet_ai import db
from sheet_ai.exceptions import EmailValidationError, SheetAIError
from sheet_ai.workbook import WorkbookData, generate_workbook_data

SESSION_RATE_LIMIT_POLICY = os.getenv("SESSION_RATE_LIMIT_POLICY", "20/day")
MAX_PROMPTS_PER_SESSION = int(os.getenv("MAX_PROMPTS_PER_SESSION", 10))

SUDO_PASSWORD = os.getenv("SUDO_PASSWORD")

app = Flask(__name__)
logger = app.logger
app.secret_key = os.getenv("FLASK_SECRET_KEY")

limiter = Limiter(key_func=get_remote_address, app=app, storage_uri=db.MONGODB_URI)


@app.route("/ping", methods=["GET"])
def ping() -> str:
    return "OK"


@app.route("/session", methods=["GET"])
@limiter.limit(SESSION_RATE_LIMIT_POLICY)
def start_session() -> str:
    session["session_id"] = get_new_session_id()
    return "OK"


@app.route("/sudo", methods=["POST"])
def start_sudo_session() -> str:
    assert SUDO_PASSWORD, "sudo access is not enabled"

    if _get_request_json().get("password") != SUDO_PASSWORD:
        abort(401)

    session["session_id"] = get_new_session_id()
    session["sudo"] = True
    return "OK"


@app.route("/converse", methods=["POST"])
def converse() -> WorkbookData:
    session_id = session.get("session_id")
    if not session_id:
        raise abort(401, "Missing session cookie")

    if not session.get("sudo"):
        check_rate_limit(session_id)

    prompt = _get_prompt()

    workbook_data = db.get_prompt_response(prompt)

    if not workbook_data:
        try:
            workbook_data = generate_workbook_data(prompt)
        except SheetAIError:
            raise abort(404, "Workbook Not Found")

    db.save_prompt_response(session_id, prompt, workbook_data)

    return workbook_data


def _get_prompt() -> list[str]:
    try:
        prompt = _get_request_json()["prompt"]
    except (KeyError, ValueError):
        raise abort(400)
    if not isinstance(prompt, list) or not all(isinstance(msg, str) for msg in prompt):
        raise abort(400)
    prompt = list(filter(None, map(str.strip, prompt)))
    if not prompt:
        raise abort(400)
    return prompt


@app.route("/signup", methods=["POST"])
def signup() -> str:
    try:
        db.save_email_address(_get_request_json()["email"])
    except (KeyError, EmailValidationError):
        raise abort(400, "Invalid POST data")
    return "OK"


def check_rate_limit(session_id: str) -> None:
    """
    Check the rate limits for given session.

    We only look at the number of associated prompts in the DB which means that repeating
    the same query or getting an error response doesn't count to the limit.
    """
    if db.get_session_prompt_count(session_id) >= MAX_PROMPTS_PER_SESSION:
        raise abort(429)


def get_new_session_id() -> str:
    return uuid4().hex


def _get_request_json() -> dict[str, Any]:
    json = request.json
    if json is None:
        raise abort(400)
    return json
