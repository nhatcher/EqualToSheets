import json
import os
from uuid import uuid4

from flask import Flask, abort, request, session
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address

from sheet_ai import db
from sheet_ai.exceptions import SheetAIError
from sheet_ai.workbook import WorkbookData, generate_workbook_data

MAX_SESSIONS_PER_IP = int(os.getenv("MAX_SESSIONS_PER_IP", 100))
MAX_PROMPTS_PER_SESSION = int(os.getenv("MAX_PROMPTS_PER_SESSION", 10))

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY")

limiter = Limiter(key_func=get_remote_address, app=app, storage_uri=db.MONGODB_URI)


@app.route("/ping", methods=["GET"])
def ping() -> str:
    return "OK"


@app.route("/session", methods=["GET"])
@limiter.limit(f"{MAX_SESSIONS_PER_IP}/hour")
def start_session() -> str:
    session["session_id"] = get_new_session_id()
    return "OK"


@app.route("/converse", methods=["POST"])
def converse() -> WorkbookData:
    session_id = session.get("session_id")
    if not session_id:
        raise abort(401, "Missing session cookie")

    check_rate_limit(session_id)

    try:
        prompt = json.loads(request.form["prompt"])
    except ValueError:
        raise abort(400, "Invalid POST data")

    workbook_data = db.get_prompt_response(prompt)

    if not workbook_data:
        try:
            workbook_data = generate_workbook_data(prompt)
        except SheetAIError:
            app.logger.exception("Workbook Not Found")
            raise abort(404, "Workbook Not Found")

    db.save_prompt_response(session_id, prompt, workbook_data)

    return workbook_data


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
