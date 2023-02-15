import json

from flask import Flask, abort, request

from sheet_ai.rate_limiter import check_rate_limit
from sheet_ai.workbook import WorkbookData, generate_workbook_data

app = Flask(__name__)


@app.route("/ping", methods=["GET"])
def ping() -> str:
    return "OK"


@app.route("/converse", methods=["POST"])
def converse() -> WorkbookData:
    check_rate_limit()

    try:
        prompt = json.loads(request.form["prompt"])
    except ValueError:
        raise abort(400, "Invalid POST data")

    return generate_workbook_data(prompt)
