from flask import abort, request


def check_rate_limit() -> None:
    # TODO: Implement API rate limiter
    if request.form.get("error") == "rate-limit":
        abort(429)
