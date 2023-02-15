import os

bind = f"0.0.0.0:{os.getenv('PORT')}"
workers = 7
timeout = 600
reload = os.environ.get("DEBUG", "") == "True"
disable_redirect_access_to_syslog = True
max_requests = 1000
max_requests_jitter = 20
