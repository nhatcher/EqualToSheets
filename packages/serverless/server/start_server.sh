#!/bin/bash

set -eu

python3 manage.py migrate
python3 manage.py collectstatic --noinput

exec gunicorn --config server/gunicorn_config.py server.wsgi:application
