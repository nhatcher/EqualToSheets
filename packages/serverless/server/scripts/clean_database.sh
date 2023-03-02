#!/bin/bash
set -euxo pipefail

# Creates a completely fresh local install of the application
docker-compose exec -T server python3 -- manage.py reset_db --noinput --close-sessions
docker-compose exec -T server python3 -- manage.py migrate
