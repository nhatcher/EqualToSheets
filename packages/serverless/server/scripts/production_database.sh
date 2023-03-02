#!/bin/bash
set -euxo pipefail

# TODO: Update this script so that it works properly in the container

# Restores database from production locally, running local migrations

# drop existing serverless database / user
sudo su -c "psql -d postgres -f scripts/sql/drop_database.sql" postgres

# download database form server
heroku pg:pull postgresql-symmetrical-99140 serverless --app serverless-sheets

sudo su -c "psql -d serverless -f scripts/sql/grant_permissions.sql" postgres

python manage.py migrate
