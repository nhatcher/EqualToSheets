#!/bin/bash
set -euxo pipefail

# Creates a completely fresh local install of the application


# drop existing serverless database / user
sudo su -c "psql -d postgres -f scripts/sql/drop_database.sql" postgres
sudo su -c "psql -d postgres -f scripts/sql/drop_user.sql" postgres

# create new serverless database / user
sudo su -c "psql -d postgres -f scripts/sql/create_database.sql" postgres
sudo su -c "psql -d serverless -f scripts/sql/create_user.sql" postgres
sudo su -c "psql -d serverless -f scripts/sql/grant_permissions.sql" postgres

python manage.py migrate

