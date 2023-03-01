GRANT ALL PRIVILEGES ON DATABASE serverless TO serverless_admin;
ALTER ROLE serverless_admin SET client_encoding TO 'utf8';
ALTER ROLE serverless_admin SET default_transaction_isolation TO 'read committed';
ALTER ROLE serverless_admin SET default_transaction_isolation TO 'read committed';
ALTER ROLE serverless_admin SET timezone TO 'UTC';
ALTER USER serverless_admin WITH SUPERUSER;