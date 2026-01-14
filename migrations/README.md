# Migrations

This project doesn't use Alembic (that's Python). For schema changes, run the SQL in order against your Postgres DB.

Example:

```bash
psql "$DATABASE_URL" -f backend/migrations/001_drop_resume_json.sql
```

Add new `.sql` files here for future changes. Keep them numbered to preserve order.
