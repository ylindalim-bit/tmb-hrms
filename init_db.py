# -*- coding: utf-8 -*-
"""
Creates a fresh, empty payroll.db from schema.sql + seed_reference_data.sql
(statutory rate tables, public holidays, leave-type lookups - no employee
or payroll data). Run this once on a new environment (e.g. right after
deploying) before using the app for the first time.

Safe to run only when payroll.db doesn't already exist - refuses to
overwrite real data.
"""
import os
import sqlite3

DB_PATH = os.path.join(os.path.dirname(__file__), "payroll.db")
SCHEMA_PATH = os.path.join(os.path.dirname(__file__), "schema.sql")
SEED_PATH = os.path.join(os.path.dirname(__file__), "seed_reference_data.sql")

if os.path.exists(DB_PATH):
    print(f"{DB_PATH} already exists - not touching it. Delete it first if you really want a fresh database.")
    raise SystemExit(1)

conn = sqlite3.connect(DB_PATH)
with open(SCHEMA_PATH, "r", encoding="utf-8") as f:
    conn.executescript(f.read())
with open(SEED_PATH, "r", encoding="utf-8") as f:
    conn.executescript(f.read())
conn.commit()
conn.close()
print(f"Created {DB_PATH} with an empty schema + reference/lookup data (no employees, no payroll).")
print("Next: log in isn't possible yet - you'll need an initial HR account. Ask Claude to create one, "
      "or add a row to hr_users directly (see app.py's hr_login for the expected fields).")
