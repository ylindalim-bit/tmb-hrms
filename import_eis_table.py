# -*- coding: utf-8 -*-
"""One-time import of the official PERKESO EIS (Act 800) contribution table."""
import sqlite3

conn = sqlite3.connect("payroll.db")
conn.execute("""CREATE TABLE IF NOT EXISTS eis_table (
    id                 INTEGER PRIMARY KEY AUTOINCREMENT,
    wage_lower_bound   REAL NOT NULL,
    bracket_label      TEXT,
    contribution       REAL NOT NULL
)""")
conn.execute("DELETE FROM eis_table")

# (wage_lower_bound, bracket_label, contribution) - transcribed directly from
# https://www.perkeso.gov.my/images/dokumen/151124-Rate%20Contribution%20ACT%20800.pdf
ROWS = [
    (0, "Wages up to RM30", 0.05),
    (30.01, "Exceed RM30 but not exceed RM50", 0.10),
    (50.01, "Exceed RM50 but not exceed RM70", 0.15),
    (70.01, "Exceed RM70 but not exceed RM100", 0.20),
    (100.01, "Exceed RM100 but not exceed RM140", 0.25),
    (140.01, "Exceed RM140 but not exceed RM200", 0.35),
    (200.01, "Exceed RM200 but not exceed RM300", 0.50),
    (300.01, "Exceed RM300 but not exceed RM400", 0.70),
    (400.01, "Exceed RM400 but not exceed RM500", 0.90),
    (500.01, "Exceed RM500 but not exceed RM600", 1.10),
    (600.01, "Exceed RM600 but not exceed RM700", 1.30),
    (700.01, "Exceed RM700 but not exceed RM800", 1.50),
    (800.01, "Exceed RM800 but not exceed RM900", 1.70),
    (900.01, "Exceed RM900 but not exceed RM1,000", 1.90),
    (1000.01, "Exceed RM1,000 but not exceed RM1,100", 2.10),
    (1100.01, "Exceed RM1,100 but not exceed RM1,200", 2.30),
    (1200.01, "Exceed RM1,200 but not exceed RM1,300", 2.50),
    (1300.01, "Exceed RM1,300 but not exceed RM1,400", 2.70),
    (1400.01, "Exceed RM1,400 but not exceed RM1,500", 2.90),
    (1500.01, "Exceed RM1,500 but not exceed RM1,600", 3.10),
    (1600.01, "Exceed RM1,600 but not exceed RM1,700", 3.30),
    (1700.01, "Exceed RM1,700 but not exceed RM1,800", 3.50),
    (1800.01, "Exceed RM1,800 but not exceed RM1,900", 3.70),
    (1900.01, "Exceed RM1,900 but not exceed RM2,000", 3.90),
    (2000.01, "Exceed RM2,000 but not exceed RM2,100", 4.10),
    (2100.01, "Exceed RM2,100 but not exceed RM2,200", 4.30),
    (2200.01, "Exceed RM2,200 but not exceed RM2,300", 4.50),
    (2300.01, "Exceed RM2,300 but not exceed RM2,400", 4.70),
    (2400.01, "Exceed RM2,400 but not exceed RM2,500", 4.90),
    (2500.01, "Exceed RM2,500 but not exceed RM2,600", 5.10),
    (2600.01, "Exceed RM2,600 but not exceed RM2,700", 5.30),
    (2700.01, "Exceed RM2,700 but not exceed RM2,800", 5.50),
    (2800.01, "Exceed RM2,800 but not exceed RM2,900", 5.70),
    (2900.01, "Exceed RM2,900 but not exceed RM3,000", 5.90),
    (3000.01, "Exceed RM3,000 but not exceed RM3,100", 6.10),
    (3100.01, "Exceed RM3,100 but not exceed RM3,200", 6.30),
    (3200.01, "Exceed RM3,200 but not exceed RM3,300", 6.50),
    (3300.01, "Exceed RM3,300 but not exceed RM3,400", 6.70),
    (3400.01, "Exceed RM3,400 but not exceed RM3,500", 6.90),
    (3500.01, "Exceed RM3,500 but not exceed RM3,600", 7.10),
    (3600.01, "Exceed RM3,600 but not exceed RM3,700", 7.30),
    (3700.01, "Exceed RM3,700 but not exceed RM3,800", 7.50),
    (3800.01, "Exceed RM3,800 but not exceed RM3,900", 7.70),
    (3900.01, "Exceed RM3,900 but not exceed RM4,000", 7.90),
    (4000.01, "Exceed RM4,000 but not exceed RM4,100", 8.10),
    (4100.01, "Exceed RM4,100 but not exceed RM4,200", 8.30),
    (4200.01, "Exceed RM4,200 but not exceed RM4,300", 8.50),
    (4300.01, "Exceed RM4,300 but not exceed RM4,400", 8.70),
    (4400.01, "Exceed RM4,400 but not exceed RM4,500", 8.90),
    (4500.01, "Exceed RM4,500 but not exceed RM4,600", 9.10),
    (4600.01, "Exceed RM4,600 but not exceed RM4,700", 9.30),
    (4700.01, "Exceed RM4,700 but not exceed RM4,800", 9.50),
    (4800.01, "Exceed RM4,800 but not exceed RM4,900", 9.70),
    (4900.01, "Exceed RM4,900 but not exceed RM5,000", 9.90),
    (5000.01, "Exceed RM5,000 but not exceed RM5,100", 10.10),
    (5100.01, "Exceed RM5,100 but not exceed RM5,200", 10.30),
    (5200.01, "Exceed RM5,200 but not exceed RM5,300", 10.50),
    (5300.01, "Exceed RM5,300 but not exceed RM5,400", 10.70),
    (5400.01, "Exceed RM5,400 but not exceed RM5,500", 10.90),
    (5500.01, "Exceed RM5,500 but not exceed RM5,600", 11.10),
    (5600.01, "Exceed RM5,600 but not exceed RM5,700", 11.30),
    (5700.01, "Exceed RM5,700 but not exceed RM5,800", 11.50),
    (5800.01, "Exceed RM5,800 but not exceed RM5,900", 11.70),
    (5900.01, "Exceed RM5,900 but not exceed RM6,000", 11.90),
    (6000.01, "Exceed RM6,000", 11.90),
]

conn.executemany(
    "INSERT INTO eis_table (wage_lower_bound, bracket_label, contribution) VALUES (?,?,?)",
    ROWS,
)
conn.commit()
print(f"Imported {len(ROWS)} EIS brackets.")
