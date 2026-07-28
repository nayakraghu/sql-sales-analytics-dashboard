"""
generate_data.py
────────────────
Generates 10,247 realistic sales records and inserts them into MySQL.
Run this AFTER running database/schema_and_seed.sql

Requirements:
    pip install faker mysql-connector-python
"""

import random
import mysql.connector
from faker import Faker
from datetime import date, timedelta

fake = Faker()

# ── DB CONFIG ────────────────────────────────────────────────
DB_CONFIG = {
    "host":     "localhost",
    "port":     3306,
    "user":     "root",
    "password": "your_password",   # ← change this
    "database": "sales_analytics",
}

# ── REFERENCE DATA ───────────────────────────────────────────
SALESPERSON_IDS = [1, 2, 3, 4, 5, 6, 7, 8]
PRODUCT_IDS     = [1, 2, 3, 4, 5, 6, 7, 8]
REGIONS         = ["North", "South", "East", "West", "Central"]
STATUSES        = ["Completed"] * 90 + ["Pending"] * 7 + ["Cancelled"] * 3

PRODUCT_PRICES = {
    1: 2400.00,   # Enterprise Suite
    2:  850.00,   # Pro Laptop X1
    3:  230.00,   # Annual Support
    4:  320.00,   # Smart Monitor 4K
    5:  140.00,   # Cloud Storage Pro
    6:   75.00,   # Wireless Headset
    7:  580.00,   # Dev Toolkit License
    8: 1200.00,   # On-site Training
}

START_DATE = date(2023, 1, 1)
END_DATE   = date(2024, 12, 31)
TARGET     = 10247


def random_date(start: date, end: date) -> date:
    delta = (end - start).days
    return start + timedelta(days=random.randint(0, delta))


def generate_records(n: int) -> list[dict]:
    records = []
    existing_orders = set()

    for i in range(n):
        # Unique order ID
        while True:
            order_id = f"ORD-{random.randint(1000, 9999)}-{i}"
            if order_id not in existing_orders:
                existing_orders.add(order_id)
                break

        product_id  = random.choice(PRODUCT_IDS)
        unit_price  = PRODUCT_PRICES[product_id]
        quantity    = random.choices([1, 2, 3, 4, 5, 10], weights=[50,25,12,6,4,3])[0]
        discount    = random.choices([0, 5, 10, 15, 20], weights=[60,20,12,5,3])[0]

        records.append({
            "order_id":       order_id,
            "salesperson_id": random.choice(SALESPERSON_IDS),
            "product_id":     product_id,
            "quantity":       quantity,
            "unit_price":     unit_price,
            "discount_pct":   discount,
            "order_date":     random_date(START_DATE, END_DATE),
            "region":         random.choice(REGIONS),
            "status":         random.choice(STATUSES),
        })

    return records


def insert_records(records: list[dict]) -> None:
    conn = mysql.connector.connect(**DB_CONFIG)
    cur  = conn.cursor()

    sql = """
        INSERT IGNORE INTO sales
            (order_id, salesperson_id, product_id, quantity,
             unit_price, discount_pct, order_date, region, status)
        VALUES
            (%(order_id)s, %(salesperson_id)s, %(product_id)s, %(quantity)s,
             %(unit_price)s, %(discount_pct)s, %(order_date)s, %(region)s, %(status)s)
    """

    batch_size = 500
    inserted   = 0

    for i in range(0, len(records), batch_size):
        batch = records[i : i + batch_size]
        cur.executemany(sql, batch)
        conn.commit()
        inserted += len(batch)
        print(f"  ✓ Inserted {inserted}/{len(records)} records...")

    cur.close()
    conn.close()
    print(f"\n✅ Done — {inserted} records inserted into sales_analytics.sales")


if __name__ == "__main__":
    print("🔧 Generating sales data...")
    records = generate_records(TARGET)
    print(f"✓ {len(records)} records generated")

    print("\n📤 Inserting into MySQL...")
    insert_records(records)
