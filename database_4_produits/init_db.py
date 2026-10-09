"""Database Initializer & Helper for 4 Products Database
Creates SQLite products.db and exports products.json
"""

import json
import os
import sqlite3
from typing import Any, Dict, List, Optional

DB_FILE = os.path.join(os.path.dirname(__file__), "products.db")
SQL_FILE = os.path.join(os.path.dirname(__file__), "schema.sql")
JSON_FILE = os.path.join(os.path.dirname(__file__), "products.json")


def init_database():
    """Initializes SQLite database from schema.sql."""
    conn = sqlite3.connect(DB_FILE)
    with open(SQL_FILE, "r", encoding="utf-8") as f:
        schema_sql = f.read()
    conn.executescript(schema_sql)
    conn.commit()
    conn.close()
    print(f"[OK] SQLite database created at: {DB_FILE}")
    export_to_json()


def get_all_products() -> List[Dict[str, Any]]:
    """Fetch all 4 products along with their category, attributes, and images."""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    cursor.execute("""
        SELECT p.*, c.name as category_name, c.slug as category_slug
        FROM products p
        LEFT JOIN categories c ON p.category_id = c.id
        ORDER BY p.id ASC
    """)
    rows = cursor.fetchall()
    products = []

    for row in rows:
        prod = dict(row)
        pid = prod["id"]

        # Fetch attributes
        cursor.execute("SELECT attribute_name, attribute_value FROM product_attributes WHERE product_id = ?", (pid,))
        prod["attributes"] = {r["attribute_name"]: r["attribute_value"] for r in cursor.fetchall()}

        # Fetch images
        cursor.execute("SELECT image_url, alt_text, is_primary FROM product_images WHERE product_id = ?", (pid,))
        prod["images"] = [dict(r) for r in cursor.fetchall()]

        products.append(prod)

    conn.close()
    return products


def export_to_json():
    """Exports all products to products.json."""
    prods = get_all_products()
    with open(JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(prods, f, indent=2, ensure_ascii=False)
    print(f"[OK] JSON exported with {len(prods)} products at: {JSON_FILE}")


if __name__ == "__main__":
    init_database()
