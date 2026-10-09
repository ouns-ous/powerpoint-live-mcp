"""Database Initializer & Helper for 1 Product Database
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


def get_product(product_id: int = 1) -> Optional[Dict[str, Any]]:
    """Fetch product by ID along with its category, attributes, and images."""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    cursor.execute("""
        SELECT p.*, c.name as category_name, c.slug as category_slug
        FROM products p
        LEFT JOIN categories c ON p.category_id = c.id
        WHERE p.id = ?
    """, (product_id,))
    row = cursor.fetchone()
    if not row:
        conn.close()
        return None

    prod = dict(row)

    # Fetch attributes
    cursor.execute("SELECT attribute_name, attribute_value FROM product_attributes WHERE product_id = ?", (product_id,))
    prod["attributes"] = {r["attribute_name"]: r["attribute_value"] for r in cursor.fetchall()}

    # Fetch images
    cursor.execute("SELECT image_url, alt_text, is_primary FROM product_images WHERE product_id = ?", (product_id,))
    prod["images"] = [dict(r) for r in cursor.fetchall()]

    conn.close()
    return prod


def export_to_json():
    """Exports the product database to products.json."""
    prod = get_product(1)
    if prod:
        with open(JSON_FILE, "w", encoding="utf-8") as f:
            json.dump([prod], f, indent=2, ensure_ascii=False)
        print(f"[OK] JSON exported at: {JSON_FILE}")


if __name__ == "__main__":
    init_database()
