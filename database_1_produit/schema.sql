-- Schema & Seed Data for 1 Product Database (Idempotent)
PRAGMA foreign_keys = ON;

DROP TABLE IF EXISTS product_images;
DROP TABLE IF EXISTS product_attributes;
DROP TABLE IF EXISTS products;
DROP TABLE IF EXISTS categories;

CREATE TABLE categories (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE,
    slug TEXT NOT NULL UNIQUE
);

CREATE TABLE products (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    sku TEXT UNIQUE NOT NULL,
    name TEXT NOT NULL,
    brand TEXT NOT NULL,
    category_id INTEGER,
    short_description TEXT,
    description TEXT,
    price REAL NOT NULL,
    compare_at_price REAL,
    cost_price REAL,
    currency TEXT DEFAULT 'USD',
    stock_quantity INTEGER DEFAULT 0,
    rating REAL DEFAULT 0.0,
    reviews_count INTEGER DEFAULT 0,
    is_active INTEGER DEFAULT 1,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (category_id) REFERENCES categories (id)
);

CREATE TABLE product_attributes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    product_id INTEGER NOT NULL,
    attribute_name TEXT NOT NULL,
    attribute_value TEXT NOT NULL,
    FOREIGN KEY (product_id) REFERENCES products (id) ON DELETE CASCADE
);

CREATE TABLE product_images (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    product_id INTEGER NOT NULL,
    image_url TEXT NOT NULL,
    alt_text TEXT,
    is_primary INTEGER DEFAULT 0,
    FOREIGN KEY (product_id) REFERENCES products (id) ON DELETE CASCADE
);

-- Seed Data: 1 Category
INSERT INTO categories (id, name, slug) VALUES 
(1, 'Audio & Casques', 'audio-casques');

-- Seed Data: 1 Complete Product
INSERT INTO products (
    id, sku, name, brand, category_id, 
    short_description, description, 
    price, compare_at_price, cost_price, currency, 
    stock_quantity, rating, reviews_count, is_active
) VALUES (
    1, 
    'SONY-WH1000XM5-BLK', 
    'Casque Sans Fil Sony WH-1000XM5 Noise-Canceling', 
    'Sony', 
    1, 
    'Casque premium circum-auriculaire avec réduction de bruit active de pointe et autonomie de 30h.', 
    'Le Sony WH-1000XM5 redéfinit l''écoute sans distraction grâce à deux processeurs contrôlant huit microphones pour une réduction de bruit exceptionnelle. Compatible Hi-Res Audio sans fil, Bluetooth 5.2 multipoint et charge ultra rapide.', 
    349.99, 
    399.99, 
    210.00, 
    'USD', 
    45, 
    4.8, 
    1240, 
    1
);

-- Product Attributes
INSERT INTO product_attributes (product_id, attribute_name, attribute_value) VALUES 
(1, 'Couleur', 'Noir Mat'),
(1, 'Connectivité', 'Bluetooth 5.2 / Câble Jack 3.5mm'),
(1, 'Autonomie', '30 heures avec ANC (40h sans ANC)'),
(1, 'Temps de charge', '3.5 heures (3 min de charge = 3 heures d''écoute)'),
(1, 'Poids', '250g'),
(1, 'Microphones', '8 micros avec technologie Precise Voice Pickup'),
(1, 'Codec Audio', 'LDAC,等级, SBC');

-- Product Images
INSERT INTO product_images (product_id, image_url, alt_text, is_primary) VALUES 
(1, 'https://images.unsplash.com/photo-1505740420928-5e560c06d30e', 'Sony WH-1000XM5 Vue Frontale', 1),
(1, 'https://images.unsplash.com/photo-1484704849700-f032a568e944', 'Sony WH-1000XM5 Vue de Côté', 0);
