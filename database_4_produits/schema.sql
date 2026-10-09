-- Schema & Seed Data for 4 Products Database
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

-- ==========================================
-- 1. CATEGORIES (3 Categories)
-- ==========================================
INSERT INTO categories (id, name, slug) VALUES 
(1, 'Audio & Casques', 'audio-casques'),
(2, 'Ordinateurs & PC Portables', 'ordinateurs-laptops'),
(3, 'Périphériques & Écrans', 'peripheriques-ecrans');

-- ==========================================
-- 2. PRODUCTS (4 Diverse Products)
-- ==========================================

-- Produit 1: Sony WH-1000XM5
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
    'Casque premium circum-auriculaire avec réduction de bruit active et autonomie 30h.', 
    'Le Sony WH-1000XM5 dispose de deux processeurs contrôlant 8 microphones pour une réduction de bruit exceptionnelle. Compatible Hi-Res Audio, Bluetooth multipoint et charge rapide.', 
    349.99, 
    399.99, 
    210.00, 
    'USD', 
    45, 
    4.8, 
    1240, 
    1
);

-- Produit 2: Apple MacBook Pro 16" M3 Pro
INSERT INTO products (
    id, sku, name, brand, category_id, 
    short_description, description, 
    price, compare_at_price, cost_price, currency, 
    stock_quantity, rating, reviews_count, is_active
) VALUES (
    2, 
    'APPL-MBP16-M3PRO-512', 
    'Apple MacBook Pro 16" (Puce M3 Pro, 18Go RAM, 512Go SSD)', 
    'Apple', 
    2, 
    'Station de travail portable surpuissante avec écran Liquid Retina XDR 120Hz et puce M3 Pro.', 
    'Conçu pour les créateurs et développeurs exigeants. Écran Liquid Retina XDR de 16,2 pouces, puce M3 Pro avec CPU 12 cœurs et GPU 18 cœurs, autonomie record jusqu''à 22 heures.', 
    2499.00, 
    2699.00, 
    1850.00, 
    'USD', 
    18, 
    4.9, 
    620, 
    1
);

-- Produit 3: Souris Logitech MX Master 3S
INSERT INTO products (
    id, sku, name, brand, category_id, 
    short_description, description, 
    price, compare_at_price, cost_price, currency, 
    stock_quantity, rating, reviews_count, is_active
) VALUES (
    3, 
    'LOGI-MXM3S-GRY', 
    'Souris Ergonomique Logitech MX Master 3S Sans Fil', 
    'Logitech', 
    3, 
    'Souris professionnelle ergonomique avec capteur 8K DPI et clics silencieux.', 
    'L''icône de la productivité repensée: capteur optique 8 000 DPI qui fonctionne sur le verre, molette MagSpeed ultra-rapide (1000 lignes/sec), clics 90% plus discrets et autonomie de 70 jours.', 
    99.99, 
    119.99, 
    55.00, 
    'USD', 
    120, 
    4.7, 
    3450, 
    1
);

-- Produit 4: Écran Gaming Samsung Odyssey G9 OLED 49"
INSERT INTO products (
    id, sku, name, brand, category_id, 
    short_description, description, 
    price, compare_at_price, cost_price, currency, 
    stock_quantity, rating, reviews_count, is_active
) VALUES (
    4, 
    'SMSG-ODY-G9-OLED-49', 
    'Écran Gaming Incurvé Samsung Odyssey OLED G9 49" (240Hz, 0.03ms)', 
    'Samsung', 
    3, 
    'Écran géant ultra-large Dual QHD 32:9 incurvé 1800R avec dalle OLED Neo Quantum.', 
    'Une immersion totale: résolution Dual QHD (5120 x 1440), taux de rafraîchissement 240Hz, temps de réponse quasi-instantané de 0.03ms, compatible AMD FreeSync Premium Pro et design métallique ultra-fin.', 
    1299.99, 
    1799.99, 
    890.00, 
    'USD', 
    12, 
    4.6, 
    410, 
    1
);

-- ==========================================
-- 3. ATTRIBUTES
-- ==========================================

-- Attributes Produit 1 (Sony)
INSERT INTO product_attributes (product_id, attribute_name, attribute_value) VALUES 
(1, 'Couleur', 'Noir Mat'),
(1, 'Autonomie', '30 heures avec ANC'),
(1, 'Poids', '250g'),
(1, 'Connexion', 'Bluetooth 5.2 / Jack 3.5mm');

-- Attributes Produit 2 (MacBook Pro)
INSERT INTO product_attributes (product_id, attribute_name, attribute_value) VALUES 
(2, 'Processeur', 'Apple M3 Pro (12 cœurs CPU, 18 cœurs GPU)'),
(2, 'Mémoire RAM', '18 Go unifiée'),
(2, 'Stockage', '512 Go SSD NVMe'),
(2, 'Écran', '16.2" Liquid Retina XDR (3456 x 2234) ProMotion 120Hz'),
(2, 'Couleur', 'Noir Sidéral');

-- Attributes Produit 3 (Logitech MX Master 3S)
INSERT INTO product_attributes (product_id, attribute_name, attribute_value) VALUES 
(3, 'Capteur', 'Darkfield haute précision 8000 DPI'),
(3, 'Molette', 'MagSpeed électromagnétique en acier'),
(3, 'Autonomie', 'Jusqu''à 70 jours (USB-C rapide)'),
(3, 'Boutons', '7 boutons personnalisables via Logi Options+'),
(3, 'Couleur', 'Gris Graphite');

-- Attributes Produit 4 (Samsung Odyssey G9)
INSERT INTO product_attributes (product_id, attribute_name, attribute_value) VALUES 
(4, 'Taille écran', '49 pouces incurvé 1800R (Format 32:9)'),
(4, 'Dalle', 'OLED Neo Quantum'),
(4, 'Résolution', 'Dual QHD 5120 x 1440 pixels'),
(4, 'Fréquence', '240 Hz'),
(4, 'Temps de réponse', '0.03 ms (GTG)');

-- ==========================================
-- 4. IMAGES
-- ==========================================
INSERT INTO product_images (product_id, image_url, alt_text, is_primary) VALUES 
(1, 'https://images.unsplash.com/photo-1505740420928-5e560c06d30e', 'Sony WH-1000XM5', 1),
(2, 'https://images.unsplash.com/photo-1517336714731-489689fd1ca8', 'Apple MacBook Pro 16 M3', 1),
(3, 'https://images.unsplash.com/photo-1615663245857-ac93bb7c39e7', 'Logitech MX Master 3S', 1),
(4, 'https://images.unsplash.com/photo-1527443224154-c4a3942d3acf', 'Samsung Odyssey G9 Monitor', 1);
