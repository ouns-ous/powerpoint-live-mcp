# Database 4 Produits (4 Products Database)

قاعدة بيانات متكاملة تحتوي على 4 منتجات متنوعة مع الفئات، الخصائص التقنية، والأسعار (SQLite + JSON).

## المحتويات:
- `products.db`: قاعدة بيانات SQLite جاهزة تحتوي على جداول:
  - `categories` (فئات المنتجات: Audio, Laptops, Périphériques)
  - `products` (المنتجات الـ 4 مع الأسعار، المخزون، والتقييمات)
  - `product_attributes` (المواصفات التقنية لكل منتج)
  - `product_images` (الصور والروابط)
- `products.json`: ملف JSON منسق يحتوي على المنتجات الـ 4 كاملة.
- `schema.sql`: ملف SQL الأصلي لإنشاء الجداول وإدخال البيانات.
- `init_db.py`: سكربت بايثون لإعادة التهيئة والاستعلام.

## المنتجات المتوفرة:
1. **Casque Sony WH-1000XM5** (Audio) - $349.99 (تقييم: 4.8★ | مخزون: 45)
2. **Apple MacBook Pro 16" M3 Pro** (PC Portables) - $2,499.00 (تقييم: 4.9★ | مخزون: 18)
3. **Souris Logitech MX Master 3S** (Périphériques) - $99.99 (تقييم: 4.7★ | مخزون: 120)
4. **Écran Samsung Odyssey OLED G9 49"** (Gaming) - $1,299.99 (تقييم: 4.6★ | مخزون: 12)
