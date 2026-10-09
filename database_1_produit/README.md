# Database 1 Produit (Single Product Database)

قاعدة بيانات مخصصة لمنتج واحد بجميع التفاصيل (SQLite + JSON).

## المحتويات:
- `products.db`: قاعدة بيانات SQLite جاهزة تحتوي على جداول:
  - `products` (المنتج الأساسي مع الأسعار والمخزون والتقييم)
  - `categories` (الفئة)
  - `product_attributes` (المواصفات التقنية: اللون، البطارية، البلوتوث...)
  - `product_images` (الصور والروابط)
- `products.json`: ملف JSON كامل للمنتج لتسهيل القراءة المباشرة من LLMs أو APIs.
- `schema.sql`: كود SQL الأصلي لإنشاء الجداول وإدخال البيانات.
- `init_db.py`: سكربت بايثون لإعادة التهيئة والتعديل.

## المنتج المتوفر:
- **الاسم**: Casque Sans Fil Sony WH-1000XM5 Noise-Canceling
- **SKU**: SONY-WH1000XM5-BLK
- **السعر**: $349.99 (تخفيض من $399.99)
- **التقييم**: 4.8 / 5 (1,240 تقييم)
- **المخزون**: 45 قطعة
