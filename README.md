# Image Processing (Luxbaz)

نیازمندی: Python 3.10 یا بالاتر

```
pip install -r requirements.txt
```

## اجرا
`run.bat` یا:

```
python process_images.py convert   # گزینه ۱
python process_images.py process   # گزینه ۲
```

### گزینه ۱ - convert
همه‌ی عکس‌های `Product images` را به WebP زیر ۱۰۰ کیلوبایت تبدیل می‌کند و در `converted/` با همان نام و پوشه ذخیره می‌کند.

### گزینه ۲ - process
1. عکس‌ها را از `Product images photoroom/<کد محصول>/` می‌خواند.
2. کد پوشه را با ستون «کد محصول» در Product Database تطبیق می‌دهد.
3. از ستون «رنگ» (مثلاً `آبی ۱ | قرمز ۲`) اعداد را حذف می‌کند و با `color_mapping.xlsx` معادل انگلیسی را پیدا می‌کند.
4. محصول تک‌رنگ: همان رنگ. محصول چندرنگ: رنگ عکس تشخیص داده می‌شود و نزدیک‌ترین رنگ مجاز محصول انتخاب می‌شود.
5. عکس بدون تبدیل، با نام `کد-رنگ-01` یا `کد-رنگ-main`، در `final image renamed/<کد محصول>/` کپی می‌شود.

خروجی‌ها: `final image renamed/manifest.json` (جلوگیری از پردازش دوباره) و `color.xlsx` (شیت `detected` و شیت `missing` برای رنگ‌ها و کدهای مشکل‌دار).
اگر فایلی از خروجی پاک شود، در اجرای بعدی دوباره ساخته می‌شود.

## تنظیمات
- `settings.json`: تنظیمات مشترک (در گیت).
- `settings.local.json`: مسیرهای مخصوص این کامپیوتر (مثلاً `product_db` و `color_map`). روی `settings.json` اعمال می‌شود و در گیت نیست.
- با متغیر محیطی `IMAGE_SETTINGS` می‌توان فایل تنظیمات دیگری داد.

## تست
```
pip install -r requirements-dev.txt
python -m pytest
```
