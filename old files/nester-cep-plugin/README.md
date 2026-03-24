# Nester DTF Preview – Illustrator CEP Prototype

این یک پنل **CEP** مینیمال برای Adobe Illustrator است که فقط برای تست
workflow «پنل پایدار + Build Preview / Clear Preview» ساخته شده.

## ساختار فولدر

```text
nester-cep-plugin/
  manifest.xml
  CSXS/
    manifest.xml
  index.html
  styles.css
  panel.js
  jsx/
    host.jsx
  README.md
```

## کارهایی که پنل انجام می‌دهد

- UI ساده در یک Panel داخل Illustrator:
  - Style: Tight / Balanced / Production
  - Blocking: 1 / 2 / 3
  - Width Fill: اسلایدر 0–100
  - دکمه Build Preview
  - دکمه Clear Preview
  - ناحیه Status کوچک
- روی **Build Preview**:
  - انتخاب فعلی (selection) در Illustrator را می‌گیرد
  - اگر لایه‌ای به نام `NEST_PREVIEW` هست پاک می‌کند
  - یک لایه جدید `NEST_PREVIEW` می‌سازد
  - آبجکت‌های انتخاب‌شده را در این لایه Duplicate می‌کند
  - آن‌ها را روی آرت‌بورد اول به صورت **Grid ساده** می‌چیند
  - پنل باز می‌ماند و Status آپدیت می‌شود
- روی **Clear Preview**:
  - لایه `NEST_PREVIEW` را اگر وجود داشته باشد حذف می‌کند
  - پنل باز می‌ماند و Status آپدیت می‌شود

منطق چیدمان Grid در فایل `jsx/host.jsx` است و عمداً ساده نوشته شده تا بعداً
بتوانید آن را با موتور Nesting واقعی جایگزین کنید.

## نصب به عنوان CEP Extension در Illustrator (Windows)

1. **فعال‌کردن اجازه اجرای CEP های unsigned (اگر قبلاً نکرده‌اید)**  
   در ویندوز باید در رجیستری مقدار زیر را تنظیم کنید:

   - Run → `regedit`
   - مسیر:
     - `HKEY_CURRENT_USER/Software/Adobe/CSXS.9`
   - اگر کلید `PlayerDebugMode` وجود ندارد، یک `String Value` با این نام بسازید
     و مقدارش را روی `1` بگذارید.

   (شماره `CSXS.9` ممکن است بسته به نسخه Illustrator کمی فرق کند؛ اگر کار نکرد،
   نسخه‌های مثل `CSXS.8` یا `CSXS.10` را هم چک کنید.)

2. **کپی‌کردن فولدر Extension**

   فولدر `nester-cep-plugin` را (یا یک کپی از آن را) به مسیر Extensions ببرید، مثلاً:

   ```text
   %AppData%\Adobe\CEP\extensions\com.example.nester.dtf.panel
   ```

   مهم این است که داخل این فولدر، فایل‌های زیر باشند:

   ```text
   CSXS\manifest.xml
   index.html
   styles.css
   panel.js
   jsx\host.jsx
   ```

3. **اجرای Illustrator و باز کردن پنل**

   - Illustrator را اجرا کنید (یا دوباره بازش کنید تا Extension لود شود).
   - از منوی بالا به مسیر تقریبی زیر بروید:
     - `Window` → `Extensions` → `Nester DTF Preview`
   - با کلیک روی آن، پنل باز می‌شود و می‌توانید آن را Dock کنید (persistent panel).

4. **تست Workflow**

   - یک سند جدید بسازید و چند آبجکت روی آرت‌بورد رسم کنید.
   - چند آبجکت را انتخاب (Selection) کنید.
   - در پنل:
     - Style / Blocking / Width Fill را تنظیم کنید.
     - روی **Build Preview** کلیک کنید:
       - لایه `NEST_PREVIEW` ساخته می‌شود.
       - آبجکت‌های انتخاب‌شده در آن Duplicate می‌شوند.
       - این Duplicate ها روی آرت‌بورد اول به صورت Grid ساده چیده می‌شوند.
       - متن Status پایین پنل به چیزی مثل `Preview built successfully.` تغییر می‌کند.
     - روی **Clear Preview** کلیک کنید:
       - اگر `NEST_PREVIEW` وجود داشته باشد حذف می‌شود.
       - Status به `Preview cleared.` یا پیغام مشابه آپدیت می‌شود.

## جای پلاگ‌کردن موتور Nesting واقعی

- در حال حاضر منطق چیدمان در Extendscript این‌جاست:
  - `jsx/host.jsx` → تابع `_layoutAsGrid(...)`
- برای اتصال موتور Nesting واقعی:
  1. همین منطق مدیریت لایه و Duplicate کردن Selection را نگه دارید.
  2. بدنه `_layoutAsGrid` را با کد موتور Nesting خودتان عوض کنید:
     - Bounds آیتم‌ها را بخوانید
     - پوزیشن‌ها/چرخش نهایی را با الگوریتم Nesting حساب کنید
     - روی هر آیتم، Transform/Translate/Rotate مناسب را اعمال کنید

به این ترتیب فقط «مغز چیدمان» عوض می‌شود و بقیه workflow (پنل، دکمه‌ها،
ساخت/پاک کردن لایه Preview) سر جای خودشان می‌مانند.

