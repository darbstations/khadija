# نشر منصة درب على الإنترنت (رابط ثابت)

الملفات جاهزة. اختر إحدى الطريقتين — كلاهما يعطيك **رابطاً عاماً دائماً**.

---

## الطريقة الأولى: Render (موصى بها — أسهل وبضغطات)

1. ادخل [render.com](https://render.com) وسجّل دخول بحساب GitHub.
2. **New +** ← **Blueprint**.
3. اختر مستودع `darbstations/khadija` وحدّد الفرع `claude/clever-wright-h6ccah`.
4. Render يقرأ ملف `render.yaml` تلقائياً ويجهّز الخدمة → اضغط **Apply / Deploy**.
5. بعد دقائق يظهر رابط مثل: `https://darb-kpi-platform.onrender.com` ✅

> الباقة المجانية تنام بعد خمول وتُوقظ عند الزيارة (تأخير أول طلب ~30 ثانية). للاستخدام الجاد اختر باقة مدفوعة.

## الطريقة الثانية: Railway

1. ادخل [railway.app](https://railway.app) ← **New Project** ← **Deploy from GitHub repo**.
2. اختر المستودع والفرع، واضبط **Root Directory = `platform`**.
3. Railway يبني تلقائياً (Procfile/Dockerfile موجودان) ← Generate Domain للحصول على الرابط.

## الطريقة الثالثة: أي خادم/سحابة (Docker)
```bash
cd platform
docker compose up -d        # المنفذ 8000
```

---

## بعد النشر (مهم)
1. **الإعدادات → Environment:** تأكد أن `DARB_SECRET` مضبوط (Render يولّده تلقائياً).
2. **غيّر كلمات المرور الافتراضية** من قاعدة البيانات أو بإعادة البذر بحسابات حقيقية.
3. **استمرارية البيانات (SQLite):** على الباقة المجانية تُعاد التهيئة عند كل نشر. للاستمرار:
   - Render: أضف **Disk** بحجم 1GB وmount على `/var/data` ثم اضبط `DARB_DB_PATH=/var/data/darb.db` (يتطلب باقة مدفوعة).
   - أو حوّل لقاعدة **PostgreSQL** (موصى به للإنتاج).

## التحقق
افتح الرابط ← سجّل دخول `admin / admin123` ← جرّب التغذية والتنقّل (موقع واحد تفاعلي كامل).
