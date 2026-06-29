# 🔗 ربط السيستم بـ OneDrive الحقيقي (Microsoft Graph API)

هذا الدليل يشرح كيفية تفعيل الرفع المباشر للشهادات الطبية على OneDrive
باستخدام **الإيميل والباسورد العادي** (بدون تحقق بخطوتين)، كما طلبتِ.

---

## ⚠️ خطوة مهمة لازم تتعمل أولاً من جانب IT الجامعة

مايكروسوفت بترفض تسجيل الدخول بالإيميل/الباسورد مباشرة من برنامج خارجي
**إلا لو** تم تسجيل "تطبيق" (App Registration) في Azure Active Directory للجامعة،
وتم تفعيل خيار اسمه **"Allow public client flows"**.

### الخطوات المطلوبة من مسؤول الـ IT بالجامعة (مرة واحدة فقط):

1. الدخول على: **portal.azure.com**
2. البحث عن: **Azure Active Directory** → **App registrations**
3. الضغط على **+ New registration**
   - الاسم: `Clinic System OneDrive Upload`
   - Supported account types: **Accounts in this organizational directory only**
   - Redirect URI: اتركيه فاضي
4. بعد إنشاء التطبيق، انسخي **Application (client) ID** — هو رقم طويل شكله:
   `a1b2c3d4-e5f6-7890-abcd-ef1234567890`
5. من القائمة الجانبية: **API permissions** → **+ Add a permission**
   → **Microsoft Graph** → **Delegated permissions**
   → ابحثي عن `Files.ReadWrite` وفعّليها
   → اضغطي **Grant admin consent**
6. من القائمة الجانبية: **Authentication**
   → في الأسفل تحت "Advanced settings"
   → فعّلي خيار: **"Allow public client flows"** → Yes
   → احفظي (Save)

بعد الخطوات دي، التطبيق يقدر يستخدم الإيميل والباسورد مباشرة بدون تحقق بخطوتين.

---

## الخطوات اللي تعمليها أنتِ

### 1. افتحي ملف `.env.example` الموجود في مجلد `clinic_system`
### 2. اعملي نسخة منه باسم `.env` (بدون .example) واملي القيم:

```
ONEDRIVE_EMAIL=studens_affairs@hti-o.edu.eg
ONEDRIVE_PASSWORD=كلمة_المرور_الحقيقية
GRAPH_CLIENT_ID=الكود_اللي_نسخته_من_خطوة_4_فوق
GRAPH_TENANT=common
```

### 3. شغّلي السيستم عادي
```bash
python app.py
```

السيستم هيقرأ ملف `.env` تلقائياً عند التشغيل.

---

## مكان حفظ الملفات على OneDrive

كل ملف هيترفع في المسار التالي على درايف حساب الشئون:
```
العياده / <كود الطالب> / <كود الطالب>_<اسم الطالب>_<تاريخ ووقت>.pdf
```

مثال:
```
العياده / 32025001 / 32025001_ابانوب_عايد_20260627_154802.pdf
```

---

## لو ظهر خطأ "فشل تسجيل الدخول لـ OneDrive"

الاحتمالات بالترتيب:

1. **الباسورد غلط** — تأكدي من نسخه صحيح في `.env`
2. **خطوة "Allow public client flows" لم تُفعّل** من جانب IT — رجعي للخطوة المطلوبة فوق
3. **صلاحية Files.ReadWrite لم تُمنح Admin Consent** — لازم Admin Consent من حساب لديه صلاحيات Azure Admin
4. **الحساب عنده تحقق بخطوتين بالفعل** — في هذه الحالة الطريقة دي مش هتنفع، ولازم نرجع لطريقة App Password بدل كده

---

## ملاحظة أمان

ملف `.env` فيه باسورد حقيقي — **لا** ترفعيه على أي مكان عام (GitHub، إيميل، الخ).
خليه على جهاز السيرفر فقط.
