# تقرير ناصع اللايف — دليل التركيب التلقائي

الهدف: عند تحديث ملف الإكسل، يتحدّث التقرير الأخضر الحيّ **تلقائياً** على رابط ثابت يُرسل للإدارة.

## ما الذي ستحصل عليه
- رابط ثابت دائم للتقرير الحالي:
  `https://<github-username>.github.io/nasie-report/latest/`
- أرشيف تلقائي لكل تحديث داخل مجلد `archive/<التاريخ>/`.
- التقرير بنفس التصميم الأخضر والهوية والقواعد المتفق عليها.

## مكونات الحزمة
| الملف | الوظيفة |
|------|---------|
| `build_report.py` | المعالج: يقرأ الإكسل ويُخرج `index.html` (لوحة التقرير) |
| `logo.b64` | شعار ناصع مضمّن (لا تحذفه) |
| `requirements.txt` | مكتبة بايثون المطلوبة (`openpyxl`) |
| `.github/workflows/build.yml` | الأتمتة: يبني التقرير وينشره تلقائياً |
| `data/Nasie_Maintenance.xlsx` | ملف البيانات (يُستبدل تلقائياً من SharePoint) |

## كيف تعمل المنظومة (نظرة عامة للمبرمج)
```
SharePoint (Excel)  ──(Power Automate)──►  GitHub: data/Nasie_Maintenance.xlsx
                                                   │  (push)
                                                   ▼
                                        GitHub Actions (build.yml)
                                                   │  build_report.py
                                                   ▼
                                        latest/index.html  ──►  GitHub Pages (الرابط اللايف)
```

## خطوات التركيب لمرة واحدة (للمبرمج / IT)
راجع ملف `SETUP_FOR_DEVELOPER_EN.md` لتفاصيل كل خطوة بالإنجليزية. باختصار:
1. ارفع محتويات هذه الحزمة إلى مستودع GitHub باسم `nasie-report`.
2. فعّل **GitHub Pages** ← Source = GitHub Actions (أو فرع `main`، مجلد الجذر).
3. أنشئ تدفّق **Power Automate**: «عند تعديل ملف في SharePoint» ← تحديث ملف `data/Nasie_Maintenance.xlsx` في GitHub عبر GitHub API.
4. خلاص — أي تعديل على الإكسل يُحدّث الرابط خلال 1–2 دقيقة.

## ملاحظات مهمة
- **القواعد مطبّقة داخل `build_report.py`**: احتساب DONE و In Review فقط، استبعاد (Demo/Inquiry/Guest Experience/General Requests-Internal)، فصل الفني الداخلي/الخارجي، احتساب المواد من Material Cost فقط، الرافع غير المحدد = AI، توحيد التصنيفات.
- **المقارنة الأسبوعية تلقائية**: يحفظ المعالج ملخص كل فترة في `latest/prev_summary.json` ويستخدمه للمقارنة في الفترة التالية.
- **هيكل ملف الإكسل المتوقّع**: ورقة `Service Calls` (التواريخ والحالات) + ورقة `Sheet1` (تفصيل التكاليف). إذا تغيّر ترتيب الأعمدة في SharePoint، يحتاج تعديل بسيط في دالة `load_data`.

## تشغيل يدوي للاختبار (بدون أتمتة)
```bash
pip install -r requirements.txt
python build_report.py data/Nasie_Maintenance.xlsx latest
# افتح latest/index.html في المتصفح
```
