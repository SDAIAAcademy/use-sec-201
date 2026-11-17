# الذكاء الاصطناعي في الأنظمة الصحية: المختبرات

أربعة مختبرات لمقرر أكاديمية سدايا **USE-SEC-201**، مدة كل منها 75 دقيقة، وعمل تحضيري متعقَّب قبل اليوم الأول. المختبرات
صفحات إرشادية بالعربية: الهدف، والملفات، وخطوات بدقائقها، وأسئلة تحقق، ومهمة إضافية ✱، وملاحظات للمدرّب. يعمل المشارك
في جدول بيانات إلكتروني ومستندات، ولا يكتب كودًا.

كل المختبرات في جهة واحدة: **شبكة أرياف الصحية الإقليمية**، شبكة افتراضية. المنتجات والمورِّدون والمرضى والملاحظات
افتراضية واصطناعية. الاستثناء الوحيد درجات الفرز في المختبر 1.2: درجات حقيقية أخرجها نموذج بحثي مفتوح على صور صدر عامة
(انظر `files/triage/ATTRIBUTION.txt`).

## ما يحتاجه المشارك

- جدول بيانات إلكتروني (Excel أو LibreOffice Calc أو Google Sheets). Orange Data Mining وKNIME بديلان لمن يفضّلهما.
- قارئ نصوص.
- المختبر 1.1 دون ذكاء اصطناعي. في المختبرات الأخرى يجوز مساعد عام بدور المعلّم (يشرح الحساب، ولا يحكم بدلًا منك)، ولا
  يُرفع إليه ملف مريض حقيقي.

## العمل التحضيري

نحو ساعتين قبل اليوم الأول: `files/prework/prework_guide.txt` و`files/prework/prework_inventory.xlsx`. حصر لخمس أدوات على
الأقل في خدمة المشارك، يفتحه المدرّب قبل الجلسة الأولى، ويُستخدم في المختبر 1.1.

## نقطة البداية والحل

كل مختبر يبدأ من ملفات جاهزة ولا يعتمد على مختبر آخر. الخريطة (`files/map/map_template.xlsx`) تكبر بسطر في آخر كل مختبر؛
من فاته مختبر يكتب سطره من بطاقة الحالة في `files/map/use_cases_14.txt`.

| المختبر | الصفحة | يبدأ من | الحل المرجعي |
|---|---|---|---|
| 1.1 | `D1_1_device_status.html` | `products/products_12.txt`، `products/intended_use_card.txt`، `products/classification_sheet.xlsx`، `prework/` | `solutions/classification_key.xlsx` |
| 1.2 | `D1_2_alert_volumes.html` | `triage/triage_scores.xlsx`، `triage/bariq_vendor_sheet.txt`، `triage/network_case_mix.txt` | `solutions/answer_key.md` |
| 2.1 | `D2_1_risk_score_audit.html` | `risk/patients_5000.xlsx`، `risk/awlawiya_datasheet.txt`، `risk/audit_note_template.txt` | `solutions/answer_key.md` |
| 2.2 | `D2_2_notes_and_plan.html` | `notes/notes_20.txt`، `notes/transcripts_20.txt`، `notes/scoring_sheet.xlsx`، `notes/vendor_claim.txt`، `notes/plan_template.txt` | `solutions/scoring_key.xlsx` |
| الخريطة | (المناقشة الختامية) | `map/map_template.xlsx`، `map/scoring_scale.txt`، `map/use_cases_14.txt` | `solutions/reference_map.xlsx` |

`solutions/answer_key.md` يجمع أرقام المختبرات الأربعة والخريطة وفخاخها.

## تشغيل الحلول

```bash
uv run labs/solutions/check.py            # كل المختبرات والخريطة
uv run labs/solutions/check.py 1.2 2.1    # بعضها
```

ملف واحد يعلن اعتمادياته، ويحتاج [uv](https://docs.astral.sh/uv/). يعيد حساب كل رقم في مفتاح الإجابات من الملفات التي
يتسلمها المشارك، ويطبع `PASS` لكل مختبر. لا يكتب شيئًا ولا يتصل بأي نموذج: لا خطوة في المختبرات تحتاج إلى نموذج.

## الملفات

| الملف | ما هو | المزروع فيه |
|---|---|---|
| `products/products_12.txt` | اثنا عشر وصف منتج كما يعرضها المورِّدون | خمسة فخاخ: إصدار غيّر الغرض، استخدام خارج الغرض، كلمات سريرية في أداة فوترة، «عافية» بغرض طبي، أداة عامة في استخدام طبي |
| `triage/triage_scores.xlsx` (و`.csv`) | 22,433 صورة: العمر والجنس ونوع الصورة والوسم ودرجة النموذج | حساسية أقل في الصور السريرية المتنقلة (AP)؛ ونشرة المورِّد تعلن «الدقة» |
| `triage/network_case_mix.txt` | توزيع الحالات في المواقع الثلاثة وطاقة القراءة | انتشار يختلف ثلاثين ضعفًا بين المواقع |
| `risk/patients_5000.xlsx` (و`.csv`) | 5,000 مريض اصطناعي بدرجة «أولوية» ودرجة بديلة | الدرجة تتعلم التكلفة، فتظلم سكان القرى النائية ومن يُعالَج خارج الشبكة |
| `notes/notes_20.txt`، `notes/transcripts_20.txt` | عشرون زيارة مكتوبة وملاحظاتها | ستة إغفالات وثلاث نتائج مختلقة، خمسة منها جسيمة؛ وادعاء توفير بلا مقارن |
| `map/use_cases_14.txt` | أربع عشرة بطاقة حالة استخدام | حالات عالية المخاطر بجاهزية بيانات ضعيفة |

## إعادة توليد الملفات

ملفات `files/` مولَّدة ببذرة ثابتة (201). النصوص في `_tools/texts.py`، والمولِّد:

```bash
uv run labs/_tools/prepare_files.py --nih-source <مجلد فيه preds.csv و nih_labels.csv>
```

`preds.csv` و`nih_labels.csv` من مستودع [reproduce-chexnet](https://github.com/jrzech/reproduce-chexnet) (رخصة BSD
3-Clause). دون `--nih-source` يُعاد استخدام `triage_scores.csv` الموجود. المحتوى نفسه في كل مرة؛ يتغير وقت الحفظ داخل
ملفات xlsx فقط.
