# /// script
# requires-python = ">=3.11"
# dependencies = ["openpyxl>=3.1", "numpy>=1.26"]
# ///
"""Writes the file pack of the Aryaf case into labs/files/ and the keys into labs/solutions/.

    uv run labs/_tools/prepare_files.py
    uv run labs/_tools/prepare_files.py --nih-source <folder with preds.csv and nih_labels.csv>

Fixed seed: the same content on every run. Without --nih-source the committed
labs/files/triage/triage_scores.csv is kept as it is (it is derived from two public files of
github.com/jrzech/reproduce-chexnet, see labs/files/triage/ATTRIBUTION.txt).
"""
import argparse
import csv
import datetime
import json
import math
import random
import sys
from pathlib import Path

import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

sys.path.insert(0, str(Path(__file__).parent))
import texts as T  # noqa: E402

SEED = 201
LABS = Path(__file__).resolve().parents[1]
FILES = LABS / "files"
SOL = LABS / "solutions"
FIXED = datetime.datetime(2026, 11, 1, 9, 0, 0)
HEAD = PatternFill("solid", fgColor="DCEBEA")
FILL = PatternFill("solid", fgColor="FFF6DA")


def text(path, body):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(body.strip() + "\n", encoding="utf-8", newline="\n")


def write_csv(path, header, rows):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, lineterminator="\n")
        w.writerow(header)
        w.writerows(rows)


def book():
    wb = Workbook()
    wb.properties.created = FIXED
    wb.properties.modified = FIXED
    wb.properties.creator = "SDAIA Academy"
    return wb


def sheet(wb, title, header, rows, widths=None, first=False, fill_cols=()):
    ws = wb.active if first else wb.create_sheet()
    ws.title = title
    ws.sheet_view.rightToLeft = True
    ws.append(header)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = HEAD
        c.alignment = Alignment(wrap_text=True, vertical="top")
    for r in rows:
        ws.append(r)
    for i, w in enumerate(widths or [], start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    for col in fill_cols:
        for row in ws.iter_rows(min_row=2, min_col=col, max_col=col):
            row[0].fill = FILL
    ws.freeze_panes = "A2"
    return ws


def save(wb, path):
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


# --------------------------------------------------------------------------------- pre-work
def prework():
    text(FILES / "prework" / "prework_guide.txt", """
العمل التحضيري · حصر أدوات الذكاء الاصطناعي في خدمتك (نحو ساعتين، قبل اليوم الأول)

ما تسلّمه: الملف prework_inventory.xlsx وفيه خمس أدوات على الأقل، يُرفع إلى مجلدك قبل الجلسة الأولى.
يُتعقَّب التسليم: المدرّب يفتح الملفات قبل اليوم الأول ويبني عليها افتتاح الجلسة.

1. اسأل ثلاثة أشخاص في خدمتك (طبيب، ممرض، إداري): أي برنامج عندنا «يتنبأ» أو «يقترح» أو «يكتب» أو «يرتّب»؟
   كثير من الأدوات لا يحمل اسم الذكاء الاصطناعي: درجة خطر في الملف الإلكتروني، ترتيب قائمة، تنبيه دوائي.
2. لكل أداة اكتب سطرًا في الجدول: ماذا تفعل، ومن يستخدمها، وماذا يقول المورِّد عن غرضها، ومن يملكها عندكم.
3. اترك عمود «جهاز طبي؟» فارغًا: تملؤه في المختبر 1.1.
4. اكتب تحت الجدول: أداة واحدة لا تعرف من يملكها، وأداة واحدة لا تعرف على أي بيانات اختُبرت.

قواعد: لا اسم مريض ولا رقم ملف ولا لقطة شاشة فيها بيانات مريض. اسم الأداة ووصفها يكفيان.
إن لم تجد خمس أدوات في خدمتك فأكمل من أدوات تعرفها في جهة صحية أخرى، واكتب ذلك.
""")
    wb = book()
    sheet(wb, "حصر الأدوات",
          ["#", "اسم الأداة", "ماذا تفعل", "من يستخدمها", "الغرض كما يقوله المورِّد", "المالك عندنا",
           "على أي بيانات اختُبرت؟", "جهاز طبي؟ (يُملأ في المختبر 1.1)"],
          [[i, "", "", "", "", "", "", ""] for i in range(1, 9)],
          [5, 22, 34, 22, 34, 20, 26, 22], first=True, fill_cols=(8,))
    save(wb, FILES / "prework" / "prework_inventory.xlsx")


# ---------------------------------------------------------------------------------- lab 1.1
def lab11():
    lines = ["اثنا عشر منتجًا معروضًا على شبكة أرياف الصحية · كل الأسماء والمورِّدين افتراضية", ""]
    for pid, name, desc, *_ in T.PRODUCTS:
        lines += [f"{pid} · {name}", desc, ""]
    text(FILES / "products" / "products_12.txt", "\n".join(lines))
    text(FILES / "products" / "intended_use_card.txt", """
بطاقة الغرض المقصود · شبكة أرياف الصحية، الهندسة الطبية والشؤون التنظيمية

السؤال الأول: ما الغرض الذي أعلنه المورِّد؟ (المواصفات، وتعليمات الاستخدام، وكل ما يقوله عن المنتج)
السؤال الثاني: هل هذا الغرض واحد مما يأتي لحالة مرضية أو لمريض؟
    فحص · كشف · تشخيص · مراقبة · علاج · تدبير حالة
    نعم  ←  يُعامل جهازًا طبيًا ويخضع لرقابة الهيئة العامة للغذاء والدواء.
    لا   ←  ليس جهازًا طبيًا بهذا الغرض.
السؤال الثالث: هل نستخدمه نحن في الغرض المعلن نفسه؟
    إن استخدمناه في غير غرضه فالمسؤولية علينا، والحكم يُعاد.
السؤال الرابع: هل تغيّر الإصدار؟ إصدار يضيف وظيفة قد يغيّر الغرض، فيتغير الحكم.

هذه البطاقة تلخيص تعليمي بكلمات المقرر لقاعدة الغرض المقصود في إرشادات الهيئة (MDS-G010).
الحكم النظامي في أي منتج حقيقي للهيئة، ويُسأل عنه مسؤول الشؤون التنظيمية.
""")
    wb = book()
    sheet(wb, "التصنيف",
          ["المنتج", "الاسم", "الحكم: جهاز طبي / ليس جهازًا طبيًا", "الكلمة الدالة في الغرض المقصود",
           "ما الذي يتطلبه منّا قبل التشغيل؟", "ملاحظة: استخدام خارج الغرض أو تغيّر إصدار؟"],
          [[p[0], p[1], "", "", "", ""] for p in T.PRODUCTS], [9, 22, 26, 30, 44, 34], first=True,
          fill_cols=(3, 4, 5, 6))
    save(wb, FILES / "products" / "classification_sheet.xlsx")
    wb = book()
    sheet(wb, "مفتاح التصنيف",
          ["المنتج", "الاسم", "الحكم", "الكلمة الدالة", "ما يتطلبه", "الفخ"],
          [[p[0], p[1], p[3], p[4], T.REQUIRES_DEVICE if p[3] == T.DEVICE else T.REQUIRES_OTHER, p[5]]
           for p in T.PRODUCTS], [9, 22, 18, 30, 60, 44], first=True)
    save(wb, SOL / "classification_key.xlsx")
    return {"devices": [p[0] for p in T.PRODUCTS if p[3] == T.DEVICE],
            "not_devices": [p[0] for p in T.PRODUCTS if p[3] != T.DEVICE],
            "traps": [p[0] for p in T.PRODUCTS if p[5]]}


# ---------------------------------------------------------------------------------- lab 1.2
SITES = [  # name, chest films a day, share of bedside (AP) films, pneumothorax prevalence
    ("الطوارئ والعناية في مستشفى أرياف المركزي", 180, 0.85, 0.030),
    ("العيادات الخارجية في مستشفى السهل العام", 220, 0.05, 0.008),
    ("مراكز الرعاية الأولية (فحوص دورية وما قبل التوظيف)", 400, 0.00, 0.001),
]
THRESHOLD = 0.05
CAPACITY = 30  # alerts a day the emergency radiologist can read at once


def derive_triage(source):
    labels = {r["Image Index"]: r for r in csv.DictReader(open(source / "nih_labels.csv", encoding="utf-8"))}
    rows = []
    for r in csv.DictReader(open(source / "preds.csv", encoding="utf-8")):
        l = labels[r["Image Index"]]
        rows.append([r["Image Index"], int(l["Patient ID"]), int(l["Patient Age"].rstrip("YMD")),
                     l["Patient Gender"], l["View Position"], int(l["Pneumothorax"]),
                     f'{float(r["prob_Pneumothorax"]):.4f}'])
    rows.sort(key=lambda x: x[0])
    with open(FILES / "triage" / "triage_scores.csv", "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f, lineterminator="\n")
        w.writerow(["image_id", "patient_id", "age", "sex", "view", "pneumothorax", "score"])
        w.writerows(rows)


def counts(rows, t):
    tp = sum(1 for r in rows if r[5] and r[6] >= t)
    fn = sum(1 for r in rows if r[5] and r[6] < t)
    fp = sum(1 for r in rows if not r[5] and r[6] >= t)
    tn = sum(1 for r in rows if not r[5] and r[6] < t)
    return tp, fn, fp, tn


def lab12(source):
    (FILES / "triage").mkdir(parents=True, exist_ok=True)
    if source:
        derive_triage(Path(source))
    rows = []
    for r in csv.DictReader(open(FILES / "triage" / "triage_scores.csv", encoding="utf-8")):
        rows.append([r["image_id"], int(r["patient_id"]), int(r["age"]), r["sex"], r["view"],
                     int(r["pneumothorax"]), float(r["score"])])
    tp, fn, fp, tn = counts(rows, THRESHOLD)
    sens, spec = tp / (tp + fn), tn / (tn + fp)
    acc = (tp + tn) / len(rows)
    sens_r, spec_r = round(sens, 2), round(spec, 2)

    wb = book()
    sheet(wb, "البيانات", ["image_id", "patient_id", "age", "sex", "view", "pneumothorax", "score"], rows,
          [20, 12, 8, 8, 8, 14, 10], first=True)
    calc = wb.create_sheet("الحساب")
    calc.sheet_view.rightToLeft = True
    n = len(rows) + 1
    D = f"'البيانات'!"
    lines = [
        ["الجزء أ · عدّ من البيانات عند العتبة", "العتبة", "تنبيه صحيح", "حالة فائتة", "تنبيه كاذب", "سليم بلا تنبيه", "الحساسية", "النوعية"],
        ["مثال محلول", THRESHOLD,
         f'=COUNTIFS({D}$F$2:$F${n},1,{D}$G$2:$G${n},">="&B2)',
         f'=COUNTIFS({D}$F$2:$F${n},1,{D}$G$2:$G${n},"<"&B2)',
         f'=COUNTIFS({D}$F$2:$F${n},0,{D}$G$2:$G${n},">="&B2)',
         f'=COUNTIFS({D}$F$2:$F${n},0,{D}$G$2:$G${n},"<"&B2)',
         "=C2/(C2+D2)", "=F2/(E2+F2)"],
        ["الصور السريرية المتنقلة (AP) عند العتبة نفسها", THRESHOLD, "", "", "", "", "", ""],
        ["الصور القياسية (PA) عند العتبة نفسها", THRESHOLD, "", "", "", "", "", ""],
        ["عتبة 0.02", 0.02, "", "", "", "", "", ""],
        ["عتبة 0.10", 0.10, "", "", "", "", "", ""],
        ["عتبة 0.20", 0.20, "", "", "", "", "", ""],
        [],
        ["الجزء ب · لكل موقع، لشهر من 30 يومًا", "صور في اليوم", "معدل الانتشار", "الحساسية", "النوعية",
         "تنبيهات في اليوم", "القيمة التنبؤية الإيجابية", "حالات فائتة في الشهر"],
    ] + [[s[0], s[1], s[3], "", "", "", "", ""] for s in SITES]
    for l in lines:
        calc.append(l)
    for r in (1, 9):
        for c in calc[r]:
            c.font = Font(bold=True)
            c.fill = HEAD
            c.alignment = Alignment(wrap_text=True, vertical="top")
    for r in list(range(3, 8)) + [10, 11, 12]:
        for c in calc[r][2 if r < 9 else 3:8]:
            c.fill = FILL
    for col, w in zip("ABCDEFGH", [46, 12, 14, 14, 14, 16, 18, 18]):
        calc.column_dimensions[col].width = w
    save(wb, FILES / "triage" / "triage_scores.xlsx")

    text(FILES / "triage" / "bariq_vendor_sheet.txt", f"""
«بارق» · فرز صور الصدر لاسترواح الصدر · نشرة المورِّد (مورِّد افتراضي)

«دقة {acc:.0%} في كشف استرواح الصدر»

- الحساسية {sens_r:.0%}، والنوعية {spec_r:.0%}، عند العتبة الموصى بها {THRESHOLD}.
- قيست على مجموعة اختبار عامة من {len(rows):,} صورة صدر.
- «يقرأ الصورة في ثوانٍ، ويرفع الحالات الحرجة إلى أول القائمة».
- العتبة قابلة للتعديل من إعدادات النظام.

ما لا تقوله النشرة: معدل انتشار الحالة في مجموعة الاختبار، ونوع الصور فيها، وكيف وُسمت.

للمشارك: «بارق» ومورِّده افتراضيان. الدرجات في ملف البيانات حقيقية: أخرجها نموذج بحثي مفتوح
على صور عامة (انظر ATTRIBUTION.txt)، وهي تقوم هنا مقام درجات المنتج. أرقام النشرة أعلاه
محسوبة من هذه الدرجات نفسها.
""")
    site_lines = "\n".join(
        f"- {s[0]}: {s[1]} صورة صدر في اليوم، منها {s[2]:.0%} سريرية متنقلة (AP). "
        f"معدل انتشار استرواح الصدر بين المصوَّرين: {s[3]:.1%}." for s in SITES)
    text(FILES / "triage" / "network_case_mix.txt", f"""
توزيع الحالات في شبكة أرياف الصحية · من قسم الأشعة (أرقام اصطناعية من صنع المقرر)

{site_lines}

طاقة القراءة الفورية: أخصائي أشعة الطوارئ يستطيع أن يقطع عمله لقراءة {CAPACITY} تنبيهًا في اليوم.
ما زاد على ذلك ينتظر في القائمة كغيره.

الصورة السريرية المتنقلة (AP) تُلتقط للمريض في سريره. الصورة القياسية (PA) تُلتقط له واقفًا في قسم الأشعة.
""")
    text(FILES / "triage" / "ATTRIBUTION.txt", """
triage_scores.csv and triage_scores.xlsx: derived data.

Scores (column "score"): the probability of pneumothorax predicted for each image of the NIH
ChestX-ray14 test split by the retrained DenseNet-121 of the open project "reproduce-chexnet"
by John Zech (https://github.com/jrzech/reproduce-chexnet, file pretrained/preds.csv),
BSD 3-Clause License, Copyright (c) 2018, John Zech. Rounded to four decimals.

Labels, view, age, sex, patient id: file nih_labels.csv of the same project, taken from the
NIH Clinical Center ChestX-ray14 dataset (Wang et al., ChestX-ray8, CVPR 2017,
https://arxiv.org/abs/1705.02315). The labels were text-mined from radiology reports: they
contain errors. No image is included here.

This is a research model, not a marketed product. "Bariq" in the course is fictional.

BSD 3-Clause notice of reproduce-chexnet: Redistribution and use in source and binary forms,
with or without modification, are permitted provided that the following conditions are met:
1. Redistributions of source code must retain the above copyright notice, this list of
conditions and the following disclaimer. 2. Redistributions in binary form must reproduce the
above copyright notice, this list of conditions and the following disclaimer in the
documentation and/or other materials provided with the distribution. 3. Neither the name of the
copyright holder nor the names of its contributors may be used to endorse or promote products
derived from this software without specific prior written permission.
THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR
IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND
FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR
CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR
OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
POSSIBILITY OF SUCH DAMAGE.
""")
    text(FILES / "triage" / "board_statement_template.txt", """
بيان إلى مجلس الحوكمة السريرية · «بارق» · أربع جمل

1. في [الموقع]، من كل 100 تنبيه يكون [عدد] صحيحًا.
2. ننتظر [عدد] تنبيهًا في اليوم، وطاقة القراءة الفورية [عدد].
3. يفوت الأداة نحو [عدد] حالة في الشهر، وأغلبها في [نوع الصور].
4. نوصي بـ [التشغيل / التشغيل بشرط / عدم التشغيل] في [الموقع]، والشرط: [ ].
""")
    return {"claimed": {"sens": sens_r, "spec": spec_r, "accuracy": round(acc, 3)}}


# ---------------------------------------------------------------------------------- lab 2.1
RES = ["المدينة", "الضواحي", "القرى النائية"]
AGES = ["18-39", "40-59", "60-74", "75+"]
ENROL = 91  # scores 91..100: the 500 places of the programme


def pct_rank(values):
    order = np.argsort(np.argsort(values, kind="stable"), kind="stable")
    return (order * 100 // len(values)) + 1


def lab21():
    rng = np.random.default_rng(SEED)
    n = 5000
    res = rng.choice(3, n, p=[0.55, 0.27, 0.18])
    outside = rng.random(n) < np.array([0.20, 0.12, 0.05])[res]
    age = rng.choice(4, n, p=[0.38, 0.32, 0.21, 0.09])
    sex = rng.choice(["ذكر", "أنثى"], n)
    cond = np.minimum(rng.poisson(np.array([0.4, 1.2, 2.4, 3.3])[age]), 8)
    unc = rng.binomial(cond, np.where(res == 2, 0.45, 0.35))
    access = np.where(res == 2, 0.5, 1.0) * np.where(outside, 0.45, 1.0)
    person = rng.normal(0, 0.35, n)
    # conditions as coded in the network's systems: coding is less complete in the village
    # centres, and a diagnosis made outside the network is often never coded inside it
    coded = rng.binomial(cond, np.where(res == 2, 0.6, 0.9) * np.where(outside, 0.65, 1.0))

    def cost():
        return np.round((800 + 4200 * cond + 3500 * unc) * access * np.exp(person + rng.normal(0, 0.35, n)), -1)
    cost_prev, cost_next = cost(), cost()
    adm_next = rng.poisson(0.02 + 0.07 * unc + 0.02 * cond)
    # the supplier's score: a linear model of next year's recorded cost
    X = np.column_stack([np.ones(n), cost_prev, coded, age])
    beta, *_ = np.linalg.lstsq(X, cost_next, rcond=None)
    score = pct_rank(X @ beta)
    # the alternative: the same recipe trained on next year's avoidable admissions, with the
    # conditions taken from a clinical review of the file
    Xn = np.column_stack([np.ones(n), cond, unc, age])
    bn, *_ = np.linalg.lstsq(Xn, adm_next, rcond=None)
    alt = pct_rank(Xn @ bn + rng.normal(0, 1e-6, n))
    header = ["patient_id", "الفئة العمرية", "الجنس", "السكن", "يتلقى جزءًا من علاجه خارج الشبكة",
              "أمراض مزمنة مرمَّزة في النظام", "أمراض مزمنة بمراجعة الملف", "أمراض مزمنة غير مضبوطة", "التكلفة المسجلة في السنة الماضية (ريال)",
              "درجة أولوية", "مسجَّل في البرنامج", "الدرجة البديلة (مدرَّبة على التنويم)",
              "التكلفة المسجلة في السنة التالية (ريال)", "تنويم يمكن تجنبه في السنة التالية"]
    rows = [[f"A{i + 1:05d}", AGES[age[i]], sex[i], RES[res[i]], "نعم" if outside[i] else "لا", int(coded[i]), int(cond[i]),
             int(unc[i]), int(cost_prev[i]), int(score[i]), "نعم" if score[i] >= ENROL else "لا", int(alt[i]),
             int(cost_next[i]), int(adm_next[i])] for i in range(n)]
    write_csv(FILES / "risk" / "patients_5000.csv", header, rows)
    wb = book()
    sheet(wb, "المرضى", header, rows, [11, 12, 8, 14, 18, 12, 12, 14, 18, 10, 12, 16, 18, 16], first=True)
    ws = wb.create_sheet("التدقيق")
    ws.sheet_view.rightToLeft = True
    for l in [
        ["أ · من في البرنامج؟", "المدينة", "الضواحي", "القرى النائية", "خارج الشبكة: نعم", "خارج الشبكة: لا"],
        ["عدد المرضى"], ["عدد المسجَّلين في البرنامج"], ["نسبة المسجَّلين"],
        [],
        ["ب · عند الدرجة نفسها (المسجَّلون فقط)"],
        ["متوسط الأمراض غير المضبوطة"], ["متوسط التنويم الذي يمكن تجنبه في السنة التالية"],
        [],
        ["ج · الخطأ بحسب الفئة: أصحاب الحاجة العالية (مرضان غير مضبوطين فأكثر)"],
        ["عددهم"], ["المسجَّلون منهم"], ["نسبة من فاتهم البرنامج"],
        [],
        ["د · لو تغيّر الوسم: أعلى 500 بالدرجة البديلة (91 فأكثر)"],
        ["عددهم في أعلى 500"], ["نسبتهم من 500"],
    ]:
        ws.append(l)
    for r in (1, 6, 10, 15):
        for c in ws[r]:
            c.font = Font(bold=True)
            c.fill = HEAD
    for col, w in zip("ABCDEF", [58, 14, 14, 16, 18, 18]):
        ws.column_dimensions[col].width = w
    save(wb, FILES / "risk" / "patients_5000.xlsx")
    text(FILES / "risk" / "awlawiya_datasheet.txt", """
«أولوية» · بطاقة النموذج كما سلّمها المورِّد (مورِّد افتراضي)

الغرض المقصود: التنبؤ بتكلفة الرعاية في السنة القادمة لأغراض تخطيط الميزانية والتعاقد.
المخرج: درجة من 1 إلى 100 لكل مريض مسجَّل. الدرجة 100 أعلى تكلفة متوقعة.
ما تتعلمه الدرجة (الوسم): مجموع التكلفة المسجلة للمريض في أنظمة الشبكة في السنة التالية.
المدخلات: التكلفة المسجلة في السنة الماضية، وعدد الأمراض المزمنة المرمَّزة، والفئة العمرية.
ما لا يدخل: السكن، والجنس، وأي بيان عن الدخل.
الأداء المعلن: «الدرجة ترتّب المرضى بحسب التكلفة الفعلية ترتيبًا جيدًا: من درجتهم 91 فأكثر
يتحملون أكثر من ثلث تكلفة الشبكة».

كيف تستخدمها الشبكة اليوم: إدارة صحة السكان تسجّل كل من درجته 91 فأكثر (500 مريض) في برنامج
إدارة الرعاية: ممرضة منسِّقة، ومواعيد مبكرة، وزيارات منزلية.

للمشارك: المرضى الخمسة آلاف اصطناعيون، ولّدهم مولِّد المقرر ببذرة ثابتة. لم تُؤخذ من سجل حقيقي.
""")
    text(FILES / "risk" / "audit_note_template.txt", """
مذكرة تدقيق · درجة «أولوية» · إلى د. إيناس الذبيب، مديرة صحة السكان

1. ما تتعلمه الدرجة: [الوسم]، وما نريده منها: [القرار].
2. الفئات المتضررة: [الفئة الأولى]، [الفئة الثانية].
3. الدليل: عند الدرجة نفسها [رقم] مقابل [رقم]. من أصحاب الحاجة العالية يفوت البرنامج [نسبة] في [الفئة] مقابل [نسبة] في [الفئة].
4. السبب: [لماذا تنخفض التكلفة المسجلة لهذه الفئة دون أن تنخفض حاجتها].
5. التوصية: [إيقاف / تغيير الوسم / تشغيل بشرط]، ومن يراجع القائمة قبل التسجيل: [الاسم أو الدور].
6. ما لا تقوله هذه البيانات: [حدّ واحد على الأقل].
""")
    return header, rows


# ---------------------------------------------------------------------------------- lab 2.2
BASELINE_MIN = 6.5


def wilson(k, n, z=1.96):
    p = k / n
    d = 1 + z * z / n
    c = (p + z * z / (2 * n)) / d
    h = z * math.sqrt(p * (1 - p) / n + z * z / (4 * n * n)) / d
    return c - h, c + h


def lab22():
    key, tr, nt = [], ["عشرون زيارة مكتوبة · شبكة أرياف الصحية · نصوص من صنع المقرر، لا مرضى حقيقيين", ""], \
        ["عشرون ملاحظة كتبها «مُحرِّر» من الزيارات العشرين · قبل مراجعة الطبيب", ""]
    for c in T.CASES:
        tr += [f"{c['id']} · {c['clinic']}"]
        by = {s: [] for s in T.SECTIONS}
        om = inv = crit = 0
        detail = []
        for k, sec, dq, pa, note in c["facts"]:
            tr += [f"الطبيب: {dq}", f"المريض: {pa}"]
            if k in c["omit"]:
                om += 1
                is_c = (c["id"], k) in T.CRITICAL_OMIT
                crit += is_c
                detail.append(f"إغفال{' جسيم' if is_c else ''}: {note}")
                if k in c.get("replace", {}):
                    by[sec].append(c["replace"][k])
            else:
                by[sec].append(note)
        for sec, line, is_c in c["invent"]:
            inv += 1
            crit += is_c
            by[sec].append(line)
            detail.append(f"مختلَق{' جسيم' if is_c else ''}: {line}")
        tr.append("")
        nt += [f"{c['id']} · {c['clinic']}"] + [f"{s}: {' '.join(by[s])}" for s in T.SECTIONS if by[s]] + [""]
        key.append([c["id"], c["clinic"], sum(len(v) for v in by.values()), om, inv, crit, c["minutes"], "؛ ".join(detail)])
    text(FILES / "notes" / "transcripts_20.txt", "\n".join(tr))
    text(FILES / "notes" / "notes_20.txt", "\n".join(nt))
    head = ["الملاحظة", "العيادة", "إغفالات", "نتائج مختلقة", "منها جسيم (قد يضر مريضًا)",
            "دقائق مراجعة الطبيب وتصحيحه (مسجَّلة)", "ما وجدته"]
    wb = book()
    sheet(wb, "التقييم", head, [[k[0], k[1], "", "", "", k[6], ""] for k in key], [10, 12, 10, 12, 18, 20, 60],
          first=True, fill_cols=(3, 4, 5, 7))
    ws = wb.create_sheet("حدود العينة")
    ws.sheet_view.rightToLeft = True
    ws.append(["ملاحظات فيها خطأ جسيم من 20", "النسبة في العينة", "أدنى ما يتوافق مع العينة", "أعلى ما يتوافق مع العينة"])
    for k in range(0, 11):
        lo, hi = wilson(k, 20)
        ws.append([k, f"{k / 20:.0%}", f"{max(lo, 0):.0%}", f"{hi:.0%}"])
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = HEAD
    for col in "ABCD":
        ws.column_dimensions[col].width = 28
    save(wb, FILES / "notes" / "scoring_sheet.xlsx")
    wb = book()
    sheet(wb, "مفتاح التقييم", ["الملاحظة", "العيادة", "أسطر الملاحظة", "إغفالات", "مختلقة", "جسيم", "دقائق", "التفصيل"],
          key, [10, 12, 12, 10, 10, 10, 10, 80], first=True)
    save(wb, SOL / "scoring_key.xlsx")
    text(FILES / "notes" / "vendor_claim.txt", f"""
«مُحرِّر» · عرض المورِّد على شبكة أرياف (مورِّد افتراضي)

- «يوفّر 70% من وقت التوثيق: ملاحظة كاملة في أقل من دقيقتين».
- «دقة 98% على مستوى الجملة».
- «أطباؤنا التجريبيون راضون جدًّا».

ما عند الشبكة: قياس سابق لوقت الكتابة اليدوية في العيادات نفسها، متوسطه {BASELINE_MIN} دقيقة للزيارة.
وعشرون ملاحظة من أسبوع تجريبي في ثلاث عيادات، مع نصوص زياراتها ودقائق مراجعة كل ملاحظة.

للمشارك: الملاحظات العشرون كتبها مؤلف المقرر وزرع أخطاءها. لم يكتبها نموذج ولا منتج.
""")
    text(FILES / "notes" / "plan_template.txt", """
خطة التحقق المحلي والتجربة الصامتة · «مُحرِّر» · إلى د. كريم المعيقل

أ. ما ندّعي أننا سنتحقق منه: [ادعاء المورِّد بكلماته] ← [ما سنقيسه نحن]
ب. المقاييس (لكل منها: كيف يُحسب، وعلى كم ملاحظة، ومن يحسبه)
   1. الإغفالات الجسيمة لكل 100 ملاحظة:
   2. النتائج المختلقة لكل 100 ملاحظة:
   3. دقائق المراجعة والتصحيح مقابل الكتابة اليدوية:
   4. مقياس للمريض (رفض التسجيل، شكوى، استفسار):
ج. التجربة الصامتة: [كم أسبوعًا]، [أي عيادات]، [كم ملاحظة]. الأداة تكتب ولا تدخل ملاحظتها الملف.
   من يقارن الملاحظتين: [دور]. بماذا: [نص الزيارة / التسجيل].
د. المسؤول عن الخطة: [اسم ودور]. من يوقّع الانتقال إلى التشغيل: [مجلس / دور].
هـ. معايير الإيقاف (رقم لكل معيار):
   - قبل التشغيل، لا ننتقل إن:
   - بعد التشغيل، نوقف إن:
و. المراقبة بعد التشغيل: [عينة كل كم؟ من يراجعها؟]، وما الذي يعيد التحقق: [إصدار جديد، عيادة جديدة، لهجة جديدة].
ز. كيف يستفسر المريض أو يعترض: [القناة، ومن يجيب، وفي كم يوم].
""")
    return key


# ------------------------------------------------------------------------------------- map
def tier(q1, q2, q3):
    if q1 and q2:
        return "عالٍ"
    if q1 or q2:
        return "متوسط"
    return "منخفض"


def the_map():
    yn = lambda b: "نعم" if b else "لا"
    cards = ["أربع عشرة حالة استخدام معروضة على مجلس الحوكمة السريرية في شبكة أرياف · كلها افتراضية", ""]
    ref = []
    for u in T.USE_CASES:
        uid, name, what, vf, rf, q1, q2, q3, v, r, lab = u
        cards += [f"{uid} · {name}", f"ما تفعله: {what}", f"القيمة: {vf}", f"البيانات: {rf}",
                  f"يؤثر في تشخيص مريض بعينه أو علاجه؟ {yn(q1)} · الضرر عند الخطأ جسيم أو يصعب تداركه؟ {yn(q2)} · "
                  f"يراه طبيب ويستطيع ردّه قبل أن يقع شيء؟ {yn(q3)}"]
        if lab:
            cards.append(f"دليلك عليها من المختبر {lab}.")
        cards.append("")
        t = tier(q1, q2, q3)
        ref.append([uid, name, v, r, t, v * r,
                    "يُسمّى (طبيب بالدور والاسم)" if t == "عالٍ" else "",
                    "تُكتب (قناة، ومن يجيب، وفي كم يوم)" if t == "عالٍ" else ""])
    text(FILES / "map" / "use_cases_14.txt", "\n".join(cards))
    text(FILES / "map" / "scoring_scale.txt", """
مقياس خريطة الفرص والمخاطر · من وضع المقرر للتعليم، وليس تصنيفًا نظاميًا

القيمة (1 إلى 5): كم مريضًا أو موظفًا تمسّه الحالة، وكم يتغير له؟
  1 راحة لقلّة · 3 وقت أو مال واضح لقسم · 5 ضرر جسيم يُتجنَّب أو خدمة تصل إلى من لا تصله

جاهزية البيانات (1 إلى 5): هل البيانات موجودة، ومرمَّزة، ومشتركة بين المواقع، ولها أساس نظامي، ولها مرجع نقيس عليه؟
  1 لا بيانات ولا مرجع · 3 موجودة وناقصة أو في نظامين · 5 كاملة ومنظمة ولها مرجع محلي

مستوى المخاطر (ثلاثة أسئلة):
  س1 هل يؤثر المخرج في تشخيص مريض بعينه أو علاجه؟
  س2 هل الضرر عند الخطأ جسيم أو يصعب تداركه؟
  س3 هل يرى المخرجَ طبيبٌ ويستطيع ردّه قبل أن يقع شيء؟
  عالٍ: س1 نعم وس2 نعم · متوسط: واحد منهما نعم · منخفض: س1 لا وس2 لا
  س3 لا يغيّر المستوى، بل يحدد الشرط: حالة عالية لا يراها طبيب قبل وقوع أثرها لا تُشغَّل حتى يُسمّى مراجعها.

قاعدة المقرر: كل حالة «عالٍ» تسمّي مراجعها البشري، وتكتب كيف يستفسر المريض عن قرار أو يعترض عليه.
الترتيب: القيمة × الجاهزية، ثم يُقرأ مع مستوى المخاطر: العالي لا يُحذف، بل يُشترط له.
""")
    head = ["الحالة", "الاسم", "القيمة (1-5)", "جاهزية البيانات (1-5)", "مستوى المخاطر", "القيمة × الجاهزية",
            "المراجع البشري (للعالي)", "كيف يستفسر المريض (للعالي)", "الدليل: من أين جاءت الدرجة؟"]
    wb = book()
    sheet(wb, "الخريطة", head, [["", "", "", "", "", "", "", "", ""] for _ in range(10)],
          [9, 34, 12, 16, 14, 14, 26, 30, 40], first=True)
    ws = wb.create_sheet("المشروع التجريبي")
    ws.sheet_view.rightToLeft = True
    for l in [["المشروع التجريبي الواحد", ""], ["الحالة المختارة", ""], ["لماذا هذه أولًا", ""],
              ["المقاييس", ""], ["المسؤول", ""], ["معايير الإيقاف", ""], ["المراجع البشري", ""],
              ["كيف يستفسر المريض", ""]]:
        ws.append(l)
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 80
    save(wb, FILES / "map" / "map_template.xlsx")
    wb = book()
    sheet(wb, "خريطة مرجعية", head[:8], sorted(ref, key=lambda x: -x[5]), [9, 44, 12, 16, 14, 14, 30, 36], first=True)
    save(wb, SOL / "reference_map.xlsx")
    return ref


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--nih-source")
    a = ap.parse_args()
    random.seed(SEED)
    prework()
    k11 = lab11()
    k12 = lab12(a.nih_source)
    lab21()
    k22 = lab22()
    ref = the_map()
    key = {"seed": SEED, "lab11": k11, "lab12": {**k12, "threshold": THRESHOLD, "capacity": CAPACITY,
                                                 "sites": [list(s) for s in SITES]},
           "lab21": {"enrol_from": ENROL, "high_need_from": 2},
           "lab22": {"baseline_minutes": BASELINE_MIN,
                     "notes": [{"id": k[0], "omissions": k[3], "invented": k[4], "critical": k[5]} for k in k22]},
           "map": [{"id": r[0], "value": r[2], "readiness": r[3], "tier": r[4]} for r in ref]}
    (SOL / "key.json").write_text(json.dumps(key, ensure_ascii=False, indent=1) + "\n", encoding="utf-8", newline="\n")
    print("files written under", FILES.name, "and", SOL.name)


if __name__ == "__main__":
    main()
