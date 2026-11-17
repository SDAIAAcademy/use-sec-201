# /// script
# requires-python = ">=3.11"
# dependencies = ["openpyxl>=3.1"]
# ///
"""Reference solutions of the four labs and the map, recomputed from the files participants get.

    uv run labs/solutions/check.py            # every lab
    uv run labs/solutions/check.py 1.2 2.1    # some

Prints what it recomputed and PASS per lab. Writes nothing and calls no model.
"""
import csv
import json
import math
import sys
from pathlib import Path

from openpyxl import load_workbook

LABS = Path(__file__).resolve().parents[1]
F = LABS / "files"
KEY = json.loads((LABS / "solutions" / "key.json").read_text(encoding="utf-8"))


def table(path, sheet=None):
    ws = load_workbook(path, read_only=True)
    ws = ws[sheet] if sheet else ws.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    return rows[0], rows[1:]


def pct(x, d=0):
    return f"{100 * x:.{d}f}%"


def lab11():
    _, sheet_rows = table(F / "products" / "classification_sheet.xlsx")
    _, key_rows = table(LABS / "solutions" / "classification_key.xlsx")
    text = (F / "products" / "products_12.txt").read_text(encoding="utf-8")
    assert len(sheet_rows) == len(key_rows) == 12
    for s, k in zip(sheet_rows, key_rows):
        assert s[0] == k[0] and f"{k[0]} · {k[1]}" in text, k[0]
        assert s[2] is None, "the participants' sheet must be empty"
    dev = [k[0] for k in key_rows if k[2] == "جهاز طبي"]
    assert dev == KEY["lab11"]["devices"] and len(dev) == 6
    traps = [k[0] for k in key_rows if k[5]]
    assert traps == KEY["lab11"]["traps"]
    print("  devices:", " ".join(dev))
    print("  not devices:", " ".join(k[0] for k in key_rows if k[2] != "جهاز طبي"))
    print("  rows with a trap:", " ".join(traps))


def counts(rows, t, view=None):
    tp = fn = fp = tn = 0
    for v, y, s in rows:
        if view and v != view:
            continue
        if y:
            if s >= t: tp += 1
            else: fn += 1
        elif s >= t: fp += 1
        else: tn += 1
    return tp, fn, fp, tn


def lab12():
    head, data = table(F / "triage" / "triage_scores.xlsx", "البيانات")
    rows = [(r[4], r[5], r[6]) for r in data]
    with open(F / "triage" / "triage_scores.csv", encoding="utf-8") as f:
        assert sum(1 for _ in f) - 1 == len(rows) == 22433
    pos = sum(r[1] for r in rows)
    print(f"  images {len(rows)}, pneumothorax {pos} ({pct(pos / len(rows), 2)}), "
          f"AP {sum(1 for r in rows if r[0] == 'AP')}, PA {sum(1 for r in rows if r[0] == 'PA')}")
    k = KEY["lab12"]
    t = k["threshold"]
    out = {}
    for name, th, view in [("all", t, None), ("AP", t, "AP"), ("PA", t, "PA"), ("0.02", 0.02, None),
                           ("0.10", 0.10, None), ("0.20", 0.20, None)]:
        tp, fn, fp, tn = counts(rows, th, view)
        se, sp = tp / (tp + fn), tn / (tn + fp)
        out[name] = (se, sp)
        print(f"  {name:>4} threshold {th}: TP {tp} FN {fn} FP {fp} TN {tn} | sensitivity {pct(se, 1)} "
              f"specificity {pct(sp, 1)} | PPV in this set {pct(tp / (tp + fp), 1)}")
    tp, fn, fp, tn = counts(rows, t)
    assert round(out["all"][0], 2) == k["claimed"]["sens"] and round(out["all"][1], 2) == k["claimed"]["spec"]
    assert round((tp + tn) / len(rows), 3) == k["claimed"]["accuracy"]
    assert out["AP"][0] < out["PA"][0] - 0.05, "the bedside films must be the weaker subgroup"
    flag_none = 1 - pos / len(rows)
    print(f"  accuracy of the tool {pct((tp + tn) / len(rows), 1)}; of a rule that never alerts {pct(flag_none, 1)}")
    se, sp = k["claimed"]["sens"], k["claimed"]["spec"]
    print("  per site, with the claimed sensitivity and specificity, 30 days:")
    res = []
    for name, per_day, ap_share, prev in k["sites"]:
        sick = per_day * prev
        alerts = sick * se + (per_day - sick) * (1 - sp)
        ppv = sick * se / alerts
        missed = sick * (1 - se) * 30
        res.append((alerts, ppv, missed))
        print(f"    {name}: alerts/day {alerts:.1f}, PPV {pct(ppv, 1)}, true alerts/day {sick * se:.2f}, "
              f"missed/month {missed:.1f}, cases/month {sick * 30:.0f}")
    assert res[0][1] > res[1][1] > res[2][1] and res[2][1] < 0.01
    name, per_day, ap_share, prev = k["sites"][0]
    se_mix = ap_share * out["AP"][0] + (1 - ap_share) * out["PA"][0]
    print(f"  emergency, sensitivity measured by view ({pct(ap_share)} AP): {pct(se_mix, 1)}; "
          f"missed/month {per_day * prev * (1 - se_mix) * 30:.1f}")
    print("  emergency, by threshold (sensitivity for its mix of views):")
    best = None
    for th in (0.02, 0.05, 0.10, 0.20):
        a = counts(rows, th, "AP"); p = counts(rows, th, "PA")
        se_t = ap_share * a[0] / (a[0] + a[1]) + (1 - ap_share) * p[0] / (p[0] + p[1])
        sp_t = ap_share * a[3] / (a[2] + a[3]) + (1 - ap_share) * p[3] / (p[2] + p[3])
        sick = per_day * prev
        alerts = sick * se_t + (per_day - sick) * (1 - sp_t)
        print(f"    {th}: sensitivity {pct(se_t, 1)}, alerts/day {alerts:.1f}, PPV {pct(sick * se_t / alerts, 1)}, "
              f"missed/month {sick * (1 - se_t) * 30:.1f}")
        if alerts <= k["capacity"] and best is None:
            best = th
    print(f"  lowest threshold within the reading capacity of {k['capacity']} alerts a day: {best}")
    assert best == 0.05


def lab21():
    head, data = table(F / "risk" / "patients_5000.xlsx", "المرضى")
    with open(F / "risk" / "patients_5000.csv", encoding="utf-8-sig") as f:
        c = list(csv.reader(f))
    assert len(data) == len(c) - 1 == 5000 and list(head) == c[0]
    H = {h: i for i, h in enumerate(head)}
    RES, OUT, UNC, SC, EN, ALT, ADM, COST = (H[x] for x in (
        "السكن", "يتلقى جزءًا من علاجه خارج الشبكة", "أمراض مزمنة غير مضبوطة", "درجة أولوية", "مسجَّل في البرنامج",
        "الدرجة البديلة (مدرَّبة على التنويم)", "تنويم يمكن تجنبه في السنة التالية",
        "التكلفة المسجلة في السنة التالية (ريال)"))
    k = KEY["lab21"]
    assert all((r[EN] == "نعم") == (r[SC] >= k["enrol_from"]) for r in data)
    enrolled = [r for r in data if r[EN] == "نعم"]
    assert len(enrolled) == 500
    print(f"  enrolled {len(enrolled)}; their share of next year's recorded cost "
          f"{pct(sum(r[COST] for r in enrolled) / sum(r[COST] for r in data))}")
    mean = lambda xs: sum(xs) / len(xs) if xs else float("nan")
    groups = [("city", lambda r: r[RES] == "المدينة"), ("suburbs", lambda r: r[RES] == "الضواحي"),
              ("remote villages", lambda r: r[RES] == "القرى النائية"),
              ("outside care: yes", lambda r: r[OUT] == "نعم"), ("outside care: no", lambda r: r[OUT] == "لا")]
    out = {}
    for name, g in groups:
        pop = [r for r in data if g(r)]
        en = [r for r in pop if r[EN] == "نعم"]
        high = [r for r in pop if r[UNC] >= k["high_need_from"]]
        high_en = [r for r in high if r[EN] == "نعم"]
        alt = [r for r in pop if r[ALT] >= k["enrol_from"]]
        out[name] = dict(n=len(pop), en=len(en), unc=mean([r[UNC] for r in en]), adm=mean([r[ADM] for r in en]),
                         high=len(high), missed=1 - len(high_en) / len(high), alt=len(alt))
        o = out[name]
        print(f"  {name:>17}: n {o['n']} ({pct(o['n'] / 5000)}), enrolled {o['en']} ({pct(o['en'] / o['n'], 1)} of group, "
              f"{pct(o['en'] / 500, 1)} of places) | enrolled: uncontrolled {o['unc']:.2f}, admissions {o['adm']:.2f} | "
              f"high need {o['high']}, missed {pct(o['missed'])} | top 500 by alternative {o['alt']} ({pct(o['alt'] / 500, 1)})")
    r, c = out["remote villages"], out["city"]
    assert r["en"] / r["n"] < 0.6 * c["en"] / c["n"] and r["unc"] > c["unc"] + 0.3 and r["missed"] > c["missed"] + 0.1
    assert r["alt"] > 1.8 * r["en"]
    o, i = out["outside care: yes"], out["outside care: no"]
    assert o["en"] / o["n"] < 0.6 * i["en"] / i["n"] and o["missed"] > i["missed"] + 0.1
    allhigh = [x for x in data if x[UNC] >= k["high_need_from"]]
    print(f"  high need overall {len(allhigh)}, enrolled {sum(1 for x in allhigh if x[EN] == 'نعم')}")


def wilson(k, n, z=1.96):
    p = k / n; d = 1 + z * z / n
    c = (p + z * z / (2 * n)) / d
    h = z * math.sqrt(p * (1 - p) / n + z * z / (4 * n * n)) / d
    return c - h, c + h


def lab22():
    _, key_rows = table(LABS / "solutions" / "scoring_key.xlsx")
    _, sheet_rows = table(F / "notes" / "scoring_sheet.xlsx", "التقييم")
    notes = (F / "notes" / "notes_20.txt").read_text(encoding="utf-8")
    trans = (F / "notes" / "transcripts_20.txt").read_text(encoding="utf-8")
    assert len(key_rows) == len(sheet_rows) == 20
    for kr, sr, kk in zip(key_rows, sheet_rows, KEY["lab22"]["notes"]):
        assert kr[0] == sr[0] == kk["id"] and kr[0] in notes and kr[0] in trans
        assert sr[2] is None and sr[5] == kr[6]
        assert (kr[3], kr[4], kr[5]) == (kk["omissions"], kk["invented"], kk["critical"])
        # every invented line is in the note and nowhere in the visit; every omitted fact is absent from the note
        for part in kr[7].split("؛ ") if kr[7] else []:
            kind, line = part.split(": ", 1)
            block = notes.split(kr[0] + " ·")[1].split("\n\n")[0]
            assert (line in block) == kind.startswith("مختلَق"), (kr[0], part)
    lines = sum(r[2] for r in key_rows); om = sum(r[3] for r in key_rows); inv = sum(r[4] for r in key_rows)
    any_err = sum(1 for r in key_rows if r[3] + r[4]); crit = sum(1 for r in key_rows if r[5])
    lo, hi = wilson(crit, 20)
    mins = [r[6] for r in key_rows]
    clean = [r[6] for r in key_rows if not r[3] + r[4]]; bad = [r[6] for r in key_rows if r[3] + r[4]]
    base = KEY["lab22"]["baseline_minutes"]
    print(f"  note lines {lines}; omissions {om}, invented {inv}; lines free of an invented finding {pct(1 - inv / lines, 1)}")
    print(f"  notes with any error {any_err} of 20; with a critical error {crit} of 20 = {pct(crit / 20)} "
          f"(compatible with {pct(lo)} to {pct(hi)})")
    print("  critical:", " ".join(r[0] for r in key_rows if r[5]), "| other errors:",
          " ".join(r[0] for r in key_rows if r[3] + r[4] and not r[5]))
    m = sum(mins) / 20
    print(f"  review minutes: mean {m:.2f} (clean notes {sum(clean) / len(clean):.2f}, notes with errors "
          f"{sum(bad) / len(bad):.2f}); baseline {base}; saving {pct(1 - m / base)} against the claimed 70%")
    assert (om, inv, any_err, crit) == (6, 3, 9, 5) and 0.4 < 1 - m / base < 0.6


def themap():
    _, ref = table(LABS / "solutions" / "reference_map.xlsx")
    cards = (F / "map" / "use_cases_14.txt").read_text(encoding="utf-8")
    assert len(ref) == 14 and all(r[0] in cards for r in ref)
    high = [r for r in ref if r[4] == "عالٍ"]
    assert all(r[6] and r[7] for r in high) and len(ref) >= 8
    for r in ref:
        print(f"  {r[0]} value {r[2]} readiness {r[3]} product {r[5]:>2} tier {r[4]}")
    print("  high tier:", " ".join(r[0] for r in high), "| each names a reviewer and a patient query route")


ALL = {"1.1": lab11, "1.2": lab12, "2.1": lab21, "2.2": lab22, "map": themap}
if __name__ == "__main__":
    for name in sys.argv[1:] or ALL:
        print(f"Lab {name}")
        ALL[name]()
        print(f"  PASS {name}")
