"""
Phase 3 (Arabic) - Render the ten showcase images in Arabic (showcase_ar/01..10 *.png).

Same data and numbers as scripts/04_showcase.py; every user-visible string is translated and
the layouts are mirrored to read right-to-left. Arabic strings go through the helper A(): with
matplotlib >= 3.11 (native text shaping) it only sets a right-to-left base direction; on older
versions it shapes with arabic_reshaper and reorders with python-bidi.

Reads only the outputs of the earlier steps: clean_data/school.db, clean_data/staging_grades.csv,
clean_data/pipeline_stats.json, reports/evaluation.json and the answer-key counts.
"""
import json
import re
import sqlite3
import time
from pathlib import Path

import arabic_reshaper
import matplotlib
from bidi.algorithm import get_display
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.colors import LinearSegmentedColormap
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch

ROOT = Path(__file__).resolve().parents[1]
SHOW = ROOT / "showcase_ar"
SHOW.mkdir(exist_ok=True)


# Matplotlib >= 3.11 shapes and reorders complex scripts itself. Applying arabic_reshaper + python-bidi on
# top of that reverses the text a second time, so the manual path is only a fallback for older versions.
# In both cases every line gets a right-to-left base direction (RLM mark), so lines that start with a
# Latin word or a number still read right-to-left.
NATIVE_BIDI = tuple(int(p) for p in matplotlib.__version__.split(".")[:2]) >= (3, 11)
RLM = "‏"
RANGE = re.compile(r"\d+(?:\.\d+)?-\d+(?:\.\d+)?")


def A(s):
    """Prepare an Arabic string (line by line) for matplotlib with a right-to-left base direction."""
    lines = str(s).split("\n")
    if NATIVE_BIDI:
        # keep numeric ranges such as 0-100 left-to-right (left-to-right isolate around them)
        return "\n".join(RLM + RANGE.sub("⁦\\g<0>⁩", line) for line in lines)
    return "\n".join(get_display(arabic_reshaper.reshape(line), base_dir="R") for line in lines)


# ---- palette (validated reference palette, light surface) -------------------
SURFACE, INK, INK2, MUTED, GRID, AXIS = "#fcfcfb", "#0b0b0b", "#52514e", "#898781", "#e1e0d9", "#c3c2b7"
BLUE, ORANGE, AQUA, YELLOW, MAGENTA, GREEN = "#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300"
GOOD, WARN, CRIT = "#0ca30c", "#fab219", "#d03b3b"
CARD = "#f1f0ec"
SEQ = LinearSegmentedColormap.from_list("blue", ["#cde2fb", "#86b6ef", "#3987e5", "#1c5cab", "#0d366b"])
plt.rcParams.update({
    "font.family": ["Segoe UI", "Tahoma", "DejaVu Sans"], "font.size": 11, "text.color": INK,
    "axes.labelcolor": INK2, "xtick.color": MUTED, "ytick.color": INK2, "axes.edgecolor": AXIS,
    "figure.facecolor": SURFACE, "axes.facecolor": SURFACE, "savefig.facecolor": SURFACE,
    "axes.spines.top": False, "axes.spines.right": False, "axes.grid": False, "legend.frameon": False})

# ---- data ---------------------------------------------------------------------
con = sqlite3.connect(ROOT / "clean_data" / "school.db")
q = lambda s: pd.read_sql(s, con)
clog, review, flags = q("select * from cleaning_log"), q("select * from review_queue"), q("select * from log_flags")
ov = q("select * from v_student_overview")
grades = q("select * from grades")
staging = pd.read_csv(ROOT / "clean_data" / "staging_grades.csv", keep_default_na=False)
stats = json.loads((ROOT / "clean_data" / "pipeline_stats.json").read_text(encoding="utf-8"))
ev = json.loads((ROOT / "reports" / "evaluation.json").read_text(encoding="utf-8"))
key = pd.read_csv(ROOT / "ground_truth" / "injected_anomalies.csv", keep_default_na=False)

SOURCES = list(stats["sources"])
SHORT = {SOURCES[0]: "سجل الطلاب", SOURCES[1]: "استمارة التسجيل", SOURCES[2]: "دفتر الدرجات",
         SOURCES[3]: "نتائج العلوم", SOURCES[4]: "تقرير التاريخ", SOURCES[5]: "سجل البوابة"}
FMT = {SOURCES[0]: "Excel", SOURCES[1]: "CSV", SOURCES[2]: "Excel", SOURCES[3]: "Word", SOURCES[4]: "PDF", SOURCES[5]: "Log"}
SUBJ = ["Math", "English", "Science", "History"]
SUBJ_AR = {"Math": "الرياضيات", "English": "اللغة الإنجليزية", "Science": "العلوم", "History": "التاريخ"}
checked = [t for block in ("student_fields", "grades_clean", "attendance", "guardians") for t in ev[block].values()]
N_CHECK = sum(sum(t.values()) for t in checked)
N_OK, N_HELD, N_WRONG = (sum(t[k] for t in checked) for k in ("correct", "withheld", "wrong"))
naive = {k: sum(t[k] for t in ev["grades_naive"].values()) for k in ("correct", "withheld", "wrong")}
clean = {k: sum(t[k] for t in ev["grades_clean"].values()) for k in ("correct", "withheld", "wrong")}

# category codes -> Arabic labels (review queue and log flags)
CAT_AR = {
    "missing_checkout": "لا تسجيل خروج", "missing_checkin": "لا تسجيل دخول", "no_form_response": "لا استمارة تسجيل",
    "text_not_available": "لا درجة مسجّلة (N/A)", "early_leave": "مغادرة مبكرة", "row_missing": "صف مفقود",
    "scale_gpa": "مقياس GPA مفترض", "letter_grade": "تقدير بالحروف", "unknown_id": "معرّف مجهول",
    "unknown_id_swipe": "تمرير بمعرّف مجهول", "absent_but_swiped": "غائب لكنه مرّر البطاقة", "text_absent": "غائب عن الامتحان",
    "dob_missing": "تاريخ ميلاد مفقود", "odd_hour_swipe": "تمرير في ساعة غير معتادة", "weekend_swipe": "تمرير في عطلة",
    "id_empty": "معرّف فارغ", "no_record": "لا سجل لليوم", "impossible_date": "تاريخ مستحيل",
    "dob_illogical": "تاريخ ميلاد غير منطقي", "total_mismatch": "مجموع غير مطابق", "duplicate_id_conflict": "معرّف مشترك",
    "duplicate_conflict": "صفان متعارضان", "name_mismatch": "اختلاف في الاسم", "illogical_value": "قيمة غير منطقية",
    "scale_percent_in_gpa_column": "نسبة في عمود GPA", "misplaced_decimal": "فاصلة عشرية في غير موضعها",
    "wrong_id": "معرّف خاطئ", "gpa_over_4": "GPA أعلى من 4", "email_invalid": "بريد غير صالح", "phone_invalid": "هاتف غير صالح",
    "blank": "حقل فارغ",
}


def frame(fig, title, subtitle, note="أكاديمية Greenfield مدرسة خيالية - جميع البيانات اصطناعية."):
    fig.text(0.96, 0.945, A(title), fontsize=21, fontweight="bold", va="center", ha="right")
    fig.text(0.96, 0.895, A(subtitle), fontsize=12.5, color=INK2, va="center", ha="right")
    fig.text(0.96, 0.025, A(note), fontsize=9, color=MUTED, va="center", ha="right")


def canvas(title, subtitle, **kw):
    fig = plt.figure(figsize=(12, 7.5), dpi=150)
    frame(fig, title, subtitle, **kw)
    return fig


def blank_ax(fig, rect=(0.04, 0.06, 0.92, 0.79)):
    ax = fig.add_axes(rect)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis("off")
    return ax


def card(ax, x, y, w, h, color=CARD, edge="none", lw=0):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0,rounding_size=1.2", fc=color, ec=edge, lw=lw,
                                mutation_aspect=ax.bbox.height / ax.bbox.width * 1.0))


def save(fig, name):
    for attempt in range(5):  # the file can be briefly locked on Windows (preview pane, antivirus)
        try:
            fig.savefig(SHOW / name)
            break
        except OSError:
            if attempt == 4:
                raise
            time.sleep(1)
    plt.close(fig)
    print("wrote", name)


def clean_axis(ax, grid="x"):
    ax.grid(axis=grid, color=GRID, lw=0.8)
    ax.set_axisbelow(True)
    ax.tick_params(length=0)


def rtl_bars(ax):
    """Bars grow right-to-left with category labels on the right."""
    ax.invert_xaxis()
    ax.yaxis.tick_right()
    ax.spines["left"].set_visible(False)
    ax.spines["right"].set_visible(False)


# =============================================================================
# 01 - project overview
# =============================================================================
fig = canvas("من ستة ملفات فوضوية إلى قاعدة بيانات مدرسية موثوقة",
             "بيانات اصطناعية لمدرسة من 240 طالبًا، أُفسدت عمدًا ثم حُلِّلت ونُظِّفت ودُقِّقت.")
ax = blank_ax(fig)
ax.text(100, 97, A("1  التوليد"), fontsize=10.5, fontweight="bold", color=MUTED, ha="right")
ax.text(64, 97, A("2  التحليل والتنظيف"), fontsize=10.5, fontweight="bold", color=MUTED, ha="right")
ax.text(28, 97, A("3  التجميع"), fontsize=10.5, fontweight="bold", color=MUTED, ha="right")
for i, s in enumerate(SOURCES):
    y = 84 - i * 10.6
    card(ax, 72, y, 28, 9)
    ax.text(98.5, y + 4.5, FMT[s], fontsize=10, fontweight="bold", color="white", va="center", ha="right",
            bbox=dict(boxstyle="round,pad=0.35", fc=BLUE, ec="none"))
    ax.text(91, y + 5.9, A(SHORT[s]), fontsize=11.5, fontweight="bold", va="center", ha="right")
    ax.text(91, y + 2.7, A(f"{stats['sources'][s]['records_read']:,} سجل - {(key.source == s).sum()} خطأ مزروع"),
            fontsize=9.5, color=INK2, va="center", ha="right")
card(ax, 36, 31, 28, 62, color="#e6eefa")
steps = [("تحديد موقع البيانات", "صفوف العناوين، الخلايا المدمجة، فواصل\nالصفحات، ضجيج السجل"),
         ("توحيد الصيغ", "المعرّفات، الأسماء، التواريخ، الكسور\nالعشرية، أرقام الهواتف"),
         ("توحيد المقاييس", "0-1 ومن 20 و GPA 4.0 والنسبة المئوية\n← مقياس واحد 0-100"),
         ("التحقق المتقاطع بين المصادر", "ملء ملف والتحقق منه\nبالاستناد إلى ملف آخر"),
         ("رفض التخمين", "القيم المستحيلة تُحجب وتُحال\nإلى شخص للمراجعة")]
for i, (h, d) in enumerate(steps):
    y = 86 - i * 11.6
    ax.text(62, y, A(h), fontsize=11.5, fontweight="bold", va="center", ha="right")
    ax.text(62, y - 4.6, A(d), fontsize=9.3, color=INK2, va="center", ha="right", linespacing=1.45)
outs = [("قاعدة بيانات SQLite", "9 جداول وعرض واحد (school.db)،\nمع مفاتيح وفحوص للنطاقات"),
        ("مصنّف Excel", "11 ورقة (school_clean.xlsx)،\nمرشّحة ومرمّزة بالألوان"),
        ("تقرير التنظيف", f"{len(clog)} إجراءً مسجّلًا،\n{len(review)} عنصر مراجعة مع طريقة لكلٍّ منها")]
for i, (h, d) in enumerate(outs):
    y = 73 - i * 21
    card(ax, 0, y, 28, 18)
    ax.text(26, y + 13, A(h), fontsize=12, fontweight="bold", va="center", ha="right")
    ax.text(26, y + 6, A(d), fontsize=9.5, color=INK2, va="center", ha="right", linespacing=1.45)
for x0, x1 in [(71.4, 64.6), (35.4, 28.6)]:
    ax.add_patch(FancyArrowPatch((x0, 62), (x1, 62), arrowstyle="-|>", mutation_scale=18, color=MUTED, lw=1.6))
tiles = [(f"{stats['students']}", "طالب"), ("6", "ملفات مصدر،\n5 صيغ"), (f"{len(key)}", "خطأ مزروع"),
         (f"{len(clog)}", "إجراء تنظيف\nمسجَّل"), (f"{len(review)}", "عنصر أُحيل\nللمراجعة"),
         (f"{N_WRONG}", f"قيمة خاطئة من\n{N_CHECK:,} قيمة مفحوصة")]
for i, (big, small) in enumerate(tiles):
    x = 100 - 15.5 - i * 16.9
    card(ax, x, 0, 15.5, 21, color=SURFACE, edge=GRID, lw=1)
    ax.text(x + 7.75, 14, big, fontsize=25, fontweight="bold", ha="center", va="center",
            color=GOOD if i == 5 else INK)
    ax.text(x + 7.75, 5.2, A(small), fontsize=9.3, color=INK2, ha="center", va="center", linespacing=1.5)
save(fig, "01_project_overview.png")

# =============================================================================
# 02 - the six sources and what was wrong with each
# =============================================================================
PROBLEMS = {
    SOURCES[0]: "صفوف عناوين فوق الترويسة، ترتيب أعمدة مختلف في إحدى الأوراق، تواريخ نصية\nبخمس صيغ، سنوات ميلاد مستحيلة، طالبان يتشاركان معرّفًا واحدًا",
    SOURCES[1]: "إجابات نصية حرة: الصف \"Tenth\"، والمعدلات 88 / 88% / 0.88 / 3.4 / B+‎،\nعدد الإخوة \"two\"، إرسالات مكررة، إدخالات تجريبية، بريد وهواتف معطوبة",
    SOURCES[2]: "\"absent\" و \"N/A\" في خانات الدرجات، أرقام مخزّنة كنص، مجاميع تتجاوز 100،\nمجاميع لا تساوي الأعمال + الامتحان، شعبة كاملة مُدخلة بمقياس 0-1",
    SOURCES[3]: "درجات بمقياس 0-1 لكن الصف 11 من 20، نسب مئوية شاردة، فواصل عشرية\nبالفاصلة، تقديرات بالحروف، صفوف باسم دون معرّف، حرف O بدل الصفر",
    SOURCES[4]: "معدل GPA بمقياس 4.0، قيم أعلى من 4، نسب مئوية في عمود GPA،\nترويسة مكررة في كل صفحة، صفوف مكررة، معرّفات مقسومة بمسافات",
    SOURCES[5]: "تمريرات مزدوجة، بطاقات مجهولة، تمريرات في الثانية فجرًا وفي العطل، أسطر\nمقطوعة، بوابة تكتب اليوم أولًا في التاريخ، إدخالات ABSENT تناقض التمريرات",
}
fig = canvas("ستة مصادر، خمس صيغ، ولكلٍّ منها أخطاؤه المعروفة",
             "ما وُلِّد في المرحلة 1. كل خطأ مزروع مسجَّل في مفتاح إجابات مخفي.")
ax = blank_ax(fig)
for x, t in [(100, "المصدر"), (75, "السجلات"), (64.5, "الأخطاء المزروعة"), (45, "المشكلات النموذجية في هذه الصيغة")]:
    ax.text(x, 97.5, A(t), fontsize=9.5, fontweight="bold", color=MUTED, va="center", ha="right")
mx = max((key.source == s).sum() for s in SOURCES)
for i, s in enumerate(SOURCES):
    y = 80.5 - i * 15.6
    card(ax, 0, y, 100, 13.4)
    ax.text(98.5, y + 6.7, FMT[s], fontsize=10, fontweight="bold", color="white", va="center", ha="right",
            bbox=dict(boxstyle="round,pad=0.35", fc=BLUE, ec="none"))
    ax.text(91.5, y + 8.6, A(SHORT[s]), fontsize=12, fontweight="bold", va="center", ha="right")
    ax.text(91.5, y + 4.4, s.split("_", 1)[1], fontsize=7.3, color=MUTED, va="center", ha="right")
    ax.text(75, y + 6.7, f"{stats['sources'][s]['records_read']:,}", fontsize=14, fontweight="bold", va="center", ha="right")
    n = (key.source == s).sum()
    w = 13 * n / mx
    ax.add_patch(plt.Rectangle((64.5 - w, y + 5.2), w, 3, fc=ORANGE, ec="none"))
    ax.text(64.5 - w - 1, y + 6.7, str(n), fontsize=11, fontweight="bold", va="center", ha="right")
    ax.text(46, y + 6.7, A(PROBLEMS[s]), fontsize=9.2, color=INK2, va="center", ha="right", linespacing=1.45)
save(fig, "02_sources_and_formats.png")

# =============================================================================
# 03 - anomaly map: family x source
# =============================================================================
FAMILY = {
    "Format and notation": ["id_format", "id_letter_O", "name_case_space", "gender_variant", "section_format",
                            "dob_text_format", "number_as_text", "comma_decimal", "email_case_space", "email_typo",
                            "timestamp_format", "scale_percent_sign", "number_as_word", "number_with_text"],
    "Wrong scale or unit": ["scale_fraction", "scale_gpa", "scale_percent_in_fraction_column",
                            "scale_percent_in_gpa_column"],
    "Text where a number belongs": ["letter_grade", "text_absent", "text_not_available"],
    "Invalid or illogical value": ["dob_illogical", "illogical_value", "illogical_total", "total_mismatch",
                                   "component_out_of_range", "misplaced_decimal", "out_of_range", "gpa_over_4",
                                   "impossible_date", "grade_mismatch", "dob_conflict", "odd_hour_swipe",
                                   "weekend_swipe", "early_leave", "late_not_marked", "absent_but_swiped",
                                   "email_invalid", "phone_invalid"],
    "Duplicates and junk rows": ["duplicate_row", "duplicate_submission", "duplicate_swipe", "duplicate_conflict",
                                 "test_submission", "truncated_line"],
    "Missing data": ["dob_missing", "gender_missing", "blank", "no_form_response", "row_missing", "total_blank",
                     "missing_checkin", "missing_checkout", "no_record", "id_empty"],
    "Identity and record linking": ["duplicate_id_conflict", "wrong_id", "unknown_id", "name_mismatch", "id_missing",
                                    "id_missing_name_typo"],
}
FAM_AR = {"Format and notation": "الصيغة والترميز", "Wrong scale or unit": "مقياس أو وحدة خاطئة",
          "Text where a number belongs": "نص في موضع رقم", "Invalid or illogical value": "قيمة غير صالحة أو غير منطقية",
          "Duplicates and junk rows": "تكرارات وصفوف زائدة", "Missing data": "بيانات مفقودة",
          "Identity and record linking": "الهوية وربط السجلات"}
cat2fam = {c: f for f, cs in FAMILY.items() for c in cs}
clog["family"] = clog.category.map(cat2fam).fillna("Invalid or illogical value")
SRC_RTL = SOURCES[::-1]  # first source on the right
mat = clog.groupby(["family", "source"]).size().unstack(fill_value=0).reindex(index=list(FAMILY), columns=SRC_RTL, fill_value=0)
fig = canvas("أين كانت المشكلات: نوع المشكلة حسب المصدر",
             f"جميع مدخلات سجل التنظيف البالغة {len(clog)}. لكل صيغة طريقتها الخاصة في الإخفاق.")
ax = fig.add_axes((0.14, 0.1, 0.56, 0.68))
im = ax.imshow(mat.values, cmap=SEQ, vmin=0, vmax=mat.values.max(), aspect="auto")
for i in range(mat.shape[0]):
    for j in range(mat.shape[1]):
        v = mat.values[i, j]
        if v:
            ax.text(j, i, v, ha="center", va="center", fontsize=13, fontweight="bold",
                    color="white" if v > mat.values.max() * 0.45 else INK)
        else:
            ax.text(j, i, "-", ha="center", va="center", fontsize=11, color=MUTED)
ax.set_xticks(range(6), [A(f"{SHORT[s]}\n({FMT[s]})") for s in SRC_RTL], fontsize=10.5, color=INK2)
ax.xaxis.tick_top()
ax.yaxis.tick_right()
ax.set_yticks(range(len(FAMILY)), [A(FAM_AR[f]) for f in FAMILY], fontsize=11.5)
ax.set_xticks(np.arange(-.5, 6), minor=True)
ax.set_yticks(np.arange(-.5, len(FAMILY)), minor=True)
ax.grid(which="minor", color=SURFACE, lw=3)
ax.tick_params(which="both", length=0)
for sp in ax.spines.values():
    sp.set_visible(False)
tot = mat.sum(axis=1)
for i, v in enumerate(tot):
    ax.text(-0.62, i, f"{v}", fontsize=13, fontweight="bold", va="center", ha="right")
ax.text(-0.62, -0.95, A("المجموع"), fontsize=9.5, fontweight="bold", color=MUTED, ha="right", va="center")
save(fig, "03_anomaly_map.png")

# =============================================================================
# 04 - what was done about it: actions by source
# =============================================================================
ACTIONS = ["standardised", "corrected", "omitted", "missing", "flagged"]
ACOL = dict(zip(ACTIONS, [BLUE, ORANGE, AQUA, YELLOW, MAGENTA]))
ALAB = {"standardised": "موحَّد - الصيغة فقط", "corrected": "مصحَّح - قيمة أُصلحت أو حُوِّلت",
        "omitted": "مستبعَد - صف أو قيمة تُركت", "missing": "مفقود - غير موجود في المصدر",
        "flagged": "مُعلَّم - دون تغيير، أُحيل للمراجعة"}
act = clog.groupby(["source", "action"]).size().unstack(fill_value=0).reindex(index=SOURCES, columns=ACTIONS, fill_value=0)
fig = canvas("ما فعله خط المعالجة بكل مشكلة وجدها",
             f"{len(clog)} إجراءً مسجَّلًا. لا شيء يتغيّر بصمت - كل صف يحتفظ بالقيمة الخام بجانب القيمة النظيفة.")
ax = fig.add_axes((0.06, 0.2, 0.74, 0.6))
left = np.zeros(6)
for a in ACTIONS:
    v = act[a].values
    ax.barh(range(6), v, left=left, height=0.56, color=ACOL[a], edgecolor=SURFACE, linewidth=2, label=A(ALAB[a]))
    for i, x in enumerate(v):
        if x >= 9:
            ax.text(left[i] + x / 2, i, x, ha="center", va="center", fontsize=10.5, fontweight="bold",
                    color=INK if a in ("missing", "flagged", "omitted") else "white")
    left += v
for i, t in enumerate(left):
    ax.text(t + 2, i, f"{int(t)}", va="center", ha="right", fontsize=12, fontweight="bold")
ax.set_xlim(0, left.max() * 1.08)
ax.set_yticks(range(6), [A(f"{SHORT[s]}  ({FMT[s]})") for s in SOURCES], fontsize=11.5)
ax.invert_yaxis()
ax.set_xlabel(A("عدد المدخلات في سجل التنظيف"))
clean_axis(ax)
rtl_bars(ax)
h, l = ax.get_legend_handles_labels()
order = [2, 4, 1, 3, 0]  # columns fill top-down, left to right: read right-to-left, row by row
ax.legend([h[k] for k in order], [l[k] for k in order], loc="upper center", bbox_to_anchor=(0.58, -0.14), ncol=3,
          fontsize=10, handlelength=1, columnspacing=1.6, markerfirst=False)
save(fig, "04_actions_by_source.png")

# =============================================================================
# 05 - scale harmonisation
# =============================================================================
KINDS = ["Percent 0-100", "Fraction 0-1", "Out of 20", "GPA 0-4", "Text, letter or blank", "Impossible number"]
KINDS_AR = dict(zip(KINDS, ["نسبة مئوية 0-100", "كسر 0-1", "من 20", "GPA 0-4", "نص أو حرف أو فراغ", "رقم مستحيل"]))
KCOL = dict(zip(KINDS, [BLUE, ORANGE, AQUA, YELLOW, MAGENTA, GREEN]))


def notation(r):
    t = r.raw_value.split(" vs ")[0].strip()
    pct = t.endswith("%")
    v = pd.to_numeric(t.rstrip("%").replace(",", ".").split("/")[0], errors="coerce")
    if pd.isna(v):
        return KINDS[4]
    if r.subject in ("Math", "English"):
        return KINDS[1] if 0 < v <= 1 else (KINDS[0] if 0 <= v <= 100 else KINDS[5])
    if r.subject == "Science":
        if pct or (r.native_scale.startswith("fraction") and 10 < v <= 100):
            return KINDS[0]
        if r.native_scale.startswith("fraction"):
            return KINDS[1] if v <= 1 else KINDS[5]
        return KINDS[2] if v <= 20 else KINDS[5]
    return KINDS[3] if 0 <= v <= 4 else (KINDS[0] if 50 <= v <= 100 else KINDS[5])


staging["notation"] = staging.apply(notation, axis=1)
nt = staging.groupby(["subject", "notation"]).size().unstack(fill_value=0).reindex(index=SUBJ, columns=KINDS, fill_value=0)
fig = canvas("أربع مواد، وأربع طرق لكتابة الدرجة - ومقياس واحد في النهاية",
             "كيف كُتبت كل درجة في ملفها المصدر (يمينًا) وتوزيعها بعد التحويل إلى 0-100 (يسارًا).")
ax = fig.add_axes((0.47, 0.2, 0.39, 0.58))
left = np.zeros(4)
for k in KINDS:
    v = nt[k].values
    ax.barh(range(4), v, left=left, height=0.5, color=KCOL[k], edgecolor=SURFACE, linewidth=2, label=A(KINDS_AR[k]))
    for i, x in enumerate(v):
        if x >= 14:
            ax.text(left[i] + x / 2, i, x, ha="center", va="center", fontsize=10.5, fontweight="bold",
                    color=INK if k in (KINDS[3], KINDS[4], KINDS[2]) else "white")
    left += v
ax.set_yticks(range(4), [A(SUBJ_AR[s]) for s in SUBJ], fontsize=12)
ax.invert_yaxis()
ax.set_xlabel(A("الطلاب"))
ax.set_title(A("كما كُتبت في المصدر"), fontsize=12.5, fontweight="bold", loc="right", pad=12)
clean_axis(ax)
rtl_bars(ax)
h, l = ax.get_legend_handles_labels()
order = [2, 5, 1, 4, 0, 3]  # three columns of two, first kind top-right
ax.legend([h[k] for k in order], [l[k] for k in order], loc="upper right", bbox_to_anchor=(1.06, -0.14), ncol=3,
          fontsize=9.8, handlelength=1, columnspacing=1.4, markerfirst=False)
bins = np.arange(30, 102.5, 5)
for i, s in enumerate(SUBJ):
    a = fig.add_axes((0.05, 0.635 - i * 0.145, 0.33, 0.115))
    vals = grades[grades.subject == s].score_pct.dropna()
    a.hist(vals, bins=bins, color=BLUE, edgecolor=SURFACE, linewidth=1.5)
    a.set_yticks([])
    a.spines["left"].set_visible(False)
    a.set_xlim(30, 100)
    a.tick_params(length=0, labelbottom=(i == 3))
    a.text(0.28, 0.82, A(SUBJ_AR[s]), transform=a.transAxes, fontsize=11.5, fontweight="bold", ha="right")
    a.text(0.28, 0.58, A(f"{len(vals)} درجة\nالمتوسط {vals.mean():.1f}%"), transform=a.transAxes, fontsize=9.5,
           color=INK2, va="top", ha="right", linespacing=1.45)
    if i == 0:
        a.set_title(A("بعد التنظيف: مقياس واحد 0-100"), fontsize=12.5, fontweight="bold", loc="right", pad=12)
    if i == 3:
        a.set_xlabel(A("الدرجة (%)"))
save(fig, "05_scale_harmonisation.png")

# =============================================================================
# 06 - naive parse vs cleaned, scored against the answer key
# =============================================================================
OUTC = [("correct", "يطابق القيمة الحقيقية", GOOD), ("withheld", "فارغ - مفقود أو محجوب للمراجعة", WARN),
        ("wrong", "قيمة خاطئة في البيانات", CRIT)]
fig = canvas("هل ينجح التنظيف؟ التقييم مقابل مفتاح الإجابات المخفي",
             "كل درجة مادة قورنت بقيمتها الحقيقية. التحليل الساذج = الثقة بالمعرّفات والمقياس المُعلن، ثم التحويل، وانتهى.")
ax = fig.add_axes((0.1, 0.17, 0.68, 0.59))
ypos, labels = [], []
for i, s in enumerate(SUBJ):
    for j, (name, src) in enumerate([("تحليل ساذج", "grades_naive"), ("بعد التنظيف", "grades_clean")]):
        y = i * 2.6 + j
        ypos.append(y)
        labels.append(A(name))
        left = 0
        for k, lab, col in OUTC:
            v = ev[src][s][k]
            ax.barh(y, v, left=left, height=0.72, color=col, edgecolor=SURFACE, linewidth=2,
                    label=A(lab) if i == 0 and j == 0 else None)
            if v >= 9:
                ax.text(left + v / 2, y, v, ha="center", va="center", fontsize=10.5, fontweight="bold",
                        color=INK if k == "withheld" else "white")
            left += v
        w = ev[src][s]["wrong"]
        ax.text(243, y, A(f"{w} خاطئة"), va="center", ha="right", fontsize=10.5, fontweight="bold" if w else "normal",
                color=CRIT if w else INK2)
    ax.text(-62, i * 2.6 + 0.5, A(SUBJ_AR[s]), fontsize=13, fontweight="bold", va="center", ha="right")
ax.set_yticks(ypos, labels, fontsize=10.5)
ax.invert_yaxis()
ax.set_xlim(0, 240)
ax.set_xlabel(A("الطلاب (240 لكل مادة)"))
clean_axis(ax)
rtl_bars(ax)
h, l = ax.get_legend_handles_labels()
ax.legend(h[::-1], l[::-1], loc="upper center", bbox_to_anchor=(0.55, -0.13), ncol=3, fontsize=10.5, handlelength=1,
          markerfirst=False)
fig.text(0.96, 0.83, A(f"المواد الأربع: التحليل الساذج {naive['wrong']} خاطئة + {naive['withheld']} فارغة من 960.   "
                       f"بعد التنظيف: {clean['wrong']} خاطئة، {clean['withheld']} محجوبة للمراجعة، {clean['correct']} صحيحة."),
         fontsize=12, fontweight="bold", color=BLUE, ha="right")
save(fig, "06_naive_vs_cleaned.png")

# =============================================================================
# 07 - the review queue
# =============================================================================
HOW = {
    "no_form_response": "إعادة إرسال استمارة التسجيل عبر معلّم الفصل",
    "missing_checkout": "التأكد من أن الطالب خرج من بوابة مراقَبة",
    "missing_checkin": "التأكد من وقت الوصول من سجل الفصل",
    "scale_gpa": "مراجعة كشف درجات المدرسة السابقة (GPA مفترض)",
    "early_leave": "المطابقة مع أذونات الانصراف المبكر الموقّعة",
    "blank": "الاتصال بولي الأمر أو مراسلته لاستكمال بيانات التواصل",
    "row_missing": "طلب الدرجة أو السبب من معلّم المادة",
    "text_not_available": "يؤكد المعلّم الحالة النهائية (N/A، معفى، غير مكتمل)",
    "text_absent": "تحديد موعد امتحان تعويضي ثم إدخال الدرجة",
    "letter_grade": "طلب الدرجة الرقمية - لا قيمة دقيقة للحروف",
    "dob_missing": "إدخال التاريخ من نسخة شهادة الميلاد",
    "dob_illogical": "مراجعة شهادة الميلاد - غالبًا خطأ في السنة",
    "unknown_id": "تحديد حامل البطاقة أو التحقق من تسجيل متأخر",
    "odd_hour_swipe": "يراجع الأمن تسجيلات الكاميرات لتلك البوابة والوقت",
    "absent_but_swiped": "يؤكد معلّم الفصل حضور الطالب",
    "total_mismatch": "يراجع المعلّم كشف الدرجات الورقي",
    "duplicate_id_conflict": "يؤكد المسجّل المعرّف ويصحّح السجل",
    "duplicate_conflict": "يحدد المعلّم أي الصفين هو الصحيح",
    "weekend_swipe": "المقارنة مع تقويم الأنشطة",
    "name_mismatch": "سؤال ولي الأمر عن التهجئة الصحيحة",
    "illogical_value": "مراجعة الاستمارة الورقية الأصلية أو كشف الدرجات",
    "scale_percent_in_gpa_column": "يؤكد المعلّم قيمة GPA المحوَّلة",
    "misplaced_decimal": "يؤكد المعلّم القيمة المقترحة",
    "wrong_id": "تأكيد المعرّف المعاد ربطه مع ولي الأمر",
    "gpa_over_4": "يعيد معلّم التاريخ إدخال قيمة GPA",
    "email_invalid": "الاتصال بولي الأمر لتأكيد العنوان",
    "phone_invalid": "تأكيد الرقم عبر البريد الإلكتروني",
}
PCOL = {"high": CRIT, "medium": WARN, "low": MUTED}
rq = review.groupby(["category", "priority"]).size().reset_index(name="n").sort_values("n", ascending=False).head(12)
fig = canvas(f"{len(review)} مدخلًا تحتاج إلى شخص - ولكلٍّ منها طريقة للمراجعة",
             "خط المعالجة لا يخمّن. قائمة المراجعة تحدد ما يجب فحصه، وأين، ومدى إلحاحه.")
ax0 = blank_ax(fig, (0.04, 0.72, 0.92, 0.12))
pc = review.priority.value_counts()
PDESC = {"high": ("أولوية عالية", "السلامة، الهوية، الامتحانات الفائتة"), "medium": ("أولوية متوسطة", "القيمة غير صالحة حتى تُفحص"),
         "low": ("أولوية منخفضة", "القيمة محفوظة، تُؤكَّد لاحقًا")}
for i, p in enumerate(["high", "medium", "low"]):
    x = 100 - 32 - i * 34
    card(ax0, x, 0, 32, 100, color=SURFACE, edge=GRID, lw=1)
    ax0.add_patch(plt.Rectangle((x + 31, 0), 1.0, 100, fc=PCOL[p], ec="none"))
    ax0.text(x + 28.5, 50, f"{pc.get(p, 0)}", fontsize=26, fontweight="bold", va="center", ha="right")
    ax0.text(x + 19.5, 66, A(PDESC[p][0]), fontsize=10, fontweight="bold", va="center", ha="right")
    ax0.text(x + 19.5, 34, A(PDESC[p][1]), fontsize=9.6, color=INK2, va="center", ha="right")
ax = fig.add_axes((0.56, 0.09, 0.22, 0.57))
ax.barh(range(len(rq)), rq.n, height=0.6, color=[PCOL[p] for p in rq.priority])
for i, r in enumerate(rq.itertuples()):
    ax.text(r.n + 0.8, i, r.n, va="center", ha="right", fontsize=10.5, fontweight="bold")
    ax.text(rq.n.max() * 1.18, i, A(HOW.get(r.category, "")), va="center", ha="right", fontsize=10, color=INK2)
ax.set_yticks(range(len(rq)), [A(CAT_AR.get(c, c.replace("_", " "))) for c in rq.category], fontsize=10.5)
ax.invert_yaxis()
ax.set_xticks([])
ax.tick_params(length=0)
for sp in ax.spines.values():
    sp.set_visible(False)
rtl_bars(ax)
fig.text(0.96, 0.675, A("أكثر عناصر المراجعة شيوعًا"), fontsize=9.5, fontweight="bold", color=MUTED, ha="right")
fig.text(0.525, 0.675, A("طريقة المراجعة"), fontsize=9.5, fontweight="bold", color=MUTED, ha="right")
save(fig, "07_review_queue.png")

# =============================================================================
# 08 - academic results from the clean data
# =============================================================================
fig = canvas("ما تقوله البيانات النظيفة: النتائج الأكاديمية",
             f"{stats['grades_available']} درجة موثّقة على مقياس واحد. أصبحت المواد قابلة للمقارنة لأول مرة.")
SUBJ_RTL = SUBJ[::-1]  # first subject on the right
ax = fig.add_axes((0.07, 0.12, 0.42, 0.63))
data = [grades[grades.subject == s].score_pct.dropna() for s in SUBJ_RTL]
bp = ax.boxplot(data, widths=0.5, patch_artist=True, medianprops=dict(color=INK, lw=2),
                boxprops=dict(facecolor="#9ec5f4", edgecolor=BLUE, lw=1.5), whiskerprops=dict(color=BLUE, lw=1.5),
                capprops=dict(color=BLUE, lw=1.5), flierprops=dict(marker="o", ms=4, mfc=SURFACE, mec=BLUE))
for i, d in enumerate(data, 1):
    ax.text(i + 0.31, d.median(), f"{d.median():.0f}", fontsize=10.5, fontweight="bold", va="center")
ax.set_xticks(range(1, 5), [A(SUBJ_AR[s]) for s in SUBJ_RTL], fontsize=11.5, color=INK2)
ax.set_ylabel(A("الدرجة (%)"))
ax.set_title(A("توزّع الدرجات حسب المادة (مع إظهار الوسيط)"), fontsize=12.5, fontweight="bold", loc="right", pad=12)
clean_axis(ax, "y")
COLS_RTL = [c.lower() for c in SUBJ_RTL]
hm = ov.groupby("grade")[COLS_RTL].mean()
ax = fig.add_axes((0.56, 0.2, 0.32, 0.5))
ax.imshow(hm.values, cmap=SEQ, aspect="auto", vmin=hm.values.min() - 1.5, vmax=hm.values.max() + 0.5)
for i in range(4):
    for j in range(4):
        v = hm.values[i, j]
        ax.text(j, i, f"{v:.1f}", ha="center", va="center", fontsize=13, fontweight="bold",
                color="white" if v > hm.values.mean() + 0.3 else INK)
ax.set_xticks(range(4), [A(SUBJ_AR[s]) for s in SUBJ_RTL], fontsize=11, color=INK2)
ax.xaxis.tick_top()
ax.yaxis.tick_right()
ax.set_yticks(range(4), [A(f"الصف {g}") for g in hm.index], fontsize=11.5)
ax.set_xticks(np.arange(-.5, 4), minor=True)
ax.set_yticks(np.arange(-.5, 4), minor=True)
ax.grid(which="minor", color=SURFACE, lw=3)
ax.tick_params(which="both", length=0)
for sp in ax.spines.values():
    sp.set_visible(False)
fig.text(0.96, 0.79, A("متوسط الدرجة حسب المستوى الدراسي (%)"), fontsize=12.5, fontweight="bold", ha="right")
lo = hm.mean(axis=1).idxmin()
fig.text(0.96, 0.14, A(f"الصف {lo} صاحب أدنى متوسط عبر المواد ({hm.mean(axis=1).min():.1f}%)."), fontsize=10.5,
         color=INK2, ha="right")
save(fig, "08_academic_results.png")

# =============================================================================
# 09 - attendance: the log as a data source
# =============================================================================
d = ov.dropna(subset=["average_pct"])
r = np.corrcoef(d.attendance_rate, d.average_pct)[0, 1]
slope, icpt = np.polyfit(d.attendance_rate * 100, d.average_pct, 1)
fig = canvas("ما يضيفه سجل البوابة: الحضور والأحداث المعلَّمة",
             f"{stats['log']['lines']:,} سطرًا في السجل أصبحت {stats['students'] * stats['log']['school_days']:,} يوم حضور للطلاب و {len(flags)} مدخلًا معلَّمًا.")
ax = fig.add_axes((0.46, 0.12, 0.44, 0.63))
jit = np.random.default_rng(1).normal(0, 0.35, len(d))
ax.scatter(d.attendance_rate * 100 + jit, d.average_pct, s=34, color=BLUE, edgecolor=SURFACE, linewidth=0.9, alpha=0.9)
xs = np.array([d.attendance_rate.min() * 100, 100])
ax.plot(xs, slope * xs + icpt, color=INK, lw=2)
ax.text(0.03, 0.95, f"r = {r:.2f}", transform=ax.transAxes, fontsize=15, fontweight="bold", va="top")
trend = "زيادة" if slope >= 0 else "انخفاض"
ax.text(0.03, 0.875, A(f"كل 10 نقاط إضافية في الحضور ≈ {trend} {abs(slope * 10):.1f} نقطة في متوسط الدرجات"),
        transform=ax.transAxes, fontsize=10, color=INK2, va="top")
ax.set_xlabel(A("نسبة الحضور خلال 20 يومًا دراسيًا (%)"))
ax.set_ylabel(A("متوسط الدرجات عبر المواد (%)"))
ax.yaxis.tick_right()
ax.yaxis.set_label_position("right")
ax.spines["left"].set_visible(False)
ax.spines["right"].set_visible(True)
ax.set_title(A("الطلاب الأكثر حضورًا يحققون درجات أعلى"), fontsize=12.5, fontweight="bold", loc="right", pad=12)
clean_axis(ax, "both")
fc = flags.category.value_counts()
ax = fig.add_axes((0.05, 0.12, 0.17, 0.63))
ax.barh(range(len(fc)), fc.values, height=0.6, color=[PCOL[flags[flags.category == c].priority.iloc[0]] for c in fc.index])
for i, v in enumerate(fc.values):
    ax.text(v + 0.5, i, v, va="center", ha="right", fontsize=10.5, fontweight="bold")
ax.set_yticks(range(len(fc)), [A(CAT_AR.get(c, c.replace("_", " "))) for c in fc.index], fontsize=10.5)
ax.invert_yaxis()
ax.set_xticks([])
ax.tick_params(length=0)
for sp in ax.spines.values():
    sp.set_visible(False)
rtl_bars(ax)
fig.text(0.37, 0.785, A("مدخلات السجل المعلَّمة للمتابعة"), fontsize=12.5, fontweight="bold", ha="right")
for i, (p, lab) in enumerate([("high", "أولوية عالية"), ("medium", "متوسطة"), ("low", "منخفضة")]):
    x = 0.37 - i * 0.115
    fig.text(x, 0.065, "■", color=PCOL[p], fontsize=12, va="center", ha="right")
    fig.text(x - 0.017, 0.065, A(lab), fontsize=10, color=INK2, va="center", ha="right")
save(fig, "09_attendance_and_log_flags.png")

# =============================================================================
# 10 - conclusions
# =============================================================================
n_open = int((ov.student_id.isin(review.record)).sum())
det = ev["detection_total"]
worst = max(SUBJ, key=lambda s: ev["grades_naive"][s]["wrong"])
finds = [
    (f"{naive['wrong'] + naive['withheld']}",
     "درجة من 960 كانت ستكون خاطئة أو مفقودة دون تنظيف",
     f"التحليل الساذج يخطئ في {naive['wrong']} درجة ويُسقط {naive['withheld']}. مادة {SUBJ_AR[worst]} هي الأسوأ\n"
     f"({ev['grades_naive'][worst]['wrong']} خاطئة): ثلاثة مقاييس تتشارك عمودًا واحدًا."),
    (f"{N_WRONG}",
     f"قيمة خاطئة من {N_CHECK:,} بعد التنظيف، مقارنةً بالمفتاح",
     f"{N_OK:,} قيمة تطابق الحقيقة تمامًا؛ و {N_HELD} حُجبت عمدًا\nلأن المصدر لم يقدّم إجابة آمنة."),
    (f"{det['handled']}/{det['injected']}",
     "خطأ مزروع التقطها خط المعالجة وسجّلها",
     "الاستثناء الوحيد اسم مكتوب خطأً دون معرّف: بعيد جدًا عن أي اسم\nفي السجل لربطه بأمان، فأُحيل إلى المراجعة بدلًا من ذلك."),
    (f"{len(review)}",
     "عنصر مراجعة، لكلٍّ منها أولوية وطريقة",
     f"{pc.get('high', 0)} بأولوية عالية (تمريرات أمنية، معرّفات مشتركة، امتحانات فائتة).\n"
     f"{n_open} من 240 طالبًا لديهم عنصر مفتوح واحد على الأقل."),
    (f"r = {r:.2f}",
     "الحضور والتحصيل يتحركان معًا",
     "لا يظهر ذلك إلا بعد أن يتشارك السجل ومصادر الدرجات الثلاثة\nوسجل الطلاب معرّفًا واحدًا للطالب ومقياسًا واحدًا."),
    ("4 ← 1",
     "مقاييس: 0-1 ومن 20 و GPA 4.0 والنسبة المئوية صارت 0-100",
     "يُقرأ المقياس من ترويسة كل جدول ويُتحقَّق منه مقابل\nنطاقات القيم - ولا يُفترض أبدًا لملف كامل."),
]
fig = canvas("الاستنتاجات", "ما يكشفه المشروع عن بيانات المدارس الفوضوية - وعن المدرسة نفسها.")
ax = blank_ax(fig)
for i, (big, head, body) in enumerate(finds):
    x, y = 51 - (i % 2) * 51, 68 - (i // 2) * 33.5
    card(ax, x, y, 49, 30)
    ax.text(x + 46.5, y + 22.5, A(big), fontsize=27, fontweight="bold", va="center", ha="right", color=BLUE)
    ax.text(x + 46.5, y + 14, A(head), fontsize=11.8, fontweight="bold", va="center", ha="right")
    ax.text(x + 46.5, y + 6.3, A(body), fontsize=9.6, color=INK2, va="center", ha="right", linespacing=1.5)
save(fig, "10_conclusions.png")

print(json.dumps(dict(checked=N_CHECK, ok=N_OK, held=N_HELD, wrong=N_WRONG, naive=naive, clean=clean, r=round(r, 3),
                      slope10=round(slope * 10, 2), open_students=n_open, review=pc.to_dict(), log=len(clog),
                      grade_means=hm.mean(axis=1).round(1).to_dict()), indent=1))
