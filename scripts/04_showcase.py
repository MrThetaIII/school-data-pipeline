"""
Phase 3 - Render the ten showcase images (showcase/01..10 *.png).

Reads only the outputs of the earlier steps: clean_data/school.db, clean_data/staging_grades.csv,
clean_data/pipeline_stats.json, reports/evaluation.json and the answer-key counts.
"""
import json
import sqlite3
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.colors import LinearSegmentedColormap
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch

ROOT = Path(__file__).resolve().parents[1]
SHOW = ROOT / "showcase"
SHOW.mkdir(exist_ok=True)

# ---- palette (validated reference palette, light surface) -------------------
SURFACE, INK, INK2, MUTED, GRID, AXIS = "#fcfcfb", "#0b0b0b", "#52514e", "#898781", "#e1e0d9", "#c3c2b7"
BLUE, ORANGE, AQUA, YELLOW, MAGENTA, GREEN = "#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300"
GOOD, WARN, CRIT = "#0ca30c", "#fab219", "#d03b3b"
CARD = "#f1f0ec"
SEQ = LinearSegmentedColormap.from_list("blue", ["#cde2fb", "#86b6ef", "#3987e5", "#1c5cab", "#0d366b"])
plt.rcParams.update({
    "font.family": ["Segoe UI", "DejaVu Sans"], "font.size": 11, "text.color": INK, "axes.labelcolor": INK2,
    "xtick.color": MUTED, "ytick.color": INK2, "axes.edgecolor": AXIS, "figure.facecolor": SURFACE,
    "axes.facecolor": SURFACE, "savefig.facecolor": SURFACE, "axes.spines.top": False, "axes.spines.right": False,
    "axes.grid": False, "legend.frameon": False})

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
SHORT = {SOURCES[0]: "Student roster", SOURCES[1]: "Enrollment form", SOURCES[2]: "Gradebook",
         SOURCES[3]: "Science results", SOURCES[4]: "History report", SOURCES[5]: "Gate log"}
FMT = {SOURCES[0]: "Excel", SOURCES[1]: "CSV", SOURCES[2]: "Excel", SOURCES[3]: "Word", SOURCES[4]: "PDF", SOURCES[5]: "Log"}
SUBJ = ["Math", "English", "Science", "History"]
checked = [t for block in ("student_fields", "grades_clean", "attendance", "guardians") for t in ev[block].values()]
N_CHECK = sum(sum(t.values()) for t in checked)
N_OK, N_HELD, N_WRONG = (sum(t[k] for t in checked) for k in ("correct", "withheld", "wrong"))
naive = {k: sum(t[k] for t in ev["grades_naive"].values()) for k in ("correct", "withheld", "wrong")}
clean = {k: sum(t[k] for t in ev["grades_clean"].values()) for k in ("correct", "withheld", "wrong")}


def frame(fig, title, subtitle, note="Greenfield Academy is fictional - all data is synthetic."):
    fig.text(0.04, 0.945, title, fontsize=21, fontweight="bold", va="center")
    fig.text(0.04, 0.895, subtitle, fontsize=12.5, color=INK2, va="center")
    fig.text(0.04, 0.025, note, fontsize=9, color=MUTED, va="center")


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
    fig.savefig(SHOW / name)
    plt.close(fig)
    print("wrote", name)


def clean_axis(ax, grid="x"):
    ax.grid(axis=grid, color=GRID, lw=0.8)
    ax.set_axisbelow(True)
    ax.tick_params(length=0)


# =============================================================================
# 01 - project overview
# =============================================================================
fig = canvas("From six messy files to one trusted school database",
             "Synthetic data for a 240-student school, deliberately broken, then parsed, cleaned and audited.")
ax = blank_ax(fig)
ax.text(0, 97, "1  SYNTHESIZE", fontsize=10.5, fontweight="bold", color=MUTED)
ax.text(36, 97, "2  PARSE AND CLEAN", fontsize=10.5, fontweight="bold", color=MUTED)
ax.text(72, 97, "3  COMPILE", fontsize=10.5, fontweight="bold", color=MUTED)
for i, s in enumerate(SOURCES):
    y = 84 - i * 10.6
    card(ax, 0, y, 28, 9)
    ax.text(1.5, y + 4.5, FMT[s], fontsize=10, fontweight="bold", color="white", va="center",
            bbox=dict(boxstyle="round,pad=0.35", fc=BLUE, ec="none"))
    ax.text(9, y + 5.9, SHORT[s], fontsize=11.5, fontweight="bold", va="center")
    ax.text(9, y + 2.7, f"{stats['sources'][s]['records_read']:,} records - {(key.source == s).sum()} planted errors",
            fontsize=9.5, color=INK2, va="center")
card(ax, 36, 31, 28, 62, color="#e6eefa")
steps = [("Locate the data", "header rows, merged cells, page\nbreaks, log noise"),
         ("Standardise notation", "IDs, names, dates, decimals,\nphone numbers"),
         ("Unify scales", "0-1, /20, GPA 4.0 and % -> one\n0-100 scale"),
         ("Cross-check sources", "fill and verify one file against\nanother"),
         ("Refuse to guess", "impossible values are withheld\nand queued for a person")]
for i, (h, d) in enumerate(steps):
    y = 86 - i * 11.6
    ax.text(38, y, h, fontsize=11.5, fontweight="bold", va="center")
    ax.text(38, y - 4.6, d, fontsize=9.3, color=INK2, va="center", linespacing=1.25)
outs = [("SQLite database", "school.db - 9 tables + 1 view,\nkeys and range checks"),
        ("Excel workbook", "school_clean.xlsx - 11 sheets,\nfiltered and colour-coded"),
        ("Cleaning report", f"{len(clog)} logged actions,\n{len(review)} review items with a method")]
for i, (h, d) in enumerate(outs):
    y = 73 - i * 21
    card(ax, 72, y, 28, 18)
    ax.text(74, y + 13, h, fontsize=12, fontweight="bold", va="center")
    ax.text(74, y + 6, d, fontsize=9.5, color=INK2, va="center", linespacing=1.3)
for x0, x1 in [(28.6, 35.4), (64.6, 71.4)]:
    ax.add_patch(FancyArrowPatch((x0, 62), (x1, 62), arrowstyle="-|>", mutation_scale=18, color=MUTED, lw=1.6))
tiles = [(f"{stats['students']}", "students"), ("6", "source files,\n5 formats"), (f"{len(key)}", "errors planted"),
         (f"{len(clog)}", "cleaning actions\nlogged"), (f"{len(review)}", "items sent\nfor review"),
         (f"{N_WRONG}", f"wrong values in\n{N_CHECK:,} checked")]
for i, (big, small) in enumerate(tiles):
    x = i * 16.9
    card(ax, x, 0, 15.5, 21, color=SURFACE, edge=GRID, lw=1)
    ax.text(x + 7.75, 14, big, fontsize=25, fontweight="bold", ha="center", va="center",
            color=GOOD if i == 5 else INK)
    ax.text(x + 7.75, 5.2, small, fontsize=9.3, color=INK2, ha="center", va="center", linespacing=1.2)
save(fig, "01_project_overview.png")

# =============================================================================
# 02 - the six sources and what was wrong with each
# =============================================================================
PROBLEMS = {
    SOURCES[0]: "Title rows above the header, a different column order on one sheet, dates as\ntext in five notations, impossible birth years, two students sharing one ID",
    SOURCES[1]: "Free-text answers: grade as \"Tenth\", averages as 88 / 88% / 0.88 / 3.4 / B+,\nsiblings as \"two\", resubmissions, test entries, broken e-mails and phones",
    SOURCES[2]: "\"absent\" and \"N/A\" in mark cells, numbers stored as text, totals above 100,\ntotals that do not equal coursework + exam, one class typed as 0-1",
    SOURCES[3]: "Scores on a 0-1 scale but Grade 11 out of 20, stray percentages, comma\ndecimals, letter grades, rows with a name but no ID, letter O typed for zero",
    SOURCES[4]: "GPA on a 4.0 scale, values above 4, percentages typed in the GPA column,\nheader repeated on every page, repeated rows, IDs split by spaces",
    SOURCES[5]: "Double swipes, unknown cards, swipes at 2 a.m. and on weekends, cut-off\nlines, a gate writing day-first timestamps, ABSENT entries that contradict swipes",
}
fig = canvas("Six sources, five formats, and the errors each one is known for",
             "What was generated in Phase 1. Every planted error is recorded in a hidden answer key.")
ax = blank_ax(fig)
for x, t in [(0, "SOURCE"), (25, "RECORDS"), (35.5, "PLANTED ERRORS"), (55, "TYPICAL PROBLEMS IN THIS FORMAT")]:
    ax.text(x, 97.5, t, fontsize=9.5, fontweight="bold", color=MUTED, va="center")
mx = max((key.source == s).sum() for s in SOURCES)
for i, s in enumerate(SOURCES):
    y = 80.5 - i * 15.6
    card(ax, 0, y, 100, 13.4)
    ax.text(1.5, y + 6.7, FMT[s], fontsize=10, fontweight="bold", color="white", va="center",
            bbox=dict(boxstyle="round,pad=0.35", fc=BLUE, ec="none"))
    ax.text(8.5, y + 8.6, SHORT[s], fontsize=12, fontweight="bold", va="center")
    ax.text(8.5, y + 4.4, s.split("_", 1)[1], fontsize=7.3, color=MUTED, va="center")
    ax.text(25, y + 6.7, f"{stats['sources'][s]['records_read']:,}", fontsize=14, fontweight="bold", va="center")
    n = (key.source == s).sum()
    ax.add_patch(plt.Rectangle((35.5, y + 5.2), 13 * n / mx, 3, fc=ORANGE, ec="none"))
    ax.text(35.5 + 13 * n / mx + 1, y + 6.7, str(n), fontsize=11, fontweight="bold", va="center")
    ax.text(55, y + 6.7, PROBLEMS[s], fontsize=9.2, color=INK2, va="center", linespacing=1.35)
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
cat2fam = {c: f for f, cs in FAMILY.items() for c in cs}
clog["family"] = clog.category.map(cat2fam).fillna("Invalid or illogical value")
mat = clog.groupby(["family", "source"]).size().unstack(fill_value=0).reindex(index=list(FAMILY), columns=SOURCES, fill_value=0)
fig = canvas("Where the problems were: issue type by source",
             f"All {len(clog)} entries of the cleaning log. Each format fails in its own way.")
ax = fig.add_axes((0.25, 0.1, 0.56, 0.68))
im = ax.imshow(mat.values, cmap=SEQ, vmin=0, vmax=mat.values.max(), aspect="auto")
for i in range(mat.shape[0]):
    for j in range(mat.shape[1]):
        v = mat.values[i, j]
        if v:
            ax.text(j, i, v, ha="center", va="center", fontsize=13, fontweight="bold",
                    color="white" if v > mat.values.max() * 0.45 else INK)
        else:
            ax.text(j, i, "-", ha="center", va="center", fontsize=11, color=MUTED)
ax.set_xticks(range(6), [f"{SHORT[s]}\n({FMT[s]})" for s in SOURCES], fontsize=10.5, color=INK2)
ax.xaxis.tick_top()
ax.set_yticks(range(len(FAMILY)), list(FAMILY), fontsize=11.5)
ax.set_xticks(np.arange(-.5, 6), minor=True)
ax.set_yticks(np.arange(-.5, len(FAMILY)), minor=True)
ax.grid(which="minor", color=SURFACE, lw=3)
ax.tick_params(which="both", length=0)
for sp in ax.spines.values():
    sp.set_visible(False)
tot = mat.sum(axis=1)
for i, v in enumerate(tot):
    ax.text(5.75, i, f"{v}", fontsize=13, fontweight="bold", va="center", ha="left")
ax.text(5.75, -0.95, "TOTAL", fontsize=9.5, fontweight="bold", color=MUTED, ha="left", va="center")
save(fig, "03_anomaly_map.png")

# =============================================================================
# 04 - what was done about it: actions by source
# =============================================================================
ACTIONS = ["standardised", "corrected", "omitted", "missing", "flagged"]
ACOL = dict(zip(ACTIONS, [BLUE, ORANGE, AQUA, YELLOW, MAGENTA]))
ADESC = {"standardised": "format only", "corrected": "value repaired or converted", "omitted": "row or value left out",
         "missing": "absent from the source", "flagged": "unchanged, queued for review"}
act = clog.groupby(["source", "action"]).size().unstack(fill_value=0).reindex(index=SOURCES, columns=ACTIONS, fill_value=0)
fig = canvas("What the pipeline did with every problem it found",
             f"{len(clog)} logged actions. Nothing is changed silently - each row keeps the raw value beside the clean one.")
ax = fig.add_axes((0.2, 0.2, 0.74, 0.6))
left = np.zeros(6)
for a in ACTIONS:
    v = act[a].values
    ax.barh(range(6), v, left=left, height=0.56, color=ACOL[a], edgecolor=SURFACE, linewidth=2, label=f"{a.capitalize()} - {ADESC[a]}")
    for i, x in enumerate(v):
        if x >= 9:
            ax.text(left[i] + x / 2, i, x, ha="center", va="center", fontsize=10.5, fontweight="bold",
                    color=INK if a in ("missing", "flagged", "omitted") else "white")
    left += v
for i, t in enumerate(left):
    ax.text(t + 2, i, f"{int(t)}", va="center", fontsize=12, fontweight="bold")
ax.set_yticks(range(6), [f"{SHORT[s]}  ({FMT[s]})" for s in SOURCES], fontsize=11.5)
ax.invert_yaxis()
ax.set_xlabel("Entries in the cleaning log")
clean_axis(ax)
ax.spines["left"].set_visible(False)
ax.legend(loc="upper center", bbox_to_anchor=(0.42, -0.14), ncol=3, fontsize=10, handlelength=1, columnspacing=1.6)
save(fig, "04_actions_by_source.png")

# =============================================================================
# 05 - scale harmonisation
# =============================================================================
KINDS = ["Percent 0-100", "Fraction 0-1", "Out of 20", "GPA 0-4", "Text, letter or blank", "Impossible number"]
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
fig = canvas("Four subjects, four ways of writing a score - one scale in the end",
             "How each mark was typed in its source file (left) and the distribution after conversion to 0-100 (right).")
ax = fig.add_axes((0.11, 0.2, 0.42, 0.58))
left = np.zeros(4)
for k in KINDS:
    v = nt[k].values
    ax.barh(range(4), v, left=left, height=0.5, color=KCOL[k], edgecolor=SURFACE, linewidth=2, label=k)
    for i, x in enumerate(v):
        if x >= 14:
            ax.text(left[i] + x / 2, i, x, ha="center", va="center", fontsize=10.5, fontweight="bold",
                    color=INK if k in (KINDS[3], KINDS[4], KINDS[2]) else "white")
    left += v
ax.set_yticks(range(4), SUBJ, fontsize=12)
ax.invert_yaxis()
ax.set_xlabel("Students")
ax.set_title("As typed in the source", fontsize=12.5, fontweight="bold", loc="left", pad=12)
clean_axis(ax)
ax.spines["left"].set_visible(False)
ax.legend(loc="upper left", bbox_to_anchor=(-0.2, -0.16), ncol=3, fontsize=9.8, handlelength=1, columnspacing=1.4)
bins = np.arange(30, 102.5, 5)
for i, s in enumerate(SUBJ):
    a = fig.add_axes((0.62, 0.635 - i * 0.145, 0.33, 0.115))
    vals = grades[grades.subject == s].score_pct.dropna()
    a.hist(vals, bins=bins, color=BLUE, edgecolor=SURFACE, linewidth=1.5)
    a.set_yticks([])
    a.spines["left"].set_visible(False)
    a.set_xlim(30, 100)
    a.tick_params(length=0, labelbottom=(i == 3))
    a.text(0.0, 0.82, s, transform=a.transAxes, fontsize=11.5, fontweight="bold")
    a.text(0.0, 0.58, f"{len(vals)} scores\nmean {vals.mean():.1f}%", transform=a.transAxes, fontsize=9.5, color=INK2, va="top", linespacing=1.3)
    if i == 0:
        a.set_title("After cleaning: one 0-100 scale", fontsize=12.5, fontweight="bold", loc="left", pad=12)
    if i == 3:
        a.set_xlabel("Score (%)")
save(fig, "05_scale_harmonisation.png")

# =============================================================================
# 06 - naive parse vs cleaned, scored against the answer key
# =============================================================================
OUTC = [("correct", "Matches the true value", GOOD), ("withheld", "Blank - lost or withheld for review", WARN),
        ("wrong", "Wrong value in the dataset", CRIT)]
fig = canvas("Does the cleaning work? Scored against the hidden answer key",
             "Every subject score compared with the true value. Naive = trust the IDs and the declared scale, convert, done.")
ax = fig.add_axes((0.2, 0.17, 0.72, 0.59))
ypos, labels = [], []
for i, s in enumerate(SUBJ):
    for j, (name, src) in enumerate([("naive parse", "grades_naive"), ("cleaned", "grades_clean")]):
        y = i * 2.6 + j
        ypos.append(y)
        labels.append(name)
        left = 0
        for k, lab, col in OUTC:
            v = ev[src][s][k]
            ax.barh(y, v, left=left, height=0.72, color=col, edgecolor=SURFACE, linewidth=2,
                    label=lab if i == 0 and j == 0 else None)
            if v >= 9:
                ax.text(left + v / 2, y, v, ha="center", va="center", fontsize=10.5, fontweight="bold",
                        color=INK if k == "withheld" else "white")
            left += v
        w = ev[src][s]["wrong"]
        ax.text(243, y, f"{w} wrong", va="center", fontsize=10.5, fontweight="bold" if w else "normal",
                color=CRIT if w else INK2)
    ax.text(-52, i * 2.6 + 0.5, s, fontsize=13, fontweight="bold", va="center", ha="left")
ax.set_yticks(ypos, labels, fontsize=10.5)
ax.invert_yaxis()
ax.set_xlim(0, 240)
ax.set_xlabel("Students (240 per subject)")
clean_axis(ax)
ax.spines["left"].set_visible(False)
ax.legend(loc="upper center", bbox_to_anchor=(0.45, -0.13), ncol=3, fontsize=10.5, handlelength=1)
fig.text(0.04, 0.83, f"All four subjects: naive parse {naive['wrong']} wrong + {naive['withheld']} blank of 960.   "
                     f"Cleaned: {clean['wrong']} wrong, {clean['withheld']} withheld for review, {clean['correct']} correct.",
         fontsize=12, fontweight="bold", color=BLUE)
save(fig, "06_naive_vs_cleaned.png")

# =============================================================================
# 07 - the review queue
# =============================================================================
HOW = {
    "no_form_response": "Re-send the enrollment form through the homeroom teacher",
    "missing_checkout": "Confirm the student left through a monitored gate",
    "missing_checkin": "Confirm arrival time from the class register",
    "scale_gpa": "Check the previous school's transcript (GPA assumed)",
    "early_leave": "Match against signed early-dismissal slips",
    "blank": "Phone or e-mail the guardian for the missing contact",
    "row_missing": "Ask the subject teacher for the mark or the reason",
    "text_not_available": "Teacher confirms final status (N/A, exempt, incomplete)",
    "text_absent": "Schedule a make-up exam, then enter the mark",
    "letter_grade": "Ask for the numeric mark - letters have no exact value",
    "dob_missing": "Enter the date from the birth-certificate copy",
    "dob_illogical": "Check the birth certificate - likely a typo in the year",
    "unknown_id": "Identify the card holder or check for a late enrolment",
    "odd_hour_swipe": "Security reviews CCTV for that gate and time",
    "absent_but_swiped": "Homeroom teacher confirms presence",
    "total_mismatch": "Teacher checks the paper mark sheet",
    "duplicate_id_conflict": "Registrar confirms the ID and fixes the register",
    "duplicate_conflict": "Teacher says which of the two rows is valid",
    "weekend_swipe": "Compare with the activities calendar",
    "name_mismatch": "Ask the parent which spelling is correct",
    "illogical_value": "Check the original paper form or mark sheet",
    "scale_percent_in_gpa_column": "Teacher confirms the converted GPA",
    "misplaced_decimal": "Teacher confirms the suggested value",
    "wrong_id": "Confirm the re-linked ID with the parent",
    "gpa_over_4": "History teacher re-enters the GPA",
    "email_invalid": "Phone the guardian to confirm the address",
    "phone_invalid": "Confirm the number by e-mail",
}
PCOL = {"high": CRIT, "medium": WARN, "low": MUTED}
rq = review.groupby(["category", "priority"]).size().reset_index(name="n").sort_values("n", ascending=False).head(12)
fig = canvas(f"{len(review)} entries need a person - each one comes with a method",
             "The pipeline does not guess. The review queue says what to check, where, and how urgently.")
ax0 = blank_ax(fig, (0.04, 0.72, 0.92, 0.12))
pc = review.priority.value_counts()
PDESC = {"high": ("HIGH PRIORITY", "safety, identity, missed exams"), "medium": ("MEDIUM", "value unusable until checked"),
         "low": ("LOW", "value kept, confirm later")}
for i, p in enumerate(["high", "medium", "low"]):
    x = i * 34
    card(ax0, x, 0, 32, 100, color=SURFACE, edge=GRID, lw=1)
    ax0.add_patch(plt.Rectangle((x, 0), 1.0, 100, fc=PCOL[p], ec="none"))
    ax0.text(x + 3.5, 50, f"{pc.get(p, 0)}", fontsize=26, fontweight="bold", va="center")
    ax0.text(x + 12.5, 66, PDESC[p][0], fontsize=10, fontweight="bold", va="center")
    ax0.text(x + 12.5, 34, PDESC[p][1], fontsize=9.6, color=INK2, va="center")
ax = fig.add_axes((0.2, 0.09, 0.3, 0.57))
ax.barh(range(len(rq)), rq.n, height=0.6, color=[PCOL[p] for p in rq.priority])
for i, r in enumerate(rq.itertuples()):
    ax.text(r.n + 0.8, i, r.n, va="center", fontsize=10.5, fontweight="bold")
    ax.text(rq.n.max() * 1.18, i, HOW.get(r.category, ""), va="center", fontsize=10, color=INK2)
ax.set_yticks(range(len(rq)), [c.replace("_", " ") for c in rq.category], fontsize=10.5)
ax.invert_yaxis()
ax.set_xticks([])
ax.tick_params(length=0)
for sp in ax.spines.values():
    sp.set_visible(False)
fig.text(0.2, 0.675, "MOST COMMON REVIEW ITEMS", fontsize=9.5, fontweight="bold", color=MUTED)
fig.text(0.535, 0.675, "HOW TO REVIEW", fontsize=9.5, fontweight="bold", color=MUTED)
save(fig, "07_review_queue.png")

# =============================================================================
# 08 - academic results from the clean data
# =============================================================================
fig = canvas("What the clean data says: academic results",
             f"{stats['grades_available']} verified scores on one scale. Subjects are comparable for the first time.")
ax = fig.add_axes((0.07, 0.12, 0.42, 0.63))
data = [grades[grades.subject == s].score_pct.dropna() for s in SUBJ]
bp = ax.boxplot(data, widths=0.5, patch_artist=True, medianprops=dict(color=INK, lw=2),
                boxprops=dict(facecolor="#9ec5f4", edgecolor=BLUE, lw=1.5), whiskerprops=dict(color=BLUE, lw=1.5),
                capprops=dict(color=BLUE, lw=1.5), flierprops=dict(marker="o", ms=4, mfc=SURFACE, mec=BLUE))
for i, d in enumerate(data, 1):
    ax.text(i + 0.31, d.median(), f"{d.median():.0f}", fontsize=10.5, fontweight="bold", va="center")
ax.set_xticks(range(1, 5), SUBJ, fontsize=11.5, color=INK2)
ax.set_ylabel("Score (%)")
ax.set_title("Spread of scores by subject (median labelled)", fontsize=12.5, fontweight="bold", loc="left", pad=12)
clean_axis(ax, "y")
hm = ov.groupby("grade")[["math", "english", "science", "history"]].mean()
ax = fig.add_axes((0.6, 0.2, 0.34, 0.5))
ax.imshow(hm.values, cmap=SEQ, aspect="auto", vmin=hm.values.min() - 1.5, vmax=hm.values.max() + 0.5)
for i in range(4):
    for j in range(4):
        v = hm.values[i, j]
        ax.text(j, i, f"{v:.1f}", ha="center", va="center", fontsize=13, fontweight="bold",
                color="white" if v > hm.values.mean() + 0.3 else INK)
ax.set_xticks(range(4), SUBJ, fontsize=11, color=INK2)
ax.xaxis.tick_top()
ax.set_yticks(range(4), [f"Grade {g}" for g in hm.index], fontsize=11.5)
ax.set_xticks(np.arange(-.5, 4), minor=True)
ax.set_yticks(np.arange(-.5, 4), minor=True)
ax.grid(which="minor", color=SURFACE, lw=3)
ax.tick_params(which="both", length=0)
for sp in ax.spines.values():
    sp.set_visible(False)
fig.text(0.6, 0.79, "Mean score by grade level (%)", fontsize=12.5, fontweight="bold")
lo = hm.mean(axis=1).idxmin()
fig.text(0.6, 0.14, f"Grade {lo} has the lowest mean across subjects ({hm.mean(axis=1).min():.1f}%).", fontsize=10.5, color=INK2)
save(fig, "08_academic_results.png")

# =============================================================================
# 09 - attendance: the log as a data source
# =============================================================================
d = ov.dropna(subset=["average_pct"])
r = np.corrcoef(d.attendance_rate, d.average_pct)[0, 1]
slope, icpt = np.polyfit(d.attendance_rate * 100, d.average_pct, 1)
fig = canvas("What the gate log adds: attendance and flagged events",
             f"{stats['log']['lines']:,} log lines became {stats['students'] * stats['log']['school_days']:,} student-days of attendance and {len(flags)} flagged entries.")
ax = fig.add_axes((0.07, 0.12, 0.44, 0.63))
jit = np.random.default_rng(1).normal(0, 0.35, len(d))
ax.scatter(d.attendance_rate * 100 + jit, d.average_pct, s=34, color=BLUE, edgecolor=SURFACE, linewidth=0.9, alpha=0.9)
xs = np.array([d.attendance_rate.min() * 100, 100])
ax.plot(xs, slope * xs + icpt, color=INK, lw=2)
ax.text(0.03, 0.95, f"r = {r:.2f}", transform=ax.transAxes, fontsize=15, fontweight="bold", va="top")
ax.text(0.03, 0.875, f"+10 points of attendance ~ {slope * 10:+.1f} points of average score", transform=ax.transAxes,
        fontsize=10, color=INK2, va="top")
ax.set_xlabel("Attendance rate over 20 school days (%)")
ax.set_ylabel("Average score across subjects (%)")
ax.set_title("Students who attend more score higher", fontsize=12.5, fontweight="bold", loc="left", pad=12)
clean_axis(ax, "both")
fc = flags.category.value_counts()
ax = fig.add_axes((0.71, 0.12, 0.2, 0.63))
ax.barh(range(len(fc)), fc.values, height=0.6, color=[PCOL[flags[flags.category == c].priority.iloc[0]] for c in fc.index])
for i, v in enumerate(fc.values):
    ax.text(v + 0.5, i, v, va="center", fontsize=10.5, fontweight="bold")
ax.set_yticks(range(len(fc)), [c.replace("_", " ") for c in fc.index], fontsize=10.5)
ax.invert_yaxis()
ax.set_xticks([])
ax.tick_params(length=0)
for sp in ax.spines.values():
    sp.set_visible(False)
ax.set_title("Log entries flagged for follow-up", fontsize=12.5, fontweight="bold", loc="left", pad=12, x=-0.62)
for i, (p, lab) in enumerate([("high", "High priority"), ("medium", "Medium"), ("low", "Low")]):
    fig.text(0.6 + i * 0.125, 0.065, "■", color=PCOL[p], fontsize=12, va="center")
    fig.text(0.615 + i * 0.125, 0.065, lab, fontsize=10, color=INK2, va="center")
save(fig, "09_attendance_and_log_flags.png")

# =============================================================================
# 10 - conclusions
# =============================================================================
n_open = int((ov.student_id.isin(review.record)).sum())
det = ev["detection_total"]
worst = max(SUBJ, key=lambda s: ev["grades_naive"][s]["wrong"])
finds = [
    (f"{naive['wrong'] + naive['withheld']}", "of 960 scores",
     "would be wrong or lost without cleaning",
     f"A naive parse gets {naive['wrong']} scores wrong and drops {naive['withheld']}. {worst} is worst\n"
     f"({ev['grades_naive'][worst]['wrong']} wrong): three scales share one column."),
    (f"{N_WRONG}", f"wrong in {N_CHECK:,}",
     "values after cleaning, checked against the key",
     f"{N_OK:,} values match the truth exactly; {N_HELD} were withheld on purpose\nbecause the source gave no safe answer."),
    (f"{det['handled']}/{det['injected']}", "planted errors",
     "were caught and logged by the pipeline",
     "The one exception is a misspelt name with no ID: too far from any\nregister name to link safely, so it went to review instead."),
    (f"{len(review)}", "review items",
     "each with a priority and a method",
     f"{pc.get('high', 0)} high priority (security swipes, shared IDs, missed exams).\n{n_open} of 240 students have at least one open item."),
    (f"r = {r:.2f}", "",
     "attendance and achievement move together",
     f"Only visible once the log, three mark sources and the register\nshare one student ID and one scale."),
    ("4 -> 1", "scales",
     "0-1, /20, GPA 4.0 and percent became 0-100",
     "The scale is read from each table's own header and checked\nagainst value ranges - never assumed for a whole file."),
]
fig = canvas("Conclusions", "What the project shows about messy school data - and about the school.")
ax = blank_ax(fig)
for i, (big, unit, head, body) in enumerate(finds):
    x, y = (i % 2) * 51, 68 - (i // 2) * 33.5
    card(ax, x, y, 49, 30)
    ax.text(x + 2.5, y + 22.5, big, fontsize=27, fontweight="bold", va="center", color=BLUE)
    ax.text(x + 2.5, y + 14, f"{unit} {head}".strip(), fontsize=11.8, fontweight="bold", va="center")
    ax.text(x + 2.5, y + 6.3, body, fontsize=9.6, color=INK2, va="center", linespacing=1.4)
save(fig, "10_conclusions.png")

print(json.dumps(dict(checked=N_CHECK, ok=N_OK, held=N_HELD, wrong=N_WRONG, naive=naive, clean=clean, r=round(r, 3),
                      slope10=round(slope * 10, 2), open_students=n_open, review=pc.to_dict(), log=len(clog),
                      grade_means=hm.mean(axis=1).round(1).to_dict()), indent=1))
