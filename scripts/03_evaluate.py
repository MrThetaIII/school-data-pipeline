"""
Score the cleaning against the ground truth written by 01_generate_sources.py.

This is the only script that reads ground_truth/. It answers three questions:
  1. How many values in the clean database equal the truth?
  2. How would a naive parse (declared scale, no cleaning) have done?
  3. Was each injected anomaly picked up by the cleaning log?
Output: reports/evaluation.json
"""
import json
import sqlite3
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
truth = pd.read_csv(ROOT / "ground_truth" / "students_truth.csv").set_index("student_id")
key = pd.read_csv(ROOT / "ground_truth" / "injected_anomalies.csv", keep_default_na=False)
con = sqlite3.connect(ROOT / "clean_data" / "school.db")
q = lambda sql: pd.read_sql(sql, con)
students = q("select * from students").set_index("student_id")
guard = q("select * from guardians").set_index("student_id")
grades = q("select * from grades")
att = q("select * from attendance_summary").set_index("student_id")
clog = q("select * from cleaning_log")
staging = pd.read_csv(ROOT / "clean_data" / "staging_grades.csv", keep_default_na=False)
SUBJ = ["Math", "English", "Science", "History"]
res = {}


def outcome(clean, true):
    """correct / withheld (blank where a true value exists) / wrong."""
    cn, tn = pd.isna(clean), pd.isna(true)
    if cn and tn:
        return "correct"
    if cn:
        return "withheld"
    if tn:
        return "wrong"
    same = abs(float(clean) - float(true)) < 0.06 if isinstance(true, (int, float, np.number)) else str(clean) == str(true)
    return "correct" if same else "wrong"


def tally(pairs):
    s = pd.Series([outcome(c, t) for c, t in pairs]).value_counts()
    return {k: int(s.get(k, 0)) for k in ("correct", "withheld", "wrong")}


# 1. clean database vs truth -------------------------------------------------
res["students_found"] = int(students.index.isin(truth.index).sum())
res["student_fields"] = {}
for col, tcol in [("first_name", "first_name"), ("last_name", "last_name"), ("gender", "gender"),
                  ("date_of_birth", "dob"), ("grade", "grade"), ("section", "section")]:
    res["student_fields"][col] = tally((students[col].get(i), truth.at[i, tcol]) for i in truth.index)

wide = grades.pivot(index="student_id", columns="subject", values="score_pct")
res["grades_clean"] = {s: tally((wide[s].get(i), truth.at[i, s]) for i in truth.index) for s in SUBJ}

res["attendance"] = {c: tally((att[c].get(i), truth.at[i, c]) for i in truth.index)
                     for c in ("days_present", "days_late", "days_absent")}

resp = truth[truth.form_submitted]
res["guardians"] = {}
for col, tcol in [("guardian_email", "guardian_email"), ("guardian_phone", "guardian_phone"),
                  ("prev_school_avg_pct", "prev_avg"), ("siblings", "siblings"), ("transport", "transport")]:
    res["guardians"][col] = tally((guard[col].get(i), resp.at[i, tcol]) for i in resp.index)

# 2. naive baseline for grades: trust the ID as typed, trust the declared scale -
conv = {"Math": lambda v: v, "English": lambda v: v, "Science": lambda v: v * 100, "History": lambda v: 50 + 12.5 * v}
naive = {}
for r in staging.itertuples():
    v = pd.to_numeric(r.raw_value.split(" vs ")[0], errors="coerce")
    if r.raw_id in truth.index and not pd.isna(v):
        naive[r.raw_id, r.subject] = conv[r.subject](v)
res["grades_naive"] = {s: tally((naive.get((i, s), np.nan), truth.at[i, s]) for i in truth.index) for s in SUBJ}

# 3. was each injected anomaly handled? --------------------------------------
ALIAS = {"scale_gpa": {"scale_gpa"}, "dob_missing": {"dob_missing"}, "unknown_id": {"unknown_id"}}
rows = []
for (src, cat), grp in key.groupby(["source", "category"]):
    logged = clog[(clog.source == src) & (clog.category == cat)]
    if src.endswith(".log") or (grp.student_id == "").all():
        hit = min(len(grp), len(logged)) if cat != "timestamp_format" else int(len(logged) > 0)
    else:
        hit = int(grp.student_id.isin(logged.record).sum())
    rows.append(dict(source=src, category=cat, injected=len(grp), handled=hit))
det = pd.DataFrame(rows)
res["detection"] = det.to_dict("records")
res["detection_total"] = dict(injected=int(det.injected.sum()), handled=int(det.handled.sum()))

(ROOT / "reports" / "evaluation.json").write_text(json.dumps(res, indent=2), encoding="utf-8")

print("students found:", res["students_found"])
for block in ("student_fields", "grades_clean", "grades_naive", "attendance", "guardians"):
    print(f"\n{block}")
    print(pd.DataFrame(res[block]).T.to_string())
print("\ndetection:", res["detection_total"])
print(det[det.injected != det.handled].to_string())
