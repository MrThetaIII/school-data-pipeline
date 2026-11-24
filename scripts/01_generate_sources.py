"""
Phase 1 - Synthesize messy school data in six formats.

A single "ground truth" school (Greenfield Academy, 240 students, grades 9-12)
is generated first. Each source file is then written from that truth with the
kind of errors that format typically carries. Every injected anomaly is written
to ground_truth/injected_anomalies.csv - the cleaning pipeline never reads that
folder; it is only used afterwards to score the cleaning.

Outputs (raw_data/):
  01_student_roster.xlsx            Excel student lists, one sheet per grade
  02_enrollment_form_responses.csv  Online-form export (free-text answers)
  03_gradebook_term1.xlsx           Teacher gradebook: Math + English, /100
  04_science_department_results.docx  Word tables: Science, 0-1 scale (/20 for Grade 11)
  05_history_gpa_report.pdf         PDF table: History on a 4.0 GPA scale
  06_gate_attendance.log            Gate swipe log, 20 school days
"""
import csv
import datetime as dt
import random
from pathlib import Path

import numpy as np
import pandas as pd
from docx import Document
from docx.shared import Pt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "raw_data"
TRUTH = ROOT / "ground_truth"
RAW.mkdir(exist_ok=True)
TRUTH.mkdir(exist_ok=True)

SEED = 2026
rng = random.Random(SEED)
nrng = np.random.default_rng(SEED)

KEY = []  # the answer key: one row per injected anomaly


def note(source, sid, field, category, raw, truth, expected):
    KEY.append(dict(source=source, student_id=sid, field=field, category=category,
                    raw_value=str(raw), true_value="" if truth is None else str(truth),
                    expected_action=expected))


def pick(pool, n, cond=None):
    cand = [x for x in pool if cond is None or cond(x)]
    chosen = rng.sample(cand, n)
    for c in chosen:
        pool.remove(c)
    return chosen


def clip(x, lo, hi):
    return max(lo, min(hi, x))


# ----------------------------------------------------------------------------
# 1. Ground truth
# ----------------------------------------------------------------------------
FIRST_F = ["Amira", "Layla", "Sofia", "Emma", "Hana", "Nour", "Maya", "Chloe", "Zara", "Isabella",
           "Yasmin", "Olivia", "Fatima", "Grace", "Leila", "Mei", "Priya", "Elena", "Salma", "Ruby",
           "Aisha", "Clara", "Dina", "Freya", "Ingrid", "Jana", "Keira", "Lucia", "Mariam", "Nadia",
           "Rania", "Sara", "Tala", "Vera", "Wendy", "Ximena", "Yara", "Zeina", "Anika", "Bianca"]
FIRST_M = ["Omar", "Liam", "Youssef", "Noah", "Karim", "Ethan", "Hassan", "Lucas", "Adam", "Mateo",
           "Tariq", "Oliver", "Khaled", "James", "Ziad", "Hiro", "Arjun", "Marco", "Samir", "Felix",
           "Bilal", "Caleb", "Daniel", "Elias", "Faris", "Gabriel", "Hamza", "Ivan", "Jonas", "Kenji",
           "Leo", "Malik", "Nabil", "Oscar", "Pablo", "Rami", "Stefan", "Tomas", "Victor", "Walid"]
LAST = ["Hassan", "Smith", "Mansour", "Johnson", "Khalil", "Garcia", "Ibrahim", "Brown", "Nasser",
        "Martinez", "Farouk", "Wilson", "Saleh", "Anderson", "Haddad", "Taylor", "Aziz", "Thomas",
        "Rahman", "Moore", "Soliman", "Jackson", "Qasim", "White", "Darwish", "Harris", "Fahmy",
        "Clark", "Yilmaz", "Lewis", "Tanaka", "Walker", "Patel", "Young", "Rossi", "King", "Novak",
        "Wright", "Petrov", "Scott", "Silva", "Green", "Kowalski", "Baker", "Okafor", "Adams",
        "Lindqvist", "Nelson", "Al-Sayed", "Carter", "O'Brien", "Mitchell", "Dubois", "Perez",
        "Schneider", "Roberts", "Costa", "Turner", "Bergstrom", "Phillips"]

students = []
used = set()
n = 0
for grade in (9, 10, 11, 12):
    for k in range(60):
        n += 1
        gender = rng.choice("MF")
        while True:
            fn = rng.choice(FIRST_F if gender == "F" else FIRST_M)
            ln = rng.choice(LAST)
            if (fn, ln) not in used:
                used.add((fn, ln))
                break
        dob = dt.date(2026 - grade - 6, 9, 1) + dt.timedelta(days=rng.randrange(365))
        ability = clip(float(nrng.normal(74, 9)), 45, 96)

        def score():
            return int(clip(round(ability + float(nrng.normal(0, 7))), 30, 100))

        hist_gpa = round(clip((score() - 50) / 12.5, 0, 4), 1)
        pfirst = rng.choice(FIRST_F + FIRST_M)
        area = rng.choice([212, 305, 415, 617, 702, 808])
        last2 = rng.randrange(100)
        students.append(dict(
            student_id=f"STU-{n:04d}", first_name=fn, last_name=ln, gender=gender, dob=dob,
            grade=grade, section="ABC"[k // 20], ability=ability,
            math=score(), english=score(), science=score(),
            history_gpa=hist_gpa, history_pct=50 + 12.5 * hist_gpa,
            guardian_name=f"{pfirst} {ln}",
            guardian_email=f"{pfirst}.{ln}".lower().replace("'", "") + f"{rng.randrange(10, 99)}@"
                           + rng.choice(["examplemail.com", "mailbox.example.org", "familynet.example.com"]),
            guardian_phone=f"+1-{area}-555-01{last2:02d}", _area=area, _last2=last2,
            prev_avg=float(int(clip(round(ability + float(nrng.normal(0, 5))), 40, 99))),
            siblings=rng.choice([0, 0, 1, 1, 1, 2, 2, 3, 4]),
            transport=rng.choice(["Bus", "Bus", "Car", "Car", "Walk", "Bicycle"]),
            medical=rng.choice([""] * 12 + ["Asthma, carries inhaler", "Peanut allergy", "Wears glasses",
                                            "Type 1 diabetes - nurse informed", "None"]),
        ))

S = {s["student_id"]: s for s in students}
IDS = [s["student_id"] for s in students]
# scores a perfect cleaner could report (None where no score exists at all)
TRUE_SCORE = {sid: dict(Math=S[sid]["math"], English=S[sid]["english"],
                        Science=S[sid]["science"], History=S[sid]["history_pct"]) for sid in IDS}


def full(s):
    return f"{s['first_name']} {s['last_name']}"


def last_first(s):
    return f"{s['last_name']}, {s['first_name']}"


def typo(word):
    i = rng.randrange(1, len(word) - 2)
    while word[i] == word[i + 1]:          # swapping a double letter would change nothing
        i = rng.randrange(1, len(word) - 2)
    return word[:i] + word[i + 1] + word[i] + word[i + 2:]


non_resp = set(rng.sample(IDS, 16))          # families that never filled the form
resp = [i for i in IDS if i not in non_resp]

# ----------------------------------------------------------------------------
# 2. Roster (Excel, one sheet per grade)
# ----------------------------------------------------------------------------
SRC = "01_student_roster.xlsx"
R = {sid: dict(id=sid, first=S[sid]["first_name"], last=S[sid]["last_name"],
               gender=S[sid]["gender"], dob=S[sid]["dob"], grade=S[sid]["grade"],
               section=S[sid]["section"]) for sid in IDS}
for sid in IDS:   # house style differs per sheet: Grade 11 spells gender out
    if S[sid]["grade"] == 11:
        R[sid]["gender"] = {"M": "Male", "F": "Female"}[S[sid]["gender"]]

pool = IDS[:]
in_resp = lambda x: x in resp
in_non = lambda x: x in non_resp
protected = set()   # students whose form answers must stay reliable

for sid in pick(pool, 6, in_resp) + pick(pool, 2, in_non):
    R[sid]["dob"] = None
    protected.add(sid)
    note(SRC, sid, "dob", "dob_missing", "", S[sid]["dob"], "fill from form / flag")
bad_years = [1911, 1912, 2026, 2020, 2001]
for sid, yr in zip(pick(pool, 3, in_resp) + pick(pool, 2, in_non), bad_years):
    d = S[sid]["dob"]
    R[sid]["dob"] = dt.date(yr, d.month, min(d.day, 28))
    protected.add(sid)
    note(SRC, sid, "dob", "dob_illogical", R[sid]["dob"], d, "flag / fill from form")
for sid in pick(pool, 20):
    d = S[sid]["dob"]
    R[sid]["dob"] = rng.choice([f"{d:%d/%m/%Y}", f"{d:%Y-%m-%d}", f"{d:%B %d, %Y}", f"{d:%d-%b-%y}",
                                f"{d.day}.{d.month}.{d.year}"])
    note(SRC, sid, "dob", "dob_text_format", R[sid]["dob"], d, "standardise")
for sid in pick(pool, 14):
    num = int(sid[-4:])
    R[sid]["id"] = rng.choice([sid.lower(), f"STU{num:04d}", num, f"  {sid} ", f"STU_{num:04d}"])
    note(SRC, sid, "student_id", "id_format", R[sid]["id"], sid, "standardise")
for sid in pick(pool, 16):
    f = rng.choice([lambda x: x.upper(), lambda x: x.lower(), lambda x: f"  {x} ", lambda x: x + "  "])
    which = rng.choice(["first", "last"])
    R[sid][which] = f(R[sid][which])
    note(SRC, sid, which + "_name", "name_case_space", R[sid][which], S[sid][which + "_name"], "standardise")
for sid in pick(pool, 18):
    g = S[sid]["gender"]
    R[sid]["gender"] = rng.choice({"M": ["m", "male", "Boy", "M "], "F": ["f", "female", "Girl", " F"]}[g])
    note(SRC, sid, "gender", "gender_variant", R[sid]["gender"], g, "standardise")
for sid in pick(pool, 4):
    R[sid]["gender"] = None
    note(SRC, sid, "gender", "gender_missing", "", S[sid]["gender"], "omit + flag")
for sid in pick(pool, 3):
    R[sid]["grade"] = S[sid]["grade"] + rng.choice([-1, 1])
    note(SRC, sid, "grade", "grade_mismatch", R[sid]["grade"], S[sid]["grade"], "correct to sheet")
for sid in pick(pool, 6):
    R[sid]["section"] = rng.choice([S[sid]["section"].lower(), "Sec " + S[sid]["section"]])
    note(SRC, sid, "section", "section_format", R[sid]["section"], S[sid]["section"], "standardise")
exact_dups = set(pick(pool, 4))
for sid in exact_dups:
    note(SRC, sid, "row", "duplicate_row", "row repeated", "", "omit duplicate")
# two students typed with their neighbour's ID
dup_x = pick(pool, 2, lambda x: x in resp and int(x[-4:]) % 60 not in (0, 1))
dup_y = []
for sid in dup_x:
    y = f"STU-{int(sid[-4:]) - 1:04d}"
    dup_y.append(y)
    R[sid]["id"] = y
    protected.update([sid, y])
    note(SRC, sid, "student_id", "duplicate_id_conflict", y, sid, "resolve by cross-reference")

wb = Workbook()
wb.remove(wb.active)
bold = Font(bold=True)
fill = PatternFill("solid", fgColor="D9E2F3")
for grade in (9, 10, 11, 12):
    ws = wb.create_sheet(f"Grade {grade}")
    ws["A1"] = "Greenfield Academy - Student Register 2026/27"
    ws["A1"].font = Font(bold=True, size=14)
    ws.merge_cells("A1:G1")
    ws["A2"] = f"Grade {grade}   |   printed 28/09/2026 by Registrar's office"
    if grade == 11:      # a different clerk: different headers and column order
        cols = [("Student No.", "id"), ("Last Name", "last"), ("First Name", "first"), ("Sex", "gender"),
                ("DOB", "dob"), ("Grade", "grade"), ("Sec.", "section")]
    else:
        cols = [("Student ID", "id"), ("First Name", "first"), ("Last Name", "last"), ("Gender", "gender"),
                ("Date of Birth", "dob"), ("Grade", "grade"), ("Section", "section")]
    hdr = 5 if grade == 12 else 4
    if grade == 12:
        ws["A3"] = "NOTE: leavers list to be confirmed in January"
        cols.append(("Notes", None))
    for c, (title, _) in enumerate(cols, 1):
        cell = ws.cell(hdr, c, title)
        cell.font, cell.fill = bold, fill
    r = hdr
    members = [sid for sid in IDS if S[sid]["grade"] == grade]
    for i, sid in enumerate(members):
        for _ in range(2 if sid in exact_dups else 1):
            r += 1
            for c, (_, k) in enumerate(cols, 1):
                if k is None:
                    continue
                cell = ws.cell(r, c, R[sid][k])
                if k == "dob" and isinstance(R[sid][k], dt.date):
                    cell.number_format = "DD/MM/YYYY"
        if i == 29:
            r += 1           # a stray blank row in the middle of the list
    ws.cell(r + 2, 1, f"Total students: {len(members)}").font = bold
    for col, w in zip("ABCDEFGH", (14, 16, 16, 10, 18, 8, 10, 20)):
        ws.column_dimensions[col].width = w
wb.save(RAW / SRC)

# ----------------------------------------------------------------------------
# 3. Enrollment form (CSV export)
# ----------------------------------------------------------------------------
SRC = "02_enrollment_form_responses.csv"
WORDS = {9: "Ninth", 10: "Tenth", 11: "Eleventh", 12: "Twelfth"}
NUMW = {0: "none", 1: "one", 2: "two", 3: "three", 4: "four"}


def ts_str(t):
    return f"{t.month}/{t.day}/{t.year} {t:%H:%M:%S}"


def phone_raw(s):
    a, x = s["_area"], f"{s['_last2']:02d}"
    return rng.choice([f"({a}) 555-01{x}", f"{a}-555-01{x}", f"{a}.555.01{x}", f"+1 {a} 555 01{x}", f"{a}55501{x}"])


F = {}
for sid in resp:
    s = S[sid]
    g = s["grade"]
    d = s["dob"]
    F[sid] = dict(ts=dt.datetime(2026, 8, 20) + dt.timedelta(minutes=rng.randrange(0, 30000)),
                  id=sid, name=full(s), dob=f"{d.month}/{d.day}/{d.year}",
                  grade=rng.choice([str(g), f"Grade {g}", f"{g}th", f"G{g}", f"grade {g}"]),
                  gname=s["guardian_name"], email=s["guardian_email"], phone=phone_raw(s),
                  prev=str(int(s["prev_avg"])), sib=str(s["siblings"]),
                  transport=rng.choice([s["transport"]] * 4 + [s["transport"].lower(), s["transport"].upper() + " "]),
                  medical=s["medical"])

free = lambda x: x not in protected
fp = resp[:]
for sid in pick(fp, 8):
    F[sid]["grade"] = WORDS[S[sid]["grade"]]
for sid in pick(fp, 15):
    F[sid]["name"] = last_first(S[sid])
wrong_id = pick(fp, 3, free)
for sid in wrong_id:
    other = rng.choice([x for x in IDS if x != sid and x not in protected])
    F[sid]["id"] = other
    note(SRC, sid, "student_id", "wrong_id", other, sid, "correct by name match")
for sid in pick(fp, 3, free):
    F[sid]["name"] = f"{S[sid]['first_name']} {typo(S[sid]['last_name'])}"
    note(SRC, sid, "name", "name_mismatch", F[sid]["name"], full(S[sid]), "flag")
for sid in pick(fp, 3, free):
    d = S[sid]["dob"] + dt.timedelta(days=rng.choice([-31, 30, 365]))
    F[sid]["dob"] = f"{d.month}/{d.day}/{d.year}"
    note(SRC, sid, "dob", "dob_conflict", F[sid]["dob"], S[sid]["dob"], "flag")

fp = resp[:]   # previous-average free-text scales
for sid in pick(fp, 15):
    F[sid]["prev"] += "%"
    note(SRC, sid, "prev_avg", "scale_percent_sign", F[sid]["prev"], S[sid]["prev_avg"], "standardise")
for sid in pick(fp, 12):
    F[sid]["prev"] = f"{S[sid]['prev_avg'] / 100:.2f}"
    note(SRC, sid, "prev_avg", "scale_fraction", F[sid]["prev"], S[sid]["prev_avg"], "correct x100")
for i, sid in enumerate(pick(fp, 17, lambda x: S[x]["prev_avg"] >= 65)):
    g = round((S[sid]["prev_avg"] - 50) / 12.5, 1)
    S[sid]["prev_avg"] = 50 + 12.5 * g
    F[sid]["prev"] = f"{g}" if i < 12 else f"{g}/4"
    note(SRC, sid, "prev_avg", "scale_gpa", F[sid]["prev"], S[sid]["prev_avg"], "convert GPA to %")
for sid in pick(fp, 5):
    F[sid]["prev"] = rng.choice(["A-", "B+", "B", "A", "C+"])
    note(SRC, sid, "prev_avg", "letter_grade", F[sid]["prev"], S[sid]["prev_avg"], "flag")
for sid in pick(fp, 2):
    F[sid]["prev"] = str(int(S[sid]["prev_avg"])) + rng.choice(["0", "5"])
    note(SRC, sid, "prev_avg", "illogical_value", F[sid]["prev"], S[sid]["prev_avg"], "flag")
for sid in pick(fp, 6):
    F[sid]["prev"] = ""
    note(SRC, sid, "prev_avg", "blank", "", S[sid]["prev_avg"], "omit")

fp = resp[:]   # siblings
for sid in pick(fp, 8):
    F[sid]["sib"] = NUMW[S[sid]["siblings"]]
    note(SRC, sid, "siblings", "number_as_word", F[sid]["sib"], S[sid]["siblings"], "correct")
for sid in pick(fp, 3, lambda x: S[x]["siblings"] > 0):
    F[sid]["sib"] = f"{S[sid]['siblings']} " + rng.choice(["brothers", "sisters", "(older)"])
    note(SRC, sid, "siblings", "number_with_text", F[sid]["sib"], S[sid]["siblings"], "correct")
for sid, v in zip(pick(fp, 2), ["-1", "25"]):
    F[sid]["sib"] = v
    note(SRC, sid, "siblings", "illogical_value", v, S[sid]["siblings"], "flag")
for sid in pick(fp, 4):
    F[sid]["sib"] = ""
    note(SRC, sid, "siblings", "blank", "", S[sid]["siblings"], "omit")

fp = resp[:]   # contact details
for sid in pick(fp, 10):
    F[sid]["email"] = rng.choice([F[sid]["email"].upper(), f" {F[sid]['email']} ", F[sid]["email"].title()])
    note(SRC, sid, "email", "email_case_space", F[sid]["email"], S[sid]["guardian_email"], "standardise")
for sid in pick(fp, 3):
    head, _, tail = F[sid]["email"].rpartition(".")
    F[sid]["email"] = head + "," + tail
    note(SRC, sid, "email", "email_typo", F[sid]["email"], S[sid]["guardian_email"], "correct")
for sid in pick(fp, 3):
    F[sid]["email"] = F[sid]["email"].replace("@", "")
    note(SRC, sid, "email", "email_invalid", F[sid]["email"], S[sid]["guardian_email"], "flag")
for sid in pick(fp, 3):
    F[sid]["email"] = ""
    note(SRC, sid, "email", "blank", "", S[sid]["guardian_email"], "omit")
for sid in pick(fp, 3):
    F[sid]["phone"] = f"555-01{S[sid]['_last2']:02d}"[: rng.choice([6, 8])]
    note(SRC, sid, "phone", "phone_invalid", F[sid]["phone"], S[sid]["guardian_phone"], "flag")
for sid in pick(fp, 2):
    F[sid]["phone"] = ""
    note(SRC, sid, "phone", "blank", "", S[sid]["guardian_phone"], "omit")

rows = []
for sid in resp:
    rows.append(F[sid])
for sid in rng.sample(resp, 12):       # an earlier submission later replaced by a corrected one
    old = dict(F[sid])
    old["ts"] = F[sid]["ts"] - dt.timedelta(hours=rng.randrange(2, 90))
    old["phone"] = f"({rng.choice([212, 305, 415])}) 555-01{rng.randrange(100):02d}"
    old["transport"] = rng.choice(["Bus", "Car", "Walk"])
    rows.append(old)
    note(SRC, sid, "row", "duplicate_submission", "earlier response", "", "keep latest")
for nm, sid in [("test test", "STU-0000"), ("asdf qwerty", "123"), ("TEST - please ignore", "")]:
    rows.append(dict(ts=dt.datetime(2026, 8, 19, 9, rng.randrange(60)), id=sid, name=nm, dob="1/1/2000",
                     grade="9", gname="test", email="test@test", phone="000", prev="100", sib="0",
                     transport="Bus", medical=""))
    note(SRC, "", "row", "test_submission", nm, "", "omit")
for nm, sid, g in [("Dalia Mostafa", "STU-0412", 9), ("Peter Lang", "STU-0388", 11)]:
    rows.append(dict(ts=dt.datetime(2026, 9, 9, 11, rng.randrange(60)), id=sid, name=nm,
                     dob=f"3/14/{2026 - g - 5}", grade=str(g), gname="Parent " + nm.split()[1],
                     email=nm.lower().replace(" ", ".") + "@examplemail.com", phone="(212) 555-0147",
                     prev="81", sib="1", transport="Car", medical=""))
    note(SRC, sid, "student_id", "unknown_id", sid, "", "flag")
rows.sort(key=lambda r: r["ts"])
with open(RAW / SRC, "w", newline="", encoding="utf-8-sig") as fh:
    w = csv.writer(fh)
    w.writerow(["Timestamp", "Student ID", "Student Full Name", "Date of Birth (MM/DD/YYYY)", "Grade",
                "Parent/Guardian Name", "Parent Email", "Parent Phone", "Previous School Average",
                "Number of Siblings", "Transportation", "Medical Notes"])
    for r in rows:
        w.writerow([ts_str(r["ts"]), r["id"], r["name"], r["dob"], r["grade"], r["gname"], r["email"],
                    r["phone"], r["prev"], r["sib"], r["transport"], r["medical"]])

# ----------------------------------------------------------------------------
# 4. Gradebook (Excel): Math + English, coursework/40 + exam/60 = total/100
# ----------------------------------------------------------------------------
SRC = "03_gradebook_term1.xlsx"
wb = Workbook()
wb.remove(wb.active)
for subject, key, teacher in [("Math", "math", "Mr. R. Delgado"), ("English", "english", "Ms. H. Whitcombe")]:
    G = {}
    for sid in IDS:
        t = S[sid][key]
        cw = int(clip(round(t * 0.4 + rng.choice([-2, -1, 0, 0, 1, 2])), max(0, t - 60), min(40, t)))
        G[sid] = dict(id=sid, name=full(S[sid]), cls=f"{S[sid]['grade']}{S[sid]['section']}",
                      cw=cw, exam=t - cw, total=t, remarks="")
    gp = IDS[:]
    frac_ids = []
    if subject == "English":   # one class teacher typed totals as fractions of 1
        frac_ids = pick(gp, 20, lambda x: S[x]["grade"] == 10 and S[x]["section"] == "C")
        for sid in frac_ids:
            G[sid]["total"] = round(S[sid][key] / 100, 2)
            note(SRC, sid, subject, "scale_fraction", G[sid]["total"], S[sid][key], "correct x100")
    for sid in pick(gp, 4):
        G[sid].update(cw="", exam="", total=rng.choice(["ABS", "absent", "Absent"]), remarks="makeup?")
        TRUE_SCORE[sid][subject] = None
        note(SRC, sid, subject, "text_absent", G[sid]["total"], None, "omit + flag")
    for sid in pick(gp, 3):
        G[sid].update(cw="", exam="", total=rng.choice(["N/A", "-", "pending"]))
        TRUE_SCORE[sid][subject] = None
        note(SRC, sid, subject, "text_not_available", G[sid]["total"], None, "omit")
    for sid in pick(gp, 8):
        t = S[sid][key]
        G[sid]["total"] = rng.choice([f"{t}%", f" {t} ", f"{t}.0", f"{t},0"])
        note(SRC, sid, subject, "number_as_text", G[sid]["total"], t, "standardise")
    for sid in pick(gp, 4):
        t = S[sid][key]
        G[sid]["total"] = rng.choice([t + 100, -t, 100 + rng.randrange(3, 9)])
        note(SRC, sid, subject, "illogical_total", G[sid]["total"], t, "correct from components")
    for sid in pick(gp, 3, lambda x: str(S[x][key])[0] != str(S[x][key])[-1] and S[x][key] < 100):
        t = S[sid][key]
        G[sid]["total"] = int(str(t)[::-1])
        note(SRC, sid, subject, "total_mismatch", G[sid]["total"], t, "correct from components + flag")
    for sid in pick(gp, 2):
        G[sid]["exam"] = G[sid]["exam"] + rng.choice([30, 40])
        if G[sid]["exam"] <= 60:
            G[sid]["exam"] = 75
        note(SRC, sid, subject, "component_out_of_range", G[sid]["exam"], S[sid][key], "keep total + flag")
    for sid in pick(gp, 3):
        G[sid]["total"] = ""
        note(SRC, sid, subject, "total_blank", "", S[sid][key], "compute from components")
    for sid in pick(gp, 4):
        G[sid]["id"] = ""
        note(SRC, sid, subject, "id_missing", "", sid, "match by name")
    missing = pick(gp, 3)
    for sid in missing:
        TRUE_SCORE[sid][subject] = None
        note(SRC, sid, subject, "row_missing", "", None, "omit + flag")
    conflict = pick(gp, 1)[0]
    note(SRC, conflict, subject, "duplicate_conflict", "two rows, different marks", S[conflict][key], "flag")

    ws = wb.create_sheet(subject)
    ws["A1"] = f"{'Mathematics' if subject == 'Math' else 'English Language'} - Term 1 2026/27 - {teacher}"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = "Coursework is marked out of 40, exam out of 60. Total = Coursework + Exam."
    for c, h in enumerate(["ID", "Student Name", "Class", "Coursework (40)", "Exam (60)", "Total (100)", "Remarks"], 1):
        cell = ws.cell(3, c, h)
        cell.font, cell.fill = bold, fill
    r = 3
    for sid in IDS:
        if sid in missing:
            continue
        reps = [G[sid]]
        if sid == conflict:
            alt = dict(G[sid])
            alt["exam"] = G[sid]["exam"] - 9 if G[sid]["exam"] >= 9 else G[sid]["exam"] + 9
            alt["total"] = alt["cw"] + alt["exam"]
            reps.append(alt)
        for g in reps:
            r += 1
            for c, k in enumerate(["id", "name", "cls", "cw", "exam", "total", "remarks"], 1):
                ws.cell(r, c, g[k] if g[k] != "" else None)
    ws.cell(r + 2, 2, "Class Average").font = bold
    ws.cell(r + 2, 6, round(float(np.mean([S[x][key] for x in IDS])), 1))
    for col, w in zip("ABCDEFG", (12, 24, 8, 16, 12, 12, 14)):
        ws.column_dimensions[col].width = w
wb.save(RAW / SRC)

# ----------------------------------------------------------------------------
# 5. Science results (Word tables): 0-1 scale, Grade 11 marked out of 20
# ----------------------------------------------------------------------------
SRC = "04_science_department_results.docx"
D = {}
for sid in IDS:
    t = S[sid]["science"]
    D[sid] = dict(id=sid, name=last_first(S[sid]),
                  score=f"{t / 5:.1f}" if S[sid]["grade"] == 11 else f"{t / 100:.2f}", remarks="")
dp = IDS[:]
not11 = lambda x: S[x]["grade"] != 11
is11 = lambda x: S[x]["grade"] == 11
for sid in pick(dp, 11, not11):
    D[sid]["score"] = rng.choice([str(S[sid]["science"]), f"{S[sid]['science']}%"])
    note(SRC, sid, "Science", "scale_percent_in_fraction_column", D[sid]["score"], S[sid]["science"], "correct /100")
for sid in pick(dp, 3, is11):
    D[sid]["score"] = f"{S[sid]['science']}%"
    note(SRC, sid, "Science", "scale_percent_in_fraction_column", D[sid]["score"], S[sid]["science"], "correct")
for sid in pick(dp, 6):
    D[sid]["score"] = D[sid]["score"].replace(".", ",")
    note(SRC, sid, "Science", "comma_decimal", D[sid]["score"], S[sid]["science"], "standardise")
for sid in pick(dp, 4):
    D[sid]["score"] = rng.choice(["B+", "A-", "B", "C+"])
    note(SRC, sid, "Science", "letter_grade", D[sid]["score"], S[sid]["science"], "flag")
for sid in pick(dp, 5):
    D[sid]["score"] = rng.choice(["exempt", "-", "", "see me"])
    D[sid]["remarks"] = rng.choice(["", "medical", "transfer"])
    TRUE_SCORE[sid]["Science"] = None
    note(SRC, sid, "Science", "text_not_available", D[sid]["score"], None, "omit")
for sid in pick(dp, 3, not11):
    D[sid]["score"] = f"{S[sid]['science'] / 10:.1f}"
    note(SRC, sid, "Science", "misplaced_decimal", D[sid]["score"], S[sid]["science"], "flag")
for sid in pick(dp, 2, is11):
    D[sid]["score"] = f"{20 + rng.randrange(2, 9)}"
    note(SRC, sid, "Science", "out_of_range", D[sid]["score"], S[sid]["science"], "flag")
id_missing = pick(dp, 8)
for i, sid in enumerate(id_missing):
    D[sid]["id"] = ""
    if i < 3:
        D[sid]["name"] = f"{typo(S[sid]['last_name'])}, {S[sid]['first_name']}"
        note(SRC, sid, "student_id", "id_missing_name_typo", D[sid]["name"], sid, "fuzzy match + flag")
    else:
        note(SRC, sid, "student_id", "id_missing", "", sid, "match by name")
for sid in pick(dp, 5):
    D[sid]["id"] = sid.replace("0", "O", 1)
    note(SRC, sid, "student_id", "id_letter_O", D[sid]["id"], sid, "standardise")
d_missing = pick(dp, 3)
for sid in d_missing:
    TRUE_SCORE[sid]["Science"] = None
    note(SRC, sid, "Science", "row_missing", "", None, "omit + flag")

doc = Document()
doc.styles["Normal"].font.size = Pt(10)
doc.add_heading("Science Department - Term 1 Results 2026/27", 0)
doc.add_paragraph("Compiled by the Head of Science from class teachers' submissions. Final scores are recorded "
                  "as a proportion of 1 (e.g. 0.85) unless the table header states otherwise. "
                  "Grade 11 follows the practical-assessment scheme and is marked out of 20.")
for grade in (9, 10, 11, 12):
    doc.add_heading(f"Grade {grade}", 1)
    members = [x for x in IDS if S[x]["grade"] == grade and x not in d_missing]
    tb = doc.add_table(rows=1, cols=5)
    tb.style = "Table Grid"
    for c, h in enumerate(["No.", "Student ID", "Student Name",
                           "Final Score (/20)" if grade == 11 else "Final Score (0-1)", "Remarks"]):
        tb.rows[0].cells[c].text = h
        tb.rows[0].cells[c].paragraphs[0].runs[0].font.bold = True
    cur, k = None, 0
    for sid in members:
        if S[sid]["section"] != cur:
            cur = S[sid]["section"]
            cells = tb.add_row().cells
            m = cells[0].merge(cells[4])
            m.text = f"Section {cur}"
        k += 1
        cells = tb.add_row().cells
        for c, v in enumerate([str(k), D[sid]["id"], D[sid]["name"], D[sid]["score"], D[sid]["remarks"]]):
            cells[c].text = v
    tb.add_row()   # trailing empty row
    doc.add_paragraph(f"Signed: Class coordinators, Grade {grade}")
doc.save(RAW / SRC)

# ----------------------------------------------------------------------------
# 6. History GPA report (PDF)
# ----------------------------------------------------------------------------
SRC = "05_history_gpa_report.pdf"
P = {sid: dict(id=sid, name=last_first(S[sid]), grade=str(S[sid]["grade"]), sec=S[sid]["section"],
               gpa=f"{S[sid]['history_gpa']:.1f}", remarks="") for sid in IDS}
pp = IDS[:]
for sid, v in zip(pick(pp, 3), ["4.5", "4.8", "5.0"]):
    P[sid]["gpa"] = v
    note(SRC, sid, "History", "gpa_over_4", v, S[sid]["history_pct"], "flag")
for sid in pick(pp, 6, lambda x: S[x]["history_gpa"] >= 0.8):
    P[sid]["gpa"] = str(round(S[sid]["history_pct"]))
    note(SRC, sid, "History", "scale_percent_in_gpa_column", P[sid]["gpa"], S[sid]["history_pct"], "convert + flag")
for sid in pick(pp, 5):
    P[sid]["gpa"] = P[sid]["gpa"].replace(".", ",")
    note(SRC, sid, "History", "comma_decimal", P[sid]["gpa"], S[sid]["history_pct"], "standardise")
for sid in pick(pp, 4):
    P[sid]["gpa"] += "/4"
    note(SRC, sid, "History", "number_as_text", P[sid]["gpa"], S[sid]["history_pct"], "standardise")
for sid in pick(pp, 5):
    P[sid]["gpa"] = rng.choice(["I", "W", "INC"])
    P[sid]["remarks"] = "see registrar"
    TRUE_SCORE[sid]["History"] = None
    note(SRC, sid, "History", "text_not_available", P[sid]["gpa"], None, "omit")
for sid in pick(pp, 1):
    P[sid]["gpa"] = "-" + P[sid]["gpa"]
    note(SRC, sid, "History", "illogical_value", P[sid]["gpa"], S[sid]["history_pct"], "flag")
for sid in pick(pp, 5):
    P[sid]["id"] = rng.choice([sid.replace("-", "- "), sid.replace("-", " ")])
    note(SRC, sid, "student_id", "id_format", P[sid]["id"], sid, "standardise")
p_dups = set(pick(pp, 3))
for sid in p_dups:
    note(SRC, sid, "row", "duplicate_row", "row repeated", "", "omit duplicate")
p_missing = pick(pp, 2)
for sid in p_missing:
    TRUE_SCORE[sid]["History"] = None
    note(SRC, sid, "History", "row_missing", "", None, "omit + flag")

styles = getSampleStyleSheet()
cellst = styles["BodyText"]
cellst.fontSize, cellst.leading = 8.5, 10
data = [["Student ID", "Student Name", "Grade", "Sec", "History GPA", "Remarks"]]
for sid in IDS:
    if sid in p_missing:
        continue
    row = [P[sid]["id"], Paragraph(P[sid]["name"], cellst), P[sid]["grade"], P[sid]["sec"], P[sid]["gpa"],
           P[sid]["remarks"]]
    data.append(row)
    if sid in p_dups:
        data.append(list(row))
table = Table(data, colWidths=[2.6 * cm, 3.6 * cm, 1.5 * cm, 1.3 * cm, 2.6 * cm, 3.4 * cm], repeatRows=1)
table.setStyle(TableStyle([
    ("GRID", (0, 0), (-1, -1), 0.5, colors.grey), ("FONTSIZE", (0, 0), (-1, -1), 8.5),
    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9E2F3")), ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))


def footer(canvas, d):
    canvas.setFont("Helvetica", 8)
    canvas.drawString(2 * cm, 1.2 * cm, "Greenfield Academy - Humanities Department - CONFIDENTIAL")
    canvas.drawRightString(A4[0] - 2 * cm, 1.2 * cm, f"Page {d.page}")


pdf = SimpleDocTemplate(str(RAW / SRC), pagesize=A4, topMargin=2 * cm, bottomMargin=2 * cm)
pdf.build([Paragraph("History - Term 1 GPA Report 2026/27", styles["Title"]),
           Paragraph("Grades are reported on the 4.0 scale. School conversion policy: "
                     "<b>Percent = 50 + 12.5 x GPA</b> (GPA 4.0 = 100%, GPA 0.0 = 50% or below). "
                     "I = incomplete, W = withdrawn.", styles["BodyText"]),
           Spacer(1, 0.4 * cm), table], onFirstPage=footer, onLaterPages=footer)

# ----------------------------------------------------------------------------
# 7. Gate attendance log
# ----------------------------------------------------------------------------
SRC = "06_gate_attendance.log"
days = [dt.date(2026, 10, 5) + dt.timedelta(days=i) for i in range(26)]
days = [d for d in days if d.weekday() < 5]            # 20 school days
ATT = {}                                                # (sid, day) -> present / late / absent
for sid in IDS:
    a = S[sid]["ability"]
    p_abs = clip(0.03 + 0.005 * (78 - a) + float(nrng.normal(0, 0.03)), 0.0, 0.45)
    p_late = clip(0.06 + 0.003 * (78 - a) + float(nrng.normal(0, 0.03)), 0.0, 0.4)
    for d in days:
        u = rng.random()
        ATT[sid, d] = "absent" if u < p_abs else ("late" if u < p_abs + p_late else "present")


def tm(d, h1, m1, h2, m2):
    a, b = h1 * 3600 + m1 * 60, h2 * 3600 + m2 * 60
    return dt.datetime.combine(d, dt.time()) + dt.timedelta(seconds=rng.randrange(a, b))


EV = {}       # (sid, day) -> dict(checkin=[ts, idtext, status], checkout=ts, absent=reason)
for (sid, d), st in ATT.items():
    if st == "absent":
        EV[sid, d] = dict(absent=rng.choice(["unexcused", "unexcused", "excused", "medical"]))
    else:
        t_in = tm(d, 8, 1, 9, 20) if st == "late" else tm(d, 7, 10, 7, 59)
        EV[sid, d] = dict(checkin=[t_in, sid, "LATE" if st == "late" else "OK"], checkout=tm(d, 14, 30, 15, 15))

present_days = [k for k, v in ATT.items() if v == "present"]
late_days = [k for k, v in ATT.items() if v == "late"]
absent_days = [k for k, v in ATT.items() if v == "absent"]
extra = []    # extra raw lines: (timestamp, text-after-timestamp)


def lg(cat, k, raw, expected):
    note(SRC, k[0], str(k[1]), cat, raw, ATT.get(k, ""), expected)


def gate():
    return f"GATE-{rng.choice([1, 1, 2, 2, 3])}"


for k in pick(present_days, 40):
    EV[k]["dup"] = rng.randrange(2, 9)
    lg("duplicate_swipe", k, "double CHECKIN", "omit duplicate")
for k in pick(present_days, 10):
    num = int(k[0][-4:])
    EV[k]["checkin"][1] = rng.choice([f"STU-{num}", k[0].lower(), f"STU{num:04d}"])
    lg("id_format", k, EV[k]["checkin"][1], "standardise")
for k in pick(present_days, 4):
    EV[k]["checkin"][1] = ""
    lg("id_empty", k, "id=", "omit line + flag")
for k in pick(present_days, 15):
    EV[k]["checkout"] = None
    lg("missing_checkout", k, "no CHECKOUT", "flag")
for k in pick(present_days, 10):
    EV[k]["checkin"] = None
    lg("missing_checkin", k, "no CHECKIN", "flag")
for k in pick(present_days, 12):
    EV[k]["checkout"] = tm(k[1], 9, 45, 11, 45)
    lg("early_leave", k, f"{EV[k]['checkout']:%H:%M}", "flag")
for k in pick(present_days, 8):
    EV[k]["absent"] = "unexcused"
    lg("absent_but_swiped", k, "ABSENT + CHECKIN", "flag, count present")
for k in pick(late_days, 10):
    EV[k]["checkin"][2] = "OK"
    lg("late_not_marked", k, f"{EV[k]['checkin'][0]:%H:%M} status=OK", "correct to late")
for k in pick(absent_days, 3):
    EV[k] = {}
    lg("no_record", k, "nothing logged", "infer absent + flag")
trunc = set(pick(present_days, 6))
for k in trunc:
    lg("truncated_line", k, "CHECKOUT line cut off", "omit line")
for i in range(8):
    uid, d = f"STU-{rng.choice([999, 731, 402, 555]):04d}", rng.choice(days)
    extra.append((tm(d, 7, 10, 8, 30), f"[WARN] {gate()} CHECKIN id={uid} card=UNREGISTERED status=OK"))
    lg("unknown_id", (uid, d), uid, "flag")
for i in range(6):
    sid, d = rng.choice(IDS), rng.choice(days)
    extra.append((tm(d, 1, 30, 3, 45), f"[INFO] {gate()} CHECKIN id={sid} card={rng.randrange(16 ** 6):06X} status=OK"))
    lg("odd_hour_swipe", (sid, d), "CHECKIN 01:30-03:45", "flag")
for i in range(5):
    sid, d = rng.choice(IDS), rng.choice([dt.date(2026, 10, 10), dt.date(2026, 10, 18), dt.date(2026, 10, 24)])
    extra.append((tm(d, 9, 0, 16, 0), f"[INFO] {gate()} CHECKIN id={sid} card={rng.randrange(16 ** 6):06X} status=OK"))
    lg("weekend_swipe", (sid, d), f"{d:%A}", "flag")
for i in range(2):
    sid = rng.choice(IDS)
    extra.append((tm(dt.date(2062, 10, 12 + i), 7, 10, 7, 50),
                  f"[INFO] {gate()} CHECKIN id={sid} card={rng.randrange(16 ** 6):06X} status=OK"))
    lg("impossible_date", (sid, dt.date(2062, 10, 12 + i)), "year 2062", "omit + flag")
for d in days:   # ordinary system chatter that is not attendance
    for i in range(7):
        extra.append((tm(d, 6, 0, 16, 0), rng.choice([
            "[DEBUG] SYSTEM heartbeat ok", f"[ERROR] {gate()} reader timeout, retrying",
            "[INFO] SYSTEM sync with SIS completed", f"[DEBUG] {gate()} door sensor cycle"])))

lines = list(extra)
ALT_DAY = dt.date(2026, 10, 14)   # GATE-3 firmware wrote DD/MM/YYYY timestamps that day
n_alt = 0
for (sid, d), e in EV.items():
    if e.get("checkin"):
        t, idt, st = e["checkin"]
        lvl = "WARN" if st == "LATE" else "INFO"
        g = gate()
        txt = f"[{lvl}] {g} CHECKIN id={idt} card={rng.randrange(16 ** 6):06X} status={st}"
        lines.append((t, txt))
        if e.get("dup"):
            lines.append((t + dt.timedelta(seconds=e["dup"]), txt))
    if e.get("checkout"):
        txt = f"[INFO] {gate()} CHECKOUT id={sid} status=OK"
        if (sid, d) in trunc:
            txt = "TRUNC" + txt
        lines.append((e["checkout"], txt))
    if e.get("absent"):
        lines.append((dt.datetime.combine(d, dt.time(9, 30)), f"[INFO] SYSTEM ABSENT id={sid} reason={e['absent']}"))
lines.sort(key=lambda x: x[0])
with open(RAW / SRC, "w", encoding="utf-8") as fh:
    fh.write("# Greenfield Academy gate controller v3.2 - attendance export 2026-10-05..2026-10-30\n")
    for t, txt in lines:
        stamp = f"{t:%Y-%m-%d %H:%M:%S}"
        if t.date() == ALT_DAY and "GATE-3" in txt:
            stamp = f"{t:%d/%m/%Y %H:%M:%S}"
            n_alt += 1
        line = f"{stamp} {txt}"
        if txt.startswith("TRUNC"):
            line = f"{stamp} {txt[5:]}"
            line = line[: rng.randrange(14, len(line) - 1)]
        fh.write(line + "\n")
note(SRC, "", str(ALT_DAY), "timestamp_format", f"{n_alt} lines in DD/MM/YYYY", "", "standardise")

# ----------------------------------------------------------------------------
# 8. Save the ground truth (never read by the cleaning pipeline)
# ----------------------------------------------------------------------------
truth = []
for sid in IDS:
    s = S[sid]
    st = [ATT[sid, d] for d in days]
    truth.append(dict(
        student_id=sid, first_name=s["first_name"], last_name=s["last_name"], gender=s["gender"],
        dob=s["dob"].isoformat(), grade=s["grade"], section=s["section"],
        Math=TRUE_SCORE[sid]["Math"], English=TRUE_SCORE[sid]["English"],
        Science=TRUE_SCORE[sid]["Science"], History=TRUE_SCORE[sid]["History"],
        days_present=sum(x != "absent" for x in st), days_late=st.count("late"), days_absent=st.count("absent"),
        form_submitted=sid in resp, guardian_email=s["guardian_email"], guardian_phone=s["guardian_phone"],
        prev_avg=s["prev_avg"], siblings=s["siblings"], transport=s["transport"]))
pd.DataFrame(truth).to_csv(TRUTH / "students_truth.csv", index=False)
key = pd.DataFrame(KEY)
key.to_csv(TRUTH / "injected_anomalies.csv", index=False)

print(f"students: {len(IDS)}   school days: {len(days)}   log lines: {len(lines)}")
print(f"injected anomalies: {len(key)}")
print(key.groupby("source").size().to_string())
