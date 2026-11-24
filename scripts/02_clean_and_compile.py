"""
Phase 2 - Parse the six raw sources, clean them, and compile one database.

Every change is written to a cleaning log with one of five actions:
  standardised  format-only change, the meaning of the value is untouched
  corrected     the value was repaired or converted (scale, typo, cross-reference)
  omitted       a row or value was deliberately left out (duplicate, test row, junk)
  missing       expected data that the source simply does not contain
  flagged       nothing changed, but the entry is unusual (gate-log events)
Any entry a human should look at carries needs_review=1, a priority and the
method for carrying out that review.

Outputs:
  clean_data/school.db            SQLite database
  clean_data/school_clean.xlsx    the same data as a formatted workbook
  clean_data/staging_grades.csv   raw-vs-clean value for every score (for auditing)
  clean_data/pipeline_stats.json  counts used by the showcase
  reports/cleaning_report.md      human-readable cleaning report
"""
import datetime as dt
import difflib
import json
import math
import re
import sqlite3
from collections import Counter, defaultdict
from pathlib import Path

import pandas as pd
import pdfplumber
from docx import Document
from docx.table import Table as DocxTable
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
RAW, OUT, REP = ROOT / "raw_data", ROOT / "clean_data", ROOT / "reports"
OUT.mkdir(exist_ok=True)
REP.mkdir(exist_ok=True)

F_ROSTER, F_FORM, F_GB = "01_student_roster.xlsx", "02_enrollment_form_responses.csv", "03_gradebook_term1.xlsx"
F_SCI, F_HIST, F_LOG = "04_science_department_results.docx", "05_history_gpa_report.pdf", "06_gate_attendance.log"
REF_DATE = dt.date(2026, 9, 1)          # start of the school year, used for age checks

# How each kind of flagged entry should be reviewed, and how urgent it is.
REVIEW = {
    "dob_missing": ("high", "Pull the birth-certificate copy from the admission file and enter the date in the register."),
    "dob_illogical": ("high", "Age is impossible for the grade. Check the birth certificate; the year is the likely typo."),
    "dob_filled_from_form": ("low", "Date taken from the parent's enrollment form. Confirm against the birth certificate at the next file audit."),
    "dob_conflict": ("medium", "Register and enrollment form disagree. The register value is kept; confirm with the birth certificate."),
    "gender_missing": ("low", "Field empty in the register. Complete it from the admission form."),
    "grade_mismatch": ("low", "Grade cell disagrees with the sheet the student is listed on. Sheet value kept; registrar to confirm placement."),
    "duplicate_id_conflict": ("high", "Two students share one ID in the register. ID re-assigned by name match with the enrollment form; registrar must confirm and fix the register."),
    "name_mismatch": ("low", "Spelling on the form differs from the register. Register spelling kept; ask the parent which is correct."),
    "wrong_id": ("medium", "Form carried another student's ID. Re-linked by exact name match; confirm with the parent contact."),
    "unknown_id": ("medium", "ID is not on the register. Check for a late enrolment or a mistyped ID before discarding."),
    "no_form_response": ("medium", "No enrollment form on file, so no guardian contact. Homeroom teacher to send the form again."),
    "scale_gpa_assumed": ("low", "Value between 1 and 4 read as a 4.0-scale GPA and converted with the school policy. Confirm against the previous school's transcript."),
    "letter_grade": ("medium", "Letter grade where a number is required; no exact conversion exists. Ask the teacher/previous school for the numeric mark."),
    "illogical_value": ("medium", "Value outside the possible range. Check the original paper form or mark sheet."),
    "email_invalid": ("medium", "Address cannot receive mail. Phone the guardian to confirm the address."),
    "phone_invalid": ("medium", "Number has the wrong number of digits. Confirm by email or through the student."),
    "text_absent": ("high", "Student was absent for the assessment. Schedule a make-up exam, then enter the mark."),
    "text_not_available": ("medium", "No mark recorded (N/A, pending, exempt, incomplete). Subject teacher to confirm the final status."),
    "total_mismatch": ("medium", "Total does not equal coursework + exam. Component sum used; teacher to check the paper mark sheet."),
    "component_out_of_range": ("medium", "A component exceeds its maximum mark. Total kept as recorded; teacher to re-enter the component."),
    "duplicate_conflict": ("high", "Two rows with different marks for one student. Score withheld; teacher to say which row is valid."),
    "row_missing": ("medium", "Student is on the register but has no row in this mark list. Subject teacher to supply the mark or the reason."),
    "misplaced_decimal": ("medium", "Value fits neither the 0-1 nor the 0-100 scale; probably a misplaced decimal point. Teacher to confirm (suggested value in the detail column)."),
    "out_of_range": ("medium", "Mark is above the maximum of the stated scale. Teacher to confirm the value and the scale used."),
    "id_missing_name_typo": ("low", "Row had no ID and the name did not match exactly; linked by closest name within the grade. Confirm the link."),
    "gpa_over_4": ("medium", "GPA above 4.0 is impossible under school policy. Score withheld; history teacher to re-enter."),
    "scale_percent_in_gpa_column": ("low", "A percentage was typed in a GPA column; converted with the school policy. Teacher to confirm."),
    "id_empty": ("low", "Swipe with an empty ID cannot be attributed. Have the gate reader/card checked."),
    "unknown_id_swipe": ("high", "Card ID not on the register was used at the gate. Security office to identify the card holder."),
    "odd_hour_swipe": ("high", "Check-in between 20:00 and 05:00. Security office to review CCTV for that gate and time."),
    "weekend_swipe": ("medium", "Swipe on a non-school day. Confirm against the activities calendar."),
    "impossible_date": ("low", "Timestamp outside the export period (clock fault). IT to check the gate controller clock."),
    "missing_checkin": ("low", "Checked out without checking in. Homeroom teacher to confirm arrival time from the class register."),
    "missing_checkout": ("low", "No check-out recorded. Confirm the student left through a monitored gate."),
    "early_leave": ("medium", "Left before 12:00. Match against signed early-dismissal slips."),
    "absent_but_swiped": ("medium", "Marked ABSENT by the system but the card was swiped. Counted as present; homeroom teacher to confirm."),
    "no_record": ("medium", "Neither a swipe nor an ABSENT entry that day. Counted as absent; confirm with the class register."),
}

LOG = []            # the cleaning log
FORMATS = []        # free-text fields harmonised in bulk
STATS = {"sources": {}}


def log(source, record, field, raw, clean, action, category, detail="", review=None):
    """review=None -> no review; review=True -> look up by category; review='key' -> look up by that key."""
    pr, method = ("", "")
    if review:
        pr, method = REVIEW[category if review is True else review]
    LOG.append(dict(source=source, record=record or "", field=field,
                    raw_value="" if raw is None else str(raw), clean_value="" if clean is None else str(clean),
                    action=action, category=category, detail=detail,
                    needs_review=int(bool(review)), priority=pr, review_method=method))


# ----------------------------------------------------------------------------
# Shared parsers
# ----------------------------------------------------------------------------
def blank(x):
    return x is None or (isinstance(x, float) and math.isnan(x)) or (isinstance(x, str) and not x.strip())


def norm_id(x):
    """'stu_0042', 'STU-O042', 'STU- 42', 42  ->  'STU-0042'. Returns None if it cannot be an ID."""
    if blank(x):
        return None
    if isinstance(x, (int, float)):
        return f"STU-{int(x):04d}"
    t = re.sub(r"\s+", "", str(x)).upper()
    if t.isdigit() and len(t) <= 4:
        return f"STU-{int(t):04d}"
    m = re.fullmatch(r"STU[-_]?([0-9O]{1,4})", t)
    return f"STU-{int(m.group(1).replace('O', '0')):04d}" if m else None


def norm_name(x):
    """Trim, collapse spaces, Title Case, and turn 'Last, First' into 'First Last'."""
    t = " ".join(str(x).replace("\n", " ").split())
    if "," in t:
        last, first = [p.strip() for p in t.split(",", 1)]
        t = f"{first} {last}"
    return t.title()


def parse_num(x):
    """Return (number or None, kind). kind tells how the text deviated from a plain number."""
    if blank(x):
        return None, "blank"
    if isinstance(x, (int, float)):
        return float(x), "num"
    t = str(x).strip()
    kind = "num" if t == str(x) else "space"
    if re.fullmatch(r"[A-Fa-f][+-]?", t):
        return None, "letter"
    if t.endswith("%"):
        t, kind = t[:-1].strip(), "pct_sign"
    m = re.fullmatch(r"(-?[\d.,]+)\s*/\s*4(\.0)?", t)
    if m:
        t, kind = m.group(1), "of4"
    if re.fullmatch(r"-?\d+,\d+", t):
        t, kind = t.replace(",", "."), "comma"
    try:
        return float(t), kind
    except ValueError:
        return None, "text"


DATE_FORMATS = ["%d/%m/%Y", "%Y-%m-%d", "%B %d, %Y", "%d-%b-%y", "%d.%m.%Y"]


def parse_date(x, formats=DATE_FORMATS):
    if blank(x):
        return None
    if isinstance(x, dt.datetime):
        return x.date()
    if isinstance(x, dt.date):
        return x
    for f in formats:
        try:
            return dt.datetime.strptime(str(x).strip(), f).date()
        except ValueError:
            pass
    return None


def age_ok(dob, grade):
    """A student in grade g is normally g+5 or g+6 years old in September; allow one year either side."""
    age = (REF_DATE - dob).days / 365.25
    return grade + 4 <= age <= grade + 7.5


def gpa_to_pct(g):
    return 50 + 12.5 * g       # school policy printed on the History report


# ----------------------------------------------------------------------------
# A light first pass over the form: name -> ID, used to resolve register conflicts
# ----------------------------------------------------------------------------
form_raw = pd.read_csv(RAW / F_FORM, dtype=str, keep_default_na=False, encoding="utf-8-sig")
form_raw.columns = ["ts", "id", "name", "dob", "grade", "gname", "email", "phone", "prev", "sib", "transport", "medical"]
form_raw["ts"] = pd.to_datetime(form_raw["ts"], format="%m/%d/%Y %H:%M:%S")
form_raw = form_raw.sort_values("ts").reset_index(drop=True)
form_name_to_id = {norm_name(r.name_): norm_id(r.id) for r in form_raw.rename(columns={"name": "name_"}).itertuples()}

# ----------------------------------------------------------------------------
# 1. Student roster (Excel)
# ----------------------------------------------------------------------------
HEAD = {"student id": "id", "student no.": "id", "first name": "first", "last name": "last", "gender": "gender",
        "sex": "gender", "date of birth": "dob", "dob": "dob", "grade": "grade", "section": "section", "sec.": "section"}
GENDER = {"m": "M", "male": "M", "boy": "M", "f": "F", "female": "F", "girl": "F"}

wb = load_workbook(RAW / F_ROSTER, data_only=True)
roster, n_read = [], 0
for ws in wb.worksheets:
    sheet_grade = int(re.search(r"\d+", ws.title).group())
    rows = list(ws.iter_rows(values_only=True))
    # the header is not on a fixed row: find the row that names an ID column
    h = next(i for i, r in enumerate(rows) if any(str(c).strip().lower() in ("student id", "student no.") for c in r if c))
    cols = {HEAD[str(c).strip().lower()]: j for j, c in enumerate(rows[h]) if c and str(c).strip().lower() in HEAD}
    for i, r in enumerate(rows[h + 1:], start=h + 2):
        if all(blank(c) for c in r):
            continue                                   # stray blank row
        if isinstance(r[0], str) and r[0].lower().startswith("total"):
            continue                                   # "Total students: 60" footer
        n_read += 1
        raw = {k: r[j] for k, j in cols.items()}
        rec = dict(sheet=ws.title, row=i, sheet_grade=sheet_grade, raw=raw)
        rec["id"] = norm_id(raw["id"])
        rec["first"], rec["last"] = norm_name(raw["first"]), norm_name(raw["last"])
        rec["name"] = f"{rec['first']} {rec['last']}"
        rec["gender"] = None if blank(raw["gender"]) else GENDER.get(str(raw["gender"]).strip().lower())
        rec["dob"] = parse_date(raw["dob"])
        rec["section"] = None if blank(raw["section"]) else str(raw["section"]).strip().upper().replace("SEC", "").strip()
        rec["grade"] = sheet_grade
        roster.append(rec)

# exact duplicate rows
seen, uniq = set(), []
for rec in roster:
    sig = (rec["id"], rec["name"], rec["dob"])
    if sig in seen:
        log(F_ROSTER, rec["id"], "row", f"{rec['sheet']} row {rec['row']}", None, "omitted", "duplicate_row",
            "Identical row already present; second copy dropped.")
        continue
    seen.add(sig)
    uniq.append(rec)
roster = uniq

# one ID used by two different students -> resolve with the enrollment form
by_id = defaultdict(list)
for rec in roster:
    by_id[rec["id"]].append(rec)
for sid, group in by_id.items():
    if len(group) < 2:
        continue
    taken = set(by_id)
    for rec in group:
        fid = form_name_to_id.get(rec["name"])
        if fid and fid != sid and fid not in taken:
            log(F_ROSTER, fid, "student_id", sid, fid, "corrected", "duplicate_id_conflict",
                f"{rec['name']} was listed with {sid}, which belongs to {[g['name'] for g in group if g is not rec][0]}. "
                f"Enrollment form gives {fid}.", review=True)
            rec["id"] = fid

for rec in roster:   # row-level log of everything else that changed
    raw, sid = rec["raw"], rec["id"]
    if str(raw["id"]) != sid and not any(l["category"] == "duplicate_id_conflict" and l["record"] == sid for l in LOG):
        log(F_ROSTER, sid, "student_id", raw["id"], sid, "standardised", "id_format")
    for k in ("first", "last"):
        if raw[k] != rec[k]:
            log(F_ROSTER, sid, f"{k}_name", repr(raw[k]), rec[k], "standardised", "name_case_space")
    if blank(raw["gender"]):
        log(F_ROSTER, sid, "gender", "", None, "missing", "gender_missing", review=True)
    elif raw["gender"] not in ("M", "F", "Male", "Female"):
        log(F_ROSTER, sid, "gender", repr(raw["gender"]), rec["gender"], "standardised", "gender_variant")
    if not blank(raw["section"]) and raw["section"] != rec["section"]:
        log(F_ROSTER, sid, "section", raw["section"], rec["section"], "standardised", "section_format")
    if not blank(raw["grade"]) and int(raw["grade"]) != rec["sheet_grade"]:
        log(F_ROSTER, sid, "grade", raw["grade"], rec["sheet_grade"], "corrected", "grade_mismatch",
            f"Listed on sheet '{rec['sheet']}'.", review=True)
    if isinstance(raw["dob"], str) and rec["dob"]:
        log(F_ROSTER, sid, "dob", raw["dob"], rec["dob"], "standardised", "dob_text_format",
            "Text date parsed (day-first, as used throughout the register).")

STATS["sources"][F_ROSTER] = dict(format="Excel student lists", records_read=n_read, records_kept=len(roster))
ROSTER = {r["id"]: r for r in roster}
NAME_TO_ID = {r["name"]: r["id"] for r in roster}

# ----------------------------------------------------------------------------
# 2. Enrollment form (CSV)
# ----------------------------------------------------------------------------
WORDNUM = {"none": 0, "zero": 0, "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
           "ninth": 9, "tenth": 10, "eleventh": 11, "twelfth": 12}
EMAIL_RE = re.compile(r"^[\w.+-]+@[\w-]+(\.[\w-]+)+$")

kept = {}
fmt_count = defaultdict(Counter)
for r in form_raw.itertuples():
    rowlabel = f"response {r.ts:%Y-%m-%d %H:%M}"
    if "test" in r.name.lower() or r.gname.strip().lower() == "test":
        log(F_FORM, r.id, "row", f"{r.name} ({rowlabel})", None, "omitted", "test_submission", "Test entry, not a real family.")
        continue
    sid, name = norm_id(r.id), norm_name(r.name)
    if name in NAME_TO_ID and NAME_TO_ID[name] != sid:
        log(F_FORM, NAME_TO_ID[name], "student_id", r.id, NAME_TO_ID[name], "corrected", "wrong_id",
            f"Form says {r.id} but that ID belongs to {ROSTER[sid]['name'] if sid in ROSTER else 'nobody'}; "
            f"'{name}' is {NAME_TO_ID[name]} on the register.", review=True)
        sid = NAME_TO_ID[name]
    elif sid not in ROSTER:
        log(F_FORM, sid or r.id, "student_id", f"{r.id} / {name}", None, "omitted", "unknown_id",
            "Neither the ID nor the name is on the register; response held back.", review=True)
        continue
    elif name != ROSTER[sid]["name"]:
        sim = difflib.SequenceMatcher(None, name, ROSTER[sid]["name"]).ratio()
        if sim < 0.8:
            log(F_FORM, sid, "student_id", f"{r.id} / {name}", None, "omitted", "unknown_id",
                f"ID belongs to {ROSTER[sid]['name']}, name does not match anyone; response held back.", review=True)
            continue
        log(F_FORM, sid, "name", name, ROSTER[sid]["name"], "corrected", "name_mismatch",
            f"Similarity {sim:.2f}; register spelling kept.", review=True)
    if sid in kept:      # rows are in time order, so the newer one replaces the older
        log(F_FORM, sid, "row", f"response {kept[sid].ts:%Y-%m-%d %H:%M}", None, "omitted", "duplicate_submission",
            f"Superseded by the later {rowlabel}.")
    kept[sid] = r

guardians = []
for sid, r in kept.items():
    g = dict(student_id=sid, guardian_name=norm_name(r.gname), form_timestamp=r.ts.isoformat(sep=" "),
             medical_notes=None if r.medical.strip().lower() in ("", "none") else r.medical.strip())

    # grade (free text) - only used as a cross-check on the register
    t = r.grade.strip().lower()
    fmt_count["grade"][re.sub(r"\d+", "N", t) if re.search(r"\d", t) else "word"] += 1
    fg = WORDNUM.get(t) or int(re.search(r"\d+", t).group())
    if t in WORDNUM:
        log(F_FORM, sid, "grade", r.grade, fg, "corrected", "number_as_word")
    if fg != ROSTER[sid]["grade"]:
        log(F_FORM, sid, "grade", r.grade, ROSTER[sid]["grade"], "corrected", "grade_mismatch", "Form and register disagree.", review=True)

    # date of birth (MM/DD/YYYY as the form states) - cross-check / fill the register
    g["_dob"] = parse_date(r.dob, ["%m/%d/%Y"])

    # previous school average: free text, several scales
    v, kind = parse_num(r.prev)
    pct = None
    if kind == "blank":
        log(F_FORM, sid, "prev_school_avg_pct", "", None, "missing", "blank")
    elif kind == "letter":
        log(F_FORM, sid, "prev_school_avg_pct", r.prev, None, "omitted", "letter_grade", review=True)
    elif kind == "of4" or (kind != "pct_sign" and 1 < v <= 4):
        pct = gpa_to_pct(v)
        log(F_FORM, sid, "prev_school_avg_pct", r.prev, pct, "corrected", "scale_gpa",
            "Read as GPA on a 4.0 scale; Percent = 50 + 12.5 x GPA.", review=None if kind == "of4" else "scale_gpa_assumed")
    elif kind != "pct_sign" and 0 < v <= 1:
        pct = round(v * 100, 1)
        log(F_FORM, sid, "prev_school_avg_pct", r.prev, pct, "corrected", "scale_fraction", "Fraction of 1 multiplied by 100.")
    elif 30 <= v <= 100:
        pct = v
        if kind == "pct_sign":
            log(F_FORM, sid, "prev_school_avg_pct", r.prev, pct, "standardised", "scale_percent_sign")
    else:
        log(F_FORM, sid, "prev_school_avg_pct", r.prev, None, "omitted", "illogical_value",
            "Not a possible percentage, fraction or GPA.", review=True)
    g["prev_school_avg_pct"] = pct

    # siblings
    t = r.sib.strip().lower()
    sib = None
    if t == "":
        log(F_FORM, sid, "siblings", "", None, "missing", "blank")
    elif t in WORDNUM:
        sib = WORDNUM[t]
        log(F_FORM, sid, "siblings", r.sib, sib, "corrected", "number_as_word")
    elif re.fullmatch(r"-?\d+", t):
        if 0 <= int(t) <= 12:
            sib = int(t)
        else:
            log(F_FORM, sid, "siblings", r.sib, None, "omitted", "illogical_value", "Outside 0-12.", review=True)
    else:
        sib = int(re.search(r"\d+", t).group())
        log(F_FORM, sid, "siblings", r.sib, sib, "corrected", "number_with_text", "Number extracted from text.")
    g["siblings"] = sib

    # e-mail
    e = r.email.strip().lower()
    email = None
    if e == "":
        log(F_FORM, sid, "guardian_email", "", None, "missing", "blank", review="email_invalid")
    elif EMAIL_RE.match(e):
        email = e
        if e != r.email:
            log(F_FORM, sid, "guardian_email", repr(r.email), e, "standardised", "email_case_space")
    elif EMAIL_RE.match(e.replace(",", ".")):
        email = e.replace(",", ".")
        log(F_FORM, sid, "guardian_email", r.email, email, "corrected", "email_typo", "Comma typed instead of a dot.")
    else:
        log(F_FORM, sid, "guardian_email", r.email, None, "omitted", "email_invalid", review=True)
    g["guardian_email"] = email

    # phone: any punctuation -> +1-AAA-555-NNNN
    digits = re.sub(r"\D", "", r.phone)
    fmt_count["phone"][re.sub(r"\d", "N", r.phone.strip())] += 1
    if len(digits) == 11 and digits[0] == "1":
        digits = digits[1:]
    phone = None
    if r.phone.strip() == "":
        log(F_FORM, sid, "guardian_phone", "", None, "missing", "blank", review="phone_invalid")
    elif len(digits) == 10:
        phone = f"+1-{digits[:3]}-{digits[3:6]}-{digits[6:]}"
    else:
        log(F_FORM, sid, "guardian_phone", r.phone, None, "omitted", "phone_invalid", review=True)
    g["guardian_phone"] = phone

    fmt_count["transport"][r.transport] += 1
    g["transport"] = r.transport.strip().title()
    guardians.append(g)

for field, c in fmt_count.items():
    FORMATS.append(dict(source=F_FORM, field=field, values=sum(c.values()), distinct_formats=len(c),
                        examples=" | ".join(k for k, _ in c.most_common(5))))
for sid in ROSTER:
    if sid not in kept:
        log(F_FORM, sid, "row", "", None, "missing", "no_form_response", review=True)
STATS["sources"][F_FORM] = dict(format="Online form (CSV)", records_read=len(form_raw), records_kept=len(guardians))

# use the form to repair / cross-check dates of birth in the register
FORM_DOB = {g["student_id"]: g.pop("_dob") for g in guardians}
for sid, rec in ROSTER.items():
    fd, raw = FORM_DOB.get(sid), rec["raw"]["dob"]
    rec["dob_source"] = "register"
    if rec["dob"] is None or not age_ok(rec["dob"], rec["grade"]):
        cat = "dob_missing" if rec["dob"] is None else "dob_illogical"
        shown = "" if rec["dob"] is None else rec["dob"]
        if fd and age_ok(fd, rec["grade"]):
            log(F_ROSTER, sid, "dob", shown, fd, "corrected", cat,
                "Register date " + ("empty" if cat == "dob_missing" else f"implies age {(REF_DATE - rec['dob']).days // 365} in grade {rec['grade']}")
                + "; filled from the enrollment form.", review="dob_filled_from_form")
            rec["dob"], rec["dob_source"] = fd, "enrollment form"
        else:
            log(F_ROSTER, sid, "dob", shown, None, "missing" if cat == "dob_missing" else "omitted", cat,
                "No usable date on the enrollment form either.", review=True)
            rec["dob"], rec["dob_source"] = None, None
    elif fd and fd != rec["dob"]:
        log(F_FORM, sid, "dob", fd, rec["dob"], "corrected", "dob_conflict", f"Register says {rec['dob']}, form says {fd}.", review=True)

# ----------------------------------------------------------------------------
# 3-5. Grades: three sources, three scales -> one percentage
# ----------------------------------------------------------------------------
GRADES = []     # one row per score as found (after de-duplication)


def add_grade(sid, subject, source, raw_id, raw_value, scale, pct, status):
    GRADES.append(dict(student_id=sid, subject=subject, source_file=source, raw_id="" if raw_id is None else str(raw_id),
                       raw_value="" if raw_value is None else str(raw_value), native_scale=scale,
                       score_pct=None if pct is None else round(pct, 1), status=status))


def resolve_student(source, raw_id, raw_name, grade=None):
    """ID if valid, else exact name, else closest name within the grade."""
    sid = norm_id(raw_id)
    if sid in ROSTER:
        if str(raw_id) != sid:
            log(source, sid, "student_id", raw_id, sid, "standardised", "id_letter_O" if "O" in str(raw_id).upper()[3:] else "id_format")
        return sid
    name = norm_name(raw_name)
    if name in NAME_TO_ID:
        log(source, NAME_TO_ID[name], "student_id", raw_id or "", NAME_TO_ID[name], "corrected", "id_missing", f"Matched by name '{name}'.")
        return NAME_TO_ID[name]
    pool = [n for n, i in NAME_TO_ID.items() if grade is None or ROSTER[i]["grade"] == grade]
    close = difflib.get_close_matches(name, pool, n=1, cutoff=0.85)
    if close:
        sid = NAME_TO_ID[close[0]]
        log(source, sid, "student_id", f"{raw_id or ''} / {name}", sid, "corrected", "id_missing_name_typo",
            f"Closest register name: '{close[0]}'.", review=True)
        return sid
    log(source, raw_id or name, "student_id", f"{raw_id} / {name}", None, "omitted", "unknown_id", "Cannot be linked to the register.", review=True)
    return None


# ---- 3. Gradebook (Excel): coursework/40 + exam/60 = total/100
n_gb = 0
for subject in ("Math", "English"):
    df = pd.read_excel(RAW / F_GB, sheet_name=subject, header=2, dtype=object)
    df.columns = ["id", "name", "cls", "cw", "exam", "total", "remarks"]
    for r in df.itertuples():
        if blank(r.name) or str(r.name).strip().lower() == "class average":
            continue
        n_gb += 1
        sid = resolve_student(F_GB, r.id, r.name)
        if sid is None:
            continue
        cw, _ = parse_num(r.cw)
        ex, _ = parse_num(r.exam)
        tot, kind = parse_num(r.total)
        comps_ok = cw is not None and ex is not None and 0 <= cw <= 40 and 0 <= ex <= 60
        s = cw + ex if comps_ok else None
        pct, status = None, "ok"
        if kind == "text":
            cat = "text_absent" if str(r.total).strip().lower().startswith("abs") else "text_not_available"
            log(F_GB, sid, subject, r.total, None, "omitted", cat, review=True)
            status = "missing"
        elif comps_ok and tot is None:
            pct, status = s, "corrected"
            log(F_GB, sid, subject, "", s, "corrected", "total_blank", f"Total empty; computed {cw:g} + {ex:g}.")
        elif comps_ok and 0 < tot <= 1 and abs(tot * 100 - s) <= 1:
            pct, status = s, "corrected"
            log(F_GB, sid, subject, r.total, s, "corrected", "scale_fraction", "Total typed as a fraction of 1; components confirm the value.")
        elif comps_ok and tot == s:
            pct = tot
            if isinstance(r.total, str):
                status = "standardised"
                log(F_GB, sid, subject, repr(r.total), tot, "standardised", "number_as_text")
        elif comps_ok and not 0 <= tot <= 100:
            pct, status = s, "corrected"
            log(F_GB, sid, subject, r.total, s, "corrected", "illogical_total", f"Impossible total; replaced by {cw:g} + {ex:g}.")
        elif comps_ok:
            pct, status = s, "corrected"
            log(F_GB, sid, subject, r.total, s, "corrected", "total_mismatch", f"Total {tot:g} but {cw:g} + {ex:g} = {s:g}.", review=True)
        elif tot is not None and 0 <= tot <= 100:
            pct = tot
            log(F_GB, sid, subject, f"cw={r.cw}, exam={r.exam}", tot, "corrected", "component_out_of_range",
                "Component above its maximum; recorded total kept.", review=True)
        else:
            status = "review"
            log(F_GB, sid, subject, r.total, None, "omitted", "illogical_value", review=True)
        add_grade(sid, subject, F_GB, r.id, r.total, "percent (0-100)", pct, status)
STATS["sources"][F_GB] = dict(format="Excel gradebook", records_read=n_gb)

# ---- 4. Science (Word tables): 0-1, except where the header says /20
doc = Document(RAW / F_SCI)
grade_now, n_sci = None, 0
for el in doc.element.body.iterchildren():
    tag = el.tag.split("}")[1]
    if tag == "p":
        m = re.fullmatch(r"Grade (\d+)", "".join(el.itertext()).strip())
        grade_now = int(m.group(1)) if m else grade_now
    elif tag == "tbl":
        tb = DocxTable(el, doc)
        header = [c.text.strip() for c in tb.rows[0].cells]
        full_mark = 20 if "/20" in header[3] else 1        # the scale is read from the header, not assumed
        scale = "out of 20" if full_mark == 20 else "fraction (0-1)"
        for row in tb.rows[1:]:
            cells = [c.text.strip() for c in row.cells]
            if len(set(cells)) == 1:
                continue                                   # merged "Section A" banner or empty row
            n_sci += 1
            _, rid, rname, rscore, _ = cells
            sid = resolve_student(F_SCI, rid, rname, grade_now)
            if sid is None:
                continue
            v, kind = parse_num(rscore)
            pct, status = None, "ok"
            if kind in ("blank", "text"):
                status = "missing"
                log(F_SCI, sid, "Science", rscore, None, "omitted", "text_not_available", review=True)
            elif kind == "letter":
                status = "review"
                log(F_SCI, sid, "Science", rscore, None, "omitted", "letter_grade", review=True)
            elif kind == "pct_sign" or (full_mark == 1 and 10 < v <= 100):
                pct, status = v, "corrected"
                log(F_SCI, sid, "Science", rscore, v, "corrected", "scale_percent_in_fraction_column",
                    f"Percentage typed in a '{scale}' column.")
            elif 0 <= v <= full_mark:
                pct = v * 100 / full_mark
                if kind == "comma":
                    status = "standardised"
                    log(F_SCI, sid, "Science", rscore, round(pct, 1), "standardised", "comma_decimal")
            elif full_mark == 1 and 1 < v <= 10:
                status = "review"
                log(F_SCI, sid, "Science", rscore, None, "omitted", "misplaced_decimal", f"Suggested: {v / 10:.2f} ({v * 10:g}%).", review=True)
            else:
                status = "review"
                log(F_SCI, sid, "Science", rscore, None, "omitted", "out_of_range", f"Scale is {scale}.", review=True)
            add_grade(sid, "Science", F_SCI, rid, rscore, scale, pct, status)
STATS["sources"][F_SCI] = dict(format="Word tables", records_read=n_sci)

# ---- 5. History (PDF table): GPA on 4.0
n_hist, hist_seen = 0, {}
with pdfplumber.open(RAW / F_HIST) as pdf:
    STATS["pdf_pages"] = len(pdf.pages)
    for page in pdf.pages:
        for tb in page.extract_tables():
            for cells in tb:
                cells = [(c or "").strip() for c in cells]
                if cells[0] == "Student ID" or not any(cells):
                    continue                               # header repeated on every page
                n_hist += 1
                rid, rname, _, _, rgpa, _ = cells
                key = (norm_id(rid), rgpa)
                if key in hist_seen:
                    log(F_HIST, key[0], "row", f"{rid} | {rgpa}", None, "omitted", "duplicate_row", "Identical row already read.")
                    continue
                hist_seen[key] = True
                sid = resolve_student(F_HIST, rid, rname)
                if sid is None:
                    continue
                v, kind = parse_num(rgpa)
                pct, status = None, "ok"
                if kind in ("blank", "text", "letter"):
                    status = "missing"
                    log(F_HIST, sid, "History", rgpa, None, "omitted", "text_not_available", review=True)
                elif 0 <= v <= 4:
                    pct = gpa_to_pct(v)
                    if kind in ("comma", "of4"):
                        status = "standardised"
                        log(F_HIST, sid, "History", rgpa, v, "standardised", "comma_decimal" if kind == "comma" else "number_as_text")
                elif 4 < v <= 5:
                    status = "review"
                    log(F_HIST, sid, "History", rgpa, None, "omitted", "gpa_over_4", review=True)
                elif 50 <= v <= 100:
                    gpa = round((v - 50) / 12.5, 1)        # the report carries GPA to one decimal
                    pct, status = gpa_to_pct(gpa), "corrected"
                    log(F_HIST, sid, "History", rgpa, f"GPA {gpa:.1f}", "corrected", "scale_percent_in_gpa_column",
                        "Percentage typed in the GPA column; converted back to the nearest 0.1 GPA.", review=True)
                else:
                    status = "review"
                    log(F_HIST, sid, "History", rgpa, None, "omitted", "illogical_value", review=True)
                add_grade(sid, "History", F_HIST, rid, rgpa, "GPA (0-4)", pct, status)
STATS["sources"][F_HIST] = dict(format="PDF table", records_read=n_hist)

# ---- same student twice in one subject, and students with no row at all
grades_df = pd.DataFrame(GRADES)
final = []
for (sid, subject), grp in grades_df.groupby(["student_id", "subject"], sort=False):
    row = grp.iloc[0].to_dict()
    if len(grp) > 1:
        if grp["score_pct"].nunique(dropna=False) == 1:
            log(row["source_file"], sid, subject, "repeated row", None, "omitted", "duplicate_row", "Same mark twice; one copy kept.")
        else:
            row.update(score_pct=None, status="review", raw_value=" vs ".join(grp["raw_value"]))
            log(row["source_file"], sid, subject, row["raw_value"], None, "omitted", "duplicate_conflict",
                f"Rows give {', '.join(f'{x:g}' for x in grp['score_pct'])}.", review=True)
    final.append(row)
have = {(r["student_id"], r["subject"]) for r in final}
for subject, src, scale in [("Math", F_GB, "percent (0-100)"), ("English", F_GB, "percent (0-100)"),
                            ("Science", F_SCI, "fraction (0-1)"), ("History", F_HIST, "GPA (0-4)")]:
    for sid in ROSTER:
        if (sid, subject) not in have:
            log(src, sid, subject, "", None, "missing", "row_missing", review=True)
            final.append(dict(student_id=sid, subject=subject, source_file=src, raw_id="", raw_value="", native_scale=scale,
                              score_pct=None, status="missing"))
grades_df = pd.DataFrame(final)
for src in (F_GB, F_SCI, F_HIST):
    STATS["sources"][src]["records_kept"] = int((grades_df[grades_df.source_file == src].score_pct.notna()).sum())

# ----------------------------------------------------------------------------
# 6. Gate log -> daily attendance + flagged entries
# ----------------------------------------------------------------------------
LINE = re.compile(r"^(\d{4}-\d{2}-\d{2}|\d{2}/\d{2}/\d{4}) (\d{2}:\d{2}:\d{2}) \[(\w+)\] (\S+) (CHECKIN|CHECKOUT|ABSENT) "
                  r"id=(\S*?)(?: card=(\S+))?(?: status=(OK|LATE)| reason=(\w+))$")
NOISE = re.compile(r"^\S+ \S+ \[\w+\] \S+ (heartbeat|reader timeout|sync with SIS|door sensor)")
FLAGS = []


def flag(cat, sid, when, detail, line_no=None):
    pr, method = REVIEW[cat]
    FLAGS.append(dict(category=cat, student_id=sid or "", date=str(when.date() if isinstance(when, dt.datetime) else when),
                      time="" if not isinstance(when, dt.datetime) else f"{when:%H:%M:%S}", detail=detail,
                      line_no=line_no, priority=pr, review_method=method))


events, n_lines, n_noise, n_alt, n_bad = [], 0, 0, 0, 0
with open(RAW / F_LOG, encoding="utf-8") as fh:
    for no, line in enumerate(fh, 1):
        line = line.rstrip("\n")
        if line.startswith("#"):
            start, end = [dt.date.fromisoformat(x) for x in re.findall(r"\d{4}-\d{2}-\d{2}", line)]
            continue
        n_lines += 1
        if NOISE.match(line):
            n_noise += 1
            continue
        m = LINE.match(line)
        if not m:
            n_bad += 1
            log(F_LOG, f"line {no}", "line", line, None, "omitted", "truncated_line", "Line is cut off or malformed.")
            continue
        d, t, lvl, gate, ev, rid, card, status, reason = m.groups()
        if "/" in d:
            n_alt += 1
            ts = dt.datetime.strptime(f"{d} {t}", "%d/%m/%Y %H:%M:%S")
        else:
            ts = dt.datetime.strptime(f"{d} {t}", "%Y-%m-%d %H:%M:%S")
        if not start <= ts.date() <= end:
            log(F_LOG, f"line {no}", "timestamp", f"{d} {t}", None, "omitted", "impossible_date", "Outside the export period.", review=True)
            flag("impossible_date", norm_id(rid), ts, f"{gate} {ev}", no)
            continue
        if rid == "":
            log(F_LOG, f"line {no}", "student_id", "id=", None, "omitted", "id_empty", f"{gate} {ev} at {ts}", review=True)
            flag("id_empty", None, ts, f"{gate} {ev} with empty ID", no)
            continue
        sid = norm_id(rid)
        if sid not in ROSTER:
            log(F_LOG, f"line {no}", "student_id", rid, None, "omitted", "unknown_id", f"{gate} {ev} at {ts}", review="unknown_id_swipe")
            flag("unknown_id_swipe", sid or rid, ts, f"{gate} {ev}, card={card}", no)
            continue
        if rid != sid:
            log(F_LOG, sid, "student_id", rid, sid, "standardised", "id_format", f"line {no}")
        if ev != "ABSENT":
            if ts.weekday() >= 5:
                log(F_LOG, sid, "line", f"{ts} {ev}", None, "omitted", "weekend_swipe", f"{ts:%A}; not counted as attendance.", review=True)
                flag("weekend_swipe", sid, ts, f"{gate} {ev} on a {ts:%A}", no)
                continue
            if ts.hour < 5 or ts.hour >= 20:
                log(F_LOG, sid, "line", f"{ts} {ev}", None, "omitted", "odd_hour_swipe", "Not counted as attendance.", review=True)
                flag("odd_hour_swipe", sid, ts, f"{gate} {ev}", no)
                continue
        events.append(dict(ts=ts, sid=sid, ev=ev, status=status, reason=reason, gate=gate, no=no))
if n_alt:
    log(F_LOG, "GATE-3", "timestamp", f"{n_alt} lines as DD/MM/YYYY", "YYYY-MM-DD", "standardised", "timestamp_format",
        "One gate wrote day-first timestamps for a day.")

ev = pd.DataFrame(events).sort_values(["sid", "ev", "ts"]).reset_index(drop=True)
gap = ev.groupby(["sid", "ev"])["ts"].diff().dt.total_seconds()
for r in ev[gap < 60].itertuples():
    log(F_LOG, r.sid, "line", f"{r.ts} {r.ev}", None, "omitted", "duplicate_swipe", f"Repeated within 60 s (line {r.no}).")
ev = ev[~(gap < 60)]

school_days = [start + dt.timedelta(days=i) for i in range((end - start).days + 1)]
school_days = [d for d in school_days if d.weekday() < 5]
ev["day"] = ev.ts.dt.date
idx = {k: g for k, g in ev.groupby(["sid", "day"])}
daily = []
for sid in ROSTER:
    for day in school_days:
        g = idx.get((sid, day))
        ins = g[g.ev == "CHECKIN"] if g is not None else []
        outs = g[g.ev == "CHECKOUT"] if g is not None else []
        absent = g is not None and (g.ev == "ABSENT").any()
        notes, first_in, last_out = [], None, None
        if len(ins) or len(outs):
            status = "present"
            if len(ins):
                first_in = ins.ts.min()
                if first_in.time() > dt.time(8, 0):
                    status = "late"
                    if ins.iloc[0].status == "OK":
                        log(F_LOG, sid, "status", f"{first_in:%Y-%m-%d %H:%M} status=OK", "late", "corrected", "late_not_marked",
                            "Arrived after 08:00 but the gate recorded OK.")
            else:
                notes.append("missing_checkin")
            if len(outs):
                last_out = outs.ts.max()
                if last_out.time() < dt.time(12, 0):
                    notes.append("early_leave")
            else:
                notes.append("missing_checkout")
            if absent:
                notes.append("absent_but_swiped")
        elif absent:
            status = "absent"
        else:
            status = "absent"
            notes.append("no_record")
        for cat in notes:
            when = dt.datetime.combine(day, (last_out if cat == "early_leave" else first_in or last_out or dt.datetime.min).time()) \
                if cat != "no_record" else day
            flag(cat, sid, when, {"missing_checkin": "CHECKOUT without CHECKIN", "missing_checkout": "CHECKIN without CHECKOUT",
                                  "early_leave": f"left at {last_out:%H:%M}" if last_out else "",
                                  "absent_but_swiped": "ABSENT entry and a swipe on the same day",
                                  "no_record": "no log entry of any kind"}[cat])
            log(F_LOG, sid, str(day), FLAGS[-1]["detail"], status, "corrected" if cat in ("absent_but_swiped", "no_record") else "flagged",
                cat, "Attendance status decided by rule; see review method.", review=True)
        daily.append(dict(student_id=sid, date=str(day), status=status,
                          first_in=None if first_in is None else f"{first_in:%H:%M:%S}",
                          last_out=None if last_out is None else f"{last_out:%H:%M:%S}", flags=", ".join(notes)))
daily_df = pd.DataFrame(daily)
att = daily_df.groupby("student_id").status.value_counts().unstack(fill_value=0).reindex(columns=["present", "late", "absent"], fill_value=0)
att_df = pd.DataFrame(dict(school_days=len(school_days), days_present=att.present + att.late, days_late=att.late,
                           days_absent=att.absent)).reset_index()
att_df["attendance_rate"] = (att_df.days_present / att_df.school_days).round(3)
flags_df = pd.DataFrame(FLAGS)
STATS["sources"][F_LOG] = dict(format="System log", records_read=n_lines, records_kept=len(ev))
STATS["log"] = dict(lines=n_lines, noise=n_noise, malformed=n_bad, alt_timestamp=n_alt, events_used=len(ev),
                    school_days=len(school_days), flags=len(flags_df))

# ----------------------------------------------------------------------------
# 7. Compile: tables -> SQLite + Excel + report
# ----------------------------------------------------------------------------
students_df = pd.DataFrame([dict(student_id=r["id"], first_name=r["first"], last_name=r["last"], gender=r["gender"],
                                 date_of_birth=None if r["dob"] is None else r["dob"].isoformat(), dob_source=r["dob_source"],
                                 grade=r["grade"], section=r["section"]) for r in ROSTER.values()]).sort_values("student_id")
guard_df = pd.DataFrame(guardians).sort_values("student_id")[
    ["student_id", "guardian_name", "guardian_email", "guardian_phone", "prev_school_avg_pct", "siblings", "transport",
     "medical_notes", "form_timestamp"]]
grades_df = grades_df.sort_values(["subject", "student_id"]).reset_index(drop=True)
grades_out = grades_df[["student_id", "subject", "score_pct", "status", "raw_value", "native_scale", "source_file"]]
log_df = pd.DataFrame(LOG)
log_df.insert(0, "log_id", range(1, len(log_df) + 1))
PR = {"high": 0, "medium": 1, "low": 2}
review_df = log_df[log_df.needs_review == 1].copy()
review_df = review_df.sort_values(by=["priority", "source", "record"], key=lambda c: c.map(PR) if c.name == "priority" else c)
review_df = review_df[["log_id", "priority", "source", "record", "field", "category", "raw_value", "clean_value", "detail", "review_method"]]
formats_df = pd.DataFrame(FORMATS)

wide = grades_df.pivot(index="student_id", columns="subject", values="score_pct")[["Math", "English", "Science", "History"]]
overview = students_df.merge(wide.reset_index(), on="student_id").merge(att_df, on="student_id")
overview["subjects_available"] = overview[["Math", "English", "Science", "History"]].notna().sum(axis=1)
overview["average_pct"] = overview[["Math", "English", "Science", "History"]].mean(axis=1).round(1)
overview["open_review_items"] = overview.student_id.map(review_df.record.value_counts()).fillna(0).astype(int)

db = OUT / "school.db"
db.unlink(missing_ok=True)
con = sqlite3.connect(db)
con.executescript("""
CREATE TABLE students (student_id TEXT PRIMARY KEY, first_name TEXT NOT NULL, last_name TEXT NOT NULL,
  gender TEXT CHECK (gender IN ('M','F')), date_of_birth TEXT, dob_source TEXT,
  grade INTEGER NOT NULL CHECK (grade BETWEEN 9 AND 12), section TEXT NOT NULL);
CREATE TABLE guardians (student_id TEXT PRIMARY KEY REFERENCES students, guardian_name TEXT, guardian_email TEXT,
  guardian_phone TEXT, prev_school_avg_pct REAL CHECK (prev_school_avg_pct BETWEEN 0 AND 100),
  siblings INTEGER CHECK (siblings BETWEEN 0 AND 12), transport TEXT, medical_notes TEXT, form_timestamp TEXT);
CREATE TABLE grades (student_id TEXT NOT NULL REFERENCES students, subject TEXT NOT NULL,
  score_pct REAL CHECK (score_pct BETWEEN 0 AND 100), status TEXT NOT NULL, raw_value TEXT, native_scale TEXT,
  source_file TEXT, PRIMARY KEY (student_id, subject));
CREATE TABLE attendance_daily (student_id TEXT NOT NULL REFERENCES students, date TEXT NOT NULL,
  status TEXT NOT NULL CHECK (status IN ('present','late','absent')), first_in TEXT, last_out TEXT, flags TEXT,
  PRIMARY KEY (student_id, date));
CREATE TABLE attendance_summary (student_id TEXT PRIMARY KEY REFERENCES students, school_days INTEGER,
  days_present INTEGER, days_late INTEGER, days_absent INTEGER, attendance_rate REAL);
""")
for name, df in [("students", students_df), ("guardians", guard_df), ("grades", grades_out), ("attendance_daily", daily_df),
                 ("attendance_summary", att_df)]:
    df.to_sql(name, con, if_exists="append", index=False)
for name, df in [("log_flags", flags_df), ("cleaning_log", log_df), ("review_queue", review_df), ("format_harmonisation", formats_df)]:
    df.to_sql(name, con, index=False)
con.executescript("""
CREATE VIEW v_student_overview AS
SELECT s.student_id, s.first_name, s.last_name, s.grade, s.section,
       MAX(CASE WHEN g.subject='Math' THEN g.score_pct END) AS math,
       MAX(CASE WHEN g.subject='English' THEN g.score_pct END) AS english,
       MAX(CASE WHEN g.subject='Science' THEN g.score_pct END) AS science,
       MAX(CASE WHEN g.subject='History' THEN g.score_pct END) AS history,
       ROUND(AVG(g.score_pct), 1) AS average_pct, a.attendance_rate, a.days_late
FROM students s LEFT JOIN grades g USING (student_id) LEFT JOIN attendance_summary a USING (student_id)
GROUP BY s.student_id;
""")
con.commit()
con.close()

grades_df.to_csv(OUT / "staging_grades.csv", index=False)

# ---- Excel workbook
act = log_df.groupby(["source", "action"]).size().unstack(fill_value=0).reindex(
    columns=["standardised", "corrected", "omitted", "missing", "flagged"], fill_value=0)
act["needs_review"] = log_df.groupby("source").needs_review.sum()
summary = pd.DataFrame(STATS["sources"]).T.join(act).reset_index(names="source_file")
corr = log_df[log_df.action.isin(["standardised", "corrected"])].drop(columns=["needs_review", "priority", "review_method"])
omit = log_df[log_df.action.isin(["omitted", "missing", "flagged"])].drop(columns=["clean_value", "priority", "review_method"])
sheets = {"Summary": summary, "Student Overview": overview, "Students": students_df, "Guardians": guard_df,
          "Grades": grades_out, "Attendance Summary": att_df, "Corrections": corr, "Omissions": omit,
          "Needs Review": review_df, "Log Flags": flags_df, "Format Harmonisation": formats_df}
xlsx = OUT / "school_clean.xlsx"
with pd.ExcelWriter(xlsx, engine="openpyxl") as xw:
    for name, df in sheets.items():
        df.to_excel(xw, sheet_name=name, index=False)
        ws = xw.sheets[name]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for j, col in enumerate(df.columns, 1):
            c = ws.cell(1, j)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1C5CAB")
            c.alignment = Alignment(vertical="center", wrap_text=True)
            width = max([len(str(col))] + [len(str(v)) for v in df[col].head(300)]) + 2
            ws.column_dimensions[get_column_letter(j)].width = min(max(width, 9), 60)
        if name == "Needs Review":
            tint = {"high": "F8D7D7", "medium": "FDECC8", "low": "E3EEDC"}
            for i, p in enumerate(df.priority, 2):
                ws.cell(i, 2).fill = PatternFill("solid", fgColor=tint[p])

# ---- stats + report
n_cells = len(log_df)
STATS.update(students=len(students_df), actions=log_df.action.value_counts().to_dict(),
             review=review_df.priority.value_counts().to_dict(),
             grades_available=int(grades_df.score_pct.notna().sum()), grades_expected=len(grades_df),
             format_harmonisation=FORMATS)
(OUT / "pipeline_stats.json").write_text(json.dumps(STATS, indent=2, default=str), encoding="utf-8")


def md(df):
    return df.to_markdown(index=False)


top_review = (review_df.groupby(["priority", "category"]).agg(entries=("log_id", "size"), method=("review_method", "first"))
              .reset_index().sort_values(by=["priority", "entries"], key=lambda c: c.map(PR) if c.name == "priority" else -c))
cats = log_df.groupby(["source", "action", "category"]).size().reset_index(name="entries")
rep = f"""# Cleaning report - Greenfield Academy, Term 1 2026/27

Generated by `scripts/02_clean_and_compile.py`. Every row behind these numbers is in the `cleaning_log`
table of `clean_data/school.db` and on the *Corrections*, *Omissions* and *Needs Review* sheets of
`clean_data/school_clean.xlsx`.

## 1. What was read and what was done

{md(summary)}

- **standardised** - format only (spacing, case, date/ID/decimal notation); the value means the same.
- **corrected** - the value was repaired or converted: scale change, typo, cross-reference with another source.
- **omitted** - a row or value was left out on purpose: duplicates, test entries, unusable text, unattributable swipes.
- **missing** - data the school expected but the source does not contain.
- **flagged** - nothing was changed, but the entry is unusual and is queued for a person (gate-log events).
- **needs_review** - entries (of any action) that a person must confirm; see section 4.

Result: **{len(students_df)} students**, **{STATS['grades_available']} of {STATS['grades_expected']} subject scores** available on one
0-100 scale, **{len(daily_df):,} student-days** of attendance from {n_lines:,} log lines.

## 2. Scales brought to a single percentage

| Source | Scale found | Rule applied |
|---|---|---|
| Gradebook (Math, English) | 0-100, one class typed as 0-1 | fraction x 100 when coursework + exam confirm it |
| Science (Word) | 0-1; Grade 11 table headed "/20"; stray percentages | scale read from each table header; /20 x 5; values 10-100 in a 0-1 column read as percent |
| History (PDF) | GPA 0-4; stray percentages | Percent = 50 + 12.5 x GPA (policy printed on the report) |
| Enrollment form | free text: 88, 88%, 0.88, 3.4, 3.4/4, B+ | <=1 -> x100; 1-4 -> GPA policy (flagged as an assumption); letters flagged |

Values that fit no scale (e.g. 8.5 in a 0-1 column, GPA 4.8) are **not guessed**: the score is withheld and queued for review.

## 3. Free-text fields harmonised in bulk

{md(formats_df)}

## 4. Entries that need a person - {len(review_df)} in total

{md(top_review)}

The full queue, one row per entry with the raw value and the student, is the *Needs Review* sheet
(sorted high -> low priority). {len(flags_df)} of these come from the gate log and are also listed with
date and time on the *Log Flags* sheet.

## 5. Everything by source, action and category

{md(cats)}
"""
(REP / "cleaning_report.md").write_text(rep, encoding="utf-8")

print(summary.to_string())
print("\nactions:", STATS["actions"], "\nreview:", STATS["review"])
print(f"grades available {STATS['grades_available']}/{STATS['grades_expected']};  log: {STATS['log']}")
