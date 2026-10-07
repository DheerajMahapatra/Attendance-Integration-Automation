"""
College Attendance Integrator
=============================

Integrates semester attendance Excel files, with or without a MASTER student
workbook - from the command line or through the Streamlit web app (app.py),
which runs this very same pipeline.

Folder layout
-------------
    attendance_integrator/
        attendance_integrator.py
        app.py                        <- Streamlit web app
        requirements.txt
        test_attendance_integrator.py
        input/
            MASTER.xlsx            <- optional student database (never modified)
            Semester_1.xlsx        <- attendance exports (.xlsx or .xls)
            ...
            Semester_7.xlsx
        output/
            Integrated_Attendance.xlsx

Any other folder of the project that holds attendance exports can be used
instead (for example an uploaded batch folder).

Run
---
    pip install -r requirements.txt
    python attendance_integrator.py                     # newest attendance folder
    python attendance_integrator.py "path/to/folder"    # explicit folder
    streamlit run app.py                                # web app

Matching rules
--------------
* Student  : MASTER 'Enrolment No.'  <-> attendance 'USN NO'  (primary key,
             normalised, case insensitive, never matched by name alone).
             When no MASTER shares enrolment numbers with the attendance data
             the roster is taken from the attendance files themselves - no
             enrolment matching happens at all.
* Semester : attendance 'SEMESTER' column (I..VII, 1..7, 'Semester -01', ...)
             or the file name, matched to the MASTER semester sections.
* Subject  : MASTER subject heading course code <-> 'COURSE_CODE1..N'
             (without a MASTER every course code of the export is used).

Attendance business rules
-------------------------
1. Below 75%      : B is added so the subject lands inside 75 - 78%.
2. 75% and above  : REGULAR students below 95% are pushed into 95 - 98%;
                    LATERAL / EX-REGULAR students and everybody already at
                    95% or above are left untouched (a blank student type is
                    treated as regular).
3. Inside a band  : the number of added classes is spread deterministically
                    over the band (student + semester + course), so the
                    result is reproducible but not identical for everybody.
4. Final %        : (A + B) / Total * 100  and never above 100%.
5. Totals         : ALWAYS weighted - sum(attended) / sum(total) * 100.
                    Subject percentages are never averaged.
6. Empty semester : left empty (never turned into 0%) - and gets no sheet.

MASTER file (optional)
----------------------
A MASTER workbook is only used when its enrolment numbers actually appear in
the attendance data.  Without one (for example a folder that only contains
attendance exports) the student roster is built straight from the attendance
files and the output keeps exactly the same sheets and columns.  Students of
another admission year (the USN carries the year, 'AJU/240916' -> 2024) are
kept out of that roster and listed in 'Other_Batch_Students'; only the
majority year of the export is used, and a MASTER flow never filters.

Output workbook  (output/Integrated_Attendance.xlsx)
----------------------------------------------------
    Semester N                 one sheet per semester that has attendance
                               (Semester 1 .. Semester 7): student wise +
                               subject wise rows and a bold weighted total
                               row for every student
    Summary                    overall Semester 1-7 attendance per student,
                               semester wise weighted totals, report metrics
    Integration_Log            every source record with its disposition
    Unmatched_Students / Unmatched_Subjects
    Other_Batch_Students       students of another admission year (roster mode)
    Integrated Master          the original MASTER sheet with its values filled
"""

from __future__ import annotations

import hashlib
import io
import math
import re
import sys
from copy import copy
from dataclasses import dataclass, field
from html.parser import HTMLParser
from pathlib import Path
from typing import Callable, Dict, List, Optional, Sequence, Tuple

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

# ==============================================================================
# CONFIGURATION
# ==============================================================================
INPUT_DIR = Path("input")
OUTPUT_DIR = Path("output")
OUTPUT_FILE = OUTPUT_DIR / "Integrated_Attendance.xlsx"

MASTER_SHEET_TITLE = "Integrated Master"
LOG_SHEET = "Integration_Log"
UNMATCHED_STUDENTS_SHEET = "Unmatched_Students"
UNMATCHED_SUBJECTS_SHEET = "Unmatched_Subjects"
OTHER_BATCH_SHEET = "Other_Batch_Students"
SUMMARY_SHEET = "Summary"
SEM_SHEET_FMT = "Semester {}"

# College template of the "Semester 1..7" sheets (row 2 = title + subject
# headings, row 3 = column headers, one row per student from row 4 on).
SEM_INFO_COLUMNS = [
    "Sr. No",
    "Student Name",
    "Father's Name",
    "Enrolment No.",
    "Course Name",
    "Batch/Year",
    "Regular /Lateral /Ex Regular",
    "Total No. of Classess",
    "Total No. of Classess Attended",
    "Current Attendance %",
]
SEM_SUBJECT_COLUMNS = [
    "Total No of Allotted Classes",
    "Total No. of Classess",
    "Total No of Classes Attended- Actual (A)",
    "Total No of Classes Added (B)",
    "Total No. of Classess Attended (A+B)",
    "Current Attendance %",
]
SEM_TITLE_COL = 8            # H - "<ordinal> Semester" merged over H:J
SEM_TITLE_SPAN = 3
SEM_FIRST_SUBJECT_COL = 11   # K - first subject section
SEM_SUBJECT_SPAN = 6
SEM_COUNT_HEADERS = {"Total No. of Classess", "Total No. of Classess Attended",
                     "Total No of Allotted Classes",
                     "Total No of Classes Attended- Actual (A)",
                     "Total No of Classes Added (B)",
                     "Total No. of Classess Attended (A+B)"}
SEM_INFO_WIDTHS = [6, 27, 14.4, 10.9, 20.9, 9.7, 11.5, 15.3, 17, 14]
SEM_SUBJECT_WIDTHS = [13.1, 15.3, 15, 14, 14.3, 14]
SEM_ORDINALS = ("1st", "2nd", "3rd", "4th", "5th", "6th", "7th")

# subject titles supplied with the college template (group heading text)
SUBJECT_NAME_HINTS = {
    "BTE21001": "Engineering Mathematics-I",
    "BTE21262": "Programming for Problem Solving",
    "BTE22010": "Engineering Physics",
    "BTE22267": "Workshop Practices",
    "BTE22370": "English for Communication",
    "BTE21261": "Engineering Physics Lab",
    "BTE25095": "Constitution of India",
}

# Student wise overall (Semester 1-7) columns of the Summary sheet.
OVERALL_COLUMNS = [
    "Student Name",
    "Enrolment No.",
    "Student Type",
    "Section",
    "Semesters with Data",
    "Total Classes",
    "Total Attended",
    "Current Attendance %",
    "Classes Required to Reach 75%",
    "Final Attendance %",
    "Final Attended (A+B)",
    "Bonus Classes (95% Rule)",
]

SEMESTERS = range(1, 8)
TARGET_PERCENT = 75.0          # 75% attendance requirement
MAX_PERCENT = 100.0            # attendance can never exceed 100%
EPS = 1e-9

# the two target bands: what an attendance percentage is pushed into
LOW_BAND = (75.0, 78.0)        # below 75%  -> land inside 75 - 78%
HIGH_BAND = (95.0, 98.0)       # 75%+ regular -> land inside 95 - 98%

# report headings follow the rules in force - nothing in the output is fixed
REQUIRED_LABEL_PREFIX = "Classes Required to Reach"
BONUS_LABEL_PREFIX = "Bonus Classes ("


def _fmt_percent(value: float) -> str:
    """75.0 -> '75', 77.5 -> '77.5' (used inside the report headings)."""
    return f"{value:g}"


def required_label() -> str:
    return f"{REQUIRED_LABEL_PREFIX} {_fmt_percent(TARGET_PERCENT)}%"


def bonus_label() -> str:
    return f"{BONUS_LABEL_PREFIX} {_fmt_percent(HIGH_BAND[0])}% Rule)"


def rules_snapshot() -> Dict[str, float]:
    """The attendance rules currently in force (UI + report headings)."""
    return {
        "target_percent": TARGET_PERCENT,
        "low_band_min": LOW_BAND[0],
        "low_band_max": LOW_BAND[1],
        "high_band_min": HIGH_BAND[0],
        "high_band_max": HIGH_BAND[1],
        "max_percent": MAX_PERCENT,
    }


def configure_rules(target_percent: Optional[float] = None,
                    low_band: Optional[Sequence[float]] = None,
                    high_band: Optional[Sequence[float]] = None) -> Dict[str, float]:
    """
    Change the attendance rules while the program is running (the Streamlit
    app exposes this in the sidebar).  Only the given arguments change and an
    impossible combination is refused instead of quietly producing nonsense.
    """
    global TARGET_PERCENT, LOW_BAND, HIGH_BAND
    target = float(target_percent) if target_percent is not None else float(TARGET_PERCENT)
    low = tuple(float(v) for v in low_band) if low_band is not None else tuple(LOW_BAND)
    high = tuple(float(v) for v in high_band) if high_band is not None else tuple(HIGH_BAND)
    if len(low) != 2 or len(high) != 2:
        raise ValueError("each band needs exactly two values: minimum and maximum")
    if not 0 < target <= MAX_PERCENT:
        raise ValueError(f"the target percentage must be inside 0 - {_fmt_percent(MAX_PERCENT)}%")
    if low[0] > low[1] or high[0] > high[1]:
        raise ValueError("a band minimum must not be above its maximum")
    if low[0] < target:
        raise ValueError(f"the lower band must start at/above the target "
                         f"({_fmt_percent(target)}%)")
    if low[1] >= high[0]:
        raise ValueError("the two bands must not overlap")
    if high[1] > MAX_PERCENT:
        raise ValueError(f"a band must not end above {_fmt_percent(MAX_PERCENT)}%")
    TARGET_PERCENT, LOW_BAND, HIGH_BAND = target, low, high
    _refresh_rule_labels()
    return rules_snapshot()


def _refresh_rule_labels() -> None:
    """Keep the column headings of the reports in sync with the rules."""
    for columns in (OVERALL_COLUMNS, LOG_COLUMNS, SEMESTER_SUMMARY_COLUMNS):
        for index, value in enumerate(columns):
            if isinstance(value, str) and value.startswith(REQUIRED_LABEL_PREFIX):
                columns[index] = required_label()
            elif isinstance(value, str) and value.startswith(BONUS_LABEL_PREFIX):
                columns[index] = bonus_label()

ROMAN_TO_INT = {"I": 1, "II": 2, "III": 3, "IV": 4, "V": 5, "VI": 6, "VII": 7}
COURSE_CODE_RE = re.compile(r"\b([A-Z]{2,6}\d{3,8})\b")
ATTENDANCE_USN_HEADERS = ("USN NO", "USN", "ENROLMENT NO", "ENROLLMENT NO", "USN/ENROLMENT NO")


# ==============================================================================
# NORMALISATION HELPERS  (spec 2, 3, 4, 21)
# ==============================================================================
def _text(value) -> str:
    """Convert any cell value to a clean single-spaced string."""
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    s = str(value).replace("\xa0", " ").replace("\u200b", "")
    return re.sub(r"\s+", " ", s).strip()


def normalize_student_id(value) -> str:
    """AJU/230064 stays AJU/230064, padded/spaced variants become identical."""
    return _text(value).upper()


def student_key(value) -> str:
    """Tolerant comparison key: case/space/punctuation insensitive."""
    return re.sub(r"[^A-Z0-9]", "", normalize_student_id(value))


def normalize_course_code(value) -> str:
    """Course codes are compared without spaces and case insensitive."""
    return re.sub(r"\s+", "", _text(value)).upper()


def course_key(value) -> str:
    return re.sub(r"[^A-Z0-9]", "", normalize_course_code(value))


def usn_batch_year(value) -> Optional[int]:
    """
    Admission year hidden in a USN such as 'AJU/240916' -> 2024.
    None when the number does not follow the '<prefix>/<YY><serial>' shape.
    """
    text = normalize_student_id(value)
    if "/" not in text:
        return None
    tail = text.rsplit("/", 1)[-1]
    if len(tail) < 4 or not tail.isdigit():
        return None
    year = int(tail[:2])
    return 2000 + year if 15 <= year <= 40 else None


def detect_batch_year(records: Sequence["CourseRecord"]) -> Optional[int]:
    """
    Admission year the majority of the students of an export belongs to.
    None when the USNs cannot be read or when no year holds a strict majority
    (a folder that really mixes batches is left untouched).
    """
    counts: Dict[int, int] = {}
    seen = set()
    for rec in records:
        key = student_key(rec.usn)
        if key in seen:
            continue
        seen.add(key)
        year = usn_batch_year(rec.usn)
        if year is not None:
            counts[year] = counts.get(year, 0) + 1
    if not counts:
        return None
    batch = sorted(counts, key=lambda y: (-counts[y], -y))[0]
    total = sum(counts.values())
    return batch if counts[batch] * 2 > total else None


def clean_subject_name(label: str, course_code: str = "") -> str:
    """
    Turn a MASTER subject heading into a subject name.
        'Sub: Engineering Chemistry - BTE201' -> 'Engineering Chemistry'
    """
    text = _text(label)
    if not text:
        return ""
    text = re.sub(r"^(SUBJECT|SUB|COURSE)\s*[:\-]\s*", "", text, flags=re.I)
    if course_code:
        text = text.replace(course_code, "")
    text = re.sub(r"\s*[-–—|:]\s*$", "", text.strip())
    return text.strip(" -–—:")


def to_number(value) -> Optional[float]:
    """Return a float for numeric cells, None for blanks / non numeric text."""
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    s = _text(value).replace("%", "").replace(",", "")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def normalize_semester(value) -> Optional[int]:
    """
    Supports: I..VII, 1..7, '1st', 'Semester -01', 'SEM III', 'SEMESTER V'.
    Never mis-reads sessions such as 'ODD 2023-24' as a semester.
    """
    s = _text(value).upper()
    if not s:
        return None

    if s in ROMAN_TO_INT:                      # 'V', 'VII'
        return ROMAN_TO_INT[s]

    if re.fullmatch(r"\d{1,2}", s):            # bare '3'
        n = int(s)
        return n if n in SEMESTERS else None

    if re.fullmatch(r"\d{1,2}(ST|ND|RD|TH)", s):   # '1st', '4th'
        n = int(re.match(r"\d{1,2}", s).group())
        return n if n in SEMESTERS else None

    if "SEM" in s:                             # 'Semester -01', 'SEM III'
        m = re.search(r"(\d{1,2})", s)
        if m and int(m.group(1)) in SEMESTERS:
            return int(m.group(1))
        for token in re.split(r"[^A-Z]+", s):
            if token in ROMAN_TO_INT:
                return ROMAN_TO_INT[token]
    return None


def semester_from_filename(name: str) -> Optional[int]:
    """Fallback when a file has no usable SEMESTER column."""
    up = Path(name).stem.upper()
    m = re.search(r"SEM(?:ESTER)?[\s_\-]*(\d{1,2})", up)
    if m and int(m.group(1)) in SEMESTERS:
        return int(m.group(1))
    m = re.search(r"SEM(?:ESTER)?[\s_\-]*(IV|III|II|VI|VII|V|I)(?![A-Z0-9])", up)
    if m:
        return ROMAN_TO_INT[m.group(1)]
    return None


def is_regular(student_type: str) -> bool:
    """
    Who receives the 95 - 98% rule: everybody except LATERAL / EX-REGULAR.
    A missing or unlabelled student type is treated as the normal (regular) case.
    """
    t = re.sub(r"\s+", " ", normalize_student_id(student_type).replace("-", " ").replace("_", " ")).strip()
    if not t:
        return True                        # no type given -> standard rule
    if "LATERAL" in t:
        return False
    if t.replace(" ", "").startswith("EX"):      # EX-REGULAR / EX REGULAR
        return False
    return "REGULAR" in t


# ==============================================================================
# FILE READING  (xlsx, xls, and ERP exports that are HTML saved as .xls)
# ==============================================================================
class _HtmlTableParser(HTMLParser):
    """Minimal stdlib parser for ERP 'xls' files that are really HTML tables."""

    def __init__(self) -> None:
        super().__init__()
        self.rows: List[List[str]] = []
        self._row: Optional[List[str]] = None
        self._cell: Optional[List[str]] = None

    def handle_starttag(self, tag, attrs):
        if tag == "tr":
            self._row = []
        elif tag in ("td", "th"):
            self._cell = []

    def handle_endtag(self, tag):
        if tag == "tr" and self._row is not None:
            self.rows.append(self._row)
            self._row = None
        elif tag in ("td", "th") and self._cell is not None and self._row is not None:
            self._row.append(_text("".join(self._cell)))
            self._cell = None

    def handle_data(self, data):
        if self._cell is not None:
            self._cell.append(data)


def _read_html_table(path: Path) -> List[List[str]]:
    parser = _HtmlTableParser()
    parser.feed(path.read_text(encoding="utf-8", errors="replace"))
    return parser.rows


def _read_xls_biff(path: Path) -> List[List[str]]:
    """Real binary .xls (needs xlrd)."""
    import xlrd  # optional dependency, only used for binary .xls files

    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_index(0)
    return [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]


def read_raw_sheets(path: Path) -> List[Tuple[str, List[List]]]:
    """
    Return [(sheet_name, rows)] for every readable sheet of a workbook.
    Handles .xlsx (openpyxl), .xls binary (xlrd) and HTML-in-.xls (stdlib).
    """
    suffix = path.suffix.lower()
    if suffix in (".xlsx", ".xlsm"):
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        sheets = []
        try:
            for ws in wb.worksheets:
                rows = [list(r) for r in ws.iter_rows(values_only=True)]
                sheets.append((ws.title, rows))
        finally:
            wb.close()
        return sheets

    if suffix == ".xls":
        head = path.read_bytes()[:512].lstrip()[:1]
        if head in (b"<", b"\xef\xbb\xbf"):          # HTML export pretending to be .xls
            return [(path.name, _read_html_table(path))]
        try:
            return [(path.name, _read_xls_biff(path))]
        except Exception as exc:                     # pragma: no cover - defensive
            print(f"  WARNING: cannot read {path.name}: {exc}")
            return []
    return []


def _pad(rows: List[List], width: int) -> List[List]:
    out = []
    for row in rows:
        row = list(row)
        if len(row) < width:
            row += [None] * (width - len(row))
        out.append(row[:width])
    return out


def find_header_row(rows: Sequence[Sequence], required: Sequence[str],
                    max_rows: int = 15) -> Optional[int]:
    """Row index (0 based) whose upper-cased cells contain every `required` token."""
    for idx in range(min(len(rows), max_rows)):
        joined = " | ".join(_text(v).upper() for v in rows[idx])
        if all(token in joined for token in required):
            return idx
    return None


def _header_index(headers: Sequence[str]) -> Dict[str, int]:
    idx: Dict[str, int] = {}
    for i, h in enumerate(headers):
        idx[_text(h).upper()] = i
    return idx


def _get(row: Dict, *names: str):
    for name in names:
        if name in row:
            v = row[name]
            if not (isinstance(v, float) and math.isnan(v)):
                return v
    return None


# ==============================================================================
# FILE DETECTION  (spec 14)
# ==============================================================================
EXCLUDED_DIR_NAMES = {"output", "__pycache__"}


def _usable_workbook(path: Path) -> bool:
    """A readable Excel file that is not a lock / hidden file."""
    return (path.is_file()
            and path.suffix.lower() in (".xlsx", ".xlsm", ".xls")
            and not path.name.startswith("~$")
            and not path.name.startswith("."))


def _excluded_path(path: Path, root: Path) -> bool:
    """True for hidden, output and cache folders below `root`."""
    try:
        parts = path.relative_to(root).parts[:-1]
    except ValueError:                       # pragma: no cover - defensive
        return False
    return any(part.startswith(".") or part.lower() in EXCLUDED_DIR_NAMES
               for part in parts)


def load_excel_files(directory: Optional[Path] = None) -> List[Path]:
    """All readable workbooks inside the folder (sub folders are included)."""
    directory = directory or INPUT_DIR
    if not directory.exists():
        return []
    files = [p for p in directory.rglob("*")
             if _usable_workbook(p) and not _excluded_path(p, directory)]
    return sorted(files, key=lambda p: p.name.lower())


def _attendance_student_keys(rows: Sequence[Sequence], hdr: int) -> set:
    """Enrolment / USN numbers of one attendance sheet."""
    headers = [_text(v).upper() for v in rows[hdr]]
    usn_col = next((i for i, h in enumerate(headers) if h in ATTENDANCE_USN_HEADERS), None)
    if usn_col is None:
        usn_col = next((i for i, h in enumerate(headers) if h.startswith("USN")), None)
    if usn_col is None:
        return set()
    keys = set()
    for row in rows[hdr + 1:]:
        if usn_col < len(row):
            key = student_key(row[usn_col])
            if key:
                keys.add(key)
    return keys


def _sniff_file(path: Path) -> Tuple[bool, bool, set, Optional[int]]:
    """
    Returns (is_master_like, is_attendance_like, student_keys, semester_hint).
    `student_keys` carries the enrolment numbers of a MASTER and the USN
    numbers of an attendance export, so both can be compared.
    """
    is_master = False
    is_attendance = False
    keys: set = set()
    for _sheet_name, rows in read_raw_sheets(path):
        hdr = find_header_row(rows, ("ENROLMENT", "STUDENT NAME"))
        if hdr is not None:
            is_master = True
            idx = _header_index(rows[hdr])
            col = idx.get("ENROLMENT NO.") or idx.get("ENROLMENT NO") or idx.get("ENROLLMENT NO")
            if col is None:
                for name, i in idx.items():
                    if "ENROLMENT" in name or "ENROLLMENT" in name:
                        col = i
                        break
            if col is not None:
                for row in rows[hdr + 1:]:
                    if col < len(row):
                        k = student_key(row[col])
                        if k:
                            keys.add(k)
        hdr_a = find_header_row(rows, ("USN",))
        if hdr_a is not None:
            headers = [_text(v).upper() for v in rows[hdr_a]]
            if any(re.fullmatch(r"COURSE_CODE\d+", h) for h in headers):
                is_attendance = True
                keys |= _attendance_student_keys(rows, hdr_a)
    sem = semester_from_filename(path.name)
    return is_master, is_attendance, keys, sem


def detect_master_file(paths: Sequence[Path],
                       attendance_keys: Optional[set] = None,
                       sniffed: Optional[Dict[Path, tuple]] = None) -> Optional[Path]:
    """
    Pick the MASTER workbook - but only when it really belongs to the
    attendance data: at least one enrolment number has to appear in both.

    Returns None when nothing matches, so the student roster can be taken
    straight from the attendance files (no enrolment matching at all).
    """
    candidates = []
    for p in paths:
        info = sniffed.get(p) if sniffed else None
        if info is None:
            info = _sniff_file(p)
        is_master, _is_att, keys, _sem = info
        if is_master:
            candidates.append((p, keys))
    if not candidates:
        return None

    attendance_keys = attendance_keys or set()
    ranked = sorted(
        candidates,
        key=lambda item: (
            len(item[1] & attendance_keys),
            1 if item[0].stem.lower() == "master" else 0,
            -len(item[1]),
        ),
        reverse=True,
    )
    best, best_keys = ranked[0]
    if not (best_keys & attendance_keys):
        return None                      # no shared enrolment number -> no MASTER
    return best


def detect_attendance_files(paths: Sequence[Path], master: Path,
                            sniffed: Optional[Dict[Path, tuple]] = None) -> List[Path]:
    """Every workbook that is not the master and has USN + COURSE_CODE columns."""
    out = []
    for p in paths:
        if master is not None and p.resolve() == master.resolve():
            continue
        info = sniffed.get(p) if sniffed else None
        if info is None:
            info = _sniff_file(p)
        _m, is_att, _keys, _sem = info
        if is_att:
            out.append(p)
    return out


def _is_attendance_workbook(path: Path) -> bool:
    try:
        return bool(_sniff_file(path)[1])
    except Exception:                       # pragma: no cover - unreadable file
        return False


def discover_input_folders() -> List[Path]:
    """
    Folders that directly contain attendance workbooks, most recently
    updated first (the input folder wins a tie).
    """
    score: Dict[Path, Tuple[float, int]] = {}
    for root, priority in ((INPUT_DIR, 0), (Path.cwd(), 1)):
        if root is None or not Path(root).exists():
            continue
        root = Path(root).resolve()          # the same folder must not be listed twice
        for path in root.rglob("*"):
            if not _usable_workbook(path) or _excluded_path(path, root):
                continue
            if not _is_attendance_workbook(path):
                continue
            try:
                mtime = path.stat().st_mtime
            except OSError:                # pragma: no cover - defensive
                continue
            best = score.get(path.parent)
            if best is None or (mtime, -priority) > (best[0], -best[1]):
                score[path.parent] = (mtime, priority)
    return sorted(score, key=lambda folder: (-score[folder][0], score[folder][1]))


def resolve_input_dir(argv: Optional[Sequence[str]] = None) -> Path:
    """
    Folder to read: the command line argument when one is given, otherwise the
    folder whose attendance files were updated most recently.
    """
    args = [a for a in list(sys.argv[1:] if argv is None else argv)
            if not a.startswith("-")]
    if args:
        target = Path(args[0]).expanduser()
        if target.is_file():
            target = target.parent
        if target.exists():
            return target
        print(f"\nWARNING: folder not found, ignoring the argument: {args[0]}")

    folders = discover_input_folders()
    if not folders:
        return INPUT_DIR
    if len(folders) > 1:
        print("\nAttendance folders found:")
        for folder in folders:
            print(("  -> " if folder == folders[0] else "     ") + str(folder))
        print("Reading the most recently updated one. Use a folder argument to "
              "change:\n     python attendance_integrator.py \"<folder>\"")
    return folders[0]


# ==============================================================================
# ATTENDANCE FILE READING  (spec 14, 15, 16)
# ==============================================================================
@dataclass
class CourseRecord:
    source_file: str
    source_sheet: str
    row_number: int
    slot: int
    usn: str
    name: str
    student_type: str
    section: str
    branch: str
    session: str
    scheme: str
    semester: Optional[int]
    course_code: str
    course_key: str
    attendance_for: str
    total: Optional[float]
    attended: Optional[float]
    source_percentage: Optional[float]


@dataclass
class AttendanceFileData:
    path: Path
    records: List[CourseRecord] = field(default_factory=list)
    source_totals: Dict[Tuple[str, int], Tuple[Optional[float], Optional[float]]] = field(default_factory=dict)
    semester_hint: Optional[int] = None
    sheets_read: List[str] = field(default_factory=list)


def read_attendance_workbook(path: Path) -> AttendanceFileData:
    """
    Read every sheet of an attendance workbook (spec 15).
    Sheets without USN + COURSE_CODE columns are ignored.
    """
    data = AttendanceFileData(path=path, semester_hint=semester_from_filename(path.name))
    for sheet_name, rows in read_raw_sheets(path):
        hdr = find_header_row(rows, ("USN",))
        if hdr is None:
            continue
        headers = [_text(v).upper() for v in rows[hdr]]
        if not any(re.fullmatch(r"COURSE_CODE\d+", h) for h in headers):
            continue

        body = [r for r in rows[hdr + 1:] if any(_text(v) for v in r)]
        if not body:
            continue

        width = len(headers)
        frame = pd.DataFrame(_pad(body, width), columns=headers)
        data.sheets_read.append(sheet_name)

        sem_hint = data.semester_hint
        idx = _header_index(headers)
        slot_numbers = sorted(
            {int(m.group(1)) for h in headers if (m := re.fullmatch(r"COURSE_CODE(\d+)", h))}
        )

        for pos, row in enumerate(frame.to_dict("records"), start=hdr + 2):
            usn = normalize_student_id(_get(row, "USN NO", "USN", "ENROLMENT NO", "ENROLLMENT NO"))
            if not usn:
                continue
            semester = normalize_semester(_get(row, "SEMESTER", "SEM"))
            if semester is None:
                semester = sem_hint

            common = dict(
                source_file=path.name,
                source_sheet=sheet_name,
                row_number=pos,
                usn=usn,
                name=_text(_get(row, "STUDENT_NAME", "STUDENT NAME")),
                student_type=_text(_get(row, "STUDENT TYPE", "TYPE")),
                section=_text(_get(row, "SECTION")),
                branch=_text(_get(row, "BRANCH")),
                session=_text(_get(row, "SESSION")),
                scheme=_text(_get(row, "SCHEME")),
                semester=semester,
            )

            for n in slot_numbers:
                code = normalize_course_code(_get(row, f"COURSE_CODE{n}"))
                if not code:
                    continue  # blank course code -> skip (spec 16)
                record = CourseRecord(
                    slot=n,
                    course_code=code,
                    course_key=course_key(code),
                    attendance_for=_text(_get(row, f"ATTENDANCE_FOR{n}")),
                    total=to_number(_get(row, f"TOTAL_CLASS{n}")),
                    attended=to_number(_get(row, f"ATTENDED_CLASS{n}")),
                    source_percentage=to_number(_get(row, f"PERCENTAGE{n}")),
                    **common,
                )
                data.records.append(record)

            # optional source provided semester totals (used for verification)
            src_total = to_number(_get(row, "TOTAL_CLASSES"))
            src_attended = to_number(_get(row, "TOTAL_ATTENDED_CLASSES"))
            if src_total is not None or src_attended is not None:
                data.source_totals[(student_key(usn), semester or 0)] = (src_total, src_attended)

    return data


# ==============================================================================
# DUPLICATE HANDLING  (spec 17)
# ==============================================================================
KEEP = "KEEP"
IGNORE = "IGNORE"
CONFLICT_KEPT = "CONFLICT_KEPT"


def resolve_duplicates(records: Sequence[CourseRecord]) -> List[Tuple[CourseRecord, str, str]]:
    """
    Group records by (student, semester, course code).

    * identical duplicates        -> keep one, ignore the rest
    * duplicate with no data      -> keep the one that has data
    * duplicates with other data  -> keep the largest TOTAL_CLASS and flag it
    """
    groups: Dict[str, List[CourseRecord]] = {}
    for rec in records:
        groups.setdefault(f"{student_key(rec.usn)}|{rec.semester}|{rec.course_key}", []).append(rec)

    resolved: List[Tuple[CourseRecord, str, str]] = []
    for _key, group in groups.items():
        if len(group) == 1:
            resolved.append((group[0], KEEP, ""))
            continue
        with_data = [r for r in group if r.total is not None and r.attended is not None]
        if not with_data:
            kept = group[0]
            resolved.append((kept, KEEP, ""))
            for other in group[1:]:
                resolved.append((other, IGNORE, "Duplicate course code without attendance values"))
            continue
        if len(with_data) == 1:
            kept = with_data[0]
            resolved.append((kept, KEEP, "Duplicate course code: kept the record with attendance values"))
            for other in group:
                if other is not kept:
                    resolved.append((other, IGNORE, "Duplicate course code without attendance values"))
            continue

        values = {(r.total, r.attended) for r in with_data}
        kept = max(with_data, key=lambda r: (r.total or 0, r.attended or 0))
        if len(values) == 1:
            resolved.append((kept, KEEP, "Identical duplicate record collapsed to one entry"))
            remark_ignore = "Identical duplicate record"
        else:
            resolved.append((kept, CONFLICT_KEPT,
                             "DUPLICATE CONFLICT: different values found, kept the record "
                             "with the highest TOTAL_CLASS"))
            remark_ignore = "Duplicate record with conflicting values (not used)"
        for other in group:
            if other is not kept:
                resolved.append((other, IGNORE, remark_ignore))
    return resolved


# ==============================================================================
# ATTENDANCE CALCULATIONS  (spec 6, 7, 8, 9, 10, 11)
# ==============================================================================
def _band_label(band: Tuple[float, float]) -> str:
    """(75.0, 78.0) -> '75-78'  (used in the log remarks)."""
    return f"{band[0]:g}-{band[1]:g}"


def spread_seed(usn, semester, course_code) -> str:
    """Stable key of one student + semester + course: drives the band spread."""
    return f"{student_key(usn)}|{semester if semester else ''}|{course_key(course_code)}"


def _stable_fraction(seed: str) -> float:
    """Deterministic 0..1 value for a seed - the same seed always gives the same value."""
    if not seed:
        return 0.5
    digest = hashlib.md5(seed.encode("utf-8")).hexdigest()
    return int(digest[:8], 16) / float(0xFFFFFFFF)


def _band_classes(total: float, actual: float, band: Tuple[float, float],
                  seed: str = "") -> int:
    """
    Classes to add so that (actual + B) / total * 100 lands inside `band`.

    * already at/above the band -> 0 (attendance is never reduced)
    * the band is reachable     -> a deterministic value inside the band, so
                                   students spread over 75-78% / 95-98%
    * the band is unreachable   -> the closest value at/above the band floor
                                   (happens only when a single class is worth
                                   more than the width of the band)
    """
    low, high = band
    current = actual / total * 100.0
    if current >= low - EPS:
        return 0                                        # already good enough

    needed = low / 100.0 * total - actual               # classes to reach the floor
    allowed = high / 100.0 * total - actual             # classes to stay in the band
    b_min = max(0, math.ceil(needed - EPS))
    b_max = math.floor(allowed + EPS)
    room = int(math.floor(total - actual + EPS))        # never above 100%

    if b_max >= b_min:
        span = int(b_max - b_min) + 1
        b = b_min + min(int(_stable_fraction(seed) * span), span - 1)
    else:
        b = b_min                                       # band too narrow -> best effort
    return max(0, min(int(b), room))


def calculate_required_classes(total: float, actual: float, student_type: str,
                               seed: str = "") -> Tuple[int, float, float]:
    """
    Core business rule. Returns (B, A+B, final percentage).

    * below 75%          -> B pushes the subject into the 75 - 78% band
    * at/above 75% and REGULAR below 95%
                         -> B pushes the subject into the 95 - 98% band
    * at/above 95%, or LATERAL / EX-REGULAR -> B = 0 (nothing is ever removed)
    """
    total = float(total)
    if total <= 0:
        return 0, 0.0, 0.0
    actual = max(0.0, min(float(actual), total))   # never negative, never > total

    current = actual / total * 100.0
    if current < TARGET_PERCENT - EPS:
        band = LOW_BAND                            # below 75% -> 75 - 78%
    elif is_regular(student_type):
        band = HIGH_BAND                           # 75 - 94.99% -> 95 - 98%
    else:
        return 0, actual, current                  # LATERAL / EX-REGULAR: untouched

    required = _band_classes(total, actual, band, seed)
    final_attended = actual + required
    pct = min(MAX_PERCENT, final_attended / total * 100.0)
    return int(required), final_attended, pct


def calculate_subject_attendance(total: Optional[float],
                                 actual: Optional[float],
                                 student_type: str,
                                 seed: str = "") -> Optional[dict]:
    """
    Per subject result (spec 6, 7, 8, 9).

    * current      = actual / total * 100                (real, untouched value)
    * required_75  = classes added by the 75 - 78% rule  (0 when already >= 75%)
    * bonus_15     = classes added by the 95 - 98% rule  (0 when below 75%)
    * final        = actual + required_75 + bonus_15     (never > total)
    * percentage   = final / total * 100                 (never > 100)

    `seed` makes the spread inside a band reproducible for one
    student + semester + course.

    None when the source data is unusable - attendance is never invented.
    """
    if total is None or actual is None or total <= 0:
        return None                                   # never invent attendance
    total_f = float(total)
    actual_f = max(0.0, min(float(actual), total_f))   # never negative, never > total
    b, final_attended, pct = calculate_required_classes(total_f, actual_f,
                                                        student_type, seed)

    current = actual_f / total_f * 100.0
    if current < TARGET_PERCENT - EPS:
        required_75, bonus = b, 0                     # 75 - 78% rule
    else:
        required_75, bonus = 0, b                     # 95 - 98% rule

    return {
        "total": total_f,
        "actual": actual_f,
        "current": current,
        "required_75": float(required_75),
        "bonus_15": float(bonus),
        "added": float(b),
        "final": float(final_attended),
        "percentage": pct,
    }


def calculate_semester_total(subject_results: Sequence[dict]) -> Optional[dict]:
    """
    Semester totals are WEIGHTED (spec 10):
        Current % = sum(attended) / sum(total) * 100
        Final   % = sum(final attended) / sum(total) * 100
    Never the average of subject percentages.
    """
    valid = [s for s in subject_results if s and s.get("total")]
    if not valid:
        return None
    total = sum(s["total"] for s in valid)
    actual = sum(s["actual"] for s in valid)
    required = sum(s.get("required_75", 0.0) for s in valid)
    bonus = sum(s.get("bonus_15", 0.0) for s in valid)
    added = sum(s["added"] for s in valid)
    final = sum(s["final"] for s in valid)
    pct = (final / total * 100.0) if total else 0.0
    current = (actual / total * 100.0) if total else 0.0
    return {"total": total, "actual": actual, "current": current,
            "required_75": required, "bonus_15": bonus, "added": added,
            "final": final, "percentage": min(MAX_PERCENT, pct)}


def calculate_overall_attendance(semester_results: Dict[int, dict]) -> Optional[dict]:
    """
    Overall attendance across semesters 1..7 (spec 11) - also weighted.
    Semesters without data are simply not included (never treated as 0%).
    """
    valid = [v for v in semester_results.values() if v and v.get("total")]
    if not valid:
        return None
    total = sum(v["total"] for v in valid)
    actual = sum(v["actual"] for v in valid)
    required = sum(v.get("required_75", 0.0) for v in valid)
    bonus = sum(v.get("bonus_15", 0.0) for v in valid)
    added = sum(v["added"] for v in valid)
    final = sum(v["final"] for v in valid)
    pct = (final / total * 100.0) if total else 0.0
    actual_pct = (actual / total * 100.0) if total else 0.0
    return {"total": total, "actual": actual, "current": actual_pct,
            "required_75": required, "bonus_15": bonus, "added": added,
            "final": final, "percentage": min(MAX_PERCENT, pct),
            "actual_percentage": actual_pct,
            "semesters": sorted(
                sem for sem, v in semester_results.items() if v and v.get("total"))}


# ==============================================================================
# MASTER WORKBOOK LAYOUT
# ==============================================================================
def _classify_field(header: str) -> Optional[str]:
    """Map a MASTER column header onto a semantic role."""
    f = _text(header).upper()
    if not f:
        return None
    if ("REGULAR" in f and "LATERAL" in f) or f in ("STUDENT TYPE", "TYPE") or "EX-REGULAR" in f:
        return "type"
    if "ALLOTTED" in f:
        return "allotted"
    if "ACTUAL" in f:
        return "actual"
    if "ADDED" in f or "ADDITIONAL" in f or "(B)" in f:
        return "added"
    if "A+B" in f:
        return "final"
    if "%" in f or "PERCENT" in f:
        return "pct"
    if "ATTENDED" in f:
        return "final"
    if "TOTAL" in f and "CLASS" in f:
        return "total"
    return None


@dataclass
class Block:
    start_col: int
    end_col: int
    label: str
    kind: str                      # 'semester' | 'subject' | 'spare'
    semester: Optional[int] = None
    course_code: str = ""
    course_key: str = ""
    columns: Dict[str, int] = field(default_factory=dict)


@dataclass
class MasterLayout:
    sheet_title: str
    header_row: int
    group_row: Optional[int]
    data_start: int
    info_cols: Dict[str, int]
    semester_blocks: Dict[int, Block]
    subject_blocks: List[Block]
    subject_index: Dict[Tuple[int, str], Block]
    student_rows: Dict[str, int]
    student_info: Dict[str, dict]
    overall_block: Optional[Block] = None
    reference_block: Optional[Block] = None

    def has_subject_catalog(self, semester: int) -> bool:
        return any(b.semester == semester and b.course_key for b in self.subject_blocks)

    def match_course(self, semester: Optional[int], code_key: str) -> Optional[Block]:
        if semester is None:
            return None
        return self.subject_index.get((semester, code_key))


def find_master_sheet(workbook) -> Optional[object]:
    """First sheet whose header contains 'Enrolment No.' and 'Student Name'."""
    for ws in workbook.worksheets:
        for r in range(1, min(ws.max_row, 20) + 1):
            joined = " | ".join(_text(ws.cell(r, c).value).upper()
                                for c in range(1, ws.max_column + 1))
            if "ENROLMENT" in joined and "STUDENT NAME" in joined:
                return ws
    return None


def _nearest_semester_left(ws, group_row: int, col: int) -> Optional[int]:
    for c in range(col - 1, 0, -1):
        text = _text(ws.cell(group_row, c).value).upper()
        if not text:
            continue
        if "SEM" in text:
            sem = normalize_semester(text)
            if sem:
                return sem
        elif text in ROMAN_TO_INT:
            return ROMAN_TO_INT[text]
    return None


def _collect_blocks(ws, group_row: Optional[int]) -> List[Tuple[int, int]]:
    """Column spans of every labelled group (merged ranges first, values fallback)."""
    if group_row is None:
        return []
    spans = [(rng.min_col, rng.max_col) for rng in ws.merged_cells.ranges
             if rng.min_row == group_row == rng.max_row
             and rng.max_col - rng.min_col + 1 >= 3]
    if not spans:
        starts = [c for c in range(1, ws.max_column + 1)
                  if _text(ws.cell(group_row, c).value)]
        spans = []
        for i, start in enumerate(starts):
            end = (starts[i + 1] - 1) if i + 1 < len(starts) else min(ws.max_column, start + 5)
            spans.append((start, max(start, end)))
    return sorted(spans)


def parse_master_layout(ws) -> MasterLayout:
    """
    Understands both known MASTER shapes:
      * semester sections (merged 'Semester -01' style headers, 5-6 columns each)
      * subject sections (course code in the heading, spec 4/5)
    """
    header_row = None
    for r in range(1, min(ws.max_row, 20) + 1):
        joined = " | ".join(_text(ws.cell(r, c).value).upper()
                            for c in range(1, ws.max_column + 1))
        if "ENROLMENT" in joined and "STUDENT NAME" in joined:
            header_row = r
            break
    if header_row is None:
        raise ValueError(f"No header row with 'Enrolment No.' + 'Student Name' in '{ws.title}'.")

    group_row = header_row - 1 if header_row > 1 else None
    raw_blocks = _collect_blocks(ws, group_row)
    first_block_col = raw_blocks[0][0] if raw_blocks else ws.max_column + 1

    def find_col(*needles: str, limit: Optional[int] = None) -> Optional[int]:
        limit = limit or first_block_col
        for c in range(1, limit):
            text = _text(ws.cell(header_row, c).value).upper()
            if any(n in text for n in needles):
                return c
        return None

    info_cols = {
        "enrollment": find_col("ENROLMENT", "ENROLLMENT"),
        "name": find_col("STUDENT NAME"),
        "type": find_col("REGULAR"),
        "semester": None,
        "father": find_col("FATHER"),
        "course": find_col("COURSE NAME"),
        "batch": find_col("BATCH"),
        "section": None,
    }
    for c in range(1, first_block_col):
        text = _text(ws.cell(header_row, c).value).upper()
        if text == "SEMESTER":
            info_cols["semester"] = c
        if text == "SECTION":
            info_cols["section"] = c
    if info_cols["enrollment"] is None:
        raise ValueError("MASTER 'Enrolment No.' column not found.")
    if info_cols["name"] is None:
        info_cols["name"] = info_cols["enrollment"]

    semester_blocks: Dict[int, Block] = {}
    subject_blocks: List[Block] = []
    spare_blocks: List[Block] = []

    for start, end in raw_blocks:
        label = _text(ws.cell(group_row, start).value) if group_row else ""
        up = label.upper()
        block = Block(start_col=start, end_col=end, label=label, kind="spare")

        if up.startswith("OVERALL"):
            block.kind = "spare"                   # our own summary section
        elif "SEM" in up:
            sem = normalize_semester(label)
            if sem:
                block.kind = "semester"
                block.semester = sem
        elif COURSE_CODE_RE.search(label) or up.startswith("SUB"):
            block.kind = "subject"
            m = COURSE_CODE_RE.search(label)
            if m:
                block.course_code = m.group(1)
                block.course_key = course_key(m.group(1))
            block.semester = _nearest_semester_left(ws, group_row, start) if group_row else None

        for c in range(start, end + 1):
            role = _classify_field(ws.cell(header_row, c).value)
            if role and role not in block.columns:
                block.columns[role] = c

        if block.kind == "semester":
            semester_blocks.setdefault(block.semester, block)
        elif block.kind == "subject":
            subject_blocks.append(block)
        else:
            spare_blocks.append(block)

    subject_index: Dict[Tuple[int, str], Block] = {}
    for block in subject_blocks:
        if block.semester and block.course_key:
            subject_index.setdefault((block.semester, block.course_key), block)

    data_start = header_row + 1
    student_rows: Dict[str, int] = {}
    student_info: Dict[str, dict] = {}
    enroll_col = info_cols["enrollment"]
    for r in range(data_start, ws.max_row + 1):
        key = student_key(ws.cell(r, enroll_col).value)
        if not key:
            continue
        if key in student_rows:
            continue
        student_rows[key] = r
        student_info[key] = {
            "usn": normalize_student_id(ws.cell(r, enroll_col).value),
            "name": _text(ws.cell(r, info_cols["name"]).value),
            "father": (_text(ws.cell(r, info_cols["father"]).value)
                       if info_cols.get("father") else ""),
            "course": (_text(ws.cell(r, info_cols["course"]).value)
                       if info_cols.get("course") else ""),
            "batch": (_text(ws.cell(r, info_cols["batch"]).value)
                      if info_cols.get("batch") else ""),
        }

    layout = MasterLayout(
        sheet_title=ws.title,
        header_row=header_row,
        group_row=group_row,
        data_start=data_start,
        info_cols=info_cols,
        semester_blocks=semester_blocks,
        subject_blocks=subject_blocks,
        subject_index=subject_index,
        student_rows=student_rows,
        student_info=student_info,
    )
    for spare in spare_blocks:
        if spare.label.upper().startswith("OVERALL"):
            layout.overall_block = spare          # already written in a previous run
            break
    layout._spare_blocks = spare_blocks           # noqa: SLF001 - used when adding new blocks
    return layout


# ------------------------------------------------------------------------------
# New blocks (Overall + source info) are appended to the RIGHT of the existing
# sections so nothing existing is moved, deleted or reordered (spec 13).
# ------------------------------------------------------------------------------
def _copy_style(src_cell, dst_cell) -> None:
    try:
        dst_cell._style = copy(src_cell._style)
    except Exception:                       # pragma: no cover - defensive
        pass


def _write_header(ws, layout: MasterLayout, col: int, text: str) -> None:
    cell = ws.cell(layout.header_row, col)
    cell.value = text
    style_src = ws.cell(layout.header_row, max(1, col - 1))
    _copy_style(style_src, cell)


def _new_block(ws, layout: MasterLayout, start: int, end: int,
               label: str, headers: Dict[str, str]) -> Block:
    if layout.group_row:
        if end > start:
            ws.merge_cells(start_row=layout.group_row, start_column=start,
                           end_row=layout.group_row, end_column=end)
        label_cell = ws.cell(layout.group_row, start)
        label_cell.value = label
        _copy_style(ws.cell(layout.group_row, max(1, start - 1)), label_cell)
    block = Block(start_col=start, end_col=end, label=label, kind="spare")
    for offset, (role, text) in enumerate(headers.items()):
        col = start + offset
        _write_header(ws, layout, col, text)
        block.columns[role] = col
        letter = get_column_letter(col)
        if letter not in ws.column_dimensions or not ws.column_dimensions[letter].width:
            ws.column_dimensions[letter].width = 17
    return block


def ensure_overall_block(ws, layout: MasterLayout) -> Block:
    """Create/claim the 6 column 'Overall (Semesters 1-7)' section."""
    if layout.overall_block is not None:          # created by an earlier run
        return layout.overall_block

    headers = {
        "type": "Regular/Lateral/Ex-Regular",
        "total": "Total No. of Classess",
        "actual": "Total No of Actual  Clases Attended (A)",
        "added": "Total No of Additional Classes(B)",
        "final": "Total No. of Classess Attended(A+B)",
        "pct": "Overall Attendance %",
    }
    last_sem_end = max((b.end_col for b in layout.semester_blocks.values()), default=0)
    for spare in getattr(layout, "_spare_blocks", []):
        if spare.start_col > last_sem_end and (spare.end_col - spare.start_col + 1) >= 6:
            if layout.group_row and not spare.label:
                ws.cell(layout.group_row, spare.start_col).value = "Overall (Semesters 1-7)"
                spare.label = "Overall (Semesters 1-7)"
            spare.columns = {}
            for offset, (role, text) in enumerate(headers.items()):
                col = spare.start_col + offset
                _write_header(ws, layout, col, text)
                spare.columns[role] = col
            layout.overall_block = spare
            return spare

    start = last_sem_end + 2
    end = start + len(headers) - 1
    block = _new_block(ws, layout, start, end, "Overall (Semesters 1-7)", headers)
    layout.overall_block = block
    return block


def ensure_reference_block(ws, layout: MasterLayout, overall: Block) -> Block:
    """Section filled from the attendance export: SECTION / BRANCH / SESSION ..."""
    headers = {
        "section": "SECTION",
        "branch": "BRANCH",
        "session": "SESSION",
        "scheme": "SCHEME",
        "type": "STUDENT TYPE",
        "semesters": "SEMESTERS WITH DATA",
    }
    start = overall.end_col + 1
    end = start + len(headers) - 1
    block = _new_block(ws, layout, start, end, "Attendance Source Info", headers)
    layout.reference_block = block
    return block


# ==============================================================================
# MATCHING  (spec 2, 4)
# ==============================================================================
def match_student(layout: MasterLayout, usn: str) -> Optional[int]:
    """MASTER row for an enrolment/USN number, or None (never name matching)."""
    return layout.student_rows.get(student_key(usn))


def match_course(layout: MasterLayout, semester: Optional[int], code_key: str) -> Optional[Block]:
    """MASTER subject block for a course code inside a semester section."""
    return layout.match_course(semester, code_key)


# ==============================================================================
# INTEGRATION
# ==============================================================================
LOG_COLUMNS = [
    "Source File", "Source Sheet", "USN", "Student Name", "Semester", "Course Code",
    "Total Classes", "Actual Attended", "Current Attendance %",
    "Classes Required to Reach 75%", "Bonus Classes (95% Rule)", "Added Classes B",
    "Final Attended A+B", "Final Percentage", "Student Type", "Match Status", "Remarks",
]


@dataclass
class IntegrationResult:
    workbook: object
    layout: MasterLayout
    master_sheet_title: str
    original_sheet_title: str
    master_generated: bool = False      # True when the roster had to be built
    log_rows: List[list] = field(default_factory=list)
    unmatched_students: List[list] = field(default_factory=list)
    unmatched_subjects: List[list] = field(default_factory=list)
    other_batch_students: List[list] = field(default_factory=list)
    stats: Dict[str, int] = field(default_factory=dict)
    semester_stats: Dict[int, dict] = field(default_factory=dict)
    student_overalls: Dict[str, dict] = field(default_factory=dict)
    # student_data[key][semester] = meta + subject rows + weighted semester total
    student_data: Dict[str, Dict[int, dict]] = field(default_factory=dict)
    # (semester, course key) -> clean subject name read from the MASTER heading
    subject_names: Dict[Tuple[int, str], str] = field(default_factory=dict)
    # semester -> every course code seen in the source, in source order
    semester_subjects: Dict[int, List[str]] = field(default_factory=dict)
    student_meta: Dict[str, dict] = field(default_factory=dict)
    source_total_checks: List[bool] = field(default_factory=list)


def _set(cell, value) -> None:
    """Write a clean value (int counts, 2 decimal percentages)."""
    if value is None:
        cell.value = None
        return
    if isinstance(value, float):
        if abs(value - round(value)) < 1e-9:
            value = int(round(value))
        else:
            value = round(value, 2)
    cell.value = value


def _write_result(ws, row: int, columns: Dict[str, int], res: dict) -> None:
    if "allotted" in columns:
        # source exports only provide one class count -> it fills both fields
        _set(ws.cell(row, columns["allotted"]), res["total"])
    if "total" in columns:
        _set(ws.cell(row, columns["total"]), res["total"])
    if "actual" in columns:
        _set(ws.cell(row, columns["actual"]), res["actual"])
    if "added" in columns:
        _set(ws.cell(row, columns["added"]), res["added"])
    if "final" in columns:
        _set(ws.cell(row, columns["final"]), res["final"])
    if "pct" in columns:
        _set(ws.cell(row, columns["pct"]), res["percentage"])


# ==============================================================================
# ROSTER WITHOUT A MASTER  -  the student list comes from the attendance exports
# ==============================================================================
ROSTER_INFO_HEADERS = [
    "Sr. No", "Student Name", "Father's Name", "Enrolment No.",
    "Course Name", "Batch/Year", "Regular /Lateral /Ex Regular",
]
ROSTER_INFO_WIDTHS = [6, 30, 14, 14, 34, 12, 15]


def collect_student_meta(records: Sequence[CourseRecord]) -> Dict[str, dict]:
    """One metadata entry per student; the latest semester wins for each detail."""
    meta: Dict[str, dict] = {}
    for rec in records:
        key = student_key(rec.usn)
        entry = meta.setdefault(key, {
            "usn": rec.usn, "name": rec.name, "type": "", "section": "",
            "branch": "", "session": "", "scheme": "", "semesters": set(),
        })
        entry["name"] = entry["name"] or rec.name
        if rec.semester:
            entry["semesters"].add(rec.semester)
        for field_name, value in (("type", rec.student_type), ("section", rec.section),
                                  ("branch", rec.branch), ("session", rec.session),
                                  ("scheme", rec.scheme)):
            if value:
                entry[field_name] = value      # latest semester wins
    return meta


def _roster_course_name(entry: dict) -> str:
    return _text(entry.get("branch")) or _text(entry.get("scheme"))


def _roster_batch(entry: dict) -> str:
    sources = f"{_text(entry.get('scheme'))} {_text(entry.get('session'))}"
    m = re.search(r"(20\d{2}\s*-\s*\d{2,4})", sources)
    if m:
        return m.group(1).replace(" ", "")
    m = re.search(r"\b(20\d{2})\b", sources)
    return m.group(1) if m else ""


def build_roster_workbook(records: Sequence[CourseRecord]):
    """
    MASTER substitute used when no workbook shares enrolment numbers with the
    attendance data: a sheet shaped exactly like a MASTER (info columns + one
    merged block per semester) that parse_master_layout understands, built from
    the attendance exports only.
    """
    meta = collect_student_meta(records)
    students = [meta[key] for key in sorted(meta)]
    semesters = sorted({rec.semester for rec in records if rec.semester})

    workbook = openpyxl.Workbook()
    ws = workbook.active
    ws.title = "MASTER"

    # row 1 = merged group labels, row 2 = column headers, row 3.. = students
    for col, text in enumerate(ROSTER_INFO_HEADERS, start=1):
        cell = ws.cell(2, col, text)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")

    col = len(ROSTER_INFO_HEADERS) + 1
    for semester in semesters:
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 5)
        label = ws.cell(1, col, f"Semester -{semester:02d}")
        label.font = Font(bold=True)
        label.alignment = Alignment(horizontal="center")
        for offset, text in enumerate(SEM_SUBJECT_COLUMNS):
            cell = ws.cell(2, col + offset, text)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
        col += 6
    ws.row_dimensions[2].height = 45

    for index, entry in enumerate(students, start=1):
        row = 2 + index
        ws.cell(row, 1, index)
        ws.cell(row, 2, _text(entry.get("name")))
        ws.cell(row, 4, _text(entry.get("usn")))
        ws.cell(row, 5, _roster_course_name(entry))
        ws.cell(row, 6, _roster_batch(entry))
        ws.cell(row, 7, _text(entry.get("type")))

    for index, width in enumerate(ROSTER_INFO_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    ws.freeze_panes = "A3"
    return workbook, ws


def split_off_other_batch(
        records: Sequence[CourseRecord],
) -> Tuple[List[CourseRecord], List[list]]:
    """
    Roster mode: keep only the students of the batch the export belongs to.
    A student whose USN carries another admission year (an older batch that is
    still listed in the file) is taken out of the roster and reported one line
    per student in the 'Other_Batch_Students' sheet - nothing is written into
    the roster or the semester sheets for him.
    """
    batch_year = detect_batch_year(records)
    if batch_year is None:
        return list(records), []
    kept: List[CourseRecord] = []
    dropped: List[CourseRecord] = []
    for rec in records:
        (kept if usn_batch_year(rec.usn) in (None, batch_year) else dropped).append(rec)
    if not dropped:
        return kept, []
    return kept, other_batch_rows(dropped, batch_year)


def other_batch_rows(records: Sequence[CourseRecord],
                     batch_year: int) -> List[list]:
    """One line per excluded student - the export rows are never lost."""
    grouped: Dict[str, dict] = {}
    for rec in records:
        entry = grouped.setdefault(student_key(rec.usn), {
            "usn": rec.usn, "name": rec.name, "years": set(),
            "semesters": set(), "records": 0, "files": set(),
        })
        entry["records"] += 1
        entry["files"].add(rec.source_file)
        if rec.semester:
            entry["semesters"].add(rec.semester)
        year = usn_batch_year(rec.usn)
        if year is not None:
            entry["years"].add(year)
    rows: List[list] = []
    for entry in sorted(grouped.values(), key=lambda e: e["usn"]):
        year_list = sorted(entry["years"])
        years = ", ".join(str(y) for y in year_list) or "unreadable"
        batch_cell = year_list[0] if len(year_list) == 1 else years
        semesters = ", ".join(f"Semester {s}" for s in sorted(entry["semesters"])) or "-"
        rows.append([
            entry["usn"], entry["name"], batch_cell, batch_year, semesters,
            entry["records"], ", ".join(sorted(entry["files"])),
            f"USN belongs to the {years} batch while the export holds the "
            f"{batch_year} batch",
        ])
    return rows


OTHER_BATCH_HEADERS = [
    "USN", "Student Name", "USN batch year", "Export batch year",
    "Semesters in the export", "Attendance records", "Source files", "Reason",
]


def integrate_attendance(master_path: Optional[Path],
                         attendance_data: Sequence[AttendanceFileData]) -> IntegrationResult:
    """
    Read the MASTER (never modified on disk) - or build the roster straight from
    the attendance files when there is no usable MASTER - write every matching
    attendance value into it and collect logs / statistics.
    """
    # ------------------------------------------------------------------ index
    records: List[CourseRecord] = []
    for data in attendance_data:
        records.extend(data.records)
    source_totals = {}
    for data in attendance_data:
        source_totals.update(data.source_totals)

    other_batch_rows: List[list] = []

    if master_path is None:
        records, other_batch_rows = split_off_other_batch(records)
        workbook, ws = build_roster_workbook(records)
    else:
        workbook = openpyxl.load_workbook(master_path)
        ws = find_master_sheet(workbook)
        if ws is None:
            raise ValueError("No MASTER sheet with 'Enrolment No.' + 'Student Name' was found.")

    original_title = ws.title
    layout = parse_master_layout(ws)
    overall_block = ensure_overall_block(ws, layout)
    reference_block = ensure_reference_block(ws, layout, overall_block)

    # subject names are read from the MASTER subject headings (when present)
    subject_names: Dict[Tuple[int, str], str] = {}
    for block in layout.subject_blocks:
        if block.semester and block.course_key:
            subject_names.setdefault(
                (block.semester, block.course_key),
                clean_subject_name(block.label, block.course_code))

    grouped: Dict[Tuple[str, Optional[int]], List[CourseRecord]] = {}
    for rec in records:
        grouped.setdefault((student_key(rec.usn), rec.semester), []).append(rec)

    # every course code of a semester in the order the source files list them
    semester_subjects: Dict[int, List[str]] = {}
    seen_codes: Dict[int, set] = {}
    for rec in records:
        if not rec.semester or not rec.course_code:
            continue
        seen = seen_codes.setdefault(rec.semester, set())
        lookup = rec.course_code.strip().upper()
        if lookup in seen:
            continue
        seen.add(lookup)
        semester_subjects.setdefault(rec.semester, []).append(rec.course_code)

    stats = {
        "records_read": len(records),
        "subjects_processed": 0,
        "students_matched": 0,
        "students_unmatched": 0,
        "master_students_without_data": 0,
        "duplicates_ignored": 0,
        "duplicate_conflicts": 0,
        "missing_attendance": 0,
        "subject_mismatches": 0,
        "records_integrated": 0,
        "semesters_processed": 0,
        "source_total_mismatches": 0,
        "students_other_batch": len(other_batch_rows),
    }

    log_rows: List[list] = []
    unmatched_students: List[list] = []
    unmatched_subjects: List[list] = []
    # student_data[key][semester] = {"meta": ..., "subjects": [...], "total": ...}
    student_data: Dict[str, Dict[int, dict]] = {}

    def log(rec: CourseRecord, res: Optional[dict], status: str, remark: str) -> None:
        log_rows.append([
            rec.source_file, rec.source_sheet, rec.usn, rec.name,
            rec.semester if rec.semester is not None else "",
            rec.course_code,
            res["total"] if res else (rec.total if rec.total is not None else ""),
            res["actual"] if res else (rec.attended if rec.attended is not None else ""),
            round(res["current"], 2) if res else "",
            res["required_75"] if res else "",
            res["bonus_15"] if res else "",
            res["added"] if res else "",
            res["final"] if res else "",
            round(res["percentage"], 2) if res else "",
            rec.student_type, status, remark,
        ])

    # ------------------------------------------------- student level metadata
    student_meta: Dict[str, dict] = collect_student_meta(records)

    matched_keys = set()

    # ------------------------------------------------------- process by group
    for (key, semester), group in sorted(
            grouped.items(),
            key=lambda kv: (kv[0][1] or 0, kv[0][0])):
        first = group[0]
        row = match_student(layout, first.usn)

        if row is None:                                        # spec 18
            for rec in group:
                reason = ("Student enrolment number not present in MASTER"
                          if key not in layout.student_rows else "Student row unreadable")
                log(rec, None, "STUDENT_NOT_FOUND", reason)
                unmatched_students.append([
                    rec.usn, rec.name, rec.semester if rec.semester is not None else "",
                    rec.course_code, reason, rec.source_file,
                ])
            continue

        if semester is None:                                   # spec 16
            for rec in group:
                log(rec, None, "SEMESTER_NOT_FOUND",
                    "Semester could not be determined from the SEMESTER column or file name")
            continue

        matched_keys.add(key)
        semester_block = layout.semester_blocks.get(semester)
        meta = student_meta.get(key, {})
        subject_results: List[dict] = []
        subject_rows: List[dict] = []          # rows of the Semester 1..7 sheets

        for rec, disposition, dup_remark in resolve_duplicates(group):
            if disposition == IGNORE:
                stats["duplicates_ignored"] += 1
                log(rec, None, "DUPLICATE_IGNORED", dup_remark)
                continue

            remark = dup_remark
            if disposition == CONFLICT_KEPT:
                stats["duplicate_conflicts"] += 1

            if rec.total is None or rec.attended is None:      # spec 16
                stats["missing_attendance"] += 1
                detail = ("TOTAL_CLASS missing in source" if rec.total is None
                          else "ATTENDED_CLASS missing in source")
                log(rec, None, "MISSING_TOTAL",
                    (remark + "; " if remark else "") + detail + " - no value invented")
                continue

            if rec.attended > rec.total:
                remark = (remark + "; " if remark else "") + \
                    "Source ATTENDED_CLASS greater than TOTAL_CLASS - clamped to 100%"

            seed = spread_seed(rec.usn, semester, rec.course_code)
            res = calculate_subject_attendance(rec.total, rec.attended,
                                               meta.get("type") or rec.student_type,
                                               seed)
            if res is None:
                stats["missing_attendance"] += 1
                log(rec, None, "MISSING_TOTAL", "TOTAL_CLASS is zero or invalid")
                continue

            band = LOW_BAND if res["current"] < TARGET_PERCENT - EPS else HIGH_BAND
            label = _band_label(band)
            if res["added"] > 0:
                remark = (remark + "; " if remark else "") + \
                    f"B={int(res['added'])} classes added to reach {label}%"
            elif not is_regular(meta.get("type") or rec.student_type) \
                    and res["current"] >= TARGET_PERCENT - EPS:
                remark = (remark + "; " if remark else "") + \
                    "LATERAL / EX-REGULAR student - no classes added"
            elif band[0] - EPS <= res["current"] <= band[1] + EPS:
                remark = (remark + "; " if remark else "") + \
                    f"Already inside {label}% - no change"
            elif res["current"] > band[1] + EPS:
                remark = (remark + "; " if remark else "") + \
                    f"Attendance already above {label}% - no change"
            else:
                remark = (remark + "; " if remark else "") + \
                    f"No class left to add - attendance stays at " \
                    f"{round(res['current'], 2)}%"

            subject_block = match_course(layout, semester, rec.course_key)
            catalog = layout.has_subject_catalog(semester)

            if subject_block is not None:
                _write_result(ws, row, subject_block.columns, res)
                status = "MATCHED"
                remark = "Written into subject block; " + remark
                stats["records_integrated"] += 1
            elif catalog:                                      # spec 19
                stats["subject_mismatches"] += 1
                status = "COURSE_NOT_IN_MASTER"
                reason = "Course code not found in the MASTER semester section"
                unmatched_subjects.append([
                    rec.usn, semester, rec.course_code, rec.name, reason, rec.source_file,
                ])
                remark = reason + "; still counted in semester totals; " + remark
            else:
                status = "MATCHED"
                reason = ("Roster taken from the attendance files (no MASTER used)"
                          if master_path is None else
                          "No subject section for this course in MASTER")
                remark = reason + " - counted in semester totals; " + remark
                stats["records_integrated"] += 1

            log(rec, res, status, remark)
            subject_results.append(res)
            subject_rows.append({
                "code": rec.course_code,
                "name": subject_names.get((semester, rec.course_key), ""),
                "attendance_for": rec.attendance_for,
                "res": res,
                "status": status,
            })
            stats["subjects_processed"] += 1

        semester_total = calculate_semester_total(subject_results)
        if semester_total and semester_block:
            _write_result(ws, row, semester_block.columns, semester_total)

        # keep per (student, semester) totals for the overall section
        if semester_total:
            store = getattr(layout, "_semester_totals", None)
            if store is None:
                store = {}
                layout._semester_totals = store      # noqa: SLF001
            store.setdefault(key, {})[semester] = semester_total

            info = layout.student_info.get(key, {})
            student_data.setdefault(key, {})[semester] = {
                "meta": {
                    "usn": info.get("usn") or first.usn,
                    "name": info.get("name") or meta.get("name") or first.name,
                    "type": meta.get("type", "") or first.student_type,
                    "section": meta.get("section", "") or first.section,
                },
                "subjects": subject_rows,
                "total": semester_total,
            }

            # verify against the totals provided by the ERP export
            src = source_totals.get((key, semester))
            if src and src[0] is not None:
                stats["source_total_mismatches"] += int(
                    abs(src[0] - semester_total["total"]) > 0.51
                    or (src[1] is not None and abs(src[1] - semester_total["actual"]) > 0.51))

        # ------------------------------------------------- student level fills
        if semester_block and meta.get("type"):
            type_col = semester_block.columns.get("type")
            if type_col:
                cell = ws.cell(row, type_col)
                if not _text(cell.value):
                    cell.value = meta["type"]

        sem_col = layout.info_cols.get("semester")
        if sem_col and meta.get("semesters"):
            cell = ws.cell(row, sem_col)
            if not _text(cell.value):
                cell.value = max(meta["semesters"])

    # ----------------------------------------------------------- overall block
    semester_totals: Dict[str, Dict[int, dict]] = getattr(layout, "_semester_totals", {})
    stats["students_matched"] = len(matched_keys)
    stats["master_students_without_data"] = len(layout.student_rows) - len(matched_keys)
    stats["students_unmatched"] = len({student_key(r.usn) for r in records
                                       if match_student(layout, r.usn) is None})
    stats["semesters_processed"] = len({sem for per_student in semester_totals.values()
                                        for sem in per_student})

    for key, row in layout.student_rows.items():
        per_sem = semester_totals.get(key, {})
        overall = calculate_overall_attendance(per_sem)
        meta = student_meta.get(key, {})
        if overall and layout.overall_block:
            _write_result(ws, row, layout.overall_block.columns, overall)
            student_overalls = getattr(layout, "_student_overalls", None)
            if student_overalls is None:
                student_overalls = {}
                layout._student_overalls = student_overalls   # noqa: SLF001
            student_overalls[key] = overall

        if layout.reference_block:
            cols = layout.reference_block.columns
            values = {
                "section": meta.get("section", ""),
                "branch": meta.get("branch", ""),
                "session": meta.get("session", ""),
                "scheme": meta.get("scheme", ""),
                "type": meta.get("type", ""),
                "semesters": ",".join(str(s) for s in overall["semesters"]) if overall else "",
            }
            for role, col in cols.items():
                if values.get(role):
                    ws.cell(row, col).value = values[role]

    # ------------------------------------------------- semester level summary
    semester_stats: Dict[int, dict] = {}
    for per_sem in semester_totals.values():
        for sem, res in per_sem.items():
            bucket = semester_stats.setdefault(
                sem, {"students": 0, "total": 0.0, "actual": 0.0, "required_75": 0.0,
                      "bonus_15": 0.0, "added": 0.0, "final": 0.0})
            bucket["students"] += 1
            bucket["total"] += res["total"]
            bucket["actual"] += res["actual"]
            bucket["required_75"] += res.get("required_75", 0.0)
            bucket["bonus_15"] += res.get("bonus_15", 0.0)
            bucket["added"] += res["added"]
            bucket["final"] += res["final"]
    for bucket in semester_stats.values():
        bucket["percentage"] = (bucket["final"] / bucket["total"] * 100.0
                                if bucket["total"] else 0.0)
        bucket["current"] = (bucket["actual"] / bucket["total"] * 100.0
                             if bucket["total"] else 0.0)

    return IntegrationResult(
        workbook=workbook,
        layout=layout,
        master_sheet_title=MASTER_SHEET_TITLE,
        original_sheet_title=original_title,
        master_generated=master_path is None,
        log_rows=log_rows,
        unmatched_students=unmatched_students,
        unmatched_subjects=unmatched_subjects,
        other_batch_students=other_batch_rows,
        stats=stats,
        semester_stats=semester_stats,
        student_overalls=getattr(layout, "_student_overalls", {}),
        student_data=student_data,
        subject_names=subject_names,
        semester_subjects=semester_subjects,
        student_meta=student_meta,
    )


# ==============================================================================
# LOG SHEETS  (spec 18, 19, 20)
# ==============================================================================
def _drop_sheet(workbook, title: str) -> None:
    if title in workbook.sheetnames:
        del workbook[title]


def _write_table(workbook, title: str, headers: Sequence[str], rows: Sequence[Sequence]):
    _drop_sheet(workbook, title)
    ws = workbook.create_sheet(title)
    ws.append(list(headers))
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in rows:
        ws.append(list(row))
    for i, _ in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(i)].width = 20
    ws.freeze_panes = "A2"
    return ws


def create_logs(workbook, result: IntegrationResult) -> None:
    _write_table(workbook, LOG_SHEET, LOG_COLUMNS, result.log_rows)
    _write_table(workbook, UNMATCHED_STUDENTS_SHEET,
                 ["USN", "Student Name", "Semester", "Course Code", "Reason", "Source File"],
                 result.unmatched_students)
    _write_table(workbook, UNMATCHED_SUBJECTS_SHEET,
                 ["USN", "Semester", "Course Code", "Student Name", "Reason", "Source File"],
                 result.unmatched_subjects)
    _write_table(workbook, OTHER_BATCH_SHEET, OTHER_BATCH_HEADERS,
                 result.other_batch_students)


# ==============================================================================
# SEMESTER 1..7 SHEETS  -  the college template, one row per student  (spec: output)
# ==============================================================================


def _num(value):
    """Clean numeric value: integers stay integers, percentages keep 2 decimals."""
    if value is None:
        return None
    if isinstance(value, float):
        if abs(value - round(value)) < 1e-9:
            return int(round(value))
        return round(value, 2)
    return value


def _pct(numerator: float, denominator: float) -> Optional[float]:
    """Weighted percentage = attended / total * 100 (None when no data)."""
    if not denominator:
        return None
    return round(numerator / denominator * 100.0, 2)


def _semester_subjects(result: IntegrationResult, semester: int) -> List[dict]:
    """
    Subject sections of one semester, in the order the source files list them.
    Every course code of the export gets a section - a subject whose source
    rows carry no TOTAL_CLASS simply stays blank instead of disappearing.
    """
    catalog = result.semester_subjects.get(semester)
    if not catalog:
        catalog = []
        seen = set()
        master_order = result.layout.student_rows
        for key in sorted(result.student_data, key=lambda k: master_order.get(k, 1 << 30)):
            info = result.student_data[key].get(semester)
            if not info:
                continue
            for subj in info["subjects"]:
                code = _text(subj.get("code")).strip().upper()
                if code and code not in seen:
                    seen.add(code)
                    catalog.append(_text(subj.get("code")))
    ordered: List[dict] = []
    for code in catalog:
        lookup = _text(code).strip().upper()
        if not lookup:
            continue
        ordered.append({
            "code": _text(code),
            "key": lookup,
            "name": result.subject_names.get((semester, lookup))
                    or SUBJECT_NAME_HINTS.get(lookup, ""),
        })
    return ordered


def semesters_with_data(result: IntegrationResult) -> List[int]:
    """
    Semesters that actually carry attendance, in ascending order.
    One input file for one semester -> exactly one semester sheet is produced;
    semesters without data get no sheet at all (never an empty placeholder).
    """
    return sorted(sem for sem, bucket in result.semester_stats.items()
                  if bucket and _semester_subjects(result, sem))


def create_semester_sheets(workbook, result: IntegrationResult) -> None:
    """
    One worksheet per semester that has attendance (Semester 1 .. Semester 7),
    built exactly like the college template:

        row 2 -> "<ordinal> Semester"  +  "Sub: <title> <code>" group headings
        row 3 -> column headers
        row 4..-> one row per MASTER student, in MASTER order (Sr. No 1..N)

    Student level : Total / Attended (A+B) / Current %  = weighted sums of the
                    subject columns of the same row.
    Subject level : Allotted, Total, A, B, A+B, % (A+B / Total).

    Nothing is invented - students and subjects without attendance stay blank,
    so a student without data never shows 0%.  Semesters without any attendance
    get no sheet at all: one file with semester I data yields one sheet.
    """
    layout = result.layout
    master_ws = (workbook[layout.sheet_title]
                 if layout.sheet_title in workbook.sheetnames else None)
    master_type_col = layout.info_cols.get("type")

    for sem in SEMESTERS:                             # drop stale/placeholder sheets
        _drop_sheet(workbook, SEM_SHEET_FMT.format(sem))
    for sem in semesters_with_data(result):
        title = SEM_SHEET_FMT.format(sem)
        ws = workbook.create_sheet(title)
        subjects = _semester_subjects(result, sem)

        # ------------------------------------------------- row 2: title+groups
        ws.cell(2, SEM_TITLE_COL).value = f"{SEM_ORDINALS[sem - 1]} Semester"
        ws.merge_cells(start_row=2, start_column=SEM_TITLE_COL,
                       end_row=2, end_column=SEM_TITLE_COL + SEM_TITLE_SPAN - 1)
        ws.cell(2, SEM_TITLE_COL).font = Font(bold=True, size=12)
        ws.cell(2, SEM_TITLE_COL).alignment = Alignment(horizontal="center")
        for i, subject in enumerate(subjects):
            start = SEM_FIRST_SUBJECT_COL + i * SEM_SUBJECT_SPAN
            ws.merge_cells(start_row=2, start_column=start, end_row=2,
                           end_column=start + SEM_SUBJECT_SPAN - 1)
            heading = (f"Sub: {subject['name']} {subject['code']}"
                       if subject["name"] else f"Sub: {subject['code']}")
            ws.cell(2, start).value = " ".join(heading.split())
            ws.cell(2, start).font = Font(bold=True)
            ws.cell(2, start).alignment = Alignment(horizontal="center")

        # ------------------------------------------------- row 3: column headers
        headers = list(SEM_INFO_COLUMNS)
        for _ in subjects:
            headers.extend(SEM_SUBJECT_COLUMNS)
        for i, text in enumerate(headers, start=1):
            cell = ws.cell(3, i, text)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(wrap_text=True, vertical="center",
                                       horizontal="center")
        ws.row_dimensions[3].height = 63

        # ------------------------------------------------------- student rows
        row_number = 4
        for index, (key, master_row) in enumerate(
                sorted(layout.student_rows.items(), key=lambda kv: kv[1]), start=1):
            info = layout.student_info.get(key, {})
            sem_data = result.student_data.get(key, {}).get(sem)
            meta = (sem_data or {}).get("meta") or result.student_meta.get(key) or {}
            student_type = _text(meta.get("type"))
            if not student_type and master_ws is not None and master_type_col:
                student_type = _text(master_ws.cell(master_row, master_type_col).value)

            ws.cell(row_number, 1).value = index
            ws.cell(row_number, 2).value = info.get("name") or _text(meta.get("name"))
            ws.cell(row_number, 3).value = info.get("father") or ""
            ws.cell(row_number, 4).value = info.get("usn") or ""
            ws.cell(row_number, 5).value = info.get("course") or ""
            ws.cell(row_number, 6).value = info.get("batch") or ""
            ws.cell(row_number, 7).value = student_type

            if sem_data:
                results = {_text(s["code"]).strip().upper(): s["res"]
                           for s in sem_data["subjects"]}
                total = sem_data["total"]
                ws.cell(row_number, SEM_TITLE_COL).value = _num(total["total"])
                ws.cell(row_number, SEM_TITLE_COL + 1).value = _num(total["final"])
                ws.cell(row_number, SEM_TITLE_COL + 2).value = _pct(total["final"],
                                                                   total["total"])
                for i, subject in enumerate(subjects):
                    res = results.get(subject["key"])
                    if not res:
                        continue
                    base = SEM_FIRST_SUBJECT_COL + i * SEM_SUBJECT_SPAN
                    ws.cell(row_number, base).value = _num(res["total"])   # allotted
                    ws.cell(row_number, base + 1).value = _num(res["total"])
                    ws.cell(row_number, base + 2).value = _num(res["actual"])
                    ws.cell(row_number, base + 3).value = _num(res["added"])
                    ws.cell(row_number, base + 4).value = _num(res["final"])
                    ws.cell(row_number, base + 5).value = _pct(res["final"], res["total"])
            row_number += 1

        # ------------------------------------------------------ presentation
        for i, width in enumerate(SEM_INFO_WIDTHS, start=1):
            ws.column_dimensions[get_column_letter(i)].width = width
        for i in range(len(subjects)):
            for offset, width in enumerate(SEM_SUBJECT_WIDTHS):
                column = SEM_FIRST_SUBJECT_COL + i * SEM_SUBJECT_SPAN + offset
                ws.column_dimensions[get_column_letter(column)].width = width
        for row in ws.iter_rows(min_row=4, max_row=row_number - 1, max_col=len(headers)):
            for cell in row:
                header = headers[cell.column - 1]
                if header.endswith("%"):
                    cell.number_format = "0.00"
                elif header in SEM_COUNT_HEADERS:
                    cell.number_format = "0"
        ws.freeze_panes = "E4"


# ==============================================================================
# SUMMARY  (spec 22)
# ==============================================================================
SUMMARY_SECTIONS = ("SUMMARY", "STUDENT-WISE OVERALL ATTENDANCE (SEMESTER 1 TO 7)",
                    "SEMESTER-WISE WEIGHTED ATTENDANCE")
SEMESTER_SUMMARY_COLUMNS = ["Semester", "Students", "Total Classes", "Total Attended",
                            "Current Attendance %", "Classes Required to Reach 75%",
                            "Bonus Classes (95% Rule)", "Final Attended", "Final Attendance %"]


def _overall_meta(result: IntegrationResult, key: str) -> dict:
    meta = result.student_meta.get(key) or {}
    info = result.layout.student_info.get(key) or {}
    return {
        "usn": meta.get("usn") or info.get("usn", ""),
        "name": info.get("name") or meta.get("name", ""),
        "type": meta.get("type", ""),
        "section": meta.get("section", ""),
    }


def create_summary(workbook, result: IntegrationResult, master_name: str,
                   attendance_files: Sequence[Path]) -> None:
    stats = result.stats
    overalls = result.student_overalls
    # 75% is judged on the ACTUAL attendance (before the extra classes B),
    # because the 75% rule itself tops every subject up to 75%.
    below = [o for o in overalls.values() if o.get("actual_percentage", o["percentage"]) < TARGET_PERCENT - EPS]
    at_above = [o for o in overalls.values() if o.get("actual_percentage", o["percentage"]) >= TARGET_PERCENT - EPS]
    below_after = [o for o in overalls.values() if o["percentage"] < TARGET_PERCENT - EPS]
    weighted_total = sum(o["total"] for o in overalls.values())
    weighted_final = sum(o["final"] for o in overalls.values())
    weighted_actual = sum(o["actual"] for o in overalls.values())
    weighted_pct = (weighted_final / weighted_total * 100.0) if weighted_total else 0.0
    weighted_actual_pct = (weighted_actual / weighted_total * 100.0) if weighted_total else 0.0

    subject_note = (
        "(no MASTER) every course code of the export is used - subject wise "
        "attendance is available in the Semester 1..7 sheets and row by row in "
        + LOG_SHEET
        if result.master_generated else
        ("Master has no subject/course-code columns - subject wise "
         "attendance is available in the Semester 1..7 sheets and row "
         "by row in " + LOG_SHEET
         if not result.layout.subject_blocks else
         "Course codes matched against the MASTER subject sections"))

    student_label = ("Total students (roster from the attendance files)"
                     if result.master_generated else "Total students in MASTER")

    rows: List[list] = [
        ["SUMMARY"],
        ["Report generated", __import__("datetime").datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        ["Master file", master_name],
        ["Master sheet exported as", f"{result.master_sheet_title} (was '{result.original_sheet_title}')"],
        ["Attendance files processed", ", ".join(p.name for p in attendance_files)],
        ["Attendance bands",
         f"below {_fmt_percent(TARGET_PERCENT)}% -> {_band_label(LOW_BAND)}%, "
         f"{_fmt_percent(TARGET_PERCENT)}% and above (regular, below "
         f"{_fmt_percent(HIGH_BAND[0])}%) -> {_band_label(HIGH_BAND)}%"],
        ["", ""],
        [student_label, len(result.layout.student_rows)],
        ["Students matched", stats["students_matched"]],
        ["Students unmatched (attendance not in MASTER)", stats["students_unmatched"]],
        ["Students excluded (other batch, listed in " + OTHER_BATCH_SHEET + ")",
         stats.get("students_other_batch", 0)],
        ["MASTER students without attendance data", stats["master_students_without_data"]],
        ["Total attendance records processed", stats["records_read"]],
        ["Total subjects processed (with valid attendance)", stats["subjects_processed"]],
        ["Records integrated (written into the sheet)", stats["records_integrated"]],
        ["Total semesters processed", stats["semesters_processed"]],
        ["Duplicate records found (ignored)", stats["duplicates_ignored"]],
        ["Duplicate conflicts (flagged)", stats["duplicate_conflicts"]],
        ["Records missing TOTAL/ATTENDED class", stats["missing_attendance"]],
        ["Subject mismatches", stats["subject_mismatches"]],
        ["Subject matching note", subject_note],
        ["Source semester totals that disagree with our sums",
         stats["source_total_mismatches"]],
        [f"Students below {_fmt_percent(TARGET_PERCENT)}% (actual attendance)", len(below)],
        [f"Students at/above {_fmt_percent(TARGET_PERCENT)}% (actual attendance)", len(at_above)],
        [f"Students still below {_fmt_percent(TARGET_PERCENT)}% after required classes",
         len(below_after)],
        ["Overall weighted attendance % (actual, before extra classes)",
         round(weighted_actual_pct, 2)],
        ["Overall weighted attendance % (after the "
         f"{_band_label(LOW_BAND)}% / {_band_label(HIGH_BAND)}% rules)",
         round(weighted_pct, 2)],
        ["Weighted total classes", int(weighted_total)],
        ["Weighted actual attended", int(weighted_actual)],
        ["Weighted final attended (A+B)", int(weighted_final)],
        ["Attendance formula", "Total Attended / Total Classes x 100 (weighted, never averaged)"],
        ["Validation", "PASS" if not getattr(result, "validation_errors", []) else "CHECK LOG"],
        ["", ""],
        ["STUDENT-WISE OVERALL ATTENDANCE (SEMESTER 1 TO 7)"],
        list(OVERALL_COLUMNS),
    ]

    # one row per MASTER student - students without data stay blank (never 0%)
    for key, _row in sorted(result.layout.student_rows.items(), key=lambda kv: kv[1]):
        meta = _overall_meta(result, key)
        overall = overalls.get(key)
        head = [meta["name"], meta["usn"], meta["type"], meta["section"]]
        if not overall:
            rows.append(head + ["", "", "", "", "", "", "", ""])
            continue
        rows.append(head + [
            ",".join(str(s) for s in overall["semesters"]),
            _num(overall["total"]), _num(overall["actual"]),
            _pct(overall["actual"], overall["total"]),
            _num(overall.get("required_75", 0.0)),
            _pct(overall["final"], overall["total"]),
            _num(overall["final"]), _num(overall.get("bonus_15", 0.0)),
        ])

    rows.append(["", ""])
    rows.append(["SEMESTER-WISE WEIGHTED ATTENDANCE"])
    rows.append(list(SEMESTER_SUMMARY_COLUMNS))
    for sem in semesters_with_data(result):
        bucket = result.semester_stats[sem]
        rows.append([f"Semester {sem}", bucket["students"], int(bucket["total"]),
                     int(bucket["actual"]), round(bucket["current"], 2),
                     int(bucket["required_75"]), int(bucket["bonus_15"]),
                     int(bucket["final"]), round(bucket["percentage"], 2)])
    rows.append(["ALL SEMESTERS", sum(b["students"] for b in result.semester_stats.values()),
                 int(weighted_total), int(weighted_actual),
                 round(weighted_actual_pct, 2),
                 int(sum(o.get("required_75", 0.0) for o in overalls.values())),
                 int(sum(o.get("bonus_15", 0.0) for o in overalls.values())),
                 int(weighted_final), round(weighted_pct, 2)])

    _drop_sheet(workbook, SUMMARY_SHEET)
    ws = workbook.create_sheet(SUMMARY_SHEET)
    for row in rows:
        ws.append(row)

    for row_cells in ws.iter_rows(min_col=1, max_col=1):
        label = _text(row_cells[0].value)
        if label in SUMMARY_SECTIONS:
            for cell in ws[row_cells[0].row]:
                cell.font = Font(bold=True)
        elif label == "Report generated":
            row_cells[0].font = Font(bold=True)

    ws.column_dimensions["A"].width = 50
    for i in range(2, len(OVERALL_COLUMNS) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 21
    ws.column_dimensions["E"].width = 20
    ws.freeze_panes = "A2"


def patch_summary_validation(errors: Sequence[str],
                             path: Optional[Path] = None) -> None:
    """Write the final validation verdict into the saved Summary sheet."""
    target = Path(path) if path is not None else OUTPUT_FILE
    if not target.exists():
        return
    wb = openpyxl.load_workbook(target)
    if SUMMARY_SHEET not in wb.sheetnames:
        wb.close()
        return
    ws = wb[SUMMARY_SHEET]
    verdict = "PASS - all checks succeeded" if not errors else \
        f"FAIL - {len(errors)} issue(s): " + "; ".join(list(errors)[:3])
    for row in ws.iter_rows(min_col=1, max_col=1):
        if _text(row[0].value) == "Validation":
            ws.cell(row=row[0].row, column=2).value = verdict
            break
    wb.save(target)


# ==============================================================================
# OUTPUT + VALIDATION  (spec 21, 22)
# ==============================================================================
def _is_empty_sheet(ws) -> bool:
    """A sheet with no values and no merged ranges carries no data to preserve."""
    if ws.merged_cells.ranges:
        return False
    return all(cell.value in (None, "") for row in ws.iter_rows() for cell in row)


def save_output(workbook, path: Optional[Path] = None) -> Path:
    """Write the workbook (default: output/Integrated_Attendance.xlsx)."""
    target = Path(path) if path is not None else OUTPUT_FILE
    target.parent.mkdir(parents=True, exist_ok=True)
    reserved = {LOG_SHEET, UNMATCHED_STUDENTS_SHEET, UNMATCHED_SUBJECTS_SHEET,
                OTHER_BATCH_SHEET, SUMMARY_SHEET} \
        | {SEM_SHEET_FMT.format(s) for s in SEMESTERS}
    # blank placeholder sheets of the MASTER (e.g. an empty 'Sem7') are dropped
    for ws in list(workbook.worksheets):
        if ws.title not in reserved and ws.title != MASTER_SHEET_TITLE \
                and _is_empty_sheet(ws):
            workbook.remove(ws)
    # the main MASTER data sheet gets the required output name
    for ws in workbook.worksheets:
        if ws.title not in reserved and ws.title != MASTER_SHEET_TITLE:
            ws.title = MASTER_SHEET_TITLE
            break
    order = [SEM_SHEET_FMT.format(s) for s in SEMESTERS] + [
        SUMMARY_SHEET, LOG_SHEET, UNMATCHED_STUDENTS_SHEET,
        UNMATCHED_SUBJECTS_SHEET, OTHER_BATCH_SHEET, MASTER_SHEET_TITLE]
    workbook._sheets.sort(key=lambda s: order.index(s.title) if s.title in order else 99)
    workbook.save(target)
    return target


def semester_only_bytes(payload: bytes) -> bytes:
    """
    The downloadable copy of a finished workbook: ONLY the semester sheets
    (Semester 1 .. Semester 7).  Summary / Integration_Log / Unmatched_* /
    Integrated Master stay in the web app - the file a teacher shares is the
    college-template sheet itself.  Without any semester sheet the payload is
    returned unchanged (never an empty workbook).
    """
    titles = {SEM_SHEET_FMT.format(sem) for sem in SEMESTERS}
    workbook = openpyxl.load_workbook(io.BytesIO(payload))
    try:
        if not any(ws.title in titles for ws in workbook.worksheets):
            return payload
        for ws in list(workbook.worksheets):
            if ws.title not in titles:
                workbook.remove(ws)
        buffer = io.BytesIO()
        workbook.save(buffer)
        return buffer.getvalue()
    finally:
        workbook.close()


def band_problem(total: float, actual: float, added: float,
                 student_type: str) -> Optional[str]:
    """
    None when `added` respects the attendance bands, otherwise the reason:
      below 75%  -> must end inside the 75 - 78% band
      75%+ regular -> must end inside the 95 - 98% band (or stay untouched)
      LATERAL / EX-REGULAR at 75%+ -> must not be changed at all
    """
    if total is None or total <= 0:
        return f"invalid Total No. of Classess {total}"
    current = actual / total * 100.0
    final = actual + added
    pct = final / total * 100.0
    if pct > MAX_PERCENT + 0.01:
        return f"{round(pct, 2)}% above 100"
    if pct + 0.01 < current:
        return f"attendance reduced from {round(current, 2)}% to {round(pct, 2)}%"
    if current < TARGET_PERCENT - EPS:
        low, high = LOW_BAND
    elif is_regular(student_type):
        low, high = HIGH_BAND
    else:
        if added > 0:
            return (f"B={int(added)} added although {student_type} "
                    f"attendance is already at/above {TARGET_PERCENT:g}%")
        return None                                   # untouched on purpose
    if added > 0 and current >= low - EPS:
        return (f"B={int(added)} added although {round(current, 2)}% "
                f"is already at/above {low:g}%")
    if pct < low - 0.01:
        return f"{round(pct, 2)}% is below the {low:g}% floor"
    if added > 0 and pct > high + 0.01 \
            and (final - 1) / total * 100.0 >= low - 0.01:
        return f"{round(pct, 2)}% is above the {high:g}% ceiling"
    return None


def validate_semester_sheets(wb, semesters: Optional[Sequence[int]] = None) -> List[str]:
    """
    Check the semester sheets against the college template and the rules.

    `semesters` names the sheets that are expected to exist (only the
    semesters the run actually had data for).  Without it every Semester
    sheet that is present is checked.

      * row 2 = "<ordinal> Semester" + one "Sub: ..." heading per subject
      * row 3 = the 10 template columns + one block of 6 columns per subject
      * subject %  = (A+B) / Total * 100, never above 100
      * B is recomputed from the business rules (75-78% / 95-98% bands), or at
        least checked against those bands when the heading carries no course code
      * student Total / Attended / % = weighted sum of that student's subjects
      * a student without data carries no numbers at all (never a fake 0%)
    """
    errors: List[str] = []
    subject_width = len(SEM_SUBJECT_COLUMNS)
    present = [sem for sem in SEMESTERS
               if SEM_SHEET_FMT.format(sem) in wb.sheetnames]
    if semesters is None:
        expected = present
    else:
        expected = sorted({int(s) for s in semesters})
        for sem in expected:
            title = SEM_SHEET_FMT.format(sem)
            if title not in wb.sheetnames:
                errors.append(f"missing sheet: {title}")
    for sem in sorted(set(present) | set(expected)):
        title = SEM_SHEET_FMT.format(sem)
        if title not in wb.sheetnames:
            continue
        ws = wb[title]
        header = [_text(c.value) for c in ws[3]]
        if header[:len(SEM_INFO_COLUMNS)] != SEM_INFO_COLUMNS:
            errors.append(f"{title}: row 3 does not carry the template columns")
            continue
        rest = header[len(SEM_INFO_COLUMNS):]
        if len(rest) % subject_width:
            errors.append(f"{title}: subject columns are not groups of "
                          f"{subject_width}")
            continue
        subjects = len(rest) // subject_width
        for i in range(subjects):
            if rest[i * subject_width:(i + 1) * subject_width] != SEM_SUBJECT_COLUMNS:
                errors.append(f"{title}: subject block {i + 1} does not carry the "
                              f"template columns")
                break
        if not _text(ws.cell(2, SEM_TITLE_COL).value).endswith("Semester"):
            errors.append(f"{title}: row 2 must carry the '<ordinal> Semester' title")
        codes: List[Optional[str]] = []
        for i in range(subjects):
            heading = _text(ws.cell(2, SEM_FIRST_SUBJECT_COL + i * SEM_SUBJECT_SPAN).value)
            if not heading.startswith("Sub:"):
                errors.append(f"{title}: subject heading {i + 1} is '{heading}'")
            match = COURSE_CODE_RE.search(heading.upper())
            codes.append(match.group(1) if match else None)

        for row in ws.iter_rows(min_row=4, max_col=len(header), values_only=True):
            usn = normalize_student_id(row[3]) if row[3] else ""
            if not usn:
                continue
            where = f"{title} {usn}"
            student_type = _text(row[6])
            student_total = to_number(row[SEM_TITLE_COL - 1])
            student_final = to_number(row[SEM_TITLE_COL])
            student_pct = to_number(row[SEM_TITLE_COL + 1])
            sums_total = sums_final = 0.0
            filled = 0

            for i in range(subjects):
                base = SEM_FIRST_SUBJECT_COL - 1 + i * SEM_SUBJECT_SPAN
                allotted = to_number(row[base])
                total = to_number(row[base + 1])
                actual = to_number(row[base + 2])
                added = to_number(row[base + 3])
                final_att = to_number(row[base + 4])
                pct = to_number(row[base + 5])
                if total is None and final_att is None and pct is None:
                    continue                                   # blank stays blank
                filled += 1
                label = f"{where} subject {i + 1}"
                if total is None or total <= 0:
                    errors.append(f"{label}: invalid Total No. of Classess {total}")
                    continue
                if allotted is not None and allotted != total:
                    errors.append(f"{label}: Allotted {allotted} != Total {total}")
                if actual is None or actual < 0 or actual > total + EPS:
                    errors.append(f"{label}: Actual (A) {actual} outside 0..{total}")
                    continue
                if added is None or added < 0:
                    errors.append(f"{label}: Added (B) {added} must not be negative")
                    continue
                if final_att is None or abs(final_att - (actual + added)) > 0.01:
                    errors.append(f"{label}: A+B {final_att} != A + B "
                                  f"{round(actual + added, 2)}")
                    continue
                if final_att > total + EPS:
                    errors.append(f"{label}: A+B {final_att} above Total {total}")
                    continue
                code = codes[i] if i < len(codes) else None
                if code:
                    # the course code is known -> the exact B can be reproduced
                    seed = spread_seed(usn, sem, code)
                    expected = calculate_subject_attendance(total, actual,
                                                            student_type, seed)
                    if expected is not None:
                        if abs(added - expected["added"]) > 0.01:
                            errors.append(f"{label}: B {added} != {expected['added']}")
                        if abs(final_att - expected["final"]) > 0.01:
                            errors.append(f"{label}: A+B {final_att} != "
                                          f"{expected['final']}")
                else:
                    # no course code in the heading -> check the bands only
                    problem = band_problem(total, actual, added, student_type)
                    if problem:
                        errors.append(f"{label}: {problem}")
                expected_pct = _pct(final_att, total)
                if pct is None or expected_pct is None or abs(pct - expected_pct) > 0.02:
                    errors.append(f"{label}: % {pct} != {expected_pct}")
                elif pct > MAX_PERCENT + EPS:
                    errors.append(f"{label}: % {pct} above 100")
                sums_total += total
                sums_final += final_att

            if filled:
                if student_total is None or abs(student_total - sums_total) > 0.01:
                    errors.append(f"{where}: Total No. of Classess {student_total} != "
                                  f"{round(sums_total, 2)} (weighted subject sum)")
                if student_final is None or abs(student_final - sums_final) > 0.01:
                    errors.append(f"{where}: Total No. of Classess Attended "
                                  f"{student_final} != {round(sums_final, 2)}")
                expected_pct = _pct(sums_final, sums_total)
                if student_pct is None or expected_pct is None \
                        or abs(student_pct - expected_pct) > 0.02:
                    errors.append(f"{where}: Current Attendance % {student_pct} != "
                                  f"{expected_pct}")
            elif student_total is not None or student_final is not None \
                    or student_pct is not None:
                errors.append(f"{where}: semester totals without subject data")
    return errors


def _semesters_listed_in_summary(wb) -> Optional[List[int]]:
    """
    The semester numbers the Summary sheet reports data for - the workbook
    must contain a sheet for each of them.  None when there is no Summary
    sheet (the caller reports that separately).
    """
    if SUMMARY_SHEET not in wb.sheetnames:
        return None
    ws = wb[SUMMARY_SHEET]
    in_section = False
    semesters: List[int] = []
    for row in ws.iter_rows(min_col=1, max_col=3, values_only=True):
        label = _text(row[0])
        if label == "SEMESTER-WISE WEIGHTED ATTENDANCE":
            in_section = True
            continue
        if not in_section:
            continue
        if label == "ALL SEMESTERS":
            break
        match = re.fullmatch(r"Semester\s+(\d+)", label, flags=re.IGNORECASE)
        if match and row[1] not in (None, "", 0):      # Students column
            semesters.append(int(match.group(1)))
    return semesters


def validate_output_file(path: Path) -> List[str]:
    """
    Re-open the generated workbook and verify the hard requirements:
      1. 0 <= percentage <= 100        5. B never negative
      2. final attended <= total       6. weighted totals, not averages
      3. percentages match A+B/Total   7. empty semesters stay empty (no sheet)
      4. required sheets exist         8. semester sheets are self consistent
    """
    errors: List[str] = []
    wb = openpyxl.load_workbook(path, data_only=True)
    for title in [SUMMARY_SHEET, LOG_SHEET, UNMATCHED_STUDENTS_SHEET,
                  UNMATCHED_SUBJECTS_SHEET, MASTER_SHEET_TITLE]:
        if title not in wb.sheetnames:
            errors.append(f"missing sheet: {title}")

    errors.extend(validate_semester_sheets(wb,
                                           semesters=_semesters_listed_in_summary(wb)))

    ws = wb[MASTER_SHEET_TITLE] if MASTER_SHEET_TITLE in wb.sheetnames else None
    if ws is None:
        wb.close()
        return errors

    layout = parse_master_layout(ws)
    blocks = list(layout.semester_blocks.values())
    if layout.overall_block:
        blocks.append(layout.overall_block)

    for key, row in layout.student_rows.items():
        for block in blocks:
            cols = block.columns
            total = to_number(ws.cell(row, cols["total"]).value) if "total" in cols else None
            final = to_number(ws.cell(row, cols["final"]).value) if "final" in cols else None
            added = to_number(ws.cell(row, cols["added"]).value) if "added" in cols else None
            pct = to_number(ws.cell(row, cols["pct"]).value) if "pct" in cols else None
            if total is None and final is None and pct is None:
                continue                                   # empty semester stays empty
            if added is not None and added < 0:
                errors.append(f"{key}: negative B in block {block.label}")
            if total is not None and final is not None and final > total + EPS:
                errors.append(f"{key}: final attended > total in block {block.label}")
            if pct is not None and not (0 - EPS <= pct <= MAX_PERCENT + EPS):
                errors.append(f"{key}: percentage out of range in block {block.label}")
            if total and final is not None and pct is not None:
                expected = round(final / total * 100.0, 2)
                if abs(expected - pct) > 0.02:
                    errors.append(f"{key}: {block.label} percentage {pct} != {expected}")

    # every log row must be self consistent
    if LOG_SHEET in wb.sheetnames:
        log = wb[LOG_SHEET]
        header = [_text(c.value) for c in log[1]]
        col = {name: i for i, name in enumerate(header)}
        for r in log.iter_rows(min_row=2, values_only=True):
            total = to_number(r[col["Total Classes"]])
            final = to_number(r[col["Final Attended A+B"]])
            added = to_number(r[col["Added Classes B"]])
            pct = to_number(r[col["Final Percentage"]])
            if total and final is not None:
                if final > total + EPS:
                    errors.append(f"log row {r[:6]}: final > total")
                if added is not None and added < 0:
                    errors.append(f"log row {r[:6]}: negative B")
                if pct is not None and not (0 <= pct <= MAX_PERCENT + EPS):
                    errors.append(f"log row {r[:6]}: percentage out of range")
    wb.close()
    return errors


# ==============================================================================
# MAIN
# ==============================================================================
def _print_banner() -> None:
    print("=" * 36)
    print("COLLEGE ATTENDANCE INTEGRATOR")
    print("=" * 36)


# ==============================================================================
# PIPELINE  -  one code path for the command line, the tests and the web app
# ==============================================================================
@dataclass
class PipelineReport:
    """Everything a run produced: the workbook, the numbers and any errors."""
    input_dir: Path
    master_path: Optional[Path]
    attendance_files: List[Path]
    semester_files: Dict[int, List[str]]
    result: "IntegrationResult"
    output_path: Path
    errors: List[str] = field(default_factory=list)

    @property
    def rules(self) -> Dict[str, float]:
        return rules_snapshot()

    @property
    def ok(self) -> bool:
        return not self.errors


def run_pipeline(input_dir: Path,
                 output_path: Optional[Path] = None,
                 log: Callable[[str], None] = print) -> PipelineReport:
    """
    Read a folder of attendance exports (plus an optional MASTER), apply the
    attendance rules and write the output workbook.

    `log` receives every progress line - print() for the command line, a
    collector for the web app.  A folder that cannot be used raises
    FileNotFoundError / ValueError instead of writing half an output.
    """
    input_dir = Path(input_dir)
    if not input_dir.exists():
        raise FileNotFoundError(f"input folder not found: {input_dir}")

    files = load_excel_files(input_dir)
    if not files:
        raise FileNotFoundError(
            f"no Excel files found in: {input_dir.resolve()}\n"
            "put the attendance exports (and a MASTER file, when you have one) "
            "inside that folder")

    sniffed = {p: _sniff_file(p) for p in files}
    attendance_keys: set = set()
    for p, (_m, is_att, keys, _s) in sniffed.items():
        if is_att:
            attendance_keys |= keys

    master_path = detect_master_file(files, attendance_keys, sniffed)
    attendance_paths = detect_attendance_files(files, master_path, sniffed)
    if not attendance_paths:
        raise ValueError("no attendance files found (need USN NO + COURSE_CODE columns)")

    log(f"\nInput folder         : {input_dir.resolve()}")
    if master_path is not None:
        log(f"Master file          : {master_path.name}")
    else:
        log("Master file          : none - student roster built from the "
            "attendance files (no enrolment matching)")
    log(f"Attendance files found: {len(attendance_paths)}")
    for p in attendance_paths:
        log(f"  - {p.name}")

    attendance_data: List[AttendanceFileData] = []
    semester_files: Dict[int, List[str]] = {}
    for path in attendance_paths:
        data = read_attendance_workbook(path)
        if not data.records:
            log(f"  Skipping (no attendance rows): {path.name}")
            continue
        attendance_data.append(data)
        sems = sorted({r.semester for r in data.records if r.semester})
        for sem in sems:
            semester_files.setdefault(sem, []).append(path.name)
        label = ", ".join(f"Semester {s}" for s in sems) or "semester not detected"
        log(f"  Loaded: {path.name} -> {label} | records: {len(data.records)} | "
            f"sheets: {', '.join(data.sheets_read)}")

    if not attendance_data:
        raise ValueError("no usable attendance data was read")

    for sem in sorted(semester_files):
        log(f"\nProcessing Semester {sem}... ({len(semester_files[sem])} file(s))")

    result = integrate_attendance(master_path, attendance_data)

    create_semester_sheets(result.workbook, result)
    create_logs(result.workbook, result)
    create_summary(result.workbook, result,
                   master_path.name if master_path is not None
                   else "(built from the attendance files)",
                   attendance_paths)

    # ------------------------------------------------------ business validation
    errors: List[str] = []
    for key, overall in result.student_overalls.items():
        if not (0 <= overall["percentage"] <= MAX_PERCENT + EPS):
            errors.append(f"{key}: overall attendance out of range")
        if overall["final"] > overall["total"] + EPS:
            errors.append(f"{key}: overall final attended exceeds total classes")
        if any(s < 0 for s in (overall["added"],)):
            errors.append(f"{key}: negative added classes")

    written = save_output(result.workbook, output_path)
    errors.extend(validate_output_file(written))
    result.validation_errors = errors
    patch_summary_validation(errors, written)

    return PipelineReport(input_dir=input_dir, master_path=master_path,
                          attendance_files=attendance_paths,
                          semester_files=semester_files, result=result,
                          output_path=written, errors=errors)


def print_report(report: PipelineReport) -> None:
    """The command line summary of a run."""
    result = report.result
    stats = result.stats
    overalls = result.student_overalls
    target = _fmt_percent(TARGET_PERCENT)
    weighted_total = sum(o["total"] for o in overalls.values())
    weighted_final = sum(o["final"] for o in overalls.values())
    weighted_pct = (weighted_final / weighted_total * 100.0) if weighted_total else 0.0

    print("\n" + "=" * 36)
    print("RESULTS")
    print("=" * 36)
    print(f"{('Students in roster (no MASTER)' if result.master_generated else 'Students in MASTER'):<29}: {len(result.layout.student_rows)}")
    print(f"Students matched              : {stats['students_matched']}")
    print(f"Students unmatched            : {stats['students_unmatched']}")
    if stats.get("students_other_batch"):
        print(f"Students excluded (other batch): {stats['students_other_batch']} "
              f"-> see the {OTHER_BATCH_SHEET} sheet")
    print(f"Subjects processed            : {stats['subjects_processed']}")
    print(f"Records processed             : {stats['records_read']}")
    print(f"Records written to the sheet  : {stats['records_integrated']}")
    print(f"Semesters processed           : {stats['semesters_processed']}")
    print(f"Duplicate records ignored     : {stats['duplicates_ignored']}")
    print(f"Duplicate conflicts flagged   : {stats['duplicate_conflicts']}")
    print(f"Records missing attendance    : {stats['missing_attendance']}")
    print(f"Subject mismatches            : {stats['subject_mismatches']}")
    below = sum(1 for o in overalls.values()
                if o.get("actual_percentage", o["percentage"]) < TARGET_PERCENT - EPS)
    weighted_actual = sum(o["actual"] for o in overalls.values())
    actual_pct_str = (f"{weighted_actual / weighted_total * 100.0:.2f}"
                      if weighted_total else "0.00")
    print(f"Students below {target}% (actual)  : {below}")
    print(f"Students at/above {target}% (actual): {len(overalls) - below}")
    print(f"Overall weighted attendance % : {weighted_pct:.2f} "
          f"(actual {actual_pct_str}% before the required extra classes)")
    print(f"Validation                    : {'PASS' if not report.errors else 'FAIL'}")
    for err in report.errors[:10]:
        print(f"  ! {err}")

    print("\nOutput:")
    print(f"  {report.output_path.as_posix()}")
    print("  Sheets: " + ", ".join(
        [SEM_SHEET_FMT.format(s) for s in semesters_with_data(result)] +
        [SUMMARY_SHEET, LOG_SHEET, UNMATCHED_STUDENTS_SHEET,
         UNMATCHED_SUBJECTS_SHEET, OTHER_BATCH_SHEET, MASTER_SHEET_TITLE]))
    print("\nCompleted successfully." if not report.errors
          else "\nCompleted with validation errors.")
    print("=" * 36)


def build_report(result: IntegrationResult) -> Dict[str, object]:
    """
    Everything a user interface needs to show a run: the rules in force, the
    detected subjects, one line per student and the weighted totals.
    Every number comes from the run itself - nothing is assumed.
    """
    stats = result.stats
    overalls = result.student_overalls

    subjects: List[dict] = []
    for sem in sorted(result.semester_subjects):
        for code in result.semester_subjects[sem]:
            key = course_key(code)
            subjects.append({
                "semester": sem,
                "code": code,
                "name": result.subject_names.get((sem, key))
                        or SUBJECT_NAME_HINTS.get(key, ""),
            })

    students: List[dict] = []
    for position, (key, _row) in enumerate(
            sorted(result.layout.student_rows.items(), key=lambda kv: kv[1]),
            start=1):
        meta = _overall_meta(result, key)
        overall = overalls.get(key)
        if overall:
            actual_pct = round(overall["actual"] / overall["total"] * 100.0, 2) \
                if overall["total"] else None
            students.append({
                "sr_no": position,
                "usn": meta["usn"],
                "name": meta["name"],
                "type": meta["type"],
                "section": meta["section"],
                "semesters": ",".join(str(s) for s in overall["semesters"]),
                "total": _num(overall["total"]),
                "actual": _num(overall["actual"]),
                "actual_percent": actual_pct,
                "required": _num(overall.get("required_75", 0.0)),
                "final_percent": round(overall["percentage"], 2),
                "final": _num(overall["final"]),
                "bonus": _num(overall.get("bonus_15", 0.0)),
                "below_target": bool(actual_pct is not None
                                     and actual_pct < TARGET_PERCENT - EPS),
            })
        else:
            students.append({
                "sr_no": position, "usn": meta["usn"], "name": meta["name"],
                "type": meta["type"], "section": meta["section"],
                "semesters": "", "total": None, "actual": None,
                "actual_percent": None, "required": None,
                "final_percent": None, "final": None, "bonus": None,
                "below_target": False,
            })

    weighted_total = sum(o["total"] for o in overalls.values())
    weighted_actual = sum(o["actual"] for o in overalls.values())
    weighted_final = sum(o["final"] for o in overalls.values())
    required = sum(o.get("required_75", 0.0) for o in overalls.values())
    bonus = sum(o.get("bonus_15", 0.0) for o in overalls.values())
    below_after = sum(1 for o in overalls.values()
                      if o["percentage"] < TARGET_PERCENT - EPS)
    semester_rows: List[dict] = []
    for sem in sorted(result.semester_stats):
        bucket = result.semester_stats[sem]
        semester_rows.append({
            "semester": sem, "students": bucket["students"],
            "total": int(bucket["total"]), "actual": int(bucket["actual"]),
            "current": round(bucket["current"], 2),
            "required": int(bucket["required_75"]), "bonus": int(bucket["bonus_15"]),
            "final": int(bucket["final"]), "percent": round(bucket["percentage"], 2),
        })

    return {
        "rules": rules_snapshot(),
        "master_generated": result.master_generated,
        "metrics": {
            "students": len(result.layout.student_rows),
            "students_matched": stats["students_matched"],
            "students_unmatched": stats["students_unmatched"],
            "students_other_batch": stats.get("students_other_batch", 0),
            "students_without_data": stats["master_students_without_data"],
            "subjects": stats["subjects_processed"],
            "records": stats["records_read"],
            "records_written": stats["records_integrated"],
            "semesters": stats["semesters_processed"],
            "duplicates": stats["duplicates_ignored"],
            "missing": stats["missing_attendance"],
            "subject_mismatches": stats["subject_mismatches"],
            "below_target": sum(1 for s in students if s["below_target"]),
            "at_or_above_target": sum(1 for s in students
                                      if not s["below_target"] and s["total"]),
            "below_target_after": below_after,
            "weighted_total": int(weighted_total),
            "weighted_actual": int(weighted_actual),
            "weighted_final": int(weighted_final),
            "required": int(required),
            "bonus": int(bonus),
            "weighted_percent": round(weighted_final / weighted_total * 100.0, 2)
                                if weighted_total else 0.0,
            "actual_percent": round(weighted_actual / weighted_total * 100.0, 2)
                              if weighted_total else 0.0,
            "validation": "PASS" if not getattr(result, "validation_errors", [])
                          else "FAIL",
        },
        "subjects_found": subjects,
        "students": students,
        "semesters": semester_rows,
        "errors": list(getattr(result, "validation_errors", [])),
        "unmatched_students": [list(r) for r in result.unmatched_students],
        "unmatched_subjects": [list(r) for r in result.unmatched_subjects],
        "other_batch": [list(r) for r in result.other_batch_students],
        "log_rows": [list(r) for r in result.log_rows],
        "log_columns": list(LOG_COLUMNS),
        "overall_columns": list(OVERALL_COLUMNS),
        "semester_columns": list(SEMESTER_SUMMARY_COLUMNS),
    }


def main() -> int:
    _print_banner()

    try:
        report = run_pipeline(resolve_input_dir())
    except (FileNotFoundError, ValueError) as exc:
        print(f"\nERROR: {exc}")
        return 1

    print_report(report)
    return 0 if not report.errors else 2


if __name__ == "__main__":
    sys.exit(main())
