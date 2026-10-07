# College Attendance Integrator

Python tool that reads per-semester attendance exports (optionally matched
against a MASTER student workbook) and writes one integrated workbook.
It runs from the command line and as a small Streamlit web app where the
teacher simply uploads the files and downloads the result.

## Folder structure

```text
attendance_integrator/
│
├── attendance_integrator.py        # the tool (pipeline + command line)
├── app.py                          # Streamlit web app (same pipeline)
├── test_attendance_integrator.py   # tests (71 cases)
├── requirements.txt
├── README.md
│
├── my_attendance/                  # ANY folder with your exports (example)
│   ├── MASTER.xlsx                 # student database  (NEVER modified, optional)
│   ├── Attendance_Sem-I.xls        # attendance exports (.xlsx or .xls)
│   ├── Attendance_Sem-II.xls       # one file per semester is enough –
│   └── ...                         # only the semesters with data get a sheet
│
└── output/
    └── Integrated_Attendance.xlsx  # result (created on the first run)
```

File names do **not** have to be exact. The tool scans every `.xlsx` / `.xls`
file of the input folder **and of its sub-folders**:

* the workbook with `Enrolment No.` + `Student Name` that shares enrolment
  numbers with the attendance files is used as **MASTER**
  (if no workbook matches, the student roster is built from the attendance
  exports – see "Without a MASTER" below);
* every workbook with `USN NO` + `COURSE_CODE1...N` is an attendance file.

`.xls` files that are actually HTML exports (as produced by the college ERP)
are read automatically, and `.xlsx` files may contain any number of sheets –
only the sheets that really contain attendance data are used.

## Run

```bash
pip install -r requirements.txt
python attendance_integrator.py
```

Without an argument the tool uses the folder whose attendance files were
updated most recently (`input/` wins a tie when it exists). A folder can also
be given explicitly – useful for a new export:

```bash
python attendance_integrator.py "C:\path\to\your\attendance-folder"
```

Result: `output/Integrated_Attendance.xlsx`

Run the tests with:

```bash
python test_attendance_integrator.py
```

## Web app (Streamlit)

```bash
pip install -r requirements.txt
streamlit run app.py
```

The app uses the same pipeline as the command line:

1. **Give the attendance data** – upload the `.xls` / `.xlsx` exports (plus
   `MASTER.xlsx` when you have one) or point the app at a folder on the
   machine. Subjects and semesters are detected from the files themselves.
2. **Rules in the sidebar** – the requirement percentage and both bands can
   be changed (`configure_rules` validates the combination); the report
   headings follow whatever is in force.
3. **Run integration** – students below the requirement are lifted into the
   low band, regular students at/above it into the high band; semester and
   overall figures are the weighted totals (never averages of percentages).
4. **Result** – overview metrics (total classes, attended, percentage before
   and after), one row per student with the below-target filter, the detected
   subjects, per-semester totals, the full workbook sheets, the integration
   log and any issues (unmatched students/subjects, other-batch exclusions) –
   then **Download the semester sheet(s)**: the downloaded file contains the
   semester sheet(s) only (the college-template sheet a teacher shares),
   while Summary, Integration_Log, Unmatched_*, Other_Batch_Students and
   Integrated Master stay visible in the app's tabs (optionally also saved
   a copy into `output/` – same semester-only content).

## Without a MASTER

When none of the workbooks shares enrolment numbers with the attendance files
(for example a folder that contains only the attendance exports), the tool
does not stop and does not guess:

* a roster sheet shaped exactly like a MASTER (`Sl. No`, `Student Name`,
  `USN`, `Course Name`, `Batch/Year`, `Regular/Lateral/Ex-Regular` + one merged
  block per semester) is built from the exports,
* every attendance student is matched against it, so
  `Unmatched_Students` stays empty,
* students whose USN carries a different admission year (for example an older
  batch that is still listed in a batch export) are kept out of the roster:
  they are reported one line each in `Other_Batch_Students` and counted in
  `Summary` as `Students excluded (other batch)`. The year is read from the
  USN (`AJU/240916` → 2024) and only the majority year of the export is used,
  so a folder that really mixes batches is left untouched. A MASTER flow never
  filters - there the MASTER defines the roster,
* the sheet is exported as `Integrated Master` exactly like the MASTER flow,
  and `Summary` records `Master file: (built from the attendance files)`.

The MASTER flow itself is unchanged: with a MASTER present the same output
sheets, columns and formulas are produced.

## Matching rules

| What | MASTER | Attendance file |
|------|--------|-----------------|
| Student | `Enrolment No.` (e.g. `AJU/230064`) | `USN NO` – primary key, trimmed, case-insensitive, never matched by name |
| Semester | merged `Semester -01 ... -07` sections | `SEMESTER` column (`I..VII`, `1..7`, `Semester -01`), file name as fallback |
| Subject | course code in the subject heading (when the MASTER has one) | `COURSE_CODE1 ... COURSE_CODE19` |

Blanks are never invented: a blank course code is skipped, a missing
`TOTAL_CLASS` / `ATTENDED_CLASS` is logged, a student that is not in the
MASTER goes to `Unmatched_Students`, a course code that is not in the MASTER
subject section goes to `Unmatched_Subjects`.

## Attendance calculation (business rules)

The rules are applied **per subject**; semester and overall figures are the
weighted sums of those subjects.

**1. Below 75% – land inside 75 … 78%**

```text
75%  →  75 … 78%        66.67% (40/60)  →  76.67% (46/60)
```

**2. At/above 75% and below 95% – land inside 95 … 98%**

```text
75%  →  95 … 98%        83.33% (50/60)  →  96.67% (58/60)
```

Both bands are **spread deterministically** over their range: the exact point
inside the band is a stable function of `student + semester + course code`, so
the same student always gets the same value and different students are not all
pinned to the floor of the band.

```text
B = the class count that lands inside the band (spread across the band)
Attendance is never reduced, A + B never exceeds Total, the result never exceeds 100%.
```

* **at/above 95%** → untouched (`B = 0`), e.g. 95% stays 95%.
* **LATERAL / EX-REGULAR at/above 75%** → untouched (they only get the
  75–78% rule when below 75%); a blank student type is treated as regular.
* when the band is too narrow for the class length (e.g. 15 classes, one class
  = 6.67 points), the closest value **at or above** the floor of the band is
  used – attendance is still never reduced.

**3. Every percentage**

```text
Current % = Total Attended / Total Classes × 100     (actual, untouched)
Final %   = (Total Attended + B) / Total Classes × 100     (0 … 100)
```

**4. Totals are WEIGHTED – never an average of subject percentages**

```text
Semester % = Σ(attended) / Σ(total) × 100
Overall %  = Σ(all semesters attended) / Σ(all semesters total) × 100
```

Semesters without data are left out entirely – they get no sheet and no
summary row, and missing data is **never** shown as 0%.

Example: subjects 45/60 and 20/40 → 65/100 = **65%**, not `(75+50)/2 = 62.5%`.

`Validation: PASS` in the console means every subject percentage of the
workbook was recomputed from the business rules and every band holds.

## Output workbook

`output/Integrated_Attendance.xlsx` – sheets in this order:

| Sheet | Content |
|-------|---------|
| `Semester N` | **one sheet per semester that actually has attendance data** – one input file with semester I data yields exactly one sheet: the college template (row 2 = semester title + one `Sub: …` heading per subject, row 3 = column headers, one row per student in MASTER order, or roster order without a MASTER). Semesters without data get no sheet at all. |
| `Summary` | overall Semester 1–7 attendance for every student, semester-wise weighted totals, report metrics, validation verdict |
| `Integration_Log` | one row per source record with its disposition |
| `Unmatched_Students` | attendance students that are not in the MASTER (empty when the roster was built from the exports) |
| `Unmatched_Subjects` | course codes that are not in the MASTER subject section |
| `Other_Batch_Students` | students of another admission year, taken out of a roster built from the exports (one line each; empty with a MASTER) |
| `Integrated Master` | the original MASTER sheet with its attendance filled in (nothing existing moved or deleted); without a MASTER it is the generated roster |

### Semester sheet layout (college template)

```
row 2   [H] "1st Semester"      [K] "Sub: Engineering Physics BTE22010"  …
row 3   Sr. No | Student Name | Father's Name | Enrolment No. | Course Name |
        Batch/Year | Regular /Lateral /Ex Regular | Total No. of Classess |
        Total No. of Classess Attended | Current Attendance % |
        (then, for every subject of the semester)
        Total No of Allotted Classes | Total No. of Classess |
        Total No of Classes Attended- Actual (A) | Total No of Classes Added (B) |
        Total No. of Classess Attended (A+B) | Current Attendance %
row 4.. one row per student, in roster order (Sr. No 1..N)
```

| Column | Meaning |
|--------|---------|
| Sr. No … Batch/Year | from the MASTER (`Father's Name`, `Course Name`, `Batch/Year` when the MASTER has them) |
| Regular /Lateral /Ex Regular | student type from the attendance export (blank when unknown) |
| `Total No. of Classess` (student) | Σ of the subject `Total No. of Classess` columns of the same row |
| `Total No. of Classess Attended` (student) | Σ of the subject `Attended (A+B)` columns |
| `Current Attendance %` (student) | `Attended (A+B) / Total × 100`, weighted |
| `Total No of Allotted Classes` | the export has a single class count, so it fills Allotted and Total (see note 2) |
| `Total No of Classes Attended- Actual (A)` | the untouched source value |
| `Total No of Classes Added (B)` | classes added by the 75–78% / 95–98% band rule |
| `Total No. of Classess Attended (A+B)` | `A + B`, capped at `Total` |
| `Current Attendance %` (subject) | `A+B / Total × 100`, never above 100% |

Subjects are listed in the order of the export; a subject whose source rows
carry no `TOTAL_CLASS` keeps its column but stays blank. Students without
usable data keep their row with blank cells – a semester without data gets
no sheet at all.

### Summary sheet

* report metrics (students, records, duplicates, mismatches, verdict),
* `STUDENT-WISE OVERALL ATTENDANCE (SEMESTER 1 TO 7)` – one row per student
  with the weighted totals of every semester that has data,
* `SEMESTER-WISE WEIGHTED ATTENDANCE` – one row per semester that has data
  (one input file → one row, never empty placeholder rows) + `ALL SEMESTERS`.

The original MASTER workbook is only opened for reading – its hash never
changes.

### Two notes about this MASTER file

1. This MASTER has **semester sections only** (no subject/course-code
   columns), so the group headings of the Semester sheets carry the course code
   (`Sub: BTE21259`) and the subject title when it is known
   (`Sub: Engineering Physics BTE22010`). Row-by-row subject data is always in
   `Integration_Log`.
2. The attendance export provides one class count per subject
   (`TOTAL_CLASS`), so that value fills both *"Total No of Allotted Classes"*
   and *"Total No. of Classess"* in the `Integrated Master` sheet.

## Extra columns added to the MASTER sheet (nothing existing was moved)

To the right of the last MASTER section:

* `Overall (Semesters 1-7)` – total, actual, B, A+B, overall weighted %
* `Attendance Source Info` – SECTION, BRANCH, SESSION, SCHEME, STUDENT TYPE,
  SEMESTERS WITH DATA (filled from the attendance export when the MASTER cell
  is blank)
