"""
Streamlit front end of the College Attendance Integrator.

    streamlit run app.py

The teacher gives the attendance data (uploaded files or a folder on this
machine), the subjects are detected by the same pipeline the command line
uses, the attendance rules are applied (below the target -> land inside the
low band, at/above the target regular students -> land inside the high band)
and the finished semester sheet(s) are offered as a download - the other
sheets of the workbook (Summary, log, unmatched, MASTER) are only shown in
the app.  Nothing about the run is assumed here: every number shown comes
from attendance_integrator.build_report.
"""

import io
import shutil
import tempfile
import time
import traceback
from pathlib import Path

import pandas as pd
import streamlit as st

import attendance_integrator as ai

PROJECT_DIR = Path(__file__).resolve().parent
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

st.set_page_config(page_title="College Attendance Integrator",
                   page_icon=None, layout="wide")


# --------------------------------------------------------------------------- #
# sidebar: the attendance rules in force (nothing is hardcoded in the report)
# --------------------------------------------------------------------------- #
def sidebar_rules() -> None:
    st.sidebar.header("Attendance rules")
    rules = ai.rules_snapshot()
    target = st.sidebar.number_input(
        "Attendance requirement (%)", min_value=1.0, max_value=100.0, step=1.0,
        value=float(rules["target_percent"]),
        help="Below this percentage extra classes are added.")
    low_min = st.sidebar.number_input(
        "Low band starts (%)", min_value=0.0, max_value=100.0, step=1.0,
        value=float(rules["low_band_min"]))
    low_max = st.sidebar.number_input(
        "Low band ends (%)", min_value=0.0, max_value=100.0, step=1.0,
        value=float(rules["low_band_max"]))
    high_min = st.sidebar.number_input(
        "High band starts (%)", min_value=0.0, max_value=100.0, step=1.0,
        value=float(rules["high_band_min"]))
    high_max = st.sidebar.number_input(
        "High band ends (%)", min_value=0.0, max_value=100.0, step=1.0,
        value=float(rules["high_band_max"]))

    st.sidebar.caption(
        f"below {ai._fmt_percent(target)}%  ->  land inside "
        f"{ai._fmt_percent(low_min)} - {ai._fmt_percent(low_max)}%  |  "
        f"at/above {ai._fmt_percent(target)}% (regular)  ->  "
        f"{ai._fmt_percent(high_min)} - {ai._fmt_percent(high_max)}%")

    if st.sidebar.button("Apply rules", width="stretch"):
        try:
            ai.configure_rules(target_percent=target,
                               low_band=(low_min, low_max),
                               high_band=(high_min, high_max))
            st.sidebar.success("Rules updated.")
        except ValueError as exc:
            st.sidebar.error(str(exc))


# --------------------------------------------------------------------------- #
# input: uploaded files or a folder the server can read
# --------------------------------------------------------------------------- #
def stage_uploads(uploads) -> Path:
    """Write the uploaded files into a fresh temp folder and return it."""
    previous = st.session_state.pop("staged_dir", None)
    if previous:
        shutil.rmtree(previous, ignore_errors=True)
    staged = Path(tempfile.mkdtemp(prefix="attendance_upload_"))
    for item in uploads:
        (staged / item.name).write_bytes(item.getbuffer())
    st.session_state["staged_dir"] = str(staged)
    return staged


def input_section() -> None:
    st.subheader("1. Give the attendance data")
    mode = st.radio("How do you want to provide the files?",
                    ("Upload files", "Folder on this machine"), horizontal=True)
    if mode == "Upload files":
        uploads = st.file_uploader(
            "Attendance exports (.xls / .xlsx) plus MASTER.xlsx when you have one",
            accept_multiple_files=True,
            type=["xls", "xlsx", "xlsm"],
            help="Semester or attendance exports; subjects and semesters are "
                 "detected from the files themselves.")
        st.session_state["uploads"] = uploads
        st.session_state["folder_input"] = ""
    else:
        default = str(PROJECT_DIR / "input") if (PROJECT_DIR / "input").exists() else ""
        folder = st.text_input("Folder containing the attendance files",
                               value=st.session_state.get("folder_input", default),
                               help="Every .xls / .xlsx below that folder is read.")
        st.session_state["folder_input"] = folder
        st.session_state["uploads"] = []


def resolve_input() -> Path | None:
    uploads = st.session_state.get("uploads") or []
    if uploads:
        return stage_uploads(uploads)
    folder = (st.session_state.get("folder_input") or "").strip()
    if folder:
        return Path(folder).expanduser()
    return None


# --------------------------------------------------------------------------- #
# the run itself - one code path with the command line
# --------------------------------------------------------------------------- #
def run_pipeline() -> None:
    input_dir = resolve_input()
    if input_dir is None:
        st.warning("Give the attendance files first (upload them or point at a folder).")
        return

    messages: list[str] = []
    out_dir = Path(tempfile.mkdtemp(prefix="attendance_out_"))
    out_path = out_dir / "Integrated_Attendance.xlsx"
    started = time.monotonic()
    try:
        with st.spinner("Reading files, matching students, applying the rules..."):
            report = ai.run_pipeline(input_dir, output_path=out_path,
                                     log=messages.append)
    except (FileNotFoundError, ValueError) as exc:
        st.error(str(exc))
        for line in messages:
            st.caption(line)
        return
    except Exception:                                   # pragma: no cover - UI
        st.exception(traceback.format_exc())
        return

    payload = report.output_path.read_bytes()
    st.session_state["att_run"] = {
        "payload": payload,
        "file_name": report.output_path.name,
        "data": ai.build_report(report.result),
        "messages": messages,
        "seconds": round(time.monotonic() - started, 2),
        "source": str(report.input_dir),
        "errors": list(report.errors),
    }
    st.session_state.pop("staged_dir", None)


# --------------------------------------------------------------------------- #
# results
# --------------------------------------------------------------------------- #
def arrow_safe(frame: pd.DataFrame) -> pd.DataFrame:
    """Give every column an Arrow-friendly type: numbers stay numbers,
    blanks become nulls, mixed columns fall back to clean strings."""
    columns = {}
    for name in frame.columns:
        series = frame[name]
        if pd.api.types.is_bool_dtype(series):
            columns[name] = series
            continue
        cleaned = series.replace({"": pd.NA})
        numeric = pd.to_numeric(cleaned, errors="coerce")
        if (cleaned.isna() | numeric.notna()).all():
            columns[name] = numeric
        else:
            columns[name] = cleaned.map(
                lambda v: "" if pd.isna(v) else str(v))
    return pd.DataFrame(columns, index=frame.index)


def df_from_rows(columns, rows):
    return arrow_safe(pd.DataFrame(rows, columns=columns))


def render_results() -> None:
    run = st.session_state["att_run"]
    data = run["data"]
    metrics = data["metrics"]
    rules = data["rules"]
    target = rules["target_percent"]

    st.subheader("2. Result")
    head = st.columns([3, 1, 1])
    head[0].markdown(
        f"**Source:** `{run['source']}`  |  **validation:** "
        f"{'PASS' if metrics['validation'] == 'PASS' else 'FAIL'}  |  "
        f"**{run['seconds']}s**")
    head[1].metric("Students", metrics["students"])
    head[2].metric("Subjects detected", len(data["subjects_found"]))

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total classes (all students)",
              f"{metrics['weighted_total']:,}")
    c2.metric("Attended (actual)", f"{metrics['weighted_actual']:,}",
              f"{metrics['actual_percent']:.2f}%")
    c3.metric("After required classes",
              f"{metrics['weighted_final']:,}",
              f"{metrics['weighted_percent']:.2f}%")
    c4.metric(f"Below {ai._fmt_percent(target)}% (actual)",
              f"{metrics['below_target']} / {metrics['students']}",
              f"{metrics['required']} classes to add")

    d1, d2, d3 = st.columns(3)
    d1.metric("Required classes (low band)", f"{metrics['required']:,}")
    d2.metric("Bonus classes (high band)", f"{metrics['bonus']:,}")
    d3.metric("Students below target after the run",
              f"{metrics['below_target_after']}")

    st.caption(
        f"rules in force: below {ai._fmt_percent(target)}% -> "
        f"{ai._fmt_percent(rules['low_band_min'])}-{ai._fmt_percent(rules['low_band_max'])}%, "
        f"regular at/above target -> "
        f"{ai._fmt_percent(rules['high_band_min'])}-{ai._fmt_percent(rules['high_band_max'])}%, "
        f"never above {ai._fmt_percent(rules['max_percent'])}%")

    tab_students, tab_subjects, tab_semesters, tab_sheets, tab_log, tab_issues = st.tabs(
        ["Students", "Subjects", "Semesters", "Workbook sheets", "Integration log", "Issues"])

    with tab_students:
        only_below = st.checkbox(
            f"Only students below {ai._fmt_percent(target)}% (actual)", value=False)
        students = data["students"]
        if only_below:
            students = [s for s in students if s["below_target"]]
        frame = pd.DataFrame(students)
        if not frame.empty:
            frame.columns = ["Sr", "Enrolment No.", "Name", "Type", "Section",
                             "Semesters", "Total classes", "Attended",
                             "Actual %", "Required", "Final %", "Final (A+B)",
                             "Bonus", "Below target"]
            # counts as whole numbers, percentages as 2-decimal numbers
            for name in ("Sr", "Total classes", "Attended", "Required",
                         "Final (A+B)", "Bonus"):
                frame[name] = pd.to_numeric(frame[name],
                                            errors="coerce").astype("Int64")
            for name in ("Actual %", "Final %"):
                frame[name] = pd.to_numeric(frame[name],
                                            errors="coerce").round(2)
            frame = arrow_safe(frame)
            st.dataframe(frame, width="stretch", hide_index=True)
            below = [s for s in students if s["below_target"]]
            if below:
                st.caption(f"{len(below)} student(s) below the requirement "
                           f"in the actual data.")
        else:
            st.info("Nothing to show.")

    with tab_subjects:
        st.caption(f"{len(data['subjects_found'])} subject(s) detected from the "
                   f"files across {metrics['semesters']} semester(s) "
                   f"({metrics['subjects']:,} subject records processed).")
        st.dataframe(arrow_safe(pd.DataFrame(data["subjects_found"])),
                     width="stretch", hide_index=True)

    with tab_semesters:
        st.dataframe(arrow_safe(pd.DataFrame(data["semesters"])),
                     width="stretch", hide_index=True)
        st.caption("Semester and overall figures are weighted sums of the "
                   "subjects, never averages of percentages.")

    with tab_sheets:
        sheets = pd.read_excel(io.BytesIO(run["payload"]), sheet_name=None,
                               header=None)
        info_width = len(ai.SEM_INFO_COLUMNS)
        usable = {}
        for name, sheet in sheets.items():
            if sheet.empty or not sheet.notna().any().any():
                continue                       # nothing at all in this sheet
            if name.startswith("Semester ") and sheet.shape[1] <= info_width:
                continue                       # roster only - no attendance given
            usable[name] = arrow_safe(sheet)
        if usable:
            st.caption(f"{len(usable)} sheet(s) of the output workbook carry data")
            choice = st.selectbox("Sheet", list(usable))
            st.dataframe(usable[choice], width="stretch", hide_index=True)
        else:
            st.info("No sheet of the workbook carries data.")

    with tab_log:
        st.dataframe(df_from_rows(data["log_columns"], data["log_rows"]),
                     width="stretch", hide_index=True)
        with st.expander("Console output"):
            st.code("\n".join(run["messages"]))

    with tab_issues:
        for title, key in (("Unmatched students", "unmatched_students"),
                           ("Unmatched subjects", "unmatched_subjects"),
                           (ai.OTHER_BATCH_SHEET + " (excluded from the roster)",
                            "other_batch")):
            st.markdown(f"**{title}**")
            rows = data[key]
            if rows:
                st.dataframe(arrow_safe(pd.DataFrame(rows)), width="stretch",
                             hide_index=True)
            else:
                st.caption("none")
        if run["errors"]:
            st.error("\n".join(run["errors"]))

    download = ai.semester_only_bytes(run["payload"])
    st.download_button("Download the semester sheet(s)", data=download,
                       file_name=run["file_name"], mime=XLSX_MIME,
                       type="primary")
    st.caption("The file contains the semester sheet(s) only - Summary, "
               "Integration_Log, Unmatched_* and Integrated Master are shown "
               "in the tabs above but are not part of the download.")

    with st.expander("Save a copy into the output folder"):
        name = st.text_input("File name", value="Integrated_Attendance.xlsx")
        if st.button("Save copy"):
            try:
                ai.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
                target_path = ai.OUTPUT_DIR / (name or "Integrated_Attendance.xlsx")
                target_path.write_bytes(download)
                st.success(f"Saved: {target_path}")
            except OSError as exc:
                st.error(f"Could not save the copy: {exc}")


# --------------------------------------------------------------------------- #
sidebar_rules()
st.title("College Attendance Integrator")
st.caption("Upload the attendance exports (or point at a folder), the subjects "
           "are detected automatically, students below the requirement are "
           "lifted into the low band and the finished workbook is ready to "
           "download.")

input_section()
if st.button("Run integration", type="primary"):
    run_pipeline()

if "att_run" in st.session_state:
    render_results()
else:
    st.info("Give the attendance data above and press **Run integration**.")
