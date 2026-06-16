"""
JoVE Quiz Generator - Streamlit App v3.2.4
Standalone tool: generates 120 questions (3 sets x 40) from PTx / Transcript docx files.
"""

from __future__ import annotations

import io
import os
import re
import tempfile
import traceback
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

from excel_export import build_excel
from quiz_generator import (
    NUM_SETS,
    QUESTIONS_PER_SET,
    QUESTION_TYPES,
    generate_quiz,
    parse_lesson_files,
)

# -----------------------------------------------------------------------------
# Page config and CSS
# -----------------------------------------------------------------------------

st.set_page_config(
    page_title="JoVE Quiz Generator",
    page_icon="JQ",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
<style>
.main-title{font-size:2rem;font-weight:800;color:#1a1a2e;border-bottom:3px solid #E63946;padding-bottom:8px;margin-bottom:4px}
.subtitle{font-size:.95rem;color:#666;margin-bottom:20px}
.section-hdr{font-size:1.05rem;font-weight:700;color:#1a1a2e;margin-top:18px;margin-bottom:6px}
.stat-box{background:#fff;border-radius:10px;padding:14px 18px;border:1px solid #ddd;box-shadow:0 1px 4px rgba(0,0,0,.06);margin:4px 0}
.stat-label{font-size:.75rem;color:#999;text-transform:uppercase;letter-spacing:.5px}
.stat-value{font-size:1.6rem;font-weight:800;color:#1a1a2e}
.warn-box{background:#fff8e1;border-left:4px solid #f9a825;padding:8px 12px;border-radius:4px;margin:4px 0;font-size:.85rem;color:#5d4037}
.ok-box{background:#e8f5e9;border-left:4px solid #388e3c;padding:8px 12px;border-radius:4px;margin:4px 0;font-size:.85rem;color:#1b5e20}
.err-box{background:#ffebee;border-left:4px solid #c62828;padding:8px 12px;border-radius:4px;margin:4px 0;font-size:.85rem;color:#b71c1c}
</style>
""",
    unsafe_allow_html=True,
)

# -----------------------------------------------------------------------------
# Session state
# -----------------------------------------------------------------------------

st.session_state.setdefault("quiz_result", None)
st.session_state.setdefault("upload_signature", None)

# -----------------------------------------------------------------------------
# API key - configured once for the team, not entered by each user
# -----------------------------------------------------------------------------

# Optional local/team slot. Leave blank when using Streamlit Secrets or env vars.
# Do not commit a real key to a shared/public repository.
TEAM_OPENAI_API_KEY = ""


def _get_configured_api_key() -> str:
    """Read the OpenAI key from team slot, Streamlit Secrets, then env-var fallback."""
    if TEAM_OPENAI_API_KEY.strip():
        return TEAM_OPENAI_API_KEY.strip()
    try:
        secret_value = st.secrets.get("OPENAI_API_KEY", "")
    except Exception:
        secret_value = ""
    return str(secret_value or os.environ.get("OPENAI_API_KEY", "")).strip()


api_key = _get_configured_api_key()

# -----------------------------------------------------------------------------
# Sidebar
# -----------------------------------------------------------------------------

with st.sidebar:
    st.markdown("## Configuration")
    if api_key:
        st.success("OpenAI API key configured for this app.")
    else:
        st.error("OPENAI_API_KEY is not configured in Streamlit Secrets or environment variables.")
    model = st.selectbox(
        "Model",
        ["gpt-5.5", "gpt-4o", "gpt-4-turbo", "gpt-4"],
        index=0,
        help="Use a strong model for better grounding and formatting adherence.",
    )
    st.markdown("---")
    st.markdown(
        f"""
**Output spec:**
- {NUM_SETS} sets x {QUESTIONS_PER_SET} = **{NUM_SETS * QUESTIONS_PER_SET} questions**
- 7 question types, minimum 5 per type per set unless fallback is flagged
- Question index restarts at 1 per set
- Summary tab appears first
- Match/Categorisation right_answer cells are blank
"""
    )
    st.markdown("---")
    st.caption("JoVE Internal Tool - v3.2.4")

# -----------------------------------------------------------------------------
# Upload and extraction helpers
# -----------------------------------------------------------------------------


def _safe_unique_path(directory: str, original_name: str, counter: int) -> str:
    basename = Path(original_name).name
    return os.path.join(directory, f"upload_{counter}_{basename}")


def _collect_uploaded_files(uploaded_files) -> tuple[list[dict[str, str]], list[str]]:
    file_records: list[dict[str, str]] = []
    errors: list[str] = []
    if not uploaded_files:
        return file_records, errors

    temp_dir = tempfile.mkdtemp()
    counter = 0

    for uploaded_file in uploaded_files:
        lower_name = uploaded_file.name.lower()
        if lower_name.endswith(".zip"):
            try:
                with zipfile.ZipFile(io.BytesIO(uploaded_file.read())) as zip_file:
                    for member in zip_file.namelist():
                        member_path = Path(member)
                        if member.endswith("/"):
                            continue
                        if "__MACOSX" in member:
                            continue
                        if not member_path.name.lower().endswith(".docx"):
                            continue
                        if member_path.name.startswith("~$"):
                            continue
                        counter += 1
                        dest = _safe_unique_path(temp_dir, member_path.name, counter)
                        with open(dest, "wb") as output_file:
                            output_file.write(zip_file.read(member))
                        # Keep the original basename for lesson-id parsing; only the disk path is made unique.
                        file_records.append({"name": member_path.name, "path": dest})
            except zipfile.BadZipFile:
                errors.append(f"{uploaded_file.name}: invalid or corrupted ZIP archive.")
            except Exception as exc:
                errors.append(f"{uploaded_file.name}: ZIP extraction failed: {exc}")
        else:
            counter += 1
            dest = _safe_unique_path(temp_dir, uploaded_file.name, counter)
            try:
                with open(dest, "wb") as output_file:
                    output_file.write(uploaded_file.getbuffer())
                file_records.append({"name": uploaded_file.name, "path": dest})
            except Exception as exc:
                errors.append(f"{uploaded_file.name}: upload save failed: {exc}")

    return file_records, errors


def _build_type_distribution(report: dict) -> pd.DataFrame:
    rows = []
    set_reports = report.get("sets", []) or []
    for question_type in QUESTION_TYPES:
        row = {"Question Type": question_type}
        total = 0
        for idx in range(NUM_SETS):
            set_report = set_reports[idx] if idx < len(set_reports) else {}
            count = int((set_report.get("type_counts", {}) or {}).get(question_type, 0))
            budget = set_report.get("budget", {}) or {}
            dropped = question_type in (set_report.get("dropped_types", []) or []) or question_type not in budget
            if dropped:
                status = "Not generated"
            elif count < 5 or count < int(budget.get(question_type, 5)):
                status = "Below minimum"
            else:
                status = "OK"
            row[f"Set {idx + 1} Count"] = count
            row[f"Set {idx + 1} Status"] = status
            total += count
        row["Total"] = total
        rows.append(row)
    return pd.DataFrame(rows)

# -----------------------------------------------------------------------------
# Main UI
# -----------------------------------------------------------------------------

st.markdown('<div class="main-title">JoVE Quiz Generator</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="subtitle">Generates 120 quiz questions across 3 non-overlapping sets from PTx and Transcript lesson files.</div>',
    unsafe_allow_html=True,
)

if not api_key:
    st.error("OpenAI API key is not configured. Add OPENAI_API_KEY to Streamlit Secrets or set it as an environment variable.")
    st.stop()

st.markdown('<div class="section-hdr">Upload Lesson Files</div>', unsafe_allow_html=True)
uploaded = st.file_uploader(
    "Drop files here",
    type=["docx", "zip"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

file_records, upload_errors = _collect_uploaded_files(uploaded)

if not uploaded:
    st.info("Upload lesson `.docx` files or a ZIP to get started.")
    st.stop()

if upload_errors:
    for error in upload_errors:
        st.markdown(f'<div class="err-box">{error}</div>', unsafe_allow_html=True)

if not file_records:
    st.error("No `.docx` files were found in the upload.")
    st.stop()

lessons, parse_report = parse_lesson_files(file_records, return_report=True)

st.markdown("---")
st.markdown('<div class="section-hdr">Detected Lessons</div>', unsafe_allow_html=True)

if not lessons:
    st.error("No valid lesson files found. Make sure lesson filenames start with a numeric Lesson ID.")
    skipped = parse_report.get("skipped_files", []) or []
    if skipped:
        with st.expander("Skipped files"):
            st.dataframe(pd.DataFrame(skipped), hide_index=True, use_container_width=True)
    st.stop()

chapter_name = parse_report.get("chapter_name", "")
content_lessons = [lesson for lesson in lessons if lesson.get("pt_text") or lesson.get("transcript_text")]
empty_lessons = [lesson["lesson_id"] for lesson in lessons if not lesson.get("pt_text") and not lesson.get("transcript_text")]
skipped_files = parse_report.get("skipped_files", []) or []
parse_warnings = list(parse_report.get("warnings", []) or [])
parse_warnings.extend(upload_errors)

if not chapter_name:
    st.markdown(
        '<div class="err-box">Chapter name was not found in any PTx file. Generation is blocked because the Excel chapter name must come from PTx content.</div>',
        unsafe_allow_html=True,
    )

st.success(
    f"Chapter: {chapter_name or 'Not extracted'} | Lessons found: {len(lessons)} | With content: {len(content_lessons)}"
)

cols = st.columns(4)
with cols[0]:
    st.markdown(
        f'<div class="stat-box"><div class="stat-label">Lessons Found</div><div class="stat-value">{len(lessons)}</div></div>',
        unsafe_allow_html=True,
    )
with cols[1]:
    st.markdown(
        f'<div class="stat-box"><div class="stat-label">With PTx</div><div class="stat-value">{sum(1 for lesson in lessons if lesson.get("pt_text"))}</div></div>',
        unsafe_allow_html=True,
    )
with cols[2]:
    st.markdown(
        f'<div class="stat-box"><div class="stat-label">With Transcript</div><div class="stat-value">{sum(1 for lesson in lessons if lesson.get("transcript_text"))}</div></div>',
        unsafe_allow_html=True,
    )
with cols[3]:
    total_words = sum(len((lesson.get("pt_text", "") + " " + lesson.get("transcript_text", "")).split()) for lesson in lessons)
    st.markdown(
        f'<div class="stat-box"><div class="stat-label">Total Words</div><div class="stat-value">{total_words:,}</div></div>',
        unsafe_allow_html=True,
    )

with st.expander("Lesson content breakdown"):
    breakdown_rows = []
    for lesson in lessons:
        pt_words = len(lesson.get("pt_text", "").split())
        transcript_words = len(lesson.get("transcript_text", "").split())
        total = pt_words + transcript_words
        breakdown_rows.append({
            "Lesson ID": lesson["lesson_id"],
            "PTx File": lesson.get("pt_filename", ""),
            "Transcript File": lesson.get("transcript_filename", ""),
            "PTx Words": pt_words,
            "Transcript Words": transcript_words,
            "Total Words": total,
            "Status": "Ready" if total > 0 else "Empty - skipped",
        })
    st.dataframe(pd.DataFrame(breakdown_rows), hide_index=True, use_container_width=True)

if skipped_files:
    with st.expander(f"Skipped / ignored files ({len(skipped_files)})"):
        st.dataframe(pd.DataFrame(skipped_files), hide_index=True, use_container_width=True)

if empty_lessons:
    st.markdown(
        f'<div class="warn-box">Lessons with no readable PTx or Transcript content will be skipped: {", ".join(empty_lessons)}</div>',
        unsafe_allow_html=True,
    )

if parse_warnings:
    with st.expander(f"Input warnings ({len(parse_warnings)})"):
        for warning in parse_warnings:
            st.markdown(f'<div class="warn-box">{warning}</div>', unsafe_allow_html=True)

if not content_lessons:
    st.markdown('<div class="err-box">All detected lesson files are empty. Cannot generate questions.</div>', unsafe_allow_html=True)
    st.stop()

if not chapter_name:
    st.stop()

st.markdown("---")
st.markdown('<div class="section-hdr">Generate Quiz</div>', unsafe_allow_html=True)
st.markdown(
    f"Will generate {NUM_SETS} sets x {QUESTIONS_PER_SET} questions = {NUM_SETS * QUESTIONS_PER_SET} total using {len(content_lessons)} content-bearing lessons."
)

upload_signature = tuple((uploaded_file.name, getattr(uploaded_file, "size", None)) for uploaded_file in uploaded)
if st.session_state.upload_signature != upload_signature:
    st.session_state.upload_signature = upload_signature
    st.session_state.quiz_result = None


def _render_results(result: dict) -> None:
    all_sets = result["all_sets"]
    report = result["report"]
    excel_bytes = result["excel_bytes"]
    result_chapter_name = result["chapter_name"]

    st.markdown("---")
    st.markdown('<div class="section-hdr">Results</div>', unsafe_allow_html=True)

    total_generated = sum(len(question_set) for question_set in all_sets)
    target = NUM_SETS * QUESTIONS_PER_SET
    result_cols = st.columns(4)
    with result_cols[0]:
        st.markdown(
            f'<div class="stat-box"><div class="stat-label">Total Generated</div><div class="stat-value">{total_generated}</div></div>',
            unsafe_allow_html=True,
        )
    with result_cols[1]:
        st.markdown(
            f'<div class="stat-box"><div class="stat-label">Target</div><div class="stat-value">{target}</div></div>',
            unsafe_allow_html=True,
        )
    with result_cols[2]:
        st.markdown(
            f'<div class="stat-box"><div class="stat-label">Sets</div><div class="stat-value">{len(all_sets)}</div></div>',
            unsafe_allow_html=True,
        )
    with result_cols[3]:
        st.markdown(
            f'<div class="stat-box"><div class="stat-label">Warnings</div><div class="stat-value">{len(report.get("warnings", []))}</div></div>',
            unsafe_allow_html=True,
        )

    st.markdown("**Question type distribution across sets:**")
    distribution_df = _build_type_distribution(report)
    st.dataframe(distribution_df, hide_index=True, use_container_width=True)

    dropped_types = report.get("fallback_types_dropped", []) or []
    if dropped_types:
        st.markdown(
            f'<div class="warn-box">Fallback applied. Dropped types: {", ".join(dropped_types)}</div>',
            unsafe_allow_html=True,
        )

    warnings = report.get("warnings", []) or []
    if warnings:
        with st.expander(f"Warnings / flags ({len(warnings)})"):
            for warning in warnings:
                st.markdown(f'<div class="warn-box">{warning}</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="ok-box">No warnings. All accepted questions passed validation.</div>', unsafe_allow_html=True)

    st.markdown("---")
    safe_chapter = re.sub(r"[^A-Za-z0-9_.-]+", "_", result_chapter_name).strip("_") if result_chapter_name else "Chapter"
    download_kwargs = {
        "label": "Download Quiz Excel",
        "data": excel_bytes,
        "file_name": f"JoVE_Quiz_{safe_chapter}.xlsx",
        "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "use_container_width": True,
        "type": "primary",
    }
    try:
        st.download_button(**download_kwargs, on_click="ignore")
    except TypeError:
        # Older Streamlit versions do not support on_click="ignore". Session state still
        # keeps the results visible after the rerun.
        st.download_button(**download_kwargs)


if st.button(f"Generate {NUM_SETS * QUESTIONS_PER_SET} Questions", type="primary", use_container_width=True):
    progress_bar = st.progress(0)
    status_area = st.empty()
    log_area = st.empty()
    logs: list[str] = []

    def progress_callback(message: str, progress: int | None = None) -> None:
        logs.append(message)
        if progress is not None:
            progress_bar.progress(max(0, min(100, int(progress))))
        log_area.markdown("\n".join(f"- {log}" for log in logs[-8:]))

    try:
        progress_bar.progress(3)
        all_sets, report = generate_quiz(
            lessons=content_lessons,
            api_key=api_key,
            model=model,
            progress_callback=progress_callback,
            all_lessons=lessons,
            skipped_lessons=empty_lessons,
            parse_warnings=parse_warnings,
            skipped_files=skipped_files,
            chapter_name=chapter_name,
        )

        progress_callback("Building Excel output", 90)
        workbook = build_excel(all_sets, report, chapter_name)
        buffer = io.BytesIO()
        workbook.save(buffer)
        excel_bytes = buffer.getvalue()
        progress_callback("Excel output ready", 100)
        status_area.success("Done.")

        st.session_state.quiz_result = {
            "all_sets": all_sets,
            "report": report,
            "excel_bytes": excel_bytes,
            "chapter_name": chapter_name,
        }

    except Exception as exc:
        st.session_state.quiz_result = None
        st.error(f"Generation failed: {exc}")
        st.code(traceback.format_exc())

if st.session_state.quiz_result is not None:
    _render_results(st.session_state.quiz_result)
