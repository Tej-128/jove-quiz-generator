"""
JoVE Quiz Generator - Streamlit App v3.3 Batch QA
Generates one quiz Excel per chapter folder and returns a ZIP.
"""

from __future__ import annotations

import gc
import io
import os
import re
import tempfile
import traceback
import zipfile
from pathlib import Path
from typing import Any

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

st.session_state.setdefault("batch_result", None)
st.session_state.setdefault("upload_signature", None)

TEAM_OPENAI_API_KEY = ""


def _get_configured_api_key() -> str:
    if TEAM_OPENAI_API_KEY.strip():
        return TEAM_OPENAI_API_KEY.strip()
    try:
        secret_value = st.secrets.get("OPENAI_API_KEY", "")
    except Exception:
        secret_value = ""
    return str(secret_value or os.environ.get("OPENAI_API_KEY", "")).strip()


api_key = _get_configured_api_key()

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

    enable_ai_accuracy_qa = st.checkbox(
        "Run AI accuracy QA",
        value=True,
        help="Reviews generated questions against the uploaded source and replaces failed questions. This adds extra API calls and time.",
    )

    st.markdown("---")
    st.markdown(
        f"""
**Output spec:**
- {NUM_SETS} sets x {QUESTIONS_PER_SET} = **{NUM_SETS * QUESTIONS_PER_SET} questions per chapter**
- One Excel file per detected chapter folder
- Final download is one ZIP containing all chapter Excel files
- Duplicate QA runs before export
- AI accuracy QA can repair/regenerate failed questions
"""
    )
    st.markdown("---")
    st.caption("JoVE Internal Tool - v3.3 Batch QA")


# -----------------------------------------------------------------------------
# Upload and chapter grouping helpers
# -----------------------------------------------------------------------------


def _safe_unique_path(directory: str, original_name: str, counter: int) -> str:
    basename = Path(original_name).name
    return os.path.join(directory, f"upload_{counter}_{basename}")


def _clean_zip_parts(member: str) -> list[str]:
    parts = []
    for part in Path(member).parts:
        if part in {"", "."}:
            continue
        if part == "__MACOSX":
            return []
        parts.append(part)
    return parts


def _infer_chapter_key(relative_path: str) -> str:
    parts = list(Path(relative_path).parts)
    if len(parts) <= 1:
        return "Uploaded_Chapter"

    # Prefer explicit Chapter_xx folders anywhere in the uploaded ZIP.
    for part in parts[:-1]:
        if part.lower().startswith("chapter_"):
            return part

    # Otherwise use the immediate parent folder as the chapter group.
    return parts[-2]


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
                        parts = _clean_zip_parts(member)
                        if not parts:
                            continue
                        member_name = parts[-1]
                        if member.endswith("/"):
                            continue
                        if not member_name.lower().endswith(".docx"):
                            continue
                        if member_name.startswith("~$"):
                            continue

                        relative_path = "/".join(parts)
                        chapter_key = _infer_chapter_key(relative_path)
                        counter += 1
                        dest = _safe_unique_path(temp_dir, member_name, counter)
                        with open(dest, "wb") as output_file:
                            output_file.write(zip_file.read(member))
                        file_records.append({
                            "name": member_name,
                            "path": dest,
                            "relative_path": relative_path,
                            "chapter_key": chapter_key,
                        })
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
                file_records.append({
                    "name": uploaded_file.name,
                    "path": dest,
                    "relative_path": uploaded_file.name,
                    "chapter_key": "Uploaded_Chapter",
                })
            except Exception as exc:
                errors.append(f"{uploaded_file.name}: upload save failed: {exc}")

    return file_records, errors


def _group_records_by_chapter(file_records: list[dict[str, str]]) -> dict[str, list[dict[str, str]]]:
    grouped: dict[str, list[dict[str, str]]] = {}
    for record in file_records:
        key = record.get("chapter_key") or "Uploaded_Chapter"
        grouped.setdefault(key, []).append(record)
    return dict(sorted(grouped.items(), key=lambda item: item[0].lower()))


def _safe_filename(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_.-]+", "_", value).strip("_")
    return cleaned or "Chapter"


def _build_type_distribution(report: dict[str, Any]) -> pd.DataFrame:
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


def _build_chapter_preview(grouped: dict[str, list[dict[str, str]]]) -> pd.DataFrame:
    rows = []
    for chapter_key, records in grouped.items():
        rows.append({
            "Detected Chapter Folder": chapter_key,
            "DOCX Files": len(records),
            "Example File": records[0].get("name", "") if records else "",
        })
    return pd.DataFrame(rows)


# -----------------------------------------------------------------------------
# Main UI
# -----------------------------------------------------------------------------

st.markdown('<div class="main-title">JoVE Quiz Generator</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="subtitle">Generates 120 quiz questions per chapter and exports one Excel file per chapter.</div>',
    unsafe_allow_html=True,
)

if not api_key:
    st.error("OpenAI API key is not configured. Add OPENAI_API_KEY to Streamlit Secrets or set it as an environment variable.")
    st.stop()

st.markdown('<div class="section-hdr">Upload Chapter ZIP</div>', unsafe_allow_html=True)
st.info("Upload one ZIP that contains all Chapter folders. The app will return one ZIP containing one Excel quiz file per detected chapter.")

uploaded = st.file_uploader(
    "Drop ZIP or DOCX files here",
    type=["docx", "zip"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

file_records, upload_errors = _collect_uploaded_files(uploaded)

if not uploaded:
    st.info("Upload a ZIP containing chapter folders to get started.")
    st.stop()

if upload_errors:
    for error in upload_errors:
        st.markdown(f'<div class="err-box">{error}</div>', unsafe_allow_html=True)

if not file_records:
    st.error("No `.docx` files were found in the upload.")
    st.stop()

grouped_records = _group_records_by_chapter(file_records)

st.markdown("---")
st.markdown('<div class="section-hdr">Detected Chapters</div>', unsafe_allow_html=True)
st.dataframe(_build_chapter_preview(grouped_records), hide_index=True, use_container_width=True)

chapter_count = len(grouped_records)
docx_count = len(file_records)
cols = st.columns(3)
with cols[0]:
    st.markdown(f'<div class="stat-box"><div class="stat-label">Chapters Detected</div><div class="stat-value">{chapter_count}</div></div>', unsafe_allow_html=True)
with cols[1]:
    st.markdown(f'<div class="stat-box"><div class="stat-label">DOCX Files</div><div class="stat-value">{docx_count}</div></div>', unsafe_allow_html=True)
with cols[2]:
    st.markdown(f'<div class="stat-box"><div class="stat-label">Expected Questions</div><div class="stat-value">{chapter_count * NUM_SETS * QUESTIONS_PER_SET}</div></div>', unsafe_allow_html=True)

upload_signature = tuple((uploaded_file.name, getattr(uploaded_file, "size", None)) for uploaded_file in uploaded)
if st.session_state.upload_signature != upload_signature:
    st.session_state.upload_signature = upload_signature
    st.session_state.batch_result = None


def _render_batch_result(result: dict[str, Any]) -> None:
    st.markdown("---")
    st.markdown('<div class="section-hdr">Batch Results</div>', unsafe_allow_html=True)

    summary_df = pd.DataFrame(result.get("chapter_summaries", []))
    if not summary_df.empty:
        st.dataframe(summary_df, hide_index=True, use_container_width=True)

    if result.get("errors"):
        with st.expander(f"Chapter errors / skipped chapters ({len(result['errors'])})"):
            for error in result["errors"]:
                st.markdown(f'<div class="err-box">{error}</div>', unsafe_allow_html=True)

    if result.get("warnings"):
        with st.expander(f"Batch warnings / flags ({len(result['warnings'])})"):
            for warning in result["warnings"][:500]:
                st.markdown(f'<div class="warn-box">{warning}</div>', unsafe_allow_html=True)
            if len(result["warnings"]) > 500:
                st.caption(f"Showing first 500 of {len(result['warnings'])} warnings.")

    download_kwargs = {
        "label": "Download All Chapter Excel Files ZIP",
        "data": result["zip_bytes"],
        "file_name": "JoVE_Chapter_Quizzes.zip",
        "mime": "application/zip",
        "use_container_width": True,
        "type": "primary",
    }
    try:
        st.download_button(**download_kwargs, on_click="ignore")
    except TypeError:
        st.download_button(**download_kwargs)


st.markdown("---")
st.markdown('<div class="section-hdr">Generate Batch</div>', unsafe_allow_html=True)
st.markdown(
    f"Will generate one Excel per detected chapter folder: **{chapter_count} chapter(s)** x **{NUM_SETS * QUESTIONS_PER_SET} questions**."
)

if st.button("Generate All Chapter Quizzes", type="primary", use_container_width=True):
    progress_bar = st.progress(0)
    status_area = st.empty()
    log_area = st.empty()
    logs: list[str] = []
    chapter_summaries: list[dict[str, Any]] = []
    batch_warnings: list[str] = []
    batch_errors: list[str] = []
    completed_chapters: set[str] = set()
    output_files: list[tuple[str, str]] = []
    output_dir = tempfile.mkdtemp(prefix="jove_quiz_outputs_")

    def log(message: str, progress: int | None = None) -> None:
        logs.append(message)
        if progress is not None:
            progress_bar.progress(max(0, min(100, int(progress))))
        log_area.markdown("\n".join(f"- {line}" for line in logs[-12:]))

    try:
        for chapter_index, (chapter_key, records) in enumerate(grouped_records.items(), 1):
            chapter_base_progress = int(((chapter_index - 1) / max(1, chapter_count)) * 100)
            chapter_done_progress = int((chapter_index / max(1, chapter_count)) * 100)

            if chapter_key in completed_chapters:
                log(f"Skipping {chapter_key}: already completed in this run", chapter_done_progress)
                continue

            log(f"Processing {chapter_key} ({chapter_index}/{chapter_count})", chapter_base_progress)

            lessons = []
            parse_report: dict[str, Any] = {}
            all_sets = None
            report = None
            workbook = None

            try:
                lessons, parse_report = parse_lesson_files(records, return_report=True)
                chapter_name = parse_report.get("chapter_name", "")
                content_lessons = [lesson for lesson in lessons if lesson.get("pt_text") or lesson.get("transcript_text")]
                empty_lessons = [lesson["lesson_id"] for lesson in lessons if not lesson.get("pt_text") and not lesson.get("transcript_text")]
                skipped_files = parse_report.get("skipped_files", []) or []
                parse_warnings = list(parse_report.get("warnings", []) or [])

                if not lessons or not content_lessons:
                    message = f"{chapter_key}: skipped because no readable lesson content was detected."
                    batch_errors.append(message)
                    log(message, chapter_done_progress)
                    completed_chapters.add(chapter_key)
                    continue

                if not chapter_name:
                    message = f"{chapter_key}: skipped because chapter name was not found in PTx content."
                    batch_errors.append(message)
                    log(message, chapter_done_progress)
                    completed_chapters.add(chapter_key)
                    continue

                def progress_callback(message: str, progress: int | None = None) -> None:
                    if progress is None:
                        log(f"{chapter_key}: {message}")
                    else:
                        scaled = chapter_base_progress + int((progress / 100) * (chapter_done_progress - chapter_base_progress))
                        log(f"{chapter_key}: {message}", scaled)

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
                    enable_ai_accuracy_qa=enable_ai_accuracy_qa,
                )

                workbook = build_excel(all_sets, report, chapter_name)
                safe_chapter = _safe_filename(chapter_key)
                safe_chapter_name = _safe_filename(chapter_name)
                output_name = f"{safe_chapter}__{safe_chapter_name}.xlsx"
                output_path = os.path.join(output_dir, output_name)
                workbook.save(output_path)
                output_files.append((output_name, output_path))

                total_generated = sum(len(question_set) for question_set in all_sets)
                warning_count = len(report.get("warnings", []) or [])
                batch_warnings.extend([f"{chapter_key}: {warning}" for warning in (report.get("warnings", []) or [])])
                chapter_summaries.append({
                    "Chapter Folder": chapter_key,
                    "Chapter Name": chapter_name,
                    "Lessons": len(lessons),
                    "Content Lessons": len(content_lessons),
                    "Questions": total_generated,
                    "Warnings": warning_count,
                    "Output File": output_name,
                    "Status": "Completed",
                })
                completed_chapters.add(chapter_key)
                log(f"Saved {chapter_key} to disk and clearing memory", chapter_done_progress)

            except Exception as chapter_exc:
                error_message = f"{chapter_key}: failed - {chapter_exc}"
                batch_errors.append(error_message)
                chapter_summaries.append({
                    "Chapter Folder": chapter_key,
                    "Chapter Name": parse_report.get("chapter_name", "") if parse_report else "",
                    "Lessons": len(lessons) if lessons else 0,
                    "Content Lessons": "",
                    "Questions": 0,
                    "Warnings": "",
                    "Output File": "",
                    "Status": "Failed",
                })
                log(error_message, chapter_done_progress)
                log(traceback.format_exc())
                # Do not mark failed chapters as completed; they are not written to output.

            finally:
                # Clear large per-chapter objects only after the chapter is fully saved or failed.
                try:
                    if workbook is not None:
                        workbook.close()
                except Exception:
                    pass
                del lessons
                del parse_report
                del all_sets
                del report
                del workbook
                gc.collect()

        if not output_files:
            raise RuntimeError("No chapter Excel files were created.")

        final_zip_path = os.path.join(output_dir, "JoVE_Chapter_Quizzes.zip")
        with zipfile.ZipFile(final_zip_path, "w", zipfile.ZIP_DEFLATED) as output_zip:
            for output_name, output_path in output_files:
                if os.path.exists(output_path):
                    output_zip.write(output_path, arcname=output_name)

        with open(final_zip_path, "rb") as zip_file:
            zip_bytes = zip_file.read()

        if not zip_bytes:
            raise RuntimeError("No output ZIP was created.")

        st.session_state.batch_result = {
            "zip_bytes": zip_bytes,
            "chapter_summaries": chapter_summaries,
            "warnings": batch_warnings,
            "errors": batch_errors,
        }
        status_area.success("Batch generation complete.")
        progress_bar.progress(100)

    except Exception as exc:
        st.session_state.batch_result = None
        st.error(f"Batch generation failed: {exc}")
        st.code(traceback.format_exc())

if st.session_state.batch_result is not None:
    _render_batch_result(st.session_state.batch_result)
