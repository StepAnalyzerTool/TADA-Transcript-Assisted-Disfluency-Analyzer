"""Streamlit interface for the Transcript-Assisted Disfluency Analyzer (TADA)."""

from __future__ import annotations

from datetime import date
from io import BytesIO
import hashlib
from pathlib import Path
import re

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from tada_core import (
    LEXICAL_DEFAULTS,
    NONLEXICAL_DEFAULTS,
    calculate_metrics,
    find_candidates,
    lexical_tokens,
    normalize_for_analysis,
    parse_transcript,
    read_transcript_upload,
)


APP_TITLE = "Smooth Talking: Transcript-Assisted Disfluency Analyzer (TADA)"

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="🎯",
    layout="wide",
)
st.markdown(
    """
    <style>
        :root {
            --tada-navy: #0b2f5b;
            --tada-teal: #0b8f9c;
            --tada-slate: #667085;
            --tada-pale: #f1f7fc;
            --tada-border: #b8d7f5;
        }
        .tada-brand { margin: 0.4rem 0 1.5rem; }
        .tada-mark {
            color: var(--tada-navy);
            font-size: clamp(3.6rem, 7vw, 6rem);
            font-weight: 800;
            letter-spacing: -0.055em;
            line-height: 0.9;
        }
        .tada-name {
            color: var(--tada-teal);
            font-size: clamp(1.25rem, 2.6vw, 2rem);
            font-weight: 750;
            letter-spacing: 0.12em;
            line-height: 1.25;
            margin-top: 1.25rem;
            text-transform: uppercase;
        }
        .tada-expanded {
            color: var(--tada-navy);
            font-size: clamp(1rem, 1.7vw, 1.3rem);
            font-weight: 600;
            margin-top: 0.3rem;
        }
        .tada-rule {
            background: linear-gradient(90deg, var(--tada-teal), var(--tada-navy));
            height: 3px;
            margin: 1.35rem 0 1.15rem;
            width: 100%;
        }
        .tada-tagline {
            color: var(--tada-slate);
            font-size: 1.15rem;
            font-style: italic;
        }
        section[data-testid="stSidebar"] { background-color: #f7fafc; }
        h2, h3 { color: var(--tada-navy); }
        div[data-testid="stAlert"] {
            background-color: var(--tada-pale);
            border-color: var(--tada-border);
            color: var(--tada-navy);
        }
        .stButton > button, .stDownloadButton > button {
            border-color: var(--tada-teal);
        }
        .stButton > button:hover, .stDownloadButton > button:hover {
            border-color: var(--tada-navy);
            color: var(--tada-navy);
        }
    </style>
    <div class="tada-brand">
        <div class="tada-mark">TADA</div>
        <div class="tada-name">Smooth Talking</div>
        <div class="tada-expanded">Transcript-Assisted Disfluency Analyzer</div>
        <div class="tada-rule"></div>
        <div class="tada-tagline">From the CVC Cosmos · Making every word count.</div>
    </div>
    """,
    unsafe_allow_html=True,
)
st.info("TADA identifies candidate speech events for human review.")

clickable_transcript = components.declare_component(
    "tada_clickable_transcript",
    path=str(Path(__file__).parent / "clickable_transcript"),
)


def split_targets(value: str) -> list[str]:
    return list(dict.fromkeys(item.strip().casefold() for item in value.split(",") if item.strip()))


def style_export_workbook(writer: pd.ExcelWriter) -> None:
    navy = "0B2F5B"
    teal = "0B8F9C"
    pale_blue = "F1F7FC"
    border_blue = "B8D7F5"
    white = "FFFFFF"
    dark = "243447"
    thin_border = Border(bottom=Side(style="thin", color=border_blue))
    tab_colors = {
        "Session_Summary": navy,
        "Consolidated_Summaries": navy,
        "Target_Metrics": teal,
        "Occurrence_Review": "6FA8DC",
        "Reviewed_Transcript": "98A2B3",
    }

    for sheet_name, worksheet in writer.sheets.items():
        worksheet.sheet_view.showGridLines = False
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
        worksheet.sheet_properties.tabColor = tab_colors.get(sheet_name, teal)
        worksheet.row_dimensions[1].height = 34

        for cell in worksheet[1]:
            cell.fill = PatternFill("solid", fgColor=navy)
            cell.font = Font(name="Arial", size=10, bold=True, color=white)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border

        for row_index in range(2, worksheet.max_row + 1):
            if row_index % 2 == 0:
                for cell in worksheet[row_index]:
                    cell.fill = PatternFill("solid", fgColor=pale_blue)
            for cell in worksheet[row_index]:
                cell.font = Font(name="Arial", size=10, color=dark)
                cell.alignment = Alignment(vertical="center")

        for column_index, column_cells in enumerate(worksheet.columns, start=1):
            header = str(worksheet.cell(1, column_index).value or "")
            sample_lengths = [len(str(cell.value)) for cell in column_cells if cell.value is not None]
            width = min(max(sample_lengths + [len(header)]) + 2, 42)
            if header in {"Context", "Reviewer_notes", "Original_Selected_Text", "Analyzed_Text"}:
                width = 55
            worksheet.column_dimensions[get_column_letter(column_index)].width = max(width, 12)

            if "Date" in header:
                for cell in list(column_cells)[1:]:
                    cell.number_format = "mm/dd/yyyy"
            elif "Per_100" in header or "Per_Minute" in header:
                for cell in list(column_cells)[1:]:
                    cell.number_format = "0.000"
            elif "Duration_Minutes" in header:
                for cell in list(column_cells)[1:]:
                    cell.number_format = "0.00"
            elif "Total" in header or "Occurrences" in header:
                for cell in list(column_cells)[1:]:
                    cell.number_format = "#,##0"


def export_workbook(
    summary: dict,
    metrics: list[dict],
    decisions: pd.DataFrame,
    original_text: str,
    analyzed_text: str,
) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        pd.DataFrame([summary]).to_excel(writer, sheet_name="Session_Summary", index=False)
        pd.DataFrame(metrics).to_excel(writer, sheet_name="Target_Metrics", index=False)
        decisions.to_excel(writer, sheet_name="Occurrence_Review", index=False)
        pd.DataFrame([{
            "Original_Selected_Text": original_text,
            "Analyzed_Text": analyzed_text,
        }]).to_excel(
            writer, sheet_name="Reviewed_Transcript", index=False
        )
        style_export_workbook(writer)
    return buffer.getvalue()


def read_review_workbook(uploaded_file) -> dict:
    sheets = pd.read_excel(uploaded_file, sheet_name=None)
    required = {"Session_Summary", "Occurrence_Review", "Reviewed_Transcript"}
    missing = required - set(sheets)
    if missing:
        raise ValueError("Missing required sheet(s): " + ", ".join(sorted(missing)))
    summary = sheets["Session_Summary"]
    decisions = sheets["Occurrence_Review"]
    required_summary = {
        "Participant_ID", "Session_Information", "Reviewer_ID", "Reviewer_Role", "Transcript_SHA256"
    }
    required_decisions = {"start", "end", "target", "category", "accepted"}
    if summary.empty or not required_summary.issubset(summary.columns):
        raise ValueError("The session summary is not from the current IOA-ready export format.")
    if not required_decisions.issubset(decisions.columns):
        raise ValueError("The occurrence review is not from the current IOA-ready export format.")
    row = summary.iloc[0]
    return {
        "summary": summary,
        "decisions": decisions,
        "participant": str(row["Participant_ID"]),
        "session": str(row["Session_Information"]),
        "reviewer": str(row["Reviewer_ID"]),
        "role": str(row["Reviewer_Role"]),
        "transcript_hash": str(row["Transcript_SHA256"]),
    }


def cell_bool(value) -> bool:
    if pd.isna(value):
        return False
    if isinstance(value, str):
        return value.strip().casefold() in {"true", "yes", "1", "accepted", "manually added"}
    return bool(value)


def occurrence_records(frame: pd.DataFrame) -> dict:
    """Create stable, duplicate-safe keys for automatic and manual occurrences."""
    records = {}
    duplicate_counts = {}
    for row in frame.to_dict("records"):
        target = str(row.get("target", "")).strip().casefold()
        if pd.notna(row.get("start")) and pd.notna(row.get("end")):
            base = ("position", int(float(row["start"])), int(float(row["end"])), target)
        else:
            base = (
                "manual",
                target,
                str(row.get("observed_text", "")).strip().casefold(),
                str(row.get("context", "")).strip().casefold(),
            )
        duplicate_counts[base] = duplicate_counts.get(base, 0) + 1
        records[base + (duplicate_counts[base],)] = row
    return records


def compare_review_workbooks(primary: dict, secondary: dict) -> tuple[pd.DataFrame, pd.DataFrame]:
    if primary["role"].casefold() != "primary" or secondary["role"].casefold() != "secondary":
        raise ValueError("Select a Primary review file first and a Secondary review file second.")
    if primary["participant"] != secondary["participant"]:
        raise ValueError("The files have different participant identifiers.")
    if primary["session"] != secondary["session"]:
        raise ValueError("The files have different session information.")
    if primary["transcript_hash"] != secondary["transcript_hash"]:
        raise ValueError("The files were generated from different transcript text.")

    records_a = occurrence_records(primary["decisions"])
    records_b = occurrence_records(secondary["decisions"])
    keys = sorted(set(records_a) | set(records_b), key=str)
    decision_agreements = both_accepted = either_accepted = 0
    discrepancies = []
    for key in keys:
        row_a, row_b = records_a.get(key), records_b.get(key)
        accepted_a = cell_bool(row_a.get("accepted")) if row_a else False
        accepted_b = cell_bool(row_b.get("accepted")) if row_b else False
        if row_a is not None and row_b is not None and accepted_a == accepted_b:
            decision_agreements += 1
        if accepted_a or accepted_b:
            either_accepted += 1
        if accepted_a and accepted_b:
            both_accepted += 1
        if row_a is None or row_b is None or accepted_a != accepted_b:
            sample = row_a or row_b
            discrepancies.append({
                "Target": sample.get("target", ""),
                "Observed_Text": sample.get("observed_text", ""),
                "Context": sample.get("context", ""),
                "Primary_Decision": "Accepted" if accepted_a else "Rejected / not present",
                "Secondary_Decision": "Accepted" if accepted_b else "Rejected / not present",
            })

    accepted_count_a = sum(cell_bool(row.get("accepted")) for row in records_a.values())
    accepted_count_b = sum(cell_bool(row.get("accepted")) for row in records_b.values())
    percent = lambda numerator, denominator: round(numerator / denominator * 100, 2) if denominator else 100.0
    result = pd.DataFrame([{
        "Participant_ID": primary["participant"],
        "Session_Information": primary["session"],
        "Primary_Reviewer": primary["reviewer"],
        "Secondary_Reviewer": secondary["reviewer"],
        "Decision_Agreement_%": percent(decision_agreements, len(keys)),
        "Occurrence_Agreement_%": percent(both_accepted, either_accepted),
        "Total_Count_Agreement_%": percent(min(accepted_count_a, accepted_count_b), max(accepted_count_a, accepted_count_b)),
        "Primary_Accepted": accepted_count_a,
        "Secondary_Accepted": accepted_count_b,
        "Discrepancies": len(discrepancies),
    }])
    return result, pd.DataFrame(discrepancies)


def export_ioa_results(results: pd.DataFrame, discrepancies: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        results.to_excel(writer, sheet_name="IOA_Results", index=False)
        discrepancies.to_excel(writer, sheet_name="Discrepancies", index=False)
    return buffer.getvalue()


def ordered_summary_columns(frame: pd.DataFrame) -> list[str]:
    preferred = [
        "Source_Workbook",
        "Participant_ID",
        "Session_Information",
        "Analysis_Date",
        "Reviewer_ID",
        "Reviewer_Role",
        "Source_File",
        "Source_Format",
        "Selected_Speaker",
        "Total_Words_Spoken",
        "Duration_Minutes",
        "Total_Disfluencies",
        "Total_Disfluencies_Per_100_Words",
        "Total_Disfluencies_Per_Minute",
        "Lexical_Total_Disfluencies",
        "Lexical_Disfluencies_Per_100_Words",
        "Lexical_Disfluencies_Per_Minute",
        "Nonlexical_Total_Disfluencies",
        "Nonlexical_Disfluencies_Per_100_Words",
        "Nonlexical_Disfluencies_Per_Minute",
    ]
    target_suffixes = (
        "_Category",
        "_Total_Disfluencies",
        "_Disfluencies_Per_100_Words",
        "_Disfluencies_Per_Minute",
    )
    target_groups: dict[str, list[str]] = {}
    for column in frame.columns:
        if not column.startswith("Target_"):
            continue
        for suffix in target_suffixes:
            if column.endswith(suffix):
                target_groups.setdefault(column[: -len(suffix)], []).append(column)
                break
    target_columns = []
    for target_prefix in sorted(target_groups, key=str.casefold):
        target_columns.extend(
            column for suffix in target_suffixes
            if (column := target_prefix + suffix) in frame.columns
        )
    audit_columns = ["Duration_Source", "Transcript_SHA256"]
    ordered = [column for column in preferred if column in frame.columns]
    ordered.extend(column for column in target_columns if column not in ordered)
    ordered.extend(column for column in audit_columns if column in frame.columns and column not in ordered)
    ordered.extend(column for column in frame.columns if column not in ordered)
    return ordered


def export_consolidated_summaries(frame: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        frame.to_excel(writer, sheet_name="Consolidated_Summaries", index=False)
        style_export_workbook(writer)
    return buffer.getvalue()


def render_summary_consolidator() -> None:
    st.header("Consolidate Summaries")
    st.caption(
        "Upload two or more Excel records generated by TADA. Only the Session_Summary sheet is consolidated."
    )
    uploaded_files = st.file_uploader(
        "TADA Excel records",
        type=["xlsx"],
        accept_multiple_files=True,
        key="summary_files",
    )
    if len(uploaded_files) < 2:
        st.info("Upload at least two TADA Excel records to create a consolidated summary.")
        return

    summary_frames = []
    errors = []
    for uploaded_file in uploaded_files:
        try:
            summary_frame = pd.read_excel(uploaded_file, sheet_name="Session_Summary")
            if summary_frame.empty:
                raise ValueError("Session_Summary is empty")
            summary_frame.insert(0, "Source_Workbook", uploaded_file.name)
            summary_frames.append(summary_frame)
        except Exception as error:
            errors.append(f"{uploaded_file.name}: {error}")
    if errors:
        st.error("Unable to consolidate:\n\n" + "\n\n".join(errors))
        return

    consolidated = pd.concat(summary_frames, ignore_index=True, sort=False)
    consolidated = consolidated[ordered_summary_columns(consolidated)]
    st.subheader("Consolidated preview")
    st.dataframe(consolidated, hide_index=True, use_container_width=True)
    st.caption(
        "Blank target cells mean that the source workbook did not contain that target column. "
        "They are not automatically interpreted as zero."
    )
    st.download_button(
        "Download consolidated summaries",
        data=export_consolidated_summaries(consolidated),
        file_name="TADA_Consolidated_Summaries.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def render_ioa_calculator() -> None:
    st.header("Interobserver Agreement Calculator")
    st.caption("Upload independently completed primary and secondary review files for the same session.")
    ioa_primary_col, ioa_secondary_col = st.columns(2)
    primary_file = ioa_primary_col.file_uploader(
        "Primary review file", type=["xlsx"], key="ioa_primary"
    )
    secondary_file = ioa_secondary_col.file_uploader(
        "Secondary review file", type=["xlsx"], key="ioa_secondary"
    )
    if not primary_file or not secondary_file:
        st.info("Upload both review files to calculate agreement.")
        return
    try:
        primary_data = read_review_workbook(primary_file)
        secondary_data = read_review_workbook(secondary_file)
        ioa_results, ioa_discrepancies = compare_review_workbooks(primary_data, secondary_data)
    except Exception as error:
        st.error(f"Unable to compare these files: {error}")
        return

    st.subheader("Agreement results")
    st.dataframe(ioa_results, hide_index=True, use_container_width=True)
    st.subheader("Discrepancies")
    if ioa_discrepancies.empty:
        st.success("No occurrence-level coding discrepancies were found.")
    else:
        st.dataframe(ioa_discrepancies, hide_index=True, use_container_width=True)
    with st.expander("How agreement is calculated"):
        st.markdown(
            "- **Decision agreement:** matching accept/reject decisions divided by all matched or unmatched candidate occurrences.\n"
            "- **Occurrence agreement:** occurrences accepted by both reviewers divided by occurrences accepted by either reviewer.\n"
            "- **Total-count agreement:** the smaller accepted-occurrence count divided by the larger count."
        )
    ioa_bytes = export_ioa_results(ioa_results, ioa_discrepancies)
    safe_ioa_participant = re.sub(
        r"[^A-Za-z0-9_-]+", "_", primary_data["participant"]
    ).strip("_") or "Participant"
    safe_ioa_session = re.sub(r"[^A-Za-z0-9_-]+", "_", primary_data["session"]).strip("_") or "Session"
    st.download_button(
        "Download IOA results",
        data=ioa_bytes,
        file_name=f"IOA_{safe_ioa_participant}_{safe_ioa_session}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


with st.sidebar:
    tool_mode = st.radio(
        "Select tool",
        ["Disfluency coding", "IOA calculator", "Consolidate summaries"],
    )

if tool_mode == "IOA calculator":
    render_ioa_calculator()
    st.stop()
if tool_mode == "Consolidate summaries":
    render_summary_consolidator()
    st.stop()

with st.sidebar:
    st.divider()
    st.header("Targets")
    lexical_value = st.text_area("Lexical words or phrases", ", ".join(LEXICAL_DEFAULTS))
    nonlexical_value = st.text_area("Nonlexical vocalizations", ", ".join(NONLEXICAL_DEFAULTS))
    st.caption("Separate targets with commas. Every literal occurrence will be presented for review.")

st.header("1. Add a transcript")
input_method = st.radio("Transcript source", ["Upload a file", "Paste text"], horizontal=True)
uploaded = None
raw_text = ""
source_name = "Pasted transcript"
source_format = "pasted text"
if input_method == "Upload a file":
    uploaded = st.file_uploader("Upload transcript", type=["txt", "docx", "vtt", "srt"])
    if uploaded:
        raw_text = read_transcript_upload(uploaded.getvalue(), uploaded.name)
        source_name = uploaded.name
        source_format = uploaded.name.rsplit(".", 1)[-1].lower()
else:
    raw_text = st.text_area("Paste transcript", height=240)

if not raw_text.strip():
    st.stop()

parsed = parse_transcript(raw_text, source_format)

st.header("2. Select speech and verify duration")
selected_text = parsed.full_text
if len(parsed.speaker_text) >= 2:
    options = ["All detected speech"] + sorted(parsed.speaker_text)
    chosen_speaker = st.selectbox("Speaker", options)
    if chosen_speaker != "All detected speech":
        selected_text = parsed.speaker_text[chosen_speaker]
elif len(parsed.speaker_text) == 1:
    only_speaker = next(iter(parsed.speaker_text))
    chosen_speaker = only_speaker
    selected_text = parsed.speaker_text[only_speaker]
    st.caption(
        f"Only one speaker label was detected ({only_speaker}). TADA cannot separate speakers in this file; "
        "remove other-speaker passages from the editable working copy below."
    )
else:
    chosen_speaker = "Not available"
    st.caption(
        "No reliable speaker labels were detected. Remove other-speaker passages from the editable working copy below."
    )

st.subheader("Edit text included in analysis")
st.caption(
    "To exclude speech from another person, select that passage below and delete it. "
    "This changes only the working copy; the original uploaded transcript is preserved."
)
editor_source_key = hashlib.sha256(
    (source_name + "|" + chosen_speaker + "|" + selected_text).encode()
).hexdigest()[:16]
editor_key = f"analysis_text_{editor_source_key}"
if editor_key not in st.session_state:
    st.session_state[editor_key] = selected_text
analysis_text = st.text_area(
    "Text included in analysis",
    key=editor_key,
    height=320,
)
if st.button("Restore original selected text"):
    st.session_state[editor_key] = selected_text
    st.rerun()

if not analysis_text.strip():
    st.warning("No text remains in the working copy. Restore the original text or retain speech to analyze.")
    st.stop()

detected_minutes = (
    parsed.detected_duration_seconds / 60 if parsed.detected_duration_seconds is not None else None
)
duration_source_default = "Detected from transcript timestamps" if detected_minutes else "Manually entered"
duration_source = st.selectbox(
    "Duration source",
    ["Detected from transcript timestamps", "Manually entered", "Not available"],
    index=0 if detected_minutes else 1,
)
if duration_source == "Detected from transcript timestamps":
    if detected_minutes is None:
        st.warning("No valid transcript timestamp span was detected. Enter the duration manually.")
        duration_seconds = None
    else:
        duration_minutes = st.number_input(
            "Verify duration (minutes)", min_value=0.01, value=float(detected_minutes), step=0.01
        )
        duration_seconds = duration_minutes * 60
elif duration_source == "Manually entered":
    duration_minutes = st.number_input(
        "Session duration (minutes)",
        min_value=0.01,
        value=None,
        step=0.01,
        placeholder="Enter verified duration",
    )
    duration_seconds = duration_minutes * 60 if duration_minutes is not None else None
else:
    duration_minutes = None
    duration_seconds = None

lexical_targets = split_targets(lexical_value)
nonlexical_targets = split_targets(nonlexical_value)
words = lexical_tokens(analysis_text, nonlexical_targets)
findings = find_candidates(analysis_text, lexical_targets, nonlexical_targets)

transcript_key = hashlib.sha256(
    (normalize_for_analysis(analysis_text) + "|" + ",".join(lexical_targets + nonlexical_targets)).encode()
).hexdigest()[:16]
state_key = f"accepted_{transcript_key}"
candidate_ids = [f"{row['start']}:{row['end']}:{row['target']}" for row in findings]
if state_key not in st.session_state:
    st.session_state[state_key] = candidate_ids.copy()

st.subheader("Visual audit")
st.caption(
    "Click a highlighted occurrence to accept or reject it. "
    "Blue = accepted lexical · Gold = accepted nonlexical · Gray = rejected"
)
component_findings = [
    {
        "id": candidate_id,
        "start": row["start"],
        "end": row["end"],
        "target": row["target"],
        "category": row["category"],
    }
    for candidate_id, row in zip(candidate_ids, findings)
]
component_value = clickable_transcript(
    text=normalize_for_analysis(analysis_text),
    findings=component_findings,
    accepted_ids=st.session_state[state_key],
    key=f"clickable_{transcript_key}",
    default=st.session_state[state_key],
)
if component_value is not None and list(component_value) != st.session_state[state_key]:
    st.session_state[state_key] = list(component_value)

accepted_ids = set(st.session_state[state_key])
for candidate_id, finding in zip(candidate_ids, findings):
    finding["accepted"] = candidate_id in accepted_ids
    finding["decision"] = "Accepted" if finding["accepted"] else "Rejected"

st.header("3. Review candidate occurrences")
st.caption("Use the visual audit to accept or reject occurrences. Add reviewer notes here when useful.")
review_df = pd.DataFrame(findings)
display_columns = ["accepted", "target", "category", "observed_text", "context", "notes"]
if review_df.empty:
    st.warning("No configured targets were detected.")
    reviewed = pd.DataFrame(columns=display_columns)
else:
    reviewed = st.data_editor(
        review_df[display_columns],
        hide_index=True,
        use_container_width=True,
        disabled=["accepted", "target", "category", "observed_text", "context"],
        column_config={
            "accepted": st.column_config.CheckboxColumn("Accept", default=True),
            "target": "Target",
            "category": "Category",
            "observed_text": "Observed text",
            "context": st.column_config.TextColumn("Context", width="large"),
            "notes": st.column_config.TextColumn("Reviewer notes", width="medium"),
        },
    )

reviewed_findings = []
for idx, row in reviewed.iterrows():
    original = findings[idx].copy()
    original["accepted"] = bool(row["accepted"])
    original["decision"] = "Accepted" if original["accepted"] else "Rejected"
    original["notes"] = str(row.get("notes", ""))
    reviewed_findings.append(original)

st.subheader("Add an occurrence missed by automatic detection")
with st.form("manual_occurrence", clear_on_submit=True):
    c1, c2, c3 = st.columns(3)
    manual_target = c1.text_input("Target")
    manual_category = c2.selectbox("Category", ["Lexical", "Nonlexical"])
    manual_context = c3.text_input("Context or location")
    manual_note = st.text_input("Reviewer note")
    add_manual = st.form_submit_button("Add occurrence")
if "manual_findings" not in st.session_state:
    st.session_state.manual_findings = []
if add_manual and manual_target.strip():
    st.session_state.manual_findings.append({
        "start": None,
        "end": None,
        "target": manual_target.strip().casefold(),
        "observed_text": manual_target.strip(),
        "category": manual_category,
        "context": manual_context,
        "accepted": True,
        "decision": "Manually added",
        "notes": manual_note,
    })
if st.session_state.manual_findings:
    manual_display = pd.DataFrame(st.session_state.manual_findings)
    manual_display.insert(0, "remove", False)
    manual_display = st.data_editor(
        manual_display,
        key=f"manual_review_{transcript_key}",
        hide_index=True,
        use_container_width=True,
        disabled=[column for column in manual_display.columns if column != "remove"],
        column_config={
            "remove": st.column_config.CheckboxColumn("Remove"),
            "target": "Target",
            "category": "Category",
            "context": st.column_config.TextColumn("Context", width="large"),
            "notes": st.column_config.TextColumn("Reviewer notes", width="medium"),
        },
        column_order=["remove", "target", "category", "context", "notes"],
    )
    remove_manual = st.button(
        "Remove selected occurrence(s)",
        disabled=not bool(manual_display["remove"].any()),
    )
    if remove_manual:
        remove_indices = set(manual_display.index[manual_display["remove"]])
        st.session_state.manual_findings = [
            row for index, row in enumerate(st.session_state.manual_findings) if index not in remove_indices
        ]
        st.rerun()

all_findings = reviewed_findings + st.session_state.manual_findings
configured_metric_targets = (
    [(target, "Lexical") for target in lexical_targets]
    + [(target, "Nonlexical") for target in nonlexical_targets]
)
metrics = calculate_metrics(
    all_findings,
    len(words),
    duration_seconds,
    configured_targets=configured_metric_targets,
)
accepted_total = sum(bool(row.get("accepted")) for row in all_findings)
total_per_100_words = accepted_total / len(words) * 100 if words else None
total_per_minute = accepted_total / (duration_seconds / 60) if duration_seconds else None

st.header("4. Results")
m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Total lexical words", f"{len(words):,}")
m2.metric("Accepted disfluencies", accepted_total)
m3.metric("Verified duration", f"{duration_seconds / 60:.2f} min" if duration_seconds else "Not available")
m4.metric(
    "Total per 100 words",
    f"{total_per_100_words:.3f}" if total_per_100_words is not None else "Not available",
)
m5.metric(
    "Total per minute",
    f"{total_per_minute:.3f}" if total_per_minute is not None else "Not available",
)
metrics_df = pd.DataFrame(metrics)
if not metrics_df.empty:
    st.dataframe(
        metrics_df.style.format({"Per_100_Lexical_Words": "{:.3f}", "Per_Minute": "{:.3f}"}, na_rep="—"),
        use_container_width=True,
        hide_index=True,
    )

st.header("5. Export")
participant_col, session_col = st.columns(2)
participant_id = participant_col.text_input("Participant ID")
session_information = session_col.text_input(
    "Session information", value=re.sub(r"\.[^.]+$", "", source_name)
)
reviewer_col, role_col = st.columns(2)
reviewer_id = reviewer_col.text_input("Reviewer identifier")
reviewer_role = role_col.selectbox("Reviewer role", ["Primary", "Secondary"])
analysis_date = st.date_input("Analysis date", value=date.today())
transcript_hash = hashlib.sha256(normalize_for_analysis(analysis_text).encode()).hexdigest()
summary = {
    "Participant_ID": participant_id,
    "Session_Information": session_information,
    "Analysis_Date": analysis_date,
    "Reviewer_ID": reviewer_id,
    "Reviewer_Role": reviewer_role,
    "Source_File": source_name,
    "Source_Format": source_format,
    "Selected_Speaker": chosen_speaker,
    "Total_Words_Spoken": len(words),
    "Duration_Minutes": duration_seconds / 60 if duration_seconds else None,
    "Total_Disfluencies": accepted_total,
    "Total_Disfluencies_Per_100_Words": total_per_100_words,
    "Total_Disfluencies_Per_Minute": total_per_minute,
}

accepted_findings = [row for row in all_findings if bool(row.get("accepted"))]
for category in ("Lexical", "Nonlexical"):
    category_count = sum(row.get("category") == category for row in accepted_findings)
    category_prefix = category.replace(" ", "_")
    summary[f"{category_prefix}_Total_Disfluencies"] = category_count
    summary[f"{category_prefix}_Disfluencies_Per_100_Words"] = (
        category_count / len(words) * 100 if words else None
    )
    summary[f"{category_prefix}_Disfluencies_Per_Minute"] = (
        category_count / (duration_seconds / 60) if duration_seconds else None
    )

for target_metric in metrics:
    target_slug = re.sub(r"[^A-Za-z0-9]+", "_", target_metric["Target"]).strip("_") or "Target"
    target_prefix = f"Target_{target_slug}"
    summary[f"{target_prefix}_Category"] = target_metric.get("Category", "")
    summary[f"{target_prefix}_Total_Disfluencies"] = target_metric["Occurrences"]
    summary[f"{target_prefix}_Disfluencies_Per_100_Words"] = target_metric["Per_100_Lexical_Words"]
    summary[f"{target_prefix}_Disfluencies_Per_Minute"] = target_metric["Per_Minute"]

summary["Duration_Source"] = duration_source
summary["Transcript_SHA256"] = transcript_hash
decision_df = pd.DataFrame(all_findings)
workbook = export_workbook(summary, metrics, decision_df, selected_text, analysis_text)
safe_participant = re.sub(r"[^A-Za-z0-9_-]+", "_", participant_id).strip("_") or "Participant"
safe_session = re.sub(r"[^A-Za-z0-9_-]+", "_", session_information).strip("_") or "Session"
st.download_button(
    "Download reviewed Excel record",
    data=workbook,
    file_name=f"TADA_{safe_participant}_{safe_session}_{reviewer_role}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    disabled=not participant_id.strip() or not reviewer_id.strip(),
)
if not participant_id.strip() or not reviewer_id.strip():
    st.caption("Enter participant and reviewer identifiers to enable the audit-ready export.")
