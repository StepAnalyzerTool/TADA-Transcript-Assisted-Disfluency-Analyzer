"""Streamlit interface for the Disfluency Analysis and Review Tool (DART)."""

from __future__ import annotations

from datetime import date
from io import BytesIO
import html
import re

import pandas as pd
import streamlit as st

from dart_core import (
    LEXICAL_DEFAULTS,
    NONLEXICAL_DEFAULTS,
    calculate_metrics,
    find_candidates,
    lexical_tokens,
    normalize_for_analysis,
    parse_transcript,
    read_transcript_upload,
)


st.set_page_config(page_title="DART", page_icon="🎯", layout="wide")
st.title("DART")
st.caption("Disfluency Analysis and Review Tool")
st.info("DART identifies candidate speech events for human review.")


def split_targets(value: str) -> list[str]:
    return list(dict.fromkeys(item.strip().casefold() for item in value.split(",") if item.strip()))


def highlighted_transcript(text: str, findings: list[dict]) -> str:
    """Render detected candidates safely with category-specific highlighting."""
    cleaned = normalize_for_analysis(text)
    pieces = []
    cursor = 0
    colors = {"Lexical": "#dbeafe", "Nonlexical": "#fef3c7"}
    for finding in findings:
        start, end = finding["start"], finding["end"]
        pieces.append(html.escape(cleaned[cursor:start]))
        observed = html.escape(cleaned[start:end])
        color = colors.get(finding["category"], "#e5e7eb")
        pieces.append(
            f'<mark style="background-color:{color};padding:0.05rem 0.18rem;'
            f'border-radius:0.2rem;font-weight:700">{observed}</mark>'
        )
        cursor = end
    pieces.append(html.escape(cleaned[cursor:]))
    return "".join(pieces).replace("\n", "<br>")


def export_workbook(summary: dict, metrics: list[dict], decisions: pd.DataFrame, cleaned_text: str) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        pd.DataFrame([summary]).to_excel(writer, sheet_name="Session_Summary", index=False)
        pd.DataFrame(metrics).to_excel(writer, sheet_name="Target_Metrics", index=False)
        decisions.to_excel(writer, sheet_name="Occurrence_Review", index=False)
        pd.DataFrame([{"Reviewed_Transcript": cleaned_text}]).to_excel(
            writer, sheet_name="Reviewed_Transcript", index=False
        )
    return buffer.getvalue()


with st.sidebar:
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
analysis_text = parsed.full_text
if parsed.speaker_text:
    options = ["All detected speech"] + sorted(parsed.speaker_text)
    chosen_speaker = st.selectbox("Speaker", options)
    if chosen_speaker != "All detected speech":
        analysis_text = parsed.speaker_text[chosen_speaker]
else:
    chosen_speaker = "Not available"
    st.caption("No reliable speaker labels were detected.")

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
    duration_minutes = st.number_input("Session duration (minutes)", min_value=0.01, value=1.00, step=0.01)
    duration_seconds = duration_minutes * 60
else:
    duration_minutes = None
    duration_seconds = None

lexical_targets = split_targets(lexical_value)
nonlexical_targets = split_targets(nonlexical_value)
words = lexical_tokens(analysis_text, nonlexical_targets)
findings = find_candidates(analysis_text, lexical_targets, nonlexical_targets)

st.subheader("Visual audit")
st.caption("Blue = lexical candidate · Gold = nonlexical candidate")
with st.container(border=True):
    st.markdown(highlighted_transcript(analysis_text, findings), unsafe_allow_html=True)

st.header("3. Review candidate occurrences")
st.caption("Reject grammatical or incorrectly transcribed instances and add a note when useful.")
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
        disabled=["target", "category", "observed_text", "context"],
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
    st.dataframe(pd.DataFrame(st.session_state.manual_findings), hide_index=True, use_container_width=True)

all_findings = reviewed_findings + st.session_state.manual_findings
metrics = calculate_metrics(all_findings, len(words), duration_seconds)

st.header("4. Results")
m1, m2, m3 = st.columns(3)
m1.metric("Total lexical words", f"{len(words):,}")
m2.metric("Accepted target occurrences", sum(bool(row.get("accepted")) for row in all_findings))
m3.metric("Verified duration", f"{duration_seconds / 60:.2f} min" if duration_seconds else "Not available")
metrics_df = pd.DataFrame(metrics)
if not metrics_df.empty:
    st.dataframe(
        metrics_df.style.format({"Per_100_Lexical_Words": "{:.3f}", "Per_Minute": "{:.3f}"}, na_rep="—"),
        use_container_width=True,
        hide_index=True,
    )

st.header("5. Export")
session_id = st.text_input("Session identifier", value=re.sub(r"\.[^.]+$", "", source_name))
reviewer_id = st.text_input("Reviewer identifier")
session_date = st.date_input("Session date", value=date.today())
summary = {
    "Session_ID": session_id,
    "Session_Date": session_date,
    "Reviewer_ID": reviewer_id,
    "Source_File": source_name,
    "Source_Format": source_format,
    "Selected_Speaker": chosen_speaker,
    "Duration_Minutes": duration_seconds / 60 if duration_seconds else None,
    "Duration_Source": duration_source,
    "Total_Lexical_Words": len(words),
    "Accepted_Target_Occurrences": sum(bool(row.get("accepted")) for row in all_findings),
}
decision_df = pd.DataFrame(all_findings)
workbook = export_workbook(summary, metrics, decision_df, analysis_text)
safe_session = re.sub(r"[^A-Za-z0-9_-]+", "_", session_id).strip("_") or "Session"
st.download_button(
    "Download reviewed Excel record",
    data=workbook,
    file_name=f"DART_{safe_session}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    disabled=not reviewer_id.strip(),
)
if not reviewer_id.strip():
    st.caption("Enter a reviewer identifier to enable the audit-ready export.")
