"""Core parsing and measurement functions for TADA."""

from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
import json
import re
from typing import Iterable

from docx import Document


NONLEXICAL_DEFAULTS = ("uh", "um", "er", "ah", "mm-hmm", "erm", "hmm", "eh", "huh")
LEXICAL_DEFAULTS = ("like", "you know", "so", "therefore", "I mean")

TIME_TOKEN = r"(?:\d{1,2}:)?\d{1,2}:\d{2}(?:[.,]\d{1,3})?"
CUE_RE = re.compile(rf"^\s*({TIME_TOKEN})\s*-->\s*({TIME_TOKEN})(?:\s+.*)?$", re.MULTILINE)
BRACKETED_TIMESTAMP_RE = re.compile(rf"^\s*\[([^\]]+)\]\s+({TIME_TOKEN})\s*$", re.MULTILINE)
SPEAKER_RE = re.compile(r"^\s*([^:\n]{1,80}):\s*(.*)$")
WORD_RE = re.compile(r"[A-Za-z0-9]+(?:[’'][A-Za-z0-9]+)*(?:-[A-Za-z0-9]+)*")
NON_SPEECH_RE = re.compile(r"\[(?:[^\]]+)\]|\((?:inaudible|unintelligible|laughter|music|silence)\)", re.I)


@dataclass(frozen=True)
class TranscriptParse:
    source_format: str
    full_text: str
    speaker_text: dict[str, str]
    detected_duration_seconds: float | None
    first_timestamp_seconds: float | None
    last_timestamp_seconds: float | None
    source_word_timings: tuple[dict, ...] = ()


def timestamp_seconds(value: str) -> float:
    clean = value.strip().replace(",", ".")
    parts = clean.split(":")
    if len(parts) == 3:
        hours, minutes, seconds = parts
    elif len(parts) == 2:
        hours, minutes, seconds = "0", parts[0], parts[1]
    else:
        raise ValueError(f"Unsupported timestamp: {value}")
    return int(hours) * 3600 + int(minutes) * 60 + float(seconds)


def _decode_upload(data: bytes, filename: str) -> str:
    suffix = Path(filename).suffix.lower()
    if suffix == ".docx":
        document = Document(BytesIO(data))
        return "\n".join(p.text for p in document.paragraphs)
    if suffix == ".json":
        return parse_elevenlabs_json_upload(data, filename).full_text
    if suffix not in {".txt", ".vtt", ".srt"}:
        raise ValueError("Supported files are .txt, .docx, .vtt, .srt, and ElevenLabs .json.")
    for encoding in ("utf-8-sig", "utf-8", "cp1252"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise ValueError("The transcript text encoding could not be read.")


def read_transcript_upload(data: bytes, filename: str) -> str:
    return _decode_upload(data, filename)


def parse_elevenlabs_json_upload(data: bytes, filename: str = "transcript.json") -> TranscriptParse:
    """Parse supported JSON transcript exports.

    ElevenLabs Speech-to-Text exports normally contain a top-level text field
    plus word-level timing records. TADA also accepts common timestamped JSON
    transcript structures that contain segment records instead of a top-level
    text field. When necessary, transcript text is reconstructed from available
    word or segment records.
    """
    payload = None
    for encoding in ("utf-8-sig", "utf-8"):
        try:
            payload = json.loads(data.decode(encoding))
            break
        except UnicodeDecodeError:
            continue
        except json.JSONDecodeError as exc:
            raise ValueError("The JSON file could not be parsed.") from exc

    timings: list[dict] = []
    starts: list[float] = []
    ends: list[float] = []
    audio_duration = None

    def add_timing(text_value, start_value, end_value, item_type="", logprob=None):
        start_num = float(start_value) if isinstance(start_value, (int, float)) else None
        end_num = float(end_value) if isinstance(end_value, (int, float)) else None
        if start_num is not None:
            starts.append(start_num)
        if end_num is not None:
            ends.append(end_num)
        timings.append({
            "text": text_value if isinstance(text_value, str) else "",
            "start": start_num,
            "end": end_num,
            "type": item_type,
            "logprob": logprob,
        })

    full_text = ""

    if isinstance(payload, dict):
        top_text = payload.get("text", "")
        if isinstance(top_text, str) and top_text.strip():
            full_text = top_text.strip()

        raw_words = payload.get("words", [])
        if not isinstance(raw_words, list):
            raw_words = []

        for item in raw_words:
            if not isinstance(item, dict):
                continue
            add_timing(
                item.get("text", ""),
                item.get("start"),
                item.get("end"),
                item.get("type", ""),
                item.get("logprob"),
            )

        if not full_text and raw_words:
            pieces = [item.get("text", "") for item in raw_words if isinstance(item, dict)]
            pieces = [piece for piece in pieces if isinstance(piece, str) and piece]
            if any(piece[:1].isspace() or piece[-1:].isspace() for piece in pieces):
                full_text = "".join(pieces).strip()
            else:
                word_pieces = [
                    item.get("text", "")
                    for item in raw_words
                    if isinstance(item, dict)
                    and isinstance(item.get("text"), str)
                    and item.get("text")
                    and item.get("type", "word") not in {"spacing", "audio_event"}
                ]
                full_text = " ".join(word_pieces).strip()

        raw_segments = payload.get("segments", [])
        if not isinstance(raw_segments, list):
            raw_segments = []
        if raw_segments and not raw_words:
            for item in raw_segments:
                if not isinstance(item, dict):
                    continue
                timestamp = item.get("timestamp")
                start_value = item.get("start")
                end_value = item.get("end")
                if isinstance(timestamp, (list, tuple)) and len(timestamp) >= 2:
                    start_value, end_value = timestamp[0], timestamp[1]
                add_timing(item.get("text", ""), start_value, end_value, "segment")
        if not full_text and raw_segments:
            full_text = " ".join(
                str(item.get("text", "")).strip()
                for item in raw_segments
                if isinstance(item, dict) and str(item.get("text", "")).strip()
            ).strip()

        audio_duration = payload.get("audio_duration_secs")
        if audio_duration is None:
            audio_duration = payload.get("audio_duration_seconds")

    elif isinstance(payload, list):
        segments = [item for item in payload if isinstance(item, dict)]
        for item in segments:
            timestamp = item.get("timestamp")
            start_value = item.get("start")
            end_value = item.get("end")
            if isinstance(timestamp, (list, tuple)) and len(timestamp) >= 2:
                start_value, end_value = timestamp[0], timestamp[1]
            add_timing(item.get("text", ""), start_value, end_value, "segment")
        full_text = " ".join(
            str(item.get("text", "")).strip()
            for item in segments
            if str(item.get("text", "")).strip()
        ).strip()
    else:
        raise ValueError("The JSON file must contain a transcript object or a list of transcript segments.")

    if not full_text:
        raise ValueError(
            "The JSON file does not contain usable transcript text. "
            "TADA supports ElevenLabs JSON exports and timestamped JSON records containing text."
        )

    if isinstance(audio_duration, (int, float)) and audio_duration > 0:
        duration = float(audio_duration)
        first = 0.0
        last = duration
    elif ends:
        first = min(starts) if starts else 0.0
        last = max(ends)
        duration = last - first if last >= first else None
    else:
        first = last = duration = None

    return TranscriptParse(
        source_format="json",
        full_text=full_text,
        speaker_text={},
        detected_duration_seconds=duration,
        first_timestamp_seconds=first,
        last_timestamp_seconds=last,
        source_word_timings=tuple(timings),
    )


def _is_cue_number(line: str) -> bool:
    return bool(re.fullmatch(r"\s*\d+\s*", line))


def parse_transcript(text: str, source_format: str = "pasted text") -> TranscriptParse:
    """Parse plain text, VTT, or SRT while preserving only spoken content."""
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    cue_matches = list(CUE_RE.finditer(normalized))
    bracketed_matches = list(BRACKETED_TIMESTAMP_RE.finditer(normalized))
    captioned = bool(cue_matches or bracketed_matches) or source_format.lower() in {"vtt", "srt", ".vtt", ".srt"}
    if cue_matches:
        first = timestamp_seconds(cue_matches[0].group(1))
        last = timestamp_seconds(cue_matches[-1].group(2))
    elif bracketed_matches:
        first = timestamp_seconds(bracketed_matches[0].group(2))
        last = timestamp_seconds(bracketed_matches[-1].group(2))
    else:
        first = last = None
    duration = last - first if first is not None and last is not None and last >= first else None

    speakers: dict[str, list[str]] = {}
    parsed_lines: list[tuple[str | None, str, str]] = []
    current_bracketed_speaker: str | None = None
    for raw_line in normalized.splitlines():
        line = raw_line.strip().lstrip("\ufeff")
        bracketed_timestamp = BRACKETED_TIMESTAMP_RE.fullmatch(line)
        if bracketed_timestamp:
            current_bracketed_speaker = bracketed_timestamp.group(1).strip()
            continue
        if not line or line.upper() == "WEBVTT" or CUE_RE.fullmatch(line) or (captioned and _is_cue_number(line)):
            continue
        # Common VTT metadata lines.
        if line.startswith(("NOTE", "STYLE", "REGION", "Kind:", "Language:")):
            continue
        line = re.sub(r"<[^>]+>", "", line).strip()
        match = SPEAKER_RE.match(line)
        if match:
            speaker, spoken = match.groups()
            speakers.setdefault(speaker.strip(), []).append(spoken.strip())
            parsed_lines.append((speaker.strip(), spoken.strip(), line))
        elif current_bracketed_speaker:
            speakers.setdefault(current_bracketed_speaker, []).append(line)
            parsed_lines.append((current_bracketed_speaker, line, line))
        else:
            parsed_lines.append((None, line, line))

    # In plain prose, one-off colon prefixes are likely content, not speaker labels.
    # Repetition is the safest general signal of a speaker label. For a caption
    # file containing only one content line, accept its single prefix as a label.
    one_line_caption = captioned and len(parsed_lines) == 1
    reliable = {
        name: " ".join(parts)
        for name, parts in speakers.items()
        if len(parts) >= 2 or one_line_caption
    }
    reliable_names = set(reliable)
    all_lines = [
        spoken if speaker in reliable_names else original
        for speaker, spoken, original in parsed_lines
    ]
    full_text = " ".join(part for part in all_lines if part)
    full_text = re.sub(r"\s+", " ", full_text).strip()
    return TranscriptParse(
        source_format=source_format,
        full_text=full_text,
        speaker_text=reliable,
        detected_duration_seconds=duration,
        first_timestamp_seconds=first,
        last_timestamp_seconds=last,
    )


def normalize_for_analysis(text: str) -> str:
    text = NON_SPEECH_RE.sub(" ", text)
    text = re.sub(r"https?://\S+|www\.\S+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def lexical_tokens(text: str, nonlexical_targets: Iterable[str]) -> list[str]:
    excluded = {item.casefold() for item in nonlexical_targets}
    return [token for token in WORD_RE.findall(normalize_for_analysis(text)) if token.casefold() not in excluded]


def find_candidates(
    text: str,
    lexical_targets: Iterable[str],
    nonlexical_targets: Iterable[str],
) -> list[dict]:
    """Return every literal target occurrence for human review."""
    cleaned = normalize_for_analysis(text)
    configured = [(x.strip(), "Lexical") for x in lexical_targets if x.strip()]
    configured += [(x.strip(), "Nonlexical") for x in nonlexical_targets if x.strip()]
    configured.sort(key=lambda item: len(item[0]), reverse=True)
    occupied: list[tuple[int, int]] = []
    findings: list[dict] = []
    for target, category in configured:
        pattern = re.compile(rf"(?<![\w'-]){re.escape(target)}(?![\w'-])", re.I)
        for match in pattern.finditer(cleaned):
            if any(match.start() < end and match.end() > start for start, end in occupied):
                continue
            occupied.append((match.start(), match.end()))
            before = cleaned[: match.start()].split()
            after = cleaned[match.end() :].split()
            context = " ".join(before[-7:] + [cleaned[match.start():match.end()]] + after[:7])
            findings.append({
                "start": match.start(),
                "end": match.end(),
                "target": target.casefold(),
                "observed_text": cleaned[match.start():match.end()],
                "category": category,
                "context": context,
                "accepted": True,
                "decision": "Accepted",
                "notes": "",
            })
    return sorted(findings, key=lambda row: (row["start"], row["end"]))


def calculate_metrics(
    findings: list[dict],
    total_lexical_words: int,
    duration_seconds: float | None,
    configured_targets: Iterable[tuple[str, str]] | None = None,
) -> list[dict]:
    accepted = [row for row in findings if bool(row.get("accepted", False))]
    targets = sorted({(row["target"], row.get("category", "")) for row in findings})
    if configured_targets:
        targets = sorted(set(targets) | {
            (target.casefold(), category) for target, category in configured_targets
        })
    rows = []
    for target, category in targets:
        count = sum(
            row["target"] == target and row.get("category", "") == category
            for row in accepted
        )
        rows.append({
            "Target": target,
            "Category": category,
            "Occurrences": count,
            "Per_100_Lexical_Words": (count / total_lexical_words * 100) if total_lexical_words else None,
            "Per_Minute": (count / (duration_seconds / 60)) if duration_seconds and duration_seconds > 0 else None,
        })
    return rows
