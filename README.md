# TADA

**Transcript-Assisted Disfluency Analyzer**

TADA is a browser-based Streamlit application for identifying, reviewing, and quantifying user-defined speech events in transcripts. It accepts pasted text and TXT, DOCX, VTT, and SRT files. TADA presents every automatically identified occurrence for human review.

## Measures

- Occurrences of each configured target
- Occurrences per 100 lexical words
- Occurrences per minute when a verified duration is available
- Total lexical words and an occurrence-level audit record

Lexical target words remain in the lexical-word denominator. Configured nonlexical vocalizations, timestamps, speaker labels, platform metadata, URLs, and non-speech annotations are excluded. See [COUNTING_RULES.md](COUNTING_RULES.md).

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Tests

```bash
python -m unittest discover -s tests
```

## Privacy

TADA analyzes transcript text during the active Streamlit session. The application does not transcribe audio. Users should not upload identifiable or protected information unless the deployment they are using is approved for that information.

## Status

TADA is under active development.
