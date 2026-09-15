import math
import unittest

from tada_core import (
    LEXICAL_DEFAULTS,
    NONLEXICAL_DEFAULTS,
    calculate_metrics,
    find_candidates,
    lexical_tokens,
    parse_transcript,
    timestamp_seconds,
)


class TadaCoreTests(unittest.TestCase):
    def test_default_target_sets(self):
        self.assertEqual(
            NONLEXICAL_DEFAULTS,
            ("uh", "um", "er", "ah", "mm-hmm", "erm", "hmm", "eh", "huh"),
        )
        self.assertEqual(LEXICAL_DEFAULTS, ("like", "you know", "so", "therefore", "I mean"))

    def test_timestamp_seconds(self):
        self.assertEqual(timestamp_seconds("00:01:30.500"), 90.5)
        self.assertEqual(timestamp_seconds("01:05,250"), 65.25)


    def test_vtt_parsing_and_duration(self):
        text = """WEBVTT

1
00:00:02.000 --> 00:00:04.000
Speaker One: So we begin.

2
00:00:05.000 --> 00:00:09.500
Speaker One: Um, this continues.
"""
        parsed = parse_transcript(text, "vtt")
        self.assertEqual(parsed.detected_duration_seconds, 7.5)
        self.assertEqual(parsed.speaker_text["Speaker One"], "So we begin. Um, this continues.")
        self.assertNotIn("00:00", parsed.full_text)

    def test_zoom_saved_caption_text_format(self):
        text = """[Speaker] 10:00:09
When you're ready to start.

[Speaker] 10:00:58
So, good morning. Um, today I will begin.

[Speaker] 10:04:09
And stop.
"""
        parsed = parse_transcript(text, "txt")
        self.assertEqual(parsed.detected_duration_seconds, 240.0)
        self.assertEqual(
            parsed.speaker_text["Speaker"],
            "When you're ready to start. So, good morning. Um, today I will begin. And stop.",
        )
        self.assertNotIn("10:00", parsed.full_text)
        self.assertNotIn("00:58", parsed.full_text)


    def test_plain_text_colon_is_not_reliable_speaker(self):
        parsed = parse_transcript("ID: This is content. Another sentence.", "txt")
        self.assertEqual(parsed.speaker_text, {})
        self.assertIn("ID: This is content", parsed.full_text)

    def test_one_off_colon_content_survives_caption_parsing(self):
        text = """WEBVTT

1
00:00:01.000 --> 00:00:02.000
Carole Van Camp: First line.

2
00:00:02.000 --> 00:00:03.000
ID: This is spoken content.

3
00:00:03.000 --> 00:00:04.000
Carole Van Camp: Last line.
"""
        parsed = parse_transcript(text, "vtt")
        self.assertEqual(parsed.speaker_text["Carole Van Camp"], "First line. Last line.")
        self.assertIn("ID: This is spoken content.", parsed.full_text)


    def test_lexical_denominator_excludes_nonlexical_only(self):
        text = "So I, um, think therefore we continue. Uh, so yes."
        tokens = lexical_tokens(text, ["uh", "um"])
        self.assertEqual([t.casefold() for t in tokens], ["so", "i", "think", "therefore", "we", "continue", "so", "yes"])


    def test_candidate_counts_are_literal_and_reviewable(self):
        text = "So I was so happy. Um, therefore, so we left."
        findings = find_candidates(text, ["so", "therefore"], ["um"])
        self.assertEqual([row["target"] for row in findings], ["so", "so", "um", "therefore", "so"])
        self.assertTrue(all(row["accepted"] for row in findings))


    def test_metrics_share_lexical_denominator(self):
        findings = find_candidates("So um therefore so", ["so", "therefore"], ["um"])
        metrics = {row["Target"]: row for row in calculate_metrics(findings, 3, 120)}
        self.assertEqual(metrics["so"]["Occurrences"], 2)
        self.assertEqual(metrics["so"]["Category"], "Lexical")
        self.assertEqual(metrics["um"]["Category"], "Nonlexical")
        self.assertTrue(math.isclose(metrics["so"]["Per_100_Lexical_Words"], 66.6667, rel_tol=1e-5))
        self.assertTrue(math.isclose(metrics["um"]["Per_100_Lexical_Words"], 33.3333, rel_tol=1e-5))
        self.assertEqual(metrics["therefore"]["Per_Minute"], 0.5)

    def test_configured_target_with_no_occurrences_is_reported_as_zero(self):
        metrics = calculate_metrics(
            find_candidates("Um, begin.", ["so"], ["um"]),
            1,
            60,
            configured_targets=[("so", "Lexical"), ("um", "Nonlexical")],
        )
        by_target = {row["Target"]: row for row in metrics}
        self.assertEqual(by_target["so"]["Occurrences"], 0)
        self.assertEqual(by_target["so"]["Per_100_Lexical_Words"], 0)
        self.assertEqual(by_target["so"]["Per_Minute"], 0)


    def test_non_speech_annotations_do_not_count_as_words(self):
        tokens = lexical_tokens("Hello [laughter] world (inaudible) um", ["um"])
        self.assertEqual(tokens, ["Hello", "world"])


if __name__ == "__main__":
    unittest.main()
