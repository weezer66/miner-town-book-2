from pathlib import Path


MANUSCRIPT_DIR = Path("manuscript")
CHAPTERS = [
    MANUSCRIPT_DIR / "chapter-01-assigned-to-witness.md",
    MANUSCRIPT_DIR / "chapter-02-the-checklist.md",
    MANUSCRIPT_DIR / "chapter-03-the-first-correction.md",
    MANUSCRIPT_DIR / "chapter-04-the-first-report.md",
]

REQUIRED_MARKERS = [
    (0, "## Day Zero: Assignment"),
    (1, "## Day One: Intake"),
    (1, "AUTOPILOT OBSERVATION ENGAGED."),
    (2, "## Day One, continued: The Pattern Forms"),
    (2, "thirty-six hours"),
    (2, "forty-eight hours"),
    (3, "## Day Three: Threshold"),
    (3, "sixty hours"),
    (3, "seventy-two hours"),
    (3, "PASSED: INITIATION SURVIVAL CRITERIA."),
]


def main() -> None:
    chapter_texts = [path.read_text(encoding="utf-8") for path in CHAPTERS]
    marker_positions = []

    for chapter_index, marker in REQUIRED_MARKERS:
        position = chapter_texts[chapter_index].find(marker)
        if position == -1:
            raise SystemExit(f"Missing Initiation timeline marker: {marker}")
        marker_positions.append((chapter_index, position))

    if marker_positions != sorted(marker_positions):
        raise SystemExit("Initiation timeline markers are out of order")

    print("Verified Book 2 Initiation timeline: Day Zero through 72-hour survival result")


if __name__ == "__main__":
    main()