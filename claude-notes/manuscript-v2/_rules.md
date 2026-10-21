# Drafting rules for every v2 chapter

These apply to both drafting models on every chapter. Each chapter's own brief (`briefs/chNN.md`) adds what that chapter needs.

## Files
- **Never edit, create or delete anything in `manuscript/`.** Read it only.
- Write only the single output file your instructions name.
- **Read before drafting:**
  1. `_rules.md` (this file);
  2. `_canon-facts.md`;
  3. the chapter's brief in `briefs/`;
  4. the source chapters the brief lists (in `manuscript/`);
  5. the **previous merged v2 chapter in the same thread** (named in the brief), so voice and continuity carry on.
- **Do not read** the other model's drafts, anything in `drafts/` other than your own output, or `comparisons/`.

## Format
- `# Chapter N: Title`, then a blank line, then the italic epigraph line exactly as given in the brief, then `— Book 1, Chapter X: Title`.
- Then a one-line heading: `## Then: <place or time>` or `## Now: <place>`.
- Scene breaks are a line with `---`.
- Plain Markdown otherwise. Use `*italic*` for thoughts and emphasis. ALL-CAPS for system messages, three at most per chapter.
- Pronunciation lives only on the front page (00.1-names-and-pronunciation-merged.md). Never put bracketed guides in chapter text. A new non-English word gets added to claude-notes/pronunciation/entries.js and to the front page.

## Voice
- **The author's voice:** close third, past tense, precise, image-led, measured, with an occasional dry line. The existing `manuscript/` chapters are the reference for the tower thread. Book 1 is the reference for Camilla, Gage, Ben and Otto: grittier, more sensory, wry.
- **Show emotion through the body or action.** Never name it.
- Vary sentence length.
- **At least one specific, slightly odd physical detail per scene.**
- **Under-25 characters use contractions in dialogue.** Grandmasters, elders, Mr. Flitzberg and the medics stay formal; the contrast is deliberate.
- **Banned:**
  - "Not X. Y." / "It was not X. It was Y." / "X wasn't A; it was B" constructions;
  - "For the first time", "For a moment", "In that moment", "a moment";
  - "as though" (use "as if" at most once per chapter);
  - "read it twice" and its variants;
  - paragraph-ending aphorisms that explain what the paragraph just showed;
  - restating events the reader has already seen;
  - recaps at the start of a chapter;
  - "simply", "merely" (except Meru's "merely"), "precisely", "quietly", "carefully" as filler.
- **Exposition:** at most one paragraph of mechanism at a time, ideally delivered in dialogue under pressure. Cut lectures into exchanges.
- **No romance anywhere for Elara and Rohan.** Comrades, close to siblings in arms. No lingering on his looks.

## Reuse
- **Keep the author's strongest existing lines** where the brief says to, verbatim or nearly. Compress everything else.
- **Hit the brief's target length within ±10%.**

## When done
Reply with:
- the file path and word count;
- under 120 words on your main choices;
- any canon question you weren't sure about.
