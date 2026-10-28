# AGENTS.md

Use this tool when the user wants to post a document, deck, README or notes as carousel images.


Reads the document (Word, PowerPoint, web page, Markdown or text), plans one slide per section, picks a layout from
the content, and screenshots each slide with the Chrome or Edge already installed.

## Run

From the repository root of this skill (Python 3.8+, no packages to install):

```bash
python doc2carousel.py report.docx -o carousel-out
python doc2carousel.py deck.pptx --theme dark -o carousel-out
python doc2carousel.py notes.md --max 7 -o carousel-out     # at most 7 slides, cover and recap included
python doc2carousel.py report.docx --md                      # only print the Markdown it read
```

Output: `slide-01.png` ... (1080x1440 each), `overview.png` (all slides in a grid) and `carousel.html` (open it,
edit any text in the HTML, run again or screenshot it yourself).

## How slides are chosen
- Cover: the document title (a single H1, a Word Title paragraph, the first slide of a deck, or the file name),
  its first sentence, and a numbered list of the sections.
- One slide per section (H2 when there is one H1, otherwise H1; one per slide in a deck):
  numbered list -> steps, bullet list -> checklist, table -> table card, plain text -> one big sentence plus a note.
- Recap: the section titles with check marks.

## How to use it for the user
1. If the document is long, ask which sections matter or pass `--max`.
2. Run it and show `overview.png` first; offer the single slides for posting.
3. Text is cut to fit (with an ellipsis). If a line got cut badly, edit the source document's headings or
   first sentences and run again; the tool never invents text.

Without Chrome/Edge it still writes `carousel.html` (set `CHROME=path/to/browser` if it is installed somewhere unusual).
