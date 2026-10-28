import os
import re
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import doc2carousel as dc  # noqa: E402

FIX = ROOT / "tests" / "fixtures"


def text_lines(md):
    """the words of a Markdown document, one entry per line, markup and images removed"""
    out = []
    for line in md.splitlines():
        if line.startswith(("![", "<!--")) or set(line.strip()) <= set("|-: "):
            continue
        line = re.sub(r"^#+\s*|^\s*[-*]\s+|^\s*\d+\.\s+|\*\*", "", line)
        line = " ".join(c.strip() for c in line.strip().strip("|").split("|") if c.strip()) if "|" in line else line
        line = " ".join(line.split())
        if line:
            out.append(line)
    return out


class Convert(unittest.TestCase):
    """Every line of text markitdown gets out of its own test files, we get too, in the same order."""

    def check(self, name):
        want = text_lines((FIX / f"markitdown_{name}.md").read_text(encoding="utf-8"))
        got = text_lines(dc.to_markdown(FIX / name))
        it = iter(got)
        missing = [w for w in want if not any(w == g for g in it)]
        self.assertEqual(missing, [], f"{len(missing)} of {len(want)} markitdown lines missing or out of order")

    def test_docx_text_matches_markitdown(self):
        self.check("test.docx")

    def test_pptx_text_matches_markitdown(self):
        self.check("test.pptx")

    def test_docx_equations_match_markitdown(self):
        self.check("equations.docx")

    def test_docx_headings_keep_their_level(self):
        md = dc.to_markdown(FIX / "test.docx")
        self.assertIn("# Abstract", md.splitlines())
        self.assertIn("## d666f1f7-46cb-42bd-9a39-9a39cf2a509f", md.splitlines())

    def test_pptx_table_is_a_markdown_table(self):
        md = dc.to_markdown(FIX / "test.pptx")
        self.assertIn("| ColA | ColB | ColC | ColD | ColE | ColF |", md)

    def test_html(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d, "a.html")
            p.write_text("<h1>Hi</h1><p>One <b>two</b>.</p><ul><li>x</li><li>y</li></ul><script>no()</script>", encoding="utf-8")
            self.assertEqual(dc.to_markdown(p).split(), "# Hi One two. - x - y".split())


class Plan(unittest.TestCase):
    MD = """# Ship a side project in a weekend

A short plan that fits in two days.

## Friday night
1. Pick one problem
2. Write the README first
3. Sketch the smallest demo

## Saturday
- Build the core loop
- Write one test
- Deploy it somewhere free

## Why it works
Small scope beats big plans. You finish because the finish line is close. Momentum does the rest.

## Numbers
| Day | Hours |
| --- | --- |
| Friday | 3 |
| Saturday | 8 |
"""

    def test_layouts_follow_the_content(self):
        slides = dc.plan(self.MD, "plan.md")
        self.assertEqual([s["layout"] for s in slides], ["cover", "steps", "checklist", "statement", "table", "end"])
        self.assertEqual(slides[0]["title"], "Ship a side project in a weekend")
        self.assertEqual(slides[1]["items"][0][0], "Pick one problem")
        self.assertEqual(slides[3]["big"], "Small scope beats big plans.")
        self.assertEqual(slides[4]["rows"][2], ["Saturday", "8"])
        self.assertEqual(slides[-1]["items"], ["Friday night", "Saturday", "Why it works", "Numbers"])

    def test_max_slides_and_fit(self):
        md = "# T\n" + "".join(f"## Part {i}\n" + "word " * 80 + ".\n" for i in range(20))
        slides = dc.plan(md, "x.md", max_slides=6)
        self.assertEqual(len(slides), 6)
        self.assertTrue(all(len(s.get("big", "")) <= 90 for s in slides))

    def test_code_blocks_links_and_badges_are_cleaned(self):
        md = """# Tool

[![CI](https://x/b.svg)](https://x) Use [the docs](https://d) and `run`.

## Install
```bash
# not a heading
pip install x
```
One line.
"""
        slides = dc.plan(md, "r.md")
        self.assertEqual([s["layout"] for s in slides], ["cover", "statement", "end"])
        self.assertEqual(slides[0]["sub"], "Use the docs and run.")
        self.assertEqual(slides[1]["big"], "One line.")

    def test_deck_title_slide_becomes_the_cover(self):
        slides = dc.plan(dc.to_markdown(FIX / "test.pptx"), "test.pptx")
        self.assertTrue(slides[0]["title"].startswith("AutoGen: Enabling"))
        self.assertNotIn("AutoGen: Enabling Next-Gen LLM…", slides[0]["items"])

    def test_html_is_escaped(self):
        html = dc.render_html(dc.plan("# <script>x</script>\n\nhi", "x.md"))
        self.assertNotIn("<script>x", html)
        self.assertIn("&lt;script&gt;x", html)


class Cli(unittest.TestCase):
    def test_run_writes_html_and_pngs(self):
        if not dc.find_chrome():
            self.skipTest("no Chrome/Edge on this machine")
        with tempfile.TemporaryDirectory() as d:
            env = {k: v for k, v in os.environ.items() if k not in ("PYTHONUTF8", "PYTHONIOENCODING")}
            r = subprocess.run([sys.executable, str(ROOT / "doc2carousel.py"), str(ROOT / "examples" / "weekend.md"), "-o", d],
                               capture_output=True, env=env, timeout=300)
            self.assertEqual(r.returncode, 0, r.stderr.decode("utf-8", "replace"))
            pngs = sorted(Path(d).glob("slide-*.png"))
            self.assertGreaterEqual(len(pngs), 4)
            self.assertEqual(pngs[0].read_bytes()[16:24], (1080).to_bytes(4, "big") + (1440).to_bytes(4, "big"))
            self.assertTrue(Path(d, "overview.png").exists() and Path(d, "carousel.html").exists())


if __name__ == "__main__":
    unittest.main()
