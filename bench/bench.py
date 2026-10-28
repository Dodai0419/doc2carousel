"""Time doc2carousel against its two upstreams on the same input.

    python bench/bench.py                          # this repo only
    python bench/bench.py path/to/ian-xhs-carousel # plus upstream; needs `pip install "markitdown[docx,pptx]" playwright pillow`

Upstream render.py launches Playwright's own Chromium; here it is pointed at the installed Chrome (channel="chrome"),
so no browser download is needed and both sides use the same browser.
"""
import runpy
import shutil
import statistics
import subprocess
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import doc2carousel as dc  # noqa: E402

FIX = ROOT / "tests" / "fixtures"
FILES = [FIX / "test.docx", FIX / "test.pptx"]


def med(fn, n):
    runs = []
    for _ in range(n):
        t = time.perf_counter()
        fn()
        runs.append((time.perf_counter() - t) * 1000)
    return statistics.median(runs)


def import_ms(stmt):
    runs = []
    for _ in range(5):
        r = subprocess.run([sys.executable, "-c", f"import sys,time;sys.path.insert(0,{str(ROOT)!r});t=time.perf_counter();{stmt};"
                            "print((time.perf_counter()-t)*1000)"], capture_output=True, text=True, check=True)
        runs.append(float(r.stdout))
    return statistics.median(runs)


rows = []
up = len(sys.argv) > 1
if up:
    from markitdown import MarkItDown
    mid = MarkItDown()
for f in FILES:
    ours = med(lambda: dc.to_markdown(f), 20)
    theirs = med(lambda: mid.convert(str(f)), 20) if up else float("nan")
    rows.append((f"convert {f.name}", theirs, ours))
rows.append(("import", import_ms("import markitdown") if up else float("nan"), import_ms("import doc2carousel")))

md = (ROOT / "examples" / "weekend.md").read_text(encoding="utf-8")
slides = dc.plan(md, "weekend.md")
chrome = dc.find_chrome()
with tempfile.TemporaryDirectory() as d:
    out = Path(d)
    page = out / "carousel.html"
    page.write_text(dc.render_html(slides), encoding="utf-8")
    t = time.perf_counter()
    dc.export_pngs(page, len(slides), out, chrome)
    ours = (time.perf_counter() - t) * 1000
    theirs = float("nan")
    if up:
        from playwright.sync_api import BrowserType
        launch = BrowserType.launch
        BrowserType.launch = lambda self, **kw: launch(self, channel="chrome", **kw)
        up_dir = out / "upstream"
        up_dir.mkdir()
        shutil.copy(page, up_dir / "index.html")
        script = Path(sys.argv[1]) / "ian-xhs-carousel" / "scripts" / "render.py"
        sys.argv = ["render.py", str(up_dir / "index.html"), str(up_dir)]
        t = time.perf_counter()
        runpy.run_path(str(script), run_name="__main__")
        theirs = (time.perf_counter() - t) * 1000
        assert len(list(up_dir.glob("[0-9][0-9].png"))) == len(slides)
    rows.append((f"render {len(slides)} slides to PNG", theirs, ours))

print(f"{'operation':30} {'upstream ms':>12} {'this repo ms':>13}")
for name, a, b in rows:
    print(f"{name:30} {a:12.1f} {b:13.1f}")
