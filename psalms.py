#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.10"
# dependencies = ["pychrome", "reportlab"]
# ///
"""Build a one-sheet, fold-into-8 PDF of the daily Responsorial Psalms (with refrains)
plus the next Sunday's Gospel, from bible.usccb.org.

Seven days of psalms fill seven panels of a landscape 4x2 grid (edge to edge, no margins
or guide lines); the eighth (bottom right by default) holds the Gospel for the Sunday
after the week. Pages are read by driving a real Chrome window over the DevTools
protocol; Chrome runs with a throwaway profile in /tmp and is closed afterwards.

Usage: uv run psalms.py [--start YYYY-MM-DD] [--out psalms.pdf] [--gospel 7]
                        [--rotate-top] [--paper letter|a4]
--start       first day of the week (default: this week's Sunday)
--gospel      grid index (0-7, left-to-right, top-to-bottom) of the Gospel panel
--rotate-top  turn the top row 180 degrees (classic one-sheet mini-zine fold)
"""
import argparse
import html
import os
import re
import socket
import subprocess
import sys
import tempfile
import threading
import time
import urllib.request
from datetime import date, timedelta

import pychrome
from reportlab.lib.colors import Color
from reportlab.lib.pagesizes import A4, LETTER, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.platypus import Frame, Paragraph

CHROME = {
    "darwin": "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "linux": "google-chrome",
    "win32": r"C:\Program Files\Google\Chrome\Application\chrome.exe",
}.get(sys.platform, "google-chrome")
URL = "https://bible.usccb.org/bible/readings/{:%m%d%y}.cfm"

COLS, ROWS = 4, 2
PAD = 0.15 * inch
MAX_PT, MIN_PT = 10.0, 4.5


# --- fetching -------------------------------------------------------------------------

class Browser:
    """Chrome with remote debugging on a free port and a temp profile; one tab."""

    def __enter__(self):
        with socket.socket() as s:
            s.bind(("127.0.0.1", 0))
            port = s.getsockname()[1]
        self.profile = tempfile.TemporaryDirectory(prefix="psalms-chrome-", ignore_cleanup_errors=True)
        self.proc = subprocess.Popen(
            [CHROME, f"--remote-debugging-port={port}", f"--user-data-dir={self.profile.name}",
             "--no-first-run", "--no-default-browser-check",
             *(["--no-sandbox"] if os.environ.get("PSALMS_NO_SANDBOX") else []), "about:blank"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        endpoint = f"http://127.0.0.1:{port}"
        for _ in range(60):
            try:
                urllib.request.urlopen(f"{endpoint}/json/version", timeout=1)
                break
            except OSError:
                time.sleep(0.5)
        else:
            self.__exit__()
            raise SystemExit(f"Chrome did not start (looked for {CHROME})")
        self.browser = pychrome.Browser(url=endpoint)
        self.tab = self.browser.new_tab()
        self.tab.start()
        self.tab.Page.enable()
        return self

    def __exit__(self, *exc):
        threading.excepthook = lambda args: None  # pychrome's reader thread complains when Chrome goes away
        try:
            self.tab.stop()
        except Exception:
            pass
        self.proc.terminate()
        self.proc.wait()
        self.profile.cleanup()

    def page(self, d: date, timeout=60) -> str:
        self.tab.Page.navigate(url=URL.format(d))
        end = time.time() + timeout
        while time.time() < end:  # the site may run a short browser check before showing the page
            time.sleep(1)
            r = self.tab.Runtime.evaluate(expression="document.documentElement.outerHTML")
            page = r.get("result", {}).get("value", "")
            if "Responsorial Psalm" in page:
                return page
        raise SystemExit(f"{d}: page never showed the readings (check the Chrome window)")


def clean(s: str) -> str:
    return " ".join(html.unescape(re.sub(r"<[^>]+>", "", s)).replace("\xa0", " ").split())


def section(page: str, name: str, d: date):
    m = re.search(
        r'<h3 class="name">\s*' + name + r'\s*</h3>.*?<div class="address">(.*?)</div>.*?'
        r'<div class="content-body">(.*?)</div>',
        page, re.S)
    if not m:
        raise SystemExit(f"{d}: could not parse the {name} block")
    return m


def parse_psalm(page: str, d: date) -> dict:
    m = section(page, "Responsorial Psalm", d)
    blocks, verse = [], []
    skip_alt = False

    def flush():
        if verse:
            blocks.append({"type": "verse", "lines": verse[:]})
            verse.clear()

    for chunk in re.split(r"<br\s*/?>|</p>\s*<p>", m.group(2)):
        line = clean(chunk)
        if not line:
            continue
        if line.lower().rstrip(":") == "or":  # alternate refrain follows; keep only the first one
            skip_alt = True
        elif re.match(r"R\.", line):
            if skip_alt:
                skip_alt = False
                continue
            flush()
            text = re.sub(r"^R\.\s*(\([^)]*\))?\s*", "", line)
            if text:
                blocks.append({"type": "refrain", "text": text})
        else:
            verse.append(line)
    flush()
    return {"heading": d.strftime("%A"), "reference": clean(m.group(1)), "blocks": blocks}


def parse_gospel(page: str, d: date) -> dict:
    m = section(page, "Gospel", d)
    paras = [clean(re.sub(r"<br\s*/?>", " ", p)) for p in m.group(2).split("</p>")]
    paras = [re.sub(r"^\d+\s+", "", p) for p in paras if p]  # drop any leading verse numbers
    return {"heading": f"Gospel \u2013 {d:%A}, {d:%b} {d.day}", "reference": clean(m.group(1)), "paragraphs": paras}


# --- layout ---------------------------------------------------------------------------

def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def build_flowables(panel, size):
    head = ParagraphStyle("h", fontName="Helvetica-Bold", fontSize=size + 2.5, leading=size + 4, spaceAfter=1)
    sub = ParagraphStyle("s", fontName="Helvetica-Oblique", fontSize=size - 0.5, leading=size + 1,
                         textColor=Color(.3, .3, .3), spaceAfter=size * 0.6)
    body = ParagraphStyle("b", fontName="Times-Roman", fontSize=size, leading=size * 1.12, spaceAfter=size * 0.35)
    refr = ParagraphStyle("r", parent=body, fontName="Times-BoldItalic")
    out = [Paragraph(esc(panel["heading"]), head), Paragraph(esc(panel["reference"]), sub)]
    for text in panel.get("paragraphs", []):  # gospel
        out.append(Paragraph(esc(text), body))
    for b in panel.get("blocks", []):  # psalm
        if b["type"] == "refrain":
            out.append(Paragraph("<b>R.</b> " + esc(b["text"]), refr))
        else:
            out.append(Paragraph("<br/>".join(esc(l) for l in b["lines"]), body))
    return out


def draw_panel(c, panel, x, y, w, h, rotate):
    c.saveState()
    if rotate:
        c.translate(x + w, y + h)
        c.rotate(180)
    else:
        c.translate(x, y)
    inner_w, inner_h = w - 2 * PAD, h - 2 * PAD
    size = MAX_PT  # largest size that fits
    while True:
        flow = build_flowables(panel, size)
        used = sum(f.wrap(inner_w, inner_h)[1] + f.getSpaceAfter() for f in flow)
        if used <= inner_h or size <= MIN_PT:
            break
        size -= 0.1
    if used > inner_h:
        print(f"WARNING: {panel['heading']} ({panel['reference']}) overflows even at {MIN_PT}pt")
    Frame(PAD, PAD, inner_w, inner_h, leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0
          ).addFromList(flow, c)
    c.restoreState()
    return size


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--start", type=date.fromisoformat, help="first day (default: this week's Sunday)")
    ap.add_argument("--out", default="psalms.pdf")
    ap.add_argument("--gospel", type=int, default=7, choices=range(8))
    ap.add_argument("--rotate-top", action="store_true")
    ap.add_argument("--paper", choices=["letter", "a4"], default="letter")
    args = ap.parse_args()
    start = args.start
    if start is None:
        today = date.today()
        start = today - timedelta(days=(today.weekday() + 1) % 7)  # back up to Sunday

    with Browser() as b:
        psalms = [parse_psalm(b.page(d), d) for d in (start + timedelta(days=i) for i in range(7))]
        sunday = start + timedelta(days=7)  # the upcoming Sunday after this week
        gospel = parse_gospel(b.page(sunday), sunday)

    page_w, page_h = landscape(LETTER if args.paper == "letter" else A4)
    pw, ph = page_w / COLS, page_h / ROWS
    c = canvas.Canvas(args.out, pagesize=(page_w, page_h))
    c.setTitle("Weekly Psalms")
    it = iter(psalms)
    for idx in range(COLS * ROWS):
        row, col = divmod(idx, COLS)
        panel = gospel if idx == args.gospel else next(it)
        size = draw_panel(c, panel, col * pw, page_h - (row + 1) * ph, pw, ph, args.rotate_top and row == 0)
        print(f"{panel['heading']:<28} {panel['reference']:<32} {size:.1f}pt")
    c.showPage()
    c.save()
    print(f"Wrote {args.out}")


if __name__ == "__main__":
    main()
