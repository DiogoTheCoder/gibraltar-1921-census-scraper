#!/usr/bin/env python3
"""
Scraper for the 1921 Gibraltar Census
Source: https://www.nationalarchives.gi/1921.aspx  (Gibraltar National Archives)

The site is a classic ASP.NET WebForms application:
  * A <select> dropdown (ctl00$MainContent$DropDownList1) lists every surname.
  * Selecting a surname is an ASP.NET postback that loads that surname's records.
  * A FormView (ctl00$MainContent$FormView1) shows ONE person's full details per
    "page"; the pager (Page$2, Page$3, ...) walks through every person for the
    selected surname.

This script drives those postbacks with the standard library only (no third-party
deps), extracts the 22 full-detail fields for every individual, and writes them
incrementally to a CSV file. It is resumable: surnames already written are skipped
on restart (tracked in a checkpoint file).

Usage:
    python3 scrape_census.py                 # scrape everything
    python3 scrape_census.py --limit 5       # scrape first 5 surnames (test run)
    python3 scrape_census.py --out data.csv  # custom output file
"""
import argparse
import csv
import html as htmlmod
import http.cookiejar
import logging
import os
import re
import threading
import time
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor

log = logging.getLogger("census")


def setup_logging(logfile):
    log.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s %(levelname)s %(message)s",
                            datefmt="%Y-%m-%d %H:%M:%S")
    fh = logging.FileHandler(logfile, encoding="utf-8")
    fh.setFormatter(fmt)
    ch = logging.StreamHandler()
    ch.setFormatter(fmt)
    log.handlers = [fh, ch]

URL = "https://www.nationalarchives.gi/1921.aspx"

# Order matters: this is the column order in the output CSV and the order the
# labels appear in the page's "Full Details" block.
FIELDS = [
    "ID", "Division", "District", "HouseNo", "Address", "NoOfPersons",
    "Surname", "Forename", "RelationToHead", "MaritalStatus", "Sex", "Age",
    "Occupation", "Employer", "Worker", "WorkingOwnAccount", "Birthplace",
    "Religion", "Education", "Disabilities", "No_of_Rooms", "Remarks",
]

DROPDOWN = "ctl00$MainContent$DropDownList1"
FORMVIEW = "ctl00$MainContent$FormView1"

# Each field is rendered as <span id="MainContent_FormView1_<Field>Label">value</span>
_field_res = {
    f: re.compile(
        r'id="MainContent_FormView1_' + re.escape(f) + r'Label">(.*?)</span>',
        re.S,
    )
    for f in FIELDS
}


class Session:
    """Wraps a cookie-aware opener and the ASP.NET hidden-field state."""

    def __init__(self):
        cj = http.cookiejar.CookieJar()
        self.opener = urllib.request.build_opener(
            urllib.request.HTTPCookieProcessor(cj)
        )
        self.opener.addheaders = [
            ("User-Agent", "Mozilla/5.0 (census-archival-scraper)"),
            ("Referer", URL),
        ]
        self.state = {}

    def _fetch(self, data=None, retries=5, tag=""):
        body = urllib.parse.urlencode(data).encode() if data else None
        last = None
        for attempt in range(1, retries + 1):
            try:
                req = urllib.request.Request(URL, data=body)
                with self.opener.open(req, timeout=60) as r:
                    status = getattr(r, "status", 200)
                    text = r.read().decode("utf-8", "replace")
                    if status != 200:
                        last = f"HTTP {status}"
                        raise RuntimeError(last)
                    return text
            except Exception as e:  # noqa: BLE001 - network/server flakiness, retry
                last = e
                wait = 2 * attempt
                log.warning("request%s failed (attempt %d/%d): %s - retrying in %ds",
                            f" [{tag}]" if tag else "", attempt, retries, e, wait)
                if attempt < retries:
                    time.sleep(wait)
        raise RuntimeError(f"request failed after {retries} tries: {last}")

    def _capture_state(self, html):
        for name in (
            "__VIEWSTATE",
            "__VIEWSTATEGENERATOR",
            "__EVENTVALIDATION",
            "__VIEWSTATEENCRYPTED",
        ):
            m = re.search(r'id="' + name + r'"[^>]*value="([^"]*)"', html)
            if m:
                self.state[name] = htmlmod.unescape(m.group(1))

    def load(self):
        html = self._fetch()
        self._capture_state(html)
        return html

    def postback(self, target, argument, surname, tag=""):
        data = dict(self.state)
        data["__EVENTTARGET"] = target
        data["__EVENTARGUMENT"] = argument
        data[DROPDOWN] = surname
        html = self._fetch(data, tag=tag)
        self._capture_state(html)
        return html


def surnames_from(html):
    opts = re.findall(r"<option[^>]*value=\"([^\"]*)\"", html)
    return [htmlmod.unescape(o) for o in opts]


def total_records(html):
    m = re.search(r'id="MainContent_Label2">(\d+)', html)
    return int(m.group(1)) if m else 0


def parse_details(html):
    if "Full Details" not in html:
        return None
    row = {}
    found = False
    for f, rx in _field_res.items():
        m = rx.search(html)
        if m:
            val = htmlmod.unescape(re.sub(r"<[^>]+>", " ", m.group(1)))
            row[f] = re.sub(r"\s+", " ", val).strip()
            found = True
        else:
            row[f] = ""
    return row if found else None


def scrape_surname(sess, surname):
    """Return a list with one dict per individual for the given surname.

    Raises if the number of records parsed does not match the total the site
    reports ("N record(s) Found"), so an incomplete surname is retried rather
    than silently under-counted.
    """
    html = sess.postback(DROPDOWN, "", surname, tag=f"{surname} select")
    total = total_records(html)
    if total == 0:
        return []
    rows = []
    for page in range(1, total + 1):
        if page > 1:
            html = sess.postback(FORMVIEW, f"Page${page}", surname,
                                 tag=f"{surname} p{page}/{total}")
        row = parse_details(html)
        # A page must yield a record; retry the postback a few times if not.
        attempt = 0
        while row is None and attempt < 3:
            attempt += 1
            log.warning("%s p%d/%d: no record parsed - retry %d",
                        surname, page, total, attempt)
            html = sess.postback(FORMVIEW, f"Page${page}", surname,
                                 tag=f"{surname} p{page}/{total} retry{attempt}")
            row = parse_details(html)
        if row is not None:
            rows.append(row)
    if len(rows) != total:
        raise RuntimeError(
            f"{surname}: parsed {len(rows)} of {total} records")
    return rows


def load_done(checkpoint):
    if not os.path.exists(checkpoint):
        return set()
    with open(checkpoint, encoding="utf-8") as f:
        return set(line.rstrip("\n") for line in f)


def fmt_hms(seconds):
    seconds = int(seconds)
    h, rem = divmod(seconds, 3600)
    m, s = divmod(rem, 60)
    return f"{h:d}:{m:02d}:{s:02d}"


def progress_bar(done, total, width=24):
    filled = int(width * done / total) if total else width
    return "[" + "#" * filled + "-" * (width - filled) + "]"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="gibraltar_1921_census.csv")
    ap.add_argument("--checkpoint", default="done_surnames.txt")
    ap.add_argument("--limit", type=int, default=0, help="only N surnames (test)")
    ap.add_argument("--delay", type=float, default=0.3, help="seconds between requests")
    ap.add_argument("--workers", type=int, default=1,
                    help="number of parallel workers (each with its own session)")
    ap.add_argument("--logfile", default="scrape.log")
    args = ap.parse_args()

    setup_logging(args.logfile)

    sess = Session()
    log.info("Loading surname list...")
    html = sess.load()
    surnames = surnames_from(html)
    log.info("Found %d surname entries.", len(surnames))

    done = load_done(args.checkpoint)
    if done:
        log.info("Resuming: %d surnames already done.", len(done))

    todo = [s for s in surnames if s not in done]
    if args.limit:
        todo = todo[: args.limit]
    log.info("To do: %d surnames with %d worker(s).", len(todo), args.workers)

    file_exists = os.path.exists(args.out)
    out = open(args.out, "a", newline="", encoding="utf-8")
    writer = csv.DictWriter(out, fieldnames=FIELDS)
    if not file_exists:
        writer.writeheader()
        out.flush()
    ck = open(args.checkpoint, "a", encoding="utf-8")

    io_lock = threading.Lock()          # guards the CSV + checkpoint files
    progress = {"surnames": 0, "records": 0, "errors": 0}

    def worker(surname):
        """Scrape one surname with its own session; return (name, rows, error)."""
        s = Session()
        s.load()
        try:
            rows = list(scrape_surname(s, surname))
            return surname, rows, None
        except Exception as e:  # noqa: BLE001
            return surname, None, e

    if args.workers <= 1:
        results = (worker(s) for s in todo)
    else:
        pool = ThreadPoolExecutor(max_workers=args.workers)
        results = pool.map(worker, todo)

    total = len(todo)
    start = time.time()
    for surname, rows, err in results:
        with io_lock:
            done_n = progress["surnames"] + progress["errors"] + 1
            if err is not None:
                progress["errors"] += 1
                log.error("%s: FAILED (%s) - will retry next run", surname, err)
            else:
                for row in rows:
                    writer.writerow(row)
                out.flush()
                ck.write(surname + "\n")
                ck.flush()
                progress["surnames"] += 1
                progress["records"] += len(rows)
            elapsed = time.time() - start
            rate = done_n / elapsed if elapsed else 0
            eta = (total - done_n) / rate if rate else 0
            pct = 100.0 * done_n / total
            log.info("%s %5.1f%% (%d/%d) | %d recs | %d err | %.1f/s | ETA %s | %s +%d",
                     progress_bar(done_n, total), pct, done_n, total,
                     progress["records"], progress["errors"], rate,
                     fmt_hms(eta), surname, len(rows) if rows else 0)
        if args.delay:
            time.sleep(args.delay)

    out.close()
    ck.close()
    log.info("Done. Wrote %d new records (%d surname errors) to %s",
             progress["records"], progress["errors"], args.out)


if __name__ == "__main__":
    main()
