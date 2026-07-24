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
import os
import re
import sys
import time
import urllib.parse
import urllib.request

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

    def _fetch(self, data=None, retries=4):
        body = urllib.parse.urlencode(data).encode() if data else None
        last = None
        for attempt in range(retries):
            try:
                req = urllib.request.Request(URL, data=body)
                with self.opener.open(req, timeout=60) as r:
                    return r.read().decode("utf-8", "replace")
            except Exception as e:  # noqa: BLE001 - network flakiness, retry
                last = e
                time.sleep(2 * (attempt + 1))
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

    def postback(self, target, argument, surname):
        data = dict(self.state)
        data["__EVENTTARGET"] = target
        data["__EVENTARGUMENT"] = argument
        data[DROPDOWN] = surname
        html = self._fetch(data)
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
    """Yield one dict per individual for the given surname."""
    html = sess.postback(DROPDOWN, "", surname)
    total = total_records(html)
    if total == 0:
        return
    row = parse_details(html)
    if row:
        yield row
    for page in range(2, total + 1):
        html = sess.postback(FORMVIEW, f"Page${page}", surname)
        row = parse_details(html)
        if row:
            yield row


def load_done(checkpoint):
    if not os.path.exists(checkpoint):
        return set()
    with open(checkpoint, encoding="utf-8") as f:
        return set(line.rstrip("\n") for line in f)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="gibraltar_1921_census.csv")
    ap.add_argument("--checkpoint", default="done_surnames.txt")
    ap.add_argument("--limit", type=int, default=0, help="only N surnames (test)")
    ap.add_argument("--delay", type=float, default=0.3, help="seconds between requests")
    args = ap.parse_args()

    sess = Session()
    print("Loading surname list...", flush=True)
    html = sess.load()
    surnames = surnames_from(html)
    print(f"Found {len(surnames)} surname entries.", flush=True)

    done = load_done(args.checkpoint)
    if done:
        print(f"Resuming: {len(done)} surnames already done.", flush=True)

    todo = [s for s in surnames if s not in done]
    if args.limit:
        todo = todo[: args.limit]

    file_exists = os.path.exists(args.out)
    out = open(args.out, "a", newline="", encoding="utf-8")
    writer = csv.DictWriter(out, fieldnames=FIELDS)
    if not file_exists:
        writer.writeheader()
        out.flush()

    ck = open(args.checkpoint, "a", encoding="utf-8")
    grand_total = 0
    for n, surname in enumerate(todo, 1):
        try:
            count = 0
            for row in scrape_surname(sess, surname):
                writer.writerow(row)
                count += 1
                grand_total += 1
                time.sleep(args.delay)
            out.flush()
            ck.write(surname + "\n")
            ck.flush()
            print(f"[{n}/{len(todo)}] {surname!r}: {count} records "
                  f"(running total {grand_total})", flush=True)
        except Exception as e:  # noqa: BLE001
            print(f"[{n}/{len(todo)}] {surname!r}: ERROR {e} - will retry next run",
                  file=sys.stderr, flush=True)
            # reset session state on error before moving on
            sess = Session()
            sess.load()
    out.close()
    ck.close()
    print(f"Done. Wrote {grand_total} new records to {args.out}", flush=True)


if __name__ == "__main__":
    main()
