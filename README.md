# 1921 Gibraltar Census Scraper

Scrapes the full **1921 Gibraltar Census** dataset (~18,700 civilian residents
recorded on 19 June 1921) from the Gibraltar National Archives and exports it to
CSV / Excel.

Source: <https://www.nationalarchives.gi/1921.aspx>

## What it does

The archive site is an ASP.NET WebForms application. Each surname is an entry in a
dropdown; selecting one triggers a server postback that loads that surname's
records, and a paged detail view (`FormView`) shows one person's full record per
page. This scraper drives those postbacks — walking every surname and every page —
and extracts the 22 fields recorded for each individual.

### Fields captured (per person)

`ID, Division, District, HouseNo, Address, NoOfPersons, Surname, Forename,
RelationToHead, MaritalStatus, Sex, Age, Occupation, Employer, Worker,
WorkingOwnAccount, Birthplace, Religion, Education, Disabilities, No_of_Rooms,
Remarks`

## Usage

Requires only Python 3 (standard library — no dependencies).

```bash
# Full run (writes gibraltar_1921_census.csv). Takes a while — ~18.7k requests.
python3 scrape_census.py

# Quick test — first 5 surnames only
python3 scrape_census.py --limit 5

# Options
python3 scrape_census.py --out data.csv --checkpoint done.txt --delay 0.3
```

The scraper is **resumable**: completed surnames are recorded in
`done_surnames.txt`, and rows are appended to the CSV as they are collected, so you
can stop and restart without losing or duplicating work.

### Convert to Excel (optional)

```bash
pip install pandas openpyxl
python3 -c "import pandas as pd; pd.read_csv('gibraltar_1921_census.csv').to_excel('gibraltar_1921_census.xlsx', index=False)"
```

## Notes / etiquette

* A polite `--delay` (default 0.3s) is applied between requests.
* This is a public historical/genealogical record published by the Gibraltar
  National Archives. Please use the data respectfully and credit the source.
