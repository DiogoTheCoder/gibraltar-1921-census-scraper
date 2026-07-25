# 1921 Gibraltar Census Scraper

Scrapes the full **1921 Gibraltar Census** dataset from the Gibraltar National
Archives and exports it to CSV / Excel.

Source: <https://www.nationalarchives.gi/1921.aspx>

## The dataset (included in this repo)

* **`gibraltar_1921_census.csv`** / **`gibraltar_1921_census.xlsx`**
* **18,697 individuals** across **3,269 surname entries**, recorded on 19 June 1921
* 22 fields per person (see below)
* Complete and verified: record IDs run contiguously 1–18,697 with **no gaps and
  no duplicates**, and every surname's row count was checked against the site's
  own "N record(s) Found". (The archive's preface rounds this to "18,700".)
* The `.xlsx` has a frozen header row and column filters ready to use.

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

### Build the Excel workbook

```bash
pip install openpyxl
python3 build_workbook.py         # Summary -> Geography -> Raw Data
# or a plain single-sheet export:
python3 make_excel.py
```

The `.xlsx` has four sheets:

1. **Summary** — headline totals and breakdowns by Ward, Sex, Religion, Marital
   Status, Education, and top Birthplaces / Occupations.
2. **Electors** — estimated electors per ward AND per police district under the 1921 franchise (see below).
3. **Geography** — Ward → Division → Police District → headcount + representative
   streets.
4. **Raw Data** — every record exactly as scraped (no added columns).

### Elector estimate (City Council Ordinance 1921, ss. 14-15)

The franchise: a British subject of full age, not under legal incapacity and not
in arrears of rates, who had occupied premises as owner, tenant or lodger for at
least six months — **women excluded**, and **Crown servants in rent-free quarters
excluded**. The census records none of nationality, rates, incapacity or tenure
directly, so the Electors sheet is an **estimate** built from proxies (Sex = M,
Age ≥ 21, Birthplace/"BS" annotation for British-subject status, RelationToHead ∈
{Head, Lodger}, Divisions 1-4 only). It comes to roughly **2,950 electors** across
the four wards (see the sheet for the per-ward figures, a boarder-inclusive
sensitivity, and full caveats).

### Ward / Division / District mapping

The census `Division` corresponds to a **Ward** and `District` to a **Police
District**, per the *City Council Ordinance 1921, First Schedule (List of
Wards)*:

| Ward | Census Division | Police Districts |
|------|-----------------|------------------|
| Old Town Ward  | 1 | 1, 2, 3, 4, 7, 9, 10, 15 |
| Castle Ward    | 2 | 5, 6, 8, 11, 12, 13, 14, 18 |
| Cathedral Ward | 3 | 16, 17, 19, 25, 26, 27 |
| Europa Ward    | 4 | 20, 21, 22, 23, 24 + the "(South)" series 1–11 |
| *Military & Government establishments* | 5 | (special enumeration, no district) |
| *Shipping (afloat)* | 6 | (ships in the Bay, no district) |

District numbers **restart** in Europa Ward (the "(South)" series), so the same
number appears under more than one Division — always read `Division` + `District`
together.

## Notes / etiquette

* A polite `--delay` (default 0.3s) is applied between requests.
* This is a public historical/genealogical record published by the Gibraltar
  National Archives. Please use the data respectfully and credit the source.
