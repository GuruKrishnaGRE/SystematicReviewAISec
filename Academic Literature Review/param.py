"""
Read titles from an Excel file, look them up via Crossref,
and fill metadata columns:

ss_title, ss_year, ss_abstract, ss_venue, ss_authors,
ss_url, ss_doi, ss_is_open_access, ss_pub_types

Install first:
    pip install pandas requests openpyxl
"""

import time
import requests
import pandas as pd

# ========== CONFIG ========== #

INPUT_FILE = "ssrn_papers_complete.xlsx"
OUTPUT_FILE = "ssrn_papers_complete_with_meta.xlsx"

# optional: change this if Crossref complains; ideally put your real email
USER_AGENT = "paper-metadata-script/1.0 (mailto:example@example.com)"

REQUEST_TIMEOUT = 30
SLEEP_BETWEEN_CALLS = 0.2  # seconds


# ========== HELPERS ========== #

def detect_title_column(df: pd.DataFrame) -> str:
    candidates = ["title", "Title", "paper_title", "Paper Title", "Paper title"]
    for c in candidates:
        if c in df.columns:
            return c
    raise ValueError(
        f"Could not find a title column. Available columns: {list(df.columns)}"
    )


def fetch_crossref_details(title: str) -> dict:
    """
    Query Crossref by title, return dict mapping to the expected ss_* columns.
    Returns empty dict if not found or on error.
    """
    if not isinstance(title, str) or not title.strip():
        return {}

    url = "https://api.crossref.org/works"
    params = {
        "query.title": title,
        "rows": 1,
    }
    headers = {"User-Agent": USER_AGENT}

    try:
        resp = requests.get(url, params=params, headers=headers, timeout=REQUEST_TIMEOUT)
        if resp.status_code != 200:
            return {}
        data = resp.json()
    except Exception:
        return {}

    items = (data.get("message") or {}).get("items") or []
    if not items:
        return {}

    item = items[0]

    # title
    cr_title = ""
    if "title" in item and item["title"]:
        cr_title = item["title"][0]

    # year
    cr_year = None
    year = None
    for key in ("published-print", "published-online", "issued"):
        if key in item:
            parts = item[key].get("date-parts") or []
            if parts and parts[0]:
                year = parts[0][0]
                break
    cr_year = year

    # abstract (may be HTML)
    cr_abstract = item.get("abstract")

    # venue / container title
    cr_venue = ""
    if "container-title" in item and item["container-title"]:
        cr_venue = item["container-title"][0]

    # authors
    cr_authors = ""
    if "author" in item:
        names = []
        for a in item["author"]:
            given = a.get("given", "")
            family = a.get("family", "")
            full = " ".join(x for x in [given, family] if x)
            if full:
                names.append(full)
        cr_authors = ", ".join(names)

    # DOI + URL
    cr_doi = item.get("DOI")
    cr_url = item.get("URL")

    # Open access info not directly in Crossref -> leave as None
    cr_is_open_access = None

    # Publication types
    cr_pub_types = ""
    if "type" in item:
        cr_pub_types = item["type"]

    return {
        "ss_title": cr_title,
        "ss_year": cr_year,
        "ss_abstract": cr_abstract,
        "ss_venue": cr_venue,
        "ss_authors": cr_authors,
        "ss_url": cr_url,
        "ss_doi": cr_doi,
        "ss_is_open_access": cr_is_open_access,
        "ss_pub_types": cr_pub_types,
    }


# ========== MAIN ========== #

def main():
    df = pd.read_excel(INPUT_FILE)

    title_col = detect_title_column(df)
    print(f"Using title column: {title_col}")

    # prepare lists for each new column
    ss_title = []
    ss_year = []
    ss_abstract = []
    ss_venue = []
    ss_authors = []
    ss_url = []
    ss_doi = []
    ss_is_open_access = []
    ss_pub_types = []

    titles = df[title_col].astype(str).tolist()

    for i, t in enumerate(titles):
        print(f"{i+1}/{len(titles)}: '{t[:80]}...'")
        meta = fetch_crossref_details(t)

        ss_title.append(meta.get("ss_title"))
        ss_year.append(meta.get("ss_year"))
        ss_abstract.append(meta.get("ss_abstract"))
        ss_venue.append(meta.get("ss_venue"))
        ss_authors.append(meta.get("ss_authors"))
        ss_url.append(meta.get("ss_url"))
        ss_doi.append(meta.get("ss_doi"))
        ss_is_open_access.append(meta.get("ss_is_open_access"))
        ss_pub_types.append(meta.get("ss_pub_types"))

        time.sleep(SLEEP_BETWEEN_CALLS)

    df["ss_title"] = ss_title
    df["ss_year"] = ss_year
    df["ss_abstract"] = ss_abstract
    df["ss_venue"] = ss_venue
    df["ss_authors"] = ss_authors
    df["ss_url"] = ss_url
    df["ss_doi"] = ss_doi
    df["ss_is_open_access"] = ss_is_open_access
    df["ss_pub_types"] = ss_pub_types

    df.to_excel(OUTPUT_FILE, index=False)
    print(f"Saved updated file to: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
