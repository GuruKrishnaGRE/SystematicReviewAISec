"""
Fetch arXiv papers for the security query from the LAST 5 YEARS
and save them to an Excel file.

Install once:
    pip install requests feedparser pandas openpyxl
"""

import datetime
import calendar
import time

import requests
import feedparser
import pandas as pd
from requests.exceptions import ReadTimeout, ConnectionError

# ---------------- CONFIG ---------------- #

QUERY = (
    '"open source security" OR '
    '"AI security" OR '
    '"machine learning security" OR '
    '"LLM security" OR '
    '"adversarial attack" OR '
    '"model poisoning" OR '
    '"supply chain attack"'
)

BASE_URL = "https://export.arxiv.org/api/query"

BATCH_SIZE = 25              # small to reduce timeout risk
MAX_BATCHES = 200            # 25 * 200 = 5000 max
REQUEST_TIMEOUT = (10, 300)  # (connect timeout, read timeout) in seconds
MAX_RETRIES_PER_BATCH = 3
SLEEP_BETWEEN_BATCHES = 5    # seconds between successful batches

OUTPUT_FILE = "arxiv_security_last5years.xlsx"

# Last 5 years cutoff (UTC)
CUTOFF_DATE = datetime.datetime.utcnow() - datetime.timedelta(days=5 * 365)


# ---------------- HELPERS ---------------- #

def parse_time(struct_time_obj):
    """Convert feedparser's struct_time to a UTC datetime."""
    if struct_time_obj is None:
        return None
    return datetime.datetime.utcfromtimestamp(calendar.timegm(struct_time_obj))


# ---------------- MAIN FETCH ---------------- #

def fetch_recent_papers():
    all_records = []
    start = 0

    headers = {"User-Agent": "paramarxiv-script/1.0"}

    for batch_idx in range(MAX_BATCHES):
        print(f"Batch {batch_idx + 1}: results {start}–{start + BATCH_SIZE - 1}")

        params = {
            "search_query": QUERY,
            "start": start,
            "max_results": BATCH_SIZE,
            "sortBy": "submittedDate",
            "sortOrder": "descending",
        }

        # retry loop for this batch
        attempt = 0
        while True:
            try:
                resp = requests.get(
                    BASE_URL,
                    params=params,
                    headers=headers,
                    timeout=REQUEST_TIMEOUT,
                )
            except (ReadTimeout, ConnectionError):
                attempt += 1
                if attempt > MAX_RETRIES_PER_BATCH:
                    print("Gave up on this batch due to repeated timeouts.")
                    return all_records
                wait = 30 * attempt
                print(f"Timeout, retry {attempt}/{MAX_RETRIES_PER_BATCH} in {wait}s...")
                time.sleep(wait)
                continue

            if resp.status_code == 200:
                break
            elif resp.status_code == 429:
                attempt += 1
                if attempt > MAX_RETRIES_PER_BATCH:
                    print("Too many 429 responses, stopping.")
                    return all_records
                wait = 60 * attempt
                print(f"HTTP 429, retry {attempt}/{MAX_RETRIES_PER_BATCH} in {wait}s...")
                time.sleep(wait)
                continue
            else:
                print(f"HTTP {resp.status_code}, stopping.")
                return all_records

        feed = feedparser.parse(resp.text)
        entries = feed.entries

        if not entries:
            print("No more entries returned. Done.")
            break

        stop_due_to_date = False

        for e in entries:
            published_dt = parse_time(getattr(e, "published_parsed", None))
            updated_dt = parse_time(getattr(e, "updated_parsed", None))

            if published_dt is None:
                continue

            # feed is sorted desc; once below cutoff we can stop
            if published_dt < CUTOFF_DATE:
                stop_due_to_date = True
                break

            title = getattr(e, "title", "").strip()
            summary = getattr(e, "summary", "").replace("\n", " ").strip()

            authors = ""
            if hasattr(e, "authors"):
                authors = ", ".join(
                    a.name for a in e.authors if hasattr(a, "name")
                )

            primary_category = ""
            categories = ""
            if hasattr(e, "tags"):
                tags = [t["term"] for t in e.tags if "term" in t]
                categories = ", ".join(tags)
                if tags:
                    primary_category = tags[0]

            entry_id = getattr(e, "id", "").strip()
            arxiv_id = ""
            if entry_id:
                # typical format: http://arxiv.org/abs/xxxx.xxxxxvY
                arxiv_id = entry_id.split("/abs/")[-1]
            pdf_url = f"https://arxiv.org/pdf/{arxiv_id}.pdf" if arxiv_id else ""

            record = {
                "title": title,
                "authors": authors,
                "summary": summary,
                "published": published_dt,
                "updated": updated_dt,
                "primary_category": primary_category,
                "categories": categories,
                "arxiv_id": arxiv_id,
                "pdf_url": pdf_url,
                "entry_id": entry_id,
            }

            all_records.append(record)

        if stop_due_to_date:
            print("Reached papers older than 5 years. Stopping.")
            break

        start += BATCH_SIZE
        time.sleep(SLEEP_BETWEEN_BATCHES)

    return all_records


def main():
    records = fetch_recent_papers()

    if not records:
        print("No records fetched.")
        return

    df = pd.DataFrame(records)
    df = df.sort_values("published", ascending=False)
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"Saved {len(df)} papers to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
