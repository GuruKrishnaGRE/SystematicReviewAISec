"""
SSRN Details Fetcher - FOR LOCAL MACHINE USE

This script should be run on YOUR LOCAL COMPUTER where network restrictions don't apply.

INSTRUCTIONS:
1. Download this script to your computer
2. Download the ssrn_all_papers.csv file to the same folder
3. Install required libraries: pip install pandas beautifulsoup4 requests openpyxl
4. Run: python fetch_ssrn_local.py
5. Wait ~8-10 minutes for completion
6. You'll get ssrn_papers_complete.xlsx with all details

"""

import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
import re

def fetch_paper_details(url, paper_id, total):
    """Fetch full details for a single SSRN paper."""
    print(f"[{paper_id}/{total}] Fetching: {url[:60]}...", end=' ', flush=True)
    
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        response = requests.get(url, headers=headers, timeout=15)
        
        if response.status_code != 200:
            print(f"HTTP {response.status_code}")
            return None
        
        soup = BeautifulSoup(response.content, 'html.parser')
        details = {}
        
        # Extract authors from meta tags
        author_metas = soup.find_all('meta', attrs={'name': 'citation_author'})
        if author_metas:
            authors = [meta.get('content', '').strip() for meta in author_metas if meta.get('content')]
            if authors:
                details['authors'] = '; '.join(authors)
        
        # Extract abstract from meta tag
        abstract_meta = soup.find('meta', attrs={'name': 'citation_abstract'})
        if abstract_meta and abstract_meta.get('content'):
            details['abstract'] = abstract_meta.get('content').strip()
        
        # Extract keywords
        keywords_section = soup.find('p', class_='keywords')
        if keywords_section:
            keywords_text = keywords_section.get_text(strip=True)
            keywords_text = re.sub(r'^Keywords?:\s*', '', keywords_text, flags=re.IGNORECASE)
            details['keywords'] = keywords_text
        
        # Extract date
        date_meta = soup.find('meta', attrs={'name': 'citation_publication_date'})
        if date_meta:
            details['date'] = date_meta.get('content', '')
        
        # Extract source/journal
        journal_meta = soup.find('meta', attrs={'name': 'citation_journal_title'})
        if journal_meta:
            details['source'] = journal_meta.get('content', '')
        
        print("✓")
        return details
        
    except Exception as e:
        print(f"Error: {str(e)[:30]}")
        return None

def main():
    # File paths - adjust these if needed
    input_csv = "ssrn_all_papers.csv"
    output_excel = "ssrn_papers_complete.xlsx"
    
    print("=" * 80)
    print("SSRN Full Details Fetcher - Local Version")
    print("=" * 80)
    
    # Read CSV
    print(f"\nReading: {input_csv}")
    df = pd.read_csv(input_csv)
    total = len(df)
    print(f"Found {total} papers")
    print(f"Estimated time: ~{int(total * 1.5 / 60)} minutes\n")
    
    # Add columns
    df['authors'] = ''
    df['abstract'] = ''
    df['keywords'] = ''
    df['source'] = ''
    df['date'] = ''
    
    # Fetch details
    start_time = time.time()
    for idx, row in df.iterrows():
        details = fetch_paper_details(row['url'], idx + 1, total)
        if details:
            for key, value in details.items():
                df.at[idx, key] = value
        time.sleep(1.5)
        
        if (idx + 1) % 50 == 0:
            elapsed = time.time() - start_time
            remaining = (total - idx - 1) * (elapsed / (idx + 1))
            print(f"\nProgress: {idx+1}/{total} | Remaining: {int(remaining/60)}m {int(remaining%60)}s\n")
    
    # Save
    print(f"\nSaving to: {output_excel}")
    column_order = ['number', 'authors', 'title', 'source', 'abstract', 'keywords', 'date', 'url']
    existing = [col for col in column_order if col in df.columns]
    remaining = [col for col in df.columns if col not in existing]
    df = df[existing + remaining]
    df.to_excel(output_excel, index=False)
    
    print("\n" + "=" * 80)
    print("COMPLETE!")
    print(f"Authors found: {df['authors'].notna().sum()}/{total}")
    print(f"Abstracts found: {df['abstract'].notna().sum()}/{total}")
    print(f"Keywords found: {df['keywords'].notna().sum()}/{total}")
    print("=" * 80)

if __name__ == "__main__":
    main()
