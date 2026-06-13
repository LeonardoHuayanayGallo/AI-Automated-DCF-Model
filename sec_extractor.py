"""
sec_extractor.py
Pulls the most recent 10-K filing for a given company ticker
from the SEC EDGAR API and extracts the MD&A section.
"""

import requests
import re
from bs4 import BeautifulSoup

HEADERS = {"User-Agent": "Leonardo Gallo leonardo@email.com"}  # keep your name/email here

def get_cik(ticker: str) -> str:
    """Resolve ticker symbol to SEC CIK number."""
    tickers_url = "https://www.sec.gov/files/company_tickers.json"
    r = requests.get(tickers_url, headers=HEADERS)
    data = r.json()
    for entry in data.values():
        if entry["ticker"].upper() == ticker.upper():
            return str(entry["cik_str"]).zfill(10)
    raise ValueError(f"Ticker {ticker} not found in SEC database.")

def get_latest_10k_url(cik: str) -> str:
    """Fetch the URL of the most recent 10-K filing."""
    url = f"https://data.sec.gov/submissions/CIK{cik}.json"
    r = requests.get(url, headers=HEADERS)
    filings = r.json()["filings"]["recent"]

    for i, form in enumerate(filings["form"]):
        if form == "10-K":
            accession = filings["accessionNumber"][i].replace("-", "")
            doc = filings["primaryDocument"][i]
            filing_url = (
                f"https://www.sec.gov/Archives/edgar/data/"
                f"{int(cik)}/{accession}/{doc}"
            )
            return filing_url
    raise ValueError("No 10-K filing found.")

def extract_mda(filing_url: str) -> str:
    """
    Download 10-K HTML and extract the MD&A section.
    Uses multiple strategies to handle different filing formats.
    """
    r = requests.get(filing_url, headers=HEADERS, timeout=30)
    soup = BeautifulSoup(r.content, "html.parser")

    # Remove script and style tags
    for tag in soup(["script", "style", "ix:header"]):
        tag.decompose()

    # Get clean text
    text = soup.get_text(separator=" ", strip=True)

    # Clean up excessive whitespace
    text = re.sub(r'\s+', ' ', text).strip()

    # Strategy 1: Look for Item 7 MD&A section
    patterns = [
        r'(Item\s+7\.?\s+Management.{0,10}Discussion.{0,200}?)(Item\s+7A|Item\s+8)',
        r'(ITEM\s+7\.?\s+MANAGEMENT.{0,10}DISCUSSION.{0,200}?)(ITEM\s+7A|ITEM\s+8)',
        r'(Management.{0,5}s Discussion and Analysis.{0,200}?)(Quantitative and Qualitative|Financial Statements)',
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
        if match:
            mda_text = match.group(1)
            words = mda_text.split()
            if len(words) > 100:  # Only accept if substantial content found
                print(f"[SEC] MD&A extracted via pattern: {len(words)} words")
                return " ".join(words[:4000])

    # Strategy 2: Find the largest text block around financial keywords
    financial_keywords = [
        "revenue", "operating income", "net income", "cash flow",
        "growth", "margin", "fiscal year", "results of operations"
    ]

    # Split text into chunks and find the most finance-dense section
    words = text.split()
    chunk_size = 500
    best_chunk_start = 0
    best_score = 0

    for i in range(0, len(words) - chunk_size, 100):
        chunk = " ".join(words[i:i + chunk_size]).lower()
        score = sum(chunk.count(kw) for kw in financial_keywords)
        if score > best_score:
            best_score = score
            best_chunk_start = i

    # Return 4000 words starting from the best financial section
    extracted = " ".join(words[best_chunk_start:best_chunk_start + 4000])
    print(f"[SEC] MD&A extracted via keyword strategy: {len(extracted.split())} words")
    return extracted

def run(ticker: str) -> str:
    print(f"[SEC] Fetching 10-K for {ticker}...")
    cik = get_cik(ticker)
    url = get_latest_10k_url(cik)
    print(f"[SEC] Filing URL: {url}")
    mda = extract_mda(url)
    return mda