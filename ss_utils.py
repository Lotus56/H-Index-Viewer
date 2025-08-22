"""
ss_utils.py

Helper module for HIndex Viewer:
- Semantic Scholar API helpers (search, details, coauthors)
- cache management (ss_author_cache.json)
- Zotero parsers

Usage: keep this file in the same folder as the GUI (hindex_gui.py) and import it.
"""

import time, json, os
from urllib.parse import quote_plus
import requests

# Configuration
SS_BASE = "https://api.semanticscholar.org/graph/v1"
SS_AUTHOR_FIELDS = "name,affiliations,hIndex,paperCount"
CACHE_FILE = "ss_author_cache.json"
DEFAULT_DELAY = 3.0

# Module-level cache dict (populated by load_cache)
cache = {}

# Cache helpers
def load_cache():
    global cache
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, 'r', encoding='utf8') as f:
                cache = json.load(f)
        except Exception as e:
            print("ss_utils.load_cache: failed to load cache:", e)
            cache = {}
    else:
        cache = {}
    return cache


def save_cache():
    global cache
    try:
        with open(CACHE_FILE, 'w', encoding='utf8') as f:
            json.dump(cache, f, indent=2)
    except Exception as e:
        print("ss_utils.save_cache: failed to save cache:", e)

# HTTP helper with retries/backoff
def safe_get(url, headers=None, base_delay=DEFAULT_DELAY, max_retries=6):
    headers = headers or {}
    last_resp = None
    for attempt in range(max_retries):
        try:
            resp = requests.get(url, headers=headers, timeout=30)
        except Exception as e:
            print(f"[safe_get] Network exception for {url}: {e}")
            return None
        last_resp = resp
        if resp.status_code == 200:
            return resp
        if resp.status_code == 429:
            wait = base_delay * (2 ** attempt) + 1.0
            print(f"[safe_get] 429 received. Waiting {wait:.1f}s before retry {attempt+1}/{max_retries}...")
            time.sleep(wait)
            continue
        if resp.status_code in (500, 502, 503, 504):
            wait = base_delay * (2 ** attempt)
            print(f"[safe_get] Server {resp.status_code}. Waiting {wait:.1f}s before retry {attempt+1}/{max_retries}...")
            time.sleep(wait)
            continue
        return resp
    return last_resp

# Semantic Scholar helpers

def ss_search_author(name, api_key=None, base_delay=DEFAULT_DELAY):
    global cache
    cache_key = f"search::{name}"
    if cache_key in cache:
        return cache[cache_key]
    q = quote_plus(name)
    url = f"{SS_BASE}/author/search?query={q}&fields={quote_plus(SS_AUTHOR_FIELDS)}&limit=5"
    headers = {}
    if api_key:
        headers['x-api-key'] = api_key
    r = safe_get(url, headers=headers, base_delay=base_delay)
    if r is None:
        print(f"[ss_search_author] Network fail for '{name}'")
        return None
    if r.status_code != 200:
        print(f"[ss_search_author] Non-200 for '{name}': {r.status_code}")
        try:
            print(r.text[:1000])
        except Exception:
            pass
        return None
    try:
        j = r.json()
    except Exception as e:
        print(f"[ss_search_author] JSON parse failed for '{name}': {e}")
        return None
    data = j.get('data', [])
    top = data[0] if data else None
    cache[cache_key] = top
    save_cache()
    return top


def ss_get_author_details(author_id, api_key=None, fields="name,hIndex,paperCount", base_delay=DEFAULT_DELAY):
    global cache
    if not author_id:
        return None
    cache_key = f"author::{author_id}"
    if cache_key in cache:
        return cache[cache_key]
    headers = {}
    if api_key:
        headers['x-api-key'] = api_key
    tried_fields = fields
    for _ in range(2):
        url = f"{SS_BASE}/author/{author_id}?fields={quote_plus(tried_fields)}"
        r = safe_get(url, headers=headers, base_delay=base_delay)
        if r is None:
            print(f"[ss_get_author_details] Network fail for id {author_id}")
            return None
        if r.status_code == 200:
            try:
                j = r.json()
            except Exception as e:
                print(f"[ss_get_author_details] JSON parse failed for id {author_id}: {e}")
                return None
            cache[cache_key] = j
            save_cache()
            return j
        # handle unsupported fields fallback
        if r.status_code == 400 and 'Unrecognized or unsupported fields' in r.text and 'aliases' in tried_fields:
            tried_fields = ",".join([f for f in tried_fields.split(",") if f.strip().lower() != 'aliases'])
            continue
        print(f"[ss_get_author_details] Non-200 for id {author_id}: {r.status_code}")
        try:
            print(r.text[:1000])
        except Exception:
            pass
        return None
    return None


def ss_get_author_coauthors(author_id, api_key=None, max_papers=50, base_delay=DEFAULT_DELAY):
    global cache
    if not author_id:
        return {}
    cache_key = f"coauthors::{author_id}"
    if cache_key in cache:
        return cache[cache_key]
    headers = {}
    if api_key:
        headers['x-api-key'] = api_key
    url = f"{SS_BASE}/author/{author_id}/papers?fields=title,authors&limit={max_papers}"
    r = safe_get(url, headers=headers, base_delay=base_delay)
    if r is None:
        print(f"[ss_get_author_coauthors] Network fail for id {author_id}")
        return {}
    if r.status_code != 200:
        print(f"[ss_get_author_coauthors] Non-200 for id {author_id}: {r.status_code}")
        try:
            print(r.text[:1000])
        except Exception:
            pass
        return {}
    try:
        j = r.json()
    except Exception as e:
        print(f"[ss_get_author_coauthors] JSON parse failed for id {author_id}: {e}")
        return {}
    co_counts = {}
    papers = j.get('data', []) if isinstance(j, dict) else []
    for p in papers:
        for a in p.get('authors', []) or []:
            name = a.get('name') or a.get('authorName') or None
            if not name:
                continue
            co_counts[name] = co_counts.get(name, 0) + 1
    cache[cache_key] = co_counts
    save_cache()
    return co_counts

# -----------------------
# Zotero / CSV parsers
import csv, re

def parse_zotero_csv(path):
    names = set()
    with open(path, newline='', encoding='utf8', errors='ignore') as f:
        reader = csv.DictReader(f)
        for row in reader:
            for key in ['Creators', 'Creator', 'Authors', 'Author', 'author']:
                if key in row and row[key]:
                    text = row[key]
                    parts = re.split(r';|\band\b', text)
                    for p in parts:
                        p = p.strip()
                        if not p: continue
                        if ',' in p:
                            last, first = [x.strip() for x in p.split(',',1)]
                            name = f"{first} {last}"
                        else:
                            name = p
                        names.add(name)
                    break
    return sorted(names)


def parse_bibtex_authors(path):
    names = set()
    txt = open(path, encoding='utf8', errors='ignore').read()
    for m in re.finditer(r'author\s*=\s*\{([^}]*)\}', txt, flags=re.I | re.S):
        authors_field = m.group(1)
        parts = [p.strip() for p in re.split(r'\s+and\s+', authors_field)]
        for p in parts:
            if ',' in p:
                last, first = [x.strip() for x in p.split(',',1)]
                name = f"{first} {last}"
            else:
                name = p
            if name:
                names.add(name)
    return sorted(names)
