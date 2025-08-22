"""
HIndex_Viewer_v3.py

Upgraded version (v3) of the Zotero -> Semantic Scholar H-index viewer.
Main new features requested:

1) On startup: if `ss_author_cache.json` exists it is loaded and used to autopopulate the Authors table
   (search:: and author:: cache entries are interpreted to pre-fill fields).

2) When importing a CSV, new names are added but existing names / information are NOT overwritten.
   Import is case-insensitive and will skip names already present in the Authors table.

3) The Co-authors tab now shows the same Semantic Scholar fields as the Authors tab
   (SS Name, H-index, papers, SSid) for each co-author. Co-author details are taken from cache
   when available; if many coauthors are present the app asks before making many API calls.

Requirements: Python 3.8+, requests. Install with:
    py -m pip install --user requests

Run:
    py HIndex_Viewer_v3.py

Notes:
 - Cache file: ss_author_cache.json in the same folder. Delete it to force fresh queries.
 - Use the API key box if you have a Semantic Scholar API key to increase rate limits.

"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import requests, time, json, csv, re, os, webbrowser
from urllib.parse import quote_plus

# -----------------------
# Configuration
SS_BASE = "https://api.semanticscholar.org/graph/v1"
SS_AUTHOR_FIELDS = "name,affiliations,hIndex,paperCount"
CACHE_FILE = "ss_author_cache.json"
DEFAULT_DELAY = 3.0  # seconds — increased default to reduce 429s

# -----------------------
# Cache helpers

def load_cache():
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, 'r', encoding='utf8') as f:
                return json.load(f)
        except Exception as e:
            print("Failed to load cache:", e)
            return {}
    return {}


def save_cache(cache):
    try:
        with open(CACHE_FILE, 'w', encoding='utf8') as f:
            json.dump(cache, f, indent=2)
    except Exception as e:
        print("Failed to save cache:", e)

# -----------------------
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

# -----------------------
# Semantic Scholar helpers

def ss_search_author(name, api_key=None, base_delay=DEFAULT_DELAY):
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
    save_cache(cache)
    return top


def ss_get_author_details(author_id, api_key=None, fields="name,hIndex,paperCount", base_delay=DEFAULT_DELAY):
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
            save_cache(cache)
            return j
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
    save_cache(cache)
    return co_counts

# -----------------------
# GUI
class HIndexApp:
    def __init__(self, root):
        self.root = root
        root.title("HIndex Viewer v3 — Zotero → Semantic Scholar")
        self.api_key_var = tk.StringVar(value="")
        self.delay_var = tk.DoubleVar(value=DEFAULT_DELAY)

        # Notebook
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True)

        # Authors tab
        self.auth_frame = tk.Frame(self.notebook)
        self.notebook.add(self.auth_frame, text="Authors")
        self._build_authors_tab()

        # Coauthors tab
        self.co_frame = tk.Frame(self.notebook)
        self.notebook.add(self.co_frame, text="Co-authors")
        self._build_coauthors_tab()

        # Data
        self.rows = []  # list of dicts for main authors
        self.co_index = {}  # coauthor_name -> metadata dict

        # Populate from cache on start
        self.load_cache_on_startup()

    # ---------- UI builders ----------
    def _build_authors_tab(self):
        frm = tk.Frame(self.auth_frame)
        frm.pack(fill='x', padx=6, pady=6)
        tk.Button(frm, text="Load Zotero CSV", command=self.load_csv).pack(side='left')
        tk.Button(frm, text="Load Zotero BibTeX", command=self.load_bib).pack(side='left')
        tk.Button(frm, text="Fetch Zotero API (opt)", command=self.fetch_zotero_api).pack(side='left')
        tk.Button(frm, text="Refresh Selected", command=self.refresh_selected).pack(side='left')
        tk.Button(frm, text="Refresh All", command=self.refresh_all).pack(side='left')
        tk.Button(frm, text="Export CSV", command=self.export_csv).pack(side='left')
        rightfrm = tk.Frame(frm)
        rightfrm.pack(side='right')
        tk.Label(rightfrm, text="API key (opt):").pack(side='left')
        tk.Entry(rightfrm, textvariable=self.api_key_var, width=30).pack(side='left')
        tk.Label(rightfrm, text="Delay(s):").pack(side='left', padx=(8,0))
        tk.Entry(rightfrm, textvariable=self.delay_var, width=4).pack(side='left')

        sfrm = tk.Frame(self.auth_frame)
        sfrm.pack(fill='x', padx=6)
        tk.Label(sfrm, text="Filter:").pack(side='left')
        self.filter_var = tk.StringVar()
        tk.Entry(sfrm, textvariable=self.filter_var).pack(side='left', fill='x', expand=True)
        self.filter_var.trace_add('write', lambda *_: self.apply_filter())

        cols = ("zotero_name","ss_name","hindex","papers","ss_id","coauthors_preview")
        self.tree = ttk.Treeview(self.auth_frame, columns=cols, show='headings')
        for c in cols:
            self.tree.heading(c, text=c.replace('_',' ').title(), command=lambda _c=c: self.sort_by(_c, False))
            self.tree.column(c, width=160, anchor='w')
        self.tree.pack(fill='both', expand=True, padx=6, pady=6)
        self.tree.bind("<Button-3>", self.on_right_click)

        self.status = tk.Label(self.auth_frame, text="Ready", anchor='w')
        self.status.pack(fill='x')

    def _build_coauthors_tab(self):
        cfrm = tk.Frame(self.co_frame)
        cfrm.pack(fill='x', padx=6, pady=6)
        tk.Button(cfrm, text="Build/Refresh Co-authors", command=self.build_coauthor_index).pack(side='left')
        tk.Button(cfrm, text="Enrich coauthors (fetch SS details)", command=self.enrich_coauthors_with_ss).pack(side='left')
        tk.Button(cfrm, text="Export Coauthors CSV", command=self.export_coauthors_csv).pack(side='left')
        tk.Button(cfrm, text="Open Authors Tab", command=lambda: self.notebook.select(self.auth_frame)).pack(side='left')

        ccols = ("coauthor_name","ss_name","hindex","papers","ss_id","mains","count")
        self.ctree = ttk.Treeview(self.co_frame, columns=ccols, show='headings')
        for c in ccols:
            self.ctree.heading(c, text=c.replace('_',' ').title(), command=lambda _c=c: self.sort_co_by(_c, False))
            self.ctree.column(c, width=160, anchor='w')
        self.ctree.pack(fill='both', expand=True, padx=6, pady=6)

        self.co_status = tk.Label(self.co_frame, text="Co-authors ready", anchor='w')
        self.co_status.pack(fill='x')

    # ---------- Utility UI methods ----------
    def set_status(self, text):
        self.status.config(text=text)
        self.root.update_idletasks()

    def set_co_status(self, text):
        self.co_status.config(text=text)
        self.root.update_idletasks()

    # ---------- Startup cache import ----------
    def load_cache_on_startup(self):
        # Use global cache loaded at module start
        # Build rows from any existing search:: entries in cache
        imported = 0
        for key, val in cache.items():
            if not key.startswith('search::'):
                continue
            zotero_name = key.split('search::',1)[1]
            top = val
            if not top:
                continue
            author_id = top.get('authorId')
            # Try to get author details from cache (author::id)
            details = None
            if author_id and f'author::{author_id}' in cache:
                details = cache.get(f'author::{author_id}')
            # Build row dict
            row = {
                'zotero_name': zotero_name,
                'ss_name': (details.get('name') if details else top.get('name')) if top else '',
                'hindex': (details.get('hIndex') if details else '') if details or top else '',
                'papers': (details.get('paperCount') if details else '') if details or top else '',
                'ss_id': author_id or '',
                'coauthors_list': cache.get(f'coauthors::{author_id}') if author_id else {},
            }
            # Avoid duplicates
            if not any(r['zotero_name'].lower() == row['zotero_name'].lower() for r in self.rows):
                self.rows.append(row)
                imported += 1
        if imported:
            self.refresh_tree()
            self.set_status(f"Loaded {imported} authors from cache")
        else:
            self.set_status("Ready (no cached authors found)")

    # ---------- Zotero importers ----------
    def load_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV files","*.csv"),("All files","*.*")])
        if not path:
            return
        names = parse_zotero_csv(path)
        if not names:
            messagebox.showinfo("No authors","No authors found in that CSV.")
            return
        self.import_names_safely(names)

    def load_bib(self):
        path = filedialog.askopenfilename(filetypes=[("BibTeX files","*.bib"),("All files","*.*")])
        if not path:
            return
        names = parse_bibtex_authors(path)
        if not names:
            messagebox.showinfo("No authors","No authors found in that BibTeX.")
            return
        self.import_names_safely(names)

    def import_names_safely(self, names):
        # Add only names that don't already exist (case-insensitive)
        existing = {r['zotero_name'].lower(): r for r in self.rows}
        added = 0
        for n in names:
            if n.lower() in existing:
                continue
            self.rows.append({
                'zotero_name': n,
                'ss_name': '',
                'hindex': '',
                'papers': '',
                'ss_id': '',
                'coauthors_list': {},
            })
            added += 1
        self.refresh_tree()
        messagebox.showinfo("Import complete", f"Imported {added} new names (skipped {len(names)-added} existing)")

    def fetch_zotero_api(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Zotero API fetch (optional)")
        tk.Label(dlg, text="Zotero user or group ID:").pack()
        uid = tk.Entry(dlg); uid.pack()
        tk.Label(dlg, text="Type (user or group):").pack()
        typ = ttk.Combobox(dlg, values=["user","group"]); typ.set("user"); typ.pack()
        tk.Label(dlg, text="API key (optional):").pack()
        key = tk.Entry(dlg); key.pack()
        def go():
            u = uid.get().strip()
            if not u:
                messagebox.showerror("Need ID","Please provide a Zotero user or group ID.")
                return
            base = f"https://api.zotero.org/{typ.get()}/{u}/items"
            headers = {}
            if key.get().strip():
                headers['Zotero-API-Key'] = key.get().strip()
            params = {'format':'json','limit':100}
            names=set()
            off=0
            self.set_status("Fetching Zotero items...")
            try:
                while True:
                    params['start']=off
                    r = requests.get(base, headers=headers, params=params, timeout=30)
                    if r.status_code != 200:
                        messagebox.showerror("Zotero error", f"Zotero API returned {r.status_code}")
                        break
                    items = r.json()
                    if not items: break
                    for it in items:
                        creators = it.get('data',{}).get('creators') or it.get('creators') or []
                        for c in creators:
                            if not isinstance(c, dict): continue
                            name = c.get('lastName')
                            if name:
                                fn = c.get('firstName') or ''
                                full = f"{fn} {name}".strip()
                                names.add(full)
                            else:
                                nm = c.get('name')
                                if nm:
                                    names.add(nm)
                    off += len(items)
                    if len(items) < params['limit']: break
                dlg.destroy()
                if names:
                    self.import_names_safely(sorted(names))
                else:
                    messagebox.showinfo("No authors", "No author names found in Zotero items.")
            except Exception as e:
                messagebox.showerror("Error", str(e))
            finally:
                self.set_status("Ready")
        tk.Button(dlg, text="Fetch", command=go).pack()

    # ---------- Tree helpers ----------
    def refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        for r in self.rows:
            preview = ", ".join(sorted(list((r.get('coauthors_list') or {}).keys())[:6]))
            vals = (r.get('zotero_name',''), r.get('ss_name',''), r.get('hindex',''), r.get('papers',''), r.get('ss_id',''), preview)
            self.tree.insert('', 'end', values=vals)
        self.apply_filter()

    def apply_filter(self):
        f = self.filter_var.get().lower().strip()
        for iid in self.tree.get_children():
            vals = self.tree.item(iid)['values']
            combined = " ".join([str(v) for v in vals]).lower()
            self.tree.item(iid, tags=() if f in combined else ("hidden",))
        self.tree.tag_configure("hidden", foreground='#999999')

    def sort_by(self, col, descending):
        data = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]
        try:
            data = [(float(d[0]) if d[0] not in [None,""] else float('-inf'), d[1]) for d in data]
        except Exception:
            data = [(d[0], d[1]) for d in data]
        data.sort(reverse=descending)
        for ix, (_, k) in enumerate(data):
            self.tree.move(k, '', ix)
        self.tree.heading(col, command=lambda: self.sort_by(col, not descending))

    def sort_co_by(self, col, descending):
        data = [(self.ctree.set(k, col), k) for k in self.ctree.get_children('')]
        try:
            data = [(int(d[0]) if d[0] not in [None,""] and str(d[0]).isdigit() else d[0], d[1]) for d in data]
        except Exception:
            data = [(d[0], d[1]) for d in data]
        data.sort(reverse=descending)
        for ix, (_, k) in enumerate(data):
            self.ctree.move(k, '', ix)
        self.ctree.heading(col, command=lambda: self.sort_co_by(col, not descending))

    def on_right_click(self, event):
        iid = self.tree.identify_row(event.y)
        if not iid: return
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Open SemanticScholar page", command=lambda: self.open_selected(iid))
        menu.add_command(label="Refresh this author", command=lambda: self.refresh_item(iid))
        menu.post(event.x_root, event.y_root)

    def open_selected(self, iid):
        vals = self.tree.item(iid)['values']
        ss_id = vals[4]
        if not ss_id:
            messagebox.showinfo("No SS id","No Semantic Scholar ID for this row.")
            return
        webbrowser.open(f"https://www.semanticscholar.org/author/{ss_id}")

    # ---------- Refresh logic ----------
    def refresh_item(self, iid):
        vals = self.tree.item(iid)['values']
        zotero_name = vals[0]
        self.set_status(f"Searching {zotero_name}...")
        api_key = self.api_key_var.get().strip() or None
        base_delay = float(self.delay_var.get() or DEFAULT_DELAY)
        top = ss_search_author(zotero_name, api_key=api_key, base_delay=base_delay)
        time.sleep(base_delay)
        if not top:
            messagebox.showinfo("No match", f"No Semantic Scholar match for {zotero_name}")
            self.set_status("Ready")
            return
        author_id = top.get('authorId')
        details = ss_get_author_details(author_id, api_key=api_key, base_delay=base_delay)
        coauthors = ss_get_author_coauthors(author_id, api_key=api_key, max_papers=50, base_delay=base_delay)
        for r in self.rows:
            if r['zotero_name'].lower() == zotero_name.lower():
                r['ss_name'] = (details.get('name') if details else top.get('name',''))
                r['hindex'] = details.get('hIndex','') if details else ''
                r['papers'] = details.get('paperCount','') if details else ''
                r['ss_id'] = author_id or ''
                r['coauthors_list'] = coauthors or {}
                break
        self.refresh_tree()
        self.set_status("Ready")

    def refresh_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Select row","Select a row first (click it) to refresh.")
            return
        for iid in sel:
            self.refresh_item(iid)

    def refresh_all(self):
        api_key = self.api_key_var.get().strip() or None
        base_delay = float(self.delay_var.get() or DEFAULT_DELAY)
        n = len(self.rows)
        if n == 0:
            messagebox.showinfo("No rows","Load some authors first.")
            return
        if n > 80:
            if not messagebox.askyesno("Many authors", f"You are about to query {n} authors. This may hit API rate limits. Continue?"):
                return
        for i, r in enumerate(self.rows, start=1):
            self.set_status(f"({i}/{n}) Searching: {r['zotero_name']}")
            top = ss_search_author(r['zotero_name'], api_key=api_key, base_delay=base_delay)
            time.sleep(base_delay)
            if not top:
                continue
            author_id = top.get('authorId')
            details = ss_get_author_details(author_id, api_key=api_key, base_delay=base_delay)
            coauthors = ss_get_author_coauthors(author_id, api_key=api_key, max_papers=50, base_delay=base_delay)
            r['ss_name'] = (details.get('name') if details else top.get('name',''))
            r['hindex'] = details.get('hIndex','') if details else ''
            r['papers'] = details.get('paperCount','') if details else ''
            r['ss_id'] = author_id or ''
            r['coauthors_list'] = coauthors or {}
            self.refresh_tree()
        self.set_status("Ready")
        # rebuild coauthor index
        self.build_coauthor_index()

    def export_csv(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV","*.csv")])
        if not path: return
        with open(path, 'w', newline='', encoding='utf8') as f:
            w = csv.DictWriter(f, fieldnames=['zotero_name','ss_name','hindex','papers','ss_id','coauthors'])
            w.writeheader()
            for r in self.rows:
                row_out = {k: (v if not isinstance(v, dict) else ", ".join(v.keys())) for k,v in r.items() if k in ['zotero_name','ss_name','hindex','papers','ss_id','coauthors']}
                w.writerow(row_out)
        messagebox.showinfo("Saved", f"Saved {len(self.rows)} rows to {path}.")

    # ---------- Coauthor aggregation & enrichment ----------
    def build_coauthor_index(self):
        self.co_index = {}
        total = len(self.rows)
        for i, r in enumerate(self.rows, start=1):
            self.set_co_status(f"({i}/{total}) Collecting coauthors for: {r['zotero_name']}")
            if not r.get('coauthors_list') and r.get('ss_id'):
                # try fetch if missing
                base_delay = float(self.delay_var.get() or DEFAULT_DELAY)
                co = ss_get_author_coauthors(r['ss_id'], api_key=self.api_key_var.get().strip() or None, max_papers=50, base_delay=base_delay)
                r['coauthors_list'] = co or {}
            for co_name, cnt in (r.get('coauthors_list') or {}).items():
                if not co_name: continue
                if co_name == r.get('ss_name'):
                    continue
                entry = self.co_index.get(co_name, {'count':0, 'mains':set(), 'ss_name':'', 'hindex':'', 'papers':'', 'ss_id':''})
                entry['count'] += cnt
                entry['mains'].add(r['zotero_name'])
                self.co_index[co_name] = entry
        # populate ctree (basic, without SS details yet)
        self.ctree.delete(*self.ctree.get_children())
        for name, meta in sorted(self.co_index.items(), key=lambda x: -x[1]['count']):
            mains = ", ".join(sorted(meta['mains']))
            vals = (name, meta.get('ss_name',''), meta.get('hindex',''), meta.get('papers',''), meta.get('ss_id',''), mains, meta['count'])
            self.ctree.insert('', 'end', values=vals)
        self.set_co_status(f"Built co-author index: {len(self.co_index)} co-authors")

    def enrich_coauthors_with_ss(self):
        # Enrich coauthors with Semantic Scholar search/details, using cache when possible.
        missing = [name for name,meta in self.co_index.items() if not meta.get('ss_id')]
        if not missing:
            messagebox.showinfo("Nothing to do","All coauthors already have SS info in cache.")
            return
        if len(missing) > 60:
            if not messagebox.askyesno("Many coauthors", f"You are about to query {len(missing)} coauthors from Semantic Scholar. This may hit rate limits. Continue?"):
                return
        api_key = self.api_key_var.get().strip() or None
        base_delay = float(self.delay_var.get() or DEFAULT_DELAY)
        for i, name in enumerate(missing, start=1):
            self.set_co_status(f"({i}/{len(missing)}) Looking up coauthor: {name}")
            top = ss_search_author(name, api_key=api_key, base_delay=base_delay)
            time.sleep(base_delay)
            if not top:
                continue
            author_id = top.get('authorId')
            details = ss_get_author_details(author_id, api_key=api_key, base_delay=base_delay)
            meta = self.co_index.get(name)
            if meta is None:
                continue
            meta['ss_name'] = details.get('name') if details else top.get('name','')
            meta['hindex'] = details.get('hIndex','') if details else ''
            meta['papers'] = details.get('paperCount','') if details else ''
            meta['ss_id'] = author_id or ''
            self.co_index[name] = meta
        # refresh ctree
        self.ctree.delete(*self.ctree.get_children())
        for name, meta in sorted(self.co_index.items(), key=lambda x: -x[1]['count']):
            mains = ", ".join(sorted(meta['mains']))
            vals = (name, meta.get('ss_name',''), meta.get('hindex',''), meta.get('papers',''), meta.get('ss_id',''), mains, meta['count'])
            self.ctree.insert('', 'end', values=vals)
        self.set_co_status("Enrichment complete")

    def export_coauthors_csv(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV","*.csv")])
        if not path: return
        with open(path, 'w', newline='', encoding='utf8') as f:
            w = csv.writer(f)
            w.writerow(['coauthor_name','ss_name','hindex','papers','ss_id','main_authors','count'])
            for name, meta in sorted(self.co_index.items(), key=lambda x: -x[1]['count']):
                mains = "; ".join(sorted(meta['mains']))
                w.writerow([name, meta.get('ss_name',''), meta.get('hindex',''), meta.get('papers',''), meta.get('ss_id',''), mains, meta['count']])
        messagebox.showinfo("Saved", f"Saved {len(self.co_index)} co-authors to {path}.")

# -----------------------
# Zotero parsers (same as before)

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

# -----------------------
if __name__ == '__main__':
    cache = load_cache()
    root = tk.Tk()
    app = HIndexApp(root)
    root.geometry('1200x700')
    root.mainloop()
