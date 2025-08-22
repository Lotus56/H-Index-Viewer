"""
hindex_gui.py

GUI front-end for HIndex Viewer v3 — imports ss_utils.py (must be in same folder).

Patches in this version:
  - Visual difference for flagged authors/coauthors (star emoji prefix)
  - Flagged entries are pinned to the top regardless of sort key
  - Flagged entries remain internally sorted among themselves according to the active sort
  - Co-author enrichment no longer flips you back to Authors tab; it keeps or returns you to Co-authors tab after completing

Run:
    py -m pip install --user requests
    py hindex_gui.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import webbrowser, time, csv
import ss_utils as su

class HIndexApp:
    def __init__(self, root):
        self.root = root
        root.title("HIndex Viewer v3 — Zotero → Semantic Scholar")
        self.api_key_var = tk.StringVar(value="")
        self.delay_var = tk.DoubleVar(value=su.DEFAULT_DELAY)

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

        # Watchlist tab
        self.watch_frame = tk.Frame(self.notebook)
        self.notebook.add(self.watch_frame, text="Watchlist")
        self._build_watchlist_tab()

        # Data
        self.rows = []         # main authors
        self.co_index = {}     # coauthor_name -> meta
        self.watchlist = {}

        # Sort state: (col, descending)
        self.current_sort = (None, False)
        self.current_co_sort = (None, False)
        self.current_watch_sort = (None, False)

        # Ensure cache loaded
        su.load_cache()

        # Populate from cache
        self.load_cache_on_startup()

    # UI builders
    def _build_authors_tab(self):
        frm = tk.Frame(self.auth_frame)
        frm.pack(fill='x', padx=6, pady=6)
        tk.Button(frm, text="Load Zotero CSV", command=self.load_csv).pack(side='left')
        tk.Button(frm, text="Load Zotero BibTeX", command=self.load_bib).pack(side='left')
        tk.Button(frm, text="Fetch Zotero API (opt)", command=self.fetch_zotero_api).pack(side='left')
        tk.Button(frm, text="Refresh Selected", command=self.refresh_selected).pack(side='left')
        tk.Button(frm, text="Refresh All", command=self.refresh_all).pack(side='left')
        tk.Button(frm, text="Export CSV", command=self.export_csv).pack(side='left')
        rightfrm = tk.Frame(frm); rightfrm.pack(side='right')
        tk.Label(rightfrm, text="API key (opt):").pack(side='left')
        tk.Entry(rightfrm, textvariable=self.api_key_var, width=30).pack(side='left')
        tk.Label(rightfrm, text="Delay(s):").pack(side='left', padx=(8,0))
        tk.Entry(rightfrm, textvariable=self.delay_var, width=4).pack(side='left')

        sfrm = tk.Frame(self.auth_frame); sfrm.pack(fill='x', padx=6)
        tk.Label(sfrm, text="Filter:").pack(side='left')
        self.filter_var = tk.StringVar()
        tk.Entry(sfrm, textvariable=self.filter_var).pack(side='left', fill='x', expand=True)
        self.filter_var.trace_add('write', lambda *_: self.apply_filter())

        cols = ("zotero_name","ss_name","hindex","papers","ss_id","coauthors_preview")
        self.tree = ttk.Treeview(self.auth_frame, columns=cols, show='headings', selectmode='extended')
        for c in cols:
            self.tree.heading(c, text=c.replace('_',' ').title(), command=lambda _c=c: self.sort_by(_c, False))
            self.tree.column(c, width=160, anchor='w')
        self.tree.pack(fill='both', expand=True, padx=6, pady=6)
        self.tree.bind("<Button-3>", self.on_right_click_authors)

        # configure tags for visual differences
        self.tree.tag_configure('flagged', background='#fff2b2')  # pale yellow background for flagged rows

        self.status = tk.Label(self.auth_frame, text="Ready", anchor='w'); self.status.pack(fill='x')

    def _build_coauthors_tab(self):
        cfrm = tk.Frame(self.co_frame); cfrm.pack(fill='x', padx=6, pady=6)
        tk.Button(cfrm, text="Build/Refresh Co-authors", command=self.build_coauthor_index).pack(side='left')
        tk.Button(cfrm, text="Enrich Selected Coauthors", command=self.enrich_selected_coauthors).pack(side='left')
        tk.Button(cfrm, text="Enrich Coauthors for Selected Main(s)", command=self.enrich_coauthors_for_selected_main).pack(side='left')
        tk.Button(cfrm, text="Enrich All Coauthors", command=self.enrich_coauthors_with_ss).pack(side='left')
        tk.Button(cfrm, text="Export Coauthors CSV", command=self.export_coauthors_csv).pack(side='left')
        tk.Button(cfrm, text="Open Authors Tab", command=lambda: self.notebook.select(self.auth_frame)).pack(side='left')

        ccols = ("coauthor_name","ss_name","hindex","papers","ss_id","mains","count")
        self.ctree = ttk.Treeview(self.co_frame, columns=ccols, show='headings', selectmode='extended')
        for c in ccols:
            self.ctree.heading(c, text=c.replace('_',' ').title(), command=lambda _c=c: self.sort_co_by(_c, False))
            self.ctree.column(c, width=160, anchor='w')
        self.ctree.pack(fill='both', expand=True, padx=6, pady=6)
        self.ctree.bind("<Button-3>", self.on_right_click_coauthors)
        self.ctree.tag_configure('flagged', background='#fff2b2')

        self.co_status = tk.Label(self.co_frame, text="Co-authors ready", anchor='w'); self.co_status.pack(fill='x')

    def _build_watchlist_tab(self):
        wfrm = tk.Frame(self.watch_frame); wfrm.pack(fill='x', padx=6, pady=6)
        tk.Button(wfrm, text="Open Authors Tab", command=lambda: self.notebook.select(self.auth_frame)).pack(side='left')
        tk.Button(wfrm, text="Open Coauthors Tab", command=lambda: self.notebook.select(self.co_frame)).pack(side='left')
        tk.Button(wfrm, text="Remove Selected from Watchlist", command=self.unflag_selected_from_watchlist).pack(side='left')

        wcols = ("name","type","ss_name","hindex","papers","ss_id")
        self.wtree = ttk.Treeview(self.watch_frame, columns=wcols, show='headings', selectmode='extended')
        for c in wcols:
            self.wtree.heading(c, text=c.replace('_',' ').title(), command=lambda _c=c: self.sort_watch_by(_c, False))
            self.wtree.column(c, width=160, anchor='w')
        self.wtree.pack(fill='both', expand=True, padx=6, pady=6)
        self.wtree.bind("<Button-3>", self.on_right_click_watchlist)
        self.wtree.tag_configure('flagged', background='#fff2b2')

        self.watch_status = tk.Label(self.watch_frame, text="Watchlist ready", anchor='w'); self.watch_status.pack(fill='x')

    # Utility UI methods
    def set_status(self, text):
        self.status.config(text=text); self.root.update_idletasks()
    def set_co_status(self, text):
        self.co_status.config(text=text); self.root.update_idletasks()
    def set_watch_status(self, text):
        self.watch_status.config(text=text); self.root.update_idletasks()

    # Startup cache import
    def load_cache_on_startup(self):
        cache = su.cache
        # load watchlist
        wl = cache.get('watchlist', {}) if isinstance(cache, dict) else {}
        self.watchlist = wl if isinstance(wl, dict) else {}
        # import search:: entries
        imported = 0
        for key, val in cache.items():
            if not key.startswith('search::'):
                continue
            zotero_name = key.split('search::',1)[1]
            top = val
            if not top:
                continue
            author_id = top.get('authorId')
            details = None
            if author_id and f'author::{author_id}' in cache:
                details = cache.get(f'author::{author_id}')
            row = {
                'zotero_name': zotero_name,
                'ss_name': (details.get('name') if details else top.get('name')) if top else '',
                'hindex': (details.get('hIndex') if details else '') if details or top else '',
                'papers': (details.get('paperCount') if details else '') if details or top else '',
                'ss_id': author_id or '',
                'coauthors_list': cache.get(f'coauthors::{author_id}') if author_id else {},
                'flagged': zotero_name in self.watchlist,
            }
            if not any(r['zotero_name'].lower() == row['zotero_name'].lower() for r in self.rows):
                self.rows.append(row)
                imported += 1
        # ensure watchlist authors present
        for name, meta in self.watchlist.items():
            if meta.get('type') == 'Author' and not any(r['zotero_name'].lower() == name.lower() for r in self.rows):
                self.rows.append({
                    'zotero_name': name,
                    'ss_name': meta.get('ss_name',''),
                    'hindex': meta.get('hindex',''),
                    'papers': meta.get('papers',''),
                    'ss_id': meta.get('ss_id',''),
                    'coauthors_list': {},
                    'flagged': True,
                })
        if imported:
            self.refresh_tree(); self.set_status(f"Loaded {imported} authors from cache")
        else:
            self.set_status("Ready (no cached authors found)")
        self.rebuild_watchlist_tree()

    # Zotero importers
    def load_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV files","*.csv"),("All files","*.*")])
        if not path: return
        names = su.parse_zotero_csv(path)
        if not names:
            messagebox.showinfo("No authors","No authors found in that CSV.")
            return
        self.import_names_safely(names)

    def load_bib(self):
        path = filedialog.askopenfilename(filetypes=[("BibTeX files","*.bib"),("All files","*.*")])
        if not path: return
        names = su.parse_bibtex_authors(path)
        if not names:
            messagebox.showinfo("No authors","No authors found in that BibTeX.")
            return
        self.import_names_safely(names)

    def import_names_safely(self, names):
        existing = {r['zotero_name'].lower(): r for r in self.rows}
        added = 0
        for n in names:
            if n.lower() in existing: continue
            self.rows.append({'zotero_name': n,'ss_name': '','hindex': '','papers': '','ss_id': '','coauthors_list': {},'flagged': False}); added += 1
        self.refresh_tree(); messagebox.showinfo("Import complete", f"Imported {added} new names (skipped {len(names)-added} existing)")

    def fetch_zotero_api(self):
        dlg = tk.Toplevel(self.root); dlg.title("Zotero API fetch (optional)")
        tk.Label(dlg, text="Zotero user or group ID:").pack(); uid = tk.Entry(dlg); uid.pack()
        tk.Label(dlg, text="Type (user or group):").pack(); typ = ttk.Combobox(dlg, values=["user","group"]); typ.set("user"); typ.pack()
        tk.Label(dlg, text="API key (optional):").pack(); key = tk.Entry(dlg); key.pack()
        def go():
            u = uid.get().strip()
            if not u: messagebox.showerror("Need ID","Please provide a Zotero user or group ID."); return
            base = f"https://api.zotero.org/{typ.get()}/{u}/items"
            headers = {}
            if key.get().strip(): headers['Zotero-API-Key'] = key.get().strip()
            params = {'format':'json','limit':100}
            names=set(); off=0
            self.set_status("Fetching Zotero items...")
            try:
                while True:
                    params['start']=off
                    r = su.requests.get(base, headers=headers, params=params, timeout=30)
                    if r.status_code != 200:
                        messagebox.showerror("Zotero error", f"Zotero API returned {r.status_code}"); break
                    items = r.json()
                    if not items: break
                    for it in items:
                        creators = it.get('data',{}).get('creators') or it.get('creators') or []
                        for c in creators:
                            if not isinstance(c, dict): continue
                            name = c.get('lastName')
                            if name:
                                fn = c.get('firstName') or ''
                                full = f"{fn} {name}".strip(); names.add(full)
                            else:
                                nm = c.get('name')
                                if nm: names.add(nm)
                    off += len(items)
                    if len(items) < params['limit']: break
                dlg.destroy()
                if names: self.import_names_safely(sorted(names))
                else: messagebox.showinfo("No authors", "No author names found in Zotero items.")
            except Exception as e:
                messagebox.showerror("Error", str(e))
            finally:
                self.set_status("Ready")
        tk.Button(dlg, text="Fetch", command=go).pack()

    # Tree helpers
    def refresh_tree(self):
        # Rebuild the Authors tree honoring flagged pinning + current sort
        self.tree.delete(*self.tree.get_children())
        col, desc = self.current_sort
        def sort_key(r):
            # flagged items will be handled in grouping, here return key for sorting within group
            val = r.get(col) if col else r.get('zotero_name')
            if col in ('hindex','papers'):
                try:
                    return float(val) if val not in (None,"") else float('-inf')
                except Exception:
                    return float('-inf')
            return (val or '').lower()
        flagged = [r for r in self.rows if r.get('flagged')]
        others = [r for r in self.rows if not r.get('flagged')]
        if col:
            flagged.sort(key=sort_key, reverse=desc)
            others.sort(key=sort_key, reverse=desc)
        else:
            # default alphabetical
            flagged.sort(key=lambda r: (r.get('zotero_name') or '').lower())
            others.sort(key=lambda r: (r.get('zotero_name') or '').lower())
        ordered = flagged + others
        for r in ordered:
            display_name = ("⭐ " + r.get('zotero_name','')) if r.get('flagged') else r.get('zotero_name','')
            preview = ", ".join(sorted(list((r.get('coauthors_list') or {}).keys())[:6]))
            vals = (display_name, r.get('ss_name',''), r.get('hindex',''), r.get('papers',''), r.get('ss_id',''), preview)
            tags = ('flagged',) if r.get('flagged') else ()
            self.tree.insert('', 'end', values=vals, tags=tags)
        self.apply_filter()

    def apply_filter(self):
        f = self.filter_var.get().lower().strip()
        for iid in self.tree.get_children():
            vals = self.tree.item(iid)['values']
            combined = " ".join([str(v) for v in vals]).lower()
            self.tree.item(iid, tags=self.tree.item(iid).get('tags',()) if f in combined else ("hidden",))
        self.tree.tag_configure("hidden", foreground='#999999')

    def sort_by(self, col, descending):
        self.current_sort = (col, descending)
        self.refresh_tree()
        # toggle next time
        # update heading callback
        self.tree.heading(col, command=lambda: self.sort_by(col, not descending))

    def sort_co_by(self, col, descending):
        self.current_co_sort = (col, descending)
        self.build_coauthor_index()
        self.ctree.heading(col, command=lambda: self.sort_co_by(col, not descending))

    def sort_watch_by(self, col, descending):
        self.current_watch_sort = (col, descending)
        self.rebuild_watchlist_tree()
        self.wtree.heading(col, command=lambda: self.sort_watch_by(col, not descending))

    # Context menus and watchlist management
    def on_right_click_authors(self, event):
        iid = self.tree.identify_row(event.y)
        if not iid: return
        vals = self.tree.item(iid)['values']
        # remove star if present in display
        name = vals[0].lstrip('⭐ ').strip()
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Open SemanticScholar page", command=lambda: webbrowser.open(f"https://www.semanticscholar.org/author/{vals[4]}") if vals[4] else messagebox.showinfo("No SS id","No Semantic Scholar ID for this row."))
        menu.add_command(label="Refresh this author", command=lambda: self.refresh_item(iid))
        is_flagged = any(r for r in self.rows if r['zotero_name'].lower()==name.lower() and r.get('flagged'))
        if not is_flagged:
            menu.add_command(label="Flag as important", command=lambda: self.flag_author(name, source='Author'))
        else:
            menu.add_command(label="Unflag", command=lambda: self.unflag_author(name, source='Author'))
        menu.post(event.x_root, event.y_root)

    def on_right_click_coauthors(self, event):
        iid = self.ctree.identify_row(event.y)
        if not iid: return
        vals = self.ctree.item(iid)['values']
        name = vals[0].lstrip('⭐ ').strip()
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Open SemanticScholar page", command=lambda: webbrowser.open(f"https://www.semanticscholar.org/author/{vals[4]}") if vals[4] else messagebox.showinfo("No SS id","No Semantic Scholar ID for this row."))
        menu.add_command(label="Enrich this coauthor", command=lambda: self.enrich_specific_coauthor(name))
        is_flagged = self.co_index.get(name, {}).get('flagged', False)
        if not is_flagged:
            menu.add_command(label="Flag as important", command=lambda: self.flag_author(name, source='Co-author'))
        else:
            menu.add_command(label="Unflag", command=lambda: self.unflag_author(name, source='Co-author'))
        menu.post(event.x_root, event.y_root)

    def on_right_click_watchlist(self, event):
        iid = self.wtree.identify_row(event.y)
        if not iid: return
        vals = self.wtree.item(iid)['values']
        name = vals[0]
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Open SemanticScholar page", command=lambda: webbrowser.open(f"https://www.semanticscholar.org/author/{vals[5]}") if vals[5] else messagebox.showinfo("No SS id","No Semantic Scholar ID for this row."))
        menu.add_command(label="Remove from Watchlist", command=lambda: self.unflag_selected_from_watchlist())
        menu.post(event.x_root, event.y_root)

    def flag_author(self, name, source='Author'):
        cache = su.cache
        if source == 'Author':
            for r in self.rows:
                if r['zotero_name'].lower() == name.lower():
                    r['flagged'] = True
                    meta = {'type':'Author','ss_name':r.get('ss_name',''),'hindex':r.get('hindex',''),'papers':r.get('papers',''),'ss_id':r.get('ss_id','')}
                    self.watchlist[name] = meta
                    cache['watchlist'] = self.watchlist
                    su.save_cache()
                    self.refresh_tree(); self.rebuild_watchlist_tree(); return
        else:
            meta = self.co_index.get(name)
            if meta is None:
                meta = {'count':0,'mains':set(),'ss_name':'','hindex':'','papers':'','ss_id':'','flagged':True}
            meta['flagged'] = True
            self.co_index[name] = meta
            wlmeta = {'type':'Co-author','ss_name':meta.get('ss_name',''),'hindex':meta.get('hindex',''),'papers':meta.get('papers',''),'ss_id':meta.get('ss_id','')}
            self.watchlist[name] = wlmeta
            cache['watchlist'] = self.watchlist
            su.save_cache()
            self.build_coauthor_index(); self.rebuild_watchlist_tree()

    def unflag_author(self, name, source='Author'):
        cache = su.cache
        if source == 'Author':
            for r in self.rows:
                if r['zotero_name'].lower() == name.lower():
                    r['flagged'] = False
                    if name in self.watchlist: del self.watchlist[name]
                    cache['watchlist'] = self.watchlist; su.save_cache()
                    self.refresh_tree(); self.rebuild_watchlist_tree(); return
        else:
            meta = self.co_index.get(name)
            if meta:
                meta['flagged'] = False
            if name in self.watchlist: del self.watchlist[name]
            cache['watchlist'] = self.watchlist; su.save_cache()
            self.build_coauthor_index(); self.rebuild_watchlist_tree()

    def rebuild_watchlist_tree(self):
        self.wtree.delete(*self.wtree.get_children())
        # show flagged with star
        for name, meta in sorted(self.watchlist.items(), key=lambda x: x[0].lower()):
            display = ("⭐ " + name) if name in [k for k in self.watchlist.keys()] else name
            typ = meta.get('type','Author')
            vals = (display, typ, meta.get('ss_name',''), meta.get('hindex',''), meta.get('papers',''), meta.get('ss_id',''))
            tags = ('flagged',)
            self.wtree.insert('', 'end', values=vals, tags=tags)
        self.set_watch_status(f"Watchlist: {len(self.watchlist)} entries")

    def unflag_selected_from_watchlist(self):
        sel = self.wtree.selection()
        if not sel:
            messagebox.showinfo("Select row","Select a row in the Watchlist to remove/unflag.")
            return
        for iid in sel:
            vals = self.wtree.item(iid)['values']
            name = vals[0].lstrip('⭐ ').strip()
            if name in self.watchlist:
                typ = self.watchlist[name].get('type','Author')
                del self.watchlist[name]
                cache = su.cache
                cache['watchlist'] = self.watchlist; su.save_cache()
                if typ == 'Author':
                    for r in self.rows:
                        if r['zotero_name'].lower() == name.lower(): r['flagged'] = False
                else:
                    if name in self.co_index: self.co_index[name]['flagged'] = False
        self.refresh_tree(); self.build_coauthor_index(); self.rebuild_watchlist_tree()

    # Refresh logic
    def refresh_item(self, iid):
        vals = self.tree.item(iid)['values']; zotero_name = vals[0].lstrip('⭐ ').strip()
        self.set_status(f"Searching {zotero_name}...")
        api_key = self.api_key_var.get().strip() or None
        base_delay = float(self.delay_var.get() or su.DEFAULT_DELAY)
        top = su.ss_search_author(zotero_name, api_key=api_key, base_delay=base_delay)
        time.sleep(base_delay)
        if not top:
            messagebox.showinfo("No match", f"No Semantic Scholar match for {zotero_name}"); self.set_status("Ready"); return
        author_id = top.get('authorId')
        details = su.ss_get_author_details(author_id, api_key=api_key, base_delay=base_delay)
        coauthors = su.ss_get_author_coauthors(author_id, api_key=api_key, max_papers=50, base_delay=base_delay)
        for r in self.rows:
            if r['zotero_name'].lower() == zotero_name.lower():
                r['ss_name'] = (details.get('name') if details else top.get('name',''))
                r['hindex'] = details.get('hIndex','') if details else ''
                r['papers'] = details.get('paperCount','') if details else ''
                r['ss_id'] = author_id or ''
                r['coauthors_list'] = coauthors or {}
                if r.get('flagged'):
                    self.watchlist[r['zotero_name']] = {'type':'Author','ss_name':r.get('ss_name',''),'hindex':r.get('hindex',''),'papers':r.get('papers',''),'ss_id':r.get('ss_id','')}
                    su.cache['watchlist'] = self.watchlist; su.save_cache()
                break
        self.refresh_tree(); self.set_status("Ready")

    def refresh_selected(self):
        sel = self.tree.selection()
        if not sel: messagebox.showinfo("Select row","Select a row first (click it) to refresh."); return
        for iid in sel: self.refresh_item(iid)

    def refresh_all(self):
        api_key = self.api_key_var.get().strip() or None
        base_delay = float(self.delay_var.get() or su.DEFAULT_DELAY)
        n = len(self.rows)
        if n == 0: messagebox.showinfo("No rows","Load some authors first."); return
        if n > 80:
            if not messagebox.askyesno("Many authors", f"You are about to query {n} authors. This may hit API rate limits. Continue?"): return
        for i, r in enumerate(self.rows, start=1):
            self.set_status(f"({i}/{n}) Searching: {r['zotero_name']}")
            top = su.ss_search_author(r['zotero_name'], api_key=api_key, base_delay=base_delay)
            time.sleep(base_delay)
            if not top: continue
            author_id = top.get('authorId')
            details = su.ss_get_author_details(author_id, api_key=api_key, base_delay=base_delay)
            coauthors = su.ss_get_author_coauthors(author_id, api_key=api_key, max_papers=50, base_delay=base_delay)
            r['ss_name'] = (details.get('name') if details else top.get('name',''))
            r['hindex'] = details.get('hIndex','') if details else ''
            r['papers'] = details.get('paperCount','') if details else ''
            r['ss_id'] = author_id or ''
            r['coauthors_list'] = coauthors or {}
            if r.get('flagged'):
                self.watchlist[r['zotero_name']] = {'type':'Author','ss_name':r.get('ss_name',''),'hindex':r.get('hindex',''),'papers':r.get('papers',''),'ss_id':r.get('ss_id','')}
                su.cache['watchlist'] = self.watchlist; su.save_cache()
            self.refresh_tree()
        self.set_status("Ready")
        self.build_coauthor_index()

    def export_csv(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV","*.csv")])
        if not path: return
        with open(path, 'w', newline='', encoding='utf8') as f:
            w = csv.DictWriter(f, fieldnames=['zotero_name','ss_name','hindex','papers','ss_id','coauthors'])
            w.writeheader()
            for r in self.rows:
                row_out = {
                    'zotero_name': r.get('zotero_name',''),
                    'ss_name': r.get('ss_name',''),
                    'hindex': r.get('hindex',''),
                    'papers': r.get('papers',''),
                    'ss_id': r.get('ss_id',''),
                    'coauthors': ", ".join(sorted(list((r.get('coauthors_list') or {}).keys())))
                }
                w.writerow(row_out)
        messagebox.showinfo("Saved", f"Saved {len(self.rows)} rows to {path}.")

    # Coauthor aggregation & enrichment
    def build_coauthor_index(self):
        self.co_index = {}
        total = len(self.rows)
        for i, r in enumerate(self.rows, start=1):
            self.set_co_status(f"({i}/{total}) Collecting coauthors for: {r['zotero_name']}")
            if not r.get('coauthors_list') and r.get('ss_id'):
                base_delay = float(self.delay_var.get() or su.DEFAULT_DELAY)
                co = su.ss_get_author_coauthors(r['ss_id'], api_key=self.api_key_var.get().strip() or None, max_papers=50, base_delay=base_delay)
                r['coauthors_list'] = co or {}
            for co_name, cnt in (r.get('coauthors_list') or {}).items():
                if not co_name: continue
                if co_name == r.get('ss_name'): continue
                entry = self.co_index.get(co_name, {'count':0, 'mains':set(), 'ss_name':'', 'hindex':'', 'papers':'', 'ss_id':'', 'flagged': False})
                entry['count'] += cnt
                entry['mains'].add(r['zotero_name'])
                self.co_index[co_name] = entry
        # populate ctree with pinned flagged first and sorted according to current_co_sort
        self.ctree.delete(*self.ctree.get_children())
        col, desc = self.current_co_sort
        def co_sort_key(item):
            name, meta = item
            if col in ('hindex','papers'):
                try:
                    return float(meta.get(col) if meta.get(col) not in (None,'') else -1)
                except Exception:
                    return -1
            if col:
                return (meta.get(col) or '').lower() if isinstance(meta.get(col,''), str) else meta.get(col)
            return name.lower()
        items = list(self.co_index.items())
        flagged_items = [it for it in items if it[1].get('flagged')]
        other_items = [it for it in items if not it[1].get('flagged')]
        if col:
            flagged_items.sort(key=co_sort_key, reverse=desc)
            other_items.sort(key=co_sort_key, reverse=desc)
        else:
            flagged_items.sort(key=lambda x: x[0].lower())
            other_items.sort(key=lambda x: x[0].lower())
        ordered = flagged_items + other_items
        for name, meta in ordered:
            display = ("⭐ " + name) if meta.get('flagged') else name
            mains = ", ".join(sorted(meta['mains']))
            vals = (display, meta.get('ss_name',''), meta.get('hindex',''), meta.get('papers',''), meta.get('ss_id',''), mains, meta['count'])
            tags = ('flagged',) if meta.get('flagged') else ()
            self.ctree.insert('', 'end', values=vals, tags=tags)
        self.set_co_status(f"Built co-author index: {len(self.co_index)} co-authors")

    def enrich_selected_coauthors(self):
        sel = self.ctree.selection()
        if not sel:
            messagebox.showinfo("Select","Select one or more co-authors in the Co-authors tab to enrich.")
            return
        names = [self.ctree.item(i)['values'][0].lstrip('⭐ ').strip() for i in sel]
        self._enrich_coauthor_names(names)
        # stay on coauthors tab
        self.notebook.select(self.co_frame)

    def enrich_coauthors_for_selected_main(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Select","Select one or more main authors in the Authors tab to target their coauthors.")
            return
        target_mains = [self.tree.item(i)['values'][0].lstrip('⭐ ').strip() for i in sel]
        names = set()
        for name, meta in self.co_index.items():
            if any(m in meta['mains'] for m in target_mains): names.add(name)
        if not names:
            messagebox.showinfo("None","No coauthors found for the selected mains.")
            return
        self._enrich_coauthor_names(sorted(names))
        self.notebook.select(self.co_frame)

    def _enrich_coauthor_names(self, names):
        api_key = self.api_key_var.get().strip() or None
        base_delay = float(self.delay_var.get() or su.DEFAULT_DELAY)
        if len(names) > 80:
            if not messagebox.askyesno("Many coauthors", f"You are about to query {len(names)} coauthors. This may hit rate limits. Continue?"):
                return
        for i, name in enumerate(names, start=1):
            self.set_co_status(f"({i}/{len(names)}) Enriching coauthor: {name}")
            top = su.ss_search_author(name, api_key=api_key, base_delay=base_delay)
            time.sleep(base_delay)
            if not top: continue
            author_id = top.get('authorId')
            details = su.ss_get_author_details(author_id, api_key=api_key, base_delay=base_delay)
            meta = self.co_index.get(name) or {'count':0,'mains':set(),'ss_name':'','hindex':'','papers':'','ss_id':'','flagged': False}
            meta['ss_name'] = details.get('name') if details else top.get('name','')
            meta['hindex'] = details.get('hIndex','') if details else ''
            meta['papers'] = details.get('paperCount','') if details else ''
            meta['ss_id'] = author_id or ''
            self.co_index[name] = meta
            if meta.get('flagged'):
                self.watchlist[name] = {'type':'Co-author','ss_name':meta.get('ss_name',''),'hindex':meta.get('hindex',''),'papers':meta.get('papers',''),'ss_id':meta.get('ss_id','')}
                su.cache['watchlist'] = self.watchlist; su.save_cache()
        # rebuild and stay on coauthors
        self.build_coauthor_index(); self.rebuild_watchlist_tree(); self.set_co_status("Enrichment complete")
        self.notebook.select(self.co_frame)

    def enrich_coauthors_with_ss(self):
        names = list(self.co_index.keys())
        if not names:
            messagebox.showinfo("None","No coauthors to enrich. Build/refresh coauthors first.")
            return
        if len(names) > 80:
            if not messagebox.askyesno("Many coauthors", f"You are about to enrich {len(names)} coauthors. This may hit rate limits. Continue?"):
                return
        self._enrich_coauthor_names(sorted(names))
        self.notebook.select(self.co_frame)

    def enrich_specific_coauthor(self, name):
        self._enrich_coauthor_names([name])
        self.notebook.select(self.co_frame)

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

# Main
if __name__ == '__main__':
    root = tk.Tk()
    app = HIndexApp(root)
    root.geometry('1200x700')
    root.mainloop()
