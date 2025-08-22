# hindex_gui_v3.py
"""
HIndex Viewer v3 — GUI (fixed initialization order)

Save this file next to ss_utils.py and run:
    py -m pip install --user requests
    py hindex_gui_v3.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
import webbrowser, time, csv, threading
import ss_utils as su

class HIndexApp:
    def __init__(self, root):
        self.root = root
        root.title("HIndex Viewer v3 — Zotero → Semantic Scholar")

        # basic vars
        self.api_key_var = tk.StringVar(value="")
        self.delay_var = tk.DoubleVar(value=su.DEFAULT_DELAY)

        # Data containers
        self.rows = []         # main authors
        self.co_index = {}     # coauthor_name -> meta
        self.watchlist = {}

        # Sort state: (col, descending)
        self.current_sort = (None, False)
        self.current_co_sort = (None, False)
        self.current_watch_sort = (None, False)

        # Load cache and settings BEFORE building UI so tag_configure uses flag_color
        su.load_cache()
        settings = su.cache.get('settings', {})
        self.flag_emoji = settings.get('emoji', '⭐')
        self.flag_color = settings.get('color', '#fff2b2')

        # Notebook + tabs (build UI)
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

        # Menu -> Settings (after UI exists)
        menubar = tk.Menu(root)
        root.config(menu=menubar)
        appmenu = tk.Menu(menubar, tearoff=0)
        appmenu.add_command(label="Settings", command=self.open_settings)
        menubar.add_cascade(label="App", menu=appmenu)

        # Apply style to tree tags now that trees exist
        self._apply_flag_style()

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

        # configure tags for visual differences (color already set during init)
        self.tree.tag_configure('flagged', background=self.flag_color)

        self.status = tk.Label(self.auth_frame, text="Ready", anchor='w'); self.status.pack(fill='x')

    def _build_coauthors_tab(self):
        cfrm = tk.Frame(self.co_frame); cfrm.pack(fill='x', padx=6, pady=6)
        tk.Button(cfrm, text="Build/Refresh Co-authors", command=self.threaded_build_coauthor_index).pack(side='left')
        tk.Button(cfrm, text="Enrich Selected Coauthors", command=self.threaded_enrich_selected_coauthors).pack(side='left')
        tk.Button(cfrm, text="Enrich Coauthors for Selected Main(s)", command=self.threaded_enrich_coauthors_for_selected_main).pack(side='left')
        tk.Button(cfrm, text="Enrich All Coauthors", command=self.threaded_enrich_coauthors_with_ss).pack(side='left')
        tk.Button(cfrm, text="Export Coauthors CSV", command=self.export_coauthors_csv).pack(side='left')
        tk.Button(cfrm, text="Open Authors Tab", command=lambda: self.notebook.select(self.auth_frame)).pack(side='left')

        ccols = ("coauthor_name","ss_name","hindex","papers","ss_id","mains","count")
        self.ctree = ttk.Treeview(self.co_frame, columns=ccols, show='headings', selectmode='extended')
        for c in ccols:
            self.ctree.heading(c, text=c.replace('_',' ').title(), command=lambda _c=c: self.sort_co_by(_c, False))
            self.ctree.column(c, width=160, anchor='w')
        self.ctree.pack(fill='both', expand=True, padx=6, pady=6)
        self.ctree.bind("<Button-3>", self.on_right_click_coauthors)
        self.ctree.tag_configure('flagged', background=self.flag_color)

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
        self.wtree.tag_configure('flagged', background=self.flag_color)

        self.watch_status = tk.Label(self.watch_frame, text="Watchlist ready", anchor='w'); self.watch_status.pack(fill='x')

    # Styling util
    def _apply_flag_style(self):
        try:
            if hasattr(self, 'tree'):
                self.tree.tag_configure('flagged', background=self.flag_color)
            if hasattr(self, 'ctree'):
                self.ctree.tag_configure('flagged', background=self.flag_color)
            if hasattr(self, 'wtree'):
                self.wtree.tag_configure('flagged', background=self.flag_color)
        except Exception:
            pass

    # Utility UI methods
    def set_status(self, text):
        self.status.config(text=text); self.root.update_idletasks()
    def set_co_status(self, text):
        self.co_status.config(text=text); self.root.update_idletasks()
    def set_watch_status(self, text):
        self.watch_status.config(text=text); self.root.update_idletasks()

    # Startup cache import
    def load_cache_on_startup(self):
        """
        Load only Zotero-imported authors into self.rows (from su.cache['zotero_names'])
        and reconstruct coauthor index from su.cache 'coauthors::{authorId}' entries.
        This prevents arbitrary 'search::' cache entries (e.g. coauthors) from
        being loaded as main authors.
        """
        cache = su.cache or {}

        # Load watchlist as before
        wl = cache.get('watchlist', {}) if isinstance(cache, dict) else {}
        self.watchlist = wl if isinstance(wl, dict) else {}

        # Get explicit Zotero-imported names (persisted during import)
        zotero_list = cache.get('zotero_names', []) if isinstance(cache, dict) else []
        imported = 0

        # Build helper mapping: authorId -> zotero_name (if available in cache)
        authorid_to_name = {}
        # If there are explicit zotero names, try to map search:: entries to ids
        for key, val in cache.items():
            if key.startswith('search::') and isinstance(val, dict):
                name = key.split('search::', 1)[1]
                aid = val.get('authorId')
                if aid:
                    # prefer mapping to the explicit Zotero name if matching
                    authorid_to_name[aid] = name

            # also check author::{id} entries for a canonical name
            if key.startswith('author::') and isinstance(val, dict):
                aid = key.split('author::',1)[1]
                nm = val.get('name')
                if aid and nm:
                    authorid_to_name[aid] = nm

        # Import main authors only if they are listed in zotero_names
        for name in zotero_list:
            top = cache.get(f'search::{name}', None)
            author_id = None
            details = None
            if top and isinstance(top, dict):
                author_id = top.get('authorId')
                if author_id and f'author::{author_id}' in cache:
                    details = cache.get(f'author::{author_id}')
            row = {
                'zotero_name': name,
                'ss_name': (details.get('name') if details else (top.get('name') if top else '')) if (details or top) else '',
                'hindex': (details.get('hIndex') if details else '') if (details or top) else '',
                'papers': (details.get('paperCount') if details else '') if (details or top) else '',
                'ss_id': author_id or '',
                'coauthors_list': cache.get(f'coauthors::{author_id}') if author_id else {},
                'flagged': name in self.watchlist,
            }
            if not any(r['zotero_name'].lower() == row['zotero_name'].lower() for r in self.rows):
                self.rows.append(row)
                imported += 1

        # Reconstruct coauthor index from any cached coauthors::{authorId} entries
        self.co_index = {}
        for key, val in cache.items():
            if not isinstance(key, str): continue
            if not key.startswith('coauthors::'): continue
            # value expected to be dict of {coauthor_name: count}
            try:
                aid = key.split('coauthors::',1)[1]
            except Exception:
                aid = None
            if not isinstance(val, dict):
                continue
            main_name = authorid_to_name.get(aid, aid or '(unknown main)')
            for co_name, cnt in val.items():
                if not co_name: continue
                entry = self.co_index.get(co_name, {'count':0, 'mains':set(), 'ss_name':'', 'hindex':'', 'papers':'', 'ss_id':'', 'flagged': False})
                entry['count'] += int(cnt or 0)
                entry['mains'].add(main_name)
                self.co_index[co_name] = entry

        # ensure watchlist authors (type=Author) present in rows
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
            self.refresh_tree()
            self.set_status(f"Loaded {imported} authors from cache")
        else:
            self.set_status("Ready (no cached Zotero authors found)")

        # Populate coauthors tree from reconstructed model
        self.build_coauthor_index()
        self.rebuild_watchlist_tree()

    # Settings dialog
    def open_settings(self):
        dlg = tk.Toplevel(self.root); dlg.title("Settings")
        tk.Label(dlg, text="Flag emoji: (e.g. ⭐)").pack()
        emoji_entry = tk.Entry(dlg); emoji_entry.insert(0, self.flag_emoji); emoji_entry.pack()
        def pick_color():
            c = colorchooser.askcolor(color=self.flag_color, title="Choose flagged row color")
            if c and c[1]:
                color_entry.delete(0, 'end'); color_entry.insert(0, c[1])
        tk.Label(dlg, text="Flag background color (hex):").pack()
        color_entry = tk.Entry(dlg); color_entry.insert(0, self.flag_color); color_entry.pack()
        tk.Button(dlg, text="Choose color...", command=pick_color).pack()
        def save_and_close():
            self.flag_emoji = emoji_entry.get().strip() or '⭐'
            self.flag_color = color_entry.get().strip() or '#fff2b2'
            # persist
            su.cache.setdefault('settings', {})['emoji'] = self.flag_emoji
            su.cache.setdefault('settings', {})['color'] = self.flag_color
            su.save_cache()
            self._apply_flag_style()
            # refresh displays to show new emoji/color
            self.refresh_tree(); self.build_coauthor_index(); self.rebuild_watchlist_tree()
            dlg.destroy()
        tk.Button(dlg, text="Save", command=save_and_close).pack()

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
        """
        Add new names to self.rows but also persist the list of Zotero-imported names
        into su.cache['zotero_names'] so they are reloaded on startup. This prevents
        accidental import of coauthor 'search::' entries as main authors later.
        """
        existing = {r['zotero_name'].lower(): r for r in self.rows}
        added = 0
        for n in names:
            if n.lower() in existing: continue
            self.rows.append({'zotero_name': n,'ss_name': '','hindex': '','papers': '','ss_id': '','coauthors_list': {},'flagged': False})
            added += 1

        # persist Zotero-imported names for future startups
        current_zotero = su.cache.get('zotero_names', [])
        # merge (preserve order as much as possible)
        for n in names:
            if n not in current_zotero:
                current_zotero.append(n)
        su.cache['zotero_names'] = current_zotero
        su.save_cache()

        self.refresh_tree()
        messagebox.showinfo("Import complete", f"Imported {added} new names (skipped {len(names)-added} existing)")

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

    # THREADING HELPERS - run heavy/network tasks off the main thread
    def _run_in_thread(self, target, on_done=None):
        def runner():
            try:
                target()
            except Exception as e:
                print("Background thread error:", e)
            finally:
                if on_done:
                    try:
                        self.root.after(10, on_done)
                    except Exception:
                        pass
        t = threading.Thread(target=runner, daemon=True)
        t.start()

    # Tree helpers
    def refresh_tree(self):
        # Rebuild the Authors tree honoring flagged pinning + current sort
        self.tree.delete(*self.tree.get_children())
        col, desc = self.current_sort
        def sort_key(r):
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
            flagged.sort(key=lambda r: (r.get('zotero_name') or '').lower())
            others.sort(key=lambda r: (r.get('zotero_name') or '').lower())
        ordered = flagged + others
        for r in ordered:
            display_name = (self.flag_emoji + ' ' + r.get('zotero_name','')) if r.get('flagged') else r.get('zotero_name','')
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
        display_name = vals[0]
        name = display_name
        if isinstance(display_name, str) and display_name.startswith(self.flag_emoji):
            name = display_name[len(self.flag_emoji):].strip()
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Open SemanticScholar page", command=lambda: webbrowser.open(f"https://www.semanticscholar.org/author/{vals[4]}") if vals[4] else messagebox.showinfo("No SS id","No Semantic Scholar ID for this row."))
        menu.add_command(label="Refresh this author", command=lambda: self.threaded_refresh_item(iid))
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
        display_name = vals[0]
        name = display_name
        if isinstance(display_name, str) and display_name.startswith(self.flag_emoji):
            name = display_name[len(self.flag_emoji):].strip()
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Open SemanticScholar page", command=lambda: webbrowser.open(f"https://www.semanticscholar.org/author/{vals[4]}") if vals[4] else messagebox.showinfo("No SS id","No Semantic Scholar ID for this row."))
        menu.add_command(label="Enrich this coauthor", command=lambda: self.threaded_enrich_specific_coauthor(name))
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
            # update ctree in main thread
            self.root.after(10, self.build_coauthor_index)
            self.root.after(10, self.rebuild_watchlist_tree)

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
            self.root.after(10, self.build_coauthor_index); self.root.after(10, self.rebuild_watchlist_tree)

    def rebuild_watchlist_tree(self):
        self.wtree.delete(*self.wtree.get_children())
        # show flagged with emoji
        for name, meta in sorted(self.watchlist.items(), key=lambda x: x[0].lower()):
            display = (self.flag_emoji + ' ' + name)
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
            name = vals[0].lstrip(self.flag_emoji + ' ').strip()
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

    # Refresh logic (threaded wrappers)
    def threaded_refresh_item(self, iid):
        self._run_in_thread(lambda: self._refresh_item_worker(iid), on_done=lambda: self.root.after(10, self.refresh_tree))

    def _refresh_item_worker(self, iid):
        vals = self.tree.item(iid)['values']; zotero_name = vals[0].lstrip(self.flag_emoji + ' ').strip()
        self.root.after(10, lambda: self.set_status(f"Searching {zotero_name}..."))
        api_key = self.api_key_var.get().strip() or None
        base_delay = float(self.delay_var.get() or su.DEFAULT_DELAY)
        top = su.ss_search_author(zotero_name, api_key=api_key, base_delay=base_delay)
        time.sleep(base_delay)
        if not top:
            self.root.after(10, lambda: messagebox.showinfo("No match", f"No Semantic Scholar match for {zotero_name}")); self.root.after(10, lambda: self.set_status("Ready")); return
        author_id = top.get('authorId')
        details = su.ss_get_author_details(author_id, api_key=api_key, base_delay=base_delay)
        coauthors = su.ss_get_author_coauthors(author_id, api_key=api_key, max_papers=50, base_delay=base_delay)
        # update model
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
        self.root.after(10, lambda: self.set_status("Ready"))

    def refresh_selected(self):
        sel = self.tree.selection()
        if not sel: messagebox.showinfo("Select row","Select a row first (click it) to refresh."); return
        for iid in sel: self.threaded_refresh_item(iid)

    def refresh_all(self):
        self._run_in_thread(self._refresh_all_worker, on_done=lambda: (self.root.after(10, self.refresh_tree), self.root.after(10, self.build_coauthor_index)))

    def _refresh_all_worker(self):
        api_key = self.api_key_var.get().strip() or None
        base_delay = float(self.delay_var.get() or su.DEFAULT_DELAY)
        n = len(self.rows)
        if n == 0:
            self.root.after(10, lambda: messagebox.showinfo("No rows","Load some authors first.")); return
        if n > 80:
            # ask on GUI thread - here we assume user already agreed
            pass
        for i, r in enumerate(self.rows, start=1):
            self.root.after(10, lambda i=i, r=r: self.set_status(f"({i}/{n}) Searching: {r['zotero_name']}"))
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

    # Coauthor aggregation & enrichment (threaded)
    def threaded_build_coauthor_index(self):
        self._run_in_thread(self._build_coauthor_index_worker, on_done=lambda: self.root.after(10, self.build_coauthor_index))

    def _build_coauthor_index_worker(self):
        # ensure coauthors are collected into the model (but do not modify GUI directly)
        total = len(self.rows)
        for i, r in enumerate(self.rows, start=1):
            self.root.after(10, lambda i=i, r=r: self.set_co_status(f"({i}/{total}) Collecting coauthors for: {r['zotero_name']}"))
            if not r.get('coauthors_list') and r.get('ss_id'):
                base_delay = float(self.delay_var.get() or su.DEFAULT_DELAY)
                co = su.ss_get_author_coauthors(r['ss_id'], api_key=self.api_key_var.get().strip() or None, max_papers=50, base_delay=base_delay)
                r['coauthors_list'] = co or {}
        self.root.after(10, lambda: self.set_co_status("Collected coauthors (ready to build index)"))

    def build_coauthor_index(self):
        # build index from model and populate ctree - called on main thread
        self.co_index = {}
        for r in self.rows:
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
            display = (self.flag_emoji + ' ' + name) if meta.get('flagged') else name
            mains = ", ".join(sorted(meta['mains']))
            vals = (display, meta.get('ss_name',''), meta.get('hindex',''), meta.get('papers',''), meta.get('ss_id',''), mains, meta['count'])
            tags = ('flagged',) if meta.get('flagged') else ()
            self.ctree.insert('', 'end', values=vals, tags=tags)
        self.set_co_status(f"Built co-author index: {len(self.co_index)} co-authors")

    def threaded_enrich_selected_coauthors(self):
        sel = self.ctree.selection()
        if not sel:
            messagebox.showinfo("Select","Select one or more co-authors in the Co-authors tab to enrich.")
            return
        names = [self.ctree.item(i)['values'][0].lstrip(self.flag_emoji + ' ').strip() for i in sel]
        self._run_in_thread(lambda: self._enrich_coauthor_names_worker(names),
                            on_done=lambda: (self.root.after(10, self.build_coauthor_index),
                                             self.root.after(10, self.rebuild_watchlist_tree),
                                             self.root.after(10, lambda: self.notebook.select(self.co_frame))))

    def threaded_enrich_coauthors_for_selected_main(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Select","Select one or more main authors in the Authors tab to target their coauthors.")
            return
        target_mains = [self.tree.item(i)['values'][0].lstrip(self.flag_emoji + ' ').strip() for i in sel]
        names = set()
        for name, meta in self.co_index.items():
            if any(m in meta['mains'] for m in target_mains): names.add(name)
        if not names:
            messagebox.showinfo("None","No coauthors found for the selected mains.")
            return
        self._run_in_thread(lambda: self._enrich_coauthor_names_worker(sorted(names)),
                            on_done=lambda: (self.root.after(10, self.build_coauthor_index),
                                             self.root.after(10, self.rebuild_watchlist_tree),
                                             self.root.after(10, lambda: self.notebook.select(self.co_frame))))

    def threaded_enrich_coauthors_with_ss(self):
        names = list(self.co_index.keys())
        if not names:
            messagebox.showinfo("None","No coauthors to enrich. Build/refresh coauthors first.")
            return
        if len(names) > 80:
            if not messagebox.askyesno("Many coauthors", f"You are about to enrich {len(names)} coauthors. This may hit rate limits. Continue?"):
                return
        self._run_in_thread(lambda: self._enrich_coauthor_names_worker(sorted(names)),
                            on_done=lambda: (self.root.after(10, self.build_coauthor_index),
                                             self.root.after(10, self.rebuild_watchlist_tree),
                                             self.root.after(10, lambda: self.notebook.select(self.co_frame))))

    def threaded_enrich_specific_coauthor(self, name):
        self._run_in_thread(lambda: self._enrich_coauthor_names_worker([name]),
                            on_done=lambda: (self.root.after(10, self.build_coauthor_index),
                                             self.root.after(10, self.rebuild_watchlist_tree),
                                             self.root.after(10, lambda: self.notebook.select(self.co_frame))))

    def _enrich_coauthor_names_worker(self, names):
        api_key = self.api_key_var.get().strip() or None
        base_delay = float(self.delay_var.get() or su.DEFAULT_DELAY)
        for i, name in enumerate(names, start=1):
            self.root.after(10, lambda i=i, name=name: self.set_co_status(f"({i}/{len(names)}) Enriching coauthor: {name}"))
            top = su.ss_search_author(name, api_key=api_key, base_delay=base_delay)
            time.sleep(base_delay)
            if not top:
                continue
            author_id = top.get('authorId')
            details = su.ss_get_author_details(author_id, api_key=api_key, base_delay=base_delay)
            meta = self.co_index.get(name) or {'count':0,'mains':set(),'ss_name':'','hindex':'','papers':'','ss_id':'','flagged': False}
            meta['ss_name'] = details.get('name') if details else top.get('name','')
            meta['hindex'] = details.get('hIndex','') if details else ''
            meta['papers'] = details.get('paperCount','') if details else ''
            meta['ss_id'] = author_id or ''
            # only update co_index (do NOT add to self.rows)
            self.co_index[name] = meta
            if meta.get('flagged'):
                self.watchlist[name] = {'type':'Co-author','ss_name':meta.get('ss_name',''),'hindex':meta.get('hindex',''),'papers':meta.get('papers',''),'ss_id':meta.get('ss_id','')}
                su.cache['watchlist'] = self.watchlist; su.save_cache()
        self.root.after(10, lambda: self.set_co_status("Enrichment complete"))

    def enrich_specific_coauthor(self, name):
        self.threaded_enrich_specific_coauthor(name)

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

if __name__ == '__main__':
    root = tk.Tk()
    app = HIndexApp(root)
    root.geometry('1200x700')
    root.mainloop()
