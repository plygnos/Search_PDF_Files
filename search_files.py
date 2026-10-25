"""Αναζήτηση περιεχομένου σε φακέλους Windows. Python 3.10+."""
from __future__ import annotations

import os
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

EXTENSIONS = {'.txt', '.md', '.csv', '.log', '.json', '.xml', '.html', '.htm',
              '.docx', '.docm', '.pdf', '.rtf', '.xlsx', '.xlsm', '.pptx'}
OFFICE_NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
IGNORED_DIRS = {'$recycle.bin', 'system volume information', '__pycache__', '.git', 'node_modules'}
MAX_PREVIEW = 2500


def normalize(s: str) -> str:
    """Case- and accent-insensitive for Greek and Latin; preserve spaces."""
    s = unicodedata.normalize('NFD', s.casefold())
    return ''.join(c for c in s if unicodedata.category(c) != 'Mn').replace('ς', 'σ')


def extract_text_file(path: Path) -> str:
    for enc in ('utf-8-sig', 'utf-16', 'cp1253', 'cp1252'):
        try:
            return path.read_text(encoding=enc)
        except (UnicodeError, LookupError):
            continue
    return path.read_text(encoding='utf-8', errors='replace')


def office_text(path: Path) -> str:
    """Parse OOXML files without Microsoft Office; only visible textual XML nodes."""
    ext = path.suffix.lower()
    collected = []
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if ext in ('.docx', '.docm'):
                wanted = name == 'word/document.xml' or (
                    name.startswith('word/') and re.search(r'(header|footer|footnotes|endnotes)\d*\.xml$', name))
                element_tags = {'t', 'tab', 'br', 'cr', 'p'}
            elif ext in ('.xlsx', '.xlsm'):
                wanted = name == 'xl/sharedStrings.xml' or name.startswith('xl/worksheets/sheet') and name.endswith('.xml')
                element_tags = {'t', 'v', 'is', 'row'}
            else:
                wanted = name.startswith('ppt/slides/slide') and name.endswith('.xml')
                element_tags = {'t', 'p'}
            if not wanted:
                continue
            try:
                root = ET.fromstring(z.read(name))
            except (ET.ParseError, KeyError):
                continue
            parts = []
            for el in root.iter():
                tag = el.tag.rsplit('}', 1)[-1]
                if tag not in element_tags:
                    continue
                if tag in ('p', 'row', 'br', 'cr'):
                    parts.append('\n')
                elif tag == 'tab':
                    parts.append('\t')
                elif el.text:
                    parts.append(el.text)
            collected.append(''.join(parts))
    return '\n'.join(collected)


def rtf_text(path: Path) -> str:
    try:
        from striprtf.striprtf import rtf_to_text
        return rtf_to_text(extract_text_file(path))
    except ImportError:
        raise RuntimeError('Για αρχεία RTF χρειάζεται η βιβλιοθήκη striprtf.')


def pdf_text(path: Path) -> str:
    try:
        import fitz  # PyMuPDF
    except ImportError:
        raise RuntimeError('Για αρχεία PDF χρειάζεται η βιβλιοθήκη pymupdf.')
    with fitz.open(str(path)) as doc:
        if doc.needs_pass:
            raise RuntimeError('Το PDF προστατεύεται με κωδικό.')
        return '\n'.join(page.get_text(sort=True) for page in doc)


def read_content(path: Path) -> str:
    ext = path.suffix.lower()
    if ext == '.pdf':
        return pdf_text(path)
    if ext in {'.docx', '.docm', '.xlsx', '.xlsm', '.pptx'}:
        return office_text(path)
    if ext == '.rtf':
        return rtf_text(path)
    return extract_text_file(path)


def find_snippet(text: str, terms: list[str]) -> str:
    # Regex search in normalized text retains approximate character offsets.
    normalized = normalize(text)
    positions = [normalized.find(t) for t in terms]
    positions = [p for p in positions if p >= 0]
    pos = min(positions) if positions else 0
    start, end = max(0, pos - 85), min(len(text), pos + 220)
    excerpt = ' '.join(text[start:end].split())
    return ('…' if start else '') + excerpt + ('…' if end < len(text) else '')


@dataclass
class Match:
    path: Path
    snippet: str
    size: int


def scan(folder: Path, words: list[str], require_all: bool, recursive: bool, cancel: threading.Event, output: queue.Queue):
    checked, found, skipped, issues = 0, 0, 0, []
    try:
        for root, dirs, files in os.walk(folder, onerror=lambda e: issues.append(str(e))):
            if cancel.is_set():
                break
            dirs[:] = sorted(d for d in dirs if d.lower() not in IGNORED_DIRS) if recursive else []
            for name in sorted(files):
                if cancel.is_set():
                    break
                path = Path(root) / name
                if path.suffix.lower() not in EXTENSIONS or name.startswith('~$'):
                    continue
                checked += 1
                try:
                    # Αναζήτηση πρώτα στο όνομα του αρχείου
                    name_norm = normalize(name)
                    name_hits = [word in name_norm for word in words]
                    if (all(name_hits) if require_all else any(name_hits)):
                        found += 1
                        output.put(('match', Match(path, "[Βρέθηκε στο όνομα του αρχείου]", path.stat().st_size)))
                        continue

                    text = read_content(path)
                    normalized = normalize(text)
                    hits = [word in normalized for word in words]
                    if (all(hits) if require_all else any(hits)):
                        found += 1
                        output.put(('match', Match(path, find_snippet(text, words), path.stat().st_size)))
                except (OSError, RuntimeError, UnicodeError, ValueError, zipfile.BadZipFile, ET.ParseError) as e:
                    skipped += 1
                    if len(issues) < 40:
                        issues.append(f'{path.name}: {e}')
                if checked % 10 == 0:
                    output.put(('progress', (checked, found, skipped, str(path))))
        output.put(('done', (checked, found, skipped, cancel.is_set(), issues)))
    except (OSError, RuntimeError, ValueError) as e:
        output.put(('fatal', str(e)))


class SearchApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title('Αναζήτηση μέσα σε αρχεία')
        self.root.geometry('1080x690')
        self.root.minsize(820, 540)
        self.root.option_add('*Font', 'Arial 10')
        self.folder = tk.StringVar(value=str(Path.home() / 'Documents'))
        self.keywords = tk.StringVar()
        self.all_terms = tk.BooleanVar(value=True)
        self.recursive = tk.BooleanVar(value=True)
        self.status = tk.StringVar(value='Επίλεξε φάκελο και γράψε τις λέξεις που ψάχνεις.')
        self.q = queue.Queue()
        self.cancel = threading.Event()
        self.running = False
        self.paths:dict[str, Path] = {}
        self.excerpts: dict[str, str] = {}
        self.search_terms: list[str] = []
        self._create_widgets()
        self.root.after(100, self.poll)

    def _create_widgets(self):
        style = ttk.Style()
        if 'vista' in style.theme_names():
            style.theme_use('vista')
        shell = ttk.Frame(self.root, padding=18)
        shell.pack(fill='both', expand=True)
        ttk.Label(shell, text='Αναζήτηση μέσα σε αρχεία', font=('Arial', 19, 'bold')).pack(anchor='w', pady=(0, 14))
        folder_frame = ttk.Frame(shell)
        folder_frame.pack(fill='x', pady=3)
        ttk.Label(folder_frame, text='Φάκελος:', width=13).pack(side='left')
        ttk.Entry(folder_frame, textvariable=self.folder).pack(side='left', fill='x', expand=True, padx=(0, 8))
        ttk.Button(folder_frame, text='Επιλογή…', command=self.pick_folder).pack(side='left')
        terms_frame = ttk.Frame(shell)
        terms_frame.pack(fill='x', pady=(10, 3))
        ttk.Label(terms_frame, text='Λέξεις / φράση:', width=13).pack(side='left')
        terms_entry = ttk.Entry(terms_frame, textvariable=self.keywords)
        terms_entry.pack(side='left', fill='x', expand=True, padx=(0, 8))
        terms_entry.bind('<Return>', lambda e: self.begin())
        self.start_btn = ttk.Button(terms_frame, text='Αναζήτηση', command=self.begin, width=16)
        self.start_btn.pack(side='left')
        options = ttk.Frame(shell)
        options.pack(fill='x', pady=(9, 15))
        ttk.Label(options, text='', width=13).pack(side='left')
        ttk.Radiobutton(options, text='Όλες οι λέξεις', variable=self.all_terms, value=True).pack(side='left', padx=(0, 14))
        ttk.Radiobutton(options, text='Οποιαδήποτε λέξη', variable=self.all_terms, value=False).pack(side='left', padx=(0, 22))
        ttk.Checkbutton(options, text='Και στους υποφακέλους', variable=self.recursive).pack(side='left')
        self.stop_btn = ttk.Button(options, text='Διακοπή', command=self.stop, state='disabled')
        self.stop_btn.pack(side='right')
        ttk.Label(shell, text='Αποτελέσματα (διπλό κλικ για άνοιγμα)', font=('Arial', 11, 'bold')).pack(anchor='w')
        result_frame = ttk.Frame(shell)
        result_frame.pack(fill='both', expand=True, pady=(6, 10))
        columns = ('name', 'kind', 'folder', 'size')
        self.tree = ttk.Treeview(result_frame, columns=columns, show='headings', selectmode='browse')
        for col, text, width, stretch in [('name', 'Όνομα αρχείου', 240, True), ('kind', 'Τύπος', 85, False), ('folder', 'Θέση', 500, True), ('size', 'KB', 70, False)]:
            self.tree.heading(col, text=text, command=lambda c=col: self.sort_column(c, False))
            self.tree.column(col, width=width, minwidth=60, stretch=stretch, anchor='e' if col == 'size' else 'w')
        vertical = ttk.Scrollbar(result_frame, orient='vertical', command=self.tree.yview)
        horizontal = ttk.Scrollbar(result_frame, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.tree.grid(row=0, column=0, sticky='nsew')
        vertical.grid(row=0, column=1, sticky='ns')
        horizontal.grid(row=1, column=0, sticky='ew')
        result_frame.rowconfigure(0, weight=1)
        result_frame.columnconfigure(0, weight=1)
        self.tree.bind('<<TreeviewSelect>>', self.show_selection)
        self.tree.bind('<Double-1>', lambda e: self.open_selected())
        bottom = ttk.LabelFrame(shell, text='Απόσπασμα όπου εντοπίστηκε η λέξη', padding=8)
        bottom.pack(fill='x')
        self.preview = tk.Text(bottom, height=5, wrap='word', font=('Arial', 10), bg='#f7f7f7', relief='flat')
        self.preview.pack(fill='both')
        self.preview.tag_configure('highlight', background='#ffff00', foreground='#000000')
        self.preview.config(state='disabled')
        buttons = ttk.Frame(shell)
        buttons.pack(fill='x', pady=(8, 0))
        ttk.Label(buttons, textvariable=self.status).pack(side='left', fill='x', expand=True)
        ttk.Button(buttons, text='Άνοιγμα φακέλου', command=self.open_folder).pack(side='right', padx=(8, 0))
        ttk.Button(buttons, text='Άνοιγμα αρχείου', command=self.open_selected).pack(side='right')
        terms_entry.focus_set()

    def pick_folder(self):
        folder = filedialog.askdirectory(initialdir=self.folder.get() if Path(self.folder.get()).exists() else None)
        if folder:
            self.folder.set(folder)

    def begin(self):
        if self.running:
            return
        folder = Path(self.folder.get().strip().strip('"'))
        if not folder.is_dir():
            messagebox.showerror('Μη έγκυρος φάκελος', 'Επίλεξε έναν φάκελο που υπάρχει.')
            return
        query = self.keywords.get().strip()
        if not query:
            messagebox.showinfo('Λείπουν λέξεις', 'Γράψε μία ή περισσότερες λέξεις ή μια φράση.')
            return
        # Space-delimited words; double-quoted groups are treated as exact phrases.
        raw = re.findall(r'"([^"]+)"|(\S+)', query)
        words = [normalize(a or b) for a, b in raw if a or b]
        if not words:
            return
        self.search_terms = words
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        self.paths.clear()
        self.set_preview('')
        self.cancel.clear()
        self.running = True
        self.start_btn.configure(state='disabled')
        self.stop_btn.configure(state='normal')
        self.status.set('Γίνεται αναζήτηση…')
        threading.Thread(target=scan, args=(folder, words, self.all_terms.get(), self.recursive.get(), self.cancel, self.q), daemon=True).start()

    def stop(self):
        self.cancel.set()
        self.status.set('Γίνεται διακοπή…')

    def sort_column(self, col, reverse):
        l = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]
        if col == 'size':
            l.sort(key=lambda t: float(t[0].replace(',', '')), reverse=reverse)
        else:
            l.sort(reverse=reverse)
        for index, (val, k) in enumerate(l):
            self.tree.move(k, '', index)
        self.tree.heading(col, command=lambda: self.sort_column(col, not reverse))

    def poll(self):
        try:
            for _ in range(200):
                typ, data = self.q.get_nowait()
                if typ == 'match':
                    iid = self.tree.insert('', 'end', values=(data.path.name, data.path.suffix.upper().lstrip('.'), str(data.path.parent), f'{data.size / 1024:,.1f}'))
                    self.paths[iid] = data.path
                    self.excerpts[iid] = data.snippet
                elif typ == 'progress':
                    checked, found, skipped, _ = data
                    self.status.set(f'Ελέγχθηκαν {checked} αρχεία · Βρέθηκαν {found} · Δεν διαβάστηκαν {skipped}')
                elif typ == 'done':
                    checked, found, skipped, cancelled, issues = data
                    self.finish()
                    state = 'Διακόπηκε.' if cancelled else 'Ολοκληρώθηκε.'
                    self.status.set(f'{state} Ελέγχθηκαν {checked} αρχεία · Βρέθηκαν {found} · Δεν διαβάστηκαν {skipped}')
                    if issues:
                        self.root.after(0, lambda entries=issues: messagebox.showwarning('Ορισμένα αρχεία παραλείφθηκαν', '\n'.join(entries[:12]) + ('\n…' if len(entries) > 12 else '')))
                elif typ == 'fatal':
                    self.finish()
                    messagebox.showerror('Πρόβλημα στην αναζήτηση', data)
        except queue.Empty:
            pass
        self.root.after(100, self.poll)

    def set_preview(self, value):
        self.preview.configure(state='normal')
        self.preview.delete('1.0', 'end')
        self.preview.insert('1.0', value)
        
        # Highlighting logic
        if value and self.search_terms:
            content = normalize(value)
            for term in self.search_terms:
                start_pos = 0
                while True:
                    idx = content.find(term, start_pos)
                    if idx == -1: break
                    
                    # Map normalized index to original text index if possible
                    # Simple approach: same index (works for most Greek/Latin)
                    start_idx = f"1.0 + {idx} chars"
                    end_idx = f"1.0 + {idx + len(term)} chars"
                    self.preview.tag_add('highlight', start_idx, end_idx)
                    start_pos = idx + len(term)
                    
        self.preview.configure(state='disabled')

    @property
    def excerpts_deprecated(self):
        # Removed since we init it in __init__ now
        pass

    def show_selection(self, _event=None):
        selection = self.tree.selection()
        self.set_preview(self.excerpts.get(selection[0], '') if selection else '')

    def selected_path(self):
        selection = self.tree.selection()
        return self.paths.get(selection[0]) if selection else None

    def open_selected(self):
        path = self.selected_path()
        if path:
            try:
                os.startfile(str(path)) if sys.platform == 'win32' else subprocess.Popen(['xdg-open', str(path)])
            except OSError as e:
                messagebox.showerror('Δεν ανοίγει το αρχείο', str(e))

    def open_folder(self):
        path = self.selected_path()
        if path:
            try:
                if sys.platform == 'win32':
                    subprocess.Popen(['explorer', '/select,', str(path)])
                else:
                    subprocess.Popen(['xdg-open', str(path.parent)])
            except OSError as e:
                messagebox.showerror('Δεν ανοίγει ο φάκελος', str(e))


if __name__ == '__main__':
    app_root = tk.Tk()
    SearchApp(app_root)
    app_root.mainloop()
