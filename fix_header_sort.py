import sys
sys.stdout.reconfigure(encoding='utf-8')
path = r'C:\Users\Abfuellung 15\Documents\Marcels Skripts\Bombadil\Bombadil.py'
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

errors = []

# 1. Sort-Dropdown entfernen, Sort-State-Variablen einfuegen
old1 = (
    '        tk.Label(filter_row, text="Sortierung:", font=("Segoe UI", 9)\n'
    '                 ).pack(side="left", padx=(14, 0))\n'
    '        self._sort_var = tk.StringVar(value="Standard")\n'
    '        _sort_opts = [\n'
    '            "Standard",\n'
    '            "Name A→Z",\n'
    '            "Ziel-Kiosk A→Z",\n'
    '            "Älteste zuerst",\n'
    '        ]\n'
    '        _sort_cb = ttk.Combobox(filter_row, textvariable=self._sort_var,\n'
    '                                values=_sort_opts, state="readonly",\n'
    '                                font=("Segoe UI", 9), width=20)\n'
    '        _sort_cb.pack(side="left", padx=4)\n'
    '        _sort_cb.bind("<<ComboboxSelected>>", lambda e: self._refresh_ui())\n'
)
new1 = (
    '        self._sort_col = None   # None = Standard, int = Spaltenindex\n'
    '        self._sort_dir = 0      # 0 = Standard, 1 = aufsteigend, 2 = absteigend\n'
)
if old1 in content:
    content = content.replace(old1, new1, 1); print('Sort-Dropdown: OK')
else:
    errors.append('Sort-Dropdown: MISS')
    idx = content.find('Sortierung:')
    print(repr(content[idx-10:idx+300]))

# 2. Header-Klick-Binding hinzufuegen
old2 = (
    '            try:\n'
    '                self._sheet.extra_bindings([("cell_select", self._on_cell_click)])\n'
    '            except Exception:\n'
    '                pass\n'
)
new2 = (
    '            try:\n'
    '                self._sheet.extra_bindings([\n'
    '                    ("cell_select", self._on_cell_click),\n'
    '                    ("column_header_select", self._on_header_click),\n'
    '                ])\n'
    '            except Exception:\n'
    '                pass\n'
)
if old2 in content:
    content = content.replace(old2, new2, 1); print('Header-Binding: OK')
else:
    errors.append('Header-Binding: MISS')

# 3. Sort-Logik in _refresh_ui ersetzen
old3 = (
    '        # ── Sortierung ─────────────────────────────────────────────────────────────────\n'
    '        s = self._sort_var.get()\n'
    '        _none_last = "ÿ"   # leere Strings nach hinten sortieren\n'
    '        if s == "Name A→Z":\n'
    '            rows = sorted(rows, key=lambda r: str(r["name"] or "").lower() or _none_last)\n'
    '        elif s == "Ziel-Kiosk A→Z":\n'
    '            rows = sorted(rows, key=lambda r: str(r["zielkiosk"] or "").lower() or _none_last)\n'
    '        elif s == "Älteste zuerst":\n'
    '            rows = sorted(rows, key=lambda r: r["verpackt_at"] or _none_last)\n'
)
new3 = (
    '        # ── Sortierung ─────────────────────────────────────────────────────────────────\n'
    '        _none_last = "ÿ"   # leere Strings nach hinten sortieren\n'
    '        _col_keys = {\n'
    '            0: lambda r: str(r.get("tour", "") or "").lower(),\n'
    '            1: lambda r: str(r["barcode"] or "").lower(),\n'
    '            2: lambda r: str(r["name"] or "").lower() or _none_last,\n'
    '            3: lambda r: str(r["tb_status"] or "").lower(),\n'
    '            5: lambda r: str(r["verpackt_at"] or _none_last),\n'
    '            6: lambda r: str(r["abholbereit_at"] or _none_last),\n'
    '            7: lambda r: str(r["zielkiosk"] or "").lower() or _none_last,\n'
    '        }\n'
    '        if self._sort_col is not None and self._sort_dir > 0 and self._sort_col in _col_keys:\n'
    '            rows = sorted(rows, key=_col_keys[self._sort_col], reverse=(self._sort_dir == 2))\n'
)
if old3 in content:
    content = content.replace(old3, new3, 1); print('Sort-Logik: OK')
else:
    errors.append('Sort-Logik: MISS')
    idx = content.find('_sort_var.get()')
    print(repr(content[idx-80:idx+200]))

# 4. Standard-Sort anpassen
old4 = (
    '        if s == "Standard":\n'
    '            rows = sorted(rows, key=lambda r: (_stage(r), _tour_order(r)))\n'
    '        else:\n'
    '            rows = sorted(rows, key=lambda r: (0 if not r["_in_db_bool"] else 1))\n'
)
new4 = (
    '        if self._sort_col is None or self._sort_dir == 0:\n'
    '            rows = sorted(rows, key=lambda r: (_stage(r), _tour_order(r)))\n'
)
if old4 in content:
    content = content.replace(old4, new4, 1); print('Standard-Sort: OK')
else:
    errors.append('Standard-Sort: MISS')
    idx = content.find('s == "Standard"')
    print(repr(content[idx-10:idx+200]))

# 5. _on_header_click Methode nach _on_cell_click einfuegen
old5 = '    def _on_cell_click(self, event):\n'
new5 = (
    '    def _on_header_click(self, event):\n'
    '        """Spaltenheader klicken: 1x aufsteigend, 2x absteigend, 3x Standard."""\n'
    '        try:\n'
    '            col = event.column if hasattr(event, "column") else event[1]\n'
    '        except Exception:\n'
    '            return\n'
    '        _sortable = {0, 1, 2, 3, 5, 6, 7}\n'
    '        if col not in _sortable:\n'
    '            return\n'
    '        if self._sort_col == col:\n'
    '            self._sort_dir = (self._sort_dir + 1) % 3\n'
    '            if self._sort_dir == 0:\n'
    '                self._sort_col = None\n'
    '        else:\n'
    '            self._sort_col = col\n'
    '            self._sort_dir = 1\n'
    '        self._refresh_ui()\n'
    '\n'
    '    def _on_cell_click(self, event):\n'
)
if old5 in content:
    content = content.replace(old5, new5, 1); print('Header-Click-Methode: OK')
else:
    errors.append('Header-Click-Methode: MISS')

if not errors:
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    print('Datei gespeichert.')
else:
    print('FEHLER:', errors)
