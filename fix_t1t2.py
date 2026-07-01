import sys
sys.stdout.reconfigure(encoding='utf-8')
path = r'C:\Users\Abfuellung 15\Documents\Marcels Skripts\Bombadil\Bombadil.py'
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

old = (
    '    # ── Tour-Abfahrt Steuerung ────────────────────────────────────────────\n'
    '    def _set_t1_abfahrt(self):\n'
    '        import datetime as _dt\n'
    '        tz = _load_tour_zeiten()\n'
    '        if tz.get("t1"):\n'
    '            return\n'
    '        jetzt = (_dt.datetime.utcnow() + _dt.timedelta(hours=2)).strftime("%H:%M")\n'
    '        t1_barcodes = [r["barcode"] for r in self._all_rows]\n'
    '        _save_tour_zeiten(t1=jetzt, t2=tz.get("t2"), t1_barcodes=t1_barcodes)\n'
    '        self._restore_tour_buttons()\n'
    '        self._recompute_tours_local()\n'
    '        self._upload_tourliste("T1")\n'
    '        self._upload_tour_zeiten_to_drive()   # ← Kollegen-Sync: andere PCs sehen Tour 1 sofort\n'
    '\n'
    '    def _set_t2_abfahrt(self):\n'
    '        import datetime as _dt\n'
    '        tz = _load_tour_zeiten()\n'
    '        if tz.get("t2"):\n'
    '            return\n'
    '        jetzt = (_dt.datetime.utcnow() + _dt.timedelta(hours=2)).strftime("%H:%M")\n'
    '        t1_bc = set(tz.get("t1_barcodes") or [])\n'
    '        t2_barcodes = [r["barcode"] for r in self._all_rows if r["barcode"] not in t1_bc]\n'
    '        _save_tour_zeiten(t1=tz.get("t1"), t2=jetzt, t2_barcodes=t2_barcodes)\n'
    '        self._restore_tour_buttons()\n'
    '        self._recompute_tours_local()\n'
    '        self._upload_tourliste("T2")\n'
    '        self._upload_tour_zeiten_to_drive()   # ← Kollegen-Sync: andere PCs sehen Tour 2 sofort\n'
)

new = (
    '    # ── Tour-Abfahrt Steuerung ────────────────────────────────────────────\n'
    '    def _lese_netz_tour_zeiten(self):\n'
    '        """Liest tour_zeiten JSON direkt vom Netzlaufwerk (synchron). Gibt {} zurueck wenn nicht gefunden."""\n'
    '        import datetime as _dt, json as _json\n'
    '        heute = (_dt.datetime.utcnow() + _dt.timedelta(hours=2)).strftime("%Y-%m-%d")\n'
    '        pfad = TOURLISTEN_DIR / f"tour_zeiten_{heute}.json"\n'
    '        try:\n'
    '            if pfad.exists():\n'
    '                return _json.loads(pfad.read_text(encoding="utf-8"))\n'
    '        except Exception:\n'
    '            pass\n'
    '        return {}\n'
    '\n'
    '    def _set_t1_abfahrt(self):\n'
    '        import datetime as _dt\n'
    '        # Zuerst Netzlaufwerk pruefen – wer zuerst klickt gewinnt\n'
    '        netz = self._lese_netz_tour_zeiten()\n'
    '        if netz.get("t1"):\n'
    '            # anderer PC war schneller – dessen Daten uebernehmen\n'
    '            _save_tour_zeiten(netz.get("t1"), netz.get("t2"),\n'
    '                              netz.get("t1_barcodes", []), netz.get("t2_barcodes", []))\n'
    '            self._restore_tour_buttons()\n'
    '            self._recompute_tours_local()\n'
    '            return\n'
    '        tz = _load_tour_zeiten()\n'
    '        if tz.get("t1"):\n'
    '            return\n'
    '        jetzt = (_dt.datetime.utcnow() + _dt.timedelta(hours=2)).strftime("%H:%M")\n'
    '        t1_barcodes = [r["barcode"] for r in self._all_rows]\n'
    '        _save_tour_zeiten(t1=jetzt, t2=tz.get("t2"), t1_barcodes=t1_barcodes)\n'
    '        self._restore_tour_buttons()\n'
    '        self._recompute_tours_local()\n'
    '        self._upload_tourliste("T1")\n'
    '        self._upload_tour_zeiten_to_drive()\n'
    '\n'
    '    def _set_t2_abfahrt(self):\n'
    '        import datetime as _dt\n'
    '        # Zuerst Netzlaufwerk pruefen – wer zuerst klickt gewinnt\n'
    '        netz = self._lese_netz_tour_zeiten()\n'
    '        if netz.get("t2"):\n'
    '            # anderer PC war schneller – dessen Daten uebernehmen\n'
    '            _save_tour_zeiten(netz.get("t1"), netz.get("t2"),\n'
    '                              netz.get("t1_barcodes", []), netz.get("t2_barcodes", []))\n'
    '            self._restore_tour_buttons()\n'
    '            self._recompute_tours_local()\n'
    '            return\n'
    '        tz = _load_tour_zeiten()\n'
    '        if tz.get("t2"):\n'
    '            return\n'
    '        jetzt = (_dt.datetime.utcnow() + _dt.timedelta(hours=2)).strftime("%H:%M")\n'
    '        # t1_barcodes vom Netzlaufwerk verwenden falls verfuegbar (konsistente Basis)\n'
    '        t1_bc = set(netz.get("t1_barcodes") or tz.get("t1_barcodes") or [])\n'
    '        t2_barcodes = [r["barcode"] for r in self._all_rows if r["barcode"] not in t1_bc]\n'
    '        _save_tour_zeiten(t1=tz.get("t1"), t2=jetzt, t2_barcodes=t2_barcodes)\n'
    '        self._restore_tour_buttons()\n'
    '        self._recompute_tours_local()\n'
    '        self._upload_tourliste("T2")\n'
    '        self._upload_tour_zeiten_to_drive()\n'
)

if old in content:
    content = content.replace(old, new, 1)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    print('OK')
else:
    print('MISS')
    # Debug
    idx = content.find('def _set_t1_abfahrt')
    print(repr(content[idx-60:idx+20]))
