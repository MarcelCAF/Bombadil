import sys
sys.stdout.reconfigure(encoding='utf-8')
path = r'C:\Users\Abfuellung 15\Documents\Marcels Skripts\Bombadil\Bombadil.py'
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

old = '                df.to_excel(pfad, index=False)\n'

new = (
    '                from openpyxl import Workbook as _WB\n'
    '                from openpyxl.styles import Font as _Fnt, PatternFill as _Fill, Alignment as _Aln, Border as _Brd, Side as _Side\n'
    '                from openpyxl.utils import get_column_letter as _gcl\n'
    '                import math as _math\n'
    '                wb = _WB()\n'
    '                ws = wb.active\n'
    '                cols = list(df.columns)\n'
    '                _thin = _Side(style=\'thin\', color=\'BFBFBF\')\n'
    '                _border = _Brd(left=_thin, right=_thin, top=_thin, bottom=_thin)\n'
    '                _col_w = [22, 28, 13, 16, 20, 20]\n'
    '                for ci, cn in enumerate(cols, 1):\n'
    '                    c = ws.cell(row=1, column=ci, value=cn)\n'
    '                    c.fill = _Fill(\'solid\', start_color=\'4F81BD\', end_color=\'4F81BD\')\n'
    '                    c.font = _Fnt(name=\'Arial\', bold=True, color=\'FFFFFF\', size=10)\n'
    '                    c.alignment = _Aln(horizontal=\'center\', vertical=\'center\', wrap_text=True)\n'
    '                    c.border = _border\n'
    '                    ws.column_dimensions[_gcl(ci)].width = _col_w[ci - 1] if ci <= len(_col_w) else 15\n'
    '                ws.row_dimensions[1].height = 30\n'
    '                for ri, (_, row) in enumerate(df.iterrows(), 2):\n'
    '                    for ci, cn in enumerate(cols, 1):\n'
    '                        val = row[cn]\n'
    '                        if isinstance(val, float) and _math.isnan(val):\n'
    '                            val = None\n'
    '                        c = ws.cell(row=ri, column=ci, value=val)\n'
    '                        c.font = _Fnt(name=\'Arial\', size=10)\n'
    '                        c.alignment = _Aln(vertical=\'center\')\n'
    '                        c.border = _border\n'
    '                ws.freeze_panes = \'A2\'\n'
    '                wb.save(pfad)\n'
)

if old in content:
    content = content.replace(old, new, 1)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    print('OK')
else:
    print('MISS')
    idx = content.find('df.to_excel(pfad')
    print(repr(content[idx-80:idx+40]))
