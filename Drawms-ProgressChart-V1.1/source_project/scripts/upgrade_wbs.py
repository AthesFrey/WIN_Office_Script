"""Upgrade the legacy WBS XML with a compact sample and the original input columns.

OOXML is edited directly so VBA, shapes and Excel extensions are not lost.
The saved example is evaluated here so it is useful before the first refresh.
"""
from copy import deepcopy
import datetime as dt
import re

from lxml import etree as ET
from openpyxl.utils.cell import get_column_letter, column_index_from_string

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS = {'s': S}

def q(name):
    return f'{{{S}}}{name}'


def upgrade(wbs, styles, strings):
    data = wbs.find('s:sheetData', NS)
    cached = {c.get('r'): deepcopy(c) for c in data.findall('.//s:c', NS)}
    shared = [''.join(si.itertext()) for si in strings]
    def value(ref):
        c = cached.get(ref)
        if c is None:
            return None
        v = c.find('s:v', NS)
        if v is None:
            return None
        return shared[int(v.text)] if c.get('t') == 's' else float(v.text)

    # The first complete project demonstrates all schedule features. Repeated
    # example projects in rows 20:64 only increase the initial workbook size.
    last = 19
    tasks = {}
    for r in range(6, last + 1):
        name = value(f'C{r}')
        if not name:
            continue
        tasks[r] = {'name': name, 'level': min(5, (len(name) - len(name.lstrip())) // 4 + 1),
                    'start': value(f'G{r}'), 'days': value(f'F{r}'), 'progress': value(f'N{r}') or 0,
                    'predecessor': value(f'O{r}')}
    ordered = list(tasks)
    for i, r in enumerate(ordered):
        t = tasks[r]
        descendants = []
        for n in ordered[i+1:]:
            if tasks[n]['level'] <= t['level']:
                break
            descendants.append(n)
        t['descendants'] = descendants
    name_rows = {''.join(t['name'].split()): r for r, t in tasks.items()}
    active = set()
    def resolve(r):
        t = tasks[r]
        if 'end' in t:
            return
        if r in active:
            raise ValueError(f'Circular dependency in sample task {r}')
        active.add(r)
        if t['descendants']:
            leaves = [n for n in t['descendants'] if not tasks[n]['descendants']]
            for n in leaves:
                resolve(n)
            t['start'] = min(tasks[n]['start'] for n in leaves)
            t['end'] = max(tasks[n]['end'] for n in leaves)
            t['days'] = t['end'] - t['start'] + 1
            total = sum(tasks[n]['days'] for n in leaves)
            t['progress'] = sum(tasks[n]['days'] * tasks[n]['progress'] for n in leaves) / total
        else:
            if t['predecessor']:
                n = name_rows[''.join(t['predecessor'].split())]
                resolve(n)
                t['start'] = tasks[n]['end'] + 1
            t['end'] = t['start'] + max(1, t['days']) - 1
        active.remove(r)
    for r in tasks:
        resolve(r)

    xfs = styles.find('s:cellXfs', NS)
    fmts = styles.find('s:numFmts', NS)
    date_fmt = max(int(f.get('numFmtId')) for f in fmts) + 1
    ET.SubElement(fmts, q('numFmt'), numFmtId=str(date_fmt), formatCode='yyyy/m/d')
    ET.SubElement(fmts, q('numFmt'), numFmtId=str(date_fmt+1), formatCode='m/d')
    derived = {}
    def style(base, *, date=False, axis=False, weekend=False, wrap=False, progress=False):
        key = (base, date, axis, weekend, wrap, progress)
        if key not in derived:
            xf = deepcopy(xfs[int(base)])
            if date or axis:
                xf.set('numFmtId', str(date_fmt + int(axis)))
                xf.set('applyNumberFormat', '1')
            if progress:
                xf.set('numFmtId', '10')  # Built-in 0.00%, keeping full stored precision.
                xf.set('applyNumberFormat', '1')
            if wrap:
                alignment = xf.find('s:alignment', NS)
                if alignment is None: alignment = ET.SubElement(xf, q('alignment'))
                alignment.set('wrapText', '1'); xf.set('applyAlignment', '1')
            if weekend:
                xf.set('fillId', str(weekend_fill)); xf.set('applyFill', '1')
            derived[key] = str(len(xfs)); xfs.append(xf)
        return derived[key]
    fills = styles.find('s:fills', NS)
    weekend_fill = len(fills)
    pf = ET.SubElement(ET.SubElement(fills, q('fill')), q('patternFill'), patternType='solid')
    ET.SubElement(pf, q('fgColor'), rgb='FFF1F5F9')
    ET.SubElement(pf, q('bgColor'), indexed='64')
    rows = {}
    data.clear()
    def row(r):
        if r not in rows:
            rows[r] = ET.SubElement(data, q('row'), r=str(r), ht=str(24 if r >= 6 else 22), customHeight='1')
        return rows[r]
    def cell(r, c, val, st='0'):
        ref = f'{get_column_letter(c)}{r}'
        node = ET.SubElement(row(r), q('c'), r=ref, s=str(st))
        if isinstance(val, str):
            node.set('t', 'inlineStr')
            t = ET.SubElement(ET.SubElement(node, q('is')), q('t'))
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve'); t.text = val
        elif val is not None:
            ET.SubElement(node, q('v')).text = str(val)
        return node
    cell(1, 2, 'WBS Gantt Planner v5.6fix3', '8'); row(1).set('ht', '28')
    cell(2, 2, 'Level', '6'); cell(2, 4, 1, '19')
    cell(2, 5, 'Select tasks, then choose Level or use Promote / Demote.', '19')
    row(3).set('ht', '34'); row(5).set('ht', '42')
    for ref, c in cached.items():
        m = re.fullmatch(r'([A-Z]+)(\d+)', ref)
        col, r = column_index_from_string(m[1]), int(m[2])
        if col > 25 or r < 5 or r > last:
            continue
        # Task date/schedule/number columns are recalculated below.
        if r >= 6 and col in [2, 6, 7, 8, 14, 24, 25]:
            continue
        if col == 15 and r in tasks and tasks[r]['predecessor']:
            predecessor = name_rows[''.join(tasks[r]['predecessor'].split())]
            text = tasks[predecessor]['name']
            c.set('t', 'str')
            for child in list(c): c.remove(child)
            ET.SubElement(c, q('f')).text = f'C{predecessor}'
            ET.SubElement(c, q('v')).text = text
        if r == 5 or col == 3:
            c.set('s', style(c.get('s','0'), wrap=True))
            if r >= 6: row(r).set('ht', str(max(24, (len(tasks[r]['name']) // 36 + 1) * 15)))
        row(r).append(c)
    for r, t in tasks.items():
        parent = bool(t['descendants'])
        base = '9' if parent else '20'
        cell(r, 2, r-5, '29' if parent else '30')
        cell(r, 6, t['days'], '82' if parent else '83')
        cell(r, 7, t['start'], style(base, date=True))
        cell(r, 8, t['end'], style(base, date=True))
        cell(r, 14, t['progress'], style('11' if parent else '22', progress=True))
        cell(r, 24, 'P' if parent else 'T', '30')
        cell(r, 25, t['level'], '30')
    first_date = int(min(t['start'] for t in tasks.values()))
    last_date = int(max(t['end'] for t in tasks.values()))
    end_col = 26 + last_date-first_date
    for serial in range(first_date, last_date+1):
        c = 26 + serial-first_date
        day = dt.date(1899, 12, 30) + dt.timedelta(days=serial)
        weekend = day.weekday() >= 5
        cell(4, c, serial, style('6', axis=True, weekend=weekend))
        cell(5, c, ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'][day.weekday()], style('6', weekend=weekend))
        for r in tasks:
            cell(r, c, None, style('30', weekend=weekend))
    # Initial Gantt bars use their real date span; Excel refresh redraws them.
    merges = wbs.find('s:mergeCells', NS)
    merges.clear()
    for ref in ['B1:O1', 'B2:C2', 'E2:O2']:
        ET.SubElement(merges, q('mergeCell'), ref=ref)
    for r, t in tasks.items():
        start_col = 26 + int(t['start'])-first_date
        finish_col = 26 + int(t['end'])-first_date
        parent = bool(t['descendants'])
        for c in rows[r]:
            idx = column_index_from_string(c.get('r').rstrip('0123456789'))
            if idx == start_col:
                ET.SubElement(c, q('v')).text = str(t['progress'])
        # A colored fill is enough to show the full saved bar without Excel.
        fill_id = len(fills)
        pf = ET.SubElement(ET.SubElement(fills, q('fill')), q('patternFill'), patternType='solid')
        ET.SubElement(pf, q('fgColor'), rgb='FF8B0000' if parent else ('FF62C384' if t['progress'] >= 1 else 'FFC6DCF6'))
        ET.SubElement(pf, q('bgColor'), indexed='64')
        xf = deepcopy(xfs[int(style('11' if parent else '22', progress=True))])
        xf.set('fillId', str(fill_id)); xf.set('applyFill', '1')
        alignment = xf.find('s:alignment', NS)
        if alignment is None:
            alignment = ET.SubElement(xf, q('alignment'))
        alignment.set('shrinkToFit', '1'); alignment.set('wrapText', '0')
        xf.set('applyAlignment', '1')
        idx = len(xfs); xfs.append(xf)
        for c in rows[r]:
            col = column_index_from_string(c.get('r').rstrip('0123456789'))
            if start_col <= col <= finish_col:
                c.set('s', str(idx))
        if finish_col > start_col:
            ET.SubElement(merges, q('mergeCell'), ref=f'{get_column_letter(start_col)}{r}:{get_column_letter(finish_col)}{r}')
    merges.set('count', str(len(merges)))
    # Derive styles instead of changing fonts shared by titles/headers or the
    # legacy template. Remove theme-font selection so Office uses YaHei exactly.
    fonts = styles.find('s:fonts', NS)
    body_fonts, body_styles = {}, {}
    for r in range(6, last + 1):
        existing = {column_index_from_string(c.get('r').rstrip('0123456789')): c for c in rows[r]}
        for col in range(2, end_col + 1):
            c = existing.get(col)
            if c is None:
                c = cell(r, col, None)
            base = int(c.get('s', '0'))
            if base not in body_styles:
                xf = deepcopy(xfs[base])
                original_font = int(xf.get('fontId', '0'))
                if original_font not in body_fonts:
                    f = deepcopy(fonts[original_font])
                    for tag in ['scheme', 'name', 'sz', 'charset']:
                        for child in list(f.findall('s:' + tag, NS)):
                            f.remove(child)
                    ET.SubElement(f, q('sz'), val='11')
                    ET.SubElement(f, q('name'), val='Microsoft YaHei')
                    ET.SubElement(f, q('charset'), val='134')
                    body_fonts[original_font] = len(fonts)
                    fonts.append(f)
                xf.set('fontId', str(body_fonts[original_font]))
                xf.set('applyFont', '1')
                body_styles[base] = len(xfs)
                xfs.append(xf)
            c.set('s', str(body_styles[base]))
    fonts.set('count', str(len(fonts)))
    # Remove obsolete conditional formats/control references and the old axis.
    for node in list(wbs):
        if ET.QName(node).localname in ['conditionalFormatting','legacyDrawing','AlternateContent','extLst','autoFilter']:
            wbs.remove(node)
    cols = wbs.find('s:cols', NS); cols.clear()
    widths = {1:3,2:7,3:40,4:14,5:13,6:8,7:15,8:15,14:12,15:40}
    for c in range(1,26):
        ET.SubElement(cols,q('col'),min=str(c),max=str(c),width=str(widths.get(c,14)),customWidth='1',hidden=str(int(c not in widths)))
    ET.SubElement(cols,q('col'),min='26',max=str(end_col),width='6',customWidth='1')
    wbs.find('s:dimension', NS).set('ref', f'A1:{get_column_letter(end_col)}{last}')
    view = wbs.find('s:sheetViews/s:sheetView',NS)
    for node in list(view): view.remove(node)
    view.set('showGridLines','0'); view.set('tabSelected','1')
    ET.SubElement(view,q('pane'),xSplit='3',ySplit='5',topLeftCell='D6',activePane='bottomRight',state='frozen')
    ET.SubElement(view,q('selection'),pane='bottomRight',activeCell='D6',sqref='D6')
    validations = wbs.find('s:dataValidations', NS)
    for dv in list(validations):
        if dv.get('sqref') != 'D2' and not dv.get('sqref','').startswith('O6:'):
            validations.remove(dv)
        elif dv.get('sqref','').startswith('O6:'):
            dv.set('sqref', f'O6:O{max(1000,last+200)}')
            dv.find('s:formula1',NS).text = f'$C$6:$C${last}'
    validations.set('count',str(len(validations)))
    for r in sorted(rows):
        node = rows[r]
        node[:] = sorted(node, key=lambda c: column_index_from_string(c.get('r').rstrip('0123456789')))
    data[:] = [rows[r] for r in sorted(rows)]
    for node in [xfs, fmts, fills]: node.set('count',str(len(node)))
    return last, end_col
