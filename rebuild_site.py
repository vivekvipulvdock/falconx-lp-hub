#!/usr/bin/env python3
"""
Regenerates FalconX_LP_Hub.html from FalconX_LP_Hub_Database.xlsx.

Usage:
    python3 rebuild_site.py [source_html] [database_xlsx] [output_html]

Defaults:
    source_html   = FalconX_LP_Hub.html        (the original/previous bundle)
    database_xlsx = FalconX_LP_Hub_Database.xlsx
    output_html   = FalconX_LP_Hub.html         (overwritten in place)

Only the data (company table + fund overview) is replaced. The page's
design, layout, CSS, fonts, and logic are read verbatim from source_html
and are left untouched.
"""
import sys
import re
import json
from openpyxl import load_workbook

FIELDS = ['id', 'name', 'mono', 'hue', 'sector', 'stage', 'status', 'inv',
          'val', 'own', 'date', 'round', 'post', 'upd', 'desc', 'thesis']
# Columns present in the spreadsheet, in order (everything except 'hue',
# which is a color seed and intentionally not exposed for editing).
SHEET_FIELDS = [f for f in FIELDS if f != 'hue']

FUND_KEYS = ['name', 'size', 'deployed', 'companies', 'value', 'invested',
             'exits', 'tvpi', 'dpi', 'irr', 'positioning']


def load_database(xlsx_path, original_core):
    """Read Companies + Fund Overview sheets back into core list / fund dict."""
    wb = load_workbook(xlsx_path, data_only=True)

    # Preserve original hue-by-id so colors don't shift just because a row
    # was edited. New companies (id not seen before) get an evenly spaced hue.
    hue_by_id = {c[0]: c[3] for c in original_core}
    used_hues = list(hue_by_id.values())

    ws = wb['Companies']
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    core = []
    next_hue_pool = [h for h in range(0, 360, 24)]
    pool_i = 0
    for row in rows:
        if row is None or all(v is None for v in row):
            continue
        rec = dict(zip(SHEET_FIELDS, row))
        cid = str(rec['id']).strip()
        if cid in hue_by_id:
            hue = hue_by_id[cid]
        else:
            hue = next_hue_pool[pool_i % len(next_hue_pool)]
            pool_i += 1
        ordered = [
            cid, rec['name'], rec['mono'], hue, rec['sector'], rec['stage'],
            rec['status'], float(rec['inv']), float(rec['val']),
            float(rec['own']), rec['date'], rec['round'],
            int(rec['post']), int(rec['upd']), rec['desc'], rec['thesis'],
        ]
        core.append(ordered)

    ws2 = wb['Fund Overview']
    fund_rows = list(ws2.iter_rows(min_row=2, max_col=2, values_only=True))
    values = [r[1] for r in fund_rows if r and r[0] is not None]
    fund = dict(zip(FUND_KEYS, values))
    fund['size'] = int(fund['size'])
    fund['deployed'] = float(fund['deployed'])
    fund['companies'] = int(fund['companies'])
    fund['value'] = float(fund['value'])
    fund['invested'] = float(fund['invested'])
    fund['exits'] = int(fund['exits'])
    fund['tvpi'] = float(fund['tvpi'])
    fund['dpi'] = float(fund['dpi'])
    return core, fund


def js_str(s):
    """Single-quoted JS string literal matching the source file's style."""
    s = str(s)
    s = s.replace('\\', '\\\\').replace("'", "\\'")
    return "'" + s + "'"


def js_num(n):
    if isinstance(n, float) and n == int(n):
        return repr(n)  # keep e.g. 2.0 -> '2.0' to match source style
    return repr(n)


def render_core_js(core):
    lines = []
    for c in core:
        cid, name, mono, hue, sector, stage, status, inv, val, own, date, \
            rnd, post, upd, desc, thesis = c
        parts = [
            js_str(cid), js_str(name), js_str(mono), js_num(hue),
            js_str(sector), js_str(stage), js_str(status), js_num(inv),
            js_num(val), js_num(own), js_str(date), js_str(rnd),
            js_num(post), js_num(upd), js_str(desc), js_str(thesis),
        ]
        lines.append('      [' + ','.join(parts) + ']')
    return 'const core=[\n' + ',\n'.join(lines) + '\n    ];\n'


def render_fund_js(fund):
    parts = []
    for k in FUND_KEYS:
        v = fund[k]
        v_js = js_str(v) if isinstance(v, str) else js_num(v)
        parts.append(f'{k}:{v_js}')
    line1 = ', '.join(parts[:4])
    line2 = ', '.join(parts[4:10])
    line3 = ', '.join(parts[10:])
    body = line1 + ',\n      ' + line2 + ',\n      ' + line3
    return 'const fund={\n      ' + body + '\n    };\n'


def main():
    source_html = sys.argv[1] if len(sys.argv) > 1 else 'FalconX_LP_Hub.html'
    db_xlsx = sys.argv[2] if len(sys.argv) > 2 else 'FalconX_LP_Hub_Database.xlsx'
    output_html = sys.argv[3] if len(sys.argv) > 3 else source_html

    with open(source_html, 'r', encoding='utf-8') as f:
        bundle_lines = f.readlines()

    template_line_idx = None
    for i, line in enumerate(bundle_lines):
        if line.startswith('"<!DOCTYPE html>'):
            template_line_idx = i
            break
    if template_line_idx is None:
        raise RuntimeError('Could not find the template data line in ' + source_html)

    template_json = bundle_lines[template_line_idx]
    template_html = json.loads(template_json)

    m_core = re.search(r"const core=\[.*?\];\n", template_html, re.S)
    m_fund = re.search(r"const fund=\{.*?\};\n", template_html, re.S)
    if not m_core or not m_fund:
        raise RuntimeError('Could not locate core/fund data blocks in template')

    original_core_text = m_core.group(0)
    import ast
    original_core = ast.literal_eval(
        re.search(r"\[.*\]", original_core_text, re.S).group(0).rstrip(';\n')
    )

    new_core, new_fund = load_database(db_xlsx, original_core)

    new_template = (
        template_html[:m_core.start()] +
        render_core_js(new_core) +
        template_html[m_core.end():m_fund.start()] +
        render_fund_js(new_fund) +
        template_html[m_fund.end():]
    )

    encoded = json.dumps(new_template, ensure_ascii=False)
    encoded = encoded.replace('</', '<\\u002F')
    bundle_lines[template_line_idx] = encoded + '\n'

    with open(output_html, 'w', encoding='utf-8') as f:
        f.writelines(bundle_lines)

    print(f'Wrote {output_html} ({len(new_core)} companies)')


if __name__ == '__main__':
    main()
