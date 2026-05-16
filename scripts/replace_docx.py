#!/usr/bin/env python3
"""Replace [公式 Fxxx] placeholders in a docx with Word native OMML.

Pipeline:
  1. Dry-run pandoc on every formula in formula_placeholders.json. FAIL-FAST on any error.
  2. Extract OMML for each formula. For display formulas, KEEP ONLY <m:oMathPara>
     (drop the descendant <m:oMath> that pandoc emits alongside) to prevent double-render.
  3. Walk docx body + every <w:tc>. Classify each placeholder occurrence:
       - numbered display  (paragraph matches ^[公式 Fxxx]\\s*(N-N)\\s*$)
         → replace entire paragraph with 1×3 borderless table, columns 200/8200/1100 twips
       - inside table cell (placeholder in a <w:tc>): lock table layout, insert OMML inline
       - inline (everywhere else): run surgery to swap placeholder for OMML
  4. Write output and report.

Usage:
  python3 replace_docx.py --docx INPUT.docx --json formula_placeholders.json --out OUTPUT.docx
"""
import argparse, io, json, os, re, shutil, subprocess, sys, zipfile
from copy import deepcopy
from lxml import etree

NSW = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NSM = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
W = '{%s}' % NSW
M = '{%s}' % NSM
NS = {'w': NSW, 'm': NSM}

# Equation table layout (twips)
TEXT_WIDTH = 9500
COL_LEFT, COL_MID, COL_RIGHT = 200, 8200, 1100

PH_RE = re.compile(r'\[公式\s*F(\d+)\]')
NUM_RE = re.compile(r'^\s*\[公式\s*F(\d+)\]\s*\(\s*([0-9]+\s*[-–]\s*[0-9]+[a-z]?)\s*\)\s*$')

def qn(tag):
    pfx, name = tag.split(':')
    return '{%s}%s' % (NS[pfx], name)

# ============================================================
# Phase A: pandoc dry-run + OMML cache build (fail-fast)
# ============================================================
def build_omml_cache(formulas, cache_dir, fail_fast=True):
    """Run pandoc on all formulas in one shot. Return dict fid -> list of OMML elements.
    For display formulas, the list contains only <m:oMathPara> (no redundant <m:oMath>)."""
    os.makedirs(cache_dir, exist_ok=True)
    md_lines = []
    for item in formulas:
        fid = item['id'].replace('公式 ', '').strip()
        src = item['source']
        s = re.sub(r'\\label\{[^}]*\}', '', src).strip()
        typ = item['type']
        if typ == 'display':
            # Strip outer environment if present, treat as $$...$$
            m = re.match(
                r'\\begin\{(equation|equation\*|align|align\*|gather|gather\*|displaymath)\}(.*)\\end\{\1\}',
                s, re.DOTALL)
            inner = m.group(2).strip() if m else s.strip('\\[\\]').strip('$').strip()
            md_lines.append(f'MARKER_{fid}\n\n$${inner}$$\n')
        else:
            inner = s.strip()
            if inner.startswith('$') and inner.endswith('$'): inner = inner[1:-1]
            if inner.startswith('\\(') and inner.endswith('\\)'): inner = inner[2:-2]
            md_lines.append(f'MARKER_{fid} ${inner}$ END_{fid}\n')
    md_path = os.path.join(cache_dir, '_all_formulas.md')
    open(md_path, 'w', encoding='utf-8').write('\n'.join(md_lines))
    docx_path = os.path.join(cache_dir, '_all_formulas.docx')
    r = subprocess.run(['pandoc', md_path, '-f', 'markdown', '-t', 'docx', '-o', docx_path],
                       capture_output=True, text=True)
    if r.returncode != 0:
        print('pandoc fatal error:', r.stderr, file=sys.stderr); sys.exit(1)

    with zipfile.ZipFile(docx_path) as z:
        doc_xml = z.read('word/document.xml')
    root = etree.fromstring(doc_xml)
    body = root.find(W+'body')
    paragraphs = list(body.findall(W+'p'))

    def para_text(p):
        return ''.join(t.text or '' for t in p.iter(W+'t'))

    cache = {}
    failed = []
    i = 0
    while i < len(paragraphs):
        p = paragraphs[i]
        m = re.match(r'MARKER_(F\d+)', para_text(p))
        if not m:
            i += 1; continue
        fid = m.group(1)
        # Inline math sits in same paragraph; display sits in next paragraph
        elems = []
        for cp in (p, paragraphs[i+1] if i+1 < len(paragraphs) else None):
            if cp is None: continue
            # Look for top-level oMathPara children of paragraph (display)
            paras = cp.findall(M+'oMathPara')
            if paras:
                elems = [deepcopy(e) for e in paras]
                break
            maths = cp.findall(M+'oMath')
            if maths:
                elems = [deepcopy(e) for e in maths]
                break
        if not elems:
            failed.append(fid)
        else:
            cache[fid] = elems
            # Save to disk for inspection
            with open(os.path.join(cache_dir, fid + '.xml'), 'w', encoding='utf-8') as f:
                f.write('\n'.join(etree.tostring(e, encoding='unicode') for e in elems))
        i += 1

    if failed and fail_fast:
        src_map = {x['id'].replace('公式 ', ''): x for x in formulas}
        print('\nPANDOC FAILED ON THESE FORMULAS — stopping (fail-fast).', file=sys.stderr)
        for fid in failed:
            info = src_map.get(fid, {})
            print(f'  {fid}  {info.get("file","?")}:{info.get("line","?")}', file=sys.stderr)
            print(f'    {info.get("source","")[:150]}', file=sys.stderr)
        sys.exit(2)

    print(f'OMML cache: {len(cache)} formulas extracted')
    return cache, failed

# ============================================================
# Phase B: docx walk + replacement
# ============================================================
def get_runs(p):
    """Return list of (run_elem, t_elem_or_None, text_len, text_str) for direct <w:r> children."""
    out = []
    for r in p.findall(W+'r'):
        t = r.find(W+'t')
        s = (t.text if (t is not None and t.text) else '')
        out.append((r, t, len(s), s))
    return out

def para_text(p):
    return ''.join(s for _,_,_,s in get_runs(p))

def split_run(p, run_idx, char_offset):
    """Split runs[run_idx] at char_offset (within run text). Returns updated runs list."""
    runs = get_runs(p)
    r, t, n, s = runs[run_idx]
    if char_offset <= 0 or char_offset >= n: return runs
    r_right = deepcopy(r)
    t_right = r_right.find(W+'t')
    t.text = s[:char_offset]
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t_right.text = s[char_offset:]
    t_right.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    r.addnext(r_right)
    return get_runs(p)

def char_to_run(runs, pos):
    acc = 0
    for i, (r,t,n,s) in enumerate(runs):
        if pos <= acc + n: return i, pos - acc
        acc += n
    return len(runs)-1, runs[-1][2] if runs else 0

def replace_span_with_omml(p, start, end, omml_elems):
    """Replace concatenated-text [start,end) in paragraph p with OMML siblings."""
    runs = get_runs(p)
    i_end, off_end = char_to_run(runs, end)
    runs = split_run(p, i_end, off_end)
    i_start, off_start = char_to_run(runs, start)
    runs = split_run(p, i_start, off_start)
    runs = get_runs(p)
    acc = 0
    first_idx = last_idx = None
    for i, (r,t,n,s) in enumerate(runs):
        if acc == start and first_idx is None: first_idx = i
        if acc + n == end: last_idx = i
        acc += n
    if first_idx is None or last_idx is None: return
    insert_before = runs[first_idx][0]
    for e in [deepcopy(x) for x in omml_elems]:
        insert_before.addprevious(e)
    for j in range(first_idx, last_idx+1):
        runs[j][0].getparent().remove(runs[j][0])

def make_equation_table(omml_elems, number_text):
    """Build a 1x3 borderless equation table (200/8200/1100 twips)."""
    tbl = etree.Element(W+'tbl')
    tblPr = etree.SubElement(tbl, W+'tblPr')
    tblW = etree.SubElement(tblPr, W+'tblW'); tblW.set(W+'w', str(TEXT_WIDTH)); tblW.set(W+'type', 'dxa')
    tblLayout = etree.SubElement(tblPr, W+'tblLayout'); tblLayout.set(W+'type', 'fixed')
    tblBorders = etree.SubElement(tblPr, W+'tblBorders')
    for b in ('top','left','bottom','right','insideH','insideV'):
        el = etree.SubElement(tblBorders, W+b); el.set(W+'val', 'nil')
    tblGrid = etree.SubElement(tbl, W+'tblGrid')
    for w in (COL_LEFT, COL_MID, COL_RIGHT):
        gc = etree.SubElement(tblGrid, W+'gridCol'); gc.set(W+'w', str(w))
    tr = etree.SubElement(tbl, W+'tr')

    def cell(width, align, kind, content, tight=False, nowrap=False):
        tc = etree.SubElement(tr, W+'tc')
        tcPr = etree.SubElement(tc, W+'tcPr')
        tcW = etree.SubElement(tcPr, W+'tcW'); tcW.set(W+'w', str(width)); tcW.set(W+'type', 'dxa')
        if tight:
            tcMar = etree.SubElement(tcPr, W+'tcMar')
            for side in ('top','left','bottom','right'):
                m = etree.SubElement(tcMar, W+side)
                m.set(W+'w', '50' if side in ('left','right') and nowrap else '0')
                m.set(W+'type','dxa')
        if nowrap:
            etree.SubElement(tcPr, W+'noWrap')
        vAlign = etree.SubElement(tcPr, W+'vAlign'); vAlign.set(W+'val','center')
        p = etree.SubElement(tc, W+'p')
        pPr = etree.SubElement(p, W+'pPr')
        jc = etree.SubElement(pPr, W+'jc'); jc.set(W+'val', align)
        spacing = etree.SubElement(pPr, W+'spacing')
        spacing.set(W+'before','0'); spacing.set(W+'after','0')
        spacing.set(W+'line','240'); spacing.set(W+'lineRule','auto')
        if kind == 'omml':
            for e in content: p.append(deepcopy(e))
        elif kind == 'text' and content:
            r = etree.SubElement(p, W+'r')
            t = etree.SubElement(r, W+'t'); t.text = content
            t.set('{http://www.w3.org/XML/1998/namespace}space','preserve')

    cell(COL_LEFT,  'left',   'text', '', tight=True)
    cell(COL_MID,   'center', 'omml', omml_elems, tight=True)
    cell(COL_RIGHT, 'right',  'text', number_text, tight=True, nowrap=True)
    return tbl

def lock_table_layout(tbl):
    """Force fixed layout + explicit tcW on every cell using gridCol widths."""
    tblPr = tbl.find(W+'tblPr')
    if tblPr is None:
        tblPr = etree.SubElement(tbl, W+'tblPr'); tbl.insert(0, tblPr)
    tblLayout = tblPr.find(W+'tblLayout')
    if tblLayout is None:
        tblLayout = etree.SubElement(tblPr, W+'tblLayout')
    tblLayout.set(W+'type', 'fixed')
    grid = tbl.find(W+'tblGrid')
    widths = []
    if grid is not None:
        for gc in grid.findall(W+'gridCol'):
            try: widths.append(int(gc.get(W+'w','0')))
            except: widths.append(0)
    for tr in tbl.findall(W+'tr'):
        for idx, tc in enumerate(tr.findall(W+'tc')):
            tcPr = tc.find(W+'tcPr')
            if tcPr is None:
                tcPr = etree.SubElement(tc, W+'tcPr'); tc.insert(0, tcPr)
            tcW = tcPr.find(W+'tcW')
            cur = tcW.get(W+'w') if tcW is not None else None
            if not cur or cur == '0' or (tcW is not None and tcW.get(W+'type')=='auto'):
                w = widths[idx] if idx < len(widths) else 2000
                if tcW is None:
                    tcW = etree.SubElement(tcPr, W+'tcW')
                tcW.set(W+'w', str(w)); tcW.set(W+'type', 'dxa')

def is_numbered_paragraph(p):
    return NUM_RE.match(para_text(p))

def in_table_ancestor(elem):
    a = elem.getparent()
    while a is not None:
        if a.tag == W+'tc': return True
        a = a.getparent()
    return False

def replace_inline_placeholders(p, cache, stats):
    """Inline replacement of all [公式 Fxxx] in this paragraph."""
    while True:
        text = para_text(p)
        m = PH_RE.search(text)
        if not m: break
        fid = 'F' + m.group(1).zfill(3)
        if fid not in cache:
            stats['missing'].append(fid)
            # Replace with a missing marker to break the loop
            _replace_span_with_text(p, m.start(), m.end(), f'⟦缺失 {fid}⟧')
            continue
        replace_span_with_omml(p, m.start(), m.end(), cache[fid])
        stats['inline'] += 1

def _replace_span_with_text(p, start, end, text):
    runs = get_runs(p)
    i_end, off_end = char_to_run(runs, end)
    runs = split_run(p, i_end, off_end)
    i_start, off_start = char_to_run(runs, start)
    runs = split_run(p, i_start, off_start)
    runs = get_runs(p)
    acc = 0; first = last = None
    for i, (r,t,n,s) in enumerate(runs):
        if acc == start and first is None: first = i
        if acc + n == end: last = i
        acc += n
    if first is None or last is None: return
    r0, t0, _, _ = runs[first]
    if t0 is None:
        t0 = etree.SubElement(r0, W+'t')
    t0.text = text
    t0.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
    for j in range(first+1, last+1):
        runs[j][0].getparent().remove(runs[j][0])

def process(docx_path, formulas, out_path, cache_dir):
    cache, _ = build_omml_cache(formulas, cache_dir, fail_fast=True)

    zin = zipfile.ZipFile(docx_path)
    root = etree.fromstring(zin.read('word/document.xml'))
    body = root.find(W+'body')

    stats = {'inline': 0, 'numbered': 0, 'in_table': 0, 'missing': []}

    # Lock all existing tables so cell widths don't rebalance when OMML is inserted
    for tbl in body.iter(W+'tbl'):
        lock_table_layout(tbl)

    # Phase 1: numbered display paragraphs at body level → replace with 1x3 table
    for elem in list(body):
        if elem.tag != W+'p': continue
        nm = is_numbered_paragraph(elem)
        if not nm: continue
        fid = 'F' + nm.group(1).zfill(3)
        if fid not in cache:
            stats['missing'].append(fid); continue
        num_text = '(' + nm.group(2) + ')'
        new_tbl = make_equation_table(cache[fid], num_text)
        elem.addprevious(new_tbl)
        body.remove(elem)
        stats['numbered'] += 1

    # Phase 2: inline pass everywhere else
    for p in list(root.iter(W+'p')):
        if not PH_RE.search(para_text(p)): continue
        before = stats['inline']
        in_tbl = in_table_ancestor(p)
        replace_inline_placeholders(p, cache, stats)
        if in_tbl:
            delta = stats['inline'] - before
            stats['in_table'] += delta
            stats['inline'] -= delta

    new_xml = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.namelist():
            if item == 'word/document.xml':
                zout.writestr(item, new_xml)
            else:
                zout.writestr(item, zin.read(item))
    open(out_path, 'wb').write(buf.getvalue())
    return stats

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--docx', required=True)
    ap.add_argument('--json', required=True)
    ap.add_argument('--out',  required=True)
    ap.add_argument('--cache-dir', default='/tmp/thesis_formula_cache')
    args = ap.parse_args()

    formulas = json.load(open(args.json, encoding='utf-8'))
    backup = os.path.splitext(args.out)[0] + '.before-replace.docx'
    if not os.path.exists(backup):
        shutil.copyfile(args.docx, backup)
        print(f'Backup: {backup}')

    stats = process(args.docx, formulas, args.out, args.cache_dir)
    print('\n=== Replacement stats ===')
    print(f'  inline:   {stats["inline"]}')
    print(f'  numbered: {stats["numbered"]}')
    print(f'  in_table: {stats["in_table"]}')
    print(f'  missing:  {len(set(stats["missing"]))}')
    if stats['missing']:
        print(f'  missing IDs: {sorted(set(stats["missing"]))}')

    report = os.path.join(os.path.dirname(args.out), 'thesis_replace_report.md')
    with open(report, 'w', encoding='utf-8') as f:
        f.write('# Replacement Report\n\n')
        f.write(f'- Source: `{args.docx}`\n- Output: `{args.out}`\n\n')
        f.write(f'## Counts\n- inline: **{stats["inline"]}**\n- numbered: **{stats["numbered"]}**\n')
        f.write(f'- in_table: **{stats["in_table"]}**\n- total: **{sum([stats["inline"],stats["numbered"],stats["in_table"]])}**\n')
        if stats['missing']:
            f.write(f'\n## Missing\n')
            for fid in sorted(set(stats['missing'])):
                f.write(f'- {fid}\n')
    print(f'Report: {report}')

if __name__ == '__main__':
    main()
