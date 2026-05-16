#!/usr/bin/env python3
"""Plan F + heading detection: clean up PDF→Word's spurious sectPr artifacts.

The problem PDF→Word leaves behind:
  - Many embedded <w:sectPr> in body paragraphs (one per PDF page)
  - Default type is nextPage (forces page break)
  - Even after changing type to "continuous", Word STILL treats them as soft page breaks
    (Word-specific behavior; OOXML spec says continuous shouldn't break — LibreOffice agrees)
  - Only fix: REMOVE the sectPr element entirely (not just retype) for redundant ones.

This script:
  1. Detects "real" chapter break sectPrs (must keep as nextPage):
     - has <w:headerReference> or <w:footerReference>, OR
     - the next NON-EMPTY paragraph has a heading-style pStyle
       (auto-detect heading styles from styles.xml: any style with outlineLvl 0-3)
     - falls back to {Heading1, Heading2, Heading3, Title} if outlineLvl missing
  2. For all other sectPrs:
     - If type=continuous AND pgMar exactly matches body-default AND no header/footer ref:
       → DELETE the sectPr element entirely (leave the containing empty paragraph)
     - Otherwise: leave alone (preserves cover/abstract/ToC sections with unique pgMar)

Heuristic: when scanning for the "next paragraph" after a sectPr, SKIP empty BodyText
spacer paragraphs — they have a pStyle but no text. Stopping there would miss the
real heading that follows.

Usage:
  python3 fix_sectpr.py --docx FILE.docx --inplace
  python3 fix_sectpr.py --docx FILE.docx --out FIXED.docx
"""
import argparse, io, json, os, shutil, sys, zipfile
from lxml import etree

NSW = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W = '{%s}' % NSW

DEFAULT_HEADING_STYLES = {'Heading1','Heading2','Heading3','Title'}

def detect_heading_styles(styles_xml):
    """Find styleIds in styles.xml that have <w:outlineLvl val='0'|'1'|'2'|'3'/>."""
    out = set(DEFAULT_HEADING_STYLES)
    root = etree.fromstring(styles_xml)
    for s in root.iter(W+'style'):
        sid = s.get(W+'styleId')
        if not sid: continue
        for ol in s.iter(W+'outlineLvl'):
            try:
                lv = int(ol.get(W+'val','99'))
                if 0 <= lv <= 3:
                    out.add(sid); break
            except: pass
    return out

def next_non_empty_paragraph_style(p, body_children, heading_styles, max_lookahead=40):
    """Walk forward from p in body_children; return (pStyle, text_preview) of next paragraph
    that has non-empty text. Returns None if not found within max_lookahead."""
    try: i = body_children.index(p)
    except ValueError: return None
    for j in range(i+1, min(i+max_lookahead, len(body_children))):
        ch = body_children[j]
        if ch.tag == W+'p':
            text = ''.join(x.text or '' for x in ch.iter(W+'t')).strip()
            if not text: continue
            pPr = ch.find(W+'pPr')
            sty = None
            if pPr is not None:
                ps = pPr.find(W+'pStyle')
                if ps is not None: sty = ps.get(W+'val')
            return sty, text[:60]
        else:
            # table or other; find first inner paragraph with text
            for inner in ch.iter(W+'p'):
                t = ''.join(x.text or '' for x in inner.iter(W+'t')).strip()
                if t: return None, t[:60]
    return None

def pgmar_dict(sp):
    pgMar = sp.find(W+'pgMar')
    if pgMar is None: return {}
    return {k.split('}')[-1]: v for k, v in pgMar.attrib.items()}

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--docx', required=True)
    ap.add_argument('--inplace', action='store_true')
    ap.add_argument('--out')
    args = ap.parse_args()
    if not args.inplace and not args.out:
        print('Either --inplace or --out is required', file=sys.stderr); sys.exit(1)

    out_path = args.docx if args.inplace else args.out
    backup = os.path.splitext(args.docx)[0] + '.before-sectpr.docx'
    if not os.path.exists(backup):
        shutil.copyfile(args.docx, backup)
        print(f'Backup: {backup}')

    zin = zipfile.ZipFile(args.docx)
    doc_xml = zin.read('word/document.xml')
    styles_xml = zin.read('word/styles.xml')
    root = etree.fromstring(doc_xml)
    body = root.find(W+'body')

    heading_styles = detect_heading_styles(styles_xml)
    print(f'Detected heading styles: {sorted(heading_styles)}')

    # Body-level final sectPr defines the "main pgMar"
    final_sp = body.find(W+'sectPr')
    main_pgmar = pgmar_dict(final_sp) if final_sp is not None else {}
    print(f'Main pgMar (from final body sectPr): {main_pgmar}')

    # Collect all embedded sectPr-bearing paragraphs in order
    body_children = list(body)
    records = []  # (paragraph, sectPr)
    for p in body.iter(W+'p'):
        pPr = p.find(W+'pPr')
        if pPr is None: continue
        sp = pPr.find(W+'sectPr')
        if sp is not None:
            records.append((p, sp))
    print(f'Embedded sectPrs found: {len(records)}')

    deleted = kept_chapter = kept_unique = 0
    deletions = []
    keeps = []

    for p, sp in records:
        typ = sp.find(W+'type')
        type_val = typ.get(W+'val') if typ is not None else 'nextPage'
        has_ref = (sp.find(W+'headerReference') is not None) or (sp.find(W+'footerReference') is not None)
        mar = pgmar_dict(sp)

        info = next_non_empty_paragraph_style(p, body_children, heading_styles)
        next_sty = info[0] if info else None
        next_text = info[1] if info else ''

        # Decision tree
        if has_ref:
            kept_chapter += 1
            keeps.append(('header_ref', next_text))
            continue
        if next_sty in heading_styles:
            kept_chapter += 1
            keeps.append((f'heading={next_sty}', next_text))
            continue
        if mar != main_pgmar:
            # Unique pgMar (cover/abstract/ToC) — leave alone
            kept_unique += 1
            keeps.append(('unique_pgmar', next_text))
            continue
        # type doesn't matter here — same pgMar + no refs + no heading → spurious
        # Plan F: physically delete sectPr element
        p.find(W+'pPr').remove(sp)
        deleted += 1
        deletions.append((type_val, next_text))

    print(f'\nDeleted spurious sectPrs: {deleted}')
    print(f'Kept (real chapter breaks): {kept_chapter}')
    print(f'Kept (unique pgMar — cover/abstract/ToC): {kept_unique}')

    new_xml = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.namelist():
            if item == 'word/document.xml':
                zout.writestr(item, new_xml)
            else:
                zout.writestr(item, zin.read(item))
    open(out_path, 'wb').write(buf.getvalue())
    print(f'Wrote: {out_path}')

    # Audit report
    report_path = os.path.join(os.path.dirname(out_path), 'thesis_sectpr_report.md')
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write('# sectPr Cleanup Report\n\n')
        f.write(f'- Docx: `{out_path}`\n- Backup: `{backup}`\n\n')
        f.write(f'## Stats\n- Deleted: **{deleted}**\n- Kept (real chapter): **{kept_chapter}**\n- Kept (unique pgMar): **{kept_unique}**\n\n')
        f.write('## Kept as nextPage (forces page break)\n')
        for reason, txt in keeps:
            f.write(f'- ({reason}) → `{txt}`\n')
        f.write('\n## Deleted (spurious)\n')
        for typ, txt in deletions:
            f.write(f'- (was type={typ}) → `{txt}`\n')
    print(f'Report: {report_path}')

if __name__ == '__main__':
    main()
