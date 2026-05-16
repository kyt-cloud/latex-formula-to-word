#!/usr/bin/env python3
"""LibreOffice pre-render verification: convert docx → PDF and audit per-page fill rate.

Reports pages whose content fill drops below thresholds. Word and LibreOffice
paginate slightly differently; this is a sanity check, not gospel.

Requires:
  - soffice (LibreOffice) on PATH or at /Applications/LibreOffice.app/Contents/MacOS/soffice
  - PyMuPDF (pip install pymupdf)

Usage:
  python3 verify_render.py --docx FILE.docx [--threshold 80]
"""
import argparse, os, re, shutil, subprocess, sys, tempfile

def find_soffice():
    if shutil.which('soffice'): return 'soffice'
    for p in ('/Applications/LibreOffice.app/Contents/MacOS/soffice',
              '/opt/homebrew/bin/soffice',
              '/usr/bin/libreoffice'):
        if os.path.exists(p): return p
    return None

def render_pdf(docx, out_dir):
    soffice = find_soffice()
    if not soffice:
        print('ERROR: soffice (LibreOffice) not found. Install LibreOffice.', file=sys.stderr)
        sys.exit(1)
    r = subprocess.run([soffice, '--headless', '--convert-to', 'pdf',
                        '--outdir', out_dir, docx],
                       capture_output=True, text=True)
    if r.returncode != 0:
        print('soffice failed:', r.stderr, file=sys.stderr); sys.exit(1)
    base = os.path.splitext(os.path.basename(docx))[0]
    pdf = os.path.join(out_dir, base + '.pdf')
    if not os.path.exists(pdf):
        print(f'PDF not produced at {pdf}', file=sys.stderr); sys.exit(1)
    return pdf

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--docx', required=True)
    ap.add_argument('--threshold', type=float, default=80.0,
                    help='Flag pages with fill below this %% (default 80)')
    ap.add_argument('--severe', type=float, default=40.0,
                    help='Mark as severe below this %% (default 40)')
    args = ap.parse_args()

    try:
        import fitz
    except ImportError:
        print('ERROR: PyMuPDF not installed. Run: pip install pymupdf', file=sys.stderr)
        sys.exit(1)

    out_dir = tempfile.mkdtemp(prefix='thesis_verify_')
    pdf_path = render_pdf(args.docx, out_dir)
    doc = fitz.open(pdf_path)
    total = len(doc)
    print(f'Rendered {total} pages → {pdf_path}\n')

    # Map "第 X 页" label → PDF page if present
    header_label = {}
    for pn, p in enumerate(doc, 1):
        m = re.search(r'第\s*(\d+)\s*页', p.get_text())
        if m:
            header_label[pn] = m.group(1)

    rows = []
    for pn, p in enumerate(doc, 1):
        ph = p.rect.height
        blocks = sorted([b for b in p.get_text('blocks') if b[4].strip()], key=lambda b: b[1])
        # Filter to "main content area": skip top ~8% (header) and bottom ~8% (footer)
        mb = [b for b in blocks if b[3] > ph*0.08 and b[1] < ph*0.92]
        if not mb:
            rows.append((pn, 0, '', '')); continue
        top = min(b[1] for b in mb); bot = max(b[3] for b in mb)
        fill = (bot - top) / (ph * 0.84) * 100
        first = mb[0][4].split('\n')[0][:60].strip()
        last  = mb[-1][4].strip().split('\n')[-1][:60]
        rows.append((pn, fill, first, last))

    # Report
    report_path = os.path.join(os.path.dirname(os.path.abspath(args.docx)),
                                'thesis_render_report.md')
    severe = []
    low = []
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write(f'# Render Verification Report\n\n')
        f.write(f'- Docx: `{args.docx}`\n- Rendered PDF: `{pdf_path}`\n- Pages: **{total}**\n\n')
        f.write(f'## Pages with fill < {args.threshold:.0f}%\n\n')
        f.write('| PDF page | 第 X 页 | Fill | First line | Last line |\n|---|---|---|---|---|\n')
        for pn, fill, first, last in rows:
            label = header_label.get(pn, '?')
            if fill < args.severe:
                severe.append((pn, fill, label, first, last))
            elif fill < args.threshold:
                low.append((pn, fill, label, first, last))
            if fill < args.threshold:
                f.write(f'| {pn} | {label} | {fill:.0f}% | `{first}` | `{last}` |\n')
    print(f'\nReport: {report_path}')

    print(f'\n=== SEVERE (fill < {args.severe:.0f}%): {len(severe)} pages ===')
    for pn, fill, label, first, last in severe:
        print(f'  p{pn} (第 {label} 页): fill={fill:.0f}%  first={first!r}')
    print(f'\n=== LOW (fill {args.severe:.0f}-{args.threshold:.0f}%): {len(low)} pages ===')
    for pn, fill, label, first, last in low:
        print(f'  p{pn} (第 {label} 页): fill={fill:.0f}%  first={first!r}')

if __name__ == '__main__':
    main()
