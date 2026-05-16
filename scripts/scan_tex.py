#!/usr/bin/env python3
"""Scan a LaTeX thesis project, extract all formulas, generate placeholder copies.

Outputs:
  <PROJECT>/formula_placeholders.json   — list of {id, file, line, column, type, source}
  <PROJECT>_formula_placeholders_only/  — modified TeX copies with placeholders

Formula types extracted:
  - inline       : $...$ or \\(...\\)
  - display      : $$...$$ , \\[...\\] , equation, equation*, align, align*, gather, gather*, displaymath

Numbered display formulas (with equation number from labels) get a placeholder followed by
the number in the PDF output. We achieve this by injecting `\\qquad(N-N)` if the equation
environment has \\label{...}, where N-N is derived later by LaTeX's auto-numbering.
Actually simpler: keep the equation environment empty except for the placeholder, and let
\\eqref still work; the PDF renders "[公式 Fxxx]    (3-1)".

Usage:
  python3 scan_tex.py --project /path/to/project --tex-entry thesis.tex
"""
import argparse, json, os, re, shutil, sys
from pathlib import Path

# Patterns ordered: longest match first
ENV_NAMES = ['equation*','equation','align*','align','gather*','gather','displaymath']
ENV_PATTERN = re.compile(
    r'\\begin\{(' + '|'.join(re.escape(e) for e in ENV_NAMES) + r')\}(.*?)\\end\{\1\}',
    re.DOTALL
)
# \[ ... \]   and  \( ... \)
BRACKET_DISPLAY = re.compile(r'\\\[(.*?)\\\]', re.DOTALL)
BRACKET_INLINE  = re.compile(r'\\\((.*?)\\\)', re.DOTALL)
# $$ ... $$ (must come before single $...$)
DOLLAR_DISPLAY = re.compile(r'\$\$(.+?)\$\$', re.DOTALL)
# $...$ — careful with $$ already handled and \$ escapes; we do a simple scan with state
DOLLAR_INLINE_HINT = '$'

VERBATIM_ENVS = re.compile(r'\\begin\{(verbatim|lstlisting|minted|Verbatim)\}.*?\\end\{\1\}', re.DOTALL)
COMMENT_LINE = re.compile(r'(?<!\\)%.*?(?=\n|$)')

def strip_protected_regions(text):
    """Return text with verbatim and line comments masked (replaced by spaces of equal length).
    Preserves line numbers and column positions."""
    def mask(m):
        return ' ' * (m.end() - m.start())
    text = VERBATIM_ENVS.sub(mask, text)
    text = COMMENT_LINE.sub(mask, text)
    return text

def find_input_files(tex_path, project_root, seen):
    """Recursively resolve \\input{} and \\include{} starting from tex_path.
    Returns ordered list of (relative_path_from_project, absolute_path)."""
    abs_path = Path(tex_path).resolve()
    if abs_path in seen: return []
    seen.add(abs_path)
    rel = abs_path.relative_to(project_root) if abs_path.is_relative_to(project_root) else abs_path
    out = [(str(rel), abs_path)]
    try:
        content = abs_path.read_text(encoding='utf-8', errors='replace')
    except FileNotFoundError:
        return out
    safe = strip_protected_regions(content)
    for m in re.finditer(r'\\(?:input|include)\{([^}]+)\}', safe):
        name = m.group(1).strip()
        # try with and without .tex
        candidates = [
            project_root / (name + '.tex'),
            project_root / name,
            abs_path.parent / (name + '.tex'),
            abs_path.parent / name,
        ]
        for c in candidates:
            if c.exists() and c.suffix == '.tex':
                out.extend(find_input_files(c, project_root, seen))
                break
    return out

def scan_dollars(text):
    """Find $...$ inline math, respecting $$..$$ (return empty) and \\$ escapes.
    Returns list of (start, end, content)."""
    results = []
    i = 0
    n = len(text)
    while i < n:
        c = text[i]
        if c == '\\' and i+1 < n:
            i += 2
            continue
        if c == '$':
            # Check double dollar
            if i+1 < n and text[i+1] == '$':
                # Skip the $$...$$ block
                end = text.find('$$', i+2)
                if end == -1:
                    i = n; break
                i = end + 2
                continue
            # single $
            start = i
            j = i + 1
            while j < n:
                if text[j] == '\\':
                    j += 2; continue
                if text[j] == '$':
                    break
                j += 1
            if j < n:
                results.append((start, j+1, text[start+1:j]))
                i = j + 1
            else:
                break
        else:
            i += 1
    return results

def line_col(text, pos):
    line = text.count('\n', 0, pos) + 1
    last_nl = text.rfind('\n', 0, pos)
    col = pos - (last_nl + 1) + 1
    return line, col

def extract_formulas(text):
    """Returns ordered list of (start, end, type, source). 'source' includes the wrapper."""
    safe = strip_protected_regions(text)
    spans = []
    # Display environments
    for m in ENV_PATTERN.finditer(safe):
        spans.append((m.start(), m.end(), 'display', text[m.start():m.end()]))
    # \[ ... \]
    for m in BRACKET_DISPLAY.finditer(safe):
        spans.append((m.start(), m.end(), 'display', text[m.start():m.end()]))
    # $$ ... $$
    for m in DOLLAR_DISPLAY.finditer(safe):
        spans.append((m.start(), m.end(), 'display', text[m.start():m.end()]))
    # \( ... \)
    for m in BRACKET_INLINE.finditer(safe):
        spans.append((m.start(), m.end(), 'inline', text[m.start():m.end()]))
    # $...$
    for start, end, content in scan_dollars(safe):
        spans.append((start, end, 'inline', text[start:end]))
    # Sort by start, drop overlapping (keep first / longer)
    spans.sort(key=lambda s: (s[0], -(s[1]-s[0])))
    filtered = []
    last_end = -1
    for s in spans:
        if s[0] >= last_end:
            filtered.append(s)
            last_end = s[1]
    return filtered

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--project', required=True, help='Project root directory')
    ap.add_argument('--tex-entry', required=True, help='Main TeX file (relative to project)')
    ap.add_argument('--out-suffix', default='_formula_placeholders_only', help='Output dir suffix')
    args = ap.parse_args()

    project = Path(args.project).resolve()
    entry = (project / args.tex_entry).resolve()
    if not entry.exists():
        print(f'ERROR: {entry} does not exist', file=sys.stderr); sys.exit(1)

    seen = set()
    files = find_input_files(entry, project, seen)
    print(f'Discovered {len(files)} TeX files reachable from {args.tex_entry}')

    json_records = []
    next_id = 1
    modified_files = {}  # abs_path → modified content

    for rel_path, abs_path in files:
        try:
            content = abs_path.read_text(encoding='utf-8', errors='replace')
        except Exception as e:
            print(f'WARN: cannot read {abs_path}: {e}'); continue
        spans = extract_formulas(content)
        if not spans:
            modified_files[abs_path] = content
            continue
        # Build modified content: replace each span with placeholder string
        out_parts = []
        last_end = 0
        for start, end, ftype, source in spans:
            fid = f'F{next_id:03d}'
            line, col = line_col(content, start)
            json_records.append({
                'id': f'公式 {fid}', 'file': str(rel_path),
                'line': line, 'column': col, 'type': ftype, 'source': source,
            })
            out_parts.append(content[last_end:start])
            placeholder = f'[公式 {fid}]'
            if ftype == 'display':
                # Wrap with same environment structure so equation numbering stays:
                # Heuristic — if source begins with \begin{equation} or \[ , wrap placeholder accordingly
                if source.lstrip().startswith('\\begin{equation}'):
                    out_parts.append('\\begin{equation}\n' + placeholder + '\n\\end{equation}')
                elif source.lstrip().startswith('\\begin{equation*}'):
                    out_parts.append('\\begin{equation*}\n' + placeholder + '\n\\end{equation*}')
                elif source.lstrip().startswith('\\begin{align}'):
                    out_parts.append('\\begin{align}\n' + placeholder + '\n\\end{align}')
                elif source.lstrip().startswith('\\['):
                    out_parts.append('\\[' + placeholder + '\\]')
                elif source.lstrip().startswith('$$'):
                    out_parts.append('$$' + placeholder + '$$')
                else:
                    out_parts.append(placeholder)
            else:
                # inline: replace with bracketed placeholder (no math wrapper)
                out_parts.append(placeholder)
            next_id += 1
            last_end = end
        out_parts.append(content[last_end:])
        modified_files[abs_path] = ''.join(out_parts)

    # Write JSON
    json_path = project / 'formula_placeholders.json'
    json_path.write_text(json.dumps(json_records, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'Wrote {json_path} with {len(json_records)} formulas')

    # Write modified TeX copies
    out_root = project.parent / (project.name + args.out_suffix)
    if out_root.exists():
        print(f'Removing existing output dir: {out_root}')
        shutil.rmtree(out_root)
    shutil.copytree(project, out_root, ignore=shutil.ignore_patterns(
        '*.aux', '*.log', '*.bbl', '*.blg', '*.toc', '*.out', '*.pdf', '*.synctex.gz',
        '.git', '__pycache__'
    ))
    for abs_path, new_content in modified_files.items():
        rel = abs_path.relative_to(project)
        target = out_root / rel
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(new_content, encoding='utf-8')
    print(f'Wrote modified TeX project to {out_root}')

if __name__ == '__main__':
    main()
