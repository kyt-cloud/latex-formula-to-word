# latex-formula-to-word

End-to-end pipeline for converting a Chinese LaTeX thesis into a Word `.docx` where every formula is a **native Word OMML object**, while preserving PDF-consistent layout.

Built as a [Claude Code skill](https://docs.claude.com/en/docs/claude-code/skills). Invoke via `/latex-formula-to-word` inside Claude Code, or run the underlying scripts directly from the command line.

## What problem it solves

Converting a LaTeX thesis to Word directly via Pandoc loses page headers, margins, and section structure. Converting the rendered PDF to Word via Acrobat/Word/Mathpix preserves layout but corrupts every formula. This pipeline keeps the best of both:

1. Rewrite the TeX so every formula becomes a tracked placeholder `[公式 Fxxx]`
2. Compile that placeholder TeX to PDF — gives you the correct layout
3. You convert that PDF to docx using whichever tool keeps your layout best
4. This pipeline swaps every placeholder for a native Word OMML formula via run-surgery
5. Then it scrubs PDF→Word's spurious section-break artifacts that cause Word to scatter blank pages

## Install

```bash
git clone <this-repo> ~/.claude/skills/latex-formula-to-word
```

After that, `/latex-formula-to-word` shows up as an available skill in any Claude Code session.

To use the scripts standalone (without Claude Code), see [Standalone usage](#standalone-usage) below.

## Requirements

| Tool | Why | Install |
|---|---|---|
| Python ≥ 3.9 | Runs the four scripts | system |
| `lxml` | OOXML editing | `pip install lxml` |
| `pandoc` ≥ 3.0 | LaTeX → OMML conversion | `brew install pandoc` |
| `xelatex` + `biber` | Compile placeholder TeX → PDF | TeX Live / MacTeX |
| LibreOffice | Final render verification | `brew install --cask libreoffice` |
| `PyMuPDF` | Per-page fill-rate measurement | `pip install pymupdf` |

The skill targets macOS but the scripts run on any POSIX system where the dependencies are present.

## Pipeline overview

```
TeX project ─[Phase 1: scan_tex.py]─▶ formula_placeholders.json
                                       + <project>_formula_placeholders_only/
                                                 │
                                                 ▼
                                       [Phase 2: xelatex 4-pass]
                                                 │
                                                 ▼
                                       placeholder PDF
                                                 │
                                                 ▼  (you convert manually)
                                                 │
                                       docx with [公式 Fxxx] text
                                                 │
                                                 ▼
                                       [Phase 4: replace_docx.py]
                                                 │
                                                 ▼
                                       docx with native OMML
                                                 │
                                                 ▼
                                       [Phase 5: fix_sectpr.py — Plan F]
                                                 │
                                                 ▼
                                       docx with clean layout
                                                 │
                                                 ▼
                                       [Phase 6: verify_render.py]
                                                 │
                                                 ▼
                                       per-page fill audit report
```

## Designed-in defaults (do not skip without override)

1. **Pandoc compile failure → terminate.** Any formula pandoc rejects causes the pipeline to stop. Fix the TeX macro and re-run.
2. **Heading detection skips empty BodyText spacer paragraphs.** A common PDF→Word habit is putting empty spacer paragraphs between section breaks and real headings. The detector ignores them.
3. **Spurious `type=continuous` sectPrs are physically deleted (Plan F).** Word treats `continuous` section breaks as soft page breaks even though OOXML says it shouldn't — LibreOffice respects continuous. The only fix that works in Word is removing the sectPr element entirely.
4. **LibreOffice pre-renders for final fill-rate audit.** Flags pages with content fill < 80% so you can spot remaining figure-driven or chapter-end blanks.

## Standalone usage

```bash
# Phase 1 — scan TeX
python3 scripts/scan_tex.py \
  --project /path/to/thesis \
  --tex-entry thesis.tex

# Phase 2 — compile placeholder TeX (run inside your thesis project)
cd /path/to/thesis_formula_placeholders_only
xelatex thesis && biber thesis && xelatex thesis && xelatex thesis

# Phase 3 — manually: convert thesis.pdf → thesis.docx (Word, Acrobat, ...)

# Phase 4 — replace placeholders
python3 scripts/replace_docx.py \
  --docx /path/to/thesis.docx \
  --json /path/to/thesis/formula_placeholders.json \
  --out  /path/to/thesis.replaced.docx

# Phase 5 — clean sectPr artifacts
python3 scripts/fix_sectpr.py \
  --docx /path/to/thesis.replaced.docx \
  --inplace

# Phase 6 — verify
python3 scripts/verify_render.py \
  --docx /path/to/thesis.replaced.docx
```

## Key technical decisions

### 1×3 equation table for numbered display formulas

Numbered display formulas (e.g. `(3-57)`) are replaced with a `1×3` borderless table:

| Cell | Width (twips) | Content | Notes |
|---|---|---|---|
| Left | 200 | empty | balancing padding |
| Middle | 8200 | OMML formula | `line=240`, tight margins |
| Right | 1100 | `(N-N)` text | `<w:noWrap/>` + tight `<w:tcMar>` 50/50/0/0 |

Right cell `noWrap` prevents the equation number from collapsing into a vertical character stack when the cell is narrow. The `1×3` table reliably keeps "centered formula + right-aligned number" together without the formula bleeding into the number area.

### Drop redundant `<m:oMath>` from pandoc output

Pandoc emits display math as:

```xml
<m:oMathPara>
  <m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>
  <m:oMath>...</m:oMath>
</m:oMathPara>
```

A naive `findall('.//m:oMath')` matches both wrapper and inner, inserting the formula twice. The cache extractor keeps only `<m:oMathPara>` for display formulas.

### Auto-detect heading styles from `styles.xml`

Different Chinese thesis templates name heading styles differently (`Heading1`, `标题1`, `章节`, …). The detector reads `styles.xml` and treats any style with `<w:outlineLvl val="0..3"/>` as a heading, with `{Heading1, Heading2, Heading3, Title}` as fallback.

### Don't unify pgMar across sections

Each section's original `pgMar` (top/bottom/left/right/header/footer) is preserved. Only sectPrs whose pgMar exactly matches the body-default are eligible for Plan F deletion. Cover page, abstract, and ToC sections keep their distinct margins.

## Outputs after a clean run

Inside the user's docx directory:

- `<name>.replaced.docx` — final output with OMML + clean layout
- `<name>.before-replace.docx` — pre-replacement backup
- `<name>.before-sectpr.docx` — post-replacement, pre-Plan-F backup
- `thesis_replace_report.md` — replacement count + missing IDs
- `thesis_sectpr_report.md` — list of kept/deleted section breaks
- `thesis_render_report.md` — LibreOffice page-fill audit

## Known limitations

- PDF→Word sometimes packs multiple consecutive `equation` blocks into a single multi-column table (e.g. a 2×3 grid holding three formulas together). The replacer handles them as inline replacements with cell width locked, but if a formula is wider than its cell it wraps. Splitting these merged formula tables is a separate manual pass.
- Pre-existing OCR alignment errors in the PDF→Word output (e.g., a sentence split mid-word with one half right-aligned) carry through. Not fixed here.
- Word and LibreOffice paginate slightly differently. LibreOffice fill-rate audit is approximate; the user's Word view is final judgement.

## License

MIT
