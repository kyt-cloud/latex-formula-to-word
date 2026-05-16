---
name: latex-formula-to-word
description: End-to-end pipeline that turns a LaTeX Chinese thesis into a Word docx with native OMML formulas + clean Word layout. Scans TeX, generates placeholder PDF, waits for user to convert PDF→Word, replaces placeholders with OMML, then surgically cleans up PDF→Word sectPr artifacts (Word-specific blank-page bug). Use ONLY when user invokes `/latex-formula-to-word`. Manual trigger; never auto-fire.
argument-hint: "[optional: project directory; defaults to cwd]"
allowed-tools: ["Read", "Bash", "Glob", "Edit", "Write", "AskUserQuestion"]
---

# Thesis Formula Replace (LaTeX → Word with Native OMML)

End-to-end pipeline for Chinese thesis projects that converts a LaTeX thesis into a Word docx where every formula is a native Word OMML object, while preserving PDF-consistent layout.

**Designed-in defaults (do not skip without user override)**:
1. Pandoc compile failure → terminate (fail-fast)
2. Heading detection skips empty BodyText spacer paragraphs
3. Plan F applied: delete redundant `type=continuous` sectPrs (Word treats them as soft page breaks even though spec says they shouldn't)
4. LibreOffice pre-render verifies per-page fill rate at the end

## Inputs the user must provide before/during the run

- Project directory containing TeX files (typically with `body/`, `thesis.tex`, `csust*.cls`)
- A PDF→Word conversion at one point (you tell the user when; they use Word/Acrobat/Mathpix/etc. themselves)

## The six phases

### Phase 1 — Scan TeX, generate placeholder copies

Run:
```bash
python3 ~/.claude/skills/latex-formula-to-word/scripts/scan_tex.py \
  --project <PROJECT_DIR> \
  --tex-entry <MAIN_TEX>      # e.g. thesis.tex
```

What it does:
- Walks all `.tex` files reachable from the entry point (resolves `\input`/`\include`)
- Extracts every formula: inline `$...$` / `\(...\)`, display `$$...$$` / `\[...\]`, environments `equation`, `equation*`, `align`, `align*`, `gather`, `gather*`, `displaymath`
- Assigns F001, F002, … sequentially in source order
- Writes `formula_placeholders.json` (list of {id, file, line, column, type, source})
- Writes modified TeX copies under `<PROJECT_DIR>_formula_placeholders_only/` where each formula is replaced by `[公式 Fxxx]` (numbered display gets the equation number to its right; the numbering macros are preserved)

**Skip rule**: if the user already has `formula_placeholders.json`, ask whether to reuse or rebuild. Default: reuse.

### Phase 2 — Compile placeholder TeX to PDF

Run XeLaTeX 4-pass on the placeholder copy:
```bash
cd <PROJECT_DIR>_formula_placeholders_only
xelatex -interaction=nonstopmode <MAIN_TEX>
biber <MAIN_TEX without ext>
xelatex -interaction=nonstopmode <MAIN_TEX>
xelatex -interaction=nonstopmode <MAIN_TEX>
```

If `make` is available with the right target, prefer `make <target>` from the original project after symlinking — see project's CLAUDE.md.

### Phase 3 — Hand off to user for PDF→Word conversion

Tell the user **exactly**:

> 占位符 PDF 已生成：`<path>.pdf`
>
> 请用你偏好的 PDF→Word 工具（Word、Acrobat 等）把它转成 docx，**不要用 Mathpix**（会丢页眉版式）。完成后告诉我 docx 路径。

Wait for the user to come back with the docx path. Don't try to auto-convert.

### Phase 4 — Replace placeholders with OMML in docx

Run:
```bash
python3 ~/.claude/skills/latex-formula-to-word/scripts/replace_docx.py \
  --docx <USER_DOCX> \
  --json <PROJECT>/formula_placeholders.json \
  --out <USER_DOCX_DIR>/<NAME>.replaced.docx
```

Internally this script:
1. **Dry-run pandoc** on all formulas. If any fail, stop immediately and report the failure list with file/line. **DO NOT auto-fix or silently skip.** User must fix the TeX macro or accept exclusion before re-running.
2. Extracts OMML for each formula into `cache/`. **For display formulas, keeps only `<m:oMathPara>` and drops the standalone `<m:oMath>`** (pandoc emits both due to its display-math wrapping; double-insertion causes the formula to render twice in Word).
3. Walks `<w:body>` and every `<w:tc>`, finds `[公式 Fxxx]` occurrences:
   - **Numbered display** (paragraph matches `^\s*\[公式 Fxxx\]\s*\(\d+[-–]\d+[a-z]?\)\s*$`): replaces whole paragraph with a 1×3 borderless table, column widths **`200 / 8200 / 1100` twips**. Right cell carries the equation number with `<w:noWrap/>` and tight `<w:tcMar>` (50/50/0/0). Middle cell carries the OMML, paragraph spacing `line=240, before=0, after=0`.
   - **Inside table cell**: lock the parent table's layout to fixed and copy gridCol widths into each `<w:tcW>` BEFORE inserting OMML (otherwise widening cells will rebalance the table).
   - **Inline**: run surgery — split runs at placeholder boundaries, replace the matched range with OMML siblings, preserve surrounding rPr.

Run surgery basics for split runs (placeholders are typically split across multiple `<w:r>` by PDF→Word):
```
1. Concatenate run texts → build offset map (run_idx, offset_within_run) per char.
2. For span [start, end]: split run at `end` first, then at `start`.
3. Delete runs between the two splits, insert OMML node(s) at that position.
```

Output report under `<USER_DOCX_DIR>/thesis_replace_report.md` listing:
- Counts (inline / numbered / in_table)
- Any IDs missing from OMML cache
- Pandoc failures (should be zero if Phase 4 didn't abort)

### Phase 5 — Clean Word layout artifacts (Plan F + Heading detection)

Run:
```bash
python3 ~/.claude/skills/latex-formula-to-word/scripts/fix_sectpr.py \
  --docx <USER_DOCX_DIR>/<NAME>.replaced.docx \
  --inplace
```

Internally:
1. Backup the docx to `<NAME>.before-sectpr.docx`.
2. Scan all embedded `<w:sectPr>` in the body. For each, find the next **non-empty** paragraph (skip empty BodyText spacer paragraphs).
3. Classify each sectPr:
   - **Keep `nextPage`** if either:
     - It already has `<w:headerReference>` or `<w:footerReference>` (real chapter break), OR
     - The next non-empty paragraph has `pStyle` in `{Heading1, Heading2, Heading3, Title}` plus the project's localized heading style names (auto-detected from `word/styles.xml` — see "Heading style detection" below)
   - **Delete the sectPr element entirely** (do NOT just retype to continuous) when ALL of:
     - Type is or would be continuous
     - `<w:pgMar>` exactly matches the body-level default
     - No header/footer reference
4. Leave 1–3 "unique pgMar" continuous sectPrs alone (cover page, abstract, ToC area — they intentionally have different margins).

**Why delete instead of retype**: Word treats `type=continuous` sectPrs as soft page breaks even though OOXML spec says it shouldn't. LibreOffice respects continuous; Word doesn't. Only removing the element entirely makes Word stop creating a section boundary.

### Heading style detection (multi-template support)

Different Chinese thesis templates name heading styles differently: `Heading1`, `标题1`, `章`, `一级标题`, etc. Detect dynamically:
- Read `word/styles.xml`
- A style is a "heading" if its `<w:pPr>` contains `<w:outlineLvl val="0..3"/>`
- Add those styleIds to the heading detection set

The default set hard-codes `Heading1, Heading2, Heading3, Title` for templates that don't set outlineLvl.

### Phase 6 — LibreOffice render verification

Run:
```bash
python3 ~/.claude/skills/latex-formula-to-word/scripts/verify_render.py \
  --docx <USER_DOCX_DIR>/<NAME>.replaced.docx
```

Internally:
1. Calls `soffice --headless --convert-to pdf` (requires LibreOffice installed; on macOS at `/Applications/LibreOffice.app/Contents/MacOS/soffice` or via Homebrew at `/opt/homebrew/bin/soffice`).
2. Uses PyMuPDF to inspect each page's content extent vs page usable area.
3. Reports any page with content fill < 80% (excluding header/footer band).
4. Maps PDF page number → header label "第 X 页" for user-friendly reference.

If verification finds blank-prone pages:
- A page filled 60–80%: usually figure-driven (image too tall, pushed to next page). Cannot fix without resizing images. Report and move on.
- A page filled < 40%: probably a kept-nextPage sectPr that user wants to convert to continuous. Show the next-paragraph pStyle and text, ask user whether to demote.
- Note: Word and LibreOffice paginate differently. LibreOffice reports are a sanity check, not gospel — final judgment is in the user's Word.

## Hard rules

1. **Never overwrite the user's original docx or TeX**. Every phase outputs to a copy with a suffix.
2. **Every destructive change has a backup**: `before-replace`, `before-sectpr`, etc.
3. **No silent failures**. If pandoc rejects a formula, stop. If a sectPr has unexpected pgMar, leave it and report.
4. **Don't unify pgMar across sections**. Each section's original pgMar is preserved. Only sectPrs with the body-default pgMar are eligible for Plan F deletion.

## Decision points where you should ask the user

- **Phase 1**: if `formula_placeholders.json` already exists — reuse or rebuild?
- **Phase 4**: if pandoc fails on any formula — show the failed IDs and source, stop and let user fix TeX macros.
- **Phase 5**: if any Heading3 sub-section break looks borderline (e.g., a section heading mid-chapter) — show its body_idx + next text, ask whether to demote.
- **Phase 6**: if any page has fill < 40% — show context, ask whether to demote a sectPr or accept.

## Known limitations

- PDF→Word can pack multiple consecutive `equation` blocks into a single multi-column table (e.g., a 2×3 grid holding 3 formulas together). The replace_docx.py script handles them as "in_table" inline replacements (with cell width locked), but if a formula is wider than its cell, it will wrap. There's no automatic fix in this pipeline — those cases need a separate "split the merged formula table" pass (not implemented here).
- Pre-existing OCR alignment errors in the PDF→Word output (e.g., a paragraph that got split mid-sentence with one half right-aligned) are visible after replacement but were already present. This skill does not fix them.
- Word renders differently from LibreOffice for some sectPr edge cases. The pipeline targets Word behavior; LibreOffice verification is approximate.

## What gets produced

In the user's docx directory after a clean run:
- `<NAME>.replaced.docx` — final output with OMML + clean layout
- `<NAME>.before-replace.docx` — pre-replacement backup
- `<NAME>.before-sectpr.docx` — post-replacement, pre-Plan-F backup
- `thesis_replace_report.md` — replacement count + any caveats
- `thesis_render_report.md` — LibreOffice page-fill audit
