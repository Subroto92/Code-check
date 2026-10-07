# Derby full EGD content analysis

Builds `Derby_Full_EGD_PO_Brand_Preference_Content_Analysis.xlsx` (repo root) from the files in `inputs/`.

## Rebuild

```bash
python derby_analysis/build_workbook.py            # writes the workbook to the repo root
python derby_analysis/build_workbook.py out.xlsx   # or to a path of your choice
```

Needs Python 3 and `openpyxl`. Theme counts are Excel formulas, so they calculate when the file opens in Excel.

## Files

| Path | What it is |
|---|---|
| `build_workbook.py` | Python builder for the 8-sheet workbook |
| `inputs/analysis_content.json` | Group profiles, 18 themes and 27 context evidence rows (from the original GPT script) |
| `inputs/egd_section_coverage.json` | 88 EGD sections and paragraph bands |
| `inputs/prior_analysis_extract.json` | 34 brand-preference evidence rows and 4 PO rows from the earlier workbook |
| `inputs/transcript_audit.json` | Full extracted text of the 8 transcripts (feeds the Source Audit sheet) |
| `original_gpt_scripts/` | The ChatGPT scripts this was ported from, kept for reference |

The original `.mjs` builder needs `@oai/artifact-tool`, which only exists inside ChatGPT. `build_workbook.py` produces the same content with openpyxl. One change: the Source Audit sheet shows text rows, character counts and a SHA-256 of the extracted text, not file bytes and file hashes, because the raw transcript files weren't available.
