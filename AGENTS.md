# aholo-

Python tooling that generates Microsoft Word (`.docx`) deliverables for the「地块智能快评」(Land Parcel Intelligent Quick Assessment) product concept.

## Cursor Cloud specific instructions

- This repo is a set of standalone Python scripts (no app/server, no tests, no linter config). The only third-party dependency is `python-docx` (imported as `docx`).
- Generate the deliverables (the end-to-end "run"):
  - `python3 scripts/generate_product_proposal.py` → writes `docs/产品方案/地块智能快评-产品方案.docx`
  - `python3 scripts/generate_kuaiping_word.py` → writes `docs/交付模板/地块前期研判报告（快评版）模板.docx`
- Both scripts write to hardcoded absolute paths under `/workspace/docs/...` and auto-create output dirs. Regenerating rewrites the committed `.docx` files with byte-identical-size but different zip metadata, so they show as `modified` in git — these are build artifacts; discard with `git checkout -- docs/` unless you intentionally changed content.
- Only `generate_product_proposal.py` imports `scripts/docx_utils.py`; `generate_kuaiping_word.py` inlines its own helpers.
- No headless Word/LibreOffice is installed. To visually inspect output, render with `python-docx` (+ Pillow and the CJK font at `/usr/share/fonts/truetype/wqy/wqy-microhei.ttc`) rather than opening in Word.
