# Arabic Administrative Decision OCR

Extract structured data from scanned Algerian administrative decisions (مقررات إدارية) and export it to Excel. The script renders PDF pages as images and sends them to a vision model running on local Ollama at `http://127.0.0.1:11434`. It also accepts JPG, JPEG, PNG, BMP, WebP, and TIFF files.

The workbook uses Arabic headers and a right-to-left layout. It includes decision numbers and dates, employee names, ranks, grades, appointment details, visa details, and source page references where available.

## Quick start on Windows

Install Python 3.12 and Ollama, then start Ollama. Run these commands in PowerShell from the project directory.

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install PyMuPDF Pillow openpyxl
ollama pull qwen3.5:9b
New-Item -ItemType Directory -Force -Path .\decisions
```

Copy one scan into `decisions\` to try a small run. The first command below displays the available options. The second processes the scan and creates an Excel workbook.

```powershell
.\.venv\Scripts\python.exe extract_decisions.py --help
.\.venv\Scripts\python.exe extract_decisions.py --limit 1
```

To process all supported files in `decisions\`, run the script without `--limit`. Use `--model` to select another locally installed vision model if you want to compare speed and accuracy.

## Review the workbook

Each run that reaches the export step writes a new timestamped workbook in `exports\`.

| Sheet | Contents |
| --- | --- |
| **المقررات** | Extracted decisions, source filenames and pages, review status, and possible duplicates |
| **للمراجعة** | Records with specific reasons to review |
| **أخطاء المعالجة** | Files or pages that raised a processing error |

Review the workbook against the original scans, including rows without specific review flags. The model can omit or misread decision numbers, stamps, names, and dates. Check a few scans by hand before processing a batch. Fresh OCR with the default model can be slow on CPU and has timed out on a single scan during local testing. Documents listing several employees in one decision need manual review because the export uses one person per record.

## Multi-page decisions

The script can merge an adjacent page when the model marks it as a continuation and the available decision identifiers are compatible. Automatic merges appear in the review sheet.

To specify the page boundaries yourself, create `page_groups.json` in the project directory. In this example, a five-page PDF contains three decisions.

```json
{
  "scans.pdf": [[1, 2], [3], [4, 5]]
}
```

Use the actual PDF filename and list each page exactly once, starting at 1.

## Rotated scans and small text

Add `rotations.json` to set the orientation of a file or a specific PDF page. Angles are counter-clockwise and can be 0, 90, 180, or 270 degrees.

```json
{
  "scans.pdf#p2": 90,
  "photo.jpg": 270
}
```

Use `--detail` to check cropped sections for missing fields. It makes additional model calls. Readings recovered from crops and disagreements are flagged; their candidate values appear in review warnings rather than being treated as certain values.

Changes to either prompt, the response schema, or the actual `--num-ctx` value automatically use a new cache key. Use `--force` to reread otherwise identical cached pages. Old cache files are retained. The default context is 12288 to leave more room for the detailed Arabic prompt and response; local model memory and speed still depend on the selected model and hardware.

## Extraction policy

Edit `SYSTEM_PROMPT` and `CROP_PROMPT` in `prompts.py`; the application imports them directly. Missing grades and effective dates stay empty. Grades are not calculated from index numbers. Arabic letter categories are accepted, and higher-function divisions (`القسم`) are not seniority grades. Ambiguous name splits leave both surname and given name empty while retaining a clearly readable full name. Uncertain values are kept as review candidates and omitted from extraction fields.

## Compare prompts before training

Use a fixed set of upright images with human-checked JSON answers. The local `evaluate_extraction.py` tool only evaluates reference rows whose `human_verified` flag is explicitly `true`. It does not treat assistant drafts as ground truth. It reports exact text accuracy, nonempty-field accuracy, unsupported fills, and omissions; fields still marked uncertain in the reference are skipped.

For the private archive dataset, a proposed review set is in `exports/docs_training/evaluation/review_set.jsonl`. Correct each selected answer against its image before setting its `human_verified` flag. Then run:

```powershell
.\.venv\Scripts\python.exe evaluate_extraction.py --references exports/docs_training/evaluation/review_set.jsonl --image-root exports/docs_training --baseline-prompt exports/docs_training/history/before_explicit_evidence_v1/project_system_prompt.txt --output exports/docs_training/evaluation/results
```

Ollama must be running with the selected model installed. Both prompts use the same current extraction pipeline, context, and images, with crops disabled, so this compares the system prompts. It does not measure historical pipeline behavior or crop-prompt improvements. No training is started by this command.

Run the regression checks with `.\.venv\Scripts\python.exe -m unittest discover -s tests -v`.

On reruns, `--limit` counts completed OCR pages that were processed during the run. Cached pages can still appear in the workbook, and failed attempts can make the script try more pages than the limit.

### CPU timeouts

The default wait is 3600 seconds per Ollama request. On a CPU, the vision model may spend more than fifteen minutes processing a single full-size image. A longer prompt also adds processing time. Set `--request-timeout 7200` if a request needs a longer wait. This increases the allowed time; it does not speed up the model. A file can involve several requests for rotations or detail crops.

Try `--limit 1` first and avoid running the evaluator and another OCR job simultaneously against the same CPU model. Reducing `--max-side` can speed up a run but may reduce small-text and stamp accuracy. Failed requests are not cached; completed pages remain cached.

## Save corrections between runs

Create a UTF-8 `corrections.csv` in the project directory and use the record ID from **معرّف السجل** in the workbook.

```csv
record_id,field,value
REPLACE_WITH_RECORD_ID,decision_number,0042
```

The `field` column accepts a code key such as `decision_number` or the corresponding Arabic header, such as `رقم المقرر`. An empty value clears the field. Use the CSV to carry corrections into future exports.

## Local data and Git

Keep scans in `decisions/`. The script creates `.cache/` for OCR results, `debug_pages/` for page previews, and `exports/` for workbooks. The provided `.gitignore` excludes these directories and the default local configuration files. If you use custom input or output directories, keep them outside the Git checkout or add their paths to `.gitignore` before staging files.
