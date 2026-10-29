# PPT Name Card Batch Generator

A **local web tool** that runs in your Windows browser: upload an existing PowerPoint (`.pptx`) template → import a name list → set repeat counts and fill order → the tool automatically fills names into the matching text boxes following the template's formatting → generates a new `.pptx` file you can open with WPS or Microsoft PowerPoint.

**Fully local — no data is uploaded to the cloud, and macros in uploaded files are never executed.**

---

## Features

- Upload and analyze a `.pptx` template: shows file name, slide count, page size, and each text box's position / size / font / font size / alignment.
- Pick the slide to use as the standard card template, and **manually tick** the text boxes to fill with names. The tool only *suggests* "likely name boxes" — it never overwrites anything without your confirmation.
- Name import: paste in the browser, upload a TXT file (auto-detects UTF-8 / GBK), or upload XLSX / XLSM (choose the worksheet and the name column).
- Generation rules: repeat count per name, fill order (top-to-bottom / left-to-right), names per page, keep blank cards on the last page, generation range, trim whitespace.
- **Folded name card**: supports "fill the same name into all selected text boxes" — for cards that fold in half and show the same name on both sides.
- Preserves template formatting: page size, background, logo, images, borders, lines, shapes, colors, fonts, font sizes, alignment, positions and dimensions are kept as much as possible; only the name text boxes you selected are changed.
- Generates a **new** file (never overwrites the original), auto-paginates, downloadable when done.
- Before generating, it computes and previews the total name count and the expected slide count.

---

## Requirements

- Windows 10 / 11
- Python 3.9+ (3.10+ recommended)
- No internet connection required to run

Verified on **Python 3.14 + Flask 3.1 + python-pptx 1.0.2 + openpyxl 3.1**.

---

## Installation

1. Make sure Python is installed (run `python --version` in a Command Prompt).
2. Enter the project directory:

   ```bat
   cd C:\Users\联想\Desktop\ppt-name-generator
   ```

3. (Recommended) Create and activate a virtual environment:

   ```bat
   python -m venv .venv
   .venv\Scripts\activate
   ```

4. Install dependencies:

   ```bat
   python -m pip install -r requirements.txt
   ```

   > For a quick start you can also just run `pip install Flask python-pptx openpyxl`.

---

## Run the web app

In the project directory, run:

```bat
python app.py
```

You should see:

```
Running on http://127.0.0.1:5000
```

Open **http://127.0.0.1:5000** in your browser.

> The tool listens on `127.0.0.1` (local only) — no external device can reach it, and nothing is sent to the internet.

---

## Usage (six steps)

1. **Upload template** — click or drag a `.pptx` file into the drop zone. After parsing, pick the slide to use as the standard card template.
2. **Select name boxes** — tick the text boxes to fill with names. The tool pre-selects "likely name boxes"; **confirm them manually**.
3. **Import names** — paste, upload TXT, or upload Excel (for Excel, choose the worksheet and the name column).
4. **Set rules** — repeat count, fill order, names per page, folded-card option, etc.
5. **Preview** — review the total name count and the estimated page count.
6. **Generate & download** — click "Generate PPT", then download the new file.

Default output file name: `原文件名_台卡生成结果_日期时间.pptx` (the original file is never modified).

---

## Testing with your own template

1. Start the tool and upload your `.pptx` template in step 1.
2. In "choose template slide", select the slide **that contains the sample name** (the name box is usually pre-selected as a "likely name box").
3. In step 2, confirm the ticked boxes are exactly the **name** boxes; do **not** tick logos, fixed text, or decorative borders.
4. If your card is "one page folded, same name on both sides", enable the **folded card** option in step 4.
5. Paste a few test names, set the repeat count, generate, then open the downloaded file in WPS / PowerPoint to check the positions, fonts, logo and borders.

---

## Tested features

- Template structure analysis (text box count, position, size, font, size, alignment).
- Repeat count, fill order (t2b / l2r), names per page, name range, trim whitespace.
- Filling empty name boxes.
- "Keep blank" and "keep original text" modes for unused slots on the last page.
- Folded name card (same name filled into multiple text boxes).
- Multi-page generation and auto-pagination; the output is re-readable, with name count / page count / images / page size verified consistent.
- Error hints for empty name list, no name box selected, invalid page number, etc.
- End-to-end API flow: upload → analyze → import → generate → download.

> Verified against a real template `10.23台卡（大1）.pptx` (26 MB, with a logo image and folded name card): names filled correctly, logo image fully preserved, page size unchanged.

---

## Known limitations

- **SmartArt, charts, OLE objects, media, complex grouped shapes**: python-pptx copies them as a whole but cannot guarantee they stay editable or lossless; a risk notice is shown on upload — check the result in WPS / PowerPoint afterwards.
- **Text inside tables**: not replaced, copied as-is.
- **Fonts**: depends on fonts installed on your machine; a missing font may fall back to the default.
- **Very long names**: the tool does not shrink the font to fit (to avoid breaking the formatting) — adjust the name length or the template yourself.
- Special animations / transitions / complex slide-master structures: copied as much as possible but not guaranteed lossless.
- XLSM macros: only data is read, **macros are not executed**; the original file is not modified.

---

## Project structure

```
ppt-name-generator/
├── app.py                  # Flask backend (upload, analyze, import, generate, download)
├── ppt_processor.py        # Core: template analysis, slide copy, text replacement, pagination
├── requirements.txt        # Python dependencies
├── conftest.py             # pytest path config
├── README.md               # Documentation (Chinese)
├── README_EN.md            # Documentation (English)
├── LICENSE                 # Apache License 2.0
├── .gitignore              # Excludes venv, uploads, generated files, personal data
├── templates/
│   └── index.html          # Web UI
├── static/
│   ├── style.css           # Styles
│   └── script.js           # Frontend logic
└── tests/
    ├── test_ppt_processor.py  # Core logic tests
    └── test_app.py            # End-to-end API tests
```

Run the tests:

```bat
python -m pytest tests/ -v
```

---

## Security & privacy

- All files are processed locally; nothing is uploaded to any server.
- Macros / scripts in uploaded files are never executed.
- Uploads and generated files are stored under `temp/` and cleaned up automatically (files older than 24 hours).
- `.gitignore` already excludes `temp/`, `*.pptx`, `*.xlsx`, `*.xlsm`, etc. — **do not manually commit files containing personal names to GitHub**.

---

## License

[Apache License 2.0](LICENSE)
