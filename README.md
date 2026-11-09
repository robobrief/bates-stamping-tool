# PDF Bates Stamper

A Python utility for batch Bates stamping of PDF documents. This tool allows you to apply sequential Bates numbering to individual PDF files or entire directories of PDFs (including subdirectories).

## Features

- **Recursive Directory Processing**: Automatically processes all PDFs in a directory and its subdirectories.
- **Customizable Bates Numbering**: Define your own prefix and starting number.
- **Page Label Numbering**: Optionally stamp using each PDF page's label instead of a global sequential counter.
- **Removable Stamp Mode**: Add Bates values as PDF annotations so they can be removed later in a PDF editor.
- **Annotation Removal Mode**: Remove BatesStamper annotations from PDFs.
- **Structure Preservation**: Preserves bookmarks/outlines, page labels, and other document structure during stamping when output flattening is not enabled.
- **Flexible Positioning**: Position stamps at various locations on the page.
- **Color Options**: Customize the color of your Bates stamps.
- **Color Presets**: Use named color combinations for common workflows.
- **Background Box Option**: Optionally place stamp text in a filled box to obscure underlying content.
- **PDF Flattening**: Both input and output flattening options for maximum compatibility with PDF viewers like Adobe Reader.
- **Optional Compression**: Compress output PDFs with selectable Ghostscript profiles.
- **Robust Error Handling**: Continues processing even if individual files fail.
- **Detailed Reporting**: Provides comprehensive summary of processed files.

## Requirements

- Python 3.6+
- Required Python packages:
  - PyPDF2
  - pypdf
  - reportlab
  - python-docx
  - pymupdf
  - pikepdf

Optional but recommended:
- Ghostscript (for better PDF flattening)

## Installation

1. Clone or download this repository
2. Install the project dependencies:

```bash
python -m pip install -r requirements.txt
```

If you prefer to install just the basic stamper packages:

```bash
python -m pip install PyPDF2 reportlab
```

3. (Optional) Install Ghostscript for better PDF flattening:

**On Ubuntu/Debian (or WSL):**
```bash
sudo apt-get update
sudo apt-get install ghostscript
```

**On Windows:**
Download and install from [Ghostscript official site](https://www.ghostscript.com/releases/gsdnld.html)

## Usage

## Authorities Bundle Workflow (Word index + ordered merge)

This repository also includes [bundle_workflow.py](/C:/github/bates-stamping-tool/bundle_workflow.py), a local-only workflow for:

- reading the authoritative row order from a Word index (`.docx`),
- inventorying and matching source PDFs,
- producing an editable manifest and review reports before merge,
- merging only approved entries in order,
- creating one top-level bookmark per entry as `Citation:: Case title`,
- generating pagination and validation reports.

### Additional dependency

Required for Word index reading:

```bash
pip install python-docx
```

### Two-step run

1) Prepare (manifest + review outputs; no merge):

```bash
python bundle_workflow.py prepare --source-folder "C:\path\to\pdfs" --word-index "C:\path\to\index.docx" --output-folder "C:\path\to\output" --output-pdf-name "Authorities Bundle.pdf"
```

2) After reviewing/editing the manifest, merge approved rows:

```bash
python bundle_workflow.py merge --source-folder "C:\path\to\pdfs" --output-folder "C:\path\to\output" --output-pdf-name "Authorities Bundle.pdf" --manifest "C:\path\to\output\authorities_bundle_manifest.csv" --include-index-pdf ask
```

3) Populate the Word index `Page` column from merged pagination (writes to a copy):

```bash
python bundle_workflow.py populate-index-pages --word-index "C:\path\to\index.docx" --pagination-report "C:\path\to\output\authorities_bundle_pagination_report.csv" --output-word-index "C:\path\to\output\Authorities Index (with pages).docx"
```

### GUI workflow (reusable local app)

Use [bundle_gui.py](/C:/github/bates-stamping-tool/bundle_gui.py) for an end-to-end local GUI (**LawPage PDF Bundling Tool v1.0.0**) that supports:

- selecting source folder, index DOCX, output folder,
- loading/reordering PDFs (up/down/top/bottom),
- merging in chosen order with `Citation:: Case title` bookmarks,
- pagination report generation,
- updating the index DOCX Page column,
- optional Word-exact index PDF insertion at page 1,
- option to skip index insertion entirely and output a bundle with no front index PDF,
- optional full-row index hyperlinks,
- optional compression and open-with-bookmarks mode.
- optional one-click stamping (including page-label numbering, no-prefix mode, presets like reverse-blue).
- save/load full reusable profiles (paths, PDF order, build options, stamp options).
- one-click **Load Preferred Defaults** button for your standard setup (no prefix, page-label numbering, reverse-blue, bottom-right, linking/compression/outlines enabled).
- combined **Inputs & PDF Order** tab.
- added **Help / Instructions** tab with a recommended run sequence and mode guidance.
- separate **Index Settings**, **Page Label Settings**, and **Output** tabs.
- separate preferred-default actions: build-only defaults, stamp-only defaults, or all defaults.
- stamp mode guidance in-app (`annotation` overlay vs `merge` content merge).
- page-label navigation bookmarks use the actual PDF page labels; optional document-bookmark grouping is configured separately.
- optional grouping of existing top-level document bookmarks under a parent folder (default `NAVIGATE BY DOCUMENT DESCRIPTION`).
- template-driven index generation with a 2-column table (`Description`, `Page Label`) using thin borders and banded row coloring.
- index table titles are rendered using Word's built-in **Caption** paragraph style.
- optional seed table placeholder support: a table captioned `[[INDEX_TABLE_PLACEHOLDER]]` is removed before index tables are generated (useful for carrying table styles in templates).
- description source option: use PDF filenames (default) or use pre-written descriptions from the index DOCX and fill page numbers only.
- optional page display mode: start page only, or start-end page range.
- optional leading **#** column (bundle document index: 1, 2, 3, ...).
- optional extra column: **PDF Page** (internal PDF page number, 1-based).
- configurable column labels for **Page Label** and **PDF Page**.
- configurable page-column width (cm), native Word table style controls (Table style family: Plain/Grid; style picker from detected built-ins; Header/Banded/First/Last flags), and optional page break before each additional index table.
- generated index tables always set the header row to repeat on subsequent pages.
- index tables: mandatory **Index (by page label)** plus optional **Index (by document description)**.
- index preview pane with row counts per selected section before running.
- style-preview button to generate a sample DOCX table using the current table/page settings.

Install dependencies:

```bash
pip install python-docx pypdf pymupdf pikepdf
```

Run:

```bash
python bundle_gui.py
```

### Basic Usage

```bash
python bates_stamper.py input_directory output_directory
```

This will process all PDFs in `input_directory` (including subdirectories), add default Bates numbers (BATES-000001, BATES-000002, etc.), and save the results in `output_directory` with the same directory structure.

### Advanced Options

```bash
python bates_stamper.py input_directory output_directory \
  --prefix "CASE-" \
  --start 1000 \
  --stamp-mode annotation \
  --number-source page-label \
  --position top-right \
  --color "#FF0000" \
  --background-box \
  --box-color "#FFFFFF" \
  --box-padding 2 \
  --margin 20 \
  --size 14 \
  --flatten-input \
  --flatten-output \
  --compress-output ebook
```

### Command-Line Arguments

| Argument | Description | Default |
|----------|-------------|---------|
| `input_directory` | Directory containing PDF files to stamp | (Required) |
| `output_directory` | Directory where stamped PDFs will be saved | (Required) |
| `--operation` | `stamp` (default) or `remove-annotations` | stamp |
| `--annotation-removal` | With remove mode: `bates` or `all-freetext` | bates |
| `--prefix` | Prefix for Bates numbers | BATES- |
| `--start` | Starting Bates number | 1 |
| `--stamp-mode` | Stamp mode: `merge` (permanent) or `annotation` (removable) | merge |
| `--number-source` | Number source: `sequential` or `page-label` | sequential |
| `--position` | Position of the Bates stamp | bottom-right |
| `--color-preset` | Preset style name (see table below) | (None) |
| `--color` | Hex color code for the Bates stamp | #000000 (black) |
| `--background-box` | Draw a filled box behind stamp text | (Disabled) |
| `--box-color` | Hex color code for the stamp background box | #FFFFFF (white) |
| `--box-padding` | Padding around text inside background box (points) | 2 |
| `--margin` | Margin from edge in points (1/72 inch) | 10 |
| `--size` | Font size for the Bates stamp | 12 |
| `--flatten-input` | Flatten input PDFs before stamping | (Disabled) |
| `--flatten-output` | Flatten output PDFs after stamping | (Disabled) |
| `--compress-output` | Compress output via Ghostscript: `none`, `screen`, `ebook`, `printer`, `prepress`, `default` | none |
| `--set-page-labels-from-stamp` | In sequential mode, set PDF page labels from stamped Bates values | (Disabled) |
| `--bookmark-page-labels` | Add one bookmark per page label under a parent folder | (Disabled) |
| `--page-label-bookmark-parent` | Parent bookmark title for page-label bookmarks | NAVIGATE BY PAGE LABEL |

Notes:
- With `--number-source page-label`, the `--start` value does not affect the stamped text.
- With `--stamp-mode annotation`, stamps are removable as PDF annotations in compatible PDF editors.
- With `--operation remove-annotations --annotation-removal bates`, only annotations created by this script are removed.
- `bates` removal also recognizes older script-created annotation files generated before the current marker was added.
- Use `--annotation-removal all-freetext` only when you explicitly want to remove every FreeText annotation.
- When using `--flatten-input`, page labels are still read from the original input PDF before flattening.
- If `--flatten-output` is used, PDF viewers may lose or alter bookmarks/outlines/page labels because flattening rewrites the PDF structure.
- If `--compress-output` is used, Ghostscript rewrites the PDF and may alter structure metadata such as page labels/bookmarks depending on the source PDF.
- If `--stamp-mode annotation` is used, flattening/compression is automatically skipped so stamps remain removable.
- Background box width automatically scales with the stamp text length, and text is centered in the box.
- If `--color-preset` is provided, preset text/box colors are used.
- If `--set-page-labels-from-stamp` is enabled with sequential numbering, output PDF page labels are set to the visible Bates values.
- If `--bookmark-page-labels` is enabled, one bookmark per page is added under the configured parent folder title.

#### Color Presets

| Preset | Text | Background | Use Case |
|--------|------|------------|----------|
| `black` | Black | Transparent | Most common |
| `white` | White | Transparent | Dark PDFs |
| `red` | Red | Transparent | Privileged / review |
| `black-box` | Black | White | Maximum readability |
| `white-box` | White | Black | Dark drawings |
| `red-box` | Red | White | Review copies |
| `reverse-red` | White | Red | Very visible |
| `blue-box` | Blue | White | Production sets |
| `reverse-blue` | White | Blue | Internal numbering |
| `highlight` | Black | Yellow | High visibility |

#### Available Positions
- `bottom-right`
- `bottom-left`
- `bottom-middle`
- `top-right`
- `top-left`
- `top-middle`
- `center`

## Examples

### Example 1: Basic Bates Numbering

```bash
python bates_stamper.py ./documents ./stamped_documents
```

### Example 2: Removable Stamps (Annotation Mode)

```bash
python bates_stamper.py ./documents ./stamped_documents --stamp-mode annotation --background-box --box-color "#000000" --color "#FFFFFF"
```

### Example 3: Remove BatesStamper Annotations

```bash
python bates_stamper.py ./documents ./cleaned_documents --operation remove-annotations --annotation-removal bates
```

### Example 4: Custom Prefix and Starting Number

```bash
python bates_stamper.py ./evidence ./stamped_evidence --prefix "CASE123-" --start 5000
```

This will create Bates numbers like CASE123-005000, CASE123-005001, etc.

### Example 5: Red Stamps in Top-Left Corner

```bash
python bates_stamper.py ./contracts ./stamped_contracts --position top-left --color "#FF0000" --size 14
```

### Example 6: White Background Box Behind Stamp Text

```bash
python bates_stamper.py ./documents ./stamped_documents --background-box --box-color "#FFFFFF" --box-padding 3
```

### Example 7: Use PDF Page Labels for Stamp Numbers

```bash
python bates_stamper.py ./documents ./stamped_documents --prefix "DOC-" --number-source page-label
```

This uses the source PDF page labels (for example `i`, `ii`, `1`, `2`, `A-1`) as the number portion of each stamp.

### Example 8: Handling Problematic PDFs with Flattening

If you encounter issues with blank pages in Adobe Reader:

```bash
python bates_stamper.py ./problematic_pdfs ./fixed_pdfs --flatten-input
```

### Example 9: Compressed Output (Smaller File Size)

```bash
python bates_stamper.py ./documents ./compressed_docs --compress-output ebook
```

### Example 10: Maximum Compatibility

For maximum compatibility with all PDF viewers:

```bash
python bates_stamper.py ./documents ./compatible_docs --flatten-input --flatten-output
```

### Example 11: Sequential Numbering That Also Sets Page Labels

```bash
python bates_stamper.py ./documents ./stamped_documents --number-source sequential --prefix "A" --set-page-labels-from-stamp
```

### Example 12: Bookmark Every Page by Label

```bash
python bates_stamper.py ./documents ./stamped_documents --bookmark-page-labels --page-label-bookmark-parent "NAVIGATE BY PAGE LABEL"
```

## Troubleshooting

### Blank Pages in Adobe Reader

If the stamped PDFs appear blank in Adobe Reader:
1. Try using the `--flatten-input` option
2. If that doesn't work, try both `--flatten-input` and `--flatten-output`
3. Ensure Ghostscript is installed for better flattening results

### Incorrect Stamp Positioning

If stamps are not positioned correctly:
1. Try adjusting the `--margin` value
2. Use a different `--position` option

### Processing Errors

If specific files fail to process:
1. Check if the original PDF is valid and not password-protected
2. Try using the `--flatten-input` option for complex PDFs

## How It Works

The script:
1. Recursively finds all PDF files in the input directory
2. Optionally reads PDF page labels (if `--number-source page-label` is used)
3. Optionally flattens input PDFs (if `--flatten-input` is used)
4. Creates a new PDF with the Bates stamp at the specified position
5. Merges the stamp with each page of the original PDF
6. Optionally flattens or compresses the output PDF (if `--flatten-output` or `--compress-output` is used)
7. Saves the result to the output directory, maintaining the original directory structure

## Disclaimer

I am not an attorney, and this tool is provided as-is, without any warranties or guarantees of fitness for legal use. This Bates stamping utility may not preserve all metadata or file integrity required for legal proceedings or compliance in your jurisdiction. It is your responsibility to ensure that the use of this tool complies with local laws, court rules, and professional standards. Use at your own risk.

## License

This project is open source and available under the MIT License.

Important licensing note: This script can optionally use Ghostscript, which is licensed under the AGPL (Affero General Public License). If you enable Ghostscript functionality:

This creates a dependency on AGPL-licensed software
If you distribute this script with Ghostscript integration enabled, you may need to comply with AGPL requirements
You should consult with a legal professional regarding licensing implications

To avoid AGPL implications:

Use the script without installing Ghostscript (the PyPDF2 fallback will be used instead)
Remove or disable the Ghostscript-dependent code if redistributing the software

For more information on Ghostscript licensing, visit Ghostscript Licensing

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
