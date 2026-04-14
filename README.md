# pypptx

A Python CLI toolkit for manipulating `.pptx` files at the OPC/XML level.

## Installation

```
pip install -e .
```

Or with [uv](https://github.com/astral-sh/uv):

```
uv pip install -e .
```

### Optional: thumbnails support

The `thumbnails` command requires [Pillow](https://pillow.readthedocs.io/) and two system tools.
Install the optional extra to get Pillow:

```
pip install 'pypptx[thumbnails]'
```

Then install the system tools:

| Tool | macOS | Debian/Ubuntu |
|---|---|---|
| LibreOffice (`soffice`) | `brew install --cask libreoffice` | `sudo apt-get install libreoffice` |
| Poppler (`pdftoppm`) | `brew install poppler` | `sudo apt-get install poppler-utils` |

## Core editing workflow

The recommended workflow for making structural edits is:

```
pypptx unpack presentation.pptx          # 1. Unpack to a directory
# ... edit XML files directly ...        # 2. Edit XML as needed
pypptx clean presentation/               # 3. Remove orphaned files
pypptx pack presentation/ output.pptx   # 4. Repack into a .pptx file
```

Most commands also accept a `.pptx` file directly and handle unpack/clean/repack
internally, so the explicit workflow above is only needed when editing XML by hand.

## Output contract

- **Default output**: every command writes a single JSON object to stdout.
- **`--plain` flag**: pass `--plain` to receive human-readable text instead.
- **Errors**: all error messages are written to stderr, never stdout.
- **Exit codes**: `0` on success, `1` on any error.

## Commands

Run `pypptx --help` or `pypptx <command> --help` for the full option reference.

---

### `extract-text`

Extract all text from a `.pptx` file. Without `--output`, raw text goes directly to
stdout (no JSON wrapper). With `--output`, text is written to the given file and
command metadata is emitted as JSON.

```
pypptx extract-text presentation.pptx
```

```
--- Slide 1 ---
Welcome to pypptx
A PowerPoint manipulation toolkit
--- Slide 2 ---
Installation
pip install -e .
```

With `--output`:

```
pypptx extract-text presentation.pptx --output extracted.txt
```

```json
{"output_file": "extracted.txt", "slide_count": 2}
```

Limit to specific slides with `--slides`:

```
pypptx extract-text presentation.pptx --slides 1,3
```

With `--plain` and `--output`:

```
pypptx extract-text presentation.pptx --output extracted.txt --plain
```

```
extracted.txt
```

---

### `unpack`

Unpack a `.pptx` file into a directory of raw OPC parts (XML, images, etc.).
If no output directory is given, defaults to the file's stem name.

```
pypptx unpack presentation.pptx
```

```json
{"unpacked_dir": "presentation"}
```

Specify an explicit output directory:

```
pypptx unpack presentation.pptx my-edits/
```

```json
{"unpacked_dir": "my-edits"}
```

With `--plain`:

```
pypptx unpack presentation.pptx --plain
```

```
presentation
```

---

### `pack`

Repack an unpacked directory back into a `.pptx` file.
`[Content_Types].xml` is written first per the OPC spec; writes atomically.

```
pypptx pack presentation/ output.pptx
```

```json
{"output_file": "output.pptx"}
```

With `--plain`:

```
pypptx pack presentation/ output.pptx --plain
```

```
output.pptx
```

---

### `clean`

Remove orphaned OPC parts from a `.pptx` file or an unpacked directory.
Orphans are files not reachable by following relationship links from the package root,
or slides that exist in relationships but are absent from `sldIdLst`.

```
pypptx clean presentation.pptx
```

```json
{"removed": ["ppt/slides/slide3.xml", "ppt/slides/_rels/slide3.xml.rels"]}
```

Returns an empty list when nothing was removed:

```json
{"removed": []}
```

With `--plain` (each removed file on its own line, empty output when nothing removed):

```
pypptx clean presentation.pptx --plain
```

```
ppt/slides/slide3.xml
ppt/slides/_rels/slide3.xml.rels
```

---

### `thumbnails`

Generate labeled thumbnail grid images from a `.pptx` file.
Requires LibreOffice, Poppler/pdftoppm, and Pillow — see [thumbnails support](#optional-thumbnails-support).

```
pypptx thumbnails presentation.pptx
```

```json
{"files": ["thumbnails.jpg"]}
```

With `--plain` (one file path per line):

```
pypptx thumbnails presentation.pptx --plain
```

```
thumbnails.jpg
```

**Options:**

| Option | Default | Description |
|---|---|---|
| `--output PREFIX` | `thumbnails` | Output filename prefix; `.jpg` is appended automatically. |
| `--cols N` | `3` (max `6`) | Number of columns in the thumbnail grid. |
| `--plain` | off | Emit one file path per line instead of JSON. |

**Multi-file output:**

When the slide count exceeds `cols × (cols + 1)`, the output is split across multiple JPEG files.
With the default `--cols 3` each grid holds up to 12 slides (3 × 4).
A deck with more than 12 slides produces multiple files suffixed `-1`, `-2`, …:

```
pypptx thumbnails big-deck.pptx
```

```json
{"files": ["thumbnails-1.jpg", "thumbnails-2.jpg"]}
```

**Hidden slides:**

Hidden slides are rendered as a hatched grey placeholder image in the grid rather than being skipped,
so the grid index always matches the presentation slide number.

---

### `slide list`

List the slides in a `.pptx` file or unpacked directory in presentation order.

```
pypptx slide list presentation.pptx
```

```json
{"slides": [
  {"index": 1, "file": "slide1.xml", "hidden": false},
  {"index": 2, "file": "slide2.xml", "hidden": true},
  {"index": 3, "file": "slide3.xml", "hidden": false}
]}
```

With `--plain`:

```
pypptx slide list presentation.pptx --plain
```

```
slide1.xml
slide2.xml [hidden]
slide3.xml
```

---

### `slide layouts`

List all slide layouts available in a `.pptx` file or unpacked directory.
Useful for finding the `--layout` index required by `slide add`.

```
pypptx slide layouts presentation.pptx
```

```json
{"layouts": [
  {"index": 1, "file": "slideLayout1.xml"},
  {"index": 2, "file": "slideLayout2.xml"},
  {"index": 3, "file": "slideLayout3.xml"}
]}
```

With `--plain`:

```
pypptx slide layouts presentation.pptx --plain
```

```
slideLayout1.xml
slideLayout2.xml
slideLayout3.xml
```

---

### `slide add`

Add a slide to a `.pptx` file or unpacked directory.
Exactly one of `--duplicate` or `--layout` must be supplied.

**Duplicate an existing slide** (notes are not copied):

```
pypptx slide add presentation.pptx --duplicate 2
```

```json
{"added_file": "slide4.xml", "position": 4}
```

**Add a blank slide** using a layout (use `slide layouts` to find the index):

```
pypptx slide add presentation.pptx --layout 1
```

```json
{"added_file": "slide4.xml", "position": 4}
```

**Insert at a specific position** with `--position`:

```
pypptx slide add presentation.pptx --duplicate 1 --position 2
```

```json
{"added_file": "slide4.xml", "position": 2}
```

With `--plain`:

```
pypptx slide add presentation.pptx --layout 1 --plain
```

```
slide4.xml at position 4
```

---

### `slide delete`

Delete a slide by its 1-based index.

```
pypptx slide delete presentation.pptx 3
```

```json
{"deleted_file": "slide3.xml", "deleted_index": 3}
```

With `--plain`:

```
pypptx slide delete presentation.pptx 3 --plain
```

```
deleted slide3.xml (index 3)
```

---

### `slide move`

Move a slide from one 1-based position to another.

```
pypptx slide move presentation.pptx 3 1
```

```json
{"file": "slide3.xml", "from": 3, "to": 1}
```

With `--plain`:

```
pypptx slide move presentation.pptx 3 1 --plain
```

```
slide3.xml: 3 -> 1
```
