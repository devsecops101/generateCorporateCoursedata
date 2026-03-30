# generateCorporateCoursedata (code)

## Setup (venv)

From the `code/` folder:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

## Run

Default input is `data/coursedata.xlsx` (repo root) or `code/data/coursedata.xlsx` (fallback):

```bash
source .venv/bin/activate
python read_courses_excel.py
```

Generate JSON for the first N rows (starting from row 1):

```bash
python read_courses_excel.py --rows 10
```

By default, JSON files are written to the repo `data/` folder, with a fallback to `code/data/` if that’s the only one that exists. You can override:

```bash
python read_courses_excel.py --rows 10 --out-dir ./data
```

## Using Claude (optional)

Set your key (recommended in your shell profile or a local env file that is **not** committed):

```bash
export ANTHROPIC_API_KEY="your_key_here"
```

If you saved it in `code/.venv/.env`, the script will also auto-load it when you run with `--use-claude`.

Then run with Claude generation enabled:

```bash
python read_courses_excel.py --rows 1 --use-claude
```

You can also choose the model:

```bash
python read_courses_excel.py --rows 1 --use-claude --claude-model "auto"
```

Or pass an explicit Excel path:

```bash
python read_courses_excel.py /path/to/file.xlsx
```

If you need a specific sheet:

```bash
python read_courses_excel.py /path/to/file.xlsx --sheet 0
python read_courses_excel.py /path/to/file.xlsx --sheet "Sheet1"
```

## What it loads into memory

The script reads these columns (case/whitespace-insensitive):

- `Course Title`
- `Description`
- `category`

And returns a Python list in memory (`list[CourseRow]`), where each item has:

- `course_title`
- `description`
- `category`

