# JSON Course Reviewer (GUI)

This folder contains a small local GUI app for reviewing and editing the course JSON files in `review/json/`.

## Setup

From the repo root:

```bash
python3 -m venv review/.venv
source review/.venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r review/requirements.txt
```

## Run

```bash
source review/.venv/bin/activate
streamlit run review/app.py
```

The app will:
- Load JSON files from `review/json/` by default
- Let you edit course fields, lesson fields, slides, and scripts
- Validate the structure before saving
- Save changes back to the same JSON file (atomic write) and create a `.bak` backup

