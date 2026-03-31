from __future__ import annotations

import json
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import streamlit as st


APP_DIR = Path(__file__).resolve().parent
DEFAULT_JSON_DIR = APP_DIR / "json"


@dataclass
class ValidationIssue:
    level: str  # "error" | "warning"
    message: str


def _load_json(path: Path) -> dict[str, Any]:
    raw = path.read_text(encoding="utf-8")
    data = json.loads(raw)
    if not isinstance(data, dict):
        raise ValueError("Top-level JSON must be an object")
    return data


def _atomic_write_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    tmp.replace(path)


def _backup_file(path: Path) -> Path:
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    bak = path.with_suffix(path.suffix + f".{ts}.bak")
    shutil.copy2(path, bak)
    return bak


def _as_str(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, (int, float, bool)):
        return str(v)
    return str(v)


def _coerce_int(v: Any, *, default: int) -> int:
    if isinstance(v, int):
        return v
    if isinstance(v, float) and v == int(v):
        return int(v)
    if isinstance(v, str) and v.strip().isdigit():
        return int(v.strip())
    return default


def _validate_course(data: dict[str, Any]) -> list[ValidationIssue]:
    issues: list[ValidationIssue] = []

    required_top = [
        "Course Title",
        "Category",
        "Description",
        "Duration",
        "Difficulty",
        "Price",
        "Instructor",
        "Published Status",
        "Prerequisites",
        "Learning objectives",
        "lessons",
    ]
    for k in required_top:
        if k not in data:
            issues.append(ValidationIssue("error", f'Missing required top-level key: "{k}"'))

    lessons = data.get("lessons")
    if lessons is None:
        return issues
    if not isinstance(lessons, list):
        issues.append(ValidationIssue("error", '"lessons" must be a list'))
        return issues
    if not lessons:
        issues.append(ValidationIssue("warning", '"lessons" is empty'))
        return issues

    for idx, lesson in enumerate(lessons, start=1):
        if not isinstance(lesson, dict):
            issues.append(ValidationIssue("error", f"Lesson {idx} must be an object"))
            continue
        for lk in ["lesson title", "order", "Description", "status", "content", "slides", "scripts"]:
            if lk not in lesson:
                issues.append(ValidationIssue("warning", f'Lesson {idx} missing key: "{lk}"'))

        slides = lesson.get("slides")
        if isinstance(slides, list) and slides:
            nums = []
            for s in slides:
                if not isinstance(s, dict):
                    issues.append(ValidationIssue("error", f"Lesson {idx} slide entry must be an object"))
                    continue
                n = s.get("number")
                if not isinstance(n, int):
                    issues.append(ValidationIssue("warning", f"Lesson {idx} slide has non-int number: {n!r}"))
                else:
                    nums.append(n)
                for sk in ["number", "slidetitle", "subtitle", "slidecontent"]:
                    if sk not in s:
                        issues.append(ValidationIssue("warning", f'Lesson {idx} slide missing key: "{sk}"'))
            if nums and len(set(nums)) != len(nums):
                issues.append(ValidationIssue("error", f"Lesson {idx} slides have duplicate numbers"))

        scripts = lesson.get("scripts")
        if isinstance(scripts, list) and scripts:
            nums = []
            for sc in scripts:
                if not isinstance(sc, dict):
                    issues.append(ValidationIssue("error", f"Lesson {idx} script entry must be an object"))
                    continue
                n = sc.get("number")
                if not isinstance(n, int):
                    issues.append(ValidationIssue("warning", f"Lesson {idx} script has non-int number: {n!r}"))
                else:
                    nums.append(n)
                for sk in ["number", "script"]:
                    if sk not in sc:
                        issues.append(ValidationIssue("warning", f'Lesson {idx} script missing key: "{sk}"'))
            if nums and len(set(nums)) != len(nums):
                issues.append(ValidationIssue("error", f"Lesson {idx} scripts have duplicate numbers"))

    return issues


def _sorted_json_files(json_dir: Path) -> list[Path]:
    if not json_dir.exists():
        return []
    return sorted([p for p in json_dir.iterdir() if p.is_file() and p.suffix.lower() == ".json"])


def _get_by_number(items: list[dict[str, Any]], number: int) -> dict[str, Any] | None:
    for it in items:
        if isinstance(it, dict) and it.get("number") == number:
            return it
    return None


def _ensure_numbered_items(
    items: Any,
    *,
    count: int,
    template: dict[str, Any],
) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    if isinstance(items, list):
        for x in items:
            if isinstance(x, dict):
                out.append(dict(x))
    # ensure 1..count exist
    for n in range(1, count + 1):
        existing = _get_by_number(out, n)
        if existing is None:
            created = dict(template)
            created["number"] = n
            out.append(created)
    out.sort(key=lambda d: _coerce_int(d.get("number"), default=10**9))
    return out


st.set_page_config(page_title="JSON Course Reviewer", layout="wide")

st.title("JSON Course Reviewer")
st.caption("Open, edit, validate, and save course JSON files.")

with st.sidebar:
    st.subheader("Files")
    json_dir_str = st.text_input("JSON folder", value=str(DEFAULT_JSON_DIR))
    json_dir = Path(json_dir_str).expanduser()

    files = _sorted_json_files(json_dir)
    if not files:
        st.warning("No JSON files found in that folder.")
        st.stop()

    file_labels = [p.name for p in files]
    if "file_index" not in st.session_state:
        st.session_state["file_index"] = 0

    # Ensure the stored index is within the current bounds
    st.session_state["file_index"] = min(
        max(st.session_state["file_index"], 0), len(file_labels) - 1
    )

    nav_prev, nav_info, nav_next = st.columns([1, 2, 1])
    with nav_prev:
        if st.button("◀ Prev", use_container_width=True, disabled=st.session_state["file_index"] == 0):
            st.session_state["file_index"] -= 1
    with nav_next:
        if st.button(
            "Next ▶",
            use_container_width=True,
            disabled=st.session_state["file_index"] >= len(file_labels) - 1,
        ):
            st.session_state["file_index"] += 1
    with nav_info:
        st.caption(
            f"File {st.session_state['file_index'] + 1} of {len(file_labels)}"
        )

    selected_label = st.selectbox(
        "Select a JSON file",
        options=file_labels,
        index=st.session_state["file_index"],
    )
    # Keep the index in sync if the user picks from the dropdown directly
    st.session_state["file_index"] = file_labels.index(selected_label)
    selected_path = next(p for p in files if p.name == selected_label)

    reload_clicked = st.button("Reload from disk")


def _load_into_state(path: Path) -> None:
    data = _load_json(path)
    st.session_state["current_path"] = str(path)
    st.session_state["original_data"] = data
    st.session_state["working_data"] = json.loads(json.dumps(data))  # deep copy (json-safe)
    st.session_state["last_loaded_mtime"] = path.stat().st_mtime


need_load = (
    reload_clicked
    or "current_path" not in st.session_state
    or st.session_state.get("current_path") != str(selected_path)
)
if need_load:
    _load_into_state(selected_path)

path = Path(st.session_state["current_path"])
data: dict[str, Any] = st.session_state["working_data"]

top_left, top_right = st.columns([2, 1], gap="large")

with top_right:
    st.subheader("Validation")
    issues = _validate_course(data)
    errors = [i for i in issues if i.level == "error"]
    warnings = [i for i in issues if i.level == "warning"]
    if not issues:
        st.success("Looks good.")
    else:
        if errors:
            st.error(f"{len(errors)} error(s)")
        if warnings:
            st.warning(f"{len(warnings)} warning(s)")
        with st.expander("Details", expanded=bool(errors)):
            for i in issues:
                (st.error if i.level == "error" else st.warning)(i.message)

    st.subheader("Save")
    make_backup = st.checkbox("Create .bak backup on save", value=True)
    can_save = len(errors) == 0
    save_clicked = st.button("Save to disk", type="primary", disabled=not can_save)
    if not can_save:
        st.caption("Fix validation errors to enable saving.")

    if save_clicked:
        try:
            if make_backup and path.exists():
                bak = _backup_file(path)
                st.info(f"Backup created: {bak.name}")
            _atomic_write_json(path, data)
            st.success("Saved.")
        except Exception as e:
            st.error(f"Save failed: {e!s}")

    st.subheader("Raw JSON")
    st.download_button(
        label="Download JSON",
        file_name=path.name,
        data=json.dumps(data, indent=2, ensure_ascii=False) + "\n",
        mime="application/json",
    )


with top_left:
    st.subheader(f"Editing: {path.name}")

    tabs = st.tabs(["Course", "Lessons", "Slides", "Scripts", "Preview"])

    with tabs[0]:
        c1, c2 = st.columns(2, gap="large")
        with c1:
            data["Course Title"] = st.text_input("Course Title", value=_as_str(data.get("Course Title", "")))
            data["Category"] = st.text_input("Category", value=_as_str(data.get("Category", "")))
            data["Description"] = st.text_area("Description", value=_as_str(data.get("Description", "")), height=120)
            data["Prerequisites"] = st.text_area(
                "Prerequisites", value=_as_str(data.get("Prerequisites", "")), height=220
            )
            data["Learning objectives"] = st.text_area(
                "Learning objectives", value=_as_str(data.get("Learning objectives", "")), height=260
            )

        with c2:
            data["Duration"] = st.text_input("Duration", value=_as_str(data.get("Duration", "")))
            data["Difficulty"] = st.text_input("Difficulty", value=_as_str(data.get("Difficulty", "")))
            data["Price"] = st.text_input("Price", value=_as_str(data.get("Price", "")))
            data["Instructor"] = st.text_input("Instructor", value=_as_str(data.get("Instructor", "")))
            data["Published Status"] = st.text_input(
                "Published Status", value=_as_str(data.get("Published Status", ""))
            )

    lessons = data.get("lessons")
    if not isinstance(lessons, list):
        lessons = []
        data["lessons"] = lessons
    if not lessons:
        lessons.append(
            {
                "lesson title": _as_str(data.get("Course Title", "")),
                "order": 1,
                "Description": "",
                "status": "draft",
                "content": _as_str(data.get("Description", "")),
                "slides": [],
                "scripts": [],
            }
        )

    lesson0 = lessons[0] if isinstance(lessons[0], dict) else {}
    if not isinstance(lesson0, dict):
        lesson0 = {}
        lessons[0] = lesson0

    with tabs[1]:
        lesson0["lesson title"] = st.text_input("Lesson title", value=_as_str(lesson0.get("lesson title", "")))
        lesson0["order"] = _coerce_int(st.text_input("Order", value=_as_str(lesson0.get("order", 1))), default=1)
        lesson0["Description"] = st.text_area(
        "Lesson description", value=_as_str(lesson0.get("Description", "")), height=240
        )
        lesson0["status"] = st.text_input("Lesson status", value=_as_str(lesson0.get("status", "draft")))
        lesson0["content"] = st.text_area("Lesson content", value=_as_str(lesson0.get("content", "")), height=100)

    lesson0_slides = _ensure_numbered_items(
        lesson0.get("slides"),
        count=8,
        template={"slidetitle": "", "subtitle": "", "slidecontent": ""},
    )
    lesson0["slides"] = lesson0_slides

    with tabs[2]:
        st.caption("Edit slide title, subtitle, and bullet content.")
        for slide in lesson0_slides:
            n = _coerce_int(slide.get("number"), default=0)
            with st.expander(f"Slide {n}", expanded=True):
                slide["slidetitle"] = st.text_input(
                    f"Slide {n} title",
                    value=_as_str(slide.get("slidetitle", "")),
                    key=f"slide_title_{n}_{path.name}",
                )
                slide["subtitle"] = st.text_input(
                    f"Slide {n} subtitle",
                    value=_as_str(slide.get("subtitle", "")),
                    key=f"slide_subtitle_{n}_{path.name}",
                )
                slide["slidecontent"] = st.text_area(
                    f"Slide {n} content (bullets)",
                    value=_as_str(slide.get("slidecontent", "")),
                    height=160,
                    key=f"slide_content_{n}_{path.name}",
                )

    lesson0_scripts = _ensure_numbered_items(
        lesson0.get("scripts"),
        count=8,
        template={"script": ""},
    )
    lesson0["scripts"] = lesson0_scripts

    with tabs[3]:
        st.caption("Edit narration scripts for each slide.")
        for sc in lesson0_scripts:
            n = _coerce_int(sc.get("number"), default=0)
            with st.expander(f"Script {n}", expanded=True):
                sc["script"] = st.text_area(
                    f"Slide {n} script",
                    value=_as_str(sc.get("script", "")),
                    height=220,
                    key=f"script_{n}_{path.name}",
                )

    with tabs[4]:
        st.json(data, expanded=False)

