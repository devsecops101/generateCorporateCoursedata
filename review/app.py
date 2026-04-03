from __future__ import annotations

import json
import os
import re
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import streamlit as st

try:
    import anthropic
except ImportError:
    anthropic = None  # type: ignore[assignment, misc]


APP_DIR = Path(__file__).resolve().parent
DEFAULT_JSON_DIR = APP_DIR / "json"
SUPPORTING_DIR = APP_DIR / "supporting_files"
HELP_TXT = SUPPORTING_DIR / "Help.txt"
SLIDES_TEMPLATE_PPTX = SUPPORTING_DIR / "AISEC.Course.Slides Template v0.5.pptx"
CAMTASIA_TEMPLATE_CMPROJ = SUPPORTING_DIR / "Untitled.cmproj"

DEFAULT_CLAUDE_MODEL = os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-4-20250514")
_CLAUDE_AUDIENCE_JSON_HINT = re.compile(
    r"```(?:json)?\s*([\s\S]*?)\s*```", re.IGNORECASE
)


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


def _build_script_export_text(data: dict[str, Any]) -> str:
    """Plain text of included scripts only (do_not_include is false)."""
    lessons = data.get("lessons")
    if not isinstance(lessons, list) or not lessons:
        return ""
    lesson0 = lessons[0]
    if not isinstance(lesson0, dict):
        return ""
    scripts = lesson0.get("scripts")
    if not isinstance(scripts, list):
        return ""
    blocks: list[str] = []
    ordered = sorted(scripts, key=lambda d: _coerce_int(d.get("number"), default=0))
    sep = "-" * 72
    slide_export_n = 0
    for sc in ordered:
        if not isinstance(sc, dict):
            continue
        if sc.get("do_not_include"):
            continue
        slide_export_n += 1
        n = _coerce_int(sc.get("number"), default=0)
        body = _as_str(sc.get("script")).rstrip()
        blocks.append(f"Script {n} (Slide {slide_export_n})\n{sep}\n{body}")
    return "\n\n".join(blocks) + ("\n" if blocks else "")


def _build_slides_export_text(data: dict[str, Any]) -> str:
    """Plain text of included slides only (do_not_include is false)."""
    lessons = data.get("lessons")
    if not isinstance(lessons, list) or not lessons:
        return ""
    lesson0 = lessons[0]
    if not isinstance(lesson0, dict):
        return ""
    slides = lesson0.get("slides")
    if not isinstance(slides, list):
        return ""
    blocks: list[str] = []
    ordered = sorted(slides, key=lambda d: _coerce_int(d.get("number"), default=0))
    sep = "-" * 72
    slide_export_n = 0
    for sl in ordered:
        if not isinstance(sl, dict):
            continue
        if sl.get("do_not_include"):
            continue
        slide_export_n += 1
        n = _coerce_int(sl.get("number"), default=0)
        title = _as_str(sl.get("slidetitle")).rstrip()
        subtitle = _as_str(sl.get("subtitle")).rstrip()
        content = _as_str(sl.get("slidecontent")).rstrip()
        body_parts = [f"Title: {title}", f"Subtitle: {subtitle}", f"Content:\n{content}"]
        body = "\n\n".join(body_parts)
        blocks.append(f"Slide {n} (Slide {slide_export_n})\n{sep}\n{body}")
    return "\n\n".join(blocks) + ("\n" if blocks else "")


_PATH_SEGMENT_INVALID = re.compile(r'[/\\:*?"<>|\x00-\x1f]')
_WIN_RESERVED_NAMES = frozenset(
    {"con", "prn", "aux", "nul"}
    | {f"com{i}" for i in range(1, 10)}
    | {f"lpt{i}" for i in range(1, 10)}
)


def _sanitize_category_folder_name(category: Any) -> str:
    """Turn course Category into a single filesystem-safe directory name."""
    raw = _as_str(category).strip()
    if not raw:
        return "Uncategorized"
    name = _PATH_SEGMENT_INVALID.sub("_", raw)
    name = re.sub(r"\s+", " ", name).strip()
    name = name.rstrip(". ")
    if not name or name in (".", ".."):
        return "Uncategorized"
    if name.casefold() in _WIN_RESERVED_NAMES:
        name = f"_{name}_"
    return name


def _export_course_package(json_path: Path, data: dict[str, Any]) -> Path:
    """Create a folder under (parent of json folder)/package exports/<Category>/<stem>/ with templates, script, and slides text."""
    stem = json_path.stem
    category_dir = _sanitize_category_folder_name(data.get("Category"))
    package_parent = json_path.parent.parent / "package exports"
    project_dir = package_parent / category_dir / stem
    if project_dir.exists():
        shutil.rmtree(project_dir)
    project_dir.mkdir(parents=True)
    for name in ("Archives", "Camtasia_Files", "Export", "Script", "Slides"):
        (project_dir / name).mkdir()

    if not HELP_TXT.is_file():
        raise FileNotFoundError(f"Missing supporting file: {HELP_TXT}")
    if not SLIDES_TEMPLATE_PPTX.is_file():
        raise FileNotFoundError(f"Missing supporting file: {SLIDES_TEMPLATE_PPTX}")
    if not CAMTASIA_TEMPLATE_CMPROJ.exists():
        raise FileNotFoundError(f"Missing supporting Camtasia project: {CAMTASIA_TEMPLATE_CMPROJ}")

    shutil.copy2(HELP_TXT, project_dir / "Help.txt")
    slides_dir = project_dir / "Slides"
    shutil.copy2(SLIDES_TEMPLATE_PPTX, slides_dir / f"{stem}.pptx")
    slides_text = _build_slides_export_text(data)
    (slides_dir / f"{stem}.txt").write_text(slides_text, encoding="utf-8")
    dest_cmproj = project_dir / "Camtasia_Files" / f"{stem}.cmproj"
    if CAMTASIA_TEMPLATE_CMPROJ.is_dir():
        shutil.copytree(CAMTASIA_TEMPLATE_CMPROJ, dest_cmproj)
    else:
        shutil.copy2(CAMTASIA_TEMPLATE_CMPROJ, dest_cmproj)

    script_text = _build_script_export_text(data)
    (project_dir / "Script" / f"{stem}.txt").write_text(script_text, encoding="utf-8")
    return project_dir


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


def _lesson0_description(data: dict[str, Any]) -> str:
    lessons = data.get("lessons")
    if not isinstance(lessons, list) or not lessons:
        return ""
    l0 = lessons[0]
    if not isinstance(l0, dict):
        return ""
    return _as_str(l0.get("Description", ""))


def _build_audience_source_text(data: dict[str, Any]) -> str:
    """Plain-text bundle for Claude from course + first-lesson description."""
    chunks = [
        f"Course Title:\n{_as_str(data.get('Course Title'))}",
        f"Category:\n{_as_str(data.get('Category'))}",
        f"Description:\n{_as_str(data.get('Description'))}",
        f"Prerequisites:\n{_as_str(data.get('Prerequisites'))}",
        f"Learning objectives:\n{_as_str(data.get('Learning objectives'))}",
        f"Lesson description (first lesson):\n{_lesson0_description(data)}",
    ]
    return "\n\n".join(chunks)


def _parse_claude_audience_json(raw: str) -> tuple[str, str]:
    """Extract who-is-this-for and team/dept strings from model output."""
    text = raw.strip()
    m = _CLAUDE_AUDIENCE_JSON_HINT.search(text)
    if m:
        text = m.group(1).strip()
    start = text.find("{")
    end = text.rfind("}")
    if start < 0 or end <= start:
        raise ValueError("Model response did not contain a JSON object")
    obj = json.loads(text[start : end + 1])
    if not isinstance(obj, dict):
        raise ValueError("JSON root must be an object")
    who = obj.get("who_is_this_for")
    teams = obj.get("team_or_dept")
    if who is None:
        who = obj.get("Who is this for") or obj.get("whoIsThisFor")
    if teams is None:
        teams = obj.get("Team or Dept this is for") or obj.get("teamOrDept")
    if who is None or teams is None:
        raise ValueError('Expected keys "who_is_this_for" and "team_or_dept" in JSON')
    return _as_str(who).strip(), _as_str(teams).strip()


def _call_claude_for_audience(*, api_key: str, source_text: str, model: str) -> tuple[str, str]:
    if not anthropic:
        raise RuntimeError('Install the anthropic package: pip install anthropic (see review/requirements.txt)')
    system = (
        "You label corporate training courses for a general workforce (non-specialists). "
        "Courses are short awareness-style modules, not aimed at dedicated security engineers—"
        "but you should still name concrete job roles and departments because copy is used for SEO and discovery. "
        "Stay faithful to the supplied content; do not invent unrelated personas or org units."
    )
    user = (
        "From the course metadata below, write two catalog fields.\n\n"
        "1) who_is_this_for: Job roles and personas that benefit (concrete titles people search for). "
        "Comma-separated phrases or short prose is fine.\n"
        "2) team_or_dept: Teams or departments that typically assign or care about this topic; "
        "include recognizable names (e.g. Operations, HR, IT, Legal) when they fit the content.\n\n"
        'Respond with ONLY a JSON object, no markdown fences, using exactly these keys: '
        '"who_is_this_for" and "team_or_dept". Both values must be strings.\n\n'
        "---\n"
        f"{source_text}\n"
        "---"
    )
    client = anthropic.Anthropic(api_key=api_key)
    message = client.messages.create(
        model=model,
        max_tokens=2048,
        messages=[{"role": "user", "content": user}],
        system=system,
    )
    parts: list[str] = []
    for block in message.content:
        if block.type == "text":
            parts.append(block.text)
    combined = "".join(parts)
    return _parse_claude_audience_json(combined)


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

    jump_to = st.number_input(
        "Jump to file #",
        min_value=1,
        max_value=len(file_labels),
        value=st.session_state["file_index"] + 1,
        step=1,
        key="jump_to_file_number",
    )
    if int(jump_to) - 1 != st.session_state["file_index"]:
        st.session_state["file_index"] = int(jump_to) - 1

    selected_label = st.selectbox(
        "Select a JSON file",
        options=file_labels,
        index=st.session_state["file_index"],
    )
    # Keep the index in sync if the user picks from the dropdown directly
    st.session_state["file_index"] = file_labels.index(selected_label)
    selected_path = next(p for p in files if p.name == selected_label)

    reload_clicked = st.button("Reload from disk")

    st.subheader("AI audience & teams (Claude)")
    st.caption(
        "Fills “Who is this for” and “Team or Dept this is for” from title, category, description, "
        "prereqs, learning objectives, and the first lesson description. Audience is general workforce "
        "(not security specialists), with concrete roles and departments for SEO."
    )
    ai_env_key = (os.environ.get("ANTHROPIC_API_KEY") or "").strip()
    ai_key_input = st.text_input(
        "API key (optional if ANTHROPIC_API_KEY is set)",
        type="password",
        key="claude_api_key_sidebar",
    )
    ai_key = ai_env_key or ai_key_input.strip()
    ai_model = st.text_input("Claude model id", value=DEFAULT_CLAUDE_MODEL, key="claude_model_sidebar")
    ai_backup = st.checkbox("Backup each file (.bak) before write", value=True, key="claude_backup_sidebar")
    r1, r2, r3 = st.columns([1, 1, 2])
    with r1:
        ai_from = st.number_input(
            "From file #",
            min_value=1,
            max_value=len(files),
            value=1,
            help="Start of range in the sorted JSON list (1-based).",
            key="ai_range_from",
        )
    with r2:
        ai_to = st.number_input(
            "To file #",
            min_value=1,
            max_value=len(files),
            value=min(5, len(files)),
            help="End of range (inclusive).",
            key="ai_range_to",
        )
    with r3:
        st.write("")
        deps_ok = anthropic is not None
        if not deps_ok:
            st.caption('Install: pip install "anthropic>=0.25"')
        ai_run = st.button(
            "Fill audience fields (Claude)",
            use_container_width=True,
            disabled=not ai_key or not deps_ok,
            key="ai_batch_run",
        )

    if ai_run and ai_key and deps_ok:
        a, b = int(ai_from), int(ai_to)
        if a > b:
            a, b = b, a
        slice_paths = files[a - 1 : b]
        prog = st.progress(0)
        status = st.empty()
        ok_names: list[str] = []
        fail_msgs: list[str] = []
        model_id = (ai_model or "").strip() or DEFAULT_CLAUDE_MODEL
        total = len(slice_paths)
        for idx, fp in enumerate(slice_paths):
            status.caption(f"Calling Claude for ({idx + 1}/{total}) {fp.name}…")
            prog.progress((idx + 1) / max(total, 1))
            try:
                disk_data = _load_json(fp)
                bundle = _build_audience_source_text(disk_data)
                who, teams = _call_claude_for_audience(
                    api_key=ai_key, source_text=bundle, model=model_id
                )
                disk_data["Who is this for"] = who
                disk_data["Team or Dept this is for"] = teams
                if ai_backup and fp.exists():
                    _backup_file(fp)
                _atomic_write_json(fp, disk_data)
                ok_names.append(fp.name)
            except Exception as e:
                fail_msgs.append(f"{fp.name}: {e!s}")
        status.caption("")
        if ok_names:
            st.success(f"Updated {len(ok_names)} file(s).")
        if fail_msgs:
            st.error("Some files failed:\n\n" + "\n\n".join(fail_msgs))
        if ok_names:
            st.info(
                "If an updated file is open in the editor, click **Reload from disk** to refresh fields."
            )


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

export_clicked = False
package_export_clicked = False
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

    export_clicked = st.button("Export Script")
    package_export_clicked = st.button("Export as Package")

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
            data["Who is this for"] = st.text_area(
                "Who is this for",
                value=_as_str(data.get("Who is this for", "")),
                height=100,
            )
            data["Team or Dept this is for"] = st.text_area(
                "Team or Dept this is for",
                value=_as_str(data.get("Team or Dept this is for", "")),
                height=100,
            )
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
        template={"slidetitle": "", "subtitle": "", "slidecontent": "", "do_not_include": False},
    )
    lesson0["slides"] = lesson0_slides

    lesson0_scripts = _ensure_numbered_items(
        lesson0.get("scripts"),
        count=8,
        template={"script": "", "do_not_include": False},
    )
    lesson0["scripts"] = lesson0_scripts
    scripts_by_number: dict[int, dict[str, Any]] = {}
    for sc in lesson0_scripts:
        n = _coerce_int(sc.get("number"), default=0)
        if n:
            scripts_by_number[n] = sc

    with tabs[2]:
        st.caption("Edit slide title, subtitle, and bullet content.")
        for slide in lesson0_slides:
            n = _coerce_int(slide.get("number"), default=0)
            with st.expander(f"Slide {n}", expanded=True):
                do_not_include = st.checkbox(
                    "Do Not Include",
                    value=bool(slide.get("do_not_include", False)),
                    key=f"slide_do_not_include_{n}_{path.name}",
                )
                slide["do_not_include"] = bool(do_not_include)
                if n in scripts_by_number:
                    scripts_by_number[n]["do_not_include"] = bool(do_not_include)

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

    with tabs[3]:
        st.caption("Edit narration scripts for each slide.")
        for sc in lesson0_scripts:
            n = _coerce_int(sc.get("number"), default=0)
            with st.expander(f"Script {n}", expanded=True):
                st.checkbox(
                    "Do Not Include (from Slide)",
                    value=bool(sc.get("do_not_include", False)),
                    key=f"script_do_not_include_{n}_{path.name}",
                    disabled=True,
                )
                sc["script"] = st.text_area(
                    f"Slide {n} script",
                    value=_as_str(sc.get("script", "")),
                    height=220,
                    key=f"script_{n}_{path.name}",
                )

    with tabs[4]:
        st.json(data, expanded=False)

if export_clicked:
    try:
        export_dir = path.parent / "script exports"
        export_dir.mkdir(parents=True, exist_ok=True)
        out_path = export_dir / path.with_suffix(".txt").name
        text = _build_script_export_text(data)
        out_path.write_text(text, encoding="utf-8")
        st.success(f"Script export saved: {out_path}")
    except Exception as e:
        st.error(f"Export failed: {e!s}")

if package_export_clicked:
    try:
        out_project = _export_course_package(path, data)
        st.success(f"Package export saved: {out_project}")
    except Exception as e:
        st.error(f"Package export failed: {e!s}")

