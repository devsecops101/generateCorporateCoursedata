from __future__ import annotations

import argparse
import json
import os
import re
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd


@dataclass(frozen=True)
class CourseRow:
    course_title: str
    description: str
    category: str


CLAUDE_PROMPTS: dict[str, str] = {
    "prerequisites": """You are an expert course designer.

Given:
- Course Title: "{course_title}"
- Category: "{category}"
- Description: "{description}"

Write exactly 1 concise sentence describing the prerequisites for this course.
Constraints:
- Beginner-friendly, realistic, and specific
- Do not mention "Claude", "AI model", or that you are generating text
- Do not use em dashes (—) or en dashes (–). Use commas or parentheses instead.
- No bullet points, no quotes
Return only the sentence.""",
    "learning_objectives": """You are an expert course designer.

Given:
- Course Title: "{course_title}"
- Category: "{category}"
- Description: "{description}"

Write exactly 2 sentences listing learning objectives.
Constraints:
- Beginner-friendly, action-oriented (use verbs like "identify", "explain", "apply")
- No bullet points, no numbering
- Do not mention "Claude", "AI model", or that you are generating text
- Do not use em dashes (—) or en dashes (–). Use commas or parentheses instead.
Return only the 2 sentences.""",
    "lesson_description_expand": """You are an expert instructional designer.

Given:
- Course Title: "{course_title}"
- Lesson Title: "{lesson_title}"
- Course Description: "{description}"

Expand the lesson description into 3-5 sentences.
Constraints:
- Keep it aligned to the course title and description
- Beginner level, clear and practical
- Do not mention "Claude", "AI model", or that you are generating text
- Do not use em dashes (—) or en dashes (–). Use commas or parentheses instead.
Return only the paragraph.""",
    "slide_title": """You are an expert instructional designer.

Given:
- Course Title: "{course_title}"
- Lesson Title: "{lesson_title}"
- Slide Number: {slide_number}
- Course Description: "{description}"

This course is 8 slides total (about 1 minute per slide). Titles should reflect a logical flow from introduction to conclusion.
For slide 8, the title must indicate a closing summary (e.g., "Key points", "What to remember", "Wrap-up") but do NOT use the word "takeaway" or "takeaways".
Write a short slide title (max 5 words).
Do not use em dashes (—) or en dashes (–).
Return only the title.""",
    "slide_subtitle": """You are an expert instructional designer.

Given:
- Course Title: "{course_title}"
- Lesson Title: "{lesson_title}"
- Slide Number: {slide_number}
- Slide Title: "{slide_title}"
- Course Description: "{description}"

This course is 8 slides total (about 1 minute per slide). Subtitles should reinforce a logical flow.
For slide 8, the subtitle must reinforce key things to remember without using the word "takeaway" or "takeaways".
Write a short slide subtitle (max 6 words) that complements the slide title.
Do not use em dashes (—) or en dashes (–).
Return only the subtitle.""",
    "slide_content": """You are an expert instructional designer.

Given:
- Course Title: "{course_title}"
- Lesson Title: "{lesson_title}"
- Slide Number: {slide_number}
- Slide Title: "{slide_title}"
- Slide Subtitle: "{slide_subtitle}"
- Course Description: "{description}"

Write slide body content as 3-6 concise bullet points.
Constraints:
- Each bullet is one sentence fragment (no periods)
- Beginner-friendly, concrete, no fluff
- Do not mention "Claude", "AI model", or that you are generating text
- Do not use em dashes (—) or en dashes (–). Use commas or parentheses instead.
 - The 8-slide flow should progress logically, from context and definitions to practical guidance and a closing summary
 - For slide 8, include key points to remember and a short next-step suggestion, but do NOT use the word "takeaway" or "takeaways"
Return only the bullet list.""",
    "slide_script": """You are an expert speaking coach and course presenter.

Given:
- Course Title: "{course_title}"
- Lesson Title: "{lesson_title}"
- Slide Number: {slide_number}
- Slide Title: "{slide_title}"
- Slide Subtitle: "{slide_subtitle}"
- Slide Content (bullets):
{slide_content}

Write a presenter narration script for this slide.
Constraints:
- Aim for about 1 minute of speaking time (roughly 130 to 160 words)
- Conversational, clear, beginner friendly
- Do not read the bullets verbatim; explain them naturally
- Do not use em dashes (—) or en dashes (–)
- Do not use the word "takeaway" or "takeaways"
Return only the narration text, as a single paragraph.""",
}

DEFAULT_CLAUDE_MODEL = "auto"


def _normalize_col_name(name: str) -> str:
    return " ".join(str(name).strip().lower().split())

def _anthropic_list_models(api_key: str) -> list[str]:
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/models",
        method="GET",
        headers={
            "anthropic-version": "2023-06-01",
            "x-api-key": api_key,
        },
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        raw = resp.read().decode("utf-8")
    data = json.loads(raw)
    return [m["id"] for m in data.get("data", []) if isinstance(m, dict) and "id" in m]


def _resolve_model_id(requested: str, *, api_key: str) -> str:
    """
    Resolve a requested model/alias to an actual model id available for this key.
    If requested is 'auto', picks a reasonable default.
    """
    ids = _anthropic_list_models(api_key)
    if not ids:
        raise RuntimeError("No models returned from /v1/models; check your API key/permissions.")

    req = (requested or "").strip()
    if not req or req.lower() == "auto":
        # Prefer Sonnet-family models if present.
        for mid in ids:
            if "sonnet" in mid.lower():
                return mid
        return ids[0]

    if req in ids:
        return req

    # Heuristic: if they asked for a family, pick the first matching id.
    lowered = req.lower()
    for mid in ids:
        if lowered in mid.lower():
            return mid

    # Otherwise, prefer sonnet, else newest.
    for mid in ids:
        if "sonnet" in mid.lower():
            return mid
    return ids[0]


def _load_env_file(path: Path) -> None:
    """
    Loads KEY=VALUE pairs into os.environ (best-effort).
    - Ignores blank lines and comments starting with '#'
    - Supports optional 'export KEY=VALUE'
    - Strips surrounding single/double quotes from VALUE
    """
    if not path.exists() or not path.is_file():
        return
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("export "):
            line = line[len("export ") :].strip()
        if "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip()
        if (value.startswith('"') and value.endswith('"')) or (value.startswith("'") and value.endswith("'")):
            value = value[1:-1]
        if key and key not in os.environ:
            os.environ[key] = value


def _ensure_anthropic_key_loaded(repo_root: Path) -> None:
    if os.environ.get("ANTHROPIC_API_KEY"):
        return
    candidates = [
        Path.cwd() / ".env",
        Path.cwd() / ".venv" / ".env",
        repo_root / ".env",
        repo_root / "code" / ".env",
        repo_root / "code" / ".venv" / ".env",
    ]
    for p in candidates:
        _load_env_file(p)
        if os.environ.get("ANTHROPIC_API_KEY"):
            return


def _safe_filename(stem: str) -> str:
    s = stem.strip().lower()
    s = re.sub(r"[^a-z0-9]+", "-", s).strip("-")
    return s or "course"

class ClaudeTextGenerator:
    def __init__(self, *, api_key: str, requested_model: str):
        try:
            from anthropic import Anthropic  # type: ignore
        except Exception as e:  # pragma: no cover
            raise RuntimeError(
                "Missing dependency 'anthropic'. Install it with: pip install -r requirements.txt"
            ) from e

        self._api_key = api_key
        self._client = Anthropic(api_key=api_key)
        self._requested_model = (requested_model or "").strip() or "auto"
        self.model_id = _resolve_model_id(self._requested_model, api_key=api_key)
        self.did_fallback = self.model_id != self._requested_model and self._requested_model.lower() != "auto"

    def generate(self, *, prompt: str, max_tokens: int = 500) -> str:
        try:
            msg = self._client.messages.create(
                model=self.model_id,
                max_tokens=max_tokens,
                temperature=0.4,
                messages=[{"role": "user", "content": prompt}],
            )
        except Exception:
            # One-time fallback if the resolved model still fails at runtime.
            fallback = _resolve_model_id("auto", api_key=self._api_key)
            if fallback != self.model_id:
                self.model_id = fallback
                self.did_fallback = True
                msg = self._client.messages.create(
                    model=self.model_id,
                    max_tokens=max_tokens,
                    temperature=0.4,
                    messages=[{"role": "user", "content": prompt}],
                )
            else:
                raise

        parts: list[str] = []
        for block in getattr(msg, "content", []) or []:
            if getattr(block, "type", None) == "text":
                parts.append(getattr(block, "text", ""))
        return "\n".join(p.strip("\n") for p in parts).strip()


def load_courses_from_excel(path: str | Path, *, sheet_name: str | int | None = 0) -> list[CourseRow]:
    """
    Reads an Excel file and returns rows in memory.

    Expected columns (case/whitespace-insensitive):
    - Course Title
    - Description
    - category
    """
    df = pd.read_excel(path, sheet_name=sheet_name)

    normalized_to_original: dict[str, str] = {_normalize_col_name(c): c for c in df.columns}

    def _pick_column(*aliases: str) -> str:
        for a in aliases:
            key = _normalize_col_name(a)
            if key in normalized_to_original:
                return normalized_to_original[key]
        raise KeyError(f"No match for any of: {list(aliases)}")

    # Accept common header variants / aliases (case/whitespace-insensitive).
    try:
        col_title = _pick_column("Course Title", "Title")
        col_description = _pick_column("Description")
        col_category = _pick_column("category", "Category")
    except KeyError as e:
        available = ", ".join(map(str, df.columns))
        raise KeyError(f"{e}. Available columns: [{available}]") from None

    subset = df[[col_title, col_description, col_category]].copy()

    def _cell_to_str(v: Any) -> str:
        if pd.isna(v):
            return ""
        return str(v).strip()

    courses: list[CourseRow] = []
    for _, row in subset.iterrows():
        courses.append(
            CourseRow(
                course_title=_cell_to_str(row[col_title]),
                description=_cell_to_str(row[col_description]),
                category=_cell_to_str(row[col_category]),
            )
        )

    return courses


def course_to_detail_json(
    course: CourseRow,
    *,
    use_claude: bool = False,
    claude_api_key: str | None = None,
    claude_model: str = DEFAULT_CLAUDE_MODEL,
) -> dict[str, Any]:
    prereq = "[PLACEHOLDER: generate 1 sentence with Claude AI]"
    objectives = "[PLACEHOLDER: generate 2 sentences with Claude AI]"
    lesson_desc = "[PLACEHOLDER: expand description with Claude AI]"

    lesson_title = course.course_title

    slide_titles: dict[int, str] = {}
    slide_subtitles: dict[int, str] = {}
    slide_contents: dict[int, str] = {}
    slide_scripts: dict[int, str] = {}

    if use_claude:
        if not claude_api_key:
            raise RuntimeError(
                "ANTHROPIC_API_KEY is not set. Export it (or pass via environment) before using --use-claude."
            )

        claude = ClaudeTextGenerator(api_key=claude_api_key, requested_model=claude_model)

        prereq = claude.generate(
            prompt=CLAUDE_PROMPTS["prerequisites"].format(
                course_title=course.course_title,
                category=course.category,
                description=course.description,
            ),
            max_tokens=120,
        )

        objectives = claude.generate(
            prompt=CLAUDE_PROMPTS["learning_objectives"].format(
                course_title=course.course_title,
                category=course.category,
                description=course.description,
            ),
            max_tokens=180,
        )

        lesson_desc = claude.generate(
            prompt=CLAUDE_PROMPTS["lesson_description_expand"].format(
                course_title=course.course_title,
                lesson_title=lesson_title,
                description=course.description,
            ),
            max_tokens=350,
        )

        for n in range(1, 9):
            title = claude.generate(
                prompt=CLAUDE_PROMPTS["slide_title"].format(
                    course_title=course.course_title,
                    lesson_title=lesson_title,
                    slide_number=n,
                    description=course.description,
                ),
                max_tokens=40,
            )
            slide_titles[n] = title

            subtitle = claude.generate(
                prompt=CLAUDE_PROMPTS["slide_subtitle"].format(
                    course_title=course.course_title,
                    lesson_title=lesson_title,
                    slide_number=n,
                    slide_title=title,
                    description=course.description,
                ),
                max_tokens=60,
            )
            slide_subtitles[n] = subtitle

            content = claude.generate(
                prompt=CLAUDE_PROMPTS["slide_content"].format(
                    course_title=course.course_title,
                    lesson_title=lesson_title,
                    slide_number=n,
                    slide_title=title,
                    slide_subtitle=subtitle,
                    description=course.description,
                ),
                max_tokens=350,
            )
            slide_contents[n] = content

            script = claude.generate(
                prompt=CLAUDE_PROMPTS["slide_script"].format(
                    course_title=course.course_title,
                    lesson_title=lesson_title,
                    slide_number=n,
                    slide_title=title,
                    slide_subtitle=subtitle,
                    slide_content=content,
                ),
                max_tokens=500,
            )
            slide_scripts[n] = script

    return {
        "Course Title": course.course_title,
        "Category": course.category,
        "Description": course.description,
        "Duration": "8",
        "Difficulty": "Beginner",
        "Price": "$29",
        "Instructor": "AI Security",
        "Published Status": "Draft",
        "Prerequisites": prereq,
        "Learning objectives": objectives,
        "lessons": [
            {
                "lesson title": course.course_title,
                "order": 1,
                "Description": lesson_desc,
                "status": "draft",
                "content": course.description,
                "slides": [
                    {
                        "number": n,
                        "slidetitle": (
                            slide_titles.get(n) or f"[PLACEHOLDER: slide {n} title via Claude AI]"
                        ),
                        "subtitle": (
                            slide_subtitles.get(n)
                            or f"[PLACEHOLDER: slide {n} subtitle via Claude AI]"
                        ),
                        "slidecontent": (
                            slide_contents.get(n)
                            or f"[PLACEHOLDER: slide {n} content via Claude AI]"
                        ),
                    }
                    for n in range(1, 9)
                ],
                "scripts": [
                    {
                        "number": n,
                        "script": (
                            slide_scripts.get(n)
                            or f"[PLACEHOLDER: slide {n} 1-minute narration script via Claude AI]"
                        ),
                    }
                    for n in range(1, 9)
                ],
            }
        ],
    }


def _default_output_dir(repo_root: Path) -> Path:
    candidates = [
        repo_root / "data",
        repo_root / "code" / "data",
    ]
    for c in candidates:
        if c.exists() and c.is_dir():
            return c
    return candidates[0]


def main() -> int:
    parser = argparse.ArgumentParser(description="Load courses from an Excel file into memory.")
    repo_root = Path(__file__).resolve().parents[1]
    default_candidates = [
        repo_root / "data" / "coursedata.xlsx",
        Path(__file__).resolve().parent / "data" / "coursedata.xlsx",
    ]
    default_excel = next((p for p in default_candidates if p.exists()), default_candidates[0])
    parser.add_argument(
        "excel_path",
        nargs="?",
        default=str(default_excel),
        help="Path to .xlsx file (default: data/coursedata.xlsx or code/data/coursedata.xlsx)",
    )
    parser.add_argument(
        "--rows",
        type=int,
        default=1,
        help="Number of rows to generate JSON for, starting from row 1 (default: 1).",
    )
    parser.add_argument(
        "--sheet",
        default="0",
        help="Sheet name or index (default: 0). Use an integer for index, otherwise a name.",
    )
    parser.add_argument(
        "--use-claude",
        action="store_true",
        help="Generate prerequisites/objectives/lesson/slides via Claude (requires ANTHROPIC_API_KEY).",
    )
    parser.add_argument(
        "--claude-model",
        default=DEFAULT_CLAUDE_MODEL,
        help='Claude model id (or "auto"). If invalid, falls back automatically (default: auto).',
    )
    parser.add_argument(
        "--out-dir",
        default="",
        help="Output directory for JSON files (default: repo data/ folder, fallback: code/data/).",
    )
    args = parser.parse_args()

    if args.rows < 1:
        raise SystemExit("--rows must be >= 1")

    sheet: str | int | None
    if args.sheet.strip().isdigit():
        sheet = int(args.sheet.strip())
    else:
        sheet = args.sheet

    courses = load_courses_from_excel(args.excel_path, sheet_name=sheet)
    print(f"Loaded {len(courses)} course rows into memory.")

    out_dir = Path(args.out_dir).expanduser().resolve() if args.out_dir else _default_output_dir(repo_root)
    out_dir.mkdir(parents=True, exist_ok=True)

    if args.use_claude:
        _ensure_anthropic_key_loaded(repo_root)

    to_generate = courses[: args.rows]
    claude_key = os.environ.get("ANTHROPIC_API_KEY")
    if args.use_claude and claude_key:
        # Resolve once and show the actual model ID used.
        resolved = _resolve_model_id(str(args.claude_model), api_key=claude_key)
        requested = str(args.claude_model).strip() or "auto"
        if requested.lower() != "auto" and resolved != requested:
            print(f'Claude model "{requested}" not available; using "{resolved}" instead.')
        else:
            print(f'Using Claude model: "{resolved}"')
    for i, course in enumerate(to_generate, start=1):
        payload = course_to_detail_json(
            course,
            use_claude=bool(args.use_claude),
            claude_api_key=claude_key,
            claude_model=str(args.claude_model),
        )
        filename = f"{i:03d}-{_safe_filename(course.course_title)}.json"
        path = out_dir / filename
        path.write_text(json.dumps(payload, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")

    print(f"Wrote {len(to_generate)} JSON file(s) to: {out_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

