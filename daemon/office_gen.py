"""
Sidekick OS — Step 5.3 v2: Human-in-the-Loop Office Generator (`daemon/office_gen.py`)

Supports `sk doc`, `sk excel`, and `sk ppt` via an iterative 2-stage workflow:
  Stage 1: Research topic via `researcher.py` + draft structured outline via
           `router.py` (with optional `.pptx` organizer template rubric & user
           instructions), shown in a Dark Cyber-HUD with:
             - Upfront live slide/section/sheet count badge
             - Inline "Revise with AI" feedback loop
             - Optional `.pptx` organizer template picker (keeps master theme/branding)
             - Editable `MERMAID:` diagram blocks + `TAKEAWAY:` callouts
  Stage 2: Render Mermaid architecture/flowcharts (Playwright headless SVG->PNG
           with local matplotlib flowchart fallback), search optimized 2-3 keyword
           visuals (Wikimedia API + Chrome CDP web image search), and compile
           dense, multi-layout `.docx`, `.xlsx`, or `.pptx` files into
           `~/Documents/Chitran_Outputs/`.
"""

from __future__ import annotations

import argparse
import asyncio
import base64
from datetime import datetime
import json
from pathlib import Path
import re
import sys
import threading
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import quote_plus

if sys.stdout is None:
    sys.stdout = open("nul", "w", encoding="utf-8")
if sys.stderr is None:
    sys.stderr = open("nul", "w", encoding="utf-8")

DAEMON_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = DAEMON_DIR.parent
if str(DAEMON_DIR) not in sys.path:
    sys.path.insert(0, str(DAEMON_DIR))

import httpx
from PIL import Image

from browser import browser_manager, is_cdp_alive
from notify import notify_windows
from researcher import discover_sources, scrape_all_sources, slugify_topic
from router import query_mentor
from win_ui import apply_win11_dwm_style, enable_high_dpi

OUTPUT_DIR = Path.home() / "Documents" / "Chitran_Outputs"
LEGACY_OUTPUT_DIR = Path.home() / "Documents" / "Sidekick_Outputs"
DRAFTS_DIR = PROJECT_ROOT / "data" / "office_drafts"
ASSETS_DIR = PROJECT_ROOT / "data" / "office_assets"
VENDOR_DIR = PROJECT_ROOT / "data" / "vendor"
MERMAID_JS_PATH = VENDOR_DIR / "mermaid.min.js"
MERMAID_CDN_URL = "https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.min.js"

VALID_TYPES = ("doc", "excel", "ppt")
EXT_MAP = {"doc": ".docx", "excel": ".xlsx", "ppt": ".pptx"}
TYPE_LABELS = {
    "doc": "WORD (.DOCX)",
    "excel": "EXCEL (.XLSX)",
    "ppt": "POWERPOINT (.PPTX)",
}

QUERY_STOPWORDS = {
    "a", "an", "and", "architecture", "diagram", "for", "from", "how", "illustration",
    "in", "into", "of", "on", "or", "overview", "system", "the", "to", "tutorial",
    "using", "visual", "with", "workflow", "flowchart",
}
BLOCKED_IMAGE_DOMAINS = (
    "shutterstock.com",
    "istockphoto.com",
    "alamy.com",
    "gettyimages.com",
    "dreamstime.com",
    "123rf.com",
    "depositphotos.com",
    "vecteezy.com",
)

# Dark Cyber-HUD palette matching daemon/hud.py & daemon/organizer.py
BG_ROOT = "#121214"
BG_CARD = "#18181B"
BG_INPUT = "#202026"
BORDER_SUBTLE = "#27272A"
TEXT_PRIMARY = "#F4F4F5"
TEXT_MUTED = "#A1A1AA"
ACCENT_CYAN = "#22D3EE"
ACCENT_EMERALD = "#10B981"
ACCENT_AMBER = "#F59E0B"
ACCENT_RED = "#F87171"
BTN_CANCEL_BG = "#27272A"


def normalize_doc_type(raw_type: str) -> str:
    val = (raw_type or "doc").strip().lower()
    if val in ("doc", "docx", "word", "report"):
        return "doc"
    if val in ("excel", "xlsx", "sheet", "sheets", "spreadsheet"):
        return "excel"
    if val in ("ppt", "pptx", "slides", "deck", "presentation"):
        return "ppt"
    raise ValueError(f"Unsupported office type '{raw_type}'. Use doc, excel, or ppt.")


def resolve_unique_output_path(output_dir: Path, base_name: str, ext: str) -> Path:
    """Return a non-colliding file path in output_dir so existing files are never overwritten."""
    output_dir.mkdir(parents=True, exist_ok=True)
    candidate = output_dir / f"{base_name}{ext}"
    counter = 1
    while candidate.exists():
        candidate = output_dir / f"{base_name} ({counter}){ext}"
        counter += 1
    return candidate


def optimize_image_query(raw_query: str, fallback_topic: str = "") -> str:
    """Distill a verbose phrase down to 2-3 high-signal search keywords for image APIs."""
    source_text = (raw_query or fallback_topic or "").strip()
    tokens = re.findall(r"[A-Za-z0-9]+", source_text)
    keywords = [t for t in tokens if len(t) >= 2 and t.lower() not in QUERY_STOPWORDS]
    if not keywords:
        keywords = [t for t in re.findall(r"[A-Za-z0-9]+", fallback_topic or "") if len(t) >= 2]
    return " ".join(keywords[:3])


def summarize_outline_counts(doc_type: str, outline: Dict[str, Any], source_count: int = 0) -> str:
    """Return a concise upfront human summary of slide/section/sheet counts."""
    if doc_type == "ppt":
        content_n = len(outline.get("slides", []))
        ref_n = 1 if source_count > 0 else 0
        total = 1 + content_n + ref_n
        parts = ["1 Cover", f"{content_n} Content"]
        if ref_n:
            parts.append("1 References")
        return f"Total: {total} slides ({' + '.join(parts)})"
    if doc_type == "doc":
        sec_n = len(outline.get("sections", []))
        return f"Total: {sec_n} section(s)"
    sheets = outline.get("sheets", [])
    rows_n = sum(len(s.get("rows", [])) for s in sheets)
    return f"Total: {len(sheets)} sheet(s) • {rows_n} row(s)"


def inspect_pptx_template(template_path: Optional[Path]) -> Dict[str, Any]:
    """Read slide dimensions and existing slide rubric hints from a hackathon/college .pptx template."""
    if not template_path:
        return {}
    t_path = Path(template_path)
    if not t_path.exists() or t_path.suffix.lower() != ".pptx":
        return {}

    try:
        from pptx import Presentation

        prs = Presentation(str(t_path))
        sw_in = round(prs.slide_width / 914400, 3)
        sh_in = round(prs.slide_height / 914400, 3)

        rubric_slides: List[str] = []
        dark_bg = False
        for idx, slide in enumerate(prs.slides, 1):
            texts: List[str] = []
            for shape in slide.shapes:
                if getattr(shape, "has_text_frame", False):
                    txt = re.sub(r"\s+", " ", shape.text_frame.text or "").strip()
                    if txt and "click to add" not in txt.lower():
                        texts.append(txt[:140])
            if texts:
                rubric_slides.append(f"Slide {idx}: {' — '.join(texts[:2])}")

        # Detect dark solid background if explicitly set on master or slide 1
        try:
            bg_fill = prs.slide_masters[0].background.fill
            if bg_fill.type == 1 and bg_fill.fore_color and bg_fill.fore_color.rgb:
                rgb = bg_fill.fore_color.rgb
                lum = 0.2126 * rgb[0] + 0.7152 * rgb[1] + 0.0722 * rgb[2]
                dark_bg = lum < 128
        except Exception:
            pass

        return {
            "path": str(t_path),
            "name": t_path.name,
            "slide_width_in": sw_in,
            "slide_height_in": sh_in,
            "template_slide_count": len(prs.slides),
            "rubric_hints": rubric_slides[:15],
            "dark_background": dark_bg,
        }
    except Exception:
        return {}


def _extract_json_block(raw_text: str) -> Optional[Dict[str, Any]]:
    """Parse JSON object from LLM text, stripping markdown code fences if present."""
    if not raw_text:
        return None
    text = raw_text.strip()
    fence = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", text, re.DOTALL)
    if fence:
        text = fence.group(1).strip()
    else:
        brace_start = text.find("{")
        brace_end = text.rfind("}")
        if brace_start != -1 and brace_end > brace_start:
            text = text[brace_start : brace_end + 1]
    try:
        parsed = json.loads(text)
        return parsed if isinstance(parsed, dict) else None
    except Exception:
        return None


def _build_outline_prompt(
    doc_type: str,
    topic: str,
    sources: List[Dict[str, Any]],
    instructions: str = "",
    template_info: Optional[Dict[str, Any]] = None,
) -> str:
    src_summary = "\n".join(
        f"[{idx}] {s.get('title', 'Source')} ({s.get('url', '')}): "
        f"{re.sub(r'\\s+', ' ', (s.get('content') or '')).strip()[:450]}"
        for idx, s in enumerate(sources, 1)
    ) or "(No external web sources fetched; use verified engineering knowledge.)"

    extra_blocks: List[str] = []
    if instructions.strip():
        extra_blocks.append(f"User Instructions (MUST FOLLOW):\n{instructions.strip()}")
    if template_info and template_info.get("rubric_hints"):
        rubric_str = "\n".join(f"  - {h}" for h in template_info["rubric_hints"])
        extra_blocks.append(
            f"Organizer .pptx Template Guidelines (preserve structure/headings):\n{rubric_str}"
        )
    extra_section = ("\n\n" + "\n\n".join(extra_blocks)) if extra_blocks else ""

    if doc_type == "doc":
        schema_hint = (
            '{"title": "Clear Document Title", "subtitle": "Concise 1-line focus", '
            '"sections": [{"heading": "Section Title", '
            '"bullets": ["**Bold Lead-in**: Concrete engineering detail (15-24 words)"], '
            '"mermaid": "flowchart LR\\n  A[Input] --> B[Process] --> C[Output]", '
            '"image_query": "2-3 keyword visual query if no mermaid"}]}'
        )
        rules = (
            "Create 4 to 6 dense technical sections with 4 substantial bullets each "
            "(formatted '**Concept**: explanation'). Include valid `mermaid` flowchart/sequence code "
            "for architecture/workflow sections; otherwise provide a crisp 2-3 word `image_query`."
        )
    elif doc_type == "ppt":
        schema_hint = (
            '{"title": "Presentation Title", "subtitle": "Concise Deck Subtitle", '
            '"slides": [{"title": "Slide Headline", '
            '"layout": "split_visual", '
            '"bullets": ["**Lead-in**: Concrete technical explanation (14-22 words)"], '
            '"takeaway": "1 concise bottom banner key takeaway", '
            '"speaker_note": "1-2 sentence presenter note", '
            '"mermaid": "flowchart LR\\n  A[Client] --> B[API] --> C[DB]", '
            '"image_query": "2-3 word query"}]}'
        )
        rules = (
            "Create 6 to 8 high-density slides (unless user instructions or template specify another count). "
            "Each slide MUST have 4 rich bullets formatted '**Key Term**: concrete detail (14-22 words)' "
            "plus a crisp `takeaway`. Choose `layout` from `split_visual` (text + visual/mermaid), "
            "`diagram_focus`, or `two_column` (6 bullets split into two comparison cards). "
            "For architecture/pipeline slides, include concise valid `mermaid` flowchart syntax (`flowchart LR` or `sequenceDiagram`, max 7 nodes). "
            "Only set 2-3 word `image_query` on slides needing a real-world diagram/photo."
        )
    else:
        schema_hint = (
            '{"title": "Workbook Title", "summary": "What this workbook analyzes", '
            '"sheets": [{"sheet_name": "Data_Overview", "headers": ["Item", "Category", "Metric_Value", "Target", "Status"], '
            '"rows": [["Sample A", "Core", 85, 90, "On Track"]], "chart_title": "Metric Comparison"}]}'
        )
        rules = (
            "Create 1 to 2 structured sheets with 4-6 columns and 6-10 realistic numeric + categorical rows "
            "(include numeric columns suitable for SUM/AVERAGE formulas and a bar chart)."
        )

    return (
        f"Prepare a structured {TYPE_LABELS[doc_type]} outline for topic: {topic}\n\n"
        f"Research Notes:\n{src_summary}"
        f"{extra_section}\n\n"
        f"Requirements:\n- {rules}\n"
        f"- Output ONLY valid JSON matching this exact schema:\n{schema_hint}"
    )


def _fallback_outline(doc_type: str, topic: str, sources: List[Dict[str, Any]]) -> Dict[str, Any]:
    clean_title = topic.strip().rstrip(".") or "Technical Brief"
    short_kw = optimize_image_query(clean_title, clean_title)
    src_bullets = []
    for idx, src in enumerate(sources[:4], 1):
        snippet = re.sub(r"\s+", " ", (src.get("content") or "")).strip()[:160]
        src_bullets.append(f"**Source [{idx}] — {src.get('title', 'Reference')[:42]}**: {snippet or src.get('url', '')}")

    default_mermaid = (
        "flowchart LR\n"
        f"    A[Input & Context] --> B[{short_kw.title()[:18] or 'Core Engine'}]\n"
        "    B --> C[Validation & Guardrails]\n"
        "    C --> D[Structured Output]"
    )

    if doc_type == "doc":
        sections = [
            {
                "heading": "Executive Overview",
                "bullets": [
                    f"**Problem Scope**: Core objectives and production requirements for {clean_title}.",
                    "**Design Constraints**: Latency, cost ceilings, reliability, and clear failure recovery paths.",
                    "**Engineering Principle**: Prefer simple deterministic workflows before adding complex orchestration.",
                    "**Deliverables**: Measurable benchmarks and automated verification checks.",
                ],
                "mermaid": "",
                "image_query": short_kw,
            },
            {
                "heading": "Architecture & Data Flow",
                "bullets": [
                    "**Boundary Validation**: Validate external inputs, schemas, and payloads at entry points.",
                    "**Core Processing**: Separate orchestration logic cleanly from persistence and UI presentation.",
                    "**Resilience**: Handle timeouts, rate limits, and offline fallback states gracefully.",
                    "**Observability**: Log actionable error context without exposing sensitive credentials.",
                ],
                "mermaid": default_mermaid,
                "image_query": "",
            },
            {
                "heading": "Research Findings & References",
                "bullets": src_bullets or [
                    f"**Ecosystem Benchmarks**: Key architectural patterns identified for {clean_title}.",
                    "**Tradeoff Analysis**: Balancing throughput, maintainability, and zero-cost hosting tiers.",
                    "**Implementation Notes**: Reusable interfaces minimize duplicate error-handling code.",
                    "**Security Posture**: Least-privilege access and sanitized outputs across stages.",
                ],
                "mermaid": "",
                "image_query": "",
            },
            {
                "heading": "Implementation Roadmap",
                "bullets": [
                    "**Milestone 1 — MVP Core**: Build single-purpose module with unit test coverage.",
                    "**Milestone 2 — Integration**: Wire interfaces cleanly and test boundary failure cases.",
                    "**Milestone 3 — Hardening**: Verify performance under realistic workload sizing.",
                    "**Milestone 4 — Documentation**: Keep operational runbooks concise and reproducible.",
                ],
                "mermaid": "",
                "image_query": "",
            },
        ]
        return {"title": clean_title, "subtitle": "Technical Report Draft", "sections": sections}

    if doc_type == "ppt":
        slides = [
            {
                "title": f"{clean_title}: Problem & Context",
                "layout": "split_visual",
                "bullets": [
                    f"**Core Challenge**: Bottlenecks and operational friction in existing {clean_title} workflows.",
                    "**Target Outcomes**: Measurable improvements in latency, accuracy, and developer velocity.",
                    "**Constraints**: Zero-cost or bounded compute budgets with strict data privacy.",
                    "**Scope**: End-to-end pipeline from ingestion through verified output delivery.",
                ],
                "takeaway": f"Focuses {clean_title} on measurable production reliability over hype.",
                "speaker_note": f"Introduce the concrete engineering problem and measurable goals for {clean_title}.",
                "mermaid": "",
                "image_query": short_kw,
            },
            {
                "title": "System Architecture & Data Flow",
                "layout": "split_visual",
                "bullets": [
                    "**Ingestion Layer**: Normalizes raw inputs and validates schemas before downstream work.",
                    "**Processing Core**: Modular pipeline stages with clear single-responsibility interfaces.",
                    "**Guardrail Gate**: Deterministic checks prevent silent failures or unsafe mutations.",
                    "**Delivery Tier**: Clean outputs with transparent status and audit logging.",
                ],
                "takeaway": "Decoupled pipeline layers enable independent testing and safe fallbacks.",
                "speaker_note": "Walk through the Mermaid flow from left to right covering validation and output.",
                "mermaid": default_mermaid,
                "image_query": "",
            },
            {
                "title": "Tradeoffs, Security & Reliability",
                "layout": "two_column",
                "bullets": src_bullets[:4] or [
                    "**Simplicity vs Overengineering**: Start with proven primitives before distributed queues.",
                    "**Determinism vs Flexibility**: Keep critical rules deterministic; reserve AI for synthesis.",
                    "**Cost Control**: Enforce usage quotas and caching before external API expansion.",
                    "**Secret Hygiene**: Keep credentials in environment config and never inside logs.",
                    "**Graceful Degradation**: Provide deterministic local fallbacks when upstream APIs stall.",
                    "**Verification Gate**: Test boundary conditions automatically after every change.",
                ],
                "takeaway": "Resilient systems pair simple defaults with strict boundary validation.",
                "speaker_note": "Highlight architectural tradeoffs and why each choice survives production edge cases.",
                "mermaid": "",
                "image_query": "",
            },
            {
                "title": "Execution Sequence & Request Lifecycle",
                "layout": "split_visual",
                "bullets": [
                    "**Stage 1 — Discovery**: Gather verified context and structure an editable plan first.",
                    "**Stage 2 — Review Gate**: Keep human approval mandatory before persistent side effects.",
                    "**Stage 3 — Execution**: Render visual assets and compile outputs deterministically.",
                    "**Stage 4 — Audit**: Return file paths, counts, and clear failure diagnostics.",
                ],
                "takeaway": "Two-stage human approval prevents wasted compute and broken deliverables.",
                "speaker_note": "Explain the lifecycle from initial trigger to verified final artifact.",
                "mermaid": (
                    "sequenceDiagram\n"
                    "    participant U as User\n"
                    "    participant S as Engine\n"
                    "    participant G as Guardrail\n"
                    "    U->>S: Request Output\n"
                    "    S->>G: Validate & Draft Plan\n"
                    "    G-->>U: Review Gate\n"
                    "    U->>S: Approve & Compile"
                ),
                "image_query": "",
            },
            {
                "title": "Implementation Roadmap & Next Steps",
                "layout": "two_column",
                "bullets": [
                    "**Phase 1 — Foundation**: Lock schema contracts and clean module boundaries first.",
                    "**Phase 2 — Core Pipeline**: Implement deterministic handlers with meaningful defaults.",
                    "**Phase 3 — Automated Checks**: Cover nominal flows, edge inputs, and error recovery.",
                    "**Phase 4 — Release Readiness**: Review diffs, latency metrics, and operator docs.",
                ],
                "takeaway": "Ship a clean verified slice early, then expand capabilities incrementally.",
                "speaker_note": "Close with concrete engineering milestones and verification gates.",
                "mermaid": "",
                "image_query": "",
            },
        ]
        return {"title": clean_title, "subtitle": "Chitran Technical Presentation", "slides": slides}

    rows = []
    for idx in range(1, 7):
        rows.append([f"{clean_title[:18]} Item {idx}", "Core" if idx % 2 else "Optimization", 60 + idx * 6, 85, "Planned"])
    return {
        "title": clean_title,
        "summary": f"Structured analysis workbook for {clean_title}",
        "sheets": [
            {
                "sheet_name": "Analysis_Data",
                "headers": ["Item", "Category", "Current_Score", "Target_Score", "Status"],
                "rows": rows,
                "chart_title": f"{clean_title[:28]} Comparison",
            }
        ],
    }


def _normalize_outline_structure(doc_type: str, raw: Dict[str, Any], topic: str) -> Dict[str, Any]:
    title = str(raw.get("title") or topic).strip() or "Untitled Output"
    if doc_type == "doc":
        sections = []
        for sec in raw.get("sections", []):
            if not isinstance(sec, dict):
                continue
            bullets = [str(b).strip() for b in sec.get("bullets", []) if str(b).strip()]
            sections.append({
                "heading": str(sec.get("heading") or "Section").strip(),
                "bullets": bullets or ["Details to be expanded."],
                "mermaid": str(sec.get("mermaid") or "").strip(),
                "image_query": optimize_image_query(str(sec.get("image_query") or ""), sec.get("heading", ""))
                if sec.get("image_query")
                else "",
            })
        if not sections:
            return _fallback_outline("doc", topic, [])
        return {
            "title": title,
            "subtitle": str(raw.get("subtitle") or "Generated by Chitran").strip(),
            "sections": sections,
        }

    if doc_type == "ppt":
        slides = []
        for sl in raw.get("slides", []):
            if not isinstance(sl, dict):
                continue
            bullets = [str(b).strip() for b in sl.get("bullets", []) if str(b).strip()]
            layout = str(sl.get("layout") or "split_visual").strip().lower()
            if layout not in ("split_visual", "two_column", "diagram_focus"):
                layout = "split_visual"
            raw_iq = str(sl.get("image_query") or "").strip()
            slides.append({
                "title": str(sl.get("title") or "Slide").strip(),
                "layout": layout,
                "bullets": bullets or ["Key takeaway"],
                "takeaway": str(sl.get("takeaway") or "").strip(),
                "speaker_note": str(sl.get("speaker_note") or "").strip(),
                "mermaid": str(sl.get("mermaid") or "").strip(),
                "image_query": optimize_image_query(raw_iq, sl.get("title", "")) if raw_iq else "",
            })
        if not slides:
            return _fallback_outline("ppt", topic, [])
        return {
            "title": title,
            "subtitle": str(raw.get("subtitle") or "Generated by Chitran").strip(),
            "slides": slides,
        }

    sheets = []
    for sh in raw.get("sheets", []):
        if not isinstance(sh, dict):
            continue
        headers = [str(h).strip() for h in sh.get("headers", []) if str(h).strip()]
        rows = [list(r) for r in sh.get("rows", []) if isinstance(r, (list, tuple)) and r]
        if headers and rows:
            sheets.append({
                "sheet_name": re.sub(r"[:\\/?*\[\]]", "_", str(sh.get("sheet_name") or "Sheet1"))[:31],
                "headers": headers,
                "rows": rows,
                "chart_title": str(sh.get("chart_title") or f"{title[:26]} Chart").strip(),
            })
    if not sheets:
        return _fallback_outline("excel", topic, [])
    return {
        "title": title,
        "summary": str(raw.get("summary") or f"Structured workbook for {title}").strip(),
        "sheets": sheets,
    }


def outline_to_editable_text(doc_type: str, outline: Dict[str, Any]) -> str:
    """Convert structured outline dict into human-friendly editable text for the Cyber-HUD."""
    lines = [f"TITLE: {outline.get('title', 'Untitled')}"]
    if doc_type in ("doc", "ppt"):
        lines.append(f"SUBTITLE: {outline.get('subtitle', '')}")
        items = outline.get("sections" if doc_type == "doc" else "slides", [])
        key = "heading" if doc_type == "doc" else "title"
        for item in items:
            lines.append("")
            lines.append(f"## {item.get(key, 'Section')}")
            if doc_type == "ppt" and item.get("layout"):
                lines.append(f"LAYOUT: {item['layout']}")
            for b in item.get("bullets", []):
                lines.append(f"- {b}")
            if doc_type == "ppt" and item.get("takeaway"):
                lines.append(f"TAKEAWAY: {item['takeaway']}")
            if doc_type == "ppt" and item.get("speaker_note"):
                lines.append(f"NOTE: {item['speaker_note']}")
            if item.get("mermaid"):
                lines.append("MERMAID:")
                for m_ln in str(item["mermaid"]).splitlines():
                    lines.append(f"    {m_ln}")
                lines.append("END_MERMAID")
            if item.get("image_query"):
                lines.append(f"IMAGE: {item['image_query']}")
        return "\n".join(lines).strip()

    lines.append(f"SUMMARY: {outline.get('summary', '')}")
    for sh in outline.get("sheets", []):
        lines.append("")
        lines.append(f"SHEET: {sh.get('sheet_name', 'Sheet1')}")
        lines.append(f"CHART: {sh.get('chart_title', 'Summary Chart')}")
        lines.append("HEADERS: " + " | ".join(str(h) for h in sh.get("headers", [])))
        for row in sh.get("rows", []):
            lines.append("ROW: " + " | ".join(str(c) for c in row))
    return "\n".join(lines).strip()


def _coerce_cell_value(raw: str) -> Any:
    val = raw.strip()
    if re.fullmatch(r"-?\d+", val):
        return int(val)
    if re.fullmatch(r"-?\d+\.\d+", val):
        return float(val)
    return val


def editable_text_to_outline(doc_type: str, text: str, fallback: Dict[str, Any]) -> Dict[str, Any]:
    """Parse user-edited outline text (including MERMAID blocks & TAKEAWAYs) back into dicts."""
    lines = (text or "").splitlines()
    if not lines:
        return fallback

    if doc_type in ("doc", "ppt"):
        title = fallback.get("title", "Untitled")
        subtitle = fallback.get("subtitle", "")
        blocks: List[Dict[str, Any]] = []
        current: Optional[Dict[str, Any]] = None
        key = "heading" if doc_type == "doc" else "title"
        in_mermaid = False
        mermaid_buf: List[str] = []

        def _make_block(heading_text: str) -> Dict[str, Any]:
            blk: Dict[str, Any] = {
                key: heading_text or "Section",
                "bullets": [],
                "mermaid": "",
                "image_query": "",
            }
            if doc_type == "ppt":
                blk["layout"] = "split_visual"
                blk["takeaway"] = ""
                blk["speaker_note"] = ""
            return blk

        for raw_ln in lines:
            s = raw_ln.strip()
            if in_mermaid:
                if s.upper() == "END_MERMAID" or s.startswith("```"):
                    if current is not None:
                        current["mermaid"] = "\n".join(mermaid_buf).strip()
                    in_mermaid = False
                    mermaid_buf = []
                elif s.startswith("##") or s.upper().startswith(("IMAGE:", "NOTE:", "TAKEAWAY:", "LAYOUT:")):
                    # Graceful closure if user forgot END_MERMAID
                    if current is not None:
                        current["mermaid"] = "\n".join(mermaid_buf).strip()
                    in_mermaid = False
                    mermaid_buf = []
                else:
                    # Preserve relative indentation cleanly
                    mermaid_buf.append(raw_ln[4:] if raw_ln.startswith("    ") else s)
                    continue

            if not s:
                continue
            if s.upper().startswith("TITLE:"):
                title = s.split(":", 1)[1].strip() or title
            elif s.upper().startswith("SUBTITLE:"):
                subtitle = s.split(":", 1)[1].strip()
            elif s.startswith("##"):
                if current:
                    blocks.append(current)
                current = _make_block(s.lstrip("#").strip())
            elif s.upper().startswith("MERMAID:") and current is not None:
                inline = s.split(":", 1)[1].strip()
                in_mermaid = True
                mermaid_buf = [inline] if inline and not inline.startswith("```") else []
            elif s.startswith("- ") or s.startswith("* "):
                if current is None:
                    current = _make_block("Overview")
                current["bullets"].append(s[2:].strip())
            elif s.upper().startswith("IMAGE:") and current is not None:
                current["image_query"] = s.split(":", 1)[1].strip()
            elif doc_type == "ppt" and s.upper().startswith("LAYOUT:") and current is not None:
                cand_layout = s.split(":", 1)[1].strip().lower()
                if cand_layout in ("split_visual", "two_column", "diagram_focus"):
                    current["layout"] = cand_layout
            elif doc_type == "ppt" and s.upper().startswith("TAKEAWAY:") and current is not None:
                current["takeaway"] = s.split(":", 1)[1].strip()
            elif doc_type == "ppt" and s.upper().startswith("NOTE:") and current is not None:
                current["speaker_note"] = s.split(":", 1)[1].strip()

        if in_mermaid and current is not None and mermaid_buf:
            current["mermaid"] = "\n".join(mermaid_buf).strip()
        if current:
            blocks.append(current)
        if not blocks:
            return fallback

        field = "sections" if doc_type == "doc" else "slides"
        return {"title": title, "subtitle": subtitle, field: blocks}

    # Excel parser
    title = fallback.get("title", "Workbook")
    summary = fallback.get("summary", "")
    sheets: List[Dict[str, Any]] = []
    cur_sh: Optional[Dict[str, Any]] = None

    for raw_ln in lines:
        s = raw_ln.strip()
        if not s:
            continue
        if s.upper().startswith("TITLE:"):
            title = s.split(":", 1)[1].strip() or title
        elif s.upper().startswith("SUMMARY:"):
            summary = s.split(":", 1)[1].strip()
        elif s.upper().startswith("SHEET:"):
            if cur_sh and cur_sh.get("headers"):
                sheets.append(cur_sh)
            raw_name = s.split(":", 1)[1].strip() or f"Sheet{len(sheets) + 1}"
            cur_sh = {
                "sheet_name": re.sub(r"[:\\/?*\[\]]", "_", raw_name)[:31],
                "headers": [],
                "rows": [],
                "chart_title": f"{title[:24]} Chart",
            }
        elif s.upper().startswith("CHART:") and cur_sh is not None:
            cur_sh["chart_title"] = s.split(":", 1)[1].strip() or cur_sh["chart_title"]
        elif s.upper().startswith("HEADERS:") and cur_sh is not None:
            cur_sh["headers"] = [c.strip() for c in s.split(":", 1)[1].split("|") if c.strip()]
        elif s.upper().startswith("ROW:") and cur_sh is not None:
            cells = [_coerce_cell_value(c) for c in s.split(":", 1)[1].split("|")]
            if cells:
                cur_sh["rows"].append(cells)

    if cur_sh and cur_sh.get("headers"):
        sheets.append(cur_sh)
    if not sheets:
        return fallback
    return {"title": title, "summary": summary, "sheets": sheets}


# ---------------------------------------------------------------------------
# Stage 1: Research + Outline Planning + AI Iteration
# ---------------------------------------------------------------------------

async def prepare_office_outline(
    doc_type_raw: str,
    topic: str,
    drafts_dir: Optional[Path] = None,
    use_llm: bool = True,
    instructions: str = "",
    template_path: Optional[Path] = None,
) -> Dict[str, Any]:
    """
    Stage 1: Discover & scrape sources for topic, inspect any organizer `.pptx`
    template rubric, draft structured outline, and persist draft JSON for HUD review.
    """
    doc_type = normalize_doc_type(doc_type_raw)
    clean_topic = (topic or "").strip()
    if not clean_topic:
        return {"status": "error", "message": "Please provide a topic (e.g., cn ppt Agentic AI)."}

    template_info = inspect_pptx_template(template_path) if doc_type == "ppt" and template_path else {}
    sources = await discover_sources(clean_topic, max_sources=4)
    scraped = await scrape_all_sources(sources, use_cdp_fallback=False)

    model_used = "deterministic-fallback"
    outline_raw: Optional[Dict[str, Any]] = None

    if use_llm:
        prompt = _build_outline_prompt(
            doc_type,
            clean_topic,
            scraped,
            instructions=instructions,
            template_info=template_info,
        )
        res = await query_mentor(
            prompt,
            system_prompt=(
                "You are an expert technical document architect. "
                "Output strictly valid JSON without commentary."
            ),
            max_tokens=2200,
        )
        if res.get("status") == "success" and res.get("model") != "offline-guide":
            parsed = _extract_json_block(res.get("answer", ""))
            if parsed:
                outline_raw = parsed
                model_used = res.get("model", "llm")

    if outline_raw is None:
        outline = _fallback_outline(doc_type, clean_topic, scraped)
    else:
        outline = _normalize_outline_structure(doc_type, outline_raw, clean_topic)

    target_drafts_dir = Path(drafts_dir) if drafts_dir else DRAFTS_DIR
    target_drafts_dir.mkdir(parents=True, exist_ok=True)
    slug = slugify_topic(clean_topic)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    draft_path = target_drafts_dir / f"{doc_type}_{slug}_{stamp}.json"

    draft_payload = {
        "doc_type": doc_type,
        "topic": clean_topic,
        "slug": slug,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "model_used": model_used,
        "instructions": instructions.strip(),
        "template_path": str(template_path) if template_path else "",
        "template_info": template_info,
        "sources": [
            {"title": s.get("title", ""), "url": s.get("url", ""), "source_type": s.get("source_type", "web")}
            for s in scraped
        ],
        "outline": outline,
    }
    draft_path.write_text(json.dumps(draft_payload, indent=2), encoding="utf-8")

    return {
        "status": "outline_ready",
        "doc_type": doc_type,
        "topic": clean_topic,
        "title": outline.get("title", clean_topic),
        "summary_badge": summarize_outline_counts(doc_type, outline, len(scraped)),
        "model_used": model_used,
        "source_count": len(scraped),
        "draft_path": str(draft_path),
        "template_path": draft_payload["template_path"],
        "outline": outline,
    }


async def revise_office_outline(
    draft_path: Path,
    current_outline: Dict[str, Any],
    feedback: str,
    template_path: Optional[Path] = None,
) -> Dict[str, Any]:
    """Iteratively revise the Stage 1 outline using user feedback before building."""
    d_path = Path(draft_path)
    if not d_path.exists():
        return {"status": "error", "message": "Draft file not found."}
    clean_feedback = (feedback or "").strip()
    if not clean_feedback:
        return {"status": "error", "message": "Enter revision feedback first."}

    draft = json.loads(d_path.read_text(encoding="utf-8"))
    doc_type = normalize_doc_type(draft.get("doc_type", "doc"))
    topic = draft.get("topic", "Technical Topic")

    tpl_path_str = str(template_path) if template_path else draft.get("template_path", "")
    tpl_info = inspect_pptx_template(Path(tpl_path_str)) if tpl_path_str and doc_type == "ppt" else {}
    rubric_note = ""
    if tpl_info.get("rubric_hints"):
        rubric_note = "\nTemplate Rubric:\n" + "\n".join(f"- {h}" for h in tpl_info["rubric_hints"])

    prompt = (
        f"Revise the {TYPE_LABELS[doc_type]} JSON outline below according to the user's feedback.\n"
        f"Topic: {topic}{rubric_note}\n\n"
        f"Current JSON Outline:\n{json.dumps(current_outline, indent=2)}\n\n"
        f"User Feedback:\n{clean_feedback}\n\n"
        "Requirements:\n"
        "- Apply ALL requested changes (slide/section count, topics, depth, Mermaid diagrams, wording).\n"
        "- Keep rich `**Lead-in**: detailed explanation` bullets, `takeaway`, and valid `mermaid` blocks where diagrams help.\n"
        "- Output ONLY valid updated JSON matching the exact same schema."
    )

    res = await query_mentor(
        prompt,
        system_prompt="You are an expert technical editor. Return ONLY valid updated JSON.",
        max_tokens=2400,
    )
    if res.get("status") != "success" or res.get("model") == "offline-guide":
        return {"status": "error", "message": res.get("message") or "AI revision unavailable (check API key/quota)."}

    parsed = _extract_json_block(res.get("answer", ""))
    if not parsed:
        return {"status": "error", "message": "AI response was not valid JSON outline — please retry."}

    updated = _normalize_outline_structure(doc_type, parsed, topic)
    draft["outline"] = updated
    draft["model_used"] = res.get("model", draft.get("model_used", "llm"))
    if tpl_path_str:
        draft["template_path"] = tpl_path_str
        draft["template_info"] = tpl_info
    d_path.write_text(json.dumps(draft, indent=2), encoding="utf-8")

    return {
        "status": "revised",
        "model_used": draft["model_used"],
        "outline": updated,
        "summary_badge": summarize_outline_counts(doc_type, updated, len(draft.get("sources", []))),
    }


# ---------------------------------------------------------------------------
# Stage 2A: Mermaid Diagram Renderer + Multi-Source Image Pipeline
# ---------------------------------------------------------------------------

def _is_valid_image_file(img_path: Path, min_w: int = 220, min_h: int = 140) -> bool:
    try:
        with Image.open(img_path) as im:
            w, h = im.size
            if w < min_w or h < min_h:
                return False
            if im.mode not in ("RGB", "RGBA"):
                im.convert("RGB").save(img_path, format="PNG")
        return True
    except Exception:
        return False


def _draw_mermaid_fallback_diagram(mermaid_code: str, title: str, out_path: Path) -> Optional[Path]:
    """
    Offline fallback renderer: parses nodes/edges (`A[Label] --> B[Label]`) from Mermaid
    code and draws a real connected box-and-arrow flowchart via matplotlib.
    """
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        from matplotlib.patches import FancyBboxPatch

        nodes: List[Tuple[str, str]] = []
        seen_ids: set = set()
        edges: List[Tuple[str, str]] = []

        def _reg_node(token: str) -> str:
            t = token.strip()
            m = re.match(r"^([A-Za-z0-9_]+)\s*[\[\(\{]+([^\]\)\}]+)[\]\)\}]+", t)
            if m:
                nid, lbl = m.group(1), m.group(2).strip()
            else:
                nid = re.sub(r"[^A-Za-z0-9_]", "", t.split()[0]) if t else f"N{len(nodes)}"
                lbl = t.replace("_", " ").strip()
            if nid and nid not in seen_ids:
                seen_ids.add(nid)
                nodes.append((nid, lbl[:34]))
            return nid

        for ln in (mermaid_code or "").splitlines():
            clean = ln.strip()
            if not clean or clean.startswith(("flowchart", "graph", "sequenceDiagram", "%%")):
                continue
            arrow_match = re.split(r"\s*(?:-->|==>|-.->|->>|-->>|->)\s*(?:\|[^|]*\|\s*)?", clean)
            if len(arrow_match) >= 2:
                prev = _reg_node(arrow_match[0])
                for nxt in arrow_match[1:]:
                    # Strip sequence ": message"
                    nxt_part = nxt.split(":", 1)[0].strip()
                    curr = _reg_node(nxt_part)
                    if prev and curr:
                        edges.append((prev, curr))
                    prev = curr
            else:
                part_m = re.match(r"^participant\s+([A-Za-z0-9_]+)(?:\s+as\s+(.+))?", clean, re.IGNORECASE)
                if part_m:
                    nid = part_m.group(1)
                    lbl = (part_m.group(2) or nid).strip()
                    if nid not in seen_ids:
                        seen_ids.add(nid)
                        nodes.append((nid, lbl[:30]))

        if not nodes:
            nodes = [("A", "Input Layer"), ("B", "Core Engine"), ("C", "Validation"), ("D", "Output")]
            edges = [("A", "B"), ("B", "C"), ("C", "D")]

        nodes = nodes[:8]
        n_count = len(nodes)
        cols = min(n_count, 3) if n_count > 4 else n_count
        rows = (n_count + cols - 1) // cols

        fig_w = max(8.0, cols * 3.1)
        fig_h = max(4.2, rows * 2.3 + 0.9)
        fig, ax = plt.subplots(figsize=(fig_w, fig_h), dpi=180)
        fig.patch.set_facecolor("#0F172A")
        ax.set_facecolor("#0F172A")
        ax.set_xlim(0, cols * 3.2)
        ax.set_ylim(0, rows * 2.2 + 0.8)
        ax.axis("off")

        ax.text(0.25, rows * 2.2 + 0.35, title[:60], color="#22D3EE", fontsize=13, fontweight="bold", va="center")

        pos_map: Dict[str, Tuple[float, float]] = {}
        box_w, box_h = 2.45, 0.95
        for idx, (nid, lbl) in enumerate(nodes):
            r = idx // cols
            c = idx % cols if r % 2 == 0 else (cols - 1 - (idx % cols))
            cx = 1.6 + c * 3.2
            cy = rows * 2.2 - 0.85 - r * 1.95
            pos_map[nid] = (cx, cy)

            rect = FancyBboxPatch(
                (cx - box_w / 2, cy - box_h / 2),
                box_w,
                box_h,
                boxstyle="round,pad=0.1,rounding_size=0.15",
                facecolor="#1E293B",
                edgecolor="#22D3EE" if idx in (0, len(nodes) - 1) else "#38BDF8",
                linewidth=1.8,
            )
            ax.add_patch(rect)
            wrapped = lbl if len(lbl) <= 18 else (lbl[:17] + "\n" + lbl[17:34])
            ax.text(cx, cy, wrapped, color="#F8FAFC", fontsize=10, fontweight="bold", ha="center", va="center")

        for src_id, dst_id in edges:
            if src_id in pos_map and dst_id in pos_map and src_id != dst_id:
                x1, y1 = pos_map[src_id]
                x2, y2 = pos_map[dst_id]
                dx, dy = x2 - x1, y2 - y1
                if abs(dx) >= abs(dy):
                    sx = x1 + (box_w / 2 + 0.06 if dx > 0 else -(box_w / 2 + 0.06))
                    ex = x2 - (box_w / 2 + 0.06 if dx > 0 else -(box_w / 2 + 0.06))
                    sy, ey = y1, y2
                else:
                    sx, ex = x1, x2
                    sy = y1 + (box_h / 2 + 0.06 if dy > 0 else -(box_h / 2 + 0.06))
                    ey = y2 - (box_h / 2 + 0.06 if dy > 0 else -(box_h / 2 + 0.06))
                ax.annotate(
                    "",
                    xy=(ex, ey),
                    xytext=(sx, sy),
                    arrowprops=dict(arrowstyle="-|>,head_width=0.3,head_length=0.45", color="#22D3EE", lw=2.0),
                )

        out_path.parent.mkdir(parents=True, exist_ok=True)
        fig.savefig(out_path, dpi=180, bbox_inches="tight", facecolor=fig.get_facecolor())
        plt.close(fig)
        return out_path if out_path.exists() else None
    except Exception:
        return None


async def _ensure_mermaid_js(client: httpx.AsyncClient) -> bool:
    """Cache mermaid.min.js locally once so diagram rendering works offline too."""
    if MERMAID_JS_PATH.exists() and MERMAID_JS_PATH.stat().st_size > 100_000:
        return True
    try:
        resp = await client.get(MERMAID_CDN_URL, timeout=15.0)
        if resp.status_code == 200 and len(resp.content) > 100_000:
            VENDOR_DIR.mkdir(parents=True, exist_ok=True)
            MERMAID_JS_PATH.write_bytes(resp.content)
            return True
    except Exception:
        pass
    return False


async def render_mermaid_to_png(
    mermaid_code: str,
    title: str,
    out_path: Path,
    client: Optional[httpx.AsyncClient] = None,
) -> Optional[Path]:
    """
    Render Mermaid syntax into a crisp PNG:
      1. Headless Playwright + local cached mermaid.min.js (highest quality, dark slate theme)
      2. mermaid.ink PNG endpoint fallback (when online)
      3. Local matplotlib connected flowchart fallback (always succeeds offline)
    """
    code = (mermaid_code or "").strip()
    if not code:
        return None

    out_path.parent.mkdir(parents=True, exist_ok=True)
    mmd_path = out_path.with_suffix(".mmd")
    try:
        mmd_path.write_text(code, encoding="utf-8")
    except Exception:
        pass

    own_client = client is None
    http_client = client or httpx.AsyncClient(timeout=12.0, follow_redirects=True)
    try:
        has_js = await _ensure_mermaid_js(http_client)
        if has_js:
            try:
                from playwright.async_api import async_playwright

                html_doc = (
                    "<!doctype html><html><head><meta charset='utf-8'>"
                    "<style>body{margin:0;padding:24px;background:#0F172A;display:inline-block;}"
                    ".mermaid svg{max-width:none !important;height:auto !important;}</style>"
                    "</head><body>"
                    f"<div class='mermaid'>{code}</div>"
                    f"<script>{MERMAID_JS_PATH.read_text(encoding='utf-8')}</script>"
                    "<script>mermaid.initialize({startOnLoad:true,theme:'dark',"
                    "themeVariables:{primaryColor:'#1E293B',primaryTextColor:'#F8FAFC',"
                    "primaryBorderColor:'#22D3EE',lineColor:'#38BDF8',secondaryColor:'#0F172A'}});</script>"
                    "</body></html>"
                )
                async with async_playwright() as pw:
                    try:
                        browser = await pw.chromium.launch(channel="chrome", headless=True)
                    except Exception:
                        browser = await pw.chromium.launch(headless=True)
                    page = await browser.new_page(viewport={"width": 1280, "height": 800}, device_scale_factor=2)
                    await page.set_content(html_doc, wait_until="load")
                    svg = page.locator(".mermaid svg").first
                    await svg.wait_for(state="visible", timeout=5000)
                    await svg.screenshot(path=str(out_path))
                    await browser.close()
                if _is_valid_image_file(out_path, min_w=160, min_h=80):
                    return out_path
            except Exception:
                pass

        # Secondary fallback: mermaid.ink API
        try:
            b64 = base64.urlsafe_b64encode(code.encode("utf-8")).decode("ascii")
            resp = await http_client.get(f"https://mermaid.ink/img/{b64}?theme=dark&bgColor=0F172A", timeout=10.0)
            if resp.status_code == 200 and len(resp.content) > 2048:
                out_path.write_bytes(resp.content)
                if _is_valid_image_file(out_path, min_w=160, min_h=80):
                    return out_path
        except Exception:
            pass
    finally:
        if own_client:
            await http_client.aclose()

    return _draw_mermaid_fallback_diagram(code, title, out_path)


def _generate_diagram_card(title: str, bullets: List[str], out_path: Path) -> Optional[Path]:
    """Clean local visual card fallback when web search & Mermaid are unavailable."""
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        from matplotlib.patches import FancyBboxPatch

        fig, ax = plt.subplots(figsize=(7.6, 5.0), dpi=170)
        fig.patch.set_facecolor("#0F172A")
        ax.set_facecolor("#0F172A")
        ax.set_xlim(0, 10)
        ax.set_ylim(0, 6.4)
        ax.axis("off")

        header_box = FancyBboxPatch(
            (0.4, 5.1), 9.2, 0.9,
            boxstyle="round,pad=0.12",
            facecolor="#1E293B",
            edgecolor="#22D3EE",
            linewidth=1.5,
        )
        ax.add_patch(header_box)
        ax.text(0.75, 5.55, title[:52], color="#F8FAFC", fontsize=13, fontweight="bold", va="center")

        y = 4.2
        for idx, bullet in enumerate(bullets[:4], 1):
            node = FancyBboxPatch(
                (0.55, y - 0.36), 8.9, 0.8,
                boxstyle="round,pad=0.08",
                facecolor="#182334",
                edgecolor="#334155",
                linewidth=1.0,
            )
            ax.add_patch(node)
            ax.text(0.85, y + 0.04, f"0{idx}", color="#22D3EE", fontsize=11, fontweight="bold", va="center")
            clean_b = re.sub(r"\*\*", "", str(bullet))
            ax.text(1.45, y + 0.04, clean_b[:64], color="#E2E8F0", fontsize=10, va="center")
            y -= 1.05

        out_path.parent.mkdir(parents=True, exist_ok=True)
        fig.savefig(out_path, dpi=170, bbox_inches="tight", facecolor=fig.get_facecolor())
        plt.close(fig)
        return out_path if out_path.exists() else None
    except Exception:
        return None


async def _fetch_wikimedia_api_image(query: str, client: httpx.AsyncClient) -> Optional[str]:
    """Query Wikimedia Commons public JSON API with short optimized keywords."""
    if not query:
        return None
    try:
        resp = await client.get(
            "https://commons.wikimedia.org/w/api.php",
            params={
                "action": "query",
                "generator": "search",
                "gsrsearch": query,
                "gsrnamespace": 6,
                "gsrlimit": 8,
                "prop": "imageinfo",
                "iiprop": "url|size|mime",
                "iiurlwidth": 1200,
                "format": "json",
            },
        )
        resp.raise_for_status()
        pages = (resp.json().get("query") or {}).get("pages") or {}
        for page in pages.values():
            infos = page.get("imageinfo") or []
            if not infos:
                continue
            info = infos[0]
            mime = (info.get("mime") or "").lower()
            w = int(info.get("width") or 0)
            h = int(info.get("height") or 0)
            if mime in ("image/jpeg", "image/png") and w >= 480 and h >= 280:
                return info.get("thumburl") or info.get("url")
    except Exception:
        return None
    return None


async def _fetch_cdp_internet_image_url(query: str) -> Optional[str]:
    """Use silent background Chrome CDP tab to search clean internet images (DuckDuckGo Images)."""
    if not query or not await is_cdp_alive(timeout_sec=0.8):
        return None
    page = None
    try:
        search_url = "https://duckduckgo.com/?q=" + quote_plus(query) + "&iax=images&ia=images"
        page, _ = await browser_manager.create_background_tab(search_url)
        await asyncio.sleep(2.2)
        img_url = await page.evaluate(
            """(blockedDomains) => {
                const imgs = Array.from(document.querySelectorAll('img'));
                for (const img of imgs) {
                    const src = img.currentSrc || img.src || '';
                    if (!src.startsWith('http')) continue;
                    const lower = src.toLowerCase();
                    if (lower.includes('duckduckgo.com/assets')) continue;
                    if (blockedDomains.some(d => lower.includes(d))) continue;
                    if (img.naturalWidth >= 360 && img.naturalHeight >= 220) {
                        return src;
                    }
                }
                return null;
            }""",
            list(BLOCKED_IMAGE_DOMAINS),
        )
        return img_url
    except Exception:
        return None
    finally:
        if page:
            try:
                await page.close()
            except Exception:
                pass


async def fetch_visual_assets(
    queries: List[Dict[str, Any]],
    asset_dir: Path,
    use_cdp: bool = True,
) -> Dict[int, Path]:
    """
    Resolve one crisp visual per item index:
      1. If `mermaid` code exists -> render crisp Mermaid diagram PNG
      2. Else optimize `image_query` to 2-3 keywords -> Wikimedia API -> CDP web search
      3. Fallback cleanly to structured local visual card
    """
    asset_dir.mkdir(parents=True, exist_ok=True)
    resolved: Dict[int, Path] = {}
    cdp_up = use_cdp and await is_cdp_alive(timeout_sec=0.8)

    async with httpx.AsyncClient(
        timeout=12.0,
        follow_redirects=True,
        headers={"User-Agent": "SidekickOS/1.0 (https://github.com/sidekick-os)"},
    ) as client:
        for item in queries:
            idx = int(item["index"])
            mermaid_code = (item.get("mermaid") or "").strip()
            raw_query = (item.get("image_query") or "").strip()
            heading = str(item.get("title") or raw_query or "Visual").strip()
            bullets = list(item.get("bullets") or [])
            dest = asset_dir / f"visual_{idx}.png"

            if mermaid_code:
                m_path = await render_mermaid_to_png(mermaid_code, heading, dest, client=client)
                if m_path:
                    resolved[idx] = m_path
                    continue

            opt_query = optimize_image_query(raw_query, heading) if raw_query else ""
            if opt_query:
                for url_getter in (
                    lambda q=opt_query: _fetch_wikimedia_api_image(q, client),
                    lambda q=opt_query: _fetch_cdp_internet_image_url(q) if cdp_up else asyncio.sleep(0, result=None),
                ):
                    remote_url = await url_getter()
                    if not remote_url:
                        continue
                    try:
                        resp = await client.get(remote_url)
                        if resp.status_code == 200 and len(resp.content) > 6144:
                            dest.write_bytes(resp.content)
                            if _is_valid_image_file(dest):
                                resolved[idx] = dest
                                break
                    except Exception:
                        pass
                if idx in resolved:
                    continue

            if opt_query:
                fallback = _generate_diagram_card(heading, bullets, dest)
                if fallback:
                    resolved[idx] = fallback

    return resolved


# ---------------------------------------------------------------------------
# Stage 2B: Format Compilers (.docx, .xlsx, .pptx with Template Preservation)
# ---------------------------------------------------------------------------

def _split_leadin_bullet(raw_bullet: str) -> Tuple[str, str]:
    """Extract optional '**Bold Lead-in**: body' cleanly."""
    clean = str(raw_bullet or "").strip()
    m = re.match(r"^\*\*(.+?)\*\*\s*[:—-]\s*(.+)$", clean)
    if m:
        return (m.group(1).strip(), m.group(2).strip())
    m2 = re.match(r"^([^:]{3,38}):\s+(.+)$", clean)
    if m2:
        return (m2.group(1).strip(), m2.group(2).strip())
    return ("", re.sub(r"\*\*", "", clean))


def _compile_docx(
    outline: Dict[str, Any],
    sources: List[Dict[str, Any]],
    visuals: Dict[int, Path],
    out_path: Path,
) -> Path:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor

    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.85)
        section.right_margin = Inches(0.85)

    title_p = doc.add_paragraph()
    title_run = title_p.add_run(outline.get("title", "Technical Report"))
    title_run.bold = True
    title_run.font.name = "Calibri"
    title_run.font.size = Pt(24)
    title_run.font.color.rgb = RGBColor(15, 23, 42)

    sub_text = outline.get("subtitle") or "Compiled by Chitran"
    sub_p = doc.add_paragraph()
    sub_run = sub_p.add_run(f"{sub_text}  •  {datetime.now().strftime('%d %b %Y')}")
    sub_run.font.name = "Calibri"
    sub_run.font.size = Pt(11)
    sub_run.font.color.rgb = RGBColor(2, 132, 199)

    for idx, sec in enumerate(outline.get("sections", [])):
        h = doc.add_heading(sec.get("heading", "Section"), level=1)
        for run in h.runs:
            run.font.name = "Calibri"
            run.font.size = Pt(15)
            run.font.color.rgb = RGBColor(15, 23, 42)

        for bullet in sec.get("bullets", []):
            lead, body = _split_leadin_bullet(bullet)
            bp = doc.add_paragraph(style="List Bullet")
            if lead:
                r1 = bp.add_run(f"{lead}: ")
                r1.bold = True
                r1.font.name = "Calibri"
                r1.font.size = Pt(11)
                r1.font.color.rgb = RGBColor(15, 23, 42)
            r2 = bp.add_run(body)
            r2.font.name = "Calibri"
            r2.font.size = Pt(11)
            r2.font.color.rgb = RGBColor(30, 41, 59)

        if idx in visuals and visuals[idx].exists():
            pic_p = doc.add_paragraph()
            pic_p.add_run().add_picture(str(visuals[idx]), width=Inches(6.1))
            cap_p = doc.add_paragraph()
            cap_r = cap_p.add_run(f"Figure {idx + 1}: {sec.get('heading', '')}")
            cap_r.italic = True
            cap_r.font.size = Pt(9.5)
            cap_r.font.color.rgb = RGBColor(100, 116, 139)

    if sources:
        doc.add_heading("Verified Sources & References", level=1)
        for s_idx, src in enumerate(sources, 1):
            sp = doc.add_paragraph()
            sr = sp.add_run(f"[{s_idx}] {src.get('title', 'Reference')} — {src.get('url', '')}")
            sr.font.size = Pt(9.5)
            sr.font.color.rgb = RGBColor(71, 85, 105)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return out_path


def _compile_xlsx(
    outline: Dict[str, Any],
    sources: List[Dict[str, Any]],
    out_path: Path,
) -> Path:
    from openpyxl import Workbook
    from openpyxl.chart import BarChart, Reference
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    default_ws = wb.active

    header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    zebra_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    total_fill = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    title_font = Font(name="Calibri", size=14, bold=True, color="0F172A")
    body_font = Font(name="Calibri", size=10, color="1E293B")
    bold_font = Font(name="Calibri", size=10, bold=True, color="0F172A")
    thin_side = Side(border_style="thin", color="CBD5E1")
    cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    for sh in outline.get("sheets", []):
        ws = wb.create_sheet(title=sh.get("sheet_name", "Data")[:31])
        ws["A1"] = outline.get("title", "Chitran Workbook")
        ws["A1"].font = title_font
        ws["A2"] = outline.get("summary", "")
        ws["A2"].font = Font(name="Calibri", size=10, italic=True, color="475569")

        headers = sh.get("headers", ["Item", "Value"])
        rows = sh.get("rows", [])

        header_row_idx = 4
        for col_idx, h_text in enumerate(headers, 1):
            c = ws.cell(row=header_row_idx, column=col_idx, value=str(h_text))
            c.fill = header_fill
            c.font = header_font
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = cell_border

        data_start_row = header_row_idx + 1
        numeric_cols: List[int] = []
        for r_offset, row_vals in enumerate(rows):
            r_num = data_start_row + r_offset
            for c_idx in range(1, len(headers) + 1):
                raw_val = row_vals[c_idx - 1] if c_idx - 1 < len(row_vals) else ""
                c = ws.cell(row=r_num, column=c_idx, value=raw_val)
                c.font = body_font
                c.border = cell_border
                if r_offset % 2 == 1:
                    c.fill = zebra_fill
                if isinstance(raw_val, (int, float)) and c_idx not in numeric_cols:
                    numeric_cols.append(c_idx)

        data_end_row = data_start_row + len(rows) - 1
        if rows and numeric_cols:
            sum_row = data_end_row + 1
            avg_row = data_end_row + 2
            ws.cell(row=sum_row, column=1, value="TOTAL").font = bold_font
            ws.cell(row=avg_row, column=1, value="AVERAGE").font = bold_font
            for c_idx in range(1, len(headers) + 1):
                sc = ws.cell(row=sum_row, column=c_idx)
                ac = ws.cell(row=avg_row, column=c_idx)
                sc.fill = total_fill
                ac.fill = total_fill
                sc.border = cell_border
                ac.border = cell_border
                if c_idx in numeric_cols:
                    col_ltr = get_column_letter(c_idx)
                    sc.value = f"=SUM({col_ltr}{data_start_row}:{col_ltr}{data_end_row})"
                    sc.font = bold_font
                    ac.value = f"=AVERAGE({col_ltr}{data_start_row}:{col_ltr}{data_end_row})"
                    ac.font = bold_font
                    ac.number_format = "0.00"

            chart = BarChart()
            chart.type = "col"
            chart.style = 10
            chart.title = sh.get("chart_title", "Metrics Comparison")
            first_num = min(numeric_cols)
            last_num = max(numeric_cols)
            data_ref = Reference(
                ws,
                min_col=first_num,
                max_col=last_num,
                min_row=header_row_idx,
                max_row=data_end_row,
            )
            cats_ref = Reference(ws, min_col=1, min_row=data_start_row, max_row=data_end_row)
            chart.add_data(data_ref, titles_from_data=True)
            chart.set_categories(cats_ref)
            chart.width = 16
            chart.height = 9
            ws.add_chart(chart, f"{get_column_letter(len(headers) + 2)}{header_row_idx}")

        ws.freeze_panes = f"A{data_start_row}"
        for col_idx in range(1, len(headers) + 1):
            col_ltr = get_column_letter(col_idx)
            max_len = max(
                [len(str(ws.cell(row=r, column=col_idx).value or "")) for r in range(header_row_idx, data_end_row + 1)]
                or [12]
            )
            ws.column_dimensions[col_ltr].width = min(max(max_len + 4, 14), 38)

    if sources:
        src_ws = wb.create_sheet(title="Sources")
        src_ws["A1"] = "Reference Title"
        src_ws["B1"] = "Source URL"
        src_ws["A1"].fill = header_fill
        src_ws["A1"].font = header_font
        src_ws["B1"].fill = header_fill
        src_ws["B1"].font = header_font
        for s_idx, src in enumerate(sources, 2):
            src_ws.cell(row=s_idx, column=1, value=src.get("title", ""))
            src_ws.cell(row=s_idx, column=2, value=src.get("url", ""))
        src_ws.column_dimensions["A"].width = 45
        src_ws.column_dimensions["B"].width = 65

    if default_ws in wb.worksheets and len(wb.worksheets) > 1:
        wb.remove(default_ws)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out_path))
    return out_path


def _clear_template_slides(prs: Any) -> None:
    """Safely remove dummy slides from a loaded .pptx template while keeping masters/backgrounds."""
    sld_id_lst = prs.slides._sldIdLst
    for sld_id in list(sld_id_lst):
        r_id = sld_id.rId
        prs.part.drop_rel(r_id)
        sld_id_lst.remove(sld_id)


def _pick_clean_layout(prs: Any) -> Any:
    """Choose the cleanest slide layout (preferring Blank / fewest placeholders)."""
    for layout in prs.slide_layouts:
        if (layout.name or "").strip().lower() == "blank" or len(layout.placeholders) == 0:
            return layout
    return min(prs.slide_layouts, key=lambda ly: len(ly.placeholders))


def _strip_empty_placeholders(slide: Any) -> None:
    """Remove unused 'Click to add title/text' placeholders cloned from a template layout."""
    for shape in list(slide.placeholders):
        sp = shape._element
        parent = sp.getparent()
        if parent is not None:
            parent.remove(sp)


def _add_fitted_picture(
    slide: Any,
    img_path: Path,
    left_in: float,
    top_in: float,
    max_w_in: float,
    max_h_in: float,
) -> None:
    """Insert picture scaled cleanly to fit within (max_w_in, max_h_in) centered vertically."""
    from pptx.util import Inches

    with Image.open(img_path) as im:
        px_w, px_h = im.size
    ratio = px_w / max(1, px_h)
    w_in = max_w_in
    h_in = w_in / ratio
    if h_in > max_h_in:
        h_in = max_h_in
        w_in = h_in * ratio

    adj_left = left_in + max(0.0, (max_w_in - w_in) / 2.0)
    adj_top = top_in + max(0.0, (max_h_in - h_in) / 2.0)
    slide.shapes.add_picture(str(img_path), Inches(adj_left), Inches(adj_top), width=Inches(w_in), height=Inches(h_in))


def _write_bullets_into_frame(tf: Any, bullets: List[str], font_pt: int, dark_text: bool) -> None:
    from pptx.dml.color import RGBColor
    from pptx.util import Pt

    lead_color = RGBColor(15, 23, 42) if dark_text else RGBColor(248, 250, 252)
    body_color = RGBColor(30, 41, 59) if dark_text else RGBColor(226, 232, 240)
    accent = RGBColor(6, 182, 212)

    tf.word_wrap = True
    for b_idx, raw_b in enumerate(bullets):
        lead, body = _split_leadin_bullet(raw_b)
        p = tf.paragraphs[0] if b_idx == 0 else tf.add_paragraph()
        p.space_after = Pt(12)

        dot = p.add_run()
        dot.text = "▪  "
        dot.font.name = "Calibri"
        dot.font.size = Pt(font_pt)
        dot.font.bold = True
        dot.font.color.rgb = accent

        if lead:
            r_lead = p.add_run()
            r_lead.text = f"{lead}: "
            r_lead.font.name = "Calibri"
            r_lead.font.size = Pt(font_pt)
            r_lead.font.bold = True
            r_lead.font.color.rgb = lead_color

        r_body = p.add_run()
        r_body.text = body
        r_body.font.name = "Calibri"
        r_body.font.size = Pt(font_pt)
        r_body.font.color.rgb = body_color


def _compile_pptx(
    outline: Dict[str, Any],
    sources: List[Dict[str, Any]],
    visuals: Dict[int, Path],
    out_path: Path,
    template_path: Optional[Path] = None,
) -> Path:
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.util import Inches, Pt

    has_template = bool(template_path and Path(template_path).exists() and Path(template_path).suffix.lower() == ".pptx")
    tpl_info = inspect_pptx_template(template_path) if has_template else {}
    dark_master = bool(tpl_info.get("dark_background", False))

    if has_template:
        prs = Presentation(str(template_path))
        _clear_template_slides(prs)
        layout = _pick_clean_layout(prs)
    else:
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        layout = prs.slide_layouts[6]

    sw = prs.slide_width / 914400.0
    sh = prs.slide_height / 914400.0

    # Safe margins keep corner logos/footers in hackathon & college templates intact
    margin_x = 0.72
    top_y = 0.95 if has_template else 0.45
    content_bottom_limit = sh - (0.65 if has_template else 0.45)

    # 1. Cover Slide
    cover = prs.slides.add_slide(layout)
    if has_template:
        _strip_empty_placeholders(cover)
    else:
        bg = cover.background.fill
        bg.solid()
        bg.fore_color.rgb = RGBColor(15, 23, 42)

    accent_bar = cover.shapes.add_shape(1, Inches(0.85), Inches(sh * 0.2), Inches(0.14), Inches(2.25))
    accent_bar.fill.solid()
    accent_bar.fill.fore_color.rgb = RGBColor(34, 211, 238)
    accent_bar.line.fill.background()

    cover_title_color = RGBColor(248, 250, 252) if (not has_template or dark_master) else RGBColor(15, 23, 42)
    cover_sub_color = RGBColor(34, 211, 238) if (not has_template or dark_master) else RGBColor(2, 132, 199)

    title_box = cover.shapes.add_textbox(Inches(1.22), Inches(sh * 0.19), Inches(sw - 2.0), Inches(1.75))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = outline.get("title", "Chitran Presentation")
    p.font.name = "Calibri"
    p.font.size = Pt(38)
    p.font.bold = True
    p.font.color.rgb = cover_title_color

    sub_box = cover.shapes.add_textbox(Inches(1.22), Inches(sh * 0.44), Inches(sw - 2.0), Inches(0.85))
    stf = sub_box.text_frame
    stf.word_wrap = True
    sp = stf.paragraphs[0]
    sp.text = f"{outline.get('subtitle', 'Technical Deck')}  •  {datetime.now().strftime('%d %b %Y')}"
    sp.font.name = "Calibri"
    sp.font.size = Pt(18)
    sp.font.color.rgb = cover_sub_color

    # Agenda pills along lower cover band to eliminate empty cover bottom
    slides_list = outline.get("slides", [])
    if slides_list:
        pill_top = sh * 0.68
        pill_count = min(len(slides_list), 4)
        pill_w = (sw - 2.2 - (pill_count - 1) * 0.22) / pill_count
        for p_idx in range(pill_count):
            px = 1.1 + p_idx * (pill_w + 0.22)
            card = cover.shapes.add_shape(5, Inches(px), Inches(pill_top), Inches(pill_w), Inches(0.95))
            card.fill.solid()
            card.fill.fore_color.rgb = RGBColor(30, 41, 59) if (not has_template or dark_master) else RGBColor(241, 245, 249)
            card.line.color.rgb = RGBColor(34, 211, 238) if not has_template else RGBColor(203, 213, 225)
            ctf = card.text_frame
            ctf.word_wrap = True
            cp = ctf.paragraphs[0]
            cp.text = f"0{p_idx + 1}  {slides_list[p_idx].get('title', '')[:34]}"
            cp.font.name = "Calibri"
            cp.font.size = Pt(12)
            cp.font.bold = True
            cp.font.color.rgb = RGBColor(226, 232, 240) if (not has_template or dark_master) else RGBColor(15, 23, 42)

    # 2. Content Slides
    use_dark_text = not dark_master
    for idx, sl in enumerate(slides_list):
        slide = prs.slides.add_slide(layout)
        if has_template:
            _strip_empty_placeholders(slide)
        else:
            sbg = slide.background.fill
            sbg.solid()
            sbg.fore_color.rgb = RGBColor(248, 250, 252)
            top_strip = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(sw), Inches(0.11))
            top_strip.fill.solid()
            top_strip.fill.fore_color.rgb = RGBColor(6, 182, 212)
            top_strip.line.fill.background()

        # Slide header
        hdr_box = slide.shapes.add_textbox(Inches(margin_x), Inches(top_y), Inches(sw - 2 * margin_x), Inches(0.75))
        htf = hdr_box.text_frame
        htf.word_wrap = True
        hp = htf.paragraphs[0]
        hp.text = sl.get("title", f"Slide {idx + 1}")
        hp.font.name = "Calibri"
        hp.font.size = Pt(26)
        hp.font.bold = True
        hp.font.color.rgb = RGBColor(15, 23, 42) if use_dark_text else RGBColor(248, 250, 252)

        has_img = idx in visuals and visuals[idx].exists()
        takeaway = (sl.get("takeaway") or "").strip()
        body_top = top_y + 0.88
        takeaway_h = 0.56 if takeaway else 0.0
        body_h = max(2.8, content_bottom_limit - body_top - takeaway_h - (0.18 if takeaway else 0.0))
        bullets = sl.get("bullets", [])

        if has_img:
            # Split layout: generous text left + large visual/diagram card right
            left_w = (sw - 2 * margin_x) * 0.44
            right_x = margin_x + left_w + 0.35
            right_w = (sw - margin_x) - right_x

            text_card = slide.shapes.add_shape(5, Inches(margin_x), Inches(body_top), Inches(left_w), Inches(body_h))
            text_card.fill.solid()
            text_card.fill.fore_color.rgb = RGBColor(255, 255, 255) if use_dark_text else RGBColor(30, 41, 59)
            text_card.line.color.rgb = RGBColor(226, 232, 240) if use_dark_text else RGBColor(51, 65, 85)

            body_box = slide.shapes.add_textbox(
                Inches(margin_x + 0.2), Inches(body_top + 0.2), Inches(left_w - 0.4), Inches(body_h - 0.35)
            )
            font_pt = 16 if len(bullets) <= 4 else 14
            _write_bullets_into_frame(body_box.text_frame, bullets, font_pt=font_pt, dark_text=use_dark_text)

            vis_card = slide.shapes.add_shape(5, Inches(right_x), Inches(body_top), Inches(right_w), Inches(body_h))
            vis_card.fill.solid()
            vis_card.fill.fore_color.rgb = RGBColor(15, 23, 42)
            vis_card.line.color.rgb = RGBColor(34, 211, 238)
            _add_fitted_picture(slide, visuals[idx], right_x + 0.14, body_top + 0.14, right_w - 0.28, body_h - 0.28)
        else:
            # Dense 2-column card layout so non-image slides fill canvas symmetrically
            mid = max(1, (len(bullets) + 1) // 2)
            col_left = bullets[:mid]
            col_right = bullets[mid:] or bullets[:1]
            col_w = (sw - 2 * margin_x - 0.32) / 2.0

            for col_i, col_items in enumerate((col_left, col_right)):
                cx = margin_x + col_i * (col_w + 0.32)
                card = slide.shapes.add_shape(5, Inches(cx), Inches(body_top), Inches(col_w), Inches(body_h))
                card.fill.solid()
                card.fill.fore_color.rgb = RGBColor(255, 255, 255) if use_dark_text else RGBColor(30, 41, 59)
                card.line.color.rgb = RGBColor(203, 213, 225) if use_dark_text else RGBColor(51, 65, 85)

                tb = slide.shapes.add_textbox(Inches(cx + 0.22), Inches(body_top + 0.22), Inches(col_w - 0.44), Inches(body_h - 0.4))
                _write_bullets_into_frame(tb.text_frame, col_items, font_pt=17, dark_text=use_dark_text)

        if takeaway:
            ty = content_bottom_limit - takeaway_h
            tbar = slide.shapes.add_shape(5, Inches(margin_x), Inches(ty), Inches(sw - 2 * margin_x), Inches(takeaway_h))
            tbar.fill.solid()
            tbar.fill.fore_color.rgb = RGBColor(224, 242, 254) if use_dark_text else RGBColor(15, 23, 42)
            tbar.line.color.rgb = RGBColor(6, 182, 212)
            ttf = tbar.text_frame
            ttf.word_wrap = True
            tp = ttf.paragraphs[0]
            tp.text = f"KEY TAKEAWAY:  {takeaway}"
            tp.font.name = "Calibri"
            tp.font.size = Pt(13)
            tp.font.bold = True
            tp.font.color.rgb = RGBColor(15, 23, 42) if use_dark_text else RGBColor(34, 211, 238)

        note_text = sl.get("speaker_note") or ""
        if note_text:
            slide.notes_slide.notes_text_frame.text = note_text

    # 3. References Slide if sources exist
    if sources:
        ref_slide = prs.slides.add_slide(layout)
        if has_template:
            _strip_empty_placeholders(ref_slide)
        else:
            rbg = ref_slide.background.fill
            rbg.solid()
            rbg.fore_color.rgb = RGBColor(248, 250, 252)

        rh = ref_slide.shapes.add_textbox(Inches(margin_x), Inches(top_y), Inches(sw - 2 * margin_x), Inches(0.8)).text_frame
        rp = rh.paragraphs[0]
        rp.text = "Verified Sources & References"
        rp.font.size = Pt(24)
        rp.font.bold = True
        rp.font.color.rgb = RGBColor(15, 23, 42) if use_dark_text else RGBColor(248, 250, 252)

        rb = ref_slide.shapes.add_textbox(Inches(margin_x), Inches(top_y + 0.9), Inches(sw - 2 * margin_x), Inches(sh - top_y - 1.6)).text_frame
        rb.word_wrap = True
        for s_idx, src in enumerate(sources, 1):
            sp = rb.paragraphs[0] if s_idx == 1 else rb.add_paragraph()
            sp.text = f"[{s_idx}] {src.get('title', 'Reference')} — {src.get('url', '')}"
            sp.space_after = Pt(10)
            sp.font.size = Pt(14)
            sp.font.color.rgb = RGBColor(71, 85, 105) if use_dark_text else RGBColor(203, 213, 225)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))
    return out_path


async def compile_office_from_draft(
    draft_path: Path,
    edited_outline: Optional[Dict[str, Any]] = None,
    include_visuals: bool = True,
    use_cdp: bool = True,
    output_dir: Optional[Path] = None,
    assets_root: Optional[Path] = None,
    template_path: Optional[Path] = None,
) -> Dict[str, Any]:
    """
    Stage 2: Compile approved draft JSON into .docx/.xlsx/.pptx in output_dir,
    preserving any organizer `.pptx` template background/branding if attached.
    """
    d_path = Path(draft_path)
    if not d_path.exists():
        return {"status": "error", "message": f"Draft file not found: {d_path}"}

    draft = json.loads(d_path.read_text(encoding="utf-8"))
    doc_type = normalize_doc_type(draft.get("doc_type", "doc"))
    outline = edited_outline or draft.get("outline", {})
    sources = draft.get("sources", [])
    slug = draft.get("slug") or slugify_topic(outline.get("title", "output"))
    target_out_dir = Path(output_dir) if output_dir else OUTPUT_DIR
    out_path = resolve_unique_output_path(target_out_dir, slug, EXT_MAP[doc_type])

    resolved_template = Path(template_path) if template_path else (
        Path(draft["template_path"]) if draft.get("template_path") else None
    )

    visuals: Dict[int, Path] = {}
    if include_visuals and doc_type in ("doc", "ppt"):
        items = outline.get("sections" if doc_type == "doc" else "slides", [])
        queries = [
            {
                "index": idx,
                "mermaid": item.get("mermaid", ""),
                "image_query": item.get("image_query", ""),
                "title": item.get("heading" if doc_type == "doc" else "title", ""),
                "bullets": item.get("bullets", []),
            }
            for idx, item in enumerate(items)
            if item.get("mermaid") or item.get("image_query")
        ]
        if queries:
            asset_dir = (Path(assets_root) if assets_root else ASSETS_DIR) / slug
            visuals = await fetch_visual_assets(queries, asset_dir=asset_dir, use_cdp=use_cdp)

    if doc_type == "doc":
        compiled = _compile_docx(outline, sources, visuals, out_path)
    elif doc_type == "ppt":
        compiled = _compile_pptx(outline, sources, visuals, out_path, template_path=resolved_template)
    else:
        compiled = _compile_xlsx(outline, sources, out_path)

    return {
        "status": "success",
        "doc_type": doc_type,
        "file": str(compiled),
        "visual_count": len(visuals),
        "template_used": resolved_template.name if (resolved_template and resolved_template.exists()) else "",
        "message": f"Compiled {compiled.name} -> Chitran_Outputs/",
    }


# ---------------------------------------------------------------------------
# Dark Cyber-HUD: Topic Prompt + Stage 1 Approval & Revision Gate
# ---------------------------------------------------------------------------

def show_office_review_hud(draft_path: Path) -> None:
    """Display editable Stage 1 outline + live count + AI Revision + .pptx template picker."""
    import os
    import subprocess
    import tkinter as tk
    from tkinter import filedialog

    d_path = Path(draft_path)
    if not d_path.exists():
        notify_windows("Chitran Office Error", f"Draft not found: {d_path}")
        return

    draft = json.loads(d_path.read_text(encoding="utf-8"))
    doc_type = normalize_doc_type(draft.get("doc_type", "doc"))
    current_outline: Dict[str, Any] = draft.get("outline", {})
    source_count = len(draft.get("sources", []))
    selected_template: Dict[str, Optional[Path]] = {
        "path": Path(draft["template_path"]) if draft.get("template_path") and Path(draft["template_path"]).exists() else None
    }

    enable_high_dpi()
    root = tk.Tk()
    root.title(f"Chitran — {TYPE_LABELS[doc_type]} Outline Approval")
    root.configure(bg=BG_ROOT)
    root.attributes("-topmost", True)

    width, height = 820, 680
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    root.geometry(f"{width}x{height}+{max(20, (sw - width) // 2)}+{max(20, int((sh - height) * 0.30))}")
    root.minsize(680, 540)
    apply_win11_dwm_style(root)

    outer = tk.Frame(root, bg=BG_ROOT, padx=20, pady=16)
    outer.pack(fill="both", expand=True)

    # Header + Live Count Summary Badge
    hdr = tk.Frame(outer, bg=BG_ROOT)
    hdr.pack(fill="x", pady=(0, 8))
    tk.Label(
        hdr,
        text=f"STAGE 1 OUTLINE REVIEW  •  {TYPE_LABELS[doc_type]}",
        font=("Segoe UI", 10, "bold"),
        fg=ACCENT_CYAN,
        bg=BG_ROOT,
    ).pack(side="left")

    count_lbl = tk.Label(
        hdr,
        text=summarize_outline_counts(doc_type, current_outline, source_count),
        font=("Segoe UI", 9, "bold"),
        fg=ACCENT_EMERALD,
        bg=BG_CARD,
        padx=10,
        pady=3,
    )
    count_lbl.pack(side="right", padx=(8, 0))

    meta_lbl = tk.Label(
        hdr,
        text=f"{source_count} sources  •  {draft.get('model_used', 'local')}",
        font=("Segoe UI", 8),
        fg=TEXT_MUTED,
        bg=BG_CARD,
        padx=8,
        pady=3,
    )
    meta_lbl.pack(side="right")

    # Optional .pptx template bar (for PPT)
    tpl_lbl: Optional[tk.Label] = None
    if doc_type == "ppt":
        tpl_bar = tk.Frame(outer, bg=BG_CARD, padx=10, pady=6)
        tpl_bar.pack(fill="x", pady=(0, 8))

        def _tpl_display() -> str:
            p = selected_template["path"]
            return f"Template: {p.name} (background & master kept)" if p else "Template: Default Chitran Cyber Theme (optional .pptx supported)"

        tpl_lbl = tk.Label(tpl_bar, text=_tpl_display(), font=("Segoe UI", 8, "bold"), fg=ACCENT_AMBER if selected_template["path"] else TEXT_MUTED, bg=BG_CARD)
        tpl_lbl.pack(side="left")

        def _pick_tpl() -> None:
            chosen = filedialog.askopenfilename(
                parent=root,
                title="Select Organizer PowerPoint Template (.pptx)",
                filetypes=[("PowerPoint Template", "*.pptx")],
            )
            if chosen:
                selected_template["path"] = Path(chosen)
                draft["template_path"] = chosen
                d_path.write_text(json.dumps(draft, indent=2), encoding="utf-8")
                if tpl_lbl:
                    tpl_lbl.config(text=_tpl_display(), fg=ACCENT_AMBER)

        def _clear_tpl() -> None:
            selected_template["path"] = None
            draft["template_path"] = ""
            d_path.write_text(json.dumps(draft, indent=2), encoding="utf-8")
            if tpl_lbl:
                tpl_lbl.config(text=_tpl_display(), fg=TEXT_MUTED)

        tk.Button(
            tpl_bar,
            text="Clear",
            font=("Segoe UI", 8),
            fg=TEXT_MUTED,
            bg=BTN_CANCEL_BG,
            relief="flat",
            padx=8,
            pady=2,
            cursor="hand2",
            command=_clear_tpl,
        ).pack(side="right", padx=(6, 0))
        tk.Button(
            tpl_bar,
            text="Attach .pptx Template...",
            font=("Segoe UI", 8, "bold"),
            fg=ACCENT_CYAN,
            bg=BTN_CANCEL_BG,
            relief="flat",
            padx=8,
            pady=2,
            cursor="hand2",
            command=_pick_tpl,
        ).pack(side="right")

    # Iterative "Revise with AI" feedback row
    revise_bar = tk.Frame(outer, bg=BG_ROOT)
    revise_bar.pack(fill="x", pady=(0, 8))

    feedback_var = tk.StringVar()
    feedback_entry = tk.Entry(
        revise_bar,
        textvariable=feedback_var,
        bg=BG_INPUT,
        fg=TEXT_PRIMARY,
        insertbackground=ACCENT_CYAN,
        relief="flat",
        font=("Segoe UI", 9),
        highlightthickness=1,
        highlightbackground=BORDER_SUBTLE,
        highlightcolor=ACCENT_CYAN,
    )
    feedback_entry.pack(side="left", fill="x", expand=True, ipady=5, padx=(0, 8))

    revise_btn = tk.Button(
        revise_bar,
        text="Revise with AI",
        font=("Segoe UI", 9, "bold"),
        fg=BG_ROOT,
        bg=ACCENT_AMBER,
        relief="flat",
        padx=12,
        pady=5,
        cursor="hand2",
    )
    revise_btn.pack(side="right")

    tk.Label(
        outer,
        text="Tip: Type 'Make it 10 slides, add architecture Mermaid diagram' above, OR directly edit headings/bullets/MERMAID below:",
        font=("Segoe UI", 8),
        fg=TEXT_MUTED,
        bg=BG_ROOT,
    ).pack(anchor="w", pady=(0, 6))

    editor_frame = tk.Frame(
        outer,
        bg=BG_CARD,
        highlightbackground=BORDER_SUBTLE,
        highlightthickness=1,
        padx=2,
        pady=2,
    )
    editor_frame.pack(fill="both", expand=True)

    editor = tk.Text(
        editor_frame,
        bg=BG_INPUT,
        fg=TEXT_PRIMARY,
        insertbackground=ACCENT_CYAN,
        font=("Consolas", 10),
        relief="flat",
        padx=12,
        pady=10,
        wrap="word",
        undo=True,
    )
    sb = tk.Scrollbar(editor_frame, command=editor.yview)
    editor.configure(yscrollcommand=sb.set)
    editor.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")
    editor.insert("1.0", outline_to_editable_text(doc_type, current_outline))

    def _refresh_live_count(_e: Any = None) -> None:
        parsed = editable_text_to_outline(doc_type, editor.get("1.0", "end"), current_outline)
        count_lbl.config(text=summarize_outline_counts(doc_type, parsed, source_count))

    editor.bind("<KeyRelease>", _refresh_live_count)

    opts = tk.Frame(outer, bg=BG_ROOT)
    opts.pack(fill="x", pady=(8, 4))
    include_visuals_var = tk.BooleanVar(value=True)
    if doc_type in ("doc", "ppt"):
        tk.Checkbutton(
            opts,
            text="Include Mermaid diagrams & web visuals (optimized 2-3 keyword search + Chrome CDP)",
            variable=include_visuals_var,
            bg=BG_ROOT,
            activebackground=BG_ROOT,
            selectcolor=BG_CARD,
            fg=TEXT_PRIMARY,
            activeforeground=TEXT_PRIMARY,
            font=("Segoe UI", 9),
            relief="flat",
            highlightthickness=0,
        ).pack(side="left")

    status_lbl = tk.Label(
        outer,
        text="Stage 1 paused: zero Office files written until you click Approve & Build.",
        font=("Segoe UI", 9),
        fg=TEXT_MUTED,
        bg=BG_ROOT,
        anchor="w",
    )
    status_lbl.pack(fill="x", pady=(4, 8))

    footer = tk.Frame(outer, bg=BG_ROOT)
    footer.pack(fill="x")

    build_btn = tk.Button(
        footer,
        text=f"Approve & Build {EXT_MAP[doc_type].upper()}",
        font=("Segoe UI", 9, "bold"),
        fg=BG_ROOT,
        bg=ACCENT_CYAN,
        activebackground="#06B6D4",
        relief="flat",
        padx=14,
        pady=7,
        cursor="hand2",
    )
    cancel_btn = tk.Button(
        footer,
        text="Cancel [Esc]",
        font=("Segoe UI", 9),
        fg=TEXT_PRIMARY,
        bg=BTN_CANCEL_BG,
        relief="flat",
        padx=12,
        pady=7,
        cursor="hand2",
        command=root.destroy,
    )
    build_btn.pack(side="right")
    cancel_btn.pack(side="right", padx=(0, 8))

    def _run_revise(_e: Any = None) -> None:
        fb = feedback_var.get().strip()
        if not fb:
            status_lbl.config(text="Type feedback in the amber bar first (e.g. 'Make it 10 slides').", fg=ACCENT_AMBER)
            return
        parsed_now = editable_text_to_outline(doc_type, editor.get("1.0", "end"), current_outline)
        revise_btn.config(state="disabled", text="Revising...")
        build_btn.config(state="disabled")
        status_lbl.config(text="Revising outline with AI based on your instructions...", fg=ACCENT_AMBER)

        def _worker() -> None:
            try:
                res = asyncio.run(
                    revise_office_outline(
                        d_path,
                        parsed_now,
                        fb,
                        template_path=selected_template["path"],
                    )
                )
            except Exception as exc:
                res = {"status": "error", "message": str(exc)}

            def _after() -> None:
                revise_btn.config(state="normal", text="Revise with AI")
                build_btn.config(state="normal")
                if res.get("status") == "revised":
                    nonlocal current_outline
                    current_outline = res["outline"]
                    editor.delete("1.0", "end")
                    editor.insert("1.0", outline_to_editable_text(doc_type, current_outline))
                    count_lbl.config(text=res["summary_badge"])
                    meta_lbl.config(text=f"{source_count} sources  •  {res.get('model_used', 'llm')}")
                    feedback_var.set("")
                    status_lbl.config(text="Outline updated! Review above, revise again, or click Approve & Build.", fg=ACCENT_EMERALD)
                else:
                    status_lbl.config(text=res.get("message", "Revision failed."), fg=ACCENT_RED)

            root.after(0, _after)

        threading.Thread(target=_worker, daemon=True).start()

    revise_btn.config(command=_run_revise)
    feedback_entry.bind("<Return>", _run_revise)

    def _on_complete(res: Dict[str, Any]) -> None:
        if res.get("status") != "success":
            status_lbl.config(text=res.get("message", "Compilation failed."), fg=ACCENT_RED)
            build_btn.config(state="normal", text=f"Retry Build {EXT_MAP[doc_type].upper()}")
            revise_btn.config(state="normal")
            return

        file_path = Path(res["file"])
        notify_windows("Chitran — Office File Ready", f"Saved {file_path.name} to Chitran_Outputs/")
        status_lbl.config(text=f"Ready ({res.get('visual_count', 0)} visuals): {file_path}", fg=ACCENT_EMERALD)
        build_btn.config(
            state="normal",
            text="Open File",
            bg=ACCENT_EMERALD,
            command=lambda: os.startfile(str(file_path)),
        )
        cancel_btn.config(text="Done [Esc]")
        tk.Button(
            footer,
            text="Open Folder",
            font=("Segoe UI", 9),
            fg=TEXT_PRIMARY,
            bg=BG_CARD,
            relief="flat",
            padx=12,
            pady=7,
            cursor="hand2",
            command=lambda: subprocess.Popen(["explorer", "/select,", str(file_path)]),
        ).pack(side="left")

    def _start_build() -> None:
        edited_text = editor.get("1.0", "end")
        updated_outline = editable_text_to_outline(doc_type, edited_text, current_outline)
        build_btn.config(state="disabled", text="Rendering Diagrams & Building...")
        revise_btn.config(state="disabled")
        status_lbl.config(text="Stage 2 running: rendering Mermaid diagrams, fetching images & building file...", fg=ACCENT_CYAN)

        def _worker() -> None:
            try:
                res = asyncio.run(
                    compile_office_from_draft(
                        d_path,
                        edited_outline=updated_outline,
                        include_visuals=include_visuals_var.get(),
                        use_cdp=True,
                        template_path=selected_template["path"],
                    )
                )
            except Exception as exc:
                res = {"status": "error", "message": str(exc)}
            root.after(0, lambda: _on_complete(res))

        threading.Thread(target=_worker, daemon=True).start()

    build_btn.config(command=_start_build)
    root.bind("<Escape>", lambda _e: root.destroy())
    root.mainloop()


def show_office_prompt_hud(default_type: str = "doc") -> None:
    """Quick Cyber-HUD launcher to pick format, enter topic + instructions, and attach optional .pptx template."""
    import tkinter as tk
    from tkinter import filedialog

    doc_type = normalize_doc_type(default_type)
    enable_high_dpi()
    root = tk.Tk()
    root.title("Chitran — Office Generator")
    root.configure(bg=BG_ROOT)
    root.attributes("-topmost", True)
    width, height = 600, 350
    sh = root.winfo_screenheight()
    root.geometry(f"{width}x{height}+24+{max(40, sh - height - 68)}")
    root.resizable(False, False)
    apply_win11_dwm_style(root)

    hdr = tk.Frame(root, bg=BG_CARD, padx=16, pady=10)
    hdr.pack(fill="x")
    tk.Label(
        hdr,
        text="●  OFFICE GENERATOR (2-STAGE APPROVAL)",
        bg=BG_CARD,
        fg=ACCENT_CYAN,
        font=("Segoe UI", 9, "bold"),
        anchor="w",
    ).pack(fill="x")
    tk.Label(
        hdr,
        text="Researches topic -> opens editable Outline HUD -> renders Mermaid/visuals -> builds file.",
        bg=BG_CARD,
        fg=TEXT_MUTED,
        font=("Segoe UI", 9),
        anchor="w",
    ).pack(fill="x", pady=(2, 0))

    body = tk.Frame(root, bg=BG_ROOT, padx=16, pady=12)
    body.pack(fill="both", expand=True)

    type_var = tk.StringVar(value=doc_type)
    fmt_row = tk.Frame(body, bg=BG_ROOT)
    fmt_row.pack(fill="x", pady=(0, 8))
    for code, lbl in (("doc", "Word (.docx)"), ("excel", "Excel (.xlsx)"), ("ppt", "PowerPoint (.pptx)")):
        tk.Radiobutton(
            fmt_row,
            text=lbl,
            variable=type_var,
            value=code,
            bg=BG_ROOT,
            activebackground=BG_ROOT,
            selectcolor=BG_CARD,
            fg=TEXT_PRIMARY,
            activeforeground=ACCENT_CYAN,
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            highlightthickness=0,
        ).pack(side="left", padx=(0, 14))

    tk.Label(body, text="Topic:", bg=BG_ROOT, fg=TEXT_MUTED, font=("Segoe UI", 8, "bold"), anchor="w").pack(fill="x")
    topic_var = tk.StringVar()
    entry = tk.Entry(
        body,
        textvariable=topic_var,
        bg=BG_INPUT,
        fg=TEXT_PRIMARY,
        insertbackground=ACCENT_CYAN,
        relief="flat",
        font=("Segoe UI", 10),
        highlightthickness=1,
        highlightbackground=BORDER_SUBTLE,
        highlightcolor=ACCENT_CYAN,
    )
    entry.pack(fill="x", ipady=5, pady=(2, 8))

    tk.Label(body, text="Instructions / Slide Count (optional, e.g. '10 slides, add cost tradeoffs'):", bg=BG_ROOT, fg=TEXT_MUTED, font=("Segoe UI", 8), anchor="w").pack(fill="x")
    inst_var = tk.StringVar()
    inst_entry = tk.Entry(
        body,
        textvariable=inst_var,
        bg=BG_INPUT,
        fg=TEXT_PRIMARY,
        insertbackground=ACCENT_CYAN,
        relief="flat",
        font=("Segoe UI", 9),
        highlightthickness=1,
        highlightbackground=BORDER_SUBTLE,
        highlightcolor=ACCENT_CYAN,
    )
    inst_entry.pack(fill="x", ipady=4, pady=(2, 8))

    tpl_state: Dict[str, Optional[Path]] = {"path": None}
    tpl_row = tk.Frame(body, bg=BG_ROOT)
    tpl_row.pack(fill="x", pady=(0, 6))
    tpl_status = tk.Label(tpl_row, text="No .pptx template attached (optional)", bg=BG_ROOT, fg=TEXT_MUTED, font=("Segoe UI", 8))
    tpl_status.pack(side="left")

    def _browse_tpl() -> None:
        chosen = filedialog.askopenfilename(
            parent=root,
            title="Select Hackathon/College .pptx Template",
            filetypes=[("PowerPoint File", "*.pptx")],
        )
        if chosen:
            tpl_state["path"] = Path(chosen)
            tpl_status.config(text=f"Template: {Path(chosen).name}", fg=ACCENT_AMBER)

    tk.Button(
        tpl_row,
        text="Attach .pptx Template...",
        bg=BG_CARD,
        fg=ACCENT_CYAN,
        font=("Segoe UI", 8, "bold"),
        relief="flat",
        padx=8,
        pady=3,
        cursor="hand2",
        command=_browse_tpl,
    ).pack(side="right")

    status_lbl = tk.Label(body, text="", bg=BG_ROOT, fg=TEXT_MUTED, font=("Segoe UI", 9), anchor="w")
    status_lbl.pack(fill="x", pady=(2, 0))

    footer = tk.Frame(root, bg=BG_ROOT, padx=16, pady=10)
    footer.pack(fill="x")
    submit_btn = tk.Button(
        footer,
        text="Research & Draft Outline",
        bg=ACCENT_CYAN,
        fg=BG_ROOT,
        font=("Segoe UI", 9, "bold"),
        relief="flat",
        padx=14,
        pady=6,
        cursor="hand2",
    )
    submit_btn.pack(side="right")
    tk.Button(
        footer,
        text="Cancel [Esc]",
        bg=BTN_CANCEL_BG,
        fg=TEXT_PRIMARY,
        font=("Segoe UI", 9),
        relief="flat",
        padx=12,
        pady=6,
        cursor="hand2",
        command=root.destroy,
    ).pack(side="right", padx=(0, 8))

    def _submit(_e: Any = None) -> None:
        topic = topic_var.get().strip()
        if not topic:
            status_lbl.config(text="Please enter a topic first.", fg=ACCENT_RED)
            return
        selected_type = type_var.get()
        submit_btn.config(state="disabled", text="Researching...")
        status_lbl.config(text="Stage 1: Scraping sources & drafting outline...", fg=ACCENT_CYAN)

        def _worker() -> None:
            try:
                res = asyncio.run(
                    prepare_office_outline(
                        selected_type,
                        topic,
                        instructions=inst_var.get().strip(),
                        template_path=tpl_state["path"],
                    )
                )
            except Exception as exc:
                res = {"status": "error", "message": str(exc)}

            def _after() -> None:
                if res.get("status") == "outline_ready" and res.get("draft_path"):
                    draft_file = Path(res["draft_path"])
                    root.destroy()
                    show_office_review_hud(draft_file)
                else:
                    status_lbl.config(text=res.get("message", "Could not draft outline."), fg=ACCENT_RED)
                    submit_btn.config(state="normal", text="Research & Draft Outline")

            root.after(0, _after)

        threading.Thread(target=_worker, daemon=True).start()

    submit_btn.config(command=_submit)
    entry.bind("<Return>", _submit)
    inst_entry.bind("<Return>", _submit)
    root.bind("<Escape>", lambda _e: root.destroy())
    entry.focus_force()
    root.mainloop()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Sidekick OS Human-in-the-Loop Office Generator")
    parser.add_argument("--review", type=str, help="Open Stage 1 outline approval HUD for a draft JSON path")
    parser.add_argument("--prompt", type=str, help="Open quick topic prompt HUD (doc | excel | ppt)")
    parser.add_argument("--build-draft", type=str, help="Directly compile a saved draft JSON path (Stage 2)")
    args = parser.parse_args()

    if args.review:
        show_office_review_hud(Path(args.review))
    elif args.build_draft:
        print(json.dumps(asyncio.run(compile_office_from_draft(Path(args.build_draft))), indent=2))
    else:
        show_office_prompt_hud(args.prompt or "doc")
