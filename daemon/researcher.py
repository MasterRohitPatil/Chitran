"""
Sidekick OS — Step 5.1: Parallel Web Researcher & Claude/ChatGPT Context Bundler (`sk research`)

Zero-cost parallel discovery & scraping pipeline:
  1. Discovers top technical sources concurrently via public developer APIs
     (HackerNews Algolia, StackOverflow, GitHub Repositories, Dev.to, ArXiv)
     plus any direct URLs supplied in the query.
  2. Scrapes and cleans article/README/Q&A text concurrently via `httpx` + `lxml.html`,
     with silent Chrome CDP (`background=True`) fallback for JS-heavy pages when
     Sidekick Chrome (port 9222) is open.
  3. Synthesizes an Executive Technical Brief via Groq / Gemini (`router.query_mentor`,
     enforcing ₹0 quota caps) with a deterministic zero-LLM fallback.
  4. Saves a structured Markdown Context Pack to `~/Daily_Notes/context_packs/` and
     exposes a 1-click "Copy Prompt for Claude/ChatGPT" block.
"""

from __future__ import annotations

import asyncio
import html as html_lib
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import urlparse

import httpx
from lxml import etree, html as lxml_html

from browser import browser_manager, is_cdp_alive
from router import query_mentor


CONTEXT_PACKS_DIR = Path.home() / "Daily_Notes" / "context_packs"
PROMPT_START_MARKER = "<!-- SIDEKICK_PROMPT_START -->"
PROMPT_END_MARKER = "<!-- SIDEKICK_PROMPT_END -->"

REQUEST_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/json;q=0.9,*/*;q=0.8",
}

RESEARCH_SYSTEM_PROMPT = (
    "You are a concise technical research synthesizer preparing engineering notes for "
    "Rohit Ganesh Patil. Extract concrete facts, architecture patterns, pros/cons, and "
    "actionable code/configuration details strictly from the provided sources. "
    "Cite sources inline as [1], [2], etc. Zero fluff, zero filler."
)

TOPIC_STOPWORDS = {
    "the", "and", "for", "with", "how", "what", "why", "when", "best", "ways",
    "using", "into", "from", "guide", "tutorial", "comparison", "overview",
}


def slugify_topic(topic: str) -> str:
    slug = re.sub(r"https?://\S+", "", topic or "").strip().lower()
    slug = re.sub(r"[^a-z0-9]+", "-", slug).strip("-")
    return (slug[:48] or "research-topic")


def extract_primary_keyword(topic: str) -> str:
    tokens = [
        t for t in re.findall(r"[a-z0-9]+", (topic or "").lower())
        if len(t) >= 3 and t not in TOPIC_STOPWORDS
    ]
    return tokens[0] if tokens else "programming"


def clean_html_to_text(raw_html: str, max_chars: int = 2400) -> Tuple[str, str]:
    """Extract `(title, clean_body_text)` from raw HTML using `lxml.html`."""
    if not raw_html or not raw_html.strip():
        return ("", "")
    try:
        doc = lxml_html.fromstring(raw_html)
    except Exception:
        text = re.sub(r"<[^>]+>", " ", raw_html)
        text = re.sub(r"\s+", " ", html_lib.unescape(text)).strip()
        return ("", text[:max_chars])

    title_nodes = doc.xpath("//title/text()")
    title = re.sub(r"\s+", " ", title_nodes[0]).strip() if title_nodes else ""

    # Strip non-content layout chrome
    for bad_node in doc.xpath("//script | //style | //noscript | //nav | //footer | //header | //aside | //svg | //form"):
        parent = bad_node.getparent()
        if parent is not None:
            parent.remove(bad_node)

    blocks: List[str] = []
    seen_lines: set = set()
    for el in doc.xpath("//h1 | //h2 | //h3 | //p | //li | //pre"):
        txt = re.sub(r"\s+", " ", el.text_content() or "").strip()
        if len(txt) < 25:
            continue
        tag = (el.tag or "").lower()
        if tag in ("h1", "h2", "h3"):
            line = f"### {txt}"
        elif tag == "pre":
            line = f"```\n{txt[:450]}\n```"
        else:
            line = txt
        if line not in seen_lines:
            seen_lines.add(line)
            blocks.append(line)
        if sum(len(b) + 2 for b in blocks) >= max_chars:
            break

    if not blocks:
        fallback = re.sub(r"\s+", " ", doc.text_content() or "").strip()
        return (title, fallback[:max_chars])

    joined = "\n\n".join(blocks)
    return (title, joined[:max_chars].strip())


def extract_prompt_from_pack(pack_markdown: str) -> str:
    """Recover the standalone Claude/ChatGPT paste-ready prompt from a saved Context Pack."""
    m = re.search(
        rf"{re.escape(PROMPT_START_MARKER)}\s*(.*?)\s*{re.escape(PROMPT_END_MARKER)}",
        pack_markdown or "",
        re.DOTALL,
    )
    if m:
        prompt = m.group(1).strip()
        # Strip outer markdown fence if wrapped for readability
        fence = re.match(r"^```(?:markdown|md|text)?\s*\n(.*?)\n```$", prompt, re.DOTALL)
        if fence:
            return fence.group(1).strip()
        return prompt
    return (pack_markdown or "").strip()


# ---------------------------------------------------------------------------
# Concurrent Source Discovery (Public Technical APIs + Direct URLs)
# ---------------------------------------------------------------------------

async def _search_hn(topic: str, client: httpx.AsyncClient) -> List[Dict[str, Any]]:
    try:
        r = await client.get(
            "https://hn.algolia.com/api/v1/search",
            params={"query": topic, "tags": "story", "hitsPerPage": 4},
        )
        r.raise_for_status()
        out = []
        for hit in r.json().get("hits", []):
            url = (hit.get("url") or "").strip()
            if not url:
                continue
            out.append({
                "source_type": "hackernews",
                "title": html_lib.unescape(hit.get("title") or url),
                "url": url,
                "seed_text": "",
                "score": int(hit.get("points") or 0),
            })
        return out
    except Exception:
        return []


async def _search_stackoverflow(topic: str, client: httpx.AsyncClient) -> List[Dict[str, Any]]:
    try:
        r = await client.get(
            "https://api.stackexchange.com/2.3/search/advanced",
            params={
                "order": "desc",
                "sort": "relevance",
                "q": topic,
                "site": "stackoverflow",
                "pagesize": 3,
                "filter": "withbody",
            },
        )
        r.raise_for_status()
        out = []
        for item in r.json().get("items", []):
            link = (item.get("link") or "").strip()
            if not link:
                continue
            _, body_txt = clean_html_to_text(item.get("body") or "", max_chars=900)
            out.append({
                "source_type": "stackoverflow",
                "title": html_lib.unescape(item.get("title") or "StackOverflow Thread"),
                "url": link,
                "seed_text": body_txt,
                "score": int(item.get("score") or 0) + 50,
            })
        return out
    except Exception:
        return []


async def _search_github(topic: str, client: httpx.AsyncClient) -> List[Dict[str, Any]]:
    try:
        q = " ".join(re.findall(r"[a-zA-Z0-9_+-]+", topic)[:4]) or topic
        r = await client.get(
            "https://api.github.com/search/repositories",
            params={"q": q, "sort": "stars", "per_page": 2},
            headers={"Accept": "application/vnd.github+json"},
        )
        r.raise_for_status()
        out = []
        for repo in r.json().get("items", []):
            full_name = repo.get("full_name", "")
            branch = repo.get("default_branch", "main")
            html_url = repo.get("html_url", "")
            desc = repo.get("description") or ""
            stars = int(repo.get("stargazers_count") or 0)
            out.append({
                "source_type": "github",
                "title": f"{full_name} ({stars} stars)",
                "url": html_url,
                "raw_readme_url": f"https://raw.githubusercontent.com/{full_name}/{branch}/README.md",
                "seed_text": desc,
                "score": min(120, 40 + stars // 100),
            })
        return out
    except Exception:
        return []


async def _search_devto(topic: str, client: httpx.AsyncClient) -> List[Dict[str, Any]]:
    try:
        tag = extract_primary_keyword(topic)
        r = await client.get(
            "https://dev.to/api/articles",
            params={"per_page": 3, "tag": tag},
        )
        r.raise_for_status()
        out = []
        for art in r.json():
            url = (art.get("url") or "").strip()
            if not url:
                continue
            out.append({
                "source_type": "devto",
                "title": (art.get("title") or url).strip(),
                "url": url,
                "seed_text": (art.get("description") or "").strip(),
                "score": int(art.get("public_reactions_count") or 10) + 25,
            })
        return out
    except Exception:
        return []


async def _search_arxiv(topic: str, client: httpx.AsyncClient) -> List[Dict[str, Any]]:
    try:
        clean_q = " ".join(re.findall(r"[a-zA-Z0-9]+", topic)[:5])
        if not clean_q:
            return []
        r = await client.get(
            "https://export.arxiv.org/api/query",
            params={"search_query": f"all:{clean_q}", "max_results": 2},
        )
        r.raise_for_status()
        root = etree.fromstring(r.content)
        ns = {"atom": "http://www.w3.org/2005/Atom"}
        out = []
        for entry in root.xpath("//atom:entry", namespaces=ns):
            t_el = entry.find("atom:title", ns)
            id_el = entry.find("atom:id", ns)
            s_el = entry.find("atom:summary", ns)
            title = re.sub(r"\s+", " ", (t_el.text if t_el is not None else "")).strip()
            url = (id_el.text if id_el is not None else "").strip()
            summary = re.sub(r"\s+", " ", (s_el.text if s_el is not None else "")).strip()
            if url and title:
                out.append({
                    "source_type": "arxiv",
                    "title": f"[Paper] {title}",
                    "url": url,
                    "seed_text": summary[:1600],
                    "score": 60,
                })
        return out
    except Exception:
        return []


async def discover_sources(topic: str, max_sources: int = 5) -> List[Dict[str, Any]]:
    """Run parallel source discovery across developer & academic APIs plus direct URLs."""
    candidates: List[Dict[str, Any]] = []

    # Direct URLs typed by user go first at top priority
    direct_urls = re.findall(r"https?://[^\s<>\"']+", topic or "")
    for u in direct_urls:
        candidates.append({
            "source_type": "direct",
            "title": u,
            "url": u,
            "seed_text": "",
            "score": 1000,
        })

    clean_topic = re.sub(r"https?://\S+", "", topic or "").strip() or topic
    async with httpx.AsyncClient(timeout=10.0, follow_redirects=True, headers=REQUEST_HEADERS) as client:
        batches = await asyncio.gather(
            _search_hn(clean_topic, client),
            _search_stackoverflow(clean_topic, client),
            _search_github(clean_topic, client),
            _search_devto(clean_topic, client),
            _search_arxiv(clean_topic, client),
            return_exceptions=True,
        )

    for batch in batches:
        if isinstance(batch, list):
            candidates.extend(batch)

    # Deduplicate by normalized URL / host+path cleanly while keeping highest-ranked diversity
    seen_urls: set = set()
    domain_counts: Dict[str, int] = {}
    ranked = sorted(candidates, key=lambda x: x.get("score", 0), reverse=True)
    selected: List[Dict[str, Any]] = []

    for item in ranked:
        url = (item.get("url") or "").strip().rstrip("/")
        if not url or url in seen_urls:
            continue
        host = urlparse(url).netloc.lower()
        if domain_counts.get(host, 0) >= 2 and item.get("source_type") != "direct":
            continue
        seen_urls.add(url)
        domain_counts[host] = domain_counts.get(host, 0) + 1
        selected.append(item)
        if len(selected) >= max_sources:
            break

    return selected


# ---------------------------------------------------------------------------
# Concurrent Page Scraping (HTTP + Silent Background CDP Fallback)
# ---------------------------------------------------------------------------

async def _scrape_via_cdp(url: str, max_chars: int = 2400) -> Tuple[str, str]:
    """Fetch rendered text in a silent background CDP tab if Chrome 9222 is active."""
    page = None
    try:
        page, _ = await browser_manager.create_background_tab(url)
        await asyncio.sleep(1.0)
        title = (await page.title()) or ""
        text = await page.evaluate("() => document.body ? document.body.innerText : ''")
        cleaned = re.sub(r"\s+", " ", text or "").strip()
        return (title.strip(), cleaned[:max_chars])
    except Exception:
        return ("", "")
    finally:
        if page:
            try:
                await page.close()
            except Exception:
                pass


async def scrape_source(
    source: Dict[str, Any],
    client: httpx.AsyncClient,
    cdp_available: bool = False,
) -> Dict[str, Any]:
    """Extract clean readable technical text for one source."""
    url = source["url"]
    stype = source.get("source_type", "web")
    seed_text = (source.get("seed_text") or "").strip()
    title = source.get("title") or url

    # ArXiv abstracts are already clean & complete without HTML scrape overhead
    if stype == "arxiv" and len(seed_text) >= 200:
        return {**source, "content": seed_text}

    # GitHub repositories: pull clean Markdown README.md directly
    fetch_url = source.get("raw_readme_url") or url
    extracted = ""
    page_title = ""

    try:
        resp = await client.get(fetch_url)
        if resp.status_code < 400:
            ctype = (resp.headers.get("content-type") or "").lower()
            if "markdown" in ctype or fetch_url.endswith(".md"):
                cleaned_md = re.sub(r"!\[[^\]]*\]\([^)]+\)", "", resp.text)
                cleaned_md = re.sub(r"\n{3,}", "\n\n", cleaned_md).strip()
                extracted = cleaned_md[:2400]
            else:
                page_title, extracted = clean_html_to_text(resp.text, max_chars=2400)
    except Exception:
        extracted = ""

    # Silent CDP fallback ONLY when HTTP text is thin/JS-rendered and Chrome 9222 is up
    if len(extracted) < 250 and cdp_available and stype != "github":
        cdp_title, cdp_text = await _scrape_via_cdp(url, max_chars=2400)
        if len(cdp_text) > len(extracted):
            extracted = cdp_text
            page_title = cdp_title or page_title

    best_content = extracted if len(extracted) >= len(seed_text) else seed_text
    if seed_text and seed_text not in best_content and len(best_content) < 1600:
        best_content = f"{seed_text}\n\n{best_content}".strip()[:2400]

    final_title = title if (stype != "direct" or not page_title) else page_title
    return {
        **source,
        "title": final_title,
        "content": best_content or "(No body text extracted — see source link above.)",
    }


async def scrape_all_sources(
    sources: List[Dict[str, Any]],
    use_cdp_fallback: bool = True,
) -> List[Dict[str, Any]]:
    if not sources:
        return []
    cdp_available = await is_cdp_alive(timeout_sec=0.8) if use_cdp_fallback else False
    async with httpx.AsyncClient(timeout=10.0, follow_redirects=True, headers=REQUEST_HEADERS) as client:
        results = await asyncio.gather(
            *(scrape_source(src, client, cdp_available=cdp_available) for src in sources),
            return_exceptions=True,
        )
    cleaned: List[Dict[str, Any]] = []
    for src, res in zip(sources, results):
        if isinstance(res, dict):
            cleaned.append(res)
        else:
            cleaned.append({**src, "content": src.get("seed_text") or ""})
    return cleaned


# ---------------------------------------------------------------------------
# Synthesis + Structured Context Pack Writer
# ---------------------------------------------------------------------------

def build_deterministic_synthesis(topic: str, scraped: List[Dict[str, Any]]) -> str:
    if not scraped:
        return (
            f"No external sources could be fetched cleanly for **{topic}** (offline or rate-limited).\n"
            "You can still copy the prompt below into Claude or ChatGPT, or pass a specific URL directly "
            "(e.g., `cn research https://docs.python.org/...`)."
        )
    lines = []
    for idx, src in enumerate(scraped, 1):
        snippet = re.sub(r"\s+", " ", (src.get("content") or "")).strip()[:280]
        lines.append(f"- **[{idx}] {src['title']}** ({src.get('source_type', 'web')}): {snippet}")
    return "\n".join(lines)


def build_paste_prompt(topic: str, synthesis_md: str, scraped: List[Dict[str, Any]]) -> str:
    source_block = []
    for idx, src in enumerate(scraped, 1):
        excerpt = (src.get("content") or "").strip()[:1100]
        source_block.append(
            f"[{idx}] {src['title']}\nURL: {src['url']}\nNotes:\n{excerpt}"
        )
    sources_joined = "\n\n---\n\n".join(source_block) if source_block else "(No scraped web sources available.)"

    return (
        f"I am researching: {topic}\n\n"
        f"Use the verified context pack and source notes below to help me design, compare tradeoffs, "
        f"and implement cleanly with citation numbers [1]..[{max(1, len(scraped))}].\n\n"
        f"### Key Takeaways So Far\n{synthesis_md}\n\n"
        f"### Verified Source Excerpts\n{sources_joined}\n\n"
        f"### What I Need From You\n"
        f"1. Clear architectural breakdown and recommended approach for **{topic}**.\n"
        f"2. Tradeoffs, production failure modes, and security/performance edge cases.\n"
        f"3. Minimal, clean implementation steps or starter code."
    )


def format_context_pack_markdown(
    topic: str,
    synthesis_md: str,
    paste_prompt: str,
    scraped: List[Dict[str, Any]],
    model_used: str,
) -> str:
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    citations = "\n".join(
        f"- **[{idx}]** [{src['title']}]({src['url']}) *({src.get('source_type', 'web')})*"
        for idx, src in enumerate(scraped, 1)
    ) or "- *(No external citations captured)*"

    source_details = []
    for idx, src in enumerate(scraped, 1):
        excerpt = (src.get("content") or "").strip()[:1400]
        source_details.append(f"#### [{idx}] {src['title']}\nURL: {src['url']}\n\n{excerpt}")

    return (
        f"# Context Pack: {topic}\n"
        f"**Generated:** {now_str} · **Sources:** {len(scraped)} · **Synthesizer:** `{model_used}`\n\n"
        f"---\n\n"
        f"## Executive Technical Brief\n"
        f"{synthesis_md}\n\n"
        f"## Citations ({len(scraped)})\n"
        f"{citations}\n\n"
        f"---\n\n"
        f"## Ready-to-Paste Prompt for Claude / ChatGPT\n"
        f"*(Click **Copy Prompt for Claude/ChatGPT** below to copy the full prompt + excerpts.)*\n\n"
        f"{PROMPT_START_MARKER}\n"
        f"{paste_prompt}\n"
        f"{PROMPT_END_MARKER}\n\n"
        f"---\n\n"
        f"## Raw Source Notes\n\n"
        + ("\n\n".join(source_details) if source_details else "_No scraped body text._")
        + "\n"
    )


async def research_topic(
    topic: str,
    max_sources: int = 5,
    use_cdp_fallback: bool = True,
    output_dir: Optional[Path] = None,
) -> Dict[str, Any]:
    """
    Execute parallel web research for `topic`, save Context Pack Markdown under
    `~/Daily_Notes/context_packs/`, and return structured data for the Copy HUD.
    """
    clean_topic = (topic or "").strip()
    if not clean_topic:
        return {"status": "error", "message": "Research topic cannot be empty."}

    discovered = await discover_sources(clean_topic, max_sources=max_sources)
    scraped = await scrape_all_sources(discovered, use_cdp_fallback=use_cdp_fallback)

    fallback_synth = build_deterministic_synthesis(clean_topic, scraped)
    synthesis_md = fallback_synth
    model_used = "deterministic"

    if scraped:
        packed_sources = "\n\n".join(
            f"[{idx}] {s['title']} ({s['url']})\n{(s.get('content') or '')[:1200]}"
            for idx, s in enumerate(scraped, 1)
        )
        synth_prompt = (
            f"Research Topic: {clean_topic}\n\n"
            "Synthesize the technical sources below into:\n"
            "1) **Core Architecture / Concepts** (3-5 crisp bullets with [n] citations)\n"
            "2) **Tradeoffs & Gotchas** (2-4 concrete engineering warnings)\n"
            "3) **Recommended Stack / Pattern** (short actionable verdict)\n\n"
            f"Sources:\n{packed_sources}"
        )
        try:
            llm_res = await query_mentor(synth_prompt, system_prompt=RESEARCH_SYSTEM_PROMPT, max_tokens=1100)
            if llm_res.get("status") == "success" and llm_res.get("model") != "offline-guide":
                synthesis_md = (llm_res.get("answer") or "").strip() or fallback_synth
                model_used = str(llm_res.get("model") or "llm")
        except Exception as e:
            print(f"[Researcher] LLM synthesis fallback ({e})")

    paste_prompt = build_paste_prompt(clean_topic, synthesis_md, scraped)
    pack_md = format_context_pack_markdown(
        topic=clean_topic,
        synthesis_md=synthesis_md,
        paste_prompt=paste_prompt,
        scraped=scraped,
        model_used=model_used,
    )

    target_dir = output_dir or CONTEXT_PACKS_DIR
    target_dir.mkdir(parents=True, exist_ok=True)
    today_str = datetime.now().strftime("%Y-%m-%d")
    slug = slugify_topic(clean_topic)
    out_file = target_dir / f"{today_str}_{slug}.md"
    out_file.write_text(pack_md, encoding="utf-8")

    return {
        "status": "success",
        "topic": clean_topic,
        "file": str(out_file),
        "source_count": len(scraped),
        "model": model_used,
        "copy_prompt": paste_prompt,
        "markdown": pack_md,
    }


if __name__ == "__main__":
    import json
    import sys

    q = " ".join(sys.argv[1:]).strip() or "FastAPI WebSockets vs SSE"
    out = asyncio.run(research_topic(q, use_cdp_fallback=False))
    print(json.dumps({k: v for k, v in out.items() if k not in ("markdown", "copy_prompt")}, indent=2))

