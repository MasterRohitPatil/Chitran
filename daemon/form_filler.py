import asyncio
import json
import os
import re
from pathlib import Path
from typing import Dict, Any, List, Optional

from browser import browser_manager

PROJECT_ROOT = Path(__file__).resolve().parent.parent
VAULT_PATH = PROJECT_ROOT / "config" / "profile_vault.json"


def load_profile_vault() -> dict:
    """Loads deterministic personal and academic data from profile_vault.json."""
    if not VAULT_PATH.exists():
        return {}
    with open(VAULT_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def select_resume_path(vault: dict, role_hint: str = "") -> Optional[str]:
    """Selects the best targeted resume PDF based on role hint."""
    resumes = vault.get("resumes", {})
    if not resumes:
        return None

    r = role_hint.lower()
    if "analyst" in r or "analytics" in r:
        return resumes.get("data_analyst")
    elif "ml" in r or "data science" in r or "scientist" in r:
        return resumes.get("data_science_ml")
    elif "genai" in r or "applied ai" in r or "llm" in r:
        return resumes.get("genai_applied_ai")
    elif "fullstack" in r or "dev" in r or "software" in r or "systems" in r:
        return resumes.get("fullstack_systems")
    return resumes.get("master_cv") or resumes.get("data_analyst")


def find_matching_vault_value(label_text: str, vault: dict) -> Optional[str]:
    """Deterministically maps field labels/names to vault values with strict prioritization."""
    text = label_text.lower().strip()
    personal = vault.get("personal", {})
    academics = vault.get("academics", {})
    links = vault.get("links", {})

    # Priority 1: Email (must be before name to prevent name@... matching name)
    if re.search(r"\b(email|e-mail|mail)\b", text):
        return personal.get("email")

    # Priority 2: Phone / Mobile
    if re.search(r"\b(phone|mobile|contact\s*no|contact\s*number|cell|tel)\b", text):
        return personal.get("phone_digits") or personal.get("phone")

    # Priority 3: Academic Institutions & Degrees (before generic name)
    if re.search(r"\b(college|institution|institute|university|school)\b", text):
        return academics.get("college")
    if re.search(r"\b(degree|course)\b", text):
        return academics.get("degree")
    if re.search(r"\b(branch|department|stream|specialization)\b", text):
        return academics.get("branch")
    if re.search(r"\b(cgpa|pointer|gpa|percentage|current\s*cgpa)\b", text):
        return academics.get("cgpa")
    if re.search(r"\b(10th|ssc|matriculation|class\s*10)\b", text):
        return academics.get("ssc_10th_percentage")
    if re.search(r"\b(12th|hsc|intermediate|class\s*12)\b", text):
        return academics.get("hsc_12th_percentage")
    if re.search(r"\b(passing\s*year|year\s*of\s*passing|graduation\s*year|batch)\b", text):
        return str(academics.get("graduation_year", "2027"))
    if re.search(r"\b(backlog|active\s*kt|live\s*kt|kt|arrears)\b", text):
        return academics.get("active_backlogs", "0")

    # Priority 4: Social / Links
    if re.search(r"\b(linkedin)\b", text):
        return links.get("linkedin")
    if re.search(r"\b(github)\b", text):
        return links.get("github")
    if re.search(r"\b(portfolio|website)\b", text):
        return links.get("portfolio")

    # Priority 5: Location
    if re.search(r"\b(city|current\s*city|current\s*location)\b", text):
        return personal.get("city")
    if re.search(r"\b(state)\b", text):
        return personal.get("state")

    # Priority 6: Personal Names (specific before generic)
    if re.search(r"\b(first\s*name|fname)\b", text):
        return personal.get("first_name")
    if re.search(r"\b(last\s*name|surname|lname)\b", text):
        return personal.get("last_name")
    if re.search(r"\b(full\s*name|candidate\s*name|your\s*name)\b", text) or (
        re.search(r"\bname\b", text) and not re.search(r"\b(college|school|company|univ|inst)\b", text)
    ):
        return personal.get("full_name")

    return None


async def auto_fill_active_form(role_hint: str = "") -> Dict[str, Any]:
    """
    Connects to active tab in Chrome, extracts form inputs, and fills them from profile_vault.json.
    STRICT GUARDRAIL: Never clicks the final Submit button.
    """
    vault = load_profile_vault()
    if not vault:
        return {
            "status": "error",
            "message": "profile_vault.json not found in config/ directory."
        }

    try:
        page, cdp_session = await browser_manager.get_target_page(prefer_active=True)
    except Exception as e:
        return {
            "status": "error",
            "message": f"Could not connect to Chrome on port 9222: {e}"
        }

    page_title = await page.title()
    page_url = page.url

    if "chrome://" in page_url:
        return {
            "status": "error",
            "message": "Active tab is a Chrome internal page. Please open a Google Form or application page in Sidekick Chrome."
        }

    filled_fields: List[str] = []

    # 1. Fill standard HTML inputs and textareas
    inputs = await page.query_selector_all("input:not([type='hidden']):not([type='submit']):not([type='button']), textarea")

    for inp in inputs:
        try:
            if not await inp.is_visible() or not await inp.is_enabled():
                continue

            inp_type = await inp.get_attribute("type") or "text"

            if inp_type in ("submit", "button", "checkbox", "radio", "file"):
                continue

            label_parts = []
            aria_label = await inp.get_attribute("aria-label")
            if aria_label:
                label_parts.append(aria_label)

            placeholder = await inp.get_attribute("placeholder")
            if placeholder:
                label_parts.append(placeholder)

            name = await inp.get_attribute("name")
            if name:
                label_parts.append(name)

            elem_id = await inp.get_attribute("id")
            if elem_id:
                label_parts.append(elem_id)

            if elem_id:
                lbl_elem = await page.query_selector(f"label[for='{elem_id}']")
                if lbl_elem:
                    lbl_text = await lbl_elem.inner_text()
                    if lbl_text:
                        label_parts.append(lbl_text)

            combined_label = " ".join(label_parts)
            val = find_matching_vault_value(combined_label, vault)

            if val:
                await inp.fill(str(val))
                filled_fields.append(f"{label_parts[0] if label_parts else 'Input'}: {val}")

        except Exception:
            continue

    # 2. Detect Google Forms listitem fields
    gf_items = await page.query_selector_all("div[role='listitem']")
    for item in gf_items:
        try:
            heading = await item.query_selector("div[role='heading']")
            if not heading:
                continue
            question_text = await heading.inner_text()
            val = find_matching_vault_value(question_text, vault)
            if not val:
                continue

            inp = await item.query_selector("input[type='text'], input[type='email'], textarea")
            if inp and await inp.is_visible() and await inp.is_enabled():
                await inp.fill(str(val))
                filled_fields.append(f"{question_text.splitlines()[0]}: {val}")
        except Exception:
            continue

    # 3. Handle File Uploads (Resumes) if present
    file_inputs = await page.query_selector_all("input[type='file']")
    if file_inputs:
        resume_file = select_resume_path(vault, role_hint)
        if resume_file and os.path.exists(resume_file):
            for finp in file_inputs:
                try:
                    await finp.set_input_files(resume_file)
                    filled_fields.append(f"Resume Attached: {os.path.basename(resume_file)}")
                    break
                except Exception:
                    continue

    return {
        "status": "success",
        "page_title": page_title,
        "url": page_url,
        "filled_count": len(filled_fields),
        "fields": filled_fields,
        "guardrail": "Paused before submission. Please review inputs on your screen."
    }
