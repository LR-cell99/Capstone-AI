import streamlit as st
import openai
import os
import re
import io
from datetime import date, datetime
from pathlib import Path
import pandas as pd
from dotenv import load_dotenv
from fpdf import FPDF
import pdfplumber
from docx import Document

load_dotenv()
OPENAI_API_KEY  = os.getenv("OPENAI_API_KEY")
SUPABASE_URL    = os.getenv("SUPABASE_URL", "")
SUPABASE_ANON_KEY = os.getenv("SUPABASE_ANON_KEY", "")

# ── Supabase client ────────────────────────────────────────────────────────────
from supabase import create_client, Client as SupabaseClient

def get_supabase() -> SupabaseClient | None:
    if SUPABASE_URL and SUPABASE_ANON_KEY:
        try:
            return create_client(SUPABASE_URL, SUPABASE_ANON_KEY)
        except Exception:
            return None
    return None

def load_tracker_from_supabase() -> pd.DataFrame:
    sb = get_supabase()
    if not sb:
        return pd.DataFrame(columns=["Company", "Role", "Date Applied", "Status", "id"])
    try:
        res = sb.table("applications").select("*").order("created_at", desc=False).execute()
        rows = res.data or []
        if not rows:
            return pd.DataFrame(columns=["Company", "Role", "Date Applied", "Status", "id"])
        df = pd.DataFrame(rows)
        df = df.rename(columns={
            "company":      "Company",
            "role":         "Role",
            "date_applied": "Date Applied",
            "status":       "Status",
        })
        cols = ["Company", "Role", "Date Applied", "Status", "id"]
        return df[[c for c in cols if c in df.columns]]
    except Exception:
        return pd.DataFrame(columns=["Company", "Role", "Date Applied", "Status", "id"])

def insert_to_supabase(company: str, role: str, date_applied: str, status: str) -> str | None:
    """Insert a new row, return the new row id or None on failure."""
    sb = get_supabase()
    if not sb:
        return None
    try:
        res = sb.table("applications").insert({
            "company":      company,
            "role":         role,
            "date_applied": date_applied,
            "status":       status,
        }).execute()
        return res.data[0]["id"] if res.data else None
    except Exception:
        return None

def update_supabase_row(row_id: str, company: str, role: str, date_applied: str, status: str):
    sb = get_supabase()
    if not sb or not row_id:
        return
    try:
        sb.table("applications").update({
            "company":      company,
            "role":         role,
            "date_applied": str(date_applied),
            "status":       status,
        }).eq("id", row_id).execute()
    except Exception:
        pass

def delete_supabase_row(row_id: str):
    sb = get_supabase()
    if not sb or not row_id:
        return
    try:
        sb.table("applications").delete().eq("id", row_id).execute()
    except Exception:
        pass


# ── Text extraction ────────────────────────────────────────────────────────────
def extract_text_from_pdf(file) -> str:
    text = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                text.append(t)
    return "\n".join(text).strip()


def extract_text_from_docx(file) -> str:
    doc = Document(file)
    paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
    return "\n".join(paragraphs).strip()


def extract_text(uploaded_file) -> str:
    name = uploaded_file.name.lower()
    file_bytes = io.BytesIO(uploaded_file.read())
    if name.endswith(".pdf"):
        return extract_text_from_pdf(file_bytes)
    elif name.endswith(".docx"):
        return extract_text_from_docx(file_bytes)
    elif name.endswith(".txt"):
        file_bytes.seek(0)
        return file_bytes.read().decode("utf-8", errors="ignore").strip()
    return ""


# ── Base resume auto-loader ───────────────────────────────────────────────────
BASE_RESUME_DIR = Path("base_resume")
SUPPORTED_EXTS = (".pdf", ".docx", ".txt")

def load_base_resume() -> tuple[str, str]:
    """
    Look for a resume file in the base_resume/ folder.
    Returns (extracted_text, filename) or ("", "") if none found.
    """
    if not BASE_RESUME_DIR.exists():
        return "", ""
    for ext in SUPPORTED_EXTS:
        matches = list(BASE_RESUME_DIR.glob(f"*{ext}"))
        if matches:
            filepath = matches[0]
            with open(filepath, "rb") as f:
                file_bytes = io.BytesIO(f.read())
            if ext == ".pdf":
                text = extract_text_from_pdf(file_bytes)
            elif ext == ".docx":
                text = extract_text_from_docx(file_bytes)
            else:
                file_bytes.seek(0)
                text = file_bytes.read().decode("utf-8", errors="ignore").strip()
            return text, filepath.name
    return "", ""


# ── PDF export ─────────────────────────────────────────────────────────────────
def sanitise_for_pdf(text: str) -> str:
    """Replace Unicode chars unsupported by Helvetica (latin-1 only)."""
    replacements = {
        "\u2013": "-",    # en dash
        "\u2014": "-",    # em dash
        "\u2012": "-",    # figure dash
        "\u2015": "-",    # horizontal bar
        "\u2018": "'",    # left single quote
        "\u2019": "'",    # right single quote
        "\u201c": '"',    # left double quote
        "\u201d": '"',    # right double quote
        "\u2022": "-",    # bullet
        "\u2026": "...",  # ellipsis
        "\u00a0": " ",    # non-breaking space
    }
    for char, replacement in replacements.items():
        text = text.replace(char, replacement)
    return text.encode("latin-1", errors="ignore").decode("latin-1")


def generate_pdf(text: str) -> bytes:
    pdf = FPDF()
    pdf.add_page()
    pdf.set_margins(20, 20, 20)
    pdf.set_auto_page_break(auto=True, margin=20)

    for line in text.split("\n"):
        stripped = sanitise_for_pdf(line.strip())
        if not stripped:
            pdf.ln(3)
            continue
        if stripped.isupper() and len(stripped) > 2:
            pdf.set_font("Helvetica", "B", 13)
            pdf.ln(3)
            pdf.cell(0, 8, stripped, ln=True)
            pdf.set_draw_color(45, 90, 61)
            pdf.set_line_width(0.5)
            pdf.line(20, pdf.get_y(), 190, pdf.get_y())
            pdf.ln(2)
        elif stripped.startswith(("-", "*")):
            pdf.set_font("Helvetica", "", 10)
            content = stripped.lstrip("-* ").strip()
            pdf.set_x(25)
            pdf.cell(5, 6, "-", ln=False)
            pdf.multi_cell(0, 6, content)
        elif len(stripped) < 60 and not stripped.endswith("."):
            pdf.set_font("Helvetica", "B", 10)
            pdf.cell(0, 6, stripped, ln=True)
        else:
            pdf.set_font("Helvetica", "", 10)
            pdf.multi_cell(0, 6, stripped)

    return bytes(pdf.output())


# ── JD helpers ─────────────────────────────────────────────────────────────────
def guess_company(jd: str) -> str:
    m = re.search(r"(?:at|@|about\s+)([A-Z][A-Za-z0-9&\s.,\']+?)(?:\n|,|\.|\s+is\s|\s+are\s)", jd)
    return m.group(1).strip()[:50] if m else ""


def guess_role(jd: str) -> str:
    m = re.search(r"(?:role|position|title|hiring(?:\s+a[n]?)?)[\s:]+([A-Za-z\s\/\-]+?)(?:\n|,|\.|\s+at\s)", jd, re.I)
    return m.group(1).strip()[:60] if m else ""


def clean_jd_text(raw: str) -> str:
    """Strip boilerplate scraped from job portals before sending to the AI."""
    noise_patterns = [
        r"(?i)by continuing to use our platform",
        r"(?i)open app",
        r"(?i)privacy policy",
        r"(?i)cookie.*consent",
        r"(?i)your application will include",
        r"(?i)employer questions",
        r"(?i)registration no\.",
        r"(?i)\bEA No\b",
        r"(?i)company information",
        r"(?i)view all jobs",
        r"(?i)full time.*per month",
        r"(?i)part time.*per month",
        r"(?i)\$[\d,]+\s*[-\u2013]\s*\$[\d,]+\s*per\s*(month|year|hour)",
        r"(?i)posted \d+\w* ago",
        r"(?i)medium application volume",
        r"(?i)high application volume",
        r"(?i)low application volume",
        r"(?i)benefits.*allowance",
        r"(?i)variable bonus",
        r"(?i)apply now",
        r"(?i)quick apply",
        r"(?i)save job",
        r"(?i)report job",
        r"(?i)share this job",
        r"(?i)similar jobs",
        r"(?i)you might also like",
        r"(?i)javascript is disabled",
        r"(?i)(central|north|south|east|west) region",
    ]
    cleaned = []
    for line in raw.split("\n"):
        stripped = line.strip()
        if not stripped:
            cleaned.append("")
            continue
        skip = any(re.search(p, stripped) for p in noise_patterns)
        if not skip:
            cleaned.append(stripped)

    # Collapse 3+ consecutive blank lines to 2
    result = re.sub(r"\n{3,}", "\n\n", "\n".join(cleaned))
    return result.strip()


# ── System prompt ──────────────────────────────────────────────────────────────
SYSTEM_PROMPT = """You are a senior professional resume writer for the Singapore job market. You are editing a candidate's resume to better align it with a specific job description — not rewriting it from scratch.

YOUR GOAL: The enhanced resume must feel like the same person wrote it, but with sharper language, better JD alignment, and stronger professional framing. Preserve the candidate's voice, their actual experience, and the essence of what they have written. Only change what needs changing.

STEP 1 — ANALYSE BEFORE EDITING
a) Read the full resume carefully. Note what is already strong — these sections need minimal or no change.
b) Extract every keyword, skill term, and requirement from the JD.
c) Identify specifically: which bullets are vague or passive? Which JD keywords are missing? Which skills are named differently from the JD?
d) These gaps are your edit targets — focus your effort here.

STEP 2 — TARGETED EDITING (change what needs changing, preserve what is already good)

SUMMARY SECTION:
- Use the original summary as your reference and foundation — do not discard the candidate's intent or voice.
- Rewrite it to be more targeted: weave in the JD's job title, 2-3 key JD skill terms, and the industry context if not already present.
- Keep the same general structure and length. The candidate should recognise their own summary.
- Example: "Motivated engineer looking to grow in manufacturing" → "Motivated Manufacturing Engineer with experience in equipment maintenance and process support, seeking to contribute to semiconductor production environments through hands-on troubleshooting and preventive maintenance."

WORK EXPERIENCE BULLETS:
- For bullets that are already strong and JD-relevant — keep them, minor wording polish only.
- For bullets that are vague, passive, or missing JD alignment — rewrite those specifically:
  * Use the formula: [Strong verb] + [what you did] + [how / with what] + [result or impact]
  * Use the EXACT terminology from the JD where applicable (e.g. "root cause analysis" not "finding problems")
  * Where no result is stated and one can be reasonably inferred from context, add it.
- Strong verbs: Executed, Implemented, Collaborated, Optimised, Reduced, Troubleshot, Calibrated, Validated, Coordinated, Monitored, Analysed, Resolved, Streamlined, Supported, Facilitated.
- Do NOT rewrite every bullet just for the sake of it — only rewrite bullets that genuinely need it.

SKILLS SECTION:
- Reorder so the most JD-critical skills appear first.
- Where a skill is named differently from the JD equivalent, use the JD's exact terminology.
- Only add a skill if the candidate's work experience clearly demonstrates it — not just because the JD lists it.

EDUCATION / CERTIFICATIONS:
- Keep factually intact. No changes unless formatting needs tidying.

HALLUCINATION RULES — non-negotiable:
- NEVER invent a job, company, title, degree, or certification not in the original.
- NEVER add a skill with no basis in the candidate's actual experience.
- NEVER fabricate metrics or achievements not supported by the original resume.
- Reframing, rewording, and reordering existing content is encouraged and is NOT hallucination.

OUTPUT RULES — strictly enforced:
- Output ONLY the resume content. No preamble, commentary, closing note, or explanation of any kind.
- Do NOT write "Here is the enhanced resume", "Note:", "I have updated...", or anything similar.
- Do NOT use markdown formatting or code fences unless they were in the original.
- The very first character must be the candidate's name or contact info.
- The very last character must be the last word of the resume."""


# ── JD extraction prompt ──────────────────────────────────────────────────────
JD_EXTRACT_PROMPT = """You are a strict job description parser. You will receive raw text scraped from a Singapore job portal.
The text is heavily polluted with noise. Your ONLY job is to extract what a hiring manager actually wrote.

WHAT TO EXTRACT (only these, nothing else):
1. Company name — the actual employer. NOT a recruiter agency, NOT an applicant badge.
   Recruiter agencies (e.g. "Inter Island Manpower", "RecruitFirst", "Adecco") are NOT the company.
   Applicant labels (e.g. "Strong applicant", "Top applicant", "Good match") are NOT the company.
   If you are not 100% certain of the real company name, output nothing for this field.

2. Job title — the actual role being hired for.

3. About the company — only if a genuine employer description exists. Skip if absent.

4. Key responsibilities — what the person in this role will actually do day-to-day.

5. Required qualifications and skills — hard requirements stated in the JD.

6. Preferred qualifications — nice-to-haves or bonus skills.

7. Anything else the hiring manager wrote that is directly relevant to performing this role.

WHAT TO STRIP — remove every single one of these, even partial mentions:
- Salary, pay range, allowances, bonuses, CPF, AWS, OT pay, benefits of any kind
- EA licence numbers, UEN/registration numbers, company registration details
- Application instructions ("Click apply", "Send resume to", "Your application will include")
- Employer questions or screening questions
- Cookie consent, privacy policy, platform terms
- Portal UI elements ("Open app", "Save job", "Report job", "Quick apply", "View all jobs")
- Applicant status tags ("Strong applicant", "Medium application volume", "Posted 5h ago")
- Recruiter agency boilerplate or descriptions
- Unrelated job listings or "similar jobs" sections
- Navigation elements, breadcrumbs, region tags ("Central Region", "Full time")
- Any sentence that does not describe the role, the company, or what is required of the candidate

STRICT OUTPUT RULES:
- Output clean plain text using the section labels above.
- If a section has no real content, skip it entirely — do not output the label.
- Do not invent, summarise, or paraphrase beyond what is written in the JD.
- When in doubt whether something belongs — leave it out.
- Your output should read like a clean, standalone job description with zero portal noise."""


# ── ATS scoring prompt ────────────────────────────────────────────────────────
ATS_SCORE_PROMPT = """You are a brutally honest ATS (Applicant Tracking System) evaluator for the Singapore job market.

You will receive a resume and a job description. Score the resume out of 100.

STEP 0 — VALIDATE THE JD FIRST:
Before scoring, check whether the job description contains real hiring content:
- Real hiring content = actual job title, responsibilities, required skills/qualifications written by a hiring manager.
- NOT real hiring content = Wikipedia articles, news articles, generic web pages, "Not applicable" across all fields, cookie notices, portal noise with no actual role described.

If the JD does NOT contain real hiring content:
- Set TOTAL_SCORE to 0
- Set all category scores to 0
- Set SUMMARY to: "Invalid job description — no real hiring content detected. ATS scoring requires an actual job posting with a role title, responsibilities, and requirements."
- Set IMPROVEMENTS to: "Upload a real job description from a job portal or employer website."
- Output in the required format and stop.

SCORING RULES (only apply if JD is valid):
- There is NO minimum score. A poor match can score 20, 30, or 40. Do not anchor to 60.
- 60+ is a PASS meaning the resume genuinely addresses the core role requirements.
- Below 60 is a FAIL — expected for weak or moderate matches.
- Score only what is actually in the resume. Do NOT give benefit of the doubt.
- If a keyword is absent, deduct. If experience is vague, deduct.

SCORING BREAKDOWN (total 100):
- Keyword Match (30 pts):
  * Missing more than half the critical JD keywords = 0-10 pts.
  * Most keywords loosely present = 11-20 pts.
  * Strong keyword alignment including SkillsFuture Singapore terms = 21-30 pts.

- Relevance of Experience (25 pts):
  * Experience in a different field or only loosely related = 0-8 pts.
  * Some relevant experience but significant gaps = 9-16 pts.
  * Experience directly maps to the role responsibilities = 17-25 pts.

- Qualifications Match (20 pts):
  * Missing required degree or certification = deduct 10+ pts immediately.
  * Wrong field of study for a technical role = deduct 8 pts.
  * Meets all requirements = 17-20 pts.

- Resume Clarity & Structure (15 pts):
  * Vague bullets, no action verbs, no results = 0-5 pts.
  * Some strong language but inconsistent = 6-10 pts.
  * Strong action verbs, quantified results, targeted summary throughout = 11-15 pts.

- ATS Formatting Compliance (10 pts):
  * Standard plain text, clean headers, consistent dates = 8-10 pts.
  * Minor issues = 5-7 pts.
  * Tables, columns, or special characters = 0-4 pts.

TOTAL_SCORE must equal the exact sum of all five category scores.

Respond ONLY in this exact format — no extra text before or after:

TOTAL_SCORE: <number>
KEYWORD_MATCH: <number>/30
RELEVANCE_OF_EXPERIENCE: <number>/25
QUALIFICATIONS_MATCH: <number>/20
RESUME_CLARITY: <number>/15
ATS_FORMATTING: <number>/10
SUMMARY: <2-3 sentences — be specific about weaknesses, name actual missing keywords or mismatches>
IMPROVEMENTS: <3 specific actionable bullet points referencing exact missing keywords or JD requirements>"""


# ── Session state ──────────────────────────────────────────────────────────────
if "enhanced_resume" not in st.session_state:
    st.session_state.enhanced_resume = ""
if "jd_text" not in st.session_state:
    st.session_state.jd_text = ""
if "jd_filename" not in st.session_state:
    st.session_state.jd_filename = ""
if "jd_editable_value" not in st.session_state:
    st.session_state.jd_editable_value = ""
if "jd_widget_version" not in st.session_state:
    st.session_state.jd_widget_version = 0
if "ats_result" not in st.session_state:
    st.session_state.ats_result = None
if "baseline_ats_result" not in st.session_state:
    st.session_state.baseline_ats_result = None
if "enhance_count" not in st.session_state:
    st.session_state.enhance_count = 0
if "tracker_edit" not in st.session_state:
    st.session_state.tracker_edit = load_tracker_from_supabase()
if "supabase_ok" not in st.session_state:
    st.session_state.supabase_ok = bool(get_supabase())

# Auto-load base resume once per session
if "resume_text" not in st.session_state or not st.session_state.resume_text:
    _base_text, _base_name = load_base_resume()
    st.session_state.resume_text = _base_text
    st.session_state.base_resume_name = _base_name
else:
    if "base_resume_name" not in st.session_state:
        st.session_state.base_resume_name = ""


# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(page_title="AI Resume Enhancer", page_icon="📄", layout="wide")

st.markdown("""
<style>
    .block-container { padding-top: 2rem; padding-bottom: 3rem; }
    .stTextArea textarea { font-family: 'DM Mono', monospace; font-size: 13px; border-radius: 10px; }
    .stButton > button {
        background-color: #2d5a3d; color: white; border: none;
        border-radius: 10px; padding: 0.5rem 1.5rem; font-weight: 500;
    }
    .stButton > button:hover { background-color: #214a2e; }
    .warn-box {
        border: 1px solid #e6c87a; border-radius: 10px;
        padding: 0.75rem 1rem; font-size: 13px; margin: 1rem 0;
        background: #fdf6e3; color: #7a5a10;
    }
    @media (prefers-color-scheme: dark) {
        .warn-box { background: #2a2200; color: #f0c040; border-color: #7a5a10; }
    }
    .output-box {
        border: 1px solid #e0dbd3; border-radius: 12px; padding: 1.25rem;
        font-family: 'DM Mono', monospace; font-size: 13px;
        line-height: 1.8; white-space: pre-wrap; min-height: 200px;
        background: transparent; color: #bbb5ad;
    }
</style>
""", unsafe_allow_html=True)


# ── Header ─────────────────────────────────────────────────────────────────────
st.markdown("# 📄 AI Resume Enhancer")
st.caption("Upload a JD — your base resume is loaded automatically from the base_resume/ folder.")
st.divider()


# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Settings")
    if OPENAI_API_KEY:
        st.success("✅ OpenAI API key loaded")
    else:
        st.error("❌ OPENAI_API_KEY not found in .env")

    if SUPABASE_URL and SUPABASE_ANON_KEY:
        st.success("✅ Supabase connected")
    else:
        st.warning("⚠️ Supabase not configured — add SUPABASE_URL and SUPABASE_ANON_KEY to .env")

    model = st.selectbox(
        "Model",
        ["gpt-4o", "gpt-4o-mini", "gpt-4-turbo", "gpt-3.5-turbo"],
        index=0,
        help="gpt-4o gives the best results. gpt-4o-mini is faster and cheaper.",
    )

    st.divider()
    output_dir = st.text_input("Save resumes to folder:", value="resume/")

    st.divider()
    st.markdown("""
**Supported file types:**
- Resume: PDF, DOCX, TXT
- JD: PDF, DOCX, TXT

**Tips:**
- Upload the full JD for best keyword matching
- Review the extracted text before enhancing
- Always validate the output before applying
""")


# ── Upload panels ──────────────────────────────────────────────────────────────
col_resume, col_jd = st.columns(2)

with col_resume:
    st.subheader("📎 Base Resume")

    if st.session_state.resume_text:
        name = st.session_state.base_resume_name or "manually entered"
        st.success(f"✅ Loaded **{name}** from `base_resume/` folder ({len(st.session_state.resume_text.split())} words)")
        with st.expander("Preview base resume"):
            st.text(st.session_state.resume_text[:1500] + ("…" if len(st.session_state.resume_text) > 1500 else ""))
        if st.button("🔄 Reload from base_resume/"):
            _base_text, _base_name = load_base_resume()
            if _base_text:
                st.session_state.resume_text = _base_text
                st.session_state.base_resume_name = _base_name
                st.rerun()
            else:
                st.error("No resume file found in base_resume/ folder.")
    else:
        st.warning("⚠️ No resume found in `base_resume/` folder.")
        st.caption("Place your resume (PDF, DOCX, or TXT) in a folder called `base_resume/` next to app.py, then restart the app.")

    with st.expander("Override: upload a different resume"):
        resume_file = st.file_uploader(
            "Upload resume",
            type=["pdf", "docx", "txt"],
            label_visibility="collapsed",
            key="resume_upload",
        )
        if resume_file:
            with st.spinner("Extracting resume text…"):
                extracted = extract_text(resume_file)
            if extracted:
                st.session_state.resume_text = extracted
                st.session_state.base_resume_name = resume_file.name
                st.success(f"✅ Using uploaded file: **{resume_file.name}**")
            else:
                st.error("Could not extract text. Try a different file format.")


with col_jd:
    st.subheader("📋 Job Description")
    jd_file = st.file_uploader(
        "Upload JD",
        type=["pdf", "docx", "txt"],
        label_visibility="collapsed",
        key="jd_upload",
    )

    if jd_file:
        # Detect new file by filename — increment widget version to force fresh widget
        if st.session_state.jd_filename != jd_file.name:
            with st.spinner("Extracting and filtering JD…"):
                extracted_jd = extract_text(jd_file)
            if extracted_jd:
                pre_cleaned = clean_jd_text(extracted_jd)
                try:
                    _client = openai.OpenAI(api_key=OPENAI_API_KEY)
                    _jd_resp = _client.chat.completions.create(
                        model="gpt-4o-mini",
                        messages=[
                            {"role": "system", "content": JD_EXTRACT_PROMPT},
                            {"role": "user", "content": pre_cleaned},
                        ],
                        temperature=0.0,
                        max_tokens=1000,
                    )
                    ai_cleaned_jd = _jd_resp.choices[0].message.content.strip()
                except Exception:
                    ai_cleaned_jd = pre_cleaned

                st.session_state.jd_text = ai_cleaned_jd
                st.session_state.jd_editable_value = ai_cleaned_jd
                st.session_state.jd_filename = jd_file.name
                # Increment version — this changes the widget key, forcing Streamlit
                # to treat it as a completely new widget with the new value
                st.session_state.jd_widget_version += 1
                st.success(f"✅ JD processed from **{jd_file.name}** — showing relevant content only")
            else:
                st.error("Could not extract text. Try a different file format.")

    # ── Editable JD preview — always shown once JD is loaded ──
    if st.session_state.jd_editable_value:
        st.caption("✏️ Edit below to correct company name, job role, or any missing info before enhancing.")
        # Key includes widget version — changes on new upload, forcing widget reset
        widget_key = f"jd_editable_{st.session_state.jd_widget_version}"
        edited_jd = st.text_area(
            "Extracted JD (editable)",
            label_visibility="collapsed",
            value=st.session_state.jd_editable_value,
            height=250,
            key=widget_key,
            help="Edit this directly — corrections here will be used for enhancement and will auto-fill the tracker.",
        )
        st.session_state.jd_text = edited_jd
        st.session_state.jd_editable_value = edited_jd

    with st.expander("Or paste JD manually"):
        manual_jd = st.text_area(
            "JD text",
            label_visibility="collapsed",
            placeholder="Paste the job description here…",
            height=200,
            key="manual_jd",
        )
        if manual_jd.strip():
            st.session_state.jd_text = manual_jd


# ── Enhance button ─────────────────────────────────────────────────────────────
st.write("")
col_btn, col_clear = st.columns([1, 5])

with col_btn:
    enhance_clicked = st.button("⚡ Enhance Resume", use_container_width=True)

with col_clear:
    if st.button("Clear all"):
        st.session_state.enhanced_resume = ""
        st.session_state.resume_text = ""
        st.session_state.jd_text = ""
        st.rerun()


if enhance_clicked:
    if not OPENAI_API_KEY:
        st.error("OPENAI_API_KEY not found. Please add it to your .env file.")
    elif not st.session_state.resume_text.strip():
        st.error("Please upload a resume or paste one manually.")
    elif not st.session_state.jd_text.strip():
        st.error("Please upload a job description or paste one manually.")
    else:
        client = openai.OpenAI(api_key=OPENAI_API_KEY)

        with st.spinner("Step 1/3 — Scoring baseline resume against JD…"):
            try:
                baseline_ats_resp = client.chat.completions.create(
                    model="gpt-4o-mini",
                    messages=[
                        {"role": "system", "content": ATS_SCORE_PROMPT},
                        {"role": "user", "content": (
                            f"## Resume (Original / Baseline)\n\n{st.session_state.resume_text}"
                            f"\n\n---\n\n## Job Description\n\n{st.session_state.jd_text}"
                        )},
                    ],
                    temperature=0.1,
                    max_tokens=600,
                )
                st.session_state.baseline_ats_result = baseline_ats_resp.choices[0].message.content.strip()
            except Exception:
                st.session_state.baseline_ats_result = None

        with st.spinner("Step 2/3 — Enhancing your resume…"):
            try:
                def similarity_ratio(a: str, b: str) -> float:
                    """Rough word-overlap similarity between two texts."""
                    words_a = set(a.lower().split())
                    words_b = set(b.lower().split())
                    if not words_a or not words_b:
                        return 1.0
                    return len(words_a & words_b) / max(len(words_a), len(words_b))

                user_content = (
                    f"## Candidate Resume\n\n{st.session_state.resume_text}"
                    f"\n\n---\n\n## Job Description\n\n{st.session_state.jd_text}"
                )

                # Try up to 2 times — retry if output is too similar to baseline
                enhanced = None
                for attempt in range(2):
                    temp = 0.7 if attempt == 0 else 0.9
                    response = client.chat.completions.create(
                        model=model,
                        messages=[
                            {"role": "system", "content": SYSTEM_PROMPT},
                            {"role": "user", "content": user_content},
                        ],
                        temperature=temp,
                        max_tokens=3500,
                    )
                    candidate = response.choices[0].message.content.strip()
                    sim = similarity_ratio(st.session_state.resume_text, candidate)
                    # Accept if similarity is below 80% or it's the last attempt
                    if sim < 0.80 or attempt == 1:
                        enhanced = candidate
                        if sim >= 0.80:
                            st.warning(f"⚠️ Enhancement is very similar to your baseline (similarity: {sim:.0%}). This may mean the JD has too little content or your resume is already well-matched.")
                        break

                st.session_state.enhanced_resume = enhanced

                Path(output_dir).mkdir(parents=True, exist_ok=True)
                timestamp = datetime.now().strftime("%Y-%m-%d_%H%M")
                filepath = Path(output_dir) / f"enhanced_{timestamp}.txt"
                filepath.write_text(enhanced, encoding="utf-8")
                st.success(f"✅ Saved to `{filepath}`")

                # ── Auto-log to tracker ──
                st.session_state.enhance_count += 1

                # Read from the editable JD widget — use versioned key to get latest edits
                _widget_key = f"jd_editable_{st.session_state.jd_widget_version}"
                _jd_source = st.session_state.get(_widget_key, st.session_state.jd_text)

                def _extract_field(text, patterns):
                    for _line in text.split("\n"):
                        _line = _line.strip()
                        for pat, sub_pat in patterns:
                            if re.match(pat, _line):
                                _val = re.sub(sub_pat, "", _line).strip()
                                if _val and _val.lower() not in ("na", "not applicable", ""):
                                    return _val
                    return "NA"

                _company = _extract_field(_jd_source, [
                    (r"(?i)^(company name|company)\s*[:\-]",
                     r"(?i)^(company name|company)\s*[:\-]\s*"),
                ])
                _role = _extract_field(_jd_source, [
                    (r"(?i)^(job title|role|position)\s*[:\-]",
                     r"(?i)^(job title|role|position)\s*[:\-]\s*"),
                ])

                # Stage row in session only — user confirms via the form before Supabase save
                auto_row = pd.DataFrame([{
                    "Company":      _company,
                    "Role":         _role,
                    "Date Applied": str(date.today()),
                    "Status":       "Not Applied",
                    "id":           "",  # No id yet — assigned on Supabase confirm
                }])
                st.session_state.tracker_edit = pd.concat(
                    [st.session_state.tracker_edit, auto_row], ignore_index=True
                )

                # ── Step 3: ATS Score enhanced resume ──
                with st.spinner("Step 3/3 — Scoring enhanced resume…"):
                    try:
                        ats_response = client.chat.completions.create(
                            model="gpt-4o-mini",
                            messages=[
                                {"role": "system", "content": ATS_SCORE_PROMPT},
                                {"role": "user", "content": (
                                    f"## Enhanced Resume\n\n{enhanced}"
                                    f"\n\n---\n\n## Job Description\n\n{st.session_state.jd_text}"
                                )},
                            ],
                            temperature=0.1,
                            max_tokens=600,
                        )
                        st.session_state.ats_result = ats_response.choices[0].message.content.strip()
                    except Exception:
                        st.session_state.ats_result = None

            except openai.AuthenticationError:
                st.error("Invalid API key. Please check OPENAI_API_KEY in your .env file.")
            except openai.RateLimitError:
                st.error("Rate limit hit. Please wait a moment and try again.")
            except Exception as e:
                st.error(f"Error: {e}")


# ── Output ─────────────────────────────────────────────────────────────────────
st.divider()
st.subheader("Enhanced Resume Output")
st.caption("Edit the resume directly below before exporting.")

if st.session_state.enhanced_resume:
    # Seed the editable copy once when AI output first arrives,
    # then let the user's edits persist independently.
    if "edited_resume" not in st.session_state or             st.session_state.get("_last_enhanced") != st.session_state.enhanced_resume:
        st.session_state.edited_resume = st.session_state.enhanced_resume
        st.session_state._last_enhanced = st.session_state.enhanced_resume

    edited = st.text_area(
        label="Edit resume",
        label_visibility="collapsed",
        value=st.session_state.edited_resume,
        height=500,
        key="editable_output",
        help="You can edit this text directly. Your changes will be used when exporting.",
    )
    # Keep session state in sync with whatever the user typed
    st.session_state.edited_resume = edited

    col_dl1, col_dl2, col_reset, col_rest = st.columns([1, 1, 1, 3])
    with col_dl1:
        st.download_button(
            label="⬇ Download .txt",
            data=edited,
            file_name=f"enhanced_resume_{date.today().isoformat()}.txt",
            mime="text/plain",
            use_container_width=True,
        )
    with col_dl2:
        pdf_bytes = generate_pdf(edited)
        st.download_button(
            label="⬇ Download .pdf",
            data=pdf_bytes,
            file_name=f"enhanced_resume_{date.today().isoformat()}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
    with col_reset:
        if st.button("↺ Reset edits", use_container_width=True, help="Revert to the original AI output"):
            st.session_state.edited_resume = st.session_state.enhanced_resume
            st.rerun()
else:
    st.markdown(
        '<div class="output-box">Your enhanced resume will appear here after clicking Enhance Resume.</div>',
        unsafe_allow_html=True,
    )

# ── ATS Score display ─────────────────────────────────────────────────────────
def parse_ats(raw: str) -> dict:
    def pf(label):
        m = re.search(rf"{label}:\s*([\d.]+)", raw)
        return float(m.group(1)) if m else None
    m_sum = re.search(r"SUMMARY:\s*(.+?)(?=IMPROVEMENTS:|$)", raw, re.S)
    m_imp = re.search(r"IMPROVEMENTS:\s*(.+?)$", raw, re.S)
    return {
        "total":       pf("TOTAL_SCORE"),
        "kw":          pf("KEYWORD_MATCH"),
        "exp":         pf("RELEVANCE_OF_EXPERIENCE"),
        "qual":        pf("QUALIFICATIONS_MATCH"),
        "clarity":     pf("RESUME_CLARITY"),
        "formatting":  pf("ATS_FORMATTING"),
        "summary":     m_sum.group(1).strip() if m_sum else "",
        "improve":     m_imp.group(1).strip() if m_imp else "",
    }

def score_badge(total, label_text):
    passed = total >= 60
    colour = "#2d5a3d" if passed else "#c0392b"
    bg     = "#e8f2eb" if passed else "#fde8e8"
    badge  = "✅ PASS" if passed else "❌ BELOW PASSING"
    return f"""
    <div style="text-align:center; background:{bg}; border:2px solid {colour};
                border-radius:12px; padding:12px 20px; min-width:130px;">
        <div style="font-size:11px; color:{colour}; font-weight:600; margin-bottom:4px;">{label_text}</div>
        <div style="font-size:2.2rem; font-weight:800; color:{colour}; line-height:1;">{int(total)}</div>
        <div style="font-size:11px; color:{colour};">/ 100</div>
        <div style="font-size:11px; font-weight:700; color:{colour}; margin-top:4px;">{badge}</div>
    </div>"""

def score_bars(scores, baseline_scores=None):
    breakdown = [
        ("Keyword Match",           "kw",        30),
        ("Relevance of Experience", "exp",       25),
        ("Qualifications Match",    "qual",      20),
        ("Resume Clarity",          "clarity",   15),
        ("ATS Formatting",          "formatting",10),
    ]
    html = ""
    for name, key, max_score in breakdown:
        score = scores.get(key)
        if score is None:
            continue
        pct = int((score / max_score) * 100)
        bar_colour = "#2d5a3d" if pct >= 60 else "#e67e22" if pct >= 40 else "#c0392b"
        delta_html = ""
        if baseline_scores:
            b = baseline_scores.get(key)
            if b is not None:
                diff = score - b
                arrow = "▲" if diff > 0 else ("▼" if diff < 0 else "–")
                d_colour = "#2d5a3d" if diff > 0 else ("#c0392b" if diff < 0 else "#888")
                delta_html = f'<span style="color:{d_colour}; font-size:12px; margin-left:8px;">{arrow} {abs(diff):.0f}</span>'
        html += f"""
        <div style="margin-bottom:12px;">
            <div style="display:flex; justify-content:space-between; font-size:13px; margin-bottom:3px;">
                <span>{name}</span>
                <span style="font-weight:600;">{int(score)} / {max_score}{delta_html}</span>
            </div>
            <div style="background:#e0dbd3; border-radius:6px; height:10px;">
                <div style="width:{pct}%; background:{bar_colour}; border-radius:6px; height:10px;"></div>
            </div>
        </div>"""
    return html

if st.session_state.ats_result:
    st.divider()
    st.subheader("🎯 ATS Score")
    st.caption("Passing score is 60 / 100 — based on Singapore ATS standards and SkillsFuture keyword alignment.")

    enhanced_scores = parse_ats(st.session_state.ats_result)
    baseline_scores = parse_ats(st.session_state.baseline_ats_result) if st.session_state.baseline_ats_result else None

    # ── Score badges ──
    if enhanced_scores["total"] is not None:
        if baseline_scores and baseline_scores["total"] is not None:
            b_total = baseline_scores["total"]
            e_total = enhanced_scores["total"]
            improvement = e_total - b_total
            imp_colour = "#2d5a3d" if improvement > 0 else "#c0392b"
            imp_sign   = "+" if improvement > 0 else ""

            col_b, col_e, col_delta = st.columns([1, 1, 1])
            with col_b:
                st.markdown(score_badge(b_total, "BASELINE RESUME"), unsafe_allow_html=True)
            with col_e:
                st.markdown(score_badge(e_total, "ENHANCED RESUME"), unsafe_allow_html=True)
            with col_delta:
                st.markdown(f"""
                <div style="text-align:center; border:2px solid {imp_colour}; border-radius:12px;
                            padding:12px 20px; min-width:130px;">
                    <div style="font-size:11px; color:{imp_colour}; font-weight:600; margin-bottom:4px;">IMPROVEMENT</div>
                    <div style="font-size:2.2rem; font-weight:800; color:{imp_colour}; line-height:1;">{imp_sign}{improvement:.0f}</div>
                    <div style="font-size:11px; color:{imp_colour};">points</div>
                </div>""", unsafe_allow_html=True)
        else:
            st.markdown(score_badge(enhanced_scores["total"], "ENHANCED RESUME"), unsafe_allow_html=True)

        # ── Breakdown in expander ──
        st.write("")
        with st.expander("📊 View score breakdown" + (" *(▲/▼ vs baseline)*" if baseline_scores else ""), expanded=False):
            st.markdown(score_bars(enhanced_scores, baseline_scores), unsafe_allow_html=True)

            if enhanced_scores["summary"]:
                st.markdown("**Summary**")
                st.info(enhanced_scores["summary"])

            if enhanced_scores["improve"]:
                st.markdown("**How to improve your score**")
                for line in enhanced_scores["improve"].split("\n"):
                    line = line.strip().lstrip("-•* ").strip()
                    if line:
                        st.markdown(f"- {line}")

st.markdown(
    '<div class="warn-box">⚠️ <strong>Always review the enhanced resume carefully.</strong> '
    'AI may occasionally embellish skills or experience — validate every change before submitting applications.</div>',
    unsafe_allow_html=True,
)

# ── Debug panel (only shown after enhancement has run) ────────────────────────
if st.session_state.enhanced_resume:
 with st.expander("🔍 Debug — verify ATS inputs & raw scores", expanded=False):
     st.markdown("**What was sent to baseline ATS scorer:**")
     if st.session_state.get("resume_text"):
         st.text(st.session_state.resume_text[:800] + ("…" if len(st.session_state.resume_text) > 800 else ""))
     else:
         st.caption("No baseline resume in session.")

     st.markdown("**What was sent to enhanced ATS scorer:**")
     if st.session_state.get("enhanced_resume"):
         st.text(st.session_state.enhanced_resume[:800] + ("…" if len(st.session_state.enhanced_resume) > 800 else ""))
     else:
         st.caption("No enhanced resume in session.")

     st.markdown("**Raw baseline ATS response:**")
     st.code(st.session_state.baseline_ats_result or "None", language="text")

     st.markdown("**Raw enhanced ATS response:**")
     st.code(st.session_state.ats_result or "None", language="text")

     st.markdown("**Are baseline and enhanced resume identical?**")
     same = st.session_state.get("resume_text","") == st.session_state.get("enhanced_resume","")
     if same:
         st.error("⚠️ YES — the enhanced resume is identical to the baseline. The enhancer may not have run correctly.")
     else:
         st.success("✅ No — the enhanced resume differs from the baseline.")


# ── Application Tracker ────────────────────────────────────────────────────────
st.divider()
st.subheader("📋 Application Tracker")
st.caption("Track every application. Edit directly in the table below.")

STATUS_OPTIONS = ["Not Applied", "Pending", "Applied", "Interview", "Offer", "Rejected"]

# ── Quick-add row form ──
# Pre-fill from latest tracker row (most recent auto-logged entry)
def get_latest_tracker_defaults() -> dict:
    """Pull the most recent row from the in-session tracker as form defaults."""
    df = st.session_state.tracker_edit
    if df.empty:
        # Fall back to JD text if tracker has no rows yet
        _company, _role = "NA", "NA"
        if st.session_state.jd_text:
            for _line in st.session_state.jd_text.split("\n"):
                _line = _line.strip()
                if re.match(r"(?i)^(company name|company)\s*[:\-]", _line) and _company == "NA":
                    _v = re.sub(r"(?i)^(company name|company)\s*[:\-]\s*", "", _line).strip()
                    if _v and _v.lower() not in ("na", "not applicable", ""):
                        _company = _v
                if re.match(r"(?i)^(job title|role|position)\s*[:\-]", _line) and _role == "NA":
                    _v = re.sub(r"(?i)^(job title|role|position)\s*[:\-]\s*", "", _line).strip()
                    if _v and _v.lower() not in ("na", "not applicable", ""):
                        _role = _v
        return {"company": _company, "role": _role, "status": "Not Applied"}

    last = df.iloc[-1]
    return {
        "company": str(last.get("Company", "NA") or "NA"),
        "role":    str(last.get("Role", "NA") or "NA"),
        "status":  str(last.get("Status", "Not Applied") or "Not Applied"),
    }

with st.expander("➕ Review & save application entry to database", expanded=False):
    st.caption("Pre-filled from the latest tracker entry. Review and correct before saving to Supabase.")

    defaults = get_latest_tracker_defaults()

    c1, c2 = st.columns(2)
    with c1:
        new_company = st.text_input(
            "Company",
            value=defaults["company"],
            key="new_company",
            help="Edit if the auto-extracted company name is wrong or NA.",
        )
        new_role = st.text_input(
            "Role",
            value=defaults["role"],
            key="new_role",
            help="Edit if the auto-extracted role is wrong or NA.",
        )
    with c2:
        new_date   = st.date_input("Date Applied", value=date.today(), key="new_date")
        new_status = st.selectbox(
            "Status",
            STATUS_OPTIONS,
            index=STATUS_OPTIONS.index(defaults["status"]) if defaults["status"] in STATUS_OPTIONS else 0,
            key="new_status",
        )

    st.info(f"📋 **Review before saving:** Company: **{new_company}** | Role: **{new_role}** | Date: **{new_date}** | Status: **{new_status}**")

    col_save, col_cancel = st.columns([1, 5])
    with col_save:
        if st.button("✅ Confirm & Save to Supabase", use_container_width=True):
            row_id = insert_to_supabase(new_company, new_role, str(new_date), new_status)
            # Update the latest matching row in session state with confirmed data
            df = st.session_state.tracker_edit
            if not df.empty:
                last_idx = df.index[-1]
                st.session_state.tracker_edit.at[last_idx, "Company"]      = new_company
                st.session_state.tracker_edit.at[last_idx, "Role"]         = new_role
                st.session_state.tracker_edit.at[last_idx, "Date Applied"] = str(new_date)
                st.session_state.tracker_edit.at[last_idx, "Status"]       = new_status
                if "id" in st.session_state.tracker_edit.columns and row_id:
                    st.session_state.tracker_edit.at[last_idx, "id"]       = row_id
            st.success("✅ Saved to Supabase.")
            st.rerun()

# ── Tracker table ──
# Build display df with entry number, hide internal id
display_df = st.session_state.tracker_edit.drop(columns=["id"], errors="ignore").copy()
# Convert Date Applied to datetime so DateColumn renders correctly
if "Date Applied" in display_df.columns:
    display_df["Date Applied"] = pd.to_datetime(
        display_df["Date Applied"], errors="coerce"
    ).dt.date
display_df.insert(0, "#", range(1, len(display_df) + 1))

edited = st.data_editor(
    display_df,
    num_rows="dynamic",
    use_container_width=True,
    disabled=["#"],
    column_config={
        "#":            st.column_config.NumberColumn("#", width="small"),
        "Company":      st.column_config.TextColumn("Company"),
        "Role":         st.column_config.TextColumn("Role"),
        "Date Applied": st.column_config.DateColumn("Date Applied", format="YYYY-MM-DD"),
        "Status":       st.column_config.SelectboxColumn("Status", options=STATUS_OPTIONS, required=True),
    },
    key="tracker_editor",
)

# Sync edits back to Supabase row by row (drop # column before comparing/saving)
edited_no_num = edited.drop(columns=["#"], errors="ignore")
display_no_num = display_df.drop(columns=["#"], errors="ignore")

if not edited_no_num.equals(display_no_num):
    id_col = st.session_state.tracker_edit["id"] if "id" in st.session_state.tracker_edit.columns else pd.Series(dtype=str)
    for i, row in edited_no_num.iterrows():
        row_id = id_col.iloc[i] if i < len(id_col) else ""
        update_supabase_row(
            str(row_id),
            str(row.get("Company", "")),
            str(row.get("Role", "")),
            str(row.get("Date Applied", "")),
            str(row.get("Status", "")),
        )
    edited_with_id = edited_no_num.copy()
    if "id" in st.session_state.tracker_edit.columns:
        edited_with_id["id"] = st.session_state.tracker_edit["id"].reindex(edited_no_num.index).values
    st.session_state.tracker_edit = edited_with_id

col_exp, _ = st.columns([1, 5])
with col_exp:
    st.download_button(
        label="⬇ Export CSV",
        data=edited.drop(columns=["#"], errors="ignore").to_csv(index=False).encode("utf-8"),
        file_name=f"applications_{date.today().isoformat()}.csv",
        mime="text/csv",
        use_container_width=True,
    )

if SUPABASE_URL and SUPABASE_ANON_KEY:
    st.caption("✅ Connected to Supabase — entries saved automatically.")
else:
    st.caption("⚠️ Supabase not configured — add SUPABASE_URL and SUPABASE_ANON_KEY to .env")