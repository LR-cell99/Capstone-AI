import streamlit as st
import openai
import os
import re
import io
from datetime import date
from pathlib import Path
import pandas as pd
from dotenv import load_dotenv
from fpdf import FPDF
import pdfplumber
from docx import Document

load_dotenv()
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")


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
SYSTEM_PROMPT = """You are an expert resume writer and ATS optimisation specialist.

You will receive a candidate's resume and a job description. Your job is to:

1. EXTRACT from the resume:
   - Candidate name, contact info, summary
   - All work experience (company, title, dates, responsibilities, achievements)
   - Skills (technical and soft)
   - Education and certifications

2. EXTRACT from the job description:
   - Role title and key responsibilities
   - Required and preferred skills/qualifications
   - Keywords and phrases used repeatedly
   - Seniority level and culture signals

3. REWRITE the resume tailored to the JD:
   - Preserve the same section structure
   - NEVER invent or hallucinate skills, roles, or achievements
   - Weave JD keywords naturally into bullet points and summary
   - Use strong action verbs (led, built, reduced, drove, optimised, delivered)
   - Quantify achievements where the original data supports it
   - Bullet points in result-first or STAR format
   - Summary: 3-5 lines tightly matched to the JD

IMPORTANT: The job description may contain scraped noise such as cookie banners, salary info, \
EA registration numbers, platform UI text, apply buttons, or unrelated job listings. \
IGNORE all such noise and focus only on the actual role title, responsibilities, and requirements.

OUTPUT RULES — strictly enforced:
- Output ONLY the resume content itself. Nothing else.
- Do NOT include any introductory sentence, closing remark, or commentary of any kind.
- Do NOT include lines like "This resume has been tailored to...", "Note:", "I have updated...", or any explanation.
- Do NOT include markdown formatting or code fences of any kind.
- The very first character of your response must be the start of the resume (e.g. the candidate's name).
- The very last character must be the end of the resume content. Nothing after it."""


# ── JD extraction prompt ──────────────────────────────────────────────────────
JD_EXTRACT_PROMPT = """You are a job description parser. You will receive raw text scraped from a job portal.
It will contain noise: cookie banners, salary ranges, EA/registration numbers, recruiter labels, portal UI elements, application instructions, and unrelated job listings.

Your job is to extract ONLY the following from the actual job posting itself:

1. Company name
   - This must be the actual hiring company or organisation, NOT a recruiter label, applicant status tag, or portal badge.
   - Examples of things that are NOT company names: "Strong applicant", "Top applicant", "Inter Island Manpower", "Quick apply", "Active today"
   - If you cannot confidently identify the real company name, leave this blank — do not guess.

2. Job title / role

3. About the company
   - Only include if there is a genuine company description in the JD.
   - Do NOT include recruiter agency descriptions.

4. Key responsibilities and duties

5. Required qualifications and skills

6. Preferred / nice-to-have qualifications

7. Any other content directly relevant to performing the role

OUTPUT RULES:
- Output clean plain text with the section labels above.
- Do NOT include salary, benefits, allowances, EA numbers, registration numbers, application instructions, cookie notices, portal UI text, or recruiter boilerplate.
- Do NOT include applicant status labels like "Strong applicant", "Medium application volume", "Posted X ago".
- If a section is not present in the JD, skip it entirely — do not invent or guess content.
- When in doubt about whether something is real JD content or portal noise, leave it out."""


# ── Session state ──────────────────────────────────────────────────────────────
if "enhanced_resume" not in st.session_state:
    st.session_state.enhanced_resume = ""
if "jd_text" not in st.session_state:
    st.session_state.jd_text = ""
if "tracker_edit" not in st.session_state:
    st.session_state.tracker_edit = pd.DataFrame(
        [{"Company": "", "Role": "", "Date Applied": str(date.today()), "Status": "Pending"}] * 2
    )

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
        st.success("✅ API key loaded from .env")
    else:
        st.error("❌ OPENAI_API_KEY not found in .env")

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
        with st.spinner("Extracting and filtering JD…"):
            extracted_jd = extract_text(jd_file)
        if extracted_jd:
            # Step 1: regex pre-clean to remove obvious portal boilerplate
            pre_cleaned = clean_jd_text(extracted_jd)

            # Step 2: AI extraction to keep only role-relevant content
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
            st.success(f"✅ JD processed from **{jd_file.name}** — showing relevant content only")
            with st.expander("Preview extracted JD"):
                st.text(ai_cleaned_jd[:1500] + ("…" if len(ai_cleaned_jd) > 1500 else ""))
        else:
            st.error("Could not extract text. Try a different file format.")

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

        with st.spinner("Enhancing your resume…"):
            try:
                response = client.chat.completions.create(
                    model=model,
                    messages=[
                        {"role": "system", "content": SYSTEM_PROMPT},
                        {"role": "user", "content": (
                            f"## Candidate Resume\n\n{st.session_state.resume_text}"
                            f"\n\n---\n\n## Job Description\n\n{st.session_state.jd_text}"
                        )},
                    ],
                    temperature=0.4,
                    max_tokens=2500,
                )
                enhanced = response.choices[0].message.content.strip()
                st.session_state.enhanced_resume = enhanced

                Path(output_dir).mkdir(parents=True, exist_ok=True)
                filepath = Path(output_dir) / f"enhanced_{date.today().isoformat()}.txt"
                filepath.write_text(enhanced, encoding="utf-8")
                st.success(f"✅ Saved to `{filepath}`")

                company = guess_company(st.session_state.jd_text)
                role    = guess_role(st.session_state.jd_text)
                if company or role:
                    new_row = pd.DataFrame([{
                        "Company": company,
                        "Role": role,
                        "Date Applied": str(date.today()),
                        "Status": "Applied",
                    }])
                    st.session_state.tracker_edit = pd.concat(
                        [st.session_state.tracker_edit, new_row], ignore_index=True
                    )

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

st.markdown(
    '<div class="warn-box">⚠️ <strong>Always review the enhanced resume carefully.</strong> '
    'AI may occasionally embellish skills or experience — validate every change before submitting applications.</div>',
    unsafe_allow_html=True,
)


# ── Application Tracker ────────────────────────────────────────────────────────
st.divider()
st.subheader("📋 Application Tracker")
st.caption("Track every application. Edit directly in the table below.")

STATUS_OPTIONS = ["Pending", "Applied", "Interview", "Offer", "Rejected"]

edited = st.data_editor(
    st.session_state.tracker_edit,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "Status": st.column_config.SelectboxColumn("Status", options=STATUS_OPTIONS, required=True),
        "Date Applied": st.column_config.TextColumn("Date Applied"),
    },
    key="tracker_editor",
)

st.session_state.tracker_edit = edited

col_exp, _ = st.columns([1, 5])
with col_exp:
    st.download_button(
        label="⬇ Export CSV",
        data=edited.to_csv(index=False).encode("utf-8"),
        file_name=f"applications_{date.today().isoformat()}.csv",
        mime="text/csv",
        use_container_width=True,
    )