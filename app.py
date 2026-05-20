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


def extract_text_from_txt(file) -> str:
    return file.read().decode("utf-8", errors="ignore").strip()


def extract_text(uploaded_file) -> str:
    """Route to correct extractor based on file type."""
    name = uploaded_file.name.lower()
    file_bytes = io.BytesIO(uploaded_file.read())
    if name.endswith(".pdf"):
        return extract_text_from_pdf(file_bytes)
    elif name.endswith(".docx"):
        return extract_text_from_docx(file_bytes)
    elif name.endswith(".txt"):
        file_bytes.seek(0)
        return file_bytes.read().decode("utf-8", errors="ignore").strip()
    else:
        return ""


# ── PDF export ─────────────────────────────────────────────────────────────────
def generate_pdf(text: str) -> bytes:
    pdf = FPDF()
    pdf.add_page()
    pdf.set_margins(20, 20, 20)
    pdf.set_auto_page_break(auto=True, margin=20)

    for line in text.split("\n"):
        stripped = line.strip()
        if stripped.isupper() and len(stripped) > 2:
            pdf.set_font("Helvetica", "B", 13)
            pdf.ln(3)
            pdf.cell(0, 8, stripped, ln=True)
            pdf.set_draw_color(45, 90, 61)
            pdf.set_line_width(0.5)
            pdf.line(20, pdf.get_y(), 190, pdf.get_y())
            pdf.ln(2)
        elif stripped.startswith(("-", "•", "*")):
            pdf.set_font("Helvetica", "", 10)
            content = stripped.lstrip("-•* ").strip()
            pdf.set_x(25)
            pdf.cell(5, 6, "•", ln=False)
            pdf.multi_cell(0, 6, content)
        elif stripped and len(stripped) < 60 and not stripped.endswith("."):
            pdf.set_font("Helvetica", "B", 10)
            pdf.cell(0, 6, stripped, ln=True)
        elif stripped:
            pdf.set_font("Helvetica", "", 10)
            pdf.multi_cell(0, 6, stripped)
        else:
            pdf.ln(3)

    return bytes(pdf.output())


# ── JD helpers ─────────────────────────────────────────────────────────────────
def guess_company(jd: str) -> str:
    m = re.search(r'(?:at|@|about\s+)([A-Z][A-Za-z0-9&\s.,\']+?)(?:\n|,|\.|\s+is\s|\s+are\s)', jd)
    return m.group(1).strip()[:50] if m else ""

def guess_role(jd: str) -> str:
    m = re.search(r'(?:role|position|title|hiring(?:\s+a[n]?)?)[\s:]+([A-Za-z\s\/\-]+?)(?:\n|,|\.|\s+at\s)', jd, re.I)
    return m.group(1).strip()[:60] if m else ""


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
   - Summary: 3–5 lines tightly matched to the JD

Output ONLY the enhanced resume as plain text — no preamble, no commentary, no markdown fences."""


# ── Session state ──────────────────────────────────────────────────────────────
if "enhanced_resume" not in st.session_state:
    st.session_state.enhanced_resume = ""
if "resume_text" not in st.session_state:
    st.session_state.resume_text = ""
if "jd_text" not in st.session_state:
    st.session_state.jd_text = ""
if "tracker_edit" not in st.session_state:
    st.session_state.tracker_edit = pd.DataFrame(
        [{"Company": "", "Role": "", "Date Applied": str(date.today()), "Status": "Pending"}] * 2
    )


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
    .extract-preview {
        border-radius: 8px; padding: 0.6rem 0.9rem;
        font-size: 12px; font-family: monospace;
        line-height: 1.6; max-height: 180px; overflow-y: auto;
    }
</style>
""", unsafe_allow_html=True)


# ── Header ─────────────────────────────────────────────────────────────────────
st.markdown("# 📄 AI Resume Enhancer")
st.caption("Upload your resume and JD — the AI extracts, analyses, and tailors automatically.")
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
    st.subheader("📎 Your Resume")
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
            st.success(f"✅ Extracted {len(extracted.split())} words from **{resume_file.name}**")
            with st.expander("Preview extracted text"):
                st.text(extracted[:1500] + ("…" if len(extracted) > 1500 else ""))
        else:
            st.error("Could not extract text. Try a different file format.")

    # Fallback: manual paste
    with st.expander("Or paste resume manually"):
        manual_resume = st.text_area(
            "Resume text",
            label_visibility="collapsed",
            placeholder="Paste your resume here as plain text…",
            height=200,
            key="manual_resume",
        )
        if manual_resume.strip():
            st.session_state.resume_text = manual_resume


with col_jd:
    st.subheader("📋 Job Description")
    jd_file = st.file_uploader(
        "Upload JD",
        type=["pdf", "docx", "txt"],
        label_visibility="collapsed",
        key="jd_upload",
    )

    if jd_file:
        with st.spinner("Extracting JD text…"):
            extracted_jd = extract_text(jd_file)
        if extracted_jd:
            st.session_state.jd_text = extracted_jd
            st.success(f"✅ Extracted {len(extracted_jd.split())} words from **{jd_file.name}**")
            with st.expander("Preview extracted text"):
                st.text(extracted_jd[:1500] + ("…" if len(extracted_jd) > 1500 else ""))
        else:
            st.error("Could not extract text. Try a different file format.")

    # Fallback: manual paste
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

        with st.spinner("Extracting key info and enhancing your resume…"):
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

                # Save txt to folder
                Path(output_dir).mkdir(parents=True, exist_ok=True)
                filepath = Path(output_dir) / f"enhanced_{date.today().isoformat()}.txt"
                filepath.write_text(enhanced, encoding="utf-8")
                st.success(f"✅ Saved to `{filepath}`")

                # Auto-log to tracker
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

if st.session_state.enhanced_resume:
    st.text_area(
        label="enhanced_output",
        label_visibility="collapsed",
        value=st.session_state.enhanced_resume,
        height=400,
        key="enhanced_output",
    )
    col_dl1, col_dl2, col_rest = st.columns([1, 1, 4])
    with col_dl1:
        st.download_button(
            label="⬇ Download .txt",
            data=st.session_state.enhanced_resume,
            file_name=f"enhanced_resume_{date.today().isoformat()}.txt",
            mime="text/plain",
            use_container_width=True,
        )
    with col_dl2:
        pdf_bytes = generate_pdf(st.session_state.enhanced_resume)
        st.download_button(
            label="⬇ Download .pdf",
            data=pdf_bytes,
            file_name=f"enhanced_resume_{date.today().isoformat()}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
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