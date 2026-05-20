import streamlit as st
import openai
import os
import re
from datetime import date
from pathlib import Path
import pandas as pd
from dotenv import load_dotenv

load_dotenv()  # loads OPENAI_API_KEY from .env
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="AI Resume Enhancer",
    page_icon="📄",
    layout="wide",
)

# ── Custom CSS ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { background-color: #f5f2ee; }
    .block-container { padding-top: 2rem; padding-bottom: 3rem; }
    h1 { font-size: 2rem !important; color: #1a1714 !important; }
    .stTextArea textarea {
        font-family: 'DM Mono', monospace;
        font-size: 13px;
        background-color: #ffffff;
        border: 1px solid #e0dbd3;
        border-radius: 10px;
    }
    .stButton > button {
        background-color: #2d5a3d;
        color: white;
        border: none;
        border-radius: 10px;
        padding: 0.5rem 1.5rem;
        font-weight: 500;
    }
    .stButton > button:hover { background-color: #214a2e; }
    .warn-box {
        background: #fdf6e3;
        border: 1px solid #e6c87a;
        border-radius: 10px;
        padding: 0.75rem 1rem;
        font-size: 13px;
        color: #7a5a10;
        margin: 1rem 0;
    }
    .output-box {
        background: #ffffff;
        border: 1px solid #e0dbd3;
        border-radius: 12px;
        padding: 1.25rem;
        font-family: 'DM Mono', monospace;
        font-size: 13px;
        line-height: 1.8;
        white-space: pre-wrap;
        min-height: 200px;
    }
    div[data-testid="stExpander"] {
        background: #ffffff;
        border: 1px solid #e0dbd3;
        border-radius: 12px;
    }
</style>
""", unsafe_allow_html=True)


# ── System prompt ──────────────────────────────────────────────────────────────
SYSTEM_PROMPT = """You are an expert resume writer and ATS optimisation specialist.

Your task: rewrite the candidate's resume so it is strongly tailored to the job description provided.

Rules:
- Preserve the same section structure (Summary, Experience, Skills, Education, etc.)
- NEVER invent or hallucinate skills, roles, or achievements the candidate does not have
- Incorporate relevant JD keywords naturally into bullet points and the summary
- Use strong action verbs (led, built, reduced, drove, optimised, delivered, etc.)
- Quantify achievements where the original data supports it
- Write bullet points in result-first or STAR format where possible
- Summary: 3–5 lines, tightly matched to the JD's key requirements
- Output ONLY the enhanced resume as plain text — no preamble, no commentary, no markdown fences"""


# ── Helper: extract company/role from JD ──────────────────────────────────────
def guess_company(jd: str) -> str:
    m = re.search(r'(?:at|@|about\s+)([A-Z][A-Za-z0-9&\s.,\']+?)(?:\n|,|\.|\s+is\s|\s+are\s)', jd)
    return m.group(1).strip()[:50] if m else ""

def guess_role(jd: str) -> str:
    m = re.search(r'(?:role|position|title|hiring(?:\s+a[n]?)?)[\s:]+([A-Za-z\s\/\-]+?)(?:\n|,|\.|\s+at\s)', jd, re.I)
    return m.group(1).strip()[:60] if m else ""


# ── Session state init ─────────────────────────────────────────────────────────
if "enhanced_resume" not in st.session_state:
    st.session_state.enhanced_resume = ""

if "tracker" not in st.session_state:
    st.session_state.tracker = pd.DataFrame(
        columns=["#", "Company", "Role", "Date Applied", "Status"]
    )

if "tracker_edit" not in st.session_state:
    st.session_state.tracker_edit = pd.DataFrame(
        [{"Company": "", "Role": "", "Date Applied": str(date.today()), "Status": "Pending"}] * 2
    )


# ── Header ─────────────────────────────────────────────────────────────────────
col_title, col_badge = st.columns([4, 1])
with col_title:
    st.markdown("# 📄 AI Resume Enhancer")
    st.caption("ATS-optimised resumes tailored to your target job description — powered by OpenAI")

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
    st.markdown("**Output folder**")
    output_dir = st.text_input("Save enhanced resumes to:", value="resume/")
    st.caption("Enhanced resumes are saved here as `.txt` files.")

    st.divider()
    st.markdown("""
**Tips for best results:**
- Paste the *full* JD, including requirements
- Include all your experience in the resume input
- Review output carefully before applying
""")


# ── Input panels ───────────────────────────────────────────────────────────────
col_resume, col_jd = st.columns(2)

with col_resume:
    st.subheader("Your Resume")
    resume_text = st.text_area(
        label="resume_input",
        label_visibility="collapsed",
        placeholder="Paste your full resume as plain text…\n\nInclude: work experience, skills, education, summary, achievements — everything.",
        height=350,
        key="resume_input",
    )

with col_jd:
    st.subheader("Job Description")
    jd_text = st.text_area(
        label="jd_input",
        label_visibility="collapsed",
        placeholder="Paste the full job description here…\n\nInclude responsibilities, requirements, and nice-to-haves for best results.",
        height=350,
        key="jd_input",
    )


# ── Enhance button ─────────────────────────────────────────────────────────────
st.write("")
col_btn, col_clear = st.columns([1, 5])

with col_btn:
    enhance_clicked = st.button("⚡ Enhance Resume", use_container_width=True)

with col_clear:
    if st.button("Clear", use_container_width=False):
        st.session_state.enhanced_resume = ""
        st.session_state.resume_input = ""
        st.session_state.jd_input = ""
        st.rerun()


if enhance_clicked:
    if not OPENAI_API_KEY:
        st.error("OPENAI_API_KEY not found. Please add it to your .env file.")
    elif not resume_text.strip():
        st.error("Please paste your resume text.")
    elif not jd_text.strip():
        st.error("Please paste the job description.")
    else:
        client = openai.OpenAI(api_key=OPENAI_API_KEY)

        with st.spinner("Enhancing your resume…"):
            try:
                response = client.chat.completions.create(
                    model=model,
                    messages=[
                        {"role": "system", "content": SYSTEM_PROMPT},
                        {"role": "user", "content": f"## My Resume\n\n{resume_text}\n\n---\n\n## Job Description\n\n{jd_text}"},
                    ],
                    temperature=0.4,
                    max_tokens=2500,
                )
                enhanced = response.choices[0].message.content.strip()
                st.session_state.enhanced_resume = enhanced

                # Save to file
                Path(output_dir).mkdir(parents=True, exist_ok=True)
                filename = f"enhanced_{date.today().isoformat()}.txt"
                filepath = Path(output_dir) / filename
                filepath.write_text(enhanced, encoding="utf-8")
                st.success(f"Saved to `{filepath}`")

                # Auto-log to tracker
                company = guess_company(jd_text)
                role    = guess_role(jd_text)
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
    st.download_button(
        label="⬇ Download as .txt",
        data=st.session_state.enhanced_resume,
        file_name=f"enhanced_resume_{date.today().isoformat()}.txt",
        mime="text/plain",
    )
else:
    st.markdown(
        '<div class="output-box" style="color:#bbb5ad;font-family:sans-serif;font-style:italic">'
        'Your enhanced resume will appear here after clicking Enhance Resume.'
        '</div>',
        unsafe_allow_html=True,
    )

# Warning box
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
        "Status": st.column_config.SelectboxColumn(
            "Status",
            options=STATUS_OPTIONS,
            required=True,
        ),
        "Date Applied": st.column_config.TextColumn("Date Applied"),
    },
    key="tracker_editor",
)

st.session_state.tracker_edit = edited

col_exp, col_dl = st.columns([1, 5])
with col_exp:
    csv_data = edited.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="⬇ Export CSV",
        data=csv_data,
        file_name=f"applications_{date.today().isoformat()}.csv",
        mime="text/csv",
        use_container_width=True,
    )