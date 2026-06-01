# Capstone-AI

> AI-powered resume enhancer — tailored to job descriptions

![Python](https://img.shields.io/badge/Python-3.9+-blue) ![OpenAI](https://img.shields.io/badge/API-OpenAI-green) ![Streamlit](https://img.shields.io/badge/UI-Streamlit-red) ![Supabase](https://img.shields.io/badge/DB-Supabase-optional-lightgrey)

---

## Section 1 — Project Title & Description

### AI Resume Enhancer

The AI Resume Enhancer takes two inputs — a base resume file and a Job Description (JD) file — runs them through OpenAI, and outputs:

- An AI-enhanced resume saved to a `resume/` folder automatically (as `.txt`)
- Downloadable as `.txt` or `.pdf` directly from the app
- An editable application tracker to log and manage job applications per session
- An ATS compatibility score (0–100) with a before/after comparison between the baseline and enhanced resume
- An exportable Excel tracker with auto-fitted columns

The base resume is loaded automatically from a `base_resume/` folder — only the JD needs to be uploaded per session. The JD is cleaned of portal noise before being used for enhancement.

| # | Company | Role | Date Applied | Status |
|---|---------|------|-------------|--------|
| 1 | _(auto-filled from JD)_ | _(auto-filled from JD)_ | _(today's date)_ | _(dropdown)_ |
| 2 | | | | |

---

## Section 2 — Problem Statement

Most resumes fail ATS (Applicant Tracking System) checks before a human ever reads them. This tool uses AI to:

1. Extract and clean the key role requirements from any JD file (stripping portal noise, salary info, and boilerplate)
2. Tailor and enhance the resume against those requirements using targeted rewriting — not a full replacement
3. Score the resume against Singapore ATS standards and SkillsFuture skill frameworks before and after enhancement
4. Present the score improvement with a breakdown across 5 categories so the user knows exactly where to focus

---

## Section 3 — Technology Stack

**Language:** Python 3.9+

**Core Libraries:**

| Package | Version | Purpose |
|---------|---------|---------|
| `streamlit` | 1.45.1 | Web UI framework |
| `openai` | 1.83.0 | AI enhancement and ATS scoring via GPT models |
| `pdfplumber` | 0.11.9 | Extract text from PDF resumes and JDs |
| `python-docx` | 1.2.0 | Extract text from Word `.docx` files |
| `fpdf2` | 2.8.7 | Export enhanced resume as formatted PDF |
| `pandas` | 3.0.0 | Application tracker table and data handling |
| `python-dotenv` | 1.1.0 | Load API keys from `.env` file |
| `openpyxl` | 3.1.5 | Export tracker to formatted Excel with auto-fitted columns |

**Optional — Database:**

| Package | Version | Purpose |
|---------|---------|---------|
| `supabase` | 2.15.1 | Cloud database for persistent application tracker (optional) |

**API:** OpenAI (GPT-4o recommended for enhancement; GPT-4o-mini used for JD extraction and ATS scoring)

---

## Section 4 — Setup Instructions

### 1. Clone the repository
```bash
git clone https://github.com/LR-cell99/Capstone-AI.git
cd Capstone-AI
```

### 2. Install dependencies

Make sure you have **Python 3.9 or higher** installed. Then run:

```bash
pip install -r requirements.txt
```

> **Note:** If you have an old `docx` package installed (not `python-docx`), remove it first as it conflicts:
> ```bash
> pip uninstall docx -y
> pip install python-docx
> ```

### 3. Set your API key

Copy `.env.example` to a new file called `.env`:

```bash
# Windows
copy .env.example .env

# Mac / Linux
cp .env.example .env
```

Open `.env` and fill in your OpenAI API key:

```
OPENAI_API_KEY=sk-...your-key-here...
```

You can get your API key from: https://platform.openai.com/api-keys

> **Important:** Never commit your `.env` file — it is already listed in `.gitignore`.

### 4. Add your base resume

Create a folder called `base_resume/` in the project root and place your resume file inside it (PDF, DOCX, or TXT). The app will auto-load it on startup — you will not need to upload your resume every session.

```
Capstone-AI/
├── app.py
├── base_resume/
│   └── my_resume.pdf     ← place your resume here
├── resume/               ← enhanced resumes saved here automatically
├── .env
└── ...
```

### 5. Run the application

```bash
streamlit run app.py
```

The app will open automatically in your browser at `http://localhost:8501`.

---

## Section 4b — Optional: Supabase Database Setup

The application tracker works fully offline using an in-session table and Excel export. Supabase is **optional** and only needed if you want entries to persist across sessions in a cloud database.

### Why Supabase is optional

Without Supabase:
- The tracker stores entries in session memory only
- Entries are lost when the app is closed
- You can still export the full tracker to Excel at any time

With Supabase:
- Selected entries are saved to a cloud database on explicit confirmation
- Entries persist across sessions and app restarts
- Data is viewable in the Supabase Table Editor or via SQL

### Setting up Supabase (optional)

1. Create a free project at https://supabase.com
2. Go to **SQL Editor** and run the following to create the table:

```sql
CREATE TABLE applications (
    id uuid DEFAULT gen_random_uuid() PRIMARY KEY,
    company text,
    role text,
    date_applied date,
    status text,
    created_at timestamp DEFAULT now()
);

ALTER TABLE applications ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Allow all for anon"
ON applications FOR ALL TO anon
USING (true) WITH CHECK (true);

GRANT ALL ON public.applications TO anon;
GRANT ALL ON public.applications TO authenticated;
```

3. Go to **Settings → API Keys → Legacy API Keys** and copy the `anon` key (starts with `eyJ...`)
4. Go to **Settings → API** and copy your Project URL
5. Add both to your `.env`:

```
OPENAI_API_KEY=sk-...
SUPABASE_URL=https://your-project-id.supabase.co
SUPABASE_ANON_KEY=eyJ...your-legacy-anon-key...
```

> **Note:** Use the legacy `anon` key (starting with `eyJ`) for best compatibility with `supabase-py`. The newer `sb_publishable_` key format has known compatibility issues with the Python client library.

---

## Section 5 — Usage

### Basic workflow
1. Place your base resume in the `base_resume/` folder — it loads automatically on startup
2. Upload a Job Description (PDF, DOCX, or TXT) in the JD panel
3. Review the cleaned JD extraction — edit directly to correct company name, job role, or any missing info
4. Click **Enhance Resume** — the app runs a 3-step process:
   - Step 1: Scores your baseline resume against the JD
   - Step 2: Enhances the resume using targeted AI rewriting
   - Step 3: Scores the enhanced resume and compares against baseline
5. Review the ATS score — baseline vs enhanced with a breakdown across 5 categories
6. Edit the enhanced resume directly in the output box if needed
7. Download as `.txt` or `.pdf`
8. Open the `💾 Select & save entry to Supabase` expander, select the entry, review, and confirm to save

### Application Tracker
- A new tracker row is auto-staged after each enhancement with company, role, date, and status pre-filled from the JD
- Delete unwanted rows before saving
- Export the full tracker to Excel at any time using **⬇ Export Excel**
- Use **Clear all** to reset the session and start fresh with a new JD

### ATS Score
- Scores range from **0 to 100** — passing is **60 and above**
- Based on Singapore ATS standards and SkillsFuture Singapore skill frameworks
- This is an **AI-simulated ATS score** — it provides indicative feedback rather than guaranteed ATS performance, similar to tools like Jobscan and Resume Worded
- The score breakdown covers: Keyword Match, Relevance of Experience, Qualifications Match, Resume Clarity, and ATS Formatting Compliance

---

## Section 6 — Known Limitations

- **Requires human validation** — always review the enhanced resume to ensure the AI has not hallucinated skills or experience that do not exist
- **ATS score is simulated** — real ATS systems (Workday, Taleo, MCF) have proprietary algorithms; this tool provides an indicative score, not a guarantee
- **PDF extraction quality** — scanned image-based PDFs will not extract well; use text-based PDF or DOCX where possible
- **JD quality matters** — thin JDs from recruiter agencies with minimal detail will produce weaker enhancements and less meaningful ATS score differences; JDs from company career pages or MCF tend to give better results
- **Session-only tracker** — without Supabase configured, tracker data is lost when the app is closed; export to Excel before closing

---

## Section 7 — Future Improvements

- Scrape JD data directly from MCF and non-dynamic job sites, removing the manual upload step
- Cover letter generation tailored to the same JD and enhanced resume
- Side-by-side diff view highlighting exactly what changed between the original and enhanced resume
- Multi-resume support — compare enhancements across different base resumes for the same JD
- Deployment to a web server for multi-user access with individual Supabase row isolation
