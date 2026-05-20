# Capstone-AI

> AI-powered resume enhancer — tailored to job descriptions

![Python](https://img.shields.io/badge/Python-3-blue) ![OpenAI](https://img.shields.io/badge/API-OpenAI-green) ![Streamlit](https://img.shields.io/badge/UI-Streamlit-red)

---

## Section 1 — Project Title & Description

### AI Enhancer

AI Enhancer takes two inputs — a user's resume and a downloaded Job Description (JD) file — runs them through the application, and outputs two files:

- An AI-enhanced resume saved to a `resume/` folder
- An iterative Excel tracker to log applications

Designed for anyone who wants stronger, ATS-optimised resumes with minimal effort.

| No. | Company | Job | Application Status |
|-----|---------|-----|--------------------|
| 1   |         |     | _(dropdown)_       |
| 2   |         |     | _(dropdown)_       |

---

## Section 2 — Problem Statement

Most resumes fail ATS (Applicant Tracking System) checks before a human ever reads them. This tool uses AI to tailor and enhance resumes against specific JDs, increasing the chance of passing automated filters and landing interviews.

---

## Section 3 — Technology Stack

**Language**
- Python

**Libraries**
- `python-dotenv`
- `streamlit`

**API**
- OpenAI

---

## Section 4 — Setup Instructions

1. Clone the repository
2. Install dependencies
3. Copy `.env.example` to `.env` and fill in your API key
4. Run the application

---

## Section 5 — Usage Examples

- Enhances resumes that can pass ATS checks

---

## Section 6 — Known Limitations

> **Note:** Requires human validation — always review the enhanced resume to ensure the AI has not hallucinated skills or experience that do not exist.
> **Note:** 

---

## Section 7 — Future Improvements

- Filter to specific roles and scrape JD data from non-dynamic job sites, removing the manual search step and increasing the volume of applications processed automatically.
