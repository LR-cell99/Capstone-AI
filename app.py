import streamlit as st
import openai, os, re, io, json, hashlib
from datetime import date, datetime
from pathlib import Path
import pandas as pd
from dotenv import load_dotenv
from fpdf import FPDF
import pdfplumber, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from docx import Document
try:
    import gspread
    from gspread.utils import ValueInputOption
except ImportError:          # Google Sheets is optional
    gspread = None

load_dotenv()
BASE_RESUME_DIR   = Path("base_resume")
SUPPORTED_EXTS    = (".pdf", ".docx", ".txt")
STATUS_OPTIONS    = ["Not Applied", "Pending", "Applied", "Interview", "Offer", "Rejected"]
# Only auto-load base_resume/ when you explicitly turn it on (keep it off on a shared/hosted app)
AUTO_LOAD_BASE_RESUME = os.getenv("AUTO_LOAD_BASE_RESUME", "false").lower() == "true"
# Keep a running history of applications in a local file across sessions.
# Turn on only for your own copy — on a shared app every user would see the same file.
TRACKER_AUTOSAVE  = os.getenv("TRACKER_AUTOSAVE", "false").lower() == "true"
TRACKER_FILE      = Path(os.getenv("TRACKER_FILE", "applications/applications_history.csv"))
TRACKER_COLS      = ["Company", "Role", "Date Applied", "Status"]
# Google Sheets (optional) — uses a Google service account; see .env.example
GOOGLE_SA_FILE    = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "").strip()
GOOGLE_SA_JSON    = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
GOOGLE_SHEET_URL  = os.getenv("GOOGLE_SHEET_URL", "").strip()   # your own sheet; pre-filled in the app

# ── AI providers ───────────────────────────────────────────────────────────────
# Every provider below speaks the OpenAI-compatible API, so one client handles all of them.
# A provider only appears in the app if its key (or, for Ollama, its URL) is in .env.
_ollama_url = os.getenv("OLLAMA_BASE_URL", "").strip().rstrip("/")
if _ollama_url and not _ollama_url.endswith("/v1"):
    _ollama_url += "/v1"

PROVIDERS = {
    "OpenAI": dict(
        env_prefix="OPENAI", key_env="OPENAI_API_KEY", base_url=None, list_models=False,
        models=["gpt-4o", "gpt-4o-mini", "gpt-4-turbo", "gpt-3.5-turbo"], fast="gpt-4o-mini", timeout=60.0),
    "Ollama (local)": dict(
        env_prefix="OLLAMA", key_env=None, base_url=_ollama_url, list_models=True,
        models=[], fast=None, timeout=300.0),
    "Groq": dict(
        env_prefix="GROQ", key_env="GROQ_API_KEY", base_url="https://api.groq.com/openai/v1", list_models=True,
        models=[], fast=None, timeout=60.0),
    "OpenRouter": dict(
        env_prefix="OPENROUTER", key_env="OPENROUTER_API_KEY", base_url="https://openrouter.ai/api/v1", list_models=True,
        models=[], fast=None, timeout=90.0),
    "Google Gemini": dict(
        env_prefix="GEMINI", key_env="GEMINI_API_KEY",
        base_url="https://generativelanguage.googleapis.com/v1beta/openai/", list_models=True,
        models=[], fast=None, timeout=60.0),
}

# Filled in by the sidebar on every run: {"provider", "model", "fast_model"}
LLM_CFG: dict | None = None

def provider_key(name: str) -> str:
    p = PROVIDERS[name]
    if p["key_env"] is None:               # Ollama needs no key, just a URL
        return "ollama" if p["base_url"] else ""
    return os.getenv(p["key_env"], "").strip()

def available_providers() -> list[str]:
    return [n for n in PROVIDERS if provider_key(n)]

@st.cache_resource
def get_client(name: str) -> openai.OpenAI:
    p = PROVIDERS[name]
    return openai.OpenAI(api_key=provider_key(name), base_url=p["base_url"] or None,
                         timeout=p["timeout"], max_retries=2)

_SKIP_MODEL_WORDS = ("embed", "whisper", "tts", "audio", "transcri", "image", "dall-e",
                     "moderation", "guard", "realtime", "aqa", "imagen", "veo")

@st.cache_data(ttl=600, show_spinner=False)
def list_models(name: str) -> list[str]:
    """Models offered for a provider. <PREFIX>_MODELS in .env (comma-separated) overrides the
    live list — use it to limit what other users can pick."""
    p = PROVIDERS[name]
    pinned = [m.strip() for m in os.getenv(f"{p['env_prefix']}_MODELS", "").split(",") if m.strip()]
    if pinned:
        return pinned
    if not p["list_models"]:
        return p["models"]
    try:
        ids = [m.id for m in get_client(name).models.list().data]
    except Exception:
        return p["models"]
    ids = [i.removeprefix("models/") for i in ids if not any(w in i.lower() for w in _SKIP_MODEL_WORDS)]
    if name == "OpenRouter":
        ids.sort(key=lambda i: (not i.endswith(":free"), i))   # free models first
    else:
        ids.sort()
    return ids or p["models"]

def llm_ready() -> bool:
    return bool(LLM_CFG and LLM_CFG.get("model"))

def _strip_think(text: str) -> str:
    # Reasoning models (deepseek-r1, qwen3, …) can put their thinking in the reply
    return re.sub(r"(?is)<think>.*?</think>", "", text).strip()

def llm(messages: list, *, fast: bool = False, temperature: float = 0.0, max_tokens: int = 1000) -> str:
    """Single entry point for every AI call. fast=True uses the scoring/extraction model."""
    if not llm_ready():
        raise RuntimeError("No AI provider/model selected — check the sidebar.")
    client = get_client(LLM_CFG["provider"])
    kwargs = dict(model=LLM_CFG["fast_model"] if fast else LLM_CFG["model"],
                  messages=messages, temperature=temperature, max_tokens=max_tokens)
    try:
        r = client.chat.completions.create(**kwargs)
    except openai.BadRequestError as e:
        # Some newer OpenAI models reject max_tokens / custom temperature — retry without them
        msg = str(e).lower()
        if "max_tokens" in msg or "temperature" in msg:
            kwargs.pop("temperature", None)
            kwargs["max_completion_tokens"] = kwargs.pop("max_tokens")
            r = client.chat.completions.create(**kwargs)
        else:
            raise
    return _strip_think(r.choices[0].message.content or "")

# ── Application tracker storage (no database) ─────────────────────────────────
def empty_tracker() -> pd.DataFrame:
    return pd.DataFrame(columns=TRACKER_COLS)

def normalise_tracker(df: pd.DataFrame) -> pd.DataFrame:
    """Coerce any tracker-like table (session, saved CSV, imported Excel) to the standard columns."""
    if df is None or df.empty:
        return empty_tracker()
    df = df.copy()
    lower = {c.lower().strip(): c for c in df.columns}
    aliases = {"Company": ["company", "company name"], "Role": ["role", "job title", "position"],
               "Date Applied": ["date applied", "date", "date_applied", "applied on"], "Status": ["status"]}
    out = pd.DataFrame()
    for col, names in aliases.items():
        src = next((lower[n] for n in names if n in lower), None)
        out[col] = df[src] if src is not None else ""
    out["Company"] = out["Company"].fillna("").astype(str).str.strip()
    out["Role"]    = out["Role"].fillna("").astype(str).str.strip()
    out["Status"]  = out["Status"].fillna("").astype(str).str.strip()
    out.loc[~out["Status"].isin(STATUS_OPTIONS), "Status"] = "Not Applied"
    out["Date Applied"] = pd.to_datetime(out["Date Applied"], errors="coerce").dt.date
    out = out[(out["Company"] != "") | (out["Role"] != "")]          # drop blank rows
    return out.reset_index(drop=True)

def merge_trackers(*dfs: pd.DataFrame) -> pd.DataFrame:
    """Combine tracker tables; the same Company + Role + Date keeps only its latest entry."""
    frames = [normalise_tracker(d) for d in dfs if d is not None and not d.empty]
    if not frames:
        return empty_tracker()
    merged = pd.concat(frames, ignore_index=True)
    key = (merged["Company"].str.lower() + "|" + merged["Role"].str.lower() + "|" + merged["Date Applied"].astype(str))
    merged = merged.loc[~key.duplicated(keep="last")]
    return merged.sort_values("Date Applied", na_position="last", kind="stable").reset_index(drop=True)

def load_local_tracker() -> pd.DataFrame:
    if not TRACKER_AUTOSAVE or not TRACKER_FILE.exists():
        return empty_tracker()
    try:
        return normalise_tracker(pd.read_csv(TRACKER_FILE))
    except Exception as e:
        st.warning(f"⚠️ Could not read saved tracker {TRACKER_FILE}: {e}")
        return empty_tracker()

def save_local_tracker(df: pd.DataFrame):
    if not TRACKER_AUTOSAVE:
        return
    try:
        TRACKER_FILE.parent.mkdir(parents=True, exist_ok=True)
        normalise_tracker(df).to_csv(TRACKER_FILE, index=False)
    except Exception as e:
        st.warning(f"⚠️ Could not save tracker to {TRACKER_FILE}: {e}")

def read_tracker_upload(f) -> pd.DataFrame:
    if f.name.lower().endswith(".csv"):
        return normalise_tracker(pd.read_csv(f))
    sheets = pd.read_excel(f, sheet_name=None)                 # all sheets
    sheet = sheets.get("Application Tracker", next(iter(sheets.values())))
    return normalise_tracker(sheet)

def current_tracker() -> pd.DataFrame:
    """Tracker including any edits made in the table on the previous run."""
    live = st.session_state.get("tracker_edit_live")
    return live if live is not None else st.session_state.tracker_edit

def set_tracker(df: pd.DataFrame):
    """Replace the tracker from outside the table (enhance, import, clear) and refresh the editor."""
    st.session_state.tracker_edit = normalise_tracker(df)
    st.session_state.tracker_edit_live = st.session_state.tracker_edit
    st.session_state.tracker_version += 1
    save_local_tracker(st.session_state.tracker_edit)

# ── Google Sheets ──────────────────────────────────────────────────────────────
GS_TRACKER_TAB = "Application Tracker"
GS_SUMMARY_TAB = "Summary"

def gs_configured() -> bool:
    return gspread is not None and bool(GOOGLE_SA_JSON or GOOGLE_SA_FILE)

def gs_service_email() -> str:
    """The service account's email — a sheet must be shared with it (Editor) before the app can use it."""
    try:
        info = json.loads(GOOGLE_SA_JSON) if GOOGLE_SA_JSON else json.loads(Path(GOOGLE_SA_FILE).read_text())
        return info.get("client_email", "")
    except Exception:
        return ""

@st.cache_resource
def get_gspread():
    if GOOGLE_SA_JSON:
        return gspread.service_account_from_dict(json.loads(GOOGLE_SA_JSON))
    return gspread.service_account(filename=GOOGLE_SA_FILE)

def gs_open(url_or_id: str):
    gc = get_gspread()
    u = url_or_id.strip()
    return gc.open_by_url(u) if u.startswith("http") else gc.open_by_key(u)

def gs_pull(url_or_id: str) -> pd.DataFrame:
    sh = gs_open(url_or_id)
    try:
        ws = sh.worksheet(GS_TRACKER_TAB)
    except gspread.exceptions.WorksheetNotFound:
        return empty_tracker()                      # new/empty sheet
    rows = ws.get_all_values()
    if len(rows) < 2:
        return empty_tracker()
    return normalise_tracker(pd.DataFrame(rows[1:], columns=rows[0]))

def _gs_tab(sh, title: str, rows: int, cols: int):
    try:
        ws = sh.worksheet(title)
        ws.clear()
        ws.resize(rows=rows, cols=cols)
        return ws
    except gspread.exceptions.WorksheetNotFound:
        return sh.add_worksheet(title=title, rows=rows, cols=cols)

def gs_push(df: pd.DataFrame, url_or_id: str) -> str:
    """Write the tracker to the sheet (replacing the tracker tab) and a live Summary tab. Returns the sheet URL."""
    sh = gs_open(url_or_id)
    df = normalise_tracker(df)
    df["Date Applied"] = df["Date Applied"].apply(lambda d: d.strftime("%Y-%m-%d") if pd.notna(d) else "")
    values = [TRACKER_COLS] + df[TRACKER_COLS].astype(str).values.tolist()
    n_rows = max(len(values) + 100, 200)                 # spare rows for adding entries in Sheets

    ws = _gs_tab(sh, GS_TRACKER_TAB, n_rows, len(TRACKER_COLS))
    ws.update(values=values, range_name="A1", value_input_option=ValueInputOption.user_entered)
    ws.format("A1:D1", {"backgroundColor": {"red": 0.176, "green": 0.353, "blue": 0.239},
                        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
                        "horizontalAlignment": "CENTER"})
    ws.format("C2:C", {"numberFormat": {"type": "DATE", "pattern": "yyyy-mm-dd"}})   # unambiguous dates
    ws.freeze(rows=1)
    sh.batch_update({"requests": [
        {"setDataValidation": {   # Status dropdown, same options as the app
            "range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": n_rows, "startColumnIndex": 3, "endColumnIndex": 4},
            "rule": {"condition": {"type": "ONE_OF_LIST", "values": [{"userEnteredValue": s} for s in STATUS_OPTIONS]},
                     "showCustomUi": True, "strict": True}}},
        {"setBasicFilter": {"filter": {"range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": n_rows,
                                                 "startColumnIndex": 0, "endColumnIndex": len(TRACKER_COLS)}}}},
        {"autoResizeDimensions": {"dimensions": {"sheetId": ws.id, "dimension": "COLUMNS",
                                                 "startIndex": 0, "endIndex": len(TRACKER_COLS)}}},
    ]})

    # Summary uses formulas, so it stays correct when you edit statuses directly in Google Sheets
    tab = f"'{GS_TRACKER_TAB}'"
    summ = [["Status", "Count"]] + [[s, f'=COUNTIF({tab}!D2:D,"{s}")'] for s in STATUS_OPTIONS]
    summ += [["Total", f"=COUNTA({tab}!A2:A)"], [], [f"Last synced from app {datetime.now().strftime('%Y-%m-%d %H:%M')}"]]
    ss = _gs_tab(sh, GS_SUMMARY_TAB, len(summ) + 5, 2)
    ss.update(values=summ, range_name="A1", value_input_option=ValueInputOption.user_entered)
    ss.format("A1:B1", {"textFormat": {"bold": True}})
    ss.format(f"A{len(STATUS_OPTIONS) + 2}:B{len(STATUS_OPTIONS) + 2}", {"textFormat": {"bold": True}})
    return sh.url

def gs_error_message(e: Exception) -> str:
    email = gs_service_email()
    share = f" Share the sheet with **{email}** as an Editor." if email else ""
    if gspread and isinstance(e, gspread.exceptions.SpreadsheetNotFound):
        return "Sheet not found — check the link." + share
    if gspread and isinstance(e, gspread.exceptions.APIError):
        code = getattr(getattr(e, "response", None), "status_code", None)
        if code in (403, 404):
            return "No access to this sheet." + share
        if code == 429:
            return "Google Sheets rate limit hit — wait a minute and try again."
    return str(e)

# ── Text extraction ────────────────────────────────────────────────────────────
def extract_pdf(file) -> str:
    with pdfplumber.open(file) as pdf:
        return "\n".join(p.extract_text() for p in pdf.pages if p.extract_text()).strip()

def extract_docx(file) -> str:
    return "\n".join(p.text for p in Document(file).paragraphs if p.text.strip())

def extract_file(f) -> str:
    b = io.BytesIO(f.read())
    n = f.name.lower()
    if n.endswith(".pdf"):   return extract_pdf(b)
    if n.endswith(".docx"):  return extract_docx(b)
    b.seek(0); return b.read().decode("utf-8", errors="ignore").strip()

def load_base_resume() -> tuple[str, str]:
    if not BASE_RESUME_DIR.exists():
        return "", ""
    for ext in SUPPORTED_EXTS:
        matches = list(BASE_RESUME_DIR.glob(f"*{ext}"))
        if matches:
            fp = matches[0]
            b = io.BytesIO(fp.read_bytes())
            text = extract_pdf(b) if ext == ".pdf" else extract_docx(b) if ext == ".docx" else fp.read_text(encoding="utf-8", errors="ignore").strip()
            return text, fp.name
    return "", ""

# ── PDF export ─────────────────────────────────────────────────────────────────
_PDF_CHARS = {"–":"-","—":"-","‒":"-","―":"-","‘":"'",
              "’":"'","“":'"',"”":'"',"•":"-","…":"..."," ":" "}

def sanitise(text: str) -> str:
    for c, r in _PDF_CHARS.items():
        text = text.replace(c, r)
    return text.encode("latin-1", errors="ignore").decode("latin-1")

def generate_pdf(text: str) -> bytes:
    pdf = FPDF()
    pdf.add_page()
    pdf.set_margins(20, 20, 20)
    pdf.set_auto_page_break(True, 20)
    pw = pdf.w - pdf.l_margin - pdf.r_margin
    line_h = 6
    indent = 5

    def est_lines(txt, width, font_size=10):
        pdf.set_font("Helvetica", "", font_size)
        char_w = pdf.get_string_width("m")
        chars_per_line = max(1, int(width / char_w))
        words = txt.split()
        lines, current = 1, 0
        for word in words:
            wlen = len(word) + 1
            if current + wlen > chars_per_line:
                lines += 1; current = wlen
            else:
                current += wlen
        return lines

    def space_left():
        return pdf.h - pdf.b_margin - pdf.get_y()

    def block_height(lines_ahead):
        """Estimate total height of a group of consecutive lines."""
        total = 0
        for ln in lines_ahead:
            ln = ln.strip()
            if not ln:
                total += 3
            elif ln.isupper() and len(ln) > 2:
                total += 8 + 2 + 4
            elif ln.startswith(("-", "*")):
                total += est_lines(ln.lstrip("-* ").strip(), pw - indent - 4) * line_h + 2
            elif len(ln) < 80 and not ln.endswith("."):
                total += est_lines(ln, pw) * line_h + 1
            else:
                total += est_lines(ln, pw) * line_h + 1
        return total

    # Pre-sanitise all lines
    all_lines = [sanitise(l) for l in text.split("\n")]

    i = 0
    while i < len(all_lines):
        s = all_lines[i].strip()

        if not s:
            pdf.ln(3); i += 1; continue

        if s.isupper() and len(s) > 2:
            # Section header — check space for header + rule + at least first bullet
            lookahead = []
            j = i + 1
            while j < len(all_lines) and len(lookahead) < 3:
                lookahead.append(all_lines[j]); j += 1
            needed = block_height([s] + lookahead)
            if space_left() < min(needed, 30):
                pdf.add_page()
            pdf.set_font("Helvetica", "B", 13); pdf.ln(3)
            pdf.set_x(pdf.l_margin)
            pdf.multi_cell(pw, 8, s)
            pdf.set_draw_color(45, 90, 61); pdf.set_line_width(0.5)
            pdf.line(pdf.l_margin, pdf.get_y(), pdf.l_margin + pw, pdf.get_y()); pdf.ln(2)

        elif s.startswith(("-", "*")):
            # Bullet — keep whole bullet together
            bullet_text = s.lstrip("-* ").strip()
            bullet_width = pw - indent - 4
            needed = est_lines(bullet_text, bullet_width) * line_h + 2
            if space_left() < needed:
                pdf.add_page()
            pdf.set_font("Helvetica", "", 10)
            pdf.set_x(pdf.l_margin + indent)
            pdf.cell(4, line_h, "-", ln=False)
            pdf.multi_cell(bullet_width, line_h, bullet_text)

        elif len(s) < 80 and not s.endswith("."):
            # Short bold line — group with following bullets on same page
            lookahead = []
            j = i + 1
            while j < len(all_lines) and len(lookahead) < 4:
                nxt = all_lines[j].strip()
                if nxt.isupper() and len(nxt) > 2:
                    break  # stop at next section header
                lookahead.append(all_lines[j]); j += 1
            needed = block_height([s] + lookahead)
            if space_left() < min(needed, 25):
                pdf.add_page()
            pdf.set_font("Helvetica", "B", 10)
            pdf.set_x(pdf.l_margin)
            pdf.multi_cell(pw, line_h, s)

        else:
            # Body text
            needed = est_lines(s, pw) * line_h + 1
            if space_left() < needed:
                pdf.add_page()
            pdf.set_font("Helvetica", "", 10)
            pdf.set_x(pdf.l_margin)
            pdf.multi_cell(pw, line_h, s)

        i += 1

    return bytes(pdf.output())

# ── Excel export ───────────────────────────────────────────────────────────────
def export_excel(df: pd.DataFrame) -> bytes:
    df = normalise_tracker(df)
    df["Date Applied"] = df["Date Applied"].apply(lambda d: d.strftime("%Y-%m-%d") if pd.notna(d) else "")
    out = io.BytesIO()
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Application Tracker"
    hfill = PatternFill("solid", fgColor="2D5A3D")
    hfont = Font(bold=True, color="FFFFFF", size=11)
    halign = Alignment(horizontal="center", vertical="center", wrap_text=True)
    border = Border(left=Side(style="thin"), right=Side(style="thin"),
                    top=Side(style="thin"), bottom=Side(style="thin"))
    for ci, col in enumerate(df.columns, 1):
        c = ws.cell(1, ci, col); c.font = hfont; c.fill = hfill
        c.alignment = halign; c.border = border
    ws.row_dimensions[1].height = 20
    afill = PatternFill("solid", fgColor="E8F2EB")
    for ri, row in enumerate(df.itertuples(index=False), 2):
        for ci, val in enumerate(row, 1):
            c = ws.cell(ri, ci, str(val) if val is not None else "")
            c.alignment = Alignment(vertical="center", wrap_text=True); c.border = border
            if ri % 2 == 0: c.fill = afill
        ws.row_dimensions[ri].height = 18
    for ci, col in enumerate(df.columns, 1):
        mx = max([len(str(col))] + [max(len(l) for l in str(v).split("\n")) for v in df[col].astype(str)])
        ws.column_dimensions[get_column_letter(ci)].width = min(max(mx + 4, 12), 60)
    ws.freeze_panes = "A2"
    if len(df):
        ws.auto_filter.ref = f"A1:{get_column_letter(len(df.columns))}{len(df) + 1}"

    # Summary sheet — count of applications per status
    ss = wb.create_sheet("Summary")
    for ci, h in enumerate(["Status", "Count"], 1):
        c = ss.cell(1, ci, h); c.font = hfont; c.fill = hfill; c.alignment = halign; c.border = border
    for ri, status in enumerate(STATUS_OPTIONS, 2):
        ss.cell(ri, 1, status).border = border
        ss.cell(ri, 2, int((df["Status"] == status).sum())).border = border
    tr = len(STATUS_OPTIONS) + 2
    ss.cell(tr, 1, "Total").font = Font(bold=True)
    ss.cell(tr, 2, len(df)).font = Font(bold=True)
    ss.cell(tr + 2, 1, f"Exported {datetime.now().strftime('%Y-%m-%d %H:%M')}").font = Font(italic=True, color="888888")
    ss.column_dimensions["A"].width = 18; ss.column_dimensions["B"].width = 10

    wb.save(out); return out.getvalue()

# ── JD helpers ─────────────────────────────────────────────────────────────────
_NOISE = [r"(?i)by continuing to use our platform", r"(?i)open app", r"(?i)privacy policy",
          r"(?i)cookie.*consent", r"(?i)your application will include", r"(?i)employer questions",
          r"(?i)registration no\.", r"(?i)\bEA No\b", r"(?i)company information",
          r"(?i)view all jobs", r"(?i)(full|part) time.*per month",
          r"(?i)\$[\d,]+\s*[-–]\s*\$[\d,]+\s*per\s*(month|year|hour)",
          r"(?i)posted \d+\w* ago", r"(?i)(medium|high|low) application volume",
          r"(?i)(benefits.*allowance|variable bonus|apply now|quick apply|save job|report job|share this job|similar jobs|you might also like|javascript is disabled)",
          r"(?i)(central|north|south|east|west) region"]

def clean_jd(raw: str) -> str:
    lines = [l.strip() for l in raw.split("\n")
             if not any(re.search(p, l.strip()) for p in _NOISE)]
    return re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip()

def clean_field_value(val: str) -> str:
    """Strip parenthetical expansions and extra whitespace from extracted field values.
    e.g. 'SSMC (Systems on Silicon Manufacturing Company Pte. Ltd.)' -> 'SSMC'
    """
    val = re.sub(r"\s*\(.*?\)", "", val).strip()  # Remove anything in parentheses
    val = re.sub(r"\s{2,}", " ", val)               # Collapse multiple spaces
    return val.strip()

def parse_jd_field(text: str, labels: list) -> str:
    lines = text.split("\n")
    for i, line in enumerate(lines):
        line = line.strip()
        for pat in labels:
            if re.match(pat + r"\s*[:\-]\s*.+", line, re.I):
                v = re.sub(pat + r"\s*[:\-]\s*", "", line, flags=re.I).strip()
                v = clean_field_value(v)
                if v and v.lower() not in ("na", "not applicable", "none", ""):
                    return v
            if re.match(pat + r"\s*[:\-]?\s*$", line, re.I):
                for nxt in lines[i+1:i+3]:
                    nxt = clean_field_value(nxt.strip())
                    if nxt and not re.match(r"(?i)^(job title|role|position|key responsibilities|required|preferred|company name|company|about)", nxt):
                        if nxt.lower() not in ("na", "not applicable", "none", ""):
                            return nxt
                        break
    return "NA"

# ── Prompts ────────────────────────────────────────────────────────────────────
SYSTEM_PROMPT = """You are a senior professional resume writer for the Singapore job market. Edit the candidate's resume to better align it with the job description — not a full rewrite.

CRITICAL OUTPUT RULES — read these first, follow them absolutely:
- Output ONLY the resume content. Start with the candidate's name. End with the last line of resume content.
- Do NOT write any sentence after the resume ends. No summary, no note, no explanation.
- Do NOT write lines like "The resume has been updated...", "Key skills have been reframed...", "This resume now better aligns...", "Note:", "I have..." or anything similar.
- If you find yourself about to write an explanatory sentence — stop. Delete it. Output only the resume.

GOAL: The enhanced resume must feel like the same person wrote it, with sharper language, better JD alignment, and stronger framing. Preserve the candidate's authentic voice and actual experience. Only change what needs changing — but everything that can be aligned to the JD, should be.

STEP 1 — ANALYSE BEFORE EDITING:
- Read every section of the resume fully.
- Extract from the JD: role title, key responsibilities, required skills, preferred skills, industry context, and any repeated terminology.
- For every section of the resume, ask: does the language, framing, and emphasis reflect what this JD is looking for?
- Note what is already strong and JD-aligned — leave those parts alone.
- Identify everything that can be reworded, reframed, or reordered to better reflect the JD — without adding anything that isn't there.

STEP 2 — TARGETED EDITING BY SECTION:

EXECUTIVE SUMMARY / PROFESSIONAL PROFILE:
- Use the original as your foundation — do not discard the candidate's intent.
- Rewrite to open with the JD's exact job title or a close variant.
- Naturally embed 3-5 JD keywords and the industry context (e.g. semiconductor, manufacturing, automation).
- Ensure the tone and focus reflects what the JD prioritises — if the JD emphasises cross-functional collaboration, the summary should mention it; if it emphasises technical depth, lead with that.
- Keep the same general length. The candidate should recognise their own summary.

WORK EXPERIENCE BULLETS — rewrite EVERY bullet using the CAR format with explicit labels. MANDATORY.
Each bullet must be structured as three separate labelled lines:

Context: [the situation, responsibility, team, equipment, or scope — 1 sentence]
Action: [what YOU specifically did — strong verb, JD-aligned terminology, tools/methods named]
Result: [outcome, impact, improvement, or value delivered — implied if not stated in original]

BEFORE:
"Assisted in maintenance tasks"

AFTER:
Context: Responsible for supporting maintenance operations across production-line equipment in a manufacturing environment.
Action: Performed scheduled preventive maintenance, fault diagnosis, and component inspections using standard operating procedures and JD-relevant tools.
Result: Contributed to reduced unplanned downtime and sustained equipment reliability across the production line.

Rules:
- Use the exact labels "Context:", "Action:", "Result:" on separate lines for every bullet. No exceptions.
- ALL THREE labels must be present for every rewritten bullet.
- Use the JD's exact terminology in the Action line (e.g. "root cause analysis", "preventive maintenance", "cross-functional collaboration").
- Do not fabricate specific numbers. Implied results drawn from context are acceptable.
- Strong verbs: Executed, Implemented, Collaborated, Optimised, Reduced, Troubleshot, Calibrated, Validated, Coordinated, Monitored, Analysed, Resolved, Streamlined, Facilitated, Engineered, Delivered, Spearheaded.

PROJECTS — rewrite EVERY project description using the CAR format. MANDATORY, no exceptions.
Each project entry must be structured with explicit labels on separate lines:

Context: [what the project was, the problem it addressed, or its scope — 1-2 sentences]
Action: [what YOU specifically did — strong verb, JD-relevant skills and tools named explicitly]
Result: [what was built, achieved, demonstrated, or delivered]

BEFORE:
"Built a cleaning robot end effector"

AFTER:
Context: Developed a robotic end effector for a capstone project addressing automated cleaning of public dining surfaces in hawker centre environments.
Action: Applied CAD modelling, materials evaluation, and iterative prototyping to design and test grip mechanisms, optimising for performance and durability.
Result: Delivered a functional prototype that validated the feasibility of the automated cleaning solution for large-scale deployment.

Rules:
- Use the exact labels "Context:", "Action:", "Result:" on separate lines for every project entry.
- ALL THREE labels must be present for every project. No exceptions.
- Name JD-relevant technologies, tools, and methodologies explicitly in the Action line.
- Do not change the project name, tech stack, or actual outcomes — only the framing and structure.
- If a project has multiple separate deliverables, write one full CAR block per deliverable.

ACHIEVEMENTS / AWARDS / CERTIFICATIONS:
- Reorder to place JD-relevant achievements first.
- Reword achievement descriptions to use JD language where applicable.
- Keep all facts intact — only rephrase, never fabricate.

SKILLS:
- Reorder so JD-critical skills appear first.
- Where a skill is named differently from the JD equivalent, use the JD's exact terminology.
- Only add a skill if the candidate's work experience or projects clearly demonstrate it.

EDUCATION:
- Keep factually intact. No changes unless a certification or qualification directly matches a JD requirement — in that case, make sure it is prominently visible.

HALLUCINATION RULES — non-negotiable:
- NEVER invent a job, company, project, degree, certification, or skill not in the original resume.
- NEVER fabricate metrics, outcomes, or achievements.
- Reframing, rewording, reordering, and re-emphasising existing content is encouraged and is NOT hallucination.
- You may make implied results explicit if the original context supports them.

OUTPUT RULES — final reminder:
- The very last character of your output must be the last character of resume content.
- Nothing after it. No summary. No note. No explanation. Just the resume."""

JD_EXTRACT_PROMPT = """You are a strict job description parser for Singapore job portals. Extract ONLY what a hiring manager actually wrote.

EXTRACT:
1. Company name — actual employer only. NOT recruiter agencies or applicant labels. Leave blank if uncertain.
2. Job title
3. About the company — only if genuine employer description exists
4. Key responsibilities
5. Required qualifications and skills
6. Preferred qualifications
7. Any other role-relevant content

STRIP everything else: salary, allowances, CPF, EA/registration numbers, application instructions, employer questions, cookie/privacy notices, portal UI ("Open app", "Save job", "Quick apply"), applicant tags ("Strong applicant", "Posted X ago"), recruiter boilerplate, unrelated listings, region tags.

OUTPUT: Clean plain text with section labels. Skip missing sections. When in doubt — leave it out."""

_ATS_RULES = """
STEP 0 — VALIDATE JD: If JD lacks real hiring content (Wikipedia, generic pages, all "Not applicable"), set all scores to 0, SUMMARY to "Invalid job description", IMPROVEMENTS to "Upload a real job posting", and stop.

SCORING (valid JD only):
- Scores range freely from 0 to 100. No floor, no ceiling, no anchor point.
- 60+ = PASS. Below 60 = FAIL.
- Score exactly what you see — reward strong alignment, penalise weak alignment, do not cluster around any number.
- A poor match can score 10-30. An average match 40-59. A good match 60-74. A strong match 75-89. An excellent match 90-100.

KEYWORD MATCH (30 pts):
- 0-3:   No JD keywords present, completely unrelated vocabulary
- 4-7:   Only 1-2 incidental keyword matches, core JD terms all missing
- 8-10:  A few relevant keywords but most critical JD terms absent
- 11-15: Some keywords present but inconsistently, many critical ones missing
- 16-18: Most keywords loosely present, a few critical ones still missing
- 19-24: Strong keyword presence across the resume, minor gaps only
- 25-30: Excellent — nearly all JD keywords and SkillsFuture terms present and naturally integrated

RELEVANCE OF EXPERIENCE (25 pts):
- 0-3:   Completely unrelated field, no transferable experience
- 4-6:   Largely unrelated, only peripheral relevance to the role
- 7-9:   Different field but some transferable skills visible
- 10-13: Partially relevant, significant gaps in core responsibilities
- 14-16: Some experience maps to the role but coverage is incomplete
- 17-20: Experience clearly maps to most role responsibilities
- 21-25: Excellent — experience directly and comprehensively addresses all JD responsibilities

QUALIFICATIONS MATCH (20 pts):
- 0-3:   Missing required degree and certification, completely wrong field
- 4-6:   Missing required degree, wrong field of study
- 7-9:   Has a degree but wrong discipline for a technical role, or missing key certifications
- 10-13: Partially meets requirements, some gaps in stated qualifications
- 14-17: Meets core requirements with minor gaps
- 18-20: Excellent — fully meets or exceeds all stated and preferred qualifications

RESUME CLARITY & STRUCTURE (15 pts):
- 0-2:   No structure, walls of text, no action verbs, no results whatsoever
- 3-5:   Very vague and passive language throughout, minimal structure
- 6-8:   Some structure but mostly weak language, few action verbs
- 9-11:  Inconsistent — some strong bullets alongside weak ones
- 12-13: Consistently clear language, good structure, most bullets action-oriented
- 14-15: Excellent — strong action verbs throughout, quantified results, tightly targeted summary

ATS FORMATTING COMPLIANCE (10 pts):
- 0-2:   Heavy use of tables, columns, graphics, or special characters — will fail most parsers
- 3-5:   Some formatting elements that may trip parsers
- 6-7:   Minor formatting issues, mostly clean
- 8-10:  Clean plain text, standard section headers, consistent date formatting

TOTAL_SCORE = exact sum of all five categories.

Respond ONLY in this format:
TOTAL_SCORE: <n>
KEYWORD_MATCH: <n>/30
RELEVANCE_OF_EXPERIENCE: <n>/25
QUALIFICATIONS_MATCH: <n>/20
RESUME_CLARITY: <n>/15
ATS_FORMATTING: <n>/10
SUMMARY: <2-3 sentences — name actual strengths AND specific weaknesses, reference exact missing keywords if any>
IMPROVEMENTS: <3 specific actionable bullet points referencing exact missing keywords or JD requirements>"""

BASELINE_ATS_PROMPT = (
    "You are a precise ATS evaluator for Singapore. Score the ORIGINAL BASELINE resume honestly. "
    "This is the benchmark — score it for what it actually is, not what it could be. "
    "A strong baseline resume against a well-matched JD can score 70, 80, or higher. Score freely across the full 0-100 range."
) + _ATS_RULES
BASELINE_ATS_TEMP   = 0.0

ENHANCED_ATS_PROMPT = (
    "You are a precise ATS evaluator for Singapore. Score this AI-ENHANCED resume against the JD. "
    "Score freely across the full 0-100 range — do not anchor to 60 or any other number. "
    "A resume that comprehensively addresses the JD with strong keywords, relevant experience, and clear structure SHOULD score 75-95+. "
    "Score what you see: reward genuine alignment, penalise genuine gaps."
) + _ATS_RULES
ENHANCED_ATS_TEMP   = 0.0

RETRY_ATS_PROMPT = (
    "You are a strict ATS evaluator for Singapore. Quick regression check — score this resume against the JD. "
    "Penalise vague or generic resumes. Reward strong JD alignment. Score freely from 0-100."
) + _ATS_RULES
RETRY_ATS_TEMP      = 0.0

SKILLS_GAP_PROMPT = """You are a career coach and skills analyst for the Singapore job market.

You will receive a resume and a job description. Analyse the skills match and return a structured gap analysis.

OUTPUT ONLY in this exact format — no extra text:

HAVE:
- <skill or competency from JD that the resume clearly demonstrates>
(repeat for each)

MISSING:
- <skill or competency required by the JD that is absent from the resume>
(repeat for each)

IRRELEVANT:
- <skill on the resume that is not mentioned or useful for this role>
(repeat for each)

Rules:
- Be specific — name actual skills, tools, methodologies, not vague categories
- HAVE: only include if clearly evidenced in the resume, not just implied
- MISSING: only include if explicitly required or strongly preferred in the JD
- IRRELEVANT: only include if clearly unrelated to the role
- If a section has nothing to list, write "None identified" under it
- Do not invent skills not mentioned in either document"""

REVISION_PROMPT = """You are a professional resume editor. You will receive:
1. The current enhanced resume
2. The original job description it was tailored for
3. A specific revision request from the candidate

CRITICAL OUTPUT RULES — follow these absolutely:
- Output ONLY the revised resume. Start with the candidate's name. End with the last line of resume content.
- Do NOT write any sentence after the resume ends. No summary, no note, no explanation of what you changed.
- Do NOT write lines like "The resume has been updated...", "Key skills have been reframed...", "I have added...", "Note:" or anything similar.
- The very last character must be the last character of resume content. Nothing after it.

Apply the requested revision carefully. Maintain the same structure and formatting.

HALLUCINATION RULES: NEVER add skills, jobs, degrees, or certifications not in the resume."""

# ── Session state ──────────────────────────────────────────────────────────────
_SS_DEFAULTS = {
    "enhanced_resume": "", "edited_resume": "", "_last_enhanced": "",
    "jd_text": "", "jd_filename": "", "jd_editable_value": "",
    "jd_widget_version": 0, "ats_result": None, "baseline_ats_result": None,
    "ats_model": "", "baseline_ats_model": "",
    "enhance_count": 0, "is_enhancing": False, "override_resume_name": "",
    "skills_gap": None, "revision_history": [], "revision_version": 0,
    "resume_versions": [], "ats_score_cache": {}, "active_version_idx": 0,
}
for k, v in _SS_DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v

if "tracker_edit" not in st.session_state:
    st.session_state.tracker_edit = load_local_tracker()   # empty unless TRACKER_AUTOSAVE=true
if "tracker_version" not in st.session_state:
    st.session_state.tracker_version = 0

if "resume_text" not in st.session_state or not st.session_state.resume_text:
    _t, _n = load_base_resume() if AUTO_LOAD_BASE_RESUME else ("", "")
    st.session_state.resume_text = _t
    st.session_state.base_resume_name = _n
elif "base_resume_name" not in st.session_state:
    st.session_state.base_resume_name = ""

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(page_title="AI Resume Enhancer", page_icon="📄", layout="wide")
st.markdown("""<style>
    .block-container{padding-top:2rem;padding-bottom:3rem}
    .stTextArea textarea{font-family:'DM Mono',monospace;font-size:13px;border-radius:10px}
    .stButton>button{background-color:#2d5a3d;color:white;border:none;border-radius:10px;padding:.5rem 1.5rem;font-weight:500}
    .stButton>button:hover{background-color:#214a2e}
    .warn-box{border:1px solid #e6c87a;border-radius:10px;padding:.75rem 1rem;font-size:13px;margin:1rem 0;background:#fdf6e3;color:#7a5a10}
    @media(prefers-color-scheme:dark){.warn-box{background:#2a2200;color:#f0c040;border-color:#7a5a10}}
    .output-box{border:1px solid #e0dbd3;border-radius:12px;padding:1.25rem;font-family:'DM Mono',monospace;font-size:13px;line-height:1.8;white-space:pre-wrap;min-height:200px;background:transparent;color:#bbb5ad}
</style>""", unsafe_allow_html=True)

st.markdown("# 📄 AI Resume Enhancer")
st.caption("Upload your resume and a job description to get a tailored version.")
st.divider()

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Settings")

    _providers = available_providers()
    if not _providers:
        st.error("❌ No AI provider configured. Add at least one of OPENAI_API_KEY, GROQ_API_KEY, "
                 "OPENROUTER_API_KEY, GEMINI_API_KEY or OLLAMA_BASE_URL to .env")
    else:
        _default_p = os.getenv("DEFAULT_PROVIDER", "")
        provider = st.selectbox("AI provider", _providers,
                                index=_providers.index(_default_p) if _default_p in _providers else 0,
                                help="Only providers with a key/URL in .env are listed.")
        _p = PROVIDERS[provider]
        with st.spinner("Loading models…"):
            _models = list_models(provider)

        _CUSTOM = "✏️ Custom model name…"
        _default_m = os.getenv(f"{_p['env_prefix']}_DEFAULT_MODEL", "")
        _opts = _models + [_CUSTOM]
        _choice = st.selectbox("Model (writing & revisions)", _opts,
                               index=_opts.index(_default_m) if _default_m in _opts else 0,
                               key=f"model_{provider}")
        if _choice == _CUSTOM:
            model = st.text_input("Model name", key=f"custom_model_{provider}",
                                  placeholder="e.g. llama3.1:8b").strip()
        else:
            model = _choice

        _fast_default = os.getenv(f"{_p['env_prefix']}_FAST_MODEL") or _p.get("fast") or model
        _fast_opts = list(dict.fromkeys(([model] if model else []) + _models))
        fast_model = st.selectbox("Model (scoring & JD extraction)", _fast_opts,
                                  index=_fast_opts.index(_fast_default) if _fast_default in _fast_opts else 0,
                                  key=f"fast_model_{provider}",
                                  help="Used for ATS scoring, JD cleanup and skills gap. A smaller, faster model is fine here.") if _fast_opts else ""

        if model:
            LLM_CFG = {"provider": provider, "model": model, "fast_model": fast_model or model}
            st.success(f"✅ {provider} ready")
        else:
            st.warning("⚠️ Pick or type a model name.")

        if provider.startswith("Ollama"):
            if not _models:
                st.warning("No local models found. Run `ollama pull <model>` first, and check OLLAMA_BASE_URL.")
            st.caption("Local models are slower and need a large context window — "
                       "start Ollama with `OLLAMA_CONTEXT_LENGTH=16384` or higher.")

    if TRACKER_AUTOSAVE:
        st.caption(f"💾 Tracker history saved to `{TRACKER_FILE}`")
    if gs_configured():
        st.caption("🟢 Google Sheets connected")
    st.divider()
    output_dir = st.text_input("Save resumes to folder:", value="resume/")
    st.divider()
    st.markdown("**File types:** PDF, DOCX, TXT\n\n**Tip:** Use full JDs from company career pages for best results.")

# ── Upload panels ──────────────────────────────────────────────────────────────
col_resume, col_jd = st.columns(2)

with col_resume:
    st.subheader("📎 Resume")

    # Show override resume info if one is uploaded, hide base resume
    if st.session_state.override_resume_name:
        st.success(f"✅ Using override: **{st.session_state.override_resume_name}** ({len(st.session_state.resume_text.split())} words)")
        with st.expander("Preview override resume"):
            st.text(st.session_state.resume_text[:1500] + ("…" if len(st.session_state.resume_text) > 1500 else ""))
        if AUTO_LOAD_BASE_RESUME and st.button("↩ Revert to base_resume/"):
            t, n = load_base_resume()
            if t:
                st.session_state.resume_text = t
                st.session_state.base_resume_name = n
                st.session_state.override_resume_name = ""
                st.session_state.resume_versions = []
                st.session_state.ats_score_cache = {}
                st.session_state.active_version_idx = 0
                st.session_state.enhanced_resume = ""
                st.session_state.edited_resume = ""
                st.session_state.ats_result = None
                st.session_state.baseline_ats_result = None
                st.session_state.skills_gap = None
                st.session_state.revision_history = []
                st.session_state.revision_version = 0
                st.rerun()
            else:
                st.error("No resume found in base_resume/.")
    else:
        if st.session_state.resume_text:
            n = st.session_state.base_resume_name or "manually entered"
            st.success(f"✅ Loaded **{n}** ({len(st.session_state.resume_text.split())} words)")
            with st.expander("Preview base resume"):
                st.text(st.session_state.resume_text[:1500] + ("…" if len(st.session_state.resume_text) > 1500 else ""))
            if AUTO_LOAD_BASE_RESUME and st.button("🔄 Reload from base_resume/"):
                t, n = load_base_resume()
                if t:
                    st.session_state.resume_text = t; st.session_state.base_resume_name = n; st.rerun()
                else:
                    st.error("No resume found in base_resume/.")
        elif AUTO_LOAD_BASE_RESUME:
            st.warning("⚠️ No resume found in `base_resume/` folder.")
            st.caption("Place your resume (PDF, DOCX, or TXT) in `base_resume/` next to app.py, or upload one below.")
        else:
            st.info("Upload your resume below to get started.")

    with st.expander("Upload a resume" if not st.session_state.resume_text else "Override: upload a different resume",
                     expanded=not st.session_state.resume_text):
        rf = st.file_uploader("Upload resume", type=["pdf","docx","txt"], label_visibility="collapsed", key="resume_upload")
        # Only process when a NEW file is uploaded — prevents re-clearing on every rerun
        if rf and rf.name != st.session_state.override_resume_name:
            t = extract_file(rf)
            if t:
                st.session_state.resume_text = t
                st.session_state.base_resume_name = rf.name
                st.session_state.override_resume_name = rf.name
                # New resume — clear version history and all results
                st.session_state.resume_versions = []
                st.session_state.ats_score_cache = {}
                st.session_state.active_version_idx = 0
                st.session_state.enhanced_resume = ""
                st.session_state.edited_resume = ""
                st.session_state.ats_result = None
                st.session_state.baseline_ats_result = None
                st.session_state.skills_gap = None
                st.session_state.revision_history = []
                st.session_state.revision_version = 0
                st.success(f"✅ Using **{rf.name}** — previous enhancements cleared.")
            else:
                st.error("Could not extract text.")

with col_jd:
    st.subheader("📋 Job Description")
    jf = st.file_uploader("Upload JD", type=["pdf","docx","txt"], label_visibility="collapsed", key="jd_upload")
    if jf and st.session_state.jd_filename != jf.name:
        with st.spinner("Extracting and filtering JD…"):
            raw = extract_file(jf)
        if raw:
            pre = clean_jd(raw)
            if llm_ready():
                try:
                    with st.spinner("Cleaning up JD with AI…"):
                        ai_jd = llm([{"role":"system","content":JD_EXTRACT_PROMPT},{"role":"user","content":pre}],
                                    fast=True, temperature=0.0, max_tokens=1000) or pre
                except openai.APITimeoutError:
                    ai_jd = pre
                    st.warning("⚠️ JD extraction timed out — using raw extracted text. Try again if the JD looks unclean.")
                except Exception as e:
                    ai_jd = pre
                    st.warning(f"⚠️ AI JD extraction failed — using raw extracted text. Reason: {e}")
            else:
                ai_jd = pre
                st.caption("No AI model selected — using raw extracted text.")
            st.session_state.jd_text = ai_jd
            st.session_state.jd_editable_value = ai_jd
            st.session_state.jd_filename = jf.name
            st.session_state.jd_widget_version += 1
            # New JD — clear version history and all results
            st.session_state.resume_versions = []
            st.session_state.ats_score_cache = {}
            st.session_state.active_version_idx = 0
            st.session_state.enhanced_resume = ""
            st.session_state.edited_resume = ""
            st.session_state.ats_result = None
            st.session_state.baseline_ats_result = None
            st.session_state.skills_gap = None
            st.session_state.revision_history = []
            st.session_state.revision_version = 0
            st.success(f"✅ JD processed from **{jf.name}**")
        else:
            st.error("Could not extract text.")

    if st.session_state.jd_editable_value:
        wk = f"jd_editable_{st.session_state.jd_widget_version}"
        if st.session_state.is_enhancing:
            st.caption("🔒 JD locked during enhancement.")
            st.text_area("JD", label_visibility="collapsed", value=st.session_state.jd_editable_value,
                         height=250, key=wk+"_locked", disabled=True)
        else:
            st.caption("✏️ Edit to correct company name, job role, or missing info before enhancing.")
            ejd = st.text_area("JD", label_visibility="collapsed", value=st.session_state.jd_editable_value,
                               height=250, key=wk, help="Edits here flow into enhancement and tracker.")
            st.session_state.jd_text = ejd
            st.session_state.jd_editable_value = ejd
    elif not st.session_state.jd_editable_value:
        with st.expander("Or paste JD manually"):
            mjd = st.text_area("JD text", label_visibility="collapsed",
                               placeholder="Paste the job description here…", height=200, key="manual_jd")
            if mjd.strip():
                st.session_state.jd_text = mjd; st.session_state.jd_editable_value = mjd

# ── Enhance button ─────────────────────────────────────────────────────────────
st.write("")
col_btn, col_clear = st.columns([1, 5])
with col_btn:
    enhance_clicked = st.button("⚡ Enhance Resume", use_container_width=True)
with col_clear:
    if st.button("Clear all"):
        for k, v in _SS_DEFAULTS.items():
            st.session_state[k] = v
        st.session_state.jd_widget_version += 1
        st.rerun()

if enhance_clicked:
    if not llm_ready():
        st.error("No AI provider/model selected — check the sidebar.")
    elif not st.session_state.resume_text.strip():
        st.error("Please upload a resume.")
    elif not st.session_state.jd_text.strip():
        st.error("Please upload a job description.")
    else:
        st.session_state.is_enhancing = True

        with st.spinner("Step 1/3 — Scoring baseline resume…"):
            try:
                st.session_state.baseline_ats_result = llm(
                    [{"role":"system","content":BASELINE_ATS_PROMPT},
                     {"role":"user","content":f"## Resume\n\n{st.session_state.resume_text}\n\n---\n\n## JD\n\n{st.session_state.jd_text}"}],
                    fast=True, temperature=BASELINE_ATS_TEMP, max_tokens=600)
                st.session_state.baseline_ats_model = f"{LLM_CFG['provider']} / {LLM_CFG['fast_model']}"
            except Exception as e:
                st.session_state.baseline_ats_result = None
                st.warning(f"⚠️ Baseline ATS scoring failed: {e}")

        with st.spinner(f"Step 2/3 — Enhancing your resume with {LLM_CFG['model']}…"):
            try:
                def sim(a, b):
                    wa, wb = set(a.lower().split()), set(b.lower().split())
                    return len(wa & wb) / max(len(wa), len(wb)) if wa and wb else 1.0

                uc = f"## Candidate Resume\n\n{st.session_state.resume_text}\n\n---\n\n## Job Description\n\n{st.session_state.jd_text}"
                baseline_total = None
                if st.session_state.baseline_ats_result:
                    m = re.search(r"TOTAL_SCORE:\s*([\d.]+)", st.session_state.baseline_ats_result)
                    if m: baseline_total = float(m.group(1))

                enhanced = None
                for attempt in range(3):
                    temp = [0.7, 0.85, 0.9][attempt]
                    candidate = llm([{"role":"system","content":SYSTEM_PROMPT},{"role":"user","content":uc}],
                                    temperature=temp, max_tokens=3500)

                    similarity = sim(st.session_state.resume_text, candidate)
                    cand_score = None
                    if baseline_total is not None:
                        try:
                            qtext = llm([{"role":"system","content":RETRY_ATS_PROMPT},
                                         {"role":"user","content":f"## Enhanced Resume\n\n{candidate}\n\n---\n\n## JD\n\n{st.session_state.jd_text}"}],
                                        fast=True, temperature=RETRY_ATS_TEMP, max_tokens=600)
                            qm = re.search(r"TOTAL_SCORE:\s*([\d.]+)", qtext)
                            if qm: cand_score = float(qm.group(1))
                        except Exception:
                            pass  # Retry check failed — continue with attempt regardless

                    regressed = baseline_total and cand_score and cand_score < baseline_total
                    if attempt == 2 or (similarity < 0.80 and not regressed):
                        enhanced = candidate
                        if similarity >= 0.80:
                            st.warning(f"⚠️ Enhancement is very similar to baseline ({similarity:.0%}). JD may have too little content.")
                        elif regressed:
                            st.warning(f"⚠️ Best available result after {attempt+1} attempt(s). Score could not be improved beyond baseline.")
                        break

                st.session_state.enhanced_resume = enhanced
                # Store as version 0 — label includes company and role from JD
                wk0 = f"jd_editable_{st.session_state.jd_widget_version}"
                src0 = st.session_state.get(wk0, st.session_state.jd_text)
                v0_company = parse_jd_field(src0, [r"(?i)company name", r"(?i)company"])
                v0_role    = parse_jd_field(src0, [r"(?i)job title", r"(?i)role", r"(?i)position"])
                v0_label   = f"Original — {v0_company} | {v0_role}" if v0_company != "NA" or v0_role != "NA" else "Original Enhancement"
                st.session_state.resume_versions = [{
                    "label": v0_label,
                    "content": enhanced,
                    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M"),
                }]
                st.session_state.active_version_idx = 0
                st.session_state.ats_score_cache = {}  # Clear cache for new JD/resume
                Path(output_dir).mkdir(parents=True, exist_ok=True)
                ts = datetime.now().strftime("%Y-%m-%d_%H%M")
                (Path(output_dir) / f"enhanced_{ts}.txt").write_text(enhanced, encoding="utf-8")
                st.success(f"✅ Saved to `{output_dir}enhanced_{ts}.txt`")

                st.session_state.enhance_count += 1
                wk = f"jd_editable_{st.session_state.jd_widget_version}"
                src = st.session_state.get(wk, st.session_state.jd_text)
                company = parse_jd_field(src, [r"(?i)company name", r"(?i)company"])
                role    = parse_jd_field(src, [r"(?i)job title", r"(?i)role", r"(?i)position"])
                set_tracker(pd.concat([current_tracker(), pd.DataFrame([{
                    "Company": company, "Role": role,
                    "Date Applied": date.today(), "Status": "Not Applied",
                }])], ignore_index=True))

                # ── Step 3: Skills gap analysis ──
                # (ATS scoring of enhanced resume is now a manual button — see below)
                try:
                    gap_content = f"## Resume\n\n{enhanced}\n\n---\n\n## JD\n\n{st.session_state.jd_text}"
                    st.session_state.skills_gap = llm(
                        [{"role":"system","content":SKILLS_GAP_PROMPT},{"role":"user","content":gap_content}],
                        fast=True, temperature=0.0, max_tokens=800)
                except Exception as e:
                    st.session_state.skills_gap = None
                    st.warning(f"⚠️ Skills gap analysis failed: {e}")

            except openai.AuthenticationError:
                st.error(f"❌ Invalid API key for {LLM_CFG['provider']} — check your .env file.")
            except openai.RateLimitError:
                st.error("❌ Rate limit hit — free tiers have low limits. Wait a moment, or switch provider/model.")
            except openai.NotFoundError:
                st.error(f"❌ Model `{LLM_CFG['model']}` not found on {LLM_CFG['provider']}. Pick another model.")
            except openai.APITimeoutError:
                st.error("❌ Request timed out — the model took too long to respond. Try again or use a smaller model.")
            except openai.APIConnectionError:
                st.error(f"❌ Could not reach {LLM_CFG['provider']}. If using Ollama, check it is running "
                         "(`ollama serve`) and that OLLAMA_BASE_URL is correct.")
            except Exception as e:
                st.error(f"❌ Enhancement failed: {e}")
            finally:
                st.session_state.is_enhancing = False

# ── Output ─────────────────────────────────────────────────────────────────────
st.divider()
st.subheader("Enhanced Resume Output")

if st.session_state.enhanced_resume:
    # ── Version switcher ──
    versions = st.session_state.resume_versions
    if len(versions) > 1:
        version_labels = [f"{'🔵' if i == 0 else '✏️'} {v['label']} ({v['timestamp']})"
                          for i, v in enumerate(versions)]
        selected_v = st.selectbox(
            "Select version to view / download:",
            version_labels,
            index=st.session_state.active_version_idx,
            key=f"version_select_{len(versions)}",
        )
        new_idx = version_labels.index(selected_v)
        if new_idx != st.session_state.active_version_idx:
            st.session_state.active_version_idx = new_idx
            active_content = versions[new_idx]["content"]
            st.session_state.edited_resume = active_content
            st.session_state.enhanced_resume = active_content
            st.session_state._last_enhanced = active_content
            st.session_state.revision_version += 1
            st.session_state.ats_result = None
    else:
        if versions:
            st.caption(f"🔵 {versions[0]['label']} ({versions[0]['timestamp']})")

    if st.session_state.get("_last_enhanced") != st.session_state.enhanced_resume:
        st.session_state.edited_resume = st.session_state.enhanced_resume
        st.session_state._last_enhanced = st.session_state.enhanced_resume

    st.caption("Edit directly below before exporting.")
    edited = st.text_area("Edit resume", label_visibility="collapsed",
                          value=st.session_state.edited_resume, height=500,
                          key=f"editable_output_{st.session_state.revision_version}",
                          help="Edit here — changes apply to exports.")
    st.session_state.edited_resume = edited

    _dl_ts = datetime.now().strftime("%Y-%m-%d_%H%M")
    col1, col2, col3, _ = st.columns([1, 1, 1, 3])
    with col1:
        st.download_button("⬇ Download .txt", data=edited,
                           file_name=f"enhanced_resume_{_dl_ts}.txt", mime="text/plain", use_container_width=True)
    with col2:
        try:
            pdf_data = generate_pdf(edited)
            st.download_button("⬇ Download .pdf", data=pdf_data,
                               file_name=f"enhanced_resume_{_dl_ts}.pdf", mime="application/pdf", use_container_width=True)
        except Exception as e:
            st.error(f"❌ PDF export failed: {e}. Try downloading as .txt instead.")
    with col3:
        if st.button("↺ Reset edits", use_container_width=True):
            active = versions[st.session_state.active_version_idx]["content"] if versions else st.session_state.enhanced_resume
            st.session_state.edited_resume = active
else:
    st.markdown('<div class="output-box">Your enhanced resume will appear here after clicking Enhance Resume.</div>', unsafe_allow_html=True)

# ── Skills Gap Analysis ───────────────────────────────────────────────────────
if st.session_state.skills_gap and st.session_state.enhanced_resume:
    st.divider()
    st.subheader("🔍 Skills Gap Analysis")
    st.caption("Based on your enhanced resume vs the JD requirements.")

    raw_gap = st.session_state.skills_gap
    def parse_gap_section(raw, label):
        m = re.search(label + r":\s*(.+?)(?=\n[A-Z]+:|$)", raw, re.S)
        if not m: return []
        return [l.strip().lstrip("-• ").strip() for l in m.group(1).strip().split("\n")
                if l.strip() and l.strip() not in ("-", "•")]

    have    = parse_gap_section(raw_gap, "HAVE")
    missing = parse_gap_section(raw_gap, "MISSING")
    irrel   = parse_gap_section(raw_gap, "IRRELEVANT")

    col_h, col_m, col_i = st.columns(3)
    with col_h:
        st.markdown("**✅ You Have**")
        for s in have:
            st.markdown(f"- {s}")
        if not have: st.caption("None identified")
    with col_m:
        st.markdown("**❌ Missing / Gaps**")
        for s in missing:
            st.markdown(f"- {s}")
        if not missing: st.caption("None identified")
    with col_i:
        st.markdown("**⚪ Irrelevant to Role**")
        for s in irrel:
            st.markdown(f"- {s}")
        if not irrel: st.caption("None identified")

# ── ATS Score button — placed after output and revisions to avoid reset ───────


def parse_ats(raw: str) -> dict:
    def pf(lbl):
        m = re.search(rf"{lbl}:\s*([\d.]+)", raw)
        return float(m.group(1)) if m else None
    ms = re.search(r"SUMMARY:\s*(.+?)(?=IMPROVEMENTS:|$)", raw, re.S)
    mi = re.search(r"IMPROVEMENTS:\s*(.+?)$", raw, re.S)
    return {"total":pf("TOTAL_SCORE"),"kw":pf("KEYWORD_MATCH"),"exp":pf("RELEVANCE_OF_EXPERIENCE"),
            "qual":pf("QUALIFICATIONS_MATCH"),"clarity":pf("RESUME_CLARITY"),"formatting":pf("ATS_FORMATTING"),
            "summary":ms.group(1).strip() if ms else "","improve":mi.group(1).strip() if mi else ""}

def badge(total, label):
    ok = total >= 60
    c = "#2d5a3d" if ok else "#c0392b"; bg = "#e8f2eb" if ok else "#fde8e8"
    return f'<div style="text-align:center;background:{bg};border:2px solid {c};border-radius:12px;padding:12px 20px"><div style="font-size:11px;color:{c};font-weight:600;margin-bottom:4px">{label}</div><div style="font-size:2.2rem;font-weight:800;color:{c};line-height:1">{int(total)}</div><div style="font-size:11px;color:{c}">/ 100</div><div style="font-size:11px;font-weight:700;color:{c};margin-top:4px">{"✅ PASS" if ok else "❌ BELOW PASSING"}</div></div>'

def bars(scores, base=None):
    cats = [("Keyword Match","kw",30),("Relevance of Experience","exp",25),
            ("Qualifications Match","qual",20),("Resume Clarity","clarity",15),("ATS Formatting","formatting",10)]
    html = ""
    for name, key, mx in cats:
        s = scores.get(key)
        if s is None: continue
        pct = int(s/mx*100)
        bc = "#2d5a3d" if pct>=60 else "#e67e22" if pct>=40 else "#c0392b"
        delta = ""
        if base and (b := base.get(key)) is not None:
            d = s - b; arrow = "▲" if d>0 else ("▼" if d<0 else "–")
            dc = "#2d5a3d" if d>0 else ("#c0392b" if d<0 else "#888")
            delta = f'<span style="color:{dc};font-size:12px;margin-left:8px">{arrow} {abs(d):.0f}</span>'
        html += f'<div style="margin-bottom:12px"><div style="display:flex;justify-content:space-between;font-size:13px;margin-bottom:3px"><span>{name}</span><span style="font-weight:600">{int(s)}/{mx}{delta}</span></div><div style="background:#e0dbd3;border-radius:6px;height:10px"><div style="width:{pct}%;background:{bc};border-radius:6px;height:10px"></div></div></div>'
    return html

st.markdown('<div class="warn-box">⚠️ <strong>Always review the enhanced resume carefully.</strong> AI may occasionally embellish skills or experience — validate before applying.</div>', unsafe_allow_html=True)

# ── ATS Score button + display ─────────────────────────────────────────────────
if st.session_state.enhanced_resume:
    st.divider()

    if not st.session_state.ats_result:
        # No result yet — show run button only
        col_ats_btn, col_ats_info = st.columns([1, 5])
        with col_ats_btn:
            run_ats = st.button("🎯 Run ATS Score", use_container_width=True,
                                help="Score the current enhanced resume against the JD.")
        with col_ats_info:
            st.caption("Click to score your enhanced resume — including after revisions.")

        if run_ats and not llm_ready():
            st.error("No AI provider/model selected — check the sidebar.")
        elif run_ats:
            current_resume = st.session_state.edited_resume or st.session_state.enhanced_resume
            scorer = f"{LLM_CFG['provider']} / {LLM_CFG['fast_model']}"
            cache_key = hashlib.md5(
                (current_resume + st.session_state.jd_text + str(st.session_state.revision_version) + scorer).encode()
            ).hexdigest()
            if cache_key in st.session_state.ats_score_cache:
                st.session_state.ats_result = st.session_state.ats_score_cache[cache_key]
                st.session_state.ats_model = scorer
                st.rerun()
            else:
                with st.spinner("Scoring enhanced resume…"):
                    try:
                        result = llm([{"role":"system","content":ENHANCED_ATS_PROMPT},
                                      {"role":"user","content":(
                                          f"## Enhanced Resume\n\n{current_resume}"
                                          f"\n\n---\n\n## JD\n\n{st.session_state.jd_text}"
                                      )}],
                                     fast=True, temperature=ENHANCED_ATS_TEMP, max_tokens=600)
                        st.session_state.ats_result = result
                        st.session_state.ats_model = scorer
                        st.session_state.ats_score_cache[cache_key] = result
                        st.rerun()
                    except openai.APITimeoutError:
                        st.error("❌ ATS scoring timed out — please try again.")
                    except Exception as e:
                        st.error(f"❌ ATS scoring failed: {e}")

    else:
        # Result exists — show re-run button + results, never alongside initial button
        col_ats_btn, col_ats_info = st.columns([1, 5])
        with col_ats_btn:
            rerun_ats = st.button("🔄 Re-run ATS Score", use_container_width=True,
                                  help="Re-score after revisions.")
        with col_ats_info:
            st.caption("ATS score shown below. Re-run after revisions to re-evaluate.")
        if rerun_ats:
            st.session_state.ats_result = None
            st.session_state.ats_score_cache = {}
            st.rerun()

        es = parse_ats(st.session_state.ats_result)
        bs = parse_ats(st.session_state.baseline_ats_result) if st.session_state.baseline_ats_result else None
        if es["total"] is not None:
            if bs and bs["total"] is not None:
                imp = es["total"] - bs["total"]
                ic = "#2d5a3d" if imp > 0 else "#c0392b"
                sign = "+" if imp > 0 else ""
                c1, c2, c3 = st.columns(3)
                c1.markdown(badge(bs["total"], "BASELINE RESUME"), unsafe_allow_html=True)
                c2.markdown(badge(es["total"], "ENHANCED RESUME"), unsafe_allow_html=True)
                c3.markdown(f'<div style="text-align:center;border:2px solid {ic};border-radius:12px;padding:12px 20px"><div style="font-size:11px;color:{ic};font-weight:600;margin-bottom:4px">IMPROVEMENT</div><div style="font-size:2.2rem;font-weight:800;color:{ic};line-height:1">{sign}{imp:.0f}</div><div style="font-size:11px;color:{ic}">points</div></div>', unsafe_allow_html=True)
                if (st.session_state.baseline_ats_model and st.session_state.ats_model
                        and st.session_state.baseline_ats_model != st.session_state.ats_model):
                    st.caption(f"⚠️ Baseline was scored by {st.session_state.baseline_ats_model} and the enhanced "
                               f"resume by {st.session_state.ats_model} — different models score differently, "
                               "so the improvement figure is only a rough guide.")
            else:
                st.markdown(badge(es["total"], "ENHANCED RESUME"), unsafe_allow_html=True)
            st.write("")
            with st.expander(f"📊 View score breakdown{' *(▲/▼ vs baseline)*' if bs else ''}"):
                st.markdown(bars(es, bs), unsafe_allow_html=True)
                if es["summary"]: st.markdown("**Summary**"); st.info(es["summary"])
                if es["improve"]:
                    st.markdown("**How to improve**")
                    for l in es["improve"].split("\n"):
                        l = l.strip().lstrip("-•* ").strip()
                        if l: st.markdown(f"- {l}")
        else:
            st.warning("⚠️ The scoring model didn't return scores in the expected format. "
                       "Try Re-run, or pick a stronger scoring model in the sidebar.")



if st.session_state.enhanced_resume:
    with st.expander("🔍 Debug — verify ATS inputs & raw scores"):
        if LLM_CFG:
            st.caption(f"Provider: {LLM_CFG['provider']} · writing model: {LLM_CFG['model']} · "
                       f"scoring model: {LLM_CFG['fast_model']}")
        st.markdown("**Baseline resume (full text sent to ATS scorer):**")
        st.text_area("Baseline", value=st.session_state.resume_text, height=200,
                     key="debug_baseline", disabled=True, label_visibility="collapsed")
        st.caption(f"{len(st.session_state.resume_text.split())} words")

        st.markdown("**Enhanced resume (full text sent to ATS scorer):**")
        st.text_area("Enhanced", value=st.session_state.enhanced_resume, height=200,
                     key="debug_enhanced", disabled=True, label_visibility="collapsed")
        st.caption(f"{len(st.session_state.enhanced_resume.split())} words")

        st.markdown("**Raw baseline ATS response:**")
        st.code(st.session_state.baseline_ats_result or "None", language="text")
        st.markdown("**Raw enhanced ATS response:**")
        st.code(st.session_state.ats_result or "None", language="text")

        same = st.session_state.resume_text == st.session_state.enhanced_resume
        (st.error if same else st.success)(
            "⚠️ Resumes are IDENTICAL — enhancer may not have run correctly." if same
            else f"✅ Resumes differ — enhancement applied ({len(st.session_state.resume_text.split())} → {len(st.session_state.enhanced_resume.split())} words)."
        )

# ── Follow-up Revision Chat ────────────────────────────────────────────────────
if st.session_state.enhanced_resume:
    st.divider()
    st.subheader("💬 Request Further Revisions")
    st.caption("Ask for specific changes to the enhanced resume. Each revision updates the output above.")

    if st.session_state.revision_history:
        with st.expander(f"📜 Revision history ({len(st.session_state.revision_history)} revision(s))", expanded=False):
            for i, entry in enumerate(st.session_state.revision_history, 1):
                st.markdown(f"**Revision {i}:** {entry['request']}")
                st.caption(f"Applied at {entry['timestamp']}")

    revision_input = st.text_area(
        "What would you like to change?",
        placeholder="e.g. 'Make the summary more focused on automation engineering' or 'Strengthen the bullet points in the most recent job'",
        height=100,
        key="revision_input",
    )

    col_rev, col_reset_rev, _ = st.columns([1, 1, 4])
    with col_rev:
        apply_revision = st.button("✏️ Apply Revision", use_container_width=True,
                                   disabled=not revision_input.strip() or not llm_ready())
    with col_reset_rev:
        if st.button("↺ Clear History", use_container_width=True):
            st.session_state.revision_history = []
            st.session_state.edited_resume = st.session_state.enhanced_resume

    if apply_revision and revision_input.strip():
        with st.spinner("Applying revision…"):
            try:
                current = st.session_state.edited_resume or st.session_state.enhanced_resume
                rev_content = (
                    f"## Current Resume\n\n{current}"
                    f"\n\n---\n\n## Job Description\n\n{st.session_state.jd_text}"
                    f"\n\n---\n\n## Revision Request\n\n{revision_input}"
                    f"\n\n---\n\n## IMPORTANT INSTRUCTION\n\n"
                    f"Integrate all requested changes directly into the resume body. "
                    f"Do not append any separate section, note, or suggestion at the end. "
                    f"If a skill is requested, add it to the Skills section. "
                    f"Output only the complete revised resume with all changes integrated."
                )
                revised = llm([{"role": "system", "content": REVISION_PROMPT},
                               {"role": "user", "content": rev_content}],
                              temperature=0.5, max_tokens=3500)

                rev_num = len([v for v in st.session_state.resume_versions if "Revision" in v["label"]]) + 1
                ts_rev = datetime.now().strftime("%Y-%m-%d %H:%M")
                st.session_state.resume_versions.append({
                    "label": f"Revision {rev_num}: {revision_input[:50]}{'…' if len(revision_input) > 50 else ''}",
                    "content": revised,
                    "timestamp": ts_rev,
                })
                new_idx = len(st.session_state.resume_versions) - 1
                st.session_state.active_version_idx = new_idx
                st.session_state.enhanced_resume = revised
                st.session_state.edited_resume = revised
                st.session_state._last_enhanced = revised
                st.session_state.revision_version += 1
                st.session_state.ats_result = None
                st.session_state.ats_score_cache = {}  # Clear so revision gets fresh score

                st.session_state.revision_history.append({
                    "request": revision_input,
                    "timestamp": ts_rev,
                })

                try:
                    gap_content = f"## Resume\n\n{revised}\n\n---\n\n## JD\n\n{st.session_state.jd_text}"
                    st.session_state.skills_gap = llm(
                        [{"role":"system","content":SKILLS_GAP_PROMPT},{"role":"user","content":gap_content}],
                        fast=True, temperature=0.0, max_tokens=800)
                except Exception as e:
                    st.caption(f"⚠️ Skills gap update failed after revision: {e}")

                st.rerun()
            except openai.APITimeoutError:
                st.error("❌ Revision timed out — the model took too long. Please try again.")
            except openai.APIConnectionError:
                st.error(f"❌ Could not reach {LLM_CFG['provider']} — check it is running / reachable.")
            except Exception as e:
                st.error(f"❌ Revision failed: {e}")

    if st.session_state.revision_history:
        st.caption(f"✅ Revision applied — {len(st.session_state.revision_history)} revision(s) made.")

# ── Application Tracker ────────────────────────────────────────────────────────
st.divider(); st.subheader("📋 Application Tracker")
st.caption("Each enhancement adds a row. Edit status and details in the table; add or delete rows directly. "
           + ("History is saved automatically between sessions." if TRACKER_AUTOSAVE
              else "To keep your history, download the Excel file and import it next time."))

# ── Carry on from previous sessions: import an earlier export ──
with st.expander("📂 Import previous applications (Excel or CSV)"):
    st.caption("Upload a tracker you downloaded before. Its rows are combined with this session's — "
               "duplicates (same company, role and date) are kept once.")
    imp = st.file_uploader("Previous tracker", type=["xlsx", "csv"], accept_multiple_files=True,
                           label_visibility="collapsed", key=f"tracker_import_{st.session_state.tracker_version}")
    if imp and st.button("➕ Combine with current tracker"):
        try:
            before = len(current_tracker())
            merged = merge_trackers(*[read_tracker_upload(f) for f in imp], current_tracker())
            set_tracker(merged)
            st.success(f"✅ Imported — tracker now has {len(merged)} application(s) (was {before}).")
            st.rerun()
        except Exception as e:
            st.error(f"❌ Could not read that file: {e}")

# ── Editable table ──
# The table's own state is keyed by tracker_version, so changes made outside the table
# (new enhancement, import, clear) refresh it without re-applying old edits.
edited_t = st.data_editor(
    normalise_tracker(st.session_state.tracker_edit),
    num_rows="dynamic", use_container_width=True, hide_index=False,
    column_config={
        "Company": st.column_config.TextColumn("Company"),
        "Role": st.column_config.TextColumn("Role"),
        "Date Applied": st.column_config.DateColumn("Date Applied", format="YYYY-MM-DD", default=date.today()),
        "Status": st.column_config.SelectboxColumn("Status", options=STATUS_OPTIONS, default="Not Applied", required=True),
    }, key=f"tracker_editor_{st.session_state.tracker_version}")

current = normalise_tracker(edited_t)
_sig = hashlib.md5(current.to_csv(index=False).encode()).hexdigest()
if st.session_state.get("_tracker_sig") != _sig:
    st.session_state._tracker_sig = _sig
    save_local_tracker(current)
st.session_state.tracker_edit_live = current

counts = current["Status"].value_counts()
st.caption(f"**{len(current)}** application(s) · " +
           " · ".join(f"{s}: {int(counts.get(s, 0))}" for s in STATUS_OPTIONS if counts.get(s, 0)))

c_dl, c_clear, _ = st.columns([1, 1, 4])
with c_dl:
    st.download_button("⬇ Export Excel", data=export_excel(current),
                       file_name=f"applications_{datetime.now().strftime('%Y-%m-%d_%H%M')}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True, disabled=current.empty)
with c_clear:
    if st.button("🗑️ Clear tracker", use_container_width=True, disabled=current.empty):
        st.session_state._confirm_clear_tracker = True
st.caption("Tip: the Excel file also opens in Google Sheets — upload it to Google Drive and open it there.")

with st.expander("🟢 Google Sheets — save / load your tracker"):
    if gspread is None:
        st.info("Google Sheets needs one more package: `pip install gspread`, then restart the app.")
    elif not gs_configured():
        st.info("Google Sheets isn't set up yet. Add GOOGLE_SERVICE_ACCOUNT_FILE (or GOOGLE_SERVICE_ACCOUNT_JSON) "
                "to .env — see .env.example for the steps.")
    else:
        email = gs_service_email()
        st.caption(f"1. Create a Google Sheet (any name) and share it with **{email or 'the service account'}** as **Editor**.  \n"
                   "2. Paste its link below.  \n"
                   "3. **Save** writes this tracker to the sheet (keeping rows that are only in the sheet). "
                   "**Load** pulls the sheet into this tracker, e.g. after you've updated statuses in Google Sheets.")
        if "gsheet_url" not in st.session_state:
            st.session_state.gsheet_url = GOOGLE_SHEET_URL
        gurl = st.text_input("Google Sheet link", key="gsheet_url",
                             placeholder="https://docs.google.com/spreadsheets/d/…").strip()
        g1, g2, _ = st.columns([1, 1, 4])
        save_gs = g1.button("⬆ Save to Sheet", use_container_width=True, disabled=not gurl)
        load_gs = g2.button("⬇ Load from Sheet", use_container_width=True, disabled=not gurl)
        if save_gs:
            with st.spinner("Saving to Google Sheets…"):
                try:
                    sheet_rows = gs_pull(gurl)
                    merged = merge_trackers(sheet_rows, current)          # app wins on duplicates
                    url = gs_push(merged, gurl)
                    set_tracker(merged)
                    st.session_state._gs_msg = f"✅ Saved {len(merged)} application(s) to [Google Sheets]({url})."
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ {gs_error_message(e)}")
        if load_gs:
            with st.spinner("Loading from Google Sheets…"):
                try:
                    sheet_rows = gs_pull(gurl)
                    merged = merge_trackers(current, sheet_rows)          # sheet wins on duplicates
                    set_tracker(merged)
                    st.session_state._gs_msg = f"✅ Loaded {len(sheet_rows)} row(s) from the sheet — tracker now has {len(merged)}."
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ {gs_error_message(e)}")
        if st.session_state.get("_gs_msg"):
            st.success(st.session_state.pop("_gs_msg"))

if st.session_state.get("_confirm_clear_tracker"):
    st.warning("Remove all rows from the tracker" + (f" and the saved history file `{TRACKER_FILE}`" if TRACKER_AUTOSAVE else "")
               + "? Export the Excel first if you want to keep them.")
    y, n, _ = st.columns([1, 1, 4])
    if y.button("Yes, clear", use_container_width=True):
        st.session_state._confirm_clear_tracker = False
        set_tracker(empty_tracker()); st.rerun()
    if n.button("Cancel", use_container_width=True):
        st.session_state._confirm_clear_tracker = False; st.rerun()