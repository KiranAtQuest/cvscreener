import os
import io
import re
import json
import anthropic
import streamlit as st
import pdfplumber
from docx import Document

st.set_page_config(page_title="CV Screener – Quest", page_icon="🔍", layout="wide")

# ── Header with logo ─────────────────────────────────────────────────────────

col_title, col_logo = st.columns([4, 1])
with col_title:
    st.title("🔍 CV Screener")
    st.caption("Powered by Claude AI · Quest Alliance")
with col_logo:
    st.image(
        "https://questalliance.net/wp-content/uploads/2023/03/Quest-logo-new.png",
        width=160,
    )

# ── Google Drive helpers ──────────────────────────────────────────────────────

@st.cache_resource
def get_drive_service():
    try:
        from googleapiclient.discovery import build
        from google.oauth2 import service_account

        creds_raw = st.secrets.get("GOOGLE_SERVICE_ACCOUNT", "")
        if not creds_raw:
            return None
        creds_dict = json.loads(creds_raw)
        creds = service_account.Credentials.from_service_account_info(
            creds_dict,
            scopes=["https://www.googleapis.com/auth/drive.readonly"],
        )
        return build("drive", "v3", credentials=creds, cache_discovery=False)
    except Exception:
        return None


def extract_folder_id(url: str) -> str:
    match = re.search(r"/folders/([a-zA-Z0-9_-]+)", url)
    return match.group(1) if match else url.strip()


def list_drive_files(folder_id: str):
    service = get_drive_service()
    if not service:
        return []
    results = service.files().list(
        q=f"'{folder_id}' in parents and trashed=false",
        fields="files(id, name, mimeType)",
        pageSize=100,
    ).execute()
    return results.get("files", [])


def download_drive_file(file_id: str, name: str, mime_type: str) -> bytes:
    from googleapiclient.http import MediaIoBaseDownload

    service = get_drive_service()
    if mime_type == "application/vnd.google-apps.document":
        request = service.files().export_media(
            fileId=file_id,
            mimeType="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    else:
        request = service.files().get_media(fileId=file_id)

    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return buf.getvalue()


# ── File parsing helpers ──────────────────────────────────────────────────────

def extract_text_from_pdf(file_bytes: bytes) -> str:
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        return "\n".join(page.extract_text() or "" for page in pdf.pages)


def extract_text_from_docx(file_bytes: bytes) -> str:
    doc = Document(io.BytesIO(file_bytes))
    return "\n".join(para.text for para in doc.paragraphs)


def parse_bytes(file_bytes: bytes, name: str) -> str:
    name_lower = name.lower()
    if name_lower.endswith(".pdf"):
        return extract_text_from_pdf(file_bytes)
    elif name_lower.endswith(".docx") or name_lower.endswith(".doc"):
        return extract_text_from_docx(file_bytes)
    else:
        return file_bytes.decode("utf-8", errors="replace")


def extract_file_text(uploaded_file) -> str:
    return parse_bytes(uploaded_file.read(), uploaded_file.name)


# ── Claude helpers ────────────────────────────────────────────────────────────

def build_prompt(jd: str, competencies: str, cvs: dict[str, str]) -> str:
    cv_block = ""
    for idx, (name, text) in enumerate(cvs.items(), 1):
        cv_block += f"\n---\nCV #{idx} – Filename: {name}\n{text}\n"

    return f"""You are an expert HR screener for Quest Alliance, an NGO focused on youth skilling.

## Job Description
{jd}

## Required Skill Competencies
{competencies}

## Candidate CVs
{cv_block}

## Task
Carefully review each CV against the job description and skill competencies.
Return a JSON array (and nothing else) with one object per candidate, sorted from best match to worst.
Each object must have these exact keys:
- "rank": integer starting at 1
- "filename": the CV filename exactly as provided
- "match_score": integer 0–100 representing overall fit
- "shortlisted": boolean (true if score >= 60)
- "strengths": array of strings, key matching strengths
- "gaps": array of strings, notable gaps or concerns
- "reasoning": a concise paragraph (3–5 sentences) explaining the overall assessment

Return ONLY the JSON array, no markdown fences, no preamble."""


def screen_cvs(jd: str, competencies: str, cvs: dict[str, str]):
    api_key = os.environ.get("ANTHROPIC_API_KEY") or st.secrets.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        st.error("ANTHROPIC_API_KEY not found. Set it as an environment variable or in Streamlit secrets.")
        return None

    client = anthropic.Anthropic(api_key=api_key)
    prompt = build_prompt(jd, competencies, cvs)

    placeholder = st.empty()
    placeholder.info("Analysing CVs with Claude… this may take up to a minute.")

    collected = []
    with client.messages.stream(
        model="claude-opus-4-8",
        max_tokens=4096,
        thinking={"type": "adaptive"},
        messages=[{"role": "user", "content": prompt}],
    ) as stream:
        for text in stream.text_stream:
            collected.append(text)

    raw = "".join(collected).strip()
    placeholder.empty()
    return raw


def render_results(raw_json: str):
    try:
        results = json.loads(raw_json)
    except json.JSONDecodeError:
        st.error("Could not parse Claude's response as JSON. Raw output:")
        st.code(raw_json)
        return

    shortlisted = [r for r in results if r.get("shortlisted")]
    not_shortlisted = [r for r in results if not r.get("shortlisted")]

    st.subheader(f"✅ Shortlisted Candidates ({len(shortlisted)})")
    for r in shortlisted:
        with st.expander(f"#{r['rank']} · {r['filename']}  —  Score: {r['match_score']}/100", expanded=True):
            st.markdown(f"**Overall Assessment**\n\n{r['reasoning']}")
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**Strengths**")
                for s in r.get("strengths", []):
                    st.markdown(f"- {s}")
            with c2:
                st.markdown("**Gaps / Concerns**")
                for g in r.get("gaps", []):
                    st.markdown(f"- {g}")

    if not_shortlisted:
        st.subheader(f"❌ Not Shortlisted ({len(not_shortlisted)})")
        for r in not_shortlisted:
            with st.expander(f"#{r['rank']} · {r['filename']}  —  Score: {r['match_score']}/100"):
                st.markdown(f"**Overall Assessment**\n\n{r['reasoning']}")
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("**Strengths**")
                    for s in r.get("strengths", []):
                        st.markdown(f"- {s}")
                with c2:
                    st.markdown("**Gaps / Concerns**")
                    for g in r.get("gaps", []):
                        st.markdown(f"- {g}")


# ── Sidebar: competencies ────────────────────────────────────────────────────

with st.sidebar:
    st.header("Skill Competencies")
    competencies_input = st.text_area(
        "competencies",
        height=300,
        placeholder="List the required skills and competencies, one per line or as a paragraph… (optional)",
        label_visibility="collapsed",
    )

    drive_configured = get_drive_service() is not None
    if not drive_configured:
        st.divider()
        st.caption("💡 Google Drive integration not configured. Add GOOGLE_SERVICE_ACCOUNT to Streamlit Secrets to enable it.")

# ── Step 1: Job Description ──────────────────────────────────────────────────

st.subheader("Step 1 — Job Description")

jd_source = st.radio(
    "Source",
    ["Paste text", "Upload file", "Import from Google Drive"],
    horizontal=True,
    label_visibility="collapsed",
)

jd_input = ""

if jd_source == "Paste text":
    jd_input = st.text_area(
        "jd_paste",
        height=220,
        placeholder="Paste the full job description here…",
        label_visibility="collapsed",
    )

elif jd_source == "Upload file":
    jd_file = st.file_uploader(
        "Upload the Job Description file",
        type=["pdf", "docx", "doc", "txt"],
        key="jd_file",
    )
    if jd_file:
        try:
            jd_input = extract_file_text(jd_file)
            st.success(f"✅ Loaded: {jd_file.name}")
        except Exception as e:
            st.error(f"Could not read file: {e}")

else:  # Google Drive
    if not get_drive_service():
        st.warning("Google Drive is not configured yet. See setup instructions below.")
    else:
        jd_folder_url = st.text_input("Google Drive folder or file URL", placeholder="https://drive.google.com/drive/folders/…")
        if jd_folder_url:
            folder_id = extract_folder_id(jd_folder_url)
            with st.spinner("Listing files…"):
                drive_files = list_drive_files(folder_id)
            if drive_files:
                jd_pick = st.selectbox(
                    "Select the JD file",
                    options=drive_files,
                    format_func=lambda f: f["name"],
                )
                if st.button("Load selected JD", key="load_jd"):
                    with st.spinner(f"Downloading {jd_pick['name']}…"):
                        try:
                            file_bytes = download_drive_file(jd_pick["id"], jd_pick["name"], jd_pick["mimeType"])
                            jd_input = parse_bytes(file_bytes, jd_pick["name"])
                            st.session_state["jd_input_gdrive"] = jd_input
                            st.success(f"✅ Loaded: {jd_pick['name']}")
                        except Exception as e:
                            st.error(f"Error downloading file: {e}")
            else:
                st.warning("No files found in that folder.")

        if "jd_input_gdrive" in st.session_state:
            jd_input = st.session_state["jd_input_gdrive"]

# ── Step 2: CVs ──────────────────────────────────────────────────────────────

st.divider()
st.subheader("Step 2 — Candidate CVs")

cv_source = st.radio(
    "CV Source",
    ["Upload files", "Import from Google Drive"],
    horizontal=True,
    label_visibility="collapsed",
)

cvs: dict[str, str] = {}

if cv_source == "Upload files":
    uploaded_files = st.file_uploader(
        "Upload one or more CVs (PDF, DOCX, or TXT)",
        type=["pdf", "docx", "doc", "txt"],
        accept_multiple_files=True,
    )
    if uploaded_files:
        st.caption(f"{len(uploaded_files)} file(s) ready: {', '.join(f.name for f in uploaded_files)}")

else:  # Google Drive
    uploaded_files = []
    if not get_drive_service():
        st.warning("Google Drive is not configured yet. See setup instructions below.")
    else:
        cv_folder_url = st.text_input("Google Drive CVs folder URL", placeholder="https://drive.google.com/drive/folders/…", key="cv_folder")
        if cv_folder_url:
            folder_id = extract_folder_id(cv_folder_url)
            with st.spinner("Listing files…"):
                cv_drive_files = list_drive_files(folder_id)
            if cv_drive_files:
                selected_cvs = st.multiselect(
                    "Select CVs to screen",
                    options=cv_drive_files,
                    default=cv_drive_files,
                    format_func=lambda f: f["name"],
                )
                if st.button("Load selected CVs", key="load_cvs"):
                    progress = st.progress(0, text="Downloading CVs…")
                    for i, f in enumerate(selected_cvs):
                        try:
                            file_bytes = download_drive_file(f["id"], f["name"], f["mimeType"])
                            cvs[f["name"]] = parse_bytes(file_bytes, f["name"])
                        except Exception as e:
                            st.warning(f"Could not load {f['name']}: {e}")
                        progress.progress((i + 1) / len(selected_cvs), text=f"Loaded {i+1}/{len(selected_cvs)}")
                    progress.empty()
                    st.session_state["cvs_gdrive"] = cvs
                    st.success(f"✅ {len(cvs)} CVs loaded from Google Drive")

            else:
                st.warning("No files found in that folder.")

        if "cvs_gdrive" in st.session_state:
            cvs = st.session_state["cvs_gdrive"]
            st.caption(f"{len(cvs)} CVs loaded: {', '.join(cvs.keys())}")

# ── Screen button ─────────────────────────────────────────────────────────────

st.divider()

if cv_source == "Upload files":
    cvs_ready = bool(uploaded_files)
else:
    cvs_ready = bool(cvs)

screen_btn = st.button("🚀 Screen CVs", type="primary", disabled=not (jd_input and cvs_ready))

if screen_btn:
    if not competencies_input.strip():
        st.warning("No skill competencies provided — Claude will rely solely on the job description.")

    if cv_source == "Upload files":
        with st.spinner("Extracting text from CVs…"):
            for f in uploaded_files:
                try:
                    cvs[f.name] = extract_file_text(f)
                except Exception as exc:
                    st.warning(f"Could not parse {f.name}: {exc}")

    if not cvs:
        st.error("No CV text could be extracted. Please check your files.")
    else:
        raw = screen_cvs(jd_input, competencies_input, cvs)
        if raw:
            st.divider()
            st.header("Screening Results")
            render_results(raw)

# ── Google Drive setup instructions ──────────────────────────────────────────

if not get_drive_service():
    with st.expander("⚙️ How to set up Google Drive integration"):
        st.markdown("""
1. Go to [Google Cloud Console](https://console.cloud.google.com/) → create a project
2. Enable **Google Drive API**
3. Go to **IAM & Admin → Service Accounts** → create a service account
4. Click the service account → **Keys → Add Key → JSON** → download the file
5. In **Streamlit Cloud → your app → Settings → Secrets**, add:
```toml
GOOGLE_SERVICE_ACCOUNT = '''
{ paste the entire contents of the downloaded JSON key file here }
'''
```
6. Share your Google Drive folders with the service account email (e.g. `myapp@project.iam.gserviceaccount.com`) — **Viewer** access is enough
7. Reboot the app — the "Import from Google Drive" option will activate
        """)
