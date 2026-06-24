import os
import io
import json
import anthropic
import streamlit as st
import pdfplumber
from docx import Document

st.set_page_config(page_title="CV Screener – Quest", page_icon="🔍", layout="wide")

st.title("🔍 CV Screener")
st.caption("Powered by Claude AI · Quest Alliance")


def extract_text_from_pdf(file_bytes: bytes) -> str:
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        return "\n".join(page.extract_text() or "" for page in pdf.pages)


def extract_text_from_docx(file_bytes: bytes) -> str:
    doc = Document(io.BytesIO(file_bytes))
    return "\n".join(para.text for para in doc.paragraphs)


def extract_cv_text(uploaded_file) -> str:
    file_bytes = uploaded_file.read()
    name = uploaded_file.name.lower()
    if name.endswith(".pdf"):
        return extract_text_from_pdf(file_bytes)
    elif name.endswith(".docx") or name.endswith(".doc"):
        return extract_text_from_docx(file_bytes)
    else:
        return file_bytes.decode("utf-8", errors="replace")


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

    result_placeholder = st.empty()
    result_placeholder.info("Analysing CVs with Claude…")

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
    result_placeholder.empty()
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
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**Strengths**")
                for s in r.get("strengths", []):
                    st.markdown(f"- {s}")
            with col2:
                st.markdown("**Gaps / Concerns**")
                for g in r.get("gaps", []):
                    st.markdown(f"- {g}")

    if not_shortlisted:
        st.subheader(f"❌ Not Shortlisted ({len(not_shortlisted)})")
        for r in not_shortlisted:
            with st.expander(f"#{r['rank']} · {r['filename']}  —  Score: {r['match_score']}/100"):
                st.markdown(f"**Overall Assessment**\n\n{r['reasoning']}")
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("**Strengths**")
                    for s in r.get("strengths", []):
                        st.markdown(f"- {s}")
                with col2:
                    st.markdown("**Gaps / Concerns**")
                    for g in r.get("gaps", []):
                        st.markdown(f"- {g}")


# ── Sidebar inputs ──────────────────────────────────────────────────────────

with st.sidebar:
    st.header("Role Details")

    jd_mode = st.radio("Job Description", ["Paste text", "Upload file"], horizontal=True)

    if jd_mode == "Paste text":
        jd_input = st.text_area(
            "JD text",
            height=220,
            placeholder="Paste the full job description here…",
            label_visibility="collapsed",
        )
    else:
        jd_file = st.file_uploader(
            "Upload JD (PDF, DOCX, or TXT)",
            type=["pdf", "docx", "doc", "txt"],
            key="jd_file",
        )
        jd_input = ""
        if jd_file:
            try:
                jd_input = extract_cv_text(jd_file)
                st.success(f"Loaded: {jd_file.name}")
            except Exception as e:
                st.error(f"Could not read file: {e}")

    competencies_input = st.text_area(
        "Skill Competencies (optional)",
        height=150,
        placeholder="List the required skills and competencies, one per line or as a paragraph…",
    )

# ── Main area ────────────────────────────────────────────────────────────────

uploaded_files = st.file_uploader(
    "Upload CVs (PDF, DOCX, or TXT)",
    type=["pdf", "docx", "doc", "txt"],
    accept_multiple_files=True,
)

if uploaded_files:
    st.caption(f"{len(uploaded_files)} file(s) uploaded: {', '.join(f.name for f in uploaded_files)}")

screen_btn = st.button("🚀 Screen CVs", type="primary", disabled=not (jd_input and uploaded_files))

if screen_btn:
    if not competencies_input.strip():
        st.warning("No skill competencies provided — Claude will rely solely on the job description.")

    with st.spinner("Extracting text from CVs…"):
        cvs: dict[str, str] = {}
        for f in uploaded_files:
            try:
                cvs[f.name] = extract_cv_text(f)
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
