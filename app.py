import os, io, re, json
from datetime import datetime
import anthropic
import streamlit as st
import pdfplumber
from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable, Table, TableStyle
from reportlab.lib.enums import TA_LEFT

st.set_page_config(page_title="CV Screener – Quest", page_icon="🔍", layout="wide",
                   initial_sidebar_state="collapsed")

# ── Global CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Work+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');

/* Hide Streamlit chrome */
#MainMenu, footer, header { visibility: hidden; }
[data-testid="stSidebar"] { display: none; }
.block-container { padding: 0 !important; max-width: 100% !important; }
[data-testid="stAppViewContainer"] { background: #F4F7FA; }

/* Shared */
* { font-family: 'Work Sans', -apple-system, sans-serif; }
.qs-header {
  height:60px; background:#fff; border-bottom:1px solid #E4E9EF;
  display:flex; align-items:center; justify-content:space-between;
  padding:0 28px; position:sticky; top:0; z-index:100;
}
.qs-logo {
  width:30px;height:30px;border-radius:9px;background:#0075BC;color:#fff;
  display:inline-flex;align-items:center;justify-content:center;
  font-weight:800;font-size:16px;position:relative;flex:none;
}
.qs-logo-dot {
  width:10px;height:10px;border-radius:50%;background:#F7941D;
  position:absolute;bottom:-2px;right:-2px;border:2px solid #fff;
}
.qs-card {
  background:#fff; border:1px solid #E4E9EF; border-radius:14px;
  padding:22px 24px; margin-bottom:16px;
}
.qs-label { font-weight:700; font-size:16px; margin-bottom:14px; }
.qs-sublabel { font:500 13px 'Work Sans'; color:#9AA1AE; }
.qs-pill {
  display:inline-flex;align-items:center;gap:7px;background:#E3F1FA;
  color:#005A91;border-radius:999px;padding:7px 13px;
  font-weight:600;font-size:13px;margin:4px;
}
.qs-pill-rm { opacity:.5;cursor:pointer;border:none;background:none;color:inherit;font-size:13px; }
.qs-pill-add {
  display:inline-flex;align-items:center;gap:6px;border:1px dashed #C5CCD6;
  color:#5E6675;border-radius:999px;padding:7px 13px;
  font-weight:600;font-size:13px;cursor:pointer;background:none;
}
.qs-upload-zone {
  border:1.5px dashed #B8DCF0;background:repeating-linear-gradient(135deg,#EAF4FB 0 10px,#fff 10px 20px);
  border-radius:11px;display:flex;flex-direction:column;align-items:center;
  justify-content:center;gap:8px;padding:24px;cursor:pointer;text-align:center;
}
.qs-file-chip {
  display:flex;align-items:center;gap:9px;background:#F4F7FA;
  border:1px solid #E4E9EF;border-radius:9px;padding:9px 11px;
}
.qs-file-type-pdf { background:#FEF0DC;color:#C76A0A; }
.qs-file-type-doc { background:#E4EDF7;color:#3251A3; }
.qs-file-icon {
  width:24px;height:24px;border-radius:6px;font-weight:700;font-size:9px;
  display:flex;align-items:center;justify-content:center;flex:none;
}
/* Candidate row */
.qs-cand-row {
  display:flex;align-items:center;gap:16px;background:#fff;
  border:1px solid #E4E9EF;border-radius:12px;padding:14px 18px;
  cursor:pointer;margin-bottom:9px;transition:box-shadow .15s,transform .15s;
}
.qs-cand-row:hover { box-shadow:0 6px 18px rgba(30,40,60,.10); transform:translateY(-1px); }
.qs-avatar {
  width:42px;height:42px;border-radius:50%;color:#fff;
  display:flex;align-items:center;justify-content:center;
  font-weight:700;font-size:15px;flex:none;
}
.qs-score-ring {
  width:54px;height:54px;border-radius:50%;
  display:flex;align-items:center;justify-content:center;flex:none;
}
.qs-score-inner {
  width:42px;height:42px;border-radius:50%;background:#fff;
  display:flex;align-items:center;justify-content:center;
  font-weight:800;font-size:15px;
}
/* Score bar */
.qs-bar-track { height:9px;background:#E4E9EF;border-radius:999px;overflow:hidden;margin-top:7px; }
.qs-bar-fill { height:100%;border-radius:999px; }
/* AI summary */
.qs-ai-box {
  background:#0075BC;border-radius:12px;padding:16px 18px;color:#fff;margin-bottom:18px;
}
/* Evidence card */
.qs-evidence {
  background:#fff;border:1px solid #E4E9EF;border-radius:10px;padding:12px 14px;margin-bottom:10px;
}
.qs-evidence-label { font:600 11px 'JetBrains Mono';color:#0075BC;margin-bottom:4px; }
/* Buttons */
.qs-btn-primary {
  display:inline-flex;align-items:center;gap:11px;background:#0075BC;color:#fff;
  border:none;border-radius:11px;padding:13px 24px;font-weight:700;font-size:15px;
  cursor:pointer;box-shadow:0 4px 14px rgba(0,117,188,.3);
}
.qs-btn-outline {
  display:inline-flex;align-items:center;gap:7px;font:600 13px 'Work Sans';
  color:#0075BC;background:#fff;border:1px solid #B8DCF0;border-radius:9px;
  padding:8px 14px;cursor:pointer;
}
/* Filter chips */
.qs-filter-btn {
  padding:8px 14px;border-radius:9px;font-weight:600;font-size:13px;cursor:pointer;
}
/* Steps */
.qs-step-bar {
  display:flex;align-items:center;padding:18px 36px;background:#fff;
  border-bottom:1px solid #E4E9EF;
}
.qs-step-num {
  width:26px;height:26px;border-radius:50%;font-weight:700;font-size:12px;
  display:flex;align-items:center;justify-content:center;flex:none;
}
.qs-step-active .qs-step-num { background:#0075BC;color:#fff; }
.qs-step-active .qs-step-label { font-weight:700;font-size:13px;color:#0075BC; }
.qs-step-inactive .qs-step-num { background:#fff;border:2px solid #CBD2DC;color:#9AA1AE; }
.qs-step-inactive .qs-step-label { font-weight:600;font-size:13px;color:#9AA1AE; }
.qs-step-line { flex:1;height:2px;margin:0 14px; }
</style>
""", unsafe_allow_html=True)

# ── Helpers ───────────────────────────────────────────────────────────────────

def extract_text_from_pdf(b):
    with pdfplumber.open(io.BytesIO(b)) as p:
        return "\n".join(pg.extract_text() or "" for pg in p.pages)

def extract_text_from_docx(b):
    return "\n".join(para.text for para in Document(io.BytesIO(b)).paragraphs)

def parse_bytes(b, name):
    n = name.lower()
    if n.endswith(".pdf"):   return extract_text_from_pdf(b)
    if n.endswith((".docx",".doc")): return extract_text_from_docx(b)
    return b.decode("utf-8", errors="replace")

def extract_file_text(f): return parse_bytes(f.read(), f.name)

def initials(name):
    parts = name.split()
    return (parts[0][0] + parts[-1][0]).upper() if len(parts) > 1 else parts[0][:2].upper()

AVATAR_COLORS = ["#0075BC","#3251A3","#27AAE1","#444A67","#6B5BA8","#1A7A5E","#C76A0A","#C23A18"]

def avatar_color(name):
    return AVATAR_COLORS[sum(ord(c) for c in name) % len(AVATAR_COLORS)]

def band(score):
    if score >= 85: return "strong",  "#E3F1FA", "#005A91", "#B8DCF0"
    if score >= 65: return "possible", "#FEF0DC", "#C76A0A", "#F7D49A"
    return "weak", "#FDE7DE", "#C23A18", "#F5B8A8"

def score_ring_colors(score):
    if score >= 85: return "#0075BC20", "#0075BC"
    if score >= 65: return "#F7941D20", "#F7941D"
    return "#E8502020", "#E85020"

# ── Claude API ────────────────────────────────────────────────────────────────

def build_prompt(jd, competencies, cvs):
    cv_block = "".join(f"\n---\nCV #{i} – {name}\n{text}\n" for i,(name,text) in enumerate(cvs.items(),1))
    return f"""You are an expert HR screener for Quest Alliance, an NGO focused on youth skilling in India.

## Job Description
{jd}

## Required Skill Competencies
{competencies or "(derive from JD)"}

## Candidate CVs
{cv_block}

## Task
Review each candidate carefully and return a JSON array sorted best-to-worst. Each object MUST have:
- "rank": integer from 1
- "filename": CV filename exactly as given
- "name": candidate's full name (extract from CV)
- "role": their current/most recent role title
- "years": integer years of total experience
- "location": city/state from CV
- "email": email address if present, else ""
- "phone": phone if present, else ""
- "overall": integer 0-100 match score
- "scores": array of 5 integers (one per competency, same order as competencies)
- "competency_labels": array of 5 strings (the competency names scored)
- "shortlisted": true if overall >= 65
- "summary": 2-3 sentence AI summary of the candidate's fit
- "strengths": array of 2-4 short strength strings
- "gaps": array of 1-2 gap strings
- "flag": a one-sentence thing to verify, or null
- "evidence": array of 2-3 objects with "label" (competency name, uppercase) and "text" (direct quote or paraphrase from CV)

Return ONLY the JSON array, no markdown fences."""

def screen_cvs(jd, competencies, cvs):
    api_key = os.environ.get("ANTHROPIC_API_KEY") or st.secrets.get("ANTHROPIC_API_KEY","")
    if not api_key:
        st.error("ANTHROPIC_API_KEY not set."); return None
    client = anthropic.Anthropic(api_key=api_key)
    ph = st.empty()
    ph.info("🔍 Analysing CVs with Claude… this may take up to a minute.")
    chunks = []
    with client.messages.stream(
        model="claude-opus-4-8", max_tokens=8000, thinking={"type":"adaptive"},
        messages=[{"role":"user","content":build_prompt(jd,competencies,cvs)}]
    ) as s:
        for t in s.text_stream: chunks.append(t)
    ph.empty()
    return "".join(chunks).strip()

# ── PDF export ────────────────────────────────────────────────────────────────

def generate_pdf(results, role_title=""):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=20*mm, rightMargin=20*mm,
                            topMargin=20*mm, bottomMargin=20*mm)
    QB, QO = colors.HexColor("#0075BC"), colors.HexColor("#F7941D")
    G, R, LG = colors.HexColor("#2E7D32"), colors.HexColor("#C62828"), colors.HexColor("#F5F5F5")
    ss = getSampleStyleSheet()
    def sty(name,**kw): return ParagraphStyle(name, **kw)
    T  = sty("T",  fontSize=22, textColor=colors.HexColor("#1A1A2E"), spaceAfter=2*mm, fontName="Helvetica-Bold")
    ST = sty("ST", fontSize=11, textColor=colors.gray, spaceAfter=6*mm, fontName="Helvetica")
    SC = sty("SC", fontSize=13, textColor=QB, spaceBefore=6*mm, spaceAfter=2*mm, fontName="Helvetica-Bold")
    NM = sty("NM", fontSize=12, textColor=colors.HexColor("#1A1A2E"), fontName="Helvetica-Bold", spaceAfter=1*mm)
    BD = sty("BD", fontSize=9,  textColor=colors.HexColor("#333"), fontName="Helvetica", spaceAfter=2*mm, leading=13)
    LB = sty("LB", fontSize=9,  textColor=colors.gray, fontName="Helvetica-Bold", spaceAfter=1*mm)
    FT = sty("FT", fontSize=8,  textColor=colors.gray, fontName="Helvetica-Oblique")

    sl = [r for r in results if r.get("shortlisted")]
    no = [r for r in results if not r.get("shortlisted")]
    date_str = datetime.now().strftime("%d %B %Y")
    story = []
    story.append(Paragraph("CV Screening Report", T))
    role_part = f"Role: {role_title} · " if role_title else ""
    story.append(Paragraph(f"{role_part}Generated: {date_str} · Quest Alliance", ST))
    story.append(HRFlowable(width="100%", thickness=2, color=QB, spaceAfter=6*mm))
    summ = Table([[str(len(results)), str(len(sl)), str(len(no))]],
                 colWidths=[55*mm,55*mm,55*mm])
    summ.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#1A1A2E")),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,-1),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),11),
        ("ALIGN",(0,0),(-1,-1),"CENTER"),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("TOPPADDING",(0,0),(-1,-1),4*mm),("BOTTOMPADDING",(0,0),(-1,-1),4*mm),
        ("BOX",(0,0),(-1,-1),.5,colors.lightgrey),("INNERGRID",(0,0),(-1,-1),.5,colors.lightgrey),
    ]))
    hdr = Table([["Total Screened","Shortlisted","Not Shortlisted"]], colWidths=[55*mm,55*mm,55*mm])
    hdr.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#1A1A2E")),
        ("TEXTCOLOR",(0,0),(-1,-1),colors.white),
        ("FONTNAME",(0,0),(-1,-1),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),10),
        ("ALIGN",(0,0),(-1,-1),"CENTER"),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("TOPPADDING",(0,0),(-1,-1),3*mm),("BOTTOMPADDING",(0,0),(-1,-1),3*mm),
    ]))
    story.extend([hdr, summ, Spacer(1,8*mm)])

    def cblock(r, is_sl):
        sc = r.get("overall",0)
        sl_label = "SHORTLISTED" if is_sl else "NOT SHORTLISTED"
        sl_color = G if is_sl else R
        story.append(Paragraph(f"#{r.get('rank','')} · {r.get('name', r.get('filename',''))}", NM))
        story.append(Paragraph(
            f'<font color="{sl_color.hexval()}">{sl_label}</font>  ·  Score: <b>{sc}/100</b>  ·  '
            f'{r.get("role","")}  ·  {r.get("location","")}', BD))
        story.append(Paragraph(r.get("summary",""), BD))
        s_text = " · ".join(r.get("strengths",[]))
        g_text = " · ".join(r.get("gaps",[]))
        ct = Table([[Paragraph("Strengths", LB), Paragraph("Gaps", LB)],
                    [Paragraph(s_text or "—", BD), Paragraph(g_text or "—", BD)]],
                   colWidths=[82*mm,82*mm])
        ct.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),LG),("VALIGN",(0,0),(-1,-1),"TOP"),
            ("LEFTPADDING",(0,0),(-1,-1),3*mm),("RIGHTPADDING",(0,0),(-1,-1),3*mm),
            ("TOPPADDING",(0,0),(-1,-1),2*mm),("BOTTOMPADDING",(0,0),(-1,-1),2*mm),
            ("BOX",(0,0),(-1,-1),.5,colors.lightgrey),("INNERGRID",(0,0),(-1,-1),.5,colors.lightgrey),
        ]))
        story.extend([ct, Spacer(1,5*mm),
                      HRFlowable(width="100%",thickness=.5,color=colors.lightgrey,spaceAfter=4*mm)])

    if sl:
        story.append(Paragraph(f"✓ Shortlisted ({len(sl)})", SC))
        story.append(HRFlowable(width="100%",thickness=1,color=G,spaceAfter=4*mm))
        for r in sl: cblock(r, True)
    if no:
        story.append(Paragraph(f"✗ Not Shortlisted ({len(no)})", SC))
        story.append(HRFlowable(width="100%",thickness=1,color=R,spaceAfter=4*mm))
        for r in no: cblock(r, False)
    story.append(Spacer(1,4*mm))
    story.append(Paragraph("Generated by Quest CV Screener · Powered by Claude AI · Scores are AI-generated and should be used alongside human review.", FT))
    doc.build(story)
    return buf.getvalue()

# ── Google Drive ──────────────────────────────────────────────────────────────

@st.cache_resource
def get_drive_service():
    try:
        from googleapiclient.discovery import build
        from google.oauth2 import service_account
        raw = st.secrets.get("GOOGLE_SERVICE_ACCOUNT","")
        if not raw: return None
        creds = service_account.Credentials.from_service_account_info(
            json.loads(raw), scopes=["https://www.googleapis.com/auth/drive.readonly"])
        return build("drive","v3",credentials=creds,cache_discovery=False)
    except: return None

def list_drive_files(folder_id):
    svc = get_drive_service()
    if not svc: return []
    return svc.files().list(q=f"'{folder_id}' in parents and trashed=false",
                             fields="files(id,name,mimeType)",pageSize=100).execute().get("files",[])

def download_drive_file(file_id, name, mime):
    from googleapiclient.http import MediaIoBaseDownload
    svc = get_drive_service()
    if mime=="application/vnd.google-apps.document":
        req=svc.files().export_media(fileId=file_id,
            mimeType="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    else:
        req=svc.files().get_media(fileId=file_id)
    buf=io.BytesIO(); dl=MediaIoBaseDownload(buf,req); done=False
    while not done: _,done=dl.next_chunk()
    return buf.getvalue()

def folder_id_from_url(url):
    m=re.search(r"/folders/([a-zA-Z0-9_-]+)",url)
    return m.group(1) if m else url.strip()

# ── Session state defaults ────────────────────────────────────────────────────
for k,v in {"screen":"setup","selected_idx":0,"screening_results":None,
            "cvs":{},"jd":"","competencies":[],"role_title":"",
            "filter":"all","search":"","sort":"match"}.items():
    if k not in st.session_state: st.session_state[k]=v

# ── Header (always visible) ───────────────────────────────────────────────────
screen = st.session_state.screen
total  = len(st.session_state.cvs)
results_data = st.session_state.screening_results or []

shortlisted = [r for r in results_data if r.get("shortlisted")]
possible    = [r for r in results_data if not r.get("shortlisted") and r.get("overall",0)>=65]
weak        = [r for r in results_data if r.get("overall",0)<65 and not r.get("shortlisted")]

header_right = ""
if screen == "results":
    header_right = f'''
      <button onclick="" style="display:flex;align-items:center;gap:7px;font:600 13px 'Work Sans';color:#0075BC;background:#fff;border:1px solid #B8DCF0;border-radius:9px;padding:8px 14px;cursor:pointer">⤓ Export shortlist ({len(shortlisted)})</button>'''

st.markdown(f"""
<div class="qs-header">
  <div style="display:flex;align-items:center;gap:11px">
    <div class="qs-logo">Q<div class="qs-logo-dot"></div></div>
    <div>
      <div style="font-weight:700;font-size:15px;line-height:1.05">CV Screener</div>
      <div style="font:500 10px 'Work Sans';color:#9AA1AE">Quest Alliance · Enabling Self Learning</div>
    </div>
  </div>
  <div style="display:flex;align-items:center;gap:14px">
    {header_right}
    <div style="width:32px;height:32px;border-radius:50%;background:#F7941D;color:#fff;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:12px">TA</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# SETUP SCREEN
# ═══════════════════════════════════════════════════════════════════════════════
if screen == "setup":
    # Step bar
    st.markdown("""
    <div class="qs-step-bar">
      <div class="qs-step-active" style="display:flex;align-items:center;gap:9px">
        <div class="qs-step-num">1</div><span class="qs-step-label">Role</span>
      </div>
      <div class="qs-step-line" style="background:linear-gradient(90deg,#0075BC,#E4E9EF)"></div>
      <div class="qs-step-inactive" style="display:flex;align-items:center;gap:9px">
        <div class="qs-step-num">2</div><span class="qs-step-label">Criteria</span>
      </div>
      <div class="qs-step-line" style="background:#E4E9EF"></div>
      <div class="qs-step-inactive" style="display:flex;align-items:center;gap:9px">
        <div class="qs-step-num">3</div><span class="qs-step-label">Upload CVs</span>
      </div>
      <div class="qs-step-line" style="background:#E4E9EF"></div>
      <div class="qs-step-inactive" style="display:flex;align-items:center;gap:9px">
        <div class="qs-step-num">4</div><span class="qs-step-label">Review</span>
      </div>
    </div>
    """, unsafe_allow_html=True)

    with st.container():
        st.markdown('<div style="max-width:1100px;margin:0 auto;padding:30px 36px">', unsafe_allow_html=True)
        st.markdown("""
        <div style="margin-bottom:20px">
          <div style="font:800 28px 'Work Sans';letter-spacing:-.02em">Let's find your strongest candidates</div>
          <div style="font:500 15px 'Work Sans';color:#5E6675;margin-top:5px">Describe the role and what great looks like. We'll rank every CV against it.</div>
        </div>
        """, unsafe_allow_html=True)

        # Role title
        role_title = st.text_input("Role title", value=st.session_state.role_title,
                                    placeholder="e.g. Placement Officer – Chennai",
                                    label_visibility="collapsed")
        st.session_state.role_title = role_title

        # JD card
        st.markdown('<div class="qs-card">', unsafe_allow_html=True)
        c1, c2 = st.columns([3,1])
        with c1: st.markdown('<div class="qs-label">Job description</div>', unsafe_allow_html=True)
        with c2:
            jd_mode = st.radio("jd_mode", ["Paste text","Upload file"], horizontal=True,
                                label_visibility="collapsed", key="jd_mode_radio")

        jd_input = ""
        if jd_mode == "Paste text":
            jd_input = st.text_area("jd_text", height=120,
                placeholder="Senior Program Associate — Employability. Lead facilitation of youth career-readiness programs…",
                label_visibility="collapsed", value=st.session_state.jd,
                key="jd_textarea")
            st.session_state.jd = jd_input
        else:
            jd_file = st.file_uploader("Upload JD", type=["pdf","docx","doc","txt"],
                                        key="jd_file_upload", label_visibility="collapsed")
            if jd_file:
                try:
                    jd_input = extract_file_text(jd_file)
                    st.session_state.jd = jd_input
                    st.success(f"✅ {jd_file.name}")
                except Exception as e:
                    st.error(str(e))
            else:
                jd_input = st.session_state.jd
        st.markdown('</div>', unsafe_allow_html=True)

        # Competencies card
        st.markdown('<div class="qs-card">', unsafe_allow_html=True)
        cc1, cc2 = st.columns([3,1])
        with cc1: st.markdown('<div class="qs-label">Skill competencies</div>', unsafe_allow_html=True)
        with cc2: st.markdown('<div class="qs-sublabel">One per line · auto-used in scoring</div>', unsafe_allow_html=True)
        comp_text = st.text_area("competencies", height=100,
            placeholder="Youth facilitation\nProgram management\nM&E / data\nStakeholder partnerships\nRegional language",
            label_visibility="collapsed",
            value="\n".join(st.session_state.competencies) if st.session_state.competencies else "",
            key="comp_textarea")
        st.session_state.competencies = [c.strip() for c in comp_text.splitlines() if c.strip()]
        st.markdown('</div>', unsafe_allow_html=True)

        # CV upload card
        st.markdown('<div class="qs-card">', unsafe_allow_html=True)
        vc1, vc2 = st.columns([3,1])
        with vc1:
            st.markdown(f'<div class="qs-label">Candidate CVs <span style="color:#0075BC">· {total} added</span></div>', unsafe_allow_html=True)
        with vc2:
            st.markdown('<div class="qs-sublabel">PDF, DOCX or TXT</div>', unsafe_allow_html=True)

        cv_src = st.radio("cv_source", ["Upload files","Google Drive"], horizontal=True,
                           label_visibility="collapsed", key="cv_source_radio")

        if cv_src == "Upload files":
            uploaded = st.file_uploader("Upload CVs", type=["pdf","docx","doc","txt"],
                                         accept_multiple_files=True, label_visibility="collapsed")
            if uploaded:
                with st.spinner("Reading files…"):
                    for f in uploaded:
                        if f.name not in st.session_state.cvs:
                            try: st.session_state.cvs[f.name] = extract_file_text(f)
                            except: pass
                st.success(f"✅ {len(st.session_state.cvs)} CVs ready")
        else:
            if not get_drive_service():
                st.info("Google Drive not configured. Add GOOGLE_SERVICE_ACCOUNT to Streamlit Secrets.")
            else:
                gd_url = st.text_input("Google Drive folder URL",
                    placeholder="https://drive.google.com/drive/folders/…", key="cv_gd_url")
                if gd_url:
                    fid = folder_id_from_url(gd_url)
                    with st.spinner("Listing…"): files = list_drive_files(fid)
                    if files:
                        sel = st.multiselect("Select CVs", options=files,
                            default=files, format_func=lambda f: f["name"])
                        if st.button("Load from Drive"):
                            prog = st.progress(0)
                            for i,f in enumerate(sel):
                                try:
                                    b = download_drive_file(f["id"],f["name"],f["mimeType"])
                                    st.session_state.cvs[f["name"]] = parse_bytes(b,f["name"])
                                except: pass
                                prog.progress((i+1)/len(sel))
                            prog.empty()
                            st.success(f"✅ {len(st.session_state.cvs)} CVs loaded")

        if st.session_state.cvs:
            names = list(st.session_state.cvs.keys())
            chips_html = "".join(
                f'<span class="qs-file-chip"><span class="qs-file-icon qs-file-type-{"pdf" if n.lower().endswith(".pdf") else "doc"}">'
                f'{"PDF" if n.lower().endswith(".pdf") else "DOC"}</span>'
                f'<span style="font-weight:600;font-size:13px;flex:1;overflow:hidden;white-space:nowrap;text-overflow:ellipsis;max-width:160px">{n}</span>'
                f'<span style="color:#0075BC;font-size:14px">✓</span></span>'
                for n in names[:6]
            )
            extra = f'<span class="qs-file-chip"><span style="font:600 12px \'Work Sans\';color:#5E6675">+{len(names)-6} more</span></span>' if len(names)>6 else ""
            st.markdown(f'<div style="display:flex;flex-wrap:wrap;gap:8px;margin-top:8px">{chips_html}{extra}</div>',
                        unsafe_allow_html=True)

            if st.button("Clear all CVs", type="secondary"):
                st.session_state.cvs = {}
                st.rerun()

        st.markdown('</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    # Footer bar
    can_screen = bool(st.session_state.jd or jd_input) and bool(st.session_state.cvs)
    st.markdown(f"""
    <div style="position:sticky;bottom:0;border-top:1px solid #E4E9EF;background:#fff;
      padding:16px 36px;display:flex;align-items:center;justify-content:space-between;z-index:99">
      <div style="font:500 13px 'Work Sans';color:#5E6675">
        Takes about 30–60 seconds for {len(st.session_state.cvs)} CVs
      </div>
    </div>
    """, unsafe_allow_html=True)

    if st.button(f"Screen {len(st.session_state.cvs)} CVs →", type="primary",
                  disabled=not can_screen, key="screen_btn"):
        jd = st.session_state.jd or jd_input
        comps = "\n".join(st.session_state.competencies) if st.session_state.competencies else ""
        raw = screen_cvs(jd, comps, st.session_state.cvs)
        if raw:
            try:
                st.session_state.screening_results = json.loads(raw)
                st.session_state.screen = "results"
                st.rerun()
            except json.JSONDecodeError:
                st.error("Could not parse response. Raw:")
                st.code(raw)

# ═══════════════════════════════════════════════════════════════════════════════
# RESULTS SCREEN
# ═══════════════════════════════════════════════════════════════════════════════
elif screen == "results":
    results_data = st.session_state.screening_results or []
    strong_list   = [r for r in results_data if r.get("overall",0) >= 85]
    possible_list = [r for r in results_data if 65 <= r.get("overall",0) < 85]
    weak_list     = [r for r in results_data if r.get("overall",0) < 65]

    # Sub-header
    role_label = st.session_state.role_title or "Screening Results"
    st.markdown(f"""
    <div style="padding:24px 36px 16px;background:#fff;border-bottom:1px solid #E4E9EF">
      <div style="display:flex;align-items:flex-end;justify-content:space-between;max-width:1200px;margin:0 auto">
        <div>
          <div style="font:500 13px 'Work Sans';color:#5E6675;display:flex;align-items:center;gap:6px;cursor:pointer">
            ← {role_label}
          </div>
          <div style="font:800 24px 'Work Sans';letter-spacing:-.02em;margin-top:5px">
            {len(results_data)} candidates screened &amp; ranked
          </div>
        </div>
        <div style="display:flex;gap:10px">
          <div style="text-align:center;background:#E3F1FA;border-radius:11px;padding:9px 16px;cursor:pointer">
            <div style="font:800 20px 'Work Sans';color:#005A91">{len(strong_list)}</div>
            <div style="font:600 11px 'Work Sans';color:#005A91">Strong</div>
          </div>
          <div style="text-align:center;background:#FEF0DC;border-radius:11px;padding:9px 16px;cursor:pointer">
            <div style="font:800 20px 'Work Sans';color:#C76A0A">{len(possible_list)}</div>
            <div style="font:600 11px 'Work Sans';color:#C76A0A">Possible</div>
          </div>
          <div style="text-align:center;background:#FDE7DE;border-radius:11px;padding:9px 16px;cursor:pointer">
            <div style="font:800 20px 'Work Sans';color:#C23A18">{len(weak_list)}</div>
            <div style="font:600 11px 'Work Sans';color:#C23A18">Weak</div>
          </div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # Filter / search bar
    fc1, fc2, fc3, fc4, fc5, fc6 = st.columns([4,1,1,1,1,2])
    with fc1: search = st.text_input("search", placeholder="🔍  Search candidates…",
                                      label_visibility="collapsed", key="search_input")
    with fc2:
        if st.button("All",     key="f_all",      type="secondary"): st.session_state.filter="all"
    with fc3:
        if st.button("Strong",  key="f_strong",   type="secondary"): st.session_state.filter="strong"
    with fc4:
        if st.button("Possible",key="f_possible", type="secondary"): st.session_state.filter="possible"
    with fc5:
        if st.button("Weak",    key="f_weak",     type="secondary"): st.session_state.filter="weak"
    with fc6:
        sort = st.selectbox("sort", ["Best match","Name A–Z","Experience"],
                             label_visibility="collapsed", key="sort_sel")

    # Filter & sort
    filt = st.session_state.filter
    filtered = results_data
    if filt == "strong":   filtered = strong_list
    elif filt == "possible": filtered = possible_list
    elif filt == "weak":   filtered = weak_list
    if search:
        filtered = [r for r in filtered if search.lower() in r.get("name","").lower()
                    or search.lower() in r.get("filename","").lower()]
    if "Name" in sort:  filtered = sorted(filtered, key=lambda r: r.get("name",""))
    elif "Exp" in sort: filtered = sorted(filtered, key=lambda r: -r.get("years",0))

    # Candidate rows
    st.markdown('<div style="padding:14px 30px;max-width:1200px;margin:0 auto">', unsafe_allow_html=True)
    if not filtered:
        st.markdown('<div style="text-align:center;font:600 14px \'Work Sans\';color:#9AA1AE;padding:40px">No candidates match these filters.</div>', unsafe_allow_html=True)

    for i, r in enumerate(filtered):
        sc = r.get("overall", 0)
        bname, bbg, bcolor, _ = band(sc)
        ring_bg, ring_color = score_ring_colors(sc)
        av_color = avatar_color(r.get("name", r.get("filename","")))
        name = r.get("name", r.get("filename",""))
        inits = initials(name)
        tags = r.get("strengths", [])[:2]
        tags_html = "".join(f'<span style="background:#E3F1FA;color:#005A91;border-radius:6px;padding:4px 9px;font-weight:600;font-size:11px;white-space:nowrap">{t}</span>' for t in tags)
        left_border = ring_color
        border_style = f"border-left:4px solid {left_border}"

        col_main, col_btn = st.columns([10,1])
        with col_main:
            st.markdown(f"""
            <div class="qs-cand-row" style="{border_style}">
              <div style="font:800 15px 'Work Sans';color:#C2C8D2;width:20px;flex:none">{r.get('rank',i+1)}</div>
              <div class="qs-avatar" style="background:{av_color}">{inits}</div>
              <div style="flex:1;min-width:0">
                <div style="font-weight:700;font-size:15px">{name}</div>
                <div style="font:500 13px 'Work Sans';color:#5E6675">{r.get('role','')} · {r.get('years','')} yrs · {r.get('location','')}</div>
              </div>
              <div style="display:flex;gap:6px;flex-wrap:wrap">{tags_html}</div>
              <div class="qs-score-ring" style="background:{ring_bg}">
                <div class="qs-score-inner" style="color:{ring_color}">{sc}</div>
              </div>
            </div>
            """, unsafe_allow_html=True)

        with col_btn:
            if st.button("View →", key=f"view_{i}", type="secondary"):
                st.session_state.selected_idx = next(
                    (j for j,x in enumerate(results_data) if x.get("rank")==r.get("rank")), i)
                st.session_state.screen = "detail"
                st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)

    # PDF download
    st.divider()
    c1, c2 = st.columns([1,4])
    with c1:
        if st.button("← New screening", key="back_to_setup"):
            st.session_state.screen = "setup"
            st.session_state.screening_results = None
            st.rerun()
    with c2:
        with st.spinner("Preparing PDF…"):
            pdf = generate_pdf(results_data, st.session_state.role_title)
        st.download_button("📄 Download PDF Report", data=pdf,
            file_name=f"CV_Screening_{datetime.now().strftime('%Y-%m-%d')}.pdf",
            mime="application/pdf", type="primary")

# ═══════════════════════════════════════════════════════════════════════════════
# DETAIL SCREEN
# ═══════════════════════════════════════════════════════════════════════════════
elif screen == "detail":
    results_data = st.session_state.screening_results or []
    idx = st.session_state.selected_idx
    if not results_data or idx >= len(results_data):
        st.session_state.screen = "results"; st.rerun()

    r = results_data[idx]
    sc = r.get("overall", 0)
    name = r.get("name", r.get("filename",""))
    inits = initials(name)
    av_color = avatar_color(name)
    ring_bg, ring_color = score_ring_colors(sc)
    bname, bbg, bcolor, _ = band(sc)
    band_label = bname.upper()
    scores = r.get("scores", [])
    labels = r.get("competency_labels", [f"Competency {i+1}" for i in range(len(scores))])

    # Nav bar
    nav1, nav2, nav3 = st.columns([2,6,2])
    with nav1:
        if st.button("← Back to candidates", key="back_results"):
            st.session_state.screen = "results"; st.rerun()
    with nav3:
        pc1, pc2 = st.columns(2)
        with pc1:
            if idx > 0 and st.button("↑ Prev"):
                st.session_state.selected_idx -= 1; st.rerun()
        with pc2:
            if idx < len(results_data)-1 and st.button("Next ↓"):
                st.session_state.selected_idx += 1; st.rerun()

    # Candidate header
    st.markdown(f"""
    <div style="padding:26px 36px;background:#fff;border-bottom:1px solid #E4E9EF;display:flex;align-items:center;gap:22px">
      <div class="qs-avatar" style="width:74px;height:74px;font-size:27px;background:{av_color}">{inits}</div>
      <div style="flex:1">
        <div style="display:flex;align-items:center;gap:10px">
          <div style="font:800 24px 'Work Sans';letter-spacing:-.02em">{name}</div>
          <span style="background:{bbg};color:{bcolor};border-radius:999px;padding:4px 11px;font-weight:700;font-size:12px">{band_label}</span>
        </div>
        <div style="font:500 14px 'Work Sans';color:#5E6675;margin-top:4px">{r.get('role','')} · {r.get('years','')} yrs · {r.get('location','')}</div>
        <div style="display:flex;gap:16px;margin-top:9px;font:500 13px 'Work Sans';color:#3A4150">
          {'<span>✉ '+r.get('email','')+'</span>' if r.get('email') else ''}
          {'<span>📞 '+r.get('phone','')+'</span>' if r.get('phone') else ''}
        </div>
      </div>
      <div style="text-align:center;flex:none">
        <div style="width:78px;height:78px;border-radius:50%;background:{ring_bg};display:flex;align-items:center;justify-content:center">
          <div style="width:60px;height:60px;border-radius:50%;background:#fff;display:flex;flex-direction:column;align-items:center;justify-content:center">
            <div style="font-weight:800;font-size:22px;color:{ring_color};line-height:1">{sc}</div>
            <div style="font:600 9px 'Work Sans';color:#9AA1AE">/ 100</div>
          </div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # Action buttons
    ac1, ac2, ac3, _, ac4 = st.columns([2,2,2,3,2])
    with ac1:
        shortlist_label = "★ Shortlisted" if r.get("shortlisted") else "★ Shortlist"
        st.button(shortlist_label, key="sl_btn", type="primary")
    with ac2:
        st.button("✕ Reject", key="rj_btn", type="secondary")
    with ac3:
        pdf_single = generate_pdf([r], st.session_state.role_title)
        st.download_button("⤓ Export PDF", data=pdf_single,
            file_name=f"{name.replace(' ','_')}_Report.pdf",
            mime="application/pdf", type="secondary")

    st.divider()

    # Main detail body
    left_col, right_col = st.columns([1.2, 1])

    with left_col:
        st.markdown('<div style="font:700 16px \'Work Sans\';margin-bottom:4px">Match breakdown</div>', unsafe_allow_html=True)
        st.markdown('<div style="font:500 13px \'Work Sans\';color:#5E6675;margin-bottom:18px">Scored against your competencies</div>', unsafe_allow_html=True)

        for label, val in zip(labels, scores):
            color = "#0075BC" if val>=85 else "#F7941D" if val>=65 else "#E85020"
            st.markdown(f"""
            <div style="margin-bottom:16px">
              <div style="display:flex;justify-content:space-between;font-weight:600;font-size:14px;margin-bottom:7px">
                <span>{label}</span><span style="color:{color}">{val}</span>
              </div>
              <div class="qs-bar-track">
                <div class="qs-bar-fill" style="width:{val}%;background:{color}"></div>
              </div>
            </div>
            """, unsafe_allow_html=True)

        if r.get("flag"):
            st.markdown(f"""
            <div style="background:#FEF0DC;border-radius:11px;padding:14px 16px;display:flex;gap:11px;margin-top:6px">
              <div style="font-size:16px">⚠</div>
              <div>
                <div style="font-weight:700;font-size:13px;color:#B0640C">One thing to verify</div>
                <div style="font:500 13px/1.5 'Work Sans';color:#9A6410;margin-top:2px">{r['flag']}</div>
              </div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown('<div style="margin-top:18px"><b>Gaps / Concerns</b></div>', unsafe_allow_html=True)
        for g in r.get("gaps", []):
            st.markdown(f'<div style="font:500 13px/1.5 \'Work Sans\';color:#5E6675;margin-top:6px">• {g}</div>', unsafe_allow_html=True)

    with right_col:
        st.markdown(f"""
        <div class="qs-ai-box">
          <div style="display:flex;align-items:center;gap:8px;font-weight:700;font-size:13px;margin-bottom:8px">
            <span style="width:18px;height:18px;border-radius:5px;background:rgba(255,255,255,.22);
              display:flex;align-items:center;justify-content:center;font-size:11px">✦</span> AI summary
          </div>
          <div style="font:500 13.5px/1.6 'Work Sans';color:rgba(255,255,255,.92)">{r.get('summary','')}</div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown('<div style="font:700 15px \'Work Sans\';margin-bottom:12px">Evidence from CV</div>', unsafe_allow_html=True)
        for ev in r.get("evidence", []):
            st.markdown(f"""
            <div class="qs-evidence">
              <div class="qs-evidence-label">{ev.get('label','')}</div>
              <div style="font:500 13px/1.5 'Work Sans';color:#3A4150">{ev.get('text','')}</div>
            </div>
            """, unsafe_allow_html=True)
