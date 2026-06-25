import os, io, re, json
from datetime import datetime
import anthropic
import streamlit as st
import streamlit.components.v1 as components
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

#MainMenu, footer, header { visibility: hidden; }
[data-testid="stSidebar"] { display: none; }
[data-testid="stAppViewContainer"] { background: #F4F7FA; }
* { font-family: 'Work Sans', -apple-system, sans-serif; box-sizing: border-box; text-align: left; }

/* Centered content wrapper — 60% on desktop, 90% on mobile */
.block-container {
  max-width: 900px !important;
  margin: 0 auto !important;
  padding: 0 0 40px 0 !important;
  width: 60% !important;
}
@media (max-width: 900px) {
  .block-container { width: 90% !important; }
}

/* Native widget overrides */
[data-testid="stTextInput"] input,
[data-testid="stTextArea"] textarea {
  border-radius: 10px !important;
  border: 1.5px solid #E4E9EF !important;
  font-family: 'Work Sans', sans-serif !important;
  font-size: 14px !important;
  background: #fff !important;
  text-align: left !important;
}
[data-testid="stTextInput"] input:focus,
[data-testid="stTextArea"] textarea:focus {
  border-color: #0075BC !important;
  box-shadow: 0 0 0 3px rgba(0,117,188,.12) !important;
}
[data-testid="stFileUploader"] {
  border: 1.5px dashed #B8DCF0 !important;
  border-radius: 11px !important;
  background: repeating-linear-gradient(135deg,#EAF4FB 0 10px,#fff 10px 20px) !important;
  padding: 8px !important;
}
[data-testid="stButton"] button[kind="primary"] {
  background: #0075BC !important;
  color: #fff !important;
  border-radius: 11px !important;
  font-weight: 700 !important;
  font-size: 14px !important;
  padding: 10px 20px !important;
  border: none !important;
  box-shadow: 0 4px 14px rgba(0,117,188,.25) !important;
}
[data-testid="stButton"] button[kind="secondary"] {
  border-radius: 9px !important;
  font-weight: 600 !important;
  font-size: 13px !important;
  border: 1px solid #E4E9EF !important;
  background: #fff !important;
}
[data-testid="stRadio"] label { font-size: 13px !important; font-weight: 600 !important; }
[data-testid="stSelectbox"] { border-radius: 9px !important; }
[data-testid="stVerticalBlock"] > [data-testid="stVerticalBlockBorderWrapper"] {
  border-radius: 14px !important;
  border: 1px solid #E4E9EF !important;
  background: #fff !important;
}
[data-testid="stVerticalBlockBorderWrapper"] > div > div {
  padding: 16px 20px !important;
}
/* File chip */
.qs-file-chip {
  display: flex; align-items: center; gap: 9px;
  background: #F4F7FA; border: 1px solid #E4E9EF;
  border-radius: 9px; padding: 9px 11px; margin-bottom: 6px;
}
.qs-file-icon {
  width: 28px; height: 28px; border-radius: 6px;
  font-weight: 700; font-size: 9px;
  display: flex; align-items: center; justify-content: center; flex: none;
}
/* Markdown left-align */
p, h1, h2, h3, h4, li, label { text-align: left !important; }

/* Shortlist / Reject toggle button states — keyed by title attribute */
button[title="undo-shortlist"] {
  background: #1B6E2E !important;
  color: #fff !important;
  border-color: #1B6E2E !important;
  box-shadow: 0 2px 8px rgba(27,110,46,.35) !important;
}
button[title="undo-shortlist"]:hover {
  background: #155724 !important;
  border-color: #155724 !important;
}
button[title="undo-reject"] {
  background: #C62828 !important;
  color: #fff !important;
  border-color: #C62828 !important;
  box-shadow: 0 2px 8px rgba(198,40,40,.35) !important;
}
button[title="undo-reject"]:hover {
  background: #a31e1e !important;
  border-color: #a31e1e !important;
}
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
    if n.endswith(".pdf"):            return extract_text_from_pdf(b)
    if n.endswith((".docx", ".doc")): return extract_text_from_docx(b)
    return b.decode("utf-8", errors="replace")

def extract_file_text(f):
    try: f.seek(0)
    except: pass
    return parse_bytes(f.read(), f.name)

def initials(name):
    parts = name.split()
    return (parts[0][0] + parts[-1][0]).upper() if len(parts) > 1 else parts[0][:2].upper()

AVATAR_COLORS = ["#0075BC","#3251A3","#27AAE1","#444A67","#6B5BA8","#1A7A5E","#C76A0A","#C23A18"]

def avatar_color(name):
    return AVATAR_COLORS[sum(ord(c) for c in name) % len(AVATAR_COLORS)]

def band(score):
    if score >= 85: return "strong",   "#E3F1FA", "#005A91", "#B8DCF0"
    if score >= 65: return "possible", "#FEF0DC", "#C76A0A", "#F7D49A"
    return "weak", "#FDE7DE", "#C23A18", "#F5B8A8"

def score_ring_colors(score):
    if score >= 85: return "#0075BC20", "#0075BC"
    if score >= 65: return "#F7941D20", "#F7941D"
    return "#E8502020", "#E85020"

def candidate_key(r):
    return r.get("filename") or r.get("name") or str(r.get("rank", ""))

# Auto-resizing iframe helper — measures actual rendered height and posts it back to Streamlit
_RESIZE_JS = """
<script>
(function(){
  function resize(){
    var h = Math.max(
      document.body ? document.body.scrollHeight : 0,
      document.body ? document.body.offsetHeight : 0,
      document.documentElement ? document.documentElement.scrollHeight : 0
    );
    window.parent.postMessage(
      {isStreamlitMessage:true, type:'streamlit:setFrameHeight', height:h}, '*'
    );
  }
  resize();
  setTimeout(resize, 80);
  setTimeout(resize, 300);
  if (document.readyState !== 'complete') window.addEventListener('load', resize);
})();
</script>"""

def auto_html(content):
    """Render HTML in an iframe that auto-sizes to its actual content height."""
    components.html(content + _RESIZE_JS, height=10, scrolling=False)

def record_history(r, action):
    """Append a history entry for a candidate action."""
    key = candidate_key(r)
    if "candidate_history" not in st.session_state:
        st.session_state.candidate_history = {}
    history = st.session_state.candidate_history.setdefault(key, [])
    history.append({
        "action": action,
        "name": r.get("name", key),
        "ts": datetime.now().strftime("%d %b %Y, %H:%M"),
    })

def record_feedback(r, ai_overall, ai_scores, ai_labels,
                    human_overall, human_scores, reason, approved):
    """Store a score calibration and add it to the feedback examples pool."""
    key = candidate_key(r)
    fb = {
        "name": r.get("name", key),
        "role": r.get("role", ""),
        "years": r.get("years", ""),
        "ai_overall": ai_overall,
        "ai_scores": ai_scores,
        "human_overall": human_overall,
        "human_scores": human_scores,
        "competency_labels": ai_labels,
        "reason": reason,
        "approved": approved,
        "ts": datetime.now().strftime("%d %b %Y, %H:%M"),
    }
    st.session_state.score_feedback[key] = fb
    # Add as a learning example only when the reviewer made a meaningful change
    if not approved or abs(human_overall - ai_overall) >= 5:
        example = (
            f"Candidate: {r.get('name','')} | Role: {r.get('role','')} | {r.get('years','')} yrs exp\n"
            f"AI scored overall {ai_overall}/100. "
        )
        if approved:
            example += f"Reviewer APPROVED this score."
        else:
            example += (
                f"Reviewer ADJUSTED overall to {human_overall}/100. "
                f"Reason: {reason}. "
            )
            deltas = []
            for label, a_sc, h_sc in zip(ai_labels, ai_scores, human_scores):
                if abs(h_sc - a_sc) >= 5:
                    deltas.append(f"{label}: AI={a_sc} → Reviewer={h_sc}")
            if deltas:
                example += "Competency adjustments: " + "; ".join(deltas) + "."
        if "feedback_examples" not in st.session_state:
            st.session_state.feedback_examples = []
        st.session_state.feedback_examples.append(example)

# ── Claude API ────────────────────────────────────────────────────────────────

def build_prompt(jd, competencies, cvs):
    cv_block = "".join(f"\n---\nCV #{i} – {name}\n{text}\n"
                       for i, (name, text) in enumerate(cvs.items(), 1))

    # Inject reviewer calibration examples so AI learns from past feedback
    examples = getattr(st.session_state, "feedback_examples", [])
    calibration_block = ""
    if examples:
        calibration_block = (
            "\n## Reviewer Calibration (learn from these past adjustments)\n"
            "The following examples show how this organisation's reviewers have "
            "adjusted or approved AI scores in previous screenings. Use these to "
            "calibrate your scoring for this batch.\n"
            + "\n".join(f"- {ex}" for ex in examples[-10:])  # keep last 10
            + "\n"
        )

    return f"""You are an expert HR screener for Quest Alliance, an NGO focused on youth skilling in India.

## Job Description
{jd}

## Required Skill Competencies
{competencies or "(derive from JD)"}
{calibration_block}
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

def extract_competencies(jd_text):
    api_key = os.environ.get("ANTHROPIC_API_KEY") or st.secrets.get("ANTHROPIC_API_KEY", "")
    if not api_key: return []
    client = anthropic.Anthropic(api_key=api_key)
    msg = client.messages.create(
        model="claude-opus-4-8", max_tokens=300,
        messages=[{"role": "user", "content":
            f"Extract exactly 5 key skill competencies from this job description. "
            f"Return ONLY a JSON array of 5 short strings (3-5 words each), nothing else.\n\n{jd_text}"}]
    )
    raw = msg.content[0].text.strip()
    raw = re.sub(r"^```[a-z]*\n?", "", raw).rstrip("` \n")
    return json.loads(raw)

def screen_cvs(jd, competencies, cvs, status_placeholder=None):
    api_key = os.environ.get("ANTHROPIC_API_KEY") or st.secrets.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        st.error("⚠️ ANTHROPIC_API_KEY is not set in Streamlit Secrets."); return None
    client = anthropic.Anthropic(api_key=api_key)
    if status_placeholder:
        status_placeholder.info(f"🔍 Analysing {len(cvs)} CVs with Claude… this takes about 30–60 seconds.")
    chunks = []
    try:
        with client.messages.stream(
            model="claude-opus-4-8", max_tokens=8000, thinking={"type": "adaptive"},
            messages=[{"role": "user", "content": build_prompt(jd, competencies, cvs)}]
        ) as s:
            for t in s.text_stream: chunks.append(t)
    except Exception as e:
        if status_placeholder: status_placeholder.empty()
        st.error(f"Claude API error: {e}"); return None
    if status_placeholder: status_placeholder.empty()
    return "".join(chunks).strip()

# ── PDF export ────────────────────────────────────────────────────────────────

def generate_pdf(results, role_title=""):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=20*mm, rightMargin=20*mm,
                            topMargin=20*mm, bottomMargin=20*mm)
    QB = colors.HexColor("#0075BC")
    G, R = colors.HexColor("#2E7D32"), colors.HexColor("#C62828")
    LG   = colors.HexColor("#F5F5F5")
    def sty(name, **kw): return ParagraphStyle(name, **kw)
    T  = sty("T",  fontSize=22, textColor=colors.HexColor("#1A1A2E"), spaceAfter=2*mm, fontName="Helvetica-Bold")
    ST = sty("ST", fontSize=11, textColor=colors.gray, spaceAfter=6*mm, fontName="Helvetica")
    SC = sty("SC", fontSize=13, textColor=QB, spaceBefore=6*mm, spaceAfter=2*mm, fontName="Helvetica-Bold")
    NM = sty("NM", fontSize=12, textColor=colors.HexColor("#1A1A2E"), fontName="Helvetica-Bold", spaceAfter=1*mm)
    BD = sty("BD", fontSize=9,  textColor=colors.HexColor("#333"), fontName="Helvetica", spaceAfter=2*mm, leading=13)
    LB = sty("LB", fontSize=9,  textColor=colors.gray, fontName="Helvetica-Bold", spaceAfter=1*mm)
    FT = sty("FT", fontSize=8,  textColor=colors.gray, fontName="Helvetica-Oblique")

    sl  = [r for r in results if r.get("shortlisted")]
    no  = [r for r in results if not r.get("shortlisted")]
    date_str = datetime.now().strftime("%d %B %Y")
    story = []
    story.append(Paragraph("CV Screening Report", T))
    role_part = f"Role: {role_title} · " if role_title else ""
    story.append(Paragraph(f"{role_part}Generated: {date_str} · Quest Alliance", ST))
    story.append(HRFlowable(width="100%", thickness=2, color=QB, spaceAfter=6*mm))
    hdr = Table([["Total Screened", "Shortlisted", "Not Shortlisted"]], colWidths=[55*mm]*3)
    hdr.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#1A1A2E")),
        ("TEXTCOLOR",(0,0),(-1,-1),colors.white),
        ("FONTNAME",(0,0),(-1,-1),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),10),
        ("ALIGN",(0,0),(-1,-1),"CENTER"),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("TOPPADDING",(0,0),(-1,-1),3*mm),("BOTTOMPADDING",(0,0),(-1,-1),3*mm),
    ]))
    summ = Table([[str(len(results)), str(len(sl)), str(len(no))]], colWidths=[55*mm]*3)
    summ.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#1A1A2E")),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,-1),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),11),
        ("ALIGN",(0,0),(-1,-1),"CENTER"),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("TOPPADDING",(0,0),(-1,-1),4*mm),("BOTTOMPADDING",(0,0),(-1,-1),4*mm),
        ("BOX",(0,0),(-1,-1),.5,colors.lightgrey),("INNERGRID",(0,0),(-1,-1),.5,colors.lightgrey),
    ]))
    story.extend([hdr, summ, Spacer(1,8*mm)])

    def cblock(r, is_sl):
        sc = r.get("overall", 0)
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
                   colWidths=[82*mm, 82*mm])
        ct.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),LG),("VALIGN",(0,0),(-1,-1),"TOP"),
            ("LEFTPADDING",(0,0),(-1,-1),3*mm),("RIGHTPADDING",(0,0),(-1,-1),3*mm),
            ("TOPPADDING",(0,0),(-1,-1),2*mm),("BOTTOMPADDING",(0,0),(-1,-1),2*mm),
            ("BOX",(0,0),(-1,-1),.5,colors.lightgrey),("INNERGRID",(0,0),(-1,-1),.5,colors.lightgrey),
        ]))
        story.extend([ct, Spacer(1,5*mm),
                      HRFlowable(width="100%", thickness=.5, color=colors.lightgrey, spaceAfter=4*mm)])

    if sl:
        story.append(Paragraph(f"✓ Shortlisted ({len(sl)})", SC))
        story.append(HRFlowable(width="100%", thickness=1, color=G, spaceAfter=4*mm))
        for r in sl: cblock(r, True)
    if no:
        story.append(Paragraph(f"✗ Not Shortlisted ({len(no)})", SC))
        story.append(HRFlowable(width="100%", thickness=1, color=R, spaceAfter=4*mm))
        for r in no: cblock(r, False)
    story.append(Spacer(1,4*mm))
    story.append(Paragraph(
        "Generated by Quest CV Screener · Powered by Claude AI · "
        "Scores are AI-generated and should be used alongside human review.", FT))
    doc.build(story)
    return buf.getvalue()

# ── Excel export ─────────────────────────────────────────────────────────────

def generate_excel(results, role_title="", history=None):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()

    # ── Sheet 1: Results ───────────────────────────────────────────────────────
    ws = wb.active
    ws.title = "Screening Results"

    BLUE  = "FF0075BC"
    GREEN = "FF1B6E2E"
    RED   = "FFC62828"
    ORG   = "FFF7941D"
    LGREY = "FFF4F7FA"
    WHITE = "FFFFFFFF"
    DARK  = "FF1A1A2E"

    hdr_font  = Font(name="Calibri", bold=True, color=WHITE, size=11)
    hdr_fill  = PatternFill("solid", fgColor=DARK)
    hdr_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

    thin = Side(style="thin", color="FFE4E9EF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    title_row = [
        f"CV Screening Report — {role_title}" if role_title else "CV Screening Report",
        "", "", "", "", f"Generated: {datetime.now().strftime('%d %B %Y')}"
    ]
    ws.append(title_row)
    ws["A1"].font = Font(name="Calibri", bold=True, size=14, color=BLUE)
    ws.merge_cells("A1:E1")
    ws["F1"].font = Font(name="Calibri", size=10, color="FF5E6675")
    ws["F1"].alignment = Alignment(horizontal="right")
    ws.append([])

    columns = ["Rank", "Name", "Role", "Years Exp", "Location",
               "Score (Reviewed)", "Band", "Status", "Review Status",
               "Strengths", "Gaps", "Summary", "Flag",
               "Email", "Phone", "Filename"]
    ws.append(columns)
    for col_idx, _ in enumerate(columns, 1):
        cell = ws.cell(row=3, column=col_idx)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = hdr_align
        cell.border = border

    band_colors = {"strong": GREEN, "possible": ORG, "weak": RED}

    for r in results:
        sc = r.get("overall", 0)
        bname, _, _, _ = band(sc)
        sl = r.get("shortlisted")
        status = "Shortlisted" if sl is True else "Rejected" if sl is False else "Pending"
        ck = r.get("filename") or r.get("name") or ""
        fb_entry = (history or {}).get(ck) if history else None
        # Use calibrated score if available (history arg reused for score_feedback here via caller)
        score_fb = {}
        if hasattr(st, 'session_state'):
            score_fb = st.session_state.score_feedback.get(ck, {})
        reviewed_score = score_fb.get("human_overall", sc) if score_fb else sc
        review_status  = ("Approved" if score_fb.get("approved") else
                          f"Calibrated ({sc}→{reviewed_score})") if score_fb else "Not reviewed"
        row = [
            r.get("rank", ""),
            r.get("name", r.get("filename", "")),
            r.get("role", ""),
            r.get("years", ""),
            r.get("location", ""),
            reviewed_score,
            bname.capitalize(),
            status,
            review_status,
            "; ".join(r.get("strengths", [])),
            "; ".join(r.get("gaps", [])),
            r.get("summary", ""),
            r.get("flag", "") or "",
            r.get("email", ""),
            r.get("phone", ""),
            r.get("filename", ""),
        ]
        ws.append(row)
        data_row = ws.max_row
        for col_idx in range(1, len(columns)+1):
            cell = ws.cell(row=data_row, column=col_idx)
            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            cell.border = border
            if data_row % 2 == 0:
                cell.fill = PatternFill("solid", fgColor=LGREY)
        # Color band cell (col 7)
        band_cell = ws.cell(row=data_row, column=7)
        fc = band_colors.get(bname, DARK)
        band_cell.font = Font(name="Calibri", bold=True, color=WHITE, size=10)
        band_cell.fill = PatternFill("solid", fgColor=fc)
        # Color status cell (col 8)
        st_cell = ws.cell(row=data_row, column=8)
        if sl is True:
            st_cell.font = Font(name="Calibri", bold=True, color=WHITE)
            st_cell.fill = PatternFill("solid", fgColor=GREEN)
        elif sl is False:
            st_cell.font = Font(name="Calibri", bold=True, color=WHITE)
            st_cell.fill = PatternFill("solid", fgColor=RED)
        # Color review status cell (col 9)
        rv_cell = ws.cell(row=data_row, column=9)
        if score_fb:
            rv_cell.font = Font(name="Calibri", bold=True,
                                color=GREEN if score_fb.get("approved") else BLUE, size=10)

    # Column widths
    col_widths = [6, 22, 22, 10, 16, 14, 10, 12, 18, 40, 30, 50, 35, 24, 16, 24]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[3].height = 22
    ws.freeze_panes = "A4"

    # ── Sheet 2: History ───────────────────────────────────────────────────────
    if history:
        wh = wb.create_sheet("Status History")
        wh.append(["Candidate", "Action", "Timestamp"])
        for col_idx in range(1, 4):
            cell = wh.cell(row=1, column=col_idx)
            cell.font = hdr_font
            cell.fill = hdr_fill
            cell.alignment = hdr_align
        all_entries = []
        for entries in history.values():
            all_entries.extend(entries)
        all_entries.sort(key=lambda e: e["ts"], reverse=True)
        for e in all_entries:
            wh.append([e.get("name",""), e.get("action","").capitalize(), e.get("ts","")])
            row_idx = wh.max_row
            for col_idx in range(1, 4):
                wh.cell(row=row_idx, column=col_idx).alignment = Alignment(horizontal="left", vertical="center")
                wh.cell(row=row_idx, column=col_idx).border = border
        wh.column_dimensions["A"].width = 24
        wh.column_dimensions["B"].width = 16
        wh.column_dimensions["C"].width = 22
        wh.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── Session state ─────────────────────────────────────────────────────────────
for k, v in {"screen": "setup", "selected_idx": 0, "screening_results": None,
             "cvs": {}, "jd": "", "competencies": [], "role_title": "",
             "filter": "all", "search": "", "sort": "match",
             "jd_last_detected": "", "candidate_history": {},
             "score_feedback": {}, "feedback_examples": []}.items():
    if k not in st.session_state: st.session_state[k] = v

# ── Shared header ─────────────────────────────────────────────────────────────
screen = st.session_state.screen

components.html(f"""
<link href="https://fonts.googleapis.com/css2?family=Work+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<div style="height:60px;background:#fff;border-bottom:1px solid #E4E9EF;display:flex;align-items:center;justify-content:space-between;padding:0 24px;font-family:'Work Sans',sans-serif">
  <div style="display:flex;align-items:center;gap:11px">
    <div style="width:32px;height:32px;border-radius:9px;background:#0075BC;color:#fff;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:16px;position:relative;flex:none">
      Q<div style="width:10px;height:10px;border-radius:50%;background:#F7941D;position:absolute;bottom:-2px;right:-2px;border:2px solid #fff"></div>
    </div>
    <div>
      <div style="font-weight:700;font-size:15px;line-height:1.05;color:#1A1A2E;text-align:left">CV Screener</div>
      <div style="font:500 10px 'Work Sans';color:#9AA1AE;text-align:left">Quest Alliance · Enabling Self Learning</div>
    </div>
  </div>
  <div style="display:flex;align-items:center;gap:14px">
    <div style="width:32px;height:32px;border-radius:50%;background:#F7941D;color:#fff;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:12px">TA</div>
  </div>
</div>
""", height=62, scrolling=False)

# ═══════════════════════════════════════════════════════════════════════════════
# SETUP SCREEN
# ═══════════════════════════════════════════════════════════════════════════════
if screen == "setup":

    st.markdown("""
<div style="padding:28px 0 12px">
  <div style="font:800 26px 'Work Sans',sans-serif;letter-spacing:-.02em;color:#1A1A2E;text-align:left">Let's find your strongest candidates</div>
  <div style="font:500 14px 'Work Sans',sans-serif;color:#5E6675;margin-top:6px;text-align:left">Describe the role and what great looks like. We'll rank every CV against it.</div>
</div>""", unsafe_allow_html=True)

    role_title = st.text_input("Role title", value=st.session_state.role_title,
                                placeholder="e.g. Placement Officer – Chennai",
                                label_visibility="visible")
    st.session_state.role_title = role_title

    # ── JD card ───────────────────────────────────────────────────────────────
    with st.container(border=True):
        hc1, hc2 = st.columns([3, 1])
        with hc1:
            st.markdown("**Job description**")
        with hc2:
            jd_mode = st.radio("jd_mode", ["Paste text", "Upload file"], horizontal=True,
                                label_visibility="collapsed", key="jd_mode_radio")
        jd_input = ""
        if jd_mode == "Paste text":
            jd_input = st.text_area("jd_text", height=130,
                placeholder="Senior Program Associate — Employability. Lead facilitation of youth career-readiness programs…",
                label_visibility="collapsed", value=st.session_state.jd, key="jd_textarea")
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

    # ── Auto-detect competencies ───────────────────────────────────────────────
    current_jd = st.session_state.jd or jd_input
    if (current_jd and current_jd != st.session_state.jd_last_detected and len(current_jd) > 80):
        st.session_state.jd_last_detected = current_jd
        with st.spinner("Auto-detecting skill competencies from JD…"):
            try:
                detected = extract_competencies(current_jd)
                if detected:
                    st.session_state.competencies = detected
                    st.rerun()
            except Exception:
                pass

    # ── Competencies card ─────────────────────────────────────────────────────
    with st.container(border=True):
        cc1, cc2 = st.columns([3, 1])
        with cc1:
            st.markdown("**Skill competencies**")
        with cc2:
            st.caption("Auto-detected from JD · tap × to remove")
        comps = st.session_state.competencies
        if comps:
            pill_cols = st.columns(min(len(comps), 5))
            for i, (col, c) in enumerate(zip(pill_cols, comps[:5])):
                with col:
                    if st.button(f"{c} ×", key=f"rm_{i}", help=f"Remove {c}", type="secondary"):
                        st.session_state.competencies = [x for x in st.session_state.competencies if x != c]
                        st.rerun()
            if len(comps) > 5:
                extra_pills = st.columns(min(len(comps)-5, 5))
                for i, (col, c) in enumerate(zip(extra_pills, comps[5:])):
                    with col:
                        if st.button(f"{c} ×", key=f"rm_{i+5}", help=f"Remove {c}", type="secondary"):
                            st.session_state.competencies = [x for x in st.session_state.competencies if x != c]
                            st.rerun()
        nc1, nc2 = st.columns([5, 1])
        with nc1:
            new_comp = st.text_input("new_comp", placeholder="+ Add skill and press Add",
                                      label_visibility="collapsed", key="new_comp_input")
        with nc2:
            if st.button("Add", key="add_comp_btn", type="secondary"):
                if new_comp.strip() and new_comp.strip() not in st.session_state.competencies:
                    st.session_state.competencies.append(new_comp.strip())
                st.rerun()

    # ── CV upload card (file upload only) ─────────────────────────────────────
    n_cvs = len(st.session_state.cvs)
    with st.container(border=True):
        cv_hc1, cv_hc2 = st.columns([3, 1])
        with cv_hc1:
            st.markdown(f"**Candidate CVs** · <span style='color:#0075BC;font-weight:700'>{n_cvs} added</span>", unsafe_allow_html=True)
        with cv_hc2:
            st.caption("PDF, DOCX or TXT")

        uploaded = st.file_uploader("Drop CVs here", type=["pdf","docx","doc","txt"],
                                     accept_multiple_files=True, label_visibility="visible")
        if uploaded:
            new_files = [f for f in uploaded if f.name not in st.session_state.cvs]
            if new_files:
                with st.spinner(f"Reading {len(new_files)} file(s)…"):
                    errors = []
                    for f in new_files:
                        try:
                            f.seek(0)
                            st.session_state.cvs[f.name] = extract_file_text(f)
                        except Exception as e:
                            errors.append(f"{f.name}: {e}")
                if errors:
                    st.warning("Some files could not be read:\n" + "\n".join(errors))
                st.rerun()

        names = list(st.session_state.cvs.keys())
        if names:
            for nm in names:
                ext = "pdf" if nm.lower().endswith(".pdf") else "doc"
                tag = "PDF" if ext == "pdf" else "DOC"
                bg  = "#FEF0DC" if ext == "pdf" else "#E4EDF7"
                fg  = "#C76A0A" if ext == "pdf" else "#3251A3"
                st.markdown(
                    f'<div class="qs-file-chip">'
                    f'<span class="qs-file-icon" style="background:{bg};color:{fg}">{tag}</span>'
                    f'<span style="font-weight:600;font-size:12px;overflow:hidden;white-space:nowrap;text-overflow:ellipsis;max-width:200px">{nm}</span>'
                    f'<span style="color:#0075BC;font-size:13px;margin-left:auto">✓</span>'
                    f'</div>',
                    unsafe_allow_html=True)
            if st.button("Clear all CVs", key="clear_cvs", type="secondary"):
                st.session_state.cvs = {}
                st.rerun()

    # ── Footer ────────────────────────────────────────────────────────────────
    jd_ready   = bool(st.session_state.jd or jd_input)
    cvs_ready  = bool(st.session_state.cvs)
    n_screen   = len(st.session_state.cvs)
    can_screen = jd_ready and cvs_ready

    status_ph = st.empty()
    st.divider()

    fl, fr = st.columns([5, 2])
    with fl:
        if not jd_ready:
            st.caption("⚠️ Add a job description to continue")
        elif not cvs_ready:
            st.caption("⚠️ Upload at least one CV to continue")
        else:
            st.caption(f"Takes about 30–60 seconds for {n_screen} CV{'s' if n_screen != 1 else ''}")
    with fr:
        if st.button(f"Screen {n_screen} CVs →", type="primary",
                      disabled=not can_screen, key="screen_btn"):
            jd = st.session_state.jd or jd_input
            comps_str = "\n".join(st.session_state.competencies)
            raw = screen_cvs(jd, comps_str, st.session_state.cvs, status_placeholder=status_ph)
            if raw:
                raw = re.sub(r"^```[a-z]*\n?", "", raw.strip()).rstrip("` \n")
                try:
                    results = json.loads(raw)
                    # stamp initial shortlist status into history
                    for r in results:
                        if r.get("shortlisted"):
                            record_history(r, "auto-shortlisted")
                    st.session_state.screening_results = results
                    st.session_state.screen = "results"
                    st.rerun()
                except json.JSONDecodeError as e:
                    st.error(f"Could not parse Claude's response as JSON: {e}")
                    st.code(raw[:2000])

# ═══════════════════════════════════════════════════════════════════════════════
# RESULTS SCREEN
# ═══════════════════════════════════════════════════════════════════════════════
elif screen == "results":
    results_data  = st.session_state.screening_results or []
    strong_list   = [r for r in results_data if r.get("overall", 0) >= 85]
    possible_list = [r for r in results_data if 65 <= r.get("overall", 0) < 85]
    weak_list     = [r for r in results_data if r.get("overall", 0) < 65]
    role_label    = st.session_state.role_title or "Screening Results"
    filt          = st.session_state.filter

    # Sub-header
    auto_html(f"""
<link href="https://fonts.googleapis.com/css2?family=Work+Sans:wght@500;700;800&display=swap" rel="stylesheet">
<div style="padding:20px 0 14px;font-family:'Work Sans',sans-serif">
  <div style="font:500 12px 'Work Sans';color:#5E6675;margin-bottom:4px;text-align:left">&#8592; {role_label}</div>
  <div style="font:800 22px 'Work Sans';letter-spacing:-.02em;color:#1A1A2E;text-align:left">{len(results_data)} candidates screened, ranked</div>
</div>
""")

    # Band filter boxes (clickable) + search + sort
    b1, b2, b3, b4 = st.columns([1, 1, 1, 3])
    with b1:
        if st.button(
            f"{'▶ ' if filt == 'strong' else ''}**{len(strong_list)}** Strong",
            key="fband_strong",
            type="primary" if filt == "strong" else "secondary",
            help="Filter: Strong (≥85)"
        ):
            st.session_state.filter = "strong" if filt != "strong" else "all"
            st.rerun()
    with b2:
        if st.button(
            f"{'▶ ' if filt == 'possible' else ''}**{len(possible_list)}** Possible",
            key="fband_possible",
            type="primary" if filt == "possible" else "secondary",
            help="Filter: Possible (65–84)"
        ):
            st.session_state.filter = "possible" if filt != "possible" else "all"
            st.rerun()
    with b3:
        if st.button(
            f"{'▶ ' if filt == 'weak' else ''}**{len(weak_list)}** Weak",
            key="fband_weak",
            type="primary" if filt == "weak" else "secondary",
            help="Filter: Weak (<65)"
        ):
            st.session_state.filter = "weak" if filt != "weak" else "all"
            st.rerun()
    with b4:
        search = st.text_input("search", placeholder="🔍  Search candidates…",
                                label_visibility="collapsed", key="search_input")

    sort_col, all_col = st.columns([4, 1])
    with sort_col:
        sort = st.selectbox("sort", ["Best match", "Name A–Z", "Experience"],
                             label_visibility="collapsed", key="sort_sel")
    with all_col:
        if st.button("All", key="f_all", type="primary" if filt == "all" else "secondary"):
            st.session_state.filter = "all"; st.rerun()

    # Filter & sort
    filtered = results_data
    if filt == "strong":     filtered = strong_list
    elif filt == "possible": filtered = possible_list
    elif filt == "weak":     filtered = weak_list
    if search:
        filtered = [r for r in filtered if search.lower() in r.get("name","").lower()
                    or search.lower() in r.get("filename","").lower()]
    if "Name" in sort:   filtered = sorted(filtered, key=lambda r: r.get("name",""))
    elif "Exp" in sort:  filtered = sorted(filtered, key=lambda r: -r.get("years", 0))

    if not filtered:
        st.info("No candidates match this filter.")

    for i, r in enumerate(filtered):
        sc         = r.get("overall", 0)
        _, bbg, bcolor, _ = band(sc)
        ring_bg, ring_color = score_ring_colors(sc)
        av_color   = avatar_color(r.get("name", r.get("filename","")))
        name       = r.get("name", r.get("filename",""))
        inits      = initials(name)
        strengths  = r.get("strengths", [])
        tags       = strengths[:2]
        extra_tags = len(strengths) - 2

        tags_html = "".join(
            f'<span style="background:#E3F1FA;color:#005A91;border-radius:6px;padding:3px 8px;'
            f'font-weight:600;font-size:11px;white-space:nowrap">{t}</span>'
            for t in tags)
        if extra_tags > 0:
            tags_html += (f'<span style="background:#E4E9EF;color:#5E6675;border-radius:6px;'
                          f'padding:3px 7px;font-weight:600;font-size:11px;white-space:nowrap">+{extra_tags}</span>')

        # Current status badge + calibration badge
        is_sl = r.get("shortlisted")
        ckey_r = candidate_key(r)
        fb_r   = st.session_state.score_feedback.get(ckey_r)
        if is_sl is True:
            status_badge = '<span style="background:#E3F1FA;color:#005A91;border-radius:5px;padding:2px 7px;font-size:10px;font-weight:700">★ Shortlisted</span>'
        elif is_sl is False:
            status_badge = '<span style="background:#FDE7DE;color:#C23A18;border-radius:5px;padding:2px 7px;font-size:10px;font-weight:700">✕ Rejected</span>'
        else:
            status_badge = ""
        if fb_r:
            cal_label = "✓ Approved" if fb_r.get("approved") else "✎ Calibrated"
            status_badge += f'<span style="background:#F0FDF4;color:#1B6E2E;border-radius:5px;padding:2px 7px;font-size:10px;font-weight:700;margin-left:4px">{cal_label}</span>'

        # SVG donut ring
        r_px  = 22
        circ  = 2 * 3.14159 * r_px
        dash_offset = circ * (1 - sc / 100)
        donut = (
            f'<svg width="54" height="54" viewBox="0 0 54 54" '
            f'style="position:absolute;top:0;left:0;transform:rotate(-90deg)">'
            f'<circle cx="27" cy="27" r="{r_px}" fill="none" stroke="#E4E9EF" stroke-width="5"/>'
            f'<circle cx="27" cy="27" r="{r_px}" fill="none" stroke="{ring_color}" stroke-width="5" '
            f'stroke-dasharray="{circ:.1f}" stroke-dashoffset="{dash_offset:.1f}" stroke-linecap="round"/>'
            f'</svg>'
        )

        # Find global index in results_data for this candidate
        global_idx = next((j for j, x in enumerate(results_data) if candidate_key(x) == candidate_key(r)), i)

        col_main, col_actions = st.columns([8, 2])
        with col_main:
            auto_html(f"""
<link href="https://fonts.googleapis.com/css2?family=Work+Sans:wght@500;700;800&display=swap" rel="stylesheet">
<style>body{{margin:0;padding:0}}</style>
<div style="display:flex;align-items:flex-start;gap:14px;background:#fff;border:1px solid #E4E9EF;border-left:4px solid {ring_color};border-radius:12px;padding:12px 16px;font-family:'Work Sans',sans-serif">
  <div style="font:800 14px 'Work Sans';color:#C2C8D2;width:18px;padding-top:2px;text-align:left;flex:none">{r.get('rank',i+1)}</div>
  <div style="width:40px;height:40px;border-radius:50%;background:{av_color};color:#fff;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:14px;flex:none;margin-top:2px">{inits}</div>
  <div style="flex:1;min-width:0">
    <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap">
      <span style="font-weight:700;font-size:14px;color:#1A1A2E">{name}</span>
      {status_badge}
    </div>
    <div style="font:500 12px 'Work Sans';color:#5E6675;margin-top:3px;text-align:left;line-height:1.4">{r.get('role','')} &middot; {r.get('years','')} yrs &middot; {r.get('location','')}</div>
    <div style="display:flex;gap:5px;margin-top:6px;flex-wrap:wrap">{tags_html}</div>
  </div>
  <div style="position:relative;width:54px;height:54px;flex:none;margin-top:2px">
    {donut}
    <div style="position:absolute;top:0;left:0;width:54px;height:54px;display:flex;align-items:center;justify-content:center">
      <span style="font-weight:800;font-size:15px;color:{ring_color};line-height:1">{sc}</span>
    </div>
  </div>
</div>
""")

        with col_actions:
            a1, a2, a3 = st.columns(3)
            is_sl = r.get("shortlisted") is True
            is_rj = r.get("shortlisted") is False
            with a1:
                if st.button("→", key=f"view_{i}", help="View detail", type="secondary"):
                    st.session_state.selected_idx = global_idx
                    st.session_state.screen = "detail"
                    st.rerun()
            with a2:
                # title="undo-shortlist" triggers green CSS when active
                sl_title = "undo-shortlist" if is_sl else "Shortlist candidate"
                if st.button("★", key=f"sl_{i}", help=sl_title, type="secondary"):
                    if is_sl:
                        st.session_state.screening_results[global_idx]["shortlisted"] = None
                        record_history(r, "un-shortlisted")
                    else:
                        st.session_state.screening_results[global_idx]["shortlisted"] = True
                        record_history(r, "shortlisted")
                    st.rerun()
            with a3:
                # title="undo-reject" triggers red CSS when active
                rj_title = "undo-reject" if is_rj else "Reject candidate"
                if st.button("✕", key=f"rj_{i}", help=rj_title, type="secondary"):
                    if is_rj:
                        st.session_state.screening_results[global_idx]["shortlisted"] = None
                        record_history(r, "un-rejected")
                    else:
                        st.session_state.screening_results[global_idx]["shortlisted"] = False
                        record_history(r, "rejected")
                    st.rerun()

    st.divider()

    # Candidate history
    history = st.session_state.candidate_history
    if history:
        with st.expander(f"📋 Candidate status history ({sum(len(v) for v in history.values())} actions)"):
            all_entries = []
            for key, entries in history.items():
                for e in entries:
                    all_entries.append(e)
            # Sort by most recent first (ts string is "DD Mon YYYY, HH:MM")
            all_entries.sort(key=lambda e: e["ts"], reverse=True)
            for e in all_entries:
                action = e["action"]
                if action == "shortlisted":
                    icon, color = "★", "#005A91"
                elif action == "rejected":
                    icon, color = "✕", "#C23A18"
                else:
                    icon, color = "•", "#5E6675"
                st.markdown(
                    f'<div style="display:flex;align-items:center;gap:10px;padding:8px 0;border-bottom:1px solid #F0F2F5">'
                    f'<span style="color:{color};font-weight:700;font-size:14px;width:16px">{icon}</span>'
                    f'<span style="font-weight:600;font-size:13px;color:#1A1A2E;flex:1">{e["name"]}</span>'
                    f'<span style="font-size:12px;color:#5E6675;font-weight:500">{action.capitalize()}</span>'
                    f'<span style="font-size:11px;color:#9AA1AE;margin-left:12px">{e["ts"]}</span>'
                    f'</div>',
                    unsafe_allow_html=True)

    c1, c2, c3, _ = st.columns([2, 2, 2, 2])
    with c1:
        if st.button("← New screening", key="back_to_setup"):
            st.session_state.screen = "setup"
            st.session_state.screening_results = None
            st.rerun()
    with c2:
        pdf = generate_pdf(results_data, st.session_state.role_title)
        st.download_button("📄 Download PDF", data=pdf,
            file_name=f"CV_Screening_{datetime.now().strftime('%Y-%m-%d')}.pdf",
            mime="application/pdf", type="primary")
    with c3:
        xlsx = generate_excel(results_data, st.session_state.role_title,
                              st.session_state.candidate_history)
        st.download_button("📊 Download Excel", data=xlsx,
            file_name=f"CV_Screening_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="secondary")

# ═══════════════════════════════════════════════════════════════════════════════
# DETAIL SCREEN
# ═══════════════════════════════════════════════════════════════════════════════
elif screen == "detail":
    results_data = st.session_state.screening_results or []
    idx = st.session_state.selected_idx
    if not results_data or idx >= len(results_data):
        st.session_state.screen = "results"; st.rerun()

    r           = results_data[idx]
    sc          = r.get("overall", 0)
    name        = r.get("name", r.get("filename", ""))
    inits       = initials(name)
    av_color    = avatar_color(name)
    ring_bg, ring_color = score_ring_colors(sc)
    bname, bbg, bcolor, _ = band(sc)
    band_label  = bname.upper()
    scores      = r.get("scores", [])
    labels      = r.get("competency_labels", [f"Competency {i+1}" for i in range(len(scores))])

    # Nav bar
    nav1, _, nav3 = st.columns([3, 4, 3])
    with nav1:
        if st.button("← Back to candidates", key="back_results"):
            st.session_state.screen = "results"; st.rerun()
    with nav3:
        pc1, pc2 = st.columns(2)
        with pc1:
            if idx > 0 and st.button("↑ Prev", key="prev_btn"):
                st.session_state.selected_idx -= 1; st.rerun()
        with pc2:
            if idx < len(results_data)-1 and st.button("Next ↓", key="next_btn"):
                st.session_state.selected_idx += 1; st.rerun()

    # Candidate header
    email_html = f'<span style="text-align:left">✉ {r["email"]}</span>' if r.get("email") else ""
    phone_html = f'<span style="text-align:left">📞 {r["phone"]}</span>' if r.get("phone") else ""

    # SVG donut for detail
    r_px  = 30
    circ  = 2 * 3.14159 * r_px
    dash_offset = circ * (1 - sc / 100)
    detail_donut = (
        f'<svg width="74" height="74" viewBox="0 0 74 74" style="position:absolute;top:0;left:0;transform:rotate(-90deg)">'
        f'<circle cx="37" cy="37" r="{r_px}" fill="none" stroke="#E4E9EF" stroke-width="6"/>'
        f'<circle cx="37" cy="37" r="{r_px}" fill="none" stroke="{ring_color}" stroke-width="6" '
        f'stroke-dasharray="{circ:.1f}" stroke-dashoffset="{dash_offset:.1f}" stroke-linecap="round"/>'
        f'</svg>'
    )

    auto_html(f"""
<link href="https://fonts.googleapis.com/css2?family=Work+Sans:wght@500;700;800&display=swap" rel="stylesheet">
<style>body{{margin:0;padding:0}}</style>
<div style="background:#fff;border:1px solid #E4E9EF;border-radius:14px;display:flex;align-items:flex-start;gap:20px;font-family:'Work Sans',sans-serif;padding:16px 20px">
  <div style="width:52px;height:52px;border-radius:50%;background:{av_color};color:#fff;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:22px;flex:none">{inits}</div>
  <div style="flex:1;min-width:0">
    <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap">
      <div style="font:800 20px 'Work Sans';letter-spacing:-.02em;color:#1A1A2E;text-align:left">{name}</div>
      <span style="background:{bbg};color:{bcolor};border-radius:999px;padding:3px 10px;font-weight:700;font-size:11px">{band_label}</span>
    </div>
    <div style="font:500 13px 'Work Sans';color:#5E6675;margin-top:3px;text-align:left;line-height:1.4">{r.get('role','')} &middot; {r.get('years','')} yrs &middot; {r.get('location','')}</div>
    <div style="display:flex;gap:14px;margin-top:6px;font:500 12px 'Work Sans';color:#3A4150;flex-wrap:wrap">{email_html}{phone_html}</div>
  </div>
  <div style="position:relative;width:74px;height:74px;flex:none">
    {detail_donut}
    <div style="position:absolute;top:0;left:0;width:74px;height:74px;display:flex;flex-direction:column;align-items:center;justify-content:center">
      <div style="font-weight:800;font-size:20px;color:{ring_color};line-height:1">{sc}</div>
      <div style="font:600 9px 'Work Sans';color:#9AA1AE">/100</div>
    </div>
  </div>
</div>
""")

    # Action buttons
    is_sl_detail = r.get("shortlisted") is True
    is_rj_detail = r.get("shortlisted") is False
    ac1, ac2, ac3, _ = st.columns([2, 2, 2, 4])
    with ac1:
        sl_title_d = "undo-shortlist" if is_sl_detail else "Shortlist candidate"
        sl_label   = "★ Shortlisted" if is_sl_detail else "★ Shortlist"
        if st.button(sl_label, key="sl_btn", help=sl_title_d,
                     type="primary" if not is_sl_detail else "secondary"):
            if is_sl_detail:
                st.session_state.screening_results[idx]["shortlisted"] = None
                record_history(r, "un-shortlisted")
            else:
                st.session_state.screening_results[idx]["shortlisted"] = True
                record_history(r, "shortlisted")
            st.rerun()
    with ac2:
        rj_title_d = "undo-reject" if is_rj_detail else "Reject candidate"
        rj_label   = "✕ Rejected" if is_rj_detail else "✕ Reject"
        if st.button(rj_label, key="rj_btn", help=rj_title_d, type="secondary"):
            if is_rj_detail:
                st.session_state.screening_results[idx]["shortlisted"] = None
                record_history(r, "un-rejected")
            else:
                st.session_state.screening_results[idx]["shortlisted"] = False
                record_history(r, "rejected")
            st.rerun()
    with ac3:
        pdf_single = generate_pdf([r], st.session_state.role_title)
        st.download_button("⤓ Export PDF", data=pdf_single,
            file_name=f"{name.replace(' ','_')}_Report.pdf",
            mime="application/pdf", type="secondary")

    # Show history for this candidate
    ckey = candidate_key(r)
    chistory = st.session_state.candidate_history.get(ckey, [])
    if chistory:
        st.caption("Status history: " + "  ·  ".join(
            f"{e['action'].capitalize()} at {e['ts']}" for e in reversed(chistory)
        ))

    st.divider()

    left_col, right_col = st.columns([1.2, 1])

    with left_col:
        # ── Score calibration panel ───────────────────────────────────────────
        ckey_d = candidate_key(r)
        existing_fb = st.session_state.score_feedback.get(ckey_d)

        if existing_fb:
            approved_txt = "✓ Scores approved" if existing_fb.get("approved") else "✎ Scores calibrated by reviewer"
            approved_color = "#1B6E2E" if existing_fb.get("approved") else "#0075BC"
            auto_html(f"""
<style>body{{margin:0;padding:0}}</style>
<div style="background:#F0FDF4;border:1px solid #BBF7D0;border-radius:10px;padding:10px 14px;font-family:'Work Sans',sans-serif">
  <div style="font-weight:700;font-size:12px;color:{approved_color}">{approved_txt}</div>
  <div style="font:500 11px 'Work Sans';color:#5E6675;margin-top:3px;line-height:1.5">
    Overall: AI={existing_fb['ai_overall']} → Reviewer={existing_fb['human_overall']}
    {"  ·  " + existing_fb['reason'] if existing_fb.get('reason') else ""}
    · {existing_fb.get('ts','')}
  </div>
</div>
""")

        with st.expander("🎯 Review & calibrate AI scores", expanded=not bool(existing_fb)):
            st.caption("Adjust scores if the AI got something wrong — your corrections help calibrate future screenings.")

            # Load prior feedback values as defaults if they exist
            prior_overall = existing_fb["human_overall"] if existing_fb else sc
            prior_comp    = existing_fb["human_scores"]   if existing_fb else list(scores)

            new_overall = st.slider(
                "Overall match score", 0, 100, prior_overall,
                key=f"cal_overall_{idx}",
                help="Drag to your assessed score for this candidate"
            )
            new_comp_scores = []
            for j, (label, val) in enumerate(zip(labels, prior_comp)):
                default_val = prior_comp[j] if j < len(prior_comp) else (scores[j] if j < len(scores) else 50)
                new_val = st.slider(
                    label, 0, 100, int(default_val),
                    key=f"cal_comp_{idx}_{j}"
                )
                new_comp_scores.append(new_val)

            cal_reason = st.text_area(
                "Reason for adjustment (optional)",
                value=existing_fb.get("reason", "") if existing_fb else "",
                placeholder="e.g. Strong field experience not reflected in CV text; personally know this candidate's work…",
                height=80, key=f"cal_reason_{idx}",
                label_visibility="visible"
            )

            cal1, cal2 = st.columns(2)
            with cal1:
                if st.button("✓ Approve AI scores as-is", key=f"cal_approve_{idx}", type="secondary"):
                    record_feedback(r, sc, list(scores), labels,
                                    sc, list(scores), "", approved=True)
                    # Sync overall back to original
                    st.session_state.screening_results[idx]["overall"] = sc
                    st.success("Scores approved and recorded.")
                    st.rerun()
            with cal2:
                if st.button("💾 Save my adjustments", key=f"cal_save_{idx}", type="primary"):
                    record_feedback(r, sc, list(scores), labels,
                                    new_overall, new_comp_scores,
                                    cal_reason.strip(), approved=False)
                    # Write adjusted scores back into the result
                    st.session_state.screening_results[idx]["overall"] = new_overall
                    st.session_state.screening_results[idx]["scores"] = new_comp_scores
                    st.success("Adjustments saved. Future screenings will learn from this.")
                    st.rerun()

        st.markdown("**Match breakdown**")
        # Use adjusted scores if available
        display_scores = existing_fb["human_scores"] if existing_fb and not existing_fb.get("approved") else scores
        display_overall = existing_fb["human_overall"] if existing_fb else sc
        for label, val in zip(labels, display_scores):
            color = "#0075BC" if val >= 85 else "#F7941D" if val >= 65 else "#E85020"
            auto_html(f"""
<style>body{{margin:0;padding:0}}</style>
<div style="font-family:'Work Sans',sans-serif;padding-bottom:4px">
  <div style="display:flex;justify-content:space-between;font-weight:600;font-size:13px;margin-bottom:6px;color:#1A1A2E">
    <span>{label}</span><span style="color:{color}">{val}</span>
  </div>
  <div style="height:8px;background:#E4E9EF;border-radius:999px;overflow:hidden">
    <div style="height:100%;width:{val}%;background:{color};border-radius:999px"></div>
  </div>
</div>
""")

        if r.get("flag"):
            auto_html(f"""
<style>body{{margin:0;padding:0}}</style>
<div style="background:#FEF0DC;border-radius:11px;padding:12px 14px;display:flex;gap:10px;font-family:'Work Sans',sans-serif">
  <div style="font-size:15px;flex:none">⚠</div>
  <div>
    <div style="font-weight:700;font-size:12px;color:#B0640C">One thing to verify</div>
    <div style="font:500 12px/1.6 'Work Sans';color:#9A6410;margin-top:4px">{r['flag']}</div>
  </div>
</div>
""")

        st.markdown("**Gaps / Concerns**")
        for g in r.get("gaps", []):
            st.markdown(f"- {g}")

    with right_col:
        auto_html(f"""
<link href="https://fonts.googleapis.com/css2?family=Work+Sans:wght@500;700&display=swap" rel="stylesheet">
<style>body{{margin:0;padding:0}}</style>
<div style="background:#0075BC;border-radius:12px;padding:14px 16px;color:#fff;font-family:'Work Sans',sans-serif">
  <div style="font-weight:700;font-size:12px;margin-bottom:7px">&#10022; AI summary</div>
  <div style="font:500 13px/1.6 'Work Sans';color:rgba(255,255,255,.92)">{r.get('summary','')}</div>
</div>
""")

        st.markdown("**Evidence from CV**")
        for ev in r.get("evidence", []):
            auto_html(f"""
<style>body{{margin:0;padding:0}}</style>
<div style="background:#fff;border:1px solid #E4E9EF;border-radius:10px;padding:10px 13px;margin-bottom:2px;font-family:'Work Sans',sans-serif">
  <div style="font:600 10px 'JetBrains Mono',monospace;color:#0075BC;margin-bottom:4px">{ev.get('label','')}</div>
  <div style="font:500 12px/1.6 'Work Sans';color:#3A4150">{ev.get('text','')}</div>
</div>
""")
