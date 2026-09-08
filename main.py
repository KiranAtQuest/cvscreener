import os, io, re, json
from datetime import datetime
from typing import List, Optional

import anthropic
import pdfplumber
from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable, Table, TableStyle
from fastapi import Cookie, Depends, FastAPI, UploadFile, File, Form, HTTPException, Response as FResponse
from fastapi.responses import Response, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

import auth as _auth

app = FastAPI(title="CV Screener – Quest Alliance")

# ── Startup ────────────────────────────────────────────────────────────────────
@app.on_event("startup")
def startup():
    _auth.init_db()

# ── Auth routes ────────────────────────────────────────────────────────────────

@app.post("/api/auth/login")
async def login(
    username: str = Form(...),
    password: str = Form(...),
    response: FResponse = None,
):
    user = _auth.get_user_by_credentials(username, password)
    if not user:
        raise HTTPException(status_code=401, detail="Invalid username or password")
    token = _auth.create_token(user["id"], user["username"], user["role"])
    resp = JSONResponse({"username": user["username"], "role": user["role"], "email": user["email"]})
    resp.set_cookie(
        _auth.COOKIE, token,
        httponly=True, samesite="lax", secure=False,  # set secure=True behind HTTPS
        max_age=_auth.TOKEN_TTL * 3600,
    )
    return resp

@app.post("/api/auth/logout")
async def logout():
    resp = JSONResponse({"ok": True})
    resp.delete_cookie(_auth.COOKIE)
    return resp

@app.get("/api/auth/me")
async def me(qs_token: Optional[str] = Cookie(default=None)):
    if not qs_token:
        raise HTTPException(status_code=401, detail="Not authenticated")
    user = _auth.get_current_user(qs_token)
    return {"username": user["username"], "role": user["role"], "email": user["email"]}

# ── Admin routes ───────────────────────────────────────────────────────────────

class NewUser(BaseModel):
    username: str
    email: str
    password: str
    role: str = "recruiter"

class UpdateUser(BaseModel):
    role: Optional[str] = None
    active: Optional[bool] = None
    password: Optional[str] = None

@app.get("/api/admin/users")
async def admin_list_users(qs_token: Optional[str] = Cookie(default=None)):
    _auth.require_admin(qs_token=qs_token)
    return _auth.list_users()

@app.post("/api/admin/users")
async def admin_create_user(body: NewUser, qs_token: Optional[str] = Cookie(default=None)):
    _auth.require_admin(qs_token=qs_token)
    return _auth.create_user(body.username, body.email, body.password, body.role)

@app.patch("/api/admin/users/{uid}")
async def admin_update_user(uid: int, body: UpdateUser, qs_token: Optional[str] = Cookie(default=None)):
    _auth.require_admin(qs_token=qs_token)
    return _auth.update_user(uid, body.role, body.active, body.password)

@app.delete("/api/admin/users/{uid}")
async def admin_delete_user(uid: int, qs_token: Optional[str] = Cookie(default=None)):
    admin = _auth.require_admin(qs_token=qs_token)
    _auth.delete_user(uid, admin["id"])
    return {"ok": True}

@app.get("/api/admin/errors")
async def admin_get_errors(qs_token: Optional[str] = Cookie(default=None)):
    _auth.require_admin(qs_token=qs_token)
    return _auth.get_upload_errors()

class ErrorStatusUpdate(BaseModel):
    status: str = "resolved"

@app.patch("/api/admin/errors/{error_id}")
async def admin_update_error(error_id: int, body: ErrorStatusUpdate,
                              qs_token: Optional[str] = Cookie(default=None)):
    _auth.require_admin(qs_token=qs_token)
    _auth.update_error_status(error_id, body.status)
    return {"ok": True}

# ── Calibration notes routes ───────────────────────────────────────────────────

class CalibNote(BaseModel):
    note: str
    jd_hash: str = ""

@app.get("/api/calibration")
async def get_calibration(jd_hash: str = "", qs_token: Optional[str] = Cookie(default=None)):
    _auth.get_current_user(qs_token)
    return _auth.get_calibration_notes(jd_hash)

@app.post("/api/calibration")
async def add_calibration(body: CalibNote, qs_token: Optional[str] = Cookie(default=None)):
    user = _auth.get_current_user(qs_token)
    return _auth.add_calibration_note(body.note, user["username"], body.jd_hash)

@app.delete("/api/calibration/{note_id}")
async def delete_calibration(note_id: int, qs_token: Optional[str] = Cookie(default=None)):
    _auth.get_current_user(qs_token)
    _auth.delete_calibration_note(note_id)
    return {"ok": True}

# ── File parsing ───────────────────────────────────────────────────────────────

def _plain_error(filename: str, exc: Exception):
    msg = str(exc).lower()
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else "unknown"
    if "password" in msg or "encrypt" in msg:
        return (
            f"The file '{filename}' is password-protected and cannot be opened.",
            "Remove the password from the file before uploading. In Word/Acrobat go to File → Protect/Security and remove the password."
        )
    if "corrupt" in msg or "invalid" in msg or "bad" in msg:
        return (
            f"The file '{filename}' appears to be damaged or incomplete.",
            "Try re-saving from the original app (Word, Acrobat) and upload again."
        )
    if ext == "pdf":
        return (
            f"The PDF '{filename}' could not be read. It may be a scanned image without selectable text.",
            "Use an OCR tool (Adobe Acrobat, online2pdf.com) to convert the scanned PDF to text-based PDF before uploading."
        )
    if ext in ("doc", "docx"):
        return (
            f"The Word document '{filename}' could not be opened.",
            "Re-save the file as .docx in Microsoft Word or Google Docs, then upload again."
        )
    return (
        f"The file '{filename}' could not be read ({ext.upper()} format).",
        "Try converting the file to PDF or DOCX and uploading again. If the problem continues, contact your administrator."
    )

def extract_text_from_pdf(b: bytes) -> str:
    with pdfplumber.open(io.BytesIO(b)) as p:
        return "\n".join(pg.extract_text() or "" for pg in p.pages)

def extract_text_from_docx(b: bytes) -> str:
    return "\n".join(para.text for para in Document(io.BytesIO(b)).paragraphs)

def parse_bytes(b: bytes, name: str) -> str:
    n = name.lower()
    if n.endswith(".pdf"):            return extract_text_from_pdf(b)
    if n.endswith((".docx", ".doc")): return extract_text_from_docx(b)
    return b.decode("utf-8", errors="replace")

# ── Helpers ────────────────────────────────────────────────────────────────────

def candidate_key(r: dict) -> str:
    return r.get("filename") or r.get("name") or str(r.get("rank", ""))

def band(score: int):
    if score >= 85: return "strong",   "#E3F1FA", "#005A91", "#0075BC"
    if score >= 65: return "possible", "#FEF0DC", "#C76A0A", "#F7941D"
    return "weak", "#FDE7DE", "#C23A18", "#F15A29"

def get_api_key() -> str:
    key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not key:
        raise HTTPException(status_code=500, detail="ANTHROPIC_API_KEY not set on server.")
    return key

# ── Prompt ─────────────────────────────────────────────────────────────────────

def build_prompt(jd: str, competencies: str, cvs: dict, calibration: list,
                 past_examples: list = None) -> str:
    cv_block = "".join(
        f"\n---\nCV #{i} – {name}\n{text}\n"
        for i, (name, text) in enumerate(cvs.items(), 1)
    )
    calibration_block = ""
    if calibration:
        calibration_block = (
            "\n## Reviewer Calibration Notes\n"
            "Use these organisational preferences when scoring.\n"
            + "\n".join(f"- {ex}" for ex in calibration[-10:]) + "\n"
        )
    examples_block = ""
    if past_examples:
        lines = []
        for ex in past_examples[:15]:
            dec  = ex.get("final_decision", "").upper()
            score = ex.get("ai_score", "?")
            name  = ex.get("candidate_name", "Candidate")
            summ  = ex.get("summary", "")
            note  = ex.get("recruiter_note", "")
            line  = f"- {name} | AI score {score} → {dec}"
            if summ: line += f" | {summ[:120]}"
            if note: line += f" | Recruiter note: {note}"
            lines.append(line)
        examples_block = (
            "\n## Past Hiring Decisions for This Role (learn from these)\n"
            "These are real recruiter decisions for the same role. Calibrate your scoring "
            "so that candidates similar to SHORTLISTED examples score ≥65 and candidates "
            "similar to REJECTED examples score <65.\n"
            + "\n".join(lines) + "\n"
        )
    return f"""You are an expert HR screener for Quest Alliance, an NGO focused on youth skilling in India.

## Screening Process — Two Levels
This screening operates in two levels:

LEVEL 1 — FILTER (automatic): Candidates with an overall match score below 30 are junk
applications that do not meet minimum requirements and should not consume recruiter time.
Set "filtered": true for these candidates. Provide only minimal details for filtered candidates.

LEVEL 2 — RANK (for human review): Candidates scoring 30 or above are eligible and must be
ranked carefully for recruiter decision-making. Within eligible candidates:
  - Score 65–100: Strong/Possible match — recommend for shortlisting
  - Score 30–64: Weak match — eligible but recruiter should review carefully before deciding

## Job Description
{jd}

## Required Skill Competencies
{competencies or "(derive from JD)"}

## Universal Scoring Factor (applies to ALL roles)
Experience working with non-profit, NGO, social sector, or development organisations must be treated
as a significant positive signal for every candidate, regardless of role. Candidates with such
experience should score meaningfully higher (add 5-10 points to overall) than equally qualified
candidates without it. Reflect this in the "strengths" array and factor it into competency scores.
{calibration_block}{examples_block}
## Candidate CVs
{cv_block}

## Task
Review every candidate and return a JSON array sorted best-to-worst (highest overall score first,
filtered candidates at the end). Each object MUST have:
- "rank": integer from 1 (rank among ALL candidates including filtered)
- "filename": CV filename exactly as given
- "name": candidate's full name (extract from CV)
- "role": their current/most recent role title
- "years": integer years of total experience
- "location": city/state from CV
- "email": email address if present, else ""
- "phone": phone if present, else ""
- "overall": integer 0-100 match score
- "filtered": true if overall < 30 (Level 1 filter — junk application), else false
- "scores": array of 5 integers (one per competency) — empty array [] for filtered candidates
- "competency_labels": array of 5 strings (competency names) — empty array [] for filtered candidates
- "shortlisted": true if overall >= 65 AND filtered is false, else false
- "summary": for eligible candidates: 2-3 sentence fit summary; for filtered: one sentence why they don't meet minimum requirements
- "strengths": array of 2-4 short strings — empty [] for filtered candidates
- "gaps": array of 1-3 gap strings (for filtered, list the critical missing requirements)
- "flag": one-sentence verification note, or null
- "evidence": array of 2-3 objects with "label" and "text" — empty [] for filtered candidates

Return ONLY the JSON array, no markdown fences."""

def extract_json(raw: str) -> list:
    cleaned = re.sub(r"^```[a-z]*\n?", "", raw.strip()).rstrip("` \n")
    if not cleaned.strip():
        cleaned = raw.strip()
    m = re.search(r'\[[\s\S]*\]', cleaned)
    if m:
        cleaned = m.group(0)
    return json.loads(cleaned)

# ── PDF export ─────────────────────────────────────────────────────────────────

def generate_pdf(results: list, role_title: str = "") -> bytes:
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
    story = [Paragraph("CV Screening Report", T)]
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
        ct = Table([[Paragraph("Strengths", LB), Paragraph("Gaps", LB)],
                    [Paragraph(" · ".join(r.get("strengths",[])) or "—", BD),
                     Paragraph(" · ".join(r.get("gaps",[])) or "—", BD)]],
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

# ── Excel export ───────────────────────────────────────────────────────────────

def generate_excel(results: list, role_title: str = "", history: dict = None, score_feedback: dict = None) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    score_feedback = score_feedback or {}
    wb = Workbook()

    BLUE  = "FF0075BC"; GREEN = "FF1B6E2E"; RED = "FFC62828"
    AMBER = "FFE65100"; LGREY = "FFF4F7FA"; WHITE = "FFFFFFFF"; DARK = "FF1A1A2E"
    COMP_HDR = "FF1565C0"

    hdr_font  = Font(name="Calibri", bold=True, color=WHITE, size=11)
    hdr_fill  = PatternFill("solid", fgColor=DARK)
    comp_fill = PatternFill("solid", fgColor=COMP_HDR)
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin      = Side(style="thin", color="FFE4E9EF")
    border    = Border(left=thin, right=thin, top=thin, bottom=thin)

    eligible = [r for r in results if not r.get("filtered")]
    filtered = [r for r in results if r.get("filtered")]

    # Derive competency labels from first eligible candidate that has them
    comp_labels = []
    for r in eligible:
        lbs = r.get("competency_labels", [])
        if lbs:
            comp_labels = lbs
            break

    def comp_color(sc):
        if sc >= 70: return GREEN
        if sc >= 40: return AMBER
        return RED

    def band_color(sc):
        if sc >= 65: return GREEN
        if sc >= 30: return AMBER
        return RED

    # ── Sheet 1: Screening Results ──────────────────────────────────────────────
    ws = wb.active
    ws.title = "Screening Results"

    title_text = f"CV Screening Report — {role_title}" if role_title else "CV Screening Report"
    ws.append([title_text, "", "", "", "", "", f"Generated: {datetime.now().strftime('%d %B %Y')}"])
    ws["A1"].font = Font(name="Calibri", bold=True, size=14, color=BLUE)
    ws.merge_cells("A1:F1")
    ws["G1"].font = Font(name="Calibri", size=10, color="FF5E6675")
    ws["G1"].alignment = Alignment(horizontal="right")
    ws.append([f"Eligible for review: {len(eligible)}   |   Auto-filtered (below 30%): {len(filtered)}"])
    ws["A2"].font = Font(name="Calibri", size=10, color="FF5E6675")
    ws.append([])

    fixed_left  = ["Rank", "Name", "Current Role", "Yrs Exp", "Location",
                   "Overall Score (/100)", "Band", "Status"]
    fixed_right = ["Evidence from CV", "Strengths", "Gaps", "Fit Summary",
                   "Flag / Verify", "Email", "Phone", "Filename"]
    columns = fixed_left + [f"{lb}\n(/100)" for lb in comp_labels] + fixed_right

    ws.append(columns)
    hdr_row = ws.max_row
    for col_idx, _ in enumerate(columns, 1):
        is_comp = len(fixed_left) < col_idx <= len(fixed_left) + len(comp_labels)
        cell = ws.cell(row=hdr_row, column=col_idx)
        cell.font = hdr_font
        cell.fill = comp_fill if is_comp else hdr_fill
        cell.alignment = hdr_align
        cell.border = border
    ws.row_dimensions[hdr_row].height = 36

    for r in eligible + filtered:
        sc    = r.get("overall", 0)
        bname, _, _, _ = band(sc)
        sl    = r.get("shortlisted")
        is_f  = r.get("filtered", False)
        status = "Auto-filtered" if is_f else ("Shortlisted" if sl is True else "Rejected" if sl is False else "Pending")
        scores = r.get("scores", [])
        evidence_text = "\n".join(
            f"{ev.get('label','')}: {ev.get('text','')}"
            for ev in r.get("evidence", []) if ev.get("text")
        )
        row_vals = (
            [r.get("rank",""), r.get("name", r.get("filename","")),
             r.get("role",""), r.get("years",""), r.get("location",""),
             sc, bname.capitalize(), status]
            + [scores[i] if i < len(scores) else "" for i in range(len(comp_labels))]
            + [evidence_text,
               "; ".join(r.get("strengths",[])), "; ".join(r.get("gaps",[])),
               r.get("summary",""), r.get("flag","") or "",
               r.get("email",""), r.get("phone",""), r.get("filename","")]
        )
        ws.append(row_vals)
        data_row = ws.max_row

        for col_idx in range(1, len(columns)+1):
            cell = ws.cell(row=data_row, column=col_idx)
            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            cell.border = border
            if data_row % 2 == 0:
                cell.fill = PatternFill("solid", fgColor=LGREY)

        # Overall score: color-coded
        sc_cell = ws.cell(row=data_row, column=6)
        sc_cell.font = Font(name="Calibri", bold=True, color=WHITE, size=11)
        sc_cell.fill = PatternFill("solid", fgColor=band_color(sc))
        sc_cell.alignment = Alignment(horizontal="center", vertical="top")

        # Band cell
        b_cell = ws.cell(row=data_row, column=7)
        b_cell.font = Font(name="Calibri", bold=True, color=WHITE, size=10)
        b_cell.fill = PatternFill("solid", fgColor=band_color(sc))

        # Status cell
        st_cell = ws.cell(row=data_row, column=8)
        if sl is True:
            st_cell.font = Font(name="Calibri", bold=True, color=WHITE)
            st_cell.fill = PatternFill("solid", fgColor=GREEN)
        elif sl is False:
            st_cell.font = Font(name="Calibri", bold=True, color=WHITE)
            st_cell.fill = PatternFill("solid", fgColor=RED)
        elif is_f:
            st_cell.font = Font(name="Calibri", italic=True, color="FF888888")

        # Competency score cells: individually color-coded
        for i in range(len(comp_labels)):
            col_idx = len(fixed_left) + 1 + i
            cell = ws.cell(row=data_row, column=col_idx)
            if isinstance(cell.value, int):
                cell.font = Font(name="Calibri", bold=True, color=WHITE, size=10)
                cell.fill = PatternFill("solid", fgColor=comp_color(cell.value))
                cell.alignment = Alignment(horizontal="center", vertical="top")

    left_widths  = [5, 22, 22, 8, 16, 10, 10, 12]
    comp_widths  = [14] * len(comp_labels)
    right_widths = [45, 35, 28, 55, 30, 22, 14, 24]
    for i, w in enumerate(left_widths + comp_widths + right_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = f"A{hdr_row+1}"

    # ── Sheet 2: Evidence Detail ────────────────────────────────────────────────
    ws2 = wb.create_sheet("Evidence Detail")
    ws2.append([title_text, "", "", f"Generated: {datetime.now().strftime('%d %B %Y')}"])
    ws2["A1"].font = Font(name="Calibri", bold=True, size=13, color=BLUE)
    ws2.merge_cells("A1:C1")
    ws2.append(["Verbatim CV evidence supporting each candidate's score — for hiring manager review."])
    ws2["A2"].font = Font(name="Calibri", size=10, color="FF5E6675")
    ws2.append([])

    ev_cols = ["Candidate", "Overall Score", "Status", "Evidence Label", "Evidence from CV", "Strengths", "Gaps"]
    ws2.append(ev_cols)
    ev_hdr = ws2.max_row
    for col_idx in range(1, len(ev_cols)+1):
        cell = ws2.cell(row=ev_hdr, column=col_idx)
        cell.font = hdr_font; cell.fill = hdr_fill
        cell.alignment = hdr_align; cell.border = border
    ws2.row_dimensions[ev_hdr].height = 22

    for r in eligible:
        sc   = r.get("overall", 0)
        sl   = r.get("shortlisted")
        name = r.get("name", r.get("filename",""))
        status = "Shortlisted" if sl is True else "Rejected" if sl is False else "Pending"
        evidence  = r.get("evidence", [])
        strengths = "; ".join(r.get("strengths", []))
        gaps      = "; ".join(r.get("gaps", []))
        items = evidence if evidence else [{"label": "Summary", "text": r.get("summary","")}]
        for i, ev in enumerate(items):
            ws2.append([
                name      if i == 0 else "",
                sc        if i == 0 else "",
                status    if i == 0 else "",
                ev.get("label",""),
                ev.get("text",""),
                strengths if i == 0 else "",
                gaps      if i == 0 else "",
            ])
            row_idx = ws2.max_row
            for col_idx in range(1, len(ev_cols)+1):
                cell = ws2.cell(row=row_idx, column=col_idx)
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
                cell.border = border

    for col, w in zip(["A","B","C","D","E","F","G"], [22, 8, 12, 22, 60, 35, 28]):
        ws2.column_dimensions[col].width = w
    ws2.freeze_panes = f"A{ev_hdr+1}"

    # ── Sheet 3: Status History ─────────────────────────────────────────────────
    if history:
        wh = wb.create_sheet("Status History")
        wh.append(["Candidate", "Action", "Timestamp"])
        for col_idx in range(1, 4):
            cell = wh.cell(row=1, column=col_idx)
            cell.font = hdr_font; cell.fill = hdr_fill; cell.alignment = hdr_align
        for e in sorted([e for v in history.values() for e in v], key=lambda e: e["ts"], reverse=True):
            wh.append([e.get("name",""), e.get("action","").capitalize(), e.get("ts","")])
            row_idx = wh.max_row
            for col_idx in range(1, 4):
                wh.cell(row=row_idx, column=col_idx).alignment = Alignment(horizontal="left", vertical="center")
                wh.cell(row=row_idx, column=col_idx).border = border
        for col, w in zip(["A","B","C"], [24, 16, 22]):
            wh.column_dimensions[col].width = w
        wh.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ── API routes ─────────────────────────────────────────────────────────────────

@app.post("/api/parse-jd")
async def parse_jd(file: UploadFile = File(...), qs_token: Optional[str] = Cookie(default=None)):
    user = _auth.get_current_user(qs_token)
    try:
        b = await file.read()
        text = parse_bytes(b, file.filename)
        return {"text": text}
    except Exception as e:
        error_plain, fix_suggestion = _plain_error(file.filename, e)
        ext = file.filename.rsplit(".", 1)[-1].lower() if "." in file.filename else "unknown"
        try:
            _auth.log_upload_error(file.filename, ext, "JD upload", user["username"],
                                   type(e).__name__, error_plain, fix_suggestion)
        except Exception:
            pass
        raise HTTPException(status_code=400, detail=error_plain)


@app.post("/api/detect-competencies")
async def detect_competencies(jd: str = Form(...), qs_token: Optional[str] = Cookie(default=None)):
    _auth.get_current_user(qs_token)
    client = anthropic.Anthropic(api_key=get_api_key())
    msg = client.messages.create(
        model="claude-opus-4-8", max_tokens=300,
        messages=[{"role": "user", "content":
            f"Extract exactly 5 key skill competencies from this job description. "
            f"Return ONLY a JSON array of 5 short strings (3-5 words each), nothing else.\n\n{jd}"}]
    )
    raw = msg.content[0].text.strip()
    raw = re.sub(r"^```[a-z]*\n?", "", raw).rstrip("` \n")
    return {"competencies": json.loads(raw)}


@app.post("/api/screen")
async def screen(
    jd: str = Form(...),
    competencies: str = Form(""),
    calibration: str = Form("[]"),
    role_title: str = Form(""),
    files: List[UploadFile] = File(...),
    qs_token: Optional[str] = Cookie(default=None),
):
    user = _auth.get_current_user(qs_token)
    cvs = {}
    for f in files:
        try:
            cvs[f.filename] = parse_bytes(await f.read(), f.filename)
        except Exception as e:
            error_plain, fix_suggestion = _plain_error(f.filename, e)
            ext = f.filename.rsplit(".", 1)[-1].lower() if "." in f.filename else "unknown"
            try:
                _auth.log_upload_error(f.filename, ext, role_title or "unknown",
                                       user["username"], type(e).__name__, error_plain, fix_suggestion)
            except Exception:
                pass
            raise HTTPException(status_code=400, detail=error_plain)

    calib = json.loads(calibration)
    past_examples = _auth.get_screening_examples(role_title) if role_title else []

    client = anthropic.Anthropic(api_key=get_api_key())
    msg = client.messages.create(
        model="claude-opus-4-8", max_tokens=16000,
        messages=[{"role": "user", "content": build_prompt(jd, competencies, cvs, calib, past_examples)}]
    )
    truncated = msg.stop_reason == "max_tokens"
    raw = msg.content[0].text if msg.content else ""
    if not raw:
        raise HTTPException(status_code=502, detail="Claude returned an empty response.")
    try:
        results = extract_json(raw)
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"Could not parse Claude response: {e}\n\n{raw[:500]}")
    return {"results": results, "truncated": truncated, "past_examples_used": len(past_examples)}


# ── Feedback / learning routes ─────────────────────────────────────────────────

class FeedbackExample(BaseModel):
    candidate_name: str
    ai_score: int
    final_decision: str   # "shortlisted" or "rejected"
    summary: str = ""
    strengths: str = ""
    gaps: str = ""
    recruiter_note: str = ""

class FeedbackBody(BaseModel):
    role_title: str
    examples: List[FeedbackExample]

@app.post("/api/feedback")
async def save_feedback(body: FeedbackBody, qs_token: Optional[str] = Cookie(default=None)):
    user = _auth.get_current_user(qs_token)
    if not body.role_title.strip():
        raise HTTPException(status_code=400, detail="role_title is required to save learning examples")
    _auth.save_screening_examples(
        body.role_title,
        [ex.model_dump() for ex in body.examples],
        user["username"],
    )
    return {"saved": len(body.examples), "role_title": body.role_title}

@app.get("/api/feedback/roles")
async def feedback_roles(qs_token: Optional[str] = Cookie(default=None)):
    _auth.get_current_user(qs_token)
    return _auth.list_example_roles()

@app.get("/api/feedback/{role_title}")
async def feedback_for_role(role_title: str, qs_token: Optional[str] = Cookie(default=None)):
    _auth.get_current_user(qs_token)
    return _auth.get_screening_examples(role_title)


class ExportBody(BaseModel):
    results: list
    role_title: str = ""
    history: dict = {}
    score_feedback: dict = {}


@app.post("/api/export/pdf")
async def export_pdf(body: ExportBody, qs_token: Optional[str] = Cookie(default=None)):
    _auth.get_current_user(qs_token)
    pdf = generate_pdf(body.results, body.role_title)
    filename = f"CV_Screening_{datetime.now().strftime('%Y-%m-%d')}.pdf"
    return Response(
        pdf,
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.post("/api/export/excel")
async def export_excel(body: ExportBody, qs_token: Optional[str] = Cookie(default=None)):
    _auth.get_current_user(qs_token)
    xlsx = generate_excel(body.results, body.role_title, body.history, body.score_feedback)
    filename = f"CV_Screening_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    return Response(
        xlsx,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# ── Serve frontend ─────────────────────────────────────────────────────────────
app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/")
async def root():
    return FileResponse("static/index.html")
