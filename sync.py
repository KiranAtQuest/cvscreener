"""
Zoho → Google Drive CV sync.
Runs hourly as a Render Cron Job.
Fetches candidates created within ZOHO_SYNC_DAYS, downloads their CV
attachments, and uploads to Google Drive with position title embedded
in filename: {candidate_id}_{position_title}_{candidate_name}_{file}
Also uploads a CSV manifest of every file synced in the run.
"""
import csv, io, json, mimetypes, os, re, time
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo

import requests
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

ZOHO_ACCOUNTS_URL  = os.environ.get("ZOHO_ACCOUNTS_URL",  "https://accounts.zoho.in")
ZOHO_RECRUIT_URL   = os.environ.get("ZOHO_RECRUIT_URL",   "https://recruit.zoho.in")
ZOHO_CLIENT_ID     = os.environ.get("ZOHO_CLIENT_ID", "")
ZOHO_CLIENT_SECRET = os.environ.get("ZOHO_CLIENT_SECRET", "")
ZOHO_REFRESH_TOKEN = os.environ.get("ZOHO_REFRESH_TOKEN", "")
GOOGLE_DRIVE_FOLDER_ID = os.environ.get("GOOGLE_DRIVE_FOLDER_ID", "")

ALLOWED_EXTENSIONS = {".pdf", ".doc", ".docx", ".rtf", ".odt", ".txt"}
SYNC_DAYS = int(os.environ.get("ZOHO_SYNC_DAYS", "2"))
TIMEOUT = 60
IST = ZoneInfo("Asia/Kolkata")


# ── Auth ───────────────────────────────────────────────────────────────────────

def get_zoho_token():
    resp = requests.post(f"{ZOHO_ACCOUNTS_URL}/oauth/v2/token", data={
        "grant_type":    "refresh_token",
        "client_id":     ZOHO_CLIENT_ID,
        "client_secret": ZOHO_CLIENT_SECRET,
        "refresh_token": ZOHO_REFRESH_TOKEN,
    }, timeout=30)
    resp.raise_for_status()
    return resp.json()["access_token"]

def zoho_headers(token):
    return {"Authorization": f"Zoho-oauthtoken {token}", "Accept": "application/json"}

def get_drive_service():
    sa_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "")
    info = json.loads(sa_json)
    creds = service_account.Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/drive"]
    )
    return build("drive", "v3", credentials=creds, cache_discovery=False)


# ── Zoho helpers ───────────────────────────────────────────────────────────────

def get_candidates_from_date(token, cutoff):
    """Fetch only candidates created on/after cutoff using the search endpoint."""
    url = f"{ZOHO_RECRUIT_URL}/recruit/v2/Candidates/search"
    cutoff_str = cutoff.isoformat(timespec="seconds")
    criteria = f"(Created_Time:greater_equal:{cutoff_str})"
    page = 1
    while True:
        print(f"  Fetching candidate page {page} (Created_Time >= {cutoff_str})…")
        resp = requests.get(url, headers=zoho_headers(token), params={
            "criteria": criteria, "page": page, "per_page": 200,
            "converted": "both", "approved": "both",
        }, timeout=TIMEOUT)
        if resp.status_code == 204:
            break
        resp.raise_for_status()
        data = resp.json()
        candidates = data.get("data", [])
        if not candidates:
            break
        yield from candidates
        if not data.get("info", {}).get("more_records"):
            break
        page += 1

def get_candidate_position(token, candidate_id):
    """Return (position_title, position_id) via the /associate sub-resource."""
    resp = requests.get(
        f"{ZOHO_RECRUIT_URL}/recruit/v2/Candidates/{candidate_id}/associate",
        headers=zoho_headers(token),
        params={"page": 1, "per_page": 200},
        timeout=TIMEOUT,
    )
    if resp.status_code == 204:
        return "Unassigned", ""
    resp.raise_for_status()
    positions = resp.json().get("data", [])
    if not positions:
        return "Unassigned", ""
    pos = positions[0]
    title = (pos.get("Posting_Title") or pos.get("Potential_Name")
             or pos.get("Job_Opening_Name") or "Unassigned")
    pid   = pos.get("id") or pos.get("Job_Opening_ID") or ""
    return str(title).strip(), str(pid).strip()

def parse_zoho_dt(value):
    if not value:
        return None
    try:
        dt = datetime.fromisoformat(str(value).strip())
    except ValueError:
        try:
            dt = datetime.strptime(str(value).strip(), "%Y-%m-%d %H:%M:%S")
        except ValueError:
            return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=IST)
    return dt

def clean_name(value):
    value = str(value or "").strip()
    for ch in '<>:"/\\|?*':
        value = value.replace(ch, "_")
    return " ".join(value.split())[:180].strip()


# ── Drive helpers ──────────────────────────────────────────────────────────────

def file_exists_in_drive(drive, name):
    safe = name.replace("'", "\\'")
    res = drive.files().list(
        q=f"name='{safe}' and '{GOOGLE_DRIVE_FOLDER_ID}' in parents and trashed=false",
        fields="files(id)", pageSize=1,
        supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()
    return bool(res.get("files"))

def upload_to_drive(drive, content, name, mime=None):
    mime = mime or mimetypes.guess_type(name)[0] or "application/octet-stream"
    media = MediaIoBaseUpload(io.BytesIO(content), mimetype=mime, resumable=False)
    return drive.files().create(
        body={"name": name, "parents": [GOOGLE_DRIVE_FOLDER_ID]},
        media_body=media, fields="id,name,webViewLink",
        supportsAllDrives=True,
    ).execute()

def upload_manifest(drive, records):
    if not records:
        return
    out = io.StringIO()
    writer = csv.DictWriter(out, fieldnames=[
        "Candidate_ID", "Candidate_Name", "Position_Title", "Position_ID",
        "CV_File_Name", "Attachment_Created_Time", "Drive_File_ID", "Drive_URL",
    ])
    writer.writeheader()
    writer.writerows(records)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    f = upload_to_drive(drive, out.getvalue().encode("utf-8-sig"),
                        f"CV_Download_Manifest_{ts}.csv", mime="text/csv")
    print(f"\n  Manifest uploaded: {f['name']}")
    print(f"  {f.get('webViewLink', '')}")


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    print(f"\n{'='*50}")
    print(f"ZOHO → DRIVE SYNC  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*50}\n")

    cutoff = datetime.now(IST) - timedelta(days=SYNC_DAYS)
    print(f"Processing candidates created on/after: {cutoff.date()}\n")

    token = get_zoho_token()
    drive = get_drive_service()
    manifest = []
    uploaded = skipped_old = skipped_dup = errors = 0

    for cand in get_candidates_from_date(token, cutoff):
        cid  = str(cand["id"])
        name = str(cand.get("Full_Name") or cand.get("Last_Name") or "Candidate").strip()

        # Refresh token per candidate to avoid expiry on large runs
        token = get_zoho_token()

        # Get position via /associate
        try:
            position_title, position_id = get_candidate_position(token, cid)
        except Exception as e:
            print(f"  WARNING: could not get position for {name}: {e}")
            position_title, position_id = "Unassigned", ""

        print(f"\nCandidate: {name}  (ID: {cid})  Position: {position_title}")

        # Fetch attachments
        try:
            resp = requests.get(
                f"{ZOHO_RECRUIT_URL}/recruit/v2/Candidates/{cid}/Attachments",
                headers=zoho_headers(token), timeout=TIMEOUT,
            )
            if resp.status_code == 204:
                continue
            resp.raise_for_status()
            attachments = resp.json().get("data", [])
        except Exception as e:
            print(f"  ERROR getting attachments: {e}")
            errors += 1
            continue

        for att in attachments:
            orig_name = att.get("File_Name") or f"{att['id']}.bin"
            if os.path.splitext(orig_name.lower())[1] not in ALLOWED_EXTENSIONS:
                continue

            created = parse_zoho_dt(att.get("Created_Time"))
            if not created or created < cutoff:
                skipped_old += 1
                continue

            # Filename: {id}_{position_title}_{candidate_name}_{original_file}
            drive_name = clean_name(
                f"{cid}_{position_title}_{name}_{orig_name}"
            )

            if file_exists_in_drive(drive, drive_name):
                skipped_dup += 1
                continue

            try:
                dl = requests.get(
                    f"{ZOHO_RECRUIT_URL}/recruit/v2/Candidates/{cid}/Attachments/{att['id']}",
                    headers=zoho_headers(token), timeout=120,
                )
                dl.raise_for_status()
                f = upload_to_drive(drive, dl.content, drive_name)
                uploaded += 1
                print(f"  ✓ {drive_name}")
                manifest.append({
                    "Candidate_ID":            cid,
                    "Candidate_Name":          name,
                    "Position_Title":          position_title,
                    "Position_ID":             position_id,
                    "CV_File_Name":            drive_name,
                    "Attachment_Created_Time": created.isoformat(),
                    "Drive_File_ID":           f["id"],
                    "Drive_URL":               f.get("webViewLink", ""),
                })
                time.sleep(0.5)
            except Exception as e:
                print(f"  ERROR uploading {orig_name}: {e}")
                errors += 1

    upload_manifest(drive, manifest)

    print(f"\n{'='*50}")
    print(f"Done.")
    print(f"  Uploaded:      {uploaded}")
    print(f"  Skipped old:   {skipped_old}")
    print(f"  Skipped dup:   {skipped_dup}")
    print(f"  Errors:        {errors}")
    print(f"{'='*50}\n")


if __name__ == "__main__":
    main()
