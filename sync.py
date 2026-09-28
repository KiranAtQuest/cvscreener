"""
Zoho → Google Drive CV sync.
Runs hourly as a Render Cron Job.
Fetches candidates from Zoho Recruit, downloads their CV attachments,
and uploads to Google Drive with role embedded in filename.
"""
import io, json, mimetypes, os, re, time
from datetime import datetime, timedelta, timezone

import requests
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

ZOHO_ACCOUNTS_URL = os.environ.get("ZOHO_ACCOUNTS_URL", "https://accounts.zoho.in")
ZOHO_RECRUIT_URL  = os.environ.get("ZOHO_RECRUIT_URL",  "https://recruit.zoho.in")
ZOHO_CLIENT_ID     = os.environ.get("ZOHO_CLIENT_ID", "")
ZOHO_CLIENT_SECRET = os.environ.get("ZOHO_CLIENT_SECRET", "")
ZOHO_REFRESH_TOKEN = os.environ.get("ZOHO_REFRESH_TOKEN", "")
GOOGLE_DRIVE_FOLDER_ID = os.environ.get("GOOGLE_DRIVE_FOLDER_ID", "")

ALLOWED_EXTENSIONS = {".pdf", ".doc", ".docx", ".rtf", ".odt", ".txt"}
SYNC_DAYS = int(os.environ.get("ZOHO_SYNC_DAYS", "2"))  # look back N days
TIMEOUT = 60


# ── Auth ───────────────────────────────────────────────────────────────────────

def get_zoho_token():
    resp = requests.post(f"{ZOHO_ACCOUNTS_URL}/oauth/v2/token", data={
        "grant_type": "refresh_token",
        "client_id": ZOHO_CLIENT_ID,
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

def fetch_all_pages(url, headers, params=None):
    params = dict(params or {})
    params.setdefault("per_page", 200)
    page, results = 1, []
    while True:
        params["page"] = page
        resp = requests.get(url, headers=headers, params=params, timeout=TIMEOUT)
        if resp.status_code == 204:
            break
        resp.raise_for_status()
        data = resp.json().get("data", [])
        results.extend(data)
        if not resp.json().get("info", {}).get("more_records"):
            break
        page += 1
    return results

def build_candidate_role_map(token):
    """
    Build {candidate_id: posting_title} using the Interviews module,
    which reliably links candidates to job openings.
    Also covers candidates not yet interviewed by falling back to
    job opening name from the most recent interview for that job.
    """
    h = zoho_headers(token)
    mapping = {}

    # Primary: Interviews module — has both Candidate_Name.id and Posting_Title.name
    print("Fetching interviews for role mapping…")
    interviews = fetch_all_pages(
        f"{ZOHO_RECRUIT_URL}/recruit/v2/Interviews", h,
        {"fields": "Candidate_Name,Posting_Title"}
    )
    for iv in interviews:
        cand = iv.get("Candidate_Name") or {}
        role = iv.get("Posting_Title") or {}
        cid   = str(cand.get("id", ""))
        title = role.get("name", "") if isinstance(role, dict) else str(role)
        if cid and title:
            mapping[cid] = title
    print(f"  {len(mapping)} candidates mapped via interviews.")

    # Fallback: fetch all job openings and try the Candidates sub-resource
    # (known to return 500 on Zoho's side, so we skip silently)

    return mapping

def parse_zoho_dt(value):
    if not value:
        return None
    try:
        dt = datetime.fromisoformat(value)
    except ValueError:
        try:
            dt = datetime.strptime(value, "%Y-%m-%d %H:%M:%S")
        except ValueError:
            return None
    if dt.tzinfo is None:
        from zoneinfo import ZoneInfo
        dt = dt.replace(tzinfo=ZoneInfo("Asia/Kolkata"))
    return dt

def clean_name(value):
    return re.sub(r'[<>:"/\\|?*]', '_', value).strip()


# ── Drive helpers ──────────────────────────────────────────────────────────────

def file_exists_in_drive(drive, name):
    safe = name.replace("'", "\\'")
    res = drive.files().list(
        q=f"name='{safe}' and '{GOOGLE_DRIVE_FOLDER_ID}' in parents and trashed=false",
        fields="files(id)", pageSize=1,
        supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()
    return bool(res.get("files"))

def upload_to_drive(drive, content, name):
    mime, _ = mimetypes.guess_type(name)
    mime = mime or "application/octet-stream"
    media = MediaIoBaseUpload(io.BytesIO(content), mimetype=mime, resumable=False)
    return drive.files().create(
        body={"name": name, "parents": [GOOGLE_DRIVE_FOLDER_ID]},
        media_body=media, fields="id,name",
        supportsAllDrives=True,
    ).execute()


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    print(f"\n{'='*50}")
    print(f"ZOHO → DRIVE SYNC  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*50}\n")

    cutoff = datetime.now(timezone.utc) - timedelta(days=SYNC_DAYS)
    print(f"Syncing attachments created after: {cutoff.date()}")

    token = get_zoho_token()
    drive = get_drive_service()
    role_map = build_candidate_role_map(token)

    print("\nFetching candidates…")
    candidates = fetch_all_pages(
        f"{ZOHO_RECRUIT_URL}/recruit/v2/Candidates",
        zoho_headers(token),
        {"fields": "id,Full_Name,Last_Name"}
    )
    print(f"  {len(candidates)} candidates found.")

    uploaded = skipped_old = skipped_dup = errors = 0

    for cand in candidates:
        cid   = str(cand["id"])
        name  = str(cand.get("Full_Name") or cand.get("Last_Name") or "Candidate").strip()
        role  = role_map.get(cid, "")

        # Refresh token periodically
        token = get_zoho_token()
        h = zoho_headers(token)

        try:
            resp = requests.get(
                f"{ZOHO_RECRUIT_URL}/recruit/v2/Candidates/{cid}/Attachments",
                headers=h, timeout=TIMEOUT
            )
            if resp.status_code == 204:
                continue
            resp.raise_for_status()
            attachments = resp.json().get("data", [])
        except Exception as e:
            print(f"ERROR getting attachments for {name}: {e}")
            errors += 1
            continue

        for att in attachments:
            orig_name = att.get("File_Name") or f"{att['id']}.bin"
            ext = os.path.splitext(orig_name.lower())[1]
            if ext not in ALLOWED_EXTENSIONS:
                continue

            created = parse_zoho_dt(att.get("Created_Time"))
            if not created or created < cutoff:
                skipped_old += 1
                continue

            role_part = f"{clean_name(role)}_" if role else ""
            drive_name = clean_name(f"{cid}_{name}_{role_part}{orig_name}")

            if file_exists_in_drive(drive, drive_name):
                skipped_dup += 1
                continue

            try:
                dl = requests.get(
                    f"{ZOHO_RECRUIT_URL}/recruit/v2/Candidates/{cid}/Attachments/{att['id']}",
                    headers=h, timeout=120
                )
                dl.raise_for_status()
                upload_to_drive(drive, dl.content, drive_name)
                uploaded += 1
                print(f"  ✓ {drive_name}")
                time.sleep(0.3)
            except Exception as e:
                print(f"  ERROR uploading {orig_name}: {e}")
                errors += 1

    print(f"\n{'='*50}")
    print(f"Done. Uploaded: {uploaded} | Skipped old: {skipped_old} | Skipped dup: {skipped_dup} | Errors: {errors}")
    print(f"{'='*50}\n")


if __name__ == "__main__":
    main()
