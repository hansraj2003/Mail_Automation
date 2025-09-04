# safe_bulk_mailer.py
"""
Safe Bulk Mailer – per spec

Key features:
- Reads baseline Excel: sheet "All Previous Year Data"; 1st (unnamed) col is CompanyName
- Uses Email ID_HR1/2/3 + HR1_Name/HR2_Name/HR3_Name
- Creates/maintains HR1_iterated/HR2_iterated/HR3_iterated with values: True, False, NA
  * NA  -> email cell blank or '0'
  * True-> successfully sent
  * False-> attempted but error (retryable)
- Runs exactly 3 sessions back-to-back in one run: 2h, 1h, 2h
- In each session: picks random 15–25 eligible HR contacts (HR slots) and sends
- Delay between mails: random 2–6 minutes + random seconds (0–59)
- Randomly uses Drive link OR attaches resume PDF
  * Attachment filename: Hansraj_Joshi_DTU_{company_name_with_underscores}.pdf
  * Logs which mode used per mail
- Output log Excel (append-only): company, HR name, email, timestamp, mode, status, error message
- Console prints: total planned mails at session start; 'Sent to - {name} {email}; next mail in - Xm Ys'; errors too
- Sends summary email after each session (to SUMMARY_EMAIL)
  * Contains session number, sent count, last mail's company/HR/email
- Sends completion email when all 1066 rows' HR slots are iterated (i.e., True or NA)
- All credentials/paths live in .env

Run: python safe_bulk_mailer.py
"""
import os
import time
import random
from datetime import datetime, timedelta
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import smtplib
from email.utils import formataddr
from dotenv import load_dotenv

# -------------------------
# Load configuration from .env
# -------------------------
load_dotenv()
BASELINE_XLSX_PATH = os.getenv("BASELINE_XLSX_PATH", "Copy of JPC.xlsx")
SHEET_NAME = os.getenv("SHEET_NAME", "All Previous Year Data")
OUTBOX_LOG_PATH = os.getenv("OUTBOX_LOG_PATH", "outbox_log.xlsx")
SUMMARY_EMAIL = os.getenv("SUMMARY_EMAIL", "anothermail@gmail.com")

SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD")
SMTP_HOST = os.getenv("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))

DRIVE_RESUME_LINK = os.getenv("DRIVE_RESUME_LINK", "https://drive.google.com/yourresume")
RESUME_PDF_PATH = os.getenv("RESUME_PDF_PATH", "./Resume.pdf")

# Subject/body templates (customize later if needed)
SUBJECT_TEMPLATE = os.getenv("SUBJECT_TEMPLATE", "Internship application – {company}")
BODY_TEMPLATE_LINK = os.getenv(
    "BODY_TEMPLATE_LINK",
    (
        "{greeting}\n\n"
        "I’m a final-year student at DTU. I’d love to be considered for an internship with {company}.\n"
        "Here’s my resume: {resume_link}\n\n"
        "Thank you for your time!\n"
        "Hansraj Joshi"
    ),
)
BODY_TEMPLATE_ATTACH = os.getenv(
    "BODY_TEMPLATE_ATTACH",
    (
        "{greeting}\n\n"
        "I’m a final-year student at DTU. I’d love to be considered for an internship with {company}.\n"
        "I’ve attached my resume for your review.\n\n"
        "Thank you for your time!\n"
        "Hansraj Joshi"
    ),
)

# Sessions configuration
SESSION_DURATIONS_MIN = [120, 60, 120]  # minutes: [2h, 1h, 2h]
PER_SESSION_MIN = int(os.getenv("PER_SESSION_MIN", 2))
PER_SESSION_MAX = int(os.getenv("PER_SESSION_MAX", 3))

# Delay configuration (seconds)
MIN_DELAY = int(os.getenv("MIN_DELAY_SECONDS", 1 * 60))
MAX_DELAY = int(os.getenv("MAX_DELAY_SECONDS", 2 * 60))

TOTAL_ROWS_HINT = 1066  # informational only

# -------------------------
# Helpers
# -------------------------
def safe_str(x):
    if pd.isna(x):
        return ""
    return str(x).strip()

def is_blank_or_zero(x: str) -> bool:
    s = safe_str(x)
    return (s == "") or (s == "0")

def company_to_filename(company_name: str) -> str:
    cname = safe_str(company_name)
    if not cname:
        cname = "Company"
    cname = "_".join(cname.split())  # spaces -> underscore
    # keep only alnum + _ to avoid weird filenames
    clean = "".join(ch for ch in cname if ch.isalnum() or ch == "_")
    return clean or "Company"

def ensure_iter_cols(df: pd.DataFrame) -> None:
    for i in [1, 2, 3]:
        col = f"HR{i}_iterated"
        if col not in df.columns:
            df[col] = ""

def send_email(
    sender: str,
    password: str,
    recipient: str,
    subject: str,
    body_text: str,
    attach_bytes: bytes | None = None,
    attach_filename: str | None = None,
) -> tuple[bool, str | None]:
    msg = MIMEMultipart()
    msg["From"] = formataddr(("Application from Hansraj Joshi | DTU | FTE/Internship", sender))
    msg["To"] = recipient
    msg["Subject"] = subject
    msg.attach(MIMEText(body_text, "plain"))

    if attach_bytes is not None and attach_filename:
        part = MIMEApplication(attach_bytes, _subtype="pdf")
        part.add_header("Content-Disposition", "attachment", filename=attach_filename)
        msg.attach(part)

    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
            server.starttls()
            server.login(sender, password)
            server.sendmail(sender, recipient, msg.as_string())
        return True, None
    except Exception as e:
        return False, str(e)

# Append a single row to the outbox log (creates or appends)
def append_log(log_row: dict):
    try:
        if os.path.exists(OUTBOX_LOG_PATH):
            existing = pd.read_excel(OUTBOX_LOG_PATH)
            combined = pd.concat([existing, pd.DataFrame([log_row])], ignore_index=True)
        else:
            combined = pd.DataFrame([log_row])
        with pd.ExcelWriter(OUTBOX_LOG_PATH, engine="openpyxl", mode="w") as writer:
            combined.to_excel(writer, index=False)
    except Exception as e:
        print(f"[WARN] Failed to write log: {e}")

# Check if all HR slots are iterated (True or NA)
def all_iterated(df: pd.DataFrame) -> bool:
    for _, row in df.iterrows():
        for i in [1, 2, 3]:
            val = safe_str(row.get(f"HR{i}_iterated", ""))
            if val not in ("True", "NA"):
                return False
    return True

# Build a flat list of eligible (row_idx, hr_no) to process
# Eligible means: email present & not '0' and iterated not in {True, NA}
def build_eligible(df: pd.DataFrame) -> list[tuple[int, int]]:
    eligible: list[tuple[int, int]] = []
    for idx, row in df.iterrows():
        for i in [1, 2, 3]:
            email_col = f"Email ID_HR{i}"
            iter_col = f"HR{i}_iterated"
            if email_col not in df.columns:
                continue
            email_val = safe_str(row.get(email_col, ""))
            iter_val = safe_str(row.get(iter_col, ""))
            if is_blank_or_zero(email_val):
                # proactively mark NA (no email available)
                if iter_val == "":
                    df.at[idx, iter_col] = "NA"
                continue
            if iter_val in ("True", "NA"):
                continue
            # iter_val is "" or "False" or an error string → eligible for (re)try
            eligible.append((idx, i))
    return eligible

# Send summary after each session
def send_session_summary(session_no: int, sent_count: int, last_info: dict | None):
    subject = f"Session {session_no} completed"
    if last_info:
        body = (
            f"Session {session_no} of the day is completed.\n"
            f"Total mails sent: {sent_count}\n\n"
            f"Last mail sent to:\n"
            f"  Company: {last_info.get('CompanyName')}\n"
            f"  HR: {last_info.get('HR_Name')}\n"
            f"  Email: {last_info.get('HR_Email')}\n"
        )
    else:
        body = (
            f"Session {session_no} of the day is completed.\n"
            f"No mails were sent in this session."
        )
    ok, err = send_email(SENDER_EMAIL, SENDER_PASSWORD, SUMMARY_EMAIL, subject, body)
    if not ok:
        print(f"[WARN] Failed to send session summary: {err}")

# Main per-session runner

def run_session(session_no: int, df: pd.DataFrame, session_minutes: int) -> None:
    session_end = datetime.now() + timedelta(minutes=session_minutes)

    # Build pool of eligible HR slots
    eligible = build_eligible(df)
    random.shuffle(eligible)

    # Choose how many to attempt this session
    target_count = random.randint(PER_SESSION_MIN, PER_SESSION_MAX)
    target_count = min(target_count, len(eligible))

    print(f"\n=== Session {session_no} start: planning to send {target_count} mails (Eligible pool: {len(eligible)}) ===")

    sent_this_session = 0
    last_sent_info: dict | None = None

    for n in range(target_count):
        if datetime.now() >= session_end:
            print(f"[INFO] Session {session_no}: time window reached. Stopping early.")
            break
        if not eligible:
            print("[INFO] No more eligible contacts to send in this session.")
            break

        row_idx, hr_no = eligible.pop(0)
        row = df.loc[row_idx]

        company = safe_str(row.iloc[0])  # first column is company name
        hr_name = safe_str(row.get(f"HR{hr_no}_Name", ""))
        hr_email = safe_str(row.get(f"Email ID_HR{hr_no}", ""))
        iter_col = f"HR{hr_no}_iterated"

        # Double-check blanks → mark NA and continue
        if is_blank_or_zero(hr_email):
            df.at[row_idx, iter_col] = "NA"
            append_log({
                "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "CompanyName": company,
                "HR_Slot": f"HR{hr_no}",
                "HR_Name": hr_name or "NA",
                "HR_Email": hr_email,
                "Status": "NA",
                "Mode": None,
                "Detail": "Email blank or 0",
            })
            continue

        # Prepare greeting/subject/body
        greeting = f"Dear {hr_name}," if hr_name else "Dear Hiring Manager,"
        subject = SUBJECT_TEMPLATE.format(company=company, hr_name=hr_name)

        # Randomly select mode: link or attachment
        use_link = bool(random.getrandbits(1))
        attach_bytes = None
        attach_name = None
        if use_link:
            body_text = BODY_TEMPLATE_LINK.format(greeting=greeting, company=company, resume_link=DRIVE_RESUME_LINK)
            mode = "LINK"
            detail = DRIVE_RESUME_LINK
        else:
            try:
                with open(RESUME_PDF_PATH, "rb") as f:
                    attach_bytes = f.read()
                sanitized = company_to_filename(company)
                attach_name = f"Hansraj_Joshi_DTU_{sanitized}.pdf"
                body_text = BODY_TEMPLATE_ATTACH.format(greeting=greeting, company=company)
                mode = "ATTACHMENT"
                detail = attach_name
            except Exception as e:
                # Fallback to link if local file missing
                body_text = BODY_TEMPLATE_LINK.format(greeting=greeting, company=company, resume_link=DRIVE_RESUME_LINK)
                mode = "LINK_FALLBACK"
                detail = f"Missing file -> {DRIVE_RESUME_LINK} ({e})"
                attach_bytes = None
                attach_name = None

        # Send email
        ok, err = send_email(
            SENDER_EMAIL,
            SENDER_PASSWORD,
            hr_email,
            subject,
            body_text,
            attach_bytes=attach_bytes,
            attach_filename=attach_name,
        )

        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        if ok:
            df.at[row_idx, iter_col] = "True"
            status = "SENT"
            last_sent_info = {"CompanyName": company, "HR_Name": hr_name or "Hiring Manager", "HR_Email": hr_email}
            print(f"Sent to - {hr_name or 'Hiring Manager'} {hr_email}")
        else:
            df.at[row_idx, iter_col] = "False"
            status = "ERROR"
            print(f"[ERROR] {hr_email} -> {err}")

        append_log({
            "Timestamp": ts,
            "CompanyName": company,
            "HR_Slot": f"HR{hr_no}",
            "HR_Name": hr_name or "NA",
            "HR_Email": hr_email,
            "Status": status,
            "Mode": mode,
            "Detail": detail if ok else f"{detail} | ERROR: {err}",
        })

        # Persist baseline after each attempt (crash-safe)
        try:
            with pd.ExcelWriter(BASELINE_XLSX_PATH, engine="openpyxl", mode="w") as writer:
                df.to_excel(writer, sheet_name=SHEET_NAME, index=False)
        except Exception as e:
            print(f"[WARN] Failed to save baseline immediately: {e}")

        # Sleep between mails
        delay = random.randint(MIN_DELAY, MAX_DELAY) + random.randint(0, 59)
        mins, secs = divmod(delay, 60)
        # Avoid sleeping past session end excessively
        if datetime.now() + timedelta(seconds=delay) > session_end:
            # shorten to fit session window but still look natural
            remaining = max(0, int((session_end - datetime.now()).total_seconds()) - random.randint(5, 20))
            mins, secs = divmod(remaining, 60)
            delay = remaining
        if delay > 0:
            print(f"Next mail in - {mins}m {secs}s")
            time.sleep(delay)

        sent_this_session += 1

    # Session wrap-up
    send_session_summary(session_no, sent_this_session, last_sent_info)

# -------------------------
# Main runner – three sessions back-to-back
# -------------------------
if __name__ == "__main__":
    if not SENDER_EMAIL or not SENDER_PASSWORD:
        raise RuntimeError("SENDER_EMAIL/SENDER_PASSWORD missing in .env")

    if not os.path.exists(BASELINE_XLSX_PATH):
        raise FileNotFoundError(f"Baseline Excel not found: {BASELINE_XLSX_PATH}")

    # Load baseline
    df = pd.read_excel(BASELINE_XLSX_PATH, sheet_name=SHEET_NAME, dtype=object)

    # Rename first/unnamed col to CompanyName (keeps position)
    first_col = df.columns[0]
    if first_col != "CompanyName":
        df = df.rename(columns={first_col: "CompanyName"})

    # Normalize column names
    df.columns = [str(c).strip() for c in df.columns]

    # Ensure iter columns exist
    ensure_iter_cols(df)

    # Persist any initial NA markings (for rows with blank/0 emails)
    _ = build_eligible(df)
    with pd.ExcelWriter(BASELINE_XLSX_PATH, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name=SHEET_NAME, index=False)

    print(f"Loaded baseline: {len(df)} rows (hint total {TOTAL_ROWS_HINT}). Starting 3 sessions...")

    for session_no, minutes in enumerate(SESSION_DURATIONS_MIN, start=1):
        run_session(session_no, df, minutes)
        # Brief pause between sessions to avoid immediate transitions (looks natural)
        if session_no < 3:
            gap = random.randint(60, 180)  # 1–3 minutes gap
            print(f"\n[INFO] Break before next session: {gap}s\n")
            time.sleep(gap)

    # Final check – if all iterated, send completion email
    if all_iterated(df):
        ok, err = send_email(
            SENDER_EMAIL,
            SENDER_PASSWORD,
            SUMMARY_EMAIL,
            "All rows iterated – completed",
            "All HR slots across the sheet have been iterated (True/NA).",
        )
        if ok:
            print("[INFO] Completion email sent.")
        else:
            print(f"[WARN] Failed to send completion email: {err}")

# -------------------------
# .env EXAMPLE (place in same folder)
# -------------------------
# BASELINE_XLSX_PATH=Copy of JPC.xlsx
# SHEET_NAME=All Previous Year Data
# OUTBOX_LOG_PATH=outbox_log.xlsx
# SUMMARY_EMAIL=anothermail@gmail.com
# SENDER_EMAIL=your_college_email@domain.edu
# SENDER_PASSWORD=your_app_password_here
# SMTP_HOST=smtp.gmail.com
# SMTP_PORT=587
# DRIVE_RESUME_LINK=https://drive.google.com/yourresume
# RESUME_PDF_PATH=./Resume.pdf
# SUBJECT_TEMPLATE=Internship application – {company}
# BODY_TEMPLATE_LINK={greeting}\n\nI’m a final-year student at DTU... link: {resume_link}
# BODY_TEMPLATE_ATTACH={greeting}\n\nI’m a final-year student at DTU... (attached)
# PER_SESSION_MIN=15
# PER_SESSION_MAX=25
# MIN_DELAY_SECONDS=120
# MAX_DELAY_SECONDS=360
