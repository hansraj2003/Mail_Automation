"""
Safe Bulk Mailer (per your spec)

Features:
- Reads sheet "All Previous Year Data" from baseline Excel.
- Uses first unnamed column as Company Name.
- Checks Email ID_HR1 / Email ID_HR2 / Email ID_HR3 presence (skip if blank or 0).
- Sends up to 50 HR mails/day (across 3 randomized slots per day).
- Random delay between mails: 2-8 minutes + random seconds.
- Randomly use DRIVE link or PDF attachment per mail.
- Updates baseline file with HR1_iterated/HR2_iterated/HR3_iterated.
- Logs attempts to outbox_log.xlsx.
- Sends summary email to anothermail@gmail.com when daily target reached or when all rows iterated.

Configure variables below before running.
"""

import pandas as pd
import numpy as np
import os
import smtplib
import random
import time
from datetime import datetime, timedelta, time as dtime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# -------------------------
# USER CONFIG (EDIT THESE)
# -------------------------
BASELINE_XLSX_PATH = "Copy of JPC.xlsx"     # path to your baseline Excel
SHEET_NAME = "All Previous Year Data1"

OUTBOX_LOG_PATH = "outbox_log.xlsx"        # log file (will be created if missing)
OUTBOX_LOG_SHEET = "Log"

# Sender (college/domain) - use app password for Gmail
SENDER_EMAIL = "hansrajjoshi_me22b9_70@dtu.ac.in"
SENDER_PASSWORD = "qypx kohx udbi djf"  # use environment variables or .env in production

# Drive link to your resume (single variable in code)
DRIVE_RESUME_LINK = "https://drive.google.com/file/d/17zHOeILK3883WVOJpDmDp4WdNRlDbRtE/view?usp=sharing"  # set your link here

# Local PDF resume path (optional). Script will attach and rename per company when chosen.
RESUME_PDF_PATH = r"C:\Users\Administrator\Desktop\Karan\Hansraj_Joshi_96_pdf.pdf"

# keep your master resume here (script renames in attachment header)

# SMTP settings (default Gmail)
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587

# Mail content placeholders (you will fill BODY later)
SUBJECT = "Internship"            # for now per your request
BODY =  """

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>Internship Inquiry</title>
</head>
<body style="font-family: Arial, sans-serif; background-color: #f6f6f6; margin: 0; padding: 0;">

<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color: #f6f6f6; padding: 20px;">
    <tr>
        <td align="center">
            <table width="600" cellpadding="0" cellspacing="0" border="0" style="max-width: 600px; background-color: #ffffff; border-radius: 8px; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);">
                <tr>
                    <td style="padding: 40px;">
                        <table width="100%" cellpadding="0" cellspacing="0" border="0">
                            <tr>
                                <td style="font-size: 16px; line-height: 1.6; color: #333333;">
                                    <p style="margin-top: 0;">Hello,</p>

                                    <p>My name is Hansraj Joshi, and I am a final-year Mechanical Engineering student at Delhi Technological University (DTU). I am writing to inquire about any potential internship openings at your company.</p>

                                    <p>My studies have provided me with a strong foundation in mechanical design, thermodynamics, and manufacturing processes, and I am eager to apply this knowledge in a professional setting.</p>

                                    <p>I have attached my resume for your consideration and would be grateful for the opportunity to discuss my qualifications further.</p>

                                    <p>Thank you for your time.</p>
                                    <p style="margin-bottom: 0;">Sincerely,</p>
                                    <p style="margin-top: 5px;">Hansraj Joshi<br>
                                    <a href="mailto:your.email@example.com" style="color: #007BFF; text-decoration: none;">your.email@example.com</a></p>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
</table>

</body>
</html>
"""           # fill this block later. Use {hr_name}, {company}, {resume_link} placeholders if desired.

# Daily target
DAILY_TARGET = 50

# Delay bounds between individual mails (in seconds)
MIN_DELAY_SECONDS = 2 * 60       # 2 minutes
MAX_DELAY_SECONDS = 8 * 60       # 8 minutes

# Slot definitions (the script will choose randomized start times in these windows)
# NOTE: You can adjust these windows as required. These are local times used by the running machine.
SLOT_WINDOWS = {
    "morning": {"duration_minutes": 120, "start_range": (6, 12)},   # start between 06:00-09:00
    "afternoon": {"duration_minutes": 60, "start_range": (12, 15)}, # start between 12:00-14:00
    "evening": {"duration_minutes": 120, "start_range": (15, 24)},  # start between 17:00-19:00
}

# Limits per slot (min,max) - used when distributing 50 across slots
SLOT_LIMITS = {
    "morning": (15, 25),
    "afternoon": (10, 20),
    "evening": (15, 25),
}

# Summary recipient (notify when daily target or all done)
SUMMARY_EMAIL = "hansjoshi2003no1@gmail.com"

# -------------------------
# END USER CONFIG
# -------------------------

# Helpers
def safe_str(x):
    return "" if pd.isna(x) else str(x).strip()

def is_blank_or_zero(x):
    s = safe_str(x)
    return (s == "") or (s == "0")

def company_to_filename(company_name):
    # replace spaces with underscores and remove problematic characters
    if pd.isna(company_name):
        cname = "UnknownCompany"
    else:
        cname = str(company_name).strip()
        # replace spaces with underscores, collapse multiple underscores
        cname = "_".join(cname.split())
    # sanitize
    safe = "".join(ch for ch in cname if ch.isalnum() or ch == "_")
    return safe if safe else "Company"

def ensure_iter_cols(df):
    # Ensure iterated columns exist
    for hrn in ["HR1_iterated", "HR2_iterated", "HR3_iterated"]:
        if hrn not in df.columns:
            df[hrn] = False

def append_to_log(log_row):
    # log_row is a dict
    if os.path.exists(OUTBOX_LOG_PATH):
        try:
            existing = pd.read_excel(OUTBOX_LOG_PATH, sheet_name=OUTBOX_LOG_SHEET)
        except Exception:
            existing = pd.DataFrame()
    else:
        existing = pd.DataFrame()
    new_df = existing.append(log_row, ignore_index=True)
    # write back
    with pd.ExcelWriter(OUTBOX_LOG_PATH, engine="openpyxl", mode="w") as writer:
        new_df.to_excel(writer, sheet_name=OUTBOX_LOG_SHEET, index=False)

def send_smtp_mail(sender, password, recipient, subject, body_text, attach_bytes=None, attach_filename=None):
    """
    Sends mail via SMTP. Returns (True, None) on success or (False, error_message) on failure.
    """
    msg = MIMEMultipart()
    msg["From"] = sender
    msg["To"] = recipient
    msg["Subject"] = subject
    msg.attach(MIMEText(body_text, "plain"))

    if attach_bytes is not None and attach_filename:
        part = MIMEApplication(attach_bytes, _subtype="pdf")
        part.add_header("Content-Disposition", "attachment", filename=attach_filename)
        msg.attach(part)

    try:
        server = smtplib.SMTP(SMTP_HOST, SMTP_PORT)
        server.starttls()
        server.login(sender, password)
        server.sendmail(sender, recipient, msg.as_string())
        server.quit()
        return True, None
    except Exception as e:
        return False, str(e)

def choose_slot_distribution(target=DAILY_TARGET):
    """
    Distribute target across the three slots obeying per-slot min/max.
    Returns dict {slot_name: count}
    """
    # available slots
    slots = list(SLOT_WINDOWS.keys())
    min_vals = {s: SLOT_LIMITS[s][0] for s in slots}
    max_vals = {s: SLOT_LIMITS[s][1] for s in slots}
    # Greedy randomized distribution: start with mins then distribute remaining
    remaining = target - sum(min_vals.values())
    distribution = {s: min_vals[s] for s in slots}

    # if remaining < 0, scale down (rare if min sums > target). We'll cap to target evenly.
    if remaining < 0:
        # distribute target proportionally to mins
        total_min = sum(min_vals.values())
        for s in slots:
            distribution[s] = int(round(target * min_vals[s] / total_min))
        # adjust rounding
        diff = target - sum(distribution.values())
        i = 0
        while diff != 0:
            distribution[slots[i % len(slots)]] += 1 if diff > 0 else -1
            diff = target - sum(distribution.values())
            i += 1
        return distribution

    # distribute remaining randomly but respecting max
    slot_order = slots.copy()
    while remaining > 0:
        random.shuffle(slot_order)
        made_progress = False
        for s in slot_order:
            if distribution[s] < max_vals[s]:
                distribution[s] += 1
                remaining -= 1
                made_progress = True
                if remaining <= 0:
                    break
        if not made_progress:
            # cannot distribute more, break
            break
    return distribution

def random_start_time_for_slot(slot_name):
    """
    Choose a randomized start datetime for today within start_range hours.
    Returns a datetime object for the start of that slot (local machine time).
    """
    now = datetime.now()
    start_hour_min, start_hour_max = SLOT_WINDOWS[slot_name]["start_range"]
    # choose a hour within [start_hour_min, start_hour_max)
    chosen_hour = random.randint(start_hour_min, max(start_hour_min, start_hour_max - 1))
    chosen_minute = random.randint(0, 59)
    # today's date with chosen hour/minute
    start_dt = now.replace(hour=chosen_hour, minute=chosen_minute, second=random.randint(0,59), microsecond=0)
    # If start_dt already in past (for same day), use next day's same time
    if start_dt < now:
        start_dt = start_dt + timedelta(days=1)
    return start_dt

# -------------------------
# Core processing function
# -------------------------
def run_day_cycle():
    # Load baseline
    if not os.path.exists(BASELINE_XLSX_PATH):
        print(f"Baseline Excel not found at {BASELINE_XLSX_PATH}")
        return

    # Read sheet (preserve first unnamed column as company)
    df = pd.read_excel(BASELINE_XLSX_PATH, sheet_name=SHEET_NAME, dtype=object)
    # Identify first column name (it may be 'Unnamed: 0' or empty). We'll treat the first column as company.
    first_col = df.columns[0]
    df = df.rename(columns={first_col: "CompanyName"})

    # Normalize column names spaces trimmed
    df.columns = [c.strip() for c in df.columns]

    # Ensure iterated columns exist
    ensure_iter_cols(df)

    total_rows = len(df)
    print(f"Loaded baseline with {total_rows} rows.")

    # Build list of eligible HR entries (flatten HR1/HR2/HR3 per row)
    # We'll iterate rows and for each HR slot that is not iterated and email present and not '0' -> eligible
    eligible_list = []  # will hold tuples (row_index, hr_slot, hr_name_col, hr_email_col)
    for idx, row in df.iterrows():
        # HR1
        for i in [1, 2, 3]:
            iter_col = f"HR{i}_iterated"
            hr_name_col = f"HR{i}_Name"
            hr_email_col = f"Email ID_HR{i}"
            # Some files may have slightly different colnames; if missing, skip that HR slot
            if hr_email_col not in df.columns:
                continue
            already_iter = False
            try:
                already_iter = bool(row.get(iter_col, False))
            except Exception:
                already_iter = False
            if already_iter:
                continue  # skip already done
            email_val = safe_str(row.get(hr_email_col, ""))
            # If the email cell is blank or "0", we will mark NA later and skip sending
            if email_val == "" or email_val == "0":
                # skip adding to eligible (we'll mark NA during a cleanup pass)
                continue
            # otherwise treat as eligible (as per your instruction: if present => it's valid)
            eligible_list.append((idx, i, hr_name_col, hr_email_col))

    # Shuffle eligible list so order is randomized daily
    random.shuffle(eligible_list)

    # If no eligible entries found, we should mark NA for blanks and exit
    if len(eligible_list) == 0:
        print("No eligible HR emails found to send to (all blank or already iterated).")
        # Mark NA for blank ones (optional behavior) - We will mark HR iterated FALSE->NA in baseline for blanks
        # but user specifically wanted NA when email not present; we will mark that now:
        for idx, row in df.iterrows():
            for i in [1,2,3]:
                hr_email_col = f"Email ID_HR{i}"
                iter_col = f"HR{i}_iterated"
                if hr_email_col in df.columns:
                    email_val = safe_str(row.get(hr_email_col, ""))
                    if email_val == "" or email_val == "0":
                        df.at[idx, iter_col] = "NA"
        # save back baseline
        with pd.ExcelWriter(BASELINE_XLSX_PATH, engine="openpyxl", mode="w") as writer:
            df.to_excel(writer, sheet_name=SHEET_NAME, index=False)
        print("Marked NA for blank email slots and updated baseline. Exiting.")
        return

    # Determine today's slot distribution
    distribution = choose_slot_distribution(DAILY_TARGET)
    print("Today's distribution across slots:", distribution)

    # Compute slot schedules (randomized start times)
    slot_schedules = {}
    for slot_name in distribution:
        start_dt = random_start_time_for_slot(slot_name)
        duration = SLOT_WINDOWS[slot_name]["duration_minutes"]
        end_dt = start_dt + timedelta(minutes=duration)
        slot_schedules[slot_name] = {"start": start_dt, "end": end_dt, "count": distribution[slot_name], "sent": 0}
        print(f"Slot '{slot_name}': start={start_dt}, end={end_dt}, count={distribution[slot_name]}")

    # Flat queue of actions: group by slot, but we will pick first N eligible entries and assign them to slots
    total_to_send = min(DAILY_TARGET, len(eligible_list))
    # slice the eligible list
    send_queue = eligible_list[:total_to_send]

    # assign items to slots respecting counts
    assigned = {s: [] for s in slot_schedules}
    current_indices = {s:0 for s in slot_schedules}
    slot_names = list(slot_schedules.keys())
    # Round-robin assign respecting counts
    slot_idx = 0
    for item in send_queue:
        # find next slot with remaining capacity
        attempts = 0
        while attempts < len(slot_names):
            s = slot_names[slot_idx % len(slot_names)]
            if len(assigned[s]) < slot_schedules[s]["count"]:
                assigned[s].append(item)
                slot_idx += 1
                break
            else:
                slot_idx += 1
                attempts += 1
        # if no slot found (shouldn't happen), break
    # debug
    for s in assigned:
        print(f"Assigned {len(assigned[s])} mails to slot {s}.")

    # Start sending per slot. The script expects to be run and wait until slot times come up.
    # For each slot: wait until slot start, then send assigned mails spaced out randomly across the slot.
    total_sent_today = 0

    for slot_name, items in assigned.items():
        if len(items) == 0:
            continue
        sched = slot_schedules[slot_name]
        start_dt = sched["start"]
        end_dt = sched["end"]
        # Wait until slot start (non-blocking sleep loop)
        now = datetime.now()
        # If start time is in the future, sleep until then
        if start_dt > now:
            wait_seconds = (start_dt - now).total_seconds()
            print(f"Waiting for slot '{slot_name}' start in {int(wait_seconds)} seconds (start at {start_dt}).")
            time.sleep(wait_seconds)
        else:
            print(f"Slot '{slot_name}' start time {start_dt} already passed. Beginning immediately.")

        # For fairness, compute average interval to fit all mails in slot if needed, but also mix with randomized 2-8min delays.
        # We'll simply follow randomized delays between each mail, but ensure we stop if we reach end_dt.
        for idx_item, item in enumerate(items):
            row_idx, hr_no, hr_name_col, hr_email_col = item
            # Re-fetch row in case baseline updated
            row = df.loc[row_idx]
            hr_name = safe_str(row.get(hr_name_col, ""))
            hr_email = safe_str(row.get(hr_email_col, ""))
            company_name = safe_str(row.get("CompanyName", ""))

            # If hr_email is blank or 0 (it may have changed), mark NA and continue
            if hr_email == "" or hr_email == "0":
                iter_col = f"HR{hr_no}_iterated"
                df.at[row_idx, iter_col] = "NA"
                # Log
                log_row = {
                    "CompanyName": company_name,
                    "HR_slot": f"HR{hr_no}",
                    "HR_name": hr_name if hr_name else "NA",
                    "HR_email_used": hr_email,
                    "HR_status": "NA",
                    "timestamp": datetime.now(),
                    "slot": slot_name,
                    "sender": SENDER_EMAIL,
                    "delivery_mode": None,
                    "error_message": "Email blank or 0"
                }
                append_to_log(log_row)
                continue

            # Prepare mail
            greeting = f"Dear {hr_name}," if hr_name != "" else "Dear Hiring Manager,"
            # prepare subject and body (placeholders). User to fill BODY later.
            subject = SUBJECT.replace("{company}", company_name).replace("{hr_name}", hr_name)
            # body template might include placeholders {company}, {hr_name}, {resume_link}
            body_text = BODY.format(hr_name=hr_name if hr_name else "Hiring Manager",
                                    company=company_name,
                                    resume_link=DRIVE_RESUME_LINK)

            # Randomly choose delivery mode: LINK or PDF
            delivery_mode = random.choice(["LINK", "PDF"])

            attach_bytes = None
            attach_filename = None

            if delivery_mode == "PDF" and os.path.exists(RESUME_PDF_PATH):
                # read bytes and attach; rename filename using company
                with open(RESUME_PDF_PATH, "rb") as f:
                    attach_bytes = f.read()
                sanitized = company_to_filename(company_name)
                attach_filename = f"Hansraj_Joshi_DTU_{sanitized}.pdf"
            else:
                # if PDF chosen but not available, fallback to LINK
                delivery_mode = "LINK"

            # Send mail
            success, error = send_smtp_mail(
                SENDER_EMAIL, SENDER_PASSWORD,
                hr_email, subject, body_text,
                attach_bytes=attach_bytes, attach_filename=attach_filename
            )

            timestamp_sent = datetime.now()

            # Update baseline & log
            iter_col = f"HR{hr_no}_iterated"
            if success:
                df.at[row_idx, iter_col] = True
                total_sent_today += 1
                slot_schedules[slot_name]["sent"] += 1
                status_str = "SENT"
                error_msg = None
            else:
                # On error, we keep iterated as False (to retry later), and log error
                df.at[row_idx, iter_col] = False
                status_str = "ERROR"
                error_msg = error

            log_row = {
                "CompanyName": company_name,
                "HR_slot": f"HR{hr_no}",
                "HR_name": hr_name if hr_name else "NA",
                "HR_email_used": hr_email,
                "HR_status": status_str,
                "timestamp": timestamp_sent,
                "slot": slot_name,
                "sender": SENDER_EMAIL,
                "delivery_mode": delivery_mode,
                "error_message": error_msg
            }
            append_to_log(log_row)

            # Save baseline progress after each attempt (crash-safety)
            try:
                with pd.ExcelWriter(BASELINE_XLSX_PATH, engine="openpyxl", mode="w") as writer:
                    df.to_excel(writer, sheet_name=SHEET_NAME, index=False)
            except Exception as e:
                print("Warning: failed to save baseline immediately:", e)

            # If daily target reached, send summary and stop
            if total_sent_today >= DAILY_TARGET:
                print("Daily target reached. Sending summary email.")
                send_summary_email(total_sent_today, slot_schedules, OUTBOX_LOG_PATH)
                return

            # Delay before next mail: randomized 2-8 minutes + seconds
            delay = random.randint(MIN_DELAY_SECONDS, MAX_DELAY_SECONDS)
            # also add random seconds (0-59) - though MIN/MAX already include minutes, adding seconds makes sense
            extra_seconds = random.randint(0, 59)
            total_delay = delay + extra_seconds
            # Ensure we don't push beyond slot end by more than some buffer; if close to end, we still proceed but warn
            now = datetime.now()
            if (now + timedelta(seconds=total_delay)) > end_dt:
                # if next send would exceed slot end, we still sleep a short random time (e.g., 10-60 sec) to avoid immediate burst
                small_wait = random.randint(10, 60)
                print(f"Approaching end of slot '{slot_name}', sleeping small wait {small_wait}s before next check.")
                time.sleep(small_wait)
            else:
                print(f"Sleeping {total_delay} seconds before next mail in slot '{slot_name}'.")
                time.sleep(total_delay)

    # End for all slots
    # After finishing all assigned sends for day:
    if total_sent_today > 0:
        print(f"Finished today's sends. Total sent: {total_sent_today}. Sending daily summary.")
        send_summary_email(total_sent_today, slot_schedules, OUTBOX_LOG_PATH)

    # Check if all rows iterated (i.e., each HR slot is True or NA)
    all_iterated = True
    for idx, row in df.iterrows():
        for i in [1,2,3]:
            iter_col = f"HR{i}_iterated"
            if iter_col in df.columns:
                val = row.get(iter_col)
                if val not in [True, "NA", "na", "NaN"]:
                    all_iterated = False
                    break
        if not all_iterated:
            break

    # Save baseline once more
    with pd.ExcelWriter(BASELINE_XLSX_PATH, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name=SHEET_NAME, index=False)

    if all_iterated:
        print("All HR slots iterated across the sheet. Sending completion email.")
        send_completion_email(OUTBOX_LOG_PATH)
    else:
        print("Not all rows iterated yet. Baseline updated.")

# Summary and completion email helpers
def send_summary_email(total_sent, slot_schedules, log_path):
    subject = f"Daily target achieved: {total_sent} mails sent"
    body_lines = [f"Daily summary: {total_sent} mails sent."]
    for s, info in slot_schedules.items():
        body_lines.append(f"- {s}: assigned {info['count']}, sent {info['sent']}")
    body_lines.append(f"Log saved at: {os.path.abspath(log_path)}")
    body = "\n".join(body_lines)
    success, error = send_smtp_mail(SENDER_EMAIL, SENDER_PASSWORD, SUMMARY_EMAIL, subject, body)
    if not success:
        print("Failed to send summary email:", error)

def send_completion_email(log_path):
    subject = "All rows iterated – check log"
    body = f"All HR slots have been iterated. Log saved at: {os.path.abspath(log_path)}"
    success, error = send_smtp_mail(SENDER_EMAIL, SENDER_PASSWORD, SUMMARY_EMAIL, subject, body)
    if not success:
        print("Failed to send completion email:", error)

# -------------------------
# Run
# -------------------------
if __name__ == "__main__":
    print("Starting safe bulk mailer. Make sure you configured variables at top of script.")
    run_day_cycle()
