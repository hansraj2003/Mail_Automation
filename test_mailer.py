# import pandas as pd
# import numpy as np
# from email.utils import formataddr
# import os
# import smtplib
# import random
# import time
# from datetime import datetime, timedelta
# from email.mime.multipart import MIMEMultipart
# from email.mime.text import MIMEText
# from email.mime.application import MIMEApplication

# # -------------------------
# # USER CONFIG (EDIT THESE)
# # -------------------------
# BASELINE_XLSX_PATH = "Copy of JPC.xlsx"     # path to your baseline Excel
# SHEET_NAME = "All Previous Year Data1"

# OUTBOX_LOG_PATH = "outbox_log.xlsx"        # log file (will be created if missing)
# OUTBOX_LOG_SHEET = "Log"

# SENDER_EMAIL = "hansrajjoshi_me22b9_70@dtu.ac.in"
# SENDER_PASSWORD = ""  # use env vars in production

# DRIVE_RESUME_LINK = "https://drive.google.com/file/d/17zHOeILK3883WVOJpDmDp4WdNRlDbRtE/view?usp=sharing"
# RESUME_PDF_PATH = r"C:\Users\Administrator\Desktop\Karan\Hansraj_Joshi_96_pdf.pdf"

# SMTP_HOST = "smtp.gmail.com"
# SMTP_PORT = 587

# SUBJECT = "Internship"
# BODY =  """
# Hello {hr_name},

# My name is Hansraj Joshi, and I am a final-year Mechanical Engineering student at Delhi Technological University (DTU).
# I am writing to inquire about any potential internship openings at {company}.

# I have attached my resume for your consideration (or you can view it here: {resume_link}). 
# I would be grateful for the opportunity to discuss my qualifications further.

# Sincerely,  
# Hansraj Joshi  
# """

# DAILY_TARGET = 50

# MIN_DELAY_SECONDS = 0 * 60
# MAX_DELAY_SECONDS = 1 * 60

# SUMMARY_EMAIL = "hansjoshi2003no1@gmail.com"

# # -------------------------
# # Helpers
# # -------------------------
# def safe_str(x):
#     return "" if pd.isna(x) else str(x).strip()

# def company_to_filename(company_name):
#     if pd.isna(company_name):
#         cname = "UnknownCompany"
#     else:
#         cname = str(company_name).strip()
#         cname = "_".join(cname.split())
#     safe = "".join(ch for ch in cname if ch.isalnum() or ch == "_")
#     return safe if safe else "Company"

# def ensure_iter_cols(df):
#     for hrn in ["HR1_iterated", "HR2_iterated", "HR3_iterated"]:
#         if hrn not in df.columns:
#             df[hrn] = False

# def append_to_log(log_row):
#     if os.path.exists(OUTBOX_LOG_PATH):
#         try:
#             existing = pd.read_excel(OUTBOX_LOG_PATH, sheet_name=OUTBOX_LOG_SHEET)
#         except Exception:
#             existing = pd.DataFrame()
#     else:
#         existing = pd.DataFrame()
#     new_df = pd.concat([existing, pd.DataFrame([log_row])], ignore_index=True)
#     with pd.ExcelWriter(OUTBOX_LOG_PATH, engine="openpyxl", mode="w") as writer:
#         new_df.to_excel(writer, sheet_name=OUTBOX_LOG_SHEET, index=False)

# def send_smtp_mail(sender, password, recipient, subject, body_text, attach_bytes=None, attach_filename=None):
#     msg = MIMEMultipart()
#     msg["From"] = formataddr(("Application from Hansraj Joshi | DTU | FTE/Internship", sender))
#     msg["To"] = recipient
#     msg["Subject"] = subject
#     msg.attach(MIMEText(body_text, "plain"))

#     if attach_bytes is not None and attach_filename:
#         part = MIMEApplication(attach_bytes, _subtype="pdf")
#         part.add_header("Content-Disposition", "attachment", filename=attach_filename)
#         msg.attach(part)

#     try:
#         server = smtplib.SMTP(SMTP_HOST, SMTP_PORT)
#         server.starttls()
#         server.login(sender, password)
#         server.sendmail(sender, recipient, msg.as_string())
#         server.quit()
#         return True, None
#     except Exception as e:
#         return False, str(e)

# # -------------------------
# # Core (no slot scheduling)
# # -------------------------
# def run_day_cycle():
#     if not os.path.exists(BASELINE_XLSX_PATH):
#         print(f"Baseline Excel not found at {BASELINE_XLSX_PATH}")
#         return

#     df = pd.read_excel(BASELINE_XLSX_PATH, sheet_name=SHEET_NAME, dtype=object)
#     first_col = df.columns[0]
#     df = df.rename(columns={first_col: "CompanyName"})
#     df.columns = [c.strip() for c in df.columns]

#     ensure_iter_cols(df)

#     eligible_list = []
#     for idx, row in df.iterrows():
#         for i in [1, 2, 3]:
#             iter_col = f"HR{i}_iterated"
#             hr_name_col = f"HR{i}_Name"
#             hr_email_col = f"Email ID_HR{i}"
#             if hr_email_col not in df.columns:
#                 continue
#             if bool(row.get(iter_col, False)):
#                 continue
#             email_val = safe_str(row.get(hr_email_col, ""))
#             if email_val not in ["", "0"]:
#                 eligible_list.append((idx, i, hr_name_col, hr_email_col))

#     random.shuffle(eligible_list)

#     if len(eligible_list) == 0:
#         print("No eligible HR emails left.")
#         return

#     send_queue = eligible_list[:DAILY_TARGET]
#     total_sent_today = 0

#     for item in send_queue:
#         row_idx, hr_no, hr_name_col, hr_email_col = item
#         row = df.loc[row_idx]
#         hr_name = safe_str(row.get(hr_name_col, ""))
#         hr_email = safe_str(row.get(hr_email_col, ""))
#         company_name = safe_str(row.get("CompanyName", ""))

#         if hr_email in ["", "0"]:
#             df.at[row_idx, f"HR{hr_no}_iterated"] = "NA"
#             continue

#         subject = SUBJECT.replace("{company}", company_name).replace("{hr_name}", hr_name)
#         body_text = BODY.format(hr_name=hr_name or "Hiring Manager",
#                                 company=company_name,
#                                 resume_link=DRIVE_RESUME_LINK)

#         delivery_mode = random.choice(["LINK", "PDF"])
#         attach_bytes = None
#         attach_filename = None
#         if delivery_mode == "PDF" and os.path.exists(RESUME_PDF_PATH):
#             with open(RESUME_PDF_PATH, "rb") as f:
#                 attach_bytes = f.read()
#             sanitized = company_to_filename(company_name)
#             attach_filename = f"Hansraj_Joshi_DTU_{sanitized}.pdf"
#         else:
#             delivery_mode = "LINK"

#         success, error = send_smtp_mail(
#             SENDER_EMAIL, SENDER_PASSWORD,
#             hr_email, subject, body_text,
#             attach_bytes=attach_bytes, attach_filename=attach_filename
#         )

#         timestamp_sent = datetime.now()
#         iter_col = f"HR{hr_no}_iterated"

#         if success:
#             df.at[row_idx, iter_col] = True
#             total_sent_today += 1
#             status_str = "SENT"
#             # 👇 Added console output
#             print(f"✅ Sent to: {hr_name} <{hr_email}>")
#         else:
#             df.at[row_idx, iter_col] = False
#             status_str = "ERROR"
#             print(f"❌ Failed to send to: {hr_name} <{hr_email}> | Error: {error}")


#         log_row = {
#             "CompanyName": company_name,
#             "HR_slot": f"HR{hr_no}",
#             "HR_name": hr_name if hr_name else "NA",
#             "HR_email_used": hr_email,
#             "HR_status": status_str,
#             "timestamp": timestamp_sent,
#             "sender": SENDER_EMAIL,
#             "delivery_mode": delivery_mode,
#             "error_message": error
#         }
#         append_to_log(log_row)

#         with pd.ExcelWriter(BASELINE_XLSX_PATH, engine="openpyxl", mode="w") as writer:
#             df.to_excel(writer, sheet_name=SHEET_NAME, index=False)

#         if total_sent_today >= DAILY_TARGET:
#             print("Daily target reached. Sending summary email.")
#             send_summary_email(total_sent_today, OUTBOX_LOG_PATH)
#             return

#         delay = random.randint(MIN_DELAY_SECONDS, MAX_DELAY_SECONDS) + random.randint(0, 59)
#         mins, secs = divmod(delay, 60)
#         print(f"⏳ Next mail in: {mins}m {secs}s\n")
#         time.sleep(delay)


#     if total_sent_today > 0:
#         print(f"Finished today's sends. Total sent: {total_sent_today}. Sending daily summary.")
#         send_summary_email(total_sent_today, OUTBOX_LOG_PATH)

#     all_iterated = True
#     for idx, row in df.iterrows():
#         for i in [1,2,3]:
#             iter_col = f"HR{i}_iterated"
#             if iter_col in df.columns:
#                 val = row.get(iter_col)
#                 if val not in [True, "NA", "na", "NaN"]:
#                     all_iterated = False
#                     break
#         if not all_iterated:
#             break

#     with pd.ExcelWriter(BASELINE_XLSX_PATH, engine="openpyxl", mode="w") as writer:
#         df.to_excel(writer, sheet_name=SHEET_NAME, index=False)

#     if all_iterated:
#         print("All HR slots iterated. Sending completion email.")
#         send_completion_email(OUTBOX_LOG_PATH)

# def send_summary_email(total_sent, log_path):
#     subject = f"Daily target achieved: {total_sent} mails sent"
#     body = f"{total_sent} mails sent today.\nLog saved at: {os.path.abspath(log_path)}"
#     send_smtp_mail(SENDER_EMAIL, SENDER_PASSWORD, SUMMARY_EMAIL, subject, body)

# def send_completion_email(log_path):
#     subject = "All rows iterated – check log"
#     body = f"All HR slots have been iterated.\nLog saved at: {os.path.abspath(log_path)}"
#     send_smtp_mail(SENDER_EMAIL, SENDER_PASSWORD, SUMMARY_EMAIL, subject, body)

# # -------------------------
# # Run
# # -------------------------
# if __name__ == "__main__":
#     print("Starting bulk mailer (TEST mode, no slot scheduling).")
#     run_day_cycle()



import pandas as pd
import numpy as np
from email.utils import formataddr
import os
import smtplib
import random
import time
from datetime import datetime, timedelta
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

SENDER_EMAIL = "hansrajjoshi_me22b9_70@dtu.ac.in"
SENDER_PASSWORD = "bhty xrms xisy yubn"  # use env vars in production

DRIVE_RESUME_LINK = "https://drive.google.com/file/d/17zHOeILK3883WVOJpDmDp4WdNRlDbRtE/view?usp=sharing"
RESUME_PDF_PATH = r"C:\Users\Administrator\Desktop\Karan\Hansraj_Joshi_96_pdf.pdf"

SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587

SUBJECT = "Internship"
BODY =  """
Hello {hr_name},

My name is Hansraj Joshi, and I am a final-year Mechanical Engineering student at Delhi Technological University (DTU).
I am writing to inquire about any potential internship openings at {company}.

{resume_line}

I would be grateful for the opportunity to discuss my qualifications further.

Sincerely,  
Hansraj Joshi  
"""

DAILY_TARGET = 50

MIN_DELAY_SECONDS = 0 * 60
MAX_DELAY_SECONDS = 1 * 60

SUMMARY_EMAIL = "hansjoshi2003no1@gmail.com"

# -------------------------
# Helpers
# -------------------------
def safe_str(x):
    return "" if pd.isna(x) else str(x).strip()

def company_to_filename(company_name):
    if pd.isna(company_name):
        cname = "UnknownCompany"
    else:
        cname = str(company_name).strip()
        cname = "_".join(cname.split())
    safe = "".join(ch for ch in cname if ch.isalnum() or ch == "_")
    return safe if safe else "Company"

def ensure_iter_cols(df):
    for hrn in ["HR1_iterated", "HR2_iterated", "HR3_iterated"]:
        if hrn not in df.columns:
            df[hrn] = False

def append_to_log(log_row):
    if os.path.exists(OUTBOX_LOG_PATH):
        try:
            existing = pd.read_excel(OUTBOX_LOG_PATH, sheet_name=OUTBOX_LOG_SHEET)
        except Exception:
            existing = pd.DataFrame()
    else:
        existing = pd.DataFrame()
    new_df = pd.concat([existing, pd.DataFrame([log_row])], ignore_index=True)
    with pd.ExcelWriter(OUTBOX_LOG_PATH, engine="openpyxl", mode="w") as writer:
        new_df.to_excel(writer, sheet_name=OUTBOX_LOG_SHEET, index=False)

def send_smtp_mail(sender, password, recipient, subject, body_text, attach_bytes=None, attach_filename=None):
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
        server = smtplib.SMTP(SMTP_HOST, SMTP_PORT)
        server.starttls()
        server.login(sender, password)
        server.sendmail(sender, recipient, msg.as_string())
        server.quit()
        return True, None
    except Exception as e:
        return False, str(e)

# -------------------------
# Core (no slot scheduling)
# -------------------------
def run_day_cycle():
    if not os.path.exists(BASELINE_XLSX_PATH):
        print(f"Baseline Excel not found at {BASELINE_XLSX_PATH}")
        return

    df = pd.read_excel(BASELINE_XLSX_PATH, sheet_name=SHEET_NAME, dtype=object)
    first_col = df.columns[0]
    df = df.rename(columns={first_col: "CompanyName"})
    df.columns = [c.strip() for c in df.columns]

    ensure_iter_cols(df)

    eligible_list = []
    for idx, row in df.iterrows():
        for i in [1, 2, 3]:
            iter_col = f"HR{i}_iterated"
            hr_name_col = f"HR{i}_Name"
            hr_email_col = f"Email ID_HR{i}"
            if hr_email_col not in df.columns:
                continue
            if bool(row.get(iter_col, False)):
                continue
            email_val = safe_str(row.get(hr_email_col, ""))
            if email_val not in ["", "0"]:
                eligible_list.append((idx, i, hr_name_col, hr_email_col))

    random.shuffle(eligible_list)

    if len(eligible_list) == 0:
        print("No eligible HR emails left.")
        return

    send_queue = eligible_list[:DAILY_TARGET]
    total_sent_today = 0

    for item in send_queue:
        row_idx, hr_no, hr_name_col, hr_email_col = item
        row = df.loc[row_idx]
        hr_name = safe_str(row.get(hr_name_col, ""))
        hr_email = safe_str(row.get(hr_email_col, ""))
        company_name = safe_str(row.get("CompanyName", ""))

        if hr_email in ["", "0"]:
            df.at[row_idx, f"HR{hr_no}_iterated"] = "NA"
            continue

        subject = SUBJECT.replace("{company}", company_name).replace("{hr_name}", hr_name)

        # Choose delivery mode
        delivery_mode = random.choice(["LINK", "PDF"])
        attach_bytes = None
        attach_filename = None
        if delivery_mode == "PDF" and os.path.exists(RESUME_PDF_PATH):
            with open(RESUME_PDF_PATH, "rb") as f:
                attach_bytes = f.read()
            sanitized = company_to_filename(company_name)
            attach_filename = f"Hansraj_Joshi_DTU_{sanitized}.pdf"
            resume_line = "I have attached my resume for your consideration."
        else:
            delivery_mode = "LINK"
            resume_line = f"You can view my resume here: {DRIVE_RESUME_LINK}"

        body_text = BODY.format(
            hr_name=hr_name or "Hiring Manager",
            company=company_name,
            resume_line=resume_line
        )

        success, error = send_smtp_mail(
            SENDER_EMAIL, SENDER_PASSWORD,
            hr_email, subject, body_text,
            attach_bytes=attach_bytes, attach_filename=attach_filename
        )

        timestamp_sent = datetime.now()
        iter_col = f"HR{hr_no}_iterated"

        if success:
            df.at[row_idx, iter_col] = True
            total_sent_today += 1
            status_str = "SENT"
            print(f"✅ Sent to: {hr_name} <{hr_email}>")
        else:
            df.at[row_idx, iter_col] = False
            status_str = "ERROR"
            print(f"❌ Failed to send to: {hr_name} <{hr_email}> | Error: {error}")

        log_row = {
            "CompanyName": company_name,
            "HR_slot": f"HR{hr_no}",
            "HR_name": hr_name if hr_name else "NA",
            "HR_email_used": hr_email,
            "HR_status": status_str,
            "timestamp": timestamp_sent,
            "sender": SENDER_EMAIL,
            "delivery_mode": delivery_mode,
            "error_message": error
        }
        append_to_log(log_row)

        with pd.ExcelWriter(BASELINE_XLSX_PATH, engine="openpyxl", mode="w") as writer:
            df.to_excel(writer, sheet_name=SHEET_NAME, index=False)

        if total_sent_today >= DAILY_TARGET:
            print("Daily target reached. Sending summary email.")
            send_summary_email(total_sent_today, OUTBOX_LOG_PATH)
            return

        delay = random.randint(MIN_DELAY_SECONDS, MAX_DELAY_SECONDS) + random.randint(0, 59)
        mins, secs = divmod(delay, 60)
        print(f"⏳ Next mail in: {mins}m {secs}s\n")
        time.sleep(delay)

    if total_sent_today > 0:
        print(f"Finished today's sends. Total sent: {total_sent_today}. Sending daily summary.")
        send_summary_email(total_sent_today, OUTBOX_LOG_PATH)

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

    with pd.ExcelWriter(BASELINE_XLSX_PATH, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name=SHEET_NAME, index=False)

    if all_iterated:
        print("All HR slots iterated. Sending completion email.")
        send_completion_email(OUTBOX_LOG_PATH)

def send_summary_email(total_sent, log_path):
    subject = f"Daily target achieved: {total_sent} mails sent"
    body = f"{total_sent} mails sent today.\nLog saved at: {os.path.abspath(log_path)}"
    send_smtp_mail(SENDER_EMAIL, SENDER_PASSWORD, SUMMARY_EMAIL, subject, body)

def send_completion_email(log_path):
    subject = "All rows iterated – check log"
    body = f"All HR slots have been iterated.\nLog saved at: {os.path.abspath(log_path)}"
    send_smtp_mail(SENDER_EMAIL, SENDER_PASSWORD, SUMMARY_EMAIL, subject, body)

# -------------------------
# Run
# -------------------------
if __name__ == "__main__":
    print("Starting bulk mailer (TEST mode, no slot scheduling).")
    run_day_cycle()