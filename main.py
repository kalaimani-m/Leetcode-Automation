"""
Daily LeetCode student progress report -- multi-batch version.

Usage:
    python main.py            # normal daily run -> one dated report per batch, one email with both attached
    python main.py --test     # test run -> TEST reports per batch, no email sent

Add or remove batches by editing the BATCHES list below. Each batch's Excel
file is read by COLUMN POSITION (S.No, Reg No, Name, LeetCode ID, Profile URL,
in that order) rather than by header text, so it doesn't matter that
different batches label their headers slightly differently
(e.g. "S.No" vs "SL.NO", "link" vs "LEETCODE LINK").
"""

import os
import csv
import time
import smtplib
import argparse
from email.message import EmailMessage
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill

from leetcode_client import extract_username, fetch_profile, count_unique_solved_today

IST = ZoneInfo("Asia/Kolkata")
BASE_DIR = os.path.dirname(__file__)
REPORTS_DIR = os.path.join(BASE_DIR, "reports")
DELAY_BETWEEN_STUDENTS = 2  # seconds -- be polite to LeetCode, avoid rate-limit bans

RECEIVER_EMAIL = "kalaimani.cybersec@prathyusha.edu.in"  # change this if you want a different destination

# ---- Batches -------------------------------------------------------------
# label          : shown in the email body
# students_file  : path to the master Excel file, relative to this script
# report_name    : prefix used for the dated report filename
# history_file   : append-only CSV this batch's daily numbers get logged to
BATCHES = [
    {
        "label": "Batch 1 (Reg. No. 111424xxxxxx)",
        "students_file": os.path.join(BASE_DIR, "students", "leetcode_Links.xlsx"),
        "report_name": "LeetCode_Daily_Report",
        "history_file": os.path.join(BASE_DIR, "data", "history.csv"),
    },
    {
        "label": "2nd Year (Reg. No. 111425xxxxxx)",
        "students_file": os.path.join(BASE_DIR, "students", "2nd_year_leetcode_links.xlsx"),
        "report_name": "LeetCode_Daily_Report_2ndYear",
        "history_file": os.path.join(BASE_DIR, "data", "history_2ndyear.csv"),
    },
]


def load_students(path):
    """
    Find the header row (search first 6 rows for something that looks like
    a serial-number column: 's.no', 'sl.no', 's no', etc.), then return the
    data below it. Columns are used by POSITION elsewhere, not by name, so
    it doesn't matter exactly what each batch calls its headers.
    """
    raw = pd.read_excel(path, header=None)
    header_row_idx = None
    for i in range(min(6, len(raw))):
        row_values = [str(v).strip().lower() for v in raw.iloc[i].tolist()]
        if any(("s.no" in v) or ("sl.no" in v) or ("s no" in v) for v in row_values):
            header_row_idx = i
            break
    if header_row_idx is None:
        raise ValueError(f"Could not find the header row in {path} (looking for a 'S.No'-like column).")

    df = pd.read_excel(path, header=header_row_idx)
    df = df.dropna(subset=[df.columns[0]])  # drop fully blank trailing rows
    return df.reset_index(drop=True)


def process_students(df, today_ist_date):
    results = []
    total = len(df)

    for i, row in df.iterrows():
        vals = row.tolist()
        sno = vals[0] if len(vals) > 0 else None
        regno = vals[1] if len(vals) > 1 else None
        name = vals[2] if len(vals) > 2 else None
        url = vals[4] if len(vals) > 4 else None

        username = extract_username(url) if isinstance(url, str) else None

        if not username:
            results.append({
                "sno": sno, "regno": regno, "name": name,
                "today": "N/A", "overall": "N/A",
                "note": "No usable profile URL in master file",
            })
            print(f"[{i+1}/{total}] {name}: SKIPPED (no profile URL)")
            continue

        profile = fetch_profile(username)

        if profile["status"] != "ok":
            results.append({
                "sno": sno, "regno": regno, "name": name,
                "today": "N/A", "overall": "N/A",
                "note": f"Profile unavailable ({profile['status']}: {profile['error']})",
            })
            print(f"[{i+1}/{total}] {name}: N/A ({profile['status']})")
        else:
            today_count, capped = count_unique_solved_today(profile["recent_ac"], today_ist_date)
            overall = profile["overall_solved"] if profile["overall_solved"] is not None else "N/A"
            note = ""
            if capped:
                note = "Recent-activity window is full (20 entries) -- today's count may be a lower bound"
            results.append({
                "sno": sno, "regno": regno, "name": name,
                "today": today_count if today_count is not None else "N/A",
                "overall": overall,
                "note": note,
            })
            print(f"[{i+1}/{total}] {name}: today={today_count} overall={overall}")

        time.sleep(DELAY_BETWEEN_STUDENTS)

    return results


def build_report(results, today_ist_date, out_path):
    wb = Workbook()

    ws = wb.active
    ws.title = "Daily Report"
    headers = ["S.No.", "Registration Number", "Student Name", "Current Date",
               "Problems Solved on Current Date", "Overall Problems Solved"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(name="Arial", bold=True)
        cell.fill = PatternFill("solid", fgColor="D9E1F2")
        cell.alignment = Alignment(horizontal="center")

    date_str = today_ist_date.strftime("%d-%b-%Y")
    for r in results:
        ws.append([r["sno"], r["regno"], r["name"], date_str, r["today"], r["overall"]])

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = Font(name="Arial")

    widths = [6, 20, 28, 14, 28, 22]
    for col, w in zip("ABCDEF", widths):
        ws.column_dimensions[col].width = w

    notes = [r for r in results if r["note"]]
    if notes:
        ws2 = wb.create_sheet("Notes")
        ws2.append(["Student Name", "Note"])
        for cell in ws2[1]:
            cell.font = Font(name="Arial", bold=True)
        for r in notes:
            ws2.append([r["name"], r["note"]])
        ws2.column_dimensions["A"].width = 28
        ws2.column_dimensions["B"].width = 70

    wb.save(out_path)


def append_history(results, today_ist_date, history_file):
    os.makedirs(os.path.dirname(history_file), exist_ok=True)
    file_exists = os.path.isfile(history_file)
    with open(history_file, "a", newline="") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(["Date", "Reg.no", "Student Name", "Problems Solved Today", "Overall Solved"])
        date_str = today_ist_date.strftime("%d-%b-%Y")
        for r in results:
            writer.writerow([date_str, r["regno"], r["name"], r["today"], r["overall"]])


def summarize(results):
    checked_ok = sum(1 for r in results if r["today"] != "N/A")
    unavailable = sum(1 for r in results if r["today"] == "N/A")
    solved_today = sum(1 for r in results if isinstance(r["today"], int) and r["today"] > 0)
    total_today = sum(r["today"] for r in results if isinstance(r["today"], int))
    return {
        "total": len(results), "checked_ok": checked_ok, "unavailable": unavailable,
        "solved_today": solved_today, "total_today": total_today,
    }


def send_email(attachments, today_ist_date, batch_summaries):
    """attachments: list of (filepath, label). batch_summaries: list of (label, summary_dict)."""
    sender = os.environ.get("SENDER_EMAIL")
    app_password = os.environ.get("SENDER_APP_PASSWORD")
    if not sender or not app_password:
        print("SENDER_EMAIL / SENDER_APP_PASSWORD not set -- skipping email send.")
        return

    date_str = today_ist_date.strftime("%d-%b-%Y")
    msg = EmailMessage()
    msg["Subject"] = f"Daily LeetCode Student Progress Report - {today_ist_date.strftime('%Y-%m-%d')}"
    msg["From"] = sender
    msg["To"] = RECEIVER_EMAIL

    body_lines = [
        "Dear Sir/Madam,",
        "",
        f"Please find attached the Daily LeetCode Student Progress Reports for {date_str}.",
        "",
    ]
    for label, summ in batch_summaries:
        body_lines.append(
            f"{label}: {summ['total']} students | {summ['checked_ok']} checked | "
            f"{summ['unavailable']} unavailable | {summ['solved_today']} solved \u22651 problem today | "
            f"{summ['total_today']} total problems solved today"
        )
    body_lines += ["", "Regards,", "Kalaimani"]
    msg.set_content("\n".join(body_lines))

    for filepath, _label in attachments:
        with open(filepath, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename=os.path.basename(filepath),
            )

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(sender, app_password)
        smtp.send_message(msg)
    print(f"Email sent to {RECEIVER_EMAIL} with {len(attachments)} attachment(s)")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--test", action="store_true", help="Test run: writes TEST reports, does not email")
    args = parser.parse_args()

    today_ist = datetime.now(IST).date()
    os.makedirs(REPORTS_DIR, exist_ok=True)

    attachments = []
    batch_summaries = []

    for batch in BATCHES:
        print(f"\n=== {batch['label']} ===")
        print(f"Loading: {batch['students_file']}")
        df = load_students(batch["students_file"])
        print(f"Loaded {len(df)} students.")

        results = process_students(df, today_ist)

        if args.test:
            out_path = os.path.join(REPORTS_DIR, f"{batch['report_name']}_TEST.xlsx")
        else:
            out_path = os.path.join(REPORTS_DIR, f"{batch['report_name']}_{today_ist.isoformat()}.xlsx")

        build_report(results, today_ist, out_path)
        print(f"Report written to {out_path}")

        append_history(results, today_ist, batch["history_file"])

        summ = summarize(results)
        batch_summaries.append((batch["label"], summ))
        attachments.append((out_path, batch["label"]))

        print(f"-- Summary: {summ}")

    print("\n--- Overall ---")
    for label, summ in batch_summaries:
        print(f"{label}: {summ}")

    if not args.test:
        send_email(attachments, today_ist, batch_summaries)
    else:
        print("\nTest run -- email NOT sent. Run without --test for the real daily run.")


if __name__ == "__main__":
    main()