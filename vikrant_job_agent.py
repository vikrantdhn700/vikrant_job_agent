import os
import csv
import re
import smtplib
from email.message import EmailMessage
from datetime import date
import pandas as pd
from jobspy import scrape_jobs

# ─────────────────────────────────────────────
# Email Configuration
# ─────────────────────────────────────────────
SENDER_EMAIL   = "soulhacker1254@gmail.com"       # <-- change to your sender email
RECEIVER_EMAIL = "vikrant.dhn007@gmail.com"        # Vikrant's email
GMAIL_PASS     = os.getenv("GMAIL_PASSWORD")

# ─────────────────────────────────────────────
# Scraper Configuration
# Tailored for: Senior WordPress Developer | 9+ YoE
# ─────────────────────────────────────────────
search_terms = [
    "senior wordpress developer",
    "wordpress developer",
    "senior PHP developer",
    "wordpress engineer",
    "full stack wordpress developer",
    "wordpress theme plugin developer",
    "WooCommerce developer",
    "senior web developer PHP",
    "headless wordpress developer",
    "wordpress backend developer",
]

# ── Titles that are a GOOD match for Vikrant's level ──────────────────────────
APPROPRIATE_TITLE_KEYWORDS = re.compile(
    r"\b(wordpress|woocommerce|php|web\s*developer|full[\s-]?stack|"
    r"senior|sr\.?|lead\s*developer|engineer)\b",
    re.IGNORECASE,
)

# ── Titles to EXCLUDE (too junior, wrong domain, or exec-level) ───────────────
EXCLUDE_TITLE_KEYWORDS = re.compile(
    r"\b(intern|trainee|fresher|junior|jr\.?|entry[\s-]?level|"
    r"data\s*scientist|data\s*analyst|business\s*analyst|"
    r"sales|marketing|recruiter|devops|mobile\s*developer|"
    r"android|ios|flutter|react\s*native|"
    r"vp|chief|cto|ceo|director|head\s*of|"
    r"ui[\s/]?ux\s*designer|graphic\s*designer)\b",
    re.IGNORECASE,
)

# ── Skills from Vikrant's resume used to score job descriptions ───────────────
RESUME_SKILL_KEYWORDS = re.compile(
    r"\b(wordpress|woocommerce|php|javascript|html5?|css3?|"
    r"gutenberg|rest\s*api|plugin\s*development|custom\s*theme|"
    r"seo|google\s*tag\s*manager|google\s*analytics|"
    r"git|github|laravel|codeigniter|node\.?js|react|"
    r"mysql|mongodb|aws|lambda|api\s*gateway|"
    r"responsive\s*design|mobile[\s-]?first|crm\s*integration|"
    r"page\s*speed|schema\s*markup|structured\s*data|"
    r"python|sql|postman)\b",
    re.IGNORECASE,
)

# ─────────────────────────────────────────────
# Email Sender Logic
# ─────────────────────────────────────────────
def send_email(filtered_df):
    if not GMAIL_PASS:
        print("\n[ERROR] GMAIL_PASSWORD environment variable is not set. Email won't be sent.")
        return

    today = date.today().strftime("%d %b %Y")
    msg = EmailMessage()
    msg["From"] = SENDER_EMAIL
    msg["To"]   = RECEIVER_EMAIL

    if filtered_df is not None and not filtered_df.empty:
        subject = f"✅ Daily WordPress Job Matches (Senior) – {today}"
        body = (
            f"Hi Vikrant,\n\n"
            f"Great news! The scraper found {len(filtered_df)} relevant Senior WordPress / PHP roles today ({today}).\n\n"
            f"I've attached the Excel file with the full list. Here are the top 5 roles for a quick glance:\n\n"
        )

        top_5 = filtered_df.head(5)
        for i, (_, row) in enumerate(top_5.iterrows(), 1):
            body += (
                f"#{i} {row['title']} @ {row['company']}\n"
                f"Location      : {row.get('location', 'N/A')}\n"
                f"Remote        : {row.get('is_remote', 'N/A')}\n"
                f"Match Score   : {row['relevance_score']}\n"
                f"Apply Here    : {row['job_url']}\n"
                f"{'─' * 45}\n"
            )

        body += (
            "\nTips for applying:\n"
            "• Highlight your 9+ years of WordPress experience upfront.\n"
            "• Mention Gutenberg block development and custom plugin work (Repair Stack, Pre Order).\n"
            "• Emphasise your SEO & GTM integration experience — it's a differentiator.\n\n"
            "Good luck! 🚀\n\n— Your Job Alert Bot 🤖"
        )
        msg.set_content(body)
        msg["Subject"] = subject

        try:
            with open("jobs.xlsx", "rb") as f:
                excel_data = f.read()
            msg.add_attachment(
                excel_data,
                maintype="application",
                subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename=f"WordPress_Senior_Jobs_{today.replace(' ', '_')}.xlsx"
            )
        except Exception as e:
            print(f"[WARN] Could not attach Excel file: {e}")

    else:
        subject = f"❌ No Senior WordPress Job Matches – {today}"
        body = (
            f"Hi Vikrant,\n\n"
            f"No highly relevant Senior WordPress / PHP roles were found today ({today}).\n\n"
            f"I'll keep checking tomorrow!\n\n"
            f"— Your Job Alert Bot 🤖"
        )
        msg.set_content(body)
        msg["Subject"] = subject

    print(f"\n[INFO] Sending email: {subject}")
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(SENDER_EMAIL, GMAIL_PASS)
            smtp.send_message(msg)
        print("[INFO] Email sent successfully ✓")
    except Exception as e:
        print(f"[ERROR] Failed to send email: {e}")


# ─────────────────────────────────────────────
# Main Scraper Execution
# ─────────────────────────────────────────────
if __name__ == "__main__":
    all_jobs = []

    for term in search_terms:
        print(f"\n🔍 Searching for: {term}")
        try:
            jobs = scrape_jobs(
                site_name=["indeed", "linkedin", "google"],
                search_term=term,
                google_search_term=f"{term} jobs India remote hybrid senior",
                location="India",
                country_indeed="India",
                results_wanted=25,
                hours_old=24,
                description_format="markdown",
                verbose=1,
            )
            print(f"   Found {len(jobs)} jobs")
            all_jobs.append(jobs)
        except Exception as e:
            print(f"   Error searching '{term}': {e}")

    if all_jobs:
        combined = pd.concat(all_jobs, ignore_index=True)
        combined = combined.drop_duplicates(subset=["job_url"], keep="first")
        before = len(combined)

        # ── Filter out excluded titles ─────────────────────────────────────────
        combined["is_excluded"] = combined["title"].apply(
            lambda t: bool(EXCLUDE_TITLE_KEYWORDS.search(str(t)))
        )
        filtered = combined[~combined["is_excluded"]].copy()
        filtered = filtered.drop(columns=["is_excluded"])

        # ── Keep only titles with at least one appropriate keyword ─────────────
        filtered = filtered[
            filtered["title"].apply(lambda t: bool(APPROPRIATE_TITLE_KEYWORDS.search(str(t))))
        ].copy()

        # ── Score by how many resume skills appear in the description ──────────
        def skill_match_score(desc):
            if pd.isna(desc):
                return 0
            return len(RESUME_SKILL_KEYWORDS.findall(str(desc)))

        filtered["relevance_score"] = filtered["description"].apply(skill_match_score)
        filtered = filtered.sort_values(
            ["relevance_score", "date_posted"],
            ascending=[False, False],
            na_position="last",
        )

        excluded = before - len(filtered)
        print(
            f"\n📊 Filtering: {before} total → {len(filtered)} relevant senior roles "
            f"(excluded {excluded} junior/irrelevant)"
        )

        keep_columns = [
            "site", "title", "company", "location", "date_posted",
            "job_type", "is_remote", "job_level", "job_url",
            "company_url", "relevance_score",
        ]
        export_columns = [c for c in keep_columns if c in filtered.columns]
        output = filtered[export_columns].copy()

        output.to_csv("jobs.csv", quoting=csv.QUOTE_NONNUMERIC, escapechar="\\", index=False)
        output.to_excel("jobs.xlsx", index=False)

        print(f"\n✅ Total relevant jobs found: {len(filtered)}")
        send_email(output)

    else:
        print("\n❌ No jobs found.")
        send_email(None)
