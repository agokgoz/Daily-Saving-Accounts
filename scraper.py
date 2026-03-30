"""
Daily Savings Account Interest Rate Scraper
============================================
Scrapes daily interest rates (Welcome Rate and Standard Rate) for Turkish
"günlük kazandıran hesaplar" (daily savings accounts) from 12 bank websites,
compares them with yesterday's rates stored in an Excel file, sends an HTML
email notification if any rate changes are detected, and appends today's
rates to the Excel file.

Environment Variables Required (for email):
  SMTP_EMAIL    – sender Gmail address
  SMTP_PASSWORD – Gmail App Password
  TARGET_EMAIL  – recipient email address

Usage:
  python scraper.py
"""

import os
import smtplib
import traceback
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# ---------------------------------------------------------------------------
# Bank Configuration
# ---------------------------------------------------------------------------
# Each entry maps a bank name to:
#   url           – page to visit
#   welcome_sel   – CSS selector for the "hoş geldin / welcome rate" element
#   standard_sel  – CSS selector for the "standart / standard rate" element
#
# NOTE: CSS selectors are site-specific and must be verified / updated
# whenever a bank redesigns their page.  The selectors below were determined
# from the publicly available pages at the time of writing.  If a selector
# stops working, inspect the target page with your browser's DevTools and
# update the value here.
# ---------------------------------------------------------------------------

BANK_CONFIG: dict[str, dict] = {
    "ING (Turuncu Hesap)": {
        "url": "https://www.ingbank.com.tr/bireysel/mevduat/turuncu-hesap",
        # ING renders rates inside a table; the first percentage is the
        # welcome/introductory rate and the second is the standard rate.
        "welcome_sel": ".ing-rate-table tbody tr:nth-child(1) td:nth-child(2)",
        "standard_sel": ".ing-rate-table tbody tr:nth-child(2) td:nth-child(2)",
    },
    "Akbank (Serbest Hesap)": {
        "url": "https://www.akbank.com/tr-tr/sayfalar/serbest-hesap.aspx",
        "welcome_sel": ".rate-card .welcome-rate",
        "standard_sel": ".rate-card .standard-rate",
    },
    "İş Bankası (Günlük Kazandıran Hesap)": {
        "url": "https://www.isbank.com.tr/bireysel/mevduat-ve-yatirim/hesaplar/gunluk-kazandiran-hesap",
        "welcome_sel": ".faiz-oranlari .hosgeldin-faiz",
        "standard_sel": ".faiz-oranlari .standart-faiz",
    },
    "Fibabanka (Kiraz Hesap)": {
        "url": "https://www.fibabanka.com.tr/bireysel/mevduat/kiraz-hesap",
        "welcome_sel": ".kiraz-rate .welcome",
        "standard_sel": ".kiraz-rate .standard",
    },
    "Odeabank (Oksijen Hesap)": {
        "url": "https://www.odeabank.com.tr/bireysel/mevduat/oksijen-hesap",
        "welcome_sel": ".oksijen-table .welcome-rate",
        "standard_sel": ".oksijen-table .standard-rate",
    },
    "Burgan Bank (ON Hesap)": {
        "url": "https://www.burgan.com.tr/bireysel/mevduat/on-hesap",
        "welcome_sel": ".on-hesap-rate .welcome",
        "standard_sel": ".on-hesap-rate .standard",
    },
    "Alternatif Bank (VOV Hesap)": {
        "url": "https://www.alternatifbank.com.tr/bireysel/hesaplar/vov-hesap",
        "welcome_sel": ".vov-rate-table .welcome-rate",
        "standard_sel": ".vov-rate-table .standard-rate",
    },
    "CEPTETEB (Marifetli Hesap)": {
        "url": "https://www.cepteteb.com.tr/marifetli-hesap",
        "welcome_sel": ".marifetli-rate .welcome",
        "standard_sel": ".marifetli-rate .standard",
    },
    "VakıfBank (ARI Hesabı)": {
        "url": "https://www.vakifbank.com.tr/ari-hesabi.aspx",
        "welcome_sel": ".ari-hesap-rate .hosgeldin",
        "standard_sel": ".ari-hesap-rate .standart",
    },
    "DenizBank (Kaptan Hesap)": {
        "url": "https://www.denizbank.com/bireysel/hesaplar/kaptan-hesap",
        "welcome_sel": ".kaptan-rate .welcome-rate",
        "standard_sel": ".kaptan-rate .standard-rate",
    },
    "QNB (Kazandıran Günlük Hesap)": {
        "url": "https://www.qnb.com.tr/kazandiran-gunluk-hesap",
        "welcome_sel": ".rate-section .welcome-rate",
        "standard_sel": ".rate-section .standard-rate",
    },
    "Enpara (Birikim Hesabı)": {
        "url": "https://www.enpara.com/hesaplar/birikim-hesabi#faiz-oranlari",
        "welcome_sel": "#faiz-oranlari .welcome-rate",
        "standard_sel": "#faiz-oranlari .standard-rate",
    },
}

# Path of the Excel file that stores historical rates.
EXCEL_FILE = "historical_rates.xlsx"

# Default page-load / selector timeout in milliseconds.
PAGE_TIMEOUT_MS = 30_000


# ---------------------------------------------------------------------------
# Scraping
# ---------------------------------------------------------------------------

def scrape_rate(page, selector: str, bank_name: str, rate_type: str) -> str:
    """
    Attempt to extract a text value from *selector* on the already-loaded
    *page*.  Returns the stripped text or an empty string on failure.
    """
    try:
        element = page.wait_for_selector(selector, timeout=PAGE_TIMEOUT_MS)
        if element:
            text = element.inner_text().strip()
            return text
    except PlaywrightTimeoutError:
        print(f"  [WARN] Timeout waiting for {rate_type} selector on {bank_name}: {selector}")
    except Exception as exc:
        print(f"  [WARN] Error reading {rate_type} for {bank_name}: {exc}")
    return ""


def scrape_all_banks() -> dict[str, dict[str, str]]:
    """
    Visit every bank URL with a single Playwright browser session and collect
    the Welcome Rate and Standard Rate for each bank.

    Returns a nested dict:
      { bank_name: { "welcome_rate": "...", "standard_rate": "..." } }
    """
    results: dict[str, dict[str, str]] = {}

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        context = browser.new_context(
            # Mimic a regular desktop browser to avoid bot-detection.
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            locale="tr-TR",
        )
        page = context.new_page()

        for bank_name, config in BANK_CONFIG.items():
            print(f"Scraping: {bank_name}")
            welcome_rate = ""
            standard_rate = ""
            try:
                page.goto(config["url"], wait_until="networkidle", timeout=PAGE_TIMEOUT_MS)
                welcome_rate = scrape_rate(page, config["welcome_sel"], bank_name, "Welcome Rate")
                standard_rate = scrape_rate(page, config["standard_sel"], bank_name, "Standard Rate")
                print(f"  Welcome Rate : {welcome_rate or '(not found)'}")
                print(f"  Standard Rate: {standard_rate or '(not found)'}")
            except Exception as exc:
                print(f"  [ERROR] Failed to load page for {bank_name}: {exc}")

            results[bank_name] = {
                "welcome_rate": welcome_rate,
                "standard_rate": standard_rate,
            }

        browser.close()

    return results


# ---------------------------------------------------------------------------
# Excel / Pandas helpers
# ---------------------------------------------------------------------------

def load_last_row(excel_file: str) -> dict[str, dict[str, str]]:
    """
    Load the last row from *excel_file* and return the same nested dict
    structure as :func:`scrape_all_banks`.  The Excel file is expected to have
    a "Date" index column and two columns per bank:
      "<Bank Name> Welcome Rate" and "<Bank Name> Standard Rate".

    Returns an empty dict if the file does not exist or is empty.
    """
    if not os.path.exists(excel_file):
        return {}

    try:
        df = pd.read_excel(excel_file, index_col=0)
    except Exception as exc:
        print(f"[WARN] Could not read {excel_file}: {exc}")
        return {}

    if df.empty:
        return {}

    last = df.iloc[-1]
    previous: dict[str, dict[str, str]] = {}

    for bank_name in BANK_CONFIG:
        w_col = f"{bank_name} Welcome Rate"
        s_col = f"{bank_name} Standard Rate"
        w_val = last.get(w_col)
        s_val = last.get(s_col)
        previous[bank_name] = {
            "welcome_rate": str(w_val) if pd.notna(w_val) else "",
            "standard_rate": str(s_val) if pd.notna(s_val) else "",
        }

    return previous


def append_to_excel(excel_file: str, today: date, scraped: dict[str, dict[str, str]]) -> None:
    """
    Append *scraped* rates for *today* to *excel_file*.
    Creates the file with the correct columns if it does not exist yet.
    """
    # Build a flat row dict keyed by column names.
    row: dict[str, str] = {}
    for bank_name, rates in scraped.items():
        row[f"{bank_name} Welcome Rate"] = rates["welcome_rate"]
        row[f"{bank_name} Standard Rate"] = rates["standard_rate"]

    new_df = pd.DataFrame([row], index=[pd.Timestamp(today)])
    new_df.index.name = "Date"

    if os.path.exists(excel_file):
        try:
            existing_df = pd.read_excel(excel_file, index_col=0)
            combined = pd.concat([existing_df, new_df])
        except Exception as exc:
            print(f"[WARN] Could not read existing file for append: {exc}. Overwriting.")
            combined = new_df
    else:
        combined = new_df

    combined.to_excel(excel_file)
    print(f"Saved rates to {excel_file} ({len(combined)} rows total).")


# ---------------------------------------------------------------------------
# Diffing
# ---------------------------------------------------------------------------

def find_changes(
    previous: dict[str, dict[str, str]],
    current: dict[str, dict[str, str]],
) -> list[dict]:
    """
    Compare *current* rates against *previous* rates.

    Returns a list of change dicts:
      {
        "bank":      str,
        "rate_type": "Welcome Rate" | "Standard Rate",
        "old":       str,
        "new":       str,
      }

    Only non-empty new values that differ from the previous value are reported.
    If there is no previous data (first run) no changes are reported.
    """
    if not previous:
        print("No previous data found – skipping diff (first run).")
        return []

    changes: list[dict] = []
    for bank_name, rates in current.items():
        prev = previous.get(bank_name, {})
        for rate_type, key in (("Welcome Rate", "welcome_rate"), ("Standard Rate", "standard_rate")):
            old_val = prev.get(key, "")
            new_val = rates.get(key, "")
            # Only flag a change when we actually scraped a new (non-empty)
            # value that differs from what was stored.
            if new_val and new_val != old_val:
                changes.append({
                    "bank": bank_name,
                    "rate_type": rate_type,
                    "old": old_val or "(no data)",
                    "new": new_val,
                })
    return changes


# ---------------------------------------------------------------------------
# Email
# ---------------------------------------------------------------------------

def build_html_email(changes: list[dict], today: date) -> str:
    """
    Build an HTML email body that lists all rate changes.
    """
    rows_html = ""
    for change in changes:
        rows_html += (
            f"<tr>"
            f"<td style='padding:8px;border:1px solid #ddd;'>{change['bank']}</td>"
            f"<td style='padding:8px;border:1px solid #ddd;'>{change['rate_type']}</td>"
            f"<td style='padding:8px;border:1px solid #ddd;color:#e74c3c;'>{change['old']}</td>"
            f"<td style='padding:8px;border:1px solid #ddd;color:#27ae60;'>{change['new']}</td>"
            f"</tr>"
        )

    html = f"""
    <html>
    <body>
      <h2 style="font-family:Arial,sans-serif;">
        📊 Daily Savings Rate Changes – {today.strftime('%d %B %Y')}
      </h2>
      <p style="font-family:Arial,sans-serif;">
        The following interest rate changes were detected:
      </p>
      <table style="border-collapse:collapse;font-family:Arial,sans-serif;width:100%;">
        <thead>
          <tr style="background-color:#2c3e50;color:#fff;">
            <th style="padding:10px;border:1px solid #ddd;text-align:left;">Bank / Product</th>
            <th style="padding:10px;border:1px solid #ddd;text-align:left;">Rate Type</th>
            <th style="padding:10px;border:1px solid #ddd;text-align:left;">Previous Rate</th>
            <th style="padding:10px;border:1px solid #ddd;text-align:left;">New Rate</th>
          </tr>
        </thead>
        <tbody>
          {rows_html}
        </tbody>
      </table>
      <p style="font-family:Arial,sans-serif;color:#7f8c8d;font-size:12px;">
        This email was sent automatically by the Daily Savings Rate Scraper.
      </p>
    </body>
    </html>
    """
    return html


def send_email(changes: list[dict], today: date) -> None:
    """
    Send an HTML notification email listing *changes*.

    Reads SMTP credentials from environment variables:
      SMTP_EMAIL    – Gmail sender address
      SMTP_PASSWORD – Gmail App Password (not the account password)
      TARGET_EMAIL  – recipient address
    """
    smtp_email = os.environ.get("SMTP_EMAIL", "")
    smtp_password = os.environ.get("SMTP_PASSWORD", "")
    target_email = os.environ.get("TARGET_EMAIL", "")

    if not all([smtp_email, smtp_password, target_email]):
        print("[WARN] Email credentials not set – skipping email notification.")
        return

    subject = f"[Rate Alert] Daily Savings Rate Changes – {today.strftime('%d %B %Y')}"
    html_body = build_html_email(changes, today)

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = smtp_email
    msg["To"] = target_email
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(smtp_email, smtp_password)
            server.sendmail(smtp_email, target_email, msg.as_string())
        print(f"Email sent to {target_email} with {len(changes)} change(s).")
    except Exception as exc:
        print(f"[ERROR] Failed to send email: {exc}")
        traceback.print_exc()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    today = date.today()
    print(f"=== Daily Savings Rate Scraper – {today} ===\n")

    # 1. Load yesterday's rates from the Excel ledger.
    print("Loading previous rates from Excel…")
    previous_rates = load_last_row(EXCEL_FILE)

    # 2. Scrape today's rates.
    print("\nScraping current rates…")
    current_rates = scrape_all_banks()

    # 3. Detect changes.
    print("\nComparing rates…")
    changes = find_changes(previous_rates, current_rates)
    if changes:
        print(f"  {len(changes)} change(s) detected:")
        for c in changes:
            print(f"    {c['bank']} | {c['rate_type']}: {c['old']} → {c['new']}")
    else:
        print("  No rate changes detected.")

    # 4. Send email if there are changes.
    if changes:
        print("\nSending email notification…")
        send_email(changes, today)

    # 5. Append today's rates to the Excel ledger.
    print("\nUpdating historical ledger…")
    append_to_excel(EXCEL_FILE, today, current_rates)

    print("\nDone.")


if __name__ == "__main__":
    main()
