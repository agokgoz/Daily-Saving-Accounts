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
import re
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
        "url": "https://www.ing.com.tr/tr/bilgi-destek/e-turuncu-faiz-oranlari",
        "custom_scraper": "ing",
    },
    "Akbank (Serbest Hesap)": {
        "url": "https://www.akbank.com/mevduat-yatirim/mevduat/vadeli-mevduat-hesaplari/serbest-plus-hesap",
        "custom_scraper": "akbank",
    },
    "CEPTETEB (Marifetli Hesap)": {
        "url": "https://www.cepteteb.com.tr/hesaplama/marifetli-hesap-faiz-hesaplama",
        "custom_scraper": "teb",
    },
    "QNB (Kazandıran Günlük Hesap)": {
        "url": "https://www.qnb.com.tr/kazandiran-gunluk-hesap",
        "custom_scraper": "qnb",
    },
}

# Path of the Excel file that stores historical rates.
EXCEL_FILE = "historical_rates.xlsx"

# Default page-load / selector timeout in milliseconds.
PAGE_TIMEOUT_MS = 30_000


# ---------------------------------------------------------------------------
# Bank-Specific Scraping Functions
# ---------------------------------------------------------------------------
# Some banks require specialized scraping logic that uses text locators
# instead of brittle CSS selectors. These custom functions navigate the DOM
# using visible text labels to find rate values more reliably.
# ---------------------------------------------------------------------------

def clean_rate_text(text: str) -> str:
    """
    Clean a rate string by removing the '%' sign, trimming whitespace,
    and replacing Turkish decimal commas with dots.
    Returns the cleaned string (e.g., '53' or '23.5').
    """
    cleaned = text.strip().replace("%", "").replace(",", ".").strip()
    return cleaned


def parse_rate_float(text: str) -> float:
    """
    Parse a cleaned rate string into a float.
    Returns 0.0 if parsing fails.
    """
    print(f"    [DEBUG] Raw extracted text: '{text}'")
    try:
        return float(clean_rate_text(text))
    except (ValueError, AttributeError):
        return 0.0


def extract_rate_via_js(page, keyword: str, bank_name: str) -> float:
    """
    Extract a rate value using JavaScript evaluation.

    Injects JS to find the keyword, go to its parent container, and extract
    the highest percentage value found (to avoid picking up term/day counts).

    Args:
        page: Playwright page object
        keyword: Text keyword to search for (e.g., "Hoş Geldin", "Standart")
        bank_name: Name of the bank for debug logging

    Returns:
        The extracted rate as a float, or 0.0 if extraction fails.
    """
    try:
        # Injects JS to find the keyword, go to its parent container, and extract the percentage
        js_code = f"""
        () => {{
            const elements = Array.from(document.querySelectorAll('*'));
            // Find element containing the keyword - check innerText to handle wrapped text in nested tags
            const target = elements.find(el => {{
                const text = el.innerText || el.textContent || "";
                return text.toLowerCase().includes('{keyword}'.toLowerCase());
            }});
            if (!target) return "";
            
            // Look at its parent row or container
            const container = target.closest('tr, .row, .flex, table') || target.parentElement?.parentElement || target;
            
            // Find all percentage values in this container - MUST have '%' sign to avoid matching plain numbers
            const text = container.innerText || "";
            const matches = text.match(/%\\s?\\d+[.,]?\\d*|\\d+[.,]?\\d*\\s?%/g);
            
            if (matches) {{
                // Return the highest percentage found in that row to ensure we get the rate, not a term/day count
                const numbers = matches.map(m => parseFloat(m.replace('%', '').replace(',', '.').trim()));
                return Math.max(...numbers).toString();
            }}
            return "";
        }}
        """
        rate_text = page.evaluate(js_code)
        print(f"    [DEBUG] {bank_name} '{keyword}' raw JS extraction: '{rate_text}'")
        return parse_rate_float(rate_text)
    except Exception as exc:
        print(f"  [WARN] JS extraction failed for {bank_name} ({keyword}): {exc}")
        return 0.0


def get_ing_rates(page) -> dict[str, float]:
    """
    Extract ING Turuncu Hesap interest rates using JavaScript evaluation.

    Returns:
        {"welcome_rate": float, "standard_rate": float}
    """
    return {
        "welcome_rate": extract_rate_via_js(page, "Hoş Geldin", "ING"),
        "standard_rate": extract_rate_via_js(page, "Standart", "ING")
    }


def get_akbank_rates(page) -> dict[str, float]:
    """
    Extract Akbank Serbest Hesap interest rates using JavaScript evaluation.

    Returns:
        {"welcome_rate": float, "standard_rate": float}
    """
    return {
        "welcome_rate": extract_rate_via_js(page, "Tanışma", "Akbank"),
        "standard_rate": extract_rate_via_js(page, "Standart", "Akbank")
    }


def get_qnb_rates(page) -> dict[str, float]:
    """
    Extract QNB Kazandıran Günlük Hesap interest rates using JavaScript evaluation.

    The QNB page renders rates client-side via JavaScript. This function uses
    DOM-based JS evaluation to extract the rates directly from the rendered page.

    Returns:
        {"welcome_rate": float, "standard_rate": float}
    """
    return {
        "welcome_rate": extract_rate_via_js(page, "Hoş Geldin", "QNB"),
        "standard_rate": extract_rate_via_js(page, "Standart", "QNB")
    }


def get_teb_rates(page) -> dict[str, float]:
    """
    Extract TEB Marifetli Hesap interest rates using JavaScript evaluation.

    Returns:
        {"welcome_rate": float, "standard_rate": float}
    """
    return {
        "welcome_rate": extract_rate_via_js(page, "Hoş Geldin", "TEB"),
        "standard_rate": extract_rate_via_js(page, "Standart", "TEB")
    }


def scrape_all_banks() -> dict[str, dict[str, str]]:
    """
    Visit every bank URL with a single Playwright browser session and collect
    the Welcome Rate and Standard Rate for each bank.

    Returns a nested dict:
      { bank_name: { "welcome_rate": "...", "standard_rate": "..." } }
    """
    # Map of custom scraper identifiers to functions
    CUSTOM_SCRAPERS = {
        "ing": get_ing_rates,
        "akbank": get_akbank_rates,
        "qnb": get_qnb_rates,
        "teb": get_teb_rates,
    }

    results: dict[str, dict[str, str]] = {}

    with sync_playwright() as pw:
        # Visible browser for debugging (headless=False)
        browser = pw.chromium.launch(headless=False)
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
                
                # Wait for observation (to see Cloudflare checks, cookie banners, etc.)
                page.wait_for_timeout(3000)

                # Attempt to dismiss cookie banners
                try:
                    # Using a CSS selector list for common Turkish accept buttons
                    cookie_btn = page.locator(
                        "button:has-text('Kabul Et'), "
                        "button:has-text('Tümünü Kabul Et'), "
                        "button:has-text('Anladım'), "
                        "button:has-text('Tercihlerimi Kaydet'), "
                        "a:has-text('Anladım')"
                    ).first
                    
                    # Try to click it if it appears within 2 seconds
                    cookie_btn.click(timeout=2000)
                    # Wait a moment for the banner animation to slide away
                    page.wait_for_timeout(1000)
                    print("  [DEBUG] Cookie banner dismissed.")
                except Exception:
                    # If no banner is found or it times out, safely ignore and continue
                    pass

                # Execute the custom scraper function
                custom_scraper_id = config["custom_scraper"]
                custom_fn = CUSTOM_SCRAPERS[custom_scraper_id]
                rates = custom_fn(page)
                # Convert float rates to strings for storage consistency
                welcome_rate = str(rates.get("welcome_rate", 0.0))
                standard_rate = str(rates.get("standard_rate", 0.0))

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
