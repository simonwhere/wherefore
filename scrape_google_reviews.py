#!/usr/bin/env python3
"""
scrape_google_reviews.py
------------------------
Local Playwright scraper for the PUBLIC Google Maps reviews panel of a single
business. Built for the SOPO Korean Eats case assessment, but works for any
place. Output matches the `raw_reviews` schema of SOPO_review_raw_analysis.xlsx
so the rows can be merged straight into that workbook.

============================ IMPORTANT / READ ME =============================
- This automates a normal browser viewing PUBLICLY visible reviews. It does
  NOT bypass logins, solve CAPTCHAs, rotate proxies, or spoof fingerprints.
- Scraping Google Maps at scale can conflict with Google's Terms of Service
  and robots restrictions. Whether you may do this is YOUR call and YOUR
  responsibility. Use a reasonable volume, keep the polite delays in place,
  and stop if Google shows a CAPTCHA / "unusual traffic" page (the script
  will detect this and exit rather than try to defeat it).
- Run it on YOUR machine, not in a restricted/cloud environment.
=============================================================================

SETUP (run once, on your laptop):
    python3 -m pip install playwright openpyxl
    python3 -m playwright install chromium

USAGE:
    # By place URL (recommended - copy the Google Maps URL of the business):
    python3 scrape_google_reviews.py \
        --url "https://www.google.com/maps/place/SOPO+Korean+Eats/..." \
        --business "SOPO Korean Eats" \
        --address "463 7th Ave, New York, NY 10018" \
        --out sopo_google_reviews.xlsx

    # Or by search query:
    python3 scrape_google_reviews.py \
        --query "SOPO Korean Eats 463 7th Ave New York" \
        --business "SOPO Korean Eats" --out sopo_google_reviews.xlsx

    # Watch it work (non-headless) and sort by newest:
    python3 scrape_google_reviews.py --url "..." --headful --sort newest

NOTE ON SELECTORS:
    Google Maps uses obfuscated, frequently-changing CSS class names. The
    selectors below worked at time of writing. If extraction returns 0 rows,
    run with --headful, open DevTools, and update the SELECTORS dict.
"""

import argparse
import datetime as dt
import random
import re
import sys
import time

TODAY = dt.date.today().isoformat()

# --- Selectors (update here if Google changes its DOM) ----------------------
SELECTORS = {
    "reviews_tab": 'button[role="tab"][aria-label*="Reviews"], button[aria-label*="Reviews for"]',
    "feed":        'div[role="main"] div.m6QErb[tabindex="-1"], div.m6QErb.DxyBCb',
    "review_card": 'div[data-review-id][jsaction]',
    "name":        '.d4r55',
    "rating_star": 'span[role="img"][aria-label*="star"]',
    "rating_text": '.kvMYJc',           # some locales render rating as element
    "date":        '.rsqaWe',
    "text":        '.wiI7pd',
    "more_btn":    'button[aria-label="See more"], button.w8nwRe',
    "sort_btn":    'button[aria-label*="Sort"], button[data-value="Sort"]',
}


def jitter(base):
    """Polite, slightly randomized delay (NOT for evasion - to let content load
    and avoid hammering the server)."""
    time.sleep(base + random.uniform(0, 0.6))


def parse_rating(aria):
    """'4.0 stars' / '별표 4개 만점에 5개' -> float or ''."""
    if not aria:
        return ""
    m = re.search(r"([0-5](?:[.,]\d)?)", aria)
    return m.group(1).replace(",", ".") if m else ""


def detect_block(page):
    """Detect CAPTCHA / 'unusual traffic' interstitials. We do NOT try to solve
    them - we stop and report, as required."""
    html = page.content().lower()
    for marker in ("unusual traffic", "recaptcha", "/sorry/", "are you a robot",
                   "비정상적인 트래픽"):
        if marker in html:
            return True
    return False


def scrape(args):
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
    rows = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not args.headful)
        ctx = browser.new_context(locale="en-US")
        page = ctx.new_page()

        target = args.url or (
            "https://www.google.com/maps/search/" + args.query.replace(" ", "+")
        )
        print(f"[+] Opening: {target}")
        page.goto(target, wait_until="domcontentloaded", timeout=60000)
        jitter(2.0)

        # Cookie consent (a normal visible user action; not a bypass)
        for label in ("Accept all", "Reject all", "I agree", "모두 수락"):
            try:
                btn = page.get_by_role("button", name=label)
                if btn.count():
                    btn.first.click(timeout=3000)
                    jitter(1.0)
                    break
            except Exception:
                pass

        if detect_block(page):
            print("[!] Google is showing a CAPTCHA / unusual-traffic page. "
                  "Stopping (this script does not defeat anti-bot challenges). "
                  "Try again later, slower, or from a normal browser session.")
            browser.close()
            return rows

        # Open the Reviews tab
        try:
            page.locator(SELECTORS["reviews_tab"]).first.click(timeout=10000)
            jitter(2.0)
        except PWTimeout:
            print("[!] Could not find the Reviews tab. If you passed --query, "
                  "open the exact place page and pass --url instead.")

        # Optional: sort by newest
        if args.sort == "newest":
            try:
                page.locator(SELECTORS["sort_btn"]).first.click(timeout=5000)
                jitter(1.0)
                page.get_by_role("menuitemradio", name=re.compile("Newest", re.I)).click(timeout=5000)
                jitter(2.0)
            except Exception:
                print("[i] Could not set sort=newest; continuing with default order.")

        # Locate the scrollable feed
        feed = page.locator(SELECTORS["feed"]).first
        try:
            feed.wait_for(timeout=10000)
        except PWTimeout:
            print("[!] Review feed not found. Update SELECTORS['feed'] (run --headful).")
            browser.close()
            return rows

        # Scroll until no new reviews load (or --max reached)
        seen = 0
        stagnant = 0
        print("[+] Scrolling to load reviews...")
        while True:
            feed.evaluate("el => el.scrollBy(0, el.scrollHeight)")
            jitter(args.delay)
            if detect_block(page):
                print("[!] Anti-bot page appeared mid-scroll. Stopping with what "
                      "we have so far.")
                break
            count = page.locator(SELECTORS["review_card"]).count()
            if count > seen:
                seen = count
                stagnant = 0
                print(f"    loaded {count} reviews...", end="\r")
            else:
                stagnant += 1
            if args.max and seen >= args.max:
                print(f"\n[+] Reached --max {args.max}.")
                break
            if stagnant >= 4:   # 4 consecutive no-growth scrolls => end of list
                print(f"\n[+] No more reviews loading. Total: {seen}.")
                break

        # Expand truncated reviews ("See more")
        for btn in page.locator(SELECTORS["more_btn"]).all():
            try:
                btn.click(timeout=1000)
            except Exception:
                pass
        jitter(1.0)

        # Extract
        cards = page.locator(SELECTORS["review_card"])
        n = cards.count()
        print(f"[+] Extracting {n} reviews...")
        for i in range(n):
            card = cards.nth(i)
            def grab(sel, attr=None):
                loc = card.locator(sel).first
                try:
                    if not loc.count():
                        return ""
                    return (loc.get_attribute(attr) if attr else loc.inner_text()).strip()
                except Exception:
                    return ""

            name = grab(SELECTORS["name"])
            rating = parse_rating(grab(SELECTORS["rating_star"], "aria-label")) \
                     or grab(SELECTORS["rating_text"])
            date = grab(SELECTORS["date"])          # relative, e.g. "2 weeks ago"
            text = grab(SELECTORS["text"])
            rid = card.get_attribute("data-review-id") or ""
            rurl = (args.url or target)
            if rid:
                rurl = f"{rurl}{'&' if '?' in rurl else '?'}review={rid}"

            rows.append({
                "platform": "Google Maps",
                "business_name": args.business,
                "location": args.location,
                "address": args.address,
                "review_date": f"{date} (relative; approximate)" if date else "",
                "rating": rating,
                "reviewer_name": name,
                "review_text": text,           # verbatim, as displayed
                "review_url": rurl,
                "source_url": args.url or target,
                "language": "",                # left blank; do not guess
                "menu_items_mentioned": "",
                "use_case": "",
                "customer_segment_guess": "",
                "notes": f"Scraped from Google Maps public reviews panel via Playwright on {TODAY}. "
                         f"Date is Google's relative label (e.g. '2 weeks ago') = approximate.",
            })

        browser.close()
    return rows


COLS = ["platform", "business_name", "location", "address", "review_date",
        "rating", "reviewer_name", "review_text", "review_url", "source_url",
        "language", "menu_items_mentioned", "use_case",
        "customer_segment_guess", "notes"]


def write_xlsx(rows, out):
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill
    wb = Workbook()
    ws = wb.active
    ws.title = "raw_reviews"
    ws.append(COLS)
    for c in range(1, len(COLS) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F3864")
    for r in rows:
        ws.append([r[c] for c in COLS])
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 24
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions["H"].width = 70
    ws.freeze_panes = "A2"
    wb.save(out)


def main():
    ap = argparse.ArgumentParser(description="Scrape PUBLIC Google Maps reviews (local use).")
    src = ap.add_mutually_exclusive_group(required=True)
    src.add_argument("--url", help="Google Maps place URL")
    src.add_argument("--query", help="Search query for the business")
    ap.add_argument("--business", default="", help="Business name")
    ap.add_argument("--location", default="", help="Neighborhood/city label")
    ap.add_argument("--address", default="", help="Street address")
    ap.add_argument("--out", default="google_reviews.xlsx", help="Output .xlsx")
    ap.add_argument("--max", type=int, default=0, help="Stop after N reviews (0 = all)")
    ap.add_argument("--sort", choices=["relevant", "newest"], default="relevant")
    ap.add_argument("--delay", type=float, default=1.8, help="Base seconds between scrolls (be polite)")
    ap.add_argument("--headful", action="store_true", help="Show the browser window")
    args = ap.parse_args()

    rows = scrape(args)
    if not rows:
        print("[!] No reviews collected. See messages above.")
        sys.exit(1)
    write_xlsx(rows, args.out)
    print(f"[✓] Wrote {len(rows)} reviews -> {args.out}")
    print("    Columns match raw_reviews in SOPO_review_raw_analysis.xlsx.")


if __name__ == "__main__":
    main()
