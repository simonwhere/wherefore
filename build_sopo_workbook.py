#!/usr/bin/env python3
"""
Build SOPO_review_raw_analysis.xlsx

Integrity rules applied (per task brief):
- No invented reviews, names, dates, ratings, or exact text.
- Verbatim social captions are taken from search-indexed page titles (which ARE
  the public caption text). Truncation is flagged where the index cut the text.
- Platform review pages (Yelp/Google/Tripadvisor/delivery apps) could NOT be
  read: WebFetch returned HTTP 403 on every URL in this environment. Where only
  a paraphrased search-engine summary exists, the row text is prefixed with
  "[PARAPHRASED SEARCH SUMMARY - NOT VERBATIM]" and reviewer_name/date/rating are
  left blank rather than guessed.
- Aggregate ratings and access blocks are documented in source_log.
- Date accessed: 2026-06-08.
"""

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

ACCESSED = "2026-06-08"
BIZ = "SOPO Korean Eats"
LOC = "Midtown / Herald Square, Manhattan, New York, NY"
ADDR = "463 7th Ave, New York, NY 10018"

# ---------------------------------------------------------------------------
# RAW REVIEWS
# Each record: id, platform, review_date, rating, reviewer_name, review_text,
#              review_url, source_url, language, menu_items, use_case,
#              segment_guess, notes
# ---------------------------------------------------------------------------
PARA = "[PARAPHRASED SEARCH SUMMARY - NOT VERBATIM] "

raw = [
    # --- Tier 1: verbatim public social captions (from search-indexed titles) ---
    dict(
        id="R01", platform="TikTok",
        review_date="2024-11-14 (approximate)", rating="",
        reviewer_name="@tiffany.elizabeth",
        review_text="lunch at Sopo near Herald Square in nyc\U0001F90D this korean fast casual eatery was sooo good and i cant wait to go back",
        review_url="https://www.tiktok.com/@tiffany.elizabeth/video/7437319271316884782",
        source_url="https://www.tiktok.com/@tiffany.elizabeth/video/7437319271316884782",
        language="en", menu_items="", use_case="Lunch (near Herald Square)",
        segment_guess="Food content creator / NYC food influencer",
        notes="Verbatim public caption (captured from search-indexed page title). Date per search summary = approximate. Full comment thread not accessible (WebFetch 403).",
    ),
    dict(
        id="R02", platform="TikTok",
        review_date="Unknown (video id suggests ~Oct 2024)", rating="",
        reviewer_name="@yourbrooklynguide",
        review_text="Best places to eat in Midtown: Sopo Korean Eats Fast casual Korean comfort food plates with banchan plus kimbap! The menu was designed by a 3 Michelin star chef! We were impressed with how good of a fast food option this was and will be keeping this as a handy spot to visit when we are out and about in Midtown. \U0001F4CD463 7th Ave #nyceats #midtownmanhattan #nycfood #koreanfood",
        review_url="https://www.tiktok.com/@yourbrooklynguide/video/7426773260227972394",
        source_url="https://www.tiktok.com/@yourbrooklynguide/video/7426773260227972394",
        language="en", menu_items="banchan plates; kimbap",
        use_case="Midtown casual meal / repeat go-to spot",
        segment_guess="NYC food/travel content creator",
        notes="Verbatim public caption (search-indexed title). Date not shown; estimate from video id only. Comments not accessible (WebFetch 403).",
    ),
    dict(
        id="R03", platform="TikTok",
        review_date="Unknown", rating="",
        reviewer_name="@jonathanchoione",
        review_text="Former Michelin Three Star Alumni At Le Bernadin Leads One Of The Hottest Fast Casual Restaurant In NYC Offering High Quality Affordable Korean Food. Sopo Is The First Fast Casual Korean Restaurant In NYC To Offer Dishes Banchan Style. For Under $20, You Can Create A Chef-Quality Korean Meal That Consists Of Your Favorite Korean Flavors And Sides. Perfect Meal For Lunch Or Dinner. \U0001F4CD Sopo Korean Eats - 463 7th Ave, New York, NY 10018",
        review_url="https://www.tiktok.com/@jonathanchoione/video/7513986853226368302",
        source_url="https://www.tiktok.com/@jonathanchoione/video/7513986853226368302",
        language="en", menu_items="banchan-style dishes; under-$20 meals",
        use_case="Lunch or dinner",
        segment_guess="Food creator (caption reads promotional/branded)",
        notes="Verbatim public caption (search-indexed title). Tone is promotional and may be sponsored/branded content - weight accordingly. Comments not accessible (WebFetch 403).",
    ),
    dict(
        id="R04", platform="TikTok",
        review_date="Unknown", rating="",
        reviewer_name="@storytimenyc",
        review_text="Discover Sopo: A Unique Korean Fast Casual Dining ...",
        review_url="https://www.tiktok.com/@storytimenyc/video/7476600656900427051",
        source_url="https://www.tiktok.com/@storytimenyc/video/7476600656900427051",
        language="en", menu_items="", use_case="Casual dining",
        segment_guess="NYC food content creator",
        notes="Caption TRUNCATED in search index ('...'); full text not accessible (WebFetch 403). Title may be auto-generated rather than user caption - treat as partial.",
    ),
    dict(
        id="R05", platform="TikTok",
        review_date="Unknown", rating="",
        reviewer_name="@codypremer",
        review_text="Sopo's Favorite Meal Review: Korean Cuisine in New York",
        review_url="https://www.tiktok.com/@codypremer/video/7469557158112251179",
        source_url="https://www.tiktok.com/@codypremer/video/7469557158112251179",
        language="en", menu_items="", use_case="Meal review",
        segment_guess="Food content creator",
        notes="Title captured from search index; likely an auto-generated/short title rather than full caption. Full caption + comments not accessible (WebFetch 403).",
    ),
    dict(
        id="R06", platform="Instagram",
        review_date="Unknown", rating="",
        reviewer_name="missy (@girlswholiketoeatfood)",
        review_text="SOPO is a new fast-casual Korean spot located right ...",
        review_url="https://www.instagram.com/girlswholiketoeatfood/reel/DABKm79xVlr/",
        source_url="https://www.instagram.com/girlswholiketoeatfood/reel/DABKm79xVlr/",
        language="en", menu_items="", use_case="Casual / new-spot discovery",
        segment_guess="NYC food influencer (account run by 'missy & jessie')",
        notes="Caption TRUNCATED in search index ('...'); full reel caption + comments not accessible (WebFetch 403).",
    ),
    dict(
        id="R07", platform="Instagram",
        review_date="Unknown", rating="",
        reviewer_name="MealPal (@mealpal)",
        review_text="Sopo serves fast casual Korean ...",
        review_url="https://www.instagram.com/reel/DK9u8j_MiBn/",
        source_url="https://www.instagram.com/reel/DK9u8j_MiBn/",
        language="en", menu_items="", use_case="Fast-casual lunch",
        segment_guess="Brand/app account (MealPal)",
        notes="Caption TRUNCATED in search index ('...'); brand account. Full caption + comments not accessible (WebFetch 403).",
    ),
    dict(
        id="R08", platform="TikTok",
        review_date="Unknown", rating="",
        reviewer_name="",
        review_text="From bulgogi perfection to saucy dumplings that hit all the right notes, this place does NOT miss.",
        review_url="",
        source_url="https://www.tiktok.com/discover/sopo-nyc",
        language="en", menu_items="bulgogi; saucy dumplings",
        use_case="Casual dining",
        segment_guess="TikTok food creator (unidentified)",
        notes="Quote surfaced via TikTok search summary; EXACT source video URL not confirmed and author not identified. Provenance uncertain - treat as indicative only.",
    ),

    # --- Tier 2: paraphrased perception summaries from blocked review platforms ---
    # reviewer_name/date/rating intentionally BLANK (not accessible / not invented).
    dict(
        id="R09", platform="Yelp",
        review_date="", rating="",
        reviewer_name="",
        review_text=PARA + "Really impressed with the quality SOPO delivers every single time; the food is super flavorful, fresh, outstanding Korean food.",
        review_url="https://www.yelp.com/biz/sopo-new-york-3",
        source_url="https://www.yelp.com/biz/sopo-new-york-3",
        language="en", menu_items="", use_case="Repeat visits",
        segment_guess="Yelp reviewer (not identified)",
        notes="NOT VERBATIM. Yelp page returned HTTP 403 to automated fetch; this is a search-engine summary of one or more positive Yelp reviews. Individual reviewer name/date/exact wording NOT accessible.",
    ),
    dict(
        id="R10", platform="Yelp",
        review_date="", rating="",
        reviewer_name="",
        review_text=PARA + "The gold pumpkin mash was out of this world, and the saucy dumplings were also something quite special.",
        review_url="https://www.yelp.com/biz/sopo-new-york-3",
        source_url="https://www.yelp.com/biz/sopo-new-york-3",
        language="en", menu_items="gold pumpkin mash; saucy dumplings",
        use_case="Dine-in / standout dishes",
        segment_guess="Yelp reviewer (not identified)",
        notes="NOT VERBATIM. Search-engine summary of Yelp content; page blocked (403). Reviewer/date/exact text not accessible.",
    ),
    dict(
        id="R11", platform="Yelp",
        review_date="", rating="",
        reviewer_name="",
        review_text=PARA + "Mixed report: the beef bulgogi was not flavorful at all, with barely any marinade, and was not the best quality meat.",
        review_url="https://www.yelp.com/biz/sopo-new-york-3",
        source_url="https://www.yelp.com/biz/sopo-new-york-3",
        language="en", menu_items="beef bulgogi",
        use_case="Dine-in / quality complaint",
        segment_guess="Yelp reviewer (not identified)",
        notes="NOT VERBATIM. Search-engine summary of a critical Yelp review; page blocked (403). Included to represent negative sentiment. Reviewer/date/exact text not accessible.",
    ),
    dict(
        id="R12", platform="Google / aggregator search summary",
        review_date="", rating="",
        reviewer_name="",
        review_text=PARA + "Best Korean fast casual in the city, with super good pricing and even better food; flavorful beef bulgogi and friendly staff who explain the dishes.",
        review_url="",
        source_url="https://www.google.com/maps/search/SOPO+Korean+Eats+463+7th+Ave",
        language="en", menu_items="beef bulgogi",
        use_case="Quick casual meal",
        segment_guess="General reviewer (not identified)",
        notes="NOT VERBATIM. Aggregated/paraphrased positive sentiment from search summaries; underlying platform not confirmed (likely Google/Maps). Could not open Google Maps reviews directly (WebFetch 403).",
    ),
    dict(
        id="R13", platform="Aggregator search summary",
        review_date="", rating="",
        reviewer_name="",
        review_text=PARA + "The kimbap is really good but a bit small for the price, especially when competing with K-town prices a few blocks over.",
        review_url="",
        source_url="https://www.postcard.inc/places/sopo-korean-eats-new-york-7HtTMM-v825",
        language="en", menu_items="kimbap",
        use_case="Value / portion comparison vs Koreatown",
        segment_guess="General reviewer (not identified)",
        notes="NOT VERBATIM. Search-engine summary; source aggregator pages (postcard.inc / wheree) blocked (403). Represents value/portion criticism. Reviewer/date/exact text not accessible.",
    ),
    dict(
        id="R14", platform="Aggregator search summary",
        review_date="", rating="",
        reviewer_name="",
        review_text=PARA + "Prices feel on the higher side, especially when adding extra sides, and seating is limited which can be inconvenient at busy hours.",
        review_url="",
        source_url="https://sopo-korean-eats.wheree.com/",
        language="en", menu_items="extra sides",
        use_case="Dine-in / value + environment criticism",
        segment_guess="General reviewer (not identified)",
        notes="NOT VERBATIM. Search-engine summary; source pages blocked (403). Represents price + limited-seating criticism. Reviewer/date/exact text not accessible.",
    ),
]

RAW_COLS = ["platform", "business_name", "location", "address", "review_date",
            "rating", "reviewer_name", "review_text", "review_url", "source_url",
            "language", "menu_items_mentioned", "use_case",
            "customer_segment_guess", "notes"]


def raw_row(r):
    return [
        r["platform"], BIZ, LOC, ADDR, r["review_date"], r["rating"],
        r["reviewer_name"], r["review_text"], r["review_url"], r["source_url"],
        r["language"], r["menu_items"], r["use_case"], r["segment_guess"],
        r["notes"],
    ]


# ---------------------------------------------------------------------------
# THEME TAGS  (1/0 flags per review + short reason)
# ---------------------------------------------------------------------------
THEMES = ["taste_flavor", "authenticity_korean_comfort", "value_for_money",
          "portion_size", "speed_convenience", "health_freshness",
          "menu_clarity_customization", "service_experience",
          "ambience_store_environment", "repeat_visit_intent",
          "catering_group_order"]

# flags keyed by review id: tuple of 11 ints in THEMES order, then reason
theme_data = {
    "R01": ((1,1,0,0,1,0,0,0,0,1,0),
            "Praises taste ('sooo good'), notes fast-casual convenience near Herald Sq, and explicit repeat intent ('cant wait to go back')."),
    "R02": ((1,1,1,0,1,0,1,0,0,1,0),
            "Korean comfort food + banchan/kimbap (authenticity, menu), 'impressed' (taste), 'fast food option' (convenience), will return (repeat)."),
    "R03": ((1,1,1,0,1,0,1,0,0,0,0),
            "High-quality affordable Korean (taste/value), banchan-style customization (menu), fast casual (convenience); promotional tone."),
    "R04": ((0,1,0,0,1,0,0,0,0,0,0),
            "Only partial caption available: signals Korean fast-casual concept (authenticity/convenience); other themes indeterminate."),
    "R05": ((1,0,0,0,0,0,0,0,0,0,0),
            "Title indicates a 'meal review' of the food (taste); too little text to flag other themes."),
    "R06": ((0,1,0,0,1,0,0,0,0,0,0),
            "Partial caption: identifies it as a new fast-casual Korean spot (authenticity/convenience); rest truncated."),
    "R07": ((0,1,0,0,1,0,0,0,0,0,0),
            "Partial caption: 'fast casual Korean' (authenticity/convenience); rest truncated."),
    "R08": ((1,1,0,0,0,0,0,0,0,0,0),
            "'bulgogi perfection... saucy dumplings... does NOT miss' = strong taste praise on Korean dishes; provenance uncertain."),
    "R09": ((1,0,0,0,0,1,0,0,0,1,0),
            "'super flavorful, fresh, outstanding' (taste + freshness) and 'every single time' (repeat). Paraphrased."),
    "R10": ((1,0,0,0,0,0,0,0,0,0,0),
            "Specific dish praise (pumpkin mash, dumplings) = taste. Paraphrased."),
    "R11": ((1,0,0,1,0,0,0,0,0,0,0),
            "Bulgogi 'not flavorful', 'barely any marinade', poor meat quality = negative taste + quality/portion concern. Paraphrased."),
    "R12": ((1,1,1,0,1,0,0,1,0,0,0),
            "Flavor + value/pricing praise, Korean fast-casual, friendly staff explaining dishes (service). Paraphrased."),
    "R13": ((1,0,1,1,0,0,0,0,0,0,0),
            "Kimbap 'really good' (taste) but 'small for the price' vs K-town (value + portion). Paraphrased."),
    "R14": ((0,0,1,0,0,0,0,0,1,0,0),
            "Higher prices esp. with extra sides (value) and limited seating (ambience/environment). Paraphrased."),
}

THEME_COLS = ["review_id", "platform", "reviewer_name", "review_excerpt"] + THEMES + ["reason"]


def theme_row(r):
    flags, reason = theme_data[r["id"]]
    excerpt = r["review_text"]
    if len(excerpt) > 90:
        excerpt = excerpt[:87] + "..."
    return [r["id"], r["platform"], r["reviewer_name"], excerpt] + list(flags) + [reason]


# ---------------------------------------------------------------------------
# SOURCE LOG
# ---------------------------------------------------------------------------
SRC_COLS = ["source", "url", "date_accessed", "reviews_collected",
            "access_method", "limitations", "notes"]

source_log = [
    ["Google Reviews / Google Maps",
     "https://www.google.com/maps/search/SOPO+Korean+Eats+463+7th+Ave",
     ACCESSED, "0 verbatim (1 paraphrased summary captured: R12)",
     "WebSearch only (WebFetch blocked: HTTP 403)",
     "Could not open Google Maps listing or read individual reviews. A widely-cited figure of '4.7 stars / 751 reviews' appeared in one search summary but the platform was NOT confirmed as Google and could not be verified on-page.",
     "Recommend manual capture from Google Maps for exact ratings, names, dates, and verbatim text."],
    ["Yelp", "https://www.yelp.com/biz/sopo-new-york-3", ACCESSED,
     "0 verbatim (3 paraphrased summaries: R09, R10, R11)",
     "WebSearch only (WebFetch blocked: HTTP 403)",
     "Yelp returned HTTP 403 to automated fetch. Listed as 'Updated June 2026 - 147 Photos & 56 Reviews'; overall rating reported as 4.5 stars in search summaries. Individual review text/names/dates NOT accessible without manual browsing.",
     "Aggregate: ~4.5 stars, ~56 reviews, 147 photos (per search index, 2026-06-08)."],
    ["TripAdvisor",
     "https://www.tripadvisor.com/Restaurant_Review-g60763-d33056321-Reviews-Sopo_Korean_Eats-New_York_City_New_York.html",
     ACCESSED, "0",
     "WebSearch only (WebFetch blocked: HTTP 403)",
     "Page blocked (403). Search summary indicated a ranking of roughly #6,444 of ~9,030 NYC restaurants; star rating and individual reviews not accessible.",
     "Not in the requested source list but found during search; logged for completeness."],
    ["DoorDash", "https://www.doordash.com/store/sopo-korean-eats-new-york-30888727/",
     ACCESSED, "0",
     "WebSearch only (WebFetch blocked: HTTP 403)",
     "Store page blocked (403); delivery ratings/reviews typically require app context/location. Star rating and review count NOT accessible.",
     "Confirmed SOPO is listed/orderable on DoorDash."],
    ["Uber Eats", "https://www.ubereats.com/ (SOPO Korean Eats, 463 7th Ave)",
     ACCESSED, "0",
     "WebSearch (no usable result) / WebFetch blocked",
     "No accessible Uber Eats listing surfaced; ratings/reviews not accessible.",
     "SOPO references multiple delivery channels; Uber Eats page not retrievable in this environment."],
    ["Grubhub", "https://www.grubhub.com/restaurant/sopo-korean-eats-463-7th-ave-new-york/8994688",
     ACCESSED, "0",
     "WebSearch only (WebFetch blocked: HTTP 403)",
     "Listing exists; star rating and individual reviews not accessible.",
     "Also present on Seamless (same platform): https://www.seamless.com/menu/sopo-korean-eats-463-7th-ave-new-york/8994688"],
    ["Toast / direct ordering", "https://eatsopo.com/ (official site)",
     ACCESSED, "0",
     "WebSearch only (WebFetch blocked: HTTP 403)",
     "Official site and ordering pages returned 403 to automated fetch. No public customer reviews/testimonials confirmed on-site.",
     "Site confirms concept, menu, chef (Dennis Hong), catering; no review data exposed."],
    ["TikTok", "https://www.tiktok.com/discover/sopo-nyc", ACCESSED,
     "5 captions captured (R01, R03, R04, R05, R08); 1 partial",
     "WebSearch (captions from search-indexed titles); WebFetch blocked (403)",
     "Full captions for some videos truncated in the search index; comment threads, like/comment counts, and exact post dates not accessible. One quote (R08) could not be tied to a confirmed video URL.",
     "Multiple genuine public TikTok posts about SOPO exist; creators include @tiffany.elizabeth, @yourbrooklynguide, @jonathanchoione, @storytimenyc, @codypremer."],
    ["Instagram", "https://www.instagram.com/eatsopo/", ACCESSED,
     "2 captions captured (R06, R07; both truncated)",
     "WebSearch (captions from search-indexed titles); WebFetch blocked (403)",
     "Caption text truncated in search index; full captions and comment threads not accessible without login. Business account is @eatsopo.",
     "Public creator posts found: @girlswholiketoeatfood (missy & jessie), MealPal. Did not access any login-walled content."],
    ["Reddit", "https://www.reddit.com/ (site search)", ACCESSED, "0",
     "WebSearch (site:reddit.com)",
     "No Reddit threads mentioning SOPO Korean Eats surfaced in search. Absence of results is not proof of absence of mentions.",
     "Recommend a manual Reddit search (r/FoodNYC, r/AskNYC, r/Korean) for any threads."],
    ["Food blogs / local press / editorial",
     "https://www.theinfatuation.com/new-york/reviews/sopo",
     ACCESSED, "0 (existence confirmed, text not captured)",
     "WebSearch only (WebFetch blocked: HTTP 403)",
     "The Infatuation has an editorial review page (blocked, 403). Other coverage found: amsterdamnews.com (2026-05-14), whatnow.com/whatnowny.com, New York Post (referenced). Editorial review text not accessible.",
     "These are press/editorial, not customer reviews; listed for context."],
    ["ezCater (catering)", "https://www.ezcater.com/catering/sopo-korean-eats-3",
     ACCESSED, "0 verbatim (aggregate rating noted)",
     "WebSearch only (WebFetch blocked: HTTP 403)",
     "Page blocked (403). Search summary reported a catering rating of ~4.9 stars from ~41 reviews; individual review text/names/dates not accessible.",
     "Relevant to catering/group-order use case. Aggregate only."],
    ["TOOL/ENVIRONMENT LIMITATION (global)", "n/a", ACCESSED, "n/a",
     "WebSearch = available; WebFetch = blocked",
     "In this execution environment WebFetch returned HTTP 403 for EVERY URL attempted (review platforms, news sites, and the business's own website). No login walls, CAPTCHAs, paywalls, or robots restrictions were bypassed. Where pages were blocked, the limitation was recorded instead of forcing access.",
     "Net effect: verbatim, per-reviewer data (names/dates/exact text) from Google/Yelp/TripAdvisor/delivery apps could not be collected here. Captured material is (a) verbatim public social captions and (b) clearly-flagged paraphrased search summaries + aggregate ratings."],
]

# ---------------------------------------------------------------------------
# BUILD WORKBOOK
# ---------------------------------------------------------------------------
wb = Workbook()

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
WRAP = Alignment(wrap_text=True, vertical="top")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def style_sheet(ws, headers, widths):
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = BORDER
    for c, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 30


def write_rows(ws, rows):
    for row in rows:
        ws.append(row)
    for r in range(2, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            cell.alignment = WRAP
            cell.border = BORDER


# Sheet 1: raw_reviews
ws1 = wb.active
ws1.title = "raw_reviews"
style_sheet(ws1, RAW_COLS,
            [14, 16, 22, 24, 22, 8, 24, 70, 42, 42, 9, 22, 26, 28, 60])
write_rows(ws1, [raw_row(r) for r in raw])

# Sheet 2: cleaned_reviews (no exact dupes present; mirror of raw, originals kept)
ws2 = wb.create_sheet("cleaned_reviews")
style_sheet(ws2, RAW_COLS,
            [14, 16, 22, 24, 22, 8, 24, 70, 42, 42, 9, 22, 26, 28, 60])
write_rows(ws2, [raw_row(r) for r in raw])  # 0 exact duplicates removed; see notes

# Sheet 3: theme_tags
ws3 = wb.create_sheet("theme_tags")
style_sheet(ws3, THEME_COLS,
            [10, 16, 24, 50] + [11] * len(THEMES) + [70])
write_rows(ws3, [theme_row(r) for r in raw])

# Sheet 4: source_log
ws4 = wb.create_sheet("source_log")
style_sheet(ws4, SRC_COLS, [30, 46, 14, 30, 30, 60, 50])
write_rows(ws4, source_log)

OUT = "/home/user/wherefore/SOPO_review_raw_analysis.xlsx"
wb.save(OUT)
print("Saved", OUT)
print("raw_reviews rows:", len(raw))
print("cleaned_reviews rows:", len(raw), "(0 exact duplicates found)")
print("theme_tags rows:", len(raw))
print("source_log rows:", len(source_log))
