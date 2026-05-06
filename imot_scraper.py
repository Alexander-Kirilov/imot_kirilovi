import pandas as pd
import re
import logging
import time
import random
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from openpyxl import load_workbook
from openpyxl.styles import Font
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import os
import sys
from pathlib import Path

# ================= GitHub Actions Environment =================
if os.getenv("GITHUB_ACTIONS"):
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = "/tmp/ms-playwright"
    print("Running in GitHub Actions environment")

# ================= LOGGING SETUP =================
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s | %(levelname)-7s | %(message)s')
file_handler = logging.FileHandler('imot_scraper.log', encoding='utf-8')
file_handler.setLevel(logging.INFO)
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)
console_handler.setFormatter(formatter)
logger.addHandler(console_handler)

logger.info("=== Starting imot.bg scraper ===")

# ================= PATHS & CONFIG =================
HISTORY_FILE = "all_listings_history.parquet"
HTML_OUTPUT  = Path("docs/index.html")   # GitHub Pages serves from /docs
HTML_OUTPUT.parent.mkdir(exist_ok=True)

TODAY     = datetime.now().strftime("%Y-%m-%d")
NOW_STR   = datetime.now().strftime("%d.%m.%Y %H:%M")
timestamp = datetime.now().strftime("%Y%m%d_%H%M")
excel_file = f"imot_bg_scraping_{timestamp}.xlsx"

# ── Email ─────────────────────────────────────────────────────────────────────
SMTP_SERVER     = "smtp.gmail.com"
SMTP_PORT       = 587
SENDER_EMAIL    = "a.kirilov74@gmail.com"
SENDER_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "vawb lshb tqrw azrf")
RECEIVERS       = ["a.kirilov74@gmail.com", "hristina.padeva@gmail.com"]

# ── Search URL ────────────────────────────────────────────────────────────────
base_url = (
    'https://www.imot.bg/obiavi/prodazhbi/grad-sofiya/darvenitsa/tristaen'
    '?type_home=4~5~&kv_min=85&price_max=280000&raioni=44~45~46~47~48~49~&ybuild_type=1~&fe1=1'
)

listings = []

# ── Column constants ──────────────────────────────────────────────────────────
COL_LINK            = 'Link'
COL_PRICE           = 'Price_EUR'
COL_SIZE            = 'Size_sqm'
COL_PRICE_PER_SQM   = 'Price_EUR_per_sqm'
COL_LOCATION        = 'Location'
COL_FLOOR           = 'Floor'
COL_TOTAL_FLOORS    = 'Total_floors'
COL_YEAR            = 'Year_built'
COL_INFO            = 'Info'
COL_TITLE           = 'Title'
COL_SCRAPED_DATE    = 'Scraped_Date'
COL_FIRST_SEEN      = 'First_Seen_Date'   # ← NEW: date when listing first appeared
COL_PRICE_HISTORY   = 'Price_History'
COL_SOLD            = 'Sold'
COL_SITE_PRICE_HISTORY = 'Site price history'


# ================= HELPERS =================

def clean_price(raw):
    if not raw: return None
    m = re.search(r'([\d\s.,]+)\s*€', raw)
    if not m: return None
    cleaned = m.group(1).replace(' ', '').replace(',', '.')
    try:
        return float(cleaned)
    except ValueError:
        logger.warning(f"Failed to parse price: {raw}")
        return None


def normalize_location(loc):
    if not loc: return loc
    rules = {
        r'Младост\s*IV-ти': 'Младост 4',
        r'Младост\s*IV':    'Младост 4',
        r'Младост\s*V':     'Младост 5',
        r'Младост\s*III':   'Младост 3',
        r'Младост\s*II':    'Младост 2',
        r'Младост\s*I\b':   'Младост 1',
    }
    orig = loc
    for pat, repl in rules.items():
        loc = re.sub(pat, repl, loc, flags=re.IGNORECASE)
    if orig != loc:
        logger.debug(f"Normalized location: {orig} → {loc}")
    return loc


def parse_total_ads(html):
    soup = BeautifulSoup(html, 'html.parser')
    text = soup.get_text(separator=' ', strip=True)
    m = re.search(r'(?:от\s*общо\s*|от\s*)(\d+)\s*обяв[иа]', text, re.IGNORECASE)
    if m: return int(m.group(1))
    m2 = re.search(r'(\d+)\s*[-–]\s*\d+\s*от\s*общо\s*(\d+)', text)
    if m2: return int(m2.group(2))
    return None


def parse_page(html, page_num=None):
    soup  = BeautifulSoup(html, 'html.parser')
    items = soup.find_all('div', class_='item')
    if not items:
        logger.warning(f"No items found on page {page_num or 1}")
        return 0

    count = 0
    for item in items:
        try:
            title_a = item.find('a', class_='title')
            if not title_a: continue

            title = title_a.get_text(separator=' ', strip=True)
            href  = title_a['href'].strip()
            if 'fakti.bg' in href.lower() or not href.startswith(('https://www.imot.bg', '//', '/')):
                continue
            if href.startswith('//'): href = 'https:' + href
            elif href.startswith('/'): href = 'https://www.imot.bg' + href
            elif not href.startswith('http'): href = 'https://www.imot.bg/' + href

            price_div = item.find('div', class_='price')
            price_raw = price_div.get_text(strip=True) if price_div else ''
            price_eur = clean_price(price_raw)

            info_div  = item.find('div', class_='info')
            info_text = info_div.get_text(strip=True) if info_div else ''

            size = None
            m = re.search(r'(\d{2,3})\s*кв\.?\s*м', info_text)
            if m: size = int(m.group(1))

            floor = total_floors = None
            m_floor = re.search(r'(\d+)\s*[-–]\s*(?:ти|ри)?', info_text)
            if m_floor: floor = int(m_floor.group(1))
            m_total = re.search(r'(?:от\s*|/\s*)(\d+)', info_text)
            if m_total: total_floors = int(m_total.group(1))

            year = None
            m_year = re.search(r'(19\d{2}|20\d{2})(?:\s*[-–]\s*(19\d{2}|20\d{2}))?', info_text)
            if m_year: year = m_year.group(0).strip()

            location = None
            if 'град София,' in title:
                location = title.split('град София,')[-1].strip()
                location = normalize_location(location)

            price_m2 = round(price_eur / size, 2) if price_eur and size and size > 0 else None

            listings.append({
                COL_TITLE:              title,
                COL_LOCATION:           location,
                COL_PRICE:              price_eur,
                COL_SIZE:               size,
                COL_PRICE_PER_SQM:      price_m2,
                COL_FLOOR:              floor,
                COL_TOTAL_FLOORS:       total_floors,
                COL_YEAR:               year,
                COL_INFO:               info_text,
                COL_LINK:               href,
                COL_SCRAPED_DATE:       TODAY,
                COL_FIRST_SEEN:         TODAY,   # will be overridden for existing
                COL_PRICE_HISTORY:      "",
                COL_SITE_PRICE_HISTORY: "",
            })
            count += 1
        except Exception as e:
            logger.warning(f"Error parsing item on page {page_num or 1}: {e}")

    logger.info(f"Scraped {count} listings from page {page_num or 1}")
    return count


def format_price_history_entry(price, date_str):
    if pd.isna(price) or price is None or pd.isna(date_str):
        return ""
    try:
        price_int = int(round(float(price)))
        return f"{price_int:,} € ({date_str})"
    except (ValueError, TypeError):
        return ""


def append_current_if_needed(hist, price, date):
    if pd.isna(price) or pd.isna(date):
        return hist if isinstance(hist, str) else ""
    curr = format_price_history_entry(price, date)
    if not isinstance(hist, str):
        hist = ""
    if curr and curr not in hist:
        return f"{hist} → {curr}" if hist.strip() else curr
    return hist


def deduplicate_history(df, link_col, price_history_col):
    def merge_history(series):
        vals   = [v.strip() for v in series if isinstance(v, str) and v.strip()]
        seen   = set()
        result = []
        for v in vals:
            if v not in seen:
                seen.add(v)
                result.append(v)
        return " → ".join(result) if result else ""

    df[price_history_col] = df[price_history_col].fillna("").astype(str)
    agg = {}
    for c in df.columns:
        agg[c] = merge_history if c == price_history_col else 'last'
    return df.groupby(link_col, as_index=False).agg(agg)


def parse_site_price_history_html(html):
    soup      = BeautifulSoup(html, 'html.parser')
    container = soup.find('div', id='priceHistory2')
    node      = None
    if container:
        node = container.find('statistiki')
    if not node:
        node = soup.find('statistiki')
    if not node:
        return ""

    divs = node.find_all('div', recursive=False)
    divs = [d for d in divs if d.get_text(strip=True)]
    if len(divs) < 4:
        return ""

    def parse_bg_datetime(s):
        m = re.search(r"(\d{2})\.(\d{2})\.(\d{4}).*?(\d{2})\.(\d{2})", s)
        if not m: return None
        day, month, year, hour, minute = map(int, m.groups())
        try: return datetime(year, month, day, hour, minute)
        except ValueError: return None

    def parse_change_number(s):
        if not s: return None
        s_clean = s.replace("\xa0", " ")
        m = re.search(r"([-+−]?\s*\d[\d\s]*)", s_clean)
        if not m: return None
        try:
            raw = m.group(1).replace(" ", "").replace("−", "-")
            return int(raw)
        except ValueError: return None

    def parse_price_number(s):
        if not s: return None
        m = re.search(r"([\d\s.,]+)", s)
        if not m: return None
        num = m.group(1).replace(" ", "").replace(",", ".")
        try: return float(num)
        except ValueError: return None

    def format_eur_int(value):
        try:
            v = int(round(float(value)))
            return f"{v:,}".replace(",", " ") + " €"
        except Exception: return ""

    rows = []
    initial_price_num = None
    i = 0
    while i + 1 < len(divs):
        group = divs[i: i + 4]
        texts = [g.get_text(" ", strip=True) for g in group]
        date_text = price_text = change_text = ""
        for t in texts:
            if not price_text and "€" in t: price_text = t
            if not date_text and ("начало" in t.lower() or re.search(r"\d{2}\.\d{2}\.\d{4}", t)):
                date_text = t
            if not change_text and any(ch in t for ch in ["+", "−", "-", "↑", "↓"]):
                change_text = t
        if price_text and date_text:
            if date_text.strip().lower().startswith("начало"):
                initial_price_num = parse_price_number(price_text)
            else:
                rows.append({"date": date_text, "change": change_text, "price_text": price_text})
        i += 4

    if not rows and initial_price_num is None:
        return ""

    for idx, r in enumerate(rows):
        rows[idx]["dt"]         = parse_bg_datetime(r["date"])
        rows[idx]["price_num"]  = parse_price_number(r["price_text"])
        rows[idx]["change_num"] = parse_change_number(r["change"])

    rows_sorted = sorted(rows, key=lambda r: r["dt"] if r["dt"] is not None else datetime.max)
    price_nums  = [r["price_num"] for r in rows_sorted if r["price_num"] is not None]
    if initial_price_num is None and price_nums:
        initial_price_num = max(price_nums)

    entries = []
    if initial_price_num is not None:
        base_str = format_eur_int(initial_price_num)
        if base_str:
            entries.append(f"Начална цена: {base_str.replace(' €', '')}")

    for r in rows_sorted:
        new_price_num = None
        if r["price_num"] is not None:
            new_price_num = r["price_num"] + r["change_num"] if r["change_num"] is not None else r["price_num"]
        price_str = format_eur_int(new_price_num) if new_price_num is not None else r["price_text"]
        if r["change_num"] is not None and r["change_num"] != 0:
            sign       = "-" if r["change_num"] < 0 else "+"
            change_abs = abs(r["change_num"])
            change_str = f"{change_abs:,}".replace(",", " ")
            entries.append(f"{r['date']} ({sign} {change_str}) - {price_str}")
        else:
            entries.append(f"{r['date']} - {price_str}")

    return " | ".join(entries)


def scrape_site_price_histories_selenium(links):
    histories = {url: "" for url in links}
    if not links: return histories

    total   = len(links)
    success = empty = failed = 0
    logger.info(f"Fetching site price history for {total} listings")

    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC

        options = Options()
        options.add_argument("--headless")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--window-size=1920,1080")
        options.add_argument(
            "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        )
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)
        driver = webdriver.Chrome(options=options)
    except Exception:
        logger.warning("Selenium init failed; skipping site price histories")
        return histories

    try:
        wait = WebDriverWait(driver, 12)
        for idx, url in enumerate(links, start=1):
            logger.info(f"Site price history [{idx}/{total}]: {url}")
            try:
                driver.get(url)
                price_history_elem = wait.until(
                    EC.presence_of_element_located((By.ID, "priceHistory2"))
                )
                try:
                    title_span = price_history_elem.find_element(
                        By.CSS_SELECTOR, 'div.title span[onclick*="showpricechange"]'
                    )
                    onclick_attr = title_span.get_attribute("onclick") or ""
                    start_i = onclick_attr.find("(") + 1
                    end_i   = onclick_attr.rfind(");")
                    params  = [p.strip().strip("'") for p in onclick_attr[start_i:end_i].split(",")]
                    if len(params) >= 5:
                        js_code = (
                            f"showpricechange('{params[0]}','{params[1]}',"
                            f"'{params[2]}','{params[3]}','{params[4]}');"
                        )
                        driver.execute_script(js_code)
                        try:
                            wait.until(EC.presence_of_element_located((By.TAG_NAME, "statistiki")))
                        except Exception:
                            pass
                        time.sleep(1)
                except Exception:
                    pass
                html = driver.page_source
                histories[url] = parse_site_price_history_html(html)
                if histories[url].strip(): success += 1
                else: empty += 1
            except Exception:
                failed += 1
                histories[url] = ""
        logger.info(f"Site price history: total={total}, ok={success}, empty={empty}, failed={failed}")
        return histories
    finally:
        driver.quit()


# ================= HTML GENERATOR =================

def _fmt_price(x):
    return f"{int(round(x)):,} €".replace(",", "\u202f") if pd.notna(x) and x else "—"

def _fmt_size(x):
    return f"{int(x)} m²" if pd.notna(x) and x else "—"

def _fmt_pm2(x):
    return f"{int(round(x))} €/m²" if pd.notna(x) and x else "—"

def _link_cell(url):
    if not url or pd.isna(url): return "—"
    return f'<a href="{url}" target="_blank" rel="noopener">🔗 Виж</a>'


def _build_rows(df, cols):
    """Build <tr> rows for given columns."""
    rows_html = []
    for _, row in df.iterrows():
        cells = []
        for c in cols:
            val = row.get(c, "")
            if c == COL_LINK:
                cells.append(f"<td>{_link_cell(val)}</td>")
            elif c == COL_PRICE:
                cells.append(f"<td>{_fmt_price(val)}</td>")
            elif c == COL_SIZE:
                cells.append(f"<td>{_fmt_size(val)}</td>")
            elif c == COL_PRICE_PER_SQM:
                cells.append(f"<td>{_fmt_pm2(val)}</td>")
            elif c == 'Price_EUR_old':
                cells.append(f"<td class='old-price'>{_fmt_price(val)}</td>")
            else:
                cells.append(f"<td>{val if pd.notna(val) and val != '' else '—'}</td>")
        rows_html.append("<tr>" + "".join(cells) + "</tr>")
    return "\n".join(rows_html)


def _table(df, cols, headers, css_id="", extra_class=""):
    if df.empty:
        return '<p class="empty-note">Няма данни.</p>'
    thead = "<tr>" + "".join(f"<th>{h}</th>" for h in headers) + "</tr>"
    tbody = _build_rows(df, cols)
    id_attr = f' id="{css_id}"' if css_id else ""
    cls     = f'data-table {extra_class}'.strip()
    return f'<table{id_attr} class="{cls}"><thead>{thead}</thead><tbody>{tbody}</tbody></table>'


def generate_html(df_all: pd.DataFrame, now_str: str):
    """Generate docs/index.html from the full history dataframe."""

    cutoff = (datetime.now() - timedelta(days=10)).strftime("%Y-%m-%d")

    # ── Derive sections ───────────────────────────────────────────────────────
    df_active = df_all[~df_all[COL_SOLD].fillna(False)].copy() if not df_all.empty else pd.DataFrame()

    # New in last 10 days: first_seen >= cutoff
    if COL_FIRST_SEEN in df_all.columns and not df_all.empty:
        df_new10 = df_active[df_active[COL_FIRST_SEEN].fillna("") >= cutoff].copy()
    else:
        df_new10 = df_active[df_active[COL_SCRAPED_DATE].fillna("") >= cutoff].copy()

    # Price changes in last 10 days: has " → " in price_history AND scraped recently
    if not df_active.empty and COL_PRICE_HISTORY in df_active.columns:
        mask_changed = (
            df_active[COL_PRICE_HISTORY].fillna("").str.contains(" → ") &
            (df_active[COL_SCRAPED_DATE].fillna("") >= cutoff)
        )
        df_changed10 = df_active[mask_changed].copy()
    else:
        df_changed10 = pd.DataFrame()

    # Sold: all sold listings
    df_sold = df_all[df_all[COL_SOLD].fillna(False)].copy() if not df_all.empty else pd.DataFrame()

    # Stats
    n_total   = len(df_active)
    n_new10   = len(df_new10)
    n_changed = len(df_changed10)
    n_sold    = len(df_sold)

    # Sort all active by scraped date desc
    if not df_active.empty:
        df_active = df_active.sort_values(COL_SCRAPED_DATE, ascending=False)
    if not df_new10.empty:
        df_new10  = df_new10.sort_values(COL_FIRST_SEEN if COL_FIRST_SEEN in df_new10.columns else COL_SCRAPED_DATE, ascending=False)
    if not df_changed10.empty:
        df_changed10 = df_changed10.sort_values(COL_SCRAPED_DATE, ascending=False)

    # ── Build table HTML ──────────────────────────────────────────────────────
    new_table = _table(
        df_new10,
        [COL_LOCATION, COL_PRICE, COL_SIZE, COL_PRICE_PER_SQM, COL_FLOOR, COL_TOTAL_FLOORS, COL_YEAR, COL_SITE_PRICE_HISTORY, COL_LINK],
        ["Локация", "Цена", "Площ", "€/m²", "Ет.", "Общо ет.", "Година", "История (сайт)", ""],
        css_id="new-table",
    )

    changed_table = _table(
        df_changed10,
        [COL_LOCATION, COL_PRICE, COL_SIZE, COL_PRICE_PER_SQM, COL_PRICE_HISTORY, COL_SITE_PRICE_HISTORY, COL_LINK],
        ["Локация", "Текуща цена", "Площ", "€/m²", "История на цената", "История (сайт)", ""],
        css_id="changed-table",
    )

    all_table = _table(
        df_active,
        [COL_LOCATION, COL_PRICE, COL_SIZE, COL_PRICE_PER_SQM, COL_FLOOR, COL_TOTAL_FLOORS, COL_YEAR, COL_SCRAPED_DATE, COL_PRICE_HISTORY, COL_LINK],
        ["Локация", "Цена", "Площ", "€/m²", "Ет.", "Общо ет.", "Год.", "Последно виждана", "История на цената", ""],
        css_id="all-table",
    )

    sold_table = _table(
        df_sold,
        [COL_LOCATION, COL_PRICE, COL_SIZE, COL_PRICE_PER_SQM, COL_SCRAPED_DATE, COL_PRICE_HISTORY, COL_LINK],
        ["Локация", "Последна цена", "Площ", "€/m²", "Последно виждана", "История на цената", ""],
        css_id="sold-table",
        extra_class="sold-table",
    )

    # ── Full HTML ─────────────────────────────────────────────────────────────
    html = f"""<!DOCTYPE html>
<html lang="bg">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Имоти Dashboard · Дървеница</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap" rel="stylesheet">
<style>
  :root {{
    --bg:        #0f1117;
    --surface:   #1a1d27;
    --border:    #2a2d3a;
    --text:      #e2e4f0;
    --muted:     #7a7d9a;
    --accent:    #4f9cf9;
    --green:     #3ecf8e;
    --orange:    #f59e0b;
    --red:       #f43f5e;
    --mono:      'IBM Plex Mono', monospace;
    --sans:      'IBM Plex Sans', sans-serif;
  }}

  * {{ box-sizing: border-box; margin: 0; padding: 0; }}

  body {{
    font-family: var(--sans);
    background: var(--bg);
    color: var(--text);
    font-size: 14px;
    line-height: 1.6;
  }}

  /* ── Header ── */
  header {{
    padding: 28px 32px 20px;
    border-bottom: 1px solid var(--border);
    display: flex;
    align-items: baseline;
    gap: 24px;
    flex-wrap: wrap;
  }}
  header h1 {{
    font-size: 22px;
    font-weight: 600;
    letter-spacing: -0.5px;
    color: var(--text);
  }}
  header h1 span {{ color: var(--accent); }}
  .updated {{
    font-family: var(--mono);
    font-size: 12px;
    color: var(--muted);
    margin-left: auto;
  }}

  /* ── Stats bar ── */
  .stats {{
    display: flex;
    gap: 1px;
    background: var(--border);
    border-bottom: 1px solid var(--border);
  }}
  .stat {{
    flex: 1;
    padding: 18px 24px;
    background: var(--surface);
    text-align: center;
  }}
  .stat-num {{
    font-family: var(--mono);
    font-size: 28px;
    font-weight: 600;
    line-height: 1;
    color: var(--accent);
  }}
  .stat-num.green  {{ color: var(--green); }}
  .stat-num.orange {{ color: var(--orange); }}
  .stat-num.red    {{ color: var(--red); }}
  .stat-label {{
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    color: var(--muted);
    margin-top: 4px;
  }}

  /* ── Nav tabs ── */
  nav {{
    padding: 0 32px;
    border-bottom: 1px solid var(--border);
    display: flex;
    gap: 0;
    overflow-x: auto;
  }}
  nav a {{
    display: inline-block;
    padding: 14px 20px;
    font-size: 13px;
    font-weight: 600;
    color: var(--muted);
    text-decoration: none;
    border-bottom: 2px solid transparent;
    white-space: nowrap;
    transition: color 0.15s, border-color 0.15s;
  }}
  nav a:hover  {{ color: var(--text); border-color: var(--border); }}
  nav a.active {{ color: var(--accent); border-color: var(--accent); }}

  /* ── Sections ── */
  main {{ padding: 0 32px 48px; }}
  section {{ padding-top: 36px; }}
  section h2 {{
    font-size: 15px;
    font-weight: 600;
    color: var(--text);
    margin-bottom: 4px;
    display: flex;
    align-items: center;
    gap: 10px;
  }}
  section h2 .badge {{
    font-family: var(--mono);
    font-size: 11px;
    padding: 1px 8px;
    border-radius: 99px;
    background: var(--border);
    color: var(--muted);
  }}
  .section-desc {{
    font-size: 12px;
    color: var(--muted);
    margin-bottom: 16px;
  }}

  /* ── Search bar ── */
  .search-wrap {{
    margin-bottom: 12px;
  }}
  .search-wrap input {{
    width: 320px;
    max-width: 100%;
    padding: 8px 14px;
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 6px;
    color: var(--text);
    font-family: var(--sans);
    font-size: 13px;
    outline: none;
    transition: border-color 0.15s;
  }}
  .search-wrap input:focus {{ border-color: var(--accent); }}
  .search-wrap input::placeholder {{ color: var(--muted); }}

  /* ── Tables ── */
  .table-wrap {{
    overflow-x: auto;
    border: 1px solid var(--border);
    border-radius: 8px;
  }}
  table.data-table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
  }}
  table.data-table thead {{
    position: sticky;
    top: 0;
    z-index: 1;
  }}
  table.data-table th {{
    background: var(--surface);
    color: var(--muted);
    font-weight: 600;
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    padding: 10px 14px;
    border-bottom: 1px solid var(--border);
    white-space: nowrap;
    cursor: pointer;
    user-select: none;
  }}
  table.data-table th:hover {{ color: var(--text); }}
  table.data-table th.sorted-asc::after  {{ content: ' ↑'; }}
  table.data-table th.sorted-desc::after {{ content: ' ↓'; }}

  table.data-table td {{
    padding: 10px 14px;
    border-bottom: 1px solid var(--border);
    vertical-align: top;
    max-width: 280px;
    word-break: break-word;
  }}
  table.data-table tr:last-child td {{ border-bottom: none; }}
  table.data-table tr:hover td {{ background: #1e2130; }}

  table.data-table td.old-price {{
    color: var(--muted);
    text-decoration: line-through;
  }}

  table.sold-table td {{ color: var(--muted); }}
  table.sold-table td a {{ color: var(--muted); }}

  table.data-table a {{
    color: var(--accent);
    text-decoration: none;
    font-weight: 600;
  }}
  table.data-table a:hover {{ text-decoration: underline; }}

  /* price history cell — smaller mono */
  table.data-table td:has(.history) {{ font-size: 11px; }}
  .history {{
    font-family: var(--mono);
    font-size: 11px;
    color: var(--muted);
  }}

  .empty-note {{
    color: var(--muted);
    font-size: 13px;
    padding: 20px 0;
  }}

  /* ── Footer ── */
  footer {{
    text-align: center;
    padding: 24px;
    font-size: 12px;
    color: var(--muted);
    border-top: 1px solid var(--border);
  }}

  @media (max-width: 640px) {{
    header {{ padding: 16px; }}
    main   {{ padding: 0 16px 32px; }}
    nav    {{ padding: 0 16px; }}
    .stats {{ flex-wrap: wrap; }}
    .stat  {{ flex: 1 1 50%; }}
    .stat-num {{ font-size: 22px; }}
  }}
</style>
</head>
<body>

<header>
  <h1>🏠 Имоти · <span>Дървеница</span></h1>
  <span class="updated">Обновено: {now_str}</span>
</header>

<div class="stats">
  <div class="stat">
    <div class="stat-num">{n_total}</div>
    <div class="stat-label">Активни обяви</div>
  </div>
  <div class="stat">
    <div class="stat-num green">{n_new10}</div>
    <div class="stat-label">Нови (10 дни)</div>
  </div>
  <div class="stat">
    <div class="stat-num orange">{n_changed}</div>
    <div class="stat-label">Промени в цена (10 дни)</div>
  </div>
  <div class="stat">
    <div class="stat-num red">{n_sold}</div>
    <div class="stat-label">Продадени / свалени</div>
  </div>
</div>

<nav>
  <a href="#new"     class="active">Нови (10 дни)</a>
  <a href="#changed">Промени (10 дни)</a>
  <a href="#all">Всички активни</a>
  <a href="#sold">Продадени</a>
</nav>

<main>

  <section id="new">
    <h2>Нови обяви <span class="badge">{n_new10}</span></h2>
    <p class="section-desc">Обяви, открити за първи път в последните 10 дни.</p>
    <div class="table-wrap">
      {new_table}
    </div>
  </section>

  <section id="changed">
    <h2>Промени в цена <span class="badge">{n_changed}</span></h2>
    <p class="section-desc">Активни обяви с регистрирана промяна на цената в последните 10 дни.</p>
    <div class="table-wrap">
      {changed_table}
    </div>
  </section>

  <section id="all">
    <h2>Всички активни обяви <span class="badge">{n_total}</span></h2>
    <p class="section-desc">Пълен списък на текущо активните обяви.</p>
    <div class="search-wrap">
      <input type="text" id="all-search" placeholder="Търси по локация, цена, площ…" oninput="filterTable('all-table', this.value)">
    </div>
    <div class="table-wrap">
      {all_table}
    </div>
  </section>

  <section id="sold">
    <h2>Продадени / свалени <span class="badge">{n_sold}</span></h2>
    <p class="section-desc">Обяви, изчезнали от сайта (вероятно продадени или свалени).</p>
    <div class="table-wrap">
      {sold_table}
    </div>
  </section>

</main>

<footer>
  Данни от imot.bg · Обновява се автоматично в 07:00 и 13:00 ч. · {now_str}
</footer>

<script>
// ── Simple table search ──────────────────────────────────────────────────────
function filterTable(tableId, query) {{
  const tbl  = document.getElementById(tableId);
  if (!tbl) return;
  const rows = tbl.querySelectorAll('tbody tr');
  const q    = query.trim().toLowerCase();
  rows.forEach(r => {{
    r.style.display = q === '' || r.textContent.toLowerCase().includes(q) ? '' : 'none';
  }});
}}

// ── Sortable columns ─────────────────────────────────────────────────────────
document.querySelectorAll('table.data-table th').forEach((th, colIdx) => {{
  th.addEventListener('click', () => {{
    const table = th.closest('table');
    const tbody = table.querySelector('tbody');
    const rows  = Array.from(tbody.querySelectorAll('tr'));
    const asc   = th.classList.contains('sorted-asc');

    table.querySelectorAll('th').forEach(t => t.classList.remove('sorted-asc','sorted-desc'));
    th.classList.add(asc ? 'sorted-desc' : 'sorted-asc');

    rows.sort((a, b) => {{
      const va = a.children[colIdx]?.textContent.trim() || '';
      const vb = b.children[colIdx]?.textContent.trim() || '';
      const na = parseFloat(va.replace(/[^\\d.]/g,''));
      const nb = parseFloat(vb.replace(/[^\\d.]/g,''));
      const cmp = !isNaN(na) && !isNaN(nb)
        ? na - nb
        : va.localeCompare(vb, 'bg');
      return asc ? -cmp : cmp;
    }});
    rows.forEach(r => tbody.appendChild(r));
  }});
}});

// ── Nav highlight on scroll ──────────────────────────────────────────────────
const sections = document.querySelectorAll('main section[id]');
const navLinks = document.querySelectorAll('nav a');
const observer = new IntersectionObserver(entries => {{
  entries.forEach(e => {{
    if (e.isIntersecting) {{
      navLinks.forEach(a => a.classList.remove('active'));
      const link = document.querySelector('nav a[href="#' + e.target.id + '"]');
      if (link) link.classList.add('active');
    }}
  }});
}}, {{ threshold: 0.25 }});
sections.forEach(s => observer.observe(s));
</script>

</body>
</html>"""

    HTML_OUTPUT.write_text(html, encoding="utf-8")
    logger.info(f"HTML dashboard written to {HTML_OUTPUT}")


# ================= SCRAPING =================
with sync_playwright() as p:
    logger.info("Opening browser (headless)…")
    browser = p.chromium.launch(headless=True)
    page    = browser.new_page()
    page.set_extra_http_headers({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    })
    total_ads      = None
    items_per_page = 40

    try:
        logger.info(f"Loading first page: {base_url}")
        page.goto(base_url, wait_until='domcontentloaded', timeout=60000)
        page.wait_for_selector('div.item', timeout=25000)

        html      = page.content()
        total_ads = parse_total_ads(html)
        if total_ads: logger.info(f"Total listings: {total_ads}")
        else:         logger.warning("Could not extract total listings count")

        parse_page(html, page_num=1)

        max_pages = (
            (total_ads // items_per_page) + (1 if total_ads % items_per_page else 0)
            if total_ads else 5
        )
        logger.info(f"Planning to scrape up to {max_pages} pages")

        page_num = 2
        while page_num <= max_pages + 2:
            try:
                url = (
                    base_url.replace('?', f'/p-{page_num}?')
                    if '?' in base_url else f"{base_url}/p-{page_num}"
                )
                logger.info(f"Loading page {page_num}: {url}")
                page.goto(url, wait_until='domcontentloaded', timeout=45000)
                try:
                    page.wait_for_selector('div.item', timeout=12000)
                except PlaywrightTimeoutError:
                    logger.info(f"No more listings on page {page_num}")
                    break
                count = parse_page(page.content(), page_num=page_num)
                if count == 0:
                    logger.info(f"Empty page {page_num} → stopping")
                    break
                time.sleep(2.8 + random.uniform(0, 1.8))
            except Exception as e:
                logger.error(f"Error on page {page_num}: {e}")
                break
            page_num += 1

    except Exception as e:
        logger.critical(f"Critical scraping error: {e}")
    finally:
        browser.close()
        logger.info("Browser closed")


# ================= PROCESSING =================
if not listings:
    logger.warning("No listings scraped → exiting")
    exit()

df_new      = pd.DataFrame(listings)
df_sold_now = pd.DataFrame()

site_histories = scrape_site_price_histories_selenium(df_new[COL_LINK].tolist())
df_new[COL_SITE_PRICE_HISTORY] = df_new[COL_LINK].map(site_histories)

# Load history
df_history = pd.DataFrame()
if os.path.exists(HISTORY_FILE):
    try:
        df_history = pd.read_parquet(HISTORY_FILE)
        logger.info(f"Loaded history: {len(df_history)} rows")
    except Exception as e:
        logger.error(f"Error reading parquet: {e}")

# Ensure all columns exist in history
if not df_history.empty:
    df_history = deduplicate_history(df_history, COL_LINK, COL_PRICE_HISTORY)
    df_history[COL_PRICE_HISTORY].fillna("", inplace=True)
    for col, default in [
        (COL_SITE_PRICE_HISTORY, ""),
        (COL_SOLD, False),
        (COL_FIRST_SEEN, ""),
    ]:
        if col not in df_history.columns:
            df_history[col] = default
    df_history[COL_SOLD].fillna(False, inplace=True)

df_all = df_history.copy()
if not df_all.empty:
    df_all = df_all.set_index(COL_LINK, drop=False)
    for col, default in [
        (COL_SOLD, False),
        (COL_SITE_PRICE_HISTORY, ""),
        (COL_FIRST_SEEN, ""),
    ]:
        if col not in df_all.columns:
            df_all[col] = default

# New listings
df_new_only = (
    df_new[~df_new[COL_LINK].isin(df_all[COL_LINK])]
    if not df_all.empty else df_new.copy()
)

# Changed prices
df_changed = pd.DataFrame()
if not df_history.empty:
    merged = df_new.merge(
        df_history[[COL_LINK, COL_PRICE, COL_LOCATION, COL_SIZE, COL_PRICE_HISTORY]],
        on=COL_LINK,
        how='inner',
        suffixes=('_new', '_old')
    )
    changed_mask = merged[f'{COL_PRICE}_new'] != merged[f'{COL_PRICE}_old']
    changed = merged[changed_mask].copy()

    if not changed.empty:
        def update_history(row):
            hist      = row[f'{COL_PRICE_HISTORY}_old']
            old_entry = format_price_history_entry(row[f'{COL_PRICE}_old'], "before")
            if old_entry and old_entry not in hist:
                return f"{hist} → {old_entry}" if hist.strip() else old_entry
            return hist

        changed[f'{COL_PRICE_HISTORY}_updated'] = changed.apply(update_history, axis=1)
        changed[f'{COL_PRICE_PER_SQM}_new'] = (
            round(changed[f'{COL_PRICE}_new'] / changed[f'{COL_SIZE}_new'], 2)
            if f'{COL_SIZE}_new' in changed.columns else None
        )
        df_changed = changed[[
            f'{COL_LOCATION}_old',
            f'{COL_PRICE}_old',
            f'{COL_PRICE}_new',
            f'{COL_SIZE}_new',
            f'{COL_PRICE_PER_SQM}_new',
            COL_LINK,
            COL_SITE_PRICE_HISTORY,
            f'{COL_PRICE_HISTORY}_updated',
        ]].rename(columns={
            f'{COL_LOCATION}_old':          COL_LOCATION,
            f'{COL_SIZE}_new':              COL_SIZE,
            f'{COL_PRICE_HISTORY}_updated': COL_PRICE_HISTORY,
        })
        logger.info(f"Found {len(df_changed)} price changes")

# Update df_all
for _, row in df_new.iterrows():
    link     = row[COL_LINK]
    row_dict = row.to_dict()

    if link in df_all.index:
        old_price        = df_all.at[link, COL_PRICE]
        old_scraped_date = df_all.at[link, COL_SCRAPED_DATE]
        new_price        = row_dict.get(COL_PRICE)
        new_scraped_date = TODAY

        # Preserve first_seen date
        row_dict[COL_FIRST_SEEN] = df_all.at[link, COL_FIRST_SEEN] or old_scraped_date

        for col in df_all.columns:
            if col in row_dict and col != COL_PRICE_HISTORY and col != COL_FIRST_SEEN:
                df_all.at[link, col] = row_dict[col]
        if COL_SITE_PRICE_HISTORY in row_dict:
            df_all.at[link, COL_SITE_PRICE_HISTORY] = row_dict[COL_SITE_PRICE_HISTORY]

        if pd.notna(new_price) and pd.notna(old_price) and old_price != new_price:
            current_hist = df_all.at[link, COL_PRICE_HISTORY] or ""
            old_entry    = format_price_history_entry(old_price, old_scraped_date)
            if old_entry and old_entry not in current_hist:
                current_hist = f"{current_hist} → {old_entry}" if current_hist.strip() else old_entry
            new_entry = format_price_history_entry(new_price, new_scraped_date)
            if new_entry and new_entry not in current_hist:
                current_hist = f"{current_hist} → {new_entry}" if current_hist.strip() else new_entry
            df_all.at[link, COL_PRICE_HISTORY] = current_hist
    else:
        # new listing
        if pd.notna(row_dict.get(COL_PRICE)):
            row_dict[COL_PRICE_HISTORY] = format_price_history_entry(row_dict[COL_PRICE], TODAY)
        row_dict[COL_SOLD]       = False
        row_dict[COL_FIRST_SEEN] = TODAY   # ← mark when first seen
        df_all = pd.concat([df_all, pd.DataFrame([row_dict])], ignore_index=True)

# Detect sold
if not df_history.empty:
    base_unsold = df_history[~df_history[COL_SOLD].fillna(False)]
    if not base_unsold.empty:
        sold_links = set(base_unsold[COL_LINK]) - set(df_new[COL_LINK])
        if sold_links:
            df_sold_now = base_unsold[base_unsold[COL_LINK].isin(sold_links)].copy()
            if COL_SOLD not in df_all.columns:
                df_all[COL_SOLD] = False
            df_all.loc[df_all[COL_LINK].isin(sold_links), COL_SOLD] = True

# Final cleanup
df_all = df_all.drop_duplicates(subset=[COL_LINK], keep='last').reset_index(drop=True)

logger.info(
    f"New: {len(df_new_only)}  |  Changed: {len(df_changed)}  |  "
    f"Sold: {len(df_sold_now)}  |  Total unique: {len(df_all)}"
)

# Save data
df_all.to_parquet(HISTORY_FILE, index=False)
df_all.to_csv("all_listings_history.csv", index=False, encoding='utf-8-sig')

# ================= GENERATE HTML DASHBOARD =================
generate_html(df_all, NOW_STR)


# ================= EXCEL EXPORT =================
df_export = df_all.copy()
df_export['Current Price']  = df_export[COL_PRICE].apply(lambda x: f"{int(round(x)):,} €" if pd.notna(x) else "")
df_export['Price per m²']   = df_export[COL_PRICE_PER_SQM].apply(lambda x: f"{int(round(x)):,} €/m²" if pd.notna(x) else "")
df_export['Size']           = df_export[COL_SIZE].apply(lambda x: f"{int(x)} m²" if pd.notna(x) else "")
df_export[COL_PRICE_HISTORY] = df_export.apply(
    lambda r: append_current_if_needed(r[COL_PRICE_HISTORY], r[COL_PRICE], r[COL_SCRAPED_DATE]), axis=1
)
df_export = df_export.rename(columns={
    COL_PRICE_HISTORY:      'Price History',
    COL_SITE_PRICE_HISTORY: 'Site price history',
    COL_PRICE:              'Current Price (numeric)',
    COL_PRICE_PER_SQM:      'Price per m² (numeric)',
    COL_SIZE:               'Size (numeric)',
    COL_SCRAPED_DATE:       'Scraped Date',
    COL_FIRST_SEEN:         'First Seen Date',
    COL_LOCATION:           'Location',
    COL_TITLE:              'Title',
    COL_FLOOR:              'Floor',
    COL_TOTAL_FLOORS:       'Total Floors',
    COL_YEAR:               'Year Built',
})
df_export.to_excel(excel_file, index=False, engine='openpyxl')

try:
    wb = load_workbook(excel_file)
    ws = wb.active
    sold_col_idx = None
    for col_idx in range(1, ws.max_column + 1):
        if ws.cell(row=1, column=col_idx).value == COL_SOLD:
            sold_col_idx = col_idx
            break
    if sold_col_idx:
        for row_idx in range(2, ws.max_row + 1):
            if ws.cell(row=row_idx, column=sold_col_idx).value:
                for col_idx in range(1, ws.max_column + 1):
                    c = ws.cell(row=row_idx, column=col_idx)
                    c.font = Font(strike=True, color="777777", name=c.font.name, size=c.font.sz)
    wb.save(excel_file)
except Exception as e:
    logger.error(f"Excel formatting error: {e}")

logger.info(f"Excel saved: {excel_file}")


# ================= EMAIL =================
if len(df_new_only) > 0 or len(df_changed) > 0 or len(df_sold_now) > 0:

    def fmt_p(x):  return f"{x:,.0f} €" if pd.notna(x) else "—"
    def fmt_s(x):  return f"{int(x)} m²" if pd.notna(x) else "—"
    def fmt_m(x):  return f"{int(round(x))} €/m²" if pd.notna(x) else "—"

    CSS = """<style>
body{font-family:Arial,sans-serif;line-height:1.6;color:#333;background:#f9f9f9}
.wrap{max-width:900px;margin:0 auto;background:#fff;border-radius:8px;overflow:hidden;border:1px solid #e0e0e0}
.hdr{background:#1a1d27;color:#fff;padding:20px 28px}
.hdr h1{font-size:18px;margin:0}
.hdr p{font-size:12px;color:#7a7d9a;margin:4px 0 0}
.body{padding:24px 28px}
h3{color:#444;font-size:14px;margin:24px 0 8px;border-bottom:2px solid #f0f0f0;padding-bottom:6px}
.imot-table{border-collapse:collapse;width:100%;font-size:13px;margin-bottom:8px}
.imot-table th{background:#f5f5f5;font-size:11px;text-transform:uppercase;letter-spacing:.04em;color:#666;padding:8px 12px;text-align:left;border-bottom:2px solid #e0e0e0}
.imot-table td{padding:9px 12px;border-bottom:1px solid #f0f0f0;vertical-align:top}
.imot-table tr:hover td{background:#fafafa}
a{color:#4f9cf9;text-decoration:none}
.pill{display:inline-block;padding:2px 10px;border-radius:99px;font-size:12px;font-weight:600}
.pill.new{background:#dcfce7;color:#166534}
.pill.chg{background:#fef3c7;color:#92400e}
.pill.sld{background:#fee2e2;color:#991b1b}
.sold-table td{color:#999;text-decoration:line-through}
.ftr{background:#f5f5f5;padding:12px 28px;font-size:12px;color:#999;text-align:center}
</style>"""

    def to_html_table(df_in, cols, headers, extra_class=""):
        thead = "<tr>" + "".join(f"<th>{h}</th>" for h in headers) + "</tr>"
        rows  = []
        for _, r in df_in.iterrows():
            cells = []
            for c in cols:
                v = r.get(c, "")
                if c == COL_LINK:
                    cells.append(f"<td><a href='{v}' target='_blank'>Виж →</a></td>")
                elif c == COL_PRICE:    cells.append(f"<td>{fmt_p(v)}</td>")
                elif c == COL_SIZE:     cells.append(f"<td>{fmt_s(v)}</td>")
                elif c == COL_PRICE_PER_SQM: cells.append(f"<td>{fmt_m(v)}</td>")
                elif c == 'Price_EUR_old': cells.append(f"<td style='text-decoration:line-through;color:#999'>{fmt_p(v)}</td>")
                else: cells.append(f"<td>{v if pd.notna(v) and v != '' else '—'}</td>")
            rows.append("<tr>" + "".join(cells) + "</tr>")
        cls = f"imot-table {extra_class}".strip()
        return f"<table class='{cls}'><thead>{thead}</thead><tbody>{''.join(rows)}</tbody></table>"

    new_section = changed_section = sold_section = ""

    if not df_new_only.empty:
        tbl = to_html_table(df_new_only,
            [COL_LOCATION, COL_PRICE, COL_SIZE, COL_PRICE_PER_SQM, COL_SITE_PRICE_HISTORY, COL_LINK],
            ["Локация", "Цена", "Площ", "€/m²", "История (сайт)", ""])
        new_section = f"""<h3><span class="pill new">НОВИ</span> &nbsp;{len(df_new_only)} обяви &mdash; {TODAY}</h3>{tbl}"""

    if not df_changed.empty:
        d = df_changed.copy()
        # Ensure price per sqm exists and is numeric
        if COL_PRICE_PER_SQM not in d.columns and f'{COL_PRICE_PER_SQM}_new' in d.columns:
            d = d.rename(columns={f'{COL_PRICE_PER_SQM}_new': COL_PRICE_PER_SQM})

        if COL_PRICE in d.columns and COL_SIZE in d.columns and COL_PRICE_PER_SQM not in d.columns:
            d[COL_PRICE_PER_SQM] = (d[COL_PRICE] / d[COL_SIZE]).round(2)

    if not df_sold_now.empty:
        tbl = to_html_table(df_sold_now,
            [COL_LOCATION, COL_PRICE, COL_SIZE, COL_PRICE_PER_SQM, COL_LINK],
            ["Локация", "Последна цена", "Площ", "€/m²", ""],
            extra_class="sold-table")
        sold_section = f"""<h3><span class="pill sld">ПРОДАДЕНИ</span> &nbsp;{len(df_sold_now)} обяви &mdash; {TODAY}</h3>{tbl}"""

    subject_parts = []
    if df_new_only:   subject_parts.append(f"{len(df_new_only)} НОВИ")
    if df_changed:    subject_parts.append(f"{len(df_changed)} ПРОМЯНА")
    if df_sold_now:   subject_parts.append(f"{len(df_sold_now)} ПРОДАДЕНИ")
    subject = f"Имоти – {' · '.join(subject_parts)} – {TODAY}"

    html_content = f"""<html><head><meta charset="utf-8">{CSS}</head><body>
<div class="wrap">
  <div class="hdr">
    <h1>git --version</h1>
    <p>Автоматичен отчет · {NOW_STR}</p>
  </div>
  <div class="body">
    {new_section}{changed_section}{sold_section}
    <p style="margin-top:24px;font-size:13px;color:#666">
      Общо уникални обяви в базата: <strong>{len(df_all)}</strong><br>
      Пълният списък е прикачен като Excel.
    </p>
  </div>
  <div class="ftr">Тук може да е вашият следващ ДОМ! 🏡</div>
</div>
</body></html>"""

    msg = MIMEMultipart()
    msg['From']    = SENDER_EMAIL
    msg['To']      = ", ".join(RECEIVERS)
    msg['Subject'] = subject
    msg.attach(MIMEText(html_content, "html", _charset="utf-8"))

    try:
        with open(excel_file, 'rb') as f:
            part = MIMEApplication(f.read(), Name=excel_file)
            part['Content-Disposition'] = f'attachment; filename="{excel_file}"'
            msg.attach(part)
    except Exception as e:
        logger.error(f"Failed to attach Excel: {e}")

    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.send_message(msg)
        server.quit()
        logger.info("✔ Email sent")
    except Exception as e:
        logger.error(f"❌ Email error: {e}")
else:
    logger.info("No changes → email not sent")

logger.info("=== Script finished ===")