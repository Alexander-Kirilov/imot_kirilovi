"""Обяви от други агенции (ERA, Home2U, Явлена) — всяка в отделен таб на dashboard-а.

Всяка агенция има fetch функция, която връща списък от речници с общите колони
по-долу (същите имена като в imot_scraper.py, за да се рисуват със същите таблици).
Историята на всяка агенция се пази в agencies/<key>.parquet: цена и ценова
история, кога е видяна за пръв път, свалена ли е.

Нова агенция = нова fetch_<име>() функция + ред в AGENCIES.
Търсенето е в GitHub secrets ERA_URL / HOME2U_URL / YAVLENA_URL: филтрираш в
сайта и копираш адреса (репото е публично, затова адресите не са в кода).
Историите се пазят криптирани (secure_store) — agencies/<key>.parquet.enc.
"""
import logging
import os
import random
import re
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import parse_qs, parse_qsl, unquote, urljoin, urlsplit

import pandas as pd
import requests
from bs4 import BeautifulSoup

import duplicates
import secure_store

logger = logging.getLogger(__name__)

AGENCY_DIR = Path("agencies")

# ── Колони (същите като в imot_scraper.py) ────────────────────────────────────
COL_LINK = 'Link'
COL_TITLE = 'Title'
COL_LOCATION = 'Location'
COL_PRICE = 'Price_EUR'
COL_SIZE = 'Size_sqm'
COL_PRICE_PER_SQM = 'Price_EUR_per_sqm'
COL_FLOOR = 'Floor'
COL_TOTAL_FLOORS = 'Total_floors'
COL_YEAR = 'Year_built'
COL_CONSTRUCTION = 'Construction_Type'
COL_IMAGES = 'Image_Paths'  # тук са директни URL-и към снимките в сайта на агенцията
COL_SCRAPED_DATE = 'Scraped_Date'
COL_FIRST_SEEN = 'First_Seen_Date'
COL_FIRST_SCRAPED = 'First_Scraped_Date'
COL_SITE_DATE = 'Site_Ad_Date'
COL_SITE_DATE_KIND = 'Site_Ad_Date_Kind'
COL_PRICE_HISTORY = 'Price_History'
COL_LAST_PRICE_CHANGE_DATE = 'Last_Price_Change_Date'
COL_SOLD = 'Sold'
COL_AGE_DAYS = 'Age_Days'
COL_SOURCE = 'Source'  # име на сайта за показване: ERA, Home2U, …
COL_SOURCE_KEY = 'Source_Key'  # ключ на сайта за филтъра в dashboard-а: era, home2u, …
COL_BULK_IMPORT = 'Bulk_Import'  # добавена при първоначално зареждане → не е "нова"
COL_OLD_PRICE = 'Price_EUR_old'  # само в df "changed" — за имейла

NUMERIC_COLS = (COL_PRICE, COL_SIZE, COL_PRICE_PER_SQM, COL_FLOOR, COL_TOTAL_FLOORS, COL_YEAR)
TEXT_COLS = (
    COL_LINK, COL_TITLE, COL_LOCATION, COL_CONSTRUCTION, COL_IMAGES, COL_SCRAPED_DATE,
    COL_FIRST_SEEN, COL_FIRST_SCRAPED, COL_SITE_DATE, COL_SITE_DATE_KIND,
    COL_PRICE_HISTORY, COL_LAST_PRICE_CHANGE_DATE, COL_SOURCE, COL_SOURCE_KEY,
    duplicates.COL_IMAGE_HASHES,
)

# Ако в един run се появят повече нови от толкова → смятаме ги за първоначално
# зареждане (смяна на филтъра, първи run) и не ги пращаме по имейла като "нови"
BULK_IMPORT_THRESHOLD = 5

USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/130.0 Safari/537.36")


# ================= HELPERS =================

def _session():
    s = requests.Session()
    s.headers.update({
        "User-Agent": USER_AGENT,
        "Accept-Language": "bg-BG,bg;q=0.9,en;q=0.8",
    })
    return s


def _get(session, url, **kwargs):
    """GET с един повторен опит при грешка."""
    for attempt in (1, 2):
        try:
            r = session.get(url, timeout=30, **kwargs)
            r.raise_for_status()
            return r
        except requests.RequestException:
            if attempt == 2:
                raise
            time.sleep(3)


def _soup(session, url):
    return BeautifulSoup(_get(session, url).text, "html.parser")


def _pause():
    time.sleep(0.6 + random.uniform(0, 0.6))


def _price(text):
    """'215 000 €' / '€ 270,000' / '270,000 €' → 215000.0. Под 10 000 → None (това е €/m²)."""
    m = re.search(r'\d[\d\s  .,]*', str(text or ""))
    if not m:
        return None
    raw = re.sub(r'[.,]\d{2}$', '', m.group().strip())  # махаме стотинките, ако има
    digits = re.sub(r'\D', '', raw)
    if not digits:
        return None
    val = float(digits)
    return val if val >= 10000 else None


def _area(text):
    """'102.12 м2' / '84 м²' / '88 кв.м.' → 102.12. Абсурдни стойности → None."""
    m = re.search(r'(\d+(?:[.,]\d+)?)', str(text or "").replace("\xa0", " "))
    if not m:
        return None
    val = float(m.group(1).replace(",", "."))
    return val if 10 <= val <= 1000 else None


def _int(text):
    """'6' → 6, 'партер' → 0, '1984 г.' → 1984, иначе None."""
    s = str(text or "").strip().lower()
    if not s:
        return None
    if "партер" in s:
        return 0
    m = re.search(r'-?\d+', s)
    return int(m.group()) if m else None


def normalize_location(loc):
    """'кв. Младост 2, София' / 'Младост 1, София' / 'Младост 1A' → 'Младост 2' / 'Младост 1' / 'Младост 1А'."""
    s = re.sub(r'\s+', ' ', str(loc or '')).strip()
    s = re.sub(r'^(кв\.|ж\.к\.|жк\.?)\s*', '', s, flags=re.IGNORECASE)
    s = re.sub(r'[,|/]?\s*(гр\.\s*)?София$', '', s).strip(' ,|/')
    s = re.sub(r'^((гр\.\s*)?София\s*[,|/]\s*)+', '', s)
    s = re.sub(r'(\d)\s*[AaАа]$', r'\1А', s)  # латинско A → кирилско А
    return s


def history_entry(price, date_str):
    try:
        return f"{int(round(float(price))):,} € ({date_str})" if price and date_str else ""
    except (TypeError, ValueError):
        return ""


def _is_missing(v):
    if v is None:
        return True
    if isinstance(v, str):
        return v.strip() == "" or v.strip().lower() in ("nan", "none")
    try:
        return bool(pd.isna(v))
    except (TypeError, ValueError):
        return False


def days_since(date_str):
    s = str(date_str or "").strip()[:10]
    try:
        d = datetime.strptime(s, "%Y-%m-%d").date()
    except ValueError:
        return None
    return max((datetime.now().date() - d).days, 0)


def _listing(link, **fields):
    """Общият формат на една обява (празните полета остават None/'')."""
    row = {
        COL_LINK: link, COL_TITLE: "", COL_LOCATION: "", COL_PRICE: None, COL_SIZE: None,
        COL_FLOOR: None, COL_TOTAL_FLOORS: None, COL_YEAR: None, COL_CONSTRUCTION: "",
        COL_IMAGES: "", COL_SITE_DATE: "", COL_SITE_DATE_KIND: "",
    }
    row.update(fields)
    row[COL_LOCATION] = normalize_location(row[COL_LOCATION])
    return row


# ================= ERA =================
# Сайтът на ERA е преправен (2026) — обявите идват от JSON API-то, което ползва
# и самият сайт. URL параметрите от браузъра се превеждат 1:1 към тялото на заявката.

ERA_URL = os.environ.get("ERA_URL", "")
ERA_API = "https://www.era.bg/server/api"

ERA_BODY = {
    "returnCountOnly": False, "page": 1, "limit": 20, "orderBy": "date_desc",
    "offerType": 2, "propertyCategory": 1,
    "excludeBuildingsInConstruction": False, "onlyShowBuildingsInConstruction": False,
    "searchText": "", "priceType": "Total", "areaType": "Gross", "floorOption": "None",
    "hasElevator": False, "withoutCommissionOnly": False,
    "propertyTypes": [], "territories": [], "localities": [],
    "constructionTypes": [], "constructionStages": [], "garageOptions": [],
    "propertyConditions": [], "furnishing": [], "heatingTypes": [],
    "propertyOrientations": [], "isExclusive": False,
    "bedrooms": [], "baths": [], "wcs": [],
}
# Числови филтри, които сайтът слага в URL-а само когато са зададени
ERA_NUM_KEYS = ("fromPrice", "toPrice", "fromArea", "toArea", "fromFloor", "toFloor",
                "constructionYear")
# Резервен речник, ако /taxonomies не отговори
ERA_CONSTRUCTION = {1: "Панел", 2: "ЕПК", 4: "Тухла", 5: "Гредоред", 6: "Сглобяема конструкция"}


def era_body_from_url(url):
    body = dict(ERA_BODY)
    for key, val in parse_qsl(urlsplit(url).query):
        current = body.get(key)
        if isinstance(current, list):
            body[key] = [int(v) for v in val.split(",") if v.strip().isdigit()]
        elif isinstance(current, bool):
            body[key] = val.lower() == "true"
        elif key in ERA_NUM_KEYS or isinstance(current, int):
            n = _int(val.replace("\xa0", "").replace(" ", ""))
            if n is not None:
                body[key] = n
        elif key in body:
            body[key] = val
    return body


def _era_construction_names(session):
    try:
        tax = _get(session, f"{ERA_API}/taxonomies").json()
        return {int(t["value"]): t["text"] for t in tax.get("constructionTypes", [])}
    except Exception as tax_err:
        logger.debug(f"ERA taxonomies failed ({tax_err}) → резервен речник")
        return ERA_CONSTRUCTION


def fetch_era():
    s = _session()
    s.headers.update({"Accept": "application/json", "Referer": ERA_URL})
    body = era_body_from_url(ERA_URL)

    r = s.post(f"{ERA_API}/offers", json={**body, "returnCountOnly": True}, timeout=30)
    r.raise_for_status()
    total = int(r.json().get("offers") or 0)
    logger.info(f"[ERA] Общо обяви по филтъра: {total}")

    offers, page = {}, 1
    while len(offers) < total and page <= 50:
        r = s.post(f"{ERA_API}/offers", json={**body, "page": page}, timeout=30)
        r.raise_for_status()
        chunk = r.json().get("offers") or []
        if not chunk:
            break
        for o in chunk:
            if o.get("number"):
                offers[str(o["number"])] = o
        page += 1

    constr_names = _era_construction_names(s)
    result = []
    for idx, (number, o) in enumerate(offers.items(), start=1):
        if o.get("isSold"):
            continue
        images = sorted(o.get("images") or [], key=lambda im: im.get("order") or 0)
        row = _listing(
            f"https://www.era.bg/imoti/oferta/{number}",
            **{
                COL_TITLE: o.get("title") or "",
                COL_LOCATION: o.get("locality") or "",
                COL_PRICE: float(o["sellingPrice"]) if o.get("sellingPrice") else None,
                COL_SIZE: _area(o.get("area")),
                COL_IMAGES: ",".join(im["file"] for im in images[:2] if im.get("file")),
                COL_SITE_DATE: str(o.get("publishedAt") or "")[:10],
                COL_SITE_DATE_KIND: "Публикувана" if o.get("publishedAt") else "",
            },
        )
        if (o.get("currency") or "eur").lower() != "eur":
            logger.warning(f"[ERA] {number}: цената е в {o.get('currency')}, не в евро")

        # Етаж, година и строителство има само в детайла
        try:
            d = _get(s, f"{ERA_API}/offers/{number}").json()
            d = d.get("offer", d)
            row[COL_FLOOR] = _int(d.get("floor"))
            row[COL_TOTAL_FLOORS] = _int(d.get("totalFloors"))
            row[COL_YEAR] = _int(d.get("constructionYear"))
            ct = _int(d.get("constructionType"))
            row[COL_CONSTRUCTION] = constr_names.get(ct, "") if ct else ""
            _pause()
        except Exception as det_err:
            logger.warning(f"[ERA] Детайлът на {number} не се зареди: {det_err}")

        logger.info(f"[ERA] [{idx}/{len(offers)}] {row[COL_LOCATION]} · "
                    f"{row[COL_PRICE] or '—'} € · {row[COL_SIZE] or '—'} m²")
        result.append(row)
    return result


# ================= HOME2U =================

HOME2U_URL = os.environ.get("HOME2U_URL", "")
# За страниците на проекти (таблица с апартаменти) — кои видове взимаме
HOME2U_ROOM_TYPES = ("3-стаен", "4-стаен", "многостаен")


def _home2u_params(soup):
    """{'Местоположение': 'Младост 1, София', 'Етаж': '2', …} от иконките в обявата."""
    params = {}
    for li in soup.select("li"):
        p = li.select_one("p")
        if not p or not li.select_one("[class*='ico-']"):
            continue
        label, sep, value = p.get_text(" ", strip=True).partition(":")
        if sep:
            params.setdefault(label.strip(), value.strip())
    return params


def _home2u_card(article):
    """Данните от картата в списъка — резерва, ако детайлът не се зареди."""
    text = re.sub(r'\s+', ' ', article.get_text(" | ", strip=True))
    img = article.select_one("img")
    m_price = re.search(r'€\s*([\d\s.,]+)|([\d\s.,]+)\s*€', text)
    m_size = re.search(r'([\d.,]+)\s*м2', text)
    return {
        COL_LOCATION: text.split("|")[0].strip(),
        COL_PRICE: _price(m_price.group(0)) if m_price else None,
        COL_SIZE: _area(m_size.group(1)) if m_size else None,
        COL_IMAGES: (img.get("data-src") or img.get("src") or "") if img else "",
    }


def fetch_home2u():
    s = _session()
    queue, pages_seen, page_url = [], set(), HOME2U_URL
    while page_url and page_url not in pages_seen and len(pages_seen) < 20:
        pages_seen.add(page_url)
        soup = _soup(s, page_url)
        for art in soup.select("article.article-catalog"):
            a = art.select_one(".article-catalog__body-title a[href]")
            if a:
                queue.append((urljoin(page_url, a["href"].strip()), _home2u_card(art)))
        nxt = soup.select_one("a.next[href], link[rel='next']")
        page_url = urljoin(page_url, nxt["href"]) if nxt else None
    logger.info(f"[Home2U] Обяви в списъка: {len(queue)}")

    result, visited = [], set()
    while queue:
        url, pre = queue.pop(0)
        if url in visited:
            continue
        visited.add(url)
        try:
            soup = _soup(s, url)
        except Exception as det_err:
            logger.warning(f"[Home2U] {url} не се зареди ({det_err}) → само данните от списъка")
            result.append(_listing(url, **pre))
            continue
        _pause()

        # Страница на проект → таблица със свободните апартаменти
        rows = soup.select(".list__table-body")
        if rows:
            project = soup.select_one(".section_body-content h5") or soup.select_one("h1")
            project_name = project.get_text(" ", strip=True) if project else ""
            project_loc = _home2u_params(soup).get("Местоположение", pre.get(COL_LOCATION, ""))
            added = 0
            for row in rows:
                status = row.select_one(".list__col-status")
                if not status or "свободен" not in status.get_text(" ", strip=True).lower():
                    continue
                type_el = row.select_one(".list__col-type")
                type_txt = (type_el.get_text(" ", strip=True) + " " + (type_el.get("data-type") or "")
                            ).lower() if type_el else ""
                if not any(t in type_txt for t in HOME2U_ROOM_TYPES):
                    continue
                a = row.select_one("a[href]")
                if not a:
                    continue
                cell = lambda sel: (row.select_one(sel).get_text(" ", strip=True)
                                    if row.select_one(sel) else "")
                queue.append((urljoin(url, a["href"]), {
                    COL_TITLE: f"{project_name} - {cell('.list__col-name')}".strip(" -"),
                    COL_LOCATION: project_loc,
                    COL_PRICE: _price(cell(".list__col-price")),
                    COL_SIZE: _area(cell(".list__col-help")),
                    COL_FLOOR: _int(cell(".list__col-floor")),
                }))
                added += 1
            logger.info(f"[Home2U] Проект '{project_name}': {added} свободни апартамента")
            continue

        params = _home2u_params(soup)
        h1 = soup.select_one("h1")
        price_el = soup.select_one(".section__head-price h2") or soup.select_one("[class*='price']")
        og_img = soup.find("meta", property="og:image")
        row = _listing(url, **{
            COL_TITLE: pre.get(COL_TITLE) or (h1.get_text(" ", strip=True) if h1 else ""),
            COL_LOCATION: params.get("Местоположение") or pre.get(COL_LOCATION, ""),
            COL_PRICE: (_price(price_el.get_text(" ", strip=True)) if price_el else None)
                       or pre.get(COL_PRICE),
            COL_SIZE: _area(params.get("Площ в квадратни метри")) or pre.get(COL_SIZE),
            COL_FLOOR: _int(params.get("Етаж")) if params.get("Етаж") else pre.get(COL_FLOOR),
            COL_CONSTRUCTION: params.get("Строителство", ""),
            COL_IMAGES: (og_img.get("content") if og_img else "") or pre.get(COL_IMAGES, ""),
        })
        logger.info(f"[Home2U] {row[COL_LOCATION]} · {row[COL_PRICE] or '—'} € · "
                    f"{row[COL_SIZE] or '—'} m² · {url}")
        result.append(row)
    return result


# ================= ЯВЛЕНА =================

YAVLENA_URL = os.environ.get("YAVLENA_URL", "")


def _yavlena_detail(html, card_title):
    soup = BeautifulSoup(html, "html.parser")
    text = soup.get_text("\n", strip=True)
    main = text.split("Подобни имоти")[0]  # отдолу са чужди обяви с други цени
    h1 = soup.find("h1")
    title = re.sub(r'\s+', ' ', h1.get_text(" ", strip=True)) if h1 else ""

    def find(pattern):
        m = re.search(pattern, main, re.IGNORECASE)
        return m.group(1).strip() if m else ""

    price = None
    for m in re.finditer(r'€\s*(\d[\d\s ]*)', main):
        price = _price(m.group(1))
        if price:
            break

    location = find(r'София\s*/\s*([^\n]+)')
    if not location and card_title:
        # "Тристаен апартамент, София, Младост 2, 88 кв.м., 220 000 €"
        parts = [p.strip() for p in card_title.split(",")]
        if "София" in parts and parts.index("София") + 1 < len(parts):
            location = parts[parts.index("София") + 1]

    date = ""
    m_date = re.search(r'(\d{2})\.(\d{2})\.(\d{4})\s*г\.', main)
    if m_date:
        date = f"{m_date.group(3)}-{m_date.group(2)}-{m_date.group(1)}"

    # Галерията е първа в страницата: images.yavlena.com/<размер>/<файл> (URL-кодирано)
    imgs = []
    for size, name in re.findall(
            r'images\.yavlena\.com(?:%2F|/)(\w+)(?:%2F|/)([\w-]+\.(?:jpe?g|png|webp))', html):
        url = f"https://images.yavlena.com/{size}/{unquote(name)}"
        if name not in "".join(imgs):
            imgs.append(url)

    return {
        COL_TITLE: title,
        COL_LOCATION: location,
        COL_PRICE: price,
        COL_SIZE: _area(find(r'(\d+(?:[.,]\d+)?\s*кв\.?\s*м)')) or _area(title),
        COL_FLOOR: _int(find(r'Етаж:\s*(\d+|партер)')),
        COL_YEAR: _int(find(r'Год(?:\.|ина)\s*на\s*построяване:?\s*(\d{4})')),
        COL_CONSTRUCTION: find(r'\nСтроителство\s*\n\s*([^\n:]+)'),
        COL_IMAGES: ",".join(imgs[:2]),
        COL_SITE_DATE: date,
        COL_SITE_DATE_KIND: "Дата в обявата" if date else "",
    }


def fetch_yavlena():
    s = _session()
    soup = _soup(s, YAVLENA_URL)
    cards = {}
    for a in soup.select("a[href]"):
        m = re.fullmatch(r'(?:https://www\.yavlena\.com)?/bg/(\d+)', a["href"].strip())
        if m:
            cards.setdefault(f"https://www.yavlena.com/bg/{m.group(1)}", a.get("title") or "")
    logger.info(f"[Явлена] Обяви в списъка: {len(cards)}")

    result = []
    for idx, (url, card_title) in enumerate(cards.items(), start=1):
        try:
            fields = _yavlena_detail(_get(s, url).text, card_title)
            _pause()
        except Exception as det_err:
            logger.warning(f"[Явлена] {url} не се зареди ({det_err}) → само данните от списъка")
            parts = [p.strip() for p in card_title.split(",")]
            fields = {COL_LOCATION: parts[2] if len(parts) > 2 else ""}
        row = _listing(url, **fields)
        logger.info(f"[Явлена] [{idx}/{len(cards)}] {row[COL_LOCATION]} · "
                    f"{row[COL_PRICE] or '—'} € · {row[COL_SIZE] or '—'} m²")
        result.append(row)
    return result


# ================= ДЕТАЙЛИ САМО ЗА НОВИ ОБЯВИ =================
# При големите сайтове (стотици обяви) етажът/строителството се теглят веднъж —
# за обява, която вече е в историята, update_history пази предишните стойности.

def _known_links(key):
    hist = load_history(key)
    return set(hist[COL_LINK]) if not hist.empty and COL_LINK in hist.columns else set()


CONSTRUCTION_WORDS = ("Тухла", "Панел", "ЕПК", "ПК", "Гредоред", "Монолит", "Сглобяема")


def _construction_word(text):
    for word in CONSTRUCTION_WORDS:
        if re.search(rf'(?<![А-Яа-я]){word}(?![А-Яа-я])', str(text or "")):
            return word
    return ""


# ================= HOMES.BG =================
# Внимание: "днес"/"вчера" в homes.bg и "Активирана на" в imoti.net са дати на
# подновяване, не на публикуване — затова не се ползват за "Добавена".
# Сайтът е React приложение — обявите идват от неговото JSON API.
# HOMES_URL е адресът на търсенето от браузъра (homes.bg/?typeId=…&neighbourhoods[]=…);
# API-то приема същите параметри.

HOMES_URL = os.environ.get("HOMES_URL", "")
HOMES_API = "https://www.homes.bg/api/"


def _homes_photo(photo):
    if not photo or not photo.get("name"):
        return ""
    portrait = int(photo.get("height") or 0) > int(photo.get("width") or 0)
    return f"https://g1.homes.bg/{photo.get('path', '')}{photo['name']}{'o' if portrait else 'b'}.jpg"


def fetch_homes():
    s = _session()
    s.headers.update({"Accept": "application/json", "Referer": "https://www.homes.bg/"})
    query = urlsplit(HOMES_URL).query
    offers, start, total = {}, 0, None
    while start < 3000:
        d = _get(s, f"{HOMES_API}offers?{query}&startIndex={start}&stopIndex={start + 99}").json()
        total = d.get("offersCount", total)
        chunk = d.get("result") or []
        for o in chunk:
            offers.setdefault(f"{o.get('type')}{o.get('id')}", o)
        if not chunk or not d.get("hasMoreItems"):
            break
        start += 100
        _pause()
    logger.info(f"[homes.bg] Обяви по филтъра: {total} | изтеглени: {len(offers)}")

    known = _known_links("homes")
    result, details = [], 0
    for o in offers.values():
        link = urljoin("https://www.homes.bg", o.get("viewHref") or "")
        title = o.get("title") or ""  # "Тристаен, 76m²"
        m_size = re.search(r'(\d+(?:[.,]\d+)?)\s*m', title)
        photos = o.get("photos") or ([o["photo"]] if o.get("photo") else [])
        row = _listing(link, **{
            COL_TITLE: title,
            COL_LOCATION: o.get("location") or "",
            COL_PRICE: _price((o.get("price") or {}).get("value")),
            COL_SIZE: _area(m_size.group(1)) if m_size else None,
            COL_CONSTRUCTION: _construction_word(o.get("description")),  # "Панел, Обзаведен, ТЕЦ"
            COL_IMAGES: ",".join(u for u in (_homes_photo(ph) for ph in photos[:2]) if u),
        })
        if link not in known:
            try:
                attrs = {a.get("key"): a.get("value") for a in
                         _get(s, f"{HOMES_API}offers/{o.get('type')}/{o.get('id')}").json()
                         .get("data", {}).get("attributes", [])}
                row[COL_FLOOR] = _int(attrs.get("floor"))            # "2-ри"
                row[COL_TOTAL_FLOORS] = _int(attrs.get("total_floors"))
                row[COL_CONSTRUCTION] = attrs.get("build_type") or row[COL_CONSTRUCTION]
                details += 1
                _pause()
            except Exception as det_err:
                logger.warning(f"[homes.bg] Детайлът на {link} не се зареди: {det_err}")
        result.append(row)
    logger.info(f"[homes.bg] Детайли (само нови обяви): {details}")
    return result


# ================= IMOTI.NET =================
# Търсенето е POST форма, която връща номер на търсене (sid) за страниците.
# IMOTINET_SEARCH са полетата на формата като query string, напр.
# ad_type_id=2&world_area_id=1&property_type_id[]=9&second_descendant_id[]=5758&…

IMOTINET_SEARCH = os.environ.get("IMOTINET_SEARCH", "")
IMOTINET = "https://www.imoti.net"


def _imotinet_cards(html):
    soup = BeautifulSoup(html, "html.parser")
    cards = []
    for li in soup.select("li.clearfix"):
        a = li.select_one("a[href*='/obiava/']")
        info = li.select_one("div.real-estate-text")
        if not a or not info:
            continue
        text = re.sub(r'\s+', ' ', info.get_text(" ", strip=True))
        heading = info.select_one("h3")
        heading_txt = heading.get_text(" ", strip=True) if heading else ""
        m_size = re.search(r'(\d+(?:[.,]\d+)?)\s*м', heading_txt)
        price_el = info.select_one(".price")  # "387 671 € 758 219 BGN" — отделно от квартала
        m_price = re.search(r'(\d[\d\s]*)\s*€', price_el.get_text(" ", strip=True)) if price_el else None
        m_floor = re.search(r'Етаж:\s*(\d+|партер)(?:\s*от\s*(\d+))?', text, re.IGNORECASE)
        loc = info.select_one("span.location")
        img = li.select_one("img")
        cards.append(_listing(IMOTINET + re.sub(r'\?.*$', '', a["href"]), **{
            COL_TITLE: heading_txt,
            COL_LOCATION: loc.get_text(strip=True) if loc else "",
            COL_PRICE: _price(m_price.group(1)) if m_price else None,
            COL_SIZE: _area(m_size.group(1)) if m_size else None,
            COL_FLOOR: _int(m_floor.group(1)) if m_floor else None,
            COL_TOTAL_FLOORS: _int(m_floor.group(2)) if m_floor and m_floor.group(2) else None,
            COL_IMAGES: urljoin(IMOTINET, img.get("src") or img.get("data-src") or "") if img else "",
        }))
    return cards, soup


def fetch_imotinet():
    s = _session()
    _get(s, f"{IMOTINET}/bg/obiavi/r/prodava/sofia/")  # бисквитка за сесията
    fields = parse_qsl(IMOTINET_SEARCH, keep_blank_values=True)
    if not any(k == "items_per_page" for k, _ in fields):
        fields.append(("items_per_page", "30"))
    r = s.post(f"{IMOTINET}/bg/obiavi/r", data=fields, timeout=30)
    r.raise_for_status()
    base, sid = r.url.split("?")[0], parse_qs(urlsplit(r.url).query).get("sid", [""])[0]
    if not sid:
        raise RuntimeError("imoti.net не върна номер на търсене (sid)")

    cards, soup = _imotinet_cards(r.text)
    last_page = max([int(n) for n in re.findall(r'[?&]page=(\d+)&(?:amp;)?sid=' + re.escape(sid), r.text)] or [1])
    for page in range(2, min(last_page, 40) + 1):
        _pause()
        more, _ = _imotinet_cards(_get(s, f"{base}?page={page}&sid={sid}").text)
        if not more:
            break
        cards += more
    by_link = {c[COL_LINK]: c for c in cards}
    logger.info(f"[imoti.net] Страници: {last_page} | обяви: {len(by_link)}")

    # Строителството е само в детайла → теглим го за обявите, които още не познаваме
    known, details = _known_links("imotinet"), 0
    for link, row in by_link.items():
        if link in known:
            continue
        try:
            text = BeautifulSoup(_get(s, link).text, "html.parser").get_text("\n", strip=True)
            # "Строителство:\nПанел" — от началото на реда, за да не хване "Година на строителство:"
            m = re.search(r'(?mi)^строителство:?\s*\n([^\n]+)', text)
            row[COL_CONSTRUCTION] = _construction_word(m.group(1)) if m else ""
            m_year = re.search(r'(?mi)^Година на строителство:?\s*\n\s*(\d{4})', text)
            row[COL_YEAR] = int(m_year.group(1)) if m_year else None
            details += 1
            _pause()
        except Exception as det_err:
            logger.warning(f"[imoti.net] Детайлът на {link} не се зареди: {det_err}")
    logger.info(f"[imoti.net] Детайли (само нови обяви): {details}")
    return list(by_link.values())


# ================= ИРИДА =================
# robots.txt на irida.bg забранява AI роботите (вкл. Claude), но не и останалите —
# скрейпърът тегли сайта, а кодът е писан по страници, запазени ръчно (Ctrl+S).
# IRIDA_URL е адресът на търсенето от браузъра; страниците са &page=N.

IRIDA_URL = os.environ.get("IRIDA_URL", "")
MLADOST = ("Младост 1", "Младост 1А", "Младост 2", "Младост 3", "Младост 4")


def _irida_cards(html, page_url):
    """Картите от страницата с резултати: квартал, площ, цена, снимка."""
    soup = BeautifulSoup(html, "html.parser")
    cards = []
    for card in soup.select("div.property-listing"):
        a = card.find("a", href=re.compile(r"/offers/\d+/"))
        if not a:
            continue
        parts = [p.strip() for p in card.get_text("|", strip=True).split("|") if p.strip()]
        text = " | ".join(parts)  # "|" пази квартала ("Младост 1") отделно от цената
        m_size = re.search(r'(\d+(?:[.,]\d+)?)\s*m²', text)
        m_price = re.search(r'(\d[\d\s]*)\s*€', text)
        img = card.find("img")
        src = (img.get("data-src") or img.get("data-lazy") or img.get("src") or "") if img else ""
        cards.append(_listing(urljoin(page_url, re.sub(r'[?#].*$', '', a["href"])), **{
            COL_TITLE: next((p for p in parts if "апартамент" in p.lower()), ""),
            # "София, гр. София, Младост 1"
            COL_LOCATION: next((p for p in parts if p.startswith("София")), ""),
            COL_PRICE: _price(m_price.group(1)) if m_price else None,
            COL_SIZE: _area(m_size.group(1)) if m_size else None,
            COL_IMAGES: urljoin(page_url, src) if src else "",
        }))
    return cards, soup


def _irida_detail(html):
    """Етаж ("9/13"), строителство и година от страницата на обявата."""
    text = BeautifulSoup(html, "html.parser").get_text("\n", strip=True)
    m_floor = re.search(r'Етаж:\s*\n\s*(\d+|партер)\s*(?:/\s*(\d+))?', text, re.IGNORECASE)
    m_constr = re.search(r'Строителство:\s*\n\s*([^\n]+)', text)
    m_year = re.search(r'Година:\s*\n\s*(\d{4})', text)
    return {
        COL_FLOOR: _int(m_floor.group(1)) if m_floor else None,
        COL_TOTAL_FLOORS: _int(m_floor.group(2)) if m_floor and m_floor.group(2) else None,
        COL_CONSTRUCTION: _construction_word(m_constr.group(1)) if m_constr else "",
        COL_YEAR: int(m_year.group(1)) if m_year else None,
    }


def fetch_irida():
    s = _session()
    sep = "&" if "?" in IRIDA_URL else "?"
    cards, soup = _irida_cards(_get(s, IRIDA_URL).text, IRIDA_URL)
    pages = [int(n) for n in re.findall(r'[?&]page=(\d+)', " ".join(a["href"] for a in soup.find_all("a", href=True)))]
    for page in range(2, min(max(pages or [1]), 30) + 1):
        _pause()
        more, _ = _irida_cards(_get(s, f"{IRIDA_URL}{sep}page={page}").text, IRIDA_URL)
        if not more:
            break
        cards += more
    # Търсенето е по дума ("младост") → оставяме само Младост 1, 1А, 2, 3, 4
    by_link = {c[COL_LINK]: c for c in cards if c[COL_LOCATION] in MLADOST}
    logger.info(f"[Ирида] Страници: {max(pages or [1])} | карти: {len(cards)} | в Младост 1–4: {len(by_link)}")

    known, details = _known_links("irida"), 0
    for link, row in by_link.items():
        if link in known:
            continue
        try:
            row.update({k: v for k, v in _irida_detail(_get(s, link).text).items() if not _is_missing(v)})
            details += 1
            _pause()
        except Exception as det_err:
            logger.warning(f"[Ирида] Детайлът на {link} не се зареди: {det_err}")
    logger.info(f"[Ирида] Детайли (само нови обяви): {details}")
    return list(by_link.values())


# ================= РЕГИСТЪР =================

AGENCIES = [
    {"key": "era", "name": "ERA", "site": "era.bg", "url": ERA_URL,
     "secret": "ERA_URL", "fetch": fetch_era},
    {"key": "home2u", "name": "Home2U", "site": "home2u.bg", "url": HOME2U_URL,
     "secret": "HOME2U_URL", "fetch": fetch_home2u},
    {"key": "yavlena", "name": "Явлена", "site": "yavlena.com", "url": YAVLENA_URL,
     "secret": "YAVLENA_URL", "fetch": fetch_yavlena},
    {"key": "homes", "name": "homes.bg", "site": "homes.bg", "url": HOMES_URL,
     "secret": "HOMES_URL", "fetch": fetch_homes},
    {"key": "imotinet", "name": "imoti.net", "site": "imoti.net", "url": IMOTINET_SEARCH and
     f"{IMOTINET}/bg/obiavi/r/prodava/sofia/", "secret": "IMOTINET_SEARCH", "fetch": fetch_imotinet},
    {"key": "irida", "name": "Ирида", "site": "irida.bg", "url": IRIDA_URL,
     "secret": "IRIDA_URL", "fetch": fetch_irida},
]


# ================= ИСТОРИЯ =================

def history_path(key):
    return AGENCY_DIR / f"{key}.parquet"


def load_history(key):
    path = history_path(key)
    try:
        # Грешна парола → SystemExit (не се хваща тук), за да не презапишем историята
        return secure_store.read_parquet(path)
    except Exception as read_err:
        logger.error(f"Не мога да прочета {path}: {read_err}")
        return pd.DataFrame()


def finalize(df, agency):
    """Типове за parquet + производни колони (€/m², възраст, източник)."""
    if df.empty:
        return df
    df = df.copy()
    for col in NUMERIC_COLS:
        if col not in df.columns:
            df[col] = None
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df[COL_PRICE_PER_SQM] = (df[COL_PRICE] / df[COL_SIZE]).where(df[COL_SIZE] > 0).round(2)
    for col in TEXT_COLS:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].fillna("").astype(str).replace({"nan": "", "None": ""})
    df[COL_SOURCE] = agency["name"]
    df[COL_SOURCE_KEY] = agency["key"]
    for flag in (COL_SOLD, COL_BULK_IMPORT):
        df[flag] = df[flag].fillna(False).astype(bool) if flag in df.columns else False
    df[COL_AGE_DAYS] = pd.array([days_since(v) for v in df[COL_FIRST_SEEN]], dtype="Int64")
    return df.drop_duplicates(subset=[COL_LINK], keep="last").reset_index(drop=True)


def update_history(agency, scraped, today):
    """Слива изтеглените обяви с историята. Връща (df_all, new, changed, sold, bulk)."""
    hist = load_history(agency["key"])
    records = {r[COL_LINK]: r for r in hist.to_dict("records")} if not hist.empty else {}
    first_run = not records
    new_rows, changed_rows, sold_rows, seen = [], [], [], set()

    for item in scraped:
        link = item[COL_LINK]
        seen.add(link)
        price = item.get(COL_PRICE)
        old = records.get(link)

        if old is None:
            rec = dict(item)
            rec[COL_PRICE_HISTORY] = history_entry(price, today)
            rec[COL_FIRST_SCRAPED] = today
            rec[COL_FIRST_SEEN] = item.get(COL_SITE_DATE) or today
            rec[COL_LAST_PRICE_CHANGE_DATE] = ""
            new_rows.append(rec)
        else:
            rec = dict(old)
            # Празно поле в този run не трие вече познатата стойност
            rec.update({k: v for k, v in item.items() if not _is_missing(v)})
            rec[COL_FIRST_SEEN] = (item.get(COL_SITE_DATE) or old.get(COL_FIRST_SEEN)
                                   or old.get(COL_FIRST_SCRAPED) or today)
            old_price = old.get(COL_PRICE)
            if not _is_missing(price) and not _is_missing(old_price) \
                    and round(float(price)) != round(float(old_price)):
                hist_txt = str(old.get(COL_PRICE_HISTORY) or "").strip()
                if not hist_txt:
                    hist_txt = history_entry(old_price, old.get(COL_SCRAPED_DATE) or "преди")
                rec[COL_PRICE_HISTORY] = f"{hist_txt} → {history_entry(price, today)}"
                rec[COL_LAST_PRICE_CHANGE_DATE] = today
                changed_rows.append({**rec, COL_OLD_PRICE: old_price})
            elif old.get(COL_SOLD):
                logger.info(f"[{agency['name']}] Обявата е отново активна: {link}")
        rec[COL_SOLD] = False
        rec[COL_SCRAPED_DATE] = today
        records[link] = rec

    for link, rec in records.items():
        if link not in seen and not rec.get(COL_SOLD):
            rec[COL_SOLD] = True
            sold_rows.append(dict(rec))

    # Първи run или много нови наведнъж (напр. сменен филтър) → първоначално зареждане
    bulk = first_run or len(new_rows) > BULK_IMPORT_THRESHOLD
    for rec in new_rows:
        rec[COL_BULK_IMPORT] = bulk

    df_all = finalize(pd.DataFrame(list(records.values())), agency)
    to_df = lambda rows: finalize(pd.DataFrame(rows), agency) if rows else pd.DataFrame()
    return df_all, to_df(new_rows), to_df(changed_rows), to_df(sold_rows), bulk


def run_agency(agency, today):
    res = {
        "key": agency["key"], "name": agency["name"], "site": agency["site"],
        "url": agency["url"], "error": None, "fetched": 0, "bulk": False,
        "new": pd.DataFrame(), "changed": pd.DataFrame(), "sold": pd.DataFrame(),
    }
    hist = load_history(agency["key"])
    try:
        if not agency["url"]:
            raise RuntimeError(f"адресът за търсене не е зададен (secret {agency['secret']})")
        scraped = agency["fetch"]()
    except Exception as fetch_err:
        logger.error(f"[{agency['name']}] Тегленето пропадна: {fetch_err}")
        scraped, res["error"] = [], str(fetch_err).splitlines()[0][:200]

    if not scraped:
        # Нищо не е изтеглено → не маркираме всичко като свалено, показваме старото
        res["error"] = res["error"] or "не са намерени обяви"
        res["df"] = finalize(hist, agency)
        logger.warning(f"[{agency['name']}] {res['error']} → историята остава непроменена")
        return res

    df_all, new, changed, sold, bulk = update_history(agency, scraped, today)
    # Отпечатъци на снимките (само за обявите без тях) — за откриване на дубликати
    df_all = duplicates.fill_image_hashes(df_all, COL_IMAGES, duplicates.url_reader(_session()))
    res.update(df=df_all, new=new, changed=changed, sold=sold, fetched=len(scraped), bulk=bulk)

    AGENCY_DIR.mkdir(exist_ok=True)
    secure_store.write_parquet(df_all, history_path(agency["key"]))
    logger.info(f"[{agency['name']}] Изтеглени: {len(scraped)} | Нови: {len(new)}"
                f"{' (първоначално зареждане)' if res['bulk'] and len(new) else ''} | "
                f"Промени: {len(changed)} | Свалени: {len(sold)} | Общо в историята: {len(df_all)}")
    return res


def run_all(today=None):
    """Тегли всички агенции. Никога не хвърля грешка — проблемите са в res['error']."""
    today = today or datetime.now().strftime("%Y-%m-%d")
    results = []
    for agency in AGENCIES:
        logger.info(f"=== {agency['name']} ===")
        try:
            results.append(run_agency(agency, today))
        except Exception as agency_err:
            logger.exception(f"[{agency['name']}] Неочаквана грешка: {agency_err}")
            results.append({
                "key": agency["key"], "name": agency["name"], "site": agency["site"],
                "url": agency["url"], "error": str(agency_err)[:200],
                "fetched": 0, "bulk": False, "df": finalize(load_history(agency["key"]), agency),
                "new": pd.DataFrame(), "changed": pd.DataFrame(), "sold": pd.DataFrame(),
            })
    return results


def has_events(results):
    """Има ли нещо за имейла: нови (без първоначално зареждане), промени или свалени."""
    return any(
        (len(r["new"]) and not r["bulk"]) or len(r["changed"]) or len(r["sold"])
        for r in results or []
    )


if __name__ == "__main__":
    # Самостоятелно пускане: python agencies.py (нужни са DASHBOARD_PASSWORD и *_URL)
    secure_store.require_password()
    logging.basicConfig(level=logging.INFO, format='%(asctime)s | %(levelname)-7s | %(message)s')
    for r in run_all():
        print(f"{r['name']}: изтеглени {r['fetched']}, грешка: {r['error']}")
