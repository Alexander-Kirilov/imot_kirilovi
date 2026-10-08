"""Един и същ апартамент в няколко обяви — по снимките.

Всяка снимка получава "отпечатък" (dHash, 64 бита): една и съща снимка дава почти
същия отпечатък и след преоразмеряване/прекомпресиране или чужд воден знак.
Отпечатъците се смятат веднъж за обява и се пазят в колона Image_Hashes.

Две обяви са един апартамент, ако имат обща снимка (не само план/чертеж) и
съвпадат кварталът, площта (±5%) и етажът (ако е известен и при двете).
Снимки, които се срещат в много обяви (лога, рендъри на нов блок), не се броят.
"""
import io
import logging
import re
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import requests
from PIL import Image

logger = logging.getLogger(__name__)

COL_IMAGE_HASHES = 'Image_Hashes'  # "a1b2…f0p,…" — 16 hex на снимка; "p" = план/чертеж; "-" = няма снимки

HASH_MAX_DIST = 8        # от 64 бита
GENERIC_LISTINGS = 5     # снимка в повече обяви = обща → не се брои
SIZE_TOL = 0.05
REPOST_TOLERANCE_DAYS = 3

# Колони само за показване в dashboard-а (не се записват в историята)
COL_DUP_REPOSTED = 'Dup_Reposted'
COL_DUP_PREV = 'Dup_Prev'
COL_DUP_ALSO = 'Dup_Also'
COL_DUP_STILL = 'Dup_Still'
COL_DUP_CHANGE_DATE = 'Dup_Price_Change_Date'


# ================= ОТПЕЧАТЪЦИ =================

def signature(data):
    """Байтове на снимка → '16 hex' (+ 'p' ако е план/чертеж), или None."""
    try:
        img = Image.open(io.BytesIO(data)).convert("L")
    except Exception:
        return None
    small = img.resize((9, 8), Image.LANCZOS)
    px = small.tobytes()
    bits = 0
    for row in range(8):
        for col in range(8):
            bits = (bits << 1) | (px[row * 9 + col] > px[row * 9 + col + 1])
    # план/чертеж: предимно бял фон — като единствено съвпадение е слабо доказателство
    thumb = img.resize((64, 64)).tobytes()
    plan = sum(v > 225 for v in thumb) / len(thumb) > 0.5
    return f"{bits:016x}{'p' if plan else ''}"


def fill_image_hashes(df, image_col, read_image):
    """Попълва Image_Hashes за редовете без отпечатъци. read_image(път|URL) → bytes|None."""
    if df.empty or image_col not in df.columns:
        return df
    if COL_IMAGE_HASHES not in df.columns:
        df[COL_IMAGE_HASHES] = ""
    hashes = df[COL_IMAGE_HASHES].fillna("").astype(str)
    todo = hashes.str.strip() == ""
    done = 0
    for idx in df.index[todo]:
        sources = [s.strip() for s in str(df.at[idx, image_col] or "").split(",") if s.strip()]
        sigs = []
        for src in sources[:2]:
            data = read_image(src)
            sig = signature(data) if data else None
            if sig:
                sigs.append(sig)
        # "-" = опитано, без снимки → да не се тегли наново при всяко пускане
        df.at[idx, COL_IMAGE_HASHES] = ",".join(sigs) if sigs else ("-" if sources else "")
        done += bool(sources)
    if done:
        logger.info(f"Отпечатъци на снимки: {done} обяви")
    return df


def local_reader(base_dir):
    def read(path):
        p = Path(base_dir) / path
        return p.read_bytes() if p.exists() else None
    return read


def url_reader(session=None):
    s = session or requests.Session()
    def read(url):
        try:
            r = s.get(url, timeout=20, headers={"User-Agent": "Mozilla/5.0"})
            return r.content if r.status_code == 200 and len(r.content) > 2000 else None
        except requests.RequestException:
            return None
    return read


# ================= ГРУПИ =================

def ad_number(link):
    """Номерът на обява в imot.bg (без двубуквения префикс, който се сменя с вида)."""
    m = re.search(r"obiava-[a-z0-9]{2}(\d{10,})", str(link or ""))
    return m.group(1) if m else ""


def imot_created_date(link):
    """imot.bg: първите 10 цифри от номера са моментът на създаване (Unix време)."""
    num = ad_number(link)
    if not num:
        return ""
    try:
        return datetime.fromtimestamp(int(num[:10])).strftime("%Y-%m-%d")
    except (ValueError, OSError):
        return ""


def _num(v):
    try:
        f = float(v)
        return None if pd.isna(f) else f
    except (TypeError, ValueError):
        return None


def _text(v):
    s = str(v if v is not None else "").strip()
    return "" if s.lower() in ("nan", "none", "nat", "<na>") else s


def _first_price(hist, fallback):
    m = re.search(r"([\d,\s ]+)\s*€", str(hist or ""))
    if m:
        digits = re.sub(r"\D", "", m.group(1))
        if digits:
            return float(digits)
    return fallback


def _fmt_eur(x):
    return f"{int(round(x)):,} €".replace(",", " ") if x else "—"


def _ddmm(date):
    d = _text(date)[:10]
    return f"{d[8:10]}.{d[5:7]}" if len(d) == 10 else d


def _sigs(text):
    out = []
    for s in str(text or "").split(","):
        s = s.strip()
        if len(s) >= 16 and s != "-":
            try:
                out.append((int(s[:16], 16), s.endswith("p")))
            except ValueError:
                pass
    return out


def annotate(df, created_col, today=None):
    """Добавя колоните Dup_* към обединената таблица (всички сайтове).

    created_col — дата на създаване на обявата ('' ако сайтът не я дава).
    """
    today = today or datetime.now().strftime("%Y-%m-%d")
    for col, default in ((COL_DUP_REPOSTED, False), (COL_DUP_PREV, ""), (COL_DUP_ALSO, ""),
                         (COL_DUP_STILL, ""), (COL_DUP_CHANGE_DATE, "")):
        df[col] = default
    if df.empty or COL_IMAGE_HASHES not in df.columns:
        return df

    rows = df.reset_index(drop=True)
    images = [(i, h, plan) for i, text in enumerate(rows[COL_IMAGE_HASHES]) for h, plan in _sigs(text)]
    close = lambda a, b: (a ^ b).bit_count() <= HASH_MAX_DIST

    # Колко различни обяви "носят" всяка снимка → общите отпадат
    reach = [len({j for j, h2, _ in images if close(h, h2)}) for _, h, _ in images]
    strong_pairs = set()
    for a in range(len(images)):
        i, h1, p1 = images[a]
        if reach[a] > GENERIC_LISTINGS:
            continue
        for b in range(a + 1, len(images)):
            j, h2, p2 = images[b]
            if i != j and reach[b] <= GENERIC_LISTINGS and close(h1, h2) and not (p1 or p2):
                strong_pairs.add((min(i, j), max(i, j)))

    def same_flat(i, j):
        A, B = rows.iloc[i], rows.iloc[j]
        la, lb = _text(A.get("Location")), _text(B.get("Location"))
        if la and lb and la != lb:
            return False
        sa, sb = _num(A.get("Size_sqm")), _num(B.get("Size_sqm"))
        if sa and sb and abs(sa - sb) > max(SIZE_TOL * max(sa, sb), 3):
            return False
        fa, fb = _num(A.get("Floor")), _num(B.get("Floor"))
        return fa is None or fb is None or fa == fb

    parent = list(range(len(rows)))
    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x
    for i, j in strong_pairs:
        if same_flat(i, j):
            parent[find(i)] = find(j)

    groups = {}
    for i in range(len(rows)):
        groups.setdefault(find(i), []).append(i)

    def info(i):
        r = rows.iloc[i]
        return dict(i=i, site=_text(r.get("Source")) or "imot.bg", link=_text(r.get("Link")),
                    price=_num(r.get("Price_EUR")), sold=bool(r.get("Sold")),
                    created=_text(r.get(created_col))[:10], end=_text(r.get("Scraped_Date"))[:10],
                    first_price=_first_price(r.get("Price_History"), _num(r.get("Price_EUR"))),
                    number=ad_number(r.get("Link")))

    n = len(rows)
    out = {COL_DUP_REPOSTED: [False] * n, COL_DUP_PREV: [""] * n, COL_DUP_ALSO: [""] * n,
           COL_DUP_STILL: [""] * n, COL_DUP_CHANGE_DATE: [""] * n}
    for members in groups.values():
        if len(members) < 2:
            continue
        L = [info(i) for i in members]
        active = [x for x in L if not x["sold"]]
        sold = sorted([x for x in L if x["sold"]], key=lambda x: x["end"])
        for x in active:
            # друга активна обява за същия апартамент (друг сайт или друга агенция)
            others = [o for o in active if o is not x and (not x["number"] or o["number"] != x["number"])]
            if others:
                out[COL_DUP_ALSO][x["i"]] = " ".join(
                    f'<a class="dup dup-also" href="{o["link"]}" target="_blank" rel="noopener" '
                    f'title="Същият апартамент в друга обява">също: {o["site"]} {_fmt_eur(o["price"])}</a>'
                    for o in others[:3])
            # предишна (вече свалена) обява за същия апартамент
            before = [s for s in sold if s["end"] <= today]
            if before:
                prev = before[-1]
                change = prev["price"] and x["first_price"] and round(prev["price"]) != round(x["first_price"])
                delta = ""
                if change:
                    diff = x["first_price"] - prev["price"]
                    pct = 100 * diff / prev["price"]
                    arrow, cls = ("↓", "price-down") if diff < 0 else ("↑", "price-up")
                    sign = "−" if diff < 0 else "+"
                    txt = f"{sign}{_fmt_eur(abs(diff))}"
                    if abs(pct) >= 0.1:
                        txt += f" ({sign}{abs(pct):.1f}%)"
                    delta = f' <span class="{cls}">{arrow} {txt}</span>'
                out[COL_DUP_PREV][x["i"]] = (
                    f'<div class="dup-prev">Предишна обява: {_fmt_eur(prev["price"])}'
                    f' ({prev["site"]}, до {_ddmm(prev["end"])}){delta}</div>')
                limit = (pd.Timestamp(prev["end"]) - timedelta(days=REPOST_TOLERANCE_DAYS)).strftime("%Y-%m-%d") \
                    if prev["end"] else ""
                if x["created"] and limit and x["created"] >= limit:
                    out[COL_DUP_REPOSTED][x["i"]] = True
                if change:
                    out[COL_DUP_CHANGE_DATE][x["i"]] = x["created"] or prev["end"]
        # свалена обява, но апартаментът още се продава в друга обява
        for s in sold:
            if active:
                a = max(active, key=lambda o: o["created"] or "")
                out[COL_DUP_STILL][s["i"]] = (
                    f'<a class="dup dup-still" href="{a["link"]}" target="_blank" rel="noopener" '
                    f'title="Не е продаден — качен е в друга обява">още се продава: {a["site"]} '
                    f'{_fmt_eur(a["price"])}</a>')

    for col, values in out.items():
        df[col] = pd.Series(values, index=df.index)
    n_rep = int(df[COL_DUP_REPOSTED].sum())
    n_also = int((df[COL_DUP_ALSO] != "").sum())
    logger.info(f"Дубликати по снимки: качени наново {n_rep}, в 2+ активни обяви {n_also}")
    return df
