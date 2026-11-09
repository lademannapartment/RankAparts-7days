"""
Скан Rank AP7d на 7 ночей.
Логика скролла/поиска карточек — из ra.py.
Запись — в лист "Rank AP7d" (Гданьск: I..AF + MMRent в C, Гдыня: AP, AQ).
"""
import asyncio
import datetime as dt
import gc as py_gc
import json
import os
import random
import re
import subprocess
import sys
import time
from typing import Dict, Tuple
from zoneinfo import ZoneInfo

import gspread
from google.oauth2.service_account import Credentials
from playwright.async_api import async_playwright, Page

# --- КОНСТАНТЫ ---
TZ = ZoneInfo("Europe/Warsaw")
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1t8jZFnJ5PxW9ry8SPN8WAGhbsgICn5F9BZCR9d4vB5U/edit"
SHEET_NAME_RANK = "Rank AP7d"
# Ключ лежит в папке RankApart (на уровень выше этой папки)
CREDENTIALS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..",
                                "silken-glyph-443313-i4-cdf14b36288a.json")
STAY_NIGHTS = 7           # ← 7 ночей
DAYS_FORWARD = 60
MAX_RANK = 400
MIN_CARDS = 10            # меньше карточек = страница не загрузилась, ячейки не трогаем
MIN_CARDS_TO_CLEAR = 50   # "не найдено" очищаем только если список загружен полностью
SCAN_PAUSE = (1, 3)       # пауза между городами/датами, сек
CARD_SELECTOR = 'div[data-testid="property-card"]'
HEADLESS = os.environ.get("HEADLESS", "0") == "1"  # на Mac по умолчанию окно видно
USER_AGENT = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36")

# Гдыня: AP, AQ
LISTING_COLUMNS_GDYNIA = {
    "Bohos Apartment - 2 min to City Center Gdynia & Seaside": 42,
    "Sailor Apartment- 10 minutes to Seaside & City Center": 43,
}

# Гданьск: I (9) ... AF (32)
LISTING_COLUMNS_GDANSK = {
    "Swish Apartment -10 Min to Old Town Gdańsk,PKP & Shopping Mall": 9,
    "Wave Apartment with Amazing View- 2 km to Baltic Sea": 10,
    "Anchor Apartment Your Docking Point in the Heart of Gdańsk": 11,
    "Boho Apartment - 10 Min to Old Town Gdańsk,PKP & Shopping Mall": 12,
    "Luna Premium Apartment at Old Town Gdansk": 13,
    "Libertas Premium Apartment at Old Town Gdansk": 14,
    "Panorama Apartment with Amazing View- 2 km to Baltic Sea, 15 Minutes to Gdańsk Old Town": 15,
    "Old Town Heaven Apartment": 16,
    "Seredino Blue Apartment- Modern Comfort in a Quiet Location, 5 Minutes to Gdańsk Old Town & Seaside": 17,
    "Seredino Green Apartment- Modern Comfort in a Quiet Location, 5 Minutes to Gdańsk Old Town & Seaside": 18,
    "Seredino Relax Apartment- Modern Comfort in a Quiet Location, 5 Minutes to Gdańsk Old Town & Seaside": 19,
    "Seredino Navy Apartment- Modern Comfort in a Quiet Location, 5 Minutes to Gdańsk Old Town & Seaside": 20,
    "Laura Apartment with Sauna & Gym, 5 Min to Baltic Sea & Polsat Plus Arena": 21,
    "Bloom Apartment with Sauna & Gym, 5 Min to Baltic Sea & Polsat Plus Arena": 22,
    "Seashell Apartment - at Gdańsk Stogi, 5 min by car to Baltic Sea": 23,
    "Coconut Apartment - 5 min to Baltic Sea": 24,
    "Granaria Apartment - Heart of Motlawa Promenade at Old Town Gdansk": 25,
    "Casa Verde Apartment at Old Town Gdańsk": 26,
    "Casa Pluma Apartment at Old Town Gdańsk": 27,
    "Rout Studio 5 min to Old Town Gdansk": 28,
    "Queem Studio 10 min to Old Town Gdansk": 29,
    "Avenue 77 Apartment - 15 Minutes to Seaside": 30,
    "Honey Apartment- Cozy Retreat, 10 Minutes to Gdańsk Old Town": 31,
    "Birch Studio - Peaceful Retreat, 20 Minutes to Gdańsk Old Town & Seaside": 32,
}

# Параметры поиска без меток трекинга (label/aid/gclid)
CITIES = [
    {
        "name": "Gdansk",
        "base": "ss=Gdańsk&dest_id=-501400&dest_type=city&order=price&nflt=ht_id%3D201",
        "columns": LISTING_COLUMNS_GDANSK,
        "write_mmrent": True,
    },
    {
        "name": "Gdynia",
        "base": "dest_id=-501414&dest_type=city&order=price&nflt=ht_id%3D201",
        "columns": LISTING_COLUMNS_GDYNIA,
        "write_mmrent": False,
    },
]

# Сторонние трекеры и реклама, которые не нужны для скрапинга
BLOCKED_HOSTS = (
    "google-analytics.com", "googletagmanager.com", "doubleclick.net", "googlesyndication.com",
    "facebook.net", "facebook.com", "criteo", "hotjar", "scorecardresearch", "taboola", "outbrain",
)

gc = None
_ws = None


# ------------------ ВСПОМОГАТЕЛЬНЫЕ ------------------
def col_to_letter(col: int) -> str:
    result = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        result = chr(65 + rem) + result
    return result


def clean_text(text: str) -> str:
    return re.sub(r'[^a-zA-Zа-яА-Я0-9]', '', text).lower() if text else ""


def authorize_gspread():
    global gc
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds_json_str = os.environ.get("GOOGLE_CREDS")
    if creds_json_str:
        creds = Credentials.from_service_account_info(json.loads(creds_json_str), scopes=scopes)
    else:
        creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
    gc = gspread.authorize(creds)


def get_worksheet():
    global _ws
    if _ws is None:
        if gc is None:
            authorize_gspread()
        _ws = gc.open_by_url(SPREADSHEET_URL).worksheet(SHEET_NAME_RANK)
    return _ws


def build_date_rows(ws) -> Dict[dt.date, int]:
    """Одно чтение колонки A: дата -> номер строки."""
    rows = {}
    for idx, cell in enumerate(ws.col_values(1), start=1):
        cell = cell.strip()
        for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
            try:
                rows.setdefault(dt.datetime.strptime(cell, fmt).date(), idx)
                break
            except ValueError:
                continue
    return rows


def build_url(base: str, checkin: dt.date, checkout: dt.date) -> str:
    return (f"https://www.booking.com/searchresults.pl.html?{base}"
            f"&checkin={checkin:%Y-%m-%d}&checkout={checkout:%Y-%m-%d}")


# ------------------ СКРЕЙПИНГ (логика из ra.py) ------------------
async def wait_for_more_cards(page: Page, before: int, timeout_ms: int = 15000) -> bool:
    try:
        await page.wait_for_function(
            "n => document.querySelectorAll('div[data-testid=\"property-card\"]').length > n",
            arg=before, timeout=timeout_ms)
        return True
    except Exception:
        return False


async def scrape_cards_and_get_ranks(page: Page, listings_map: dict) -> Tuple[dict, int, int]:
    found_ranks = {title: None for title in listings_map.keys()}
    remaining = set(listings_map.keys())
    mmrent_count = 0
    seen_cards_count = 0

    print(f"   📈 Поиск объектов (цель: до {MAX_RANK})...")

    for _ in range(40):  # итерации скроллинга
        cards = await page.query_selector_all(CARD_SELECTOR)
        current_count = len(cards)

        if current_count > seen_cards_count:
            print(f"      [Страница] Загружено карточек: {current_count}")

        # Обрабатываем только новые карточки
        for rank, card in enumerate(cards[seen_cards_count:], start=seen_cards_count + 1):
            if rank > MAX_RANK:
                break
            try:
                element = await card.query_selector('div[data-testid="title"]')
                if not element:
                    continue
                title_clean = clean_text(await element.inner_text())
                if not title_clean:
                    continue

                if "mmrent" in title_clean:
                    mmrent_count += 1

                for listing_name in list(remaining):
                    if clean_text(listing_name) in title_clean or title_clean in clean_text(listing_name):
                        found_ranks[listing_name] = rank
                        remaining.remove(listing_name)
                        print(f"   ✅ НАЙДЕНО: #{rank} {listing_name[:20]}...")
            except Exception:
                continue

        seen_cards_count = current_count

        if not remaining:
            print(f"   🎯 Все объекты найдены (загружено карточек: {seen_cards_count}).")
            break

        if seen_cards_count >= MAX_RANK:
            print(f"   ⏹ Достигнут лимит в {MAX_RANK} карточек.")
            break

        # Подгрузка: кнопка или прокрутка, затем ждём новые карточки
        try:
            show_more = page.locator('button:has-text("Załaduj"), button:has-text("Show more")').first
            if await show_more.is_visible(timeout=1000):
                await show_more.click(force=True)
            else:
                await page.evaluate("window.scrollTo(0, document.body.scrollHeight);")
            if not await wait_for_more_cards(page, current_count):
                break  # новых карточек нет — выдача закончилась
        except Exception:
            break

    return found_ranks, mmrent_count, seen_cards_count


async def block_aggressively(route):
    req = route.request
    if req.resource_type in ["image", "media", "font"] or any(h in req.url for h in BLOCKED_HOSTS):
        await route.abort()
    else:
        await route.continue_()


# ------------------ ЗАПИСЬ ------------------
def write_city(ws, row: int, city: dict, ranks: dict, mmrent: int, seen: int):
    updates = []
    if city["write_mmrent"]:
        updates.append({"range": f"C{row}", "values": [[mmrent]]})

    for title, col in city["columns"].items():
        rank = ranks.get(title)
        if rank is None and seen < MIN_CARDS_TO_CLEAR:
            print(f"   ⏭️ {title[:20]}: не найден, список неполный ({seen}) — старое значение оставлено")
            continue
        updates.append({"range": f"{col_to_letter(col)}{row}", "values": [[rank if rank is not None else ""]]})
        print(f"   {'✅' if rank else '🗑️'} {title[:20]}: {'#' + str(rank) if rank else 'очищено'}")

    if updates:
        ws.batch_update(updates, value_input_option="USER_ENTERED")


# ------------------ СКАН ГОРОДА ------------------
async def scan_city(context, city: dict, ws, row: int, date_obj: dt.date) -> bool:
    checkin = date_obj
    checkout = date_obj + dt.timedelta(days=STAY_NIGHTS)
    url = build_url(city["base"], checkin, checkout)

    for attempt in range(1, 3):
        page = await context.new_page()
        try:
            await page.goto(url, timeout=90000, wait_until="domcontentloaded")
            try:
                await page.wait_for_selector(CARD_SELECTOR, timeout=20000)
            except Exception:
                pass

            try:
                btn = page.locator('button#onetrust-accept-btn-handler, button:has-text("Akceptuj")').first
                if await btn.is_visible(timeout=1000):
                    await btn.click()
            except Exception:
                pass

            ranks, mmrent, seen = await scrape_cards_and_get_ranks(page, city["columns"])
            print(f"   🃏 {city['name']}: карточек просмотрено {seen}")
            if seen < MIN_CARDS:
                raise RuntimeError(f"загружено только {seen} карточек — похоже на блокировку")

            await asyncio.to_thread(write_city, ws, row, city, ranks, mmrent, seen)
            return True
        except Exception as e:
            print(f"   ⚠️ {city['name']} попытка {attempt}/2: {e}")
            await asyncio.sleep(5)
        finally:
            await page.close()

    print(f"   ❌ {city['name']} пропущен, строка {row} не изменена")
    return False


# ------------------ MAIN ------------------
async def main_async():
    print("Checking/Installing Playwright browser...")
    subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=False)

    ws = get_worksheet()
    ws.update_cell(1, 1, f"📅 Skan: {dt.datetime.now(TZ):%d.%m.%Y %H:%M:%S}")
    date_rows = build_date_rows(ws)
    today = dt.datetime.now(TZ).date()

    async with async_playwright() as p:
        # HEADLESS=0 — окно браузера видно (Booking реже блокирует обычный браузер, чем headless)
        # BROWSER_CHANNEL=chrome — использовать установленный Chrome (на Mac), по умолчанию Chromium Playwright
        channel = os.environ.get("BROWSER_CHANNEL") or None
        browser = await p.chromium.launch(
            headless=HEADLESS,
            channel=channel,
            # убираем флаг автоматизации, чтобы Booking не отдавал страницу города вместо выдачи
            ignore_default_args=["--enable-automation"],
            args=["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage", "--disable-gpu",
                  "--disable-blink-features=AutomationControlled"],
        )
        context = await browser.new_context(
            user_agent=USER_AGENT,
            locale="pl-PL",
            timezone_id="Europe/Warsaw",
            viewport={"width": 1400, "height": 900},
            service_workers="block",
        )
        # скрываем признак WebDriver у navigator
        await context.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined});")
        await context.route("**/*", block_aggressively)

        for i in range(DAYS_FORWARD):
            date_obj = today + dt.timedelta(days=i)
            row = date_rows.get(date_obj)
            if not row:
                print(f"⚠️ {date_obj:%d.%m.%Y} нет в таблице, пропуск")
                continue

            print(f"\n>>>> {date_obj:%d.%m.%Y} (строка {row}), {STAY_NIGHTS} noce <<<<")
            for city in CITIES:
                await asyncio.sleep(random.uniform(*SCAN_PAUSE))
                print(f"   --- {city['name']} ---")
                await scan_city(context, city, ws, row, date_obj)

            py_gc.collect()

        await browser.close()
    print("\n🏁 Все даты обработаны.")


if __name__ == "__main__":
    asyncio.run(main_async())
