import asyncio
import datetime as dt
import json
import os
import sys
import time
import re
import gc as py_gc
import gspread
from zoneinfo import ZoneInfo
from google.oauth2.service_account import Credentials
from playwright.async_api import async_playwright, Page
from typing import Tuple, Dict

# --- КОНСТАНТЫ ---
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1t8jZFnJ5PxW9ry8SPN8WAGhbsgICn5F9BZCR9d4vB5U/edit"
SHEET_NAME_RANK = "Rank AP7d"  # Общее название для листа
CREDENTIALS_FILE = "silken-glyph-443313-i4-cdf14b36288a.json"
STAY_NIGHTS = 7
DAYS_FORWARD = 60
MAX_RANK = 400
MIN_CARDS = 10  # меньше карточек = страница не загрузилась (блокировка), ячейки не трогаем

# Гданьск (Колонки I–AL)
LISTING_COLUMNS_GDANSK_APARTMENTS = {
    "Swish Apartment -10 Min to Old Town Gdańsk,PKP & Shopping Mall": 9,
    "Milky Apartment with Balcony - 3 km to Baltic Sea": 10,
    "Homely Apartment with Balcony - 3 km to Baltic Sea": 11,
    "Wine Apartment - 5 min to Old Town Gdańsk": 12,
    "Wave Apartment with Amazing View- 2 km to Baltic Sea": 13,
    "Anchor Apartment Your Docking Point in the Heart of Gdańsk": 14,
    "Boho Apartment - 10 Min to Old Town Gdańsk,PKP & Shopping Mall": 15,
    "Luna Premium Apartment at Old Town Gdansk": 16,
    "Libertas Premium Apartment at Old Town Gdansk": 17,
    "Postcard Apartment Scenic Escape 10 Minutes to Gdańsk Old Town and Seaside": 18,
    "Panorama Apartment with Amazing View- 2 km to Baltic Sea, 15 Minutes to Gdańsk Old Town": 19,
    "Old Town Heaven Apartment": 20,
    "Seredino Blue Apartment- Modern Comfort in a Quiet Location, 5 Minutes to Gdańsk Old Town & Seaside": 21,
    "Seredino Green Apartment- Modern Comfort in a Quiet Location, 5 Minutes to Gdańsk Old Town & Seaside": 22,
    "Seredino Relax Apartment- Romantic Escape with a Bathtub in the Bedroom, 5 Minutes to Gdańsk Old Town & Seaside": 23,
    "Seredino Navy Apartment- Modern Comfort in a Quiet Location, 5 Minutes to Gdańsk Old Town & Seaside": 24,
    "Laura Apartment with Sauna & Gym, 5 Min to Baltic Sea & Polsat Plus Arena": 25,
    "Coast Apartment with Sauna & Gym, 5 Min to Baltic Sea & Polsat Plus Arena": 26,
    "Golden Apartment - Premium Apartment at Old Town Gdansk": 27,
    "Amber Apartment - Premium Apartment at Old Town Gdansk": 28,
    "Bloom Apartment with Sauna & Gym, 5 Min to Baltic Sea & Polsat Plus Arena": 29,
    "Legit Apartment - Old Town Gdansk": 30,
    "Elite Apartment 15 min to Old Town Gdansk": 31,
    "Seashell Apartment - at Gdańsk Stogi, 5 min by car to Baltic Sea": 32,
    "Rout Studio 5 min to Old Town Gdansk": 33,
    "Queem Studio 10 min to Old Town Gdansk": 34,
    "Exclusive Apartment at Old Town Gdansk": 35,
    "Riverview Apartment at Old Town Gdańsk": 36,
    "Coconut Apartment - 5 min to Baltic Sea": 37,
    "Granaria Apartment - Heart of Motlawa Promenade at Old Town Gdansk": 38
}

# Гдыня (Колонки AP, AQ)
LISTING_COLUMNS_GDYNIA_APARTMENTS = {
    "Bohos Apartment - 2 min to City Center Gdynia & Seaside": 42,  # AP
    "Sailor Apartment- 10 minutes to Seaside & City Center": 43  # AQ
}

gc = None


# ------------------ ВСПОМОГАТЕЛЬНЫЕ ------------------
def col_to_letter(col: int) -> str:
    result = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        result = chr(65 + rem) + result
    return result


def authorize_gspread():
    global gc
    creds_json_str = os.environ.get("GOOGLE_CREDS")
    if creds_json_str:
        creds_json = json.loads(creds_json_str)
        creds = Credentials.from_service_account_info(
            creds_json, scopes=["https://www.googleapis.com/auth/spreadsheets"]
        )
    else:
        creds = Credentials.from_service_account_file(
            CREDENTIALS_FILE, scopes=["https://www.googleapis.com/auth/spreadsheets"]
        )
    gc = gspread.authorize(creds)


def initialize_gspread(sheet_name: str):
    if gc is None: authorize_gspread()
    try:
        sh = gc.open_by_url(SPREADSHEET_URL)
        return sh.worksheet(sheet_name)
    except Exception as e:
        print(f"❌ Ошибка листа {sheet_name}: {e}")
        return None


def find_row_by_date(ws, target_date: dt.date, retries: int = 3):
    for _ in range(retries):
        try:
            values = ws.col_values(1)
            for idx, cell in enumerate(values, start=1):
                if not cell.strip(): continue
                for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
                    try:
                        if dt.datetime.strptime(cell.strip(), fmt).date() == target_date:
                            return idx
                    except ValueError:
                        continue
            return None
        except:
            time.sleep(2)
    return None


# ------------------ ОБНОВЛЕНИЕ ДАННЫХ ------------------
def update_spreadsheet_data(ws, row_index, ranks_map, listings_map, mmrent_count, update_mmrent=True):
    SLEEP_TIME = 1.2

    # Запись MMRent в колонку C (только если нужно, например только при скане Гданьска)
    if update_mmrent:
        try:
            ws.update(values=[[mmrent_count]], range_name=f"C{row_index}")
            print(f"   📊 MMRent: {mmrent_count} карточек")
        except Exception as e:
            print(f"   ❌ Ошибка записи MMRent в C{row_index}: {e}")

    # Запись рангов
    for title, col_idx in listings_map.items():
        raw_rank = ranks_map.get(title)
        cell = f"{col_to_letter(col_idx)}{row_index}"
        try:
            ws.update(range_name=cell, values=[[raw_rank]] if raw_rank else [[""]], value_input_option='USER_ENTERED')
            print(f"   {'✅' if raw_rank else '🗑️'} {title[:20]}: {'#' + str(raw_rank) if raw_rank else 'очищено'}")
        except Exception as e:
            print(f"   ❌ Ошибка записи {cell}: {e}")
        time.sleep(SLEEP_TIME)


# ------------------ СКРЕЙПИНГ ------------------
def build_gdansk_apartments_url(checkin: dt.date, checkout: dt.date) -> str:
    checkin_str, checkout_str = checkin.strftime("%Y-%m-%d"), checkout.strftime("%Y-%m-%d")
    base = ("ss=Gdańsk&ssne=Gdańsk&ssne_untouched=Gdańsk&dest_id=-501400&dest_type=city"
            "&lang=pl&src=searchresults&order=price&nflt=ht_id%3D201"
            "&group_adults=2&no_rooms=1&group_children=0")
    return f"https://www.booking.com/searchresults.pl.html?{base}&checkin={checkin_str}&checkout={checkout_str}"


def build_gdynia_apartments_url(checkin: dt.date, checkout: dt.date) -> str:
    checkin_str, checkout_str = checkin.strftime("%Y-%m-%d"), checkout.strftime("%Y-%m-%d")
    base = ("dest_id=-501414&dest_type=city&lang=pl&src=searchresults&order=price&nflt=ht_id%3D201"
            "&group_adults=2&no_rooms=1&group_children=0")
    return f"https://www.booking.com/searchresults.pl.html?{base}&checkin={checkin_str}&checkout={checkout_str}"


async def wait_for_more_cards(page: Page, before: int, timeout_ms: int = 8000) -> bool:
    try:
        await page.wait_for_function(
            "n => document.querySelectorAll('div[data-testid=\"property-card\"]').length > n",
            arg=before, timeout=timeout_ms)
        return True
    except Exception:
        return False


async def scrape_cards_and_get_ranks(page: Page, listings_map: dict) -> Tuple[dict, int, int]:
    PROPERTY_CARD_SELECTOR = 'div[data-testid="property-card"]'

    def clean_text(text: str) -> str:
        return re.sub(r'[^a-zA-Zа-яА-Я0-9]', '', text).lower() if text else ""

    found_ranks = {title: None for title in listings_map.keys()}
    remaining = set(listings_map.keys())
    mmrent_count = 0
    seen_cards_count = 0

    print(f"   📈 Поиск объектов (цель: до {MAX_RANK})...")

    for _ in range(40):  # Итерации скроллинга
        cards = await page.query_selector_all(PROPERTY_CARD_SELECTOR)
        current_count = len(cards)

        if current_count > seen_cards_count:
            print(f"      [Страница] Загружено карточек: {current_count}")

        # Обрабатываем только новые карточки
        for rank, card in enumerate(cards[seen_cards_count:], start=seen_cards_count + 1):
            if rank > MAX_RANK: break
            try:
                element = await card.query_selector('div[data-testid="title"]')
                if not element: continue
                title_clean = clean_text(await element.inner_text())
                if not title_clean: continue

                if "mmrent" in title_clean:
                    mmrent_count += 1

                for listing_name in list(remaining):
                    if clean_text(listing_name) in title_clean or title_clean in clean_text(listing_name):
                        found_ranks[listing_name] = rank
                        remaining.remove(listing_name)
                        print(f"   ✅ НАЙДЕНО: #{rank} {listing_name[:20]}...")
            except:
                continue

        seen_cards_count = current_count

        if not remaining:
            print(f"   🎯 Все объекты найдены (загружено карточек: {seen_cards_count}). Останавливаю скролл.")
            break

        if seen_cards_count >= MAX_RANK:
            print(f"   ⏹ Достигнут лимит в {MAX_RANK} карточек.")
            break

        # Скроллим дальше: ждём появления новых карточек, а не фиксированную паузу
        try:
            show_more_button = page.locator('button:has-text("Załaduj"), button:has-text("Show more")').first
            if await show_more_button.is_visible(timeout=1000):
                await show_more_button.click(force=True)
            else:
                await page.evaluate("window.scrollTo(0, document.body.scrollHeight);")
            if not await wait_for_more_cards(page, current_count):
                break  # новых карточек нет — выдача закончилась
        except:
            break

    return found_ranks, mmrent_count, seen_cards_count


async def process_location(page: Page, city_name: str, ws: gspread.Worksheet, url_builder: callable, listings_map: dict,
                           date_obj: dt.date, row_index: int, update_mmrent: bool):
    print(f"   --- {city_name} ---")
    url = url_builder(date_obj, date_obj + dt.timedelta(days=STAY_NIGHTS))

    async def block_aggressively(route):
        if route.request.resource_type in ["image", "media", "font"]:
            await route.abort()
        else:
            await route.continue_()

    try:
        await page.route("**/*", block_aggressively)
        await page.goto(url, timeout=100000, wait_until="domcontentloaded")
        await asyncio.sleep(2)

        # Куки
        try:
            cookie_btn = page.locator('button#onetrust-accept-btn-handler')
            if await cookie_btn.is_visible(timeout=5000):
                await cookie_btn.click()
        except:
            pass

        # Ждём появления карточек (до 30 сек), а не фиксированные 2 секунды
        try:
            await page.wait_for_selector('div[data-testid="property-card"]', timeout=30000)
        except Exception:
            pass

        ranks_map, mmrent_count, cards_seen = await scrape_cards_and_get_ranks(page, listings_map)
        print(f"   🃏 Карточек просмотрено: {cards_seen}")
        if cards_seen < MIN_CARDS:
            # Мало карточек = страница не загрузилась / блокировка: не трогаем таблицу
            print(f"⚠️ {city_name}: загружено только {cards_seen} карточек (нужно {MIN_CARDS}), строку {row_index} не меняю")
            try:
                print(f"   URL: {page.url}")
                print(f"   Title: {await page.title()}")
                body_text = (await page.inner_text("body"))[:200]
                print(f"   Текст страницы: {body_text!r}")
                shot = f"debug_{city_name}_{date_obj}.png"
                await page.screenshot(path=shot)
                print(f"   Скриншот: {os.path.abspath(shot)}")
            except Exception as dbg_err:
                print(f"   ❌ Не удалось собрать диагностику: {dbg_err}")
        else:
            update_spreadsheet_data(ws, row_index, ranks_map, listings_map, mmrent_count, update_mmrent)
        await page.unroute("**/*")
    except Exception as e:
        print(f"❌ Ошибка {city_name} на {date_obj}: {e}")


# ------------------ MAIN ------------------
import subprocess


async def main_async():
    subprocess.run(["python", "-m", "playwright", "install", "chromium"])
    authorize_gspread()
    ws = initialize_gspread(SHEET_NAME_RANK)

    now_str = dt.datetime.now(ZoneInfo("Europe/Warsaw")).strftime("%d.%m.%Y %H:%M:%S")
    if ws: ws.update_cell(1, 1, f"📅 Skan: {now_str}")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0",
            locale="pl-PL"
        )

        today = dt.datetime.now(ZoneInfo("Europe/Warsaw")).date()
        for i in range(DAYS_FORWARD):
            date_obj = today + dt.timedelta(days=i)
            row_index = find_row_by_date(ws, date_obj)

            if not row_index:
                print(f"⚠️ Дата {date_obj} не найдена в таблице. Пропуск.")
                continue

            print(f"\n>>>> ОБРАБОТКА: {date_obj.strftime('%d.%m.%Y')} (Строка {row_index}) <<<<")

            # --- ГДАНЬСК ---
            page_gdansk = await context.new_page()
            await process_location(page_gdansk, "Gdansk", ws, build_gdansk_apartments_url,
                                   LISTING_COLUMNS_GDANSK_APARTMENTS, date_obj, row_index, update_mmrent=True)
            await page_gdansk.close()

            # --- ГДЫНЯ ---
            page_gdynia = await context.new_page()
            await process_location(page_gdynia, "Gdynia", ws, build_gdynia_apartments_url,
                                   LISTING_COLUMNS_GDYNIA_APARTMENTS, date_obj, row_index, update_mmrent=False)
            await page_gdynia.close()

            py_gc.collect()
            print(f"🧹 ОЗУ очищена после {date_obj}")

        await browser.close()
    print("\n🏁 Все города и даты обработаны.")


if __name__ == "__main__":
    asyncio.run(main_async())

