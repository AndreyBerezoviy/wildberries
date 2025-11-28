import time
import random
from typing import Optional

import requests
import pandas as pd
from bs4 import BeautifulSoup

BASE_URL = "https://www.wildberries.ru"
SEARCH_QUERY = "пальто из натуральной шерсти"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/121.0.0.0 Safari/537.36"
    ),
    "Accept": (
        "text/html,application/xhtml+xml,application/xml;q=0.9,"
        "image/avif,image/webp,*/*;q=0.8"
    ),
    "Accept-Language": "ru-RU,ru;q=0.9",
    "Connection": "keep-alive",
}


def fetch_html(
    session: requests.Session,
    url: str,
    params: Optional[dict] = None,
) -> Optional[BeautifulSoup]:
    try:
        resp = session.get(url, params=params, headers=HEADERS, timeout=20)
    except requests.RequestException:
        return None

    if not (200 <= resp.status_code < 300):
        return None

    return BeautifulSoup(resp.text, "html.parser")


def load_search_page(
    session: requests.Session,
    query: str,
    page: int,
) -> Optional[BeautifulSoup]:
    url = f"{BASE_URL}/catalog/0/search.aspx"
    params = {
        "search": query,
        "page": page,
        "sort": "popular",
    }
    return fetch_html(session, url, params)


def extract_cards_from_search_page(soup: BeautifulSoup) -> list[dict]:
    result: list[dict] = []
    cards = soup.select("article.product-card, div.product-card")

    for card in cards:
        link = card.select_one(
            "a.product-card__main, a.product-card__link, a.j-open-full-product-card"
        )
        if not link:
            continue

        href = link.get("href")
        if not href:
            continue

        if href.startswith("/"):
            href = BASE_URL + href

        nm_id = card.get("data-nm-id")

        result.append(
            {
                "url": href,
                "nm_id": nm_id,
            }
        )

    return result


def parse_product_page(
    session: requests.Session,
    url: str,
    nm_id: Optional[str],
) -> Optional[dict]:
    soup = fetch_html(session, url)
    if soup is None:
        return None

    # Название
    title_el = soup.select_one("h1.product-page__title, h1")
    title = title_el.get_text(strip=True) if title_el else ""

    # Артикул
    article = nm_id
    article_label = soup.find(string=lambda t: isinstance(t, str) and "Артикул" in t)
    if article_label and article_label.parent:
        span = article_label.parent.find("span")
        if span:
            article = span.get_text(strip=True)

    # Цена
    price_val: Optional[int] = None
    price_el = soup.select_one(
        ".price-block__final-price, .final-cost, .price-block__price"
    )
    if price_el:
        digits = "".join(ch for ch in price_el.get_text(strip=True) if ch.isdigit())
        if digits.isdigit():
            price_val = int(digits)

    # Описание
    desc_el = soup.select_one("#description, .product-description__text")
    description = desc_el.get_text(" ", strip=True) if desc_el else ""

    # Картинки
    image_urls: list[str] = []
    for img in soup.select(
        "div.product-page__gallery img, .photo-zoom__preview img"
    ):
        src = img.get("src") or img.get("data-src")
        if not src:
            continue
        if src.startswith("//"):
            src = "https:" + src
        elif src.startswith("/"):
            src = BASE_URL + src
        if src not in image_urls:
            image_urls.append(src)

    # Характеристики
    characteristics_block = soup.select_one(
        ".product-params, .product-characteristics, #characteristics"
    )
    characteristics: list[str] = []
    country = ""

    if characteristics_block:
        rows = characteristics_block.select("tr, li")
        for row in rows:
            name_el = row.select_one("th, .product-params__cell-name")
            value_el = row.select_one("td, .product-params__cell-value")
            if not name_el or not value_el:
                continue

            name = name_el.get_text(strip=True)
            value = value_el.get_text(strip=True)
            if not name or not value:
                continue

            characteristics.append(f"{name}: {value}")

            if "страна" in name.lower():
                country = value

    # Продавец
    seller_name = ""
    seller_url = ""
    seller_el = soup.select_one("a.seller-info__name, a.seller-link")
    if seller_el:
        seller_name = seller_el.get_text(strip=True)
        href = seller_el.get("href")
        if href:
            seller_url = BASE_URL + href if href.startswith("/") else href

    # Размеры
    sizes_block = soup.select_one(".sizes-list, .sizes-table")
    sizes: list[str] = []
    if sizes_block:
        for el in sizes_block.select("button, .sizes-list__item"):
            txt = el.get_text(" ", strip=True)
            if txt:
                sizes.append(txt)

    # Рейтинг
    rating: Optional[float] = None
    rating_el = soup.select_one(".product-page__rating, span.rating")
    if rating_el:
        txt = rating_el.get_text(strip=True).replace(",", ".")
        filtered = "".join(c for c in txt if c.isdigit() or c == ".")
        try:
            rating = float(filtered)
        except ValueError:
            rating = None

    # Количество отзывов
    feedbacks: Optional[int] = None
    fb_el = soup.find(string=lambda t: isinstance(t, str) and "отзыв" in t.lower())
    if fb_el:
        digits = "".join(ch for ch in fb_el if ch.isdigit())
        if digits.isdigit():
            feedbacks = int(digits)

    return {
        "Ссылка на товар": url,
        "Артикул": article,
        "Название": title,
        "Цена": price_val,
        "Описание": description,
        "Ссылки на изображения через запятую": ", ".join(image_urls),
        "Все характеристики с сохранением их структуры": "; ".join(characteristics),
        "Название селлера": seller_name,
        "Ссылка на селлера": seller_url,
        "Размеры товара через запятую": ", ".join(sizes),
        "Остатки по товару (число)": None,  # без доп. API WB общий остаток не отдаёт
        "Рейтинг": rating,
        "Количество отзывов": feedbacks,
        "Страна производства": country,  # служебное поле для фильтрации
    }


def main() -> None:
    session = requests.Session()
    items: list[dict] = []

    # Проходим по нескольким страницам выдачи
    for page in range(1, 6):
        soup = load_search_page(session, SEARCH_QUERY, page)
        if soup is None:
            break

        cards = extract_cards_from_search_page(soup)
        if not cards:
            break

        for card in cards:
            data = parse_product_page(session, card["url"], card["nm_id"])
            if data:
                items.append(data)

            # небольшая задержка, чтобы не долбить сайт
            time.sleep(random.uniform(1.0, 2.5))

    if not items:
        print("Данные не собраны (скорее всего, WB режет запросы с этого IP).")
        return

    df = pd.DataFrame(items)

    # основной файл по ТЗ
    df_for_export = df.drop(columns=["Страна производства"])
    df_for_export.to_excel("wb_coats_full.xlsx", index=False)

    # выборка
    filtered = df[
        (df["Цена"].notna())
        & (df["Рейтинг"].notna())
        & (df["Рейтинг"] >= 4.5)
        & (df["Цена"] <= 10_000)
        & (df["Страна производства"].str.contains("Россия", na=False))
    ]

    filtered_for_export = filtered.drop(columns=["Страна производства"])
    filtered_for_export.to_excel("wb_coats_filtered.xlsx", index=False)

    print("Готово: wb_coats_full.xlsx и wb_coats_filtered.xlsx")


if __name__ == "__main__":
    main()
