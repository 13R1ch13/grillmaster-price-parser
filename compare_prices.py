import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
import re
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

OUR_URL = "https://grillmaster.dp.ua/hazovi-hryli/"
COMPETITOR_URL = "https://bbq24.com.ua/ua/gazovye-grili/"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/115.0 Safari/537.36"
    )
}

def clean_name(name):
    name = name.lower()
    name = re.sub(r"[^a-zа-я0-9\s]", "", name)
    return name.strip()

def parse_price(price_text):
    digits = re.sub(r"[^\d]", "", price_text)
    return int(digits) if digits else None

# ===== Универсальный парсер Grill Master =====
def parse_grillmaster():
    response = requests.get(OUR_URL, headers=HEADERS)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "html.parser")

    products = {}
    cards = soup.find_all(class_=lambda x: x and "product" in x)
    for card in cards:
        title_tag = card.find("h2")
        price_tag = card.find("span", class_=lambda x: x and "amount" in x)
        if title_tag and price_tag:
            title = clean_name(title_tag.get_text(strip=True))
            price = parse_price(price_tag.get_text(strip=True))
            if price:
                products[title] = price
    return products

# ===== Универсальный парсер BBQ24 =====
def parse_bbq24():
    response = requests.get(COMPETITOR_URL, headers=HEADERS)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "html.parser")

    products = {}
    cards = soup.find_all(class_=lambda x: x and ("ty-grid-list" in x or "ut2-gl" in x))
    for card in cards:
        title_tag = card.find("a") or card.find("div")
        price_tag = card.find(class_=lambda x: x and "price" in x)
        if title_tag and price_tag:
            title = clean_name(title_tag.get_text(strip=True))
            price = parse_price(price_tag.get_text(strip=True))
            if price:
                products[title] = price
    return products

def compare_prices(our_prices, competitor_prices):
    rows = []
    for our_name, our_price in our_prices.items():
        matched_competitor = None
        competitor_price = None

        for comp_name, comp_price in competitor_prices.items():
            if all(word in comp_name for word in our_name.split()[:2]):
                matched_competitor = comp_name
                competitor_price = comp_price
                break

        if competitor_price is not None:
            diff = our_price - competitor_price
            rows.append([our_name, our_price, competitor_price, diff])
        else:
            rows.append([our_name, our_price, None, None])
    return rows

def save_to_excel(data):
    df = pd.DataFrame(data, columns=["Товар", "Наша цена", "Цена конкурента", "Разница"])
    filename = f"comparison_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    df.to_excel(filename, index=False)

    wb = load_workbook(filename)
    ws = wb.active

    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    gray_fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")

    for row in range(2, ws.max_row + 1):
        diff = ws.cell(row=row, column=4).value
        if diff is not None:
            if diff < 0:
                ws.cell(row=row, column=4).fill = green_fill
            elif diff > 0:
                ws.cell(row=row, column=4).fill = red_fill
            else:
                ws.cell(row=row, column=4).fill = gray_fill

    wb.save(filename)
    print(f"✅ Результат сохранён в {filename}")

if __name__ == "__main__":
    print("📦 Парсим наш сайт...")
    our_prices = parse_grillmaster()
    print(f"  Найдено {len(our_prices)} товаров.")

    print("📦 Парсим сайт конкурента...")
    competitor_prices = parse_bbq24()
    print(f"  Найдено {len(competitor_prices)} товаров.")

    print("📊 Сравниваем цены...")
    comparison_data = compare_prices(our_prices, competitor_prices)

    print("💾 Сохраняем в Excel...")
    save_to_excel(comparison_data)
