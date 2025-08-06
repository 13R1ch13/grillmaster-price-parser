import time
import re
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# ====== URLs ======
OUR_URL = "https://grillmaster.dp.ua/hazovi-hryli/"
COMPETITOR_URL = "https://bbq24.com.ua/ua/gazovye-grili/"

# ====== Настройки Selenium (автоустановка драйвера) ======
def get_driver():
    options = Options()
    options.add_argument("--headless")  # без графического интерфейса
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--window-size=1920,1080")
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# ====== Вспомогательные ======
def clean_name(name):
    name = name.lower()
    name = re.sub(r"[^a-zа-я0-9\s]", "", name)
    return name.strip()

def parse_price(price_text):
    digits = re.sub(r"[^\d]", "", price_text)
    return int(digits) if digits else None

# ====== Прокрутка страницы ======
def scroll_page(driver):
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height

# ====== Парсер Grill Master ======
def parse_grillmaster(driver):
    driver.get(OUR_URL)
    time.sleep(2)
    scroll_page(driver)

    products = {}
    cards = driver.find_elements(By.CSS_SELECTOR, ".product")
    for card in cards:
        try:
            title = clean_name(card.find_element(By.CSS_SELECTOR, "h2").text)
            price = parse_price(card.find_element(By.CSS_SELECTOR, ".amount").text)
            if price:
                products[title] = price
        except:
            continue
    return products

# ====== Парсер BBQ24 ======
def parse_bbq24(driver):
    driver.get(COMPETITOR_URL)
    time.sleep(2)
    scroll_page(driver)

    products = {}
    cards = driver.find_elements(By.CSS_SELECTOR, ".ut2-gl__item, .ty-grid-list__item")
    for card in cards:
        try:
            title = clean_name(card.text.split("\n")[0])
            price_elements = card.find_elements(By.CSS_SELECTOR, ".ty-price, .price")
            if price_elements:
                price = parse_price(price_elements[0].text)
                if price:
                    products[title] = price
        except:
            continue
    return products

# ====== Сравнение ======
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

# ====== Сохранение в Excel ======
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

# ====== Запуск ======
if __name__ == "__main__":
    driver = get_driver()

    print("📦 Парсим наш сайт...")
    our_prices = parse_grillmaster(driver)
    print(f"  Найдено {len(our_prices)} товаров.")

    print("📦 Парсим сайт конкурента...")
    competitor_prices = parse_bbq24(driver)
    print(f"  Найдено {len(competitor_prices)} товаров.")

    driver.quit()

    print("📊 Сравниваем цены...")
    comparison_data = compare_prices(our_prices, competitor_prices)

    print("💾 Сохраняем в Excel...")
    save_to_excel(comparison_data)
