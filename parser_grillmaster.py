import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
import os

# URL каталога газовых грилей
URL = "https://grillmaster.dp.ua/hazovi-hryli/"

# Папка для сохранения данных
os.makedirs("data", exist_ok=True)

def get_prices():
    response = requests.get(URL)
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "html.parser")

    products = []

    for item in soup.select(".product"):
        title_tag = item.select_one(".woocommerce-loop-product__title")
        price_tag = item.select_one(".woocommerce-Price-amount.amount")

        if title_tag and price_tag:
            title = title_tag.get_text(strip=True)
            price = price_tag.get_text(strip=True)
            products.append({
                "Название": title,
                "Цена": price,
                "Дата": datetime.now().strftime("%d.%m.%Y")
            })

    return products

def save_to_excel(products):
    df = pd.DataFrame(products)
    file_path = os.path.join("data", "prices.xlsx")
    df.to_excel(file_path, index=False)
    print(f"✅ Данные сохранены: {file_path}")

if __name__ == "__main__":
    prices = get_prices()
    save_to_excel(prices)
