import requests
from bs4 import BeautifulSoup
import json
import time

CATEGORIES = [
    "bathroom-remodel",
    "kitchen-remodel", 
    "tile-installation",
    "flooring",
    "painting",
    "plumbing",
    "electrical",
    "drywall",
    "roofing",
    "deck-building",
    "fence-installation",
    "window-installation",
    "door-installation",
    "hvac",
    "handyman",
    "carpentry",
    "concrete",
    "landscaping",
    "pressure-washing",
    "gutter-cleaning"
]

def scrape_category(category):
    url = f"https://www.homeadvisor.com/cost/{category}/"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    }
    
    try:
        response = requests.get(url, headers=headers, timeout=10)
        soup = BeautifulSoup(response.content, "html.parser")
        
        # Берем основной текст со страницы
        text_blocks = []
        
        # Заголовок
        title = soup.find("h1")
        if title:
            text_blocks.append(title.get_text(strip=True))
        
        # Параграфы с ценами
        paragraphs = soup.find_all("p")
        for p in paragraphs:
            text = p.get_text(strip=True)
            if any(word in text.lower() for word in ["cost", "price", "average", "$", "per"]):
                if len(text) > 30:
                    text_blocks.append(text)
        
        return {
            "category": category,
            "url": url,
            "data": " ".join(text_blocks[:10])  # берем топ 10 блоков
        }
    
    except Exception as e:
        print(f"❌ Ошибка {category}: {e}")
        return None

def main():
    print("🔍 Начинаем сбор данных с HomeAdvisor...")
    results = []
    
    for category in CATEGORIES:
        print(f"📊 Скрапим: {category}")
        data = scrape_category(category)
        if data:
            results.append(data)
        time.sleep(2)  # пауза чтобы не забанили
    
    with open("prices_data.json", "w") as f:
        json.dump(results, f, indent=2)
    
    print(f"✅ Готово! Собрано {len(results)} категорий")
    print("📁 Сохранено в prices_data.json")

if __name__ == "__main__":
    main()
