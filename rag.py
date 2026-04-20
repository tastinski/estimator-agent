import chromadb
import json
import os

client = chromadb.PersistentClient(path="./chroma_db")


def init_database():
    """
    Инициализация ChromaDB.
    Пересобирает индекс если количество категорий в JSON не совпадает с БД —
    так новые цены подхватываются автоматически после обновления prices_data.json.
    """
    # Читаем JSON первым делом — чтобы понять сколько должно быть записей
    json_path = "prices_data.json"
    if not os.path.exists(json_path):
        print(f"❌ {json_path} not found")
        return None

    with open(json_path, "r") as f:
        data = json.load(f)

    expected_count = len(data)

    # Получаем коллекцию
    collection = client.get_or_create_collection(name="prices")
    existing = collection.count()

    # Если количество совпадает — ничего не делаем
    if existing == expected_count:
        print(f"✅ База актуальна: {existing} записей")
        return collection

    # Количество не совпало — пересоздаём
    print(f"⚠️ Количество изменилось (DB: {existing}, JSON: {expected_count}) — пересобираем индекс")

    try:
        client.delete_collection(name="prices")
    except Exception as e:
        print(f"delete_collection: {e}")

    collection = client.get_or_create_collection(name="prices")

    documents = []
    metadatas = []
    ids = []

    for i, item in enumerate(data):
        documents.append(item["data"])
        metadatas.append({"category": item["category"]})
        ids.append(f"price_{i}")

    collection.add(
        documents=documents,
        metadatas=metadatas,
        ids=ids
    )

    print(f"✅ Загружено {len(documents)} категорий в ChromaDB")
    return collection


def search_prices(query: str, n_results: int = 3) -> str:
    collection = client.get_or_create_collection(name="prices")

    results = collection.query(
        query_texts=[query],
        n_results=n_results
    )

    if results["documents"]:
        return "\n".join(results["documents"][0])
    return ""


if __name__ == "__main__":
    init_database()

    # Тест
    result = search_prices("bathroom tile work")
    print("\n🔍 Тест поиска 'bathroom tile work':")
    print(result)
