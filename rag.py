import chromadb
import json
import os

client = chromadb.PersistentClient(path="./chroma_db")

def init_database():
    collection = client.get_or_create_collection(name="prices")
    
    existing = collection.count()
    if existing > 0:
        print(f"✅ База уже загружена: {existing} записей")
        return collection
    
    with open("prices_data.json", "r") as f:
        data = json.load(f)
    
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
