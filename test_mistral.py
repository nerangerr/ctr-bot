from mistralai import Mistral
import os
from dotenv import load_dotenv

load_dotenv()
api_key = os.getenv("MISTRAL_API_KEY")
print("🔑 API Key exists:", bool(api_key))

try:
    client = Mistral(api_key=api_key)
    models = client.models.list()
    print(f"📚 Найдено моделей: {len(models.data)}")
    for m in models.data:
        print(m.id)
except Exception as e:
    print("❌ Ошибка:", e)