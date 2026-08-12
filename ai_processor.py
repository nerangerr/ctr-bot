"""
Модуль для работы с Groq API (безопасное извлечение поручений)
"""

import os
import json
import re
import logging
import traceback
from typing import Optional, List, Dict, Any

from groq import Groq
from dotenv import load_dotenv

load_dotenv()

GROQ_API_KEY = os.getenv("GROQ_API_KEY", "")
GROQ_ENABLED = os.getenv("GROQ_ENABLED", "false").lower() == "true"
GROQ_MODEL = os.getenv("GROQ_MODEL", "llama-3.3-70b-versatile")

print(f"🔍 GROQ_ENABLED = {GROQ_ENABLED}")

logger = logging.getLogger(__name__)

groq_client = Groq(api_key=GROQ_API_KEY) if GROQ_API_KEY else None

EXTRACTION_PROMPT = """
Ты извлекаешь ПРОТОКОЛЬНЫЕ ПОРУЧЕНИЯ из текста документа.

Правила для определения ответственных:
1. "Исп.: ФИО" — это исполнитель для ВСЕГО документа, НЕ создавай отдельное поручение для этого человека.
2. "Ответственный: ФИО" — это ответственный за конкретное поручение.
3. Подпись "И.о. директора ФИО" — это просто подпись, НЕ создавай поручение для этого человека.
4. Одно поручение = одна задача. Не создавай два поручения из одного действия.
5. Если в тексте есть "уведомляем Вас об одностороннем отказе" — это ОДНО поручение.

Верни ТОЛЬКО валидный JSON в формате:
{
    "protocol": {
        "number": "номер протокола или null",
        "date": "YYYY-MM-DD или null",
        "title": "название или null"
    },
    "assignments": [
        {
            "short_task": "Краткая формулировка поручения",
            "description": "Подробное описание",
            "deadline": "YYYY-MM-DD или null",
            "responsible": ["ФИО ответственного"]
        }
    ]
}

Текст документа:
{document_text}
"""

def safe_json_loads(content: str) -> dict:
    """Безопасно парсит JSON, сохраняет ошибки в файл."""
    if not content:
        return {}
    
    # Сохраняем сырой ответ
    with open("groq_raw_response.txt", "w", encoding="utf-8") as f:
        f.write(content)
    print("✅ Сырой ответ Groq сохранён в groq_raw_response.txt")
    
    # Убираем markdown-разметку
    content = re.sub(r'^```(?:json)?\s*', '', content)
    content = re.sub(r'\s*```$', '', content)
    content = content.strip()
    
    # Ищем JSON-объект или массив
    json_match = re.search(r'(\{.*\}|\[.*\])', content, re.DOTALL)
    if json_match:
        content = json_match.group(1)
    
    # Пробуем распарсить
    try:
        return json.loads(content)
    except json.JSONDecodeError:
        # Пробуем исправить: убираем лишние запятые
        content = re.sub(r',\s*}', '}', content)
        content = re.sub(r',\s*]', ']', content)
        try:
            return json.loads(content)
        except json.JSONDecodeError as e:
            print(f"❌ Ошибка парсинга JSON: {e}")
            with open("groq_invalid_json.txt", "w", encoding="utf-8") as f:
                f.write(content)
            return {}

def extract_assignments_from_text(text: str) -> Optional[Dict[str, Any]]:
    print("⏳ extract_assignments_from_text вызвана")
    try:
        if not GROQ_ENABLED:
            print("❌ Groq отключён")
            return None
        if not GROQ_API_KEY:
            print("❌ GROQ_API_KEY не задан")
            return None
        if not text or len(text.strip()) < 10:
            print("❌ Текст слишком короткий")
            return None

        if len(text) > 24000:
            text = text[:24000]

        prompt = EXTRACTION_PROMPT.replace("{document_text}", text)

        print("⏳ Отправляем запрос в Groq...")
        response = groq_client.chat.completions.create(
            model=GROQ_MODEL,
            messages=[
                {"role": "system", "content": "Ты строгий JSON-парсер протокольных поручений."},
                {"role": "user", "content": prompt},
            ],
            temperature=0.1,
            max_tokens=3000,
        )
        print("✅ Groq ответил")

        content = response.choices[0].message.content or ""
        print("✅ Сырой ответ получен, сохраняем...")
        with open("groq_raw_response.txt", "w", encoding="utf-8") as f:
            f.write(content)
        print("✅ Сырой ответ сохранён в groq_raw_response.txt")

        # Парсим
        data = safe_json_loads(content)
        if "assignments" in data and data["assignments"]:
            print(f"✅ Найдено {len(data['assignments'])} поручений")
            return data
        else:
            print("⚠️ Groq не нашёл поручений")
            return None

    except Exception as e:
        print(f"❌ Критическая ошибка в Groq: {e}")
        traceback.print_exc()
        with open("groq_error.txt", "w", encoding="utf-8") as f:
            f.write(str(e))
            f.write("\n")
            f.write(traceback.format_exc())
        return None

def normalize_assignment_from_groq(raw: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "short_task": raw.get("short_task", "Без названия"),
        "description": raw.get("description", ""),
        "deadline": raw.get("deadline"),
        "responsible": raw.get("responsible", ["Не указан"]),
        "source_fragment": raw.get("source_fragment", ""),
        "confidence": raw.get("confidence", 0.5),
    }