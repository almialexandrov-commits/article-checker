import os
import re
import uuid
import requests
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from docx import Document
from docx.shared import Pt

app = FastAPI()

# Создаем папку для временных файлов
os.makedirs("uploads", exist_ok=True)

# Подключаем папку static для отдачи HTML-страницы
app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/", response_class=HTMLResponse)
async def read_index():
    with open("static/index.html", "r", encoding="utf-8") as f:
        return f.read()

@app.post("/check")
async def check_article(file: UploadFile = File(...)):
    if not file.filename.endswith('.docx'):
        raise HTTPException(status_code=400, detail="Принимаются только файлы в формате .docx")

    file_id = str(uuid.uuid4())
    file_path = f"uploads/{file_id}.docx"
    
    # Сохраняем файл на сервере
    with open(file_path, "wb") as buffer:
        content = await file.read()
        buffer.write(content)

    try:
        doc = Document(file_path)
        report = analyze_document(doc, file.filename)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка чтения файла: {str(e)}")
    finally:
        if os.path.exists(file_path):
            os.remove(file_path)

    return report

def analyze_document(doc: Document, filename: str) -> dict:
    errors = []
    paragraphs = doc.paragraphs
    full_text = "\n".join([p.text.strip() for p in paragraphs])

    # 1. Проверка УДК
    if not re.search(r'^УДК\s+\d+', full_text, re.MULTILINE | re.IGNORECASE):
        errors.append(" <b>УДК:</b> Не найден индекс УДК в начале статьи (пример: УДК 271.2).")

    # 2. Проверка порядка: Аннотация должна быть раньше Ключевых слов
    annot_idx = -1
    kw_idx = -1
    for i, p in enumerate(paragraphs):
        lower_text = p.text.lower()
        if "аннотация" in lower_text and annot_idx == -1:
            annot_idx = i
        if "ключевые слова" in lower_text and kw_idx == -1:
            kw_idx = i

    if annot_idx != -1 and kw_idx != -1:
        if kw_idx < annot_idx:
            errors.append("❌ <b>Структура:</b> Нарушен порядок! 'Ключевые слова' идут перед 'Аннотацией'. Сначала должна быть аннотация.")
    else:
        if annot_idx == -1:
            errors.append("❌ <b>Структура:</b> Не найдено слово 'Аннотация'.")
        if kw_idx == -1:
            errors.append("❌ <b>Структура:</b> Не найдены 'Ключевые слова'.")

    # 3. Проверка шрифта (выборочно, для полных абзацев)
    font_errors = []
    for i, p in enumerate(paragraphs):
        if len(p.text.strip()) > 100:  
            for run in p.runs:
                if run.font.name and run.font.name != 'Times New Roman':
                    font_errors.append(f"шрифт '{run.font.name}'")
                    break
                if run.font.size and run.font.size != Pt(12):
                    font_errors.append(f"размер {run.font.size.pt} пт")
                    break
    
    if font_errors:
        unique_font_errors = list(set(font_errors))[:3]
        errors.append(f"⚠️ <b>Шрифт:</b> В некоторых абзацах обнаружены несоответствия (например: {', '.join(unique_font_errors)}). Требование: Times New Roman, 12 пт.")

    # 4. Проверка орфографии (Яндекс.Спеллер)
    main_text = extract_main_text(paragraphs)
    if main_text:
        spelling_errors = check_spelling_yandex(main_text)
        if spelling_errors:
            errors.append(f"⚠️ <b>Орфография:</b> Найдено потенциальных ошибок: {len(spelling_errors)}. (Пример: '{spelling_errors[0]['word']}' → варианты: {', '.join(spelling_errors[0]['s'][:3])}). <br><i>Примечание: церковнославянизмы и специфические термины могут определяться как ошибки.</i>")

    # 5. Напоминание о сносках
    errors.append("ℹ️ <b>Напоминание:</b> Убедитесь, что сноски оформлены автоматически и располагаются внизу страницы, а имена авторов в библиографии выделены *курсивом*.")

    return {
        "filename": filename,
        "is_valid": len([e for e in errors if e.startswith("❌")]) == 0,
        "errors": errors
    }

def extract_main_text(paragraphs):
    in_body = False
    text_parts = []
    for p in paragraphs:
        lower_text = p.text.lower().strip()
        if "ключевые слова" in lower_text:
            in_body = True
            continue
        if "список литературы" in lower_text or "источники" in lower_text or "references" in lower_text:
            in_body = False
            break
        if in_body and len(p.text.strip()) > 20:
            text_parts.append(p.text)
    return " ".join(text_parts)

def check_spelling_yandex(text):
    try:
        url = "https://speller.yandex.net/services/spellservice.json/checkText"
        response = requests.post(url, data={"text": text[:10000], "lang": "ru"}, timeout=5)
        return response.json()
    except Exception:
        return []