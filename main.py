import os
import re
import uuid
import requests
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

app = FastAPI()

os.makedirs("uploads", exist_ok=True)
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

def is_church_slavonic_or_greek(text):
    """Проверяет, содержит ли текст церковнославянские или греческие символы"""
    # Греческие буквы
    greek_pattern = r'[\u0370-\u03FF\u1F00-\u1FFF]'
    # Церковнославянские специфические символы (ятть, фита, ижица и т.д.)
    church_slavonic_pattern = r'[ѣѢѳѲѵѴѧѧѩѫѭѯѱ]'
    
    if re.search(greek_pattern, text) or re.search(church_slavonic_pattern, text):
        return True
    
    # Проверяем, есть ли слова с характерными церковнославянскими окончаниями
    church_endings = r'\b\w+(овъ|евъ|інъ|їнъ|іе|їе)\b'
    if re.search(church_endings, text, re.IGNORECASE):
        return True
    
    return False

def analyze_document(doc: Document, filename: str) -> dict:
    errors = []
    highlighted_text = []
    paragraphs = doc.paragraphs
    full_text = "\n".join([p.text.strip() for p in paragraphs])
    
    error_id = 0

    # 1. Проверка УДК
    if not re.search(r'^УДК\s+\d+', full_text, re.MULTILINE | re.IGNORECASE):
        errors.append({
            "id": error_id,
            "type": "error",
            "category": "УДК",
            "message": "Не найден индекс УДК в начале статьи (пример: УДК 271.2).",
            "paragraph_idx": 0
        })
        error_id += 1

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
            errors.append({
                "id": error_id,
                "type": "error",
                "category": "Структура",
                "message": "Нарушен порядок! 'Ключевые слова' идут перед 'Аннотацией'. Сначала должна быть аннотация.",
                "paragraph_idx": kw_idx
            })
            error_id += 1
    else:
        if annot_idx == -1:
            errors.append({
                "id": error_id,
                "type": "error",
                "category": "Структура",
                "message": "Не найдено слово 'Аннотация'.",
                "paragraph_idx": 0
            })
            error_id += 1
        if kw_idx == -1:
            errors.append({
                "id": error_id,
                "type": "error",
                "category": "Структура",
                "message": "Не найдены 'Ключевые слова'.",
                "paragraph_idx": 0
            })
            error_id += 1

    # 3. Проверка шрифта (только для русского текста, исключаем церковнославянские и греческие слова)
    font_errors = []
    for i, p in enumerate(paragraphs):
        if len(p.text.strip()) > 50:  # Проверяем только содержательные абзацы
            for run in p.runs:
                text = run.text.strip()
                if not text:
                    continue
                    
                # Пропускаем церковнославянские и греческие слова
                if is_church_slavonic_or_greek(text):
                    continue
                
                if run.font.name and run.font.name not in ['Times New Roman', 'Times', None]:
                    font_errors.append({
                        "id": error_id,
                        "type": "warning",
                        "category": "Шрифт",
                        "message": f"Шрифт '{run.font.name}' вместо Times New Roman",
                        "text": text[:50],
                        "paragraph_idx": i
                    })
                    error_id += 1
                    if len(font_errors) >= 5:  # Ограничиваем количество ошибок
                        break
                elif run.font.size and run.font.size != Pt(12):
                    font_errors.append({
                        "id": error_id,
                        "type": "warning",
                        "category": "Шрифт",
                        "message": f"Размер {run.font.size.pt} пт вместо 12 пт",
                        "text": text[:50],
                        "paragraph_idx": i
                    })
                    error_id += 1
                    if len(font_errors) >= 5:
                        break
        if len(font_errors) >= 5:
            break
    
    errors.extend(font_errors)

    # 4. Детальная проверка орфографии с выделением слов
    main_text = extract_main_text(paragraphs)
    if main_text:
        spelling_errors = check_spelling_yandex_detailed(main_text)
        for err in spelling_errors[:10]:  # Показываем первые 10 ошибок
            errors.append({
                "id": error_id,
                "type": "warning",
                "category": "Орфография",
                "message": f"Возможная ошибка: '{err['word']}'",
                "suggestions": err.get('s', []),
                "paragraph_idx": -1  # Для орфографии не указываем конкретный абзац
            })
            error_id += 1

    # 5. Проверка сносок
    footnote_errors = check_footnotes(doc)
    errors.extend(footnote_errors)
    for err in footnote_errors:
        error_id += 1

    # 6. Проверка наличия и оформления списка литературы
    if not re.search(r'(список\s+литературы|источники|references)', full_text, re.IGNORECASE):
        errors.append({
            "id": error_id,
            "type": "error",
            "category": "Структура",
            "message": "Не найден список литературы/источников.",
            "paragraph_idx": len(paragraphs) - 1
        })
        error_id += 1

    # Формируем текст статьи с разметкой для отображения
    article_html = generate_article_html(paragraphs, errors)

    return {
        "filename": filename,
        "is_valid": len([e for e in errors if e["type"] == "error"]) == 0,
        "errors": errors,
        "article_html": article_html,
        "total_errors": len([e for e in errors if e["type"] == "error"]),
        "total_warnings": len([e for e in errors if e["type"] == "warning"])
    }

def check_footnotes(doc):
    """Проверяет оформление сносок"""
    errors = []
    
    # Проверяем, есть ли сноски в документе
    try:
        # Получаем все сноски из документа
        footnotes_part = doc.part.related_parts.get('footnotes')
        if footnotes_part:
            footnotes = footnotes_part.element.findall('.//w:footnote', 
                namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
            
            if len(footnotes) == 0:
                errors.append({
                    "type": "warning",
                    "category": "Сноски",
                    "message": "В документе не найдены сноски. Убедитесь, что сноски расставлены.",
                    "paragraph_idx": -1
                })
            else:
                # Проверяем тип сносок (должны быть внизу страницы, не концевые)
                for footnote in footnotes:
                    footnote_type = footnote.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type')
                    if footnote_type == 'endnote':
                        errors.append({
                            "type": "error",
                            "category": "Сноски",
                            "message": "Обнаружены концевые сноски. По правилам журнала сноски должны быть внизу страницы (постраничные).",
                            "paragraph_idx": -1
                        })
                        break
        else:
            errors.append({
                "type": "warning",
                "category": "Сноски",
                "message": "В документе не найдены сноски. Убедитесь, что сноски расставлены автоматически через Word.",
                "paragraph_idx": -1
            })
    except Exception as e:
        errors.append({
            "type": "warning",
            "category": "Сноски",
            "message": f"Не удалось проверить сноски автоматически. Убедитесь, что сноски оформлены правильно.",
            "paragraph_idx": -1
        })
    
    return errors

def generate_article_html(paragraphs, errors):
    """Генерирует HTML статьи с подсветкой ошибок"""
    html_parts = []
    
    # Создаем словарь ошибок по индексам абзацев
    errors_by_paragraph = {}
    for err in errors:
        p_idx = err.get('paragraph_idx', -1)
        if p_idx not in errors_by_paragraph:
            errors_by_paragraph[p_idx] = []
        errors_by_paragraph[p_idx].append(err)
    
    for i, p in enumerate(paragraphs):
        if not p.text.strip():
            continue
            
        paragraph_errors = errors_by_paragraph.get(i, [])
        
        if paragraph_errors:
            # Если есть ошибки в абзаце, добавляем подсветку
            error_classes = ' '.join([f"error-{err['id']}" for err in paragraph_errors])
            tooltip_content = '<br>'.join([f"{err['category']}: {err['message']}" for err in paragraph_errors])
            
            html_parts.append(f'''
                <div class="paragraph error {error_classes}" 
                     data-tooltip="{tooltip_content}"
                     title="{tooltip_content}">
                    {p.text}
                </div>
            ''')
        else:
            html_parts.append(f'<div class="paragraph">{p.text}</div>')
    
    return ''.join(html_parts)

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

def check_spelling_yandex_detailed(text):
    """Детальная проверка орфографии с возвратом конкретных слов"""
    try:
        url = "https://speller.yandex.net/services/spellservice.json/checkText"
        response = requests.post(url, data={"text": text[:10000], "lang": "ru"}, timeout=5)
        return response.json()
    except Exception:
        return []
