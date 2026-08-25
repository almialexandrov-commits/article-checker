import os
import re
import uuid
import requests
import json
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from docx import Document
from docx.shared import Pt

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
    greek_pattern = r'[\u0370-\u03FF\u1F00-\u1FFF]'
    church_slavonic_pattern = r'[ѣѢѳѲѵѴѧѧѩѫѭѯѱ]'
    
    if re.search(greek_pattern, text) or re.search(church_slavonic_pattern, text):
        return True
    
    church_endings = r'\b\w+(овъ|евъ|інъ|їнъ|іе|їе)\b'
    if re.search(church_endings, text, re.IGNORECASE):
        return True
    
    return False

def find_section_in_text(paragraphs, section_names):
    """
    Улучшенный поиск раздела в тексте
    section_names - список возможных названий (например: ['список литературы', 'источники'])
    """
    for i, p in enumerate(paragraphs):
        text_lower = p.text.lower().strip()
        # Убираем лишние пробелы и знаки препинания
        text_clean = re.sub(r'[\s\-\–\—]+', ' ', text_lower)
        text_clean = re.sub(r'[^\w\sа-яё]', '', text_clean)
        
        for section in section_names:
            section_clean = re.sub(r'[\s]+', ' ', section.lower())
            if section_clean in text_clean or section.lower() in text_lower:
                return i
    
    # Дополнительная проверка по всему тексту
    full_text = "\n".join([p.text.lower() for p in paragraphs])
    for section in section_names:
        if section.lower() in full_text:
            return -2  # Найдено, но не в начале абзаца
    
    return -1  # Не найдено

def analyze_document(doc: Document, filename: str) -> dict:
    errors = []
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

    # 2. Проверка DOI
    if not re.search(r'https?://doi\.org/[\d\.\-]+', full_text, re.IGNORECASE):
        errors.append({
            "id": error_id,
            "type": "warning",
            "category": "DOI",
            "message": "Не найден DOI (пример: https://doi.org/10.24412/2224-5391-0000-00-00-00).",
            "paragraph_idx": 0
        })
        error_id += 1

    # 3. Проверка EDN
    if not re.search(r'EDN\s+[A-Z]{6}', full_text, re.IGNORECASE):
        errors.append({
            "id": error_id,
            "type": "warning",
            "category": "EDN",
            "message": "Не найден код EDN (уникальный номер статьи в eLIBRARY).",
            "paragraph_idx": 0
        })
        error_id += 1

    # 4. Проверка порядка элементов
    annot_idx = find_section_in_text(paragraphs, ['аннотация', 'abstract'])
    kw_idx = find_section_in_text(paragraphs, ['ключевые слова', 'keywords'])
    biblio_idx = find_section_in_text(paragraphs, [
        'список литературы', 
        'источники', 
        'литература',
        'библиографический список',
        'references',
        'библиография'
    ])
    
    # Проверяем порядок: Аннотация -> Ключевые слова
    if annot_idx != -1 and kw_idx != -1 and annot_idx != -2 and kw_idx != -2:
        if kw_idx < annot_idx:
            errors.append({
                "id": error_id,
                "type": "error",
                "category": "Структура",
                "message": "Нарушен порядок! 'Ключевые слова' идут перед 'Аннотацией'.",
                "paragraph_idx": kw_idx
            })
            error_id += 1
    
    if annot_idx == -1:
        errors.append({
            "id": error_id,
            "type": "error",
            "category": "Структура",
            "message": "Не найдена 'Аннотация'.",
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

    # 5. Проверка наличия информации для цитирования
    if not re.search(r'для\s+цитирования|for\s+citation', full_text, re.IGNORECASE):
        errors.append({
            "id": error_id,
            "type": "warning",
            "category": "Цитирование",
            "message": "Не найден блок 'Для цитирования'.",
            "paragraph_idx": 0
        })
        error_id += 1

    # 6. Проверка шрифта
    font_errors = []
    for i, p in enumerate(paragraphs):
        if len(p.text.strip()) > 50:
            for run in p.runs:
                text = run.text.strip()
                if not text or is_church_slavonic_or_greek(text):
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
                    if len(font_errors) >= 5:
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

    # 7. Проверка орфографии
    main_text = extract_main_text(paragraphs)
    if main_text:
        spelling_errors = check_spelling_yandex_detailed(main_text)
        for err in spelling_errors[:10]:
            errors.append({
                "id": error_id,
                "type": "warning",
                "category": "Орфография",
                "message": f"Возможная ошибка: '{err['word']}'",
                "suggestions": err.get('s', []),
                "paragraph_idx": -1
            })
            error_id += 1

    # 8. Проверка списка литературы
    if biblio_idx == -1:
        errors.append({
            "id": error_id,
            "type": "error",
            "category": "Структура",
            "message": "Не найден список литературы/источников. Проверьте наличие разделов: 'Источники', 'Литература', 'Список литературы' или 'References'.",
            "paragraph_idx": len(paragraphs) - 1
        })
        error_id += 1
    else:
        # Проверяем, есть ли позиции в списке
        biblio_section_found = False
        for i in range(biblio_idx if biblio_idx >= 0 else 0, len(paragraphs)):
            p_text = paragraphs[i].text.strip()
            # Проверяем наличие библиографических записей (автор, год, страницы)
            if re.search(r'\d{4}[\.\s]|с\.\s*\d+|pp\.\s*\d+', p_text, re.IGNORECASE):
                biblio_section_found = True
                break
        
        if not biblio_section_found and biblio_idx >= 0:
            errors.append({
                "id": error_id,
                "type": "warning",
                "category": "Библиография",
                "message": "Список литературы найден, но в нем не обнаружены библиографические записи.",
                "paragraph_idx": biblio_idx
            })
            error_id += 1

    # 9. Проверка сносок
    footnote_errors = check_footnotes(doc)
    for err in footnote_errors:
        err["id"] = error_id
        errors.append(err)
        error_id += 1

    # 10. Проверка объема (примерная)
    char_count = len(full_text.replace(" ", ""))
    if char_count > 80000:
        errors.append({
            "id": error_id,
            "type": "warning",
            "category": "Объем",
            "message": f"Объем статьи ({char_count} знаков) превышает рекомендуемый (80 000 знаков).",
            "paragraph_idx": 0
        })
        error_id += 1

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
    
    try:
        footnotes_part = doc.part.related_parts.get('footnotes')
        if footnotes_part:
            footnotes = footnotes_part.element.findall(
                './/w:footnote', 
                namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            )
            
            if len(footnotes) == 0:
                errors.append({
                    "type": "info",
                    "category": "Сноски",
                    "message": "В документе не найдены сноски. Убедитесь, что сноски расставлены.",
                    "paragraph_idx": -1
                })
            else:
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
                "type": "info",
                "category": "Сноски",
                "message": "В документе не найдены сноски. Убедитесь, что сноски оформлены автоматически через Word (Вставка → Сноска).",
                "paragraph_idx": -1
            })
    except Exception as e:
        errors.append({
            "type": "info",
            "category": "Сноски",
            "message": "Не удалось проверить сноски автоматически.",
            "paragraph_idx": -1
        })
    
    return errors

def generate_article_html(paragraphs, errors):
    """Генерирует HTML статьи с подсветкой ошибок"""
    html_parts = []
    
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
            error_classes = ' '.join([f"error-{err['id']}" for err in paragraph_errors])
            has_critical = any(err['type'] == 'error' for err in paragraph_errors)
            class_type = "critical" if has_critical else ""
            
            tooltip_content = '<br>'.join([f"{err['category']}: {err['message']}" for err in paragraph_errors])
            
            html_parts.append(f'''
                <div class="paragraph error {class_type} {error_classes}" 
                     data-error-id="{paragraph_errors[0]['id']}"
                     onclick="showErrorDetails({paragraph_errors[0]['id']})">
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
        if "ключевые слова" in lower_text or "keywords" in lower_text:
            in_body = True
            continue
        if re.search(r'(список\s+литературы|источники|references|литература)', lower_text):
            in_body = False
            break
        if in_body and len(p.text.strip()) > 20:
            text_parts.append(p.text)
    return " ".join(text_parts)

def check_spelling_yandex_detailed(text):
    """Детальная проверка орфографии"""
    try:
        url = "https://speller.yandex.net/services/spellservice.json/checkText"
        response = requests.post(url, data={"text": text[:10000], "lang": "ru"}, timeout=5)
        return response.json()
    except Exception:
        return []
