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
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = FastAPI()

os.makedirs("uploads", exist_ok=True)
app.mount("/static", StaticFiles(directory="static"), name="static")

# DeepSeek API key (получите на https://platform.deepseek.com)
DEEPSEEK_API_KEY = "sk-2dfe45323edf45fcb88961b41cf91a7b"  # Замените на ваш ключ
DEEPSEEK_API_URL = "https://api.deepseek.com/v1/chat/completions"

@app.get("/", response_class=HTMLResponse)
async def read_index():
    with open("static/index.html", "r", encoding="utf-8") as f:
        return f.read()

@app.post("/check")
async def check_article(file: UploadFile = File(...), use_ai: bool = False):
    if not file.filename.endswith('.docx'):
        raise HTTPException(status_code=400, detail="Принимаются только файлы в формате .docx")

    file_id = str(uuid.uuid4())
    file_path = f"uploads/{file_id}.docx"
    
    with open(file_path, "wb") as buffer:
        content = await file.read()
        buffer.write(content)

    try:
        doc = Document(file_path)
        report = analyze_document(doc, file.filename, use_ai)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка чтения файла: {str(e)}")
    finally:
        if os.path.exists(file_path):
            os.remove(file_path)

    return report

def is_church_slavonic_or_greek(text):
    """Проверяет церковнославянские и греческие символы"""
    greek_pattern = r'[\u0370-\u03FF\u1F00-\u1FFF]'
    church_slavonic_pattern = r'[ѣѢѳѲѵѴѧѧѩѫѭѯѱ]'
    
    if re.search(greek_pattern, text) or re.search(church_slavonic_pattern, text):
        return True
    
    church_endings = r'\b\w+(овъ|евъ|інъ|їнъ|іе|їе)\b'
    if re.search(church_endings, text, re.IGNORECASE):
        return True
    
    return False

def check_footnotes_advanced(doc):
    """Улучшенная проверка сносок через прямой доступ к XML"""
    errors = []
    footnote_count = 0
    
    try:
        # Метод 1: Проверяем через related_parts
        for rel in doc.part.rels.values():
            if "footnotes" in rel.reltype:
                footnotes_part = rel.target_part
                footnotes = footnotes_part.element.findall(
                    './/w:footnote', 
                    namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
                )
                footnote_count = len(footnotes)
                
                for footnote in footnotes:
                    footnote_type = footnote.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type')
                    if footnote_type == 'endnote':
                        errors.append({
                            "type": "error",
                            "category": "Сноски",
                            "message": "Обнаружены концевые сноски. Требуются постраничные (внизу страницы).",
                            "paragraph_idx": -1
                        })
                        break
                break
        
        # Метод 2: Ищем сноски в тексте (маркеры сносок)
        if footnote_count == 0:
            for para in doc.paragraphs:
                for run in para.runs:
                    # Ищем элементы сносок в XML
                    if 'w:footnoteReference' in run._element.xml:
                        footnote_count += 1
        
        if footnote_count == 0:
            errors.append({
                "type": "info",
                "category": "Сноски",
                "message": "Сноски не обнаружены. Убедитесь, что они расставлены через Вставка → Сноска.",
                "paragraph_idx": -1
            })
        else:
            errors.append({
                "type": "success",
                "category": "Сноски",
                "message": f"Найдено {footnote_count} сносок. Убедитесь, что все оформлены автоматически.",
                "paragraph_idx": -1
            })
            
    except Exception as e:
        errors.append({
            "type": "info",
            "category": "Сноски",
            "message": f"Не удалось проверить сноски: {str(e)}",
            "paragraph_idx": -1
        })
    
    return errors, footnote_count

def analyze_document(doc: Document, filename: str, use_ai: bool = False) -> dict:
    errors = []
    checklist = []
    paragraphs = doc.paragraphs
    full_text = "\n".join([p.text.strip() for p in paragraphs])
    
    error_id = 0

    # 1. Проверка УДК
    udk_match = re.search(r'^УДК\s+(\d[\d\.\-]+)', full_text, re.MULTILINE | re.IGNORECASE)
    if udk_match:
        checklist.append({"item": "УДК указан", "status": "success", "details": f"УДК {udk_match.group(1)}"})
    else:
        errors.append({
            "id": error_id, "type": "error", "category": "УДК",
            "message": "Не найден индекс УДК в начале статьи.",
            "paragraph_idx": 0
        })
        checklist.append({"item": "УДК указан", "status": "error", "details": "Не найден"})
        error_id += 1

    # 2. Проверка DOI
    doi_match = re.search(r'https?://doi\.org/([\d\.\-]+)', full_text, re.IGNORECASE)
    if doi_match:
        checklist.append({"item": "DOI указан", "status": "success", "details": doi_match.group(0)})
    else:
        errors.append({
            "id": error_id, "type": "warning", "category": "DOI",
            "message": "Не найден DOI.",
            "paragraph_idx": 0
        })
        checklist.append({"item": "DOI указан", "status": "warning", "details": "Не найден"})
        error_id += 1

    # 3. Проверка EDN
    edn_match = re.search(r'EDN\s+([A-Z]{6})', full_text, re.IGNORECASE)
    if edn_match:
        checklist.append({"item": "EDN указан", "status": "success", "details": edn_match.group(1)})
    else:
        checklist.append({"item": "EDN указан", "status": "warning", "details": "Не найден"})
        error_id += 1

    # 4. Проверка аннотации
    annot_pattern = re.search(r'аннотация[.\s\:]*(.+?)(?=ключевые\s+слова|$)', full_text, re.IGNORECASE | re.DOTALL)
    if annot_pattern:
        annot_text = annot_pattern.group(1).strip()
        annot_length = len(annot_text)
        if 2000 <= annot_length <= 2200:
            checklist.append({"item": "Аннотация (2000-2200 знаков)", "status": "success", "details": f"{annot_length} знаков"})
        elif annot_length < 2000:
            checklist.append({"item": "Аннотация (2000-2200 знаков)", "status": "warning", "details": f"Всего {annot_length} знаков (нужно 2000-2200)"})
            errors.append({
                "id": error_id, "type": "warning", "category": "Аннотация",
                "message": f"Аннотация слишком короткая ({annot_length} знаков). Требуется 2000-2200 знаков.",
                "paragraph_idx": -1
            })
            error_id += 1
        else:
            checklist.append({"item": "Аннотация (2000-2200 знаков)", "status": "warning", "details": f"{annot_length} знаков (превышение)"})
    else:
        checklist.append({"item": "Аннотация (2000-2200 знаков)", "status": "error", "details": "Не найдена"})
        error_id += 1

    # 5. Проверка ключевых слов
    kw_pattern = re.search(r'ключевые\s+слова[.\s\:]*(.+?)(?=для\s+цитирования|$)', full_text, re.IGNORECASE | re.DOTALL)
    if kw_pattern:
        kw_text = kw_pattern.group(1).strip()
        kw_list = re.split(r'[,;]', kw_text)
        kw_count = len([kw for kw in kw_list if kw.strip()])
        if 5 <= kw_count <= 10:
            checklist.append({"item": "Ключевые слова (5-10)", "status": "success", "details": f"{kw_count} слов"})
        else:
            checklist.append({"item": "Ключевые слова (5-10)", "status": "warning", "details": f"{kw_count} слов (нужно 5-10)"})
    else:
        checklist.append({"item": "Ключевые слова (5-10)", "status": "error", "details": "Не найдены"})
        error_id += 1

    # 6. Проверка шрифта
    font_issues = []
    for i, p in enumerate(paragraphs):
        if len(p.text.strip()) > 50:
            for run in p.runs:
                text = run.text.strip()
                if not text or is_church_slavonic_or_greek(text):
                    continue
                
                if run.font.name and run.font.name not in ['Times New Roman', 'Times', None]:
                    font_issues.append(f"шрифт '{run.font.name}'")
                    break
                if run.font.size and run.font.size != Pt(12):
                    font_issues.append(f"размер {run.font.size.pt} пт")
                    break
    
    if not font_issues:
        checklist.append({"item": "Шрифт Times New Roman 12 пт", "status": "success", "details": "Соответствует"})
    else:
        unique_issues = list(set(font_issues))[:3]
        checklist.append({"item": "Шрифт Times New Roman 12 пт", "status": "warning", "details": f"Найдено: {', '.join(unique_issues)}"})

    # 7. Проверка сносок
    footnote_errors, footnote_count = check_footnotes_advanced(doc)
    errors.extend(footnote_errors)
    if footnote_count > 0:
        checklist.append({"item": "Сноски оформлены", "status": "success", "details": f"{footnote_count} сносок"})
    else:
        checklist.append({"item": "Сноски оформлены", "status": "warning", "details": "Не обнаружены"})

    # 8. Проверка списка литературы
    biblio_patterns = [
        r'(источники|литература|список\s+литературы|библиографический\s+список)',
        r'references'
    ]
    biblio_found = False
    for pattern in biblio_patterns:
        if re.search(pattern, full_text, re.IGNORECASE):
            biblio_found = True
            break
    
    if biblio_found:
        checklist.append({"item": "Список литературы", "status": "success", "details": "Найден"})
    else:
        checklist.append({"item": "Список литературы", "status": "error", "details": "Не найден"})
        error_id += 1

    # 9. Проверка орфографии через Яндекс
    main_text = extract_main_text(paragraphs)
    spelling_errors = []
    if main_text:
        spelling_errors = check_spelling_yandex_detailed(main_text)
        if spelling_errors:
            for err in spelling_errors[:15]:
                errors.append({
                    "id": error_id, "type": "warning", "category": "Орфография",
                    "message": f"Возможная ошибка: '{err['word']}'",
                    "suggestions": err.get('s', []),
                    "paragraph_idx": -1
                })
                error_id += 1

    # 10. Глубокая проверка через DeepSeek AI (если включено)
    ai_analysis = None
    if use_ai and DEEPSEEK_API_KEY != "sk-xxxxxxxxxxxxxxxxxxxxxxxx":
        try:
            ai_analysis = check_with_deepseek(full_text, main_text)
            if ai_analysis:
                for issue in ai_analysis.get('issues', []):
                    errors.append({
                        "id": error_id, "type": issue.get('type', 'warning'), 
                        "category": "AI-анализ",
                        "message": issue.get('message', ''),
                        "paragraph_idx": -1
                    })
                    error_id += 1
                checklist.append({
                    "item": "AI-проверка качества", 
                    "status": "success" if not ai_analysis.get('critical_issues') else "warning",
                    "details": f"Найдено замечаний: {len(ai_analysis.get('issues', []))}"
                })
        except Exception as e:
            errors.append({
                "id": error_id, "type": "info", "category": "AI",
                "message": f"Не удалось выполнить AI-проверку: {str(e)}",
                "paragraph_idx": -1
            })
            error_id += 1

    article_html = generate_article_html(paragraphs, errors)

    return {
        "filename": filename,
        "is_valid": len([e for e in errors if e["type"] == "error"]) == 0,
        "errors": errors,
        "article_html": article_html,
        "checklist": checklist,
        "total_errors": len([e for e in errors if e["type"] == "error"]),
        "total_warnings": len([e for e in errors if e["type"] == "warning"]),
        "ai_analysis": ai_analysis
    }

def check_with_deepseek(full_text: str, main_text: str) -> dict:
    """Проверка статьи через DeepSeek AI"""
    
    prompt = f"""
Ты - редактор научного богословского журнала. Проверь статью по следующим критериям:

1. Орфография и пунктуация (найди реальные ошибки, игнорируя церковнославянизмы)
2. Структура статьи (логика изложения, наличие введения и выводов)
3. Оформление библиографических ссылок
4. Научный стиль изложения
5. Грамматические ошибки

Текст статьи:
{main_text[:8000]}  # Ограничиваем до 8000 символов

Верни ответ в формате JSON:
{{
    "issues": [
        {{"type": "error|warning", "message": "Описание проблемы"}}
    ],
    "critical_issues": true/false,
    "recommendations": ["рекомендация 1", "рекомендация 2"]
}}

Найди конкретные ошибки, например:
- Повторы слов
- Грамматические ошибки
- Неправильное оформление ссылок
- Отсутствие выводов

Игнорируй:
- Церковнославянские слова
- Цитаты из Библии
- Специфические богословские термины
"""

    try:
        headers = {
            "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
            "Content-Type": "application/json"
        }
        
        payload = {
            "model": "deepseek-chat",
            "messages": [
                {"role": "system", "content": "Ты помощник редактора научного журнала."},
                {"role": "user", "content": prompt}
            ],
            "temperature": 0.3,
            "response_format": {"type": "json_object"}
        }
        
        response = requests.post(DEEPSEEK_API_URL, json=payload, timeout=30)
        response.raise_for_status()
        
        result = response.json()
        content = result['choices'][0]['message']['content']
        
        return json.loads(content)
        
    except Exception as e:
        return {"issues": [], "critical_issues": False, "error": str(e)}

def generate_article_html(paragraphs, errors):
    """Генерирует HTML с сохранением базового форматирования"""
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
        
        # Сохраняем форматирование через runs
        paragraph_html = ""
        for run in p.runs:
            text = run.text
            if not text:
                continue
            
            # Применяем форматирование
            style = ""
            if run.font.bold:
                style += "font-weight: bold; "
            if run.font.italic:
                style += "font-style: italic; "
            if run.font.underline:
                style += "text-decoration: underline; "
            
            if style:
                paragraph_html += f'<span style="{style}">{text}</span>'
            else:
                paragraph_html += text
        
        paragraph_errors = errors_by_paragraph.get(i, [])
        
        if paragraph_errors:
            error_classes = ' '.join([f"error-{err['id']}" for err in paragraph_errors])
            has_critical = any(err['type'] == 'error' for err in paragraph_errors)
            class_type = "critical" if has_critical else ""
            
            tooltip_content = '<br>'.join([f"{err['category']}: {err['message']}" for err in paragraph_errors])
            
            html_parts.append(f'''
                <div class="paragraph error {class_type} {error_classes}" 
                     data-error-id="{paragraph_errors[0]['id']}"
                     title="{tooltip_content}">
                    {paragraph_html}
                </div>
            ''')
        else:
            html_parts.append(f'<div class="paragraph">{paragraph_html}</div>')
    
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
