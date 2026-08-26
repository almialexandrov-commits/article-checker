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
from datetime import datetime

app = FastAPI()

os.makedirs("uploads", exist_ok=True)
app.mount("/static", StaticFiles(directory="static"), name="static")

# DeepSeek API key
DEEPSEEK_API_KEY = "sk-2dfe45323edf45fcb88961b41cf91a7b"
DEEPSEEK_API_URL = "https://api.deepseek.com/v1/chat/completions"

# Журнал для проверки
JOURNAL_NAME = "Вестник Екатеринбургской духовной семинарии"
JOURNAL_URL = "https://journals.epds.ru/index.php/VEDS/about/submissions"

# Библейские сокращения (игнорировать при проверке орфографии)
BIBLE_ABBREVIATIONS = {
    'Быт', 'Исх', 'Лев', 'Чис', 'Втор', 'Нав', 'Суд', 'Руф', '1Цар', '2Цар', '3Цар', '4Цар',
    '1Пар', '2Пар', '1Ездр', 'Ездр', 'Неем', 'Есф', 'Иов', 'Пс', 'Притч', 'Еккл', 'Песн',
    'Прем', 'Сир', 'Ис', 'Иер', 'Плач', 'Посл', 'Иез', 'Дан', 'Ос', 'Иоил', 'Ам', 'Авд',
    'Иона', 'Мих', 'Наум', 'Авв', 'Соф', 'Агг', 'Зах', 'Мал', '1Макк', '2Макк', '3Макк',
    'Мф', 'Мк', 'Лк', 'Ин', 'Деян', 'Рим', '1Кор', '2Кор', 'Гал', 'Еф', 'Флп', 'Кол',
    '1Фес', '2Фес', '1Тим', '2Тим', 'Тит', 'Флм', 'Евр', 'Иак', '1Пет', '2Пет', '1Ин',
    '2Ин', '3Ин', 'Иуд', 'Откр', 'Апок'
}

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

def check_footnotes(doc):
    """Проверяет сноски через прямой поиск в XML"""
    errors = []
    footnote_count = 0
    footnote_positions = []
    
    try:
        # Ищем ссылки на сноски в тексте
        footnote_refs = doc.element.body.findall('.//' + qn('w:footnoteReference'))
        footnote_count = len(footnote_refs)
        
        # Получаем информацию о сносках
        for i, ref in enumerate(footnote_refs):
            footnote_id = ref.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
            footnote_positions.append({
                'id': footnote_id,
                'index': i
            })
        
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
                "message": f"Найдено {footnote_count} сносок. Проверьте оформление по ГОСТ.",
                "paragraph_idx": -1,
                "footnote_count": footnote_count
            })
            
    except Exception as e:
        errors.append({
            "type": "info",
            "category": "Сноски",
            "message": f"Не удалось проверить сноски: {str(e)}",
            "paragraph_idx": -1
        })
    
    return errors, footnote_count, footnote_positions

def count_keywords(kw_text):
    """Правильно считает ключевые слова"""
    if not kw_text:
        return 0
    
    # Разбиваем по запятым, точкам с запятой, переносам строк
    keywords = re.split(r'[,;\n]', kw_text)
    # Убираем пустые и слишком короткие
    keywords = [kw.strip() for kw in keywords if kw.strip() and len(kw.strip()) > 2]
    
    return len(keywords)

def check_journal_citation(full_text):
    """Проверяет, правильный ли журнал указан в цитировании"""
    # Ищем блок цитирования
    citation_pattern = re.search(
        r'цитирование[.\s\:]+(.+?)(?=сведения\s+об\s+авторе|$)',
        full_text,
        re.IGNORECASE | re.DOTALL
    )
    
    if citation_pattern:
        citation_text = citation_pattern.group(1)
        # Проверяем, есть ли название нашего журнала
        if JOURNAL_NAME.lower() not in citation_text.lower():
            # Ищем, какой журнал указан
            journal_match = re.search(r'//\s*(.+?)[\.\d]', citation_text)
            if journal_match:
                found_journal = journal_match.group(1).strip()
                return {
                    "found": False,
                    "journal": found_journal,
                    "expected": JOURNAL_NAME
                }
    
    return {"found": True}

def check_dash_in_bibliography(biblio_text):
    """Проверяет правильность тире в библиографии"""
    errors = []
    
    # Ищем страницы в формате "С. 162—198" (длинное тире)
    long_dash_pattern = re.finditer(r'С\.\s*\d+\s*—\s*\d+', biblio_text)
    
    for match in long_dash_pattern:
        errors.append({
            "type": "warning",
            "category": "Оформление",
            "message": f"Использовано длинное тире вместо среднего. Найдено: '{match.group()}', должно быть: '{match.group().replace('—', '–')}'",
            "paragraph_idx": -1
        })
    
    return errors

def analyze_document(doc: Document, filename: str, use_ai: bool = False) -> dict:
    errors = []
    checklist = []
    paragraphs = doc.paragraphs
    full_text = "\n".join([p.text.strip() for p in paragraphs])
    
    error_id = 0
    error_positions = []  # Для подсветки в тексте

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
        error_positions.append({"paragraph": 0, "error_id": error_id})
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
        error_positions.append({"paragraph": 0, "error_id": error_id})
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

    # 5. Проверка ключевых слов (ИСПРАВЛЕНО!)
    kw_pattern = re.search(r'ключевые\s+слова[.\s\:]*(.+?)(?=для\s+цитирования|$)', full_text, re.IGNORECASE | re.DOTALL)
    if kw_pattern:
        kw_text = kw_pattern.group(1).strip()
        kw_count = count_keywords(kw_text)  # Правильный подсчет
        
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
    footnote_errors, footnote_count, footnote_positions = check_footnotes(doc)
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
    biblio_start_idx = -1
    
    for i, p in enumerate(paragraphs):
        for pattern in biblio_patterns:
            if re.search(pattern, p.text, re.IGNORECASE):
                biblio_found = True
                biblio_start_idx = i
                break
        if biblio_found:
            break
    
    if biblio_found:
        checklist.append({"item": "Список литературы", "status": "success", "details": "Найден"})
        
        # Проверяем тире в библиографии (ИСПРАВЛЕНО!)
        if biblio_start_idx >= 0:
            biblio_text = "\n".join([p.text for p in paragraphs[biblio_start_idx:]])
            dash_errors = check_dash_in_bibliography(biblio_text)
            for err in dash_errors:
                err["id"] = error_id
                errors.append(err)
                error_id += 1
    else:
        checklist.append({"item": "Список литературы", "status": "error", "details": "Не найден"})
        error_id += 1

    # 9. Проверка названия журнала в цитировании (НОВОЕ!)
    journal_check = check_journal_citation(full_text)
    if not journal_check.get("found", True):
        errors.append({
            "id": error_id,
            "type": "error",
            "category": "Цитирование",
            "message": f"В цитировании указан другой журнал: \"{journal_check['journal']}\". Статья подается в \"{JOURNAL_NAME}\".",
            "paragraph_idx": -1
        })
        error_id += 1

    # 10. Проверка орфографии через LanguageTool (ИСПРАВЛЕНО!)
    main_text = extract_main_text(paragraphs)
    if main_text:
        grammar_errors = check_grammar_languagetool(main_text)
        for err in grammar_errors[:20]:
            # Пропускаем библейские сокращения
            context = err.get('context', '')
            skip_error = False
            
            for bible_abbr in BIBLE_ABBREVIATIONS:
                if bible_abbr in context:
                    skip_error = True
                    break
            
            if not skip_error:
                errors.append({
                    "id": error_id, "type": "warning", "category": "Грамматика/Орфография",
                    "message": f"{err['message']} (пример: '{err['context'][:100]}')",
                    "suggestions": err.get('replacements', []),
                    "paragraph_idx": -1
                })
                error_id += 1

    # 11. AI-проверка (если включено)
    ai_analysis = None
    if use_ai and DEEPSEEK_API_KEY:
        try:
            ai_analysis = check_with_deepseek(full_text, main_text)
            if ai_analysis and 'issues' in ai_analysis:
                for issue in ai_analysis.get('issues', [])[:10]:
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

    article_html = generate_article_html_with_errors(paragraphs, errors, footnote_positions)

    return {
        "filename": filename,
        "is_valid": len([e for e in errors if e["type"] == "error"]) == 0,
        "errors": errors,
        "article_html": article_html,
        "checklist": checklist,
        "total_errors": len([e for e in errors if e["type"] == "error"]),
        "total_warnings": len([e for e in errors if e["type"] == "warning"]),
        "ai_analysis": ai_analysis,
        "error_positions": error_positions
    }

def check_grammar_languagetool(text):
    """Проверка орфографии и пунктуации через LanguageTool API"""
    try:
        url = "https://api.languagetool.org/v2/check"
        data = {
            "text": text[:10000],
            "language": "ru-RU",
            "enabledOnly": "false",
            "disabledCategories": "RU_COMPOUND_WORDS"  # Отключаем некоторые ложные срабатывания
        }
        response = requests.post(url, data=data, timeout=10)
        result = response.json()
        
        errors = []
        for match in result.get('matches', []):
            # Пропускаем некоторые типы ошибок
            if match.get('rule', {}).get('id') in ['RU_SENTENCE_WHITESPACE', 'WHITESPACE_RULE']:
                # Проверяем, не является ли это новым абзацем
                context = match.get('context', {}).get('text', '')
                if '.\n' in context or '?\n' in context or '!\n' in context:
                    continue  # Это новый абзац, не ошибка
            
            errors.append({
                "message": match.get('message', ''),
                "context": match.get('context', {}).get('text', ''),
                "replacements": [r.get('value', '') for r in match.get('replacements', [])[:3]],
                "offset": match.get('offset', 0),
                "length": match.get('errorLength', 0)
            })
        
        return errors
    except Exception as e:
        print(f"LanguageTool error: {e}")
        return []

def check_with_deepseek(full_text: str, main_text: str) -> dict:
    """Проверка статьи через DeepSeek AI"""
    
    prompt = f"""Проверь текст статьи на ошибки. Найди:
1. Орфографические и пунктуационные ошибки
2. Грамматические ошибки
3. Стилистические проблемы
4. Логические несоответствия

Текст:
{main_text[:6000]}

Верни JSON:
{{"issues": [{{"type": "error или warning", "message": "описание"}}], "critical_issues": true или false}}

Найди конкретные ошибки. Игнорируй церковнославянизмы и библейские сокращения."""

    try:
        headers = {
            "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
            "Content-Type": "application/json"
        }
        
        payload = {
            "model": "deepseek-chat",
            "messages": [
                {"role": "system", "content": "Ты редактор научного журнала. Отвечай только валидным JSON."},
                {"role": "user", "content": prompt}
            ],
            "temperature": 0.3
        }
        
        response = requests.post(DEEPSEEK_API_URL, json=payload, headers=headers, timeout=30)
        response.raise_for_status()
        
        result = response.json()
        content = result['choices'][0]['message']['content']
        
        try:
            return json.loads(content)
        except:
            return {"issues": [{"type": "warning", "message": content}], "critical_issues": False}
        
    except Exception as e:
        return {"issues": [], "critical_issues": False, "error": str(e)}

def generate_article_html_with_errors(paragraphs, errors, footnote_positions):
    """Генерирует HTML с сохранением форматирования и подсветкой ошибок"""
    html_parts = []
    
    # Создаем словарь ошибок по индексам абзацев
    errors_by_paragraph = {}
    for err in errors:
        p_idx = err.get('paragraph_idx', -1)
        if p_idx not in errors_by_paragraph:
            errors_by_paragraph[p_idx] = []
        errors_by_paragraph[p_idx].append(err)
    
    # Создаем множество позиций сносок для отображения
    footnote_ids = {fp.get('id') for fp in footnote_positions} if footnote_positions else set()
    
    for i, p in enumerate(paragraphs):
        if not p.text.strip():
            continue
        
        # Сохраняем форматирование
        paragraph_html = ""
        for run in p.runs:
            text = run.text
            if not text:
                continue
            
            # Проверяем, есть ли здесь сноска
            if 'w:footnoteReference' in run._element.xml:
                # Добавляем маркер сноски
                footnote_num = len([fp for fp in footnote_positions if fp.get('index') == len([x for x in paragraph_html if x == '†'])])
                paragraph_html += f'<sup style="color: #800020; cursor: pointer;" title="Сноска">†</sup>'
            
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
            
            # Собираем все сообщения об ошибках
            error_messages = [f"{err['category']}: {err['message']}" for err in paragraph_errors]
            tooltip_content = '<br>'.join(error_messages)
            
            html_parts.append(f'''
                <div class="paragraph error {class_type} {error_classes}" 
                     data-error-ids="{','.join(str(err['id']) for err in paragraph_errors)}"
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

# Для запуска локально
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
