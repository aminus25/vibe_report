# -*- coding: utf-8 -*-
"""
report_tools.py — «Parser» и работа с Word в схеме Parser -> LLM -> Assembler.

Содержимое:
  * низкоуровневые помощники python-docx (замена текста поверх runs, клонирование
    абзацев, моноширинный код);
  * Parser:  extract_methodichka_text()  — текст методички из PDF/DOCX/TXT;
             describe_template_for_llm() — список абзацев шаблона с индексами;
  * Шаг 1:   fill_title_page()           — «Лабораторная работа №...» и тема;
  * Шаг 3:   add_images_with_captions()  — скриншоты + подписи «Рисунок N — ...».

Картинки можно передавать и путями к файлам, и кортежами (имя, байты) — второе
удобно для Streamlit, где файлы приходят из file_uploader.
"""

from __future__ import annotations

import copy
import json
import re
from io import BytesIO
from pathlib import Path
from typing import Callable, Optional, Sequence, Union

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Emu, Pt
from docx.table import Table
from docx.text.paragraph import Paragraph
from PIL import Image, ImageOps

# Функция-обёртка над LLM: (system_prompt, user_message) -> текст ответа
LLMFunc = Callable[[str, str], str]
# Картинка: путь к файлу ИЛИ (имя_файла, байты)
ImageItem = Union[str, Path, tuple]

CODE_FONT = "Courier New"


# =========================================================================== #
#  Низкоуровневые помощники
# =========================================================================== #

def paragraph_text(par: Paragraph) -> str:
    """Текст абзаца, склеенный из всех runs (Word дробит текст на куски)."""
    return "".join(run.text for run in par.runs)


def replace_span(par: Paragraph, start: int, end: int, new_text: str):
    """
    Заменяет символы [start, end) склеенного текста абзаца на new_text.

    Новый текст получает форматирование run, в котором начинался заменяемый
    фрагмент; из остальных затронутых runs вырезается только их часть.
    Соседние runs не трогаются. Если start == end — вставка в позицию start.
    Возвращает run, получивший новый текст (или None).
    """
    runs = par.runs
    if start == end:
        pos = 0
        for run in runs:
            n = len(run.text)
            if n and pos <= start <= pos + n:
                cut = start - pos
                run.text = run.text[:cut] + new_text + run.text[cut:]
                return run
            pos += n
        return None

    first_run, pos = None, 0
    for run in runs:
        text = run.text
        run_start, run_end = pos, pos + len(text)
        pos = run_end
        if run_end <= start or run_start >= end:
            continue
        local_start = max(start - run_start, 0)
        local_end = min(end - run_start, len(text))
        if first_run is None:
            run.text = text[:local_start] + new_text + text[local_end:]
            first_run = run
        else:
            run.text = text[local_end:]
    return first_run


def make_shell(ref_par: Paragraph) -> Paragraph:
    """Пустой абзац с теми же свойствами (<w:pPr>), что у ref_par: без содержимого,
    без разрыва раздела, нумерации списка и «с новой страницы»."""
    new_p = copy.deepcopy(ref_par._p)
    for child in list(new_p):
        if child.tag != qn("w:pPr"):
            new_p.remove(child)
    pPr = new_p.pPr
    if pPr is not None:
        for tag in ("w:sectPr", "w:numPr", "w:pageBreakBefore"):
            for el in pPr.findall(qn(tag)):
                pPr.remove(el)
    return Paragraph(new_p, ref_par._parent)


def first_run_rpr(par: Paragraph, drop_bold: bool = False):
    """Копия <w:rPr> первого непустого run (шрифт, размер, цвет)."""
    for run in par.runs:
        if run.text.strip() and run._r.rPr is not None:
            rpr = copy.deepcopy(run._r.rPr)
            if drop_bold:
                for tag in ("w:b", "w:bCs"):
                    for el in rpr.findall(qn(tag)):
                        rpr.remove(el)
            return rpr
    return None


def paragraph_mark_rpr(par: Paragraph):
    """Копия свойств шрифта «знака абзаца» (<w:pPr><w:rPr>) — это «обычное»
    форматирование ввода в данном абзаце, даже если он пустой."""
    pPr = par._p.pPr
    if pPr is None:
        return None
    mark = pPr.find(qn("w:rPr"))
    if mark is None:
        return None
    rpr = copy.deepcopy(mark)
    for tag in ("w:ins", "w:del", "w:moveFrom", "w:moveTo", "w:rPrChange"):
        for el in rpr.findall(qn(tag)):
            rpr.remove(el)
    return rpr


def apply_rpr(run, rpr) -> None:
    """Применяет сохранённые свойства шрифта к новому run."""
    if rpr is None:
        return
    if run._r.rPr is not None:
        run._r.remove(run._r.rPr)
    run._r.insert(0, copy.deepcopy(rpr))


def set_run_font(run, name: str, size_pt: float) -> None:
    """Шрифт для run. Для кириллицы проставляем все 4 атрибута и убираем
    theme-атрибуты (они перекрывают явно заданный шрифт)."""
    run.font.name = name
    rFonts = run._r.rPr.rFonts
    rFonts.set(qn("w:eastAsia"), name)
    rFonts.set(qn("w:cs"), name)
    for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rFonts.get(qn(attr)) is not None:
            del rFonts.attrib[qn(attr)]
    run.font.size = Pt(size_pt)


def set_paragraph_mark_font(par: Paragraph, name: str, size_pt: float) -> None:
    """Шрифт «знака абзаца»: пустые строки кода получают ту же высоту, что и
    непустые (иначе берётся размер из шаблона, например 14 пт)."""
    pPr = par._p.get_or_add_pPr()
    for old in pPr.findall(qn("w:rPr")):
        pPr.remove(old)
    rPr = OxmlElement("w:rPr")
    rFonts = OxmlElement("w:rFonts")
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), name)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(size_pt * 2)))
    rPr.append(rFonts)
    rPr.append(sz)
    pPr.append(rPr)


def strip_code_fences(code: str) -> str:
    """Убирает ```lang ... ``` вокруг кода, если LLM всё же их добавила."""
    code = code.strip("\n")
    code = re.sub(r"^\s*```[\w+#-]*\s*\n", "", code)
    code = re.sub(r"\n\s*```\s*$", "", code)
    return code


def build_code_paragraphs(ref_par: Paragraph, code: str, size_pt: float = 10,
                          font: str = CODE_FONT) -> Paragraph:
    """
    Вставляет код после ref_par: каждая строка — отдельный абзац, моноширинный
    шрифт, без отступов, одинарный интервал, выравнивание влево.
    Возвращает последний вставленный абзац.
    """
    code = strip_code_fences(code)
    lines = code.replace("\r\n", "\n").replace("\r", "\n").expandtabs(4).split("\n")
    while lines and not lines[0].strip():
        lines.pop(0)
    while lines and not lines[-1].strip():
        lines.pop()
    anchor = ref_par
    for line in lines:
        p = make_shell(ref_par)
        pf = p.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.first_line_indent = Cm(0)
        pf.left_indent = Cm(0)
        pf.line_spacing = 1.0
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        set_run_font(p.add_run(line.rstrip()), font, size_pt)
        set_paragraph_mark_font(p, font, size_pt)
        anchor._p.addnext(p._p)
        anchor = p
    return anchor


def parse_json_object(raw: str) -> dict:
    """
    Достаёт JSON-объект из ответа LLM. Порядок попыток: весь ответ -> содержимое
    обёртки ```json ... ``` -> от первой «{» до последней «}». Ограждения ``` ВНУТРИ
    строковых значений (например, в коде) не трогаем. strict=False разрешает
    «сырые» переводы строк и табуляции внутри строк — так часто отвечают LLM с кодом.
    """
    raw = (raw or "").strip()
    candidates = [raw]
    m = re.match(r"^```(?:json)?\s*(.*?)\s*```$", raw, re.S | re.I)
    if m:
        candidates.append(m.group(1))
    start, end = raw.find("{"), raw.rfind("}")
    if start != -1 and end > start:
        candidates.append(raw[start:end + 1])
    for candidate in candidates:
        try:
            data = json.loads(candidate, strict=False)
        except json.JSONDecodeError:
            continue
        if isinstance(data, dict):
            return data
    return {}


# =========================================================================== #
#  PARSER 1: текст методички
# =========================================================================== #

def extract_methodichka_text(filename: str, data: bytes, max_chars: int = 40000) -> tuple:
    """Обёртка: любые ошибки разбора файла превращаются в понятный ValueError."""
    try:
        return _extract_methodichka_text(filename, data, max_chars)
    except ValueError:
        raise
    except Exception as exc:
        raise ValueError(f"Не удалось прочитать методичку «{filename}»: файл повреждён или защищён паролем "
                         f"({type(exc).__name__}).") from exc


def _extract_methodichka_text(filename: str, data: bytes, max_chars: int) -> tuple:
    """
    Извлекает текст методички из PDF / DOCX / TXT / MD.
    Возвращает (текст, был_ли_обрезан). Длинные методички режем до max_chars
    (≈10–13 тыс. токенов), чтобы уложиться в лимиты бесплатного API.
    """
    ext = Path(filename).suffix.lower()
    if ext == ".pdf":
        import pdfplumber                       # pypdf склеивает слова в русских PDF — берём pdfminer
        chunks, total = [], 0
        with pdfplumber.open(BytesIO(data)) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text() or ""
                chunks.append(page_text)
                total += len(page_text)
                if total > max_chars * 1.2:     # дальше читать нет смысла: всё равно обрежем
                    break
        text = "\n".join(chunks)
        text = re.sub(r"-\n(?=[a-zа-яё])", "", text)      # перенос слова: «опера-\nционная»
    elif ext == ".docx":
        doc = Document(BytesIO(data))
        parts = []
        for el in doc.element.body.iterchildren():
            if el.tag == qn("w:p"):
                parts.append(Paragraph(el, doc._body).text)
            elif el.tag == qn("w:tbl"):
                seen = set()
                for row in Table(el, doc._body).rows:
                    cells = []
                    for cell in row.cells:
                        if cell._tc not in seen:               # объединённые ячейки — один раз
                            seen.add(cell._tc)
                            cells.append(cell.text.strip())
                    parts.append(" | ".join(cells))
        text = "\n".join(parts)
    elif ext in (".txt", ".md"):
        try:
            text = data.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = data.decode("cp1251", errors="replace")
    else:
        raise ValueError(f"Неподдерживаемый формат методички: {ext or filename}")

    text = "\n".join(line.strip() for line in text.splitlines())
    text = re.sub(r"\n{3,}", "\n\n", text).strip()
    if not text:
        raise ValueError("В методичке не найден текст (возможно, это скан без распознавания).")
    if len(text) > max_chars:
        return text[:max_chars], True
    return text, False


# =========================================================================== #
#  PARSER 2: структура шаблона для LLM
# =========================================================================== #

def is_bold_paragraph(par: Paragraph) -> bool:
    runs = [r for r in par.runs if r.text.strip()]
    return bool(runs) and all(r.bold for r in runs)


def describe_template_for_llm(doc, max_chars: int = 140) -> str:
    """
    Строки вида:
        [30] TEXT,BOLD,PAGE_BREAK | 1. Цель работы
        [31] EMPTY |
    Индекс — позиция в doc.paragraphs (абзацы внутри таблиц не нумеруются).
    """
    lines = []
    for i, par in enumerate(doc.paragraphs):
        text = " ".join(par.text.split())
        flags = ["TEXT" if text else "EMPTY"]
        if text and is_bold_paragraph(par):
            flags.append("BOLD")
        if par.style is not None and par.style.name.lower().startswith(("heading", "заголовок")):
            flags.append("HEADING_STYLE")
        if par._p.xpath(".//w:br[@w:type='page']"):
            flags.append("PAGE_BREAK")
        if par._p.xpath(".//w:drawing | .//w:pict"):
            flags.append("IMAGE")
        if len(text) > max_chars:
            text = text[:max_chars] + "…"
        lines.append(f"[{i}] {','.join(flags)} | {text}")
    return "\n".join(lines)


# =========================================================================== #
#  ШАГ 1. Титульный лист: «Лабораторная работа №...» и тема
# =========================================================================== #

LAB_RE = re.compile(
    r"Лабораторн\w*\s+работ\w*\s*(?:(?:№|N|No\.?)\s*)?"
    r"(?P<num>\d+(?:[.\-]\d+)*[A-Za-zА-Яа-я]?|[_.…]+|[xXхХ]+(?!\w))",
    re.IGNORECASE,
)
COURSE_RE = re.compile(r"^\s*(по\s+(курсу|дисциплине|предмету)|дисциплина|курс|предмет)\b", re.IGNORECASE)
STOP_RE = re.compile(
    r"^\s*(выполнил[аи]?|выполнено|студент\w*|принял[аи]?|проверил[аи]?|"
    r"преподаватель|руководитель|группа|гр\.|подпись|оценка|\w+,\s*\d{4}|г\.\s)",
    re.IGNORECASE,
)
TOPIC_RE = re.compile(
    r"^(\s*(?:на\s+тему|тема(?:\s+работы)?)\s*[:\-–—]?\s*)?([«\"“„]?)\s*(?P<body>.*?)\s*([»\"”“]?)\s*$",
    re.IGNORECASE | re.DOTALL,
)


def first_page_paragraphs(doc, limit: int = 60) -> list:
    """Абзацы первой страницы (включая таблицы): до разрыва страницы,
    «с новой страницы» или разрыва раздела."""
    result = []
    for p_el in doc.element.body.iter(qn("w:p")):
        par = Paragraph(p_el, doc._body)
        if p_el.xpath("./w:pPr/w:pageBreakBefore"):
            break
        before, hit_break = "", False
        for run in par.runs:
            if run._r.xpath("./w:br[@w:type='page']"):
                hit_break = True
                break
            before += run.text
        if hit_break:
            if before.strip():
                result.append(par)
            break
        result.append(par)
        if p_el.xpath("./w:pPr/w:sectPr") or len(result) >= limit:
            break
    return result


def _replace_topic_in_paragraph(par: Paragraph, topic: str, offset: int = 0) -> bool:
    """Заменяет «тело» темы (без метки «Тема:» и без кавычек)."""
    full = paragraph_text(par)
    m = TOPIC_RE.match(full[offset:])
    if not m:
        return False
    start, end = m.span("body")
    return replace_span(par, offset + start, offset + end, topic) is not None or start == end


def fill_title_page(doc, number, topic: str, insert_topic_if_missing: bool = True) -> dict:
    """
    Находит на первой странице «Лабораторная работа №...» и подставляет номер и тему.

    Тема ищется: (1) в той же строке после номера; (2) в ближайшем следующем
    непустом абзаце, пропуская «по курсу/дисциплине…» (метка «на тему:» и кавычки
    «» сохраняются); (3) если темы нет — вставляется новый абзац.
    Стиль берётся у run, где стоял старый номер/тема.
    Возвращает {"number_replaced": bool, "topic_mode": "inline|separate|inserted|skipped"}.
    """
    number = str(number).strip().lstrip("№").strip()
    topic = " ".join(str(topic).split())

    paragraphs = first_page_paragraphs(doc)
    lab_idx, m = None, None
    for i, par in enumerate(paragraphs):
        m = LAB_RE.search(paragraph_text(par))
        if m:
            lab_idx = i
            break
    if lab_idx is None:
        raise ValueError("На первой странице не найдена строка «Лабораторная работа №...»")

    lab_par = paragraphs[lab_idx]
    report = {"number_replaced": False, "topic_mode": "skipped"}

    if number:
        n_start, n_end = m.span("num")
        replace_span(lab_par, n_start, n_end, number)
        report["number_replaced"] = True
    else:
        n_start, number = m.start("num"), m.group("num")      # номер не меняем
    if not topic:
        return report

    # 1) тема в той же строке
    after_start = n_start + len(number)
    after = paragraph_text(lab_par)[after_start:]
    stripped = re.sub(r"^[\s.:,;\-–—]+", "", after)
    if stripped.strip() and not COURSE_RE.match(stripped):
        offset = after_start + (len(after) - len(stripped))
        if _replace_topic_in_paragraph(lab_par, topic, offset):
            report["topic_mode"] = "inline"
            return report

    # 2) тема в следующем абзаце
    anchor = lab_par
    for par in paragraphs[lab_idx + 1:]:
        text = paragraph_text(par)
        if not text.strip():
            continue
        if COURSE_RE.match(text):
            anchor = par
            continue
        if STOP_RE.match(text):
            break
        if _replace_topic_in_paragraph(par, topic):
            report["topic_mode"] = "separate"
            return report
        break

    # 3) темы в шаблоне нет — вставляем свой абзац
    if insert_topic_if_missing:
        new_par = make_shell(lab_par)
        new_par.paragraph_format.space_before = Pt(12)
        apply_rpr(new_par.add_run(topic), first_run_rpr(lab_par))
        anchor._p.addnext(new_par._p)
        report["topic_mode"] = "inserted"
    return report


# =========================================================================== #
#  ШАГ 3. Картинки + подписи
# =========================================================================== #

DEFAULT_TITLE = "Результат выполнения"

CAPTION_SYSTEM_PROMPT = """Ты помогаешь оформлять отчёты по лабораторным работам в российском вузе.
На вход приходит JSON-массив объектов {"id": номер, "filename": имя файла скриншота}.
Для каждого файла придумай короткое название рисунка по смыслу имени файла.

Правила:
1. Название — именная группа на русском языке, 2–8 слов, с заглавной буквы, без точки в конце.
   Примеры: «Результаты тестирования», «Окно выбора языка установки», «Вывод команды ip a».
2. НЕ добавляй слово «Рисунок» и номер — их добавит программа.
3. Опирайся только на имя файла. Не выдумывай подробностей, которых в имени нет.
   Имена на английском переводи на русский; команды и названия программ оставляй как есть.
4. Если имя неинформативно (IMG_001, Screenshot 2026-03-01, «снимок экрана») — верни пустую строку "".
5. Ответ — ТОЛЬКО JSON-объект без пояснений и без markdown: {"1": "название", "2": "название", ...}.
   Ключи — те же id, по одному на каждый файл."""


def _item_name(item: ImageItem) -> str:
    return str(item[0]) if isinstance(item, tuple) else Path(item).name


def _item_bytes(item: ImageItem) -> bytes:
    return item[1] if isinstance(item, tuple) else Path(item).read_bytes()


def natural_sort_images(items: Sequence[ImageItem]) -> list:
    """Сортировка по имени «по-человечески»: 2.png идёт перед 10.png."""
    def key(item):
        return [int(t) if t.isdigit() else t.lower()
                for t in re.split(r"(\d+)", _item_name(item))]
    return sorted(items, key=key)


natural_sort_paths = natural_sort_images        # старое имя


def caption_from_filename(path) -> str:
    """Запасной вариант без LLM: «03_результаты_тестирования.png» -> «Результаты тестирования».
    Системные имена («Снимок экрана 2026-03-01 в 12.00.00», «Screenshot ... at ...», «IMG_0001»)
    дают нейтральное «Результат выполнения»."""
    stem = Path(str(path)).stem
    stem = re.sub(r"^\s*\d+\s*[-_.)\s]+", "", stem)                                   # «01_», «3. »
    stem = re.sub(r"\d{2,4}[-.:]\d{2}[-.:]\d{2,4}|\d{1,2}[.:]\d{2}[.:]\d{2}", " ", stem)  # дата/время
    stem = re.sub(r"[_\-]+", " ", stem)
    stem = re.sub(r"\s+", " ", stem).strip()
    if re.fullmatch(r"(?i)(screenshot|screen shot|скриншот|снимок экрана|img|image|photo|фото|pic)?"
                    r"[\s\d()_.]*(?:(?:в|at)[\s\d()_.]*)?", stem):
        return DEFAULT_TITLE
    return stem[0].upper() + stem[1:]


def _clean_title(title: str) -> str:
    """Чистим ответ LLM: кавычки, «Рисунок N —» и точка в конце."""
    title = " ".join(str(title).split()).strip(" \"'«»“”")
    title = re.sub(r"^(рисунок|рис\.?)\s*\d*\s*[-–—:.]*\s*", "", title, flags=re.IGNORECASE)
    title = title.rstrip(".").strip()
    return title[0].upper() + title[1:] if title else ""


def generate_captions(items: Sequence[ImageItem], llm: Optional[LLMFunc] = None,
                      context: str = "") -> list:
    """
    Названия рисунков (без «Рисунок N»), по одному на файл. Все имена уходят в LLM
    ОДНИМ запросом; для файлов, по которым LLM ничего не вернул (или упала) —
    эвристика по имени файла.
    """
    names = [_item_name(i) for i in items]
    titles = [""] * len(names)
    if llm is not None and names:
        payload = [{"id": i + 1, "filename": n} for i, n in enumerate(names)]
        user = json.dumps(payload, ensure_ascii=False)
        if context:
            user = f"Тема лабораторной работы: {context}\n{user}"
        try:
            data = parse_json_object(llm(CAPTION_SYSTEM_PROMPT, user))
            titles = [_clean_title(data.get(str(i + 1), "")) for i in range(len(names))]
        except Exception:
            titles = [""] * len(names)             # сеть/ключ/лимиты — не роняем отчёт
    return [t or caption_from_filename(n) for t, n in zip(titles, names)]


def _docx_can_read(raw: bytes) -> bool:
    """Умеет ли САМ python-docx распознать эти байты. Он строже Pillow: например, JPEG
    принимает только с маркером JFIF/Exif в заголовке, а Pillow читает и такие, и без них."""
    try:
        from docx.image.image import Image as DocxImage
        DocxImage.from_blob(raw)
        return True
    except Exception:
        return False


def _load_image_bytes(item: ImageItem) -> tuple:
    """
    Готовит картинку к вставке в Word. Возвращает (bytes, (ширина_px, высота_px)).
    Исходные байты используем, только если python-docx их распознаёт и поворот по EXIF
    не нужен. Иначе (WebP, JPEG без JFIF/Exif, MPO, фото «на боку» и т.п.) перекодируем в PNG.
    """
    name = _item_name(item)
    raw = _item_bytes(item)
    try:
        img = Image.open(BytesIO(raw))
        img.load()                                    # заодно ловим обрезанные/битые файлы
    except Exception as exc:
        raise ValueError(f"Не удалось прочитать изображение «{name}»: файл повреждён или формат "
                         f"не поддерживается ({type(exc).__name__}).") from exc

    with img:
        needs_rotation = img.getexif().get(0x0112, 1) != 1
        if not needs_rotation and _docx_can_read(raw):
            return raw, img.size
        try:
            fixed = ImageOps.exif_transpose(img)
            has_alpha = fixed.mode in ("RGBA", "LA", "PA") or "transparency" in fixed.info
            fixed = fixed.convert("RGBA" if has_alpha else "RGB")
            buf = BytesIO()
            fixed.save(buf, "PNG")
        except Exception as exc:
            raise ValueError(f"Не удалось преобразовать изображение «{name}» в PNG "
                             f"({type(exc).__name__}).") from exc
        return buf.getvalue(), fixed.size


def _pick_reference_paragraph(doc) -> Paragraph:
    """Образец оформления: последний абзац с «настоящим» текстом (≥20 символов)."""
    for par in reversed(doc.paragraphs):
        if len(par.text.strip()) >= 20:
            return par
    return doc.paragraphs[-1] if doc.paragraphs else doc.add_paragraph()


def add_images_with_captions(
    doc,
    images: Sequence[ImageItem],
    *,
    llm: Optional[LLMFunc] = None,
    captions: Optional[Sequence[str]] = None,
    context: str = "",
    after: Optional[Paragraph] = None,
    reference_paragraph: Optional[Paragraph] = None,
    caption_format: str = "Рисунок {n} — {title}",
    start_number: int = 1,
    max_width_cm: float = 15.0,
    max_height_cm: float = 20.0,
) -> list:
    """
    Вставляет картинки по центру, под каждой — подпись по центру.

    images              — пути к файлам или (имя, байты); порядок = порядок рисунков.
    llm                 — функция для генерации названий по именам файлов (None -> эвристика).
    captions            — готовые названия (если заданы, LLM не вызывается).
    after               — абзац, ПОСЛЕ которого вставлять (None -> в конец документа).
    reference_paragraph — образец шрифта/интервалов (None -> последний абзац с текстом).
    Возвращает список (имя_файла, итоговая_подпись).
    """
    images = list(images)
    if not images:
        return []

    titles = list(captions) if captions is not None else generate_captions(images, llm, context)
    if len(titles) != len(images):
        raise ValueError("Число подписей не совпадает с числом картинок")

    ref = reference_paragraph or _pick_reference_paragraph(doc)
    caption_rpr = first_run_rpr(ref, drop_bold=True)
    if caption_rpr is None:
        caption_rpr = paragraph_mark_rpr(ref)

    sec = doc.sections[-1]
    try:
        page_text_width = int(sec.page_width - sec.left_margin - sec.right_margin)
    except TypeError:
        page_text_width = int(Cm(16))
    max_w = min(page_text_width, int(Cm(max_width_cm)))

    def insert(new_par: Paragraph, anchor: Optional[Paragraph]) -> Paragraph:
        if anchor is not None:
            anchor._p.addnext(new_par._p)
        else:
            body = doc.element.body
            if body.sectPr is not None:
                body.sectPr.addprevious(new_par._p)
            else:
                body.append(new_par._p)
        return new_par

    anchor, result = after, []
    for offset, (item, title) in enumerate(zip(images, titles)):
        n = start_number + offset
        data, (w_px, h_px) = _load_image_bytes(item)

        width = min(max_w, w_px * 9525)                  # не больше натурального размера (96 dpi)
        if width * h_px / w_px > Cm(max_height_cm):
            width = Cm(max_height_cm) * w_px / h_px

        pic = make_shell(ref)
        pf = pic.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf.first_line_indent = Cm(0)
        pf.left_indent = Cm(0)
        pf.line_spacing = 1.0                            # «точный» интервал обрезал бы картинку
        pf.space_before = Pt(6)
        pf.space_after = Pt(6)
        pf.keep_with_next = True                         # картинка и подпись на одной странице
        pic.add_run().add_picture(BytesIO(data), width=Emu(int(width)))
        anchor = insert(pic, anchor)

        cap = make_shell(ref)
        cf = cap.paragraph_format
        cf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cf.first_line_indent = Cm(0)
        cf.left_indent = Cm(0)
        cf.space_after = Pt(12)
        text = caption_format.replace("{n}", str(n)).replace("{title}", title)
        apply_rpr(cap.add_run(text), caption_rpr)
        anchor = insert(cap, anchor)

        result.append((_item_name(item), text))
    return result