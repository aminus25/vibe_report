# -*- coding: utf-8 -*-
"""
assembler.py — «Assembler» и оркестратор всей цепочки Parser -> LLM -> Assembler.

generate_report() — единственная точка входа для интерфейса:
    1. Parser:    снимок абзацев шаблона + текст методички
    2. Титульник: fill_title_page()                       (без LLM)
    3. LLM:       JSON-план {"paragraph_N": "..."}        (1 запрос)
    4. Assembler: вставка текста/кода по плану, сохраняя стиль абзаца
    5. Скрины:    подписи (1 запрос к LLM) + вставка в раздел «Результаты»

Режимы кода:
    (а) студент дал код   -> LLM лишь указывает абзац (возвращает метку),
                             код вставляется БЕЗ изменений;
    (б) кода нет          -> LLM пишет код по методичке (и варианту).

Безопасность вставки: Assembler не затирает настоящий текст шаблона, заголовки
и титульный лист — такие указания LLM пропускаются с пояснением в журнале.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from io import BytesIO
from typing import Optional

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

import report_tools as rt
from llm_client import STUDENT_CODE_SENTINEL, build_user_message, request_plan

# --------------------------------------------------------------------------- #
#  Эвристики разметки шаблона
# --------------------------------------------------------------------------- #

# Заглушка: пусто, «…», «___», «{{ЦЕЛЬ}}», «[вставьте текст]», «<текст>», «Здесь должен быть…»
PLACEHOLDER_RE = re.compile(
    r"^\s*(?:[\s._…\-–—]*|\{\{.*?\}\}|\[[^\]]*\]|<[^>]*>|текст|"
    r"(?:здесь|тут)\b.*|(?:вставьте|укажите|заполните|введите|впишите|напишите)\b.*)\s*$",
    re.IGNORECASE | re.DOTALL,
)
CODE_HEADING_RE = re.compile(r"листинг|исходн\w*\s+код|код\w*\s+программ|программн\w*\s+(?:код|реализац)|\bкод\b", re.IGNORECASE)
RESULTS_RE = re.compile(r"результат|скриншот|ход\s+работ|порядок\s+выполнен", re.IGNORECASE)
CONCLUSION_RE = re.compile(r"вывод|заключени", re.IGNORECASE)


def is_placeholder(text: str) -> bool:
    text = text.strip()
    return not text or (len(text) <= 120 and bool(PLACEHOLDER_RE.match(text)))


def is_heading(par: Paragraph) -> bool:
    """Короткий абзац, целиком жирный или в стиле «Заголовок»."""
    text = par.text.strip()
    if not text or len(text) > 120:
        return False
    if par.style is not None and par.style.name.lower().startswith(("heading", "заголовок")):
        return True
    return rt.is_bold_paragraph(par)


def heading_above(paragraphs: list, idx: int) -> str:
    """Текст ближайшего заголовка выше абзаца idx."""
    for j in range(idx - 1, -1, -1):
        if is_heading(paragraphs[j]):
            return paragraphs[j].text.strip()
    return ""


def looks_like_code(text: str) -> bool:
    """Страховка: многострочный текст, в котором большинство строк похожи на код.
    Пункты нумерованных списков («1. Изучить…») за код не считаются."""
    lines = [ln for ln in text.splitlines() if ln.strip()]
    if len(lines) < 3:
        return False
    code_re = re.compile(
        r"[{};]|==?|\):\s*$|^\s{2,}\S|"
        r"^\s*(?:def|class|import|from|for|if|while|return|int|void|print|printf|cout|std::|#include)\b")
    hits = sum(1 for ln in lines if not re.match(r"^\s*\d+[.)]\s", ln) and code_re.search(ln))
    return hits / len(lines) >= 0.6


def title_region(doc) -> set:
    """Элементы <w:p> титульного листа (до первого разрыва). Если разрыва в документе
    нет — пустое множество (защита титульника отключается)."""
    pars = rt.first_page_paragraphs(doc, limit=10 ** 6)
    total = sum(1 for _ in doc.element.body.iter(qn("w:p")))
    return {p._p for p in pars} if len(pars) < total else set()


def find_code_slot(paragraphs: list, title: set) -> Optional[int]:
    """Запасной поиск абзаца под код: первая заглушка под заголовком «Листинг/Код…»."""
    for i, par in enumerate(paragraphs):
        if par._p in title or not (is_heading(par) and CODE_HEADING_RE.search(par.text)):
            continue
        for j in range(i + 1, min(i + 4, len(paragraphs))):
            if is_heading(paragraphs[j]):
                break
            if is_placeholder(paragraphs[j].text):
                return j
        break
    return None


def find_images_anchor(paragraphs: list, title: set) -> tuple:
    """
    Куда класть скриншоты. Возвращает (after, placeholder_to_remove):
      1) раздел «Результаты / Ход работы…»: если под заголовком заглушка — на её место,
         иначе в конец раздела;
      2) нет такого раздела — перед заголовком «Выводы»;
      3) иначе — в конец документа (after=None).
    """
    n = len(paragraphs)
    for h, par in enumerate(paragraphs):
        if par._p in title or not (is_heading(par) and RESULTS_RE.search(par.text)):
            continue
        nxt = paragraphs[h + 1] if h + 1 < n else None
        if nxt is not None and is_placeholder(nxt.text) and not is_heading(nxt):
            return par, nxt
        end = h
        for k in range(h + 1, n):
            if is_heading(paragraphs[k]):
                break
            end = k
        return paragraphs[end], None
    for c, par in enumerate(paragraphs):
        if par._p not in title and is_heading(par) and CONCLUSION_RE.search(par.text) and c > 0:
            return paragraphs[c - 1], None
    return None, None


# --------------------------------------------------------------------------- #
#  Assembler: вставка по плану
# --------------------------------------------------------------------------- #

def _typing_rpr(par: Paragraph):
    """Шрифт «обычного ввода» в абзаце: знак абзаца, иначе первый run (без подсветки)."""
    rpr = rt.paragraph_mark_rpr(par)
    if rpr is None:
        rpr = rt.first_run_rpr(par)
        if rpr is not None:
            for el in rpr.findall(qn("w:highlight")):
                rpr.remove(el)
    return rpr


def fill_text_paragraph(par: Paragraph, text: str) -> None:
    """Подставляет текст в абзац-образец: каждая строка — отдельный абзац-клон
    (стиль, отступы и шрифт как у образца), пустые строки отбрасываются."""
    lines = [ln.strip() for ln in text.replace("\r", "").split("\n") if ln.strip()]
    if not lines:
        return
    rpr = _typing_rpr(par)
    for child in list(par._p):                        # убираем заглушку, оставляя <w:pPr>
        if child.tag != qn("w:pPr"):
            par._p.remove(child)
    anchor = par
    for i, line in enumerate(lines):
        target = par
        if i > 0:
            target = rt.make_shell(par)
            anchor._p.addnext(target._p)
            anchor = target
        rt.apply_rpr(target.add_run(line), rpr)


def apply_plan(paragraphs: list, plan: dict, *, student_code: str, code_size_pt: float,
               title: set) -> tuple:
    """
    Вставляет тексты по плану. Возвращает (журнал, множество_затронутых_<w:p>).
    Запись журнала: {"index", "section", "kind": text|code, "status": ok|skipped, "note"}.
    """
    log, used = [], set()
    for idx in sorted(plan):
        value = plan[idx]
        entry = {"index": idx, "section": "", "kind": "text", "status": "skipped", "note": ""}
        log.append(entry)

        if not 0 <= idx < len(paragraphs):
            entry["note"] = "индекс вне документа"
            continue
        par = paragraphs[idx]
        entry["section"] = heading_above(paragraphs, idx)

        is_code = (value.strip() == STUDENT_CODE_SENTINEL
                   or bool(CODE_HEADING_RE.search(entry["section"]))
                   or looks_like_code(value))
        entry["kind"] = "code" if is_code else "text"

        # --- что можно трогать ---
        if par._p in title:
            entry["note"] = "абзац на титульном листе — не трогаем"
            continue
        text = par.text.strip()
        overwritten = False
        if not is_placeholder(text):
            if is_heading(par):
                entry["note"] = "LLM указала заголовок — не затираем"
                continue
            if len(text) > 150:
                entry["note"] = "в абзаце уже есть содержательный текст — не затираем"
                continue
            overwritten = True

        # --- что вставлять ---
        if is_code:
            if student_code.strip():
                payload = student_code                 # режим (а): без изменений
            elif value.strip() == STUDENT_CODE_SENTINEL:
                entry["note"] = "LLM вернула метку кода, но кода студента нет"
                continue
            else:
                payload = rt.strip_code_fences(value)  # режим (б): код от LLM
            rt.build_code_paragraphs(par, payload, code_size_pt)
            par._p.getparent().remove(par._p)          # заглушка больше не нужна
        else:
            fill_text_paragraph(par, value)

        used.add(par._p)
        entry["status"] = "ok"
        if overwritten:
            entry["note"] = "заменён короткий существующий текст"
    return log, used


# --------------------------------------------------------------------------- #
#  Оркестратор
# --------------------------------------------------------------------------- #

@dataclass
class ReportResult:
    docx_bytes: bytes
    plan: dict = field(default_factory=dict)
    raw_llm_answer: str = ""
    log: list = field(default_factory=list)
    warnings: list = field(default_factory=list)
    title_report: dict = field(default_factory=dict)
    images: list = field(default_factory=list)       # [(имя файла, подпись)]
    template_dump: str = ""


def generate_report(
    template_bytes: bytes,
    method_text: str,
    images: list,
    *,
    llm,
    number: str = "",
    topic: str = "",
    variant: str = "",
    student_code: str = "",
    code_size_pt: float = 10,
    caption_format: str = "Рисунок {n} — {title}",
    llm_captions: bool = True,
    sort_images: bool = True,
) -> ReportResult:
    """images — список (имя_файла, байты). llm — функция (system, user) -> str."""
    warnings: list = []
    try:
        doc = Document(BytesIO(template_bytes))
    except Exception as exc:
        raise ValueError("Не удалось открыть шаблон: файл повреждён или это не .docx "
                         f"({type(exc).__name__}).") from exc

    # 1. PARSER: снимок абзацев ДО любых вставок — индексы LLM ссылаются на него
    paragraphs = list(doc.paragraphs)
    dump = rt.describe_template_for_llm(doc)
    title = title_region(doc)

    # 2. Титульник (без LLM; не фатально)
    title_report: dict = {}
    if number or topic:
        try:
            title_report = rt.fill_title_page(doc, number, topic)
            if topic and title_report.get("topic_mode") == "inserted":
                warnings.append("В шаблоне не нашлось строки с темой — тема добавлена отдельным абзацем.")
        except ValueError as exc:
            warnings.append(f"Титульный лист: {exc}")

    # 3. LLM: план
    if sort_images:
        images = rt.natural_sort_images(images)
    message = build_user_message(
        dump, method_text, number=number, topic=topic, variant=variant.strip(),
        has_student_code=bool(student_code.strip()),
        screenshot_names=[rt._item_name(i) for i in images] or None)
    plan, plan_warnings, raw = request_plan(llm, message)
    warnings += plan_warnings
    if not plan:
        warnings.append("LLM не вернула корректный план заполнения — текстовые разделы остались пустыми.")

    # Режим (а): если LLM не нашла раздел кода, ищем его сами
    if student_code.strip() and not any(
            v.strip() == STUDENT_CODE_SENTINEL or CODE_HEADING_RE.search(heading_above(paragraphs, i))
            for i, v in plan.items() if 0 <= i < len(paragraphs)):
        slot = find_code_slot(paragraphs, title)
        if slot is not None and slot not in plan:
            plan[slot] = STUDENT_CODE_SENTINEL
            warnings.append("LLM не указала раздел для кода — найден по заголовку «Листинг/Код».")
        else:
            warnings.append("В шаблоне не найден раздел для кода — код не вставлен.")

    # 4. ASSEMBLER: текст и код
    log, used = apply_plan(paragraphs, plan, student_code=student_code,
                           code_size_pt=code_size_pt, title=title)
    for e in log:
        if e["status"] == "skipped":
            warnings.append(f"Абзац {e['index']} ({e['section'] or 'без заголовка'}): {e['note']}.")

    # 5. Скриншоты
    placed: list = []
    if images:
        after, placeholder = find_images_anchor(paragraphs, title)
        if placeholder is not None and placeholder._p not in used:
            placeholder._p.getparent().remove(placeholder._p)
        elif placeholder is not None:                    # заглушку уже заняли — кладём после неё
            after = placeholder
        ref = next((p for p in paragraphs if p._p in used and p._p.getparent() is not None), None)
        placed = rt.add_images_with_captions(
            doc, images, llm=llm if llm_captions else None, context=topic, after=after,
            reference_paragraph=ref, caption_format=caption_format)

    out = BytesIO()
    doc.save(out)
    return ReportResult(out.getvalue(), plan, raw, log, warnings, title_report, placed, dump)
