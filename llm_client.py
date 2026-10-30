# -*- coding: utf-8 -*-
"""
llm_client.py — «LLM» в схеме Parser -> LLM -> Assembler.

* Клиент для ЛЮБОГО OpenAI-совместимого API (POST {base_url}/chat/completions).
  По умолчанию — бесплатный тариф Google Gemini. Переключение на другого
  провайдера — это три переменные: base_url, api_key, model
  (в UI или через переменные окружения LLM_BASE_URL / LLM_API_KEY / LLM_MODEL).
* Реализован на `requests`, без SDK `openai`: на одну зависимость меньше, а повтор
  запросов при лимитах (429) и откат с JSON-режима контролируем сами.
* SYSTEM_PROMPT, сборка пользовательского сообщения и разбор JSON-плана.
"""

from __future__ import annotations

import json
import os
import random
import re
import time
from dataclasses import dataclass
from typing import Optional

import requests

from report_tools import LLMFunc, parse_json_object

# Метка, которую LLM возвращает вместо кода, когда код предоставил студент.
# Сам код в LLM не отправляется (экономия токенов + гарантия «без изменений»).
STUDENT_CODE_SENTINEL = "[[STUDENT_CODE]]"


# --------------------------------------------------------------------------- #
#  Конфигурация и пресеты
# --------------------------------------------------------------------------- #

@dataclass
class LLMConfig:
    base_url: str
    api_key: str
    model: str
    json_mode: bool = True        # response_format={"type": "json_object"}; при ошибке 400 отключится само
    timeout: int = 120
    max_retries: int = 4          # повторы для ПОСЛЕДНЕЙ (или единственной) модели в списке
    fallback_models: tuple = ()   # запасные модели: пробуются, если основная перегружена / недоступна
    backoff_base: float = 3.0     # пауза перед повтором = base * 2**попытка (+ случайная добавка), с


# Название -> (base_url, модель по умолчанию, переменная окружения с ключом).
# Модели Gemini сверены с документацией Google на 29.10.2026; остальные — примеры,
# их имена могли устареть: поле «Модель» в интерфейсе редактируется.
PRESETS = {
    "Google Gemini (бесплатный тариф)": (
        "https://generativelanguage.googleapis.com/v1beta/openai/", "gemini-3.8-flash", "GEMINI_API_KEY"),
    "OpenAI": ("https://api.openai.com/v1", "gpt-4o-mini", "OPENAI_API_KEY"),
    "Groq": ("https://api.groq.com/openai/v1", "llama-3.3-70b-versatile", "GROQ_API_KEY"),
    "OpenRouter": ("https://openrouter.ai/api/v1", "google/gemini-3.8-flash", "OPENROUTER_API_KEY"),
    "Ollama (локально)": ("http://localhost:11434/v1", "qwen2.5:7b", ""),
    "Другой (вручную)": ("", "", "LLM_API_KEY"),
}
DEFAULT_PRESET = "Google Gemini (бесплатный тариф)"

# Запасные модели по умолчанию: у каждой модели своя «ёмкость», и когда одна перегружена (503),
# соседняя обычно отвечает. Меняется в боковой панели или переменной LLM_FALLBACK_MODELS.
DEFAULT_FALLBACKS = {
    "Google Gemini (бесплатный тариф)": ("gemini-3.5-flash-lite", "gemini-3.6-flash"),
}


def parse_models(text: str) -> tuple:
    """«a, b ,c» -> ('a', 'b', 'c')."""
    return tuple(m.strip() for m in (text or "").split(",") if m.strip())


def config_from_env(preset: str = DEFAULT_PRESET) -> LLMConfig:
    """Конфиг из окружения: LLM_* перекрывают пресет; ключ — LLM_API_KEY или переменная пресета."""
    base_url, model, key_env = PRESETS[preset]
    api_key = os.getenv("LLM_API_KEY") or (os.getenv(key_env) if key_env else "") or ""
    fb_env = os.getenv("LLM_FALLBACK_MODELS")
    if fb_env is not None:
        fallbacks = parse_models(fb_env)
    elif os.getenv("LLM_BASE_URL") is None:         # свой адрес API -> имена запасных моделей нам неизвестны
        fallbacks = DEFAULT_FALLBACKS.get(preset, ())
    else:
        fallbacks = ()
    return LLMConfig(
        base_url=os.getenv("LLM_BASE_URL", base_url),
        api_key=api_key,
        model=os.getenv("LLM_MODEL", model),
        fallback_models=fallbacks,
    )


class LLMError(RuntimeError):
    """Ошибка обращения к LLM с понятным текстом для пользователя."""


class _Unavailable(Exception):
    """Модель не ответила (перегрузка, лимит, 404, сеть) — можно пробовать следующую."""

    def __init__(self, status, message):
        super().__init__(message)
        self.status, self.message = status, message


# --------------------------------------------------------------------------- #
#  HTTP-клиент
# --------------------------------------------------------------------------- #

RETRYABLE = (429, 500, 502, 503, 504)


def _short_error(body: str) -> str:
    """Достаёт из тела ошибки читаемое сообщение (у Gemini это [{"error": {"message": ...}}])."""
    try:
        data = json.loads(body)
        if isinstance(data, list) and data:
            data = data[0]
        err = data.get("error", data) if isinstance(data, dict) else {}
        msg = err.get("message") if isinstance(err, dict) else err
        if msg:
            return " ".join(str(msg).split())[:250]
    except (ValueError, AttributeError):
        pass
    return " ".join((body or "").split())[:250]


def _explain(errors: list) -> str:
    """Человеческое объяснение, почему не удалось получить ответ, и что делать."""
    statuses = {exc.status for _, exc in errors}
    if statuses and statuses <= {500, 502, 503, 504}:
        reason = ("Сервис модели перегружен (HTTP 503) — это временный сбой на стороне провайдера, "
                  "а не ваш лимит. Подождите 1–2 минуты и нажмите «Сгенерировать отчет» ещё раз "
                  "или выберите другую модель в боковой панели.")
    elif statuses == {429}:
        reason = ("Исчерпан лимит запросов бесплатного тарифа (HTTP 429). Подождите минуту "
                  "(лимит «запросов в минуту») или до завтра (дневной лимит), либо смените модель.")
    elif statuses == {404}:
        reason = ("Модель не найдена (HTTP 404): возможно, она устарела или недоступна для вашего ключа. "
                  "Проверьте имя в поле «Модель».")
    elif statuses == {0}:
        reason = "Нет связи с сервером API. Проверьте интернет и поле Base URL."
    else:
        reason = "Не удалось получить ответ нейросети."
    tried = "; ".join(f"{model}: {('HTTP ' + str(e.status)) if e.status else 'сеть'} — {e.message}"
                      for model, e in errors)
    return f"{reason}\n\nПробовали: {tried}"


def make_llm(cfg: LLMConfig) -> LLMFunc:
    """
    Возвращает функцию (system, user) -> текст ответа для OpenAI-совместимого API.

    Устойчивость к сбоям: при 429/5xx и сетевых ошибках делаются повторы с нарастающей паузой;
    если модель так и не ответила — пробуется следующая из cfg.fallback_models.
    Атрибуты результата: .last_model и .used_models (какие модели реально отвечали).
    """
    if not cfg.base_url:
        raise LLMError("Не указан base_url LLM.")
    if not cfg.model:
        raise LLMError("Не указана модель LLM.")
    url = cfg.base_url.rstrip("/") + "/chat/completions"
    headers = {"Content-Type": "application/json"}
    if cfg.api_key:                                   # у локальных серверов (Ollama) ключа нет
        headers["Authorization"] = f"Bearer {cfg.api_key}"
    models = [cfg.model] + [m for m in cfg.fallback_models if m and m != cfg.model]

    def pause(attempt: int, retry_after: str = "") -> None:
        try:
            delay = float(retry_after)
        except ValueError:
            delay = cfg.backoff_base * 2 ** attempt
        time.sleep(min(delay, 30) + random.uniform(0, 1) * cfg.backoff_base / 3)

    def ask_model(model: str, system: str, user: str, retries: int) -> str:
        payload = {"model": model,
                   "messages": [{"role": "system", "content": system},
                                {"role": "user", "content": user}]}
        # temperature намеренно не задаём: у новых моделей (в т.ч. Gemini 3) разумное
        # значение по умолчанию, а часть провайдеров не принимает нестандартное.
        if cfg.json_mode:
            payload["response_format"] = {"type": "json_object"}

        last_status, last_msg = None, ""
        for attempt in range(retries + 1):
            try:
                resp = requests.post(url, headers=headers, json=payload, timeout=cfg.timeout)
                if resp.status_code == 400 and "response_format" in payload:
                    payload.pop("response_format")    # провайдер не знает JSON-режим — повторяем без него
                    resp = requests.post(url, headers=headers, json=payload, timeout=cfg.timeout)
            except requests.RequestException as exc:
                last_status, last_msg = 0, f"сетевая ошибка: {exc}"
                if attempt < retries:
                    pause(attempt)
                continue

            code = resp.status_code
            if code in RETRYABLE:
                last_status, last_msg = code, _short_error(resp.text)
                if attempt < retries:
                    pause(attempt, resp.headers.get("Retry-After", ""))
                continue
            if code in (401, 403):
                raise LLMError("API отклонил ключ (401/403). Проверьте ключ и base_url.")
            if code == 404:
                raise _Unavailable(404, _short_error(resp.text))
            if code >= 400:
                raise LLMError(f"HTTP {code}: {_short_error(resp.text)}")
            try:
                return resp.json()["choices"][0]["message"]["content"] or ""
            except (KeyError, IndexError, ValueError) as exc:
                raise LLMError(f"Неожиданный формат ответа API: {resp.text[:300]}") from exc
        raise _Unavailable(last_status, last_msg)

    start = {"i": 0}      # с какой модели начинать: после сбоя основной не тратим попытки на неё снова

    def call(system: str, user: str) -> str:
        errors = []
        order = models[start["i"]:] + models[:start["i"]]
        for pos, model in enumerate(order):
            is_last = pos == len(order) - 1
            # у не последней модели повторов мало: лучше быстро перейти на соседнюю, чем долго ждать
            retries = cfg.max_retries if is_last else min(2, cfg.max_retries)
            try:
                text = ask_model(model, system, user, retries)
            except _Unavailable as exc:
                errors.append((model, exc))
                continue
            start["i"] = models.index(model)           # следующие запросы этого запуска — сразу к рабочей модели
            call.last_model = model
            if model not in call.used_models:
                call.used_models.append(model)
            return text
        raise LLMError(_explain(errors))

    call.last_model = None
    call.used_models = []
    return call


# --------------------------------------------------------------------------- #
#  Системный промпт
# --------------------------------------------------------------------------- #

SYSTEM_PROMPT = f"""Ты — ассистент, который составляет план заполнения отчёта по лабораторной работе в российском вузе. Ты НЕ редактируешь документ сам: ты возвращаешь JSON-план, по которому программа вставит тексты в готовый шаблон Word.

# ВХОДНЫЕ ДАННЫЕ
В сообщении пользователя будут блоки:

<template_paragraphs> — абзацы шаблона, по одному на строку в формате:
  [индекс] ФЛАГИ | текст
Флаги: EMPTY (пустой абзац), TEXT (есть текст), BOLD (весь текст жирный — как правило, заголовок раздела), HEADING_STYLE (стиль «Заголовок»), PAGE_BREAK (разрыв страницы), IMAGE (в абзаце картинка).
Индекс — число в квадратных скобках. Используй ТОЛЬКО эти числа; строки сам не пересчитывай.

<methodical_text> — текст методички: цель, задание, требования, иногда несколько вариантов.
<user_params> — параметры от студента: номер и тема работы, ВАРИАНТ и РЕЖИМ КОДА.
<screenshots> — (может отсутствовать) имена файлов скриншотов; это подсказка о ходе работы.

# ЗАДАЧА
Сопоставь по СМЫСЛУ разделы методички и заголовки шаблона и определи, в какой абзац шаблона вставить каждый из четырёх текстов. Названия в шаблоне и методичке часто различаются, ориентируйся на смысл:
- ЦЕЛЬ — заголовки «Цель», «Цель работы», «Цели».
- ЗАДАНИЕ — «Задание», «Задачи», «Учебные задачи», «Постановка задачи», «Индивидуальное задание».
- КОД — «Листинг», «Листинг программы», «Исходный код», «Код программы», «Программная реализация».
- ВЫВОД — «Вывод», «Выводы», «Заключение».
Разделы «Теоретические сведения», «Ход работы», «Результаты», «Порядок выполнения» и т.п. НЕ заполняй: скриншоты и титульный лист программа обрабатывает сама.
Если заголовок объединяет разделы («Цель и задачи»), положи в один абзац оба текста: «Цель: …\\nЗадачи: …».

# КУДА ВСТАВЛЯТЬ (выбор индекса)
Целевой абзац — первый абзац ПОСЛЕ нужного заголовка, который имеет флаг EMPTY либо содержит заглушку: «…», «___», «[вставьте текст]», «{{{{ЦЕЛЬ}}}}», «Текст», «Здесь должен быть …» и т.п.
Запрещено выбирать: сам заголовок; абзац из другого раздела; абзац с настоящим содержимым (его нельзя затирать); абзацы титульного листа (всё до первого разрыва страницы).
Если подходящего абзаца под заголовком нет, не угадывай: пропусти раздел и опиши причину в "_warnings".

# ЕСЛИ МЕТОДИЧКИ НЕТ
Если вместо текста методички написано, что она не предоставлена, составь цель, задание, код и вывод по ТЕМЕ работы из <user_params>: сформулируй типовые задачи для такой лабораторной, не выдумывая конкретных данных, чисел и результатов, и добавь в "_warnings" предупреждение, что материалы составлены без методички.

# ВАРИАНТ
Если в <user_params> указан вариант, сгенерируй содержимое именно под него: найди в методичке условие этого варианта и используй его исходные данные в задании, коде и выводе. Если вариант указан, а в методичке вариантов нет, — игнорируй его и добавь предупреждение в "_warnings". Если вариант не указан, а в методичке их несколько, — возьми общую часть и добавь предупреждение в "_warnings".

# СОДЕРЖАНИЕ ТЕКСТОВ
Язык — русский, научно-деловой стиль, безличные обороты («в ходе работы были изучены…»).
- Цель: 1–3 предложения. Опирайся на цель из методички, формулируй своими словами; если цели в методичке нет — выведи её из задания, а не придумывай.
- Задание: перечень задач из методички (для выбранного варианта), близко к тексту, нумерованный список — каждый пункт с новой строки в виде «1. …».
- Код — зависит от режима в <user_params>:
  • «Режим кода: USER» — код предоставил студент. В абзац для кода верни ровно строку {STUDENT_CODE_SENTINEL} (ничего больше), сам код НЕ пиши.
  • «Режим кода: GENERATE» — напиши код сам: полное работающее решение на языке из методички (или из тем работы); комментарии на русском; отступы пробелами (4 на уровень); только код, без markdown-ограждений (```) и пояснений вне комментариев. Если программной части нет, но есть последовательность команд (консоль, администрирование) — вставь команды. Если нет ни того ни другого — не добавляй ключ.
- Вывод: 3–6 предложений в прошедшем времени: что сделано, чему научились, достигнута ли цель. Связывай с целью и задачами. Не выдумывай конкретные числа, замеры и результаты, которых нет во входных данных.
Никакого markdown (**, #, `) — текст попадёт в Word как есть. Абзацы разделяй символом \\n. Не копируй большие куски методички дословно (кроме списка заданий).

# ФОРМАТ ОТВЕТА
Верни ТОЛЬКО один валидный JSON-объект: без пояснений до и после, без ```json. Ключ — "paragraph_<индекс>", значение — строка. Переводы строк в значениях записывай как \\n, кавычки внутри строк экранируй как \\", обратные слеши как \\\\.
Дополнительно можно добавить ключ "_warnings" со списком строк (только если есть что сообщить).
Нет данных для раздела — не включай его. Один индекс — не более одного раза.

# ПРИМЕР (режим кода GENERATE, вариант не указан)
Вход:
<template_paragraphs>
[4] TEXT,BOLD | Лабораторная работа №3
[8] TEXT,BOLD,PAGE_BREAK | 1. Цель работы
[9] EMPTY |
[10] TEXT,BOLD | 2. Задание
[11] TEXT | [вставьте текст]
[12] TEXT,BOLD | 3. Листинг программы
[13] EMPTY |
[14] TEXT,BOLD | 4. Выводы
[15] EMPTY |
</template_paragraphs>
Выход:
{{"paragraph_9": "Освоить работу с массивами в языке Python.", "paragraph_11": "1. Создать массив из 10 чисел.\\n2. Найти максимальный элемент.", "paragraph_13": "numbers = [3, 7, 1]\\nprint(max(numbers))", "paragraph_15": "В ходе работы были освоены основные приёмы работы с массивами. Цель работы достигнута.", "_warnings": ["Вариант не указан, использована общая часть методички."]}}"""


# --------------------------------------------------------------------------- #
#  Сообщение пользователя и разбор плана
# --------------------------------------------------------------------------- #

def build_user_message(template_dump: str, method_text: str, *, number: str = "", topic: str = "",
                       variant: str = "", has_student_code: bool = False,
                       screenshot_names: Optional[list] = None) -> str:
    """Собирает пользовательское сообщение по формату, описанному в SYSTEM_PROMPT."""
    params = []
    if number:
        params.append(f"Номер работы: {number}")
    if topic:
        params.append(f"Тема работы: {topic}")
    if variant:
        params.append(f"Вариант: {variant}. Сгенерируй задание, код и вывод именно под этот вариант: "
                      f"найди в методичке условие варианта {variant} и используй его данные.")
    else:
        params.append("Вариант: не указан")
    if has_student_code:
        params.append(f"Режим кода: USER — код предоставит студент; в абзац для кода верни ровно "
                      f"строку {STUDENT_CODE_SENTINEL}.")
    else:
        params.append("Режим кода: GENERATE — напиши код сам.")

    parts = [
        "<template_paragraphs>\n" + template_dump + "\n</template_paragraphs>",
        "<methodical_text>\n" + method_text + "\n</methodical_text>",
        "<user_params>\n" + "\n".join(params) + "\n</user_params>",
    ]
    if screenshot_names:
        parts.append("<screenshots>\n" + "\n".join(screenshot_names) + "\n</screenshots>")
    return "\n\n".join(parts)


def parse_plan(data: dict) -> tuple:
    """JSON от LLM -> ({индекс: текст}, [предупреждения]). Лишние/битые ключи игнорируются."""
    plan, warnings = {}, []
    for key, value in data.items():
        m = re.fullmatch(r"paragraph_(\d+)", str(key).strip())
        if not m:
            continue
        if isinstance(value, list):
            value = "\n".join(str(v) for v in value)
        if isinstance(value, str) and value.strip():
            plan[int(m.group(1))] = value
    raw_warn = data.get("_warnings", [])
    if isinstance(raw_warn, str):
        raw_warn = [raw_warn]
    if isinstance(raw_warn, list):
        warnings = [str(w) for w in raw_warn if str(w).strip()]
    return plan, warnings


def request_plan(llm: LLMFunc, user_message: str, attempts: int = 2) -> tuple:
    """Запрашивает план у LLM. Если ответ — не JSON-план, один раз просит исправиться.
    Возвращает (plan, warnings, raw_text)."""
    message, raw = user_message, ""
    for attempt in range(attempts):
        raw = llm(SYSTEM_PROMPT, message)
        data = parse_json_object(raw)
        plan, warnings = parse_plan(data)
        if plan:
            return plan, warnings, raw
        message = (user_message + "\n\nПредыдущий ответ не был корректным JSON-планом. "
                   "Верни ТОЛЬКО JSON-объект с ключами paragraph_<индекс>.")
    return {}, [], raw