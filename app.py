# -*- coding: utf-8 -*-
"""
ИИ-генератор отчётов по лабораторным работам — интерфейс Streamlit.

Цепочка Parser -> LLM -> Assembler живёт в соседних модулях:
    report_tools.py — чтение методички, структура шаблона, титульник, скриншоты
    llm_client.py   — клиент любого OpenAI-совместимого API (по умолчанию Gemini)
    assembler.py    — generate_report(): единая точка входа для этого файла

Запуск:
    pip install -r requirements.txt
    streamlit run app.py
Ключ можно ввести в боковой панели или задать переменной окружения
GEMINI_API_KEY / LLM_API_KEY.
"""

import traceback

import streamlit as st

from assembler import generate_report
from llm_client import DEFAULT_PRESET, PRESETS, LLMConfig, LLMError, config_from_env, make_llm
from report_tools import extract_methodichka_text

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

# Что получает нейросеть вместо методички, если пользователь её не загрузил
NO_METHOD_TEXT = ("Методичка не предоставлена. Составь материалы отчёта по теме работы "
                  "из <user_params>.")


# --------------------------------------------------------------------------- #
#  Боковая панель: LLM (ключ, модель, провайдер) и оформление
# --------------------------------------------------------------------------- #

def sidebar_settings() -> dict:
    """Рисует боковую панель и возвращает настройки: конфиг LLM и параметры оформления."""
    st.sidebar.header("⚙️ Нейросеть")
    presets = list(PRESETS)
    preset = st.sidebar.selectbox("Провайдер (OpenAI-совместимый API)", presets,
                                  index=presets.index(DEFAULT_PRESET))
    env_cfg = config_from_env(preset)            # значения по умолчанию: пресет + переменные окружения
    # key=… с именем пресета: у каждого провайдера свои запомненные значения полей
    api_key = st.sidebar.text_input(
        "API-ключ", value=env_cfg.api_key, type="password", key=f"api_key_{preset}",
        help="Можно не вводить, если ключ задан в окружении (GEMINI_API_KEY / LLM_API_KEY).")
    model = st.sidebar.text_input("Модель", value=env_cfg.model, key=f"model_{preset}")
    base_url = st.sidebar.text_input("Base URL", value=env_cfg.base_url, key=f"base_url_{preset}")
    if preset == DEFAULT_PRESET:
        st.sidebar.caption(
            "Бесплатный ключ: https://aistudio.google.com/apikey. На бесплатном тарифе Google, "
            "по его условиям, может использовать запросы для улучшения своих продуктов — "
            "не отправляйте конфиденциальные данные.")

    with st.sidebar.expander("🎨 Оформление"):
        code_size = st.number_input("Размер шрифта кода, пт", 6, 16, 10)
        caption_format = st.text_input("Формат подписи под рисунком", value="Рисунок {n} — {title}",
                                       help="{n} — номер, {title} — название.")
        llm_captions = st.checkbox("Названия рисунков придумывает нейросеть", value=True,
                                   help="Выключено — название берётся из имени файла без LLM.")
        sort_images = st.checkbox("Сортировать скриншоты по имени файла", value=True)

    return {
        "preset": preset,
        "cfg": LLMConfig(base_url=base_url.strip(), api_key=api_key.strip(), model=model.strip()),
        "code_size": float(code_size),
        "caption_format": caption_format,
        "llm_captions": llm_captions,
        "sort_images": sort_images,
    }


# --------------------------------------------------------------------------- #
#  Запуск генерации
# --------------------------------------------------------------------------- #

def run_generation(settings: dict, template_file, method_file, topic: str, number: str,
                   variant: str, screenshots: list, student_code: str) -> None:
    """Выполняет шаги 1–5 из ТЗ. Успех кладёт результат в st.session_state, ошибка — в st.error."""
    # Прошлый результат убираем сразу: при ошибке не должно остаться «чужой» кнопки скачивания
    for key in ("docx_bytes", "warnings", "file_name", "details"):
        st.session_state.pop(key, None)

    # --- проверки ввода ---
    if template_file is None:
        st.error("Загрузите шаблон отчёта (.docx).")
        return
    if method_file is None and not topic:
        st.error("Загрузите методичку или укажите тему работы — иначе нейросети не из чего составлять отчёт.")
        return
    cfg = settings["cfg"]
    if not cfg.api_key and settings["preset"] != "Ollama (локально)":
        st.error("Не задан API-ключ: введите его в боковой панели или задайте переменную окружения.")
        return

    try:
        # 1. Методичка (если загружена)
        method_text = NO_METHOD_TEXT
        if method_file is not None:
            with st.spinner("Читаю методичку…"):
                method_text, truncated = extract_methodichka_text(method_file.name, method_file.getvalue())
            if truncated:
                st.warning("Методичка очень длинная — в нейросеть отправлено только начало (≈40 000 символов).")

        # 2. LLM-клиент: ключ, модель и адрес API берём из настроек
        llm = make_llm(cfg)

        # 3. Генерация отчёта
        with st.spinner("Нейросеть составляет план и собирает документ… (до минуты)"):
            result = generate_report(
                template_file.getvalue(),
                method_text,
                [(f.name, f.getvalue()) for f in screenshots or []],
                llm=llm,
                number=number,
                topic=topic,
                variant=variant,
                student_code=student_code,
                code_size_pt=settings["code_size"],
                caption_format=settings["caption_format"],
                llm_captions=settings["llm_captions"],
                sort_images=settings["sort_images"],
            )

        # 4. Успех: сохраняем файл в session_state (кнопка скачивания рисуется ниже, вне этой функции,
        #    поэтому переживает перезапуск скрипта после клика по ней)
        st.session_state["docx_bytes"] = result.docx_bytes
        st.session_state["warnings"] = result.warnings
        st.session_state["file_name"] = f"report_lab{number}.docx" if number else "report.docx"
        st.session_state["details"] = {
            "log": result.log, "images": result.images,
            "raw": result.raw_llm_answer, "dump": result.template_dump,
        }

    # 5. Ошибки
    except (LLMError, ValueError) as exc:           # ожидаемые: ключ, квота, битый файл, нет текста в PDF
        st.error(str(exc))
    except Exception as exc:                         # непредвиденные: показываем причину и трассировку
        st.error(f"Непредвиденная ошибка: {exc}")
        with st.expander("Подробности"):
            st.code(traceback.format_exc())


# --------------------------------------------------------------------------- #
#  Основной экран
# --------------------------------------------------------------------------- #

def main() -> None:
    st.set_page_config(page_title="ИИ-генератор отчётов", page_icon="📄", layout="wide")
    st.title("📄 ИИ-генератор отчётов по лабораторным работам")
    st.caption("Шаблон + методичка + скриншоты → готовый .docx в оформлении вашего шаблона.")
    settings = sidebar_settings()

    # ---- Файлы ----
    st.subheader("1. Файлы")
    c1, c2 = st.columns(2)
    template_file = c1.file_uploader("Шаблон отчёта (.docx) — с вашим оформлением и титульным листом",
                                     type=["docx"])
    method_file = c2.file_uploader("Методичка (.pdf или .docx) — необязательно", type=["pdf", "docx"])
    screenshots = st.file_uploader(
        "Скриншоты (в названиях файлов — подсказки, что на них, например «02_ping_проверка.png»)",
        type=["png", "jpg", "jpeg", "bmp", "gif", "webp"], accept_multiple_files=True)

    # ---- Параметры ----
    st.subheader("2. Параметры работы")
    p1, p2, p3 = st.columns([1, 3, 1])
    number = p1.text_input("Номер работы", placeholder="6",
                           help="Необязательно. Подставляется на титульный лист.").strip()
    topic = p2.text_input("Тема работы", placeholder="Установка ОС Astra Linux в виртуальной среде",
                          help="Обязательна, если методичка не загружена. Подставляется на титульный лист.").strip()
    variant = p3.text_input("Вариант", placeholder="12",
                            help="Необязательно. Нейросеть найдёт в методичке условие этого варианта "
                                 "и составит задание, код и вывод под него.").strip()
    if method_file is None and not topic:
        st.caption("ℹ️ Методичка не загружена — укажите тему работы.")

    student_code = st.text_area(
        "Код студента (необязательно)", height=220,
        placeholder="Вставьте свой код — он попадёт в отчёт БЕЗ изменений. "
                    "Оставьте пустым, и нейросеть напишет код сама.")
    if student_code.strip():
        st.info("Режим кода: **ваш код** — вставляется без изменений, нейросеть его не видит.")
    else:
        st.info("Режим кода: **генерация** — нейросеть напишет код"
                + (f" для варианта {variant}." if variant else "."))

    # ---- Кнопка ----
    st.divider()
    if st.button("🚀 Сгенерировать отчет", type="primary"):
        run_generation(settings, template_file, method_file, topic, number, variant,
                       screenshots, student_code)

    # ---- Результат ----
    if "docx_bytes" in st.session_state:
        st.success("Готово! Отчёт собран в оформлении вашего шаблона.")
        st.download_button("⬇️ Скачать отчёт (.docx)", data=st.session_state["docx_bytes"],
                           file_name=st.session_state.get("file_name", "report.docx"),
                           mime=DOCX_MIME, type="primary")
        for warning in st.session_state.get("warnings", []):
            st.warning(warning)

        details = st.session_state.get("details", {})
        with st.expander("Что нейросеть вставила и куда"):
            st.table([{"Абзац №": e["index"], "Раздел шаблона": e["section"] or "—",
                       "Тип": "код" if e["kind"] == "code" else "текст",
                       "Статус": "✅ вставлено" if e["status"] == "ok" else "⏭ пропущено",
                       "Примечание": e["note"]} for e in details.get("log", [])]
                     or [{"Результат": "план пуст"}])
            if details.get("images"):
                st.caption("Подписи к рисункам:")
                st.table([{"Файл": n, "Подпись": c} for n, c in details["images"]])
        with st.expander("Отладка: ответ нейросети и структура шаблона"):
            st.code(details.get("raw") or "—", language="json")
            st.code(details.get("dump") or "—", language="text")


if __name__ == "__main__":
    main()
