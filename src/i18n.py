"""Локализация интерфейса."""

TRANSLATIONS = {
    'ru': {
        'title': "Форматировщик перекрестных ссылок",
        'app_name': "Форматировщик перекрестных ссылок (ФПС)",
        'subtitle': "Выберите .docx – программа расставит поля SEQ/REF и сохранит копию «_автонумерация».",
        'file_frame': "Файл",
        'no_file': "– файл не выбран –",
        'btn_select': "Выбрать…",
        'btn_process': "Обработать и сохранить",
        'btn_check': "Только проверить (без записи)",
        'btn_open': "Открыть папку",
        'btn_exit': "Выход",
        'log_frame': "Журнал",
        'hint_f9': "После обработки откройте файл в Word и нажмите Ctrl+A → F9, чтобы Word пересчитал все поля.",
        'hint_ref': "Готовые перекрестные ссылки доступны в Word: Ссылки → Перекрестная ссылка → Тип ссылки → Закладка.",
        'lang_label': "Язык интерфейса:",
        'doc_lang_label': "Язык документа:",
        'log_start': "{name}: обработка «{path}»",
        'log_check_mode': "Режим ПРОВЕРКИ: файл не будет изменен.",
        'log_result': "Результат: {path}",
        'log_paragraphs': "Абзацев всего (с таблицами): {count}",
        'log_step1a': "Шаг 1а/3: формулы в таблицах…",
        'log_step1b': "Шаг 1б/3: подписи и формулы в тексте…",
        'log_step2': "Шаг 2/3: список литературы…",
        'log_step3': "Шаг 3/3: замена ссылок в тексте…",
        'log_found_stats': "  рисунков={figs}, таблиц={tabs}, формул={eqs}",
        'log_total_refs': "  всего ссылок сделано: {refs}",
        'log_saved': "Сохранено: {path}",
        'log_not_saved': "Режим проверки: файл НЕ сохранен.",
        'log_done': "Готово.",
        'log_file_selected': "Выбран файл: {path}",
        'log_error': "ОШИБКА: {exc}",
        'log_bib_found': "  → раздел литературы: «{text}»",
        'log_bib_end': "  ← конец раздела: «{text}»",
        'log_bib_warn': "    [WARN] не распознан: «{text}»",
        'log_bib_item': "    [source {bm_name}]",
        'log_bib_total': "  → источников: {count}",
        'msg_wrong_format_title': "Неверный формат",
        'msg_wrong_format_text': "Работаем только с .docx.",
        'msg_success': "Готово",
        'msg_check': "Проверка",
        'msg_done_path': "Готово: {path}",
        'msg_check_done': "Проверка завершена (файл не изменен).",
        'msg_stats': (
            "Рисунков:   {figures}\n"
            "Таблиц:     {tables}\n"
            "Формул:     {equations}\n"
            "Источников: {sources}\n"
            "Ссылок:     {refs}\n"
            "Абзацев:    {paragraphs}"
        ),
        'msg_error_title': "Ошибка",
        'msg_error_word': "\n\nУбедитесь, что файл не открыт в Word.",
        'msg_folder_err': "Не удалось открыть папку: {e}"
    },
    'en': {
        'title': "Cross-Reference Formatter",
        'app_name': "Cross-Reference Formatter (CRF)",
        'subtitle': "Select a .docx file – the program will insert SEQ/REF fields and save a copy with the suffix '_autonumbering'.",
        'file_frame': "File",
        'no_file': "– no file selected –",
        'btn_select': "Browse...",
        'btn_process': "Process & Save",
        'btn_check': "Check Only (No Save)",
        'btn_open': "Open Folder",
        'btn_exit': "Exit",
        'log_frame': "Log",
        'hint_f9': "After processing, open the file in Word and press Ctrl+A → F9 to update all fields.",
        'hint_ref': "Cross-references are available in Word: References → Cross-reference → Reference type → Bookmark.",
        'lang_label': "UI Language:",
        'doc_lang_label': "Document Language:",
        'log_start': "{name}: processing «{path}»",
        'log_check_mode': "CHECK mode: file will not be modified.",
        'log_result': "Result: {path}",
        'log_paragraphs': "Total paragraphs (with tables): {count}",
        'log_step1a': "Step 1a/3: formulas in tables...",
        'log_step1b': "Step 1b/3: captions and formulas in text...",
        'log_step2': "Step 2/3: bibliography...",
        'log_step3': "Step 3/3: replacing references in text...",
        'log_found_stats': "  figures={figs}, tables={tabs}, equations={eqs}",
        'log_total_refs': "  total references made: {refs}",
        'log_saved': "Saved: {path}",
        'log_not_saved': "Check mode: file NOT saved.",
        'log_done': "Done.",
        'log_file_selected': "File selected: {path}",
        'log_error': "ERROR: {exc}",
        'log_bib_found': "  → bibliography section: «{text}»",
        'log_bib_end': "  ← end of section: «{text}»",
        'log_bib_warn': "    [WARN] not recognized: «{text}»",
        'log_bib_item': "    [source {bm_name}]",
        'log_bib_total': "  → sources: {count}",
        'msg_wrong_format_title': "Wrong Format",
        'msg_wrong_format_text': "Only .docx files are supported.",
        'msg_success': "Success",
        'msg_check': "Check",
        'msg_done_path': "Done: {path}",
        'msg_check_done': "Check completed (file not modified).",
        'msg_stats': (
            "Figures:    {figures}\n"
            "Tables:     {tables}\n"
            "Equations:  {equations}\n"
            "Sources:    {sources}\n"
            "References: {refs}\n"
            "Paragraphs: {paragraphs}"
        ),
        'msg_error_title': "Error",
        'msg_error_word': "\n\nMake sure the file is not open in Word.",
        'msg_folder_err': "Failed to open folder: {e}"
    }
}

current_ui_lang = 'ru'
current_doc_lang = 'ru'


def tr(key, **kwargs):
    text = TRANSLATIONS[current_ui_lang].get(key, key)
    if kwargs:
        return text.format(**kwargs)
    return text