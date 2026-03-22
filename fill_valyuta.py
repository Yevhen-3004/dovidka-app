"""
fill_valyuta.py — заповнює шаблон заяви на купівлю валюти (docx → pdf)

Використовується з app.py через subprocess або імпортом.
"""
import re
import shutil
import subprocess
import sys
import os
import tempfile
from datetime import date
from pathlib import Path

# Місячні назви для дати
MONTHS_UK = {
    1: "січня", 2: "лютого", 3: "березня", 4: "квітня",
    5: "травня", 6: "червня", 7: "липня", 8: "серпня",
    9: "вересня", 10: "жовтня", 11: "листопада", 12: "грудня",
}

CURRENCY_CODES = {"USD": "840", "EUR": "978"}
CURRENCY_NAMES = {"USD": "Долар США", "EUR": "Євро"}


def get_date_parts(date_str: str):
    """Повертає (день, місяць_укр, рік) з рядка ДД.ММ.РРРР або сьогодні."""
    if date_str and date_str.strip():
        parts = date_str.strip().split(".")
        day, month, year = int(parts[0]), int(parts[1]), int(parts[2])
    else:
        today = date.today()
        day, month, year = today.day, today.month, today.year
    return str(day), MONTHS_UK[month], str(year)


def replace_nth_match(content: str, pattern: str, replacement: str, n: int) -> str:
    """Замінює n-е (0-базоване) входження pattern на replacement."""
    matches = list(re.finditer(pattern, content))
    if n >= len(matches):
        return content
    m = matches[n]
    return content[:m.start()] + replacement + content[m.end():]


def replace_text_in_cell(content: str, old_text: str, new_text: str) -> str:
    """Замінює текст у <w:t> тегах."""
    # Шукаємо точне входження в w:t
    pattern = r'(<w:t[^>]*>)' + re.escape(old_text) + r'(</w:t>)'
    return re.sub(pattern, r'\g<1>' + new_text + r'\g<2>', content)


def replace_individual_digits(content: str, start_marker_text: str,
                               old_digits: list, new_digits: list,
                               occurrence: int = 0) -> str:
    """
    Замінює окремі цифри у послідовних <w:t> клітинках після маркера.
    old_digits і new_digits — списки рядків (один символ кожен).
    """
    # Знаходимо всі w:t з їх позиціями
    wt_pattern = re.compile(r'<w:t(?:[^>]*)>([^<]*)</w:t>')
    all_matches = list(wt_pattern.finditer(content))

    # Знаходимо маркер (наприклад 'IBAN', 'купівлі')
    marker_indices = [i for i, m in enumerate(all_matches)
                      if start_marker_text in m.group(1)]

    if occurrence >= len(marker_indices):
        return content

    start_idx = marker_indices[occurrence] + 1  # наступний після маркера

    # Збираємо позиції для заміни
    replacements = []
    old_pos = 0
    new_pos = 0

    i = start_idx
    while old_pos < len(old_digits) and i < len(all_matches):
        cell_text = all_matches[i].group(1)
        if cell_text == old_digits[old_pos]:
            if new_pos < len(new_digits):
                replacements.append((all_matches[i].start(), all_matches[i].end(),
                                     all_matches[i].group(0).replace(
                                         f'>{old_digits[old_pos]}<',
                                         f'>{new_digits[new_pos]}<')))
            old_pos += 1
            new_pos += 1
        i += 1

    # Застосовуємо заміни у зворотньому порядку
    for start, end, new_val in reversed(replacements):
        content = content[:start] + new_val + content[end:]

    return content


def iban_to_digits(iban: str) -> list:
    """Перетворює IBAN рядок у список символів (без пробілів)."""
    return list(iban.replace(" ", "").replace("-", ""))


def amount_to_digits(amount_str: str) -> list:
    """
    Перетворює суму у список символів для клітинок.
    Формат у документі: 5 8 0 0 , 0 0  → ['5','8','0','0',',','0','0']
    """
    # Форматуємо як ціле + дробова частина
    try:
        val = float(amount_str)
        int_part = str(int(val))
        dec_part = f"{val:.2f}".split('.')[1]
        return list(int_part) + [','] + list(dec_part)
    except Exception:
        return list(amount_str)


def currency_code_to_digits(code: str) -> list:
    return list(code)


def fill_template(params: dict, template_path: str, output_docx: str):
    """
    Заповнює docx-шаблон даними з params і зберігає у output_docx.

    params:
        pib         : str  — "Прізвище Ім'я По-батькові"
        ipn         : str  — 10 цифр
        address     : str  — повна адреса
        date        : str  — "ДД.ММ.РРРР" або "" (сьогодні)
        purpose_en  : str  — мета купівлі англійською
        purpose_uk  : str  — мета купівлі українською (генерується автоматично)
        amount      : str  — сума, напр. "5800.00"
        currency    : str  — "USD" або "EUR"
        iban_debit  : str  — IBAN для купівлі (гривневий)
        iban_credit : str  — IBAN для зарахування валюти
        iban_commission: str — IBAN для списання комісії
    """
    # Розпакуємо шаблон у тимчасову папку
    tmpdir = tempfile.mkdtemp()
    unpack_dir = os.path.join(tmpdir, "unpacked")
    scripts_dir = "/mnt/skills/public/docx/scripts/office"

    subprocess.run(
        [sys.executable, f"{scripts_dir}/unpack.py",
         template_path, unpack_dir],
        check=True, capture_output=True
    )

    doc_path = os.path.join(unpack_dir, "word", "document.xml")
    content = open(doc_path, encoding="utf-8").read()

    pib = params["pib"]
    pib_parts = pib.split()
    ipn = params["ipn"]
    address = params["address"]
    day, month_uk, year = get_date_parts(params.get("date", ""))
    purpose_uk = params["purpose_uk"]
    amount = params["amount"]
    currency = params["currency"]
    cur_code = CURRENCY_CODES.get(currency, "840")
    cur_name = CURRENCY_NAMES.get(currency, "Долар США")
    iban_debit = params["iban_debit"].replace(" ", "")
    iban_credit = params["iban_credit"].replace(" ", "")
    iban_commission = params["iban_commission"].replace(" ", "")

    # Поточні значення у шаблоні
    OLD_PIB_PARTS = ["Каллаш", "Леонід", "Юрійович"]
    OLD_IPN = "3073810850"
    OLD_ADDRESS_1 = "Україна, м. Київ, пров. Костя Гордієнка, буд."
    OLD_ADDRESS_2 = "10, кв. 7"
    OLD_DAY = "23"
    OLD_MONTH = "березня"
    OLD_YEAR = "2026"
    OLD_PURPOSE_UK = ["Передплата", " ", "за", " ", "товари", " ", "згідно", " ",
                      "контракту", " ", "01/26", " ", "від", " ", "16.03.2026,",
                      " ", "інвойсу", " ", "PI2026031201від", " ", "17.03.2026"]
    OLD_AMOUNT_DIGITS = ["5", "8", "0", "0", ",", "0", "0"]
    OLD_CUR_CODE_DIGITS = ["8", "4", "0"]
    OLD_CUR_NAME = ["Долар", "США"]
    OLD_IBAN1 = list("UA293220010000026006310115156")   # гривневий (купівля)
    OLD_IBAN2 = list("UA353220010000026000370076190")   # валютний
    OLD_IBAN3 = list("UA293220010000026006310115156")   # комісія (той самий що 1)

    # ── 1. ДАТА ──────────────────────────────────────────────────────────────
    content = replace_text_in_cell(content, OLD_DAY, day)
    content = replace_text_in_cell(content, OLD_MONTH, month_uk)
    content = replace_text_in_cell(content, OLD_YEAR, year)

    # ── 2. ПІБ (вгорі — 3 окремі runs) ──────────────────────────────────────
    # Замінюємо кожен рун окремо (Прізвище / Ім'я / По-батькові)
    for i, old_part in enumerate(OLD_PIB_PARTS):
        new_part = pib_parts[i] if i < len(pib_parts) else ""
        content = replace_text_in_cell(content, old_part, new_part)

    # ── 3. АДРЕСА ─────────────────────────────────────────────────────────────
    # Адреса зберігається у двох run-ах, об'єднаємо все в перший
    content = replace_text_in_cell(content, OLD_ADDRESS_1, address)
    content = replace_text_in_cell(content, OLD_ADDRESS_2, " ")

    # ── 4. ІПН ────────────────────────────────────────────────────────────────
    content = replace_text_in_cell(content, OLD_IPN, ipn)

    # ── 5. МЕТА КУПІВЛІ (українська) ─────────────────────────────────────────
    # Мета розбита на окремі runs — замінюємо весь блок першим run-ом
    # Знаходимо "Передплата" і замінюємо на повний текст, решту — очищаємо
    wt_pattern = re.compile(r'(<w:t(?:[^>]*)>)([^<]*)(</w:t>)')
    all_wt = list(wt_pattern.finditer(content))

    # Знаходимо індекс "Передплата" run
    purpose_start = None
    for i, m in enumerate(all_wt):
        if m.group(2) == "Передплата":
            purpose_start = i
            break

    if purpose_start is not None:
        # Замінюємо перший run на повний текст мети
        m0 = all_wt[purpose_start]
        content = content[:m0.start()] + m0.group(1) + purpose_uk + m0.group(3) + content[m0.end():]

        # Перераховуємо після заміни
        all_wt = list(wt_pattern.finditer(content))
        for i, m in enumerate(all_wt):
            if m.group(2) == purpose_uk:
                purpose_start = i
                break

        # Очищаємо наступні runs до "(Оплата/передплата"
        all_wt = list(wt_pattern.finditer(content))
        i = purpose_start + 1
        replacements = []
        while i < len(all_wt):
            txt = all_wt[i].group(2)
            if "(Оплата/передплата" in txt:
                break
            if txt.strip():
                replacements.append((all_wt[i].start(), all_wt[i].end(),
                                     all_wt[i].group(1) + "" + all_wt[i].group(3)))
            i += 1
        for start, end, new_val in reversed(replacements):
            content = content[:start] + new_val + content[end:]

    # ── 6. СУМА ───────────────────────────────────────────────────────────────
    new_amount_digits = amount_to_digits(amount)
    content = replace_individual_digits(content, "купівлі", OLD_AMOUNT_DIGITS,
                                        new_amount_digits, occurrence=1)

    # ── 7. КОД ВАЛЮТИ ─────────────────────────────────────────────────────────
    new_cur_code_digits = list(cur_code)
    content = replace_individual_digits(content, "Currency", OLD_CUR_CODE_DIGITS,
                                        new_cur_code_digits, occurrence=0)

    # ── 8. НАЗВА ВАЛЮТИ ───────────────────────────────────────────────────────
    # "Долар" + " " + "США"  →  перший рун отримує нову назву
    if currency == "EUR":
        content = replace_text_in_cell(content, "Долар", "Євро")
        content = replace_text_in_cell(content, "США", "")
    # USD залишається як є

    # ── 9. IBAN №1 (гривневий — списання для купівлі) ─────────────────────────
    new_iban1 = iban_to_digits(iban_debit)
    content = replace_individual_digits(content, "IBAN", OLD_IBAN1, new_iban1,
                                        occurrence=0)

    # ── 10. IBAN №2 (валютний — зарахування) ─────────────────────────────────
    new_iban2 = iban_to_digits(iban_credit)
    content = replace_individual_digits(content, "IBAN", OLD_IBAN2, new_iban2,
                                        occurrence=1)

    # ── 11. IBAN №3 (комісія) ─────────────────────────────────────────────────
    new_iban3 = iban_to_digits(iban_commission)
    content = replace_individual_digits(content, "IBAN", OLD_IBAN3, new_iban3,
                                        occurrence=2)

    # Зберігаємо XML
    with open(doc_path, "w", encoding="utf-8") as f:
        f.write(content)

    # Пакуємо назад у docx
    subprocess.run(
        [sys.executable, f"{scripts_dir}/pack.py",
         unpack_dir, output_docx,
         "--original", template_path, "--validate", "false"],
        check=True, capture_output=True
    )

    shutil.rmtree(tmpdir)
    return output_docx


def docx_to_pdf(docx_path: str, output_pdf: str) -> str:
    """Конвертує docx у PDF через LibreOffice."""
    scripts_dir = "/mnt/skills/public/docx/scripts/office"
    out_dir = os.path.dirname(output_pdf)

    result = subprocess.run(
        [sys.executable, f"{scripts_dir}/soffice.py",
         "--headless", "--convert-to", "pdf",
         "--outdir", out_dir, docx_path],
        capture_output=True, text=True
    )

    # LibreOffice створює файл з тим самим ім'ям але .pdf
    base = os.path.splitext(os.path.basename(docx_path))[0]
    generated = os.path.join(out_dir, base + ".pdf")

    if os.path.exists(generated) and generated != output_pdf:
        shutil.move(generated, output_pdf)

    return output_pdf


if __name__ == "__main__":
    # Тестовий запуск
    test_params = {
        "pib": "Мацола Євгеній Володимирович",
        "ipn": "3522908011",
        "address": "Україна, м. Дніпро, вул. Центральна, буд. 1, кв. 1",
        "date": "",
        "purpose_uk": "Передплата за товари згідно контракту 01/26 від 16.03.2026, інвойсу PI2026031201 від 17.03.2026",
        "amount": "5800.00",
        "currency": "USD",
        "iban_debit": "UA293220010000026006310115156",
        "iban_credit": "UA353220010000026000370076190",
        "iban_commission": "UA293220010000026006310115156",
    }
    template = "/home/claude/dovidka_app/template_valyuta.docx"
    out_docx = "/tmp/test_valyuta.docx"
    out_pdf = "/tmp/test_valyuta.pdf"

    fill_template(test_params, template, out_docx)
    print(f"✅ DOCX: {out_docx}")
    docx_to_pdf(out_docx, out_pdf)
    print(f"✅ PDF:  {out_pdf}")
