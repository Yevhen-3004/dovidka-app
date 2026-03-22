"""
fill_valyuta.py — заповнює шаблон заяви через python-docx.
Повертає bytes готового .docx файлу (без конвертації в PDF).
"""
import io, os, copy
from datetime import date
from docx import Document
from docx.oxml.ns import qn

MONTHS_UK = {
    1:"січня",2:"лютого",3:"березня",4:"квітня",
    5:"травня",6:"червня",7:"липня",8:"серпня",
    9:"вересня",10:"жовтня",11:"листопада",12:"грудня",
}
CURRENCY_CODES    = {"USD":"840","EUR":"978"}
CURRENCY_NAMES_UK = {"USD":"Долар США","EUR":"Євро"}


def get_date_parts(date_str):
    if date_str and date_str.strip():
        p = date_str.strip().split(".")
        d, m, y = int(p[0]), int(p[1]), int(p[2])
    else:
        t = date.today(); d, m, y = t.day, t.month, t.year
    return str(d), MONTHS_UK[m], str(y)


def set_run(run, new_text):
    """Замінює текст run, зберігаючи форматування."""
    run.text = new_text


def replace_runs_in_cell(cell, replacements):
    """
    replacements: dict {old_text: new_text}
    Знаходить run з точним текстом і замінює його.
    """
    for para in cell.paragraphs:
        for run in para.runs:
            if run.text in replacements:
                set_run(run, replacements[run.text])


def replace_all_matching_runs(tables, old_text, new_text):
    """Замінює всі runs з точним текстом у всіх таблицях."""
    for table in tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        if run.text == old_text:
                            set_run(run, new_text)


def replace_iban(tables, table_idx, old_iban, new_iban):
    """
    Замінює символи IBAN у таблиці. Символи зберігаються як окремі runs
    в послідовних клітинках.
    old_iban / new_iban — рядки без пробілів.
    """
    old_chars = list(old_iban)
    new_chars = list(new_iban.ljust(len(old_chars)))[:len(old_chars)]

    table = tables[table_idx]
    # Збираємо всі runs по порядку
    all_runs = []
    for row in table.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                for run in para.runs:
                    all_runs.append(run)

    run_texts = [r.text for r in all_runs]

    # Шукаємо послідовність old_chars
    for i in range(len(run_texts) - len(old_chars) + 1):
        if run_texts[i:i+len(old_chars)] == old_chars:
            for j, ch in enumerate(new_chars):
                set_run(all_runs[i+j], ch)
            return True
    return False


def replace_amount(tables, table_idx, old_digits, new_amount_str):
    """Замінює цифри суми (кожна в окремому run/клітинці)."""
    try:
        val = float(new_amount_str)
        int_p = str(int(val))
        dec_p = f"{val:.2f}".split(".")[1]
        new_chars = list(int_p) + [","] + list(dec_p)
    except Exception:
        return False

    # Вирівнюємо по довжині старих digits
    while len(new_chars) < len(old_digits):
        new_chars.insert(0, " ")
    new_chars = new_chars[-len(old_digits):]

    table = tables[table_idx]
    all_runs = []
    for row in table.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                for run in para.runs:
                    all_runs.append(run)

    run_texts = [r.text for r in all_runs]
    for i in range(len(run_texts) - len(old_digits) + 1):
        if run_texts[i:i+len(old_digits)] == old_digits:
            for j, ch in enumerate(new_chars):
                set_run(all_runs[i+j], ch if ch.strip() else " ")
            return True
    return False


def replace_purpose(table, new_purpose):
    """
    Замінює мету купівлі. Мета в T1 r0c1 — кілька runs що утворюють текст.
    Ми записуємо весь новий текст в перший run, решту очищаємо.
    """
    cell = table.rows[0].cells[1]  # r0c1 — права клітинка з метою
    para = cell.paragraphs[0]
    runs = para.runs
    if not runs:
        return
    # Перший run отримує весь текст
    set_run(runs[0], new_purpose)
    # Решта runs — очищаємо
    for run in runs[1:]:
        set_run(run, "")


def build_zayava_docx(params):
    """Заповнює шаблон і повертає bytes готового .docx."""
    template_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "template_valyuta.docx")

    pib         = params["pib"]
    pib_parts   = pib.split()
    ipn         = params["ipn"]
    address     = params["address"]
    purpose_uk  = params["purpose_uk"]
    amount      = params["amount"]
    currency    = params["currency"]
    cur_code    = CURRENCY_CODES.get(currency, "840")
    cur_name_uk = CURRENCY_NAMES_UK.get(currency, "Долар США")
    iban_debit  = params["iban_debit"].replace(" ", "")
    iban_credit = params["iban_credit"].replace(" ", "")
    iban_comm   = params["iban_commission"].replace(" ", "")
    day, month_uk, year = get_date_parts(params.get("date", ""))

    doc = Document(template_path)
    tables = doc.tables

    # ── T0: Дата ─────────────────────────────────────────────────────────────
    # r3c1: runs: №, ' ', 5, ' ', ВІД/DATED, ' ', ", ' ', 23, ' ', ", ' ', березня, ' ', 2026\t, р.
    replace_all_matching_runs(tables, "23",      day)
    replace_all_matching_runs(tables, "березня", month_uk)
    replace_all_matching_runs(tables, "2026\t",  year + "\t")

    # ── T0: ПІБ вгорі (r6c1) ─────────────────────────────────────────────────
    replace_all_matching_runs(tables, "Каллаш",   pib_parts[0] if len(pib_parts) > 0 else "")
    replace_all_matching_runs(tables, "Леонід",   pib_parts[1] if len(pib_parts) > 1 else "")
    replace_all_matching_runs(tables, "Юрійович", pib_parts[2] if len(pib_parts) > 2 else "")

    # ── T0: Адреса (r7c1) ────────────────────────────────────────────────────
    # run0: 'Україна, м. Київ, пров. Костя Гордієнка, буд.'
    # run2: '10, кв. 7'
    replace_all_matching_runs(
        tables,
        "Україна, м. Київ, пров. Костя Гордієнка, буд.",
        address)
    replace_all_matching_runs(tables, "10, кв. 7", "")

    # ── T0: ІПН (r7c3) ───────────────────────────────────────────────────────
    replace_all_matching_runs(tables, "3073810850", ipn)

    # ── T1: Мета купівлі ─────────────────────────────────────────────────────
    replace_purpose(tables[1], purpose_uk)

    # ── T2: Сума ─────────────────────────────────────────────────────────────
    # Старі digits: ["5","8","0","0",",","0","0"]
    replace_amount(tables, 2, ["5","8","0","0",",","0","0"], amount)

    # ── T2: Код валюти ───────────────────────────────────────────────────────
    old_cur = ["8","4","0"]
    new_cur = list(cur_code.ljust(3))[:3]
    table2 = tables[2]
    all_runs2 = []
    for row in table2.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                for run in para.runs:
                    all_runs2.append(run)
    texts2 = [r.text for r in all_runs2]
    # Шукаємо "840" після "Currency"
    for i in range(len(texts2) - 3):
        if texts2[i:i+3] == old_cur:
            # Переконуємось що перед ним є "Currency"
            context = texts2[max(0,i-5):i]
            if any("Currency" in t or "Валюта" in t for t in context):
                for j in range(3):
                    set_run(all_runs2[i+j], new_cur[j])
                break

    # ── T2: Назва валюти ─────────────────────────────────────────────────────
    replace_all_matching_runs(tables, "Долар США", cur_name_uk)

    # ── T2: IBAN 1 ───────────────────────────────────────────────────────────
    replace_iban(tables, 2, "UA293220010000026006310115156", iban_debit)

    # ── T3: IBAN 2 ───────────────────────────────────────────────────────────
    replace_iban(tables, 3, "UA353220010000026000370076190", iban_credit)

    # ── T4: IBAN 3 ───────────────────────────────────────────────────────────
    replace_iban(tables, 4, "UA293220010000026006310115156", iban_comm)

    # ── T5: ПІБ внизу (підпис) ───────────────────────────────────────────────
    # r0c2p3: runs Каллаш, Леонід, Юрійович — вже замінені вище
    # (replace_all_matching_runs вже замінив їх)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# Заглушки для сумісності
def build_zayava_pdf(params):
    return build_zayava_docx(params)

def fill_template(params, template_path, output_docx): pass
def docx_to_pdf(docx_path, output_pdf): pass


if __name__ == "__main__":
    params = {
        "pib":    "Мацола Євгеній Володимирович",
        "ipn":    "3522908011",
        "address":"Україна, обл. Закарпатська, р-н. Тячівський, с. Грушово, вул. Центральна, буд. 53-А",
        "date":   "",
        "purpose_uk": "Передоплата за товар згідно з договором № 01/26 від 16.03.2026, рахунком № PI2026031201 від 17.03.2026",
        "amount": "5800.00", "currency": "USD",
        "iban_debit":      "UA293220010000026006310115156",
        "iban_credit":     "UA353220010000026000370076190",
        "iban_commission": "UA293220010000026006310115156",
    }
    docx_bytes = build_zayava_docx(params)
    with open("/tmp/test_zayava.docx","wb") as f:
        f.write(docx_bytes)
    print(f"✅ DOCX: {len(docx_bytes)} bytes → /tmp/test_zayava.docx")
