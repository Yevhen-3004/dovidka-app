"""
fill_valyuta.py — заповнює шаблон заяви на купівлю валюти (docx → pdf)
Без залежностей від зовнішніх скриптів — тільки стандартна бібліотека Python.
"""
import re, shutil, subprocess, sys, os, tempfile, zipfile
from datetime import date
from lxml import etree

MONTHS_UK = {
    1:"січня",2:"лютого",3:"березня",4:"квітня",
    5:"травня",6:"червня",7:"липня",8:"серпня",
    9:"вересня",10:"жовтня",11:"листопада",12:"грудня",
}
CURRENCY_CODES = {"USD":"840","EUR":"978"}
CURRENCY_NAMES = {"USD":"Долар США","EUR":"Євро"}


def get_date_parts(date_str):
    if date_str and date_str.strip():
        p = date_str.strip().split(".")
        d,m,y = int(p[0]),int(p[1]),int(p[2])
    else:
        t = date.today(); d,m,y = t.day,t.month,t.year
    return str(d), MONTHS_UK[m], str(y)


def replace_text_in_cell(content, old_text, new_text):
    pattern = r'(<w:t[^>]*>)' + re.escape(old_text) + r'(</w:t>)'
    return re.sub(pattern, r'\g<1>' + new_text.replace('\\',r'\\') + r'\g<2>', content)


def replace_individual_digits(content, start_marker_text, old_digits, new_digits, occurrence=0):
    wt_pattern = re.compile(r'<w:t(?:[^>]*)>([^<]*)</w:t>')
    all_matches = list(wt_pattern.finditer(content))
    marker_indices = [i for i,m in enumerate(all_matches) if start_marker_text in m.group(1)]
    if occurrence >= len(marker_indices):
        return content
    start_idx = marker_indices[occurrence] + 1
    replacements = []
    old_pos = new_pos = 0
    i = start_idx
    while old_pos < len(old_digits) and i < len(all_matches):
        cell_text = all_matches[i].group(1)
        if cell_text == old_digits[old_pos]:
            if new_pos < len(new_digits):
                orig = all_matches[i].group(0)
                repl = orig.replace(f'>{old_digits[old_pos]}<', f'>{new_digits[new_pos]}<', 1)
                replacements.append((all_matches[i].start(), all_matches[i].end(), repl))
            old_pos += 1; new_pos += 1
        i += 1
    for start,end,new_val in reversed(replacements):
        content = content[:start] + new_val + content[end:]
    return content


def iban_to_digits(iban):
    return list(iban.replace(" ","").replace("-",""))


def amount_to_digits(amount_str):
    try:
        val = float(amount_str)
        int_part = str(int(val))
        dec_part = f"{val:.2f}".split('.')[1]
        return list(int_part) + [','] + list(dec_part)
    except Exception:
        return list(amount_str)


# ── Розпакування / запакування docx без зовнішніх скриптів ──────────────────

def unpack_docx(docx_path, out_dir):
    """Розпаковує docx (zip) у папку."""
    os.makedirs(out_dir, exist_ok=True)
    with zipfile.ZipFile(docx_path, 'r') as z:
        z.extractall(out_dir)


def pack_docx(src_dir, output_path, original_path):
    """Пакує папку назад у docx."""
    # Беремо список файлів із оригінального архіву щоб зберегти порядок
    with zipfile.ZipFile(original_path, 'r') as orig_z:
        orig_names = orig_z.namelist()

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        # Спочатку файли з оригінального архіву (в тому ж порядку)
        for name in orig_names:
            file_path = os.path.join(src_dir, name.replace('/', os.sep))
            if os.path.isfile(file_path):
                zf.write(file_path, name)
        # Потім будь-які нові файли яких не було в оригіналі
        for root, dirs, files in os.walk(src_dir):
            for fname in files:
                fpath = os.path.join(root, fname)
                arcname = os.path.relpath(fpath, src_dir).replace(os.sep, '/')
                if arcname not in orig_names:
                    zf.write(fpath, arcname)


def find_libreoffice():
    """Шукає виконуваний файл LibreOffice."""
    candidates = [
        "libreoffice", "soffice",
        "/usr/bin/libreoffice", "/usr/bin/soffice",
        "/usr/lib/libreoffice/program/soffice",
        "/opt/libreoffice/program/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for c in candidates:
        try:
            r = subprocess.run([c,"--version"], capture_output=True, timeout=5)
            if r.returncode == 0:
                return c
        except Exception:
            pass
    return None


def docx_to_pdf(docx_path, output_pdf):
    """Конвертує docx → pdf через LibreOffice."""
    lo = find_libreoffice()
    if not lo:
        raise RuntimeError("LibreOffice не знайдено. Встановіть: apt-get install libreoffice")
    out_dir = os.path.dirname(output_pdf)
    subprocess.run(
        [lo, "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx_path],
        check=True, capture_output=True, timeout=60
    )
    base = os.path.splitext(os.path.basename(docx_path))[0]
    generated = os.path.join(out_dir, base + ".pdf")
    if os.path.exists(generated) and generated != output_pdf:
        shutil.move(generated, output_pdf)
    return output_pdf


# ── Основна функція заповнення шаблону ──────────────────────────────────────

def fill_template(params, template_path, output_docx):
    tmpdir = tempfile.mkdtemp()
    unpack_dir = os.path.join(tmpdir, "unpacked")

    unpack_docx(template_path, unpack_dir)

    doc_path = os.path.join(unpack_dir, "word", "document.xml")
    content = open(doc_path, encoding="utf-8").read()

    pib = params["pib"]
    pib_parts = pib.split()
    ipn = params["ipn"]
    address = params["address"]
    day, month_uk, year = get_date_parts(params.get("date",""))
    purpose_uk = params["purpose_uk"]
    amount = params["amount"]
    currency = params["currency"]
    cur_code = CURRENCY_CODES.get(currency,"840")
    iban_debit = params["iban_debit"].replace(" ","")
    iban_credit = params["iban_credit"].replace(" ","")
    iban_commission = params["iban_commission"].replace(" ","")

    OLD_PIB_PARTS = ["Каллаш","Леонід","Юрійович"]
    OLD_IPN = "3073810850"
    OLD_ADDRESS_1 = "Україна, м. Київ, пров. Костя Гордієнка, буд."
    OLD_ADDRESS_2 = "10, кв. 7"
    OLD_DAY = "23"; OLD_MONTH = "березня"; OLD_YEAR = "2026"
    OLD_AMOUNT_DIGITS = ["5","8","0","0",",","0","0"]
    OLD_CUR_CODE_DIGITS = ["8","4","0"]
    OLD_IBAN1 = list("UA293220010000026006310115156")
    OLD_IBAN2 = list("UA353220010000026000370076190")
    OLD_IBAN3 = list("UA293220010000026006310115156")

    # 1. Дата
    content = replace_text_in_cell(content, OLD_DAY, day)
    content = replace_text_in_cell(content, OLD_MONTH, month_uk)
    content = replace_text_in_cell(content, OLD_YEAR, year)

    # 2. ПІБ (3 окремі runs вгорі + внизу)
    for i,old_part in enumerate(OLD_PIB_PARTS):
        new_part = pib_parts[i] if i < len(pib_parts) else ""
        content = replace_text_in_cell(content, old_part, new_part)

    # 3. Адреса
    content = replace_text_in_cell(content, OLD_ADDRESS_1, address)
    content = replace_text_in_cell(content, OLD_ADDRESS_2, " ")

    # 4. ІПН
    content = replace_text_in_cell(content, OLD_IPN, ipn)

    # 5. Мета купівлі (українська) — замінюємо перший run, решту очищаємо
    wt_pattern = re.compile(r'(<w:t(?:[^>]*)>)([^<]*)(</w:t>)')
    all_wt = list(wt_pattern.finditer(content))
    purpose_start = next((i for i,m in enumerate(all_wt) if m.group(2)=="Передплата"), None)
    if purpose_start is not None:
        m0 = all_wt[purpose_start]
        content = content[:m0.start()] + m0.group(1) + purpose_uk + m0.group(3) + content[m0.end():]
        all_wt = list(wt_pattern.finditer(content))
        purpose_start = next((i for i,m in enumerate(all_wt) if m.group(2)==purpose_uk), purpose_start)
        i = purpose_start + 1
        replacements = []
        while i < len(all_wt):
            if "(Оплата/передплата" in all_wt[i].group(2):
                break
            if all_wt[i].group(2).strip():
                replacements.append((all_wt[i].start(), all_wt[i].end(),
                                      all_wt[i].group(1) + "" + all_wt[i].group(3)))
            i += 1
        for s,e,v in reversed(replacements):
            content = content[:s] + v + content[e:]

    # 6. Сума
    content = replace_individual_digits(content,"купівлі",OLD_AMOUNT_DIGITS,
                                        amount_to_digits(amount),occurrence=1)

    # 7. Код валюти
    content = replace_individual_digits(content,"Currency",OLD_CUR_CODE_DIGITS,
                                        list(cur_code),occurrence=0)

    # 8. Назва валюти
    if currency == "EUR":
        content = replace_text_in_cell(content,"Долар","Євро")
        content = replace_text_in_cell(content,"США"," ")

    # 9-11. IBAN
    content = replace_individual_digits(content,"IBAN",OLD_IBAN1,iban_to_digits(iban_debit),occurrence=0)
    content = replace_individual_digits(content,"IBAN",OLD_IBAN2,iban_to_digits(iban_credit),occurrence=1)
    content = replace_individual_digits(content,"IBAN",OLD_IBAN3,iban_to_digits(iban_commission),occurrence=2)

    with open(doc_path,"w",encoding="utf-8") as f:
        f.write(content)

    pack_docx(unpack_dir, output_docx, template_path)
    shutil.rmtree(tmpdir)
    return output_docx


if __name__ == "__main__":
    params = {
        "pib":"Мацола Євгеній Володимирович","ipn":"3522908011",
        "address":"Україна, обл. Закарпатська, р-н. Тячівський, с. Грушово, вул. Центральна, буд. 53-А",
        "date":"","purpose_uk":"Передоплата за товар згідно з договором № 01/26 від 16.03.2026",
        "amount":"5800.00","currency":"USD",
        "iban_debit":"UA293220010000026006310115156",
        "iban_credit":"UA353220010000026000370076190",
        "iban_commission":"UA293220010000026006310115156",
    }
    out = "/tmp/test_fill.docx"
    fill_template(params, "/home/claude/dovidka_app/template_valyuta.docx", out)
    print(f"DOCX: {out}")
    docx_to_pdf(out, "/tmp/test_fill.pdf")
    print("PDF: /tmp/test_fill.pdf")
