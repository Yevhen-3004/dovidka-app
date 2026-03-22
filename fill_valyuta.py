"""
fill_valyuta.py — заповнює шаблон заяви через python-docx (XML-рівень).
Структура нового шаблону (Валюта_ФОП.docx):
  T0  — заголовок, дата, ПІБ, адреса, ІПН
  T1  — мета купівлі
  T2  — сума (9 клітинок: 6 цілих + кома + 2 дробових), код і назва валюти
  T3  — курс банку
  T4  — строк дії
  T5  — IBAN 1 (гривневий, 30 клітинок)
  T6  — назва банку
  T7  — IBAN 2 (валютний, 30 клітинок)
  T8  — IBAN 3 (комісія, 30 клітинок)
  T9  — підпис (ПІБ внизу)
  T10 — позначки банку
  T11 — уповноважений працівник
"""
import io, os
from datetime import date
from docx import Document

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

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


def get_tcs(row):
    return row._tr.findall('.//w:tc', NS)


def set_tc_text(tc, new_text):
    wt_list = tc.findall('.//w:t', NS)
    if wt_list:
        wt_list[0].text = new_text
        for wt in wt_list[1:]:
            wt.text = ""


def replace_iban_tcs(tcs, new_iban):
    """tc[0]=IBAN мітка, tc[1..29]=символи."""
    chars = list(new_iban.replace(" ", "").ljust(29))[:29]
    for i, ch in enumerate(chars):
        idx = i + 1
        if idx < len(tcs):
            set_tc_text(tcs[idx], ch)


def replace_amount_tcs(tcs, new_amount_str):
    """
    tc[0]=мітка, tc[1..9]=символи суми.
    9 позицій: до 6 цілих + кома + 2 дробових = 999999,99.
    """
    try:
        val = float(new_amount_str)
        int_p = str(int(val))
        dec_p = f"{val:.2f}".split(".")[1]
        chars = list(int_p) + [","] + list(dec_p)
    except Exception:
        return

    # Записуємо зліва без вирівнювання, зайві клітинки — порожні
    for i in range(1, len(tcs)):
        char_idx = i - 1
        ch = chars[char_idx] if char_idx < len(chars) else ""
        set_tc_text(tcs[i], ch)


def build_zayava_docx(params):
    template_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "template_valyuta.docx")

    pib         = params["pib"]
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
    T = doc.tables

    # ── T0 r2c1: Дата ─────────────────────────────────────────────────────
    # run3 = '\u200223  ' (день), run5 = ' березня', run6 = '\u20022026\u2002...'
    tc_date = get_tcs(T[0].rows[2])[1]
    runs = tc_date.findall('.//w:r', NS)
    for run in runs:
        wts = run.findall('w:t', NS)
        full = ''.join(w.text or '' for w in wts)
        if '23' in full and '\u2002' not in full and '"' not in full:
            # run3: '\u200223  '
            for w in wts:
                if w.text and '23' in w.text:
                    w.text = w.text.replace('23', day)
        elif 'березня' in full:
            for w in wts:
                if w.text and 'березня' in w.text:
                    w.text = w.text.replace('березня', month_uk)
        elif '2026' in full:
            for w in wts:
                if w.text and '2026' in w.text:
                    w.text = w.text.replace('2026', year)

    # ── T0 r5c1: ПІБ вгорі ────────────────────────────────────────────────
    tc_pib = get_tcs(T[0].rows[5])[1]
    for wt in tc_pib.findall('.//w:t', NS):
        if 'Палажій' in (wt.text or '') or 'ФОП' in (wt.text or ''):
            wt.text = f"ФОП {pib}"
            break

    # ── T0 r6c1: Адреса ───────────────────────────────────────────────────
    tc_addr = get_tcs(T[0].rows[6])[1]
    for wt in tc_addr.findall('.//w:t', NS):
        if wt.text and wt.text.strip():
            wt.text = address
            break

    # ── T0 r6c3: ІПН ──────────────────────────────────────────────────────
    tc_ipn = get_tcs(T[0].rows[6])[3]
    for wt in tc_ipn.findall('.//w:t', NS):
        if wt.text and wt.text.strip():
            wt.text = ipn
            break

    # ── T1 r0c1: Мета купівлі ─────────────────────────────────────────────
    tc_purpose = get_tcs(T[1].rows[0])[1]
    wts = tc_purpose.findall('.//w:t', NS)
    if wts:
        wts[0].text = purpose_uk
        for wt in wts[1:]:
            wt.text = ""

    # ── T2 r0: Сума (tc[1..9]) ────────────────────────────────────────────
    tcs_amount = get_tcs(T[2].rows[0])
    replace_amount_tcs(tcs_amount, amount)

    # ── T2 r0: Код валюти (tc[14..16]) ────────────────────────────────────
    cur_chars = list(cur_code.ljust(3))[:3]
    for i, ch in enumerate(cur_chars):
        idx = 14 + i
        if idx < len(tcs_amount):
            set_tc_text(tcs_amount[idx], ch)

    # ── T2 r1: Назва валюти (tc[1]) ───────────────────────────────────────
    tcs_r1 = get_tcs(T[2].rows[1])
    if len(tcs_r1) > 1:
        set_tc_text(tcs_r1[1], cur_name_uk)

    # ── T5: IBAN 1 (гривневий) ────────────────────────────────────────────
    replace_iban_tcs(get_tcs(T[5].rows[0]), iban_debit)

    # ── T7: IBAN 2 (валютний) ─────────────────────────────────────────────
    replace_iban_tcs(get_tcs(T[7].rows[0]), iban_credit)

    # ── T8: IBAN 3 (комісія) ──────────────────────────────────────────────
    replace_iban_tcs(get_tcs(T[8].rows[0]), iban_comm)

    # ── T9 r0c2: ПІБ внизу ────────────────────────────────────────────────
    tc_pib_bot = get_tcs(T[9].rows[0])[2]
    for wt in tc_pib_bot.findall('.//w:t', NS):
        if 'Палажій' in (wt.text or '') or (wt.text and len(wt.text.strip()) > 5):
            # Зберігаємо пробіли перед ПІБ
            spaces = len(wt.text) - len(wt.text.lstrip())
            wt.text = ' ' * spaces + pib
            break

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


def build_zayava_pdf(params):
    return build_zayava_docx(params)

def fill_template(params, template_path, output_docx): pass
def docx_to_pdf(docx_path, output_pdf): pass


if __name__ == "__main__":
    tests = [
        ("125.00",    "EUR"),
        ("9999.99",   "USD"),
        ("12500.34",  "EUR"),
        ("102003.37", "USD"),
    ]
    base = {
        "pib":    "Мацола Євгеній Володимирович",
        "ipn":    "3522908011",
        "address":"Україна, обл. Закарпатська, р-н. Тячівський, с. Грушово, вул. Центральна, буд. 53-А",
        "date":   "22.03.2026",
        "purpose_uk": "Передоплата за товар згідно з договором № 01/26 від 16.03.2026",
        "iban_debit":      "UA823220010000026205356515049",
        "iban_credit":     "UA943220010000026204354428300",
        "iban_commission": "UA963220010000026206310213171",
    }
    import io as _io
    for amount, currency in tests:
        data = build_zayava_docx(dict(base, amount=amount, currency=currency))
        doc2 = Document(_io.BytesIO(data))
        tcs = doc2.tables[2].rows[0]._tr.findall('.//w:tc', NS)
        vals = [''.join(t.text or '' for t in tcs[i].findall('.//w:t',NS)) for i in range(1,10)]
        cur = [''.join(t.text or '' for t in tcs[14+i].findall('.//w:t',NS)) for i in range(3)]
        print(f"  {amount:>12} {currency} → {''.join(vals)}  код={''.join(cur)}")

    # Зберігаємо останній для перевірки
    data = build_zayava_docx(dict(base, amount="12500.34", currency="EUR"))
    with open("/tmp/test_new.docx", "wb") as f:
        f.write(data)
    print("\n✅ /tmp/test_new.docx")
