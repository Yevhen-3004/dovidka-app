from flask import Flask, request, send_file, render_template_string, jsonify
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY, TA_RIGHT
from datetime import date
import io, os, sys, zipfile, tempfile, json, urllib.request, urllib.error

app = Flask(__name__)

# ── Шлях до шаблону заяви ────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_VALYUTA = os.path.join(BASE_DIR, "template_valyuta.docx")

# ── Системний промпт для перекладу мети ──────────────────────────────────────
TRANSLATION_SYSTEM = """Ти — банківський перекладач. Переклади призначення платежу з англійської на українську згідно з правилами:

- Payment/Prepayment for goods → Оплата/Передоплата за товар
- Payment/Prepayment for services → Оплата/Передоплата за послуги
- contract nr X dd DD.MM.YYYY → договором № X від DD.MM.YYYY
- Invoice nr X dd DD.MM.YYYY → рахунком № X від DD.MM.YYYY (або рахунком-фактурою якщо інвойс виступає договором)
- Prepayment → передоплата (платіж ДО отримання товару/послуги, до митної декларації)
- Payment → оплата (товар вже розмитнений або послуга надана)
- according to → згідно з

Правила:
- Зберігай всі номери, дати, назви документів точно
- Відповідай ТІЛЬКИ перекладом без будь-яких пояснень, лапок чи додаткового тексту
- Перший символ — велика літера
"""

MONTHS_UK = {
    1:"січня",2:"лютого",3:"березня",4:"квітня",
    5:"травня",6:"червня",7:"липня",8:"серпня",
    9:"вересня",10:"жовтня",11:"листопада",12:"грудня"
}

HTML = """<!DOCTYPE html>
<html lang="uk">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Документи — Валютний контроль</title>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Sans:wght@300;400;500&display=swap" rel="stylesheet">
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  :root {
    --bg: #f5f3ef; --surface: #ffffff; --border: #d8d4cc;
    --text: #1a1916; --text-muted: #7a756c; --accent: #1a1916;
    --sans: 'IBM Plex Sans', sans-serif; --mono: 'IBM Plex Mono', monospace;
  }
  body { font-family: var(--sans); background: var(--bg); color: var(--text);
    min-height: 100vh; display: flex; justify-content: center; padding: 3rem 1rem; }
  .container { width: 100%; max-width: 580px; }
  header { margin-bottom: 2.5rem; }
  .badge { font-family: var(--mono); font-size: 10px; letter-spacing: 0.12em;
    text-transform: uppercase; color: var(--text-muted); border: 1px solid var(--border);
    display: inline-block; padding: 4px 10px; border-radius: 2px; margin-bottom: 1rem; }
  h1 { font-weight: 300; font-size: 28px; line-height: 1.2; letter-spacing: -0.02em; }
  h1 span { font-weight: 500; }
  .subtitle { margin-top: 0.5rem; font-size: 13px; color: var(--text-muted); font-weight: 300; }
  .card { background: var(--surface); border: 1px solid var(--border); border-radius: 4px; padding: 2rem; margin-bottom: 1rem; }
  .section-title { font-family: var(--mono); font-size: 10px; letter-spacing: 0.12em;
    text-transform: uppercase; color: var(--text-muted); margin-bottom: 1.25rem;
    padding-bottom: 8px; border-bottom: 1px solid var(--border); }
  .field { margin-bottom: 1.25rem; }
  .field:last-child { margin-bottom: 0; }
  label { display: block; font-family: var(--mono); font-size: 11px; letter-spacing: 0.08em;
    text-transform: uppercase; color: var(--text-muted); margin-bottom: 6px; }
  input[type="text"], input[type="number"], textarea { width: 100%; padding: 10px 14px;
    border: 1px solid var(--border); border-radius: 3px; background: var(--bg);
    color: var(--text); font-family: var(--sans); font-size: 14px; outline: none;
    transition: border-color 0.15s; }
  textarea { resize: vertical; min-height: 70px; line-height: 1.5; }
  input:focus, textarea:focus { border-color: var(--accent); }
  input::placeholder, textarea::placeholder { color: var(--text-muted); opacity: 0.6; }
  .two-col { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
  .divider { border: none; border-top: 1px solid var(--border); margin: 1.5rem 0; }
  .toggle-label { font-family: var(--mono); font-size: 11px; letter-spacing: 0.08em;
    text-transform: uppercase; color: var(--text-muted); margin-bottom: 8px; display: block; }
  .toggle-row { display: grid; grid-template-columns: 1fr 1fr;
    border: 1px solid var(--border); border-radius: 3px; overflow: hidden; }
  .toggle-btn { padding: 9px; font-family: var(--sans); font-size: 14px;
    background: var(--bg); color: var(--text-muted); border: none; cursor: pointer;
    transition: background 0.15s, color 0.15s; }
  .toggle-btn:first-child { border-right: 1px solid var(--border); }
  .toggle-btn.active { background: var(--accent); color: #fff; font-weight: 500; }
  .balance-fields { margin-top: 1rem; padding-top: 1rem; border-top: 1px dashed var(--border); }
  .amount-row { display: grid; grid-template-columns: 1fr auto; gap: 10px; align-items: flex-end; }
  .cur-row { display: flex; border: 1px solid var(--border); border-radius: 3px; overflow: hidden; }
  .cur-btn { padding: 10px 16px; font-family: var(--mono); font-size: 13px; font-weight: 500;
    background: var(--bg); color: var(--text-muted); border: none; cursor: pointer; }
  .cur-btn:first-child { border-right: 1px solid var(--border); }
  .cur-btn.active { background: var(--accent); color: #fff; }
  .submit-btn { width: 100%; padding: 13px; background: var(--accent); color: #fff;
    border: none; border-radius: 3px; font-family: var(--sans); font-size: 15px;
    font-weight: 500; cursor: pointer; transition: opacity 0.15s;
    display: flex; align-items: center; justify-content: center; gap: 8px; }
  .submit-btn:hover { opacity: 0.85; }
  .submit-btn.loading { opacity: 0.6; pointer-events: none; }
  .error-msg { margin-top: 1rem; padding: 10px 14px; background: #fee;
    border: 1px solid #fcc; border-radius: 3px; font-size: 13px; color: #a00; display: none; }
  .status-msg { margin-top: 1rem; padding: 10px 14px; background: #eef7ee;
    border: 1px solid #c3e6cb; border-radius: 3px; font-size: 13px; color: #2d6a4f; display: none; }
  .hint { font-size: 11px; color: var(--text-muted); margin-top: 5px; font-style: italic; }
  footer { margin-top: 1.5rem; text-align: center; font-family: var(--mono);
    font-size: 11px; color: var(--text-muted); letter-spacing: 0.04em; }
</style>
</head>
<body>
<div class="container">
  <header>
    <div class="badge">АТ Універсал Банк · Валютний контроль</div>
    <h1>Генератор<br><span>банківських документів</span></h1>
    <p class="subtitle">Заповніть форму — отримайте обидва PDF одразу</p>
  </header>

  <!-- Блок 1: Дані клієнта -->
  <div class="card">
    <div class="section-title">Дані клієнта</div>
    <div class="field">
      <label>ПІБ клієнта</label>
      <input type="text" id="pib" placeholder="Іваненко Іван Іванович" />
    </div>
    <div class="two-col">
      <div class="field">
        <label>ІПН</label>
        <input type="text" id="ipn" placeholder="1234567890" maxlength="10" />
      </div>
      <div class="field">
        <label>Дата</label>
        <input type="text" id="date" placeholder="сьогодні" />
      </div>
    </div>
    <div class="field">
      <label>Адреса клієнта</label>
      <input type="text" id="address" placeholder="Україна, обл. ..., вул. ..., буд. ..." />
    </div>
  </div>

  <!-- Блок 2: Залишок (довідка) -->
  <div class="card">
    <div class="section-title">Довідка про залишок валюти</div>
    <div class="field">
      <span class="toggle-label">Залишок на рахунку</span>
      <div class="toggle-row">
        <button class="toggle-btn active" id="btn-no" onclick="setBalance(false)">Відсутній</button>
        <button class="toggle-btn" id="btn-yes" onclick="setBalance(true)">Є залишок</button>
      </div>
      <div class="balance-fields" id="balance-section" style="display:none;">
        <div class="amount-row">
          <div class="field" style="margin-bottom:0;">
            <label>Сума залишку</label>
            <input type="number" id="balance_amount" placeholder="1238.33" step="0.01" min="0" />
          </div>
          <div>
            <label style="margin-bottom:6px;">Валюта</label>
            <div class="cur-row">
              <button class="cur-btn active" id="bal-usd" onclick="setBalCurrency('USD')">USD</button>
              <button class="cur-btn" id="bal-eur" onclick="setBalCurrency('EUR')">EUR</button>
            </div>
          </div>
        </div>
      </div>
    </div>
  </div>

  <!-- Блок 3: Заява на купівлю -->
  <div class="card">
    <div class="section-title">Заява на купівлю валюти</div>

    <div class="field">
      <label>Мета купівлі (англійською)</label>
      <textarea id="purpose_en" placeholder="Prepayment for goods according to contract nr 01/26 dd 16.03.2026, Invoice nr PI2026031201 dd 17.03.2026"></textarea>
      <p class="hint">Claude перекладе автоматично</p>
    </div>

    <div class="two-col">
      <div class="field">
        <label>Сума купівлі</label>
        <input type="number" id="buy_amount" placeholder="5800.00" step="0.01" min="0" />
      </div>
      <div class="field">
        <label>Валюта купівлі</label>
        <div class="cur-row" style="margin-top:2px;">
          <button class="cur-btn active" id="buy-usd" onclick="setBuyCurrency('USD')">USD</button>
          <button class="cur-btn" id="buy-eur" onclick="setBuyCurrency('EUR')">EUR</button>
        </div>
      </div>
    </div>

    <hr class="divider">

    <div class="field">
      <label>IBAN — списання гривні (купівля валюти)</label>
      <input type="text" id="iban_debit" placeholder="UA293220010000026006310115156" maxlength="29" />
    </div>
    <div class="field">
      <label>IBAN — зарахування валюти</label>
      <input type="text" id="iban_credit" placeholder="UA353220010000026000370076190" maxlength="29" />
    </div>
    <div class="field">
      <label>IBAN — списання комісії</label>
      <input type="text" id="iban_commission" placeholder="UA293220010000026006310115156" maxlength="29" />
    </div>
  </div>

  <div class="error-msg" id="error"></div>
  <div class="status-msg" id="status"></div>

  <button class="submit-btn" id="submit-btn" onclick="generate()">
    Згенерувати обидва PDF ↓
  </button>

  <footer style="margin-top:1.5rem;">Для внутрішнього використання</footer>
</div>

<script>
let hasBalance = false, balCurrency = 'USD', buyCurrency = 'USD';

function setBalance(val) {
  hasBalance = val;
  document.getElementById('btn-no').classList.toggle('active', !val);
  document.getElementById('btn-yes').classList.toggle('active', val);
  document.getElementById('balance-section').style.display = val ? 'block' : 'none';
}
function setBalCurrency(c) {
  balCurrency = c;
  document.getElementById('bal-usd').classList.toggle('active', c==='USD');
  document.getElementById('bal-eur').classList.toggle('active', c==='EUR');
}
function setBuyCurrency(c) {
  buyCurrency = c;
  document.getElementById('buy-usd').classList.toggle('active', c==='USD');
  document.getElementById('buy-eur').classList.toggle('active', c==='EUR');
}
function showError(msg) {
  const el = document.getElementById('error');
  el.textContent = msg; el.style.display = 'block';
  setTimeout(() => el.style.display = 'none', 4000);
}
function showStatus(msg) {
  const el = document.getElementById('status');
  el.textContent = msg; el.style.display = 'block';
}
function hideStatus() {
  document.getElementById('status').style.display = 'none';
}

function validate() {
  const pib = document.getElementById('pib').value.trim();
  const ipn = document.getElementById('ipn').value.trim();
  const address = document.getElementById('address').value.trim();
  const purpose = document.getElementById('purpose_en').value.trim();
  const buyAmt = document.getElementById('buy_amount').value.trim();
  const ibanD = document.getElementById('iban_debit').value.trim();
  const ibanC = document.getElementById('iban_credit').value.trim();
  const ibanK = document.getElementById('iban_commission').value.trim();
  const balAmt = document.getElementById('balance_amount').value.trim();

  if (!pib) return 'Заповніть ПІБ клієнта';
  if (!ipn || ipn.length !== 10) return 'ІПН має містити 10 цифр';
  if (!address) return 'Заповніть адресу клієнта';
  if (!purpose) return 'Вкажіть мету купівлі (англійською)';
  if (!buyAmt) return 'Вкажіть суму купівлі';
  if (!ibanD || ibanD.length < 20) return 'Перевірте IBAN для списання гривні';
  if (!ibanC || ibanC.length < 20) return 'Перевірте IBAN для зарахування валюти';
  if (!ibanK || ibanK.length < 20) return 'Перевірте IBAN для списання комісії';
  if (hasBalance && !balAmt) return 'Вкажіть суму залишку';
  return null;
}

async function generate() {
  const err = validate();
  if (err) { showError(err); return; }

  const btn = document.getElementById('submit-btn');
  btn.classList.add('loading');
  showStatus('⏳ Перекладаємо мету та генеруємо документи...');

  const payload = {
    pib: document.getElementById('pib').value.trim(),
    ipn: document.getElementById('ipn').value.trim(),
    address: document.getElementById('address').value.trim(),
    date: document.getElementById('date').value.trim(),
    purpose_en: document.getElementById('purpose_en').value.trim(),
    buy_amount: document.getElementById('buy_amount').value.trim(),
    buy_currency: buyCurrency,
    iban_debit: document.getElementById('iban_debit').value.trim(),
    iban_credit: document.getElementById('iban_credit').value.trim(),
    iban_commission: document.getElementById('iban_commission').value.trim(),
    has_balance: hasBalance,
    balance_amount: document.getElementById('balance_amount').value.trim(),
    balance_currency: balCurrency,
  };

  try {
    const resp = await fetch('/generate', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify(payload)
    });

    if (!resp.ok) {
      const data = await resp.json();
      showError(data.error || 'Помилка генерації');
      hideStatus();
    } else {
      const blob = await resp.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      const name = payload.pib.split(' ')[0];
      a.href = url; a.download = `Документи_${name}.zip`;
      document.body.appendChild(a); a.click(); document.body.removeChild(a);
      URL.revokeObjectURL(url);
      hideStatus();
      showStatus('✅ Документи успішно згенеровано!');
      setTimeout(() => hideStatus(), 4000);
    }
  } catch(e) {
    showError('Помилка з\'єднання з сервером');
    hideStatus();
  }

  btn.classList.remove('loading');
  btn.textContent = 'Згенерувати обидва PDF ↓';
}
</script>
</body>
</html>"""


# ── Допоміжні функції ────────────────────────────────────────────────────────

def translate_purpose(purpose_en: str) -> str:
    """Перекладає мету купівлі з англійської на українську через Claude API."""
    try:
        payload = json.dumps({
            "model": "claude-sonnet-4-20250514",
            "max_tokens": 1000,
            "system": TRANSLATION_SYSTEM,
            "messages": [{"role": "user", "content": purpose_en}]
        }).encode("utf-8")

        req = urllib.request.Request(
            "https://api.anthropic.com/v1/messages",
            data=payload,
            headers={
                "Content-Type": "application/json",
                "anthropic-version": "2023-06-01",
            },
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = json.loads(resp.read())
            return data["content"][0]["text"].strip()
    except Exception as e:
        # Fallback: базовий переклад
        return purpose_en


def register_fonts():
    paths = [
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
         "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        ("/usr/share/fonts/dejavu/DejaVuSans.ttf",
         "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf"),
    ]
    for reg, bold in paths:
        if os.path.exists(reg) and os.path.exists(bold):
            try:
                pdfmetrics.registerFont(TTFont("AppFont", reg))
                pdfmetrics.registerFont(TTFont("AppFont-Bold", bold))
                return "AppFont", "AppFont-Bold"
            except Exception:
                pass
    return "Helvetica", "Helvetica-Bold"


def get_date_str(date_str):
    if date_str and date_str.strip():
        return date_str.strip()
    return date.today().strftime("%d.%m.%Y")


def format_balance(amount, currency):
    try:
        val = float(amount)
        if val > 0 and currency:
            fmt = f"{val:,.2f}".replace(",", " ").replace(".", ",")
            return f"{fmt} {currency} в АТ УніверсалБанк"
    except Exception:
        pass
    return None


def get_date_parts(date_str):
    if date_str and date_str.strip():
        parts = date_str.strip().split(".")
        d, m, y = int(parts[0]), int(parts[1]), int(parts[2])
    else:
        today = date.today()
        d, m, y = today.day, today.month, today.year
    return str(d), MONTHS_UK[m], str(y)


# ── Генерація довідки (PDF через ReportLab) ──────────────────────────────────

def build_dovidka_pdf(pib, ipn, doc_date, balance_line):
    font_name, font_bold = register_fonts()
    p1 = (f"1. на поточних рахунках - {balance_line};"
          if balance_line else "1. на поточних рахунках - відсутні;")

    sc  = ParagraphStyle("c",  fontName=font_name, fontSize=11, leading=16, alignment=TA_CENTER)
    scb = ParagraphStyle("cb", fontName=font_bold,  fontSize=11, leading=16, alignment=TA_CENTER)
    sj  = ParagraphStyle("j",  fontName=font_name, fontSize=11, leading=16, alignment=TA_JUSTIFY)
    sl  = ParagraphStyle("l",  fontName=font_name, fontSize=11, leading=16, alignment=TA_LEFT)
    ss  = ParagraphStyle("s",  fontName=font_name, fontSize=10, leading=14, alignment=TA_LEFT)
    ssr = ParagraphStyle("sr", fontName=font_name, fontSize=10, leading=14, alignment=TA_RIGHT)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            rightMargin=20*mm, leftMargin=20*mm,
                            topMargin=20*mm, bottomMargin=20*mm)
    s = []
    s.append(Paragraph("До уваги Управління з валютного контролю", sc))
    s.append(Spacer(1, 6*mm))
    s.append(Paragraph("ДОВІДКА", scb))
    s.append(Spacer(1, 5*mm))
    s.append(Paragraph(
        f"На додаток до заяви на купівлю іноземної валюти надаємо наступну інформацію "
        f"щодо наявності іноземної валюти на рахунках ФОП {pib} "
        f"{ipn} станом на початок операційного дня в уповноважених банках України "
        f"(суми необхідно вказати в розрізі валют та назв уповноважених банків):", sj))
    s.append(Spacer(1, 4*mm))
    s.append(Paragraph(p1, sl))
    s.append(Spacer(1, 2*mm))
    s.append(Paragraph("2. на вкладних (депозитних) рахунках - залишки в іноземних валютах відсутні;", sl))
    s.append(Spacer(1, 2*mm))
    s.append(Paragraph(
        "3. продана клієнтом іноземна валюта за першою частиною валютних операцій на "
        "умовах \"своп\" за незавершеними угодами з банками - відсутні.", sj))
    s.append(Spacer(1, 2*mm))
    s.append(Paragraph(
        "4. виключення відповідно до пункту 12-15 Постанови НБУ № 18 від 24.02.2022: "
        "кошти, що знаходяться під заставою/ накладено арешт/ у неплатоспроможних "
        "банках/ отримані за договорами комісії/ були куплені та не використані в "
        "установлений законодавством строк/ інші: відсутні", sj))
    s.append(Spacer(1, 8*mm))
    s.append(Paragraph("Підпис", sl))
    s.append(Spacer(1, 2*mm))
    s.append(Paragraph(f"ФОП {pib}", sl))
    s.append(Spacer(1, 2*mm))
    s.append(Paragraph("Електронний підпис", sl))
    s.append(Spacer(1, 4*mm))
    sig_table = Table(
        [[Paragraph(doc_date, ss), Paragraph(pib, ssr)]],
        colWidths=["40%", "60%"]
    )
    sig_table.setStyle(TableStyle([
        ("VALIGN",(0,0),(-1,-1),"TOP"),
        ("LEFTPADDING",(0,0),(-1,-1),0),
        ("RIGHTPADDING",(0,0),(-1,-1),0),
        ("TOPPADDING",(0,0),(-1,-1),0),
        ("BOTTOMPADDING",(0,0),(-1,-1),0),
    ]))
    s.append(sig_table)
    doc.build(s)
    buf.seek(0)
    return buf


# ── Генерація заяви на купівлю (через fill_valyuta) ──────────────────────────

def build_valyuta_pdf(params: dict) -> bytes:
    sys.path.insert(0, BASE_DIR)
    import fill_valyuta as fv

    tmpdir = tempfile.mkdtemp()
    out_docx = os.path.join(tmpdir, "valyuta.docx")
    out_pdf  = os.path.join(tmpdir, "valyuta.pdf")

    fv.fill_template(params, TEMPLATE_VALYUTA, out_docx)
    fv.docx_to_pdf(out_docx, out_pdf)

    with open(out_pdf, "rb") as f:
        data = f.read()

    import shutil
    shutil.rmtree(tmpdir)
    return data


# ── Flask маршрути ────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/generate", methods=["POST"])
def generate():
    try:
        data = request.get_json()

        pib          = data.get("pib", "").strip()
        ipn          = data.get("ipn", "").strip()
        address      = data.get("address", "").strip()
        date_str     = data.get("date", "").strip()
        purpose_en   = data.get("purpose_en", "").strip()
        buy_amount   = data.get("buy_amount", "0")
        buy_currency = data.get("buy_currency", "USD")
        iban_debit   = data.get("iban_debit", "").strip()
        iban_credit  = data.get("iban_credit", "").strip()
        iban_commission = data.get("iban_commission", "").strip()
        has_balance  = data.get("has_balance", False)
        bal_amount   = data.get("balance_amount", "0")
        bal_currency = data.get("balance_currency", "USD")

        # Переклад мети
        purpose_uk = translate_purpose(purpose_en)

        doc_date_str = get_date_str(date_str)
        balance_line = format_balance(bal_amount, bal_currency) if has_balance else None

        # ── Довідка PDF ───────────────────────────────────────────────────────
        dovidka_buf = build_dovidka_pdf(pib, ipn, doc_date_str, balance_line)

        # ── Заява PDF ─────────────────────────────────────────────────────────
        valyuta_params = {
            "pib": pib,
            "ipn": ipn,
            "address": address,
            "date": date_str,
            "purpose_uk": purpose_uk,
            "amount": buy_amount,
            "currency": buy_currency,
            "iban_debit": iban_debit,
            "iban_credit": iban_credit,
            "iban_commission": iban_commission,
        }
        valyuta_bytes = build_valyuta_pdf(valyuta_params)

        # ── Пакуємо в ZIP ─────────────────────────────────────────────────────
        safe = pib.split()[0] if pib else "doc"
        d    = doc_date_str.replace(".", "_")

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr(f"Dovidka_{safe}_{d}.pdf", dovidka_buf.read())
            zf.writestr(f"Zayava_{safe}_{d}.pdf", valyuta_bytes)
        zip_buf.seek(0)

        return send_file(
            zip_buf,
            mimetype="application/zip",
            as_attachment=True,
            download_name=f"Dokumenty_{safe}_{d}.zip"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"\n✅ Додаток запущено на http://localhost:{port}\n")
    app.run(host="0.0.0.0", port=port)
