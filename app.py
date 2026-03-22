from flask import Flask, request, send_file, render_template_string
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY, TA_RIGHT
from datetime import date
import io, os

app = Flask(__name__)

HTML = """<!DOCTYPE html>
<html lang="uk">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Довідка — Валютний контроль</title>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Sans:wght@300;400;500&display=swap" rel="stylesheet">
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  :root {
    --bg: #f5f3ef; --surface: #ffffff; --border: #d8d4cc;
    --text: #1a1916; --text-muted: #7a756c; --accent: #1a1916;
    --sans: 'IBM Plex Sans', sans-serif; --mono: 'IBM Plex Mono', monospace;
  }
  body { font-family: var(--sans); background: var(--bg); color: var(--text);
    min-height: 100vh; display: flex; align-items: flex-start;
    justify-content: center; padding: 3rem 1rem; }
  .container { width: 100%; max-width: 540px; }
  header { margin-bottom: 2.5rem; }
  .badge { font-family: var(--mono); font-size: 10px; letter-spacing: 0.12em;
    text-transform: uppercase; color: var(--text-muted); border: 1px solid var(--border);
    display: inline-block; padding: 4px 10px; border-radius: 2px; margin-bottom: 1rem; }
  h1 { font-weight: 300; font-size: 28px; line-height: 1.2; letter-spacing: -0.02em; }
  h1 span { font-weight: 500; }
  .subtitle { margin-top: 0.5rem; font-size: 13px; color: var(--text-muted); font-weight: 300; }
  .card { background: var(--surface); border: 1px solid var(--border); border-radius: 4px; padding: 2rem; }
  .field { margin-bottom: 1.5rem; }
  .field:last-child { margin-bottom: 0; }
  label { display: block; font-family: var(--mono); font-size: 11px; letter-spacing: 0.08em;
    text-transform: uppercase; color: var(--text-muted); margin-bottom: 8px; }
  input[type="text"], input[type="number"] { width: 100%; padding: 10px 14px;
    border: 1px solid var(--border); border-radius: 3px; background: var(--bg);
    color: var(--text); font-family: var(--sans); font-size: 15px; outline: none;
    transition: border-color 0.15s; }
  input:focus { border-color: var(--accent); }
  input::placeholder { color: var(--text-muted); opacity: 0.6; }
  .divider { border: none; border-top: 1px solid var(--border); margin: 1.75rem 0; }
  .toggle-label { font-family: var(--mono); font-size: 11px; letter-spacing: 0.08em;
    text-transform: uppercase; color: var(--text-muted); margin-bottom: 10px; display: block; }
  .toggle-row { display: grid; grid-template-columns: 1fr 1fr;
    border: 1px solid var(--border); border-radius: 3px; overflow: hidden; }
  .toggle-btn { padding: 10px; font-family: var(--sans); font-size: 14px;
    background: var(--bg); color: var(--text-muted); border: none; cursor: pointer;
    transition: background 0.15s, color 0.15s; }
  .toggle-btn:first-child { border-right: 1px solid var(--border); }
  .toggle-btn.active { background: var(--accent); color: #fff; font-weight: 500; }
  .balance-fields { margin-top: 1.25rem; padding-top: 1.25rem; border-top: 1px dashed var(--border); }
  .amount-row { display: grid; grid-template-columns: 1fr auto; gap: 10px; align-items: flex-start; }
  .currency-group { display: flex; flex-direction: column; gap: 6px; }
  .currency-group label { margin-bottom: 0; }
  .cur-row { display: flex; border: 1px solid var(--border); border-radius: 3px; overflow: hidden; }
  .cur-btn { padding: 10px 18px; font-family: var(--mono); font-size: 13px; font-weight: 500;
    background: var(--bg); color: var(--text-muted); border: none; cursor: pointer;
    transition: background 0.15s, color 0.15s; }
  .cur-btn:first-child { border-right: 1px solid var(--border); }
  .cur-btn.active { background: var(--accent); color: #fff; }
  .submit-btn { width: 100%; margin-top: 2rem; padding: 13px; background: var(--accent);
    color: #fff; border: none; border-radius: 3px; font-family: var(--sans); font-size: 15px;
    font-weight: 500; cursor: pointer; transition: opacity 0.15s;
    display: flex; align-items: center; justify-content: center; gap: 8px; }
  .submit-btn:hover { opacity: 0.85; }
  .submit-btn.loading { opacity: 0.6; pointer-events: none; }
  .error-msg { margin-top: 1rem; padding: 10px 14px; background: #fee;
    border: 1px solid #fcc; border-radius: 3px; font-size: 13px; color: #a00; display: none; }
  footer { margin-top: 1.5rem; text-align: center; font-family: var(--mono);
    font-size: 11px; color: var(--text-muted); letter-spacing: 0.04em; }
</style>
</head>
<body>
<div class="container">
  <header>
    <div class="badge">НБУ · Валютний контроль</div>
    <h1>Генератор<br><span>довідки про залишки</span></h1>
    <p class="subtitle">Заповніть форму — отримайте PDF одразу</p>
  </header>
  <div class="card">
    <div class="field">
      <label>ПІБ клієнта</label>
      <input type="text" id="pib" placeholder="Іваненко Іван Іванович" />
    </div>
    <div class="field">
      <label>ІПН клієнта</label>
      <input type="text" id="ipn" placeholder="1234567890" maxlength="10" />
    </div>
    <div class="field">
      <label>Дата довідки</label>
      <input type="text" id="date" placeholder="Залиш порожнім — підставиться сьогоднішня" />
    </div>
    <hr class="divider">
    <div class="field">
      <span class="toggle-label">Залишок на рахунку</span>
      <div class="toggle-row">
        <button class="toggle-btn active" id="btn-no" onclick="setBalance(false)">Відсутній</button>
        <button class="toggle-btn" id="btn-yes" onclick="setBalance(true)">Є залишок</button>
      </div>
      <div class="balance-fields" id="balance-section" style="display:none;">
        <div class="amount-row">
          <div class="field" style="margin-bottom:0;">
            <label>Сума</label>
            <input type="number" id="amount" placeholder="1238.33" step="0.01" min="0" />
          </div>
          <div class="currency-group">
            <label>Валюта</label>
            <div class="cur-row">
              <button class="cur-btn active" id="btn-usd" onclick="setCurrency('USD')">USD</button>
              <button class="cur-btn" id="btn-eur" onclick="setCurrency('EUR')">EUR</button>
            </div>
          </div>
        </div>
      </div>
    </div>
    <div class="error-msg" id="error"></div>
    <button class="submit-btn" id="submit-btn" onclick="generate()">
      Згенерувати PDF ↓
    </button>
  </div>
  <footer>Для внутрішнього використання</footer>
</div>
<script>
  let hasBalance = false, currency = 'USD';
  function setBalance(val) {
    hasBalance = val;
    document.getElementById('btn-no').classList.toggle('active', !val);
    document.getElementById('btn-yes').classList.toggle('active', val);
    document.getElementById('balance-section').style.display = val ? 'block' : 'none';
  }
  function setCurrency(cur) {
    currency = cur;
    document.getElementById('btn-usd').classList.toggle('active', cur === 'USD');
    document.getElementById('btn-eur').classList.toggle('active', cur === 'EUR');
  }
  function showError(msg) {
    const el = document.getElementById('error');
    el.textContent = msg; el.style.display = 'block';
    setTimeout(() => el.style.display = 'none', 3000);
  }
  function generate() {
    const pib = document.getElementById('pib').value.trim();
    const ipn = document.getElementById('ipn').value.trim();
    const dateVal = document.getElementById('date').value.trim();
    const amount = document.getElementById('amount').value.trim();
    if (!pib) { showError('Заповніть ПІБ клієнта'); return; }
    if (!ipn || ipn.length !== 10) { showError('ІПН має містити 10 цифр'); return; }
    if (hasBalance && !amount) { showError('Вкажіть суму залишку'); return; }
    const btn = document.getElementById('submit-btn');
    btn.classList.add('loading'); btn.textContent = 'Генерується...';
    const params = new URLSearchParams({
      pib, ipn, date: dateVal,
      has_balance: hasBalance ? '1' : '0',
      amount: hasBalance ? amount : '0',
      currency: hasBalance ? currency : ''
    });
    const link = document.createElement('a');
    link.href = '/generate?' + params.toString();
    link.download = ''; document.body.appendChild(link);
    link.click(); document.body.removeChild(link);
    setTimeout(() => { btn.classList.remove('loading'); btn.textContent = 'Згенерувати PDF ↓'; }, 2000);
  }
</script>
</body>
</html>"""


def register_fonts():
    paths = [
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
         "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        ("/usr/share/fonts/dejavu/DejaVuSans.ttf",
         "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf"),
    ]
    for reg, bold in paths:
        if os.path.exists(reg) and os.path.exists(bold):
            pdfmetrics.registerFont(TTFont("AppFont", reg))
            pdfmetrics.registerFont(TTFont("AppFont-Bold", bold))
            return "AppFont", "AppFont-Bold"
    return "Helvetica", "Helvetica-Bold"


def format_amount(amount, currency):
    try:
        val = float(amount)
        if val > 0 and currency:
            formatted = f"{val:,.2f}".replace(",", " ").replace(".", ",")
            return f"{formatted} {currency} в АТ УніверсалБанк"
    except (ValueError, TypeError):
        pass
    return None


def get_date(date_str):
    if date_str and date_str.strip():
        return date_str.strip()
    return date.today().strftime("%d.%m.%Y")


def build_pdf(pib, ipn, doc_date, balance_line):
    font_name, font_bold = register_fonts()
    p1 = f"1. на поточних рахунках - {balance_line};" if balance_line else "1. на поточних рахунках - відсутні;"
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
        ("VALIGN",        (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING",   (0,0), (-1,-1), 0),
        ("RIGHTPADDING",  (0,0), (-1,-1), 0),
        ("TOPPADDING",    (0,0), (-1,-1), 0),
        ("BOTTOMPADDING", (0,0), (-1,-1), 0),
    ]))
    s.append(sig_table)
    doc.build(s)
    buf.seek(0)
    return buf


@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/generate")
def generate():
    pib         = request.args.get("pib", "").strip()
    ipn         = request.args.get("ipn", "").strip()
    date_str    = request.args.get("date", "").strip()
    has_balance = request.args.get("has_balance", "0") == "1"
    amount      = request.args.get("amount", "0")
    currency    = request.args.get("currency", "")
    doc_date     = get_date(date_str)
    balance_line = format_amount(amount, currency) if has_balance else None
    buf          = build_pdf(pib, ipn, doc_date, balance_line)
    safe_pib     = pib.split()[0] if pib else "dovidka"
    filename     = f"Dovidka_{safe_pib}_{doc_date.replace('.','_')}.pdf"
    return send_file(buf, mimetype="application/pdf",
                     as_attachment=True, download_name=filename)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"\n✅ Додаток запущено на http://localhost:{port}\n")
    app.run(host="0.0.0.0", port=port)
