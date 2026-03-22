from flask import Flask, request, send_file, render_template_string, jsonify
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY, TA_RIGHT
from datetime import date
import io, os, sys, zipfile, tempfile, json, urllib.request

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_VALYUTA = os.path.join(BASE_DIR, "template_valyuta.docx")
HAS_API_KEY = bool(os.environ.get("ANTHROPIC_API_KEY", "").strip())

TRANSLATION_SYSTEM = """Ти — банківський перекладач. Переклади призначення платежу з англійської на українську.
Правила:
- Payment for goods → Оплата за товар
- Prepayment for goods → Передоплата за товар
- Payment for services → Оплата за послуги
- Prepayment for services → Передоплата за послуги
- according to → згідно з
- contract nr X dd DD.MM.YYYY → договором № X від DD.MM.YYYY
- Invoice nr X dd DD.MM.YYYY → рахунком № X від DD.MM.YYYY
- Зберігай всі номери та дати точно
- Відповідай ТІЛЬКИ перекладом, без пояснень і лапок
- Перший символ — велика літера"""

MONTHS_UK = {
    1:"січня",2:"лютого",3:"березня",4:"квітня",
    5:"травня",6:"червня",7:"липня",8:"серпня",
    9:"вересня",10:"жовтня",11:"листопада",12:"грудня"
}

HTML = r"""<!DOCTYPE html>
<html lang="uk">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Документи — Валютний контроль</title>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Sans:wght@300;400;500&display=swap" rel="stylesheet">
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#f5f3ef;--surface:#fff;--border:#d8d4cc;
  --text:#1a1916;--muted:#7a756c;--accent:#1a1916;
  --sans:'IBM Plex Sans',sans-serif;--mono:'IBM Plex Mono',monospace;
}
body{font-family:var(--sans);background:var(--bg);color:var(--text);min-height:100vh;display:flex;justify-content:center;padding:3rem 1rem}
.wrap{width:100%;max-width:580px}
h1{font-weight:300;font-size:26px;line-height:1.2;letter-spacing:-.02em;margin-bottom:.3rem}
h1 b{font-weight:500}
.sub{font-size:13px;color:var(--muted);margin-bottom:2rem}
.badge{font-family:var(--mono);font-size:10px;letter-spacing:.1em;text-transform:uppercase;color:var(--muted);border:1px solid var(--border);display:inline-block;padding:3px 8px;border-radius:2px;margin-bottom:.8rem}
.card{background:var(--surface);border:1px solid var(--border);border-radius:4px;padding:1.5rem;margin-bottom:1rem}
.sec{font-family:var(--mono);font-size:10px;letter-spacing:.1em;text-transform:uppercase;color:var(--muted);padding-bottom:8px;border-bottom:1px solid var(--border);margin-bottom:1.2rem}
.field{margin-bottom:1.1rem}
.field:last-child{margin-bottom:0}
.lbl{display:block;font-family:var(--mono);font-size:11px;letter-spacing:.07em;text-transform:uppercase;color:var(--muted);margin-bottom:5px}
input[type=text],input[type=number],textarea{width:100%;padding:9px 12px;border:1px solid var(--border);border-radius:3px;background:var(--bg);color:var(--text);font-family:var(--sans);font-size:14px;outline:none;transition:border-color .15s}
textarea{resize:vertical;min-height:65px;line-height:1.5}
input:focus,textarea:focus{border-color:var(--accent)}
input::placeholder,textarea::placeholder{color:var(--muted);opacity:.6}
.two{display:grid;grid-template-columns:1fr 1fr;gap:10px}
.hr{border:none;border-top:1px solid var(--border);margin:1.2rem 0}
.toggle-row{display:grid;grid-template-columns:1fr 1fr;border:1px solid var(--border);border-radius:3px;overflow:hidden}
.tbtn{padding:9px;font-family:var(--sans);font-size:14px;background:var(--bg);color:var(--muted);border:none;cursor:pointer;transition:background .15s,color .15s}
.tbtn:first-child{border-right:1px solid var(--border)}
.tbtn.on{background:var(--accent);color:#fff;font-weight:500}
.cur-row{display:flex;border:1px solid var(--border);border-radius:3px;overflow:hidden;margin-top:2px}
.cbtn{flex:1;padding:9px;font-family:var(--mono);font-size:13px;font-weight:500;background:var(--bg);color:var(--muted);border:none;cursor:pointer;transition:background .15s,color .15s}
.cbtn:first-child{border-right:1px solid var(--border)}
.cbtn.on{background:var(--accent);color:#fff}
.hidden{display:none}
.bal-extra{margin-top:.9rem;padding-top:.9rem;border-top:1px dashed var(--border)}
.amtrow{display:grid;grid-template-columns:1fr auto;gap:10px;align-items:flex-end}
.hint{font-size:11px;color:var(--muted);margin-top:4px;font-style:italic}
.api-badge{font-family:var(--mono);font-size:10px;padding:2px 7px;border-radius:2px;margin-left:6px;vertical-align:middle}
.api-on{background:#eef7ee;color:#2d6a4f;border:1px solid #c3e6cb}
.api-off{background:#fff8e6;color:#856404;border:1px solid #ffdda0}

/* Кнопки завантаження */
.dl-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:8px}
.btn-zip{width:100%;padding:12px;background:var(--accent);color:#fff;border:none;border-radius:3px;font-family:var(--sans);font-size:15px;font-weight:500;cursor:pointer;transition:opacity .15s}
.btn-zip:hover{opacity:.85}
.btn-zip:disabled{opacity:.5;cursor:default}
.btn-sep{width:100%;padding:10px;background:var(--surface);color:var(--text);border:1px solid var(--border);border-radius:3px;font-family:var(--sans);font-size:13px;cursor:pointer;transition:background .15s}
.btn-sep:hover{background:var(--bg)}
.btn-sep:disabled{opacity:.5;cursor:default}

.msg{margin-top:.8rem;padding:9px 12px;border-radius:3px;font-size:13px;display:none}
.msg.err{background:#fee;border:1px solid #fcc;color:#a00}
.msg.ok{background:#eef7ee;border:1px solid #c3e6cb;color:#2d6a4f}
footer{margin-top:1.5rem;text-align:center;font-family:var(--mono);font-size:11px;color:var(--muted)}
</style>
</head>
<body>
<div class="wrap">
  <div class="badge">АТ Універсал Банк · Валютний контроль</div>
  <h1>Генератор <b>банківських документів</b></h1>
  <p class="sub">Заповніть форму — отримайте обидва PDF одразу</p>

  <div class="card">
    <div class="sec">Дані клієнта</div>
    <div class="field">
      <span class="lbl">ПІБ клієнта</span>
      <input type="text" id="pib" placeholder="Іваненко Іван Іванович">
    </div>
    <div class="two">
      <div class="field">
        <span class="lbl">ІПН</span>
        <input type="text" id="ipn" placeholder="1234567890" maxlength="10">
      </div>
      <div class="field">
        <span class="lbl">Дата (порожньо = сьогодні)</span>
        <input type="text" id="date" placeholder="22.03.2026">
      </div>
    </div>
    <div class="field">
      <span class="lbl">Адреса клієнта</span>
      <input type="text" id="address" placeholder="Україна, обл. ..., вул. ..., буд. ...">
    </div>
  </div>

  <div class="card">
    <div class="sec">Довідка про залишок валюти</div>
    <div class="field">
      <span class="lbl">Залишок на рахунку</span>
      <div class="toggle-row">
        <button class="tbtn on" id="bal-no"  type="button">Відсутній</button>
        <button class="tbtn"    id="bal-yes" type="button">Є залишок</button>
      </div>
    </div>
    <div id="bal-extra" class="bal-extra hidden">
      <div class="amtrow">
        <div class="field" style="margin:0">
          <span class="lbl">Сума залишку</span>
          <input type="number" id="bal-amt" placeholder="1238.33" step="0.01" min="0">
        </div>
        <div>
          <span class="lbl">Валюта</span>
          <div class="cur-row">
            <button class="cbtn on" id="bcur-usd" type="button">USD</button>
            <button class="cbtn"    id="bcur-eur" type="button">EUR</button>
          </div>
        </div>
      </div>
    </div>
  </div>

  <div class="card">
    <div class="sec">Заява на купівлю валюти</div>

    <!-- Блок мети: залежить від наявності API ключа -->
    <div id="purpose-block-auto" class="__PURPOSE_BLOCK__">
      <div class="field">
        <span class="lbl">
          Мета купівлі (англійською)
          <span class="api-badge api-on">автопереклад увімкнено</span>
        </span>
        <textarea id="purpose-en" placeholder="Prepayment for goods according to contract nr 01/26 dd 16.03.2026, Invoice nr PI2026031201 dd 17.03.2026"></textarea>
        <p class="hint">Claude перекладе автоматично на українську</p>
      </div>
    </div>

    <div id="purpose-block-manual" class="__PURPOSE_BLOCK_MAN__">
      <div class="field">
        <span class="lbl">
          Мета купівлі (англійською)
          <span class="api-badge api-off">автопереклад вимкнено</span>
        </span>
        <textarea id="purpose-en-m" placeholder="Prepayment for goods according to contract nr 01/26 dd 16.03.2026, Invoice nr PI2026031201 dd 17.03.2026"></textarea>
      </div>
      <div class="field">
        <span class="lbl">Мета купівлі (українською)</span>
        <textarea id="purpose-uk" placeholder="Передоплата за товар згідно з договором № 01/26 від 16.03.2026, рахунком № PI2026031201 від 17.03.2026"></textarea>
      </div>
    </div>

    <div class="two">
      <div class="field">
        <span class="lbl">Сума купівлі</span>
        <input type="number" id="buy-amt" placeholder="5800.00" step="0.01" min="0">
      </div>
      <div class="field">
        <span class="lbl">Валюта купівлі</span>
        <div class="cur-row">
          <button class="cbtn on" id="bcur2-usd" type="button">USD</button>
          <button class="cbtn"    id="bcur2-eur" type="button">EUR</button>
        </div>
      </div>
    </div>
    <hr class="hr">
    <div class="field">
      <span class="lbl">IBAN — списання гривні (купівля)</span>
      <input type="text" id="iban1" placeholder="UA293220010000026006310115156" maxlength="29">
    </div>
    <div class="field">
      <span class="lbl">IBAN — зарахування валюти</span>
      <input type="text" id="iban2" placeholder="UA353220010000026000370076190" maxlength="29">
    </div>
    <div class="field">
      <span class="lbl">IBAN — списання комісії</span>
      <input type="text" id="iban3" placeholder="UA293220010000026006310115156" maxlength="29">
    </div>
  </div>

  <div class="msg err" id="err-msg"></div>
  <div class="msg ok"  id="ok-msg"></div>

  <!-- Три кнопки завантаження -->
  <div class="dl-grid">
    <button class="btn-sep" id="btn-dovidka" type="button">↓ Тільки довідка</button>
    <button class="btn-sep" id="btn-zayava"  type="button">↓ Тільки заява</button>
  </div>
  <button class="btn-zip" id="btn-both" type="button">↓ Завантажити обидва PDF (zip)</button>

  <footer style="margin-top:1.2rem">Для внутрішнього використання</footer>
</div>

<script>
(function(){
  var hasBalance = false;
  var balCur = 'USD', buyCur = 'USD';
  var hasApiKey = __HAS_API_KEY__;

  function q(id){ return document.getElementById(id); }

  function setActive(a, b, aIsOn){
    a.classList.toggle('on', aIsOn);
    b.classList.toggle('on', !aIsOn);
  }

  // Показуємо потрібний блок мети
  if(hasApiKey){
    q('purpose-block-auto').classList.remove('hidden');
    q('purpose-block-manual').classList.add('hidden');
  } else {
    q('purpose-block-auto').classList.add('hidden');
    q('purpose-block-manual').classList.remove('hidden');
  }

  q('bal-no').addEventListener('click', function(){
    hasBalance=false; setActive(q('bal-no'),q('bal-yes'),true);
    q('bal-extra').classList.add('hidden');
  });
  q('bal-yes').addEventListener('click', function(){
    hasBalance=true; setActive(q('bal-no'),q('bal-yes'),false);
    q('bal-extra').classList.remove('hidden');
  });
  q('bcur-usd').addEventListener('click',function(){ balCur='USD'; setActive(q('bcur-usd'),q('bcur-eur'),true); });
  q('bcur-eur').addEventListener('click',function(){ balCur='EUR'; setActive(q('bcur-usd'),q('bcur-eur'),false); });
  q('bcur2-usd').addEventListener('click',function(){ buyCur='USD'; setActive(q('bcur2-usd'),q('bcur2-eur'),true); });
  q('bcur2-eur').addEventListener('click',function(){ buyCur='EUR'; setActive(q('bcur2-usd'),q('bcur2-eur'),false); });

  function showMsg(id,text){
    ['err-msg','ok-msg'].forEach(function(x){q(x).style.display='none';});
    var el=q(id); el.textContent=text; el.style.display='block';
  }

  function getPurposeEn(){
    return hasApiKey ? q('purpose-en').value.trim() : q('purpose-en-m').value.trim();
  }
  function getPurposeUk(){
    return hasApiKey ? '' : q('purpose-uk').value.trim();
  }

  function validate(){
    if(!q('pib').value.trim()) return 'Заповніть ПІБ клієнта';
    if(q('ipn').value.trim().length!==10) return 'ІПН має містити 10 цифр';
    if(!q('address').value.trim()) return 'Заповніть адресу клієнта';
    if(!getPurposeEn()) return 'Вкажіть мету купівлі (англійською)';
    if(!hasApiKey && !getPurposeUk()) return 'Вкажіть мету купівлі (українською)';
    if(!q('buy-amt').value.trim()) return 'Вкажіть суму купівлі';
    if(q('iban1').value.trim().length<20) return 'Перевірте IBAN для списання гривні';
    if(q('iban2').value.trim().length<20) return 'Перевірте IBAN для зарахування валюти';
    if(q('iban3').value.trim().length<20) return 'Перевірте IBAN для списання комісії';
    if(hasBalance && !q('bal-amt').value.trim()) return 'Вкажіть суму залишку';
    return null;
  }

  function buildPayload(docType){
    return {
      pib:         q('pib').value.trim(),
      ipn:         q('ipn').value.trim(),
      address:     q('address').value.trim(),
      date:        q('date').value.trim(),
      purpose_en:  getPurposeEn(),
      purpose_uk:  getPurposeUk(),
      buy_amount:  q('buy-amt').value.trim(),
      buy_currency: buyCur,
      iban_debit:  q('iban1').value.trim(),
      iban_credit: q('iban2').value.trim(),
      iban_commission: q('iban3').value.trim(),
      has_balance: hasBalance,
      balance_amount:   q('bal-amt').value.trim(),
      balance_currency: balCur,
      doc_type: docType
    };
  }

  function setLoading(on){
    ['btn-both','btn-dovidka','btn-zayava'].forEach(function(id){
      q(id).disabled = on;
    });
  }

  function doDownload(docType){
    var err=validate(); if(err){showMsg('err-msg',err);return;}
    setLoading(true);
    var label = docType==='both'?'Генеруємо обидва документи...':
                docType==='dovidka'?'Генеруємо довідку...':'Генеруємо заяву...';
    showMsg('ok-msg', label);

    fetch('/generate', {
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify(buildPayload(docType))
    })
    .then(function(resp){
      if(!resp.ok) return resp.json().then(function(d){throw new Error(d.error||'Помилка сервера');});
      return resp.blob();
    })
    .then(function(blob){
      var pib0 = q('pib').value.trim().split(' ')[0];
      var ext  = docType==='both' ? '.zip' : '.pdf';
      var pre  = docType==='both'  ? 'Dokumenty_' :
                 docType==='dovidka'? 'Dovidka_' : 'Zayava_';
      var url=URL.createObjectURL(blob);
      var a=document.createElement('a');
      a.href=url; a.download=pre+pib0+ext;
      document.body.appendChild(a); a.click();
      document.body.removeChild(a); URL.revokeObjectURL(url);
      showMsg('ok-msg','Готово! Документ(и) завантажено.');
    })
    .catch(function(e){ showMsg('err-msg', e.message||'Помилка зʼєднання'); })
    .finally(function(){ setLoading(false); });
  }

  q('btn-both').addEventListener('click',    function(){ doDownload('both'); });
  q('btn-dovidka').addEventListener('click', function(){ doDownload('dovidka'); });
  q('btn-zayava').addEventListener('click',  function(){ doDownload('zayava'); });
})();
</script>
</body>
</html>"""


def translate_purpose(purpose_en):
    try:
        api_key = os.environ.get("ANTHROPIC_API_KEY","")
        if not api_key:
            return purpose_en
        payload = json.dumps({
            "model": "claude-sonnet-4-20250514",
            "max_tokens": 500,
            "system": TRANSLATION_SYSTEM,
            "messages": [{"role":"user","content":purpose_en}]
        }).encode("utf-8")
        req = urllib.request.Request(
            "https://api.anthropic.com/v1/messages", data=payload,
            headers={"Content-Type":"application/json",
                     "anthropic-version":"2023-06-01",
                     "x-api-key":api_key},
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read())["content"][0]["text"].strip()
    except Exception:
        return purpose_en


def register_fonts():
    for reg,bold in [
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
         "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        ("/usr/share/fonts/dejavu/DejaVuSans.ttf",
         "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf"),
    ]:
        if os.path.exists(reg):
            try:
                pdfmetrics.registerFont(TTFont("AppFont",reg))
                pdfmetrics.registerFont(TTFont("AppFont-Bold",bold))
                return "AppFont","AppFont-Bold"
            except Exception:
                pass
    return "Helvetica","Helvetica-Bold"


def get_date_str(date_str):
    if date_str and date_str.strip():
        return date_str.strip()
    return date.today().strftime("%d.%m.%Y")


def format_balance(amount, currency):
    try:
        val=float(amount)
        if val>0 and currency:
            fmt=f"{val:,.2f}".replace(",", " ").replace(".", ",")
            return f"{fmt} {currency} в АТ УніверсалБанк"
    except Exception:
        pass
    return None


def build_dovidka_pdf(pib, ipn, doc_date, balance_line):
    fn,fb=register_fonts()
    p1=(f"1. на поточних рахунках - {balance_line};"
        if balance_line else "1. на поточних рахунках - відсутні;")
    sc =ParagraphStyle("c", fontName=fn,fontSize=11,leading=16,alignment=TA_CENTER)
    scb=ParagraphStyle("cb",fontName=fb,fontSize=11,leading=16,alignment=TA_CENTER)
    sj =ParagraphStyle("j", fontName=fn,fontSize=11,leading=16,alignment=TA_JUSTIFY)
    sl =ParagraphStyle("l", fontName=fn,fontSize=11,leading=16,alignment=TA_LEFT)
    ss =ParagraphStyle("s", fontName=fn,fontSize=10,leading=14,alignment=TA_LEFT)
    ssr=ParagraphStyle("sr",fontName=fn,fontSize=10,leading=14,alignment=TA_RIGHT)
    buf=io.BytesIO()
    doc=SimpleDocTemplate(buf,pagesize=A4,
                          rightMargin=20*mm,leftMargin=20*mm,
                          topMargin=20*mm,bottomMargin=20*mm)
    s=[]
    s.append(Paragraph("До уваги Управління з валютного контролю",sc))
    s.append(Spacer(1,6*mm))
    s.append(Paragraph("ДОВІДКА",scb))
    s.append(Spacer(1,5*mm))
    s.append(Paragraph(
        f"На додаток до заяви на купівлю іноземної валюти надаємо наступну інформацію "
        f"щодо наявності іноземної валюти на рахунках ФОП {pib} "
        f"{ipn} станом на початок операційного дня в уповноважених банках України "
        f"(суми необхідно вказати в розрізі валют та назв уповноважених банків):",sj))
    s.append(Spacer(1,4*mm))
    s.append(Paragraph(p1,sl))
    s.append(Spacer(1,2*mm))
    s.append(Paragraph("2. на вкладних (депозитних) рахунках - залишки в іноземних валютах відсутні;",sl))
    s.append(Spacer(1,2*mm))
    s.append(Paragraph(
        '3. продана клієнтом іноземна валюта за першою частиною валютних операцій на '
        'умовах "своп" за незавершеними угодами з банками - відсутні.',sj))
    s.append(Spacer(1,2*mm))
    s.append(Paragraph(
        "4. виключення відповідно до пункту 12-15 Постанови НБУ № 18 від 24.02.2022: "
        "кошти, що знаходяться під заставою/ накладено арешт/ у неплатоспроможних "
        "банках/ отримані за договорами комісії/ були куплені та не використані в "
        "установлений законодавством строк/ інші: відсутні",sj))
    s.append(Spacer(1,8*mm))
    s.append(Paragraph("Підпис",sl))
    s.append(Spacer(1,2*mm))
    s.append(Paragraph(f"ФОП {pib}",sl))
    s.append(Spacer(1,2*mm))
    s.append(Paragraph("Електронний підпис",sl))
    s.append(Spacer(1,4*mm))
    tbl=Table([[Paragraph(doc_date,ss),Paragraph(pib,ssr)]],colWidths=["40%","60%"])
    tbl.setStyle(TableStyle([
        ("VALIGN",(0,0),(-1,-1),"TOP"),
        ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),
        ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0),
    ]))
    s.append(tbl)
    doc.build(s)
    buf.seek(0)
    return buf.read()


def build_valyuta_pdf(params):
    sys.path.insert(0,BASE_DIR)
    import fill_valyuta as fv
    tmpdir=tempfile.mkdtemp()
    out_docx=os.path.join(tmpdir,"v.docx")
    out_pdf =os.path.join(tmpdir,"v.pdf")
    fv.fill_template(params,TEMPLATE_VALYUTA,out_docx)
    fv.docx_to_pdf(out_docx,out_pdf)
    with open(out_pdf,"rb") as f: data=f.read()
    import shutil; shutil.rmtree(tmpdir)
    return data


@app.route("/")
def index():
    html = HTML
    if HAS_API_KEY:
        html = html.replace("__PURPOSE_BLOCK__","").replace("__PURPOSE_BLOCK_MAN__","hidden")
    else:
        html = html.replace("__PURPOSE_BLOCK__","hidden").replace("__PURPOSE_BLOCK_MAN__","")
    html = html.replace("__HAS_API_KEY__", "true" if HAS_API_KEY else "false")
    return render_template_string(html)


@app.route("/generate", methods=["POST"])
def generate():
    try:
        data=request.get_json(force=True)
        pib             =data.get("pib","").strip()
        ipn             =data.get("ipn","").strip()
        address         =data.get("address","").strip()
        date_str        =data.get("date","").strip()
        purpose_en      =data.get("purpose_en","").strip()
        purpose_uk_in   =data.get("purpose_uk","").strip()
        buy_amount      =data.get("buy_amount","0")
        buy_currency    =data.get("buy_currency","USD")
        iban_debit      =data.get("iban_debit","").strip()
        iban_credit     =data.get("iban_credit","").strip()
        iban_commission =data.get("iban_commission","").strip()
        has_balance     =bool(data.get("has_balance",False))
        bal_amount      =data.get("balance_amount","0")
        bal_currency    =data.get("balance_currency","USD")
        doc_type        =data.get("doc_type","both")

        # Визначаємо переклад
        if HAS_API_KEY:
            purpose_uk = translate_purpose(purpose_en)
        else:
            purpose_uk = purpose_uk_in or purpose_en

        doc_date = get_date_str(date_str)
        bal_line = format_balance(bal_amount, bal_currency) if has_balance else None

        valyuta_params = {
            "pib":pib,"ipn":ipn,"address":address,"date":date_str,
            "purpose_uk":purpose_uk,"amount":buy_amount,"currency":buy_currency,
            "iban_debit":iban_debit,"iban_credit":iban_credit,"iban_commission":iban_commission,
        }
        safe = (pib.split()[0] if pib else "doc")
        d    = doc_date.replace(".","_")

        if doc_type == "dovidka":
            pdf = build_dovidka_pdf(pib, ipn, doc_date, bal_line)
            return send_file(io.BytesIO(pdf), mimetype="application/pdf",
                             as_attachment=True, download_name=f"Dovidka_{safe}_{d}.pdf")

        elif doc_type == "zayava":
            pdf = build_valyuta_pdf(valyuta_params)
            return send_file(io.BytesIO(pdf), mimetype="application/pdf",
                             as_attachment=True, download_name=f"Zayava_{safe}_{d}.pdf")

        else:  # both
            dov = build_dovidka_pdf(pib, ipn, doc_date, bal_line)
            zay = build_valyuta_pdf(valyuta_params)
            zip_buf=io.BytesIO()
            with zipfile.ZipFile(zip_buf,"w",zipfile.ZIP_DEFLATED) as zf:
                zf.writestr(f"Dovidka_{safe}_{d}.pdf", dov)
                zf.writestr(f"Zayava_{safe}_{d}.pdf",  zay)
            zip_buf.seek(0)
            return send_file(zip_buf, mimetype="application/zip",
                             as_attachment=True, download_name=f"Dokumenty_{safe}_{d}.zip")

    except Exception as e:
        return jsonify({"error":str(e)}),500


if __name__=="__main__":
    port=int(os.environ.get("PORT",5000))
    print(f"\n Додаток: http://localhost:{port}")
    print(f" API ключ: {'знайдено' if HAS_API_KEY else 'відсутній — ручний режим'}\n")
    app.run(host="0.0.0.0",port=port)
