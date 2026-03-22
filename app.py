from flask import Flask, request, send_file, render_template_string, jsonify
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY, TA_RIGHT
from datetime import date
import io, os, sys, tempfile, json

app = Flask(__name__)  # v3.0 — no LibreOffice
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_VALYUTA = os.path.join(BASE_DIR, "template_valyuta.docx")

HTML = r"""<!DOCTYPE html>
<html lang="uk">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Документи — Валютний контроль</title>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Sans:wght@300;400;500&display=swap" rel="stylesheet">
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{--bg:#f5f3ef;--surface:#fff;--border:#d8d4cc;--text:#1a1916;--muted:#7a756c;--accent:#1a1916;--sans:'IBM Plex Sans',sans-serif;--mono:'IBM Plex Mono',monospace}
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
.hidden{display:none!important}
.extra{margin-top:.9rem;padding-top:.9rem;border-top:1px dashed var(--border)}
.amtrow{display:grid;grid-template-columns:1fr auto;gap:10px;align-items:flex-end}
.dl-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:8px}
.btn-main{width:100%;padding:12px;background:var(--accent);color:#fff;border:none;border-radius:3px;font-family:var(--sans);font-size:15px;font-weight:500;cursor:pointer;transition:opacity .15s}
.btn-main:hover{opacity:.85}
.btn-main:disabled{opacity:.5;cursor:default}
.btn-sec{width:100%;padding:10px;background:var(--surface);color:var(--text);border:1px solid var(--border);border-radius:3px;font-family:var(--sans);font-size:13px;cursor:pointer;transition:background .15s}
.btn-sec:hover{background:var(--bg)}
.btn-sec:disabled{opacity:.5;cursor:default}
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
  <p class="sub">Заповніть форму — завантажте один або обидва документи</p>

  <!-- Спільні дані -->
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
  </div>

  <!-- Довідка -->
  <div class="card">
    <div class="sec">Довідка про залишок валюти</div>
    <div class="field">
      <span class="lbl">Залишок на рахунку</span>
      <div class="toggle-row">
        <button class="tbtn on" id="bal-no"  type="button">Відсутній</button>
        <button class="tbtn"    id="bal-yes" type="button">Є залишок</button>
      </div>
    </div>
    <div id="bal-extra" class="extra hidden">
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

  <!-- Заява -->
  <div class="card">
    <div class="sec">Заява на купівлю валюти</div>
    <div class="field">
      <span class="lbl">Адреса клієнта</span>
      <input type="text" id="address" placeholder="Україна, обл. ..., вул. ..., буд. ...">
    </div>
    <div class="field">
      <span class="lbl">Мета купівлі (українською)</span>
      <textarea id="purpose" placeholder="Передоплата за товар згідно з договором № 01/26 від 16.03.2026, рахунком № PI2026031201 від 17.03.2026"></textarea>
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

  <div class="dl-grid">
    <button class="btn-sec" id="btn-dovidka" type="button">↓ Тільки довідка</button>
    <button class="btn-sec" id="btn-zayava"  type="button">↓ Тільки заява</button>
  </div>
  <button class="btn-main" id="btn-both" type="button">↓ Завантажити обидва документи</button>

  <footer>Для внутрішнього використання</footer>
</div>

<script>
(function(){
  var hasBalance=false, balCur='USD', buyCur='USD';
  function q(id){return document.getElementById(id)}
  function on(a,b,aOn){a.classList.toggle('on',aOn);b.classList.toggle('on',!aOn)}

  q('bal-no').addEventListener('click',function(){
    hasBalance=false; on(q('bal-no'),q('bal-yes'),true);
    q('bal-extra').classList.add('hidden');
  });
  q('bal-yes').addEventListener('click',function(){
    hasBalance=true; on(q('bal-no'),q('bal-yes'),false);
    q('bal-extra').classList.remove('hidden');
  });
  q('bcur-usd').addEventListener('click',function(){balCur='USD';on(q('bcur-usd'),q('bcur-eur'),true)});
  q('bcur-eur').addEventListener('click',function(){balCur='EUR';on(q('bcur-usd'),q('bcur-eur'),false)});
  q('bcur2-usd').addEventListener('click',function(){buyCur='USD';on(q('bcur2-usd'),q('bcur2-eur'),true)});
  q('bcur2-eur').addEventListener('click',function(){buyCur='EUR';on(q('bcur2-usd'),q('bcur2-eur'),false)});

  function showMsg(id,txt){
    ['err-msg','ok-msg'].forEach(function(x){q(x).style.display='none'});
    var el=q(id); el.textContent=txt; el.style.display='block';
  }

  function validateDovidka(){
    if(!q('pib').value.trim()) return 'Заповніть ПІБ клієнта';
    if(q('ipn').value.trim().length!==10) return 'ІПН має містити 10 цифр';
    if(hasBalance && !q('bal-amt').value.trim()) return 'Вкажіть суму залишку';
    return null;
  }
  function validateZayava(){
    if(!q('pib').value.trim()) return 'Заповніть ПІБ клієнта';
    if(q('ipn').value.trim().length!==10) return 'ІПН має містити 10 цифр';
    if(!q('address').value.trim()) return 'Заповніть адресу клієнта';
    if(!q('purpose').value.trim()) return 'Вкажіть мету купівлі';
    if(!q('buy-amt').value.trim()) return 'Вкажіть суму купівлі';
    if(q('iban1').value.trim().length<20) return 'Перевірте IBAN для списання гривні';
    if(q('iban2').value.trim().length<20) return 'Перевірте IBAN для зарахування валюти';
    if(q('iban3').value.trim().length<20) return 'Перевірте IBAN для списання комісії';
    return null;
  }

  function payload(){
    return {
      pib:q('pib').value.trim(), ipn:q('ipn').value.trim(),
      date:q('date').value.trim(), address:q('address').value.trim(),
      purpose_uk:q('purpose').value.trim(),
      buy_amount:q('buy-amt').value.trim(), buy_currency:buyCur,
      iban_debit:q('iban1').value.trim(), iban_credit:q('iban2').value.trim(),
      iban_commission:q('iban3').value.trim(),
      has_balance:hasBalance, balance_amount:q('bal-amt').value.trim(),
      balance_currency:balCur
    };
  }

  function setLoading(on){
    ['btn-both','btn-dovidka','btn-zayava'].forEach(function(id){q(id).disabled=on});
  }

  function downloadPdf(url, filename){
    return fetch(url,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload())})
      .then(function(r){
        if(!r.ok) return r.json().then(function(d){throw new Error(d.error||'Помилка сервера')});
        return r.blob();
      })
      .then(function(blob){
        var a=document.createElement('a');
        a.href=URL.createObjectURL(blob); a.download=filename;
        document.body.appendChild(a); a.click();
        document.body.removeChild(a);
      });
  }

  function pib0(){return q('pib').value.trim().split(' ')[0]||'doc'}

  q('btn-dovidka').addEventListener('click',function(){
    var err=validateDovidka(); if(err){showMsg('err-msg',err);return;}
    setLoading(true); showMsg('ok-msg','Генеруємо довідку...');
    downloadPdf('/gen/dovidka','Dovidka_'+pib0()+'.pdf')
      .then(function(){showMsg('ok-msg','Довідку завантажено.')})
      .catch(function(e){showMsg('err-msg',e.message)})
      .finally(function(){setLoading(false)});
  });

  q('btn-zayava').addEventListener('click',function(){
    var err=validateZayava(); if(err){showMsg('err-msg',err);return;}
    setLoading(true); showMsg('ok-msg','Генеруємо заяву...');
    downloadPdf('/gen/zayava','Zayava_'+pib0()+'.docx')
      .then(function(){showMsg('ok-msg','Заяву завантажено (Word .docx).')})
      .catch(function(e){showMsg('err-msg',e.message)})
      .finally(function(){setLoading(false)});
  });

  q('btn-both').addEventListener('click',function(){
    var err=validateDovidka()||validateZayava(); if(err){showMsg('err-msg',err);return;}
    setLoading(true); showMsg('ok-msg','Генеруємо довідку...');
    var p=pib0();
    downloadPdf('/gen/dovidka','Dovidka_'+p+'.pdf')
      .then(function(){
        showMsg('ok-msg','Генеруємо заяву...');
        return downloadPdf('/gen/zayava','Zayava_'+p+'.docx');
      })
      .then(function(){showMsg('ok-msg','Обидва документи завантажено.')})
      .catch(function(e){showMsg('err-msg',e.message)})
      .finally(function(){setLoading(false)});
  });
})();
</script>
</body>
</html>"""


def register_fonts():
    for reg,bold in [
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
         "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        ("/usr/share/fonts/dejavu/DejaVuSans.ttf",
         "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf"),
    ]:
        if os.path.exists(reg):
            try:
                pdfmetrics.registerFont(TTFont("F",reg))
                pdfmetrics.registerFont(TTFont("FB",bold))
                return "F","FB"
            except Exception:
                pass
    return "Helvetica","Helvetica-Bold"


def get_date_str(d):
    return d.strip() if d and d.strip() else date.today().strftime("%d.%m.%Y")


def format_balance(amount, currency):
    try:
        v=float(amount)
        if v>0 and currency:
            return f"{v:,.2f}".replace(",","·").replace(".",",").replace("·"," ")+" "+currency+" в АТ УніверсалБанк"
    except Exception:
        pass
    return None


def build_dovidka(pib, ipn, doc_date, balance_line):
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


def build_zayava(params):
    sys.path.insert(0,BASE_DIR)
    import fill_valyuta as fv
    return fv.build_zayava_docx(params)


def parse_request():
    d=request.get_json(force=True)
    return {
        "pib":            d.get("pib","").strip(),
        "ipn":            d.get("ipn","").strip(),
        "date":           d.get("date","").strip(),
        "address":        d.get("address","").strip(),
        "purpose_uk":     d.get("purpose_uk","").strip(),
        "buy_amount":     d.get("buy_amount","0"),
        "buy_currency":   d.get("buy_currency","USD"),
        "iban_debit":     d.get("iban_debit","").strip(),
        "iban_credit":    d.get("iban_credit","").strip(),
        "iban_commission":d.get("iban_commission","").strip(),
        "has_balance":    bool(d.get("has_balance",False)),
        "balance_amount": d.get("balance_amount","0"),
        "balance_currency":d.get("balance_currency","USD"),
    }


@app.route("/health")
def health():
    return jsonify({"status":"ok","version":"3.0","engine":"python-docx"})

@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/gen/dovidka", methods=["POST"])
def gen_dovidka():
    try:
        p=parse_request()
        doc_date=get_date_str(p["date"])
        bal_line=format_balance(p["balance_amount"],p["balance_currency"]) if p["has_balance"] else None
        pdf=build_dovidka(p["pib"],p["ipn"],doc_date,bal_line)
        safe=(p["pib"].split()[0] if p["pib"] else "doc")
        return send_file(io.BytesIO(pdf),mimetype="application/pdf",
                         as_attachment=True,
                         download_name=f"Dovidka_{safe}_{doc_date.replace('.','_')}.pdf")
    except Exception as e:
        return jsonify({"error":str(e)}),500


@app.route("/gen/zayava", methods=["POST"])
def gen_zayava():
    try:
        p=parse_request()
        doc_date=get_date_str(p["date"])
        params={
            "pib":p["pib"],"ipn":p["ipn"],"address":p["address"],
            "date":p["date"],"purpose_uk":p["purpose_uk"],
            "amount":p["buy_amount"],"currency":p["buy_currency"],
            "iban_debit":p["iban_debit"],"iban_credit":p["iban_credit"],
            "iban_commission":p["iban_commission"],
        }
        pdf=build_zayava(params)
        safe=(p["pib"].split()[0] if p["pib"] else "doc")
        return send_file(io.BytesIO(pdf),
                         mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                         as_attachment=True,
                         download_name=f"Zayava_{safe}_{doc_date.replace('.','_')}.docx")
    except Exception as e:
        return jsonify({"error":str(e)}),500


if __name__=="__main__":
    port=int(os.environ.get("PORT",5000))
    print(f"\n Додаток: http://localhost:{port}\n")
    app.run(host="0.0.0.0",port=port)
