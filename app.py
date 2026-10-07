"""
Invoice -> Excel.  Run:  python app.py   then open http://127.0.0.1:5000
Upload KSBCL invoice PDFs, check the numbers on the review screen, download the workbook.
"""
import os, io, uuid, json, tempfile, datetime as dt
from flask import Flask, request, render_template_string, send_file, jsonify, abort
import engine

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024
SESS = {}   # id -> parsed invoices (kept in memory only)

PAGE = """<!doctype html><html><head><meta charset="utf-8"><title>Invoice to Excel</title>
<style>
 body{font-family:Cambria,Georgia,serif;margin:0;background:#f4f6fa;color:#1b2540}
 header{background:#002060;color:#fff;padding:14px 28px;font-size:20px}
 main{max-width:1200px;margin:22px auto;padding:0 18px}
 .card{background:#fff;border:1px solid #d5dbe8;border-radius:8px;padding:18px 22px;margin-bottom:18px}
 label{display:block;margin:10px 0 3px;font-weight:bold;font-size:14px}
 input[type=text],input[type=number],input[type=date]{padding:6px 8px;border:1px solid #b9c2d6;border-radius:5px;font:inherit}
 button{background:#002060;color:#fff;border:0;border-radius:6px;padding:10px 22px;font:inherit;font-size:16px;cursor:pointer}
 button.sec{background:#e6eaf4;color:#002060;padding:5px 12px;font-size:13px}
 table{border-collapse:collapse;width:100%;font-size:13px}
 th{background:#002060;color:#fff;padding:5px;text-align:left}
 td{border-bottom:1px solid #e3e7f0;padding:2px}
 td input{width:100%;box-sizing:border-box;border:1px solid transparent;background:transparent;font:inherit;padding:3px}
 td input:focus{border-color:#7a8bb5;background:#fff}
 tr.flag td{background:#fff4d6} tr.bad td{background:#ffe0e0}
 .warn{background:#ffe0e0;border:1px solid #e09a9a;border-radius:6px;padding:8px 12px;margin:8px 0}
 .ok{background:#e1f5e4;border:1px solid #95cf9d;border-radius:6px;padding:8px 12px;margin:8px 0}
 .note{font-size:12px;color:#a05a00} .row{display:flex;gap:22px;flex-wrap:wrap}
 .num{text-align:right}
</style></head><body><header>Invoice &rarr; Excel &nbsp;<small>(Purchase + Sales + Summary)</small></header><main>
{% if not invoices %}
<div class="card"><form method="post" action="/parse" enctype="multipart/form-data">
 <label>Invoice PDFs (scanned KSBCL invoices, permit page + bill pages)</label>
 <input type="file" name="pdfs" accept="application/pdf" multiple required>
 <div class="row">
  <div><label>First sales invoice no.</label><input type="number" name="sales_start" value="{{ start }}"></div>
  <div><label>Sales mark-up %</label><input type="number" step="0.1" name="markup" value="10"></div>
  <div><label>First sale: days after first purchase</label><input type="number" name="offset" value="2"></div>
  <div><label>Gap between sales invoices (days)</label><input type="number" name="gap" value="3"></div>
 </div>
 <p><button>Read invoices</button></p>
 <p class="note">Reading takes ~15-30 s per PDF (OCR). Nothing leaves your computer.</p>
</form></div>
{% else %}
<form id="f" method="post" action="/build">
<input type="hidden" name="sid" value="{{ sid }}"><input type="hidden" name="payload" id="payload">
<div class="card"><b>Settings</b>
 <div class="row">
  <div><label>First sales invoice no.</label><input type="number" id="sales_start" value="{{ cfg.sales_start }}"></div>
  <div><label>Mark-up %</label><input type="number" step="0.1" id="markup" value="{{ cfg.markup }}"></div>
  <div><label>First sale: days after first purchase</label><input type="number" id="offset" value="{{ cfg.offset }}"></div>
  <div><label>Gap (days)</label><input type="number" id="gap" value="{{ cfg.gap }}"></div>
 </div></div>
<div id="invs"></div>
<p><button type="button" onclick="go()">Create Excel workbook</button></p>
</form>
<script>
const DATA={{ data|safe }};
const fmt=x=>Number(x).toLocaleString('en-IN',{minimumFractionDigits:2,maximumFractionDigits:2});
function rowHtml(it){
 const cls=(it.flags&&it.flags.length)?(it.flags.join(' ').match(/not read|check qty/)?'bad':'flag'):'';
 return `<tr class="${cls}"><td style="width:34px">${it.sl}</td>
 <td><input value="${(it.name||'').replace(/"/g,'&quot;')}" data-k="name"></td>
 <td style="width:60px"><input class="num" value="${it.cb}" data-k="cb"></td>
 <td style="width:60px"><input class="num" value="${it.btls}" data-k="btls"></td>
 <td style="width:100px"><input class="num" value="${it.rate}" data-k="rate"></td>
 <td style="width:110px"><input class="num" value="${it.amt}" data-k="amt"></td>
 <td style="width:30px"><button type="button" class="sec" onclick="delRow(this)">x</button></td>
 <td class="note" style="width:230px">${(it.flags||[]).join('; ')}</td></tr>`;}
function render(){
 const box=document.getElementById('invs'); box.innerHTML='';
 DATA.forEach((inv,i)=>{
  const d=document.createElement('div'); d.className='card'; d.dataset.i=i;
  d.innerHTML=`<b>${inv.invoice_no||'(no. not read)'}</b> &nbsp; date <input type="date" value="${inv.date}" data-k="date"> &nbsp; <small>${inv.file}</small>
   <div class="chk"></div>
   <table><thead><tr><th>#</th><th>Item</th><th>CB</th><th>Btls</th><th>Rate/CB</th><th>Amount</th><th></th><th>Check</th></tr></thead>
   <tbody>${inv.items.map(rowHtml).join('')}</tbody></table>
   <p><button type="button" class="sec" onclick="addRow(${i})">+ add row</button></p>`;
  box.appendChild(d); check(d,inv);
  d.addEventListener('input',()=>{collect(d,inv);check(d,inv)});
 });}
function collect(d,inv){
 inv.date=d.querySelector('[data-k=date]').value;
 inv.items=[...d.querySelectorAll('tbody tr')].map((tr,k)=>{const o={sl:k+1,flags:[]};
  tr.querySelectorAll('input').forEach(x=>{const key=x.dataset.k;o[key]=key==='name'?x.value:parseFloat(x.value)||0});return o;});}
function check(d,inv){
 const sum=inv.items.reduce((a,b)=>a+(parseFloat(b.amt)||0),0), cb=inv.items.reduce((a,b)=>a+(parseInt(b.cb)||0),0);
 let msgs=[];
 if(inv.total_read!=null){const diff=Math.round((sum-inv.total_read)*100)/100;
   msgs.push(Math.abs(diff)<0.06?`<div class="ok">Items add up to ${fmt(sum)} = invoice total ${fmt(inv.total_read)} &#10003;</div>`:
   `<div class="warn">Items add up to ${fmt(sum)} but the invoice total is ${fmt(inv.total_read)} (difference ${fmt(diff)}). Check the highlighted rows against the paper invoice.</div>`);}
 else msgs.push(`<div class="warn">Invoice total could not be read - please verify by eye. Items add up to ${fmt(sum)}.</div>`);
 (inv.warnings||[]).filter(w=>!w.startsWith('Line amounts')).forEach(w=>msgs.push(`<div class="warn">${w}</div>`));
 d.querySelector('.chk').innerHTML=msgs.join('');}
function delRow(b){const d=b.closest('.card'),i=+d.dataset.i;b.closest('tr').remove();collect(d,DATA[i]);check(d,DATA[i]);}
function addRow(i){const d=document.querySelector(`.card[data-i="${i}"]`);collect(d,DATA[i]);
 DATA[i].items.push({sl:DATA[i].items.length+1,name:'',cb:0,btls:0,rate:0,amt:0,flags:[]});const keep=DATA.map(x=>x);render();}
function go(){document.querySelectorAll('.card[data-i]').forEach(d=>collect(d,DATA[+d.dataset.i]));
 const cfg={sales_start:+sales_start.value,markup:(+markup.value)/100,offset:+offset.value,gap:+gap.value};
 payload.value=JSON.stringify({cfg,invoices:DATA});document.getElementById('f').submit();}
render();
</script>
{% endif %}
</main></body></html>"""

@app.get("/")
def home():
    return render_template_string(PAGE, invoices=None, start=911)

@app.post("/parse")
def parse():
    files = request.files.getlist("pdfs")
    if not files: abort(400)
    cat = engine.load_catalog()
    tmp = tempfile.mkdtemp()
    invs = []
    for f in files:
        p = os.path.join(tmp, os.path.basename(f.filename) or f"{uuid.uuid4().hex}.pdf")
        f.save(p)
        invs.append(engine.read_invoice(p, catalog=cat))
    invs.sort(key=lambda v: (v["date"], v["invoice_no"]))
    sid = uuid.uuid4().hex
    cfg = dict(sales_start=int(request.form.get("sales_start") or 911), markup=float(request.form.get("markup") or 10),
               offset=int(request.form.get("offset") or 2), gap=int(request.form.get("gap") or 3))
    SESS[sid] = invs
    return render_template_string(PAGE, invoices=True, sid=sid, cfg=cfg, data=json.dumps(invs, default=str).replace("</", "<\\/"))

@app.post("/build")
def build():
    pl = json.loads(request.form["payload"])
    cfg, invs = pl["cfg"], pl["invoices"]
    for inv in invs:
        for it in inv["items"]:
            for k in ("cb", "btls"): it[k] = int(it.get(k) or 0)
            for k in ("rate", "amt"): it[k] = engine.money(it.get(k) or 0)
        inv["items"] = [it for it in inv["items"] if it["name"].strip()]
    out = os.path.join(tempfile.mkdtemp(), "Liquor_Purchase_Sales.xlsx")
    info = engine.build_workbook(invs, out, sales_start=cfg["sales_start"], markup=cfg["markup"],
                                 first_sale_offset=cfg["offset"], sale_gap_days=cfg["gap"])
    engine.learn([it for inv in invs for it in inv["items"] if it["rate"]])
    name = f"Hotel_Shanth_{info['month'].replace(' ', '-')}.xlsx"
    return send_file(out, as_attachment=True, download_name=name,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
