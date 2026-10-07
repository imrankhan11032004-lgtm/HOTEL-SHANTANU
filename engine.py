"""
engine.py - KSBCL invoice PDFs  ->  Purchase / Sales / Summary workbook
(Hotel Shanth / Anuradha Kagod layout).

 1. read_invoice(pdf)    OCR a scanned KSBCL "Bill of Invoice" and parse the line items
 2. build_workbook(...)  write Summary + Purchase <Month> + Sales <Month> using the
                         formatting of template.xlsx (your original workbook)
"""
import os, re, json, glob, math, shutil, subprocess, tempfile, datetime as dt
from copy import copy
from collections import Counter
from decimal import Decimal, ROUND_HALF_UP
from difflib import SequenceMatcher

import openpyxl
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "template.xlsx")
LEARNED = os.path.join(HERE, "learned_catalog.json")

SIZES = {60, 90, 180, 275, 330, 375, 500, 650, 700, 750, 1000}
DEFAULT_PACK = {180: 48, 750: 12, 650: 12, 330: 24, 275: 24, 500: 24, 375: 24, 90: 96, 60: 96, 1000: 9}
BAD_LINE = re.compile(r"Invoice No|PartyID|Party Name|KSBCL|Sagar|Soraba|Bill Of|Page \d|Corporation|"
                      r"GSTIN|FSSAI|Bengal|\bBlock\b|Karnataka|Shantinagar|Building|Permit No|Valid Date|"
                      r"Assessment|District|Taluk|\bRange\b|P\.A\.N|Item Name|Amount|Rate for|AROED|Sy\.No|rporation|Limited|BNTC|BMTC|Beverages Corp|Shant|Digitally signed|Signature valid", re.I)

def D(x): return Decimal(str(x))
def money(x): return float(D(x).quantize(Decimal("0.01"), ROUND_HALF_UP))

# ---------------------------------------------------------------- catalogue ---
def _key(s): return re.sub(r"[^A-Z0-9]", "", s.upper())

def _fix_size(n):
    if n in SIZES: return n
    t = str(n)
    for k in range(len(t)):
        v = int(t[:k] + t[k + 1:] or 0)
        if v in SIZES: return v
    return None

def _size_of(name):
    m = re.search(r"(\d{2,4})\s*ML", name, re.I)
    return _fix_size(int(m.group(1))) if m else None

def load_catalog():
    """name -> set(rates). Seeded from template.xlsx, extended by learned_catalog.json"""
    cat = {}
    wb = openpyxl.load_workbook(TEMPLATE, data_only=False)
    ws = [w for w in wb if w.title.lower().startswith("purchase")][0]
    for c0 in range(2, ws.max_column, 7):
        for r in range(14, ws.max_row + 1):
            n, rate = ws.cell(r, c0 + 1).value, ws.cell(r, c0 + 4).value
            if isinstance(n, str) and n != "Total" and isinstance(rate, (int, float)):
                cat.setdefault(n, set()).add(float(rate))
    if os.path.exists(LEARNED):
        for n, rates in json.load(open(LEARNED)).items():
            cat.setdefault(n, set()).update(rates)
    return cat

def learn(items):
    data = json.load(open(LEARNED)) if os.path.exists(LEARNED) else {}
    for it in items:
        data.setdefault(it["name"], [])
        if it["rate"] not in data[it["name"]]: data[it["name"]].append(it["rate"])
    json.dump(data, open(LEARNED, "w"), indent=1)

def smart_title(s):
    out = []
    for w in s.split():
        if any(ch.isdigit() for ch in w) or w.isupper() and len(w) <= 3 and w in ("PET", "CAN", "XXX", "DSP", "TI", "UB"):
            out.append(w)
        else:
            out.append("-".join(p.capitalize() if not p.isupper() or len(p) > 1 else p for p in w.split("-")))
    return " ".join(out)

def clean_name(s):
    s = re.sub(r"Bt\s*[1lI]\s*s", "Btls", s)
    s = re.sub(r"\bL(\d{2})ML", r"1\1ML", s)
    s = re.sub(r"(\d{2,4})\s*ML", lambda m: (str(_fix_size(int(m.group(1)))) if _fix_size(int(m.group(1))) else m.group(1)) + "ML", s, flags=re.I)
    s = re.sub(r"(?<=\d)S(?=\d)", "5", s)
    s = re.sub(r"(?<=[\d(])[OQ]|[OQ](?=\d)", "0", s)
    s = re.sub(r"\s+\(", "(", s)                 # "Btls (0020)" -> "Btls(0020)"
    s = re.sub(r"\.\.", ".", s)
    s = re.sub(r"(?<=\d)\s+(?=ML)", "", s, flags=re.I)
    s = re.sub(r"[|~\\_^*©«»=¢“”`]", " ", s)
    s = re.sub(r"\s+", " ", s).strip(" .-,:;'\"")
    return s

def snap(name, size, catalog):
    """Return (name, status)  status: 'catalog' | 'new'"""
    k = _key(name)
    best, bscore = None, 0
    for cn in catalog:
        if size and _size_of(cn) and _size_of(cn) != size:
            continue
        ck = _key(cn)
        a = SequenceMatcher(None, k, ck).ratio()
        b = SequenceMatcher(None, k, ck[:len(k) + 3]).ratio() * 0.96   # last line of name lost by OCR
        sc = max(a, b)
        if sc > bscore: best, bscore = cn, sc
    if best and bscore >= 0.80: return best, "catalog"
    return smart_title(name), "new"

# ---------------------------------------------------------------------- OCR ---
def _skew_angle(img):
    import numpy as np
    from PIL import Image
    g = img.convert("L"); g.thumbnail((900, 1200))
    a = (np.array(g) < 140) * 255
    base = Image.fromarray(a.astype("uint8"))
    best = (0.0, -1.0)
    for ang in np.arange(-3, 3.01, 0.1):
        p = np.array(base.rotate(float(ang), resample=Image.BILINEAR, fillcolor=0)).sum(axis=1).astype(float)
        v = float(np.var(np.diff(p)))
        if v > best[1]: best = (float(ang), v)
    return best[0]

def ocr_pdf(pdf, dpi=250):
    from PIL import Image
    tmp = tempfile.mkdtemp()
    try:
        subprocess.run(["pdftoppm", "-r", str(dpi), "-png", pdf, os.path.join(tmp, "p")], check=True)
        pages = []
        for img in sorted(glob.glob(os.path.join(tmp, "p*.png"))):
            im = Image.open(img).convert("L")
            ang = _skew_angle(im)
            if abs(ang) > 0.05:
                im = im.rotate(ang, resample=Image.BICUBIC, fillcolor=255)
            im.save(img)
            r = subprocess.run(["tesseract", img, "-", "--psm", "6"], capture_output=True, text=True)
            W, H = im.size
            im.crop((int(.36 * W), 0, W, H)).save(img + ".r.png")
            rr = subprocess.run(["tesseract", img + ".r.png", "-", "--psm", "6", "-c",
                                 "tessedit_char_whitelist=0123456789.,-"], capture_output=True, text=True)
            pages.append(dict(full=r.stdout, right=rr.stdout))
        return pages
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

# ------------------------------------------------------------------- parsing ---
def _norm_nums(line):
    s = re.sub(r"(\d)\s+\.(\d\d)\b", r"\1.\2", line)       # "14508 .28"
    s = re.sub(r"(\d)\.\s+(\d\d)\b", r"\1.\2", s)             # "15624. 32"
    s = re.sub(r"(\d),(\d\d)\b", r"\1.\2", s)                 # "30108,84"
    s = re.sub(r"(?<=\d)[~:](?=\d\d\b)", ".", s)              # "11873~32"
    return s

def _numtoks(s):
    """[(start, text)] numeric tokens; joins '6 327.04' -> '6327.04'"""
    toks = [(m.start(), m.group(0).rstrip(".")) for m in re.finditer(r"\d+(?:\.\d+)?\.?", s)]
    out = []
    for p, t in toks:
        if out and re.fullmatch(r"\d{3}\.\d\d", t) and re.fullmatch(r"\d{1,3}", out[-1][1]) and p - (out[-1][0] + len(out[-1][1])) <= 2:
            out[-1] = (out[-1][0], out[-1][1] + t)
        else:
            out.append((p, t))
    return out

_SER = {"]": "1", "!": "1", "g": "9", "q": "9", "i": "1", "l": "1", "I": "1", "|": "1", "O": "0", "o": "0", "S": "5", "s": "5", "Z": "2", "z": "2", "B": "8", "G": "6"}
def _serial_variants(tok):
    outs = [""]
    for ch in tok:
        opts = ["8", "9"] if ch in "gq" else [_SER.get(ch, ch)]
        outs = [o + x for o in outs for x in opts]
    return outs

def _serial_is(tok, expected):
    for t in _serial_variants(tok):
        if t.isdigit() and int(t) == expected: return True
        if len(t) == 3 and t.isdigit() and int(t[:2]) == expected: return True     # stray trailing mark, e.g. "283"
    return False

def _serial_near(tok, expected):
    for t in _serial_variants(tok):
        if t.isdigit() and len(t) <= 2 and max(1, expected - 3) <= int(t) <= expected + 3: return True
    return False

def _has_tail(s):
    s = _norm_nums(s)
    toks = _numtoks(s)
    for i, (p, t) in enumerate(toks):
        if re.fullmatch(r"\d{3,6}\.\d{1,2}", t) and float(t) >= 500 and len(toks) - i >= 3:
            return True
    return False

def _resolve_tail(line, pack, cat_rates=()):
    """Numbers of an item line -> dict(rate, cb, btls, amt, verified, size, cut).
    'verified' means amount == CB*rate + Btls*rate/pack exactly (so rate, qty and amount agree)."""
    s = _norm_nums(line)
    toks = _numtoks(s)
    ocr_rates = []
    for i, (p, t) in enumerate(toks):
        if re.fullmatch(r"\d{3,6}\.\d{1,2}", t) and float(t) >= 500: ocr_rates.append((float(t), i, p))
        elif re.fullmatch(r"\d{6}", t): ocr_rates.append((float(t[:-2] + "." + t[-2:]), i, p))
        elif re.fullmatch(r"\d{7,8}", t): ocr_rates.append((float(t[:-3] + "." + t[-2:]), i, p))   # '.' read as a digit
    cands = [(r, i, p) for r, i, p in ocr_rates] + [(float(r), None, None) for r in cat_rates]
    if not cands: return None

    def attempt(rate, ri):
        after = [t for q, t in (toks[ri + 1:] if ri is not None else toks)]
        best = None
        for k, t in enumerate(after):
            try: a = float(t)
            except ValueError: continue
            if a <= 0: continue
            if ri is None and abs(a - rate) < 1e-6:
                twice = sum(1 for q, tt in toks if tt == t) >= 2
                misread = any("." in tt and tt != t and abs(float(tt) - a) / a < 0.15 for q, tt in toks if re.fullmatch(r"\d+\.\d+", tt))
                if not (twice or misread): continue
            units = D(a) / D(rate) * pack
            if abs(units - units.to_integral_value()) > D("0.02") or units < 1: continue
            u = int(units.to_integral_value()); cb, bt = divmod(u, pack)
            if cb > 300: continue
            if abs(money(D(rate) * cb + D(rate) * bt / pack) - a) > 0.02: continue
            ocr = [x for x in after[:k] if x.isdigit()][-2:]
            score = 2 if ocr == [str(cb), str(bt)] else 1
            if best is None or score > best[0]: best = (score, cb, bt, a)
        return best

    for rate, ri, rp in cands:
        best = attempt(rate, ri)
        if best:
            size, cut = None, rp
            if ri is not None and ri > 0:
                pp, pt = toks[ri - 1]
                if pt.isdigit() and 2 <= len(pt) <= 4 and not pt.startswith("0") and rp - (pp + len(pt)) <= 10: size, cut = (int(pt) if int(pt) in SIZES else None), pp
            return dict(rate=rate, cb=best[1], btls=best[2], amt=best[3], verified=True, size=size, cut=cut,
                        from_catalog=ri is None)
    # unverified: best effort from the first OCR rate
    rate, ri, rp = cands[0]
    if ri is None: return None
    after = [t for q, t in toks[ri + 1:]]
    ints = [t for t in after if t.isdigit() and int(t) <= 300]
    cb = int(ints[0]) if ints else 0
    bt = int(ints[1]) if len(ints) > 1 else 0
    a = float(next((t for t in after if "." in t), "0") or 0)
    size, cut = None, rp
    if ri > 0:
        pp, pt = toks[ri - 1]
        if pt.isdigit() and 2 <= len(pt) <= 4 and not pt.startswith("0") and rp - (pp + len(pt)) <= 10: size, cut = (int(pt) if int(pt) in SIZES else None), pp
    return dict(rate=rate, cb=cb, btls=bt, amt=a, verified=False, size=size, cut=cut, from_catalog=False)

def _cut_junk(txt):
    keep = []
    for w in txt.split():
        if re.fullmatch(r"[A-Za-z][A-Za-z'’\-\.&,:]*", w) or re.search(r"ML|\(\d", w, re.I) or re.fullmatch(r"\d{1,2}|\(?\d+%\)?", w):
            keep.append(w)
        else:
            break
    return " ".join(keep)

def _trim_junk_words(txt):
    w = txt.split()
    pops = 0
    while w and pops < 4 and not re.fullmatch(r"[A-Za-z][A-Za-z'’\-\.&]*", w[-1]):
        w.pop(); pops += 1
    return " ".join(w)

def read_invoice(pdf, catalog=None, pages=None):
    catalog = catalog if catalog is not None else load_catalog()
    pages = pages or ocr_pdf(pdf)
    right_lines = []
    if pages and isinstance(pages[0], dict):
        for p in pages: right_lines += [l.strip() for l in p["right"].splitlines() if l.strip()]
        pages = [p["full"] for p in pages]
    text = "\n".join(pages)
    warn = []

    nos = Counter(re.findall(r"SSGR\d{8}", text))
    inv_no = nos.most_common(1)[0][0] if nos else ""
    dts = Counter(re.findall(r"(?:SSGR\d{8}\s+|Date\s*:?\s*)(\d\d/\d\d/20\d\d)", text))
    date = dt.datetime.strptime(dts.most_common(1)[0][0], "%d/%m/%Y").date() if dts else None
    if not inv_no: warn.append("Invoice number not read")
    if not date: warn.append("Invoice date not read")

    pm = re.search(r"\bTotal\s+(\d+)\s+(\d+)\s+([\d.]+)\s*\n", pages[0]) if pages else None
    permit = (int(pm.group(1)), int(pm.group(2))) if pm else None

    # ---- split the invoice pages into item blocks
    blocks, cur, started = [], None, False
    for pg in pages:
        if "Bill Of" not in pg: continue
        for ln in pg.splitlines():
            s = ln.strip()
            if not s: continue
            if not started:
                if re.search(r"Item\s*Name|Amount\s*\(", s): started = True
                continue
            if re.match(r"Total\b", s) and cur is not None and re.search(r"\d{4,}\.\d\d", s): break
            if BAD_LINE.search(s): continue
            expected = len(blocks) + 1
            lead = re.match(r"^\W*([0-9A-Za-z|\]!]{1,3})\s+(\S.*)$", s)
            if lead and not _serial_is(lead.group(1), expected) and len(lead.group(1)) <= 2:
                l2 = re.match(r"^\W*[0-9A-Za-z|\]!]{1,2}\s+([0-9A-Za-z|\]!]{1,3})\s+(\S.*)$", s)
                if l2 and _serial_is(l2.group(1), expected): lead = l2
            if lead and _serial_is(lead.group(1), expected) and re.search(r"[A-Za-z]{3,}", lead.group(2)):
                cur = dict(lines=[lead.group(2)]); blocks.append(cur)
            elif lead and _has_tail(s) and re.search(r"[A-Za-z]{3,}", lead.group(2)) and _serial_near(lead.group(1), expected):
                cur = dict(lines=[lead.group(2)]); blocks.append(cur)
            elif cur is not None:
                if _has_tail(s) and not any(_has_tail(x) for x in cur["lines"]):
                    cur["lines"].append(s)
                elif re.search(r"[A-Za-z]{4,}|\d\s*ML", s, re.I) and sum(ch.isalnum() for ch in s) >= 0.6 * len(s):
                    m = re.search(r"\)[^)]*$", s)
                    cur["lines"].append(_cut_junk(s[:m.start() + 1] if m else s))
        # page break keeps `cur` so a wrapped name line joins the previous item
    tot_amt = None
    for pg in reversed(pages):
        m = re.search(r"^\s*Total\b[^\n]*?(\d{4,}\.\d\d)\s*$", pg, re.M)
        if m: tot_amt = float(m.group(1)); break
    iv = re.search(r"nvoice\s*Value\s*:\s*([\d.]+)", text); tcs = re.search(r"TCS\s*Value\s*:\s*([\d.]+)", text)
    if tot_amt is None and iv and tcs:
        try: tot_amt = money(D(iv.group(1)) - D(tcs.group(1)))
        except Exception: pass

    out, used_right = [], set()
    for i, b in enumerate(blocks, 1):
        joined = " ".join(b["lines"])
        guess_raw = clean_name(re.sub(r"[^A-Za-z0-9()%.\-' ]", " ", joined))
        size_guess = _size_of(guess_raw)
        # first pass: find the line holding the numbers (any pack, any catalogue rate of items of that size)
        def cat_rates_for(sz):
            rs = set()
            for n, rr in catalog.items():
                if not sz or _size_of(n) == sz: rs |= set(rr)
            return rs
        m = re.search(r"ML\s*x\s*(\d+)", guess_raw, re.I)
        pack0 = int(m.group(1)) if m else DEFAULT_PACK.get(size_guess or 180, 48)
        tail_line, t0 = None, None
        for li, ln in enumerate(b["lines"]):
            r = _resolve_tail(ln, pack0, ())
            if r and r["verified"]: tail_line, t0 = ln, r; break
        if t0 is None:
            for li, ln in enumerate(b["lines"]):
                r = _resolve_tail(ln, pack0, cat_rates_for(size_guess))
                if r and (r["verified"] or tail_line is None): tail_line, t0 = ln, r
                if r and r["verified"]: break
        parts = []
        for li, ln in enumerate(b["lines"]):
            if ln is tail_line:
                if li == 0:
                    s_ = _norm_nums(ln)
                    cut = t0["cut"] if t0 and t0["cut"] is not None else None
                    if cut is None:
                        mm = re.search(r"\s\d{2,4}\b", s_); cut = mm.start() if mm else len(s_)
                    parts.append(_trim_junk_words(s_[:cut]))
            else:
                parts.append(ln)
        name_raw = clean_name(" ".join(p for p in parts if p))
        size = _size_of(name_raw) or (t0 or {}).get("size") or size_guess
        name, status = snap(name_raw, size, catalog)
        flags = []
        if status == "new" and t0 and t0["verified"]:
            byrate = [n for n, rr in catalog.items() if any(abs(r - t0["rate"]) < 0.005 for r in rr) and (not size or _size_of(n) in (None, size))]
            if byrate:
                kk = _key(name_raw)
                name = max(byrate, key=lambda n: SequenceMatcher(None, kk, _key(n)).ratio()) if len(byrate) > 1 else byrate[0]
                if len(byrate) == 1 or SequenceMatcher(None, kk, _key(name)).ratio() > 0.45:
                    status = "catalog"; flags.append("name matched by rate - check")
        m = re.search(r"ML\s*x\s*(\d+)", name, re.I)
        pack = int(m.group(1)) if m else DEFAULT_PACK.get(size or _size_of(name) or 180, 48)
        if status == "new": flags.append("new item - check name")
        t = None
        if tail_line:
            known_rates = catalog.get(name, ())
            t = _resolve_tail(tail_line, pack, ())
            if not (t and t["verified"]) and known_rates:
                t2 = _resolve_tail(tail_line, pack, known_rates)
                if t2 and t2["verified"]: t = t2; flags.append("rate corrected from catalogue")
        if (not t or not t["verified"]) and right_lines:
            want = [t["rate"]] if t else list(catalog.get(name, ()))
            for ri, rl in enumerate(right_lines):
                if ri in used_right: continue
                r = _resolve_tail(rl, pack, ())
                if r and r["verified"] and (not want or any(abs(r["rate"] - w) < 0.005 for w in want)):
                    t = r; used_right.add(ri); flags.append("numbers read from column scan"); break
        elif t and t["verified"] and right_lines:
            for ri, rl in enumerate(right_lines):
                if ri in used_right: continue
                r = _resolve_tail(rl, pack, ())
                if r and r["verified"] and abs(r["rate"] - t["rate"]) < 0.005 and abs(r["amt"] - t["amt"]) < 0.005:
                    used_right.add(ri); break
        if not t:
            out.append(dict(sl=i, name=name, cb=0, btls=0, rate=0.0, amt=0.0, pack=pack, flags=flags + ["numbers not read - fill in"]))
            continue
        rate = t["rate"]
        calc = money(D(rate) * t["cb"] + D(rate) * t["btls"] / pack)
        amt = t["amt"]
        if not t["verified"] or abs(calc - amt) >= 0.03:
            flags.append(f"check qty/amount (read {t['cb']} CB {t['btls']} Btl = {amt})")
            amt = calc
        elif catalog.get(name) and rate not in catalog[name] and status == "catalog":
            flags.append("rate differs from earlier invoices")
        out.append(dict(sl=i, name=name, cb=t["cb"], btls=t["btls"], rate=rate, amt=amt, pack=pack, flags=flags))

    # single missing row: derive it from the invoice total
    bad = [o for o in out if "numbers not read - fill in" in o["flags"]]
    if len(bad) == 1 and tot_amt is not None:
        o = bad[0]; rs = catalog.get(o["name"], ())
        rest = sum(D(x["amt"]) for x in out if x is not o)
        gap = D(str(tot_amt)) - rest
        if len(rs) == 1 and gap > 0:
            r = list(rs)[0]; units = gap / D(r) * o["pack"]
            if abs(units - units.to_integral_value()) < D("0.02") and units >= 1:
                u = int(units.to_integral_value()); o["cb"], o["btls"] = divmod(u, o["pack"]); o["rate"] = r
                o["amt"] = money(D(r) * o["cb"] + D(r) * o["btls"] / o["pack"])
                o["flags"] = [f for f in o["flags"] if f != "numbers not read - fill in"] + ["derived from invoice total - check"]
    s_amt = money(sum(D(o["amt"]) for o in out))
    s_cb, s_bt = sum(o["cb"] for o in out), sum(o["btls"] for o in out)
    if tot_amt is not None and abs(tot_amt - s_amt) > 0.05:
        warn.append(f"Line amounts add to {s_amt:,.2f} but invoice total reads {tot_amt:,.2f}")
    if permit and permit != (s_cb, s_bt):
        warn.append(f"Permit page says {permit[0]} CB / {permit[1]} Btls, items add to {s_cb} / {s_bt}")
    if not out: warn.append("No line items found")
    return dict(file=os.path.basename(pdf), invoice_no=inv_no, date=date.isoformat() if date else "",
                items=out, total_read=tot_amt, total_items=s_amt, cb=s_cb, btls=s_bt, warnings=warn)

# ------------------------------------------------------------------ building ---
def _cp(dst, src):
    dst._style = copy(src._style)

def _block_merges(ws, c0, r1, r2):
    res = []
    for m in ws.merged_cells.ranges:
        if c0 <= m.min_col <= c0 + 5 and r1 <= m.min_row <= r2 and m.max_col <= c0 + 5:
            res.append((m.min_row, m.min_col - c0, m.max_row, m.max_col - c0))
    return res

def split_rows(n, parts=2):
    """sales invoices per purchase invoice: first half = ceil(n/2)"""
    if n <= 1 or parts == 1: return [(0, n)]
    h = math.ceil(n / 2)
    return [(0, h), (h, n)]

def build_workbook(invoices, out_path, sales_start=911, markup=0.10, first_sale_offset=2, sale_gap_days=3,
                   owner=None, parts=2):
    invoices = sorted(invoices, key=lambda v: (v["date"], v["invoice_no"]))
    first = dt.date.fromisoformat(invoices[0]["date"])
    mon = first.strftime("%B %Y")
    pname, sname = f"Purchase {mon}", f"Sales {mon}"

    wb = openpyxl.load_workbook(TEMPLATE)
    tS, tP, tL = wb["Summary"], [w for w in wb if w.title.startswith("Purchase")][0], [w for w in wb if w.title.startswith("Sales")][0]
    tS.title, tP.title, tL.title = "_tS", "_tP", "_tL"          # free the real names for the new sheets
    P, S, Su = wb.create_sheet(pname), wb.create_sheet(sname), wb.create_sheet("Summary")
    for w in (P, S, Su): w.sheet_view.showGridLines = False; w.sheet_view.zoomScale = tP.sheet_view.zoomScale or 63
    Su.sheet_view.zoomScale = tS.sheet_view.zoomScale

    def header_proto(tws): return {(r, o): tws.cell(r, 2 + o) for r in range(3, 14) for o in range(6)}
    hP, hL = header_proto(tP), header_proto(tL)
    mP, mL = _block_merges(tP, 2, 3, 12), _block_merges(tL, 2, 3, 12)
    txt = {k: tP.cell(*k).value for k in [(3, 2), (4, 2), (5, 2), (8, 2), (8, 4), (9, 2), (9, 4), (10, 2), (10, 4), (11, 2), (11, 4), (12, 2)]}
    if owner: txt[(9, 2)] = owner
    # item/total row prototypes (taken from template Purchase block 3 / Sales block 1)
    def row_proto(tws, c0, first_r, mid_r, tot_r):
        return ([tws.cell(first_r, c0 + o) for o in range(6)], [tws.cell(mid_r, c0 + o) for o in range(6)],
                [tws.cell(tot_r, c0 + o) for o in range(6)])
    pf, pm_, pt = row_proto(tP, 2, 14, 15, 46)
    sf, sm_, st = row_proto(tL, 2, 14, 15, 30)

    def setup_block(ws, tws, hp, merges, c0, title_no, d, widths_from, sales):
        for (r, o), src in hp.items():
            _cp(ws.cell(r, c0 + o), src)
        for r in range(3, 14): ws.row_dimensions[r].height = 18
        for (r1, o1, r2, o2) in merges:
            ws.merge_cells(start_row=r1, start_column=c0 + o1, end_row=r2, end_column=c0 + o2)
        ws.cell(3, c0, txt[(3, 2)]); ws.cell(4, c0, txt[(4, 2)]); ws.cell(5, c0, txt[(5, 2)])
        ws.cell(6, c0, f"Invoice No : {title_no}"); ws.cell(6, c0 + 2, "Date :")
        dc = ws.cell(6, c0 + 3, d); dc.number_format = "dd/mm/yyyy"
        ws.cell(8, c0, txt[(8, 2)]); ws.cell(8, c0 + 2, txt[(8, 4)])
        for r in (9, 10, 11): ws.cell(r, c0, txt[(r, 2)]); ws.cell(r, c0 + 2, txt[(r, 4)])
        ws.cell(12, c0, txt[(12, 2)])
        for o, h in enumerate(["SL.No", "ITEMS NAME", "CB" if sales else "CBs", "Blts.", "RATE for CB",
                               "SALES AMOUNT" if sales else "Purchase Amt"]):
            ws.cell(13, c0 + o, h)
        for o in range(6):
            ws.column_dimensions[L(c0 + o)].width = widths_from.column_dimensions[L(9 + o)].width or 10
        ws.column_dimensions[L(c0 + 6)].width = 3

    # ------------------------------------------------ Purchase sheet
    p_total_cells, p_rows = [], []
    for bi, inv in enumerate(invoices):
        c0 = 2 + 7 * bi
        d = dt.date.fromisoformat(inv["date"])
        setup_block(P, tP, hP, mP, c0, inv["invoice_no"], d, tP, False)
        r = 14
        for k, it in enumerate(inv["items"]):
            proto = pf if k == 0 else pm_
            for o in range(6): _cp(P.cell(r, c0 + o), proto[o])
            for o, v in enumerate([k + 1, it["name"], it["cb"], it["btls"], it["rate"], it["amt"]]):
                P.cell(r, c0 + o, v)
            P.row_dimensions[r].height = 18
            r += 1
        for o in range(6): _cp(P.cell(r, c0 + o), pt[o])
        P.cell(r, c0 + 1, "Total")
        for o in (2, 3, 5):
            col = L(c0 + o); P.cell(r, c0 + o, f"=SUM({col}14:{col}{r - 1})")
        P.row_dimensions[r].height = 18
        p_total_cells.append(f"{L(c0 + 5)}{r}")
        p_rows.append((c0, 14, r - 1))

    # ------------------------------------------------ Sales sheet
    s_total_cells, s_nos = [], []
    no, sdate, bi = sales_start, first + dt.timedelta(days=first_sale_offset), 0
    qp = f"'{pname}'"
    for inv, (pc0, pr1, pr2) in zip(invoices, p_rows):
        n = pr2 - pr1 + 1
        for (a, b) in split_rows(n, parts):
            if b <= a: continue
            c0 = 2 + 7 * bi
            setup_block(S, tL, hL, mL, c0, no, sdate, tL, True)
            r = 14
            for k, pr in enumerate(range(pr1 + a, pr1 + b)):
                proto = sf if k == 0 else sm_
                for o in range(6): _cp(S.cell(r, c0 + o), proto[o])
                S.cell(r, c0, k + 1)
                S.cell(r, c0 + 1, f"={qp}!{L(pc0 + 1)}{pr}")
                S.cell(r, c0 + 2, f"={qp}!{L(pc0 + 2)}{pr}")
                S.cell(r, c0 + 3, f"={qp}!{L(pc0 + 3)}{pr}")
                S.cell(r, c0 + 4, f"={qp}!{L(pc0 + 4)}{pr}")
                S.cell(r, c0 + 5, f"=({qp}!{L(pc0 + 5)}{pr})*{1 + markup:g}")
                S.row_dimensions[r].height = 18
                r += 1
            for o in range(6): _cp(S.cell(r, c0 + o), st[o])
            S.cell(r, c0 + 1, "Total")
            for o in (2, 3, 5):
                col = L(c0 + o); S.cell(r, c0 + o, f"=SUM({col}14:{col}{r - 1})")
            S.row_dimensions[r].height = 18
            s_total_cells.append(f"{L(c0 + 5)}{r}"); s_nos.append(no)
            no += 1; sdate += dt.timedelta(days=sale_gap_days); bi += 1

    # ------------------------------------------------ Summary sheet
    for r in range(2, 9):
        for c in range(2, 8): _cp(Su.cell(r, c), tS.cell(r, c))
    for m in tS.merged_cells.ranges:
        if m.min_row <= 8: Su.merge_cells(str(m))
    Su["B2"], Su["B3"] = tS["B2"].value, tS["B3"].value
    Su["B5"] = f"Liquor Details for the month of {mon}"
    Su["B7"], Su["E7"] = "Purchase Details", "Sales Details"
    Su["B8"], Su["C8"] = "Sl.No", dt.datetime(first.year, first.month, 1); Su["C8"].number_format = "mmm-yy"
    Su["E8"], Su["F8"], Su["G8"] = "Sl.No", "Invoice no.", dt.datetime(first.year, first.month, 1); Su["G8"].number_format = "mmm-yy"
    for r in range(1, 9): Su.row_dimensions[r].height = tS.row_dimensions[r].height
    for col, w in tS.column_dimensions.items(): Su.column_dimensions[col].width = w.width
    for i, cell in enumerate(p_total_cells):
        r = 9 + i
        for c, srcc in ((2, 9), (3, 9)): _cp(Su.cell(r, c), tS.cell(srcc, c))
        Su.cell(r, 2, i + 1); Su.cell(r, 3, f"={qp}!{cell}")
    pr_tot = 9 + len(p_total_cells)
    for c in (2, 3): _cp(Su.cell(pr_tot, c), tS.cell(13, c))
    Su.cell(pr_tot, 2, "Purchase Total"); Su.cell(pr_tot, 3, f"=SUM(C9:C{pr_tot - 1})")
    qs = f"'{sname}'"
    for i, (cell, no_) in enumerate(zip(s_total_cells, s_nos)):
        r = 9 + i
        for c in (5, 6, 7): _cp(Su.cell(r, c), tS.cell(10, c))
        Su.cell(r, 5, i + 1); Su.cell(r, 6, no_); Su.cell(r, 7, f"={qs}!{cell}")
        Su.cell(r, 7).number_format = "#,##0"
    st_row = 9 + len(s_total_cells)
    for c in (5, 6, 7): _cp(Su.cell(st_row, c), tS.cell(17, c))
    Su.cell(st_row, 6, "Sales Total"); Su.cell(st_row, 7, f"=SUM(G9:G{st_row - 1})")
    Su.cell(st_row, 7).number_format = "#,##0"

    for w in (tS, tP, tL): wb.remove(w)
    wb._sheets = [Su, P, S]
    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    wb.save(out_path)
    return dict(purchase_sheet=pname, sales_sheet=sname, sales_numbers=s_nos, month=mon)
