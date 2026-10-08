import os, re, json, csv, traceback
from pathlib import Path
from typing import List, Dict, Any

NUM_RE = re.compile(r'^-?\d[\d,]*(?:\.\d+)?$')
ITEM_START_RE = re.compile(r'^\s*(\d{1,4})\s+(\d{3,})\s+(.+?)\s*$')
META_PATTERNS = {
    'poNumber': [r'PO\s*No\s*[:\-]\s*(?:-\s*)?([^\s]+)', r'Purchase\s*Order\s*(?:No|Number)?\s*[:\-]\s*(?:-\s*)?([^\s]+)'],
    'grnNumber': [r'GRN\s*No\s*[:\-]\s*(?:-\s*)?([^\s]+)', r'GRN\s*(?:Number|#)\s*[:\-]\s*(?:-\s*)?([^\s]+)'],
    'invoiceNumber': [r'Invoice\s*No\s*[:\-]\s*(?:-\s*)?([^\s]+)', r'Invoice\s*(?:Number|#)\s*[:\-]\s*(?:-\s*)?([^\s]+)'],
}

def clean(s):
    return re.sub(r'\s+', ' ', str(s or '')).strip()

def to_num(s):
    s = str(s).strip().replace(',', '')
    if not s: return None
    try:
        return float(s)
    except: return None

def extract_meta(text):
    flat = clean(text.replace('\n', ' '))
    out = {}
    for key, pats in META_PATTERNS.items():
        for p in pats:
            m = re.search(p, flat, re.I)
            if m:
                val = m.group(1).strip().lstrip('-—:').strip()
                out[key] = val
                break
    return out

def extract_vendor(text):
    m = re.search(r'Vendor\s*Name\s*[:\-]\s*(.*?)(?:\s+PO\s*No\s*[:\-]|$)', clean(text.replace('\n',' ')), re.I)
    return m.group(1).strip() if m else ''

def is_num(x):
    return bool(NUM_RE.match(x.replace('%','').strip()))

def parse_numeric_tail(line):
    toks = line.split()
    # We expect exactly 15 numeric fields after the 3 identifying fields:
    # MRP, Exp, Recv, UnitPrice, Taxable, CGST rate/amount, SGST rate/amount,
    # IGST rate/amount, CESS rate/amount, AddCess amount, Total.
    for n in (15, 16, 17):
        if len(toks) >= n + 3 and all(is_num(x) for x in toks[-n:]):
            return toks, [to_num(x) for x in toks[-n:]]
    return None, None

def parse_layout_rows(text):
    lines = [x.strip() for x in text.replace('\r','').split('\n') if x.strip()]
    starts = []
    for i, line in enumerate(lines):
        m = ITEM_START_RE.match(line)
        if m:
            # Must look like a genuine item: SKU >= 3 digits and not a total/header.
            if not line.lower().startswith(('total', 'amount')):
                starts.append((i, int(m.group(1)), m.group(2), m.group(3)))

    detail_candidates = []
    for i, line in enumerate(lines):
        toks, nums = parse_numeric_tail(line)
        if nums is not None and len(toks) >= 18:
            # first three tokens identify vendor SKU/bin/lot in the common GRN layout
            detail_candidates.append((i, toks, nums))

    rows=[]
    # Pair item starts and detail lines in document order. This is deliberate:
    # descriptions can wrap over multiple lines and the detail line is a separate line.
    for idx, (si, sr, sku, first_desc) in enumerate(starts):
        next_start = starts[idx+1][0] if idx+1 < len(starts) else len(lines)
        details = [(di,t,n) for di,t,n in detail_candidates if si < di < next_start]
        if not details:
            continue
        di, toks, nums = details[-1]
        # Description is everything from first_desc through the lines before detail line,
        # excluding stray table/header text.
        desc_parts=[first_desc]
        for j in range(si+1, di):
            ln=lines[j]
            if ln.lower().startswith(('total:', 'amount')): continue
            # Exclude obvious duplicate/detail fragments.
            if parse_numeric_tail(ln)[1] is not None: continue
            if re.match(r'^\d+\s+\S+\s+\S+', ln): continue
            desc_parts.append(ln)
        desc=clean(' '.join(desc_parts))
        # Last 15 numeric fields are the financial/quantity tail.
        if len(nums) != 15:
            continue
        mrp, expq, recvq, unit, taxable, cgr, cga, sgr, sga, igr, iga, cer, cea, addcess, total = nums
        rows.append({
            'Sr. No.': sr, 'SKU Code': sku, 'SKU Desc': desc,
            'Lot MRP (INR)': mrp, 'Exp Qty': expq, 'Recv Qty': recvq,
            'Unit Price (INR)': unit, 'Taxable Value (INR)': taxable,
            'CGST Rate': cgr, 'CGST Amount': cga,
            'SGST/UGST Rate': sgr, 'SGST/UGST Amount': sga,
            'IGST Rate': igr, 'IGST Amount': iga,
            'CESS Rate': cer, 'CESS Amount': cea,
            'Add. Cess Rate': '', 'Add. Cess Amount': addcess,
            'Total (INR)': total,
            '_vendorSku': toks[0], '_skuBin': toks[1], '_lotNo': toks[2]
        })
    return rows

def parse_pdf_text(text, source_name):
    meta=extract_meta(text)
    vendor=extract_vendor(text)
    rows=parse_layout_rows(text)
    if not rows:
        # Strategy 2: generic row parser. Looks for any item-start line followed by a
        # detail line with a 15-number tail, regardless of bin name.
        rows=parse_layout_rows(re.sub(r'\s{2,}', ' ', text))
    for r in rows:
        r.update({'PO Number': meta.get('poNumber',''), 'GRN Number': meta.get('grnNumber',''),
                  'Invoice Number': meta.get('invoiceNumber','')})
    return rows, meta, vendor

def parse_pdf_file(path):
    # Parser 1: pypdf text extraction. It preserves the row/detail-line structure
    # of the supplied GRN PDFs and is the primary table parser.
    text=''
    try:
        from pypdf import PdfReader
        reader=PdfReader(path)
        text='\n'.join((p.extract_text() or '') for p in reader.pages)
    except Exception:
        pass
    rows, meta, vendor=parse_pdf_text(text, os.path.basename(path)) if text else ([],{},'')

    # Parser 2: PyMuPDF fallback. Useful when pypdf cannot extract usable text.
    if not rows:
        try:
            import fitz
            doc=fitz.open(path)
            alt='\n'.join((page.get_text('text') or '') for page in doc)
            rows, meta, vendor=parse_pdf_text(alt, os.path.basename(path))
            if alt: text=alt
        except Exception:
            pass

    # Parser 3: OCR fallback for scanned/image-only PDFs.
    if not rows:
        try:
            import fitz, pytesseract, io
            from PIL import Image
            doc=fitz.open(path); ocr_parts=[]
            for page in doc:
                pix=page.get_pixmap(matrix=fitz.Matrix(2.5,2.5), alpha=False)
                img=Image.open(io.BytesIO(pix.tobytes('png')))
                ocr_parts.append(pytesseract.image_to_string(img, config='--psm 6'))
            ocr_text='\n'.join(ocr_parts)
            rows, meta, vendor=parse_pdf_text(ocr_text, os.path.basename(path))
            if ocr_text: text=ocr_text
        except Exception:
            pass
    return rows, meta, vendor, text

def parse_excel_file(path):
    from openpyxl import load_workbook
    wb=load_workbook(path, data_only=True, read_only=True)
    rows=[]
    aliases={
        'Sr. No.':['sr no','sr. no','sr','serial no','s no'],
        'SKU Code':['sku code','skucode','item code','product code','material code','vendor sku'],
        'SKU Desc':['sku desc','sku description','item description','product description','material description'],
        'PO Number':['po number','po no','po no.','purchase order','purchase order no'],
        'GRN Number':['grn number','grn no','grn no.','receipt no'],
        'Invoice Number':['invoice number','invoice no','invoice no.','bill no'],
        'Lot MRP (INR)':['lot mrp','mrp'], 'Exp Qty':['exp qty','expected qty'],
        'Recv Qty':['recv qty','received qty','accepted qty'], 'Unit Price (INR)':['unit price','rate'],
        'Taxable Value (INR)':['taxable value','taxable amount'], 'CGST Rate':['cgst rate'],
        'CGST Amount':['cgst amount'], 'SGST/UGST Rate':['sgst rate','ugst rate','sgst/ugst rate'],
        'SGST/UGST Amount':['sgst amount','ugst amount','sgst/ugst amount'], 'IGST Rate':['igst rate'],
        'IGST Amount':['igst amount'], 'CESS Rate':['cess rate'], 'CESS Amount':['cess amount'],
        'Add. Cess Rate':['add cess rate','add.cess rate'], 'Add. Cess Amount':['add cess amount','add.cess amount'],
        'Total (INR)':['total','total amount','amount']}
    def norm(v): return clean(v).lower().replace('_',' ')
    for ws in wb.worksheets:
        data=list(ws.iter_rows(values_only=True))
        if not data: continue
        header_i=None; mapping={}
        for i,row in enumerate(data[:40]):
            vals=[norm(x) for x in row]
            score=0
            mp={}
            for c,aliases_list in aliases.items():
                for j,v in enumerate(vals):
                    if v and any(a==v or a in v for a in aliases_list): mp[c]=j; score+=1; break
            if score>=4:
                header_i=i; mapping=mp; break
        if header_i is None: continue
        for ridx,row in enumerate(data[header_i+1:], start=1):
            if not any(x not in (None,'') for x in row): continue
            out={c:(row[j] if j<len(row) else '') for c,j in mapping.items()}
            # Keep only rows that look like items.
            if out.get('SKU Code') in (None,'') and out.get('SKU Desc') in (None,''): continue
            rows.append(out)
    return rows, {}, '', ''

def parse_file(path):
    ext=Path(path).suffix.lower()
    if ext=='.pdf':
        return parse_pdf_file(path)
    if ext in ('.xlsx','.xlsm'):
        rows,meta,vendor,text=parse_excel_file(path)
        return rows,meta,vendor,text
    raise ValueError('Unsupported file type')
