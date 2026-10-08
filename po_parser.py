
import re
from pathlib import Path
import fitz
from openpyxl import load_workbook

COLUMNS = [
    "Sr. No.","Item Code","Item Desc","HSN Code","PO Number","PO Date",
    "PO Release Date","Payment Terms","Expected Delivery Date","PO Expiry Date",
    "Vendor Name","Qty","MRP (INR)","Unit Base Cost (INR)","Taxable Value (INR)",
    "CGST Rate","CGST Amount","SGST/UGST Rate","SGST/UGST Amount",
    "IGST Rate","IGST Amount","CESS Rate","CESS Amount",
    "Additional CESS Rate","Additional CESS Amount","Total (INR)"
]
NUM_COLS = {
    "Qty","MRP (INR)","Unit Base Cost (INR)","Taxable Value (INR)",
    "CGST Rate","CGST Amount","SGST/UGST Rate","SGST/UGST Amount",
    "IGST Rate","IGST Amount","CESS Rate","CESS Amount",
    "Additional CESS Rate","Additional CESS Amount","Total (INR)"
}

def num(v):
    if v is None or str(v).strip()=="":
        return 0.0
    s = str(v).replace(",","").replace("₹","").replace("INR","").strip()
    try: return float(s)
    except: return 0.0

def clean(v):
    return re.sub(r"\s+", " ", str(v or "")).strip()

def fmt(v):
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v

def extract_lines(pdf_path):
    doc = fitz.open(pdf_path)
    out=[]
    for page in doc:
        groups=[]
        for w in sorted(page.get_text("words"), key=lambda z:(z[1],z[0])):
            y=w[1]
            g=next((x for x in groups if abs(x["y"]-y)<=2.5),None)
            if g is None:
                g={"y":y,"words":[]}; groups.append(g)
            g["words"].append(w)
        for g in sorted(groups,key=lambda x:x["y"]):
            out.append(" ".join(w[4] for w in sorted(g["words"],key=lambda z:z[0])).strip())
    return [clean(x) for x in out if clean(x)]

def first(patterns, text):
    for p in patterns:
        m=re.search(p,text,re.I|re.S)
        if m: return clean(m.group(1))
    return ""

def parse_meta(lines):
    text="\n".join(lines)
    po = first([r"PO\s*No\s*:\s*([A-Z0-9./_-]+)", r"Purchase Order\s*(?:No|Number)\s*:\s*([A-Z0-9./_-]+)"], text)
    po_date = first([r"PO\s*Date\s*:\s*([A-Za-z0-9, -]+?)(?:\n|PO\s*Release)", r"PO\s*Date\s*:\s*([A-Za-z0-9, -]+)"], text)
    release = first([r"PO\s*Release\s*Date\s*:\s*([A-Za-z0-9, -]+?)(?:\n|Payment Terms)",], text)
    payment = first([r"Payment\s*Terms\s*:\s*([^\n]+)"], text)
    expected = first([r"Expected\s*Delivery\s*Date\s*:\s*([^\n]+)"], text)
    expiry = first([r"PO\s*Expiry\s*Date\s*:\s*([^\n]+)"], text)
    vendor=""
    for line in lines[:20]:
        if "PO Date" in line:
            left=line.split("PO Date",1)[0].strip()
            if left and "SAYZA TRADING" in left.upper():
                vendor="SAYZA TRADING"
                break
    if not vendor:
        for line in lines[:25]:
            if line.upper().strip() == "SAYZA TRADING":
                vendor=line.strip(); break
    return dict(po_number=po,po_date=po_date,po_release_date=release,
                payment_terms=payment,expected_delivery_date=expected,
                po_expiry_date=expiry,vendor_name=vendor)

def parse_pdf(path):
    lines=extract_lines(path)
    meta=parse_meta(lines)
    rows=[]
    current=None
    for line in lines:
        m=re.match(r"^(\d+)\s+(\S+)\s+(.+)$", line)
        if m:
            toks=line.split()
            nums=[]; start=None
            for i in range(2,len(toks)):
                if re.fullmatch(r"-?\d[\d,]*(?:\.\d+)?",toks[i]):
                    # HSN is the first numeric token after the item description begins.
                    # In this PO template it is the 8-digit HSN, followed by 15 value columns.
                    if len(toks[i].replace(',',''))>=6 and i+14 < len(toks):
                        tail=toks[i+1:i+15]
                        if len(tail)==14 and all(re.fullmatch(r"-?\d[\d,]*(?:\.\d+)?",x) for x in tail):
                            start=i; nums=tail; break
            if start is not None:
                sr=int(num(toks[0])); item=toks[1]; hsn=toks[start]
                desc=" ".join(toks[2:start])
                vals=[num(x) for x in nums]
                current={
                    "Sr. No.":sr,"Item Code":item,"Item Desc":desc,"HSN Code":hsn,
                    "PO Number":meta["po_number"],"PO Date":meta["po_date"],
                    "PO Release Date":meta["po_release_date"],"Payment Terms":meta["payment_terms"],
                    "Expected Delivery Date":meta["expected_delivery_date"],"PO Expiry Date":meta["po_expiry_date"],
                    "Vendor Name":meta["vendor_name"],"Qty":vals[0],"MRP (INR)":vals[1],
                    "Unit Base Cost (INR)":vals[2],"Taxable Value (INR)":vals[3],
                    "CGST Rate":vals[4],"CGST Amount":vals[5],"SGST/UGST Rate":vals[6],
                    "SGST/UGST Amount":vals[7],"IGST Rate":vals[8],"IGST Amount":vals[9],
                    "CESS Rate":vals[10],"CESS Amount":vals[11],
                    "Additional CESS Rate":0,"Additional CESS Amount":vals[12],
                    "Total (INR)":vals[13]
                }
                rows.append(current)
                continue
        # Description continuation lines in the item table.
        if current and len(rows)>0 and not re.match(r"^(Total Amount|GST Compensation|GST Additional|Total Tax|Grand Total|Amount in Words|S\.|No\.|Unit Taxable|Base Value|Cost|Rate Amt|Code|Item Desc|HSN|Qty|MRP|Additional|CESS|CGST|SGST/UGST|IGST|Total|Prepared By|Verified By|Authorised Signature|Terms And Conditions|Annexure)\b", line, re.I):
            if not re.search(r"PO No|PO Date|Payment Terms|Expected Delivery|PO Expiry|Vendor Name|Billing Address|Shipping Address|Terms And Conditions|Annexure", line, re.I):
                if "Brand:" not in current["Item Desc"] and len(line) <= 80 and not re.match(r"^\d{4,}\b", line):
                    current["Item Desc"]=(current["Item Desc"]+" "+line).strip()
    if not rows:
        raise ValueError("No PO item rows could be extracted from PDF")
    for r in rows:
        for c in NUM_COLS: r[c]=fmt(r.get(c,0))
        r["Item Desc"]=clean(r["Item Desc"])
    return rows

def parse_excel(path):
    wb=load_workbook(path,data_only=True,read_only=True)
    ws=wb.active
    header=None; hrow=None
    for i,row in enumerate(ws.iter_rows(min_row=1,max_row=min(100,ws.max_row),values_only=True),1):
        vals=[clean(v).lower() for v in row]
        joined=" | ".join(vals)
        if ("item code" in vals or "sku code" in vals) and ("po number" in vals or "po no" in vals or "po" in vals):
            header=vals; hrow=i; break
        if "item code" in vals and "qty" in vals:
            header=vals; hrow=i; break
    if not header:
        raise ValueError("Could not find PO item header in Excel")
    aliases={
        "sr. no.":["sr. no.","sr no","s. no","s.no","sr"],
        "item code":["item code","sku code","item"],
        "item desc":["item desc","item description","sku desc","description"],
        "hsn code":["hsn code","hsn","hsn/sac"],
        "po number":["po number","po no","po"],
        "po date":["po date","purchase order date","order date"],
        "po release date":["po release date","release date"],
        "payment terms":["payment terms","payment"],
        "expected delivery date":["expected delivery date","delivery date"],
        "po expiry date":["po expiry date","expiry date","expiry"],
        "vendor name":["vendor name","vendor","supplier"],
        "qty":["qty","quantity"],
        "mrp (inr)":["mrp (inr)","mrp"],
        "unit base cost (inr)":["unit base cost (inr)","unit base cost","unit cost","base cost"],
        "taxable value (inr)":["taxable value (inr)","taxable value"],
        "cgst rate":["cgst rate"],"cgst amount":["cgst amount","cgst amt"],
        "sgst/ugst rate":["sgst/ugst rate","sgst rate"],"sgst/ugst amount":["sgst/ugst amount","sgst amount"],
        "igst rate":["igst rate"],"igst amount":["igst amount","igst amt"],
        "cess rate":["cess rate"],"cess amount":["cess amount","cess amt"],
        "additional cess rate":["additional cess rate","add. cess rate","add cess rate"],
        "additional cess amount":["additional cess amount","add. cess amount","add cess amount"],
        "total (inr)":["total (inr)","total"]
    }
    idx={}
    for c in COLUMNS:
        for j,v in enumerate(header):
            if v in aliases.get(c.lower(),[c.lower()]):
                idx[c]=j; break
    rows=[]
    for vals in ws.iter_rows(min_row=hrow+1,values_only=True):
        if not any(v is not None and clean(v) for v in vals): continue
        item=vals[idx["Item Code"]] if "Item Code" in idx and idx["Item Code"]<len(vals) else None
        if item is None or not clean(item) or clean(item).lower()=="item code": continue
        r={c:(vals[idx[c]] if c in idx and idx[c]<len(vals) else "") for c in COLUMNS}
        for c in NUM_COLS: r[c]=fmt(num(r[c]))
        for c in COLUMNS:
            if c not in NUM_COLS: r[c]=clean(r[c])
        if not r["Sr. No."]: r["Sr. No."]=len(rows)+1
        rows.append(r)
    if not rows: raise ValueError("No PO item rows could be extracted from Excel")
    return rows

def parse_file(path):
    ext=Path(path).suffix.lower()
    if ext==".pdf": return parse_pdf(path)
    if ext in (".xlsx",".xlsm"): return parse_excel(path)
    raise ValueError("Unsupported file type")
