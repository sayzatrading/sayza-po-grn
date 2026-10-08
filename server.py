import json, sqlite3, uuid, shutil, io, re
from pathlib import Path
from datetime import datetime
from flask import Flask, request, jsonify, send_file, send_from_directory
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from grn_parser import parse_file as parse_grn
from po_parser import parse_file as parse_po

ROOT=Path(__file__).resolve().parent
DATA=ROOT/'data'; ORIGINALS=DATA/'originals'; GENERATED=DATA/'generated'; DB=DATA/'control_center.db'
for p in (ORIGINALS,GENERATED): p.mkdir(parents=True,exist_ok=True)
app=Flask(__name__,static_folder='static',static_url_path='')

PO_COLUMNS=["Sr. No.","Item Code","Item Desc","HSN Code","PO Number","PO Date","PO Release Date","Payment Terms","Expected Delivery Date","PO Expiry Date","Vendor Name","Qty","MRP (INR)","Unit Base Cost (INR)","Taxable Value (INR)","CGST Rate","CGST Amount","SGST/UGST Rate","SGST/UGST Amount","IGST Rate","IGST Amount","CESS Rate","CESS Amount","Additional CESS Rate","Additional CESS Amount","Total (INR)"]
GRN_COLUMNS=["Sr. No.","SKU Code","SKU Desc","PO Number","GRN Number","Invoice Number","Lot MRP (INR)","Exp Qty","Recv Qty","Unit Price (INR)","Taxable Value (INR)","CGST Rate","CGST Amount","SGST/UGST Rate","SGST/UGST Amount","IGST Rate","IGST Amount","CESS Rate","CESS Amount","Add. Cess Rate","Add. Cess Amount","Total (INR)"]

# ---------- DB ----------
def conn():
    c=sqlite3.connect(DB); c.row_factory=sqlite3.Row; return c

def init():
    c=conn()
    c.execute('''CREATE TABLE IF NOT EXISTS documents(
      id TEXT PRIMARY KEY, doc_type TEXT, filename TEXT, stored_name TEXT, uploaded_at TEXT,
      po_number TEXT, grn_number TEXT, invoice_number TEXT, vendor_name TEXT, meta_json TEXT,
      item_count INTEGER, status TEXT, error TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS lines(id TEXT PRIMARY KEY, doc_id TEXT, doc_type TEXT, data_json TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS appointments(po_number TEXT PRIMARY KEY, appointment_date TEXT, updated_at TEXT)''')
    c.commit(); c.close()
init()

def clean_num(v):
    try: return float(v or 0)
    except: return 0.0

def normalize_key(v): return str(v or '').strip().upper()

def save_doc(doc_type, f):
    did=uuid.uuid4().hex; safe=f'{did}_{Path(f.filename).name}'.replace('/','_').replace('\\','_')
    tmp=DATA/(did+'_tmp_'+Path(f.filename).name); f.save(tmp)
    try:
        if doc_type=='PO': rows=parse_po(str(tmp)); meta={}
        else:
            rows,meta,vendor,text=parse_grn(str(tmp)); meta=meta or {}
            if vendor: meta['vendorName']=vendor
        if not rows: raise ValueError('No item rows could be extracted')
        shutil.move(str(tmp), ORIGINALS/safe)
        if doc_type=='PO':
            po=rows[0].get('PO Number',''); grn=''; inv=''; vendor=rows[0].get('Vendor Name','')
        else:
            po=rows[0].get('PO Number',''); grn=rows[0].get('GRN Number',''); inv=rows[0].get('Invoice Number',''); vendor=meta.get('vendorName','')
        c=conn(); now=datetime.now().isoformat(timespec='seconds')
        c.execute('INSERT INTO documents VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)',(did,doc_type,f.filename,safe,now,po,grn,inv,vendor,json.dumps(meta),len(rows),'Success',''))
        for row in rows: c.execute('INSERT INTO lines VALUES(?,?,?,?)',(uuid.uuid4().hex,did,doc_type,json.dumps(row,ensure_ascii=False)))
        c.commit(); c.close()
        return {'id':did,'filename':f.filename,'type':doc_type,'status':'Success','items':len(rows),'po_number':po,'grn_number':grn}
    except Exception as e:
        try: tmp.unlink(missing_ok=True)
        except: pass
        c=conn(); now=datetime.now().isoformat(timespec='seconds')
        c.execute('INSERT INTO documents VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)',(did,doc_type,f.filename,'',now,'','','','', '{}',0,'Failed',str(e)))
        c.commit(); c.close(); return {'id':did,'filename':f.filename,'type':doc_type,'status':'Failed','error':str(e)}

# ---------- API ----------
@app.get('/')
def home(): return send_from_directory('static','index.html')
@app.get('/api/health')
def health(): return jsonify({'ok':True})

@app.post('/api/upload/<doc_type>')
def upload(doc_type):
    doc_type=doc_type.upper()
    if doc_type not in ('PO','GRN'): return jsonify({'error':'Invalid type'}),400
    files=request.files.getlist('files'); return jsonify([save_doc(doc_type,f) for f in files if f.filename])

@app.get('/api/documents/<doc_type>')
def documents(doc_type):
    c=conn(); rs=c.execute('SELECT * FROM documents WHERE doc_type=? ORDER BY uploaded_at DESC',(doc_type.upper(),)).fetchall(); c.close()
    out=[]
    for r in rs:
        d=dict(r)
        if doc_type.upper()=='PO':
            a=conn().execute('SELECT appointment_date FROM appointments WHERE po_number=?',(d['po_number'],)).fetchone()
            d['appointment_booked_on']=a['appointment_date'] if a else ''
        out.append(d)
    return jsonify(out)

@app.get('/api/lines/<doc_type>')
def lines(doc_type):
    c=conn(); rs=c.execute('SELECT l.id AS line_id,l.data_json,l.doc_id,d.filename,d.uploaded_at FROM lines l JOIN documents d ON d.id=l.doc_id WHERE l.doc_type=? AND d.status="Success" ORDER BY d.uploaded_at DESC,l.rowid',(doc_type.upper(),)).fetchall(); c.close()
    out=[]
    for r in rs:
        x=json.loads(r['data_json']); x['_line_id']=r['line_id']; x['_doc_id']=r['doc_id']; x['_filename']=r['filename']; x['_uploaded_at']=r['uploaded_at']; out.append(x)
    return jsonify(out)

@app.delete('/api/line/<line_id>')
def delete_line(line_id):
    c=conn()
    row=c.execute('SELECT id,doc_id,doc_type FROM lines WHERE id=?',(line_id,)).fetchone()
    if not row:
        c.close(); return jsonify({'error':'Row not found'}),404
    c.execute('DELETE FROM lines WHERE id=?',(line_id,))
    remaining=c.execute('SELECT COUNT(*) n FROM lines WHERE doc_id=?',(row['doc_id'],)).fetchone()['n']
    c.execute('UPDATE documents SET item_count=? WHERE id=?',(remaining,row['doc_id']))
    if remaining==0:
        # Keep the uploaded document record for audit/history, but mark it as empty.
        c.execute('UPDATE documents SET status=?, error=? WHERE id=?',('Empty','All item rows deleted',row['doc_id']))
    c.commit(); c.close()
    return jsonify({'ok':True,'line_id':line_id,'doc_id':row['doc_id'],'remaining_rows':remaining})

@app.delete('/api/document/<did>')
def delete_document(did):
    c=conn()
    d=c.execute('SELECT * FROM documents WHERE id=?',(did,)).fetchone()
    if not d:
        c.close(); return jsonify({'error':'Document not found'}),404
    # Remove parsed item rows first, then the document record.
    c.execute('DELETE FROM lines WHERE doc_id=?',(did,))
    # Appointment belongs to the PO number rather than a particular upload. Only remove it
    # when this was the last successful PO document for that PO number.
    if d['doc_type']=='PO' and d['po_number']:
        other=c.execute('SELECT COUNT(*) n FROM documents WHERE doc_type="PO" AND status="Success" AND po_number=? AND id<>?',(d['po_number'],did)).fetchone()['n']
        if other==0:
            c.execute('DELETE FROM appointments WHERE po_number=?',(d['po_number'],))
    c.execute('DELETE FROM documents WHERE id=?',(did,))
    c.commit(); c.close()
    try:
        if d['stored_name']:
            (ORIGINALS/d['stored_name']).unlink(missing_ok=True)
    except Exception:
        pass
    return jsonify({'ok':True,'id':did,'doc_type':d['doc_type'],'filename':d['filename']})

@app.post('/api/appointment')
def appointment():
    data=request.get_json(force=True); po=str(data.get('po_number','')).strip(); date=str(data.get('appointment_booked_on','')).strip()
    if not po: return jsonify({'error':'PO Number is required'}),400
    c=conn(); now=datetime.now().isoformat(timespec='seconds')
    if date: c.execute('INSERT INTO appointments(po_number,appointment_date,updated_at) VALUES(?,?,?) ON CONFLICT(po_number) DO UPDATE SET appointment_date=excluded.appointment_date,updated_at=excluded.updated_at',(po,date,now))
    else: c.execute('DELETE FROM appointments WHERE po_number=?',(po,))
    c.commit(); c.close(); return jsonify({'ok':True,'po_number':po,'appointment_booked_on':date})

def get_reconciliation():
    c=conn(); po_docs=c.execute('SELECT id,* FROM documents WHERE doc_type="PO" AND status="Success"').fetchall(); grn_docs=c.execute('SELECT id,* FROM documents WHERE doc_type="GRN" AND status="Success"').fetchall()
    po_lines={}; po_meta={}; grn_lines={}; grn_numbers={}
    def add_fin(bucket,key,row,qty_key,taxable_key,total_key):
        x=bucket.setdefault(key,{'qty':0.0,'taxable':0.0,'total':0.0})
        x['qty'] += clean_num(row.get(qty_key))
        x['taxable'] += clean_num(row.get(taxable_key))
        x['total'] += clean_num(row.get(total_key))
    for d in po_docs:
        po_key=normalize_key(d['po_number'])
        po_meta[po_key]=dict(po_number=d['po_number'],vendor=d['vendor_name'],file_id=d['id'])
        rs=c.execute('SELECT data_json FROM lines WHERE doc_id=?',(d['id'],)).fetchall()
        for rr in rs:
            r=json.loads(rr['data_json']); key=(po_key,normalize_key(r.get('Item Code')))
            add_fin(po_lines,key,r,'Qty','Taxable Value (INR)','Total (INR)')
    for d in grn_docs:
        p=normalize_key(d['po_number']); g=d['grn_number']; grn_numbers.setdefault(p,[]).append(g)
        rs=c.execute('SELECT data_json FROM lines WHERE doc_id=?',(d['id'],)).fetchall()
        for rr in rs:
            r=json.loads(rr['data_json']); key=(p,normalize_key(r.get('SKU Code')))
            add_fin(grn_lines,key,r,'Recv Qty','Taxable Value (INR)','Total (INR)')
    appointments={r['po_number']:r['appointment_date'] for r in c.execute('SELECT * FROM appointments').fetchall()}; c.close()
    keys=sorted(set(po_lines)|set(grn_lines))
    out=[]
    for key in keys:
        p,item=key; po=po_lines.get(key,{'qty':0.0,'taxable':0.0,'total':0.0}); gr=grn_lines.get(key,{'qty':0.0,'taxable':0.0,'total':0.0})
        ordered=po['qty']; received=gr['qty']; diff=ordered-received
        po_tax=po['taxable']; grn_tax=gr['taxable']; tax_diff=po_tax-grn_tax
        po_total=po['total']; grn_total=gr['total']; total_diff=po_total-grn_total
        if p not in po_meta: status='PO Not Found'
        elif key not in po_lines: status='Item Not in PO'
        elif received==0: status='Not Received'
        elif diff>0: status='Short Received'
        elif diff<0: status='Over Received'
        else: status='Matched'
        out.append({'po_number':po_meta.get(p,{}).get('po_number',p),'item_code':item,'ordered_qty':ordered,'received_qty':received,'difference_qty':diff,'po_taxable':po_tax,'grn_taxable':grn_tax,'taxable_difference':tax_diff,'po_total':po_total,'grn_total':grn_total,'total_difference':total_diff,'grn_numbers':', '.join(sorted(set(grn_numbers.get(p,[])))),'appointment_booked_on':appointments.get(po_meta.get(p,{}).get('po_number',''),'') if p in po_meta else '','status':status})
    return out

@app.get('/api/reconciliation')
def reconciliation(): return jsonify(get_reconciliation())

@app.get('/api/stats')
def stats():
    c=conn(); po=c.execute('SELECT COUNT(*) n,COALESCE(SUM(item_count),0) items FROM documents WHERE doc_type="PO" AND status="Success"').fetchone(); gr=c.execute('SELECT COUNT(*) n,COALESCE(SUM(item_count),0) items FROM documents WHERE doc_type="GRN" AND status="Success"').fetchone(); c.close()
    rec=get_reconciliation(); return jsonify({'po_files':po['n'],'po_items':po['items'],'grn_files':gr['n'],'grn_items':gr['items'],'matched':sum(x['status']=='Matched' for x in rec),'short':sum(x['status']=='Short Received' for x in rec),'over':sum(x['status']=='Over Received' for x in rec),'not_received':sum(x['status']=='Not Received' for x in rec),'po_not_found':sum(x['status']=='PO Not Found' for x in rec),'pending_qty':sum(max(0,x['difference_qty']) for x in rec)})

# ---------- exports ----------
def make_xlsx(headers,rows,name):
    wb=Workbook(); ws=wb.active; ws.title='Report'; ws.append(headers)
    for c in ws[1]: c.font=Font(bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor='1F4E78'); c.alignment=Alignment(wrap_text=True)
    for row in rows: ws.append([row.get(h,'') for h in headers])
    ws.freeze_panes='A2'; ws.auto_filter.ref=ws.dimensions
    for col in ws.columns:
        width=min(45,max(10,max(len(str(c.value or '')) for c in col)+2)); ws.column_dimensions[col[0].column_letter].width=width
    p=GENERATED/name; wb.save(p); return p

def all_lines(t):
    c=conn(); rs=c.execute('SELECT l.data_json FROM lines l JOIN documents d ON d.id=l.doc_id WHERE l.doc_type=? AND d.status="Success" ORDER BY d.uploaded_at,l.rowid',(t,)).fetchall(); c.close(); return [json.loads(x['data_json']) for x in rs]

RECON_HEADERS=['PO Number','Item Code','Ordered Qty','Received Qty','Difference Qty','PO Taxable Value (INR)','GRN Taxable Value (INR)','Taxable Difference (INR)','PO Total (INR)','GRN Total (INR)','Total Difference (INR)','GRN Number(s)','Appointment Booked On','Status']

def recon_export_rows(rows):
    """Convert internal reconciliation keys to the human-readable export headers."""
    out=[]
    mapping={
        'PO Number':'po_number','Item Code':'item_code','Ordered Qty':'ordered_qty','Received Qty':'received_qty',
        'Difference Qty':'difference_qty','PO Taxable Value (INR)':'po_taxable','GRN Taxable Value (INR)':'grn_taxable',
        'Taxable Difference (INR)':'taxable_difference','PO Total (INR)':'po_total','GRN Total (INR)':'grn_total',
        'Total Difference (INR)':'total_difference','GRN Number(s)':'grn_numbers',
        'Appointment Booked On':'appointment_booked_on','Status':'status'
    }
    for r in rows:
        out.append({h:r.get(k,'') for h,k in mapping.items()})
    return out

@app.get('/api/export/<kind>.xlsx')
def export_xlsx(kind):
    kind=kind.lower()
    if kind=='po': p=make_xlsx(PO_COLUMNS,all_lines('PO'),'PO_Consolidated.xlsx')
    elif kind=='grn': p=make_xlsx(GRN_COLUMNS,all_lines('GRN'),'GRN_Consolidated.xlsx')
    elif kind=='reconciliation': p=make_xlsx(RECON_HEADERS,recon_export_rows(get_reconciliation()),'PO_GRN_Reconciliation.xlsx')
    else:return 'Invalid',400
    return send_file(p,as_attachment=True,download_name=p.name)

@app.post('/api/export/reconciliation-visible.xlsx')
def export_reconciliation_visible_xlsx():
    payload=request.get_json(silent=True) or {}
    rows=payload.get('rows') or []
    p=make_xlsx(RECON_HEADERS,recon_export_rows(rows),'PO_GRN_Reconciliation_Visible.xlsx')
    return send_file(p,as_attachment=True,download_name=p.name)

@app.get('/api/document/<did>/original')
def original(did):
    c=conn(); r=c.execute('SELECT * FROM documents WHERE id=?',(did,)).fetchone(); c.close()
    if not r or not r['stored_name']: return 'Not found',404
    return send_file(ORIGINALS/r['stored_name'],as_attachment=True,download_name=r['filename'])

@app.get('/api/document/<did>/excel')
def doc_excel(did):
    c=conn(); d=c.execute('SELECT * FROM documents WHERE id=?',(did,)).fetchone(); rs=c.execute('SELECT data_json FROM lines WHERE doc_id=? ORDER BY rowid',(did,)).fetchall(); c.close()
    if not d:return 'Not found',404
    rows=[json.loads(x['data_json']) for x in rs]; headers=PO_COLUMNS if d['doc_type']=='PO' else GRN_COLUMNS; name=f"{d['po_number'] or d['grn_number'] or did}_{d['doc_type']}.xlsx"; p=make_xlsx(headers,rows,name); return send_file(p,as_attachment=True,download_name=p.name)

def pdf_table(title,headers,rows,path):
    styles=getSampleStyleSheet(); doc=SimpleDocTemplate(str(path),pagesize=landscape(A4),leftMargin=16,rightMargin=16,topMargin=16,bottomMargin=16)
    data=[[str(x) for x in headers]]+[[str(r.get(x,'')) for x in headers] for r in rows]
    t=Table(data,repeatRows=1); t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#1F4E78')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),.25,colors.grey),('FONTSIZE',(0,0),(-1,-1),5.5),('VALIGN',(0,0),(-1,-1),'TOP')]))
    doc.build([Paragraph(title,styles['Title']),Spacer(1,8),t])
    return path

@app.get('/api/export/po.pdf')
def po_pdf():
    p=GENERATED/'PO_Consolidated.pdf'; pdf_table('Purchase Orders - Consolidated',PO_COLUMNS,all_lines('PO'),p); return send_file(p,as_attachment=True,download_name=p.name)

@app.get('/api/export/grn.pdf')
def grn_pdf():
    p=GENERATED/'GRN_Consolidated.pdf'; pdf_table('GRNs - Consolidated',GRN_COLUMNS,all_lines('GRN'),p); return send_file(p,as_attachment=True,download_name=p.name)

@app.get('/api/export/reconciliation.pdf')
def recon_pdf():
    p=GENERATED/'PO_GRN_Reconciliation.pdf'; pdf_table('PO ↔ GRN Reconciliation',RECON_HEADERS,recon_export_rows(get_reconciliation()),p); return send_file(p,as_attachment=True,download_name=p.name)

@app.post('/api/export/reconciliation-visible.pdf')
def export_reconciliation_visible_pdf():
    payload=request.get_json(silent=True) or {}
    rows=payload.get('rows') or []
    p=GENERATED/'PO_GRN_Reconciliation_Visible.pdf'
    pdf_table('PO ↔ GRN Reconciliation - Visible Rows',RECON_HEADERS,recon_export_rows(rows),p)
    return send_file(p,as_attachment=True,download_name=p.name)

@app.get('/api/document/<did>/pdf')
def doc_pdf(did):
    c=conn(); d=c.execute('SELECT * FROM documents WHERE id=?',(did,)).fetchone(); rs=c.execute('SELECT data_json FROM lines WHERE doc_id=? ORDER BY rowid',(did,)).fetchall(); c.close()
    if not d:return 'Not found',404
    rows=[json.loads(x['data_json']) for x in rs]; headers=PO_COLUMNS if d['doc_type']=='PO' else GRN_COLUMNS; base=d['po_number'] or d['grn_number'] or did; p=GENERATED/f'{base}_{d["doc_type"]}.pdf'; pdf_table(f'{d["doc_type"]} {base}',headers,rows,p); return send_file(p,as_attachment=True,download_name=p.name)

if __name__=='__main__': app.run(host='127.0.0.1',port=3002,debug=False)
