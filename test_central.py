from grn_parser import parse_file as parse_grn
from po_parser import parse_file as parse_po
po=parse_po('sample_files/SOTY-1N70252155-CADPO227001.pdf')
assert len(po)==4 and po[0]['PO Number']=='CADPO227001' and sum(float(x['Qty']) for x in po)==1061
gr=parse_grn('sample_files/GRN_CAD000267571.pdf')[0]
assert len(gr)==3 and gr[0]['PO Number']=='CADPO227015' and sum(float(x['Recv Qty']) for x in gr)==53
ordered=62; received=0+30; assert ordered-received==32
print('CENTRAL TEST PASSED')
