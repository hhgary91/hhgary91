import json, datetime, os
HERE = os.path.dirname(os.path.abspath(__file__))
from openpyxl import load_workbook
wb = load_workbook(os.path.join(HERE,'..','Design_Library_Register.xlsx'), data_only=True)
def val(v):
    if isinstance(v,(datetime.date,datetime.datetime)): return v.strftime('%Y-%m-%d')
    return v
regs={}
for sh in ['Drain_Register','RW_Register','Embedded_Wall_Register','Manhole_Register','ERSS_Pit_Register','Linkway_Register','Other_Register']:
    ws=wb[sh]; hdr=[ws.cell(4,c).value for c in range(1,ws.max_column+1)]
    rows=[]
    for r in range(5,ws.max_row+1):
        if not ws.cell(r,1).value: continue
        d={h:val(ws.cell(r,c+1).value) for c,h in enumerate(hdr) if h and h!='Selector Key' and ws.cell(r,c+1).value not in (None,'')}
        rows.append(d)
    regs[sh]=rows
crit=wb['SG_Criteria']; C={}
for r in range(4,crit.max_row+1):
    C[crit.cell(r,1).value]=crit.cell(r,2).value
criteria={'fb':C['Freeboard used for capacity (fraction of depth)'],'n':C["Manning's n, concrete"],'vmin':C['Min self-cleansing velocity'],
          'vmax':C['Max velocity, concrete-lined drain'],'root':C['Library root (normalised)']}
lw=wb['Lists']; lists={}
def listcol(name):
    for c in range(1,lw.max_column+1):
        if lw.cell(3,c).value==name:
            out=[];r=4
            while lw.cell(r,c).value: out.append(lw.cell(r,c).value); r+=1
            return out
lists={'Drain_Register':listcol('DrainType'),'RW_Register':listcol('WallType'),'Embedded_Wall_Register':listcol('EmbType'),
       'Manhole_Register':listcol('MHType'),'ERSS_Pit_Register':listcol('ERSSType'),'Linkway_Register':listcol('LkType')}
tp=wb['Tender_Pack']; items=[]
for r in range(6,61):
    if tp.cell(r,2).value: items.append({'id':tp.cell(r,2).value,'loc':tp.cell(r,3).value or '','qty':tp.cell(r,4).value,'sheet':tp.cell(r,5).value})
data={'reg':regs,'criteria':criteria,'lists':lists,'tender':{'name':tp['B3'].value,'items':items}}
html=open(os.path.join(HERE,'tpl_css.html')).read()+open(os.path.join(HERE,'tpl_body.html')).read().replace('/*DATA*/',json.dumps(data,ensure_ascii=False))
open(os.path.join(HERE,'..','design-selector-prototype.html'),'w').write(html)
print(len(html), {k:len(v) for k,v in regs.items()})
