import io, re, unicodedata
from typing import Dict
import numpy as np
import pandas as pd
from openpyxl import load_workbook

NEEDED = ['SHAREPOINT','LOGS','BOPS','SL','DEFINITION BOOK']
REPORTS = ['LOGS','BOPS','SL']
ALIASES = {
 'entity':['entity','entidade','unit name','unidade','bu','location','plant','dc'],
 'kpi':['kpi code','kpi_code','codigo kpi','código kpi','cod kpi'],
 'value':['current_value ac','current value ac','valor anaplan','valor origem','actual','ac','value','valor'],
 'sp_value':['valor sharepoint','sharepoint value','value sharepoint'],
 'status':['check manual','check','status','resultado'],
 'kpi_name':['kpi name','kpi_name','nome kpi','description','descrição','descricao'],
 'formula':['formula','fórmula','calculation rule','regra'],
 'uom':['unit_of_measure','unit of measure','uom','unidade de medida'],
 'owner':['owner','responsavel','responsável']
}

def noacc(x): return ''.join(c for c in unicodedata.normalize('NFKD',str(x)) if not unicodedata.combining(c))
def norm(x):
    if pd.isna(x): return ''
    return re.sub(r'\s+',' ',noacc(x).strip().upper())
def ncol(x): return re.sub(r'\s+',' ',re.sub(r'[^a-z0-9]+',' ',noacc(x).strip().lower())).strip()
def nkpi(x):
    s=norm(x); m=re.search(r'\b([A-Z]{2,3}-[KR]\d{3,5}(?:_\d{4})?)\b',s)
    return m.group(1) if m else s
def find_col(df,role):
    cols={ncol(c):c for c in df.columns}
    for a in ALIASES[role]:
        if ncol(a) in cols:return cols[ncol(a)]
    candidates=[]
    for key,col in cols.items():
        score=max((len(ncol(a)) for a in ALIASES[role] if ncol(a) in key),default=0)
        if score:candidates.append((score,col))
    return max(candidates)[1] if candidates else None
def match_sheet(names,target):
    t=norm(target)
    exact=[n for n in names if norm(n)==t]
    if exact:return exact[0]
    aliases={'DEFINITION BOOK':['DEFINITION BOOK','DEFINITIONBOOK','DEFINITIONS']}.get(t,[t])
    for a in aliases:
        found=[n for n in names if norm(a) in norm(n)]
        if found:return found[0]
    return None

def detect_header(raw,max_rows=80):
    best=(0,-1)
    for i in range(min(max_rows,len(raw))):
        vals=[norm(v) for v in raw.iloc[i].tolist()]
        has_kpi=any(('KPI' in v and ('CODE' in v or 'CODIGO' in v)) for v in vals)
        has_entity=any(v in {'ENTITY','ENTIDADE','UNIT NAME','UNIDADE','BU','LOCATION'} or 'ENTITY' in v for v in vals)
        has_value=any(v in {'AC','VALUE','VALOR','CURRENT_VALUE AC','CURRENT VALUE AC','VALOR ANAPLAN'} for v in vals)
        score=1000*has_kpi+800*has_entity+200*has_value+sum(bool(v) for v in vals)
        if score>best[1]:best=(i,score)
    return best[0]
def read_sheet(data,sheet,header=None):
    raw=pd.read_excel(io.BytesIO(data),sheet_name=sheet,header=None,engine='openpyxl')
    h=detect_header(raw) if header is None else header
    cols=[];seen={}
    for j,v in enumerate(raw.iloc[h].tolist()):
        c=str(v).strip() if not pd.isna(v) else f'COL_{j+1}'
        if c in seen:seen[c]+=1;c=f'{c}_{seen[c]}'
        else:seen[c]=1
        cols.append(c)
    df=raw.iloc[h+1:].copy();df.columns=cols
    return df.dropna(how='all').reset_index(drop=True),h

def to_num(s):
    if pd.api.types.is_numeric_dtype(s):return pd.to_numeric(s,errors='coerce')
    x=s.astype(str).str.strip().replace({'':np.nan,'-':np.nan,'nan':np.nan,'None':np.nan})
    br=x.str.contains(r'^-?\d{1,3}(?:\.\d{3})+,\d+(?:[Ee][+-]?\d+)?$',regex=True,na=False)
    x.loc[br]=x.loc[br].str.replace('.','',regex=False).str.replace(',','.',regex=False)
    comma=x.str.contains(r'^-?\d+,\d+(?:[Ee][+-]?\d+)?$',regex=True,na=False)
    x.loc[comma]=x.loc[comma].str.replace(',','.',regex=False)
    return pd.to_numeric(x,errors='coerce')
def auto_map(df,source):
    m={r:find_col(df,r) for r in ALIASES}
    # SharePoint must prefer CURRENT_VALUE AC. Reports prefer AC/Valor Anaplan, never Valor Sharepoint.
    normalized={ncol(c):c for c in df.columns}
    if source=='SHAREPOINT':
        for key in ['current value ac','current value','ac','value','valor']:
            if key in normalized:m['value']=normalized[key];break
    else:
        for key in ['valor anaplan','ac','actual','valor origem','value','valor']:
            if key in normalized:m['value']=normalized[key];break
    return m

def standard(df,m,source):
    if not m.get('entity') or not m.get('kpi') or not m.get('value'):
        missing=[x for x in ['entity','kpi','value'] if not m.get(x)]
        raise ValueError(f'{source}: colunas não localizadas: {", ".join(missing)}. Detectadas: {list(df.columns)}')
    out=pd.DataFrame({
      'SOURCE':source,
      'ENTITY':df[m['entity']].map(norm),
      'KPI_CODE':df[m['kpi']].map(nkpi),
      'VALUE':to_num(df[m['value']]),
      'SOURCE_ROW':df.index+2,
    })
    out['KPI_NAME']=df[m['kpi_name']].astype(str) if m.get('kpi_name') else ''
    out['EXISTING_SP_VALUE']=to_num(df[m['sp_value']]) if m.get('sp_value') else np.nan
    out['EXISTING_STATUS']=df[m['status']].astype(str) if m.get('status') else ''
    out=out[(out.ENTITY!='')&(out.KPI_CODE!='')]
    out['KEY']=out.ENTITY+'|'+out.KPI_CODE
    return out

def definitions(df,m):
    if not m.get('kpi'):return pd.DataFrame(columns=['KPI_CODE'])
    d=pd.DataFrame({'KPI_CODE':df[m['kpi']].map(nkpi)})
    d['DEF_KPI_NAME']=df[m['kpi_name']].astype(str) if m.get('kpi_name') else ''
    d['DEF_FORMULA']=df[m['formula']].astype(str).replace('nan','') if m.get('formula') else ''
    d['DEF_UOM']=df[m['uom']].astype(str) if m.get('uom') else ''
    d['DEF_OWNER']=df[m['owner']].astype(str) if m.get('owner') else ''
    return d[d.KPI_CODE!=''].drop_duplicates('KPI_CODE')
def compare(rep,sp,defs,atol,rtol):
    ra=rep.groupby(['ENTITY','KPI_CODE'],dropna=False).agg(REPORT_VALUE=('VALUE','sum'),REPORT_ROWS=('VALUE','size'),KPI_NAME=('KPI_NAME','first'),SOURCE=('SOURCE','first'),EXISTING_SP_VALUE=('EXISTING_SP_VALUE','first'),EXISTING_STATUS=('EXISTING_STATUS','first')).reset_index()
    sa=sp.groupby(['ENTITY','KPI_CODE'],dropna=False).agg(SHAREPOINT_VALUE=('VALUE','sum'),SHAREPOINT_ROWS=('VALUE','size'),SP_KPI_NAME=('KPI_NAME','first')).reset_index()
    x=ra.merge(sa,on=['ENTITY','KPI_CODE'],how='outer',indicator=True)
    x['KPI_NAME']=x.KPI_NAME.combine_first(x.SP_KPI_NAME)
    x=x.merge(defs,on='KPI_CODE',how='left')
    x['DIFFERENCE']=x.REPORT_VALUE-x.SHAREPOINT_VALUE
    x['DIFFERENCE_PCT']=np.where(x.SHAREPOINT_VALUE.abs()>atol,x.DIFFERENCE/x.SHAREPOINT_VALUE,np.nan)
    close=(x.DIFFERENCE.abs()<=atol+rtol*x.SHAREPOINT_VALUE.abs()).fillna(False)
    x['STATUS']=np.select([x._merge.eq('left_only'),x._merge.eq('right_only'),close],['NÃO ESTÁ NO SHAREPOINT','SOMENTE NO SHAREPOINT','OK'],default='DIVERGENTE')
    x['DUPLICATE_ALERT']=np.where((x.REPORT_ROWS.fillna(0)>1)|(x.SHAREPOINT_ROWS.fillna(0)>1),'CHAVE REPETIDA - VALORES SOMADOS','')
    x['KEY_USED']='ENTITY + KPI_CODE'
    return x.drop(columns=['_merge','SP_KPI_NAME'])

def run(data,atol=.01,rtol=.0001):
    xls=pd.ExcelFile(io.BytesIO(data),engine='openpyxl'); names=xls.sheet_names
    resolved={t:match_sheet(names,t) for t in NEEDED}
    required=['SHAREPOINT','LOGS','BOPS','SL']
    absent=[t for t in required if not resolved[t]]
    if absent:return {'error':'Abas não encontradas: '+', '.join(absent),'resolved':resolved,'names':names}
    raws={};maps={};headers={};std={}
    # Only five useful sheets are read. No full-workbook inventory and no second formula pass.
    for t,sheet in resolved.items():
        if not sheet:continue
        df,h=read_sheet(data,sheet);raws[t]=df;headers[t]=h+1;maps[t]=auto_map(df,t)
    try:
        sp=standard(raws['SHAREPOINT'],maps['SHAREPOINT'],'SHAREPOINT');std['SHAREPOINT']=sp
        defs=definitions(raws['DEFINITION BOOK'],maps['DEFINITION BOOK']) if 'DEFINITION BOOK' in raws else pd.DataFrame(columns=['KPI_CODE'])
        results=[]
        for t in REPORTS:
            frame=standard(raws[t],maps[t],t);std[t]=frame;results.append(compare(frame,sp,defs,atol,rtol))
    except ValueError as e:return {'error':str(e),'resolved':resolved,'names':names,'raws':raws,'maps':maps,'headers':headers}
    return {'error':None,'resolved':resolved,'names':names,'raws':raws,'maps':maps,'headers':headers,'standardized':std,'definitions':defs,'results':results}
def excel_report(p):
    out=io.BytesIO(); allr=pd.concat(p['results'],ignore_index=True)
    with pd.ExcelWriter(out,engine='xlsxwriter') as w:
        pd.DataFrame([{'TYPE':k,'SHEET':v,'HEADER_ROW':p['headers'].get(k),'ENTITY_COL':p['maps'].get(k,{}).get('entity'),'KPI_COL':p['maps'].get(k,{}).get('kpi'),'VALUE_COL':p['maps'].get(k,{}).get('value')} for k,v in p['resolved'].items()]).to_excel(w,index=False,sheet_name='Mapeamento')
        allr.to_excel(w,index=False,sheet_name='Comparacao_Completa')
        allr[allr.STATUS=='OK'].to_excel(w,index=False,sheet_name='OK')
        allr[allr.STATUS!='OK'].to_excel(w,index=False,sheet_name='Achados')
        allr.groupby(['SOURCE','STATUS']).size().reset_index(name='QUANTIDADE').to_excel(w,index=False,sheet_name='Resumo')
        allr.groupby(['KPI_CODE','STATUS']).size().reset_index(name='QUANTIDADE').to_excel(w,index=False,sheet_name='Resumo_KPI')
        allr.groupby(['ENTITY','STATUS']).size().reset_index(name='QUANTIDADE').to_excel(w,index=False,sheet_name='Resumo_Entidade')
        p['definitions'].to_excel(w,index=False,sheet_name='Definition_Book')
        header=w.book.add_format({'bold':True,'bg_color':'#001E60','font_color':'white'})
        for ws in w.sheets.values():ws.freeze_panes(1,0);ws.set_row(0,24,header);ws.set_column(0,30,22)
    return out.getvalue()
