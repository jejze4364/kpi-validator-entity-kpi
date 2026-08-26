import io,re,unicodedata
import numpy as np,pandas as pd
NEEDED=['SHAREPOINT','LOGS','BOPS','SL','DEFINITION BOOK']; REPORTS=['LOGS','BOPS','SL']
ALIASES={'entity':['entity','entidade','unit name','unidade','bu','location','plant','dc'],'kpi':['kpi code','kpi_code','codigo kpi','código kpi','cod kpi'],'value':['current_value ac','current value ac','valor anaplan','valor origem','actual','ac','value','valor'],'kpi_name':['kpi name','nome kpi','description','descrição'],'formula':['formula','fórmula','regra'],'uom':['unit_of_measure','unit of measure','uom'],'owner':['owner','responsavel','responsável']}
def noacc(x):return ''.join(c for c in unicodedata.normalize('NFKD',str(x)) if not unicodedata.combining(c))
def norm(x):return '' if pd.isna(x) else re.sub(r'\s+',' ',noacc(x).strip().upper())
def nc(x):return re.sub(r'\s+',' ',re.sub(r'[^a-z0-9]+',' ',noacc(x).strip().lower())).strip()
def nkpi(x):
 s=norm(x);m=re.search(r'\b([A-Z]{2,3}-[KR]\d{3,5}(?:_\d{4})?)\b',s);return m.group(1) if m else s
def sheet(names,t):
 x=[n for n in names if norm(n)==norm(t)];
 if x:return x[0]
 aliases=['DEFINITION BOOK','DEFINITIONBOOK','DEFINITIONS'] if t=='DEFINITION BOOK' else [t]
 for a in aliases:
  x=[n for n in names if norm(a) in norm(n)]
  if x:return x[0]
def header(raw):
 best=(0,-1)
 for i in range(min(80,len(raw))):
  v=[norm(x) for x in raw.iloc[i]];score=1000*any('KPI' in x and ('CODE' in x or 'CODIGO' in x) for x in v)+800*any(x in ['ENTITY','ENTIDADE','UNIT NAME','UNIDADE','BU','LOCATION'] for x in v)+sum(bool(x) for x in v)
  if score>best[1]:best=(i,score)
 return best[0]
def read(data,s):
 raw=pd.read_excel(io.BytesIO(data),sheet_name=s,header=None,engine='openpyxl');h=header(raw);cols=[];seen={}
 for j,x in enumerate(raw.iloc[h]):
  c=str(x).strip() if not pd.isna(x) else f'COL_{j+1}';seen[c]=seen.get(c,0)+1;c=f'{c}_{seen[c]}' if seen[c]>1 else c;cols.append(c)
 d=raw.iloc[h+1:].copy();d.columns=cols;return d.dropna(how='all').reset_index(drop=True),h+1
def find(d,r):
 c={nc(x):x for x in d.columns}
 for a in ALIASES[r]:
  if nc(a) in c:return c[nc(a)]
 for k,v in c.items():
  if any(nc(a) in k for a in ALIASES[r]):return v
def mapping(d,src):
 m={r:find(d,r) for r in ALIASES};c={nc(x):x for x in d.columns}
 pref=['current value ac','current value','ac','value','valor'] if src=='SHAREPOINT' else ['valor anaplan','ac','actual','valor origem','value','valor']
 for k in pref:
  if k in c:m['value']=c[k];break
 return m
def num(s):
 if pd.api.types.is_numeric_dtype(s):return pd.to_numeric(s,errors='coerce')
 x=s.astype(str).str.strip().replace({'':np.nan,'-':np.nan,'nan':np.nan});br=x.str.contains(r'^-?\d{1,3}(?:\.\d{3})+,\d+$',regex=True,na=False);x.loc[br]=x.loc[br].str.replace('.','',regex=False).str.replace(',','.',regex=False);x.loc[~br]=x.loc[~br].str.replace(',','.',regex=False);return pd.to_numeric(x,errors='coerce')
def std(d,m,src):
 miss=[x for x in ['entity','kpi','value'] if not m.get(x)]
 if miss:raise ValueError(f'{src}: colunas não localizadas: {miss}. Detectadas: {list(d.columns)}')
 o=pd.DataFrame({'SOURCE':src,'ENTITY':d[m['entity']].map(norm),'KPI_CODE':d[m['kpi']].map(nkpi),'VALUE':num(d[m['value']])});o['KPI_NAME']=d[m['kpi_name']].astype(str) if m.get('kpi_name') else '';return o[(o.ENTITY!='')&(o.KPI_CODE!='')]
def defs(d,m):
 if not m.get('kpi'):return pd.DataFrame(columns=['KPI_CODE'])
 o=pd.DataFrame({'KPI_CODE':d[m['kpi']].map(nkpi)});o['DEF_KPI_NAME']=d[m['kpi_name']].astype(str) if m.get('kpi_name') else '';o['DEF_FORMULA']=d[m['formula']].astype(str) if m.get('formula') else '';o['DEF_UOM']=d[m['uom']].astype(str) if m.get('uom') else '';o['DEF_OWNER']=d[m['owner']].astype(str) if m.get('owner') else '';return o[o.KPI_CODE!=''].drop_duplicates('KPI_CODE')
def compare(r,sp,de,at,rt):
 a=r.groupby(['ENTITY','KPI_CODE']).agg(REPORT_VALUE=('VALUE','sum'),REPORT_ROWS=('VALUE','size'),KPI_NAME=('KPI_NAME','first'),SOURCE=('SOURCE','first')).reset_index();b=sp.groupby(['ENTITY','KPI_CODE']).agg(SHAREPOINT_VALUE=('VALUE','sum'),SHAREPOINT_ROWS=('VALUE','size')).reset_index();x=a.merge(b,on=['ENTITY','KPI_CODE'],how='outer',indicator=True).merge(de,on='KPI_CODE',how='left');x['DIFFERENCE']=x.REPORT_VALUE-x.SHAREPOINT_VALUE;x['DIFFERENCE_PCT']=np.where(x.SHAREPOINT_VALUE.abs()>at,x.DIFFERENCE/x.SHAREPOINT_VALUE,np.nan);ok=(x.DIFFERENCE.abs()<=at+rt*x.SHAREPOINT_VALUE.abs()).fillna(False);x['STATUS']=np.select([x._merge.eq('left_only'),x._merge.eq('right_only'),ok],['NÃO ESTÁ NO SHAREPOINT','SOMENTE NO SHAREPOINT','OK'],default='DIVERGENTE');x['KEY_USED']='ENTITY + KPI_CODE';return x.drop(columns='_merge')
def run(data,at=.01,rt=.0001):
 names=pd.ExcelFile(io.BytesIO(data),engine='openpyxl').sheet_names;res={t:sheet(names,t) for t in NEEDED};missing=[t for t in REPORTS+['SHAREPOINT'] if not res[t]]
 if missing:return {'error':'Abas não encontradas: '+', '.join(missing),'resolved':res}
 raw={};maps={};heads={}
 for t,s in res.items():
  if s:raw[t],heads[t]=read(data,s);maps[t]=mapping(raw[t],t)
 try:
  sp=std(raw['SHAREPOINT'],maps['SHAREPOINT'],'SHAREPOINT');de=defs(raw['DEFINITION BOOK'],maps['DEFINITION BOOK']) if 'DEFINITION BOOK' in raw else pd.DataFrame(columns=['KPI_CODE']);results=[compare(std(raw[t],maps[t],t),sp,de,at,rt) for t in REPORTS]
 except ValueError as e:return {'error':str(e),'resolved':res,'maps':maps}
 return {'error':None,'resolved':res,'maps':maps,'heads':heads,'definitions':de,'results':results}
def report(p):
 out=io.BytesIO();x=pd.concat(p['results'],ignore_index=True)
 with pd.ExcelWriter(out,engine='xlsxwriter') as w:
  x.to_excel(w,index=False,sheet_name='Comparacao_Completa');x[x.STATUS=='OK'].to_excel(w,index=False,sheet_name='OK');x[x.STATUS!='OK'].to_excel(w,index=False,sheet_name='Achados');x.groupby(['SOURCE','STATUS']).size().reset_index(name='QUANTIDADE').to_excel(w,index=False,sheet_name='Resumo');p['definitions'].to_excel(w,index=False,sheet_name='Definition_Book')
 return out.getvalue()
