import io,re,unicodedata,zipfile,hashlib
from datetime import datetime
import numpy as np
import pandas as pd

REPORTS=['LOGS','BOPS','SL']
ALIASES={
'entity':['entity','entidade','unit name','unidade','location','plant','dc'],
'kpi':['kpi code','kpi_code','codigo kpi','cod kpi'],
'value':['current_value ac','current value ac','current value','valor anaplan','valor origem','actual','ac','value','valor'],
'kpi_name':['kpi name','nome kpi','description','descricao'],
'formula':['formula','regra'],'uom':['unit_of_measure','unit of measure','uom','unidade de medida'],
'owner':['owner','responsavel'],'classification':['classificacao','classification','categoria','category'],
'source':['source','origem','grupo','business area','theme']}

def noacc(v):return ''.join(c for c in unicodedata.normalize('NFKD',str(v)) if not unicodedata.combining(c))
def norm(v):return '' if pd.isna(v) else re.sub(r'\s+',' ',noacc(v).strip().upper())
def nc(v):return re.sub(r'\s+',' ',re.sub('[^a-z0-9]+',' ',noacc(v).lower())).strip()
def nkpi(v):
 s=norm(v);m=re.search(r'\b([A-Z]{2,3}-[KR]\d{3,5}(?:_\d{4})?)\b',s);return m.group(1) if m else s
def safe(v):return re.sub('[^a-z0-9]+','_',noacc(v).lower()).strip('_')[:100] or 'sem_responsavel'
def locate(names,target):
 exact=[n for n in names if norm(n)==norm(target)]
 if exact:return exact[0]
 opts={'SHAREPOINT':['SHAREPOINT','SHARE POINT'],'DEFINITION BOOK':['DEFINITION BOOK','DEFINITIONBOOK','DEFINITIONS'],'LOGS':['LOGS','LOGISTICS','LOGISTICA'],'BOPS':['BOPS'],'SL':['SL','SERVICE LEVEL']}.get(target,[target])
 for a in opts:
  found=[n for n in names if norm(a) in norm(n)]
  if found:return found[0]
def header(raw):
 best=(0,-1)
 for i in range(min(80,len(raw))):
  vals=[norm(x) for x in raw.iloc[i]]
  score=1000*any('KPI' in x and ('CODE' in x or 'COD' in x) for x in vals)+800*any(x in ['ENTITY','ENTIDADE','UNIT NAME','UNIDADE','LOCATION','PLANT','DC'] for x in vals)+sum(bool(x) for x in vals)
  if score>best[1]:best=(i,score)
 return best[0]
def read(data,sheet):
 raw=pd.read_excel(io.BytesIO(data),sheet_name=sheet,header=None,engine='openpyxl');h=header(raw);cols=[];seen={}
 for j,x in enumerate(raw.iloc[h]):
  c=str(x).strip() if not pd.isna(x) else f'COL_{j+1}';seen[c]=seen.get(c,0)+1;cols.append(f'{c}_{seen[c]}' if seen[c]>1 else c)
 d=raw.iloc[h+1:].copy();d.columns=cols;return d.dropna(how='all').reset_index(drop=True),h+1
def find(d,role):
 cols={nc(x):x for x in d.columns}
 for a in ALIASES[role]:
  if nc(a) in cols:return cols[nc(a)]
 for a in ALIASES[role]:
  for k,v in cols.items():
   if nc(a) in k:return v
def mapping(d,src):
 m={r:find(d,r) for r in ALIASES};cols={nc(x):x for x in d.columns}
 pref=['current value ac','current_value ac','current value','ac','value','valor'] if src=='SHAREPOINT' else ['valor anaplan','valor origem','actual','ac','value','valor']
 for p in pref:
  if nc(p) in cols:m['value']=cols[nc(p)];break
 return m
def num(s):
 if pd.api.types.is_numeric_dtype(s):return pd.to_numeric(s,errors='coerce')
 x=s.astype(str).str.strip().replace({'':np.nan,'-':np.nan,'nan':np.nan});br=x.str.match(r'^-?\d{1,3}(?:\.\d{3})+(?:,\d+)?$',na=False)|x.str.match(r'^-?\d+,\d+$',na=False);x.loc[br]=x.loc[br].str.replace('.','',regex=False).str.replace(',','.',regex=False);x.loc[~br]=x.loc[~br].str.replace(',','',regex=False);return pd.to_numeric(x,errors='coerce')
def col(d,c):return d[c].fillna('').astype(str) if c else pd.Series(['']*len(d),index=d.index)
def standard(d,m,src,need_value=True):
 missing=[r for r in ['entity','kpi']+(['value'] if need_value else []) if not m.get(r)]
 if missing:raise ValueError(f'{src}: colunas não localizadas: {missing}. Detectadas: {list(d.columns)}')
 o=pd.DataFrame({'SOURCE':col(d,m.get('source')).map(norm),'ENTITY':d[m['entity']].map(norm),'KPI_CODE':d[m['kpi']].map(nkpi),'KPI_NAME':col(d,m.get('kpi_name')).str.strip(),'CLASSIFICATION':col(d,m.get('classification')).str.strip(),'FORMULA':col(d,m.get('formula')).str.strip(),'UNIT_OF_MEASURE':col(d,m.get('uom')).str.strip(),'OWNER':col(d,m.get('owner')).str.strip(),'VALUE':num(d[m['value']]) if m.get('value') else np.nan})
 if src in REPORTS:o['SOURCE']=src
 o=o[(o.ENTITY!='')&(o.KPI_CODE!='')].copy();prefix=o.KPI_CODE.str.extract(r'^(SL|BOPS|LOGS)',expand=False).fillna('');o.loc[~o.SOURCE.isin(REPORTS),'SOURCE']=prefix;return o.reset_index(drop=True)
def definitions(d,m):
 if not m.get('kpi'):return pd.DataFrame(columns=['KPI_CODE','DEF_KPI_NAME','DEF_CLASSIFICATION','DEF_FORMULA','DEF_UOM','DEF_OWNER','DEF_SOURCE'])
 o=pd.DataFrame({'KPI_CODE':d[m['kpi']].map(nkpi),'DEF_KPI_NAME':col(d,m.get('kpi_name')).str.strip(),'DEF_CLASSIFICATION':col(d,m.get('classification')).str.strip(),'DEF_FORMULA':col(d,m.get('formula')).str.strip(),'DEF_UOM':col(d,m.get('uom')).str.strip(),'DEF_OWNER':col(d,m.get('owner')).str.strip(),'DEF_SOURCE':col(d,m.get('source')).map(norm)});return o[o.KPI_CODE!=''].drop_duplicates('KPI_CODE')
def load_reference(data,filename='Consolidador 2.0.xlsm',modified_at=None):
 names=pd.ExcelFile(io.BytesIO(data),engine='openpyxl').sheet_names;resolved={x:locate(names,x) for x in ['SHAREPOINT','DEFINITION BOOK']}
 if not resolved['SHAREPOINT']:raise ValueError(f'Aba SHAREPOINT não encontrada. Abas disponíveis: {names}')
 raw={};heads={};maps={};raw['SHAREPOINT'],heads['SHAREPOINT']=read(data,resolved['SHAREPOINT']);maps['SHAREPOINT']=mapping(raw['SHAREPOINT'],'SHAREPOINT');ref=standard(raw['SHAREPOINT'],maps['SHAREPOINT'],'SHAREPOINT')
 if resolved['DEFINITION BOOK']:
  raw['DEFINITION BOOK'],heads['DEFINITION BOOK']=read(data,resolved['DEFINITION BOOK']);maps['DEFINITION BOOK']=mapping(raw['DEFINITION BOOK'],'DEFINITION BOOK');defs=definitions(raw['DEFINITION BOOK'],maps['DEFINITION BOOK'])
 else:defs=definitions(pd.DataFrame(),{})
 if not defs.empty:
  ref=ref.merge(defs,on='KPI_CODE',how='left')
  for a,b in [('KPI_NAME','DEF_KPI_NAME'),('CLASSIFICATION','DEF_CLASSIFICATION'),('FORMULA','DEF_FORMULA'),('UNIT_OF_MEASURE','DEF_UOM'),('OWNER','DEF_OWNER'),('SOURCE','DEF_SOURCE')]:ref[a]=ref[a].fillna('').where(ref[a].fillna('').astype(str).str.strip().ne(''),ref[b].fillna(''))
  ref=ref[['SOURCE','ENTITY','KPI_CODE','KPI_NAME','CLASSIFICATION','FORMULA','UNIT_OF_MEASURE','OWNER','VALUE']]
 ref['OWNER_KEY']=ref.OWNER.map(norm);ref['CLASSIFICATION_KEY']=ref.CLASSIFICATION.map(norm)
 return {'filename':filename,'modified_at':modified_at,'fingerprint':hashlib.sha256(data).hexdigest(),'reference_bytes':data,'resolved':resolved,'heads':heads,'maps':maps,'reference_data':ref,'definitions':defs}
def get_kpi_catalog(b):return b['reference_data'][['SOURCE','ENTITY','KPI_CODE','KPI_NAME','CLASSIFICATION','FORMULA','UNIT_OF_MEASURE','OWNER']].drop_duplicates().sort_values(['SOURCE','ENTITY','KPI_CODE']).reset_index(drop=True)
def get_responsibles(b):
 d=b['reference_data'];rows=[]
 for key,g in d.groupby('OWNER_KEY',dropna=False):rows.append({'key':key,'display_name':g.OWNER.replace('','Sem responsável').mode().iloc[0] if len(g) else 'Sem responsável','kpi_count':g.KPI_CODE.nunique(),'entity_count':g.ENTITY.nunique()})
 return pd.DataFrame(rows).sort_values('display_name').reset_index(drop=True)
def get_classifications(b):
 d=b['reference_data'];x=d[d.CLASSIFICATION_KEY!=''].groupby('CLASSIFICATION_KEY').agg(display_name=('CLASSIFICATION','first'),kpi_count=('KPI_CODE','nunique'),entity_count=('ENTITY','nunique')).reset_index().rename(columns={'CLASSIFICATION_KEY':'key'});return {'available':not x.empty,'data':x}
def get_reference_summary(b):
 d=b['reference_data'];return {'total_kpis':d.KPI_CODE.nunique(),'total_entities':d.ENTITY.nunique(),'total_responsibles':d.loc[d.OWNER_KEY!='','OWNER_KEY'].nunique(),'total_classifications':d.loc[d.CLASSIFICATION_KEY!='','CLASSIFICATION_KEY'].nunique(),'logs_kpis':d.loc[d.SOURCE=='LOGS','KPI_CODE'].nunique(),'bops_kpis':d.loc[d.SOURCE=='BOPS','KPI_CODE'].nunique(),'sl_kpis':d.loc[d.SOURCE=='SL','KPI_CODE'].nunique(),'without_owner':d.loc[d.OWNER_KEY=='','KPI_CODE'].nunique(),'without_classification':d.loc[d.CLASSIFICATION_KEY=='','KPI_CODE'].nunique()}
def compare(r,ref,at,rt,src):
 a=r.groupby(['ENTITY','KPI_CODE'],as_index=False).agg(REPORT_VALUE=('VALUE',lambda s:s.sum(min_count=1)),REPORT_ROWS=('VALUE','size'),KPI_NAME=('KPI_NAME','first'));rr=ref[(ref.SOURCE==src)|(ref.SOURCE=='')];b=rr.groupby(['ENTITY','KPI_CODE'],as_index=False).agg(SHAREPOINT_VALUE=('VALUE',lambda s:s.sum(min_count=1)),SHAREPOINT_ROWS=('VALUE','size'),CLASSIFICATION=('CLASSIFICATION','first'),FORMULA=('FORMULA','first'),UNIT_OF_MEASURE=('UNIT_OF_MEASURE','first'),OWNER=('OWNER','first'));x=a.merge(b,on=['ENTITY','KPI_CODE'],how='outer',indicator=True);x['SOURCE']=src;x['DIFFERENCE']=x.REPORT_VALUE-x.SHAREPOINT_VALUE;x['DIFFERENCE_PCT']=np.where(x.SHAREPOINT_VALUE.abs()>at,x.DIFFERENCE/x.SHAREPOINT_VALUE,np.nan);ok=(x.DIFFERENCE.abs()<=at+rt*x.SHAREPOINT_VALUE.abs()).fillna(False);x['STATUS']=np.select([x._merge.eq('left_only'),x._merge.eq('right_only'),ok],['NÃO ESTÁ NA REFERÊNCIA','SOMENTE NA REFERÊNCIA','OK'],default='DIVERGENTE');x['KEY_USED']='ENTITY + KPI_CODE';return x.drop(columns='_merge')
def run(upload_bytes,reference_bundle,abs_tol=.01,rel_tol=.0001,upload_filename='arquivo.xlsx'):
 try:names=pd.ExcelFile(io.BytesIO(upload_bytes),engine='openpyxl').sheet_names
 except Exception as e:return {'error':f'Não foi possível abrir o arquivo enviado: {e}'}
 resolved={s:locate(names,s) for s in REPORTS};missing=[s for s in REPORTS if not resolved[s]]
 if missing:return {'error':'Abas obrigatórias não encontradas: '+', '.join(missing),'missing_sheets':missing,'resolved':resolved,'maps':{}}
 maps={};heads={};results=[]
 try:
  for s in REPORTS:
   d,heads[s]=read(upload_bytes,resolved[s]);maps[s]=mapping(d,s);results.append(compare(standard(d,maps[s],s),reference_bundle['reference_data'],abs_tol,rel_tol,s))
 except Exception as e:return {'error':str(e),'resolved':resolved,'maps':maps,'heads':heads}
 combined=pd.concat(results,ignore_index=True);summary=combined.groupby(['SOURCE','STATUS']).size().reset_index(name='QUANTIDADE');return {'error':None,'resolved':resolved,'maps':maps,'heads':heads,'results':results,'combined':combined,'summary':summary,'definitions':reference_bundle['definitions'],'parameters':{'KEY_USED':'ENTITY + KPI_CODE','REFERENCE_FILE':reference_bundle['filename'],'UPLOAD_FILE':upload_filename,'ABS_TOL':abs_tol,'REL_TOL':rel_tol,'GENERATED_AT':datetime.now().isoformat(timespec='seconds')}}
def excel(sheets):
 out=io.BytesIO()
 with pd.ExcelWriter(out,engine='xlsxwriter') as w:
  for name,d in sheets.items():d.to_excel(w,index=False,sheet_name=name[:31])
  for ws in w.sheets.values():ws.freeze_panes(1,0);ws.autofilter(0,0,ws.dim_rowmax,ws.dim_colmax);ws.set_column(0,max(0,ws.dim_colmax),20)
 return out.getvalue()
def dataframe_to_excel(d,sheet_name='Dados'):return excel({sheet_name:d})
def report(p):
 x=p['combined'];return excel({'Comparacao_Completa':x,'OK':x[x.STATUS=='OK'],'Achados':x[x.STATUS!='OK'],'Resumo':p['summary'],'Definition_Book':p['definitions'],'Parametros':pd.DataFrame([p['parameters']])})
def owner_data(b,name):
 key='' if name=='Sem responsável' else norm(name);return b['reference_data'][b['reference_data'].OWNER_KEY==key].copy()
def generate_responsible_template(b,responsible):
 d=owner_data(b,responsible);cols=['SOURCE','ENTITY','KPI_CODE','KPI_NAME','CLASSIFICATION','FORMULA','UNIT_OF_MEASURE','OWNER','VALUE'];d=d[cols].drop_duplicates(['SOURCE','ENTITY','KPI_CODE']).sort_values(['SOURCE','ENTITY','KPI_CODE']);d['VALUE']=np.nan;control=pd.DataFrame({'CAMPO':['ARQUIVO_REFERENCIA','DATA_HORA_GERACAO','RESPONSAVEL','REGISTROS','KPIS','ENTIDADES','VERSAO'],'VALOR':[b['filename'],datetime.now().isoformat(timespec='seconds'),responsible,len(d),d.KPI_CODE.nunique(),d.ENTITY.nunique(),'2.1']});instructions=pd.DataFrame({'INSTRUÇÕES':[f'Responsável: {responsible}','Preencha somente VALUE.','Não altere ENTITY nem KPI_CODE.']});sheets={s:d[d.SOURCE==s] for s in REPORTS};sheets.update({'INSTRUÇÕES':instructions,'CONTROLE':control});return excel(sheets),f'template_{safe(responsible)}.xlsx'
def generate_all_templates(b):
 out=io.BytesIO();rows=[]
 with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as z:
  for n in get_responsibles(b).display_name:
   content,fn=generate_responsible_template(b,n);z.writestr(fn,content);d=owner_data(b,n);rows.append({'RESPONSAVEL':n,'ARQUIVO':fn,'KPIS':d.KPI_CODE.nunique(),'ENTIDADES':d.ENTITY.nunique()})
  z.writestr('resumo_templates.xlsx',excel({'Resumo':pd.DataFrame(rows)}))
 return out.getvalue()
def generate_classification_catalog(b):return excel({'Classificacoes':get_classifications(b)['data'],'KPIs':get_kpi_catalog(b),'Responsaveis':get_responsibles(b),'Entidades':pd.DataFrame({'ENTITY':sorted(b['reference_data'].ENTITY.unique())}),'Resumo':pd.DataFrame([get_reference_summary(b)]),'Definition_Book':b['definitions']})
def get_reference_file_download(data):return data
