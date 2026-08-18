from datetime import datetime
import pandas as pd
import streamlit as st
from engine import excel_report, run

st.set_page_config(page_title='KPI Validator | Entity + KPI',page_icon='📊',layout='wide')
BLUE='#001E60';YELLOW='#FFC72C';BG='#F4F7FB'
st.markdown(f'''<style>.stApp{{background:{BG}}}[data-testid="stSidebar"]{{background:linear-gradient(180deg,{BLUE},#063D91)}}[data-testid="stSidebar"] *{{color:white}}.hero{{background:linear-gradient(110deg,{BLUE},#0750B8);padding:26px 30px;border-radius:24px;color:white;margin-bottom:18px}}.pill{{display:inline-block;background:{YELLOW};color:{BLUE};font-weight:800;padding:5px 10px;border-radius:999px}}.card{{background:white;padding:17px;border-radius:17px;border:1px solid #E4EAF3;min-height:105px}}.label{{color:#67758C;font-size:.75rem;font-weight:700}}.value{{color:{BLUE};font-size:1.6rem;font-weight:800;margin-top:7px}}footer{{visibility:hidden}}</style>''',unsafe_allow_html=True)
st.markdown('<div class="hero"><div class="pill">AMBEV | DATA QUALITY</div><h1>KPI Validator Otimizado</h1><p>Comparação exclusiva pela chave ENTITY + KPI_CODE.</p></div>',unsafe_allow_html=True)
with st.sidebar:
    st.markdown('## Configuração')
    atol=st.number_input('Tolerância absoluta',0.0,value=.01,format='%.6f')
    rtol=st.number_input('Tolerância relativa',0.0,value=.0001,format='%.6f')
    st.caption('Definition Book enriquece o relatório, mas não participa da chave.')
up=st.file_uploader('Envie o Consolidador (.xlsx ou .xlsm)',type=['xlsx','xlsm'])
if not up:st.info('A análise lê somente SHAREPOINT, LOGS, BOPS, SL e Definition Book.');st.stop()
with st.spinner('Lendo somente as cinco abas necessárias...'):
    p=run(up.getvalue(),atol,rtol)
if p.get('error'):
    st.error(p['error']);st.write('Abas:',p.get('resolved'));st.write('Mapeamento:',p.get('maps'));st.stop()
allr=pd.concat(p['results'],ignore_index=True)
total=len(allr);ok=int((allr.STATUS=='OK').sum());div=int((allr.STATUS=='DIVERGENTE').sum());miss=int(allr.STATUS.isin(['NÃO ESTÁ NO SHAREPOINT','SOMENTE NO SHAREPOINT']).sum());conf=ok/total*100 if total else 0
summary,detail,mapping,definition,export=st.tabs(['Resumo','Comparação','Mapeamento','Definition Book','Exportar'])
with summary:
    cols=st.columns(5);cards=[('Comparados',total),('OK',ok),('Divergentes',div),('Ausências',miss),('Conformidade',f'{conf:.2f}%')]
    for c,(l,v) in zip(cols,cards):c.markdown(f'<div class="card"><div class="label">{l}</div><div class="value">{v}</div></div>',unsafe_allow_html=True)
    st.markdown('### Resultado por aba');st.dataframe(allr.groupby(['SOURCE','STATUS']).size().reset_index(name='QUANTIDADE'),use_container_width=True,hide_index=True)
    st.markdown('### Top KPIs com achados');st.dataframe(allr[allr.STATUS!='OK'].groupby('KPI_CODE').size().sort_values(ascending=False).head(30).reset_index(name='ACHADOS'),use_container_width=True,hide_index=True)
with detail:
    a,b,c,d=st.columns(4);sources=a.multiselect('Fonte',sorted(allr.SOURCE.dropna().unique()),default=sorted(allr.SOURCE.dropna().unique()));statuses=b.multiselect('Status',sorted(allr.STATUS.unique()),default=sorted(allr.STATUS.unique()));kpi=c.text_input('KPI');entity=d.text_input('Entity')
    view=allr[allr.SOURCE.isin(sources)&allr.STATUS.isin(statuses)]
    if kpi:view=view[view.KPI_CODE.str.contains(kpi,case=False,na=False)]
    if entity:view=view[view.ENTITY.str.contains(entity,case=False,na=False)]
    cols=[x for x in ['SOURCE','ENTITY','KPI_CODE','KPI_NAME','REPORT_VALUE','SHAREPOINT_VALUE','DIFFERENCE','DIFFERENCE_PCT','STATUS','REPORT_ROWS','SHAREPOINT_ROWS','DUPLICATE_ALERT','DEF_KPI_NAME','DEF_FORMULA','DEF_UOM','DEF_OWNER','KEY_USED'] if x in view.columns]
    st.dataframe(view[cols].sort_values(['STATUS','SOURCE','KPI_CODE','ENTITY']),use_container_width=True,hide_index=True,height=620)
with mapping:
    mapdf=pd.DataFrame([{'TIPO':k,'ABA':v,'CABEÇALHO':p['headers'].get(k),'ENTITY':p['maps'].get(k,{}).get('entity'),'KPI_CODE':p['maps'].get(k,{}).get('kpi'),'VALOR':p['maps'].get(k,{}).get('value')} for k,v in p['resolved'].items()])
    st.success('A única chave usada na comparação é ENTITY + KPI_CODE.')
    st.dataframe(mapdf,use_container_width=True,hide_index=True)
with definition:st.dataframe(p['definitions'],use_container_width=True,hide_index=True,height=620)
with export:
    st.download_button('Baixar relatório completo',excel_report(p),f'KPI_Entity_KPI_{datetime.now():%Y%m%d_%H%M%S}.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary',use_container_width=True)
