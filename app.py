from datetime import datetime
import pandas as pd,streamlit as st
from engine import run,report
st.set_page_config(page_title='KPI Validator',page_icon='📊',layout='wide');st.title('KPI Validator | ENTITY + KPI_CODE');st.caption('Leitura otimizada de SHAREPOINT, LOGS, BOPS, SL e Definition Book.')
with st.sidebar:at=st.number_input('Tolerância absoluta',0.0,value=.01);rt=st.number_input('Tolerância relativa',0.0,value=.0001)
u=st.file_uploader('Consolidador',type=['xlsx','xlsm'])
if not u:st.stop()
with st.spinner('Processando...'):p=run(u.getvalue(),at,rt)
if p.get('error'):st.error(p['error']);st.write(p.get('resolved'));st.write(p.get('maps'));st.stop()
x=pd.concat(p['results'],ignore_index=True);ok=int(x.STATUS.eq('OK').sum());a,b,c,d=st.columns(4);a.metric('Comparados',len(x));b.metric('OK',ok);c.metric('Achados',len(x)-ok);d.metric('Conformidade',f'{ok/len(x)*100:.2f}%' if len(x) else '0%')
t1,t2,t3=st.tabs(['Resumo','Comparação','Exportar'])
with t1:st.dataframe(x.groupby(['SOURCE','STATUS']).size().reset_index(name='QUANTIDADE'),use_container_width=True,hide_index=True)
with t2:
 s=st.multiselect('Status',sorted(x.STATUS.unique()),default=sorted(x.STATUS.unique()));q=st.text_input('Buscar Entity ou KPI');v=x[x.STATUS.isin(s)]
 if q:v=v[v.ENTITY.str.contains(q,case=False,na=False)|v.KPI_CODE.str.contains(q,case=False,na=False)]
 st.dataframe(v,use_container_width=True,hide_index=True,height=620)
with t3:st.download_button('Baixar Excel',report(p),f'KPI_Entity_KPI_{datetime.now():%Y%m%d_%H%M%S}.xlsx',type='primary')