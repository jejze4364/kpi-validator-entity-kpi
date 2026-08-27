from datetime import datetime
from pathlib import Path
import streamlit as st
import engine
st.set_page_config(page_title='KPI Validator',page_icon='📊',layout='wide')
BASE=Path(__file__).resolve().parent
candidates=[BASE/'Consolidador 2.0.xlsm',BASE/'consolidador.xlsm']+sorted(BASE.glob('Consolidador*.xlsm'))+sorted(BASE.glob('Consolidador*.xlsx'))
REFERENCE_PATH=next((p for p in candidates if p.exists()),None)
@st.cache_data(show_spinner=False)
def load_ref(data,name,modified):return engine.load_reference(data,name,modified)
st.title('KPI Validator | ENTITY + KPI_CODE')
st.caption('Consulta da referência, templates e validação automática de LOGS, BOPS e SL.')
if REFERENCE_PATH is None:st.error('Arquivo de referência não encontrado. Coloque Consolidador 2.0.xlsm na mesma pasta de app.py.');st.stop()
try:
 reference_bytes=REFERENCE_PATH.read_bytes();modified=datetime.fromtimestamp(REFERENCE_PATH.stat().st_mtime);reference=load_ref(reference_bytes,REFERENCE_PATH.name,modified.isoformat())
except Exception as e:st.error(f'Não foi possível carregar o arquivo de referência: {e}');st.stop()
summary=engine.get_reference_summary(reference);catalog=engine.get_kpi_catalog(reference);owners=engine.get_responsibles(reference);classes=engine.get_classifications(reference)
tabs=st.tabs(['Início','Classificações','Templates','Downloads','Validação','Resultados','Diagnóstico','Exportar'])
with tabs[0]:
 st.success(f'Referência carregada: {REFERENCE_PATH.name}')
 a,b,c,d=st.columns(4);a.metric('KPIs',summary['total_kpis']);b.metric('Entidades',summary['total_entities']);c.metric('Responsáveis',summary['total_responsibles']);d.metric('Classificações',summary['total_classifications'])
 a,b,c,d,e=st.columns(5);a.metric('LOGS',summary['logs_kpis']);b.metric('BOPS',summary['bops_kpis']);c.metric('SL',summary['sl_kpis']);d.metric('Sem responsável',summary['without_owner']);e.metric('Sem classificação',summary['without_classification'])
 st.write({'arquivo':REFERENCE_PATH.name,'tamanho_mb':round(len(reference_bytes)/1048576,2),'modificação':modified.strftime('%d/%m/%Y %H:%M:%S'),'abas':reference['resolved']})
with tabs[1]:
 st.subheader('Classificações e KPIs');q=st.text_input('Buscar');view=catalog
 if q:view=view[view.astype(str).apply(lambda s:s.str.contains(q,case=False,na=False,regex=False)).any(axis=1)]
 st.dataframe(view,use_container_width=True,hide_index=True,height=620)
 if not classes['available']:st.warning('A informação de classificação não foi encontrada no consolidador.')
with tabs[2]:
 st.subheader('Templates');opts=owners.display_name.tolist();owner=st.selectbox('Responsável',opts,index=None)
 if owner:
  content,name=engine.generate_responsible_template(reference,owner);st.download_button('Baixar template do responsável',content,name,type='primary')
 if opts:st.download_button('Baixar todos os templates',engine.generate_all_templates(reference),'templates_kpi.zip')
with tabs[3]:
 st.subheader('Central de Downloads');a,b,c=st.columns(3);a.download_button('Baixar consolidador completo',reference_bytes,REFERENCE_PATH.name);b.download_button('Baixar catálogo',engine.generate_classification_catalog(reference),'catalogo_classificacoes_kpi.xlsx');c.download_button('Baixar relação de KPIs',engine.dataframe_to_excel(catalog,'KPIs'),'relacao_kpis.xlsx')
 a,b=st.columns(2);a.download_button('Baixar responsáveis',engine.dataframe_to_excel(owners,'Responsaveis'),'responsaveis.xlsx');
 if not reference['definitions'].empty:b.download_button('Baixar Definition Book',engine.dataframe_to_excel(reference['definitions'],'Definition_Book'),'definition_book.xlsx')
with st.sidebar:at=st.number_input('Tolerância absoluta',0.0,value=.01,format='%.8f');rt=st.number_input('Tolerância relativa',0.0,value=.0001,format='%.8f')
with tabs[4]:
 u=st.file_uploader('Arquivo preenchido',type=['xlsx','xlsm'])
 if not u:st.info('Envie o arquivo preenchido para iniciar a validação. As consultas e downloads já estão disponíveis.')
 else:
  with st.spinner('Processando...'):st.session_state.payload=engine.run(u.getvalue(),reference,at,rt,u.name)
  if st.session_state.payload.get('error'):st.error(st.session_state.payload['error'])
  else:st.success('Validação concluída.')
p=st.session_state.get('payload');valid=p if p and not p.get('error') else None
with tabs[5]:
 if not valid:st.info('Execute a validação para visualizar os resultados.')
 else:
  x=valid['combined'];ok=int(x.STATUS.eq('OK').sum());a,b,c,d=st.columns(4);a.metric('Comparados',len(x));b.metric('OK',ok);c.metric('Achados',len(x)-ok);d.metric('Conformidade',f'{ok/len(x)*100:.2f}%' if len(x) else '0%');st.dataframe(valid['summary'],use_container_width=True,hide_index=True);st.dataframe(x,use_container_width=True,hide_index=True,height=620)
with tabs[6]:
 st.json({'referência':{'abas':reference['resolved'],'cabeçalhos':reference['heads'],'mapeamentos':reference['maps']}})
 if valid:st.json({'arquivo_enviado':{'abas':valid['resolved'],'cabeçalhos':valid['heads'],'mapeamentos':valid['maps']}})
with tabs[7]:
 if valid:st.download_button('Baixar relatório Excel',engine.report(valid),f'KPI_Entity_KPI_{datetime.now():%Y%m%d_%H%M%S}.xlsx',type='primary')
 else:st.info('Execute uma validação para gerar o relatório.')
