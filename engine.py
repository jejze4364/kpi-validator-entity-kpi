from datetime import datetime
import pandas as pd, streamlit as st
from engine import run, report

st.set_page_config(page_title='KPI Validator', page_icon='📊', layout='wide')

st.title('KPI Validator | ENTITY + KPI_CODE')
st.caption('Leitura otimizada de SHAREPOINT, LOGS, BOPS, SL e Definition Book.')

with st.sidebar:
    at = st.number_input('Tolerância absoluta', 0.0, value=.01)
    rt = st.number_input('Tolerância relativa', 0.0, value=.0001)

u = st.file_uploader('Consolidador', type=['xlsx', 'xlsm'])

if not u:
    st.stop()

with st.spinner('Processando...'):
    p = run(u.getvalue(), at, rt)

if p.get('error'):
    st.error(p['error'])
    st.write(p.get('resolved'))
    st.write(p.get('maps'))
    st.stop()

x = pd.concat(p['results'], ignore_index=True)

ok = int((x.STATUS == 'OK').sum())
div = int((x.STATUS == 'DIVERGENTE').sum())

nao_comparados = int(
    x.STATUS.isin(
        [
            'NÃO ESTÁ NO SHAREPOINT',
            'SOMENTE NO SHAREPOINT'
        ]
    ).sum()
)

base_conformidade = ok + div

conf = (
    ok / base_conformidade * 100
    if base_conformidade > 0
    else 0
)

a, b, c, d = st.columns(4)

a.metric('OK', f'{ok:,}')
b.metric('Divergentes', f'{div:,}')
c.metric('Não Comparados', f'{nao_comparados:,}')
d.metric('Conformidade', f'{conf:.2f}%')

t1, t2, t3 = st.tabs(
    [
        'Resumo',
        'Comparação',
        'Exportar'
    ]
)

with t1:

    st.dataframe(
        x.groupby(
            ['SOURCE', 'STATUS']
        )
        .size()
        .reset_index(
            name='QUANTIDADE'
        ),
        use_container_width=True,
        hide_index=True
    )

    st.dataframe(
        x.groupby(
            'STATUS'
        )
        .size()
        .reset_index(
            name='TOTAL'
        ),
        use_container_width=True,
        hide_index=True
    )

with t2:

    s = st.multiselect(
        'Status',
        sorted(
       
