# Roteirização NIP - baseada em 2_Lista_Continua(6).py.
# K = COLABORADOR PLANEJADO
from __future__ import annotations
from pathlib import Path
from io import BytesIO
from collections import defaultdict
from datetime import datetime
import math
import re
import html
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import streamlit as st
import folium
from folium.plugins import MarkerCluster
from streamlit_folium import st_folium

ROOT = Path(__file__).resolve().parents[1]
COLUMNS = ['NOTA','STATUS','TIPO DEMANDA','MUNICIPIO','REGIONAL','ENDEREÇO','BAIRRO','LATITUDE','LONGITUDE','COLABORADOR','MATRICULA']
COLS_ORIG = {'NOTA':'NOTA','STATUS':'STATUS','TIPO DEMANDA':'TIPO DEMANDA','MUNICIPIO':'MUNICÍPIO','REGIONAL':'REGIONAL SERVIÇO','ENDEREÇO':'ENDEREÇO','BAIRRO':'BAIRRO','LATITUDE':'LATITUDE','LONGITUDE':'LONGITUDE','COLABORADOR':'COLABORADOR PLANEJADO','MATRICULA':'MATRÍCULA'}

def esc(x):
    if pd.isna(x): return ''
    return html.escape(str(x),quote=True)

def safe_filename(s):
    return re.sub(r'[^a-zA-Z0-9 _.-]+','_',str(s)).strip(' ._')[:90] or 'SEM_COLABORADOR'

@st.cache_data(show_spinner=False)
def load_file(raw:bytes):
    df=pd.read_excel(BytesIO(raw),dtype=str,engine='openpyxl')
    df.columns=[str(c).strip() for c in df.columns]
    missing=[v for v in COLS_ORIG.values() if v not in df.columns]
    if missing: raise ValueError('Colunas ausentes: '+', '.join(missing))
    # Verifica explicitamente a posicao solicitada, sem adivinhar o responsavel.
    if df.columns[10]!='COLABORADOR PLANEJADO':
        raise ValueError('A coluna K deve ser COLABORADOR PLANEJADO; confira o arquivo.')
    df=df.rename(columns={v:k for k,v in COLS_ORIG.items()})[COLUMNS].copy()
    for c in COLUMNS: df[c]=df[c].fillna('').astype(str).str.strip()
    for c in ['LATITUDE','LONGITUDE']:
        df[c+'_NUM']=pd.to_numeric(df[c].str.replace(',','.',regex=False),errors='coerce')
    df['VALIDA']=df['LATITUDE_NUM'].between(-90,90)&df['LONGITUDE_NUM'].between(-180,180)&(df['LATITUDE_NUM']!=0)&(df['LONGITUDE_NUM']!=0)
    return df

def haversine(a,b,c,d):
    r=6371.0088
    a,b,c,d=map(math.radians,[a,b,c,d]);dl=c-a;dn=d-b
    h=math.sin(dl/2)**2+math.cos(a)*math.cos(c)*math.sin(dn/2)**2
    return 2*r*math.asin(min(1,math.sqrt(h)))

def order_route(frame,reverse=False):
    rows=frame.to_dict('records')
    if not rows:return []
    la=sum(float(x['LATITUDE_NUM']) for x in rows)/len(rows)
    lo=sum(float(x['LONGITUDE_NUM']) for x in rows)/len(rows)
    if reverse:
        start=max(range(len(rows)),key=lambda i:haversine(la,lo,float(rows[i]['LATITUDE_NUM']),float(rows[i]['LONGITUDE_NUM'])))
        rows[0],rows[start]=rows[start],rows[0]
        result=[rows.pop(0)];la=float(result[-1]['LATITUDE_NUM']);lo=float(result[-1]['LONGITUDE_NUM'])
    else:result=[]
    while rows:
        i=min(range(len(rows)),key=lambda i:haversine(la,lo,float(rows[i]['LATITUDE_NUM']),float(rows[i]['LONGITUDE_NUM'])))
        x=rows.pop(i);x['DISTANCIA_PONTO_ANTERIOR_KM']=round(haversine(la,lo,float(x['LATITUDE_NUM']),float(x['LONGITUDE_NUM'])),3)
        result.append(x);la=float(x['LATITUDE_NUM']);lo=float(x['LONGITUDE_NUM'])
    for i,x in enumerate(result):
        x['ORDEM']=i+1
        if 'DISTANCIA_PONTO_ANTERIOR_KM' not in x:x['DISTANCIA_PONTO_ANTERIOR_KM']=0
    return result

def popup(row):
    # Mesmo conteudo de card no mapa e nos arquivos KML/KMZ.
    cells=''.join(f'<tr><td style="padding:4px 7px;border-bottom:1px solid #e4eaf0"><b>{esc(c)}</b></td><td style="padding:4px 7px;border-bottom:1px solid #e4eaf0">{esc(row.get(c,""))}</td></tr>' for c in COLUMNS)
    return '<div style="font-family:Arial,sans-serif;min-width:280px;color:#14344b"><b style="color:#087888">NIP | Roteirização</b><br><b>Ordem:</b> '+esc(row.get('ORDEM',''))+'<table style="border-collapse:collapse;width:100%;font-size:12px;margin-top:6px">'+cells+'</table></div>'

def kml_content(records,title):
    kml=ET.Element('kml',xmlns='http://www.opengis.net/kml/2.2');doc=ET.SubElement(kml,'Document');ET.SubElement(doc,'name').text=title
    by=defaultdict(list)
    for r in records:by[r['COLABORADOR']].append(r)
    for name,rows in by.items():
        folder=ET.SubElement(doc,'Folder');ET.SubElement(folder,'name').text=name or 'SEM COLABORADOR'
        for r in rows:
            point=ET.SubElement(folder,'Placemark');ET.SubElement(point,'name').text=f"{r['ORDEM']:03d} | {r['NOTA']}"
            ET.SubElement(point,'description').text=popup(r)
            ET.SubElement(ET.SubElement(point,'Point'),'coordinates').text=f"{r['LONGITUDE_NUM']},{r['LATITUDE_NUM']},0"
        if len(rows)>1:
            line=ET.SubElement(folder,'Placemark');ET.SubElement(line,'name').text='Sequência da rota (linha reta)'
            ET.SubElement(ET.SubElement(line,'LineString'),'coordinates').text=' '.join(f"{r['LONGITUDE_NUM']},{r['LATITUDE_NUM']},0" for r in rows)
    return ET.tostring(kml,encoding='utf-8',xml_declaration=True)

def gpx_content(records,title):
    root=ET.Element('gpx',version='1.1',creator='NIP Roteirizacao',xmlns='http://www.topografix.com/GPX/1/1')
    ET.SubElement(root,'name').text=title
    by=defaultdict(list)
    for r in records:
        w=ET.SubElement(root,'wpt',lat=str(r['LATITUDE_NUM']),lon=str(r['LONGITUDE_NUM']))
        ET.SubElement(w,'name').text=str(r['NOTA']);ET.SubElement(w,'desc').text=f"{r['COLABORADOR']} | {r['STATUS']} | {r['MUNICIPIO']}"
        by[r['COLABORADOR']].append(r)
    for name,rows in by.items():
        trk=ET.SubElement(root,'trk');ET.SubElement(trk,'name').text=name
        seg=ET.SubElement(trk,'trkseg')
        for r in rows:ET.SubElement(seg,'trkpt',lat=str(r['LATITUDE_NUM']),lon=str(r['LONGITUDE_NUM']))
    return ET.tostring(root,encoding='utf-8',xml_declaration=True)

def txt_content(rows):
    return ('\n\n'.join('ORDEM: '+str(r['ORDEM'])+'\n'+'\n'.join(f'{c}: {r.get(c,"")}' for c in COLUMNS) for r in rows)).encode('utf-8-sig')

def excel_content(rows):
    out=BytesIO();pd.DataFrame(rows).reindex(columns=['ORDEM']+COLUMNS).to_excel(out,index=False,engine='openpyxl');return out.getvalue()

def export_zip(rows):
    b=BytesIO();today=datetime.now().strftime('%Y%m%d')
    with zipfile.ZipFile(b,'w',zipfile.ZIP_DEFLATED) as z:
        z.writestr('ROTA_TOTAL.kml',kml_content(rows,'NIP Rota Total'))
        z.writestr('ROTA_TOTAL.gpx',gpx_content(rows,'NIP Rota Total'))
        z.writestr('ROTA_TOTAL.xlsx',excel_content(rows))
        z.writestr('ROTA_TOTAL.txt',txt_content(rows))
        kmz=BytesIO()
        with zipfile.ZipFile(kmz,'w',zipfile.ZIP_DEFLATED) as km:zkm=km.writestr('doc.kml',kml_content(rows,'NIP Rota Total'))
        z.writestr('ROTA_TOTAL.kmz',kmz.getvalue())
        by=defaultdict(list)
        for r in rows:by[r['COLABORADOR']].append(r)
        for name,group in by.items():
            stem='POR_COLABORADOR/'+safe_filename(name)
            z.writestr(stem+'.kml',kml_content(group,name));z.writestr(stem+'.gpx',gpx_content(group,name))
            z.writestr(stem+'.xlsx',excel_content(group));z.writestr(stem+'.txt',txt_content(group))
            kmz=BytesIO()
            with zipfile.ZipFile(kmz,'w',zipfile.ZIP_DEFLATED) as km:km.writestr('doc.kml',kml_content(group,name))
            z.writestr(stem+'.kmz',kmz.getvalue())
    return b.getvalue()

st.markdown("""<style>
[data-testid="stAppViewContainer"]{background:#edf4f8}
.block-container{max-width:100%;padding-top:0.8rem}
.nip-route-header{background:#123551;color:white;padding:18px 24px;border-radius:5px;margin-bottom:16px}
.nip-route-header h2{color:white;margin:0;font-size:23px}
.nip-route-header small{color:#d5e6f0}
</style><div class="nip-route-header"><h2>NIP | Roteirização Massiva</h2><small>Roteirização · Lista contínua por colaborador planejado</small></div>""",unsafe_allow_html=True)
st.page_link('pages/01_visao_geral.py',label='← Visão Geral',icon='🗺️')
st.markdown('### Roteirização')
file=next((p for p in [ROOT/'BASE_SANEAMENTO.xlsx',ROOT/'BASE_SANEAMENTO(1).xlsx',ROOT/'BASE_SANEAMENTO_ATUALIZADA.xlsx'] if p.is_file()),None)
with st.expander('Base de dados',expanded=file is None):
    if file:st.success(f'Planilha do repositório: {file.name}')
    up=st.file_uploader('Opcional: selecionar outra BASE_SANEAMENTO.xlsx',type=['xlsx'],key='nip_route_upload')
if up is None and file is None:
    st.warning('Coloque BASE_SANEAMENTO.xlsx na raiz do GitHub ou carregue a planilha acima.');st.stop()
try:
    raw=up.getvalue() if up else file.read_bytes()
    df=load_file(raw)
except Exception as e:st.error(f'Falha na leitura da planilha: {e}');st.stop()

f1,f2,f3,f4=st.columns(4)
with f1: colaboradores=st.multiselect('Colaborador planejado',sorted(x for x in df['COLABORADOR'].unique() if x),key='route_staff')
with f2: regionais=st.multiselect('Regional',sorted(x for x in df['REGIONAL'].unique() if x),key='route_region')
with f3: municipios=st.multiselect('Município',sorted(x for x in df['MUNICIPIO'].unique() if x),key='route_city')
with f4: status=st.multiselect('Status',sorted(x for x in df['STATUS'].unique() if x),key='route_status')
query=st.text_input('Pesquisar nota, endereço, bairro ou matrícula',placeholder='Digite para pesquisar...')
for key,sel in [('COLABORADOR',colaboradores),('REGIONAL',regionais),('MUNICIPIO',municipios),('STATUS',status)]:
    if sel:df=df[df[key].isin(sel)]
if query:
    q=re.escape(query)
    df=df[df[['NOTA','ENDEREÇO','BAIRRO','MATRICULA']].apply(lambda c:c.str.contains(q,case=False,na=False,regex=True)).any(axis=1)]
valid=df[df['VALIDA']].copy();invalid=df[~df['VALIDA']].copy()
c1,c2,c3,c4=st.columns(4)
c1.metric('Obras selecionadas',f'{len(df):,}'.replace(',','.'))
c2.metric('Colaboradores',valid['COLABORADOR'].replace('',pd.NA).nunique())
c3.metric('Com coordenadas',f'{len(valid):,}'.replace(',','.'))
c4.metric('Retidas para correção',f'{len(invalid):,}'.replace(',','.'))
if not invalid.empty:
    with st.expander('Obras sem coordenadas válidas'):
        st.dataframe(invalid[COLUMNS],use_container_width=True,hide_index=True)
if valid.empty:st.info('Nenhuma obra com coordenadas válidas atende aos filtros.');st.stop()
col1,col2,col3=st.columns([1,1,1])
with col1:reverse=st.checkbox('Varredura reversa',value=False)
with col2:show_routes=st.checkbox('Exibir sequência das rotas',value=True)
with col3:st.caption('Distâncias em linha reta; não são quilômetros rodoviários.')
records=[]
for name,group in valid.groupby('COLABORADOR',sort=True,dropna=False):
    records.extend(order_route(group,reverse=reverse))
# Evita executar milhares de marcadores na primeira carga sem necessidade.
st.markdown('#### Mapa operacional')
if len(records)>2000:
    st.info(f'{len(records)} obras selecionadas. O mapa exibe até 2.000 marcadores para preservar desempenho; a exportação contém todas.')
map_rows=records[:2000]
mp=folium.Map(location=[valid['LATITUDE_NUM'].mean(),valid['LONGITUDE_NUM'].mean()],zoom_start=7,tiles='OpenStreetMap',control_scale=True)
cluster=MarkerCluster(name='Obras').add_to(mp)
colors=['blue','green','purple','orange','red','cadetblue','darkgreen','darkpurple']
for i,(name,group) in enumerate(pd.DataFrame(map_rows).groupby('COLABORADOR',sort=True)):
    col=colors[i%len(colors)]
    pts=[]
    for r in group.to_dict('records'):
        pts.append([float(r['LATITUDE_NUM']),float(r['LONGITUDE_NUM'])])
        folium.Marker(pts[-1],popup=folium.Popup(popup(r),max_width=400),tooltip=f"{r['ORDEM']} | {r['NOTA']} | {name}",icon=folium.Icon(color=col,icon='info-sign')).add_to(cluster)
    if show_routes and len(pts)>1:folium.PolyLine(pts,weight=3,opacity=.7,color=col,tooltip=f'Rota: {name} (linha reta)').add_to(mp)
folium.LayerControl().add_to(mp)
st_folium(mp,height=560,use_container_width=True,returned_objects=[])
st.markdown('#### Obras roteirizadas')
st.dataframe(pd.DataFrame(records).reindex(columns=['ORDEM']+COLUMNS),hide_index=True,use_container_width=True,height=380)
if st.button('Gerar arquivos de roteirização',type='primary'):
    with st.spinner('Gerando GPX, KML, KMZ, Excel e TXT por colaborador...'):
        st.session_state['nip_route_zip']=export_zip(records)
        st.session_state['nip_route_signature']=(len(records),str(colaboradores),str(regionais),str(municipios),str(status),query,reverse)
sig=(len(records),str(colaboradores),str(regionais),str(municipios),str(status),query,reverse)
if st.session_state.get('nip_route_signature')==sig and st.session_state.get('nip_route_zip'):
    st.download_button('📦 Baixar rotas — ZIP (KMZ, KML, GPX, Excel e TXT)',st.session_state['nip_route_zip'],file_name='NIP_ROTEIRIZACAO_'+datetime.now().strftime('%Y%m%d')+'.zip',mime='application/zip',type='primary')
