import re, io
from datetime import datetime, timedelta
import pandas as pd
import streamlit as st

st.set_page_config(page_title="RKAV Volendam | Kleedkamerplanner", page_icon="⚽", layout="wide", initial_sidebar_state="expanded")
st.markdown("""<style>
:root{--rkav:#f47721;--rkav-dark:#c94d0c;--ink:#202938}
.stApp{background:linear-gradient(180deg,#fff0e5 0,#fff8f3 260px,#f8f9fb 560px);color:var(--ink)}
.block-container{max-width:1320px;padding-top:2rem;padding-bottom:4rem}
[data-testid="stSidebar"]{background:linear-gradient(180deg,#ad3d09,#e56816 65%,#f47721)}
[data-testid="stSidebar"] *{color:#fff!important}
[data-testid="stSidebar"] input{color:#202938!important}
[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p{font-weight:650}
h1,h2,h3{letter-spacing:-.035em;color:#202938}
[data-testid="stMetric"]{background:#fff;border:1px solid #ffdcc6;border-top:4px solid #f47721;border-radius:14px;padding:20px;box-shadow:0 5px 20px #8d3b1610}
[data-testid="stFileUploader"]{background:white;border:2px dashed #f5a16c;border-radius:14px;padding:15px}
.stButton button[kind="primary"],.stDownloadButton button{background:#ec6718!important;color:white!important;border:0!important;border-radius:9px!important;font-weight:750!important;padding:.7rem 1.2rem!important}
.stButton button:hover,.stDownloadButton button:hover{border-color:#b94409!important;box-shadow:0 2px 10px #ed6c2d44!important}
[data-testid="stDataFrame"], [data-testid="stDataEditor"]{border:1px solid #f9d4ba;border-radius:12px;overflow:hidden}
.rkav-hero{background:linear-gradient(115deg,#ad3c09 0%,#ed6b1a 58%,#ff9944 100%);border-radius:20px;padding:36px;margin-bottom:24px;border-left:8px solid #ffcf9c;box-shadow:0 12px 28px #bb4c1528;color:white}
.rkav-kicker{color:#fff0dd;font-size:.82rem;font-weight:800;letter-spacing:.14em;text-transform:uppercase}
.rkav-hero h1{color:white!important;margin:.5rem 0!important;font-size:2.5rem!important}
.rkav-hero p{color:#fff3e8;margin:0;font-size:1.03rem}
.rkav-section{font-size:1.15rem;font-weight:800;color:#bd4a0b;margin:1.4rem 0 .65rem;padding-left:12px;border-left:4px solid #f47721}
.rkav-foot{font-size:.8rem;color:#6a7888;border-top:1px solid #f5d5be;padding-top:14px;margin-top:32px}
</style>""",unsafe_allow_html=True)
st.markdown("""<div class="rkav-hero"><div class="rkav-kicker">RKAV VOLENDAM · WEDSTRIJDDAG</div><h1>⚽ Kleedkamerplanner</h1><p>Van wedstrijdschema naar een overzichtelijke kleedkamerindeling.</p></div>""",unsafe_allow_html=True)
st.markdown('<div class="rkav-section">01 / Wedstrijdschema importeren</div>',unsafe_allow_html=True)
upload=st.file_uploader('Upload wedstrijdschema (.xlsx)',type='xlsx')
st.sidebar.markdown('## ⚙️ Planninginstellingen')
st.sidebar.caption('Stel de tijden en verdelingsregels in.')
durations={}
for k,default in [('O8-O10',50),('O11-O12',60),('O13-O15',70),('O16-O19',80),('O20+',90)]:
 durations[k]=st.sidebar.number_input('Wedstrijdduur '+k+' (min)',30,150,default,5)
before=st.sidebar.number_input('Omkleedtijd vóór wedstrijd (min)',0,180,30,5)
after=st.sidebar.number_input('Omkleedtijd na wedstrijd (min)',0,180,30,5)
allow_sharing=st.sidebar.checkbox('Gebruik gescheiden kooien bij tekort',True)
rooms={i:([f'{i}A',f'{i}B',f'{i}C',f'{i}D'] if i in (1,2,4,9,10,13,14,17,18,19,21,22) else [f'{i}k1',f'{i}k2']) for i in range(1,24)}
# Exacte onderverdeling afgeleid uit Blad2 van het voorbeeldbestand.
rooms={1:['1A','1B','1C','1D'],2:['2A','2B','2C','2D'],3:['3k1','3k2'],4:['4A','4B','4C','4D'],5:['5k1','5k2'],6:['6A','6B'],7:['7k1','7k2'],8:['8k1','8k2'],9:['9A','9B','9C','9D'],10:['10A','10B','10C','10D'],11:['11k1','11k2'],12:['12k1','12k2'],13:['13A','13B','13C','13D'],14:['14A','14B','14C','14D'],15:['15k1','15k2'],16:['16k1','16k2'],17:['17A','17B','17C','17D'],18:['18A','18B','18C','18D'],19:['19A','19B','19C','19D'],20:['20k1','20k2'],21:['21A','21B','21C','21D'],22:['22A','22B','22C','22D'],23:['23']}

def category(team):
 m=re.search(r'\b(?:JO|MO|O)(\d{1,2})\b',str(team),re.I)
 if not m:return 'O20+'
 n=int(m.group(1))
 return 'O8-O10' if n<=10 else 'O11-O12' if n<=12 else 'O13-O15' if n<=15 else 'O16-O19' if n<=19 else 'O20+'
def gender(team):
 t=str(team).upper()
 return 'M' if re.search(r'\bMO\d+',t) else 'J' if re.search(r'\bJO\d+',t) else 'O'
def parsed_time(v):
 if isinstance(v,(float,int)):return datetime(2026,1,1)+timedelta(days=float(v)%1)
 if hasattr(v,'hour'):return datetime(2026,1,1,v.hour,v.minute)
 return datetime.strptime(str(v).strip()[:5],'%H:%M').replace(year=2026,month=1,day=1)

def run(df):
 occupied={r:[] for slots in rooms.values() for r in slots}; result=[]
 for idx,row in df.iterrows():
  if pd.isna(row.iloc[2]) or pd.isna(row.iloc[1]):continue
  start=parsed_time(row.iloc[1]); home=str(row.iloc[2]); away=str(row.iloc[4]); cat=category(home)
  dur=int(durations[cat]); extra=120 if re.search(r'\bRKAV\s+Volendam\b',home,re.I) else 60 if re.search(r'Volendam\s+(?:2|O23-1)\b',home,re.I) else before
  begin=start-timedelta(minutes=extra); end=start+timedelta(minutes=dur+after)
  for role,team in [('Thuis',home),('Uit',away)]:
   g=gender(team); chosen=None
   # Eerst een volledig vrije fysieke kamer, daarna pas een gescheiden kooi.
   for shared in (False,True):
    if shared and not allow_sharing:continue
    for room,slots in rooms.items():
     if (len(slots)>1)!=shared:continue
     if any(any(begin<e and end>b for b,e,_,_ in occupied[s]) for s in slots):continue
     chosen=slots[0];break
    if chosen:break
   if chosen is None and allow_sharing:
    for room,slots in rooms.items():
     for slot in slots:
      if any(begin<e and end>b for b,e,_,_ in occupied[slot]):continue
      # Voor de zekerheid mogen jongens en meiden niet gelijktijdig in dezelfde fysieke kamer.
      other=[x for s in slots for x in occupied[s] if begin<x[1] and end>x[0]]
      if any(g!='O' and og!='O' and g!=og for _,_,og,_ in other):continue
      chosen=slot;break
     if chosen:break
   if chosen:occupied[chosen].append((begin,end,g,team))
   result.append({'Excelrij':idx+2,'Tijd':start.strftime('%H:%M'),'Veld':row.iloc[6], 'Rol':role,'Team':team,'Kleedkamer':chosen or 'CONFLICT','Vanaf':begin.strftime('%H:%M'),'Tot':end.strftime('%H:%M'),'Status':'OK' if chosen else 'NIET INDEELBAAR'})
 return pd.DataFrame(result)

# Scheidsrechters worden apart gepland: deze kamers zijn niet beschikbaar voor teams.
st.sidebar.markdown('### 🧑‍⚖️ Scheidsrechters')
ref_before=st.sidebar.number_input('Scheidsrechter aanwezig vóór aftrap (min)',0,180,30,5)
ref_after=st.sidebar.number_input('Scheidsrechter na afloop (min)',0,180,30,5)
ref_rooms=['SR1','SR2','SR3']

def plan_referees(source):
    # Eén arbitragegroep per wedstrijd; meerdere officials van één wedstrijd delen een ruimte.
    occupied={room:[] for room in ref_rooms}
    assignments=[]
    for idx,row in source.iterrows():
        if pd.isna(row.iloc[1]) or pd.isna(row.iloc[2]):
            continue
        start=parsed_time(row.iloc[1]); home=str(row.iloc[2]); away=str(row.iloc[4]); cat=category(home)
        begin=start-timedelta(minutes=int(ref_before))
        end=start+timedelta(minutes=int(durations[cat])+int(ref_after))
        chosen=next((room for room in ref_rooms if all(not(begin<e and end>b) for b,e in occupied[room])),None)
        if chosen: occupied[chosen].append((begin,end))
        assignments.append({'Excelrij':idx+2,'Tijd':start.strftime('%H:%M'),'Wedstrijd':f'{home} - {away}',
            'Arbitrage':'Nog in te vullen','Scheidsrechterkamer':chosen or 'CONFLICT',
            'Vanaf':begin.strftime('%H:%M'),'Tot':end.strftime('%H:%M'),
            'Status':'OK' if chosen else 'NIET INDEELBAAR'})
    return pd.DataFrame(assignments)

if upload:
 try:
  df=pd.read_excel(upload,sheet_name=0)
  if df.shape[1]<7:st.error('Onvoldoende kolommen. Verwacht: Datum, Tijd, Thuisteam, ..., Uitteam, ..., Veld.');st.stop()
  result=run(df)
  referees=plan_referees(df)
  st.markdown('<div class="rkav-section">02 / Resultaat in één oogopslag</div>',unsafe_allow_html=True)
  c1,c2,c3=st.columns(3)
  c1.metric('Teams',len(result))
  c2.metric('Ingepland',int((result.Status=='OK').sum()))
  c3.metric('Niet ingedeeld',int((result.Status!='OK').sum()))
  st.markdown('<div class="rkav-section">03 / Kleedkamerindeling controleren</div>',unsafe_allow_html=True)
  st.caption('Kleedkamers zijn voorlopig automatisch toegewezen. Wijzigingen in de tabel zijn voor controle; een herberekening na handmatig wijzigen is nog niet beschikbaar.')
  edited=st.data_editor(result,hide_index=True,num_rows='fixed',use_container_width=True)
  buf=io.BytesIO()
  with pd.ExcelWriter(buf,engine='xlsxwriter') as writer:
   edited.to_excel(writer,index=False,sheet_name='Kleedkamerindeling')
   ref_edited.to_excel(writer,index=False,sheet_name='Scheidsrechters')
  st.markdown('<div class="rkav-section">04 / Scheidsrechterkleedkamers</div>',unsafe_allow_html=True)
  st.caption('Per wedstrijd wordt voorlopig één arbitragegroep ingepland. Vul namen en eventuele bijzonderheden zelf in. Bij overlap krijgt een groep een andere kamer of een conflictmelding.')
  rc1,rc2=st.columns(2)
  rc1.metric('Arbitragegroepen',len(referees))
  rc2.metric('Conflicten arbitrage',int((referees.Status!='OK').sum()))
  ref_edited=st.data_editor(referees,hide_index=True,num_rows='fixed',use_container_width=True,key='referee_editor')
  st.markdown('<div class="rkav-section">05 / Indeling exporteren</div>',unsafe_allow_html=True)
  st.download_button('⬇️ Download indeling als Excel',buf.getvalue(),'rkav_kleedkamerindeling.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
  st.warning('Let op: arbitrage wordt voorlopig per wedstrijd als één groep gepland. Geslacht, individuele scheidsrechters, eventuele gedeelde kamers, volledige uitzonderingen en handmatige wijzigingen worden nog niet automatisch gevalideerd. Controle door planner vereist.')
 except Exception as e:st.error(f'Inlezen/plannen mislukt: {e}')

st.markdown('<div class="rkav-foot">RKAV Volendam · Kleedkamerplanner · Prototype — controleer de indeling vóór publicatie.</div>',unsafe_allow_html=True)
