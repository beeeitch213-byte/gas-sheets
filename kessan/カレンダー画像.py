# 決算カレンダー画像を作る(kessan/CLAUDE.md の「決算カレンダー画像」)。
# 使い方: python3 -I kessan/カレンダー画像.py YYYY MM 銘柄.tsv 出力.png 時点の文字列
#   銘柄.tsv は「決算情報取得銘柄」シートの A列(コード)と B列(銘柄名)をタブ区切りで1行1銘柄。
#   発表日・予定時刻は SBI 証券から、祝日は内閣府の「国民の祝日」CSV から、その場で取り直す。
#   HTML を作り、Playwright(node)でスクリーンショットを撮って PNG にする。
import os,re,subprocess,sys,tempfile,calendar,datetime as dt,html,urllib.request
Y,M=int(sys.argv[1]),int(sys.argv[2]); codes=[l.rstrip('\n').split('\t') for l in open(sys.argv[3],encoding='utf-8') if l.strip()]
UA={'User-Agent':'Mozilla/5.0','Referer':'https://www.sbisec.co.jp/'}
get=lambda u,enc='cp932':urllib.request.urlopen(urllib.request.Request(u,headers=UA),timeout=30).read().decode(enc,'replace')
# 祝日(内閣府 CSV)
hol={}; csv_ok=True
try:
    for line in get('https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv').splitlines()[1:]:
        d,n=line.split(',')[:2]; y,m,dd=map(int,d.split('/')); hol[dt.date(y,m,dd)]=n
except Exception as e: csv_ok=False
def closed(d): return d.weekday()>=5 or d in hol or (d.month==1 and d.day<=3) or (d.month==12 and d.day==31)
fr=[d for d in (dt.date(Y,M,x) for x in range(1,calendar.monthrange(Y,M)[1]+1)) if d.weekday()==4][1]
sq=fr
while closed(sq): sq-=dt.timedelta(days=1)
# SBI
page=get('https://www.sbisec.co.jp/ETGate/?_ControlID=WPLETmgR001Control&_PageID=WPLETmgR001Mdtl20&_DataStoreID=DSWPLETmgR001Control&_ActionID=DefaultAID&burl=iris_economicCalendar&cat1=market&cat2=economicCalender&dir=tl1-cal%7Ctl2-schedule%7Ctl3-stock%7Ctl4-calsel&file=index.html&getFlg=on')
P=re.search(r"ANNOUNCE_INFO_PARAM = '([^']+)'",page).group(1)
ev={}; undated=[]; tbd=[]
for c,n in codes:
    t=get('https://vc.iris.sbisec.co.jp/calendar/settlement/stock/announcement_info_stock.do'+P+'&productCode='+c+'&inputWord=&callback=cb')
    r=re.search(r'\{[^{}]*"productCode":"%s"[^{}]*\}'%re.escape(c),t,re.S)
    if not r: continue
    f=lambda k:(re.search(r'"%s":"([^"]*)"'%k,r.group(0)) or [None,''])[1].replace('<br>',' ')
    date,tm,typ=f('date'),f('time').replace(' (予定)',''),re.sub('<[^>]+>','',f('type'))
    m=re.match(r'(\d{4})/(\d{2})/(\d{2})',date)
    if m and int(m[1])==Y and int(m[2])==M:
        tm='未定' if tm.strip() in ('--','') else tm
        ev.setdefault(int(m[3]),[]).append((tm,n,c)); 
        if tm=='未定': tbd.append(n)
    elif re.match(r'%d/%02d '%(Y,M),date): undated.append(f'{n}({c}) {M}月{date.split(" ")[1]} {tm}予定({typ.replace(" ","")})')
for d in ev: ev[d].sort(key=lambda x:(x[0]=='未定',x[0]))
cells=''
for w in calendar.Calendar(0).monthdayscalendar(Y,M):
    cells+='<tr>'
    for i,d in enumerate(w):
        if d==0: cells+='<td class="empty'+(' we' if i>=5 else '')+'"></td>';continue
        day=dt.date(Y,M,d); hn=hol.get(day)
        cls='sat we' if i==5 else 'sun we' if i==6 else ('sun' if closed(day) else '')
        lab=(f'<span class="hol">{hn}</span>' if hn and i<5 else '')+(f'<span class="hol">休場</span>' if closed(day) and not hn and i<5 else '')+('<span class="evt">SQ</span>' if day==sq else '')
        items=''.join(f'<div class="ev{" tbd" if t=="未定" else ""}"><span class="t">{t}</span> {html.escape(n)}<span class="c">{c}</span></div>' for t,n,c in ev.get(d,[]))
        cells+=f'<td class="{cls}{" has" if d in ev else ""}"><div class="d">{d}{lab}</div>{items}</td>'
    cells+='</tr>'
notes=[]
if undated: notes.append('<b>日付が未確定:</b> '+' / '.join(undated))
if tbd: notes.append('<b>時刻が未定:</b> '+'・'.join(tbd)+'は 16:00 と 19:30 に評価')
if not csv_ok: notes.append('<b>注意:</b> 内閣府の祝日 CSV が取れなかったため、祝日が抜けている可能性があります')
page=f"""<!doctype html><html><head><meta charset="utf-8"><style>
body{{margin:0;background:#fff;font-family:'IPAPGothic','IPAGothic',sans-serif;color:#1f2933}}
.wrap{{padding:28px 32px;width:1180px}} h1{{font-size:28px;margin:0 0 4px}} .sub{{color:#5b6670;font-size:14px;margin-bottom:16px}}
table{{border-collapse:collapse;width:100%;table-layout:fixed}} col.we{{width:46px}}
th{{background:#24364a;color:#fff;padding:8px;font-size:15px}} th.sat{{background:#2f5d8a}} th.sun{{background:#9c3b3b}}
td{{border:1px solid #d5dbe1;vertical-align:top;height:150px;padding:6px 7px}} td.we{{padding:6px 4px;text-align:center}}
td.empty{{background:#f5f6f8}} td.has{{background:#fbfaf3}}
.d{{font-weight:bold;font-size:17px;margin-bottom:4px}} .sat .d{{color:#2f5d8a}} .sun .d{{color:#b03a3a}}
.evt{{display:inline-block;margin-left:6px;padding:1px 7px;font-size:12px;font-weight:bold;color:#fff;background:#6b46c1;border-radius:9px;vertical-align:2px}}
.hol{{font-size:11px;font-weight:normal;margin-left:6px}}
.ev{{font-size:13.5px;line-height:1.35;margin:2px 0;padding:2px 4px;border-left:3px solid #c98a1b;background:#fff4dc;border-radius:2px;letter-spacing:-0.2px}}
.ev.tbd{{border-left-color:#8a96a3;background:#eef1f4}} .t{{font-weight:bold}} .c{{color:#6b7580;font-size:12px;margin-left:4px}}
.foot{{margin-top:14px;font-size:13px;line-height:1.7;color:#3d4852}} .foot b{{color:#1f2933}}
</style></head><body><div class="wrap"><h1>決算発表スケジュール {Y}年{M}月</h1>
<div class="sub">{len(codes)}銘柄のうち{M}月発表分 / 時刻は SBI証券の予定時刻(前年同期の開示時刻) / 祝日は内閣府の CSV / {sys.argv[5]} 時点</div>
<table><colgroup>{'<col class="wd">'*5}<col class="we"><col class="we"></colgroup><tr><th>月</th><th>火</th><th>水</th><th>木</th><th>金</th><th class="sat">土</th><th class="sun">日</th></tr>{cells}</table>
<div class="foot">{' ／ '.join(notes)}</div></div></body></html>"""
html_path=os.path.join(tempfile.mkdtemp(),'cal.html')
open(html_path,'w',encoding='utf-8').write(page)
shot="""const {chromium}=require('playwright');(async()=>{const b=await chromium.launch();
const p=await b.newPage({viewport:{width:1244,height:900},deviceScaleFactor:2});
await p.goto('file://'+process.argv[1]);await (await p.$('.wrap')).screenshot({path:process.argv[2]});await b.close();})();"""
subprocess.run(['node','-e',shot,html_path,os.path.abspath(sys.argv[4])],check=True)
print('holidays:',[ (d.isoformat(),n) for d,n in hol.items() if d.year==Y and d.month==M],'SQ:',sq,'csv_ok:',csv_ok,'events:',sum(len(v) for v in ev.values()))
