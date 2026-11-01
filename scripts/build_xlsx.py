import json, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.comments import Comment
from openpyxl.chart import BarChart, LineChart, Reference

"""data/*.json 으로 엑셀 마스터(downloads/jepq_dividend_tracker.xlsx)를 만든다.
수식은 엑셀·구글시트가 파일을 열 때 계산한다(fullCalcOnLoad)."""
from pathlib import Path
import calendar
ROOT=Path(__file__).resolve().parent.parent
J=lambda n: json.loads((ROOT/'data'/f'{n}.json').read_text(encoding='utf-8'))
META=J('meta'); KRJ=J('kr_etf'); HOLDJ=J('holdings')
# JEPQ는 매월 분배. 1월분은 12월 말에 배당락하고 1월 초에 지급하므로 '분배월'은 지급일의 연·월로 센다.
rows=[dict(ex=d['ex'],rec=d['record'],pay=d['pay'],amt=d['amount'],close=d.get('close'),fx=d.get('fx'),src=d['source']) for d in J('dividends')]
LAST=rows[-1]; MAXY=int(LAST['pay'][:4]); LASTM=int(LAST['pay'][5:7]); FIRSTY=int(rows[0]['pay'][:4])
order=KRJ['order']; names={c:KRJ['etfs'][c]['name'] for c in order}
fee={c:KRJ['etfs'][c]['fee'] for c in order}; cost={c:KRJ['etfs'][c].get('cost') for c in order}; timing={c:KRJ['etfs'][c]['timing'] for c in order}
kr={c:[dict(gijunYmd=h[0],divAmt=h[1],divRt=h[2]) for h in KRJ['etfs'][c]['history']] for c in order}
_lk=max(KRJ['etfs'][c]['history'][-1][0] for c in order); _ky,_km=int(_lk[:4]),int(_lk[4:6])
KR_ASOF=dt.date(_ky,_km,calendar.monthrange(_ky,_km)[1])
FONT='맑은 고딕'
f_in=Font(name=FONT,size=10,color='0000FF')
f_fx=Font(name=FONT,size=10,color='000000')
f_link=Font(name=FONT,size=10,color='008000')
f_hd=Font(name=FONT,size=10,bold=True,color='FFFFFF')
f_t=Font(name=FONT,size=14,bold=True,color='1F2A44')
f_b=Font(name=FONT,size=10,bold=True)
f_n=Font(name=FONT,size=10)
f_s=Font(name=FONT,size=9,color='666666')
hd_fill=PatternFill('solid',fgColor='1F2A44')
in_fill=PatternFill('solid',fgColor='FFF9DB')
key_fill=PatternFill('solid',fgColor='FFFF00')
band=PatternFill('solid',fgColor='F4F6FA')
thin=Side(style='thin',color='D0D5DD'); bd=Border(bottom=thin)
C=Alignment(horizontal='center',vertical='center',wrap_text=True)
R=Alignment(horizontal='right',vertical='center')
USD4='$0.0000;-$0.0000;-'; PCT='0.00%;-0.00%;-'; KRW='#,##0.0"원";-#,##0.0"원";-'; KRW0='#,##0"원";-#,##0"원";-'
DATE='yyyy-mm-dd'

def header(ws,row,labels,widths=None):
    for i,t in enumerate(labels,1):
        c=ws.cell(row,i,t); c.font=f_hd; c.fill=hd_fill; c.alignment=C
    ws.row_dimensions[row].height=34
    if widths:
        for i,w in enumerate(widths,1): ws.column_dimensions[L(i)].width=w

wb=Workbook()
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.chart.label import DataLabelList
OF=META.get('official') or {}; SCH=(META.get('schedule') or {}).get('rows',[])
# ---------------- 사용법 ----------------
G=wb.active; G.title='사용법'
G.column_dimensions['A'].width=3; G.column_dimensions['B'].width=26; G.column_dimensions['C'].width=90
G['B2']='JEPQ 분배 트래커 (마스터)'; G['B2'].font=f_t
G['B3']=f"최종 갱신: {META['updated']} · 반영 범위: {rows[0]['pay'][:7]} ~ {LAST['pay'][:7]} 지급분 (배당락 {LAST['ex']}) · GitHub Actions 자동 생성"; G['B3'].font=f_s
r=5
G.cell(r,2,'설정값').font=f_b; r+=1
G.cell(r,2,'미국 원천징수율').font=f_n
c=G.cell(r,3,META['withholding']); c.font=f_in; c.fill=key_fill; c.number_format='0.0%'; c.alignment=Alignment(horizontal='left')
c.comment=Comment('미국 상장 ETF 분배금: 한국 거주자는 미국에서 15% 원천징수 (한미 조세조약). 국내 추가 원천징수 없음, 금융소득종합과세(연 2,000만원 초과) 합산 대상.','Claude')
WHT='사용법!$C$6'
r+=2
G.cell(r,2,'색상 규칙').font=f_b; r+=1
for k,v,font,fill in [('파란 글씨 + 연노랑 칸','직접 입력하는 값 (공시 분배금·날짜·종가·환율·분배금)',f_in,in_fill),('검정 글씨','수식 — 건드리지 마세요. 자동 계산됩니다',f_fx,None),('노란 칸','핵심 설정값',f_in,key_fill)]:
    a=G.cell(r,2,k); a.font=font
    if fill: a.fill=fill
    G.cell(r,3,v).font=f_n; r+=1
r+=1
G.cell(r,2,'월 갱신 절차 (5분)').font=f_b; r+=1
steps=[
 ('① JEPQ 분배 입력','[JEPQ_배당] 시트 맨 아래 빈 줄(수식이 미리 깔린 줄)에 C~E, I, K, N 열만 입력합니다.\n  C 배당락일 · D 지급일 · E 공시 분배금 · I 배당락일 종가 · K 지급일 환율 · N 출처'),
 ('② 국내 JEPQ형 분배 입력','[국내JEPQ형_분배] 시트 맨 아래에 3종목의 월 분배금을 붙여넣습니다 (종목코드·종목명·기준일·분배금·분배율).'),
 ('③ 비교 기준일 변경','[국내JEPQ형_비교] B3 칸의 기준일을 월말로 바꿉니다 (자동 생성 시 최신 분배월 말일로 맞춰짐).'),
 ('④ 결과 확인','[최신월_요약] 시트가 자동으로 최신 분배 기준으로 바뀝니다. 카페 글 숫자는 여기서 가져가면 됩니다.'),
 ('⑤ 파일명','JEPQ_분배트래커_YY년MM월_YYMMDD.xlsx 형식으로 저장'),
]
for k,v in steps:
    G.cell(r,2,k).font=f_n; c=G.cell(r,3,v); c.font=f_n; c.alignment=Alignment(wrap_text=True,vertical='top'); G.row_dimensions[r].height=32 if '\n' in v else 18; r+=1
r+=1
G.cell(r,2,'데이터 출처').font=f_b; r+=1
PR=J('prices')
src=[
 ('JEPQ 최신 분배','J.P. Morgan Asset Management JEPQ 상품 데이터(공식) — 다음 분배가 발표되면 배당락 전에 표시됨'),
 ('JEPQ 분배 이력 (2022.06~)','Nasdaq.com 분배 이력(배당락일·기준일·지급일·금액 소수 5자리). JPMorgan은 과거 이력 파일을 따로 제공하지 않음. Yahoo Finance(소수 3자리)·stockanalysis(최근 20회)와 교차 확인'),
 ('배당락일 종가',f"{PR.get('source','')} — 분배 미조정 종가"),
 ('지급일 환율','하나은행 고시 매매기준율 (다음금융 일별 환율, 지급일 당일 최종 회차)'),
 ('공식 수익률·NAV','J.P. Morgan Asset Management JEPQ 상품 데이터 — 30일 SEC 수익률·12개월 롤링 분배수익률 (기준일 함께 표기)'),
 ('분배 일정','J.P. Morgan Exchange-Traded Fund Trust Dividend Calendar (연 1회 PDF)'),
 ('국내 JEPQ형 분배금','FunETF 분배금 이력 (기준일·분배금·분배율) — 운용사 공시와 같은 원천'),
 ('국내 ETF 실부담비용',KRJ.get('cost_source','')+' — 기준일은 [국내JEPQ형_비교] 시트 메모'),
 ('국내 ETF 총보수(참고)',KRJ['fee_source']),
 ('구성종목·월별 비중','JPMorgan 공식 일별 보유종목 엑셀(Daily Holdings) — 매일 자동 수집. 과거 날짜 파일은 제공되지 않아 2026-10부터 매월 첫 기록을 쌓습니다'),
]
for k,v in src:
    G.cell(r,2,k).font=f_n; c=G.cell(r,3,v); c.font=f_n; c.alignment=Alignment(wrap_text=True); r+=1
r+=1
G.cell(r,2,'알아둘 점').font=f_b; r+=1
notes=[
 '· JEPQ는 매월 분배합니다. 1월분은 전년 12월 말에 배당락하고 1월 초에 지급하므로, 이 파일의 연도·월은 지급일 기준입니다.',
 '· JEPQ 분배금은 주식 배당 + ELN(나스닥100 콜옵션 매도) 수익에서 나와 매달 금액이 달라집니다. 성장률은 \'배당 성장\'이 아니라 옵션 프리미엄 변화의 영향이 큽니다.',
 '· 연간 성장률은 12회 분배가 모두 확정된 해만 계산합니다. 2022년은 상장(5월) 첫해라 7회뿐이고, 진행 중인 해는 [연도별] 시트의 \'올해 누적 vs 작년 동기\'로 봅니다.',
 '· 진행 중인 해의 남은 달 \'예상\' 값은 그해 확정 월의 단순 평균입니다. 옵션 프리미엄은 변동성에 따라 크게 달라지니 참고용입니다.',
 '· 이 파일은 GitHub에서 새 분배가 반영될 때마다 다시 만들어집니다(차트 범위도 자동). 손으로 고친 내용은 다음 생성 때 덮어써지니, 고칠 값은 data 폴더 JSON에서 고치세요.',
 '· 국내 JEPQ형 ETF 분배율은 매월 기준일 가격 대비 분배율을 12개월 합산한 값입니다(근사치).',
]
for n in notes:
    c=G.cell(r,2,n); c.font=f_n; G.merge_cells(start_row=r,start_column=2,end_row=r,end_column=3); r+=1

# ---------------- JEPQ_배당 ----------------
S=wb.create_sheet('JEPQ_배당')
heads=['연도\n(지급)','월\n(지급)','배당락일','지급일','공시 분배금\n(USD)','전월\n대비','전년 동월\n대비','최근 12회\n합계(TTM)','배당락일\n종가(USD)','TTM\n분배수익률','지급일\n환율(원)','1주 분배\n원화(세전)','1주 분배\n원화(미국15%후)','출처']
header(S,1,heads,[7,6,11.5,11.5,13,9,9,11,11,10,10,11,12,22])
S.freeze_panes='C2'
N=len(rows); FUT=60; last=1+N+FUT
for i in range(N+FUT):
    rr=i+2; d=rows[i] if i<N else None
    S.cell(rr,1,f'=IF($D{rr}="","",YEAR($D{rr}))').number_format='0'
    S.cell(rr,2,f'=IF($D{rr}="","",MONTH($D{rr}))').number_format='0"월"'
    for col in (3,4,5,9,11,14):
        c=S.cell(rr,col); c.font=f_in; c.fill=in_fill
    if d:
        S.cell(rr,3,dt.date.fromisoformat(d['ex'])); S.cell(rr,4,dt.date.fromisoformat(d['pay']))
        S.cell(rr,5,d['amt']); S.cell(rr,9,d['close']); S.cell(rr,11,d['fx']); S.cell(rr,14,d['src'])
    S.cell(rr,3).number_format=DATE; S.cell(rr,4).number_format=DATE
    S.cell(rr,5).number_format='$0.00000'; S.cell(rr,9).number_format='$0.00'; S.cell(rr,11).number_format='#,##0.0'
    S.cell(rr,6,'' if rr==2 else f'=IF(OR($C{rr}="",E{rr-1}=""),"",E{rr}/E{rr-1}-1)').number_format=PCT
    S.cell(rr,7,'' if rr<14 else f'=IF(OR($C{rr}="",E{rr-12}=""),"",E{rr}/E{rr-12}-1)').number_format=PCT
    S.cell(rr,8,'' if rr<13 else f'=IF($C{rr}="","",SUM(E{rr-11}:E{rr}))').number_format=USD4
    S.cell(rr,10,f'=IF(OR(H{rr}="",I{rr}=""),"",H{rr}/I{rr})').number_format=PCT
    S.cell(rr,12,f'=IF(OR(E{rr}="",K{rr}=""),"",E{rr}*K{rr})').number_format=KRW
    S.cell(rr,13,f'=IF(L{rr}="","",L{rr}*(1-{WHT}))').number_format=KRW
    for col in range(1,15):
        c=S.cell(rr,col)
        if col not in (3,4,5,9,11,14): c.font=f_fx
        c.border=bd
        if col in (1,2): c.alignment=Alignment(horizontal='center')
S['E1'].comment=Comment('운용사(JPMorgan) 공시 금액 그대로 입력. 과거분은 Nasdaq.com 이력(소수 5자리).','Claude')
S['B1'].comment=Comment('지급일의 달. 1월분은 12월 말 배당락·1월 초 지급이라 지급일 기준으로 셉니다.','Claude')
S['K1'].comment=Comment('지급일 하나은행 매매기준율(최종 회차). 실제 입금 환율은 증권사·지급 시점에 따라 다릅니다.','Claude')
S['M1'].comment=Comment('미국 원천징수율은 [사용법] C6 칸에서 변경','Claude')
MS="JEPQ_배당"; RNG=lambda col: f"{MS}!${col}$2:${col}${last}"

# ---------------- 연도별 ----------------
Y=wb.create_sheet('연도별')
header(Y,1,['연도','분배\n횟수','연간 분배\n(USD)','상태','전년 대비\n증감률','','올해 누적\nvs 작년 동기','연간 분배\n원화(세전)','미지급 월\n예상분','연간 분배\n(예상 포함)','예상\n증감률'],[8,8,15,15,11,3,13,14,12,13,10])
Y.freeze_panes='A2'
years=list(range(FIRSTY,MAXY+6))
for i,yv in enumerate(years):
    rr=i+2
    Y.cell(rr,1,yv).font=f_fx; Y.cell(rr,1).alignment=Alignment(horizontal='center')
    Y.cell(rr,2,f'=COUNTIFS({RNG("A")},A{rr})')
    Y.cell(rr,3,f'=IF(B{rr}=0,"",SUMIFS({RNG("E")},{RNG("A")},A{rr}))').number_format=USD4
    st=f'=IF(B{rr}=0,"",IF(B{rr}>=12,"확정","진행중("&B{rr}&"회)"))' if yv!=FIRSTY else f'=IF(B{rr}=0,"","상장 첫해("&B{rr}&"회)")'
    Y.cell(rr,4,st)
    Y.cell(rr,5,'' if rr==2 else f'=IF(AND(B{rr}>=12,B{rr-1}>=12),C{rr}/C{rr-1}-1,"")').number_format=PCT
    Y.cell(rr,7,'' if rr==2 else f'=IF(AND(B{rr}>0,B{rr}<12,B{rr-1}>=12),C{rr}/SUMIFS({RNG("E")},{RNG("A")},A{rr-1},{RNG("B")},"<="&_xlfn.MAXIFS({RNG("B")},{RNG("A")},A{rr}))-1,"")').number_format=PCT
    Y.cell(rr,8,f'=IF(B{rr}=0,"",SUMIFS({RNG("L")},{RNG("A")},A{rr}))').number_format=KRW
    Y.cell(rr,9,'' if yv==FIRSTY else f'=IF(AND(B{rr}>0,B{rr}<12),(12-B{rr})*C{rr}/B{rr},"")').number_format=USD4
    Y.cell(rr,10,f'=IF(B{rr}=0,"",C{rr}+IF(I{rr}="",0,I{rr}))').number_format=USD4
    Y.cell(rr,11,'' if rr==2 else f'=IF(AND(I{rr}<>"",B{rr-1}>=12),J{rr}/C{rr-1}-1,"")').number_format=PCT
    for col in (9,10,11): Y.cell(rr,col).fill=PatternFill('solid',fgColor='FBEFDF')
    for col in range(1,12):
        if col==6: continue
        c=Y.cell(rr,col); c.border=bd
        if col>1: c.font=f_fx
        if col in (2,4): c.alignment=Alignment(horizontal='center')
Yl=len(years)+1
Y['H1'].comment=Comment('각 달 지급일 환율로 환산한 세전 원화 합계','Claude')
Y['G1'].comment=Comment('올해 지급된 달까지의 누적 ÷ 작년 같은 달까지의 누적 − 1','Claude')
b=Yl+2
Y.cell(b,1,'핵심 지표 (최근 확정연도 기준, 자동)').font=f_b
YR=lambda col: f'$'+col+'$2:$'+col+f'${Yl}'
items=[
 ('최근 확정연도', f'=_xlfn.MAXIFS({YR("A")},{YR("B")},">=12")','0'),
 ('그해 연간 분배', f'=INDEX({YR("C")},MATCH(B{b+1},{YR("A")},0))',USD4),
 ('그해 증감률', f'=INDEX({YR("E")},MATCH(B{b+1},{YR("A")},0))',PCT),
 ('확정연도 평균 연간 분배', f'=AVERAGEIFS({YR("C")},{YR("B")},">=12")',USD4),
 ('확정연도 수', f'=COUNTIFS({YR("B")},">=12")','0"년"'),
 ('진행 중인 해', f'=MAX({RNG("A")})','0'),
 ('  예상 연간 분배 (남은 달 = 올해 확정 월 평균)', f'=IF(INDEX({YR("I")},MATCH(B{b+6},{YR("A")},0))="","",INDEX({YR("J")},MATCH(B{b+6},{YR("A")},0)))',USD4),
 ('  예상 증감률 (남은 달 = 올해 확정 월 평균)', f'=IF(B{b+7}="","",B{b+7}/INDEX({YR("C")},MATCH(B{b+6}-1,{YR("A")},0))-1)',PCT),
]
for j,(k,fml,fmt) in enumerate(items,1):
    Y.cell(b+j,1,k).font=f_n
    c=Y.cell(b+j,2,fml); c.number_format=fmt; c.font=f_fx
Y.column_dimensions['A'].width=40
Y.cell(b+len(items)+1,1,'※ 2022년은 5월 상장 첫해라 6~12월 7회만 있어 연간 비교에서 뺍니다.').font=f_s
Y.cell(b+len(items)+2,1,'※ 예상치: 진행 중인 해의 남은 달을 그해 확정 월 평균으로 채운 단순 계산. 옵션 프리미엄은 시장 변동성에 따라 달라 실제와 차이가 큽니다.').font=f_s
Y['I1'].comment=Comment('진행 중인 해의 미지급 달을, 그해 이미 확정된 월 분배의 단순 평균으로 채운 값 (예측치)','Claude')
# chart helper (M~T)
for j,t in enumerate(['차트용\n연도','확정\n분배','예상\n추가분','확정\n증감률','예상\n증감률','확정연도\n평균 분배','평균\n증감률','연간 합계\n(예상 포함)'],13):
    c=Y.cell(1,j,t); c.font=f_hd; c.fill=PatternFill('solid',fgColor='5E6E78'); c.alignment=C
    Y.column_dimensions[L(j)].width=9
for i,yv in enumerate(years):
    rr=i+2
    Y.cell(rr,13,f'=IF(B{rr}=0,"",IF(A{rr}={FIRSTY},TEXT(A{rr},"0")&"*",TEXT(A{rr},"0")))')
    Y.cell(rr,14,f'=IF(B{rr}=0,"",C{rr})').number_format=USD4
    Y.cell(rr,15,f'=IF(I{rr}="","",I{rr})').number_format=USD4
    Y.cell(rr,16,f'=IF(E{rr}="","",E{rr})').number_format=PCT
    Y.cell(rr,17,f'=IF(K{rr}="","",K{rr})').number_format=PCT
    Y.cell(rr,18,f'=IF(B{rr}=0,"",AVERAGEIFS({YR("C")},{YR("B")},">=12"))').number_format=USD4
    Y.cell(rr,19,f'=IF(B{rr}=0,"",IFERROR(AVERAGE({YR("E")}),""))').number_format=PCT
    Y.cell(rr,20,f'=IF(B{rr}=0,"",J{rr})').number_format=USD4
    for col in range(13,21): Y.cell(rr,col).font=f_s
CH_END=MAXY-FIRSTY+2
def avg_line(ref,cats):
    lc=LineChart(); lc.add_data(ref,titles_from_data=True); lc.set_categories(cats)
    sr=lc.series[0]; sr.marker.symbol='none'; sr.smooth=False
    sr.graphicalProperties=GraphicalProperties(); sr.graphicalProperties.line=LineProperties(solidFill='2F5D8A',w=22000,prstDash='dash')
    return lc
def label_line(ref,cats,fmt):
    lc=LineChart(); lc.add_data(ref,titles_from_data=True); lc.set_categories(cats)
    sr=lc.series[0]; sr.marker.symbol='none'
    sr.graphicalProperties=GraphicalProperties(); sr.graphicalProperties.line=LineProperties(noFill=True)
    sr.dLbls=DataLabelList(); sr.dLbls.showVal=True; sr.dLbls.position='t'; sr.dLbls.numFmt=fmt
    for k in ('showSerName','showCatName','showLegendKey','showPercent'): setattr(sr.dLbls,k,False)
    return lc
def solid(series,hexc,dash=False):
    series.graphicalProperties=GraphicalProperties(solidFill=hexc)
    series.graphicalProperties.line=LineProperties(solidFill='A8661B' if dash else hexc, prstDash='dash' if dash else None)
ch=BarChart(); ch.type='col'; ch.grouping='stacked'; ch.overlap=100; ch.height=9; ch.width=20
ch.add_data(Reference(Y,min_col=14,min_row=1,max_row=CH_END),titles_from_data=True)
ch.add_data(Reference(Y,min_col=15,min_row=1,max_row=CH_END),titles_from_data=True)
ch.set_categories(Reference(Y,min_col=13,min_row=2,max_row=CH_END))
solid(ch.series[0],'C98A3A'); solid(ch.series[1],'F1D9B5',dash=True)
ch.y_axis.numFmt='$0.00'; ch.y_axis.majorGridlines=None; ch.legend.position='b'
ch.y_axis.delete=False; ch.x_axis.delete=False
_cats=Reference(Y,min_col=13,min_row=2,max_row=CH_END)
ch+=avg_line(Reference(Y,min_col=18,min_row=1,max_row=CH_END),_cats)
ch+=label_line(Reference(Y,min_col=20,min_row=1,max_row=CH_END),_cats,'$0.0000')
ch.title=f'JEPQ 연간 분배 (USD) · 진한색 확정 / 연한색 예상 / 점선 확정연도 평균 · {FIRSTY}*는 상장 첫해(7회)'
Y.add_chart(ch,'V2')
g=BarChart(); g.type='col'; g.grouping='clustered'; g.overlap=100; g.height=9; g.width=20
g.add_data(Reference(Y,min_col=16,min_row=1,max_row=CH_END),titles_from_data=True)
g.add_data(Reference(Y,min_col=17,min_row=1,max_row=CH_END),titles_from_data=True)
g.set_categories(Reference(Y,min_col=13,min_row=2,max_row=CH_END))
solid(g.series[0],'17784A'); solid(g.series[1],'F1D9B5',dash=True)
g.y_axis.numFmt='0%'; g.legend.position='b'; g.y_axis.delete=False; g.x_axis.delete=False
g.title='JEPQ 연간 분배 증감률 · 진한색 확정 / 연한색 예상'
Y.add_chart(g,'V22')

# ---------------- 월별표 ----------------
Q=wb.create_sheet('월별표')
header(Q,1,['연도']+[f'{m}월' for m in range(1,13)]+['연간','']+['연도']+[f'{m}월\n전년비' for m in range(1,13)],
       [8]+[9]*12+[10,3]+[8]+[8]*12)
Q.cell(1,15).fill=PatternFill(None)
for i,yv in enumerate(years):
    rr=i+2
    Q.cell(rr,1,yv); Q.cell(rr,16,yv)
    for m in range(1,13):
        Q.cell(rr,1+m,f'=IF(COUNTIFS({RNG("A")},$A{rr},{RNG("B")},{m})=0,"",SUMIFS({RNG("E")},{RNG("A")},$A{rr},{RNG("B")},{m}))').number_format=USD4
        if rr>2:
            col=L(1+m)
            Q.cell(rr,16+m,f'=IF(OR({col}{rr}="",{col}{rr-1}=""),"",{col}{rr}/{col}{rr-1}-1)').number_format=PCT
    Q.cell(rr,14,f'=IF(SUM(B{rr}:M{rr})=0,"",SUM(B{rr}:M{rr}))').number_format=USD4
    for col in range(1,29):
        if col==15: continue
        c=Q.cell(rr,col); c.font=f_fx; c.border=bd
        if col in (1,16): c.alignment=Alignment(horizontal='center')
Q.freeze_panes='B2'
Q.cell(len(years)+3,1,'※ 월 = 지급일의 달 (1월분은 전년 12월 말 배당락, 1월 초 지급)').font=f_s

# ---------------- 국내JEPQ형_분배 ----------------
K=wb.create_sheet('국내JEPQ형_분배')
header(K,1,['종목코드','종목명','기준일','분배금(원)','분배율','비고'],[10,34,12,11,9,30])
K.freeze_panes='A2'
recs=[]
for code in order:
    for x in sorted(kr[code],key=lambda z:z['gijunYmd']):
        recs.append((code,names[code],dt.datetime.strptime(x['gijunYmd'],'%Y%m%d').date(),x['divAmt'],x['divRt']/100))
for i,(code,nm,d,amt,rt) in enumerate(recs):
    rr=i+2
    for col,v,fmt in [(1,code,'@'),(2,nm,None),(3,d,DATE),(4,amt,'#,##0'),(5,rt,'0.00%')]:
        c=K.cell(rr,col,v); c.font=f_in; c.fill=in_fill; c.border=bd
        if fmt: c.number_format=fmt
KL=len(recs)+1+200
for rr in range(len(recs)+2,KL+1):
    for col in range(1,6):
        c=K.cell(rr,col); c.fill=in_fill; c.font=f_in; c.border=bd
    K.cell(rr,1).number_format='@'; K.cell(rr,3).number_format=DATE; K.cell(rr,4).number_format='#,##0'; K.cell(rr,5).number_format='0.00%'
K['C1'].comment=Comment('분배금 지급 기준일(권리 확정일). 출처: FunETF 분배 이력','Claude')
K['E1'].comment=Comment('기준일 가격 대비 1회 분배율','Claude')
KR=lambda col: f"국내JEPQ형_분배!${col}$2:${col}${KL}"

# ---------------- 국내JEPQ형_비교 ----------------
V=wb.create_sheet('국내JEPQ형_비교')
V['A1']='국내상장 JEPQ형(나스닥100 커버드콜) ETF vs JEPQ'; V['A1'].font=f_t
V['A3']='비교 기준일'; V['A3'].font=f_b
V['B3']=KR_ASOF; V['B3'].font=f_in; V['B3'].fill=key_fill; V['B3'].number_format=DATE
V['C3']='← 매월 월말로 바꾸세요. 아래 표가 이 날짜 기준 최근 12개월로 다시 계산됩니다.'; V['C3'].font=f_s
hrow=5
header(V,hrow,['종목','코드','실부담비용\n(연)','분배\n시점','최근 분배일','최근 분배금','12개월\n분배 횟수','12개월\n분배금 합계','12개월\n분배율 합계','기초지수'],[34,9,9,10,12,11,9,12,11,44])
V.row_dimensions[hrow].height=34
for j,code in enumerate(order):
    rr=hrow+1+j
    V.cell(rr,1,names[code]).font=f_n
    c=V.cell(rr,2,code); c.number_format='@'; c.font=f_n
    c=V.cell(rr,3,round(cost[code]['real']/100,8) if cost[code] and cost[code].get('real') is not None else fee[code]); c.font=f_in; c.fill=in_fill; c.number_format='0.0000%'
    if cost[code] and cost[code].get('real') is not None: V.cell(rr,3).comment=Comment(f"총보수 {cost[code]['total_fee']}% + 기타비용 {cost[code]['other']}% + 매매·중개수수료 {cost[code]['trading']}% (KOFIA {cost[code]['basis']})",'tracker')
    c=V.cell(rr,4,timing[code]); c.font=f_in; c.fill=in_fill
    V.cell(rr,5,f'=_xlfn.MAXIFS({KR("C")},{KR("A")},$B{rr},{KR("C")},"<="&$B$3)').number_format=DATE
    V.cell(rr,6,f'=SUMIFS({KR("D")},{KR("A")},$B{rr},{KR("C")},E{rr})').number_format=KRW0
    V.cell(rr,7,f'=COUNTIFS({KR("A")},$B{rr},{KR("C")},">"&EDATE($B$3,-12),{KR("C")},"<="&$B$3)')
    V.cell(rr,8,f'=SUMIFS({KR("D")},{KR("A")},$B{rr},{KR("C")},">"&EDATE($B$3,-12),{KR("C")},"<="&$B$3)').number_format=KRW0
    V.cell(rr,9,f'=SUMIFS({KR("E")},{KR("A")},$B{rr},{KR("C")},">"&EDATE($B$3,-12),{KR("C")},"<="&$B$3)').number_format=PCT
    V.cell(rr,10,KRJ['etfs'][code].get('index','')).font=f_s
    for col in range(1,11):
        V.cell(rr,col).border=bd
        if 5<=col<=9: V.cell(rr,col).font=f_fx
sr=hrow+len(order)+2
V.cell(sr,1,'JEPQ (미국 직투, 참고)').font=f_b
V.cell(sr,2,'JEPQ')
c=V.cell(sr,3,META['jepq_fee']); c.font=f_in; c.fill=in_fill; c.number_format='0.00%'
V.cell(sr,4,'월(배당락 월초)')
V.cell(sr,5,f'=_xlfn.MAXIFS({RNG("C")},{RNG("C")},"<="&$B$3)').number_format=DATE
V.cell(sr,6,f'=INDEX({RNG("E")},MATCH(E{sr},{RNG("C")},0))').number_format=USD4
V.cell(sr,7,f'=COUNTIFS({RNG("C")},">"&EDATE($B$3,-12),{RNG("C")},"<="&$B$3)')
V.cell(sr,8,f'=INDEX({RNG("H")},MATCH(E{sr},{RNG("C")},0))').number_format=USD4
V.cell(sr,9,f'=INDEX({RNG("J")},MATCH(E{sr},{RNG("C")},0))').number_format=PCT
V.cell(sr,10,'Nasdaq-100 주식 + ELN(나스닥100 콜옵션 매도), 액티브').font=f_s
for col in range(1,11): V.cell(sr,col).border=bd
V.cell(sr+1,1,'※ JEPQ 행: 금액은 USD, 수익률은 최근 12회 분배 ÷ 배당락일 종가. 0.35%는 JPMorgan 공시 순보수(펀드 내부 매매비용 미포함)라 국내 ETF 실부담비용과 기준이 다릅니다. 투자자가 내는 증권사 수수료·환전 비용은 양쪽 모두 제외.').font=f_s
_cb=max((v['basis'] for v in cost.values() if v and v.get('real') is not None),default='')
V.cell(sr+2,1,f"※ 국내 ETF 실부담비용 = 총보수 + 기타비용 + 매매·중개수수료(+판매수수료), 금융투자협회 공시 {_cb} 기준. 셀 메모에 항목별 금액.").font=f_s
V.cell(sr+3,1,'※ 국내 3종은 옵션 전략(기초지수)이 서로 다르고 JEPQ와도 다릅니다. 분배율만으로 우열을 가릴 수 없고, 같은 기간 총수익률을 함께 보세요.').font=f_s
V.cell(sr+4,1,'※ 국내상장 해외주식형 ETF 분배금은 배당소득세 15.4% 과세 (연금저축·IRP·ISA에서는 과세이연/비과세 혜택).').font=f_s

# monthly table
mr=sr+7
V.cell(mr-1,1,'월별 분배금 (국내: 기준일의 달 · JEPQ: 지급일의 달)').font=f_b
header_cells=['월','시작일','종료일']+[f'{c}\n{names[c]}' for c in order]+['JEPQ\n(USD)','JEPQ 원화\n(세전)']
for i,t in enumerate(header_cells,1):
    c=V.cell(mr,i,t); c.font=f_hd; c.fill=hd_fill; c.alignment=C
V.row_dimensions[mr].height=62
V.column_dimensions['B'].width=11; V.column_dimensions['C'].width=11
_k0=min(min(x['gijunYmd'] for x in kr[c]) for c in order)
months=[(y,m) for y in range(int(_k0[:4]),MAXY+2) for m in range(1,13) if (y,m)>=(int(_k0[:4]),int(_k0[4:6]))]
nk=len(order)
for i,(yv,m) in enumerate(months):
    rr=mr+1+i
    V.cell(rr,1,f'{yv}-{m:02d}').alignment=Alignment(horizontal='center')
    V.cell(rr,2,dt.date(yv,m,1)).number_format=DATE
    V.cell(rr,3,f'=EDATE(B{rr},1)-1').number_format=DATE
    for j,code in enumerate(order):
        V.cell(rr,4+j,f'=IF(COUNTIFS({KR("A")},"{code}",{KR("C")},">="&$B{rr},{KR("C")},"<="&$C{rr})=0,"",SUMIFS({KR("D")},{KR("A")},"{code}",{KR("C")},">="&$B{rr},{KR("C")},"<="&$C{rr}))').number_format=KRW0
    V.cell(rr,4+nk,f'=IF(COUNTIFS({RNG("D")},">="&$B{rr},{RNG("D")},"<="&$C{rr})=0,"",SUMIFS({RNG("E")},{RNG("D")},">="&$B{rr},{RNG("D")},"<="&$C{rr}))').number_format=USD4
    V.cell(rr,5+nk,f'=IF({L(4+nk)}{rr}="","",SUMIFS({RNG("L")},{RNG("D")},">="&$B{rr},{RNG("D")},"<="&$C{rr}))').number_format=KRW
    for col in range(1,6+nk):
        c=V.cell(rr,col); c.border=bd
        if col>1: c.font=f_fx
V.freeze_panes='B6'

# ---------------- 최신월_요약 ----------------
Z=wb.create_sheet('최신월_요약')
Z.column_dimensions['A'].width=34; Z.column_dimensions['B'].width=16; Z.column_dimensions['C'].width=56
Z['A1']='JEPQ 최신 분배 요약 (카페 글용 숫자)'; Z['A1'].font=f_t
Z['A2']='[JEPQ_배당]에 새 분배를 입력하면 자동으로 바뀝니다.'; Z['A2'].font=f_s
LR=f'(COUNT({RNG("C")})+1)'
IX=lambda col: f'=IF(INDEX({MS}!${col}:${col},{LR})="","",INDEX({MS}!${col}:${col},{LR}))'
Yn=lambda col,yr: f'INDEX(연도별!${col}$2:${col}${Yl},MATCH({yr},연도별!$A$2:$A${Yl},0))'
lines=[
 ('최신 분배', None, None),
 ('연도 (지급)', IX('A'),'0'),
 ('월 (지급)', IX('B'),'0"월"'),
 ('배당락일', IX('C'),DATE),
 ('지급일', IX('D'),DATE),
 ('분배금 (USD)', IX('E'),'$0.00000'),
 ('전월 대비', IX('F'),'+0.00%;-0.00%;0.00%'),
 ('전년 동월 대비', IX('G'),'+0.00%;-0.00%;0.00%'),
 ('전년 동월 분배금 (USD)', f'=INDEX({MS}!$E:$E,{LR}-12)','$0.00000'),
 ('최근 12회 합계 TTM (USD)', IX('H'),USD4),
 ('배당락일 종가 (USD)', IX('I'),'$0.00'),
 ('TTM 분배수익률', IX('J'),PCT),
 ('지급일 환율', IX('K'),'#,##0.0"원"'),
 ('1주 분배 원화 (세전)', IX('L'),KRW),
 ('1주 분배 원화 (미국 15% 원천징수 후)', IX('M'),KRW),
 ('', None,None),
 ('올해 흐름', None,None),
 ('올해 누적 분배 (USD)', f'={Yn("C","B4")}',USD4),
 ('작년 같은 기간 누적 (USD)', f'=SUMIFS({RNG("E")},{RNG("A")},B4-1,{RNG("B")},"<="&B5)',USD4),
 ('올해 누적 vs 작년 동기', f'=B20/B21-1','+0.00%;-0.00%;0.00%'),
 ('', None,None),
 ('연간 흐름', None,None),
 ('최근 확정연도', f'=_xlfn.MAXIFS(연도별!$A$2:$A${Yl},연도별!$B$2:$B${Yl},">=12")','0'),
 ('그해 연간 분배 (USD)', f'={Yn("C","B25")}',USD4),
 ('그해 증감률', f'={Yn("E","B25")}',PCT),
 ('확정연도 평균 연간 분배 (USD)', f'=AVERAGEIFS(연도별!$C$2:$C${Yl},연도별!$B$2:$B${Yl},">=12")',USD4),
]
r=3
for k,fml,fmt in lines:
    if fml is None:
        if k: Z.cell(r,1,k).font=Font(name=FONT,size=11,bold=True,color='1F2A44')
    else:
        Z.cell(r,1,k).font=f_n
        c=Z.cell(r,2,fml); c.number_format=fmt; c.font=f_fx; c.alignment=R; c.border=bd
        Z.cell(r,1).border=bd
    r+=1
assert Z['A4'].value.startswith('연도') and Z['A5'].value.startswith('월') and Z['A20'].value=='올해 누적 분배 (USD)' and Z['A21'].value.startswith('작년') and Z['A25'].value=='최근 확정연도' and Z['A26'].value.startswith('그해 연간')
Z['C9'].value='옵션 프리미엄에 따라 매달 오르내림'; Z['C9'].font=f_s
Z['C22'].value='같은 달(지급월)까지끼리 비교'; Z['C22'].font=f_s
r+=1
Z.cell(r,1,'공식 수익률 · 다음 분배').font=Font(name=FONT,size=11,bold=True,color='1F2A44'); r+=1
def zrow(k,v,fmt,note):
    global r
    Z.cell(r,1,k).font=f_n; Z.cell(r,1).border=bd
    c=Z.cell(r,2,v); c.number_format=fmt; c.font=f_fx; c.alignment=R; c.border=bd
    if note: Z.cell(r,3,note).font=f_s
    r+=1
if OF.get('sec_yield'):
    zrow('30일 SEC 수익률 (JPMorgan 공식)',OF['sec_yield']['value'],PCT,f"{OF['sec_yield']['asOf']} 기준 · 최근 30일 순투자수익(보수 차감) 연 환산. 옵션 프리미엄 대부분이 빠져 분배수익률보다 낮을 수 있음")
if OF.get('rolling_12m'):
    zrow('12개월 롤링 분배수익률 (JPMorgan 공식)',OF['rolling_12m']['value'],PCT,f"{OF['rolling_12m']['asOf']} 기준 · 최근 12회 분배금 ÷ 각 배당락일 NAV 합계")
if OF.get('nav'):
    zrow('NAV (USD)',OF['nav']['value'],'$0.00',f"{OF['nav']['asOf']} 기준")
nx=next((x for x in SCH if x['ex']>LAST['ex'] and x.get('kind')!='capgain'),None)
if nx:
    zrow('다음 배당락일 (예정)',dt.date.fromisoformat(nx['ex']),DATE,f"{nx['pay'][:7]} 지급분 · 금액 미발표 · JPMorgan 분배 일정 기준(변경 가능)")
    zrow('다음 지급일 (예정)',dt.date.fromisoformat(nx['pay']),DATE,None)
else:
    zrow('다음 배당락일 (예정)','일정 미공시','@',None)
cg=next((x for x in SCH if x['ex']>LAST['ex'] and x.get('kind')=='capgain'),None)
if cg:
    zrow('자본이득 분배 예정일 (있을 때만)',dt.date.fromisoformat(cg['ex']),DATE,f"지급 {cg['pay']} · 상장 이후 실제 자본이득 분배는 없었음")
_dl=J('dividends')[-1]
Z.cell(r,1,f"금액 출처: {LAST['src']}" + (f" · 트래커 반영 {_dl['added']}" if _dl.get('added') else '') + " · 실제 세후 입금액은 증권사 거래내역 기준").font=f_s

# ---------------- 차트데이터 (월별 분배 막대) ----------------
Q2=wb.create_sheet('차트데이터')
header(Q2,1,['연도','월','라벨','분배금','전체 평균'],[8,6,9,11,11])
Q2['G1']='월별 차트용 표 (지급월 기준)'; Q2['G1'].font=f_s
qrow=2
first_m=int(rows[0]['pay'][5:7])
for yv in range(FIRSTY,MAXY+1):
    for m in range(1,13):
        if (yv,m)<(FIRSTY,first_m) or (yv,m)>(MAXY,LASTM): continue
        rr=qrow
        Q2.cell(rr,1,yv); Q2.cell(rr,2,m)
        Q2.cell(rr,3,f"{str(yv)[2:]}.{m:02d}")
        Q2.cell(rr,4,f'=IF(COUNTIFS({RNG("A")},A{rr},{RNG("B")},B{rr})=0,"",SUMIFS({RNG("E")},{RNG("A")},A{rr},{RNG("B")},B{rr}))').number_format=USD4
        Q2.cell(rr,5,f'=AVERAGE($D$2:$D$__END__)').number_format=USD4
        for col in range(1,6): Q2.cell(rr,col).font=f_fx; Q2.cell(rr,col).border=bd
        qrow+=1
for rr in range(2,qrow): Q2.cell(rr,5).value=Q2.cell(rr,5).value.replace('__END__',str(qrow-1))
qc=BarChart(); qc.type='col'; qc.grouping='clustered'; qc.overlap=100; qc.gapWidth=40; qc.height=9; qc.width=30
qc.add_data(Reference(Q2,min_col=4,min_row=1,max_row=qrow-1),titles_from_data=True)
qc.set_categories(Reference(Q2,min_col=3,min_row=2,max_row=qrow-1))
solid(qc.series[0],'C98A3A')
qc.y_axis.numFmt='$0.00'; qc.legend.position='b'; qc.y_axis.delete=False; qc.x_axis.delete=False
qc.x_axis.tickLblSkip=6
qc+=avg_line(Reference(Q2,min_col=5,min_row=1,max_row=qrow-1),Reference(Q2,min_col=3,min_row=2,max_row=qrow-1))
qc.title='JEPQ 월 분배 (USD, 지급월 기준) · 점선 전체 평균'
S.add_chart(qc,'P2')

# ---------------- 구성종목 ----------------
from openpyxl.chart import DoughnutChart
from openpyxl.chart.series import DataPoint
H=wb.create_sheet('구성종목')
KN=META['names_ko']; SKD=META['sectors_ko']
ELN_K=SKD.get('ELN','ELN(옵션 프리미엄)'); CASH_K=SKD.get('','현금·단기자금')
def sec_ko(i):
    if i['type']=='eln': return ELN_K
    if i['type'] in ('cash','other') or not i.get('sector'): return CASH_K
    return SKD.get(i['sector'],i['sector'])
hold=HOLDJ['items']
SK=[(e,k) for e,k in SKD.items() if e not in ('','ELN')]
H['A1']='JEPQ 구성종목'; H['A1'].font=f_t
H['A2']='기준일'; H['A2'].font=f_b
H['B2']=dt.date.fromisoformat(HOLDJ['asOf']); H['B2'].font=f_in; H['B2'].fill=key_fill; H['B2'].number_format=DATE
H['C2']='← JPMorgan 공식 일별 보유종목 엑셀 기준일. 매일 자동으로 최신 파일을 받습니다. 비중 = 시장가치 ÷ 전체 시장가치(합계 100%)'; H['C2'].font=f_s
hs=4
for j,tl in enumerate(['티커','비중','종목명(영문)','구분','섹터(GICS)','한글명','섹터(한글)'],8):
    c=H.cell(hs,j,tl); c.font=f_hd; c.fill=hd_fill; c.alignment=C
for j,w in zip(range(8,15),[9,9,30,8,22,16,16]): H.column_dimensions[L(j)].width=w
TYK={'stock':'주식','eln':'ELN','cash':'현금','other':'기타'}
HLN=hs+len(hold)
for i,it in enumerate(hold):
    rr=hs+1+i
    for col,v,fmt in [(8,it['symbol'],None),(9,it['weight']/100,'0.00%'),(10,it['name'],None),(11,TYK.get(it['type'],it['type']),None),(12,it.get('sector',''),None),(13,KN.get(it['symbol'],''),None),(14,sec_ko(it),None)]:
        c=H.cell(rr,col,v); c.font=f_in; c.fill=in_fill
        if fmt: c.number_format=fmt
HR=lambda col: f'${col}${hs+1}:${col}${HLN}'
# TOP10 주식 + ELN 합계 + 기타 (A~E)
for j,tl in enumerate(['순위','티커','종목','섹터','비중'],1):
    c=H.cell(hs,j,tl); c.font=f_hd; c.fill=hd_fill; c.alignment=C
for j,w in zip(range(1,6),[6,8,22,16,10]): H.column_dimensions[L(j)].width=w
top=[i for i in hold if i['type']=='stock'][:10]
for k,it in enumerate(top,1):
    rr=hs+k
    H.cell(rr,1,k).alignment=Alignment(horizontal='center')
    H.cell(rr,2,it['symbol']); H.cell(rr,3,KN.get(it['symbol']) or it['name'].title()); H.cell(rr,4,sec_ko(it))
    H.cell(rr,5,f'=INDEX({HR("I")},MATCH(B{rr},{HR("H")},0))').number_format='0.00%'
    for col in range(1,6): H.cell(rr,col).font=f_fx; H.cell(rr,col).border=bd
r11=hs+11
H.cell(r11,3,'ELN 합계 (옵션 프리미엄)').font=f_b
H.cell(r11,5,f'=SUMIFS({HR("I")},{HR("K")},"ELN")').number_format='0.00%'
H.cell(r11+1,3,'기타 (나머지 주식·현금)').font=f_n
H.cell(r11+1,5,f'=SUM({HR("I")})-SUM(E{hs+1}:E{r11})').number_format='0.00%'
H.cell(r11+2,3,'합계').font=f_b
H.cell(r11+2,5,f'=SUM(E{hs+1}:E{r11+1})').number_format='0.00%'; H.cell(r11+2,5).font=f_b
H.cell(r11+3,3,'TOP10 주식 합계').font=f_n
H.cell(r11+3,5,f'=SUM(E{hs+1}:E{hs+10})').number_format='0.00%'
H.cell(r11+4,3,'보유 주식 수 / ELN 수').font=f_n
H.cell(r11+4,5,f'=COUNTIFS({HR("K")},"주식")&" / "&COUNTIFS({HR("K")},"ELN")')
sr0=r11+7
H.cell(sr0-1,1,'섹터 비중 (ELN·현금 포함, 합계 100%)').font=f_b
for j,tl in enumerate(['','','섹터','','비중'],1):
    if tl: c=H.cell(sr0,j,tl); c.font=f_hd; c.fill=hd_fill; c.alignment=C
seclist=[k for _,k in SK]+[ELN_K,CASH_K]
for i,k in enumerate(seclist):
    rr=sr0+1+i
    H.cell(rr,3,k).font=f_n
    H.cell(rr,5,f'=SUMIFS({HR("I")},{HR("N")},C{rr})').number_format='0.00%'
    H.cell(rr,5).font=f_fx; H.cell(rr,3).border=bd; H.cell(rr,5).border=bd
H.cell(sr0+len(seclist)+2,1,'출처: J.P. Morgan Asset Management, JEPQ Daily Holdings. ELN = 나스닥100 콜옵션 매도 수익을 담은 주가연계채권. 종목 정보는 투자 권유가 아닙니다.').font=f_s
dc=DoughnutChart(); dc.title='JEPQ TOP10 주식 + ELN + 기타 (합계 100%)'; dc.holeSize=55; dc.height=10; dc.width=14
dc.add_data(Reference(H,min_col=5,min_row=hs,max_row=r11+1),titles_from_data=True)
dc.set_categories(Reference(H,min_col=3,min_row=hs+1,max_row=r11+1))
PAL=['A8661B','C98A3A','E0B36A','2F5D8A','4F7FAF','86A9CC','17784A','4E9E72','8CC3A2','7A5C8E','D9534F','C9D1D6']
ser=dc.series[0]
for idx,col in enumerate(PAL):
    pt=DataPoint(idx=idx); pt.graphicalProperties=GraphicalProperties(solidFill=col); ser.dPt.append(pt)
ser.dLbls=DataLabelList(); ser.dLbls.showVal=True; ser.dLbls.numFmt='0.0%'
for k in ('showSerName','showCatName','showLegendKey','showPercent'): setattr(ser.dLbls,k,False)
dc.legend.position='r'
H.add_chart(dc,'Q4')

# ---------------- 주가_월별 ----------------
_m={}
for d_,c_,a_ in PR['rows']: _m[d_[:7]]=(d_,c_,a_)
P=wb.create_sheet('주가_월별')
header(P,1,['월','월말 거래일','종가(USD)','분배 재투자 지수','종가 전년 대비','재투자 전년 대비'],[9,12,11,14,12,12])
P.freeze_panes='A2'
mrows=sorted(_m.items())
for i,(ym,(d_,c_,a_)) in enumerate(mrows):
    rr=i+2
    P.cell(rr,1,ym); P.cell(rr,2,dt.date.fromisoformat(d_)).number_format=DATE
    c=P.cell(rr,3,c_); c.number_format='$0.00'; c.font=f_in
    c=P.cell(rr,4,a_); c.number_format='0.0000'; c.font=f_in
    if rr>13:
        P.cell(rr,5,f'=C{rr}/C{rr-12}-1').number_format=PCT
        P.cell(rr,6,f'=D{rr}/D{rr-12}-1').number_format=PCT
    for col in range(1,7): P.cell(rr,col).border=bd
P['H1']=f"출처: {PR['source']} · 기준일 {PR['asOf']}"; P['H1'].font=f_s
pc=LineChart(); pc.title='JEPQ 월말 종가 (USD)'; pc.height=9; pc.width=26
pc.add_data(Reference(P,min_col=3,min_row=1,max_row=len(mrows)+1),titles_from_data=True)
pc.set_categories(Reference(P,min_col=1,min_row=2,max_row=len(mrows)+1))
pc.series[0].graphicalProperties=GraphicalProperties(); pc.series[0].graphicalProperties.line=LineProperties(solidFill='16232C',w=19000)
pc.series[0].marker.symbol='none'; pc.legend=None; pc.y_axis.numFmt='$0'; pc.x_axis.tickLblSkip=6
pc.y_axis.delete=False; pc.x_axis.delete=False
P.add_chart(pc,'H3')

# ---------------- 월별 비중 (매월 첫 기록) ----------------
HMJ=J('holdings_monthly'); MSX=HMJ['months']
M=wb.create_sheet('월별_비중')
M['A1']='JEPQ 월별 구성 비중 (매월 처음 받은 공식 보유종목)'; M['A1'].font=f_t
M['A2']=(f"{MSX[0]['month']} ~ {MSX[-1]['month']} · JPMorgan 공식 일별 보유종목 · 빈칸 = 그 달엔 보유하지 않음 · JPMorgan이 과거 날짜 파일을 주지 않아 {MSX[0]['month']}부터 쌓습니다") if MSX else '아직 기록 없음'; M['A2'].font=f_s
hr=4
for j,(tl,w) in enumerate([('티커',8),('종목',24),('섹터',14)],1):
    c=M.cell(hr,j,tl); c.font=f_hd; c.fill=hd_fill; c.alignment=C; M.column_dimensions[L(j)].width=w
for k,m in enumerate(MSX):
    c=M.cell(hr,4+k,f"{m['month']}\n{m['asOf'][5:]}"); c.font=f_hd; c.fill=hd_fill; c.alignment=C
    M.column_dimensions[L(4+k)].width=9
M.row_dimensions[hr].height=30
if MSX:
    # ELN·현금 합계 줄 먼저
    for i,(lab,key) in enumerate([('ELN 합계 (옵션 프리미엄)','eln'),('현금·단기자금','cash')]):
        rr=hr+1+i
        M.cell(rr,1,key.upper()).font=f_b; M.cell(rr,2,lab).font=f_b
        for kk,m in enumerate(MSX):
            c=M.cell(rr,4+kk,(m.get(key) or 0)/100); c.number_format='0.00%'; c.font=f_in
        M.cell(rr,4+len(MSX)+1,f"ELN 개수 {MSX[-1].get('elnCount','')}" if key=='eln' else '').font=f_s
    W={m['month']:{x[0]:x[1] for x in m['items']} for m in MSX}
    lastw=W[MSX[-1]['month']]
    syms=sorted(HMJ['names'],key=lambda k:(-lastw.get(k,-1),-max(W[m['month']].get(k,0) for m in MSX)))
    for i,k in enumerate(syms):
        rr=hr+3+i; nm,sec=HMJ['names'][k]
        M.cell(rr,1,k).font=f_n; M.cell(rr,2,KN.get(k) or nm.title()).font=f_n; M.cell(rr,3,SKD.get(sec,sec)).font=f_n
        for kk,m in enumerate(MSX):
            v=W[m['month']].get(k)
            if v is not None:
                c=M.cell(rr,4+kk,v/100); c.number_format='0.00%'; c.font=f_in
        if i%2:
            for col in range(1,4+len(MSX)): M.cell(rr,col).fill=band
M.freeze_panes=M.cell(hr+1,4)
S2=wb.create_sheet('월별_섹터')
S2['A1']='JEPQ 월별 섹터·ELN 비중 (매월 처음 받은 공식 보유종목)'; S2['A1'].font=f_t
secs=[k for _,k in SK]
header(S2,3,['월','기준일','방식']+secs+[ELN_K,'현금'],[9,12,9]+[11]*(len(secs)+2))
for i,m in enumerate(MSX):
    rr=4+i; agg={}
    for kk,w,*_ in m['items']:
        sk_=SKD.get(HMJ['names'][kk][1],HMJ['names'][kk][1]); agg[sk_]=agg.get(sk_,0)+w
    S2.cell(rr,1,m['month']); S2.cell(rr,2,dt.date.fromisoformat(m['asOf'])).number_format=DATE; S2.cell(rr,3,m['basis']).font=f_s
    for j,sname in enumerate(secs):
        if agg.get(sname): c=S2.cell(rr,4+j,agg[sname]/100); c.number_format='0.00%'; c.font=f_in
    c=S2.cell(rr,4+len(secs),m.get('eln',0)/100); c.number_format='0.00%'; c.font=f_in
    c=S2.cell(rr,5+len(secs),m.get('cash',0)/100); c.number_format='0.00%'; c.font=f_in
    for col in range(1,6+len(secs)): S2.cell(rr,col).border=bd
S2.freeze_panes='D4'

# ---------------- 분배_일정 ----------------
SC=wb.create_sheet('분배_일정')
SC['A1']='JEPQ 분배 일정 (JPMorgan Dividend Calendar)'; SC['A1'].font=f_t
SC['A2']=f"출처: {(META.get('schedule') or {}).get('source','')} · 확인 {(META.get('schedule') or {}).get('checked','')} · 배당락일 = 기준일. 일정은 운용사 사정으로 바뀔 수 있습니다."; SC['A2'].font=f_s
header(SC,4,['구분','배당락일·기준일','지급일','지급월','반영 여부'],[16,14,12,9,14])
have={d['ex'] for d in rows}
for i,x in enumerate(SCH):
    rr=5+i
    vals=['자본이득(있을 때만)' if x.get('kind')=='capgain' else '정기 분배',dt.date.fromisoformat(x['ex']),dt.date.fromisoformat(x['pay']),x['pay'][:7],
          '반영됨' if x['ex'] in have else ('지남·미분배' if x['ex']<=LAST['ex'] else '예정')]
    for j,v in enumerate(vals,1):
        c=SC.cell(rr,j,v); c.font=f_n; c.border=bd
        if j in (2,3): c.number_format=DATE

wb._sheets=[wb[n] for n in ['사용법','최신월_요약','JEPQ_배당','연도별','월별표','국내JEPQ형_분배','국내JEPQ형_비교','구성종목','월별_비중','월별_섹터','분배_일정','주가_월별','차트데이터']]
wb.calculation.fullCalcOnLoad=True
out=ROOT/'downloads'/'jepq_dividend_tracker.xlsx'
wb.save(out); print(out)
