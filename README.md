# JEPQ 분배 트래커

JEPQ(JPMorgan Nasdaq Equity Premium Income ETF) 월 분배 이력, 연간 분배·성장률, 원화 환산, 구성종목(ELN 비중 따로), 국내상장 나스닥100 커버드콜 ETF 3종 비교를 한 페이지로 보여 줍니다.
GitHub Actions가 매일 새 데이터를 확인해 자동으로 갱신합니다. 구조·화면은 `schd-tracker`와 같고, 분기 대신 월 단위로 동작합니다.

## 폴더 구조

| 경로 | 내용 |
|---|---|
| `index.html` | 트래커 화면. `data/*.json`을 읽어서 그립니다 |
| `data/dividends.json` | JEPQ 분배 (배당락일·기준일·지급일·금액·배당락일 종가·지급일 환율·출처). 2022-06 첫 분배부터 전체 |
| `data/holdings.json` | 구성종목 전체 목록(주식·ELN·현금). `asOf` = JPMorgan 공시 기준일 |
| `data/holdings_monthly.json` | 매월 첫 공식 파일의 비중 기록 (2026-10부터, 한 달 = 한 줄, ELN 합계 포함) |
| `data/kr_etf.json` | 국내 JEPQ형(나스닥100 커버드콜) ETF 3종 분배 이력, 총보수, 실부담비용 |
| `data/prices.json` | JEPQ 일별 종가·분배 재투자 수정주가 (2022-05-04 상장 이후 전체) |
| `data/meta.json` | 갱신일, 한글 종목명·섹터명, 원천징수율, 보수, 공식 수익률(`official`), 분배 일정(`schedule`) |
| `downloads/jepq_dividend_tracker.xlsx` | 엑셀 마스터 (데이터가 바뀔 때마다 자동 생성) |
| `scripts/update.py` | 데이터 수집·검증 |
| `scripts/holdings_lib.py` | 구성종목 공통 처리 (JPMorgan 보유종목 엑셀 해석·월별 기록) |
| `scripts/build_xlsx.py` | 엑셀 생성 |
| `.github/workflows/update.yml` | 매일 한국시간 08:30 자동 실행 |

## 출처와 접속 방법 (2026-11-01 실측)

| 자료 | 출처 (우선순위 순) | 접속 방법 | 실측 결과 |
|---|---|---|---|
| 최신 분배 1건 | JPMorgan 상품 데이터 `FundsMarketingHandler/product-data?cusip=46654Q203` | 일반 `requests` (브라우저 UA) | 성공. 배당락·기준·지급일·금액(소수 5자리) |
| 과거 분배 전체 | Nasdaq.com `api.nasdaq.com/api/quote/JEPQ/dividends` | 일반 `requests` | 성공. 53건, 소수 5자리, 기준일·지급일 포함 |
| 〃 예비 | stockanalysis `etf/jepq/dividend/` | 일반 `requests` | 성공. 최근 20건만 |
| 〃 교차 확인 | Yahoo Finance | `yfinance` | 53건(소수 3자리 반올림이라 금액 원본으로는 쓰지 않음) |
| JPMorgan 과거 분배 이력 | — | dividend-history 등 4종 404, 엑셀 `type=distributions` 빈 응답 | 못 찾음 → Nasdaq.com 이력 사용 |
| 주가·수정주가 | Yahoo Finance → Nasdaq.com 종가 + 분배 재투자 직접 계산 | `yfinance` / `requests` | 성공 (PC에서는 사용자 폴더 경로의 한글 때문에 Yahoo 인증서 오류 → GitHub 서버에서는 문제 없음) |
| 구성종목 | JPMorgan 일별 보유종목 엑셀 `FundsMarketingHandler/excel?type=dailyETFHoldings` | 일반 `requests` | 성공. 주식 94 + ELN 15 + 단기자금·현금 |
| 공식 수익률 | JPMorgan 상품 데이터 (30일 SEC·12개월 롤링 분배수익률·분배수익률·NAV·순자산·보수) | 일반 `requests` | 성공 |
| 분배 일정 | JPMorgan 'ETF Dividend Calendar' PDF (연 1회) | 일반 `requests` + `pypdf` | 성공. 2026년판, 2027년판은 아직 미공개 |
| 지급일 환율 | 하나은행 매매기준율(다음금융) | 일반 `requests` | 성공. 지급일이 지나야 채워짐 |
| 국내 ETF 분배금 | FunETF 분배 이력 API | 일반 `requests` | 성공 |
| 국내 ETF 실부담비용 | 금융투자협회 전자공시 '펀드별 보수비용비교' | 일반 `requests` (XML 서비스) | 성공. 2026-09-30 기준 |

SCHD와 달리 JPMorgan은 일반 파이썬 요청을 막지 않아 `curl_cffi`가 필요 없습니다(설치 목록에는 예비로 남김).

## 자동 갱신 규칙

- **JEPQ 분배**: Nasdaq.com 이력(막히면 stockanalysis)으로 빠진 달을 채우고, 가장 최근 1건은 JPMorgan 공식 값으로 덮어씁니다. 새 분배가 직전 대비 ±50% 넘게 다르면 반영하지 않고 이슈로 알림
  - ±50% 근거: 2022-06 ~ 2026-11 월간 변화 53회의 표준편차 17%, 2023년 이후 최대 ±30.5%(상장 초기 2022-11 +79% 제외). 그 1.6배로 잡아 소수점 밀림·종목 혼동 같은 자료 오류만 거릅니다
- **분배 월**: 지급일이 속한 달로 셉니다(배당락일은 보통 매월 첫 영업일, 지급은 2~3영업일 뒤)
- **배당락일 종가**: Yahoo Finance (배당 미조정 종가). 배당락일이 지나야 채워짐
- **주가 흐름**: 전체 이력을 7일마다(새 분배 반영 시엔 즉시) 통째로 다시 받음
- **지급일 환율**: 하나은행 매매기준율(다음금융). 지급일이 지나야 채워짐
- **구성종목**: 매일 JPMorgan 공식 보유종목 엑셀을 받아 기준일이 새로우면 반영. ELN(주가연계채권, 나스닥100 콜옵션 매도 수익을 담는 상품)은 따로 묶어 비중을 보여 줌. 공식 파일을 10일 넘게 못 받으면 `확인 필요` 이슈
- **월별 비중 기록**: 그 달 첫 공식 파일을 `holdings_monthly.json`에 추가. JPMorgan은 과거 보유종목 파일을 공개하지 않아 2026-10부터 쌓입니다. JEPQ는 정해진 지수 리밸런싱일이 없어 SCHD의 리밸런싱 기록은 없습니다
- **공식 수익률**: 30일 SEC 수익률·12개월 롤링 분배수익률·분배수익률·NAV를 기준일과 함께 `data/meta.json`의 `official`에 저장. SEC 기준일이 7일 이상 새로워졌을 때만 갱신(매일 커밋 방지). 실패하면 기존 값 유지
- **다음 분배 예정일**: JPMorgan 'ETF Dividend Calendar' PDF에서 JEPQ가 속한 그룹의 배당락·지급일을 읽어 `data/meta.json`의 `schedule`에 넣음. 올해·내년 PDF를 확인해 아직 없는 해만 읽습니다. 12월 자본이득 분배일은 `kind: capgain`으로 구분(자본이득이 없으면 지급되지 않을 수 있음)
- **국내 JEPQ형 분배금**: FunETF 분배 이력
- **국내 JEPQ형 실부담비용**: 금융투자협회 전자공시 '펀드별 보수비용비교'(총보수+기타비용+매매·중개수수료+판매수수료). 월말 기준·약 한 달 뒤 공개라, 지난달 말 기준이 공개되면 자동으로 바꿈(`data/kr_etf.json`의 `cost`)
- 새 분배가 반영되면 `JEPQ 새 분배 반영` 이슈가, 수집 실패가 있으면 `JEPQ 자동 갱신 확인 필요` 이슈가 생깁니다 (GitHub 알림 메일로 받음)

## 국내 비교 대상 (나스닥100 커버드콜)

| 코드 | 이름 | 고른 이유 |
|---|---|---|
| 486290 | TIGER 미국나스닥100타겟데일리커버드콜 | 기초지수 NASDAQ 100 Daily Covered Call Target Premium 15%(TR). 매일 옵션 매도, 연 15% 프리미엄 목표 |
| 494300 | KODEX 미국나스닥100데일리커버드콜OTM | 기초지수 NASDAQ-100 Daily Covered Call 101 Index(TR). 매일 외가격(OTM) 옵션 매도 |
| 441680 | TIGER 미국나스닥100커버드콜(합성) | 기초지수 Cboe Nasdaq-100 BuyWrite V2(TR). 월 단위 옵션 매도로 JEPQ와 가장 오래 비교 가능(2022-09 상장) |

RISE 미국테크100데일리고정커버드콜(491620)은 기초지수가 나스닥100이 아니라(KEDI 테크100) 제외했습니다.

## 국내 ETF 12개월 분배율 계산 (화면·엑셀 공통)

12개월 분배율 = 12개월 분배금 합계 ÷ 기준 가격. 회차별 분배율을 더한 값이 아닙니다.

- 기간: 비교 기준일(최신 분배월 말일) 직전 12개월에 분배 기준일이 든 회차 전부
- 기준 가격: 최근 분배 시점 1좌 가격. 국내 ETF는 `최근 분배금 ÷ 그 회차 분배율`로 역산(FunETF는 가격을 따로 주지 않음), JEPQ는 최근 배당락일 종가
- JEPQ 행은 기준일 이전 마지막 배당락까지 12회라, 위쪽 요약의 TTM(최신 분배 포함)과 기준일 이후 분배만큼 다를 수 있음
- 화면은 `index.html`의 국내 비교표, 엑셀은 `[국내JEPQ형_비교]` 시트(기준 가격·12개월 분배율 열)가 같은 식을 씁니다

## SCHD 트래커와 다른 점

- 분기(연 4회) → 월(연 12회). 표·차트·엑셀 모두 월 기준, 변동 경고 범위 ±40% → ±50%
- 구성종목에 ELN이 있어 주식·ELN·현금을 나눠 보여 줌. 리밸런싱 기록 없음
- 공식 수익률에 12개월 롤링 분배수익률·NAV 추가
- JPMorgan 분배 일정은 PDF에서 자동으로 읽음(SCHD는 손 입력)
- 엑셀 파일 이름 `downloads/jepq_dividend_tracker.xlsx`

## 손으로 고칠 때

- 값은 `data/*.json`에서 고칩니다. 엑셀은 다음 갱신 때 새로 만들어지므로 엑셀에서 고친 내용은 남지 않습니다.
- 분배 일정 PDF 형식이 바뀌어 자동으로 못 읽으면 `data/meta.json`의 `schedule.rows`에 `ex`·`record`·`pay`·`kind`(`income`/`capgain`)를 직접 넣고 `source`·`checked`를 고칩니다.
- 공식 수익률을 손으로 고칠 때는 `official`의 `value`(소수, 11.47% → 0.1147)와 `asOf`를 같이 고칩니다.
- 국내 ETF 총보수가 바뀌면 `data/kr_etf.json`의 `fee`와 `fee_source`를 고칩니다. 새 ETF를 추가하면 상장일 `listed`(YYYY-MM-DD)도 넣습니다.
- 한글 종목명이 없는 종목은 영문명으로 나옵니다. `data/meta.json`의 `names_ko`에 티커별로 추가합니다.
- Actions 탭 → `JEPQ 데이터 자동 갱신` → `Run workflow`로 언제든 수동 실행할 수 있습니다.

## 로컬에서 보기

`index.html`을 더블클릭하면 데이터를 못 읽습니다(브라우저 보안). 폴더에서 `python -m http.server` 실행 후 http://localhost:8000 으로 엽니다.

출처: J.P. Morgan Asset Management, Nasdaq.com, stockanalysis, Yahoo Finance, 다음금융(하나은행 고시 환율), FunETF, 금융투자협회. 특정 종목의 매수·매도를 권하는 자료가 아닙니다.
