"""JEPQ 배당 트래커 데이터 자동 갱신.

GitHub Actions가 매일 실행한다. 새 데이터가 있을 때만 data/*.json 을 고치고,
결과를 .run/ 폴더에 남겨 워크플로가 커밋·이슈 생성 여부를 판단하게 한다.

  .run/changed         데이터가 바뀌었으면 생성 (엑셀 재생성·커밋 트리거)
  .run/new_dividend.md 새 JEPQ 분배가 추가됐으면 생성 (알림 이슈 본문)
  .run/errors.md       수집 실패·검증 실패가 있으면 생성 (경고 이슈 본문)

출처 우선순위
  분배   JPMorgan 상품 데이터(최신 1건, 공식) → Nasdaq.com 분배 이력(전체) → stockanalysis(최근 20건)
  주가   Yahoo(종가·수정주가) → Nasdaq.com 종가 + 분배 재투자 수정주가 직접 계산
  보유   JPMorgan 일별 보유종목 엑셀(공식)
  수익률 JPMorgan 상품 데이터(30일 SEC·12개월 롤링 분배수익률·NAV·순자산)
  일정   JPMorgan 'ETF Dividend Calendar' PDF (연 1회)

로컬 실행:  python scripts/update.py          (전체)
            python scripts/update.py --only dividends
"""
import argparse
import datetime as dt
import io
import json
import re
import sys
import time
from pathlib import Path

import requests

sys.path.insert(0, str(Path(__file__).resolve().parent))
import holdings_lib as HL  # noqa: E402
import tr_compare as TRC  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
RUN = ROOT / ".run"
KST = dt.timezone(dt.timedelta(hours=9))
TODAY = dt.datetime.now(KST).date()

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36")
CUSIP = "46654Q203"
JPM_PAGE = ("https://am.jpmorgan.com/us/en/asset-management/adv/products/"
            "jpmorgan-nasdaq-equity-premium-income-etf-etf-shares-46654q203")
JPM_DATA = ("https://am.jpmorgan.com/FundsMarketingHandler/product-data?cusip=46654Q203"
            "&country=us&role=adv&language=en&userLoggedIn=false&version=6.0")
JPM_HOLD = ("https://am.jpmorgan.com/FundsMarketingHandler/excel?type=dailyETFHoldings"
            "&cusip=46654Q203&country=us&role=adv&fundType=N_ETF")
JPM_CAL = ("https://am.jpmorgan.com/content/dam/jpm-am-aem/americas/us/en/supplemental/"
           "income-distribution-rates/jpmorgan-etfs-{y}-distribution-notice.pdf")
NASDAQ_DIV = "https://api.nasdaq.com/api/quote/JEPQ/dividends?assetclass=etf"
NASDAQ_PX = ("https://api.nasdaq.com/api/quote/JEPQ/historical?assetclass=etf"
             "&fromdate=2022-05-01&limit=9999&todate={d}")
SA_DIV = "https://stockanalysis.com/etf/jepq/dividend/"

# 새 분배가 직전 분배 대비 이 범위를 벗어나면 반영하지 않고 경고만 남긴다.
# 2022-06 ~ 2026-11 월간 변화(53회): 표준편차 17%, 최대 +79%(2022-11, 상장 초기), 2023년 이후 최대 ±30.5%.
# 2023년 이후 최대 변동의 약 1.6배인 ±50% — 소수점 밀림·종목 혼동 같은 자료 오류만 걸러낸다.
MAX_CHANGE = 0.50

_cache: dict = {}

errors: list[str] = []
notes: list[str] = []  # 매일 이슈를 만들 정도는 아닌 참고사항 (새 분배 알림에만 덧붙임)
changed: list[str] = []


def load(name):
    return json.loads((DATA / f"{name}.json").read_text(encoding="utf-8"))


def save(name, obj):
    (DATA / f"{name}.json").write_text(json.dumps(obj, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")


def get(url, *, referer=None, timeout=30, tries=3, accept="text/html,application/json,*/*"):
    """JPMorgan·Nasdaq·다음·FunETF 모두 일반 요청으로 열린다(2026-11-01 실측, curl_cffi 불필요).
    막히면 curl_cffi(크롬 흉내)로 한 번 더 시도한다."""
    h = {"User-Agent": UA, "Accept": accept, "Accept-Language": "en-US,en;q=0.9,ko;q=0.8"}
    if referer:
        h["Referer"] = referer
    last = None
    for i in range(tries):
        try:
            r = requests.get(url, headers=h, timeout=timeout)
            if r.status_code == 200 and r.content:
                return r
            last = f"HTTP {r.status_code}"
            if r.status_code in (401, 403) and i == tries - 1:
                from curl_cffi import requests as cr
                r = cr.get(url, impersonate="chrome", headers={"Referer": referer} if referer else None,
                           timeout=timeout)
                if r.status_code == 200 and r.content:
                    return r
                last = f"HTTP {r.status_code} (curl_cffi)"
        except Exception as e:  # noqa: BLE001  네트워크 오류는 재시도
            last = str(e)
        time.sleep(2 * (i + 1))
    raise RuntimeError(f"{url} → {last}")


def jpm_data():
    """JPMorgan 상품 데이터 JSON. 분배·수익률이 같이 쓰므로 한 번만 받는다."""
    if "jpm" not in _cache:
        j = get(JPM_DATA, referer=JPM_PAGE, accept="application/json").json()
        fd = j.get("fundData") or {}
        if (fd.get("shareClass") or {}).get("cusip") != CUSIP:
            raise RuntimeError("JPMorgan 상품 데이터의 CUSIP이 JEPQ가 아님")
        _cache["jpm"] = fd
    return _cache["jpm"]


def mdy(s):
    return dt.datetime.strptime(s.strip(), "%m/%d/%Y").date().isoformat()


# ---------------------------------------------------------------- JEPQ 분배
def fetch_jpm_latest():
    """JPMorgan 공식 최신 분배 1건(발표되면 배당락 전날부터 보인다)."""
    d = jpm_data()["shareClass"].get("fundShareclassDividend") or {}
    if not (d.get("exDate") and d.get("dividendAmount")):
        raise RuntimeError("JPMorgan 상품 데이터에 최신 분배가 없음")
    return [dict(ex=d["exDate"], record=d.get("recordDate") or d["exDate"], pay=d["payDate"],
                 amount=round(float(d["dividendAmount"]), 6), source="JPMorgan 공식")]


def fetch_nasdaq_distributions():
    """Nasdaq.com 분배 이력 전체(배당락일·기준일·지급일·금액 소수 5자리)."""
    j = get(NASDAQ_DIV, accept="application/json").json()
    rows = []
    for x in ((j.get("data") or {}).get("dividends") or {}).get("rows") or []:
        if (x.get("type") or "").lower() != "cash":
            continue
        rows.append(dict(ex=mdy(x["exOrEffDate"]), record=mdy(x["recordDate"]), pay=mdy(x["paymentDate"]),
                         amount=round(float(x["amount"].replace("$", "").replace(",", "")), 6),
                         source="Nasdaq.com(예비)"))
    if not rows:
        raise RuntimeError("Nasdaq.com 분배 표가 비어 있음")
    return rows


def fetch_stockanalysis_distributions():
    """Nasdaq.com도 막혔을 때의 예비 출처(무료 화면은 최근 20건)."""
    r = get(SA_DIV)
    rows = []
    pat = re.compile(r"(\w{3} \d{1,2}, \d{4})\s*\|\s*\$([\d.]+)\s*\|\s*(\w{3} \d{1,2}, \d{4})\s*\|\s*(\w{3} \d{1,2}, \d{4})")
    plain = re.sub(r"<[^>]+>", "|", r.text)
    plain = re.sub(r"(\|\s*)+", "| ", plain)
    f = lambda s: dt.datetime.strptime(s, "%b %d, %Y").date().isoformat()
    for m in pat.finditer(plain):
        rows.append(dict(ex=f(m.group(1)), record=f(m.group(3)), pay=f(m.group(4)),
                         amount=float(m.group(2)), source="stockanalysis(예비)"))
    if not rows:
        raise RuntimeError("stockanalysis 분배 표를 읽지 못함")
    return rows


def update_dividends():
    divs = load("dividends")
    known = {d["ex"] for d in divs}
    src: dict[str, dict] = {}
    official = None
    try:
        backup = fetch_nasdaq_distributions()
    except Exception as e:  # noqa: BLE001
        notes.append(f"Nasdaq.com 분배 이력 접근 실패({e}) → stockanalysis로 대신 확인")
        try:
            backup = fetch_stockanalysis_distributions()
        except Exception as e2:  # noqa: BLE001
            backup = []
            notes.append(f"stockanalysis 분배 표도 실패: {e2}")
    for r in backup:
        src[r["ex"]] = r
    try:
        official = fetch_jpm_latest()[0]
    except Exception as e:  # noqa: BLE001
        notes.append(f"JPMorgan 공식 최신 분배 접근 실패: {e}")
    if official:
        b = src.get(official["ex"])
        if b and abs(b["amount"] - official["amount"]) > 1e-5:
            errors.append(f"최신 분배 금액이 출처마다 다름: JPMorgan {official['amount']} / {b['source']} {b['amount']} "
                          f"({official['ex']}) — JPMorgan 값으로 반영")
        src[official["ex"]] = official
    if not src:
        errors.append("분배 출처(JPMorgan·Nasdaq.com·stockanalysis)를 모두 읽지 못함")

    added = []
    boot = not divs  # 빈 파일 첫 채우기: 과거 이력은 출처 대조(Nasdaq·Yahoo·stockanalysis)로 확인했으므로 변화율 보류를 쓰지 않는다
    for r in sorted(src.values(), key=lambda x: x["ex"]):
        if r["ex"] in known:
            continue
        if r["pay"] < r["ex"] or (divs and r["ex"] <= divs[-1]["ex"]):
            errors.append(f"날짜 이상으로 보류: {r}")
            continue
        chg = None
        if divs:
            prev = divs[-1]
            chg = r["amount"] / prev["amount"] - 1
            if abs(chg) > MAX_CHANGE and not boot:
                errors.append(f"직전 대비 {chg:+.1%} 변화로 보류 (자본이득·특별분배 또는 자료 오류 여부 확인 필요): {r}")
                continue
            # JEPQ는 1월분을 12월 말에 배당락한다(지급은 1월) → '분배월'은 지급일의 달로 센다
            if r["pay"][:7] == prev["pay"][:7]:
                errors.append(f"같은 달(지급월) 두 번째 분배라 보류 (자본이득·특별분배 여부 확인 필요): {r}")
                continue
        r = {**r, "close": None, "fx": None}
        if not boot:  # 처음 채울 때(빈 파일)는 added 표시를 남기지 않는다
            r["added"] = TODAY.isoformat()
        divs.append(r)
        known.add(r["ex"])
        added.append((r, chg))

    # 기존 행 금액·날짜 정정 감지 — 자동 반영하지 않고 알림만
    for d in divs:
        s = src.get(d["ex"])
        if not s:
            continue
        if abs(s["amount"] - d["amount"]) > 1e-5:
            errors.append(f"{s['source']} 금액이 기존 기록과 다름: {d['ex']} 기존 {d['amount']} / 출처 {s['amount']}")
        if s["pay"] != d["pay"]:
            errors.append(f"{s['source']} 지급일이 기존 기록과 다름: {d['ex']} 기존 {d['pay']} / 출처 {s['pay']}")
        # 예비 출처로 넣은 행을 공식 값으로 확인하면 출처 표기를 올린다
        if s["source"] == "JPMorgan 공식" and d["source"] != s["source"] and abs(s["amount"] - d["amount"]) <= 1e-5:
            d["source"] = s["source"]
            changed.append("dividends")

    save("dividends", divs)
    if added:
        changed.append("dividends")
        if not boot:  # 첫 채우기는 알림 이슈를 만들지 않는다
            lines = ["| 배당락일 | 지급일 | 분배금 | 직전 대비 | 출처 |", "|---|---|---|---|---|"]
            for r, chg in added:
                lines.append(f"| {r['ex']} | {r['pay']} | ${r['amount']:.5f} | "
                             f"{'' if chg is None else f'{chg:+.2%}'} | {r['source']} |")
            (RUN / "new_dividend.md").write_text(
                "JEPQ 새 분배가 트래커에 반영됐습니다.\n\n" + "\n".join(lines) +
                "\n\n배당락일 종가는 배당락일 장 마감 뒤, 지급일 환율은 지급일이 지난 뒤 자동으로 채워집니다.\n" +
                ("".join(f"\n> {n}" for n in notes)), encoding="utf-8")
    return [r for r, _ in added]


def fill_close(divs, prices):
    """배당락일 종가(분배 미반영 실제 종가)를 주가 이력에서 채운다."""
    px = {r[0]: r[1] for r in prices["rows"]}
    for d in divs:
        if d.get("close") is None and d["ex"] in px:
            d["close"] = px[d["ex"]]
            changed.append("close")


def fill_fx(divs):
    need = [d for d in divs if d.get("fx") is None and d["pay"] <= TODAY.isoformat()]
    if not need:
        return
    try:
        rates = {}
        for page in range(1, 20):
            r = get("https://finance.daum.net/api/exchanges/FRX.KRWUSD/days?symbolCode=FRX.KRWUSD"
                    f"&page={page}&perPage=100&pagination=true",
                    referer="https://finance.daum.net/exchanges/FRX.KRWUSD")
            data = r.json().get("data") or []
            for x in data:
                rates[x["date"][:10]] = x["basePrice"]
            if not data or min(rates) <= min(d["pay"] for d in need):
                break
            time.sleep(0.5)
        for d in need:
            day = dt.date.fromisoformat(d["pay"])
            for back in range(7):  # 지급일이 국내 휴일이면 직전 영업일
                k = (day - dt.timedelta(days=back)).isoformat()
                if k in rates:
                    d["fx"] = rates[k]
                    changed.append("fx")
                    break
    except Exception as e:  # noqa: BLE001
        errors.append(f"지급일 환율(다음금융, 하나은행 고시) 수집 실패: {e}")


# ---------------------------------------------------------------- 주가 흐름
def adj_from_dividends(rows, divs):
    """종가 + 분배금으로 배당 재투자 수정주가 계산 (Yahoo Adj Close와 같은 방식).
    배당락일마다 그 전날까지의 값에 (1 - 분배금 ÷ 전날 종가)를 곱한다."""
    amt = {d["ex"]: d["amount"] for d in divs}
    out = [list(r) for r in rows]
    f = 1.0
    for i in range(len(out) - 1, -1, -1):
        out[i][2] = round(out[i][1] * f, 4)
        a = amt.get(out[i][0])
        if a and i > 0:
            f *= 1 - a / out[i - 1][1]
    return out


def fetch_prices_yahoo():
    import yfinance as yf
    px = yf.download("JEPQ", period="max", progress=False, auto_adjust=False)
    c, a = px["Close"], px["Adj Close"]
    if hasattr(c, "columns"):
        c, a = c.iloc[:, 0], a.iloc[:, 0]
    rows = [[i.strftime("%Y-%m-%d"), round(float(cv), 4), round(float(av), 4)]
            for i, cv, av in zip(px.index, c, a) if cv == cv and av == av]
    if len(rows) < 100:
        raise RuntimeError(f"Yahoo 주가가 {len(rows)}행뿐")
    return rows, "Yahoo Finance (종가·Adj Close)"


def fetch_prices_nasdaq(divs):
    j = get(NASDAQ_PX.format(d=TODAY.isoformat()), accept="application/json").json()
    rows = []
    for x in ((j.get("data") or {}).get("tradesTable") or {}).get("rows") or []:
        rows.append([mdy(x["date"]), round(float(x["close"].replace("$", "").replace(",", "")), 4), None])
    rows.sort()
    if len(rows) < 100:
        raise RuntimeError(f"Nasdaq.com 주가가 {len(rows)}행뿐")
    return adj_from_dividends(rows, divs), "Nasdaq.com 종가 + 분배 재투자 수정주가(직접 계산)"


def update_prices(divs, force=False):
    """일별 종가·배당재투자 수정주가 전체 이력. 7일마다(또는 새 분배 반영 시) 통째로 다시 받는다."""
    p = load("prices")
    if p.get("rows") and not force and (TODAY - dt.date.fromisoformat(p["asOf"])).days < 7:
        return p
    rows = src = None
    fails = []
    for fn in (fetch_prices_yahoo, lambda: fetch_prices_nasdaq(divs)):
        try:
            rows, src = fn()
            if p.get("rows") and (len(rows) < len(p["rows"]) - 5 or rows[0][0] != p["rows"][0][0]):
                raise RuntimeError(f"주가 이력이 기존보다 짧거나 시작일이 다름 ({len(rows)}행, 시작 {rows[0][0]})")
            break
        except Exception as e:  # noqa: BLE001
            fails.append(str(e)[:200])
            rows = None
    if not rows:
        notes.append("주가 갱신 실패: " + " / ".join(fails))
        return p
    if fails:
        notes.append("Yahoo 주가 실패 → Nasdaq.com 종가로 대신 갱신: " + fails[0])
    if rows[-1][0] != p.get("asOf") or force or src != p.get("source"):
        p = dict(asOf=rows[-1][0], source=src, rows=rows)
        save("prices", p)
        changed.append("prices")
    return p


# ---------------------------------------------------------------- 구성종목
OFFICIAL_STALE_DAYS = 10  # 공식 파일(매 영업일 갱신)을 이만큼 못 받으면 '확인 필요' 이슈


def update_holdings(force=False):
    h = HL.load("holdings")
    try:
        as_of, rows = HL.parse_jpm_xlsx(get(JPM_HOLD, referer=JPM_PAGE).content)
        if force or as_of > h.get("asOf", ""):
            h = HL.official_holdings(as_of, rows)
            HL.save("holdings", h)
            changed.append("holdings")
    except Exception as e:  # noqa: BLE001
        age = (TODAY - dt.date.fromisoformat(h["asOf"])).days if h.get("asOf") else 999
        msg = f"구성종목 공식 파일 수집 실패: {e} (기준일 {h.get('asOf')}, {age}일 경과)"
        (errors if age >= OFFICIAL_STALE_DAYS else notes).append(msg)
    if h.get("items"):
        monthly = HL.load("holdings_monthly")
        if HL.add_month(monthly, h):
            HL.save_monthly(monthly)
            changed.append("holdings_monthly")


# ---------------------------------------------------------------- 공식 수익률·분배 일정
def update_official(meta, force=False):
    """JPMorgan 상품 데이터의 30일 SEC 수익률·12개월 롤링 분배수익률·분배수익률·NAV·순자산을 기준일과 함께 저장한다.
    매일 바뀌므로 SEC 기준일이 7일 이상 새로워졌을 때(또는 새 분배 반영 시)만 갱신해 매일 커밋을 막는다."""
    o = meta.setdefault("official", {})
    try:
        fd = jpm_data()
        sc = fd["shareClass"]
        sec, dy = sc.get("etfSecYield") or {}, sc.get("dailyYield") or {}
        got = dict(
            sec_yield=dict(value=sec.get("thirtyDaySecYield"), asOf=sec.get("effectiveDate")),
            rolling_12m=dict(value=dy.get("twlvMthSimpleYield"), asOf=dy.get("effectiveDate")),
            dist_yield=dict(value=dy.get("dividendYield"), asOf=dy.get("effectiveDate")),
            nav=dict(value=(sc.get("nav") or {}).get("price"), asOf=(sc.get("nav") or {}).get("date")),
            aum=dict(value=(fd.get("aum") or {}).get("netAsset"), asOf=(fd.get("aum") or {}).get("date")),
            expense=round(float((sc.get("expenses") or {}).get("netExpense")) / 100, 6),
        )
        for k in ("sec_yield", "rolling_12m", "dist_yield"):
            v = got[k]["value"]
            if v is None or not 0.02 < float(v) < 0.40 or not got[k]["asOf"]:
                raise RuntimeError(f"{k} 값이 이상함: {got[k]}")
            got[k]["value"] = round(float(v), 4)
        if not got["nav"]["value"]:
            raise RuntimeError("NAV 없음")
        got["nav"]["value"] = round(float(got["nav"]["value"]), 4)
        got["aum"]["value"] = round(float(got["aum"]["value"] or 0))
        old = (o.get("sec_yield") or {}).get("asOf")
        if (force or not old or (dt.date.fromisoformat(got["sec_yield"]["asOf"]) - dt.date.fromisoformat(old)).days >= 7):
            o.update(got, source="J.P. Morgan Asset Management JEPQ 상품 데이터", url=JPM_PAGE,
                     checked=TODAY.isoformat())
            changed.append("official")
    except Exception as e:  # noqa: BLE001
        notes.append(f"JPMorgan 공식 수익률 갱신 실패: {e} (현재 SEC 기준일 {(o.get('sec_yield') or {}).get('asOf')})")


def parse_calendar_pdf(content, year):
    """JPMorgan 'Dividend Calendar' PDF에서 JEPQ가 속한 월분배 그룹의 배당락·지급일을 읽는다."""
    import pypdf
    text = "\n".join(p.extract_text() or "" for p in pypdf.PdfReader(io.BytesIO(content)).pages)
    blocks = re.split(r"(?=Monthly Distributing)", text)
    blk = next((b for b in blocks if "(JEPQ)" in b), None)
    if not blk:
        raise RuntimeError(f"{year} 분배 일정 PDF에서 JEPQ 그룹을 찾지 못함")
    ex_m = re.search(r"Ex-Date/Record Date\s+([\d/ ]+)", blk)
    pay_m = re.search(r"Pay Date\s+([\d/ ]+)", blk)
    if not (ex_m and pay_m):
        raise RuntimeError(f"{year} 분배 일정 PDF 형식이 바뀜")
    exs = [mdy(x) for x in ex_m.group(1).split()]
    pays = [mdy(x) for x in pay_m.group(1).split()]
    if len(exs) != len(pays) or len(exs) < 6:
        raise RuntimeError(f"{year} 분배 일정 날짜 수가 맞지 않음 ({len(exs)}/{len(pays)})")
    cg = "Capital Gains" in blk.split("Ex-Date")[0]
    rows = []
    for n, (e, p) in enumerate(zip(exs, pays)):
        kind = "capgain" if cg and n == len(exs) - 1 else "income"
        rows.append(dict(ex=e, record=e, pay=p, kind=kind))
    return rows


def update_schedule(meta):
    """올해·내년 분배 일정 PDF를 읽어 meta.schedule.rows 를 채운다(이미 있는 해는 다시 읽지 않음)."""
    sch = meta.setdefault("schedule", dict(rows=[]))
    have_years = {r["ex"][:4] for r in sch.get("rows", [])}
    got_any = False
    for y in (TODAY.year, TODAY.year + 1):
        if str(y) in have_years:
            continue
        url = JPM_CAL.format(y=y)
        try:
            r = requests.get(url, headers={"User-Agent": UA}, timeout=30)
            if r.status_code != 200 or not r.content.startswith(b"%PDF"):
                continue  # 아직 안 나온 해
            rows = [x for x in parse_calendar_pdf(r.content, y) if x["ex"][:4] == str(y)]
            sch["rows"] = sorted([x for x in sch.get("rows", []) if x["ex"][:4] != str(y)] + rows,
                                 key=lambda x: x["ex"])
            sch.setdefault("urls", {})[str(y)] = url
            got_any = True
        except Exception as e:  # noqa: BLE001
            notes.append(f"{y} 분배 일정 PDF 읽기 실패: {e}")
    if got_any:
        sch["source"] = "J.P. Morgan Exchange-Traded Fund Trust Dividend Calendar (PDF)"
        sch["checked"] = TODAY.isoformat()
        changed.append("schedule")


def check_schedule(meta, divs):
    sch = meta.get("schedule") or {}
    last_ex = divs[-1]["ex"] if divs else ""
    if not any(r["ex"] > last_ex and r.get("kind") != "capgain" for r in sch.get("rows", [])):
        notes.append("다음 분배 예정일이 data/meta.json 의 schedule에 없습니다. "
                     "JPMorgan 'ETF Dividend Calendar' 새해 PDF가 나왔는지 확인해 주세요.")


# ---------------------------------------------------------------- 국내 JEPQ형 ETF
def update_kr_etf():
    k = load("kr_etf")
    before = json.dumps(k, ensure_ascii=False)
    for code in k["order"]:
        e = k["etfs"][code]
        try:
            r = get(f"https://www.funetf.co.kr/api/public/product/view/etfdividend?itemId={e['isin']}",
                    referer="https://www.funetf.co.kr/")
            hist = {x[0]: x for x in e["history"]}
            for x in r.json():
                d = x.get("gijunYmd")
                if d and x.get("divAmt") is not None:
                    hist[d] = [d, float(x["divAmt"]), float(x.get("divRt") or 0)]
            e["history"] = [hist[d] for d in sorted(hist)]
        except Exception as ex:  # noqa: BLE001
            errors.append(f"{e['name']} 분배금(FunETF) 수집 실패: {ex}")
        time.sleep(1)
    if json.dumps(k, ensure_ascii=False) != before:
        save("kr_etf", k)
        changed.append("kr_etf")


# ---------------------------------------------------------------- 국내 JEPQ형 실부담비용 (KOFIA)
KOFIA_URL = "https://dis.kofia.or.kr/proframeWeb/XMLSERVICES/"


def fetch_kofia_costs(basis):
    """금융투자협회 전자공시 '펀드별 보수비용비교' (상장지수 펀드 전체). 단위 %. 미공개 기준일은 빈 목록."""
    import xml.etree.ElementTree as ET
    payload = ('<?xml version="1.0" encoding="utf-8"?><message><proframeHeader>'
               '<pfmAppName>FS-DIS2</pfmAppName><pfmSvcName>DISFundFeeCmsSO</pfmSvcName>'
               '<pfmFnName>select</pfmFnName></proframeHeader><systemHeader></systemHeader>'
               f'<DISCondFuncDTO><tmpV30>{basis}</tmpV30><tmpV11></tmpV11><tmpV12>상장지수</tmpV12>'
               '<tmpV3></tmpV3><tmpV5></tmpV5><tmpV4></tmpV4></DISCondFuncDTO></message>')
    last = None
    for i in range(4):  # 응답이 1.6MB라 중간에 끊기는 일이 잦다 → 재시도
        try:
            r = requests.post(KOFIA_URL, data=payload.encode("utf-8"), timeout=120,
                              headers={"Content-Type": "application/xml; charset=UTF-8", "User-Agent": UA})
            root = ET.fromstring(r.content.decode("utf-8"))
            break
        except Exception as e:  # noqa: BLE001
            last = e
            time.sleep(3 * (i + 1))
    else:
        raise RuntimeError(f"KOFIA 조회 실패: {last}")
    num = lambda x: float(x) if x not in (None, "") else 0.0
    out = {}
    for m in root.findall(".//selectMeta"):
        g = lambda t: ((m.find(t).text or "").strip() if m.find(t) is not None else "")
        if not g("tmpV2"):
            continue
        ter, sales, trade = num(g("tmpV12")), num(g("tmpV13")) + num(g("tmpV14")), num(g("tmpV16"))
        out[g("tmpV15")] = dict(total_fee=num(g("tmpV9")), other=num(g("tmpV11")), ter=ter, trading=trade,
                                sales=sales, real=round(ter + sales + trade, 4))
    return out


def update_real_cost():
    """월말 기준·약 한 달 뒤 공개. 저장된 기준일보다 새 달(지난달 말)이 공개됐으면 실부담비용을 바꾼다."""
    k = load("kr_etf")
    have = min((k["etfs"][c].get("cost") or {}).get("basis", "0000-00-00") for c in k["order"])
    first = TODAY.replace(day=1)
    month_end = first - dt.timedelta(days=1)          # 지난달 말일
    prev_end = month_end.replace(day=1) - dt.timedelta(days=1)  # 지지난달 말일 (지난달 말 기준이 아직 미공개일 때)
    cands = []
    for me in (month_end, prev_end):
        if me.isoformat() <= have:
            break
        d, n = me, 0
        while n < 4:
            if d.weekday() < 5:
                cands.append(d)
                n += 1
            d -= dt.timedelta(days=1)
    if not cands:
        return
    try:
        for b in cands:
            rows = fetch_kofia_costs(b.strftime("%Y%m%d"))
            if len(rows) < 500:  # 미공개(0건) 또는 휴일
                continue
            missing = []
            for c in k["order"]:
                e = k["etfs"][c]
                code = (e.get("cost") or {}).get("kofia_code")
                if code in rows:
                    e["cost"] = dict(basis=b.isoformat(), **rows[code], kofia_code=code)
                    e["fee"] = round(rows[code]["total_fee"] / 100, 6)
                else:
                    missing.append(e["name"])
            if missing:
                errors.append(f"KOFIA {b} 실부담비용에서 못 찾은 종목: {', '.join(missing)} (kofia_code 확인 필요)")
            k["fee_source"] = f"KOFIA 금융투자협회 전자공시 {b.isoformat()}"
            save("kr_etf", k)
            changed.append("kr_etf")
            notes.append(f"국내 JEPQ형 실부담비용을 KOFIA {b} 기준으로 갱신")
            return
    except Exception as e:  # noqa: BLE001
        notes.append(f"KOFIA 실부담비용 조회 실패: {e}")


# ---------------------------------------------------------------- TR 비교 (5종)
def update_tr_compare():
    """JEPI·JEPQ·QQQ·SCHD·SPY 배당 재투자 비교 지수(data/tr_compare.json). scripts/tr_compare.py 참고.
    7일마다 통째로 다시 받는다. 실패하면 기존 값을 두고, 14일 넘게 못 받으면 '확인 필요' 이슈."""
    old = load("tr_compare") if (DATA / "tr_compare.json").exists() else None
    age = (TODAY - dt.date.fromisoformat(old["asOf"])).days if old else 999
    if age < TRC.REFRESH_DAYS:
        return
    try:
        new = TRC.fetch(old)
    except Exception as e:  # noqa: BLE001
        (errors if age >= TRC.STALE_DAYS else notes).append(f"TR 비교(Yahoo 5종) 갱신 실패, 기존 값 유지 ({old and old['asOf']}): {str(e)[:200]}")
        return
    if new != old:
        save("tr_compare", new)
        changed.append("tr_compare")


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--only", choices=["dividends", "prices", "holdings", "kr_etf"])
    ap.add_argument("--force-holdings", action="store_true")
    a = ap.parse_args()
    RUN.mkdir(exist_ok=True)
    for f in RUN.glob("*"):
        f.unlink()

    meta = load("meta")
    divs = load("dividends")
    if a.only in (None, "dividends"):
        added = update_dividends()
        divs = load("dividends")
        update_official(meta, force=bool(added))
        update_schedule(meta)
        check_schedule(meta, divs)
    if a.only in (None, "dividends", "prices"):
        prices = update_prices(divs, force="dividends" in changed)
        update_tr_compare()
        n0 = len(changed)
        fill_close(divs, prices)
        fill_fx(divs)
        if len(changed) > n0:
            save("dividends", divs)
    if a.only in (None, "holdings"):
        update_holdings(force=a.force_holdings)
    if a.only in (None, "kr_etf"):
        update_kr_etf()
        update_real_cost()

    if changed:
        meta["updated"] = TODAY.isoformat()
        save("meta", meta)
        (RUN / "changed").write_text(",".join(sorted(set(changed))), encoding="utf-8")
    if errors:
        (RUN / "errors.md").write_text("자동 갱신 중 확인이 필요한 항목이 있습니다.\n\n" +
                                       "\n".join(f"- {e}" for e in errors) + "\n", encoding="utf-8")
    print("changed:", sorted(set(changed)) or "없음")
    if notes:
        print("notes:", *notes, sep="\n  ")
    print("errors:", *errors, sep="\n  ") if errors else print("errors: 없음")
    return 0


if __name__ == "__main__":
    sys.exit(main())
