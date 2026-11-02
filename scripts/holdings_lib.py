"""JEPQ 구성종목 공통 함수 (표준 라이브러리 + openpyxl).

JPMorgan 일별 보유종목 엑셀(type=dailyETFHoldings) 한 장을 읽어 holdings.json 형태로 바꾸고,
매월 첫 기록을 holdings_monthly.json 에 남긴다.

종목 구분(type)
  stock : 나스닥100 주식 (DOMESTIC COMMON STOCK · ADR · REIT 등 섹터가 있는 줄)
  eln   : 주가연계채권(Equity Linked Notes, NDX_숫자). 나스닥100 콜옵션 매도 수익을 담는 상품
  cash  : 단기자금(MMF)·달러 현금
비중(weight)은 '그 줄 시장가치 ÷ 전체 시장가치 × 100'(합계 100). 공시 '% of Net Assets'는 officialWeight 로 함께 둔다.

과거 분기말(basis="nport")은 SEC EDGAR Form N-PORT(primary_doc.xml)에서 읽는다(parse_nport).
N-PORT 비중 pctVal 은 순자산 대비라 합계가 100이 아니므로, 위와 같은 '시장가치 ÷ 포함한 줄 합계'로 다시 계산한다.
"""
import datetime as dt
import io
import json
import re
from pathlib import Path
from urllib.parse import quote_plus

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
SOURCE = "J.P. Morgan Asset Management JEPQ Daily Holdings (엑셀)"


def is_stock(item):
    return item.get("type") == "stock"


def is_eln(item):
    return item.get("type") == "eln"


def is_cash(item):
    return item.get("type") == "cash"


def _num(v):
    if v in (None, ""):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    return float(str(v).replace(",", "").replace("%", "").strip() or 0)


def parse_jpm_xlsx(content):
    """엑셀 바이트 → (기준일 YYYY-MM-DD, 행 목록). 형식이 바뀌었으면 RuntimeError."""
    import warnings
    import openpyxl
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        ws = openpyxl.load_workbook(io.BytesIO(content), read_only=True).active
        rows = list(ws.iter_rows(values_only=True))
    as_of, head_i = None, None
    for n, r in enumerate(rows[:20]):
        for c in r:
            m = re.search(r"As of Date:\s*(\d{1,2})/(\d{1,2})/(\d{4})", str(c or ""))
            if m:
                as_of = dt.date(int(m.group(3)), int(m.group(1)), int(m.group(2))).isoformat()
        if r and r[0] == "Ticker":
            head_i = n
            break
    if not as_of or head_i is None:
        raise RuntimeError("보유종목 엑셀에서 기준일 또는 머리글(Ticker)을 찾지 못함")
    head = list(rows[head_i])
    need = ["Ticker", "Security Description", "Security Type", "Shares/Par", "Market Value (USD)",
            "Sector", "% of Net Assets"]
    miss = [k for k in need if k not in head]
    if miss:
        raise RuntimeError(f"보유종목 엑셀 머리글이 바뀜: {miss} 없음")
    ix = {k: head.index(k) for k in head if k}
    out = []
    for r in rows[head_i + 1:]:
        if not r or not (r[ix["Ticker"]] or r[ix["Security Description"]]):
            continue
        if r[ix["Market Value (USD)"]] in (None, ""):
            continue
        typ_raw = str(r[ix["Security Type"]] or "")
        desc = str(r[ix["Security Description"]] or "").strip()
        tick = str(r[ix["Ticker"]] or "").strip()
        sector = str(r[ix["Sector"]] or "").strip()
        if re.search(r"Equity Linked", typ_raw, re.I):
            typ = "eln"
        elif re.search(r"MONEY MARKET|CURRENC|CASH", typ_raw, re.I) or desc.upper() == "CASH":
            typ = "cash"
        elif sector:
            typ = "stock"
        else:
            typ = "other"
        out.append(dict(symbol=tick or desc, name=desc, type=typ, secType=typ_raw, sector=sector,
                        industry=str(r[ix.get("Industry", ix["Sector"])] or "").strip(),
                        shares=_num(r[ix["Shares/Par"]]), mv=round(_num(r[ix["Market Value (USD)"]]), 2),
                        officialWeight=_num(r[ix["% of Net Assets"]])))
    if len([x for x in out if x["type"] == "stock"]) < 80:
        raise RuntimeError(f"주식 줄이 {len(out)}개뿐 — 형식 확인 필요")
    return as_of, out


NPORT_NS = "{http://www.sec.gov/edgar/nport}"


def parse_nport(content, series_id):
    """N-PORT XML 바이트 → (기준일 YYYY-MM-DD, 행 목록). 시리즈가 다르거나 형식이 이상하면 RuntimeError.
    구분: assetCat EC → stock, SN(구조화채권) 중 제목에 ELN·linked to 가 있으면 → eln, STIV(단기자금 펀드) → cash, 나머지(권리 등) → other.
    뺀 것: 증권대여 담보 재투자 MMF(같은 금액의 반환 부채가 있고 JPMorgan 일별 파일에도 없음),
           선물 등 파생(DE·derivativeInfo — valUSD가 보유 금액이 아니라 평가손익)."""
    import xml.etree.ElementTree as ET
    n = NPORT_NS
    root = ET.fromstring(content)
    sid = root.findtext(f".//{n}genInfo/{n}seriesId")
    as_of = root.findtext(f".//{n}genInfo/{n}repPdDate") or ""
    if sid != series_id or not re.fullmatch(r"\d{4}-\d{2}-\d{2}", as_of):
        raise RuntimeError(f"N-PORT 시리즈·기준일이 이상함 ({sid}, {as_of})")
    out = []
    for s in root.iter(n + "invstOrSec"):
        cat = s.findtext(n + "assetCat") or ""
        title = (s.findtext(n + "title") or s.findtext(n + "name") or "").strip()
        if cat.startswith("D") or s.find(n + "derivativeInfo") is not None or re.search(r"securities lending", title, re.I):
            continue
        if cat == "EC":
            typ = "stock"
        elif cat == "SN" and re.search(r"\bELN\b|linked to", title, re.I):
            typ = "eln"
        elif cat == "STIV":
            typ = "cash"
        else:
            typ = "other"
        ids = s.find(n + "identifiers")
        isin = next((x.get("value") for x in (ids if ids is not None else []) if x.tag == n + "isin"), "") or ""
        isin = "" if isin == "N/A" else isin
        cusip = s.findtext(n + "cusip") or ""
        out.append(dict(symbol=isin or cusip or title, name=(s.findtext(n + "name") or title).strip(), type=typ,
                        secType=cat, sector="", isin=isin, title=title, shares=_num(s.findtext(n + "balance")),
                        mv=round(_num(s.findtext(n + "valUSD")), 2), officialWeight=_num(s.findtext(n + "pctVal"))))
    if sum(1 for x in out if x["type"] == "stock") < 50 or not any(x["type"] == "eln" for x in out):
        raise RuntimeError(f"N-PORT {as_of} 주식·ELN 줄이 너무 적음 — 형식 확인 필요")
    return as_of, out


def official_holdings(as_of, rows, basis="official", source=SOURCE):
    total = sum(r["mv"] for r in rows)
    items = []
    for r in rows:
        items.append({**r, "weight": round(r["mv"] / total * 100, 6)})
    items.sort(key=lambda x: -x["weight"])
    eln = [i for i in items if is_eln(i)]
    summary = dict(
        stock=round(sum(i["weight"] for i in items if is_stock(i)), 4),
        eln=round(sum(i["weight"] for i in eln), 4),
        cash=round(sum(i["weight"] for i in items if is_cash(i)), 4),
        other=round(sum(i["weight"] for i in items if i["type"] == "other"), 4),
        stockCount=sum(1 for i in items if is_stock(i)), elnCount=len(eln),
        totalMV=round(total, 2))
    return dict(asOf=as_of, basis=basis, source=source,
                official=dict(asOf=as_of, source=source), summary=summary, items=items)


def load(name):
    return json.loads((DATA / f"{name}.json").read_text(encoding="utf-8"))


def save(name, obj):
    (DATA / f"{name}.json").write_text(json.dumps(obj, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")


def save_monthly(monthly):
    """한 달 = 한 줄로 저장 (파일 크기·커밋 차이를 작게)."""
    c = lambda o: json.dumps(o, ensure_ascii=False, separators=(",", ":"))
    head = {k: v for k, v in monthly.items() if k not in ("names", "months")}
    parts = ["{"] + [f"{c(k)}:{c(v)}," for k, v in head.items()]
    parts.append('"names":{')
    names = sorted(monthly["names"].items())
    parts += [f"{c(k)}:{c(v)}" + ("," if n < len(names) - 1 else "") for n, (k, v) in enumerate(names)]
    parts.append('},"months":[')
    ms = monthly["months"]
    parts += [c(m) + ("," if n < len(ms) - 1 else "") for n, m in enumerate(ms)]
    parts.append("]}")
    (DATA / "holdings_monthly.json").write_text("\n".join(parts) + "\n", encoding="utf-8")


def add_month(monthly, h):
    """그 달 기록이 없을 때만 추가(매월 처음 받은 공식 파일, 또는 SEC N-PORT 분기말). 추가했으면 True.
    items 는 주식만 [기호, 비중%, 보유 주식 수]. ELN·현금은 합계 비중만 남긴다.
    N-PORT 에는 섹터가 없으므로 names 는 JPMorgan 공식 값을 덮어쓰지 않는다."""
    month = h["asOf"][:7]
    if any(m["month"] == month for m in monthly["months"]):
        return False
    items = h["items"]
    s = h["summary"]
    rec = dict(month=month, asOf=h["asOf"], basis=h["basis"], eln=s["eln"], elnCount=s["elnCount"],
               cash=s["cash"], **({"accession": h["accession"]} if h.get("accession") else {}),
               items=[[i["symbol"], round(i["weight"], 4), i["shares"]]
                      for i in sorted(items, key=lambda x: -x["weight"]) if is_stock(i)])
    for i in items:
        if is_stock(i) and (h["basis"] == "official" or i["symbol"] not in monthly["names"]):
            monthly["names"][i["symbol"]] = [i["name"], i["sector"]]
    monthly["months"].append(rec)
    monthly["months"].sort(key=lambda m: m["month"])
    return True


# ---------------------------------------------------------------- 섹터 보충 (SEC N-PORT 에는 섹터가 없다)
# 순서: ① JPMorgan 공식 파일에 나온 섹터(holdings_monthly names) 그대로 ② data/sector_map.json ③ 위키백과 S&P 500 표의 GICS 섹터
#       ④ SEC EDGAR 제출 정보의 SIC 코드 → 섹터 대응(상장폐지·인수 종목) ⑤ 못 찾으면 빈 값('섹터 미확인')
# 섹터 이름은 JPMorgan 파일이 쓰는 GICS 11개 이름 그대로(새 이름 없음).
SEC_UA = "ttokjaeTV research ttokjae@gmail.com"  # SEC 규칙: 연락처가 든 User-Agent, 초당 10회 이하
WIKI_SP500 = "https://en.wikipedia.org/w/index.php?title=List_of_S%26P_500_companies&action=raw"
SEC_TICKERS = "https://www.sec.gov/files/company_tickers.json"
SEC_SUBMISSIONS = "https://data.sec.gov/submissions/CIK{c:010d}.json"
SEC_SEARCH = "https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&type=10-K&owner=exclude&count=10&output=atom&company="
GICS = {"Information Technology", "Health Care", "Financials", "Consumer Discretionary", "Communication Services",
        "Industrials", "Consumer Staples", "Energy", "Utilities", "Real Estate", "Materials"}
# SIC 코드 범위 → 섹터 (처음 맞는 줄). 옛 종목 보충용 거친 대응이라 출처에 SIC 번호를 남긴다.
_SIC = [(3411, 3412, "Materials"), (3559, 3559, "Information Technology"), (4922, 4922, "Energy"),  # 금속캔·반도체장비·가스 송유관
        (100, 999, "Consumer Staples"), (1000, 1099, "Materials"), (1200, 1399, "Energy"), (1400, 1499, "Materials"),
        (1500, 1799, "Industrials"), (2000, 2199, "Consumer Staples"), (2200, 2599, "Consumer Discretionary"),
        (2600, 2699, "Materials"), (2700, 2799, "Communication Services"), (2833, 2836, "Health Care"),
        (2840, 2849, "Consumer Staples"), (2800, 2899, "Materials"), (2900, 2999, "Energy"),
        (3000, 3199, "Consumer Discretionary"), (3200, 3399, "Materials"), (3400, 3569, "Industrials"),
        (3570, 3579, "Information Technology"), (3580, 3599, "Industrials"), (3651, 3652, "Consumer Discretionary"),
        (3600, 3659, "Industrials"), (3660, 3699, "Information Technology"), (3711, 3716, "Consumer Discretionary"),
        (3700, 3799, "Industrials"), (3820, 3829, "Information Technology"), (3840, 3851, "Health Care"),
        (3800, 3899, "Industrials"), (3900, 3999, "Consumer Discretionary"), (4000, 4799, "Industrials"),
        (4800, 4899, "Communication Services"), (4950, 4959, "Industrials"), (4900, 4999, "Utilities"),
        (5000, 5199, "Industrials"), (5400, 5499, "Consumer Staples"), (5912, 5912, "Consumer Staples"),
        (5200, 5999, "Consumer Discretionary"), (6500, 6599, "Real Estate"), (6798, 6798, "Real Estate"),
        (6000, 6799, "Financials"), (7000, 7299, "Consumer Discretionary"), (7310, 7319, "Communication Services"),
        (7370, 7379, "Information Technology"), (7300, 7399, "Industrials"), (7800, 7899, "Communication Services"),
        (7900, 7999, "Consumer Discretionary"), (8000, 8099, "Health Care"), (8200, 8299, "Consumer Discretionary"),
        (8731, 8731, "Health Care"), (8700, 8799, "Industrials")]


def sic_sector(sic):
    sic = int(sic)
    return next((s for lo, hi, s in _SIC if lo <= sic <= hi), "")


def load_sector_map():
    p = DATA / "sector_map.json"
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else {"note": "", "map": {}}


def save_sector_map(sm):
    """종목 하나 = 한 줄 (커밋 차이를 작게)."""
    sm["note"] = "N-PORT 종목 섹터 보충표. map = 기호 → [GICS 섹터(JPMorgan 이름), 출처]. 출처: wikipedia-sp500 · sec-sic:SIC번호"
    c = lambda o: json.dumps(o, ensure_ascii=False, separators=(",", ":"))
    rows = ",\n".join(f"{c(k)}:{c(v)}" for k, v in sorted(sm["map"].items()))
    (DATA / "sector_map.json").write_text(f'{{"note":{c(sm["note"])},\n"map":{{\n{rows}\n}}}}\n', encoding="utf-8")


def apply_sectors(monthly, sm):
    """섹터가 빈 names 항목에만 sector_map 값을 넣는다(JPMorgan 공식 섹터는 그대로). 채운 종목 수를 돌려준다."""
    n = 0
    for sym, v in monthly["names"].items():
        if not v[1] and sm["map"].get(sym):
            v[1] = sm["map"][sym][0]
            n += 1
    return n


def parse_wiki_sp500(text):
    """위키백과 S&P 500 구성종목 표(wikitext) → {티커 또는 정리한 회사 이름: GICS 섹터}. 표 모양이 바뀌었으면 RuntimeError.
    이름 키는 옛 ISIN 이 이상한 기호로 잡힌 종목(예: Linde plc)을 이름으로 찾을 때 쓴다."""
    out = {}
    for row in text.split("\n|-"):
        m = re.search(r"Symbol\|([^}|]+)\}\}", row)
        c = re.split(r"\n\|\|?", "\n" + row.strip())
        if m and len(c) > 3 and c[3].strip() in GICS:
            out[m.group(1).strip()] = c[3].strip()
            out.setdefault(_sec_name_query(re.sub(r"\[\[(?:[^\]|]*\|)?|\]\]", "", c[2]))[0].lower(), c[3].strip())
    if len(out) < 400:
        raise RuntimeError(f"위키백과 S&P 500 표에서 {len(out)}종목만 읽음 — 형식 확인 필요")
    return out


def _sec_name_query(name):
    """회사 이름 → EDGAR 검색어: 괄호·구두점·끝의 Inc/Corp 같은 말을 떼고, 없으면 앞 두 단어."""
    w = re.sub(r"\(.*?\)|[^\w ]", " ", name).split()
    while len(w) > 1 and w[-1].lower() in ("inc", "corp", "co", "ltd", "plc", "the", "holdings", "company", "nv", "sa"):
        w.pop()
    return [" ".join(w)] + ([" ".join(w[:2])] if len(w) > 2 else [])


def resolve_sectors(missing, get):
    """missing {기호: 이름} → {기호: [섹터, 출처]}. get(url) 은 requests 응답을 돌려주는 함수(SEC 요청 간격은 get 이 지킨다).
    못 찾은 종목은 결과에 없다. 위키백과 표를 못 읽으면 SIC 로만 찾는다."""
    out, wiki = {}, {}
    try:
        wiki = parse_wiki_sp500(get(WIKI_SP500).text)
    except Exception:  # noqa: BLE001
        pass
    tickers = None
    for sym, name in sorted(missing.items()):
        key = sym if sym in wiki else _sec_name_query(name)[0].lower()
        if key in wiki:
            out[sym] = [wiki[key], "wikipedia-sp500" + ("" if key == sym else "(이름)")]
            continue
        try:
            if tickers is None:
                tickers = {v["ticker"]: v["cik_str"] for v in get(SEC_TICKERS).json().values()}
            cik = tickers.get(sym.replace(".", "-"))
            if not cik:  # 상장폐지·인수 종목: 이름으로 검색 (딱 한 회사로 좁혀진 경우만 — 여러 곳이면 이름을 알 수 없어 건너뜀)
                for q in _sec_name_query(name):
                    t = get(SEC_SEARCH + quote_plus(q)).text
                    if "<conformed-name>" in t:
                        cik = int(re.search(r"<cik>(\d+)</cik>", t).group(1))
                        break
            sic = cik and get(SEC_SUBMISSIONS.format(c=int(cik))).json().get("sic")
            if sic and sic_sector(sic):
                out[sym] = [sic_sector(sic), f"sec-sic:{sic}"]
        except Exception:  # noqa: BLE001
            continue
    return out


def fill_sectors(monthly, get):
    """sector_map 을 먼저 보고, 없는 종목만 resolve_sectors 로 새로 조회해 sector_map 에 저장한 뒤 names 에 적용한다.
    (monthly 는 호출한 쪽이 저장.) → (새로 붙인 종목 수, 못 찾은 {기호: 이름})"""
    sm = load_sector_map()
    apply_sectors(monthly, sm)
    miss = {k: v[0] for k, v in monthly["names"].items() if not v[1]}
    got = resolve_sectors(miss, get) if miss else {}
    if got:
        sm["map"].update(got)
        save_sector_map(sm)
        apply_sectors(monthly, sm)
    return len(got), {k: v for k, v in miss.items() if k not in got}


if __name__ == "__main__":  # 자체 점검: python scripts/holdings_lib.py
    sec = lambda cat, t, v: (f"<invstOrSec><name>{t}</name><title>{t}</title><balance>1</balance>"
                             f"<valUSD>{v}</valUSD><pctVal>{v}</pctVal><assetCat>{cat}</assetCat></invstOrSec>")
    x = (f'<edgarSubmission xmlns="{NPORT_NS[1:-1]}"><formData><genInfo><seriesId>S1</seriesId>'
         "<repPdDate>2026-06-30</repPdDate></genInfo><invstOrSecs>"
         + "".join(sec("EC", f"S{i}", 1) for i in range(50)) + sec("SN", "BARCLAYS ELN", 40)
         + sec("SN", "Note linked to S&amp;P 500", 5) + sec("STIV", "JPM Prime MMF", 5)
         + sec("STIV", "JPM Securities Lending MMF", 99) + sec("DE", "Future", 3) + "</invstOrSecs></formData></edgarSubmission>")
    d, rows = parse_nport(x, "S1")
    s = official_holdings(d, rows, basis="nport")["summary"]
    assert (d, len(rows), s["stock"], s["eln"], s["cash"], s["elnCount"]) == ("2026-06-30", 53, 50.0, 45.0, 5.0, 2), s
    print("parse_nport ok")
    assert (sic_sector(7372), sic_sector(2911), sic_sector(3559), sic_sector(6798), sic_sector(9995)) == (
        "Information Technology", "Energy", "Information Technology", "Real Estate", "")
    wk = "{| id=constituents\n|-\n! h" + "".join(f"\n|-\n|| {{{{NyseSymbol|T{i}}}}}\n|| [[Co {i} Inc.|Co {i}]]\n|| Energy \n|| x" for i in range(400))
    w = parse_wiki_sp500(wk)
    assert w["T7"] == "Energy" and w["co 7"] == "Energy", list(w)[:3]
    mo = {"names": {"A": ["a", "Financials"], "B": ["b", ""], "C": ["c", ""]}}
    assert apply_sectors(mo, {"map": {"A": ["Energy", "x"], "B": ["Materials", "x"]}}) == 1 and mo["names"] == {
        "A": ["a", "Financials"], "B": ["b", "Materials"], "C": ["c", ""]}
    print("sector helpers ok")
