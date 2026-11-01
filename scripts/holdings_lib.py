"""JEPQ 구성종목 공통 함수 (표준 라이브러리 + openpyxl).

JPMorgan 일별 보유종목 엑셀(type=dailyETFHoldings) 한 장을 읽어 holdings.json 형태로 바꾸고,
매월 첫 기록을 holdings_monthly.json 에 남긴다.

종목 구분(type)
  stock : 나스닥100 주식 (DOMESTIC COMMON STOCK · ADR · REIT 등 섹터가 있는 줄)
  eln   : 주가연계채권(Equity Linked Notes, NDX_숫자). 나스닥100 콜옵션 매도 수익을 담는 상품
  cash  : 단기자금(MMF)·달러 현금
비중(weight)은 '그 줄 시장가치 ÷ 전체 시장가치 × 100'(합계 100). 공시 '% of Net Assets'는 officialWeight 로 함께 둔다.
"""
import datetime as dt
import io
import json
import re
from pathlib import Path

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


def official_holdings(as_of, rows):
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
    return dict(asOf=as_of, basis="official", source=SOURCE,
                official=dict(asOf=as_of, source=SOURCE), summary=summary, items=items)


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
    """그 달 기록이 없을 때만 추가(매월 처음 받은 공식 파일). 추가했으면 True.
    items 는 주식만 [기호, 비중%, 보유 주식 수]. ELN·현금은 합계 비중만 남긴다."""
    month = h["asOf"][:7]
    if any(m["month"] == month for m in monthly["months"]):
        return False
    items = h["items"]
    s = h["summary"]
    rec = dict(month=month, asOf=h["asOf"], basis=h["basis"], eln=s["eln"], elnCount=s["elnCount"],
               cash=s["cash"],
               items=[[i["symbol"], round(i["weight"], 4), i["shares"]]
                      for i in sorted(items, key=lambda x: -x["weight"]) if is_stock(i)])
    for i in items:
        if is_stock(i):
            monthly["names"][i["symbol"]] = [i["name"], i["sector"]]
    monthly["months"].append(rec)
    monthly["months"].sort(key=lambda m: m["month"])
    return True
