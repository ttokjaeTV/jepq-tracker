"""JEPI·JEPQ·QQQ·SCHD·SPY 배당 재투자 총수익(TR) 비교 데이터 — data/tr_compare.json

schd-tracker·jepq-tracker·jepi-tracker 세 저장소에 똑같은 파일을 둔다. 고칠 때는 세 곳을 같이 고친다.

  시작일  5종이 모두 거래된 첫날 = 가장 늦게 상장한 JEPQ의 첫 거래일(2022-05-04)
  값      Yahoo Finance Adj Close(배당 재투자 수정주가)를 시작일 = 100 으로 나눈 지수
          새 배당이 생기면 Adj Close 과거 값은 모두 같은 비율로 바뀌므로 지수의 과거 값은 그대로다.
  주기    prices.json 과 같은 7일. 새 배당이 있어도 앞당기지 않는다(세 저장소가 같은 날 같은 값을 갖도록).
"""
import datetime as dt

TICKERS = ["JEPI", "JEPQ", "QQQ", "SCHD", "SPY"]
NAMES = {"JEPI": "S&P500 커버드콜", "JEPQ": "나스닥100 커버드콜", "QQQ": "나스닥100",
         "SCHD": "미국 배당성장", "SPY": "S&P500"}
SOURCE = "Yahoo Finance Adj Close (배당 재투자 반영 수정주가)"
REFRESH_DAYS = 7   # prices.json 전체 재다운로드 주기와 같다
STALE_DAYS = 14    # 이만큼 못 받으면 '확인 필요' 이슈
MAX_DRIFT = 0.02   # 기존 지수와 겹치는 날 값이 이보다 크게 다르면 자료 오류로 보고 반영하지 않는다


def build(adj):
    """adj: {티커: {날짜: Adj Close}} → tr_compare.json 내용. 5종 모두 값이 있는 날만 쓴다."""
    days = sorted(set.intersection(*(set(adj[t]) for t in TICKERS)))
    if len(days) < 100:
        raise RuntimeError(f"5종 공통 거래일이 {len(days)}일뿐")
    base = [adj[t][days[0]] for t in TICKERS]
    rows = [[d, *[round(adj[t][d] / b * 100, 4) for t, b in zip(TICKERS, base)]] for d in days]
    return dict(asOf=days[-1], start=days[0], source=SOURCE, tickers=TICKERS, names=NAMES, rows=rows)


def fetch(old=None):
    import yfinance as yf
    px = yf.download(TICKERS, start="2022-01-01", progress=False, auto_adjust=False, threads=False)
    a = px["Adj Close"]
    adj = {t: {i.strftime("%Y-%m-%d"): float(v) for i, v in a[t].items() if v == v} for t in TICKERS}
    new = build(adj)
    if old:
        if new["start"] != old["start"] or len(new["rows"]) < len(old["rows"]) - 5:
            raise RuntimeError(f"시작일이 다르거나 이력이 짧아짐 ({new['start']}, {len(new['rows'])}행)")
        prev = {r[0]: r for r in old["rows"]}
        for r in new["rows"]:
            o = prev.get(r[0])
            if o and any(abs(x / y - 1) > MAX_DRIFT for x, y in zip(r[1:], o[1:])):
                raise RuntimeError(f"{r[0]} 지수가 기존 값과 {MAX_DRIFT:.0%} 넘게 다름 (배당·분할 자료 오류 확인 필요)")
        # Yahoo Adj Close 는 받을 때마다 소수 끝자리가 조금씩 달라진다 → 기존 날짜는 그대로 두고 새 날짜만 덧붙인다
        new["rows"] = old["rows"] + [r for r in new["rows"] if r[0] > old["asOf"]]
    return new


def period_start(rows, years):
    """기간 버튼과 같은 규칙: 마지막 거래일에서 years년 전 날짜 이후 첫 거래일의 행 번호 (0이면 시작일)."""
    if not years:
        return 0
    last = dt.date.fromisoformat(rows[-1][0])
    try:
        frm = last.replace(year=last.year - years)
    except ValueError:  # 2월 29일 → 3월 1일 (브라우저 Date 와 같다)
        frm = dt.date(last.year - years, 3, 1)
    frm = frm.isoformat()
    return next(i for i, r in enumerate(rows) if r[0] >= frm)


def summary(tr, years):
    """(시작 행, 끝 행, {티커: (누적, 연환산 또는 None)}). 화면 표·엑셀 시트와 같은 계산."""
    rows = tr["rows"]
    i0, i1 = period_start(rows, years), len(rows) - 1
    yrs = (dt.date.fromisoformat(rows[i1][0]) - dt.date.fromisoformat(rows[i0][0])).days / 365.25
    out = {}
    for k, t in enumerate(tr["tickers"], 1):
        c = rows[i1][k] / rows[i0][k] - 1
        out[t] = (c, (1 + c) ** (1 / yrs) - 1 if years != 1 and yrs >= 1 else None)
    return i0, i1, out


if __name__ == "__main__":  # 점검: python scripts/tr_compare.py → 원자료로 받아 요약 출력
    import json
    from pathlib import Path
    p = Path(__file__).resolve().parent.parent / "data" / "tr_compare.json"
    tr = json.loads(p.read_text(encoding="utf-8"))
    assert period_start([["2024-02-28"], ["2024-03-01"], ["2025-02-28"]], 1) == 0
    assert period_start([["2023-02-28"], ["2023-03-01"], ["2024-02-29"]], 1) == 1
    for y in (1, 3, 0):
        i0, i1, s = summary(tr, y)
        print(f"{y or '전체'}년 {tr['rows'][i0][0]}~{tr['rows'][i1][0]}",
              " ".join(f"{t} {c:+.2%}" + (f"(연 {g:+.2%})" if g is not None else "") for t, (c, g) in s.items()))
