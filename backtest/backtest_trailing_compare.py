"""
SOXL 그리드 매도 전략 2주 백테스트 (정밀版)
- 고정 TP: +0.7% 지정가 (봉 high가 목표가 터치 시 정확히 0.7%에 체결)
- 트레일: 봇과 동일 — 3초 폴링, 시장가, peak_pct 엑셀 저장(포지션 상태 유지)
- 그리드 매수: -0.7% / 50 USDT / 20차
"""
import ccxt
import pandas as pd
from datetime import datetime, timedelta, timezone
from dataclasses import dataclass, field
from typing import List, Optional, Dict, Literal

SYMBOL = "SOXL/USDT:USDT"
DAYS = 14
POLL_SEC = 3
BUY_GAP = -0.7
SELL_GAP = 0.7
TRAIL_ACTIVATE = 0.7
MAX_LEVEL = 20
BUY_USDT = 50.0
FEE = 0.0008

StrategyKind = Literal["limit", "trail", "market"]


@dataclass
class Position:
    level: int
    buy_price: float
    buy_usdt: float
    peak_pct: float = 0.0


@dataclass
class ClosedTrade:
    level: int
    buy_price: float
    sell_price: float
    profit_pct: float
    profit_usdt: float


@dataclass
class SimState:
    name: str
    kind: StrategyKind
    trail_distance: float = 0.0
    positions: List[Position] = field(default_factory=list)
    closed: List[ClosedTrade] = field(default_factory=list)
    next_level: int = 1


def pct(price: float, base: float) -> float:
    return (price / base - 1.0) * 100.0


def bar_poll_prices(o: float, h: float, l: float, c: float, n: int = 20) -> List[float]:
    """1분봉을 3초×20회 폴링 — O→H→L→C 경로 보간"""
    if n <= 1:
        return [c]
    seg = max(n // 3, 1)
    prices: List[float] = []
    for i in range(n):
        if i < seg:
            t = i / max(seg - 1, 1)
            prices.append(o + (h - o) * t)
        elif i < 2 * seg:
            t = (i - seg) / max(seg - 1, 1)
            prices.append(h + (l - h) * t)
        else:
            t = (i - 2 * seg) / max(n - 2 * seg - 1, 1)
            prices.append(l + (c - l) * t)
    return prices


def try_sell_limit(pos: Position, high: float) -> Optional[float]:
    """+0.7% 지정가: 봉 high가 목표 터치 시 정확히 목표가 체결"""
    target = pos.buy_price * (1 + SELL_GAP / 100.0)
    if high >= target:
        return target
    return None


def try_sell_trailing_bot(pos: Position, poll_price: float, distance: float) -> Optional[float]:
    """봇 매도실행과 동일: peak 갱신 → 활성 → 시장가(=poll_price)"""
    cur_pct = pct(poll_price, pos.buy_price)
    pos.peak_pct = max(pos.peak_pct, cur_pct)
    if pos.peak_pct < TRAIL_ACTIVATE:
        return None
    trail_line = pos.peak_pct - distance
    if cur_pct <= trail_line:
        return poll_price
    return None


def try_sell_market_fixed(pos: Position, poll_price: float) -> Optional[float]:
    """TRAILING_ENABLE=False 봇: 현재가 >= 목표가 시 시장가"""
    target = pos.buy_price * (1 + SELL_GAP / 100.0)
    if poll_price >= target:
        return poll_price
    return None


def record_close(state: SimState, pos: Position, sell_px: float):
    gross = sell_px / pos.buy_price
    net = gross - 1 - FEE * 2
    state.closed.append(
        ClosedTrade(pos.level, pos.buy_price, sell_px, net * 100, pos.buy_usdt * net)
    )


def process_sells_limit(state: SimState, high: float):
    remaining = []
    for pos in state.positions:
        sell_px = try_sell_limit(pos, high)
        if sell_px is not None:
            record_close(state, pos, sell_px)
        else:
            remaining.append(pos)
    state.positions = remaining


def process_sells_polled(state: SimState, poll_prices: List[float]):
    for price in poll_prices:
        remaining = []
        for pos in state.positions:
            if state.kind == "trail":
                sell_px = try_sell_trailing_bot(pos, price, state.trail_distance)
            else:
                sell_px = try_sell_market_fixed(pos, price)
            if sell_px is not None:
                record_close(state, pos, sell_px)
            else:
                remaining.append(pos)
        state.positions = remaining


def try_buy(state: SimState, poll_price: float, low: float):
    open_levels = {p.level for p in state.positions}
    for level in range(1, MAX_LEVEL + 1):
        if level in open_levels:
            continue
        if level > state.next_level:
            break
        if level == 1:
            if state.positions:
                continue
            state.positions.append(Position(1, poll_price, BUY_USDT))
            state.next_level = 2
            return
        prev = [p for p in state.positions if p.level == level - 1]
        if not prev:
            continue
        target = prev[0].buy_price * (1 + BUY_GAP / 100.0)
        if low <= target:
            state.positions.append(Position(level, target, BUY_USDT))
            state.next_level = level + 1
            return


def process_bar(state: SimState, o: float, h: float, l: float, c: float):
    polls = bar_poll_prices(o, h, l, c, n=max(1, 60 // POLL_SEC))

    if state.kind == "limit":
        process_sells_limit(state, h)
    else:
        process_sells_polled(state, polls)

    try_buy(state, polls[-1], l)


def summarize(state: SimState) -> Dict:
    df = pd.DataFrame([t.__dict__ for t in state.closed])
    if df.empty:
        return {
            "strategy": state.name,
            "closed_trades": 0,
            "total_profit_usdt": 0.0,
            "avg_profit_pct": 0.0,
            "median_profit_pct": 0.0,
            "win_rate": 0.0,
            "below_054_net": 0.0,
            "open_positions": len(state.positions),
        }
    net_target = SELL_GAP - FEE * 200
    return {
        "strategy": state.name,
        "closed_trades": len(df),
        "total_profit_usdt": round(df["profit_usdt"].sum(), 4),
        "avg_profit_pct": round(df["profit_pct"].mean(), 4),
        "median_profit_pct": round(df["profit_pct"].median(), 4),
        "win_rate": round((df["profit_pct"] > 0).mean() * 100, 1),
        "below_054_net": round((df["profit_pct"] < net_target).mean() * 100, 1),
        "open_positions": len(state.positions),
    }


def fetch_ohlcv() -> pd.DataFrame:
    ex = ccxt.bitget({"enableRateLimit": True, "options": {"defaultType": "swap"}})
    ex.load_markets()
    since = int((datetime.now(timezone.utc) - timedelta(days=DAYS)).timestamp() * 1000)
    all_rows = []
    while True:
        batch = ex.fetch_ohlcv(SYMBOL, "1m", since=since, limit=1000)
        if not batch:
            break
        all_rows.extend(batch)
        since = batch[-1][0] + 60_000
        if len(batch) < 1000:
            break
    df = pd.DataFrame(all_rows, columns=["ts", "open", "high", "low", "close", "volume"])
    df["datetime"] = pd.to_datetime(df["ts"], unit="ms", utc=True)
    return df.drop_duplicates("ts").sort_values("ts")


def main():
    print(f"Fetching {DAYS} days of 1m data for {SYMBOL}...")
    df = fetch_ohlcv()
    print(f"  bars: {len(df)} | {df['datetime'].iloc[0]} ~ {df['datetime'].iloc[-1]}")

    states = [
        SimState("1. 고정 지정가 +0.7%", kind="limit"),
        SimState("2. 고정 시장가 (봇 False)", kind="market"),
        SimState("3. 트레일 0.5% (봇 동일)", kind="trail", trail_distance=0.5),
        SimState("4. 트레일 0% (봇 동일)", kind="trail", trail_distance=0.0),
    ]

    for _, row in df.iterrows():
        for st in states:
            process_bar(st, row["open"], row["high"], row["low"], row["close"])

    print("\n=== 2주 백테스트 (정밀) ===")
    print(f"기간: {df['datetime'].iloc[0].strftime('%Y-%m-%d')} ~ {df['datetime'].iloc[-1].strftime('%Y-%m-%d')}")
    print(f"고정 지정가: high >= 목표 시 +0.7% 정확 체결")
    print(f"트레일/시장가: {POLL_SEC}초 폴링, O-H-L-C 경로, peak 저장, 시장가 체결")
    print(f"그리드: -0.7% / {BUY_USDT} USDT / {MAX_LEVEL}차 / 수수료 0.08%x2\n")

    results = [summarize(st) for st in states]
    print(pd.DataFrame(results).to_string(index=False))

    best = max(results, key=lambda x: x["total_profit_usdt"])
    print(f"\n총 실현손익 1위: {best['strategy']} ({best['total_profit_usdt']} USDT)")

    for st in states:
        if st.closed:
            profits = [t.profit_pct for t in st.closed]
            print(
                f"\n[{st.name}] 청산 {len(st.closed)} | 미청산 {len(st.positions)} | "
                f"최대 {max(profits):+.2f}% | 최소 {min(profits):+.2f}% | "
                f"중앙값 {sorted(profits)[len(profits)//2]:+.2f}%"
            )


if __name__ == "__main__":
    main()
