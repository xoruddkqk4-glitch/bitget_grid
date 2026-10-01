r"""
================================================================================
 비트겟 그리드롱 자동매매 봇   |   Grid_Bitget.py
================================================================================

[무엇을 하는 봇인가]
--------------------------------------------------------------------------------
Bitget USDT 무기한선물에서 '그리드 롱' 전략을 실행합니다.

  - 1차 매수 후, 가격이 정해진 % 만큼 하락할 때마다 분할 매수 (차수 증가)
  - 각 차수는 자기 매수가 대비 목표 % 상승 시 그 차수만 개별 청산
  - 3초마다 가격을 확인하는 상주 프로그램 (한 번 켜두면 24시간 실행)
  - 매수/매도 발생 시 텔레그램 알림, 모든 거래는 엑셀 파일에 기록

--------------------------------------------------------------------------------
[파일 위치]   ※ 모두 /var/autobot 폴더 안에 있습니다
--------------------------------------------------------------------------------
  봇 파일       /var/autobot/bitget_grid/Grid_Bitget.py
  거래 기록     /var/autobot/bitget_grid/BTC-USDT_비트겟그리드_LONG폴.xlsx   (종목별로 자동 생성)
  에러 로그     /var/autobot/bitget_grid/bitget-grid-error.log
  청산 락 파일  /var/autobot/bitget_grid/.BTC_USDT_USDT_total_profit_target_reached.lock

--------------------------------------------------------------------------------
[시작하기 전에 할 일]
--------------------------------------------------------------------------------
  1) 비트겟 → 우상단 프로필 → API keys 에서 키 발급 (선물 거래 권한 필요)
  2) 아래 'API 키' 항목에 access / secret / password 입력
  3) 선물(Futures) 지갑에 USDT 가 있는지 확인
     ※ 현물 지갑에 있으면 주문이 나가지 않습니다. Transfer 로 옮기세요.
  4) '매매 종목 및 그리드 템플릿' 항목에서 종목과 차수별 금액 설정

--------------------------------------------------------------------------------
[실행 / 확인 / 정지]   ※ SSH 접속 후 아래 명령을 그대로 복사해서 사용
--------------------------------------------------------------------------------
  ● 실행 (백그라운드 - 터미널 창을 닫아도 계속 돌아갑니다)

      nohup python3 -u /var/autobot/bitget_grid/Grid_Bitget.py > /dev/null 2>> /var/autobot/bitget_grid/bitget-grid-error.log &

  ● 실행 중인지 확인 (결과가 한 줄이라도 나오면 실행 중)

      ps aux | grep Grid_Bitget | grep -v grep

  ● 에러 확인 (Ctrl + C 로 빠져나옴)

      tail -f /var/autobot/bitget_grid/bitget-grid-error.log

  ● 정지

      pkill -f Grid_Bitget.py

  ● 다시 실행 → 위 '실행' 명령을 그대로 다시 입력

--------------------------------------------------------------------------------
[그리드가 돌아가는 순서]
--------------------------------------------------------------------------------
  1차 매수 → 가격 -0.7% 하락 → 2차 매수 → 또 -0.7% 하락 → 3차 매수 ...
  각 차수는 자기 매수가에서 +0.7% 오르면 그 차수만 청산됩니다.

  예) 1차를 100 USDT 매수 → 하락하여 2차 100 USDT 매수 (평단가 내려감)
      2차 매수가 +0.7% 도달 → 2차만 청산 → 다시 하락하면 2차를 재매수

  간격(-0.7%, +0.7%)과 금액(100 USDT)은 GRID_TEMPLATE 에서 바꿉니다.

--------------------------------------------------------------------------------
[봇은 엑셀로 기억합니다 - 매우 중요]
--------------------------------------------------------------------------------
  봇은 자기가 무엇을 얼마에 샀는지 오직 엑셀 파일로 기억합니다.

  - 봇이 꺼졌다 켜져도 엑셀만 있으면 이어서 매매합니다.
  - 포지션이 열려 있는 상태에서 엑셀을 지우면 봇이 기존 물량을 잊어버립니다.
    → 거래소에는 포지션이 남아 있는데 봇은 1차부터 새로 매수 (매우 위험)
  - 엑셀은 반드시 '미청산 차수 0개' 일 때만 지우거나 옮기세요.

--------------------------------------------------------------------------------
[총 수익 목표 자동 청산]
--------------------------------------------------------------------------------
  TOTAL_PROFIT_TARGET_ENABLE = True 일 때만 작동합니다.
  (현재 미실현 손익 + 엑셀 누적실현손익) 합계가 목표에 도달하면 전 차수를 청산합니다.

  청산 후 동작은 TOTAL_PROFIT_AFTER_LIQUIDATION 으로 정합니다.

      "STOP"      청산 후 봇 종료          + 락 파일 생성
      "PAUSE"     봇은 살아있고 매수만 정지  + 락 파일 생성
      "RESTART"   즉시 1차부터 다시 시작    + 락 파일 없음

  락 파일이 있으면 봇을 다시 켜도 매수하지 않습니다. 새로 시작하려면
  아래 두 가지를 순서대로 하세요.

      1) 엑셀을 다른 이름으로 옮기기 (누적 수익이 0부터 다시 쌓이도록)
         mv /var/autobot/bitget_grid/BTC-USDT_비트겟그리드_LONG폴.xlsx /var/autobot/bitget_grid/백업_1차.xlsx

      2) 락 파일 삭제
         rm /var/autobot/bitget_grid/.*_total_profit_target_reached.lock

  ※ 엑셀을 옮기지 않고 락만 지우면, 누적 수익이 그대로 남아 있어
     다음 순간 목표 달성으로 판단하고 곧바로 다시 락이 걸립니다.

--------------------------------------------------------------------------------
[문제가 생겼을 때 확인 순서]
--------------------------------------------------------------------------------
  1. 봇이 살아있나?          ps aux | grep Grid_Bitget | grep -v grep
  2. 에러가 있나?            tail -n 50 /var/autobot/bitget_grid/bitget-grid-error.log
  3. 텔레그램 알림이 오나?    10분마다 생존 알림 도착 (끊기면 봇 사망 의심)
  4. 선물 지갑에 USDT 있나?   현물 지갑에 있으면 주문이 나가지 않습니다
  5. 락 파일이 걸려있나?      ls -la /var/autobot/bitget_grid/*.lock

  자주 나오는 상황
  - "이미 실행 중" 메시지    → 봇이 두 개 뜨지 않도록 막은 것. 정상입니다.
  - 매수가 안 들어감         → 락 파일 / 선물 지갑 잔고 / 최소 주문 수량 확인
  - 주문 수량 오류           → 차수별 매수금액이 거래소 최소 주문 금액보다 작은 경우

--------------------------------------------------------------------------------
[주의]
--------------------------------------------------------------------------------
  - 선물·레버리지 거래는 큰 손실 위험을 동반하며 모든 책임은 본인에게 있습니다.
  - 반드시 소액으로 충분히 테스트한 뒤 금액을 늘리세요.
  - API 키는 절대 외부에 공유하지 마세요. 출금 권한은 부여하지 않는 것이 안전합니다.

================================================================================
"""

import sys
import os
import time
import signal
from datetime import datetime
import copy
from typing import Dict, Any, Tuple, Optional

import ccxt
import pandas as pd
from tendo import singleton

import telegram_alert  # 사용자 정의 모듈: SendMessage(text: str) 제공


# ===============================
# 싱글톤: 봇이 두 개 실행되는 것을 막습니다
# ===============================
# 실수로 실행 명령을 두 번 입력해도, 이미 돌고 있으면 새로 뜬 쪽이 바로 종료됩니다.
# 봇이 두 개 돌면 같은 종목을 중복 매매하게 되므로 반드시 필요한 안전장치입니다.
try:
    me = singleton.SingleInstance()
except singleton.SingleInstanceException:
    print("[싱글톤] 이미 실행 중 → 이번 실행은 즉시 종료합니다.")
    sys.exit(0)


# ===============================
# ⚠️ API 키 (실제 값으로 교체)
# ===============================
access = "bg_c5fde21271d5162915719e689bfc5cc9"  # 본인 API 키
secret = "6e6f1784ccb6cada8a39f45183e3f5f0cc5c77bba9f3f1873fbadd459bbee891"  # 본인 Secret 키
password = "0503kmw761"  # Bitget은 패스프레이즈 필요


# ===============================
# 거래소 연결 (Bitget USDT 무기한)
# ===============================
bitgetX = ccxt.bitget({
    "apiKey": access,
    "secret": secret,
    "password": password,
    "enableRateLimit": True,             # ccxt 자체 rate limit 활성화 (안전)
    "options": {"defaultType": "swap"},  # USDT 무기한선물 모드
})

# 마켓 메타데이터 로드 (수량 정밀도, 최소 주문 단위 등)
try:
    bitgetX.load_markets()
except Exception as e:
    print(f"[거래소 메타 로드 실패] {e}")


# ===============================
# 📌 실시간 폴링 설정
# ===============================

# 📌 POLL_INTERVAL = 3
# - 봇이 가격을 조회하고 매수/매도 조건을 체크하는 주기 (초 단위)
# - 3초 권장: Bitget rate limit + EC2 t3.micro 자원 + 다른 봇 공존 고려한 절충값
# - 1초 미만: rate limit 위험 + 다른 봇과 자원 경합
# - 5초 이상: 그리드 봇 의미 약화 (체결 기회 놓침)
POLL_INTERVAL = 3

# 📌 HEARTBEAT_INTERVAL = 600
# - 봇이 살아있다는 신호를 텔레그램으로 보내는 주기 (초 단위)
# - 600초 = 10분마다 1회 발송
# - 이 알림이 끊기면 봇 사망/네트워크 문제로 인지 가능
# - 너무 짧으면 텔레그램 스팸, 너무 길면 사망 인지 지연
HEARTBEAT_INTERVAL = 86400

# 📌 ERROR_BACKOFF_BASE = 5
# - 메인 루프에서 예외 발생 시 첫 번째 재시도 대기 시간 (초)
# - 이후 5 → 10 → 20 → 40 → 60(최대) 지수 증가 (백오프 전략)
ERROR_BACKOFF_BASE = 5

# 📌 ERROR_BACKOFF_MAX = 60
# - 백오프 최대 대기 시간 (초)
# - 연속 에러가 많아도 1분 이상은 대기하지 않음
ERROR_BACKOFF_MAX = 60

# 📌 RATE_LIMIT_BACKOFF = 30
# - Bitget API rate limit에 걸렸을 때 대기 시간 (초)
# - 30초 권장 (대부분의 거래소는 1분 단위로 rate limit 리셋)
RATE_LIMIT_BACKOFF = 30

# --- 콘솔/로그 실시간 모니터링 설정 ---

# 📌 PRICE_LOG_INTERVAL = 10
# - 현재가를 콘솔/로그에 출력하는 주기 (폴링 횟수 단위)
# - 10 = POLL_INTERVAL(3초) × 10 = 30초마다 1줄 출력
# - 1로 두면 매 폴링(3초)마다 출력 → 빠른 모니터링 가능, 단 로그 파일이 빠르게 커짐
# - 20으로 두면 60초마다 1줄 → 로그 절약
# - 출력 예: 📊 [09:23:42] 폴링#20 | Q=0.073180
# - 실시간 확인: $ tail -f /var/autobot/bitget_grid/bitget-grid.log
PRICE_LOG_INTERVAL = 120

# 📌 PRICE_LOG_DECIMALS = 6
# - 가격 출력 시 표시할 소수점 자릿수
# - 저가 알트코인(0.0001 단위) → 6~8 권장
# - BTC/ETH 같은 고가 코인 → 2~4면 충분
PRICE_LOG_DECIMALS = 6


# --- 실제 체결가 조회 설정 ---

# 📌 FILL_FETCH_RETRIES = 3
# - 시장가 주문 후 fetch_order로 평균체결가를 조회할 때 최대 재시도 횟수
# - Bitget는 보통 즉시 체결되지만, 네트워크 지연 등으로 즉시 조회 시 'open'일 수 있음
# - 너무 크면 봇 응답성 저하, 너무 작으면 fallback(현재가) 빈도 증가
FILL_FETCH_RETRIES = 3

# 📌 FILL_FETCH_DELAY = 0.5
# - fetch_order 재시도 간 대기 시간 (초)
# - 0.3~1.0 권장 (너무 짧으면 미체결 상태 반복, 너무 길면 봇 폴링 지연)
FILL_FETCH_DELAY = 0.5


# --- 총 수익 목표 자동 청산 설정 ⭐ (사용자 설정 핵심 영역) ---
# 이 설정으로 "특정 USDT 금액에 도달하면 모든 차수 자동 청산" 기능을 제어합니다.
# 자세한 동작 설명은 파일 상단 docstring의 [🎯 총 수익 목표 자동 청산] 섹션 참고.

# 📌 TOTAL_PROFIT_TARGET_ENABLE = True
# - True : 총 수익 목표 도달 시 전체 청산 기능 활성화
# - False: 기능 비활성화 → 그리드 봇만 무한 반복
# - 기본값을 False로 두면 기존 봇과 동일하게 동작
TOTAL_PROFIT_TARGET_ENABLE = True

# 📌 TOTAL_PROFIT_TARGET_USDT = 5.0
# - 목표 누적 수익금 (USDT 단위)
# - "현재 미실현 손익 + 엑셀 누적실현손익" 합이 이 값 이상이면 전체 청산 트리거
# - 예: 5.0 → 5 USDT 모이면 청산
# - 예: 0.5 → 0.5 USDT만 모여도 청산 (소액 단위 익절)
# - 예: 50.0 → 50 USDT 모일 때까지 그리드 계속
TOTAL_PROFIT_TARGET_USDT = 10.0

# 📌 TOTAL_PROFIT_AFTER_LIQUIDATION = "STOP"
# - 전체 청산 후 봇 동작 모드 (3가지 중 하나):
#   "STOP"    : 봇 자체 종료 (가장 안전, 추천)
#               + 락 파일 유지 → 크론탭이 봇 부활시켜도 매수 안 함
#               + 완전 정지 원하면 크론탭도 주석 처리하는 게 좋음
#   "PAUSE"   : 봇은 계속 살아있되 매매만 정지 (관망 모드)
#               + 가격 출력 + 텔레그램 하트비트는 계속
#               + 락 파일 수동 해제 시 매매 재개
#   "RESTART" : 청산 직후 락 자동 해제 → 다음 폴링부터 1차 매수 재시작 (연속 익절)
#               ⚠️ 시장이 계속 우상향이면 무한 익절 가능, 하락 시 누적 손실 주의
TOTAL_PROFIT_AFTER_LIQUIDATION = "RESTART"

# 📌 PROFIT_SOURCE = "EXCHANGE"
# - 미실현 손익을 어디서 가져올지 결정:
#   "EXCHANGE" : 거래소 fetch_positions의 unrealizedPnl 사용 (정확, 권장)
#   "EXCEL"    : 엑셀 매수가 + 현재가로 자체 계산 (단순, API 호출 절약)
# - EXCHANGE 사용 시 fetch_positions 실패하면 자동으로 EXCEL fallback
PROFIT_SOURCE = "EXCEL"

# 📌 TOTAL_PROFIT_CHECK_INTERVAL = 5
# - 총 수익 목표를 체크하는 주기 (폴링 횟수 단위)
# - 5 = 매 5폴링(15초)마다 1회 체크 → 폴링 부담 최소화
# - 1 = 매 폴링(3초)마다 체크 → 빠른 반응
# - 너무 크면 목표 도달 후 청산 지연 발생 가능
TOTAL_PROFIT_CHECK_INTERVAL = 5

# 📌 TOTAL_PROFIT_LOG_INTERVAL = 20
# - 현재 총 수익을 콘솔/로그에 출력하는 주기 (폴링 횟수 단위)
# - 20 = 매 20폴링(60초)마다 1회 출력 → 디버깅/모니터링용
# - 출력 예: 💰 [총수익] 미실현=+1.23 + 누적실현=+3.45 = +4.68 USDT (목표 5.0)
# - 0 으로 두면 출력 안 함
TOTAL_PROFIT_LOG_INTERVAL = 20


# ===============================
# 📌 매매 전략 설정
# ===============================

# 📌 즉시매수허용 = True
# - True : 1차 매수 시 RSI 조건을 무시하고 즉시 매수 가능
# - False : 1차 매수 시 반드시 RSI_THRESHOLD 이하일 때만 매수
즉시매수허용 = True

# 📌 RSI_TIMEFRAME = "5m"
# - RSI 계산 시 사용할 차트 주기
# - 가능 값: "1m", "3m", "5m", "15m", "1h", "4h", "1d" 등
RSI_TIMEFRAME = "5m"

# 📌 RSI_PERIOD = 14
# - RSI 계산에 사용되는 기간 (봉 개수)
# - 14가 표준. 작을수록 민감, 클수록 둔감
RSI_PERIOD = 14

# 📌 RSI_THRESHOLD = 30.0
# - RSI가 이 값 이하일 때 '과매도'로 판단하여 1차 매수 트리거
# - 즉시매수허용 = False일 때만 작동
RSI_THRESHOLD = 30.0

# 📌 재진입여부 = True
# - 1차 청산 후 같은 종목에 다시 1차 매수를 넣을지 여부
# - True 로 두고 사용하세요 (그리드를 계속 반복)
# - False 기능은 아직 구현되지 않았습니다. 한 사이클만 돌리고 멈추려면
#   아래 '총 수익 목표 자동 청산' 의 STOP / PAUSE 를 사용하세요.
재진입여부 = True

# 📌 THREE_YIN_CANDLE_TIMEFRAME = "5m"
# - 3연속 음봉 조건 체크 시 사용할 차트 주기
THREE_YIN_CANDLE_TIMEFRAME = "5m"

# 📌 THREE_YIN_CANDLE_APPLY_LEVELS = []
# - 특정 차수에서만 '3연속 음봉 조건'을 추가로 적용하도록 지정
# - 예: [2, 4] → 2차/4차 매수 시 가격조건 + 3음봉조건 병합
# - 빈 리스트 [] : 모든 차수 가격조건만 사용
THREE_YIN_CANDLE_APPLY_LEVELS = []

# 📌 THREE_YIN_CANDLE_CONDITION_TYPE = "AND"
# - 가격조건과 3음봉 조건을 병합하는 방식
# - "AND" : 둘 다 만족해야 매수 (보수적)
# - "OR"  : 둘 중 하나만 만족해도 매수 (공격적)
THREE_YIN_CANDLE_CONDITION_TYPE = "AND"

# 📌 음봉전용차수 = []
# - 가격조건을 무시하고 오직 '3연속 음봉'만으로 매수 결정하는 차수
# - 예: [5, 7] → 5차/7차 매수 시 가격 무관, 3음봉만 체크
음봉전용차수 = []


# ===============================
# 📌 매매 종목 및 그리드 템플릿
# ===============================

# 📌 종목코드목록
# - 거래할 심볼 리스트 (Bitget 무기한선물 형식)
# - 여러 종목 동시 매매 시 메모리/API 호출 부담 증가
# 종목코드목록 = [
#     "BTC/USDT:USDT", "SOXL/USDT:USDT"
# ]

종목코드목록 = [
    "SOXL/USDT:USDT"
]


# 📌 GRID_TEMPLATE
# - 차수별 매수/매도 간격(%) 및 매수금액(USDT) 정의
# - 매수간격: 직전 차수 매수가 대비 % (음수 = 하락 시 매수)
# - 매도간격: 본인 차수 매수가 대비 % (양수 = 상승 시 매도)
# - 매수금액: 차수별 USDT 투입액
# - 차수가 깊어질수록 간격/금액을 키워 평균단가를 빠르게 낮추는 전략
GRID_TEMPLATE: Dict[int, Dict[str, Any]] = {
    1: {"매수간격": 0.0, "매도간격": 0.7, "매수금액": 10},
    2: {"매수간격": -0.7, "매도간격": 0.7, "매수금액": 10},
    3: {"매수간격": -0.7, "매도간격": 0.7, "매수금액": 10},
    4: {"매수간격": -0.7, "매도간격": 0.7, "매수금액": 10},
    5: {"매수간격": -0.7, "매도간격": 0.7, "매수금액": 10},
    6: {"매수간격": -0.7, "매도간격": 0.7, "매수금액": 10},
    7: {"매수간격": -0.7, "매도간격": 0.7, "매수금액": 10},
    8: {"매수간격": -0.7, "매도간격": 0.7, "매수금액": 10},
    9: {"매수간격": -0.7, "매도간격": 0.7, "매수금액": 10},
    10: {"매수간격": -0.7, "매도간격": 0.7, "매수금액": 10},
    **{i: {"매수간격": -0.7, "매도간격": 0.7, "매수금액": 10} for i in range(11, 13)},
    **{i: {"매수간격": -0.7, "매도간격": 0.7, "매수금액": 10} for i in range(13, 21)},
#    **{i: {"매수간격": -0.7, "매도간격": 0.7, "매수금액": 10} for i in range(21, 41)},
#    **{i: {"매수간격": -0.7, "매도간격": 0.7, "매수금액": 10} for i in range(41, 100)},
}

# 📌 PER_SYMBOL_GRID
# - 종목별로 GRID_TEMPLATE을 오버라이드하고 싶을 때 사용
PER_SYMBOL_GRID: Dict[str, Dict[int, Dict[str, Any]]] = {
    # "ETH/USDT:USDT": {
    #     1: {"매수간격": 0.0, "매도간격": 0.7, "매수금액": 150},
    # },
}

# 트레이드 설정 빌드 (종목별 그리드 + 파일 경로 + 수수료율)
trade_config: Dict[str, Dict[str, Any]] = {}
for ticker in 종목코드목록:
    base_symbol = ticker.split('/')[0]
    grid = copy.deepcopy(GRID_TEMPLATE)
    if ticker in PER_SYMBOL_GRID:
        grid.update(PER_SYMBOL_GRID[ticker])
    trade_config[ticker] = {
        # 거래 기록 엑셀 저장 경로 (EC2)
        "file_name": f"/var/autobot/bitget_grid/{base_symbol}-USDT_비트겟그리드_LONG폴.xlsx",
        # "file_name": f"./{base_symbol}-USDT_비트겟그리드_LONG폴.xlsx",
        "grid_config": grid,
        # 📌 수수료율 = 0.0008 (0.08%)
        # - Bitget USDT 무기한 taker 수수료 기준
        # - VIP 등급에 따라 변동 가능
        "수수료율": 0.0008,
    }

# --- 엑셀 컬럼 정의 (실제 체결가 + 슬리피지 컬럼 추가) ---
# - "매수가" / "매도가" 는 이제 실제 평균 체결가로 저장됨
# - "매수시현재가" / "매도시현재가" : 주문 직전 fetch_ticker로 조회한 시장가
# - "매수슬리피지%" / "매도슬리피지%" : 시장가 대비 체결가 차이 (%)
# ─────────────────────────────────────────────────────────────────
COLUMNS = [
    "종목", "차수",
    "매수날짜", "매수가", "매수수량", "매수금액",
    "매수시현재가", "매수슬리피지%",
    "매도날짜", "매도가", "매도수량", "매도금액",
    "매도시현재가", "매도슬리피지%",
    "수수료", "실현손익", "수익률", "누적실현손익"
]


# ===============================
# 안전 종료 처리 (SIGINT / SIGTERM)
# ===============================
# Ctrl+C(SIGINT) 또는 'pkill -TERM' 수신 시
# 진행 중인 작업을 마치고 깔끔하게 종료
shutdown_requested = False

def signal_handler(signum, frame):
    """종료 신호 수신 핸들러"""
    global shutdown_requested
    sig_name = "SIGINT" if signum == signal.SIGINT else "SIGTERM"
    print(f"\n[신호 수신] {sig_name} → 진행 중 작업 마치고 안전 종료합니다.")
    shutdown_requested = True

signal.signal(signal.SIGINT, signal_handler)
signal.signal(signal.SIGTERM, signal_handler)


# ===============================
# 파일 유틸 (엑셀 거래 기록)
# ===============================
def initialize_excel_file(file_name: str):
    """엑셀 파일이 없으면 컬럼만 있는 빈 파일 생성"""
    try:
        pd.read_excel(file_name)
    except FileNotFoundError:
        os.makedirs(os.path.dirname(file_name), exist_ok=True)
        df = pd.DataFrame(columns=COLUMNS)
        df.to_excel(file_name, index=False, engine="openpyxl")

def 매매기록불러오기(file_name: str) -> pd.DataFrame:
    """엑셀에서 거래 기록 로드 (날짜 컬럼 datetime 변환 포함)"""
    try:
        df = pd.read_excel(file_name, engine="openpyxl")
        if "매수날짜" in df:
            df["매수날짜"] = pd.to_datetime(df["매수날짜"], errors="coerce")
        if "매도날짜" in df:
            df["매도날짜"] = pd.to_datetime(df["매도날짜"], errors="coerce")

        # --- 누락된 신규 컬럼 자동 보완 ---
        # - 기존 엑셀 파일이 있는 경우 신규 컬럼이 없을 수 있음
        # - 누락된 컬럼은 NaN으로 채워서 호환성 유지
        # ─────────────────────────────────────────────────────
        for col in COLUMNS:
            if col not in df.columns:
                df[col] = pd.NA
        # 컬럼 순서를 표준 순서로 재정렬
        df = df[COLUMNS]

    except FileNotFoundError:
        df = pd.DataFrame(columns=COLUMNS)
        os.makedirs(os.path.dirname(file_name), exist_ok=True)
        df.to_excel(file_name, index=False, engine="openpyxl")
    return df

def 매매기록저장(df: pd.DataFrame, file_name: str):
    """
    거래 기록 엑셀 저장

    [안전장치] 저장 직전에 누적실현손익 컬럼을 자동 재계산합니다.
    청산 완료된 행을 매도날짜 순으로 정렬해 실현손익의 누적합을 다시 기록하므로,
    어떤 상황에서도 "실현손익 합계 = 마지막 행의 누적실현손익" 이 항상 맞습니다.
    차수별 손익 값 자체는 건드리지 않습니다.
    """
    try:
        # 청산 완료된 행만 추출 (매도날짜가 NaT가 아닌 행)
        cleared_mask = df["매도날짜"].notna()
        if cleared_mask.any():
            # 매도날짜 시간순으로 정렬한 인덱스 순서로 cumsum 계산
            cleared_df = df[cleared_mask].copy()
            cleared_df = cleared_df.sort_values(by=["매도날짜"], kind="stable")
            # 실현손익 NaN은 0으로 처리하여 누적합 계산
            cleared_df["누적실현손익"] = cleared_df["실현손익"].fillna(0).cumsum()
            # 원본 df의 해당 인덱스에 재계산된 누적값 반영
            df.loc[cleared_df.index, "누적실현손익"] = cleared_df["누적실현손익"]
            # 미청산 행의 누적실현손익은 명시적으로 NaN 유지
            df.loc[~cleared_mask, "누적실현손익"] = pd.NA
    except Exception as e:
        # 재계산 실패해도 원본 저장은 진행 (안전성)
        print(f"[누적실현손익 재계산 실패 - 원본대로 저장] {e}")

    df.to_excel(file_name, index=False, engine="openpyxl")


# --- 누적실현손익 시작값 안전 조회 (max() 버그 수정) ---
def 직전누적실현손익(df: pd.DataFrame) -> float:
    """
    가장 최근 청산된 행의 '누적실현손익' 값을 반환.
    - dropna().max() 사용 X: 음수 누적이 발생하면 옛날 양수를 잘못 가져옴
    - 매도날짜 기준 시간순 정렬 후 마지막 행의 값을 사용 → 진짜 직전 누적값
    - 청산 기록이 없으면 0.0 반환

    [중요] 이 함수는 한 청산 시작 시점에서 1회만 호출하고,
           이후 같은 청산 함수 내부에서는 '변수로 누적'해야 합니다.
           (이 함수를 매 차수 청산마다 호출하면 같은 값을 반복 반환할 수 있음)
    """
    try:
        # "매도날짜"가 NaN이 아닌 행만 (즉 청산 완료된 행만)
        cleared = df[df["매도날짜"].notna()].copy()
        if cleared.empty:
            return 0.0
        # 매도날짜 시간순 정렬, 마지막 행의 누적실현손익 추출
        cleared_sorted = cleared.sort_values(by=["매도날짜"], na_position="last")
        last_value = cleared_sorted.iloc[-1]["누적실현손익"]
        if pd.isna(last_value):
            return 0.0
        return float(last_value)
    except Exception as e:
        print(f"[직전누적실현손익 조회 오류] {e}")
        return 0.0


# ===============================
# 재진입 차단 유틸 (락 파일 기반)
# ===============================
def _reentry_lock_path(Ticker: str) -> str:
    """재진입 차단 락 파일 경로"""
    safe = Ticker.replace('/', '_').replace(':', '_')
    return f"/var/autobot/bitget_grid/.{safe}_no_reentry_gridlong.lock"

def 재진입차단여부(Ticker: str) -> bool:
    """현재 재진입이 차단되어 있는지 확인"""
    return os.path.exists(_reentry_lock_path(Ticker))

def 재진입차단설정(Ticker: str):
    """재진입여부=False일 때 락 파일 생성"""
    if not 재진입여부:
        try:
            os.makedirs("/var/autobot", exist_ok=True)
            with open(_reentry_lock_path(Ticker), "w") as f:
                f.write("locked")
        except Exception as e:
            print(f"[재진입 차단 저장 오류] {e}")

def 재진입차단해제(Ticker: str):
    """락 파일 제거 (수동 재개용)"""
    path = _reentry_lock_path(Ticker)
    if os.path.exists(path):
        try:
            os.remove(path)
        except Exception as e:
            print(f"[재진입 차단 해제 오류] {e}")


# --- 총 수익 청산 락 파일 유틸 ---
# 총 수익 목표 도달로 전체 청산이 한 번 일어나면 이 락 파일이 생성됩니다.
# 이후 봇이 재실행되어도 이 락이 있으면 매수가 차단됩니다.
# 다시 그리드를 시작하려면 락 파일을 수동으로 삭제해야 합니다.
def _총수익청산락_경로(Ticker: str) -> str:
    """총 수익 청산 락 파일 경로"""
    safe = Ticker.replace('/', '_').replace(':', '_')
    return f"/var/autobot/bitget_grid/.{safe}_total_profit_target_reached.lock"

def 총수익청산락_여부(Ticker: str) -> bool:
    """총 수익 청산 락이 걸려있는지 확인"""
    return os.path.exists(_총수익청산락_경로(Ticker))

def 총수익청산락_설정(Ticker: str, info: str = ""):
    """총 수익 청산 락 파일 생성 (정보 기록 포함)"""
    try:
        os.makedirs("/var/autobot", exist_ok=True)
        with open(_총수익청산락_경로(Ticker), "w") as f:
            f.write(f"locked at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n{info}")
    except Exception as e:
        print(f"[총수익청산락 설정 실패] {e}")

def 총수익청산락_해제(Ticker: str):
    """총 수익 청산 락 파일 제거 (수동 재시작용)"""
    path = _총수익청산락_경로(Ticker)
    if os.path.exists(path):
        try:
            os.remove(path)
        except Exception as e:
            print(f"[총수익청산락 해제 실패] {e}")


# ===============================
# 공통 유틸
# ===============================
def 현재가조회(Ticker: str) -> float:
    """현재가 조회 (실패 시 0.0 반환)"""
    try:
        ticker = bitgetX.fetch_ticker(Ticker)
        return float(ticker["last"])
    except Exception as e:
        print(f"[현재가 조회 실패] {Ticker} / {e}")
        return 0.0

def USDT금액을_수량으로변환(Ticker: str, 금액: float, 현재가: float = 0.0) -> float:
    """
    USDT 금액을 거래소 수량 단위로 변환
    - 현재가 인자를 받으면 fetch_ticker 중복 호출 방지 (성능 최적화)
    - 최소 주문 수량 미만이면 0.0 반환
    """
    try:
        가격 = 현재가 if 현재가 > 0 else 현재가조회(Ticker)
        if 가격 <= 0:
            raise ValueError("가격이 0이거나 음수입니다.")
        market = bitgetX.market(Ticker) if Ticker in bitgetX.markets else bitgetX.load_markets().get(Ticker, {})
        precision_raw = market.get('precision', {}).get('amount') if market else None
        precision = 4 if (precision_raw is None or precision_raw < 1) else int(precision_raw)
        min_qty = market.get('limits', {}).get('amount', {}).get('min', 0.0001) if market else 0.0001
        수량 = round(금액 / 가격, precision)
        if 수량 < min_qty:
            print(f"❌ 주문 실패: 계산된 수량 {수량} < 최소 수량 {min_qty}")
            return 0.0
        return float(수량)
    except Exception as e:
        print(f"[수량 변환 에러] {e}")
        return 0.0


# --- 실제 체결가 조회 함수 ---
def 체결가_조회(Ticker: str,
              order_result: Optional[dict],
              fallback_price: float,
              fallback_qty: float) -> Tuple[float, float, float]:
    """
    시장가 주문의 '실제 평균 체결가'를 조회합니다.

    [동작 순서]
    1. create_market_*_order 반환값(order_result)에서 즉시 average 확인
       (Bitget는 보통 즉시 체결되어 average가 채워져서 반환됨)
    2. 즉시 정보가 없으면 order_id로 fetch_order를 FILL_FETCH_RETRIES만큼 재시도
       (각 시도 사이에 FILL_FETCH_DELAY 초 대기)
    3. 모두 실패 시 fallback_price/fallback_qty 사용 (시장가 조회값)

    [반환값]
    (실제체결가, 실제체결수량, 실제체결금액)
    - 실제체결금액은 order에 cost가 있으면 그 값을, 없으면 가격×수량으로 계산
    """
    # 1단계: 즉시 응답에서 추출 시도
    if order_result:
        try:
            avg = order_result.get('average')
            filled = order_result.get('filled', 0) or 0
            cost = order_result.get('cost', 0) or 0
            if avg and float(avg) > 0 and float(filled) > 0:
                avg_f = float(avg)
                filled_f = float(filled)
                cost_f = float(cost) if cost else avg_f * filled_f
                return avg_f, filled_f, cost_f
        except Exception as e:
            print(f"[체결가 즉시 추출 실패] {e}")

    # 2단계: order_id로 재조회
    order_id = None
    try:
        if order_result:
            order_id = order_result.get('id')
    except Exception:
        pass

    if order_id:
        for attempt in range(1, FILL_FETCH_RETRIES + 1):
            try:
                time.sleep(FILL_FETCH_DELAY)
                fetched = bitgetX.fetch_order(order_id, Ticker)
                if fetched:
                    avg = fetched.get('average')
                    filled = fetched.get('filled', 0) or 0
                    cost = fetched.get('cost', 0) or 0
                    if avg and float(avg) > 0 and float(filled) > 0:
                        avg_f = float(avg)
                        filled_f = float(filled)
                        cost_f = float(cost) if cost else avg_f * filled_f
                        return avg_f, filled_f, cost_f
            except Exception as e:
                if attempt == FILL_FETCH_RETRIES:
                    print(f"[fetch_order 최종 실패 {attempt}/{FILL_FETCH_RETRIES}] {e}")

    # 3단계: fallback (시장가 조회값 사용)
    print(f"⚠️ 체결가 조회 실패 → fallback 사용 (가격={fallback_price}, 수량={fallback_qty})")
    return float(fallback_price), float(fallback_qty), float(fallback_price) * float(fallback_qty)


def 슬리피지_계산(시장가: float, 체결가: float) -> float:
    """
    슬리피지(%) 계산 = (체결가 - 시장가) / 시장가 × 100
    - 매수: 양수 = 시장가보다 비싸게 체결 (불리), 음수 = 싸게 체결 (유리)
    - 매도: 양수 = 시장가보다 비싸게 체결 (유리), 음수 = 싸게 체결 (불리)
    """
    if 시장가 <= 0:
        return 0.0
    return (체결가 - 시장가) / 시장가 * 100.0


# --- 미실현 손익 / 누적실현손익 / 총 수익 체크 함수 ---
def 미실현손익조회_거래소(Ticker: str) -> Optional[float]:
    """
    거래소 fetch_positions로 현재 포지션의 미실현 손익(USDT)을 조회.
    [반환] float : 정상 조회 (포지션 없으면 0.0)
            None  : API 실패 (호출 측에서 fallback 처리)
    """
    try:
        positions = bitgetX.fetch_positions([Ticker])
        unrealized = 0.0
        for pos in positions:
            if pos.get('symbol') != Ticker:
                continue
            contracts = float(pos.get('contracts', 0) or 0)
            if contracts == 0:
                continue  # 포지션 없음
            upnl = pos.get('unrealizedPnl', 0)
            if upnl is None:
                # ccxt 표준 필드 없으면 raw info에서 시도
                upnl = pos.get('info', {}).get('unrealisedPnl', 0) or 0
            unrealized += float(upnl)
        return unrealized
    except Exception as e:
        print(f"[미실현 손익 조회 실패 - 거래소] {e}")
        return None


def 미실현손익조회_엑셀(Ticker: str, 현재가: float) -> float:
    """
    엑셀 매수가/수량 + 현재가로 미실현 손익(USDT)을 자체 계산.
    수수료는 매도 예상가 × 수수료율로 추정 반영.
    """
    file_name = trade_config[Ticker]["file_name"]
    df = 매매기록불러오기(file_name)
    미청산 = df[df["매도날짜"].isnull()]
    if 미청산.empty:
        return 0.0

    수수료율 = float(trade_config[Ticker]["수수료율"])
    unrealized = 0.0
    for _, row in 미청산.iterrows():
        매수수량 = float(row["매수수량"])
        매수금액 = float(row["매수금액"])
        예상매도금액 = 현재가 * 매수수량
        예상수수료 = 예상매도금액 * 수수료율
        unrealized += (예상매도금액 - 매수금액 - 예상수수료)
    return unrealized


def 누적실현손익_조회(Ticker: str) -> float:
    """
    엑셀에서 가장 최근의 누적실현손익 값 조회.
    max() 대신 시간순 마지막 행의 누적값 사용 (음수 누적 시 정확).
    """
    file_name = trade_config[Ticker]["file_name"]
    try:
        df = 매매기록불러오기(file_name)
        return 직전누적실현손익(df)  # 헬퍼 함수 사용
    except Exception as e:
        print(f"[누적실현손익 조회 실패] {e}")
        return 0.0


def 총수익_계산(Ticker: str, 현재가: float) -> Tuple[float, float, float]:
    """
    총 수익 계산: 미실현 손익 + 누적실현손익
    [반환] (미실현, 누적실현, 합계)
    PROFIT_SOURCE 설정에 따라 미실현 손익 소스 선택.
    EXCHANGE 실패 시 EXCEL로 자동 fallback.
    """
    if PROFIT_SOURCE == "EXCHANGE":
        미실현 = 미실현손익조회_거래소(Ticker)
        if 미실현 is None:  # 거래소 API 실패 → EXCEL fallback
            미실현 = 미실현손익조회_엑셀(Ticker, 현재가)
    else:
        미실현 = 미실현손익조회_엑셀(Ticker, 현재가)

    누적실현 = 누적실현손익_조회(Ticker)
    총합 = 미실현 + 누적실현
    return 미실현, 누적실현, 총합


def 총수익_체크_및_청산(Ticker: str, 현재가: float) -> bool:
    """
    매 폴링마다 호출되어 총 수익이 목표를 도달하면 전체 차수를 시장가로 청산.
    [반환] True : 청산 발생 (이번 폴링은 매도/매수 모두 스킵)
            False: 청산 없음 (정상 매도/매수 진행)

    [동작 순서]
    1. 기능 활성화 여부 + 락 파일 체크
    2. 미실현 + 누적실현 합산
    3. 목표 미달 → False 반환
    4. 목표 도달 → 미청산 차수 전체 시장가 매도
    5. 엑셀 갱신 + 텔레그램 알림 + 락 파일 생성
    6. TOTAL_PROFIT_AFTER_LIQUIDATION 설정에 따라 후속 동작
    """
    global shutdown_requested

    # ── 1) 기능 비활성 또는 이미 청산 완료 락이 있으면 스킵 ──
    if not TOTAL_PROFIT_TARGET_ENABLE:
        return False
    if 총수익청산락_여부(Ticker):
        return False

    # ── 2) 총 수익 계산 ──
    미실현, 누적실현, 총수익 = 총수익_계산(Ticker, 현재가)

    # ── 3) 목표 미달 시 종료 ──
    if 총수익 < TOTAL_PROFIT_TARGET_USDT:
        return False

    # ── 4) 목표 도달 → 전체 청산 시작 ──
    print(
        f"🎯 총 수익 목표 도달! "
        f"미실현 {미실현:+.4f} + 누적실현 {누적실현:+.4f} = {총수익:+.4f} USDT "
        f"(목표 {TOTAL_PROFIT_TARGET_USDT})",
        flush=True
    )

    file_name = trade_config[Ticker]["file_name"]
    df = 매매기록불러오기(file_name)
    미청산 = df[df["매도날짜"].isnull()]

    청산_차수_수 = 0
    총_청산_손익 = 0.0
    수수료율 = float(trade_config[Ticker]["수수료율"])

    # --- 핵심: 누적값을 변수로 추적 ---
    # - 청산 루프 시작 전, 직전 청산 행의 누적값을 1회 조회 (시간순 마지막 행)
    # - 이후 청산되는 차수마다 이 변수에 실현손익을 더해가며 갱신
    # - 더 이상 dropna().max()로 매번 재조회하지 않음 → 음수 누적도 정확
    누적_추적기 = 직전누적실현손익(df)

    if 미청산.empty:
        # 미청산 포지션 없는데 수익 목표 도달 (거의 누적실현만으로) → 락만 걸고 종료
        print(f"ℹ️ 미청산 포지션 없음. 누적실현만으로 목표 도달.")
    else:
        for idx, row in 미청산.sort_values(by=["매수날짜", "차수"]).iterrows():
            차수 = int(row["차수"])
            매수가 = float(row["매수가"])
            매수수량 = float(row["매수수량"])
            매수금액 = float(row["매수금액"])

            try:
                order_result = bitgetX.create_market_sell_order(
                    symbol=Ticker, amount=매수수량, params={"reduceOnly": True}
                )
            except Exception as e:
                print(f"[전체 청산 - 차수 {차수} 실패] {e}")
                continue

            실제체결가, 실제체결수량, 실제체결금액 = 체결가_조회(
                Ticker, order_result, fallback_price=현재가, fallback_qty=매수수량
            )
            매도슬리피지 = 슬리피지_계산(현재가, 실제체결가)

            매도금액 = 실제체결금액
            수수료 = 매도금액 * 수수료율
            실현손익 = 매도금액 - 매수금액 - 수수료
            수익률 = (실현손익 / 매수금액) * 100.0 if 매수금액 else 0.0

            # --- 변수로 누적 (max() 미사용) ---
            누적_추적기 += 실현손익
            현재누적 = 누적_추적기

            df.at[idx, "매도날짜"] = datetime.now()
            df.at[idx, "매도가"] = 실제체결가
            df.at[idx, "매도수량"] = 실제체결수량
            df.at[idx, "매도금액"] = 매도금액
            df.at[idx, "매도시현재가"] = 현재가
            df.at[idx, "매도슬리피지%"] = round(매도슬리피지, 4)
            df.at[idx, "수수료"] = 수수료
            df.at[idx, "실현손익"] = 실현손익
            df.at[idx, "수익률"] = 수익률
            df.at[idx, "누적실현손익"] = 현재누적

            print(
                f"💸 [전체청산] {Ticker} 차수 {차수} | "
                f"체결가 {실제체결가:.{PRICE_LOG_DECIMALS}f} | "
                f"손익 {실현손익:+.4f} USDT | "
                f"누적 {현재누적:+.4f}",
                flush=True
            )

            청산_차수_수 += 1
            총_청산_손익 += 실현손익

        매매기록저장(df, file_name)

    # ── 5) 텔레그램 종합 알림 ──
    title_symbol = Ticker.split('/')[0]
    try:
        telegram_alert.SendMessage(
            f"🎯 *[{title_symbol}_총 수익 목표 달성 - 전체 청산]*\n"
            f"━━━━━━━━━━━━━━━\n"
            f"💵 청산 시점가: `{현재가:.6f} USDT`\n"
            f"💰 청산 전 미실현: `{미실현:+.4f} USDT`\n"
            f"💼 청산 전 누적실현: `{누적실현:+.4f} USDT`\n"
            f"🎯 합계: `{총수익:+.4f} USDT` (목표 {TOTAL_PROFIT_TARGET_USDT})\n"
            f"📋 청산 차수: `{청산_차수_수}개`\n"
            f"💸 청산 시 실현손익: `{총_청산_손익:+.4f} USDT`\n"
            f"⚙️ 후속 동작: `{TOTAL_PROFIT_AFTER_LIQUIDATION}`\n"
            f"🕒 시간: `{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}`\n"
            f"━━━━━━━━━━━━━━━"
        )
    except Exception as e:
        print(f"[총수익 청산 알림 실패] {e}")

    # ── 6) 후속 동작 처리 ──
    if TOTAL_PROFIT_AFTER_LIQUIDATION == "STOP":
        # 락 파일 생성 + 봇 종료 신호
        총수익청산락_설정(Ticker, info=f"reason=STOP, total={총수익:.4f}, cleared={청산_차수_수}")
        shutdown_requested = True
        print(f"⏹️ [후속동작=STOP] 봇 종료 진행. 락 파일 유지 → 재실행해도 매수 안 함.", flush=True)

    elif TOTAL_PROFIT_AFTER_LIQUIDATION == "PAUSE":
        # 락 파일만 생성, 봇은 계속 살아있음
        총수익청산락_설정(Ticker, info=f"reason=PAUSE, total={총수익:.4f}, cleared={청산_차수_수}")
        print(f"⏸️ [후속동작=PAUSE] 매매 정지. 모니터링/하트비트는 계속.", flush=True)

    elif TOTAL_PROFIT_AFTER_LIQUIDATION == "RESTART":
        # 락 파일 만들지 않음 → 다음 폴링부터 1차 매수 재시작
        print(f"🔄 [후속동작=RESTART] 다음 폴링부터 1차 매수 재시작.", flush=True)

    else:
        # 알 수 없는 옵션 → 안전상 STOP 처리
        총수익청산락_설정(Ticker, info=f"reason=UNKNOWN({TOTAL_PROFIT_AFTER_LIQUIDATION}) → STOP")
        shutdown_requested = True
        print(f"⚠️ TOTAL_PROFIT_AFTER_LIQUIDATION='{TOTAL_PROFIT_AFTER_LIQUIDATION}' 알 수 없음 → STOP 처리", flush=True)

    return True


# ===============================
# 3음봉 조건 체크
# ===============================
def is_three_yin_candles(Ticker, timeframe):
    """
    최근 3개의 봉이 모두 음봉(종가 < 시가)인지 체크
    - 가장 최근 봉(아직 미완성)은 제외하고 그 이전 3개를 검사
    """
    try:
        ohlcv = bitgetX.fetch_ohlcv(Ticker, timeframe=timeframe, limit=5)
        if len(ohlcv) < 4:
            return False
        for i in range(2, 5):
            candle = ohlcv[-i]
            # candle: [timestamp, open, high, low, close, volume]
            if candle[4] >= candle[1]:  # 종가 >= 시가 → 양봉이거나 도지
                return False
        return True
    except Exception as e:
        print(f"[3음봉 체크 오류] {e}")
        return False


# ===============================
# RSI 계산 (Wilder's smoothing 방식)
# ===============================
def calculate_rsi(close_prices, period: int = 14) -> float:
    """종가 시계열로부터 RSI 값을 계산"""
    if not isinstance(close_prices, pd.Series):
        close_prices = pd.Series(close_prices)
    delta = close_prices.diff()
    gain = delta.where(delta > 0, 0)
    loss = -delta.where(delta < 0, 0)
    avg_gain = gain.iloc[1:period + 1].mean()
    avg_loss = loss.iloc[1:period + 1].mean()
    rsi_values = [0] * (period - 1)
    for i in range(period - 1, len(close_prices)):
        avg_gain = (avg_gain * (period - 1) + gain.iloc[i]) / period
        avg_loss = (avg_loss * (period - 1) + loss.iloc[i]) / period
        rsi = 100 if avg_loss == 0 else 100 - (100 / (1 + (avg_gain / avg_loss)))
        rsi_values.append(rsi)
    rsi_series = pd.Series(rsi_values, index=close_prices.index[max(0, len(close_prices) - len(rsi_values)):])
    return float(rsi_series.iloc[-1]) if not rsi_series.empty else 0.0

def 최근_RSI(Ticker: str, timeframe: str = "5m", period: int = 14) -> float:
    """현재 시점의 RSI 값 반환 (오류 시 100.0 반환 → 매수 차단 효과)"""
    try:
        ohlcv = bitgetX.fetch_ohlcv(Ticker, timeframe=timeframe, limit=period + 100)
        if not ohlcv or len(ohlcv) < period + 1:
            return 100.0
        df_ohlcv = pd.DataFrame(ohlcv, columns=["timestamp", "open", "high", "low", "close", "volume"])
        return calculate_rsi(df_ohlcv["close"], period=period)
    except Exception as e:
        print(f"[RSI 계산 실패] {Ticker} / {e}")
        return 100.0


# ===============================
# 매수 실행 함수
# ===============================
def 매수실행(Ticker: str, 현재가: float):
    """
    그리드 매수 로직
    - 1차: 즉시매수허용 또는 RSI 조건 충족 시 진입
    - 2차+: 직전 차수 매수가 기준 % 하락 시 진입
    - 한 번의 폴링에서는 최대 1개 차수만 진입 (break)
    - 실제 체결가/수량을 조회해서 엑셀에 기록
    - 총 수익 청산 락이 걸려있으면 매수 차단
    - 총수익 옵션 활성화 시 텔레그램에 미실현/누적실현/총수익 표시
    """
    # --- 총 수익 청산 락 체크 ---
    if 총수익청산락_여부(Ticker):
        return  # 락 걸림 → 매수 안 함

    file_name = trade_config[Ticker]["file_name"]
    df = 매매기록불러오기(file_name)
    if 현재가 <= 0:
        return

    for 차수, config in trade_config[Ticker]["grid_config"].items():
        # 이미 해당 차수 미청산 포지션이 있으면 스킵
        if not df[(df["차수"] == 차수) & (df["매도날짜"].isnull())].empty:
            continue

        # ─── 가격 조건 판정 ───
        if 차수 == 1:
            # 1차 매수: 재진입 차단 체크 → RSI 또는 즉시매수
            if (not 재진입여부) and 재진입차단여부(Ticker):
                print(f"🚫 1차 매수 차단: 재진입 차단 활성화")
                continue
            if 즉시매수허용:
                rsi_val = 최근_RSI(Ticker, RSI_TIMEFRAME, RSI_PERIOD)
                가격조건 = True
            else:
                rsi_val = 최근_RSI(Ticker, RSI_TIMEFRAME, RSI_PERIOD)
                가격조건 = rsi_val <= RSI_THRESHOLD
                if not 가격조건:
                    print(f"🚫 매수 미실행: RSI 미충족 (RSI {rsi_val:.2f} > {RSI_THRESHOLD})")
                    continue
        else:
            # 2차+ 매수: 직전 차수 미청산 포지션 필요 + 가격 하락 조건
            prev_open = df[(df["차수"] == 차수 - 1) & (df["매도날짜"].isnull())]
            if prev_open.empty:
                continue
            이전매수가 = float(prev_open.iloc[-1]["매수가"])
            target_price = 이전매수가 * (1 + float(config["매수간격"]) / 100.0)
            가격조건 = 현재가 <= target_price
            if not 가격조건:
                continue  # 폴링이 빠르므로 매번 로그 찍지 않음 (로그 폭주 방지)

        # ─── 3음봉 조건 판정 ───
        삼음봉조건 = is_three_yin_candles(Ticker, THREE_YIN_CANDLE_TIMEFRAME)
        음봉전용 = 차수 in 음봉전용차수

        # ─── 조건 병합 ───
        if 음봉전용:
            매수트리거 = 삼음봉조건
        elif 차수 in THREE_YIN_CANDLE_APPLY_LEVELS:
            if THREE_YIN_CANDLE_CONDITION_TYPE.upper() == "AND":
                매수트리거 = 가격조건 and 삼음봉조건
            else:
                매수트리거 = 가격조건 or 삼음봉조건
        else:
            매수트리거 = 가격조건

        if not 매수트리거:
            continue

        # ─── 주문 실행 ───
        수량 = USDT금액을_수량으로변환(Ticker, float(config["매수금액"]), 현재가)
        if 수량 == 0.0:
            print(f"🚫 매수 미실행: 최소 수량 미만")
            continue
        try:
            # --- 주문 결과를 변수로 받음 (이전엔 결과 무시) ---
            order_result = bitgetX.create_market_buy_order(
                symbol=Ticker, amount=수량, params={"reduceOnly": False}
            )
        except Exception as e:
            print(f"[매수 주문 실패] {e}")
            continue

        # --- 실제 체결가/수량/금액 조회 + 슬리피지 계산 ---
        실제체결가, 실제체결수량, 실제체결금액 = 체결가_조회(
            Ticker, order_result, fallback_price=현재가, fallback_qty=수량
        )
        매수슬리피지 = 슬리피지_계산(현재가, 실제체결가)

        # ─── 거래 기록 추가 ───
        new_row = pd.DataFrame([{
            "종목": Ticker,
            "차수": 차수,
            "매수날짜": datetime.now(),
            "매수가": 실제체결가,                # 실제 평균 체결가
            "매수수량": 실제체결수량,            # 실제 체결 수량
            "매수금액": 실제체결금액,            # 실제 체결 금액
            "매수시현재가": 현재가,              # 주문 직전 시장가
            "매수슬리피지%": round(매수슬리피지, 4),
        }])
        df = pd.concat([df, new_row], ignore_index=True)
        매매기록저장(df, file_name)

        # ─── 텔레그램 알림 ───
        title_symbol = Ticker.split('/')[0]
        다음차 = 차수 + 1
        다음매수가 = 0.0
        if 다음차 in trade_config[Ticker]["grid_config"]:
            # 다음 차수 목표가는 '실제 체결가' 기준으로 계산
            다음매수가 = 실제체결가 * (1 + float(trade_config[Ticker]["grid_config"][다음차]["매수간격"]) / 100.0)

        이유 = f"가격조건={가격조건}, 삼음봉={삼음봉조건}, 음봉전용={음봉전용}"
        message = (
            f"🛒🚨📈💥🌟 *[{title_symbol}_그리드롱 진입 완료]* 🌟💥📈🚨🛒\n"
            f"━━━━━━━━━━━━━━━\n"
            f"📊 *종목*: `{Ticker}`\n"
            f"🔢 *차수*: `{차수}`\n"
            f"💵 *시장가*: `{현재가:.6f} USDT`\n"
            f"🎯 *실제 체결가*: `{실제체결가:.6f} USDT`\n"
            f"📐 *슬리피지*: `{매수슬리피지:+.4f}%`\n"
            f"📦 *체결수량*: `{실제체결수량}`\n"
            f"💰 *체결금액*: `{실제체결금액:.6f} USDT`\n"
            f"🎯 *목표 매도가*: `{실제체결가 * (1 + config['매도간격'] / 100.0):.6f} USDT`\n"
            f"📉 *다음 차수 매수 목표가*: `{다음매수가:.6f} USDT`\n"
            f"🧩 *매수 트리거*: `{이유}`\n"
            f"🕒 *시간*: `{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}`\n"
            f"━━━━━━━━━━━━━━━"
        )

        # --- 추가: 총수익 옵션 활성화 시 손익 정보를 메시지에 추가 ---
        # - 매수 직후(엑셀에 새 행 기록 완료 후) 총수익 계산
        # - 미실현(현재가 기준) + 누적실현 + 합계를 시간 라인 앞에 삽입
        # - TOTAL_PROFIT_TARGET_ENABLE = False면 추가 안 함 (기존 메시지 그대로 유지)
        # - 이번에 매수한 차수도 미실현 손익 계산에 자동 포함됨
        #   (방금 엑셀 저장 → 미실현손익조회 함수가 새 행도 읽어옴)
        # - 거래소 API 사용 시(PROFIT_SOURCE=EXCHANGE): 약간의 시간 지연 발생 가능
        #   (Bitget이 새 포지션 반영하는 데 0.5~2초 정도 걸릴 수 있음)
        #   → fetch_positions가 아직 반영 전 값이면 EXCEL 모드 결과와 미세 차이 가능
        # ─────────────────────────────────────────────────────
        if TOTAL_PROFIT_TARGET_ENABLE:
            try:
                미실현_msg, 누적실현_msg, 총수익_msg = 총수익_계산(Ticker, 현재가)
                손익라인 = (
                    f"💰 *현재 미실현 손익*: `{미실현_msg:+.4f} USDT`\n"
                    f"💼 *엑셀 누적실현*: `{누적실현_msg:+.4f} USDT`\n"
                    f"🎯 *총수익*: `{총수익_msg:+.4f} USDT` (목표 {TOTAL_PROFIT_TARGET_USDT} USDT)\n"
                )
                # 시간 라인 앞에 손익 라인 삽입 (기존 메시지 구조 유지)
                message = message.replace(
                    f"🕒 *시간*",
                    f"{손익라인}🕒 *시간*"
                )
            except Exception as e:
                print(f"[매수 텔레그램 - 총수익 추가 실패] {e}")

        telegram_alert.SendMessage(message)

        # --- +v2: 매수 완료 콘솔 출력 (실제 체결가 + 슬리피지 포함) ---
        print(
            f"✅ 매수 완료 - {Ticker} 차수 {차수} | "
            f"체결가 {실제체결가:.{PRICE_LOG_DECIMALS}f} "
            f"(시장가 {현재가:.{PRICE_LOG_DECIMALS}f}, 슬리피지 {매수슬리피지:+.4f}%)",
            flush=True
        )

        break  # 한 폴링에서는 한 차수만 진입 (안전성)


# ===============================
# 매도 실행 함수
# ===============================
def 매도실행(Ticker: str, 현재가: float) -> bool:
    """
    그리드 매도 로직
    - 미청산 포지션 중 목표가에 도달한 차수를 모두 청산
    - 한 폴링에서 여러 차수가 동시에 청산될 수 있음 (큰 변동성 대응)
    - 반환값: 매도 발생 시 True (이번 폴링에서 매수는 스킵)
    - 실제 체결가/수량/슬리피지 엑셀 기록
    - 누적실현손익 갱신을 변수 누적 방식으로 변경 (max() 버그 수정)
    """
    file_name = trade_config[Ticker]["file_name"]
    df = 매매기록불러오기(file_name)
    if 현재가 <= 0:
        return False

    미청산 = df[df["매도날짜"].isnull()]
    if 미청산.empty:
        return False

    청산된_차수_수 = 0
    총실현손익 = 0.0
    총수익률합 = 0.0
    청산된_포지션수 = 0
    title_symbol = Ticker.split('/')[0]
    매도발생 = False

    # --- 핵심: 누적값을 변수로 추적 ---
    # - 청산 루프 시작 전, 직전 청산 행의 누적값을 1회 조회 (시간순 마지막 행)
    # - 이후 청산되는 차수마다 이 변수에 실현손익을 더해가며 갱신
    # - 더 이상 dropna().max()로 매번 재조회하지 않음 → 음수 누적도 정확
    누적_추적기 = 직전누적실현손익(df)

    # 매수 시간/차수 순으로 정렬하여 청산
    for idx, row in 미청산.sort_values(by=["매수날짜", "차수"]).iterrows():
        차수 = int(row["차수"])
        매수가 = float(row["매수가"])
        매수수량 = float(row["매수수량"])
        매수금액 = float(row["매수금액"])
        config = trade_config[Ticker]["grid_config"].get(차수)
        if not config:
            continue
        목표가 = 매수가 * (1 + float(config["매도간격"]) / 100.0)

        if 현재가 >= 목표가:
            try:
                # --- 주문 결과를 변수로 받음 ---
                order_result = bitgetX.create_market_sell_order(
                    symbol=Ticker, amount=매수수량, params={"reduceOnly": True}
                )
            except Exception as e:
                print(f"[매도 주문 실패] {e}")
                continue

            # --- 실제 체결가/수량/금액 조회 + 슬리피지 계산 ---
            실제체결가, 실제체결수량, 실제체결금액 = 체결가_조회(
                Ticker, order_result, fallback_price=현재가, fallback_qty=매수수량
            )
            매도슬리피지 = 슬리피지_계산(현재가, 실제체결가)

            # 손익 계산도 실제 체결값 기반으로 변경
            매도금액 = 실제체결금액
            수수료 = 매도금액 * float(trade_config[Ticker]["수수료율"])
            실현손익 = 매도금액 - 매수금액 - 수수료
            수익률 = (실현손익 / 매수금액) * 100.0 if 매수금액 else 0.0

            # --- 변수로 누적 (max() 미사용) ---
            누적_추적기 += 실현손익
            현재누적손익 = 누적_추적기

            df.at[idx, "매도날짜"] = datetime.now()
            df.at[idx, "매도가"] = 실제체결가
            df.at[idx, "매도수량"] = 실제체결수량
            df.at[idx, "매도금액"] = 매도금액
            df.at[idx, "매도시현재가"] = 현재가
            df.at[idx, "매도슬리피지%"] = round(매도슬리피지, 4)
            df.at[idx, "수수료"] = 수수료
            df.at[idx, "실현손익"] = 실현손익
            df.at[idx, "수익률"] = 수익률
            df.at[idx, "누적실현손익"] = 현재누적손익

            message = (
                f"🛒🚨📈💥🌟 *[{title_symbol}_그리드롱 청산 완료]* 🌟💥📈🚨🛒\n"
                f"━━━━━━━━━━━━━━━\n"
                f"📊 *종목*: `{Ticker}`\n"
                f"🔢 *차수*: `{차수}`\n"
                f"💵 *시장가*: `{현재가:.6f} USDT`\n"
                f"🎯 *실제 체결가*: `{실제체결가:.6f} USDT`\n"
                f"📐 *슬리피지*: `{매도슬리피지:+.4f}%`\n"
                f"📦 *체결수량*: `{실제체결수량}`\n"
                f"💰 *매도금액*: `{매도금액:.6f} USDT`\n"
                f"💸 *실현 손익*: `{실현손익:.4f} USDT`\n"
                f"📈 *수익률*: `{수익률:.4f}%`\n"
                f"💼 *누적 실현손익*: `{현재누적손익:.4f} USDT`\n"
                f"🕒 *시간*: `{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}`\n"
                f"━━━━━━━━━━━━━━━"
            )
            telegram_alert.SendMessage(message)

            # --- +v2: 매도 완료 콘솔 출력 (실제 체결가 + 슬리피지 포함) ---
            print(
                f"💸 매도 완료 - {Ticker} 차수 {차수} | "
                f"체결가 {실제체결가:.{PRICE_LOG_DECIMALS}f} "
                f"(시장가 {현재가:.{PRICE_LOG_DECIMALS}f}, 슬리피지 {매도슬리피지:+.4f}%) | "
                f"손익 {실현손익:+.4f} USDT | "
                f"수익률 {수익률:+.4f}% | "
                f"누적 {현재누적손익:+.4f}",
                flush=True
            )

            청산된_차수_수 += 1
            총실현손익 += 실현손익
            총수익률합 += 수익률
            청산된_포지션수 += 1
            매도발생 = True

    매매기록저장(df, file_name)

    # 여러 차수가 동시 청산되면 종합 요약 알림
    if 청산된_포지션수 > 0:
        평균수익률 = 총수익률합 / 청산된_포지션수
        summary_message = (
            f"📊 *[그리드롱 청산 요약]*\n"
            f"━━━━━━━━━━━━━━━\n"
            f"📊 *종목*: `{Ticker}`\n"
            f"📋 *총 청산 차수*: `{청산된_차수_수}개`\n"
            f"💸 *총 실현 손익*: `{총실현손익:.4f} USDT`\n"
            f"📈 *평균 수익률*: `{평균수익률:.4f}%`\n"
            f"🕒 *시간*: `{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}`\n"
            f"━━━━━━━━━━━━━━━"
        )
        telegram_alert.SendMessage(summary_message)

        # --- 다중 청산 시 요약 콘솔 출력 ---
        if 청산된_포지션수 >= 2:
            print(
                f"📋 [{Ticker}] 다중 청산 요약 - "
                f"{청산된_차수_수}개 차수 | "
                f"총손익 {총실현손익:+.4f} USDT | "
                f"평균수익률 {평균수익률:+.4f}%",
                flush=True
            )

    return 매도발생


# ===============================
# 메인 루프 (실시간 폴링 데몬)
# ===============================
def main_loop():
    """
    봇의 핵심 루프
    - shutdown_requested가 True가 될 때까지 무한 반복
    - 매 POLL_INTERVAL초마다 종목별로 매도 → 매수 순서 체크
    - 예외 발생 시 지수 백오프로 재시도
    - HEARTBEAT_INTERVAL마다 텔레그램으로 생존 신호 전송
    - PRICE_LOG_INTERVAL 폴링마다 콘솔/로그에 현재가 출력 
    - TOTAL_PROFIT_CHECK_INTERVAL 폴링마다 총 수익 목표 체크 
    """
    global shutdown_requested

    print(f"🚀 비트겟 그리드 봇 실시간 모드 시작")
    print(f"   - 폴링 주기: {POLL_INTERVAL}초")
    print(f"   - 하트비트: {HEARTBEAT_INTERVAL}초")
    print(f"   - 가격 출력: {PRICE_LOG_INTERVAL * POLL_INTERVAL}초마다 (PRICE_LOG_INTERVAL={PRICE_LOG_INTERVAL})")
    print(f"   - 체결가 조회: 최대 {FILL_FETCH_RETRIES}회 재시도, {FILL_FETCH_DELAY}초 간격")
    # --- 총 수익 목표 설정 표시 ---
    if TOTAL_PROFIT_TARGET_ENABLE:
        print(f"   - 총 수익 목표: {TOTAL_PROFIT_TARGET_USDT} USDT 도달 시 전체 청산 (소스={PROFIT_SOURCE}, 후속={TOTAL_PROFIT_AFTER_LIQUIDATION})")
    else:
        print(f"   - 총 수익 목표: 비활성화 (TOTAL_PROFIT_TARGET_ENABLE=False)")
    print(f"   - 종목: {종목코드목록}", flush=True)

    # 시작 알림
    try:
        telegram_alert.SendMessage(
            f"🚀 *[비트겟 그리드봇 시작]*\n"
            f"━━━━━━━━━━━━━━━\n"
            f"⏱ 폴링 주기: `{POLL_INTERVAL}초`\n"
            f"📊 종목: `{', '.join(종목코드목록)}`\n"
            f"🎯 총수익 목표: `{TOTAL_PROFIT_TARGET_USDT} USDT`"
            f"{' (활성)' if TOTAL_PROFIT_TARGET_ENABLE else ' (비활성)'}\n"
            f"🕒 시간: `{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}`\n"
            f"━━━━━━━━━━━━━━━"
        )
    except Exception as e:
        print(f"[시작 알림 실패] {e}")

    # 종목별 엑셀 파일 초기화
    for Ticker in 종목코드목록:
        try:
            initialize_excel_file(trade_config[Ticker]["file_name"])
        except Exception as e:
            print(f"[파일 초기화 실패] {Ticker} / {e}")

    last_heartbeat = time.time()
    consecutive_errors = 0
    loop_count = 0

    # ─── 메인 무한 루프 ───
    while not shutdown_requested:
        loop_start = time.time()
        loop_count += 1

        # --- 이번 폴링에서 조회한 가격을 보관 (출력용 재사용) ---
        # - 같은 가격을 매도/매수/출력 모두에 사용 → fetch_ticker 중복 호출 방지
        # ─────────────────────────────────────────────────────────
        가격_캐시: Dict[str, float] = {}

        try:
            for Ticker in 종목코드목록:
                if shutdown_requested:
                    break
                try:
                    # 루프당 fetch_ticker 1회만 호출 (매도/매수에 동일 가격 사용)
                    현재가 = 현재가조회(Ticker)
                    if 현재가 <= 0:
                        continue

                    가격_캐시[Ticker] = 현재가  # 출력용 캐시 저장

                    # --- 총 수익 목표 체크 (매도/매수 보다 먼저) ---
                    # - 일정 폴링마다 체크 (TOTAL_PROFIT_CHECK_INTERVAL)
                    # - 청산 발생 시 이번 폴링은 그리드 매도/매수 스킵
                    # - 총 수익 로그 출력 (TOTAL_PROFIT_LOG_INTERVAL 폴링마다)
                    # ─────────────────────────────────────────────────
                    if TOTAL_PROFIT_TARGET_ENABLE and (loop_count % TOTAL_PROFIT_CHECK_INTERVAL == 0):
                        총수익_청산_발생 = 총수익_체크_및_청산(Ticker, 현재가)
                        if 총수익_청산_발생:
                            continue  # 다음 종목으로 (이번 폴링 매도/매수 스킵)

                    # 총 수익 모니터링 로그 (선택)
                    if (TOTAL_PROFIT_TARGET_ENABLE
                        and TOTAL_PROFIT_LOG_INTERVAL > 0
                        and loop_count % TOTAL_PROFIT_LOG_INTERVAL == 0):
                        try:
                            미실현_log, 누적실현_log, 총수익_log = 총수익_계산(Ticker, 현재가)
                            락_표시 = " 🔒락" if 총수익청산락_여부(Ticker) else ""
                            print(
                                f"💰 [총수익] {Ticker.split('/')[0]} | "
                                f"미실현={미실현_log:+.4f} + 누적실현={누적실현_log:+.4f} "
                                f"= {총수익_log:+.4f} USDT (목표 {TOTAL_PROFIT_TARGET_USDT}){락_표시}",
                                flush=True
                            )
                        except Exception as e:
                            print(f"[총수익 로그 실패] {e}")

                    # 매도 먼저 → 매도 발생 시 같은 폴링에서 매수 스킵
                    매도발생 = 매도실행(Ticker, 현재가)
                    if not 매도발생:
                        매수실행(Ticker, 현재가)
                except Exception as e:
                    print(f"[{Ticker} 처리 오류] {e}")

            # 한 사이클 정상 완료 → 에러 카운터 리셋
            consecutive_errors = 0

            # --- 실시간 현재가 콘솔/로그 출력 ---
            # - PRICE_LOG_INTERVAL 폴링마다 한 줄 출력 (기본 30초마다)
            # - 추가 API 호출 없이 가격_캐시 재사용 → rate limit 안전
            # - flush=True : 크론탭 리다이렉션 환경에서 tail -f 즉시 반영
            # - 출력 예: 📊 [09:23:42] 폴링#20 | Q=0.073180
            # ─────────────────────────────────────────────────────
            if loop_count % PRICE_LOG_INTERVAL == 0 and 가격_캐시:
                가격_표시 = " | ".join(
                    f"{t.split('/')[0]}={p:.{PRICE_LOG_DECIMALS}f}"
                    for t, p in 가격_캐시.items()
                )
                print(
                    f"📊 [{datetime.now().strftime('%H:%M:%S')}] "
                    f"폴링#{loop_count} | {가격_표시}",
                    flush=True
                )

            # 하트비트 알림 (10분마다)
            now = time.time()
            if now - last_heartbeat >= HEARTBEAT_INTERVAL:
                try:
                    가격정보 = []
                    for t in 종목코드목록:
                        p = 현재가조회(t)
                        if p > 0:
                            가격정보.append(f"{t.split('/')[0]}: {p:.4f}")
                    가격라인 = "\n".join(가격정보) if 가격정보 else "가격조회 실패"
                    telegram_alert.SendMessage(
                        f"💓 *[비트겟 그리드봇 생존]*\n"
                        f"━━━━━━━━━━━━━━━\n"
                        f"🔁 누적 폴링: `{loop_count}회`\n"
                        f"💵 현재가:\n{가격라인}\n"
                        f"🕒 `{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}`\n"
                        f"━━━━━━━━━━━━━━━"
                    )
                except Exception as e:
                    print(f"[하트비트 알림 실패] {e}")
                last_heartbeat = now

        except ccxt.RateLimitExceeded as e:
            # API 호출 한도 초과 → 30초 대기
            print(f"[Rate Limit] {e} → {RATE_LIMIT_BACKOFF}초 대기")
            time.sleep(RATE_LIMIT_BACKOFF)
            continue

        except ccxt.NetworkError as e:
            # 네트워크 오류 → 지수 백오프
            consecutive_errors += 1
            backoff = min(ERROR_BACKOFF_BASE * (2 ** (consecutive_errors - 1)), ERROR_BACKOFF_MAX)
            print(f"[네트워크 오류 #{consecutive_errors}] {e} → {backoff}초 대기")
            time.sleep(backoff)
            continue

        except Exception as e:
            # 기타 예외 → 지수 백오프 + 3회 연속 실패 시 텔레그램 알림 (스팸 방지)
            consecutive_errors += 1
            backoff = min(ERROR_BACKOFF_BASE * (2 ** (consecutive_errors - 1)), ERROR_BACKOFF_MAX)
            print(f"[메인 루프 오류 #{consecutive_errors}] {e} → {backoff}초 대기")
            try:
                if consecutive_errors == 3:
                    telegram_alert.SendMessage(
                        f"⚠️ *[봇 연속 오류 3회]*\n"
                        f"오류: `{str(e)[:200]}`\n"
                        f"{backoff}초 후 재시도"
                    )
            except Exception:
                pass
            time.sleep(backoff)
            continue

        # ─── 폴링 주기 유지 (루프 처리 시간 보정) ───
        elapsed = time.time() - loop_start
        sleep_time = max(0, POLL_INTERVAL - elapsed)
        if sleep_time > 0:
            # sleep을 0.5초씩 쪼개서 종료 신호를 빠르게 감지
            sleep_step = 0.5
            slept = 0.0
            while slept < sleep_time and not shutdown_requested:
                time.sleep(min(sleep_step, sleep_time - slept))
                slept += sleep_step

    # ─── 종료 처리 ───
    print("✅ 안전 종료 완료")
    try:
        telegram_alert.SendMessage(
            f"⏹️ *[비트겟 그리드봇 종료]*\n"
            f"━━━━━━━━━━━━━━━\n"
            f"🔁 총 폴링: `{loop_count}회`\n"
            f"🕒 `{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}`\n"
            f"━━━━━━━━━━━━━━━"
        )
    except Exception:
        pass


if __name__ == "__main__":
    main_loop()
