# -*- coding: utf-8 -*-
# =========================================================
#  비트겟 그리드 봇 — 수강생 배포용
# =========================================================
#  ※ 이 파일은 배포 버전입니다. (선생님 운영 서버본 기준)
#  ※ 이 봇은 LS증권을 쓰지 않습니다. 시세·주문 모두 비트겟에서 처리합니다.
#     따라서 LS 앱키(myLS_StockInfo.yaml)가 필요 없습니다.
#
#  [ 서버에 올릴 파일 ]
#   /var/autobot/bitget_grid 에 아래를 함께 올리세요.
#     - Grid_Bitget_배포용.py   ← 이 파일
#     - 공통모듈 폴더 안의 .py
#         · telegram_alert.py
#         · bitget_keys.py
#
#  [ 필요한 파이썬 패키지 ]
#     pip3 install ccxt pandas openpyxl tendo python-telegram-bot
#     ※ tendo 는 같은 봇이 두 번 실행되는 것을 막는 데 씁니다.
#
#  [ 사용 전 반드시 확인 ]
#   1) bitget_keys.py 에 본인 비트겟 API 키를 넣으세요.
#      비우면 봇이 시작 즉시 멈춥니다(주문 안 나감).
#         chmod 600 /var/autobot/bitget_gridbitget_keys.py
#   2) telegram_alert.py 에 본인 TOKEN / CHAT_ID 를 넣으세요.
#   3) 종목 : BTC/USDT:USDT (종목코드목록에서 변경 가능)
#   4) 차수별 매수금액 : 15 USDT (원본 100 USDT)
#      ※ BTC 최소 주문수량 0.0001개 ≈ 6.5 USDT 입니다.
#        이보다 작게 잡으면 주문이 거부됩니다.
#      ※ 차수가 최대 99개까지 있으므로 전부 채워지면
#        15 × 99 = 1,485 USDT 가 투입됩니다. 본인 규모에 맞게 조정하세요.
#   5) 이 봇은 '상주형' 입니다. 크론 1회 호출이 아니라 계속 떠 있습니다.
#         cd ~ && nohup python3 -u /var/autobot/bitget_gridGrid_Bitget_배포용.py \
#              >> /var/autobot/bitget_gridgrid_bitget.log 2>&1 &
#
#  [ ★ 하나의 계좌를 여러 봇이 함께 쓸 때 ]
#   비트겟은 계정·심볼당 포지션이 하나로 합쳐집니다(단방향).
#   이 봇은 '자기 엑셀' 을 유일한 원장으로 삼고, 매도는 reduceOnly 로
#   자기 엑셀 수량만 청산하므로 다른 봇 물량을 건드리지 않습니다.
#   ※ 포지션이 열려 있는데 엑셀을 지우면 봇이 기존 물량을 잊고
#     1차부터 새로 삽니다. 절대 지우지 마세요.
#
#  [ 면책 ]
#   레버리지 선물 + 물타기 그리드는 원금 전액 손실이 가능합니다.
#   투자 판단과 그 결과에 대한 책임은 사용자 본인에게 있습니다.
#   반드시 소액으로 충분히 검증한 뒤 사용하세요.
# =========================================================
r"""
================================================================================
 비트겟 그리드롱 자동매매 봇   |   Grid_Bitget.py
================================================================================

[무엇을 하는 봇인가]
--------------------------------------------------------------------------------
Bitget USDT 무기한선물에서 '그리드 롱' 전략을 실행합니다.

  - 1차 매수 후, 가격이 정해진 % 만큼 하락할 때마다 분할 매수 (차수 증가)
  - 각 차수는 자기 매수가 대비 목표 % 상승 시 그 차수만 개별 청산
  - 각 차수는 2단계 트레일링 매도: 1차 +0.7% 도달 후 하회 시 0.7% 보존 매도, 2차 +0.9%부터 고점 대비 -0.2% 트레일링 매도
  - 30분 내 서로 다른 차수 5개 매수 시 30분 매수 쿨링(재개 후 하이브리드: >A 즉시 / ≤A 저점 반등)
  - 3초마다 가격을 확인하는 상주 프로그램 (한 번 켜두면 24시간 실행)
  - 매수/매도 발생 시 텔레그램 알림, 모든 거래는 엑셀 파일에 기록

--------------------------------------------------------------------------------
[파일 위치]   ※ 모두 /var/autobot/bitget_grid 폴더 안에 있습니다
--------------------------------------------------------------------------------
  봇 파일       /var/autobot/bitget_grid/Grid_Bitget.py
  거래 기록     /var/autobot/bitget_grid/BTC-USDT_비트겟그리드_LONG폴.xlsx   (종목별로 자동 생성)
  에러 로그     /var/autobot/bitget_grid/bitget-grid-error.log
  청산 락 파일  /var/autobot/bitget_grid/.BTC_USDT_USDT_total_profit_target_reached.lock
  수익 baseline /var/autobot/bitget_grid/.SOXL_USDT_USDT_total_profit_baseline.json
  매수 쿨링 상태 /var/autobot/bitget_grid/.SOXL_USDT_USDT_buy_cooldown.json

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
  각 차수는 1차 +0.7% 보존 트리거(하회 시 0.7% 매도), 2차 +0.9%부터 고점 대비 -0.2% 트레일링 청산.

  예) 1차를 100 USDT 매수 → 하락하여 2차 100 USDT 매수 (평단가 내려감)
      2차 수익률이 +0.8% 찍고 내려오면 0.7%에서 보존 매도, +1.3%까지 올랐다 밀리면 1.1%에서 청산

  급락 시(서로 다른 차수 5개 매수) 30분 매수 금지 후, 재개 시점 가격 A를 기준으로
  현재가 > A 이면 즉시 1차수, ≤ A 이면 저점 대비 매수간격만큼 반등 시 1차수만 이어갑니다.
  간격(-0.7%)과 금액은 GRID_TEMPLATE, 트레일/쿨링 수치는 하단 상수에서 바꿉니다.

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
  판정: (미실현 + 엑셀 누적실현) - baseline ≥ 목표 USDT 이면 전 차수 청산.
  RESTART 시 청산 직후 누적실현을 새 baseline으로 저장 → 다음 사이클은 +목표만큼만 다시 모읍니다.

  청산 후 동작은 TOTAL_PROFIT_AFTER_LIQUIDATION 으로 정합니다.

      "STOP"      청산 후 봇 종료          + 락 파일 생성
      "PAUSE"     봇은 살아있고 매수만 정지  + 락 파일 생성
      "RESTART"   즉시 1차부터 다시 시작    + baseline 갱신 (락 없음)

  락 파일이 있으면 봇을 다시 켜도 매수하지 않습니다. 새로 시작하려면
  아래 두 가지를 순서대로 하세요.

      1) 엑셀을 다른 이름으로 옮기기 (누적 수익이 0부터 다시 쌓이도록)
         mv /var/autobot/bitget_grid/BTC-USDT_비트겟그리드_LONG폴.xlsx /var/autobot/bitget_grid/백업_1차.xlsx

      2) 락 파일 삭제
         rm /var/autobot/bitget_grid/.*_total_profit_target_reached.lock

  ※ RESTART는 baseline으로 사이클을 리셋하므로, 엑셀을 안 옮겨도
     같은 누적실현으로 즉시 재청산되지 않습니다.
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
  - 매수가 안 들어감         → 락 파일 / 매수 쿨링 / 선물 지갑 잔고 / 최소 주문 수량 확인
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
import json
from datetime import datetime, timedelta, timezone
import copy
from typing import Dict, Any, Tuple, Optional

import ccxt
import pandas as pd
try:
    from tendo import singleton
except ImportError:
    singleton = None

import telegram_alert  # 사용자 정의 모듈: SendMessage(text: str) 제공

# 서버 OS 시간대(UTC 등)와 무관하게 엑셀·쿨링·텔레그램은 서울 시각을 쓴다.
# 일일 정산(_kst_now)과 거래 시각이 어긋나면 정산 구간이 9시간 밀린다.
KST = timezone(timedelta(hours=9))

def 지금() -> datetime:
    """서울(KST) 기준 naive datetime."""
    return datetime.now(KST).replace(tzinfo=None)


# ===============================
# 싱글톤: 봇이 두 개 실행되는 것을 막습니다
# ===============================
# 실수로 실행 명령을 두 번 입력해도, 이미 돌고 있으면 새로 뜬 쪽이 바로 종료됩니다.
# 봇이 두 개 돌면 같은 종목을 중복 매매하게 되므로 반드시 필요한 안전장치입니다.
if singleton is not None:
    try:
        me = singleton.SingleInstance()
    except singleton.SingleInstanceException:
        print("[싱글톤] 이미 실행 중 → 이번 실행은 즉시 종료합니다.")
        sys.exit(0)


# ===============================
# ⚠️ API 키 (실제 값으로 교체)
# ===============================
#   키는 bitget_keys.py 에서 읽는다. 이 파일에는 키를 두지 않는다.
#   bitget_keys.py 를 이 파일과 같은 폴더(/var/autobot/bitget_grid)에 두세요.
try:
    from bitget_keys import BITGET_API_KEY, BITGET_SECRET, BITGET_PASSPHRASE
except ImportError:
    BITGET_API_KEY = "여기에_API_KEY"
    BITGET_SECRET = "여기에_SECRET"
    BITGET_PASSPHRASE = "여기에_PASSPHRASE"

access = BITGET_API_KEY.strip()
secret = BITGET_SECRET.strip()
password = BITGET_PASSPHRASE.strip()

if not access or access.startswith('여기에_'):
    raise SystemExit('비트겟 API 키가 설정되지 않았습니다. bitget_keys.py 를 확인하세요.')


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
# - 목표 사이클 수익금 (USDT 단위)
# - 판정: (미실현 + 엑셀 누적실현) - baseline ≥ 이 값 이면 전체 청산
# - RESTART 후 baseline이 갱신되므로, 매 사이클마다 목표만큼 새로 모아야 함
TOTAL_PROFIT_TARGET_USDT = 150.0

# 📌 TOTAL_PROFIT_AFTER_LIQUIDATION = "STOP"
# - 전체 청산 후 봇 동작 모드 (3가지 중 하나):
#   "STOP"    : 봇 자체 종료 (가장 안전, 추천)
#               + 락 파일 유지 → 크론탭이 봇 부활시켜도 매수 안 함
#               + 완전 정지 원하면 크론탭도 주석 처리하는 게 좋음
#   "PAUSE"   : 봇은 계속 살아있되 매매만 정지 (관망 모드)
#               + 가격 출력 + 텔레그램 하트비트는 계속
#               + 락 파일 수동 해제 시 매매 재개
#   "RESTART" : 청산 직후 baseline=청산후 누적실현 저장 → 다음 폴링부터 1차 재시작
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
# - 출력 예: 💰 [총수익] 사이클=+4.68 (미실현=+1.23 + 누적=+53.45 - baseline=+50.00) (목표 50.0)
# - 0 으로 두면 출력 안 함
TOTAL_PROFIT_LOG_INTERVAL = 20


# --- 3-State Switching Engine (5일선 추세 라이딩) 설정 ⭐ ---
# 📌 TREND_RIDING_ENABLE = True
# - True : 5일선 상승 시 매도 정지(포지션 100% 홀딩) & 추세 꺾임 시 총수익 청산(Flash Close) 또는 매도 재개
# - False: 기능 비활성화 → 기존 그리드 매매만 수행
TREND_RIDING_ENABLE = True

# 📌 이동평균 계산 기준 캔들 및 기간
# - 1차 검증: 일봉(1d) 5일선
# - 2차 검증: 4시간봉(4h) 5선 (1차 이탈 시 즉시 조회 크로스체크)
MA_TIMEFRAME_1D = "1d"
MA_TIMEFRAME_4H = "4h"
MA_TIMEFRAME = MA_TIMEFRAME_1D  # 기존 코드 호환용
MA_PERIOD = 5

# 📌 추세 종료(지지선 이탈) 판정 버퍼 (%)
# - 0.5% 이탈 기준: 현재가 < 5MA * (1 - 0.005) = 5MA * 0.995
TREND_EXIT_BUFFER_PCT = 0.5

# 📌 fetch_ohlcv 거래소 호출 캐싱 주기 (초 단위, API Rate Limit 방어)
MA_CACHE_SECONDS = 30       # 일봉 캐시 30초
MA_4H_CACHE_SECONDS = 15    # 4시간봉 캐시 15초

# 📌 TREND_RIDING_SELL_MODE = "STAGE1_ONLY"
# - 추세 라이딩(상태 2) 중 매도 허용 범위를 결정합니다.
# - "STAGE1_ONLY" (권장):
#     우상향 중인 차수는 100% 홀딩하여 추세를 끝까지 타되,
#     수익률이 +TRAIL_STAGE1_PCT(0.7%) 이상 도달 후 그 이하로 반락하는 차수만
#     '이익 보존 매도'를 실행하여 수익 반납 원천 차단.
# - "ALL":
#     1단계(0.7% 보존선) 및 2단계(0.9% 이상 도달 후 고점 대비 -0.2%p 하락) 모두 매도 허용.
# - "OFF":
#     기존 방식 (5MA 지지선 이탈 전까지 매도 100% 전면 금지, 무조건 홀딩).
TREND_RIDING_SELL_MODE = "STAGE1_ONLY"  # "STAGE1_ONLY" | "ALL" | "OFF"


# --- 트레일링 매도 설정 (2단계 하이브리드) ---
# 📌 TRAILING_ENABLE = True
# - True : 2단계 하이브리드 트레일링 스탑 적용
#          1단계: +TRAIL_STAGE1_PCT 도달 후 하회 시 최소 이익 보존 매도
#          2단계: +TRAIL_STAGE2_PCT 도달 후 고점 대비 -TRAIL_DISTANCE_PCT 하락 시 매도
# - False: GRID_TEMPLATE 매도간격 고정 목표가 매도 (기존 방식)
TRAILING_ENABLE = True

# 📌 TRAIL_STAGE1_PCT = 0.7
# - 1차 트리거(최소 이익 보존선, %): +0.7% 도달 후 0.7% 이하로 내려가면 익절
TRAIL_STAGE1_PCT = 0.7

# 📌 TRAIL_STAGE2_PCT = 0.9
# - 2차 트리거(다이내믹 트레일링 시작 기준, %): +0.9% 도달 시 고점 추적 시작
TRAIL_STAGE2_PCT = 0.9
TRAIL_ACTIVATE_PCT = TRAIL_STAGE2_PCT  # 하위 호환용 별칭

# 📌 TRAIL_DISTANCE_PCT = 0.2
# - 고점 대비 하락폭(%) — 2단계 진입 후 이만큼 빠지면 해당 차수 청산 (예: 고점 1.3% → 매도선 1.1%)
TRAIL_DISTANCE_PCT = 0.2


# --- 매수 쿨링타임 설정 ---
# 📌 BUY_COOLDOWN_ENABLE = True
# - True : 짧은 시간에 다수 차수 매수 시 매수 일시 정지
#          + 재개 후 하이브리드 rebase
#            · 현재가 > A(해제시 가격) → 즉시 다음 차수 1회
#            · 현재가 ≤ A → 저점 트레일, 저점 대비 |매수간격|% 반등 시 1회
BUY_COOLDOWN_ENABLE = True

# 📌 BUY_COOLDOWN_WINDOW_SEC = 600
# - 매수 횟수 집계 구간 (초) — 10분
BUY_COOLDOWN_WINDOW_SEC = 2400

# 📌 BUY_COOLDOWN_BUY_COUNT = 5
# - 위 구간 내 '서로 다른 차수' 매수가 이 값 이상이면 쿨링 발동
# - 동일 차수 반복 매수는 1회로만 집계
BUY_COOLDOWN_BUY_COUNT = 5

# 📌 BUY_COOLDOWN_PAUSE_SEC = 1800
# - 쿨링 동안 매수 금지 시간 (초) — 30분 (매도는 계속)
BUY_COOLDOWN_PAUSE_SEC = 3600


# --- 쿨링 공백 복원 롤다운 그리드 설정 (v2.4 비상 자가치유 롤다운 탑재) ---
# 📌 GAP_RECONSTRUCTION_ENABLE = True
# - True : 쿨링 해제 후 Rebase 첫 매수 체결(또는 증거금 고갈 시 비상 선손절) 시,
#          직전 차수와의 공백을 계산하여 상위 차수 강제 손절 및 공백 예약 매수 그물망 배치
# - False: 기능 비활성화 (기존 하이브리드 Rebase 방식만 유지)
GAP_RECONSTRUCTION_ENABLE = True

# 📌 MAX_GAP_CUT_LEVELS = 10
# - 1회 공백 복원 시 최대 손절 차수 캡 제한 (기본값: 10개)
# - 아무리 공백이 커도 1회에 최대 10개까지만 손절하여 계좌 손실 충격 완화
MAX_GAP_CUT_LEVELS = 10


# --- 더미(가상) 매수/매도 알림 모드 스위치 ---
# 📌 DUMMY_SELL_ALERT_MODE = "EACH"
# - "EACH" : 더미 차수별 목표가 도달 시마다 가상 매수 & 가상 매도 발송 (검증/테스트용)
# - "OFF"  : 더미 매도 알림 끔 (조용히 침묵하며 실제 포지션 청산만 대기, 안정 운영용)
# - "ONCE" : 최저 바닥 차수 첫 반등 시 1회만 요약 알림 발송
DUMMY_SELL_ALERT_MODE = "EACH"


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
    1: {"매수간격": 0.0, "매도간격": 0.7, "매수금액": 143},
    2: {"매수간격": -0.7, "매도간격": 0.7, "매수금액": 143},
    3: {"매수간격": -0.7, "매도간격": 0.7, "매수금액": 143},
    4: {"매수간격": -0.7, "매도간격": 0.7, "매수금액": 143},
    5: {"매수간격": -0.7, "매도간격": 0.7, "매수금액": 143},
    6: {"매수간격": -0.7, "매도간격": 0.7, "매수금액": 143},
    7: {"매수간격": -0.7, "매도간격": 0.7, "매수금액": 143},
    8: {"매수간격": -0.7, "매도간격": 0.7, "매수금액": 143},
    9: {"매수간격": -0.7, "매도간격": 0.7, "매수금액": 143},
    10: {"매수간격": -0.7, "매도간격": 0.7, "매수금액": 143},
    **{i: {"매수간격": -0.7, "매도간격": 0.7, "매수금액": 143} for i in range(11, 26)},
    **{i: {"매수간격": -0.7, "매도간격": 0.7, "매수금액": 143} for i in range(26, 46)},
    **{i: {"매수간격": -0.7, "매도간격": 0.7, "매수금액": 143} for i in range(46, 71)},
    **{i: {"매수간격": -0.7, "매도간격": 0.7, "매수금액": 143} for i in range(71, 91)},
    **{i: {"매수간격": -0.7, "매도간격": 0.7, "매수금액": 143} for i in range(91, 101)},
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
    "최고수익률",
    "매도날짜", "매도가", "매도수량", "매도금액",
    "매도시현재가", "매도슬리피지%",
    "수수료", "실현손익", "수익률", "누적실현손익",
    "청산유형",
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

    # pandas to_excel() 기본 동작은 워크북 전체를 새로 만들어
    # '일일정산' 시트를 지워버린다. 거래 시트만 교체하고 나머지는 유지.
    if os.path.exists(file_name):
        try:
            with pd.ExcelWriter(
                file_name, engine="openpyxl", mode="a", if_sheet_exists="replace"
            ) as writer:
                names = list(writer.book.sheetnames)
                trade_sheet = "Sheet1"
                if names and names[0] != "일일정산":
                    trade_sheet = names[0]
                elif "Sheet1" in names:
                    trade_sheet = "Sheet1"
                df.to_excel(writer, index=False, sheet_name=trade_sheet)
            return
        except Exception as e:
            print(f"[엑셀 시트 유지 저장 실패 → 전체 덮어쓰기] {e}")
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
            os.makedirs("/var/autobot/bitget_grid", exist_ok=True)
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
        os.makedirs("/var/autobot/bitget_grid", exist_ok=True)
        with open(_총수익청산락_경로(Ticker), "w") as f:
            f.write(f"locked at {지금().strftime('%Y-%m-%d %H:%M:%S')}\n{info}")
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


# --- 총 수익 사이클 baseline (RESTART 무한 재청산 방지) ---
# 사이클 수익 = (미실현 + 누적실현) - baseline
# RESTART/부트스트랩 시 baseline 갱신 + cycle_count(+1)
# cycle_count: 현재 진행 중인 회차 (1부터, 파일 없으면 1)
def _총수익_baseline_경로(Ticker: str) -> str:
    safe = Ticker.replace('/', '_').replace(':', '_')
    return f"/var/autobot/bitget_grid/.{safe}_total_profit_baseline.json"

def 총수익_baseline_상태조회(Ticker: str) -> Tuple[float, int]:
    """[반환] (baseline, cycle_count) — 파일 없으면 (0.0, 1)"""
    path = _총수익_baseline_경로(Ticker)
    if not os.path.exists(path):
        return 0.0, 1
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        baseline = float(data.get("baseline", 0.0))
        raw_c = data.get("cycle_count", 1)
        try:
            cycle_count = int(raw_c)
        except (TypeError, ValueError):
            cycle_count = 1
        if cycle_count < 1:
            cycle_count = 1
        return baseline, cycle_count
    except Exception as e:
        print(f"[총수익 baseline 로드 실패] {e}")
        return 0.0, 1

def 총수익_baseline_조회(Ticker: str) -> float:
    baseline, _ = 총수익_baseline_상태조회(Ticker)
    return baseline

def 총수익_사이클회차_조회(Ticker: str) -> int:
    _, cycle_count = 총수익_baseline_상태조회(Ticker)
    return cycle_count

def 총수익_baseline_저장(
    Ticker: str,
    baseline: float,
    cycle_count: Optional[int] = None,
):
    """cycle_count를 넘기면 그대로 저장, None이면 기존 회차 유지."""
    try:
        _, prev_c = 총수익_baseline_상태조회(Ticker)
        if cycle_count is None:
            cycle_count = prev_c
        cycle_count = max(1, int(cycle_count))
        os.makedirs("/var/autobot/bitget_grid", exist_ok=True)
        with open(_총수익_baseline_경로(Ticker), "w", encoding="utf-8") as f:
            json.dump(
                {
                    "baseline": float(baseline),
                    "cycle_count": cycle_count,
                    "updated_at": 지금().strftime("%Y-%m-%d %H:%M:%S"),
                },
                f,
                ensure_ascii=False,
                indent=2,
            )
        print(
            f"📌 [baseline] {Ticker} → {float(baseline):+.4f} USDT | "
            f"사이클 {cycle_count}회차",
            flush=True,
        )
    except Exception as e:
        print(f"[총수익 baseline 저장 실패] {e}")

def 총수익_baseline_부트스트랩(Ticker: str):
    """
    시작 시 사이클 수익이 이미 목표 이상이면 baseline을 현재 누적으로 맞추고 회차 +1.
    (구버전 RESTART 무한 청산 상태 탈출용)
    """
    if not TOTAL_PROFIT_TARGET_ENABLE:
        return
    누적 = 누적실현손익_조회(Ticker)
    baseline, cycle_count = 총수익_baseline_상태조회(Ticker)
    사이클 = 누적 - baseline
    if 사이클 >= TOTAL_PROFIT_TARGET_USDT:
        다음회차 = cycle_count + 1
        총수익_baseline_저장(Ticker, 누적, cycle_count=다음회차)
        print(
            f"📌 [baseline 부트스트랩] {Ticker} 사이클수익 {사이클:+.4f} ≥ 목표 "
            f"→ baseline {baseline:+.4f} → {누적:+.4f} | "
            f"{cycle_count}회차 → {다음회차}회차",
            flush=True,
        )


# ===============================
# 매수 쿨링타임 상태 (JSON 파일)
# ===============================
# 필드:
#   cooldown_until (ISO|null)
#   rebase_anchor  (float|null) — 쿨링 해제 시점 가격 A (고정)
#   trail_low      (float|null) — 해제 후 ≤A 구간에서 추적한 저점
#   activated_at / pending_rebase
def _쿨링상태_경로(Ticker: str) -> str:
    safe = Ticker.replace('/', '_').replace(':', '_')
    base_dir = "/var/autobot/bitget_grid" if os.path.exists("/var/autobot/bitget_grid") else "."
    return os.path.join(base_dir, f".{safe}_buy_cooldown.json")

def 쿨링상태로드(Ticker: str) -> Dict[str, Any]:
    path = _쿨링상태_경로(Ticker)
    default: Dict[str, Any] = {
        "cooldown_until": None,
        "rebase_anchor": None,
        "trail_low": None,
        "activated_at": None,
        "pending_rebase": False,
    }
    if not os.path.exists(path):
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        for k, v in default.items():
            if k not in data:
                data[k] = v
        return data
    except Exception as e:
        print(f"[쿨링상태 로드 실패] {e}")
        return default

def 쿨링상태저장(Ticker: str, state: Dict[str, Any]):
    try:
        path = _쿨링상태_경로(Ticker)
        base_dir = os.path.dirname(path)
        if base_dir and not os.path.exists(base_dir):
            os.makedirs(base_dir, exist_ok=True)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
    except Exception as e:
        print(f"[쿨링상태 저장 실패] {e}")

def 쿨링여부(Ticker: str) -> bool:
    """쿨링 중이면 True (매수 차단). 만료됐으면 pending_rebase 표시 후 False."""
    if not BUY_COOLDOWN_ENABLE:
        return False
    state = 쿨링상태로드(Ticker)
    until_s = state.get("cooldown_until")
    if not until_s:
        return False
    try:
        until = datetime.fromisoformat(until_s)
        if until.tzinfo is not None:
            until = until.astimezone(KST).replace(tzinfo=None)
    except Exception:
        return False
    if 지금() < until:
        return True
    # 쿨링 만료 → 재개 시 rebase를 위해 플래그 남김
    if not state.get("pending_rebase") and state.get("rebase_anchor") is None:
        state["cooldown_until"] = None
        state["pending_rebase"] = True
        쿨링상태저장(Ticker, state)
    elif state.get("cooldown_until"):
        state["cooldown_until"] = None
        쿨링상태저장(Ticker, state)
    return False

def 쿨링시작(Ticker: str):
    """30분 매수 쿨링 시작"""
    now = 지금()
    state = 쿨링상태로드(Ticker)
    state["cooldown_until"] = (now + timedelta(seconds=BUY_COOLDOWN_PAUSE_SEC)).isoformat()
    state["activated_at"] = now.isoformat()
    state["rebase_anchor"] = None
    state["trail_low"] = None
    state["pending_rebase"] = False
    쿨링상태저장(Ticker, state)

def 쿨링_rebase_적용(Ticker: str, 현재가: float) -> bool:
    """
    쿨링 만료 후 첫 매수 폴링에서 해제 기준가 A와 저점(trail_low)을 설정.
    A는 고정, trail_low는 ≤A 구간에서만 하향 갱신.
    적용했으면 True (텔레그램 알림용).
    """
    if not BUY_COOLDOWN_ENABLE or 현재가 <= 0:
        return False
    state = 쿨링상태로드(Ticker)
    if not state.get("pending_rebase"):
        return False
    if state.get("rebase_anchor") is not None:
        state["pending_rebase"] = False
        쿨링상태저장(Ticker, state)
        return False
    px = float(현재가)
    state["rebase_anchor"] = px  # A (고정)
    state["trail_low"] = px
    state["pending_rebase"] = False
    state["cooldown_until"] = None
    쿨링상태저장(Ticker, state)
    return True

def 쿨링_rebase_저점갱신(Ticker: str, 현재가: float) -> bool:
    """
    쿨링 해제 후 현재가 ≤ A 일 때 trail_low를 저점 트레일(하향만).
    갱신 시 True.
    """
    if not BUY_COOLDOWN_ENABLE or 현재가 <= 0:
        return False
    state = 쿨링상태로드(Ticker)
    anchor = state.get("rebase_anchor")
    if anchor is None:
        return False
    try:
        A = float(anchor)
    except (TypeError, ValueError):
        return False
    px = float(현재가)
    if px > A:
        return False
    try:
        low = float(state["trail_low"]) if state.get("trail_low") is not None else A
    except (TypeError, ValueError):
        low = A
    if px < low:
        state["trail_low"] = px
        쿨링상태저장(Ticker, state)
        return True
    if state.get("trail_low") is None:
        state["trail_low"] = low
        쿨링상태저장(Ticker, state)
    return False

def 쿨링_rebase_앵커조회(Ticker: str) -> Optional[float]:
    """쿨링 해제 시점 기준가 A 조회."""
    state = 쿨링상태로드(Ticker)
    anchor = state.get("rebase_anchor")
    if anchor is None:
        return None
    try:
        return float(anchor)
    except (TypeError, ValueError):
        return None

def 쿨링_rebase_저점조회(Ticker: str) -> Optional[float]:
    """저점 트레일 값. 없으면 A로 대체."""
    state = 쿨링상태로드(Ticker)
    A = state.get("rebase_anchor")
    if A is None:
        return None
    try:
        A_f = float(A)
    except (TypeError, ValueError):
        return None
    low = state.get("trail_low")
    if low is None:
        return A_f
    try:
        return float(low)
    except (TypeError, ValueError):
        return A_f

def 쿨링_rebase_클리어(Ticker: str):
    state = 쿨링상태로드(Ticker)
    state["rebase_anchor"] = None
    state["trail_low"] = None
    state["pending_rebase"] = False
    쿨링상태저장(Ticker, state)


# =========================================================
# 📌 쿨링 공백 복원 롤다운 그리드 모듈 (v2.1)
# =========================================================

def _gap_state_path(Ticker: str) -> str:
    """공백 복원 예약 주문 상태 파일 경로"""
    safe = Ticker.replace("/", "_").replace(":", "_")
    base_dir = "/var/autobot/bitget_grid" if os.path.exists("/var/autobot/bitget_grid") else "."
    return os.path.join(base_dir, f".{safe}_gap_orders.json")

def 공백복원_상태로드(Ticker: str) -> Dict[str, Any]:
    """공백 예약 주문 상태 로드"""
    path = _gap_state_path(Ticker)
    if not os.path.exists(path):
        return {"orders": [], "last_rebalance": None}
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"[공백복원 상태 로드 실패] {e}")
        return {"orders": [], "last_rebalance": None}

def 공백복원_상태저장(Ticker: str, state: Dict[str, Any]):
    """공백 예약 주문 상태 원자적(Atomic) 영속화 저장"""
    path = _gap_state_path(Ticker)
    try:
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
    except Exception as e:
        print(f"[공백복원 상태 저장 실패] {e}")

def 공백복원_상태클리어(Ticker: str):
    """공백 예약 주문 상태 초기화 (총수익 청산 / 사이클 리셋 시)"""
    공백복원_상태저장(Ticker, {"orders": [], "last_rebalance": None})
    잔고부족_트래커_클리어(Ticker)


# =========================================================
# 📌 잔고 부족 가상 매수 시도 트래커 (v2.4 비상 자가치유 롤다운 모듈)
# =========================================================

def _insufficient_funds_state_path(Ticker: str) -> str:
    """잔고 부족 매수 시도 상태 파일 경로"""
    safe = Ticker.replace('/', '_').replace(':', '_')
    base_dir = "/var/autobot/bitget_grid" if os.path.exists("/var/autobot/bitget_grid") else "."
    return os.path.join(base_dir, f".{safe}_insufficient_funds.json")

def 잔고부족_상태로드(Ticker: str) -> Dict[str, Any]:
    """잔고 부족 매수 시도 상태 로드"""
    path = _insufficient_funds_state_path(Ticker)
    default: Dict[str, Any] = {"attempts": []}
    if not os.path.exists(path):
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict) or "attempts" not in data:
            return default
        return data
    except Exception as e:
        print(f"[잔고부족 상태 로드 실패] {e}")
        return default

def 잔고부족_상태저장(Ticker: str, state: Dict[str, Any]) -> None:
    """잔고 부족 상태 원자적(Atomic) 저장"""
    try:
        path = _insufficient_funds_state_path(Ticker)
        base_dir = os.path.dirname(path)
        if base_dir and not os.path.exists(base_dir):
            os.makedirs(base_dir, exist_ok=True)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
    except Exception as e:
        print(f"[잔고부족 상태 저장 실패] {e}")

def 잔고부족_트래커_클리어(Ticker: str) -> None:
    """정상 매수 체결 또는 롤다운 완료 시 잔고부족 트래커 초기화"""
    잔고부족_상태저장(Ticker, {"attempts": []})

def 잔고부족_시도기록(Ticker: str, 차수: int, 현재가: float) -> None:
    """
    증거금 부족(InsufficientFunds)으로 매수 실패 시 차수, 시세, 타임스탬프 기록.
    엑셀 원장을 오염시키지 않고 독립 상태 파일에만 누적.
    동일 차수가 이미 존재할 경우 중복 누적(스팸 원인)을 방지하고 최신 타임스탬프 갱신(Upsert).
    """
    state = 잔고부족_상태로드(Ticker)
    now_str = 지금().isoformat()
    attempts = state.get("attempts", [])
    cutoff = 지금() - timedelta(hours=24)

    # 24시간 이내 유효한 기록 필터링 및 기존 동일 차수 정제 (중복 누적 방지)
    by_level: Dict[int, Dict[str, Any]] = {}
    for a in attempts:
        try:
            lvl = int(a.get("차수", 0))
            if lvl <= 0:
                continue
            ts = datetime.fromisoformat(a.get("timestamp", ""))
            if ts.tzinfo is not None:
                ts = ts.astimezone(KST).replace(tzinfo=None)
            if ts >= cutoff:
                if lvl not in by_level or ts > by_level[lvl]["ts_obj"]:
                    by_level[lvl] = {
                        "entry": a,
                        "ts_obj": ts
                    }
        except Exception:
            continue

    target_lvl = int(차수)
    if target_lvl in by_level:
        # 이미 존재하는 차수: 최초 진입가는 유지하고 타임스탬프만 최신으로 갱신 (쿨다운 판별 유지)
        by_level[target_lvl]["entry"]["timestamp"] = now_str
    else:
        # 신규 차수 추가
        by_level[target_lvl] = {
            "entry": {
                "차수": target_lvl,
                "price": float(현재가),
                "timestamp": now_str
            },
            "ts_obj": 지금()
        }

    # 차수 오름차순으로 정렬하여 저장
    sorted_attempts = [v["entry"] for k, v in sorted(by_level.items(), key=lambda item: item[0])]
    state["attempts"] = sorted_attempts
    잔고부족_상태저장(Ticker, state)


def 잔고부족_최근시도횟수(Ticker: str, window_sec: int = 2400) -> int:
    """
    최근 window_sec (기본 2400초 = 40분) 이내에 기록된 고유 차수(unique level) 개수 반환.
    급락형 파동 판별: 40분 내 서로 다른 차수 5회 이상 발생 시에만 쿨링 발동.
    """
    state = 잔고부족_상태로드(Ticker)
    attempts = state.get("attempts", [])
    if not attempts:
        return 0
    cutoff = 지금() - timedelta(seconds=window_sec)
    recent_levels = set()
    for a in attempts:
        try:
            ts = datetime.fromisoformat(a.get("timestamp", ""))
            if ts.tzinfo is not None:
                ts = ts.astimezone(KST).replace(tzinfo=None)
            if ts >= cutoff:
                recent_levels.add(int(a["차수"]))
        except Exception:
            continue
    return len(recent_levels)

def 잔고부족_가상차수_조회(Ticker: str, 차수: int) -> Optional[Dict[str, Any]]:
    """
    df에 기록되지 못한 직전 가상 차수(차수 - 1 등)의 시도 기록 조회.
    최근 6시간 이내 기록된 유효한 시도 중 가장 최신 항목 반환.
    """
    state = 잔고부족_상태로드(Ticker)
    attempts = state.get("attempts", [])
    cutoff = 지금() - timedelta(hours=6)
    matched = []
    for a in attempts:
        if int(a.get("차수", 0)) == int(차수):
            try:
                ts = datetime.fromisoformat(a.get("timestamp", ""))
                if ts.tzinfo is not None:
                    ts = ts.astimezone(KST).replace(tzinfo=None)
                if ts >= cutoff:
                    matched.append(a)
            except Exception:
                continue
    if matched:
        return matched[-1]
    return None

def 잔고부족_최근기록_여부(Ticker: str, 차수: int, cooldown_sec: int = 60) -> bool:
    """해당 차수가 최근 cooldown_sec 초 이내에 이미 시도된 적이 있는지 여부 (거래소 API 반복 호출 폭주 방지)"""
    v = 잔고부족_가상차수_조회(Ticker, 차수)
    if not v:
        return False
    try:
        ts = datetime.fromisoformat(v.get("timestamp", ""))
        if ts.tzinfo is not None:
            ts = ts.astimezone(KST).replace(tzinfo=None)
        if (지금() - ts).total_seconds() < cooldown_sec:
            return True
    except Exception:
        pass
    return False

def 최근_총매수시도_고유차수수(df: pd.DataFrame, Ticker: str, window_sec: int = 2400) -> int:
    """
    최근 window_sec 초 이내에 발생한 '실제 매수 체결 차수'와 '잔고 부족 가상 매수 시도 차수'의
    고유(unique) 차수 합집합 개수 계산.
    급락형 파동(40분 내 5개 차수 급락) 감지용.
    """
    unique_levels = set()
    cutoff = 지금() - timedelta(seconds=window_sec)

    # 1) df에서 최근 체결된 차수
    if df is not None and not df.empty and "매수날짜" in df.columns and "차수" in df.columns:
        dates = pd.to_datetime(df["매수날짜"], errors="coerce")
        recent = df.loc[dates >= cutoff]
        if not recent.empty:
            for lvl in recent["차수"].dropna():
                try:
                    unique_levels.add(int(lvl))
                except Exception:
                    pass

    # 2) 잔고부족 가상 시도 차수
    state = 잔고부족_상태로드(Ticker)
    for a in state.get("attempts", []):
        try:
            ts = datetime.fromisoformat(a.get("timestamp", ""))
            if ts.tzinfo is not None:
                ts = ts.astimezone(KST).replace(tzinfo=None)
            if ts >= cutoff:
                unique_levels.add(int(a["차수"]))
        except Exception:
            pass

    return len(unique_levels)

def 더미매수_텔레그램알림(Ticker: str, 차수: int, 현재가: float, config: Dict[str, Any], 총시도수: int) -> None:
    """
    증거금 부족으로 실제 매수가 불가능하여 더미 매수로 기록될 때,
    차수당 최초 1회 발송되는 텔레그램 안내 메시지.
    """
    try:
        title_symbol = Ticker.split('/')[0]
        필요금액 = float(config.get("매수금액", 0.0))
        간격 = config.get("매수간격", "")
        간격_str = f"직전가 대비 {간격}% 하락 도달" if 간격 else "매수 조건 충족"
        msg = (
            f"⚠️ *[{title_symbol}_그리드롱 가상 매수 진입 (더미 기록)]* ⚠️\n"
            f"━━━━━━━━━━━━━━━\n"
            f"📊 *종목*: `{Ticker}`\n"
            f"🔢 *가상 차수*: `{차수}차` (더미 기록)\n"
            f"💵 *진입 시세*: `{현재가:.6f} USDT`\n"
            f"💰 *필요 금액*: `{필요금액:.2f} USDT` (가용 잔고 부족)\n"
            f"📉 *하락 조건*: `{간격_str}`\n"
            f"🔢 *최근 {BUY_COOLDOWN_WINDOW_SEC // 60}분 누적*: `{총시도수} / {BUY_COOLDOWN_BUY_COUNT}개` (5개 도달 시 쿨링)\n"
            f"🕒 *시간*: `{지금().strftime('%Y-%m-%d %H:%M:%S')}`\n"
            f"━━━━━━━━━━━━━━━"
        )
        telegram_alert.SendMessage(msg)
        print(f"📢 [가상 매수 알림 전송] {Ticker} {차수}차 더미 기록 텔레그램 발송 완료", flush=True)
    except Exception as e:
        print(f"⚠️ [가상 매수 텔레그램 발송 실패] {e}", flush=True)


def 더미매도_텔레그램알림(Ticker: str, 차수: int, entry_price: float, 현재가: float, config: Dict[str, Any]) -> None:
    """
    증거금 부족으로 가상 기록된 더미 매수 차수가 반등하여 익절 목표가에 도달했을 때,
    차수당 발송되는 텔레그램 안내 메시지 (가상 매도/청산).
    """
    try:
        title_symbol = Ticker.split('/')[0]
        매수금액 = float(config.get("매수금액", 0.0))
        수익률 = (현재가 / entry_price - 1.0) * 100.0 if entry_price > 0 else 0.0
        가상손익 = 매수금액 * (수익률 / 100.0)
        msg = (
            f"📈 *[{title_symbol}_그리드롱 가상 매도 (더미 청산)]* 📈\n"
            f"━━━━━━━━━━━━━━━\n"
            f"📊 *종목*: `{Ticker}`\n"
            f"🔢 *가상 차수*: `{차수}차` (더미 청산)\n"
            f"💵 *진입 시세*: `{entry_price:.6f} USDT`\n"
            f"🎯 *청산 시세*: `{현재가:.6f} USDT`\n"
            f"✨ *가상 수익률*: `+{수익률:.2f}%` (`+{가상손익:+.4f} USDT`)\n"
            f"💡 *안내*: 바닥 반등 회복 진행 중 (실제 잔고 보유분 회복 추적)\n"
            f"🕒 *시간*: `{지금().strftime('%Y-%m-%d %H:%M:%S')}`\n"
            f"━━━━━━━━━━━━━━━"
        )
        telegram_alert.SendMessage(msg)
        print(f"📢 [가상 매도 알림 전송] {Ticker} {차수}차 더미 청산 텔레그램 발송 완료", flush=True)
    except Exception as e:
        print(f"⚠️ [가상 매도 텔레그램 발송 실패] {e}", flush=True)


def 더미_가상매도_감시_및_실행(Ticker: str, 현재가: float) -> None:
    """
    더미 매수 기록된 가상 차수들의 반등 여부를 감시하고,
    목표가 도달 시 가상 매도 텔레그램 발송 및 상태 파일 원자적 갱신(청산 차수 제거).
    DUMMY_SELL_ALERT_MODE 설정에 따라 작동 ("EACH", "ONCE", "OFF").
    """
    if DUMMY_SELL_ALERT_MODE == "OFF" or 현재가 <= 0:
        return

    try:
        state = 잔고부족_상태로드(Ticker)
        attempts = state.get("attempts", [])
        if not attempts:
            return

        grid_cfg = trade_config.get(Ticker, {}).get("grid_config", {})
        cleared_levels = set()
        new_attempts = []

        # 각 더미 차수의 목표가 도달 여부 확인
        for a in attempts:
            try:
                차수 = int(a.get("차수", 0))
                entry_price = float(a.get("price", 0.0))
            except Exception:
                continue

            if 차수 <= 0 or entry_price <= 0:
                continue

            config = grid_cfg.get(차수, {})
            매도간격 = float(config.get("매도간격", 0.7))
            목표가 = entry_price * (1.0 + 매도간격 / 100.0)

            # 가상 목표가 도달 시
            if 현재가 >= 목표가:
                # 동일 차수에 대해 이번 루프에서 중복 발송 방지 (차수당 최초 1회만 발송)
                if 차수 not in cleared_levels:
                    cleared_levels.add(차수)
                    if DUMMY_SELL_ALERT_MODE in ["EACH", "ONCE"]:
                        더미매도_텔레그램알림(Ticker, 차수, entry_price, 현재가, config)
                    print(
                        f"✨ [더미 가상 청산] {Ticker} {차수}차 도달! "
                        f"진입 {entry_price:.6f} -> 현재 {현재가:.6f} (목표 {목표가:.6f})",
                        flush=True
                    )
            else:
                # 미청산된 차수는 유지 (이미 청산 처리된 차수는 제외)
                if 차수 not in cleared_levels:
                    new_attempts.append(a)

        # 청산된 차수가 있으면 상태 원자적 저장 (해당 차수 기록 제거로 중복 스팸 방지)
        # 미청산 목록 내에서도 동일 차수 중복을 최종 정제(deduplication)
        if cleared_levels:
            dedup_new = []
            seen_levels = set()
            for item in new_attempts:
                lvl = int(item.get("차수", 0))
                if lvl not in seen_levels and lvl not in cleared_levels:
                    seen_levels.add(lvl)
                    dedup_new.append(item)
            state["attempts"] = dedup_new
            잔고부족_상태저장(Ticker, state)

    except Exception as e:
        print(f"⚠️ [더미 가상 매도 감시 오류] {Ticker}: {e}", flush=True)


def 엑셀_스냅샷_백업(Ticker: str, file_name: str) -> str:
    """리밸런싱 직전 타임스탬프 스냅샷 백업 생성 (데이터 무결성 안전장치)"""
    try:
        if not os.path.exists(file_name):
            return ""
        ts = 지금().strftime("%Y%m%d_%H%M%S")
        base, ext = os.path.splitext(file_name)
        backup_path = f"{base}_backup_{ts}{ext}"
        import shutil
        shutil.copy2(file_name, backup_path)
        print(f"💾 [엑셀 자동 스냅샷 백업 완료] {backup_path}", flush=True)
        return backup_path
    except Exception as e:
        print(f"⚠️ [엑셀 스냅샷 백업 실패] {e}", flush=True)
        return ""

def 쿨링공백복원_실행(
    Ticker: str,
    rebase_px: float,
    rebase_qty: float,
    rebase_cost: float,
    rebase_slip: float
) -> None:
    """
    쿨링 만료 후 하이브리드 Rebase 첫 매수가 체결된 시점에 호출되어,
    직전 체결가와의 공백(Gap)을 산출하고 상위 차수 강제 손절 및 바닥 위 공백 예약 그물을 배치합니다.
    """
    if not GAP_RECONSTRUCTION_ENABLE:
        return

    file_name = trade_config[Ticker]["file_name"]
    df = 매매기록불러오기(file_name)
    미청산 = df[df["매도날짜"].isnull()].copy()

    # 미청산 포지션이 비어있는 경우 (1차수 신규 진입)
    if 미청산.empty:
        new_row = pd.DataFrame([{
            "종목": Ticker, "차수": 1, "매수날짜": 지금(),
            "매수가": rebase_px, "매수수량": rebase_qty, "매수금액": rebase_cost,
            "매수시현재가": rebase_px, "매수슬리피지%": round(rebase_slip, 4), "최고수익률": 0.0,
        }])
        df = pd.concat([df, new_row], ignore_index=True)
        매매기록저장(df, file_name)
        return

    # 1. 직전 체결가 P_prev 확인 (미청산 중 가장 낮은 매수가 = 최하위 차수)
    min_row = 미청산.sort_values(by=["매수가"]).iloc[0]
    P_prev = float(min_row["매수가"])

    # 만약 직전 체결가보다 현재가가 더 높거나 같으면 공백 없음
    if P_prev <= rebase_px:
        max_open_level = int(미청산["차수"].max())
        new_row = pd.DataFrame([{
            "종목": Ticker, "차수": max_open_level + 1, "매수날짜": 지금(),
            "매수가": rebase_px, "매수수량": rebase_qty, "매수금액": rebase_cost,
            "매수시현재가": rebase_px, "매수슬리피지%": round(rebase_slip, 4), "최고수익률": 0.0,
        }])
        df = pd.concat([df, new_row], ignore_index=True)
        매매기록저장(df, file_name)
        return

    # 2. 공백률 및 공백 차수 수 N 계산
    Gap_pct = ((P_prev - rebase_px) / P_prev) * 100.0
    grid_cfg = trade_config[Ticker]["grid_config"]
    interval_pct = abs(float(grid_cfg.get(2, {}).get("매수간격", 0.7)))
    N = int(Gap_pct // interval_pct)

    # 공백 차수가 2개 미만이면(단순 1개 차수 하락) 일반 매수로 처리
    if N < 2:
        max_open_level = int(미청산["차수"].max())
        new_row = pd.DataFrame([{
            "종목": Ticker, "차수": max_open_level + 1, "매수날짜": 지금(),
            "매수가": rebase_px, "매수수량": rebase_qty, "매수금액": rebase_cost,
            "매수시현재가": rebase_px, "매수슬리피지%": round(rebase_slip, 4), "최고수익률": 0.0,
        }])
        df = pd.concat([df, new_row], ignore_index=True)
        매매기록저장(df, file_name)
        return

    # 3. 실제 손절 차수 수 K 결정 (최대 캡 적용)
    K = min(N - 1, MAX_GAP_CUT_LEVELS, len(미청산))
    if K <= 0:
        max_open_level = int(미청산["차수"].max())
        new_row = pd.DataFrame([{
            "종목": Ticker, "차수": max_open_level + 1, "매수날짜": 지금(),
            "매수가": rebase_px, "매수수량": rebase_qty, "매수금액": rebase_cost,
            "매수시현재가": rebase_px, "매수슬리피지%": round(rebase_slip, 4), "최고수익률": 0.0,
        }])
        df = pd.concat([df, new_row], ignore_index=True)
        매매기록저장(df, file_name)
        return

    print(
        f"🌪️ [공백 복원 롤다운 발동] {Ticker} | 직전가={P_prev:.4f} ➔ 체결가={rebase_px:.4f} "
        f"| 공백={Gap_pct:.2f}% ({N}개) | 상위 {K}개 손절 청산 개시",
        flush=True
    )

    # 4. 안전장치: 엑셀 자동 스냅샷 백업
    엑셀_스냅샷_백업(Ticker, file_name)

    # 5. 상위 K개 차수 시장가 강제 손절 (reduceOnly)
    손절대상 = 미청산.sort_values(by=["매수가"], ascending=False).iloc[:K]
    cut_indices = list(손절대상.index)
    cut_levels = list(손절대상["차수"].values)
    total_cut_qty = float(손절대상["매수수량"].sum())

    sell_px = rebase_px
    try:
        sell_res = bitgetX.create_market_sell_order(
            symbol=Ticker, amount=total_cut_qty, params={"reduceOnly": True}
        )
        sell_px, sell_qty, sell_cost = 체결가_조회(
            Ticker, sell_res, fallback_price=rebase_px, fallback_qty=total_cut_qty
        )
    except Exception as e:
        print(f"⚠️ [공백 복원 강제 손절 거래소 주문 실패 - fallback 체결가 적용] {e}", flush=True)

    # 엑셀 손절 마감 처리
    total_loss = 0.0
    for c_idx in cut_indices:
        b_px = float(df.loc[c_idx, "매수가"])
        b_qty = float(df.loc[c_idx, "매수수량"])
        b_fee = float(df.loc[c_idx, "수수료"]) if pd.notna(df.loc[c_idx, "수수료"]) else 0.0
        pnl = (sell_px - b_px) * b_qty - b_fee
        total_loss += pnl
        df.loc[c_idx, "매도날짜"] = 지금()
        df.loc[c_idx, "매도가"] = sell_px
        df.loc[c_idx, "매도수량"] = b_qty
        df.loc[c_idx, "매도금액"] = sell_px * b_qty
        df.loc[c_idx, "실현손익"] = round(pnl, 4)
        df.loc[c_idx, "수익률"] = round(((sell_px - b_px) / b_px) * 100.0, 4)
        df.loc[c_idx, "청산유형"] = "공백복원손절"

    # 6. 잔여 미청산 차수 Re-indexing (높은 매수가 순으로 1, 2... M 차수 부여)
    잔여 = df[df["매도날짜"].isnull()].copy()
    잔여_정렬 = 잔여.sort_values(by=["매수가"], ascending=False)
    M = len(잔여_정렬)
    for new_level, (r_idx, _) in enumerate(잔여_정렬.iterrows(), start=1):
        df.loc[r_idx, "차수"] = new_level

    # 7. 바닥 Rebase 체결 차수 번호 확정 (M + K + 1 차수)
    rebase_level = M + K + 1
    new_row = pd.DataFrame([{
        "종목": Ticker,
        "차수": rebase_level,
        "매수날짜": 지금(),
        "매수가": rebase_px,
        "매수수량": rebase_qty,
        "매수금액": rebase_cost,
        "매수시현재가": rebase_px,
        "매수슬리피지%": round(rebase_slip, 4),
        "최고수익률": 0.0,
    }])
    df = pd.concat([df, new_row], ignore_index=True)
    매매기록저장(df, file_name)

    # 8. 바닥 체결가 바로 위 1차 동적 예약 목표가 산출 (P_prev와 0.7% 미만 근접 시 생략)
    default_amount = float(trade_config[Ticker]["grid_config"].get(1, {}).get("매수금액", 143.0))
    first_target = round(rebase_px * (1.0 + interval_pct / 100.0), 6)
    safety_ceiling = P_prev * (1.0 - interval_pct / 100.0) if P_prev > 0 else 0.0

    orders = []
    remaining_count = 0
    if P_prev > 0 and first_target >= safety_ceiling:
        print(
            f"⚠️ [공백 복원] 첫 예약 목표가({first_target:.4f})가 기준가({P_prev:.4f})와 {interval_pct}% 이내로 근접하여 예약 배치 생략",
            flush=True
        )
    else:
        orders.append({
            "step": 1,
            "target_price": first_target,
            "amount_usdt": default_amount,
            "created_at": 지금().strftime("%Y-%m-%d %H:%M:%S")
        })
        remaining_count = K

    공백복원_상태저장(Ticker, {
        "orders": orders,
        "remaining_count": remaining_count,
        "total_planned": K,
        "last_rebalance": 지금().strftime("%Y-%m-%d %H:%M:%S"),
        "base_price": rebase_px,
        "base_level": rebase_level,
        "top_ceiling": M,
        "prev_price": P_prev,
        "gap_pct": round(Gap_pct, 4)
    })
    잔고부족_트래커_클리어(Ticker)

    # 9. 텔레그램 종합 리포트 발송
    try:
        title_sym = Ticker.split("/")[0]
        cut_str = ", ".join([f"{lvl}차" for lvl in cut_levels])
        if orders:
            targets_str = f"   · 1차 동적 예약가: `{first_target:.4f} USDT` (체결 시 실체결가 +{interval_pct}% 자동 재조정 / 총 {K}회 한도)"
        else:
            targets_str = "   · 기준가 근접으로 예약 생략 (안전 종료)"
        msg = (
            f"🌪️ *[{title_sym}_쿨링 공백 복원 롤다운 완료]*\n"
            f"━━━━━━━━━━━━━━━\n"
            f"📊 *종목*: `{Ticker}`\n"
            f"📉 *공백 발생*: `{Gap_pct:.2f}%` (추정 공백 `{N}개 차수`)\n"
            f"✂️ *상위 손절 마감*: `{len(cut_levels)}개` (`{cut_str}`)\n"
            f"💰 *확정 손실*: `{total_loss:+.4f} USDT`\n"
            f"🔄 *잔여 차수 재배치*: `1차 ~ {M}차` (총 {M}개 승계)\n"
            f"⚓ *바닥 Rebase 체결*: `{rebase_level}차` (`{rebase_px:.4f} USDT`)\n"
            f"🕸️ *공백 복원 동적 예약 ({len(orders)}개 대기)*:\n{targets_str}\n"
            f"🕒 *시간*: `{지금().strftime('%Y-%m-%d %H:%M:%S')}`\n"
            f"━━━━━━━━━━━━━━━"
        )
        telegram_alert.SendMessage(msg)
    except Exception as e:
        print(f"⚠️ [공백 복원 텔레그램 발송 실패] {e}", flush=True)

    if not orders:
        공백복원_종료_미매도차수_점검(Ticker, 사유="바닥 목표가가 기준가에 근접하여 예약 배치 생략 (안전 종료)", 현재가=rebase_px)

def 공백복원_종료_미매도차수_점검(Ticker: str, 사유: str, 현재가: float = 0.0) -> None:
    """
    공백 복원 세션이 종료(완료 또는 취소/만료)될 때 호출되어:
    1) 엑셀 스냅샷 백업
    2) 미청산(미매도) 포지션을 매수가 내림차순(높은 가격 ➔ 낮은 가격)으로 정렬
    3) 중간 결번(Gap)이 있는 경우 1차부터 N차까지 빈틈없이 압축 재정렬(Normalization)
    4) 비트겟 거래소 실제 롱 포지션 수량과 엑셀 미매도 총 수량 1:1 정합성 대조
    5) 가격 역전(Price Inversion) 무결성 검증
    6) 최하위 N차 기준 다음 신규 하락 매수 목표가(N+1차) 산출
    7) 공백 복원 상태 완전 초기화 (공백복원_상태클리어)
    8) 텔레그램 종합 점검 리포트 발송
    """
    try:
        file_name = trade_config[Ticker]["file_name"]
        엑셀_스냅샷_백업(Ticker, file_name)
        df = 매매기록불러오기(file_name)
        미청산 = df[df["매도날짜"].isnull()].copy()

        if 미청산.empty:
            공백복원_상태클리어(Ticker)
            print(f"ℹ️ [미매도 차수 점검] {Ticker} 미청산 포지션 없음", flush=True)
            return

        # 1) 매수가 내림차순 정렬 (높은 가격 -> 낮은 가격)
        미청산_정렬 = 미청산.sort_values(by=["매수가"], ascending=False)
        old_levels = list(미청산_정렬["차수"].values)
        total_open = len(미청산_정렬)
        target_levels = list(range(1, total_open + 1))

        reindexed = False
        if old_levels != target_levels:
            # 결번 또는 불일치 발견 ➔ 1부터 N까지 압축 재정렬
            for new_level, (r_idx, _) in enumerate(미청산_정렬.iterrows(), start=1):
                df.loc[r_idx, "차수"] = new_level
            매매기록저장(df, file_name)
            reindexed = True
            reindex_desc = f"`{old_levels[0]}~{old_levels[-1]}차` ({len(old_levels)}개) ➔ `1차 ~ {total_open}차` 결번 압축 정규화 완료"
        else:
            reindex_desc = f"`1차 ~ {total_open}차` (결번 없이 완벽 연속)"

        # 2) 가격 역전(Price Inversion) 검증
        prices = list(미청산_정렬["매수가"].values)
        is_price_ordered = all(prices[i] >= prices[i + 1] for i in range(len(prices) - 1))
        price_order_str = "✅ 가격 내림차순 정합성 100% 정상" if is_price_ordered else "⚠️ 가격 역전 감지 (확인 필요)"

        # 3) 비트겟 거래소 실제 롱 포지션 수량 대조
        excel_total_qty = float(pd.to_numeric(미청산_정렬["매수수량"], errors="coerce").fillna(0.0).sum())
        exchange_pos = 현재_롱포지션_현황_조회(Ticker)
        exchange_qty = float(exchange_pos.get("contracts", 0.0) or 0.0) if exchange_pos else 0.0
        qty_diff = abs(excel_total_qty - exchange_qty)
        qty_match_str = (
            f"✅ 거래소 실포지션(`{exchange_qty}`)과 엑셀 수량(`{excel_total_qty:.4f}`) 100% 일치"
            if qty_diff < 0.0001 else
            f"⚠️ 거래소 수량(`{exchange_qty}`) vs 엑셀 수량(`{excel_total_qty:.4f}`) 차이({qty_diff:.4f})"
        )

        # 4) 최하위 N차 기준 다음 신규 하락 매수 목표가(N+1차) 산출
        lowest_px = float(prices[-1])
        next_level = total_open + 1
        grid_cfg = trade_config[Ticker]["grid_config"]
        next_cfg = grid_cfg.get(next_level, grid_cfg.get(2, {}))
        next_interval = abs(float(next_cfg.get("매수간격", 0.7)))
        next_buy_target = lowest_px * (1.0 - next_interval / 100.0)

        # 5) 공백 복원 상태 완전 클리어
        공백복원_상태클리어(Ticker)

        # 6) 콘솔 및 텔레그램 종합 점검 리포트 발송
        title_sym = Ticker.split("/")[0]
        print(
            f"🔍 [공백 복원 종료 미매도 차수 점검] {Ticker} | 사유: {사유} | "
            f"보유: {total_open}개 차수 | {reindex_desc} | {qty_match_str}",
            flush=True
        )

        msg = (
            f"🔍 *[{title_sym}_공백 복원 종료 및 미매도 차수 점검]*\n"
            f"━━━━━━━━━━━━━━━\n"
            f"📊 *종목*: `{Ticker}`\n"
            f"📌 *종료 사유*: {사유}\n"
            f"🔢 *미매도 차수*: `1차 ~ {total_open}차` (총 {total_open}개 포지션)\n"
            f"🔄 *차수 정규화*: {reindex_desc}\n"
            f"⚖️ *거래소 대조*: {qty_match_str}\n"
            f"📐 *가격 정합성*: {price_order_str}\n"
            f"⚓ *최하위 보유가*: `{lowest_px:.4f} USDT` ({total_open}차)\n"
            f"🔜 *다음 매수 대기*: `{next_level}차` (`{next_buy_target:.4f} USDT`, -{next_interval:.2f}%)\n"
            f"🕒 *시간*: `{지금().strftime('%Y-%m-%d %H:%M:%S')}`\n"
            f"━━━━━━━━━━━━━━━"
        )
        telegram_alert.SendMessage(msg)
    except Exception as e:
        print(f"⚠️ [미매도 차수 점검 실행 실패] {e}", flush=True)


def 비상_증거금확보_쿨링공백복원_실행(
    Ticker: str,
    현재가: float,
    차수: int,
    config: Dict[str, Any],
    rebase_mode: str
) -> bool:
    """
    [v2.4 비상 증거금 부족 자가 치유(Self-Healing) 롤다운 실행 함수]
    - 하이브리드 Rebase 반등 시점에 가용 잔고/증거금이 부족하여 바닥 매수가 불가능한 경우 호출됨.
    - 1단계: 현재가(P_virtual)를 가상 기준가로 삼아 공백 N 및 필요 손절 개수 K를 사전에 100% 정밀 계산.
             * N < 2 인 경우 K = 0 (과잉 손절 완벽 차단, 불필요한 실현 손실 방지).
    - 2단계: K >= 1 인 경우, 상위 K개 물량을 즉시 시장가 선(先) 손절(reduceOnly)하여 증거금 환급 & 유동성 확보.
    - 3단계: 환급된 현금으로 바닥 Rebase 실제 시장가 매수 체결.
    - 4단계: 실제 체결가(P_real_fill) 기준으로 첫 동적 예약 매수가(+0.7%) 재산출 및 배치.
    - 5단계: 잔여 차수 재배치(1~M) 및 텔레그램 종합 알림 발송.
    """
    if not GAP_RECONSTRUCTION_ENABLE:
        return False

    file_name = trade_config[Ticker]["file_name"]
    df = 매매기록불러오기(file_name)
    미청산 = df[df["매도날짜"].isnull()].copy()

    if 미청산.empty:
        print(f"⚠️ [비상 롤다운] {Ticker} 미청산 포지션이 없어 손절 불가", flush=True)
        return False

    # 1. 쿨링 전 마지막 체결가 P_prev 확인 (미청산 중 가장 낮은 매수가 = 최하위 차수)
    min_row = 미청산.sort_values(by=["매수가"]).iloc[0]
    P_prev = float(min_row["매수가"])

    # 가상 바닥 기준가 P_virtual = 현재가
    P_virtual = float(현재가)
    if P_prev <= P_virtual:
        print(f"⚠️ [비상 롤다운] 직전가({P_prev:.4f}) <= 현재가({P_virtual:.4f}): 공백 없음 (손절 취소)", flush=True)
        return False

    # 2. 공백률 및 공백 차수 수 N 사전 계산
    Gap_pct = ((P_prev - P_virtual) / P_prev) * 100.0
    grid_cfg = trade_config[Ticker]["grid_config"]
    interval_pct = abs(float(grid_cfg.get(2, {}).get("매수간격", 0.7)))
    N = int(Gap_pct // interval_pct)

    # ⭐ 핵심 원칙: 공백 차수가 2개 미만이면(단순 1칸 하락) 손절 개수 K = 0 (과잉 손절 방지!)
    if N < 2:
        print(
            f"ℹ️ [비상 롤다운 스킵] {Ticker} 공백 {Gap_pct:.2f}% (N={N} < 2, 단순 하락): 과잉 손절 방지를 위해 손절 0개 적용",
            flush=True
        )
        return False

    # 3. 필요 손절 차수 수 K 결정 (최대 캡 적용)
    K = min(N - 1, MAX_GAP_CUT_LEVELS, len(미청산))
    if K <= 0:
        print(f"ℹ️ [비상 롤다운 스킵] 손절 대상 수량 0개 (K={K})", flush=True)
        return False

    print(
        f"🚨 [비상 증거금 자가치유 롤다운 개시] {Ticker} | 가상기준가={P_virtual:.4f} | 직전가={P_prev:.4f} "
        f"| 공백={Gap_pct:.2f}% ({N}개) | 상위 {K}개 선(先) 손절을 통한 증거금 확보 시작",
        flush=True
    )

    # 4. 안전장치: 엑셀 자동 스냅샷 백업
    엑셀_스냅샷_백업(Ticker, file_name)

    # 5. [선(先) 손절] 상위 K개 차수 시장가 강제 손절 (reduceOnly)
    손절대상 = 미청산.sort_values(by=["매수가"], ascending=False).iloc[:K]
    cut_indices = list(손절대상.index)
    cut_levels = list(손절대상["차수"].values)
    total_cut_qty = float(손절대상["매수수량"].sum())

    sell_px = P_virtual
    try:
        sell_res = bitgetX.create_market_sell_order(
            symbol=Ticker, amount=total_cut_qty, params={"reduceOnly": True}
        )
        sell_px, sell_qty, sell_cost = 체결가_조회(
            Ticker, sell_res, fallback_price=P_virtual, fallback_qty=total_cut_qty
        )
    except Exception as e:
        print(f"⚠️ [비상 롤다운 강제 손절 거래소 주문 실패 - fallback 체결가 적용] {e}", flush=True)

    # 엑셀 손절 마감 처리
    total_loss = 0.0
    for c_idx in cut_indices:
        b_px = float(df.loc[c_idx, "매수가"])
        b_qty = float(df.loc[c_idx, "매수수량"])
        b_fee = float(df.loc[c_idx, "수수료"]) if pd.notna(df.loc[c_idx, "수수료"]) else 0.0
        pnl = (sell_px - b_px) * b_qty - b_fee
        total_loss += pnl
        df.loc[c_idx, "매도날짜"] = 지금()
        df.loc[c_idx, "매도가"] = sell_px
        df.loc[c_idx, "매도수량"] = b_qty
        df.loc[c_idx, "매도금액"] = sell_px * b_qty
        df.loc[c_idx, "실현손익"] = round(pnl, 4)
        df.loc[c_idx, "수익률"] = round(((sell_px - b_px) / b_px) * 100.0, 4)
        df.loc[c_idx, "청산유형"] = "공백복원손절"

    # 손절 후 원장 저장 (손절 확정)
    매매기록저장(df, file_name)
    time.sleep(0.5)

    # 6. [후(後) 실체결] 손절로 확보된 가용 현금/증거금으로 바닥 Rebase 실제 시장가 매수
    수량 = USDT금액을_수량으로변환(Ticker, float(config["매수금액"]), 현재가)
    if 수량 == 0.0:
        print(f"🚫 [비상 롤다운] 매수 최소 수량 미만", flush=True)
        return False

    실제체결가 = float(현재가)
    실제체결수량 = 수량
    실제체결금액 = 수량 * 실제체결가
    매수슬리피지 = 0.0

    try:
        buy_res = bitgetX.create_market_buy_order(
            symbol=Ticker, amount=수량, params={"reduceOnly": False}
        )
        실제체결가, 실제체결수량, 실제체결금액 = 체결가_조회(
            Ticker, buy_res, fallback_price=현재가, fallback_qty=수량
        )
        매수슬리피지 = 슬리피지_계산(현재가, 실제체결가)
    except Exception as e:
        print(f"⚠️ [비상 롤다운 바닥 매수 주문 실패 - fallback 체결가 적용] {e}", flush=True)

    # 7. 잔여 미청산 차수 Re-indexing (높은 매수가 순으로 1, 2... M 차수 부여)
    잔여 = df[df["매도날짜"].isnull()].copy()
    잔여_정렬 = 잔여.sort_values(by=["매수가"], ascending=False)
    M = len(잔여_정렬)
    for new_level, (r_idx, _) in enumerate(잔여_정렬.iterrows(), start=1):
        df.loc[r_idx, "차수"] = new_level

    # 8. 바닥 Rebase 체결 차수 번호 확정 (M + K + 1 차수)
    rebase_level = M + K + 1
    new_row = pd.DataFrame([{
        "종목": Ticker,
        "차수": rebase_level,
        "매수날짜": 지금(),
        "매수가": 실제체결가,
        "매수수량": 실제체결수량,
        "매수금액": 실제체결금액,
        "매수시현재가": 현재가,
        "매수슬리피지%": round(매수슬리피지, 4),
        "최고수익률": 0.0,
    }])
    df = pd.concat([df, new_row], ignore_index=True)
    매매기록저장(df, file_name)

    # 9. 바닥 실제체결가(P_real_fill) 기준 동적 예약 목표가 산출 (+0.7%)
    default_amount = float(trade_config[Ticker]["grid_config"].get(1, {}).get("매수금액", 143.0))
    first_target = round(실제체결가 * (1.0 + interval_pct / 100.0), 6)
    safety_ceiling = P_prev * (1.0 - interval_pct / 100.0) if P_prev > 0 else 0.0

    orders = []
    remaining_count = 0
    if P_prev > 0 and first_target >= safety_ceiling:
        print(
            f"⚠️ [비상 롤다운] 첫 예약 목표가({first_target:.4f})가 기준가({P_prev:.4f})와 {interval_pct}% 이내로 근접하여 예약 배치 생략",
            flush=True
        )
    else:
        orders.append({
            "step": 1,
            "target_price": first_target,
            "amount_usdt": default_amount,
            "created_at": 지금().strftime("%Y-%m-%d %H:%M:%S")
        })
        remaining_count = K

    공백복원_상태저장(Ticker, {
        "orders": orders,
        "remaining_count": remaining_count,
        "total_planned": K,
        "last_rebalance": 지금().strftime("%Y-%m-%d %H:%M:%S"),
        "base_price": 실제체결가,
        "base_level": rebase_level,
        "top_ceiling": M,
        "prev_price": P_prev,
        "gap_pct": round(Gap_pct, 4)
    })

    # 상태 클리어
    쿨링_rebase_클리어(Ticker)
    잔고부족_트래커_클리어(Ticker)

    # 10. 텔레그램 종합 리포트 발송
    try:
        title_sym = Ticker.split("/")[0]
        cut_str = ", ".join([f"{lvl}차" for lvl in cut_levels])
        if orders:
            targets_str = f"   · 1차 동적 예약가: `{first_target:.4f} USDT` (실체결가 기준 +{interval_pct}% 동적 산출 / 총 {K}회 한도)"
        else:
            targets_str = "   · 기준가 근접으로 예약 생략 (안전 종료)"
        msg = (
            f"🌪️ *[{title_sym}_비상 자가치유 롤다운 완료]* 🛡️\n"
            f"━━━━━━━━━━━━━━━\n"
            f"📊 *종목*: `{Ticker}`\n"
            f"💡 *유형*: 증거금 고갈 비상 자가치유 (선 손절 ➔ 증거금 확보 ➔ Rebase 실체결)\n"
            f"📉 *공백 발생*: `{Gap_pct:.2f}%` (추정 공백 `{N}개 차수`)\n"
            f"✂️ *상위 선(先)손절*: `{len(cut_levels)}개` (`{cut_str}`)\n"
            f"💰 *확정 손실*: `{total_loss:+.4f} USDT`\n"
            f"🔄 *잔여 차수 재배치*: `1차 ~ {M}차` (총 {M}개 승계)\n"
            f"⚓ *바닥 Rebase 실체결*: `{rebase_level}차` (`{실제체결가:.4f} USDT`, 슬립 `{매수슬리피지:+.4f}%`)\n"
            f"🕸️ *공백 복원 동적 예약 ({len(orders)}개 대기)*:\n{targets_str}\n"
            f"🕒 *시간*: `{지금().strftime('%Y-%m-%d %H:%M:%S')}`\n"
            f"━━━━━━━━━━━━━━━"
        )
        telegram_alert.SendMessage(msg)
    except Exception as e:
        print(f"⚠️ [비상 롤다운 텔레그램 발송 실패] {e}", flush=True)

    if not orders:
        공백복원_종료_미매도차수_점검(Ticker, 사유="바닥 목표가가 기준가에 근접하여 예약 배치 생략 (안전 종료)", 현재가=실제체결가)

    return True


def 공백복원_예약매수_감시(Ticker: str, 현재가: float) -> bool:
    """
    메인 폴링마다 호출되어 공백 예약 매수 목표가 도달을 감시하고,
    도달 시 시장가 매수를 실행한 뒤:
    1) 바닥 Rebase 차수 기준 역순(-1씩) 차수 부여 및 엑셀 기록 (기존 차수 불변 보존)
    2) 일반 그리드 매수 표준 템플릿의 단일 통합 텔레그램 알림 발송
    3) 실제 체결가 + interval_pct%로 다음 목표가를 실시간 동적 재조정 (Dynamic Retargeting)
    4) 다음 목표가가 P_prev와 interval_pct% 이내로 근접하거나 현재가가 P_prev 이상 상향 돌파 시,
       잔여 예약 자동 취소 후 미매도 차수 점검(결번 압축 재정렬 및 거래소 대조) 자동 실행
    """
    if not GAP_RECONSTRUCTION_ENABLE or 현재가 <= 0:
        return False

    state = 공백복원_상태로드(Ticker)
    orders = state.get("orders", [])
    if not orders:
        return False

    P_prev = float(state.get("prev_price", 0.0))
    grid_cfg = trade_config[Ticker]["grid_config"]
    interval_pct = abs(float(grid_cfg.get(2, {}).get("매수간격", 0.7)))

    # [보완 1] 현재가가 쿨링 전 최하위 기준가(P_prev) 이상으로 상승 돌파 시 잔여 예약 전량 만료 폐기 및 미매도 점검
    if P_prev > 0 and 현재가 >= P_prev:
        cleared_count = len(orders)
        state["orders"] = []
        state["remaining_count"] = 0
        공백복원_상태저장(Ticker, state)
        사유 = f"현재가(`{현재가:.4f} USDT`)가 기준가(`{P_prev:.4f} USDT`) 상향 돌파 ➔ 잔여 예약 {cleared_count}개 자동 취소"
        print(f"🌪️ [공백 복원 예약 만료] {Ticker} {사유}", flush=True)
        공백복원_종료_미매도차수_점검(Ticker, 사유=사유, 현재가=현재가)
        return False

    target_order = orders[0]
    target_price = float(target_order.get("target_price", 0.0))
    if 현재가 < target_price:
        return False

    amount_usdt = float(target_order.get("amount_usdt", 143.0))
    수량 = USDT금액을_수량으로변환(Ticker, amount_usdt, 현재가)
    if 수량 <= 0.0:
        return False

    try:
        order_res = bitgetX.create_market_buy_order(
            symbol=Ticker, amount=수량, params={"reduceOnly": False}
        )
        실제체결가, 실제수량, 실제금액 = 체결가_조회(
            Ticker, order_res, fallback_price=현재가, fallback_qty=수량
        )
        매수슬리피지 = 슬리피지_계산(현재가, 실제체결가)
    except Exception as e:
        print(f"⚠️ [공백 복원 예약 매수 주문 실행 실패] {e}", flush=True)
        return False

    # 1) 역순 차수 배정: 바닥 Rebase 차수(예: 25) 기준 -1씩 카운트다운 (24차, 23차...)
    rebase_level = int(state.get("base_level", 25))
    current_step = int(target_order.get("step", 1))
    assigned_level = rebase_level - current_step

    file_name = trade_config[Ticker]["file_name"]
    df = 매매기록불러오기(file_name)

    temp_row = pd.DataFrame([{
        "종목": Ticker,
        "차수": assigned_level,
        "매수날짜": 지금(),
        "매수가": 실제체결가,
        "매수수량": 실제수량,
        "매수금액": 실제금액,
        "매수시현재가": 현재가,
        "매수슬리피지%": round(매수슬리피지, 4),
        "최고수익률": 0.0,
    }])
    # 기존 행들의 차수는 절대로 덮어쓰지 않고 새 행만 추가!
    df = pd.concat([df, temp_row], ignore_index=True)
    매매기록저장(df, file_name)

    # 2) ── [동적 재조정: Dynamic Retargeting] 실제 체결가 기준으로 다음 목표가 재산출 ──
    rem_count = int(state.get("remaining_count", 1)) - 1
    next_orders = []
    status_note = ""
    session_ended = False
    end_reason = ""

    if rem_count > 0:
        next_target = round(실제체결가 * (1.0 + interval_pct / 100.0), 6)
        safety_ceiling = P_prev * (1.0 - interval_pct / 100.0) if P_prev > 0 else 0.0

        # 다음 목표가가 쿨링 전 마지막 매수가(P_prev)와 interval_pct% 이내로 좁혀지면 잔여 예약 전량 취소
        if P_prev > 0 and next_target >= safety_ceiling:
            status_note = (
                f"🛑 다음 예약가(`{next_target:.4f}`)가 기준가(`{P_prev:.4f}`)와 "
                f"`{interval_pct}%` 이내로 근접하여 잔여 예약 `{rem_count}개` 자동 취소 (공백 복원 정상 종료)"
            )
            session_ended = True
            end_reason = status_note
            rem_count = 0
            next_orders = []
            print(
                f"🛑 [공백 복원 자동 취소] {Ticker} 다음 목표가({next_target:.4f})가 기준가({P_prev:.4f})에 {interval_pct}% 이내 근접 ➔ 잔여 예약 전량 취소",
                flush=True
            )
        else:
            next_orders = [{
                "step": current_step + 1,
                "target_price": next_target,
                "amount_usdt": amount_usdt,
                "created_at": 지금().strftime("%Y-%m-%d %H:%M:%S")
            }]
            status_note = f"🔜 다음 동적 예약가: `{next_target:.4f} USDT` (+{interval_pct}% 재조정 / 남은 {rem_count}회)"
    else:
        status_note = "🏁 계획된 공백 예약 횟수 모두 소진 완료"
        session_ended = True
        end_reason = status_note
        next_orders = []

    state["orders"] = next_orders
    state["remaining_count"] = rem_count
    공백복원_상태저장(Ticker, state)

    print(
        f"🎯 [공백 복원 예약 매수 체결] {Ticker} | 차수={assigned_level}차 | 체결가={실제체결가:.4f} USDT "
        f"| 슬리피지={매수슬리피지:+.4f}% | {status_note}",
        flush=True
    )

    # 3) 단일 통합 텔레그램 매수 알림 발송 (일반 그리드 매수 표준 템플릿 활용)
    try:
        title_symbol = Ticker.split('/')[0]
        손익블록 = ""
        if TOTAL_PROFIT_TARGET_ENABLE:
            try:
                손익블록 = 총수익_텔레그램라인(Ticker, 현재가)
            except Exception as e:
                print(f"[예약 매수 텔레그램 - 총수익 추가 실패] {e}")
        보유금라인 = 보유금_텔레그램라인(Ticker, df)

        cfg = grid_cfg.get(assigned_level, grid_cfg.get(1, {}))
        sell_interval = float(cfg.get("매도간격", 0.7))
        목표매도가 = 실제체결가 * (1.0 + sell_interval / 100.0)

        msg = (
            f"🚨*[{title_symbol}_그리드롱 진입 완료 (공백 복원 예약)]*🚨\n"
            f"━━━━━━━━━━━━━━━\n"
            f"📊 *종목*: `{Ticker}`\n"
            f"🔢 *차수*: `{assigned_level}차` (바닥 역순: {rebase_level}차 ➔ {assigned_level}차)\n"
            f"💵 *시장가*: `{현재가:.6f} USDT`\n"
            f"🎯 *실제 체결가*: `{실제체결가:.6f} USDT`\n"
            f"📐 *슬리피지*: `{매수슬리피지:+.4f}%`\n"
            f"📦 *체결수량*: `{실제수량}`\n"
            f"💰 *체결금액*: `{실제금액:.6f} USDT`\n"
            f"🎯 *목표 매도가*: `{목표매도가:.6f} USDT` (+{sell_interval:.2f}%)\n"
            f"🌪️ *공백 복원 현황*: {status_note}\n"
            f"{손익블록}"
            f"{보유금라인}"
            f"🕒 *시간*: `{지금().strftime('%Y-%m-%d %H:%M:%S')}`\n"
            f"━━━━━━━━━━━━━━━"
        )
        telegram_alert.SendMessage(msg)
    except Exception as e:
        print(f"⚠️ [공백 복원 예약 매수 텔레그램 발송 실패] {e}", flush=True)

    # 4) 공백 복원 종료 시점 미매도 차수 점검 로직 실행
    if session_ended:
        공백복원_종료_미매도차수_점검(Ticker, 사유=end_reason, 현재가=실제체결가)

    return True


def 시작시_매수기준_요약(Ticker: str, 현재가: float = 0.0) -> str:
    """
    봇 시작 알림용: 다음 매도 목표가 / 현재가 / 다음 매수 기준가 / 매수 목표가 요약.
    - 가격 내림차순(높은 가격 ➔ 낮은 가격)으로 정렬하여 표시
    - 미청산 포지션이 있을 경우 가격 상승 시 최우선 도달하는 '다음 매도 목표가' 포함
    - rebase 중: A 및 (>A 즉시 / ≤A 저점 반등) 안내
    """
    sym = Ticker.split("/")[0]

    if 현재가 <= 0:
        현재가 = 현재가조회(Ticker)

    cur_item = {
        "price": 현재가 if 현재가 > 0 else None,
        "text": f"💰 *{sym} 시작 시 현재가*: `{현재가:.6f} USDT`" if 현재가 > 0 else f"💰 *{sym} 시작 시 현재가*: `조회 실패`",
        "order": 3
    }

    try:
        df = 매매기록불러오기(trade_config[Ticker]["file_name"])
    except Exception as e:
        base_item = {
            "price": None,
            "text": f"📌 *{sym} 시작 시 기준가*: `엑셀 로드 실패 ({e})`",
            "order": 2
        }
        items = [cur_item, base_item]
        items.sort(key=lambda x: (0 if x["price"] is not None and x["price"] > 0 else 1, -(x["price"] or 0), x["order"]))
        return "\n".join(x["text"] for x in items) + "\n"

    미청산 = df[df["매도날짜"].isnull()] if (df is not None and not df.empty) else pd.DataFrame()
    grid = trade_config[Ticker].get("grid_config", GRID_TEMPLATE)
    rebase_A = 쿨링_rebase_앵커조회(Ticker)

    # 1) 다음 매도 목표가 계산 (미청산 포지션 중 가격 상승 시 가장 먼저 도달하는 최우선 매도 목표가)
    sell_item = None
    if not 미청산.empty:
        매도_후보 = []
        for _, row in 미청산.iterrows():
            try:
                c_차수 = int(row["차수"])
                c_매수가 = float(row["매수가"])
                c_간격 = float(grid[c_차수]["매도간격"]) if c_차수 in grid else 0.7
                c_목표가 = c_매수가 * (1 + c_간격 / 100.0)
                매도_후보.append({
                    "차수": c_차수,
                    "매수가": c_매수가,
                    "간격": c_간격,
                    "목표가": c_목표가
                })
            except Exception:
                continue

        if 매도_후보:
            다음_매도 = min(매도_후보, key=lambda x: x["목표가"])
            sell_item = {
                "price": 다음_매도["목표가"],
                "text": f"📈 *{sym} 다음 매도 목표가*: `{다음_매도['목표가']:.6f} USDT` ({다음_매도['차수']}차, {다음_매도['간격']:+.2f}%)",
                "order": 1
            }

    # 2) 다음 매수 차수 판별
    다음차수: Optional[int] = None
    open_lvls_init = 미청산["차수"] if not 미청산.empty else pd.Series(dtype=int)
    max_open_lvl_init = int(open_lvls_init.max()) if not open_lvls_init.empty else 0
    gap_state_init = 공백복원_상태로드(Ticker) if GAP_RECONSTRUCTION_ENABLE else {}
    has_gap_orders_init = bool(gap_state_init.get("orders", []))

    for 차수 in sorted(grid.keys()):
        해당_미청산 = 미청산[미청산["차수"] == 차수] if not 미청산.empty else pd.DataFrame()
        if not 해당_미청산.empty:
            continue
        if 차수 == 1:
            다음차수 = 1
            break
        # 공백 복원 예약 대기 중이면 상단 미복원 공백 차수는 건너뛰고 하향 차수에서 다음 차수 탐색
        if has_gap_orders_init and max_open_lvl_init > 0 and 차수 < max_open_lvl_init:
            continue
        prev_open = 미청산[미청산["차수"] == 차수 - 1] if not 미청산.empty else pd.DataFrame()
        if prev_open.empty:
            v_prev = 잔고부족_가상차수_조회(Ticker, 차수 - 1)
            if v_prev is None:
                break
        다음차수 = 차수
        break

    base_item = None
    buy_item = None

    if 다음차수 is None:
        if 미청산.empty:
            base_item = {
                "price": None,
                "text": f"📌 *{sym} 시작 시 기준가*: `1차 대기 (기준가 없음)`",
                "order": 2
            }
        else:
            max_open = int(미청산["차수"].max())
            base_item = {
                "price": None,
                "text": f"📌 *{sym} 시작 시 기준가*: `다음 매수 없음 (미청산 최대 {max_open}차)`",
                "order": 2
            }
    elif 다음차수 == 1:
        base_item = {
            "price": None,
            "text": f"📌 *{sym} 시작 시 기준가*: `1차 대기 (기준가 없음)`",
            "order": 2
        }
        if 즉시매수허용:
            buy_item = {
                "price": None,
                "text": f"📉 *{sym} 다음 매수*: `1차 즉시매수 가능`",
                "order": 4
            }
        else:
            buy_item = {
                "price": None,
                "text": f"📉 *{sym} 다음 매수*: `1차 RSI≤{RSI_THRESHOLD}`",
                "order": 4
            }
    else:
        간격 = float(grid[다음차수]["매수간격"])
        bounce = abs(간격)

        if rebase_A is not None:
            low = 쿨링_rebase_저점조회(Ticker)
            base_item = {
                "price": float(rebase_A),
                "text": f"📌 *{sym} 시작 시 기준가 A*: `{rebase_A:.6f} USDT` (쿨링 해제가)",
                "order": 2
            }
            if 현재가 > rebase_A:
                buy_item = {
                    "price": None,
                    "text": f"📉 *{sym} 다음 매수*: `{다음차수}차 즉시매수 가능` (현재가 > A)",
                    "order": 4
                }
            else:
                반등목표가 = float(low) * (1 + bounce / 100.0)
                buy_item = {
                    "price": 반등목표가,
                    "text": (
                        f"📉 *{sym} 저점*: `{float(low):.6f}` → "
                        f"반등 목표가 `{반등목표가:.6f}` ({다음차수}차, +{bounce:.2f}%)"
                    ),
                    "order": 4
                }
        else:
            prev_open = 미청산[미청산["차수"] == 다음차수 - 1] if not 미청산.empty else pd.DataFrame()
            if not prev_open.empty:
                기준가 = float(prev_open.iloc[-1]["매수가"])
                출처 = f"{다음차수 - 1}차 매수가"
            else:
                v_prev = 잔고부족_가상차수_조회(Ticker, 다음차수 - 1)
                기준가 = float(v_prev["price"]) if v_prev else 0.0
                출처 = f"{다음차수 - 1}차 가상기준가"
            목표가 = 기준가 * (1 + 간격 / 100.0)
            base_item = {
                "price": 기준가,
                "text": f"📌 *{sym} 시작 시 기준가*: `{기준가:.6f} USDT` ({출처})",
                "order": 2
            }
            buy_item = {
                "price": 목표가,
                "text": f"📉 *{sym} 다음 매수 목표가*: `{목표가:.6f} USDT` ({다음차수}차, {간격:+.2f}%)",
                "order": 4
            }

    # 3) 내림차순(높은 가격 ➔ 낮은 가격) 정렬
    items = [item for item in [sell_item, cur_item, base_item, buy_item] if item is not None]
    items.sort(key=lambda x: (
        0 if x["price"] is not None and x["price"] > 0 else 1,
        -(x["price"] or 0.0),
        x.get("order", 99)
    ))

    return "\n".join(x["text"] for x in items) + "\n"

def 최근매수횟수(df: pd.DataFrame, window_sec: int) -> int:
    """
    엑셀 기준 최근 window_sec 초 이내 '고유 차수' 개수.
    동일 차수 반복 매수(매도 후 재진입 등)는 1회로만 집계 → 급락 연타만 쿨링 대상으로.
    """
    if df is None or df.empty or "매수날짜" not in df.columns:
        return 0
    if "차수" not in df.columns:
        return 0
    cutoff = 지금() - timedelta(seconds=window_sec)
    dates = pd.to_datetime(df["매수날짜"], errors="coerce")
    recent = df.loc[dates >= cutoff]
    if recent.empty:
        return 0
    return int(recent["차수"].nunique())


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


def 총수익_계산(Ticker: str, 현재가: float) -> Tuple[float, float, float, float, float]:
    """
    총 수익 계산
    [반환] (미실현, 누적실현, 총합, baseline, 사이클수익)
    사이클수익 = 총합 - baseline  (목표 판정·텔레그램 표시에 사용)
    """
    if PROFIT_SOURCE == "EXCHANGE":
        미실현 = 미실현손익조회_거래소(Ticker)
        if 미실현 is None:  # 거래소 API 실패 → EXCEL fallback
            미실현 = 미실현손익조회_엑셀(Ticker, 현재가)
    else:
        미실현 = 미실현손익조회_엑셀(Ticker, 현재가)

    누적실현 = 누적실현손익_조회(Ticker)
    총합 = 미실현 + 누적실현
    baseline = 총수익_baseline_조회(Ticker)
    사이클수익 = 총합 - baseline
    return 미실현, 누적실현, 총합, baseline, 사이클수익


def 총수익_텔레그램라인(Ticker: str, 현재가: float) -> str:
    """매도/시작 등 텔레그램에 붙이는 총수익 요약 블록 (사이클 핵심 정보만 표시)"""
    미실현, 누적실현, 총합, baseline, 사이클 = 총수익_계산(Ticker, 현재가)
    회차 = 총수익_사이클회차_조회(Ticker)
    return (
        f"🔁 *사이클*: `{회차}회차`\n"
        f"💰 *미실현*: `{미실현:+.4f} USDT`\n"
        f"🎯 *사이클 수익*: `{사이클:+.4f} USDT` "
        f"(목표 {TOTAL_PROFIT_TARGET_USDT} USDT)\n"
    )


# --- 거래소 잔고 및 보유금 현황 함수 ---
_최근_잔여보유금_캐시: Optional[float] = None

def 거래소_USDT_잔고조회() -> Optional[float]:
    """
    거래소 fetch_balance로 선물 계좌의 남아있는 가용 USDT(free) 잔고 조회.
    ccxt 표준 필드 및 비트겟 v2 raw data를 탐색하여 안전하게 float로 반환.
    API 일시 오류 시 직전 성공 캐시값을 fallback으로 사용.
    """
    global _최근_잔여보유금_캐시
    try:
        balance = bitgetX.fetch_balance(params={"type": "swap"})

        free_val = None
        # 1) 표준 ccxt 구조 탐색
        if "USDT" in balance and isinstance(balance["USDT"], dict):
            free_val = balance["USDT"].get("free")
        if free_val is None and "free" in balance and isinstance(balance["free"], dict):
            free_val = balance["free"].get("USDT")

        # 2) raw info 탐색 (fallback)
        if free_val is None:
            info = balance.get("info", {})
            data_items = []
            if isinstance(info, list):
                data_items = info
            elif isinstance(info, dict):
                data = info.get("data")
                if isinstance(data, list):
                    data_items = data
                elif isinstance(data, dict):
                    data_items = [data]

            for item in data_items:
                if not isinstance(item, dict):
                    continue
                coin = str(item.get("marginCoin") or item.get("coin") or "").upper()
                if coin == "USDT":
                    avail = (
                        item.get("available")
                        or item.get("availableBalance")
                        or item.get("crossMarginAvailable")
                        or item.get("maxTransferOut")
                    )
                    if avail is not None:
                        free_val = avail
                        break

        if free_val is not None:
            _최근_잔여보유금_캐시 = float(free_val)
            return _최근_잔여보유금_캐시

    except Exception as e:
        print(f"[거래소 가용 잔고 조회 실패] {e}")

    return _최근_잔여보유금_캐시


def 보유금_현황_조회(Ticker: str, df: Optional[pd.DataFrame] = None) -> Tuple[float, Optional[float], Optional[float]]:
    """
    현재 진입한 총 금액, 남아있는 보유금(가용잔고), 총 보유금을 계산하여 반환.
    [반환] (현재진입총금액, 남아있는보유금, 총보유금)
    - 현재진입총금액: 엑셀 미청산 차수들의 매수금액 합계 (USDT)
    - 남아있는보유금: 거래소 선물 계좌의 가용 USDT 잔고 (USDT)
    - 총보유금: 현재진입총금액 + 남아있는보유금
    """
    if df is None:
        file_name = trade_config[Ticker]["file_name"]
        df = 매매기록불러오기(file_name)

    미청산 = df[df["매도날짜"].isnull()]
    if 미청산.empty:
        현재진입 = 0.0
    else:
        현재진입 = float(pd.to_numeric(미청산["매수금액"], errors="coerce").fillna(0.0).sum())

    남아있는보유금 = 거래소_USDT_잔고조회()
    if 남아있는보유금 is not None:
        총보유금 = 현재진입 + 남아있는보유금
    else:
        총보유금 = None

    return 현재진입, 남아있는보유금, 총보유금


def 보유금_텔레그램라인(Ticker: str, df: Optional[pd.DataFrame] = None) -> str:
    """
    그리드 매매 텔레그램 알림용 [현재 진입한 총 금액 + 남아있는 보유금의 금액 = 총 보유금] 문자열 라인 생성.
    """
    try:
        현재진입, 잔여, 총보유 = 보유금_현황_조회(Ticker, df)
        if 잔여 is not None and 총보유 is not None:
            return f"💼 *보유금 현황*: `[진입 {현재진입:,.4f} + 잔여 {잔여:,.4f} = 총 {총보유:,.4f} USDT]`\n"
        else:
            return f"💼 *보유금 현황*: `[진입 {현재진입:,.4f} + 잔여 조회실패 = 총 N/A USDT]`\n"
    except Exception as e:
        print(f"[보유금 텔레그램라인 생성 실패] {e}")
        return ""


def 전체_보유금_현황_조회() -> Tuple[float, Optional[float], Optional[float]]:
    """
    모든 운용 종목의 현재 진입한 총 금액 합계, 거래소 가용 잔고(잔여), 총 보유금 계산.
    [반환] (총진입금액, 남아있는보유금, 총보유금)
    """
    총진입 = 0.0
    for ticker in 종목코드목록:
        try:
            cfg = trade_config.get(ticker) or {}
            fp = cfg.get("file_name")
            if fp and os.path.exists(fp):
                df = 매매기록불러오기(fp)
                if df is not None and not df.empty:
                    미청산 = df[df["매도날짜"].isnull()]
                    if not 미청산.empty:
                        총진입 += float(pd.to_numeric(미청산["매수금액"], errors="coerce").fillna(0.0).sum())
        except Exception as e:
            print(f"[전체 보유금 현황 조회 오류 - {ticker}] {e}", flush=True)

    남아있는보유금 = 거래소_USDT_잔고조회()
    if 남아있는보유금 is not None:
        총보유금 = 총진입 + 남아있는보유금
    else:
        총보유금 = None

    return 총진입, 남아있는보유금, 총보유금



def 총수익_체크_및_청산(Ticker: str, 현재가: float) -> bool:
    """
    사이클 수익이 목표에 도달하면 전체 차수를 시장가로 청산.
    사이클 수익 = (미실현 + 누적실현) - baseline
    RESTART 시 청산 후 baseline을 새 누적실현으로 갱신.
    """
    global shutdown_requested

    # ── 1) 기능 비활성 또는 이미 청산 완료 락이 있으면 스킵 ──
    if not TOTAL_PROFIT_TARGET_ENABLE:
        return False
    if 총수익청산락_여부(Ticker):
        return False

    # ── 2) 총 수익 계산 ──
    미실현, 누적실현, 총합, baseline, 사이클수익 = 총수익_계산(Ticker, 현재가)

    # ── 3) 목표 미달 시 종료 (사이클 수익 기준) ──
    if 사이클수익 < TOTAL_PROFIT_TARGET_USDT:
        return False

    # ── 4) 목표 도달 → 전체 청산 시작 ──
    print(
        f"🎯 총 수익 목표 도달! "
        f"사이클 {사이클수익:+.4f} = (미실현 {미실현:+.4f} + 누적 {누적실현:+.4f}) "
        f"- baseline {baseline:+.4f} "
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
    누적_추적기 = 직전누적실현손익(df)

    if 미청산.empty:
        # 미청산 포지션 없는데 수익 목표 도달 (거의 누적실현만으로) → 락만 걸고 종료
        print(f"ℹ️ 미청산 포지션 없음. 누적실현만으로 목표 도달.")
    else:
        for idx, row in 미청산.sort_values(by=["차수"], ascending=False).iterrows():
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

            df.at[idx, "매도날짜"] = 지금()
            df.at[idx, "매도가"] = 실제체결가
            df.at[idx, "매도수량"] = 실제체결수량
            df.at[idx, "매도금액"] = 매도금액
            df.at[idx, "매도시현재가"] = 현재가
            df.at[idx, "매도슬리피지%"] = round(매도슬리피지, 4)
            df.at[idx, "수수료"] = 수수료
            df.at[idx, "실현손익"] = 실현손익
            df.at[idx, "수익률"] = 수익률
            df.at[idx, "누적실현손익"] = 현재누적
            df.at[idx, "청산유형"] = "전체청산"

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
        # 전체 청산 시 공백 복원 예약 주문 상태 파일도 깔끔히 초기화
        공백복원_상태클리어(Ticker)

    # 청산 후 누적 (미청산 없으면 기존 누적_추적기 = 직전누적)
    청산후_누적 = float(누적_추적기)
    현재회차 = 총수익_사이클회차_조회(Ticker)

    # ── 5) 텔레그램 종합 알림 ──
    title_symbol = Ticker.split('/')[0]
    try:
        telegram_alert.SendMessage(
            f"🎯 *[{title_symbol}_총 수익 목표 달성 - 전체 청산]*\n"
            f"━━━━━━━━━━━━━━━\n"
            f"🔁 사이클: `{현재회차}회차` 완료\n"
            f"💵 청산 시점가: `{현재가:.6f} USDT`\n"
            f"💰 청산 전 미실현: `{미실현:+.4f} USDT`\n"
            f"💼 청산 전 누적실현: `{누적실현:+.4f} USDT`\n"
            f"📌 청산 전 baseline: `{baseline:+.4f} USDT`\n"
            f"🎯 사이클 수익: `{사이클수익:+.4f} USDT` (목표 {TOTAL_PROFIT_TARGET_USDT})\n"
            f"📊 청산 전 총합: `{총합:+.4f} USDT`\n"
            f"📋 청산 차수: `{청산_차수_수}개`\n"
            f"💸 청산 시 실현손익: `{총_청산_손익:+.4f} USDT`\n"
            f"📌 청산 후 누적/새 baseline: `{청산후_누적:+.4f} USDT`\n"
            f"⚙️ 후속 동작: `{TOTAL_PROFIT_AFTER_LIQUIDATION}`\n"
            f"🕒 시간: `{지금().strftime('%Y-%m-%d %H:%M:%S')}`\n"
            f"━━━━━━━━━━━━━━━"
        )
    except Exception as e:
        print(f"[총수익 청산 알림 실패] {e}")

    # ── 6) 후속 동작 처리 ──
    if TOTAL_PROFIT_AFTER_LIQUIDATION == "STOP":
        총수익청산락_설정(
            Ticker,
            info=f"reason=STOP, cycle_pnl={사이클수익:.4f}, round={현재회차}, cleared={청산_차수_수}",
        )
        shutdown_requested = True
        print(f"⏹️ [후속동작=STOP] 봇 종료 진행. 락 파일 유지 → 재실행해도 매수 안 함.", flush=True)

    elif TOTAL_PROFIT_AFTER_LIQUIDATION == "PAUSE":
        총수익청산락_설정(
            Ticker,
            info=f"reason=PAUSE, cycle_pnl={사이클수익:.4f}, round={현재회차}, cleared={청산_차수_수}",
        )
        print(f"⏸️ [후속동작=PAUSE] 매매 정지. 모니터링/하트비트는 계속.", flush=True)

    elif TOTAL_PROFIT_AFTER_LIQUIDATION == "RESTART":
        # 다음 사이클은 청산 후 누적부터 다시 +목표 모아야 함 + 회차 +1
        다음회차 = 현재회차 + 1
        총수익_baseline_저장(Ticker, 청산후_누적, cycle_count=다음회차)
        print(
            f"🔄 [후속동작=RESTART] baseline={청산후_누적:+.4f} | "
            f"{현재회차}회차 → {다음회차}회차. 다음 폴링부터 1차 매수 재시작.",
            flush=True,
        )

    else:
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
    """종가 시계열로부터 RSI 값을 계산 (Wilder smoothing)."""
    if not isinstance(close_prices, pd.Series):
        close_prices = pd.Series(close_prices)
    if len(close_prices) < period + 1:
        return 0.0
    delta = close_prices.diff()
    gain = delta.clip(lower=0)
    loss = (-delta).clip(lower=0)
    # 첫 period개 변화량(인덱스 1..period) 단순평균 후, 그 다음 봉부터 평활
    avg_gain = float(gain.iloc[1:period + 1].mean())
    avg_loss = float(loss.iloc[1:period + 1].mean())
    rsi = 100.0 if avg_loss == 0 else 100.0 - (100.0 / (1.0 + (avg_gain / avg_loss)))
    for i in range(period + 1, len(close_prices)):
        avg_gain = (avg_gain * (period - 1) + float(gain.iloc[i])) / period
        avg_loss = (avg_loss * (period - 1) + float(loss.iloc[i])) / period
        rsi = 100.0 if avg_loss == 0 else 100.0 - (100.0 / (1.0 + (avg_gain / avg_loss)))
    return float(rsi)

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
# 3-State Switching Engine: 5일선(MA5) MTF 지지선 이탈 기반 추세 라이딩
# ===============================
_ma5_cache: Dict[str, Dict[str, dict]] = {}
_trend_riding_state: Dict[str, dict] = {}

def 이동평균_조회(Ticker: str, timeframe: str = "1d") -> Tuple[float, float, bool]:
    """
    지정 타임프레임(1d, 4h)의 5선(MA5) 계산 및 캐싱
    [반환] (ma_today, ma_yesterday, is_rising)
    """
    now = time.time()
    ticker_cache = _ma5_cache.setdefault(Ticker, {})
    cached = ticker_cache.get(timeframe)
    cache_ttl = MA_4H_CACHE_SECONDS if timeframe == "4h" else MA_CACHE_SECONDS

    if cached and (now - cached["time"] < cache_ttl):
        return cached["ma_today"], cached["ma_yesterday"], cached["is_rising"]

    try:
        ohlcv = bitgetX.fetch_ohlcv(Ticker, timeframe=timeframe, limit=MA_PERIOD + 2)
        if not ohlcv or len(ohlcv) < MA_PERIOD + 1:
            return (cached["ma_today"], cached["ma_yesterday"], cached["is_rising"]) if cached else (0.0, 0.0, False)

        # 직전 완결봉 기준 5개 봉 종가 평균 (인덱스 -6 ~ -2)
        ma_yesterday = sum(float(c[4]) for c in ohlcv[-(MA_PERIOD + 1):-1]) / float(MA_PERIOD)
        # 현재 진행 중 봉 포함 5개 봉 종가 평균 (인덱스 -5 ~ -1)
        ma_today = sum(float(c[4]) for c in ohlcv[-MA_PERIOD:]) / float(MA_PERIOD)

        is_rising = ma_today > ma_yesterday
        ticker_cache[timeframe] = {
            "time": now,
            "is_rising": is_rising,
            "ma_today": ma_today,
            "ma_yesterday": ma_yesterday,
        }
        return ma_today, ma_yesterday, is_rising
    except Exception as e:
        print(f"[MA5({timeframe}) 계산 실패] {Ticker} / {e}")
        if cached:
            return cached["ma_today"], cached["ma_yesterday"], cached["is_rising"]
        return 0.0, 0.0, False

def is_ma5_rising(Ticker: str) -> bool:
    """일봉 5일선 상승 여부 (하위 호환)"""
    _, _, is_rising = 이동평균_조회(Ticker, timeframe=MA_TIMEFRAME_1D)
    return is_rising

def 현재_MA5_정보(Ticker: str, timeframe: str = "1d") -> Tuple[float, float, bool]:
    """캐시된 ma_today, ma_yesterday, is_rising 반환 (모니터링용)"""
    return 이동평균_조회(Ticker, timeframe=timeframe)

def 추세라이딩_상태체크(Ticker: str, 현재가: float) -> Tuple[bool, bool]:
    """
    MTF(일봉 + 4시간봉) 지지선 이탈 필터가 적용된 추세 라이딩(상태 2) 활성 여부 판정
    [반환] (라이딩_중: bool, 방금_꺾임: bool)
    
    1. 진입 조건:
       - 일봉 5일선 상승 중 (ma5_today > ma5_yesterday)
       - AND 현재가가 일봉 5MA 지지선 이상 (현재가 >= ma_1d_today) -> 진입 즉시 튕겨나감 방지
    
    2. 종료(꺾임) 조건 (MTF 2단계 검증):
       - 1단계: 현재가 < 일봉 5MA * (1 - TREND_EXIT_BUFFER_PCT/100)  [기본 0.995]
       - 2단계 (1단계 발생 시에만 4시간봉 즉시 조회):
               현재가 < 4시간 5MA * (1 - TREND_EXIT_BUFFER_PCT/100) [기본 0.995]
       - 1단계 및 2단계 동시 이탈 시 상태 2 종료 -> 방금_꺾임=True 반환 (청산/수비 검사 진입)
    """
    now = time.time()
    state = _trend_riding_state.setdefault(Ticker, {"is_riding": False, "entered_at": 0.0})
    ma_1d_today, ma_1d_yesterday, rising_1d = 이동평균_조회(Ticker, timeframe=MA_TIMEFRAME_1D)
    title_symbol = Ticker.split('/')[0]

    buffer_multiplier = 1.0 - (TREND_EXIT_BUFFER_PCT / 100.0)

    if not state["is_riding"]:
        # 진입: 일봉 5일선 상승 + 현재가가 일봉 5일선 지지선 이상
        if rising_1d and ma_1d_today > 0 and 현재가 >= ma_1d_today:
            state["is_riding"] = True
            state["entered_at"] = now
            msg = (
                f"🚀 *[{title_symbol}_추세 라이딩 진입]*\n"
                f"━━━━━━━━━━━━━━━\n"
                f"📊 *종목*: `{Ticker}`\n"
                f"📈 *일봉 5MA*: 상승 전환 (`{ma_1d_today:.4f} > {ma_1d_yesterday:.4f}`)\n"
                f"📌 *현재가*: `{현재가:.4f} USDT`\n"
                f"🛡️ *종료 조건*: 일봉 및 4시간봉 5MA -{TREND_EXIT_BUFFER_PCT}% 동시 이탈 시\n"
                f"🕒 *시간*: `{지금().strftime('%Y-%m-%d %H:%M:%S')}`\n"
                f"━━━━━━━━━━━━━━━"
            )
            telegram_alert.SendMessage(msg)
            print(
                f"🚀 [{title_symbol}] 일봉 5MA 상승 & 안착 -> 상태 2 [추세 라이딩] 진입 "
                f"(종료조건: 일봉+4시간 5MA -{TREND_EXIT_BUFFER_PCT}% 동시 이탈)",
                flush=True
            )
            return True, False
        return False, False
    else:
        # 이미 상태 2(라이딩) 중인 경우
        if ma_1d_today <= 0:
            return True, False

        # 1차 검증: 일봉 5MA 지지선 이탈 (-0.5%) 확인
        이탈기준_1d = ma_1d_today * buffer_multiplier
        if 현재가 < 이탈기준_1d:
            # ➔ 1차 이탈 발생 시에만 4시간봉 즉시 조회 (Lazy Evaluation)
            ma_4h_today, _, _ = 이동평균_조회(Ticker, timeframe=MA_TIMEFRAME_4H)
            if ma_4h_today > 0:
                이탈기준_4h = ma_4h_today * buffer_multiplier
                # 2차 검증: 4시간봉 5MA 지지선도 동시 이탈했는가?
                if 현재가 < 이탈기준_4h:
                    state["is_riding"] = False
                    state["entered_at"] = 0.0
                    print(
                        f"📉 [{title_symbol}] 일봉/4시간봉 5MA 지지선(-{TREND_EXIT_BUFFER_PCT}%) 동시 이탈 감지 "
                        f"(현재가 {현재가:.4f} < 일봉기준 {이탈기준_1d:.4f} & 4시간기준 {이탈기준_4h:.4f}) -> 추세 라이딩 종료",
                        flush=True
                    )
                    return False, True

        # 지지선 유지 중: 라이딩 지속
        return True, False


# ===============================
# 매수 실행 함수
# ===============================
def 매수실행(Ticker: str, 현재가: float):
    """
    그리드 매수 로직
    - 1차: 즉시매수허용 또는 RSI 조건 충족 시 진입
    - 2차+: 직전 차수 매수가 기준 % 하락 시 진입
      (쿨링 해제 후 rebase 중이면 하이브리드:
       현재가 > A → 즉시, ≤ A → 저점 대비 |매수간격|% 반등)
    - 한 번의 폴링에서는 최대 1개 차수만 진입 (break)
    - 다수 매수 시 쿨링 + 재개 후 하이브리드 rebase
    - 총 수익 청산 락 / 쿨링 중이면 매수 차단
    """
    if 총수익청산락_여부(Ticker):
        return

    # 공백 복원 예약 매수 목표가 도달 감시 (도달 시 시장가 매수 및 동적 정렬)
    if 공백복원_예약매수_감시(Ticker, 현재가):
        return

    # 공백 복원 예약 대기 여부 확인 (예약 대기 중 일반 루프의 공백 차수 침범 방지)
    gap_state = 공백복원_상태로드(Ticker) if GAP_RECONSTRUCTION_ENABLE else {}
    has_pending_gap_orders = bool(gap_state.get("orders", []))
    gap_base_level = int(gap_state.get("base_level", 0))

    # 쿨링 중이면 매수만 차단 (매도는 별도)
    if 쿨링여부(Ticker):
        return

    # 쿨링 만료 직후: 해제 기준가 A + 저점 초기화
    if 쿨링_rebase_적용(Ticker, 현재가):
        title_symbol = Ticker.split('/')[0]
        msg = (
            f"⏳ *[{title_symbol}_매수 쿨링 해제]*\n"
            f"━━━━━━━━━━━━━━━\n"
            f"📊 *종목*: `{Ticker}`\n"
            f"📌 *기준가 A*: `{현재가:.6f} USDT`\n"
            f"📈 *> A*: 즉시 다음 차수 1회\n"
            f"📉 *≤ A*: 저점 트레일 → 저점 대비 |매수간격|% 반등 시 1회\n"
            f"🕒 *시간*: `{지금().strftime('%Y-%m-%d %H:%M:%S')}`\n"
            f"━━━━━━━━━━━━━━━"
        )
        telegram_alert.SendMessage(msg)
        print(
            f"⏳ 쿨링 해제 - {Ticker} 기준가 A={현재가:.{PRICE_LOG_DECIMALS}f} "
            f"(하이브리드: >A 즉시 / ≤A 저점 반등)",
            flush=True
        )

    # ≤A 구간: 저점 하향 갱신
    쿨링_rebase_저점갱신(Ticker, 현재가)

    file_name = trade_config[Ticker]["file_name"]
    df = 매매기록불러오기(file_name)
    if 현재가 <= 0:
        return

    rebase_A = 쿨링_rebase_앵커조회(Ticker)
    rebase_mode = ""

    open_levels = df[df["매도날짜"].isnull()]["차수"] if (df is not None and not df.empty and "매도날짜" in df.columns) else pd.Series(dtype=int)
    max_open_level = int(open_levels.max()) if not open_levels.empty else 0

    for 차수, config in trade_config[Ticker]["grid_config"].items():
        if not df[(df["차수"] == 차수) & (df["매도날짜"].isnull())].empty:
            continue

        # 공백 복원 예약이 대기 중일 때, 상단 미복원 공백 차수(현재 열려있는 최하단 차수보다 작은 미보유 차수)는 일반 하향 매수 차단
        # 단, 현재 열려있는 최하단 차수보다 큰 하향 차수(차수 > max_open_level)는 정상적인 하락 매수/더미 매수 진행
        if has_pending_gap_orders and max_open_level > 0 and 차수 < max_open_level:
            continue

        if 차수 == 1:
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
            prev_open = df[(df["차수"] == 차수 - 1) & (df["매도날짜"].isnull())]
            if prev_open.empty:
                # 잔고 부족으로 df에 기록되지 못한 직전 가상 차수 시도가 있는지 확인
                v_prev = 잔고부족_가상차수_조회(Ticker, 차수 - 1)
                if v_prev is None:
                    continue
                기준가 = float(v_prev["price"])
            else:
                기준가 = float(prev_open.iloc[-1]["매수가"])

            # 일반 매수 시(rebase_A가 아닐 때), 해당 차수가 최근 60초 내 이미 잔고부족으로 시도되었으면 중복 시도 스킵
            if rebase_A is None and 잔고부족_최근기록_여부(Ticker, 차수, cooldown_sec=60):
                continue

            if rebase_A is not None:
                A = float(rebase_A)
                if 현재가 > A:
                    가격조건 = True
                    rebase_mode = "즉시(>A)"
                else:
                    low = 쿨링_rebase_저점조회(Ticker)
                    if low is None:
                        continue
                    bounce = abs(float(config["매수간격"]))
                    target_price = float(low) * (1 + bounce / 100.0)
                    가격조건 = 현재가 >= target_price
                    rebase_mode = f"저점반등(low={float(low):.6f}, tgt={target_price:.6f})"
                if not 가격조건:
                    continue
            else:
                target_price = 기준가 * (1 + float(config["매수간격"]) / 100.0)
                가격조건 = 현재가 <= target_price
                if not 가격조건:
                    continue

        삼음봉조건 = is_three_yin_candles(Ticker, THREE_YIN_CANDLE_TIMEFRAME)
        음봉전용 = 차수 in 음봉전용차수

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

        has_open = not df[df["매도날짜"].isnull()].empty
        used_rebase = (rebase_A is not None) and (차수 > 1 or has_open)

        # Rebase 진입 시 잔고 부족 사전 점검 (자금 고갈 시 비상 선손절 자가치유 롤다운 가동)
        if used_rebase and GAP_RECONSTRUCTION_ENABLE:
            free_usdt = 거래소_USDT_잔고조회()
            buy_cost = float(config["매수금액"])
            if free_usdt is not None and free_usdt < (buy_cost * 0.9):
                print(
                    f"⚠️ [비상 롤다운 감지] 가용 USDT({free_usdt:.2f}) < 매수필요금액({buy_cost:.2f}) "
                    f"➔ 선(先) 손절 비상 자가치유 롤다운 가동!",
                    flush=True
                )
                if 비상_증거금확보_쿨링공백복원_실행(Ticker, 현재가, 차수, config, rebase_mode):
                    break

        수량 = USDT금액을_수량으로변환(Ticker, float(config["매수금액"]), 현재가)
        if 수량 == 0.0:
            print(f"🚫 매수 미실행: 최소 수량 미만")
            continue
        try:
            order_result = bitgetX.create_market_buy_order(
                symbol=Ticker, amount=수량, params={"reduceOnly": False}
            )
        except ccxt.InsufficientFunds as e:
            if used_rebase and GAP_RECONSTRUCTION_ENABLE:
                print(f"⚠️ [Rebase 증거금 부족] 비상 선(先)손절 자가치유 롤다운 가동: {e}", flush=True)
                if 비상_증거금확보_쿨링공백복원_실행(Ticker, 현재가, 차수, config, rebase_mode):
                    break
            print(f"⚠️ [증거금/잔고 부족] {Ticker} {차수}차 매수 스킵 (포지션 감시 및 청산 루프는 정상 지속): {e}")
            is_new_dummy = 잔고부족_가상차수_조회(Ticker, 차수) is None
            dummy_ref_price = target_price if 차수 > 1 and "target_price" in locals() else 현재가
            잔고부족_시도기록(Ticker, 차수, dummy_ref_price)
            총시도수 = 최근_총매수시도_고유차수수(df, Ticker, BUY_COOLDOWN_WINDOW_SEC)
            if is_new_dummy:
                더미매수_텔레그램알림(Ticker, 차수, 현재가, config, 총시도수)
            # 오직 40분 내 5회 급락 시에만 발동! (완만 하락형 제외)
            if BUY_COOLDOWN_ENABLE:
                if 총시도수 >= BUY_COOLDOWN_BUY_COUNT and not 쿨링여부(Ticker):
                    쿨링시작(Ticker)
                    until = 지금() + timedelta(seconds=BUY_COOLDOWN_PAUSE_SEC)
                    title_symbol = Ticker.split('/')[0]
                    cool_msg = (
                        f"🚨 *[{title_symbol}_비상 증거금부족 매수 쿨링 발동]* 🚨\n"
                        f"━━━━━━━━━━━━━━━\n"
                        f"📊 *종목*: `{Ticker}`\n"
                        f"⚠️ *원인*: 증거금 부족 중 급락 연쇄 시도 감지\n"
                        f"🔢 *최근 {BUY_COOLDOWN_WINDOW_SEC // 60}분 고유차수*: `{총시도수}개` (기준 {BUY_COOLDOWN_BUY_COUNT}개)\n"
                        f"⏸ *매수 금지*: `{BUY_COOLDOWN_PAUSE_SEC // 60}분` (~{until.strftime('%H:%M:%S')})\n"
                        f"✅ *매도*: 계속 동작\n"
                        f"📌 *재개 시*: 비상 자가치유 롤다운 (가상 계산 ➔ 상위 손절 ➔ 증거금 확보 ➔ Rebase 실체결)\n"
                        f"🕒 *시간*: `{지금().strftime('%Y-%m-%d %H:%M:%S')}`\n"
                        f"━━━━━━━━━━━━━━━"
                    )
                    telegram_alert.SendMessage(cool_msg)
                    print(f"🚨 비상 쿨링 발동 - {Ticker} 잔고부족 급락 40분 내 {총시도수}회 시도 감지", flush=True)
            return  # 해당 폴링 매수만 스킵하고 메인 감시 루프는 정상 유지
        except ccxt.ExchangeError as e:
            err_msg = str(e).lower()
            if any(k in err_msg for k in ["insufficient", "margin", "balance"]):
                if used_rebase and GAP_RECONSTRUCTION_ENABLE:
                    print(f"⚠️ [Rebase 거래소 잔고 부족] 비상 선(先)손절 자가치유 롤다운 가동: {e}", flush=True)
                    if 비상_증거금확보_쿨링공백복원_실행(Ticker, 현재가, 차수, config, rebase_mode):
                        break
                print(f"⚠️ [거래소 잔고 부족] {Ticker} {차수}차 매수 스킵 (감시 루프 지속): {e}")
                is_new_dummy = 잔고부족_가상차수_조회(Ticker, 차수) is None
                dummy_ref_price = target_price if 차수 > 1 and "target_price" in locals() else 현재가
                잔고부족_시도기록(Ticker, 차수, dummy_ref_price)
                총시도수 = 최근_총매수시도_고유차수수(df, Ticker, BUY_COOLDOWN_WINDOW_SEC)
                if is_new_dummy:
                    더미매수_텔레그램알림(Ticker, 차수, 현재가, config, 총시도수)
                if BUY_COOLDOWN_ENABLE:
                    if 총시도수 >= BUY_COOLDOWN_BUY_COUNT and not 쿨링여부(Ticker):
                        쿨링시작(Ticker)
                        until = 지금() + timedelta(seconds=BUY_COOLDOWN_PAUSE_SEC)
                        title_symbol = Ticker.split('/')[0]
                        cool_msg = (
                            f"🚨 *[{title_symbol}_비상 증거금부족 매수 쿨링 발동]* 🚨\n"
                            f"━━━━━━━━━━━━━━━\n"
                            f"📊 *종목*: `{Ticker}`\n"
                            f"⚠️ *원인*: 증거금 부족 중 급락 연쇄 시도 감지\n"
                            f"🔢 *최근 {BUY_COOLDOWN_WINDOW_SEC // 60}분 고유차수*: `{총시도수}개` (기준 {BUY_COOLDOWN_BUY_COUNT}개)\n"
                            f"⏸ *매수 금지*: `{BUY_COOLDOWN_PAUSE_SEC // 60}분` (~{until.strftime('%H:%M:%S')})\n"
                            f"✅ *매도*: 계속 동작\n"
                            f"📌 *재개 시*: 비상 자가치유 롤다운 (가상 계산 ➔ 상위 손절 ➔ 증거금 확보 ➔ Rebase 실체결)\n"
                            f"🕒 *시간*: `{지금().strftime('%Y-%m-%d %H:%M:%S')}`\n"
                            f"━━━━━━━━━━━━━━━"
                        )
                        telegram_alert.SendMessage(cool_msg)
                        print(f"🚨 비상 쿨링 발동 - {Ticker} 잔고부족 급락 40분 내 {총시도수}회 시도 감지", flush=True)
                return
            print(f"[매수 거래소 오류] {e}")
            continue
        except Exception as e:
            print(f"[매수 주문 실패] {e}")
            continue

        실제체결가, 실제체결수량, 실제체결금액 = 체결가_조회(
            Ticker, order_result, fallback_price=현재가, fallback_qty=수량
        )
        매수슬리피지 = 슬리피지_계산(현재가, 실제체결가)
        잔고부족_트래커_클리어(Ticker)

        # rebase로 진입했으면 앵커/저점 제거 → 이후 차수는 직전 체결가 기준
        if used_rebase:
            쿨링_rebase_클리어(Ticker)
            rebase_A = None
            if GAP_RECONSTRUCTION_ENABLE:
                쿨링공백복원_실행(
                    Ticker,
                    실제체결가,
                    실제체결수량,
                    실제체결금액,
                    매수슬리피지
                )
                break

        new_row = pd.DataFrame([{
            "종목": Ticker,
            "차수": 차수,
            "매수날짜": 지금(),
            "매수가": 실제체결가,
            "매수수량": 실제체결수량,
            "매수금액": 실제체결금액,
            "매수시현재가": 현재가,
            "매수슬리피지%": round(매수슬리피지, 4),
            "최고수익률": 0.0,
        }])
        df = pd.concat([df, new_row], ignore_index=True)
        매매기록저장(df, file_name)

        title_symbol = Ticker.split('/')[0]
        다음차 = 차수 + 1
        다음매수가 = 0.0
        다음목표가설명 = "없음"
        if 다음차 in trade_config[Ticker]["grid_config"]:
            다음간격 = float(trade_config[Ticker]["grid_config"][다음차]["매수간격"])
            다음매수가 = 실제체결가 * (1 + 다음간격 / 100.0)
            다음목표가설명 = f"{다음매수가:.6f} USDT (직전 체결가 기준)"

        if TRAILING_ENABLE:
            매도안내 = (
                f"🎯 *트레일링 매도*: 1차보존`+{TRAIL_STAGE1_PCT}%` / "
                f"2차활성`+{TRAIL_STAGE2_PCT}%`(고점`-{TRAIL_DISTANCE_PCT}%`)\n"
            )
        else:
            매도안내 = (
                f"🎯 *목표 매도가*: "
                f"`{실제체결가 * (1 + config['매도간격'] / 100.0):.6f} USDT`\n"
            )

        rebase_note = f" (rebase {rebase_mode})" if used_rebase and rebase_mode else (
            " (rebase 하이브리드 진입)" if used_rebase else ""
        )
        이유 = f"가격조건={가격조건}, 삼음봉={삼음봉조건}, 음봉전용={음봉전용}{rebase_note}"
        손익블록 = ""
        if TOTAL_PROFIT_TARGET_ENABLE:
            try:
                손익블록 = 총수익_텔레그램라인(Ticker, 현재가)
            except Exception as e:
                print(f"[매수 텔레그램 - 총수익 추가 실패] {e}")

        총시도수 = 최근_총매수시도_고유차수수(df, Ticker, BUY_COOLDOWN_WINDOW_SEC)
        보유금라인 = 보유금_텔레그램라인(Ticker, df)
        message = (
            f"🚨*[{title_symbol}_그리드롱 진입 완료]*🚨\n"
            f"━━━━━━━━━━━━━━━\n"
            f"📊 *종목*: `{Ticker}`\n"
            f"🔢 *차수*: `{차수}`\n"
            f"💵 *시장가*: `{현재가:.6f} USDT`\n"
            f"🎯 *실제 체결가*: `{실제체결가:.6f} USDT`\n"
            f"📐 *슬리피지*: `{매수슬리피지:+.4f}%`\n"
            f"📦 *체결수량*: `{실제체결수량}`\n"
            f"💰 *체결금액*: `{실제체결금액:.6f} USDT`\n"
            f"🔢 *최근 {BUY_COOLDOWN_WINDOW_SEC // 60}분 누적*: `{총시도수} / {BUY_COOLDOWN_BUY_COUNT}개` ({BUY_COOLDOWN_BUY_COUNT}개 도달 시 쿨링)\n"
            f"{손익블록}"
            f"{보유금라인}"
            f"🕒 *시간*: `{지금().strftime('%Y-%m-%d %H:%M:%S')}`\n"
            f"━━━━━━━━━━━━━━━"
        )

        telegram_alert.SendMessage(message)

        print(
            f"✅ 매수 완료 - {Ticker} 차수 {차수} | "
            f"체결가 {실제체결가:.{PRICE_LOG_DECIMALS}f} "
            f"(시장가 {현재가:.{PRICE_LOG_DECIMALS}f}, 슬리피지 {매수슬리피지:+.4f}%)"
            f"{' | rebase ' + rebase_mode if used_rebase else ''}",
            flush=True
        )

        # 급락 연쇄 매수 쿨링 체크 (여기 도달 시 쿨링 중이 아님)
        if BUY_COOLDOWN_ENABLE:
            if 총시도수 >= BUY_COOLDOWN_BUY_COUNT:
                쿨링시작(Ticker)
                until = 지금() + timedelta(seconds=BUY_COOLDOWN_PAUSE_SEC)
                cool_msg = (
                    f"🧊 *[{title_symbol}_매수 쿨링 발동]*\n"
                    f"━━━━━━━━━━━━━━━\n"
                    f"📊 *종목*: `{Ticker}`\n"
                    f"🔢 *최근 {BUY_COOLDOWN_WINDOW_SEC // 60}분 고유차수*: `{총시도수}개` "
                    f"(기준 {BUY_COOLDOWN_BUY_COUNT}개)\n"
                    f"⏸ *매수 금지*: `{BUY_COOLDOWN_PAUSE_SEC // 60}분` "
                    f"(~{until.strftime('%H:%M:%S')})\n"
                    f"✅ *매도*: 계속 동작\n"
                    f"📌 *재개 시*: >A 즉시 / ≤A 저점 반등 하이브리드\n"
                    f"🕒 *시간*: `{지금().strftime('%Y-%m-%d %H:%M:%S')}`\n"
                    f"━━━━━━━━━━━━━━━"
                )
                telegram_alert.SendMessage(cool_msg)
                print(
                    f"🧊 쿨링 발동 - {Ticker} 최근 고유차수 {총시도수}개 / "
                    f"{BUY_COOLDOWN_PAUSE_SEC // 60}분 매수 금지",
                    flush=True
                )

        break


# ===============================
# 매도 실행 함수
# ===============================
def 매도실행(Ticker: str, 현재가: float, riding_mode: str = None) -> bool:
    """
    그리드 매도 로직
    - TRAILING_ENABLE: 2단계 하이브리드 트레일링
      · 1단계 (+TRAIL_STAGE1_PCT ~ +TRAIL_STAGE2_PCT 미만): 0.7% 이하 하회 시 보존 매도
      · 2단계 (+TRAIL_STAGE2_PCT 이상): 고점 대비 TRAIL_DISTANCE_PCT 하락 시 청산
    - TRAILING_ENABLE=False: 매도간격 고정 목표가 청산
    - riding_mode: 추세 라이딩 중 매도 제어 모드
      · None     : 일반 그리드 모드 (TRAILING_ENABLE 기준 그대로 적용)
      · "STAGE1_ONLY": 라이딩 중, 최고수익률이 +0.7% 이상 찍은 차수가 +0.7% 이하로 하회 시에만 매도
      · "ALL"    : 라이딩 중, 1단계(0.7%) + 2단계(고점-0.2%) 모두 허용
    - 한 폴링에서 여러 차수가 동시에 청산될 수 있음
    - 반환값: 매도 발생 시 True (이번 폴링에서 매수는 스킵)
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
    고점갱신됨 = False

    누적_추적기 = 직전누적실현손익(df)

    for idx, row in 미청산.sort_values(by=["차수"], ascending=False).iterrows():
        차수 = int(row["차수"])
        매수가 = float(row["매수가"])
        매수수량 = float(row["매수수량"])
        매수금액 = float(row["매수금액"])
        config = trade_config[Ticker]["grid_config"].get(차수)
        if not config:
            continue
        if 매수가 <= 0:
            continue

        현재수익률 = (현재가 / 매수가 - 1.0) * 100.0
        try:
            raw_peak = row["최고수익률"] if "최고수익률" in row.index else None
            기존고점 = float(raw_peak) if raw_peak is not None and pd.notna(raw_peak) else 0.0
        except (TypeError, ValueError, KeyError):
            기존고점 = 0.0
        최고수익률 = max(기존고점, 현재수익률)
        if 최고수익률 != 기존고점:
            df.at[idx, "최고수익률"] = round(최고수익률, 4)
            고점갱신됨 = True

        매도조건 = False
        트레일매도선 = 0.0
        트레일단계 = 0
        라이딩_보존매도 = False  # 라이딩 중 이익 보존 매도 여부 플래그
        if TRAILING_ENABLE:
            if 최고수익률 >= TRAIL_STAGE2_PCT:
                # 2단계: 0.9% 이상 도달 후 다이내믹 트레일링 (고점 - 0.2%)
                트레일단계 = 2
                트레일매도선 = 최고수익률 - TRAIL_DISTANCE_PCT
                if riding_mode == "STAGE1_ONLY":
                    # 라이딩 STAGE1_ONLY: 2단계 고점 트레일링은 홀딩하되,
                    # 최소 이익 보존선(+0.7%) 아래로 내려오면 이익 반납 방지 매도
                    매도조건 = (현재수익률 <= TRAIL_STAGE1_PCT)
                    if 매도조건:
                        라이딩_보존매도 = True
                        트레일매도선 = TRAIL_STAGE1_PCT  # 보존선 기준으로 표기
                else:
                    # 일반 모드(None) 또는 ALL 모드: 기존 2단계 다이내믹 트레일링
                    매도조건 = (현재수익률 <= 트레일매도선)
            elif 최고수익률 >= TRAIL_STAGE1_PCT:
                # 1단계: 0.7% 이상 도달 후 1차 이익 보존선 (0.7% 하회 시 매도)
                트레일단계 = 1
                트레일매도선 = TRAIL_STAGE1_PCT
                매도조건 = (현재수익률 <= 트레일매도선)
        else:
            목표가 = 매수가 * (1 + float(config["매도간격"]) / 100.0)
            매도조건 = 현재가 >= 목표가

        if not 매도조건:
            continue

        try:
            order_result = bitgetX.create_market_sell_order(
                symbol=Ticker, amount=매수수량, params={"reduceOnly": True}
            )
        except Exception as e:
            print(f"[매도 주문 실패] {e}")
            continue

        실제체결가, 실제체결수량, 실제체결금액 = 체결가_조회(
            Ticker, order_result, fallback_price=현재가, fallback_qty=매수수량
        )
        매도슬리피지 = 슬리피지_계산(현재가, 실제체결가)

        매도금액 = 실제체결금액
        수수료 = 매도금액 * float(trade_config[Ticker]["수수료율"])
        실현손익 = 매도금액 - 매수금액 - 수수료
        수익률 = (실현손익 / 매수금액) * 100.0 if 매수금액 else 0.0

        누적_추적기 += 실현손익
        현재누적손익 = 누적_추적기

        df.at[idx, "최고수익률"] = round(최고수익률, 4)
        df.at[idx, "매도날짜"] = 지금()
        df.at[idx, "매도가"] = 실제체결가
        df.at[idx, "매도수량"] = 실제체결수량
        df.at[idx, "매도금액"] = 매도금액
        df.at[idx, "매도시현재가"] = 현재가
        df.at[idx, "매도슬리피지%"] = round(매도슬리피지, 4)
        df.at[idx, "수수료"] = 수수료
        df.at[idx, "실현손익"] = 실현손익
        df.at[idx, "수익률"] = 수익률
        df.at[idx, "누적실현손익"] = 현재누적손익
        df.at[idx, "청산유형"] = "그리드익절"
        매매기록저장(df, file_name)

        if TRAILING_ENABLE:
            if 라이딩_보존매도:
                단계_태그 = "1차-라이딩보존"
            elif 트레일단계 > 0:
                단계_태그 = f"{트레일단계}차"
            else:
                단계_태그 = "트레일"
            트레일라인 = (
                f"📍 *트레일링({단계_태그})*: 고점`{최고수익률:.4f}%` / "
                f"매도선`{트레일매도선:.4f}%` / 청산시`{현재수익률:.4f}%`\n"
            )
        else:
            트레일라인 = (
                f"🎯 *고정 목표가*: `{매수가 * (1 + float(config['매도간격']) / 100.0):.6f} USDT`\n"
            )

        손익블록 = ""
        if TOTAL_PROFIT_TARGET_ENABLE:
            try:
                손익블록 = 총수익_텔레그램라인(Ticker, 현재가)
            except Exception as e:
                print(f"[매도 텔레그램 - 총수익 추가 실패] {e}")

        보유금라인 = 보유금_텔레그램라인(Ticker, df)

        message = (
            f"🌟[{title_symbol}_그리드롱 청산 완료]🌟\n"
            f"━━━━━━━━━━━━━━━\n"
            f"📊 *종목*: `{Ticker}`\n"
            f"🔢 *차수*: `{차수}`\n"
            f"💵 *시장가*: `{현재가:.6f} USDT`\n"
            f"🎯 *실제 체결가*: `{실제체결가:.6f} USDT`\n"
            f"📐 *슬리피지*: `{매도슬리피지:+.4f}%`\n"
            f"📦 *체결수량*: `{실제체결수량}`\n"
            f"💰 *매도금액*: `{매도금액:.6f} USDT`\n"
            f"{트레일라인}"
            f"✨ *【 실현 손익 】*: *`{실현손익:+.4f} USDT`* (`{수익률:+.4f}%`)\n"
            f"{손익블록}"
            f"{보유금라인}"
            f"🕒 *시간*: `{지금().strftime('%Y-%m-%d %H:%M:%S')}`\n"
            f"━━━━━━━━━━━━━━━"
        )

        telegram_alert.SendMessage(message)

        print(
            f"💸 매도 완료 - {Ticker} 차수 {차수} | "
            f"체결가 {실제체결가:.{PRICE_LOG_DECIMALS}f} "
            f"(시장가 {현재가:.{PRICE_LOG_DECIMALS}f}, 슬리피지 {매도슬리피지:+.4f}%) | "
            f"손익 {실현손익:+.4f} USDT | "
            f"수익률 {수익률:+.4f}% | "
            f"고점 {최고수익률:+.4f}% | "
            f"누적 {현재누적손익:+.4f}",
            flush=True
        )

        청산된_차수_수 += 1
        총실현손익 += 실현손익
        총수익률합 += 수익률
        청산된_포지션수 += 1
        매도발생 = True

        # [보완 3] 쿨링 전 상위 차수(매수가 >= P_prev) 익절 매도 체결 시 잔여 공백 예약 즉시 클리어 및 미매도 점검
        try:
            st = 공백복원_상태로드(Ticker)
            st_orders = st.get("orders", [])
            st_prev = float(st.get("prev_price", 0.0))
            if st_orders and st_prev > 0 and 매수가 >= (st_prev * 0.999):
                cleared_n = len(st_orders)
                사유 = f"쿨링 전 상위 {차수}차수(매수가 {매수가:.4f} USDT) 익절 매도 ➔ 잔여 예약 {cleared_n}개 자동 클리어"
                공백복원_종료_미매도차수_점검(Ticker, 사유=사유, 현재가=현재가)
        except Exception as e:
            print(f"⚠️ [공백 복원 매도 연동 오류] {e}", flush=True)

    if 고점갱신됨 and not 매도발생:
        매매기록저장(df, file_name)

    if 청산된_포지션수 >= 2:
        평균수익률 = 총수익률합 / 청산된_포지션수
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
# ======================================================
# 📊 일일 정산 (LS_Grid.py v4.1 방식 이식)
# ======================================================
# - 하루 1회, SETTLEMENT_HOUR_KST 시각에 직전 24시간 거래를 집계
# - 종목별 엑셀에 "일일정산" 시트를 만들어 날짜별 1행 누적
# - 텔레그램으로 종목별 상세 + 전체 합계 1건 발송
# - 10분마다 오던 하트비트를 대체한다 (알림 하루 1건)
SETTLEMENT_SHEET = "일일정산"
SETTLEMENT_HOUR_KST = 9          # KST 09:00 기준으로 하루를 끊는다
SETTLEMENT_COLUMNS = [
    "정산일", "매수건수", "매수수량", "매수금액(USDT)",
    "매도건수", "매도수량", "매도금액(USDT)",
    "그리드익절(USDT)", "공백손절(USDT)",
    "수수료(USDT)", "실현수익(USDT)", "누적수익(USDT)",
]

_last_settlement_date = None     # 중복 발송 방지 (메모리)
SETTLEMENT_STATE_FILE = "/var/autobot/bitget_grid/.grid_bitget_last_settlement"  # 재기동 후에도 유지


def _kst_now():
    return datetime.now(KST)


def _정산_데이터_집계(df: pd.DataFrame, start_dt: datetime, end_dt: datetime) -> dict:
    """[start_dt, end_dt) 구간 거래 집계 (메모리 DataFrame 대상)."""
    empty_res = {
        "매수건수": 0, "매수수량": 0.0, "매수금액": 0.0,
        "매도건수": 0, "매도수량": 0.0, "매도금액": 0.0,
        "그리드건수": 0, "그리드수량": 0.0, "그리드금액": 0.0,
        "그리드수익": 0.0, "그리드수수료": 0.0,
        "공백손절건수": 0, "공백손절수량": 0.0, "공백손절금액": 0.0,
        "공백실현손실": 0.0, "공백수수료": 0.0,
        "수수료": 0.0, "실현수익": 0.0,
        "거래일수": 0, "거래일목록": set(),
    }
    if df is None or df.empty:
        return empty_res

    b = pd.to_datetime(df.get("매수날짜"), errors="coerce")
    s = pd.to_datetime(df.get("매도날짜"), errors="coerce")
    cs = start_dt.replace(tzinfo=None) if start_dt else datetime.min
    ce = end_dt.replace(tzinfo=None) if end_dt else datetime.max

    buy_mask = b.notna() & (b >= cs) & (b < ce)
    sell_mask = s.notna() & (s >= cs) & (s < ce)

    buy_df = df[buy_mask]
    sell_df = df[sell_mask]

    # 구간 내 실제 거래(매수 또는 매도)가 발생한 고유 일자 집계
    b_dates = set(b[buy_mask].dt.date)
    s_dates = set(s[sell_mask].dt.date)
    trade_dates = b_dates.union(s_dates)
    trading_days = len(trade_dates)

    def _sum(d, col):
        return round(float(pd.to_numeric(d.get(col), errors="coerce").fillna(0).sum()), 6)

    # 🟢 순수 그리드 익절 vs 🔴 공백복원 손절 분리
    # - 청산유형 컬럼이 명시되어 있으면 우선 참조
    # - 과거 데이터이거나 결측치인 경우: 실현손익 < 0 이면 공백복원손절, >= 0 이면 그리드익절
    pnl_series = pd.to_numeric(sell_df.get("실현손익"), errors="coerce").fillna(0.0)
    if "청산유형" in sell_df.columns:
        type_series = sell_df["청산유형"].astype(str)
        is_gap_cut = (type_series == "공백복원손절") | (sell_df["청산유형"].isna() & (pnl_series < 0))
    else:
        is_gap_cut = pnl_series < 0

    grid_df = sell_df[~is_gap_cut]
    gap_df = sell_df[is_gap_cut]

    grid_pnl = _sum(grid_df, "실현손익")
    gap_pnl = _sum(gap_df, "실현손익")
    tot_pnl = _sum(sell_df, "실현손익")
    tot_fee = _sum(sell_df, "수수료")

    return {
        "매수건수": int(len(buy_df)), "매수수량": _sum(buy_df, "매수수량"),
        "매수금액": _sum(buy_df, "매수금액"),
        "매도건수": int(len(sell_df)), "매도수량": _sum(sell_df, "매도수량"),
        "매도금액": _sum(sell_df, "매도금액"),
        "그리드건수": int(len(grid_df)), "그리드수량": _sum(grid_df, "매도수량"),
        "그리드금액": _sum(grid_df, "매도금액"), "그리드수익": grid_pnl,
        "그리드수수료": _sum(grid_df, "수수료"),
        "공백손절건수": int(len(gap_df)), "공백손절수량": _sum(gap_df, "매도수량"),
        "공백손절금액": _sum(gap_df, "매도금액"), "공백실현손실": gap_pnl,
        "공백수수료": _sum(gap_df, "수수료"),
        "수수료": tot_fee, "실현수익": tot_pnl,
        "거래일수": trading_days, "거래일목록": trade_dates,
    }


def _정산_집계(file_path: str, start_dt: datetime, end_dt: datetime) -> dict:
    """[start_dt, end_dt) 구간 거래 집계. 매수=매수날짜, 매도=매도날짜 기준 (하위 호환)."""
    try:
        df = pd.read_excel(file_path, engine="openpyxl", sheet_name=0)
    except Exception:
        return None
    if df.empty:
        return None
    return _정산_데이터_집계(df, start_dt, end_dt)


def _정산_시트갱신(file_path: str, 정산일: str, st: dict) -> float:
    """'일일정산' 시트에 1행 추가/갱신. 같은 날짜가 있으면 덮어쓴다. 누적수익 반환."""
    from openpyxl import load_workbook
    wb = load_workbook(file_path)
    if SETTLEMENT_SHEET not in wb.sheetnames:
        ws = wb.create_sheet(SETTLEMENT_SHEET)
        ws.append(SETTLEMENT_COLUMNS)
    ws = wb[SETTLEMENT_SHEET]

    # 기존 시트 헤더 확인 및 갱신 (신규 컬럼 반영)
    existing_headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    if existing_headers != SETTLEMENT_COLUMNS:
        for c, col_name in enumerate(SETTLEMENT_COLUMNS, start=1):
            ws.cell(row=1, column=c, value=col_name)

    cum_col_idx = SETTLEMENT_COLUMNS.index("누적수익(USDT)") + 1
    target_row, prev_cum = None, 0.0
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(row=r, column=1).value) == 정산일:
            target_row = r
            break
        try:
            val = ws.cell(row=r, column=cum_col_idx).value
            if val is not None:
                prev_cum = float(val)
        except Exception:
            pass

    cum = round(prev_cum + st["실현수익"], 6)
    row = [
        정산일, st["매수건수"], st["매수수량"], st["매수금액"],
        st["매도건수"], st["매도수량"], st["매도금액"],
        st.get("그리드수익", 0.0), st.get("공백실현손실", 0.0),
        st["수수료"], st["실현수익"], cum
    ]
    if target_row:
        for i, v in enumerate(row, start=1):
            ws.cell(row=target_row, column=i, value=v)
    else:
        ws.append(row)
    wb.save(file_path)
    return cum


def _텔레그램_안전발송(text: str):
    """메시지 길이가 긴 경우 텔레그램 4,096자 제한을 고려해 안전하게 분할 전송."""
    MAX_LEN = 3800
    if len(text) <= MAX_LEN:
        telegram_alert.SendMessage(text)
        return

    chunks = []
    current = []
    current_len = 0
    for line in text.split("\n"):
        if current_len + len(line) + 1 > MAX_LEN:
            chunks.append("\n".join(current))
            current = [line]
            current_len = len(line)
        else:
            current.append(line)
            current_len += len(line) + 1
    if current:
        chunks.append("\n".join(current))

    for chunk in chunks:
        try:
            telegram_alert.SendMessage(chunk)
        except Exception as e:
            print(f"[정산 분할알림 전송 실패] {e}", flush=True)


def _포맷_일일_정산_섹션(
    title: str,
    daily_items: list,
    day2_items: list,
    day3_items: list,
    tot_daily: dict,
    tot_day2: dict,
    tot_day3: dict,
    date_1d: str,
    date_2d: str,
    date_3d: str,
) -> list:
    """
    일일 정산(최근 3일: 1일전, 2일전, 3일전) 전용 포맷터.
    - 순실현 손익과 함께 🟢 순수 그리드 익절과 🔴 공백복원 손절을 분리 표시.
    """
    sec = ["━━━━━━━━━━━━━━━", title]

    d2_map = {item[0]: item[1] for item in day2_items}
    d3_map = {item[0]: item[1] for item in day3_items}

    empty_stat = {
        "실현수익": 0.0, "매도건수": 0, "수수료": 0.0,
        "그리드건수": 0, "그리드수익": 0.0, "공백손절건수": 0, "공백실현손실": 0.0
    }

    def _day_block(tag, dt_str, st):
        p = "+" if st["실현수익"] >= 0 else ""
        lines = [
            f"  [{tag}({dt_str})] 순실현 {p}{st['실현수익']:,.4f} USDT (수수료 {st['수수료']:,.4f} USDT)"
        ]
        g_cnt = st.get("그리드건수", 0)
        g_profit = st.get("그리드수익", 0.0)
        g_avg = (g_profit / g_cnt) if g_cnt > 0 else 0.0
        g_p = "+" if g_profit >= 0 else ""
        lines.append(f"    • 🟢 순수 그리드 익절: {g_p}{g_profit:,.4f} USDT ({g_cnt}건 | 건당 +{g_avg:,.4f} USDT)")

        c_cnt = st.get("공백손절건수", 0)
        c_loss = st.get("공백실현손실", 0.0)
        if c_cnt > 0:
            lines.append(f"    • 🔴 공백복원 손절:   {c_loss:,.4f} USDT ({c_cnt}건)")
        else:
            lines.append(f"    • ⚪ 공백복원 손절:   0.0000 USDT (0건)")
        return lines

    for name, st1, _ in daily_items:
        st2 = d2_map.get(name) or empty_stat
        st3 = d3_map.get(name) or empty_stat

        sec += [
            f"📌 {name}",
            f"  매수 {st1['매수건수']}건 | 수량 {st1['매수수량']:g} | {st1['매수금액']:,.2f} USDT",
            f"  매도 {st1['매도건수']}건 | 수량 {st1['매도수량']:g} | {st1['매도금액']:,.2f} USDT",
        ]
        sec += _day_block("1일전", date_1d, st1)
        sec += _day_block("2일전", date_2d, st2)
        sec += _day_block("3일전", date_3d, st3)
        sec.append("")

    # 다중 종목일 경우 합계 블록
    if len(daily_items) > 1:
        sec += [
            "📊 합계",
            f"  매수 {tot_daily['매수건수']}건 / {tot_daily['매수수량']:g} / {tot_daily['매수금액']:,.2f} USDT",
            f"  매도 {tot_daily['매도건수']}건 / {tot_daily['매도수량']:g} / {tot_daily['매도금액']:,.2f} USDT",
        ]
        sec += _day_block("1일전", date_1d, tot_daily)
        sec += _day_block("2일전", date_2d, tot_day2)
        sec += _day_block("3일전", date_3d, tot_day3)
        sec.append("")

    return sec


def 현재_롱포지션_현황_조회(Ticker: str) -> Optional[dict]:
    """
    거래소 fetch_positions로 Ticker의 현재 롱 포지션 상태를 상세 조회 (비트겟 앱 Positions 카드 정보).
    [반환] dict: {
        'symbol': str, 'side': str, 'contracts': float, 'entryPrice': float,
        'markPrice': float, 'notional': float, 'unrealizedPnl': float,
        'percentage': float (ROE %), 'margin': float, 'leverage': int,
        'marginMode': str,
    } 또는 None (API 오류 시)
    """
    try:
        positions = bitgetX.fetch_positions([Ticker])
        for pos in positions:
            sym = pos.get('symbol')
            # ccxt symbol 또는 raw info symbol 매칭
            raw_sym = pos.get('info', {}).get('symbol')
            target_clean = Ticker.replace('/', '').split(':')[0]
            if sym != Ticker and raw_sym != target_clean:
                continue

            contracts = float(pos.get('contracts', 0) or 0)
            if contracts <= 0:
                continue

            side = str(pos.get('side', '')).lower()
            hold_side = str(pos.get('info', {}).get('holdSide', '')).lower()
            if side != 'long' and hold_side != 'long':
                continue

            upnl = pos.get('unrealizedPnl')
            if upnl is None:
                upnl = pos.get('info', {}).get('unrealisedPnl', 0) or 0
            upnl = float(upnl)

            # ROE (%) 계산: ccxt percentage 또는 raw uplRatio
            roe = pos.get('percentage')
            if roe is None:
                upl_ratio = pos.get('info', {}).get('uplRatio')
                if upl_ratio is not None:
                    roe = float(upl_ratio) * 100.0
                else:
                    roe = 0.0
            else:
                roe = float(roe)

            entry_price = float(pos.get('entryPrice') or pos.get('info', {}).get('openPriceAvg') or 0.0)
            mark_price = float(pos.get('markPrice') or pos.get('info', {}).get('markPrice') or 0.0)
            notional = float(pos.get('notional') or pos.get('info', {}).get('notionalUsd') or 0.0)
            if notional == 0.0 and mark_price > 0 and contracts > 0:
                notional = mark_price * contracts

            margin = float(pos.get('initialMargin') or pos.get('collateral') or pos.get('info', {}).get('margin') or 0.0)
            leverage = int(pos.get('leverage') or pos.get('info', {}).get('leverage') or 1)
            margin_mode = str(pos.get('marginMode') or pos.get('marginType') or pos.get('info', {}).get('marginMode') or 'cross').capitalize()

            return {
                'symbol': Ticker,
                'side': 'Long',
                'contracts': contracts,
                'entryPrice': entry_price,
                'markPrice': mark_price,
                'notional': notional,
                'unrealizedPnl': upnl,
                'percentage': roe,
                'margin': margin,
                'leverage': leverage,
                'marginMode': margin_mode,
            }
        return {'symbol': Ticker, 'contracts': 0.0}
    except Exception as e:
        print(f"[현재 포지션 현황 조회 실패 - {Ticker}] {e}", flush=True)
        return None


def _포맷_포지션_현황_섹션() -> list:
    """
    일일 정산 메시지용 통합 롱 포지션 현황 섹션 (거래소 앱 Positions 탭 정보 포맷팅).
    미청산 차수들의 미실현 손익(Unrealized PnL) 및 ROE 표기.
    """
    sec = ["━━━━━━━━━━━━━━━", "5️⃣ [현재 포지션 및 미실현 손익]"]
    for ticker in 종목코드목록:
        name = ticker.split("/")[0]
        open_levels_cnt = 0
        try:
            cfg = trade_config.get(ticker) or {}
            fp = cfg.get("file_name")
            if fp and os.path.exists(fp):
                _df = pd.read_excel(fp, engine="openpyxl", sheet_name=0)
                open_levels_cnt = int(len(_df[_df["매도날짜"].isnull()]))
        except Exception:
            pass

        pos = 현재_롱포지션_현황_조회(ticker)
        open_info = f" (미청산 {open_levels_cnt}개 차수 보유)" if open_levels_cnt > 0 else ""
        if pos is None:
            sec += [
                f"📈 {name}{open_info}",
                f"  • 포지션: 조회 실패 (거래소 API 확인 필요)",
                "",
            ]
        elif pos.get("contracts", 0.0) <= 0:
            sec += [
                f"📈 {name}",
                f"  • 포지션: 보유 포지션 없음 (0.0 USDT)",
                "",
            ]
        else:
            pnl = pos["unrealizedPnl"]
            pnl_p = "+" if pnl >= 0 else ""
            roe = pos["percentage"]
            roe_p = "+" if roe >= 0 else ""
            sec += [
                f"📈 {name}{open_info}",
                f"  • 🧊 그리드 미실현 손익: {pnl_p}{pnl:,.4f} USDT (ROE {roe_p}{roe:,.2f}%)",
                f"  • 🎯 진입 평단가: {pos['entryPrice']:,.4f} | 현재 시장가: {pos['markPrice']:,.4f} USDT",
                f"  • 📦 포지션 규모: {pos['notional']:,.2f} USDT ({pos['contracts']:g} {name} | {pos['side']} {pos['leverage']}x {pos['marginMode']})",
                "",
            ]
    return sec


def _포맷_정산_섹션(
    title: str,
    item_list: list,
    total_stat: dict,
    profit_label: str = "실현수익",
    days: Optional[int] = None,
    show_divider: bool = True,
    item_days_map: Optional[dict] = None,
) -> list:
    """
    특정 구간(주간/월간/전체)의 상세 및 합계 블록 서식 생성.
    - 🟢 순수 그리드 익절
    - 🔴 공백복원 손절
    - ⚖️ 순실현손익(또는 누적수익)
    을 명확하게 분리 표기.
    """
    sec = []
    if show_divider:
        sec.append("━━━━━━━━━━━━━━━")
    sec.append(title)

    for name, st, _ in item_list:
        p = "+" if st["실현수익"] >= 0 else ""
        g_cnt = st.get("그리드건수", 0)
        g_profit = st.get("그리드수익", 0.0)
        g_avg = (g_profit / g_cnt) if g_cnt > 0 else 0.0

        it_days = (item_days_map.get(name) if item_days_map else None)
        if it_days is None:
            it_days = st.get("거래일수")
        if it_days is None:
            it_days = days

        if it_days and it_days > 0:
            g_day_avg = g_profit / it_days
            grid_line = (
                f"  • 🟢 순수 그리드 익절: +{g_profit:,.4f} USDT "
                f"({g_cnt}건 | 건당 +{g_avg:,.4f} USDT | 일평균 +{g_day_avg:,.4f} USDT ({it_days}일 기준))"
            )
        else:
            grid_line = f"  • 🟢 순수 그리드 익절: +{g_profit:,.4f} USDT ({g_cnt}건 | 건당 +{g_avg:,.4f} USDT)"

        c_cnt = st.get("공백손절건수", 0)
        c_loss = st.get("공백실현손실", 0.0)
        if c_cnt > 0:
            gap_line = f"  • 🔴 공백복원 손절:   {c_loss:,.4f} USDT ({c_cnt}건)"
        else:
            gap_line = f"  • ⚪ 공백복원 손절:   0.0000 USDT (0건)"

        sec += [
            f"📌 {name}",
            f"  매수 {st['매수건수']}건 | 수량 {st['매수수량']:g} | {st['매수금액']:,.2f} USDT",
            f"  매도 {st['매도건수']}건 | 수량 {st['매도수량']:g} | {st['매도금액']:,.2f} USDT",
            grid_line,
            gap_line,
            f"  • ⚖️ {profit_label}: {p}{st['실현수익']:,.4f} USDT (수수료 {st['수수료']:,.4f} USDT)",
        ]
        sec.append("")

    # 종목이 2개 이상일 때만 전체 합계 블록 출력
    if len(item_list) > 1:
        tp = total_stat["실현수익"]
        p = "+" if tp >= 0 else ""
        tot_g_cnt = total_stat.get("그리드건수", 0)
        tot_g_profit = total_stat.get("그리드수익", 0.0)
        tot_g_avg = (tot_g_profit / tot_g_cnt) if tot_g_cnt > 0 else 0.0

        tot_days = total_stat.get("거래일수") if (total_stat.get("거래일수") is not None and total_stat.get("거래일수") > 0) else days
        if tot_days and tot_days > 0:
            tot_g_day_avg = tot_g_profit / tot_days
            tot_grid_line = (
                f"  • 🟢 순수 그리드 익절: +{tot_g_profit:,.4f} USDT "
                f"({tot_g_cnt}건 | 건당 +{tot_g_avg:,.4f} USDT | 일평균 +{tot_g_day_avg:,.4f} USDT ({tot_days}일 기준))"
            )
        else:
            tot_grid_line = f"  • 🟢 순수 그리드 익절: +{tot_g_profit:,.4f} USDT ({tot_g_cnt}건 | 건당 +{tot_g_avg:,.4f} USDT)"

        tot_c_cnt = total_stat.get("공백손절건수", 0)
        tot_c_loss = total_stat.get("공백실현손실", 0.0)
        if tot_c_cnt > 0:
            tot_gap_line = f"  • 🔴 공백복원 손절:   {tot_c_loss:,.4f} USDT ({tot_c_cnt}건)"
        else:
            tot_gap_line = f"  • ⚪ 공백복원 손절:   0.0000 USDT (0건)"

        sec += [
            "📊 합계",
            f"  매수 {total_stat['매수건수']}건 / {total_stat['매수수량']:g} / {total_stat['매수금액']:,.2f} USDT",
            f"  매도 {total_stat['매도건수']}건 / {total_stat['매도수량']:g} / {total_stat['매도금액']:,.2f} USDT",
            tot_grid_line,
            tot_gap_line,
            f"  • 💰 {profit_label}: {p}{tp:,.4f} USDT (수수료 {total_stat['수수료']:,.4f} USDT)",
            "",
        ]
    return sec


def 일일정산_실행(정산일: str, start_dt: datetime = None, end_dt: datetime = None):
    """종목별 집계 (일일 3일치/주간/월간/전체) → 시트 갱신 → 텔레그램 통합 발송."""
    if end_dt is None:
        now = _kst_now()
        end_dt = now.replace(hour=SETTLEMENT_HOUR_KST, minute=0, second=0, microsecond=0)
    if start_dt is None:
        start_dt = end_dt - timedelta(days=1)

    daily_start = start_dt
    day2_start = end_dt - timedelta(days=2)
    day2_end = end_dt - timedelta(days=1)
    day3_start = end_dt - timedelta(days=3)
    day3_end = end_dt - timedelta(days=2)

    date_day2 = (end_dt - timedelta(days=1)).strftime("%Y-%m-%d")
    date_day3 = (end_dt - timedelta(days=2)).strftime("%Y-%m-%d")

    weekly_start = end_dt - timedelta(days=7)
    monthly_start = end_dt - timedelta(days=30)
    total_start = datetime(2000, 1, 1)

    daily_items = []
    day2_items = []
    day3_items = []
    weekly_items = []
    monthly_items = []
    total_items = []

    def _calc_total(items):
        t = {
            "매수건수": 0, "매수수량": 0.0, "매수금액": 0.0,
            "매도건수": 0, "매도수량": 0.0, "매도금액": 0.0,
            "그리드건수": 0, "그리드수량": 0.0, "그리드금액": 0.0,
            "그리드수익": 0.0, "그리드수수료": 0.0,
            "공백손절건수": 0, "공백손절수량": 0.0, "공백손절금액": 0.0,
            "공백실현손실": 0.0, "공백수수료": 0.0,
            "수수료": 0.0, "실현수익": 0.0,
            "거래일수": 0,
        }
        all_trade_dates = set()
        for _, st, _ in items:
            t["매수건수"] += st["매수건수"]; t["매수수량"] += st["매수수량"]; t["매수금액"] += st["매수금액"]
            t["매도건수"] += st["매도건수"]; t["매도수량"] += st["매도수량"]; t["매도금액"] += st["매도금액"]
            t["그리드건수"] += st.get("그리드건수", 0); t["그리드수량"] += st.get("그리드수량", 0.0); t["그리드금액"] += st.get("그리드금액", 0.0)
            t["그리드수익"] += st.get("그리드수익", 0.0); t["그리드수수료"] += st.get("그리드수수료", 0.0)
            t["공백손절건수"] += st.get("공백손절건수", 0); t["공백손절수량"] += st.get("공백손절수량", 0.0); t["공백손절금액"] += st.get("공백손절금액", 0.0)
            t["공백실현손실"] += st.get("공백실현손실", 0.0); t["공백수수료"] += st.get("공백수수료", 0.0)
            t["수수료"] += st["수수료"];   t["실현수익"] += st["실현수익"]
            dates = st.get("거래일목록")
            if isinstance(dates, set):
                all_trade_dates.update(dates)
            elif isinstance(dates, (list, tuple)):
                all_trade_dates.update(dates)
        t["거래일수"] = len(all_trade_dates)
        return t

    total_cum_daily = 0.0
    ref_end = end_dt.replace(tzinfo=None) if end_dt else datetime.now()
    overall_min_date = None
    item_total_days = {}

    for ticker in 종목코드목록:
        cfg = trade_config.get(ticker) or {}
        fp = cfg.get("file_name")
        if not fp or not os.path.exists(fp):
            continue

        try:
            df = pd.read_excel(fp, engine="openpyxl", sheet_name=0)
        except Exception as e:
            print(f"   ⚠️ {ticker} 엑셀 로딩 실패: {e}", flush=True)
            continue

        if df.empty:
            continue

        st_daily = _정산_데이터_집계(df, daily_start, end_dt)
        st_day2 = _정산_데이터_집계(df, day2_start, day2_end)
        st_day3 = _정산_데이터_집계(df, day3_start, day3_end)
        st_weekly = _정산_데이터_집계(df, weekly_start, end_dt)
        st_monthly = _정산_데이터_집계(df, monthly_start, end_dt)
        st_total = _정산_데이터_집계(df, total_start, end_dt)

        cum = _정산_시트갱신(fp, 정산일, st_daily)
        total_cum_daily += cum

        # 종목 첫 거래일 계산 (전체 정산 일일 평균용)
        b_dates = pd.to_datetime(df.get("매수날짜"), errors="coerce").dropna()
        s_dates = pd.to_datetime(df.get("매도날짜"), errors="coerce").dropna()
        all_dates = pd.concat([b_dates, s_dates])
        if not all_dates.empty:
            t_min = all_dates.min()
            if hasattr(t_min, "tz_localize") and t_min.tzinfo is not None:
                t_min = t_min.tz_localize(None)
            t_days = max(1, (ref_end.date() - t_min.date()).days)
            if overall_min_date is None or t_min < overall_min_date:
                overall_min_date = t_min
        else:
            t_days = 1

        name = ticker.split("/")[0]
        item_total_days[name] = t_days

        daily_items.append((name, st_daily, cum))
        day2_items.append((name, st_day2, None))
        day3_items.append((name, st_day3, None))
        weekly_items.append((name, st_weekly, None))
        monthly_items.append((name, st_monthly, None))
        total_items.append((name, st_total, None))

    tot_daily = _calc_total(daily_items)
    tot_day2 = _calc_total(day2_items)
    tot_day3 = _calc_total(day3_items)
    tot_weekly = _calc_total(weekly_items)
    tot_monthly = _calc_total(monthly_items)
    tot_total = _calc_total(total_items)

    tot_total_days = max(1, (ref_end.date() - overall_min_date.date()).days) if overall_min_date else 1

    main_symbol = 종목코드목록[0].split("/")[0] if len(종목코드목록) == 1 else ""
    header_title = f"🏁 *[비트겟 그리드봇 종합 정산 ({main_symbol})]*" if main_symbol else "🏁 *[비트겟 그리드봇 종합 정산]*"
    lines = [
        header_title,
        f"📅 {정산일} (KST {SETTLEMENT_HOUR_KST}시 기준)",
    ]

    # 1️⃣ [일일 정산] (1일전/2일전/3일전 통합)
    date_1d = 정산일
    lines += _포맷_일일_정산_섹션(
        "1️⃣ [일일 정산]",
        daily_items, day2_items, day3_items,
        tot_daily, tot_day2, tot_day3,
        date_1d, date_day2, date_day3,
    )

    # 2️⃣ [주간 정산] (최근 7일 / 거래 N일)
    주간_거래일수 = tot_weekly.get("거래일수", 0) if len(weekly_items) > 1 else (weekly_items[0][1].get("거래일수", 0) if weekly_items else 0)
    lines += _포맷_정산_섹션(
        f"2️⃣ [주간 정산] (최근 7일 / 거래 {주간_거래일수}일)",
        weekly_items, tot_weekly,
        profit_label="주간 순실현손익",
        days=주간_거래일수,
    )

    # 3️⃣ [월간 정산] (최근 30일 / 거래 N일)
    월간_거래일수 = tot_monthly.get("거래일수", 0) if len(monthly_items) > 1 else (monthly_items[0][1].get("거래일수", 0) if monthly_items else 0)
    lines += _포맷_정산_섹션(
        f"3️⃣ [월간 정산] (최근 30일 / 거래 {월간_거래일수}일)",
        monthly_items, tot_monthly,
        profit_label="월간 순실현손익",
        days=월간_거래일수,
    )

    # 4️⃣ [전체 정산] (전체 기간: N일)
    lines += _포맷_정산_섹션(
        f"4️⃣ [전체 정산] (전체 기간: {tot_total_days}일)",
        total_items, tot_total,
        profit_label="누적 순실현손익",
        days=tot_total_days,
        item_days_map=item_total_days
    )

    # 5️⃣ [현재 포지션 및 미실현 손익]
    try:
        lines += _포맷_포지션_현황_섹션()
    except Exception as e:
        print(f"[포지션 현황 섹션 생성 실패] {e}", flush=True)

    # 💼 [종합 손익 대시보드 요약] (순수 그리드 익절 vs 공백 손절 vs 현재 미실현 손익)
    lines.append("━━━━━━━━━━━━━━━")
    lines.append("💼 [종합 손익 대시보드 요약]")
    tot_all_g_profit = tot_total.get("그리드수익", 0.0)
    tot_all_c_loss = tot_total.get("공백실현손실", 0.0)
    lines.append(f"  • 🟢 누적 그리드 익절: +{tot_all_g_profit:,.4f} USDT")
    if tot_all_c_loss < 0:
        lines.append(f"  • 🔴 누적 공백 손절:   {tot_all_c_loss:,.4f} USDT")
    else:
        lines.append(f"  • ⚪ 누적 공백 손절:   0.0000 USDT")

    # 미실현 손익 합산
    total_unrealized = 0.0
    has_pos_data = False
    for ticker in 종목코드목록:
        pos = 현재_롱포지션_현황_조회(ticker)
        if pos and pos.get("contracts", 0.0) > 0:
            total_unrealized += float(pos.get("unrealizedPnl", 0.0) or 0.0)
            has_pos_data = True
    if has_pos_data:
        unr_p = "+" if total_unrealized >= 0 else ""
        lines.append(f"  • 🧊 현재 미실현 손익: {unr_p}{total_unrealized:,.4f} USDT")
    else:
        lines.append("  • 🧊 현재 미실현 손익: 0.0000 USDT (보유 포지션 없음)")

    # 현재 보유금 현황
    try:
        총진입, 잔여, 총보유 = 전체_보유금_현황_조회()
        if 잔여 is not None and 총보유 is not None:
            lines.append(f"  • 💼 계좌 잔고 현황: [포지션 {총진입:,.2f} + 잔여 {잔여:,.2f} = 총 {총보유:,.2f} USDT]")
        else:
            lines.append(f"  • 💼 계좌 잔고 현황: [포지션 {총진입:,.2f} + 잔여 조회실패 = 총 N/A USDT]")
    except Exception as e:
        print(f"[보유금 현황 조회 실패] {e}", flush=True)

    # 🔎 봇 정상 가동 중 (메시지의 가장 마지막!)
    lines.append(f"🔎 봇 정상 가동 중 ({지금().strftime('%Y-%m-%d %H:%M:%S')})")


    msg = "\n".join(lines)
    print(msg, flush=True)
    try:
        _텔레그램_안전발송(msg)
    except Exception as e:
        print(f"[정산 알림 발송 실패] {e}", flush=True)


def 일일정산_체크및실행():
    """KST SETTLEMENT_HOUR_KST 를 지났고 오늘 아직 정산 안 했으면 실행."""
    global _last_settlement_date
    now = _kst_now()
    # 정산 대상일 = 직전 09:00 이 속한 날. 09시 이전이면 아직 실행하지 않는다.
    anchor = now.replace(hour=SETTLEMENT_HOUR_KST, minute=0, second=0, microsecond=0)
    if now < anchor:
        return
    정산일 = anchor.strftime("%Y-%m-%d")

    if _last_settlement_date == 정산일:
        return

    # ★ 워치독이 봇을 재기동하면 메모리 변수가 초기화되므로, 같은 날 정산이
    #   반복 발송될 수 있다. 상태 파일로 재기동 후에도 중복을 막는다.
    try:
        if os.path.exists(SETTLEMENT_STATE_FILE):
            with open(SETTLEMENT_STATE_FILE) as f:
                if f.read().strip() == 정산일:
                    _last_settlement_date = 정산일
                    return
    except Exception as e:
        print(f"[정산 상태파일 읽기 실패] {e}", flush=True)

    _last_settlement_date = 정산일
    일일정산_실행(정산일, anchor - timedelta(days=1), anchor)

    try:
        with open(SETTLEMENT_STATE_FILE, "w") as f:
            f.write(정산일)
    except Exception as e:
        print(f"[정산 상태파일 기록 실패] {e}", flush=True)


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
    if TRAILING_ENABLE:
        print(f"   - 트레일링 매도: 1차보존 +{TRAIL_STAGE1_PCT}% / 2차활성 +{TRAIL_STAGE2_PCT}% (고점 -{TRAIL_DISTANCE_PCT}%)")
    else:
        print(f"   - 트레일링 매도: 비활성 (고정 매도간격)")
    if BUY_COOLDOWN_ENABLE:
        print(
            f"   - 매수 쿨링: {BUY_COOLDOWN_WINDOW_SEC // 60}분 내 "
            f"고유차수 {BUY_COOLDOWN_BUY_COUNT}개 → {BUY_COOLDOWN_PAUSE_SEC // 60}분 금지 + 하이브리드 rebase"
        )
    else:
        print(f"   - 매수 쿨링: 비활성")
    if TOTAL_PROFIT_TARGET_ENABLE:
        print(f"   - 총 수익 목표: {TOTAL_PROFIT_TARGET_USDT} USDT 도달 시 전체 청산 (소스={PROFIT_SOURCE}, 후속={TOTAL_PROFIT_AFTER_LIQUIDATION})")
    else:
        print(f"   - 총 수익 목표: 비활성화 (TOTAL_PROFIT_TARGET_ENABLE=False)")
    if TREND_RIDING_ENABLE:
        _sell_mode_label = {
            "STAGE1_ONLY": "0.7% 보존선 하회 시만 매도",
            "ALL"        : "1단계+2단계 모두 허용",
            "OFF"        : "100% 홀딩 (모든 매도 괈지)",
        }.get(TREND_RIDING_SELL_MODE, TREND_RIDING_SELL_MODE)
        print(f"   - 3-State 엔진: 활성화 (일봉+4시간봉 5MA -{TREND_EXIT_BUFFER_PCT}% MTF 지지선 이탈 필터)")
        print(f"   - 라이딩 중 매도: [{TREND_RIDING_SELL_MODE}] {_sell_mode_label}")
    else:
        print(f"   - 3-State 엔진: 비활성화 (기존 그리드)")
    print(f"   - 종목: {종목코드목록}", flush=True)

    # 종목별 엑셀 파일 초기화 (시작 알림의 기준가 조회보다 먼저)
    for Ticker in 종목코드목록:
        try:
            initialize_excel_file(trade_config[Ticker]["file_name"])
        except Exception as e:
            print(f"[파일 초기화 실패] {Ticker} / {e}")
        try:
            총수익_baseline_부트스트랩(Ticker)
        except Exception as e:
            print(f"[baseline 부트스트랩 실패] {Ticker} / {e}")

    # 시작 알림
    try:
        _sell_mode_desc = {
            "STAGE1_ONLY": f"0.7% 보존선 하회 시만 매도 (우상향 홀딩)",
            "ALL"        : f"1/2단계 트레일링 전체 허용",
            "OFF"        : f"100% 홀딩 (매도 전면 금지)",
        }.get(TREND_RIDING_SELL_MODE, TREND_RIDING_SELL_MODE)
        trend_line = (
            f"🚀 3-State 엔진: `일봉+4시간봉 5MA MTF 추세 라이딩` (이탈: -{TREND_EXIT_BUFFER_PCT}% 동시하회)\n"
            f"🔄 라이딩 중 매도: `[{TREND_RIDING_SELL_MODE}] {_sell_mode_desc}`\n"
            if TREND_RIDING_ENABLE else "🚀 3-State 엔진: `비활성`\n"
        )
        trail_line = (
            f"📍 트레일링: 1차`+{TRAIL_STAGE1_PCT}%` 보존 / 2차`+{TRAIL_STAGE2_PCT}%` 활성(고점`-{TRAIL_DISTANCE_PCT}%`)\n"
            if TRAILING_ENABLE else "📍 트레일링: `비활성`\n"
        )
        cool_line = (
            f"🧊 쿨링: `{BUY_COOLDOWN_WINDOW_SEC // 60}분 고유차수 {BUY_COOLDOWN_BUY_COUNT}개` "
            f"→ `{BUY_COOLDOWN_PAUSE_SEC // 60}분` + 하이브리드 rebase (>A 즉시/≤A 저점반등)\n"
            if BUY_COOLDOWN_ENABLE else "🧊 쿨링: `비활성`\n"
        )
        gap_line = (
            f"🌪️ 공백 복원: `활성 (v2.4 비상 자가치유 롤다운)` (최대 {MAX_GAP_CUT_LEVELS}개 손절 캡 + 바닥 {MAX_GAP_CUT_LEVELS}개 그물망)\n"
            if GAP_RECONSTRUCTION_ENABLE else "🌪️ 공백 복원: `비활성`\n"
        )
        price_lines = ""
        profit_lines = ""
        balance_lines = ""
        for t in 종목코드목록:
            px = 현재가조회(t)
            if px <= 0:
                try:
                    _df = 매매기록불러오기(trade_config[t]["file_name"])
                    if _df is not None and not _df.empty:
                        _cleared = _df[_df["매수가"].notna()]
                        if not _cleared.empty:
                            px = float(_cleared.iloc[-1]["매수가"])
                except Exception:
                    pass

            price_lines += 시작시_매수기준_요약(t, 현재가=px)

            if TOTAL_PROFIT_TARGET_ENABLE:
                try:
                    profit_lines += 총수익_텔레그램라인(t, px if px > 0 else 0.0)
                except Exception as e:
                    print(f"[시작 알림 총수익 조회 실패] {t} / {e}")

            try:
                balance_lines += 보유금_텔레그램라인(t)
            except Exception as e:
                print(f"[시작 알림 보유금 조회 실패] {t} / {e}")

        telegram_alert.SendMessage(
            f"🚀 *[비트겟 그리드봇 시작]*\n"
            f"━━━━━━━━━━━━━━━\n"
            f"⏱ 폴링 주기: `{POLL_INTERVAL}초`\n"
            f"📊 종목: `{', '.join(종목코드목록)}`\n"
            f"{trend_line}"
            f"{trail_line}"
            f"{cool_line}"
            f"{gap_line}"
            f"🎯 총수익 목표: `{TOTAL_PROFIT_TARGET_USDT} USDT` / 후속 `{TOTAL_PROFIT_AFTER_LIQUIDATION}`"
            f"{' (활성)' if TOTAL_PROFIT_TARGET_ENABLE else ' (비활성)'}\n"
            f"{price_lines}"
            f"{profit_lines}"
            f"{balance_lines}"
            f"🕒 시간: `{지금().strftime('%Y-%m-%d %H:%M:%S')}`\n"
            f"━━━━━━━━━━━━━━━"
        )
        time.sleep(1.0)
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

                    # --- 3-State Switching Engine (5일선 추세 라이딩 판정) ---
                    라이딩_중 = False
                    방금_꺾임 = False
                    if TREND_RIDING_ENABLE:
                        라이딩_중, 방금_꺾임 = 추세라이딩_상태체크(Ticker, 현재가)

                    # 총 수익 모니터링 로그 (선택)
                    if (TOTAL_PROFIT_TARGET_ENABLE
                        and TOTAL_PROFIT_LOG_INTERVAL > 0
                        and loop_count % TOTAL_PROFIT_LOG_INTERVAL == 0):
                        try:
                            미실현_log, 누적실현_log, 총합_log, baseline_log, 사이클_log = 총수익_계산(Ticker, 현재가)
                            회차_log = 총수익_사이클회차_조회(Ticker)
                            락_표시 = " 🔒락" if 총수익청산락_여부(Ticker) else ""
                            모드_표시 = " [🚀라이딩]" if 라이딩_중 else ""
                            print(
                                f"💰 [총수익] {Ticker.split('/')[0]} | "
                                f"{회차_log}회차{모드_표시} | "
                                f"사이클={사이클_log:+.4f} "
                                f"(미실현={미실현_log:+.4f} + 누적={누적실현_log:+.4f} "
                                f"- baseline={baseline_log:+.4f}) "
                                f"총합={총합_log:+.4f} "
                                f"(목표 {TOTAL_PROFIT_TARGET_USDT}){락_표시}",
                                flush=True
                            )
                        except Exception as e:
                            print(f"[총수익 로그 실패] {e}")

                    if 라이딩_중:
                        # ── 상태 2: [추세 라이딩] ──
                        # TREND_RIDING_SELL_MODE에 따라 일부 매도 허용 여부 제어
                        매도발생 = False
                        if TREND_RIDING_SELL_MODE in ("STAGE1_ONLY", "ALL"):
                            # 이익 보존선 하회 시 방어 매도 허용
                            매도발생 = 매도실행(Ticker, 현재가, riding_mode=TREND_RIDING_SELL_MODE)
                            더미_가상매도_감시_및_실행(Ticker, 현재가)
                        # 매도 미발생 시 매수 유지 (하단 방어)
                        if not 매도발생:
                            매수실행(Ticker, 현재가)
                    else:
                        # ── 상태 3 및 상태 1: [추세 꺾임 / 그리드 수비 모드] ──
                        # 1) 목표 순익 달성 검사 및 전량 청산 판단 (꺾임 직후 또는 정기 주기)
                        총수익_청산_발생 = False
                        if TOTAL_PROFIT_TARGET_ENABLE:
                            if 방금_꺾임 or (loop_count % TOTAL_PROFIT_CHECK_INTERVAL == 0):
                                총수익_청산_발생 = 총수익_체크_및_청산(Ticker, 현재가)
                                if 총수익_청산_발생:
                                    continue  # Flash Close 청산 완료 -> 이번 폴링 종료 (다음 사이클 준비)

                        # 방금 꺾였으나 목표 미달로 청산되지 않은 경우 (Case B 안내)
                        if 방금_꺾임 and not 총수익_청산_발생:
                            try:
                                _, _, _, _, 사이클_val = 총수익_계산(Ticker, 현재가)
                                title_sym = Ticker.split('/')[0]
                                resume_msg = (
                                    f"🛡️ *[{title_sym}_그리드 매도 재개]*\n"
                                    f"━━━━━━━━━━━━━━━\n"
                                    f"📉 *추세 꺾임*: 일봉/4시간봉 5MA 지지선(-{TREND_EXIT_BUFFER_PCT}%) 동시 이탈 확인\n"
                                    f"💰 *사이클 순익*: `{사이클_val:+.4f} USDT` (목표 {TOTAL_PROFIT_TARGET_USDT} USDT 미달)\n"
                                    f"🔄 *조치*: 멈췄던 그리드 매도를 재개하여 고점 순차 익절 수비 모드로 복귀합니다.\n"
                                    f"🕒 *시간*: `{지금().strftime('%Y-%m-%d %H:%M:%S')}`\n"
                                    f"━━━━━━━━━━━━━━━"
                                )
                                telegram_alert.SendMessage(resume_msg)
                            except Exception as e:
                                print(f"[매도 재개 알림 오류] {e}")

                        # 2) 그리드 매도 실행 (고점 순차 익절) -> 매도 미발생 시 매수 실행
                        매도발생 = 매도실행(Ticker, 현재가)
                        더미_가상매도_감시_및_실행(Ticker, 현재가)
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
                    f"📊 [{지금().strftime('%H:%M:%S')}] "
                    f"폴링#{loop_count} | {가격_표시}",
                    flush=True
                )

            # ── 생존 신호: 텔레그램 발송 없이 콘솔에만 기록 ──────────────
            #   생존 확인은 크론 워치독(1분마다 재기동)이 담당한다.
            #   봇이 죽으면 워치독이 되살리고, 그때 '시작 알림'이 한 번 나가므로
            #   그 알림 자체가 사망 신호 역할을 한다. (10분마다 오던 하트비트 제거)
            now = time.time()
            if now - last_heartbeat >= HEARTBEAT_INTERVAL:
                print(f"💓 [생존] 폴링 {loop_count}회 | "
                      f"{지금().strftime('%Y-%m-%d %H:%M:%S')}", flush=True)
                last_heartbeat = now

            # ── 일일 정산 (하루 1회) ────────────────────────────────
            try:
                일일정산_체크및실행()
            except Exception as e:
                print(f"[일일정산 체크 실패] {e}", flush=True)

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
            f"🕒 `{지금().strftime('%Y-%m-%d %H:%M:%S')}`\n"
            f"━━━━━━━━━━━━━━━"
        )
        time.sleep(2.0)  # 텔레그램 비동기 네트워크 패킷 전송 완료 대기 (프로세스 즉시 종료로 인한 유실 방지)
    except Exception as e:
        print(f"[종료 알림 실패] {e}")


if __name__ == "__main__":
    main_loop()
