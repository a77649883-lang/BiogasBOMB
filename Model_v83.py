시발
# =======================================================================
# Model v79 — 강원도 유기성폐자원 바이오가스 슈퍼스트럭처 최적화
#
#   v78 대비 변경점 — 원료별 COD 상수 정정 (§2 COD_*_KG_M3 주석 참조)
#
#     v78 까지 COD_ANI/FW/SL = 40.79 / 6.66 / 0.75 를 썼다. 이는
#     Paper_feed721.xlsx Sheet1 'COD(kg)' 열인데, 그 열은 7:2:1 혼합비가
#     이미 곱해진 '혼합물 1 m3 당 기여분' 이라 원료별 농도가 아니다.
#     지역별 원료유량에 곱하면 비율이 이중 적용되고, 지역 조성에 따라
#     1.4 ~ 81 배까지 제각각 어긋난다(태백시 3.6 배).
#
#     혼합비를 뺀 물성값으로 다시 계산한다:
#         COD = TS(0.05) x 1000 x (VS/TS) x (COD/VS)
#         -> 축분 58.29 / 음식물 75.76 / 슬러지 61.08 kgCOD/m3
#
#     이 상수는 열역학 검산과 보고용 열에만 쓰이고 NN 입력에는 안 들어가므로
#     생산량·비용·LCOE 결과는 바뀌지 않는다. 태백시 기준 메탄 수율이
#     이론최대의 2.07 배(위반) -> 0.57 배(정상)가 되어 경고가 해소된다.
#
# =======================================================================
# Model v78 — 강원도 유기성폐자원 바이오가스 슈퍼스트럭처 최적화
#
#   v77 대비 변경점 — 둘 다 '조용히 틀린 결과'를 막는 방어장치다.
#
#     1. [grid_hint 미적용 행 무효화]
#        probe 가 실패하면 grid_hint 없이 스윕하는데, 이중선형(총가스량 x
#        CH4분율) 근사격자가 넓어져 근사오차가 route 선택을 뒤집는다.
#        태백시 실측(동일 epsilon, 동일 모델):
#            grid_hint 있음 -> purification, 11,279 kWh/d, TAC 3,115,854
#            grid_hint 없음 -> generator,     4,523 kWh/d, TAC 2,985,074
#        공급량이 2.5 배 차이인데 solver 는 둘 다 'optimal' 을 보고한다.
#        -> grid_hint_applied 플래그를 세우고 pareto_valid=0 으로 떨어뜨린다.
#           값은 Raw_results 에 남지만 Pareto_results 와 LCOE 최소점 선정에서
#           제외된다. 검증 블록에 해당 row_index 를 찍는다.
#
#     2. [fallback 경로의 demand 누락 수정]
#        probe 실패 시 fallback 균등 스윕이 solve_case 를 demand 없이
#        호출하고 있었다. 수요 상한이 안 걸리면 목적함수가 '총생산량 최대화'
#        로 조용히 퇴화한다(v73 이후 핵심 장치가 무력화). demand 를 넘긴다.
#
# =======================================================================
# Model v77 — 강원도 유기성폐자원 바이오가스 슈퍼스트럭처 최적화
#
#   v76 대비 변경점
#     1. [폐수처리 약품비] 메탄올·소석회(질소부하 기반) + Alum 을 변동 OPEX 에
#        추가. v75/v76 은 "질소농도·단가 자료 없음"으로 미반영이었으나
#        TN=2833 mg/L(소화액 탈리액 실측)와 capex_selector 단가로 확보됨.
#        태백 발생량 92 t/d 기준 메탄올+소석회 460,303 $/y 로, 이 항 하나가
#        변동 OPEX 의 85% 다.
#     2. 메탄올·소석회는 A2O/Bardenpho 에만 부과(Anammox 는 원리적으로 불필요).
#        v76 까지 약품비가 0 이라 Anammox 가 구조적으로 불리했는데, 이 항이
#        들어가면서 폐수처리 공법 선택이 비로소 실제 경쟁이 된다.
#     3. TN 은 지역별 실측이 없어 상수(WWTP_TN_MG_L)로 둔다.
#
#   [주의] TN 은 CAPEX 에 전혀 영향을 주지 않는다. 폐수처리 반응조 부피는
#     유량(x1.3)만으로 정해지고 질소부하는 반영돼 있지 않다. 즉 TN 을 바꾸면
#     약품비만 움직인다. 고농도 질소가 반응조 용량을 키우는 효과를 보려면
#     별도 설계식이 필요하다(현재 자료 없음).
#
# =======================================================================
# Model v82 — 강원도 유기성폐자원 바이오가스 슈퍼스트럭처 최적화
#
#   v81 대비 변경점 — 폐수처리 OPEX 를 New4.xlsx Data_OPEX 원자료와 대조해 수정.
#   v81 의 EL_WWTP 는 원자료에 없는 값(0.39/0.12/0.33)을 쓰고 있었고 A2O 에서
#   호기조(0.412)가 누락돼 있었다. capex_selector.py 와 대조해 확인했다.
#
#     [C1] EL_WWTP 전기 원단위 정정  (Data_OPEX 열 98~109)
#            A2O       v81 0.680 -> v82 0.650 (+공통 0.150) = 0.800
#            Bardenpho v81 1.482 -> v82 0.650 (+공통 0.150) = 0.800
#          Bardenpho 2단 무산소·호기조는 원자료에 전기 원단위가 비어 있다.
#     [C2] DAF(0.05) + UV(0.1) 를 공법 무관 공통항으로 분리.
#          v81 은 Anammox 에서 이 둘을 통째로 건너뛰어 0 으로 계산했다.
#     [C3] 메탄올을 무산소조 단수에 비례시킴 (Data_OPEX 열 99·103·105).
#            A2O       무산소 1단 -> 메탄올 x1, 소석회 x1
#            Bardenpho 무산소 2단 -> 메탄올 x2, 소석회 x1  (2단은 소석회 없음)
#          v81·capex_selector 모두 Bardenpho 메탄올을 1배만 넣고 있었다.
#     [C4] Anammox 황산비(8,760,000 KRW/y) 를 토글로 분리, 기본 비활성.
#          Data_OPEX 의 Anammox 열은 '황' 행이 비어 있어 근거를 찾지 못했다.
#     [C5] 가동률 F_UTIL 을 전기 사용량 항에도 적용 (기본요금은 제외).
#          v81 은 약품비에만 F_UTIL 을 곱하고 전기에는 곱하지 않았다.
#
#   ※ 위 다섯 건 모두 F_UTIL=1.0 · Anammox 선택 조건에서는 [C2][C3][C4] 만
#     결과를 바꾼다. A2O/Bardenpho 가 선택되는 경우에만 [C1][C3] 이 드러난다.
#
#   연구 목표: 수익 극대화가 아니라 "가성비 좋게 도시가스·전력을 공급"하는 것.
#
#   다목적 정식화 (epsilon-constraint)
#       목적   max  유효공급에너지 E [kWh/d]  = min(생산, 수요) 합
#       제약   TAC <= epsilon  [USD/y]        = 연환산 CAPEX + 연간 OPEX
#     -> epsilon 을 훑어 (TAC, E) Pareto front 를 얻고,
#        각 점의 LCOE = TAC / (E x 365) 중 최소가 '가성비 최적점'이다.
#        (LCOE 최소 설계는 반드시 이 front 위에 있다 — S14 주석의 증명 참조)
#
#   v75 대비 변경점
#     1. [비용기준] C_agg = 구입기기비 + 표준사업비 x 77.1% (부대비용).
#        v75 는 부대비용을 통째로 빼고 있었다. 태백시 기준 부대비용이
#        17,792,271 USD 로 구입기기비 4,298,759 USD 의 4.1 배였고,
#        그 결과 LCOE 가 0.174 -> 0.752 $/kWh 로 4.3 배 바뀐다.
#        더 중요한 것은 LCOE 최소 설계가 wet_scrubber -> psa 로 뒤집힌다는 점이다.
#        (부대비용은 상수라 TAC 순위는 안 바뀌지만, LCOE 는 분수라 분모가 큰
#         설계가 유리해진다.)  INCLUDE_INDIRECT_COST=False 로 두면 v75 동작.
#     2. [산출물] 설계 전환 임계 TAC 자동 추출. 균등격자로 front 를 그린 뒤
#        이진 선택변수가 바뀐 구간만 이분탐색해 "TAC 가 얼마를 넘으면 어느
#        기술로 갈아타는가"를 Design_switch 시트로 낸다.
#     3. [경고] 메탄 수율이 물리 COD 이론상한을 넘으면 지역 요약에 경고를 찍는다.
#        (태백시 실측 2.07 배. COD 상수 출처 확인 전까지 값은 건드리지 않는다.)
#
#   v74 대비 변경점 (v75 에서)
#     1. epsilon 축을 CAPEX -> TAC 로. v74 는 변동이 작은 CAPEX(+10%)를 묶고
#        변동이 큰 OPEX(+56%)를 풀어놓아, 약품비가 비싼 정제기술이 공짜로 통과했다.
#     2. 목적함수를 판매수익 -> 유효공급에너지로. 사업수익이 아니라 공급량이 목표.
#     3. 전기요금을 단일단가 -> TOU(계절x부하) + 기본요금 + 8h/24h 가동 구분.
#        (capex_selector.py 의 Elec 모델 이관)
#
# ---------------------------------------------------------------- 목차
#   §1  실행 설정            INPUT_EXCEL / REGIONS_TO_RUN / N_EPS 등
#   §2  설비 원단위·비용 상수  기준용량·기준비용·scale factor 표
#   §3  공용 헬퍼            scale_cost, cost_capacity_breakpoints,
#                            add_cost_capacity_pw, add_bigm_select
#   §4  NN surrogate         scaler / bounds / OMLT ReLU 블록
#   §5  MODEL BUILD          공정 순서 스테이지 S1~S14 + build_model
#   §6  SOLVE                solve_model / retry / maxobj-then-mincost
#   §7  PROBE · 진단 · 요약    probe_cost_range, get_model_summary 등
#   §8  EPSILON SWEEP        solve_case / adaptive / run_region_epsilon_sweep
#   §9  RUN ALL REGIONS      지역 루프 · 로깅 · __main__
#
#   공정 순서(§5):
#     S1  기초·단위환산      S8  압축·정제 비용
#     S2  반입               S9  정제 성능
#     S3  전처리·중간저장     S10 발전(CHP)
#     S4  혐기소화(AD)       S11 소화액 탈수·물질수지
#     S5  가스계통 설비선택   S12 폐수처리
#     S6  가스 물질수지      S13 CAPEX 집계
#     S7  탈황·제습·저장     S14 OPEX·수익·목적함수
# =======================================================================

import os
import sys
import json
import time
import math
import contextlib
import traceback
from pathlib import Path
from functools import lru_cache
from types import SimpleNamespace

import numpy as np
import pandas as pd
import pyomo.environ as pyo

# keras/tensorflow 가 없는 환경에서도 .keras 파일을 직접 읽어 OMLT 망을 만들 수 있게
# import 를 선택적으로 처리한다(가중치가 같으므로 결과는 동일).
try:
    from tensorflow import keras
except Exception:
    try:
        import keras
    except Exception:
        keras = None

from omlt import OffsetScaling, OmltBlock
from omlt.neuralnet import ReluBigMFormulation

# =======================================================================
# §1  실행 설정
# =======================================================================

# omlt.io.keras 는 import 시점에 tensorflow 를 요구하므로 선택적으로 처리한다.
try:
    from omlt.io.keras import load_keras_sequential
except Exception:
    load_keras_sequential = None

# =========================================================
# Settings
# =========================================================
INPUT_EXCEL = "2020년_지역별_발생량_강원_수요포함.xlsx"

# A열(REGION_COL)의 지역별로 나눠 각각 실행하고, 지역이 끝날 때마다
# ModelOutPut_v68_{지역명}.xlsx 로 저장한다.
OUTPUT_PREFIX = "ModelOutPut_v82"   # 실제 파일: {OUTPUT_PREFIX}_{지역명}.xlsx
REGION_COL = "시군구"                # A열: 지역 분할 기준

# None 이면 A열의 모든 지역을 실행. 특정 지역만 돌리려면
# ["원주시", "강릉시"] 처럼 리스트로 지정한다.
REGIONS_TO_RUN = None

# stdout 출력 내용을 지역별 {지역명}.txt 로도 저장한다(콘솔에도 그대로 표시).
LOG_PER_REGION = True

NN_MODEL_PATH = "biogas_nn_relu.keras"

NN_OUTPUTS = [
    "final_methane_fraction_dry",
    "final_q_gas(m3/d)",
    "final_co2_fraction_dry",
]
SCALER_PATH = "biogas_nn_relu_scaler.json"
TRAINING_DATA_PATH = "ADMoutput_0512.xlsx"

# HRT 고정값 [day]
HRT_DAYS = 40

# 테스트용: 처음 몇 개 지역만 실행. 전체 실행은 None.
MAX_ROWS = None

# 지역별 트레이드오프 곡선 해상도 (= 지역당 최대 solve 횟수 = 계산비용 상한)
N_EPS = 40

# ---------------------------------------------------------
# 요약 시트(Key_results)에 실을 열
# ---------------------------------------------------------
# Raw_results 는 413열이라 눈으로 읽기 어렵다. 아래 목록만 뽑아 Key_results
# 시트를 만들고 첫 시트로 둔다(엑셀을 열면 이 시트가 먼저 보인다).
# Raw_results / Pareto_results 는 그대로 두므로 상세 추적은 언제든 가능하다.
#
# 열을 더/덜 보고 싶으면 이 목록만 고치면 된다. 여기 없는 이름을 적으면
# 실행 시 경고만 찍고 건너뛴다. 빈 리스트로 두면 Key_results 를 만들지 않는다.
KEY_COLUMNS = [
    # 식별 · 실행조건
    "시군", "읍면동", "리", "epsilon", "termination",
    # 원료 투입
    "W_raw_t_d", "F_raw_t_d", "S_raw_t_d", "Q_total",
    # 설계 선택 결과
    "selected_route", "AD_train_count", "V_reactor",
    "selected_purification_tech", "selected_generator",
    # 생산 성능
    "pred_q_gas", "pred_ch4_frac_dry", "up_ch4", "gen_power_kW",
    "city_gas_production_kWh_d", "electricity_production_kWh_d",
    # 경제성  (Cost 는 CAPEX_usd 와 같은 값이라 뺐다)
    # [v76] CAPEX 는 구입기기비 / 부대비용 / 합계(C_agg) 로 나눠 본다.
    "equipment_cost_usd", "indirect_cost_usd", "C_agg_usd",
    "annualized_capex_usd_y",
    "CAPEX_usd", "TAC_usd_y", "useful_energy_kWh_d", "objective_value",
    "subsidy_usd", "C_agg_net_usd", "annualized_capex_net_usd_y",
    "tipping_revenue_usd_y", "TAC_net_usd_y", "LCOE_net_usd_per_kWh",
    "levelized_energy_cost_usd_per_kWh", "levelized_energy_cost_krw_per_kWh",
    "LCOB_usd_per_m3_CH4", "LCOB_krw_per_m3_CH4",
    "LCOB_net_usd_per_m3_CH4", "LCOB_net_krw_per_m3_CH4",
    "LCOB_carbon_usd_per_m3_CH4", "LCOB_carbon_krw_per_m3_CH4",
    "LCOB_carbon_net_usd_per_m3_CH4", "LCOB_carbon_net_krw_per_m3_CH4",
    "CO2_reduction_tCO2_y", "carbon_credit_krw_y", "carbon_credit_usd_y",
    "LCOB_opex_annual_usd_y", "LCOB_elec_cost_usd_y",
    "wwtp_chem_methanol_lime_usd_y", "wwtp_chem_alum_usd_y", "wwtp_N_load_kgN_d",
    "selected_wwtp",
    "elec_kwh_d_24h", "elec_kwh_d_8h", "daily_revenue_usd_d",
    # 신뢰성 플래그 — 이 값들이 정상이어야 위 수치를 믿을 수 있다
    "is_min_lcoe_point",
    "elec_peak_kW", "elec_contract_kW",
    "nn_all_inputs_in_training_range", "surrogate_reliable",
    "AD_mass_balance_error_kg_d", "solids_balance_residual_kg_d",
    "grid_hint_applied", "pareto_valid",
]

# =========================================================
# 적응형 epsilon 스텝 설정 (v44.1)
# =========================================================
# 기존 방식은 지역 비용범위 [C_min, C_sat] 를 무조건 N_EPS 등분(np.linspace)
# 했기 때문에, 설계 조합이 이산적으로만 바뀌는 특성상 40개 solve 중 상당수가
# '같은 최적해'를 다시 푸는 데 낭비되었다(=옵션이 찔끔찔끔만 바뀌는 원인).
#
# ADAPTIVE_EPS=True 이면 균등 등분 대신,
#   '설계 조합(이진 선택변수 집합)이 실제로 바뀌는 임계 epsilon' 만
#   이분탐색(bisection)으로 찾아 solve 한다. 같은 solve 예산(N_EPS)으로
#   훨씬 많은 서로 다른 설계(Pareto 변곡점)를 잡아낸다.
ADAPTIVE_EPS = False   # 균등 CAPEX epsilon grid 사용

# 이분탐색 종료 허용오차(상대):
#   구간폭이 (지역 비용범위 span * 이 값) 이하로 좁아지면
#   그 사이의 임계 epsilon 을 충분히 특정한 것으로 보고 refine 을 멈춘다.
EPS_BREAKPOINT_TOL_FRAC = 1.0e-4

# 출력 단계에서 연속으로 동일한 조합이 나오면
#   임계 epsilon(그 조합이 처음 등장한 지점)만 남기고 제거한다.
DEDUPE_CONSECUTIVE = False  # 동일 설계조합도 전부 남긴다 (eps 격자 그대로 출력)
DROP_DOMINATED_POINTS = False  # 지배해도 전부 남긴다 (필요하면 후처리에서 필터)


# =========================================================
# v49 신규 설정  (v48 대비 변경점 6가지)
# =========================================================
# --- [Fix 1] NN 입력 COD -----------------------------------------------
#   NN 입력 COD 는 원료별 실측 COD(40.79/6.66/0.75)가 아니라 학습 당시 basis
#   (SURROGATE_COD_COEF = 75.76)를 쓴다. 이유는 선택이 아니라 제약이다:
#
#   [정정] surrogate 는 '농축 후 AD 유입' 기준으로 학습됐고(Paper_feed721.xlsx
#   Sheet2 의 TS 132 / VS 125 / COD 200 kg/m3), 75.76 은 그걸 원료(TS 5%)
#   기준으로 환산한 값이다. 즉 기준 자체는 일관되며 현재 해는 전부 학습범위 안이다.
#
#   남은 한계는 두 가지다.
#   (1) 원료 종류를 구분하지 않고 75.76 을 일괄 적용한다. Paper_feed721 Sheet1 의
#       원료별 물성(TS x VS/TS x COD/VS)으로는 가축분뇨 349.8 / 음식물 200.0 /
#       슬러지 45.2 kgCOD/m3 이므로, 조성이 크게 다른 지역일수록 오차가 커진다.
#   (2) NN 이 q_ad <= 10 구간에서 학습데이터를 -46% 로 과소재현한다
#       (q_ad > 20 에서는 +-2%). surrogate_reliable 열로 표시된다.

# --- [Fix 2] 이중선형(총가스량 x CH4분율) 근사 ------------------------------
GAS_GRID_N = 13                 # v48: 11 (대신 격자 '범위'를 probe로 대폭 축소)
QGAS_HINT_MARGIN = 1.35         # probe 최대가스량 대비 격자 상한 여유
CH4_FRAC_MARGIN = 0.10          # probe CH4분율 대비 격자 여유
# 목적함수가 NPV(수백만 USD 규모의 음수)라 '상대' gap 은 척도로 부적절하다.
# (rel 1e-4 로 두면 같은 문제가 6초 -> 300초 초과로 느려지면서도 해는 동일했다)
# 절대 gap 을 주 기준으로 쓰고 상대 gap 은 끈다.
# 목적함수가 up_ch4 [m3 CH4/d] 이므로 상대 gap 이 다시 의미를 갖는다.
# 절대 gap 은 소규모 지역(up_ch4 약 40)에서의 하한으로만 쓴다.
SOLVER_MIP_REL_GAP = 1.0e-4         # v48: 1.0e-3
SOLVER_MIP_ABS_GAP = 0.02           # m3 CH4/d
SOLVER_TIME_LIMIT = 600.0           # v48: 300.0

# --- [Fix 5] cost-capacity 지수 정정 (v48 원자료 오류로 판단) ---------------
SF_FEEDING_HOPPER = 0.6         # v48: 1.5  (SF>1 = 규모의 불경제)
SF_SCREENING_SHREDDER = 0.6     # v48: 1.5
SF_TWO_SHAFT_SHREDDER = 0.6     # v48: 0.01 (용량 무관 상수가 되어버림)

# --- [Fix 6] 목적함수: CAPEX vs 정제메탄량 이원목적 ------------------------
#   v48: 목적 = USD/day 수익, 제약 = USD 자본비  -> 단위 불일치.
#   v49~v51: NPV 목적. 단위는 맞았지만 OPEX 계수 7개를 가정해야 했고,
#            그 가정 하나가 "몇 개 지역이 흑자냐"를 결정해 버렸다
#            (OPEX 제외 3/11 흑자, 포함 0/11).
#   v52: 두 목적을 '더하지 않는다'.
#          목적   max up_ch4        [m3 CH4/d]  정제 후 판매 가능 메탄량
#          제약   Cost <= epsilon   [USD]       CAPEX
#        epsilon 을 훑어 CAPEX-메탄량 Pareto front 를 그린다.
#        할인율, 설비수명, 가동률, OPEX, 판매단가 — 어느 것도 가정하지 않는다.
#
#   주의: 목적이 up_ch4(정제 후 메탄)이므로 발전 route 는 up_ch4 = 0 이라
#         구조상 절대 선택되지 않는다. 발전 설비 자체는 모델에 남아 있으나
#         이 목적함수 하에서는 사실상 비활성이다.
#         정제/발전을 함께 비교하려면 목적을 총 바이오가스 에너지나
#         일일 판매수익으로 바꿔야 한다.
#
#   판매단가(CITY_GAS_PRICE / ELEC_PRICE)로 계산되는 수익 열은 참고용으로만
#   출력되며 최적화에는 전혀 쓰이지 않는다.
# =========================================================
# 지역 수요 대조 (v53)
# =========================================================
# 입력 엑셀에 아래 두 열이 있으면 읽어서 '대체 가능량'의 상한으로 쓴다.
# 열이 없거나 비어 있으면 해당 에너지원의 수요를 무제한으로 간주한다.
#
# 단위는 모두 kWh/d (발열량 기준). 도시가스는 부피가 아니라 에너지로 받는다.
#   - 바이오메탄 생산 에너지 = up_ch4 [m3/d] x 9.94 [kWh/m3]
#   - 발전 전기 생산량      = gen_power_kW x 24
#
# 주의 1) 시군구 단위 수요밖에 없으면 읍면동 안분을 '엑셀에서' 해서 넣을 것.
#         코드 안에 숨기지 말고 데이터에 드러나 있어야 방어가 된다.
# 주의 2) 도시가스는 배관망이 없는 지역이면 인구 안분이 아니라 0 이어야 한다.
#         면지역에 배관망이 없다는 사실이 정제/발전 선택을 가르는 핵심이다.
# 주의 3) 여기서는 '수요를 대체한 만큼'만 가치로 계산한다. 잉여 전력을
#         계통에 역송해 판매하는 경우는 반영돼 있지 않다.
DEMAND_GAS_COL = "도시가스수요_kWh_d"
DEMAND_ELEC_COL = "전력수요_kWh_d"


# =========================================================
# LCOB (Levelized Cost Of Biomethane)  — New4.xlsx 'LCOB' 시트 정의
# =========================================================
#   LCOB = {CAPEX_tot * CRF + OPEX_tot} / (Q_CH4 * 365 * f_util)   [$/m3 CH4]
#   CAPEX_tot = f_inv * C_agg
#   CRF       = i / {1 - (1+i)^-n}
#   OPEX_tot  = f_MLC * C_agg + OPEX_fixed + OPEX_var + C_elec
#
# 결정사항:
#   C_agg   = 모델이 계산한 장비구입비(model.Cost).
#             New4.xlsx 'LCOB' 시트의 C_agg = 24,317,639 $ 는 Base 시트의
#             '표준 사업비'(3.528e10 ￦)와 일치하는 값이라 라벨과 달리
#             장비구입비가 아니다. 같은 파일 Sheet3 의 장비 합계는
#             4,456,633 ~ 5,350,179 $ 이고 우리 모델값 4,628,175 $ 와 일치한다.
#   Q_CH4   = 정제 후 생산 메탄 up_ch4 [m3/d]
#   발전전기 = 부산물 크레딧으로 분자에서 차감 (LCOE 관행)
#
# LCOB 는 분수형이라 MILP 에 직접 못 넣는다 -> Dinkelbach 반복법으로 푼다.
LCOB_I = 0.09                   # 가중평균자본비용
LCOB_N = 30.0                   # 설비 수명 [year]
LCOB_F_INV = 1.0                # 설치/투자계수
LCOB_F_MLC = 0.03               # 유지보수/인건/약품 [1/year]
LCOB_F_UTIL = 1.0               # 가동률
LCOB_OPEX_FIXED_USD_Y = 0.0
LCOB_CRF = LCOB_I / (1.0 - (1.0 + LCOB_I) ** (-LCOB_N))

EXCHANGE_RATE_KRW_PER_USD = 1434.0   # New4.xlsx Data_CAPEX B1
# 전기단가: New4.xlsx 'Elec' 시트 (산업용전력(을) 고압A 선택3) 실효단가
# [v75] 아래 TOU 모델로 대체됨. 비교용으로만 남긴다.
OPEX_ELEC_PRICE_USD_PER_KWH = 0.09858

# =========================================================
# [v75] 전기요금 TOU 모델  (capex_selector.py 'Elec' 시트에서 이관)
# =========================================================
# v74 는 전 설비에 단일 단가 0.09858 $/kWh 를 곱했다. 실제 산업용 요금은
#   (1) 계절(여름/봄가을/겨울) x 부하(경/중간/최대) 별 전력량요금
#   (2) 계약전력 기준 기본요금
#   (3) 기후환경요금 + 연료비조정요금
# 으로 구성되고, 무엇보다 전처리 설비는 주간 8시간만 돌아서 같은 일일 kWh 라도
# 평균전력(kW)이 24h 설비의 3배가 된다. 최대부하 시간대에 몰리므로 단가도 높다.
#
# 이 구조를 MILP 에 넣기 위해, 요금을 '일일 소비전력(kWh/d)에 대한 선형식'으로
# 미리 적분해 둔다. 계절·부하 루프는 설계와 무관한 상수이므로 import 시점에
# 한 번만 계산하면 된다.
#
#   P24 = e24 / 24 [kW],  P8 = e8 / 8 [kW]
#   전력량요금[￦/y] = Σ_계절 Σ_부하 (P24·h24·days + P8·h8·days) · 단가
#                   = P24·A24 + P8·A8
#   ->  e24·(A24/24) + e8·(A8/8)      (e = kWh/day)
#
# (이름, 일수, [(24h설비 가동시간, 8h설비 가동시간, 단가 ￦/kWh), ...부하 3단계])
ELEC_SEASONS = [
    ("여름",   92,  [(10, 0, 115.1), (8, 7, 163.2), (6, 1, 216.6)]),
    ("봄가을", 153, [(10, 0, 115.1), (8, 7, 132.1), (6, 1, 142.6)]),
    ("겨울",   120, [(10, 0, 122.4), (8, 5, 163.4), (6, 3, 193.4)]),
]
ELEC_BASE_KRW_PER_KW_MONTH = 9810.0  # 기본요금 [￦/kW/월]

# 계약전력 산정 방식
#   "fixed" : ELEC_CONTRACT_KW 고정 (capex_selector.py 원본 가정)
#   "peak"  : 모델이 계산한 실제 피크전력(P24 + P8)에 연동
#
# 주의: 원본 가정인 300 kW 는 태백시 실측 피크(P24 19.6 + P8 51.6 = 71 kW)의
# 4배다. 기본요금이 설계와 무관한 상수가 되어 LCOE 를 약 0.046 $/kWh 부풀린다
# (설계 순위는 안 바뀌지만 절대값이 왜곡된다). "peak" 로 두면 기본요금도
# 설계에 반응하고, 8시간 가동 설비를 줄이려는 압력이 생긴다.
ELEC_CONTRACT_MODE = "fixed"
ELEC_CONTRACT_KW = 300.0             # ELEC_CONTRACT_MODE="fixed" 일 때만 사용
ELEC_BASE_CHARGE_KRW_Y = ELEC_BASE_KRW_PER_KW_MONTH * ELEC_CONTRACT_KW * 12.0
ELEC_CLIMATE_KRW_PER_KWH = 9.0       # 기후환경요금
ELEC_FUEL_ADJ_KRW_PER_KWH = 5.0      # 연료비조정요금


def _tou_coefficients():
    """(24h계수, 8h계수, 24h연간시간, 8h연간시간) 반환.

    계수 단위는 [￦/year] per [kWh/day] 이므로, 모델의 일일 소비전력에
    그대로 곱하면 연간 전력량요금이 된다.
    """
    a24 = a8 = h24_y = h8_y = 0.0
    for _name, days, loads in ELEC_SEASONS:
        for h24, h8, price in loads:
            a24 += h24 * days * price
            a8 += h8 * days * price
            h24_y += h24 * days
            h8_y += h8 * days
    return a24 / 24.0, a8 / 8.0, h24_y, h8_y


(TOU_KRW_PER_KWHD_Y_24H, TOU_KRW_PER_KWHD_Y_8H,
 _TOU_H24_PER_YEAR, _TOU_H8_PER_YEAR) = _tou_coefficients()

# 검산: 24h 설비는 연 8,760시간, 8h 설비는 연 2,920시간이어야 한다.
assert abs(_TOU_H24_PER_YEAR - 8760.0) < 1e-6, _TOU_H24_PER_YEAR
assert abs(_TOU_H8_PER_YEAR - 2920.0) < 1e-6, _TOU_H8_PER_YEAR

# 기후환경 + 연료비조정은 시간대와 무관하게 소비량에 비례한다.
TOU_SURCHARGE_KRW_PER_KWH = ELEC_CLIMATE_KRW_PER_KWH + ELEC_FUEL_ADJ_KRW_PER_KWH

# ---------------------------------------------------------
# 설비별 OPEX 원단위 — 전부 New4.xlsx 'Data_OPEX' 시트 값
# ---------------------------------------------------------
# 전력 원단위 [kWh/ton] (원료·유량 기준)
EL_CONVEYOR = 0.0165
EL_AM_SCREEN = 0.5
EL_AM_CYCLONE = 0.29
EL_AM_DRUM_FILTER = 0.5
EL_SL_GRAVITY = 0.33
EL_SL_ROTARY_DRUM = 0.14
EL_FW_SCREEN = {"drum": 0.157, "screw": 0.13, "step": 0.157}
EL_FW_SORTING = 37.0
EL_FW_SHREDDER = {
    "screening": 0.019,
    "rotary": 34.0,
    "two_shaft": 3.1,
    "hammer": 31.0,
}
EL_BM_MIXER = 0.19
EL_AD = 3.35
EL_BELT_PRESS = 0.1
EL_SCREW_PRESS = 0.019
EL_AIR_CURTAIN_KW = 0.177          # 상시 [kW]

# 가스계통 전력 원단위 [kWh/m3 gas]
EL_DEHUM = {"Cooling": 0.0033, "Compression_cooling": 0.1}
EL_GAS_COMPRESSOR = 0.1

# 폐수처리 전력 원단위 [kWh/m3]
# [v82-C1][v82-C2] New4.xlsx Data_OPEX 열 96~109 원자료 그대로.
#
#   98  Anaerobic tank (A2O)             0.23  kWh/m3
#   99  Anoxic tank (A2O)                 -     (메탄올·소석회만)
#  100  Aerobic tank (A2O)               0.412 kWh/m3
#  101  Secondary Clarifier (A2O)        0.008 kWh/m3
#  102  Anaerobic tank (Bardenpho)       0.23  kWh/m3
#  103  Anoxic tank (Bardenpho) 1단       -     (메탄올·소석회만)
#  104  Aerobic tank (Bardenpho)         0.412 kWh/m3
#  105  Anoxic tank (Bardenpho) 2단       -     (메탄올만)
#  106  Aerobic tank (Bardenpho) 2단      -     (원단위 없음)
#  107  Secondary Clarifier (Bardenpho)  0.008 kWh/m3
#  108  DAF                              0.05  kWh/m3   <- 공법 무관 공통
#  109  UV system                        0.1   kWh/m3   <- 공법 무관 공통
#
# v81 은 0.39/0.12/0.33 을 썼는데 원자료 어디에도 없는 값이고,
# A2O 에서 호기조 0.412 가 빠져 있었다.
EL_WWTP_COMMON = 0.05 + 0.1        # DAF + UV = 0.150, 세 공법 모두 부과
EL_WWTP = {
    # Anammox 반응조 전기는 원단위가 아니라 고정비(KRW/y)로 따로 잡는다.
    "Anammox": 0.0,
    "A2O": 0.23 + 0.412 + 0.008,            # = 0.650  (+공통 0.150 = 0.800)
    "Bardenpho": 0.23 + 0.412 + 0.008,      # = 0.650  (+공통 0.150 = 0.800)
}
# ---------------------------------------------------------
# [v69] Anammox 반응조 건설비 — 상수에서 유량 1차식으로 변경
#   공사비 [억원] = 0.3148 x Q + 9.17      (Q = 폐수 설계유량 [m3/d])
#   USD 환산      = (0.3148 x Q + 9.17) x 1e8 / 환율 x 0.229
#   0.229 = 1 - 0.771 (New4.xlsx 부대비용 12개 항목 합계 77.1%) = 구입기기비 비율
#   엑셀의 A2O/Bardenpho 는 이미 구입기기비라 0.229 를 곱하지 않는다.
#   v68 은 2,967,442 USD 고정이었다 (이 식에 Q=106 을 넣은 값과 같다).
#   1차식이라 구간선형화가 필요 없다. 선형 제약 한 줄이면 된다.
# ---------------------------------------------------------
# ---------------------------------------------------------
# [v72] 폐수처리 공정 비용식 — New4.xlsx Data_CAPEX 최신본 반영
#   설계유량 V = 폐수발생량 x 1.3  (엑셀 CE3/CE9~CE20 = Base!AE2*1.3)
#
#   혐기조·무산소조 : 1246   x V^0.71   x (CEPCI_now / 357.6)
#   호기조          : 1114.1 x V^0.8324 x (CEPCI_now / 386.5)
#   2차침전지       : -0.0006 V^2 + 98.952 V + 191806      (지수 없음)
#   균등조          : 10000 x (V/18)^0.6
#   DAF             : 20000 고정
#   UV              : 7500  x (V/24)^0.6
#   출처 표기: Watertap, CEPCI 기준
#
#   A2O       = 균등조 + 혐기 + 무산소 + 호기 + 2차침전 + DAF + UV
#   Bardenpho = 균등조 + 혐기 + 무산소 + 호기 + 무산소 + 호기 + 2차침전 + DAF + UV
#   Anammox   = 균등조 + Anammox반응조 + DAF + UV
# ---------------------------------------------------------
WWTP_DESIGN_MARGIN = 1.3           # 설계여유 (엑셀 Base!AE2*1.3)
CEPCI_NOW = 797.9                  # Base!C32
CEPCI_ANAEROBIC = 357.6            # Base!C28 (혐기조·무산소조)
CEPCI_AEROBIC = 386.5              # Base!C30 (폭기조)
WWTP_ANAEROBIC_A, WWTP_ANAEROBIC_B = 1246.0, 0.71
WWTP_AEROBIC_A, WWTP_AEROBIC_B = 1114.1, 0.8324
WWTP_CLARIFIER_C2 = -0.0006
WWTP_CLARIFIER_C1 = 98.952
WWTP_CLARIFIER_C0 = 191806.0

ANAMMOX_TANK_SLOPE_EOK_PER_M3D = 0.3148     # 억원 per (m3/d)
ANAMMOX_TANK_INTERCEPT_EOK = 9.17           # 억원
# 위 식은 총공사비 기준이라 장비구입비로 환산하는 계수를 곱한다.
# 모델의 다른 CAPEX 항목은 전부 장비구입비 기준이므로 기준을 맞춰야 한다.
ANAMMOX_TANK_COST_FACTOR = 0.229
KRW_PER_EOK = 1.0e8                         # 1억원

# =========================================================
# [v76] 표준사업비 부대비용 — capex_selector.py 와 기준을 맞춘다
# =========================================================
# v75 까지 model.Cost 는 '구입기기비'만이었다. 그런데 capex_selector.py 는
#   C_agg = 구매장비비 + 표준사업비 x 77.1%
# 로 잡는다(971행). 태백시 기준 부대비용이 17,792,271 USD 로 구입기기비
# 4,298,759 USD 의 4.1 배라, 이걸 빼면 LCOE 가 4.3 배 과소평가된다.
#
#   표준사업비 [￦] = (2.7322 x 총원료[t/d] + 79.56) x 1억원      (Base!D4)
#   부대비용   [$] = 표준사업비 x 0.771 / 환율
#     0.771 = New4.xlsx 부대비용 12개 항목(설계·감리·토목·배관·시운전 등) 합계
#
# 이 값은 '총 원료량만의 함수'이므로 설계 선택과 무관한 지역별 상수다.
# 따라서 제약에 상수로 더하기만 하면 되고 MILP 구조는 그대로다.
#
# [기준 일관성] Anammox 반응조의 ANAMMOX_TANK_COST_FACTOR(0.229) 는 그대로 둔다.
#   원식(0.3148Q + 9.17 억원)은 '총공사비'라 자체 부대비용을 이미 포함한다.
#   여기서 0.229 를 빼고 총공사비로 쓰면 아래 부대비용과 이중계상이 된다.
#   즉 model.Cost 는 전 항목이 구입기기비 기준으로 통일되어 있고,
#   부대비용은 플랜트 전체에 대해 한 번만 더해진다.
INCLUDE_INDIRECT_COST = True                # False 면 v75 와 완전히 동일하게 동작
STD_PROJECT_SLOPE_EOK_PER_TON = 2.7322      # 억원 per (t/d)
STD_PROJECT_INTERCEPT_EOK = 79.56           # 억원
STD_COST_INDIRECT_RATE = 0.771              # 부대비용 12개 항목 합계 비율


def standard_project_cost_krw(total_ton_d):
    """표준사업비 [￦] = (2.7322 x 총원료[t/d] + 79.56) x 1억원."""
    return (
        STD_PROJECT_SLOPE_EOK_PER_TON * float(total_ton_d)
        + STD_PROJECT_INTERCEPT_EOK
    ) * KRW_PER_EOK


def indirect_project_cost_usd(total_ton_d):
    """설계·감리·토목·배관·시운전 등 부대비용 [USD]. 설계 선택과 무관한 상수."""
    if not INCLUDE_INDIRECT_COST:
        return 0.0
    return (
        standard_project_cost_krw(total_ton_d)
        * STD_COST_INDIRECT_RATE
        / EXCHANGE_RATE_KRW_PER_USD
    )


# =======================================================================
# [v79b] 국고보조금 + 폐기물 반입료 — capex_selector.py 기준을 v79 로 이관
# =======================================================================
APPLY_SUBSIDY = True          # False 면 보조금 미반영(기존 v79 동작)
SUBSIDY_RATE = 0.60           # 표준사업비의 60% (실제 CAPEX 아님).
                              # 자본비(CRF)에서만 차감, O&M(f_MLC×C_agg)은 총 C_agg 기준.
APPLY_TIPPING = True          # False 면 반입료 미반영
TIPPING_FEE_KRW_PER_TON = {   # 폐기물 반입료 수익 [￦/ton]
    "animal_manure": 28000.0,
    "sludge":        180468.0,
    "food_waste":   161000.0,
}


def subsidy_usd(total_ton_d):
    """국고보조금 [USD] = 표준사업비 × SUBSIDY_RATE ÷ 환율. 설계 무관 상수."""
    if not APPLY_SUBSIDY:
        return 0.0
    return SUBSIDY_RATE * standard_project_cost_krw(total_ton_d) / EXCHANGE_RATE_KRW_PER_USD


def tipping_revenue_usd_y(W_ton_d, S_ton_d, F_ton_d):
    """폐기물 반입료 수익 [USD/year] = Σ(발생량 × 단가) × 365 ÷ 환율."""
    if not APPLY_TIPPING:
        return 0.0
    f = TIPPING_FEE_KRW_PER_TON
    krw_d = (W_ton_d * f["animal_manure"]
             + S_ton_d * f["sludge"]
             + F_ton_d * f["food_waste"])
    return krw_d * 365.0 / EXCHANGE_RATE_KRW_PER_USD


# =======================================================================
# [v79c] 탄소배출권 수익 — 바이오가스 CO2 감축분을 판매 가능한 크레딧으로 계산
#   CO2 감축량 [tCO2/y] = pred_q_gas [m3/d] × 환산계수 × 배출계수 × 365
#   탄소배출권 [₩/y]    = CO2 감축량 × 단가(₩/tCO2)
# =======================================================================
APPLY_CARBON = True
CARBON_TJ_PER_M3 = 40.0e-6            # 환산계수 [TJ/m3]
CARBON_EMISSION_TCO2_PER_TJ = 56.1   # 배출계수 [tCO2/TJ]
CARBON_PRICE_KRW_PER_TCO2 = 10000.0  # 탄소배출권 단가 [₩/tCO2]


def carbon_credit_usd_y(pred_q_gas_m3_d):
    """탄소배출권 수익 [USD/year] = CO2감축[tCO2/y] × 단가 ÷ 환율."""
    if not APPLY_CARBON:
        return 0.0, 0.0
    co2_tco2_y = (float(pred_q_gas_m3_d) * CARBON_TJ_PER_M3
                  * CARBON_EMISSION_TCO2_PER_TJ * 365.0)
    krw_y = co2_tco2_y * CARBON_PRICE_KRW_PER_TCO2
    return co2_tco2_y, krw_y / EXCHANGE_RATE_KRW_PER_USD


# Anammox 는 유량무관 고정비 [￦/year]
WWTP_ANAMMOX_ELEC_KRW_Y = 137_340_000.0
# [v82-C4] 황산비 — New4.xlsx Data_OPEX 의 Anammox 열(97)은 '황' 행이 비어 있어
#   이 값의 근거를 찾지 못했다. capex_selector.py 에도 없다.
#   근거 확인 전까지 기본 비활성. 되살리려면 아래 플래그만 True 로 바꾸면 된다.
#   (참고: 같은 열 'etc' 행에 508,000 KRW/kg-N 이 있는데 두 코드 모두 미반영이다.)
APPLY_ANAMMOX_SULFUR = False
_WWTP_ANAMMOX_SULFUR_KRW_Y_RAW = 8_760_000.0
WWTP_ANAMMOX_SULFUR_KRW_Y = (
    _WWTP_ANAMMOX_SULFUR_KRW_Y_RAW if APPLY_ANAMMOX_SULFUR else 0.0
)

# =========================================================
# [v77] 폐수처리 약품비 — capex_selector.py 883~893행 이관
# =========================================================
# v75/v76 은 "유입 질소농도와 약품 단가가 자료에 없어 계산 불가"로 미반영이었다.
# 이제 둘 다 확보됐으므로 반영한다. 태백 실측(TN=2833 mg/L, 발생량 92 t/d)에서
# 메탄올+소석회가 460,303 $/y 로 변동 OPEX 의 85% 를 차지한다. 빼놓을 수 없다.
#
#   질소부하 N [tonN/d] = TN[mg/L] x 폐수량[m3/d] x 1e-6
#   메탄올 : 2.47 tonMeOH/tonN x 340 $/ton      (탈질 외부탄소원)
#   소석회 : 7.14 tonCa(OH)2/tonN x 950 ￦/kg   (질산화 알칼리도 보충)
#   Alum   : 100 mg/L x 폐수량 x 778.8 ￦/kg    (DAF 응집제, TN 과 무관)
#
# [중요] 메탄올·소석회는 A2O/Bardenpho 에만 든다. Anammox 는 부분아질산화-
#   혐기성암모늄산화라 외부탄소원과 알칼리도 보충이 원리적으로 불필요하다
#   (capex_selector 878행 주석과 동일). 대신 전기 고정비가 붙는다.
#   -> v76 까지는 약품비가 0 이라 Anammox 가 구조적으로 불리했는데,
#      이 항이 들어가면서 비로소 공정 선택이 실제 경쟁이 된다.
#
# TN 은 지역별 실측이 없으므로 상수로 둔다. 엑셀에 열이 생기면
# WWTP_TN_COL 을 지정해 행별로 받도록 바꾸면 된다.
WWTP_TN_MG_L = 2833.0               # 소화액 탈리액 TN [mg/L]
METHANOL_USD_PER_TON = 340.0
METHANOL_TON_PER_TON_N = 2.47
LIME_KRW_PER_KG = 19000.0 / 20.0    # 소석회 950 ￦/kg (19,000￦/20kg)
LIME_TON_PER_TON_N = 7.14
ALUM_KRW_PER_KG = 19470.0 / 25.0    # Alum 778.8 ￦/kg (19,470￦/25kg)
ALUM_MG_PER_L = 100.0

# 질소부하 1 m3 폐수당 약품비 [USD/m3] — 선형계수로 미리 접어둔다.
#   N[tonN] per m3 = TN[mg/L] x 1e-6
_N_TON_PER_M3 = WWTP_TN_MG_L * 1.0e-6
_METHANOL_USD_PER_M3 = (
    _N_TON_PER_M3 * METHANOL_TON_PER_TON_N * METHANOL_USD_PER_TON
)
_LIME_USD_PER_M3 = (
    _N_TON_PER_M3 * LIME_TON_PER_TON_N
    * LIME_KRW_PER_KG * 1000.0 / EXCHANGE_RATE_KRW_PER_USD
)
# [v82-C3] Data_OPEX 열 99·103·105 기준 — 약품은 무산소조 '단수'에 비례한다.
#   A2O       : 무산소 1단        -> 메탄올 x1, 소석회 x1
#   Bardenpho : 무산소 2단(103,105) -> 메탄올 x2, 소석회 x1
#               (105 열에는 메탄올만 있고 소석회는 비어 있다)
#   v81 과 capex_selector.py 모두 Bardenpho 메탄올을 1배만 넣고 있었다.
WWTP_METHANOL_STAGES = {"A2O": 1, "Bardenpho": 2}
WWTP_LIME_STAGES = {"A2O": 1, "Bardenpho": 1}
WWTP_CHEM_N_USD_PER_M3 = {
    _o: (_METHANOL_USD_PER_M3 * WWTP_METHANOL_STAGES[_o]
         + _LIME_USD_PER_M3 * WWTP_LIME_STAGES[_o])
    for _o in ("A2O", "Bardenpho")
}
# Alum 은 공법과 무관하게 DAF 에 공통으로 든다.
WWTP_CHEM_ALUM_USD_PER_M3 = (
    ALUM_MG_PER_L * 1000.0 / 1.0e6 * ALUM_KRW_PER_KG / EXCHANGE_RATE_KRW_PER_USD
)
# 메탄올·소석회가 드는 공법
WWTP_OPTIONS_NEEDING_N_CHEM = ("A2O", "Bardenpho")

# [v83] Anammox 변동 OPEX — 전기·황 대신 질소부하 기반 'etc' 508 원/kg-N 으로 교체.
#   New4.xlsx Data_OPEX 의 Anammox 열 'etc' 행 값을 508 KRW/kg-N 로 반영한다.
#   메탄올·소석회는 원리상 불필요하므로 Anammox 는 이 항만 든다(전기·황 제외).
#     kg-N per m3 = TN[mg/L] x 1e-3       (2833 mg/L -> 2.833 kg-N/m3)
#     USD per m3  = kg-N/m3 x 508 KRW/kg-N / 환율
WWTP_ANAMMOX_ETC_KRW_PER_KG_N = 508.0
WWTP_ANAMMOX_ETC_USD_PER_M3 = (
    WWTP_TN_MG_L * 1.0e-3
    * WWTP_ANAMMOX_ETC_KRW_PER_KG_N
    / EXCHANGE_RATE_KRW_PER_USD
)

# 기술별 변동 OPEX
OPEX_DESULF_USD_PER_M3 = {
    "wet_desulfurization": 0.0225,
    "dry_desulfurization": 0.014,
}
OPEX_PURIFICATION_USD_PER_M3_CH4 = {
    "psa": 0.144,
    "membrane": 0.194,
    "wet_scrubber": 0.162,
}

# 소화조 가온용 가스 자가소비 = 가스 생산량의 35% (Data_OPEX 'LNG(CH4)' 행).
#   중온소화(35degC) 유지를 위한 원료 승온 + 방열 보상 열량을,
#   생산한 바이오가스의 일부를 태워 직접 공급한다(자가소비).
#   -> 연료를 '사는' 것이 아니므로 OPEX 에 비용을 더하지 않는다.
#      대신 소화조 출구에서 이만큼 빼내므로 후단으로 가는(=팔 수 있는) 가스가 줄어든다.
#   AD 물질수지와 VS 분해량은 '전량 생산 기준'을 써야 하므로 별도로 환산한다.
AD_HEATING_CH4_FRAC_OF_GAS = 0.35

# 판매 에너지 최소치 [kWh/d]. 분모가 0 이면 균등화원가가 정의되지 않는다.
MIN_DELIVERED_ENERGY_KWH_D = 1.0e-3
DAYS_PER_YEAR = 365.0 * LCOB_F_UTIL

# surrogate 신뢰 하한 [m3/d]. 학습데이터 재현오차가 q_ad>20 에서 ±2%,
# q_ad<=10 에서 -46% 이므로 20 을 기준으로 삼는다.
SURROGATE_MIN_Q_AD = 20.0
# 학습데이터의 정상 비수율 = 가스/(q_ad x COD). COD 130~160, OLR 3~3.8 구간 평균.
SURROGATE_REF_YIELD = 0.47

# 비용상한이 '걸렸다'고 볼 허용오차 [USD].
COST_CAP_BINDING_TOL = 1.0

# 미반영: 폐수처리 약품비(메탄올 2.47 ton/tonN, 소석회 7.14 ton/tonN,
#   Alum 100 mg/L). 유입 질소농도와 약품 단가가 자료에 없어 계산 불가.


RAW_RANGE_TOL = 1e-2

# =========================================================
# Unit assumptions
# =========================================================
# 원자료 W, F, S의 단위는 t/d이다.
# 원료와 소화슬러지 밀도는 모두 1,000 kg/m3로 가정한다.
# 따라서 수치상 1 t/d = 1 m3/d이다.
#
# Q_* 변수는 m3/d 기준으로 사용하고,
# AD 물질수지는 kg/d 기준으로 계산한다.
RAW_FEED_DENSITY_KG_M3 = 1000.0
DIGESTATE_DENSITY_KG_M3 = 1000.0
TON_TO_KG = 1000.0

# 표준상태 건식가스 밀도 [kg/m3]
CH4_GAS_DENSITY_KG_M3 = 0.716
CO2_GAS_DENSITY_KG_M3 = 1.977

# =========================================================
# 원료별 총 COD 농도 [kgCOD/m3]
# =========================================================
# [v79] 값 정정. v78 까지 40.79 / 6.66 / 0.75 를 썼는데, 이는
#   Paper_feed721.xlsx Sheet1 'Feed ratio' 블록의 'COD(kg)' 열을 그대로
#   옮긴 것이고 그 열의 계산식은 다음과 같다.
#
#       COD(kg) = Ratio x TS_희석 x (VS/TS) x (COD/VS)
#                  ^^^^^
#                  혼합비가 이미 곱해져 있다
#
#   즉 40.79 의 단위는 [kgCOD / '7:2:1 혼합물' 1 m3] 이지
#      [kgCOD / '축분' 1 m3] 이 아니다. 셋을 더하면 48.21 이 되고
#      그것이 Sheet1 이 구하려던 혼합물의 COD 농도(13행 'COD 48.2 g/kg')다.
#
#   코드는 이 값을 지역별 원료유량에 곱했다(physical_COD_ani = W x 40.79).
#   '축분 W m3' 에 '혼합물 1m3당 기여분' 을 곱한 셈이라 단위가 맞지 않고,
#   이미 반영된 비율(0.7/0.2/0.1)을 한 번 더 곱하는 결과가 된다.
#
#   지역마다 원료 조성이 다르므로 오차 크기도 지역마다 달라진다(균일 편향이
#   아니라 보정계수로 덮을 수 없다). 슬러지는 비율 0.1 로 나뉜 값이라 가장 심하다.
#       축분만 있는 지역   : 1.43 배 과소
#       음식물만 있는 지역 : 11.4 배 과소
#       슬러지만 있는 지역 : 81.4 배 과소
#       태백시(슬러지 50%) :  3.6 배 과소
#
#   상수는 혼합비와 무관한 '물성값' 이어야 한다. Ratio 를 제거하고
#   모델이 가정하는 TS(=0.05)를 적용해 다시 계산한다.
#
#       COD = TS x 1000 x (VS/TS) x (COD/VS)
#
#   VS/TS 와 COD/VS 는 Sheet1 의 원료별 물성이며 배합과 무관하다.
#
#   [영향] 이 상수들은 열역학 검산(ch4_yield_exceeds_theoretical)과 보고용
#   열에만 쓰이고 NN 입력에는 들어가지 않는다(NN 입력은 SURROGATE_COD_COEF).
#   따라서 생산량·비용·LCOE 결과값은 바뀌지 않고, 검산 경고만 해소된다.
#   태백시 기준 메탄 수율이 이론최대의 2.07 배(위반) -> 0.57 배(정상)가 된다.
#
#   [남은 한계] SURROGATE_COD_COEF(75.76)는 Sheet2 가 전 원료에 음식물 물성
#   (VS/TS 0.947, COD/VS 1.6)을 일괄 적용해 만든 값이다. 즉 surrogate 는
#   축분(58.29)과 슬러지(61.08)도 음식물로 간주하므로 가스를 20~30% 과대
#   추정하는 방향이다. 원료별 물성으로 재학습하기 전까지 남는 오차다.
TS_FEED_FRACTION = 0.05          # 모델 가정 (TS_ani = TS_fw = TS_sl = 0.05)

# Paper_feed721.xlsx Sheet1 — 원료별 물성 (혼합비와 무관)
VSTS_ANI, CODVS_ANI = 0.783, 1.489     # 가축분뇨
VSTS_FW, CODVS_FW = 0.947, 1.600       # 음식물류
VSTS_SL, CODVS_SL = 0.722, 1.692       # 하수슬러지

COD_ANI_KG_M3 = TS_FEED_FRACTION * 1000.0 * VSTS_ANI * CODVS_ANI   # 58.29
COD_FW_KG_M3 = TS_FEED_FRACTION * 1000.0 * VSTS_FW * CODVS_FW      # 75.76
COD_SL_KG_M3 = TS_FEED_FRACTION * 1000.0 * VSTS_SL * CODVS_SL      # 61.08

# v78 까지의 값. 재현 비교가 필요하면 위 세 줄을 이걸로 바꾼다.
COD_LEGACY_BLEND_BASIS = (40.79, 6.66, 0.75)

# =========================================================
# Surrogate input feature basis
# =========================================================
# 중요:
# 현재 Keras/OMLT surrogate는 아래 공통 COD 정의로 생성한 입력으로 학습됐다.
# 원료별 물리 COD(40.79/6.66/0.75)를 NN 입력에 직접 넣으면
# COD_total 및 ADM 상태가 학습범위와 완전히 달라져 OMLT 입력 bounds와
# 충돌하면서 비용상한을 제거해도 모델 자체가 infeasible이 된다.
#
# 물리 COD 질량은 별도 변수(COD_*_kg_d)로 유지하고,
# NN 입력용 COD/ADM feature만 학습 당시 기준을 그대로 사용한다.
SURROGATE_MOISTURE = 0.95
SURROGATE_TS_KG_M3 = 132.0
SURROGATE_VS_KG_M3 = 125.0
SURROGATE_COD_KG_M3 = 200.0
SURROGATE_COD_COEF = (
    (1.0 - SURROGATE_MOISTURE)
    * (SURROGATE_VS_KG_M3 / SURROGATE_TS_KG_M3)
    * (SURROGATE_COD_KG_M3 / SURROGATE_VS_KG_M3)
    * 1000.0
)

# =========================================================
# 혐기소화 화학량론 (물리 검증 및 고형물 수지용)
# =========================================================
# 이론 최대 메탄수율 [m3 CH4 / kgCOD 제거] — 표준상태(0degC, 1atm) 기준.
#   35degC 기준으로 보려면 0.40 으로 바꾼다.
# COD 제거량은 '유입 COD' 이하이므로, CH4/유입COD 가 이 값을 넘으면
# 열역학적으로 불가능한 해라는 뜻이다.
CH4_YIELD_M3_PER_KGCOD = 0.35
# COD/VS 환산 [kgCOD / kgVS] — 유기물 1 kg 이 갖는 COD
COD_PER_VS_KG_PER_KG = 1.5


# 소화슬러지 탈수 가정
# 주의: v48~v50 은 소화액 TS 를 5% 로 '가정'했고, 그 결과 고형물 수지가
#       닫히지 않았다(원료 고형물 - 바이오가스 - 소화액 고형물 = 최대 -4,260 kg/d).
#       v51 은 소화액 고형물을 VS 분해량으로부터 계산해 수지를 닫는다.
#       따라서 소화액 TS 는 가정값이 아니라 계산 결과이며,
#       digestate_TS_computed 열로 출력된다.
CAKE_MOISTURE = 0.80
CAKE_TS_FRACTION = 1.0 - CAKE_MOISTURE
SOLIDS_CAPTURE_EFFICIENCY = 0.95
CAKE_DENSITY_KG_M3 = 1000.0
CENTRATE_DENSITY_KG_M3 = 1000.0

# =========================================================
# ★★★ 판매 단가 (수익 계산용) — 나중에 이 두 값만 바꾸면 됩니다 ★★★  (모두 USD)
# =========================================================
# 도시가스(바이오메탄) 판매단가 [USD / m3 CH4]
#   참고: 2025년 한국 천연가스 ~0.065 USD/kWh, CH4 발열량 9.94 kWh/m3 → ~0.65
CITY_GAS_PRICE_USD_PER_M3_CH4 = 0.65
# 에너지 기준 단가 [USD/kWh]. 도시가스(LNG)와 바이오메탄은 발열량이 달라
# 부피 1:1 대체가 아니므로, 수요 대조는 전부 에너지(kWh) 기준으로 한다.
CITY_GAS_PRICE_USD_PER_KWH = CITY_GAS_PRICE_USD_PER_M3_CH4 / 9.94
# 전기 판매단가 [USD / kWh]
#   참고: 2025년 한국 SMP ~130 KRW/kWh, 환율 ~1,350 → ~0.10
ELEC_PRICE_USD_PER_KWH = 0.10

# OMLT 입력 bounds
# 학습범위를 그대로 강제하면 일부 지역이 구조적으로 infeasible이 될 수 있어
# scaled 범위를 제한적으로 확장한다. 단, 과도한 외삽은 막기 위해 [-8, 8]로 제한.
NN_BOUND_MIN_EXPANSION_STD = 1.0
NN_BOUND_SPAN_EXPANSION = 0.50
NN_BOUND_ABS_LIMIT = 8.0

# Reactor CAPEX — 고정 모듈형 AD 기준
# AD 1기 기준: 처리용량 50 m3/d, 반응조 부피 2,300 m3, 건설비 24억원
# 24억원 ≈ 1,600,910 USD (적용 환율 기준값)
# AD는 2계열 병렬이며 각 계열에 최소 1기씩 설치한다.
# 각 계열 유량이 50 m3/d를 초과하면 해당 계열에 AD 모듈을 추가한다.
# 주의: v48에는 REACTOR_VOLUME_REF(2000) / REACTOR_SCALE_EXPONENT(0.67) 도 있었으나
#       비용식에 전혀 쓰이지 않는 死상수라 삭제했다. AD CAPEX 는 계열수 x 기준비용의
#       순수 계단함수이며, 반응조 부피(V_reactor)는 비용에 영향을 주지 않는다.
REACTOR_CAPEX_REF = 1_600_910.0
REACTOR_FLOW_REF = 50.0

# AD 반응조 교반기(vertical mixer): AD 1기당 개수와 1개당 기준비용 [USD]
AD_MIXER_PER_TRAIN = 4
AD_MIXER_REF_COST = 660.0
# AD 계열 수는 Q_ad_in/50에 따라 모델 내부에서 자동 결정

# 발전 관련 상수
CH4_LHV_KWH = 9.94      # 메탄 저위발열량 [kWh/m3]
GEN_ELEC_EFF = 0.40     # 발전 전기효율
NN_INPUTS = [
    "q_ad(m3/d)",
    "total_influent_COD(kgCOD/m3)",
    "input_S_su(kgCOD/m3)",
    "input_S_aa(kgCOD/m3)",
    "input_S_fa(kgCOD/m3)",
    "input_S_I(kgCOD/m3)",
    "input_X_ch(kgCOD/m3)",
    "input_X_pr(kgCOD/m3)",
    "input_X_li(kgCOD/m3)",
    "input_X_I(kgCOD/m3)",
    "OLR(kgCOD/m3/d)",
]


# =======================================================================
# §2  설비 원단위 · 비용 상수  (build_model 에서 승격)
# =======================================================================

TS = 0.05

TS_ani = 0.05

TS_fw = 0.05

TS_sl = 0.05

water_removal_ani = [0.3, 0.5]

# 전처리 완료 후 AD로 이송되는 feeding hopper 및 conveyor
# 반입설비와는 별도이며, 비용은 원료 총유량 W/F/S가 아니라
# 실제 전처리 완료 유량 Q_ani/Q_sl/Q_fw를 기준으로 Piecewise 계산
PRETREATMENT_AUX_COST_DATA = {
    "AM_feeding_hopper": (2000.0, 2.0, 0.6),
    "SL_feeding_hopper": (2000.0, 2.0, 0.6),
    "FW_feeding_hopper": (2000.0, 2.0, 0.6),
    "AM_conveyor_1": (4500.0, 360.0, 0.6),
    "AM_conveyor_2": (4500.0, 360.0, 0.6),
    "SL_conveyor_1": (4500.0, 360.0, 0.6),
    "SL_conveyor_2": (4500.0, 360.0, 0.6),
    "FW_conveyor_1": (4500.0, 360.0, 0.6),
    "FW_conveyor_2": (4500.0, 360.0, 0.6),
}

# screen에서만 등가 수분 제거율을 적용
# sorting machine과 shredder는 유량 손실 없이 비용만 반영
FW_SCREEN_WATER_REMOVAL = {
    "drum": 0.3,
    "screw": 0.2,
    "step": 0.1,
}

FW_SORTING_MACHINE_1_MIN_COST = 10_000

FW_SORTING_MACHINE_1_MIN_VOL = 2

FW_SORTING_MACHINE_1_SF = 0.6

FW_SORTING_MACHINE_2_MIN_COST = 10_000

FW_SORTING_MACHINE_2_MIN_VOL = 2

FW_SORTING_MACHINE_2_SF = 0.6

FW_SCREENING_SHREDDER_MIN_COST = 5_000

FW_SCREENING_SHREDDER_MIN_VOL = 9_600

FW_ROTARY_SHREDDER_MIN_COST = 9_000

FW_ROTARY_SHREDDER_MIN_VOL = 12

FW_ROTARY_SHREDDER_SF = 0.6

FW_TWO_SHAFT_SHREDDER_MIN_COST = 23_700

FW_TWO_SHAFT_SHREDDER_MIN_VOL = 72

FW_HAMMER_CRUSHER_MIN_COST = 7_500

FW_HAMMER_CRUSHER_MIN_VOL = 19

FW_HAMMER_CRUSHER_SF = 0.6

# screen 3은 두 계열 합류 후 실제 유입유량을 기준으로 계산
FW_SCREEN3_COST_DATA = {
    "drum": (17000.0, 48.0, 0.6),
    "screw": (5500.0, 480.0, 0.6),
    "step": (1800.0, 96.0, 0.6),
}

# 스크린에서 분리된 폐기물 배출 설비
# screen 1·2 및 screen 3에서 빠져나온 총 waste 유량을 모두 처리
FW_WASTE_HANDLING_COST_DATA = {
    "discharge_hopper": (2000.0, 2.0, 1.5),
    "conveyor_1": (4500.0, 360.0, 0.6),
    "conveyor_2": (4500.0, 360.0, 0.6),
}

water_removal_sl = [0.20, 0.4]

ADM_STATES = ["X_I", "X_ch", "X_pr", "X_li", "S_I", "S_su", "S_aa", "S_fa"]

RATIO_ANI = {"X_I":0.1617886114462387,"X_ch":0.5328739910019253,"X_pr":0.125057590035531,"X_li":0.03188506512808813,"S_I":0.02819076750576085,"S_su":0.09285877534817646,"S_aa":0.02179298600808249,"S_fa":0.005552213526197075}

RATIO_FW = {"X_I":0.0837120524969625,"X_ch":0.24524685131930646,"X_pr":0.14928100956956466,"X_li":0.11809416159247088,"S_I":0.05631793637394176,"S_su":0.16617323438612126,"S_aa":0.10114798045309979,"S_fa":0.08002677380853272}

RATIO_SL = {"X_I":0.38429249746518954,"X_ch":0.08498339415934248,"X_pr":0.34959444804334835,"X_li":0.18292758460536516,"S_I":0.0005773037291433472,"S_su":0.0001276676619417676,"S_aa":0.0005251806796803454,"S_fa":0.0002745283750804516}

AD_SWITCH_EPS = 1.0e-6

# Biomass pulping (ref=240, floor). 전 구간 ref 이하이면 상수 → Piecewise 경고 방지.
_PULP_REF = 240.0

# =========================================================
# Gas purification process
# =========================================================
tanks = ["membrane_type", "piston_deck_type", "steel_storage", "wet_structure"]

desulf_techs = ["wet_desulfurization", "dry_desulfurization"]

dehum_techs = ["Cooling", "Compression_cooling"]

# Gas compressor는 공통 필수설비이고,
# 정제기술은 PSA 또는 membrane 중 정확히 하나만 선택
purification_techs = ["psa", "membrane", "wet_scrubber"]

# 후단 설비 유량 하한도 0으로 두어 저유량 surrogate 출력을 허용
QGAS_LB = 0.0

# =========================================================
# 가스 전처리 비용 공통 함수
#
# Q <= min_vol : min_cost
# Q >  min_vol : min_cost * (Q / min_vol)^SF
#
# 각 공정은 후보 기술 중 정확히 하나만 선택한다.
# =========================================================
# ---------------------------------------------------------
# 1) 탈황공정: wet / dry 중 택 1
# 엑셀 기준값 반영
# ---------------------------------------------------------
desulf_cost_data = {
    # (min_cost USD, min_vol m3/d, SF)
    "wet_desulfurization": (50000.0, 960.0, 0.6),
    "dry_desulfurization": (5000.0, 24.0, 0.6),
}

# ---------------------------------------------------------
# 2) 제습공정: Cooling / Compression cooling 중 택 1
#
# Cooling은 첨부 엑셀값 반영:
# min_vol=2400, min_cost=10849, SF=0.6
# Compression cooling은 기존 기준값 유지
# ---------------------------------------------------------
dehum_cost_data = {
    "Cooling": (10849.0, 2400.0, 0.6),
    "Compression_cooling": (16200.0, 1200.0, 0.6),
}

# Gas storage cost (floor 적용)
tank_cost_data = {
    "membrane_type": (8800, 100, 0.6),
    "piston_deck_type": (9999, 50, 0.6),
    "steel_storage": (9999, 50, 0.6),
    "wet_structure": (5000, 50, 0.6),
}

# =========================================================
# 잉여가스 연소기 (Waste Gas thermal oxidizer) — 필수 설비
# 엑셀: ref_vol=50 m3/hr, ref_cost=7,000, SF=0.6,
#       설계유량 i_vol[m3/hr] = (Total_Gas[m3/d]/24)*2.5
# 일유량 기준 환산: ref_vol_daily = 50*24/2.5 = 480 m3/d
# 총 바이오가스 dehum_out_q_gas 에 대해 ref_vol=480 으로 계산.
# =========================================================
THERMAL_OXIDIZER_MIN_COST = 7000.0

THERMAL_OXIDIZER_MIN_VOL = 480.0

THERMAL_OXIDIZER_SF = 0.6

# =========================================================
# 3) 정제공정: 4개 기술 중 정확히 1개 선택
#
# 현재 보유 비용을 기준값으로 사용하고, 모든 기술에
# cost-capacity 식을 동일하게 적용한다.
# 추후 실제 min_vol/min_cost/SF 자료로 이 dictionary만 교체하면 된다.
# =========================================================
# Gas compressor는 PSA/membrane 선택과 무관하게 반드시 설치
GAS_COMPRESSOR_MIN_COST = 16200.0

GAS_COMPRESSOR_MIN_VOL = 1200.0

GAS_COMPRESSOR_SF = 0.6

purification_cost_data = {
    # PSA 또는 membrane 중 택 1
    # (min_cost USD, min_vol m3/d, SF)
    "psa": (500_000.0, 2400.0, 0.6),
    "membrane": (193_567, 1200.0, 0.6),
    "wet_scrubber": (50_000.0, 960.0, 0.6),
}

# Purification performance
ch4_recovery = {
    "psa": 0.98,
    "membrane": 0.96,
    "wet_scrubber": 0.95,
}

co2_removal = {
    "psa": 0.98,
    "membrane": 0.95,
    "wet_scrubber": 0.97,
}

product_ch4_purity = {
    "psa": 0.98,
    "membrane": 0.97,
    "wet_scrubber": 0.96,
}

# =========================================================
# 발전기 (에너지화 / CHP) — 4개 중 택 1
#
# 발전기는 PSA/membrane 정제 후의 up_ch4가 아니라,
# 탈황·제습을 마친 전체 바이오가스 dehum_out_q_gas를 처리한다.
# 현재 가스 손실 0 가정에서는 dehum_out_q_gas = pred_q_gas이다.
#
# 발전기 처리유량 [m3 biogas/d]: dehum_out_q_gas
# 발전출력 계산용 CH4 유량 [m3 CH4/d]: pur_in_ch4_flow
#   pur_in_ch4_flow ≈ pred_q_gas × pred_ch4_frac_dry
#
#   P_gen [kW]
#   = gen_ch4_flow [m3 CH4/d] × CH4_LHV_KWH [kWh/m3]
#     × 발전 전기효율 / 24 [h/d]
#
# 발전기 비용:
#
#   P <= P_min:
#       C = C_min
#
#   P > P_min:
#       C = C_min × (P/P_min)^SF
# =========================================================
gens = [
    "fuel_cell",
    "gas_turbine",
    "gas_engine",
    "diesel_engine",
]

# (min_generation_kW, min_cost_USD, scale_factor)
gen_cost_data = {
    "fuel_cell":   (5.5, 80_000.0, 0.6),
    "gas_turbine": (30.0, 25_000.0, 0.6),
    "gas_engine":  (30.0, 28_500.0, 0.6),
    "diesel_engine": (30.0, 28_500.0, 0.6),
}

# =========================================================
# 소화슬러지 탈수시설
#
# AD 소화액 유량은 별도 감소율 자료가 없으므로
# 현재는 AD 실제 유입유량 Q_ad_in과 동일하다고 가정한다.
#
# 탈수기:
#   belt press 또는 screw press 중 정확히 1개 선택
#
# 이송설비:
#   DS conveyor와 SC conveyor는 모두 필수 설치
#
# 비용식:
#   Q <= min_vol : min_cost
#   Q >  min_vol : min_cost × (Q/min_vol)^SF
# =========================================================
digestate_dewatering_techs = [
    "belt_press",
    "screw_press",
]

# 소화슬러지 탈수기 비용자료
# 첨부 표 기준: (ref_vol [m3/d], ref_cost [USD], SF)
#
# belt press  : ref_vol=3.00,  ref_cost=20,000, SF=0.60
# screw press : ref_vol=0.07,  ref_cost= 5,000, SF=0.60
digestate_dewatering_cost_data = {
    "belt_press":  (3.0, 20_000.0, 0.6),
    "screw_press": (0.07, 5_000.0, 0.6),
}

# 소화슬러지 이송설비 비용자료
# 첨부 표 기준: (ref_vol [m3/d], ref_cost [USD], SF)
#
# DS conveyor : ref_vol=360, ref_cost=4,500, SF=0.60
# SC conveyor : ref_vol=360, ref_cost=4,500, SF=0.60
digestate_conveyor_cost_data = {
    "DS_conveyor": (360.0, 4_500.0, 0.6),
    "SC_conveyor": (360.0, 4_500.0, 0.6),
}

# =========================================================
# 폐수처리 공정 — A2O 또는 Bardenpho 중 정확히 1개 선택
#
# 현재 모델에는 탈수 후 탈리액 유량을 별도로 계산할 자료가 없으므로,
# 보수적으로 AD 출구 소화액 전체 유량 digestate_flow를
# 폐수처리시설 설계유량으로 사용한다.
#
# 비용식:
#   Q <= ref_vol : ref_cost
#   Q >  ref_vol : ref_cost * (Q / ref_vol) ** SF
# =========================================================
wwtp_options = ["Anammox", "A2O", "Bardenpho"]

# (ref_vol [m3/d], ref_cost [USD], SF)
WWTP_EQUIPMENT_COST_DATA = {
    # Anammox
    "Anammox_equalization_tank": (18.0, 10_000.0, 0.6),
    # [v69] Anammox_tank 는 아래에서 1차식으로 따로 계산한다.
    #       이 자리 값은 더 이상 비용식이 아니다.
    "Anammox_tank": (1.0e12, 0.0, 0.6),
    "Anammox_DAF_system": (120.0, 20_000.0, 0.0),
    "Anammox_UV_system": (24.0, 7_500.0, 0.6),

    # A2O  [v72] 반응조 3종 + 2차침전지는 아래에서 별도 식으로 계산
    "A2O_equalization_tank": (18.0, 10_000.0, 0.6),
    "A2O_anaerobic_tank": None,
    "A2O_anoxic_tank": None,
    "A2O_aeration_tank": None,
    "A2O_secondary_clarifier": None,
    "A2O_DAF_system": (120.0, 20_000.0, 0.0),
    "A2O_UV_system": (24.0, 7_500.0, 0.6),

    # Bardenpho  [v72] 무산소조·호기조가 2단씩
    "Bardenpho_equalization_tank": (18.0, 10_000.0, 0.6),
    "Bardenpho_anaerobic_tank": None,
    "Bardenpho_anoxic_tank": None,
    "Bardenpho_aeration_tank": None,
    "Bardenpho_anoxic_tank_2": None,
    "Bardenpho_aeration_tank_2": None,
    "Bardenpho_secondary_clarifier": None,
    "Bardenpho_DAF_system": (120.0, 20_000.0, 0.0),
    "Bardenpho_UV_system": (24.0, 7_500.0, 0.6),
}

# [v72] 별도 비용식을 쓰는 설비 -> 식 종류
WWTP_SPECIAL_COST = {
    "A2O_anaerobic_tank": "anaerobic",
    "A2O_anoxic_tank": "anaerobic",
    "A2O_aeration_tank": "aerobic",
    "A2O_secondary_clarifier": "clarifier",
    "Bardenpho_anaerobic_tank": "anaerobic",
    "Bardenpho_anoxic_tank": "anaerobic",
    "Bardenpho_aeration_tank": "aerobic",
    "Bardenpho_anoxic_tank_2": "anaerobic",
    "Bardenpho_aeration_tank_2": "aerobic",
    "Bardenpho_secondary_clarifier": "clarifier",
}

WWTP_OPTION_EQUIPMENT = {
    "Anammox": [
        "Anammox_equalization_tank",
        "Anammox_tank",
        "Anammox_DAF_system",
        "Anammox_UV_system",
    ],
    "A2O": [
        "A2O_equalization_tank",
        "A2O_anaerobic_tank",
        "A2O_anoxic_tank",
        "A2O_aeration_tank",
        "A2O_secondary_clarifier",
        "A2O_DAF_system",
        "A2O_UV_system",
    ],
    "Bardenpho": [
        "Bardenpho_equalization_tank",
        "Bardenpho_anaerobic_tank",
        "Bardenpho_anoxic_tank",
        "Bardenpho_aeration_tank",
        "Bardenpho_anoxic_tank_2",
        "Bardenpho_aeration_tank_2",
        "Bardenpho_secondary_clarifier",
        "Bardenpho_DAF_system",
        "Bardenpho_UV_system",
    ],
}

# =======================================================================
# §3-4  공용 헬퍼 · NN surrogate
# =======================================================================


def add_cost_capacity_pw(model, name, cost_var, flow_var, pts,
                         ref_cost=None, ref_vol=None, exponent=None,
                         f_rule=None, warning_tol=None):
    """cost-capacity 곡선  C(Q) = ref_cost * (max(Q, ref_vol) / ref_vol) ** exponent
    를 구간선형(DCC, 등식)으로 model.<name> 에 붙인다.

    max(Q, ref_vol) 로 하한을 거는 것은 기준용량 미만에서도 최소비용 floor 를
    유지하기 위해서다. 절점 0 을 넣고 직선보간하면 floor 가 무너진다([Fix 5]).

    f_rule 을 직접 주면 그 함수를 그대로 쓴다(예: 폐수처리 A2O/Bardenpho 상관식).
    warning_tol 은 None 이면 Pyomo 기본값을 쓴다(원본 동작 보존).
    """
    if f_rule is None:
        def f_rule(m, x, c=ref_cost, v=ref_vol, n=exponent):
            return c * (max(x, v) / v) ** n

    kw = {} if warning_tol is None else {"warning_tol": warning_tol}
    setattr(
        model,
        name,
        pyo.Piecewise(
            cost_var, flow_var,
            pw_pts=pts,
            f_rule=f_rule,
            pw_constr_type="EQ",
            pw_repn="DCC",
            **kw,
        ),
    )


def add_bigm_select(cons, selected, cost_if, sel, big_m):
    """선택변수 sel 에 따라  selected = sel * cost_if  가 되도록 big-M 선형화.

    sel=1 이면 selected == cost_if, sel=0 이면 selected == 0.
    제약 추가 순서는 원본과 동일하게 (상한 M, 상한 cost, 하한) 순이다.
    """
    cons.add(selected <= big_m * sel)
    cons.add(selected <= cost_if)
    cons.add(selected >= cost_if - big_m * (1 - sel))


def get_raw_streams_from_excel_row(row):
    q_hanwoo = row["한우"]
    q_dairy = row["젖소"]
    q_pig = row["돼지"]
    q_chicken = row["치킨"]

    q_ani = q_hanwoo + q_dairy + q_pig + q_chicken
    q_sl = row["슬러지"]
    q_fw = row["음식물"]

    return {
        "Q_hanwoo_raw": q_hanwoo,
        "Q_dairy_raw": q_dairy,
        "Q_pig_raw": q_pig,
        "Q_chicken_raw": q_chicken,
        "Q_ani_raw": q_ani,
        "Q_sl_raw": q_sl,
        "Q_fw_raw": q_fw,
    }


def scale_cost(cost_ref, flow, flow_ref, n):
    flow = float(flow)
    flow_ref = float(flow_ref)

    if flow <= 0:
        return 0.0

    # floor: 유량이 기준(flow_ref) 이하이면 factor를 쓰지 않고 cost_ref로 고정.
    # 반응조 CAPEX에는 적용하지 않음.
    flow = max(flow, flow_ref)

    return cost_ref * (flow / flow_ref) ** n


@lru_cache(maxsize=1)
def load_biogas_scaler():
    if os.path.exists(SCALER_PATH):
        with open(SCALER_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)

        x_offset = data["x_offset"]
        x_factor = data["x_factor"]
        y_offset = data["y_offset"]
        y_factor = data["y_factor"]
        scaled_lb = data["scaled_lb"]
        scaled_ub = data["scaled_ub"]

        x_raw_min = data.get("x_raw_min")
        x_raw_max = data.get("x_raw_max")

        if x_raw_min is None or x_raw_max is None:
            x_raw_min = {}
            x_raw_max = {}
            for name in NN_INPUTS:
                offset = float(x_offset[name])
                factor = float(x_factor[name])
                lb = offset + float(scaled_lb[name]) * factor
                ub = offset + float(scaled_ub[name]) * factor
                x_raw_min[name] = min(lb, ub)
                x_raw_max[name] = max(lb, ub)

        return x_offset, x_factor, y_offset, y_factor, scaled_lb, scaled_ub, x_raw_min, x_raw_max

    if not os.path.exists(TRAINING_DATA_PATH):
        raise FileNotFoundError(
            "OMLT surrogate를 쓰려면 biogas_nn_relu_scaler.json 또는 "
            "ADM_5000_results.xlsx가 필요합니다."
        )

    df = pd.read_excel(TRAINING_DATA_PATH, usecols=NN_INPUTS + NN_OUTPUTS)
    dfin = df[NN_INPUTS]
    dfout = df[NN_OUTPUTS]

    x_offset = dfin.mean().to_dict()
    x_factor = dfin.std().to_dict()
    y_offset = dfout.mean().to_dict()
    y_factor = dfout.std().to_dict()

    dfin_scaled = (dfin - dfin.mean()).divide(dfin.std())
    scaled_lb = dfin_scaled.min()[NN_INPUTS].to_dict()
    scaled_ub = dfin_scaled.max()[NN_INPUTS].to_dict()

    x_raw_min = dfin.min()[NN_INPUTS].to_dict()
    x_raw_max = dfin.max()[NN_INPUTS].to_dict()

    data = {
        "x_offset": x_offset,
        "x_factor": x_factor,
        "y_offset": y_offset,
        "y_factor": y_factor,
        "scaled_lb": scaled_lb,
        "scaled_ub": scaled_ub,
        "x_raw_min": x_raw_min,
        "x_raw_max": x_raw_max,
    }

    with open(SCALER_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

    return x_offset, x_factor, y_offset, y_factor, scaled_lb, scaled_ub, x_raw_min, x_raw_max


@lru_cache(maxsize=1)
def load_biogas_output_bounds():
    if os.path.exists(SCALER_PATH):
        with open(SCALER_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)

        y_raw_min = data.get("y_raw_min")
        y_raw_max = data.get("y_raw_max")

        if y_raw_min is not None and y_raw_max is not None:
            return y_raw_min, y_raw_max

    if os.path.exists(TRAINING_DATA_PATH):
        df = pd.read_excel(TRAINING_DATA_PATH, usecols=NN_OUTPUTS)
        y_raw_min = df[NN_OUTPUTS].min().to_dict()
        y_raw_max = df[NN_OUTPUTS].max().to_dict()
        return y_raw_min, y_raw_max

    y_raw_min = {
        "final_methane_fraction_dry": 0.0,
        "final_q_gas(m3/d)": 0.0,
        "final_co2_fraction_dry": 0.0,
    }
    y_raw_max = {
        "final_methane_fraction_dry": 1.0,
        "final_q_gas(m3/d)": 1.0e6,
        "final_co2_fraction_dry": 1.0,
    }
    return y_raw_min, y_raw_max


def make_even_points(lb, ub, n=5):
    lb = float(lb)
    ub = float(ub)

    if ub <= lb:
        ub = lb + 1.0

    return [lb + i * (ub - lb) / (n - 1) for i in range(n)]


def cost_capacity_breakpoints(lb, ub, ref, n_segments=1):
    """비용-용량 곡선의 Piecewise 절점 리스트.
    ref가 [lb,ub] 내부면 절점 추가, ref 초과 구간을 n_segments 등분.
    n_segments=1이면 3점([lb,ref,ub]) 또는 2점([lb,ub]) 근사."""
    lb=float(lb); ub=float(ub); ref=float(ref)
    if ub<=lb: ub=lb+1.0e-6
    points=[lb]
    if lb<ref<ub:
        points.append(ref); start=ref
    else:
        start=lb
    if ub>start:
        for i in range(1,n_segments):
            points.append(start+i*(ub-start)/n_segments)
        points.append(ub)
    points=sorted(set(float(x) for x in points))
    points[0]=lb; points[-1]=ub
    return points


def add_piecewise_bilinear_product(
    model,
    x_var,
    y_var,
    z_var,
    x_pts,
    y_pts,
    name,
):
    """
    z = x*y를 삼각형 기반 piecewise-linear surface로 근사한다.

    기존 사각 셀의 네 꼭짓점 convex combination 방식은 동일한 x, y에서도
    z가 여러 값을 가질 수 있어 목적함수가 z를 비물리적으로 과대평가할 수 있었다.

    각 사각 셀을 두 삼각형으로 나눈 뒤 정확히 하나의 삼각형만 선택하고,
    선택된 삼각형의 세 꼭짓점으로 barycentric interpolation을 수행한다.
    따라서 주어진 x, y에 대한 z가 선택된 삼각형 안에서 하나로 결정된다.
    """
    triangles = []
    vertices = {}
    tri_id = 0

    for i in range(len(x_pts) - 1):
        for j in range(len(y_pts) - 1):
            x0 = float(x_pts[i])
            x1 = float(x_pts[i + 1])
            y0 = float(y_pts[j])
            y1 = float(y_pts[j + 1])

            # [Fix 2] 대각선 방향
            #  v48이 쓰던 두 삼각형은 둘 다 xy 의 McCormick '상한' 평면과 정확히
            #  일치한다. 그래서 z=x*y 가 구조적으로 과대평가됐다
            #  (v48 출력 192행 전부 +오차, CH4 대비 평균 +7.07%).
            #  반대 대각선을 쓰면 두 평면 모두 '하한' 평면이 되어 보수적(과소)이 된다.
            tri_a = [(x0, y0), (x1, y0), (x0, y1)]
            tri_b = [(x1, y0), (x0, y1), (x1, y1)]

            triangles.append(tri_id)
            vertices[tri_id] = tri_a
            tri_id += 1

            triangles.append(tri_id)
            vertices[tri_id] = tri_b
            tri_id += 1

    tri_set = pyo.Set(initialize=triangles)
    vertex_set = pyo.Set(
        initialize=[
            (t, v)
            for t in triangles
            for v in range(3)
        ],
        dimen=2,
    )

    setattr(model, f"{name}_TRI", tri_set)
    setattr(model, f"{name}_VERTEX", vertex_set)

    tri_bin = pyo.Var(tri_set, domain=pyo.Binary)
    lam = pyo.Var(vertex_set, domain=pyo.NonNegativeReals)

    setattr(model, f"{name}_tri_bin", tri_bin)
    setattr(model, f"{name}_lambda", lam)

    cons = pyo.ConstraintList()
    setattr(model, f"{name}_cons", cons)

    # 정확히 하나의 삼각형만 활성화
    cons.add(sum(tri_bin[t] for t in triangles) == 1)

    # 선택된 삼각형의 lambda 합은 1, 비선택 삼각형은 0
    for t in triangles:
        cons.add(
            sum(lam[t, v] for v in range(3))
            == tri_bin[t]
        )

    cons.add(
        x_var
        == sum(
            lam[t, v] * vertices[t][v][0]
            for t in triangles
            for v in range(3)
        )
    )

    cons.add(
        y_var
        == sum(
            lam[t, v] * vertices[t][v][1]
            for t in triangles
            for v in range(3)
        )
    )

    cons.add(
        z_var
        == sum(
            lam[t, v]
            * vertices[t][v][0]
            * vertices[t][v][1]
            for t in triangles
            for v in range(3)
        )
    )


def _build_network_from_keras_file(path, scaler, scaled_input_bounds):
    """keras 없이 .keras(zip) 파일에서 직접 OMLT NetworkDefinition을 만든다.

    keras.models.load_model + load_keras_sequential 과 동일한 망을 구성한다.
    (Dense + relu/linear 로만 구성된 Sequential 모델에 한함)
    """
    import io as _io
    import json as _json
    import zipfile as _zipfile

    import h5py
    from omlt.neuralnet import NetworkDefinition
    from omlt.neuralnet.layer import DenseLayer, InputLayer

    z = _zipfile.ZipFile(path)
    cfg = _json.loads(z.read("config.json"))
    all_layers = cfg["config"]["layers"]
    dense_cfgs = [l for l in all_layers if l["class_name"] == "Dense"]

    h5 = h5py.File(_io.BytesIO(z.read("model.weights.h5")), "r")
    keys = list(h5["layers"].keys())

    def _order(k):
        tail = k.rsplit("_", 1)[-1]
        return int(tail) if tail.isdigit() else 0

    keys = sorted(keys, key=_order)
    if len(keys) != len(dense_cfgs):
        raise ValueError("keras 파일의 Dense 층 수와 가중치 수가 일치하지 않습니다.")

    n_in = int(all_layers[0]["config"]["batch_shape"][1])

    net = NetworkDefinition(
        scaling_object=scaler,
        scaled_input_bounds=scaled_input_bounds,
    )
    prev = InputLayer([n_in])
    net.add_layer(prev)

    for lcfg, key in zip(dense_cfgs, keys):
        Wm = np.array(h5[f"layers/{key}/vars/0"], dtype=float)
        bv = np.array(h5[f"layers/{key}/vars/1"], dtype=float)
        act = lcfg["config"].get("activation", "linear")
        layer = DenseLayer(
            prev.output_size,
            [Wm.shape[1]],
            weights=Wm,
            biases=bv,
            activation=("relu" if act == "relu" else "linear"),
        )
        net.add_layer(layer)
        net.add_edge(prev, layer)
        prev = layer

    return net


@lru_cache(maxsize=1)
def load_biogas_network_definition():
    if not os.path.exists(NN_MODEL_PATH):
        raise FileNotFoundError(f"Keras 모델 파일을 찾을 수 없습니다: {NN_MODEL_PATH}")

    x_offset, x_factor, y_offset, y_factor, scaled_lb, scaled_ub, x_raw_min, x_raw_max = load_biogas_scaler()

    scaled_input_bounds = {}
    for i, name in enumerate(NN_INPUTS):
        offset = float(x_offset[name])
        factor = float(x_factor[name])

        raw_lb = float(x_raw_min[name]) - RAW_RANGE_TOL
        raw_ub = float(x_raw_max[name]) + RAW_RANGE_TOL

        scaled_a = (raw_lb - offset) / factor
        scaled_b = (raw_ub - offset) / factor

        # 학습데이터의 실제 raw 최소·최대값을 scaled 영역으로 변환해 사용
        # 임의의 (-20, 20)은 지나치게 넓어 ReLU Big-M를 약화시키고
        # 학습범위를 크게 벗어난 외삽을 허용하므로 사용하지 않는다.
        base_lb = min(scaled_a, scaled_b)
        base_ub = max(scaled_a, scaled_b)
        base_span = max(base_ub - base_lb, 1.0e-6)

        expansion = max(
            NN_BOUND_MIN_EXPANSION_STD,
            NN_BOUND_SPAN_EXPANSION * base_span,
        )

        scaled_input_bounds[i] = (
            max(-NN_BOUND_ABS_LIMIT, base_lb - expansion),
            min(NN_BOUND_ABS_LIMIT, base_ub + expansion),
        )

    scaler = OffsetScaling(
        offset_inputs={i: x_offset[NN_INPUTS[i]] for i in range(len(NN_INPUTS))},
        factor_inputs={i: x_factor[NN_INPUTS[i]] for i in range(len(NN_INPUTS))},
        offset_outputs={i: y_offset[NN_OUTPUTS[i]] for i in range(len(NN_OUTPUTS))},
        factor_outputs={i: y_factor[NN_OUTPUTS[i]] for i in range(len(NN_OUTPUTS))},
    )

    if keras is not None and load_keras_sequential is not None:
        nn_biogas = keras.models.load_model(NN_MODEL_PATH, compile=False)
        net = load_keras_sequential(
            nn_biogas,
            scaling_object=scaler,
            scaled_input_bounds=scaled_input_bounds,
        )
    else:
        net = _build_network_from_keras_file(
            NN_MODEL_PATH, scaler, scaled_input_bounds
        )

    return net


def add_biogas_omlt_surrogate(model):
    net = load_biogas_network_definition()
    x_offset, x_factor, y_offset, y_factor, scaled_lb, scaled_ub, x_raw_min, x_raw_max = load_biogas_scaler()

    model.biogas = OmltBlock()
    model.biogas.build_formulation(ReluBigMFormulation(net))

    model.connect_nn_inputs = pyo.ConstraintList()
    model.nn_input_range_cons = pyo.ConstraintList()

    nn_input_exprs = [
        model.Q_ad_in,
        model.COD_total,
        model.ADM_total["S_su"],
        model.ADM_total["S_aa"],
        model.ADM_total["S_fa"],
        model.ADM_total["S_I"],
        model.ADM_total["X_ch"],
        model.ADM_total["X_pr"],
        model.ADM_total["X_li"],
        model.ADM_total["X_I"],
        model.OLR,
    ]

    for i, expr in enumerate(nn_input_exprs):
        name = NN_INPUTS[i]
        # 제한적 외삽 허용: 확장된 OMLT scaled bounds를 사용함.
        model.connect_nn_inputs.add(model.biogas.inputs[i] == expr)

    y_raw_min, y_raw_max = load_biogas_output_bounds()

    ch4_lb = 0.0
    ch4_ub = 1.0
    # 저유량 지역도 허용하기 위해 학습 출력 최소값을 하한으로 강제하지 않음
    qgas_lb = 0.0
    qgas_ub = max(
        qgas_lb + 1.0,
        float(y_raw_max["final_q_gas(m3/d)"]) + RAW_RANGE_TOL,
    )
    co2_lb = 0.0
    co2_ub = 1.0

    model.pred_ch4_frac_dry_var = pyo.Var(bounds=(ch4_lb, ch4_ub))
    model.pred_ch4_frac_dry_eq = pyo.Constraint(
        expr=model.pred_ch4_frac_dry_var == model.biogas.outputs[0]
    )

    model.pred_q_gas_var = pyo.Var(bounds=(qgas_lb, qgas_ub))
    model.pred_q_gas_eq = pyo.Constraint(
        expr=model.pred_q_gas_var == model.biogas.outputs[1]
    )

    # NN의 CO2 출력은 검산용으로 보존
    model.pred_co2_frac_nn_var = pyo.Var(bounds=(co2_lb, co2_ub))
    model.pred_co2_frac_nn_eq = pyo.Constraint(
        expr=model.pred_co2_frac_nn_var == model.biogas.outputs[2]
    )

    # 건식 바이오가스는 CH4 + CO2 = 1로 강제
    model.pred_co2_frac_var = pyo.Var(bounds=(co2_lb, co2_ub))
    model.pred_co2_frac_eq = pyo.Constraint(
        expr=model.pred_co2_frac_var
        == 1.0 - model.pred_ch4_frac_dry_var
    )


# =========================================================
# 1. 모델 생성 함수
# =========================================================
# =======================================================================
# §5  MODEL BUILD — 공정 순서 스테이지
# =======================================================================

# -----------------------------------------------------------------------
# S1  기초 · 단위환산 · Big-M
# -----------------------------------------------------------------------
def _stage_setup(model, ctx):
    # 입력 원자료 단위: t/d
    ctx.W = float(ctx.W)
    ctx.F = float(ctx.F)
    ctx.S = float(ctx.S)

    # 밀도 1,000 kg/m3 가정에 따른 부피유량 환산
    # Q[m3/d] = M[t/d] * 1,000[kg/t] / rho[kg/m3]
    ctx.W_m3_d = ctx.W * TON_TO_KG / RAW_FEED_DENSITY_KG_M3
    ctx.F_m3_d = ctx.F * TON_TO_KG / RAW_FEED_DENSITY_KG_M3
    ctx.S_m3_d = ctx.S * TON_TO_KG / RAW_FEED_DENSITY_KG_M3

    # [v76] 표준사업비 부대비용 [USD]. 총 원료량만의 함수라 지역별 상수다.
    ctx.TOTAL_TON_D = ctx.W + ctx.F + ctx.S
    ctx.INDIRECT_USD = indirect_project_cost_usd(ctx.TOTAL_TON_D)


    # 유량 선택 제약용 Big-M
    # 지역별 총 원료유량의 2배를 사용해 기존 10,000,000보다 안정화
    ctx.M_FLOW = max(1.0, 2.0 * (ctx.W_m3_d + ctx.F_m3_d + ctx.S_m3_d))


    ctx.Solid_ani = ctx.W_m3_d * TS_ani
    ctx.Water_ani = ctx.W_m3_d * (1 - TS_ani)

    ctx.Solid_fw = ctx.F_m3_d * TS_fw
    ctx.Water_fw = ctx.F_m3_d * (1 - TS_fw)

    ctx.Solid_sl = ctx.S_m3_d * TS_sl
    ctx.Water_sl = ctx.S_m3_d * (1 - TS_sl)
    
    # ===================== 반입 설비 ===================== 


# -----------------------------------------------------------------------
# S2  반입 설비 (feeding hopper · conveyor)
# -----------------------------------------------------------------------
def _stage_receiving(model, ctx):
    W_per_line = ctx.W_m3_d / 2.0

    AM_stroage_cost = 2.0 * scale_cost(16783, W_per_line, 50, 0.6)  # AM storage 1 + AM storage 2

    SL_stroage_cost = scale_cost(16783, ctx.S, 50, 0.6)
    # 반입설비용 FW feeding hopper (전처리용 hopper와 별도)
    FW_receiving_feeding_hopper = scale_cost(2000, ctx.F, 2, SF_FEEDING_HOPPER)
    
    model.Biomass_pulping_cost = pyo.Var(domain=pyo.NonNegativeReals)

    AM_air_curtain_1 = 79.0 if ctx.W > 0 else 0.0
    AM_air_curtain_2 = 79.0 if ctx.W > 0 else 0.0
    SL_air_curtain = 79.0 if ctx.S > 0 else 0.0
    FW_air_curtain = 79.0 if ctx.F > 0 else 0.0

    AM_shutter_1 = 1019.0 if ctx.W > 0 else 0.0
    AM_shutter_2 = 1019.0 if ctx.W > 0 else 0.0
    SL_shutter = 1019.0 if ctx.S > 0 else 0.0
    FW_shutter = 1019.0 if ctx.F > 0 else 0.0

    # Conveyor
    AM_conveyor_1 = scale_cost(4500, W_per_line, 360, 0.6)
    AM_conveyor_2 = scale_cost(4500, W_per_line, 360, 0.6)
    SL_conveyor_1 = scale_cost(4500, ctx.S, 360, 0.6)
    FW_conveyor_1 = scale_cost(4500, ctx.F, 360, 0.6)
    FW_conveyor_2 = scale_cost(4500, ctx.F, 360, 0.6)

    model.receiving_cost = pyo.Var(domain=pyo.NonNegativeReals)
    model.eqreceiving = pyo.Constraint(
        expr=model.receiving_cost ==
        AM_stroage_cost + SL_stroage_cost +
        FW_receiving_feeding_hopper +
        AM_air_curtain_1 + AM_air_curtain_2 +
        SL_air_curtain + FW_air_curtain +
        AM_conveyor_1 + AM_conveyor_2 + SL_conveyor_1 +
        FW_conveyor_1 + FW_conveyor_2 + AM_shutter_1 + AM_shutter_2
        + SL_shutter + FW_shutter
    )

    # ================== 전처리 설비 ==================


# -----------------------------------------------------------------------
# S3  전처리 — 가축분뇨 · 음식물 · 하수슬러지 · BM 중간저장조
# -----------------------------------------------------------------------
def _stage_pretreatment(model, ctx):
    AM_screen_total_cost = scale_cost(
        3000, ctx.W, 72, 0.6
    )
    AM_cyclone_total_cost = scale_cost(
        40000, ctx.W, 14400, 0.6
    )
    AM_drum_filter_total_cost = scale_cost(
        3000, ctx.W, 72, 0.6
    )

    # option 1: AM screen + AM cyclone
    # option 2: AM screen + AM cyclone + AM drum filter
    option_ani_pr_cost = [
        AM_screen_total_cost + AM_cyclone_total_cost,
        AM_screen_total_cost + AM_cyclone_total_cost + AM_drum_filter_total_cost,
    ]

    # 가축분뇨 전처리 비용 선택용 Big-M
    M_ANI_COST = max(1.0, 2.0 * max(option_ani_pr_cost))


    # =========================================================
    # 전처리 후 중간 저장조(BM storage) 설비
    # 전처리 완료 혼합유량 Q_total을 2계열로 균등 분할하여 처리
    # BM feeding hopper 2대, storage tank 2기,
    # vertical mixer 8대(각 저장조 4대), conveyor 2대
    # =========================================================
    BM_STORAGE_COST_DATA = {
        "feeding_hopper_1": (2000.0, 2.0, SF_FEEDING_HOPPER),
        "feeding_hopper_2": (2000.0, 2.0, SF_FEEDING_HOPPER),
        "storage_tank_1": (16783.0, 50.0, 0.6),
        "storage_tank_2": (16783.0, 50.0, 0.6),
        "conveyor_1": (4500.0, 360.0, 0.6),
        "conveyor_2": (4500.0, 360.0, 0.6),
    }

    # 음식물 전처리 공정 흐름
    # 1계열: screen 1 + sorting machine 1
    # 2계열: screen 2 + sorting machine 2
    # 두 계열 합류 -> shredder 1 -> screen 3 -> shredder 2
    F_per_line = ctx.F_m3_d / 2.0


    # screen 1·2는 각 병렬 계열의 원유량 F/2를 기준으로 계산
    FW_SCREEN_LINE_COST = {
        "drum": scale_cost(17000, F_per_line, 48, 0.6),
        "screw": scale_cost(5500, F_per_line, 480, 0.6),
        "step": scale_cost(1800, F_per_line, 96, 0.6),
    }



    FW_SCREENING_SHREDDER_SF = SF_SCREENING_SHREDDER


    FW_TWO_SHAFT_SHREDDER_SF = SF_TWO_SHAFT_SHREDDER


    FW_SORTING_COST_DATA = {
        1: (
            FW_SORTING_MACHINE_1_MIN_COST,
            FW_SORTING_MACHINE_1_MIN_VOL,
            FW_SORTING_MACHINE_1_SF,
        ),
        2: (
            FW_SORTING_MACHINE_2_MIN_COST,
            FW_SORTING_MACHINE_2_MIN_VOL,
            FW_SORTING_MACHINE_2_SF,
        ),
    }

    FW_SHREDDER_COST_DATA = {
        "screening": (
            FW_SCREENING_SHREDDER_MIN_COST,
            FW_SCREENING_SHREDDER_MIN_VOL,
            FW_SCREENING_SHREDDER_SF,
        ),
        "rotary": (
            FW_ROTARY_SHREDDER_MIN_COST,
            FW_ROTARY_SHREDDER_MIN_VOL,
            FW_ROTARY_SHREDDER_SF,
        ),
        "two_shaft": (
            FW_TWO_SHAFT_SHREDDER_MIN_COST,
            FW_TWO_SHAFT_SHREDDER_MIN_VOL,
            FW_TWO_SHAFT_SHREDDER_SF,
        ),
        "hammer": (
            FW_HAMMER_CRUSHER_MIN_COST,
            FW_HAMMER_CRUSHER_MIN_VOL,
            FW_HAMMER_CRUSHER_SF,
        ),
    }



    option_sl_pr_cost = [
        scale_cost(4000, ctx.S, 6, 0.6),        # gravity thickening
        scale_cost(12000, ctx.S, 120, 0.6),     # rotary drum thickening
    ]

    # ------------------- Variables -------------------
    model.W_ani_pr1 = pyo.Var(domain=pyo.NonNegativeReals)
    model.C_ani_pr1_cost = pyo.Var(domain=pyo.NonNegativeReals)
    model.Q_ani_pr1 = pyo.Var(domain=pyo.NonNegativeReals)
    model.Q_ani = pyo.Var(domain=pyo.NonNegativeReals)

    # 음식물 전처리 공정별 수분 및 총 유량
    model.W_fw_l1_after_screen = pyo.Var(domain=pyo.NonNegativeReals)
    model.W_fw_l2_after_screen = pyo.Var(domain=pyo.NonNegativeReals)
    model.W_fw_after_merge = pyo.Var(domain=pyo.NonNegativeReals)
    model.W_fw_after_screen3 = pyo.Var(domain=pyo.NonNegativeReals)

    model.Q_fw_l1_after_screen = pyo.Var(domain=pyo.NonNegativeReals)
    model.Q_fw_l2_after_screen = pyo.Var(domain=pyo.NonNegativeReals)
    model.Q_fw_l1_after_sorting = pyo.Var(domain=pyo.NonNegativeReals)
    model.Q_fw_l2_after_sorting = pyo.Var(domain=pyo.NonNegativeReals)
    model.Q_fw_after_merge = pyo.Var(domain=pyo.NonNegativeReals)
    model.Q_fw_after_shredder1 = pyo.Var(domain=pyo.NonNegativeReals)
    model.Q_fw_after_shredder2 = pyo.Var(domain=pyo.NonNegativeReals)
    model.Q_fw_after_screen3 = pyo.Var(domain=pyo.NonNegativeReals)

    # 각 스크린에서 분리되어 waste 계통으로 빠지는 유량
    model.Q_fw_waste_line1 = pyo.Var(domain=pyo.NonNegativeReals)
    model.Q_fw_waste_line2 = pyo.Var(domain=pyo.NonNegativeReals)
    model.Q_fw_waste_screen3 = pyo.Var(domain=pyo.NonNegativeReals)
    model.Q_fw_waste_total = pyo.Var(domain=pyo.NonNegativeReals)

    model.fw_pr_cost = pyo.Var(domain=pyo.NonNegativeReals)
    model.fw_waste_handling_cost = pyo.Var(domain=pyo.NonNegativeReals)
    model.Q_fw_pr = pyo.Var(domain=pyo.NonNegativeReals)
    model.Q_fw = pyo.Var(domain=pyo.NonNegativeReals)
    model.W_sl_pr1 = pyo.Var(domain=pyo.NonNegativeReals)
    model.Q_sl_pr = pyo.Var(domain=pyo.NonNegativeReals)
    model.Q_sl = pyo.Var(domain=pyo.NonNegativeReals)

    model.C_sl_pr = pyo.Var(domain=pyo.NonNegativeReals)

    model.PRETREATMENT_AUX = pyo.Set(
        initialize=list(PRETREATMENT_AUX_COST_DATA.keys())
    )
    model.pretreatment_aux_cost_if = pyo.Var(
        model.PRETREATMENT_AUX,
        domain=pyo.NonNegativeReals,
    )
    model.pretreatment_aux_cost = pyo.Var(domain=pyo.NonNegativeReals)

    # 전처리 후 중간 저장조 유량 및 비용
    model.BM_STORAGE_EQUIP = pyo.Set(
        initialize=list(BM_STORAGE_COST_DATA.keys())
    )
    model.bm_storage_cost_if = pyo.Var(
        model.BM_STORAGE_EQUIP,
        domain=pyo.NonNegativeReals,
    )
    model.bm_storage_cost = pyo.Var(domain=pyo.NonNegativeReals)

    model.Q_bm_storage_in = pyo.Var(domain=pyo.NonNegativeReals)
    model.Q_bm_line1 = pyo.Var(domain=pyo.NonNegativeReals)
    model.Q_bm_line2 = pyo.Var(domain=pyo.NonNegativeReals)
    model.Q_bm_storage_out = pyo.Var(domain=pyo.NonNegativeReals)
    model.Q_ad_in = pyo.Var(domain=pyo.NonNegativeReals)

    model.Cost = pyo.Var(domain=pyo.NonNegativeReals)

    # ------------------- Animal manure -------------------
    model.y1 = pyo.Var(domain=pyo.Binary)

    model.eq2_1_1 = pyo.Constraint(expr=model.W_ani_pr1 <= ctx.Water_ani * (1 - water_removal_ani[1]) + ctx.M_FLOW * (1 - model.y1))
    model.eq2_1_2 = pyo.Constraint(expr=model.W_ani_pr1 >= ctx.Water_ani * (1 - water_removal_ani[1]) - ctx.M_FLOW * (1 - model.y1))

    model.eq2_1_3 = pyo.Constraint(expr=model.C_ani_pr1_cost <= option_ani_pr_cost[1] + M_ANI_COST * (1 - model.y1))
    model.eq2_1_4 = pyo.Constraint(expr=model.C_ani_pr1_cost >= option_ani_pr_cost[1] - M_ANI_COST * (1 - model.y1))

    model.eq2_2_1 = pyo.Constraint(expr=model.W_ani_pr1 <= ctx.Water_ani * (1 - water_removal_ani[0]) + ctx.M_FLOW * model.y1)
    model.eq2_2_2 = pyo.Constraint(expr=model.W_ani_pr1 >= ctx.Water_ani * (1 - water_removal_ani[0]) - ctx.M_FLOW * model.y1)

    model.eq2_2_3 = pyo.Constraint(expr=model.C_ani_pr1_cost <= option_ani_pr_cost[0] + M_ANI_COST * model.y1)
    model.eq2_2_4 = pyo.Constraint(expr=model.C_ani_pr1_cost >= option_ani_pr_cost[0] - M_ANI_COST * model.y1)

    model.eq2_5 = pyo.Constraint(expr=model.Q_ani_pr1 == ctx.Solid_ani + model.W_ani_pr1)
    model.eq2_6 = pyo.Constraint(expr=model.Q_ani == model.Q_ani_pr1)

    # ------------------- Food waste pretreatment -------------------
    # 1계열: screen 1 + sorting machine 1
    # 2계열: screen 2 + sorting machine 2
    # 합류 -> shredder 1 -> screen 3 -> shredder 2

    model.FW_SCREEN = pyo.Set(initialize=["drum", "screw", "step"])
    model.FW_SHREDDER = pyo.Set(
        initialize=["screening", "rotary", "two_shaft", "hammer"]
    )
    model.FW_LINE = pyo.Set(initialize=[1, 2])
    model.FW_SHREDDER_STAGE = pyo.Set(initialize=[1, 2])

    model.fw_line_screen_sel = pyo.Var(
        model.FW_LINE, model.FW_SCREEN, domain=pyo.Binary
    )
    model.fw_shredder_sel = pyo.Var(
        model.FW_SHREDDER_STAGE, model.FW_SHREDDER, domain=pyo.Binary
    )
    model.fw_screen3_sel = pyo.Var(model.FW_SCREEN, domain=pyo.Binary)

    model.fw_line_screen_one = pyo.ConstraintList()
    for line in model.FW_LINE:
        model.fw_line_screen_one.add(
            sum(model.fw_line_screen_sel[line, s] for s in model.FW_SCREEN) == 1
        )

    model.fw_shredder_one = pyo.ConstraintList()
    for stage in model.FW_SHREDDER_STAGE:
        model.fw_shredder_one.add(
            sum(model.fw_shredder_sel[stage, h] for h in model.FW_SHREDDER) == 1
        )

    model.fw_screen3_one = pyo.Constraint(
        expr=sum(model.fw_screen3_sel[s] for s in model.FW_SCREEN) == 1
    )

    Water_fw_per_line = ctx.Water_fw / 2.0
    Solid_fw_per_line = ctx.Solid_fw / 2.0

    # screen 1·2에서 선택된 수분 제거율 적용
    model.fw_l1_screen_cons = pyo.ConstraintList()
    model.fw_l2_screen_cons = pyo.ConstraintList()
    for s in model.FW_SCREEN:
        eff = FW_SCREEN_WATER_REMOVAL[s]
        target = Water_fw_per_line * (1 - eff)

        model.fw_l1_screen_cons.add(
            model.W_fw_l1_after_screen
            <= target + ctx.M_FLOW * (1 - model.fw_line_screen_sel[1, s])
        )
        model.fw_l1_screen_cons.add(
            model.W_fw_l1_after_screen
            >= target - ctx.M_FLOW * (1 - model.fw_line_screen_sel[1, s])
        )

        model.fw_l2_screen_cons.add(
            model.W_fw_l2_after_screen
            <= target + ctx.M_FLOW * (1 - model.fw_line_screen_sel[2, s])
        )
        model.fw_l2_screen_cons.add(
            model.W_fw_l2_after_screen
            >= target - ctx.M_FLOW * (1 - model.fw_line_screen_sel[2, s])
        )

    # screen 후 총 유량 = 계열별 고형물 + 남은 수분
    model.fw_l1_after_screen_eq = pyo.Constraint(
        expr=model.Q_fw_l1_after_screen ==
        Solid_fw_per_line + model.W_fw_l1_after_screen
    )
    model.fw_l2_after_screen_eq = pyo.Constraint(
        expr=model.Q_fw_l2_after_screen ==
        Solid_fw_per_line + model.W_fw_l2_after_screen
    )

    # sorting machine은 별도 제거율이 없으므로 유량 유지
    model.fw_l1_after_sorting_eq = pyo.Constraint(
        expr=model.Q_fw_l1_after_sorting == model.Q_fw_l1_after_screen
    )
    model.fw_l2_after_sorting_eq = pyo.Constraint(
        expr=model.Q_fw_l2_after_sorting == model.Q_fw_l2_after_screen
    )

    # 두 계열 합류
    model.fw_merge_eq = pyo.Constraint(
        expr=model.Q_fw_after_merge ==
        model.Q_fw_l1_after_sorting + model.Q_fw_l2_after_sorting
    )
    model.fw_water_merge_eq = pyo.Constraint(
        expr=model.W_fw_after_merge ==
        model.W_fw_l1_after_screen + model.W_fw_l2_after_screen
    )

    # 합류 후 shredder 1은 입도만 감소시키므로 유량 유지
    model.fw_after_shredder1_eq = pyo.Constraint(
        expr=model.Q_fw_after_shredder1 == model.Q_fw_after_merge
    )

    # screen 3에서 shredder 1 출구의 남은 수분에 추가 제거율 적용
    model.fw_screen3_cons = pyo.ConstraintList()
    for s in model.FW_SCREEN:
        eff = FW_SCREEN_WATER_REMOVAL[s]
        model.fw_screen3_cons.add(
            model.W_fw_after_screen3
            <= model.W_fw_after_merge * (1 - eff)
            + ctx.M_FLOW * (1 - model.fw_screen3_sel[s])
        )
        model.fw_screen3_cons.add(
            model.W_fw_after_screen3
            >= model.W_fw_after_merge * (1 - eff)
            - ctx.M_FLOW * (1 - model.fw_screen3_sel[s])
        )

    model.fw_after_screen3_eq = pyo.Constraint(
        expr=model.Q_fw_after_screen3 == ctx.Solid_fw + model.W_fw_after_screen3
    )

    # 각 스크린에서 제거된 유량은 waste 계통으로 이동
    model.fw_waste_line1_eq = pyo.Constraint(
        expr=model.Q_fw_waste_line1 == F_per_line - model.Q_fw_l1_after_screen
    )
    model.fw_waste_line2_eq = pyo.Constraint(
        expr=model.Q_fw_waste_line2 == F_per_line - model.Q_fw_l2_after_screen
    )
    model.fw_waste_screen3_eq = pyo.Constraint(
        expr=model.Q_fw_waste_screen3 ==
        model.Q_fw_after_shredder1 - model.Q_fw_after_screen3
    )
    model.fw_waste_total_eq = pyo.Constraint(
        expr=model.Q_fw_waste_total ==
        model.Q_fw_waste_line1 +
        model.Q_fw_waste_line2 +
        model.Q_fw_waste_screen3
    )

    # screen 3 후 shredder 2는 입도만 감소시키므로 유량 유지
    model.fw_after_shredder2_eq = pyo.Constraint(
        expr=model.Q_fw_after_shredder2 == model.Q_fw_after_screen3
    )

    # 유량 변수 bounds: 실제 가능한 최소·최대 유량을 사용
    # 이렇게 해야 최소비용 floor가 0~첫 절점 사이에서 선형으로 깎이지 않음
    _fw_eff_min = min(FW_SCREEN_WATER_REMOVAL.values())
    _fw_eff_max = max(FW_SCREEN_WATER_REMOVAL.values())

    _fw_line_lb = (
        Solid_fw_per_line + Water_fw_per_line * (1 - _fw_eff_max)
    )
    _fw_line_ub = (
        Solid_fw_per_line + Water_fw_per_line * (1 - _fw_eff_min)
    )

    _fw_merge_lb = ctx.Solid_fw + ctx.Water_fw * (1 - _fw_eff_max)
    _fw_merge_ub = ctx.Solid_fw + ctx.Water_fw * (1 - _fw_eff_min)

    _fw_screen3_out_lb = ctx.Solid_fw + ctx.Water_fw * (1 - _fw_eff_max) ** 2
    _fw_screen3_out_ub = ctx.Solid_fw + ctx.Water_fw * (1 - _fw_eff_min) ** 2

    _fw_waste_lb = ctx.F - _fw_screen3_out_ub
    _fw_waste_ub = ctx.F - _fw_screen3_out_lb

    for var in [
        model.Q_fw_l1_after_screen,
        model.Q_fw_l2_after_screen,
        model.Q_fw_l1_after_sorting,
        model.Q_fw_l2_after_sorting,
    ]:
        var.setlb(_fw_line_lb)
        var.setub(_fw_line_ub)

    for var in [
        model.Q_fw_after_merge,
        model.Q_fw_after_shredder1,
    ]:
        var.setlb(_fw_merge_lb)
        var.setub(_fw_merge_ub)

    for var in [
        model.Q_fw_after_screen3,
        model.Q_fw_after_shredder2,
    ]:
        var.setlb(_fw_screen3_out_lb)
        var.setub(_fw_screen3_out_ub)

    model.Q_fw_waste_line1.setlb(0.0)
    model.Q_fw_waste_line1.setub(F_per_line)
    model.Q_fw_waste_line2.setlb(0.0)
    model.Q_fw_waste_line2.setub(F_per_line)
    model.Q_fw_waste_screen3.setlb(0.0)
    model.Q_fw_waste_screen3.setub(ctx.F)
    model.Q_fw_waste_total.setlb(max(0.0, _fw_waste_lb))
    model.Q_fw_waste_total.setub(max(_fw_waste_ub, 1.0e-6))

    # ---------------------------------------------------------
    # 실제 전단 유량을 이용한 cost-capacity 비용식
    # ---------------------------------------------------------
    model.fw_sorting_cost = pyo.Var(model.FW_LINE, domain=pyo.NonNegativeReals)
    model.fw_shredder_cost_if = pyo.Var(
        model.FW_SHREDDER_STAGE,
        model.FW_SHREDDER,
        domain=pyo.NonNegativeReals,
    )
    model.fw_shredder_cost_selected = pyo.Var(
        model.FW_SHREDDER_STAGE,
        model.FW_SHREDDER,
        domain=pyo.NonNegativeReals,
    )
    model.fw_screen3_cost_if = pyo.Var(
        model.FW_SCREEN, domain=pyo.NonNegativeReals
    )
    model.fw_screen3_cost_selected = pyo.Var(
        model.FW_SCREEN, domain=pyo.NonNegativeReals
    )

    model.FW_WASTE_EQUIP = pyo.Set(
        initialize=["discharge_hopper", "conveyor_1", "conveyor_2"]
    )
    model.fw_waste_cost_if = pyo.Var(
        model.FW_WASTE_EQUIP, domain=pyo.NonNegativeReals
    )

    if ctx.F <= 0:
        # 음식물 유입이 없으면 모든 단계 유량과 비용을 0으로 고정
        model.fw_zero_flow_cons = pyo.ConstraintList()
        for var in [
            model.W_fw_l1_after_screen,
            model.W_fw_l2_after_screen,
            model.W_fw_after_merge,
            model.W_fw_after_screen3,
            model.Q_fw_l1_after_screen,
            model.Q_fw_l2_after_screen,
            model.Q_fw_l1_after_sorting,
            model.Q_fw_l2_after_sorting,
            model.Q_fw_after_merge,
            model.Q_fw_after_shredder1,
            model.Q_fw_after_shredder2,
            model.Q_fw_after_screen3,
            model.Q_fw_waste_line1,
            model.Q_fw_waste_line2,
            model.Q_fw_waste_screen3,
            model.Q_fw_waste_total,
        ]:
            model.fw_zero_flow_cons.add(var == 0.0)

        for line in model.FW_LINE:
            model.fw_zero_flow_cons.add(model.fw_sorting_cost[line] == 0.0)

        for stage in model.FW_SHREDDER_STAGE:
            for h in model.FW_SHREDDER:
                model.fw_zero_flow_cons.add(
                    model.fw_shredder_cost_if[stage, h] == 0.0
                )
                model.fw_zero_flow_cons.add(
                    model.fw_shredder_cost_selected[stage, h] == 0.0
                )

        for s in model.FW_SCREEN:
            model.fw_zero_flow_cons.add(model.fw_screen3_cost_if[s] == 0.0)
            model.fw_zero_flow_cons.add(model.fw_screen3_cost_selected[s] == 0.0)

        for e in model.FW_WASTE_EQUIP:
            model.fw_zero_flow_cons.add(model.fw_waste_cost_if[e] == 0.0)

        model.fw_zero_flow_cons.add(model.fw_waste_handling_cost == 0.0)

    else:
        # sorting machine 1·2: 각 계열 screen 출구 유량 기준
        # 비용식:
        #   Q = 0            -> 0
        #   0 < Q <= min_vol -> min_cost
        #   Q > min_vol      -> min_cost * (Q / min_vol) ** SF
        #
        # 가능한 유량범위가 min_vol 이하이면 최소비용으로 고정하고,
        # min_vol을 가로지르면 [lb, min_vol, ub]만 사용한다.
        # 이렇게 해야 평평한 구간을 여러 Piecewise segment로 쪼개는 경고가 사라진다.
        _sorting_input = {
            1: model.Q_fw_l1_after_screen,
            2: model.Q_fw_l2_after_screen,
        }

        model.fw_sorting_constant_cost_cons = pyo.ConstraintList()

        for line in model.FW_LINE:
            cost_ref, flow_ref, exponent = FW_SORTING_COST_DATA[line]

            if _fw_line_ub <= flow_ref:
                # 가능한 전 구간이 기준용량 이하
                model.fw_sorting_constant_cost_cons.add(
                    model.fw_sorting_cost[line] == cost_ref
                )

            else:
                # 기준용량을 가로지르는 경우에는 기준용량을 절점으로 포함
                if _fw_line_lb < flow_ref < _fw_line_ub:
                    pts = [
                        float(_fw_line_lb),
                        float(flow_ref),
                        float(_fw_line_ub),
                    ]
                else:
                    # 가능한 전 구간이 기준용량보다 큼
                    pts = [
                        float(_fw_line_lb),
                        float(_fw_line_ub),
                    ]

                add_cost_capacity_pw(
                    model,
                    f"pw_fw_sorting_cost_{line}",
                    model.fw_sorting_cost[line],
                    _sorting_input[line],
                    pts,
                    cost_ref, flow_ref, exponent,
                )

        # 파쇄기 단계별 실제 유입유량
        _shredder_input = {
            1: model.Q_fw_after_merge,
            2: model.Q_fw_after_screen3,
        }

        # 가능한 유량 전 구간이 기준용량 이하이면 최소비용으로 고정한다.
        # 기준용량을 넘는 경우에는 [lb, ub] 또는 [lb, ref, ub]만 사용하여
        # 동일 기울기 경고와 불필요한 binary 변수를 줄인다.
        _max_fw_shredder_cost = 0.0
        model.fw_shredder_constant_cost_cons = pyo.ConstraintList()

        for stage in model.FW_SHREDDER_STAGE:
            for h in model.FW_SHREDDER:
                cost_ref, flow_ref, exponent = FW_SHREDDER_COST_DATA[h]

                if stage == 1:
                    stage_lb, stage_ub = _fw_merge_lb, _fw_merge_ub
                else:
                    stage_lb, stage_ub = _fw_screen3_out_lb, _fw_screen3_out_ub

                _max_fw_shredder_cost = max(
                    _max_fw_shredder_cost,
                    cost_ref * (max(stage_ub, flow_ref) / flow_ref) ** exponent,
                )

                if stage_ub <= flow_ref:
                    model.fw_shredder_constant_cost_cons.add(
                        model.fw_shredder_cost_if[stage, h] == cost_ref
                    )
                else:
                    if stage_lb < flow_ref < stage_ub:
                        pts = [float(stage_lb), float(flow_ref), float(stage_ub)]
                    else:
                        pts = [float(stage_lb), float(stage_ub)]

                    add_cost_capacity_pw(
                        model,
                        f"pw_fw_shredder_cost_{stage}_{h}",
                        model.fw_shredder_cost_if[stage, h],
                        _shredder_input[stage],
                        pts,
                        cost_ref, flow_ref, exponent,
                        warning_tol=0.0,
                    )

        FW_SHREDDER_COST_M = max(1.0, _max_fw_shredder_cost * 10.0)
        model.fw_shredder_cost_select_cons = pyo.ConstraintList()
        for stage in model.FW_SHREDDER_STAGE:
            for h in model.FW_SHREDDER:
                add_bigm_select(
                    model.fw_shredder_cost_select_cons,
                    model.fw_shredder_cost_selected[stage, h],
                    model.fw_shredder_cost_if[stage, h],
                    model.fw_shredder_sel[stage, h],
                    FW_SHREDDER_COST_M,
                )

        # screen 3: 2차 파쇄 출구 유량 기준
        # 가능한 유량 전 구간이 기준용량 이하이면 Piecewise를 만들지 않고
        # 최소비용 cost_ref로 고정하여 동일 기울기 경고를 방지
        _max_fw_screen3_cost = 0.0
        model.fw_screen3_constant_cost_cons = pyo.ConstraintList()

        for s in model.FW_SCREEN:
            cost_ref, flow_ref, exponent = FW_SCREEN3_COST_DATA[s]
            _max_fw_screen3_cost = max(
                _max_fw_screen3_cost,
                cost_ref * (max(_fw_merge_ub, flow_ref) / flow_ref) ** exponent,
            )

            if _fw_merge_ub <= flow_ref:
                model.fw_screen3_constant_cost_cons.add(
                    model.fw_screen3_cost_if[s] == cost_ref
                )
            else:
                pts = cost_capacity_breakpoints(_fw_merge_lb, _fw_merge_ub, flow_ref)
                add_cost_capacity_pw(
                    model,
                    f"pw_fw_screen3_cost_{s}",
                    model.fw_screen3_cost_if[s],
                    model.Q_fw_after_shredder1,
                    pts,
                    cost_ref, flow_ref, exponent,
                )

        FW_SCREEN3_COST_M = max(1.0, _max_fw_screen3_cost * 10.0)
        model.fw_screen3_cost_select_cons = pyo.ConstraintList()
        for s in model.FW_SCREEN:
            add_bigm_select(
                model.fw_screen3_cost_select_cons,
                model.fw_screen3_cost_selected[s],
                model.fw_screen3_cost_if[s],
                model.fw_screen3_sel[s],
                FW_SCREEN3_COST_M,
            )

        # waste discharge hopper 및 conveyor 2대:
        # screen 1·2와 screen 3에서 제거된 총 waste 유량 기준
        # 가능한 유량 전 구간이 기준용량 이하이면 최소비용으로 고정
        model.fw_waste_constant_cost_cons = pyo.ConstraintList()

        for e in model.FW_WASTE_EQUIP:
            cost_ref, flow_ref, exponent = FW_WASTE_HANDLING_COST_DATA[e]

            if _fw_waste_ub <= flow_ref:
                # 가능한 유량 전체가 min_vol 이하이면 min_cost 고정
                model.fw_waste_constant_cost_cons.add(
                    model.fw_waste_cost_if[e] == cost_ref
                )
            else:
                pts = cost_capacity_breakpoints(
                    _fw_waste_lb,
                    _fw_waste_ub,
                    flow_ref,
                )

                add_cost_capacity_pw(
                    model,
                    f"pw_fw_waste_cost_{e}",
                    model.fw_waste_cost_if[e],
                    model.Q_fw_waste_total,
                    pts,
                    cost_ref, flow_ref, exponent,
                )

        model.fw_waste_handling_cost_eq = pyo.Constraint(
            expr=model.fw_waste_handling_cost ==
            sum(model.fw_waste_cost_if[e] for e in model.FW_WASTE_EQUIP)
        )

    # screen 1·2 + sorting 1·2 + shredder 1 + screen 3 + shredder 2 + waste 배출설비 비용
    model.eq8 = pyo.Constraint(
        expr=model.fw_pr_cost ==
        sum(
            model.fw_line_screen_sel[line, s] * FW_SCREEN_LINE_COST[s]
            for line in model.FW_LINE
            for s in model.FW_SCREEN
        )
        + sum(model.fw_sorting_cost[line] for line in model.FW_LINE)
        + sum(
            model.fw_shredder_cost_selected[stage, h]
            for stage in model.FW_SHREDDER_STAGE
            for h in model.FW_SHREDDER
        )
        + sum(
            model.fw_screen3_cost_selected[s]
            for s in model.FW_SCREEN
        )
        + model.fw_waste_handling_cost
    )

    model.eq8_1 = pyo.Constraint(
        expr=model.Q_fw_pr == model.Q_fw_after_shredder2
    )
    model.eq8_2 = pyo.Constraint(expr=model.Q_fw == model.Q_fw_pr)

    # ------------------- Sludge -------------------
    model.s1 = pyo.Var(domain=pyo.Binary)
    model.s2 = pyo.Var(domain=pyo.Binary)

    model.eq11 = pyo.Constraint(expr=model.s1 + model.s2 == 1)

    model.eq12_2_1 = pyo.Constraint(expr=model.W_sl_pr1 <= ctx.Water_sl * (1 - water_removal_sl[0]) + ctx.M_FLOW * (1 - model.s1))
    model.eq12_2_2 = pyo.Constraint(expr=model.W_sl_pr1 >= ctx.Water_sl * (1 - water_removal_sl[0]) - ctx.M_FLOW * (1 - model.s1))
    model.eq12_3_1 = pyo.Constraint(expr=model.W_sl_pr1 <= ctx.Water_sl * (1 - water_removal_sl[1]) + ctx.M_FLOW * (1 - model.s2))
    model.eq12_3_2 = pyo.Constraint(expr=model.W_sl_pr1 >= ctx.Water_sl * (1 - water_removal_sl[1]) - ctx.M_FLOW * (1 - model.s2))

    model.eq13_1 = pyo.Constraint(expr=model.Q_sl_pr == ctx.Solid_sl + model.W_sl_pr1)
    model.eq13_2 = pyo.Constraint(expr=model.Q_sl == model.Q_sl_pr)

    model.eq13_cost = pyo.Constraint(
        expr=model.C_sl_pr ==
        option_sl_pr_cost[0] * model.s1 +
        option_sl_pr_cost[1] * model.s2
    )

    # ------------------- Final TS by Piecewise -------------------
    Q_ani_lb = ctx.Solid_ani + ctx.Water_ani * (1 - max(water_removal_ani))
    Q_ani_ub = ctx.Solid_ani + ctx.Water_ani * (1 - min(water_removal_ani))

    # 음식물은 screen 1·2와 screen 3에서만 수분 제거
    _fw_screen_eff_values = list(FW_SCREEN_WATER_REMOVAL.values())
    Q_fw_lb = ctx.Solid_fw + ctx.Water_fw * (1 - max(_fw_screen_eff_values)) ** 2
    Q_fw_ub = ctx.Solid_fw + ctx.Water_fw * (1 - min(_fw_screen_eff_values)) ** 2

    Q_sl_lb = ctx.Solid_sl + ctx.Water_sl * (1 - max(water_removal_sl))
    Q_sl_ub = ctx.Solid_sl + ctx.Water_sl * (1 - min(water_removal_sl))

    model.COD_total = pyo.Var(domain=pyo.NonNegativeReals)
    model.Q_total = pyo.Var(domain=pyo.NonNegativeReals)

    # 원료별 COD(물리/surrogate)는 보고용 상수·나눗셈일 뿐 어떤 제약에도 들어가지
    # 않는다. v48/v49는 이걸 Var + Piecewise(이진변수 포함)로 만들어 MIP 크기만
    # 키웠다. 값은 get_model_summary()에서 산술로 계산한다.
    model.raw_flow_m3_d = pyo.Param(
        pyo.Set(initialize=["ani", "fw", "sl"]),
        initialize={"ani": ctx.W_m3_d, "fw": ctx.F_m3_d, "sl": ctx.S_m3_d},
        mutable=False,
    )

    model.Q_ani.setlb(Q_ani_lb)
    model.Q_ani.setub(Q_ani_ub)
    model.Q_fw.setlb(Q_fw_lb)
    model.Q_fw.setub(Q_fw_ub)
    model.Q_sl.setlb(Q_sl_lb)
    model.Q_sl.setub(Q_sl_ub)


    # ---------------------------------------------------------
    # 전처리 완료 후 AD 이송설비 비용
    # 각 설비는 실제 전처리 완료 유량 Q_ani/Q_sl/Q_fw를 기준으로 계산
    # ---------------------------------------------------------
    _pretreatment_aux_flow_var = {
        "AM_feeding_hopper": model.Q_ani,
        "AM_conveyor_1": model.Q_ani,
        "AM_conveyor_2": model.Q_ani,
        "SL_feeding_hopper": model.Q_sl,
        "SL_conveyor_1": model.Q_sl,
        "SL_conveyor_2": model.Q_sl,
        "FW_feeding_hopper": model.Q_fw,
        "FW_conveyor_1": model.Q_fw,
        "FW_conveyor_2": model.Q_fw,
    }

    _pretreatment_aux_raw_flow = {
        "AM_feeding_hopper": ctx.W,
        "AM_conveyor_1": ctx.W,
        "AM_conveyor_2": ctx.W,
        "SL_feeding_hopper": ctx.S,
        "SL_conveyor_1": ctx.S,
        "SL_conveyor_2": ctx.S,
        "FW_feeding_hopper": ctx.F,
        "FW_conveyor_1": ctx.F,
        "FW_conveyor_2": ctx.F,
    }

    model.pretreatment_aux_zero_cons = pyo.ConstraintList()

    for equip in model.PRETREATMENT_AUX:
        cost_ref, flow_ref, exponent = PRETREATMENT_AUX_COST_DATA[equip]
        flow_var = _pretreatment_aux_flow_var[equip]

        if _pretreatment_aux_raw_flow[equip] <= 0:
            # 해당 원료 자체가 없으면 설비비용 0
            model.pretreatment_aux_zero_cons.add(
                model.pretreatment_aux_cost_if[equip] == 0.0
            )

        elif flow_var.ub <= flow_ref:
            # 가능한 유량 전 구간이 기준용량 이하이면 최소비용으로 고정
            model.pretreatment_aux_zero_cons.add(
                model.pretreatment_aux_cost_if[equip] == cost_ref
            )

        else:
            # 일부 또는 전체 구간이 기준용량을 초과할 때만 Piecewise 생성
            pts = cost_capacity_breakpoints(
                flow_var.lb,
                flow_var.ub,
                flow_ref,
            )

            add_cost_capacity_pw(
                model,
                f"pw_pretreatment_aux_cost_{equip}",
                model.pretreatment_aux_cost_if[equip],
                flow_var,
                pts,
                cost_ref, flow_ref, exponent,
            )
    model.pretreatment_aux_cost_eq = pyo.Constraint(
        expr=model.pretreatment_aux_cost ==
        sum(
            model.pretreatment_aux_cost_if[equip]
            for equip in model.PRETREATMENT_AUX
        )
    )

    # 개별 원료 유량이 0이면 Piecewise를 만들지 않고 Q와 COD를 0으로 고정
    # pw_pts가 [0, 0, 0, 0, 0]이 되거나 0/0 계산이 발생하는 문제 방지
    if ctx.W <= 0:
        model.Q_ani_zero = pyo.Constraint(expr=model.Q_ani == 0.0)

    if ctx.F <= 0:
        model.Q_fw_zero = pyo.Constraint(expr=model.Q_fw == 0.0)

    if ctx.S <= 0:
        model.Q_sl_zero = pyo.Constraint(expr=model.Q_sl == 0.0)



    model.ADM_STATE = pyo.Set(initialize=ADM_STATES)
    model.eq102_flow = pyo.Constraint(expr=model.Q_total == model.Q_ani + model.Q_fw + model.Q_sl)

    ctx.Q_total_lb = Q_ani_lb + Q_fw_lb + Q_sl_lb
    ctx.Q_total_ub = Q_ani_ub + Q_fw_ub + Q_sl_ub

    model.Q_total.setlb(ctx.Q_total_lb)
    model.Q_total.setub(ctx.Q_total_ub)

    # =========================================================
    # 전처리 후 중간 저장조(BM storage) 유량 연결
    # Q_total -> 2계열 균등 분할 -> 저장 후 재합류 -> Q_ad_in
    # 저장·혼합·이송 과정의 유량 손실은 0으로 가정
    # =========================================================
    model.Q_bm_storage_in.setlb(ctx.Q_total_lb)
    model.Q_bm_storage_in.setub(ctx.Q_total_ub)
    model.Q_bm_line1.setlb(ctx.Q_total_lb / 2.0)
    model.Q_bm_line1.setub(ctx.Q_total_ub / 2.0)
    model.Q_bm_line2.setlb(ctx.Q_total_lb / 2.0)
    model.Q_bm_line2.setub(ctx.Q_total_ub / 2.0)
    model.Q_bm_storage_out.setlb(ctx.Q_total_lb)
    model.Q_bm_storage_out.setub(ctx.Q_total_ub)
    model.Q_ad_in.setlb(ctx.Q_total_lb)
    model.Q_ad_in.setub(ctx.Q_total_ub)

    model.bm_storage_in_eq = pyo.Constraint(
        expr=model.Q_bm_storage_in == model.Q_total
    )
    model.bm_line1_eq = pyo.Constraint(
        expr=model.Q_bm_line1 == model.Q_bm_storage_in / 2.0
    )
    model.bm_line2_eq = pyo.Constraint(
        expr=model.Q_bm_line2 == model.Q_bm_storage_in / 2.0
    )
    model.bm_storage_out_eq = pyo.Constraint(
        expr=model.Q_bm_storage_out == model.Q_bm_line1 + model.Q_bm_line2
    )
    model.q_ad_in_eq = pyo.Constraint(
        expr=model.Q_ad_in == model.Q_bm_storage_out
    )

    # 2계열 장비는 각 계열 유량 Q_total/2를 처리한다.
    _bm_flow_var = {
        "feeding_hopper_1": model.Q_bm_line1,
        "feeding_hopper_2": model.Q_bm_line2,
        "storage_tank_1": model.Q_bm_line1,
        "storage_tank_2": model.Q_bm_line2,
        "conveyor_1": model.Q_bm_line1,
        "conveyor_2": model.Q_bm_line2,
    }

    model.bm_storage_constant_cost_cons = pyo.ConstraintList()

    for equip in model.BM_STORAGE_EQUIP:
        cost_ref, flow_ref, exponent = BM_STORAGE_COST_DATA[equip]
        flow_var = _bm_flow_var[equip]

        if ctx.Q_total_ub <= 0:
            model.bm_storage_constant_cost_cons.add(
                model.bm_storage_cost_if[equip] == 0.0
            )
        elif flow_var.ub <= flow_ref:
            # 전 가능한 유량이 min_vol 이하이면 min_cost 적용
            model.bm_storage_constant_cost_cons.add(
                model.bm_storage_cost_if[equip] == cost_ref
            )
        else:
            pts = cost_capacity_breakpoints(flow_var.lb, flow_var.ub, flow_ref)
            add_cost_capacity_pw(
                model,
                f"pw_bm_storage_cost_{equip}",
                model.bm_storage_cost_if[equip],
                flow_var,
                pts,
                cost_ref, flow_ref, exponent,
            )

    model.bm_storage_cost_eq = pyo.Constraint(
        expr=model.bm_storage_cost ==
        sum(
            model.bm_storage_cost_if[equip]
            for equip in model.BM_STORAGE_EQUIP
        )
    )

    # =========================================================
    # 유입유량에 따른 AD 계열 구성
    #
    # 원칙:
    # - 0 < Q_ad_in <= 50 m3/d   : 1계열, AD 1기
    # - 50 < Q_ad_in <= 100      : 2계열, 계열당 AD 1기
    # - 100 < Q_ad_in <= 150     : 3계열, 계열당 AD 1기
    # - 150 < Q_ad_in <= 200     : 4계열, 계열당 AD 1기
    # - 이후에도 50 m3/d 증가할 때마다 계열 1개와 AD 1기 추가
    #
    # 한 계열에 AD 여러 기를 두지 않으며, 계열 수 = AD 설치 대수이다.
    # 선택된 모든 계열에는 총 AD 유입유량을 균등 분배한다.
    #
    # 기준점:
    # - AD 1기 처리 기준유량: 50 m3/d
    # - AD 1기 기준부피: 2,300 m3
    # - AD 1기 기준비용: 1,600,910 USD
    # - cost-capacity factor: 0.67
    # =========================================================

    # 해당 지역에서 필요할 수 있는 최대 계열 수


# -----------------------------------------------------------------------
# S4  혐기소화 — AD 계열 구성 · HRT · 반응조/교반기 비용 · NN surrogate
# -----------------------------------------------------------------------
def _stage_digestion(model, ctx):
    MAX_AD_TRAINS = max(
        1,
        int(math.ceil(ctx.Q_total_ub / REACTOR_FLOW_REF)),
    )

    model.AD_TRAIN_COUNT_OPTION = pyo.RangeSet(1, MAX_AD_TRAINS)
    model.AD_TRAIN = pyo.RangeSet(1, MAX_AD_TRAINS)

    # 정확히 하나의 계열 수를 선택
    model.ad_train_count_sel = pyo.Var(
        model.AD_TRAIN_COUNT_OPTION,
        domain=pyo.Binary,
    )

    model.ad_train_count_one = pyo.Constraint(
        expr=sum(
            model.ad_train_count_sel[n]
            for n in model.AD_TRAIN_COUNT_OPTION
        ) == 1
    )

    model.ad_train_count = pyo.Var(
        domain=pyo.NonNegativeIntegers,
        bounds=(1, MAX_AD_TRAINS),
    )
    model.ad_train_count_eq = pyo.Constraint(
        expr=model.ad_train_count
        == sum(
            n * model.ad_train_count_sel[n]
            for n in model.AD_TRAIN_COUNT_OPTION
        )
    )

    # 선택 계열 수 n은 ((n-1)*50, n*50] 구간에서만 활성화
    AD_SWITCH_M = max(
        1.0,
        ctx.Q_total_ub + REACTOR_FLOW_REF + 1.0,
    )

    model.ad_train_count_range_cons = pyo.ConstraintList()

    for n in model.AD_TRAIN_COUNT_OPTION:
        upper = n * REACTOR_FLOW_REF
        lower = (n - 1) * REACTOR_FLOW_REF

        # 선택된 n계열의 총 처리용량 이내
        model.ad_train_count_range_cons.add(
            model.Q_ad_in
            <= upper
            + AD_SWITCH_M * (1 - model.ad_train_count_sel[n])
        )

        # n=1은 Q_ad_in > 0이면 충분하므로 별도 양의 하한을 강제하지 않음
        if n > 1:
            model.ad_train_count_range_cons.add(
                model.Q_ad_in
                >= lower + AD_SWITCH_EPS
                - AD_SWITCH_M * (1 - model.ad_train_count_sel[n])
            )

    # 각 계열의 활성 여부
    # n계열 선택 시 1~n번 계열은 활성, 나머지는 비활성
    model.ad_train_active = pyo.Var(
        model.AD_TRAIN,
        domain=pyo.Binary,
    )

    model.ad_train_active_cons = pyo.ConstraintList()
    for t in model.AD_TRAIN:
        model.ad_train_active_cons.add(
            model.ad_train_active[t]
            == sum(
                model.ad_train_count_sel[n]
                for n in model.AD_TRAIN_COUNT_OPTION
                if n >= t
            )
        )

    # 각 계열 유량
    model.Q_ad_train = pyo.Var(
        model.AD_TRAIN,
        domain=pyo.NonNegativeReals,
        bounds=(0.0, ctx.Q_total_ub),
    )

    AD_FLOW_M = max(1.0, ctx.Q_total_ub)

    model.ad_train_flow_cons = pyo.ConstraintList()

    # n계열 선택 시 활성 계열 각각에 Q_ad_in/n 균등 분배
    for t in model.AD_TRAIN:
        model.ad_train_flow_cons.add(
            model.Q_ad_train[t]
            <= AD_FLOW_M * model.ad_train_active[t]
        )

        for n in model.AD_TRAIN_COUNT_OPTION:
            if t <= n:
                model.ad_train_flow_cons.add(
                    model.Q_ad_train[t]
                    <= model.Q_ad_in / n
                    + AD_FLOW_M * (1 - model.ad_train_count_sel[n])
                )
                model.ad_train_flow_cons.add(
                    model.Q_ad_train[t]
                    >= model.Q_ad_in / n
                    - AD_FLOW_M * (1 - model.ad_train_count_sel[n])
                )

    model.ad_train_flow_balance = pyo.Constraint(
        expr=sum(
            model.Q_ad_train[t]
            for t in model.AD_TRAIN
        ) == model.Q_ad_in
    )

    # HRT 40일 고정: 50 m3/d 모듈의 설계부피는 2,000 m3
    AD_VOLUME_PER_FLOW = HRT_DAYS

    model.V_reactor_train = pyo.Var(
        model.AD_TRAIN,
        domain=pyo.NonNegativeReals,
    )
    model.V_reactor = pyo.Var(domain=pyo.NonNegativeReals)

    model.ad_volume_train_cons = pyo.ConstraintList()
    for t in model.AD_TRAIN:
        model.ad_volume_train_cons.add(
            model.V_reactor_train[t]
            == AD_VOLUME_PER_FLOW * model.Q_ad_train[t]
        )

    model.ad_total_volume_eq = pyo.Constraint(
        expr=model.V_reactor
        == sum(
            model.V_reactor_train[t]
            for t in model.AD_TRAIN
        )
    )

    # ---------------------------------------------------------
    # AD 비용
    #
    # 50 m3/d 처리 모듈 1기의 기준비용은 1,600,910 USD로 고정.
    # 총유량이 50 m3/d를 넘으면 계열과 모듈 수가 증가한다.
    # ---------------------------------------------------------
    model.C_reactor = pyo.Var(domain=pyo.NonNegativeReals)
    model.ad_total_cost_eq = pyo.Constraint(
        expr=model.C_reactor
        == REACTOR_CAPEX_REF * model.ad_train_count
    )

    # AD 반응조 교반기(vertical mixer): AD 1기당 4개.
    # 각 계열 유량 Q_ad_train ≤ 50 m3/d 로 mixer 기준용량(1000)보다 항상 작아
    # cost-capacity floor 가 적용되어 1개당 비용은 기준비용 660 으로 고정된다.
    # 따라서 총 교반기 비용 = 4 × 660 × ad_train_count.
    model.ad_mixer_cost = pyo.Var(domain=pyo.NonNegativeReals)
    model.ad_mixer_cost_eq = pyo.Constraint(
        expr=model.ad_mixer_cost
        == AD_MIXER_PER_TRAIN * AD_MIXER_REF_COST * model.ad_train_count
    )

    if ctx.Q_total_ub <= _PULP_REF:
        model.pw_Biomass_pulping = pyo.Constraint(expr=model.Biomass_pulping_cost == 65000.0)
    else:
        if ctx.Q_total_lb >= _PULP_REF:
            _pulp_pts = [ctx.Q_total_lb, ctx.Q_total_lb+0.25*(ctx.Q_total_ub-ctx.Q_total_lb), ctx.Q_total_lb+0.50*(ctx.Q_total_ub-ctx.Q_total_lb), ctx.Q_total_lb+0.75*(ctx.Q_total_ub-ctx.Q_total_lb), ctx.Q_total_ub]
        else:
            _pulp_pts = [ctx.Q_total_lb, _PULP_REF, _PULP_REF+(ctx.Q_total_ub-_PULP_REF)/3.0, _PULP_REF+2.0*(ctx.Q_total_ub-_PULP_REF)/3.0, ctx.Q_total_ub]
        model.pw_Biomass_pulping = pyo.Piecewise(
            model.Biomass_pulping_cost, model.Q_total, pw_pts=_pulp_pts,
            f_rule=lambda m, x: 65000 * (max(x, _PULP_REF) / _PULP_REF) ** 0.6,
            pw_constr_type="EQ", pw_repn="DCC",
        warning_tol=0.0,
        )

    Q_raw_total = ctx.W_m3_d + ctx.F_m3_d + ctx.S_m3_d
    # [Fix 1] COD_BASIS 에 따라 원료별 실측 COD 또는 학습 basis 를 사용한다.
    total_cod_mass = (
        ctx.W_m3_d * SURROGATE_COD_COEF
        + ctx.F_m3_d * SURROGATE_COD_COEF
        + ctx.S_m3_d * SURROGATE_COD_COEF
    )

    if Q_raw_total <= 0:
        raise ValueError("W + F + S must be positive for surrogate inputs.")
    if total_cod_mass <= 0:
        raise ValueError("total_cod_mass must be positive for surrogate inputs.")

    model.COD_total.setlb(total_cod_mass / ctx.Q_total_ub if ctx.Q_total_ub > 0 else 0.0)
    model.COD_total.setub(total_cod_mass / ctx.Q_total_lb if ctx.Q_total_lb > 0 else 0.0)

    q_total_cod_pts = [ctx.Q_total_lb, ctx.Q_total_lb+0.25*(ctx.Q_total_ub-ctx.Q_total_lb), ctx.Q_total_lb+0.50*(ctx.Q_total_ub-ctx.Q_total_lb), ctx.Q_total_lb+0.75*(ctx.Q_total_ub-ctx.Q_total_lb), ctx.Q_total_ub]

    model.pw_COD_total = pyo.Piecewise(
        model.COD_total, model.Q_total, pw_pts=q_total_cod_pts,
        f_rule=lambda m, x: total_cod_mass / x, pw_constr_type="EQ", pw_repn="DCC", warning_tol=0.0)

    model.ADM_total = pyo.Var(model.ADM_STATE, domain=pyo.NonNegativeReals)

    # NN 입력용 ADM 상태도 학습 당시 COD feature basis를 유지한다.
    total_adm_mass = {
        s: (
            ctx.W_m3_d * SURROGATE_COD_COEF * RATIO_ANI[s]
            + ctx.F_m3_d * SURROGATE_COD_COEF * RATIO_FW[s]
            + ctx.S_m3_d * SURROGATE_COD_COEF * RATIO_SL[s]
        )
        for s in ADM_STATES
    }
    mixed_ratio = {s: total_adm_mass[s] / total_cod_mass for s in ADM_STATES}

    for s in ADM_STATES:
        adm_lb = (total_adm_mass[s] / ctx.Q_total_ub) if ctx.Q_total_ub > 0 else 0.0
        adm_ub = (total_adm_mass[s] / ctx.Q_total_lb) if ctx.Q_total_lb > 0 else 0.0
        model.ADM_total[s].setlb(min(adm_lb, adm_ub))
        model.ADM_total[s].setub(max(adm_lb, adm_ub))
        setattr(model, f"ADM_total_pretreated_eq_{s}",
            pyo.Constraint(expr=model.ADM_total[s] == model.COD_total * mixed_ratio[s]))

    model.OLR = pyo.Var(domain=pyo.NonNegativeReals)
    model.OLR.setlb((model.COD_total.lb or 0.0) / HRT_DAYS)
    model.OLR.setub((model.COD_total.ub or 0.0) / HRT_DAYS)
    model.OLR_eq = pyo.Constraint(expr=model.OLR == model.COD_total / HRT_DAYS)

    # OMLT 연결 전 NN 입력범위 진단값을 모델에 저장
    model.surrogate_cod_basis = pyo.Param(
        initialize=SURROGATE_COD_COEF,
        mutable=False,
    )

    add_biogas_omlt_surrogate(model)


# -----------------------------------------------------------------------
# S5  가스계통 — 설비 선택변수 · route(정제/발전) · 경계값
# -----------------------------------------------------------------------
def _stage_gas_setup(model, ctx):
    model.TANK = pyo.Set(initialize=tanks)
    model.DESULF = pyo.Set(initialize=desulf_techs)
    model.DEHUM = pyo.Set(initialize=dehum_techs)
    model.TECH = pyo.Set(initialize=purification_techs)

    model.t = pyo.Var(model.TANK, domain=pyo.Binary)
    model.ds = pyo.Var(model.DESULF, domain=pyo.Binary)
    model.dh = pyo.Var(model.DEHUM, domain=pyo.Binary)
    model.y = pyo.Var(model.TECH, domain=pyo.Binary)

    # 정제(도시가스) vs 발전(전기) 택1 선택변수.
    # 뒤의 가스유량 분기와 정제/발전 설비 선택을 모두 이 변수로 게이팅한다.
    model.ROUTE = pyo.Set(initialize=["purification", "generator"])
    model.route_sel = pyo.Var(model.ROUTE, domain=pyo.Binary)
    model.one_route = pyo.Constraint(
        expr=sum(model.route_sel[r] for r in model.ROUTE) == 1
    )
    # [v56] route 는 택1이며(가스를 나누지 않는다), 어느 쪽을 평가할지는
    #  시나리오로 강제한다. LCOB(정제)와 LCOE(발전)는 단위가 달라 하나의
    #  목적함수로 비교할 수 없으므로, 각각 따로 풀어 나란히 보고한다.
    if ctx.force_route in ("purification", "generator"):
        model.force_route_con = pyo.Constraint(
            expr=model.route_sel[ctx.force_route] == 1
        )

    # 각 공정에서 정확히 하나의 기술만 선택
    model.one_tank = pyo.Constraint(
        expr=sum(model.t[i] for i in model.TANK) == 1
    )
    model.one_desulf = pyo.Constraint(
        expr=sum(model.ds[d] for d in model.DESULF) == 1
    )
    model.one_dehum = pyo.Constraint(
        expr=sum(model.dh[d] for d in model.DEHUM) == 1
    )
    # 정제 route 일 때만 정제기술 1개 선택, 발전 route 면 0개(=정제설비 미설치→비용0)
    model.one_tech = pyo.Constraint(
        expr=sum(model.y[j] for j in model.TECH)
        == model.route_sel["purification"]
    )

    model.pur_in_q_gas = pyo.Var(domain=pyo.NonNegativeReals)
    model.pur_in_ch4_frac = pyo.Var(bounds=(0, 1))
    model.pur_in_co2_frac = pyo.Var(bounds=(0, 1))

    y_raw_min, y_raw_max = load_biogas_output_bounds()

    ctx.QGAS_UB = max(
        QGAS_LB + 1.0,
        float(y_raw_max["final_q_gas(m3/d)"]) + RAW_RANGE_TOL,
    )

    ctx.CH4_FRAC_LB = 0.0
    ctx.CH4_FRAC_UB = 1.0
    ctx.CO2_FRAC_LB = 0.0
    ctx.CO2_FRAC_UB = 1.0

    # [Fix 2] 이중선형 격자 '범위' 축소
    #  v48은 [0, 학습최대가스량 44,336] x [0, 1] 전 구간에 11점 균등격자를 깔았다.
    #  실제 해는 가스 8,400 m3/d 이하, CH4 분율 0.53~0.63 이라 격자가 지나치게 성기고,
    #  셀 하나의 근사오차 상한이 dx*dy/4 이므로 그대로 +7% 편향이 됐다.
    #  probe 해에서 얻은 실제 범위로 격자를 좁힌다(해가 벗어나면 solve_case가 전 범위로 재시도).
    if ctx.grid_hint:
        _qh = ctx.grid_hint.get("qgas")
        if _qh is not None and float(_qh) > 0.0:
            ctx.QGAS_UB = float(
                min(ctx.QGAS_UB, max(float(_qh) * QGAS_HINT_MARGIN, 1.0))
            )
        _lo = ctx.grid_hint.get("ch4_lo")
        _hi = ctx.grid_hint.get("ch4_hi")
        if _lo is not None and _hi is not None:
            ctx.CH4_FRAC_LB = max(0.0, float(_lo) - CH4_FRAC_MARGIN)
            ctx.CH4_FRAC_UB = min(1.0, float(_hi) + CH4_FRAC_MARGIN)
            ctx.CO2_FRAC_LB = max(0.0, 1.0 - ctx.CH4_FRAC_UB)
            ctx.CO2_FRAC_UB = min(1.0, 1.0 - ctx.CH4_FRAC_LB)

    model.pur_in_q_gas.setlb(QGAS_LB)
    model.pur_in_q_gas.setub(ctx.QGAS_UB)

    # =========================================================
    # AD 이후 가스계통 유량 연결
    #
    # pred_q_gas
    #   -> 가스 저장조
    #   -> 탈황
    #   -> 제습
    #   -> 정제
    #
    # 현재는 저장·탈황·제습 과정의 총가스 손실을 0으로 가정한다.
    # =========================================================


# -----------------------------------------------------------------------
# S6  가스계통 — 유량 연결 · 성분(CH4/CO2) 물질수지
# -----------------------------------------------------------------------
def _stage_gas_balance(model, ctx):
    model.storage_in_q_gas = pyo.Var(
        domain=pyo.NonNegativeReals,
        bounds=(0.0, ctx.QGAS_UB),
    )
    model.storage_out_q_gas = pyo.Var(
        domain=pyo.NonNegativeReals,
        bounds=(0.0, ctx.QGAS_UB),
    )
    model.desulf_in_q_gas = pyo.Var(
        domain=pyo.NonNegativeReals,
        bounds=(0.0, ctx.QGAS_UB),
    )
    model.desulf_out_q_gas = pyo.Var(
        domain=pyo.NonNegativeReals,
        bounds=(0.0, ctx.QGAS_UB),
    )
    model.dehum_in_q_gas = pyo.Var(
        domain=pyo.NonNegativeReals,
        bounds=(0.0, ctx.QGAS_UB),
    )
    model.dehum_out_q_gas = pyo.Var(
        domain=pyo.NonNegativeReals,
        bounds=(0.0, ctx.QGAS_UB),
    )

    # [v57] 소화조 가온 자가소비: 생산가스의 일부를 출구에서 바로 빼낸다.
    #  후단(저장-탈황-제습-정제/발전) 설비는 남은 가스만 처리한다.
    ctx._gas_keep = 1.0 - AD_HEATING_CH4_FRAC_OF_GAS
    model.ad_heating_gas_m3_d = pyo.Var(
        domain=pyo.NonNegativeReals, bounds=(0.0, ctx.QGAS_UB)
    )
    model.ad_heating_gas_eq = pyo.Constraint(
        expr=model.ad_heating_gas_m3_d
        == AD_HEATING_CH4_FRAC_OF_GAS * model.pred_q_gas_var
    )
    model.storage_in_q_gas_eq = pyo.Constraint(
        expr=model.storage_in_q_gas
        == model.pred_q_gas_var - model.ad_heating_gas_m3_d
    )
    model.storage_out_q_gas_eq = pyo.Constraint(
        expr=model.storage_out_q_gas == model.storage_in_q_gas
    )
    model.desulf_in_q_gas_eq = pyo.Constraint(
        expr=model.desulf_in_q_gas == model.storage_out_q_gas
    )
    model.desulf_out_q_gas_eq = pyo.Constraint(
        expr=model.desulf_out_q_gas == model.desulf_in_q_gas
    )
    model.dehum_in_q_gas_eq = pyo.Constraint(
        expr=model.dehum_in_q_gas == model.desulf_out_q_gas
    )
    model.dehum_out_q_gas_eq = pyo.Constraint(
        expr=model.dehum_out_q_gas == model.dehum_in_q_gas
    )
    # 제습 후 총 바이오가스를 정제분기와 발전분기로 분할
    model.gen_biogas_flow = pyo.Var(
        domain=pyo.NonNegativeReals,
        bounds=(0.0, ctx.QGAS_UB),
    )
    model.gas_branch_balance = pyo.Constraint(
        expr=model.dehum_out_q_gas
        == model.pur_in_q_gas + model.gen_biogas_flow
    )
    # 택1: 선택 안 된 분기 유량은 0 (route_sel 은 앞에서 정의)
    model.pur_route_flow = pyo.Constraint(
        expr=model.pur_in_q_gas <= ctx.QGAS_UB * model.route_sel["purification"]
    )
    model.gen_route_flow = pyo.Constraint(
        expr=model.gen_biogas_flow <= ctx.QGAS_UB * model.route_sel["generator"]
    )

    model.pur_in_ch4_frac_eq = pyo.Constraint(
        expr=model.pur_in_ch4_frac == model.pred_ch4_frac_dry_var
    )
    model.pur_in_co2_frac_eq = pyo.Constraint(
        expr=model.pur_in_co2_frac == model.pred_co2_frac_var
    )

    # 삼각형 기반 bilinear surface grid
    # 전체 0~QGAS_UB, 0~1 범위가 아니라 실제 surrogate 출력범위만 사용해
    # 근사오차를 줄이면서 binary 수는 과도하게 늘리지 않는다.
    # probe는 비용 범위 추정이 목적이므로 grid를 줄여 계산을 안정화한다.
    # 최종 epsilon sweep은 기존 11점 grid를 유지한다.
    gas_grid_n = 5 if ctx.probe_mode else int(GAS_GRID_N)

    qgas_pts = make_even_points(0.0, ctx.QGAS_UB, n=gas_grid_n)
    ch4_frac_pts = make_even_points(
        ctx.CH4_FRAC_LB,
        ctx.CH4_FRAC_UB,
        n=gas_grid_n,
    )
    co2_frac_pts = make_even_points(
        ctx.CO2_FRAC_LB,
        ctx.CO2_FRAC_UB,
        n=gas_grid_n,
    )

    model.pur_in_ch4_frac.setlb(ctx.CH4_FRAC_LB)
    model.pur_in_ch4_frac.setub(ctx.CH4_FRAC_UB)
    model.pur_in_co2_frac.setlb(ctx.CO2_FRAC_LB)
    model.pur_in_co2_frac.setub(ctx.CO2_FRAC_UB)

    # ---------------------------------------------------------
    # AD 출구 전체 건식 바이오가스 성분유량
    #
    # CH4와 CO2를 각각 독립 Piecewise로 근사하면
    # CH4_frac + CO2_frac = 1이어도 보간오차 때문에
    # CH4_flow + CO2_flow = total_gas가 정확히 성립하지 않아
    # 모델 전체가 구조적으로 infeasible이 될 수 있다.
    #
    # 따라서 CH4만 Piecewise로 계산하고,
    # CO2는 총가스에서 CH4를 뺀 잔차로 정의한다.
    # ---------------------------------------------------------
    model.pur_in_ch4_flow = pyo.Var(
        domain=pyo.NonNegativeReals,
        bounds=(0.0, ctx.QGAS_UB),
    )
    model.pur_in_co2_flow = pyo.Var(
        domain=pyo.NonNegativeReals,
        bounds=(0.0, ctx.QGAS_UB),
    )

    add_piecewise_bilinear_product(
        model,
        model.dehum_out_q_gas,
        model.pur_in_ch4_frac,
        model.pur_in_ch4_flow,
        qgas_pts,
        ch4_frac_pts,
        "pw_pur_in_ch4_flow",
    )

    model.pur_in_co2_flow_eq = pyo.Constraint(
        expr=model.pur_in_co2_flow
        == model.dehum_out_q_gas - model.pur_in_ch4_flow
    )

    # ---------------------------------------------------------
    # 정제분기와 발전분기의 성분유량
    #
    # 정제분기도 CH4만 Piecewise로 계산하고,
    # CO2는 정제분기 총가스에서 CH4를 뺀 잔차로 정의한다.
    # 발전분기는 전체 성분유량에서 정제분기를 뺀다.
    # ---------------------------------------------------------
    model.purification_ch4_flow = pyo.Var(
        domain=pyo.NonNegativeReals,
        bounds=(0.0, ctx.QGAS_UB),
    )
    model.purification_co2_flow = pyo.Var(
        domain=pyo.NonNegativeReals,
        bounds=(0.0, ctx.QGAS_UB),
    )
    model.gen_ch4_flow = pyo.Var(
        domain=pyo.NonNegativeReals,
        bounds=(0.0, ctx.QGAS_UB),
    )
    model.gen_co2_flow = pyo.Var(
        domain=pyo.NonNegativeReals,
        bounds=(0.0, ctx.QGAS_UB),
    )

    # [Fix 2] 정제분기 CH4 유량에는 두 번째 이중선형 근사가 필요 없다.
    #  route 는 택1이므로
    #     purification_ch4_flow = pur_in_ch4_flow x route_sel["purification"]
    #  이고, 이 이진 x 연속 곱은 big-M 으로 '정확히' 선형화된다.
    #  v48은 여기서 근사를 한 번 더 해 오차원과 이진변수를 두 배로 늘렸다.
    model.purification_ch4_gate_cons = pyo.ConstraintList()
    add_bigm_select(
        model.purification_ch4_gate_cons,
        model.purification_ch4_flow,
        model.pur_in_ch4_flow,
        model.route_sel["purification"],
        ctx.QGAS_UB,
    )

    model.purification_co2_flow_eq = pyo.Constraint(
        expr=model.purification_co2_flow
        == model.pur_in_q_gas - model.purification_ch4_flow
    )

    model.gen_ch4_flow_eq = pyo.Constraint(
        expr=model.gen_ch4_flow
        == model.pur_in_ch4_flow - model.purification_ch4_flow
    )
    model.gen_co2_flow_eq = pyo.Constraint(
        expr=model.gen_co2_flow
        == model.pur_in_co2_flow - model.purification_co2_flow
    )

    # 가스 성분 물질수지 검산 제약
    model.total_component_sum = pyo.Constraint(
        expr=model.pur_in_ch4_flow
        + model.pur_in_co2_flow
        == model.dehum_out_q_gas
    )
    model.purification_branch_component_sum = pyo.Constraint(
        expr=model.purification_ch4_flow
        + model.purification_co2_flow
        == model.pur_in_q_gas
    )
    model.generator_branch_component_sum = pyo.Constraint(
        expr=model.gen_ch4_flow
        + model.gen_co2_flow
        == model.gen_biogas_flow
    )


# -----------------------------------------------------------------------
# S7  탈황 · 제습 · 가스저장 · 잉여가스 연소
# -----------------------------------------------------------------------
def _stage_gas_cleanup(model, ctx):
    model.desulf_cost_if = pyo.Var(
        model.DESULF,
        domain=pyo.NonNegativeReals,
    )
    model.desulf_cost_selected = pyo.Var(
        model.DESULF,
        domain=pyo.NonNegativeReals,
    )
    model.desulf_cost = pyo.Var(domain=pyo.NonNegativeReals)

    max_desulf_cost = 0.0
    model.desulf_constant_cost_cons = pyo.ConstraintList()

    for d in desulf_techs:
        cost_ref, flow_ref, exponent = desulf_cost_data[d]

        max_desulf_cost = max(
            max_desulf_cost,
            cost_ref
            * (max(ctx.QGAS_UB, flow_ref) / flow_ref) ** exponent,
        )

        if ctx.QGAS_UB <= flow_ref:
            model.desulf_constant_cost_cons.add(
                model.desulf_cost_if[d] == cost_ref
            )
        else:
            pts = cost_capacity_breakpoints(0.0, ctx.QGAS_UB, flow_ref)
            add_cost_capacity_pw(
                model,
                f"pw_desulf_cost_{d}",
                model.desulf_cost_if[d],
                model.desulf_in_q_gas,
                pts,
                cost_ref, flow_ref, exponent,
            )

    DESULF_COST_M = max(1.0, 2.0 * max_desulf_cost)
    model.desulf_cost_select_cons = pyo.ConstraintList()

    for d in desulf_techs:
        add_bigm_select(
            model.desulf_cost_select_cons,
            model.desulf_cost_selected[d],
            model.desulf_cost_if[d],
            model.ds[d],
            DESULF_COST_M,
        )

    model.desulf_cost_eq = pyo.Constraint(
        expr=model.desulf_cost
        == sum(
            model.desulf_cost_selected[d]
            for d in model.DESULF
        )
    )


    model.dehum_cost_if = pyo.Var(
        model.DEHUM,
        domain=pyo.NonNegativeReals,
    )
    model.dehum_cost_selected = pyo.Var(
        model.DEHUM,
        domain=pyo.NonNegativeReals,
    )
    model.dehum_cost = pyo.Var(domain=pyo.NonNegativeReals)

    max_dehum_cost = 0.0
    model.dehum_constant_cost_cons = pyo.ConstraintList()

    for d in dehum_techs:
        cost_ref, flow_ref, exponent = dehum_cost_data[d]

        max_dehum_cost = max(
            max_dehum_cost,
            cost_ref
            * (max(ctx.QGAS_UB, flow_ref) / flow_ref) ** exponent,
        )

        if ctx.QGAS_UB <= flow_ref:
            model.dehum_constant_cost_cons.add(
                model.dehum_cost_if[d] == cost_ref
            )
        else:
            pts = cost_capacity_breakpoints(0.0, ctx.QGAS_UB, flow_ref)
            add_cost_capacity_pw(
                model,
                f"pw_dehum_cost_{d}",
                model.dehum_cost_if[d],
                model.dehum_in_q_gas,
                pts,
                cost_ref, flow_ref, exponent,
            )

    DEHUM_COST_M = max(1.0, 2.0 * max_dehum_cost)
    model.dehum_cost_select_cons = pyo.ConstraintList()

    for d in dehum_techs:
        add_bigm_select(
            model.dehum_cost_select_cons,
            model.dehum_cost_selected[d],
            model.dehum_cost_if[d],
            model.dh[d],
            DEHUM_COST_M,
        )

    model.dehum_cost_eq = pyo.Constraint(
        expr=model.dehum_cost
        == sum(
            model.dehum_cost_selected[d]
            for d in model.DEHUM
        )
    )


    model.tank_cost_if = pyo.Var(model.TANK, domain=pyo.NonNegativeReals)
    model.tank_cost_selected = pyo.Var(model.TANK, domain=pyo.NonNegativeReals)
    model.tank_cost = pyo.Var(domain=pyo.NonNegativeReals)

    # [Fix 5] v48은 (a) 절점을 0~QGAS_UB 균등 11점으로 깔고 (b) f(0)=0 으로 두어
    #  (0, 0) 과 (첫 절점, 기준비용) 사이가 직선보간되면서 최소기준비용 floor 가 무너졌다.
    #  (v48 결과 192행 중 28행의 tank_cost 가 자기 최소기준비용보다 낮았다)
    #  다른 설비와 동일하게 floor 를 유지하고 기준용량을 절점으로 포함시킨다.
    max_tank_cost = 0.0
    for tank in tanks:
        cost_ref, flow_ref, exponent = tank_cost_data[tank]
        max_tank_cost = max(
            max_tank_cost,
            cost_ref * (max(ctx.QGAS_UB, flow_ref) / flow_ref) ** exponent,
        )
        _tank_pts = cost_capacity_breakpoints(0.0, ctx.QGAS_UB, flow_ref, n_segments=8)
        add_cost_capacity_pw(
            model,
            f"pw_tank_cost_{tank}",
            model.tank_cost_if[tank],
            model.storage_in_q_gas,
            _tank_pts,
            cost_ref, flow_ref, exponent,
            warning_tol=0.0,
        )

    TANK_COST_M = max(1.0, max_tank_cost * 10.0)
    model.tank_cost_select_cons = pyo.ConstraintList()
    for tank in tanks:
        add_bigm_select(
            model.tank_cost_select_cons,
            model.tank_cost_selected[tank],
            model.tank_cost_if[tank],
            model.t[tank],
            TANK_COST_M,
        )
    model.tank_cost_eq = pyo.Constraint(expr=model.tank_cost == sum(model.tank_cost_selected[tank] for tank in model.TANK))


    model.thermal_oxidizer_cost = pyo.Var(domain=pyo.NonNegativeReals)

    if ctx.QGAS_UB <= THERMAL_OXIDIZER_MIN_VOL:
        model.thermal_oxidizer_cost_eq = pyo.Constraint(
            expr=model.thermal_oxidizer_cost == THERMAL_OXIDIZER_MIN_COST
        )
    else:
        thermal_oxidizer_pts = cost_capacity_breakpoints(
            0.0, ctx.QGAS_UB, THERMAL_OXIDIZER_MIN_VOL,
        )
        model.pw_thermal_oxidizer_cost = pyo.Piecewise(
            model.thermal_oxidizer_cost,
            model.dehum_out_q_gas,
            pw_pts=thermal_oxidizer_pts,
            f_rule=lambda m, x:
                THERMAL_OXIDIZER_MIN_COST
                * (max(x, THERMAL_OXIDIZER_MIN_VOL) / THERMAL_OXIDIZER_MIN_VOL)
                ** THERMAL_OXIDIZER_SF,
            pw_constr_type="EQ", pw_repn="DCC", warning_tol=0.0,
        )


# -----------------------------------------------------------------------
# S8  가스압축 · 정제(PSA / membrane) 비용
# -----------------------------------------------------------------------
def _stage_upgrading(model, ctx):
    # ---------------------------------------------------------
    # 필수 Gas compressor 비용
    # ---------------------------------------------------------
    # 정제 route 일 때만 gas compressor 비용 발생 (발전 route 면 0).
    # 원가는 _if 로 계산하고, route 로 big-M 게이팅한다.
    model.gas_compressor_cost_if = pyo.Var(domain=pyo.NonNegativeReals)
    model.gas_compressor_cost = pyo.Var(domain=pyo.NonNegativeReals)

    if ctx.QGAS_UB <= GAS_COMPRESSOR_MIN_VOL:
        model.gas_compressor_cost_eq = pyo.Constraint(
            expr=model.gas_compressor_cost_if == GAS_COMPRESSOR_MIN_COST
        )
    else:
        gas_compressor_pts = cost_capacity_breakpoints(
            0.0,
            ctx.QGAS_UB,
            GAS_COMPRESSOR_MIN_VOL,
        )
        model.pw_gas_compressor_cost = pyo.Piecewise(
            model.gas_compressor_cost_if,
            model.pur_in_q_gas,
            pw_pts=gas_compressor_pts,
            f_rule=lambda m, x:
                GAS_COMPRESSOR_MIN_COST
                * (
                    max(x, GAS_COMPRESSOR_MIN_VOL)
                    / GAS_COMPRESSOR_MIN_VOL
                ) ** GAS_COMPRESSOR_SF,
            pw_constr_type="EQ",
            pw_repn="DCC",
        )

    GAS_COMPRESSOR_COST_M = max(
        1.0,
        2.0 * GAS_COMPRESSOR_MIN_COST
        * (max(ctx.QGAS_UB, GAS_COMPRESSOR_MIN_VOL) / GAS_COMPRESSOR_MIN_VOL)
        ** GAS_COMPRESSOR_SF,
    )
    model.gas_compressor_gate_cons = pyo.ConstraintList()
    add_bigm_select(
        model.gas_compressor_gate_cons,
        model.gas_compressor_cost,
        model.gas_compressor_cost_if,
        model.route_sel["purification"],
        GAS_COMPRESSOR_COST_M,
    )

    # ---------------------------------------------------------
    # 선택 정제기술 비용: PSA 또는 membrane 중 1개
    # ---------------------------------------------------------
    model.pure_cost_if = pyo.Var(
        model.TECH,
        domain=pyo.NonNegativeReals,
    )
    model.pure_cost_selected = pyo.Var(
        model.TECH,
        domain=pyo.NonNegativeReals,
    )
    model.pure_cost = pyo.Var(domain=pyo.NonNegativeReals)

    max_pure_cost = 0.0
    model.pure_constant_cost_cons = pyo.ConstraintList()

    for tech in purification_techs:
        cost_ref, flow_ref, exponent = purification_cost_data[tech]

        max_pure_cost = max(
            max_pure_cost,
            cost_ref
            * (max(ctx.QGAS_UB, flow_ref) / flow_ref) ** exponent,
        )

        if ctx.QGAS_UB <= flow_ref:
            model.pure_constant_cost_cons.add(
                model.pure_cost_if[tech] == cost_ref
            )
        else:
            pts = cost_capacity_breakpoints(0.0, ctx.QGAS_UB, flow_ref)
            add_cost_capacity_pw(
                model,
                f"pw_pure_cost_{tech}",
                model.pure_cost_if[tech],
                model.pur_in_q_gas,
                pts,
                cost_ref, flow_ref, exponent,
            )

    PURE_COST_M = max(1.0, 2.0 * max_pure_cost)
    model.pure_cost_select_cons = pyo.ConstraintList()

    for tech in purification_techs:
        add_bigm_select(
            model.pure_cost_select_cons,
            model.pure_cost_selected[tech],
            model.pure_cost_if[tech],
            model.y[tech],
            PURE_COST_M,
        )

    # 최종 정제비용 = 필수 compressor + 선택된 PSA/membrane 비용
    model.pure_cost_eq = pyo.Constraint(
        expr=model.pure_cost
        == model.gas_compressor_cost
        + sum(
            model.pure_cost_selected[tech]
            for tech in model.TECH
        )
    )


# -----------------------------------------------------------------------
# S9  정제 성능 — 회수율 · CO2 제거율 · 제품 순도
# -----------------------------------------------------------------------
def _stage_upgrading_perf(model, ctx):
    model.up_ch4_if = pyo.Var(model.TECH, domain=pyo.NonNegativeReals)
    model.up_co2_if = pyo.Var(model.TECH, domain=pyo.NonNegativeReals)
    model.up_gas_if = pyo.Var(model.TECH, domain=pyo.NonNegativeReals)

    model.up_ch4_selected = pyo.Var(model.TECH, domain=pyo.NonNegativeReals)
    model.up_co2_selected = pyo.Var(model.TECH, domain=pyo.NonNegativeReals)
    model.up_gas_selected = pyo.Var(model.TECH, domain=pyo.NonNegativeReals)

    model.up_ch4 = pyo.Var(domain=pyo.NonNegativeReals)
    model.up_co2 = pyo.Var(domain=pyo.NonNegativeReals)
    model.up_gas = pyo.Var(domain=pyo.NonNegativeReals)
    model.up_ch4_frac = pyo.Var(bounds=(0, 1))

    model.up_perf_cons = pyo.ConstraintList()
    for tech in purification_techs:
        model.up_perf_cons.add(model.up_ch4_if[tech] == model.purification_ch4_flow * ch4_recovery[tech])
        model.up_perf_cons.add(model.up_co2_if[tech] == model.purification_co2_flow * (1 - co2_removal[tech]))
        model.up_perf_cons.add(model.up_gas_if[tech] == model.up_ch4_if[tech] + model.up_co2_if[tech])

    FLOW_M = max(1.0, ctx.QGAS_UB * 2.0)
    model.up_select_cons = pyo.ConstraintList()
    for tech in purification_techs:
        add_bigm_select(
            model.up_select_cons,
            model.up_ch4_selected[tech],
            model.up_ch4_if[tech],
            model.y[tech],
            FLOW_M,
        )
        add_bigm_select(
            model.up_select_cons,
            model.up_co2_selected[tech],
            model.up_co2_if[tech],
            model.y[tech],
            FLOW_M,
        )
        add_bigm_select(
            model.up_select_cons,
            model.up_gas_selected[tech],
            model.up_gas_if[tech],
            model.y[tech],
            FLOW_M,
        )

    model.up_ch4_eq = pyo.Constraint(expr=model.up_ch4 == sum(model.up_ch4_selected[tech] for tech in model.TECH))
    model.up_co2_eq = pyo.Constraint(expr=model.up_co2 == sum(model.up_co2_selected[tech] for tech in model.TECH))
    model.up_gas_eq = pyo.Constraint(expr=model.up_gas == sum(model.up_gas_selected[tech] for tech in model.TECH))

    model.up_ch4_frac_eq = pyo.Constraint(
        expr=model.up_ch4_frac == sum(product_ch4_purity[tech] * model.y[tech] for tech in model.TECH))


# -----------------------------------------------------------------------
# S10 발전(CHP) — 설비 선택 · 출력 · 비용
# -----------------------------------------------------------------------
def _stage_power(model, ctx):
    model.GEN = pyo.Set(initialize=gens)

    # get_model_summary()에서도 접근할 수 있도록
    # 발전기 비용 기준값을 model Param으로 저장
    model.gen_min_capacity_kW = pyo.Param(
        model.GEN,
        initialize={
            g: gen_cost_data[g][0]
            for g in gens
        },
        mutable=False,
    )
    model.gen_min_cost_USD = pyo.Param(
        model.GEN,
        initialize={
            g: gen_cost_data[g][1]
            for g in gens
        },
        mutable=False,
    )
    model.gen_scale_factor = pyo.Param(
        model.GEN,
        initialize={
            g: gen_cost_data[g][2]
            for g in gens
        },
        mutable=False,
    )
    model.g = pyo.Var(model.GEN, domain=pyo.Binary)

    # 발전 route 일 때만 발전기 1개 선택, 정제 route 면 0개(=발전기 미설치→비용0)
    model.one_gen = pyo.Constraint(
        expr=sum(model.g[g] for g in model.GEN)
        == model.route_sel["generator"]
    )

    model.up_ch4.setlb(0.0)
    model.up_ch4.setub(ctx.QGAS_UB)

    # 가능한 최대 CH4 분율을 1로 둔 보수적 발전용량 상한
    ctx.GEN_POWER_UB = max(
        1.0,
        ctx.QGAS_UB * CH4_LHV_KWH * GEN_ELEC_EFF / 24.0,
    )

    model.gen_power_kW = pyo.Var(
        domain=pyo.NonNegativeReals,
        bounds=(0.0, ctx.GEN_POWER_UB),
    )

    model.gen_power_kW_eq = pyo.Constraint(
        expr=model.gen_power_kW
        == model.gen_ch4_flow
        * CH4_LHV_KWH
        * GEN_ELEC_EFF
        / 24.0
    )

    model.gen_cost_if = pyo.Var(
        model.GEN,
        domain=pyo.NonNegativeReals,
    )
    model.gen_cost_selected = pyo.Var(
        model.GEN,
        domain=pyo.NonNegativeReals,
    )
    model.gen_cost = pyo.Var(domain=pyo.NonNegativeReals)

    max_gen_cost = 0.0
    model.gen_constant_cost_cons = pyo.ConstraintList()

    for gname in gens:
        min_generation_kW, min_cost, exponent = gen_cost_data[gname]

        candidate_max_cost = (
            min_cost
            * (
                max(ctx.GEN_POWER_UB, min_generation_kW)
                / min_generation_kW
            ) ** exponent
        )

        max_gen_cost = max(
            max_gen_cost,
            candidate_max_cost,
        )

        if ctx.GEN_POWER_UB <= min_generation_kW:
            # 가능한 전 구간이 최소 발전용량 이하이면 최소비용 고정
            model.gen_constant_cost_cons.add(
                model.gen_cost_if[gname] == min_cost
            )

        else:
            gen_pts = cost_capacity_breakpoints(
                0.0,
                ctx.GEN_POWER_UB,
                min_generation_kW,
                n_segments=12,
            )

            # [Fix 5] v48은 x<=0 에서 0을 반환해, 절점 0 과 최소용량 사이가
            #  직선보간되며 최소기준비용 floor 가 무너졌다
            #  (예: 8.17 kW 발전기 비용 6,811 USD, 정확식 25,000 USD, -73%).
            #  미설치 시 비용 0 은 g[gname] big-M 게이팅이 이미 처리한다.
            add_cost_capacity_pw(
                model,
                f"pw_gen_cost_{gname}",
                model.gen_cost_if[gname],
                model.gen_power_kW,
                gen_pts,
                min_cost, min_generation_kW, exponent,
                warning_tol=0.0,
            )

    GEN_COST_M = max(1.0, 2.0 * max_gen_cost)

    model.gen_cost_select_cons = pyo.ConstraintList()

    for gname in gens:
        # 비선택 설비의 selected cost는 0
        add_bigm_select(
            model.gen_cost_select_cons,
            model.gen_cost_selected[gname],
            model.gen_cost_if[gname],
            model.g[gname],
            GEN_COST_M,
        )

    model.gen_cost_eq = pyo.Constraint(
        expr=model.gen_cost
        == sum(
            model.gen_cost_selected[gname]
            for gname in model.GEN
        )
    )


# -----------------------------------------------------------------------
# S11 소화액 탈수 · AD 물질수지
# -----------------------------------------------------------------------
def _stage_digestate(model, ctx):
    model.DIGESTATE_DEWATERING = pyo.Set(
        initialize=digestate_dewatering_techs
    )

    model.dd = pyo.Var(
        model.DIGESTATE_DEWATERING,
        domain=pyo.Binary,
    )

    model.one_digestate_dewatering = pyo.Constraint(
        expr=sum(
            model.dd[d]
            for d in model.DIGESTATE_DEWATERING
        ) == 1
    )

    # ---------------------------------------------------------
    # AD 출구 물질수지
    #
    # 액상 유입 질량
    #   = 소화슬러지 질량 + CH4 질량 + CO2 질량
    #
    # 주의:
    # Q_ad_in과 pred_q_gas는 모두 m3/d이지만 액체와 기체이므로
    # 부피끼리 직접 빼지 않고, 먼저 질량[kg/d]으로 환산한다.
    #
    # 현재 surrogate가 건식 CH4와 CO2만 출력하므로,
    # 수증기·H2S·NH3 등 미량가스는 제외한 근사 물질수지이다.
    # ---------------------------------------------------------
    # 보수적인 소화슬러지 유량 bounds
    MAX_DRY_BIOGAS_MASS_KG_D = (
        ctx.QGAS_UB * max(CH4_GAS_DENSITY_KG_M3, CO2_GAS_DENSITY_KG_M3)
    )
    ctx.DIGESTATE_FLOW_LB = max(
        0.0,
        (
            ctx.Q_total_lb * RAW_FEED_DENSITY_KG_M3
            - MAX_DRY_BIOGAS_MASS_KG_D
        )
        / DIGESTATE_DENSITY_KG_M3,
    )
    ctx.DIGESTATE_FLOW_UB = (
        ctx.Q_total_ub * RAW_FEED_DENSITY_KG_M3
        / DIGESTATE_DENSITY_KG_M3
    )

    model.ad_feed_mass_kg_d = pyo.Var(
        domain=pyo.NonNegativeReals,
    )
    model.biogas_ch4_mass_kg_d = pyo.Var(
        domain=pyo.NonNegativeReals,
    )
    model.biogas_co2_mass_kg_d = pyo.Var(
        domain=pyo.NonNegativeReals,
    )
    model.biogas_dry_mass_kg_d = pyo.Var(
        domain=pyo.NonNegativeReals,
    )
    model.digestate_mass_kg_d = pyo.Var(
        domain=pyo.NonNegativeReals,
    )
    model.digestate_flow = pyo.Var(
        domain=pyo.NonNegativeReals,
        bounds=(ctx.DIGESTATE_FLOW_LB, ctx.DIGESTATE_FLOW_UB),
    )

    # Pyomo Piecewise와 완전히 같은 float endpoint를 사용하도록
    # 실제 변수 bound를 다시 읽어 동기화한다.
    ctx.DIGESTATE_FLOW_LB = float(model.digestate_flow.lb)
    ctx.DIGESTATE_FLOW_UB = float(model.digestate_flow.ub)

    # AD 액상 유입질량
    model.ad_feed_mass_eq = pyo.Constraint(
        expr=model.ad_feed_mass_kg_d
        == model.Q_ad_in * RAW_FEED_DENSITY_KG_M3
    )

    # 건식 바이오가스 성분별 질량
    # pur_in_ch4_flow / pur_in_co2_flow는 AD 예측가스의 성분 유량이며,
    # 저장·탈황·제습 전까지 총가스 손실 0 가정을 사용한다.
    # [v57] pur_in_ch4_flow 는 가온분을 뺀 '후단' CH4 이므로, 소화조가 실제로
    #  내보낸 총 CH4 로 되돌려 물질수지에 쓴다(가온분도 소화조에서 나온 가스다).
    model.biogas_ch4_mass_eq = pyo.Constraint(
        expr=model.biogas_ch4_mass_kg_d
        == (model.pur_in_ch4_flow / ctx._gas_keep) * CH4_GAS_DENSITY_KG_M3
    )
    model.biogas_co2_mass_eq = pyo.Constraint(
        expr=model.biogas_co2_mass_kg_d
        == (model.pur_in_co2_flow / ctx._gas_keep) * CO2_GAS_DENSITY_KG_M3
    )
    model.biogas_dry_mass_eq = pyo.Constraint(
        expr=model.biogas_dry_mass_kg_d
        == model.biogas_ch4_mass_kg_d
        + model.biogas_co2_mass_kg_d
    )

    # 전체 AD 물질수지
    model.digestate_mass_balance_eq = pyo.Constraint(
        expr=model.digestate_mass_kg_d
        == model.ad_feed_mass_kg_d
        - model.biogas_dry_mass_kg_d
    )

    # 소화슬러지 질량을 부피유량으로 환산
    model.digestate_flow_eq = pyo.Constraint(
        expr=model.digestate_flow
        == model.digestate_mass_kg_d / DIGESTATE_DENSITY_KG_M3
    )

    # =========================================================
    # 소화슬러지 탈수 물질수지  [v51: 고형물 수지를 실제로 닫음]
    #
    # v48~v50:  digestate_dry_solids = digestate_mass x 0.05  (TS 5% 가정)
    #   -> 원료 고형물, 바이오가스 발생량과 아무 연결이 없어
    #      "원료 고형물 - 가스 - 소화액 고형물" 잔차가 최대 -4,260 kg/d 였다.
    #      (없는 고형물을 만들어내고 있었다는 뜻)
    #
    # v51:
    #   COD 제거량   = CH4 발생량 / 0.35            [kgCOD/d]
    #   VS 분해량    = COD 제거량 / 1.5             [kgVS/d]
    #   소화액 고형물 = 유입 건고형물 - VS 분해량     [kg/d]
    #
    # 0.35 는 이론 최대수율이므로 COD 제거량은 하한, 즉 VS 분해량도 하한이고
    # 소화액 고형물은 상한(보수적: 케이크량과 탈수비를 과소평가하지 않음)이다.
    # 소화액 TS 는 가정이 아니라 이 수지의 결과로 결정된다.
    # =========================================================
    ctx.ad_feed_dry_solids_kg_d = float(
        (ctx.Solid_ani + ctx.Solid_fw + ctx.Solid_sl) * RAW_FEED_DENSITY_KG_M3
    )

    model.ad_feed_dry_solids_kg_d = pyo.Param(
        initialize=ctx.ad_feed_dry_solids_kg_d, mutable=False
    )
    model.cod_removed_kg_d = pyo.Var(domain=pyo.NonNegativeReals)
    model.vs_destroyed_kg_d = pyo.Var(domain=pyo.NonNegativeReals)

    model.cod_removed_eq = pyo.Constraint(
        expr=model.cod_removed_kg_d
        == (model.pur_in_ch4_flow / ctx._gas_keep) / CH4_YIELD_M3_PER_KGCOD
    )
    model.vs_destroyed_eq = pyo.Constraint(
        expr=model.vs_destroyed_kg_d
        == model.cod_removed_kg_d / COD_PER_VS_KG_PER_KG
    )

    model.digestate_dry_solids_kg_d = pyo.Var(domain=pyo.NonNegativeReals)
    model.captured_solids_kg_d = pyo.Var(domain=pyo.NonNegativeReals)
    model.cake_mass_kg_d = pyo.Var(domain=pyo.NonNegativeReals)
    model.cake_flow_m3_d = pyo.Var(
        domain=pyo.NonNegativeReals,
        bounds=(0.0, ctx.DIGESTATE_FLOW_UB),
    )
    model.centrate_mass_kg_d = pyo.Var(domain=pyo.NonNegativeReals)
    model.centrate_flow_m3_d = pyo.Var(
        domain=pyo.NonNegativeReals,
        bounds=(0.0, ctx.DIGESTATE_FLOW_UB),
    )

    model.digestate_dry_solids_eq = pyo.Constraint(
        expr=model.digestate_dry_solids_kg_d
        == model.ad_feed_dry_solids_kg_d - model.vs_destroyed_kg_d
    )
    model.captured_solids_eq = pyo.Constraint(
        expr=model.captured_solids_kg_d
        == model.digestate_dry_solids_kg_d
        * SOLIDS_CAPTURE_EFFICIENCY
    )
    model.cake_mass_eq = pyo.Constraint(
        expr=model.cake_mass_kg_d
        == model.captured_solids_kg_d / CAKE_TS_FRACTION
    )
    model.cake_flow_eq = pyo.Constraint(
        expr=model.cake_flow_m3_d
        == model.cake_mass_kg_d / CAKE_DENSITY_KG_M3
    )
    model.centrate_mass_eq = pyo.Constraint(
        expr=model.centrate_mass_kg_d
        == model.digestate_mass_kg_d - model.cake_mass_kg_d
    )
    model.centrate_flow_eq = pyo.Constraint(
        expr=model.centrate_flow_m3_d
        == model.centrate_mass_kg_d / CENTRATE_DENSITY_KG_M3
    )
    model.dewatering_mass_balance_eq = pyo.Constraint(
        expr=model.digestate_mass_kg_d
        == model.cake_mass_kg_d + model.centrate_mass_kg_d
    )


    # get_model_summary()에서도 접근할 수 있도록
    # 탈수기 비용 기준값을 model Param으로 저장
    model.digestate_min_vol_m3d = pyo.Param(
        model.DIGESTATE_DEWATERING,
        initialize={
            d: digestate_dewatering_cost_data[d][0]
            for d in digestate_dewatering_techs
        },
        mutable=False,
    )
    model.digestate_min_cost_USD = pyo.Param(
        model.DIGESTATE_DEWATERING,
        initialize={
            d: digestate_dewatering_cost_data[d][1]
            for d in digestate_dewatering_techs
        },
        mutable=False,
    )
    model.digestate_scale_factor = pyo.Param(
        model.DIGESTATE_DEWATERING,
        initialize={
            d: digestate_dewatering_cost_data[d][2]
            for d in digestate_dewatering_techs
        },
        mutable=False,
    )


    model.digestate_dewatering_cost_if = pyo.Var(
        model.DIGESTATE_DEWATERING,
        domain=pyo.NonNegativeReals,
    )

    model.digestate_dewatering_cost_selected = pyo.Var(
        model.DIGESTATE_DEWATERING,
        domain=pyo.NonNegativeReals,
    )

    model.DIGESTATE_CONVEYOR = pyo.Set(
        initialize=list(digestate_conveyor_cost_data.keys())
    )

    model.digestate_conveyor_ref_vol_m3d = pyo.Param(
        model.DIGESTATE_CONVEYOR,
        initialize={
            c: digestate_conveyor_cost_data[c][0]
            for c in digestate_conveyor_cost_data
        },
        mutable=False,
    )
    model.digestate_conveyor_ref_cost_USD = pyo.Param(
        model.DIGESTATE_CONVEYOR,
        initialize={
            c: digestate_conveyor_cost_data[c][1]
            for c in digestate_conveyor_cost_data
        },
        mutable=False,
    )
    model.digestate_conveyor_scale_factor = pyo.Param(
        model.DIGESTATE_CONVEYOR,
        initialize={
            c: digestate_conveyor_cost_data[c][2]
            for c in digestate_conveyor_cost_data
        },
        mutable=False,
    )

    model.digestate_conveyor_cost_if = pyo.Var(
        model.DIGESTATE_CONVEYOR,
        domain=pyo.NonNegativeReals,
    )

    model.digestate_dewatering_cost = pyo.Var(
        domain=pyo.NonNegativeReals,
    )

    # ---------------------------
    # belt press / screw press 비용
    # ---------------------------
    max_digestate_dewatering_cost = 0.0
    model.digestate_dewatering_constant_cost_cons = (
        pyo.ConstraintList()
    )

    for d in digestate_dewatering_techs:
        min_vol, min_cost, exponent = (
            digestate_dewatering_cost_data[d]
        )

        candidate_max_cost = (
            min_cost
            * (
                max(ctx.DIGESTATE_FLOW_UB, min_vol)
                / min_vol
            ) ** exponent
        )

        max_digestate_dewatering_cost = max(
            max_digestate_dewatering_cost,
            candidate_max_cost,
        )

        if ctx.DIGESTATE_FLOW_UB <= min_vol:
            model.digestate_dewatering_constant_cost_cons.add(
                model.digestate_dewatering_cost_if[d]
                == min_cost
            )

        else:
            pts = cost_capacity_breakpoints(
                ctx.DIGESTATE_FLOW_LB,
                ctx.DIGESTATE_FLOW_UB,
                min_vol,
                n_segments=12,
            )

            add_cost_capacity_pw(
                model,
                f"pw_digestate_dewatering_cost_{d}",
                model.digestate_dewatering_cost_if[d],
                model.digestate_flow,
                pts,
                min_cost, min_vol, exponent,
                warning_tol=0.0,
            )

    DIGESTATE_DEWATERING_COST_M = max(
        1.0,
        2.0 * max_digestate_dewatering_cost,
    )

    model.digestate_dewatering_cost_select_cons = (
        pyo.ConstraintList()
    )

    for d in digestate_dewatering_techs:
        add_bigm_select(
            model.digestate_dewatering_cost_select_cons,
            model.digestate_dewatering_cost_selected[d],
            model.digestate_dewatering_cost_if[d],
            model.dd[d],
            DIGESTATE_DEWATERING_COST_M,
        )

    # ---------------------------
    # DS conveyor / SC conveyor 비용
    # DS: 탈수 전 소화슬러지, SC: 탈수 후 케이크
    # ---------------------------
    digestate_conveyor_flow_var = {
        "DS_conveyor": model.digestate_flow,
        "SC_conveyor": model.cake_flow_m3_d,
    }

    model.digestate_conveyor_constant_cost_cons = pyo.ConstraintList()

    for conveyor in model.DIGESTATE_CONVEYOR:
        min_vol, min_cost, exponent = digestate_conveyor_cost_data[conveyor]
        flow_var = digestate_conveyor_flow_var[conveyor]
        flow_lb = float(flow_var.lb or 0.0)
        flow_ub = float(flow_var.ub or ctx.DIGESTATE_FLOW_UB)

        if flow_ub <= min_vol:
            model.digestate_conveyor_constant_cost_cons.add(
                model.digestate_conveyor_cost_if[conveyor] == min_cost
            )
        else:
            pts = cost_capacity_breakpoints(
                flow_lb,
                flow_ub,
                min_vol,
                n_segments=12,
            )

            add_cost_capacity_pw(
                model,
                f"pw_digestate_conveyor_cost_{conveyor}",
                model.digestate_conveyor_cost_if[conveyor],
                flow_var,
                pts,
                min_cost, min_vol, exponent,
                warning_tol=0.0,
            )

    # 총 소화슬러지 탈수시설 비용
    model.digestate_dewatering_cost_eq = pyo.Constraint(
        expr=model.digestate_dewatering_cost
        == sum(
            model.digestate_dewatering_cost_selected[d]
            for d in model.DIGESTATE_DEWATERING
        )
        + sum(
            model.digestate_conveyor_cost_if[c]
            for c in model.DIGESTATE_CONVEYOR
        )
    )


# -----------------------------------------------------------------------
# S12 폐수처리 (A2O / Bardenpho)
# -----------------------------------------------------------------------
def _stage_wwtp(model, ctx):
    model.WWTP_OPTION = pyo.Set(initialize=wwtp_options)
    model.wwtp_select = pyo.Var(
        model.WWTP_OPTION,
        domain=pyo.Binary,
    )

    model.one_wwtp_option = pyo.Constraint(
        expr=sum(
            model.wwtp_select[o]
            for o in model.WWTP_OPTION
        ) == 1
    )

    # 폐수처리 설계유량 [m3/d]
    # v51에서 소화액 TS 가 계산값이 되면서 centrate 비율도 더는 상수가 아니다.
    # 케이크 질량 상한은 '유입 건고형물이 전혀 분해되지 않은 경우'이므로
    #   cake_mass_max = 유입건고형물 x capture / TS_cake
    # 이고, centrate 하한은 digestate 하한에서 그만큼 뺀 값이다.
    _cake_mass_ub = (
        ctx.ad_feed_dry_solids_kg_d
        * SOLIDS_CAPTURE_EFFICIENCY
        / CAKE_TS_FRACTION
    )
    WASTEWATER_FLOW_LB = max(
        0.0,
        (ctx.DIGESTATE_FLOW_LB * DIGESTATE_DENSITY_KG_M3 - _cake_mass_ub)
        / CENTRATE_DENSITY_KG_M3,
    )
    ctx.WASTEWATER_FLOW_UB = max(
        WASTEWATER_FLOW_LB + 1.0e-6,
        ctx.DIGESTATE_FLOW_UB
        * DIGESTATE_DENSITY_KG_M3
        / CENTRATE_DENSITY_KG_M3,
    )

    model.wastewater_in_flow = pyo.Var(
        domain=pyo.NonNegativeReals,
        bounds=(WASTEWATER_FLOW_LB, ctx.WASTEWATER_FLOW_UB),
    )
    model.wastewater_in_flow_eq = pyo.Constraint(
        expr=model.wastewater_in_flow == model.centrate_flow_m3_d
    )




    model.WWTP_EQUIPMENT = pyo.Set(
        initialize=list(WWTP_EQUIPMENT_COST_DATA.keys())
    )
    model.wwtp_equipment_cost_if = pyo.Var(
        model.WWTP_EQUIPMENT,
        domain=pyo.NonNegativeReals,
    )
    model.wwtp_option_cost_if = pyo.Var(
        model.WWTP_OPTION,
        domain=pyo.NonNegativeReals,
    )
    model.wwtp_option_cost_selected = pyo.Var(
        model.WWTP_OPTION,
        domain=pyo.NonNegativeReals,
    )
    model.wastewater_cost = pyo.Var(
        domain=pyo.NonNegativeReals,
    )

    model.wwtp_constant_cost_cons = pyo.ConstraintList()
    ctx.max_wwtp_option_cost = 0.0

    # =========================================================
    # [v72] 폐수처리 설비비
    #   설계유량 V = 폐수유입량 x 1.3 (엑셀 Base!AE2*1.3)
    # =========================================================
    model.wwtp_design_flow = pyo.Var(
        domain=pyo.NonNegativeReals,
        bounds=(WASTEWATER_FLOW_LB * WWTP_DESIGN_MARGIN,
                ctx.WASTEWATER_FLOW_UB * WWTP_DESIGN_MARGIN),
    )
    model.wwtp_design_flow_eq = pyo.Constraint(
        expr=model.wwtp_design_flow
        == WWTP_DESIGN_MARGIN * model.wastewater_in_flow
    )
    V_LB = WASTEWATER_FLOW_LB * WWTP_DESIGN_MARGIN
    V_UB = ctx.WASTEWATER_FLOW_UB * WWTP_DESIGN_MARGIN

    # 비용식 (New4.xlsx Data_CAPEX)
    def _f_anaerobic(V):
        return (WWTP_ANAEROBIC_A * max(V, 1.0e-6) ** WWTP_ANAEROBIC_B
                * (CEPCI_NOW / CEPCI_ANAEROBIC))

    def _f_aerobic(V):
        return (WWTP_AEROBIC_A * max(V, 1.0e-6) ** WWTP_AEROBIC_B
                * (CEPCI_NOW / CEPCI_AEROBIC))

    def _f_clarifier(V):
        return (WWTP_CLARIFIER_C2 * V ** 2
                + WWTP_CLARIFIER_C1 * V + WWTP_CLARIFIER_C0)

    _WWTP_FORM = {"anaerobic": _f_anaerobic,
                  "aerobic": _f_aerobic,
                  "clarifier": _f_clarifier}

    # [v72] Anammox 반응조 : (0.3148 V + 9.17) x 1e8 / 환율 x 0.229  (1차식)
    _anammox_slope_usd = (
        ANAMMOX_TANK_SLOPE_EOK_PER_M3D * KRW_PER_EOK
        / EXCHANGE_RATE_KRW_PER_USD
        * ANAMMOX_TANK_COST_FACTOR
    )
    _anammox_intercept_usd = (
        ANAMMOX_TANK_INTERCEPT_EOK * KRW_PER_EOK
        / EXCHANGE_RATE_KRW_PER_USD
        * ANAMMOX_TANK_COST_FACTOR
    )
    model.anammox_tank_cost_eq = pyo.Constraint(
        expr=model.wwtp_equipment_cost_if["Anammox_tank"]
        == _anammox_slope_usd * model.wwtp_design_flow
        + _anammox_intercept_usd
    )
    _anammox_tank_cost_ub = _anammox_slope_usd * V_UB + _anammox_intercept_usd

    # 설비별 상한 (Big-M 산정용)
    ctx._wwtp_cost_ub = {"Anammox_tank": _anammox_tank_cost_ub}

    for equip in model.WWTP_EQUIPMENT:
        if equip == "Anammox_tank":
            continue

        if equip in WWTP_SPECIAL_COST:          # [v72] 별도 비용식
            fn = _WWTP_FORM[WWTP_SPECIAL_COST[equip]]
            ctx._wwtp_cost_ub[equip] = fn(V_UB)
            if V_UB - V_LB < 1.0e-9:
                model.wwtp_constant_cost_cons.add(
                    model.wwtp_equipment_cost_if[equip] == fn(V_LB)
                )
                continue
            pts = [V_LB + (V_UB - V_LB) * i / 8.0 for i in range(9)]
            add_cost_capacity_pw(
                model,
                f"pw_wwtp_cost_{equip}",
                model.wwtp_equipment_cost_if[equip],
                model.wwtp_design_flow,
                pts,
                f_rule=lambda m, x, fn=fn: fn(x),
                warning_tol=0.0,
            )
            continue

        ref_vol, ref_cost, exponent = WWTP_EQUIPMENT_COST_DATA[equip]
        ctx._wwtp_cost_ub[equip] = (
            ref_cost * (max(V_UB, ref_vol) / ref_vol) ** exponent
        )

        if exponent == 0.0 or V_UB <= ref_vol:   # DAF 는 유량무관 고정
            model.wwtp_constant_cost_cons.add(
                model.wwtp_equipment_cost_if[equip] == ref_cost
            )
            ctx._wwtp_cost_ub[equip] = ref_cost
        else:
            pts = cost_capacity_breakpoints(
                V_LB, V_UB, ref_vol, n_segments=8,
            )
            add_cost_capacity_pw(
                model,
                f"pw_wwtp_cost_{equip}",
                model.wwtp_equipment_cost_if[equip],
                model.wwtp_design_flow,
                pts,
                ref_cost, ref_vol, exponent,
                warning_tol=0.0,
            )


# -----------------------------------------------------------------------
# S13 CAPEX 집계 · 부대비용
# -----------------------------------------------------------------------
def _stage_capex(model, ctx):
    model.wwtp_option_cost_eq = pyo.ConstraintList()

    for option in model.WWTP_OPTION:
        model.wwtp_option_cost_eq.add(
            model.wwtp_option_cost_if[option]
            == sum(
                model.wwtp_equipment_cost_if[equip]
                for equip in WWTP_OPTION_EQUIPMENT[option]
            )
        )

        option_max_cost = sum(
            ctx._wwtp_cost_ub[equip]
            for equip in WWTP_OPTION_EQUIPMENT[option]
        )
        ctx.max_wwtp_option_cost = max(
            ctx.max_wwtp_option_cost,
            option_max_cost,
        )

    WWTP_COST_M = max(
        1.0,
        2.0 * ctx.max_wwtp_option_cost,
    )

    model.wwtp_cost_select_cons = pyo.ConstraintList()

    for option in model.WWTP_OPTION:
        add_bigm_select(
            model.wwtp_cost_select_cons,
            model.wwtp_option_cost_selected[option],
            model.wwtp_option_cost_if[option],
            model.wwtp_select[option],
            WWTP_COST_M,
        )

    model.wastewater_cost_eq = pyo.Constraint(
        expr=model.wastewater_cost
        == sum(
            model.wwtp_option_cost_selected[option]
            for option in model.WWTP_OPTION
        )
    )

    # =========================================================
    # 수익 (USD / day)
    #   정제 수익 = 정제메탄량[m3/d] × 도시가스단가
    #   발전 수익 = 발전량[kWh/d] × 전기단가  (발전량 = gen_power_kW × 24)
    # =========================================================
    model.purification_revenue = pyo.Var(domain=pyo.NonNegativeReals)
    model.electricity_revenue = pyo.Var(domain=pyo.NonNegativeReals)
    model.total_revenue = pyo.Var(domain=pyo.NonNegativeReals)
    model.purification_revenue_eq = pyo.Constraint(
        expr=model.purification_revenue
        == model.up_ch4 * CITY_GAS_PRICE_USD_PER_M3_CH4
    )
    model.electricity_revenue_eq = pyo.Constraint(
        expr=model.electricity_revenue
        == model.gen_power_kW * 24.0 * ELEC_PRICE_USD_PER_KWH
    )
    model.total_revenue_eq = pyo.Constraint(
        expr=model.total_revenue
        == model.purification_revenue + model.electricity_revenue
    )

    # ------------------- Cost -------------------
    model.eq101 = pyo.Constraint(
        expr=model.Cost ==
        model.C_ani_pr1_cost +
        model.fw_pr_cost +
        model.C_sl_pr +
        model.pretreatment_aux_cost +
        model.bm_storage_cost +
        model.C_reactor +
        model.ad_mixer_cost +
        model.receiving_cost +
        model.Biomass_pulping_cost +
        model.tank_cost +
        model.thermal_oxidizer_cost +
        model.desulf_cost +
        model.dehum_cost +
        model.pure_cost +
        model.gen_cost +
        model.digestate_dewatering_cost +
        model.wastewater_cost
    )

    # [v75] epsilon 은 CAPEX 가 아니라 TAC(연환산 CAPEX + 연간 OPEX)에 건다.
    #  TAC 는 OPEX 가 확정되는 _stage_objective 에서 정의되므로 제약도 거기 있다.
    #  (v74 의 model.cost_cap = Cost <= epsilon 을 대체)

    # =========================================================
    # [v53] 지역 수요 대체
    #
    # 목적: max  대체편익 = 대체 도시가스 x 단가 + 대체 전기 x 단가  [USD/d]
    #       s.t. CAPEX <= epsilon
    #
    # 수요 상한이 있으므로 과잉생산에는 보상이 없다. 그 결과
    #   - 도시가스 수요가 작은(배관망 없는) 지역 -> 정제해도 팔 곳이 없어 발전이 유리
    #   - 도시가스 수요가 큰 지역               -> 정제가 유리
    # 로 route 선택이 지역별로 갈린다. 수요를 넣기 전에는 단가 구조상
    # 정제가 항상 이겨서(1 m3 CH4 당 정제 0.65 vs 발전 0.40 USD) 무의미했다.
    #
    # 가정은 판매단가 2개뿐이며 OPEX/할인/수명은 쓰지 않는다.
    # =========================================================
    _gas_prod_ub = float(ctx.QGAS_UB * CH4_LHV_KWH)
    _elec_prod_ub = float(ctx.GEN_POWER_UB * 24.0)

    _gas_dem = (
        _gas_prod_ub
        if ctx.gas_demand_kwh_d is None
        else max(0.0, float(ctx.gas_demand_kwh_d))
    )
    _elec_dem = (
        _elec_prod_ub
        if ctx.elec_demand_kwh_d is None
        else max(0.0, float(ctx.elec_demand_kwh_d))
    )

    model.gas_demand_kwh_d = pyo.Param(initialize=_gas_dem, mutable=False)
    model.elec_demand_kwh_d = pyo.Param(initialize=_elec_dem, mutable=False)
    model.gas_demand_is_given = pyo.Param(
        initialize=int(ctx.gas_demand_kwh_d is not None), mutable=False
    )
    model.elec_demand_is_given = pyo.Param(
        initialize=int(ctx.elec_demand_kwh_d is not None), mutable=False
    )

    # 생산량 [kWh/d]
    model.gas_product_kwh_d = pyo.Var(
        domain=pyo.NonNegativeReals, bounds=(0.0, _gas_prod_ub)
    )
    model.elec_product_kwh_d = pyo.Var(
        domain=pyo.NonNegativeReals, bounds=(0.0, _elec_prod_ub)
    )
    model.gas_product_eq = pyo.Constraint(
        expr=model.gas_product_kwh_d == model.up_ch4 * CH4_LHV_KWH
    )
    model.elec_product_eq = pyo.Constraint(
        expr=model.elec_product_kwh_d == model.gen_power_kW * 24.0
    )

    # 수요를 실제로 대체한 양 [kWh/d] = min(생산량, 수요)
    model.gas_displaced_kwh_d = pyo.Var(
        domain=pyo.NonNegativeReals, bounds=(0.0, min(_gas_prod_ub, _gas_dem))
    )
    model.elec_displaced_kwh_d = pyo.Var(
        domain=pyo.NonNegativeReals, bounds=(0.0, min(_elec_prod_ub, _elec_dem))
    )
    model.gas_displaced_cons = pyo.ConstraintList()
    model.gas_displaced_cons.add(
        model.gas_displaced_kwh_d <= model.gas_product_kwh_d
    )
    model.gas_displaced_cons.add(
        model.gas_displaced_kwh_d <= model.gas_demand_kwh_d
    )
    model.elec_displaced_cons = pyo.ConstraintList()
    model.elec_displaced_cons.add(
        model.elec_displaced_kwh_d <= model.elec_product_kwh_d
    )
    model.elec_displaced_cons.add(
        model.elec_displaced_kwh_d <= model.elec_demand_kwh_d
    )

    model.displacement_benefit_usd_d = pyo.Var(domain=pyo.NonNegativeReals)
    model.displacement_benefit_eq = pyo.Constraint(
        expr=model.displacement_benefit_usd_d
        == model.gas_displaced_kwh_d * CITY_GAS_PRICE_USD_PER_KWH
        + model.elec_displaced_kwh_d * ELEC_PRICE_USD_PER_KWH
    )

    # =========================================================
    # [v54] 연간 OPEX — 전부 New4.xlsx 'Data_OPEX' 원단위
    # =========================================================
    ctx.PE = OPEX_ELEC_PRICE_USD_PER_KWH
    ctx.DAYS = 365.0 * LCOB_F_UTIL


# -----------------------------------------------------------------------
# S14 OPEX · 수익 · 목적함수 · epsilon 제약
# -----------------------------------------------------------------------
def _stage_objective(model, ctx):
    def _gate(name, binvar, expr, ub):
        """selected = binvar * expr 를 big-M 으로 정확 선형화."""
        v = pyo.Var(domain=pyo.NonNegativeReals, bounds=(0.0, max(ub, 1.0)))
        setattr(model, name, v)
        cl = pyo.ConstraintList()
        setattr(model, name + "_cons", cl)
        M = max(ub, 1.0)
        add_bigm_select(
            cl,
            v,
            expr,
            binvar,
            M,
        )
        return v

    # =========================================================
    # [v75] 소비전력을 가동시간대별로 분리한다.
    #
    #   el_24h : 24시간 연속 가동 (컨베이어·에어커튼·농축·교반·AD·가스계통·탈수·폐수)
    #   el_8h  : 주간 8시간만 가동 (가축분뇨/음식물 전처리의 스크린·선별·파쇄)
    #
    # 같은 일일 kWh 라도 8h 설비는 평균전력(kW)이 3배라 기본요금·최대부하
    # 단가에 훨씬 크게 걸린다. v74 는 이 구분 없이 단일 단가를 곱했다.
    # 구분 기준은 capex_selector.py 의 equipment_daily_kwh() 와 동일하다.
    # (슬러지 농축은 24h — capex_selector 에서도 명시적 예외)
    # =========================================================
    el_24h = 0.0
    el_8h = 0.0

    # --- 원료량 기준 (지역 상수 x 이진변수 -> 선형) ---
    el_24h += EL_AIR_CURTAIN_KW * 24.0 * (
        (2 if ctx.W > 0 else 0) + (1 if ctx.S > 0 else 0) + (1 if ctx.F > 0 else 0)
    )
    el_24h += EL_CONVEYOR * (2 * ctx.W + ctx.S + 2 * ctx.F)          # 반입 컨베이어
    el_24h += EL_CONVEYOR * (2 * ctx.W + 2 * ctx.S + 2 * ctx.F)      # 전처리 이송 컨베이어

    # 슬러지 농축 (24h)
    el_24h += EL_SL_GRAVITY * ctx.S * model.s1
    el_24h += EL_SL_ROTARY_DRUM * ctx.S * model.s2

    # 가축분뇨 전처리 (8h): y1=1 이면 screen+cyclone+drum filter, 아니면 screen+cyclone
    el_8h += (EL_AM_SCREEN + EL_AM_CYCLONE) * ctx.W
    el_8h += EL_AM_DRUM_FILTER * ctx.W * model.y1

    # 음식물 전처리 (8h): screen 1/2, sorting 1/2, shredder 1/2, screen3
    for line in model.FW_LINE:
        for sc in model.FW_SCREEN:
            el_8h += EL_FW_SCREEN[sc] * (ctx.F / 2.0) * model.fw_line_screen_sel[line, sc]
        el_8h += EL_FW_SORTING * (ctx.F / 2.0)
    for sc in model.FW_SCREEN:
        el_8h += EL_FW_SCREEN[sc] * ctx.F * model.fw_screen3_sel[sc]
    for stage in model.FW_SHREDDER_STAGE:
        for h in model.FW_SHREDDER:
            el_8h += EL_FW_SHREDDER[h] * ctx.F * model.fw_shredder_sel[stage, h]

    # --- 유량 변수 기준 (전부 24h) ---
    el_24h += EL_BM_MIXER * model.Q_total * 8.0 / 8.0     # 저장조 교반
    el_24h += EL_CONVEYOR * model.Q_total * 2.0
    el_24h += EL_AD * model.Q_ad_in
    el_24h += EL_GAS_COMPRESSOR * model.pur_in_q_gas
    el_24h += EL_CONVEYOR * (model.digestate_flow + model.cake_flow_m3_d)
    # 제습 기술별 (24h)
    for dtech in model.DEHUM:
        v = _gate(
            f"el_dehum_{dtech}", model.dh[dtech],
            EL_DEHUM[dtech] * model.dehum_in_q_gas,
            EL_DEHUM[dtech] * ctx.QGAS_UB,
        )
        el_24h += v
    # 탈수기 기술별 (24h)
    for dtech in model.DIGESTATE_DEWATERING:
        rate = EL_BELT_PRESS if dtech == "belt_press" else EL_SCREW_PRESS
        v = _gate(
            f"el_dewater_{dtech}", model.dd[dtech],
            rate * model.digestate_flow,
            rate * ctx.DIGESTATE_FLOW_UB,
        )
        el_24h += v
    # 폐수처리 옵션별 (24h)
    # [v82-C2] DAF + UV 는 세 공법 모두에 들어가는 필수설비라 게이팅 없이 더한다.
    #   v81 은 Anammox 를 통째로 건너뛰어 이 0.15 kWh/m3 를 0 으로 계산했다.
    el_24h += EL_WWTP_COMMON * model.wastewater_in_flow
    for opt in model.WWTP_OPTION:
        if EL_WWTP.get(opt, 0.0) <= 0.0:
            continue                     # Anammox: 반응조 전기는 고정비로 별도
        v = _gate(
            f"el_wwtp_{opt}", model.wwtp_select[opt],
            EL_WWTP[opt] * model.wastewater_in_flow,
            EL_WWTP[opt] * ctx.WASTEWATER_FLOW_UB,
        )
        el_24h += v

    model.elec_kwh_d_24h = pyo.Var(domain=pyo.NonNegativeReals)
    model.elec_kwh_d_8h = pyo.Var(domain=pyo.NonNegativeReals)
    model.elec_kwh_d_24h_eq = pyo.Constraint(expr=model.elec_kwh_d_24h == el_24h)
    model.elec_kwh_d_8h_eq = pyo.Constraint(expr=model.elec_kwh_d_8h == el_8h)

    model.elec_kwh_d = pyo.Var(domain=pyo.NonNegativeReals)
    model.elec_kwh_d_eq = pyo.Constraint(
        expr=model.elec_kwh_d == model.elec_kwh_d_24h + model.elec_kwh_d_8h
    )

    # 피크전력 [kW]. 8h 설비는 주간에 24h 설비와 동시에 돌므로 합이 피크다.
    model.elec_peak_kw = pyo.Var(domain=pyo.NonNegativeReals)
    model.elec_peak_eq = pyo.Constraint(
        expr=model.elec_peak_kw
        == model.elec_kwh_d_24h / 24.0 + model.elec_kwh_d_8h / 8.0
    )

    # 계약전력: 고정 or 피크 연동 (ELEC_CONTRACT_MODE)
    model.elec_contract_kw = pyo.Var(domain=pyo.NonNegativeReals)
    if ELEC_CONTRACT_MODE == "peak":
        model.elec_contract_cons = pyo.Constraint(
            expr=model.elec_contract_kw >= model.elec_peak_kw
        )
    else:
        model.elec_contract_cons = pyo.Constraint(
            expr=model.elec_contract_kw == ELEC_CONTRACT_KW
        )

    # [v75] 연간 전기요금 = 기본요금 + 전력량요금(TOU) + 기후환경 + 연료비조정
    #  전력량요금은 가동시간대별 계수를 곱해 선형으로 계산한다(위 _tou_coefficients).
    model.elec_cost_usd_y = pyo.Var(domain=pyo.NonNegativeReals)
    # [v82-C5] 가동률 F_UTIL 은 '사용량' 항에만 곱한다. 기본요금은 계약전력
    #   기준이라 가동률과 무관하다. v81 은 약품비(DAYS=365*F_UTIL)에만 곱하고
    #   전기에는 곱하지 않아 두 항의 기준이 어긋나 있었다. F_UTIL=1.0 이면
    #   v81 과 값이 같다.
    model.elec_cost_eq = pyo.Constraint(
        expr=model.elec_cost_usd_y * EXCHANGE_RATE_KRW_PER_USD
        == ELEC_BASE_KRW_PER_KW_MONTH * model.elec_contract_kw * 12.0
        + LCOB_F_UTIL * (
            model.elec_kwh_d_24h * TOU_KRW_PER_KWHD_Y_24H
            + model.elec_kwh_d_8h * TOU_KRW_PER_KWHD_Y_8H
            + model.elec_kwh_d * 365.0 * TOU_SURCHARGE_KRW_PER_KWH
        )
    )

    # --- 변동 OPEX (약품·소모품) ---
    opex_var_d = 0.0
    for dtech in model.DESULF:
        v = _gate(
            f"op_desulf_{dtech}", model.ds[dtech],
            OPEX_DESULF_USD_PER_M3[dtech] * model.desulf_in_q_gas,
            OPEX_DESULF_USD_PER_M3[dtech] * ctx.QGAS_UB,
        )
        opex_var_d += v
    # 정제는 이미 기술별로 게이팅된 up_ch4_selected 를 그대로 쓸 수 있다.
    for tech in model.TECH:
        opex_var_d += (
            OPEX_PURIFICATION_USD_PER_M3_CH4[tech] * model.up_ch4_selected[tech]
        )
    # [v57] 소화조 가온은 자가소비이므로 연료 '구매비'가 없다.
    #  대신 후단으로 가는 가스가 줄어드는 형태로 이미 반영돼 있다.

    # =========================================================
    # [v77] 폐수처리 약품비 — 메탄올·소석회(질소부하 기반) + Alum
    #
    #   메탄올·소석회는 A2O/Bardenpho 에만 든다(Anammox 는 불필요).
    #   폐수유량 x 이진선택변수는 곱이므로 _gate 로 big-M 선형화한다.
    #   Alum 은 공법 무관 공통이라 게이팅 없이 그냥 더한다.
    # =========================================================
    _ww_ub = max(float(ctx.WASTEWATER_FLOW_UB), 1.0)
    for opt in WWTP_OPTIONS_NEEDING_N_CHEM:
        if opt not in model.wwtp_select:
            continue
        # [v82-C3] 공법별 무산소조 단수를 반영한 계수
        _rate = WWTP_CHEM_N_USD_PER_M3[opt]
        opex_var_d += _gate(
            f"op_wwtp_nchem_{opt}",
            model.wwtp_select[opt],
            _rate * model.wastewater_in_flow,
            _rate * _ww_ub,
        )
    # [v83] Anammox 변동 OPEX — 질소부하 기반 'etc'(508 원/kg-N). 전기·황 제외.
    #   폐수유량 x 이진선택변수 곱이므로 _gate 로 big-M 선형화한다.
    if "Anammox" in model.wwtp_select:
        opex_var_d += _gate(
            "op_wwtp_anammox_etc",
            model.wwtp_select["Anammox"],
            WWTP_ANAMMOX_ETC_USD_PER_M3 * model.wastewater_in_flow,
            WWTP_ANAMMOX_ETC_USD_PER_M3 * _ww_ub,
        )
    opex_var_d += WWTP_CHEM_ALUM_USD_PER_M3 * model.wastewater_in_flow

    model.opex_var_usd_d = pyo.Var(domain=pyo.NonNegativeReals)
    model.opex_var_eq = pyo.Constraint(expr=model.opex_var_usd_d == opex_var_d)

    # [v83] Anammox 고정비 제거 — 전기(137,340,000)·황(8,760,000 KRW/y) 대신
    #   질소부하 기반 'etc'(508 원/kg-N) 변동비로 대체했다(위 opex_var_d 참고).
    #   따라서 고정비 항은 0.
    anammox_fixed_usd_y = 0.0

    # =========================================================
    # [v76] C_agg = 구입기기비(model.Cost) + 표준사업비 부대비용(상수)
    #
    # capex_selector.py 971행과 같은 정의다. 부대비용은 총 원료량만의 함수라
    # 설계 선택과 무관한 상수이므로, 선형 제약 한 줄로 끝난다.
    #
    # 주의: 부대비용은 모든 설계에 동일하게 더해지는 상수라 TAC '순위'는
    #  바뀌지 않는다. 그러나 LCOE = TAC / E 는 분수라 분모(공급에너지)가 큰
    #  설계가 유리해진다. 태백시 실측에서 LCOE 최소가 wet_scrubber 에서
    #  psa 로 뒤집혔다. 즉 v75 의 '습식스크러버 최적'은 부대비용 누락이
    #  만든 인공물이었다.
    # =========================================================
    model.indirect_cost_usd = pyo.Param(
        initialize=float(ctx.INDIRECT_USD), mutable=False
    )
    model.C_agg_usd = pyo.Var(domain=pyo.NonNegativeReals)
    model.C_agg_eq = pyo.Constraint(
        expr=model.C_agg_usd == model.Cost + model.indirect_cost_usd
    )

    # [v79b] 국고보조금: 표준사업비×60% 를 자본비에서만 차감 (capex_selector.py 기준)
    #   부대비율(0.771) > 보조율(0.60) 이라 C_agg > 보조금 이 항상 성립 → 등식으로 고정.
    model.subsidy_usd = pyo.Param(
        initialize=float(subsidy_usd(ctx.TOTAL_TON_D)), mutable=False
    )
    model.C_agg_net_usd = pyo.Var(domain=pyo.NonNegativeReals)
    model.C_agg_net_eq = pyo.Constraint(
        expr=model.C_agg_net_usd == model.C_agg_usd - model.subsidy_usd
    )
    # [v79b] 폐기물 반입료 수익 [USD/y] (설계 무관 상수)
    model.tipping_usd_y = pyo.Param(
        initialize=float(tipping_revenue_usd_y(ctx.W, ctx.S, ctx.F)), mutable=False
    )

    model.opex_annual_usd_y = pyo.Var(domain=pyo.NonNegativeReals)
    model.opex_annual_eq = pyo.Constraint(
        expr=model.opex_annual_usd_y
        == LCOB_F_MLC * model.C_agg_usd   # [v76] Cost -> C_agg
        + LCOB_OPEX_FIXED_USD_Y
        + model.elec_cost_usd_y          # [v75] TOU 전기요금 (기본요금 포함)
        + model.opex_var_usd_d * ctx.DAYS
        + anammox_fixed_usd_y * model.wwtp_select["Anammox"]
    )

    # =========================================================
    # [v75] TAC (Total Annualized Cost) = 연환산 CAPEX + 연간 OPEX
    #
    # v74 는 epsilon 을 CAPEX(model.Cost) 에만 걸었다. 그런데 태백시 실측에서
    # 설계별 CAPEX 변동은 +10.3% 인 반면 OPEX 변동은 +55.9% 였다.
    # 즉 '변동이 작은 쪽을 묶고 큰 쪽을 풀어놓는' 제약이었고, 정제기술처럼
    # CAPEX 는 싸지만 약품비가 비싼 선택지가 공짜로 통과했다.
    # v75 는 총연간비용을 epsilon 축으로 삼는다.
    # [v76] 분모가 아니라 분자 기준을 고쳤다. CAPEX 는 C_agg 로 환산한다.
    # =========================================================
    model.TAC_usd_y = pyo.Var(domain=pyo.NonNegativeReals)
    model.TAC_eq = pyo.Constraint(
        expr=model.TAC_usd_y
        == LCOB_F_INV * LCOB_CRF * model.C_agg_net_usd   # [v79b] 보조금 반영: 자본비는 순 CAPEX
        + model.opex_annual_usd_y
    )
    # [v79b] 반입료 차감 순 TAC (리포트용; 상수 오프셋이라 설계선택은 불변)
    model.TAC_net_usd_y = pyo.Var(domain=pyo.Reals)
    model.TAC_net_eq = pyo.Constraint(
        expr=model.TAC_net_usd_y == model.TAC_usd_y - model.tipping_usd_y
    )

    # =========================================================
    # [v54] LCOB
    #   분자 = CAPEX*f_inv*CRF + OPEX_annual - 전기판매수익(부산물 크레딧)
    #   분모 = up_ch4 * 365 * f_util
    # 분수형이라 Dinkelbach 로 푼다: min (분자 - lambda * 분모)
    # =========================================================
    model.lcob_numerator_usd_y = pyo.Var(domain=pyo.Reals)
    model.lcob_denominator_m3_y = pyo.Var(domain=pyo.NonNegativeReals)

    # 택1 구조에서는 부산물 크레딧이 의미가 없다(한 쪽을 고르면 다른 쪽 생산은 0).
    # [v75] 분자는 TAC 와 동일하다. 정의를 한 곳에 두기 위해 TAC 를 참조한다.
    model.lcob_numerator_eq = pyo.Constraint(
        expr=model.lcob_numerator_usd_y == model.TAC_usd_y
    )
    # 정제 시나리오 분모: 정제메탄 [m3/yr]
    model.lcob_denominator_eq = pyo.Constraint(
        expr=model.lcob_denominator_m3_y == model.up_ch4 * ctx.DAYS
    )
    # 발전 시나리오 분모: 발전량 [kWh/yr]
    model.lcoe_denominator_kwh_y = pyo.Var(domain=pyo.NonNegativeReals)
    model.lcoe_denominator_eq = pyo.Constraint(
        expr=model.lcoe_denominator_kwh_y == model.gen_power_kW * 24.0 * ctx.DAYS
    )

    model.lcob_lambda = pyo.Param(
        initialize=float(ctx.lcob_lambda or 0.0), mutable=True
    )
    # =========================================================
    # [v66] route 중립 성능지표 = '판매 에너지' [kWh/d]
    #
    #  LCOB(정제, $/m3)와 LCOE(발전, $/kWh)는 단위가 달라 하나의 목적함수로
    #  못 쓴다. 그래서 v63~v65 는 route 를 강제하고 따로 풀었고,
    #  그 결과 모델이 정제/발전을 '고르지' 못했다.
    #
    #  균등화 생산원가를 '에너지 단위'로 일반화하면 두 경로가 같은 분모를 갖는다.
    #      정제 -> up_ch4 [m3/d] x 9.94 [kWh/m3]
    #      발전 -> P [kW] x 24 [h/d]
    #  route 는 택1이라 둘 중 하나는 0 이고 분모는 항상 정의된다.
    #  덕분에 route_sel 을 자유 이진변수로 되돌릴 수 있다.
    # =========================================================
    model.delivered_energy_kwh_d = pyo.Var(domain=pyo.NonNegativeReals)
    model.delivered_energy_eq = pyo.Constraint(
        expr=model.delivered_energy_kwh_d
        == model.up_ch4 * CH4_LHV_KWH + model.gen_power_kW * 24.0
    )
    model.delivered_energy_kwh_y = pyo.Var(domain=pyo.NonNegativeReals)
    model.delivered_energy_y_eq = pyo.Constraint(
        expr=model.delivered_energy_kwh_y
        == model.delivered_energy_kwh_d * ctx.DAYS
    )
    model.levelized_denominator = pyo.Expression(
        expr=model.delivered_energy_kwh_y
    )
    #  분모(판매 에너지)가 0 이면 균등화원가가 정의되지 않는다.
    model.min_output_cons = pyo.Constraint(
        expr=model.delivered_energy_kwh_d >= MIN_DELIVERED_ENERGY_KWH_D
    )

    # =========================================================
    # [v62] Pareto 생성 목적함수
    #   목적: upgraded CH4 생산량 최대화
    #   제약: CAPEX(model.Cost) <= epsilon
    #
    # LCOB는 목적함수로 사용하지 않는다. 각 Pareto 해를 얻은 뒤
    # 연간화 비용 / 연간 CH4 생산량으로 사후 평가한다.
    # =========================================================
    # =========================================================
    # [v67] Pareto 생성 목적 = 일일 판매수익 [$/d]  (route 무관)
    #
    #   total_revenue = up_ch4 x 도시가스단가 + P x 24 x 전기단가
    #
    #  v66 은 '판매 에너지 [kWh/d]' 를 썼는데, 그러면 route 선택이
    #  "어느 경로가 최종 에너지를 더 많이 보존하는가" 가 된다.
    #  메탄 1 m3 기준 정제 9.94 kWh vs 발전 3.98 kWh (효율 40%) 라서
    #  단가가 어떻게 바뀌든 정제가 구조적으로 이긴다. 즉 route 를 자유변수로
    #  풀어놓고도 사실상 정제를 강제하는 셈이었다.
    #
    #  수익 기준이면 route 선택이 경제 문제가 되고 단가에 반응한다.
    #    현재 단가:   정제 0.617 vs 발전 0.398 $/m3CH4  -> 정제 우세
    #    전기 2배면:  정제 0.617 vs 발전 0.795          -> 발전 우세
    # =========================================================
    # =========================================================
    # [v75] 다목적 정식화 —  min TAC  /  max 유효공급에너지
    #
    #   목적:  max  E_useful [kWh/d]
    #   제약:  TAC <= epsilon   [USD/y]
    #
    # epsilon 을 훑으면 (TAC, E) Pareto front 가 나오고, 각 점의
    #   LCOE = TAC / (E x 365)   [USD/kWh]
    # 를 계산해 최소가 되는 점이 '가성비 최적'이다.
    #
    # 왜 LCOE 를 직접 최소화하지 않는가:
    #   (a) LCOE 최소 설계는 반드시 (TAC 최소, E 최대) Pareto front 위에 있다.
    #       [증명] 설계 X 가 front 밖이면 X 를 지배하는 front 점 P 가 있어
    #       TAC_P <= TAC_X, E_P >= E_X 이므로 LCOE_P <= LCOE_X. 따라서 front 를
    #       모두 훑으면 LCOE 최소점은 그 안에 반드시 포함된다.
    #   (b) 분수형 목적은 Dinkelbach 반복이 필요해 점당 solve 가 3~6 배로 늘고,
    #       epsilon 이 커지면 LCOE 가 상수로 평평해져 front 상단 정보가 사라진다.
    #   (c) 선형 목적이면 기존 2단계 solve/단조성 제약을 그대로 쓸 수 있다.
    #
    # 유효공급에너지 E_useful:
    #   수요가 주어지면 min(생산, 수요) 를 쓴다. 수요를 넘는 생산은 팔 곳이
    #   없으므로 분모에 넣지 않는다. 이렇게 해야 도시가스 배관이 없는 지역이
    #   정제설비를 짓고 가지도 못하는 설계가 걸러진다(v73 대체편익과 같은 취지).
    #   수요가 없으면 총생산량을 그대로 쓴다.
    # =========================================================
    _use_demand = (ctx.gas_demand_kwh_d is not None) or (ctx.elec_demand_kwh_d is not None)
    model.objective_uses_demand = pyo.Param(
        initialize=int(_use_demand), mutable=False
    )

    # gas_displaced / elec_displaced 는 수요가 없으면 상한이 '생산 상한'으로
    # 잡히므로(위 _gas_dem/_elec_dem 참조), 목적함수가 이들을 밀어올리면
    # 자동으로 min(생산, 수요) = 생산 이 된다. 따라서 분기가 필요 없다.
    model.useful_energy_kwh_d = pyo.Var(domain=pyo.NonNegativeReals)
    model.useful_energy_eq = pyo.Constraint(
        expr=model.useful_energy_kwh_d
        == model.gas_displaced_kwh_d + model.elec_displaced_kwh_d
    )
    model.useful_energy_kwh_y = pyo.Var(domain=pyo.NonNegativeReals)
    model.useful_energy_y_eq = pyo.Constraint(
        expr=model.useful_energy_kwh_y == model.useful_energy_kwh_d * ctx.DAYS
    )

    # 균등화 원가의 분모는 '유효공급에너지'로 통일한다.
    model.levelized_denominator_v75 = pyo.Expression(
        expr=model.useful_energy_kwh_y
    )

    model.objective_expr = pyo.Expression(expr=model.useful_energy_kwh_d)
    model.obj = pyo.Objective(expr=model.objective_expr, sense=pyo.maximize)

    # epsilon 제약: 총연간비용 상한
    model.tac_cap = pyo.Constraint(expr=model.TAC_usd_y <= ctx.epsilon)

    # [Fix 4] 단조성 강제.
    #  epsilon 이 커지면 실행가능영역은 포함관계로 커지므로 최적 목적값은 절대
    #  감소할 수 없다. 그런데 v48 결과에서는 192행 중 20행이 이를 위반했고
    #  두 stage 모두 'optimal' 을 보고했다.
    #  이미 확보한 목적값을 하한으로 명시해 solver 가 그보다 나쁜 해를 반환하지 못하게 한다.
    if ctx.obj_floor is not None:
        model.objective_floor = pyo.Constraint(
            expr=model.objective_expr
            >= float(ctx.obj_floor) - max(1.0e-6, 1.0e-6 * abs(float(ctx.obj_floor)))
        )

    # [v75] LCOE 최소화 전용 모드 (sweep 끝에서 최적점을 정확히 집을 때만 사용).
    #  Dinkelbach:  min (TAC - lambda x E_y).  최적값이 0 이 되는 lambda 가 LCOE 최소값.
    if ctx.lcob_lambda is not None and float(ctx.lcob_lambda) > 0.0:
        model.obj.deactivate()
        model.dinkelbach_obj = pyo.Objective(
            expr=model.TAC_usd_y
            - float(ctx.lcob_lambda) * model.useful_energy_kwh_y,
            sense=pyo.minimize,
        )


# -----------------------------------------------------------------------
# 조립: 위 스테이지를 공정 순서대로 호출
# -----------------------------------------------------------------------
def build_model(epsilon, W, F, S, probe_mode=False, grid_hint=None, obj_floor=None,
                gas_demand_kwh_d=None, elec_demand_kwh_d=None, lcob_lambda=0.0,
                force_route=None):
    """공정 순서대로 스테이지를 조립한다. 상세는 각 _stage_* 함수 참조."""
    model = pyo.ConcreteModel("Animal manure flow")
    ctx = SimpleNamespace(
        epsilon=epsilon,
        W=W,
        F=F,
        S=S,
        probe_mode=probe_mode,
        grid_hint=grid_hint,
        obj_floor=obj_floor,
        gas_demand_kwh_d=gas_demand_kwh_d,
        elec_demand_kwh_d=elec_demand_kwh_d,
        lcob_lambda=lcob_lambda,
        force_route=force_route,
    )
    _stage_setup(model, ctx)
    _stage_receiving(model, ctx)
    _stage_pretreatment(model, ctx)
    _stage_digestion(model, ctx)
    _stage_gas_setup(model, ctx)
    _stage_gas_balance(model, ctx)
    _stage_gas_cleanup(model, ctx)
    _stage_upgrading(model, ctx)
    _stage_upgrading_perf(model, ctx)
    _stage_power(model, ctx)
    _stage_digestate(model, ctx)
    _stage_wwtp(model, ctx)
    _stage_capex(model, ctx)
    _stage_objective(model, ctx)

    return model


# =========================================================
# 3. solve helper
# =========================================================
def solve_model(
    model,
    *,
    retry=False,
    time_limit=None,
    mip_rel_gap=None,
):
    # [Fix 4] 기본 gap/시간을 전역설정으로 통일 (v48: 1e-3 / 300s)
    time_limit = SOLVER_TIME_LIMIT if time_limit is None else time_limit
    mip_rel_gap = SOLVER_MIP_REL_GAP if mip_rel_gap is None else mip_rel_gap
    solver = pyo.SolverFactory("appsi_highs")

    # 병렬 실행 시 스레드 과다구독을 막기 위해 환경변수로 HiGHS 스레드 수를 제한한다.
    # (미설정이면 기존 동작 그대로 HiGHS 기본값 사용)
    _highs_threads = os.environ.get("HIGHS_THREADS")
    if _highs_threads:
        solver.options["threads"] = int(_highs_threads)

    solver.options["presolve"] = "off" if retry else "on"
    solver.options["time_limit"] = float(time_limit)
    solver.options["mip_rel_gap"] = float(mip_rel_gap)
    solver.options["mip_abs_gap"] = float(SOLVER_MIP_ABS_GAP)
    solver.options["mip_feasibility_tolerance"] = 1.0e-6
    solver.options["primal_feasibility_tolerance"] = 1.0e-7
    solver.options["dual_feasibility_tolerance"] = 1.0e-7

    result = solver.solve(model, load_solutions=False)
    term = str(result.solver.termination_condition)

    if term.lower() in ("optimal", "feasible"):
        model.solutions.load_from(result)
        return result, term

    # [Fix 4-c] 시간/반복 한계로 멈췄더라도 incumbent 가 있으면 그것을 쓴다.
    #  v48은 이 경우를 실패로 처리해 presolve 를 끄고 처음부터 다시 풀었고
    #  (같은 시간 한계를 한 번 더 소모), 그래도 실패하면 그 epsilon 을 통째로 버렸다.
    if "limit" in term.lower() or "max" in term.lower():
        try:
            model.solutions.load_from(result)
            return result, "feasible"
        except Exception:
            pass

    return result, term


def solve_model_with_retry(
    model,
    *,
    time_limit=None,
    mip_rel_gap=None,
):
    result, term = solve_model(
        model,
        retry=False,
        time_limit=time_limit,
        mip_rel_gap=mip_rel_gap,
    )

    if term.lower() in ("optimal", "feasible"):
        return result, term

    result_retry, term_retry = solve_model(
        model,
        retry=True,
        time_limit=time_limit,
        mip_rel_gap=mip_rel_gap,
    )

    return result_retry, term_retry


def solve_maxobj_then_mincost(model):
    """1단계: 유효공급에너지 최대화 / 2단계: 같은 공급량을 유지하며 TAC 최소화.

    [v75] 2단계 목적을 CAPEX(model.Cost)에서 TAC 로 바꿨다.
      epsilon 축이 TAC 이므로 2단계도 같은 축에서 눌러야 Pareto front 위의
      점이 된다. CAPEX 만 줄이면 OPEX 가 늘어 TAC 가 오히려 커지는 해가
      front 점으로 기록될 수 있다(정제기술 선택이 정확히 그런 경우다).
    """
    _, term1 = solve_model_with_retry(model)
    model.stage1_termination = term1

    if term1.lower() not in ("optimal", "feasible"):
        model.stage2_termination = "not_run"
        return term1

    m_star = pyo.value(model.objective_expr)

    model.obj.deactivate()
    # up_ch4 는 0 이상 스칼라라 상대오차가 적절하다.
    model.lock_obj = pyo.Constraint(
        expr=model.objective_expr
        >= m_star * (1.0 - 1.0e-4) - 1.0e-6
    )
    model.obj_cost = pyo.Objective(
        expr=model.TAC_usd_y,          # [v75] CAPEX 가 아니라 TAC 를 최소화
        sense=pyo.minimize,
    )

    _, term2 = solve_model_with_retry(model)
    model.stage2_termination = term2

    if term2.lower() in ("optimal", "feasible"):
        return term2

    return "stage1_only"


# =========================================================
# 3-b. 지역별 적응형 epsilon 범위 산정
# =========================================================
PROBE_TIME_LIMIT = 180.0
PROBE_MIP_GAP = 0.0


def _build_probe_model(W, F, S, grid_hint=None, probe_mode=True, demand=None,
                       force_route=None):
    # 비용상한을 비활성화한 상태로 비용범위를 탐색한다.
    demand = demand or {}
    m = build_model(
        epsilon=1.0e12,
        W=W,
        F=F,
        S=S,
        probe_mode=probe_mode,
        grid_hint=grid_hint,
        gas_demand_kwh_d=demand.get("gas"),
        elec_demand_kwh_d=demand.get("elec"),
        force_route=force_route,
    )

    # [v75] epsilon 제약은 TAC 에 걸려 있다.
    if hasattr(m, "tac_cap"):
        m.tac_cap.deactivate()

    return m


def _probe_min_cost(W, F, S, ch4_floor=None, grid_hint=None, probe_mode=True, demand=None,
                    force_route=None):
    m = _build_probe_model(W, F, S, grid_hint=grid_hint, probe_mode=probe_mode, demand=demand,
                           force_route=force_route)
    m.obj.deactivate()

    if ch4_floor is not None:
        m.probe_ch4_floor = pyo.Constraint(
            expr=m.up_ch4 >= float(ch4_floor)
        )

    # [v75] epsilon 축이 TAC 이므로 최소비용도 TAC 기준으로 찾는다.
    m.probe_min_cost_obj = pyo.Objective(
        expr=m.TAC_usd_y,
        sense=pyo.minimize,
    )

    _, term = solve_model_with_retry(
        m,
        time_limit=PROBE_TIME_LIMIT,
        mip_rel_gap=PROBE_MIP_GAP,
    )

    if term.lower() not in ("optimal", "feasible"):
        print(
            f"    [probe min-cost 실패] termination={term}, "
            f"W={W:.6f}, F={F:.6f}, S={S:.6f}, "
            f"ch4_floor={ch4_floor}"
        )
        return None

    return {
        "cost": float(pyo.value(m.Cost)),                 # 장비 CAPEX [USD]
        "tac": float(pyo.value(m.TAC_usd_y)),             # [v75] 총연간비용 [USD/y]
        "useful_kwh_d": float(pyo.value(m.useful_energy_kwh_d)),
        "ch4": float(pyo.value(m.up_ch4)),
        "q_gas": float(pyo.value(m.pred_q_gas_var)),
        "ch4_frac": float(pyo.value(m.pred_ch4_frac_dry_var)),
        "objective": float(pyo.value(m.objective_expr)),
        "termination": term,
    }


def _probe_max_ch4(W, F, S, grid_hint=None, probe_mode=True, demand=None,
                   force_route=None):
    m = _build_probe_model(W, F, S, grid_hint=grid_hint, probe_mode=probe_mode, demand=demand,
                           force_route=force_route)
    m.obj.deactivate()

    # [v75] 포화점은 '유효공급에너지 최대' 해로 잡는다(본 sweep 목적함수와 동일).
    m.probe_max_ch4_obj = pyo.Objective(
        expr=m.useful_energy_kwh_d,
        sense=pyo.maximize,
    )

    _, term = solve_model_with_retry(
        m,
        time_limit=PROBE_TIME_LIMIT,
        mip_rel_gap=PROBE_MIP_GAP,
    )

    if term.lower() not in ("optimal", "feasible"):
        print(
            f"    [probe max-CH4 실패] termination={term}, "
            f"W={W:.6f}, F={F:.6f}, S={S:.6f}"
        )
        return None

    return {
        "cost": float(pyo.value(m.Cost)),                 # 장비 CAPEX [USD]
        "tac": float(pyo.value(m.TAC_usd_y)),             # [v75] 총연간비용 [USD/y]
        "useful_kwh_d": float(pyo.value(m.useful_energy_kwh_d)),
        "ch4": float(pyo.value(m.up_ch4)),
        "q_gas": float(pyo.value(m.pred_q_gas_var)),
        "ch4_frac": float(pyo.value(m.pred_ch4_frac_dry_var)),
        "objective": float(pyo.value(m.objective_expr)),
        "termination": term,
    }


def probe_cost_range(W, F, S, demand=None, force_route=None):
    """
    지역별 epsilon 범위를 두 번의 probe solve로 계산한다.

    1) 순수 최소비용 해 → C_min
    2) 최대 upgraded CH4 해 → C_sat

    기존의 세 번째 solve인
    '최대 CH4의 99.9% 이상 유지 + 최소비용' 단계는 제거한다.
    이 단계는 이미 존재하는 최대-CH4 feasible 해를 다시 비용최소화로
    풀면서 HiGHS가 수치적으로 infeasible을 반환하는 경우가 많았고,
    epsilon 상한 산정에는 꼭 필요하지 않다.
    """
    # [Fix 4-b] v48은 probe 를 경량 grid(5점, 전 범위)로 풀고 그렇게 얻은
    #  C_min/C_sat 을 '본 sweep 모델(11점 grid)'의 비용상한으로 그대로 썼다.
    #  두 모델의 이산화가 달라 본 모델의 실제 최소비용이 C_sat 보다 커지는 경우가 있고,
    #  그러면 낮은 epsilon 구간이 통째로 infeasible 이 된다
    #  (v48 출력의 infeasible 2행이 이 현상이며 실행시간도 크게 낭비된다).
    #  v49는 2단계로 probe 한다:
    #    1차) 경량 grid 로 이중선형 격자 힌트만 얻고
    #    2차) 본 sweep 과 '동일한' grid 로 다시 풀어 C_min/C_sat 을 일치시킨다.
    coarse_min = _probe_min_cost(W, F, S, ch4_floor=None, demand=demand, force_route=force_route)
    coarse_max = _probe_max_ch4(W, F, S, demand=demand, force_route=force_route)

    if coarse_max is None:
        return None

    _sols0 = [x for x in (coarse_min, coarse_max) if x is not None]
    _q0 = max([x.get("q_gas", 0.0) or 0.0 for x in _sols0] + [0.0])
    _f0 = [x.get("ch4_frac") for x in _sols0 if x.get("ch4_frac") is not None]

    grid_hint0 = None
    if _q0 > 0.0 and _f0:
        grid_hint0 = {"qgas": _q0, "ch4_lo": min(_f0), "ch4_hi": max(_f0)}

    fine_min = _probe_min_cost(
        W, F, S, ch4_floor=None, grid_hint=grid_hint0, probe_mode=False,
        demand=demand, force_route=force_route,
    )
    fine_max = _probe_max_ch4(
        W, F, S, grid_hint=grid_hint0, probe_mode=False, demand=demand,
        force_route=force_route,
    )

    min_solution = fine_min if fine_min is not None else coarse_min
    max_solution = fine_max if fine_max is not None else coarse_max

    if max_solution is None:
        return None

    ch4_max = max_solution["ch4"]

    # 최대 upgraded CH4 해 자체가 실제 feasible 해이므로
    # 그 비용을 포화비용 C_sat로 직접 사용한다.
    C_sat = float(max_solution["tac"])          # [v75] TAC 기준

    if min_solution is not None:
        C_min = float(min_solution["tac"])      # [v75] TAC 기준
    else:
        C_min = 0.90 * C_sat
        print(
            "    [probe pure min-cost 실패 → "
            f"지역별 추정 하한 사용: {C_min:,.0f} $/y]"
        )

    C_min = max(0.0, C_min)
    C_sat = max(C_sat, C_min)

    # [Fix 2] probe 해에서 이중선형 격자 범위 힌트를 만든다.
    _sols = [x for x in (min_solution, max_solution) if x is not None]
    _qgas = max([x.get("q_gas", 0.0) or 0.0 for x in _sols] + [0.0])
    _fracs = [
        x.get("ch4_frac")
        for x in _sols
        if x.get("ch4_frac") is not None
    ]
    grid_hint = None
    if _qgas > 0.0 and _fracs:
        grid_hint = {
            "qgas": _qgas,
            "ch4_lo": min(_fracs),
            "ch4_hi": max(_fracs),
        }

    print(
        f"    [probe 성공] TAC_min={C_min:,.0f} $/y, "
        f"TAC_sat={C_sat:,.0f} $/y, "
        f"max_useful_energy={max_solution['useful_kwh_d']:,.1f} kWh/d "
        f"(up_ch4={ch4_max:,.1f} m3/d), "
        f"grid_hint={grid_hint}"
    )

    return C_min, C_sat, ch4_max, grid_hint


def get_selected_options(model):
    selected = {
        "y1": pyo.value(model.y1),
        "s1": pyo.value(model.s1),
        "s2": pyo.value(model.s2),
    }

    if hasattr(model, "FW_LINE"):
        for line in model.FW_LINE:
            for s in model.FW_SCREEN:
                selected[f"FW_line{line}_screen_{s}"] = pyo.value(
                    model.fw_line_screen_sel[line, s]
                )

    if hasattr(model, "FW_SHREDDER_STAGE"):
        for stage in model.FW_SHREDDER_STAGE:
            for h in model.FW_SHREDDER:
                selected[f"FW_shredder_stage{stage}_{h}"] = pyo.value(
                    model.fw_shredder_sel[stage, h]
                )

    if hasattr(model, "fw_screen3_sel"):
        for s in model.FW_SCREEN:
            selected[f"FW_screen3_{s}"] = pyo.value(model.fw_screen3_sel[s])

    if hasattr(model, "TECH"):
        for tech in model.TECH:
            selected[f"purification_{tech}"] = pyo.value(model.y[tech])

    if hasattr(model, "DESULF"):
        for d in model.DESULF:
            selected[f"desulfurization_{d}"] = pyo.value(model.ds[d])

    if hasattr(model, "DEHUM"):
        for d in model.DEHUM:
            selected[f"dehumidification_{d}"] = pyo.value(model.dh[d])

    if hasattr(model, "TANK"):
        for tank in model.TANK:
            selected[f"gas_storage_{tank}"] = pyo.value(model.t[tank])

    if hasattr(model, "GEN"):
        for gname in model.GEN:
            selected[f"generator_{gname}"] = pyo.value(model.g[gname])

    if hasattr(model, "DIGESTATE_DEWATERING"):
        for d in model.DIGESTATE_DEWATERING:
            selected[f"digestate_dewatering_{d}"] = pyo.value(
                model.dd[d]
            )

    if hasattr(model, "WWTP_OPTION"):
        for option in model.WWTP_OPTION:
            selected[f"wastewater_treatment_{option}"] = pyo.value(
                model.wwtp_select[option]
            )

    if hasattr(model, "ROUTE"):
        for r in model.ROUTE:
            selected[f"route_{r}"] = pyo.value(model.route_sel[r])

    return selected


def nn_input_diagnostics(model):
    """[Fix 1] NN 입력 11개가 학습범위 안인지 진단한다.

    범위를 벗어난 경우 표준편차 기준 초과폭(sigma)을 함께 낸다.
    surrogate 재학습이 필요한지, 어느 입력이 얼마나 벗어났는지 수치로 남긴다.
    """
    (
        _x_offset,
        x_factor,
        _y_offset,
        _y_factor,
        _slb,
        _sub,
        x_raw_min,
        x_raw_max,
    ) = load_biogas_scaler()

    out = {}
    all_in = 1
    worst = 0.0

    for i, name in enumerate(NN_INPUTS):
        try:
            v = float(pyo.value(model.biogas.inputs[i]))
        except Exception:
            continue

        lo = float(x_raw_min[name])
        hi = float(x_raw_max[name])
        sd = abs(float(x_factor[name])) or 1.0

        in_range = 1 if (lo - 1.0e-9) <= v <= (hi + 1.0e-9) else 0
        if in_range:
            excess = 0.0
        elif v < lo:
            excess = (lo - v) / sd
        else:
            excess = (v - hi) / sd

        out[f"nn_in_{name}"] = v
        out[f"nn_train_min_{name}"] = lo
        out[f"nn_train_max_{name}"] = hi
        out[f"nn_in_range_{name}"] = in_range
        out[f"nn_sigma_excess_{name}"] = excess

        all_in = all_in and in_range
        worst = max(worst, excess)

    out["nn_all_inputs_in_training_range"] = int(all_in)
    out["nn_max_sigma_excess"] = worst
    return out


def feed_cod_summary(model):
    """원료별 COD 보고값 + 물리 메탄수율 검증.

    [v51 추가] 물리 메탄수율 검증
      실측 COD 로 계산한 CH4 수율이 이론 최대(0.35 m3/kgCOD)를 넘는지 본다.
      COD 제거량 <= COD 유입량 이므로 CH4/유입COD 조차 0.35 를 넘으면
      그 해는 열역학적으로 불가능하다. surrogate 가 실측 COD 범위 밖에서
      학습됐다는 사실이 결과 파일 안에서 바로 드러나게 하기 위한 열이다.
    """
    out = {}
    phys = {"ani": COD_ANI_KG_M3, "fw": COD_FW_KG_M3, "sl": COD_SL_KG_M3}
    flow_var = {"ani": model.Q_ani, "fw": model.Q_fw, "sl": model.Q_sl}

    physical_cod_kg_d = 0.0
    for k in ("ani", "fw", "sl"):
        q_raw = float(pyo.value(model.raw_flow_m3_d[k]))
        q_out = float(pyo.value(flow_var[k]))
        out[f"surrogate_COD_{k}"] = (
            q_raw * SURROGATE_COD_COEF / q_out if q_out > 1.0e-9 else 0.0
        )
        out[f"physical_COD_{k}_kg_d"] = q_raw * phys[k]
        physical_cod_kg_d += q_raw * phys[k]

    raw_ch4 = float(pyo.value(model.pur_in_ch4_flow))
    yield_phys = (
        raw_ch4 / physical_cod_kg_d if physical_cod_kg_d > 1.0e-9 else None
    )

    out["physical_COD_total_kg_d"] = physical_cod_kg_d
    out["surrogate_COD_total_kg_d"] = sum(
        float(pyo.value(model.raw_flow_m3_d[k])) * SURROGATE_COD_COEF
        for k in ("ani", "fw", "sl")
    )
    out["theoretical_CH4_yield_m3_per_kgCOD"] = CH4_YIELD_M3_PER_KGCOD
    out["physical_CH4_yield_m3_per_kgCOD"] = yield_phys
    out["physical_CH4_yield_ratio_to_theoretical"] = (
        yield_phys / CH4_YIELD_M3_PER_KGCOD if yield_phys is not None else None
    )
    out["ch4_yield_exceeds_theoretical"] = (
        int(yield_phys > CH4_YIELD_M3_PER_KGCOD)
        if yield_phys is not None
        else None
    )
    out["max_CH4_from_physical_COD_m3_d"] = (
        physical_cod_kg_d * CH4_YIELD_M3_PER_KGCOD
    )
    return out


def capex_summary(model):
    """CAPEX, 수요 대체량·충족률, 단위 생산량당 자본비."""
    cost = float(pyo.value(model.Cost))
    ch4 = float(pyo.value(model.up_ch4))
    biogas = float(pyo.value(model.dehum_out_q_gas))
    feed = float(pyo.value(model.Q_ad_in))

    gas_prod = float(pyo.value(model.gas_product_kwh_d))
    elec_prod = float(pyo.value(model.elec_product_kwh_d))
    gas_disp = float(pyo.value(model.gas_displaced_kwh_d))
    elec_disp = float(pyo.value(model.elec_displaced_kwh_d))
    gas_dem = float(pyo.value(model.gas_demand_kwh_d))
    elec_dem = float(pyo.value(model.elec_demand_kwh_d))
    gas_given = int(pyo.value(model.gas_demand_is_given))
    elec_given = int(pyo.value(model.elec_demand_is_given))

    return {
        "objective_metric": "LCOB_usd_per_m3_CH4_minimized",
        # --- surrogate 신뢰도 진단 (학습데이터 재현오차 기준) ---
        #  NN 은 q_ad > 20 에서 학습데이터를 오차 ±2% 로 재현하지만
        #  q_ad <= 10 구간(학습샘플 5,000중 114개)에서는 평균 -46% 과소예측하고
        #  24%가 음수를 예측한다. 해당 구간 결과는 신뢰할 수 없다.
        "surrogate_reliable": int(
            float(pyo.value(model.Q_ad_in)) >= SURROGATE_MIN_Q_AD
        ),
        "specific_gas_yield": (
            float(pyo.value(model.pred_q_gas_var))
            / (float(pyo.value(model.Q_ad_in)) * float(pyo.value(model.COD_total)))
            if float(pyo.value(model.Q_ad_in)) > 1.0e-9
            else None
        ),
        "reference_specific_yield": SURROGATE_REF_YIELD,
        "gas_if_reference_yield_m3_d": (
            float(pyo.value(model.Q_ad_in))
            * float(pyo.value(model.COD_total))
            * SURROGATE_REF_YIELD
        ),
        "route_note": "purification_and_generation_are_mutually_exclusive",
        "AD_gas_produced_m3_d": float(pyo.value(model.pred_q_gas_var)),
        "AD_heating_self_use_frac": AD_HEATING_CH4_FRAC_OF_GAS,
        "AD_heating_gas_m3_d": float(pyo.value(model.ad_heating_gas_m3_d)),
        "gas_to_downstream_m3_d": float(pyo.value(model.storage_in_q_gas)),
        "LCOB_CRF": LCOB_CRF,
        "LCOB_annualized_capex_usd_y": LCOB_F_INV * LCOB_CRF * cost,
        "LCOB_opex_annual_usd_y": float(pyo.value(model.opex_annual_usd_y)),
        "LCOB_elec_kwh_d": float(pyo.value(model.elec_kwh_d)),
        # [v75] TOU 전기요금 실측치 및 가동시간대별 내역
        "elec_kwh_d_24h": float(pyo.value(model.elec_kwh_d_24h)),
        "elec_kwh_d_8h": float(pyo.value(model.elec_kwh_d_8h)),
        "elec_P24_kW": float(pyo.value(model.elec_kwh_d_24h)) / 24.0,
        "elec_P8_kW": float(pyo.value(model.elec_kwh_d_8h)) / 8.0,
        "LCOB_elec_cost_usd_y": float(pyo.value(model.elec_cost_usd_y)),
        "elec_peak_kW": float(pyo.value(model.elec_peak_kw)),
        "elec_contract_kW": float(pyo.value(model.elec_contract_kw)),
        "elec_base_charge_usd_y": (
            ELEC_BASE_KRW_PER_KW_MONTH * float(pyo.value(model.elec_contract_kw))
            * 12.0 / EXCHANGE_RATE_KRW_PER_USD
        ),
        "elec_effective_price_usd_per_kWh": (
            float(pyo.value(model.elec_cost_usd_y))
            / (float(pyo.value(model.elec_kwh_d)) * 365.0)
            if float(pyo.value(model.elec_kwh_d)) > 1.0e-9 else None
        ),
        # [v76] CAPEX 기준 분해 — 구입기기비 / 부대비용 / 합계
        "equipment_cost_usd": float(pyo.value(model.Cost)),
        "indirect_cost_usd": float(pyo.value(model.indirect_cost_usd)),
        "C_agg_usd": float(pyo.value(model.C_agg_usd)),
        "indirect_cost_included": int(INCLUDE_INDIRECT_COST),
        "annualized_capex_usd_y": float(
            LCOB_F_INV * LCOB_CRF * pyo.value(model.C_agg_usd)
        ),
        # [v75] 총연간비용과 유효공급에너지 — epsilon 축과 목적함수 그 자체
        "TAC_usd_y": float(pyo.value(model.TAC_usd_y)),
        # [v79b] 국고보조금·반입료 반영 결과
        "subsidy_usd": float(pyo.value(model.subsidy_usd)),
        "C_agg_net_usd": float(pyo.value(model.C_agg_net_usd)),
        "annualized_capex_net_usd_y": float(
            LCOB_F_INV * LCOB_CRF * pyo.value(model.C_agg_net_usd)
        ),
        "tipping_revenue_usd_y": float(pyo.value(model.tipping_usd_y)),
        "TAC_net_usd_y": float(pyo.value(model.TAC_net_usd_y)),
        "LCOE_net_usd_per_kWh": (
            float(pyo.value(model.TAC_net_usd_y))
            / float(pyo.value(model.useful_energy_kwh_y))
            if float(pyo.value(model.useful_energy_kwh_y)) > 1.0e-9 else None
        ),
        "useful_energy_kWh_d": float(pyo.value(model.useful_energy_kwh_d)),
        "useful_energy_kWh_y": float(pyo.value(model.useful_energy_kwh_y)),
        "LCOB_opex_var_usd_d": float(pyo.value(model.opex_var_usd_d)),
        # [v77] 폐수처리 약품비 내역 [USD/y]
        "wwtp_TN_mg_L": WWTP_TN_MG_L,
        "wwtp_N_load_kgN_d": (
            WWTP_TN_MG_L * float(pyo.value(model.wastewater_in_flow)) * 1.0e-3
        ),
        "wwtp_chem_methanol_lime_usd_y": sum(
            float(pyo.value(getattr(model, f"op_wwtp_nchem_{o}")))
            for o in WWTP_OPTIONS_NEEDING_N_CHEM
            if hasattr(model, f"op_wwtp_nchem_{o}")
        ) * 365.0,
        "wwtp_chem_alum_usd_y": (
            WWTP_CHEM_ALUM_USD_PER_M3
            * float(pyo.value(model.wastewater_in_flow)) * 365.0
        ),
        "LCOB_elec_credit_usd_y": float(
            pyo.value(model.electricity_revenue)
        ) * 365.0 * LCOB_F_UTIL,
        "LCOB_numerator_usd_y": float(pyo.value(model.lcob_numerator_usd_y)),
        "LCOB_denominator_m3_y": float(pyo.value(model.lcob_denominator_m3_y)),
        "displacement_benefit_objective_value": float(
            pyo.value(model.displacement_benefit_usd_d)
        ),
        # 주의: 후처리(dominated 판정·단조성 검증)에 쓰이는 objective_value 는
        #  solve_case() 에서 -LCOB(또는 -LCOE)로 세팅한다. 여기서 같은 키를
        #  쓰면 세팅 순서에 의존하는 버그가 되므로 다른 이름을 쓴다.
        "displacement_benefit_usd_d_value": float(
            pyo.value(model.displacement_benefit_usd_d)
        ),
        "CAPEX_usd": cost,

        # --- 생산량 ---
        "city_gas_production_m3CH4_d": ch4,
        "city_gas_production_kWh_d": gas_prod,
        "electricity_production_kWh_d": elec_prod,
        "biogas_production_m3_d": biogas,

        # --- 지역 수요 및 충족률 ---
        "gas_demand_given": gas_given,
        "elec_demand_given": elec_given,
        "gas_demand_kWh_d": (gas_dem if gas_given else None),
        "elec_demand_kWh_d": (elec_dem if elec_given else None),
        "gas_displaced_kWh_d": gas_disp,
        "elec_displaced_kWh_d": elec_disp,
        "gas_demand_coverage_frac": (
            gas_disp / gas_dem if (gas_given and gas_dem > 1.0e-9) else None
        ),
        "elec_demand_coverage_frac": (
            elec_disp / elec_dem if (elec_given and elec_dem > 1.0e-9) else None
        ),
        "gas_surplus_kWh_d": max(0.0, gas_prod - gas_disp),
        "elec_surplus_kWh_d": max(0.0, elec_prod - elec_disp),

        # --- 편익 (운영비 차감 전) ---
        "displacement_benefit_usd_d": float(
            pyo.value(model.displacement_benefit_usd_d)
        ),
        "benefit_note": "operating_cost_not_deducted",

        # --- 단위 자본비 (단가 가정 불필요) ---
        "capex_per_m3d_CH4": (cost / ch4 if ch4 > 1.0e-9 else None),
        "capex_per_m3d_biogas": (cost / biogas if biogas > 1.0e-9 else None),
        "capex_per_m3d_feed": (cost / feed if feed > 1.0e-9 else None),
        "capex_per_usd_d_benefit": (
            cost / float(pyo.value(model.displacement_benefit_usd_d))
            if float(pyo.value(model.displacement_benefit_usd_d)) > 1.0e-9
            else None
        ),
        "revenue_reference_only_usd_d": float(
            pyo.value(model.total_revenue)
        ),
    }


def solids_balance_check(model, W, F, S):
    """[검증] AD 고형물 수지 잔차.

    v51 부터는 소화액 고형물을 VS 분해량으로부터 계산하므로
        유입 건고형물 - VS 분해량 - 소화액 건고형물 = 0
    이 항등적으로 성립해야 한다. 이 열이 0 이 아니면 버그다.
    (v48~v50 은 소화액 TS 를 5% 로 가정해 최대 -4,260 kg/d 어긋났다)
    """
    feed = float(pyo.value(model.ad_feed_dry_solids_kg_d))
    vs = float(pyo.value(model.vs_destroyed_kg_d))
    dig = float(pyo.value(model.digestate_dry_solids_kg_d))
    return {
        "feed_dry_solids_kg_d": feed,
        "solids_balance_residual_kg_d": feed - vs - dig,
    }


def get_model_summary(model):
    return {
        "stage1_termination": getattr(model, "stage1_termination", None),
        "stage2_termination": getattr(model, "stage2_termination", None),
        "Q_ani": pyo.value(model.Q_ani),
        "Q_fw": pyo.value(model.Q_fw),
        "Q_sl": pyo.value(model.Q_sl),
        "Q_total": pyo.value(model.Q_total),
        "V_reactor": pyo.value(model.V_reactor),
        "C_reactor": pyo.value(model.C_reactor),
        "ad_mixer_cost": pyo.value(model.ad_mixer_cost),
        "ad_mixer_count": AD_MIXER_PER_TRAIN * pyo.value(model.ad_train_count) if hasattr(model,"ad_train_count") else None,
        "AD_train_count": pyo.value(model.ad_train_count),
        "AD_capacity_per_train_m3d": REACTOR_FLOW_REF,
        **{
            f"AD_train_{t}_active":
                pyo.value(model.ad_train_active[t])
            for t in model.AD_TRAIN
        },
        **{
            f"AD_train_{t}_flow":
                pyo.value(model.Q_ad_train[t])
            for t in model.AD_TRAIN
        },
        **{
            f"AD_train_{t}_volume":
                pyo.value(model.V_reactor_train[t])
            for t in model.AD_TRAIN
        },
        **{
            f"AD_{n}_train_selected":
                pyo.value(model.ad_train_count_sel[n])
            for n in model.AD_TRAIN_COUNT_OPTION
        },
        **feed_cod_summary(model),
        "surrogate_COD_coefficient_kg_m3":
            SURROGATE_COD_COEF,
        # [v51] 이 값들은 원료 실측 COD 가 아니라 surrogate 학습 basis
        #  (SURROGATE_COD_COEF = 75.76 kgCOD/m3 일괄)로 만든 NN 입력이다.
        #  이름만 COD_total 이라 실측 COD 로 오해되기 쉬워 접두어를 붙였다.
        #  실측 COD 질량은 physical_COD_*_kg_d 열에 따로 있다.
        "surrogate_COD_total": pyo.value(model.COD_total),
        "surrogate_OLR": pyo.value(model.OLR),
        "surrogate_input_S_su": pyo.value(model.ADM_total["S_su"]),
        "surrogate_input_S_aa": pyo.value(model.ADM_total["S_aa"]),
        "surrogate_input_S_fa": pyo.value(model.ADM_total["S_fa"]),
        "surrogate_input_S_I": pyo.value(model.ADM_total["S_I"]),
        "surrogate_input_X_ch": pyo.value(model.ADM_total["X_ch"]),
        "surrogate_input_X_pr": pyo.value(model.ADM_total["X_pr"]),
        "surrogate_input_X_li": pyo.value(model.ADM_total["X_li"]),
        "surrogate_input_X_I": pyo.value(model.ADM_total["X_I"]),
        "pred_ch4_frac_dry": pyo.value(model.pred_ch4_frac_dry_var),
        "pred_co2_frac_nn_raw": pyo.value(model.pred_co2_frac_nn_var),
        "gas_fraction_sum_enforced":
            pyo.value(model.pred_ch4_frac_dry_var)
            + pyo.value(model.pred_co2_frac_var),
        "pred_q_gas": pyo.value(model.pred_q_gas_var),
        "surrogate_low_flow_case": (
            1 if pyo.value(model.Q_ad_in) < 10.0 else 0
        ),
        "pred_co2_frac_dry": pyo.value(model.pred_co2_frac_var),
        "storage_in_q_gas": pyo.value(model.storage_in_q_gas),
        "storage_out_q_gas": pyo.value(model.storage_out_q_gas),
        "desulf_in_q_gas": pyo.value(model.desulf_in_q_gas),
        "desulf_out_q_gas": pyo.value(model.desulf_out_q_gas),
        "desulf_cost": pyo.value(model.desulf_cost),
        **{
            f"desulf_selected_cost_{d}":
                pyo.value(model.desulf_cost_selected[d])
            for d in model.DESULF
        },
        "dehum_in_q_gas": pyo.value(model.dehum_in_q_gas),
        "dehum_out_q_gas": pyo.value(model.dehum_out_q_gas),
        "dehum_cost": pyo.value(model.dehum_cost),
        "AD_raw_CH4_m3_d": pyo.value(model.pur_in_ch4_flow),
        "purification_branch_biogas_m3_d":
            pyo.value(model.pur_in_q_gas),
        "generator_branch_biogas_m3_d":
            pyo.value(model.gen_biogas_flow),
        "purification_branch_CH4_m3_d":
            pyo.value(model.purification_ch4_flow),
        "generator_branch_CH4_m3_d":
            pyo.value(model.gen_ch4_flow),
        "pur_in_q_gas": pyo.value(model.pur_in_q_gas),
        "pur_in_ch4_flow_pw": pyo.value(model.pur_in_ch4_flow),
        "pur_in_co2_flow_pw": pyo.value(model.pur_in_co2_flow),
        "pur_in_ch4_flow_exact_check":
            pyo.value(model.dehum_out_q_gas)
            * pyo.value(model.pur_in_ch4_frac),
        "pur_in_ch4_flow_approx_error":
            pyo.value(model.pur_in_ch4_flow)
            - (
                pyo.value(model.dehum_out_q_gas)
                * pyo.value(model.pur_in_ch4_frac)
            ),
        "pur_in_co2_flow_exact_check":
            pyo.value(model.dehum_out_q_gas)
            * pyo.value(model.pur_in_co2_frac),
        "pur_in_co2_flow_approx_error":
            pyo.value(model.pur_in_co2_flow)
            - (
                pyo.value(model.dehum_out_q_gas)
                * pyo.value(model.pur_in_co2_frac)
            ),
        "up_ch4": pyo.value(model.up_ch4),
        "purification_revenue_usd_d": pyo.value(model.purification_revenue),
        "electricity_revenue_usd_d": pyo.value(model.electricity_revenue),
        "total_revenue_usd_d": pyo.value(model.total_revenue),
        # [v73] 목적함수가 대체편익인지(1) 판매수익인지(0)
        "objective_uses_demand": int(pyo.value(model.objective_uses_demand)),
        "route_purification": pyo.value(model.route_sel["purification"]),
        "route_generator": pyo.value(model.route_sel["generator"]),
        "optimized_methane_target_m3_d":
            pyo.value(model.up_ch4),
        "up_co2": pyo.value(model.up_co2),
        "up_gas": pyo.value(model.up_gas),
        "up_ch4_frac_design": pyo.value(model.up_ch4_frac),
        "gen_biogas_flow_m3_d":
            pyo.value(model.gen_biogas_flow),
        "gen_CH4_flow_m3_d":
            pyo.value(model.gen_ch4_flow),
        "gen_power_kW": pyo.value(model.gen_power_kW),
        "gen_cost": pyo.value(model.gen_cost),
        "gen_cost_exact_check":
            next(
                (
                    pyo.value(model.gen_min_cost_USD[g])
                    * (
                        max(
                            pyo.value(model.gen_power_kW),
                            pyo.value(model.gen_min_capacity_kW[g]),
                        )
                        / pyo.value(model.gen_min_capacity_kW[g])
                    ) ** pyo.value(model.gen_scale_factor[g])
                    for g in model.GEN
                    if pyo.value(model.g[g]) > 0.5
                ),
                0.0,
            ),
        "gen_cost_piecewise_error":
            (
                pyo.value(model.gen_cost)
                - next(
                    (
                        pyo.value(model.gen_min_cost_USD[g])
                        * (
                            max(
                                pyo.value(model.gen_power_kW),
                                pyo.value(model.gen_min_capacity_kW[g]),
                            )
                            / pyo.value(model.gen_min_capacity_kW[g])
                        ) ** pyo.value(model.gen_scale_factor[g])
                        for g in model.GEN
                        if pyo.value(model.g[g]) > 0.5
                    ),
                    0.0,
                )
            ),
        **{
            f"generator_cost_if_{g}":
                pyo.value(model.gen_cost_if[g])
            for g in model.GEN
        },
        **{
            f"generator_selected_cost_{g}":
                pyo.value(model.gen_cost_selected[g])
            for g in model.GEN
        },
        "AD_inlet_flow_m3_d":
            pyo.value(model.Q_ad_in),
        "AD_inlet_equivalent_t_d":
            (
                pyo.value(model.Q_ad_in)
                * RAW_FEED_DENSITY_KG_M3
                / TON_TO_KG
            ),
        "AD_feed_mass_kg_d":
            pyo.value(model.ad_feed_mass_kg_d),
        "biogas_CH4_mass_kg_d":
            pyo.value(model.biogas_ch4_mass_kg_d),
        "biogas_CO2_mass_kg_d":
            pyo.value(model.biogas_co2_mass_kg_d),
        "biogas_dry_mass_kg_d":
            pyo.value(model.biogas_dry_mass_kg_d),
        "digestate_mass_kg_d":
            pyo.value(model.digestate_mass_kg_d),
        "digestate_TS_computed":
            (
                pyo.value(model.digestate_dry_solids_kg_d)
                / pyo.value(model.digestate_mass_kg_d)
                if pyo.value(model.digestate_mass_kg_d) > 1.0e-9
                else None
            ),
        "AD_feed_dry_solids_kg_d":
            pyo.value(model.ad_feed_dry_solids_kg_d),
        "COD_removed_kg_d":
            pyo.value(model.cod_removed_kg_d),
        "VS_destroyed_kg_d":
            pyo.value(model.vs_destroyed_kg_d),
        "VS_destruction_fraction":
            (
                pyo.value(model.vs_destroyed_kg_d)
                / pyo.value(model.ad_feed_dry_solids_kg_d)
                if pyo.value(model.ad_feed_dry_solids_kg_d) > 1.0e-9
                else None
            ),
        "solids_capture_assumption":
            SOLIDS_CAPTURE_EFFICIENCY,
        "cake_TS_assumption":
            CAKE_TS_FRACTION,
        "digestate_dry_solids_kg_d":
            pyo.value(model.digestate_dry_solids_kg_d),
        "captured_solids_kg_d":
            pyo.value(model.captured_solids_kg_d),
        "cake_mass_kg_d":
            pyo.value(model.cake_mass_kg_d),
        "cake_flow_m3_d":
            pyo.value(model.cake_flow_m3_d),
        "centrate_mass_kg_d":
            pyo.value(model.centrate_mass_kg_d),
        "centrate_flow_m3_d":
            pyo.value(model.centrate_flow_m3_d),
        "dewatering_mass_balance_error_kg_d":
            (
                pyo.value(model.digestate_mass_kg_d)
                - pyo.value(model.cake_mass_kg_d)
                - pyo.value(model.centrate_mass_kg_d)
            ),
        "digestate_flow_m3_d":
            pyo.value(model.digestate_flow),
        "digestate_equivalent_t_d":
            (
                pyo.value(model.digestate_flow)
                * DIGESTATE_DENSITY_KG_M3
                / TON_TO_KG
            ),
        "digestate_flow":
            pyo.value(model.digestate_flow),
        "AD_mass_balance_error_kg_d":
            (
                pyo.value(model.ad_feed_mass_kg_d)
                - pyo.value(model.biogas_dry_mass_kg_d)
                - pyo.value(model.digestate_mass_kg_d)
            ),
        "digestate_flow_ratio_to_AD_inlet":
            (
                pyo.value(model.digestate_flow)
                / pyo.value(model.Q_ad_in)
                if pyo.value(model.Q_ad_in) > 0
                else 0.0
            ),
        "digestate_dewatering_cost":
            pyo.value(model.digestate_dewatering_cost),
        "digestate_press_cost_exact_check":
            next(
                (
                    pyo.value(model.digestate_min_cost_USD[d])
                    * (
                        max(
                            pyo.value(model.digestate_flow),
                            pyo.value(model.digestate_min_vol_m3d[d]),
                        )
                        / pyo.value(model.digestate_min_vol_m3d[d])
                    ) ** pyo.value(model.digestate_scale_factor[d])
                    for d in model.DIGESTATE_DEWATERING
                    if pyo.value(model.dd[d]) > 0.5
                ),
                0.0,
            ),
        **{
            f"digestate_dewatering_cost_if_{d}":
                pyo.value(model.digestate_dewatering_cost_if[d])
            for d in model.DIGESTATE_DEWATERING
        },
        **{
            f"digestate_dewatering_selected_cost_{d}":
                pyo.value(model.digestate_dewatering_cost_selected[d])
            for d in model.DIGESTATE_DEWATERING
        },
        **{
            f"digestate_conveyor_cost_{c}":
                pyo.value(model.digestate_conveyor_cost_if[c])
            for c in model.DIGESTATE_CONVEYOR
        },
        **{
            f"digestate_conveyor_ref_vol_m3d_{c}":
                pyo.value(model.digestate_conveyor_ref_vol_m3d[c])
            for c in model.DIGESTATE_CONVEYOR
        },
        **{
            f"digestate_conveyor_ref_cost_USD_{c}":
                pyo.value(model.digestate_conveyor_ref_cost_USD[c])
            for c in model.DIGESTATE_CONVEYOR
        },
        **{
            f"digestate_conveyor_scale_factor_{c}":
                pyo.value(model.digestate_conveyor_scale_factor[c])
            for c in model.DIGESTATE_CONVEYOR
        },
        **{
            f"digestate_conveyor_exact_cost_check_{c}":
                (
                    pyo.value(model.digestate_conveyor_ref_cost_USD[c])
                    * (
                        max(
                            pyo.value(
                                model.digestate_flow
                                if c == "DS_conveyor"
                                else model.cake_flow_m3_d
                            ),
                            pyo.value(model.digestate_conveyor_ref_vol_m3d[c]),
                        )
                        / pyo.value(model.digestate_conveyor_ref_vol_m3d[c])
                    ) ** pyo.value(model.digestate_conveyor_scale_factor[c])
                )
            for c in model.DIGESTATE_CONVEYOR
        },
        "tank_cost": pyo.value(model.tank_cost),
        "thermal_oxidizer_cost": pyo.value(model.thermal_oxidizer_cost),
        "pure_cost": pyo.value(model.pure_cost),
        "gas_compressor_cost": pyo.value(model.gas_compressor_cost),
        **{
            f"purification_selected_cost_{tech}":
                pyo.value(model.pure_cost_selected[tech])
            for tech in model.TECH
        },
        "FW_pretreatment_cost": pyo.value(model.fw_pr_cost),
        "FW_line1_flow_after_screen1": pyo.value(model.Q_fw_l1_after_screen),
        "FW_line1_flow_after_sorting1": pyo.value(model.Q_fw_l1_after_sorting),
        "FW_line2_flow_after_screen2": pyo.value(model.Q_fw_l2_after_screen),
        "FW_line2_flow_after_sorting2": pyo.value(model.Q_fw_l2_after_sorting),
        "FW_flow_after_merge": pyo.value(model.Q_fw_after_merge),
        "FW_flow_after_shredder1": pyo.value(model.Q_fw_after_shredder1),
        "FW_flow_after_screen3": pyo.value(model.Q_fw_after_screen3),
        "FW_flow_after_shredder2": pyo.value(model.Q_fw_after_shredder2),
        "FW_waste_from_screen1": pyo.value(model.Q_fw_waste_line1),
        "FW_waste_from_screen2": pyo.value(model.Q_fw_waste_line2),
        "FW_waste_from_screen3": pyo.value(model.Q_fw_waste_screen3),
        "FW_waste_total": pyo.value(model.Q_fw_waste_total),
        "FW_waste_handling_cost": pyo.value(model.fw_waste_handling_cost),
        "receiving_equipment_cost": pyo.value(model.receiving_cost),
        "pretreatment_aux_equipment_cost": pyo.value(model.pretreatment_aux_cost),
        "BM_storage_in_flow": pyo.value(model.Q_bm_storage_in),
        "BM_storage_line1_flow": pyo.value(model.Q_bm_line1),
        "BM_storage_line2_flow": pyo.value(model.Q_bm_line2),
        "BM_storage_out_flow": pyo.value(model.Q_bm_storage_out),
        "AD_actual_inlet_flow": pyo.value(model.Q_ad_in),
        "BM_storage_cost": pyo.value(model.bm_storage_cost),
        **{
            f"BM_storage_cost_{equip}":
                pyo.value(model.bm_storage_cost_if[equip])
            for equip in model.BM_STORAGE_EQUIP
        },
        "AM_post_pretreatment_flow": pyo.value(model.Q_ani),
        "SL_post_pretreatment_flow": pyo.value(model.Q_sl),
        "FW_post_pretreatment_flow": pyo.value(model.Q_fw),
        **{
            f"pretreatment_aux_cost_{equip}":
                pyo.value(model.pretreatment_aux_cost_if[equip])
            for equip in model.PRETREATMENT_AUX
        },
        "wastewater_in_flow_m3_d":
            pyo.value(model.wastewater_in_flow),
        # [v72] 폐수처리 설비 설계유량 = 유입 x 1.3
        "wwtp_design_flow_m3_d":
            pyo.value(model.wwtp_design_flow),
        "wastewater_cost":
            pyo.value(model.wastewater_cost),
        **{
            f"wwtp_option_cost_if_{option}":
                pyo.value(model.wwtp_option_cost_if[option])
            for option in model.WWTP_OPTION
        },
        **{
            f"wwtp_selected_cost_{option}":
                pyo.value(model.wwtp_option_cost_selected[option])
            for option in model.WWTP_OPTION
        },
        **{
            f"wwtp_equipment_cost_{equip}":
                pyo.value(model.wwtp_equipment_cost_if[equip])
            for equip in model.WWTP_EQUIPMENT
        },
        "Cost": pyo.value(model.Cost),
        **capex_summary(model),
        **nn_input_diagnostics(model),
        **get_selected_options(model),
    }


# =========================================================
# 적응형 epsilon 변곡점 탐색 (v44.1)
# =========================================================
def _design_signature(model):
    """
    solve 된 모델에서 '설계 조합'을 비교 가능한 튜플로 뽑는다.

    get_selected_options() 가 반환하는 이진/정수 선택변수 전체를 사용한다.
    수치오차를 무시하기 위해 반올림해서 정수화한다.
    두 해의 signature 가 같으면 = 동일한 설계 조합 = Pareto 상의 같은 점.
    """
    opts = get_selected_options(model)
    sig = []
    for k in sorted(opts.keys()):
        v = opts[k]
        if v is None:
            sig.append((k, None))
        else:
            sig.append((k, int(round(float(v)))))
    return tuple(sig)


def solve_case(epsilon, W, F, S, grid_hint=None, obj_floor=None, demand=None):
    """
    단일 CAPEX 상한 epsilon에서 CH4를 최대화한다.

    1단계: max upgraded CH4
    2단계: 1단계 CH4를 유지하면서 CAPEX 최소화
    3단계: 최종 Pareto 해의 LCOB를 사후 계산

    반환: (termination, summary, design_signature)
    """
    demand = demand or {}
    attempts = [grid_hint] if grid_hint is not None else []
    attempts.append(None)
    last = ("error", {"error_message": "no attempt"}, ("__error__",))

    for hint in attempts:
        try:
            model = build_model(
                epsilon=epsilon,
                W=W, F=F, S=S,
                probe_mode=False,
                grid_hint=hint,
                obj_floor=obj_floor,
                gas_demand_kwh_d=demand.get("gas"),
                elec_demand_kwh_d=demand.get("elec"),
            )

            term = solve_maxobj_then_mincost(model)
            if term.lower() not in ("optimal", "feasible", "stage1_only"):
                last = (term, {}, ("__" + str(term).lower() + "__",))
                continue

            summary = get_model_summary(model)
            summary.update(solids_balance_check(model, W, F, S))

            annual_cost = float(pyo.value(model.lcob_numerator_usd_y))
            annual_ch4 = float(pyo.value(model.lcob_denominator_m3_y))
            if annual_ch4 > 1.0e-9:
                lcob = annual_cost / annual_ch4
                summary["LCOB_usd_per_m3_CH4"] = lcob
                summary["LCOB_krw_per_m3_CH4"] = (
                    lcob * EXCHANGE_RATE_KRW_PER_USD
                )
                summary["LCOB_to_gas_price_ratio"] = (
                    lcob / CITY_GAS_PRICE_USD_PER_M3_CH4
                )
                # [v79b] 반입료 차감 순 LCOB (LCOB_net) — capex_selector lcob_net 대응
                lcob_net = float(pyo.value(model.TAC_net_usd_y)) / annual_ch4
                summary["LCOB_net_usd_per_m3_CH4"] = lcob_net
                summary["LCOB_net_krw_per_m3_CH4"] = (
                    lcob_net * EXCHANGE_RATE_KRW_PER_USD
                )
                # [v79c] 탄소배출권 수익 → LCOB 4종 (기존 / 반입료 / 탄소 / 탄소+반입료)
                _pgas = float(pyo.value(model.pred_q_gas_var))
                _co2_tco2_y, _carbon_usd_y = carbon_credit_usd_y(_pgas)
                _tip_usd_y = float(pyo.value(model.tipping_usd_y))
                summary["CO2_reduction_tCO2_y"] = _co2_tco2_y
                summary["carbon_credit_krw_y"] = _co2_tco2_y * CARBON_PRICE_KRW_PER_TCO2
                summary["carbon_credit_usd_y"] = _carbon_usd_y
                # ③ 탄소배출권만 차감
                lcob_carbon = (annual_cost - _carbon_usd_y) / annual_ch4
                summary["LCOB_carbon_usd_per_m3_CH4"] = lcob_carbon
                summary["LCOB_carbon_krw_per_m3_CH4"] = (
                    lcob_carbon * EXCHANGE_RATE_KRW_PER_USD
                )
                # ④ 탄소배출권 + 반입료 둘 다 차감
                lcob_both = (annual_cost - _carbon_usd_y - _tip_usd_y) / annual_ch4
                summary["LCOB_carbon_net_usd_per_m3_CH4"] = lcob_both
                summary["LCOB_carbon_net_krw_per_m3_CH4"] = (
                    lcob_both * EXCHANGE_RATE_KRW_PER_USD
                )
            else:
                summary["LCOB_usd_per_m3_CH4"] = None
                summary["LCOB_krw_per_m3_CH4"] = None
                summary["LCOB_to_gas_price_ratio"] = None
                summary["LCOB_net_usd_per_m3_CH4"] = None
                summary["LCOB_net_krw_per_m3_CH4"] = None
                summary["LCOB_carbon_usd_per_m3_CH4"] = None
                summary["LCOB_carbon_krw_per_m3_CH4"] = None
                summary["LCOB_carbon_net_usd_per_m3_CH4"] = None
                summary["LCOB_carbon_net_krw_per_m3_CH4"] = None
                summary["CO2_reduction_tCO2_y"] = None
                summary["carbon_credit_krw_y"] = None
                summary["carbon_credit_usd_y"] = None

            # Pareto 후처리는 일일 판매수익을 최대화하는 방향으로 판단한다.
            # 목적함수와 반드시 같은 값이어야 한다. 다르면 obj_floor 가
            # 단위 안 맞는 제약이 되어 가짜 infeasible 을 만든다.
            # [v73.1] objective_value 는 반드시 objective_expr 와 같아야 한다.
            #   obj_floor 로 되먹임되기 때문에 다른 값을 넣으면
            #   달성 불가능한 하한이 걸려 이후 epsilon 이 전부 infeasible 이 된다.
            #   (v73 에서 목적을 대체편익으로 바꾸고도 여기는 total_revenue 를
            #    기록해, 수요가 포화된 동산면이 40점 중 38점 infeasible 이 됐다.)
            summary["objective_value"] = float(
                pyo.value(model.objective_expr)
            )
            summary["pareto_objective"] = (
                "maximize_useful_delivered_energy_subject_to_TAC"
                if int(pyo.value(model.objective_uses_demand))
                else "maximize_delivered_energy_subject_to_TAC"
            )
            summary["LCOB_role"] = "post_evaluation_only"
            summary["grid_hint_used"] = 0 if hint is None else 1

            # [v65] Pareto 유효성 — 1·2단계 모두 'optimal' 로 증명된 해만
            #  최종 Pareto 시트에 쓴다. feasible/stage1_only 는 최적성이
            #  증명되지 않아 가짜 비단조·지배해를 만들 수 있다.
            _s1 = str(summary.get("stage1_termination", "")).lower()
            _s2 = str(summary.get("stage2_termination", "")).lower()
            summary["pareto_valid"] = int(_s1 == "optimal" and _s2 == "optimal")

            # [v65] epsilon 의 의미를 열 이름으로 명시.
            #  같은 행의 정제·발전은 '같은 CAPEX' 가 아니라
            #  '같은 CAPEX 상한' 아래에서 각각 얻은 최적해다.
            summary["epsilon_meaning"] = "max_allowed_TAC_usd_per_year"
            summary["selected_route_CAPEX_usd"] = summary.get("Cost")
            summary["LCOB_converged"] = None
            summary["LCOB_dinkelbach_iters"] = 0
            # =========================================================
            # [v66] route 는 모델이 골랐다. 선택 결과에 따라 지표를 나눈다.
            #   정제 선택 -> LCOB [$/m3 CH4]
            #   발전 선택 -> LCOE [$/kWh]
            #   공통      -> 균등화 에너지원가 [$/kWh]  (목적함수와 동일 기준)
            # =========================================================
            _pur = float(pyo.value(model.route_sel["purification"]))
            route = "purification" if _pur > 0.5 else "generator"
            # [v75] 균등화 원가의 분모는 '유효공급에너지'(수요 상한 반영).
            #  총생산량은 delivered_* 로 따로 기록해 비교할 수 있게 둔다.
            e_kwh_d = float(pyo.value(model.useful_energy_kwh_d))
            e_kwh_y = float(pyo.value(model.useful_energy_kwh_y))
            kw = float(pyo.value(model.gen_power_kW))
            ch4 = float(pyo.value(model.up_ch4))

            summary["selected_route"] = route
            # [v78] 폐수처리 공법 선택 결과. v77 에서 약품비가 들어가면서
            #  Anammox 와 A2O/Bardenpho 가 실제로 경쟁하게 됐으므로
            #  요약 시트에서 바로 보여야 한다.
            summary["selected_wwtp"] = next(
                (o for o in wwtp_options
                 if float(pyo.value(model.wwtp_select[o])) > 0.5),
                None,
            )
            if route == "purification":
                summary["PUR_CAPEX_usd"] = summary.get("Cost")
                summary["GEN_CAPEX_usd"] = None
            else:
                summary["PUR_CAPEX_usd"] = None
                summary["GEN_CAPEX_usd"] = summary.get("Cost")
            summary["route_choice_driver"] = "useful_delivered_energy_kWh_per_day"
            summary["daily_revenue_usd_d"] = float(
                pyo.value(model.total_revenue)
            )
            summary["delivered_energy_kWh_d"] = float(
                pyo.value(model.delivered_energy_kwh_d)
            )
            summary["delivered_energy_kWh_y"] = float(
                pyo.value(model.delivered_energy_kwh_y)
            )
            summary["levelized_energy_cost_usd_per_kWh"] = (
                annual_cost / e_kwh_y if e_kwh_y > 1.0e-9 else None
            )
            summary["levelized_energy_cost_krw_per_kWh"] = (
                annual_cost / e_kwh_y * EXCHANGE_RATE_KRW_PER_USD
                if e_kwh_y > 1.0e-9 else None
            )

            if route == "purification":
                summary["selected_purification_tech"] = next(
                    (u for u in model.TECH if pyo.value(model.y[u]) > 0.5), None
                )
                summary["selected_generator"] = None
                summary["GEN_power_kW"] = 0.0
                summary["GEN_electricity_kWh_d"] = 0.0
                summary["GEN_LCOE_usd_per_kWh"] = None
                summary["sale_price_usd_per_kWh"] = CITY_GAS_PRICE_USD_PER_KWH
            else:
                summary["selected_purification_tech"] = None
                summary["selected_generator"] = next(
                    (g for g in model.GEN if pyo.value(model.g[g]) > 0.5), None
                )
                summary["GEN_power_kW"] = kw
                summary["GEN_electricity_kWh_d"] = kw * 24.0
                summary["GEN_LCOE_usd_per_kWh"] = (
                    annual_cost / (kw * 24.0 * DAYS_PER_YEAR)
                    if kw > 1.0e-9 else None
                )
                summary["LCOB_usd_per_m3_CH4"] = None
                summary["LCOB_krw_per_m3_CH4"] = None
                summary["LCOB_to_gas_price_ratio"] = None
                summary["LCOB_net_usd_per_m3_CH4"] = None
                summary["LCOB_net_krw_per_m3_CH4"] = None
                summary["LCOB_carbon_usd_per_m3_CH4"] = None
                summary["LCOB_carbon_krw_per_m3_CH4"] = None
                summary["LCOB_carbon_net_usd_per_m3_CH4"] = None
                summary["LCOB_carbon_net_krw_per_m3_CH4"] = None
                summary["CO2_reduction_tCO2_y"] = None
                summary["carbon_credit_krw_y"] = None
                summary["carbon_credit_usd_y"] = None
                summary["sale_price_usd_per_kWh"] = ELEC_PRICE_USD_PER_KWH

            # 원가/판매가 비율 (선택된 경로 기준, 무차원)
            _lec = summary["levelized_energy_cost_usd_per_kWh"]
            _p = summary["sale_price_usd_per_kWh"]
            summary["cost_to_price_ratio"] = (
                _lec / _p if (_lec is not None and _p) else None
            )
            summary["breaks_even"] = (
                int(summary["cost_to_price_ratio"] <= 1.0)
                if summary["cost_to_price_ratio"] is not None else None
            )

            _s1 = str(summary.get("stage1_termination", "")).lower()
            _s2 = str(summary.get("stage2_termination", "")).lower()
            summary["pareto_valid"] = int(_s1 == "optimal" and _s2 == "optimal")
            summary["epsilon_meaning"] = "max_allowed_CAPEX_usd"

            try:
                sig = _design_signature(model)
            except Exception:
                sig = ("__sig_error__",)
            return term, summary, sig

        except Exception as e:
            last = ("error", {"error_message": str(e)}, ("__error__",))

    return last


def adaptive_epsilon_solves(W, F, S, budget=None, demand=None):
    """
    지역 하나에 대해 적응형(이분탐색) epsilon 스윕을 수행한다.

    아이디어:
      - 비용범위 [C_min, C_sat] 의 양 끝을 먼저 solve.
      - 두 끝점의 설계 조합(signature)이 다르면 그 사이에 '변곡점'이 있다.
        구간을 이분하며 조합이 바뀌는 임계 epsilon 을 찾아간다.
      - 조합이 같은 구간은 더 파지 않는다(계단형 Pareto 가정).
      - solve 횟수는 budget(=N_EPS) 로 제한한다.

    반환: (solved, mode)
      solved : [(epsilon, term, extra, signature), ...] epsilon 오름차순
      mode   : 'adaptive' | 'uniform' | 'fallback_uniform' | 'degenerate_single'
    """
    if budget is None:
        budget = N_EPS

    # --- 적응형 비활성화 시: 기존 균등 등분 동작 유지(하위호환) ---
    if not ADAPTIVE_EPS:
        # [수정] v57까지 이 경로는 grid_hint 를 넘기지 않아 축소격자가 적용되지 않았고,
        #        CH4 이중선형 오차가 0.000% -> -8.5% 로 되돌아갔다.
        # [v65] eps 범위를 정제·발전 두 route 의 합집합으로 잡는다.
        #  v63 은 정제 기준으로만 만들어서, 발전설비가 더 싼 경우
        #  eps 하한이 이미 발전 C_min 보다 커 발전 route 가 전 구간에서
        #  한 번도 비용제약을 받지 않고 같은 해만 반복됐다.
        # [v66] route 가 자유변수이므로 probe 한 번이면 두 경로가 모두 고려된다.
        rng0 = probe_cost_range(W, F, S, demand=demand)
        if rng0 is None:
            _tot = max(0.0, float(W) + float(F) + float(S))
            _c = 2_000_000.0 + 45_000.0 * _tot
            _lo, _hi = max(500_000.0, 0.75 * _c), max(500_001.0, 1.25 * _c)
            _hint = None
        else:
            _cmin, _csat, _, _hint = rng0
            _span = max(_csat - _cmin, 1.0)
            _lo = _cmin + max(1000.0, 0.002 * _span)
            _hi = _csat + max(1000.0, 1.0e-5 * _csat)
        # =========================================================
        # [v78] grid_hint 미적용 경고
        #
        # grid_hint 는 probe 가 찾아준 (총가스량, CH4분율) 범위로 이중선형
        # 근사격자를 좁히는 값이다. 이게 없으면 격자가 넓어져 근사오차가 커지고,
        # 그 오차가 route 선택 자체를 뒤집는다.
        #
        # 태백시 실측 (동일 epsilon, 동일 모델):
        #   grid_hint 있음 -> purification, 유효공급 11,279 kWh/d, TAC 3,115,854
        #   grid_hint 없음 -> generator,    유효공급  4,523 kWh/d, TAC 2,985,074
        # 공급량이 2.5 배 차이인데 두 경우 모두 solver 는 'optimal' 을 보고한다.
        #
        # 따라서 probe 실패 행은 '풀렸지만 믿을 수 없는 해'다. 조용히 섞이면
        # 안 되므로 플래그를 세워 pareto_valid 에서 떨어뜨린다.
        # =========================================================
        _hint_ok = int(_hint is not None)
        if not _hint_ok:
            print(
                "    [경고] probe 실패로 grid_hint 없이 스윕한다. "
                "이중선형 근사오차가 route 선택을 뒤집을 수 있어 "
                "이 행은 pareto_valid=0 으로 표시된다."
            )
        eps_list = [float(x) for x in np.linspace(_lo, _hi, budget)]
        solved = []
        best_obj = None
        for i, eps in enumerate(eps_list, start=1):
            _t = time.time()
            term, extra, sig = solve_case(
                eps, W, F, S, grid_hint=_hint, obj_floor=best_obj,
                demand=demand,
            )
            ex = extra if isinstance(extra, dict) else {}
            ex["grid_hint_applied"] = _hint_ok
            objv = ex.get("objective_value")
            cost = ex.get("TAC_usd_y")      # [v75] epsilon 축은 TAC
            ex["epsilon_binding"] = 1 if cost is None else int(
                cost >= eps - COST_CAP_BINDING_TOL
            )
            ex["solve_reused"] = 0
            if objv is not None:
                best_obj = objv if best_obj is None else max(best_obj, objv)
            # [v75] 진행표시: 목적함수가 '유효공급에너지' 이므로 그것과 LCOE 를 보여준다.
            lcoe = ex.get("levelized_energy_cost_usd_per_kWh")
            print(
                f"      solve {i:2d}/{budget} eps={eps:,.0f} {term} "
                f"{time.time() - _t:5.1f}s"
                + (f" | {ex.get('selected_route', '?'):12s}" if objv is not None else "")
                + (f" 공급 {objv:8,.0f} kWh/d" if objv is not None else "")
                + (f" TAC {ex.get('TAC_usd_y', 0):9,.0f} $/y" if objv is not None else "")
                + (f" | LCOE {lcoe:6.4f} $/kWh"
                   f" ({lcoe * EXCHANGE_RATE_KRW_PER_USD:5,.0f} W/kWh)"
                   if lcoe is not None else ""),
                flush=True,
            )
            solved.append((eps, term, extra, sig))

        return solved, "uniform_pareto"

    rng = probe_cost_range(W, F, S, demand=demand)

    # --- probe 실패: 유량기반 fallback 범위를 균등 스윕 ---
    if rng is None:
        total_raw_flow = max(0.0, float(W) + float(F) + float(S))
        estimated_center = 2_000_000.0 + 45_000.0 * total_raw_flow
        lo = max(500_000.0, 0.75 * estimated_center)
        hi = max(lo + 1.0, 1.25 * estimated_center)
        print(
            "  [probe 실패 → 유량기반 fallback 균등 스윕: "
            f"{lo:,.0f} ~ {hi:,.0f} USD]"
        )
        print(
            "    [경고] grid_hint 없이 스윕한다. 이중선형 근사오차가 "
            "route 선택을 뒤집을 수 있어 이 행은 pareto_valid=0 이다."
        )
        solved = []
        for eps in np.linspace(lo, hi, budget):
            # [v78] demand 를 빠뜨리고 있었다. 수요 상한이 안 걸리면
            #  목적함수가 '총생산량 최대화'로 조용히 퇴화한다.
            term, extra, sig = solve_case(float(eps), W, F, S, demand=demand)
            if isinstance(extra, dict):
                extra["grid_hint_applied"] = 0
            solved.append((float(eps), term, extra, sig))
        return solved, "fallback_uniform"

    C_min, C_sat, _, _grid_hint = rng

    # --- 범위가 사실상 0: 중앙 1점만 ---
    if C_sat - C_min < 1.0:
        eps = max(C_sat, 1.0)
        # [수정] 이 분기에서 grid_hint 를 넘기지 않으면 축소격자가 적용되지 않아
        #        CH4 이중선형 오차가 다시 커진다(실측 0.000% -> -8.5%).
        term, extra, sig = solve_case(
            eps, W, F, S, grid_hint=_grid_hint, demand=demand
        )
        return [(eps, term, extra, sig)], "degenerate_single"

    span = C_sat - C_min
    # C_min 은 본 sweep 과 동일한 모델에서 얻은 값이라 소폭 여유만 둔다.
    lo = C_min + max(1000.0, 0.002 * span)
    hi = C_sat + max(1000.0, 1.0e-5 * C_sat)
    tol = max(1.0, EPS_BREAKPOINT_TOL_FRAC * span)

    cache = {}          # round(eps,2) -> (eps, term, extra, sig)
    solves_used = [0]   # 리스트로 감싸 클로저에서 수정

    def _best_obj_at_or_below(e):
        """[Fix 4] 이미 푼 eps' <= e 중 최고 목적값.

        eps' <= e 의 해는 e 에서도 실행가능하므로 이 값은 유효한 하한이다.
        """
        best = None
        for (ee, _tt, ex, _sg) in cache.values():
            if ee <= e + 1.0e-9 and isinstance(ex, dict):
                ov = ex.get("objective_value")
                if ov is not None and (best is None or ov > best):
                    best = ov
        return best

    def _solve(e):
        key = round(e, 2)
        if key in cache:
            return cache[key]
        if solves_used[0] >= budget:
            return None
        _t0 = time.time()
        term, extra, sig = solve_case(
            e,
            W,
            F,
            S,
            grid_hint=_grid_hint,
            obj_floor=None,
            demand=demand,
        )
        solves_used[0] += 1
        _ex = extra if isinstance(extra, dict) else {}
        _lcob = _ex.get("LCOB_usd_per_m3_CH4")
        _lcoe = _ex.get("GEN_LCOE_usd_per_kWh")
        _ch4 = _ex.get("up_ch4")
        _msg = f"      solve {solves_used[0]:2d}/{budget} eps={e:,.0f} {term} {time.time() - _t0:5.1f}s"
        if _lcob is not None:
            _msg += f" | LCOB {_lcob:7.3f} $/m3 ({_lcob * EXCHANGE_RATE_KRW_PER_USD:6,.0f} W/m3)"
        if _ch4 is not None:
            _msg += f" CH4 {_ch4:8,.1f} m3/d"
        if _lcoe is not None:
            _msg += f" | LCOE {_lcoe:7.3f} $/kWh"
        print(_msg, flush=True)
        cache[key] = (e, term, extra, sig)
        return cache[key]

    # 양 끝 anchor
    a = _solve(lo)
    b = _solve(hi)

    # 조합이 다른 구간을 '폭이 넓은 순'으로 이분 (최소힙, 폭은 음수로)
    import heapq
    heap = []
    counter = 0
    if a is not None and b is not None and a[3] != b[3]:
        heapq.heappush(heap, (-(hi - lo), counter, lo, hi, a[3], b[3]))
        counter += 1

    while heap and solves_used[0] < budget:
        neg_w, _, ea, eb, sa, sb = heapq.heappop(heap)
        if (eb - ea) <= tol:
            continue  # 임계 epsilon 을 충분히 좁혔다
        mid = 0.5 * (ea + eb)
        m = _solve(mid)
        if m is None:
            break  # solve 예산 소진
        sm = m[3]
        if sm != sa:
            heapq.heappush(heap, (-(mid - ea), counter, ea, mid, sa, sm))
            counter += 1
        if sb != sm:
            heapq.heappush(heap, (-(eb - mid), counter, mid, eb, sm, sb))
            counter += 1

    solved = sorted(cache.values(), key=lambda r: r[0])
    return solved, "adaptive"


# =======================================================================
# [v76] 설계 전환 임계 TAC 추출
# =======================================================================
# 균등격자 sweep 은 front 를 그리기엔 좋지만 "어느 TAC 에서 기술이 갈리는가"는
# 격자 간격만큼의 오차로만 알 수 있다. 설계가 이산이라 front 는 계단함수이고,
# 임계값은 두 격자점 사이 어딘가에 있다.
#
# 여기서는 균등격자 결과를 받아, 인접한 두 점의 이진 선택변수 조합(signature)이
# 다른 구간만 골라 이분탐색한다. 조합이 같은 구간은 건드리지 않으므로
# 추가 solve 는 '실제로 설계가 바뀐 횟수 x log(정밀도)' 로 제한된다.
#
# 이 추가 solve 들도 그 자체로 front 위의 유효한 점이므로 결과에 합쳐서
# 반환한다(= 그림도 촘촘해진다). obj_floor 는 구간 하한의 목적값을 넘겨
# 단조성이 깨지지 않게 한다.
DESIGN_SWITCH_ENABLED = True
DESIGN_SWITCH_TOL_FRAC = 1.0e-3      # 구간폭 / 전체 TAC 범위. 이하로 좁아지면 종료
DESIGN_SWITCH_MAX_SOLVES = 24        # 행(읍면동)당 추가 solve 상한


def _sig_changed_keys(sig_a, sig_b):
    """두 설계 signature 사이에서 값이 바뀐 선택변수 이름 목록."""
    da = dict(sig_a or ())
    db = dict(sig_b or ())
    out = []
    for k in sorted(set(da) | set(db)):
        if da.get(k) != db.get(k):
            out.append(f"{k}: {da.get(k)} -> {db.get(k)}")
    return out


def _record_switch(switches, lo_rec, hi_rec):
    """브래킷 [lo_rec, hi_rec] 하나를 '설계 전환 1건'으로 기록한다.

    lo_rec / hi_rec 는 (epsilon, term, extra, signature) 튜플이고,
    두 signature 가 다르다는 것이 전제다. 임계 TAC 는 브래킷 중앙,
    불확실성은 브래킷 폭의 절반이다.
    """
    e_lo, _, ex_lo, sig_lo = lo_rec
    e_hi, _, ex_hi, sig_hi = hi_rec
    ex_lo = ex_lo if isinstance(ex_lo, dict) else {}
    ex_hi = ex_hi if isinstance(ex_hi, dict) else {}
    changed = _sig_changed_keys(sig_lo, sig_hi)

    switches.append({
        "threshold_TAC_usd_y": 0.5 * (e_lo + e_hi),
        "bracket_lo_usd_y": e_lo,
        "bracket_hi_usd_y": e_hi,
        "bracket_width_usd_y": e_hi - e_lo,
        "route_before": ex_lo.get("selected_route"),
        "route_after": ex_hi.get("selected_route"),
        "purification_before": ex_lo.get("selected_purification_tech"),
        "purification_after": ex_hi.get("selected_purification_tech"),
        "generator_before": ex_lo.get("selected_generator"),
        "generator_after": ex_hi.get("selected_generator"),
        "useful_energy_before_kWh_d": ex_lo.get("useful_energy_kWh_d"),
        "useful_energy_after_kWh_d": ex_hi.get("useful_energy_kWh_d"),
        "TAC_before_usd_y": ex_lo.get("TAC_usd_y"),
        "TAC_after_usd_y": ex_hi.get("TAC_usd_y"),
        "LCOE_before_usd_per_kWh": ex_lo.get(
            "levelized_energy_cost_usd_per_kWh"),
        "LCOE_after_usd_per_kWh": ex_hi.get(
            "levelized_energy_cost_usd_per_kWh"),
        "n_changed_options": len(changed),
        "changed_options": " | ".join(changed[:12]),
    })


def find_design_switches(W, F, S, solved, demand=None, grid_hint=None,
                         tol_frac=None, max_solves=None, verbose=True):
    """균등격자 결과에서 설계가 바뀐 구간만 이분탐색해 임계 TAC 를 찾는다.

    반환: (switches, extra_solved)
      switches     : [{...}, ...] 임계 epsilon 과 바뀐 선택변수
      extra_solved : [(eps, term, extra, sig), ...] 탐색 중 얻은 추가 front 점
    """
    tol_frac = DESIGN_SWITCH_TOL_FRAC if tol_frac is None else tol_frac
    max_solves = DESIGN_SWITCH_MAX_SOLVES if max_solves is None else max_solves

    pts = [r for r in sorted(solved, key=lambda r: r[0]) if r[3] is not None]
    if len(pts) < 2:
        return [], []

    span = max(pts[-1][0] - pts[0][0], 1.0)
    tol = max(span * tol_frac, 1.0)

    def _obj_of(rec):
        ex = rec[2] if isinstance(rec[2], dict) else {}
        return ex.get("objective_value")

    switches = []
    extra = []
    used = 0

    # 작업 스택: (하한 rec, 상한 rec). 한 구간 안에 전환이 여러 개 있을 수 있으므로
    # 이분 후 남는 두 반구간을 다시 스택에 넣어 전부 훑는다.
    stack = [
        (lo_rec, hi_rec)
        for lo_rec, hi_rec in zip(pts[:-1], pts[1:])
        if lo_rec[3] != hi_rec[3]
    ]
    stack.reverse()          # epsilon 오름차순으로 처리되게

    while stack and used < max_solves:
        lo_rec, hi_rec = stack.pop()
        e_lo, sig_lo = lo_rec[0], lo_rec[3]
        e_hi, sig_hi = hi_rec[0], hi_rec[3]

        if (e_hi - e_lo) <= tol:
            _record_switch(switches, lo_rec, hi_rec)
            continue

        mid = 0.5 * (e_lo + e_hi)
        term, ex, sig = solve_case(
            mid, W, F, S, grid_hint=grid_hint,
            obj_floor=_obj_of(lo_rec), demand=demand,
        )
        used += 1
        if sig is None:                    # 못 풀면 현재 브래킷을 그대로 기록
            _record_switch(switches, lo_rec, hi_rec)
            continue
        ex = ex if isinstance(ex, dict) else {}
        ex["sweep_point_type"] = "switch_probe"
        mid_rec = (mid, term, ex, sig)
        extra.append(mid_rec)

        # 중점 조합이 어느 쪽과도 다르면 전환이 둘 이상이라는 뜻이다.
        # 그때는 양쪽 반구간을 모두 큐에 넣는다.
        if sig != sig_lo:
            stack.append((lo_rec, mid_rec))
        if sig != sig_hi:
            stack.append((mid_rec, hi_rec))

    # 예산 소진으로 남은 구간도 현재 브래킷 상태로 기록해 둔다(폭이 곧 오차).
    for lo_rec, hi_rec in stack:
        _record_switch(switches, lo_rec, hi_rec)

    switches.sort(key=lambda s: s["threshold_TAC_usd_y"])

    if verbose and used >= max_solves and stack:
        print(f"      [switch] 추가 solve 상한({max_solves}) 도달 — "
              f"미해결 구간 {len(stack)}개는 브래킷 폭 그대로 기록")

    if verbose and switches:
        print(f"      [switch] 설계 전환 {len(switches)}개 "
              f"(추가 solve {used}회)")
        for s in switches:
            print(f"        TAC {s['threshold_TAC_usd_y']:>11,.0f} $/y  "
                  f"±{s['bracket_width_usd_y']/2:>8,.0f}  "
                  f"{s['route_before']}/{s['purification_before'] or s['generator_before']}"
                  f" -> "
                  f"{s['route_after']}/{s['purification_after'] or s['generator_after']}")

    return switches, extra


def run_region_epsilon_sweep(df_run, output_excel):
    """한 지역(df_run: 해당 지역 읍면동들)에 대해 스윕을 돌리고
    결과를 output_excel 로 저장한다."""
    # 장시간 계산 전에 출력 파일 확장자를 검증
    if Path(output_excel).suffix.lower() != ".xlsx":
        raise ValueError(
            f"output_excel must end with '.xlsx': {output_excel}"
        )

    results = []
    switch_records = []          # [v76] 설계 전환 임계 TAC
    total_cases_est = len(df_run) * N_EPS
    case_no = 0
    t0 = time.time()

    for row_idx, row in df_run.iterrows():
        raw = get_raw_streams_from_excel_row(row)
        W = float(raw["Q_ani_raw"])
        F = float(raw["Q_fw_raw"])
        S = float(raw["Q_sl_raw"])

        region_info = {
            "row_index": row_idx,
            "시군": row.get("시군", row.get("시군구", row.get("지역", None))),
            "읍면동": row.get("읍면동", row.get("행정동", None)),
            "리": row.get("리", None),
        }

        if W + F + S <= 0:
            results.append({
                **region_info,
                "epsilon": None,
                **raw,
                "W_raw_t_d": W,
                "F_raw_t_d": F,
                "S_raw_t_d": S,
                "W_converted_m3_d": W * TON_TO_KG / RAW_FEED_DENSITY_KG_M3,
                "F_converted_m3_d": F * TON_TO_KG / RAW_FEED_DENSITY_KG_M3,
                "S_converted_m3_d": S * TON_TO_KG / RAW_FEED_DENSITY_KG_M3,
                "raw_feed_density_kg_m3": RAW_FEED_DENSITY_KG_M3,
                "digestate_density_kg_m3": DIGESTATE_DENSITY_KG_M3,
                "W": W,
                "F": F,
                "S": S,
                            "termination": "skipped_zero_flow"})
            continue

        # --- 적응형 변곡점 스윕: 조합이 실제로 바뀌는 epsilon 만 solve ---
        demand = {}
        for key, col in (("gas", DEMAND_GAS_COL), ("elec", DEMAND_ELEC_COL)):
            if col in df_run.columns:
                v = row.get(col)
                if v is not None and not pd.isna(v):
                    demand[key] = float(v)

        try:
            solved, sweep_mode = adaptive_epsilon_solves(W, F, S, demand=demand)
        except Exception as e:
            solved, sweep_mode = [], "error"
            print(f"  [adaptive sweep error] row={row_idx}: {e}")

        # --- [v76] 설계 전환 임계 TAC: 조합이 바뀐 구간만 이분탐색 ---
        if DESIGN_SWITCH_ENABLED and solved:
            try:
                _hint = None
                for _r in solved:
                    _ex = _r[2] if isinstance(_r[2], dict) else {}
                    if _ex.get("grid_hint") is not None:
                        _hint = _ex["grid_hint"]
                        break
                _sw, _extra = find_design_switches(
                    W, F, S, solved, demand=demand, grid_hint=_hint,
                )
                for _s in _sw:
                    switch_records.append({**region_info, **_s})
                if _extra:
                    solved = sorted(solved + _extra, key=lambda r: r[0])
            except Exception as e:
                print(f"  [design switch error] row={row_idx}: {e}")

        if not solved:
            results.append({
                **region_info,
                "epsilon": None,
                **raw,
                "W_raw_t_d": W,
                "F_raw_t_d": F,
                "S_raw_t_d": S,
                "W_converted_m3_d": W * TON_TO_KG / RAW_FEED_DENSITY_KG_M3,
                "F_converted_m3_d": F * TON_TO_KG / RAW_FEED_DENSITY_KG_M3,
                "S_converted_m3_d": S * TON_TO_KG / RAW_FEED_DENSITY_KG_M3,
                "raw_feed_density_kg_m3": RAW_FEED_DENSITY_KG_M3,
                "digestate_density_kg_m3": DIGESTATE_DENSITY_KG_M3,
                "W": W,
                "F": F,
                "S": S,
                            "termination": "probe_infeasible"})
            continue

        # [v62] CAPEX-CH4 dominated 점 제거.
        #  epsilon 오름차순에서 CH4가 실제로 증가하는 점만 Pareto frontier로 남긴다.
        #  (infeasible 등 목적값이 없는 레코드는 '실현가능 임계점' 정보로 유지)
        if DROP_DOMINATED_POINTS:
            kept = []
            best_obj = None
            n_dropped = 0
            for rec in sorted(solved, key=lambda r: r[0]):
                ex = rec[2] if isinstance(rec[2], dict) else {}
                ov = ex.get("objective_value")
                if ov is None:
                    kept.append(rec)
                    continue
                if best_obj is None or ov > best_obj + 1.0e-6:
                    best_obj = ov
                    kept.append(rec)
                else:
                    n_dropped += 1
            if n_dropped:
                print(f"    [dominated 제거] {n_dropped}점")
            solved = kept

        # 연속으로 동일한 조합이 나오면 임계 epsilon(첫 등장)만 남긴다.
        if DEDUPE_CONSECUTIVE:
            deduped = []
            _prev_sig = object()  # 절대 같지 않은 초기값
            for rec in solved:
                if rec[3] != _prev_sig:
                    deduped.append(rec)
                    _prev_sig = rec[3]
            solved = deduped

        # =========================================================
        # [v78] grid_hint 미적용 행은 pareto_valid 에서 떨어뜨린다.
        #   solve 자체는 'optimal' 로 끝나지만 이중선형 근사오차가 route 를
        #   뒤집을 수 있어 신뢰할 수 없다. 값은 Raw_results 에 남기되
        #   Pareto_results / LCOE 최소점 선정에서는 제외된다.
        # =========================================================
        _n_nohint = 0
        for _rec in solved:
            _ex = _rec[2] if isinstance(_rec[2], dict) else None
            if _ex is None:
                continue
            _ex.setdefault("grid_hint_applied", 1)
            if not _ex["grid_hint_applied"]:
                _n_nohint += 1
                if "pareto_valid" in _ex:
                    _ex["pareto_valid"] = 0
        if _n_nohint:
            print(
                f"    [주의] grid_hint 미적용 {_n_nohint}점 -> pareto_valid=0"
            )

        print(
            f"  row={row_idx} [{sweep_mode}] "
            f"epsilon {solved[0][0]:,.0f} ~ {solved[-1][0]:,.0f}, "
            f"고유조합 {len(solved)}개"
        )

        for (epsilon, term, extra, _sig) in solved:
            case_no += 1
            print(f"[{case_no}/~{total_cases_est}] row={row_idx}, epsilon={epsilon:,.0f}, W={W}, F={F}, S={S}")

            base_record = {
                **region_info,
                "epsilon": epsilon,
                **raw,
                "W_raw_t_d": W,
                "F_raw_t_d": F,
                "S_raw_t_d": S,
                "W_converted_m3_d": W * TON_TO_KG / RAW_FEED_DENSITY_KG_M3,
                "F_converted_m3_d": F * TON_TO_KG / RAW_FEED_DENSITY_KG_M3,
                "S_converted_m3_d": S * TON_TO_KG / RAW_FEED_DENSITY_KG_M3,
                "raw_feed_density_kg_m3": RAW_FEED_DENSITY_KG_M3,
                "digestate_density_kg_m3": DIGESTATE_DENSITY_KG_M3,
                "W": W,
                "F": F,
                "S": S,
            }

            results.append({**base_record, "termination": term, **extra})

    out_df = pd.DataFrame(results)

    # ---------------------------------------------------------
    # [검증] 최종 결과 자체 점검
    # ---------------------------------------------------------
    print("\n[검증]")
    if "objective_value" in out_df.columns:
        viol = 0
        for _r, g in out_df.dropna(subset=["objective_value"]).groupby("row_index"):
            g = g.sort_values("epsilon")
            v = g["objective_value"].to_numpy(dtype=float)
            viol += int((v[1:] < np.maximum.accumulate(v)[:-1] - 1.0e-6).sum())
        print(f"  - Pareto 단조성 위반: {viol}건 (v48: 20건)")

    if "TAC_usd_y" in out_df.columns and "epsilon" in out_df.columns:
        cc = out_df.dropna(subset=["TAC_usd_y", "epsilon"])
        print(
            "  - TAC <= epsilon 위반: "
            f"{int((cc['TAC_usd_y'] - cc['epsilon'] > 1.0e-6).sum())}건"
        )

    # ---------------------------------------------------------
    # [v76] 열역학 검산 경고
    #   메탄 생산량이 '원료 물리 COD 로부터의 이론 최대'를 넘으면 경고한다.
    #   원인은 NN 입력 COD(SURROGATE_COD_COEF=75.76)와 원료별 물리 COD
    #   (COD_ANI/FW/SL = 40.79/6.66/0.75)의 기준 불일치다.
    #   어느 쪽이 맞는지 확정되기 전이므로 값은 건드리지 않고 경고만 남긴다.
    #   -> 이 경고가 떠 있는 동안 생산량·LCOE 수치는 인용하면 안 된다.
    # ---------------------------------------------------------
    if "ch4_yield_exceeds_theoretical" in out_df.columns:
        _f = out_df["ch4_yield_exceeds_theoretical"].dropna().astype(float)
        _n_bad = int((_f > 0.5).sum())
        if _n_bad:
            _ratio = out_df.get("physical_CH4_yield_ratio_to_theoretical")
            _rmax = float(_ratio.dropna().astype(float).max()) if _ratio is not None else float("nan")
            print(
                f"  ! [경고] 메탄 수율이 물리 COD 이론상한 초과: {_n_bad}행 "
                f"(최대 {_rmax:.2f}배)"
            )
            print(
                "    NN 입력 COD(75.76 kg/m3)와 원료별 물리 COD"
                "(40.79/6.66/0.75)의 기준 불일치가 원인이다."
            )
            print(
                "    COD_FW / COD_SL 상수의 출처를 확인하기 전까지 "
                "생산량·LCOE 는 잠정치로만 취급할 것."
            )
        else:
            print("  - 메탄 수율 열역학 검산: 위반 없음")

    # [v78] grid_hint 미적용(= probe 실패) 행 집계
    if "grid_hint_applied" in out_df.columns:
        _gh = out_df["grid_hint_applied"].fillna(1).astype(float)
        _n = int((_gh < 0.5).sum())
        if _n:
            _rows = sorted(
                out_df.loc[_gh < 0.5, "row_index"].dropna().unique().tolist()
            )
            print(
                f"  ! [경고] grid_hint 미적용 {_n}행 (row_index {_rows}) "
                "-> pareto_valid=0 처리됨"
            )
            print(
                "    probe 가 실패한 지역이다. solve 는 optimal 이지만 "
                "이중선형 근사오차가 route 를 뒤집을 수 있어 신뢰 불가."
            )
            print(
                "    GAS_GRID_N 을 올리거나 해당 지역 probe 실패 원인을 "
                "확인한 뒤 재실행할 것."
            )
        else:
            print("  - grid_hint: 전 행 적용됨")

    for col, label in [
        ("pur_in_ch4_flow_approx_error", "CH4 이중선형 근사오차"),
        ("AD_mass_balance_error_kg_d", "AD 총질량 수지오차"),
        ("solids_balance_residual_kg_d", "AD 고형물 수지잔차(0이어야 정상)"),
        ("gen_cost_piecewise_error", "발전기 piecewise 오차"),
    ]:
        if col in out_df.columns:
            v = out_df[col].dropna().astype(float).abs()
            if len(v):
                print(f"  - {label}: max|.|={v.max():,.3f}")

    if "nn_all_inputs_in_training_range" in out_df.columns:
        v = out_df["nn_all_inputs_in_training_range"].dropna()
        w = out_df["nn_max_sigma_excess"].dropna().astype(float)
        print(
            f"  - NN 입력 학습범위 내: {int(v.sum())}/{len(v)}행, "
            f"최대 외삽 {w.max() if len(w) else 0:.2f}sigma"
        )

    if "ch4_yield_exceeds_theoretical" in out_df.columns:
        v = out_df["ch4_yield_exceeds_theoretical"].dropna()
        r = out_df["physical_CH4_yield_ratio_to_theoretical"].dropna().astype(float)
        if len(v):
            print(
                f"  - 실측COD 기준 메탄수율이 이론최대(0.35) 초과: "
                f"{int(v.sum())}/{len(v)}행, 최대 {r.max():.1f}배"
            )

    if "LCOB_usd_per_m3_CH4" in out_df.columns:
        v = out_df["LCOB_usd_per_m3_CH4"].dropna().astype(float)
        c = out_df["LCOB_converged"].dropna()
        if len(v):
            print(
                f"  - LCOB: {v.min():.3f} ~ {v.max():.3f} $/m3 CH4 "
                f"({v.min() * EXCHANGE_RATE_KRW_PER_USD:,.0f} ~ "
                f"{v.max() * EXCHANGE_RATE_KRW_PER_USD:,.0f} ￦/m3), "
                f"Dinkelbach 수렴 {int(c.sum())}/{len(c)}행"
            )

    if "surrogate_reliable" in out_df.columns:
        v = out_df["surrogate_reliable"].dropna()
        bad = out_df.loc[v[v == 0].index, "읍면동"].unique() if len(v) else []
        print(
            f"  - surrogate 신뢰구간(q_ad>={SURROGATE_MIN_Q_AD:.0f}) 내: "
            f"{int(v.sum())}/{len(v)}행"
            + (f", 구간 밖 지역: {list(bad)}" if len(bad) else "")
        )

    if "GEN_LCOE_usd_per_kWh" in out_df.columns:
        v = out_df["GEN_LCOE_usd_per_kWh"].dropna().astype(float)
        if len(v):
            print(
                f"  - 발전 시나리오 LCOE: {v.min():.3f} ~ {v.max():.3f} $/kWh "
                f"({v.min() * EXCHANGE_RATE_KRW_PER_USD:,.0f} ~ "
                f"{v.max() * EXCHANGE_RATE_KRW_PER_USD:,.0f} ￦/kWh)"
            )

    if "preferred_route_by_cost_price_ratio" in out_df.columns:
        vc = out_df["preferred_route_by_cost_price_ratio"].dropna().value_counts()
        print(f"  - 원가/판매가 비율 기준 우세 route: {dict(vc)}")

    if "capex_per_m3d_CH4" in out_df.columns:
        v = out_df["capex_per_m3d_CH4"].dropna().astype(float)
        if len(v):
            print(
                f"  - 메탄 1 m3/d 당 CAPEX: "
                f"{v.min():,.0f} ~ {v.max():,.0f} USD"
            )

    # =========================================================
    # [v65] 결과 후처리 및 Excel 저장
    # =========================================================
    # Raw_results
    #   - 모든 epsilon 결과를 그대로 보존한다.
    # Purification_Pareto
    #   - 정제 1·2단계가 모두 optimal인 해만 사용한다.
    #   - 실제 정제 CAPEX가 증가할 때 CH4 생산량이 엄격히 증가하는 점만 남긴다.
    # Generator_Pareto
    #   - 발전 1·2단계가 모두 optimal인 해만 사용한다.
    #   - 실제 발전 CAPEX가 증가할 때 발전량이 엄격히 증가하는 점만 남긴다.
    # Route_Comparison
    #   - 같은 CAPEX 상한에서 두 route가 모두 유효한 결과를 나란히 비교한다.

    def _strict_pareto_front(
        frame,
        valid_col,
        cost_col,
        production_col,
        cost_tol=1.0e-6,
        production_tol=1.0e-9,
    ):
        """지역별 비용 최소화-생산량 최대화의 엄격한 비지배점만 반환한다."""
        required = {"row_index", valid_col, cost_col, production_col}
        if frame.empty or not required.issubset(frame.columns):
            return frame.iloc[0:0].copy()

        work = frame.loc[frame[valid_col] == 1].copy()
        work[cost_col] = pd.to_numeric(work[cost_col], errors="coerce")
        work[production_col] = pd.to_numeric(
            work[production_col], errors="coerce"
        )
        work = work.dropna(subset=[cost_col, production_col])
        if work.empty:
            return work

        keep_indices = []
        for _, group in work.groupby("row_index", sort=False):
            # 동일 또는 거의 동일한 실제 CAPEX에서는 생산량이 가장 큰 행만 후보로 둔다.
            group = group.sort_values(
                [cost_col, production_col], ascending=[True, False]
            )
            candidates = []
            last_cost = None
            for idx, row in group.iterrows():
                cost = float(row[cost_col])
                if last_cost is None or abs(cost - last_cost) > cost_tol:
                    candidates.append(idx)
                    last_cost = cost

            best_production = -float("inf")
            for idx in candidates:
                production = float(work.at[idx, production_col])
                if production > best_production + production_tol:
                    keep_indices.append(idx)
                    best_production = production

        result = work.loc[keep_indices].copy()
        return result.sort_values(["row_index", cost_col]).reset_index(drop=True)

    # [v66] route 는 모델이 골랐으므로 Pareto 는 하나다.
    #  (CAPEX 최소, 판매에너지 최대) 의 비지배점만 남긴다.
    # [v75] 가성비 최적점 = front 위에서 LCOE 가 최소인 행.
    #  LCOE 최소 설계는 반드시 (TAC 최소, E 최대) Pareto front 위에 있으므로
    #  front 를 훑어 최소를 고르면 그것이 전역 최적이다(S14 주석의 증명).
    out_df["is_min_lcoe_point"] = 0
    _lc = "levelized_energy_cost_usd_per_kWh"
    if _lc in out_df.columns and "row_index" in out_df.columns:
        _ok = out_df[(out_df.get("pareto_valid", 1) == 1) & out_df[_lc].notna()]
        for _ri, _g in _ok.groupby("row_index"):
            out_df.loc[_g[_lc].idxmin(), "is_min_lcoe_point"] = 1

    # [v75] Pareto 축 = (TAC 최소, 유효공급에너지 최대)
    pareto_df = _strict_pareto_front(
        out_df,
        valid_col="pareto_valid",
        cost_col="TAC_usd_y",
        production_col="useful_energy_kWh_d",
    )
    pareto_df["pareto_type"] = "TAC_min_vs_useful_energy_max"

    # KEY_COLUMNS 만 뽑은 요약본. 없는 열은 경고만 하고 건너뛴다.
    key_cols = [c for c in KEY_COLUMNS if c in out_df.columns]
    missing_keys = [c for c in KEY_COLUMNS if c not in out_df.columns]
    if missing_keys:
        print(f"  [주의] KEY_COLUMNS 에 없는 열 {len(missing_keys)}개 무시: "
              f"{missing_keys}")
    key_df = out_df[key_cols].copy() if key_cols else None

    with pd.ExcelWriter(output_excel, engine="openpyxl") as _xw:
        # 요약을 첫 시트로 두어 엑셀을 열면 바로 보이게 한다.
        if key_df is not None:
            key_df.to_excel(_xw, sheet_name="Key_results", index=False)
        out_df.to_excel(
            _xw, sheet_name="Raw_results", index=False
        )
        pareto_df.to_excel(_xw, sheet_name="Pareto_results", index=False)
        # [v76] 설계 전환 임계 TAC — 논문 Table 2 가 되는 시트
        switch_df = pd.DataFrame(switch_records)
        if not switch_df.empty:
            switch_df = switch_df.sort_values(
                ["row_index", "threshold_TAC_usd_y"]
            ).reset_index(drop=True)
        switch_df.to_excel(_xw, sheet_name="Design_switch", index=False)

        # 가독성: 첫 행 고정, 자동필터, 지나치게 긴 열 폭 제한
        for ws in _xw.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for column_cells in ws.columns:
                values = [
                    "" if cell.value is None else str(cell.value)
                    for cell in column_cells[:200]
                ]
                width = min(max((len(v) for v in values), default=8) + 2, 35)
                ws.column_dimensions[column_cells[0].column_letter].width = width

    n_invalid = (
        int((out_df["pareto_valid"] != 1).sum())
        if "pareto_valid" in out_df.columns else len(out_df)
    )
    print(
        "  - 저장 시트:\n"
        + (f"      Key_results    : {len(key_df)}행 x {len(key_cols)}열 (요약)\n"
           if key_df is not None else "")
        + f"      Raw_results    : {len(out_df)}행 x {len(out_df.columns)}열\n"
        f"      Pareto_results : {len(pareto_df)}행 (비유효 {n_invalid}행 제외)\n"
        f"      Design_switch  : {len(switch_records)}행 (설계 전환 임계 TAC)"
    )
    # [v75] 지역 요약: 읍면동별 가성비 최적 설계
    if "is_min_lcoe_point" in out_df.columns:
        _b = out_df[out_df["is_min_lcoe_point"] == 1]
        if not _b.empty:
            print("  - 가성비 최적점 (읍면동별 LCOE 최소):")
            for _, r in _b.iterrows():
                _t = r.get("selected_purification_tech") or r.get("selected_generator")
                _l = r.get("levelized_energy_cost_usd_per_kWh")
                print(
                    f"      {str(r.get('읍면동', '')):<12}"
                    f"{str(r.get('selected_route', '')):<13}{str(_t):<14}"
                    f"LCOE {_l:.4f} $/kWh ({_l * EXCHANGE_RATE_KRW_PER_USD:,.0f} W/kWh) | "
                    f"TAC {r.get('TAC_usd_y', 0):,.0f} $/y, "
                    f"공급 {r.get('useful_energy_kWh_d', 0):,.0f} kWh/d"
                )

    if "selected_route" in out_df.columns:
        vc = out_df["selected_route"].dropna().value_counts().to_dict()
        print(f"  - 모델이 고른 route: {vc}")
    if "breaks_even" in out_df.columns:
        b = out_df["breaks_even"].dropna()
        print(f"  - 판매가로 원가 회수(비율<=1): {int(b.sum())}/{len(b)}행")

    print("\nDONE")
    print(f"saved: {output_excel}")
    print(f"rows: {len(out_df)}")
    print(f"elapsed_sec: {time.time() - t0:.2f}")

    return out_df


class _Tee:
    """여러 스트림에 동시에 write. (콘솔 + 지역별 로그파일 동시 출력용)"""

    def __init__(self, *streams):
        self.streams = streams

    def write(self, data):
        for s in self.streams:
            s.write(data)
            s.flush()
        return len(data)

    def flush(self):
        for s in self.streams:
            s.flush()


def _sanitize_filename(name):
    """파일명으로 못 쓰는 문자를 '_' 로 치환."""
    bad = '\\/:*?"<>|'
    return "".join("_" if c in bad else c for c in str(name)).strip()


def run_all_regions():
    """A열(REGION_COL)의 지역별로 나눠 순서대로 실행한다.
    지역이 끝날 때마다
      - 결과 엑셀:  {OUTPUT_PREFIX}_{지역명}.xlsx
      - stdout 로그: {지역명}.txt
    를 각각 저장한다."""
    if not os.path.exists(INPUT_EXCEL):
        raise FileNotFoundError(f"입력 엑셀 파일을 찾을 수 없습니다: {INPUT_EXCEL}")

    df = pd.read_excel(INPUT_EXCEL)

    if REGION_COL not in df.columns:
        raise KeyError(
            f"지역 분할 열 '{REGION_COL}'이 입력 엑셀에 없습니다. "
            f"사용 가능한 열: {list(df.columns)}"
        )

    region_series = df[REGION_COL].astype(str).str.strip()

    # 등장 순서를 보존하며 중복 제거 (엑셀 A열 순서 그대로)
    all_regions = list(dict.fromkeys(region_series))

    if REGIONS_TO_RUN is not None:
        wanted = [str(r).strip() for r in REGIONS_TO_RUN]
        missing = [r for r in wanted if r not in all_regions]
        if missing:
            raise ValueError(
                f"REGIONS_TO_RUN 에 지정한 지역이 A열에 없습니다: {missing}\n"
                f"사용 가능한 지역: {all_regions}"
            )
        regions = [r for r in all_regions if r in wanted]
    else:
        regions = all_regions

    print(f"[전체] 총 {len(regions)}개 지역 실행 예정: {regions}")

    t_all = time.time()
    summary = []

    for i, region in enumerate(regions, 1):
        df_region = df[region_series == region].copy()

        if MAX_ROWS is not None:
            df_region = df_region.head(MAX_ROWS).copy()

        safe = _sanitize_filename(region)
        output_excel = f"{OUTPUT_PREFIX}_{safe}.xlsx"
        # OUTPUT_PREFIX 에서 버전 태그(예: v79)만 떼어 txt 파일명에 붙인다.
        # 이전 버전 로그(원주시.txt 등)를 덮어쓰지 않도록 원주시_v79.txt 형태로 저장.
        version_tag = OUTPUT_PREFIX.split("_")[-1]  # "ModelOutPut_v79" -> "v79"
        log_path = f"{safe}_{version_tag}.txt"

        header = (
            f"\n{'=' * 70}\n"
            f"[{i}/{len(regions)}] 지역: {region}  "
            f"({len(df_region)}개 읍면동)\n"
            f"  -> 엑셀: {output_excel}\n"
            f"  -> 로그: {log_path}\n"
            f"{'=' * 70}"
        )

        # stdout 을 콘솔 + 지역별 txt 에 동시 기록
        if LOG_PER_REGION:
            logf = open(log_path, "w", encoding="utf-8")
            out_target = _Tee(sys.__stdout__, logf)
        else:
            logf = None
            out_target = sys.__stdout__

        status = "OK"
        try:
            with contextlib.redirect_stdout(out_target):
                print(header)
                try:
                    run_region_epsilon_sweep(df_region, output_excel)
                except Exception as e:
                    status = f"FAILED: {e}"
                    print(f"\n[ERROR] '{region}' 실행 중 예외 발생: {e}")
                    traceback.print_exc(file=out_target)
        finally:
            if logf is not None:
                logf.close()

        summary.append((region, status, output_excel, log_path))
        print(f"[{i}/{len(regions)}] {region} 완료 ({status})")

    print(f"\n{'#' * 70}")
    print("[전체 요약]")
    for region, status, output_excel, log_path in summary:
        print(f"  - {region:<8} {status:<12} {output_excel} / {log_path}")
    print(f"총 소요: {time.time() - t_all:.1f} s")
    print(f"{'#' * 70}")


if __name__ == "__main__":
    # 명령행 인자로 지역명을 주면 그 지역만 실행 (REGIONS_TO_RUN 을 덮어씀).
    #   예) python3 Model_v69.py 동해시 태백시
    # 인자가 없으면 REGIONS_TO_RUN(None 이면 전체)을 따른다.
    if len(sys.argv) > 1:
        REGIONS_TO_RUN = sys.argv[1:]
    run_all_regions()
