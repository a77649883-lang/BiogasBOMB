# -*- coding: utf-8 -*-
"""
바이오가스 플랜트 TEA 계산기 (LCOB / LCOE)  — 순차 계산판(v2)
====================================================================
· 입력값(발생량·설비 선택·모드)을 파일 맨 위에서 지정하면,
  아래로 쭉 계산되어 LCOB(정제) 또는 LCOE(발전)를 산출한다.
· 변수명은 Nomenclature(SJNA Meeting) 표기를 따른다.
      LCOB = (CAPEX_total x f_crf + OPEX_total) / (Q_CH4 x f_util)      ...(1)
      CAPEX_total = C_equipment + C_e_install + C_pipe + C_instrumentation
                    + C_electrical + C_building + C_yard + C_facility
                    + C_engineering + C_construction + C_legal
                    + C_contractor + C_contingency                      ...(2)
      OPEX_total  = OPEX_fixed + OPEX_var                               ...(4)
      OPEX_fixed  = O_salary + O_benefit + O_maintenance + O_lab + O_insurance ...(5)
      OPEX_var    = O_chemical + O_electricity                          ...(6)
      발전가능량   = 총가스 x 발열량 x 발전효율 / 단위전력당발열량        ...(7)
· 설비 구매비(C_equipment)로부터 표준사업비(파생) = C_equipment / 0.229 를
  구하고, 배관·설치·전기 등 부대비용을 그 %로 계산한다.
"""

import math

# ====================================================================
# 1. 입력값 (INPUT)  ── 여기 값만 바꿔서 실행
# ====================================================================
AM = 90.0            # 가축분뇨 발생량 (ton/day)
SL = 8.0             # 하수슬러지 발생량 (ton/day)
FW = 2.0             # 음식물류폐기물 발생량 (ton/day)
TOTAL_GAS = 3500.0   # 총 가스 발생량 (m3/day)
BIOGAS    = 2050.0   # 메탄 생산량   (m3/day)
TN  = 1500.0         # 폐수 총질소 TN (mg/L)
HRT = 40.0           # 혐기성 소화조 체류시간 (day)
MODE = "정제"         # "정제"(가스판매 -> LCOB) 또는 "발전"(전기 -> LCOE)

# 공정별 설비 선택 (옵션 번호)
OPT_AM_PRETREAT   = 1   # 1: screen+cyclone / 2: +drum filter
OPT_SLUDGE_THICK  = 1   # 1: gravity thickening / 2: rotary drum thickening
OPT_FW_SCREEN_1   = 1   # 1: drum / 2: screw / 3: step   (1계열 스크린)
OPT_FW_SCREEN_2   = 1   # (2계열 스크린)
OPT_FW_SCREEN_MID = 1   # (파쇄 후 스크린)
OPT_FW_SHRED_1    = 1   # 1: screening / 2: rotary / 3: two-shaft / 4: hammer (1차 파쇄)
OPT_FW_SHRED_2    = 1   # (2차 파쇄)
OPT_GAS_STORAGE   = 1   # 1: membrane / 2: piston deck / 3: steel / 4: wet
OPT_DESULF        = 1   # 1: wet / 2: dry
OPT_DEHUM         = 1   # 1: cooling / 2: compressed cooling
OPT_UPGRADE       = 1   # 1: PSA / 2: Membrane / 3: Wet scrubber   (정제 모드만)
OPT_WWT           = 1   # 1: Anammox / 2: A2O / 3: Bardenpho
OPT_DEWATER       = 1   # 1: belt press / 2: screw press
OPT_GENERATOR     = 1   # 1: gas engine / 2: gas turbine / 3: fuel cell (발전 모드만)

# ====================================================================
# 2. 파라미터 (PARAMETERS)
# ====================================================================
EXCHANGE_RATE = 1434       # 환율 (원/$)
DESIGN_MARGIN = 1.2        # 설계여유계수
DEFAULT_SF = 0.6           # 설비 공통 Scale Factor
STORAGE_DAYS = 3           # 저장조 설계 일수
FOOD_WASTE_DENSITY = 0.8   # 음식물 밀도 (ton/m3)
OPERATION_FACTOR = 3       # 전처리 8시간 가동 보정 (=24/8), 슬러지 제외
MIXER_COVER_VOL = 200      # 교반기 1대 커버 부피 (m3)
DIGESTER_TRAINS = 2        # 소화조 계열 수

# LCOB/LCOE 파라미터
f_wacc = 0.045             # 가중평균자본비용 (할인율)
L = 30                     # 플랜트 수명 (year)
f_util = 1.0               # 가동률
HEATING_FRACTION = 0.35    # 생산 메탄의 35%를 소화조 가온에 사용 (정제 판매 메탄 차감)
CH4_HV_MJ_PER_M3 = 35.77   # 정제 후 메탄 발열량 (MJ/m3) — 에너지단위 환산용
MJ_PER_KWH = 3.6           # 1 kWh = 3.6 MJ

# 정부 보조금: 회귀식 표준사업비의 60% (실제 CAPEX가 아니라 표준사업비 기준, 고정)
SUBSIDY_RATE = 0.60

# 고정 O&M 계수 (WaterTap Zero-Order, 약품 제외)
OM_SALARY_RATE     = 0.001   # 급여
OM_BENEFIT_RATE    = 0.90    # 복리후생 = 급여의 90%
OM_MAINTENANCE_RATE= 0.008   # 유지보수
OM_LAB_RATE        = 0.003   # 시험
OM_INSURANCE_RATE  = 0.002   # 보험·세금
# O&M 적용 base 비율 = 구입기기비(0.229) + 배관(0.073) + 전기설비(0.046) = 0.348
PURCHASED_EQUIP_RATIO = 0.229
STD_RATES = {   # 표준사업비 대비 부대비용 항목 % (합 0.771)
    "C_pipe": 0.073, "C_e_install": 0.083, "C_instrumentation": 0.092,
    "C_electrical": 0.046, "C_building": 0.046, "C_yard": 0.018,
    "C_facility": 0.138, "C_engineering": 0.118, "C_construction": 0.092,
    "C_legal": 0.018, "C_contractor": 0.018, "C_contingency": 0.029,
}
OM_CAPEX_FRACTION = PURCHASED_EQUIP_RATIO + STD_RATES["C_pipe"] + STD_RATES["C_electrical"]  # 0.348

# 폐수 물질수지 (가스 질량 + 고형물 포집 기반)
TS_FEED_FRACTION = 0.05
CH4_GAS_DENSITY = 0.716
CO2_GAS_DENSITY = 1.977
CH4_YIELD_M3_PER_KGCOD = 0.35
COD_PER_VS = 1.5
SOLIDS_CAPTURE = 0.95
CAKE_TS = 0.20

# 폐수처리 반응조 비용식 (CEPCI 물가보정)
CEPCI_ANAEROBIC = 357.6    # 혐기조·무산소조 (1990)
CEPCI_AEROBIC   = 386.5    # 폭기조 (1997)
CEPCI_CLARIFIER = 358.2    # 2차 침전지 (1992)
CEPCI_CURRENT   = 797.9    # 현재 (2023)
M3_DAY_PER_MGD  = 3785.41  # 1 MGD = 3785.41 m3/day
ANAMMOX_ETC_KRW_PER_KGN = 508.0   # Anammox 변동 OPEX (원/kg-N)

# 약품 단가·원단위
METHANOL_USD_PER_TON = 340.0
LIME_KRW_PER_KG = 950.0
ALUM_KRW_PER_KG = 778.8
METHANOL_TON_PER_TON_N = 2.47
LIME_TON_PER_TON_N = 7.14
ALUM_MG_PER_L = 100.0
DESULF_USD_PER_M3  = {1: 0.0225, 2: 0.014}                 # wet / dry
UPGRADE_USD_PER_M3 = {1: 0.144, 2: 0.194, 3: 0.162}        # PSA / Membrane / scrubber

# 전기요금 (계절 x 부하 TOU)
#   각 계절: (일수, [(24h설비 가동h, 8h설비 가동h, 단가 원/kWh) x 3])
ELEC_SEASONS = [
    (92,  [(10, 0, 115.1), (8, 7, 163.2), (6, 1, 216.6)]),   # 여름
    (153, [(10, 0, 115.1), (8, 7, 132.1), (6, 1, 142.6)]),   # 봄가을
    (120, [(10, 0, 122.4), (8, 5, 163.4), (6, 3, 193.4)]),   # 겨울
]
ELEC_BASE_CHARGE = 9810 * 300 * 12   # 기본요금 (원/year, 300kW 계약)
ELEC_CLIMATE_RATE = 9                # 기후환경요금 (원/kWh)
ELEC_FUEL_ADJ_RATE = 5               # 연료비조정요금 (원/kWh)

# 발전 설비 (식 7, Table 1 범위 중앙값)
CH4_HEATING_VALUE_KCAL = 5550.0      # 바이오가스 발열량 (kcal/m3)
ELEC_HEATING_VALUE_KCAL = 860.0      # 단위 전력당 발열량 (kcal/kWh)
GEN_OPERATING_HOURS = 24.0
GEN_SPECS = [   # (이름, 발전효율, Cost $/kW, O&M $/kWh)
    ("gas engine",  (0.28 + 0.40) / 2, (465 + 1600) / 2,  (0.010 + 0.025) / 2),
    ("gas turbine", (0.20 + 0.35) / 2, (1100 + 2000) / 2, (0.008 + 0.010) / 2),
    ("fuel cell",   (0.36 + 0.50) / 2, (3800 + 5280) / 2, (0.004 + 0.019) / 2),
]

# ====================================================================
# 3. 공통 함수
# ====================================================================
def sf_cost(ref_cost, ref_vol=None, vol=None):
    """Scale Factor 비용: vol<ref_vol이면 ref_cost, 아니면 ref_cost x (vol/ref_vol)^0.6.
    ref_vol 없으면(부피 무관) ref_cost 고정."""
    if ref_vol is None or vol is None:
        return ref_cost
    if vol < ref_vol:
        return ref_cost
    return ref_cost * (vol / ref_vol) ** DEFAULT_SF

def wwt_anaerobic(vol):   # 혐기조·무산소조 (부피 m3)
    return 1246 * vol ** 0.71 * (CEPCI_CURRENT / CEPCI_ANAEROBIC)

def wwt_aerobic(vol):     # 폭기조 (부피 m3)
    return 1114.1 * vol ** 0.8324 * (CEPCI_CURRENT / CEPCI_AEROBIC)

def wwt_clarifier(flow):  # 2차 침전지 CAPEX (유량 MGD): ln(Y)=12.834601+0.688675lnX+0.035432(lnX)^2
    x = flow / M3_DAY_PER_MGD
    if x <= 0:
        return 0.0
    lnx = math.log(x)
    return math.exp(12.834601 + 0.688675 * lnx + 0.035432 * lnx ** 2) * (CEPCI_CURRENT / CEPCI_CLARIFIER)

def wwt_clarifier_om(flow):  # 2차 침전지 O&M ($/yr): ln(Y)=10.197762+0.339952lnX+0.015822(lnX)^2
    x = flow / M3_DAY_PER_MGD
    if x <= 0:
        return 0.0
    lnx = math.log(x)
    return math.exp(10.197762 + 0.339952 * lnx + 0.015822 * lnx ** 2) * (CEPCI_CURRENT / CEPCI_CLARIFIER)

def anammox_capex(vol):   # Anammox 총공사비 전액 ($, 부대 포함)
    return (0.3148 * vol + 9.17) * 1e8 / EXCHANGE_RATE

def elec_tariff(p24, p8):
    """평균전력(P24,P8) -> (전력량요금 원, 연간 소비전력 kWh). 계절·부하 TOU 합산."""
    energy = kwh = 0.0
    for days, loads in ELEC_SEASONS:
        for h24, h8, price in loads:
            k = p24 * h24 * days + p8 * h8 * days
            energy += k * price
            kwh += k
    return energy, kwh

# ====================================================================
# 4. 계산 (위 -> 아래 순차 진행)
# ====================================================================
def calculate():
    total_ton = AM + SL + FW
    fw_m3 = FW / FOOD_WASTE_DENSITY

    # ---- 폐수 물질수지 (탈수·폐수 유량 산정) ----
    co2 = max(0.0, TOTAL_GAS - BIOGAS)
    biogas_mass = BIOGAS * CH4_GAS_DENSITY + co2 * CO2_GAS_DENSITY
    digestate_mass = total_ton * 1000.0 - biogas_mass
    Y4 = digestate_mass / 1000.0                                 # 소화 후 슬러지 (m3/day)
    feed_dry = total_ton * TS_FEED_FRACTION * 1000.0
    vs_destroyed = (BIOGAS / CH4_YIELD_M3_PER_KGCOD) / COD_PER_VS
    cake_mass = max(0.0, feed_dry - vs_destroyed) * SOLIDS_CAPTURE / CAKE_TS
    Y5 = cake_mass / 1000.0                                      # 탈수 케이크 (m3/day)
    AE2 = (digestate_mass - cake_mass) / 1000.0                  # 폐수 유량 (m3/day)
    N_ton = TN * AE2 * 1e-6                                      # 질소부하 (ton N/day)

    def storage_vol(x):
        return x * STORAGE_DAYS * DESIGN_MARGIN

    # 공정별 누적 상자: equip(구매장비비) / anammox(all-inc) / e24 / e8(kWh) / chem($/yr)
    P = {}
    def box():
        return {"equip": 0.0, "anammox": 0.0, "e24": 0.0, "e8": 0.0, "chem": 0.0}
    def add(a, cost=0.0, e=0.0, h8=False, chem=0.0, all_inc=False, n=1):
        for _ in range(n):
            a["anammox" if all_inc else "equip"] += cost
            a["e8" if h8 else "e24"] += e
            a["chem"] += chem

    def fw_screen(a, opt):     # 음식물 스크린 (8h)
        if opt == 1:   add(a, cost=sf_cost(17000, 48, fw_v), e=0.157 * FW, h8=True)   # drum
        elif opt == 2: add(a, cost=sf_cost(5500, 480, fw_v), e=0.13 * FW, h8=True)    # screw
        else:          add(a, cost=sf_cost(1800, 96, fw_v), e=0.0, h8=True)           # step
    def fw_shred(a, opt, hammer_rv):   # 음식물 파쇄기 (8h)
        if opt == 1:   add(a, cost=sf_cost(2000, 552, fw_v), e=1.1 * FW, h8=True)     # screening
        elif opt == 2: add(a, cost=sf_cost(9000, 12, fw_v), e=34 * FW, h8=True)       # rotary
        elif opt == 3: add(a, cost=sf_cost(23700, 72, fw_v), e=3.1 * FW, h8=True)     # two-shaft
        else:          add(a, cost=sf_cost(7500, hammer_rv, fw_v), e=31 * FW, h8=True)# hammer

    # ---- 공정 1: 반입 ----
    a = box()
    if AM > 0:
        add(a, cost=79.2, e=0.177 * 24, n=2); add(a, cost=1019, n=2)      # air curtain·shutter x2
    if SL > 0:
        add(a, cost=79.2, e=0.177 * 24); add(a, cost=1019)
    if FW > 0:
        add(a, cost=79.2, e=0.177 * 24); add(a, cost=1019)
    if AM > 0:
        add(a, cost=sf_cost(16783, 50, storage_vol(AM)), n=2)             # AM storage x2
    if SL > 0:
        add(a, cost=sf_cost(16783, 50, storage_vol(SL)))
    if FW > 0:
        add(a, cost=sf_cost(2000, 2, storage_vol(fw_m3)))                 # feeding storage hopper
    if AM > 0:
        add(a, cost=sf_cost(4500, 360, AM), e=0.0165 * AM, n=2)           # conveyor x2
    if SL > 0:
        add(a, cost=sf_cost(4500, 360, SL), e=0.0165 * SL)
    if FW > 0:
        add(a, cost=sf_cost(4500, 360, FW), e=0.0165 * FW, n=2)
    P["반입"] = a

    # ---- 공정 2: 전처리 ----  (AM·FW는 vol_i x3, 전력은 원래 발생량)
    a = box()
    am_v = AM * OPERATION_FACTOR
    sl_v = SL
    fw_v = fw_m3 * OPERATION_FACTOR
    if AM > 0:
        add(a, cost=sf_cost(3000, 72, am_v), e=0.5 * AM, h8=True)         # screen
        add(a, cost=sf_cost(40000, 14400, am_v), e=0.29 * AM, h8=True)    # cyclone
        if OPT_AM_PRETREAT == 2:
            add(a, cost=sf_cost(3000, 72, am_v), e=0.5 * AM, h8=True)     # drum filter
    if SL > 0:
        if OPT_SLUDGE_THICK == 1:
            add(a, cost=sf_cost(4000, 5.6, sl_v), e=0.33 * SL)            # gravity (24h)
        else:
            add(a, cost=sf_cost(12000, 120, sl_v), e=0.14 * SL)          # rotary drum
    if FW > 0:
        fw_screen(a, OPT_FW_SCREEN_1)
        fw_screen(a, OPT_FW_SCREEN_2)
        fw_shred(a, OPT_FW_SHRED_1, 19.2)
        fw_screen(a, OPT_FW_SCREEN_MID)
        fw_shred(a, OPT_FW_SHRED_2, 19)
        add(a, cost=sf_cost(10000, 2.4, fw_v), e=37 * FW, h8=True, n=2)   # sorting machine x2
    add(a, cost=2000)                                                     # waste discharge hopper (공용)
    n_waste = sum(1 for x in (AM, SL, FW) if x > 0)
    add(a, cost=4500, e=0.0165 * FW, n=n_waste)                          # waste conveyor (원료계열당 1)
    if AM > 0: add(a, cost=2000)                                          # feeding hopper
    if SL > 0: add(a, cost=2000)
    if FW > 0: add(a, cost=2000)
    if AM > 0: add(a, cost=sf_cost(4500, 360, AM), e=0.0165 * AM, n=2)    # conveyor x2
    if SL > 0: add(a, cost=sf_cost(4500, 360, SL), e=0.0165 * SL, n=2)
    if FW > 0: add(a, cost=sf_cost(4500, 360, FW), e=0.0165 * FW, n=2)
    P["전처리"] = a

    # ---- 공정 3: 중간저장조 ----
    a = box()
    tank_vol = storage_vol(total_ton)
    mixers = math.ceil(tank_vol / MIXER_COVER_VOL)
    add(a, cost=2000, n=2)                                                # feeding hopper x2
    add(a, cost=sf_cost(16783, 50, tank_vol), n=2)                       # storage tank x2
    add(a, cost=660, e=0.19 * total_ton, n=mixers * 2)                   # 교반기 (저장조당 x2계열)
    add(a, cost=sf_cost(4500, 360, total_ton), e=0.0165 * total_ton, n=2)  # conveyor x2
    P["중간저장조"] = a

    # ---- 공정 4: 혐기성 소화조 ----
    a = box()
    dig_vol = total_ton * HRT * DESIGN_MARGIN / DIGESTER_TRAINS
    add(a, cost=sf_cost(1600000, 2300, dig_vol),
        e=3.35 * (total_ton / DIGESTER_TRAINS), n=DIGESTER_TRAINS)       # 소화조 x계열수
    P["혐기성소화조"] = a

    # ---- 공정 5: 가스 저장조 ----
    a = box()
    add(a, cost=sf_cost(7000, 50, TOTAL_GAS / 24 * 2.5))                 # 잉여가스연소기
    gs = {1: (8800, 100), 2: (9999, 50), 3: (9999, 50), 4: (5000, 50)}[OPT_GAS_STORAGE]
    add(a, cost=sf_cost(gs[0], gs[1], TOTAL_GAS / 4))                    # 저장조
    P["가스저장조"] = a

    # ---- 공정 6: 가스 전처리 ----  (발전 모드는 고질화 제외)
    a = box()
    g = TOTAL_GAS
    add(a, cost=sf_cost(16200, 1200, g), e=0.1 * g)                      # compressor
    if OPT_DESULF == 1:
        add(a, cost=sf_cost(50000, 960, g), chem=DESULF_USD_PER_M3[1] * g * 365)   # wet 탈황
    else:
        add(a, cost=sf_cost(5000, 24, g), chem=DESULF_USD_PER_M3[2] * g * 365)     # dry 탈황
    if OPT_DEHUM == 1:
        add(a, cost=sf_cost(10849, 2400, g), e=0.0033 * g)              # cooling
    else:
        add(a, cost=sf_cost(16200, 1200, g), e=0.1 * g)                # compressed cooling
    if MODE == "정제":
        up = {1: (500000, 2400), 2: (193567, 1200), 3: (50000, 960)}[OPT_UPGRADE]
        add(a, cost=sf_cost(up[0], up[1], g), chem=UPGRADE_USD_PER_M3[OPT_UPGRADE] * BIOGAS * 365)
    P["가스전처리"] = a

    # ---- 공정 7: 발전 (발전 모드만) ----  식(7)
    gen_kwh_day = 0.0
    if MODE == "발전":
        a = box()
        _, eff, cost_kw, om_kwh = GEN_SPECS[OPT_GENERATOR - 1]
        gen_kwh_day = TOTAL_GAS * CH4_HEATING_VALUE_KCAL * eff / ELEC_HEATING_VALUE_KCAL
        size_kw = gen_kwh_day / GEN_OPERATING_HOURS
        add(a, cost=size_kw * cost_kw, chem=om_kwh * gen_kwh_day * 365)   # 장비비 + O&M
        P["발전"] = a

    # ---- 공정 8: 탈수기 ----
    a = box()
    if OPT_DEWATER == 1:
        add(a, cost=sf_cost(20000, 3, Y4), e=0.1 * Y4)                   # belt press
    else:
        add(a, cost=sf_cost(5000, 0.072, Y4), e=0.019 * Y4)             # screw press
    add(a, cost=sf_cost(16783, 50, Y4 * DESIGN_MARGIN))                  # 저장탱크
    add(a, cost=sf_cost(4500, 360, Y4), e=0.0165 * Y4)                  # 소화슬러지 conveyor
    add(a, cost=sf_cost(4500, 360, Y5), e=0.0165 * Y5)                  # 케이크 conveyor
    P["탈수기"] = a

    # ---- 공정 9: 폐수처리 ----  (Anammox / A2O / Bardenpho)
    a = box()
    v = AE2 * DESIGN_MARGIN
    methanol = METHANOL_TON_PER_TON_N * N_ton * 365 * METHANOL_USD_PER_TON
    lime = LIME_TON_PER_TON_N * N_ton * 365 * (LIME_KRW_PER_KG * 1000 / EXCHANGE_RATE)
    alum = (ALUM_MG_PER_L * AE2 / 1000) * 365 * (ALUM_KRW_PER_KG / EXCHANGE_RATE)
    if OPT_WWT == 1:      # Anammox (총공사비 all-inclusive, 508원/kgN)
        add(a, cost=anammox_capex(v), all_inc=True,
            chem=(N_ton * 1000) * ANAMMOX_ETC_KRW_PER_KGN * 365 / EXCHANGE_RATE)
    elif OPT_WWT == 2:    # A2O
        add(a, cost=wwt_anaerobic(v), e=0.23 * AE2, chem=methanol + lime)   # 혐기조
        add(a, cost=wwt_anaerobic(v), e=0.0)                               # 무산소조
        add(a, cost=wwt_aerobic(v), e=0.412 * AE2)                         # 폭기조
        add(a, cost=wwt_clarifier(AE2), e=0.008 * AE2, chem=wwt_clarifier_om(AE2))  # 2차침전지
    else:                # Bardenpho
        add(a, cost=wwt_anaerobic(v), e=0.23 * AE2, chem=2 * methanol + lime)
        add(a, cost=wwt_anaerobic(v), e=0.0)
        add(a, cost=wwt_aerobic(v), e=0.412 * AE2)
        add(a, cost=wwt_anaerobic(v), e=0.0)                               # 2차 무산소조
        add(a, cost=wwt_aerobic(v), e=0.0)                                 # 2차 폭기조
        add(a, cost=wwt_clarifier(AE2), e=0.008 * AE2, chem=wwt_clarifier_om(AE2))
    add(a, cost=sf_cost(10000, 18, v), e=0.0)                             # 균등조
    add(a, cost=sf_cost(20000, 120, v), e=0.05 * AE2, chem=alum)          # 가압부상조(DAF)
    add(a, cost=sf_cost(7500, 24, v), e=0.1 * AE2)                        # UV
    P["폐수처리"] = a

    # ================================================================
    # CAPEX  (식 2)
    # ================================================================
    C_equipment = sum(p["equip"] for p in P.values())      # 장비 구매비
    C_anammox   = sum(p["anammox"] for p in P.values())    # Anammox(부대 이미 포함)
    std_cost = C_equipment / PURCHASED_EQUIP_RATIO         # 표준사업비(파생) = 장비비/0.229
    budae = {name: rate * std_cost for name, rate in STD_RATES.items()}   # 부대비용 항목
    CAPEX_total = C_equipment + sum(budae.values()) + C_anammox           # = std_cost + Anammox

    # 정부 보조금 = 회귀식 표준사업비 x 60% (고정)
    reg_std_cost = (2.7322 * total_ton + 79.56) * 1e8                     # 원
    subsidy = SUBSIDY_RATE * reg_std_cost / EXCHANGE_RATE                 # $
    CAPEX_net = max(0.0, CAPEX_total - subsidy)                          # 순 CAPEX (자본비 회수 대상)

    # 자본회수계수 (식: f_crf = i / (1-(1+i)^-L))
    f_crf = f_wacc / (1 - (1 + f_wacc) ** -L)

    # ================================================================
    # OPEX  (식 4·5·6)
    # ================================================================
    e24_tot = sum(p["e24"] for p in P.values())
    e8_tot  = sum(p["e8"] for p in P.values())
    energy_krw, total_kwh = elec_tariff(e24_tot / 24.0, e8_tot / 8.0)
    O_electricity = (ELEC_BASE_CHARGE + energy_krw
                     + ELEC_CLIMATE_RATE * total_kwh
                     + ELEC_FUEL_ADJ_RATE * total_kwh) / EXCHANGE_RATE    # 전기비
    O_chemical = sum(p["chem"] for p in P.values())                      # 약품비(+발전기 O&M)

    base_om = OM_CAPEX_FRACTION * CAPEX_total                             # O&M 적용 base (34.8%)
    O_salary      = OM_SALARY_RATE * base_om
    O_benefit     = OM_BENEFIT_RATE * O_salary
    O_maintenance = OM_MAINTENANCE_RATE * base_om
    O_lab         = OM_LAB_RATE * base_om
    O_insurance   = OM_INSURANCE_RATE * base_om
    OPEX_fixed = O_salary + O_benefit + O_maintenance + O_lab + O_insurance
    OPEX_var   = O_chemical + O_electricity
    OPEX_total = OPEX_fixed + OPEX_var

    annual_cost = CAPEX_net * f_crf + OPEX_total

    # ================================================================
    # LCOB (정제) / LCOE (발전)  (식 1·7)
    # ================================================================
    R = dict(total_ton=total_ton, AE2=AE2, Y4=Y4, Y5=Y5,
             C_equipment=C_equipment, C_anammox=C_anammox, std_cost=std_cost,
             budae=budae, CAPEX_total=CAPEX_total, reg_std_cost=reg_std_cost,
             subsidy=subsidy, CAPEX_net=CAPEX_net, f_crf=f_crf,
             O_electricity=O_electricity, O_chemical=O_chemical,
             OPEX_fixed=OPEX_fixed, OPEX_var=OPEX_var, OPEX_total=OPEX_total,
             total_kwh=total_kwh, P=P, mode=MODE)

    if MODE == "정제":
        Q_CH4 = BIOGAS * (1 - HEATING_FRACTION)                          # 판매가능 메탄
        denom = Q_CH4 * 365 * f_util
        LCOB = annual_cost / denom
        R.update(Q_CH4=Q_CH4,
                 LCOB_m3=LCOB, LCOB_kwh=LCOB / (CH4_HV_MJ_PER_M3 / MJ_PER_KWH),
                 LCOB_MJ=LCOB / CH4_HV_MJ_PER_M3)
    else:
        gen_kwh_yr = gen_kwh_day * 365 * f_util
        LCOE_A = annual_cost / gen_kwh_yr if gen_kwh_yr else float("nan")
        net_export = gen_kwh_yr - total_kwh                              # 잉여전기
        cost_b = CAPEX_net * f_crf + OPEX_fixed + O_chemical             # C_elec 제외
        LCOE_B = cost_b / net_export if net_export > 0 else float("nan")
        R.update(gen_kwh_day=gen_kwh_day, gen_kwh_yr=gen_kwh_yr,
                 net_export=net_export, LCOE_A=LCOE_A, LCOE_B=LCOE_B)
    return R

# ====================================================================
# 5. 결과 출력
# ====================================================================
def print_report(R):
    W = EXCHANGE_RATE
    print("=" * 60)
    print(f"CAPEX_total  [장비구매비/0.229 + Anammox]")
    print("=" * 60)
    print(f"  C_equipment (장비구매비)  ${R['C_equipment']:>14,.0f}")
    print(f"  표준사업비(파생 =/0.229)  ${R['std_cost']:>14,.0f}")
    print(f"  부대비용 합               ${sum(R['budae'].values()):>14,.0f}")
    if R['C_anammox']:
        print(f"  Anammox(부대포함·별도)    ${R['C_anammox']:>14,.0f}")
    print(f"  CAPEX_total               ${R['CAPEX_total']:>14,.0f}  (₩{R['CAPEX_total']*W:,.0f})")
    print(f"  - 보조금(회귀표준x60%)    ${R['subsidy']:>14,.0f}")
    print(f"  CAPEX_net (순)            ${R['CAPEX_net']:>14,.0f}")
    print(f"  f_crf (i={f_wacc}, L={L})     {R['f_crf']:>14.6f}")
    print("=" * 60)
    print("OPEX")
    print("=" * 60)
    print(f"  OPEX_fixed (고정 O&M)     ${R['OPEX_fixed']:>14,.0f}")
    print(f"  O_electricity (전기비)    ${R['O_electricity']:>14,.0f}")
    print(f"  O_chemical (약품비)       ${R['O_chemical']:>14,.0f}")
    print(f"  OPEX_total                ${R['OPEX_total']:>14,.0f}/yr")
    print("=" * 60)
    if R['mode'] == "정제":
        print(f"LCOB  (Q_CH4 = {R['Q_CH4']:,.0f} m3/day, 가온 {HEATING_FRACTION*100:.0f}% 제외)")
        print("=" * 60)
        print(f"  LCOB = ${R['LCOB_m3']:.4f}/m3  (₩{R['LCOB_m3']*W:,.0f}/m3)")
        print(f"       = ${R['LCOB_kwh']:.4f}/kWh (₩{R['LCOB_kwh']*W:,.1f}/kWh)")
        print(f"       = ${R['LCOB_MJ']:.4f}/MJ  (₩{R['LCOB_MJ']*W:,.1f}/MJ)")
    else:
        print(f"LCOE  (발전가능량 {R['gen_kwh_day']:,.0f} kWh/day, 연간 {R['gen_kwh_yr']:,.0f} kWh)")
        print("=" * 60)
        print(f"  [A] 자체소비=그리드구매, 전량판매 : ${R['LCOE_A']:.4f}/kWh (₩{R['LCOE_A']*W:,.1f})")
        if R['net_export'] > 0:
            print(f"  [B] 자가소비 상쇄, 잉여판매       : ${R['LCOE_B']:.4f}/kWh (₩{R['LCOE_B']*W:,.1f})")
        else:
            print(f"  [B] 발전량 < 자체소비 -> 잉여전기 없음")

# ====================================================================
# 6. 그래프  (CAPEX_total 위치 / 공정별 CAPEX·OPEX 비용 / breakdown)
# ====================================================================
def _loaded_capex(P):
    return {n: p["equip"] / PURCHASED_EQUIP_RATIO + p["anammox"] for n, p in P.items()}

def _process_opex(R):
    """공정별 OPEX = 전기 + 유지보수(부대포함 자본비 배분) + 약품."""
    P = R["P"]
    loaded = _loaded_capex(P)
    total_loaded = sum(loaded.values())
    total_kwh = R["total_kwh"]
    out = {}
    for n, p in P.items():
        energy_p, kwh_p = elec_tariff(p["e24"] / 24.0, p["e8"] / 8.0)
        base_p = ELEC_BASE_CHARGE * (kwh_p / total_kwh) if total_kwh else 0.0
        elec_p = (energy_p + ELEC_CLIMATE_RATE * kwh_p + ELEC_FUEL_ADJ_RATE * kwh_p + base_p) / EXCHANGE_RATE
        maint_p = R["OPEX_fixed"] * (loaded[n] / total_loaded) if total_loaded else 0.0
        out[n] = elec_p + maint_p + p["chem"]
    return out

PROCESS_EN = {"반입": "Intake", "전처리": "Pretreatment", "중간저장조": "Intermediate storage",
              "혐기성소화조": "Anaerobic digestion", "가스저장조": "Gas storage",
              "가스전처리": "Gas pretreatment", "발전": "Power generation",
              "탈수기": "Dewatering", "폐수처리": "Wastewater treatment"}

def make_plots(R):
    try:
        import numpy as np
        import matplotlib.pyplot as plt
    except ImportError:
        print("\n(그래프 생략: matplotlib 미설치)")
        return

    def stacked_bar(items, title, ylabel, fmt="${:,.0f}"):
        items = [(PROCESS_EN.get(n, n), v) for n, v in items if v > 0]
        items.sort(key=lambda t: t[1], reverse=True)
        names = [n for n, _ in items]; vals = [v for _, v in items]
        colors = plt.get_cmap("tab10").colors
        _, ax = plt.subplots(figsize=(4.5, 7))
        bottom = 0.0
        for i, (n, v) in enumerate(zip(names, vals)):
            ax.bar(title, v, bottom=bottom, color=colors[i % 10], width=0.6, edgecolor="white", label=n)
            if v / sum(vals) > 0.04:
                ax.text(0, bottom + v / 2, fmt.format(v), ha="center", va="center",
                        color="white", fontsize=8, fontweight="bold")
            bottom += v
        h, l = ax.get_legend_handles_labels()
        ax.legend(h[::-1], l[::-1], bbox_to_anchor=(1.02, 1), loc="upper left", fontsize=9)
        ax.set_ylabel(ylabel); ax.set_title(title); ax.set_xlim(-0.6, 0.6)
        plt.tight_layout(); plt.show()

    # (1) CAPEX_total 위치 (회귀식 표준사업비 직선 +-10%)
    cap_eok = R["CAPEX_total"] * EXCHANGE_RATE / 1e8
    ton = R["total_ton"]
    x = np.linspace(0, max(200.0, ton * 1.5), 300)
    y = 2.7322 * x + 79.56
    plt.figure(figsize=(8, 6))
    plt.plot(x, y, color="steelblue", lw=2, label="Standard project cost")
    plt.plot(x, y * 1.1, "--", color="steelblue", lw=1, alpha=0.7, label="+10%")
    plt.plot(x, y * 0.9, "--", color="steelblue", lw=1, alpha=0.7, label="-10%")
    plt.fill_between(x, y * 0.9, y * 1.1, color="steelblue", alpha=0.08)
    plt.scatter([ton], [cap_eok], color="crimson", s=45, zorder=5, label=f"CAPEX_total = {cap_eok:.1f}")
    plt.xlabel("Organic waste input (ton/day)"); plt.ylabel("Cost (10^8 KRW)")
    plt.title("CAPEX_total vs standard project cost"); plt.legend(); plt.grid(alpha=0.3)
    plt.tight_layout(); plt.show()

    # (2) 공정별 CAPEX 비용($)  (3) 공정별 OPEX 비용($)
    stacked_bar(list(_loaded_capex(R["P"]).items()), "CAPEX", "CAPEX ($)")
    stacked_bar(list(_process_opex(R).items()), "OPEX", "OPEX ($/year)")

    # (4) LCOB / LCOE breakdown
    if R["mode"] == "정제":
        denom = R["Q_CH4"] * 365 * f_util
        unit, title = "LCOB ($/m3)", "LCOB breakdown"
    else:
        denom = R["gen_kwh_yr"]
        unit, title = "LCOE ($/kWh)", "LCOE breakdown"
    if denom > 0:
        comp = [("CAPEX (annualized)", R["CAPEX_net"] * R["f_crf"] / denom),
                ("O&M (fixed)", R["OPEX_fixed"] / denom),
                ("Electricity", R["O_electricity"] / denom),
                ("Chemicals/Gen O&M", R["O_chemical"] / denom)]
        stacked_bar(comp, title.split()[0], unit, fmt="${:.4f}")


# ====================================================================
if __name__ == "__main__":
    R = calculate()
    print_report(R)
    make_plots(R)
