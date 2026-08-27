"""
가전팀 대시보드 데이터 생성 스크립트 (전체)
=================================================
매주 월요일 실행. Google Drive에서 받은 pgm.csv / competitor.xlsx / weight_targets.csv를
같은 폴더에 두고 실행하면 data/ 5개 JSON을 전부 재생성한다.

입력 파일 (스크립트와 같은 폴더에 위치):
  - pgm.csv              : PGM 실적 (fileId 1figvjnk_pR6sq3cc2vC-rs_29K92GnPTIddEdzQCUyY, exportMimeType text/csv)
  - competitor.xlsx       : 경쟁사 편성 (fileId 1qMHEjeHpFbhxDrdQqmBirOEKI9l8S0EJUUe4Rz_lZJ8,
                             exportMimeType application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,
                             시트: KT알파 / SSG)
  - weight_targets.csv    : 가중분 목표 (fileId 1dTDL0L-ZilM7mBYq8QZ85iyE3gvtBygrsEByLs8dVOU, exportMimeType text/csv)
  - current_dashboard-data.json : 현재 배포된 dashboard-data.json (kpiTarget 등 고정값 유지용, GitHub Contents API로 미리 받아둘 것)

출력: ./output/ 아래에 6개 JSON 생성
  dashboard-data.json, newproduct-data.json, search-data.json,
  competitor-data.json, product-heatmap-data.json, corevendor-data.json

사용법 예:
  python3 build_all.py --today 2026-07-19

주의:
  - 취급고는 항상 컬럼17(VAT포함) 사용. 컬럼18(V-) 사용 금지.
  - 신상품 비중 공식 지표는 "단순 태그 기준"(신상품 태그 붙은 방송 그대로 집계). "생애귀속"(같은 상품명이
    한번이라도 신상품 태그를 받으면 그 상품의 전체 생애 실적을 귀속)은 MD 기여도 분석(newproduct-data.json)에만 사용.
  - 대형가전은 KT알파/SSG 카테고리 매핑이 없어 competitor-data.json / product-heatmap-data.json 에서는 제외.
  - 8월 MD 카테고리 고정 담당제 폐지 이후에는 mdData의 "cat" 라벨(고정 카테고리 설명 문구)이 더 이상 유효하지
    않을 수 있으니 매번 확인. cats/top5 breakdown 자체는 실제 데이터 기반이라 자동으로 맞게 나온다.
"""
import argparse, re, json
import pandas as pd

CATS = ['생활가전', '주방가전', '주방용품', '대형가전']
CATS3 = ['생활가전', '주방가전', '주방용품']  # 대형가전 제외 (경쟁사 비교용)
CAT_EMOJI = {'생활가전': '🏠 생활가전', '주방가전': '🍳 주방가전', '주방용품': '🍳 주방용품', '대형가전': '📦 대형가전'}
DOW_LABELS = ['월', '화', '수', '목', '금', '토', '일']
MDMAP = {'ko': '권오석', 'ma': '마영호', 'kim': '김응도', 'lim': '임동진', 'baek': '백혜정'}
COLORS = {'ko': '#4d8fff', 'ma': '#ff4d6a', 'kim': '#22c87a', 'lim': '#f5a623', 'baek': '#9b7fff'}
# 8월 개편 이후 이 라벨은 더 이상 고정 담당 의미가 없을 수 있음 — 매번 형님께 최신 담당 확인 후 갱신할 것
CATLABEL = {'ko': '주방가전 · 주방용품', 'ma': '주방용품 · 주방가전', 'kim': '생활가전', 'lim': '생활가전', 'baek': '대형가전'}

# 핵심 협력사 트래커 (2026-07-28 신설). 벤더명↔판매 브랜드명이 다른 경우가 있으니(예: 벤더 에코센스 =
# 판매 브랜드 히트락) 별칭 목록으로 관리한다. 새 별칭이 발견되면 aliases 리스트에 한 줄만 추가할 것.
# compareTarget은 벤더별로 실제 경쟁 중인 채널이 다르므로 개별 지정 (마이하우스/코지마=KT알파, 히트락=SSG).
VENDOR_DEFS = [
    {"id": "myhouse", "name": "마이하우스", "subLabel": "마이어·라포레·안방·AM 등 10개 서브브랜드",
     "aliases": ["마이어", "라포레", "파베르", "안방", "델리원", "프리파라", "렌지프로", "부가티", "모나쿡", "AM"],
     "compareTarget": "KT알파"},
    {"id": "kojima", "name": "코지마", "subLabel": "안마의자·목어깨마사지기·문체어 등",
     "aliases": ["코지마"], "compareTarget": "KT알파"},
    {"id": "hitrock", "name": "히트락", "subLabel": "벤더사: 에코센스 (올트라이탄 밀폐용기)",
     "aliases": ["히트락", "에코센스", "조이락"], "compareTarget": "SSG"},
]


def num(s):
    return pd.to_numeric(s.astype(str).str.replace(',', '', regex=False), errors='coerce').fillna(0)


def brand_of(name):
    mt = re.match(r'^\[([^\]]+)\]', str(name))
    return mt.group(1) if mt else str(name)


def pm(a, b):
    return round(a / b / 1e4, 1) if b else 0.0


def load_pgm(path='pgm.csv'):
    df = pd.read_csv(path, encoding='utf-8-sig', low_memory=False, dtype=str)
    c = df.columns.tolist()
    cols = dict(YEAR=c[0], MON=c[1], WK=c[2], DATE=c[3], TIME=c[5],
                CAT=c[12], BRAND=c[11], WMIN=c[14], SALES=c[17], REV=c[21],
                MARGIN=c[39], MD=c[70], SHIN=c[71], TEAM=c[72])
    df = df[(df[cols['TEAM']] == '가전팀') & (df[cols['YEAR']] == '2026')].copy()
    for col in [cols['WMIN'], cols['SALES'], cols['REV'], cols['MARGIN']]:
        df[col] = num(df[col])
    df['MONI'] = pd.to_numeric(df[cols['MON']], errors='coerce').astype(int)
    df['dt'] = pd.to_datetime(df[cols['DATE']], errors='coerce')
    df['DOW'] = df['dt'].dt.dayofweek
    df['HOUR'] = pd.to_numeric(df[cols['TIME']].astype(str).str.slice(0, 2), errors='coerce').fillna(0).astype(int)
    df['brand'] = df[cols['BRAND']].apply(brand_of)
    return df, cols


def build_monthData(f, cols):
    monthData = {}
    for m in sorted(f['MONI'].unique()):
        sub = f[f['MONI'] == m]
        w, s, r, mg = sub[cols['WMIN']].sum(), sub[cols['SALES']].sum(), sub[cols['REV']].sum(), sub[cols['MARGIN']].sum()
        total = {"weight": f"{w:,.0f}분", "sales": f"{s/1e8:.1f}억", "perMin": f"{pm(s,w):.1f}만",
                 "revenue": f"{r/1e8:.1f}억", "margin": f"{mg/1e8:.1f}억", "perMargin": f"{pm(mg,w):.1f}만"}
        cats = []
        for cat in CATS:
            csub = sub[sub[cols['CAT']] == cat]
            cw, cs, cm = csub[cols['WMIN']].sum(), csub[cols['SALES']].sum(), csub[cols['MARGIN']].sum()
            cats.append({"name": CAT_EMOJI[cat], "weight": f"{cw:,.0f}분", "sales": f"{cs/1e8:.1f}억",
                         "perMin": f"{pm(cs,cw):.1f}만", "margin": f"{cm/1e8:.1f}억", "perMargin": f"{pm(cm,cw):.1f}만"})
        monthData[str(m)] = {"label": f"{m}월", "total": total, "cats": cats}
    return monthData


def build_weeklyData(f_all_team, cols):
    """f_all_team: 가전팀 전체(연도 필터 없이, 주차 라벨 연속성을 위해 전년도 말 포함)"""
    grp = f_all_team.groupby(cols['WK'])
    order = grp['dt'].min().sort_values()
    byWeek, weeks = {}, []
    for wk in order.index:
        sub = grp.get_group(wk)
        dr = f"{sub['dt'].min().strftime('%m.%d')}~{sub['dt'].max().strftime('%m.%d')}"
        w, s, mg = sub[cols['WMIN']].sum(), sub[cols['SALES']].sum(), sub[cols['MARGIN']].sum()
        total = {"weight": round(w,1), "sales": round(s/1e8,3), "margin": round(mg/1e8,3),
                 "perMin": pm(s,w), "perMargin": pm(mg,w)}
        cats = {}
        for cat in CATS:
            csub = sub[sub[cols['CAT']] == cat]
            cw, cs, cm = csub[cols['WMIN']].sum(), csub[cols['SALES']].sum(), csub[cols['MARGIN']].sum()
            cats[cat] = {"weight": round(cw,1), "sales": round(cs/1e8,3), "margin": round(cm/1e8,3),
                        "perMin": pm(cs,cw), "perMargin": pm(cm,cw)}
        byWeek[wk] = {"label": wk, "dateRange": dr, "total": total, "cats": cats}
        weeks.append(wk)
    return {"weeks": weeks, "byWeek": byWeek}


def build_mdData(f, cols):
    mdData = {}
    for key, name in MDMAP.items():
        sub = f[f[cols['MD']] == name]
        w, s, mg, r = sub[cols['WMIN']].sum(), sub[cols['SALES']].sum(), sub[cols['MARGIN']].sum(), sub[cols['REV']].sum()
        total = {"cnt": int(len(sub)), "weight": f"{w:,.0f}분", "sales": f"{s/1e8:.1f}억",
                 "perMin": f"{pm(s,w):.1f}만", "margin": f"{mg/1e8:.2f}억", "perMargin": f"{pm(mg,w):.1f}만",
                 "revenue": f"{r/1e8:.1f}억"}
        monthly_sales, monthly_margin = [], []
        for m in range(1, 8):
            msub = sub[sub['MONI'] == m]
            monthly_sales.append(round(msub[cols['SALES']].sum()/1e8, 1))
            monthly_margin.append(round(msub[cols['MARGIN']].sum()/1e8, 1))
        cat_present = list(sub[cols['CAT']].dropna().unique())
        cat_sales = {ct: sub[sub[cols['CAT']] == ct][cols['SALES']].sum() for ct in cat_present}
        cats = []
        for ct in sorted(cat_sales, key=lambda x: -cat_sales[x]):
            csub = sub[sub[cols['CAT']] == ct]
            cw, cs, cm = csub[cols['WMIN']].sum(), csub[cols['SALES']].sum(), csub[cols['MARGIN']].sum()
            pct = round(cs/s*100) if s else 0
            cats.append({"name": CAT_EMOJI.get(ct, ct), "sales": f"{cs/1e8:.2f}억", "perMin": f"{pm(cs,cw):.1f}만",
                        "margin": f"{cm/1e8:.2f}억", "pct": f"{pct}%"})
        prod = sub.groupby(cols['BRAND']).agg(sales=(cols['SALES'],'sum'), weight=(cols['WMIN'],'sum'), margin=(cols['MARGIN'],'sum')).reset_index()
        prod = prod.sort_values('sales', ascending=False).head(5)
        top5 = []
        for _, row in prod.iterrows():
            marginRate = (row['margin']/row['sales']*100) if row['sales'] else 0
            top5.append({"name": row[cols['BRAND']], "sales": f"{row['sales']/1e8:.2f}억",
                        "perMin": f"{pm(row['sales'],row['weight']):.1f}만", "marginRate": f"{marginRate:.1f}%"})
        mdData[key] = {"name": name, "color": COLORS[key], "cat": CATLABEL[key], "total": total,
                       "monthly": {"sales": monthly_sales, "margin": monthly_margin}, "cats": cats, "top5": top5}
    return mdData


def build_shin(f, cols):
    def grp_metrics(sub):
        w, s, mg = sub[cols['WMIN']].sum(), sub[cols['SALES']].sum(), sub[cols['MARGIN']].sum()
        return {"cnt": int(len(sub)), "w": round(w,2), "s": round(s/1e8,2), "ma": round(mg/1e8,2),
                "perMin": pm(s,w), "perMar": pm(mg,w)}
    shin_rows = f[f[cols['SHIN']] == '신상품']
    redo_rows = f[f[cols['SHIN']].isin(['재녹화','재편집'])]
    existing_rows = f[f[cols['SHIN']] == '기존']
    shinSummary = {"shin": grp_metrics(shin_rows), "redo": grp_metrics(redo_rows), "existing": grp_metrics(existing_rows)}
    shinMonthly = {"shin": [], "redo": []}
    for m in range(1, 8):
        shinMonthly["shin"].append(round(shin_rows[shin_rows['MONI']==m][cols['SALES']].sum()/1e8, 2))
        shinMonthly["redo"].append(round(redo_rows[redo_rows['MONI']==m][cols['SALES']].sum()/1e8, 2))
    shin_products = set(shin_rows[cols['BRAND']].dropna().unique())
    shinTrendData = {}
    for key, name in MDMAP.items():
        sub = f[f[cols['MD']] == name]
        rows = []
        for m in range(1, 8):
            msub = sub[sub['MONI'] == m]
            total_s = msub[cols['SALES']].sum()
            shin_s = msub[msub[cols['BRAND']].isin(shin_products)][cols['SALES']].sum()
            pct = round(shin_s/total_s*100, 1) if total_s else 0.0
            rows.append({"month": str(m), "total_s": round(total_s/1e8,2), "shin_s": round(shin_s/1e8,2), "shin_pct": pct})
        shinTrendData[name] = rows
    return shinSummary, shinMonthly, shinTrendData, shin_products


def build_vendor_concentration(f, cols):
    """협력사 편성 집중도. 편성코드명 대괄호 브랜드를 협력사 단위로 통합(VENDOR_DEFS aliases 사용).
    가중분 기준 점유율 / 누적 집중도(CR3·CR5·CR10) / HHI / 월별 상위 3사 집중도 추이."""
    vmap = {}
    for v in VENDOR_DEFS:
        for a in v['aliases']:
            vmap[a] = v['name']
    g = f.copy()
    g['vendor'] = g['brand'].map(lambda b: vmap.get(b, b))
    tot_w = g[cols['WMIN']].sum()
    tot_s = g[cols['SALES']].sum()
    agg = g.groupby('vendor').agg(w=(cols['WMIN'], 'sum'), s=(cols['SALES'], 'sum'),
                                  mg=(cols['MARGIN'], 'sum'), cnt=(cols['WMIN'], 'size')
                                  ).sort_values('w', ascending=False)
    sh = (agg['w'] / tot_w).values
    cum = (sh.cumsum() * 100)

    def cr(k):
        return round(float(cum[k - 1]), 1) if len(cum) >= k else 100.0

    top = []
    for name, r in agg.head(10).iterrows():
        top.append({"name": name, "w": int(round(r.w)), "wsh": round(r.w / tot_w * 100, 1),
                    "s": round(r.s / 1e8, 1), "ssh": round(r.s / tot_s * 100, 1),
                    "cnt": int(r.cnt), "perMin": pm(r.s, r.w), "perMar": pm(r.mg, r.w)})
    top3names = list(agg.index[:3])
    trend = []
    for m in sorted(g['MONI'].unique()):
        sub = g[g['MONI'] == m]
        mw = sub[cols['WMIN']].sum()
        t3 = sub[sub['vendor'].isin(top3names)][cols['WMIN']].sum()
        trend.append({"month": f"{int(m)}월", "pct": round(t3 / mw * 100, 1) if mw else 0.0,
                      "vcnt": int(sub['vendor'].nunique())})
    return {"totalVendors": int(agg.shape[0]), "totalW": int(round(tot_w)),
            "top3": cr(3), "top5": cr(5), "top10": cr(10),
            "top3Sales": round(float(agg['s'].head(3).sum() / tot_s * 100), 1),
            "hhi": int(round(float((sh ** 2).sum() * 10000))),
            "top3Names": top3names, "top": top, "trend": trend}


def build_shin_settle(f, cols, target_q3=55, target_year=60, year_goal=70, min_w=5.0, mature_months=2):
    """신상품 정착률 = 런칭 후 '유효방송' 3회 이상 비중.

    측정 규칙 (2026-08-27 확정):
      1) 상품 식별 = 편성코드명(cols['BRAND']), build_shin과 동일
      2) 런칭일   = 해당 상품의 첫 '신상품' 태그 편성일. 태그 이전 편성은 카운트하지 않는다.
      3) 유효방송 = 가중분 >= min_w(5분). 가중분 0~0.6분짜리 등록성 행과 협력사 케어용
         저가중분 새벽 편성이 '방송 1회'로 잡혀 정착률이 부풀려지는 것을 막는다.
         (새벽대라도 가중분 5분 이상 실방송은 재편성이 일어난 것이므로 카운트한다)
      4) 관찰기간 = 런칭 후 mature_months(2개월) 경과 상품만 본지표에 집계.
         최근 런칭은 3회를 채울 시간이 없었을 뿐이므로 'pending'으로 분리한다.
    """
    valid = f[f[cols['WMIN']] >= min_w]
    shin_rows = f[f[cols['SHIN']] == '신상품']
    launch = shin_rows.groupby(cols['BRAND'])['dt'].min()
    maxdt = f['dt'].max()
    cutoff = maxdt - pd.DateOffset(months=mature_months)

    v = valid[valid[cols['BRAND']].isin(launch.index)].copy()
    v['_ld'] = v[cols['BRAND']].map(launch)
    lifecnt = v[v['dt'] >= v['_ld']].groupby(cols['BRAND']).size()

    sp = shin_rows.drop_duplicates(subset=[cols['BRAND']])[[cols['BRAND'], cols['CAT']]].copy()
    sp.columns = ['name', 'cat']
    sp['ld'] = sp['name'].map(launch)
    sp['lm'] = sp['ld'].dt.month
    sp['cnt'] = sp['name'].map(lifecnt).fillna(0).astype(int)
    sp['mature'] = sp['ld'] <= cutoff

    allT = int(len(sp))
    mt = sp[sp['mature']]
    n = int(len(mt))
    base = {"targetQ3": target_q3, "targetYear": target_year, "yearGoal": year_goal,
            "minW": min_w, "matureMonths": mature_months,
            "rule": f"유효방송 = 가중분 {min_w:.0f}분 이상 · 신상품 태그일 이후 · 런칭 {mature_months}개월 경과 상품만 집계",
            "allTotal": allT, "pending": allT - n}
    if n == 0:
        base.update({"total": 0, "ge3": 0, "pct": 0.0, "allGe3": 0, "allPct": 0.0,
                     "avgCnt": 0.0, "buckets": [], "cohort": [], "byCat": []})
        return base
    ge3 = int((mt['cnt'] >= 3).sum())
    allGe3 = int((sp['cnt'] >= 3).sum())
    buckets = [{"label": "1회", "n": int((mt['cnt'] <= 1).sum())},
               {"label": "2회", "n": int((mt['cnt'] == 2).sum())},
               {"label": "3~4회", "n": int(((mt['cnt'] >= 3) & (mt['cnt'] <= 4)).sum())},
               {"label": "5회 이상", "n": int((mt['cnt'] >= 5).sum())}]
    for b in buckets:
        b['pct'] = round(b['n'] / n * 100, 1)
    cohort = []
    for m in sorted(sp['lm'].dropna().unique()):
        sub = sp[sp['lm'] == m]
        sm = sub[sub['mature']]
        cohort.append({"month": f"{int(m)}월", "n": int(len(sm)),
                       "ge3": int((sm['cnt'] >= 3).sum()) if len(sm) else 0,
                       "pct": round(float((sm['cnt'] >= 3).mean() * 100), 1) if len(sm) else 0.0,
                       "pending": int(len(sub) - len(sm))})
    byCat = []
    for c in CATS:
        sub = mt[mt['cat'] == c]
        if len(sub) == 0:
            continue
        byCat.append({"cat": c, "n": int(len(sub)), "ge3": int((sub['cnt'] >= 3).sum()),
                      "pct": round(float((sub['cnt'] >= 3).mean() * 100), 1)})
    base.update({"total": n, "ge3": ge3, "pct": round(ge3 / n * 100, 1),
                 "allGe3": allGe3, "allPct": round(allGe3 / allT * 100, 1),
                 "avgCnt": round(float(mt['cnt'].mean()), 1),
                 "buckets": buckets, "cohort": cohort, "byCat": byCat})
    return base


def build_newproduct(f, cols, shin_products):
    MD_DATA = []
    for key, name in MDMAP.items():
        sub = f[f[cols['MD']] == name]
        total_s, total_cnt = sub[cols['SALES']].sum(), len(sub)
        shin_sub = sub[sub[cols['BRAND']].isin(shin_products)]
        shin_s, shin_cnt, shin_w = shin_sub[cols['SALES']].sum(), len(shin_sub), shin_sub[cols['WMIN']].sum()
        shin_prod_cnt = shin_sub[cols['BRAND']].nunique()
        shin_pct = round(shin_s/total_s*100, 1) if total_s else 0.0
        gy_s, gy_cnt = total_s - shin_s, total_cnt - shin_cnt
        MD_DATA.append({"md": name, "total_s": round(total_s/1e8,2), "total_cnt": int(total_cnt),
                        "shin_s": round(shin_s/1e8,2), "shin_cnt": int(shin_cnt), "shin_prod_cnt": int(shin_prod_cnt),
                        "shin_pct": shin_pct, "shin_pm": pm(shin_s, shin_w),
                        "gy_s": round(gy_s/1e8,2), "gy_cnt": int(gy_cnt), "gy_pct": round(100-shin_pct,1)})
    SHIN_DATA = []
    for prod in shin_products:
        psub = f[f[cols['BRAND']] == prod]
        w, s, mg = psub[cols['WMIN']].sum(), psub[cols['SALES']].sum(), psub[cols['MARGIN']].sum()
        cat = psub[cols['CAT']].mode().iloc[0] if not psub[cols['CAT']].mode().empty else ''
        tagged = psub[psub[cols['SHIN']] == '신상품']
        md = tagged[cols['MD']].iloc[0] if len(tagged) else psub[cols['MD']].mode().iloc[0]
        months = sorted(psub['MONI'].unique().tolist())
        mr = round(mg/s*100, 1) if s else 0.0
        SHIN_DATA.append({"name": prod, "cat": cat, "md": md, "cnt": int(len(psub)), "w": round(w,1),
                          "s": round(s/1e8,3), "pm": pm(s,w), "m": round(mg/1e8,3), "pmm": pm(mg,w),
                          "mr": mr, "months": [str(x) for x in months]})
    SHIN_DATA.sort(key=lambda x: (-x['s'], x['name']))
    return MD_DATA, SHIN_DATA


def build_search(f, cols):
    out = []
    for prod, psub in f.groupby(cols['BRAND']):
        psub = psub.sort_values('dt')
        w, s, mg = psub[cols['WMIN']].sum(), psub[cols['SALES']].sum(), psub[cols['MARGIN']].sum()
        cat = psub[cols['CAT']].mode().iloc[0] if not psub[cols['CAT']].mode().empty else ''
        md = psub[cols['MD']].mode().iloc[0] if not psub[cols['MD']].mode().empty else ''
        shin_tag = psub.iloc[-1][cols['SHIN']]
        months = sorted(psub['MONI'].unique().tolist())
        mr = round(mg/s*100, 1) if s else 0.0
        bc = []
        for _, row in psub.iterrows():
            rw, rs, rm = row[cols['WMIN']], row[cols['SALES']], row[cols['MARGIN']]
            bc.append({"d": str(row['dt'].date()).replace('-','/'), "t": row[cols['TIME']], "m": str(int(row['MONI'])),
                      "w": round(rw,2), "s": round(rs/1e8,3), "pm": pm(rs,rw), "im": round(rm/1e8,3), "pmm": pm(rm,rw)})
        out.append({"name": prod, "brand": brand_of(prod), "cat": cat, "md": md, "shin": shin_tag,
                    "cnt": int(len(psub)), "w": round(w,1), "s": round(s/1e8,2), "pm": pm(s,w), "m": round(mg/1e8,2),
                    "pmm": pm(mg,w), "mr": mr, "months": [str(x) for x in months], "bc": bc})
    out.sort(key=lambda x: -x['s'])
    return out


def build_competitor(f_all3, cols, kt, ssg):
    def cat_weight(dfc, catcol, wcol='가중분'):
        return {ct: round(float(dfc[dfc[catcol]==ct][wcol].sum()),2) for ct in CATS3}
    def pct_of(catdict, total):
        return {k: (v/total*100 if total else 0.0) for k, v in catdict.items()}
    kt_cat, ssg_cat = cat_weight(kt, 'MD분류'), cat_weight(ssg, 'MD분류')
    sk_cat = {ct: round(float(f_all3[f_all3[cols['CAT']]==ct][cols['WMIN']].sum()),2) for ct in CATS3}
    kt_total, ssg_total, sk_total = round(sum(kt_cat.values()),2), round(sum(ssg_cat.values()),2), round(sum(sk_cat.values()),2)
    category = {"categories": CATS3, "kt": kt_cat, "kt_total": kt_total, "kt_pct": pct_of(kt_cat, kt_total),
               "ssg": ssg_cat, "ssg_total": ssg_total, "ssg_pct": pct_of(ssg_cat, ssg_total),
               "sk": sk_cat, "sk_total": sk_total, "sk_pct": pct_of(sk_cat, sk_total)}

    def monthly_block(dfc, catcol, wcol, monthcol):
        tot_w = {ct: round(float(dfc[dfc[catcol]==ct][wcol].sum()),2) for ct in CATS3}
        tot = round(sum(tot_w.values()),2)
        out = {"0": {"total": tot, "w": tot_w, "pct": pct_of(tot_w, tot)}}
        for m in range(1, 8):
            msub = dfc[dfc[monthcol]==m]
            mw = {ct: round(float(msub[msub[catcol]==ct][wcol].sum()),2) for ct in CATS3}
            mt = round(sum(mw.values()),2)
            out[str(m)] = {"total": mt, "w": mw, "pct": pct_of(mw, mt)}
        return out
    categoryMonthly = {"SK스토아": monthly_block(f_all3, cols['CAT'], cols['WMIN'], 'MONI'),
                       "KT알파": monthly_block(kt, 'MD분류', '가중분', 'MON'),
                       "SSG": monthly_block(ssg, 'MD분류', '가중분', 'MON')}

    def top_brands(dfc, catcol, wcol, brandcol='brand'):
        sub = dfc[dfc[catcol].isin(CATS3)]
        g = sub.groupby(brandcol).agg(w=(wcol,'sum'), cnt=(wcol,'count')).reset_index().sort_values('w', ascending=False).head(12)
        return [{"brand": r[brandcol], "w": round(float(r['w']),1), "cnt": int(r['cnt'])} for _, r in g.iterrows()]
    brands = {"SK스토아": top_brands(f_all3, cols['CAT'], cols['WMIN'], 'brand'),
             "KT알파": top_brands(kt, 'MD분류', '가중분', 'brand'), "SSG": top_brands(ssg, 'MD분류', '가중분', 'brand')}

    def channel_mix(dfc):
        sub = dfc[dfc['MD분류'].isin(CATS3)]
        return {k: round(float(v),1) for k, v in sub.groupby('채널구분')['가중분'].sum().items()}
    channelMix = {"KT알파": channel_mix(kt), "SSG": channel_mix(ssg)}

    KEYWORDS = ['신상품','신상','NEW','런칭','론칭','첫방송']
    def kw_new_pct(dfc):
        sub = dfc[dfc['MD분류'].isin(CATS3)]
        pat = '|'.join(KEYWORDS)
        mask = sub['상품명'].astype(str).str.contains(pat, case=False, na=False) | sub['PGM명'].astype(str).str.contains(pat, case=False, na=False)
        tot = sub['가중분'].sum()
        return round(float(sub[mask]['가중분'].sum()/tot*100),2) if tot else 0.0
    sk_new_pct = round(float(f_all3[f_all3[cols['SHIN']]=='신상품'][cols['WMIN']].sum() / f_all3[cols['WMIN']].sum() * 100),2)
    newProduct = {"kt_new_pct": kw_new_pct(kt), "ssg_new_pct": kw_new_pct(ssg), "sk_new_pct": sk_new_pct,
                 "note": "SK스토아는 PGM \"신상품구분\" 태그 기준(단일 방송분), KT알파/SSG는 상품명 키워드 추정치로 산출 방식이 다름"}

    def build_grid(dfc, catcol, wcol, dowcol, hourcol):
        def grid_for(sub):
            g = [[0.0]*24 for _ in range(7)]
            for (d,h), v in sub.groupby([dowcol,hourcol])[wcol].sum().items():
                if 0<=d<7 and 0<=h<24: g[int(d)][int(h)] = round(float(v),2)
            return g
        out = {"전체": grid_for(dfc)}
        for ct in CATS3: out[ct] = grid_for(dfc[dfc[catcol]==ct])
        return out
    hourDowGrid = {"SK스토아": build_grid(f_all3, cols['CAT'], cols['WMIN'], 'DOW', 'HOUR'),
                  "KT알파": build_grid(kt, 'MD분류', '가중분', 'DOW', 'HOUR'),
                  "SSG": build_grid(ssg, 'MD분류', '가중분', 'DOW', 'HOUR')}

    productList = []
    sk_g = f_all3.groupby(['MONI', cols['WK'], cols['CAT'], cols['BRAND'], 'brand', cols['MD']]).agg(
        w=(cols['WMIN'],'sum'), cnt=(cols['WMIN'],'count')).reset_index()
    for _, r in sk_g.iterrows():
        productList.append({"co":"SK스토아","m":int(r['MONI']),"wk":r[cols['WK']],"cat":r[cols['CAT']],
                            "name":r[cols['BRAND']],"brand":r['brand'],"md":r[cols['MD']],
                            "w":round(float(r['w']),1),"cnt":int(r['cnt'])})
    def kt_ssg_products(dfc, coname):
        sub = dfc[dfc['MD분류'].isin(CATS3)].copy()
        namecol = '정제상품명 (아이템명)' if '정제상품명 (아이템명)' in sub.columns else '정제상품명(아이템명)'
        g = sub.groupby(['MON','YYYYWW','MD분류',namecol,'brand']).agg(w=('가중분','sum'), cnt=('가중분','count')).reset_index()
        return [{"co":coname,"m":int(r['MON']),"wk":r['YYYYWW'],"cat":r['MD분류'],"name":r[namecol],
                "brand":r['brand'],"w":round(float(r['w']),1),"cnt":int(r['cnt'])} for _, r in g.iterrows()]
    productList += kt_ssg_products(kt, 'KT알파') + kt_ssg_products(ssg, 'SSG')

    meta = {"period": None, "kt_rows": int(len(kt[kt['MD분류'].isin(CATS3)])), "ssg_rows": int(len(ssg[ssg['MD분류'].isin(CATS3)])),
           "sk_rows": int(len(f_all3)), "updated": None, "companies": ["SK스토아","KT알파","SSG"]}
    return {"category": category, "categoryMonthly": categoryMonthly, "brands": brands, "channelMix": channelMix,
           "newProduct": newProduct, "hourDowGrid": hourDowGrid, "hourDowDays": DOW_LABELS,
           "productList": productList, "meta": meta}


def build_corevendor(team_all, cols, kt_raw, ssg_raw, today):
    """핵심 협력사(마이하우스/코지마/히트락) 자사 vs 지정 경쟁사 월별 비교.
    벤더명↔판매 브랜드명 불일치 케이스(예: 에코센스=히트락)를 놓치지 않도록
    브랜드/정제상품명/상품명 3개 컬럼에서 alias를 OR 검색한다."""
    def contains_any(series, aliases):
        pat = '|'.join(r'(?<![A-Za-z0-9가-힣])' + re.escape(a) + r'(?![A-Za-z0-9가-힣])' for a in aliases)
        return series.astype(str).str.contains(pat, na=False, regex=True)

    def self_monthly(aliases):
        mask = contains_any(team_all[cols['BRAND']], aliases)
        sub = team_all[mask]
        out = {}
        for m, g in sub.groupby('MONI'):
            out[str(int(m))] = {"weight": round(float(g[cols['WMIN']].sum()), 1), "count": int(len(g))}
        return out

    def comp_monthly(dfc, aliases):
        namecol = '정제상품명 (아이템명)' if '정제상품명 (아이템명)' in dfc.columns else '정제상품명(아이템명)'
        cols_to_check = [c for c in ['brand', namecol, '상품명'] if c in dfc.columns]
        mask = False
        for c in cols_to_check:
            m = contains_any(dfc[c], aliases)
            mask = m if mask is False else (mask | m)
        sub = dfc[mask]
        out = {}
        for m, g in sub.groupby('MON'):
            out[str(int(m))] = {"weight": round(float(g['가중분'].sum()), 1), "count": int(len(g))}
        return out

    vendors_out = []
    for vd in VENDOR_DEFS:
        self_m = self_monthly(vd['aliases'])
        comp_df = ssg_raw if vd['compareTarget'] == 'SSG' else kt_raw
        comp_m = comp_monthly(comp_df, vd['aliases'])
        months = sorted(set(self_m) | set(comp_m), key=lambda x: int(x))
        self_series = [self_m.get(m, {"weight": 0.0})['weight'] for m in months]
        comp_series = [comp_m.get(m, {"weight": 0.0})['weight'] for m in months]
        last_self = self_series[-1] if self_series else 0.0
        last_comp = comp_series[-1] if comp_series else 0.0
        penetration = round(last_comp / last_self * 100, 1) if last_self else (0.0 if not last_comp else None)
        vendors_out.append({
            "id": vd['id'], "name": vd['name'], "subLabel": vd['subLabel'],
            "aliases": vd['aliases'], "compareTarget": vd['compareTarget'],
            "months": [f"2026-{int(m):02d}" for m in months],
            "monthLabels": [f"{int(m)}월" for m in months],
            "self": self_series, "comp": comp_series,
            "penetration": penetration,
            # risk/trendLabel/flags는 수치 변화 폭이 커서 매주 지미가 직접 판단 후 수기 갱신 권장
            # (기본값만 채워두고, 이상 신호 있으면 대시보드 페이지에서 직접 수정)
            "risk": "high" if (penetration or 0) >= 100 else ("mid" if (penetration or 0) >= 50 else "low"),
            "trendLabel": "", "flags": []
        })

    return {
        "meta": {"period": None, "updated": None,
                 "note": "핵심 협력사(마이하우스/코지마/히트락) 자사 편성 vs 지정 경쟁사 편성 비교. 가중분 기준.",
                 "matchRule": "브랜드/정제상품명/상품명 3개 컬럼 OR 검색 + 별칭(alias) 리스트 기반. 벤더마다 비교 경쟁사가 다름."},
        "vendors": vendors_out
    }


def build_heatmap(kt, ssg):
    def rows(dfc, co):
        namecol = '정제상품명 (아이템명)' if '정제상품명 (아이템명)' in dfc.columns else '정제상품명(아이템명)'
        out = []
        for _, r in dfc.iterrows():
            out.append({"co": co, "brand": r['brand'], "name": r[namecol], "cat": r['MD분류'],
                       "dow": int(r['DOW']), "hour": int(r['HOUR']), "w": round(float(r['가중분']),1)})
        return out
    kt3 = kt[kt['MD분류'].isin(CATS3)]
    ssg3 = ssg[ssg['MD분류'].isin(CATS3)]
    broadcasts = rows(kt3, 'KT알파') + rows(ssg3, 'SSG')
    return {"meta": {"period": None, "companies": ["KT알파","SSG"], "categories": CATS3, "updated": None,
                     "note": "HOUR 컬럼 기준 방송 시작 시(0~23시), 가중분 단위 합산. DATE 요일 기준 월~일."},
           "days": DOW_LABELS, "broadcasts": broadcasts}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--today', required=True, help='YYYY-MM-DD, 이번 주 데이터 마지막 날짜(일요일)')
    ap.add_argument('--outdir', default='output')
    args = ap.parse_args()
    today = pd.Timestamp(args.today)

    import os
    os.makedirs(args.outdir, exist_ok=True)

    # ---- PGM (연도 필터 있는 버전 / 없는 버전 둘 다 필요) ----
    raw = pd.read_csv('pgm.csv', encoding='utf-8-sig', low_memory=False, dtype=str)
    c = raw.columns.tolist()
    cols = dict(YEAR=c[0], MON=c[1], WK=c[2], DATE=c[3], TIME=c[5], CAT=c[12], BRAND=c[11], WMIN=c[14],
               SALES=c[17], REV=c[21], MARGIN=c[39], MD=c[70], SHIN=c[71], TEAM=c[72])
    team_all = raw[raw[cols['TEAM']] == '가전팀'].copy()
    for col in [cols['WMIN'], cols['SALES'], cols['REV'], cols['MARGIN']]:
        team_all[col] = num(team_all[col])
    team_all['MONI'] = pd.to_numeric(team_all[cols['MON']], errors='coerce')
    team_all['dt'] = pd.to_datetime(team_all[cols['DATE']], errors='coerce')
    team_all['DOW'] = team_all['dt'].dt.dayofweek
    team_all['HOUR'] = pd.to_numeric(team_all[cols['TIME']].astype(str).str.slice(0,2), errors='coerce').fillna(0).astype(int)
    team_all['brand'] = team_all[cols['BRAND']].apply(brand_of)

    f = team_all[team_all[cols['YEAR']] == '2026'].copy()
    f['MONI'] = f['MONI'].astype(int)
    f3 = f[f[cols['CAT']].isin(CATS3)].copy()  # 대형가전 제외 (경쟁사 비교용)

    monthData = build_monthData(f, cols)
    weeklyData = build_weeklyData(team_all, cols)
    mdData = build_mdData(f, cols)
    shinSummary, shinMonthly, shinTrendData, shin_products = build_shin(f, cols)
    MD_DATA, SHIN_DATA = build_newproduct(f, cols, shin_products)
    search_data = build_search(f, cols)
    vendorConcentration = build_vendor_concentration(f, cols)
    shinSettle = build_shin_settle(f, cols)

    cur = json.load(open('current_dashboard-data.json')) if __import__('os').path.exists('current_dashboard-data.json') else {}
    dashboard_data = {
        "monthData": monthData, "mdData": mdData, "shinTrendData": shinTrendData,
        "kpiTarget": cur.get('kpiTarget', {"sales": 1235, "revenue": 403, "margin": 323}),
        "shinSummary": shinSummary, "shinMonthly": shinMonthly,
        "weightTargets": cur.get('weightTargets', {}),  # weight_targets.csv 갱신 시 별도 반영 필요
        "weeklyData": weeklyData,
        "vendorConcentration": vendorConcentration,
        "shinSettle": shinSettle,
    }
    json.dump(dashboard_data, open(f'{args.outdir}/dashboard-data.json','w'), ensure_ascii=False, indent=1)
    json.dump({"MD_DATA": MD_DATA, "SHIN_DATA": SHIN_DATA}, open(f'{args.outdir}/newproduct-data.json','w'), ensure_ascii=False, indent=1)
    json.dump(search_data, open(f'{args.outdir}/search-data.json','w'), ensure_ascii=False, indent=1)

    # ---- 경쟁사 (Excel 필요) ----
    try:
        xl = pd.ExcelFile('competitor.xlsx')
        kt, ssg = xl.parse('KT알파'), xl.parse('SSG')
        def prep(dfc):
            dfc = dfc.copy()
            dfc['DATE'] = pd.to_datetime(dfc['DATE'], errors='coerce')
            dfc = dfc[dfc['DATE'] <= today]
            dfc['가중분'] = pd.to_numeric(dfc['가중분'], errors='coerce').fillna(0)
            dfc['brand'] = dfc['당사브랜드'].fillna(dfc['브랜드'])
            dfc['MON'] = dfc['DATE'].dt.month
            dfc['DOW'] = dfc['DATE'].dt.dayofweek
            dfc['HOUR'] = pd.to_numeric(dfc['HOUR'], errors='coerce').fillna(0).astype(int)
            return dfc
        kt, ssg = prep(kt), prep(ssg)
        competitor_data = build_competitor(f3, cols, kt, ssg)
        period_str = f"2026년 1월 ~ {today.month}월(~{today.month}/{today.day})"
        competitor_data['meta']['period'] = period_str
        competitor_data['meta']['updated'] = today.strftime('%Y.%m.%d')
        json.dump(competitor_data, open(f'{args.outdir}/competitor-data.json','w'), ensure_ascii=False, indent=1, allow_nan=False)

        heatmap_data = build_heatmap(kt, ssg)
        heatmap_data['meta']['period'] = period_str
        heatmap_data['meta']['updated'] = today.strftime('%Y.%m.%d')
        json.dump(heatmap_data, open(f'{args.outdir}/product-heatmap-data.json','w'), ensure_ascii=False, indent=1, allow_nan=False)

        # ---- 핵심 협력사 (마이하우스/코지마/히트락) ----
        corevendor_data = build_corevendor(f, cols, kt, ssg, today)
        corevendor_data['meta']['period'] = period_str
        corevendor_data['meta']['updated'] = today.strftime('%Y.%m.%d')
        json.dump(corevendor_data, open(f'{args.outdir}/corevendor-data.json','w'), ensure_ascii=False, indent=1, allow_nan=False)
        print("⚠ corevendor-data.json의 risk는 침투율로 자동 산출되지만 trendLabel/flags는 빈 값입니다 —",
              "appliance-corevendor-77ac0c66.html 배포 전 지미가 직접 채워넣을 것 (다른 수동 영역과 동일한 패턴).")
    except FileNotFoundError:
        print("competitor.xlsx 없음 — 경쟁사 데이터는 건너뜀 (경쟁사 갱신 없는 주는 정상)")

    print("완료:", args.outdir, "폴더 확인")


if __name__ == '__main__':
    main()
