# 가전팀 대시보드 — 데이터 생성 스크립트

매주 월요일 대시보드 업데이트 시 사용. 이 폴더가 생기기 전까지는 매번 대화 기록을 검색해서
로직을 기억으로 재현했음 (2026-07-20 이전). 이제부터는 스크립트를 그대로 실행하면 된다.

## 매주 실행 절차

1. **원본 데이터 3개 다운로드** (Google Drive)
   - PGM 실적: fileId `1figvjnk_pR6sq3cc2vC-rs_29K92GnPTIddEdzQCUyY` → `pgm.csv`로 저장
   - 경쟁사 편성: fileId `1qMHEjeHpFbhxDrdQqmBirOEKI9l8S0EJUUe4Rz_lZJ8` (시트 KT알파/SSG) → `competitor.xlsx`로 저장
   - 가중분 목표: fileId `1dTDL0L-ZilM7mBYq8QZ85iyE3gvtBygrsEByLs8dVOU` → 변경 있을 때만 확인
   - (파싱 방법은 이 저장소가 아니라 Claude 메모리에 기록돼 있음 — Google Drive 다운로드 결과가
     `json.loads(raw[0]['text'])['content']` base64 이중 인코딩이라는 점)

2. **현재 배포본 dashboard-data.json도 받아둔다** (kpiTarget/weightTargets 고정값 유지용)
   ```
   GET https://api.github.com/repos/huiju9388/-/contents/data/dashboard-data.json?ref=main
   → current_dashboard-data.json 으로 저장
   ```

3. **스크립트 실행**
   ```
   python3 build_all.py --today 2026-07-19   # 이번 주 일요일 날짜
   ```
   `pgm.csv`, `competitor.xlsx`, `current_dashboard-data.json`이 스크립트와 같은 폴더에 있어야 함.
   `output/` 폴더에 5개 JSON 생성됨.

4. **검증 (필수)** — 배포 전에 항상:
   - `output/dashboard-data.json`의 `monthData` 1~(지난달)월이 현재 배포본과 완전히 일치하는지 확인
   - 불일치하면 절대 배포하지 말고 원인 파악 먼저

5. **GitHub 배포** — `deploy.py` 참고 (Contents API로 SHA 재조회 후 PUT)

6. **스크립트가 자동으로 안 하는 것** (매번 손으로 확인/작업)
   - HTML 5개 페이지의 날짜 태그 `(~M/D)` 갱신 (sed로 일괄 치환 가능)
   - `appliance-ranking-77ac0c66.html`의 TOP10 카드 재생성 (search-data.json 상위 10개 기준,
     `#top10-list` ~ `<!-- /top10-list -->` 사이 블록 교체)
   - **`appliance-md-77ac0c66.html`의 상단 요약 카드 5개(권오석/마영호/김응도/임동진/백혜정)** —
     JSON이 아니라 HTML에 직접 박힌 하드코딩 값(취급고/방송횟수/가중분/분당취급고/한계이익/분당한계이익).
     `<div class="kpi-card">` 블록 5개를 매주 `mdData`의 `total` 값으로 교체 필요. 안 하면 몇 주째
     예전 숫자가 그대로 남는다 (2026-07-20에 실제로 이 문제를 겪음 — 최대 23.7억 차이). 분당취급고
     1위 트로피(🏆) 위치도 매번 재확인. 정수란 인수인계 이력 문구(과거 고정 기록)는 갱신 대상 아님.
   - **`appliance-md-77ac0c66.html`의 "분당취급고 비교" 바 차트(요약카드 아래, MD별 월별 추이 차트
     다음 섹션)** — 이것도 하드코딩. `.cat-row` 5개 각각의 `.cat-bar-fill` width%(최댓값 대비 비율)와
     `.cat-bar-label` 텍스트(예: "281.4만원/분 🏆")를 `mdData`의 `total.perMin` 기준으로 갱신 필요.
     🏆(1위)·🥈(2위) 이모지 위치도 순위 바뀌면 함께 이동. 2026-08-03에 요약카드만 갱신하고 이 섹션은
     놓칠 뻔했다 — 매주 요약카드 갱신할 때 반드시 같이 확인할 것.
   - **`appliance-dashboard-77ac0c66.html`의 `#weeklyCompareInsights`(주간 실적 비교 리뷰
     인사이트 텍스트)** — 최신 완료 주차 vs 직전 주차의 `weeklyData`(total/cats)를 비교해서
     증감율을 직접 계산 후 "개선된 점"/"부족한 점" 각 3~4개 불릿으로 수동 작성. 인사이트 제목의
     "(N월N주차)" 라벨도 매주 최신 주차로 갱신.
   - 오버뷰 페이지 `#weeklyCompareInsights` — 주간 비교 인사이트 텍스트는 항상 지미가 데이터 보고
     수동 작성 (사용자 명시적 선호, 규칙기반 자동생성 금지)
   - `weightTargets`는 연초 1회 세팅 이후 보통 안 바뀜. 바뀌면 weight_targets.csv 다시 받아서
     `current_dashboard-data.json`의 `weightTargets` 자리 대신 새 값 반영
   - **매주 업데이트 후 다른 페이지에도 하드코딩된 부분이 남아있는지 한 번씩 훑어볼 것.** 지금까지
     오버뷰(인사이트 텍스트는 원래 의도된 수동 영역), 상품랭킹(TOP10), MD실적(요약카드) 세 군데에서
     하드코딩을 발견했다 — 다른 페이지에도 비슷한 게 더 있을 수 있으니 `grep -n "억\|만"` 정도로
     JSON fetch 콜백 밖에 있는 숫자가 있는지 가끔 점검.

## 핵심 협력사 트래커 (2026-07-28 신설)

`appliance-corevendor-77ac0c66.html` / `data/corevendor-data.json` — 마이하우스·코지마·히트락(벤더사 에코센스)
3개 핵심 협력사의 자사 편성 vs 지정 경쟁사 편성을 월별로 비교. 마이하우스·코지마는 KT알파, 히트락은 SSG와 비교
(벤더마다 실제 경쟁 채널이 다름).

- `build_all.py`의 `VENDOR_DEFS`에 벤더별 별칭(aliases) 리스트로 관리. 새 별칭 발견 시 코드 아래쪽이 아니라
  이 리스트에 한 줄만 추가하면 됨 (예: 벤더사명과 판매 브랜드명이 다른 케이스 — 에코센스=히트락처럼).
- `build_corevendor()`가 브랜드/정제상품명/상품명 3개 컬럼을 alias로 OR 검색 — 편성코드명(브랜드 태그) 한
  컬럼만 보면 벤더명↔브랜드명 불일치를 놓치니 주의.
- **수동 갱신 필요**: `risk`(침투율로 자동 산출됨)와 달리 `trendLabel`, `flags`는 스크립트가 빈 값으로만
  채움 — 매주 지미가 수치 보고 직접 채워넣어야 함 (오버뷰 `weeklyCompareInsights`와 같은 패턴).
- 사이드바 nav에 "🛡️ 핵심 협력사" 링크가 8개 페이지 전체에 추가되어 있음 — 신규 페이지 만들 때는 항상
  전체 페이지 nav에 링크 추가하는 것 잊지 말 것.

## 핵심 규칙 (절대 위반 금지)
- 취급고는 **항상 컬럼17 (예상 취급액, VAT포함)**. 컬럼18(V-)은 사용 금지.
- 신상품 비중 공식 지표(팀 목표 관리용)는 **단순 태그 기준**. "생애귀속"(같은 상품명이 한번이라도
  신상품 태그를 받으면 전체 생애 실적을 그 상품에 귀속)은 `newproduct-data.json`의 MD 기여도
  분석에만 사용 — 이 두 방식을 섞어서 team-level 신상품 비중으로 쓰면 안 됨.
- 대형가전은 KT알파/SSG 카테고리 매핑이 없어서 경쟁사 관련 파일(competitor-data.json,
  product-heatmap-data.json)에서는 제외. dashboard-data.json/search-data.json에는 포함.
- 8월부터 MD 카테고리 고정 담당제 폐지됨. `MDMAP`/`CATLABEL`의 "cat" 설명 문구는 더 이상
  실제 담당을 의미하지 않을 수 있으니 매번 형님께 최신 업무 분장 확인 후 필요시 스크립트 상단 갱신.

## 파일 구성
- `build_all.py` — 5개 JSON 전체 생성 (핵심 로직)
- `deploy.py` — GitHub Contents API 업로드 헬퍼
