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
   - 오버뷰 페이지 `#weeklyCompareInsights` — 주간 비교 인사이트 텍스트는 항상 지미가 데이터 보고
     수동 작성 (사용자 명시적 선호, 규칙기반 자동생성 금지)
   - `weightTargets`는 연초 1회 세팅 이후 보통 안 바뀜. 바뀌면 weight_targets.csv 다시 받아서
     `current_dashboard-data.json`의 `weightTargets` 자리 대신 새 값 반영
   - **매주 업데이트 후 다른 페이지에도 하드코딩된 부분이 남아있는지 한 번씩 훑어볼 것.** 지금까지
     오버뷰(인사이트 텍스트는 원래 의도된 수동 영역), 상품랭킹(TOP10), MD실적(요약카드) 세 군데에서
     하드코딩을 발견했다 — 다른 페이지에도 비슷한 게 더 있을 수 있으니 `grep -n "억\|만"` 정도로
     JSON fetch 콜백 밖에 있는 숫자가 있는지 가끔 점검.

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
