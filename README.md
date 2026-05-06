# 📊 Market Stats Monthly — 매월 자동 업데이트 대시보드

S&P 500, NASDAQ 100, KOSPI 월간 수익률을 매월 1일 자동 업데이트하는 대시보드.

- ✅ **완전 자동** — 매월 1일 09:00 KST에 GitHub Actions가 자동 실행
- ✅ **무료** — GitHub Actions 무료 분량 안에서 충분 (이 작업은 1분 안에 끝남)
- ✅ **항상 같은 URL** — 북마크 하나만 저장하면 끝
- ✅ **로컬 PC 안 켜져 있어도 됨** — 클라우드에서 실행

---

## 🚀 셋업 가이드 (1회, 약 30분)

### 사전 준비물

- GitHub 계정 1개 (개인 계정 권장 — 회사 망에서 차단되어 있을 수 있음)
- 모바일 핫스팟 또는 집 인터넷 (셋업할 때만)
- 브라우저

### Step 1: GitHub 저장소 만들기 (5분)

1. https://github.com 접속 → 우상단 `+` → `New repository`
2. **Repository name**: `market-stats-monthly` (자유롭게)
3. **Public** 선택 (Private도 가능하지만 GitHub Pages 무료 사용하려면 Public)
4. **README** 체크박스는 비워두기
5. `Create repository` 클릭

### Step 2: 파일 업로드 (10분)

**옵션 A — 웹 업로드 (가장 간단)**

1. 방금 만든 저장소 페이지에서 `uploading an existing file` 링크 클릭
2. 이 ZIP 파일 압축 풀고 **모든 파일을 통째로 드래그 앤 드롭**
   - `update_stats.py`
   - `requirements.txt`
   - `templates/index_template.html`
   - `.github/workflows/monthly-update.yml`
   - `README.md`
3. 하단 `Commit changes` 클릭

> ⚠️ `.github` 폴더는 숨김 폴더라 일부 파일 매니저에서 안 보일 수 있음. macOS는 `Cmd+Shift+.`, Windows는 폴더 옵션에서 "숨김 항목 보기" 켜기.

**옵션 B — git CLI (Git 익숙한 경우)**

```bash
git clone https://github.com/YOUR_USERNAME/market-stats-monthly.git
cd market-stats-monthly
# ZIP 압축 풀어서 모든 파일 복사
git add .
git commit -m "초기 셋업"
git push
```

### Step 3: GitHub Pages 활성화 (3분)

1. 저장소 페이지에서 `Settings` 탭
2. 좌측 메뉴 `Pages`
3. **Source**: `Deploy from a branch`
4. **Branch**: `main` 선택, 폴더는 `/docs` 선택
5. `Save` 클릭

> 1~2분 후 페이지 상단에 `Your site is live at https://YOUR_USERNAME.github.io/market-stats-monthly/` 메시지 표시. 이 URL을 북마크.

### Step 4: 첫 실행 (5분)

자동 스케줄을 기다리지 말고 지금 한 번 돌려보자.

1. 저장소 페이지 → `Actions` 탭
2. 좌측 `Monthly Market Stats Update` 클릭
3. 우측 `Run workflow` 버튼 → `Run workflow` 확인
4. 약 1~2분 후 초록 체크 ✅ 표시되면 성공

> 빨간 X가 뜨면 클릭해서 어떤 step에서 실패했는지 로그 확인. 가장 흔한 이슈: yfinance 일시적 다운(재실행하면 됨), `permissions: contents: write` 누락 (이미 설정됨).

### Step 5: 결과 확인 (2분)

- 브라우저에서 `https://YOUR_USERNAME.github.io/market-stats-monthly/` 접속
- 대시보드 5개 탭 (요약 / 연간 / 월간 통계 / 계절성 / 상관관계) 정상 작동 확인
- "📊 Excel 데이터 파일 다운로드" 버튼으로 xlsx 다운로드 가능 확인

---

## 📅 자동 실행 일정

```
매월 1일 09:00 KST (= UTC 00:00)
```

크론 표기: `0 0 1 * *`

다른 시간 원하면 `.github/workflows/monthly-update.yml` 수정. 예:
- 매월 1일·15일 둘 다: `0 0 1,15 * *`
- 매주 월요일 09:00 KST: `0 0 * * 1`

---

## 🔧 유지보수 (연 1~2회)

연말이나 큰 정책 이벤트 후 `update_stats.py`의 두 딕셔너리만 수동 업데이트:

```python
FED_BOK_RATES = {
    ...,
    2026: (3.75, 2.50),  # 연말에 마지막 값 확인 후 업데이트
    2027: (?.??, ?.??),  # 새해 1월에 새 항목 추가
}

YEAR_EVENTS = {
    ...,
    2027: "올해 주요 이벤트 한 줄 메모",  # 연말에 추가
}
```

지수 가격·통계는 yfinance가 매월 자동으로 받아오므로 손댈 필요 없음.

---

## 🐛 트러블슈팅

### "yfinance 다운로드 실패"
- yfinance가 일시적으로 응답 없을 때 발생. Actions 탭에서 `Re-run all jobs` 클릭하면 대부분 해결.

### "Pages 사이트가 404"
- Settings > Pages에서 Source가 `main / docs`로 정확히 설정됐는지 재확인
- 첫 배포는 5~10분 걸릴 수 있음

### "회사망에서 GitHub Pages URL 차단됨"
- IT 정책상 차단되어 있으면 개인 휴대폰 핫스팟 또는 집에서 접속

### "이메일 알림 받고 싶음"
- GitHub 우상단 프로필 → Settings > Notifications > Actions에서 활성화

### "Excel 파일이 매월 동일하게 보임"
- 브라우저 캐시 때문일 가능성. `Ctrl+F5` (강력 새로고침) 또는 시크릿 모드에서 접속

---

## 📂 프로젝트 구조

```
market-stats-monthly/
├── README.md                          # 이 문서
├── requirements.txt                   # Python 의존성
├── update_stats.py                    # 메인 스크립트 (데이터 수집·통계·생성)
├── templates/
│   └── index_template.html            # HTML 대시보드 템플릿
├── .github/
│   └── workflows/
│       └── monthly-update.yml         # GitHub Actions 스케줄
└── docs/                              # ← 자동 생성 (GitHub Pages가 서빙)
    ├── index.html                     # 대시보드 (매월 자동 갱신)
    └── data.xlsx                      # Excel 데이터 (매월 자동 갱신)
```

---

## 💡 확장 아이디어 (필요 시)

- **이메일/슬랙 알림**: workflow에 `actions/send-mail` 추가
- **추가 지수**: `update_stats.py`의 `TICKERS` 딕셔너리에 항목 추가 (예: `^N225` 일본 닛케이, `000001.SS` 상해종합)
- **환율 보정**: `KRW=X` 추가해서 USD 자산을 KRW 기준으로 비교
- **drawdown 차트**: 추가 통계 함수 작성 후 템플릿에 섹션 추가

---

## 📝 라이선스 / 면책

이 자료는 정보 제공 목적이며 투자 추천이 아님. 본인의 판단·책임 하에 사용.
yfinance는 Yahoo Finance 비공식 API라 서비스 변경 시 영향 받을 수 있음.
