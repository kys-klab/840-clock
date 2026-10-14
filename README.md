# 840 시계

수능 국어 80분 실전 모의고사용 웹 시계. 지문(세트) 9개 버튼으로 세트별 소요 시간을 기록하고, OMR 입력·채점·설문·결과 이미지 저장·선생님께 보내기까지 한다.

- `public/` 학생이 여는 웹 페이지(홈 화면 앱, 오프라인 캐시 포함). Cloudflare Pages가 이 폴더만 배포한다.
- `apps-script/Code.gs` 구글 시트 '학생 기록 전송(응답)'에 붙는 Apps Script. 기록 받기('840 기록' 탭)와 정답 보관·채점('840 회차' 탭). 정답은 웹 페이지에 넣지 않는다.

## 배포

- 웹: Cloudflare Pages, Framework None, Build command 없음, Build output directory `public`. main에 올리면 자동 반영.
- Apps Script: 코드를 바꾸면 배포 → 배포 관리 → 새 버전으로 다시 배포(주소 유지). 웹 앱 주소는 `public/index.html`의 `SEND_URL`.
- 페이지를 고치면 `public/sw.js`의 `CACHE` 이름을 올려야 홈 화면 앱이 새 버전을 받는다(페이지 자체는 네트워크 우선이라 대개 바로 반영).
