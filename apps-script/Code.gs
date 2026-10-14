// 840 시계 → 구글 시트 (기록 받기 + 정답 보관·채점)
// 붙여 넣을 곳: 시트 '학생 기록 전송(응답)' → 확장 프로그램 → Apps Script → 이 내용으로 Code.gs 전체 교체
// 배포: 배포 → 새 배포 → 유형 '웹 앱' → 실행 사용자 '나', 액세스 권한 '모든 사용자' → 배포 → 웹 앱 URL(…/exec) 복사
// 코드를 고친 뒤에는 배포 → 배포 관리 → 연필 → 버전 '새 버전' → 배포 (주소는 그대로)
//
// 탭 두 개를 씀(없으면 처음 접속 때 자동으로 만들어짐)
// - '840 회차': 정답을 등록하는 곳. 한 줄 = 한 회차. 정답은 여기에만 있고 시계 페이지 코드에는 들어 있지 않음
//              (학생이 채점을 누르면 그때 이 스크립트가 그 학생에게 정답을 돌려줌)
// - '840 기록': 학생 기록. 같은 학생의 같은 시험(기록 ID)은 다시 보내면 그 줄을 덮어씀
// 등급컷: '840 회차' 탭 맨 오른쪽 두 칸에 1~8등급 원점수 컷을 쉼표로 (예: 92,85,77,68,58,47,37,27)
// 오답률: 같은 회차를 채점해서 보낸 학생들의 답안으로 문항마다 계산(학생이 '등급컷 오답률' 카드를 열 때)

const REC = '840 기록';
const EXAM = '840 회차';
const FIXED = ['기록 ID', '받은 시각', '상태', '이름', '회차'];
const SETS = ['독서론', '독서 세트 1', '독서 세트 2', '독서 세트 3', '현대운문', '현대소설', '고전산문', '고전시가', '선택'];
const EXAM_HEAD = ['회차 이름', '공통 정답(1~34)', '화법과 작문 정답(35~45)', '언어와 매체 정답(35~45)',
  '3점 문항(공통)', '3점 문항(화작)', '3점 문항(언매)'].concat(SETS.map(n => n + ' 문항'), ['등급컷(화작)', '등급컷(언매)']);

// ── 웹 앱 입구 ──
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  if (action === 'exams') return out({ ok: true, exams: listExams() });
  return out({ ok: true, app: '840 시계' });
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return out({ ok: false, error: 'bad_json' }); }
  if (d && d.action === 'grade') return out(grade(d));
  if (d && d.action === 'stats') return out(stats(d));
  return out(saveRecord(d));
}

// ── 회차 탭 ──
function examSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(EXAM);
  if (!sh) {
    sh = ss.insertSheet(EXAM);
    sh.getRange(1, 1, 2, EXAM_HEAD.length).setNumberFormat('@').setValues([EXAM_HEAD, [
      '이감 파이널Ⅱ 5차', '4134352233151542513123232525415443', '42335541132', '13354152524',
      '3,8,12,17,21,24,30,34', '36,45', '36,40',
      '1~3', '4~9', '10~13', '14~17', '22~26', '18~21', '31~34', '27~30', '35~45']]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, EXAM_HEAD.length).setFontWeight('bold');
    sh.getRange('A:R').setNumberFormat('@');
  }
  // 예전에 만든 탭에는 등급컷 칸 머리글이 없으니 채워 둠
  const head = sh.getRange(1, 1, 1, EXAM_HEAD.length).getDisplayValues()[0];
  if (head.some((h, i) => String(h).trim() !== EXAM_HEAD[i])) {
    sh.getRange(1, 1, 1, EXAM_HEAD.length).setValues([EXAM_HEAD]).setFontWeight('bold');
    sh.getRange(2, 17, Math.max(1, sh.getMaxRows() - 1), 2).setNumberFormat('@');
  }
  return sh;
}

function readExams() {
  const sh = examSheet();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const nums = s => String(s).split(/[^0-9]+/).filter(Boolean).map(Number);
  return sh.getRange(2, 1, last - 1, EXAM_HEAD.length).getDisplayValues()
    .filter(r => String(r[0]).trim())
    .map(r => {
      const ranges = {};
      SETS.forEach((n, i) => { const v = String(r[7 + i]).replace(/\s/g, '').replace(/[-–]/g, '~'); if (v) ranges[n] = v; });
      return {
        name: String(r[0]).trim(),
        key: { common: String(r[1]).replace(/\D/g, ''), hwa: String(r[2]).replace(/\D/g, ''), eon: String(r[3]).replace(/\D/g, '') },
        three: { common: nums(r[4]), hwa: nums(r[5]), eon: nums(r[6]) },
        cuts: { hwa: nums(r[16]).slice(0, 8), eon: nums(r[17]).slice(0, 8) },
        ranges: ranges
      };
    });
}

// 시계 시작 화면용: 회차 이름과 세트별 문항 범위만(정답 없음)
function listExams() {
  return readExams().map(x => ({ name: x.name, ranges: x.ranges, gradable: x.key.common.length === 34 }));
}

const SUBJ_NAME = { hwa: '화법과 작문', eon: '언어와 매체' };
const simple = t => String(t || '').replace(/\s+/g, '').toLowerCase();

// ── 채점용 정답: 학생이 시험을 끝내고 채점을 누를 때만 그 회차·선택 과목 정답을 돌려줌 ──
function grade(d) {
  const x = readExams().find(e => simple(e.name) === simple(d.exam));
  if (!x) return { ok: false, error: 'unknown_exam' };
  if (d.subj !== 'hwa' && d.subj !== 'eon') return { ok: false, error: 'no_subject' };
  if (x.key.common.length !== 34 || x.key[d.subj].length !== 11) return { ok: false, error: 'key_incomplete' };
  return { ok: true, key: (x.key.common + x.key[d.subj]).split('').map(Number), three: x.three.common.concat(x.three[d.subj]) };
}

// ── 등급컷·오답률: 채점한 학생이 볼 때만. 같은 회차를 채점해서 보낸 학생들의 답안으로 계산 ──
function stats(d) {
  const x = readExams().find(e => simple(e.name) === simple(d.exam));
  if (!x) return { ok: false, error: 'unknown_exam' };
  const subj = d.subj === 'eon' ? 'eon' : 'hwa', subjName = SUBJ_NAME[subj];
  const key = (x.key.common + x.key[subj]).split('').map(Number);
  const cuts = x.cuts[subj].length === 8 ? x.cuts[subj] : [];
  const wrong = new Array(45).fill(0);
  let n = 0, nSel = 0;
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REC);
  if (sh && sh.getLastRow() > 1 && key.length === 45) {
    const all = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getDisplayValues();
    const h = all[0].map(v => String(v).trim());
    const cE = h.indexOf('회차'), cS = h.indexOf('선택 과목'), cA = h.indexOf('답안(1~45)'), cP = h.indexOf('원점수');
    if (cE >= 0 && cA >= 0 && cP >= 0) {
      all.slice(1).forEach(r => {
        if (simple(r[cE]) !== simple(x.name)) return;
        if (String(r[cP]).trim() === '') return;                  // 채점 안 한 기록은 빼기
        const ans = String(r[cA]).replace(/\D/g, '');
        if (ans.length !== 45) return;
        const same = String(r[cS]).trim() === subjName;              // 선택 문항은 같은 과목 학생끼리
        n++; if (same) nSel++;
        for (let q = 0; q < 45; q++) {
          if (q >= 34 && !same) continue;
          if (+ans[q] !== key[q]) wrong[q]++;
        }
      });
    }
  }
  const rate = wrong.map((w, q) => { const m = q >= 34 ? nSel : n; return m ? Math.round(w / m * 1000) / 10 : null; });
  return { ok: true, cuts: cuts, n: n, nSel: nSel, rate: rate };
}

// ── 기록 저장 ──
function saveRecord(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (!d || !d.id || !Array.isArray(d.cols)) return { ok: false, error: 'bad_request' };
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(REC) || ss.insertSheet(REC);
    const headers = FIXED.concat((d.headers || []).map(String), ['원문']);
    if (sh.getLastRow() === 0) {
      sh.appendRow(headers);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
    // 점수·시간 같은 숫자는 숫자로(정렬·합계 가능), 나머지는 글자 그대로(회차 '2-5'가 날짜로 바뀌지 않게 앞에 ')
    const clean = v => {
      const t = String(v == null ? '' : v).slice(0, 5000);
      return /^-?\d+(\.\d+)?$/.test(t) ? Number(t) : (t ? "'" + t : '');
    };
    const now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    const row = [clean(d.id), now, clean(d.status), clean(d.name), clean(d.exam)]
      .concat(d.cols.map(clean), [clean(d.text)]);
    const last = sh.getLastRow();
    let at = 0;
    if (last > 1) {
      const ids = sh.getRange(2, 1, last - 1, 1).getValues();
      for (let i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === String(d.id)) { at = i + 2; break; }
    }
    const range = at ? sh.getRange(at, 1, 1, row.length) : sh.getRange(last + 1, 1, 1, row.length);
    range.setValues([row]);
    return { ok: true, row: at || last + 1 };
  } catch (err) {
    return { ok: false, error: String(err) };
  } finally {
    lock.releaseLock();
  }
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
