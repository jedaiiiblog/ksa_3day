# 팀 주간 보고 웹 입력 → PPT 생성 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 담당자가 브라우저 한 화면에 팀원별 주간 보고·공휴일·휴가·요약을 입력하면 기여도를 계산해 상위부서 보고용 PPT(.pptx)를 내려받는 서버 없는 웹 페이지를 만든다.

**Architecture:** 빌드 없는 정적 페이지(`web/`). 순수 로직 `core.js`(날짜·근무일·기여도·경고/차단·상태 검증), 슬라이드 계획·생성 `deck.js`(PptxGenJS 생성자를 인자로 받음), 화면 `app.js`(상태·렌더링·localStorage·백업). `core.js`/`deck.js`는 브라우저(`window.TR`, `window.TRDeck`)와 Node(`module.exports`) 양쪽에서 동작해 `node --test`로 검증한다. PptxGenJS는 `web/vendor/`에 파일로 함께 둔다(오프라인 동작).

**Tech Stack:** 바닐라 JavaScript(ES2020, `<script>` 태그), PptxGenJS 4.0.1(vendored bundle), Node 24 내장 `node:test`, 화면 검증은 내장 브라우저.

**Spec:** `docs/superpowers/specs/2026-09-28-web-report-builder-design.md` (기존 Python 구현 `team_report/`는 수정하지 않는다)

> 이 폴더는 git 저장소가 아니다. "Commit" 단계는 생략하고 "테스트 통과 확인"으로 대신한다.

## Global Constraints

- 근무시간 1인 하루 8시간, 월~금. 개인 월 기준 시간 = (해당 월 월~금 일수 − 공휴일 − 그 팀원 휴가일) × 8.
- 프로젝트당 시간 = 개인 월 기준 시간 ÷ 그 팀원의 프로젝트 수. 개인 기준 % = 100 ÷ 프로젝트 수. 팀 대비 % = 개인 시간 ÷ 팀 합계. 프로젝트 팀 대비 % = 프로젝트당 시간 ÷ 팀 합계.
- 프로젝트 수 = 대상 월 보고 전체에서 정규화(공백 정리 + 소문자화)한 프로젝트 이름의 합집합. 팀원 이름·휴가 이름도 같은 정규화로 비교.
- 날짜는 `YYYY-MM-DD` 문자열, 계산은 UTC. 제출 시각 `YYYY-MM-DDTHH:MM`. 지연 = 제출 시각 > 주차 날짜 + 마감 시각(정확히 마감 시각은 정상). 지연은 표시만.
- 팀원 순서 = 대상 월 보고 중 가장 이른 제출 시각 순, 같으면 입력 순.
- 기여도는 코드가 계산. 요약은 담당자가 직접 입력(AI 호출 없음).
- 표시: 시간·비율 소수 1자리(`53.3h`, `21.1%`). 파일명 `팀_월간보고_YYYY-MM.pptx`, 백업 `팀보고_백업_YYYY-MM.json`. 글꼴 "맑은 고딕", 16:9.
- 슬라이드: 팀 요약 1장 + 팀원별 슬라이드(프로젝트 6개까지, 초과 시 `(계속)`). 요약이 비면 `(요약 없음) ` + 원문 앞 200자(초과 시 `…`).
- 하지 않는 것: 팀원 각자 접속/제출, md 파일 가져오기, 서버·계정, 자동 발송·알림, 반차, 휴가 기간 범위, 순위·인사평가, 월 외 기간, 이미지·차트.
- `file://`에서 열려야 하므로 ES 모듈 import 금지. 사용자 입력은 항상 텍스트로만 DOM에 넣는다(`innerHTML` 금지).
- 기존 파일(`index.html`, `공사스케줄_*.xlsx`, `team_report/`, `tests/`, 기획서)은 건드리지 않는다. 새 파일은 `web/`, `web-tests/`에만 만든다.
- **결정(Ruling):** 스펙 §4의 `vacations: {이름: [날짜]}`는 화면에서 반쯤 채워진 행을 다뤄야 하므로 `vacations: [{name, date}]` 배열로 저장한다(스펙 §4를 같이 수정). 기능은 동일.

## Review Focus

1. **저장 데이터 손상·구버전** — localStorage의 JSON이 깨졌거나 `v`가 다르면 빈 상태로 시작하고 알린다. 원본은 다음 저장 전까지 그대로. (Task 3, 6)
2. **datetime-local 형식 변형** — 브라우저가 `HH:MM:SS`(초 포함)를 줄 수 있음. 초는 무시하고 마감 판정이 흔들리지 않아야 한다. (Task 1)
3. **이름에 HTML 특수문자** — 팀원·프로젝트명에 `<b>x</b>`, `&`가 있어도 화면·PPT에 글자 그대로 나오고 스크립트가 실행되지 않아야 한다. (Task 6, 7)
4. **월 경계** — 12월(23일), 윤년 2월(2028-02는 21일), 다음 해 1월이 정확해야 한다. 다른 달·주말 날짜 입력은 무시. (Task 1)
5. **프로젝트가 많은 팀원 / 이름 표기 차이** — 13개 프로젝트가 6+6+1로 나뉘고, `Alpha Web`/`alpha  web`이 하나로 합쳐지며 합쳤다고 알린다. (Task 2, 4)

---

## File Structure

```
실습/
  web/
    index.html            # 뼈대 + 스크립트 로드 순서: vendor → core → deck → app
    style.css
    core.js               # 순수 로직 (window.TR / module.exports)
    deck.js               # 슬라이드 계획·생성 (window.TRDeck / module.exports)
    app.js                # 화면·상태·저장·백업
    vendor/pptxgen.bundle.js
  web-tests/
    core.test.js
    deck.test.js
    pptx-smoke.test.js
```

전체 테스트: `node --test "web-tests/*.test.js"` (작업 디렉터리 `C:\Users\KSA\Desktop\실습`)

---

### Task 1: core.js — 날짜·이름·근무일·마감 판정 헬퍼

**Files:**
- Create: `web/core.js`
- Test: `web-tests/core.test.js`

**Interfaces:**
- Consumes: 없음
- Produces (`module.exports` / `window.TR`):
  - `normName(s: string): string` — 앞뒤·중복 공백 정리 + 소문자
  - `cleanName(s: string): string` — 공백만 정리(대소문자 유지)
  - `isValidDate(s: any): boolean` — 엄격한 `YYYY-MM-DD`
  - `weekdayOf(date: string): number` — 0(일)~6(토), UTC
  - `workdays(year: number, month: number, holidays: Set<string>, vacations: Set<string>): number`
  - `isValidSubmit(s: any): boolean` — `YYYY-MM-DDTHH:MM`(초 이하는 허용·무시)
  - `isLate(week: string, submittedAt: string, deadline: string): boolean`
  - `lastFriday(year: number, month: number): string`
  - `fmtHours(v: number): string`, `fmtPct(v: number): string`
  - `summaryKey(member: string, project: string): string` — `normName(member) + '|' + normName(project)`
  - 상수 `HOURS_PER_DAY = 8`, `TEAM_SIZE = 5`

- [ ] **Step 1: 실패하는 테스트 작성**

`web-tests/core.test.js`:

```js
const test = require('node:test');
const assert = require('node:assert/strict');
const C = require('../web/core.js');

test('normName / cleanName', () => {
  assert.equal(C.normName('  Alpha   Web '), 'alpha web');
  assert.equal(C.cleanName('  Alpha   Web '), 'Alpha Web');
  assert.equal(C.normName(undefined), '');
});

test('isValidDate is strict', () => {
  assert.ok(C.isValidDate('2026-09-25'));
  assert.ok(C.isValidDate('2028-02-29'));
  assert.ok(!C.isValidDate('2026-02-30'));
  assert.ok(!C.isValidDate('2026-9-5'));
  assert.ok(!C.isValidDate('2026-13-01'));
  assert.ok(!C.isValidDate(''));
  assert.ok(!C.isValidDate(undefined));
});

test('weekdayOf uses calendar weekday', () => {
  assert.equal(C.weekdayOf('2026-09-25'), 5); // 금
  assert.equal(C.weekdayOf('2026-09-27'), 0); // 일
  assert.equal(C.weekdayOf('2026-09-28'), 1); // 월
});

const HOL = new Set(['2026-09-24', '2026-09-25']);
const none = new Set();

test('workdays: weekdays, holidays, vacations', () => {
  assert.equal(C.workdays(2026, 9, none, none), 22);
  assert.equal(C.workdays(2026, 9, HOL, none), 20);
  assert.equal(C.workdays(2026, 9, HOL, new Set(['2026-09-01', '2026-09-02'])), 18);
});

test('workdays: weekend, other-month and duplicate days are not double counted', () => {
  const extra = new Set(['2026-09-26', '2026-09-27', '2026-08-31']);
  assert.equal(C.workdays(2026, 9, new Set([...HOL, ...extra]), extra), 20);
  assert.equal(C.workdays(2026, 9, HOL, new Set(['2026-09-24'])), 20); // 공휴일=휴가 같은 날
});

test('workdays: month boundaries (12월, 윤년 2월, 1월)', () => {
  assert.equal(C.workdays(2026, 12, none, none), 23);
  assert.equal(C.workdays(2028, 2, none, none), 21);
  assert.equal(C.workdays(2027, 1, none, none), 21);
});

test('isValidSubmit accepts seconds, rejects garbage', () => {
  assert.ok(C.isValidSubmit('2026-09-25T17:00'));
  assert.ok(C.isValidSubmit('2026-09-25T17:00:30'));
  assert.ok(!C.isValidSubmit('2026-09-25'));
  assert.ok(!C.isValidSubmit('2026-09-25T25:00'));
  assert.ok(!C.isValidSubmit('2026-02-30T10:00'));
  assert.ok(!C.isValidSubmit(''));
});

test('isLate: deadline boundary, seconds ignored, next-day submit', () => {
  assert.equal(C.isLate('2026-09-25', '2026-09-25T17:00', '17:00'), false);
  assert.equal(C.isLate('2026-09-25', '2026-09-25T17:00:59', '17:00'), false);
  assert.equal(C.isLate('2026-09-25', '2026-09-25T17:01', '17:00'), true);
  assert.equal(C.isLate('2026-09-25', '2026-09-25T09:00', '17:00'), false);
  assert.equal(C.isLate('2026-09-25', '2026-09-28T09:00', '17:00'), true);
});

test('lastFriday', () => {
  assert.equal(C.lastFriday(2026, 9), '2026-09-25');
  assert.equal(C.lastFriday(2026, 10), '2026-10-30');
  assert.equal(C.lastFriday(2026, 2), '2026-02-27');
});

test('formatting and summaryKey', () => {
  assert.equal(C.fmtHours(53.3333), '53.3h');
  assert.equal(C.fmtPct(160 / 760 * 100), '21.1%');
  assert.equal(C.summaryKey(' 김민수 ', 'Alpha  Web'), '김민수|alpha web');
  assert.equal(C.HOURS_PER_DAY, 8);
  assert.equal(C.TEAM_SIZE, 5);
});
```

- [ ] **Step 2: 실패 확인**

Run: `node --test "web-tests/*.test.js"`
Expected: FAIL — `Cannot find module '../web/core.js'`

- [ ] **Step 3: 구현**

`web/core.js`:

```js
(function (root) {
  'use strict';

  const HOURS_PER_DAY = 8;
  const TEAM_SIZE = 5;
  const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;
  const SUBMIT_RE = /^\d{4}-\d{2}-\d{2}T([01]\d|2[0-3]):[0-5]\d/;
  const DEADLINE_RE = /^([01]\d|2[0-3]):[0-5]\d$/;

  const pad = (n) => String(n).padStart(2, '0');
  const cleanName = (s) => String(s == null ? '' : s).trim().replace(/\s+/g, ' ');
  const normName = (s) => cleanName(s).toLowerCase();
  const summaryKey = (member, project) => normName(member) + '|' + normName(project);

  function isValidDate(s) {
    if (typeof s !== 'string' || !DATE_RE.test(s)) return false;
    const [y, m, d] = s.split('-').map(Number);
    const dt = new Date(Date.UTC(y, m - 1, d));
    return dt.getUTCFullYear() === y && dt.getUTCMonth() === m - 1 && dt.getUTCDate() === d;
  }

  function weekdayOf(s) {
    const [y, m, d] = s.split('-').map(Number);
    return new Date(Date.UTC(y, m - 1, d)).getUTCDay();
  }

  function workdays(year, month, holidays, vacations) {
    const last = new Date(Date.UTC(year, month, 0)).getUTCDate();
    let count = 0;
    for (let d = 1; d <= last; d++) {
      const key = `${year}-${pad(month)}-${pad(d)}`;
      const dow = new Date(Date.UTC(year, month - 1, d)).getUTCDay();
      if (dow >= 1 && dow <= 5 && !holidays.has(key) && !vacations.has(key)) count++;
    }
    return count;
  }

  function isValidSubmit(s) {
    return typeof s === 'string' && SUBMIT_RE.test(s) && isValidDate(s.slice(0, 10));
  }

  function isLate(week, submittedAt, deadline) {
    return submittedAt.slice(0, 16) > `${week}T${deadline}`;
  }

  function lastFriday(year, month) {
    for (let d = new Date(Date.UTC(year, month, 0)).getUTCDate(); d >= 1; d--) {
      if (new Date(Date.UTC(year, month - 1, d)).getUTCDay() === 5) return `${year}-${pad(month)}-${pad(d)}`;
    }
    return '';
  }

  const fmtHours = (v) => v.toFixed(1) + 'h';
  const fmtPct = (v) => v.toFixed(1) + '%';

  // ---- exports ----
  const api = {
    HOURS_PER_DAY, TEAM_SIZE, DEADLINE_RE,
    normName, cleanName, summaryKey, isValidDate, weekdayOf, workdays,
    isValidSubmit, isLate, lastFriday, fmtHours, fmtPct,
  };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.TR = api;
})(typeof self !== 'undefined' ? self : this);
```

- [ ] **Step 4: 통과 확인**

Run: `node --test "web-tests/*.test.js"`
Expected: 9 tests pass, 0 fail

---

### Task 2: core.js — buildModel (기여도·순서·경고·차단)

**Files:**
- Modify: `web/core.js` (헬퍼 아래, `// ---- exports ----` 위에 함수 추가하고 exports에 `buildModel` 추가)
- Test: `web-tests/core.test.js` (끝에 추가)

**Interfaces:**
- Consumes: Task 1의 모든 헬퍼
- Produces: `buildModel(state): Model`
  - `state` = `{v, year, month, deadline:'HH:MM', holidays:[{date,name}], vacations:[{name,date}], members:[{name, reports:[{week, submittedAt, projects:[{name, body}]}]}], summaries:{team:string, projects:{[summaryKey]: string}}}`
  - `Model` = `{year, month, teamSummary: string, members: MemberModel[], warnings: string[], blockers: string[]}`
  - `MemberModel` = `{name, projects:[{name, bodies:string[], summary:string}], lateWeeks:string[], earliest:string, workdays:number, baseHours:number, perProjectHours:number, personalPct:number, teamSharePct:number, projectSharePct:number}` (`members`는 이미 저장 시각 순으로 정렬됨)
  - 연·월이 잘못이면 `members: []`로 즉시 반환(blockers에 사유)

- [ ] **Step 1: 실패하는 테스트 추가**

`web-tests/core.test.js` 끝에 추가:

```js
function state(over = {}) {
  return Object.assign({
    v: 1, year: 2026, month: 9, deadline: '17:00',
    holidays: [{ date: '2026-09-24', name: '추석' }, { date: '2026-09-25', name: '추석' }],
    vacations: [], members: [], summaries: { team: '', projects: {} },
  }, over);
}
function member(name, projects, opts = {}) {
  return {
    name,
    reports: [{
      week: opts.week || '2026-09-25', submittedAt: opts.at || '2026-09-25T10:00',
      projects: projects.map((p) => (typeof p === 'string' ? { name: p, body: 'x' } : p)),
    }],
  };
}
const names = (n, prefix) => Array.from({ length: n }, (_, i) => `${prefix}${i + 1}`);
const vac = (name, ...dates) => dates.map((date) => ({ name, date }));

function fiveMembers() {
  return state({
    members: [
      member('A', names(4, 'a')), member('B', names(2, 'b'), { at: '2026-09-25T11:00' }),
      member('C', names(1, 'c'), { at: '2026-09-25T12:00' }), member('D', names(3, 'd'), { at: '2026-09-25T13:00' }),
      member('E', names(5, 'e'), { at: '2026-09-25T14:00' }),
    ],
    vacations: [...vac('D', '2026-09-01', '2026-09-02'), ...vac('E', '2026-09-01', '2026-09-02', '2026-09-03')],
  });
}

test('buildModel: spec example (160h → 21.1%)', () => {
  const m = C.buildModel(fiveMembers());
  assert.deepEqual(m.blockers, []);
  const a = m.members[0];
  assert.equal(a.name, 'A');
  assert.equal(a.workdays, 20);
  assert.equal(a.baseHours, 160);
  assert.equal(a.perProjectHours, 40);
  assert.equal(a.personalPct, 25);
  assert.ok(Math.abs(a.teamSharePct - 160 / 760 * 100) < 1e-9);
  assert.ok(Math.abs(a.projectSharePct - 40 / 760 * 100) < 1e-9);
  assert.equal(m.members[3].baseHours, 144);
  assert.equal(m.members[4].baseHours, 136);
  const sum = m.members.reduce((s, x) => s + x.teamSharePct, 0);
  assert.ok(Math.abs(sum - 100) < 1e-9);
  assert.ok(m.members[4].teamSharePct < a.teamSharePct);
  assert.ok(!m.warnings.some((w) => w.includes('5명')));
});

test('buildModel: projects merge across weeks ignoring case/space, bodies in week order, warns', () => {
  const s = state({
    members: [{
      name: 'A',
      reports: [
        { week: '2026-09-25', submittedAt: '2026-09-25T10:00', projects: [{ name: 'alpha  web', body: 'w2' }, { name: 'Beta', body: 'b' }] },
        { week: '2026-09-18', submittedAt: '2026-09-18T10:00', projects: [{ name: 'Alpha Web', body: 'w1' }] },
      ],
    }],
    summaries: { team: 'T', projects: { 'a|alpha web': '요약됨' } },
  });
  const m = C.buildModel(s);
  const a = m.members[0];
  assert.deepEqual(a.projects.map((p) => p.name), ['Alpha Web', 'Beta']);
  assert.deepEqual(a.projects[0].bodies, ['w1', 'w2']);
  assert.equal(a.projects[0].summary, '요약됨');
  assert.equal(m.teamSummary, 'T');
  assert.ok(m.warnings.some((w) => w.includes('같은 프로젝트로 합쳤습니다')));
});

test('buildModel: late flags and member order by earliest submit', () => {
  const s = state({
    members: [
      { name: 'B', reports: [
        { week: '2026-09-18', submittedAt: '2026-09-18T09:00', projects: [{ name: 'P', body: '' }] },
        { week: '2026-09-25', submittedAt: '2026-09-25T17:01', projects: [{ name: 'P', body: '' }] }] },
      member('A', ['P'], { at: '2026-09-18T12:00', week: '2026-09-25' }),
    ],
  });
  const m = C.buildModel(s);
  assert.deepEqual(m.members.map((x) => x.name), ['B', 'A']);
  assert.deepEqual(m.members[0].lateWeeks, ['2026-09-25']);
  assert.deepEqual(m.members[1].lateWeeks, []);
});

test('buildModel: same earliest submit keeps input order', () => {
  const m = C.buildModel(state({ members: [member('Z', ['P']), member('A', ['P'])] }));
  assert.deepEqual(m.members.map((x) => x.name), ['Z', 'A']);
});

test('buildModel: warnings', () => {
  const s = state({
    members: [member('김민수', ['P'], { week: '2026-09-23' })],
    vacations: [{ name: ' 김민수 ', date: '2026-09-10' }, { name: '없는사람', date: '2026-09-11' }],
  });
  const m = C.buildModel(s);
  assert.deepEqual(m.blockers, []);
  assert.ok(m.warnings.some((w) => w.includes('1명') || w.includes('5명')));
  assert.ok(m.warnings.some((w) => w.includes('금요일')));
  assert.ok(m.warnings.some((w) => w.includes('없는사람')));
  assert.ok(m.warnings.some((w) => w.includes("'P'") && w.includes('요약')));
  assert.equal(m.members[0].workdays, 19); // 22 - 공휴일 2 - 휴가 1 (이름 공백 무시)
});

test('buildModel: other-month report is ignored with a warning; only-other-month member is blocked', () => {
  const s = state({
    members: [
      { name: 'A', reports: [
        { week: '2026-08-28', submittedAt: '2026-08-28T10:00', projects: [{ name: 'Old', body: '' }] },
        { week: '2026-09-25', submittedAt: '2026-09-25T10:00', projects: [{ name: 'P', body: '' }] }] },
      member('B', ['P'], { week: '2026-08-28', at: '2026-08-28T10:00' }),
    ],
  });
  const m = C.buildModel(s);
  assert.deepEqual(m.members[0].projects.map((p) => p.name), ['P']);
  assert.ok(m.warnings.some((w) => w.includes('무시')));
  assert.ok(m.blockers.some((b) => b.includes('B') && b.includes('2026-09')));
});

test('buildModel: blockers', () => {
  const has = (s, text) => C.buildModel(s).blockers.some((b) => b.includes(text));
  assert.ok(has(state(), '팀원이 없습니다'));
  assert.ok(has(state({ members: [member('  ', ['P'])] }), '이름'));
  assert.ok(has(state({ members: [member('A', ['P']), member(' a ', ['P'])] }), '중복'));
  assert.ok(has(state({ members: [member('A', [{ name: ' ', body: 'x' }])] }), '프로젝트'));
  assert.ok(has(state({ members: [member('A', [])] }), '프로젝트'));
  assert.ok(has(state({ holidays: [{ date: '2026-02-30', name: '' }], members: [member('A', ['P'])] }), '공휴일'));
  assert.ok(has(state({ vacations: [{ name: 'A', date: '' }], members: [member('A', ['P'])] }), '휴가'));
  assert.ok(has(state({ vacations: [{ name: '', date: '2026-09-01' }], members: [member('A', ['P'])] }), '휴가'));
  assert.ok(has(state({ members: [member('A', ['P'], { week: '2026-99-01' })] }), '주차'));
  assert.ok(has(state({ members: [member('A', ['P'], { at: '' })] }), '제출 시각'));
  assert.ok(has(state({ deadline: '25:00', members: [member('A', ['P'])] }), '마감 시각'));
  assert.ok(has(state({ month: 13, members: [member('A', ['P'])] }), '연·월'));
  assert.deepEqual(C.buildModel(state({ month: 13 })).members, []);
});

test('buildModel: zero team hours does not divide by zero', () => {
  const all = Array.from({ length: 30 }, (_, i) => ({ date: `2026-09-${String(i + 1).padStart(2, '0')}`, name: 'x' }));
  const m = C.buildModel(state({ holidays: all, members: [member('A', ['P', 'Q'])] }));
  assert.equal(m.members[0].baseHours, 0);
  assert.equal(m.members[0].teamSharePct, 0);
  assert.equal(m.members[0].projectSharePct, 0);
  assert.equal(m.members[0].perProjectHours, 0);
});

test('buildModel: filled summary produces no empty-summary warning', () => {
  const s = state({
    members: [member('A', ['P'])],
    summaries: { team: '', projects: { 'a|p': '됨' } },
  });
  assert.ok(!C.buildModel(s).warnings.some((w) => w.includes('요약이 비어')));
});
```

- [ ] **Step 2: 실패 확인**

Run: `node --test "web-tests/*.test.js"`
Expected: FAIL — `C.buildModel is not a function` (기존 9개는 통과)

- [ ] **Step 3: 구현**

`web/core.js`의 `// ---- exports ----` 바로 위에 추가:

```js
  function buildModel(state) {
    const warnings = [];
    const blockers = [];
    const teamSummary = (state.summaries && state.summaries.team) || '';
    const year = Number(state.year);
    const month = Number(state.month);
    const monthOk = Number.isInteger(year) && year >= 2000 && year <= 2100 && Number.isInteger(month) && month >= 1 && month <= 12;
    if (!monthOk) blockers.push('대상 연·월이 올바르지 않습니다');
    if (!DEADLINE_RE.test(state.deadline || '')) blockers.push('마감 시각이 올바르지 않습니다 (HH:MM)');
    if (!monthOk) return { year, month, teamSummary, members: [], warnings, blockers };
    const ym = `${year}-${pad(month)}`;

    const holidays = new Set();
    (state.holidays || []).forEach((h, i) => {
      if (isValidDate(h.date)) holidays.add(h.date);
      else blockers.push(`공휴일 ${i + 1}행: 날짜가 올바르지 않습니다 (${h.date || '비어 있음'})`);
    });

    const vacations = new Map();
    const vacationNames = new Map();
    (state.vacations || []).forEach((v, i) => {
      const name = cleanName(v.name);
      if (!name) { blockers.push(`휴가 ${i + 1}행: 이름이 비어 있습니다`); return; }
      if (!isValidDate(v.date)) { blockers.push(`휴가 ${i + 1}행: 날짜가 올바르지 않습니다 (${v.date || '비어 있음'})`); return; }
      const k = normName(name);
      if (!vacations.has(k)) { vacations.set(k, new Set()); vacationNames.set(k, name); }
      vacations.get(k).add(v.date);
    });

    const inputMembers = state.members || [];
    if (!inputMembers.length) blockers.push('팀원이 없습니다');
    const seen = new Set();
    const list = [];
    inputMembers.forEach((m, idx) => {
      const name = cleanName(m.name);
      if (!name) { blockers.push(`${idx + 1}번째 팀원의 이름이 비어 있습니다`); return; }
      const key = normName(name);
      if (seen.has(key)) { blockers.push(`팀원 이름이 중복됩니다: ${name}`); return; }
      seen.add(key);

      const inMonth = [];
      let ignored = 0;
      (m.reports || []).forEach((r) => {
        if (!isValidDate(r.week)) { blockers.push(`${name}: 주차 날짜가 올바르지 않습니다 (${r.week || '비어 있음'})`); return; }
        if (!isValidSubmit(r.submittedAt)) { blockers.push(`${name} ${r.week}: 제출 시각이 비어 있거나 올바르지 않습니다`); return; }
        if (Number(r.week.slice(0, 4)) !== year || Number(r.week.slice(5, 7)) !== month) { ignored++; return; }
        if (weekdayOf(r.week) !== 5) warnings.push(`${name}: 주차 날짜 ${r.week}가 금요일이 아닙니다 (마감 판정이 틀릴 수 있음)`);
        inMonth.push(r);
      });
      if (ignored) warnings.push(`${name}: 다른 달 보고 ${ignored}건은 무시되었습니다`);
      if (!inMonth.length) { blockers.push(`${name}: ${ym}에 해당하는 보고가 없습니다`); return; }

      inMonth.sort((a, b) => a.week.localeCompare(b.week) || a.submittedAt.localeCompare(b.submittedAt));
      const projects = new Map();
      inMonth.forEach((r) => {
        (r.projects || []).forEach((p) => {
          const pname = cleanName(p.name);
          if (!pname) { blockers.push(`${name} ${r.week}: 프로젝트 이름이 비어 있습니다`); return; }
          const pkey = normName(pname);
          if (!projects.has(pkey)) projects.set(pkey, { name: pname, bodies: [], summary: '' });
          const entry = projects.get(pkey);
          if (entry.name !== pname && !entry.merged) {
            entry.merged = true;
            warnings.push(`${name}: 프로젝트 '${entry.name}'와 '${pname}'는 같은 프로젝트로 합쳤습니다`);
          }
          const body = String(p.body == null ? '' : p.body).trim();
          if (body) entry.bodies.push(body);
        });
      });
      if (!projects.size) { blockers.push(`${name}: 프로젝트가 없습니다`); return; }

      const summaries = (state.summaries && state.summaries.projects) || {};
      const plist = [...projects.values()].map((p) => ({
        name: p.name, bodies: p.bodies, summary: String(summaries[summaryKey(name, p.name)] || '').trim(),
      }));
      list.push({
        name, key, idx, projects: plist,
        lateWeeks: inMonth.filter((r) => isLate(r.week, r.submittedAt, state.deadline)).map((r) => r.week),
        earliest: inMonth.map((r) => r.submittedAt.slice(0, 16)).sort()[0],
      });
    });

    list.sort((a, b) => a.earliest.localeCompare(b.earliest) || a.idx - b.idx);
    list.forEach((mm) => {
      mm.workdays = workdays(year, month, holidays, vacations.get(mm.key) || new Set());
      mm.baseHours = mm.workdays * HOURS_PER_DAY;
    });
    const total = list.reduce((s, mm) => s + mm.baseHours, 0);
    list.forEach((mm) => {
      const n = mm.projects.length;
      mm.perProjectHours = n ? mm.baseHours / n : 0;
      mm.personalPct = n ? 100 / n : 0;
      mm.teamSharePct = total ? (mm.baseHours / total) * 100 : 0;
      mm.projectSharePct = total ? (mm.perProjectHours / total) * 100 : 0;
      delete mm.key; delete mm.idx;
    });

    if (inputMembers.length && inputMembers.length !== TEAM_SIZE) {
      warnings.push(`팀원이 ${inputMembers.length}명입니다 (기획서 기준 ${TEAM_SIZE}명). 팀 합계와 비중이 달라집니다`);
    }
    vacationNames.forEach((display, k) => {
      if (!seen.has(k)) warnings.push(`휴가 명단의 '${display}'은(는) 팀원에 없는 이름입니다 (오타 확인)`);
    });
    list.forEach((mm) => mm.projects.forEach((p) => {
      if (!p.summary) warnings.push(`${mm.name}: '${p.name}'의 요약이 비어 있습니다`);
    }));

    return { year, month, teamSummary, members: list, warnings, blockers };
  }
```

exports 객체에 `buildModel`을 추가한다: `... isValidSubmit, isLate, lastFriday, fmtHours, fmtPct, buildModel,`

- [ ] **Step 4: 통과 확인**

Run: `node --test "web-tests/*.test.js"`
Expected: 18 tests pass (Task 1의 9 + 9), 0 fail

---

### Task 3: core.js — 상태 만들기·검증·파싱

**Files:**
- Modify: `web/core.js`
- Test: `web-tests/core.test.js` (끝에 추가)

**Interfaces:**
- Consumes: 없음
- Produces:
  - `STATE_VERSION = 1`
  - `emptyState(year: number, month: number): State`
  - `validateState(obj: any): {ok: true, state: State} | {ok: false, error: string}`
  - `parseState(text: string): {ok: true, state: State} | {ok: false, error: string}` — JSON 파싱 실패도 `ok:false`

- [ ] **Step 1: 실패하는 테스트 추가**

```js
test('emptyState is valid and matches buildModel input shape', () => {
  const s = C.emptyState(2026, 9);
  assert.equal(s.v, C.STATE_VERSION);
  assert.equal(s.deadline, '17:00');
  assert.ok(C.validateState(s).ok);
  assert.ok(C.buildModel(s).blockers.some((b) => b.includes('팀원이 없습니다')));
});

test('validateState rejects bad shapes and versions', () => {
  const bad = (o) => C.validateState(o);
  assert.equal(bad(null).ok, false);
  assert.equal(bad([]).ok, false);
  assert.equal(bad('x').ok, false);
  assert.match(bad({ ...C.emptyState(2026, 9), v: 2 }).error, /버전/);
  assert.equal(bad({ ...C.emptyState(2026, 9), members: {} }).ok, false);
  assert.equal(bad({ ...C.emptyState(2026, 9), holidays: null }).ok, false);
  assert.equal(bad({ ...C.emptyState(2026, 9), summaries: null }).ok, false);
  assert.equal(bad({ ...C.emptyState(2026, 9), members: [{ name: 'a' }] }).ok, false);
  assert.equal(bad({ ...C.emptyState(2026, 9), members: [{ name: 'a', reports: [{ week: 'x' }] }] }).ok, false);
});

test('parseState: round trip, corrupt JSON, wrong version', () => {
  const s = C.emptyState(2026, 9);
  s.members.push({ name: '김민수', reports: [{ week: '2026-09-25', submittedAt: '2026-09-25T10:00', projects: [{ name: 'P', body: '내용' }] }] });
  const r = C.parseState(JSON.stringify(s));
  assert.ok(r.ok);
  assert.deepEqual(r.state, s);
  const broken = C.parseState('{not json');
  assert.equal(broken.ok, false);
  assert.match(broken.error, /JSON/);
  assert.equal(C.parseState(JSON.stringify({ v: 99 })).ok, false);
  assert.equal(C.parseState('').ok, false);
});
```

- [ ] **Step 2: 실패 확인**

Run: `node --test "web-tests/*.test.js"`
Expected: FAIL — `C.emptyState is not a function`

- [ ] **Step 3: 구현**

`web/core.js`의 `// ---- exports ----` 위에 추가:

```js
  const STATE_VERSION = 1;

  function emptyState(year, month) {
    return {
      v: STATE_VERSION, year, month, deadline: '17:00',
      holidays: [], vacations: [], members: [],
      summaries: { team: '', projects: {} },
    };
  }

  function validateState(o) {
    const fail = (error) => ({ ok: false, error });
    if (!o || typeof o !== 'object' || Array.isArray(o)) return fail('데이터 형식이 올바르지 않습니다');
    if (o.v !== STATE_VERSION) return fail(`지원하지 않는 데이터 버전입니다 (v${o.v})`);
    for (const k of ['holidays', 'vacations', 'members']) {
      if (!Array.isArray(o[k])) return fail(`'${k}' 항목이 목록이 아닙니다`);
    }
    const s = o.summaries;
    if (!s || typeof s !== 'object' || !s.projects || typeof s.projects !== 'object') return fail("'summaries' 항목이 올바르지 않습니다");
    for (const m of o.members) {
      if (!m || !Array.isArray(m.reports)) return fail('팀원 데이터에 보고 목록이 없습니다');
      for (const r of m.reports) {
        if (!r || !Array.isArray(r.projects)) return fail('보고 데이터에 프로젝트 목록이 없습니다');
      }
    }
    return { ok: true, state: o };
  }

  function parseState(text) {
    let o;
    try { o = JSON.parse(text); } catch (e) { return { ok: false, error: '데이터를 읽을 수 없습니다 (JSON 오류)' }; }
    return validateState(o);
  }
```

exports에 `STATE_VERSION, emptyState, validateState, parseState` 추가.

- [ ] **Step 4: 통과 확인**

Run: `node --test "web-tests/*.test.js"`
Expected: 21 tests pass, 0 fail

---

### Task 4: deck.js — 슬라이드 계획과 PptxGenJS 생성

**Files:**
- Create: `web/deck.js`
- Test: `web-tests/deck.test.js`

**Interfaces:**
- Consumes: `core.buildModel`, `fmtHours`, `fmtPct` (Task 1–2), `Model`/`MemberModel` 구조 (Task 2)
- Produces (`module.exports` / `window.TRDeck`):
  - `MAX_ROWS = 6`
  - `planSlides(model): Slide[]` — `Slide = {kind: 'team'|'member', title: string, header: string[], rows: string[][], text: string}`
  - `buildPptx(PptxGenJS: new () => Pptx, model): Pptx` — 슬라이드를 채운 인스턴스를 반환(저장은 호출자가 `writeFile`/`write`)

- [ ] **Step 1: 실패하는 테스트**

`web-tests/deck.test.js`:

```js
const test = require('node:test');
const assert = require('node:assert/strict');
const C = require('../web/core.js');
const D = require('../web/deck.js');

const names = (n, p) => Array.from({ length: n }, (_, i) => ({ name: `${p}${i + 1}`, body: '내용' }));
function model(members, over = {}) {
  return C.buildModel(Object.assign({
    v: 1, year: 2026, month: 9, deadline: '17:00',
    holidays: [{ date: '2026-09-24', name: '' }, { date: '2026-09-25', name: '' }],
    vacations: [], members, summaries: { team: '팀 요약', projects: {} },
  }, over));
}
const mem = (name, projects, at = '2026-09-25T10:00') => ({
  name, reports: [{ week: '2026-09-25', submittedAt: at, projects }],
});

test('planSlides: team slide + one slide per member, numbers formatted', () => {
  const m = model([mem('A', names(4, 'p')), mem('B', names(1, 'q'), '2026-09-25T11:00')]);
  const plan = D.planSlides(m);
  assert.equal(plan.length, 3);
  assert.equal(plan[0].title, '2026년 9월 팀 주간 보고 취합');
  assert.deepEqual(plan[0].rows[0], ['A', '4', '160.0h', '50.0%']);
  assert.equal(plan[0].text, '팀 요약');
  assert.equal(plan[1].title, 'A — 2026년 9월');
  assert.deepEqual(plan[1].rows[0].slice(2), ['40.0h', '25.0%', '12.5%']);
  assert.match(plan[1].text, /근무일 20일 · 월 기준 시간 160\.0h · 팀 대비 50\.0%/);
  assert.equal(plan[2].title, 'B — 2026년 9월');
});

test('planSlides: 13 projects split into 6 + 6 + 1 with (계속)', () => {
  const plan = D.planSlides(model([mem('A', names(13, 'p'))]));
  const member = plan.filter((s) => s.kind === 'member');
  assert.deepEqual(member.map((s) => s.rows.length), [6, 6, 1]);
  assert.equal(member[0].title, 'A — 2026년 9월');
  assert.equal(member[1].title, 'A — 2026년 9월 (계속)');
  assert.equal(member[2].title, 'A — 2026년 9월 (계속)');
});

test('planSlides: empty summary falls back to marked, truncated original', () => {
  const m = model([mem('A', [{ name: 'P', body: '가'.repeat(500) }, { name: 'Q', body: '- 짧음' }])], {
    summaries: { team: '', projects: { 'a|q': '' } },
  });
  const plan = D.planSlides(m);
  const long = plan[1].rows[0][1];
  assert.ok(long.startsWith('(요약 없음) '));
  assert.ok(long.endsWith('…'));
  assert.ok(long.length <= '(요약 없음) '.length + 201);
  assert.equal(plan[1].rows[1][1], '(요약 없음) - 짧음');
  assert.equal(plan[0].text, '(요약 없음)');
});

test('planSlides: filled summary is used; late weeks noted on member slide', () => {
  const m = model([{
    name: 'A',
    reports: [{ week: '2026-09-25', submittedAt: '2026-09-28T09:00', projects: [{ name: 'P', body: 'x' }] }],
  }], { summaries: { team: '', projects: { 'a|p': '요약된 문장' } } });
  const plan = D.planSlides(m);
  assert.equal(plan[1].rows[0][1], '요약된 문장');
  assert.match(plan[1].text, /마감 후 제출: 2026-09-25/);
});

test('planSlides: member name with HTML characters stays literal text', () => {
  const plan = D.planSlides(model([mem('<b>A&B</b>', names(1, 'p'))]));
  assert.equal(plan[1].title, '<b>A&B</b> — 2026년 9월');
});

class FakePptx {
  constructor() { this.slides = []; }
  addSlide() {
    const s = { texts: [], tables: [] };
    s.addText = (t, o) => s.texts.push([t, o]);
    s.addTable = (r, o) => s.tables.push([r, o]);
    this.slides.push(s);
    return s;
  }
}

test('buildPptx: one slide per plan entry, wide layout, 맑은 고딕, table + text', () => {
  const m = model([mem('A', names(13, 'p'))]);
  const pptx = D.buildPptx(FakePptx, m);
  assert.equal(pptx.layout, 'LAYOUT_WIDE');
  assert.equal(pptx.slides.length, 4);
  const first = pptx.slides[0];
  assert.equal(first.tables.length, 1);
  assert.equal(first.tables[0][1].fontFace, '맑은 고딕');
  assert.equal(first.texts[0][0], '2026년 9월 팀 주간 보고 취합');
  const rows = pptx.slides[1].tables[0][0];
  assert.equal(rows.length, 7); // 머리글 + 6
  assert.equal(rows[0][0].text, '프로젝트');
});
```

- [ ] **Step 2: 실패 확인**

Run: `node --test "web-tests/deck.test.js"`
Expected: FAIL — `Cannot find module '../web/deck.js'`

- [ ] **Step 3: 구현**

`web/deck.js`:

```js
(function (root) {
  'use strict';
  const core = (typeof module === 'object' && module.exports) ? require('./core.js') : root.TR;

  const MAX_ROWS = 6;
  const FALLBACK_LEN = 200;
  const FONT = '맑은 고딕';
  const TEAM_WIDTHS = [3.5, 2.6, 3.1, 3.1];
  const MEMBER_WIDTHS = [2.4, 5.6, 1.1, 1.6, 1.6];

  function summaryText(p) {
    if (p.summary) return p.summary;
    let text = p.bodies.join('\n');
    if (text.length > FALLBACK_LEN) text = text.slice(0, FALLBACK_LEN) + '…';
    return '(요약 없음) ' + text;
  }

  function planSlides(model) {
    const { year, month } = model;
    const slides = [{
      kind: 'team',
      title: `${year}년 ${month}월 팀 주간 보고 취합`,
      header: ['팀원', '프로젝트 수', '월 기준 시간', '팀 대비 비중'],
      rows: model.members.map((m) => [m.name, String(m.projects.length), core.fmtHours(m.baseHours), core.fmtPct(m.teamSharePct)]),
      text: (model.teamSummary || '').trim() || '(요약 없음)',
    }];
    model.members.forEach((m) => {
      let memo = `근무일 ${m.workdays}일 · 월 기준 시간 ${core.fmtHours(m.baseHours)} · 팀 대비 ${core.fmtPct(m.teamSharePct)}`;
      if (m.lateWeeks.length) memo += `\n마감 후 제출: ${m.lateWeeks.join(', ')}`;
      for (let i = 0; i < m.projects.length; i += MAX_ROWS) {
        slides.push({
          kind: 'member',
          title: `${m.name} — ${year}년 ${month}월` + (i > 0 ? ' (계속)' : ''),
          header: ['프로젝트', '요약', '시간', '개인 기준 %', '팀 대비 %'],
          rows: m.projects.slice(i, i + MAX_ROWS).map((p) => [
            p.name, summaryText(p), core.fmtHours(m.perProjectHours), core.fmtPct(m.personalPct), core.fmtPct(m.projectSharePct),
          ]),
          text: memo,
        });
      }
    });
    return slides;
  }

  function buildPptx(PptxGenJS, model) {
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE';
    planSlides(model).forEach((s) => {
      const slide = pptx.addSlide();
      slide.addText(s.title, { x: 0.5, y: 0.3, w: 12.3, h: 0.8, fontSize: 28, bold: true, fontFace: FONT });
      const head = s.header.map((t) => ({ text: t, options: { bold: true, fill: { color: 'D9E2F3' } } }));
      const body = s.rows.map((r) => r.map((t) => ({ text: String(t) })));
      const team = s.kind === 'team';
      slide.addTable([head, ...body], {
        x: 0.5, y: 1.4, w: 12.3, colW: team ? TEAM_WIDTHS : MEMBER_WIDTHS,
        fontSize: team ? 14 : 12, fontFace: FONT, valign: 'middle',
        border: { type: 'solid', pt: 0.5, color: '999999' },
      });
      const textY = team ? 1.4 + 0.45 * (s.rows.length + 1) + 0.4 : 6.2;
      slide.addText(s.text, { x: 0.5, y: textY, w: 12.3, h: 1.0, fontSize: team ? 14 : 11, fontFace: FONT, valign: 'top' });
    });
    return pptx;
  }

  const api = { MAX_ROWS, planSlides, buildPptx };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.TRDeck = api;
})(typeof self !== 'undefined' ? self : this);
```

- [ ] **Step 4: 통과 확인**

Run: `node --test "web-tests/*.test.js"`
Expected: 27 tests pass (21 + 6), 0 fail

---

### Task 5: PptxGenJS 파일 확보와 실제 .pptx 생성 스모크 테스트

**Files:**
- Create: `web/vendor/pptxgen.bundle.js` (npm 패키지 `pptxgenjs@4.0.1`의 `dist/pptxgen.bundle.js` 복사)
- Test: `web-tests/pptx-smoke.test.js`

**Interfaces:**
- Consumes: `TRDeck.buildPptx`, `core.buildModel` (Task 2, 4)
- Produces: 브라우저 전역 `window.PptxGenJS`(vendor 스크립트가 제공), Node에서 `require('../web/vendor/pptxgen.bundle.js')`로 생성자 사용

- [ ] **Step 1: (사람 허락 필요) 라이브러리 받기**

파일을 받는 단계이므로 실행 직전에 사용자에게 파일명·출처·크기를 알리고 명시적 허락을 받는다. 허락 후:

```bash
mkdir -p web/vendor
cd "$SCRATCHPAD" && npm pack pptxgenjs@4.0.1 && tar -xzf pptxgenjs-4.0.1.tgz package/dist/pptxgen.bundle.js
cp "$SCRATCHPAD/package/dist/pptxgen.bundle.js" "C:/Users/KSA/Desktop/실습/web/vendor/pptxgen.bundle.js"
```

(`$SCRATCHPAD` = 세션 scratchpad 디렉터리.) 복사 후 파일 크기와 첫 줄의 라이선스 헤더(MIT)를 확인해 기록한다. 허락이 없으면 이 Task와 Task 7의 PPT 생성 검증은 보류하고 나머지를 진행한다.

- [ ] **Step 2: 실패하는(또는 첫 실행하는) 스모크 테스트 작성**

`web-tests/pptx-smoke.test.js`:

```js
const test = require('node:test');
const assert = require('node:assert/strict');
const C = require('../web/core.js');
const D = require('../web/deck.js');
const PptxGenJS = require('../web/vendor/pptxgen.bundle.js');

const mem = (name, n, at) => ({
  name,
  reports: [{ week: '2026-09-25', submittedAt: at, projects: Array.from({ length: n }, (_, i) => ({ name: `P${i + 1}`, body: '내용' })) }],
});

test('real PptxGenJS produces a valid pptx zip with the planned slides', async () => {
  const model = C.buildModel({
    v: 1, year: 2026, month: 9, deadline: '17:00',
    holidays: [{ date: '2026-09-24', name: '' }, { date: '2026-09-25', name: '' }],
    vacations: [], members: [mem('김민수', 13, '2026-09-25T10:00'), mem('<b>이서연</b>', 2, '2026-09-25T11:00')],
    summaries: { team: '팀 요약 문장', projects: {} },
  });
  const pptx = D.buildPptx(PptxGenJS, model);
  const buf = await pptx.write({ outputType: 'nodebuffer' });
  assert.equal(buf.slice(0, 2).toString('latin1'), 'PK');
  assert.ok(buf.length > 10000);
  const names = new Set(buf.toString('latin1').match(/ppt\/slides\/slide\d+\.xml/g));
  assert.equal(names.size, D.planSlides(model).length); // 1 + 3 + 1
});
```

- [ ] **Step 3: 실행해 확인**

Run: `node --test "web-tests/pptx-smoke.test.js"`
Expected: PASS. 실패하면(예: `write` 시그니처, UMD `require` 방식) 원인을 확인해 최소 변경으로 고치고 Ruling으로 기록한다.

- [ ] **Step 4: python-pptx로 실제 파일 열어 보기 (1회, 수동)**

```bash
node -e "const C=require('./web/core.js'),D=require('./web/deck.js'),P=require('./web/vendor/pptxgen.bundle.js');const m=C.buildModel({v:1,year:2026,month:9,deadline:'17:00',holidays:[],vacations:[],members:[{name:'김민수',reports:[{week:'2026-09-25',submittedAt:'2026-09-25T10:00',projects:[{name:'P',body:'내용'}]}]}],summaries:{team:'팀 요약',projects:{}}});D.buildPptx(P,m).write({outputType:'nodebuffer'}).then(b=>require('fs').writeFileSync(process.env.SCRATCHPAD+'/smoke.pptx',b))"
python -c "from pptx import Presentation;import os;p=Presentation(os.environ['SCRATCHPAD']+'/smoke.pptx');print(len(p.slides),[s.shapes.title.text if s.shapes.title else [sh.text_frame.text for sh in s.shapes if sh.has_text_frame][0] for s in p.slides])"
```

Expected: `2 [...]`에 "2026년 9월 팀 주간 보고 취합"과 "김민수 — 2026년 9월"이 보인다.

- [ ] **Step 5: 전체 테스트**

Run: `node --test "web-tests/*.test.js"`
Expected: 28 tests pass, 0 fail

---

### Task 6: 화면 — index.html, style.css, app.js

**Files:**
- Create: `web/index.html`, `web/style.css`, `web/app.js`

**Interfaces:**
- Consumes: `window.TR`(`buildModel`, `emptyState`, `parseState`, `summaryKey`, `lastFriday`, `cleanName`, `fmtHours`, `fmtPct`), `window.TRDeck.buildPptx`, `window.PptxGenJS`
- Produces: 사용자 화면. 검증은 Task 7(브라우저). 여기서는 코드 작성 후 문법 검사(`node --check web/app.js`)와 전체 테스트 회귀만 확인한다.

동작 규칙:
- 모든 사용자 입력은 `h()` 헬퍼로 텍스트 노드/속성으로만 DOM에 넣는다(`innerHTML` 사용 금지).
- 입력 이벤트는 상태만 바꾸고 `changed()`(저장 예약 + 확인 패널 + 월 요약 영역 갱신)를 호출한다. 월 요약 텍스트칸의 입력은 포커스를 잃지 않도록 `save() + refreshPanel()`만 호출한다. 추가/삭제 버튼은 해당 영역을 다시 그린다.
- 저장: localStorage `teamReport.state`, 300ms 디바운스, 모든 storage 접근은 try/catch.

- [ ] **Step 1: `web/index.html`**

```html
<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>팀 주간 보고 취합</title>
  <link rel="stylesheet" href="style.css">
</head>
<body>
  <header>
    <h1>팀 주간 보고 취합</h1>
    <p class="hint">팀원별 주간 보고를 입력하면 월 기여도를 계산해 PPT로 만듭니다. 입력한 내용은 이 브라우저에만 저장됩니다.</p>
    <div id="notice" class="notice" hidden></div>
  </header>
  <main>
    <section><h2>1. 기본 설정</h2><div id="settings" class="row"></div></section>
    <section>
      <h2>2. 달력</h2>
      <h3>공휴일</h3><div id="holidays"></div>
      <h3>휴가 (하루 단위)</h3><div id="vacations"></div>
      <datalist id="member-names"></datalist>
    </section>
    <section><h2>3. 팀원 보고</h2><div id="members"></div><button type="button" id="add-member">+ 팀원 추가</button></section>
    <section>
      <h2>4. 요약 문장</h2>
      <label class="block">팀 전체 요약<textarea id="team-summary" rows="3"></textarea></label>
      <div id="summaries"></div>
    </section>
    <section>
      <h2>5. 확인</h2>
      <div id="panel"></div>
      <div class="actions">
        <button type="button" id="generate" class="primary">PPT 만들기</button>
        <button type="button" id="backup">백업 저장</button>
        <button type="button" id="restore">백업 불러오기</button>
        <input type="file" id="restore-file" accept=".json,application/json" hidden>
      </div>
    </section>
  </main>
  <script src="vendor/pptxgen.bundle.js"></script>
  <script src="core.js"></script>
  <script src="deck.js"></script>
  <script src="app.js"></script>
</body>
</html>
```

- [ ] **Step 2: `web/style.css`**

```css
:root { color-scheme: light dark; --line: #8886; --accent: #2b579a; --warn: #b26a00; --err: #c62828; }
* { box-sizing: border-box; }
body { font-family: "맑은 고딕", system-ui, sans-serif; margin: 0 auto; max-width: 1100px; padding: 16px; line-height: 1.5; }
h1 { margin: 0 0 4px; } h2 { margin: 0 0 8px; font-size: 1.15rem; } h3 { margin: 12px 0 6px; font-size: 1rem; }
section { border: 1px solid var(--line); border-radius: 8px; padding: 12px 16px; margin: 12px 0; }
.hint { color: #888; margin: 0; }
.notice { margin-top: 8px; padding: 8px 12px; border-left: 4px solid var(--warn); background: #ff980022; }
.row { display: flex; flex-wrap: wrap; gap: 12px; align-items: end; }
label { display: flex; flex-direction: column; font-size: .85rem; gap: 2px; }
label.block { width: 100%; }
input, textarea, select, button { font: inherit; padding: 4px 6px; }
textarea { width: 100%; resize: vertical; }
button { cursor: pointer; }
button.primary { background: var(--accent); color: #fff; border: 0; border-radius: 6px; padding: 8px 16px; }
button:disabled { opacity: .45; cursor: not-allowed; }
.card { border: 1px solid var(--line); border-radius: 8px; padding: 10px 12px; margin: 10px 0; }
.report { border-top: 1px dashed var(--line); margin-top: 8px; padding-top: 8px; }
.project { display: grid; grid-template-columns: 200px 1fr auto; gap: 8px; margin: 6px 0; align-items: start; }
.item { display: flex; gap: 8px; margin: 4px 0; flex-wrap: wrap; }
.actions { display: flex; gap: 8px; flex-wrap: wrap; margin-top: 12px; }
table { border-collapse: collapse; width: 100%; font-size: .9rem; }
th, td { border: 1px solid var(--line); padding: 4px 8px; text-align: left; vertical-align: top; }
th { background: #d9e2f344; }
.late { color: var(--err); font-weight: 600; }
ul.warn li { color: var(--warn); } ul.err li { color: var(--err); font-weight: 600; }
@media (max-width: 700px) { .project { grid-template-columns: 1fr; } }
```

- [ ] **Step 3: `web/app.js`**

```js
(function () {
  'use strict';
  const TR = window.TR;
  const TRDeck = window.TRDeck;
  const KEY = 'teamReport.state';
  const $ = (id) => document.getElementById(id);

  // 사용자 입력은 항상 텍스트로만 넣는다 (innerHTML 금지)
  function h(tag, props, ...kids) {
    const el = document.createElement(tag);
    let value;
    for (const [k, v] of Object.entries(props || {})) {
      if (k === 'class') el.className = v;
      else if (k === 'value') value = v;
      else if (k.startsWith('on')) el.addEventListener(k.slice(2), v);
      else if (v === true) el.setAttribute(k, '');
      else if (v !== false && v != null) el.setAttribute(k, v);
    }
    for (const kid of kids.flat()) el.append(kid instanceof Node ? kid : document.createTextNode(String(kid)));
    if (value !== undefined) el.value = value;
    return el;
  }
  const field = (label, input) => h('label', {}, label, input);
  const pad = (n) => String(n).padStart(2, '0');
  function nowLocal() {
    const d = new Date();
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
  }

  // ---- 상태·저장 ----
  let state;
  let model;
  let timer;

  function load() {
    try {
      const raw = localStorage.getItem(KEY);
      if (raw) {
        const r = TR.parseState(raw);
        if (r.ok) return r.state;
        showNotice(`저장된 데이터를 불러오지 못해 빈 화면으로 시작합니다: ${r.error}`);
      }
    } catch (e) {
      showNotice('이 브라우저에서는 자동 저장을 사용할 수 없습니다. 백업 저장 버튼으로 직접 보관하세요.');
    }
    const now = new Date();
    return TR.emptyState(now.getFullYear(), now.getMonth() + 1);
  }
  function save() {
    clearTimeout(timer);
    timer = setTimeout(() => { try { localStorage.setItem(KEY, JSON.stringify(state)); } catch (e) { /* 저장 불가 */ } }, 300);
  }
  function showNotice(text) { const n = $('notice'); n.textContent = text; n.hidden = !text; }
  function changed() { save(); refreshPanel(); renderSummaries(); updateNames(); }
  const bind = (obj, key, num) => (e) => { obj[key] = num ? Number(e.target.value) : e.target.value; changed(); };

  // ---- 1. 기본 설정 ----
  function renderSettings() {
    $('settings').replaceChildren(
      field('연', h('input', { type: 'number', min: 2000, max: 2100, value: state.year, oninput: bind(state, 'year', true) })),
      field('월', h('input', { type: 'number', min: 1, max: 12, value: state.month, oninput: bind(state, 'month', true) })),
      field('마감 시각', h('input', { type: 'time', value: state.deadline, oninput: bind(state, 'deadline') })),
    );
  }

  // ---- 2. 달력 ----
  function renderCalendar() {
    $('holidays').replaceChildren(
      ...state.holidays.map((row, i) => h('div', { class: 'item' },
        h('input', { type: 'date', value: row.date, 'aria-label': '공휴일 날짜', oninput: bind(row, 'date') }),
        h('input', { type: 'text', value: row.name, placeholder: '이름(선택)', oninput: bind(row, 'name') }),
        h('button', { type: 'button', onclick: () => { state.holidays.splice(i, 1); renderCalendar(); changed(); } }, '삭제'))),
      h('button', { type: 'button', onclick: () => { state.holidays.push({ date: '', name: '' }); renderCalendar(); changed(); } }, '+ 공휴일 추가'),
    );
    $('vacations').replaceChildren(
      ...state.vacations.map((row, i) => h('div', { class: 'item' },
        h('input', { type: 'text', list: 'member-names', value: row.name, placeholder: '팀원 이름', 'aria-label': '휴가 팀원', oninput: bind(row, 'name') }),
        h('input', { type: 'date', value: row.date, 'aria-label': '휴가 날짜', oninput: bind(row, 'date') }),
        h('button', { type: 'button', onclick: () => { state.vacations.splice(i, 1); renderCalendar(); changed(); } }, '삭제'))),
      h('button', { type: 'button', onclick: () => { state.vacations.push({ name: '', date: '' }); renderCalendar(); changed(); } }, '+ 휴가 추가'),
    );
  }
  function updateNames() {
    $('member-names').replaceChildren(...state.members.filter((m) => m.name.trim()).map((m) => h('option', { value: m.name.trim() })));
  }

  // ---- 3. 팀원 ----
  function newReport() { return { week: TR.lastFriday(state.year, state.month), submittedAt: nowLocal(), projects: [{ name: '', body: '' }] }; }

  function renderMembers() {
    $('members').replaceChildren(...state.members.map((m, mi) => h('div', { class: 'card' },
      h('div', { class: 'row' },
        field('팀원 이름', h('input', { type: 'text', value: m.name, oninput: bind(m, 'name') })),
        h('button', { type: 'button', onclick: () => { if (confirm(`'${m.name || '이름 없음'}' 팀원을 삭제할까요?`)) { state.members.splice(mi, 1); renderMembers(); changed(); } } }, '팀원 삭제')),
      ...m.reports.map((r, ri) => h('div', { class: 'report' },
        h('div', { class: 'row' },
          field('주차(금요일)', h('input', { type: 'date', value: r.week, oninput: bind(r, 'week') })),
          field('제출 시각', h('input', { type: 'datetime-local', value: r.submittedAt, oninput: bind(r, 'submittedAt') })),
          h('button', { type: 'button', onclick: () => { m.reports.splice(ri, 1); renderMembers(); changed(); } }, '이 주차 삭제')),
        ...r.projects.map((p, pi) => h('div', { class: 'project' },
          h('input', { type: 'text', value: p.name, placeholder: '프로젝트명', 'aria-label': '프로젝트명', oninput: bind(p, 'name') }),
          h('textarea', { rows: 2, value: p.body, placeholder: '주간 내용', 'aria-label': '주간 내용', oninput: bind(p, 'body') }),
          h('button', { type: 'button', onclick: () => { r.projects.splice(pi, 1); renderMembers(); changed(); } }, '삭제'))),
        h('button', { type: 'button', onclick: () => { r.projects.push({ name: '', body: '' }); renderMembers(); changed(); } }, '+ 프로젝트 추가'))),
      h('button', { type: 'button', onclick: () => { m.reports.push(newReport()); renderMembers(); changed(); } }, '+ 주차 보고 추가'),
    )));
  }

  // ---- 4. 요약 ----
  function renderSummaries() {
    const box = $('summaries');
    box.replaceChildren(...model.members.map((mm) => h('div', { class: 'card' },
      h('h3', {}, mm.name),
      ...mm.projects.map((p) => {
        const key = TR.summaryKey(mm.name, p.name);
        return field(`${p.name}`, h('textarea', {
          rows: 2, value: state.summaries.projects[key] || '',
          placeholder: p.bodies.join(' / ').slice(0, 80) || '상위부서 보고용 요약 문장',
          oninput: (e) => { state.summaries.projects[key] = e.target.value; save(); refreshPanel(); },
        }));
      }))));
  }

  // ---- 5. 확인 패널 ----
  function refreshPanel() {
    model = TR.buildModel(state);
    const rows = model.members.map((m) => h('tr', {},
      h('td', {}, m.name),
      h('td', {}, `${m.projects.length}개: ${m.projects.map((p) => p.name).join(', ')}`),
      h('td', {}, `${m.workdays}일`),
      h('td', {}, TR.fmtHours(m.baseHours)),
      h('td', {}, TR.fmtPct(m.teamSharePct)),
      h('td', {}, `${TR.fmtHours(m.perProjectHours)} / ${TR.fmtPct(m.personalPct)} / ${TR.fmtPct(m.projectSharePct)}`),
      h('td', { class: m.lateWeeks.length ? 'late' : '' }, m.lateWeeks.length ? `지연 ${m.lateWeeks.join(', ')}` : '-')));
    const table = model.members.length
      ? h('table', {}, h('thead', {}, h('tr', {}, ...['팀원', '프로젝트', '근무일', '월 기준 시간', '팀 대비', '프로젝트당 시간 / 개인 기준 % / 프로젝트 팀 대비', '마감'].map((t) => h('th', {}, t)))), h('tbody', {}, rows))
      : h('p', { class: 'hint' }, '입력하면 여기에 계산 결과가 보입니다.');
    $('panel').replaceChildren(
      table,
      model.blockers.length ? h('div', {}, h('h3', {}, 'PPT를 만들 수 없는 이유'), h('ul', { class: 'err' }, ...model.blockers.map((b) => h('li', {}, b)))) : '',
      model.warnings.length ? h('div', {}, h('h3', {}, '경고 (확인 후 진행 가능)'), h('ul', { class: 'warn' }, ...model.warnings.map((w) => h('li', {}, w)))) : '',
    );
    $('generate').disabled = model.blockers.length > 0;
  }

  // ---- 동작 ----
  function download(name, blob) {
    const a = h('a', { href: URL.createObjectURL(blob), download: name });
    document.body.append(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  }
  const ymText = () => `${state.year}-${pad(state.month)}`;

  async function generate() {
    const m = TR.buildModel(state);
    if (m.blockers.length) return;
    if (m.warnings.length && !confirm(`경고 ${m.warnings.length}건이 있습니다:\n- ${m.warnings.join('\n- ')}\n\n그래도 PPT를 만들까요?`)) return;
    if (!window.PptxGenJS) { alert('PPT 라이브러리(vendor/pptxgen.bundle.js)를 불러오지 못했습니다.'); return; }
    try {
      await TRDeck.buildPptx(window.PptxGenJS, m).writeFile({ fileName: `팀_월간보고_${ymText()}.pptx` });
    } catch (e) { alert('PPT를 만들지 못했습니다: ' + e.message); }
  }

  function renderAll() {
    $('team-summary').value = state.summaries.team || '';
    renderSettings(); renderCalendar(); renderMembers(); refreshPanel(); renderSummaries(); updateNames();
  }

  function init() {
    state = load();
    $('team-summary').addEventListener('input', (e) => { state.summaries.team = e.target.value; save(); refreshPanel(); });
    $('add-member').addEventListener('click', () => { state.members.push({ name: '', reports: [newReport()] }); renderMembers(); changed(); });
    $('generate').addEventListener('click', generate);
    $('backup').addEventListener('click', () => download(`팀보고_백업_${ymText()}.json`, new Blob([JSON.stringify(state, null, 2)], { type: 'application/json' })));
    $('restore').addEventListener('click', () => $('restore-file').click());
    $('restore-file').addEventListener('change', async (e) => {
      const file = e.target.files[0];
      e.target.value = '';
      if (!file) return;
      const r = TR.parseState(await file.text());
      if (!r.ok) { alert(`백업 파일을 불러올 수 없습니다: ${r.error}`); return; }
      if (!confirm('현재 입력한 내용이 백업 내용으로 바뀝니다. 계속할까요?')) return;
      state = r.state; save(); renderAll();
    });
    renderAll();
  }
  init();
})();
```

- [ ] **Step 4: 문법 검사와 회귀**

Run: `node --check web/app.js && node --test "web-tests/*.test.js"`
Expected: 오류 없음, 28 tests pass

---

### Task 7: 브라우저 종단 검증

**Files:**
- 코드 변경 없음(문제가 나오면 해당 파일 수정 + 재현 테스트 추가)

**Interfaces:**
- Consumes: Task 6의 화면, Task 5의 vendor 파일
- Produces: 검증 기록(원장)

내장 브라우저(`mcp__Claude_Browser__*`)로 `file:///C:/Users/KSA/Desktop/실습/web/index.html`을 연다. `file://`이 막히면 `python -m http.server`를 `web/`에서 띄우고 `http://localhost`로 확인한다(사용자는 `file://`로 쓰므로 가능하면 `file://`도 한 번 확인).

- [ ] **Step 1: 로드 확인** — 콘솔 오류 없음, `window.TR`, `window.TRDeck`, `window.PptxGenJS` 모두 존재, 확인 패널에 "팀원이 없습니다" 차단 사유가 보이고 PPT 버튼이 비활성.

- [ ] **Step 2: 입력 흐름** — `javascript_tool`로 `input` 이벤트를 발생시켜 실제 핸들러를 거치게 한다: 연 2026 / 월 9, 공휴일 2개(9/24, 9/25), 팀원 3명(이름 하나는 `<b>이서연</b>`), 각각 프로젝트 1~4개, 휴가 1건, 한 팀원 제출 시각을 마감 후로. 확인할 것:
  - 확인 패널의 근무일·시간·%가 Task 2 계산과 일치, 팀원 수 경고 표시
  - `<b>이서연</b>`이 굵은 글씨가 아니라 글자 그대로 보임, `document.querySelector('#panel b')`가 null
  - 지연 표시가 마감 후 제출자에게만 나옴
  - 프로젝트 이름을 바꾸면 월 요약 칸 목록이 따라 바뀜, 요약 칸에 타이핑 중 포커스가 유지됨

- [ ] **Step 3: 저장·복원** — 페이지 새로고침 후 입력이 그대로 복원되는지. `localStorage.setItem('teamReport.state','{깨진')` 후 새로고침하면 빈 화면 + 안내 문구(`#notice`)가 뜨는지. 백업 저장 파일 내용이 `parseState`를 통과하는지(`TR.parseState(...)`로 확인).

- [ ] **Step 4: PPT 생성** — 차단 사유를 해소한 뒤(요약 입력 포함) PPT 버튼이 활성화되는지. 클릭 시 오류 알림이 없는지(경고 확인 대화상자는 `confirm` 재정의로 통과). 페이지 안에서 `await TRDeck.buildPptx(PptxGenJS, TR.buildModel(state)).write({outputType:'base64'})`가 문자열을 반환하고 길이가 충분한지 확인.

- [ ] **Step 5: 전체 테스트 재실행**

Run: `node --test "web-tests/*.test.js"`
Expected: 28 tests pass, 0 fail

- [ ] **Step 6: 사람 확인 항목 정리** — 최종 보고에 "PPT 글자 배치·넘침은 PowerPoint에서 육안 확인 필요", 스크린샷은 참고용임을 명시한다.

---

## Self-Review

- **Spec coverage:** §1 결정(서버 없음·담당자 단독 입력·직접 요약) → Task 6 화면. §3 화면 1–6 → Task 6(설정, 달력, 팀원 카드, 팀 요약, 확인 패널, 생성·백업). §4 상태·자동저장·백업 → Task 3, 6. §5 계산 → Task 1–2. §6 경고·차단 각 항목 → Task 2 테스트(팀원 수, 휴가 이름, 금요일, 빈 요약, 다른 달, 표기 합침 / 팀원 0, 이름 빈/중복, 보고 없음, 프로젝트 빈, 날짜 오류, 제출 시각, 연·월). §7 PPT → Task 4, 5. §8 테스트 → Task 1–5, 7. §9 가정 → Global Constraints·Task 5 Step 1.
- **Placeholder scan:** 없음(모든 코드 단계에 실제 코드).
- **Type consistency:** `buildModel` 반환 필드(`members[].projects[].{name,bodies,summary}`, `lateWeeks`, `workdays`, `baseHours`, `perProjectHours`, `personalPct`, `teamSharePct`, `projectSharePct`)가 Task 2 정의와 Task 4·6 사용에서 일치. `summaryKey`는 Task 1 정의, Task 2·6 사용. 스펙 §4 vacations 표현 차이는 Global Constraints의 Ruling으로 기록.
- **Known gaps:** 반차, 휴가 범위 입력, 팀원 각자 제출은 범위 밖. PPT 시각 검증은 사람이 함.
