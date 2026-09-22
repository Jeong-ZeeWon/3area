/*******************************************************************
 * 성민교회 3교구 출석 현황 (관리자용, 읽기 전용)
 * 파일: admin/Code.gs  — 지역장 출석 앱과 별개인 Apps Script 프로젝트
 *
 * - 출석부 시트에 있는 정보(이름·연락처·순·순장·주일/순모임 출석)만 읽어
 *   교구 전체 현황을 보여 줍니다. 시트에 쓰는 기능은 없습니다.
 * - '웹 앱에 액세스하는 사용자' 권한으로 실행되도록 배포합니다(appsscript.json).
 *   그래서 출석부를 공유받은 사람(목사님·교구장)만 열 수 있고, 링크가 퍼져도
 *   출석부 권한이 없는 계정에는 아무것도 보이지 않습니다.
 *   출석부 공유는 '제한됨'으로 두고 볼 사람의 Gmail만 추가하세요.
 * - 아래 "시트 해석" 함수들은 지역장 앱(src/Code.gs)과 똑같은 것을 복사해 둔 것입니다.
 *   두 앱의 숫자가 어긋나지 않도록, 한쪽을 고치면 다른 쪽도 같이 고치세요.
 *******************************************************************/

const CONFIG = {
  // 비워 두면 스크립트 속성 SPREADSHEET_ID 를 쓰고, 그것도 없으면 붙어 있는 시트를 씁니다.
  SPREADSHEET_ID: '',

  APP_TITLE: '3교구 출석 현황',

  // 날짜가 시작되는 열 (D열 = 4)
  FIRST_DATE_COL: 4,

  // 계산 결과를 이 시간(초) 동안 저장해 두었다가 바로 보여 줍니다. [새로고침]은 항상 새로 계산합니다.
  CACHE_SECONDS: 600
};

/* ------------------------------ 웹앱 진입점 ------------------------------ */

function doGet() {
  const t = HtmlService.createTemplateFromFile('Index');
  t.appTitle = CONFIG.APP_TITLE;
  return t.evaluate()
    .setTitle(CONFIG.APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ------------------------------ 현황 데이터 ------------------------------ */

const CACHE_KEY = 'dashboard_v1';

/** 웹앱에서 호출. force 가 참이면 저장해 둔 결과를 버리고 시트를 다시 읽는다. */
function getDashboard(force) {
  // 저장해 둔 결과는 모든 사용자가 함께 쓰므로, 먼저 이 사람이 출석부를 열 수 있는지 확인한다.
  // 권한이 없으면 여기서 오류가 나고 아무것도 돌려주지 않는다.
  ss_();

  const cache = CacheService.getScriptCache();
  if (!force) {
    const hit = cache.get(CACHE_KEY);
    if (hit) return JSON.parse(hit);
  }
  const data = buildDashboard_();
  try {
    cache.put(CACHE_KEY, JSON.stringify(data), CONFIG.CACHE_SECONDS);
  } catch (err) {
    // 결과가 너무 크면 저장만 건너뛴다(화면에는 그대로 보여 준다)
  }
  return data;
}

/**
 * 모든 지역 시트를 읽어 '오늘까지 지난 주'의 출결을 모은다.
 * 출석 표시는 주마다 한 글자: '1' 출석, '0' 결석, '-' 그 지역 시트에 그 주 칸이 없음.
 * 지역장 행은 순 안의 같은 이름과 겹치므로 명단에서 빼고, 연락처만 기록 점검용으로 남긴다.
 */
function buildDashboard_() {
  const now = new Date();
  const tz = ss_().getSpreadsheetTimeZone();
  const todayKey = Number(Utilities.formatDate(now, tz, 'yyyyMMdd'));

  const master = {};
  const parsed = [];
  const errors = [];

  getRegions().forEach(function (r) {
    let d;
    try {
      d = parseSheet_(sheetByRegionId_(r.id));
    } catch (err) {
      errors.push(r.title + ': ' + err.message);
      return;
    }
    const keys = dateKeys_(d.dates, now);
    keys.forEach(function (k, i) {
      if (!master[k]) master[k] = { month: d.dates[i].month, day: d.dates[i].day, label: d.dates[i].label };
    });
    parsed.push({ d: d, keys: keys });
  });

  const weekKeys = Object.keys(master).map(Number)
    .filter(function (k) { return k <= todayKey; })
    .sort(function (a, b) { return a - b; });

  const regions = parsed.map(function (p) {
    const at = {};
    p.keys.forEach(function (k, i) { at[k] = i; });

    const encode = function (marks, colKey) {
      return weekKeys.map(function (k) {
        const i = at[k];
        if (i == null || !p.d.dates[i][colKey]) return '-';
        return marks[i] ? '1' : '0';
      }).join('');
    };

    let leaderPhone = '';
    const groups = [];
    p.d.groups.forEach(function (g) {
      if (g.name === '지역장') {
        if (g.members.length) leaderPhone = g.members[0].phone;
        return;
      }
      groups.push({
        name: g.name,
        members: g.members.map(function (m) {
          return { name: m.name, phone: m.phone, lead: !!m.lead, w: encode(m.marks, 'col'), g: encode(m.gmarks, 'gcol') };
        })
      });
    });

    return {
      id: p.d.id,
      title: p.d.title,
      leader: p.d.leader,
      leaderPhone: leaderPhone,
      hasGroup: p.d.hasGroup,
      warnings: p.d.warnings,
      groups: groups
    };
  });

  const first = regions.length ? regions[0].title : '';
  return {
    title: first.indexOf('_') > 0 ? first.split('_')[0] : CONFIG.APP_TITLE,
    at: Utilities.formatDate(now, tz, 'M월 d일 HH:mm'),
    weeks: weekKeys.map(function (k) { return master[k]; }),
    regions: regions,
    errors: errors
  };
}

/* ------------------------------ 시트 해석 (src/Code.gs 와 동일) ------------------------------ */

function prop_(key) {
  try {
    return PropertiesService.getScriptProperties().getProperty(key) || '';
  } catch (err) {
    return '';
  }
}

function ss_() {
  const id = CONFIG.SPREADSHEET_ID || prop_('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);

  const book = SpreadsheetApp.getActiveSpreadsheet();
  if (book) return book;

  throw new Error(
    '출석부 스프레드시트를 찾을 수 없습니다. 이 스크립트가 시트에 붙어 있지 않은 ' +
    '독립 프로젝트라면, setSpreadsheetId() 함수 안에 시트 ID를 넣고 편집기에서 ' +
    '한 번 실행해 주세요.');
}

/**
 * 독립 프로젝트로 만든 경우 출석부 시트 ID를 한 번 저장해 둡니다.
 * 시트 주소 https://docs.google.com/spreadsheets/d/<이부분>/edit 의 가운데 문자열입니다.
 * 아래 id 에 넣고 편집기에서 이 함수를 한 번 실행하세요.
 */
function setSpreadsheetId() {
  const id = '';  // ← 여기에 출석부 스프레드시트 ID를 넣으세요

  if (!id) throw new Error('setSpreadsheetId() 안의 id 에 스프레드시트 ID를 먼저 넣어 주세요.');
  const book = SpreadsheetApp.openById(id);   // 잘못된 ID면 여기서 바로 오류가 납니다
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', id);
  Logger.log('저장했습니다: %s (시트 %s개)', book.getName(), book.getSheets().length);
  return book.getName();
}

function sheetByRegionId_(id) {
  const sheet = ss_().getSheetByName(String(id));
  if (!sheet) throw new Error('지역 시트를 찾을 수 없습니다: ' + id);
  return sheet;
}

function parseMonth_(v) {
  if (v instanceof Date) return v.getMonth() + 1;
  const s = String(v == null ? '' : v).trim();
  if (!s) return 0;
  const m = s.match(/(\d{1,2})\s*월/);
  if (m) return Number(m[1]);
  const n = Number(s);
  return n >= 1 && n <= 12 ? n : 0;
}

function parseDay_(v) {
  if (v instanceof Date) return v.getDate();
  const s = String(v == null ? '' : v).trim();
  if (!s) return 0;
  const m = s.match(/(\d{1,2})/);
  if (!m) return 0;
  const n = Number(m[1]);
  return n >= 1 && n <= 31 ? n : 0;
}

function findHeaderRow_(head) {
  for (let i = 0; i < head.length; i++) {
    const a = String(head[i][0] == null ? '' : head[i][0]).trim();
    const b = String(head[i][1] == null ? '' : head[i][1]).trim();
    if (a === '구분' && b === '순원') return i + 1;
  }
  return -1;
}

function getRegions() {
  const out = [];
  ss_().getSheets().forEach(function (sh) {
    if (sh.isSheetHidden()) return;
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 3 || lastCol < 3) return;

    const head = sh.getRange(1, 1, Math.min(8, lastRow), Math.min(lastCol, 40)).getValues();
    const hr = findHeaderRow_(head);
    if (hr === -1) return;

    const title = String(head[0][0] == null ? '' : head[0][0]).trim() || sh.getName();
    let leader = '';
    for (let i = hr; i < head.length; i++) {
      if (String(head[i][0] == null ? '' : head[i][0]).trim() === '지역장') {
        leader = String(head[i][1] == null ? '' : head[i][1]).trim();
        break;
      }
    }
    out.push({ id: sh.getName(), title: title, leader: leader });
  });
  return out;
}

function parseSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 3 || lastCol < CONFIG.FIRST_DATE_COL) {
    throw new Error('출석부 형식이 아닌 시트입니다: ' + sheet.getName());
  }

  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  // B열(순원) 이름이 굵은 글씨면 순장으로 봅니다. 시트 서식을 그대로 읽습니다.
  let weights = [];
  try {
    weights = sheet.getRange(1, 2, lastRow, 1).getFontWeights();
  } catch (err) {
    weights = [];
  }

  const headerRow = findHeaderRow_(values.slice(0, Math.min(8, lastRow)));
  if (headerRow === -1) {
    throw new Error('머리글(구분·순원) 행을 찾을 수 없습니다: ' + sheet.getName());
  }

  const dateRow = values[headerRow - 1];

  // 월 줄: 날짜 줄 위쪽에서 'N월'이 적힌 가장 가까운 줄
  let monthRow = [];
  let monthIdx = -1;
  for (let i = headerRow - 2; i >= 0; i--) {
    if (hasLabel_(values[i], lastCol, /\d{1,2}\s*월/)) { monthRow = values[i]; monthIdx = i; break; }
  }

  // 소제목 줄(주일출석 | 순모임출석): 날짜 줄 위쪽, 없으면 바로 아랫줄
  let subRow = null;
  let subBelow = false;
  for (let i = headerRow - 2; i >= 0; i--) {
    if (i !== monthIdx && hasLabel_(values[i], lastCol, /순모임|예배|주일/)) { subRow = values[i]; break; }
  }
  if (!subRow && findSubRow_(values[headerRow], lastCol)) {
    subRow = values[headerRow];
    subBelow = true;
  }

  // 날짜 칸을 주일 칸과 순모임 칸으로 나눈다.
  // 병합된 월 표시는 첫 칸에만 값이 있으므로 오른쪽으로 이어서 적용한다.
  const wList = [];
  const gList = [];
  const warnings = [];
  let month = 0;
  let inGroup = false;
  let lastDay = 0;
  let prev = null;
  for (let c = CONFIG.FIRST_DATE_COL - 1; c < lastCol; c++) {
    const head = String(monthRow[c] == null ? '' : monthRow[c]).trim();
    if (head) {
      inGroup = isGroupText_(head);
      month = parseMonth_(head) || (inGroup ? 0 : month);
    }
    const cell = dateRow[c];
    const sub = subRow ? String(subRow[c] == null ? '' : subRow[c]).trim() : '';
    let day = parseDay_(cell);
    let guessed = false;
    if (!day && subRow) {
      if (/^\d{1,2}$/.test(sub)) {        // 날짜가 소제목 줄에 잘못 들어간 경우
        day = Number(sub);
        guessed = true;
      } else if (sub && lastDay) {        // 날짜 한 칸이 (주일|순모임) 두 칸 위로 병합된 경우
        day = lastDay;
      }
    }
    if (!day) { prev = null; continue; }
    lastDay = day;

    const m = cell instanceof Date ? cell.getMonth() + 1 : month;
    let isGroup = inGroup || isGroupText_(cell) || isGroupText_(sub);
    // 소제목이 비었거나 잘못 적힌 칸은 바로 앞 주일 칸의 짝(순모임)으로 본다
    if (!isGroup && subRow && !/주일|예배/.test(sub) &&
        prev && !prev.isGroup && prev.day === day && prev.month === m) {
      isGroup = true;
      guessed = true;
    }
    const item = { col: c + 1, month: m, day: day, isGroup: isGroup };
    if (guessed) {
      warnings.push(colLetter_(c + 1) + '열(' + m + '월 ' + day + '일 ' + (isGroup ? '순모임' : '주일') +
        ')의 머리글이 비었거나 다른 줄에 적혀 있어 짐작해서 읽었습니다');
    }
    (isGroup ? gList : wList).push(item);
    prev = item;
  }

  const dates = [];
  const byLabel = {};
  wList.forEach(function (x) {
    if (!x.month) return;
    const label = x.month + '월 ' + x.day + '일';
    if (byLabel[label]) return;
    const d = { col: x.col, gcol: null, month: x.month, day: x.day, label: label };
    byLabel[label] = d;
    dates.push(d);
  });
  if (!dates.length) throw new Error('날짜 열을 찾을 수 없습니다: ' + sheet.getName());

  // 순모임 칸은 같은 날짜의 주일 칸과 짝을 짓는다. 월 표시가 없으면 순서대로 맞춘다.
  gList.forEach(function (x, k) {
    let d = x.month ? byLabel[x.month + '월 ' + x.day + '일'] : null;
    if (!d && !x.month && dates[k] && dates[k].day === x.day) d = dates[k];
    if (d && !d.gcol) d.gcol = x.col;
  });
  const hasGroup = dates.some(function (d) { return !!d.gcol; });

  const filled = function (r, col) {
    if (!col) return false;
    return String(values[r][col - 1] == null ? '' : values[r][col - 1]).trim() !== '';
  };

  // 명단: A열 값이 나오면 새 그룹, B열 이름이 있으면 순원
  const groups = [];
  let cur = null;
  for (let r = headerRow + (subBelow ? 1 : 0); r < lastRow; r++) {
    const row = r + 1;
    const gname = String(values[r][0] == null ? '' : values[r][0]).trim();
    const name = String(values[r][1] == null ? '' : values[r][1]).trim();
    if (gname) {
      cur = { name: gname, members: [] };
      groups.push(cur);
    }
    if (!name) continue;
    if (!cur) {
      cur = { name: '명단', members: [] };
      groups.push(cur);
    }
    const w = weights[r] ? String(weights[r][0] == null ? '' : weights[r][0]).toLowerCase() : '';
    cur.members.push({
      row: row,
      name: name,
      phone: String(values[r][2] == null ? '' : values[r][2]).trim(),
      lead: (w === 'bold' || w === '700' || w === 'bolder'),
      marks: dates.map(function (d) { return filled(r, d.col); }),
      gmarks: dates.map(function (d) { return filled(r, d.gcol); })
    });
  }

  const title = String(values[0][0] == null ? '' : values[0][0]).trim() || sheet.getName();
  let leader = '';
  groups.forEach(function (g) {
    if (g.name === '지역장' && g.members.length) leader = g.members[0].name;
  });

  return {
    id: sheet.getName(),
    title: title,
    leader: leader,
    headerRow: headerRow,
    dates: dates,
    hasGroup: hasGroup,
    warnings: warnings,
    groups: groups.filter(function (g) { return g.members.length; })
  };
}

function isGroupText_(v) {
  return /순모임/.test(String(v == null ? '' : v));
}

function hasLabel_(row, lastCol, re) {
  if (!row) return false;
  for (let c = CONFIG.FIRST_DATE_COL - 1; c < lastCol; c++) {
    if (re.test(String(row[c] == null ? '' : row[c]))) return true;
  }
  return false;
}

function colLetter_(n) {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/**
 * 머리글 바로 아랫줄이 (예배 | 순모임) 같은 소제목 줄인지 본다.
 * A·B열이 비어 있고 날짜 칸 어딘가에 '순모임'·'예배'·'주일'이 적혀 있으면 소제목 줄이다.
 */
function findSubRow_(row, lastCol) {
  if (!row) return null;
  if (String(row[0] == null ? '' : row[0]).trim() || String(row[1] == null ? '' : row[1]).trim()) return null;
  for (let c = CONFIG.FIRST_DATE_COL - 1; c < lastCol; c++) {
    if (/순모임|예배|주일/.test(String(row[c] == null ? '' : row[c]))) return row;
  }
  return null;
}

/**
 * 시트의 월·일에 연도를 붙여 yyyymmdd 숫자로 만든다.
 * 9월~이듬해 2월처럼 해를 넘기는 출석부도 순서가 유지되도록 월이 줄어들면 해를 올린다.
 * 첫 달이 오늘보다 넉 달 넘게 뒤라면 작년에 시작한 출석부로 본다(예: 1월에 9월 시작 출석부).
 */
function dateKeys_(dates, today) {
  if (!dates.length) return [];
  const first = dates[0].month;
  let year = today.getFullYear();
  if (first > today.getMonth() + 1 + 4) year -= 1;
  let prev = first;
  return dates.map(function (d) {
    if (d.month < prev) year += 1;
    prev = d.month;
    return year * 10000 + d.month * 100 + d.day;
  });
}
