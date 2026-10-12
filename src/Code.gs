/*******************************************************************
 * 성민교회 3교구 지역별 출석 체크 웹앱
 * 파일: Code.gs
 *
 * - 지역별 시트 구조를 매번 읽어 들이므로 탭 이름·월·명단이 바뀌어도 코드 수정이 필요 없습니다.
 *     1행 월(병합) / 2행 주일출석·순모임출석 / 3행 구분·순원·연락처·날짜 / 4행부터 명단
 *   순모임 소제목 줄이 없는 옛 모양(1행 월 / 2행 날짜)도 그대로 읽습니다.
 * - 앱에서 체크하면 해당 지역 시트의 해당 날짜 칸(주일 또는 순모임)에 바로 기록됩니다.
 *******************************************************************/

const CONFIG = {
  // 비워 두면 스크립트가 붙어 있는 스프레드시트를 그대로 사용합니다.
  // 다른 파일을 쓰려면 스크립트 속성 SPREADSHEET_ID 에 넣으세요(코드에는 적지 않습니다).
  SPREADSHEET_ID: '',

  APP_TITLE: '3교구 출석 체크',

  // 출석 표시 문자 ('O', '○', '1' 등으로 바꿔 쓸 수 있습니다)
  MARK: 'O',

  // 날짜가 시작되는 열 (D열 = 4)
  FIRST_DATE_COL: 4,

  // true 로 바꾸면 지역별 접속코드를 입력해야 명단이 열립니다
  USE_ACCESS_CODE: false,

  // 접속코드는 코드가 아니라 스크립트 속성 ACCESS_CODES(JSON)에 저장합니다.
  // setAccessCodes() 를 한 번 실행해 넣어 두세요.
  ACCESS_CODES: {},

  // 변경 기록 남기기
  ENABLE_LOG: true,
  LOG_SHEET_NAME: '_변경기록',
  LINK_SHEET_NAME: '_배부링크'
};

/* ------------------------------ 설정 값 ------------------------------ */

function prop_(key) {
  try {
    return PropertiesService.getScriptProperties().getProperty(key) || '';
  } catch (err) {
    return '';
  }
}

function accessCodes_() {
  const raw = prop_('ACCESS_CODES');
  if (raw) {
    try { return JSON.parse(raw); } catch (err) { /* 형식 오류는 무시 */ }
  }
  return CONFIG.ACCESS_CODES || {};
}

/**
 * 지역별 접속코드를 스크립트 속성에 저장합니다.
 * 아래 codes 를 교회 상황에 맞게 고친 뒤 편집기에서 한 번 실행하세요.
 */
function setAccessCodes() {
  const codes = {
    '3교구_1지역': '1001', '3교구_2지역': '1002', '3교구_3지역': '1003',
    '3교구_4지역': '1004', '3교구_5지역': '1005', '3교구_6지역': '1006',
    '3교구_7지역': '1007', '3교구_8지역': '1008', '3교구_9지역': '1009',
    '3교구_10지역': '1010'
  };
  // 전 지역 '출석 현황' 화면을 여는 관리자 코드
  const adminCode = '9999';

  const props = PropertiesService.getScriptProperties();
  props.setProperty('ACCESS_CODES', JSON.stringify(codes));
  props.setProperty('ADMIN_CODE', adminCode);
  Logger.log('접속코드 %s개와 관리자 코드를 저장했습니다.', Object.keys(codes).length);
}

/* ------------------------------ 웹앱 진입점 ------------------------------ */

function doGet(e) {
  const params = (e && e.parameter) || {};
  const t = HtmlService.createTemplateFromFile('Index');
  t.appTitle = CONFIG.APP_TITLE;
  t.useCode = CONFIG.USE_ACCESS_CODE;
  t.lockedRegion = params.r ? String(params.r) : '';
  return t.evaluate()
    .setTitle(CONFIG.APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ------------------------------ 시트 접근 ------------------------------ */

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

/* ------------------------------ 파싱 도우미 ------------------------------ */

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

/* ------------------------------ 지역 목록 ------------------------------ */

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

/* ------------------------------ 지역 명단 ------------------------------ */

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

function defaultDateIndex_(dates) {
  const now = new Date();
  const y = now.getFullYear();
  let best = 0;
  let bestDiff = Infinity;
  dates.forEach(function (d, i) {
    const diff = Math.abs(new Date(y, d.month - 1, d.day).getTime() - now.getTime());
    if (diff < bestDiff) { bestDiff = diff; best = i; }
  });
  return best;
}

function checkCode_(id, code) {
  if (!CONFIG.USE_ACCESS_CODE) return;
  const sheet = sheetByRegionId_(id);
  const title = String(sheet.getRange(1, 1).getValue()).trim() || sheet.getName();
  const codes = accessCodes_();
  const expect = codes[title] || codes[sheet.getName()];
  if (!expect) return;
  if (String(code == null ? '' : code).trim() !== String(expect)) {
    throw new Error('접속코드가 올바르지 않습니다.');
  }
}

/** 웹앱에서 호출: 지역 명단 + 기존 출석 표시를 모두 가져온다 */
function getRegionData(id, code) {
  checkCode_(id, code);
  const data = parseSheet_(sheetByRegionId_(id));
  data.defaultDateIndex = defaultDateIndex_(data.dates);
  data.mark = CONFIG.MARK;

  // 오늘까지 지난 주인지 표시한다(순원별 누적 출석 횟수 계산용)
  const now = new Date();
  const todayKey = Number(Utilities.formatDate(now, ss_().getSpreadsheetTimeZone(), 'yyyyMMdd'));
  const keys = dateKeys_(data.dates, now);
  data.dates.forEach(function (d, i) { d.past = keys[i] <= todayKey; });
  return data;
}

/* ------------------------------ 저장 ------------------------------ */

/**
 * payload = { id: '시트이름', code: '접속코드', items: [{row, col, checked}] }
 * 같은 열은 한 번에 묶어 기록하므로 체크를 여러 개 해도 호출이 가볍습니다.
 */
function saveAttendance(payload) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) throw new Error('다른 저장이 진행 중입니다. 잠시 후 다시 시도해 주세요.');
  try {
    payload = payload || {};
    checkCode_(payload.id, payload.code);

    const sheet = sheetByRegionId_(payload.id);
    const layout = parseSheet_(sheet);

    const validRow = {};
    layout.groups.forEach(function (g) {
      g.members.forEach(function (m) { validRow[m.row] = true; });
    });
    const validCol = {};
    layout.dates.forEach(function (d) {
      if (d.col) validCol[d.col] = true;
      if (d.gcol) validCol[d.gcol] = true;
    });

    const byCol = {};
    (payload.items || []).forEach(function (it) {
      const row = Number(it.row);
      const col = Number(it.col);
      if (!validRow[row] || !validCol[col]) return;
      if (!byCol[col]) byCol[col] = [];
      byCol[col].push({ row: row, checked: !!it.checked });
    });

    let saved = 0;
    Object.keys(byCol).forEach(function (key) {
      const col = Number(key);
      const rows = byCol[key];
      const minRow = Math.min.apply(null, rows.map(function (o) { return o.row; }));
      const maxRow = Math.max.apply(null, rows.map(function (o) { return o.row; }));
      const range = sheet.getRange(minRow, col, maxRow - minRow + 1, 1);
      const vals = range.getValues();
      rows.forEach(function (o) {
        vals[o.row - minRow][0] = o.checked ? CONFIG.MARK : '';
        saved++;
      });
      range.setValues(vals);
    });

    SpreadsheetApp.flush();
    writeLog_(layout, byCol);

    const tz = ss_().getSpreadsheetTimeZone();
    return { ok: true, saved: saved, at: Utilities.formatDate(new Date(), tz, 'HH:mm:ss') };
  } finally {
    lock.releaseLock();
  }
}

function writeLog_(layout, byCol) {
  if (!CONFIG.ENABLE_LOG) return;
  try {
    const book = ss_();
    let log = book.getSheetByName(CONFIG.LOG_SHEET_NAME);
    if (!log) {
      log = book.insertSheet(CONFIG.LOG_SHEET_NAME);
      log.appendRow(['시각', '지역', '날짜', '변경 건수', '작성자']);
      log.setFrozenRows(1);
      log.hideSheet();
    }
    let email = '';
    try { email = Session.getActiveUser().getEmail() || ''; } catch (err) { email = ''; }

    const colLabel = {};
    layout.dates.forEach(function (d) {
      if (d.col) colLabel[d.col] = d.label + ' 예배';
      if (d.gcol) colLabel[d.gcol] = d.label + ' 순모임';
    });

    Object.keys(byCol).forEach(function (key) {
      log.appendRow([
        new Date(),
        layout.title,
        colLabel[Number(key)] || key,
        byCol[key].length,
        email
      ]);
    });
  } catch (err) {
    // 기록 실패가 출석 저장을 막지 않도록 조용히 넘어감
  }
}

/* ------------------------------ 출석 현황 ------------------------------ */

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

function checkAdmin_(code) {
  if (!CONFIG.USE_ACCESS_CODE) return;
  const expect = prop_('ADMIN_CODE');
  if (!expect || String(code == null ? '' : code).trim() !== String(expect)) {
    throw new Error('관리자 코드가 올바르지 않습니다.');
  }
}

/**
 * 웹앱에서 호출: 모든 지역의 지난 주일 출결을 모아 돌려준다.
 * 지역장 행은 순 안의 이름과 겹치므로 빼고, 오늘 이후 날짜는 넣지 않는다.
 * marks 는 1(출석) / 0(결석) / null(그 지역 시트에 없는 날짜)이다.
 */
function getOverview(code) {
  checkAdmin_(code);

  const now = new Date();
  const tz = ss_().getSpreadsheetTimeZone();
  const todayKey = Number(Utilities.formatDate(now, tz, 'yyyyMMdd'));

  const master = {};
  const parsed = [];

  getRegions().forEach(function (r) {
    let d;
    try {
      d = parseSheet_(sheetByRegionId_(r.id));
    } catch (err) {
      return;
    }
    const keys = dateKeys_(d.dates, now);
    keys.forEach(function (k, i) {
      if (!master[k]) {
        master[k] = { month: d.dates[i].month, day: d.dates[i].day, label: d.dates[i].label };
      }
    });
    parsed.push({ d: d, keys: keys });
  });

  const past = Object.keys(master).map(Number)
    .filter(function (k) { return k <= todayKey; })
    .sort(function (a, b) { return a - b; });

  const regions = parsed.map(function (p) {
    const at = {};
    p.keys.forEach(function (k, i) { at[k] = i; });

    const groups = p.d.groups
      .filter(function (g) { return g.name !== '지역장'; })
      .map(function (g) {
        return {
          name: g.name,
          members: g.members.map(function (m) {
            return {
              name: m.name,
              lead: !!m.lead,
              marks: past.map(function (k) {
                return at[k] == null ? null : (m.marks[at[k]] ? 1 : 0);
              }),
              gmarks: past.map(function (k) {
                if (at[k] == null || !p.d.dates[at[k]].gcol) return null;
                return m.gmarks[at[k]] ? 1 : 0;
              })
            };
          })
        };
      });

    // 그 지역에서 한 명이라도 체크한 주만 '기록된 주'로 본다(예배·순모임 따로)
    const anyOn = function (key, j) {
      return groups.some(function (g) {
        return g.members.some(function (m) { return m[key][j] === 1; });
      });
    };

    return {
      id: p.d.id,
      title: p.d.title,
      leader: p.d.leader,
      hasGroup: p.d.hasGroup,
      recorded: past.map(function (k, j) { return anyOn('marks', j); }),
      grecorded: past.map(function (k, j) { return anyOn('gmarks', j); }),
      groups: groups
    };
  });

  const firstTitle = regions.length ? regions[0].title : '';
  return {
    title: firstTitle.indexOf('_') > 0 ? firstTitle.split('_')[0] : CONFIG.APP_TITLE,
    dates: past.map(function (k) { return master[k]; }),
    regions: regions,
    at: Utilities.formatDate(now, tz, 'M월 d일 HH:mm')
  };
}

/* ------------------------------ 배부용 도구 ------------------------------ */

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('출석 체크 앱')
      .addItem('지역장 배부 링크 만들기', 'makeRegionLinks')
      .addItem('출석부 구조 점검', 'checkStructure')
      .addToUi();
  } catch (err) {
    // 독립 프로젝트에는 메뉴를 붙일 수 없습니다. 편집기에서 함수를 직접 실행하세요.
  }
}

/** 시트 화면이 없을 때는 조용히 넘어가는 알림 */
function notify_(title, msg) {
  Logger.log('%s\n%s', title, msg);
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert(title, msg, ui.ButtonSet.OK);
  } catch (err) {
    // 편집기에서 실행한 경우: 실행 기록(Logger)으로만 확인합니다.
  }
}

/** 지역별 개인 링크(?r=시트이름)를 _배부링크 시트에 정리해 줍니다 */
function makeRegionLinks() {
  const url = ScriptApp.getService().getUrl();
  if (!url) {
    notify_('배포가 필요합니다', '먼저 [배포] → [새 배포] 로 웹 앱을 배포해 주세요.');
    return;
  }
  const book = ss_();
  let sheet = book.getSheetByName(CONFIG.LINK_SHEET_NAME);
  if (!sheet) sheet = book.insertSheet(CONFIG.LINK_SHEET_NAME);
  sheet.clear();
  sheet.appendRow(['지역', '지역장', '접속코드', '배부 링크']);
  sheet.setFrozenRows(1);

  const codes = accessCodes_();
  getRegions().forEach(function (r) {
    const code = CONFIG.USE_ACCESS_CODE ? (codes[r.title] || codes[r.id] || '') : '';
    sheet.appendRow([
      r.title,
      r.leader,
      code,
      url + '?r=' + encodeURIComponent(r.id)
    ]);
  });
  sheet.autoResizeColumns(1, 4);
  book.setActiveSheet(sheet);
}

/** 각 지역 시트가 제대로 인식되는지 확인 (실행 후 [실행 기록] 확인) */
function checkStructure() {
  const lines = [];
  getRegions().forEach(function (r) {
    try {
      const d = parseSheet_(sheetByRegionId_(r.id));
      const total = d.groups.reduce(function (s, g) { return s + g.members.length; }, 0);
      const leads = [];
      d.groups.forEach(function (g) {
        if (g.name === '지역장') return;
        g.members.forEach(function (m) { if (m.lead) leads.push(g.name + ' ' + m.name); });
      });
      lines.push(d.title + ' — 지역장 ' + (d.leader || '미지정') +
        ' / 순 ' + d.groups.length + '개 / 인원 ' + total + '명 / 날짜 ' +
        d.dates.length + '개 (' + d.dates[0].label + ' ~ ' + d.dates[d.dates.length - 1].label + ')' +
        '\n    순모임 칸: ' + (d.hasGroup
          ? d.dates.filter(function (x) { return x.gcol; }).length + '주 인식'
          : '없음 (날짜 위 줄에 주일출석·순모임출석을 적어 날짜마다 두 칸으로 나누세요)') +
        '\n    순장(굵은 글씨): ' + (leads.length ? leads.join(', ') : '없음') +
        (d.warnings.length ? '\n    ⚠ ' + d.warnings.join('\n    ⚠ ') : ''));
    } catch (err) {
      lines.push(r.title + ' — 오류: ' + err.message);
    }
  });
  const msg = lines.join('\n');
  notify_('출석부 구조 점검', msg);
  return msg;
}
