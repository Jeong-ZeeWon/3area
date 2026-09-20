/*******************************************************************
 * 성민교회 3교구 지역별 출석 체크 웹앱
 * 파일: Code.gs
 *
 * - 지역별 시트 구조(1행 월 병합 / 2행 구분·순원·연락처·날짜 / 3행 지역장)를
 *   자동으로 인식하므로 탭 이름이나 월이 추가되어도 코드 수정이 필요 없습니다.
 * - 앱에서 체크하면 해당 지역 시트의 해당 날짜 칸에 바로 기록됩니다.
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
  PropertiesService.getScriptProperties().setProperty('ACCESS_CODES', JSON.stringify(codes));
  Logger.log('접속코드 %s개를 저장했습니다.', Object.keys(codes).length);
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
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
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

  const monthRow = headerRow >= 2 ? values[headerRow - 2] : [];
  const dateRow = values[headerRow - 1];

  // 병합된 월 표시는 첫 칸에만 값이 있으므로 오른쪽으로 이어서 채운다
  const dates = [];
  let curMonth = 0;
  for (let c = CONFIG.FIRST_DATE_COL - 1; c < lastCol; c++) {
    const m = parseMonth_(monthRow[c]);
    if (m) curMonth = m;
    const cell = dateRow[c];
    const day = parseDay_(cell);
    if (!day) continue;
    const month = cell instanceof Date ? cell.getMonth() + 1 : curMonth;
    if (!month) continue;
    dates.push({ col: c + 1, month: month, day: day, label: month + '월 ' + day + '일' });
  }
  if (!dates.length) throw new Error('날짜 열을 찾을 수 없습니다: ' + sheet.getName());

  // 명단: A열 값이 나오면 새 그룹, B열 이름이 있으면 순원
  const groups = [];
  let cur = null;
  for (let r = headerRow; r < lastRow; r++) {
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
      marks: dates.map(function (d) {
        return String(values[r][d.col - 1] == null ? '' : values[r][d.col - 1]).trim() !== '';
      })
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
    groups: groups.filter(function (g) { return g.members.length; })
  };
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
    layout.dates.forEach(function (d) { validCol[d.col] = true; });

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
    layout.dates.forEach(function (d) { colLabel[d.col] = d.label; });

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

/* ------------------------------ 배부용 도구 ------------------------------ */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('출석 체크 앱')
    .addItem('지역장 배부 링크 만들기', 'makeRegionLinks')
    .addItem('출석부 구조 점검', 'checkStructure')
    .addToUi();
}

/** 지역별 개인 링크(?r=시트이름)를 _배부링크 시트에 정리해 줍니다 */
function makeRegionLinks() {
  const url = ScriptApp.getService().getUrl();
  if (!url) {
    SpreadsheetApp.getUi().alert('먼저 [배포] → [새 배포] 로 웹 앱을 배포해 주세요.');
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
        '\n    순장(굵은 글씨): ' + (leads.length ? leads.join(', ') : '없음'));
    } catch (err) {
      lines.push(r.title + ' — 오류: ' + err.message);
    }
  });
  const msg = lines.join('\n');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert('출석부 구조 점검', msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) {}
  return msg;
}
