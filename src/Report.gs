/*******************************************************************
 * 주간 출석 리포트 (구글 문서)
 * 파일: Report.gs — 지역장 출석 앱(src) 프로젝트에 함께 넣습니다.
 *
 * - 관리자 현황 앱과 같은 기준으로 셈한 한 주의 출석 현황을 구글 문서 한 편으로 만듭니다.
 *   문서는 내 드라이브의 '주간 출석 리포트' 폴더에 주일 날짜별로 쌓입니다.
 * - 시트 메뉴 [출석 체크 앱 → 주간 리포트 지금 만들기]로 바로 만들고,
 *   [주간 리포트 매주 자동 만들기]를 한 번 누르면 매주 월요일 아침에 저절로 만들어집니다.
 * - 같은 주일의 리포트를 다시 만들면 새 문서를 만들지 않고 그 문서 내용을 새로 씁니다.
 * - 이 프로젝트는 목사님 계정으로 실행되므로 리포트 문서도 목사님 드라이브에 생깁니다.
 *   교구장 등 함께 볼 사람은 폴더를 공유하거나 setReportEmails() 로 메일 받을 사람을 넣으세요.
 *******************************************************************/

const REPORT = {
  FOLDER_NAME: '주간 출석 리포트',

  // 자동 생성 요일과 시각(이 시각부터 한 시간 안에 실행됩니다)
  WEEKDAY: 'MONDAY',
  HOUR: 8,

  // '최근 흐름' 표에 넣을 주 수
  TREND_WEEKS: 6
};

/* ------------------------------ 메뉴에서 실행 ------------------------------ */

/** 가장 최근 주일의 리포트를 만들고 문서 주소를 돌려준다 (시간 트리거도 이 함수를 부른다) */
function makeWeeklyReport() {
  const D = reportData_();
  if (!D.weekKeys.length) {
    notify_('주간 리포트', '아직 지난 주일이 없어 리포트를 만들 수 없습니다.');
    return '';
  }

  // 오늘이 주일이라 아직 아무도 기록하지 않았다면 그 전 주일로 만든다
  let j = D.weekKeys.length - 1;
  while (j > 0 && !anyRecorded_(D, j)) j--;

  const doc = reportDoc_(D, j);
  writeReport_(doc, D, j);
  doc.saveAndClose();

  const url = doc.getUrl();
  mailReport_(D, j, doc, url);
  notify_('주간 리포트를 만들었습니다', doc.getName() + '\n\n' + url);
  return url;
}

/** 매주 월요일 아침에 리포트를 만들도록 예약한다. 여러 번 눌러도 예약은 하나만 남는다. */
function installWeeklyReport() {
  removeTriggers_();
  ScriptApp.newTrigger('makeWeeklyReport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay[REPORT.WEEKDAY])
    .atHour(REPORT.HOUR)
    .inTimezone(ss_().getSpreadsheetTimeZone())
    .create();
  notify_('주간 리포트 자동 만들기',
    '매주 ' + weekdayName_(REPORT.WEEKDAY) + ' ' + REPORT.HOUR + '시~' + (REPORT.HOUR + 1) +
    '시 사이에 지난 주일 리포트가 내 드라이브의 \'' + REPORT.FOLDER_NAME + '\' 폴더에 만들어집니다.');
}

function removeWeeklyReport() {
  const n = removeTriggers_();
  notify_('주간 리포트 자동 만들기', n ? '예약을 해제했습니다.' : '예약된 리포트가 없습니다.');
}

/**
 * 리포트가 만들어질 때 문서 링크를 메일로 받을 사람을 저장합니다(쉼표로 여러 명).
 * 받는 사람에게는 그 문서의 보기 권한이 함께 주어집니다.
 * 아래 emails 에 넣고 편집기에서 한 번 실행하세요. 비워 두고 실행하면 메일 보내기를 끕니다.
 * 주소는 스크립트 속성에 저장되므로 실행한 뒤에는 지워도 됩니다.
 */
function setReportEmails() {
  const emails = '';  // ← 예: 'a@gmail.com, b@gmail.com'

  const props = PropertiesService.getScriptProperties();
  if (emails.trim()) {
    props.setProperty('REPORT_EMAILS', emails);
    Logger.log('리포트 메일 받는 사람을 저장했습니다: %s', emails);
  } else {
    props.deleteProperty('REPORT_EMAILS');
    Logger.log('리포트 메일 보내기를 껐습니다.');
  }
}

function removeTriggers_() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'makeWeeklyReport') {
      ScriptApp.deleteTrigger(t);
      n++;
    }
  });
  return n;
}

function weekdayName_(w) {
  return { SUNDAY: '주일', MONDAY: '월요일', TUESDAY: '화요일', WEDNESDAY: '수요일',
    THURSDAY: '목요일', FRIDAY: '금요일', SATURDAY: '토요일' }[w] || w;
}

/* ------------------------------ 데이터 ------------------------------ */

/**
 * 관리자 현황 앱(admin/Code.gs buildDashboard_)과 같은 방식으로 '오늘까지 지난 주'의 출결을 모은다.
 * 주마다 '1' 출석, '0' 결석, '-' 그 지역 시트에 그 주 칸이 없음.
 * 지역장 행은 순 안의 같은 이름과 겹치므로 셈에서 빼고, 연락처만 기록 점검용으로 남긴다.
 */
function reportData_() {
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
      if (!master[k]) master[k] = { month: d.dates[i].month, day: d.dates[i].day };
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
      });
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
          return { name: m.name, lead: !!m.lead, w: encode(m.marks, 'col'), g: encode(m.gmarks, 'gcol') };
        })
      });
    });

    const r = {
      title: p.d.title,
      leader: p.d.leader,
      leaderPhone: leaderPhone,
      hasGroup: p.d.hasGroup,
      warnings: p.d.warnings,
      groups: groups
    };
    r.all = [];
    groups.forEach(function (g) { r.all = r.all.concat(g.members); });
    // 그 지역에서 한 명이라도 체크한 주만 '기록된 주'로 본다
    r.rec = {
      w: weekKeys.map(function (k, j) { return r.all.some(function (m) { return m.w[j] === '1'; }); }),
      g: weekKeys.map(function (k, j) { return r.all.some(function (m) { return m.g[j] === '1'; }); })
    };
    return r;
  });

  const first = regions.length ? regions[0].title : '';
  return {
    title: first.indexOf('_') > 0 ? first.split('_')[0] : CONFIG.APP_TITLE,
    tz: tz,
    now: now,
    weekKeys: weekKeys,
    weeks: weekKeys.map(function (k) { return master[k]; }),
    regions: regions,
    errors: errors
  };
}

/* ------------------------------ 계산 (admin/Index.html 과 같은 기준) ------------------------------ */

function applies_(r, kind) { return kind === 'w' || r.hasGroup; }

function anyRecorded_(D, j) {
  return D.regions.some(function (r) { return r.rec.w[j] || r.rec.g[j]; });
}

/** members 가운데 j주 출석 인원 / 대상 인원 (그 지역이 그 주를 기록한 경우에만) */
function tally_(r, members, kind, j) {
  if (!applies_(r, kind) || !r.rec[kind][j]) return null;
  let att = 0, base = 0;
  members.forEach(function (m) {
    const c = m[kind][j];
    if (c === '-') return;
    base++;
    if (c === '1') att++;
  });
  return base ? { att: att, base: base, rate: att / base } : null;
}

function weekStats_(D, kind, j) {
  let att = 0, base = 0, done = 0, total = 0;
  D.regions.forEach(function (r) {
    if (!applies_(r, kind)) return;
    total++;
    const t = tally_(r, r.all, kind, j);
    if (!t) return;
    done++; att += t.att; base += t.base;
  });
  return { att: att, base: base, done: done, total: total, rate: base ? att / base : null };
}

// 순원 가운데 한 명이라도 순모임에 체크되어 있으면 그 순은 모임을 한 것으로 센다.
function groupMet_(g, j) {
  return g.members.some(function (m) { return m.g[j] === '1'; });
}

function meetings_(regions, j) {
  let did = 0, total = 0;
  regions.forEach(function (r) {
    if (!r.hasGroup) return;
    r.groups.forEach(function (g) {
      total++;
      if (groupMet_(g, j)) did++;
    });
  });
  return { did: did, total: total };
}

// 지난주와 비교할 때는 두 주 모두 기록한 지역끼리만 센다(미기록 지역 때문에 줄어 보이지 않도록)
function delta_(D, kind, j) {
  if (j <= 0) return '';
  let now = 0, prev = 0, n = 0, total = 0;
  D.regions.forEach(function (r) {
    if (!applies_(r, kind)) return;
    total++;
    const a = tally_(r, r.all, kind, j), b = tally_(r, r.all, kind, j - 1);
    if (!a || !b) return;
    now += a.att; prev += b.att; n++;
  });
  if (!n) return '';
  const note = n < total ? ' (' + n + '개 지역 비교)' : '';
  return diffText_(now - prev) + note;
}

function diffText_(d) {
  if (d > 0) return '지난주보다 ▲ ' + d + '명';
  if (d < 0) return '지난주보다 ▼ ' + (-d) + '명';
  return '지난주와 같음';
}

/** 0주부터 j주까지 기록된 주의 평균 출석 인원 */
function avgCount_(r, members, kind, j) {
  let att = 0, weeks = 0;
  for (let k = 0; k <= j; k++) {
    const t = tally_(r, members, kind, k);
    if (!t) continue;
    att += t.att; weeks++;
  }
  return weeks ? Math.round(att / weeks) : null;
}

function pct_(x) { return x == null ? '—' : Math.round(x * 100) + '%'; }

function countText_(t) { return t ? t.att + '명 · ' + pct_(t.rate) : '미기록'; }

function rateColor_(rate) {
  if (rate == null) return '#8c847a';
  return rate >= 0.75 ? '#2E7D5B' : rate >= 0.5 ? '#A87B22' : '#B04936';
}

function regionNo_(r, i) {
  const m = String(r.title).match(/(\d+)\s*지역/);
  return m ? m[1] + '지역' : r.title || (i + 1) + '지역';
}

function leadOf_(g) {
  for (let k = 0; k < g.members.length; k++) if (g.members[k].lead) return g.members[k].name;
  return '';
}

function weekLabel_(D, j) {
  const k = String(D.weekKeys[j]);
  return k.slice(0, 4) + '년 ' + Number(k.slice(4, 6)) + '월 ' + Number(k.slice(6, 8)) + '일';
}

/* ------------------------------ 문서 ------------------------------ */

const COLOR = { brand: '#2B3A55', line: '#d9d4cc', soft: '#f4f2ee', muted: '#8c847a', ink: '#1c1a17' };

function reportFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('REPORT_FOLDER_ID');
  if (id) {
    try {
      const f = DriveApp.getFolderById(id);
      if (!f.isTrashed()) return f;
    } catch (err) {
      // 폴더가 지워졌으면 새로 만든다
    }
  }
  const folder = DriveApp.createFolder(REPORT.FOLDER_NAME);
  props.setProperty('REPORT_FOLDER_ID', folder.getId());
  return folder;
}

/** 같은 주일의 리포트가 폴더에 있으면 그 문서를, 없으면 새 문서를 연다 */
function reportDoc_(D, j) {
  const name = D.title + ' 주간 출석 리포트 (' +
    String(D.weekKeys[j]).replace(/^(\d{4})(\d{2})(\d{2})$/, '$1-$2-$3') + ')';
  const folder = reportFolder_();
  const files = folder.getFilesByName(name);
  while (files.hasNext()) {
    const f = files.next();
    if (!f.isTrashed()) return DocumentApp.openById(f.getId());
  }
  const doc = DocumentApp.create(name);
  DriveApp.getFileById(doc.getId()).moveTo(folder);
  return doc;
}

function writeReport_(doc, D, j) {
  const body = doc.getBody();
  body.clear();
  body.setMarginTop(48).setMarginBottom(48).setMarginLeft(54).setMarginRight(54);

  const title = body.getParagraphs()[0];
  title.setText(D.title + ' 주간 출석 리포트').setHeading(DocumentApp.ParagraphHeading.TITLE);
  title.editAsText().setFontSize(22).setBold(true).setForegroundColor(COLOR.brand);

  const sub = body.appendParagraph(weekLabel_(D, j) + ' 주일 · ' +
    Utilities.formatDate(D.now, D.tz, 'M월 d일 HH:mm') + ' 작성');
  sub.editAsText().setFontSize(10).setForegroundColor(COLOR.muted);

  writeSummary_(body, D, j);
  writeRegions_(body, D, j);
  writeTrend_(body, D, j);
  writeGroups_(body, D, j);
  writeChecks_(body, D, j);

  const foot = body.appendParagraph(
    '시트에 있는 정보만 사용했습니다. 한 지역에서 아무도 체크하지 않은 주는 아직 기록하지 않은 주로 보고 ' +
    '출석률 계산에서 뺐습니다. 지역장 행은 순 안의 같은 이름과 겹치므로 셈에서 뺐습니다. ' +
    '순모임은 순원 한 명이라도 체크되면 그 순이 모인 것으로 셉니다.');
  foot.editAsText().setFontSize(8.5).setForegroundColor(COLOR.muted);
  foot.setSpacingBefore(14);
}

function heading_(body, text) {
  const p = body.appendParagraph(text).setHeading(DocumentApp.ParagraphHeading.HEADING2);
  p.editAsText().setFontSize(14).setBold(true).setForegroundColor(COLOR.brand);
  p.setSpacingBefore(18).setSpacingAfter(6);
  return p;
}

/**
 * 표를 붙이고 모양을 다듬는다.
 * widths: 열 너비(pt), colors: { 'row,col': 글자색 }, boldRows: 굵게 할 본문 행 번호(0부터)
 */
function table_(body, header, rows, opt) {
  opt = opt || {};
  const cells = (header ? [header] : []).concat(rows).map(function (r) {
    return r.map(function (v) { return String(v == null ? '' : v); });
  });
  const t = body.appendTable(cells);
  t.setBorderColor(COLOR.line).setBorderWidth(0.5);
  for (let r = 0; r < t.getNumRows(); r++) {
    const row = t.getRow(r);
    for (let c = 0; c < row.getNumCells(); c++) {
      const cell = row.getCell(c);
      cell.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(6).setPaddingRight(6);
      const tx = cell.editAsText();
      tx.setFontSize(opt.fontSize || 10).setForegroundColor(COLOR.ink).setBold(false);
      const bodyRow = header ? r - 1 : r;
      if (header && r === 0) {
        cell.setBackgroundColor(COLOR.brand);
        tx.setForegroundColor('#ffffff').setBold(true);
        continue;
      }
      if (opt.labelCol && c === 0) {
        cell.setBackgroundColor(COLOR.soft);
        tx.setBold(true);
      }
      if (opt.boldRows && opt.boldRows.indexOf(bodyRow) !== -1) {
        cell.setBackgroundColor(COLOR.soft);
        tx.setBold(true);
      }
      const color = opt.colors && opt.colors[bodyRow + ',' + c];
      if (color) tx.setForegroundColor(color);
    }
  }
  if (opt.widths) opt.widths.forEach(function (w, c) { if (w) t.setColumnWidth(c, w); });
  return t;
}

function writeSummary_(body, D, j) {
  heading_(body, '1. 이번 주 한눈에');
  const w = weekStats_(D, 'w', j), g = weekStats_(D, 'g', j);
  const cm = meetings_(D.regions, j);
  const dw = delta_(D, 'w', j), dg = delta_(D, 'g', j);

  const rows = [
    ['순모임 한 순', cm.total ? cm.did + '순 / ' + cm.total + '순 (' + pct_(cm.did / cm.total) + ')' : '순모임 칸 없음'],
    ['주일예배 출석', w.done ? w.att + '명 · 출석률 ' + pct_(w.rate) + (dw ? '  |  ' + dw : '') : '미기록'],
    ['순모임 참석', g.done ? g.att + '명 · 참석률 ' + pct_(g.rate) + (dg ? '  |  ' + dg : '') : '미기록'],
    ['기록 완료 지역', '주일예배 ' + w.done + ' / ' + w.total + ' 지역 · 순모임 ' + g.done + ' / ' + g.total + ' 지역']
  ];
  const colors = {};
  if (cm.total) colors['0,1'] = rateColor_(cm.did / cm.total);
  if (w.done) colors['1,1'] = rateColor_(w.rate);
  if (g.done) colors['2,1'] = rateColor_(g.rate);
  if (w.done < w.total || g.done < g.total) colors['3,1'] = '#B04936';
  table_(body, null, rows, { labelCol: true, widths: [110], fontSize: 11, colors: colors });
}

function writeRegions_(body, D, j) {
  heading_(body, '2. 지역별 현황');
  const rows = [];
  const colors = {};
  D.regions.forEach(function (r, i) {
    const tw = tally_(r, r.all, 'w', j), tg = tally_(r, r.all, 'g', j);
    const pw = tally_(r, r.all, 'w', j - 1), pg = tally_(r, r.all, 'g', j - 1);
    const cm = meetings_([r], j);
    const arrow = function (a, b) {
      if (!a || !b) return '';
      const d = a.att - b.att;
      return d > 0 ? ' ▲' + d : d < 0 ? ' ▼' + (-d) : '';
    };
    rows.push([
      regionNo_(r, i),
      r.leader || '미지정',
      r.all.length + '명',
      countText_(tw) + arrow(tw, pw),
      r.hasGroup ? countText_(tg) + arrow(tg, pg) : '칸 없음',
      !r.hasGroup ? '—' : r.rec.g[j] ? cm.did + ' / ' + cm.total : '미기록'
    ]);
    const n = rows.length - 1;
    colors[n + ',3'] = rateColor_(tw && tw.rate);
    colors[n + ',4'] = rateColor_(tg && tg.rate);
  });

  const w = weekStats_(D, 'w', j), g = weekStats_(D, 'g', j), cm = meetings_(D.regions, j);
  const people = D.regions.reduce(function (s, r) { return s + r.all.length; }, 0);
  rows.push([
    '합계', '', people + '명',
    w.done ? w.att + '명 · ' + pct_(w.rate) : '미기록',
    g.done ? g.att + '명 · ' + pct_(g.rate) : '미기록',
    cm.total ? cm.did + ' / ' + cm.total : '—'
  ]);
  table_(body, ['지역', '지역장', '인원', '주일예배', '순모임', '모인 순'], rows, {
    widths: [52, 62, 44, 104, 104, 54],
    boldRows: [rows.length - 1],
    colors: colors
  });
  note_(body, '▲▼는 지난주와 비교한 인원 차이입니다. 출석률 75% 이상은 초록, 50% 이상은 노랑, 그 아래는 붉은색입니다.');
}

function writeTrend_(body, D, j) {
  heading_(body, '3. 최근 ' + REPORT.TREND_WEEKS + '주 흐름');
  const rows = [];
  const colors = {};
  for (let k = Math.max(0, j - REPORT.TREND_WEEKS + 1); k <= j; k++) {
    const w = weekStats_(D, 'w', k), g = weekStats_(D, 'g', k), cm = meetings_(D.regions, k);
    rows.push([
      D.weeks[k].month + '월 ' + D.weeks[k].day + '일',
      w.done ? w.att + '명 · ' + pct_(w.rate) : '미기록',
      g.done ? g.att + '명 · ' + pct_(g.rate) : '미기록',
      cm.total ? cm.did + ' / ' + cm.total : '—',
      w.done + ' / ' + w.total
    ]);
    const n = rows.length - 1;
    colors[n + ',1'] = rateColor_(w.rate);
    colors[n + ',2'] = rateColor_(g.rate);
  }
  table_(body, ['주일', '주일예배', '순모임', '모인 순', '기록 지역'], rows, {
    widths: [70, 110, 110, 70, 70],
    boldRows: [rows.length - 1],
    colors: colors
  });
  note_(body, '주별 합계에는 그 주를 기록한 지역만 들어갑니다. 기록 지역 수가 다르면 인원을 바로 비교하기 어렵습니다.');
}

function writeGroups_(body, D, j) {
  heading_(body, '4. 순별 현황');
  D.regions.forEach(function (r, i) {
    const h = body.appendParagraph(regionNo_(r, i) + '  ·  지역장 ' + (r.leader || '미지정') + '  ·  ' + r.all.length + '명')
      .setHeading(DocumentApp.ParagraphHeading.HEADING3);
    h.editAsText().setFontSize(11.5).setBold(true).setForegroundColor(COLOR.ink);
    h.setSpacingBefore(10).setSpacingAfter(4);

    const rows = [];
    const colors = {};
    r.groups.forEach(function (g) {
      const tw = tally_(r, g.members, 'w', j), tg = tally_(r, g.members, 'g', j);
      const aw = avgCount_(r, g.members, 'w', j), ag = avgCount_(r, g.members, 'g', j);
      rows.push([
        g.name,
        leadOf_(g) || '—',
        g.members.length + '명',
        tw ? tw.att + ' / ' + tw.base : '미기록',
        !r.hasGroup ? '칸 없음' : tg ? tg.att + ' / ' + tg.base : '미기록',
        '예배 ' + (aw == null ? '—' : aw + '명') + (r.hasGroup ? ' · 순모임 ' + (ag == null ? '—' : ag + '명') : '')
      ]);
      const n = rows.length - 1;
      colors[n + ',3'] = rateColor_(tw && tw.rate);
      colors[n + ',4'] = rateColor_(tg && tg.rate);
    });
    table_(body, ['순', '순장', '인원', '주일예배', '순모임', '주 평균'], rows, {
      widths: [50, 58, 40, 66, 66, 140],
      fontSize: 9.5,
      colors: colors
    });
  });
}

function writeChecks_(body, D, j) {
  heading_(body, '5. 기록 점검');
  const items = [];

  ['w', 'g'].forEach(function (kind) {
    const miss = [];
    D.regions.forEach(function (r, i) {
      if (!applies_(r, kind) || r.rec[kind][j]) return;
      miss.push(regionNo_(r, i) + '(지역장 ' + (r.leader || '미지정') + (r.leaderPhone ? ' ' + r.leaderPhone : '') + ')');
    });
    const what = kind === 'w' ? '주일예배' : '순모임';
    items.push(miss.length
      ? what + ' 미기록 ' + miss.length + '개 지역: ' + miss.join(', ')
      : what + '는 모든 지역이 기록했습니다.');
  });

  const noGroup = D.regions.filter(function (r) { return !r.hasGroup; })
    .map(function (r, i) { return regionNo_(r, i); });
  if (noGroup.length) items.push('순모임 칸이 없는 지역: ' + noGroup.join(', '));

  const noLead = [];
  D.regions.forEach(function (r, i) {
    r.groups.forEach(function (g) { if (!leadOf_(g)) noLead.push(regionNo_(r, i) + ' ' + g.name); });
  });
  if (noLead.length) items.push('순장이 굵은 글씨로 표시되지 않은 순: ' + noLead.join(', '));

  D.regions.forEach(function (r, i) {
    (r.warnings || []).forEach(function (w) { items.push('시트 머리글 확인 필요(' + regionNo_(r, i) + '): ' + w); });
  });
  D.errors.forEach(function (e) { items.push('읽지 못한 탭: ' + e); });

  items.forEach(function (t) {
    const li = body.appendListItem(t).setGlyphType(DocumentApp.GlyphType.BULLET);
    li.editAsText().setFontSize(10).setForegroundColor(COLOR.ink).setBold(false);
  });
}

function note_(body, text) {
  const p = body.appendParagraph(text);
  p.editAsText().setFontSize(8.5).setForegroundColor(COLOR.muted).setBold(false);
  p.setSpacingBefore(3);
  return p;
}

/* ------------------------------ 메일 ------------------------------ */

function mailReport_(D, j, doc, url) {
  const to = String(prop_('REPORT_EMAILS') || '').split(/[,\s;]+/).filter(function (s) { return /@/.test(s); });
  if (!to.length) return;

  try {
    DriveApp.getFileById(doc.getId()).addViewers(to);
  } catch (err) {
    Logger.log('보기 권한을 주지 못했습니다: %s', err.message);
  }

  const w = weekStats_(D, 'w', j), g = weekStats_(D, 'g', j), cm = meetings_(D.regions, j);
  const lines = [
    '순모임 한 순: ' + (cm.total ? cm.did + '순 / ' + cm.total + '순' : '순모임 칸 없음'),
    '주일예배 출석: ' + (w.done ? w.att + '명 (' + pct_(w.rate) + ')' : '미기록'),
    '순모임 참석: ' + (g.done ? g.att + '명 (' + pct_(g.rate) + ')' : '미기록'),
    '기록 완료: 주일예배 ' + w.done + '/' + w.total + ' 지역, 순모임 ' + g.done + '/' + g.total + ' 지역'
  ];
  const esc = function (s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };
  MailApp.sendEmail({
    to: to.join(','),
    subject: doc.getName(),
    body: weekLabel_(D, j) + ' 주일 출석 리포트\n\n' + lines.join('\n') + '\n\n전체 리포트: ' + url,
    htmlBody: '<p><b>' + esc(weekLabel_(D, j)) + ' 주일 출석 리포트</b></p><p>' +
      lines.map(esc).join('<br>') + '</p><p><a href="' + esc(url) + '">전체 리포트 열기</a></p>'
  });
}
