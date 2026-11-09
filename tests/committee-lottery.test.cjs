const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const vm = require('node:vm');
const lottery = require('../apps-script/CommitteeLottery.js');

const roles = lottery.LOTTERY_SETTINGS.committees;
const classes = lottery.LOTTERY_SETTINGS.classes;
const header = ['タイムスタンプ', 'メールアドレス', '名前', 'クラス', '出席番号',
  '【質問2】役職希望', '【質問3】第1希望', '【質問4】第2希望', '【質問5】出演',
  '【質問6】出演内容', '【質問1】参加', '【質問7】準備参加', '【質問10】自由記述', '重複立候補不可チェック'];
const opHeader = ['クラス', '番号', '11月21日不参加特別理由', '打ち合わせ不参加特別理由', '重複不可委員就任済み', 'メモ'];
const digest = bytes => Array.from(crypto.createHash('sha256').update(Buffer.from(bytes)).digest());
const clone = value => structuredClone(value);
function freeze(value) {
  if (value && typeof value === 'object') { Object.values(value).forEach(freeze); Object.freeze(value); }
  return value;
}
function response(className, number, first, options = {}) {
  return ['fixture-time', 'fixture@example.invalid', options.name || `架空${className}${number}`, className, number,
    options.chair === false ? '希望しない' : '希望する', first, options.second || '特になし',
    'まだ未定', '架空の記述', options.orientation || '参加できます',
    options.meetings || '予定によっては参加できない日があります', '', options.oldN || ''];
}
function balanced() {
  return [header.slice(), ...classes.flatMap(c => roles.map((role, i) => response(c, i + 1, role)))];
}
function operations(rows = []) { return [opHeader.slice(), ...rows]; }
function engine(rows, ops = operations(), options = {}) {
  return lottery.runLotteryEngine_(rows, ops, { digest, ...options });
}

test('SHA-256 chain and UTF-8 encoding agree with the independent Node crypto implementation', () => {
  const events = [];
  const random = lottery.createRandom_('00'.repeat(32), (type, detail) => events.push(detail), digest);
  const first = digest(Array(32).fill(0));
  const second = digest(first);
  random.nextInt(0x100000000, 'one');
  random.nextInt(0x100000000, 'two');
  assert.equal(events[0].stateAfter, Buffer.from(first).toString('hex'));
  assert.equal(events[1].stateBefore, events[0].stateAfter);
  assert.equal(events[1].stateAfter, Buffer.from(second).toString('hex'));
  assert.equal(events[0].word, Buffer.from(first).readUInt32BE());
  const text = '委員抽選😀\ud800';
  assert.equal(lottery.hashText_(text, digest), crypto.createHash('sha256').update(text).digest('hex'));
});

test('presentation changes preserve the version 1.0 deterministic audit trace', () => {
  const data = balanced();
  const extra = response('A組', 90, '特に希望なし', { chair: false, second: roles[1], meetings: '参加できます' });
  data.push(extra);
  // The version 1.0 reference hash belongs to this fixture seed, independent of the operational seed.
  const referenceSeed = '72f7c0a2629261e23b8f7627a22233d17c72740d6305acdc9716789bf70ec83c';
  const result = engine(data, operations(), { seed: referenceSeed });
  assert.equal(crypto.createHash('sha256').update(JSON.stringify(result.audit)).digest('hex'),
    '41bb13fc68eb788fdf6908b94eea9a9ad3cf6265bf30293abdac46140538a516');
});

test('integer conversion rejects the modulo-biased tail and records the discarded word', () => {
  const words = [0xffffffff, 8];
  const events = [];
  const fakeDigest = () => { const b = Buffer.alloc(32); b.writeUInt32BE(words.shift()); return b; };
  const random = lottery.createRandom_('01'.repeat(32), (type, detail) => events.push(detail), fakeDigest);
  assert.equal(random.nextInt(3, 'fixture'), 2);
  assert.equal(random.position(), 2);
  assert.deepEqual(events.map(e => e.accepted), [false, true]);
  assert.equal(events[0].limit, 4294967295);
});

test('same seed and frozen inputs reproduce the full decision trace without mutating inputs', () => {
  const data = balanced();
  data.push(response('A組', 90, roles[1], { second: roles[3] }));
  const before = JSON.stringify(data);
  const frozenData = freeze(clone(data));
  const frozenOps = freeze(operations());
  const first = engine(frozenData, frozenOps);
  const second = engine(frozenData, frozenOps);
  assert.deepEqual(first, second);
  assert.equal(JSON.stringify(data), before);
  assert.equal(first.status, 'ACCEPTED');
  assert.equal(first.acceptedAttempt, 1);
});

test('latest response wins; operations match independently; old N flags and performance do not exclude', () => {
  const data = balanced();
  data.find(r => r[3] === 'B組' && r[4] === 2)[10] = '参加できません';
  const c3 = data.find(r => r[3] === 'C組' && r[4] === 3);
  c3[10] = '参加できません'; c3[11] = '参加できません';
  data.find(r => r[3] === 'D組' && r[4] === 4)[11] = '参加できません';
  data.find(r => r[3] === 'E組' && r[4] === 5)[11] = '未知の回答でも参加不可とは判定しない';
  data.find(r => r[3] === 'E組' && r[4] === 1)[13] = '旧Nの印';
  data.push(response('Ａ組', '０１番', roles[2], { orientation: '参加できません', meetings: '参加できません' }));
  const ops = operations([
    ['A組', 1, '公欠', '', '確認済みの委員', 'fixture'],
    ['B組', '２', '公欠', '', '', 'fixture'],
    ['C組', 3, '', '公欠', '', 'fixture']
  ]);
  const prepared = lottery.prepareLotteryInput_(freeze(data), freeze(ops));
  const people = new Map(prepared.candidates.map(p => [p.key, p]));
  assert.equal(prepared.responseCount, 26);
  assert.equal(prepared.candidates.length, 25);
  assert.equal(prepared.duplicateCount, 1);
  assert.equal(people.get('A組:1').firstPreference, roles[2]);
  assert.deepEqual(people.get('A組:1').exclusionReasons, ['OTHER_COMMITTEE', 'MEETINGS_UNAVAILABLE']);
  assert.deepEqual(people.get('A組:1').appliedExceptions, ['ORIENTATION_EXCUSED']);
  assert.deepEqual(people.get('B組:2').exclusionReasons, []);
  assert.deepEqual(people.get('C組:3').exclusionReasons, ['ORIENTATION_UNAVAILABLE']);
  assert.deepEqual(people.get('C組:3').appliedExceptions, ['MEETINGS_EXCUSED']);
  assert.deepEqual(people.get('D組:4').exclusionReasons, ['MEETINGS_UNAVAILABLE']);
  assert.deepEqual(people.get('E組:5').exclusionReasons, []);
  assert.deepEqual(people.get('E組:1').exclusionReasons, []);
});

test('discarded old response contents do not invalidate a valid final response', () => {
  const data = balanced();
  data[1][6] = '古い回答の未対応値';
  data.push(response('A組', 1, roles[0]));
  assert.equal(lottery.prepareLotteryInput_(data, operations()).candidates.length, 25);
});

test('only first-round losers compete for second preference vacancies; capacity and uniqueness hold', () => {
  const data = balanced().filter(r => !(r[3] === 'E組' && r[4] === 4));
  data.filter(r => r[6] === roles[1]).forEach(r => { r[7] = roles[3]; });
  data.push(response('A組', 90, roles[1], { second: roles[3] }));
  const result = engine(data);
  const attempt = result.attempts[0];
  assert.equal(result.status, 'ACCEPTED');
  assert.equal(attempt.assignments.length, 25);
  assert.equal(new Set(attempt.assignments.map(a => a.key)).size, 25);
  assert.equal(attempt.assignments.filter(a => a.stage === 2).length, 1);
  assert.equal(attempt.assignments.find(a => a.stage === 2).committee, roles[3]);
  assert.ok(Object.values(attempt.committeeCounts).every(n => n <= 5));
  for (const row of data.slice(1).filter(r => r[6] === roles[3])) {
    assert.ok(attempt.assignments.some(a => a.key === `${row[3]}:${Number(row[4])}` && a.stage === 1 && a.committee === roles[3]));
  }
});

test('empty second preferences leave vacancies and do not cause off-preference assignment', () => {
  const data = balanced().filter(r => !(r[3] === 'E組' && r[4] === 4));
  data.push(response('A組', 90, roles[1]));
  const result = engine(data);
  assert.equal(result.status, 'ACCEPTED');
  assert.equal(result.assignments.length, 24);
  assert.equal(result.attempts[0].committeeCounts[roles[3]], 4);
  assert.ok(result.assignments.every(a => a.stage === 1));
  assert.ok(result.audit.some(e => e.type === 'TRIAL_CANDIDATE_STATUS' && e.detail.reason === 'NO_SECOND_PREFERENCE'));
});

test('audit distinguishes full second-choice committees from second-choice lottery losses for every candidate', () => {
  const full = balanced();
  full.push(response('A組', 90, roles[1], { second: roles[0] }));
  full.filter(r => r[6] === roles[1]).forEach(r => { r[7] = roles[0]; });
  const fullResult = engine(full);
  assert.ok(fullResult.audit.some(e => e.type === 'TRIAL_CANDIDATE_STATUS' && e.detail.reason === 'SECOND_PREFERENCE_FULL'));
  const vacancy = balanced().filter(r => !(r[3] === 'E組' && r[4] === 4));
  vacancy.push(response('A組', 90, roles[1]), response('A組', 91, roles[1]));
  vacancy.filter(r => r[6] === roles[1]).forEach(r => { r[7] = roles[3]; });
  const result = engine(vacancy);
  assert.ok(result.audit.some(e => e.type === 'TRIAL_CANDIDATE_STATUS' && e.detail.reason === 'SECOND_PREFERENCE_DRAW_LOSS'));
  for (const attempt of result.attempts) {
    const statuses = result.audit.filter(e => e.type === 'TRIAL_CANDIDATE_STATUS' && e.detail.attempt === attempt.number);
    assert.equal(statuses.length, result.prepared.eligible.length);
    assert.equal(new Set(statuses.map(e => e.detail.key)).size, statuses.length);
    assert.ok(statuses.every(e => e.detail.attemptAccepted === attempt.accepted));
  }
});

function oneLeaderCommittee() {
  const data = balanced();
  for (const r of data.slice(1)) {
    r[5] = r[6] === roles[1] ? '希望する' : '希望しない';
    if (r[6] === roles[1]) r[7] = roles[0];
  }
  return data;
}

test('rejects all six attempts when leaders occupy only one committee; no failed winners are adopted', () => {
  const result = engine(oneLeaderCommittee());
  assert.equal(result.precheck.reasons.length, 0);
  assert.equal(result.status, 'ATTEMPT_LIMIT_REACHED');
  assert.deepEqual(result.attempts.map(a => a.number), [1, 2, 3, 4, 5, 6]);
  assert.ok(result.attempts.every(a => !a.accepted && a.leadershipCommittees.length === 1));
  assert.deepEqual(result.assignments, []);
  assert.deepEqual(result.leadershipCandidates, []);
  assert.equal(result.acceptedAttempt, null);
});

test('no-first-preference assignment is fixed once and the random stream continues through retries', () => {
  const data = oneLeaderCommittee();
  data.push(response('A組', 90, '特に希望なし', { chair: false }));
  const result = engine(data);
  assert.equal(result.status, 'ATTEMPT_LIMIT_REACHED');
  assert.equal(result.audit.filter(e => e.type === 'FIXED_FIRST_PREFERENCE').length, 1);
  assert.equal(result.firstPreferenceAssignments.length, 1);
  const randoms = result.audit.filter(e => e.type === 'RANDOM').map(e => e.detail);
  assert.deepEqual(randoms.map(e => e.position), Array.from({ length: randoms.length }, (_, i) => i + 1));
  for (let i = 1; i < randoms.length; i++) assert.equal(randoms[i].stateBefore, randoms[i - 1].stateAfter);
  const starts = result.audit.filter(e => e.type === 'ATTEMPT_START');
  for (let i = 1; i < starts.length; i++) {
    assert.equal(starts[i].detail.randomPosition, result.attempts[i - 1].randomPosition);
    assert.equal(starts[i].detail.randomState, result.attempts[i - 1].randomState);
  }
  for (const attempt of result.attempts) {
    for (const a of attempt.assignments.filter(a => a.key === 'A組:90' && a.stage === 1)) {
      assert.equal(a.committee, result.firstPreferenceAssignments[0].committee);
    }
  }
});

test('rechecks all conditions and stops at the first valid retry without resetting the stream', () => {
  const data = [header.slice()];
  for (const [i, role] of roles.entries()) {
    if (role === roles[1]) continue;
    for (const c of classes.slice(1)) data.push(response(c, 100 + i * 10, role, { chair: i === 0 && c === 'B組' }));
    data.push(response('B組', 101 + i * 10, role, { chair: false }));
  }
  for (const [c, n] of [['A組', 1], ['A組', 2], ['B組', 9], ['C組', 9], ['D組', 9], ['E組', 9]]) {
    data.push(response(c, n, roles[1]));
  }
  const words = [0, 0, 0, 0, 0, 5, 4, 3, 2, 1];
  let calls = 0;
  const controlledDigest = () => { const b = Buffer.alloc(32); b.writeUInt32BE(words[calls] ?? 0); b.writeUInt32BE(++calls, 28); return b; };
  const result = engine(data, operations(), { digest: controlledDigest });
  assert.equal(result.status, 'ACCEPTED');
  assert.equal(result.acceptedAttempt, 2);
  assert.equal(result.attempts.length, 2);
  assert.equal(result.attempts[0].classCounts['A組'], 1);
  assert.equal(result.attempts[1].classCounts['A組'], 2);
  assert.equal(result.attempts[0].randomPosition, 5);
  assert.equal(result.randomPosition, 10);
});

test('proven impossible class coverage and missing leadership candidates stop before random draws', () => {
  const narrow = balanced();
  narrow.slice(1).forEach(r => { r[6] = roles[0]; r[7] = '特になし'; });
  const result = engine(narrow);
  assert.equal(result.status, 'PRECHECK_FAILED');
  assert.equal(result.precheck.optimisticClassCoverage, 5);
  assert.equal(result.attempts.length, 0);
  assert.equal(result.randomPosition, 0);
  const noLeaders = balanced(); noLeaders.slice(1).forEach(r => { r[5] = '希望しない'; });
  assert.ok(engine(noLeaders).precheck.reasons.includes('役職希望者が2人未満'));
  const missingClass = balanced().filter(r => r[3] !== 'E組');
  assert.ok(engine(missingClass).precheck.reasons.includes('E組の対象者が2人未満'));
});

test('performance, contact details, names and submission time cannot change the selected identities', () => {
  const first = balanced(); first.push(response('A組', 90, roles[1]));
  const second = clone(first);
  for (const r of second.slice(1)) {
    r[0] = 'different time'; r[1] = 'other@example.invalid'; r[2] = '別の架空名';
    r[8] = '出演確定'; r[9] = '別の架空記述'; r[12] = '主観的評価に使わない';
  }
  assert.deepEqual(engine(first).assignments, engine(second).assignments);
});

test('ambiguous headers, invalid identities, duplicate operations and unmatched operations fail explicitly', () => {
  const bad = balanced(); bad[1][4] = '不明';
  assert.throws(() => engine(bad), /クラスまたは出席番号/);
  assert.throws(() => engine(balanced(), operations([['A組', 1, '', '', '', ''], ['A組', '01番', '', '', '', '']])), /重複/);
  assert.throws(() => engine(balanced(), operations([['A組', 999, '', '', '', '']])), /対応する応募者/);
  assert.throws(() => engine(balanced(), operations([['A組', 1, '未確認', '', '', '']])), /公欠/);
  const ambiguous = balanced(); ambiguous[0].push('クラス'); ambiguous.slice(1).forEach(r => r.push('A組'));
  assert.throws(() => engine(ambiguous), /一意に特定/);
  assert.throws(() => engine(balanced(), operations(), { seed: 'bad' }), /64桁/);
});

test('header-based input mapping survives reordered columns and ignores the obsolete exclusion field', () => {
  const original = balanced(); original.push(response('A組', 90, roles[1]));
  const order = [11, 4, 7, 0, 10, 3, 5, 2, 9, 8, 6, 1, 12, 13];
  const reordered = original.map(row => order.map(i => row[i]));
  const ops = operations([['A組', 1, '', '', '架空の委員', 'fixture']]);
  const opOrder = [4, 2, 0, 5, 1, 3];
  const movedOps = ops.map(row => opOrder.map(i => row[i]));
  assert.deepEqual(engine(original, ops), engine(reordered, movedOps));
});

class FakeRange {
  constructor(sheet, row, column, height, width) { Object.assign(this, { sheet, row, column, height, width }); }
  getValues() {
    return Array.from({ length: this.height }, (_, y) => Array.from({ length: this.width }, (_, x) =>
      clone(this.sheet.values[this.row - 1 + y]?.[this.column - 1 + x] ?? '')));
  }
  getA1Notation() { return `A1:N${this.height}`; }
  getNote() { return this.sheet.notes[`${this.row}:${this.column}`] || ''; }
  setNote(value) { this.sheet.writable(); this.sheet.notes[`${this.row}:${this.column}`] = value; return this; }
  setNumberFormat() { this.sheet.writable(); return this; }
  setFontWeight() { this.sheet.writable(); return this; }
  setBackground() { this.sheet.writable(); return this; }
  setFontColor() { this.sheet.writable(); return this; }
  setWrapStrategy() { this.sheet.writable(); return this; }
  createFilter() { this.sheet.writable(); this.sheet.filter = true; return this; }
  setValues(values) {
    this.sheet.writable();
    if (++this.sheet.owner.writes === this.sheet.owner.failWriteAt) throw Error('simulated write failure');
    assert.equal(values.length, this.height);
    for (const [y, row] of values.entries()) {
      assert.equal(row.length, this.width);
      for (const [x, value] of row.entries()) {
        assert.ok(typeof value !== 'string' || !value.startsWith('='), 'formula injection');
        this.sheet.values[this.row - 1 + y] ||= [];
        this.sheet.values[this.row - 1 + y][this.column - 1 + x] = value;
      }
    }
    return this;
  }
  setValue(value) { this.height = 1; this.width = 1; return this.setValues([[value]]); }
}
class FakeSheet {
  constructor(owner, name, values, input = false) {
    Object.assign(this, { owner, name, values: clone(values), input, id: ++owner.nextId, rows: 1000, columns: 26, notes: {} });
  }
  getName() { return this.name; }
  getSheetId() { return this.id; }
  getDataRange() { return new FakeRange(this, 1, 1, this.values.length || 1, Math.max(1, ...this.values.map(r => r.length))); }
  getRange(row, column, height = 1, width = 1) { return new FakeRange(this, row, column, height, width); }
  getMaxRows() { return this.rows; }
  getMaxColumns() { return this.columns; }
  writable() { assert.ok(!this.input, 'input sheet must never be mutated'); }
  insertRowsAfter(after, count) { this.writable(); this.rows += count; }
  insertColumnsAfter(after, count) { this.writable(); this.columns += count; }
  setFrozenRows(count) { this.writable(); this.frozenRows = count; }
  setColumnWidth() { this.writable(); }
  setRowHeight() { this.writable(); }
  getLastRow() {
    for (let i = this.values.length - 1; i >= 0; i--) if (this.values[i].some(v => v !== '' && v != null)) return i + 1;
    return 0;
  }
  hideSheet() { this.writable(); this.hidden = true; }
  showSheet() { this.writable(); this.hidden = false; }
  setName(name) {
    this.writable();
    if (++this.owner.renames === this.owner.failRenameAt) throw Error('simulated rename failure');
    assert.ok(!this.owner.getSheetByName(name) || this.owner.getSheetByName(name) === this);
    this.name = name; return this;
  }
}
class FakeSpreadsheet {
  constructor(data = balanced(), ops = operations()) {
    this.nextId = 0; this.writes = 0; this.renames = 0;
    this.sheets = [new FakeSheet(this, 'フォームの回答 1', data, true), new FakeSheet(this, '運営確認', ops, true)];
  }
  getId() { return 'fixture-spreadsheet'; }
  getSpreadsheetTimeZone() { return 'Asia/Tokyo'; }
  getSheets() { return this.sheets.slice(); }
  getSheetByName(name) { return this.sheets.find(s => s.name === name) || null; }
  insertSheet(name) {
    assert.ok(!this.getSheetByName(name), 'must not overwrite existing sheets');
    const sheet = new FakeSheet(this, name, []); this.sheets.push(sheet); return sheet;
  }
  deleteSheet(sheet) { sheet.writable(); this.sheets.splice(this.sheets.indexOf(sheet), 1); sheet.deleted = true; }
  setActiveSheet(sheet) { sheet.writable(); this.activeSheet = sheet; }
  moveActiveSheet(position) {
    this.activeSheet.writable();
    this.sheets.splice(this.sheets.indexOf(this.activeSheet), 1);
    this.sheets.splice(position - 1, 0, this.activeSheet);
  }
}
const runOptions = { seed: lottery.LOTTERY_SEED, digest, runId: 'fixture-run', startedAt: '2026-01-01T00:00:00.000Z' };

test('operations headers may have a title above them; ambiguous or missing input sheets never cause writes', () => {
  const ops = [['運営追記'], [], ...operations([['A組', 1, '', '', '架空の委員', 'fixture']])];
  const spreadsheet = new FakeSpreadsheet(balanced(), ops);
  const captured = lottery.readInputSheets_(spreadsheet);
  assert.equal(captured.operations.headerRow, 3);
  const prepared = lottery.prepareLotteryInput_(captured.responses.data, captured.operations.data, captured);
  assert.equal(prepared.candidates.find(p => p.key === 'A組:1').operationRow, 4);
  assert.equal(prepared.eligible.length, 24);
  const ambiguous = new FakeSpreadsheet();
  ambiguous.sheets.push(new FakeSheet(ambiguous, '別の回答', balanced(), true));
  assert.throws(() => lottery.executeLotteryOnSpreadsheet_(ambiguous, runOptions), /一意に特定/);
  assert.equal(ambiguous.writes, 0);
  const missing = new FakeSpreadsheet(); missing.sheets = missing.sheets.slice(0, 1);
  assert.throws(() => lottery.executeLotteryOnSpreadsheet_(missing, runOptions), /運営確認/);
  assert.equal(missing.writes, 0);
});

test('latest tables are at the front, have fixed names and preserve inputs while compact history is append-only', () => {
  const data = balanced(); data[1][2] = '=1+1';
  const spreadsheet = new FakeSpreadsheet(data);
  const original = clone(spreadsheet.sheets.map(s => s.values));
  const first = lottery.executeLotteryOnSpreadsheet_(spreadsheet, runOptions);
  assert.equal(first.status, 'ACCEPTED');
  assert.equal(first.sheetNames.length, 7);
  assert.deepEqual(spreadsheet.sheets.filter(s => s.input).map(s => s.values), original);
  const outputs = spreadsheet.sheets.filter(s => lottery.readOutputNote_(s)?.kind === 'output');
  assert.ok(outputs.every(s => lottery.readOutputNote_(s).phase === 'published'));
  assert.deepEqual(spreadsheet.sheets.slice(0, 3).map(s => s.name), ['抽選_結果', '抽選_役職希望者', '抽選_実行情報']);
  const result = outputs.find(s => s.name.endsWith('_結果'));
  assert.ok(result.values.some(r => r.includes("'=1+1")));
  const leaders = outputs.find(s => s.name.endsWith('_役職希望者'));
  assert.equal(result.values[0][0], '班');
  assert.equal(result.values.length, 26);
  assert.equal(result.frozenRows, 1);
  assert.ok(result.filter);
  assert.equal(leaders.values.slice(1).length, 25);
  const snapshot = outputs.find(s => s.name.endsWith('_入力控え'));
  assert.equal(snapshot.values.slice(1).length, data.length + 1);
  const historySheet = spreadsheet.getSheetByName(lottery.LOTTERY_HISTORY_NAME);
  const previousHistory = clone(historySheet.values);
  const saved = clone(outputs.map(s => s.values));
  const second = lottery.executeLotteryOnSpreadsheet_(spreadsheet, runOptions);
  assert.equal(second.runId, 'fixture-run_1');
  assert.equal(spreadsheet.sheets.length, 10);
  assert.ok(outputs.every(s => s.deleted));
  assert.deepEqual(outputs.map(s => s.values), saved);
  assert.deepEqual(historySheet.values.slice(0, previousHistory.length), previousHistory);
  assert.ok(!JSON.stringify(historySheet.values).includes('架空'));
  const entries = lottery.historyRows_(historySheet);
  assert.equal(entries.filter(e => e.event === '保存完了').length, 2);
  const checked = lottery.verifyHistoryEntries_(lottery.readInputSheets_(spreadsheet), entries, digest);
  assert.deepEqual(checked.map(c => c.verification), ['一致', '一致']);
});

function legacyRun(spreadsheet, id = 'legacy-run') {
  const captured = lottery.readInputSheets_(spreadsheet);
  const result = engine(captured.responses.data, captured.operations.data);
  const record = lottery.makeRunRecord_(captured, result, { ...runOptions, runId: id });
  const metadata = [
    ['実行ID', id], ['状態', result.status === 'ACCEPTED' ? '成立' : '試行上限到達・選出未成立'],
    ['方式の版', '1.0.0'], ['シード', record.seed], ['乱数方式', lottery.LOTTERY_PRNG],
    ['元回答SHA256', record.responsesHash], ['運営確認SHA256', record.operationsHash],
    ['採用試行', result.acceptedAttempt || 'なし'], ['実施試行数', result.attempts.length],
    ['消費した乱数の個数', result.randomPosition], ['最終乱数状態', result.finalRandomState],
    ['実行開始日時', record.startedAt], ['タイムゾーン', captured.timezone],
    ['元回答シート・範囲', captured.responses.name + '!' + captured.responses.range],
    ['運営確認シート・範囲', captured.operations.name + '!' + captured.operations.range]
  ];
  const sheets = lottery.buildOutputPlans_(captured, result, runOptions).filter(p => p.label !== '実行情報').map(plan => {
    const sheet = new FakeSheet(spreadsheet, `抽選_${id}_${plan.label}`, [...metadata, [], ...plan.rows]);
    spreadsheet.sheets.push(sheet);
    return sheet;
  });
  return sheets;
}

test('complete legacy runs are recorded with reproducible hashes before their six output sheets are removed', () => {
  const spreadsheet = new FakeSpreadsheet();
  const old = legacyRun(spreadsheet);
  const summary = lottery.executeLotteryOnSpreadsheet_(spreadsheet, runOptions);
  assert.equal(summary.warnings.length, 0);
  assert.ok(old.every(s => s.deleted));
  assert.equal(spreadsheet.sheets.length, 10);
  const entries = lottery.historyRows_(spreadsheet.getSheetByName(lottery.LOTTERY_HISTORY_NAME));
  assert.equal(entries.filter(e => e.event === '旧版移行').length, 1);
  assert.deepEqual(lottery.verifyHistoryEntries_(lottery.readInputSheets_(spreadsheet), entries, digest)
    .map(item => item.verification), ['一致', '一致']);
});

test('incomplete legacy output sets are preserved with a warning', () => {
  const spreadsheet = new FakeSpreadsheet();
  const old = legacyRun(spreadsheet);
  spreadsheet.deleteSheet(old.pop());
  const summary = lottery.executeLotteryOnSpreadsheet_(spreadsheet, runOptions);
  assert.ok(old.every(sheet => spreadsheet.sheets.includes(sheet)));
  assert.ok(summary.warnings.some(warning => warning.includes('6種類がそろわない')));
  assert.equal(lottery.historyRows_(spreadsheet.getSheetByName(lottery.LOTTERY_HISTORY_NAME))
    .filter(entry => entry.event === '旧版移行').length, 0);
});

test('old-output cleanup failure keeps the newly saved tables and is recovered on the next run', () => {
  const spreadsheet = new FakeSpreadsheet();
  lottery.executeLotteryOnSpreadsheet_(spreadsheet, runOptions);
  const old = spreadsheet.getSheetByName('抽選_結果');
  const remove = spreadsheet.deleteSheet.bind(spreadsheet);
  spreadsheet.deleteSheet = sheet => { if (sheet === old) throw Error('simulated cleanup failure'); remove(sheet); };
  const summary = lottery.executeLotteryOnSpreadsheet_(spreadsheet, runOptions);
  assert.equal(summary.status, 'ACCEPTED');
  assert.equal(summary.warnings.length, 1);
  assert.ok(spreadsheet.sheets.includes(old));
  assert.notEqual(spreadsheet.getSheetByName('抽選_結果'), old);
  assert.equal(spreadsheet.sheets[0].name, '抽選_結果');
  spreadsheet.deleteSheet = remove;
  lottery.executeLotteryOnSpreadsheet_(spreadsheet, runOptions);
  assert.ok(old.deleted);
  assert.equal(spreadsheet.sheets.length, 10);
});

test('changed answers or operations halt replacement before any prior output or history is modified', () => {
  for (const kind of ['answers', 'operations']) {
    const spreadsheet = new FakeSpreadsheet();
    lottery.executeLotteryOnSpreadsheet_(spreadsheet, runOptions);
    const outputs = spreadsheet.sheets.filter(s => !s.input);
    const previous = clone(outputs.map(s => ({ id: s.id, name: s.name, values: s.values, notes: s.notes })));
    if (kind === 'answers') spreadsheet.sheets.find(s => s.input && s.name !== '運営確認').values[1][12] = '変更された回答';
    else spreadsheet.getSheetByName('運営確認').values.push(['A組', 1, '', '', '', '変更された運営記録']);
    const writes = spreadsheet.writes;
    assert.throws(() => lottery.executeLotteryOnSpreadsheet_(spreadsheet, runOptions), error => {
      assert.match(error.message, /過去の実行から変わっています/);
      if (kind === 'answers') {
        assert.match(error.message, /「フォームの回答 1」2行目 M2/);
        assert.match(error.message, /【質問10】自由記述/);
        assert.match(error.message, /""（文字列） → "変更された回答"（文字列）/);
      } else {
        assert.match(error.message, /「運営確認」2行目 F2/);
        assert.match(error.message, /メモ/);
        assert.match(error.message, /追加: （セルなし） → "変更された運営記録"（文字列）/);
      }
      return true;
    });
    assert.equal(spreadsheet.writes, writes);
    assert.deepEqual(outputs.map(s => ({ id: s.id, name: s.name, values: s.values, notes: s.notes })), previous);
    assert.ok(outputs.every(s => !s.deleted));
  }
});

test('input-change errors show both sheets, dates, value types and the actual operations header row', () => {
  const data = balanced();
  data[1][0] = new Date('2026-01-01T00:00:00.000Z');
  const ops = [['運営追記'], ...operations([['A組', 1, '', '', '', '以前のメモ']])];
  const spreadsheet = new FakeSpreadsheet(data, ops);
  lottery.executeLotteryOnSpreadsheet_(spreadsheet, runOptions);
  const responses = spreadsheet.sheets.find(sheet => sheet.input && sheet.name !== '運営確認');
  responses.values[1][0] = new Date('2026-01-02T00:00:00.000Z');
  responses.values[1][4] = '1'; // Same identity, different cell type.
  spreadsheet.getSheetByName('運営確認').values[2][5] = '新しいメモ';
  const writes = spreadsheet.writes;
  assert.throws(() => lottery.executeLotteryOnSpreadsheet_(spreadsheet, runOptions), error => {
    assert.match(error.message, /2行目 A2/);
    assert.match(error.message, /2026-01-01T00:00:00.000Z.*日時・UTC.* → .*2026-01-02T00:00:00.000Z/);
    assert.match(error.message, /2行目 E2.*出席番号.*1（数値） → "1"（文字列）/);
    assert.match(error.message, /「運営確認」3行目 F3.*メモ.*"以前のメモ".* → "新しいメモ"/);
    return true;
  });
  assert.equal(spreadsheet.writes, writes);
});

test('input-change errors identify deleted rows and the obsolete final column without changing old outputs', () => {
  for (const kind of ['row', 'column']) {
    const spreadsheet = new FakeSpreadsheet();
    lottery.executeLotteryOnSpreadsheet_(spreadsheet, runOptions);
    const responses = spreadsheet.sheets.find(sheet => sheet.input && sheet.name !== '運営確認');
    if (kind === 'row') responses.values.pop();
    else responses.values.forEach(row => row.pop());
    const saved = clone(spreadsheet.sheets.filter(sheet => !sheet.input).map(sheet => sheet.values));
    assert.throws(() => lottery.executeLotteryOnSpreadsheet_(spreadsheet, runOptions), error => {
      if (kind === 'row') {
        assert.match(error.message, /26行×14列 → 25行×14列/);
        assert.match(error.message, /26行目 A26.*削除: .* → （セルなし）/);
      } else {
        assert.match(error.message, /26行×14列 → 26行×13列/);
        assert.match(error.message, /1行目 N1.*重複立候補不可チェック.*見出し.*削除:/);
        assert.match(error.message, /差分 26セル/);
        assert.match(error.message, /ほか 6セル/);
      }
      return true;
    });
    assert.deepEqual(spreadsheet.sheets.filter(sheet => !sheet.input).map(sheet => sheet.values), saved);
  }
});

test('legacy input snapshots also produce cell-level change diagnostics', () => {
  const spreadsheet = new FakeSpreadsheet();
  const previous = legacyRun(spreadsheet);
  spreadsheet.sheets.find(sheet => sheet.input && sheet.name !== '運営確認').values[1][12] = '旧版からの変更';
  const writes = spreadsheet.writes;
  assert.throws(() => lottery.executeLotteryOnSpreadsheet_(spreadsheet, runOptions), error => {
    assert.match(error.message, /比較元の実行ID: legacy-run/);
    assert.match(error.message, /2行目 M2.*"旧版からの変更"/);
    return true;
  });
  assert.equal(spreadsheet.writes, writes);
  assert.ok(previous.every(sheet => !sheet.deleted));
});

test('missing or altered snapshots report unavailable differences rather than trusting changed copies', () => {
  for (const kind of ['missing', 'altered']) {
    const spreadsheet = new FakeSpreadsheet();
    lottery.executeLotteryOnSpreadsheet_(spreadsheet, runOptions);
    const snapshot = spreadsheet.getSheetByName('抽選_入力控え');
    if (kind === 'missing') spreadsheet.deleteSheet(snapshot);
    else {
      const cell = snapshot.values[2];
      const values = JSON.parse(cell[4]); values[12].value = '信用してはいけない控え';
      cell[4] = JSON.stringify(values);
    }
    spreadsheet.sheets.find(sheet => sheet.input && sheet.name !== '運営確認').values[1][12] = '現在の回答';
    const writes = spreadsheet.writes;
    assert.throws(() => lottery.executeLotteryOnSpreadsheet_(spreadsheet, runOptions), error => {
      assert.match(error.message, /行・値の差分を特定できません/);
      assert.match(error.message, /過去SHA256: [0-9a-f]{64}/);
      assert.doesNotMatch(error.message, /信用してはいけない控え|2行目 M2/);
      return true;
    });
    assert.equal(spreadsheet.writes, writes);
  }
});

test('large change errors show total counts and truncate values without changing source strings', () => {
  const spreadsheet = new FakeSpreadsheet();
  lottery.executeLotteryOnSpreadsheet_(spreadsheet, runOptions);
  const responses = spreadsheet.sheets.find(sheet => sheet.input && sheet.name !== '運営確認');
  const longValue = '架空の長いメモ😀\n'.repeat(50);
  responses.values.slice(1).forEach(row => { row[12] = longValue; });
  assert.throws(() => lottery.executeLotteryOnSpreadsheet_(spreadsheet, runOptions), error => {
    assert.match(error.message, /差分 25セル/);
    assert.match(error.message, /ほか 5セル。先頭20件のみ表示/);
    assert.match(error.message, /…（省略）/);
    assert.doesNotMatch(error.message, /26行目 M26/);
    return true;
  });
  assert.equal(responses.values[1][12], longValue);
});

test('unrelated sheets with output-like names are preserved and reserved-name collisions are rejected', () => {
  const spreadsheet = new FakeSpreadsheet();
  const unrelated = new FakeSheet(spreadsheet, '抽選_私用メモ', [['keep me']]);
  spreadsheet.sheets.push(unrelated);
  const summary = lottery.executeLotteryOnSpreadsheet_(spreadsheet, runOptions);
  assert.ok(summary.warnings.length);
  assert.ok(spreadsheet.sheets.includes(unrelated));
  assert.deepEqual(unrelated.values, [['keep me']]);
  const collision = new FakeSpreadsheet();
  collision.sheets.push(new FakeSheet(collision, '抽選_結果', [['user-owned']]));
  assert.throws(() => lottery.executeLotteryOnSpreadsheet_(collision, runOptions), /管理対象外シートと衝突/);
  assert.equal(collision.writes, 0);
});

test('a rename failure restores the previous latest set; the next successful run cleans only managed failed outputs', () => {
  const spreadsheet = new FakeSpreadsheet();
  lottery.executeLotteryOnSpreadsheet_(spreadsheet, runOptions);
  const old = spreadsheet.sheets.filter(s => lottery.readOutputNote_(s)?.kind === 'output');
  const saved = clone(old.map(s => ({ name: s.name, values: s.values })));
  spreadsheet.failRenameAt = spreadsheet.renames + 10;
  assert.throws(() => lottery.executeLotteryOnSpreadsheet_(spreadsheet, runOptions), /旧出力は削除していません/);
  assert.ok(old.every(s => !s.deleted));
  assert.deepEqual(old.map(s => ({ name: s.name, values: s.values })), saved);
  spreadsheet.failRenameAt = undefined;
  const summary = lottery.executeLotteryOnSpreadsheet_(spreadsheet, runOptions);
  assert.equal(summary.runId, 'fixture-run_2');
  assert.equal(spreadsheet.sheets.length, 10);
  assert.ok(old.every(s => s.deleted));
  const history = lottery.historyRows_(spreadsheet.getSheetByName(lottery.LOTTERY_HISTORY_NAME));
  assert.ok(history.some(e => e.event === '保存失敗'));
});

test('failure to append the new history record never deletes or renames previous latest outputs', () => {
  const spreadsheet = new FakeSpreadsheet();
  lottery.executeLotteryOnSpreadsheet_(spreadsheet, runOptions);
  const old = spreadsheet.sheets.filter(s => lottery.readOutputNote_(s)?.kind === 'output');
  const saved = clone(old.map(s => ({ name: s.name, values: s.values })));
  spreadsheet.failWriteAt = spreadsheet.writes + 8; // 7 new tables, then append-only history.
  assert.throws(() => lottery.executeLotteryOnSpreadsheet_(spreadsheet, runOptions), /旧出力は削除していません/);
  assert.ok(old.every(s => !s.deleted));
  assert.deepEqual(old.map(s => ({ name: s.name, values: s.values })), saved);
});

test('a final flush failure overrides a saved completion event and restores previous latest tables to the front', () => {
  const spreadsheet = new FakeSpreadsheet();
  lottery.executeLotteryOnSpreadsheet_(spreadsheet, runOptions);
  const oldResult = spreadsheet.getSheetByName('抽選_結果');
  let flushes = 0;
  assert.throws(() => lottery.executeLotteryOnSpreadsheet_(spreadsheet, {
    ...runOptions, flush() { if (++flushes === 5) throw Error('final flush failed'); }
  }), /旧出力は削除していません/);
  assert.equal(spreadsheet.getSheetByName('抽選_結果'), oldResult);
  assert.equal(spreadsheet.sheets[0], oldResult);
  const entries = lottery.historyRows_(spreadsheet.getSheetByName(lottery.LOTTERY_HISTORY_NAME));
  assert.ok(entries.some(e => e.record.runId === 'fixture-run_1' && e.event === '保存完了'));
  assert.ok(entries.some(e => e.record.runId === 'fixture-run_1' && e.event === '保存失敗'));
  assert.deepEqual(lottery.verifyHistoryEntries_(lottery.readInputSheets_(spreadsheet), entries, digest)
    .map(item => item.runId), ['fixture-run']);
});

test('user-made duplicates of managed outputs are not silently deleted', () => {
  const spreadsheet = new FakeSpreadsheet();
  lottery.executeLotteryOnSpreadsheet_(spreadsheet, runOptions);
  const result = spreadsheet.getSheetByName('抽選_結果');
  const duplicate = new FakeSheet(spreadsheet, '抽選_結果のコピー', result.values);
  duplicate.notes = clone(result.notes); spreadsheet.sheets.push(duplicate);
  assert.throws(() => lottery.executeLotteryOnSpreadsheet_(spreadsheet, runOptions), /同じ種類の出力が複数/);
  assert.ok(spreadsheet.sheets.includes(result));
  assert.ok(spreadsheet.sheets.includes(duplicate));
});

test('read-only replay detects mismatching hashes and changed input without modifying sheets', () => {
  const spreadsheet = new FakeSpreadsheet();
  lottery.executeLotteryOnSpreadsheet_(spreadsheet, runOptions);
  const captured = lottery.readInputSheets_(spreadsheet);
  const entries = lottery.historyRows_(spreadsheet.getSheetByName(lottery.LOTTERY_HISTORY_NAME));
  const altered = clone(entries);
  altered.find(e => e.event === '保存完了').record.auditHash = '00'.repeat(32);
  const writes = spreadsheet.writes;
  assert.equal(lottery.verifyHistoryEntries_(captured, altered, digest)[0].verification, '結果または全試行ログが不一致');
  captured.operations.values.push(['A組', 1, '', '', '', 'changed']);
  assert.equal(lottery.verifyHistoryEntries_(captured, entries, digest)[0].verification, '入力が当時と異なる');
  assert.equal(spreadsheet.writes, writes);
});

test('unsuccessful attempts leave result and leadership output empty while preserving all attempt statistics', () => {
  const spreadsheet = new FakeSpreadsheet(oneLeaderCommittee());
  const summary = lottery.executeLotteryOnSpreadsheet_(spreadsheet, runOptions);
  assert.equal(summary.status, 'ATTEMPT_LIMIT_REACHED');
  for (const suffix of ['_結果', '_役職希望者']) {
    const sheet = spreadsheet.sheets.find(s => s.name.endsWith(suffix));
    assert.equal(sheet.values.slice(1).length, 0);
    assert.equal(lottery.readOutputNote_(sheet).record.status, 'ATTEMPT_LIMIT_REACHED');
  }
  assert.equal(spreadsheet.sheets[0].name, '抽選_実行情報');
  assert.equal(spreadsheet.getSheetByName('抽選_実行情報').values[2][1], '試行上限到達・選出未成立');
  const stats = spreadsheet.sheets.find(s => s.name.endsWith('_結果統計'));
  assert.deepEqual([...new Set(stats.values.slice(1).map(r => r[0]))], [1, 2, 3, 4, 5, 6]);
});

test('input validation writes nothing; partial output failure is retained and marked unconfirmed', () => {
  const invalid = balanced(); invalid[1][4] = '';
  const invalidSheet = new FakeSpreadsheet(invalid);
  assert.throws(() => lottery.executeLotteryOnSpreadsheet_(invalidSheet, runOptions));
  assert.equal(invalidSheet.sheets.length, 2);
  const spreadsheet = new FakeSpreadsheet(); spreadsheet.failWriteAt = 2;
  const original = clone(spreadsheet.sheets.map(s => s.values));
  assert.throws(() => lottery.executeLotteryOnSpreadsheet_(spreadsheet, runOptions), /出力の保存に失敗/);
  assert.equal(spreadsheet.sheets.length, 5);
  assert.deepEqual(spreadsheet.sheets.filter(s => s.input).map(s => s.values), original);
  assert.ok(spreadsheet.sheets.filter(s => lottery.readOutputNote_(s)?.kind === 'output')
    .every(s => lottery.readOutputNote_(s).phase === 'failed'));
});

test('menu setup and the legacy preview entry do not run lotteries or modify sheets', () => {
  const spreadsheet = new FakeSpreadsheet();
  const menu = { addItem() { return this; }, addToUi() { return this; } };
  const alerts = [];
  const ui = { createMenu: () => menu, alert: (...args) => alerts.push(args), ButtonSet: { OK: 'OK' } };
  const context = { SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet, getUi: () => ui } };
  vm.createContext(context);
  vm.runInContext(fs.readFileSync(require.resolve('../apps-script/CommitteeLottery.js'), 'utf8'), context);
  context.onOpen();
  const summary = context.myFunction();
  assert.equal(summary.eligible, 25);
  assert.equal(spreadsheet.sheets.length, 2);
  assert.equal(spreadsheet.writes, 0);
  assert.equal(alerts.length, 1);
});

test('buffered output failure cannot confirm winners', () => {
  const spreadsheet = new FakeSpreadsheet();
  let calls = 0;
  assert.throws(() => lottery.executeLotteryOnSpreadsheet_(spreadsheet, {
    ...runOptions, flush() { if (++calls === 1) throw Error('simulated buffered failure'); }
  }), /出力の保存に失敗/);
  assert.ok(spreadsheet.sheets.filter(s => lottery.readOutputNote_(s)?.kind === 'output')
    .every(s => lottery.readOutputNote_(s).phase === 'failed'));
});

test('Apps Script entry point uses signed digest bytes, flushes outputs and releases the document lock', () => {
  const spreadsheet = new FakeSpreadsheet();
  const notifications = [];
  spreadsheet.toast = (...args) => notifications.push(args);
  let releases = 0, flushes = 0;
  const context = {
    LOTTERY_BUILD_REVISION: '1'.repeat(40),
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet, flush: () => { flushes++; } },
    LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock: () => { releases++; } }) },
    Utilities: {
      DigestAlgorithm: { SHA_256: 'SHA_256' },
      computeDigest: (algorithm, bytes) => {
        assert.equal(algorithm, 'SHA_256');
        assert.ok(bytes.every(b => Number.isInteger(b) && b >= -128 && b <= 127));
        return digest(bytes).map(b => b > 127 ? b - 256 : b);
      },
      formatDate: () => 'fixture-run', getUuid: () => '00000000-fixture'
    }
  };
  vm.createContext(context);
  vm.runInContext(fs.readFileSync(require.resolve('../apps-script/CommitteeLottery.js'), 'utf8'), context);
  const summary = context.runCommitteeLottery();
  assert.equal(summary.status, 'ACCEPTED');
  assert.equal(summary.selected, 25);
  assert.equal(releases, 1);
  assert.equal(flushes, 5);
  assert.equal(notifications.length, 1);
  const saved = lottery.historyRows_(spreadsheet.getSheetByName(lottery.LOTTERY_HISTORY_NAME));
  assert.equal(saved.find(entry => entry.event === '保存完了').record.sourceRevision, '1'.repeat(40));
  const denied = { ...context, LockService: { getDocumentLock: () => ({ tryLock: () => false }) } };
  vm.createContext(denied);
  vm.runInContext(fs.readFileSync(require.resolve('../apps-script/CommitteeLottery.js'), 'utf8'), denied);
  assert.throws(() => denied.runCommitteeLottery(), /別の抽選処理/);
  assert.equal(spreadsheet.sheets.length, 10);
});
