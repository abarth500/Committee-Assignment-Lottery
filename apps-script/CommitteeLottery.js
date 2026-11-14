const LOTTERY_SEED = 'a4e9d4bfe3a04a2eb53914b029132873c0d2774400784ecabdc141db0d6348fb';
const LOTTERY_VERSION = '1.2.1';
const LOTTERY_ALGORITHM_VERSION = '1.0.0';
const LOTTERY_RESERVE_VERSION = '1.1.0';
const LOTTERY_RESERVE_LABELS = Object.freeze(['補欠名簿', '補欠記録']);
const LOTTERY_PRNG = 'SHA-256-chain / uint32-be / rejection-sampling / Fisher-Yates';
const LOTTERY_OUTPUT_OWNER = 'CommitteeAssignmentLottery/output-v2';
const LOTTERY_OUTPUT_LABELS = Object.freeze(['結果', '役職希望者', '実行情報', '詳細ログ', '元データ統計', '結果統計', '入力控え']);
const LOTTERY_HISTORY_NAME = '抽選_実行履歴';
const LOTTERY_HISTORY_HEADER = Object.freeze(['実行日時', '実行ID', '記録種別', '抽選状態', '採用試行', '結果SHA256', '全試行ログSHA256', '再検証情報（JSON）']);
const LOTTERY_SETTINGS = Object.freeze({
  committees: Object.freeze(['企画班', '司会進行班', '音響照明班', '会場演出班', '有志団体統括班']),
  classes: Object.freeze(['A組', 'B組', 'C組', 'D組', 'E組']),
  capacity: 5,
  minimumPerClass: 2,
  maximumAttempts: 6,
  responseSheetName: '', // 空欄なら、先頭行の質問見出しから回答シートを一意に特定する。
  operationsSheetName: '運営確認',
  outputPrefix: '抽選_'
});

/** メニュー追加だけを行う。開いたときに抽選や書き込みを実行しない。 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('委員抽選')
    .addItem('入力・成立条件を確認（書込なし）', 'previewCommitteeLottery')
    .addItem('抽選を実行して最新結果を保存', 'runCommitteeLottery')
    .addItem('保存済み結果から補欠名簿を作成', 'createCommitteeReserves')
    .addItem('実行履歴をハッシュ照合（書込なし）', 'verifyLotteryHistory')
    .addToUi();
}

/** 以前のデプロイ確認用エントリーは、乱数を消費しない入力確認として残す。 */
function myFunction() {
  return previewCommitteeLottery();
}

function previewCommitteeLottery() {
  const captured = readInputSheets_(SpreadsheetApp.getActiveSpreadsheet());
  const prepared = prepareLotteryInput_(captured.responses.data, captured.operations.data, captured);
  const check = checkPossible_(prepared.eligible);
  const summary = {
    responses: prepared.responseCount,
    applicants: prepared.candidates.length,
    excluded: prepared.candidates.length - prepared.eligible.length,
    eligible: prepared.eligible.length,
    reasons: check.reasons
  };
  SpreadsheetApp.getUi().alert('入力確認（抽選・書込なし）',
    '回答 ' + summary.responses + '件／応募者 ' + summary.applicants + '人\n' +
    '除外 ' + summary.excluded + '人／抽選対象 ' + summary.eligible + '人\n' +
    (check.reasons.length ? check.reasons.join('\n') : '事前の必要条件を満たします。抽選成立の保証ではありません。'),
    SpreadsheetApp.getUi().ButtonSet.OK);
  return summary;
}

/** 最新出力の置換は、入力照合・履歴保存・新出力の保存が済んだ後に行う。 */
function runCommitteeLottery() {
  const lock = LockService.getDocumentLock();
  if (!lock || !lock.tryLock(1000)) {
    throw new Error('別の抽選処理が実行中か、スプレッドシートに紐づいていません。');
  }
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const now = new Date();
    const runId = Utilities.formatDate(now, spreadsheet.getSpreadsheetTimeZone(), 'yyyyMMdd_HHmmss_SSS') +
      '_' + Utilities.getUuid().slice(0, 8);
    const summary = executeLotteryOnSpreadsheet_(spreadsheet, {
      seed: LOTTERY_SEED, runId: runId, startedAt: now.toISOString(),
      flush: function () { SpreadsheetApp.flush(); }
    });
    const message = summary.status === 'ACCEPTED' ? '最新の抽選結果を保存しました。' :
      '今回は選出未成立です。実行情報とログを確認してください。';
    spreadsheet.toast(message + (summary.warnings.length ? '\n旧出力の整理に注意事項があります。' : ''), '委員抽選', 10);
    return summary;
  } finally {
    lock.releaseLock();
  }
}

function text_(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

function normalizedText_(value) {
  return text_(value).normalize('NFKC');
}

function personKey_(classValue, numberValue, location) {
  const className = normalizedText_(classValue);
  const numberText = normalizedText_(numberValue).replace(/\s/g, '').replace(/番$/, '');
  const number = Number(numberText);
  if (!LOTTERY_SETTINGS.classes.includes(className) || !/^\d+$/.test(numberText) ||
      !Number.isSafeInteger(number) || number < 1) {
    throw new Error(location + ': クラスまたは出席番号を確認してください。');
  }
  return { key: className + ':' + number, className: className, number: number };
}

function resolveColumns_(header, rules, location) {
  const columns = {};
  Object.keys(rules).forEach(function (field) {
    const matches = [];
    header.forEach(function (value, index) {
      const label = normalizedText_(value).replace(/\s/g, '');
      if (rules[field](label)) matches.push(index);
    });
    if (matches.length !== 1) throw new Error(location + ': 見出し「' + field + '」を一意に特定できません。');
    columns[field] = matches[0];
  });
  return columns;
}

function formColumns_(header) {
  return resolveColumns_(header, {
    name: function (s) { return s === '名前'; },
    className: function (s) { return s === 'クラス'; },
    number: function (s) { return s === '出席番号'; },
    chair: function (s) { return s.startsWith('【質問2】'); },
    first: function (s) { return s.startsWith('【質問3】'); },
    second: function (s) { return s.startsWith('【質問4】'); },
    orientation: function (s) { return s.startsWith('【質問1】'); },
    meetings: function (s) { return s.startsWith('【質問7】'); }
  }, '回答シート');
}

function operationsColumns_(header) {
  const fields = ['クラス', '番号', '11月21日不参加特別理由', '打ち合わせ不参加特別理由', '重複不可委員就任済み', 'メモ'];
  const rules = {};
  fields.forEach(function (field) { rules[field] = function (s) { return s === field; }; });
  return resolveColumns_(header, rules, '運営確認／運営追記');
}

function readInputSheets_(spreadsheet) {
  if (!spreadsheet) throw new Error('対象のスプレッドシートがありません。');
  const operationsSheet = spreadsheet.getSheetByName(LOTTERY_SETTINGS.operationsSheetName);
  if (!operationsSheet) throw new Error('「運営確認」シートがありません。');
  const candidates = [];
  spreadsheet.getSheets().forEach(function (sheet) {
    if (sheet.getSheetId() === operationsSheet.getSheetId() ||
        sheet.getName().startsWith(LOTTERY_SETTINGS.outputPrefix)) return;
    if (LOTTERY_SETTINGS.responseSheetName && sheet.getName() !== LOTTERY_SETTINGS.responseSheetName) return;
    const range = sheet.getDataRange();
    const values = range.getValues();
    try { formColumns_(values[0] || []); } catch (error) { return; }
    candidates.push({ sheet: sheet, range: range, values: values });
  });
  if (candidates.length !== 1) {
    throw new Error('回答シートを一意に特定できません。見出し、または responseSheetName の設定を確認してください。');
  }
  const opRange = operationsSheet.getDataRange();
  const opValues = opRange.getValues();
  const headerRows = [];
  opValues.forEach(function (row, index) {
    try { operationsColumns_(row); headerRows.push(index); } catch (error) { /* 見出し以外の行 */ }
  });
  if (headerRows.length !== 1) {
    throw new Error('「運営確認」上で「運営追記」の6項目の見出しを一意に特定できません。');
  }
  function capturedSheet(sheet, range, values, headerIndex) {
    return {
      name: sheet.getName(), id: sheet.getSheetId(), range: range.getA1Notation(),
      values: values, headerRow: headerIndex + 1, data: values.slice(headerIndex)
    };
  }
  return {
    spreadsheetId: spreadsheet.getId(), timezone: spreadsheet.getSpreadsheetTimeZone(),
    responses: capturedSheet(candidates[0].sheet, candidates[0].range, candidates[0].values, 0),
    operations: capturedSheet(operationsSheet, opRange, opValues, headerRows[0])
  };
}

function preference_(value, optional, location) {
  const label = normalizedText_(value);
  if (['特に希望なし', '特になし', '希望なし'].includes(label) || (optional && label === '')) return null;
  if (!LOTTERY_SETTINGS.committees.includes(label)) throw new Error(location + ': 希望役割を確認してください。');
  return label;
}

function prepareLotteryInput_(responseValues, operationValues, locations) {
  locations = locations || {};
  const responseHeader = locations.responses ? locations.responses.headerRow : 1;
  const operationHeader = locations.operations ? locations.operations.headerRow : 1;
  const columns = formColumns_(responseValues[0] || []);
  const opColumns = operationsColumns_(operationValues[0] || []);
  const audit = [];
  function record(type, detail) { audit.push({ sequence: audit.length + 1, type: type, detail: detail }); }
  const people = new Map();
  const lastResponses = new Map();
  let responseCount = 0;
  responseValues.slice(1).forEach(function (row, index) {
    if (row.every(function (v) { return text_(v) === ''; })) return;
    responseCount += 1;
    const sourceRow = responseHeader + index + 1;
    const location = '回答シート ' + sourceRow + '行目';
    const identity = personKey_(row[columns.className], row[columns.number], location);
    if (lastResponses.has(identity.key)) {
      record('DUPLICATE_REPLACED', {
        key: identity.key, discardedRow: lastResponses.get(identity.key).sourceRow, adoptedRow: sourceRow
      });
    }
    lastResponses.set(identity.key, { identity: identity, row: row, sourceRow: sourceRow });
  });
  lastResponses.forEach(function (response) {
    const identity = response.identity; const row = response.row; const sourceRow = response.sourceRow;
    const location = '回答シート ' + sourceRow + '行目';
    const chair = normalizedText_(row[columns.chair]);
    const orientation = normalizedText_(row[columns.orientation]);
    if (!['希望する', '希望しない'].includes(chair)) throw new Error(location + ': 役職希望の回答を確認してください。');
    if (!['参加できます', '参加できません'].includes(orientation)) {
      throw new Error(location + ': オリエンテーション参加可否を確認してください。');
    }
    const name = text_(row[columns.name]);
    if (!name) throw new Error(location + ': 名前が空欄です。');
    people.set(identity.key, Object.assign({}, identity, {
      name: name, sourceRow: sourceRow, chairHope: chair === '希望する',
      firstPreference: preference_(row[columns.first], false, location),
      secondPreference: preference_(row[columns.second], true, location),
      orientation: orientation, meetings: normalizedText_(row[columns.meetings]),
      exclusionReasons: [], appliedExceptions: [], operationRow: null
    }));
  });
  const operations = new Map();
  operationValues.slice(1).forEach(function (row, index) {
    if (row.every(function (v) { return text_(v) === ''; })) return;
    const sourceRow = operationHeader + index + 1;
    const location = '運営確認 ' + sourceRow + '行目';
    const identity = personKey_(row[opColumns['クラス']], row[opColumns['番号']], location);
    if (!people.has(identity.key)) throw new Error(location + ': 対応する応募者がいません。');
    if (operations.has(identity.key)) throw new Error(location + ': 同一人物の運営確認が重複しています。');
    const orientation = normalizedText_(row[opColumns['11月21日不参加特別理由']]);
    const meetings = normalizedText_(row[opColumns['打ち合わせ不参加特別理由']]);
    if (![orientation, meetings].every(function (value) { return value === '' || value === '公欠'; })) {
      throw new Error(location + ': 特別理由は空欄または承認済みの「公欠」で記録してください。');
    }
    operations.set(identity.key, Object.assign({}, identity, {
      sourceRow: sourceRow, orientationExcused: orientation === '公欠', meetingsExcused: meetings === '公欠',
      otherCommittee: text_(row[opColumns['重複不可委員就任済み']]), memo: text_(row[opColumns['メモ']])
    }));
  });
  const candidates = Array.from(people.values()).sort(comparePeople_);
  candidates.forEach(function (person) {
    const operation = operations.get(person.key);
    person.operationRow = operation ? operation.sourceRow : null;
    if (operation && operation.otherCommittee) person.exclusionReasons.push('OTHER_COMMITTEE');
    if (person.orientation === '参加できません') {
      if (operation && operation.orientationExcused) person.appliedExceptions.push('ORIENTATION_EXCUSED');
      else person.exclusionReasons.push('ORIENTATION_UNAVAILABLE');
    }
    if (person.meetings === '参加できません') {
      if (operation && operation.meetingsExcused) person.appliedExceptions.push('MEETINGS_EXCUSED');
      else person.exclusionReasons.push('MEETINGS_UNAVAILABLE');
    }
    record('ELIGIBILITY', {
      key: person.key, sourceRow: person.sourceRow, operationRow: person.operationRow,
      orientation: person.orientation, meetings: person.meetings,
      exclusions: person.exclusionReasons.slice(), exceptions: person.appliedExceptions.slice()
    });
  });
  return {
    responseCount: responseCount, duplicateCount: responseCount - candidates.length,
    candidates: candidates, eligible: candidates.filter(function (p) { return p.exclusionReasons.length === 0; }),
    operations: Array.from(operations.values()).sort(comparePeople_), audit: audit
  };
}

function comparePeople_(a, b) {
  return LOTTERY_SETTINGS.classes.indexOf(a.className) - LOTTERY_SETTINGS.classes.indexOf(b.className) ||
    a.number - b.number;
}

function possibleCommittees_(person) {
  const roles = person.firstPreference ? [person.firstPreference] : LOTTERY_SETTINGS.committees.slice();
  if (person.secondPreference && !roles.includes(person.secondPreference)) roles.push(person.secondPreference);
  return roles;
}

/** 希望順位を緩和した最大流。必要条件の確認であり、実際の抽選を置き換えない。 */
function maximumClassCoverage_(people) {
  const classStart = 1;
  const personStart = classStart + LOTTERY_SETTINGS.classes.length;
  const committeeStart = personStart + people.length;
  const sink = committeeStart + LOTTERY_SETTINGS.committees.length;
  const graph = Array.from({ length: sink + 1 }, function () { return []; });
  function edge(from, to, capacity) {
    const forward = { to: to, reverse: graph[to].length, capacity: capacity };
    const backward = { to: from, reverse: graph[from].length, capacity: 0 };
    graph[from].push(forward); graph[to].push(backward);
  }
  LOTTERY_SETTINGS.classes.forEach(function (name, i) { edge(0, classStart + i, LOTTERY_SETTINGS.minimumPerClass); });
  people.forEach(function (person, i) {
    edge(classStart + LOTTERY_SETTINGS.classes.indexOf(person.className), personStart + i, 1);
    possibleCommittees_(person).forEach(function (name) {
      edge(personStart + i, committeeStart + LOTTERY_SETTINGS.committees.indexOf(name), 1);
    });
  });
  LOTTERY_SETTINGS.committees.forEach(function (name, i) { edge(committeeStart + i, sink, LOTTERY_SETTINGS.capacity); });
  let flow = 0;
  while (true) {
    const previous = Array(graph.length).fill(null);
    const queue = [0]; previous[0] = { from: -1 };
    for (let head = 0; head < queue.length && previous[sink] === null; head += 1) {
      const from = queue[head];
      graph[from].forEach(function (item, index) {
        if (item.capacity > 0 && previous[item.to] === null) {
          previous[item.to] = { from: from, edge: index }; queue.push(item.to);
        }
      });
    }
    if (previous[sink] === null) return flow;
    for (let to = sink; to !== 0;) {
      const step = previous[to]; const item = graph[step.from][step.edge];
      item.capacity -= 1; graph[to][item.reverse].capacity += 1; to = step.from;
    }
    flow += 1;
  }
}

function checkPossible_(people) {
  const reasons = [];
  const classCounts = countBy_(people, function (p) { return p.className; }, LOTTERY_SETTINGS.classes);
  LOTTERY_SETTINGS.classes.forEach(function (name) {
    if (classCounts[name] < LOTTERY_SETTINGS.minimumPerClass) reasons.push(name + 'の対象者が2人未満');
  });
  const leaders = people.filter(function (p) { return p.chairHope; });
  if (leaders.length < 2) reasons.push('役職希望者が2人未満');
  const leaderRoles = new Set();
  leaders.forEach(function (p) { possibleCommittees_(p).forEach(function (role) { leaderRoles.add(role); }); });
  if (leaderRoles.size < 2) reasons.push('役職希望者の希望先が2班未満');
  const coverage = maximumClassCoverage_(people);
  if (coverage < LOTTERY_SETTINGS.classes.length * LOTTERY_SETTINGS.minimumPerClass) {
    reasons.push('希望順位を緩和しても各クラス2人を班の定員内に配置できない');
  }
  return { reasons: reasons, classCounts: classCounts, leadershipPeople: leaders.length, optimisticClassCoverage: coverage };
}

function utf8Bytes_(value) {
  const result = [];
  for (const character of String(value)) {
    let code = character.codePointAt(0);
    if (code >= 0xd800 && code <= 0xdfff) code = 0xfffd;
    if (code < 0x80) result.push(code);
    else if (code < 0x800) result.push(0xc0 | (code >> 6), 0x80 | (code & 63));
    else if (code < 0x10000) result.push(0xe0 | (code >> 12), 0x80 | ((code >> 6) & 63), 0x80 | (code & 63));
    else result.push(0xf0 | (code >> 18), 0x80 | ((code >> 12) & 63), 0x80 | ((code >> 6) & 63), 0x80 | (code & 63));
  }
  return result;
}

function digestBytes_(bytes) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    bytes.map(function (b) { return b > 127 ? b - 256 : b; })).map(function (b) { return (b + 256) % 256; });
}

function hex_(bytes) {
  return Array.from(bytes, function (b) { return b.toString(16).padStart(2, '0'); }).join('');
}

function hashText_(value, digest) {
  return hex_((digest || digestBytes_)(utf8Bytes_(value)));
}

function createRandom_(seed, record, digest) {
  if (typeof seed !== 'string' || !/^[0-9a-fA-F]{64}$/.test(seed)) {
    throw new Error('シードは64桁の16進数としてプログラム冒頭で固定してください。');
  }
  let state = seed.match(/../g).map(function (part) { return parseInt(part, 16); });
  let count = 0;
  const hash = digest || digestBytes_;
  return {
    nextInt: function (bound, context) {
      if (!Number.isSafeInteger(bound) || bound < 1 || bound > 0x100000000) throw new Error('乱数の範囲が不正です。');
      const limit = Math.floor(0x100000000 / bound) * bound;
      while (true) {
        const before = hex_(state);
        state = Array.from(hash(state.slice()), function (b) { return (b + 256) % 256; });
        if (state.length !== 32) throw new Error('SHA-256の結果が不正です。');
        count += 1;
        const word = state[0] * 0x1000000 + state[1] * 0x10000 + state[2] * 0x100 + state[3];
        const accepted = word < limit;
        const value = accepted ? word % bound : null;
        record('RANDOM', { position: count, context: context, stateBefore: before, stateAfter: hex_(state),
          word: word, bound: bound, limit: limit, accepted: accepted, value: value });
        if (accepted) return value;
      }
    },
    position: function () { return count; },
    state: function () { return hex_(state); }
  };
}

function pick_(pool, capacity, random, context) {
  if (pool.length <= capacity) return pool.slice();
  const shuffled = pool.slice();
  for (let i = shuffled.length - 1; i > 0; i -= 1) {
    const j = random.nextInt(i + 1, context + '/swap=' + i);
    const item = shuffled[i]; shuffled[i] = shuffled[j]; shuffled[j] = item;
  }
  return shuffled.slice(0, capacity);
}

function countBy_(items, getKey, initialKeys) {
  const result = {};
  (initialKeys || []).forEach(function (key) { result[key] = 0; });
  items.forEach(function (item) { const key = getKey(item); result[key] = (result[key] || 0) + 1; });
  return result;
}

/** Apps Scriptサービスに依存しない抽選本体。digestだけテストで差し替え可能。 */
function runLotteryEngine_(responseValues, operationValues, options) {
  options = options || {};
  const seed = options.seed === undefined ? LOTTERY_SEED : options.seed;
  if (typeof seed !== 'string' || !/^[0-9a-fA-F]{64}$/.test(seed)) throw new Error('シードは64桁の16進数が必要です。');
  const prepared = prepareLotteryInput_(responseValues, operationValues, options.locations);
  const audit = prepared.audit;
  function record(type, detail) { audit.push({ sequence: audit.length + 1, type: type, detail: detail }); }
  const check = checkPossible_(prepared.eligible);
  record('PRECHECK', check);
  const result = {
    version: LOTTERY_VERSION, prng: LOTTERY_PRNG, seed: seed.toLowerCase(), prepared: prepared,
    precheck: check, audit: audit, firstPreferenceAssignments: [], effectiveFirstByKey: {},
    attempts: [], assignments: [], leadershipCandidates: [], acceptedAttempt: null,
    status: 'PRECHECK_FAILED', randomPosition: 0, finalRandomState: seed.toLowerCase()
  };
  if (check.reasons.length) {
    record('FINISHED', { status: result.status, acceptedAttempt: null, attempts: 0,
      randomPosition: 0, randomState: result.finalRandomState });
    return result;
  }
  const random = createRandom_(seed, record, options.digest);
  const lookup = new Map(prepared.eligible.map(function (p) { return [p.key, p]; }));
  prepared.eligible.forEach(function (person) {
    let first = person.firstPreference;
    if (first === null) {
      first = LOTTERY_SETTINGS.committees[random.nextInt(LOTTERY_SETTINGS.committees.length, 'NO_FIRST_PREFERENCE/' + person.key)];
      const assignment = { key: person.key, committee: first };
      result.firstPreferenceAssignments.push(assignment);
      record('FIXED_FIRST_PREFERENCE', assignment);
    }
    result.effectiveFirstByKey[person.key] = first;
  });
  for (let number = 1; number <= LOTTERY_SETTINGS.maximumAttempts; number += 1) {
    record('ATTEMPT_START', { attempt: number, randomPosition: random.position(), randomState: random.state() });
    const assignments = []; const selected = new Set();
    function draw(committee, stage, pool, capacity) {
      const chosen = pick_(pool, capacity, random, 'attempt=' + number + '/stage=' + stage + '/' + committee);
      const chosenKeys = new Set(chosen.map(function (p) { return p.key; }));
      record('POOL', { attempt: number, stage: stage, committee: committee, capacity: capacity,
        pool: pool.map(function (p) { return p.key; }), chosen: Array.from(chosenKeys) });
      pool.forEach(function (person) {
        record('DECISION', { attempt: number, stage: stage, committee: committee, key: person.key,
          selected: chosenKeys.has(person.key), method: pool.length <= capacity ? 'ALL_WITHIN_CAPACITY' : 'SHUFFLE' });
      });
      chosen.forEach(function (person) {
        selected.add(person.key); assignments.push({ key: person.key, committee: committee, stage: stage });
      });
    }
    LOTTERY_SETTINGS.committees.forEach(function (committee) {
      draw(committee, 1, prepared.eligible.filter(function (p) { return result.effectiveFirstByKey[p.key] === committee; }),
        LOTTERY_SETTINGS.capacity);
    });
    const firstCommitteeCounts = countBy_(assignments, function (a) { return a.committee; }, LOTTERY_SETTINGS.committees);
    LOTTERY_SETTINGS.committees.forEach(function (committee) {
      const vacancies = LOTTERY_SETTINGS.capacity - assignments.filter(function (a) { return a.committee === committee; }).length;
      if (vacancies <= 0) return;
      draw(committee, 2, prepared.eligible.filter(function (p) {
        return !selected.has(p.key) && p.secondPreference === committee;
      }), vacancies);
    });
    const classCounts = countBy_(assignments, function (a) { return lookup.get(a.key).className; }, LOTTERY_SETTINGS.classes);
    const committeeCounts = countBy_(assignments, function (a) { return a.committee; }, LOTTERY_SETTINGS.committees);
    const leadershipCandidates = assignments.filter(function (a) { return lookup.get(a.key).chairHope; });
    const leadershipCommittees = LOTTERY_SETTINGS.committees.filter(function (name) {
      return leadershipCandidates.some(function (a) { return a.committee === name; });
    });
    const failures = [];
    LOTTERY_SETTINGS.classes.forEach(function (name) {
      if (classCounts[name] < LOTTERY_SETTINGS.minimumPerClass) failures.push(name + 'の当選者が2人未満');
    });
    if (leadershipCommittees.length < 2) failures.push('当選した役職希望者の所属が2班未満');
    const attempt = { number: number, assignments: assignments, classCounts: classCounts, committeeCounts: committeeCounts,
      leadershipCandidates: leadershipCandidates, leadershipCommittees: leadershipCommittees,
      failures: failures, accepted: failures.length === 0, randomPosition: random.position(), randomState: random.state() };
    result.attempts.push(attempt);
    const assignedByKey = new Map(assignments.map(function (a) { return [a.key, a]; }));
    prepared.eligible.forEach(function (person) {
      const assignment = assignedByKey.get(person.key);
      let reason;
      if (assignment) reason = assignment.stage === 1 ? 'FIRST_PREFERENCE_SELECTED' : 'SECOND_PREFERENCE_SELECTED';
      else if (person.secondPreference === null) reason = 'NO_SECOND_PREFERENCE';
      else if (firstCommitteeCounts[person.secondPreference] >= LOTTERY_SETTINGS.capacity) reason = 'SECOND_PREFERENCE_FULL';
      else reason = 'SECOND_PREFERENCE_DRAW_LOSS';
      record('TRIAL_CANDIDATE_STATUS', { attempt: number, attemptAccepted: attempt.accepted,
        key: person.key, selected: Boolean(assignment), committee: assignment ? assignment.committee : null,
        stage: assignment ? assignment.stage : null, reason: reason });
    });
    record('ATTEMPT_END', { attempt: number, accepted: attempt.accepted, failures: failures,
      classCounts: classCounts, committeeCounts: committeeCounts, leadershipCommittees: leadershipCommittees,
      randomPosition: random.position(), randomState: random.state() });
    if (attempt.accepted) {
      result.acceptedAttempt = number; result.assignments = assignments;
      result.leadershipCandidates = leadershipCandidates; break;
    }
  }
  result.status = result.acceptedAttempt === null ? 'ATTEMPT_LIMIT_REACHED' : 'ACCEPTED';
  result.randomPosition = random.position(); result.finalRandomState = random.state();
  record('FINISHED', { status: result.status, acceptedAttempt: result.acceptedAttempt, attempts: result.attempts.length,
    randomPosition: result.randomPosition, randomState: result.finalRandomState });
  return result;
}

function serializedRow_(row) {
  return JSON.stringify(row.map(function (value) {
    if (value instanceof Date) return { type: 'date', value: value.toISOString() };
    return { type: typeof value, value: value };
  }));
}

function snapshotHash_(values, digest) {
  return hashText_(JSON.stringify(values.map(serializedRow_)), digest);
}

function displayStatus_(status) {
  return { ACCEPTED: '成立', PRECHECK_FAILED: '事前条件不成立', ATTEMPT_LIMIT_REACHED: '試行上限到達・選出未成立' }[status];
}

function sortedAssignments_(assignments, people) {
  return assignments.slice().sort(function (a, b) {
    return LOTTERY_SETTINGS.committees.indexOf(a.committee) - LOTTERY_SETTINGS.committees.indexOf(b.committee) ||
      comparePeople_(people.get(a.key), people.get(b.key));
  });
}

function buildOutputPlans_(captured, result, context) {
  const people = new Map(result.prepared.candidates.map(function (p) { return [p.key, p]; }));
  const metadata = [
    ['実行ID', context.runId], ['状態', '保存中・当選未確定'], ['コードの版', result.version], ['シード', result.seed],
    ['乱数方式', result.prng], ['元回答SHA256', snapshotHash_(captured.responses.values, context.digest)],
    ['運営確認SHA256', snapshotHash_(captured.operations.values, context.digest)], ['採用試行', result.acceptedAttempt || 'なし'],
    ['実施試行数', result.attempts.length], ['消費した乱数の個数', result.randomPosition],
    ['最終乱数状態', result.finalRandomState], ['実行開始日時', context.startedAt], ['タイムゾーン', captured.timezone],
    ['元回答シート・範囲', captured.responses.name + '!' + captured.responses.range],
    ['運営確認シート・範囲', captured.operations.name + '!' + captured.operations.range]
  ];
  const record = context.record || makeRunRecord_(captured, result, context);
  metadata.push(['抽選アルゴリズムの版', record.algorithmVersion], ['Gitリビジョン', record.sourceRevision],
    ['結果SHA256', record.resultHash], ['全試行ログSHA256', record.auditHash]);
  function plan(label, header, rows) { return { label: label, rows: [header].concat(rows), record: record }; }
  const selectedRows = sortedAssignments_(result.assignments, people).map(function (assignment) {
    const person = people.get(assignment.key);
    return [assignment.committee, person.className, person.number, person.name,
      assignment.stage === 2 ? '第2希望' : person.firstPreference === null ? '第1希望（仮割当）' : '第1希望',
      person.sourceRow, person.chairHope ? '希望する' : '希望しない'];
  });
  const leaderRows = sortedAssignments_(result.leadershipCandidates, people).map(function (assignment) {
    const person = people.get(assignment.key);
    return [assignment.committee, person.className, person.number, person.name, person.sourceRow];
  });
  const inputStats = [
    ['回答', '元回答件数', result.prepared.responseCount, '空行を除く'],
    ['回答', '重複により採用しなかった回答', result.prepared.duplicateCount, '原本は変更しない'],
    ['人数', '応募者', result.prepared.candidates.length, '最後の回答を採用'],
    ['人数', '除外者', result.prepared.candidates.length - result.prepared.eligible.length, '人数は重複計上しない'],
    ['人数', '抽選対象者', result.prepared.eligible.length, '参加条件適用後']
  ];
  [['応募者', result.prepared.candidates], ['抽選対象者', result.prepared.eligible]].forEach(function (group) {
    const label = group[0]; const members = group[1];
    LOTTERY_SETTINGS.classes.forEach(function (name) {
      inputStats.push([label + '・クラス', name, members.filter(function (p) { return p.className === name; }).length, '']);
    });
    [true, false].forEach(function (hope) {
      inputStats.push([label + '・役職希望', hope ? '希望する' : '希望しない', members.filter(function (p) { return p.chairHope === hope; }).length, '']);
    });
    ['firstPreference', 'secondPreference'].forEach(function (field) {
      LOTTERY_SETTINGS.committees.concat([null]).forEach(function (role) {
        inputStats.push([label + (field === 'firstPreference' ? '・第1希望' : '・第2希望'), role || '希望なし',
          members.filter(function (p) { return p[field] === role; }).length, '仮割当前の回答']);
      });
    });
  });
  ['OTHER_COMMITTEE', 'ORIENTATION_UNAVAILABLE', 'MEETINGS_UNAVAILABLE'].forEach(function (reason) {
    inputStats.push(['除外理由', reason, result.prepared.candidates.filter(function (p) { return p.exclusionReasons.includes(reason); }).length,
      '複数理由の重なりあり。理由別合計は除外人数と一致しない場合がある']);
  });
  ['ORIENTATION_EXCUSED', 'MEETINGS_EXCUSED'].forEach(function (reason) {
    inputStats.push(['適用した公欠', reason, result.prepared.candidates.filter(function (p) { return p.appliedExceptions.includes(reason); }).length, '']);
  });
  const resultStats = [];
  result.precheck.reasons.forEach(function (reason) { resultStats.push(['なし', '否', '事前停止', '', '', reason]); });
  result.attempts.forEach(function (attempt) {
    const adopted = attempt.accepted ? '採用' : '不採用';
    LOTTERY_SETTINGS.committees.forEach(function (name) {
      resultStats.push([attempt.number, adopted, '班の当選者', name, attempt.committeeCounts[name], '定員5人']);
      resultStats.push([attempt.number, adopted, '班の欠員', name, LOTTERY_SETTINGS.capacity - attempt.committeeCounts[name], '']);
    });
    LOTTERY_SETTINGS.classes.forEach(function (name) {
      resultStats.push([attempt.number, adopted, 'クラスの当選者', name, attempt.classCounts[name], '最低2人']);
    });
    resultStats.push([attempt.number, adopted, '役職希望者', '人数', attempt.leadershipCandidates.length, '']);
    resultStats.push([attempt.number, adopted, '役職希望者', '所属する班数', attempt.leadershipCommittees.length, '最低2班']);
    attempt.failures.forEach(function (reason) { resultStats.push([attempt.number, adopted, '不採用理由', '', '', reason]); });
  });
  const snapshots = [];
  [['元回答', captured.responses], ['運営確認', captured.operations]].forEach(function (item) {
    item[1].values.forEach(function (row, index) {
      snapshots.push([item[0], item[1].id, item[1].name, index + 1, serializedRow_(row)]);
    });
  });
  return [
    plan('結果', ['班', 'クラス', '出席番号', '名前', '当選段階', '元回答行', '役職希望'], selectedRows),
    plan('役職希望者', ['班', 'クラス', '出席番号', '名前', '元回答行'], leaderRows),
    plan('実行情報', ['項目', '内容'], metadata),
    plan('詳細ログ', ['連番', '種別', '内容（JSON）'], result.audit.map(function (event) {
      return [event.sequence, event.type, JSON.stringify(event.detail)];
    })),
    plan('元データ統計', ['項目', '区分', '人数・件数', '注記'], inputStats),
    plan('結果統計', ['試行', '採否', '項目', '区分', '人数', '条件・備考'], resultStats),
    plan('入力控え', ['入力種別', '元シートID', '元シート名', '元行番号', '値のJSON（型付き）'], snapshots)
  ];
}

function safeCell_(value) {
  if (typeof value === 'string' && /^[=+\-@\t\r]/.test(value)) return "'" + value;
  return value === null || value === undefined ? '' : value;
}

function sourceRevision_() {
  return typeof LOTTERY_BUILD_REVISION === 'string' ? LOTTERY_BUILD_REVISION : 'manual-build';
}

function hashResult_(status, acceptedAttempt, assignments, leaders, digest) {
  function ordered(items) {
    return items.map(function (a) { return { key: a.key, committee: a.committee, stage: a.stage }; })
      .sort(function (a, b) {
        const first = a.key.split(':'); const second = b.key.split(':');
        return LOTTERY_SETTINGS.committees.indexOf(a.committee) - LOTTERY_SETTINGS.committees.indexOf(b.committee) ||
          LOTTERY_SETTINGS.classes.indexOf(first[0]) - LOTTERY_SETTINGS.classes.indexOf(second[0]) ||
          Number(first[1]) - Number(second[1]);
      });
  }
  return hashText_(JSON.stringify({ schema: 'lottery-result/v1', status: status,
    acceptedAttempt: acceptedAttempt, assignments: ordered(assignments), leadershipCandidates: ordered(leaders) }), digest);
}

function makeRunRecord_(captured, result, context) {
  return {
    schema: 'lottery-history/v1', runId: context.runId, startedAt: context.startedAt,
    version: result.version, algorithmVersion: LOTTERY_ALGORITHM_VERSION,
    sourceRevision: context.sourceRevision || sourceRevision_(), seed: result.seed, status: result.status,
    acceptedAttempt: result.acceptedAttempt, attempts: result.attempts.length,
    responsesHash: snapshotHash_(captured.responses.values, context.digest),
    operationsHash: snapshotHash_(captured.operations.values, context.digest),
    resultHash: hashResult_(result.status, result.acceptedAttempt, result.assignments, result.leadershipCandidates, context.digest),
    auditHash: hashText_(JSON.stringify(result.audit), context.digest),
    randomPosition: result.randomPosition, finalRandomState: result.finalRandomState
  };
}

function validRunRecord_(record) {
  return record && record.schema === 'lottery-history/v1' &&
    typeof record.runId === 'string' && record.runId.length > 0 &&
    ['startedAt', 'version', 'algorithmVersion', 'sourceRevision'].every(function (field) {
      return typeof record[field] === 'string' && record[field].length > 0;
    }) &&
    Number.isInteger(record.attempts) && record.attempts >= 0 && record.attempts <= LOTTERY_SETTINGS.maximumAttempts &&
    Number.isInteger(record.randomPosition) && record.randomPosition >= 0 &&
    /^[0-9a-f]{64}$/.test(record.finalRandomState || '') &&
    (record.status === 'ACCEPTED' ? Number.isInteger(record.acceptedAttempt) && record.acceptedAttempt >= 1 &&
      record.acceptedAttempt <= record.attempts : record.acceptedAttempt === null) &&
    /^[0-9a-f]{64}$/.test(record.seed || '') &&
    ['responsesHash', 'operationsHash', 'resultHash', 'auditHash'].every(function (field) {
      return /^[0-9a-f]{64}$/.test(record[field] || '');
    }) &&
    ['ACCEPTED', 'PRECHECK_FAILED', 'ATTEMPT_LIMIT_REACHED'].includes(record.status);
}

function readOutputNote_(sheet) {
  let note;
  try { note = JSON.parse(sheet.getRange(1, 1).getNote() || 'null'); } catch (error) { return null; }
  return note && note.owner === LOTTERY_OUTPUT_OWNER ? note : null;
}

function markOutput_(sheet, label, record, phase) {
  sheet.getRange(1, 1).setNote(JSON.stringify({
    owner: LOTTERY_OUTPUT_OWNER, kind: 'output', schema: 2, label: label, phase: phase, record: record
  }));
  if (label === '実行情報') {
    sheet.getRange(3, 2).setValue(phase === 'published' ? displayStatus_(record.status) :
      phase === 'failed' ? '出力失敗・当選未確定' : '保存中・当選未確定');
  }
}

function historyRows_(sheet) {
  if (!sheet) return [];
  const note = readOutputNote_(sheet);
  if (!note || note.kind !== 'history') throw new Error('「抽選_実行履歴」がプログラム管理外のシートと衝突しています。');
  const values = sheet.getDataRange().getValues();
  if (values.every(function (row) { return row.every(function (v) { return text_(v) === ''; }); })) return [];
  if (JSON.stringify(values[0].slice(0, LOTTERY_HISTORY_HEADER.length)) !== JSON.stringify(LOTTERY_HISTORY_HEADER)) {
    throw new Error('実行履歴の見出しが変更されています。履歴を保全したまま確認してください。');
  }
  return values.slice(1).filter(function (row) { return row.some(function (v) { return text_(v) !== ''; }); })
    .map(function (row) {
      let record;
      try { record = JSON.parse(row[7]); } catch (error) { throw new Error('実行履歴の再検証情報が不正です。'); }
      if (!validRunRecord_(record) || text_(row[0]) !== record.startedAt || text_(row[1]) !== record.runId ||
          text_(row[3]) !== displayStatus_(record.status) || text_(row[4]) !== text_(record.acceptedAttempt || 'なし') ||
          row[5] !== record.resultHash || row[6] !== record.auditHash) {
        throw new Error('実行履歴のハッシュと再検証情報が一致しません。');
      }
      return { event: text_(row[2]), record: record };
    });
}

function appendHistory_(spreadsheet, entries) {
  if (!entries.length) return;
  let sheet = spreadsheet.getSheetByName(LOTTERY_HISTORY_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(LOTTERY_HISTORY_NAME);
    sheet.getRange(1, 1).setNote(JSON.stringify({ owner: LOTTERY_OUTPUT_OWNER, kind: 'history', schema: 1 }));
    sheet.getRange(1, 1, 1, LOTTERY_HISTORY_HEADER.length).setValues([LOTTERY_HISTORY_HEADER.slice()]);
    formatOutput_(sheet, '実行履歴', 1, LOTTERY_HISTORY_HEADER.length);
  } else {
    historyRows_(sheet);
    if (sheet.getRange(1, 1, 1, LOTTERY_HISTORY_HEADER.length).getValues()[0].every(function (v) { return text_(v) === ''; })) {
      sheet.getRange(1, 1, 1, LOTTERY_HISTORY_HEADER.length).setValues([LOTTERY_HISTORY_HEADER.slice()]);
    }
  }
  const rows = entries.map(function (entry) {
    const r = entry.record;
    return [r.startedAt, r.runId, entry.event, displayStatus_(r.status), r.acceptedAttempt || 'なし',
      r.resultHash, r.auditHash, JSON.stringify(r)].map(safeCell_);
  });
  const first = sheet.getLastRow() + 1;
  if (sheet.getMaxRows() < first + rows.length - 1) {
    sheet.insertRowsAfter(sheet.getMaxRows(), first + rows.length - 1 - sheet.getMaxRows());
  }
  sheet.getRange(first, 1, rows.length, LOTTERY_HISTORY_HEADER.length).setNumberFormat('@').setValues(rows);
}

function legacyOutput_(sheet) {
  const values = sheet.getDataRange().getValues();
  const keys = ['実行ID', '状態', '方式の版', 'シード', '乱数方式', '元回答SHA256', '運営確認SHA256',
    '採用試行', '実施試行数', '消費した乱数の個数', '最終乱数状態', '実行開始日時', 'タイムゾーン',
    '元回答シート・範囲', '運営確認シート・範囲'];
  if (values.length < 17 || !keys.every(function (key, i) { return values[i][0] === key; })) return null;
  const meta = {};
  keys.forEach(function (key, i) { meta[key] = values[i][1]; });
  if (meta['方式の版'] !== '1.0.0' || meta['乱数方式'] !== LOTTERY_PRNG) return null;
  const headers = {
    '結果': ['班', 'クラス', '出席番号', '名前', '当選段階', '元回答行', '役職希望'],
    '役職希望者': ['班', 'クラス', '出席番号', '名前', '元回答行'],
    '詳細ログ': ['連番', '種別', '内容（JSON）'],
    '元データ統計': ['項目', '区分', '人数・件数', '注記'],
    '結果統計': ['試行', '採否', '項目', '区分', '人数', '条件・備考'],
    '入力控え': ['入力種別', '元シートID', '元シート名', '元行番号', '値のJSON（型付き）']
  };
  const labels = Object.keys(headers).filter(function (label) {
    return JSON.stringify(values[16].slice(0, headers[label].length)) === JSON.stringify(headers[label]);
  });
  return labels.length === 1 ? { sheet: sheet, label: labels[0], meta: meta, body: values.slice(17) } : null;
}

function legacyRecord_(items, digest) {
  if (items.length !== 6 || new Set(items.map(function (item) { return item.label; })).size !== 6) return null;
  const meta = items[0].meta;
  if (!items.every(function (item) { return JSON.stringify(item.meta) === JSON.stringify(meta); })) {
    throw new Error('旧出力の実行情報が同じ実行内で一致しません。過去の出力を確認してください。');
  }
  const statuses = { '成立': 'ACCEPTED', '事前条件不成立': 'PRECHECK_FAILED', '試行上限到達・選出未成立': 'ATTEMPT_LIMIT_REACHED' };
  const status = statuses[meta['状態']];
  if (!status) return null;
  const resultItem = items.find(function (item) { return item.label === '結果'; });
  const leaderItem = items.find(function (item) { return item.label === '役職希望者'; });
  const auditItem = items.find(function (item) { return item.label === '詳細ログ'; });
  const assignments = resultItem.body.filter(function (row) { return text_(row[0]); }).map(function (row) {
    const person = personKey_(row[1], row[2], '旧結果');
    const stage = text_(row[4]).startsWith('第2希望') ? 2 : 1;
    if (!LOTTERY_SETTINGS.committees.includes(row[0])) throw new Error('旧結果の班名が不正です。');
    return { key: person.key, committee: row[0], stage: stage };
  });
  const leaders = leaderItem.body.filter(function (row) { return text_(row[0]); }).map(function (row) {
    const key = personKey_(row[1], row[2], '旧役職希望者').key;
    const assignment = assignments.find(function (a) { return a.key === key && a.committee === row[0]; });
    if (!assignment) throw new Error('旧役職希望者と旧結果が一致しません。');
    return assignment;
  });
  const audit = auditItem.body.filter(function (row) { return text_(row[0]); }).map(function (row) {
    return { sequence: Number(row[0]), type: row[1], detail: JSON.parse(row[2]) };
  });
  const acceptedAttempt = status === 'ACCEPTED' ? Number(meta['採用試行']) : null;
  const record = {
    schema: 'lottery-history/v1', runId: text_(meta['実行ID']), startedAt: text_(meta['実行開始日時']),
    version: '1.0.0', algorithmVersion: '1.0.0', sourceRevision: 'legacy-1.0.0 (Git baseline 17c3cd72f5c677f762c2de15fc235edd5a116402)',
    seed: text_(meta['シード']), status: status, acceptedAttempt: acceptedAttempt, attempts: Number(meta['実施試行数']),
    responsesHash: text_(meta['元回答SHA256']), operationsHash: text_(meta['運営確認SHA256']),
    resultHash: hashResult_(status, acceptedAttempt, assignments, leaders, digest),
    auditHash: hashText_(JSON.stringify(audit), digest),
    randomPosition: Number(meta['消費した乱数の個数']), finalRandomState: text_(meta['最終乱数状態'])
  };
  if (!validRunRecord_(record)) throw new Error('旧出力の再検証情報が不正です。');
  return record;
}

function collectOutputRuns_(spreadsheet, protectedIds, digest) {
  const groups = new Map(); const legacy = new Map(); const warnings = []; const reserves = [];
  spreadsheet.getSheets().forEach(function (sheet) {
    if (protectedIds.has(sheet.getSheetId()) || !sheet.getName().startsWith(LOTTERY_SETTINGS.outputPrefix) ||
        sheet.getName() === LOTTERY_HISTORY_NAME) return;
    const note = readOutputNote_(sheet);
    if (note && note.kind === 'reserve') {
      if (!validRunRecord_(note.record) || !LOTTERY_RESERVE_LABELS.includes(note.label) ||
          !note.reserve || typeof note.reserve.id !== 'string' || !/^[0-9a-f]{64}$/.test(note.reserve.listHash || '')) {
        throw new Error('補欠シートの管理情報が不正です。自動削除せず確認してください。');
      }
      reserves.push({ sheet: sheet, label: note.label, record: note.record, reserve: note.reserve });
      return;
    }
    if (note && note.kind === 'output') {
      if (!validRunRecord_(note.record) || !LOTTERY_OUTPUT_LABELS.includes(note.label)) {
        throw new Error('出力シートの管理情報が不正です。自動削除せず確認してください。');
      }
      const id = note.record.runId;
      if (!groups.has(id)) groups.set(id, { record: note.record, items: [], legacy: false });
      const group = groups.get(id);
      if (JSON.stringify(group.record) !== JSON.stringify(note.record)) throw new Error('出力間の管理情報が一致しません。');
      group.items.push({ sheet: sheet, label: note.label });
      return;
    }
    const old = legacyOutput_(sheet);
    if (old) {
      const id = text_(old.meta['実行ID']);
      if (!legacy.has(id)) legacy.set(id, []);
      legacy.get(id).push(old);
    } else {
      warnings.push('管理対象と確認できない出力名のシートを保持: ' + sheet.getName());
    }
  });
  legacy.forEach(function (items, id) {
    const record = legacyRecord_(items, digest);
    if (!record) { warnings.push('6種類がそろわない、または未確定の旧出力を保持: ' + id); return; }
    if (groups.has(id)) throw new Error('旧版と新版で実行IDが重複しています。');
    groups.set(id, { record: record, items: items, legacy: true });
  });
  reserves.forEach(function (item) {
    if (!groups.has(item.record.runId)) groups.set(item.record.runId, { record: item.record, items: [], legacy: false });
    const group = groups.get(item.record.runId);
    if (JSON.stringify(group.record) !== JSON.stringify(item.record)) throw new Error('補欠と本抽選の管理情報が一致しません。');
    group.reserves = (group.reserves || []).concat([item]);
  });
  groups.forEach(function (group) {
    if (new Set(group.items.map(function (item) { return item.label; })).size !== group.items.length) {
      throw new Error('同一実行に同じ種類の出力が複数あります。複製したシートを保全したまま確認してください。');
    }
    const keys = (group.reserves || []).map(function (item) { return item.reserve.id + '/' + item.label; });
    if (new Set(keys).size !== keys.length) throw new Error('同じ補欠シートの複製があります。自動削除せず確認してください。');
  });
  return { groups: Array.from(groups.values()), warnings: warnings };
}

function moveOutputFront_(spreadsheet, items, status) {
  const order = status === 'ACCEPTED' ? ['結果', '役職希望者', '実行情報'] : ['実行情報', '結果', '役職希望者'];
  order.forEach(function (label, index) {
    const item = items.find(function (entry) { return entry.label === label; });
    if (item) { spreadsheet.setActiveSheet(item.sheet); spreadsheet.moveActiveSheet(index + 1); }
  });
  const first = items.find(function (item) { return item.label === order[0]; });
  if (first) spreadsheet.setActiveSheet(first.sheet);
}

function formatOutput_(sheet, label, height, width) {
  sheet.setFrozenRows(1);
  const head = sheet.getRange(1, 1, 1, width);
  head.setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff');
  sheet.setRowHeight(1, 30);
  for (let column = 1; column <= width; column += 1) sheet.setColumnWidth(column, 130);
  if (label === '結果' || label === '役職希望者' || label === '補欠名簿') {
    sheet.setColumnWidth(1, 180); sheet.setColumnWidth(2, 70); sheet.setColumnWidth(3, 80); sheet.setColumnWidth(4, 160);
    if (height > 1) sheet.getRange(1, 1, height, width).createFilter();
    if (label === '補欠名簿') { sheet.setColumnWidth(4, 80); sheet.setColumnWidth(5, 160); sheet.setColumnWidth(width, 260); }
  } else if (label === '実行情報') {
    sheet.setColumnWidth(1, 220); sheet.setColumnWidth(2, 560);
  } else {
    sheet.setColumnWidth(width, 460);
  }
  const clip = typeof SpreadsheetApp !== 'undefined' && SpreadsheetApp.WrapStrategy ? SpreadsheetApp.WrapStrategy.CLIP : 'CLIP';
  sheet.getRange(1, 1, Math.max(height, 1), width).setWrapStrategy(clip);
}

/** 控えの全文が当時のハッシュと一致する場合だけ、差分の比較元として使う。 */
function storedInputForDiff_(group, kind, expectedHash, digest) {
  const item = group.items.find(function (entry) { return entry.label === '入力控え'; });
  if (!item) return null;
  try {
    let body;
    if (group.legacy) body = item.body;
    else {
      const values = item.sheet.getDataRange().getValues();
      const header = ['入力種別', '元シートID', '元シート名', '元行番号', '値のJSON（型付き）'];
      if (JSON.stringify(values[0].slice(0, header.length)) !== JSON.stringify(header)) return null;
      body = values.slice(1);
    }
    const rows = body.filter(function (row) { return row[0] === kind; })
      .sort(function (a, b) { return Number(a[3]) - Number(b[3]); });
    if (!rows.length || !rows.every(function (row, index) {
      return Number(row[3]) === index + 1 && row[2] === rows[0][2] && typeof row[4] === 'string';
    })) return null;
    const serialized = rows.map(function (row) { return row[4]; });
    if (hashText_(JSON.stringify(serialized), digest) !== expectedHash) return null;
    const cells = serialized.map(function (row) {
      const parsed = JSON.parse(row);
      if (!Array.isArray(parsed) || !parsed.every(function (cell) {
        return cell && ['string', 'number', 'boolean', 'date'].includes(cell.type) &&
          (cell.type === 'date' ? typeof cell.value === 'string' : typeof cell.value === cell.type);
      })) throw new Error('入力控えのセル形式が不正です。');
      return parsed;
    });
    return { name: text_(rows[0][2]), cells: cells, runId: group.record.runId };
  } catch (error) { return null; }
}

function diffText_(value, maximum) {
  const text = JSON.stringify(value);
  const characters = Array.from(text === undefined ? String(value) : text);
  return characters.length > maximum ? characters.slice(0, maximum).join('') + '…（省略）' : characters.join('');
}

function diffCellText_(cell) {
  if (!cell) return '（セルなし）';
  const types = { string: '文字列', number: '数値', boolean: '真偽値', date: '日時・UTC' };
  return diffText_(cell.value, 160) + '（' + types[cell.type] + '）';
}

function diffColumnName_(column) {
  let name = '';
  for (let value = column + 1; value > 0; value = Math.floor((value - 1) / 26)) {
    name = String.fromCharCode(65 + (value - 1) % 26) + name;
  }
  return name;
}

function inputDiffLines_(previous, current, kind) {
  const before = previous.cells;
  const after = current.values.map(function (row) { return JSON.parse(serializedRow_(row)); });
  function headerIndex(rows) {
    if (kind === '元回答') return 0;
    return rows.findIndex(function (row) {
      try { operationsColumns_(row.map(function (cell) { return cell.value; })); return true; }
      catch (error) { return false; }
    });
  }
  const oldHeaderIndex = headerIndex(before); const newHeaderIndex = headerIndex(after);
  const oldHeader = before[oldHeaderIndex] || []; const newHeader = after[newHeaderIndex] || [];
  const lines = []; let count = 0;
  for (let row = 0; row < Math.max(before.length, after.length); row += 1) {
    const oldRow = before[row] || []; const newRow = after[row] || [];
    for (let column = 0; column < Math.max(oldRow.length, newRow.length); column += 1) {
      const oldCell = oldRow[column]; const newCell = newRow[column];
      if (JSON.stringify(oldCell) === JSON.stringify(newCell)) continue;
      count += 1;
      if (lines.length >= 20) continue;
      const oldLabel = oldHeader[column] ? text_(oldHeader[column].value) : '';
      const newLabel = newHeader[column] ? text_(newHeader[column].value) : '';
      const label = oldLabel && newLabel && oldLabel !== newLabel ? oldLabel + ' → ' + newLabel : newLabel || oldLabel || '項目名なし';
      const position = diffColumnName_(column) + (row + 1);
      const change = !oldCell ? '追加' : !newCell ? '削除' : '変更';
      lines.push('「' + current.name + '」' + (row + 1) + '行目 ' + position + '［' +
        diffText_(label, 100) + '］' + (row === oldHeaderIndex || row === newHeaderIndex ? '（見出し）' : '') +
        ' ' + change + ': ' + diffCellText_(oldCell) + ' → ' + diffCellText_(newCell));
    }
  }
  const oldWidth = Math.max.apply(null, before.map(function (row) { return row.length; }));
  const newWidth = Math.max.apply(null, after.map(function (row) { return row.length; }));
  const heading = kind + '（比較元の実行ID: ' + previous.runId + '、入力控えの元シート: ' + previous.name + '）\n' +
    '読取範囲: ' + before.length + '行×' + oldWidth + '列 → ' + after.length + '行×' + newWidth + '列。差分 ' + count + 'セル。';
  if (count > lines.length) lines.push('ほか ' + (count - lines.length) + 'セル。先頭20件のみ表示しています。全文は入力控えと照合してください。');
  return [heading].concat(lines);
}

function assertUnchangedInputs_(record, groups, history, captured, digest) {
  const previous = groups.map(function (group) { return group.record; })
    .concat(history.map(function (entry) { return entry.record; }));
  const lines = []; const checked = new Set();
  previous.forEach(function (old) {
    [['元回答', 'responses', 'responsesHash'], ['運営確認', 'operations', 'operationsHash']].forEach(function (source) {
      const kind = source[0]; const field = source[2];
      if (old[field] === record[field] || checked.has(kind + old[field])) return;
      checked.add(kind + old[field]);
      let stored = null;
      if (captured) {
        groups.filter(function (group) { return group.record[field] === old[field]; }).some(function (group) {
          stored = storedInputForDiff_(group, kind, old[field], digest);
          return stored !== null;
        });
      }
      if (stored) lines.push.apply(lines, inputDiffLines_(stored, captured[source[1]], kind));
      else lines.push(kind + '（過去の実行ID: ' + old.runId + '）: 当時のハッシュと一致する入力控えが見つからないため、行・値の差分を特定できません。\n' +
        '過去SHA256: ' + old[field] + '\n現在SHA256: ' + record[field]);
    });
  });
  if (lines.length) throw new Error('元回答または運営確認が過去の実行から変わっています。旧出力の整理を停止します。\n' +
    '以下は行・列の位置で比較した変更前 → 変更後です。行・列の挿入や削除では後続セルも差分になります。\n' + lines.join('\n'));
}

function writeOutputPlans_(spreadsheet, plans, runId, status, flush, options) {
  flush = flush || function () {};
  options = options || {};
  const protectedIds = new Set(options.protectedIds || []);
  const history = historyRows_(spreadsheet.getSheetByName(LOTTERY_HISTORY_NAME));
  const inventory = collectOutputRuns_(spreadsheet, protectedIds, options.digest);
  let record = Object.assign({}, plans[0].record);
  assertUnchangedInputs_(record, inventory.groups, history, options.captured, options.digest);
  const usedIds = new Set(history.map(function (entry) { return entry.record.runId; })
    .concat(inventory.groups.map(function (group) { return group.record.runId; })));
  let suffix = 1;
  while (usedIds.has(record.runId)) record.runId = runId + '_' + suffix++;
  const oldSheets = inventory.groups.flatMap(function (group) {
    return group.items.concat(group.reserves || []).map(function (item) { return item.sheet; });
  });
  const oldIds = new Set(oldSheets.map(function (sheet) { return sheet.getSheetId(); }));
  LOTTERY_OUTPUT_LABELS.forEach(function (label) {
    const existing = spreadsheet.getSheetByName(LOTTERY_SETTINGS.outputPrefix + label);
    if (existing && !oldIds.has(existing.getSheetId())) throw new Error('固定の出力名が管理対象外シートと衝突しています: ' + label);
  });
  const prepared = plans.map(function (plan) {
    const width = Math.max.apply(null, plan.rows.map(function (row) { return row.length; }));
    const rows = plan.rows.map(function (row) {
      const values = row.map(safeCell_);
      if (plan.label === '実行情報' && row[0] === '実行ID') values[1] = record.runId;
      while (values.length < width) values.push('');
      values.forEach(function (value) {
        if (typeof value === 'string' && value.length > 45000) throw new Error('入力またはログが1セルの保存上限を超えます。');
      });
      return values;
    });
    return { label: plan.label, rows: rows, width: width,
      temporaryName: LOTTERY_SETTINGS.outputPrefix + '保存中_' + record.runId + '_' + plan.label };
  });
  const created = []; const renamedOld = []; const renamedNew = [];
  let published = false;
  try {
    prepared.forEach(function (plan) {
      const sheet = spreadsheet.insertSheet(plan.temporaryName);
      created.push({ sheet: sheet, plan: plan });
      sheet.hideSheet();
      sheet.getRange(1, 1).setNote(JSON.stringify({ owner: LOTTERY_OUTPUT_OWNER, kind: 'output',
        schema: 2, label: plan.label, phase: 'staging', record: record }));
      if (sheet.getMaxRows() < plan.rows.length) sheet.insertRowsAfter(sheet.getMaxRows(), plan.rows.length - sheet.getMaxRows());
      if (sheet.getMaxColumns() < plan.width) sheet.insertColumnsAfter(sheet.getMaxColumns(), plan.width - sheet.getMaxColumns());
      for (let offset = 0; offset < plan.rows.length; offset += 500) {
        const chunk = plan.rows.slice(offset, offset + 500);
        sheet.getRange(offset + 1, 1, chunk.length, plan.width).setNumberFormat('@').setValues(chunk);
      }
      formatOutput_(sheet, plan.label, plan.rows.length, plan.width);
    });
    flush();
    const entries = inventory.groups.filter(function (group) {
      return !history.some(function (entry) { return entry.record.runId === group.record.runId; });
    }).map(function (group) { return { event: group.legacy ? '旧版移行' : '旧出力記録', record: group.record }; });
    entries.push({ event: '保存準備', record: record });
    appendHistory_(spreadsheet, entries); flush();
    oldSheets.forEach(function (sheet) {
      const oldName = sheet.getName();
      if (LOTTERY_OUTPUT_LABELS.some(function (label) { return oldName === LOTTERY_SETTINGS.outputPrefix + label; })) {
        let replacement = LOTTERY_SETTINGS.outputPrefix + '退避_' + sheet.getSheetId() + '_' + record.runId;
        while (spreadsheet.getSheetByName(replacement)) replacement += '_';
        sheet.setName(replacement); renamedOld.push({ sheet: sheet, name: oldName });
      }
    });
    created.forEach(function (item) {
      item.sheet.setName(LOTTERY_SETTINGS.outputPrefix + item.plan.label); renamedNew.push(item);
      markOutput_(item.sheet, item.plan.label, record, 'staging');
    });
    flush();
    created.filter(function (item) { return item.plan.label !== '実行情報'; }).forEach(function (item) {
      markOutput_(item.sheet, item.plan.label, record, 'published');
    });
    const summary = created.find(function (item) { return item.plan.label === '実行情報'; });
    markOutput_(summary.sheet, summary.plan.label, record, 'published'); flush();
    created.forEach(function (item) { item.sheet.showSheet(); });
    moveOutputFront_(spreadsheet, created.map(function (item) { return { sheet: item.sheet, label: item.plan.label }; }), status);
    appendHistory_(spreadsheet, [{ event: '保存完了', record: record }]); flush();
    published = true;
  } catch (error) {
    renamedNew.slice().reverse().forEach(function (item) {
      try { item.sheet.setName(item.plan.temporaryName); } catch (ignored) { /* 旧出力を削除しない */ }
    });
    renamedOld.slice().reverse().forEach(function (item) {
      try { item.sheet.setName(item.name); } catch (ignored) { /* 退避状態を保持 */ }
    });
    created.forEach(function (item) {
      try { markOutput_(item.sheet, item.plan.label, record, 'failed'); item.sheet.showSheet(); } catch (ignored) { /* 原因を保持 */ }
    });
    try {
      const oldFront = LOTTERY_OUTPUT_LABELS.map(function (label) {
        return { label: label, sheet: spreadsheet.getSheetByName(LOTTERY_SETTINGS.outputPrefix + label) };
      }).filter(function (item) { return item.sheet && oldIds.has(item.sheet.getSheetId()); });
      if (oldFront.length) moveOutputFront_(spreadsheet, oldFront, readOutputNote_(oldFront[0].sheet).record.status);
    } catch (ignored) { /* 旧出力は保持したまま、元の保存エラーを伝える */ }
    try { appendHistory_(spreadsheet, [{ event: '保存失敗', record: record }]); flush(); } catch (ignored) { /* 旧出力を保持 */ }
    throw new Error('出力の保存に失敗しました。旧出力は削除していません。保存中・出力失敗のシートを当選結果として使用しないでください。');
  }
  const warnings = inventory.warnings.slice();
  if (published) {
    oldSheets.forEach(function (sheet) {
      if (protectedIds.has(sheet.getSheetId()) || !oldIds.has(sheet.getSheetId())) throw new Error('入力シートは削除できません。');
      try { spreadsheet.deleteSheet(sheet); } catch (error) { warnings.push('旧出力を削除できず保持: ' + sheet.getName()); }
    });
  }
  return { runId: record.runId, sheetNames: created.map(function (item) { return item.sheet.getName(); }), warnings: warnings };
}

function executeLotteryOnSpreadsheet_(spreadsheet, options) {
  const captured = readInputSheets_(spreadsheet);
  const result = runLotteryEngine_(captured.responses.data, captured.operations.data,
    { seed: options.seed, digest: options.digest, locations: captured });
  const context = Object.assign({}, options);
  context.record = makeRunRecord_(captured, result, context);
  const plans = buildOutputPlans_(captured, result, context);
  const output = writeOutputPlans_(spreadsheet, plans, options.runId, result.status, options.flush, {
    digest: options.digest, captured: captured, protectedIds: [captured.responses.id, captured.operations.id]
  });
  return {
    status: result.status, runId: output.runId, sheetNames: output.sheetNames, warnings: output.warnings,
    acceptedAttempt: result.acceptedAttempt, attempts: result.attempts.length,
    applicants: result.prepared.candidates.length, eligible: result.prepared.eligible.length,
    selected: result.assignments.length
  };
}

/** 本抽選の乱数や結果に触れず、未選出者の第1・第2希望ごとに補欠順位を作る。 */
function buildReserveList_(prepared, assignments, fixedFirstByKey, seed, digest) {
  if (typeof seed !== 'string' || !/^[0-9a-fA-F]{64}$/.test(seed)) throw new Error('補欠の元シードは64桁の16進数が必要です。');
  const selected = new Set(assignments.map(function (item) { return item.key; }));
  const audit = []; const candidates = []; const unassigned = []; const seeds = {};
  function record(type, detail) { audit.push({ sequence: audit.length + 1, type: type, detail: detail }); }
  prepared.candidates.slice().sort(comparePeople_).forEach(function (person) {
    if (selected.has(person.key)) return;
    const blocked = person.exclusionReasons.filter(function (reason) { return reason !== 'MEETINGS_UNAVAILABLE'; });
    record('RESERVE_ELIGIBILITY', { key: person.key, included: blocked.length === 0, exclusions: blocked });
    if (blocked.length) return;
    const first = person.firstPreference || fixedFirstByKey[person.key] || null;
    const memberships = [];
    if (first) memberships.push({ committee: first, preference: person.firstPreference ? '第1希望' : '第1希望（仮割当）' });
    if (person.secondPreference) {
      const existing = memberships.find(function (item) { return item.committee === person.secondPreference; });
      if (existing) existing.preference += '・第2希望';
      else memberships.push({ committee: person.secondPreference, preference: '第2希望' });
    }
    const notes = [];
    if (person.exclusionReasons.includes('MEETINGS_UNAVAILABLE')) notes.push('繰上げ前に参加可否の確認が必要');
    if (person.firstPreference === null) notes.push('第1希望は「特になし」');
    const candidate = { key: person.key, className: person.className, number: person.number, name: person.name,
      sourceRow: person.sourceRow, committee: null, preference: '', rank: null, meetings: person.meetings,
      category: person.exclusionReasons.includes('MEETINGS_UNAVAILABLE') ? '準備・打ち合わせ参加不可' : '落選',
      note: notes.join('／') };
    if (memberships.length) memberships.forEach(function (membership) { candidates.push(Object.assign({}, candidate, membership)); });
    else { candidate.note = notes.concat(['本抽選の仮割当なし・希望班未割当']).join('／'); unassigned.push(candidate); }
  });
  const ranked = [];
  LOTTERY_SETTINGS.committees.forEach(function (committee) {
    const derived = hashText_(JSON.stringify(['committee-reserve/v1', seed.toLowerCase(), committee]), digest);
    seeds[committee] = derived;
    const pool = candidates.filter(function (person) { return person.committee === committee; });
    const random = createRandom_(derived, function (type, detail) {
      record('RESERVE_' + type, Object.assign({ committee: committee }, detail));
    }, digest);
    const shuffled = pool.slice();
    for (let i = shuffled.length - 1; i > 0; i -= 1) {
      const j = random.nextInt(i + 1, 'reserve/' + committee + '/swap=' + i);
      const item = shuffled[i]; shuffled[i] = shuffled[j]; shuffled[j] = item;
    }
    shuffled.forEach(function (person, index) { ranked.push(Object.assign({}, person, { rank: index + 1 })); });
    record('RESERVE_RANKING', { committee: committee, seed: derived, pool: pool.map(function (p) { return p.key; }),
      ranked: shuffled.map(function (p) { return p.key; }), randomPosition: random.position(), finalRandomState: random.state() });
  });
  return { ranked: ranked, unassigned: unassigned, seeds: seeds, audit: audit,
    listHash: hashText_(JSON.stringify({ version: LOTTERY_RESERVE_VERSION,
      ranked: ranked.map(function (p) { return { key: p.key, committee: p.committee, rank: p.rank,
        category: p.category, preference: p.preference, note: p.note }; }),
      unassigned: unassigned.map(function (p) { return { key: p.key, note: p.note }; }) }), digest),
    auditHash: hashText_(JSON.stringify(audit), digest) };
}

function readSavedLotteryForReserves_(spreadsheet, captured, digest) {
  const resultSheet = spreadsheet.getSheetByName(LOTTERY_SETTINGS.outputPrefix + '結果');
  const note = resultSheet && readOutputNote_(resultSheet);
  if (!note || note.kind !== 'output' || note.label !== '結果' || note.phase !== 'published' ||
      !validRunRecord_(note.record) || note.record.status !== 'ACCEPTED') {
    throw new Error('保存が完了した成立済みの「抽選_結果」が必要です。本抽選は自動実行しません。');
  }
  const base = note.record;
  if (base.algorithmVersion !== LOTTERY_ALGORITHM_VERSION) throw new Error('本抽選のアルゴリズム版に対応したコードで補欠名簿を作成してください。');
  const history = historyRows_(spreadsheet.getSheetByName(LOTTERY_HISTORY_NAME));
  if (!history.some(function (entry) { return entry.event === '保存完了' && entry.record.runId === base.runId &&
      JSON.stringify(entry.record) === JSON.stringify(base); }) || history.some(function (entry) {
    return entry.event === '保存失敗' && entry.record.runId === base.runId;
  })) throw new Error('本抽選の保存完了を実行履歴で確認できません。補欠名簿を作成しません。');
  const inventory = collectOutputRuns_(spreadsheet, new Set([captured.responses.id, captured.operations.id]), digest);
  const group = inventory.groups.find(function (item) { return item.record.runId === base.runId; });
  if (!group || !LOTTERY_OUTPUT_LABELS.every(function (label) {
    const item = group.items.find(function (entry) { return entry.label === label; });
    return item && readOutputNote_(item.sheet).phase === 'published';
  })) throw new Error('本抽選の出力7種類の保存状態が一致しません。補欠名簿を作成しません。');
  const current = Object.assign({}, base, { responsesHash: snapshotHash_(captured.responses.values, digest),
    operationsHash: snapshotHash_(captured.operations.values, digest) });
  assertUnchangedInputs_(current, inventory.groups, history, captured, digest);
  const prepared = prepareLotteryInput_(captured.responses.data, captured.operations.data, captured);
  const people = new Map(prepared.candidates.map(function (p) { return [p.key, p]; }));
  const values = resultSheet.getDataRange().getValues();
  if (JSON.stringify(values[0]) !== JSON.stringify(['班', 'クラス', '出席番号', '名前', '当選段階', '元回答行', '役職希望'])) {
    throw new Error('抽選結果の見出しが変更されています。補欠名簿を作成しません。');
  }
  const assignments = values.slice(1).filter(function (row) { return row.some(function (value) { return text_(value); }); })
    .map(function (row) {
      const key = personKey_(row[1], row[2], '保存済み抽選結果').key;
      if (!people.has(key) || !LOTTERY_SETTINGS.committees.includes(row[0]) ||
          !['第1希望', '第1希望（仮割当）', '第2希望'].includes(row[4])) throw new Error('保存済み抽選結果の班・本人・当選段階が不正です。');
      return { key: key, committee: row[0], stage: row[4] === '第2希望' ? 2 : 1 };
    });
  const leaders = assignments.filter(function (item) { return people.get(item.key).chairHope; });
  if (new Set(assignments.map(function (item) { return item.key; })).size !== assignments.length ||
      hashResult_(base.status, base.acceptedAttempt, assignments, leaders, digest) !== base.resultHash) {
    throw new Error('保存済みの当選者・所属班が本抽選の結果ハッシュと一致しません。');
  }
  const log = group.items.find(function (item) { return item.label === '詳細ログ'; }).sheet.getDataRange().getValues();
  if (JSON.stringify(log[0]) !== JSON.stringify(['連番', '種別', '内容（JSON）'])) throw new Error('本抽選ログの見出しが不正です。');
  let audit;
  try { audit = log.slice(1).map(function (row) { return { sequence: row[0], type: row[1], detail: JSON.parse(row[2]) }; }); }
  catch (error) { throw new Error('本抽選ログのJSONが不正です。'); }
  if (hashText_(JSON.stringify(audit), digest) !== base.auditHash) throw new Error('本抽選ログが保存時のハッシュと一致しません。');
  const fixedFirstByKey = {};
  audit.filter(function (event) { return event.type === 'FIXED_FIRST_PREFERENCE'; }).forEach(function (event) {
    fixedFirstByKey[event.detail.key] = event.detail.committee;
  });
  return { record: base, prepared: prepared, assignments: assignments, fixedFirstByKey: fixedFirstByKey, inventory: inventory, history: history };
}

function writeReservePlans_(spreadsheet, source, result, options) {
  const flush = options.flush || function () {};
  const old = source.inventory.groups.flatMap(function (group) { return group.reserves || []; });
  const used = new Set(old.map(function (item) { return item.reserve.id; }).concat(source.history
    .filter(function (entry) { return entry.record.reserve; }).map(function (entry) { return entry.record.reserve.id; })));
  let id = options.runId; let suffix = 1;
  while (used.has(id)) id = options.runId + '_' + suffix++;
  LOTTERY_RESERVE_LABELS.forEach(function (label) {
    const sheet = spreadsheet.getSheetByName(LOTTERY_SETTINGS.outputPrefix + label);
    if (sheet && !old.some(function (item) { return item.sheet === sheet; })) throw new Error('補欠の出力名が管理対象外シートと衝突しています: ' + label);
  });
  const metadata = { id: id, createdAt: options.startedAt, version: LOTTERY_RESERVE_VERSION,
    codeVersion: LOTTERY_VERSION, sourceRevision: sourceRevision_(),
    baseRunId: source.record.runId, seed: source.record.seed, seeds: result.seeds,
    listHash: result.listHash, auditHash: result.auditHash, rankedCount: result.ranked.length, unassignedCount: result.unassigned.length };
  const historyRecord = Object.assign({}, source.record, { reserve: metadata });
  const listRows = result.ranked.concat(result.unassigned).map(function (p) {
    return [p.committee || '希望なし（班未割当）', p.rank || '', p.className, p.number, p.name, p.category,
      p.preference, p.meetings, p.sourceRow, p.note];
  });
  const logRows = [['生成情報', JSON.stringify(metadata)], ['本抽選', JSON.stringify(source.record)]]
    .concat(result.audit.map(function (event) { return [event.type, JSON.stringify(event.detail)]; }));
  const plans = [
    { label: '補欠名簿', rows: [['班', '補欠順位', 'クラス', '出席番号', '名前', '区分', '希望順位', '準備・打ち合わせの回答', '元回答行', '確認事項']].concat(listRows) },
    { label: '補欠記録', rows: [['種別', '内容（JSON）']].concat(logRows) }
  ];
  plans.forEach(function (plan) {
    plan.rows = plan.rows.map(function (row) { return row.map(function (value) {
      const cell = safeCell_(value);
      if (typeof cell === 'string' && cell.length > 45000) throw new Error('補欠の記録が1セルの保存上限を超えます。');
      return cell;
    }); });
  });
  const created = []; const renamedOld = []; const renamedNew = [];
  function mark(item, phase) {
    item.sheet.getRange(1, 1).setNote(JSON.stringify({ owner: LOTTERY_OUTPUT_OWNER, kind: 'reserve', schema: 1,
      label: item.plan.label, phase: phase, record: source.record, reserve: metadata }));
  }
  try {
    plans.forEach(function (plan) {
      const temporary = LOTTERY_SETTINGS.outputPrefix + '補欠保存中_' + id + '_' + plan.label;
      const sheet = spreadsheet.insertSheet(temporary);
      const item = { sheet: sheet, plan: plan, temporary: temporary }; created.push(item);
      mark(item, 'staging'); sheet.hideSheet();
      if (sheet.getMaxRows() < plan.rows.length) sheet.insertRowsAfter(sheet.getMaxRows(), plan.rows.length - sheet.getMaxRows());
      for (let offset = 0; offset < plan.rows.length; offset += 500) {
        const chunk = plan.rows.slice(offset, offset + 500);
        sheet.getRange(offset + 1, 1, chunk.length, plan.rows[0].length).setNumberFormat('@').setValues(chunk);
      }
      formatOutput_(sheet, plan.label, plan.rows.length, plan.rows[0].length);
    });
    flush(); appendHistory_(spreadsheet, [{ event: '補欠保存準備', record: historyRecord }]); flush();
    old.forEach(function (item) {
      if (!LOTTERY_RESERVE_LABELS.some(function (label) { return item.sheet.getName() === LOTTERY_SETTINGS.outputPrefix + label; })) return;
      const previousName = item.sheet.getName();
      item.sheet.setName(LOTTERY_SETTINGS.outputPrefix + '補欠退避_' + item.sheet.getSheetId() + '_' + id);
      renamedOld.push({ sheet: item.sheet, name: previousName });
    });
    created.forEach(function (item) { item.sheet.setName(LOTTERY_SETTINGS.outputPrefix + item.plan.label); renamedNew.push(item); mark(item, 'published'); });
    flush();
    created.forEach(function (item) { item.sheet.showSheet(); });
    spreadsheet.setActiveSheet(created[0].sheet); spreadsheet.moveActiveSheet(4);
    appendHistory_(spreadsheet, [{ event: '補欠保存完了', record: historyRecord }]); flush();
  } catch (error) {
    renamedNew.slice().reverse().forEach(function (item) { try { item.sheet.setName(item.temporary); } catch (ignored) { /* 保持 */ } });
    renamedOld.slice().reverse().forEach(function (item) { try { item.sheet.setName(item.name); } catch (ignored) { /* 保持 */ } });
    created.forEach(function (item) { try { mark(item, 'failed'); item.sheet.showSheet(); } catch (ignored) { /* 保持 */ } });
    try { appendHistory_(spreadsheet, [{ event: '補欠保存失敗', record: historyRecord }]); flush(); } catch (ignored) { /* 本抽選は変更しない */ }
    throw new Error('補欠名簿の保存に失敗しました。本抽選と以前の補欠は削除していません。保存中・失敗の名簿を使用しないでください。');
  }
  const warnings = [];
  old.forEach(function (item) {
    try { spreadsheet.deleteSheet(item.sheet); } catch (error) { warnings.push('以前の補欠シートを保持: ' + item.sheet.getName()); }
  });
  return { baseRunId: source.record.runId, reserveId: id, ranked: result.ranked.length, unassigned: result.unassigned.length,
    sheetNames: created.map(function (item) { return item.sheet.getName(); }), warnings: warnings };
}

function executeReservesOnSpreadsheet_(spreadsheet, options) {
  const captured = readInputSheets_(spreadsheet);
  const source = readSavedLotteryForReserves_(spreadsheet, captured, options.digest);
  const result = buildReserveList_(source.prepared, source.assignments, source.fixedFirstByKey, source.record.seed, options.digest);
  return writeReservePlans_(spreadsheet, source, result, options);
}

function createCommitteeReserves() {
  const lock = LockService.getDocumentLock();
  if (!lock || !lock.tryLock(1000)) throw new Error('別の抽選・補欠処理が実行中か、スプレッドシートに紐づいていません。');
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); const now = new Date();
    const summary = executeReservesOnSpreadsheet_(spreadsheet, {
      runId: Utilities.formatDate(now, spreadsheet.getSpreadsheetTimeZone(), 'yyyyMMdd_HHmmss_SSS') + '_' + Utilities.getUuid().slice(0, 8),
      startedAt: now.toISOString(), flush: function () { SpreadsheetApp.flush(); }
    });
    spreadsheet.toast('本抽選の当選者を維持し、補欠名簿を保存しました。' + (summary.warnings.length ? '\n旧補欠の整理に注意事項があります。' : ''), '委員抽選', 10);
    return summary;
  } finally { lock.releaseLock(); }
}

function verifyHistoryEntries_(captured, entries, digest) {
  const responsesHash = snapshotHash_(captured.responses.values, digest);
  const operationsHash = snapshotHash_(captured.operations.values, digest);
  const records = new Map();
  const failed = new Set(entries.filter(function (entry) { return entry.event === '保存失敗'; })
    .map(function (entry) { return entry.record.runId; }));
  entries.filter(function (entry) {
    return !failed.has(entry.record.runId) && (entry.event === '保存完了' || entry.event === '旧版移行');
  })
    .forEach(function (entry) { records.set(entry.record.runId, entry.record); });
  const cache = new Map();
  return Array.from(records.values()).map(function (record) {
    if (record.responsesHash !== responsesHash || record.operationsHash !== operationsHash) {
      return { runId: record.runId, verification: '入力が当時と異なる' };
    }
    if (record.algorithmVersion !== LOTTERY_ALGORITHM_VERSION) {
      return { runId: record.runId, verification: '当時のプログラムでの確認が必要' };
    }
    if (!cache.has(record.seed)) {
      const result = runLotteryEngine_(captured.responses.data, captured.operations.data,
        { seed: record.seed, digest: digest, locations: captured });
      cache.set(record.seed, makeRunRecord_(captured, result, { runId: record.runId, startedAt: record.startedAt, digest: digest }));
    }
    const regenerated = cache.get(record.seed);
    return { runId: record.runId, verification: record.resultHash === regenerated.resultHash &&
      record.auditHash === regenerated.auditHash ? '一致' : '結果または全試行ログが不一致' };
  });
}

function verifyLotteryHistory() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const captured = readInputSheets_(spreadsheet);
  const checked = verifyHistoryEntries_(captured, historyRows_(spreadsheet.getSheetByName(LOTTERY_HISTORY_NAME)));
  const matched = checked.filter(function (item) { return item.verification === '一致'; }).length;
  SpreadsheetApp.getUi().alert('実行履歴の照合（書込なし）',
    '保存済み実行 ' + checked.length + '件／一致 ' + matched + '件\n' +
    checked.filter(function (item) { return item.verification !== '一致'; }).map(function (item) {
      return item.runId + ': ' + item.verification;
    }).join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
  return checked;
}

// Node.js組み込みテスト用。Apps Scriptではmoduleが存在せず、この分岐は実行されない。
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    LOTTERY_SEED, LOTTERY_VERSION, LOTTERY_ALGORITHM_VERSION, LOTTERY_PRNG, LOTTERY_SETTINGS,
    LOTTERY_OUTPUT_OWNER, LOTTERY_HISTORY_HEADER, LOTTERY_HISTORY_NAME,
    personKey_, prepareLotteryInput_, checkPossible_, maximumClassCoverage_, utf8Bytes_, hashText_,
    createRandom_, runLotteryEngine_, readInputSheets_, buildOutputPlans_, writeOutputPlans_,
    executeLotteryOnSpreadsheet_, safeCell_, serializedRow_, makeRunRecord_, hashResult_,
    historyRows_, collectOutputRuns_, verifyHistoryEntries_, readOutputNote_,
    LOTTERY_RESERVE_VERSION, LOTTERY_RESERVE_LABELS, buildReserveList_, readSavedLotteryForReserves_, executeReservesOnSpreadsheet_
  };
}
