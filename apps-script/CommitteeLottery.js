const LOTTERY_SEED = '72f7c0a2629261e23b8f7627a22233d17c72740d6305acdc9716789bf70ec83c';
const LOTTERY_VERSION = '1.0.0';
const LOTTERY_PRNG = 'SHA-256-chain / uint32-be / rejection-sampling / Fisher-Yates';
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
    .addItem('抽選を実行して新規シートに保存', 'runCommitteeLottery')
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

/** この関数を明示的に実行したときだけ、新しい出力シートを追加する。 */
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
    spreadsheet.toast(summary.status === 'ACCEPTED' ? '抽選結果を保存しました。' :
      '選出未成立です。詳細ログと統計を確認してください。', '委員抽選', 10);
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
    ['実行ID', context.runId], ['状態', '保存中・当選未確定'], ['方式の版', result.version], ['シード', result.seed],
    ['乱数方式', result.prng], ['元回答SHA256', snapshotHash_(captured.responses.values, context.digest)],
    ['運営確認SHA256', snapshotHash_(captured.operations.values, context.digest)], ['採用試行', result.acceptedAttempt || 'なし'],
    ['実施試行数', result.attempts.length], ['消費した乱数の個数', result.randomPosition],
    ['最終乱数状態', result.finalRandomState], ['実行開始日時', context.startedAt], ['タイムゾーン', captured.timezone],
    ['元回答シート・範囲', captured.responses.name + '!' + captured.responses.range],
    ['運営確認シート・範囲', captured.operations.name + '!' + captured.operations.range]
  ];
  function plan(label, header, rows) { return { label: label, rows: metadata.map(function (r) { return r.slice(); }).concat([[], header], rows) }; }
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
    plan('詳細ログ', ['連番', '種別', '内容（JSON）'], result.audit.map(function (event) {
      return [event.sequence, event.type, JSON.stringify(event.detail)];
    })),
    plan('元データ統計', ['項目', '区分', '人数・件数', '注記'], inputStats),
    plan('結果統計', ['試行', '採否', '項目', '区分', '人数', '条件・備考'], resultStats),
    plan('役職希望者', ['班', 'クラス', '出席番号', '名前', '元回答行'], leaderRows),
    plan('入力控え', ['入力種別', '元シートID', '元シート名', '元行番号', '値のJSON（型付き）'], snapshots)
  ];
}

function safeCell_(value) {
  if (typeof value === 'string' && /^[=+\-@\t\r]/.test(value)) return "'" + value;
  return value === null || value === undefined ? '' : value;
}

function writeOutputPlans_(spreadsheet, plans, runId, status, flush) {
  flush = flush || function () {};
  let uniqueId = runId; let suffix = 1;
  while (plans.some(function (plan) { return spreadsheet.getSheetByName(LOTTERY_SETTINGS.outputPrefix + uniqueId + '_' + plan.label); })) {
    uniqueId = runId + '_' + suffix; suffix += 1;
  }
  const preparedPlans = plans.map(function (plan) {
    const width = Math.max.apply(null, plan.rows.map(function (row) { return row.length; }));
    const rows = plan.rows.map(function (row, index) {
      const result = row.map(safeCell_);
      if (index === 0) result[1] = uniqueId;
      while (result.length < width) result.push('');
      result.forEach(function (cell) {
        if (typeof cell === 'string' && cell.length > 45000) throw new Error('入力またはログが1セルの保存上限を超えます。');
      });
      return result;
    });
    return { label: plan.label, name: LOTTERY_SETTINGS.outputPrefix + uniqueId + '_' + plan.label, width: width, rows: rows };
  });
  const created = [];
  try {
    preparedPlans.forEach(function (plan) {
      const sheet = spreadsheet.insertSheet(plan.name);
      created.push(sheet);
      if (sheet.getMaxRows() < plan.rows.length) sheet.insertRowsAfter(sheet.getMaxRows(), plan.rows.length - sheet.getMaxRows());
      if (sheet.getMaxColumns() < plan.width) sheet.insertColumnsAfter(sheet.getMaxColumns(), plan.width - sheet.getMaxColumns());
      for (let offset = 0; offset < plan.rows.length; offset += 500) {
        const chunk = plan.rows.slice(offset, offset + 500);
        const range = sheet.getRange(offset + 1, 1, chunk.length, plan.width);
        range.setNumberFormat('@'); range.setValues(chunk);
      }
      sheet.setFrozenRows(2);
    });
    flush();
    // 結果シートの状態を最後に確定する。他の出力が欠けた状態を成立として扱わない。
    for (let i = 1; i < created.length; i += 1) created[i].getRange(2, 2).setValue(displayStatus_(status));
    flush();
    created[0].getRange(2, 2).setValue(displayStatus_(status));
    flush();
  } catch (error) {
    created.forEach(function (sheet) {
      try { sheet.getRange(2, 2).setValue('出力失敗・当選未確定'); } catch (ignored) { /* 元の例外を保持する */ }
    });
    try { flush(); } catch (ignored) { /* 元の例外を保持する */ }
    throw new Error('出力の保存に失敗しました。作成済みシートは削除せず残しました。当選結果として使用しないでください。');
  }
  return { runId: uniqueId, sheetNames: created.map(function (sheet) { return sheet.getName(); }) };
}

function executeLotteryOnSpreadsheet_(spreadsheet, options) {
  const captured = readInputSheets_(spreadsheet);
  const result = runLotteryEngine_(captured.responses.data, captured.operations.data,
    { seed: options.seed, digest: options.digest, locations: captured });
  const plans = buildOutputPlans_(captured, result, options);
  const output = writeOutputPlans_(spreadsheet, plans, options.runId, result.status, options.flush);
  return {
    status: result.status, runId: output.runId, sheetNames: output.sheetNames,
    acceptedAttempt: result.acceptedAttempt, attempts: result.attempts.length,
    applicants: result.prepared.candidates.length, eligible: result.prepared.eligible.length,
    selected: result.assignments.length
  };
}

// Node.js組み込みテスト用。Apps Scriptではmoduleが存在せず、この分岐は実行されない。
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    LOTTERY_SEED, LOTTERY_VERSION, LOTTERY_PRNG, LOTTERY_SETTINGS,
    personKey_, prepareLotteryInput_, checkPossible_, maximumClassCoverage_, utf8Bytes_, hashText_,
    createRandom_, runLotteryEngine_, readInputSheets_, buildOutputPlans_, writeOutputPlans_,
    executeLotteryOnSpreadsheet_, safeCell_, serializedRow_
  };
}
