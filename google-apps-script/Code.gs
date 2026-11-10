const SPREADSHEET_ID = '1K9WbtUH3S-IsgXPeLDLhnN0yDyruh1ksfz7RicAfNb8';
const ALLOWED_SOURCES = ['talk.varis.kr', 'talkwith.varis.kr'];
const PROJECTS = {
  'talk.varis.kr': 'talk-1on1',
  'talkwith.varis.kr': 'talk-group',
};
const SYNC_TOKEN = 'talk-varis-sync-v1';
const SHEETS = {
  sync: 'sync_log',
  active: 'active_match_history',
  archive: 'archive_match_history',
  presses: 'match_press_count',
  topics: 'interview_topics',
  pairs: 'pair_history',
};

function doGet(e) {
  const params = (e && e.parameter) || {};
  if (params.action !== 'history') {
    return json_({ ok: true, service: 'VARIS Workshop Matching Sync' });
  }

  const source = String(params.source || '');
  const projectKey = String(params.projectKey || PROJECTS[source] || '');
  if (!isAllowedProject_(source, projectKey) || params.token !== SYNC_TOKEN) {
    return json_({ ok: false, error: 'unauthorized' });
  }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const pairs = readPairHistory_(ss, source, projectKey);
    return json_({ ok: true, source, projectKey, pairs: pairs, count: pairs.length });
  } catch (error) {
    return json_({ ok: false, error: String(error) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const source = String(payload.source || '');
    const projectKey = String(payload.projectKey || PROJECTS[source] || '');
    if (!isAllowedProject_(source, projectKey) || payload.token !== SYNC_TOKEN) {
      return json_({ ok: false, error: 'unauthorized' });
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const recordedAt = payload.recordedAt || new Date().toISOString();
    const sessionId = String(payload.sessionId || 'unknown');
    const event = String(payload.event || 'sync');
    const reason = String(payload.reason || '');
    const activeHistory = Array.isArray(payload.activeHistory) ? payload.activeHistory : [];
    const archiveSnapshots = Array.isArray(payload.archiveSnapshots) ? payload.archiveSnapshots : [];
    const topics = Array.isArray(payload.usedInterviewTopics) ? payload.usedInterviewTopics : [];
    const participants = Array.isArray(payload.participants) ? payload.participants : [];
    const participantText = participants.join(' | ');

    const pairRows = [];
    activeHistory.forEach(pairKey => {
      if (pairKey) pairRows.push([recordedAt, projectKey, source, sessionId, String(pairKey), event]);
    });
    archiveSnapshots.forEach(snapshot => {
      const history = Array.isArray(snapshot.history) ? snapshot.history : [];
      history.forEach(pairKey => {
        if (pairKey) pairRows.push([snapshot.archivedAt || recordedAt, projectKey, source, String(snapshot.sessionId || sessionId), String(pairKey), 'archive']);
      });
    });
    append_(ss, SHEETS.pairs,
      ['recordedAt','projectKey','source','sessionId','pairKey','event'], pairRows);

    append_(ss, SHEETS.active,
      ['recordedAt','sessionId','pairKey','participants','event','reason','source','projectKey'],
      (activeHistory.length ? activeHistory : ['']).map(pairKey => [recordedAt, sessionId, String(pairKey), participantText, event, reason, source, projectKey]));

    append_(ss, SHEETS.presses,
      ['recordedAt','sessionId','matchPresses','event','reason','source','projectKey'],
      [[recordedAt, sessionId, Number(payload.matchPresses || 0), event, reason, source, projectKey]]);

    append_(ss, SHEETS.topics,
      ['recordedAt','sessionId','topicId','topicTitle','event','source','projectKey'],
      (topics.length ? topics : ['']).map(topic => [recordedAt, sessionId, String(topic), '', event, source, projectKey]));

    archiveSnapshots.forEach(snapshot => {
      const pairs = Array.isArray(snapshot.history) ? snapshot.history : [''];
      append_(ss, SHEETS.archive,
        ['archivedAt','archiveId','sessionId','pairKey','participants','matchPresses','topics','reason','source','projectKey'],
        pairs.map(pairKey => [snapshot.archivedAt || recordedAt, String(snapshot.archiveId || ''), String(snapshot.sessionId || sessionId), String(pairKey), (snapshot.participants || participants).join(' | '), Number(snapshot.matchPresses || 0), (snapshot.topics || topics).join(' | '), String(snapshot.reason || reason), source, projectKey]));
    });

    append_(ss, SHEETS.sync,
      ['recordedAt','sessionId','event','storageKeys','status','source','projectKey'],
      [[recordedAt, sessionId, event, 'history,archive,presses,topics,pairs', 'ok', source, projectKey]]);

    return json_({ ok: true, recordedAt: recordedAt, event: event, projectKey: projectKey });
  } catch (error) {
    return json_({ ok: false, error: String(error) });
  } finally {
    lock.releaseLock();
  }
}

function isAllowedProject_(source, projectKey) {
  return ALLOWED_SOURCES.indexOf(source) !== -1 && PROJECTS[source] === projectKey;
}

function readPairHistory_(ss, source, projectKey) {
  const seen = {};
  const pairsSheet = ss.getSheetByName(SHEETS.pairs);
  if (pairsSheet && pairsSheet.getLastRow() > 1) {
    const values = pairsSheet.getDataRange().getValues();
    const headers = values[0].map(String);
    const sourceIndex = headers.indexOf('source');
    const projectIndex = headers.indexOf('projectKey');
    const pairIndex = headers.indexOf('pairKey');
    values.slice(1).forEach(row => {
      if (pairIndex < 0) return;
      const rowSource = sourceIndex >= 0 ? String(row[sourceIndex] || '') : '';
      const rowProject = projectIndex >= 0 ? String(row[projectIndex] || '') : '';
      const pair = String(row[pairIndex] || '');
      if (rowSource === source && rowProject === projectKey && pair) seen[pair] = true;
    });
  }

  // Backward compatibility with records written before pair_history/projectKey existed.
  [SHEETS.active, SHEETS.archive].forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) return;
    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(String);
    const sourceIndex = headers.indexOf('source');
    const projectIndex = headers.indexOf('projectKey');
    const pairIndex = headers.indexOf('pairKey');
    values.slice(1).forEach(row => {
      const rowSource = sourceIndex >= 0 ? String(row[sourceIndex] || '') : '';
      const rowProject = projectIndex >= 0 ? String(row[projectIndex] || '') : PROJECTS[rowSource] || '';
      const pair = pairIndex >= 0 ? String(row[pairIndex] || '') : '';
      if (rowSource === source && rowProject === projectKey && pair) seen[pair] = true;
    });
  });
  return Object.keys(seen);
}

function append_(ss, sheetName, headers, rows) {
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (!rows || !rows.length) return;
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}
