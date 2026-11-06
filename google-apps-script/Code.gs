const SPREADSHEET_ID = '1K9WbtUH3S-IsgXPeLDLhnN0yDyruh1ksfz7RicAfNb8';
const ALLOWED_SOURCES = ['talk.varis.kr', 'talkwith.varis.kr'];
const SYNC_TOKEN = 'talk-varis-sync-v1';
const SHEETS = {
  sync: 'sync_log',
  active: 'active_match_history',
  archive: 'archive_match_history',
  presses: 'match_press_count',
  topics: 'interview_topics',
};

function doGet() {
  return json_({ ok: true, service: 'VARIS Workshop Matching Sync' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!ALLOWED_SOURCES.includes(payload.source) || payload.token !== SYNC_TOKEN) {
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
    const source = String(payload.source || '');

    append_(ss, SHEETS.active,
      ['recordedAt','sessionId','pairKey','participants','event','reason','source'],
      (activeHistory.length ? activeHistory : ['']).map(pairKey => [recordedAt, sessionId, String(pairKey), participantText, event, reason, source]));

    append_(ss, SHEETS.presses,
      ['recordedAt','sessionId','matchPresses','event','reason','source'],
      [[recordedAt, sessionId, Number(payload.matchPresses || 0), event, reason, source]]);

    append_(ss, SHEETS.topics,
      ['recordedAt','sessionId','topicId','topicTitle','event','source'],
      (topics.length ? topics : ['']).map(topic => [recordedAt, sessionId, String(topic), '', event, source]));

    archiveSnapshots.forEach(snapshot => {
      const pairs = Array.isArray(snapshot.history) ? snapshot.history : [''];
      append_(ss, SHEETS.archive,
        ['archivedAt','archiveId','sessionId','pairKey','participants','matchPresses','topics','reason','source'],
        pairs.map(pairKey => [snapshot.archivedAt || recordedAt, String(snapshot.archiveId || ''), sessionId, String(pairKey), (snapshot.participants || participants).join(' | '), Number(snapshot.matchPresses || 0), (snapshot.topics || topics).join(' | '), String(snapshot.reason || reason), source]));
    });

    append_(ss, SHEETS.sync,
      ['recordedAt','sessionId','event','storageKeys','status','source'],
      [[recordedAt, sessionId, event, 'history,archive,presses,topics', 'ok', source]]);

    return json_({ ok: true, recordedAt: recordedAt, event: event });
  } catch (error) {
    return json_({ ok: false, error: String(error) });
  } finally {
    lock.releaseLock();
  }
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
