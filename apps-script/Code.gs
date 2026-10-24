/**
 * Hive Operations Hub - Google Apps Script web app.
 * The Google Sheet this script is attached to is the database.
 * Tabs: Team, Tasks, EOD, Findings, Notes, Runs (created by setupHub).
 * Passwords are checked here on Google's server; the page never sees them.
 */

var HUB = {
  name: 'Hive Operations Hub',
  sessionDays: 14,
  hashRounds: 1000,
  levels: ['Expert', 'Admin', 'Organizer'],
  statuses: ['Open', 'In progress', 'Done', 'Blocked'],
  findingStatuses: ['Open', 'Done', 'Ignore'],
  roles: ['Build', 'Client Support', 'Guest Experience']
};

var TABLES = {
  Team: [['Name', 'name'], ['Level', 'level'], ['Role', 'role'], ['Active', 'active'], ['Salt', 'salt'], ['Password hash', 'hash']],
  Tasks: [['ID', 'id'], ['Date', 'date'], ['Person', 'person'], ['Task', 'task'], ['Client', 'client'], ['Role', 'ownerRole'], ['Due', 'deadline'], ['Status', 'status'], ['What they did', 'doneNote'], ['Updated by', 'updatedBy'], ['Updated at', 'updatedAt'], ['Completed at', 'completedAt'], ['Source', 'source'], ['Notes', 'notes'], ['Created at', 'createdAt']],
  EOD: [['ID', 'id'], ['Date', 'date'], ['Person', 'person'], ['Done', 'done'], ['Blockers', 'blockers'], ['Next', 'next'], ['Created at', 'createdAt']],
  Findings: [['ID', 'id'], ['Date', 'date'], ['Client', 'client'], ['Sub-account ID', 'subAccountId'], ['Plan', 'plan'], ['Stage', 'stage'], ['Issue', 'issue'], ['Severity', 'severity'], ['Evidence', 'evidence'], ['Owner role', 'ownerRole'], ['Next action', 'nextAction'], ['Status', 'status'], ['Created at', 'createdAt']],
  Notes: [['ID', 'id'], ['Date', 'date'], ['Applies to', 'appliesTo'], ['Note', 'text'], ['Created at', 'createdAt']],
  Runs: [['ID', 'id'], ['Agent', 'agent'], ['Timestamp', 'ts'], ['Rows read', 'rowsRead'], ['Items added', 'itemsAdded'], ['Notes', 'notes']]
};

/* ---------- Web app entry ---------- */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(HUB.name)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Hive Hub')
    .addItem('Set up or repair tabs', 'setupHub')
    .addItem('Set a password', 'setPasswordFromMenu')
    .addToUi();
}

/* ---------- Sheet helpers ---------- */

function tz_() { return Session.getScriptTimeZone() || 'America/Chicago'; }
function today_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }
function nowIso_() { return new Date().toISOString(); }

function sheet_(name) {
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('The ' + name + ' tab is missing. Open the sheet and use Hive Hub > Set up or repair tabs.');
  return sh;
}

function cell_(v) {
  if (v instanceof Date) {
    var timed = v.getHours() || v.getMinutes() || v.getSeconds();
    return Utilities.formatDate(v, tz_(), timed ? "yyyy-MM-dd'T'HH:mm:ss" : 'yyyy-MM-dd');
  }
  return v == null ? '' : String(v).trim();
}

/* Text typed into the portal is stored as plain text, never as a formula. */
function safe_(v) {
  if (typeof v !== 'string') return v;
  v = v.replace(/\r\n/g, '\n');
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function readTable_(name) {
  var sh = sheet_(name), spec = TABLES[name];
  var values = sh.getDataRange().getValues();
  var head = (values[0] || []).map(function (h) { return String(h).trim().toLowerCase(); });
  var cols = {};
  spec.forEach(function (p) { cols[p[1]] = head.indexOf(p[0].toLowerCase()); });
  var rows = [];
  for (var r = 1; r < values.length; r++) {
    var o = { _row: r + 1 }, any = false;
    spec.forEach(function (p) {
      var i = cols[p[1]];
      var v = i >= 0 ? cell_(values[r][i]) : '';
      o[p[1]] = v;
      if (v) any = true;
    });
    if (any) rows.push(o);
  }
  return { name: name, sh: sh, cols: cols, rows: rows };
}

function setFields_(t, rowNum, obj) {
  Object.keys(obj).forEach(function (k) {
    var i = t.cols[k];
    if (i == null || i < 0) throw new Error('The ' + t.name + ' tab is missing a column. Use Hive Hub > Set up or repair tabs.');
    t.sh.getRange(rowNum, i + 1).setValue(safe_(obj[k]));
  });
}

function appendRow_(t, obj) {
  var width = Math.max(t.sh.getLastColumn(), TABLES[t.name].length);
  var row = [];
  for (var c = 0; c < width; c++) row.push('');
  TABLES[t.name].forEach(function (p) {
    var i = t.cols[p[1]];
    if (i >= 0 && obj[p[1]] != null) row[i] = safe_(obj[p[1]]);
  });
  t.sh.appendRow(row);
}

function strip_(o) { var c = {}; Object.keys(o).forEach(function (k) { if (k !== '_row') c[k] = o[k]; }); return c; }
function newId_(prefix) { return prefix + Utilities.getUuid().replace(/-/g, '').slice(0, 12); }

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function findRow_(t, id) {
  var r = t.rows.filter(function (x) { return x.id === id; })[0];
  if (!r) throw new Error('That item was not found. Refresh and try again.');
  return r;
}

/* ---------- People, passwords, sessions ---------- */

function normLevel_(v) {
  v = String(v || '').trim().toLowerCase();
  if (v === 'organizer' || v === 'organiser') return 'Organizer';
  if (v === 'admin') return 'Admin';
  return 'Expert';
}
function isInactive_(m) { return /^(no|false|inactive|0|off)$/i.test(String(m.active || '').trim()); }

function members_() { return readTable_('Team').rows.filter(function (m) { return m.name && !isInactive_(m); }); }
function findMember_(name) {
  name = String(name || '').trim().toLowerCase();
  return readTable_('Team').rows.filter(function (m) { return m.name.toLowerCase() === name; })[0] || null;
}

function hash_(salt, pw) {
  var d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pw, Utilities.Charset.UTF_8);
  for (var i = 1; i < HUB.hashRounds; i++) d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, d);
  return d.map(function (b) { var x = (b + 256) % 256; return (x < 16 ? '0' : '') + x.toString(16); }).join('');
}
function newSalt_() { return Utilities.getUuid().replace(/-/g, ''); }

function props_() { return PropertiesService.getScriptProperties(); }

function cleanSessions_(onlyName) {
  var all = props_().getProperties(), now = Date.now();
  Object.keys(all).forEach(function (k) {
    if (k.indexOf('s_') !== 0) return;
    try {
      var s = JSON.parse(all[k]);
      if (s.exp < now || (onlyName && s.name.toLowerCase() === onlyName.toLowerCase())) props_().deleteProperty(k);
    } catch (e) { props_().deleteProperty(k); }
  });
}

function who_(token) {
  if (!token || !/^[a-f0-9]{64}$/.test(String(token))) throw new Error('SIGNED_OUT');
  var raw = props_().getProperty('s_' + token);
  if (!raw) throw new Error('SIGNED_OUT');
  var s = JSON.parse(raw);
  if (s.exp < Date.now()) { props_().deleteProperty('s_' + token); throw new Error('SIGNED_OUT'); }
  var m = findMember_(s.name);
  if (!m || isInactive_(m)) throw new Error('SIGNED_OUT');
  return { name: m.name, level: normLevel_(m.level) };
}

function need_(me, levels) {
  if (levels.indexOf(me.level) < 0) throw new Error('You do not have access to do that.');
}

/* ---------- Public API (called from the page) ---------- */

function getNames() {
  return members_().filter(function (m) { return m.hash; }).map(function (m) { return m.name; }).sort();
}

function login(name, password) {
  name = String(name || '').trim();
  password = String(password || '');
  var cache = CacheService.getScriptCache(), key = 'fail_' + name.toLowerCase();
  var fails = Number(cache.get(key) || 0);
  if (fails >= 5) throw new Error('Too many attempts. Wait 10 minutes and try again.');
  var m = findMember_(name);
  if (!m || isInactive_(m) || !m.salt || !m.hash || hash_(m.salt, password) !== m.hash) {
    cache.put(key, String(fails + 1), 600);
    throw new Error('That name and password do not match.');
  }
  cache.remove(key);
  cleanSessions_();
  var token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  props_().setProperty('s_' + token, JSON.stringify({ name: m.name, exp: Date.now() + HUB.sessionDays * 864e5 }));
  return { token: token, name: m.name, level: normLevel_(m.level) };
}

function logout(token) {
  if (token && /^[a-f0-9]{64}$/.test(String(token))) props_().deleteProperty('s_' + token);
  return true;
}

function fillTaskIds_() {
  var t = readTable_('Tasks');
  var missing = t.rows.filter(function (r) { return r.task && !r.id; });
  if (!missing.length) return;
  withLock_(function () {
    var t2 = readTable_('Tasks');
    t2.rows.forEach(function (r) {
      if (!r.task || r.id) return;
      var f = { id: newId_('t_') };
      if (!r.status) f.status = 'Open';
      if (!r.date) f.date = today_();
      if (!r.createdAt) f.createdAt = nowIso_();
      if (!r.source) f.source = 'Sheet';
      setFields_(t2, r._row, f);
    });
  });
}

function getData(token, viewAs) {
  var me = who_(token);
  var viewer = { name: me.name, level: me.level }, preview = false;
  if (viewAs && me.level === 'Organizer') {
    var v = findMember_(viewAs);
    if (v && !isInactive_(v) && normLevel_(v.level) !== 'Organizer') { viewer = { name: v.name, level: normLevel_(v.level) }; preview = true; }
  }
  fillTaskIds_();
  var team = members_().map(function (m) { return { name: m.name, level: normLevel_(m.level), role: m.role, hasPassword: !!m.hash }; });
  var tasks = readTable_('Tasks').rows.filter(function (r) { return r.task; }).map(strip_);
  var eod = readTable_('EOD').rows.map(strip_);
  var out = {
    me: me.name, level: me.level, viewer: viewer.name, viewerLevel: viewer.level, preview: preview,
    today: today_(), statuses: HUB.statuses, findingStatuses: HUB.findingStatuses, roles: HUB.roles, levels: HUB.levels
  };
  if (me.level === 'Organizer') out.viewAsNames = team.filter(function (m) { return m.level !== 'Organizer'; }).map(function (m) { return m.name; });
  if (viewer.level === 'Expert') {
    out.tasks = tasks.filter(function (t) { return t.person === viewer.name; });
    out.eod = eod.filter(function (e) { return e.person === viewer.name; });
    out.people = [];
  } else {
    out.tasks = tasks;
    out.eod = eod;
    out.people = team.filter(function (m) { return m.level !== 'Admin'; }).map(function (m) { return m.name; });
  }
  if (viewer.level !== 'Organizer') out.tasks.forEach(function (t) { t.source = ''; });
  if (viewer.level === 'Organizer') {
    out.findings = readTable_('Findings').rows.map(strip_);
    out.notes = readTable_('Notes').rows.map(strip_);
    out.runs = readTable_('Runs').rows.map(strip_);
    out.team = team;
  }
  return out;
}

function updateTask(token, id, status, doneNote) {
  var me = who_(token);
  if (HUB.statuses.indexOf(status) < 0) throw new Error('Unknown status.');
  return withLock_(function () {
    var t = readTable_('Tasks'), r = findRow_(t, id);
    if (me.level === 'Expert' && r.person !== me.name) throw new Error('You can only update your own tasks.');
    var now = nowIso_();
    var f = { status: status, updatedBy: me.name, updatedAt: now };
    if (doneNote != null) f.doneNote = String(doneNote).slice(0, 4000);
    if (status === 'Done' && r.status !== 'Done') f.completedAt = now;
    setFields_(t, r._row, f);
    return true;
  });
}

function addTask(token, o) {
  var me = who_(token);
  need_(me, ['Admin', 'Organizer']);
  o = o || {};
  var task = String(o.task || '').trim(), person = String(o.person || '').trim();
  if (!task || !person) throw new Error('Add the task and who it is for.');
  if (!findMember_(person)) throw new Error('That person is not on the team list.');
  return withLock_(function () {
    appendRow_(readTable_('Tasks'), {
      id: newId_('t_'), date: today_(), person: person, task: task.slice(0, 2000),
      client: String(o.client || '').slice(0, 200), ownerRole: String(o.ownerRole || ''),
      deadline: String(o.deadline || ''), status: 'Open', source: me.name,
      notes: String(o.notes || '').slice(0, 1000), createdAt: nowIso_()
    });
    return true;
  });
}

function reassignTask(token, id, person) {
  var me = who_(token);
  need_(me, ['Admin', 'Organizer']);
  if (person && !findMember_(person)) throw new Error('That person is not on the team list.');
  return withLock_(function () {
    var t = readTable_('Tasks'), r = findRow_(t, id);
    setFields_(t, r._row, { person: person || '', updatedBy: me.name, updatedAt: nowIso_() });
    return true;
  });
}

/* ---------- Voice tasks (Gemini) ----------
 * Script properties (Project Settings, Script properties):
 *   GEMINI_API_KEY           required, from aistudio.google.com
 *   GEMINI_MODEL             optional, defaults to VOICE.model
 *   The HIVE Master Ops Gem brain (Assistant tab + Client Master Sheet) is added to the prompt; see hiveBrain_.
 */
var VOICE = { model: 'gemini-3.8-flash', maxAudioBytes: 15 * 1024 * 1024, maxTasks: 40 };

function assignable_() { return members_().filter(function (m) { return normLevel_(m.level) !== 'Admin'; }); }

function voiceTasks(token, audio, mimeType, typed) {
  var me = who_(token);
  need_(me, ['Admin', 'Organizer']);
  return voiceTasks_(audio, mimeType, typed);
}

function voiceTasks_(audio, mimeType, typed) {
  audio = String(audio || '');
  typed = String(typed || '').trim().slice(0, 20000);
  if (!audio && !typed) throw new Error('Record a voice note or type what needs doing.');
  if (audio.length * 3 / 4 > VOICE.maxAudioBytes) throw new Error('That voice note is too long. Keep it under about 15 minutes.');

  var people = assignable_();
  var system = [
    'Background knowledge about H.I.V.E. (from the HIVE Master Ops Gem):\n' + hiveBrain_(),
    'YOUR JOB NOW: turn the operations manager\'s spoken brief into clear, assignable tasks for the team. Ignore the output formats above; reply only in the JSON shape asked for.',
    'Team members you may assign (use these exact names; names in the recording may be misheard, so match each to the closest name here): ' +
      people.map(function (m) { return m.name + (m.role ? ' (' + m.role + ')' : ''); }).join(', ') + '.',
    'Roles: ' + HUB.roles.join(', ') + '. Set ownerRole to one of these only when it is clear, otherwise leave it empty.',
    'Today is ' + Utilities.formatDate(new Date(), tz_(), 'EEEE yyyy-MM-dd') + '. Turn spoken dates such as "tomorrow" or "by Friday" into YYYY-MM-DD in deadline; leave deadline empty when no date was said.',
    'Split the brief into separate tasks. Each task starts with a verb and is specific and refined, but keeps every detail that was said (links, numbers, client names). ' +
      'One person per task: when a task is for several people, add one copy for each of them. When nobody was named for a task, set person to an empty string. ' +
      'Group the tasks by person in the order people were first mentioned, and order each person\'s tasks by priority. ' +
      'Put the client or business name in client when one is mentioned. Put useful context or priority in notes. Put the full transcript of the recording in transcript.'
  ].filter(String).join('\n\n');

  var schema = {
    type: 'OBJECT', required: ['tasks'],
    properties: {
      transcript: { type: 'STRING' },
      tasks: { type: 'ARRAY', items: { type: 'OBJECT', required: ['person', 'task'], properties: {
        person: { type: 'STRING' }, task: { type: 'STRING' }, client: { type: 'STRING' },
        ownerRole: { type: 'STRING' }, deadline: { type: 'STRING' }, notes: { type: 'STRING' } } } }
    }
  };
  var parts = [];
  if (audio) parts.push({ inline_data: { mime_type: String(mimeType || 'audio/webm').split(';')[0], data: audio } });
  parts.push({ text: (typed ? 'Typed brief:\n' + typed + '\n\n' : '') + (audio ? 'Listen to the voice note, then make the tasks.' : 'Make the tasks.') });

  var text = gemini_({
    system_instruction: { parts: [{ text: system }] },
    contents: [{ role: 'user', parts: parts }],
    generationConfig: { temperature: 0.2, responseMimeType: 'application/json', responseSchema: schema }
  });
  var out;
  try { out = JSON.parse(text); } catch (e) { throw new Error('Gemini did not return tasks. Try again.'); }

  var byName = {};
  people.forEach(function (m) { byName[m.name.toLowerCase()] = m.name; });
  var tasks = (out.tasks || []).slice(0, VOICE.maxTasks).map(function (t) {
    return {
      person: byName[String(t.person || '').trim().toLowerCase()] || '',
      task: String(t.task || '').trim().slice(0, 2000), client: String(t.client || '').trim().slice(0, 200),
      ownerRole: HUB.roles.indexOf(t.ownerRole) >= 0 ? t.ownerRole : '',
      deadline: /^\d{4}-\d{2}-\d{2}$/.test(String(t.deadline || '')) ? t.deadline : '',
      notes: String(t.notes || '').trim().slice(0, 1000)
    };
  }).filter(function (t) { return t.task; });
  return { transcript: String(out.transcript || typed).slice(0, 20000), tasks: tasks };
}

/* ---------- Hive AI chat (Gemini) ----------
 * Two modes: 'hive' answers like the HIVE Master Ops Gem (instructions in the Assistant tab plus a live read
 * of the Client Master Sheet, its knowledge file); 'gemini' is plain Gemini with Google Search.
 * Assistant tab rows: Name | Text. "Instructions" = the Gem's instructions; any other row is added as background.
 * Script property CLIENT_MASTER_SHEET_ID overrides ASSIST.masterSheetId.
 */
var ASSIST = { masterSheetId: '1r3RR53Q5NCdtoaHqURSrcz6kyfyBx53eMLQKlScC9Ao', maxTurns: 24, maxSheetChars: 60000 };

function gemini_(body) {
  var key = props_().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('Gemini is not set up yet: add GEMINI_API_KEY to the Apps Script project settings.');
  var model = props_().getProperty('GEMINI_MODEL') || VOICE.model;
  var res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true, headers: { 'x-goog-api-key': key }, payload: JSON.stringify(body)
  });
  var out = {};
  try { out = JSON.parse(res.getContentText()); } catch (e) {}
  if (res.getResponseCode() !== 200) throw new Error('Gemini could not process that: ' + ((out.error && out.error.message) || 'error ' + res.getResponseCode()));
  return ((((out.candidates || [])[0] || {}).content || {}).parts || []).map(function (p) { return p.text || ''; }).join('');
}

function hiveBrain_() {
  var parts = [], sh = SpreadsheetApp.getActive().getSheetByName('Assistant');
  if (sh) {
    var rows = sh.getDataRange().getValues().slice(1).filter(function (r) { return String(r[1] || '').trim(); });
    rows.sort(function (a, b) { return (String(a[0]).trim() === 'Instructions' ? 0 : 1) - (String(b[0]).trim() === 'Instructions' ? 0 : 1); });
    rows.forEach(function (r) { parts.push((String(r[0]).trim() === 'Instructions' ? '' : String(r[0]).trim().toUpperCase() + ':\n') + String(r[1]).trim()); });
  }
  if (!parts.length) parts.push('You are the H.I.V.E. Master Operations AI for Powered by H.I.V.E., a GoHighLevel CRM agency for U.S. shared-housing owners.');
  try {
    var ss = SpreadsheetApp.openById(props_().getProperty('CLIENT_MASTER_SHEET_ID') || ASSIST.masterSheetId);
    var live = ss.getSheets().map(function (s) {
      var csv = s.getDataRange().getDisplayValues().filter(function (r) { return r.join('').trim(); })
        .map(function (r) { return r.map(function (c) { return String(c).replace(/\s+/g, ' ').trim(); }).join(' | '); }).join('\n');
      return 'TAB "' + s.getName() + '":\n' + csv;
    }).join('\n\n').slice(0, ASSIST.maxSheetChars);
    parts.push('LIVE CLIENT MASTER SHEET, read just now (' + Utilities.formatDate(new Date(), tz_(), 'MMM d, yyyy h:mm a') +
      '). This is newer than any snapshot above and wins over it:\n' + live);
  } catch (e) {
    parts.push('(The live Client Master Sheet could not be read just now: ' + e.message + '. Use the snapshot above and say it may be stale.)');
  }
  return parts.join('\n\n');
}

function askAssistant(token, mode, history, audio, mimeType) {
  var me = who_(token);
  need_(me, ['Admin', 'Organizer']);
  return askAssistant_(me, mode, history, audio, mimeType);
}

function askAssistant_(me, mode, history, audio, mimeType) {
  history = (Array.isArray(history) ? history : []).slice(-ASSIST.maxTurns);
  audio = String(audio || '');
  if (audio.length * 3 / 4 > VOICE.maxAudioBytes) throw new Error('That voice message is too long.');
  var contents = history.map(function (m) {
    return { role: m && m.role === 'model' ? 'model' : 'user', parts: [{ text: String((m && m.text) || '').slice(0, 12000) || '(voice message)' }] };
  });
  if (!contents.length || contents[contents.length - 1].role !== 'user') throw new Error('Ask something first.');
  if (audio) {
    var last = contents[contents.length - 1];
    last.parts = [{ inline_data: { mime_type: String(mimeType || 'audio/webm').split(';')[0], data: audio } },
      { text: 'This is a voice message. Answer it directly.' }];
  }
  var when = 'Today is ' + Utilities.formatDate(new Date(), tz_(), 'EEEE MMM d, yyyy') + '. You are talking with ' + me.name + ' (' + me.level + ') in the Hive Operations Hub.';
  var body = { contents: contents, generationConfig: { temperature: mode === 'hive' ? 0.3 : 0.7 } };
  if (mode === 'hive') body.system_instruction = { parts: [{ text: hiveBrain_() + '\n\n' + when }] };
  else { body.system_instruction = { parts: [{ text: 'You are Gemini, a helpful assistant. ' + when + ' They work at Powered by H.I.V.E., a GoHighLevel CRM agency for shared-housing owners.' }] }; body.tools = [{ google_search: {} }]; }
  var reply = gemini_(body).trim();
  if (!reply) throw new Error('Gemini returned an empty answer. Try again.');
  return { reply: reply };
}

function addTasks(token, list) {
  var me = who_(token);
  need_(me, ['Admin', 'Organizer']);
  list = Array.isArray(list) ? list.slice(0, VOICE.maxTasks) : [];
  if (!list.length) throw new Error('There are no tasks to assign.');
  var byName = {};
  members_().forEach(function (m) { byName[m.name.toLowerCase()] = m.name; });
  var rows = list.map(function (o) {
    o = o || {};
    var task = String(o.task || '').trim(), person = byName[String(o.person || '').trim().toLowerCase()];
    if (!task) throw new Error('One of the tasks is empty. Fill it in or remove it.');
    if (!person) throw new Error('Choose who this task is for: ' + task.slice(0, 60));
    return { person: person, task: task.slice(0, 2000), client: String(o.client || '').slice(0, 200), ownerRole: String(o.ownerRole || ''),
      deadline: String(o.deadline || ''), notes: String(o.notes || '').slice(0, 1000) };
  });
  return withLock_(function () {
    var t = readTable_('Tasks'), now = nowIso_();
    rows.forEach(function (r) {
      appendRow_(t, { id: newId_('t_'), date: today_(), person: r.person, task: r.task, client: r.client, ownerRole: r.ownerRole,
        deadline: r.deadline, status: 'Open', source: me.name + ' (voice)', notes: r.notes, createdAt: now });
    });
    return rows.length;
  });
}

function addEod(token, o) {
  var me = who_(token);
  o = o || {};
  if (me.level === 'Admin') throw new Error('Admins do not file end-of-day reports.');
  var person = me.level === 'Organizer' && o.person ? String(o.person) : me.name;
  if (!findMember_(person)) throw new Error('That person is not on the team list.');
  var done = String(o.done || '').trim();
  if (!done) throw new Error('Say what got done today.');
  return withLock_(function () {
    appendRow_(readTable_('EOD'), {
      id: newId_('e_'), date: today_(), person: person, done: done.slice(0, 6000),
      blockers: String(o.blockers || '').slice(0, 2000), next: String(o.next || '').slice(0, 2000), createdAt: nowIso_()
    });
    return true;
  });
}

function setFindingStatus(token, id, status) {
  var me = who_(token);
  need_(me, ['Organizer']);
  if (HUB.findingStatuses.indexOf(status) < 0) throw new Error('Unknown status.');
  return withLock_(function () {
    var t = readTable_('Findings'), r = findRow_(t, id);
    setFields_(t, r._row, { status: status });
    return true;
  });
}

function addNote(token, appliesTo, text) {
  var me = who_(token);
  need_(me, ['Organizer']);
  text = String(text || '').trim();
  if (!text) throw new Error('Write the note first.');
  if (['Both', 'Comb', 'Waggle'].indexOf(appliesTo) < 0) appliesTo = 'Both';
  return withLock_(function () {
    appendRow_(readTable_('Notes'), { id: newId_('n_'), date: today_(), appliesTo: appliesTo, text: text.slice(0, 2000), createdAt: nowIso_() });
    return true;
  });
}

function deleteNote(token, id) {
  var me = who_(token);
  need_(me, ['Organizer']);
  return withLock_(function () {
    var t = readTable_('Notes'), r = findRow_(t, id);
    t.sh.deleteRow(r._row);
    return true;
  });
}

function addPerson(token, name, level, role) {
  var me = who_(token);
  need_(me, ['Organizer']);
  name = String(name || '').trim();
  if (!name) throw new Error('Add a name.');
  if (findMember_(name)) throw new Error('That name is already on the team.');
  return withLock_(function () {
    appendRow_(readTable_('Team'), { name: name.slice(0, 60), level: normLevel_(level), role: String(role || '') });
    return true;
  });
}

function setLevel(token, name, level) {
  var me = who_(token);
  need_(me, ['Organizer']);
  level = normLevel_(level);
  return withLock_(function () {
    var t = readTable_('Team');
    var r = t.rows.filter(function (m) { return m.name.toLowerCase() === String(name).toLowerCase(); })[0];
    if (!r) throw new Error('That person was not found.');
    if (normLevel_(r.level) === 'Organizer' && level !== 'Organizer') {
      var organizers = t.rows.filter(function (m) { return normLevel_(m.level) === 'Organizer' && !isInactive_(m); });
      if (organizers.length < 2) throw new Error('Keep at least one Organizer.');
    }
    setFields_(t, r._row, { level: level });
    return true;
  });
}

function setPassword(token, name, password) {
  var me = who_(token);
  need_(me, ['Organizer']);
  return setPassword_(name, password);
}

function changeMyPassword(token, current, next) {
  var me = who_(token);
  var m = findMember_(me.name);
  if (!m || hash_(m.salt, String(current || '')) !== m.hash) throw new Error('Your current password is not right.');
  return setPassword_(me.name, next, token);
}

function setPassword_(name, password, keepToken) {
  password = String(password || '');
  if (password.length < 6) throw new Error('Use at least 6 characters.');
  return withLock_(function () {
    var t = readTable_('Team');
    var r = t.rows.filter(function (m) { return m.name.toLowerCase() === String(name).toLowerCase(); })[0];
    if (!r) throw new Error('That person was not found.');
    var salt = newSalt_();
    setFields_(t, r._row, { salt: salt, hash: hash_(salt, password) });
    var keep = keepToken ? props_().getProperty('s_' + keepToken) : null;
    cleanSessions_(r.name);
    if (keep) props_().setProperty('s_' + keepToken, keep);
    return true;
  });
}

/* ---------- Spreadsheet menu ---------- */

function setPasswordFromMenu() {
  var ui = SpreadsheetApp.getUi();
  var a = ui.prompt('Set a password', 'Name (exactly as in the Team tab):', ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK) return;
  var b = ui.prompt('Set a password', 'New password for ' + a.getResponseText() + ' (6+ characters):', ui.ButtonSet.OK_CANCEL);
  if (b.getSelectedButton() !== ui.Button.OK) return;
  try { setPassword_(a.getResponseText().trim(), b.getResponseText()); ui.alert('Password updated.'); }
  catch (e) { ui.alert(e.message); }
}

/* ---------- One-time setup ---------- */

function setupHub() {
  var ss = SpreadsheetApp.getActive();
  Object.keys(TABLES).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    var spec = TABLES[name];
    var lastCol = sh.getLastColumn();
    var head = lastCol ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim().toLowerCase(); }) : [];
    spec.forEach(function (p) {
      if (head.indexOf(p[0].toLowerCase()) < 0) { head.push(p[0].toLowerCase()); sh.getRange(1, head.length).setValue(p[0]); }
    });
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold').setBackground('#2a0f1e').setFontColor('#fff4e6');
  });
  var first = ss.getSheetByName('Sheet1');
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
  if (typeof SEED !== 'undefined') seedHub_(SEED);
  var team = ss.getSheetByName('Team');
  var cols = readTable_('Team').cols;
  if (cols.salt >= 0) team.hideColumns(cols.salt + 1);
  if (cols.hash >= 0) team.hideColumns(cols.hash + 1);
  return 'Hive Operations Hub is set up.';
}

function seedHub_(seed) {
  Object.keys(seed).forEach(function (name) {
    var t = readTable_(name);
    if (t.rows.length) return;
    seed[name].forEach(function (o) { appendRow_(t, o); });
  });
}

/* ---------- JSON API (used by the page hosted on GitHub Pages) ---------- */

function apiFns_() {
  return {
    getNames: getNames, login: login, logout: logout, getData: getData,
    updateTask: updateTask, addTask: addTask, addTasks: addTasks, voiceTasks: voiceTasks, askAssistant: askAssistant, reassignTask: reassignTask, addEod: addEod,
    setFindingStatus: setFindingStatus, addNote: addNote, deleteNote: deleteNote,
    addPerson: addPerson, setLevel: setLevel, setPassword: setPassword, changeMyPassword: changeMyPassword
  };
}

function doPost(e) {
  var out;
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var fns = apiFns_();
    if (!req.fn || !Object.prototype.hasOwnProperty.call(fns, req.fn)) throw new Error('Unknown action.');
    var args = Array.isArray(req.args) ? req.args.slice(0, 6) : [];
    out = { ok: true, result: fns[req.fn].apply(null, args) };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// Run once from the editor after changing the code: grants permissions and checks Gemini and the Gem brain.
function testGemini() {
  Logger.log('Gem brain: ' + hiveBrain_().length + ' characters');
  Logger.log('Gemini says: ' + gemini_({ contents: [{ role: 'user', parts: [{ text: 'Reply with just: OK' }] }] }));
}

// Editor-only check of Hive AI and voice tasks without signing in (not reachable from the page).
function testHiveAi() {
  var me = { name: 'Fouz', level: 'Organizer' };
  var hive = askAssistant_(me, 'hive', [{ role: 'user', text: 'Who is stuck right now and who has not paid? Short table, then flags.' }]);
  Logger.log('HIVE MASTER OPS: ' + hive.reply.slice(0, 1500));
  var gem = askAssistant_(me, 'gemini', [{ role: 'user', text: 'In one sentence: what is A2P 10DLC registration?' }]);
  Logger.log('GEMINI: ' + gem.reply.slice(0, 600));
  var vt = voiceTasks_('', '', 'Kamran, clear the pending A2P approvals today and fix the negative balance email by Friday. Merab and Chris, each send three referral outreach messages tomorrow. Someone needs to call Anika Williams about her GoDaddy access.');
  vt.tasks.forEach(function (t) { Logger.log('TASK: ' + (t.person || '(nobody)') + ' | ' + t.task + ' | ' + t.client + ' | ' + t.deadline); });
}
