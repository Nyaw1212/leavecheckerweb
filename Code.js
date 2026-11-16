/** Leave Checker Web: read-only Apps Script lookup. Configuration is in Script Properties. */
const SHEET_NAME = 'LeaveBalances';
const SESSION_SECONDS = 1800;

function doGet(e) {
  try {
    const config = config_();
    if (config.anonymous) return page_('', '', '', true);
    if (e && e.parameter && e.parameter.error) return page_('', 'Sign-in was cancelled. Please try again.');
    if (e && e.parameter && e.parameter.code) {
      const session = finishSignIn_(e.parameter, config);
      return page_(session, '');
    }
    return page_('', '', signInUrl_(config));
  } catch (err) {
    console.error('Leave Checker unavailable: ' + err.message);
    return page_('', 'The app is not configured or sign-in could not be completed. Contact the administrator.');
  }
}

function page_(session, message, url, anonymous) {
  const page = HtmlService.createTemplateFromFile('Index');
  page.initialSession = session || '';
  page.anonymousMode = anonymous === true;
  page.initialMessage = message || '';
  page.signInUrl = url || '';
  page.appUrl = ScriptApp.getService().getUrl() || '';
  return page.evaluate().setTitle('Leave Balance Lookup').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function config_() {
  const p = PropertiesService.getScriptProperties();
  const config = {
    spreadsheet: p.getProperty('SPREADSHEET_ID'),
    anonymous: true, // Temporary: allow access without Google sign-in.
    client: p.getProperty('GOOGLE_CLIENT_ID'),
    secret: p.getProperty('GOOGLE_CLIENT_SECRET'),
    emails: (p.getProperty('ALLOWED_EMAILS') || '').split(',').map(x => x.trim().toLowerCase()).filter(Boolean),
    redirect: ScriptApp.getService().getUrl()
  };
  if (!config.spreadsheet || (!config.anonymous && (!config.client || !config.secret || !config.emails.length || !config.redirect))) throw Error('Missing configuration');
  return config;
}

function random_() { return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, ''); }
function signInUrl_(config) {
  const state = random_();
  CacheService.getScriptCache().put('oauth:' + state, '1', 600);
  const params = {
    client_id: config.client, redirect_uri: config.redirect, response_type: 'code',
    scope: 'openid email', state: state, prompt: 'select_account'
  };
  return 'https://accounts.google.com/o/oauth2/v2/auth?' + Object.keys(params).map(k => encodeURIComponent(k) + '=' + encodeURIComponent(params[k])).join('&');
}

function finishSignIn_(p, config) {
  const state = String(p.state || '');
  if (!/^[0-9a-f]{64}$/i.test(state)) throw Error('Invalid state');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const cache = CacheService.getScriptCache();
    if (!cache.get('oauth:' + state)) throw Error('Expired state');
    cache.remove('oauth:' + state);
  } finally { lock.releaseLock(); }
  const tokenResponse = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post', muteHttpExceptions: true,
    payload: { code: p.code, client_id: config.client, client_secret: config.secret, redirect_uri: config.redirect, grant_type: 'authorization_code' }
  });
  if (tokenResponse.getResponseCode() !== 200) throw Error('OAuth exchange failed');
  const accessToken = JSON.parse(tokenResponse.getContentText()).access_token;
  if (!accessToken) throw Error('No access token');
  const userResponse = UrlFetchApp.fetch('https://openidconnect.googleapis.com/v1/userinfo', {
    headers: { Authorization: 'Bearer ' + accessToken }, muteHttpExceptions: true
  });
  if (userResponse.getResponseCode() !== 200) throw Error('User verification failed');
  const user = JSON.parse(userResponse.getContentText());
  const email = String(user.email || '').toLowerCase();
  if (user.email_verified !== true || !config.emails.includes(email)) throw Error('Account not authorized');
  const session = random_();
  CacheService.getScriptCache().put('session:' + session, email, SESSION_SECONDS);
  return session;
}

/** Access follows the anonymous setting in config_(). */
function lookupBalance(session, employeeId) {
  const config = authorizedConfig_(session);
  employeeId = String(employeeId || '').trim();
  if (!employeeId || employeeId.length > 100 || !normalizeEmployeeId_(employeeId)) throw Error('Enter an employee number.');
  const sheet = SpreadsheetApp.openById(config.spreadsheet).getSheetByName(SHEET_NAME);
  if (!sheet) throw Error('The balance sheet is unavailable.');
  const data = sheet.getDataRange().getValues();
  const expected = ['employee_id', 'name', 'as_of', 'vl', 'sl', 'fl', 'wl', 'spl', 'updated_at'];
  const headings = (data.shift() || []).map(x => String(x).trim().toLowerCase());
  const indexes = expected.map(x => headings.indexOf(x));
  if (indexes.some(i => i < 0)) throw Error('The balance sheet is not configured correctly.');
  const idIndex = indexes[0];
  const idKey = normalizeEmployeeId_(employeeId);
  const matches = data.filter(r => normalizeEmployeeId_(r[idIndex]) === idKey);
  if (matches.length !== 1) return null;
  const row = matches[0];
  const result = {};
  expected.forEach((key, i) => { const v = row[indexes[i]]; result[key] = v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : (v === '' ? null : String(v)); });
  return result;
}

// Both name search and balance lookup enforce the same access settings.
function authorizedConfig_(session) {
  const config = config_();
  if (!config.anonymous) {
    if (!/^[0-9a-f]{64}$/i.test(String(session || ''))) throw Error('Please sign in again.');
    const email = CacheService.getScriptCache().get('session:' + session);
    if (!email || !config.emails.includes(email)) throw Error('Please sign in again.');
  }
  return config;
}

function searchPersonnel(session, query) {
  const config = authorizedConfig_(session);
  const q = normalize_(query);
  if (q.length > 200) throw Error('Search text is too long.');
  const sheet = SpreadsheetApp.openById(config.spreadsheet).getSheetByName(SHEET_NAME);
  if (!sheet) throw Error('The balance sheet is unavailable.');
  const rows = sheet.getDataRange().getValues();
  const headings = (rows.shift() || []).map(x => String(x).trim().toLowerCase());
  const ni = headings.indexOf('name'), ii = headings.indexOf('employee_id'), asOfIndex = headings.indexOf('as_of');
  const balanceKeys = ['vl', 'sl', 'fl', 'wl', 'spl'];
  const balanceIndexes = balanceKeys.map(key => headings.indexOf(key));
  if (ni < 0 || ii < 0 || balanceIndexes.some(i => i < 0)) throw Error('The balance sheet is not configured correctly.');
  const counts = {};
  rows.forEach(row => {
    const id = normalizeEmployeeId_(row[ii]);
    if (id) counts[id] = (counts[id] || 0) + 1;
  });
  const found = rows.filter(row => row[ni] && normalize_(row[ni]).includes(q))
    .map(row => {
      const id = String(row[ii] == null ? '' : row[ii]).trim();
      const key = normalizeEmployeeId_(id);
      const balances = {};
      balanceKeys.forEach((field, index) => {
        const value = row[balanceIndexes[index]];
        balances[field] = value == null || value === '' ? null : String(value);
      });
      const asOfValue = asOfIndex < 0 ? '' : row[asOfIndex];
      const asOf = asOfValue instanceof Date ? Utilities.formatDate(asOfValue, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(asOfValue == null ? '' : asOfValue);
      return {name: String(row[ni]), employeeId: id, selectable: !!key && counts[key] === 1, asOf, ...balances};
    }).sort((a, b) => a.name.localeCompare(b.name) || a.employeeId.localeCompare(b.employeeId));
  return {items: found, more: false};
}

function normalizeEmployeeId_(value) {
  return String(value == null ? '' : value).normalize('NFKC').replace(/[\s-]+/g, '').toLowerCase();
}

function normalize_(v) { return String(v || '').normalize('NFKC').trim().replace(/\s+/g, ' ').toLowerCase(); }
function signOut(session) { if (/^[0-9a-f]{64}$/i.test(String(session || ''))) CacheService.getScriptCache().remove('session:' + session); }

// MONE writes follow the web app's access settings.
function applyMone(session, input) {
  const config = authorizedConfig_(session);
  input = input || {};
  const id = normalizeEmployeeId_(input.employeeId);
  const requestId = String(input.requestId || '');
  const startDate = String(input.startDate || ''), endDate = String(input.endDate || '');
  function validDate(value) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
    const parsed = new Date(value + 'T00:00:00Z');
    return !isNaN(parsed.getTime()) && parsed.toISOString().slice(0, 10) === value;
  }
  const year = Number(startDate.slice(0, 4)), month = Number(startDate.slice(5, 7));
  function amount(value, signed) {
    if (value === '' || value === null || value === undefined) return 0;
    const n = Number(value);
    if (!Number.isFinite(n) || (!signed && n < 0) || Math.abs(n) > 100000 || Math.abs(n * 1000 - Math.round(n * 1000)) > 1e-6) throw Error('Enter nonnegative amounts with at most three decimal places.');
    return Math.round(n * 1000);
  }
  const adjusting = input.adjustment === true;
  const vlDeduction = amount(input.vl, adjusting), slDeduction = amount(input.sl, adjusting);
  if (!id || id.length > 100 || !/^[a-zA-Z0-9-]{16,80}$/.test(requestId) ||
      !validDate(startDate) || !validDate(endDate) || startDate > endDate ||
      !Number.isInteger(year) || year < 1900 || year > 2100 ||
      !Number.isInteger(month) || month < 1 || month > 12 || (!vlDeduction && !slDeduction)) {
    throw Error('Select an employee and enter valid start and end dates and a MONE amount.');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const book = SpreadsheetApp.openById(config.spreadsheet);
    const sheet = book.getSheetByName(SHEET_NAME);
    if (!sheet) throw Error('The balance sheet is unavailable.');
    const rows = sheet.getDataRange().getValues();
    const headers = (rows.shift() || []).map(x => String(x).trim().toLowerCase());
    const idIndex = headers.indexOf('employee_id'), nameIndex = headers.indexOf('name');
    const vlIndex = headers.indexOf('vl'), slIndex = headers.indexOf('sl'), asOfIndex = headers.indexOf('as_of');
    if ([idIndex, nameIndex, vlIndex, slIndex, asOfIndex].some(i => i < 0) || Math.abs(vlIndex - slIndex) !== 1) throw Error('The balance sheet must contain employee_id, name, as_of, and adjacent VL/SL columns.');
    const matches = rows.map((row, index) => ({row, index})).filter(x => normalizeEmployeeId_(x.row[idIndex]) === id);
    if (matches.length !== 1) throw Error('Employee number is missing or duplicated.');
    let log = book.getSheetByName('MoneLog');
    const baseHeaders = ['request_id','employee_id','name','month','year','vl_deducted','sl_deducted','vl_before','sl_before','vl_after','sl_after','recorded_at','status'];
    const logHeaders = baseHeaders.concat(['start_date', 'end_date']);
    if (!log) {
      log = book.insertSheet('MoneLog');
      log.getRange(1, 1, 1, logHeaders.length).setValues([logHeaders]);
    }
    const logRows = log.getDataRange().getValues();
    if (baseHeaders.some((h, i) => String((logRows[0] || [])[i] || '').trim() !== h)) throw Error('MoneLog headers do not match the expected format.');
    const existingDateHeaders = [13, 14].map(i => String((logRows[0] || [])[i] || '').trim());
    if (existingDateHeaders.every(value => !value)) {
      if (logRows.slice(1).some(row => row[13] != null && row[13] !== '' || row[14] != null && row[14] !== ''))
        throw Error('MoneLog date columns contain existing data. Review the sheet before adding date headers.');
      if (log.getMaxColumns() >= 15 && log.getRange(1, 14, Math.max(log.getLastRow(), 1), 2).getFormulas().some(row => row.some(Boolean)))
        throw Error('MoneLog date columns contain formulas. Review the sheet before adding date headers.');
      if (log.getMaxColumns() < 15) log.insertColumnsAfter(log.getMaxColumns(), 15 - log.getMaxColumns());
      log.getRange(1, 14, 1, 2).setValues([['start_date', 'end_date']]);
    } else if (existingDateHeaders[0] !== 'start_date' || existingDateHeaders[1] !== 'end_date') {
      throw Error('MoneLog date columns do not match the expected format.');
    }
    const previous = logRows.slice(1).find(row => String(row[0]) === requestId);
    if (previous) {
      if (normalizeEmployeeId_(previous[1]) !== id || Number(previous[3]) !== month || Number(previous[4]) !== year ||
          Math.round(Number(previous[5]) * 1000) !== vlDeduction || Math.round(Number(previous[6]) * 1000) !== slDeduction ||
          String(previous[13]) !== startDate || String(previous[14]) !== endDate)
        throw Error('This request ID has different details. Review the log before retrying.');
      if (previous[12] !== 'APPLIED') throw Error('This MONE entry is pending. Check MoneLog and balances before retrying.');
      return {name: String(matches[0].row[nameIndex]), vl: String(matches[0].row[vlIndex]), sl: String(matches[0].row[slIndex]), alreadyApplied: true};
    }
    if (logRows.slice(1).some(row => normalizeEmployeeId_(row[1]) === id && String(row[12]) === 'PENDING'))
      throw Error('This employee has a pending MONE entry. Check MoneLog and balances before recording another.');
    if (adjusting) {
      const dateValue = value => value instanceof Date ? Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(value == null ? '' : value).trim();
      let currentVl = 0, currentSl = 0;
      logRows.slice(1).forEach(row => {
        if (normalizeEmployeeId_(row[1]) !== id || String(row[12]).trim().toUpperCase() !== 'APPLIED' || dateValue(row[13]) !== startDate || dateValue(row[14]) !== endDate) return;
        if (!Number.isFinite(Number(row[5])) || !Number.isFinite(Number(row[6]))) throw Error('MoneLog has invalid amounts. Review the sheet.');
        currentVl += Math.round(Number(row[5]) * 1000); currentSl += Math.round(Number(row[6]) * 1000);
      });
      if (amount(input.expectedMoneVl) !== currentVl || amount(input.expectedMoneSl) !== currentSl) throw Error('MONE totals changed since they were displayed. Reload and review them.');
      if (currentVl + vlDeduction < 0 || currentSl + slDeduction < 0) throw Error('MONE totals cannot be negative.');
    }
    function balance(value, label) {
      if (value === '' || value === null || value === undefined || !Number.isFinite(Number(value))) throw Error(label + ' balance is unavailable. Review it before recording MONE.');
      return Math.round(Number(value) * 1000);
    }
    const match = matches[0], vlBefore = balance(match.row[vlIndex], 'VL'), slBefore = balance(match.row[slIndex], 'SL');
    const asOfValue = match.row[asOfIndex];
    const asOf = asOfValue instanceof Date ? Utilities.formatDate(asOfValue, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(asOfValue || '').trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(asOf) || year * 100 + month > Number(asOf.slice(0, 4)) * 100 + Number(asOf.slice(5, 7)))
      throw Error('MONE month must not be later than the snapshot month. Export a newer balance first.');
    if (amount(input.expectedVl) !== vlBefore || amount(input.expectedSl) !== slBefore)
      throw Error('The balance changed since it was displayed. Reload the employee and review it again.');
    if (vlDeduction > vlBefore || slDeduction > slBefore) throw Error('A MONE deduction exceeds the available balance.');
    const rowNumber = match.index + 2, startCol = Math.min(vlIndex, slIndex) + 1;
    const range = sheet.getRange(rowNumber, startCol, 1, 2);
    if (range.getFormulas()[0].some(Boolean)) throw Error('VL or SL is a formula. Review the sheet before recording MONE.');
    const after = [0, 0];
    after[vlIndex - (startCol - 1)] = (vlBefore - vlDeduction) / 1000;
    after[slIndex - (startCol - 1)] = (slBefore - slDeduction) / 1000;
    const logRow = log.getLastRow() + 1;
    log.getRange(logRow, 14, 1, 2).setNumberFormat('@');
    log.getRange(logRow, 1, 1, logHeaders.length).setValues([[
      requestId, String(match.row[idIndex]), String(match.row[nameIndex]), month, year,
      vlDeduction / 1000, slDeduction / 1000, vlBefore / 1000, slBefore / 1000,
      after[vlIndex - (startCol - 1)], after[slIndex - (startCol - 1)], new Date(), 'PENDING', startDate, endDate
    ]]);
    SpreadsheetApp.flush();
    range.setValues([after]);
    SpreadsheetApp.flush();
    log.getRange(logRow, 13).setValue('APPLIED');
    SpreadsheetApp.flush();
    return {name: String(match.row[nameIndex]), vl: String(after[vlIndex - (startCol - 1)]), sl: String(after[slIndex - (startCol - 1)]), alreadyApplied: false};
  } finally {
    lock.releaseLock();
  }
}


/** Process a group of individual, idempotent requests in one server call.
 * Each entry retains applyMone's lock, validation and pending-log recovery.
 * The group stops at the first failure; it is not an atomic transaction.
 */
function applyMoneBatch(session, inputs) {
  authorizedConfig_(session);
  if (!Array.isArray(inputs) || !inputs.length || inputs.length > 10) throw Error('A save group must contain 1 to 10 entries.');
  const ids = inputs.map(input => String(input && input.requestId || ''));
  if (new Set(ids).size !== ids.length) throw Error('Duplicate request IDs in save group.');
  const results = [];
  for (const input of inputs) {
    try { results.push({requestId:String(input.requestId), result:applyMone(session, input)}); }
    catch (error) { return {results, error:{requestId:String(input && input.requestId || ''), message:String(error && error.message || 'Save failed.')}}; }
  }
  return {results, error:null};
}



// Totals are for confirmed entries with the exact selected date range. Older
// month-only records cannot be safely assigned to a specific range.
function getRecordedMone(session, startDate, endDate) {
  const config = authorizedConfig_(session);
  const isoDate = value => {
    if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    return String(value == null ? '' : value).trim();
  };
  const valid = value => /^\d{4}-\d{2}-\d{2}$/.test(value) && !isNaN(new Date(value + 'T00:00:00Z').getTime()) && new Date(value + 'T00:00:00Z').toISOString().slice(0, 10) === value;
  startDate = String(startDate || ''); endDate = String(endDate || '');
  if (!valid(startDate) || !valid(endDate) || startDate > endDate) throw Error('Choose a valid MONE date range.');
  const log = SpreadsheetApp.openById(config.spreadsheet).getSheetByName('MoneLog');
  if (!log) return {};
  const rows = log.getDataRange().getValues(), headers = (rows.shift() || []).map(x => String(x).trim().toLowerCase());
  const fields = ['employee_id', 'vl_deducted', 'sl_deducted', 'status', 'start_date', 'end_date'];
  const indexes = fields.map(field => headers.indexOf(field));
  if (indexes.slice(0, 4).some(i => i < 0)) throw Error('MoneLog headers do not match the expected format.');
  if (indexes[4] < 0 && indexes[5] < 0) return {};
  if (indexes[4] < 0 || indexes[5] < 0) throw Error('MoneLog date columns are incomplete.');
  const totals = Object.create(null);
  rows.forEach(row => {
    if (String(row[indexes[3]]).trim().toUpperCase() !== 'APPLIED' || isoDate(row[indexes[4]]) !== startDate || isoDate(row[indexes[5]]) !== endDate) return;
    const id = normalizeEmployeeId_(row[indexes[0]]), vl = Number(row[indexes[1]]), sl = Number(row[indexes[2]]);
    if (!id || !Number.isFinite(vl) || !Number.isFinite(sl)) return;
    if (!totals[id]) totals[id] = {vl: 0, sl: 0};
    totals[id].vl += Math.round(vl * 1000); totals[id].sl += Math.round(sl * 1000);
  });
  Object.keys(totals).forEach(id => { totals[id].vl /= 1000; totals[id].sl /= 1000; });
  return totals;
}


/** Save a report-only snapshot of all pasted rows; never updates leave balances. */
function saveMoneReportBatch(session, input) {
  const config = authorizedConfig_(session);
  input = input || {};
  const batchId = String(input.batchId || ''), office = String(input.office || '').trim();
  const start = String(input.startDate || ''), end = String(input.endDate || ''), rows = input.rows;
  function dateOk(v) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(v)) return false;
    const d = new Date(v + 'T00:00:00Z');
    return !isNaN(d.getTime()) && d.toISOString().slice(0, 10) === v;
  }
  if (!/^[a-zA-Z0-9-]{16,100}$/.test(batchId) || office.length > 150 || !dateOk(start) || !dateOk(end) || start > end || !Array.isArray(rows) || !rows.length || rows.length > 500) throw Error('Enter a valid date range and 1 to 500 preview rows; office may be blank.');
  const fields = ['number','rank','inputName','matchedName','employeeId','category','requested','currentVl','currentSl','moneVl','moneSl','previewVl','previewSl','tag','manuallySelected','detail'];
  const tags = ['READY','INVALID_INPUT','NAME_REVIEW','DUPLICATE','EMPLOYEE_ID_REVIEW','NO_LEAVE_CARD','BALANCE_UNAVAILABLE','INSUFFICIENT_BALANCE','REQUEST_ADJUSTED','NEEDS_REVIEW'];
  const groups = ['Commissioned','Non-commissioned','Civilian','Unclassified'];
  const cleanRows = rows.map(row => {
    if (!row || typeof row !== 'object' || !groups.includes(row.category) || !tags.includes(row.tag)) throw Error('Invalid personnel group or review tag. Refresh the preview.');
    const output = {};
    fields.forEach(key => {
      const value = key === 'manuallySelected' ? (row[key] === true ? 'YES' : 'NO') : String(row[key] == null ? '' : row[key]);
      if (value.length > 500) throw Error('A report field is too long.');
      output[key] = value;
    });
    return output;
  });
  const canonical = JSON.stringify({office,start,end,rows:cleanRows});
  const fingerprint = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, canonical, Utilities.Charset.UTF_8));
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const book = SpreadsheetApp.openById(config.spreadsheet);
    let sheet = book.getSheetByName('MoneBatchData');
    const headers = ['batch_id','fingerprint','saved_at','office','start_date','end_date'].concat(fields);
    if (!sheet) sheet = book.insertSheet('MoneBatchData');
    if (sheet.getMaxColumns() < headers.length) sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
    if (!sheet.getLastRow()) sheet.getRange(1,1,1,headers.length).setValues([headers]);
    const values = sheet.getDataRange().getValues();
    if (headers.some((value,index) => String((values[0] || [])[index] || '').trim() !== value)) throw Error('MoneBatchData headers do not match. Review the sheet.');
    const prior = values.slice(1).filter(row => String(row[0]) === batchId);
    if (prior.length) {
      if (prior.length !== cleanRows.length || prior.some(row => String(row[1]) !== fingerprint)) throw Error('Batch ID already exists with different or incomplete data. Review MoneBatchData.');
      return {batchId,rowCount:prior.length,alreadySaved:true};
    }
    // Mark possible prior-batch duplicates for the same exact MONE period.
    // Employee number is authoritative when both rows have one. Name-only
    // comparisons are possible matches, not proof of identity.
    const dateValue = value => value instanceof Date ? Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(value == null ? '' : value).trim();
    const byId = new Map(), byName = new Map();
    values.slice(1).forEach(priorRow => {
      if (String(priorRow[0]) === batchId || dateValue(priorRow[4]) !== start || dateValue(priorRow[5]) !== end) return;
      const item = {batch:String(priorRow[0]), id:normalizeEmployeeId_(priorRow[10]), name:normalize_(priorRow[8])};
      if (!item.batch) return;
      if (item.id && !byId.has(item.id)) byId.set(item.id,item);
      if (item.name) { if (!byName.has(item.name)) byName.set(item.name,[]); byName.get(item.name).push(item); }
    });
    // Sheets can interpret leading = as a formula. Keep all imported data as text.
    const safe = value => /^[\s]*[=+\-@]/.test(String(value)) ? "'" + String(value) : value;
    const stamp = new Date();
    const output = cleanRows.map(row => {
      const id = normalizeEmployeeId_(row.employeeId), name = normalize_(row.inputName);
      const matchById = id && byId.get(id);
      const matchByName = name && (byName.get(name) || []).find(item => !id || !item.id);
      const match = matchById || matchByName;
      const detail = match ? ('POSSIBLE_DUPLICATE_PREVIOUS_BATCH: ' + match.batch + ' (' + (matchById ? 'employee ID' : 'name; verify identity') + '). ' + row.detail).slice(0,500) : row.detail;
      return [batchId,fingerprint,stamp,safe(office),start,end].concat(fields.map(key => safe(key === 'detail' ? detail : row[key])));
    });
    const finalRow = sheet.getLastRow() + output.length;
    if (finalRow > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), finalRow - sheet.getMaxRows());
    sheet.getRange(sheet.getLastRow()+1,1,output.length,headers.length).setValues(output);
    SpreadsheetApp.flush();
    return {batchId,rowCount:output.length,alreadySaved:false};
  } finally { lock.releaseLock(); }
}


/** Read user-selected DOCX/XLSX table data into the existing MONE preview. */
function parseMoneAttachment(session, fileName, base64) {
  authorizedConfig_(session);
  fileName = String(fileName || '').trim(); base64 = String(base64 || '');
  const type = /\.(docx|xlsx)$/i.exec(fileName);
  if (!type || !base64 || base64.length > 4 * 1024 * 1024 || !/^[A-Za-z0-9+/]+={0,2}$/.test(base64)) throw Error('Choose a DOCX or XLSX file no larger than 3 MB.');
  let entries;
  try { entries = Utilities.unzip(Utilities.newBlob(Utilities.base64Decode(base64), 'application/zip', fileName)); }
  catch (_) { throw Error('The file could not be opened. Check that it is an unencrypted DOCX or XLSX file.'); }
  if (entries.length > 1000) throw Error('The file has too many parts to import.');
  const parts = Object.create(null); let totalBytes = 0;
  entries.forEach(part => { totalBytes += part.getBytes().length; parts[part.getName()] = part; });
  if (totalBytes > 20 * 1024 * 1024) throw Error('The file is too large after decompression.');
  const child = (element, name) => element.getChildren().filter(node => node.getName() === name);
  const descendants = (element, name) => {
    const result = [];
    function visit(node) { node.getChildren().forEach(item => { if (item.getName() === name) result.push(item); visit(item); }); }
    visit(element); return result;
  };
  const attribute = (element, name) => { const attr = element.getAttributes().find(item => item.getName() === name); return attr ? attr.getValue() : ''; };
  const text = element => descendants(element, 't').map(node => node.getValue()).join('');
  const readXml = part => { try { return XmlService.parse(part.getDataAsString('UTF-8')).getRootElement(); } catch (_) { throw Error('The file contains XML that could not be read.'); } };
  let tables = [];
  if (type[1].toLowerCase() === 'docx') {
    if (!parts['word/document.xml']) throw Error('The Word document has no readable document content.');
    const root = readXml(parts['word/document.xml']);
    tables = descendants(root, 'tbl').map(table => child(table, 'tr').map(row => {
      const result = [];
      child(row, 'tc').forEach(cell => {
        const value = text(cell).replace(/\s+/g, ' ').trim();
        const properties = child(cell, 'tcPr')[0], spanNode = properties && child(properties, 'gridSpan')[0];
        const span = spanNode ? Number(attribute(spanNode, 'val')) : 1;
        if (!Number.isInteger(span) || span < 1 || span > 50) throw Error('The Word table has an unsupported merged cell.');
        for (let i = 0; i < span; i++) result.push(value);
      });
      return result;
    }));
  } else {
    const shared = parts['xl/sharedStrings.xml'] ? child(readXml(parts['xl/sharedStrings.xml']), 'si').map(text) : [];
    const sheets = Object.keys(parts).filter(name => /^xl\/worksheets\/sheet\d+\.xml$/.test(name)).sort((a,b) => a.localeCompare(b, undefined, {numeric:true}));
    if (!sheets.length) throw Error('The Excel file has no readable worksheets.');
    tables = sheets.map(name => descendants(readXml(parts[name]), 'row').map(row => {
      const values = [];
      child(row, 'c').forEach(cell => {
        const ref = attribute(cell, 'r'), letters = /^[A-Z]+/i.exec(ref);
        if (!letters) return;
        let index = 0;
        for (const letter of letters[0].toUpperCase()) index = index * 26 + letter.charCodeAt(0) - 64;
        if (index < 1 || index > 100) return;
        const kind = attribute(cell, 't'), valueNode = child(cell, 'v')[0];
        let value = kind === 'inlineStr' ? text(cell) : valueNode ? valueNode.getValue() : '';
        if (kind === 's') value = shared[Number(value)] == null ? '' : shared[Number(value)];
        values[index - 1] = String(value || '').replace(/\s+/g, ' ').trim();
      });
      return values;
    }));
  }
  const found = [];
  tables.forEach(table => table.forEach(row => {
    for (let i = 0; i <= Math.min(9, row.length - 1); i++) {
      const serial = String(row[i] || '').trim(), rank = String(row[i+1] || '').trim();
      const numbered = /^\d+\.?$/.test(serial);
      // Some official forms leave the No. column blank for every person.
      // Require a rank containing a digit for this case, so headings are skipped.
      const unnumbered = !serial && /^[a-z]{1,8}\s*\d[a-z0-9.\/-]*$/i.test(rank);
      if ((!numbered && !unnumbered) || !/^[a-z][a-z0-9 .\/-]{1,24}$/i.test(rank) || !/[a-z]/i.test(rank) || !String(row[i+2] || '').trim() || !String(row[i+3] || '').trim()) continue;
      found.push(Array.from({length:7}, (_, column) => String(row[i+column] == null ? '' : row[i+column]).replace(/[\t\r\n]+/g, ' ').trim()));
      break;
    }
  }));
  if (!found.length) throw Error('No supported personnel rows were found. Check the rank and name columns.');
  if (found.length > 500) throw Error('This file contains more than 500 personnel rows. Split it into smaller batches.');
  return {fileName, rows:found, count:found.length, missingAmounts:found.filter(row => !row[5]).length};
}

// Imported batches are stored in the existing spreadsheet, not in browser storage.
function moneReviewSheet_(config) {
  const book = SpreadsheetApp.openById(config.spreadsheet);
  let sheet = book.getSheetByName('MoneReviewQueue');
  const headers = ['batch_id','file_name','created_at','status','rows_json'];
  if (!sheet) sheet = book.insertSheet('MoneReviewQueue');
  if (!sheet.getLastRow()) sheet.getRange(1,1,1,headers.length).setValues([headers]);
  const actual = sheet.getRange(1,1,1,headers.length).getValues()[0];
  if (headers.some((h,i) => actual[i] !== h)) throw Error('MoneReviewQueue headers do not match.');
  return sheet;
}
function saveMoneReviewImport(session, batchId, fileName, rows) {
  const config = authorizedConfig_(session);
  batchId = String(batchId || ''); fileName = String(fileName || '').trim();
  if (!/^[a-zA-Z0-9-]{16,100}$/.test(batchId) || !fileName || fileName.length > 250 || !Array.isArray(rows) || !rows.length || rows.length > 500 || rows.some(row => !Array.isArray(row) || row.length !== 7 || row.some(v => typeof v !== 'string' || v.length > 1000))) throw Error('Invalid imported batch.');
  const json = JSON.stringify(rows);
  if (json.length > 45000) throw Error('This batch is too large for the review queue. Split the file into smaller batches.');
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sheet = moneReviewSheet_(config), values = sheet.getDataRange().getValues();
    const existing = values.slice(1).find(row => String(row[0]) === batchId);
    if (existing) {
      if (String(existing[4]) !== json) throw Error('Batch ID already exists with different data.');
      return {batchId, alreadySaved:true};
    }
    // Prefix potentially formula-like filenames. JSON always begins with '['.
    const safeName = /^[\s]*[=+\-@]/.test(fileName) ? "'" + fileName : fileName;
    sheet.appendRow([batchId,safeName,new Date(),'PENDING',json]);
    return {batchId,alreadySaved:false};
  } finally { lock.releaseLock(); }
}
function listMoneReviewImports(session) {
  const config = authorizedConfig_(session), sheet = SpreadsheetApp.openById(config.spreadsheet).getSheetByName('MoneReviewQueue');
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues(), expected = ['batch_id','file_name','created_at','status','rows_json'];
  if (expected.some((h,i) => (values[0] || [])[i] !== h)) throw Error('MoneReviewQueue headers do not match.');
  return values.slice(1).reverse().filter(row => row[0] && row[3] === 'PENDING').map(row => ({batchId:String(row[0]),fileName:String(row[1]),createdAt:row[2] instanceof Date ? row[2].toISOString() : String(row[2] || '')}));
}
function loadMoneReviewImport(session, batchId) {
  const config = authorizedConfig_(session), sheet = SpreadsheetApp.openById(config.spreadsheet).getSheetByName('MoneReviewQueue');
  if (!sheet) throw Error('Review queue is empty.');
  const rows = sheet.getDataRange().getValues(), match = rows.slice(1).find(row => String(row[0]) === String(batchId));
  if (!match || match[3] !== 'PENDING') throw Error('Batch is unavailable or already reviewed.');
  const data = JSON.parse(String(match[4]));
  return {batchId:String(match[0]),fileName:String(match[1]),rows:data};
}
function markMoneReviewImportReviewed(session, batchId) {
  const config = authorizedConfig_(session), lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.openById(config.spreadsheet).getSheetByName('MoneReviewQueue');
    if (!sheet) throw Error('Review queue is empty.');
    const values = sheet.getDataRange().getValues(), index = values.findIndex((row,i) => i > 0 && String(row[0]) === String(batchId));
    if (index < 0) throw Error('Batch not found.');
    sheet.getRange(index + 1,4).setValue('REVIEWED');
    return {batchId:String(batchId)};
  } finally { lock.releaseLock(); }
}

// Drive access is restricted to the configured folder and follows the app access setting.
function moneDriveContext_(session) {
  const config = authorizedConfig_(session);
  const folderId = String(PropertiesService.getScriptProperties().getProperty('MONE_FOLDER_ID') || '').trim();
  if (!/^[a-zA-Z0-9_-]{10,200}$/.test(folderId)) throw Error('Set MONE_FOLDER_ID in Apps Script project properties.');
  return {config,folder:DriveApp.getFolderById(folderId)};
}
function moneDriveFiles_(folder) {
  const iter = folder.getFiles(), files = []; let scanned = 0;
  while (iter.hasNext()) {
    if (++scanned > 1000) throw Error('Folder contains more than 1,000 files. Move older files to another folder before refreshing.');
    const file = iter.next(), name = file.getName();
    if (/\.(docx|xlsx)$/i.test(name) && !file.isTrashed()) files.push(file);
  }
  return files;
}
function moneDriveReviewData_(config) {
  const sheet = SpreadsheetApp.openById(config.spreadsheet).getSheetByName('MoneDriveReviews');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues(), headers = ['file_id','modified_ms','reviewed_at'];
  if (headers.some((header,index) => (data[0] || [])[index] !== header)) throw Error('MoneDriveReviews headers do not match.');
  return data.slice(1);
}
function listMoneDriveFiles(session) {
  const context = moneDriveContext_(session), reviewed = new Set(moneDriveReviewData_(context.config).map(r => String(r[0]) + ':' + String(r[1])));
  return moneDriveFiles_(context.folder).map(f => {
    const version = f.getLastUpdated().getTime(), id = f.getId();
    return {id,name:f.getName(),modified:version,size:f.getSize(),status:reviewed.has(id + ':' + version) ? 'REVIEWED' : 'PENDING'};
  }).sort((a,b) => (a.status === b.status ? b.modified-a.modified : a.status === 'PENDING' ? -1 : 1));
}
function moneDriveFile_(context, fileId) {
  if (!/^[a-zA-Z0-9_-]{10,200}$/.test(String(fileId || ''))) throw Error('Invalid file selection.');
  const match = moneDriveFiles_(context.folder).find(f => f.getId() === fileId);
  if (!match) throw Error('The selected DOCX or XLSX file is no longer in the configured folder.');
  return match;
}
function importMoneDriveFile(session, fileId, expectedModified) {
  const context = moneDriveContext_(session), file = moneDriveFile_(context,fileId), version = file.getLastUpdated().getTime();
  if (Number(expectedModified) !== version) throw Error('This file changed. Refresh the folder list before importing.');
  if (file.getSize() <= 0 || file.getSize() > 3 * 1024 * 1024) throw Error('File must be between 1 byte and 3 MB.');
  const result = parseMoneAttachment(session,file.getName(),Utilities.base64Encode(file.getBlob().getBytes()));
  if (file.getLastUpdated().getTime() !== version) throw Error('This file changed while being read. Refresh and try again.');
  return result;
}
function markMoneDriveReviewed(session, fileId, expectedModified) {
  const context = moneDriveContext_(session), file = moneDriveFile_(context,fileId), version = file.getLastUpdated().getTime();
  if (Number(expectedModified) !== version) throw Error('This file changed. Refresh the folder list before marking it reviewed.');
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const book = SpreadsheetApp.openById(context.config.spreadsheet);
    let sheet = book.getSheetByName('MoneDriveReviews');
    const headers = ['file_id','modified_ms','reviewed_at'];
    if (!sheet) sheet = book.insertSheet('MoneDriveReviews');
    if (!sheet.getLastRow()) sheet.getRange(1,1,1,headers.length).setValues([headers]);
    const rows = moneDriveReviewData_(context.config);
    if (!rows.some(r => String(r[0]) === fileId && Number(r[1]) === version)) sheet.appendRow([fileId,version,new Date()]);
    return {reviewed:true};
  } finally { lock.releaseLock(); }
}


// Controlled CSV import: rows with any recorded MONE are protected from replacement.
function balanceImportPlan_(config, csvText) {
  csvText = String(csvText || '').replace(/^\uFEFF/, '');
  if (!csvText || csvText.length > 1500000) throw Error('Choose a CSV file up to 1.5 MB.');
  let input;
  try { input = Utilities.parseCsv(csvText); } catch (e) { throw Error('Could not read this CSV. Export it again from the desktop app.'); }
  const fields = ['employee_id','name','as_of','vl','sl','fl','wl','spl','updated_at'];
  const csvHeaders = (input.shift() || []).map(x => String(x || '').trim().toLowerCase());
  if (csvHeaders.length !== fields.length || fields.some((x,i) => csvHeaders[i] !== x)) throw Error('The CSV must use the Leave Checker export columns in the original order.');
  input = input.filter(row => row.some(v => String(v || '').trim()));
  if (!input.length || input.length > 5000) throw Error('The CSV must contain 1 to 5000 data rows.');
  const book = SpreadsheetApp.openById(config.spreadsheet);
  const sheet = book.getSheetByName(SHEET_NAME);
  if (!sheet) throw Error('LeaveBalances is unavailable.');
  const rowCount = Math.max(sheet.getLastRow(), 1);
  const values = sheet.getRange(1,1,rowCount,9).getValues();
  const formulas = sheet.getRange(1,1,rowCount,9).getFormulas();
  if (fields.some((x,i) => String((values[0] || [])[i] || '').trim().toLowerCase() !== x)) throw Error('LeaveBalances columns differ from the expected export order; no rows changed.');
  const log = book.getSheetByName('MoneLog');
  const logRows = log ? log.getDataRange().getValues() : [];
  const protectedIds = new Set();
  if (logRows.length) {
    const head = logRows[0].map(x => String(x || '').trim().toLowerCase());
    const idIndex = head.indexOf('employee_id'), statusIndex = head.indexOf('status');
    if (idIndex < 0 || statusIndex < 0) throw Error('MoneLog headers could not be verified; import stopped.');
    logRows.slice(1).forEach(row => {
      const id = normalizeEmployeeId_(row[idIndex]), status = String(row[statusIndex] || '').trim().toUpperCase();
      if (id && status !== 'APPLIED' && status !== 'PENDING') throw Error('MoneLog has an entry with an unknown status; review it before importing.');
      if (id && (status === 'APPLIED' || status === 'PENDING')) protectedIds.add(id);
      if (!id && row.some(v => String(v || '').trim())) throw Error('MoneLog has a record without an employee number; review it before importing.');
    });
  }
  const currentById = Object.create(null), currentNames = [];
  values.slice(1).forEach((row,i) => {
    const id = normalizeEmployeeId_(row[0]), name = normalize_(row[1]);
    if (id) (currentById[id] || (currentById[id] = [])).push(i+2);
    if (name) currentNames.push({id,name});
  });
  const inputCounts = Object.create(null), inputNames = Object.create(null), inputBySurname = Object.create(null);
  input.forEach(row => {
    const id = normalizeEmployeeId_(row[0]), name = normalize_(row[1]);
    if (id) inputCounts[id] = (inputCounts[id] || 0) + 1;
    if (name) {
      (inputNames[name] || (inputNames[name] = new Set())).add(id);
      const surname = name.includes(',') ? name.split(',')[0].trim() : name;
      (inputBySurname[surname] || (inputBySurname[surname] = [])).push({id,name});
    }
  });
  function similarName(a,b) {
    if (a === b) return true;
    const x = a.split(','), y = b.split(',');
    if (x.length !== 2 || y.length !== 2 || x[0].trim() !== y[0].trim()) return false;
    const ag = x[1].trim(), bg = y[1].trim();
    return !!ag && !!bg && (ag.startsWith(bg + ' ') || bg.startsWith(ag + ' '));
  }
  const results = [], writes = [];
  input.forEach((row,index) => {
    const idText = String(row[0] || '').trim(), id = normalizeEmployeeId_(idText);
    const name = String(row[1] || '').trim(), date = String(row[2] || '').trim();
    const asDate = /^\d{4}-\d{2}-\d{2}$/.test(date) ? new Date(date + 'T00:00:00Z') : null;
    let status = '', target = null;
    if (row.length !== 9 || !id || !/^[a-z0-9 -]+$/i.test(idText) || /^man-/i.test(idText) || !name || /^[=+\-@]/.test(name) || /^[=+\-@]/.test(String(row[8] || '').trim())) status = 'Review: missing or unsafe ID/name or incorrect columns';
    else if (!asDate || isNaN(asDate.getTime()) || asDate.toISOString().slice(0,10) !== date) status = 'Review: invalid as-of date';
    else if (row.slice(3,8).some(v => String(v).trim() !== '' && !Number.isFinite(Number(v)))) status = 'Review: invalid leave balance';
    else if (inputCounts[id] !== 1) status = 'Review: duplicate employee number in CSV';
    else if (inputNames[normalize_(name)].size > 1 || (inputBySurname[normalize_(name).split(',')[0].trim()] || []).some(other => other.id !== id && similarName(other.name, normalize_(name)))) status = 'Review: similar names have different employee numbers in CSV';
    else if (protectedIds.has(id)) status = 'Protected: MONE recorded; existing balance unchanged';
    else if ((currentById[id] || []).length > 1) status = 'Review: duplicate employee number in sheet';
    else if ((currentById[id] || []).length === 1) {
      target = currentById[id][0];
      const oldDate = values[target-1][2] instanceof Date ? Utilities.formatDate(values[target-1][2], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(values[target-1][2] || '').trim();
      if (normalize_(values[target-1][1]) !== normalize_(name)) status = 'Review: employee number has a different name';
      else if (formulas[target-1].some(Boolean)) status = 'Review: sheet row contains a formula';
      else if (!/^\d{4}-\d{2}-\d{2}$/.test(oldDate)) status = 'Review: existing as-of date is missing or invalid';
      else if (oldDate > date) status = 'Review: CSV balance is older than the sheet';
      else status = 'Update existing unprotected row';
    } else if (currentNames.some(item => similarName(item.name, normalize_(name)))) status = 'Review: similar existing name; check for a blank ID or duplicate';
    else status = 'Add new employee';
    if (status === 'Add new employee' || status === 'Update existing unprotected row') {
      const clean = [idText, name, date].concat(row.slice(3,8).map(v => String(v).trim() === '' ? '' : Number(v)), [String(row[8] || '').trim()]);
      writes.push({target, row:clean});
    }
    results.push({line:index+2, name, employeeId:idText, status});
  });
  // Include the underlying snapshot and CSV in the token; stale previews cannot be applied.
  const raw = JSON.stringify({csvText, values, formulas, logRows});
  const token = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8)
    .map(v => ('0' + ((v+256)%256).toString(16)).slice(-2)).join('');
  return {sheet, results, writes, token};
}

function previewBalanceImport(session, csvText) {
  const config = authorizedConfig_(session);
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const plan = balanceImportPlan_(config, csvText);
    return {token:plan.token, rows:plan.results, eligible:plan.writes.length};
  } finally { lock.releaseLock(); }
}

function saveBalanceImport(session, csvText, token) {
  const config = authorizedConfig_(session);
  if (!/^[a-f0-9]{64}$/.test(String(token || ''))) throw Error('Preview the file before saving.');
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const plan = balanceImportPlan_(config, csvText);
    if (plan.token !== token) throw Error('The sheet, MONE log, or CSV changed. Preview again before saving.');
    if (!plan.writes.length) return {saved:0};
    const sheet = plan.sheet;
    let next = sheet.getLastRow() + 1, saved = 0;
    const needed = plan.writes.filter(item => item.target === null).length;
    if (next + needed - 1 > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), next + needed - 1 - sheet.getMaxRows());
    try {
      plan.writes.forEach(item => {
        const row = item.target === null ? next++ : item.target;
        // Ensure IDs such as 002-1020 are kept as text.
        sheet.getRange(row, 1).setNumberFormat('@');
        sheet.getRange(row, 1, 1, 9).setValues([item.row]);
        saved++;
      });
      SpreadsheetApp.flush();
    } catch (error) {
      throw Error('Import stopped after ' + saved + ' row(s) were written. Review the sheet and preview again before retrying. ' + error.message);
    }
    return {saved};
  } finally { lock.releaseLock(); }
}


const MONE_NAME_MAP_SHEET = 'MoneNameMappings';
const MONE_NAME_MAP_HEADERS = ['input_key','employee_id','employee_name','recorded_at'];
function moneNameKey_(input) {
  if (!Array.isArray(input) || input.length !== 3) throw Error('The input name is incomplete.');
  const pieces = input.map(value => String(value || '').normalize('NFKC').trim().replace(/\s+/g, ' ').toLowerCase());
  if (!pieces[0] || !pieces[1] || pieces.some(value => value.length > 150)) throw Error('Enter a valid surname and first name.');
  return JSON.stringify(pieces);
}
function moneMapSheet_(book, create) {
  let sheet = book.getSheetByName(MONE_NAME_MAP_SHEET);
  if (!sheet && create) {
    sheet = book.insertSheet(MONE_NAME_MAP_SHEET);
    sheet.getRange(1, 1, 1, MONE_NAME_MAP_HEADERS.length).setValues([MONE_NAME_MAP_HEADERS]);
  }
  if (!sheet) return null;
  const rows = sheet.getDataRange().getValues();
  if (MONE_NAME_MAP_HEADERS.some((value, index) => String((rows[0] || [])[index] || '').trim() !== value)) throw Error('MoneNameMappings headers need review.');
  return sheet;
}
function getMoneNameMappings(session) {
  const config = authorizedConfig_(session);
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sheet = moneMapSheet_(SpreadsheetApp.openById(config.spreadsheet), false);
    if (!sheet) return [];
    const rows = sheet.getDataRange().getValues().slice(1), seen = Object.create(null), result = [];
    rows.forEach(row => { const key = String(row[0] || ''); if (key) seen[key] = (seen[key] || 0) + 1; });
    rows.forEach(row => {
      const key = String(row[0] || ''), employeeId = String(row[1] || '').trim(), employeeName = String(row[2] || '').trim();
      if (key && seen[key] !== 1) { if (!result.some(item => item.key === key)) result.push({key, invalid:true}); }
      else if (key && employeeId && employeeName) result.push({key, employeeId, employeeName});
      else if (key) result.push({key, invalid:true});
    });
    return result;
  } finally { lock.releaseLock(); }
}
function rememberMoneNameMapping(session, input, employeeId, expectedName) {
  const config = authorizedConfig_(session), key = moneNameKey_(input);
  employeeId = String(employeeId || '').trim(); expectedName = String(expectedName || '').trim();
  if (employeeId && (!/^[a-z0-9 -]{1,100}$/i.test(employeeId) || !expectedName || /^[=+\-@]/.test(expectedName))) throw Error('The selected employee is invalid.');
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const book = SpreadsheetApp.openById(config.spreadsheet);
    if (employeeId) {
      const balance = book.getSheetByName(SHEET_NAME);
      if (!balance) throw Error('LeaveBalances is unavailable.');
      const rows = balance.getDataRange().getValues(), header = (rows.shift() || []).map(v => String(v || '').trim().toLowerCase());
      const idIndex = header.indexOf('employee_id'), nameIndex = header.indexOf('name');
      if (idIndex < 0 || nameIndex < 0) throw Error('LeaveBalances headers need review.');
      const matches = rows.filter(row => normalizeEmployeeId_(row[idIndex]) === normalizeEmployeeId_(employeeId));
      if (matches.length !== 1 || normalize_(matches[0][nameIndex]) !== normalize_(expectedName)) throw Error('The employee number or name changed; reload and review the employee.');
    }
    const sheet = moneMapSheet_(book, !!employeeId);
    if (!sheet) return {removed:true};
    const rows = sheet.getDataRange().getValues();
    const matches = [];
    rows.slice(1).forEach((row,index) => { if (String(row[0] || '') === key) matches.push(index+2); });
    if (matches.length > 1) throw Error('This input name has duplicate saved mappings; review MoneNameMappings.');
    if (!employeeId) {
      if (matches.length) sheet.deleteRow(matches[0]);
      return {removed:true};
    }
    const target = matches.length ? matches[0] : sheet.getLastRow() + 1;
    const range = sheet.getRange(target, 1, 1, 4);
    if (range.getFormulas()[0].some(Boolean)) throw Error('The mapping row contains a formula; review it first.');
    sheet.getRange(target, 2).setNumberFormat('@');
    range.setValues([[key, employeeId, expectedName, new Date()]]);
    return {key, employeeId, employeeName:expectedName};
  } finally { lock.releaseLock(); }
}


// Read-only compilation of saved batch snapshots. It does not imply that a deduction was applied.
function getMoneCompilerData(session) {
  const config = authorizedConfig_(session), book = SpreadsheetApp.openById(config.spreadsheet);
  const sheet = book.getSheetByName('MoneBatchData');
  if (!sheet) return {rows:[]};
  const values = sheet.getDataRange().getValues(), headings = (values.shift() || []).map(v => String(v || '').trim());
  const fields = ['number','rank','inputName','requested','moneVl','moneSl','tag','detail','office'];
  const idx = Object.create(null); fields.forEach(field => { idx[field] = headings.indexOf(field); });
  if (fields.some(field => idx[field] < 0)) throw Error('MoneBatchData is missing expected columns. Review the sheet before exporting.');
  const text = value => String(value == null ? '' : value).trim();
  const rows = values.filter(row => fields.some(field => text(row[idx[field]]))).map(row => {
    const get = field => text(row[idx[field]]), detail = get('detail');
    const source = detail.match(/(?:^|\s·\s)Source file:\s*([\s\S]*?)(?=\s·\s(?:Camp:|Name parts:|Source remarks:)|$)/i);
    const remarks = detail.match(/(?:^|\s·\s)Source remarks:\s*([\s\S]*)$/i);
    const campMatch = detail.match(/(?:^|\s·\s)Camp:\s*(MAXIMUM|MED|MIN|RDC)(?=\s·\s|$)/i);
    const duplicate = /POSSIBLE_DUPLICATE_PREVIOUS_BATCH/i.test(detail) ? 'Possible duplicate; review' : '';
    const vl = get('moneVl'), sl = get('moneSl');
    const approved = vl !== '' && sl !== '' && Number.isFinite(Number(vl)) && Number.isFinite(Number(sl)) && Number(vl) >= 0 && Number(sl) >= 0 ? Math.round((Number(vl) + Number(sl)) * 1000) / 1000 : '';
    return {number:get('number'), rank:get('rank'), inputName:get('inputName'), requested:get('requested'), approved,
      remarks:remarks ? remarks[1].trim() : '', controlNumber:source ? source[1].trim() : '', camp:campMatch ? campMatch[1].toUpperCase() : '', office:get('office'),
      review:[get('tag'), duplicate].filter(Boolean).join(' · ')};
  });
  const locations = moneLocations_(book, false);
  if (locations) {
    const normalize = value => String(value || '').normalize('NFKC').replace(/\s+/g,' ').trim().toLowerCase();
    const saved = locations.getDataRange().getValues().slice(1), lookup = new Map(), counts = new Map();
    saved.forEach(row => { const key = normalize(row[0]); if (key) counts.set(key,(counts.get(key)||0)+1); });
    saved.forEach(row => { const key = normalize(row[0]); if (key && counts.get(key) === 1) lookup.set(key,row); });
    rows.forEach(row => {
      const key = normalize(row.controlNumber);
      if (!key) return;
      if (counts.get(key) > 1) { row.review += ' · Duplicate source location; review'; return; }
      const item = lookup.get(key);
      if (!item) return;
      const camp = String(item[1] || '').trim().toUpperCase(), office = String(item[2] || '').trim();
      if (!row.camp && ['MAXIMUM','MED','MIN','RDC'].includes(camp)) row.camp = camp;
      else if (!row.camp && camp) row.review += ' · Invalid source camp; review';
      if (!row.office && office) row.office = office;
    });
  }
  return {rows};
}


// Cache only; scans run explicitly and never overwrite an existing mapping.
const MONE_SOURCE_LOCATIONS = 'MoneSourceLocations';
const MONE_SOURCE_HEADERS = ['source_filename','camp','office'];
function moneLocations_(book, create) {
  let sheet = book.getSheetByName(MONE_SOURCE_LOCATIONS);
  if (!sheet && create) { sheet = book.insertSheet(MONE_SOURCE_LOCATIONS); sheet.getRange(1,1,1,3).setValues([MONE_SOURCE_HEADERS]); }
  if (!sheet) return null;
  const head = sheet.getRange(1,1,1,3).getValues()[0];
  if (MONE_SOURCE_HEADERS.some((name,i) => String(head[i] || '').trim() !== name)) throw Error('MoneSourceLocations must have source_filename, camp, office headers.');
  return sheet;
}
function sourceLocationFromLines_(lines) {
  const camps = new Set(), offices = new Set();
  lines.forEach(raw => {
    const line = String(raw || '').replace(/\s+/g,' ').trim();
    if (!line || line.length > 200) return;
    const upper = line.toUpperCase();
    if (/\b(?:MAX(?:IMUM)?\s*[- ]?SEC(?:URITY)?\s*[- ]?CAMP|MAXIMUM\s+SECURITY|MAXSECAMP)\b/.test(upper)) camps.add('MAXIMUM');
    if (/\b(?:MED(?:IUM)?\s*[- ]?SEC(?:URITY)?\s*[- ]?CAMP|MEDIUM\s+SECURITY|MEDSECAMP)\b/.test(upper)) camps.add('MED');
    if (/\b(?:MIN(?:IMUM)?\s*[- ]?SEC(?:URITY)?\s*[- ]?CAMP|MINIMUM\s+SECURITY|MINSECAMP)\b/.test(upper)) camps.add('MIN');
    if (/\bRDC\b|\bRECEPTION\s+(?:AND|&)\s+DIAGNOSTIC\s+CENT(?:ER|RE)\b/.test(upper)) camps.add('RDC');
    let office = '';
    const labelled = line.match(/^(?:OFFICE|UNIT|SECTION|DIVISION)\s*:\s*(.+)$/i);
    if (labelled) office = labelled[1].trim();
    else if (/\b(?:OFFICE|UNIT|SECTION|DIVISION)\b/i.test(line)) office = line.split(/,\s*(?:NBP\b|MAX(?:IMUM)?\b|MED(?:IUM)?\b|MIN(?:IMUM)?\b|RDC\b)/i)[0].trim();
    office = office.replace(/\s*[-–,]\s*(?:NBP\s*)?(?:MAXSECAMP|MEDSECAMP|MINSECAMP|RDC)\s*$/i,'').trim();
    if (office && office.length <= 150 && !/^(?:OFFICE|UNIT|SECTION|DIVISION|BUREAU OF CORRECTIONS)$/i.test(office) && !/^[=+\-@]/.test(office)) offices.add(office);
  });
  return {camp:camps.size === 1 ? [...camps][0] : '', office:offices.size === 1 ? [...offices][0] : '', ambiguous:camps.size > 1 || offices.size > 1};
}
function sourceHeadingLines_(file) {
  if (file.getSize() <= 0 || file.getSize() > 3*1024*1024) throw Error('Source file size is unsupported.');
  const entries = Utilities.unzip(Utilities.newBlob(file.getBlob().getBytes(), 'application/zip', file.getName()));
  if (entries.length > 1000) throw Error('Too many file parts.');
  const parts = Object.create(null); let size = 0;
  entries.forEach(part => { size += part.getBytes().length; parts[part.getName()] = part; });
  if (size > 20*1024*1024) throw Error('Source file is too large.');
  const children = (node,name) => node.getChildren().filter(child => child.getName() === name);
  const descendants = (node,name) => { let found=[]; function visit(n){ n.getChildren().forEach(child => {if(child.getName()===name) found.push(child); visit(child);}); } visit(node); return found; };
  const xml = part => XmlService.parse(part.getDataAsString('UTF-8')).getRootElement();
  const content = node => descendants(node,'t').map(n=>n.getValue()).join('');
  const lines = [];
  if (/\.docx$/i.test(file.getName())) {
    if (!parts['word/document.xml']) throw Error('Document body unavailable.');
    const body = descendants(xml(parts['word/document.xml']),'body')[0];
    if (!body) throw Error('Document body unavailable.');
    for (const node of body.getChildren()) {
      if (node.getName() === 'tbl') break;
      if (node.getName() === 'p') { const line = content(node).trim(); if (line) lines.push(line); }
      if (lines.length >= 25) break;
    }
  } else if (/\.xlsx$/i.test(file.getName())) {
    let shared = [];
    if (parts['xl/sharedStrings.xml']) shared = descendants(xml(parts['xl/sharedStrings.xml']),'si').map(content);
    const names = Object.keys(parts).filter(name => /^xl\/worksheets\/sheet\d+\.xml$/.test(name)).sort();
    for (const name of names.slice(0,5)) {
      const rows = descendants(xml(parts[name]),'row').slice(0,8);
      rows.forEach(row => children(row,'c').forEach(cell => {
        const attrs = cell.getAttributes(), a = attrs.find(x => x.getName()==='t');
        const kind = a ? a.getValue() : '', value = children(cell,'v')[0];
        const text = kind === 's' && value ? shared[Number(value.getValue())] : kind === 'inlineStr' ? content(cell) : '';
        if (text && String(text).trim()) lines.push(String(text).trim());
      }));
    }
  } else throw Error('Only DOCX and XLSX source files are supported.');
  return lines;
}
function scanMoneSourceLocations(session) {
  const config = authorizedConfig_(session), context = moneDriveContext_(session), book = SpreadsheetApp.openById(config.spreadsheet);
  const report = book.getSheetByName('MoneBatchData');
  if (!report) throw Error('No saved batches were found.');
  const values = report.getDataRange().getValues(), header = (values.shift() || []).map(v => String(v || '').trim()), detailCol = header.indexOf('detail');
  if (detailCol < 0) throw Error('MoneBatchData detail column is missing.');
  const normalize = v => String(v || '').normalize('NFKC').replace(/\s+/g,' ').trim().toLowerCase();
  const names = new Map();
  values.forEach(row => { const m = String(row[detailCol] || '').match(/(?:^|\s·\s)Source file:\s*([\s\S]*?)(?=\s·\s(?:Camp:|Name parts:|Source remarks:)|$)/i); if (m && m[1].trim()) names.set(normalize(m[1]),m[1].trim()); });
  const files = new Map();
  moneDriveFiles_(context.folder).forEach(file => { const key = normalize(file.getName()); if (names.has(key)) { if (!files.has(key)) files.set(key,[]); files.get(key).push(file); } });
  const prepared = [], issues = [];
  // Read only filenames not already cached; the second check under lock below
  // protects user edits or concurrent scans while the source files are read.
  let sheet = moneLocations_(book, false), existing = new Set();
  if (sheet) sheet.getDataRange().getValues().slice(1).forEach(row=>existing.add(normalize(row[0])));
  for (const [key,name] of names.entries()) {
    if (existing.has(key)) continue;
    const candidates = files.get(key) || [];
    if (candidates.length !== 1) { prepared.push([name,'','']); issues.push(name + ': ' + (candidates.length ? 'duplicate filenames' : 'file not found')); continue; }
    try {
      const file = candidates[0], before = file.getLastUpdated().getTime();
      const parsed = sourceLocationFromLines_(sourceHeadingLines_(file));
      if (file.getLastUpdated().getTime() !== before) throw Error('File changed while reading.');
      prepared.push([name,parsed.camp,parsed.office]);
      if (!parsed.camp || !parsed.office || parsed.ambiguous) issues.push(name + ': verify missing or ambiguous location');
    } catch (error) { prepared.push([name,'','']); issues.push(name + ': could not read heading'); }
  }
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    sheet = moneLocations_(book, true);
    const current = sheet.getDataRange().getValues().slice(1);
    existing = new Set(current.map(row=>normalize(row[0])));
    const append = prepared.filter(row=>!existing.has(normalize(row[0])));
    if (append.length) {
      const last = sheet.getLastRow(), needed = last + append.length - sheet.getMaxRows();
      if (needed > 0) sheet.insertRowsAfter(sheet.getMaxRows(),needed);
      // Prevent file names and office labels from being interpreted as formulas.
      const safe = value => /^\s*[=+\-@]/.test(value) ? "'" + value : value;
      sheet.getRange(last+1,1,append.length,3).setValues(append.map(row=>row.map(safe)));
    }
    return {added:append.length, existing:existing.size, issues:issues.slice(0,30), issueCount:issues.length, sheetUrl:book.getUrl() + '#gid=' + sheet.getSheetId()};
  } finally { lock.releaseLock(); }
}
