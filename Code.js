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
  const ni = headings.indexOf('name'), ii = headings.indexOf('employee_id');
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
      return {name: String(row[ni]), employeeId: id, selectable: !!key && counts[key] === 1, ...balances};
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
