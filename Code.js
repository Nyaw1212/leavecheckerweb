/** Leave Checker Web: read-only Apps Script lookup. Configuration is in Script Properties. */
const SHEET_NAME = 'LeaveBalances';
const SESSION_SECONDS = 1800;

function doGet(e) {
  try {
    const config = config_();
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

function page_(session, message, url) {
  const page = HtmlService.createTemplateFromFile('Index');
  page.initialSession = session || '';
  page.initialMessage = message || '';
  page.signInUrl = url || '';
  page.appUrl = ScriptApp.getService().getUrl() || '';
  return page.evaluate().setTitle('Leave Balance Lookup').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function config_() {
  const p = PropertiesService.getScriptProperties();
  const config = {
    spreadsheet: p.getProperty('SPREADSHEET_ID'),
    client: p.getProperty('GOOGLE_CLIENT_ID'),
    secret: p.getProperty('GOOGLE_CLIENT_SECRET'),
    emails: (p.getProperty('ALLOWED_EMAILS') || '').split(',').map(x => x.trim().toLowerCase()).filter(Boolean),
    redirect: ScriptApp.getService().getUrl()
  };
  if (!config.spreadsheet || !config.client || !config.secret || !config.emails.length || !config.redirect) throw Error('Missing configuration');
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

/** Only an authorized, unexpired session may retrieve a matching employee's snapshot. */
function lookupBalance(session, badge, name) {
  if (!/^[0-9a-f]{64}$/i.test(String(session || ''))) throw Error('Please sign in again.');
  const email = CacheService.getScriptCache().get('session:' + session);
  const config = config_();
  if (!email || !config.emails.includes(email)) throw Error('Please sign in again.');
  badge = String(badge || '').trim();
  name = normalize_(name);
  if (!badge || !name || badge.length > 100 || name.length > 200) throw Error('Enter a badge number and full name.');
  const sheet = SpreadsheetApp.openById(config.spreadsheet).getSheetByName(SHEET_NAME);
  if (!sheet) throw Error('The balance sheet is unavailable.');
  const data = sheet.getDataRange().getValues();
  const expected = ['employee_id', 'name', 'as_of', 'vl', 'sl', 'fl', 'wl', 'spl', 'updated_at'];
  const headings = (data.shift() || []).map(x => String(x).trim().toLowerCase());
  const indexes = expected.map(x => headings.indexOf(x));
  if (indexes.some(i => i < 0)) throw Error('The balance sheet is not configured correctly.');
  const idIndex = indexes[0], nameIndex = indexes[1];
  const matches = data.filter(r => String(r[idIndex]).trim().toLowerCase() === badge.toLowerCase());
  if (matches.length !== 1 || normalize_(matches[0][nameIndex]) !== name) return null;
  const row = matches[0];
  const result = {};
  expected.forEach((key, i) => { const v = row[indexes[i]]; result[key] = v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : (v === '' ? null : String(v)); });
  return result;
}
function normalize_(v) { return String(v || '').normalize('NFKC').trim().replace(/\s+/g, ' ').toLowerCase(); }
function signOut(session) { if (/^[0-9a-f]{64}$/i.test(String(session || ''))) CacheService.getScriptCache().remove('session:' + session); }
