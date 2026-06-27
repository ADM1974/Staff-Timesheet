// Shared backend for the STAFF timesheet app.
//
// Identity: staff sign in with Microsoft (MSAL) in the browser; the browser
// sends their ID token, which we VERIFY here against Microsoft (signature,
// tenant, audience) — so identity can't be faked. We then write to SharePoint
// using the existing APP-ONLY connection (same creds as the contractor app).
//
// This file is also the /staff ("me") endpoint, and exports helpers the other
// functions reuse — mirroring the contractor app's validate.js pattern.
//
// Environment variables (Netlify):
//   TENANT_ID, CLIENT_ID, CLIENT_SECRET   - app-only creds (SharePoint writes)
//   SP_SITE_ID, LIST_ID                   - the SharePoint site + Timesheets list
//   STAFF_CLIENT_ID                       - the Staff sign-in app registration (token audience)

const TOKEN_URL = t => `https://login.microsoftonline.com/${t}/oauth2/v2.0/token`;

// ---- app-only token for SharePoint writes (client credentials) ----
async function getAppToken() {
  const body = new URLSearchParams({
    client_id: process.env.CLIENT_ID,
    client_secret: process.env.CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  });
  const r = await fetch(TOKEN_URL(process.env.TENANT_ID), { method: 'POST', body });
  if (!r.ok) throw new Error('token ' + r.status + ' ' + (await r.text()));
  return (await r.json()).access_token;
}

// ---- verify a staff member's Microsoft ID token, return their identity ----
let _jwks;
async function validateStaffToken(authHeader) {
  const m = /^Bearer (.+)$/.exec(authHeader || '');
  if (!m) return null;
  const tenant = process.env.TENANT_ID;
  const { jwtVerify, createRemoteJWKSet } = await import('jose');
  if (!_jwks) _jwks = createRemoteJWKSet(new URL(`https://login.microsoftonline.com/${tenant}/discovery/v2.0/keys`));
  try {
    const { payload } = await jwtVerify(m[1], _jwks, {
      issuer: `https://login.microsoftonline.com/${tenant}/v2.0`,
      audience: process.env.STAFF_CLIENT_ID,
    });
    const id = String(payload.oid || payload.sub || '').trim();
    if (!id) return null;
    return {
      id,                                                   // stable per-user key
      name: String(payload.name || '').trim() || 'Staff',
      email: String(payload.preferred_username || payload.email || '').trim(),
    };
  } catch (e) { return null; }
}

// ===== reused SharePoint helpers (same logic as the contractor app) =====
const WEEKDAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

async function getActiveSites(token) {
  const base = `https://graph.microsoft.com/v1.0/sites/${process.env.SP_SITE_ID}`;
  let r = await fetch(`${base}/lists?$select=id,displayName`, { headers: { Authorization: 'Bearer ' + token } });
  if (!r.ok) throw new Error('graph lists ' + r.status + ' ' + (await r.text()));
  const list = ((await r.json()).value || []).find(l => String(l.displayName || '').toLowerCase() === 'sites');
  if (!list) throw new Error('Sites list not found');
  r = await fetch(`${base}/lists/${list.id}/items?expand=fields&$top=999`, { headers: { Authorization: 'Bearer ' + token } });
  if (!r.ok) throw new Error('graph items ' + r.status + ' ' + (await r.text()));
  return ((await r.json()).value || [])
    .filter(it => (it.fields || {}).Active !== false)
    .map(it => String((it.fields || {}).Title || '').trim())
    .filter(Boolean).sort((a, b) => a.localeCompare(b));
}

// Sites list → { byName: {lowerName:{id,name,closeDayIndex}}, byId: {id:name} }.
async function getSitesMap(token) {
  const base = `https://graph.microsoft.com/v1.0/sites/${process.env.SP_SITE_ID}`;
  let r = await fetch(`${base}/lists?$select=id,displayName`, { headers: { Authorization: 'Bearer ' + token } });
  if (!r.ok) throw new Error('graph lists ' + r.status + ' ' + (await r.text()));
  const list = ((await r.json()).value || []).find(l => String(l.displayName || '').toLowerCase() === 'sites');
  if (!list) throw new Error('Sites list not found');
  r = await fetch(`${base}/lists/${list.id}/items?expand=fields&$top=999`, { headers: { Authorization: 'Bearer ' + token } });
  if (!r.ok) throw new Error('graph items ' + r.status + ' ' + (await r.text()));
  const byName = {}, byId = {};
  for (const it of ((await r.json()).value || [])) {
    const f = it.fields || {};
    const name = String(f.Title || '').trim();
    if (!name) continue;
    const idx = WEEKDAYS.findIndex(d => d.toLowerCase() === String(f.CloseDay || '').trim().toLowerCase());
    byName[name.toLowerCase()] = { id: it.id, name, closeDayIndex: idx < 0 ? 0 : idx };
    byId[String(it.id)] = name;
  }
  return { byName, byId };
}

// The staff member's HOME site name, from the Workers list matched by Email.
// HomeSite may be a Lookup (resolved via the sites map) or plain text/choice.
async function getHomeSite(token, email, sitesMap) {
  const want = String(email || '').trim().toLowerCase();
  if (!want) return '';
  const base = `https://graph.microsoft.com/v1.0/sites/${process.env.SP_SITE_ID}`;
  let r = await fetch(`${base}/lists?$select=id,displayName`, { headers: { Authorization: 'Bearer ' + token } });
  if (!r.ok) return '';
  const list = ((await r.json()).value || []).find(l => String(l.displayName || '').toLowerCase() === 'workers');
  if (!list) return '';
  r = await fetch(`${base}/lists/${list.id}/items?expand=fields&$top=999`, { headers: { Authorization: 'Bearer ' + token } });
  if (!r.ok) return '';
  for (const it of ((await r.json()).value || [])) {
    const f = it.fields || {};
    if (String(f.Email || '').trim().toLowerCase() !== want) continue;
    let hs = String(f.HomeSite || '').trim();
    if (!hs && f.HomeSiteLookupId != null) hs = sitesMap.byId[String(f.HomeSiteLookupId)] || '';
    return hs;
  }
  return '';
}

// All active allowances grouped by site name → [allowance titles].
async function getAllowancesBySite(token, sitesMap) {
  const base = `https://graph.microsoft.com/v1.0/sites/${process.env.SP_SITE_ID}`;
  let r = await fetch(`${base}/lists?$select=id,displayName`, { headers: { Authorization: 'Bearer ' + token } });
  if (!r.ok) return {};
  const list = ((await r.json()).value || []).find(l => String(l.displayName || '').toLowerCase() === 'allowances');
  if (!list) return {};
  r = await fetch(`${base}/lists/${list.id}/items?expand=fields&$top=999`, { headers: { Authorization: 'Bearer ' + token } });
  if (!r.ok) return {};
  const out = {};
  for (const it of ((await r.json()).value || [])) {
    const f = it.fields || {};
    if (f.Active === false) continue;
    const siteName = f.SiteLookupId != null ? sitesMap.byId[String(f.SiteLookupId)] : '';
    const title = String(f.Title || '').trim();
    if (!siteName || !title) continue;
    (out[siteName] = out[siteName] || []).push(title);
  }
  for (const k in out) out[k].sort((a, b) => a.localeCompare(b));
  return out;
}

function nzDateInfo() {
  const s = new Intl.DateTimeFormat('en-CA', { timeZone: 'Pacific/Auckland', year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date());
  const [y, m, d] = s.split('-').map(Number);
  return { y, m, d, dow: new Date(Date.UTC(y, m - 1, d)).getUTCDay() };
}
const _pad = n => String(n).padStart(2, '0');
const _iso = ms => { const dt = new Date(ms); return `${dt.getUTCFullYear()}-${_pad(dt.getUTCMonth() + 1)}-${_pad(dt.getUTCDate())}`; };

function computeOpenWeek(closeDayIdx) {
  const { y, m, d, dow } = nzDateInfo();
  const today = Date.UTC(y, m - 1, d);
  const end = today + ((closeDayIdx - dow + 7) % 7) * 86400000;
  const days = [];
  for (let i = 6; i >= 0; i--) days.push(_iso(end - i * 86400000));
  return { weekStart: days[0], weekEnd: days[6], days };
}

// Build the open-week context from a close-day index (computed from the home site).
function weekContext(closeDayIdx) {
  return { ...computeOpenWeek(closeDayIdx), closeDayName: WEEKDAYS[closeDayIdx] };
}

// A user's Timesheets rows within [weekStart,weekEnd], by the ContractorId key
// (for staff this holds their 365 object id).
async function getUserEntries(token, userKey, weekStart, weekEnd) {
  const url = `https://graph.microsoft.com/v1.0/sites/${process.env.SP_SITE_ID}/lists/${process.env.LIST_ID}/items?expand=fields&$top=999&$filter=fields/ContractorId eq '${userKey}'`;
  const r = await fetch(url, { headers: { Authorization: 'Bearer ' + token, Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' } });
  if (!r.ok) throw new Error('graph entries ' + r.status + ' ' + (await r.text()));
  return ((await r.json()).value || []).filter(it => {
    const ed = String((it.fields || {}).EntryDate || '').slice(0, 10);
    return ed >= weekStart && ed <= weekEnd;
  });
}

const itemsUrl = () => `https://graph.microsoft.com/v1.0/sites/${process.env.SP_SITE_ID}/lists/${process.env.LIST_ID}/items`;
async function createItem(token, fields) {
  const r = await fetch(itemsUrl(), { method: 'POST', headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' }, body: JSON.stringify({ fields }) });
  if (!r.ok) throw new Error('graph create ' + r.status + ' ' + (await r.text()));
}
async function deleteItem(token, id) {
  const r = await fetch(`${itemsUrl()}/${id}`, { method: 'DELETE', headers: { Authorization: 'Bearer ' + token } });
  if (!r.ok && r.status !== 204) throw new Error('graph delete ' + r.status + ' ' + (await r.text()));
}

// ---- /staff endpoint: confirm sign-in and return the staff member's name ----
exports.handler = async (event) => {
  const headers = { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' };
  const user = await validateStaffToken(event.headers && (event.headers.authorization || event.headers.Authorization));
  if (!user) return { statusCode: 401, headers, body: '{"ok":false}' };
  return { statusCode: 200, headers, body: JSON.stringify({ ok: true, name: user.name, email: user.email }) };
};

exports.getAppToken = getAppToken;
exports.validateStaffToken = validateStaffToken;
exports.getActiveSites = getActiveSites;
exports.getSitesMap = getSitesMap;
exports.getHomeSite = getHomeSite;
exports.getAllowancesBySite = getAllowancesBySite;
exports.weekContext = weekContext;
exports.getUserEntries = getUserEntries;
exports.createItem = createItem;
exports.deleteItem = deleteItem;
