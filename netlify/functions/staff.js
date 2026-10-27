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
//
// Teams Shifts (optional): a site whose Sites row has a ShiftsTeamId (the Team's id)
// gets its published roster read from Shifts. Needs the Graph APPLICATION permission
// Schedule.Read.All (admin-consented) on the CLIENT_ID app registration.

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

// Sites list → { byName: {lowerName:{id,name,closeDayIndex,shiftsTeamId}}, byId: {id:name} }.
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
    const shiftsTeamId = String(f.ShiftsTeamId || '').trim();   // optional: Teams Shifts roster for this site
    byName[name.toLowerCase()] = { id: it.id, name, closeDayIndex: idx < 0 ? 0 : idx, shiftsTeamId };
    byId[String(it.id)] = name;
  }
  return { byName, byId };
}

// Generate the company email from a worker's name: first-initial.surname@ael.co
// (so the Workers list only needs a name + a site — no Email column required).
function emailFromName(name) {
  const parts = String(name || '').trim().split(/\s+/).filter(Boolean);
  if (parts.length < 2) return '';
  const first = parts[0].replace(/[^a-z]/gi, '').charAt(0).toLowerCase();
  const last = parts[parts.length - 1].replace(/[^a-z0-9-]/gi, '').toLowerCase();
  return (first && last) ? `${first}.${last}@ael.co` : '';
}

// The staff member's HOME site name, from the Workers list. A row matches the
// signed-in email by EITHER an explicit "Email" column OR the email generated from
// the worker's name (Title). Site may be plain text/choice or a Lookup.
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
    const rowEmail = String(f.Email || '').trim().toLowerCase();
    const genEmail = emailFromName(f.Title);                    // derive from the worker's name
    if (rowEmail !== want && genEmail !== want) continue;       // explicit Email OR generated
    let hs = String(f.HomeSite || f.Site || '').trim();                 // text/choice (HomeSite or Site)
    if (!hs) {
      const lid = f.HomeSiteLookupId != null ? f.HomeSiteLookupId : f.SiteLookupId;  // lookup column
      if (lid != null) hs = sitesMap.byId[String(lid)] || '';
    }
    return hs;
  }
  return '';
}

// Optional "Days" column on Allowances (multi-choice Monday..Sunday) → weekday
// indexes (0 = Sunday). Blank → null = every day. Accepts an array or "Sat, Sun".
const DAY_KEYS = ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'];
function parseDays(v) {
  const raw = Array.isArray(v) ? v : String(v || '').split(/[,;]/);
  const idx = [...new Set(raw.map(x => DAY_KEYS.indexOf(String(x || '').trim().toLowerCase().slice(0, 3))).filter(i => i >= 0))];
  return idx.length ? idx.sort() : null;
}
// Is daily allowance `title` allowed on ISO date `date`, given a { title:[weekday idx] } map?
function allowedOnDate(daysMap, title, date) {
  const days = daysMap && daysMap[title];
  return !days || days.includes(new Date(String(date).slice(0, 10) + 'T00:00:00Z').getUTCDay());
}

// All active allowances grouped by site name → { daily:[], weekly:[], days:{}, shiftMap:{}, shiftNames:[] }.
// days = { title: [weekday idx] } for daily allowances limited to certain days.
// shiftMap = { lower-case Shifts label: [daily allowance titles] } — a shift pre-ticks
// these. Labels come from the optional "ShiftName" column ("Night" or "Night, Late");
// the allowance's own title always counts too. shiftNames = those labels as typed.
async function getAllowancesBySite(token, sitesMap) {
  const base = `https://graph.microsoft.com/v1.0/sites/${process.env.SP_SITE_ID}`;
  let r = await fetch(`${base}/lists?$select=id,displayName`, { headers: { Authorization: 'Bearer ' + token } });
  if (!r.ok) return {};
  const list = ((await r.json()).value || []).find(l => String(l.displayName || '').toLowerCase() === 'allowances');
  if (!list) return {};
  r = await fetch(`${base}/lists/${list.id}/items?expand=fields&$top=999`, { headers: { Authorization: 'Bearer ' + token } });
  if (!r.ok) return {};
  const out = {};
  const anyDay = {};                            // site → daily titles with a row open to every day
  for (const it of ((await r.json()).value || [])) {
    const f = it.fields || {};
    if (f.Active === false) continue;
    const siteName = f.SiteLookupId != null ? sitesMap.byId[String(f.SiteLookupId)] : '';
    const title = String(f.Title || '').trim();
    if (!siteName || !title) continue;
    const bucket = (out[siteName] = out[siteName] || { daily: [], weekly: [], days: {}, shiftMap: {}, shiftNames: [] });
    if (String(f.Frequency || 'Daily').trim().toLowerCase() === 'weekly') { if (!bucket.weekly.includes(title)) bucket.weekly.push(title); continue; }
    if (!bucket.daily.includes(title)) bucket.daily.push(title);
    const labels = String(f.ShiftName || '').split(/[,;]/).map(x => x.trim()).filter(Boolean);
    for (const lab of labels) if (!bucket.shiftNames.includes(lab)) bucket.shiftNames.push(lab);
    for (const lab of [...labels, title]) {
      const k = lab.toLowerCase();
      if (!(bucket.shiftMap[k] = bucket.shiftMap[k] || []).includes(title)) bucket.shiftMap[k].push(title);
    }
    const days = parseDays(f.Days);
    if (!days) (anyDay[siteName] = anyDay[siteName] || new Set()).add(title);
    else bucket.days[title] = [...new Set([...(bucket.days[title] || []), ...days])].sort();
  }
  for (const k in out) {
    for (const t of (anyDay[k] || [])) delete out[k].days[t];
    out[k].daily.sort((a, b) => a.localeCompare(b)); out[k].weekly.sort((a, b) => a.localeCompare(b));
    out[k].shiftNames.sort((a, b) => a.localeCompare(b));
  }
  return out;
}

// Active work orders for a site, from the "WorkOrders" list (Site = lookup or text).
// Returns [] if there's no list yet, so the typeahead simply has no suggestions.
async function getWorkOrders(token, siteId, siteName) {
  const base = `https://graph.microsoft.com/v1.0/sites/${process.env.SP_SITE_ID}`;
  let r = await fetch(`${base}/lists?$select=id,displayName`, { headers: { Authorization: 'Bearer ' + token } });
  if (!r.ok) return [];
  const list = ((await r.json()).value || []).find(l => String(l.displayName || '').toLowerCase().replace(/\s+/g, '') === 'workorders');
  if (!list) return [];
  r = await fetch(`${base}/lists/${list.id}/items?expand=fields&$top=999`, { headers: { Authorization: 'Bearer ' + token } });
  if (!r.ok) return [];
  const wantId = siteId != null ? String(siteId) : '';
  const wantName = String(siteName || '').trim().toLowerCase();
  return ((await r.json()).value || [])
    .filter(it => {
      const f = it.fields || {};
      if (f.Active === false) return false;
      const byId = wantId && String(f.SiteLookupId != null ? f.SiteLookupId : '') === wantId;
      const byName = wantName && String(f.Site || '').trim().toLowerCase() === wantName;
      return byId || byName;
    })
    .map(it => String((it.fields || {}).Title || '').trim())
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b));
}

function nzDateInfo() {
  const s = new Intl.DateTimeFormat('en-CA', { timeZone: 'Pacific/Auckland', year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date());
  const [y, m, d] = s.split('-').map(Number);
  return { y, m, d, dow: new Date(Date.UTC(y, m - 1, d)).getUTCDay() };
}
const _pad = n => String(n).padStart(2, '0');
const _iso = ms => { const dt = new Date(ms); return `${dt.getUTCFullYear()}-${_pad(dt.getUTCMonth() + 1)}-${_pad(dt.getUTCDate())}`; };

function computeOpenWeek(closeDayIdx) {
  // Stay open until 11:00 NZ the day AFTER the close day (the report-send time), not
  // midnight — compute the week from a "now" shifted back 11h (= the send hour), so
  // people can finalise the just-closed week right up to the cut-off.
  const shifted = new Date(Date.now() - 11 * 3600 * 1000);
  const s = new Intl.DateTimeFormat('en-CA', { timeZone: 'Pacific/Auckland', year: 'numeric', month: '2-digit', day: '2-digit' }).format(shifted);
  const [y, m, d] = s.split('-').map(Number);
  const dow = new Date(Date.UTC(y, m - 1, d)).getUTCDay();
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

// The pay week (7 ISO days, ending on closeDayIdx) that CONTAINS dateStr.
function weekFor(dateStr, closeDayIdx) {
  const [y, m, d] = String(dateStr).slice(0, 10).split('-').map(Number);
  const base = Date.UTC(y, m - 1, d);
  const dow = new Date(base).getUTCDay();
  const end = base + ((closeDayIdx - dow + 7) % 7) * 86400000;
  const days = [];
  for (let i = 6; i >= 0; i--) days.push(_iso(end - i * 86400000));
  return { weekStart: days[0], weekEnd: days[6], days };
}

// If the staff member has a REJECTED week, return it (so they can fix + re-submit
// even after that week closed). Most recent rejected week wins. null if none.
async function getRejectedWeek(token, userKey, closeDayIdx) {
  if (!userKey) return null;
  const url = `https://graph.microsoft.com/v1.0/sites/${process.env.SP_SITE_ID}/lists/${process.env.LIST_ID}/items?expand=fields&$top=999&$filter=fields/ContractorId eq '${userKey}'`;
  const r = await fetch(url, { headers: { Authorization: 'Bearer ' + token, Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' } });
  if (!r.ok) return null;
  const rejected = ((await r.json()).value || []).filter(it => String((it.fields || {}).Status || '') === 'Rejected');
  if (!rejected.length) return null;
  let latest = '';
  for (const it of rejected) { const dd = String((it.fields || {}).EntryDate || '').slice(0, 10); if (dd > latest) latest = dd; }
  if (!latest) return null;
  const wk = weekFor(latest, closeDayIdx);
  let reason = '';
  for (const it of rejected) {
    const f = it.fields || {};
    const dd = String(f.EntryDate || '').slice(0, 10);
    if (dd >= wk.weekStart && dd <= wk.weekEnd && f.RejectionReason) { reason = String(f.RejectionReason).trim(); break; }
  }
  return { ...wk, reason };
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

// Create a Timesheets row, tolerating optional columns that don't exist yet: if
// SharePoint rejects the row (400), retry without the shift columns, then also
// without Comment. `state` (shared across one save) remembers which columns were
// missing so later rows skip them straight away.
const OPTIONAL_GROUPS = [['ShiftWorked', 'RosteredShift'], ['Comment']];
async function createItemResilient(token, fields, state) {
  const missing = (state.missing = state.missing || new Set());
  const f = { ...fields };
  for (const k of missing) delete f[k];
  for (;;) {
    try { await createItem(token, f); return; }
    catch (e) {
      if (!/^graph create 400\b/.test(String(e.message))) throw e;   // only a rejected row is worth retrying
      const group = OPTIONAL_GROUPS.find(g => g.some(k => k in f));
      if (!group) throw e;
      for (const k of group) { delete f[k]; missing.add(k); }
    }
  }
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
exports.allowedOnDate = allowedOnDate;
exports.getWorkOrders = getWorkOrders;
exports.weekContext = weekContext;
exports.getUserEntries = getUserEntries;
exports.getRejectedWeek = getRejectedWeek;
exports.createItem = createItem;
exports.createItemResilient = createItemResilient;
exports.weekFor = weekFor;
exports.deleteItem = deleteItem;
