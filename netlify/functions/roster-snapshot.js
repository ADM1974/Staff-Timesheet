// Teams Shifts roster, FIXED at the start of each pay week.
//
// A site is linked to Shifts by putting its Team's id in the Sites list
// "ShiftsTeamId" column. At the start of that site's pay week (the day after its
// CloseDay) this scheduled function reads the Team's PUBLISHED shifts for the week
// and saves them as one row in the "RosterSnapshots" SharePoint list. From then on
// the app only reads that row — later edits in Shifts don't change the week; staff
// correct anything themselves with "Shift worked" in the app.
//
// If a week has no snapshot yet (first deploy, a missed run) the first read takes
// one on the spot, and that becomes the fixed roster for the week.
//
// Needs the Graph APPLICATION permission Schedule.Read.All (admin-consented) on the
// CLIENT_ID app registration. Optional env SHIFTS_ACT_AS = a Team owner's user id,
// used only if Shifts insists on an MS-APP-ACTS-AS user for the scheduled run.
// Schedule: netlify.toml [functions."roster-snapshot"] (hourly).

const { getAppToken, getSitesMap, weekFor } = require('./staff');

const SNAP_LIST = 'RosterSnapshots';
const DEFAULT_TZ = 'Australia/Brisbane';
const spBase = () => `https://graph.microsoft.com/v1.0/sites/${process.env.SP_SITE_ID}`;

// ---- Shifts (Graph) ----
// App-only Shifts calls may need a user named in MS-APP-ACTS-AS; try with it (when we
// have one) and then without, so either tenant behaviour works.
async function shiftsGet(token, url, actAs) {
  const tries = actAs ? [{ 'MS-APP-ACTS-AS': actAs }, {}] : [{}];
  let r;
  for (const extra of tries) {
    r = await fetch(url, { headers: { Authorization: 'Bearer ' + token, ...extra } });
    if (r.ok || (r.status !== 400 && r.status !== 403)) break;
  }
  if (!r.ok) throw new Error('shifts ' + r.status + ' ' + (await r.text()).slice(0, 300));
  return r.json();
}
const _ymd = (d, tz) => new Intl.DateTimeFormat('en-CA', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit' }).format(d);
const _hm = (d, tz) => new Intl.DateTimeFormat('en-GB', { timeZone: tz, hour: '2-digit', minute: '2-digit', hour12: false }).format(d).replace(/^24/, '00');

async function teamTimeZone(token, teamId, actAs) {
  let tz = DEFAULT_TZ;
  try { const sch = await shiftsGet(token, `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent(teamId)}/schedule`, actAs); if (sch.timeZone) tz = sch.timeZone; } catch (e) { /* keep default */ }
  try { _ymd(new Date(), tz); } catch (e) { tz = DEFAULT_TZ; }   // a zone name Node doesn't know
  return tz;
}

// Every published shift in the Team that STARTS on one of the week's days (in the
// schedule's time zone — so a night shift sits on the day it starts).
async function readTeamWeek(token, site, wk, actAs) {
  const tz = await teamTimeZone(token, site.shiftsTeamId, actAs);
  const want = new Set(wk.days);
  const from = new Date(Date.parse(wk.weekStart + 'T00:00:00Z') - 86400000).toISOString();
  const to = new Date(Date.parse(wk.weekEnd + 'T00:00:00Z') + 2 * 86400000).toISOString();
  const all = `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent(site.shiftsTeamId)}/schedule/shifts`;
  let url = `${all}?$filter=` + encodeURIComponent(`sharedShift/startDateTime ge ${from} and sharedShift/endDateTime le ${to}`);
  const shifts = [];
  for (let page = 0; url && page < 50; page++) {
    let j;
    try { j = await shiftsGet(token, url, actAs); }
    catch (e) {                                      // date filter refused → read unfiltered, filter below
      if (page > 0 || url === all || !/^shifts 400\b/.test(String(e.message))) throw e;
      url = all; j = await shiftsGet(token, url, actAs);
    }
    for (const it of (j.value || [])) {
      const sh = it.sharedShift;                     // published version only (not drafts)
      if (!sh || !it.userId || !sh.startDateTime) continue;
      const st = new Date(sh.startDateTime), en = new Date(sh.endDateTime || sh.startDateTime);
      const d = _ymd(st, tz);
      if (!want.has(d)) continue;
      shifts.push({ u: String(it.userId), d, n: String(sh.displayName || '').trim() || 'Shift', s: _hm(st, tz), e: _hm(en, tz), t: st.getTime() });
    }
    url = j['@odata.nextLink'] || '';
  }
  shifts.sort((a, b) => a.t - b.t);
  return { tz, shifts: shifts.map(({ t, ...x }) => x) };
}

// ---- RosterSnapshots list (SharePoint) ----
let _snapListId = null;
async function snapListId(token, create) {
  if (_snapListId) return _snapListId;
  let r = await fetch(`${spBase()}/lists?$select=id,displayName`, { headers: { Authorization: 'Bearer ' + token } });
  if (!r.ok) throw new Error('graph lists ' + r.status);
  const found = ((await r.json()).value || []).find(l => String(l.displayName || '').toLowerCase() === SNAP_LIST.toLowerCase());
  if (found) return (_snapListId = found.id);
  if (!create) return null;
  // First run: create the list (falls back to the README's manual steps if the app isn't allowed to).
  r = await fetch(`${spBase()}/lists`, {
    method: 'POST', headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      displayName: SNAP_LIST,
      list: { template: 'genericList' },
      columns: [
        { name: 'Site', text: {} },
        { name: 'WeekStart', text: {} },
        { name: 'TakenAt', text: {} },
        { name: 'Data', text: { allowMultipleLines: true, textType: 'plain' } },
      ],
    }),
  });
  if (!r.ok) throw new Error('create ' + SNAP_LIST + ' list ' + r.status + ' ' + (await r.text()).slice(0, 300));
  return (_snapListId = (await r.json()).id);
}

// The saved snapshot for one site + week, or null. If two exist, the earliest wins.
async function loadSnapshot(token, listId, siteName, weekStart) {
  const url = `${spBase()}/lists/${listId}/items?expand=fields&$top=50&$filter=` + encodeURIComponent(`fields/WeekStart eq '${weekStart}'`);
  const r = await fetch(url, { headers: { Authorization: 'Bearer ' + token, Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' } });
  if (!r.ok) throw new Error('snapshot read ' + r.status);
  const rows = ((await r.json()).value || [])
    .map(it => it.fields || {})
    .filter(f => String(f.Site || '').toLowerCase() === siteName.toLowerCase())
    .sort((a, b) => String(a.TakenAt || '').localeCompare(String(b.TakenAt || '')));
  for (const f of rows) { try { return JSON.parse(f.Data || ''); } catch (e) { /* unreadable row: try the next */ } }
  return null;
}

async function takeSnapshot(token, site, wk, actAs) {
  const listId = await snapListId(token, true);
  const existing = await loadSnapshot(token, listId, site.name, wk.weekStart);
  if (existing) return existing;                     // already fixed — never overwrite
  const { tz, shifts } = await readTeamWeek(token, site, wk, actAs || process.env.SHIFTS_ACT_AS || '');
  const snap = { site: site.name, weekStart: wk.weekStart, weekEnd: wk.weekEnd, tz, takenAt: new Date().toISOString(), shifts };
  const r = await fetch(`${spBase()}/lists/${listId}/items`, {
    method: 'POST', headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify({ fields: { Title: `${site.name} ${wk.weekStart}`, Site: site.name, WeekStart: wk.weekStart, TakenAt: snap.takenAt, Data: JSON.stringify(snap) } }),
  });
  if (!r.ok) throw new Error('snapshot write ' + r.status + ' ' + (await r.text()).slice(0, 300));
  return snap;
}

// The user's FIXED rostered shifts on the given ISO days, from every linked site.
// Never throws: a site whose roster can't be read is listed in `errors`.
// → { byDate: { 'YYYY-MM-DD': [{ site, name, start, end }] }, errors: [site] }
async function getRoster(token, userId, sitesMap, days) {
  const byDate = {}, errors = [];
  const sites = Object.values(sitesMap.byName).filter(s => s.shiftsTeamId);
  if (!userId || !days.length || !sites.length) return { byDate, errors };
  const want = new Set(days);
  for (const site of sites) {
    try {
      const weeks = {};                              // this site's pay weeks covering the days
      for (const d of days) { const w = weekFor(d, site.closeDayIndex); weeks[w.weekStart] = w; }
      for (const wk of Object.values(weeks)) {
        const listId = await snapListId(token, false);
        const snap = (listId && await loadSnapshot(token, listId, site.name, wk.weekStart)) || await takeSnapshot(token, site, wk, userId);
        for (const x of (snap.shifts || [])) {
          if (x.u !== userId || !want.has(x.d)) continue;
          (byDate[x.d] = byDate[x.d] || []).push({ site: site.name, name: x.n, start: x.s, end: x.e });
        }
      }
    } catch (e) {
      console.error('roster ' + site.name + ': ' + (e.message || e));
      errors.push(site.name);
    }
  }
  return { byDate, errors };
}

// One-line summary saved on timesheet rows, e.g. "Night 18:00–06:00 (BNE)".
function rosterLabel(list) {
  return (list || []).map(x => `${x.name} ${x.start}–${x.end} (${x.site})`).join('; ');
}

// ---- scheduled run (hourly): fix each linked site's roster once its week has begun ----
exports.handler = async () => {
  try {
    const token = await getAppToken();
    const sitesMap = await getSitesMap(token);
    const out = [];
    for (const site of Object.values(sitesMap.byName).filter(s => s.shiftsTeamId)) {
      try {
        const tz = await teamTimeZone(token, site.shiftsTeamId, process.env.SHIFTS_ACT_AS || '');
        const wk = weekFor(_ymd(new Date(), tz), site.closeDayIndex);   // the week that's running now, in the site's zone
        const snap = await takeSnapshot(token, site, wk, '');
        out.push(`${site.name} ${wk.weekStart}: ${snap.shifts.length} shifts (taken ${snap.takenAt})`);
      } catch (e) { out.push(`${site.name}: ERROR ${e.message || e}`); }
    }
    console.log('roster-snapshot:', out.join(' | ') || 'no linked sites');
    return { statusCode: 200, body: 'ok' };
  } catch (err) {
    console.error('roster-snapshot error', err);
    return { statusCode: 500, body: 'error' };
  }
};

exports.getRoster = getRoster;
exports.rosterLabel = rosterLabel;
