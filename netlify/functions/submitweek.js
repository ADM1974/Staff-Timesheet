// Saves the signed-in staff member's whole week — REPLACES their rows for the
// open week (delete + rewrite). Week boundary is the home site's; hours are per
// the day's chosen site; allowances are validated against that day's site.
const { getAppToken, validateStaffToken, getActiveSites, getSitesMap, getHomeSite, getAllowancesBySite, allowedOnDate, weekContext, getUserEntries, getRejectedWeek, createItemResilient, deleteItem } = require('./staff');
const { getRoster, rosterLabel } = require('./roster-snapshot');

exports.handler = async (event) => {
  const headers = { 'Content-Type': 'application/json' };
  const bad = msg => ({ statusCode: 400, headers, body: JSON.stringify({ ok: false, error: msg }) });

  const user = await validateStaffToken(event.headers && (event.headers.authorization || event.headers.Authorization));
  if (!user) return { statusCode: 401, headers, body: '{"ok":false}' };

  let data;
  try { data = JSON.parse(event.body || '{}'); } catch { return bad('bad json'); }

  try {
    const token = await getAppToken();
    const sitesMap = await getSitesMap(token);
    const homeSite = await getHomeSite(token, user.email, sitesMap);
    const closeDayIdx = (sitesMap.byName[String(homeSite || '').toLowerCase()] || {}).closeDayIndex || 0;
    const wk = weekContext(closeDayIdx);
    const rejected = await getRejectedWeek(token, user.id, closeDayIdx);
    const submittedDates = (Array.isArray(data.days) ? data.days : []).map(d => String(d.date || '').slice(0, 10));
    const target = (rejected && submittedDates.some(d => rejected.days.includes(d))) ? rejected : wk;  // re-submitting a sent-back week, or the open week
    const open = new Set(target.days);
    const activeSites = new Set(await getActiveSites(token));
    const allowancesBySite = await getAllowancesBySite(token, sitesMap);
    const weeklySet = new Set(((allowancesBySite[homeSite] || {}).weekly) || []);
    const roster = await getRoster(token, user.id, sitesMap, target.days);   // the week's fixed roster (not trusted from the phone)

    const clean = [];
    for (const day of (Array.isArray(data.days) ? data.days : [])) {
      const date = String(day.date || '').slice(0, 10);
      if (!open.has(date)) continue;
      const site = String(day.site || '').trim();
      if (!site) continue;
      if (!activeSites.has(site)) return bad('Unknown or inactive site: ' + site);
      const lines = (Array.isArray(day.lines) ? day.lines : [])
        .map(l => ({ wo: String(l.wo || '').trim().slice(0, 200), hr: parseFloat(l.hr), comment: String(l.comment || '').trim().slice(0, 500) }))
        .filter(l => l.wo && l.hr > 0 && l.hr <= 24);
      const siteAllow = allowancesBySite[site] || {};
      const dailyOpts = new Set(siteAllow.daily || []);
      const allowances = (Array.isArray(day.allowances) ? day.allowances : [])
        .map(a => String(a || '').trim())
        .filter(a => dailyOpts.has(a) && allowedOnDate(siteAllow.days, a, date));   // daily allowances for the day's site, on their allowed days
      if (!lines.length && !allowances.length) continue;
      const shiftWorked = String(day.shiftWorked || '').trim().slice(0, 100);
      clean.push({ date, site, lines, allowances, shiftWorked, rosteredShift: rosterLabel(roster.byDate[date]).slice(0, 255) });
    }
    const weekAllowances = (Array.isArray(data.weekAllowances) ? data.weekAllowances : [])
      .map(a => String(a || '').trim())
      .filter(a => weeklySet.has(a));             // weekly allowances = the home site's

    const existing = await getUserEntries(token, user.id, target.weekStart, target.weekEnd);
    for (const it of existing) await deleteItem(token, it.id);

    const batchId = 'S-' + target.weekStart + '-' + Math.random().toString(36).slice(2, 8).toUpperCase();
    let created = 0;
    const cols = {};                                  // which optional columns exist (learned on first write)
    for (const day of clean) {
      const shiftCols = {};                           // stamped on every row of the day
      if (day.shiftWorked) shiftCols.ShiftWorked = day.shiftWorked;
      if (day.rosteredShift) shiftCols.RosteredShift = day.rosteredShift;
      for (const l of day.lines) {
        const fields = {
          Title: user.name, ContractorId: user.id, EntryDate: day.date + 'T00:00:00Z',
          Site: day.site, WorkOrder: l.wo, Hours: l.hr, Status: 'Submitted', BatchID: batchId,
          Notes: 'Staff app · ' + user.email, ...shiftCols,
        };
        if (l.comment) fields.Comment = l.comment;
        await createItemResilient(token, fields, cols);   // Comment / shift columns may not exist yet
        created++;
      }
      for (const a of day.allowances) {
        await createItemResilient(token, {
          Title: user.name, ContractorId: user.id, EntryDate: day.date + 'T00:00:00Z',
          Site: day.site, RowType: 'Allowance', Allowance: a, Hours: 0, Status: 'Submitted', BatchID: batchId,
          Notes: 'Staff app allowance · ' + user.email, ...shiftCols,
        }, cols);
        created++;
      }
    }
    for (const a of weekAllowances) {
      await createItemResilient(token, {
        Title: user.name, ContractorId: user.id, EntryDate: target.weekStart + 'T00:00:00Z',
        Site: homeSite, RowType: 'Allowance', Allowance: a, Hours: 0, Status: 'Submitted', BatchID: batchId,
        Notes: 'Staff app weekly allowance · ' + user.email,
      }, cols);
      created++;
    }
    return { statusCode: 200, headers, body: JSON.stringify({ ok: true, created, weekStart: wk.weekStart, weekEnd: wk.weekEnd }) };
  } catch (err) {
    console.error(err);
    return { statusCode: 502, headers, body: JSON.stringify({ ok: false, error: 'server' }) };
  }
};
