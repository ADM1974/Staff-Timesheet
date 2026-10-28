// Saves the signed-in staff member's whole week — REPLACES their rows for the
// open week: the new rows are written first and only then are the old ones removed,
// so a save that fails part-way is undone and leaves the week as it was. Week boundary is the home site's; hours are per
// the day's chosen site; allowances are validated against that day's site.
const { getAppToken, validateStaffToken, getActiveSites, getSitesMap, getHomeSite, getAllowancesBySite, allowedOnDate, weekContext, getUserEntries, getRejectedWeek, createItemResilient, deleteItem, pool } = require('./staff');
const { getRoster, rosterLabel, allowanceForShift } = require('./roster-snapshot');

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
        .map(l => ({ wo: String(l.wo || '').trim().slice(0, 200), hr: parseFloat(l.hr), comment: String(l.comment || '').trim().slice(0, 500), co: l.co === true }))
        .filter(l => l.wo && l.hr > 0 && l.hr <= 24);
      const siteAllow = allowancesBySite[site] || {};
      const dailyOpts = new Set(siteAllow.daily || []);
      const allowances = (Array.isArray(day.allowances) ? day.allowances : [])
        .map(a => String(a || '').trim())
        .filter(a => dailyOpts.has(a) && allowedOnDate(siteAllow.days, a, date));   // daily allowances for the day's site, on their allowed days
      if (!lines.length && !allowances.length) continue;
      const shiftWorked = String(day.shiftWorked || '').trim().slice(0, 100);
      const workShifts = (roster.byDate[date] || []).filter(x => !allowanceForShift(x.name, allowancesBySite[x.site]));   // On call etc. are allowances, not the shift
      clean.push({ date, site, lines, allowances, shiftWorked, rosteredShift: rosterLabel(workShifts).slice(0, 255) });
    }
    const weekAllowances = (Array.isArray(data.weekAllowances) ? data.weekAllowances : [])
      .map(a => String(a || '').trim())
      .filter(a => weeklySet.has(a));             // weekly allowances = the home site's

    // The new rows, grouped so each day is written in order (line order is kept);
    // different days are written side by side.
    const batchId = 'S-' + target.weekStart + '-' + Math.random().toString(36).slice(2, 8).toUpperCase();
    const groups = [];
    for (const day of clean) {
      const shiftCols = {};                           // stamped on every row of the day
      if (day.shiftWorked) shiftCols.ShiftWorked = day.shiftWorked;
      if (day.rosteredShift) shiftCols.RosteredShift = day.rosteredShift;
      const rows = [];
      for (const l of day.lines) {
        const fields = {
          Title: user.name, ContractorId: user.id, EntryDate: day.date + 'T00:00:00Z',
          Site: day.site, WorkOrder: l.wo, Hours: l.hr, Status: 'Submitted', BatchID: batchId,
          Notes: 'Staff app · ' + user.email, ...shiftCols,
        };
        if (l.comment) fields.Comment = l.comment;
        if (l.co) fields.Callout = 'Yes';                 // this job was a call-out
        rows.push(fields);
      }
      for (const a of day.allowances) {
        rows.push({
          Title: user.name, ContractorId: user.id, EntryDate: day.date + 'T00:00:00Z',
          Site: day.site, RowType: 'Allowance', Allowance: a, Hours: 0, Status: 'Submitted', BatchID: batchId,
          Notes: 'Staff app allowance · ' + user.email, ...shiftCols,
        });
      }
      groups.push(rows);
    }
    if (weekAllowances.length) groups.push(weekAllowances.map(a => ({
      Title: user.name, ContractorId: user.id, EntryDate: target.weekStart + 'T00:00:00Z',
      Site: homeSite, RowType: 'Allowance', Allowance: a, Hours: 0, Status: 'Submitted', BatchID: batchId,
      Notes: 'Staff app weekly allowance · ' + user.email,
    })));

    // REPLACE the week: write the new rows first; if anything fails, remove what was
    // written and leave the old week untouched. Only after a full write are the old rows removed.
    const existing = await getUserEntries(token, user.id, target.weekStart, target.weekEnd);
    const cols = {};                                  // optional columns found missing + comment counts (shared by the save)
    const createdIds = [];
    try {
      await pool(groups, 4, async rows => { for (const f of rows) createdIds.push(await createItemResilient(token, f, cols)); });
    } catch (e) {
      await pool(createdIds.slice(), 4, id => deleteItem(token, id).catch(() => {}));
      throw e;
    }
    await pool(existing, 4, it => deleteItem(token, it.id));
    const created = createdIds.length;
    return { statusCode: 200, headers, body: JSON.stringify({ ok: true, created, commentsShortened: cols.shortened || 0, commentsDropped: cols.dropped || 0, weekStart: wk.weekStart, weekEnd: wk.weekEnd }) };
  } catch (err) {
    console.error(err);
    return { statusCode: 502, headers, body: JSON.stringify({ ok: false, error: 'server' }) };
  }
};
