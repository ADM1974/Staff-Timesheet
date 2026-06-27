// Saves the signed-in staff member's whole week — REPLACES their rows for the
// open week (delete + rewrite). Week boundary is the home site's; hours are per
// the day's chosen site; allowances are validated against that day's site.
const { getAppToken, validateStaffToken, getActiveSites, getSitesMap, getHomeSite, getAllowancesBySite, weekContext, getUserEntries, createItem, deleteItem } = require('./staff');

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
    const open = new Set(wk.days);
    const activeSites = new Set(await getActiveSites(token));
    const allowancesBySite = await getAllowancesBySite(token, sitesMap);

    const clean = [];
    for (const day of (Array.isArray(data.days) ? data.days : [])) {
      const date = String(day.date || '').slice(0, 10);
      if (!open.has(date)) continue;
      const site = String(day.site || '').trim();
      if (!site) continue;
      if (!activeSites.has(site)) return bad('Unknown or inactive site: ' + site);
      const lines = (Array.isArray(day.lines) ? day.lines : [])
        .map(l => ({ wo: String(l.wo || '').trim().slice(0, 200), hr: parseFloat(l.hr) }))
        .filter(l => l.wo && l.hr > 0 && l.hr <= 24);
      const allowOpts = new Set(allowancesBySite[site] || []);
      const allowances = (Array.isArray(day.allowances) ? day.allowances : [])
        .map(a => String(a || '').trim())
        .filter(a => allowOpts.has(a));           // only valid for the day's site
      if (!lines.length && !allowances.length) continue;
      clean.push({ date, site, lines, allowances });
    }

    const existing = await getUserEntries(token, user.id, wk.weekStart, wk.weekEnd);
    for (const it of existing) await deleteItem(token, it.id);

    const batchId = 'S-' + wk.weekStart + '-' + Math.random().toString(36).slice(2, 8).toUpperCase();
    let created = 0;
    for (const day of clean) {
      for (const l of day.lines) {
        await createItem(token, {
          Title: user.name, ContractorId: user.id, EntryDate: day.date + 'T00:00:00Z',
          Site: day.site, WorkOrder: l.wo, Hours: l.hr, Status: 'Submitted', BatchID: batchId,
          Notes: 'Staff app · ' + user.email,
        });
        created++;
      }
      for (const a of day.allowances) {
        await createItem(token, {
          Title: user.name, ContractorId: user.id, EntryDate: day.date + 'T00:00:00Z',
          Site: day.site, RowType: 'Allowance', Allowance: a, Hours: 0, Status: 'Submitted', BatchID: batchId,
          Notes: 'Staff app allowance · ' + user.email,
        });
        created++;
      }
    }
    return { statusCode: 200, headers, body: JSON.stringify({ ok: true, created, weekStart: wk.weekStart, weekEnd: wk.weekEnd }) };
  } catch (err) {
    console.error(err);
    return { statusCode: 502, headers, body: JSON.stringify({ ok: false, error: 'server' }) };
  }
};
