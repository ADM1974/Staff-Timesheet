// Saves the signed-in staff member's whole week — REPLACES their rows for the
// open week (delete + rewrite), so re-submitting overwrites cleanly. Days
// outside the open week are ignored (closed weeks can't be edited).
const { getAppToken, validateStaffToken, getActiveSites, getWeekContext, getUserEntries, createItem, deleteItem } = require('./staff');

exports.handler = async (event) => {
  const headers = { 'Content-Type': 'application/json' };
  const bad = msg => ({ statusCode: 400, headers, body: JSON.stringify({ ok: false, error: msg }) });

  const user = await validateStaffToken(event.headers && (event.headers.authorization || event.headers.Authorization));
  if (!user) return { statusCode: 401, headers, body: '{"ok":false}' };

  let data;
  try { data = JSON.parse(event.body || '{}'); } catch { return bad('bad json'); }

  try {
    const token = await getAppToken();
    const wk = await getWeekContext(token);
    const open = new Set(wk.days);
    const activeSites = new Set(await getActiveSites(token));

    const clean = [];
    for (const day of (Array.isArray(data.days) ? data.days : [])) {
      const date = String(day.date || '').slice(0, 10);
      if (!open.has(date)) continue;
      const site = String(day.site || '').trim();
      const lines = (Array.isArray(day.lines) ? day.lines : [])
        .map(l => ({ wo: String(l.wo || '').trim().slice(0, 200), hr: parseFloat(l.hr) }))
        .filter(l => l.wo && l.hr > 0 && l.hr <= 24);
      if (!site || !lines.length) continue;
      if (!activeSites.has(site)) return bad('Unknown or inactive site: ' + site);
      clean.push({ date, site, lines });
    }

    const existing = await getUserEntries(token, user.id, wk.weekStart, wk.weekEnd);
    for (const it of existing) await deleteItem(token, it.id);

    const batchId = 'S-' + wk.weekStart + '-' + Math.random().toString(36).slice(2, 8).toUpperCase();
    let created = 0;
    for (const day of clean) {
      for (const l of day.lines) {
        await createItem(token, {
          Title: user.name,
          ContractorId: user.id,                 // staff 365 object id (stable key)
          EntryDate: day.date + 'T00:00:00Z',
          Site: day.site,
          WorkOrder: l.wo,
          Hours: l.hr,
          Status: 'Submitted',
          BatchID: batchId,
          Notes: 'Staff app · ' + user.email,
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
