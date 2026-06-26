// Returns the signed-in staff member's current open week + what they've entered.
const { getAppToken, validateStaffToken, getActiveSites, getWeekContext, getUserEntries } = require('./staff');

exports.handler = async (event) => {
  const headers = { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' };
  const user = await validateStaffToken(event.headers && (event.headers.authorization || event.headers.Authorization));
  if (!user) return { statusCode: 401, headers, body: '{"ok":false}' };

  try {
    const token = await getAppToken();
    const wk = await getWeekContext(token);
    const sites = await getActiveSites(token);
    const entries = await getUserEntries(token, user.id, wk.weekStart, wk.weekEnd);

    const byDate = {};
    for (const it of entries) {
      const f = it.fields || {};
      const date = String(f.EntryDate || '').slice(0, 10);
      if (!byDate[date]) byDate[date] = { site: '', lines: [] };
      if (f.Site && !byDate[date].site) byDate[date].site = f.Site;
      byDate[date].lines.push({ wo: String(f.WorkOrder || ''), hr: Number(f.Hours) || 0 });
    }
    const days = wk.days.map(date => ({
      date,
      site: (byDate[date] && byDate[date].site) || '',
      lines: (byDate[date] && byDate[date].lines) || [],
    }));

    return { statusCode: 200, headers, body: JSON.stringify({ ok: true, name: user.name, closeDayName: wk.closeDayName, weekStart: wk.weekStart, weekEnd: wk.weekEnd, sites, days }) };
  } catch (err) {
    console.error(err);
    return { statusCode: 502, headers, body: '{"ok":false,"error":"server"}' };
  }
};
