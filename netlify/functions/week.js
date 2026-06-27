// Returns the signed-in staff member's current open week + what they've entered.
// The week is driven by their HOME site's close day (Workers list, by email).
const { getAppToken, validateStaffToken, getActiveSites, getSitesMap, getHomeSite, getAllowancesBySite, weekContext, getUserEntries } = require('./staff');

exports.handler = async (event) => {
  const headers = { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' };
  const user = await validateStaffToken(event.headers && (event.headers.authorization || event.headers.Authorization));
  if (!user) return { statusCode: 401, headers, body: '{"ok":false}' };

  try {
    const token = await getAppToken();
    const sitesMap = await getSitesMap(token);
    const homeSite = await getHomeSite(token, user.email, sitesMap);
    const closeDayIdx = (sitesMap.byName[String(homeSite || '').toLowerCase()] || {}).closeDayIndex || 0;
    const wk = weekContext(closeDayIdx);
    const sites = await getActiveSites(token);
    const allowancesBySite = await getAllowancesBySite(token, sitesMap);
    const entries = await getUserEntries(token, user.id, wk.weekStart, wk.weekEnd);

    // group existing rows by date → site + hours lines + ticked allowances
    const byDate = {};
    for (const it of entries) {
      const f = it.fields || {};
      const date = String(f.EntryDate || '').slice(0, 10);
      if (!byDate[date]) byDate[date] = { site: '', lines: [], allowances: [] };
      if (f.Site && !byDate[date].site) byDate[date].site = f.Site;
      if (String(f.RowType || 'Hours') === 'Allowance') {
        const a = String(f.Allowance || f.WorkOrder || '').trim();
        if (a) byDate[date].allowances.push(a);
      } else {
        byDate[date].lines.push({ wo: String(f.WorkOrder || ''), hr: Number(f.Hours) || 0 });
      }
    }
    const days = wk.days.map(date => ({
      date,
      site: (byDate[date] && byDate[date].site) || '',
      lines: (byDate[date] && byDate[date].lines) || [],
      allowances: (byDate[date] && byDate[date].allowances) || [],
    }));

    return { statusCode: 200, headers, body: JSON.stringify({ ok: true, name: user.name, homeSite, closeDayName: wk.closeDayName, weekStart: wk.weekStart, weekEnd: wk.weekEnd, sites, allowancesBySite, days }) };
  } catch (err) {
    console.error(err);
    return { statusCode: 502, headers, body: '{"ok":false,"error":"server"}' };
  }
};
