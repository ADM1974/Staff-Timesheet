// Returns the signed-in staff member's current open week + what they've entered.
// The week is driven by their HOME site's close day (Workers list, by email).
const { getAppToken, validateStaffToken, getActiveSites, getSitesMap, getHomeSite, getAllowancesBySite, weekContext, getUserEntries, getWorkOrders, getRejectedWeek } = require('./staff');

exports.handler = async (event) => {
  const headers = { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' };
  const user = await validateStaffToken(event.headers && (event.headers.authorization || event.headers.Authorization));
  if (!user) return { statusCode: 401, headers, body: '{"ok":false}' };

  try {
    const token = await getAppToken();
    const sitesMap = await getSitesMap(token);
    const homeSite = await getHomeSite(token, user.email, sitesMap);
    const homeInfo = sitesMap.byName[String(homeSite || '').toLowerCase()] || {};
    const closeDayIdx = homeInfo.closeDayIndex || 0;
    const workOrders = await getWorkOrders(token, homeInfo.id, homeSite);   // typeahead for the home site
    const wk = weekContext(closeDayIdx);
    const rejected = await getRejectedWeek(token, user.id, closeDayIdx);    // a sent-back week to fix?
    const active = rejected ? { ...rejected, closeDayName: wk.closeDayName } : wk;  // edit that week, else the open week
    const sites = await getActiveSites(token);
    const allowancesBySite = await getAllowancesBySite(token, sitesMap);   // {site:{daily,weekly}}
    const weeklyOptions = ((allowancesBySite[homeSite] || {}).weekly) || []; // staff weekly = home site
    const weeklySet = new Set(weeklyOptions);
    const entries = await getUserEntries(token, user.id, active.weekStart, active.weekEnd);

    // group existing rows → per-day site/hours/daily-allowances + week-level weekly allowances
    const byDate = {};
    const weekAllowances = [];
    const day = d => (byDate[d] = byDate[d] || { site: '', lines: [], allowances: [] });
    for (const it of entries) {
      const f = it.fields || {};
      const date = String(f.EntryDate || '').slice(0, 10);
      const dd = day(date);
      if (f.Site && !dd.site) dd.site = f.Site;
      if (String(f.RowType || 'Hours') === 'Allowance') {
        const a = String(f.Allowance || f.WorkOrder || '').trim();
        if (!a) continue;
        if (weeklySet.has(a)) { if (!weekAllowances.includes(a)) weekAllowances.push(a); }
        else dd.allowances.push(a);
      } else {
        dd.lines.push({ wo: String(f.WorkOrder || ''), hr: Number(f.Hours) || 0 });
      }
    }
    const days = active.days.map(date => ({
      date,
      site: (byDate[date] && byDate[date].site) || '',
      lines: (byDate[date] && byDate[date].lines) || [],
      allowances: (byDate[date] && byDate[date].allowances) || [],
    }));

    return { statusCode: 200, headers, body: JSON.stringify({ ok: true, name: user.name, homeSite, closeDayName: wk.closeDayName, weekStart: active.weekStart, weekEnd: active.weekEnd, sites, allowancesBySite, weeklyOptions, workOrders, weekAllowances, days, rejected: !!rejected, rejectionReason: rejected ? rejected.reason : '' }) };
  } catch (err) {
    console.error(err);
    return { statusCode: 502, headers, body: '{"ok":false,"error":"server"}' };
  }
};
