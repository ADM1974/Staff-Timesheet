// Returns the signed-in staff member's current open week + what they've entered.
// The week is driven by their HOME site's close day (Workers list, by email).
// Also returns each day's rostered shift (Teams Shifts, as captured at the start of the
// week — see roster-snapshot.js) and what was saved: the shift they say they worked.
const { getAppToken, validateStaffToken, getSitesMap, getHomeSite, getAllowancesBySite, weekContext, getUserEntries, getUserRows, getWorkOrders, getRejectedWeek } = require('./staff');
const { getRoster } = require('./roster-snapshot');

exports.handler = async (event) => {
  const headers = { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' };
  const user = await validateStaffToken(event.headers && (event.headers.authorization || event.headers.Authorization));
  if (!user) return { statusCode: 401, headers, body: '{"ok":false}' };

  try {
    // Independent SharePoint reads run side by side — one after another they took ~8s.
    const token = await getAppToken();
    const [sitesMap, rows] = await Promise.all([getSitesMap(token), getUserRows(token, user.id)]);
    const sites = sitesMap.active;
    const [homeSite, allowancesBySite] = await Promise.all([
      getHomeSite(token, user.email, sitesMap),
      getAllowancesBySite(token, sitesMap),                                // {site:{daily,weekly}}
    ]);
    const homeInfo = sitesMap.byName[String(homeSite || '').toLowerCase()] || {};
    const closeDayIdx = homeInfo.closeDayIndex || 0;
    const wk = weekContext(closeDayIdx);
    const rejected = await getRejectedWeek(token, user.id, closeDayIdx, rows);    // a sent-back week to fix?
    const active = rejected ? { ...rejected, closeDayName: wk.closeDayName } : wk;  // edit that week, else the open week
    const weeklyOptions = ((allowancesBySite[homeSite] || {}).weekly) || []; // staff weekly = home site
    const weeklySet = new Set(weeklyOptions);
    const entries = await getUserEntries(token, user.id, active.weekStart, active.weekEnd, rows);
    const [workOrders, roster] = await Promise.all([
      getWorkOrders(token, homeInfo.id, homeSite),                         // typeahead for the home site
      getRoster(token, user.id, sitesMap, active.days),                    // fixed at the start of each week
    ]);

    // group existing rows → per-day site/hours/daily-allowances + week-level weekly allowances
    const byDate = {};
    const weekAllowances = [];
    const day = d => (byDate[d] = byDate[d] || { site: '', lines: [], allowances: [], shiftWorked: '', rosteredAtSave: '' });
    for (const it of entries) {
      const f = it.fields || {};
      const date = String(f.EntryDate || '').slice(0, 10);
      const dd = day(date);
      if (f.Site && !dd.site) dd.site = f.Site;
      if (f.ShiftWorked && !dd.shiftWorked) dd.shiftWorked = String(f.ShiftWorked);
      if (f.RosteredShift && !dd.rosteredAtSave) dd.rosteredAtSave = String(f.RosteredShift);
      if (String(f.RowType || 'Hours') === 'Allowance') {
        const a = String(f.Allowance || f.WorkOrder || '').trim();
        if (!a) continue;
        if (weeklySet.has(a)) { if (!weekAllowances.includes(a)) weekAllowances.push(a); }
        else dd.allowances.push(a);
      } else {
        dd.lines.push({ wo: String(f.WorkOrder || ''), hr: Number(f.Hours) || 0, comment: String(f.Comment || '') });
      }
    }
    const days = active.days.map(date => ({
      date,
      site: (byDate[date] && byDate[date].site) || '',
      lines: (byDate[date] && byDate[date].lines) || [],
      allowances: (byDate[date] && byDate[date].allowances) || [],
      saved: !!byDate[date],
      shiftWorked: (byDate[date] && byDate[date].shiftWorked) || '',
      rosteredAtSave: (byDate[date] && byDate[date].rosteredAtSave) || '',
      rostered: roster.byDate[date] || [],
    }));

    return { statusCode: 200, headers, body: JSON.stringify({ ok: true, name: user.name, homeSite, closeDayName: wk.closeDayName, weekStart: active.weekStart, weekEnd: active.weekEnd, sites, allowancesBySite, weeklyOptions, workOrders, weekAllowances, days, rosterSites: Object.values(sitesMap.byName).filter(x => x.shiftsTeamId).map(x => x.name), rosterErrors: roster.errors, rosterErrorDetail: roster.details || {}, rosterShiftNames: roster.shiftNames || {}, rejected: !!rejected, rejectionReason: rejected ? rejected.reason : '' }) };
  } catch (err) {
    console.error(err);
    return { statusCode: 502, headers, body: '{"ok":false,"error":"server"}' };
  }
};
