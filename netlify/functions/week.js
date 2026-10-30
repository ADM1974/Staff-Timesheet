// Returns the signed-in staff member's current open week + what they've entered.
// The week is driven by their HOME site's close day (Workers list, by email).
// Also returns each day's rostered shift (Teams Shifts, as captured at the start of the
// week — see roster-snapshot.js) and what was saved: the shift they say they worked.
const { getAppToken, validateStaffToken, getSitesMap, getWorkerInfo, getAllowancesBySite, weekContext, getUserEntries, getUserRows, getWorkOrders, getRejectedWeek, shiftKey } = require('./staff');
const { getRoster, allowanceForShift } = require('./roster-snapshot');
const { holidaysBySite } = require('./holidays');

exports.handler = async (event) => {
  const headers = { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' };
  const user = await validateStaffToken(event.headers && (event.headers.authorization || event.headers.Authorization));
  if (!user) return { statusCode: 401, headers, body: '{"ok":false}' };

  try {
    // Independent SharePoint reads run side by side — one after another they took ~8s.
    const token = await getAppToken();
    const [sitesMap, rows] = await Promise.all([getSitesMap(token), getUserRows(token, user.id)]);
    const sites = sitesMap.active;
    const [{ homeSite, contractHours }, allowancesBySite] = await Promise.all([
      getWorkerInfo(token, user.email, sitesMap),                          // home site + weekly contract hours
      getAllowancesBySite(token, sitesMap),                                // {site:{daily,weekly,perJob}}
    ]);
    const homeInfo = sitesMap.byName[String(homeSite || '').toLowerCase()] || {};
    const closeDayIdx = homeInfo.closeDayIndex || 0;
    const wk = weekContext(closeDayIdx);
    const rejected = await getRejectedWeek(token, user.id, closeDayIdx, rows);    // a sent-back week to fix?
    const active = rejected ? { ...rejected, closeDayName: wk.closeDayName } : wk;  // edit that week, else the open week
    const weeklyOptions = ((allowancesBySite[homeSite] || {}).weekly) || []; // staff weekly = home site
    const weeklySet = new Set(weeklyOptions);
    const entries = await getUserEntries(token, user.id, active.weekStart, active.weekEnd, rows);
    const [workOrders, roster, holidays] = await Promise.all([
      getWorkOrders(token, homeInfo.id, homeSite),                         // typeahead for the home site
      getRoster(token, user.id, sitesMap, active.days),                    // fixed at the start of each week
      holidaysBySite(token, sitesMap, active.weekStart, active.weekEnd).catch(() => ({})),   // {site:{date:name}}
    ]);

    // Split the roster: work shifts vs shifts that stand for a weekly allowance (On call).
    // The allowance ones tick that allowance instead of being the shift worked.
    const workRoster = {}, allowRoster = {}, rosterWeekAllowances = [];
    for (const [date, list] of Object.entries(roster.byDate)) {
      for (const x of list) {
        const a = allowanceForShift(x.name, allowancesBySite[x.site]);
        if (!a) { (workRoster[date] = workRoster[date] || []).push(x); continue; }
        if (!(allowRoster[date] = allowRoster[date] || []).includes(a)) allowRoster[date].push(a);
        if (weeklySet.has(a) && !rosterWeekAllowances.includes(a)) rosterWeekAllowances.push(a);
      }
    }
    const rosterShiftNames = {};
    for (const [site, names] of Object.entries(roster.shiftNames || {})) rosterShiftNames[site] = names.filter(n => !allowanceForShift(n, allowancesBySite[site]));

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
        // Per-job allowances (JobAllowances). Rows saved before v27 have Callout = Yes instead.
        let jobs = String(f.JobAllowances || '').split(/[;,]/).map(x => x.trim()).filter(Boolean);
        if (!jobs.length && (f.Callout === true || /^(yes|true|1)$/i.test(String(f.Callout || '')))) {
          const co = (((allowancesBySite[String(f.Site || '').trim()] || {}).perJob) || []).find(a => shiftKey(a) === 'callout');
          jobs = [co || 'Call-out'];
        }
        dd.lines.push({ wo: String(f.WorkOrder || ''), hr: Number(f.Hours) || 0, comment: String(f.Comment || ''), jobs });
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
      rostered: workRoster[date] || [],
      rosterAllowances: allowRoster[date] || [],
    }));

    return { statusCode: 200, headers, body: JSON.stringify({ ok: true, name: user.name, homeSite, contractHours, holidays, closeDayName: wk.closeDayName, weekStart: active.weekStart, weekEnd: active.weekEnd, sites, allowancesBySite, weeklyOptions, workOrders, weekAllowances, days, rosterSites: Object.values(sitesMap.byName).filter(x => x.shiftsTeamId).map(x => x.name), rosterErrors: roster.errors, rosterErrorDetail: roster.details || {}, rosterShiftNames, rosterWeekAllowances, weekSaved: entries.length > 0, rejected: !!rejected, rejectionReason: rejected ? rejected.reason : '' }) };
  } catch (err) {
    console.error(err);
    return { statusCode: 502, headers, body: '{"ok":false,"error":"server"}' };
  }
};
