// Leave requests for signed-in staff. POST { action, ... }:
//   init                          → switch state, leave types for the home site, my requests
//   preview { type,start,end,hours }   → working days, hours, any rules it breaks (live warnings)
//   submit  { type,start,end,hours,note,evidence? }   evidence = { name, mime, data(base64) }
//   cancel  { id }                → withdraw a pending request, or future approved leave
//   evidence { id, evidence }     → add a medical certificate / stat dec afterwards
// Approval happens in the supervisor app (see leave-core.js for the flow and lists).
const { getAppToken, validateStaffToken, getSitesMap, getWorkerInfo } = require('./staff');
const { holidaysForSite } = require('./holidays');
const L = require('./leave-core');

const headers = { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' };
const reply = (code, obj) => ({ statusCode: code, headers, body: JSON.stringify(obj) });

// What the person sees of their own request.
const mine = r => ({ id: r.id, type: r.type, start: r.start, end: r.end, hours: r.hours, days: L.daysOf(r).length, status: r.status, rules: r.rules, note: r.note, hasEvidence: !!r.evidenceUrl, needsEvidence: L.needsEvidence(r) });

exports.handler = async (event) => {
  const user = await validateStaffToken(event.headers && (event.headers.authorization || event.headers.Authorization));
  if (!user) return reply(401, { ok: false });
  let data;
  try { data = JSON.parse(event.body || '{}'); } catch (_) { return reply(400, { ok: false, error: 'bad json' }); }
  const action = String(data.action || 'init');
  const email = String(user.email || '').trim().toLowerCase();

  try {
    const token = await getAppToken();
    const sw = await L.leaveSwitch(token);
    if (!L.leaveOnFor(sw, email)) return reply(200, { ok: true, on: false });

    const sitesMap = await getSitesMap(token);
    const [{ homeSite, contractHours }, types, leaveSites, { listId, requests }] = await Promise.all([
      getWorkerInfo(token, email, sitesMap), L.getLeaveTypes(token), L.getLeaveSites(token), L.getRequests(token),
    ]);
    if (!listId) return reply(200, { ok: true, on: true, setup: 'The LeaveRequests list hasn’t been set up in SharePoint yet.' });
    const site = homeSite || '';
    const siteInfo = leaveSites[site] || {};
    const myTypes = L.typesForSite(types, siteInfo.id || '');
    const hoursPerDay = contractHours > 0 ? Math.round(contractHours / 5 * 100) / 100 : 8;
    const tz = L.tzForSite(site), today = L.todayIn(tz);
    const own = requests.filter(r => r.email === email || (r.personId && r.personId === user.id));

    if (action === 'init') {
      const cutoff = L.isoOf(L.msOf(today) - 400 * 86400000);
      return reply(200, {
        ok: true, on: true, site, hoursPerDay, today,
        types: myTypes.map(t => ({ title: t.title, entitlement: t.entitlement, needsAttachment: t.needsAttachment, allowPast: t.allowPast })),
        requests: own.filter(r => r.end >= cutoff).sort((a, b) => b.start.localeCompare(a.start)).map(mine),
      });
    }

    // ---- build / check a request ----
    const plan = async () => {
      const type = myTypes.find(t => t.title === String(data.type || ''));
      if (!type) return { error: 'Choose a leave type.' };
      const start = String(data.start || ''), end = String(data.end || data.start || '');
      if (!L.validDate(start) || !L.validDate(end)) return { error: 'Choose the dates.' };
      if (end < start) return { error: 'The end date is before the start date.' };
      if (L.datesBetween(start, end).length > 180) return { error: 'That’s more than 6 months — split it into separate requests.' };
      if (start < today && !type.allowPast) return { error: `${type.title} can’t be backdated — pick today or later.` };
      const sm = sitesMap.byName[site.toLowerCase()];
      // a day either side too, for the "next to a public holiday" certificate rule
      const hols = sm ? await holidaysForSite(token, sm, L.isoOf(L.msOf(start) - 86400000), L.isoOf(L.msOf(end) + 86400000)).catch(() => ({})) : {};
      const days = L.workingDays(start, end, hols);
      if (!days.length) return { error: 'Those dates are all weekends or public holidays.' };
      const clash = own.find(r => L.ACTIVE.includes(r.status) && L.daysOf(r).some(d => days.includes(d)));
      if (clash) return { error: `You already have ${clash.type} on some of those days.` };
      let per = hoursPerDay;
      const asked = parseFloat(data.hours);
      if (days.length === 1 && asked > 0 && asked <= 24) per = Math.round(asked * 100) / 100;   // part day
      const dayHours = Object.fromEntries(days.map(d => [d, per]));
      const others = L.othersOff(requests, site, days, '', email);
      const rules = L.evaluateRules({ type, start, days, today, site, siteMax: siteInfo.max || 0, others });
      const skipped = Object.entries(hols).filter(([d]) => d >= start && d <= end).map(([d, n]) => `${L.fmtDay(d)} is ${n} — not counted`);
      return { type, start, end, days, dayHours, hours: Math.round(per * days.length * 100) / 100, rules, skipped, evidence: L.evidenceReason(type, days, hols) };
    };

    if (action === 'preview') {
      const p = await plan();
      if (p.error) return reply(200, { ok: true, error: p.error });
      return reply(200, { ok: true, days: p.days.length, hours: p.hours, rules: p.rules, skipped: p.skipped, entitlement: p.type.entitlement, evidence: p.evidence });
    }

    if (action === 'submit') {
      const p = await plan();
      if (p.error) return reply(400, { ok: false, error: p.error });
      let evidenceUrl = '';
      if (data.evidence && data.evidence.data) evidenceUrl = await L.uploadEvidence(token, user.name, data.evidence.name, data.evidence.mime, data.evidence.data);
      const r = await L.createRequest(token, listId, {
        Title: user.name, Email: email, PersonType: 'Staff', PersonId: user.id, Site: site,
        LeaveType: p.type.title, StartDate: p.start, EndDate: p.end, Hours: p.hours, DayHours: JSON.stringify(p.dayHours),
        Note: String(data.note || '').trim().slice(0, 1000), Status: 'Pending', RulesBroken: p.rules.join('\n'),
        History: L.stamp(user.name, 'Requested') + (p.evidence ? '\n' + L.stamp('', 'Certificate needed — ' + p.evidence) : ''),
        ...(evidenceUrl ? { EvidenceUrl: evidenceUrl } : {}),
      });
      return reply(200, { ok: true, request: mine(r), rules: p.rules });
    }

    const r = own.find(x => x.id === String(data.id || ''));
    if (!r) return reply(404, { ok: false, error: 'Request not found.' });

    if (action === 'cancel') {
      const future = r.start > today;
      if (!(['Pending', 'Supervisor OK'].includes(r.status) || (L.FINAL_OK.includes(r.status) && future)))
        return reply(400, { ok: false, error: 'Leave that has started can’t be cancelled here — talk to your supervisor.' });
      const teamId = (leaveSites[r.site] || {}).teamId;
      if (r.timeOffId && teamId) await L.removeTimeOff(token, teamId, r.timeOffId).catch(e => console.warn('leave cancel shifts', e.message));
      await L.updateRequest(token, listId, r.id, { Status: 'Cancelled', ShiftsTimeOffId: '', History: L.addHistory(r, L.stamp(user.name, 'Cancelled')) });
      return reply(200, { ok: true });
    }

    if (action === 'evidence') {
      if (!data.evidence || !data.evidence.data) return reply(400, { ok: false, error: 'No file.' });
      const url = await L.uploadEvidence(token, user.name, data.evidence.name, data.evidence.mime, data.evidence.data);
      await L.updateRequest(token, listId, r.id, { EvidenceUrl: url, History: L.addHistory(r, L.stamp(user.name, 'Evidence added')) });
      return reply(200, { ok: true });
    }

    return reply(400, { ok: false, error: 'bad action' });
  } catch (err) {
    console.error('leave', err);
    if (/not recognized|does not exist|invalidRequest/i.test(String(err.message))) return reply(500, { ok: false, error: 'SharePoint setup: ' + String(err.message).replace(/^.*?"message":"?/, '').slice(0, 160) });
    return reply(502, { ok: false, error: 'Couldn’t reach SharePoint — try again.' });
  }
};
