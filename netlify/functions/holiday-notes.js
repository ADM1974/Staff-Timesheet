// Scheduled (hourly): puts each public holiday on the Teams Shifts schedule of every
// Shifts-linked site (Sites.ShiftsTeamId) as a day note — "🎉 Public holiday – Christmas Day" —
// for the next 120 days. A day that already has a note naming the holiday is left alone,
// so it's safe to run again and again. Holidays are the site's (see holidays.js).
//
// Writing to Shifts needs the Graph APPLICATION permission Schedule.ReadWrite.All
// (admin-consented) on the CLIENT_ID app registration — until then each site just
// reports a 403 and nothing else changes.
const { getAppToken, getSitesMap } = require('./staff');
const { holidaysForSite } = require('./holidays');

const DAYS_AHEAD = 120;
const iso = ms => new Date(ms).toISOString().slice(0, 10);
const noteText = n => (n && ((n.sharedDayNote || {}).content || (n.draftDayNote || {}).content)) || '';

async function existingNotes(token, teamId) {
  const out = [];
  let url = `https://graph.microsoft.com/v1.0/teams/${teamId}/schedule/dayNotes`;
  while (url) {
    const r = await fetch(url, { headers: { Authorization: 'Bearer ' + token } });
    if (!r.ok) throw new Error('read day notes ' + r.status + ' ' + (await r.text()).slice(0, 300));
    const j = await r.json();
    out.push(...(j.value || []));
    url = j['@odata.nextLink'] || '';
  }
  return out;
}

async function addNote(token, teamId, date, text) {
  const body = { contentType: 'text', content: text };
  const r = await fetch(`https://graph.microsoft.com/v1.0/teams/${teamId}/schedule/dayNotes`, {
    method: 'POST',
    headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify({ dayNoteDate: date, sharedDayNote: body, draftDayNote: body }),   // shared = what the team sees
  });
  if (!r.ok) throw new Error('add day note ' + r.status + ' ' + (await r.text()).slice(0, 300));
}

exports.handler = async () => {
  const out = [];
  try {
    const token = await getAppToken();
    const sitesMap = await getSitesMap(token);
    const from = iso(Date.now()), to = iso(Date.now() + DAYS_AHEAD * 86400000);
    for (const name of sitesMap.active) {
      const site = sitesMap.byName[name.toLowerCase()];
      if (!site || !site.shiftsTeamId) continue;
      try {
        const hols = await holidaysForSite(token, site, from, to);
        if (!Object.keys(hols).length) { out.push(`${name}: no holidays`); continue; }
        const notes = await existingNotes(token, site.shiftsTeamId);
        let added = 0;
        for (const [date, hol] of Object.entries(hols)) {
          const already = notes.some(n => String(n.dayNoteDate || '').slice(0, 10) === date && noteText(n).toLowerCase().includes(hol.toLowerCase()));
          if (already) continue;
          await addNote(token, site.shiftsTeamId, date, `🎉 Public holiday – ${hol}`);
          added++;
        }
        out.push(`${name}: ${added} added`);
      } catch (e) {
        out.push(`${name}: ${e.message}`);
      }
    }
  } catch (e) {
    out.push('failed: ' + e.message);
  }
  console.log('holiday-notes', out.join(' | '));
  return { statusCode: 200, body: out.join('\n') };
};
