// Public holidays per site, shown on the timesheet and put on the Teams Shifts schedule
// as day notes (see holiday-notes.js).
//
// Each site's region comes from the optional Sites column "HolidayRegion": a country
// ("AU", "NZ") or a state / province ("AU-QLD", "AU-NSW", "NZ-AUK"). With no column the
// region is guessed from the site code (BNE → AU-QLD, SYD → AU-NSW, AKL → NZ-AUK …).
// Holidays come from the free public-holiday service date.nager.at; extras it doesn't
// know (show days, one-offs) can be added to an optional SharePoint list "PublicHolidays":
// Title = the holiday's name, HolidayDate = the date, Region = blank (every site) or a
// region / site name, e.g. "AU-QLD; BNE".

const { findList } = require('./staff');

// Airport codes → holiday region, used when a site has no HolidayRegion.
const CODE_REGION = {
  BNE: 'AU-QLD', OOL: 'AU-QLD', CNS: 'AU-QLD', TSV: 'AU-QLD',
  SYD: 'AU-NSW', WSI: 'AU-NSW', NTL: 'AU-NSW',
  MEL: 'AU-VIC', AVV: 'AU-VIC', ADL: 'AU-SA', PER: 'AU-WA', CBR: 'AU-ACT', HBA: 'AU-TAS', DRW: 'AU-NT',
  AKL: 'NZ-AUK', WLG: 'NZ-WGN', CHC: 'NZ-CAN', ZQN: 'NZ-OTA', DUD: 'NZ-OTA',
};
function regionForSite(site) {
  const set = String((site && site.holidayRegion) || '').trim().toUpperCase();
  if (set) return set;
  const code = (/^[A-Z]{3}\b/.exec(String((site && site.name) || '').trim().toUpperCase()) || [])[0];
  return CODE_REGION[code] || '';
}

// date.nager.at, one request per country-year (kept while the function stays warm).
const _years = {};
function countryYear(country, year) {
  const k = country + year;
  if (!_years[k]) {
    _years[k] = fetch(`https://date.nager.at/api/v3/PublicHolidays/${year}/${country}`)
      .then(r => (r.ok ? r.json() : []))
      .catch(() => [])
      .then(list => { if (!list.length) delete _years[k]; return list; });   // try again next time
  }
  return _years[k];
}

// The optional "PublicHolidays" list (cached 10 min).
let _extra = null, _extraAt = 0;
async function extraHolidays(token) {
  if (_extra && Date.now() - _extraAt < 600000) return _extra;
  const out = [];
  try {
    const list = await findList(token, 'publicholidays');
    if (list) {
      const r = await fetch(`https://graph.microsoft.com/v1.0/sites/${process.env.SP_SITE_ID}/lists/${list.id}/items?expand=fields&$top=999`, { headers: { Authorization: 'Bearer ' + token } });
      if (r.ok) for (const it of ((await r.json()).value || [])) {
        const f = it.fields || {};
        const date = String(f.HolidayDate || f.Date || '').slice(0, 10);
        const name = String(f.Title || '').trim();
        if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !name) continue;
        out.push({ date, name, regions: String(f.Region || '').split(/[;,]/).map(x => x.trim().toUpperCase()).filter(Boolean) });
      }
    }
  } catch (e) { console.warn('PublicHolidays list', e.message); }
  _extra = out; _extraAt = Date.now();
  return out;
}

// { 'YYYY-MM-DD': 'Holiday name' } for one site between two ISO dates (inclusive).
async function holidaysForSite(token, site, from, to) {
  const out = {};
  const region = regionForSite(site);
  const country = region.split('-')[0];
  if (country) {
    const years = [...new Set([from.slice(0, 4), to.slice(0, 4)])];
    for (const list of await Promise.all(years.map(y => countryYear(country, y)))) {
      for (const h of list) {
        if (h.date < from || h.date > to) continue;
        if (Array.isArray(h.types) && !h.types.includes('Public')) continue;
        const here = h.global || (region.includes('-') && (h.counties || []).includes(region));
        if (here && !out[h.date]) out[h.date] = h.name || h.localName;
      }
    }
  }
  const siteName = String((site && site.name) || '').trim().toUpperCase();
  for (const h of await extraHolidays(token)) {
    if (h.date < from || h.date > to) continue;
    if (!h.regions.length || h.regions.includes(region) || h.regions.includes(country) || h.regions.includes(siteName)) out[h.date] = out[h.date] || h.name;
  }
  return out;
}

// Every active site → { date: name } over [from, to]. Sites with no holidays are left out.
async function holidaysBySite(token, sitesMap, from, to) {
  const out = {};
  await Promise.all(sitesMap.active.map(async name => {
    const site = sitesMap.byName[name.toLowerCase()];
    if (!site) return;
    const h = await holidaysForSite(token, site, from, to);
    if (Object.keys(h).length) out[name] = h;
  }));
  return out;
}

exports.regionForSite = regionForSite;
exports.holidaysForSite = holidaysForSite;
exports.holidaysBySite = holidaysBySite;
