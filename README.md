# Staff Timesheet — Microsoft 365 sign-in

A separate staff app: the same weekly timesheet as the contractor app, but staff
**sign in with their Microsoft 365 account** instead of a code. On a company phone
that's signed into 365, it's usually one tap (SSO). Hours write to the **same
Timesheets list**, keyed by each person's 365 account id.

```
Staff phone → "Sign in with Microsoft" (MSAL) → ID token
   → Netlify Function verifies the token (signature, tenant, audience)
   → writes to Timesheets via the existing app-only connection (stamped with the 365 identity)
```

It reuses your existing SharePoint setup — **no new lists or columns** (needs the
`ContractorId` column on Timesheets, plus the `Sites` and `Settings` lists, all from
the contractor app).

## One-time setup

### 1. Create the "Staff Timesheet" app registration (admin)
1. entra.microsoft.com → **App registrations → + New registration**.
2. Name `Staff Timesheet`. **Supported account types:** *Accounts in this organizational directory only* (single tenant).
3. **Redirect URI:** platform **Single-page application (SPA)** → enter your Netlify
   site URL once it exists (e.g. `https://your-staff-site.netlify.app`). You can add
   this after the first deploy (step 3).
4. Register, then copy the **Application (client) ID** → this is `STAFF_CLIENT_ID`.
   No client secret is needed for sign-in.

### 2. Deploy (Git-connected, so functions + the `jose` dependency build)
1. Put this folder in a **new GitHub repo** (e.g. `staff-timesheet`).
2. Netlify → **Add new site → Import an existing project → GitHub** → pick the repo.
3. Build command: empty. Publish directory: `.`  (functions are read from `netlify.toml`).
4. Deploy → note the site URL.

### 3. Wire sign-in to the site
- Put the site URL into the app registration's **SPA redirect URI** (step 1.3).
- Edit **`config.js`** → paste the **client id** into `clientId`. Commit (auto-deploys).

### 4. Environment variables (Netlify → Site configuration → Environment variables)
```
TENANT_ID            3efd78a4-4c46-434e-b653-4d0b65d18caa
CLIENT_ID            (the app-only "Contractor Timesheet" client id — reused for SharePoint writes)
CLIENT_SECRET        (its secret Value)
SP_SITE_ID           jdmclennan.sharepoint.com,7925a1f9-...,7d6a8277-...
LIST_ID              013ec528-5efd-4b78-a86b-6b8c148c2ff5   (Timesheets)
STAFF_CLIENT_ID      (the Staff Timesheet registration's client id)
```
Redeploy after setting them.

### 5. Test
Open the site → **Sign in with Microsoft** → you should land on your week → enter
hours → Submit → check a row appears in Timesheets with your name and `Notes` =
"Staff app · <your email>".

## Allowances (shift types, e.g. Night Shift)

Set up in the **Allowances** SharePoint list, one row per allowance per site — no redeploy:

- **Title** — the name shown on the tick box.
- **Site** — Lookup to the Sites list.
- **Frequency** — `Daily` (tick box on each day; the default if blank) or `Weekly` (one tick box for the whole week).
- **Active** — set to **No** to hide it.
- **Days** *(optional, Choice with "allow multiple selections": Monday … Sunday)* — limits a
  **Daily** allowance to those weekdays, e.g. Saturday + Sunday for a weekend allowance. Leave
  blank for every day. The tick box only shows on those days, and the server rejects it on any
  other day. Ignored for Weekly allowances.
- **ShiftName** *(optional, single line of text)* — the Teams Shifts label(s) that pre-tick this
  allowance, e.g. `Night` or `Night, Late`. A shift whose name equals the allowance's Title matches
  without it. See *Teams Shifts roster* below.

## Teams Shifts roster (rostered shift + "Shift worked")

For a site linked to a Team, each day card shows the person's **rostered shift** from Teams Shifts
and a **Shift worked** dropdown. On days not saved yet the app starts from the roster: that site,
that shift, and the allowances its name maps to (Allowances → `ShiftName`). The person can change
**Shift worked** if they worked something else — the shift's allowances swap over, anything they
ticked by hand stays — and the day shows "Different from the roster".

**The roster is fixed at the start of each pay week.** The hourly `roster-snapshot` function
(schedule in `netlify.toml`) saves each linked site's published shifts for the week into the
**RosterSnapshots** list in the first hour of the week (the day after the site's CloseDay, in the
Team's time zone). After that the app only reads the snapshot — later edits in Teams don't change
the week; staff fix differences themselves with **Shift worked**. If a week has no snapshot (first
deploy, a missed run), the first person to open the app takes it, and that becomes the week's roster.
Only **published** (shared) shifts count, on the day each shift **starts**.

**On call is an allowance, not a shift.** A Teams shift whose name matches a **Weekly** allowance
(its title, or a label in its `ShiftName` — compared ignoring case, spaces and hyphens, so
"On Call" = "Oncall") ticks that weekly allowance (until the week is first saved) and is shown on
the day as "Also rostered: Oncall (allowance)". It is never the rostered/worked shift and isn't
offered in **Shift worked**. Staff can untick or tick it themselves.

**Per-job allowances (v27):** any Allowances row with **Frequency = `Per job`** (e.g. `Call-out`
for an on-call call-out that costs extra, `Call-Out - Working Hours` for one taken while already on site during working hours)
becomes a tick on every work-order line for that site. Ticks are saved on the row as
**JobAllowances** (Timesheets, single line of text, comma-separated); if that column is missing the
week still saves and the person is told the ticks weren't kept. Rows saved before v27 used
`Callout = Yes` — still read as "Call-out".

**Contract hours (v27):** Workers → optional **ContractHours** (number, hours per week). The week
total shows "of 40 h" and warns when over. In the supervisor app a week over contract needs the
site's **SeniorApprover** as well: the supervisor's Approve marks it "awaiting senior approval"
(still Submitted, so not in payroll) and the senior approver's Approve approves it.

**Public holidays (v27):** each day shows "Public holiday — <name>" for its site, and ticks the
site's daily allowance whose name contains "Public holiday" (if there is one) on days not yet
saved. Holidays come from date.nager.at by the Sites column **HolidayRegion** (`AU-QLD`, `AU-NSW`,
`NZ-AUK`, or just `AU` / `NZ`); with no column it's guessed from the site code (BNE → AU-QLD,
SYD/WSI → AU-NSW, AKL → NZ-AUK …). Extra local days go in an optional list **PublicHolidays**
(Title = name, **HolidayDate** = date, **Region** = blank for all sites, or regions / site names).
The hourly `holiday-notes` function adds each holiday for the next 120 days to every Shifts-linked
site's Teams schedule as a day note — needs Graph **Schedule.ReadWrite.All** (Application, admin
consent) on the `CLIENT_ID` app; until then it just logs a 403.

Each saved row records **ShiftWorked** (what they say they worked) and **RosteredShift** (the
week's roster for that day, e.g. `Night 18:00–06:00 (BNE)`), so a supervisor can spot differences.

### Setup
1. **Permission (admin):** entra.microsoft.com → App registrations → the app-only app (`CLIENT_ID`)
   → API permissions → Microsoft Graph → **Application** → `Schedule.Read.All` → **Grant admin consent**.
2. **Sites list:** add **ShiftsTeamId** (single line of text); on the site's row put its Team id —
   BNE = `2d11d80c-1357-4fb5-9325-d5646295c102` (AEL AU - BNE). Leave blank for unlinked sites.
3. **Timesheets list:** add **ShiftWorked**, **RosteredShift** and **JobAllowances** (single line of text). Until they
   exist the app still saves, just without them.
4. **RosterSnapshots list:** created automatically on first run. If the app isn't allowed to create
   lists, create it yourself. Only **Data** (multiple lines of text, plain) is required — each row is
   keyed by its Title ("BNE 2026-10-28"); **Site**, **WeekStart**, **TakenAt** (single line of text)
   are filled in if present. Don't edit **Data** by hand; delete a row to re-capture that week.
5. Optional: Allowances → **ShiftName** for shift labels that differ from the allowance name.
6. If the hourly run's log (Netlify → Logs → Functions → roster-snapshot) shows Shifts refusing an
   app-only call without a user, set env **SHIFTS_ACT_AS** to a Team owner's Entra object id.

Contractors aren't in Teams, so this is staff-app only.

## Notes
- **Identity is verified server-side** — the function checks the Microsoft token's
  signature, your tenant, and audience, so a forged sign-in can't write data.
- SharePoint writes still use the **app-only** connection (the contractor app's creds),
  so staff don't each need SharePoint permissions.
- The `Worker` (person) column isn't set yet — staff identity is captured as name +
  email. Mapping to the person field can be added later.
- Approvals still live in your Power Automate flow for now.
