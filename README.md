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

## Notes
- **Identity is verified server-side** — the function checks the Microsoft token's
  signature, your tenant, and audience, so a forged sign-in can't write data.
- SharePoint writes still use the **app-only** connection (the contractor app's creds),
  so staff don't each need SharePoint permissions.
- The `Worker` (person) column isn't set yet — staff identity is captured as name +
  email. Mapping to the person field can be added later.
- Approvals still live in your Power Automate flow for now.
