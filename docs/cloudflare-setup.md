# Running the rota on Cloudflare

The website runs on Cloudflare's free plan: a Worker (the server), static files (the pages) and a D1 database (the rota data). There are no patient details in it. It holds doctors' names, grades, work emails, where they work and their leave.

## 1. Deploy

Cloudflare is connected to this GitHub repository. Each push to the production branch deploys automatically. The build runs `npx wrangler deploy`, which also:

- runs `node tools/build.js` to make the `public/` folder
- creates the D1 database (`haematology-rota`) the first time and connects it as `DB`

In the Cloudflare dashboard (Workers & Pages → your Worker → Settings → Build), check:

| Setting | Value |
|---|---|
| Production branch | the branch you want live (for example `main`) |
| Build command | *(leave empty)* |
| Deploy command | `npx wrangler deploy` |
| Root directory | `/` |

The Worker name in `wrangler.jsonc` (`"name"`) must match the Worker name in the dashboard. If the build log says the name doesn't match, change `"name"` in `wrangler.jsonc` to the dashboard name.

## 2. Set the authoriser password (do this straight away)

1. Open the website. A yellow bar says no password has been set.
2. Click **Set it now**. Type the authorisers' names (one per line) and choose the password.
3. Share the password with the other authorisers in person or by phone. Never email it.

Until this is done, anyone with the link could set the password. To prevent that, add a **setup code** first: in the dashboard go to your Worker → Settings → Variables and Secrets → Add → type *Secret*, name `SETUP_CODE`, any value. The setup screen will then ask for it.

The password is stored only as a salted PBKDF2 hash. Nobody can read it back. If it is forgotten, delete the row in the `auth` table (D1 → haematology-rota → Console: `DELETE FROM auth;`) and set a new one.

Good password: long and unguessable, for example 3 or 4 unrelated words mixed with a haematology term only the authorisers would think of. Avoid ward names, the department name, or anything shown on the site.

## 3. Start for real

The first visit loads **example data with fictional doctors** so you can try everything. When ready: sign in → **Set up → Settings → Start fresh**. This keeps the rules, authoriser names and alert emails. Then:

1. **Set up → Data file:** download the rota data file, fill in the Doctors tab (with NHS emails: leave requests must use the same email), Consultant of the week, Leave restrictions and On-calls, and upload it.
2. **Each week:** paste the trust on-calls into the On-calls tab and upload the file again.
3. **Set up → Master sheet / People & leave:** small changes can also be made by hand there.
4. **Set up → Rotations:** rotation dates and changeover bands.
5. **Rules:** check the numbers and the text.

## 4. Emails (optional)

Without this, decisions are shown in the app ("Check my request") only. To send emails:

1. Create a free account at resend.com and verify a sending domain you control.
2. In the Worker's Variables and Secrets add:
   - `RESEND_API_KEY` (Secret): the API key from Resend
   - `MAIL_FROM` (Text): for example `Haematology Rota <rota@your-domain>`
   - `SITE_URL` (Text, optional): the website address, added to the end of each email
3. Sign in → **Set up → Settings → Alert emails**: who is emailed about new requests and sickness.

Emails only go to addresses already in the rota (doctors, requesters and alert emails).

## 5. Before sharing the link

- Tell your Information Governance / Data Protection Officer: the site holds staff names, work emails, rota and leave data (no patient data, no sickness reasons: sickness shows only as "Unavailable").
- Viewers see names, grades and where people work. They do not see emails, notes, reasons or the change log.
- Search engines are asked not to index the site.

## Backups

The rota lives in the Cloudflare D1 database `haematology-rota` (not in a file). Copies are kept automatically:

- **After every change:** a compressed copy (the last 500 changes).
- **Every night** (01:15 UTC, a Cloudflare Cron Trigger set in `wrangler.jsonc`): a daily copy, kept for 2 years.
- **Cloudflare's own restore points:** D1 Time Travel keeps the last 30 days, in case the database itself has a problem.

Authorisers can see and use these under **Set up → Backups**: download any copy, or put one back (the restore is itself logged and backed up, so it can be undone). The same page downloads a full **Excel backup**. Its Doctors, Consultant of the week and Leave restrictions tabs can be imported again. It also downloads an exact **JSON copy**, which can be uploaded there to restore everything.

For a full database export from your computer: `npx wrangler d1 export haematology-rota --remote --output backup.sql`.

## Setting the rota up from Excel

**Set up → Data file** downloads the **rota data file**, one Excel workbook filled in with what is on the rota now. Its tabs:

- **Doctors:** name, grade, category, start and finish dates, whose column they take over, usual Mon–Fri locations and leave allowances.
- **Consultant of the week**
- **Leave restrictions**
- **On-calls:** the trust on-calls (Date, Doctor, Shift).
- **Rotation dates** and **Codes:** reference only.

Edit it in Excel and upload it on the same page. Doctors are matched by name, so nobody is deleted. On-calls already on the rota for the dates in the On-calls tab are replaced. The Backups page downloads the same file with extra read-only tabs. A CSV holds only one tab (for example the Doctors tab); the Excel file is the one file that holds everything.
