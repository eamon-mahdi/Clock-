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

1. **Set up → People:** add doctors (or import the staff list template). Add their NHS email: leave requests must use the same email.
2. **Set up → Weekly plan:** each doctor's usual site per weekday (or import the weekly plan template).
3. **Set up → Import:** each week, import the trust on-call rota using the on-call template.
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

Download an Excel copy from the Rota page (Whole rotation → Download Excel). For a full database backup: `npx wrangler d1 export haematology-rota --remote --output backup.sql`.
