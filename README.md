# Haematology Rota

A rota and leave website for the haematology department (GUH · RGH · NHH). It replaces the Excel roster, the Microsoft Forms leave request and the Power Automate emails.

- **Live website:** hosted on Cloudflare (Workers + D1 database). See [docs/cloudflare-setup.md](docs/cloudflare-setup.md).
- **Who can do what:** anyone with the link can view the rota, send a leave request or report sickness, and check their own requests (name + NHS email). The rota authorisers sign in with one shared password and can change everything.
- **Offline demo:** `dist/haematology-rota.html` works without a server (fictional data, saved in the browser only).
- **Research and specification:** [docs/doctor-roster-spec.md](docs/doctor-roster-spec.md)
- **Import templates:** [templates/](templates/) (on-call rota, weekly plan, staff list)

## Layout

| Path | What it is |
|---|---|
| `app/` | The website (plain HTML and JavaScript, no framework) |
| `app/js/rules.js` | Cover rules, leave and staffing checks, EDT/SDT, entitlements |
| `app/js/store.js` | Data, leave requests, authoriser actions, what viewers can see |
| `worker/` | The server: JSON API on Cloudflare with a D1 database |
| `public/` | Built website that Cloudflare serves (made by `npm run build`) |
| `wrangler.jsonc` | Cloudflare settings |

## Development

```
npm install
npm test          # unit tests
npm run dev       # local server with a local database at http://localhost:8787
npm run deploy    # deploy from your computer (Cloudflare also deploys on every push)
```
