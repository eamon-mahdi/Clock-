# Haematology Rota

A rota and leave system for the haematology department (GUH · RGH · NHH). It replaces the Excel roster, the Microsoft Forms leave request and the Power Automate emails.

- **Prototype:** `dist/haematology-rota.html`. Open it in a browser. Fictional data only.
- **How it works:** [docs/prototype-guide.md](docs/prototype-guide.md)
- **Research and specification:** [docs/doctor-roster-spec.md](docs/doctor-roster-spec.md)
- **Import templates:** [templates/](templates/) (on-call rota, weekly plan, staff list)

Development: `node tools/build.js` builds the page and templates. `node --test tests/rota.test.js` runs the tests. No dependencies.
