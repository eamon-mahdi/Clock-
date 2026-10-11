# Haematology Rota prototype: how it works

The prototype is a working web app with **fictional people**. It runs entirely in the browser and saves changes in that browser only. Nothing is sent anywhere.

- **Open it:** use the published link, or open `dist/haematology-rota.html` in Chrome or Edge.
- **Switch person:** use the **"Viewing as"** menu at the top to act as a rota coordinator, a doctor or a ward viewer.

## The weekly cycle

| Step | Who | What happens |
|---|---|---|
| 1. Set up once per rotation | Coordinator | **Set up → People**: dates in post, grade, leave allowance. **Set up → Weekly plan**: each doctor's usual site per weekday (A/B alternate weeks allowed) |
| 2. Import on-calls weekly | Coordinator | Put the trust on-call rota into the on-call template and upload it under **Set up → Import**. Earlier imports for those dates are replaced |
| 3. Roster builds itself | Automatic | Weekly plan, then on-calls on top, then authorised leave on top. Coordinator changes override any cell and are logged |
| 4. Request leave / report sickness | Doctor | Days, balance, notice, restricted dates and on-call clashes are worked out. The app suggests cover that leaves no gap |
| 5. Authorise | One of the 3 coordinators | Only now does the leave reach the roster, and the chosen cover move is applied. Declining needs a reason. The doctor is notified |
| 6. Excel | Anyone | **Roster → Download Excel**: your familiar layout plus a Leave log and Change log |

**Sickness** goes straight onto the roster (shown only as "Sick"). The coordinators are notified, and it is never taken from annual leave.

## Cover rules (Mon–Fri, not bank holidays)

| Post | Days | Counts as cover |
|---|---|---|
| Resident at GUH | Mon–Fri | Resident on `GUH` |
| Resident at RGH | Mon–Fri | Resident on `RGH`, `RGH - DAY UNIT` or `RGH MDCU` |
| Resident at NHH | Tue–Fri | Resident on `NHH` or `NHH - DAY UNIT` |
| Middle grade, GUH ward (mandatory) | Mon–Fri | Middle grade on `GUH` |
| RGH day unit | Mon–Fri | Middle grade on any RGH code. A resident only = amber exception |
| NHH day unit | Tue–Fri | Middle grade on any NHH code. A resident only = amber exception |

`GUH LAB` is separate and never counts as cover.

The cover check marks each post as:

| Mark | Meaning |
|---|---|
| ✓ | Covered |
| res | Resident-only exception |
| ! | At risk: a pending leave request would open a gap |
| ✕ | Gap |

## Import formats (templates in `templates/`)

### On-call rota: `oncall-import-template.xlsx`
Use either layout. The importer detects which one it is.

- **List** (sheet "On-call (list)"): columns `Date | Doctor | Shift | Notes`, one row per shift.
- **Grid** (sheet "On-call (grid)"): `Date | Day | <one column per doctor>`, with shift codes in the cells and blanks for no on-call. This is the same shape as the roster.

For both layouts:
- **Dates:** dd/mm/yyyy, or real Excel dates.
- **Shift codes:**
  - `LONG DAY`
  - `NIGHTS`
  - `TWILIGHT`
  - `EARLY TWILIGHT`
  - `WARD COVER`
  - `W/E SD`
  - `FLEXIBLE W/C`
  - `OFF PRE ON-CALL`
  - `POST NIGHTS`

  Common variants such as "Night", "LD" or "off pre-on-call" are recognised.
- **Names** are matched by:
  - the exact name
  - a remembered spelling
  - surname plus initial (so "MALIK S (IMT)" and "Dr Sara Malik" both match "S. Malik")

  New spellings are matched once in the preview and remembered.
- **Re-importing** a file replaces the imported on-calls for the dates it covers, so a corrected week can simply be uploaded again.

### Weekly plan: `weekly-plan-template.xlsx`
Columns: `Doctor | Week (A or B) | Mon | Tue | Wed | Thu | Fri | Applies from`.

Site codes:
- `GUH`, `GUH LAB`
- `RGH`, `RGH - DAY UNIT`, `RGH MDCU`
- `NHH`, `NHH - DAY UNIT`
- `OFF`, `TEACHING`, `SpR TRAINING DAY`

### Staff list: `staff-list-template.xlsx`
Columns: `Name | Grade | Group (Resident or Middle grade) | Email | In post from | In post to | Annual leave days | Study leave days | Other spellings`.

Existing doctors are updated by name.

## Prototype vs. live version
| | Prototype | Live version |
|---|---|---|
| Sign-in | "Viewing as" menu | NHS email sign-in. 3 named coordinators |
| Data | This browser only | Secure server, shared by everyone |
| Permissions | Checked in the app | Enforced by the server |
| Notifications | In-app list | Email (plus the in-app list) |
| Excel | Download | Download, or weekly email |

## For developers
- `app/`: source.
  - `index.html`: page shell and styles.
  - `js/xlsx.js`: dependency-free .xlsx reader and writer.
  - `js/rules.js`: dates, codes, cell resolution, cover rules, leave days, suggestions.
  - `js/store.js`: state, demo data, permission-checked actions, audit log.
  - `js/excel.js`: templates, import parsing, roster export.
  - `js/ui.js`: screens.
- `node tools/build.js`: builds `dist/haematology-rota.html` (one self-contained file) and the templates.
- `node --test tests/rota.test.js`: unit tests.

No external packages are needed.
