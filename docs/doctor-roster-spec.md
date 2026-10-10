# Doctor Roster: Research & Specification (v0.2, Haematology)

**Status:** draft for review. No code yet.
**Goal:** replace the Excel rota plus the Power Automate leave form with one web-hosted system. Leave, swaps and redeployments should update the live rota automatically.

---

## 0. Haematology: current setup and what changes (v0.2)

> This section is the most current and overrides the generic parts of §2–§7 where they disagree.

### What exists today
- **Roster (Excel):** dates run down the rows and doctors across the columns.
  - Columns: Date, Day, Consultant (of the week), then resident doctors (F1, F2, F2 on calls NHH, IMT2, IMT3, JCF), then registrars (middle grade).
  - Each cell holds free text, colour-coded:
    - **GUH** red, **RGH** blue, **NHH** green
    - sub-units: *day unit*, *LAB*, *MDCU*
    - leave and off in yellow: ANNUAL / STUDY / SICK LEAVE, OFF, day off in lieu, SpR training day
    - general medical on-call shifts: LONG DAY, NIGHTS, WARD COVER, (EARLY) TWILIGHT, W/E SD, FLEXIBLE W/C, OFF pre on-call
    - half days, e.g. "RGH AM; PM Annual Leave"
  - Weekends are blacked out except for on-call.
  - Red and purple date cells mark days when leave is discouraged because the department is short-staffed.
- **On-call shifts** come from the **trust medical rota** and are copied in by hand.
- **Leave request (MS Forms + Power Automate):** fields are:
  1. date of request
  2. requestor (free text)
  3. NHS email
  4. leave type (Annual / Study / Sick / Other)
  5. start date
  6. end date
  7. total days (typed by the doctor)
  8. doctor covering the day unit (free text)
  9. doctor covering GUH (free text)
  10. remaining AL/SL (typed by the doctor)
  11. notes

  The output is a summary with an **empty approval section** (name, date, signature), which is emailed around. Nothing updates the roster.
- **Consultants** are not on the roster apart from the consultant-of-the-week column.

### Problems this causes
Everything is retyped by hand. Names, day counts and balances are self-reported and unchecked. Approval isn't recorded. The roster doesn't change when leave is approved. Cover is arranged by free text. There's no automatic check that the sites are staffed.

### Staffing rules (configurable in Settings)
| Group | GUH | RGH | NHH |
|---|---|---|---|
| **Resident doctors** (F1/F2/IMT/JCF, interchangeable) | 1 every weekday | 1 every weekday | 1 **Tue–Fri** (not Mon) |
| **Registrars** (middle grade) | **1 mandatory** every weekday | 1 ideally | 1 ideally |

The cover check shows ✓ / ✕ for the mandatory posts, "–" (amber) for the ideal posts, and "at risk" when a *pending* request would create a gap.

### Decisions from you
| Topic | Decision |
|---|---|
| On-call shifts | Imported from the **trust medical rota** (CSV/Excel upload each rotation). They're shown with a ⇣ marker and locked against accidental edits |
| Coloured dates | Become **leave-restricted dates**: "NO AL" (red: requests are flagged to the rota lead) and "AVOID" (purple: warning to the requester) |
| Consultants | Their own **Consultants tab** in the same layout (site, clinic, MDT, LAB, SPA, on call, leave). Consultant of the week still appears on the residents' tab |
| Excel | **The website is the master copy.** After every change it rewrites an Excel file in SharePoint/OneDrive in today's layout and colours, plus a *Leave log* sheet. The file is read-only, and edits are made on the website |
| Approver | **Rota lead consultant.** For the rota lead's own leave, the clinical director (to confirm) |
| Cover | **The system suggests and the doctor confirms.** It lists colleagues who can move without breaking a mandatory post. When leave is approved, the covering doctor's cell moves automatically |
| Allowances | **NHS standard defaults, editable in Settings** (see below) |

### Default allowances (editable)
These are starting values to check with medical staffing, not confirmed figures.

| Grade | Annual leave / year | Study leave |
|---|---|---|
| Resident doctors (2016 TCS) | 27 days (32 after 5 years' NHS service), pro-rata to rotation length and LTFT % | 30 days / year. Foundation doctors' study leave follows foundation school rules |
| Registrars / SAS | 27–32 days depending on contract and service | Per contract |
| Consultants | 32 days (34 after 7 years as a consultant) | 30 days over 3 years |

Bank holidays are handled separately. "Total days" excludes weekends and bank holidays and supports half days.

### New leave form (replaces MS Forms)
| Old field | New behaviour |
|---|---|
| Date of request, requestor, email | Filled in automatically from the sign-in |
| Leave type | Annual, study, professional/exam, teaching (not from SL allowance), sick, compassionate, day off in lieu, other |
| Start / end, total days | Date pickers with AM/PM. **Days calculated automatically** |
| Remaining AL/SL | **Calculated automatically** from allowances and approved leave |
| Doctor covering (day unit / GUH) | **Suggested by the system** from the roster for each affected site, then confirmed by the doctor |
| Approval section | Electronic sign-off by the rota lead (name, role, timestamp). A PDF in the old layout is emailed to the requester, the covering doctor and the rota admin |

### Excel export
`docs/samples/haematology-roster-export-sample.xlsx` (generated by `tools/sample_excel_export.py`) shows the planned output:
- a *Residents & Registrars* sheet in your layout, with cover-check columns
- a *Leave log* sheet listing every request, status, approver and timestamp
- an *About* sheet

The real system writes it with Microsoft Graph to a SharePoint/OneDrive file after each change. This needs the trust's M365 tenant. As a fallback, the file can be downloaded on demand.

---

## 1. Research summary: what Medirota does (and what we copy)

Medirota is made by Rotamap Ltd (a sibling of CLWRota, which serves anaesthetics). It is used by around 18,000 staff in 400+ departments across about 60 NHS trusts, and it is aimed at consultant-led services.

| Medirota capability | What it means | Take into our design? |
|---|---|---|
| **Repeating patterns of work** | Job plans and rota lines are entered once and roll forward automatically | **Yes.** This is the core of our rota-line model (§4.3) |
| **Integrated leave management** | Leave is requested in the same system. Approved leave removes the person from their sessions, which then show as gaps | **Yes.** This is the main pain point today |
| **Minimum staffing levels** | Each activity or location has a required headcount. Gaps are highlighted | **Yes**, per ward/clinic and per grade |
| **Internal bank / "bidding"** | Uncovered sessions are offered out and clinicians bid via the app. The rota team allocates | **Yes**, as gap bidding (§4.6) |
| **Mobile app** | Personal rota, leave diary, leave requests, bids, and who else is on | **Yes**, as a responsive web app (PWA), not native apps |
| **Notifications** | Automatic email/SMS/push on leave approval and late changes affecting you. User picks urgent/non-urgent channels | **Yes**, email first, then Teams (§4.8) |
| **Annualised contract tracking** | Sessions planned vs. delivered against job plans (consultant PAs) | **Phase 2** (consultant job-plan tracking) |
| **Reporting / benchmarking** | Planned sessions, sessions worked, leave taken, % session cover | **Yes**, basic reports in v1 |
| **Junior doctor compliance checks** | Live checks against 2016 contract rules | **Yes** (§5) |
| **Cross-department view** | Colleagues on other Rotamap rotas visible where activities are mapped | Out of scope for now |
| **Shift swaps** | No evidence Medirota supports doctor-to-doctor swaps directly; it relies on bidding | **We go further:** formal like-for-like swaps (§4.5) |

**Other products in the same space:** Allocate/RLDatix HealthRoster & Medic Rota (with compliance and EWTD calculators), Patchwork (bank/locum and leave), and HealthRota. An open-source example is `RudolfCardinal/rota` on GitHub, a Python rota and compliance checker.

**Key guidance documents:**
- BMA / NHS Employers *Good Rostering Guide* (2018) recommends using technology for rostering and for managing the live rota, and designing rotas with fatigue in mind.
- **NHS England Minimum Standards for Annual Leave for Resident Doctors (1 Apr 2026)** requires:
  - every trust has a published resident-doctor leave policy
  - leave is approved or declined promptly *with clear reasons*
  - rotas have enough flexibility for doctors to take their full annual, study and bank holiday entitlement

  → Our system must record a **decision reason** and **time-to-decision** for every request.
- **Exception reporting reforms (in force 4 Feb 2026)** let doctors choose pay or TOIL without supervisor sign-off. Employer fines rose to £500 per doctor per week from 4 Aug 2026.
  → In v1 the system only needs to **record TOIL as a leave type**. Exception reporting itself stays in the trust's existing tool.

---

## 2. Requirements captured from you

| Topic | Your answer |
|---|---|
| Staff | Resident doctors (FY/CT/ST), Consultants, SAS / specialty doctors |
| Scale | 30–100 doctors, a few locations |
| Pain points | Day-to-day redeployment, rotations, rota line swaps, no live "who's where" view |
| Hosting / IT | **Not decided yet**, so the design must be portable (§7) |
| Leave types | Annual, study, sickness, compassionate, professional/other, **not-on-call requests** |
| Leave approval | Automatic checks, then **one approver** with a mandatory reason |
| Leave rules | **6 weeks' notice**. **Minimum staffing per grade** |
| Rota build | **Rota lines in Excel** (a rolling pattern) |
| Rota cycle | **Both**: residents by rotation (Aug / Dec / Apr changeover), consultants on a rolling job-plan rota |
| Swaps | **Like-for-like, both doctors agree, admin approves** |
| Locations | **Flat list of wards and clinics**, each with required cover |
| Roles | Rota coordinator/admin, Clinical lead / CD, Doctors, Read-only viewers |
| v1 scope | Live rota + leave, Swaps + gap bidding, Compliance checks, Notifications + calendar feed |
| Integrations | **Import from Excel** (staff list and current rota) |
| Next step | Research and spec only. Review before any code |

---

## 3. Users & permissions

| Role | Can do |
|---|---|
| **Doctor** | View the full rota and their own "My rota". Request leave and not-on-call dates. Propose and accept swaps. Bid on gaps. Subscribe to a personal calendar feed. Choose notification preferences |
| **Rota coordinator (admin)** | Everything a doctor can do, plus: edit the rota, manage rota lines and allocation, approve or decline leave, swaps and bids, redeploy people, manage staff, wards and shift codes, import from Excel, override compliance warnings (with a logged reason) |
| **Clinical lead / CD** | Read everything. Dashboards and reports. Approve escalated requests (inside 6 weeks, or below minimum staffing). Sign off compliance overrides |
| **Read-only viewer** | "Who's on now / today" board by ward, with bleep or phone numbers (ward managers, switchboard). No personal leave details, which show only as "Unavailable" |

Every change is audit-logged: who, what, when, before and after, and why.

---

## 4. Functional specification

### 4.1 Core data model

```
Doctor ──< Assignment >── Shift (date, shift code, ward)
  │  grade, contract type (2016 TCS / SAS / Consultant), FTE %,
  │  competencies, rota line, leave allowances
  │
  ├──< LeaveRequest  (type, dates, half-day AM/PM, status, approver, reason, decided_at)
  ├──< SwapRequest   (shift A ⇄ shift B, doctor A, doctor B, status)
  └──< Bid           (on an open Gap)

RotaTemplate ──< RotaLine ──< PatternDay (day N → shift code + ward)
Ward/Clinic  ──< StaffingRequirement (shift code, grade, min count, days of week)
ShiftCode    (code, label, start, end, crosses midnight, hours, is_on_call, is_leave, colour)
RotaPeriod   (start, end, template, published_at)
AuditLog
```

### 4.2 Shift and absence codes (pre-loaded, fully editable)

Every working code is tied to a **ward or clinic**. Leave codes are not.

| Code | Label | Typical times | Notes |
|---|---|---|---|
| **D** | Normal day | 09:00–17:00 | Ward-based |
| **LD** | Long day | 08:00–20:30 | Over 10 hours, so counts as a "long shift" for the rules |
| **E** | Evening / twilight | 14:00–22:00 | |
| **N** | Night | 20:30–09:00 | Crosses midnight. Night-shift rules apply |
| **OC** | On call (resident) | as defined | e.g. weekend on call 08:00–20:30 |
| **NROC** | Non-resident on call | e.g. 17:00–09:00 | From home. Separate NROC rules apply. Common for consultants and some registrars |
| **WE** | Weekend day | 08:00–20:30 | Counts toward weekend frequency |
| **CL** | Clinic | AM / PM session | Clinic-based (consultant/SAS sessions) |
| **TH** | Theatre / list | AM / PM session | |
| **TEACH** | Teaching / regional training day | 09:00–17:00 | Protected, counted as working time |
| **ADMIN / SPA** | Admin / SPA session | AM / PM | Consultant and SAS job plan sessions |
| **OFF** | Day off / zero day | none | Rostered rest |
| **POST** | Post-nights rest | none | Auto-inserted after a run of nights |
| **AL** | Annual leave | none | Deducted from allowance |
| **SL** | Study leave | none | Deducted from the study leave allowance |
| **PL** | Professional leave | none | Exams, courses, conferences, interviews |
| **SICK** | Sickness | none | Entered by admin or the doctor. Triggers a cover gap |
| **COMP** | Compassionate leave | none | |
| **PAR** | Parental leave | none | |
| **TOIL** | Time off in lieu | none | From exception reports or extra hours |
| **BH** | Bank holiday (in lieu) | none | |
| **NOC** | *Not-on-call request* | none | A **pre-rota preference**, not leave. Used when rota lines are allocated |

### 4.3 Building the rota: rota lines (matches how you work in Excel)

1. The admin builds a **rota template**, e.g. a 10-week, 10-line pattern. Each line is a sequence of shift codes plus a ward for each day. Templates can be **imported from the existing Excel sheet** (§4.9).
2. For each **rota period**, the admin assigns each doctor to a line, with a start offset for rotations.
   - Residents use rotation-length periods (Aug / Dec / Apr).
   - Consultants use a rolling job-plan pattern rolled forward monthly.
3. When allocating lines, the system shows each doctor's pre-submitted **NOC requests and booked leave**, and suggests the best-fitting line.
4. **Compliance check runs on the template** (§5) *before* publication. Any line that breaks a rule is flagged.
5. **Publish.** The rota should be published at least **6 weeks before the period starts** (Code of Practice). The system warns if it is late.
6. After publication, the **live rota** is the source of truth. Every change (leave, swap, redeployment, sickness) is an *override* on top of the pattern. This means the original pattern is never lost and changes are reportable.

### 4.4 Leave workflow

```
Doctor submits request (type, dates, AM/PM, note)
        │
        ▼
Automatic checks ─────────────────────────────────────────────┐
  • Allowance remaining (pro-rata by FTE and start date)      │
  • Notice ≥ 6 weeks?               no → flag "short notice"  │
  • Minimum staffing per grade, per day, per ward after       │
    removing this doctor            breach → flag "below min" │
  • Clashes: on a night block / on call / existing leave      │
  • Shows doctor & approver what shifts they'd lose           │
        │                                                      │
        ▼                                                      ▼
Rota coordinator approves / declines            Flags → escalate to Clinical lead
  (reason is MANDATORY on decline)
        │
        ▼ approved
Rota auto-updates: shifts become leave, displaced shifts become GAPS
→ gap goes to cover workflow (§4.6) → notifications sent (§4.8)
```

- **Sickness** is entered by the admin (or the doctor via "report sick"). It applies immediately with no approval step and creates gaps.
- **Not-on-call requests** have a separate deadline before rota allocation, e.g. "submit by 12 weeks before changeover".
- Leave balances are shown on the doctor's dashboard: entitlement, taken, booked, remaining.
- **Metrics for the NHSE 2026 standards:** median time-to-decision, % declined (with reasons), and % of entitlement taken per doctor (to flag people who can't get their leave).

### 4.5 Swaps (like-for-like, admin approves)

1. Doctor A picks a shift and selects "Offer swap". They choose a specific colleague or "anyone eligible".
2. The system only lists **eligible** colleagues: same grade band, required competencies, not on leave, and the swap would not break compliance for *either* doctor.
3. Doctor B accepts, then the coordinator approves in one click. The system shows a compliance check for both doctors.
4. The live rota updates, both doctors are notified, and the swap is audit-logged.
5. Supported swap types:
   - shift ⇄ shift
   - shift → "give away" (becomes a gap if nobody takes it)
   - whole **rota line swap** for a period, for the rotation case

### 4.6 Gaps, redeployment and bidding (the "moving people" problem)

- **Gap board:** every ward/clinic × shift × grade where cover is below the minimum is shown in red, with the reason (leave, sickness, vacancy).
- **Redeploy (admin):** drag a doctor from an over-covered ward to an under-covered one *on the same day*. The system shows the effect on both wards and logs it as a redeployment, not a rota change.
- **Offer out (bidding):** the admin publishes a gap. Eligible doctors are notified and bid ("I can do it"). The admin picks a bidder. A compliance check runs for that doctor before assignment. This is Medirota's "bank bids" model, kept internal only in v1.
- **Who's where now:** a live board showing, for today or any date, each ward → who is on → grade → contact/bleep. It is also the read-only viewer's landing page.

### 4.7 Views

- **Department grid:** doctors × dates, colour-coded by shift code, filterable by grade, ward and team. Weekly, 4-weekly and full-rotation views.
- **Ward view:** wards × dates, showing headcount against the minimum.
- **My rota:** personal list and calendar, leave balance, pending requests.
- **Requests inbox (admin):** leave, swaps and bids in one queue, sortable by urgency, with SLA timers.
- **Print / PDF export** of the rota for notice boards.

### 4.8 Notifications & calendar

- Email in v1. Teams via webhook as an optional add-on. Each doctor picks urgent and non-urgent channels, as Medirota does.
- Triggers:
  - leave decided
  - swap proposed, accepted or approved
  - gap offered that you're eligible for
  - change to *your* shifts within 14 days ("late change")
  - rota published
- **Personal iCal feed** (secret URL per doctor) that phones and Outlook subscribe to. It updates automatically.

### 4.9 Excel import (v1 integration)

- **Staff list:** name, email, grade, contract, FTE, start/end date, rota line, AL/SL entitlement.
- **Rota template / current rota:** a grid with doctors (or line numbers) down the side, dates across the top, and shift codes in the cells.
  - Import maps your existing cell codes to system shift codes. Unknown codes are listed for you to map once.
  - The importer previews changes before saving and reports errors by row and cell.
- We'll provide a downloadable **template .xlsx**. A sample of your *real* sheet (anonymised) is needed to design the mapping.

---

## 5. Compliance rules engine

Rules are **configurable data, not hard-coded**, and attached per contract type. They run on the template, on every live change, and on swap and bid approval.

Breaches show as:
- 🔴 **block**, which needs a clinical lead override with a reason
- 🟠 **warn**

> ⚠️ The values below reflect the **2016 TCS for resident doctors in England (as amended by the 2019 refresh)**, as best understood. **They must be verified against the current TCS Schedule 3 and your trust's guardian of safe working / medical staffing before go-live.** The rules are editable, so corrections are configuration rather than code changes.

**Resident doctors (2016 TCS):**

| Rule | Default | Level |
|---|---|---|
| Average weekly hours over the reference period | ≤ 48h (≤ 56h with an EWTD opt-out) | 🔴 |
| Hours in any 168-hour period | ≤ 72h | 🔴 |
| Maximum shift length | ≤ 13h | 🔴 |
| Rest between shifts | ≥ 11h | 🔴 |
| Consecutive night shifts | ≤ 4, then ≥ 46h rest | 🔴 |
| Consecutive long shifts (> 10h) | ≤ 4, then ≥ 48h rest | 🔴 |
| Consecutive shifts | ≤ 7, then ≥ 48h rest | 🔴 |
| Weekend frequency | no more than 1 in 2. Warn above 1 in 3 | 🔴 / 🟠 |
| Non-resident on-call | ≤ 3 consecutive days (4 at weekends). ≥ 8h rest per 24h with ≥ 5h continuous between 22:00–07:00 | 🟠 |
| Rota published | ≥ 6 weeks before start | 🟠 |
| Leave notice | ≥ 6 weeks | 🟠 (escalates) |

**Consultants and SAS:**
- EWTD 48-hour average and 11-hour daily rest.
- Job-plan tracking (planned vs. delivered PAs) comes in Phase 2.

---

## 6. Phasing

| Phase | Contents |
|---|---|
| **v1 (MVP)** | Staff and wards setup, shift codes, Excel import, rota lines → live rota, leave workflow, swaps, gaps/redeploy/bidding, compliance engine, email notifications, iCal feed, who's-where board, audit log, basic reports |
| **v2** | Teams notifications, consultant job-plan / PA tracking, payroll/ESR export, rota monitoring (hours actually worked), auto-suggested line allocation, multi-department support |
| **v3** | Auto-generated rotas (constraint solver), locum/bank integration, SSO-everywhere, analytics dashboards |

---

## 7. Architecture (portable, since hosting isn't decided yet)

- **Web app:** TypeScript (Next.js / React), responsive and installable as a PWA on phones. No app store needed.
- **Database:** PostgreSQL. This works on any host: Azure (for a trust tenant), Supabase, or on-prem.
- **Authentication:** pluggable.
  - **Option A:** Microsoft Entra ID / NHSmail single sign-on if a trust tenant is available.
  - **Option B:** email magic-link login for a standalone pilot.
- **Background jobs:** notifications, iCal regeneration, nightly compliance re-check.
- **Hosting choices:**

  | Option | Pros | Cons |
  |---|---|---|
  | Trust Azure tenant | IG-friendly, SSO | Needs trust IT |
  | Supabase / Vercel (UK/EU region) | Fastest pilot | Needs DPIA |
  | Self-hosted Docker | Full control | Someone has to run it |

- **Information governance:** the system holds staff personal data and sickness information (special category health data). A **DPIA** will be needed before real data goes in. Use the UK/EU region, encrypt data, keep audit logs, and keep sickness reasons out of rota views (they show only as "Unavailable").

---

## 8. Open questions for you

1. **Excel sample:** can you share an anonymised copy of the current rota sheet and the Power Automate leave form fields? These drive the importer and the leave form.
2. **Grades and minimums:** what grade bands do you staff by (e.g. FY1 / SHO-level / Registrar / Consultant), and what is the minimum per ward per shift?
3. **Entitlements:** what are the annual and study leave entitlements by grade and seniority? Do you use half-days?
4. **Not-on-call deadline:** how many weeks before changeover must requests be in?
5. **Who approves:** one named coordinator, or a shared admin team? Who covers when they are away?
6. **Competencies:** are there any shifts only certain doctors can do (e.g. ALS-certified, procedure-trained, registrar-on-call eligible)?
7. **Hosting:** can you check with your trust's IT/IG whether an Azure/NHSmail tenant is available, or whether a standalone pilot is acceptable?
8. **Pilot:** which ward or team, and roughly when (ideally ahead of a rotation changeover)?
9. **Consultant job plans:** do consultants need PA tracking in v1, or is the rota grid enough to start?
10. **Payroll:** is the hours and enhancements export for ESR needed soon, or later?

---

## Sources

- Rotamap Medirota product page: https://rotamap.net/medirota
- Rotamap mobile apps: https://rotamap.net/apps · Messaging docs: https://www.rotamap.net/docs/messages
- Medirota app (App Store): https://apps.apple.com/gb/app/medirota/id967224308 · (Google Play): https://play.google.com/store/apps/details?id=com.erota.medirota
- G-Cloud listing (Rotamap): https://www.applytosupply.digitalmarketplace.service.gov.uk/g-cloud-14/services/534640821858651
- BMA / NHS Employers Good Rostering Guide (2018): https://www.bma.org.uk/media/1979/bma-nhse-good-rostering-guidance-may2018.pdf
- NHS England Minimum standards for annual leave for resident doctors (2026): https://www.stgeorges.nhs.uk/wp-content/uploads/2026/04/NHS-England-Minimum-standards-for-annual-leave-for-resident-doctors.pdf
- BMA exception reporting in England (2026 reforms): https://www.bma.org.uk/our-campaigns/resident-doctor-campaigns/pay-in-england/exception-reporting-in-england-summary-and-faqs
- Open-source rota checker: https://github.com/RudolfCardinal/rota
- HealthRota evaluation: https://transform.england.nhs.uk/documents/145/HealthRota_Evaluation_-_Future_Healthcare_Journal.pdf
