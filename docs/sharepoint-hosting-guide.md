# Putting the Haematology Rota on SharePoint

## Can I just share the prototype via OneDrive or SharePoint?

- **OneDrive:** you can share the `.html` file, but people download it and open their own copy. Each copy keeps its own data, so requests and approvals never reach anyone else. This is fine for letting colleagues click through the prototype, but not for running the rota.
- **SharePoint:** a document library stores `.html` files but does not open them as web pages. SharePoint Online blocks custom scripts by default, so the file downloads instead. Workarounds that rename the file to `.aspx` are blocked or reset by Microsoft. A plain upload therefore won't work.

## What works: Option A, SharePoint + Lists + Power Apps + Power Automate

Everything stays inside your NHS Wales Microsoft 365, and people sign in with their normal account. The two sign-ins come from two SharePoint groups:

| Group | Who | Permission |
|---|---|---|
| Rota Authorisers | The 3 rota coordinators | Edit everything, approve leave |
| Rota Viewers | Doctors, ward managers, switchboard | Read only (optionally they can send leave requests) |

### Step by step
1. **Get the site.** SharePoint start page → **+ Create site** → **Team site** → "Haematology Rota" → Private → add the other two authorisers as owners. If you can't create sites, ask IT (email below).
2. **Set the two sign-ins.** Settings (gear) → **Site permissions** → **Advanced permissions settings** → **Create group**:
   - "Rota Authorisers" with **Edit** (or Full control): add the 3 coordinators.
   - "Rota Viewers" with **Read**: add everyone else.

   Remove "Everyone except external users" if it appears.
3. **Create the lists from Excel.** Site → **+ New** → **List** → **From Excel** → upload `templates/staff-list-template.xlsx` → name it "Doctors". Then:
   - Repeat for "Roster" (Date, Doctor, Site/shift, Source), "Changeovers" and "Leave allowances".
   - Create "Leave requests" as a blank list with these columns: Doctor, Type (choice), Start, End, Part of day, Days, Status (Pending/Approved/Declined), Cover, Reason.
4. **(Optional) Let viewers send requests.** Leave requests → **List settings** → **Permissions for this list** → **Stop inheriting** → give "Rota Viewers" **Contribute** on this list only. Then **Advanced settings** → Item-level permissions: *Read all items*; *Create and edit only their own*.
5. **First version, no coding: list views.** Roster → **+ Add view** "Who is where" → Group by Date, then Site → Filter Date = [Today]. This gives "GUH: Name (grade), …". Add a Calendar view for leave and a "Pending" view of requests.
6. **Screens in Power Apps.** make.powerapps.com → **Create** → **Start with data** → SharePoint → Roster list. Add the roster grid, list view, leave form, and the EDT/SDT and leave calculations as screens. **Publish** → **Share** with both groups.
7. **Automatic approvals in Power Automate.** make.powerautomate.com → **Automated cloud flow** → trigger *When an item is created* (Leave requests) → *Start and wait for an approval* (the 3 authorisers, first to respond).
   - **Approved:** set Status = Approved, add the leave days to Roster, and email the doctor.
   - **Declined:** set Status = Declined and email the reason.
8. **Home page.** Site home → **Edit** → add the **Power Apps** web part, a **List** web part ("Who is where") and **Quick links**. Then republish.
9. **Share the link** with the department, or add it as a Teams tab (channel → **+** → SharePoint).
10. **Weekly on-calls and Excel.**
    - **On-calls:** paste rows from the on-call template into Roster using **Edit in grid view**.
    - **Excel:** Roster → **Export** → **Export to Excel**, or set up a scheduled flow that emails it every Monday.

### Email to IT
> Please could you create a private SharePoint team site called "Haematology Rota" with me as owner, and confirm that Power Apps and Power Automate (standard connectors only) are enabled for our accounts? It will hold staff names, rota and leave, no patient data. Do we need an information governance or DPIA check?

### Notes
- **Licensing:** SharePoint with standard connectors is normally covered by an enterprise Microsoft 365 licence. Your health board's IT controls whether Power Apps and Power Automate are switched on.
- **What needs building:** the cover checks, cover suggestions and EDT/SDT maths have to be rebuilt as Power Apps formulas. The first version can be lists only (step 5), with the clever parts added later.

## Option B: IT hosts the web app
IT (or DHCW) hosts the app built in this repo, for example on Azure with NHS Wales sign-in. A SharePoint page then shows it through the **Embed** web part, with the site address added to the allowed list in *HTML Field Security*.

This keeps every feature exactly as designed, but needs IT approval and set-up.
