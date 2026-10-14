# Ennovantiv Books

A web and mobile front end for the `Bookkeeping 2026.xlsx` workbook on the Ennovantiv SharePoint. You can log time, scan receipts, enter expenses, track mileage, add projects, and print invoices and expense reports. The workbook stays on SharePoint, so you can keep editing it directly in Excel at any time. **The Excel file is the only place data is stored.** The app writes rows straight into the workbook's existing Excel Tables, so your formulas, Calendar, Weekly, Projects and Invoice Creator sheets keep working as they do today.

```
Phone / laptop browser ──sign-in──▶ Microsoft 365 (your tenant)
        │                                   │
        └──── Microsoft Graph Excel API ───▶ Bookkeeping 2026.xlsx (SharePoint "Hub")
GitHub Pages only serves the app's code. It never sees your data.
```

| Screen | Writes to (table → sheet) |
|---|---|
| Time | `Tasks` → Time |
| Expenses + receipt scan | `Expenses` → Expenses, photo → `Receipts/2026/` next to the workbook |
| Mileage | `Mileage` → Mileage |
| Invoices | reads `Projects`, `Disbursements`; records to `Invoices`, flags `Tasks`/`Disbursements` as Invoiced |
| Projects → Add a project | new row in `Projects` → Projects (editing and deleting projects stays in Excel) |
| Expense report | reads `Expenses`, `Mileage`, per diem and km rates on the Mileage sheet; *Mark as reimbursed* sets `Expenses[Paid]` and `Mileage[PD Paid]` |

Read-only: `Clients`, `Workers`, `Categories`, `Status`, `Table12` (project types). Drop-downs come from these, so anything you add in Excel appears in the app after a refresh (↻).

## Security model

Anyone with an Ennovantiv Microsoft 365 account can use the app; nobody else can.

1. **Microsoft single sign-on, Ennovantiv tenant only.** The app is registered for *this organization only*, so Microsoft rejects sign-ins from any other organization or a personal account. Your normal sign-in rules (MFA etc.) apply.
2. **Ennovantiv addresses only.** The app also checks that the signed-in address ends in `@ennovantiv.com` (`allowedDomains` in `js/config.js`). That keeps out guest accounts invited into your tenant from other companies.
3. **SharePoint permissions.** The app acts *as the signed-in person*, so they also need edit access to the workbook on the SharePoint site (normally true for members of the site).
4. *(Optional)* To limit it to a specific Entra group instead of the whole organization, set **Assignment required = Yes** on the app and assign that group (step 2.4 below). `allowedUsers` in `js/config.js` can also list named people.

Nothing secret is stored in the repo. The Client ID is a public identifier, and the workbook link only works for people who already have access to it.

## Setup (about 20 minutes, one time)

### 1. Put the code on GitHub Pages
1. Create a repository, e.g. `ennovantiv-books`. A public repo is fine, because no data lives in it. GitHub Pages on a *private* repo needs a paid GitHub plan.
2. Upload the contents of this folder to the repository root.
3. Go to **Settings → Pages → Build and deployment**, choose *Deploy from a branch*, select `main` and `/ (root)`, then save.
4. Note the site URL, e.g. `https://YOUR-GITHUB-NAME.github.io/ennovantiv-books/`.

### 2. Register the app with Microsoft (Entra ID)
1. Go to <https://entra.microsoft.com> → **Applications → App registrations → New registration**.
   - Name: `Ennovantiv Books`
   - Supported account types: **Accounts in this organizational directory only**
   - Redirect URI: platform **Single-page application (SPA)**, URL = your Pages URL from step 1 (with the trailing `/`)
2. On the app's **Overview**, copy the **Application (client) ID** and **Directory (tenant) ID**.
3. Go to **API permissions → Add → Microsoft Graph → Delegated**, add `Files.ReadWrite.All` (`User.Read` is already there), then **Grant admin consent**.
4. *(Optional, only to restrict to a group)* **Enterprise applications → Ennovantiv Books → Properties → Assignment required = Yes**, then add the group under **Users and groups**. Leave it at *No* to let everyone in Ennovantiv sign in.

### 3. Configure
Edit `js/config.js`: paste `clientId` and `tenantId`, and check that `allowedDomains` is `['ennovantiv.com']`. Commit the change.

### 4. Connect the workbook
Open the site and sign in. On **Settings** (gear icon on phones), paste the workbook's SharePoint link, then save. To get the link, open the file in SharePoint or Excel for the web and use **Share → Copy link**; a link for *People in Ennovantiv* is fine. Then turn on **Settings → Invoice tracking** before the next billing day.

### 5. First test
Log a small time entry and check that it appears at the bottom of the Tasks table in Excel, with the `Short` column calculated. Then open it in the app and **Delete row**.

### 6. Put it on your phone
Open the site in Safari (iPhone) or Chrome (Android), then choose **Share → Add to Home Screen**. It opens like an app.

### Optional
- **Logo:** save your logo as `assets/logo.png` (light-background version). It's used in the header and on invoices. Until then, a text wordmark is shown.
- **Receipt links in Excel:** add a column named `Receipt` to the Expenses table. The app will then fill it with the link to the uploaded photo.
- **Invoice tracking (recommended, do this before the next billing day):** in **Settings → Invoice tracking**, click *Add Invoice column*. Excel adds an `Invoice` column to the Tasks and Disbursements tables. On the Time sheet, the `x` in J187 and the `=SUM(Projects[This Invoice])` in K1 move one column right, and the formula keeps working. From then on, recording an invoice writes its number (e.g. `WF2611`) on every included row. That's the record of what was billed when, which can't be reconstructed later because time can become billable long after it was worked. Invoices recorded before tracking are left as they are.
- **Try it without signing in:** open the site with `?demo` on the end (e.g. `…/ennovantiv-books/?demo`). It uses made-up sample data and saves nothing.
- **Local preview:** `python3 -m http.server 8765`, then open <http://localhost:8765/?demo>. To sign in locally, also add `http://localhost:8765/` as a SPA redirect URI.

## How it behaves

- **Adds** go through the table's `rows/add`, so the table grows and calculated columns fill automatically. If Excel ever leaves a calculated cell blank, the app copies the formula from the row above.
- **Edits** write only the non-formula cells of that one row.
- **Before an edit or delete**, the app re-reads the row from Excel and refuses if it changed (e.g. you sorted the table in Excel meanwhile). If that happens, press refresh (↻) and try again.
- **What you enter vs. what's filled in.**
  - *Time:* Internal-type projects (Internal, Training, …) are listed first and default to non-billable. Opportunity and Hold projects are listed last, also non-billable by default.
  - *PO approved:* on the Time screen, pick the project in the list filter (and *This year*). If it has non-billable, uninvoiced entries, a **Make billable** button sets them all billable so they go on the next invoice, whatever month they were worked in. Then set the project to Active in the Projects table.
  - *Expenses:* Paid is set to TRUE unless you paid personally (Reimburse). Billable is FALSE when there's no project. Scan is TRUE when you attach a receipt in the app. HST is left to Excel's formula. Untick *Includes 13% HST* to enter a different tax amount (or 0) for that row only.
  - *Mileage:* PD Paid starts FALSE; set it on the spreadsheet when per diems are paid out.
- **Lists** filter by Today, This week (Mon–Sun), Last week, Last 7 / 14 days, This month, Last month, or This year (the whole workbook).
- **Invoice date** defaults to this month's billing day: the first Monday, or the next business day if it's an Ontario stat holiday or the Civic Holiday (e.g. Tue Sep 8, 2026 after Labour Day). Add other days off to `extraHolidays` in `js/config.js`.
- **Mark paid** asks for the payment date, sets `Paid` on the invoice, and (optionally) sets `Paid = TRUE` on that invoice's time entries and disbursements, matched by invoice number. For an invoice recorded before tracking, it can only do this if that's the client's one unpaid invoice from then (all their invoiced, unpaid entries must be on it). Otherwise it says so, and you mark those rows in the spreadsheet as before. The dialog shows the count and the list before anything is written.
- **Project budget while logging time:** after you pick a fixed-fee project, the form shows its hours remaining: (fee − billed − time not yet invoiced) ÷ rate, the same figures as the Projects table. As you type the hours, it shows what's left after this entry and warns if it goes over. Hourly projects show the rate and unbilled hours.
- **Add a project** (Projects page, or *+ New project* on the Time screen): name, client, project ID, client ref/PO #, status, type, rate, fee (fixed fee only), prime and notes. Started from the Time screen, it returns there with the new project selected. After saving, the app re-reads the Projects table to confirm the project is there. The project ID is suggested from your numbering (e.g. `WF-024`, `FWC-005`) and can be changed. Names must be unique, because time and invoices look projects up by name. New clients are added in Excel. The Projects page also lists open projects with their budget status.
- **Expense report** (Expenses page → *Expense report*): the same rule as your *Expense Rep Creator* sheet. It includes expenses marked Reimburse that aren't Paid, for the chosen period, with the business portion as the total. It adds **per diems** from the Mileage table (trips where PD Paid is unticked), priced from `Mileage!Q22:R26`. Checkboxes include **motor vehicle expenses** and the **mileage allowance** (km × the rates in `Mileage!T7:V8`, tiered on the year's cumulative km); both are off by default. The report runs to the next billing day by default and shows it on screen and in print. **Print / Save PDF** gives a branded report with sign-off lines. **Mark as reimbursed** sets Paid on those expenses and PD Paid on those trips. The km allowance has no paid flag in the workbook, so it's always based on the period you choose.
- **Invoice pages:** click any issued invoice to see every time entry and disbursement billed on it, grouped by project. **Print time detail** gives the client a branded breakdown.
- **Expected invoice:** in the Time list, billable time not yet invoiced shows the invoice it's expected to go on (e.g. `→ WF2611` for the November billing day). This is display only. The number is written only when the invoice is actually recorded, so a skipped or delayed billing run never leaves wrong numbers in the workbook.
- **Before printing or recording an invoice**, a checklist shows the billing period (the day after last month's billing day, up to the invoice date). It lists each person's hours in that period and their latest time entry, and flags:
  - invoiced time with no invoice number,
  - non-billable time on the client's projects,
  - included entries dated outside the period.

  You tick *All time for this period has been entered* to continue.
- **Invoice check** (Invoices page): once tracking is on, any time entry marked Invoiced without an invoice number appears here, usually because it was entered or ticked by hand in Excel. In the Time list it shows as *Invoiced, no #*. Pick which invoice each one was on and save; the app doesn't guess. After recording an invoice, the app also re-reads the table and warns if any included entry didn't get the number.
- **Invoices** use the same rules as the FILTER formulas on *Invoice Creator*: hourly projects with uninvoiced billable hours, fixed-fee projects with an amount this invoice or status `Active*`, and uninvoiced disbursements. HST is 13%, terms are net 30, and the number is client `Short` + `YYMM`. On fixed-fee lines, *This invoice* is the last column so the amounts sum to the subtotal. **Print / Save PDF** uses the browser's print dialog. **Record invoice** adds the row to `Invoices` and (optionally) sets `Invoiced = TRUE` on the included Tasks and Disbursements rows. It tells you exactly how many rows before it writes.
- Each write shows the exact cell range it touched (e.g. `Time!A755:I755`). The dashboard keeps a log of recent writes made from that device.
- Receipt photos are shrunk, read on the device with Tesseract OCR to pre-fill vendor, date and total, and saved as `YYYY-MM-DD_Vendor_Amount.jpg`.

## Rules for the workbook

- Keep the table names (`Tasks`, `Expenses`, `Mileage`, `Invoices`, `Disbursements`, `Projects`, `Clients`, `Workers`, `Categories`) and the column headers the app uses. You can add new columns and reorder columns freely.
- The file must stay a `.xlsx` in SharePoint/OneDrive. Leaving it open in desktop Excel with AutoSave on is fine.
- **New year:** create `Bookkeeping 2027.xlsx`, then in **Settings** paste its link and change the receipts folder to `Receipts/2027`.

## Updating the app

Each file is referenced with a version tag (`?v=20261014` in `index.html`, `js/app.js` and `js/store-demo.js`). When you change the code, change that number everywhere (same value), so browsers and GitHub Pages' cache load the new files together instead of mixing old and new.

## Files

```
index.html            page shell
app.css               styles (Ennovantiv palette, mobile layout, B&W-safe invoice print)
js/config.js          your settings: IDs, tax, invoice terms
js/app.js             screens: dashboard, time, expenses, mileage, invoices, settings
js/store-graph.js     Microsoft Graph Excel read/write
js/store-demo.js      sample-data mode (?demo)
js/receipt.js         photo shrink + OCR
```
