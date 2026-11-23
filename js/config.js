// ---------------------------------------------------------------------------
// App configuration. Nothing in this file is secret: the Client ID identifies
// the app to Microsoft, it does not grant access. Access is granted only when
// a person signs in with an account that (a) is allowed into the app
// registration and (b) has permission to the workbook in SharePoint/OneDrive.
// ---------------------------------------------------------------------------
window.BOOKS_CONFIG = {
  appName: 'Ennovantiv Books',

  // From Entra ID (Azure) > App registrations > your app > Overview
  clientId: '749232ca-bd55-4054-b691-84d534ea147e',
  tenantId: '2159ca4c-7474-40a2-b704-8040dd9494fb',

  // Sharing link to the workbook (Excel for the web > Share > Copy link).
  // Can also be set/changed from the app's Settings screen (e.g. each new year).
  fileShareUrl: '',

  // Receipt photos are saved here, relative to the workbook's folder.
  receiptFolder: 'Receipts/2026',

  // Who can use the app. Sign-in is limited to the Ennovantiv Microsoft 365
  // tenant (tenantId above). On top of that, the signed-in address must end in
  // one of these domains, which keeps out guest accounts invited from other
  // organizations. allowedUsers can narrow it further to named people.
  allowedDomains: ['ennovantiv.com'],
  allowedUsers: [],

  // Invoice settings (mirrors the "Invoice Creator" sheet)
  taxRate: 0.13,
  taxLabel: 'HST',
  taxNumber: '735771420-RT0001',
  invoiceTermsDays: 30,
  invoiceFooter: [
    'Terms: Net 30 Days.',
    'Interest: 1.5% per month (18% per annum)',
  ],
  companyName: 'Ennovantiv Services Inc.',
  companyClientRow: 'Internal', // Clients row holding your own address/phone/email
  logo: {                       // Ennovantiv mark, made from the brand files
    onDark: 'assets/mark-blue.png',   // header (navy background)
    onLight: 'assets/mark-navy.png',  // sign-in screen
    print: 'assets/mark-black.png',   // invoices, time detail, expense reports
  },

  // Invoices are dated the first Monday of the month, or the next business
  // day when that Monday is a holiday. Ontario holidays are built in; add any
  // other days off here as 'YYYY-MM-DD'.
  extraHolidays: [],

  // Invoice tracking: a column with this name on the Tasks and Disbursements
  // tables (Settings → Invoice tracking adds it). Recording an invoice writes
  // its number there, so "Mark paid" and the invoice's time detail are exact.
  invoiceLinkColumn: 'Invoice',

  // Project statuses offered for new time entries. Opportunity / on-hold
  // projects are listed separately and default to non-billable until a PO.
  activeStatuses: ['Active', 'Active*'],
  opportunityStatuses: ['Opportunity', 'Hold'],

  // Starting a new year (Settings → New year).
  rolloverKeepStatuses: ['Active', 'Active*', 'Opportunity', 'Hold'],   // projects carried into the new workbook
  // Tables emptied in the NEW workbook (never the current one). Clients, Workers, Categories and
  // the Status / Type lists are kept. A table that doesn't exist is skipped.
  rolloverClear: ['Tasks', 'Expenses', 'Mileage', 'Invoices', 'Disbursements', 'Payroll', 'Sales', 'MainAct', 'TaxAct', 'Visa', 'ChangeLog'],
  // Cells tied to the year that are typed in by hand, so the app lists them rather than changing them.
  // {from} and {to} become the two years.
  rolloverChecklist: [
    'Calendar!A1: the year start date',
    'Calendar!B25 and B26: opening bank balances (typed into the formulas)',
    'Mileage!Q2 and R2: the year start date and the odometer reading on 1 January',
    'Mileage!R21:R26 (the "{from} Rate" heading and the per diem rates) and T7:V8 (the km rates)',
    'Weekly!B2: the date the weekly grid starts',
    'Payroll!R12:R21: the remittance dates, and the dated notes in column U',
    'Search the workbook for "{from}" to catch anything else',
  ],

  // Expense report: per diem rates and mileage allowance rates are read from
  // these cells on the Mileage sheet, so updating the sheet updates the report.
  perDiemRange: 'Mileage!Q22:R26',      // label, rate: Breakfast … Private Acc
  mileageRateRange: 'Mileage!T7:V8',    // first-rate, label, threshold km / after-rate
  vehicleCategory: 'Motor Vehicle Expenses',
};

// Excel table schema. `computed` columns are Excel formulas: the app never
// writes them, Excel calculates them.
window.BOOKS_SCHEMA = {
  Tasks:         { computed: ['Short'] },
  Expenses:      { computed: ['Tax', 'Bus Total', 'Bus Tax'] },
  Mileage:       { computed: [] },
  Invoices:      { computed: ['Tax', 'Due Date', 'Days'] },
  Disbursements: { computed: ['Markup  $', 'Billable'] },
  Payroll:       { computed: ['Net', 'Remittance', 'Total'] },
  Sales:         { computed: ['Tax'] },
  MainAct:       { computed: ['Date'] },
  TaxAct:        { computed: ['Date'] },
  Visa:          { computed: ['Date'] },
  Projects:      { computed: ['%Comp', 'Hours', 'This Invoice', 'Billed', 'Total', 'Remaining', 'Overage', 'Total Hours', 'Short'] },
  Status:        { readOnly: true },
  Table12:       { readOnly: true }, // project Types
  Clients:       { readOnly: true },
  Workers:       { readOnly: true },
  Categories:    { readOnly: true },
};
