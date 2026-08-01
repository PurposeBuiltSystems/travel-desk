# Travel Desk

Outlook add-in that turns one filled-out travel request into **both** artifacts
an email-plus-spreadsheet travel process needs:

1. the **Travel Authorization** form as a printable email draft to the travel
   coordinator (structured subject: `Travel Auth - <Name> - <Event> - <Date>`),
2. the matching **row in the shared travel planner** workbook on
   SharePoint/OneDrive — mapped by column header, so year-to-year template
   changes don't break it.

One entry, no re-typing, totals that always agree.

## Architecture

Same as the rest of the PurposeBuilt suite: XML manifest + static files on
GitHub Pages, MSAL nested app authentication, Microsoft Graph with delegated
scopes only (`Files.ReadWrite.All`, `Mail.ReadWrite`), no backend, no data
collection, deterministic logic (no AI).

- `src/form.js` — pure logic: totals, Iowa SFY/FFY, structured subject,
  planner-row mapping by fuzzy header match, printable form HTML.
- `src/graph.js` — Graph layer: share-link → driveItem resolution, workbook
  tables (list/headers/add row), draft creation.
- `src/taskpane/` — the form UI.
- `test/form.test.js` — offline tests (`node test/form.test.js`).

## Org profiles

Coordinators configure once (workbook, table, coordinator email, org name,
fiscal-year calendar, funding label) and click **Copy profile**; travelers
paste the code and click **Apply** — fully set up in one step. Iowa DOT is
simply one profile; any org that tracks travel in a shared spreadsheet and
emails authorization forms can use its own.

## Analytics

The planner is an Excel table on SharePoint — Power BI's native food.
[POWERBI.md](POWERBI.md) is the one-hour build recipe: budget burn per
division, the coordinator's follow-up queue, a destination map, and
net-cost-after-reimbursement measures.

## Deployment

Internal / pilot tool — not on AppSource. Sideload via aka.ms/olksideload with
`manifest.xml`, or have IT centrally deploy it.

**Planner prerequisite:** the planner sheet must be an Excel **Table**
(Insert > Table over the header + data rows).
