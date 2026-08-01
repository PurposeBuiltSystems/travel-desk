# Power BI dashboard recipe — Travel Desk planner

Travel Desk appends every trip as a row in your shared planner — an
**Excel table on SharePoint** — which is Power BI's happiest data source.
No exports, no gateway: connect once, schedule refresh, and the
coordinator gets live travel analytics. About an hour to build.

## What you get

- **Budget burn** — travel spend by division against each division's
  travel budget for the fiscal year (the chart budget requests are
  argued with).
- **Coordinator queue** — trips by Status (Requested → Booked →
  Approved), oldest first.
- **Travel map** — destinations plotted; conference repeat-offenders
  listed.
- **Net cost after reimbursement** — third-party/grant reimbursements
  netted out, so "what this actually costs us" is a number, not a guess.

## Connect (10 minutes)

1. In SharePoint, open the planner workbook's library → ⋯ → **Copy link**
   (a direct file link).
2. Power BI Desktop → **Get Data → Excel workbook** → paste the SharePoint
   URL (strip anything after `.xlsx`) → in Navigator, pick the planner
   **table** (Travel Desk requires the sheet to be a real Insert > Table,
   so it shows by name — not just a sheet).
3. Set column types: date columns → Date, the cost/estimate column →
   Decimal (Travel Desk writes rounded numbers; if older hand-typed rows
   carry `$` text, add a `Text.Select([Cost], {"0".."9","."})` cleanup
   step first).

### Multiple fiscal years

Planner templates drift year to year (new workbook per SFY, headers
shift). Two options:

- **Simple:** add each year's workbook as its own query, rename columns
  to one canonical set (Traveler, Division, Bureau, Event, Destination,
  Date, EstCost, Funding, FY, ReimbursePct, Status), then **Append**.
- **Automatic:** use the **SharePoint folder** connector on the planner
  library and combine files — more setup, survives future years without
  edits.

The planner already carries a fiscal-year column (Travel Desk computes
it), so FY is your slicer for free.

## The budget table (the chart that matters)

Division travel budgets usually live in a planning document, not a
system. Create them as data once:

- Power BI → **Enter Data** → a small table: `Division`, `FY`, `Budget`
  (or keep it as a SharePoint list so the coordinator can edit it
  without opening the report).
- Relate `Planner[Division]` → `Budgets[Division]` (and FY if you budget
  per year).

```
Travel Spend = SUM ( Planner[EstCost] )

Net Spend = SUMX ( Planner,
    Planner[EstCost] * ( 1 - DIVIDE ( Planner[ReimbursePct], 100 ) ) )

Budget = SUM ( Budgets[Budget] )

Remaining = [Budget] - [Travel Spend]

Burn % = DIVIDE ( [Travel Spend], [Budget] )

Trips = COUNTROWS ( Planner )

Avg Cost / Trip = DIVIDE ( [Travel Spend], [Trips] )
```

When planning next year's request, a common convention is current-year
budget plus an inflation factor (e.g. +10%) — add a
`Next FY Ask = [Budget] * 1.10` measure if that matches your practice.

## Suggested pages

1. **Budget** — bar per division: Travel Spend vs Budget with Remaining
   labels; card row: Burn %, Trips, Avg Cost/Trip; FY slicer.
2. **Queue** — table of trips with Status ≠ done, oldest travel date
   first (the coordinator's follow-up list, mirroring the Status column
   Travel Desk stamps as "Requested").
3. **Map** — the Destination column on a map visual, bubble size =
   EstCost; alongside: top events by total cost (the "do we really send
   six people to this conference every year?" view).
4. **Reimbursement** — Travel Spend vs Net Spend by division; third-party
   funded trips listed.

## Publish

Publish to your workspace → scheduled refresh (SharePoint credentials =
your org account; daily is fine) → share with the coordinator and
division heads, or pin as a tab in whatever Teams channel travel
planning lives in. Available in government (GCC) tenants.

> Privacy note: the report reads the same planner workbook your team
> already shares, under each viewer's own permissions. Nothing new is
> collected and nothing leaves your tenant.
