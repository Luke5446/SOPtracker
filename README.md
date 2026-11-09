# SOPtracker – Order tracker for the sales office

One page, `index.html`, served by GitHub Pages at
https://luke5446.github.io/SOPtracker/. Type a sales order number and see the
order, what is in stock, what is missing and why (works order, container,
supplier PO, or nothing on order), the date it was added to Sage, the promised
date and the expected completion date. Tibard and Oliver Harvey, intercompany
orders left out.

## Tabs

- **Find an order** – SO number (leading zeros optional), or a customer, their
  order ref or a product code. Click a line to see every live order that wants
  the same code, in the order the stock goes out.
- **Late & at risk** – live orders past their promised date (late) or whose
  cover arrives after it / on an unknown date / not at all (at risk), split
  into **Bought in**, **Made in house** (works order or Clockwork container,
  including codes usually bought from Clockwork that have a works order on)
  and **Waiting on both**. Tiles filter; each group exports to CSV.
- **Can jump the queue** – orders that could go out in full by borrowing stock
  held for an order that cannot go out anyway. Shows the lender, what it is
  waiting on and whether the next delivery of the code backfills it in time.
  Nothing here changes Sage; it is for the sales office to decide.
- **All live orders** – every live order with its status and expected date.
- **Data** – where each feed came from and when.

## Data

1. **Sage** – `ordertracker_live.sql` (the POtracker query with every product
   line on a live order and the stock row for each code in both companies plus
   its BULK item). The sales manager refreshes the Excel query, copies columns
   A to AD, opens the page with `?edit`, presses **Paste today's data** and
   then **Publish for team**. Everyone else opens the plain address.
2. **Works orders** – the production planner's published `data.json`
   (`WOs`, `smPending`, buffer `rows`), read when the page opens.
3. **Containers** – the stock planner's published `data.json` (`cPOs`),
   read when the page opens.

Who gets what, per code: Sage allocation on the line; free stock to the oldest
promised date first (then order date, SO, line – the planner's rule); a
special-make works order raised for the line; the back-to-back PO Sage
generated from the order; live works orders oldest raised first; containers in
landing order (+1 working day); purchase orders by date; anything left is
"not on order". Made-in-house codes pool Tibard + BULK + OH stock (Settings
can switch this off); bought-in codes are per company. **Contested** marks a
line that has stock only because it is older than another order that goes
short on the same code.

## Publishing

Same as POtracker: a fine-grained GitHub token on the publishing PC with
**Contents: read and write** on this repo, pasted once via **Set up this PC**.
`data.json` holds the last published paste; viewers read it. Settings are
published with it. A PC that has not pasted today takes the published data when
it is newer than its own.

## Later

Built so the Sage paste can be swapped for a live Sage link when the apps move
on-prem: the model is built from three plain feeds (SO/PO/STK rows, works
orders, containers) and nothing else.

## Changelog

- v1.8 (9 Nov 2026) – two more tiles on Late & at risk: late orders waiting on
  production (a works order, or a special make with no WO yet – Clockwork not
  counted) and late orders waiting on suppliers (a container or supplier PO,
  Clockwork included). An order waiting on both shows in both.
- v1.7 (9 Nov 2026) – the back arrow now steps back: an order opened from a
  search list returns to that list, one opened from another tab returns to that
  tab; with nothing to go back to it clears the search as before.
- v1.6 (9 Nov 2026) – back arrow (←) beside the Find box, shown once a search
  or order is on screen. It clears the search and returns to the blank Find
  tab, the same as the × in the box. Asked for by the sales office.
- v1.5 (6 Nov 2026) – clearing the Find box (backspace or its ×) clears the
  results straight away; before, the last search stayed on screen until Find
  was pressed again.
- v1.4 (6 Nov 2026) – Special makes tile on Late & at risk: orders with a
  made-in-house line whose code is not stock held in Sage (or that the planner
  lists as a special make), with how many are still in production. SM pill on
  the order and the line. Orders needing embroidery: planned for later.
- v1.3 (6 Nov 2026) – the Find an order tab opens blank: just the search box
  until you type an order number, customer, order ref or code. The late /
  at-risk counts are on the Late & at risk tab.
- v1.2 (6 Nov 2026) – whole-column pastes from Excel bring 150,000+ empty
  rows with them, more than the browser will store, so Publish found nothing
  saved ("Paste today's data first"). Blank rows and the header are now
  dropped on load, Publish sends the paste held in memory, and a paste the
  browser cannot keep says so and still publishes.
- v1.1 (6 Nov 2026) – tested on the first real paste (6,127 SO lines).
  A works order or PO whose date has passed and is still not in now makes
  the line **at risk** and shows "overdue"; proforma and web accounts
  (PROFORMA, PROFEURO, XONLINE) show the customer's name from the order
  reference. SQL rewritten to read each Sage table once (table variables)
  and run without locks.
- v1.0 (6 Nov 2026) – first version: find an order, late & at risk in three
  groups, can jump the queue, all live orders, Sage paste + planner data.
