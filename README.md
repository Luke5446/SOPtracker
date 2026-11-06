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

- v1.0 (6 Nov 2026) – first version: find an order, late & at risk in three
  groups, can jump the queue, all live orders, Sage paste + planner data.
