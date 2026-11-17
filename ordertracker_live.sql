/* ============================================================================
   ORDER TRACKER v2 - LIVE SALES ORDERS, PURCHASE ORDERS, STOCK AND COMMENTS
   Sage 200 Professional, server TIB-SQL-002, both companies (S200_LIVE =
   Tibard, OliverHarveyLive = Oliver Harvey) in one pass.

   This is the POtracker v3 query (POtracker_live.sql) with TWO changes:

     1. SO rows include EVERY product line on a live or held sales order,
        despatched lines too (Qty3 outstanding = 0), so an order shows in
        full - what has gone and what is still to go.
     2. STK rows cover each code in both companies plus its BULK item (see
        the codes CTE), so the tracker can pool Tibard, BULK and OH stock for
        a made-in-house code the way the Buffer sheet does.

   POtracker reads this paste too (it ignores outstanding-0 lines, and the
   extra STK rows only add a few entries to its Stock held tab), but the
   buyer's own query stays the one for POtracker.

   ONE PASTE. Copy columns A to AD, all rows, without the header. Four row
   types share one column layout and the app splits them on column A:

     SO   one row per product line on a live or held sales order
     CMT  one row per COMMENT line on those same orders (LineTypeID 3), the
          text in column E - added 12 Nov 2026 (tracker v2.8) so the order
          drill-down shows the comments typed on the Sage order. Only the
          columns A, B, C, E, F, G and AB are filled. POtracker and tracker
          versions before 2.8 skip the row type, so the paste still works
          there unchanged.
     PO   one row per live purchase order line with quantity outstanding
     STK  one row per stock code per company, for every code that appears on
          an SO or PO row - the classification and the stock figures

   The STK rows also differ a little from POtracker's: see the codes CTE.

   Intercompany lines (Tibard SO to OLIVER, OH SO to TIB003, OH PO on TIB001)
   are returned with column S = Y. The tracker keeps them out of the search
   and the late list but follows them to show where an OH order bought via
   Tibard really is.

   COLUMN LAYOUT - FIXED. The app parses by position. Blank = not applicable.

     A  RowType         SO / PO / STK
     B  Company         TIBARD / OLIVER HARVEY
     C  RowKey          TIB-SO-<lineid>, TIB-CMT-<lineid>, OH-PO-<lineid>,
                        TIB-STK-<code> ...
     D  ProductCode
     E  Description     SO/PO: line description. CMT: the comment text.
                        STK: stock item name, or CODE NOT IN STOCK FILE
     F  DocNo           SO/CMT: sales order no. PO: purchase order no
     G  LineSeq         print sequence on the order
     H  AccountNo       SO: customer account. PO: supplier account.
                        STK: preferred supplier account
     I  AccountName     as above
     J  Reference       SO: customer's own order ref. PO: supplier's ref.
                        STK: supplier's stock code
     K  DateISO         SO: promised date, ORDER HEADER FIRST (amending the
                        date in Sage changes the header only).
                        PO: requested delivery date - header, POP has no line date
     L  Qty1            SO/PO: ordered.             STK: confirmed in stock
     M  Qty2            SO: despatched. PO: received. STK: allocated (SOP+stock+BOM)
     N  Qty3            SO/PO: OUTSTANDING.         STK: FREE STOCK (Sage's figure)
     O  Qty4            SO: allocated to this line.  STK: Sage's cached on-PO qty
     P  FulfilMethodID  SO: fulfilment method on the line. STK: the item's default
     Q  LinkID          SO: the first back-to-back PO generated from this order.
                        PO: the sales order this PO was generated from
     R  B2BCount        SO: how many POs this order generated. PO: how many SOs
                        it was generated from. STK: back-to-back POs with this
                        code in the last 12 months
     S  Intercompany    Y on Tibard SO to OLIVER, OH SO to TIB003, OH PO on TIB001
     T  Manufacturer    STK, exactly as typed in Sage. Blank is meaningful
     U  StockHeld       STK: Yes or blank (Tibard AnalysisCode3; OH also
                        AnalysisCode7 Website or the Tibard record)
     V  WebsiteOH       STK: Yes when Oliver Harvey AnalysisCode7 = Yes
     W  ProductGroup    STK: product group code (54 = Additional Charges)
     X  BOMItemTypeID   STK
     Y  LeadTime        STK: preferred supplier's lead time
     Z  LeadTimeUnitID  STK: unit of Y
     AA MOQ             STK: preferred supplier's minimum order quantity
     AB DocStatusID     SO/CMT/PO: order status. 0 = live, 1 = on hold
     AC OrderDate       SO/PO: header DocumentDate (the date the order was
                        raised in Sage), yyyy-mm-dd
     AD EnteredBy       SO/PO: DocumentCreatedBy, the Sage login that raised it

   MADE IN HOUSE: the tracker's test is an exact match on Manufacturer
   (Tibard, Oliver Harvey, Clockwork, Clockwork Clothing, Clockwork Giant
   Clothing - editable in its Settings), OR the code is on the production
   planner's buffer list, a live works order, or the stock planner's
   container list. Everything else is bought in.

   STATUS FILTERS: DocumentTypeID 0 = order (not return). DocumentStatusID
   0 = live, 1 = on hold, 2 = complete, 4 = cancelled/disputed. Live and held
   are both returned; the app shows held orders with an On hold badge.

   Dates are yyyy-mm-dd text so the browser never guesses dd/mm against
   mm/dd; tabs and line breaks are stripped from names so the paste survives
   Excel. QUANTITIES ARE IN STOCK UNITS.

   SQL 1.xlsx: Power Query returns the first result set only - this script is
   one SELECT. Run it in a new workbook connection, Refresh All, then copy
   A to AD into the tracker's Paste box (editor mode, ?edit on the URL).
   DEADLOCKS: on 6 Nov 2026 the first run was picked as a deadlock victim
   (error 1205) by a Sage posting. This is a report, so it reads without
   taking locks - READ UNCOMMITTED below is the standard reporting setting.

   SPEED (6 Nov 2026): the first cut, written as one chain of CTEs, was still
   running after 15 minutes. SQL Server re-runs a CTE every time it is named,
   and the SO/PO blocks were named five or six times each, every time with two
   correlated subqueries per line into MMS123BackToBackPO. This version
   builds each step ONCE into a table variable (@so, @po, @codes, @stk) and
   looks the back-to-back links up from two small pre-grouped tables, so each
   Sage table is read once. Table variables rather than #temp tables because
   Power Query reads the result shape before it runs the batch and #temp
   tables are invisible to that check ("Invalid object name '#so'"). SET
   NOCOUNT ON keeps the row-count messages out of the way so the final SELECT
   is the only result. Text columns carry COLLATE DATABASE_DEFAULT because
   tempdb on TIB-SQL-002 collates differently from S200_LIVE (error 468
   otherwise). OPTION (RECOMPILE) lets SQL Server see the real row counts in
   the table variables when it plans each step.
   ========================================================================= */


SET NOCOUNT ON;
SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;

/* Working tables. Text columns carry COLLATE DATABASE_DEFAULT - see SPEED. */
DECLARE @b2bso TABLE (Company nvarchar(20) COLLATE DATABASE_DEFAULT, SOPOrderReturnID bigint,
                      FirstDoc nvarchar(40) COLLATE DATABASE_DEFAULT, N int,
                      PRIMARY KEY (Company, SOPOrderReturnID));
DECLARE @b2bpo TABLE (Company nvarchar(20) COLLATE DATABASE_DEFAULT, POPOrderReturnID bigint,
                      FirstDoc nvarchar(40) COLLATE DATABASE_DEFAULT, N int,
                      PRIMARY KEY (Company, POPOrderReturnID));
DECLARE @so TABLE (
    Company nvarchar(20) COLLATE DATABASE_DEFAULT, RowKey nvarchar(60) COLLATE DATABASE_DEFAULT,
    ProductCode nvarchar(60) COLLATE DATABASE_DEFAULT, Description nvarchar(255) COLLATE DATABASE_DEFAULT,
    DocNo nvarchar(40) COLLATE DATABASE_DEFAULT, LineSeq int,
    AccountNo nvarchar(40) COLLATE DATABASE_DEFAULT, AccountName nvarchar(255) COLLATE DATABASE_DEFAULT,
    Reference nvarchar(255) COLLATE DATABASE_DEFAULT, DateISO nvarchar(10) COLLATE DATABASE_DEFAULT,
    Qty1 decimal(18,2), Qty2 decimal(18,2), Qty3 decimal(18,2), Qty4 decimal(18,2),
    FulfilMethodID nvarchar(10) COLLATE DATABASE_DEFAULT, LinkID nvarchar(40) COLLATE DATABASE_DEFAULT,
    B2BCount nvarchar(10) COLLATE DATABASE_DEFAULT, Intercompany nvarchar(1) COLLATE DATABASE_DEFAULT,
    DocStatusID nvarchar(10) COLLATE DATABASE_DEFAULT, OrderDate nvarchar(10) COLLATE DATABASE_DEFAULT,
    EnteredBy nvarchar(255) COLLATE DATABASE_DEFAULT);
DECLARE @po TABLE (
    Company nvarchar(20) COLLATE DATABASE_DEFAULT, RowKey nvarchar(60) COLLATE DATABASE_DEFAULT,
    ProductCode nvarchar(60) COLLATE DATABASE_DEFAULT, Description nvarchar(255) COLLATE DATABASE_DEFAULT,
    DocNo nvarchar(40) COLLATE DATABASE_DEFAULT, LineSeq int,
    AccountNo nvarchar(40) COLLATE DATABASE_DEFAULT, AccountName nvarchar(255) COLLATE DATABASE_DEFAULT,
    Reference nvarchar(255) COLLATE DATABASE_DEFAULT, DateISO nvarchar(10) COLLATE DATABASE_DEFAULT,
    Qty1 decimal(18,2), Qty2 decimal(18,2), Qty3 decimal(18,2), Qty4 decimal(18,2),
    FulfilMethodID nvarchar(10) COLLATE DATABASE_DEFAULT, LinkID nvarchar(40) COLLATE DATABASE_DEFAULT,
    B2BCount nvarchar(10) COLLATE DATABASE_DEFAULT, Intercompany nvarchar(1) COLLATE DATABASE_DEFAULT,
    DocStatusID nvarchar(10) COLLATE DATABASE_DEFAULT, OrderDate nvarchar(10) COLLATE DATABASE_DEFAULT,
    EnteredBy nvarchar(255) COLLATE DATABASE_DEFAULT);
DECLARE @cmt TABLE (
    Company nvarchar(20) COLLATE DATABASE_DEFAULT, RowKey nvarchar(60) COLLATE DATABASE_DEFAULT,
    Description nvarchar(max) COLLATE DATABASE_DEFAULT, DocNo nvarchar(40) COLLATE DATABASE_DEFAULT,
    LineSeq int, DocStatusID nvarchar(10) COLLATE DATABASE_DEFAULT);
DECLARE @codes0 TABLE (Company nvarchar(20) COLLATE DATABASE_DEFAULT, ProductCode nvarchar(60) COLLATE DATABASE_DEFAULT,
                       PRIMARY KEY (Company, ProductCode));
DECLARE @codes  TABLE (Company nvarchar(20) COLLATE DATABASE_DEFAULT, ProductCode nvarchar(60) COLLATE DATABASE_DEFAULT,
                       PRIMARY KEY (Company, ProductCode));
DECLARE @b2b12  TABLE (Company nvarchar(20) COLLATE DATABASE_DEFAULT, ProductCode nvarchar(60) COLLATE DATABASE_DEFAULT, N int,
                       PRIMARY KEY (Company, ProductCode));
DECLARE @stk TABLE (
    Company nvarchar(20) COLLATE DATABASE_DEFAULT, RowKey nvarchar(70) COLLATE DATABASE_DEFAULT,
    ProductCode nvarchar(60) COLLATE DATABASE_DEFAULT, Description nvarchar(255) COLLATE DATABASE_DEFAULT,
    AccountNo nvarchar(40) COLLATE DATABASE_DEFAULT, AccountName nvarchar(255) COLLATE DATABASE_DEFAULT,
    Reference nvarchar(255) COLLATE DATABASE_DEFAULT,
    Qty1 decimal(18,2), Qty2 decimal(18,2), Qty3 decimal(18,2), Qty4 decimal(18,2),
    FulfilMethodID nvarchar(10) COLLATE DATABASE_DEFAULT, Manufacturer nvarchar(255) COLLATE DATABASE_DEFAULT,
    StockHeld nvarchar(3) COLLATE DATABASE_DEFAULT, WebsiteOH nvarchar(3) COLLATE DATABASE_DEFAULT,
    ProductGroup nvarchar(40) COLLATE DATABASE_DEFAULT, BOMItemTypeID nvarchar(10) COLLATE DATABASE_DEFAULT,
    LeadTime nvarchar(10) COLLATE DATABASE_DEFAULT, LeadTimeUnitID nvarchar(10) COLLATE DATABASE_DEFAULT,
    MOQ decimal(18,2), B2B12m nvarchar(10) COLLATE DATABASE_DEFAULT);

/* ------------------------------------------------------------------------
   1. Back-to-back links, grouped once. Sage links each generated PO to its
   sales order in MMS123BackToBackPO. @b2bso: per sales order, the first PO
   it generated and how many. @b2bpo: per purchase order, the first sales
   order it came from and how many. If the Oliver Harvey half errors on
   MMS123BackToBackPO, the table is not in that database - delete the OH half
   of each statement (from UNION ALL to the end of the inner SELECT).
   ------------------------------------------------------------------------ */
INSERT INTO @b2bso (Company, SOPOrderReturnID, FirstDoc, N)
SELECT 'TIBARD', b.SOPOrderReturnID, MIN(lpo.DocumentNo), COUNT(*)
FROM   S200_LIVE.dbo.MMS123BackToBackPO b
JOIN   S200_LIVE.dbo.POPOrderReturn lpo ON lpo.POPOrderReturnID = b.POPOrderReturnID
GROUP BY b.SOPOrderReturnID
UNION ALL
SELECT 'OLIVER HARVEY', b.SOPOrderReturnID, MIN(lpo.DocumentNo), COUNT(*)
FROM   OliverHarveyLive.dbo.MMS123BackToBackPO b
JOIN   OliverHarveyLive.dbo.POPOrderReturn lpo ON lpo.POPOrderReturnID = b.POPOrderReturnID
GROUP BY b.SOPOrderReturnID;

INSERT INTO @b2bpo (Company, POPOrderReturnID, FirstDoc, N)
SELECT 'TIBARD', b.POPOrderReturnID, MIN(lso.DocumentNo), COUNT(*)
FROM   S200_LIVE.dbo.MMS123BackToBackPO b
JOIN   S200_LIVE.dbo.SOPOrderReturn lso ON lso.SOPOrderReturnID = b.SOPOrderReturnID
GROUP BY b.POPOrderReturnID
UNION ALL
SELECT 'OLIVER HARVEY', b.POPOrderReturnID, MIN(lso.DocumentNo), COUNT(*)
FROM   OliverHarveyLive.dbo.MMS123BackToBackPO b
JOIN   OliverHarveyLive.dbo.SOPOrderReturn lso ON lso.SOPOrderReturnID = b.SOPOrderReturnID
GROUP BY b.POPOrderReturnID;

/* ------------------------------------------------------------------------
   2. SO rows: every product line on a live or held sales order, both
   companies, despatched lines too (Qty3 = 0) so the tracker shows the whole
   order. Tabs and line breaks stripped - the app splits on TAB.
   ------------------------------------------------------------------------ */
INSERT INTO @so (Company, RowKey, ProductCode, Description, DocNo, LineSeq, AccountNo, AccountName,
                 Reference, DateISO, Qty1, Qty2, Qty3, Qty4, FulfilMethodID, LinkID, B2BCount,
                 Intercompany, DocStatusID, OrderDate, EnteredBy)
/* ---------------------------------------------------------------- TIBARD -- */
SELECT
    'TIBARD',
    'TIB-SO-' + CAST(sorl.SOPOrderReturnLineID AS varchar(20)),
    ISNULL(LTRIM(RTRIM(sorl.ItemCode)),''),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(sorl.ItemDescription,''), CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))),
    sor.DocumentNo,
    sorl.PrintSequenceNumber,
    ISNULL(cust.CustomerAccountNumber,''),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(cust.CustomerAccountName,''), CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(sor.CustomerDocumentNo,''), CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))),
    /* Promised date: order header first (amending the date in Sage changes
       the header only), then the line, then requested. yyyy-mm-dd text. */
    CONVERT(varchar(10), COALESCE(sor.PromisedDeliveryDate,  sorl.PromisedDeliveryDate,
                                  sor.RequestedDeliveryDate, sorl.RequestedDeliveryDate), 23),
    CAST(sorl.StockUnitLineQuantity AS decimal(18,2)),
    CAST(ISNULL(sorl.StockUnitDespRcptQuantity,0) AS decimal(18,2)),
    CAST(sorl.StockUnitLineQuantity - ISNULL(sorl.StockUnitDespRcptQuantity,0) AS decimal(18,2)),
    CAST(ISNULL(sorl.StockUnitAllocatedQuantity,0) AS decimal(18,2)),
    ISNULL(CAST(sorl.SOPOrderFulfilmentMethodID AS varchar(10)),''),
    ISNULL(bs.FirstDoc,''),
    CAST(ISNULL(bs.N,0) AS varchar(10)),
    /* Matched on ACCOUNT NUMBER, never name. */
    CASE WHEN cust.CustomerAccountNumber = 'OLIVER' THEN 'Y' ELSE '' END,
    CAST(sor.DocumentStatusID AS varchar(10)),
    ISNULL(CONVERT(varchar(10), sor.DocumentDate, 23),''),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(sor.DocumentCreatedBy,''), CHAR(9),' '), CHAR(13),' '), CHAR(10),' ')))
FROM        S200_LIVE.dbo.SOPOrderReturn      sor
INNER JOIN  S200_LIVE.dbo.SOPOrderReturnLine  sorl ON sorl.SOPOrderReturnID    = sor.SOPOrderReturnID
LEFT  JOIN  S200_LIVE.dbo.SLCustomerAccount   cust ON cust.SLCustomerAccountID = sor.CustomerID
LEFT  JOIN  @b2bso                            bs   ON bs.Company = 'TIBARD' AND bs.SOPOrderReturnID = sor.SOPOrderReturnID
WHERE
        sor.DocumentTypeID   = 0                  -- sales orders, not returns
    AND sor.DocumentStatusID IN (0,1)             -- live and on hold
    AND sorl.LineTypeID      = 0                  -- product lines only; no free text, no comments

UNION ALL

/* -------------------------------------------------------- OLIVER HARVEY -- */
SELECT
    'OLIVER HARVEY',
    'OH-SO-' + CAST(sorl.SOPOrderReturnLineID AS varchar(20)),
    ISNULL(LTRIM(RTRIM(sorl.ItemCode)),''),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(sorl.ItemDescription,''), CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))),
    sor.DocumentNo,
    sorl.PrintSequenceNumber,
    ISNULL(cust.CustomerAccountNumber,''),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(cust.CustomerAccountName,''), CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(sor.CustomerDocumentNo,''), CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))),
    CONVERT(varchar(10), COALESCE(sor.PromisedDeliveryDate,  sorl.PromisedDeliveryDate,
                                  sor.RequestedDeliveryDate, sorl.RequestedDeliveryDate), 23),
    CAST(sorl.StockUnitLineQuantity AS decimal(18,2)),
    CAST(ISNULL(sorl.StockUnitDespRcptQuantity,0) AS decimal(18,2)),
    CAST(sorl.StockUnitLineQuantity - ISNULL(sorl.StockUnitDespRcptQuantity,0) AS decimal(18,2)),
    CAST(ISNULL(sorl.StockUnitAllocatedQuantity,0) AS decimal(18,2)),
    ISNULL(CAST(sorl.SOPOrderFulfilmentMethodID AS varchar(10)),''),
    ISNULL(bs.FirstDoc,''),
    CAST(ISNULL(bs.N,0) AS varchar(10)),
    CASE WHEN cust.CustomerAccountNumber = 'TIB003' THEN 'Y' ELSE '' END,
    CAST(sor.DocumentStatusID AS varchar(10)),
    ISNULL(CONVERT(varchar(10), sor.DocumentDate, 23),''),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(sor.DocumentCreatedBy,''), CHAR(9),' '), CHAR(13),' '), CHAR(10),' ')))
FROM        OliverHarveyLive.dbo.SOPOrderReturn      sor
INNER JOIN  OliverHarveyLive.dbo.SOPOrderReturnLine  sorl ON sorl.SOPOrderReturnID    = sor.SOPOrderReturnID
LEFT  JOIN  OliverHarveyLive.dbo.SLCustomerAccount   cust ON cust.SLCustomerAccountID = sor.CustomerID
LEFT  JOIN  @b2bso                                   bs   ON bs.Company = 'OLIVER HARVEY' AND bs.SOPOrderReturnID = sor.SOPOrderReturnID
WHERE
        sor.DocumentTypeID   = 0
    AND sor.DocumentStatusID IN (0,1)
    AND sorl.LineTypeID      = 0
OPTION (RECOMPILE);

/* ------------------------------------------------------------------------
   2b. CMT rows: the comment lines (LineTypeID 3) on the same live or held
   sales orders, both companies. Sage keeps the comment text in
   ItemDescription; tabs and line breaks become spaces so the paste survives
   Excel. Free text item lines (LineTypeID 1) are not comments and stay out.
   ------------------------------------------------------------------------ */
INSERT INTO @cmt (Company, RowKey, Description, DocNo, LineSeq, DocStatusID)
SELECT
    'TIBARD',
    'TIB-CMT-' + CAST(sorl.SOPOrderReturnLineID AS varchar(20)),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(sorl.ItemDescription,''), CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))),
    sor.DocumentNo,
    sorl.PrintSequenceNumber,
    CAST(sor.DocumentStatusID AS varchar(10))
FROM        S200_LIVE.dbo.SOPOrderReturn      sor
INNER JOIN  S200_LIVE.dbo.SOPOrderReturnLine  sorl ON sorl.SOPOrderReturnID = sor.SOPOrderReturnID
WHERE
        sor.DocumentTypeID   = 0
    AND sor.DocumentStatusID IN (0,1)
    AND sorl.LineTypeID      = 3                  -- comment lines only
    AND LTRIM(RTRIM(ISNULL(sorl.ItemDescription,''))) <> ''

UNION ALL

SELECT
    'OLIVER HARVEY',
    'OH-CMT-' + CAST(sorl.SOPOrderReturnLineID AS varchar(20)),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(sorl.ItemDescription,''), CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))),
    sor.DocumentNo,
    sorl.PrintSequenceNumber,
    CAST(sor.DocumentStatusID AS varchar(10))
FROM        OliverHarveyLive.dbo.SOPOrderReturn      sor
INNER JOIN  OliverHarveyLive.dbo.SOPOrderReturnLine  sorl ON sorl.SOPOrderReturnID = sor.SOPOrderReturnID
WHERE
        sor.DocumentTypeID   = 0
    AND sor.DocumentStatusID IN (0,1)
    AND sorl.LineTypeID      = 3
    AND LTRIM(RTRIM(ISNULL(sorl.ItemDescription,''))) <> ''
OPTION (RECOMPILE);

/* ------------------------------------------------------------------------
   3. PO rows: live purchase order lines with quantity outstanding. POP lines
   carry no delivery date of their own - the header's requested date is used;
   blank when the buyer never entered one (the app shows "no date").
   ------------------------------------------------------------------------ */
INSERT INTO @po (Company, RowKey, ProductCode, Description, DocNo, LineSeq, AccountNo, AccountName,
                 Reference, DateISO, Qty1, Qty2, Qty3, Qty4, FulfilMethodID, LinkID, B2BCount,
                 Intercompany, DocStatusID, OrderDate, EnteredBy)
/* ---------------------------------------------------------------- TIBARD -- */
SELECT
    'TIBARD',
    'TIB-PO-' + CAST(porl.POPOrderReturnLineID AS varchar(20)),
    ISNULL(LTRIM(RTRIM(porl.ItemCode)),''),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(porl.ItemDescription,''), CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))),
    por.DocumentNo,
    porl.PrintSequenceNumber,
    ISNULL(supp.SupplierAccountNumber,''),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(supp.SupplierAccountName,''), CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(por.SupplierDocumentNo,''), CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))),
    ISNULL(CONVERT(varchar(10), por.RequestedDeliveryDate, 23),''),
    CAST(porl.StockUnitLineQuantity AS decimal(18,2)),
    CAST(ISNULL(porl.StockUnitRcptRtnQuantity,0) AS decimal(18,2)),
    CAST(porl.StockUnitLineQuantity - ISNULL(porl.StockUnitRcptRtnQuantity,0) AS decimal(18,2)),
    NULL,
    '',
    ISNULL(bp.FirstDoc,''),
    CAST(ISNULL(bp.N,0) AS varchar(10)),
    /* Tibard does not currently buy from OH. If it ever does, add that
       supplier account number here the same way as TIB001 below. */
    '',
    CAST(por.DocumentStatusID AS varchar(10)),
    ISNULL(CONVERT(varchar(10), por.DocumentDate, 23),''),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(por.DocumentCreatedBy,''), CHAR(9),' '), CHAR(13),' '), CHAR(10),' ')))
FROM        S200_LIVE.dbo.POPOrderReturn      por
INNER JOIN  S200_LIVE.dbo.POPOrderReturnLine  porl ON porl.POPOrderReturnID    = por.POPOrderReturnID
LEFT  JOIN  S200_LIVE.dbo.PLSupplierAccount   supp ON supp.PLSupplierAccountID = por.SupplierID
LEFT  JOIN  @b2bpo                            bp   ON bp.Company = 'TIBARD' AND bp.POPOrderReturnID = por.POPOrderReturnID
WHERE
        por.DocumentTypeID   = 0                  -- purchase orders, not returns
    AND por.DocumentStatusID IN (0,1)             -- live and on hold
    AND porl.LineTypeID      = 0                  -- product lines only
    AND (porl.StockUnitLineQuantity - ISNULL(porl.StockUnitRcptRtnQuantity,0)) > 0

UNION ALL

/* -------------------------------------------------------- OLIVER HARVEY -- */
SELECT
    'OLIVER HARVEY',
    'OH-PO-' + CAST(porl.POPOrderReturnLineID AS varchar(20)),
    ISNULL(LTRIM(RTRIM(porl.ItemCode)),''),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(porl.ItemDescription,''), CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))),
    por.DocumentNo,
    porl.PrintSequenceNumber,
    ISNULL(supp.SupplierAccountNumber,''),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(supp.SupplierAccountName,''), CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(por.SupplierDocumentNo,''), CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))),
    ISNULL(CONVERT(varchar(10), por.RequestedDeliveryDate, 23),''),
    CAST(porl.StockUnitLineQuantity AS decimal(18,2)),
    CAST(ISNULL(porl.StockUnitRcptRtnQuantity,0) AS decimal(18,2)),
    CAST(porl.StockUnitLineQuantity - ISNULL(porl.StockUnitRcptRtnQuantity,0) AS decimal(18,2)),
    NULL,
    '',
    ISNULL(bp.FirstDoc,''),
    CAST(ISNULL(bp.N,0) AS varchar(10)),
    /* OH buying from Tibard. TIB001 = "Tibard Ltd" in OH's purchase ledger. */
    CASE WHEN supp.SupplierAccountNumber = 'TIB001' THEN 'Y' ELSE '' END,
    CAST(por.DocumentStatusID AS varchar(10)),
    ISNULL(CONVERT(varchar(10), por.DocumentDate, 23),''),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(por.DocumentCreatedBy,''), CHAR(9),' '), CHAR(13),' '), CHAR(10),' ')))
FROM        OliverHarveyLive.dbo.POPOrderReturn      por
INNER JOIN  OliverHarveyLive.dbo.POPOrderReturnLine  porl ON porl.POPOrderReturnID    = por.POPOrderReturnID
LEFT  JOIN  OliverHarveyLive.dbo.PLSupplierAccount   supp ON supp.PLSupplierAccountID = por.SupplierID
LEFT  JOIN  @b2bpo                                   bp   ON bp.Company = 'OLIVER HARVEY' AND bp.POPOrderReturnID = por.POPOrderReturnID
WHERE
        por.DocumentTypeID   = 0
    AND por.DocumentStatusID IN (0,1)
    AND porl.LineTypeID      = 0
    AND (porl.StockUnitLineQuantity - ISNULL(porl.StockUnitRcptRtnQuantity,0)) > 0
OPTION (RECOMPILE);

/* ------------------------------------------------------------------------
   4. Which codes get a STK row. The stock file is far too big to send whole:
   every code on a live SO or PO line, in BOTH companies (an OH order for a
   Tibard-made garment is covered from Tibard's shelf, so the tracker needs
   the Tibard stock figure even when no Tibard order carries the code), plus
   the BULK item for each code (container stock sits under BULKCJ0193MM01,
   not CJ0193MM01). A company only gets the row when it has the stock record
   or the code is on one of its own orders, so there are no empty rows for
   codes a company never trades.
   ------------------------------------------------------------------------ */
INSERT INTO @codes0 (Company, ProductCode)
SELECT Company, ProductCode FROM @so
UNION
SELECT Company, ProductCode FROM @po;

INSERT INTO @codes (Company, ProductCode)
SELECT 'TIBARD', a.ProductCode
FROM  (SELECT ProductCode FROM @codes0
       UNION
       SELECT 'BULK' + ProductCode FROM @codes0 WHERE ProductCode NOT LIKE 'BULK%') a
WHERE  EXISTS (SELECT 1 FROM @codes0 c WHERE c.Company = 'TIBARD' AND c.ProductCode = a.ProductCode)
   OR  EXISTS (SELECT 1 FROM S200_LIVE.dbo.StockItem si WHERE si.Code = a.ProductCode)
UNION ALL
SELECT 'OLIVER HARVEY', a.ProductCode
FROM  (SELECT ProductCode FROM @codes0
       UNION
       SELECT 'BULK' + ProductCode FROM @codes0 WHERE ProductCode NOT LIKE 'BULK%') a
WHERE  EXISTS (SELECT 1 FROM @codes0 c WHERE c.Company = 'OLIVER HARVEY' AND c.ProductCode = a.ProductCode)
   OR  EXISTS (SELECT 1 FROM OliverHarveyLive.dbo.StockItem si WHERE si.Code = a.ProductCode)
OPTION (RECOMPILE);

/* ------------------------------------------------------------------------
   5. Back-to-back history per code: generated POs in the last 12 months
   that contained it. > 0 marks a code normally ordered back to back.
   ------------------------------------------------------------------------ */
INSERT INTO @b2b12 (Company, ProductCode, N)
SELECT 'TIBARD', LTRIM(RTRIM(lpl.ItemCode)), COUNT(DISTINCT b.POPOrderReturnID)
FROM   S200_LIVE.dbo.MMS123BackToBackPO  b
JOIN   S200_LIVE.dbo.POPOrderReturn      lpo ON lpo.POPOrderReturnID = b.POPOrderReturnID
JOIN   S200_LIVE.dbo.POPOrderReturnLine  lpl ON lpl.POPOrderReturnID = lpo.POPOrderReturnID
WHERE  lpo.DocumentDate >= DATEADD(month, -12, GETDATE())
  AND  lpl.ItemCode IS NOT NULL
GROUP BY LTRIM(RTRIM(lpl.ItemCode))
UNION ALL
SELECT 'OLIVER HARVEY', LTRIM(RTRIM(lpl.ItemCode)), COUNT(DISTINCT b.POPOrderReturnID)
FROM   OliverHarveyLive.dbo.MMS123BackToBackPO  b
JOIN   OliverHarveyLive.dbo.POPOrderReturn      lpo ON lpo.POPOrderReturnID = b.POPOrderReturnID
JOIN   OliverHarveyLive.dbo.POPOrderReturnLine  lpl ON lpl.POPOrderReturnID = lpo.POPOrderReturnID
WHERE  lpo.DocumentDate >= DATEADD(month, -12, GETDATE())
  AND  lpl.ItemCode IS NOT NULL
GROUP BY LTRIM(RTRIM(lpl.ItemCode));

/* ------------------------------------------------------------------------
   6. STK rows: the stock record, figures and classification per code per
   company. Stock figures sum every warehouse the item sits in.
   ------------------------------------------------------------------------ */
INSERT INTO @stk (Company, RowKey, ProductCode, Description, AccountNo, AccountName, Reference,
                  Qty1, Qty2, Qty3, Qty4, FulfilMethodID, Manufacturer, StockHeld, WebsiteOH,
                  ProductGroup, BOMItemTypeID, LeadTime, LeadTimeUnitID, MOQ, B2B12m)
/* ---------------------------------------------------------------- TIBARD -- */
SELECT
    c.Company,
    'TIB-STK-' + c.ProductCode,
    c.ProductCode,
    CASE WHEN si.Code IS NULL THEN 'CODE NOT IN STOCK FILE'
         ELSE LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(si.Name,''), CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))) END,
    ISNULL(ps.SupplierAccountNumber,''),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(ps.SupplierAccountName,''), CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(ps.SupplierStockCode,''), CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))),
    CAST(ISNULL(wh.InStock,0)           AS decimal(18,2)),
    CAST(ISNULL(wh.Allocated,0)         AS decimal(18,2)),
    CAST(ISNULL(si.FreeStockQuantity,0) AS decimal(18,2)),
    CAST(ISNULL(wh.OnPO,0)              AS decimal(18,2)),
    ISNULL(CAST(si.SOPOrderFulfilmentMethodID AS varchar(10)),''),
    LTRIM(RTRIM(ISNULL(si.Manufacturer,''))),
    /* Tibard: Stock Held is its own analysis code. */
    CASE WHEN ISNULL(si.AnalysisCode3,'') = 'Yes' THEN 'Yes' ELSE '' END,
    '',
    ISNULL(pg.Code,''),
    ISNULL(CAST(si.BOMItemTypeID AS varchar(10)),''),
    ISNULL(CAST(ps.LeadTime AS varchar(10)),''),
    ISNULL(CAST(ps.LeadTimeUnitID AS varchar(10)),''),
    CAST(ps.MinimumOrderQuantity AS decimal(18,2)),
    ISNULL(CAST(bb.N AS varchar(10)),'0')
FROM        @codes c
LEFT  JOIN  S200_LIVE.dbo.StockItem     si ON si.Code             = c.ProductCode
LEFT  JOIN  @b2b12                      bb ON bb.Company = 'TIBARD' AND bb.ProductCode = c.ProductCode
LEFT  JOIN  S200_LIVE.dbo.ProductGroup  pg ON pg.ProductGroupID   = si.ProductGroupID
OUTER APPLY (
    SELECT  SUM(wi.ConfirmedQtyInStock)                                 AS InStock,
            SUM(ISNULL(wi.QuantityAllocatedSOP,0) + ISNULL(wi.QuantityAllocatedStock,0)
              + ISNULL(wi.QuantityAllocatedBOM,0))                      AS Allocated,
            SUM(wi.QuantityOnPOPOrder)                                  AS OnPO
    FROM    S200_LIVE.dbo.WarehouseItem wi
    WHERE   wi.ItemID = si.ItemID
) wh
/* Preferred supplier first; otherwise the one most recently ordered from. */
OUTER APPLY (
    SELECT TOP 1 supp.SupplierAccountNumber, supp.SupplierAccountName,
           sis.SupplierStockCode, sis.LeadTime, sis.LeadTimeUnitID, sis.MinimumOrderQuantity
    FROM   S200_LIVE.dbo.StockItemSupplier sis
    LEFT JOIN S200_LIVE.dbo.PLSupplierAccount supp ON supp.PLSupplierAccountID = sis.SupplierID
    WHERE  sis.ItemID = si.ItemID
    ORDER BY sis.Preferred DESC, sis.DateLastOrder DESC
) ps
WHERE c.Company = 'TIBARD'

UNION ALL

/* -------------------------------------------------------- OLIVER HARVEY -- */
SELECT
    c.Company,
    'OH-STK-' + c.ProductCode,
    c.ProductCode,
    CASE WHEN si.Code IS NULL THEN 'CODE NOT IN STOCK FILE'
         ELSE LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(si.Name,''), CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))) END,
    ISNULL(ps.SupplierAccountNumber,''),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(ps.SupplierAccountName,''), CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(ps.SupplierStockCode,''), CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))),
    CAST(ISNULL(wh.InStock,0)           AS decimal(18,2)),
    CAST(ISNULL(wh.Allocated,0)         AS decimal(18,2)),
    CAST(ISNULL(si.FreeStockQuantity,0) AS decimal(18,2)),
    CAST(ISNULL(wh.OnPO,0)              AS decimal(18,2)),
    ISNULL(CAST(si.SOPOrderFulfilmentMethodID AS varchar(10)),''),
    LTRIM(RTRIM(ISNULL(si.Manufacturer,''))),
    /* OH: Stock Held OR Website OR stock held at Tibard under the same code.
       Same three-way rule as the special makes query. */
    CASE WHEN ISNULL(si.AnalysisCode3,'')  = 'Yes'
           OR ISNULL(si.AnalysisCode7,'')  = 'Yes'
           OR ISNULL(tsi.AnalysisCode3,'') = 'Yes' THEN 'Yes' ELSE '' END,
    CASE WHEN ISNULL(si.AnalysisCode7,'')  = 'Yes' THEN 'Yes' ELSE '' END,
    ISNULL(pg.Code,''),
    ISNULL(CAST(si.BOMItemTypeID AS varchar(10)),''),
    ISNULL(CAST(ps.LeadTime AS varchar(10)),''),
    ISNULL(CAST(ps.LeadTimeUnitID AS varchar(10)),''),
    CAST(ps.MinimumOrderQuantity AS decimal(18,2)),
    ISNULL(CAST(bb.N AS varchar(10)),'0')
FROM        @codes c
LEFT  JOIN  OliverHarveyLive.dbo.StockItem     si  ON si.Code            = c.ProductCode
LEFT  JOIN  @b2b12                             bb  ON bb.Company = 'OLIVER HARVEY' AND bb.ProductCode = c.ProductCode
LEFT  JOIN  S200_LIVE.dbo.StockItem            tsi ON tsi.Code           = c.ProductCode   -- Tibard's record, same code
LEFT  JOIN  OliverHarveyLive.dbo.ProductGroup  pg  ON pg.ProductGroupID  = si.ProductGroupID
OUTER APPLY (
    SELECT  SUM(wi.ConfirmedQtyInStock)                                 AS InStock,
            SUM(ISNULL(wi.QuantityAllocatedSOP,0) + ISNULL(wi.QuantityAllocatedStock,0)
              + ISNULL(wi.QuantityAllocatedBOM,0))                      AS Allocated,
            SUM(wi.QuantityOnPOPOrder)                                  AS OnPO
    FROM    OliverHarveyLive.dbo.WarehouseItem wi
    WHERE   wi.ItemID = si.ItemID
) wh
OUTER APPLY (
    SELECT TOP 1 supp.SupplierAccountNumber, supp.SupplierAccountName,
           sis.SupplierStockCode, sis.LeadTime, sis.LeadTimeUnitID, sis.MinimumOrderQuantity
    FROM   OliverHarveyLive.dbo.StockItemSupplier sis
    LEFT JOIN OliverHarveyLive.dbo.PLSupplierAccount supp ON supp.PLSupplierAccountID = sis.SupplierID
    WHERE  sis.ItemID = si.ItemID
    ORDER BY sis.Preferred DESC, sis.DateLastOrder DESC
) ps
WHERE c.Company = 'OLIVER HARVEY'
OPTION (RECOMPILE);

/* ------------------------------------------------------------------------
   7. One shape for all four row types - THE RESULT. Columns a row type does
   not use are blank (text) or NULL (numbers); both paste as an empty cell.
   ------------------------------------------------------------------------ */
SELECT  RowType, Company, RowKey, ProductCode, Description, DocNo, LineSeq,
        AccountNo, AccountName, Reference, DateISO, Qty1, Qty2, Qty3, Qty4,
        FulfilMethodID, LinkID, B2BCount, Intercompany,
        Manufacturer, StockHeld, WebsiteOH, ProductGroup,
        BOMItemTypeID, LeadTime, LeadTimeUnitID, MOQ, DocStatusID, OrderDate, EnteredBy
FROM (
    SELECT 'SO' AS RowType, Company, RowKey, ProductCode, Description, DocNo, LineSeq,
           AccountNo, AccountName, Reference, DateISO, Qty1, Qty2, Qty3, Qty4,
           FulfilMethodID, LinkID, B2BCount, Intercompany,
           '' AS Manufacturer, '' AS StockHeld, '' AS WebsiteOH, '' AS ProductGroup,
           '' AS BOMItemTypeID, '' AS LeadTime, '' AS LeadTimeUnitID,
           CAST(NULL AS decimal(18,2)) AS MOQ, DocStatusID, OrderDate, EnteredBy
    FROM   @so
    UNION ALL
    SELECT 'CMT', Company, RowKey, '' AS ProductCode, Description, DocNo, LineSeq,
           '' AS AccountNo, '' AS AccountName, '' AS Reference, '' AS DateISO,
           CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)), CAST(NULL AS decimal(18,2)),
           '' AS FulfilMethodID, '' AS LinkID, '' AS B2BCount, '' AS Intercompany,
           '', '', '', '', '', '', '', CAST(NULL AS decimal(18,2)), DocStatusID, '' AS OrderDate, '' AS EnteredBy
    FROM   @cmt
    UNION ALL
    SELECT 'PO', Company, RowKey, ProductCode, Description, DocNo, LineSeq,
           AccountNo, AccountName, Reference, DateISO, Qty1, Qty2, Qty3, Qty4,
           FulfilMethodID, LinkID, B2BCount, Intercompany,
           '', '', '', '', '', '', '', CAST(NULL AS decimal(18,2)), DocStatusID, OrderDate, EnteredBy
    FROM   @po
    UNION ALL
    SELECT 'STK', Company, RowKey, ProductCode, Description, '' AS DocNo, CAST(NULL AS int) AS LineSeq,
           AccountNo, AccountName, Reference, '' AS DateISO, Qty1, Qty2, Qty3, Qty4,
           FulfilMethodID, '' AS LinkID, B2B12m AS B2BCount, '' AS Intercompany,
           Manufacturer, StockHeld, WebsiteOH, ProductGroup,
           BOMItemTypeID, LeadTime, LeadTimeUnitID, MOQ, '' AS DocStatusID, '' AS OrderDate, '' AS EnteredBy
    FROM   @stk
) allrows
ORDER BY CASE RowType WHEN 'SO' THEN 1 WHEN 'CMT' THEN 2 WHEN 'PO' THEN 3 ELSE 4 END,
         Company, ProductCode, DocNo, LineSeq
OPTION (RECOMPILE);
