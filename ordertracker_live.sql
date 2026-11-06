/* ============================================================================
   ORDER TRACKER v1 - LIVE SALES ORDERS, PURCHASE ORDERS AND STOCK
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

   ONE PASTE. Copy columns A to AD, all rows, without the header. Three row
   types share one column layout and the app splits them on column A:

     SO   one row per product line on a live or held sales order
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
     C  RowKey          TIB-SO-<lineid>, OH-PO-<lineid>, TIB-STK-<code> ...
     D  ProductCode
     E  Description     SO/PO: line description. STK: stock item name, or
                        CODE NOT IN STOCK FILE
     F  DocNo           SO: sales order no. PO: purchase order no
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
     AB DocStatusID     SO/PO: order status. 0 = live, 1 = on hold
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
   ========================================================================= */

WITH so AS (

/* ---------------------------------------------------------------- TIBARD -- */
SELECT
    'TIBARD'                                                    AS Company,
    'TIB-SO-' + CAST(sorl.SOPOrderReturnLineID AS varchar(20))  AS RowKey,
    LTRIM(RTRIM(sorl.ItemCode))                                 AS ProductCode,
    /* Tabs and line breaks stripped everywhere - the app splits pasted rows on TAB. */
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(sorl.ItemDescription,''),
        CHAR(9),' '), CHAR(13),' '), CHAR(10),' ')))            AS Description,
    sor.DocumentNo                                              AS DocNo,
    sorl.PrintSequenceNumber                                    AS LineSeq,
    ISNULL(cust.CustomerAccountNumber,'')                       AS AccountNo,
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(cust.CustomerAccountName,''),
        CHAR(9),' '), CHAR(13),' '), CHAR(10),' ')))            AS AccountName,
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(sor.CustomerDocumentNo,''),
        CHAR(9),' '), CHAR(13),' '), CHAR(10),' ')))            AS Reference,
    /* Header before line - see PROMISED DATE above. Text yyyy-mm-dd so the
       browser never has to guess dd/mm against mm/dd. */
    CONVERT(varchar(10), COALESCE(sor.PromisedDeliveryDate,  sorl.PromisedDeliveryDate,
                                  sor.RequestedDeliveryDate, sorl.RequestedDeliveryDate), 23)
                                                                AS DateISO,
    CAST(sorl.StockUnitLineQuantity AS decimal(18,2))           AS Qty1,
    CAST(ISNULL(sorl.StockUnitDespRcptQuantity,0) AS decimal(18,2))
                                                                AS Qty2,
    CAST(sorl.StockUnitLineQuantity - ISNULL(sorl.StockUnitDespRcptQuantity,0) AS decimal(18,2))
                                                                AS Qty3,
    CAST(ISNULL(sorl.StockUnitAllocatedQuantity,0) AS decimal(18,2))
                                                                AS Qty4,
    ISNULL(CAST(sorl.SOPOrderFulfilmentMethodID AS varchar(10)),'')
                                                                AS FulfilMethodID,
    /* Back to back: the PO(s) Sage generated from this order, via
       MMS123BackToBackPO. Q = first linked PO number, R = how many. Blank Q
       on a code the buyer normally orders back to back (STK column R > 0) is
       the miss this app exists to catch. */
    ISNULL((SELECT TOP 1 lpo.DocumentNo
            FROM  S200_LIVE.dbo.MMS123BackToBackPO b
            JOIN  S200_LIVE.dbo.POPOrderReturn lpo ON lpo.POPOrderReturnID = b.POPOrderReturnID
            WHERE b.SOPOrderReturnID = sor.SOPOrderReturnID
            ORDER BY lpo.DocumentNo),'')                        AS LinkID,
    CAST((SELECT COUNT(*) FROM S200_LIVE.dbo.MMS123BackToBackPO b
          WHERE b.SOPOrderReturnID = sor.SOPOrderReturnID) AS varchar(10))
                                                                AS B2BCount,
    /* Matched on ACCOUNT NUMBER, never name - same reason as special makes. */
    CASE WHEN cust.CustomerAccountNumber = 'OLIVER' THEN 'Y' ELSE '' END
                                                                AS Intercompany,
    CAST(sor.DocumentStatusID AS varchar(10))                   AS DocStatusID,
    /* Date the order was raised in Sage - see ORDER DATE above. */
    ISNULL(CONVERT(varchar(10), sor.DocumentDate, 23),'')       AS OrderDate,
    /* Sage user who raised the order - see ENTERED BY above. */
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(sor.DocumentCreatedBy,''),
        CHAR(9),' '), CHAR(13),' '), CHAR(10),' ')))            AS EnteredBy
FROM        S200_LIVE.dbo.SOPOrderReturn      sor
INNER JOIN  S200_LIVE.dbo.SOPOrderReturnLine  sorl ON sorl.SOPOrderReturnID    = sor.SOPOrderReturnID
LEFT  JOIN  S200_LIVE.dbo.SLCustomerAccount   cust ON cust.SLCustomerAccountID = sor.CustomerID
WHERE
        sor.DocumentTypeID   = 0                  -- sales orders, not returns
    AND sor.DocumentStatusID IN (0,1)             -- live and on hold
    AND sorl.LineTypeID      = 0                  -- product lines only; no free text, no comments
    /* Every product line, despatched ones too (Qty3 = 0): the tracker shows
       the whole order. POtracker ignored this, so the paste suits both. */

UNION ALL

/* -------------------------------------------------------- OLIVER HARVEY -- */
SELECT
    'OLIVER HARVEY',
    'OH-SO-' + CAST(sorl.SOPOrderReturnLineID AS varchar(20)),
    LTRIM(RTRIM(sorl.ItemCode)),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(sorl.ItemDescription,''),
        CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))),
    sor.DocumentNo,
    sorl.PrintSequenceNumber,
    ISNULL(cust.CustomerAccountNumber,''),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(cust.CustomerAccountName,''),
        CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(sor.CustomerDocumentNo,''),
        CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))),
    CONVERT(varchar(10), COALESCE(sor.PromisedDeliveryDate,  sorl.PromisedDeliveryDate,
                                  sor.RequestedDeliveryDate, sorl.RequestedDeliveryDate), 23),
    CAST(sorl.StockUnitLineQuantity AS decimal(18,2)),
    CAST(ISNULL(sorl.StockUnitDespRcptQuantity,0) AS decimal(18,2)),
    CAST(sorl.StockUnitLineQuantity - ISNULL(sorl.StockUnitDespRcptQuantity,0) AS decimal(18,2)),
    CAST(ISNULL(sorl.StockUnitAllocatedQuantity,0) AS decimal(18,2)),
    ISNULL(CAST(sorl.SOPOrderFulfilmentMethodID AS varchar(10)),''),
    ISNULL((SELECT TOP 1 lpo.DocumentNo
            FROM  OliverHarveyLive.dbo.MMS123BackToBackPO b
            JOIN  OliverHarveyLive.dbo.POPOrderReturn lpo ON lpo.POPOrderReturnID = b.POPOrderReturnID
            WHERE b.SOPOrderReturnID = sor.SOPOrderReturnID
            ORDER BY lpo.DocumentNo),''),
    CAST((SELECT COUNT(*) FROM OliverHarveyLive.dbo.MMS123BackToBackPO b
          WHERE b.SOPOrderReturnID = sor.SOPOrderReturnID) AS varchar(10)),
    CASE WHEN cust.CustomerAccountNumber = 'TIB003' THEN 'Y' ELSE '' END,
    CAST(sor.DocumentStatusID AS varchar(10)),
    ISNULL(CONVERT(varchar(10), sor.DocumentDate, 23),''),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(sor.DocumentCreatedBy,''),
        CHAR(9),' '), CHAR(13),' '), CHAR(10),' ')))
FROM        OliverHarveyLive.dbo.SOPOrderReturn      sor
INNER JOIN  OliverHarveyLive.dbo.SOPOrderReturnLine  sorl ON sorl.SOPOrderReturnID    = sor.SOPOrderReturnID
LEFT  JOIN  OliverHarveyLive.dbo.SLCustomerAccount   cust ON cust.SLCustomerAccountID = sor.CustomerID
WHERE
        sor.DocumentTypeID   = 0
    AND sor.DocumentStatusID IN (0,1)
    AND sorl.LineTypeID      = 0
    /* every product line, despatched ones too - see the Tibard block */
),

po AS (

/* ---------------------------------------------------------------- TIBARD -- */
SELECT
    'TIBARD'                                                    AS Company,
    'TIB-PO-' + CAST(porl.POPOrderReturnLineID AS varchar(20))  AS RowKey,
    LTRIM(RTRIM(porl.ItemCode))                                 AS ProductCode,
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(porl.ItemDescription,''),
        CHAR(9),' '), CHAR(13),' '), CHAR(10),' ')))            AS Description,
    por.DocumentNo                                              AS DocNo,
    porl.PrintSequenceNumber                                    AS LineSeq,
    ISNULL(supp.SupplierAccountNumber,'')                       AS AccountNo,
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(supp.SupplierAccountName,''),
        CHAR(9),' '), CHAR(13),' '), CHAR(10),' ')))            AS AccountName,
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(por.SupplierDocumentNo,''),
        CHAR(9),' '), CHAR(13),' '), CHAR(10),' ')))            AS Reference,
    /* POP lines carry no delivery date of their own - header only. Blank if
       the buyer never entered one; the app treats blank as "date unknown". */
    ISNULL(CONVERT(varchar(10), por.RequestedDeliveryDate, 23),'')
                                                                AS DateISO,
    CAST(porl.StockUnitLineQuantity AS decimal(18,2))           AS Qty1,
    CAST(ISNULL(porl.StockUnitRcptRtnQuantity,0) AS decimal(18,2))
                                                                AS Qty2,
    CAST(porl.StockUnitLineQuantity - ISNULL(porl.StockUnitRcptRtnQuantity,0) AS decimal(18,2))
                                                                AS Qty3,
    CAST(NULL AS decimal(18,2))                                 AS Qty4,
    ''                                                          AS FulfilMethodID,
    /* The sales order this PO was generated from, via MMS123BackToBackPO. */
    ISNULL((SELECT TOP 1 lso.DocumentNo
            FROM  S200_LIVE.dbo.MMS123BackToBackPO b
            JOIN  S200_LIVE.dbo.SOPOrderReturn lso ON lso.SOPOrderReturnID = b.SOPOrderReturnID
            WHERE b.POPOrderReturnID = por.POPOrderReturnID
            ORDER BY lso.DocumentNo),'')                        AS LinkID,
    CAST((SELECT COUNT(*) FROM S200_LIVE.dbo.MMS123BackToBackPO b
          WHERE b.POPOrderReturnID = por.POPOrderReturnID) AS varchar(10))
                                                                AS B2BCount,
    /* Tibard does not currently buy from OH. If it ever does, add that
       supplier account number here the same way as TIB001 below. */
    ''                                                          AS Intercompany,
    CAST(por.DocumentStatusID AS varchar(10))                   AS DocStatusID,
    ISNULL(CONVERT(varchar(10), por.DocumentDate, 23),'')       AS OrderDate,
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(por.DocumentCreatedBy,''),
        CHAR(9),' '), CHAR(13),' '), CHAR(10),' ')))            AS EnteredBy
FROM        S200_LIVE.dbo.POPOrderReturn      por
INNER JOIN  S200_LIVE.dbo.POPOrderReturnLine  porl ON porl.POPOrderReturnID    = por.POPOrderReturnID
LEFT  JOIN  S200_LIVE.dbo.PLSupplierAccount   supp ON supp.PLSupplierAccountID = por.SupplierID
WHERE
        por.DocumentTypeID   = 0                  -- purchase orders, not returns
    AND por.DocumentStatusID IN (0,1)             -- live and on hold (see STATUS FILTERS)
    AND porl.LineTypeID      = 0                  -- product lines only
    AND (porl.StockUnitLineQuantity - ISNULL(porl.StockUnitRcptRtnQuantity,0)) > 0

UNION ALL

/* -------------------------------------------------------- OLIVER HARVEY -- */
SELECT
    'OLIVER HARVEY',
    'OH-PO-' + CAST(porl.POPOrderReturnLineID AS varchar(20)),
    LTRIM(RTRIM(porl.ItemCode)),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(porl.ItemDescription,''),
        CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))),
    por.DocumentNo,
    porl.PrintSequenceNumber,
    ISNULL(supp.SupplierAccountNumber,''),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(supp.SupplierAccountName,''),
        CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(por.SupplierDocumentNo,''),
        CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))),
    ISNULL(CONVERT(varchar(10), por.RequestedDeliveryDate, 23),''),
    CAST(porl.StockUnitLineQuantity AS decimal(18,2)),
    CAST(ISNULL(porl.StockUnitRcptRtnQuantity,0) AS decimal(18,2)),
    CAST(porl.StockUnitLineQuantity - ISNULL(porl.StockUnitRcptRtnQuantity,0) AS decimal(18,2)),
    CAST(NULL AS decimal(18,2)),
    '',
    ISNULL((SELECT TOP 1 lso.DocumentNo
            FROM  OliverHarveyLive.dbo.MMS123BackToBackPO b
            JOIN  OliverHarveyLive.dbo.SOPOrderReturn lso ON lso.SOPOrderReturnID = b.SOPOrderReturnID
            WHERE b.POPOrderReturnID = por.POPOrderReturnID
            ORDER BY lso.DocumentNo),''),
    CAST((SELECT COUNT(*) FROM OliverHarveyLive.dbo.MMS123BackToBackPO b
          WHERE b.POPOrderReturnID = por.POPOrderReturnID) AS varchar(10)),
    /* OH buying from Tibard. TIB001 = "Tibard Ltd" in OH's purchase ledger. */
    CASE WHEN supp.SupplierAccountNumber = 'TIB001' THEN 'Y' ELSE '' END,
    CAST(por.DocumentStatusID AS varchar(10)),
    ISNULL(CONVERT(varchar(10), por.DocumentDate, 23),''),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(por.DocumentCreatedBy,''),
        CHAR(9),' '), CHAR(13),' '), CHAR(10),' ')))
FROM        OliverHarveyLive.dbo.POPOrderReturn      por
INNER JOIN  OliverHarveyLive.dbo.POPOrderReturnLine  porl ON porl.POPOrderReturnID    = por.POPOrderReturnID
LEFT  JOIN  OliverHarveyLive.dbo.PLSupplierAccount   supp ON supp.PLSupplierAccountID = por.SupplierID
WHERE
        por.DocumentTypeID   = 0
    AND por.DocumentStatusID IN (0,1)
    AND porl.LineTypeID      = 0
    AND (porl.StockUnitLineQuantity - ISNULL(porl.StockUnitRcptRtnQuantity,0)) > 0
),

/* Which codes get a STK row. The stock file is far too big to send whole, so:
   every code on a live SO or PO line, in BOTH companies (an OH order for a
   Tibard-made garment is covered from Tibard's shelf, so the tracker needs
   the Tibard stock figure even when no Tibard order carries the code), plus
   the BULK item for each code (container stock sits under BULKCJ0193MM01,
   not CJ0193MM01). A company only gets the row when it has the stock record
   or the code is on one of its own orders, so there are no empty rows for
   codes a company never trades. POtracker's query lists only the codes on
   that company's own orders - that is the one difference in the STK rows. */
codes0 AS (
    SELECT Company, ProductCode FROM so
    UNION
    SELECT Company, ProductCode FROM po
),
anycode AS (
    SELECT ProductCode FROM codes0
    UNION
    SELECT 'BULK' + ProductCode FROM codes0 WHERE ProductCode NOT LIKE 'BULK%'
),
codes AS (
    SELECT 'TIBARD' AS Company, a.ProductCode
    FROM   anycode a
    WHERE  EXISTS (SELECT 1 FROM codes0 c WHERE c.Company = 'TIBARD' AND c.ProductCode = a.ProductCode)
       OR  EXISTS (SELECT 1 FROM S200_LIVE.dbo.StockItem si WHERE si.Code = a.ProductCode)
    UNION ALL
    SELECT 'OLIVER HARVEY', a.ProductCode
    FROM   anycode a
    WHERE  EXISTS (SELECT 1 FROM codes0 c WHERE c.Company = 'OLIVER HARVEY' AND c.ProductCode = a.ProductCode)
       OR  EXISTS (SELECT 1 FROM OliverHarveyLive.dbo.StockItem si WHERE si.Code = a.ProductCode)
),

/* Back-to-back history per code: how many generated POs in the last 12
   months contained it. Any value > 0 marks a code the buyer normally orders
   back to back, so an SO line for it with no linked PO is a miss. */
b2b12 AS (
    SELECT 'TIBARD' AS Company, LTRIM(RTRIM(lpl.ItemCode)) AS ProductCode,
           COUNT(DISTINCT b.POPOrderReturnID) AS N
    FROM   S200_LIVE.dbo.MMS123BackToBackPO  b
    JOIN   S200_LIVE.dbo.POPOrderReturn      lpo ON lpo.POPOrderReturnID = b.POPOrderReturnID
    JOIN   S200_LIVE.dbo.POPOrderReturnLine  lpl ON lpl.POPOrderReturnID = lpo.POPOrderReturnID
    WHERE  lpo.DocumentDate >= DATEADD(month, -12, GETDATE())
    GROUP BY LTRIM(RTRIM(lpl.ItemCode))
    UNION ALL
    SELECT 'OLIVER HARVEY', LTRIM(RTRIM(lpl.ItemCode)), COUNT(DISTINCT b.POPOrderReturnID)
    FROM   OliverHarveyLive.dbo.MMS123BackToBackPO  b
    JOIN   OliverHarveyLive.dbo.POPOrderReturn      lpo ON lpo.POPOrderReturnID = b.POPOrderReturnID
    JOIN   OliverHarveyLive.dbo.POPOrderReturnLine  lpl ON lpl.POPOrderReturnID = lpo.POPOrderReturnID
    WHERE  lpo.DocumentDate >= DATEADD(month, -12, GETDATE())
    GROUP BY LTRIM(RTRIM(lpl.ItemCode))
),

stk AS (

/* ---------------------------------------------------------------- TIBARD -- */
SELECT
    'TIBARD'                                                    AS Company,
    'TIB-STK-' + c.ProductCode                                  AS RowKey,
    c.ProductCode,
    CASE WHEN si.Code IS NULL THEN 'CODE NOT IN STOCK FILE'
         ELSE LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(si.Name,''),
              CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))) END  AS Description,
    ISNULL(ps.SupplierAccountNumber,'')                         AS AccountNo,
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(ps.SupplierAccountName,''),
        CHAR(9),' '), CHAR(13),' '), CHAR(10),' ')))            AS AccountName,
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(ps.SupplierStockCode,''),
        CHAR(9),' '), CHAR(13),' '), CHAR(10),' ')))            AS Reference,
    CAST(ISNULL(wh.InStock,0)          AS decimal(18,2))        AS Qty1,
    CAST(ISNULL(wh.Allocated,0)        AS decimal(18,2))        AS Qty2,
    CAST(ISNULL(si.FreeStockQuantity,0) AS decimal(18,2))       AS Qty3,
    CAST(ISNULL(wh.OnPO,0)             AS decimal(18,2))        AS Qty4,
    ISNULL(CAST(si.SOPOrderFulfilmentMethodID AS varchar(10)),'')
                                                                AS FulfilMethodID,
    LTRIM(RTRIM(ISNULL(si.Manufacturer,'')))                    AS Manufacturer,
    /* Tibard: Stock Held is its own analysis code. */
    CASE WHEN ISNULL(si.AnalysisCode3,'') = 'Yes' THEN 'Yes' ELSE '' END
                                                                AS StockHeld,
    ''                                                          AS WebsiteOH,
    ISNULL(pg.Code,'')                                          AS ProductGroup,
    ISNULL(CAST(si.BOMItemTypeID AS varchar(10)),'')            AS BOMItemTypeID,
    ISNULL(CAST(ps.LeadTime AS varchar(10)),'')                 AS LeadTime,
    ISNULL(CAST(ps.LeadTimeUnitID AS varchar(10)),'')           AS LeadTimeUnitID,
    CAST(ps.MinimumOrderQuantity AS decimal(18,2))              AS MOQ,
    ISNULL(CAST(bb.N AS varchar(10)),'0')                       AS B2B12m
FROM        codes c
LEFT  JOIN  S200_LIVE.dbo.StockItem     si ON si.Code             = c.ProductCode
LEFT  JOIN  b2b12                       bb ON bb.Company = 'TIBARD' AND bb.ProductCode = c.ProductCode
LEFT  JOIN  S200_LIVE.dbo.ProductGroup  pg ON pg.ProductGroupID   = si.ProductGroupID
/* Stock across every warehouse the item is set up in. */
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
    'OLIVER HARVEY',
    'OH-STK-' + c.ProductCode,
    c.ProductCode,
    CASE WHEN si.Code IS NULL THEN 'CODE NOT IN STOCK FILE'
         ELSE LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(si.Name,''),
              CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))) END,
    ISNULL(ps.SupplierAccountNumber,''),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(ps.SupplierAccountName,''),
        CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(ps.SupplierStockCode,''),
        CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))),
    CAST(ISNULL(wh.InStock,0)          AS decimal(18,2)),
    CAST(ISNULL(wh.Allocated,0)        AS decimal(18,2)),
    CAST(ISNULL(si.FreeStockQuantity,0) AS decimal(18,2)),
    CAST(ISNULL(wh.OnPO,0)             AS decimal(18,2)),
    ISNULL(CAST(si.SOPOrderFulfilmentMethodID AS varchar(10)),''),
    LTRIM(RTRIM(ISNULL(si.Manufacturer,''))),
    /* OH: Stock Held OR Website OR stock held at Tibard under the same code.
       Same three-way rule as the special makes query - do not tidy it to
       match the Tibard block, the two companies use the codes differently. */
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
FROM        codes c
LEFT  JOIN  OliverHarveyLive.dbo.StockItem     si  ON si.Code            = c.ProductCode
LEFT  JOIN  b2b12                              bb  ON bb.Company = 'OLIVER HARVEY' AND bb.ProductCode = c.ProductCode
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
),

/* One shape for all three row types. Columns the row type does not use are
   blank (text) or NULL (numbers) - both arrive as an empty cell when pasted. */
allrows AS (
    SELECT 'SO' AS RowType, Company, RowKey, ProductCode, Description, DocNo, LineSeq,
           AccountNo, AccountName, Reference, DateISO, Qty1, Qty2, Qty3, Qty4,
           FulfilMethodID, LinkID, B2BCount, Intercompany,
           '' AS Manufacturer, '' AS StockHeld, '' AS WebsiteOH, '' AS ProductGroup,
           '' AS BOMItemTypeID, '' AS LeadTime, '' AS LeadTimeUnitID,
           CAST(NULL AS decimal(18,2)) AS MOQ, DocStatusID, OrderDate, EnteredBy
    FROM   so
    UNION ALL
    SELECT 'PO', Company, RowKey, ProductCode, Description, DocNo, LineSeq,
           AccountNo, AccountName, Reference, DateISO, Qty1, Qty2, Qty3, Qty4,
           FulfilMethodID, LinkID, B2BCount, Intercompany,
           '', '', '', '', '', '', '', CAST(NULL AS decimal(18,2)), DocStatusID, OrderDate, EnteredBy
    FROM   po
    UNION ALL
    SELECT 'STK', Company, RowKey, ProductCode, Description, '' AS DocNo, CAST(NULL AS smallint) AS LineSeq,
           AccountNo, AccountName, Reference, '' AS DateISO, Qty1, Qty2, Qty3, Qty4,
           FulfilMethodID, '' AS LinkID, B2B12m AS B2BCount, '' AS Intercompany,
           Manufacturer, StockHeld, WebsiteOH, ProductGroup,
           BOMItemTypeID, LeadTime, LeadTimeUnitID, MOQ, '' AS DocStatusID, '' AS OrderDate, '' AS EnteredBy
    FROM   stk
)

SELECT  RowType, Company, RowKey, ProductCode, Description, DocNo, LineSeq,
        AccountNo, AccountName, Reference, DateISO, Qty1, Qty2, Qty3, Qty4,
        FulfilMethodID, LinkID, B2BCount, Intercompany,
        Manufacturer, StockHeld, WebsiteOH, ProductGroup,
        BOMItemTypeID, LeadTime, LeadTimeUnitID, MOQ, DocStatusID, OrderDate, EnteredBy
FROM    allrows
ORDER BY CASE RowType WHEN 'SO' THEN 1 WHEN 'PO' THEN 2 ELSE 3 END,
         Company, ProductCode, DocNo, LineSeq;
