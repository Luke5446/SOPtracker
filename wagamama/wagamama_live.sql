/* ============================================================================
   WAGAMAMA T-SHIRT CODE SWITCH - LIVE ORDERS, POs AND STOCK (v1, 19 Nov 2026)
   Sage 200 Professional, TIB-SQL-002, S200_LIVE (Tibard) only. Wagamama is a
   Tibard customer (WAG001A etc.) and every WAGTS code is a Tibard stock item,
   so Oliver Harvey is not read.

   Same 30-column A-AD layout as ordertracker_live.sql (SOPtracker), limited
   to the Wagamama T-shirt codes, so the paste works in the Wagamama dashboard
   AND would be accepted by the Order tracker. Three row types:

     SO   every live/held sales order line whose item code starts WAGTS,
          despatched lines too (Qty3 outstanding = 0)
     PO   live purchase order lines on WAGTS codes with quantity outstanding
     STK  EVERY WAGTS stock item in Sage (not just the ones on orders), so the
          dashboard can see an alternative colour/style/size with stock even
          when nobody is ordering it.

   Columns (fixed, the app parses by position):
     A RowType  B Company  C RowKey  D ProductCode  E Description  F DocNo
     G LineSeq  H AccountNo  I AccountName  J Reference (SO: customer's order
     ref; PO: supplier ref; STK: supplier stock code)  K DateISO (SO promised,
     header first; PO requested)  L Qty1 (ordered / in stock)  M Qty2
     (despatched / received / allocated)  N Qty3 (OUTSTANDING / FREE STOCK)
     O Qty4 (SO allocated / STK on PO)  P FulfilMethodID  Q LinkID (B2B PO/SO)
     R B2BCount  S Intercompany  T Manufacturer  U StockHeld  V WebsiteOH
     W ProductGroup  X BOMItemTypeID  Y LeadTime  Z LeadTimeUnitID  AA MOQ
     AB DocStatusID (0 live, 1 hold)  AC OrderDate  AD EnteredBy

   SQL 1.xlsx returns the first result set only: one SELECT at the end.
   Copy A to AD without the header into the dashboard's Paste box (?edit).
   READ UNCOMMITTED: report only, takes no locks. Table variables, not #temp
   (Power Query cannot see #temp tables). COLLATE DATABASE_DEFAULT because
   tempdb collates differently from S200_LIVE.
   ========================================================================= */

SET NOCOUNT ON;
SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;

DECLARE @b2bso TABLE (SOPOrderReturnID bigint PRIMARY KEY, FirstDoc nvarchar(40) COLLATE DATABASE_DEFAULT, N int);
DECLARE @b2bpo TABLE (POPOrderReturnID bigint PRIMARY KEY, FirstDoc nvarchar(40) COLLATE DATABASE_DEFAULT, N int);

DECLARE @so TABLE (
    RowKey nvarchar(60) COLLATE DATABASE_DEFAULT, ProductCode nvarchar(60) COLLATE DATABASE_DEFAULT,
    Description nvarchar(255) COLLATE DATABASE_DEFAULT, DocNo nvarchar(40) COLLATE DATABASE_DEFAULT, LineSeq int,
    AccountNo nvarchar(40) COLLATE DATABASE_DEFAULT, AccountName nvarchar(255) COLLATE DATABASE_DEFAULT,
    Reference nvarchar(255) COLLATE DATABASE_DEFAULT, DateISO nvarchar(10) COLLATE DATABASE_DEFAULT,
    Qty1 decimal(18,2), Qty2 decimal(18,2), Qty3 decimal(18,2), Qty4 decimal(18,2),
    FulfilMethodID nvarchar(10) COLLATE DATABASE_DEFAULT, LinkID nvarchar(40) COLLATE DATABASE_DEFAULT,
    B2BCount nvarchar(10) COLLATE DATABASE_DEFAULT, Intercompany nvarchar(1) COLLATE DATABASE_DEFAULT,
    DocStatusID nvarchar(10) COLLATE DATABASE_DEFAULT, OrderDate nvarchar(10) COLLATE DATABASE_DEFAULT,
    EnteredBy nvarchar(255) COLLATE DATABASE_DEFAULT);
DECLARE @po TABLE (
    RowKey nvarchar(60) COLLATE DATABASE_DEFAULT, ProductCode nvarchar(60) COLLATE DATABASE_DEFAULT,
    Description nvarchar(255) COLLATE DATABASE_DEFAULT, DocNo nvarchar(40) COLLATE DATABASE_DEFAULT, LineSeq int,
    AccountNo nvarchar(40) COLLATE DATABASE_DEFAULT, AccountName nvarchar(255) COLLATE DATABASE_DEFAULT,
    Reference nvarchar(255) COLLATE DATABASE_DEFAULT, DateISO nvarchar(10) COLLATE DATABASE_DEFAULT,
    Qty1 decimal(18,2), Qty2 decimal(18,2), Qty3 decimal(18,2),
    LinkID nvarchar(40) COLLATE DATABASE_DEFAULT, B2BCount nvarchar(10) COLLATE DATABASE_DEFAULT,
    DocStatusID nvarchar(10) COLLATE DATABASE_DEFAULT, OrderDate nvarchar(10) COLLATE DATABASE_DEFAULT,
    EnteredBy nvarchar(255) COLLATE DATABASE_DEFAULT);
DECLARE @stk TABLE (
    RowKey nvarchar(70) COLLATE DATABASE_DEFAULT, ProductCode nvarchar(60) COLLATE DATABASE_DEFAULT,
    Description nvarchar(255) COLLATE DATABASE_DEFAULT,
    AccountNo nvarchar(40) COLLATE DATABASE_DEFAULT, AccountName nvarchar(255) COLLATE DATABASE_DEFAULT,
    Reference nvarchar(255) COLLATE DATABASE_DEFAULT,
    Qty1 decimal(18,2), Qty2 decimal(18,2), Qty3 decimal(18,2), Qty4 decimal(18,2),
    FulfilMethodID nvarchar(10) COLLATE DATABASE_DEFAULT, Manufacturer nvarchar(255) COLLATE DATABASE_DEFAULT,
    StockHeld nvarchar(3) COLLATE DATABASE_DEFAULT, ProductGroup nvarchar(40) COLLATE DATABASE_DEFAULT,
    BOMItemTypeID nvarchar(10) COLLATE DATABASE_DEFAULT, LeadTime nvarchar(10) COLLATE DATABASE_DEFAULT,
    LeadTimeUnitID nvarchar(10) COLLATE DATABASE_DEFAULT, MOQ decimal(18,2), B2B12m nvarchar(10) COLLATE DATABASE_DEFAULT);

/* 1. Back-to-back links (first PO per SO, first SO per PO, and counts). */
INSERT INTO @b2bso (SOPOrderReturnID, FirstDoc, N)
SELECT b.SOPOrderReturnID, MIN(lpo.DocumentNo), COUNT(*)
FROM   S200_LIVE.dbo.MMS123BackToBackPO b
JOIN   S200_LIVE.dbo.POPOrderReturn lpo ON lpo.POPOrderReturnID = b.POPOrderReturnID
GROUP BY b.SOPOrderReturnID;

INSERT INTO @b2bpo (POPOrderReturnID, FirstDoc, N)
SELECT b.POPOrderReturnID, MIN(lso.DocumentNo), COUNT(*)
FROM   S200_LIVE.dbo.MMS123BackToBackPO b
JOIN   S200_LIVE.dbo.SOPOrderReturn lso ON lso.SOPOrderReturnID = b.SOPOrderReturnID
GROUP BY b.POPOrderReturnID;

/* 2. SO rows: every WAGTS line on a live or held sales order. */
INSERT INTO @so (RowKey, ProductCode, Description, DocNo, LineSeq, AccountNo, AccountName, Reference, DateISO,
                 Qty1, Qty2, Qty3, Qty4, FulfilMethodID, LinkID, B2BCount, Intercompany, DocStatusID, OrderDate, EnteredBy)
SELECT
    'TIB-SO-' + CAST(sorl.SOPOrderReturnLineID AS varchar(20)),
    LTRIM(RTRIM(sorl.ItemCode)),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(sorl.ItemDescription,''), CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))),
    sor.DocumentNo,
    sorl.PrintSequenceNumber,
    ISNULL(cust.CustomerAccountNumber,''),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(cust.CustomerAccountName,''), CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(sor.CustomerDocumentNo,''), CHAR(9),' '), CHAR(13),' '), CHAR(10),' '))),
    CONVERT(varchar(10), COALESCE(sor.PromisedDeliveryDate, sorl.PromisedDeliveryDate,
                                  sor.RequestedDeliveryDate, sorl.RequestedDeliveryDate), 23),
    CAST(sorl.StockUnitLineQuantity AS decimal(18,2)),
    CAST(ISNULL(sorl.StockUnitDespRcptQuantity,0) AS decimal(18,2)),
    CAST(sorl.StockUnitLineQuantity - ISNULL(sorl.StockUnitDespRcptQuantity,0) AS decimal(18,2)),
    CAST(ISNULL(sorl.StockUnitAllocatedQuantity,0) AS decimal(18,2)),
    ISNULL(CAST(sorl.SOPOrderFulfilmentMethodID AS varchar(10)),''),
    ISNULL(bs.FirstDoc,''),
    CAST(ISNULL(bs.N,0) AS varchar(10)),
    CASE WHEN cust.CustomerAccountNumber = 'OLIVER' THEN 'Y' ELSE '' END,
    CAST(sor.DocumentStatusID AS varchar(10)),
    ISNULL(CONVERT(varchar(10), sor.DocumentDate, 23),''),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(sor.DocumentCreatedBy,''), CHAR(9),' '), CHAR(13),' '), CHAR(10),' ')))
FROM        S200_LIVE.dbo.SOPOrderReturn      sor
INNER JOIN  S200_LIVE.dbo.SOPOrderReturnLine  sorl ON sorl.SOPOrderReturnID    = sor.SOPOrderReturnID
LEFT  JOIN  S200_LIVE.dbo.SLCustomerAccount   cust ON cust.SLCustomerAccountID = sor.CustomerID
LEFT  JOIN  @b2bso                            bs   ON bs.SOPOrderReturnID      = sor.SOPOrderReturnID
WHERE   sor.DocumentTypeID   = 0
    AND sor.DocumentStatusID IN (0,1)
    AND sorl.LineTypeID      = 0
    AND sorl.ItemCode LIKE 'WAGTS%'
OPTION (RECOMPILE);

/* 3. PO rows: live WAGTS purchase order lines with quantity outstanding. */
INSERT INTO @po (RowKey, ProductCode, Description, DocNo, LineSeq, AccountNo, AccountName, Reference, DateISO,
                 Qty1, Qty2, Qty3, LinkID, B2BCount, DocStatusID, OrderDate, EnteredBy)
SELECT
    'TIB-PO-' + CAST(porl.POPOrderReturnLineID AS varchar(20)),
    LTRIM(RTRIM(porl.ItemCode)),
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
    ISNULL(bp.FirstDoc,''),
    CAST(ISNULL(bp.N,0) AS varchar(10)),
    CAST(por.DocumentStatusID AS varchar(10)),
    ISNULL(CONVERT(varchar(10), por.DocumentDate, 23),''),
    LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(por.DocumentCreatedBy,''), CHAR(9),' '), CHAR(13),' '), CHAR(10),' ')))
FROM        S200_LIVE.dbo.POPOrderReturn      por
INNER JOIN  S200_LIVE.dbo.POPOrderReturnLine  porl ON porl.POPOrderReturnID    = por.POPOrderReturnID
LEFT  JOIN  S200_LIVE.dbo.PLSupplierAccount   supp ON supp.PLSupplierAccountID = por.SupplierID
LEFT  JOIN  @b2bpo                            bp   ON bp.POPOrderReturnID      = por.POPOrderReturnID
WHERE   por.DocumentTypeID   = 0
    AND por.DocumentStatusID IN (0,1)
    AND porl.LineTypeID      = 0
    AND porl.ItemCode LIKE 'WAGTS%'
    AND (porl.StockUnitLineQuantity - ISNULL(porl.StockUnitRcptRtnQuantity,0)) > 0
OPTION (RECOMPILE);

/* 4. STK rows: EVERY WAGTS stock item, plus any code on an order that has no
   stock record (shown as CODE NOT IN STOCK FILE). */
INSERT INTO @stk (RowKey, ProductCode, Description, AccountNo, AccountName, Reference, Qty1, Qty2, Qty3, Qty4,
                  FulfilMethodID, Manufacturer, StockHeld, ProductGroup, BOMItemTypeID, LeadTime, LeadTimeUnitID, MOQ, B2B12m)
SELECT
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
    CASE WHEN ISNULL(si.AnalysisCode3,'') = 'Yes' THEN 'Yes' ELSE '' END,
    ISNULL(pg.Code,''),
    ISNULL(CAST(si.BOMItemTypeID AS varchar(10)),''),
    ISNULL(CAST(ps.LeadTime AS varchar(10)),''),
    ISNULL(CAST(ps.LeadTimeUnitID AS varchar(10)),''),
    CAST(ps.MinimumOrderQuantity AS decimal(18,2)),
    ISNULL(CAST(bb.N AS varchar(10)),'0')
FROM (
    SELECT LTRIM(RTRIM(Code)) COLLATE DATABASE_DEFAULT AS ProductCode FROM S200_LIVE.dbo.StockItem WHERE Code LIKE 'WAGTS%'
    UNION SELECT ProductCode FROM @so
    UNION SELECT ProductCode FROM @po
) c
LEFT  JOIN  S200_LIVE.dbo.StockItem     si ON si.Code           = c.ProductCode
LEFT  JOIN  S200_LIVE.dbo.ProductGroup  pg ON pg.ProductGroupID = si.ProductGroupID
OUTER APPLY (
    SELECT  SUM(wi.ConfirmedQtyInStock) AS InStock,
            SUM(ISNULL(wi.QuantityAllocatedSOP,0) + ISNULL(wi.QuantityAllocatedStock,0) + ISNULL(wi.QuantityAllocatedBOM,0)) AS Allocated,
            SUM(wi.QuantityOnPOPOrder) AS OnPO
    FROM    S200_LIVE.dbo.WarehouseItem wi
    WHERE   wi.ItemID = si.ItemID
) wh
OUTER APPLY (
    SELECT TOP 1 supp.SupplierAccountNumber, supp.SupplierAccountName,
           sis.SupplierStockCode, sis.LeadTime, sis.LeadTimeUnitID, sis.MinimumOrderQuantity
    FROM   S200_LIVE.dbo.StockItemSupplier sis
    LEFT JOIN S200_LIVE.dbo.PLSupplierAccount supp ON supp.PLSupplierAccountID = sis.SupplierID
    WHERE  sis.ItemID = si.ItemID
    ORDER BY sis.Preferred DESC, sis.DateLastOrder DESC
) ps
OUTER APPLY (
    SELECT COUNT(DISTINCT b.POPOrderReturnID) AS N
    FROM   S200_LIVE.dbo.MMS123BackToBackPO  b
    JOIN   S200_LIVE.dbo.POPOrderReturn      lpo ON lpo.POPOrderReturnID = b.POPOrderReturnID
    JOIN   S200_LIVE.dbo.POPOrderReturnLine  lpl ON lpl.POPOrderReturnID = lpo.POPOrderReturnID
    WHERE  lpo.DocumentDate >= DATEADD(month, -12, GETDATE())
      AND  LTRIM(RTRIM(lpl.ItemCode)) = c.ProductCode
) bb
OPTION (RECOMPILE);

/* 5. THE RESULT: one shape, 30 columns A-AD. */
SELECT  RowType, Company, RowKey, ProductCode, Description, DocNo, LineSeq,
        AccountNo, AccountName, Reference, DateISO, Qty1, Qty2, Qty3, Qty4,
        FulfilMethodID, LinkID, B2BCount, Intercompany,
        Manufacturer, StockHeld, WebsiteOH, ProductGroup,
        BOMItemTypeID, LeadTime, LeadTimeUnitID, MOQ, DocStatusID, OrderDate, EnteredBy
FROM (
    SELECT 'SO' AS RowType, 'TIBARD' AS Company, RowKey, ProductCode, Description, DocNo, LineSeq,
           AccountNo, AccountName, Reference, DateISO, Qty1, Qty2, Qty3, Qty4,
           FulfilMethodID, LinkID, B2BCount, Intercompany,
           '' AS Manufacturer, '' AS StockHeld, '' AS WebsiteOH, '' AS ProductGroup,
           '' AS BOMItemTypeID, '' AS LeadTime, '' AS LeadTimeUnitID,
           CAST(NULL AS decimal(18,2)) AS MOQ, DocStatusID, OrderDate, EnteredBy
    FROM   @so
    UNION ALL
    SELECT 'PO', 'TIBARD', RowKey, ProductCode, Description, DocNo, LineSeq,
           AccountNo, AccountName, Reference, DateISO, Qty1, Qty2, Qty3, CAST(NULL AS decimal(18,2)),
           '', LinkID, B2BCount, '',
           '', '', '', '', '', '', '', CAST(NULL AS decimal(18,2)), DocStatusID, OrderDate, EnteredBy
    FROM   @po
    UNION ALL
    SELECT 'STK', 'TIBARD', RowKey, ProductCode, Description, '', CAST(NULL AS int),
           AccountNo, AccountName, Reference, '', Qty1, Qty2, Qty3, Qty4,
           FulfilMethodID, '', B2B12m, '',
           Manufacturer, StockHeld, '', ProductGroup,
           BOMItemTypeID, LeadTime, LeadTimeUnitID, MOQ, '', '', ''
    FROM   @stk
) allrows
ORDER BY CASE RowType WHEN 'SO' THEN 1 WHEN 'PO' THEN 2 ELSE 3 END, ProductCode, DocNo, LineSeq
OPTION (RECOMPILE);
