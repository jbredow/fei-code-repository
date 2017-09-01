CREATE OR REPLACE FORCE VIEW "AAE0376"."GP_TRACKER_WRITER_SUMS"
AS

SELECT GP_DATA.YEARMONTH,
       CASE
          WHEN GP_DATA.TYPE_OF_SALE IN ('Showroom', 'Showroom Direct')
          THEN
             'Showroom'
          ELSE
             GP_DATA.TYPE_OF_SALE
       END
          AS CHANNEL,
       DENSE_RANK () OVER (ORDER BY GP_DATA.YEARMONTH ASC) AS MM,
       CASE
          WHEN ACCT.REGION_NAME IS NULL THEN ACCT.ACCOUNT_NAME
          WHEN ACCT.REGION_NAME = 'EAST' THEN 'EASTERN'
          WHEN ACCT.REGION_NAME = 'WEST' THEN 'WESTERN'
          ELSE ACCT.REGION_NAME
       END
          REGION,
       GP_DATA.ACCOUNT_NUMBER,
       ACCT.ACCOUNT_NAME,
       NVL (GP_DATA.WRITER, '#N/A') WRITER_INIT,
          CASE
             WHEN ACCT.REGION_NAME IS NULL THEN ACCT.ACCOUNT_NAME
             WHEN ACCT.REGION_NAME = 'EAST' THEN 'EASTERN'
             WHEN ACCT.REGION_NAME = 'WEST' THEN 'WESTERN'
             ELSE ACCT.REGION_NAME
          END
       || '*'
       || GP_DATA.ACCOUNT_NUMBER
       || '*'
       || NVL (GP_DATA.WRITER, '#N/A')
       || CASE
             WHEN GP_DATA.TYPE_OF_SALE IN ('Showroom', 'Showroom Direct')
             THEN
                'Showroom'
             ELSE
                GP_DATA.TYPE_OF_SALE
          END
          AS GPTRACK_KEY,
       SUM (GP_DATA.EXT_SALES_AMOUNT) "Total Sales",
       SUM (EXT_AVG_COGS_AMOUNT) "Total Cost",
       SUM (GP_DATA.EXT_SALES_AMOUNT - EXT_AVG_COGS_AMOUNT) "Total GP$",
       SUM (GP_DATA.INVOICE_LINES) "Total # Lines",
       COUNT (GP_DATA.INVOICE_NUMBER_NK) "Total # Invoices",
       COUNT (
          DISTINCT (CASE
                       WHEN GP_DATA.INVOICE_NUMBER_NK NOT IN ('CM%', '%-%')
                       THEN
                          1
                       ELSE
                          0
                    END))
          "Total Invoice Count",
       SUM (
          CASE
             WHEN GP_DATA.PRICE_CATEGORY IN ('MATRIX', 'MATRIX_BID')
             THEN
                (GP_DATA.EXT_SALES_AMOUNT)
             ELSE
                0
          END)
          "Price Matrix Sales",
       SUM (
          CASE
             WHEN GP_DATA.PRICE_CATEGORY IN ('MATRIX', 'MATRIX_BID')
             THEN
                (GP_DATA.EXT_AVG_COGS_AMOUNT)
             ELSE
                0
          END)
          "Price Matrix Cost",
       SUM (
          CASE
             WHEN GP_DATA.PRICE_CATEGORY IN ('MATRIX', 'MATRIX_BID')
             THEN
                (GP_DATA.EXT_SALES_AMOUNT - GP_DATA.EXT_AVG_COGS_AMOUNT)
             ELSE
                0
          END)
          "Price Matrix GP$",
       SUM (
          CASE
             WHEN GP_DATA.PRICE_CATEGORY IN ('MATRIX', 'MATRIX_BID')
             THEN
                (GP_DATA.INVOICE_LINES)
             ELSE
                0
          END)
          "Price Matrix # Lines",
       COUNT (
          CASE
             WHEN( GP_DATA.PRICE_CATEGORY IN ('MATRIX', 'MATRIX_BID')AND (GP_DATA.INVOICE_NUMBER_NK) NOT IN ('CM%', '%-%'))
             THEN
                (GP_DATA.INVOICE_NUMBER_NK)
             
             ELSE
                NULL
          END)
          "Price Matrix Invoice Cnt",
       SUM (
          CASE
             WHEN GP_DATA.PRICE_CATEGORY IN 'OVERRIDE'
             THEN
                (GP_DATA.EXT_SALES_AMOUNT)
             ELSE
                0
          END)
          "Contract Sales",
       SUM (
          CASE
             WHEN GP_DATA.PRICE_CATEGORY IN 'OVERRIDE'
             THEN
                (GP_DATA.EXT_AVG_COGS_AMOUNT)
             ELSE
                0
          END)
          "Contract Cost",
       SUM (
          CASE
             WHEN GP_DATA.PRICE_CATEGORY IN 'OVERRIDE'
             THEN
                (GP_DATA.EXT_SALES_AMOUNT - GP_DATA.EXT_AVG_COGS_AMOUNT)
             ELSE
                0
          END)
          "Contract GP$",
       SUM (
          CASE
             WHEN GP_DATA.PRICE_CATEGORY IN 'OVERRIDE'
             THEN
                (GP_DATA.INVOICE_LINES)
             ELSE
                0
          END)
          "Contract # Lines",
       COUNT (
          CASE
             WHEN (GP_DATA.PRICE_CATEGORY IN ('OVERRIDE') AND (GP_DATA.INVOICE_NUMBER_NK) NOT IN ('CM%', '%-%'))
             THEN
                (GP_DATA.INVOICE_NUMBER_NK)
             ELSE
                NULL
          END)
          "Contract Invoice Cnt",
       SUM (
          CASE
             WHEN GP_DATA.PRICE_CATEGORY IN ('MANUAL', 'TOOLS', 'QUOTE')
             THEN
                (GP_DATA.EXT_SALES_AMOUNT)
             ELSE
                0
          END)
          "Manual Sales",
       SUM (
          CASE
             WHEN GP_DATA.PRICE_CATEGORY IN ('MANUAL', 'TOOLS', 'QUOTE')
             THEN
                (GP_DATA.EXT_AVG_COGS_AMOUNT)
             ELSE
                0
          END)
          "Manual Cost",
       SUM (
          CASE
             WHEN GP_DATA.PRICE_CATEGORY IN ('MANUAL', 'TOOLS', 'QUOTE')
             THEN
                (GP_DATA.EXT_SALES_AMOUNT - GP_DATA.EXT_AVG_COGS_AMOUNT)
             ELSE
                0
          END)
          "Manual GP$",
       SUM (
          CASE
             WHEN GP_DATA.PRICE_CATEGORY IN ('MANUAL', 'TOOLS', 'QUOTE')
             THEN
                (GP_DATA.INVOICE_LINES)
             ELSE
                0
          END)
          "Manual # Lines",
        COUNT (
          CASE
             WHEN (GP_DATA.PRICE_CATEGORY IN ('MANUAL', 'TOOLS', 'QUOTE')AND (GP_DATA.INVOICE_NUMBER_NK) NOT IN ('CM%', '%-%'))
             THEN
                (GP_DATA.INVOICE_NUMBER_NK)
             ELSE
                NULL
          END)
          "Manual Invoice Cnt",
       SUM (CASE
               WHEN GP_DATA.PRICE_CATEGORY NOT IN ('MATRIX',
                                                   'OVERRIDE',
                                                   'MANUAL',
                                                   'CREDITS',
                                                   'TOOLS',
                                                   'QUOTE',
                                                   'MATRIX_BID',
                                                   'Total')
               THEN
                  (GP_DATA.EXT_SALES_AMOUNT)
               ELSE
                  0
            END)
          "Other Sales",
       SUM (CASE
               WHEN GP_DATA.PRICE_CATEGORY NOT IN ('MATRIX',
                                                   'OVERRIDE',
                                                   'MANUAL',
                                                   'CREDITS',
                                                   'TOOLS',
                                                   'QUOTE',
                                                   'MATRIX_BID',
                                                   'Total')
               THEN
                  (GP_DATA.EXT_AVG_COGS_AMOUNT)
               ELSE
                  0
            END)
          "Other Cost",
       SUM (CASE
               WHEN GP_DATA.PRICE_CATEGORY NOT IN ('MATRIX',
                                                   'OVERRIDE',
                                                   'MANUAL',
                                                   'CREDITS',
                                                   'TOOLS',
                                                   'QUOTE',
                                                   'MATRIX_BID',
                                                   'Total')
               THEN
                  (GP_DATA.EXT_SALES_AMOUNT - GP_DATA.EXT_AVG_COGS_AMOUNT)
               ELSE
                  0
            END)
          "Other GP$",
       SUM (CASE
               WHEN GP_DATA.PRICE_CATEGORY NOT IN ('MATRIX',
                                                   'OVERRIDE',
                                                   'MANUAL',
                                                   'CREDITS',
                                                   'TOOLS',
                                                   'QUOTE',
                                                   'MATRIX_BID',
                                                   'Total')
               THEN
                  (GP_DATA.INVOICE_LINES)
               ELSE
                  0
            END)
          "Other # Lines",
       COUNT (
          CASE
             WHEN (GP_DATA.PRICE_CATEGORY NOT IN ('MATRIX',
                                                   'OVERRIDE',
                                                   'MANUAL',
                                                   'CREDITS',
                                                   'TOOLS',
                                                   'QUOTE',
                                                   'MATRIX_BID',
                                                   'Total')AND (GP_DATA.INVOICE_NUMBER_NK) NOT IN ('CM%', '%-%'))
             THEN
                (GP_DATA.INVOICE_NUMBER_NK)
             ELSE
                NULL
          END)
          "Other Invoice Cnt",
       ROUND (
          SUM (
             CASE
                WHEN GP_DATA.PRICE_CATEGORY IN ('CREDITS')
                THEN
                   (GP_DATA.EXT_SALES_AMOUNT)
                ELSE
                   0
             END),
          3)
          "Credits $",
       SUM (
          CASE
             WHEN GP_DATA.PRICE_CATEGORY IN ('CREDITS')
             THEN
                (GP_DATA.INVOICE_LINES)
             ELSE
                0
          END)
          "Credits Lines",
       
       SUM (
          CASE
             WHEN GP_DATA.PRICE_CATEGORY NOT IN ('CREDITS', 'Total')
             THEN
                (GP_DATA.EXT_SALES_AMOUNT)
             ELSE
                0
          END)
          "Outbound Sales"
  FROM AAE0376.PR_VICT2_CUST_12MO GP_DATA,
       (SELECT WD.REGION_NAME,
               WD.ACCOUNT_NAME,
               WD.ACCOUNT_NUMBER_NK,
               WD.WAREHOUSE_NUMBER_NK
          FROM SALES_MART.SALES_WAREHOUSE_DIM WD
        GROUP BY WD.REGION_NAME,
                 WD.ACCOUNT_NAME,
                 WD.ACCOUNT_NUMBER_NK,
                 WD.WAREHOUSE_NUMBER_NK) ACCT                               --
 WHERE     GP_DATA.WAREHOUSE_NUMBER = ACCT.WAREHOUSE_NUMBER_NK(+)
       AND GP_DATA.WRITER IS NOT NULL
GROUP BY GP_DATA.YEARMONTH,
         CASE
            WHEN ACCT.REGION_NAME IS NULL THEN ACCT.ACCOUNT_NAME
            WHEN ACCT.REGION_NAME = '%EAST%' THEN 'EASTERN'
            WHEN ACCT.REGION_NAME = '%WEST%' THEN 'WESTERN'
            ELSE ACCT.REGION_NAME
         END,
         GP_DATA.ACCOUNT_NUMBER,
         ACCT.ACCOUNT_NAME,
         ACCT.REGION_NAME,
         GP_DATA.INVOICE_NUMBER_NK,
         GP_DATA.WRITER,
         CASE
            WHEN GP_DATA.TYPE_OF_SALE IN ('Showroom', 'Showroom Direct')
            THEN
               'Showroom'
            ELSE
               GP_DATA.TYPE_OF_SALE
         END
ORDER BY    CASE
               WHEN ACCT.REGION_NAME IS NULL THEN ACCT.ACCOUNT_NAME
               WHEN ACCT.REGION_NAME = '%EAST%' THEN 'EASTERN'
               WHEN ACCT.REGION_NAME = '%WEST%' THEN 'WESTERN'
               ELSE ACCT.REGION_NAME
            END
         || '*'
         || GP_DATA.ACCOUNT_NUMBER
         || '*'
         || NVL (GP_DATA.WRITER, '#N/A')
         || CASE
               WHEN GP_DATA.TYPE_OF_SALE IN ('Showroom', 'Showroom Direct')
               THEN
                  'Showroom'
               ELSE
                  GP_DATA.TYPE_OF_SALE
            END
  ;
