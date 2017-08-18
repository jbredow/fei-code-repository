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
          WHEN GP_DATA.REGION IS NULL THEN ACCT.ACCOUNT_NAME
          WHEN GP_DATA.REGION = 'EAST' THEN 'EASTERN'
          WHEN GP_DATA.REGION = 'WEST' THEN 'WESTERN'
          ELSE GP_DATA.REGION
       END
          REGION,
       GP_DATA.ACCOUNT_NUMBER,
       ACCT.ACCOUNT_NAME,
       NVL (GP_DATA.WRITER, '#N/A') WRITER_INIT,
       --NVL(REPS.ASSOC_NAME,'#N/A') WRITER_NAME,
       CASE
          WHEN GP_DATA.REGION IS NULL THEN ACCT.ACCOUNT_NAME
          WHEN GP_DATA.REGION = 'EAST' THEN 'EASTERN'
          WHEN GP_DATA.REGION = 'WEST' THEN 'WESTERN'
          ELSE GP_DATA.REGION
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
       SUM (
            GP_DATA.SLS_SUBTOTAL
          + GP_DATA.SLS_FREIGHT
          + GP_DATA.SLS_MISC
          + GP_DATA.SLS_RESTOCK)
          "Total Sales",
       SUM (GP_DATA.SLS_SUBTOTAL) "Sales, Sub-Total",
       SUM (GP_DATA.SLS_FREIGHT) "Sales, Freight",
       SUM (GP_DATA.SLS_MISC) "Sales, Misc",
       SUM (GP_DATA.SLS_RESTOCK) "Sales, Restock",
       SUM (
            GP_DATA.AVG_COST_SUBTOTAL
          + GP_DATA.AVG_COST_FREIGHT
          + GP_DATA.AVG_COST_MISC)
          "Total Cost",
       SUM (GP_DATA.AVG_COST_SUBTOTAL) "Cost, Sub-Total",
       SUM (GP_DATA.AVG_COST_FREIGHT) "Cost, Freight",
       SUM (GP_DATA.AVG_COST_MISC) "Cost, Misc",
       SUM (
          (  GP_DATA.SLS_SUBTOTAL
           + GP_DATA.SLS_FREIGHT
           + GP_DATA.SLS_MISC
           + GP_DATA.SLS_RESTOCK)
          - (  GP_DATA.AVG_COST_SUBTOTAL
             + GP_DATA.AVG_COST_FREIGHT
             + GP_DATA.AVG_COST_MISC))
          "Total GP$",
       ROUND (
          SUM (
             (  GP_DATA.SLS_SUBTOTAL
              + GP_DATA.SLS_FREIGHT
              + GP_DATA.SLS_MISC
              + GP_DATA.SLS_RESTOCK)
             - (  GP_DATA.AVG_COST_SUBTOTAL
                + GP_DATA.AVG_COST_FREIGHT
                + GP_DATA.AVG_COST_MISC))
          / SUM (
                 GP_DATA.SLS_SUBTOTAL
               + GP_DATA.SLS_FREIGHT
               + GP_DATA.SLS_MISC
               + GP_DATA.SLS_RESTOCK),
          3)
          "Total GP%",
       SUM (GP_DATA.INVOICE_LINES) "Total # Lines",
       SUM (
          CASE
             WHEN GP_DATA.PRICE_CATEGORY IN 'MATRIX' THEN (GP_DATA.EXT_SALES)
             ELSE 0
          END)
          "Price Matrix Sales",
       SUM (
          CASE
             WHEN GP_DATA.PRICE_CATEGORY IN 'MATRIX' THEN (GP_DATA.AVG_COGS)
             ELSE 0
          END)
          "Price Matrix Cost",
       SUM (
          CASE
             WHEN GP_DATA.PRICE_CATEGORY IN 'MATRIX'
             THEN
                (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
             ELSE
                0
          END)
          "Price Matrix GP$",
       ROUND (
          SUM (
             CASE
                WHEN GP_DATA.PRICE_CATEGORY IN 'MATRIX'
                THEN
                   (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                ELSE
                   0
             END)
          / SUM (
               CASE
                  WHEN GP_DATA.PRICE_CATEGORY IN 'MATRIX'
                  THEN
                     CASE
                        WHEN GP_DATA.EXT_SALES > 0 THEN (GP_DATA.EXT_SALES)
                        ELSE 1
                     END
                  ELSE
                     1
               END),
          3)
          "Price Matrix GP%",
       ROUND (
          SUM (
             CASE
                WHEN GP_DATA.PRICE_CATEGORY IN 'MATRIX'
                THEN
                   (GP_DATA.EXT_SALES)
                ELSE
                   0
             END)
          / SUM (
               CASE
                  WHEN GP_DATA.SLS_SUBTOTAL > 0 THEN GP_DATA.SLS_SUBTOTAL
                  ELSE 1
               END),
          3)
          "Price Matrix Use%$",
       ROUND (
          SUM (
             CASE
                WHEN GP_DATA.PRICE_CATEGORY IN 'MATRIX'
                THEN
                   (GP_DATA.INVOICE_LINES)
                ELSE
                   0
             END)
          / SUM (
               CASE
                  WHEN GP_DATA.INVOICE_LINES > 0 THEN GP_DATA.INVOICE_LINES
                  ELSE 1
               END),
          3)
          "Price Matrix Use%#",
       ROUND (
          SUM (
             CASE
                WHEN GP_DATA.PRICE_CATEGORY IN 'MATRIX'
                THEN
                   (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                ELSE
                   0
             END)
          / SUM (
               CASE
                  WHEN GP_DATA.ROLLUP = 'Total'
                  THEN
                     (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                  ELSE
                     0
               END),
          3)
          "Price Matrix Profit%$",
       SUM (
          CASE
             WHEN GP_DATA.PRICE_CATEGORY IN 'MATRIX'
             THEN
                (GP_DATA.INVOICE_LINES)
             ELSE
                0
          END)
          "Price Matrix # Lines",
       SUM (
          CASE
             WHEN GP_DATA.PRICE_CATEGORY IN 'OVERRIDE'
             THEN
                (GP_DATA.EXT_SALES)
             ELSE
                0
          END)
          "Contract Sales",
       SUM (
          CASE
             WHEN GP_DATA.PRICE_CATEGORY IN 'OVERRIDE'
             THEN
                (GP_DATA.AVG_COGS)
             ELSE
                0
          END)
          "Contract Cost",
       SUM (
          CASE
             WHEN GP_DATA.PRICE_CATEGORY IN 'OVERRIDE'
             THEN
                (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
             ELSE
                0
          END)
          "Contract GP$",
       ROUND (
          SUM (
             CASE
                WHEN GP_DATA.PRICE_CATEGORY IN 'OVERRIDE'
                THEN
                   (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                ELSE
                   0
             END)
          / SUM (
               CASE
                  WHEN GP_DATA.PRICE_CATEGORY IN 'OVERRIDE'
                  THEN
                     CASE
                        WHEN GP_DATA.EXT_SALES > 0 THEN (GP_DATA.EXT_SALES)
                        ELSE 1
                     END
                  ELSE
                     1
               END),
          3)
          "Contract GP%",
       ROUND (
          SUM (
             CASE
                WHEN GP_DATA.PRICE_CATEGORY IN 'OVERRIDE'
                THEN
                   (GP_DATA.EXT_SALES)
                ELSE
                   0
             END)
          / SUM (
               CASE
                  WHEN GP_DATA.SLS_SUBTOTAL > 0 THEN GP_DATA.SLS_SUBTOTAL
                  ELSE 1
               END),
          3)
          "Contract Use%$",
       ROUND (
          SUM (
             CASE
                WHEN GP_DATA.PRICE_CATEGORY IN 'OVERRIDE'
                THEN
                   (GP_DATA.INVOICE_LINES)
                ELSE
                   0
             END)
          / SUM (
               CASE
                  WHEN GP_DATA.INVOICE_LINES > 0 THEN GP_DATA.INVOICE_LINES
                  ELSE 1
               END),
          3)
          "Contract Use%#",
       ROUND (
          SUM (
             CASE
                WHEN GP_DATA.PRICE_CATEGORY IN 'OVERRIDE'
                THEN
                   (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                ELSE
                   0
             END)
          / SUM (
               CASE
                  WHEN GP_DATA.ROLLUP = 'Total'
                  THEN
                     (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                  ELSE
                     0
               END),
          3)
          "Contract Profit%$",
       SUM (
          CASE
             WHEN GP_DATA.PRICE_CATEGORY IN 'OVERRIDE'
             THEN
                (GP_DATA.INVOICE_LINES)
             ELSE
                0
          END)
          "Contract # Lines",
       SUM (
          CASE
             WHEN GP_DATA.PRICE_CATEGORY IN ('MANUAL', 'TOOLS', 'QUOTE')
             THEN
                (GP_DATA.EXT_SALES)
             ELSE
                0
          END)
          "Manual Sales",
       SUM (
          CASE
             WHEN GP_DATA.PRICE_CATEGORY IN ('MANUAL', 'TOOLS', 'QUOTE')
             THEN
                (GP_DATA.AVG_COGS)
             ELSE
                0
          END)
          "Manual Cost",
       SUM (
          CASE
             WHEN GP_DATA.PRICE_CATEGORY IN ('MANUAL', 'TOOLS', 'QUOTE')
             THEN
                (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
             ELSE
                0
          END)
          "Manual GP$",
       ROUND (
          SUM (
             CASE
                WHEN GP_DATA.PRICE_CATEGORY IN ('MANUAL', 'TOOLS', 'QUOTE')
                THEN
                   (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                ELSE
                   0
             END)
          / SUM (
               CASE
                  WHEN GP_DATA.PRICE_CATEGORY IN ('MANUAL', 'TOOLS', 'QUOTE')
                  THEN
                     CASE
                        WHEN GP_DATA.EXT_SALES > 0 THEN (GP_DATA.EXT_SALES)
                        ELSE 1
                     END
                  ELSE
                     1
               END),
          3)
          "Manual GP%",
       ROUND (
          SUM (
             CASE
                WHEN GP_DATA.PRICE_CATEGORY IN ('MANUAL', 'TOOLS', 'QUOTE')
                THEN
                   (GP_DATA.EXT_SALES)
                ELSE
                   0
             END)
          / SUM (
               CASE
                  WHEN GP_DATA.SLS_SUBTOTAL > 0 THEN GP_DATA.SLS_SUBTOTAL
                  ELSE 1
               END),
          3)
          "Manual Use%$",
       ROUND (
          SUM (
             CASE
                WHEN GP_DATA.PRICE_CATEGORY IN ('MANUAL', 'TOOLS', 'QUOTE')
                THEN
                   (GP_DATA.INVOICE_LINES)
                ELSE
                   0
             END)
          / SUM (
               CASE
                  WHEN GP_DATA.INVOICE_LINES > 0 THEN GP_DATA.INVOICE_LINES
                  ELSE 1
               END),
          3)
          "Manual Use%#",
       ROUND (
          SUM (
             CASE
                WHEN GP_DATA.PRICE_CATEGORY IN ('MANUAL', 'TOOLS', 'QUOTE')
                THEN
                   (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                ELSE
                   0
             END)
          / SUM (
               CASE
                  WHEN GP_DATA.ROLLUP = 'Total'
                  THEN
                     (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                  ELSE
                     0
               END),
          3)
          "Manual Profit%$",
       SUM (
          CASE
             WHEN GP_DATA.PRICE_CATEGORY IN ('MANUAL', 'TOOLS', 'QUOTE')
             THEN
                (GP_DATA.INVOICE_LINES)
             ELSE
                0
          END)
          "Manual # Lines",
       SUM (
          CASE
             WHEN GP_DATA.PRICE_CATEGORY NOT IN
                     ('MATRIX',
                      'OVERRIDE',
                      'MANUAL',
                      'CREDITS',
                      'TOOLS',
                      'QUOTE')
             THEN
                (GP_DATA.EXT_SALES)
             ELSE
                0
          END)
          "Other Sales",
       SUM (
          CASE
             WHEN GP_DATA.PRICE_CATEGORY NOT IN
                     ('MATRIX',
                      'OVERRIDE',
                      'MANUAL',
                      'CREDITS',
                      'TOOLS',
                      'QUOTE')
             THEN
                (GP_DATA.AVG_COGS)
             ELSE
                0
          END)
          "Other Cost",
       SUM (
          CASE
             WHEN GP_DATA.PRICE_CATEGORY NOT IN
                     ('MATRIX',
                      'OVERRIDE',
                      'MANUAL',
                      'CREDITS',
                      'TOOLS',
                      'QUOTE')
             THEN
                (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
             ELSE
                0
          END)
          "Other GP$",
       ROUND (
          SUM (
             CASE
                WHEN GP_DATA.PRICE_CATEGORY NOT IN
                        ('MATRIX',
                         'OVERRIDE',
                         'MANUAL',
                         'CREDITS',
                         'TOOLS',
                         'QUOTE')
                THEN
                   (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                ELSE
                   0
             END)
          / SUM (
               CASE
                  WHEN GP_DATA.PRICE_CATEGORY NOT IN
                          ('MATRIX',
                           'OVERRIDE',
                           'MANUAL',
                           'CREDITS',
                           'TOOLS',
                           'QUOTE')
                  THEN
                     CASE
                        WHEN GP_DATA.EXT_SALES > 0 THEN (GP_DATA.EXT_SALES)
                        ELSE 1
                     END
                  ELSE
                     1
               END),
          3)
          "Other GP%",
       ROUND (
          SUM (
             CASE
                WHEN GP_DATA.PRICE_CATEGORY NOT IN
                        ('MATRIX',
                         'OVERRIDE',
                         'MANUAL',
                         'CREDITS',
                         'TOOLS',
                         'QUOTE')
                THEN
                   (GP_DATA.EXT_SALES)
                ELSE
                   0
             END)
          / SUM (
               CASE
                  WHEN GP_DATA.SLS_SUBTOTAL > 0 THEN GP_DATA.SLS_SUBTOTAL
                  ELSE 1
               END),
          3)
          "Other Use%$",
       ROUND (
          SUM (
             CASE
                WHEN GP_DATA.PRICE_CATEGORY NOT IN
                        ('MATRIX',
                         'OVERRIDE',
                         'MANUAL',
                         'CREDITS',
                         'TOOLS',
                         'QUOTE')
                THEN
                   (GP_DATA.INVOICE_LINES)
                ELSE
                   0
             END)
          / SUM (
               CASE
                  WHEN GP_DATA.INVOICE_LINES > 0 THEN GP_DATA.INVOICE_LINES
                  ELSE 1
               END),
          3)
          "Other Use%#",
       ROUND (
          SUM (
             CASE
                WHEN GP_DATA.PRICE_CATEGORY NOT IN
                        ('MATRIX',
                         'OVERRIDE',
                         'MANUAL',
                         'CREDITS',
                         'TOOLS',
                         'QUOTE')
                THEN
                   (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                ELSE
                   0
             END)
          / SUM (
               CASE
                  WHEN GP_DATA.ROLLUP = 'Total'
                  THEN
                     (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                  ELSE
                     0
               END),
          3)
          "Other Profit%$",
       SUM (
          CASE
             WHEN GP_DATA.PRICE_CATEGORY NOT IN
                     ('MATRIX',
                      'OVERRIDE',
                      'MANUAL',
                      'CREDITS',
                      'TOOLS',
                      'QUOTE')
             THEN
                (GP_DATA.INVOICE_LINES)
             ELSE
                0
          END)
          "Other # Lines",
       ROUND (
          SUM (
             CASE
                WHEN GP_DATA.PRICE_CATEGORY IN ('CREDITS')
                THEN
                   (GP_DATA.EXT_SALES)
                ELSE
                   0
             END)
          / SUM (
               CASE
                  WHEN GP_DATA.SLS_SUBTOTAL > 0 THEN GP_DATA.SLS_SUBTOTAL
                  ELSE 1
               END),
          3)
          "Credits Use%$",
       ROUND (
          SUM (
             CASE
                WHEN GP_DATA.PRICE_CATEGORY IN ('CREDITS')
                THEN
                   (GP_DATA.INVOICE_LINES)
                ELSE
                   0
             END)
          / SUM (
               CASE
                  WHEN GP_DATA.INVOICE_LINES > 0 THEN GP_DATA.INVOICE_LINES
                  ELSE 1
               END),
          3)
          "Credits Use%#",
       SUM (GP_DATA.SLS_FREIGHT - GP_DATA.AVG_COST_FREIGHT)
          "Freight Profit (Loss)",
       SUM (
          CASE
             WHEN GP_DATA.ROLLUP = 'Total'
             THEN
                (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
             ELSE
                0
          END)
          "Subtotal GP$"
  FROM AAD9606.GP_TRACKER_WRITER GP_DATA,
       (SELECT WD.ACCOUNT_NAME, WD.ACCOUNT_NUMBER_NK
          FROM DW_FEI.WAREHOUSE_DIMENSION WD
         WHERE (WD.ACTIVE_FLAG = 1) AND (WD.DELETE_DATE IS NULL)
        GROUP BY WD.ACCOUNT_NAME, WD.ACCOUNT_NUMBER_NK) ACCT
 WHERE     GP_DATA.ACCOUNT_NUMBER = ACCT.ACCOUNT_NUMBER_NK(+)
      -- AND GP_DATA.TYPE_OF_SALE IN ('Counter', 'Showroom', 'Showroom Direct')
       AND GP_DATA.WRITER IS NOT NULL
-- AND GP_DATA.EXT_SALES<>0
-- AND GP_DATA.INVOICE_LINES<>0
-- AND GP_DATA.SLS_SUBTOTAL<>0
 AND GP_DATA.YEARMONTH BETWEEN TO_CHAR (
                                        TRUNC (
                                           SYSDATE
                                           - NUMTOYMINTERVAL (12, 'MONTH'),
                                           'MONTH'),
                                        'YYYYMM')
                                 AND TO_CHAR (TRUNC (SYSDATE, 'MM') - 11,
                                              'YYYYMM')
       --AND GP_DATA.YEARMONTH = (SELECT MAX (GP_DATA.YEARMONTH)
       --                           FROM AAD9606.GP_TRACKER_12MO GP_DATA)
--AND GP_DATA.TYPE_OF_SALE='Counter'   
HAVING SUM (
            GP_DATA.SLS_SUBTOTAL
          + GP_DATA.SLS_FREIGHT
          + GP_DATA.SLS_MISC
          + GP_DATA.SLS_RESTOCK) <> 0
       AND SUM (
               CASE
                  WHEN GP_DATA.ROLLUP = 'Total'
                  THEN
                     (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                  ELSE
                     0
               END) <>0
GROUP BY GP_DATA.YEARMONTH,
         CASE
            WHEN GP_DATA.REGION IS NULL THEN ACCT.ACCOUNT_NAME
            WHEN GP_DATA.REGION = 'EAST' THEN 'EASTERN'
            WHEN GP_DATA.REGION = 'WEST' THEN 'WESTERN'
            ELSE GP_DATA.REGION
         END,
         GP_DATA.ACCOUNT_NUMBER,
         ACCT.ACCOUNT_NAME,
         GP_DATA.WRITER,
         CASE
            WHEN GP_DATA.TYPE_OF_SALE IN ('Showroom', 'Showroom Direct')
            THEN
               'Showroom'
            ELSE
               GP_DATA.TYPE_OF_SALE
         END
ORDER BY CASE
            WHEN GP_DATA.REGION IS NULL THEN ACCT.ACCOUNT_NAME
            WHEN GP_DATA.REGION = 'EAST' THEN 'EASTERN'
            WHEN GP_DATA.REGION = 'WEST' THEN 'WESTERN'
            ELSE GP_DATA.REGION
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
