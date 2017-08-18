SELECT AT_A_GLANCE.TYPE_OF_SALE,
       AT_A_GLANCE.RPT_PERIOD,
       AT_A_GLANCE.ACTUAL,
       --AT_A_GLANCE.REGION,
       SLS_DIV.DIVISION_NAME,
       SLS_DIV.REGION_NAME,
       AT_A_GLANCE.ACCOUNT_NUMBER,
       AT_A_GLANCE.ACCOUNT_NAME,
       AT_A_GLANCE.GPTRACK_KEY,
       AT_A_GLANCE."Total Sales",
       AT_A_GLANCE."Sales, Sub-Total",
       AT_A_GLANCE."Sales, Freight",
       AT_A_GLANCE."Sales, Misc",
       AT_A_GLANCE."Sales, Restock",
       AT_A_GLANCE."Total Cost",
       AT_A_GLANCE."Cost, Sub-Total",
       AT_A_GLANCE."Cost, Freight",
       AT_A_GLANCE."Cost, Misc",
       AT_A_GLANCE."Total GP$",
       AT_A_GLANCE."Total GP%",
       AT_A_GLANCE."Total # Lines",
       AT_A_GLANCE."Price Matrix Sales",
       AT_A_GLANCE."Price Matrix Cost",
       AT_A_GLANCE."Price Matrix GP$",
       AT_A_GLANCE."Price Matrix GP%",
       AT_A_GLANCE."Price Matrix Use%$",
       AT_A_GLANCE."Price Matrix Use%#",
       AT_A_GLANCE."Price Matrix Profit%$",
       AT_A_GLANCE."Price Matrix # Lines",
       AT_A_GLANCE."Contract Sales",
       AT_A_GLANCE."Contract Cost",
       AT_A_GLANCE."Contract GP$",
       AT_A_GLANCE."Contract GP%",
       AT_A_GLANCE."Contract Use%$",
       AT_A_GLANCE."Contract Use%#",
       AT_A_GLANCE."Contract Profit%$",
       AT_A_GLANCE."Contract # Lines",
       AT_A_GLANCE."Manual Sales",
       AT_A_GLANCE."Manual Cost",
       AT_A_GLANCE."Manual GP$",
       AT_A_GLANCE."Manual GP%",
       AT_A_GLANCE."Manual Use%$",
       AT_A_GLANCE."Manual Use%#",
       AT_A_GLANCE."Manual Profit%$",
       AT_A_GLANCE."Manual # Lines",
       AT_A_GLANCE."Other Sales",
       AT_A_GLANCE."Other Cost",
       AT_A_GLANCE."Other GP$",
       AT_A_GLANCE."Other GP%",
       AT_A_GLANCE."Other Use%$",
       AT_A_GLANCE."Other Use%#",
       AT_A_GLANCE."Other Profit%$",
       AT_A_GLANCE."Other # Lines",
       AT_A_GLANCE."Credits Use%$",
       AT_A_GLANCE."Credits Use%#",
       AT_A_GLANCE."Freight Profit (Loss)"
  FROM ( (SELECT 'THIS_YM' RPT_PERIOD,
                 TO_CHAR (GP_DATA.YEARMONTH) ACTUAL,
                 --DENSE_RANK () OVER (ORDER BY GP_DATA.YEARMONTH ASC) AS MM,
                 CASE
                    WHEN GP_DATA.REGION IS NULL THEN ACCT.ACCOUNT_NAME
                    WHEN GP_DATA.REGION = 'EAST' THEN 'EASTERN'
                    WHEN GP_DATA.REGION = 'WEST' THEN 'WESTERN'
                    ELSE GP_DATA.REGION
                 END
                    REGION,
                 GP_DATA.ACCOUNT_NUMBER,
                 ACCT.ACCOUNT_NAME,
                 CASE
                    WHEN GP_DATA.REGION IS NULL THEN ACCT.ACCOUNT_NAME
                    WHEN GP_DATA.REGION = 'EAST' THEN 'EASTERN'
                    WHEN GP_DATA.REGION = 'WEST' THEN 'WESTERN'
                    ELSE GP_DATA.REGION
                 END
                 || '*'
                 || GP_DATA.ACCOUNT_NUMBER
                 || '*'
                 || 'MM'
                 || TO_CHAR (
                       DENSE_RANK () OVER (ORDER BY GP_DATA.YEARMONTH ASC),
                       'FM00')
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
                       WHEN GP_DATA.PRICE_CATEGORY IN 'MATRIX'
                       THEN
                          (GP_DATA.EXT_SALES)
                       ELSE
                          0
                    END)
                    "Price Matrix Sales",
                 SUM (
                    CASE
                       WHEN GP_DATA.PRICE_CATEGORY IN 'MATRIX'
                       THEN
                          (GP_DATA.AVG_COGS)
                       ELSE
                          0
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
                                  WHEN GP_DATA.EXT_SALES > 0
                                  THEN
                                     (GP_DATA.EXT_SALES)
                                  ELSE
                                     1
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
                            WHEN GP_DATA.SLS_SUBTOTAL > 0
                            THEN
                               GP_DATA.SLS_SUBTOTAL
                            ELSE
                               1
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
                            WHEN GP_DATA.INVOICE_LINES > 0
                            THEN
                               GP_DATA.INVOICE_LINES
                            ELSE
                               1
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
                               CASE
                                  WHEN (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS) <>
                                          0
                                  THEN
                                     (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                                  ELSE
                                     1
                               END
                            ELSE
                               1
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
                                  WHEN GP_DATA.EXT_SALES > 0
                                  THEN
                                     (GP_DATA.EXT_SALES)
                                  ELSE
                                     1
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
                            WHEN GP_DATA.SLS_SUBTOTAL > 0
                            THEN
                               GP_DATA.SLS_SUBTOTAL
                            ELSE
                               1
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
                            WHEN GP_DATA.INVOICE_LINES > 0
                            THEN
                               GP_DATA.INVOICE_LINES
                            ELSE
                               1
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
                               CASE
                                  WHEN (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS) <>
                                          0
                                  THEN
                                     (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                                  ELSE
                                     1
                               END
                            ELSE
                               1
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
                       WHEN GP_DATA.PRICE_CATEGORY IN
                               ('MANUAL', 'TOOLS', 'QUOTE')
                       THEN
                          (GP_DATA.EXT_SALES)
                       ELSE
                          0
                    END)
                    "Manual Sales",
                 SUM (
                    CASE
                       WHEN GP_DATA.PRICE_CATEGORY IN
                               ('MANUAL', 'TOOLS', 'QUOTE')
                       THEN
                          (GP_DATA.AVG_COGS)
                       ELSE
                          0
                    END)
                    "Manual Cost",
                 SUM (
                    CASE
                       WHEN GP_DATA.PRICE_CATEGORY IN
                               ('MANUAL', 'TOOLS', 'QUOTE')
                       THEN
                          (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                       ELSE
                          0
                    END)
                    "Manual GP$",
                 ROUND (
                    SUM (
                       CASE
                          WHEN GP_DATA.PRICE_CATEGORY IN
                                  ('MANUAL', 'TOOLS', 'QUOTE')
                          THEN
                             (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                          ELSE
                             0
                       END)
                    / SUM (
                         CASE
                            WHEN GP_DATA.PRICE_CATEGORY IN
                                    ('MANUAL', 'TOOLS', 'QUOTE')
                            THEN
                               CASE
                                  WHEN GP_DATA.EXT_SALES > 0
                                  THEN
                                     (GP_DATA.EXT_SALES)
                                  ELSE
                                     1
                               END
                            ELSE
                               1
                         END),
                    3)
                    "Manual GP%",
                 ROUND (
                    SUM (
                       CASE
                          WHEN GP_DATA.PRICE_CATEGORY IN
                                  ('MANUAL', 'TOOLS', 'QUOTE')
                          THEN
                             (GP_DATA.EXT_SALES)
                          ELSE
                             0
                       END)
                    / SUM (
                         CASE
                            WHEN GP_DATA.SLS_SUBTOTAL > 0
                            THEN
                               GP_DATA.SLS_SUBTOTAL
                            ELSE
                               1
                         END),
                    3)
                    "Manual Use%$",
                 ROUND (
                    SUM (
                       CASE
                          WHEN GP_DATA.PRICE_CATEGORY IN
                                  ('MANUAL', 'TOOLS', 'QUOTE')
                          THEN
                             (GP_DATA.INVOICE_LINES)
                          ELSE
                             0
                       END)
                    / SUM (
                         CASE
                            WHEN GP_DATA.INVOICE_LINES > 0
                            THEN
                               GP_DATA.INVOICE_LINES
                            ELSE
                               1
                         END),
                    3)
                    "Manual Use%#",
                 ROUND (
                    SUM (
                       CASE
                          WHEN GP_DATA.PRICE_CATEGORY IN
                                  ('MANUAL', 'TOOLS', 'QUOTE')
                          THEN
                             (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                          ELSE
                             0
                       END)
                    / SUM (
                         CASE
                            WHEN GP_DATA.ROLLUP = 'Total'
                            THEN
                               CASE
                                  WHEN (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS) <>
                                          0
                                  THEN
                                     (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                                  ELSE
                                     1
                               END
                            ELSE
                               1
                         END),
                    3)
                    "Manual Profit%$",
                 SUM (
                    CASE
                       WHEN GP_DATA.PRICE_CATEGORY IN
                               ('MANUAL', 'TOOLS', 'QUOTE')
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
                                  WHEN GP_DATA.EXT_SALES > 0
                                  THEN
                                     (GP_DATA.EXT_SALES)
                                  ELSE
                                     1
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
                            WHEN GP_DATA.SLS_SUBTOTAL > 0
                            THEN
                               GP_DATA.SLS_SUBTOTAL
                            ELSE
                               1
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
                            WHEN GP_DATA.INVOICE_LINES > 0
                            THEN
                               GP_DATA.INVOICE_LINES
                            ELSE
                               1
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
                               CASE
                                  WHEN (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS) <>
                                          0
                                  THEN
                                     (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                                  ELSE
                                     1
                               END
                            ELSE
                               1
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
                            WHEN GP_DATA.SLS_SUBTOTAL > 0
                            THEN
                               GP_DATA.SLS_SUBTOTAL
                            ELSE
                               1
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
                            WHEN GP_DATA.INVOICE_LINES > 0
                            THEN
                               GP_DATA.INVOICE_LINES
                            ELSE
                               1
                         END),
                    3)
                    "Credits Use%#",
                 SUM (GP_DATA.SLS_FREIGHT - GP_DATA.AVG_COST_FREIGHT)
                    "Freight Profit (Loss)",
                 GP_DATA.TYPE_OF_SALE
            FROM AAD9606.GP_TRACKER_13MO GP_DATA,
                 (SELECT WD.ACCOUNT_NAME, WD.ACCOUNT_NUMBER_NK
                    FROM DW_FEI.WAREHOUSE_DIMENSION WD
                   WHERE (WD.ACTIVE_FLAG = 1) AND (WD.DELETE_DATE IS NULL)
                  GROUP BY WD.ACCOUNT_NAME, WD.ACCOUNT_NUMBER_NK) ACCT
           WHERE GP_DATA.ACCOUNT_NUMBER = ACCT.ACCOUNT_NUMBER_NK(+)
                 AND GP_DATA.YEARMONTH =
                        (SELECT MAX (GP_DATA.YEARMONTH)
                           FROM AAD9606.GP_TRACKER_13MO GP_DATA)
                 AND GP_DATA.TYPE_OF_SALE = 'Counter'
          HAVING SUM (GP_DATA.SLS_SUBTOTAL) <> 0
          GROUP BY GP_DATA.YEARMONTH,
                   GP_DATA.TYPE_OF_SALE,
                   CASE
                      WHEN GP_DATA.REGION IS NULL THEN ACCT.ACCOUNT_NAME
                      WHEN GP_DATA.REGION = 'EAST' THEN 'EASTERN'
                      WHEN GP_DATA.REGION = 'WEST' THEN 'WESTERN'
                      ELSE GP_DATA.REGION
                   END,
                   GP_DATA.ACCOUNT_NUMBER,
                   ACCT.ACCOUNT_NAME)
        UNION
        (SELECT 'ROLLING_12MO' RPT_PERIOD,
                TO_CHAR (COUNT (DISTINCT GP_DATA.YEARMONTH)) || '_MONTHS'
                   ACTUAL,
                --DENSE_RANK () OVER (ORDER BY GP_DATA.YEARMONTH ASC) AS MM,
                CASE
                   WHEN GP_DATA.REGION IS NULL THEN ACCT.ACCOUNT_NAME
                   WHEN GP_DATA.REGION = 'EAST' THEN 'EASTERN'
                   WHEN GP_DATA.REGION = 'WEST' THEN 'WESTERN'
                   ELSE GP_DATA.REGION
                END
                   REGION,
                GP_DATA.ACCOUNT_NUMBER,
                ACCT.ACCOUNT_NAME,
                CASE
                   WHEN GP_DATA.REGION IS NULL THEN ACCT.ACCOUNT_NAME
                   WHEN GP_DATA.REGION = 'EAST' THEN 'EASTERN'
                   WHEN GP_DATA.REGION = 'WEST' THEN 'WESTERN'
                   ELSE GP_DATA.REGION
                END
                || '*'
                || GP_DATA.ACCOUNT_NUMBER
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
                      WHEN GP_DATA.PRICE_CATEGORY IN 'MATRIX'
                      THEN
                         (GP_DATA.EXT_SALES)
                      ELSE
                         0
                   END)
                   "Price Matrix Sales",
                SUM (
                   CASE
                      WHEN GP_DATA.PRICE_CATEGORY IN 'MATRIX'
                      THEN
                         (GP_DATA.AVG_COGS)
                      ELSE
                         0
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
                                 WHEN GP_DATA.EXT_SALES > 0
                                 THEN
                                    (GP_DATA.EXT_SALES)
                                 ELSE
                                    1
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
                           WHEN GP_DATA.SLS_SUBTOTAL > 0
                           THEN
                              GP_DATA.SLS_SUBTOTAL
                           ELSE
                              1
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
                           WHEN GP_DATA.INVOICE_LINES > 0
                           THEN
                              GP_DATA.INVOICE_LINES
                           ELSE
                              1
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
                              CASE
                                 WHEN (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS) <>
                                         0
                                 THEN
                                    (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                                 ELSE
                                    1
                              END
                           ELSE
                              1
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
                                 WHEN GP_DATA.EXT_SALES > 0
                                 THEN
                                    (GP_DATA.EXT_SALES)
                                 ELSE
                                    1
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
                           WHEN GP_DATA.SLS_SUBTOTAL > 0
                           THEN
                              GP_DATA.SLS_SUBTOTAL
                           ELSE
                              1
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
                           WHEN GP_DATA.INVOICE_LINES > 0
                           THEN
                              GP_DATA.INVOICE_LINES
                           ELSE
                              1
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
                              CASE
                                 WHEN (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS) <>
                                         0
                                 THEN
                                    (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                                 ELSE
                                    1
                              END
                           ELSE
                              1
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
                      WHEN GP_DATA.PRICE_CATEGORY IN
                              ('MANUAL', 'TOOLS', 'QUOTE')
                      THEN
                         (GP_DATA.EXT_SALES)
                      ELSE
                         0
                   END)
                   "Manual Sales",
                SUM (
                   CASE
                      WHEN GP_DATA.PRICE_CATEGORY IN
                              ('MANUAL', 'TOOLS', 'QUOTE')
                      THEN
                         (GP_DATA.AVG_COGS)
                      ELSE
                         0
                   END)
                   "Manual Cost",
                SUM (
                   CASE
                      WHEN GP_DATA.PRICE_CATEGORY IN
                              ('MANUAL', 'TOOLS', 'QUOTE')
                      THEN
                         (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                      ELSE
                         0
                   END)
                   "Manual GP$",
                ROUND (
                   SUM (
                      CASE
                         WHEN GP_DATA.PRICE_CATEGORY IN
                                 ('MANUAL', 'TOOLS', 'QUOTE')
                         THEN
                            (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                         ELSE
                            0
                      END)
                   / SUM (
                        CASE
                           WHEN GP_DATA.PRICE_CATEGORY IN
                                   ('MANUAL', 'TOOLS', 'QUOTE')
                           THEN
                              CASE
                                 WHEN GP_DATA.EXT_SALES > 0
                                 THEN
                                    (GP_DATA.EXT_SALES)
                                 ELSE
                                    1
                              END
                           ELSE
                              1
                        END),
                   3)
                   "Manual GP%",
                ROUND (
                   SUM (
                      CASE
                         WHEN GP_DATA.PRICE_CATEGORY IN
                                 ('MANUAL', 'TOOLS', 'QUOTE')
                         THEN
                            (GP_DATA.EXT_SALES)
                         ELSE
                            0
                      END)
                   / SUM (
                        CASE
                           WHEN GP_DATA.SLS_SUBTOTAL > 0
                           THEN
                              GP_DATA.SLS_SUBTOTAL
                           ELSE
                              1
                        END),
                   3)
                   "Manual Use%$",
                ROUND (
                   SUM (
                      CASE
                         WHEN GP_DATA.PRICE_CATEGORY IN
                                 ('MANUAL', 'TOOLS', 'QUOTE')
                         THEN
                            (GP_DATA.INVOICE_LINES)
                         ELSE
                            0
                      END)
                   / SUM (
                        CASE
                           WHEN GP_DATA.INVOICE_LINES > 0
                           THEN
                              GP_DATA.INVOICE_LINES
                           ELSE
                              1
                        END),
                   3)
                   "Manual Use%#",
                ROUND (
                   SUM (
                      CASE
                         WHEN GP_DATA.PRICE_CATEGORY IN
                                 ('MANUAL', 'TOOLS', 'QUOTE')
                         THEN
                            (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                         ELSE
                            0
                      END)
                   / SUM (
                        CASE
                           WHEN GP_DATA.ROLLUP = 'Total'
                           THEN
                              CASE
                                 WHEN (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS) <>
                                         0
                                 THEN
                                    (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                                 ELSE
                                    1
                              END
                           ELSE
                              1
                        END),
                   3)
                   "Manual Profit%$",
                SUM (
                   CASE
                      WHEN GP_DATA.PRICE_CATEGORY IN
                              ('MANUAL', 'TOOLS', 'QUOTE')
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
                                 WHEN GP_DATA.EXT_SALES > 0
                                 THEN
                                    (GP_DATA.EXT_SALES)
                                 ELSE
                                    1
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
                           WHEN GP_DATA.SLS_SUBTOTAL > 0
                           THEN
                              GP_DATA.SLS_SUBTOTAL
                           ELSE
                              1
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
                           WHEN GP_DATA.INVOICE_LINES > 0
                           THEN
                              GP_DATA.INVOICE_LINES
                           ELSE
                              1
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
                              CASE
                                 WHEN (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS) <>
                                         0
                                 THEN
                                    (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                                 ELSE
                                    1
                              END
                           ELSE
                              1
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
                           WHEN GP_DATA.SLS_SUBTOTAL > 0
                           THEN
                              GP_DATA.SLS_SUBTOTAL
                           ELSE
                              1
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
                           WHEN GP_DATA.INVOICE_LINES > 0
                           THEN
                              GP_DATA.INVOICE_LINES
                           ELSE
                              1
                        END),
                   3)
                   "Credits Use%#",
                SUM (GP_DATA.SLS_FREIGHT - GP_DATA.AVG_COST_FREIGHT)
                   "Freight Profit (Loss)",
                GP_DATA.TYPE_OF_SALE
           FROM AAD9606.GP_TRACKER_13MO GP_DATA,
                (SELECT WD.ACCOUNT_NAME, WD.ACCOUNT_NUMBER_NK
                   FROM DW_FEI.WAREHOUSE_DIMENSION WD
                  WHERE (WD.ACTIVE_FLAG = 1) AND (WD.DELETE_DATE IS NULL)
                 GROUP BY WD.ACCOUNT_NAME, WD.ACCOUNT_NUMBER_NK) ACCT
          WHERE GP_DATA.ACCOUNT_NUMBER = ACCT.ACCOUNT_NUMBER_NK(+)
                AND GP_DATA.YEARMONTH BETWEEN TO_CHAR (
                                                 TRUNC (
                                                    SYSDATE
                                                    - NUMTOYMINTERVAL (
                                                         12,
                                                         'MONTH'),
                                                    'MONTH'),
                                                 'YYYYMM')
                                          AND TO_CHAR (
                                                 TRUNC (SYSDATE, 'MM') - 1,
                                                 'YYYYMM')
                /*AND GP_DATA.YEARMONTH BETWEEN (SELECT FISCAL_YEARMONTH_START
                                                 FROM SALES_MART.TIME_DIM)
                                          AND (SELECT MAX (YEARMONTH)
                                                 FROM AAD9606.GP_TRACKER_12MO)*/
                AND GP_DATA.TYPE_OF_SALE = 'Counter'
         HAVING SUM (GP_DATA.SLS_SUBTOTAL) <> 0
         GROUP BY GP_DATA.TYPE_OF_SALE,                   --GP_DATA.YEARMONTH,
                  CASE
                     WHEN GP_DATA.REGION IS NULL THEN ACCT.ACCOUNT_NAME
                     WHEN GP_DATA.REGION = 'EAST' THEN 'EASTERN'
                     WHEN GP_DATA.REGION = 'WEST' THEN 'WESTERN'
                     ELSE GP_DATA.REGION
                  END,
                  GP_DATA.ACCOUNT_NUMBER,
                  ACCT.ACCOUNT_NAME)
        UNION
        (SELECT 'LAST_YM' RPT_PERIOD,
                TO_CHAR (GP_DATA.YEARMONTH) ACTUAL,
                --DENSE_RANK () OVER (ORDER BY GP_DATA.YEARMONTH ASC) AS MM,
                CASE
                   WHEN GP_DATA.REGION IS NULL THEN ACCT.ACCOUNT_NAME
                   WHEN GP_DATA.REGION = 'EAST' THEN 'EASTERN'
                   WHEN GP_DATA.REGION = 'WEST' THEN 'WESTERN'
                   ELSE GP_DATA.REGION
                END
                   REGION,
                GP_DATA.ACCOUNT_NUMBER,
                ACCT.ACCOUNT_NAME,
                CASE
                   WHEN GP_DATA.REGION IS NULL THEN ACCT.ACCOUNT_NAME
                   WHEN GP_DATA.REGION = 'EAST' THEN 'EASTERN'
                   WHEN GP_DATA.REGION = 'WEST' THEN 'WESTERN'
                   ELSE GP_DATA.REGION
                END
                || '*'
                || GP_DATA.ACCOUNT_NUMBER
                || '*'
                || 'MM'
                || TO_CHAR (
                      DENSE_RANK () OVER (ORDER BY GP_DATA.YEARMONTH ASC),
                      'FM00')
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
                      WHEN GP_DATA.PRICE_CATEGORY IN 'MATRIX'
                      THEN
                         (GP_DATA.EXT_SALES)
                      ELSE
                         0
                   END)
                   "Price Matrix Sales",
                SUM (
                   CASE
                      WHEN GP_DATA.PRICE_CATEGORY IN 'MATRIX'
                      THEN
                         (GP_DATA.AVG_COGS)
                      ELSE
                         0
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
                                 WHEN GP_DATA.EXT_SALES > 0
                                 THEN
                                    (GP_DATA.EXT_SALES)
                                 ELSE
                                    1
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
                           WHEN GP_DATA.SLS_SUBTOTAL > 0
                           THEN
                              GP_DATA.SLS_SUBTOTAL
                           ELSE
                              1
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
                           WHEN GP_DATA.INVOICE_LINES > 0
                           THEN
                              GP_DATA.INVOICE_LINES
                           ELSE
                              1
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
                              CASE
                                 WHEN (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS) <>
                                         0
                                 THEN
                                    (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                                 ELSE
                                    1
                              END
                           ELSE
                              1
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
                                 WHEN GP_DATA.EXT_SALES > 0
                                 THEN
                                    (GP_DATA.EXT_SALES)
                                 ELSE
                                    1
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
                           WHEN GP_DATA.SLS_SUBTOTAL > 0
                           THEN
                              GP_DATA.SLS_SUBTOTAL
                           ELSE
                              1
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
                           WHEN GP_DATA.INVOICE_LINES > 0
                           THEN
                              GP_DATA.INVOICE_LINES
                           ELSE
                              1
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
                              CASE
                                 WHEN (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS) <>
                                         0
                                 THEN
                                    (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                                 ELSE
                                    1
                              END
                           ELSE
                              1
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
                      WHEN GP_DATA.PRICE_CATEGORY IN
                              ('MANUAL', 'TOOLS', 'QUOTE')
                      THEN
                         (GP_DATA.EXT_SALES)
                      ELSE
                         0
                   END)
                   "Manual Sales",
                SUM (
                   CASE
                      WHEN GP_DATA.PRICE_CATEGORY IN
                              ('MANUAL', 'TOOLS', 'QUOTE')
                      THEN
                         (GP_DATA.AVG_COGS)
                      ELSE
                         0
                   END)
                   "Manual Cost",
                SUM (
                   CASE
                      WHEN GP_DATA.PRICE_CATEGORY IN
                              ('MANUAL', 'TOOLS', 'QUOTE')
                      THEN
                         (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                      ELSE
                         0
                   END)
                   "Manual GP$",
                ROUND (
                   SUM (
                      CASE
                         WHEN GP_DATA.PRICE_CATEGORY IN
                                 ('MANUAL', 'TOOLS', 'QUOTE')
                         THEN
                            (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                         ELSE
                            0
                      END)
                   / SUM (
                        CASE
                           WHEN GP_DATA.PRICE_CATEGORY IN
                                   ('MANUAL', 'TOOLS', 'QUOTE')
                           THEN
                              CASE
                                 WHEN GP_DATA.EXT_SALES > 0
                                 THEN
                                    (GP_DATA.EXT_SALES)
                                 ELSE
                                    1
                              END
                           ELSE
                              1
                        END),
                   3)
                   "Manual GP%",
                ROUND (
                   SUM (
                      CASE
                         WHEN GP_DATA.PRICE_CATEGORY IN
                                 ('MANUAL', 'TOOLS', 'QUOTE')
                         THEN
                            (GP_DATA.EXT_SALES)
                         ELSE
                            0
                      END)
                   / SUM (
                        CASE
                           WHEN GP_DATA.SLS_SUBTOTAL > 0
                           THEN
                              GP_DATA.SLS_SUBTOTAL
                           ELSE
                              1
                        END),
                   3)
                   "Manual Use%$",
                ROUND (
                   SUM (
                      CASE
                         WHEN GP_DATA.PRICE_CATEGORY IN
                                 ('MANUAL', 'TOOLS', 'QUOTE')
                         THEN
                            (GP_DATA.INVOICE_LINES)
                         ELSE
                            0
                      END)
                   / SUM (
                        CASE
                           WHEN GP_DATA.INVOICE_LINES > 0
                           THEN
                              GP_DATA.INVOICE_LINES
                           ELSE
                              1
                        END),
                   3)
                   "Manual Use%#",
                ROUND (
                   SUM (
                      CASE
                         WHEN GP_DATA.PRICE_CATEGORY IN
                                 ('MANUAL', 'TOOLS', 'QUOTE')
                         THEN
                            (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                         ELSE
                            0
                      END)
                   / SUM (
                        CASE
                           WHEN GP_DATA.ROLLUP = 'Total'
                           THEN
                              CASE
                                 WHEN (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS) <>
                                         0
                                 THEN
                                    (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                                 ELSE
                                    1
                              END
                           ELSE
                              1
                        END),
                   3)
                   "Manual Profit%$",
                SUM (
                   CASE
                      WHEN GP_DATA.PRICE_CATEGORY IN
                              ('MANUAL', 'TOOLS', 'QUOTE')
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
                                 WHEN GP_DATA.EXT_SALES > 0
                                 THEN
                                    (GP_DATA.EXT_SALES)
                                 ELSE
                                    1
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
                           WHEN GP_DATA.SLS_SUBTOTAL > 0
                           THEN
                              GP_DATA.SLS_SUBTOTAL
                           ELSE
                              1
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
                           WHEN GP_DATA.INVOICE_LINES > 0
                           THEN
                              GP_DATA.INVOICE_LINES
                           ELSE
                              1
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
                              CASE
                                 WHEN (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS) <>
                                         0
                                 THEN
                                    (GP_DATA.EXT_SALES - GP_DATA.AVG_COGS)
                                 ELSE
                                    1
                              END
                           ELSE
                              1
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
                           WHEN GP_DATA.SLS_SUBTOTAL > 0
                           THEN
                              GP_DATA.SLS_SUBTOTAL
                           ELSE
                              1
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
                           WHEN GP_DATA.INVOICE_LINES > 0
                           THEN
                              GP_DATA.INVOICE_LINES
                           ELSE
                              1
                        END),
                   3)
                   "Credits Use%#",
                SUM (GP_DATA.SLS_FREIGHT - GP_DATA.AVG_COST_FREIGHT)
                   "Freight Profit (Loss)",
                GP_DATA.TYPE_OF_SALE
           FROM AAD9606.GP_TRACKER_13MO GP_DATA,
                (SELECT WD.ACCOUNT_NAME, WD.ACCOUNT_NUMBER_NK
                   FROM DW_FEI.WAREHOUSE_DIMENSION WD
                  WHERE (WD.ACTIVE_FLAG = 1) AND (WD.DELETE_DATE IS NULL)
                 GROUP BY WD.ACCOUNT_NAME, WD.ACCOUNT_NUMBER_NK) ACCT
          WHERE GP_DATA.ACCOUNT_NUMBER = ACCT.ACCOUNT_NUMBER_NK(+)
                AND GP_DATA.YEARMONTH =
                       (SELECT MIN (GP_DATA.YEARMONTH)
                          FROM AAD9606.GP_TRACKER_13MO GP_DATA)
                AND GP_DATA.TYPE_OF_SALE = 'Counter'
         HAVING SUM (GP_DATA.SLS_SUBTOTAL) <> 0
         GROUP BY GP_DATA.YEARMONTH,
                  GP_DATA.TYPE_OF_SALE,
                  CASE
                     WHEN GP_DATA.REGION IS NULL THEN ACCT.ACCOUNT_NAME
                     WHEN GP_DATA.REGION = 'EAST' THEN 'EASTERN'
                     WHEN GP_DATA.REGION = 'WEST' THEN 'WESTERN'
                     ELSE GP_DATA.REGION
                  END,
                  GP_DATA.ACCOUNT_NUMBER,
                  ACCT.ACCOUNT_NAME)) AT_A_GLANCE,
       EBUSINESS.SALES_DIVISIONS SLS_DIV
 WHERE AT_A_GLANCE.ACCOUNT_NAME = SLS_DIV.ACCOUNT_NAME(+)
ORDER BY AT_A_GLANCE.RPT_PERIOD ASC, SLS_DIV.REGION_NAME ASC
