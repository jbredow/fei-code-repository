TRUNCATE TABLE AAE0376.GP_TRACKER_WRITER_YTD;
DROP TABLE AAE0376.GP_TRACKER_WRITER_YTD;

CREATE TABLE AAE0376.GP_TRACKER_WRITER_YTD

AS
   (SELECT LINE_HIST.YEARMONTH,
           LINE_HIST.REGION,
           LINE_HIST.ACCOUNT_NUMBER,
           LINE_HIST.WAREHOUSE_NUMBER,
           LINE_HIST.KOB,
           LINE_HIST.WRITER,
           LINE_HIST.TYPE_OF_SALE,
           COUNT (DISTINCT LINE_HIST.INVOICE_NUMBER_NK) INVOICE_CNT,
           SUM (LINE_HIST.INVOICE_LINES) invoice_lines,
           SUM (LINE_HIST.EXT_AVG_COGS_AMOUNT) avg_cogs,
           SUM (LINE_HIST.EXT_ACTUAL_COGS_AMOUNT) actual_cogs,
           SUM (LINE_HIST.EXT_SALES_AMOUNT) ext_sales,
           LINE_HIST.PRICE_CATEGORY,
       CASE
          WHEN LINE_HIST.ORDER_ENTRY_DATE BETWEEN COALESCE (
                                                     PR_OVR_JOB.INSERT_TIMESTAMP,
                                                     PR_OVR_BASE.INSERT_TIMESTAMP)
                                              AND COALESCE (
                                                     PR_OVR_JOB.EXPIRE_DATE,
                                                     PR_OVR_BASE.EXPIRE_DATE,
                                                     LINE_HIST.ORDER_ENTRY_DATE)
          THEN
             CASE
                WHEN LINE_HIST.UNIT_NET_PRICE_AMOUNT =
                        COALESCE (PR_OVR_JOB.MULTIPLIER,
                                  PR_OVR_BASE.MULTIPLIER)
                THEN
                   'OVERRIDE'
                WHEN LINE_HIST.UNIT_NET_PRICE_AMOUNT =
                        (  TRUNC (
                              COALESCE (PR_OVR_JOB.MULTIPLIER,
                                        PR_OVR_BASE.MULTIPLIER),
                              2)
                         + .01)
                THEN
                   'OVERRIDE'
                WHEN LINE_HIST.UNIT_NET_PRICE_AMOUNT =
                        (ROUND (
                            COALESCE (PR_OVR_JOB.MULTIPLIER,
                                      PR_OVR_BASE.MULTIPLIER),
                            2))
                THEN
                   'OVERRIDE'
                WHEN LINE_HIST.UNIT_NET_PRICE_AMOUNT =
                        (  TRUNC (
                              COALESCE (PR_OVR_JOB.MULTIPLIER,
                                        PR_OVR_BASE.MULTIPLIER),
                              1)
                         + .1)
                THEN
                   'OVERRIDE'
                WHEN LINE_HIST.UNIT_NET_PRICE_AMOUNT =
                          FLOOR (
                             COALESCE (PR_OVR_JOB.MULTIPLIER,
                                       PR_OVR_BASE.MULTIPLIER))
                        + 1
                THEN
                   'OVERRIDE'
                WHEN REPLACE (LINE_HIST.PRICE_FORMULA, '0.', '.') =
                        REPLACE (
                           COALESCE (PR_OVR_JOB.FORMULA, PR_OVR_BASE.FORMULA),
                           '0.',
                           '.')
                THEN
                   'OVERRIDE'
             END
       END
          PRICE_CATEGORY_OVR_PR,
       CASE
          WHEN LINE_HIST.ORDER_ENTRY_DATE BETWEEN COALESCE (
                                                     GR_OVR_JOB.INSERT_TIMESTAMP,
                                                     GR_OVR_BASE.INSERT_TIMESTAMP)
                                              AND COALESCE (
                                                     GR_OVR_JOB.EXPIRE_DATE,
                                                     GR_OVR_BASE.EXPIRE_DATE,
                                                     LINE_HIST.ORDER_ENTRY_DATE)
          THEN
             CASE
                WHEN REPLACE (LINE_HIST.PRICE_FORMULA, '0.', '.') =
                        REPLACE (
                           COALESCE (GR_OVR_JOB.FORMULA, GR_OVR_BASE.FORMULA),
                           '0.',
                           '.')
                THEN
                   'OVERRIDE'
                ELSE
                   NULL
             END
       END
          PRICE_CATEGORY_OVR_GR,
           'Subtotal' AS ROLLUP,
           (0) sls_total,
           (0) sls_subtotal,
           (0) sls_freight,
           (0) sls_misc,
           (0) sls_restock,
           (0) avg_cost_subtotal,
           (0) avg_cost_freight,
           (0) avg_cost_misc
      FROM (SELECT SUBSTR (ihf.YEARMONTH, 0, 4) YYYY,
                   ihf.YEARMONTH,
                   ps.division_name REGION,
                   ps.ACCOUNT_NUMBER_NK ACCOUNT_NUMBER,
                   ps.WAREHOUSE_NUMBER_nk warehouse_number,
                   ihf.ORDER_ENTRY_DATE,
                   ihf.PROCESS_DATE,
                   ihf.CONTRACT_NUMBER,
                   ihf.INVOICE_NUMBER_GK,
                   ihf.INVOICE_NUMBER_NK,
                   ihf.CUSTOMER_ACCOUNT_GK,
                   NVL (ihf.WRITER, ihf.OML_ASSOC_INI) WRITER,
                   PROD.DISCOUNT_GROUP_NK,
                   CUST.MAIN_CUSTOMER_NK,
                   ilf.PRODUCT_NUMBER_NK,
                   ilf.PRICE_CODE,
                   ilf.UNIT_NET_PRICE_AMOUNT,
                   ilf.PRICE_FORMULA,
                   DECODE (ps.DIVISION_NAME,
                   'EAST REGION', 'BLENDED',
                   'WEST REGION', 'BLENDED',
                   'NORTH CENTRAL REGION', 'BLENDED',
                   'SOUTH CENTRAL REGION', 'BLENDED',
                   'INDUSTRIAL REGION', 'INDUSTRIAL',
                   'WATERWORKS REGION', 'WW',
                   'HVAC REGION', 'HVAC',
                    ps.DIVISION_NAME)
                      KOB,
                   DECODE (ihf.SALE_TYPE,
                           '1', 'Our Truck',
                           '2', 'Counter',
                           '3', 'Direct',
                           '4', 'Counter',
                           '5', 'Credit Memo',
                           '6', 'Showroom',
                           '7', 'Showroom Direct',
                           '8', 'eBusiness')
                      TYPE_OF_SALE,
                   CASE WHEN ilf.SHIPPED_QTY > 0 THEN 1 ELSE 0 END
                      invoice_lines,
                   ilf.EXT_AVG_COGS_AMOUNT,
                   ilf.EXT_ACTUAL_COGS_AMOUNT,
                   ilf.EXT_SALES_AMOUNT,
                   CASE
                      WHEN ihf.order_code = 'IC'
                      THEN
                         'CREDITS'
                      WHEN ihf.order_code = 'SP'
                      THEN
                         'SPECIALS'
                      WHEN ilf.price_code = 'Q'
                      THEN
                         CASE
                            WHEN ilf.UNIT_NET_PRICE_AMOUNT = ilf.MATRIX_PRICE
                            THEN
                               'MATRIX'
                            WHEN ilf.UNIT_NET_PRICE_AMOUNT = ilf.MATRIX
                            THEN
                               'MATRIX_BID'
                            WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                    ROUND (ilf.MATRIX_PRICE, 2)
                            THEN
                               'MATRIX'
                            WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                    ROUND (ilf.MATRIX, 2)
                            THEN
                               'MATRIX_BID'
                            WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                    (TRUNC (ilf.MATRIX_PRICE, 2) + .01)
                            THEN
                               'MATRIX'
                            WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                    (TRUNC (ilf.MATRIX, 2) + .01)
                            THEN
                               'MATRIX_BID'
                            WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                    (TRUNC (ilf.MATRIX_PRICE, 1) + .1)
                            THEN
                               'MATRIX'
                            WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                    (TRUNC (ilf.MATRIX, 1) + .1)
                            THEN
                               'MATRIX_BID'
                            WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                    FLOOR (ilf.MATRIX_PRICE) + 1
                            THEN
                               'MATRIX'
                            WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                    FLOOR (ilf.MATRIX) + 1
                            THEN
                               'MATRIX_BID'
                            ELSE
                               'QUOTE'
                         END
                      WHEN REGEXP_LIKE (ilf.price_code, '[0-9]?[0-9]?[0-9]')
                      THEN
                         'MATRIX'
                      WHEN ilf.price_code IN ('FC', 'PM')
                      THEN
                         'MATRIX'
                      WHEN ilf.price_code LIKE 'M%'
                      THEN
                         'MATRIX'
                      WHEN ilf.price_formula IN ('CPA', 'CPO')
                      THEN
                         'OVERRIDE'
                      WHEN ilf.price_code IN
                              ('PR',
                               'GR',
                               'CB',
                               'GJ',
                               'PJ',
                               '*G',
                               '*P',
                               'G*',
                               'P*',
                               'G',
                               'GJ',
                               'P')
                      THEN
                         'OVERRIDE'
                      WHEN ilf.price_code IN
                              ('GI', 'GPC', 'HPF', 'HPN', 'NC')
                      THEN
                         'MANUAL'
                      WHEN ilf.price_code = '*E'
                      THEN
                         'OTH/ERROR'
                      WHEN ilf.price_code = 'SKC'
                      THEN
                         'OTH/ERROR'
                      WHEN ilf.price_code IN ('%', '$', 'N', 'F', 'B', 'PO')
                      THEN
                         'TOOLS'
                      WHEN ilf.price_code IS NULL
                      THEN
                         'MANUAL'
                      WHEN ilf.price_code IN ('R', 'N/A')
                      THEN
                         CASE
                            WHEN ilf.UNIT_NET_PRICE_AMOUNT = ilf.MATRIX_PRICE
                            THEN
                               'MATRIX'
                            WHEN ilf.UNIT_NET_PRICE_AMOUNT = ilf.MATRIX
                            THEN
                               'MATRIX_BID'
                            WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                    ROUND (ilf.MATRIX_PRICE, 2)
                            THEN
                               'MATRIX'
                            WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                    ROUND (ilf.MATRIX, 2)
                            THEN
                               'MATRIX_BID'
                            WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                    (TRUNC (ilf.MATRIX_PRICE, 2) + .01)
                            THEN
                               'MATRIX'
                            WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                    (TRUNC (ilf.MATRIX, 2) + .01)
                            THEN
                               'MATRIX_BID'
                            WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                    (TRUNC (ilf.MATRIX_PRICE, 1) + .1)
                            THEN
                               'MATRIX'
                            WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                    (TRUNC (ilf.MATRIX, 1) + .1)
                            THEN
                               'MATRIX_BID'
                            WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                    FLOOR (ilf.MATRIX_PRICE) + 1
                            THEN
                               'MATRIX'
                            WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                    FLOOR (ilf.MATRIX) + 1
                            THEN
                               'MATRIX_BID'
                            ELSE
                               'MANUAL'
                         END
                      ELSE
                         'MANUAL'
                   END
                      AS PRICE_CATEGORY
              FROM DW_FEI.INVOICE_LINE_FACT ilf,
                   DW_FEI.INVOICE_HEADER_FACT ihf,
                   DW_FEI.CUSTOMER_DIMENSION cust,
                   DW_FEI.PRODUCT_DIMENSION PROD,
                   sales_mart.sales_warehouse_dim ps
             WHERE ihf.CUSTOMER_ACCOUNT_GK = cust.CUSTOMER_GK
                   AND (ilf.INVOICE_NUMBER_GK = ihf.INVOICE_NUMBER_GK)
                   AND ILF.PRODUCT_GK = PROD.PRODUCT_GK(+)
                   AND TO_CHAR (ihf.WAREHOUSE_NUMBER) = TO_CHAR (ps.WAREHOUSE_NUMBER_NK)
                   AND ihf.IC_FLAG = 0
                   AND ihf.PO_WAREHOUSE_NUMBER IS NULL
                   AND NVL (ihf.CONSIGN_TYPE, 'N') <> 'R'
                   AND ILF.YEARMONTH BETWEEN TO_CHAR (
                                                 TRUNC (
                                                    SYSDATE
                                                    - NUMTOYMINTERVAL (
                                                         13,
                                                         'MONTH'),
                                                    'MONTH'),
                                                 'YYYYMM')
                                                   AND TO_CHAR (TRUNC (SYSDATE, 'MM') - 1, 'YYYYMM')
                   AND IHF.YEARMONTH BETWEEN TO_CHAR (
                                                 TRUNC (
                                                    SYSDATE
                                                    - NUMTOYMINTERVAL (
                                                         13,
                                                         'MONTH'),
                                                    'MONTH'),
                                                 'YYYYMM')
                                                   AND TO_CHAR (TRUNC (SYSDATE, 'MM') - 1, 'YYYYMM')
                   AND DECODE (NVL (cust.ar_gl_number, '9999'),
                               '1320', 0,
                               '1360', 0,
                               '1380', 0,
                               '1400', 0,
                               '1401', 0,
                               '1500', 0,
                               '4000', 0,
                               '7100', 0,
                               '9999', 0,
                               1) <> 0) LINE_HIST
           LEFT OUTER JOIN (SELECT COD.BASIS,
                       COD.BRANCH_NUMBER_NK,
                       COD.CONTRACT_ID,
                       COD.CUSTOMER_GK,
                       COD.CUSTOMER_NK,
                       COD.DISC_GROUP,
                       COD.INSERT_TIMESTAMP,
                       COD.EXPIRE_DATE,
                       COD.MASTER_PRODUCT,
                       COD.MULTIPLIER,
                       COD.OPERATOR_USED,
                       COD.OVERRIDE_ID_NK,
                       COD.OVERRIDE_TYPE,
                          COD.BASIS
                       || COD.OPERATOR_USED
                       || '0'
                       || COD.MULTIPLIER
                          FORMULA
                  FROM DW_FEI.CUSTOMER_OVERRIDE_DIMENSION COD,
                       DW_FEI.CUSTOMER_DIMENSION CUST
                 WHERE     COD.OVERRIDE_TYPE = 'G'
                       AND COD.CUSTOMER_GK = CUST.CUSTOMER_GK
                       AND COD.DELETE_DATE IS NULL
                       AND CUST.JOB_YN = 'Y'
                       AND NVL (COD.EXPIRE_DATE, SYSDATE) >= (SYSDATE - 395)) GR_OVR_JOB
                  ON (    LINE_HIST.DISCOUNT_GROUP_NK =
                             (LTRIM (GR_OVR_JOB.DISC_GROUP, '0'))
                      AND LINE_HIST.ACCOUNT_NUMBER =
                             GR_OVR_JOB.BRANCH_NUMBER_NK
                      AND LINE_HIST.CUSTOMER_ACCOUNT_GK =
                             GR_OVR_JOB.CUSTOMER_GK
                      AND NVL (LINE_HIST.CONTRACT_NUMBER, 'DEFAULT_MATCH') =
                             NVL (GR_OVR_JOB.CONTRACT_ID, 'DEFAULT_MATCH'))
               LEFT OUTER JOIN
               (SELECT COD.BASIS,
                       COD.BRANCH_NUMBER_NK,
                       COD.CONTRACT_ID,
                       COD.CUSTOMER_GK,
                       COD.CUSTOMER_NK,
                       COD.DISC_GROUP,
                       COD.INSERT_TIMESTAMP,
                       COD.EXPIRE_DATE,
                       COD.MASTER_PRODUCT,
                       COD.MULTIPLIER,
                       COD.OPERATOR_USED,
                       COD.OVERRIDE_ID_NK,
                       COD.OVERRIDE_TYPE,
                          COD.BASIS
                       || COD.OPERATOR_USED
                       || '0'
                       || COD.MULTIPLIER
                          FORMULA
                  FROM DW_FEI.CUSTOMER_OVERRIDE_DIMENSION COD,
                       DW_FEI.CUSTOMER_DIMENSION CUST
                 WHERE     COD.OVERRIDE_TYPE = 'G'
                       AND COD.CUSTOMER_GK = CUST.CUSTOMER_GK
                       AND COD.DELETE_DATE IS NULL
                       AND CUST.JOB_YN = 'N'
                       AND NVL (COD.EXPIRE_DATE, SYSDATE) >= (SYSDATE - 395)) GR_OVR_BASE
                  ON (    LINE_HIST.DISCOUNT_GROUP_NK =
                             (LTRIM (GR_OVR_BASE.DISC_GROUP, '0'))
                      AND LINE_HIST.ACCOUNT_NUMBER =
                             GR_OVR_BASE.BRANCH_NUMBER_NK
                      AND LINE_HIST.MAIN_CUSTOMER_NK = GR_OVR_BASE.CUSTOMER_NK
                      AND NVL (LINE_HIST.CONTRACT_NUMBER, 'DEFAULT_MATCH') =
                             NVL (GR_OVR_BASE.CONTRACT_ID, 'DEFAULT_MATCH'))
               LEFT OUTER JOIN
               (SELECT COD.BASIS,
                       COD.BRANCH_NUMBER_NK,
                       COD.CONTRACT_ID,
                       COD.CUSTOMER_GK,
                       COD.CUSTOMER_NK,
                       COD.DISC_GROUP,
                       COD.INSERT_TIMESTAMP,
                       COD.EXPIRE_DATE,
                       COD.MASTER_PRODUCT,
                       TO_NUMBER (COD.MULTIPLIER) MULTIPLIER,
                       COD.OPERATOR_USED,
                       COD.OVERRIDE_ID_NK,
                       COD.OVERRIDE_TYPE,
                       CASE
                          WHEN COD.OPERATOR_USED <> '$'
                          THEN
                                COD.BASIS
                             || COD.OPERATOR_USED
                             || '0'
                             || COD.MULTIPLIER
                          ELSE
                             TO_CHAR (COD.MULTIPLIER)
                       END
                          FORMULA
                  FROM DW_FEI.CUSTOMER_OVERRIDE_DIMENSION COD,
                       DW_FEI.CUSTOMER_DIMENSION CUST
                 WHERE     COD.OVERRIDE_TYPE = 'P'
                       AND COD.CUSTOMER_GK = CUST.CUSTOMER_GK
                       AND COD.DELETE_DATE IS NULL
                       AND CUST.JOB_YN = 'Y'
                       AND NVL (COD.EXPIRE_DATE, SYSDATE) >= (SYSDATE - 395)) PR_OVR_JOB
                  ON (    LINE_HIST.PRODUCT_NUMBER_NK = PR_OVR_JOB.MASTER_PRODUCT
                      AND LINE_HIST.ACCOUNT_NUMBER =
                             PR_OVR_JOB.BRANCH_NUMBER_NK
                      AND LINE_HIST.CUSTOMER_ACCOUNT_GK =
                             PR_OVR_JOB.CUSTOMER_GK
                      AND NVL (LINE_HIST.CONTRACT_NUMBER, 'DEFAULT_MATCH') =
                             NVL (PR_OVR_JOB.CONTRACT_ID, 'DEFAULT_MATCH'))
               LEFT OUTER JOIN
               (SELECT COD.BASIS,
                       COD.BRANCH_NUMBER_NK,
                       COD.CONTRACT_ID,
                       COD.CUSTOMER_GK,
                       COD.CUSTOMER_NK,
                       COD.DISC_GROUP,
                       COD.INSERT_TIMESTAMP,
                       COD.EXPIRE_DATE,
                       COD.MASTER_PRODUCT,
                       TO_NUMBER (COD.MULTIPLIER) MULTIPLIER,
                       COD.OPERATOR_USED,
                       COD.OVERRIDE_ID_NK,
                       COD.OVERRIDE_TYPE,
                       CASE
                          WHEN COD.OPERATOR_USED <> '$'
                          THEN
                                COD.BASIS
                             || COD.OPERATOR_USED
                             || '0'
                             || COD.MULTIPLIER
                          ELSE
                             TO_CHAR (COD.MULTIPLIER)
                       END
                          FORMULA
                  FROM DW_FEI.CUSTOMER_OVERRIDE_DIMENSION COD,
                       DW_FEI.CUSTOMER_DIMENSION CUST
                 WHERE     COD.OVERRIDE_TYPE = 'P'
                       AND COD.CUSTOMER_GK = CUST.CUSTOMER_GK
                       AND COD.DELETE_DATE IS NULL
                       AND CUST.JOB_YN = 'N'
                       AND NVL (COD.EXPIRE_DATE, SYSDATE) >= (SYSDATE - 395)) PR_OVR_BASE
                  ON (    LINE_HIST.PRODUCT_NUMBER_NK = PR_OVR_BASE.MASTER_PRODUCT
                      AND LINE_HIST.ACCOUNT_NUMBER =
                             PR_OVR_BASE.BRANCH_NUMBER_NK
                      AND LINE_HIST.MAIN_CUSTOMER_NK = PR_OVR_BASE.CUSTOMER_NK
                      AND NVL (LINE_HIST.CONTRACT_NUMBER, 'DEFAULT_MATCH') =
                             NVL (PR_OVR_BASE.CONTRACT_ID, 'DEFAULT_MATCH'))
    GROUP BY LINE_HIST.YEARMONTH,
             LINE_HIST.REGION,
             LINE_HIST.KOB,
             LINE_HIST.ACCOUNT_NUMBER,
             LINE_HIST.WAREHOUSE_NUMBER,
             LINE_HIST.WRITER,
             LINE_HIST.PRICE_CATEGORY,
             LINE_HIST.TYPE_OF_SALE,
             CASE
            WHEN LINE_HIST.ORDER_ENTRY_DATE BETWEEN COALESCE (
                                                       PR_OVR_JOB.INSERT_TIMESTAMP,
                                                       PR_OVR_BASE.INSERT_TIMESTAMP)
                                                AND COALESCE (
                                                       PR_OVR_JOB.EXPIRE_DATE,
                                                       PR_OVR_BASE.EXPIRE_DATE,
                                                       LINE_HIST.ORDER_ENTRY_DATE)
            THEN
               CASE
                  WHEN LINE_HIST.UNIT_NET_PRICE_AMOUNT =
                          COALESCE (PR_OVR_JOB.MULTIPLIER,
                                    PR_OVR_BASE.MULTIPLIER)
                  THEN
                     'OVERRIDE'
                  WHEN LINE_HIST.UNIT_NET_PRICE_AMOUNT =
                          (  TRUNC (
                                COALESCE (PR_OVR_JOB.MULTIPLIER,
                                          PR_OVR_BASE.MULTIPLIER),
                                2)
                           + .01)
                  THEN
                     'OVERRIDE'
                  WHEN LINE_HIST.UNIT_NET_PRICE_AMOUNT =
                          (ROUND (
                              COALESCE (PR_OVR_JOB.MULTIPLIER,
                                        PR_OVR_BASE.MULTIPLIER),
                              2))
                  THEN
                     'OVERRIDE'
                  WHEN LINE_HIST.UNIT_NET_PRICE_AMOUNT =
                          (  TRUNC (
                                COALESCE (PR_OVR_JOB.MULTIPLIER,
                                          PR_OVR_BASE.MULTIPLIER),
                                1)
                           + .1)
                  THEN
                     'OVERRIDE'
                  WHEN LINE_HIST.UNIT_NET_PRICE_AMOUNT =
                            FLOOR (
                               COALESCE (PR_OVR_JOB.MULTIPLIER,
                                         PR_OVR_BASE.MULTIPLIER))
                          + 1
                  THEN
                     'OVERRIDE'
                  WHEN REPLACE (LINE_HIST.PRICE_FORMULA, '0.', '.') =
                          REPLACE (
                             COALESCE (PR_OVR_JOB.FORMULA,
                                       PR_OVR_BASE.FORMULA),
                             '0.',
                             '.')
                  THEN
                     'OVERRIDE'
               END
         END,
         CASE
            WHEN LINE_HIST.ORDER_ENTRY_DATE BETWEEN COALESCE (
                                                       GR_OVR_JOB.INSERT_TIMESTAMP,
                                                       GR_OVR_BASE.INSERT_TIMESTAMP)
                                                AND COALESCE (
                                                       GR_OVR_JOB.EXPIRE_DATE,
                                                       GR_OVR_BASE.EXPIRE_DATE,
                                                       LINE_HIST.ORDER_ENTRY_DATE)
            THEN
               CASE
                  WHEN REPLACE (LINE_HIST.PRICE_FORMULA, '0.', '.') =
                          REPLACE (
                             COALESCE (GR_OVR_JOB.FORMULA,
                                       GR_OVR_BASE.FORMULA),
                             '0.',
                             '.')
                  THEN
                     'OVERRIDE'
                  ELSE
                     NULL
               END
         END)
   UNION
   (SELECT ihf.YEARMONTH,
   ps.division_name REGION,
           ps.ACCOUNT_NUMBER_NK ACCOUNT_NUMBER,
           ps.WAREHOUSE_NUMBER_nk,
           DECODE (ps.DIVISION_NAME,
                   'EAST REGION', 'BLENDED',
                   'WEST REGION', 'BLENDED',
                   'NORTH CENTRAL REGION', 'BLENDED',
                   'SOUTH CENTRAL REGION', 'BLENDED',
                   'INDUSTRIAL REGION', 'INDUSTRIAL',
                   'WATERWORKS REGION', 'WW',
                   'HVAC REGION', 'HVAC',
                    ps.DIVISION_NAME)
              KOB,
           NVL (ihf.WRITER, ihf.OML_ASSOC_INI) WRITER,
           DECODE (ihf.SALE_TYPE,
                   '1', 'Our Truck',
                   '2', 'Counter',
                   '3', 'Direct',
                   '4', 'Counter',
                   '5', 'Credit Memo',
                   '6', 'Showroom',
                   '7', 'Showroom Direct',
                   '8', 'eBusiness')
              TYPE_OF_SALE,
           (0) invoice_lines,
           COUNT (distinct ihf.invoice_number_nk) invoice_cnt,
           SUM (NVL (ihf.AVG_COST_SUBTOTAL_AMOUNT, '0')) avg_cogs,
           SUM (NVL (ihf.COST_SUBTOTAL_AMOUNT, '0')) actual_cogs,
           SUM (ihf.SALES_SUBTOTAL_AMOUNT) ext_sales,
           'Total' AS PRICE_CATEGORY,
           'PROD' AS PRICE_CATEGORY_OVR_PR,
           'GROUP' AS PRICE_CATEGORY_OVR_GR,
           'Total' AS ROLLUP,
           SUM (ihf.TOTAL_SALES_AMOUNT) sls_total,
           SUM (NVL (ihf.SALES_SUBTOTAL_AMOUNT, '0')) sls_subtotal,
           SUM (NVL (ihf.FREIGHT_SALES_AMOUNT, '0')) sls_freight,
           SUM (NVL (ihf.MISC_SALES_AMOUNT, '0')) sls_misc,
           SUM (NVL (ihf.RESTOCKING_SALES_AMOUNT, '0')) sls_restock,
           SUM (NVL (ihf.AVG_COST_SUBTOTAL_AMOUNT, '0')) avg_cost_subtotal,
           SUM (NVL (ihf.FREIGHT_COST_AMOUNT, '0')) avg_cost_freight,
           SUM (NVL (ihf.MISC_COST_AMOUNT, '0')) avg_cost_misc
      FROM DW_FEI.INVOICE_HEADER_FACT ihf,
           DW_FEI.CUSTOMER_DIMENSION cust,
           sales_mart.sales_warehouse_dim ps
     WHERE ihf.CUSTOMER_ACCOUNT_GK = cust.CUSTOMER_GK
           AND TO_CHAR (ihf.WAREHOUSE_NUMBER) = TO_CHAR (ps.WAREHOUSE_NUMBER_NK)
           AND ihf.IC_FLAG = 0
           AND ihf.PO_WAREHOUSE_NUMBER IS NULL
           AND NVL (ihf.CONSIGN_TYPE, 'N') <> 'R'
           AND IHF.YEARMONTH BETWEEN TO_CHAR (
                                                 TRUNC (
                                                    SYSDATE
                                                    - NUMTOYMINTERVAL (
                                                         13,
                                                         'MONTH'),
                                                    'MONTH'),
                                                 'YYYYMM')
                                           AND TO_CHAR (TRUNC (SYSDATE, 'MM') - 1, 'YYYYMM')
           AND DECODE (NVL (cust.ar_gl_number, '9999'),
                                   '1320', 0,
                                   '1360', 0,
                                   '1380', 0,
                                   '1400', 0,
                                   '1401', 0,
                                   '1500', 0,
                                   '4000', 0,
                                   '7100', 0,
                                   '9999', 0,
                                   1) <> 0

    GROUP BY  ihf.YEARMONTH,
    ps.division_name,
             ps.ACCOUNT_NUMBER_NK,
             ps.WAREHOUSE_NUMBER_nk,
             DECODE (ps.DIVISION_NAME,
                   'EAST REGION', 'BLENDED',
                   'WEST REGION', 'BLENDED',
                   'NORTH CENTRAL REGION', 'BLENDED',
                   'SOUTH CENTRAL REGION', 'BLENDED',
                   'INDUSTRIAL REGION', 'INDUSTRIAL',
                   'WATERWORKS REGION', 'WW',
                   'HVAC REGION', 'HVAC',
                    ps.DIVISION_NAME),
             NVL (ihf.WRITER, ihf.OML_ASSOC_INI),
             DECODE (ihf.SALE_TYPE,
                     '1', 'Our Truck',
                     '2', 'Counter',
                     '3', 'Direct',
                     '4', 'Counter',
                     '5', 'Credit Memo',
                     '6', 'Showroom',
                     '7', 'Showroom Direct',
                     '8', 'eBusiness'));
                     GRANT SELECT ON AAE0376.GP_TRACKER_WRITER_YTD TO PUBLIC
         ;
