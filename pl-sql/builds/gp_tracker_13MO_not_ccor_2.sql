TRUNCATE TABLE AAE0376.GP_TRACKER_13MO_NOT_CCOR;
DROP TABLE AAE0376.GP_TRACKER_13MO_NOT_CCOR;

CREATE TABLE AAE0376.GP_TRACKER_13MO_NOT_CCOR

AS
(SELECT SUBSTR (ihf.YEARMONTH, 0, 4) YYYY,
          ihf.YEARMONTH,
           DECODE (
          ihf.YEARMONTH,
          TO_CHAR (TRUNC (SYSDATE - NUMTOYMINTERVAL (12, 'MONTH'), 'MONTH'),
                   'YYYYMM'), 'ROLL_Q1',
          TO_CHAR (TRUNC (SYSDATE - NUMTOYMINTERVAL (11, 'MONTH'), 'MONTH'),
                   'YYYYMM'), 'ROLL_Q1',
          TO_CHAR (TRUNC (SYSDATE - NUMTOYMINTERVAL (10, 'MONTH'), 'MONTH'),
                   'YYYYMM'), 'ROLL_Q1',
          TO_CHAR (TRUNC (SYSDATE - NUMTOYMINTERVAL (9, 'MONTH'), 'MONTH'),
                   'YYYYMM'), 'ROLL_Q2',
          TO_CHAR (TRUNC (SYSDATE - NUMTOYMINTERVAL (8, 'MONTH'), 'MONTH'),
                   'YYYYMM'), 'ROLL_Q2',
          TO_CHAR (TRUNC (SYSDATE - NUMTOYMINTERVAL (7, 'MONTH'), 'MONTH'),
                   'YYYYMM'), 'ROLL_Q2',
          TO_CHAR (TRUNC (SYSDATE - NUMTOYMINTERVAL (6, 'MONTH'), 'MONTH'),
                   'YYYYMM'), 'ROLL_Q3',
          TO_CHAR (TRUNC (SYSDATE - NUMTOYMINTERVAL (5, 'MONTH'), 'MONTH'),
                   'YYYYMM'), 'ROLL_Q3',
          TO_CHAR (TRUNC (SYSDATE - NUMTOYMINTERVAL (4, 'MONTH'), 'MONTH'),
                   'YYYYMM'), 'ROLL_Q3',
          TO_CHAR (TRUNC (SYSDATE - NUMTOYMINTERVAL (3, 'MONTH'), 'MONTH'),
                   'YYYYMM'), 'ROLL_Q4',
          TO_CHAR (TRUNC (SYSDATE - NUMTOYMINTERVAL (2, 'MONTH'), 'MONTH'),
                   'YYYYMM'), 'ROLL_Q4',
          TO_CHAR (TRUNC (SYSDATE - NUMTOYMINTERVAL (1, 'MONTH'), 'MONTH'),
                   'YYYYMM'), 'ROLL_Q4')
          ROLLING_QTR,
           ps.DIVISION_NAME REGION,
           ps.ACCOUNT_NUMBER_NK ACCOUNT_NUMBER,
           PS.WAREHOUSE_NUMBER_NK WAREHOUSE_NUMBER,
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
           SUM (CASE WHEN ilf.SHIPPED_QTY > 0 THEN 1 ELSE 0 END)
              invoice_lines,
           SUM (ilf.EXT_AVG_COGS_AMOUNT) avg_cogs,
           SUM (ilf.EXT_ACTUAL_COGS_AMOUNT) actual_cogs,
           SUM (ilf.EXT_SALES_AMOUNT) ext_sales,
           CASE
              WHEN ihf.order_code = 'IC'
              THEN
                 'CREDITS'
              WHEN ilf.special_product_gk IS NOT NULL
              THEN
                 'SPECIALS'
              WHEN ilf.price_code = 'Q'
              THEN
                 CASE
                                   WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                    ilf.MATRIX_PRICE
                            THEN
                               'MATRIX'
                            WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                    ilf.MATRIX
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
                                           (TRUNC (ilf.MATRIX_PRICE, 2)+.01) 
                                   THEN
                                      'MATRIX'
                                   WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                           (TRUNC (ilf.MATRIX, 2)+.01)
                                   THEN
                                      'MATRIX_BID'
                                   WHEN  ilf.UNIT_NET_PRICE_AMOUNT =
                                           (TRUNC (ilf.MATRIX_PRICE, 1)+.1)  
                                   THEN
                                      'MATRIX'
                                   WHEN  ilf.UNIT_NET_PRICE_AMOUNT =
                                           (TRUNC (ilf.MATRIX, 1)+.1)
                                   THEN
                                      'MATRIX_BID'
                                   WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                           FLOOR (ilf.MATRIX_PRICE)+1
                                   THEN
                                      'MATRIX'
                                   WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                           FLOOR (ilf.MATRIX)+1
                                   THEN
                                      'MATRIX_BID'
                                      ELSE 'QUOTE'
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
              WHEN ilf.price_code IN ('GI', 'GPC', 'HPF', 'HPN', 'NC')
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
                                   WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                    ilf.MATRIX_PRICE
                            THEN
                               'MATRIX'
                            WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                    ilf.MATRIX
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
                                           (TRUNC (ilf.MATRIX_PRICE, 2)+.01) 
                                   THEN
                                      'MATRIX'
                                   WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                           (TRUNC (ilf.MATRIX, 2)+.01)
                                   THEN
                                      'MATRIX_BID'
                                   WHEN  ilf.UNIT_NET_PRICE_AMOUNT =
                                           (TRUNC (ilf.MATRIX_PRICE, 1)+.1)  
                                   THEN
                                      'MATRIX'
                                   WHEN  ilf.UNIT_NET_PRICE_AMOUNT =
                                           (TRUNC (ilf.MATRIX, 1)+.1)
                                   THEN
                                      'MATRIX_BID'
                                   WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                           FLOOR (ilf.MATRIX_PRICE)+1
                                   THEN
                                      'MATRIX'
                                   WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                           FLOOR (ilf.MATRIX)+1
                                   THEN
                                      'MATRIX_BID'
                                   WHEN ILF.MATRIX_PRICE IS NULL AND ILF.PRICE_FORMULA LIKE 'L-0.%'  
            THEN 
                'NDP'

                                     ELSE 'MANUAL'
                                END
              ELSE
                 'MANUAL'
           END
              AS PRICE_CATEGORY,
           'Subtotal' AS ROLLUP,
           (0) sls_total,
           (0) sls_subtotal,
           (0) sls_freight,
           (0) sls_misc,
           (0) sls_restock,
           (0) avg_cost_subtotal,
           (0) avg_cost_freight,
           (0) avg_cost_misc
      FROM DW_FEI.INVOICE_LINE_FACT ilf,
           DW_FEI.INVOICE_HEADER_FACT ihf,
           DW_FEI.CUSTOMER_DIMENSION cust,
           SALES_MART.SALES_WAREHOUSE_DIM ps
     WHERE ihf.CUSTOMER_ACCOUNT_GK = cust.CUSTOMER_GK
           AND (ilf.INVOICE_NUMBER_GK = ihf.INVOICE_NUMBER_GK)
           AND TO_CHAR (ihf.WAREHOUSE_NUMBER) = TO_CHAR (ps.WAREHOUSE_NUMBER_NK)
           AND ihf.IC_FLAG = 0
           AND ihf.PO_WAREHOUSE_NUMBER IS NULL
           AND NVL (ihf.CONSIGN_TYPE, 'N') <> 'R'
           AND DECODE(NVL (ilf.PRICE_CODE,'N/A'), 'R', 0, 'Q', 0,'N/A',0,1) <> 0
           AND ILF.YEARMONTH BETWEEN TO_CHAR (
                                        TRUNC (
                                           SYSDATE
                                           - NUMTOYMINTERVAL (12, 'MONTH'),
                                           'MONTH'),
                                        'YYYYMM')
                                        AND TO_CHAR (TRUNC (SYSDATE, 'MM') - 1,
                                              'YYYYMM')
           AND IHF.YEARMONTH BETWEEN TO_CHAR (
                                        TRUNC (
                                           SYSDATE
                                           - NUMTOYMINTERVAL (12, 'MONTH'),
                                           'MONTH'),
                                        'YYYYMM')
                                        AND TO_CHAR (TRUNC (SYSDATE, 'MM') - 1,
                                              'YYYYMM')
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


    GROUP BY SUBSTR (ihf.YEARMONTH, 0, 4),
    ihf.YEARMONTH,
             DECODE (
          ihf.YEARMONTH,
          TO_CHAR (TRUNC (SYSDATE - NUMTOYMINTERVAL (12, 'MONTH'), 'MONTH'),
                   'YYYYMM'), 'ROLL_Q1',
          TO_CHAR (TRUNC (SYSDATE - NUMTOYMINTERVAL (11, 'MONTH'), 'MONTH'),
                   'YYYYMM'), 'ROLL_Q1',
          TO_CHAR (TRUNC (SYSDATE - NUMTOYMINTERVAL (10, 'MONTH'), 'MONTH'),
                   'YYYYMM'), 'ROLL_Q1',
          TO_CHAR (TRUNC (SYSDATE - NUMTOYMINTERVAL (9, 'MONTH'), 'MONTH'),
                   'YYYYMM'), 'ROLL_Q2',
          TO_CHAR (TRUNC (SYSDATE - NUMTOYMINTERVAL (8, 'MONTH'), 'MONTH'),
                   'YYYYMM'), 'ROLL_Q2',
          TO_CHAR (TRUNC (SYSDATE - NUMTOYMINTERVAL (7, 'MONTH'), 'MONTH'),
                   'YYYYMM'), 'ROLL_Q2',
          TO_CHAR (TRUNC (SYSDATE - NUMTOYMINTERVAL (6, 'MONTH'), 'MONTH'),
                   'YYYYMM'), 'ROLL_Q3',
          TO_CHAR (TRUNC (SYSDATE - NUMTOYMINTERVAL (5, 'MONTH'), 'MONTH'),
                   'YYYYMM'), 'ROLL_Q3',
          TO_CHAR (TRUNC (SYSDATE - NUMTOYMINTERVAL (4, 'MONTH'), 'MONTH'),
                   'YYYYMM'), 'ROLL_Q3',
          TO_CHAR (TRUNC (SYSDATE - NUMTOYMINTERVAL (3, 'MONTH'), 'MONTH'),
                   'YYYYMM'), 'ROLL_Q4',
          TO_CHAR (TRUNC (SYSDATE - NUMTOYMINTERVAL (2, 'MONTH'), 'MONTH'),
                   'YYYYMM'), 'ROLL_Q4',
          TO_CHAR (TRUNC (SYSDATE - NUMTOYMINTERVAL (1, 'MONTH'), 'MONTH'),
                   'YYYYMM'), 'ROLL_Q4'),
             ps.DIVISION_NAME,
             DECODE (ihf.SALE_TYPE,
                     '1', 'Our Truck',
                     '2', 'Counter',
                     '3', 'Direct',
                     '4', 'Counter',
                     '5', 'Credit Memo',
                     '6', 'Showroom',
                     '7', 'Showroom Direct',
                     '8', 'eBusiness'),
                     ilf.price_code,
             ps.ACCOUNT_NUMBER_NK,
             PS.WAREHOUSE_NUMBER_NK,
             DECODE (ps.DIVISION_NAME,
                   'EAST REGION', 'BLENDED',
                   'WEST REGION', 'BLENDED',
                   'NORTH CENTRAL REGION', 'BLENDED',
                   'SOUTH CENTRAL REGION', 'BLENDED',
                   'INDUSTRIAL REGION', 'INDUSTRIAL',
                   'WATERWORKS REGION', 'WW',
                   'HVAC REGION', 'HVAC',
                    ps.DIVISION_NAME),
             CASE
              WHEN ihf.order_code = 'IC'
              THEN
                 'CREDITS'
              WHEN ilf.special_product_gk IS NOT NULL
              THEN
                 'SPECIALS'
              WHEN ilf.price_code = 'Q'
              THEN
                 CASE
                                   WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                    ilf.MATRIX_PRICE
                            THEN
                               'MATRIX'
                            WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                    ilf.MATRIX
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
                                           (TRUNC (ilf.MATRIX_PRICE, 2)+.01) 
                                   THEN
                                      'MATRIX'
                                   WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                           (TRUNC (ilf.MATRIX, 2)+.01)
                                   THEN
                                      'MATRIX_BID'
                                   WHEN  ilf.UNIT_NET_PRICE_AMOUNT =
                                           (TRUNC (ilf.MATRIX_PRICE, 1)+.1)  
                                   THEN
                                      'MATRIX'
                                   WHEN  ilf.UNIT_NET_PRICE_AMOUNT =
                                           (TRUNC (ilf.MATRIX, 1)+.1)
                                   THEN
                                      'MATRIX_BID'
                                   WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                           FLOOR (ilf.MATRIX_PRICE)+1
                                   THEN
                                      'MATRIX'
                                   WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                           FLOOR (ilf.MATRIX)+1
                                   THEN
                                      'MATRIX_BID'
                                      ELSE 'QUOTE'
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
              WHEN ilf.price_code IN ('GI', 'GPC', 'HPF', 'HPN', 'NC')
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
                                   WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                    ilf.MATRIX_PRICE
                            THEN
                               'MATRIX'
                            WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                    ilf.MATRIX
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
                                           (TRUNC (ilf.MATRIX_PRICE, 2)+.01) 
                                   THEN
                                      'MATRIX'
                                   WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                           (TRUNC (ilf.MATRIX, 2)+.01)
                                   THEN
                                      'MATRIX_BID'
                                   WHEN  ilf.UNIT_NET_PRICE_AMOUNT =
                                           (TRUNC (ilf.MATRIX_PRICE, 1)+.1)  
                                   THEN
                                      'MATRIX'
                                   WHEN  ilf.UNIT_NET_PRICE_AMOUNT =
                                           (TRUNC (ilf.MATRIX, 1)+.1)
                                   THEN
                                      'MATRIX_BID'
                                   WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                           FLOOR (ilf.MATRIX_PRICE)+1
                                   THEN
                                      'MATRIX'
                                   WHEN ilf.UNIT_NET_PRICE_AMOUNT =
                                           FLOOR (ilf.MATRIX)+1
                                   THEN
                                      'MATRIX_BID'
                                   WHEN ILF.MATRIX_PRICE IS NULL AND ILF.PRICE_FORMULA LIKE 'L-0.%'  
            THEN 
                'NDP'

                                      ELSE 'MANUAL'
                                END
              ELSE
                 'MANUAL'
           END)
       ;
