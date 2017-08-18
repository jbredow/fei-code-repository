/*price category by writer
Leigh North
Formatted 10/25/2011
*/
--DROP TABLE AAD9606.GP_TRACKER_12MO;
DROP TABLE AAD9606.GP_TRACKER_WRITER;

CREATE TABLE AAD9606.GP_TRACKER_WRITER
--DROP TABLE AAD9606.GP_TRACKER_TEST;
--CREATE TABLE AAD9606.GP_TRACKER_TEST
AS
   (SELECT SUBSTR (ihf.YEARMONTH, 0, 4) YYYY,
           ihf.YEARMONTH,
           ihf.WRITER,
           ps.REGION,
           ps.ACCOUNT_NUMBER_NK ACCOUNT_NUMBER,
           --cust.MAIN_CUSTOMER_NK,
           --cust.CUSTOMER_NAME,
           --MIN (ihf.YEARMONTH) BEGIN_MM,
           --MAX (ihf.YEARMONTH) END_MM,
           DECODE (ps.KIND_OF_BUSINESS, 'PLB', 'PLBG', ps.KIND_OF_BUSINESS)
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
              WHEN ihf.order_code = 'SP'
              THEN
                 'SPECIALS'
              WHEN ilf.price_code = 'Q'
              THEN
                 'QUOTE'
              WHEN REGEXP_LIKE (ilf.price_code, '[0-9]?[0-9]?[0-9]')
              THEN
                 'MATRIX'
              --WHEN NVL (ihf.split_ticket_flag, '0') = '1'
              --THEN
              --   'MATRIX'
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
                    WHEN ROUND (NVL (ilf.MATRIX_PRICE, ilf.MATRIX), 2) =
                            ilf.UNIT_NET_PRICE_AMOUNT
                    THEN
                       'MATRIX'
                    WHEN ihf.CONTRACT_NUMBER IS NOT NULL
                    THEN
                       'OVERRIDE'
                    ELSE
                       'MANUAL'
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
           SCORECARD1.PS_HIERARCHY ps
     WHERE ihf.CUSTOMER_ACCOUNT_GK = cust.CUSTOMER_GK
           AND (ilf.INVOICE_NUMBER_GK = ihf.INVOICE_NUMBER_GK)
           AND TO_CHAR (ihf.WAREHOUSE_NUMBER) =
                  TO_CHAR (ps.WAREHOUSE_NUMBER_NK)
           --AND ihf.ACCOUNT_NUMBER = 39
           AND ihf.IC_FLAG = 0
           AND ihf.PO_WAREHOUSE_NUMBER IS NULL
           AND NVL (ihf.CONSIGN_TYPE, 'N') <> 'R'
           --AND ILF.YEARMONTH = 201209
           --AND IHF.YEARMONTH = 201209
           AND ILF.YEARMONTH =
                  TO_NUMBER (TO_CHAR (TRUNC (SYSDATE, 'MM') - 1, 'YYYYMM'))
           AND IHF.YEARMONTH =
                  TO_NUMBER (TO_CHAR (TRUNC (SYSDATE, 'MM') - 1, 'YYYYMM'))
           --AND ihf.ACCOUNT_NUMBER = 4
           AND (cust.ar_gl_number NOT IN
                   ('1320',
                    '1360',
                    '1380',
                    '1400',
                    '1401',
                    '1500',
                    '4000',
                    '7100')
                AND cust.ar_gl_number IS NOT NULL)
    GROUP BY SUBSTR (ihf.YEARMONTH, 0, 4),
             ihf.YEARMONTH,
             ihf.WRITER,
             ps.REGION,
             DECODE (ihf.SALE_TYPE,
                     '1', 'Our Truck',
                     '2', 'Counter',
                     '3', 'Direct',
                     '4', 'Counter',
                     '5', 'Credit Memo',
                     '6', 'Showroom',
                     '7', 'Showroom Direct',
                     '8', 'eBusiness'),
             ps.ACCOUNT_NUMBER_NK,
             --cust.MAIN_CUSTOMER_NK,
             --cust.CUSTOMER_NAME,
             DECODE (ps.KIND_OF_BUSINESS, 'PLB', 'PLBG', ps.KIND_OF_BUSINESS),
             CASE
                WHEN ihf.order_code = 'IC'
                THEN
                   'CREDITS'
                WHEN ihf.order_code = 'SP'
                THEN
                   'SPECIALS'
                WHEN ilf.price_code = 'Q'
                THEN
                   'QUOTE'
                WHEN REGEXP_LIKE (ilf.price_code, '[0-9]?[0-9]?[0-9]')
                THEN
                   'MATRIX'
                --WHEN NVL (ihf.split_ticket_flag, '0') = '1'
                --THEN
                --   'MATRIX'
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
                      WHEN ROUND (NVL (ilf.MATRIX_PRICE, ilf.MATRIX), 2) =
                              ilf.UNIT_NET_PRICE_AMOUNT
                      THEN
                         'MATRIX'
                      WHEN ihf.CONTRACT_NUMBER IS NOT NULL
                      THEN
                         'OVERRIDE'
                      ELSE
                         'MANUAL'
                   END
                ELSE
                   'MANUAL'
             END)
   UNION
   (SELECT SUBSTR (ihf.YEARMONTH, 0, 4) YYYY,
           ihf.YEARMONTH,
           ihf.WRITER,
           ps.REGION,
           ps.ACCOUNT_NUMBER_NK ACCOUNT_NUMBER,
           --cust.MAIN_CUSTOMER_NK,
           --cust.CUSTOMER_NAME,
           --MIN (ihf.YEARMONTH) BEGIN_MM,
           --MAX (ihf.YEARMONTH) END_MM,
           DECODE (ps.KIND_OF_BUSINESS,
                   'PLB', 'PLBG',
                   'HFMB', 'PLBG',
                   ps.KIND_OF_BUSINESS)
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
           (0) invoice_lines,
           SUM (NVL (ihf.AVG_COST_SUBTOTAL_AMOUNT, '0')) avg_cogs,
           SUM (NVL (ihf.COST_SUBTOTAL_AMOUNT, '0')) actual_cogs,
           SUM (ihf.SALES_SUBTOTAL_AMOUNT) ext_sales,
           'Total' AS PRICE_CATEGORY,
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
           SCORECARD1.PS_HIERARCHY ps
     WHERE ihf.CUSTOMER_ACCOUNT_GK = cust.CUSTOMER_GK
           --AND ihf.ACCOUNT_NUMBER = 39
           --AND (ilf.INVOICE_NUMBER_GK = ihf.INVOICE_NUMBER_GK)
           AND TO_CHAR (ihf.WAREHOUSE_NUMBER) =
                  TO_CHAR (ps.WAREHOUSE_NUMBER_NK)
           AND ihf.IC_FLAG = 0
           AND ihf.PO_WAREHOUSE_NUMBER IS NULL
           AND NVL (ihf.CONSIGN_TYPE, 'N') <> 'R'
           AND IHF.YEARMONTH =
                  TO_NUMBER (TO_CHAR (TRUNC (SYSDATE, 'MM') - 1, 'YYYYMM'))
           AND (cust.ar_gl_number NOT IN
                   ('1320',
                    '1360',
                    '1380',
                    '1400',
                    '1401',
                    '1500',
                    '4000',
                    '7100')
                AND cust.ar_gl_number IS NOT NULL)
    GROUP BY SUBSTR (ihf.YEARMONTH, 0, 4),
             ihf.YEARMONTH,
             ihf.WRITER,
             ps.REGION,
             ps.ACCOUNT_NUMBER_NK,
             --cust.MAIN_CUSTOMER_NK,
             --cust.CUSTOMER_NAME,
             DECODE (ps.KIND_OF_BUSINESS,
                     'PLB', 'PLBG',
                     'HFMB', 'PLBG',
                     ps.KIND_OF_BUSINESS),
             DECODE (ihf.SALE_TYPE,
                     '1', 'Our Truck',
                     '2', 'Counter',
                     '3', 'Direct',
                     '4', 'Counter',
                     '5', 'Credit Memo',
                     '6', 'Showroom',
                     '7', 'Showroom Direct',
                     '8', 'eBusiness'))
