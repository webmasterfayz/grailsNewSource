package exceldataexport

class MyController {

    def myService

    def index() {
        render "Hello"
    }
    def export() {
        String dd="  SELECT D.DIVISION,\n" +
                "         D.AREA,\n" +
                "         D.PLAZA,\n" +
                "         D.ACCOUNT_NO,\n" +
                "         D.CUSTOMER_NAME,\n" +
                "         D.MOBILE_PHONE,\n" +
                "         D.ID SALE_MST_ID,\n" +
                "         D.INVOICE_NO,\n" +
                "         D.SALE_DATE,\n" +
                "         D.DOWNPAYMENT,\n" +
                "         D.INSTALLMENT_NO,\n" +
                "         D.MRP_OR_CASH,\n" +
                "         D.HMRP_VALUE,\n" +
                "         D.HIRE_VALUE,\n" +
                "         D.COLLECTED_AMT COLLECTION_BEFORE_LPR,\n" +
                "         D.HIRE_VALUE - D.COLLECTED_AMT COLLECTIBLE_AMOUNT,\n" +
                "         D.MRP_OR_CASH - D.COLLECTED_AMT + D.EARNED_AMOUNT\n" +
                "            COLLECTIBLE_AMT_UPTO_MRP,\n" +
                "         D.EARNED_AMOUNT REVENUE_EARNED_FROM_HIRE,\n" +
                "         D.REMARKS\n" +
                "    FROM (SELECT D.TITLE DIVISION,\n" +
                "                 Z.TITLE AREA,\n" +
                "                 P.NAME PLAZA,\n" +
                "                 C.CODE ACCOUNT_NO,\n" +
                "                 C.NAME CUSTOMER_NAME,\n" +
                "                 C.MOBILE_PHONE,\n" +
                "                 M.ID,\n" +
                "                 M.INVOICE_NO,\n" +
                "                 M.CREATED SALE_DATE,\n" +
                "                 M.CASH_RECEIVED DOWNPAYMENT,\n" +
                "                 M.INSTALLMENT_NO,\n" +
                "                 --M.G_TOTAL - NVL (M.REBATE, 0) CASH_PRICE,\n" +
                "                 M.MRP_VALUE MRP_OR_CASH,\n" +
                "                 M.HMRP_VALUE,\n" +
                "                 M.HIRE_VALUE,\n" +
                "                 M.CASH_RECEIVED + NVL (M.COLLECTED_AMT, 0) COLLECTED_AMT,\n" +
                "                 ROUND (\n" +
                "                    (SELECT SUM (L.CREDIT_AMOUNT - L.DEBIT_AMOUNT) AMT\n" +
                "                       FROM ACC_LEDGER L\n" +
                "                      WHERE L.SALE_MST_ID = M.ID AND L.ACC_COA_LEDGER_ID = 45))\n" +
                "                    EARNED_AMOUNT,\n" +
                "                 M.REMARKS\n" +
                "            FROM POS_SALE_MST M,\n" +
                "                 POS_CUSTOMERS C,\n" +
                "                 POS_PLAZAS P,\n" +
                "                 HR_DEVELOPMENT D,\n" +
                "                 HR_ZONE Z\n" +
                "           WHERE     1 = 1\n" +
                "                 AND M.CUSTOMER_ID = C.ID\n" +
                "                 AND C.PLAZA_ID = P.ID\n" +
                "                 AND P.DEVELOPMENT_ID = D.ID\n" +
                "                 AND P.ZONE_ID = Z.ID\n" +
                "                 AND P.P_STATUS = 'Active'\n" +
                "                 AND M.S_TYPE = 'Hire'\n" +
                "                 AND C.C_TYPE = 'Hire'\n" +
                "                 AND NVL (M.IS_RETURN, 0) = 0\n" +
                "                 AND NVL (M.IS_EMPLOYEE, 0) = 0\n" +
                "                 AND NVL (M.IS_LPR, 0) = 0) D\n" +
                "ORDER BY D.DIVISION, D.AREA, D.PLAZA"
        myService.generateExcelFile(dd, "mytable2.xlsx")
        render "Exported to Excel successfully"
    }



}