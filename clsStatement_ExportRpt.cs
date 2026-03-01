using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Data;
using Oracle.DataAccess.Client;
using System.Xml;
using System.Collections;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using PdfSharp;
using PdfSharp.Pdf.IO;

public class clsStatement_ExportRpt : clsBasStatement
    {
    private DateTime vCurDate;
    private string strOutputPath, strFileName;
    private string StrStatLable = string.Empty, strWhereCond = string.Empty, strWhereCondd = string.Empty;
    private ArrayList aryLstFiles = new ArrayList();
    private string selFormula = string.Empty;
    private string strCloseBalance = string.Empty;
    private string fileSummaryName = string.Empty;
    private frmStatementFile frmMain;
    protected string prepaidCond = string.Empty;
    private string strProductCond = string.Empty;
    private clsExportReport expReportdbca = new clsExportReport();

    public clsStatement_ExportRpt()
        {
        }

    public bool mantainBank(int pBankCode)
        {
        clsMaintainData maintainData = new clsMaintainData();
        //maintainData.matchCardBranch4Account(pBankCode);
        maintainData.makeBranchAsMainCard(pBankCode);
        maintainData.fixArbicAddress(pBankCode);

        return true;
        }

    //the function combine the pdf with the VO
    public void CombinePDF(string StatementPath, string BovPath, string DestinationOutPut)
        {
        try
            {
            PdfDocument Original = PdfReader.Open(StatementPath, PdfDocumentOpenMode.Import);
            PdfDocument Bov = PdfReader.Open(BovPath, PdfDocumentOpenMode.Import);

            PdfDocument outputDocument = new PdfDocument();



            outputDocument.PageLayout = PdfPageLayout.SinglePage;


            XFont font = new XFont("Verdana", 10, XFontStyle.Bold);

            XStringFormat format = new XStringFormat();

            format.Alignment = XStringAlignment.Center;

            format.LineAlignment = XLineAlignment.Far;

            int count = Math.Max(Bov.PageCount, Original.PageCount);
            PdfPage page2 = Bov.PageCount > 0 ? Bov.Pages[0] : new PdfPage();
            page2.Size = PageSize.A4;

            for (int idx = 0; idx < count; idx++)
                {
                PdfPage page1 = Original.PageCount > idx ? Original.Pages[idx] : new PdfPage();
                page1.Size = PageSize.A4;
                outputDocument.AddPage(page1);
                outputDocument.AddPage(page2);
                }
            string filename = DestinationOutPut;
            outputDocument.Save(filename);

            }
        catch (Exception ex)
            {
            clsBasErrors.catchError(ex);
            }
        }

    public bool export(string pStrFileName, string pBankName
      , int pBankCode, string pStrFile, string pReportName, ExportFormatType pCurrType, DateTime pDate, string pStmntType, bool pAppendData)
        {
        bool rtrnVal = true;
        string strCloseBalanceFrml = string.Empty, strCloseBalanceSql = string.Empty;
        clsExportReport expReport = new clsExportReport();
        expReport.bankCode = pBankCode;
        if (strCloseBalance == "ALL")
            {
            strCloseBalanceFrml = "";
            strCloseBalanceSql = "";
            }
        else if (strCloseBalance == string.Empty)
            {
            strCloseBalanceFrml = "";//" and {@f_closingbalance} <> 0 ";
            strCloseBalanceSql = "";//" and closingbalance <> 0";
            }
        else
            {
            strCloseBalanceFrml = " and {@f_closingbalance} " + strCloseBalance;
            strCloseBalanceSql = " and closingbalance " + strCloseBalance;
            }
        strCloseBalanceSql += strCloseBalanceSql;

        if (selFormula == string.Empty)
            //expReport.selectionFormula = "if not isnull({@f_BranchTrans}) then {@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) else {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 ";
            expReport.selectionFormula = "((IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y') or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + ") and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y')";

        else
            expReport.selectionFormula = selFormula;

        try
            {
            string exportResult;

            pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);
            vCurDate = pDate;
            strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
            clsBasFile.createDirectory(pStrFileName + vCurDate.ToString("yyyyMM") + pBankName);
            strFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM"); // + ".pdf"

            curBranchVal = pBankCode;
            MainTableCond = strWhereCond;
            frmMain.BeginInvoke(frmMain.setStatusDelegate, new object[] { 0, "Data Retrieved" });//\r\n\r\n
            expReport.StatLable = StrStatLable;

            if (pBankName == "ICBG_Prepaid")
                {
                //expReport.selectionFormula = "if not isnull({@f_BranchTrans}) then {@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) and {@f_cardproduct} in " + prepaidCond.Replace('(', '[').Replace(')', ']') + " else {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {@f_cardproduct} in " + prepaidCond.Replace('(', '[').Replace(')', ']') + ""; //and {@f_closingbalance} <> 0 
                expReport.selectionFormula = "((IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {@f_cardproduct} in " + prepaidCond.Replace('(', '[').Replace(')', ']') + ") and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y') or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_cardproduct} in " + prepaidCond.Replace('(', '[').Replace(')', ']') + " and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y')";
                expReport.ExportSetup(clsBasFile.getPathWithoutFile(strFileName), clsBasFile.getFileFromPath(strFileName), pReportName, DSstatement);
                expReport.ExportSelection(pCurrType);
                exportResult = expReport.ExportCompletion();
                MainTableCond += " m.cardproduct in " + PrepaidCondition;
                FillStatementDataSet(pBankCode, "");
                }
            else if (pBankCode == 74)
                {
                // merge mark-up fee with original transaction
                clsMaintainData maintainData = new clsMaintainData();
                maintainData.mergeMarkUpFees(pBankCode);
                maintainData = null;
                MainTableCond = " m.contracttype in " + strProductCond + "";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + ")";
                //expReport.selectionFormula = "if not isnull({@f_BranchTrans}) then {@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} in " + strProductCond.Replace('(', '[').Replace(')', ']') + " else {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} in " + strProductCond.Replace('(', '[').Replace(')', ']') + ""; //and {@f_closingbalance} <> 0 
                //expReport.selectionFormula = "((IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} in " + strProductCond.Replace('(', '[').Replace(')', ']') + " and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y')) or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} in " + strProductCond.Replace('(', '[').Replace(')', ']') + " and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y')";
                //expReport.selectionFormula = "((IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} like " + strProductCond.Replace('%', '*') + " and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y')) or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} like " + strProductCond.Replace('%', '*') + " and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y')";


                expReport.selectionFormula = "((IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} like " + strProductCond.Replace('%', '*') + " and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y')) or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} like " + strProductCond.Replace('%', '*') + " and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y')  and not({TSTATEMENTDETAILTABLE.TRANDESCRIPTION} like \"*MARK-UP*\" )";
                expReport.ExportSetup(clsBasFile.getPathWithoutFile(strFileName), clsBasFile.getFileFromPath(strFileName), pReportName, DSstatement);
                expReport.ExportSelection(pCurrType);
                exportResult = expReport.ExportCompletion();
                FillStatementDataSetWithRemovingMarkupFee(pBankCode); //FillStatementDataSet(pBankCode, ""); MAMR Update to exclude markup fee after combined
                getClientEmail(pBankCode);
                DSstatement.Merge(DSemails.Tables[0]);
                }
            else if (pBankName == "NCB_Credit")
                {
                expReport.ExportSetup(clsBasFile.getPathWithoutFile(strFileName), clsBasFile.getFileFromPath(strFileName), pReportName, DSstatement);
                expReport.ExportSelection(pCurrType);
                exportResult = expReport.ExportCompletion();
                FillStatementDataSet(pBankCode, "");
                getClientEmail(pBankCode);
                DSstatement.Merge(DSemails.Tables[0]);
                }
            else if (pBankName == "NCB_Prepaid")
                {
                MainTableCond = " m.cardproduct in " + prepaidCond + "";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.cardproduct in " + prepaidCond + ")";
                expReport.selectionFormula = "((IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.CARDPRODUCT} in " + prepaidCond.Replace('(', '[').Replace(')', ']') + ")and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y') or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.CARDPRODUCT} in " + prepaidCond.Replace('(', '[').Replace(')', ']') + " and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y')";
                expReport.ExportSetup(clsBasFile.getPathWithoutFile(strFileName), clsBasFile.getFileFromPath(strFileName), pReportName, DSstatement);
                expReport.ExportSelection(pCurrType);
                exportResult = expReport.ExportCompletion();
                FillStatementDataSet(pBankCode, "");
                getClientEmail(pBankCode);
                DSstatement.Merge(DSemails.Tables[0]);
                }
            else if (pBankCode == 24)
                {
                expReport.ExportSetup(clsBasFile.getPathWithoutFile(strFileName), clsBasFile.getFileFromPath(strFileName), pReportName, DSstatement);
                expReport.ExportSelection(pCurrType);
                exportResult = expReport.ExportCompletion();
                FillStatementDataSet(pBankCode, "vip");
                }
            else if (pBankCode == 45)
                {
                expReport.ExportSetup(clsBasFile.getPathWithoutFile(strFileName), clsBasFile.getFileFromPath(strFileName), pReportName, DSstatement);
                expReport.ExportSelection(pCurrType);
                exportResult = expReport.ExportCompletion();
                FillStatementDataSet(pBankCode, "vip");
                }
            else if (pBankCode == 67)
                {
                expReport.ExportSetup(clsBasFile.getPathWithoutFile(strFileName), clsBasFile.getFileFromPath(strFileName), pReportName, DSstatement);
                expReport.ExportSelection(pCurrType);
                exportResult = expReport.ExportCompletion();
                FillStatementDataSet(pBankCode, "");
                }
            else if (pBankCode == 1) // nsgb (QNB)
            {
                MainTableCond = " m.cardproduct in " + strProductCond + "";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.cardproduct in " + strProductCond + ")";
                expReport.selectionFormula = "((IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.CARDPRODUCT} in " + strProductCond.Replace('(', '[').Replace(')', ']') + ")and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y') or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.CARDPRODUCT} in " + strProductCond.Replace('(', '[').Replace(')', ']') + " and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y')";
                expReport.ExportSetup(clsBasFile.getPathWithoutFile(strFileName), clsBasFile.getFileFromPath(strFileName), pReportName, DSstatement);
                expReport.ExportSelection(pCurrType);
                exportResult = expReport.ExportCompletion();
                FillStatementDataSet(pBankCode, "");
                }
            else if (pBankName == "GTBK_Prepaid")
                {
                MainTableCond = " m.cardproduct in " + prepaidCond + "";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.cardproduct in " + prepaidCond + ")";
                expReport.selectionFormula = "((IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.CARDPRODUCT} in " + prepaidCond.Replace('(', '[').Replace(')', ']') + ")and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y') or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.CARDPRODUCT} in " + prepaidCond.Replace('(', '[').Replace(')', ']') + " and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y')";
                expReport.ExportSetup(clsBasFile.getPathWithoutFile(strFileName), clsBasFile.getFileFromPath(strFileName), pReportName, DSstatement);
                expReport.ExportSelection(pCurrType);
                exportResult = expReport.ExportCompletion();
                FillStatementDataSet(pBankCode, "");
                }
            else if (pBankName == "I&M_Prepaid")
                {
                MainTableCond = " m.cardproduct in " + prepaidCond + "";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.cardproduct in " + prepaidCond + ")";
                expReport.selectionFormula = "((IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.CARDPRODUCT} in " + prepaidCond.Replace('(', '[').Replace(')', ']') + ")and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y') or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.CARDPRODUCT} in " + prepaidCond.Replace('(', '[').Replace(')', ']') + " and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y')";
                expReport.ExportSetup(clsBasFile.getPathWithoutFile(strFileName), clsBasFile.getFileFromPath(strFileName), pReportName, DSstatement);
                expReport.ExportSelection(pCurrType);
                exportResult = expReport.ExportCompletion();
                FillStatementDataSet(pBankCode, "");
                }
            else if (pBankCode == 128)//EGB Egyptian Gulf Bank of Egyp
            {
                //edit 7/12/2021 to select some 
                expReport.selectionFormula = "not isNull ({@F_accountno}) and not isNull({@F_cardno}) and ((IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = "+ pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y') or (not IsNull({@f_POSTINGDATE})) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = "+ pBankCode+ ") and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y' and IsNull({TSTATEMENTMASTERTABLE.CREDITCONTRACTS})";
                //expReport.selectionFormula = "not isNull ({@F_accountno}) and not isNull({@F_cardno}) and ((IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = "+ pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y') or (not IsNull({@f_POSTINGDATE})) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = "+ pBankCode+") and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y' and (({TSTATEMENTMASTERTABLE.CARDNO} = '4644800374601505') or ({TSTATEMENTMASTERTABLE.CARDNO} = '4644800200079231' )) and IsNull({TSTATEMENTMASTERTABLE.CREDITCONTRACTS})";
                
                //expReport.selectionFormula = "not isNull ({@F_accountno}) and not isNull({@F_cardno}) and ((IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y') or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + ") and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y' and IsNull({TSTATEMENTMASTERTABLE.CREDITCONTRACTS}))";
                expReport.ExportSetup(clsBasFile.getPathWithoutFile(strFileName), clsBasFile.getFileFromPath(strFileName), pReportName, DSstatement);
                expReport.ExportSelection(pCurrType);
                exportResult = expReport.ExportCompletion();
                FillStatementDataSet(pBankCode, "");
                CombinePDF(strFileName + ".pdf", @"D:\pC#\ProjData\Statement\EGB\BOV.PDF", strFileName + ".pdf");
                }
            //else if (pBankCode == 15)
            //    {
            //    decimal closebal = 0;

             //    FillStatementDataSet(pBankCode);
            //    if (DSstatement.Tables[1].Rows.Count == 0)
            //        {
            //        for (int i = 0; DSstatement.Tables[0].Rows.Count > 0; i++)
            //            {
            //            if (decimal.Parse(DSstatement.Tables[0].Rows[i].ItemArray[91].ToString()) != 0)
            //                closebal++;
            //            }
            //        if (closebal == 0)
            //            return false;
            //        }
            //    expReport.ExportSetup(clsBasFile.getPathWithoutFile(strFileName), clsBasFile.getFileFromPath(strFileName), pReportName, DSstatement);
            //    expReport.ExportSelection(pCurrType);
            //    exportResult = expReport.ExportCompletion();
            //    }
            else if (pBankName == "AIBK_Credit")
                {
                //MainTableCond = " m.contracttype in " + strProductCond + "";//strWhereCond
                //supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + ")";
                //expReport.selectionFormula = "if not isnull({@f_BranchTrans}) then {@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} in " + strProductCond.Replace('(', '[').Replace(')', ']') + " else {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} in " + strProductCond.Replace('(', '[').Replace(')', ']') + ""; //and {@f_closingbalance} <> 0 
                //expReport.selectionFormula = "((IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} in " + strProductCond.Replace('(', '[').Replace(')', ']') + " and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y')) or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} in " + strProductCond.Replace('(', '[').Replace(')', ']') + " and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y')";
                expReport.selectionFormula = "(IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y') or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y')";
                expReport.ExportSetup(clsBasFile.getPathWithoutFile(strFileName), clsBasFile.getFileFromPath(strFileName), pReportName, DSstatement);
                expReport.ExportSelection(pCurrType);
                exportResult = expReport.ExportCompletion();
                FillStatementDataSet(pBankCode, "");
                getClientEmail(pBankCode);
                DSstatement.Merge(DSemails.Tables[0]);
                }
            else if (pBankName == "AIBK_Credit_Customer")
                {
                //MainTableCond = " m.contracttype in " + strProductCond + "";//strWhereCond
                //supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + ")";
                //expReport.selectionFormula = "if not isnull({@f_BranchTrans}) then {@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} in " + strProductCond.Replace('(', '[').Replace(')', ']') + " else {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} in " + strProductCond.Replace('(', '[').Replace(')', ']') + ""; //and {@f_closingbalance} <> 0 
                //expReport.selectionFormula = "((IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} in " + strProductCond.Replace('(', '[').Replace(')', ']') + " and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y')) or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} in " + strProductCond.Replace('(', '[').Replace(')', ']') + " and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y')";
                expReport.selectionFormula = "(IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y' and not({TCLIENTPERSONE.EXTERNALID} like '9*' )) or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y' and not({TCLIENTPERSONE.EXTERNALID} like '9*'))";
                expReport.ExportSetup(clsBasFile.getPathWithoutFile(strFileName), clsBasFile.getFileFromPath(strFileName), pReportName, DSstatement);
                expReport.ExportSelection(pCurrType);
                exportResult = expReport.ExportCompletion();
                FillStatementDataSet(pBankCode, "");
                getClientEmail(pBankCode);
                DSstatement.Merge(DSemails.Tables[0]);
                }
            else if (pBankName == "AIBK_Credit_Staff")
                {
                //MainTableCond = " m.contracttype in " + strProductCond + "";//strWhereCond
                //supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + ")";
                //expReport.selectionFormula = "if not isnull({@f_BranchTrans}) then {@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} in " + strProductCond.Replace('(', '[').Replace(')', ']') + " else {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} in " + strProductCond.Replace('(', '[').Replace(')', ']') + ""; //and {@f_closingbalance} <> 0 
                //expReport.selectionFormula = "((IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} in " + strProductCond.Replace('(', '[').Replace(')', ']') + " and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y')) or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} in " + strProductCond.Replace('(', '[').Replace(')', ']') + " and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y')";
                expReport.selectionFormula = "(IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y' and {TCLIENTPERSONE.EXTERNALID} like '9*' ) or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y' and {TCLIENTPERSONE.EXTERNALID} like '9*' )";
                expReport.ExportSetup(clsBasFile.getPathWithoutFile(strFileName), clsBasFile.getFileFromPath(strFileName), pReportName, DSstatement);
                expReport.ExportSelection(pCurrType);
                exportResult = expReport.ExportCompletion();
                FillStatementDataSet(pBankCode, "");
                getClientEmail(pBankCode);
                DSstatement.Merge(DSemails.Tables[0]);
                }
            //else if (pBankName == "AIBK_Installment")
            //    {
            //    //MainTableCond = " m.contracttype in " + strProductCond + "";//strWhereCond
            //    //supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + ")";
            //    //expReport.selectionFormula = "if not isnull({@f_BranchTrans}) then {@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} in " + strProductCond.Replace('(', '[').Replace(')', ']') + " else {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} in " + strProductCond.Replace('(', '[').Replace(')', ']') + ""; //and {@f_closingbalance} <> 0 
            //    //expReport.selectionFormula = "((IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} in " + strProductCond.Replace('(', '[').Replace(')', ']') + " and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y')) or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} in " + strProductCond.Replace('(', '[').Replace(')', ']') + " and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y')";
            //    expReport.selectionFormula = "((IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y')) or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y')";
            //    expReport.ExportSetup(clsBasFile.getPathWithoutFile(strFileName), clsBasFile.getFileFromPath(strFileName), pReportName, DSstatement);
            //    expReport.ExportSelection(pCurrType);
            //    exportResult = expReport.ExportCompletion();
            //    FillStatementDataSet(pBankCode, "");
            //    getClientEmail(pBankCode);
            //    DSstatement.Merge(DSemails.Tables[0]);
            //    }
            else if (pBankName == "AIBK_Prepaid_Customer")
                {
                //MainTableCond = " m.contracttype in " + strProductCond + "";//strWhereCond
                //supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + ")";
                //expReport.selectionFormula = "if not isnull({@f_BranchTrans}) then {@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} in " + strProductCond.Replace('(', '[').Replace(')', ']') + " else {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} in " + strProductCond.Replace('(', '[').Replace(')', ']') + ""; //and {@f_closingbalance} <> 0 
                //expReport.selectionFormula = "((IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} in " + strProductCond.Replace('(', '[').Replace(')', ']') + " and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y')) or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} in " + strProductCond.Replace('(', '[').Replace(')', ']') + " and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y')";
                expReport.selectionFormula = "((IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y' and not({TCLIENTPERSONE.EXTERNALID} like '9*' )) or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y' and not({TCLIENTPERSONE.EXTERNALID} like '9*' )))";
                expReport.ExportSetup(clsBasFile.getPathWithoutFile(strFileName), clsBasFile.getFileFromPath(strFileName), pReportName, DSstatement);
                expReport.ExportSelection(pCurrType);
                exportResult = expReport.ExportCompletion();
                FillStatementDataSet(pBankCode, "");
                getClientEmail(pBankCode);
                DSstatement.Merge(DSemails.Tables[0]);
                }
            else if (pBankName == "AIBK_Prepaid_Staff")
                {
                //MainTableCond = " m.contracttype in " + strProductCond + "";//strWhereCond
                //supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + ")";
                //expReport.selectionFormula = "if not isnull({@f_BranchTrans}) then {@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} in " + strProductCond.Replace('(', '[').Replace(')', ']') + " else {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} in " + strProductCond.Replace('(', '[').Replace(')', ']') + ""; //and {@f_closingbalance} <> 0 
                //expReport.selectionFormula = "((IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} in " + strProductCond.Replace('(', '[').Replace(')', ']') + " and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y')) or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} in " + strProductCond.Replace('(', '[').Replace(')', ']') + " and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y')";
                expReport.selectionFormula = "((IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y' and {TCLIENTPERSONE.EXTERNALID} like '9*' ) or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.HOLSTMT} <> 'Y' and {TCLIENTPERSONE.EXTERNALID} like '9*' ))";
                expReport.ExportSetup(clsBasFile.getPathWithoutFile(strFileName), clsBasFile.getFileFromPath(strFileName), pReportName, DSstatement);
                expReport.ExportSelection(pCurrType);
                exportResult = expReport.ExportCompletion();
                FillStatementDataSet(pBankCode, "");
                getClientEmail(pBankCode);
                DSstatement.Merge(DSemails.Tables[0]);
                }
            else
                {
                expReport.ExportSetup(clsBasFile.getPathWithoutFile(strFileName), clsBasFile.getFileFromPath(strFileName), pReportName, DSstatement);
                expReport.ExportSelection(pCurrType);
                exportResult = expReport.ExportCompletion();
                FillStatementDataSet(pBankCode, "");
                }
            getCardProduct(pBankCode);
            frmMain.BeginInvoke(frmMain.setStatusDelegate, new object[] { 0, "Data Exported - " + exportResult });//\r\n   + "\r\n" 

            aryLstFiles.Add(strFileName + expReport.getExportType(pCurrType));
            if (!(pBankCode == 5 || pBankCode == 7 || pBankCode == 15 || pBankCode == 20 || pBankCode == 23 || pBankCode == 28 || pBankCode == 33 || pBankCode == 51 || pBankCode == 74 || pBankName == "AIBK_Credit_Customer" || pBankName == "AIBK_Credit_Staff" || pBankName == "AIBK_Prepaid_Staff"))
                {
                clsCalcStatSummary calcStatSummary = new clsCalcStatSummary();
                calcStatSummary.statRelation = StatementNoDRel;
                calcStatSummary.Statement(pStrFileName, pBankName, pBankCode
                , pStrFile, pDate, pStmntType, pAppendData, DSstatement, DSProducts);
                }
            }
        catch (OracleException ex)
            {
            clsDbOracleLayer.catchError(ex);
            rtrnVal = false;
            }
        catch (NotSupportedException ex)  //(Exception ex)  //
            {
            clsBasErrors.catchError(ex);
            rtrnVal = false;
            }
        //finally
        //{
        //    DSstatement = null;
        //    expReport = null;
        //}
        return rtrnVal;
        }

    public bool exportAttachmentBDCA(DataSet data, string pStrFileName, int pBankCode, string pExportName, string pReportName, string selformula)
    {
        bool rtrnVal = true;
        //var expReportdbca = new clsExportReport();
        try
        {
            curBranchVal = pBankCode;
            MainTableCond = strWhereCond;
            supTableCond = strWhereCondd;

            DSstatementForReports = new DataSet("MasterDetailDS");
            //DSstatement.Tables["tStatementDetailTable"].Clone();
            DataTable StatementDetailTableClone = data.Tables["tStatementDetailTable"].Clone();
            DataTable StatementMasterTableClone = data.Tables["tStatementMasterTable"].Clone();
            DSstatementForReports.Tables.Add(StatementDetailTableClone);
            DSstatementForReports.Tables.Add(StatementMasterTableClone);


            foreach (var item in data.Tables["tStatementDetailTable"].Select(strWhereCond))
            {
                StatementDetailTableClone.ImportRow(item);
            }

            foreach (var item in data.Tables["tStatementMasterTable"].Select(strWhereCond))
            {
                StatementMasterTableClone.ImportRow(item);
            }

            expReportdbca.selectionFormula = selformula;
            expReportdbca.exportPDF_ForBDCA(clsBasFile.getPathWithoutFile(pStrFileName), pExportName.Replace('.', ' ').Replace('/', ' ').Replace(' ', '_'), pReportName, DSstatementForReports, selformula);
        }
        catch (Exception ex)
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message + ex.InnerException + "\n" + ex.Source + "\n" + ex.StackTrace, System.IO.FileMode.Append);

            //clsBasErrors.catchError(ex.InnerException);
            rtrnVal = false;
            if (pBankCode == 41)
            {
                rtrnVal = true;
            }
        }
        finally
        {
            DSstatementForReports = null;
        }
        return rtrnVal;
    }
    public bool exportAttachment(string pStrFileName, int pBankCode, string pExportName, string pReportName)
        {
        bool rtrnVal = true;
        clsExportReport expReport = new clsExportReport();
        try
            {
            curBranchVal = pBankCode;
            MainTableCond = strWhereCond;
            supTableCond = strWhereCondd;
            FillStatementDataSet(pBankCode, "vip");
            expReport.selectionFormula = string.Empty;
            expReport.exportPDF(clsBasFile.getPathWithoutFile(pStrFileName), pExportName.Replace('.', ' ').Replace('/', ' ').Replace(' ', '_'), pReportName, DSstatement);
            }
        catch (NotSupportedException ex)
            {
            clsBasErrors.catchError(ex);
            rtrnVal = false;
            }
        return rtrnVal;
        }

    //public bool exportAttachment_UNBN(string pStrFileName, int pBankCode, string pExportName, string pReportName)
    //{
    //    bool rtrnVal = true;
    //    clsExportReport expReport = new clsExportReport();
    //    try
    //    {
    //        curBranchVal = pBankCode;
    //        MainTableCond = strWhereCond;
    //        supTableCond = strWhereCondd;
    //        FillStatementDataSet(pBankCode, "vip");
    //        expReport.selectionFormula = string.Empty;
    //        expReport.exportPDF_UNBN(clsBasFile.getPathWithoutFile(pStrFileName), pExportName.Replace('.', ' ').Replace('/', ' ').Replace(' ', '_'), pReportName, DSstatement);
    //    }
    //    catch (NotSupportedException ex)
    //    {
    //        clsBasErrors.catchError(ex);
    //        rtrnVal = false;
    //    }
    //    return rtrnVal;
    //}

    public bool exportAttachment(string pStrFileName, int pBankCode, string pExportName, string pReportName, string selformula)
        {
        bool rtrnVal = true;
        clsExportReport expReport = new clsExportReport();
        try
        {
            curBranchVal = pBankCode;
            MainTableCond = strWhereCond;
            supTableCond = strWhereCondd;
            if (pBankCode==122)
            {
                //FillStatementDataSet_SortCardPriority_Exclude_VisaCards
                FillStatementDataSet_Exclude_VisaCards(pBankCode, "vip");

            }else if (pBankCode == 41)
            {
                FillStatementDataSet_SortCardPriority(pBankCode, "");
            }
            else
            {
                FillStatementDataSet(pBankCode, "vip");

            }
            expReport.selectionFormula = selformula;
            expReport.exportPDF(clsBasFile.getPathWithoutFile(pStrFileName), pExportName.Replace('.', ' ').Replace('/', ' ').Replace(' ', '_'), pReportName, DSstatement, selformula);
        }
        catch (Exception ex)
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message + ex.InnerException + "\n" + ex.Source + "\n" + ex.StackTrace, System.IO.FileMode.Append);

            //clsBasErrors.catchError(ex.InnerException);
            rtrnVal = false;
            if (pBankCode == 41)
            {
              rtrnVal = true;
            }
        }
        finally
        {
            expReport = null;
        }
        return rtrnVal;
        }


    public bool exportAttachmentPDF(string pStrFileName, int pBankCode, string pExportName, string pReportName, string selformula)
    {
        bool rtrnVal = true;
        clsExportReport expReport = new clsExportReport();
        try
        {
            curBranchVal = pBankCode;
            MainTableCond = strWhereCond;
            supTableCond = strWhereCondd;
            if (pBankCode == 122)
            {
                //FillStatementDataSet_SortCardPriority_Exclude_VisaCards
                FillStatementDataSet_Exclude_VisaCards(pBankCode, "vip");

            }
            else if (pBankCode == 41)
            {
                FillStatementDataSet_SortCardPriority(pBankCode, "");
            }
            else
            {
                FillStatementDataSet(pBankCode, "vip");

            }
            expReport.selectionFormula = selformula;
            expReport.exportPDFPDF(clsBasFile.getPathWithoutFile(pStrFileName), pExportName.Replace('.', ' ').Replace('/', ' ').Replace(' ', '_'), pReportName, DSstatement, selformula);
        }
        catch (Exception ex)
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message + ex.InnerException + "\n" + ex.Source + "\n" + ex.StackTrace, System.IO.FileMode.Append);

            //clsBasErrors.catchError(ex.InnerException);
            rtrnVal = false;
            if (pBankCode == 41)
            {
                rtrnVal = true;
            }
        }
        finally
        {
            expReport = null;
        }
        return rtrnVal;
    }

    public bool exportAttachment_Rewards(string pStrFileName, int pBankCode, string pExportName, string pReportName, string selformula)
    {
        bool rtrnVal = true;
        clsExportReport expReport = new clsExportReport();

        try
        {


            //OracleDataAdapter da_UNBN = new OracleDataAdapter(strQry, connection);



            //(new OracleCommand("alter session set optimizer_goal = RULE", conn)).ExecuteNonQuery();
            //ds_UNBN = new DataSet();
            //da_UNBN.Fill(ds_UNBN, "TSTATEMENTMASTERTABLE");
            //connection.Close();

            curBranchVal = pBankCode;
            MainTableCond = strWhereCond; //"accountno = '0078749309'";//strWhereCond;
            supTableCond = strWhereCondd; //"accountno = '0078749309'"; //strWhereCondd;
            FillStatementDataSet(pBankCode, "vip");
            //ds_UNBN = DSstatement;
            //strQry += " AND accountno = '0078749309' "; //+ strWhereCondd;
            //cmd.CommandText = strQry;
            //connection.Open();
            //ds_UNBN.Tables.Add("TSTATEMENTMASTERTABLE_1").Load(cmd.ExecuteReader());
            //connection.Close();
            expReport.selectionFormula = selformula;
            try
            {
                expReport.exportPDF_Rewards(clsBasFile.getPathWithoutFile(pStrFileName), pExportName.Replace('.', ' ').Replace('/', ' ').Replace(' ', '_'), pReportName, DSstatement, selformula);

            }
            catch (Exception ex)
            {
                clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine +ex.Message+"\n"+ ex.InnerException, System.IO.FileMode.Append);
                rtrnVal = false;
            }

    }
        catch (Exception ex)
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + pExportName + "\n"+ex.Message + "\n" + ex.InnerException, System.IO.FileMode.Append);
            clsBasErrors.catchError(ex);
            rtrnVal = false;
        }
      
        return rtrnVal;
    }

    public bool SplitByProfile(string pStrFileName, string pBankName
    , int pBankCode, string pStrFile, string pReportName, ExportFormatType pCurrType, DateTime pDate, string pStmntType, bool pAppendData)
        {
        bool rtrnVal = true;
        string queryString, whereCond;
        //selFormula = "{@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) ";
        queryString = "select distinct t.contracttype from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " t where t.branch = " + pBankCode + strWhereCond + "";
        try
            {
            OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
            OracleCommand cmd = new OracleCommand(queryString, conn);
            conn.Open();
            OracleDataReader rdr;
            rdr = cmd.ExecuteReader();
            whereCond = strWhereCond;
            strCloseBalance = "ALL";
            while (rdr.Read())
                {
                strWhereCond = string.Empty;
                if (!rdr.IsDBNull(0))
                    {
                    strWhereCond = " contracttype = '" + rdr.GetValue(0) + "'" + whereCond;
                    supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and " + strWhereCond + ")";
                    //selFormula = "if not isnull({@f_BranchTrans}) then {@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} = '" + rdr.GetValue(0) + "' else {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} = '" + rdr.GetValue(0) + "'" + ""; //and {@f_closingbalance} <> 0 
                    //selFormula = "(IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 ) or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + ")";
                    selFormula = "(IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} = '" + rdr.GetValue(0) + "') or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} = '" + rdr.GetValue(0) + "')";
                    export(pStrFileName, pBankName, pBankCode, pStrFile + rdr.GetValue(0) + "_", pReportName, pCurrType, pDate, pStmntType, pAppendData);
                    }
                }
            rdr.Close();
            strWhereCond = whereCond;
            strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
            strFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM"); // + ".pdf"
            return true;
            }
        catch (OracleException ex)
            {
            clsDbOracleLayer.catchError(ex);
            rtrnVal = false;
            }
        catch (NotSupportedException ex)  //(Exception ex)  //
            {
            clsBasErrors.catchError(ex);
            rtrnVal = false;
            }

        //finally
        //{
        //}
        return rtrnVal;
        }

    public bool SplitByCurrency(string pStrFileName, string pBankName
    , int pBankCode, string pStrFile, string pReportName, ExportFormatType pCurrType, DateTime pDate, string pStmntType, bool pAppendData)
        {
        bool rtrnVal = true;
        string queryString, whereCond;
        //selFormula = "{@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) ";
        selFormula = "(IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0) or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " )";
        queryString = "select distinct t.accountcurrency from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " t where t.branch = " + pBankCode + strWhereCond + "";
        try
            {
            OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
            OracleCommand cmd = new OracleCommand(queryString, conn);
            conn.Open();
            OracleDataReader rdr;
            rdr = cmd.ExecuteReader();
            whereCond = strWhereCond;
            strCloseBalance = "ALL";
            while (rdr.Read())
                {
                strWhereCond = string.Empty;
                if (!rdr.IsDBNull(0))
                    {
                    strWhereCond = " accountcurrency = '" + Convert.ToInt32(rdr.GetValue(0)) + "'" + whereCond;
                    supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and " + strWhereCond + ")";
                    export(pStrFileName, pBankName, pBankCode, pStrFile + Convert.ToInt32(rdr.GetValue(0)) + "_", pReportName, pCurrType, pDate, pStmntType, pAppendData);
                    }
                }
            rdr.Close();
            strWhereCond = whereCond;
            strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
            strFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM"); // + ".pdf"
            return true;
            }
        catch (OracleException ex)
            {
            clsDbOracleLayer.catchError(ex);
            rtrnVal = false;
            }
        catch (NotSupportedException ex)  //(Exception ex)  //
            {
            clsBasErrors.catchError(ex);
            rtrnVal = false;
            }

        //finally
        //{
        //}
        return rtrnVal;
        }

    public bool SplitByBranchByProfile(string pStrFileName, string pBankName
    , int pBankCode, string pStrFile, string pReportName, ExportFormatType pCurrType, DateTime pDate, string pStmntType, bool pAppendData)
        {
        bool rtrnVal = true;
        string queryString, whereCond;
        queryString = "select distinct t.cardbranchpart, t.cardbranchpartname,t.contracttype from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " t where t.branch = " + pBankCode + strWhereCond + " ORDER BY CARDBRANCHPARTNAME";
        try
            {
            OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
            OracleCommand cmd = new OracleCommand(queryString, conn);
            conn.Open();
            OracleDataReader rdr;
            rdr = cmd.ExecuteReader();
            whereCond = strWhereCond;
            strCloseBalance = "ALL";
            while (rdr.Read())
                {
                strWhereCond = string.Empty;
                if (!rdr.IsDBNull(1))
                    {
                    //selFormula = "(IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} = " + strProductCond + " and {TSTATEMENTMASTERTABLE.CARDBRANCHPART} = " + Convert.ToInt32(rdr.GetValue(0)).ToString() + ") or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} = " + strProductCond + " and {TSTATEMENTMASTERTABLE.CARDBRANCHPART} = " + Convert.ToInt32(rdr.GetValue(0)).ToString() + ")";
                        if (pBankName.Equals("ZENG_Prepaid")) // condition over cardproduct
                        {
                            selFormula = "(IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.CARDPRODUCT} in [" + strProductCond + "] and {TSTATEMENTMASTERTABLE.CARDBRANCHPART} = " + Convert.ToInt32(rdr.GetValue(0)).ToString() + ") or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.CARDPRODUCT} in [" + strProductCond + "] and {TSTATEMENTMASTERTABLE.CARDBRANCHPART} = " + Convert.ToInt32(rdr.GetValue(0)).ToString() + ")";
                        }
                        else // condition for contracttype
                        {
                            selFormula = "(IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} = " + strProductCond + " and {TSTATEMENTMASTERTABLE.CARDBRANCHPART} = " + Convert.ToInt32(rdr.GetValue(0)).ToString() + ") or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} = " + strProductCond + " and {TSTATEMENTMASTERTABLE.CARDBRANCHPART} = " + Convert.ToInt32(rdr.GetValue(0)).ToString() + ")";
                        }

                    strWhereCond = " cardbranchpart = " + Convert.ToInt32(rdr.GetValue(0)).ToString() + whereCond;
                    supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and " + strWhereCond + ")";
                    //export(pStrFileName, pBankName, pBankCode, pStrFile + rdr.GetString(1) + "_" + rdr.GetString(2).Substring(rdr.GetString(2).IndexOf('-') + 2) + "_", pReportName, pCurrType, pDate, pStmntType, pAppendData);
                    export(pStrFileName, pBankName, pBankCode, pStrFile + rdr.GetString(1) + "_", pReportName, pCurrType, pDate, pStmntType, pAppendData);
                    }
                else //UnDefined brnaches
                    {
                    //selFormula = "{@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) and isNull({TSTATEMENTMASTERTABLE.CARDBRANCHPART}) "; //and {@f_closingbalance} <> 0 
                    selFormula = "(IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and isNull({TSTATEMENTMASTERTABLE.CARDBRANCHPART}) or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and isNull({TSTATEMENTMASTERTABLE.CARDBRANCHPART})";
                    strWhereCond = " cardbranchpartname is null " + whereCond;
                    supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and " + strWhereCond + ")";
                    export(pStrFileName, pBankName, pBankCode, pStrFile + "Not_Defined_Branches_", pReportName, pCurrType, pDate, pStmntType, pAppendData);
                    break;
                    }
                }
            rdr.Close();
            if (pBankCode == 7 || pBankCode == 15)
                {
                getCardProduct(pBankCode);
                FillStatementHistoryDataSet(pBankCode);
                clsCalcStatSummary calcStatSummary = new clsCalcStatSummary();
                calcStatSummary.statRelation = StatementNoDRel;
                if (pBankCode == 7)
                    calcStatSummary.emailService = true;
                else
                    calcStatSummary.emailService = false;
                calcStatSummary.Statement(pStrFileName, pBankName, pBankCode, pStrFile, pDate, pStmntType, pAppendData, DSstatementHist, DSProducts);
                }

            if (pBankCode == 33)
                {
                //MainTableCond = " m.contracttype in " + strProductCond + "";//strWhereCond
                //supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + ")";
                    if (pBankName.Equals("ZENG_Prepaid")) // condition over cardproduct
                    {
                        MainTableCond = " m.cardproduct in (" + strProductCond + ")";//strWhereCond
                        supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.cardproduct in (" + strProductCond + "))";
                    }
                    else // condition over contracttype
                    {
                        MainTableCond = " m.contracttype in " + strProductCond + "";//strWhereCond
                        supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + ")";
                    }
                FillStatementDataSet(pBankCode, "");
                clsCalcStatSummary calcStatSummary = new clsCalcStatSummary();
                calcStatSummary.statRelation = StatementNoDRel;
                calcStatSummary.Statement(pStrFileName, pBankName, pBankCode
                , pStrFile, pDate, pStmntType, pAppendData, DSstatement, DSProducts);
                }

            strWhereCond = whereCond;
            strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
            strFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM"); // + ".pdf"
            return true;
            }
        catch (OracleException ex)
            {
            clsDbOracleLayer.catchError(ex);
            rtrnVal = false;
            }
        catch (NotSupportedException ex)  //(Exception ex)  //
            {
            clsBasErrors.catchError(ex);
            rtrnVal = false;
            }

        //finally
        //{
        //}
        return rtrnVal;
        }

    public bool SplitByBranch(string pStrFileName, string pBankName
    , int pBankCode, string pStrFile, string pReportName, ExportFormatType pCurrType, DateTime pDate, string pStmntType, bool pAppendData)
        {
        bool rtrnVal = true;
        string queryString, whereCond;
        //selFormula = "if not isnull({@f_BranchTrans}) then {@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) else {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 " + ""; //and {@f_closingbalance} <> 0 
        queryString = "select distinct t.cardbranchpart, t.cardbranchpartname from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " t where t.branch = " + pBankCode + strWhereCond + " ORDER BY CARDBRANCHPARTNAME";
        try
            {
            OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
            OracleCommand cmd = new OracleCommand(queryString, conn);
            conn.Open();
            OracleDataReader rdr;
            rdr = cmd.ExecuteReader();
            whereCond = strWhereCond;
            strCloseBalance = "ALL";
            while (rdr.Read())
                {
                strWhereCond = string.Empty;
                if (!rdr.IsDBNull(1))
                    {
                    if (pBankCode == 23)
                        //selFormula = "if not isnull({@f_BranchTrans}) then {@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} = " + strProductCond + " and {TSTATEMENTMASTERTABLE.CARDBRANCHPART} = " + Convert.ToInt32(rdr.GetValue(0)).ToString() + " else {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0  and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} = " + strProductCond + " and {TSTATEMENTMASTERTABLE.CARDBRANCHPART} = " + Convert.ToInt32(rdr.GetValue(0)).ToString() + ""; //and {@f_closingbalance} <> 0 
                        //selFormula = "{@F_branch} = " + pBankCode + " and {@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} = " + strProductCond + " and {TSTATEMENTMASTERTABLE.CARDBRANCHPART} =" + Convert.ToInt32(rdr.GetValue(0)).ToString();
                        selFormula = "(IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} = " + strProductCond + " and {TSTATEMENTMASTERTABLE.CARDBRANCHPART} = " + Convert.ToInt32(rdr.GetValue(0)).ToString() + ") or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} = " + strProductCond + " and {TSTATEMENTMASTERTABLE.CARDBRANCHPART} = " + Convert.ToInt32(rdr.GetValue(0)).ToString() + ")";
                    else if (pBankCode == 7 && !string.IsNullOrEmpty(strProductCond))
                        selFormula = "(IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.CARDPRODUCT} = " + strProductCond + " and {TSTATEMENTMASTERTABLE.CARDBRANCHPART} = " + Convert.ToInt32(rdr.GetValue(0)).ToString() + ") or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.CARDPRODUCT} = " + strProductCond + " and {TSTATEMENTMASTERTABLE.CARDBRANCHPART} = " + Convert.ToInt32(rdr.GetValue(0)).ToString() + ")";
                    else
                        //selFormula = "{@F_branch} = " + pBankCode + " and {@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and {TSTATEMENTMASTERTABLE.CARDBRANCHPART} = " + Convert.ToInt32(rdr.GetValue(0)).ToString(); //and {@f_closingbalance} <> 0 
                        selFormula = "(IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.CARDBRANCHPART} = " + Convert.ToInt32(rdr.GetValue(0)).ToString() + ") or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.CARDBRANCHPART} = " + Convert.ToInt32(rdr.GetValue(0)).ToString() + ")";
                    //selFormula = "{@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) and {TSTATEMENTMASTERTABLE.CARDBRANCHPART} = " + Convert.ToInt32(rdr.GetValue(0)).ToString() + " else {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.CARDBRANCHPART} = " + Convert.ToInt32(rdr.GetValue(0)).ToString() + "";

                    strWhereCond = " cardbranchpart = " + Convert.ToInt32(rdr.GetValue(0)).ToString() + whereCond;
                    supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and " + strWhereCond + ")";
                    export(pStrFileName, pBankName, pBankCode, pStrFile + rdr.GetString(1) + "_", pReportName, pCurrType, pDate, pStmntType, pAppendData);
                    }
                //else //UnDefined brnaches
                //    {
                //    //selFormula = "{@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) and isNull({TSTATEMENTMASTERTABLE.CARDBRANCHPART}) "; //and {@f_closingbalance} <> 0 
                //    selFormula = "(IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and isNull({TSTATEMENTMASTERTABLE.CARDBRANCHPART}) or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and isNull({TSTATEMENTMASTERTABLE.CARDBRANCHPART})";
                //    strWhereCond = " cardbranchpartname is null " + whereCond;
                //    supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and " + strWhereCond + ")";
                //    export(pStrFileName, pBankName, pBankCode, pStrFile + "Not_Defined_Branches_", pReportName, pCurrType, pDate, pStmntType, pAppendData);
                //    break;
                //    }
                }
            rdr.Close();
            if (pBankCode == 7 || pBankCode == 15)
                {
                getCardProduct(pBankCode);
                FillStatementHistoryDataSet(pBankCode);
                clsCalcStatSummary calcStatSummary = new clsCalcStatSummary();
                calcStatSummary.statRelation = StatementNoDRel;
                if (pBankCode == 7)
                    calcStatSummary.emailService = true;
                else
                    calcStatSummary.emailService = false;
                calcStatSummary.Statement(pStrFileName, pBankName, pBankCode, pStrFile, pDate, pStmntType, pAppendData, DSstatementHist, DSProducts);
                }
            strWhereCond = whereCond;
            strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
            strFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM"); // + ".pdf"
            return true;
            }
        catch (OracleException ex)
            {
            clsDbOracleLayer.catchError(ex);
            rtrnVal = false;
            }
        catch (NotSupportedException ex)  //(Exception ex)  //
            {
            clsBasErrors.catchError(ex);
            rtrnVal = false;
            }

        //finally
        //{
        //}
        return rtrnVal;
        }

    public bool SplitByProduct(string pStrFileName, string pBankName
    , int pBankCode, string pStrFile, string pReportName, ExportFormatType pCurrType, DateTime pDate, string pStmntType, bool pAppendData)
        {
        bool rtrnVal = true;
        string queryString, whereCond;
        queryString = "select distinct t.cardproduct from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " t where t.branch = " + pBankCode + strWhereCond + " ORDER BY cardproduct";
        try
            {
            OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
            OracleCommand cmd = new OracleCommand(queryString, conn);
            conn.Open();
            OracleDataReader rdr;
            rdr = cmd.ExecuteReader();
            whereCond = strWhereCond;
            strCloseBalance = "ALL";
            while (rdr.Read())
                {
                strWhereCond = string.Empty;
                if (!rdr.IsDBNull(0))
                    {
                    strWhereCond = " cardproduct = '" + rdr.GetValue(0) + "'" + whereCond;
                    supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and " + strWhereCond + ")";
                    //selFormula = "if not isnull({@f_BranchTrans}) then {@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) and {@f_cardproduct} = '" + rdr.GetValue(0) + "' else {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {@f_cardproduct} = '" + rdr.GetValue(0) + "'" + ""; //and {@f_closingbalance} <> 0 
                    selFormula = "(IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {@f_cardproduct} = " + rdr.GetValue(0).ToString() + ") or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_cardproduct} = " + rdr.GetValue(0).ToString() + ")";
                    export(pStrFileName, pBankName, pBankCode, pStrFile + rdr.GetValue(0) + "_", pReportName, pCurrType, pDate, pStmntType, pAppendData);
                    }
                }
            rdr.Close();
            strWhereCond = whereCond;
            strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
            strFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM"); // + ".pdf"
            return true;
            }
        catch (OracleException ex)
            {
            clsDbOracleLayer.catchError(ex);
            rtrnVal = false;
            }
        catch (NotSupportedException ex)  //(Exception ex)  //
            {
            clsBasErrors.catchError(ex);
            rtrnVal = false;
            }

        //finally
        //{
        //}
        return rtrnVal;
        }

    public bool SplitByDbCr(string pStrFileName, string pBankName
  , int pBankCode, string pStrFile, string pReportName, ExportFormatType pCurrType, DateTime pDate, string pStmntType, bool pAppendData)
        {
        string orgStrWhere;
        orgStrWhere = strWhereCond;

        //if (pBankCode == 23)
        //    {
        //    FillStatementDataSet(pBankCode, "");
        //    clsCalcStatSummary calcStatSummary = new clsCalcStatSummary();
        //    calcStatSummary.statRelation = StatementNoDRel;
        //    calcStatSummary.Statement(pStrFileName, pBankName, pBankCode
        //    , pStrFile, pDate, pStmntType, pAppendData, DSstatement, DSProducts);
        //    }
        selFormula = "{@F_branch} = " + pBankCode + " and {@f_closingbalance} = 0 and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} = " + strProductCond + "";
        //selFormula = "(IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} = 0 and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} = " + strProductCond + ") or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} = " + strProductCond + ")";
        //strWhereCond += " and closingbalance = 0";
        strWhereCond = " closingbalance = 0 " + strWhereCond;
        supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and " + strWhereCond + ")";
        export(pStrFileName, pBankName, pBankCode, pStrFile + "Balance Equal Zero_", pReportName, pCurrType, pDate, pStmntType, pAppendData);

        strWhereCond = orgStrWhere;
        selFormula = "{@F_branch} = " + pBankCode + " and {@f_closingbalance} > 0 and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} = " + strProductCond + "";
        //selFormula = "(IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} > 0 and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} = " + strProductCond + ") or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} = " + strProductCond + ")";
        //strWhereCond += " and closingbalance > 0";
        strWhereCond = " closingbalance > 0 " + strWhereCond;
        supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and " + strWhereCond + ")";
        export(pStrFileName, pBankName, pBankCode, pStrFile + "Balance Greater Zero_", pReportName, pCurrType, pDate, pStmntType, pAppendData);

        strWhereCond = orgStrWhere;
        selFormula = "{@F_branch} = " + pBankCode + " and {@f_closingbalance} < 0 and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} = " + strProductCond + "";
        //selFormula = "(IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} < 0 and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} = " + strProductCond + ") or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.CONTRACTTYPE} = " + strProductCond + ")";
        //strWhereCond += " and closingbalance < 0";
        strWhereCond = " closingbalance < 0 " + strWhereCond;
        supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and " + strWhereCond + ")";
        export(pStrFileName, pBankName, pBankCode, pStrFile + "Balance Less Zero_", pReportName, pCurrType, pDate, pStmntType, pAppendData);

        //selFormula = "{@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) ";//and {@f_closingbalance} <> 0 
        selFormula = "(IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0) or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " )";
        strWhereCond = orgStrWhere;
        strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
        strFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM"); // + ".pdf"

        if (pBankCode == 23)
            {
            MainTableCond = " m.contracttype in " + strProductCond + "";//strWhereCond
            supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + ")";
            FillStatementDataSet(pBankCode, "");
            clsCalcStatSummary calcStatSummary = new clsCalcStatSummary();
            calcStatSummary.statRelation = StatementNoDRel;
            calcStatSummary.Statement(pStrFileName, pBankName, pBankCode
            , pStrFile, pDate, pStmntType, pAppendData, DSstatement, DSProducts);
            }

        return true;
        }

    public bool SplitByCompany(string pStrFileName, string pBankName
    , int pBankCode, string pStrFile, string pReportName, ExportFormatType pCurrType, DateTime pDate, string pStmntType, bool pAppendData)
        {
        bool rtrnVal = true;
        string queryString, whereCond;
        //selFormula = "{@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) ";
        selFormula = "(IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0) or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " )";
        queryString = "select distinct tm.branch,tcp.company,tcb.name from " + MainSchema + "tclientbank tcb," + MainSchema + "tclientpersone tcp," + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " tm "
        + "where tcb.branch = tcp.branch and tcb.companycode = tcp.company and tcb.branch = tm.branch and TCP.IDCLIENT = tm.clientid and tcp.branch = tm.branch "
        + "and tm.branch = " + pBankCode + strWhereCond + " ORDER BY tcp.company";
        try
            {
            OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
            OracleCommand cmd = new OracleCommand(queryString, conn);
            conn.Open();
            OracleDataReader rdr;
            rdr = cmd.ExecuteReader();
            whereCond = strWhereCond;
            strCloseBalance = "ALL";
            while (rdr.Read())
                {
                strWhereCond = string.Empty;
                if (!rdr.IsDBNull(1))
                    {
                    //strWhereCond = " clientid in (select idclient from tclientpersone where company in "
                    //+ "(select companycode from tclientbank) and company = " + rdr.GetInt32(1).ToString() + ") and branch =" + pBankCode + whereCond;
                    strWhereCond = " clientid in (select idclient from tclientpersone where company = " + rdr.GetInt32(1).ToString() + ") and branch =" + pBankCode + whereCond;
                    supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and " + strWhereCond + ")";
                    export(pStrFileName, pBankName, pBankCode, pStrFile + rdr.GetInt32(1) + "_" + rdr.GetString(2).Replace("/", " ") + "_", pReportName, pCurrType, pDate, pStmntType, pAppendData);
                    }
                }
            rdr.Close();
            //UnDefined brnaches
            queryString = "select distinct tm.branch,tcp.company from " + MainSchema + "tclientpersone tcp," + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " tm "
            + "where tcp.company is null and TCP.IDCLIENT = tm.clientid and tcp.branch = tm.branch "
            + "and tm.branch = " + pBankCode;
            OracleConnection conn1 = new OracleConnection(clsDbCon.sConOracle);
            OracleCommand cmd1 = new OracleCommand(queryString, conn1);
            conn1.Open();
            OracleDataReader rdr1;
            rdr1 = cmd1.ExecuteReader();
            whereCond = strWhereCond;
            strCloseBalance = "ALL";
            while (rdr1.Read())
                {
                strWhereCond = string.Empty;
                if (rdr1.IsDBNull(1))
                    {
                    strWhereCond = " clientid in (select idclient from " + MainSchema + "tclientpersone where company is null and branch =" + pBankCode + ")";
                    supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and " + strWhereCond + ")";
                    export(pStrFileName, pBankName, pBankCode, pStrFile + "Company Not Defined" + "_", pReportName, pCurrType, pDate, pStmntType, pAppendData);
                    break;
                    }
                }
            rdr1.Close();
            if (pBankCode == 1)
                {
                FillStatementHistoryDataSet(pBankCode);
                clsCalcStatSummary calcStatSummary = new clsCalcStatSummary();
                calcStatSummary.statRelation = StatementNoDRel;
                calcStatSummary.Statement(pStrFileName, pBankName, pBankCode
                , pStrFile, pDate, pStmntType, pAppendData, DSstatementHist, DSProducts);
                }
            strWhereCond = whereCond;
            strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
            strFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM"); // + ".pdf"
            return true;
            }
        catch (OracleException ex)
            {
            clsDbOracleLayer.catchError(ex);
            rtrnVal = false;
            }
        catch (NotSupportedException ex)  //(Exception ex)  //
            {
            clsBasErrors.catchError(ex);
            rtrnVal = false;
            }

        finally
            {
            }
        return rtrnVal;
        }

    public void CreateZip()
        {
        clsBasFile.generateFileMD5(aryLstFiles, strFileName + ".MD5");
        aryLstFiles.Add(strFileName + ".MD5");
        fileSummaryName = strFileName + ".txt";
        fileSummaryName = clsBasFile.getPathWithoutExtn(fileSummaryName) +
          "_Summary." + clsBasFile.getFileExtn(fileSummaryName);
        aryLstFiles.Add(@fileSummaryName);
        SharpZip zip = new SharpZip();
        zip.createZip(aryLstFiles, strFileName + ".zip", "");
        }

    public string closeBalance // " <> 0"
        {
        set { strCloseBalance = value; }
        }// setCloseBanace

    public string StatLable
        {
        get { return StrStatLable; }
        set { StrStatLable = value; }
        }// StatLable

    public string whereCond
        {
        get { return strWhereCond; }
        set { strWhereCond = value; }
        }  // whereCond

    public string whereCondD
        {
        get { return strWhereCondd; }
        set { strWhereCondd = value; }
        }  // whereCondD

    public frmStatementFile setFrm
        {
        set { frmMain = value; }
        }// setFrm

    public string PrepaidCondition
        {
        get { return prepaidCond; }
        set { prepaidCond = value; }
        }// PrepaidCondition

    public string productCond
        {
        get { return strProductCond; }
        set { strProductCond = value; }
        }  // productCond

    ~clsStatement_ExportRpt()
        {
        }
    }

