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

public class clsStatement_ExportRptRAS : clsBasStatement
    {
    private DateTime vCurDate;
    private string strOutputPath, strFileName;
    private string StrStatLable = string.Empty, strWhereCond = string.Empty;
    private ArrayList aryLstFiles = new ArrayList();
    private string selFormula = string.Empty;
    private string strCloseBalance = string.Empty;
    private string fileSummaryName = string.Empty;
    private frmStatementFile frmMain;
    protected string prepaidCond = string.Empty;
    private string strProductCond = string.Empty;

    public clsStatement_ExportRptRAS()
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

    public bool export(string pStrFileName, string pBankName
      , int pBankCode, string pStrFile, string pReportName, ExportFormatType pCurrType, DateTime pDate, string pStmntType, bool pAppendData)
        {
        bool rtrnVal = true;
        string strCloseBalanceFrml = string.Empty, strCloseBalanceSql = string.Empty;
        clsExportReportRAS expReport = new clsExportReportRAS();
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
            expReport.selectionFormula = "if not isnull({@f_BranchTrans}) then {@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) else {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 " + ""; //and {@f_closingbalance} <> 0 
        else
            expReport.selectionFormula = selFormula;

        try
            {
            string exportResult;

            pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);
            vCurDate = pDate; // DateTime.Now.AddMonths(-1);
            strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
            clsBasFile.createDirectory(pStrFileName + vCurDate.ToString("yyyyMM") + pBankName);
            strFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM"); // + ".pdf"

            curBranchVal = pBankCode;
            MainTableCond = strWhereCond;//strWhereCond
            if (pBankName == "ICBG_Prepaid")
                {
                MainTableCond += " m.cardproduct in " + PrepaidCondition;
                FillStatementDataSet(pBankCode, "vip");
                expReport.selectionFormula = "if not isnull({@f_BranchTrans}) then {@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) and {@f_cardproduct} in " + prepaidCond.Replace('(', '[').Replace(')', ']') + " else {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {@f_cardproduct} in " + prepaidCond.Replace('(', '[').Replace(')', ']') + ""; //and {@f_closingbalance} <> 0 
                }
            else if (pBankCode == 74)
                {
                // merge mark-up fee with original transaction
                clsMaintainData maintainData = new clsMaintainData();
                maintainData.mergeMarkUpFees(pBankCode);
                maintainData = null;
                MainTableCond = " m.contracttype in " + strProductCond + "";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + ")";
                FillStatementDataSet(pBankCode);
                expReport.selectionFormula = "if not isnull({@f_BranchTrans}) then {@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) and {@f_cardproduct} in " + strProductCond.Replace('(', '[').Replace(')', ']') + " else {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {@f_cardproduct} in " + strProductCond.Replace('(', '[').Replace(')', ']') + ""; //and {@f_closingbalance} <> 0 
                }
            else if (pBankCode == 24)
                {
                //MainTableCond = "rownum < 2000 ";//strWhereCond
                FillStatementDataSet(pBankCode, "vip");
                }
            else if (pBankCode == 45)
                {
                //MainTableCond = "rownum < 2000 ";//strWhereCond
                FillStatementDataSet(pBankCode, "vip");
                }
            else
                {
                FillStatementDataSet(pBankCode);
                }
            getCardProduct(pBankCode);
            if ((Convert.ToString(pBankCode) == "56") || (Convert.ToString(pBankCode) == "74"))
                {
                getClientEmail(pBankCode);
                DSstatement.Merge(DSemails.Tables[0]);
                }
            frmMain.BeginInvoke(frmMain.setStatusDelegate, new object[] { 0, "Data Retrieved" });//\r\n\r\n
            expReport.StatLable = StrStatLable;
            //expReport.ExportSetup(clsBasFile.getPathWithoutFile(strFileName), clsBasFile.getFileFromPath(strFileName), pReportName, arrTables);
            expReport.ExportSetup(clsBasFile.getPathWithoutFile(strFileName), clsBasFile.getFileFromPath(strFileName), pReportName);
            //expReport.exportPDF(clsBasFile.getPathWithoutFile(strFileName), clsBasFile.getFileFromPath(strFileName), pReportName, DSstatement);
            //expReport.ExportSelection(pCurrType);
            exportResult = expReport.ExportCompletion();
            frmMain.BeginInvoke(frmMain.setStatusDelegate, new object[] { 0, "Data Exported - " + exportResult });//\r\n   + "\r\n" 

            aryLstFiles.Add(strFileName + expReport.getExportType(pCurrType));

            //if (!(exportResult.IndexOf("Report Export Done Successfully.") > -1))
            //{
            //    rtrnVal = false;
            //}

            //if (pStmntType.ToLower() != "debit")
            //{
            if (!(pBankCode == 5 || pBankCode == 7 || pBankCode == 20 || pBankCode == 23 || pBankCode == 28 || pBankCode == 51 || pBankCode == 74))
            //if (!(pBankCode == 7 || pBankCode == 23 || pBankCode == 1 || pBankCode == 5 || pBankCode == 51))
                {
                clsCalcStatSummary calcStatSummary = new clsCalcStatSummary();
                //calcStatSummary.statType = "PDF";
                calcStatSummary.statRelation = StatementNoDRel;
                //calcStatSummary.isDebitVal = pStmntType.ToLower() == "debit" ? true : false;
                calcStatSummary.Statement(pStrFileName, pBankName, pBankCode
                , pStrFile, pDate, pStmntType, pAppendData, DSstatement, DSProducts);
                //}
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

    public bool SplitByProfile(string pStrFileName, string pBankName
    , int pBankCode, string pStrFile, string pReportName, ExportFormatType pCurrType, DateTime pDate, string pStmntType, bool pAppendData)
        {
        bool rtrnVal = true;
        string queryString, whereCond;
        selFormula = "{@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) ";
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
                    strWhereCond = " contracttype = '" + Convert.ToInt32(rdr.GetValue(0)) + "'" + whereCond;
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

    public bool SplitByCurrency(string pStrFileName, string pBankName
    , int pBankCode, string pStrFile, string pReportName, ExportFormatType pCurrType, DateTime pDate, string pStmntType, bool pAppendData)
        {
        bool rtrnVal = true;
        string queryString, whereCond;
        selFormula = "{@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) ";
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
            if (pBankCode == 7)
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
        selFormula = "if not isnull({@f_BranchTrans}) then {@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) else {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 " + ""; //and {@f_closingbalance} <> 0 
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
                    selFormula = "if not isnull({@f_BranchTrans}) then {@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) and {TSTATEMENTMASTERTABLE.CARDBRANCHPART} = " + Convert.ToInt32(rdr.GetValue(0)).ToString() + " else {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0  and {TSTATEMENTMASTERTABLE.CARDBRANCHPART} = " + Convert.ToInt32(rdr.GetValue(0)).ToString() + ""; //and {@f_closingbalance} <> 0 
                    strWhereCond = " cardbranchpart = " + Convert.ToInt32(rdr.GetValue(0)).ToString() + whereCond;
                    supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and " + strWhereCond + ")";
                    export(pStrFileName, pBankName, pBankCode, pStrFile + rdr.GetString(1) + "_", pReportName, pCurrType, pDate, pStmntType, pAppendData);
                    }
                else //UnDefined brnaches
                    {
                    selFormula = "if not isnull({@f_BranchTrans}) then {@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) and isNull({TSTATEMENTMASTERTABLE.CARDBRANCHPART}) else {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and isNull({TSTATEMENTMASTERTABLE.CARDBRANCHPART}) " + ""; //and {@f_closingbalance} <> 0 
                    strWhereCond = " cardbranchpartname is null " + whereCond;
                    supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and " + strWhereCond + ")";
                    export(pStrFileName, pBankName, pBankCode, pStrFile + "Branches Not Defined" + "_", pReportName, pCurrType, pDate, pStmntType, pAppendData);
                    break;
                    }
                }
            rdr.Close();
            if (pBankCode == 7)
                {
                getCardProduct(pBankCode);
                FillStatementHistoryDataSet(pBankCode);
                clsCalcStatSummary calcStatSummary = new clsCalcStatSummary();
                calcStatSummary.statRelation = StatementNoDRel;
                calcStatSummary.emailService = true;
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
        //selFormula = "{@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) ";
        selFormula = "if not isnull({@f_BranchTrans}) then {@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) else {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 " + ""; //and {@f_closingbalance} <> 0 
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
                    strWhereCond = " cardproduct = '" + Convert.ToInt32(rdr.GetValue(0)) + "'" + whereCond;
                    supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and " + strWhereCond + ")";
                    export(pStrFileName, pBankName, pBankCode, pStrFile + Convert.ToInt32(rdr.GetValue(0)) + "_", pReportName, pCurrType, pDate, pStmntType, pAppendData);
                    }
                }
            rdr.Close();
            //if (pBankCode == 15)
            //{
            //    FillStatementHistoryDataSet(pBankCode);
            //    clsCalcStatSummary calcStatSummary = new clsCalcStatSummary();
            //    calcStatSummary.statRelation = StatementNoDRel;
            //    calcStatSummary.Statement(pStrFileName, pBankName, pBankCode
            //    , pStrFile, pDate, pStmntType, pAppendData, DSstatementHist, DSProducts);
            //}
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
        selFormula = "{@F_branch} = " + pBankCode + " and {@f_closingbalance} = 0 and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) ";
        //strWhereCond += " and closingbalance = 0";
        strWhereCond = " closingbalance = 0 " + strWhereCond;
        supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and " + strWhereCond + ")";
        export(pStrFileName, pBankName, pBankCode, pStrFile + "Balance Equal Zero_", pReportName, pCurrType, pDate, pStmntType, pAppendData);

        strWhereCond = orgStrWhere;
        selFormula = "{@F_branch} = " + pBankCode + " and {@f_closingbalance} > 0 and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) ";
        //strWhereCond += " and closingbalance > 0";
        strWhereCond = " closingbalance > 0 " + strWhereCond;
        supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and " + strWhereCond + ")";
        export(pStrFileName, pBankName, pBankCode, pStrFile + "Balance Greater Zero_", pReportName, pCurrType, pDate, pStmntType, pAppendData);

        strWhereCond = orgStrWhere;
        selFormula = "{@F_branch} = " + pBankCode + " and {@f_closingbalance} < 0 and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) ";
        //strWhereCond += " and closingbalance < 0";
        strWhereCond = " closingbalance < 0 " + strWhereCond;
        supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and " + strWhereCond + ")";
        export(pStrFileName, pBankName, pBankCode, pStrFile + "Balance Less Zero_", pReportName, pCurrType, pDate, pStmntType, pAppendData);

        selFormula = "{@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) ";//and {@f_closingbalance} <> 0 
        strWhereCond = orgStrWhere;
        strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
        strFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM"); // + ".pdf"

        if (pBankCode == 23)
            {
            FillStatementDataSet(pBankCode);
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
        selFormula = "{@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) ";
        queryString = "select distinct tm.branch,tcp.company,tcb.name from "+MainSchema+"tclientbank tcb,"+MainSchema+"tclientpersone tcp," + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " tm "
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
                    strWhereCond = " clientid in (select idclient from " + MainSchema + "tclientpersone where company = " + rdr.GetInt32(1).ToString() + ") and branch =" + pBankCode + whereCond;
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
                    strWhereCond = " clientid in (select idclient from tclientpersone where company is null and branch =" + pBankCode + ")";
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

    ~clsStatement_ExportRptRAS()
        {
        }
    }

