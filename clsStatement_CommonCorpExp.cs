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

// Branch (SSB 15)
public class clsStatement_CommonCorpExp : clsBasStatement
    {
    private DateTime vCurDate;
    private string strOutputPath, strFileName, reportNameComp, reportNameIndv, strFileNameIndv, strFileNameComp, strFileMD5;
    private string StrStatLable = string.Empty, strWhereCond = string.Empty, strProductCond = string.Empty;
    private bool createCorporateVal = false;
    private ArrayList aryLstFiles = new ArrayList();

    public clsStatement_CommonCorpExp()
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

    public bool SplitByCurrency(string pStrFileName, string pBankName
    , int pBankCode, string pStrFile, string pReportName, ExportFormatType pCurrType, DateTime pDate, string pStmntType, bool pAppendData)
        {
        bool rtrnVal = true;
        string queryString, whereCond;
        queryString = "select distinct t.accountcurrency from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " t where t.branch = " + pBankCode + strWhereCond + "";
        try
            {
            OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
            OracleCommand cmd = new OracleCommand(queryString, conn);
            conn.Open();
            OracleDataReader rdr;
            rdr = cmd.ExecuteReader();
            whereCond = strWhereCond;

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
    public bool export(string pStrFileName, string pBankName
    , int pBankCode, string pStrFile, string pReportName, ExportFormatType pCurrType, DateTime pDate, string pStmntType, bool pAppendData)
        {
        string rtrnStr = "Successfully Generate " + pBankName;
        int curMonth = pDate.Month;
        bool preExit = true, rtrnVal = true;

        //clsMaintainData maintainData = new clsMaintainData();
        //maintainData.matchCardBranch4Account(pBankCode);

        clsExportReport expReport = new clsExportReport();
        expReport.bankCode = pBankCode;
        try
            {
            string exportResult;

            //pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);
            //reportNameIndv = pReportName.ToUpper().Replace("STATEMENT_CORP_COM_", "STATEMENT_CORP_IND_");
            vCurDate = pDate; // DateTime.Now.AddMonths(-1);
            strOutputPath = clsBasFile.getPathWithoutFile(pStrFileName) + vCurDate.ToString("yyyyMM") + pBankName;
            clsBasFile.createDirectory(strOutputPath);//pStrFileName + vCurDate.ToString("yyyyMM") + pBankName
            //pStrFileName = reportNameComp;// = clsBasFile.makeStrAsPath(pStrFileName);
            //strFileNameComp = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + "Company_" + vCurDate.ToString("yyyyMM"); // + ".pdf"
            strFileNameComp = strOutputPath + "\\" + pBankName + pStrFile + "Company_" + vCurDate.ToString("yyyyMM"); // + ".pdf"
            //strFileNameIndv = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + "Individual_" + vCurDate.ToString("yyyyMM"); // + ".pdf"
            strFileNameIndv = strOutputPath + "\\" + pBankName + pStrFile + "Individual_" + vCurDate.ToString("yyyyMM"); // + ".pdf"
            //strFileMD5 = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM"); // + ".pdf"
            strFileMD5 = strOutputPath + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM"); // + ".pdf"
            curBranchVal = pBankCode;
            MainTableCond = strWhereCond;//strWhereCond
            //supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and instr(x.contracttype,'" + strProductCond + "') > 0)";
            //FillStatementDataSet(pBankCode);
            if (createCorporateVal)
                {
                isCorporateVal = true;
                }
            FillStatementDataSet(pBankCode, "");

            //Company

            string[,] arrTables = new string[1, 2];
            //expReport.selectionFormula = "{@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0";
            //if (pBankCode == 7)
            //    expReport.selectionFormula = "{@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.CARDPRODUCT} in " + strProductCond;//" and {@f_closingbalance} <> 0";
            //else
            expReport.selectionFormula = "{@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.PACKAGENAME} = 'STTOTABLECORP'";//" and {@f_closingbalance} <> 0";
            //arrTables[0, 0] = "select * from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " t where t.branch = " + pBankCode;//7"
            //arrTables[0, 1] = "TSTATEMENTMASTERTABLE";//clsSessionValues.mainDbSchema + 
            //expReport.ExportSetup(clsBasFile.getPathWithoutFile(strFileNameComp), clsBasFile.getFileFromPath(strFileNameComp), reportNameComp, arrTables);//pReportName@"D:\pC#\exe\Reports" clsBasFile.getPathWithoutExtn(pStrFileName), strReportDbName
            expReport.ExportSetup(clsBasFile.getPathWithoutFile(strFileNameComp), clsBasFile.getFileFromPath(strFileNameComp), reportNameComp, DSstatement);
            expReport.ExportSelection(pCurrType);
            exportResult = expReport.ExportCompletion();

            //Individual or CardHolder
            expReport = new clsExportReport();
            //expReport.selectionFormula = "{@F_branch} = " + pBankCode + " and (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}))";
            //expReport.selectionFormula = "{@F_branch} = " + pBankCode + " and (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) or {@f_closingbalance} <> 0)";
            //expReport.selectionFormula = "{@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 or ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) ";
            //if (pBankCode == 50)
            //expReport.selectionFormula = "if not isnull({@f_BranchTrans}) then {@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) else {@F_branch} = " + pBankCode + ""; //and {@f_closingbalance} <> 0 
            //expReport.selectionFormula = "if not isnull({@f_BranchTrans}) then {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 " + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) else {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 " + ""; //and {@f_closingbalance} <> 0 
            //expReport.selectionFormula = "if not isnull({@f_BranchTrans}) then {@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) else {@F_branch} = " + pBankCode ; //and {@f_closingbalance} <> 0 
            //if (pBankCode == 7)
            //    expReport.selectionFormula = "(IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0) and {TSTATEMENTMASTERTABLE.CARDPRODUCT} in " + strProductCond + ") or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + ")";
            //else
            expReport.selectionFormula = "(IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.PACKAGENAME} = 'STTOTABLECORP') or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.PACKAGENAME} = 'STTOTABLECORP')";

            //else
            //expReport.selectionFormula = "{@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) "; //and {@f_closingbalance} <> 0 
            //string[,] arrTablesIndv = new string[2, 2];
            //arrTablesIndv[0, 0] = "select * from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " t where t.branch = " + pBankCode;//7"
            //arrTablesIndv[0, 1] =  "TSTATEMENTMASTERTABLE";//clsSessionValues.mainDbSchema +
            //arrTablesIndv[1, 0] = "select * from " + clsSessionValues.mainDbSchema + clsSessionValues.detailTable + " t where t.branch = " + pBankCode;//7"
            //arrTablesIndv[1, 1] = "tstatementdetailtable";
            //expReport.ExportSetup(clsBasFile.getPathWithoutFile(strFileNameIndv), clsBasFile.getFileFromPath(strFileNameIndv), reportNameIndv, arrTablesIndv);//@"D:\pC#\exe\Reports" clsBasFile.getPathWithoutExtn(pStrFileName), strReportDbName
            expReport.ExportSetup(clsBasFile.getPathWithoutFile(strFileNameIndv), clsBasFile.getFileFromPath(strFileNameIndv), reportNameIndv, DSstatement);
            expReport.ExportSelection(pCurrType);
            exportResult = expReport.ExportCompletion();

            //ArrayList aryLstFiles = new ArrayList();
            ////aryLstFiles.Add("");
            aryLstFiles.Add(strFileNameComp + expReport.getExportType(pCurrType));
            aryLstFiles.Add(strFileNameIndv + expReport.getExportType(pCurrType));

            //clsBasFile.generateFileMD5(aryLstFiles, strFileMD5 + ".MD5");
            //aryLstFiles.Add(strFileMD5 + ".MD5");
            //SharpZip zip = new SharpZip();
            //zip.createZip(aryLstFiles, strFileMD5 + ".zip", "");

            if (!(exportResult.IndexOf("Report Export Done Successfully.") > -1))
                {
                rtrnVal = false;
                }
            getCardProduct(pBankCode);
            clsCalcStatSummary calcStatSummary = new clsCalcStatSummary();
            //calcStatSummary.statType = "PDF";
            calcStatSummary.statRelation = StatementNoDRel;
            calcStatSummary.CreateCorporate = true;
            calcStatSummary.Statement(pStrFileName, pBankName, pBankCode
            , pStrFile, pDate, pStmntType, pAppendData, DSstatement, DSProducts);
            //if (pStmntType != "No")
            //{
            //  curBranchVal = pBankCode;
            //  FillStatementDataSet(pBankCode);
            //  fillStatementHistory(pStmntType, pAppendData);
            //}
            }
        catch (OracleException ex)
            {
            clsDbOracleLayer.catchError(ex);
            }
        catch (NotSupportedException ex)  //(Exception ex)  //
            {
            clsBasErrors.catchError(ex);
            }
        finally
            {
            }
        return rtrnVal;
        }



    public void CreateZip()
        {
        clsBasFile.generateFileMD5(aryLstFiles, strFileMD5 + ".MD5");
        aryLstFiles.Add(strFileMD5 + ".MD5");
        SharpZip zip = new SharpZip();
        zip.createZip(aryLstFiles, strFileMD5 + ".zip", "");
        }


    public string reportCompany
        {
        get { return reportNameComp; }
        set { reportNameComp = value; }
        }// reportCompany

    public string reportIndividual
        {
        get { return reportNameIndv; }
        set { reportNameIndv = value; }
        }// reportIndividual

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

    public string productCond
        {
        get { return strProductCond; }
        set { strProductCond = value; }
        }  // productCond

    public bool CreateCorporate
        {
        get { return createCorporateVal; }
        set { createCorporateVal = value; }
        }// CreateCorporate

    ~clsStatement_CommonCorpExp()
        {
        }
    }

