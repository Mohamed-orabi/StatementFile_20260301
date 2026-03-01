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

public class clsStatement_Export : clsBasStatement
{
  private DateTime vCurDate;
  private string strOutputPath, strFileName;
  private string StrStatLable = string.Empty, strWhereCond = string.Empty;
  private ArrayList aryLstFiles = new ArrayList();
  private string selFormula = string.Empty;
  private string strCloseBalance = string.Empty;

  public clsStatement_Export()
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
    bool rtrnVal = true ;
    string strCloseBalanceFrml = string.Empty, strCloseBalanceSql = string.Empty;
    //(new OracleCommand("alter session set optimizer_goal = RULE", new OracleConnection(clsDbCon.sConOracle))).ExecuteNonQuery();
    clsExportReport expReport = new clsExportReport();
    expReport.bankCode = pBankCode;
    //expReport.selectionFormula = "{@F_branch} = " + pBankCode + " and (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}))";
    //expReport.selectionFormula = "{@F_branch} = " + pBankCode + " and (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) or {@f_closingbalance} <> 0)";
    //expReport.selectionFormula = "{@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 or ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) ";
    //expReport.selectionFormula = "{@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) ";
    if (strCloseBalance == "ALL")
    {
      strCloseBalanceFrml = "" ;
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

    if(selFormula == string.Empty)
      expReport.selectionFormula = "{@F_branch} = " + pBankCode + strCloseBalanceFrml + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) ";
    //expReport.selectionFormula = "{@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) ";//and {@f_closingbalance} <> 0 
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

      string[,] arrTables = new string[2,2];
      arrTables[0, 0] = "select /*+ index (" + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " iBranchTstatementmastertable) */ * from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " where branch = " + pBankCode + strWhereCond;//+ " and accountno in('Credit_AOA_000002090') "
      arrTables[0, 1] =  "TSTATEMENTMASTERTABLE";//clsSessionValues.mainDbSchema +clsSessionValues.mainDbSchema + 
      arrTables[1, 0] = "select /*+ index (" + clsSessionValues.mainDbSchema + clsSessionValues.detailTable + " iBranchTstatementdetailtable) */ * from " + clsSessionValues.mainDbSchema + clsSessionValues.detailTable + " where branch = " + pBankCode + " and not (postingdate is null and docno is null)";//+ " and accountno in('Credit_AOA_000002090') " //and (postingdate != '' and docno  != '')
      arrTables[1, 1] = "tstatementdetailtable";//clsSessionValues.mainDbSchema + 
      
//      expReport.ExportSetup(@"D:\TEMP\P20Files\Statement\200709BAI\", "BAI_Statement_File_200709", pReportName, arrTables);//@"D:\pC#\exe\Reports" clsBasFile.getPathWithoutExtn(pStrFileName), strReportDbName
      expReport.StatLable = StrStatLable;
      expReport.ExportSetup(clsBasFile.getPathWithoutFile(strFileName), clsBasFile.getFileFromPath(strFileName), pReportName, arrTables);//@"D:\pC#\exe\Reports" clsBasFile.getPathWithoutExtn(pStrFileName), strReportDbName
      expReport.ExportSelection(pCurrType);
      exportResult = expReport.ExportCompletion();

      //ArrayList aryLstFiles = new ArrayList();
      ////aryLstFiles.Add("");
      aryLstFiles.Add(strFileName + expReport.getExportType(pCurrType));

      //clsBasFile.generateFileMD5(aryLstFiles, pStrFileName + ".MD5");
      //aryLstFiles.Add(pStrFileName + ".MD5");
      //SharpZip zip = new SharpZip();
      //zip.createZip(aryLstFiles, pStrFileName + ".zip", "");

      if (!(exportResult.IndexOf("Report Export Done Successfully.") > -1))
      {
        rtrnVal = false;
      }
      getCardProduct(pBankCode);
      clsCalcStatSummary calcStatSummary = new clsCalcStatSummary();
      //calcStatSummary.statType = "PDF";
      calcStatSummary.statRelation = StatementNoDRel;
      calcStatSummary.Statement(pStrFileName, pBankName, pBankCode
      , pStrFile, pDate, pStmntType, pAppendData, DSstatement,DSProducts);
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

  public bool SplitByBranch(string pStrFileName, string pBankName
  , int pBankCode, string pStrFile, string pReportName, ExportFormatType pCurrType, DateTime pDate, string pStmntType, bool pAppendData)
  {
    bool rtrnVal = true;
    string queryString, whereCond;
    selFormula = "{@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) ";
    queryString = "select distinct t.cardbranchpart, t.cardbranchpartname from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " t where t.branch = " + pBankCode + strWhereCond;
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
        strWhereCond = whereCond + " and cardbranchpart = " + rdr.GetInt32(0);
        export(pStrFileName, pBankName, pBankCode, pStrFile + rdr.GetString(1) + "_", pReportName, pCurrType, pDate, pStmntType, pAppendData);
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
    finally
    {
    }
    return rtrnVal;
  }

  public bool SplitByDbCr(string pStrFileName, string pBankName
, int pBankCode, string pStrFile, string pReportName, ExportFormatType pCurrType, DateTime pDate, string pStmntType, bool pAppendData)
  {
    string orgStrWhere;
    orgStrWhere = strWhereCond;
    selFormula = "{@F_branch} = " + pBankCode + " and {@f_closingbalance} = 0 and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) ";
    strWhereCond += " and closingbalance = 0";
    export(pStrFileName, pBankName, pBankCode, pStrFile + "Balance Equal Zero_", pReportName, pCurrType, pDate, pStmntType, pAppendData);
    
    strWhereCond = orgStrWhere;
    selFormula = "{@F_branch} = " + pBankCode + " and {@f_closingbalance} > 0 and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) ";
    strWhereCond += " and closingbalance > 0";
    export(pStrFileName, pBankName, pBankCode, pStrFile + "Balance Greater Zero_", pReportName, pCurrType, pDate, pStmntType, pAppendData);
    
    strWhereCond = orgStrWhere;
    selFormula = "{@F_branch} = " + pBankCode + " and {@f_closingbalance} < 0 and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) ";
    strWhereCond += " and closingbalance < 0";
    export(pStrFileName, pBankName, pBankCode, pStrFile + "Balance Less Zero_", pReportName, pCurrType, pDate, pStmntType, pAppendData);

    selFormula = "{@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) ";//and {@f_closingbalance} <> 0 
    strWhereCond = orgStrWhere;
    strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
    strFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM"); // + ".pdf"
    return true;
  }

  public void CreateZip()
  {
    clsBasFile.generateFileMD5(aryLstFiles, strFileName + ".MD5");
    aryLstFiles.Add(strFileName + ".MD5");
    SharpZip zip = new SharpZip();
    zip.createZip(aryLstFiles, strFileName + ".zip", "");
  }


  public string closeBanace // " <> 0"
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


  ~clsStatement_Export()
  {
  }
}

