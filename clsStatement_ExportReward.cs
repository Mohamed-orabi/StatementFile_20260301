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

public class clsStatement_ExportReward : clsBasStatement
{
  private DateTime vCurDate, pStrFileName;
  private string strOutputPath, strFileName;
  private ArrayList aryLstFiles ;
  private string StrStatLable = string.Empty;

  public clsStatement_ExportReward()
  {
    aryLstFiles = new ArrayList();
  }

  public bool export(string pStrFileName, string pBankName
    , int pBankCode, string pStrFile, string pReportName, ExportFormatType pCurrType, DateTime pDate, string pStmntType, bool pAppendData)
  {
    bool rtrnVal = true ;
    //(new OracleCommand("alter session set optimizer_goal = RULE", new OracleConnection(clsDbCon.sConOracle))).ExecuteNonQuery();
    clsExportReport expReport = new clsExportReport();
    expReport.bankCode = pBankCode;
    //expReport.selectionFormula = "{@F_branch} = " + pBankCode + " and (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}))";
    //expReport.selectionFormula = "{@F_branch} = " + pBankCode + " and {@f_contracttype} <> 'Reward Program (Airmile)' and {@f_TRANDESCRIPTION} <> 'Calculated Points' and (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}))";
    //expReport.selectionFormula = "({@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {@f_contracttype} <> 'Reward Program (Airmile)') and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and {@f_TRANDESCRIPTION} <> 'Calculated Points') ";
    //expReport.selectionFormula = "({@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {@f_contracttype} <> 'Reward Program (Airmile)') or ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and {@f_TRANDESCRIPTION} <> 'Calculated Points') ";
    expReport.selectionFormula = "({@F_branch} = " + pBankCode + " and {@f_contracttype} <> 'Reward Program (Airmile)') and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and {@f_TRANDESCRIPTION} <> 'Calculated Points') "; // and {@f_closingbalance} <> 0
    try
    {
      string exportResult;

      clsMaintainData maintainData = new clsMaintainData();
      //maintainData.matchCardBranch4Account(pBankCode);
      //maintainData.fixArbicAddress(pBankCode);
      maintainData.fixArbicAddressLang(pBankCode);
      
      pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);
      vCurDate = pDate; // DateTime.Now.AddMonths(-1);
      strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
      clsBasFile.createDirectory(pStrFileName + vCurDate.ToString("yyyyMM") + pBankName);
      pStrFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM"); // + ".pdf"
      strFileName = pStrFileName;

      string[,] arrTables = new string[3,2];
      arrTables[0, 0] = "select /*+ index (" + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " iBranchTstatementmastertable) */ * from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " m where m.branch = " + pBankCode + " AND m.CONTRACTTYPE <> 'Reward Program (Airmile)'";
      arrTables[0,1] = "TSTATEMENTMASTERTABLE";//clsSessionValues.mainDbSchema +
      arrTables[1, 0] = "select /*+ index (" + clsSessionValues.mainDbSchema + clsSessionValues.detailTable + " iBranchTstatementdetailtable) */ * from tstatementdetailtable d where d.branch = " + pBankCode + " and d.TRANDESCRIPTION <> 'Calculated Points' AND d.POSTINGDATE IS NOT NULL AND d.DOCNO IS NOT NULL ";
      arrTables[1,1] = "tstatementdetailtable";//clsSessionValues.mainDbSchema + 
      arrTables[2, 0] = "select * from V_StatementMasterReward r where r.branch = " + pBankCode;
      arrTables[2, 1] = "V_StatementMasterReward";
      
//      expReport.ExportSetup(@"D:\TEMP\P20Files\Statement\200709BAI\", "BAI_Statement_File_200709", pReportName, arrTables);//@"D:\pC#\exe\Reports" clsBasFile.getPathWithoutExtn(pStrFileName), strReportDbName
      expReport.StatLable = StrStatLable;
      expReport.ExportSetup(clsBasFile.getPathWithoutFile(pStrFileName), clsBasFile.getFileFromPath(pStrFileName), pReportName, arrTables);//@"D:\pC#\exe\Reports" clsBasFile.getPathWithoutExtn(pStrFileName), strReportDbName
      expReport.ExportSelection(pCurrType);
      //expReport.selectionFormula = "{@F_branch} = " + pBankCode + " and {@f_contracttype} <> 'Reward Program (Airmile)' and {@f_TRANDESCRIPTION} <> 'Calculated Points' and (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}))";
      exportResult = expReport.ExportCompletion();

      //aryLstFiles.Add("");
      aryLstFiles.Add(pStrFileName + expReport.getExportType(pCurrType));

      if (!(exportResult.IndexOf("Report Export Done Successfully.") > -1))
      {
        rtrnVal = false;
      }

      if (pStmntType != "No")
      {
        curBranchVal = pBankCode;
        FillStatementDataSet(pBankCode);
        getCardProduct(pBankCode);
        //fillStatementHistory(pStmntType,pAppendData);
      }
      clsCalcStatSummary calcStatSummary = new clsCalcStatSummary();
      //calcStatSummary.statType = "PDF";
      calcStatSummary.statRelation = StatementNoDRel;
      calcStatSummary.Statement(pStrFileName, pBankName, pBankCode
      , pStrFile, pDate, pStmntType, pAppendData, DSstatement,DSProducts);
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

  public bool exportContactData(string pStrFileName, string pBankName
    , int pBankCode, string pStrFile, string pReportName, ExportFormatType pCurrType, DateTime pDate, string pStmntType, bool pAppendData)
  {
    bool rtrnVal = true;
    //(new OracleCommand("alter session set optimizer_goal = RULE", new OracleConnection(clsDbCon.sConOracle))).ExecuteNonQuery();
    clsExportReport expReport = new clsExportReport();
    expReport.bankCode = pBankCode;
    //expReport.selectionFormula = "{@F_branch} = " + pBankCode + " and (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}))";
    //expReport.selectionFormula = "{@F_branch} = " + pBankCode + " and {@f_contracttype} <> 'Reward Program (Airmile)' and {@f_TRANDESCRIPTION} <> 'Calculated Points' and (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}))";
    //expReport.selectionFormula = "({@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {@f_contracttype} <> 'Reward Program (Airmile)') and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and {@f_TRANDESCRIPTION} <> 'Calculated Points') and {@f_ClientBranch} = " + pBankCode ;
    //expReport.selectionFormula = "({@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {@f_contracttype} <> 'Reward Program (Airmile)') or ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and {@f_TRANDESCRIPTION} <> 'Calculated Points') ";
    expReport.selectionFormula = "({@F_branch} = " + pBankCode + " and {@f_contracttype} <> 'Reward Program (Airmile)') and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and {@f_TRANDESCRIPTION} <> 'Calculated Points') "; //and {@f_closingbalance} <> 0 
    try
    {
      string exportResult;

      clsMaintainData maintainData = new clsMaintainData();
      //maintainData.matchCardBranch4Account(pBankCode);
      maintainData.fixArbicAddress(pBankCode);

      pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);
      vCurDate = pDate; // DateTime.Now.AddMonths(-1);
      strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
      clsBasFile.createDirectory(pStrFileName + vCurDate.ToString("yyyyMM") + pBankName);
      pStrFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + "Contacts_Data_" + vCurDate.ToString("yyyyMM"); // + ".pdf"

      string[,] arrTables = new string[4, 2];
      arrTables[0, 0] = "select /*+ index (" + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " iBranchTstatementmastertable) */ * from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " m where m.branch = " + pBankCode + " AND m.CONTRACTTYPE <> 'Reward Program (Airmile)'";
      arrTables[0, 1] = clsSessionValues.mainDbSchema + clsSessionValues.mainTable + "";
      arrTables[1, 0] = "select /*+ index (" + clsSessionValues.mainDbSchema + clsSessionValues.detailTable + " iBranchTstatementdetailtable) */ * from " + clsSessionValues.mainDbSchema + clsSessionValues.detailTable + " d where d.branch = " + pBankCode + " and d.TRANDESCRIPTION <> 'Calculated Points' AND d.POSTINGDATE IS NOT NULL AND d.DOCNO IS NOT NULL ";
      arrTables[1, 1] = clsSessionValues.mainDbSchema + clsSessionValues.detailTable + "";
      arrTables[2, 0] = "select * from V_StatementMasterReward r where r.branch = " + pBankCode;
      arrTables[2, 1] = clsSessionValues.mainDbSchema + "V_StatementMasterReward";
      arrTables[3, 0] = "select /*+ index (" + MainSchema + "TCLIENTPERSONE UK_CLIENTPERSONE) */ * from " + clsSessionValues.mainDbSchema + "tClientPersone c where c.branch = " + pBankCode;
      //arrTables[3, 0] = "select * from tClientPersone c where c.branch = " + pBankCode;
      arrTables[3, 1] = "tClientPersone";

      //      expReport.ExportSetup(@"D:\TEMP\P20Files\Statement\200709BAI\", "BAI_Statement_File_200709", pReportName, arrTables);//@"D:\pC#\exe\Reports" clsBasFile.getPathWithoutExtn(pStrFileName), strReportDbName
      expReport.ExportSetup(clsBasFile.getPathWithoutFile(pStrFileName), clsBasFile.getFileFromPath(pStrFileName), pReportName, arrTables);//@"D:\pC#\exe\Reports" clsBasFile.getPathWithoutExtn(pStrFileName), strReportDbName
      expReport.ExportSelection(pCurrType);
      //expReport.selectionFormula = "{@F_branch} = " + pBankCode + " and {@f_contracttype} <> 'Reward Program (Airmile)' and {@f_TRANDESCRIPTION} <> 'Calculated Points' and (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}))";
      exportResult = expReport.ExportCompletion();

      //aryLstFiles.Add("");
      aryLstFiles.Add(pStrFileName + expReport.getExportType(pCurrType));

      if (!(exportResult.IndexOf("Report Export Done Successfully.") > -1))
      {
        rtrnVal = false;
      }

      if (pStmntType != "No")
      {
        curBranchVal = pBankCode;
        FillStatementDataSet(pBankCode);
        //fillStatementHistory(pStmntType, pAppendData);
      }
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


  public string StatLable
  {
    get { return StrStatLable; }
    set { StrStatLable = value; }
  }// StatLable

  ~clsStatement_ExportReward()
  {
  }
}

