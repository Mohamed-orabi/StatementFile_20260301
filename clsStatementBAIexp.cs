using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Data;
using System.Data.OracleClient;
using System.Xml;
using System.Collections;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;

// Branch 7
public class clsStatementBAIexp : clsBasStatement
{
  DateTime vCurDate, pStrFileName;
  string strOutputPath;
  public clsStatementBAIexp()
  {
  }

  public bool export(string pStrFileName, string pBankName
    , int pBankCode, string pStrFile, string pReportName, ExportFormatType pCurrType, DateTime pCurDate, string pStmntType, bool pAppendData)
  {
    bool rtrnVal = true ;
    clsMaintainData maintainData = new clsMaintainData();
    maintainData.matchCardBranch4Account(pBankCode);

    clsExportReport expReport = new clsExportReport();
    expReport.bankCode = pBankCode;
    //expReport.selectionFormula = "{@F_branch} = " + pBankCode + " and (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}))";
    //expReport.selectionFormula = "{@F_branch} = " + pBankCode + " and (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) or {@f_closingbalance} <> 0)";
    //expReport.selectionFormula = "{@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 or ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) ";
    expReport.selectionFormula = "{@F_branch} = " + pBankCode + " and ({@f_BranchTrans} = " + pBankCode + " and not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO})) ";//and {@f_closingbalance} <> 0 
    try
    {
      string exportResult;

      pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);
      vCurDate = pCurDate; //DateTime.Now.AddMonths(-1);
      strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
      clsBasFile.createDirectory(pStrFileName + vCurDate.ToString("yyyyMM") + pBankName);
      pStrFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM"); // + ".pdf"

      string[,] arrTables = new string[2,2];
      arrTables[0, 0] = "select * from TSTATEMENTMASTERTABLE t where t.branch = " + pBankCode;//7"
      arrTables[0,1] = "TSTATEMENTMASTERTABLE";
      arrTables[1, 0] = "select * from tstatementdetailtable t where t.branch = " + pBankCode;//7"
      arrTables[1,1] = "tstatementdetailtable";

      
//      expReport.ExportSetup(@"C:\TEMP\P20Files\Statement\200709BAI\", "BAI_Statement_File_200709", pReportName, arrTables);//@"D:\pC#\exe\Reports" clsBasFile.getPathWithoutExtn(pStrFileName), strReportDbName
      expReport.ExportSetup(clsBasFile.getPathWithoutFile(pStrFileName), clsBasFile.getFileFromPath(pStrFileName), pReportName, arrTables);//@"D:\pC#\exe\Reports" clsBasFile.getPathWithoutExtn(pStrFileName), strReportDbName
      expReport.ExportSelection(pCurrType);
      exportResult = expReport.ExportCompletion();

      ArrayList aryLstFiles = new ArrayList();
      //aryLstFiles.Add("");
      aryLstFiles.Add(pStrFileName + expReport.getExportType(pCurrType));
      clsBasFile.generateFileMD5(aryLstFiles, pStrFileName + ".MD5");
      aryLstFiles.Add(pStrFileName + ".MD5");
      SharpZip zip = new SharpZip();
      zip.createZip(aryLstFiles, pStrFileName + ".zip", "");
      if (!(exportResult.IndexOf("Report Export Done Successfully.") > -1))
      {
        rtrnVal = false;
      }
      if (pStmntType != "No")
      {
        curBranchVal = pBankCode;
        FillStatementDataSet(pBankCode);
        fillStatementHistory(pStmntType,pAppendData);
      }
      //fillStatementHistory(pStmntType,pAppendData);
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

 


  ~clsStatementBAIexp()
  {
  }
}

