using System;
using System.Collections.Generic;
using System.Text;
using Oracle.DataAccess.Client;
using System.Collections;
using System.Data;
using System.IO;


class clsStatementSummaryDetail
{
    private int vBankCode, vOpeningBalance, vClosingBalance, vMinimumAmount, vOverDueAmount, vOverDueDays, vInterset, vTotalDebits, vTotalCredits, vTotalPayements, vTotalCash, vTotalRetail, vTotalFees, vTotalOthers, vTotalPoints, vTotalRedeemedPoints, vTotalExpiredPoints, vCreditLimit;
    private string vAccountNo, vAccountCurrency, vExternalAccount;
    private DateTime vBilingDate, vGenerationDate, vDueDate;
    public static bool isUpdateDatble = true;

    private string STMTsummaryFileName;

    private frmStatementFile frmMain;

    public clsStatementSummaryDetail()
    {
        vBilingDate = vGenerationDate = vDueDate = DateTime.Now;
        vAccountNo = vExternalAccount = string.Empty;
        if (clsBasUserData.sType == "CR")
            STMTsummaryFileName = clsBasFile.Filepath + @"STMTSummaryDetail" + clsBasFile.TableName + ".txt";
        else if (clsBasUserData.sType == "DB")
            STMTsummaryFileName = clsBasFile.Filepath + @"STMTSummaryDetail" + clsBasFile.TableName + ".txt";
        else if (clsBasUserData.sType == "CP")
            STMTsummaryFileName = clsBasFile.Filepath + @"STMTSummaryDetail" + clsBasFile.TableName + ".txt";
    }

    //this is not used, you can use the InsertRecordDb in clsStatementSummary
    public void InsertRecordDb()
    {
        string strSqlActn;
        if (isUpdateDatble)
        {
            strSqlActn = "INSERT /*+ APPEND */ INTO a4m.MSCC_PROD_STAT_DETAIL values(";
            strSqlActn += vBankCode;
            strSqlActn += ",'" + vAccountNo + "'"; ;
            strSqlActn += ",'" + vAccountCurrency + "'"; ;
            strSqlActn += ",'" + vExternalAccount + "'";
            strSqlActn += ",to_date('" + String.Format("{0:dd/MM/yyyy hh:MM:ss}", vBilingDate) + "','dd/mm/yyyy HH:MI:SS')" ;
            strSqlActn += ",to_date('" + String.Format("{0:dd/MM/yyyy hh:MM:ss}", vGenerationDate) + "','dd/mm/yyyy HH:MI:SS')";
            strSqlActn += "," + vOpeningBalance;
            strSqlActn += "," + vClosingBalance;
            strSqlActn += "," + vMinimumAmount;
            strSqlActn += ",to_date('" + String.Format("{0:dd/MM/yyyy hh:MM:ss}", vDueDate) + "','dd/mm/yyyy HH:MI:SS')"  ;
            strSqlActn += "," + vOverDueAmount;
            strSqlActn += "," + vOverDueDays;
            strSqlActn += "," + vInterset;
            strSqlActn += "," + vTotalDebits;
            strSqlActn += "," + vTotalCredits;
            strSqlActn += "," + vTotalPayements;
            strSqlActn += "," + vTotalCash;
            strSqlActn += "," + vTotalRetail;
            strSqlActn += "," + vTotalFees;
            strSqlActn += "," + vTotalOthers;
            strSqlActn += "," + vTotalPoints;
            strSqlActn += "," + vTotalRedeemedPoints;
            strSqlActn += "," + vTotalExpiredPoints;
            strSqlActn += "," + vCreditLimit + ")"; ;
            clsBasFile.writeTxtFile(STMTsummaryFileName, Environment.NewLine + strSqlActn + ";", System.IO.FileMode.Append);
            clsBasFile.writeTxtFile(STMTsummaryFileName, Environment.NewLine + "Commit;", System.IO.FileMode.Append);
        }
        else
            return;
    }

    public int BankCode { set { vBankCode = value; } }
    public string AccountNo { set { vAccountNo = value; } }
    public string AccountCurrency { set { vAccountCurrency = value; } }
    public string ExternalAccount { set { vExternalAccount = value; } }
    public DateTime BilingDate { set { vBilingDate = value; } }
    public DateTime GenerationDate { set { vGenerationDate = value; } }
    public int OpeningBalance { set { vOpeningBalance = value; } }
    public int ClosingBalance { set { vClosingBalance = value; } }
    public int MinimumAmount { set { vMinimumAmount = value; } }
    public DateTime DueDate { set { vDueDate = value; } }
    public int OverDueAmount { set { vOverDueAmount = value; } }
    public int OverDueDays { set { vOverDueDays = value; } }
    public int Interset { set { vInterset = value; } }
    public int TotalDebits { set { vTotalDebits = value; } }
    public int TotalCredits { set { vTotalCredits = value; } }
    public int TotalPayements { set { vTotalPayements = value; } }
    public int TotalCash { set { vTotalCash = value; } }
    public int TotalRetail { set { vTotalRetail = value; } }
    public int TotalFees { set { vTotalFees = value; } }
    public int TotalOthers { set { vTotalOthers = value; } }
    public int TotalPoints { set { vTotalPoints = value; } }
    public int TotalRedeemedPoints { set { vTotalRedeemedPoints = value; } }
    public int TotalExpiredPoints { set { vTotalExpiredPoints = value; } }
    public int CreditLimit { set { vCreditLimit = value; } }
    public frmStatementFile setFrm { set { frmMain = value; } }
}

