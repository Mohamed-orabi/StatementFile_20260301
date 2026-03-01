using System;
using System.Collections.Generic;
using System.Text;
using Oracle.DataAccess.Client;
using System.Collections;
using System.Data;
using System.IO;
using System.Linq;

class clsStatementSummary
{
    private int vStatementCode, vProductCode, vBankCode, vNoOfStatements, vNoOfPages, vNoOfTransactions, vNoOfTransactionsInt;
    private string vBankName, vStatementProduct, vProductName, vProductType, vStatementType;
    private DateTime vStatementDate, vCreationDate;
    public static bool isUpdatedatble = true;
    public static bool isUpdatable = false;

    private ArrayList aryLstFiles = new ArrayList();
    List<string> listOFStrings = new List<string>();
    string OriginalBillingPath = @"D:\TEMP\P20Files\Statement\Billing\";
    string BillingPath="";
    string BillingPathWithoutDup = "";
    //private string summaryFileName = clsBasFile.Filepath+@"StatementSummary.txt";
    private string STMTsummaryFileName;

    private frmStatementFile frmMain;


    public clsStatementSummary()
    {
        vStatementDate = vCreationDate = DateTime.Now;
        vBankName = vStatementProduct = vStatementType = vProductType = string.Empty;
        if (clsBasUserData.sType == "CR")
            STMTsummaryFileName = clsBasFile.Filepath + @"STMTSummary" + clsBasFile.TableName + ".txt";
        else if (clsBasUserData.sType == "DB")
            STMTsummaryFileName = clsBasFile.Filepath + @"STMTSummary" + clsBasFile.TableName + ".txt";
        else if (clsBasUserData.sType == "CP")
            STMTsummaryFileName = clsBasFile.Filepath + @"STMTSummary" + clsBasFile.TableName + ".txt";
    }


    public void InsertRecordDb(string fileName)
    {
        if (fileName.Contains("UNB_Credit_txt"))
            return;

        BillingPath = OriginalBillingPath + "\\" + vStatementDate.ToString("yyyyMM") + "\\";
        if (!Directory.Exists(BillingPath))
        {
            Directory.CreateDirectory(BillingPath);
        }
        string filePath = BillingPath + fileName;

        //string strConn, strSqlActn;
        string strSqlActn;
        if (isUpdatedatble)
        {
            strSqlActn = "INSERT /*+ APPEND */ INTO a4m.MSCC_PROD_STAT_MASTER values(";
            strSqlActn += vBankCode;
            strSqlActn += "," + vProductCode + "";
            strSqlActn += ",'" + vProductName;
            //strSqlActn += "'," + "to_date('" + String.Format("{0:dd/MM/yyyy}", vStatementDate) + "','dd/mm/yyyy')";
            strSqlActn += "',null";
            strSqlActn += "," + vStatementDate.ToString("yyyyMM");
            strSqlActn += "," + vNoOfStatements;
            strSqlActn += "," + vNoOfTransactionsInt;
            strSqlActn += ",to_date('" + String.Format("{0:dd/MM/yyyy hh:MM:ss}", vCreationDate) + "','dd/mm/yyyy HH:MI:SS')" + ")";

            //if (File.Exists(filePath + ".txt"))
            if (File.Exists(filePath + ".txt"))
            {
                string[] lines = File.ReadAllLines(filePath + ".txt", Encoding.UTF8);
                foreach (var item in lines)
                {
                    if (string.IsNullOrWhiteSpace(item))
                        continue;

                    if (item[0] == 'C' || item[0] == 'c')
                        continue;

                    var match = listOFStrings.FirstOrDefault(stringToCheck => stringToCheck.Contains($"{vBankCode},{vProductCode},'{vProductName}'"));

                    if (match != null)
                    {
                        if (match.Contains($"{vBankCode},{vProductCode},'{vProductName}'") != item.Contains($"{vBankCode},{vProductCode},'{vProductName}'"))
                            continue;

                        var CurrentQuery = GetDateTimeFromQuery(match);
                        var IncompingQuery = GetDateTimeFromQuery(item);


                        if (CurrentQuery > IncompingQuery)
                        {
                            listOFStrings = listOFStrings.Select(x => x.Replace(item.ToString(), match.ToString())).ToList();
                            continue;
                        }
                        else
                        {
                            listOFStrings = listOFStrings.Select(x => x.Replace(match.ToString(), item.ToString())).ToList();
                            continue;
                        }
                    }

                    if (item.Contains($"{vBankCode},{vProductCode},'{vProductName}'"))
                    {
                        var match2 = listOFStrings.FirstOrDefault(stringToCheck => stringToCheck.Contains($"{vBankCode},{vProductCode},'{vProductName}'"));

                        if (match2 != null)
                        {
                            var CurrentQuery = GetDateTimeFromQuery(match2);
                            var IncompingQuery = GetDateTimeFromQuery(item);


                            if (IncompingQuery > CurrentQuery)
                            {
                                listOFStrings = listOFStrings.Select(x => x.Replace(match2.ToString(), item.ToString())).ToList();
                                continue;
                            }
                            continue;
                        }

                        var split = GetDateTimeFromQuery(item);
                        var split2 = GetDateTimeFromQuery(strSqlActn);



                        if (split2 > split)
                        {
                            listOFStrings.Add(strSqlActn);
                            listOFStrings.Add("Commit;");
                            continue;
                        }
                        else
                        {
                            listOFStrings.Add(item);
                            listOFStrings.Add("Commit;");
                            continue;
                        }
                    }
                    
                    listOFStrings.Add(strSqlActn);
                    listOFStrings.Add("Commit;");
                }
                if (!listOFStrings.Any())
                {
                    listOFStrings.Add(strSqlActn);
                    listOFStrings.Add("Commit;");
                }
                File.WriteAllLines(filePath + ".txt", listOFStrings);
            }
            else
            {
                listOFStrings.Add(strSqlActn);
                listOFStrings.Add("Commit;");
                clsBasFile.writeTxtFile(filePath + ".txt", Environment.NewLine + strSqlActn + ";", System.IO.FileMode.Append);
                clsBasFile.writeTxtFile(filePath + ".txt", Environment.NewLine + "Commit;", System.IO.FileMode.Append);
                clsDbOracleLayer.doAction(strSqlActn);
            }
        }
        else if (isUpdatable)
        {
            strSqlActn = "SELECT COUNT_STAT,COUNT_INT FROM MSCC_PROD_STAT_MASTER WHERE branch = " + vBankCode + " and product_code = " + vProductCode + " and stat_m = " + vStatementDate.ToString("yyyyMM") + "";
            DataSet ds = clsDbOracleLayer.getDataset(strSqlActn);
            strSqlActn = "UPDATE a4m.MSCC_PROD_STAT_MASTER SET COUNT_STAT=" + ds.Tables[0].Rows[0].ItemArray[0] + "+" + vNoOfStatements + ", COUNT_INT=" + ds.Tables[0].Rows[0].ItemArray[1] + "+" + vNoOfTransactionsInt + " where branch = " + vBankCode + " and product_code = " + vProductCode + " and stat_m = " + vStatementDate.ToString("yyyyMM") + "";
            //clsDbOracleLayer.doAction(strSqlActn);
            clsBasFile.writeTxtFile(STMTsummaryFileName, Environment.NewLine + strSqlActn + ";", System.IO.FileMode.Append);
            clsBasFile.writeTxtFile(STMTsummaryFileName, Environment.NewLine + "Commit;", System.IO.FileMode.Append);
        }
        else
            return;
    }

    private DateTime GetDateTimeFromQuery(string query)
    {
        return DateTime.ParseExact(query.Split(',')[7]
                                       .Remove(0, 8)
                                       .Trim('\''), "dd/MM/yyyy HH:mm:ss", null);
    }

    private bool IsSamePath(string fileName)
    {
        var getAllFiles = Directory
                        .GetFiles(@"D:\Temp\P20Files\Statement\Billing\" + $@"{vStatementDate.ToString("yyyyMM")}\", "*");
        string[] lastItem = null;
        string lastItem2 = "";


        foreach (var file in getAllFiles)
        {
            lastItem = file.Split('\\');
            lastItem2 = lastItem[lastItem.Length - 1].Split('_')[0];

            if (lastItem2 == fileName.Split('_')[0])
            {
                BillingPathWithoutDup = file;
                return true;
            }
        }
        return false;
    }
    public void UpdateRecordDb()
    {
        string strSqlActn;
        if (!isUpdatable)
            return;

        strSqlActn = "SELECT COUNT_STAT,COUNT_INT FROM MSCC_PROD_STAT_MASTER WHERE where branch = " + vBankCode + " and product_code = " + vProductCode + " and stat_m = " + vStatementDate.ToString("yyyyMM") + "";
        DataSet ds = clsDbOleDbLayer.getDataset(strSqlActn);
        strSqlActn = "UPDATE MSCC_PROD_STAT_MASTER SET COUNT_STAT=" + ds.Tables[0].Rows[0].ItemArray[0] + "+" + vNoOfStatements + " COUNT_INT=" + ds.Tables[0].Rows[0].ItemArray[1] + "+" + vNoOfTransactionsInt + " where branch = " + vBankCode + " and product_code = " + vProductCode + " and stat_m = " + vStatementDate.ToString("yyyyMM") + "";
        clsDbOracleLayer.doAction(strSqlActn);
    }

    public int StatementCode { set { vStatementCode = value; } }

    public int BankCode { set { vBankCode = value; } }
    public int ProductCode { set { vProductCode = value; } }
    public string ProductName { set { vProductName = value; } }
    public DateTime StatementDate { set { vStatementDate = value; } }
    public int NoOfStatements { set { vNoOfStatements = value; } }
    public int NoOfTransactionsInt { set { vNoOfTransactionsInt = value; } }
    public DateTime CreationDate { set { vCreationDate = value; } }

    public int NoOfPages { set { vNoOfPages = value; } }
    public int NoOfTransactions { set { vNoOfTransactions = value; } }
    public string BankName { set { vBankName = value; } }
    public string StatementProduct { set { vStatementProduct = value; } }
    public string StatementType { set { vStatementType = value; } }
    public frmStatementFile setFrm { set { frmMain = value; } }// setFrm



}

