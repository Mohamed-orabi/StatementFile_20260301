using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Collections;
using System.Data.SqlClient;
//using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;

// Branch X
public class clsStatementMrchXml
    {
    private string strBankName;
    private FileStream fileStrmBasic, fileStrmTrans;//, fileStrmSubtotal
    private string strFileBasic, strFileTrans;//, strFileSubtotal
    private StreamWriter streamWritBasic, streamWritTrans;//, streamWritSubtotal
    private DataRow masterRow;
    private DataRow detailRow;
    private const int MaxDetailInPage = 20; //
    private const int linesInLastPage = 67; //
    private int CurPageRec4Dtl = 0;
    private int pageNo = 0, totalCardPages = 0 //, totalPages=0
      , totalAccPages = 0, totCardRows = 0, totAccRows = 0, totAccCards = 0;
    //	private string lastPageTotal ;
    private string curCardNo, CardNumber = String.Empty, curCardNumber = String.Empty, PrevCardNumber = String.Empty;// 
    private string curAccountNo, prevAccountNo = String.Empty;//
    private string strAccountFooter;
    private int intAccountFooter;
    private decimal totNetUsage = 0;
    private decimal totAccountValue = 0;
    private DataRow[] cardsRows, accountRows;
    private string CrDbDetail;
    private const string strFileSpr = "','";//#
    private bool isPrimaryOnly;
    private string curMainCard;
    private int curAccRows = 0;
    private int totCrdNoInAcc, curCrdNoInAcc;
    private string stmNo;
    private DataRow[] mainRows;

    private string extAccNum;
    private string strOutputPath, strOutputFile, fileSummaryName;
    private DateTime vCurDate;
    //private DataSet DSstatementRaw;
    //private DataRelation StatementNoDRels;

    private OleDbConnection conn;
    private string strSqlActn;

    private string StrStatLable = string.Empty, strWhereCond = string.Empty;
    private ArrayList aryLstFiles = new ArrayList();
    private string strFileName;

    private string strStatementPath;
    private string strDestDataFile;
    private string curBranchVal;
    private string curBankName, curBankFullName;
    private DataSet DSstatement = new DataSet();
    private int cntMaster, cntDetail;
    private DateTime curStatDate = DateTime.Now;
    private string xmlSourceFile;

    private bool isHaveEstatVal = false;

    public clsStatementMrchXml()
        {
        }

    public string Statement(string pStrFileName, string pBankFullName, string pBankName, string pBankCode, DateTime pDate)
        {
        string rtrnStr = "Successfully Generate " + pBankName;
        bool preExit = true;
        curBankFullName = pBankFullName;
        curBankName = pBankName;
        curStatDate = pDate;
        xmlSourceFile = pStrFileName;

        strStatementPath = clsBasFile.getPathWithoutFile(pStrFileName) + pBankName + "_MerchantStatement_" + pDate.ToString("yyyyMMdd_HHmmss");
        clsBasFile.createDirectory(strStatementPath);
        strDestDataFile = strStatementPath + "\\" + pBankName + "_MerchantStatement_" + pDate.ToString("yyyyMMdd_HHmmss") + ".mdb";
        clsBasFile.moveFile(xmlSourceFile, clsBasFile.getPathWithoutExtn(strDestDataFile) + ".XML");
        pStrFileName = xmlSourceFile = clsBasFile.getPathWithoutExtn(strDestDataFile) + ".XML";
        clsBasFile.copyFile(@"D:\pC#\ProjData\Statement\_Data\MerchantStatementTemplate.mdb", strDestDataFile);
        curBranchVal = pBankCode; // 4; //4  = real   1 = test
        //FillStatementDataSet(pBankCode); //DSstatement =  //6); // 6
        DSstatement.ReadXml(pStrFileName);
        if (DSstatement.Tables.Count > 1)
            {
            DataRelation StatementNoDRel = DSstatement.Relations.Add("StaementNoDR",
            DSstatement.Tables["Statement"].Columns["StatementNo"],
            DSstatement.Tables["Operation"].Columns["StatementNo"]);
            }

        Statement(pStrFileName, DSstatement);
        mergeTrans();
        DSstatement = null;

        OleDbConnection conn = new OleDbConnection(clsDbCon.sConMsAccess(strDestDataFile, ""));
        DSstatement = new DataSet("merchStatement");
        OleDbDataAdapter StatDataAdapter = new OleDbDataAdapter("SELECT * FROM Statement where right(Account,2) <> '-1'", conn);
        StatDataAdapter.Fill(DSstatement, "Statement");
        OleDbDataAdapter OperDataAdapter = new OleDbDataAdapter("SELECT * FROM Operation where D <> 'Reimbursemnt - Payment'", conn);
        OperDataAdapter.Fill(DSstatement, "Operation");
        conn.Close();

        //StatementBackup(pStrFileName, DSstatement);
        exportReport();

        if (isHaveEstatVal)
            {
            clsStatHtmlMrch statHtmlMrch = new clsStatHtmlMrch();
            statHtmlMrch.emailFromName = curBankFullName + " - Statement";
            statHtmlMrch.emailFrom = "cardservices@emp-group.com";
            statHtmlMrch.bankWebLink = "www.emp-group.com";
            statHtmlMrch.bankLogo = @"D:\pC#\ProjData\Statement\EMP\logo.gif";
            statHtmlMrch.visaLogo = @"D:\pC#\ProjData\Statement\EMP\VisaLogo.gif";
            //statHtmlMrch.setFrm = this;
            statHtmlMrch.waitPeriod = 10000;
            statHtmlMrch.Statement(pStrFileName, DSstatement, curBankName, curBankFullName, curStatDate);
            statHtmlMrch = null;
            //aryLstFiles.Add(@clsBasFile.getPathWithoutExtn(pStrFileName) + "_MerchantEmails.txt");
            //aryLstFiles.Add(@clsBasFile.getPathWithoutExtn(pStrFileName) + "_MerchantNoEmails.txt");
            //aryLstFiles.Add(@clsBasFile.getPathWithoutExtn(pStrFileName) + "_SummaryEmail.txt");
            }
        sendEmail2Bank(curBankName);

        return "";
        }

    public string Statement(string pStrFileName, DataSet pDSstatement)
        {
        string sqlStr = string.Empty, fieldsName = string.Empty;
        try
            {
            conn = new OleDbConnection(clsDbCon.sConMsAccess(strDestDataFile, ""));
            conn.Open();

            foreach (DataRow mRow in DSstatement.Tables["Statement"].Rows)
                {
                //printHeader();//>>
                cntMaster++;
                sqlStr = "insert into Statement ";
                sqlStr += " (StatMasterCode,Branch,BankName,BankFullName,StatDate,";
                foreach (DataColumn myColumn in DSstatement.Tables["Statement"].Columns)
                    {
                    sqlStr += "[" + myColumn.ColumnName + "],";
                    }
                sqlStr = sqlStr.Substring(0, sqlStr.Length - 1);
                sqlStr += ") VALUES (" + cntMaster + "," + curBranchVal + ",'" + curBankName + "','" + curBankFullName + "',#" + curStatDate.ToString("dd/MM/yyyy HH:mm:ss") + "#,";

                foreach (DataColumn myColumn in DSstatement.Tables["Statement"].Columns)
                    {
                    if (myColumn.ColumnName.ToUpper() == "EndDate".ToUpper() || myColumn.ColumnName == "StartDate".ToUpper())
                        sqlStr += "#" + mRow[myColumn] + "#,";//sqlStr += "TO_DATE('" + string.Format("{0,19:dd/MM/yyyy HH:mm:ss}", mRow[myColumn]) + "','DD/MM/YYYY HH24:MI:SS'),";
                    else
                        sqlStr += "'" + basText.replaceSpecialChar(mRow[myColumn].ToString()) + "',";
                    }
                sqlStr = sqlStr.Substring(0, sqlStr.Length - 1);
                sqlStr += ")";

                (new OleDbCommand(sqlStr, conn)).ExecuteNonQuery();
                cardsRows = mRow.GetChildRows("StaementNoDR");

                foreach (DataRow dRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
                    {
                    cntDetail++;
                    sqlStr = "insert into Operation ";
                    sqlStr += " (StatDetailCode,StatMasterCode,Branch,";
                    foreach (DataColumn myColumn in DSstatement.Tables["Operation"].Columns)
                        {
                        sqlStr += "[" + myColumn.ColumnName + "],";
                        }
                    sqlStr = sqlStr.Substring(0, sqlStr.Length - 1);
                    sqlStr += ") VALUES (" + cntDetail + "," + cntMaster + "," + curBranchVal + ",";

                    foreach (DataColumn myColumn in DSstatement.Tables["Operation"].Columns)
                        {
                        if (myColumn.ColumnName.ToUpper() == "OD" || myColumn.ColumnName.ToUpper() == "TD")
                            sqlStr += (string.IsNullOrEmpty(dRow[myColumn].ToString().Trim()) ? "null," : "#" + dRow[myColumn] + "#,");//"#" + dRow[myColumn] + "#,"   sqlStr += "TO_DATE('" + string.Format("{0,19:dd/MM/yyyy HH:mm:ss}", mRow[myColumn]) + "','DD/MM/YYYY HH24:MI:SS'),";
                        else if (myColumn.ColumnName.ToUpper() == "O" || myColumn.ColumnName.ToUpper() == "A" || myColumn.ColumnName.ToUpper() == "OA" || myColumn.ColumnName.ToUpper() == "CF" || myColumn.ColumnName.ToUpper() == "S")
                            sqlStr += "'" + (string.IsNullOrEmpty(dRow[myColumn].ToString().Trim()) ? "0" : dRow[myColumn].ToString()) + "',";
                        else
                            sqlStr += "'" + basText.replaceSpecialChar(dRow[myColumn].ToString()) + "',";
                        }
                    sqlStr = sqlStr.Substring(0, sqlStr.Length - 1);
                    sqlStr += ")";

                    (new OleDbCommand(sqlStr, conn)).ExecuteNonQuery();
                    }//dRow
                }//mRow
            (new OleDbCommand("update Statement set ExternalAccount = Account where ExternalAccount is null or ExternalAccount = ''", conn)).ExecuteNonQuery();
            (new OleDbCommand("update Operation set TD = OD where TD is null", conn)).ExecuteNonQuery();
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
            conn.Close();

            //ArrayList aryLstFiles = new ArrayList();
            //aryLstFiles.Add("");
            //>>aryLstFiles.Add(pStrFileName);
            //aryLstFiles.Add(@strFileSubtotal);

            //clsBasFile.generateFileMD5(aryLstFiles, @clsBasFile.getPathWithoutExtn(pStrFileName)+ ".MD5"); //+ "_Raw" 
            //aryLstFiles.Add(@clsBasFile.getPathWithoutExtn(pStrFileName)+ ".MD5"); //+ "_Raw" 
            //SharpZip zip = new SharpZip();
            //zip.createZip(aryLstFiles, @clsBasFile.getPathWithoutExtn(pStrFileName)+ ".zip", ""); //+ "_Raw" 

            //printFileMD5();
            }
        return "";

        }


    public string StatementBackup(string pStrFileName, DataSet pDSstatement)
        {
        SqlConnection connSql;

        string sqlStr = string.Empty, fieldsName = string.Empty;
        try
            {
            connSql = new SqlConnection(clsDbCon.sConSqlServer("MMOHAMED", "Statement", "apach", "apach4ever"));
            connSql.Open();

            foreach (DataRow mRow in DSstatement.Tables["Statement"].Rows)
                {
                //printHeader();//>>
                cntMaster++;
                sqlStr = "insert into mrchStatement ";
                sqlStr += " (StatMasterCode,Branch,BankName,BankFullName,StatDate,";
                foreach (DataColumn myColumn in DSstatement.Tables["Statement"].Columns)
                    {
                    sqlStr += myColumn.ColumnName + ",";
                    }
                sqlStr = sqlStr.Substring(0, sqlStr.Length - 1);
                sqlStr += ") VALUES (" + cntMaster + "," + curBranchVal + ",'" + curBankName + "','" + curBankFullName + "',#" + curStatDate.ToString("dd/MM/yyyy HH:mm:ss") + "#,";

                foreach (DataColumn myColumn in DSstatement.Tables["Statement"].Columns)
                    {
                    if (myColumn.ColumnName.ToUpper() == "EndDate".ToUpper() || myColumn.ColumnName == "StartDate".ToUpper())
                        sqlStr += "#" + mRow[myColumn] + "#,";//sqlStr += "TO_DATE('" + string.Format("{0,19:dd/MM/yyyy HH:mm:ss}", mRow[myColumn]) + "','DD/MM/YYYY HH24:MI:SS'),";
                    else
                        sqlStr += "'" + basText.replaceSpecialChar(mRow[myColumn].ToString()) + "',";
                    }
                sqlStr = sqlStr.Substring(0, sqlStr.Length - 1);
                sqlStr += ")";

                (new SqlCommand(sqlStr, connSql)).ExecuteNonQuery();
                cardsRows = mRow.GetChildRows("StaementNoDR");

                foreach (DataRow dRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
                    {
                    cntDetail++;
                    sqlStr = "insert into mrchOperation ";
                    sqlStr += " (StatDetailCode,StatMasterCode,Branch,";
                    foreach (DataColumn myColumn in DSstatement.Tables["Operation"].Columns)
                        {
                        sqlStr += myColumn.ColumnName + ",";
                        }
                    sqlStr = sqlStr.Substring(0, sqlStr.Length - 1);
                    sqlStr += ") VALUES (" + cntDetail + "," + cntMaster + "," + curBranchVal + ",";

                    foreach (DataColumn myColumn in DSstatement.Tables["Operation"].Columns)
                        {
                        if (myColumn.ColumnName.ToUpper() == "OD" || myColumn.ColumnName.ToUpper() == "TD")
                            sqlStr += (string.IsNullOrEmpty(dRow[myColumn].ToString().Trim()) ? "null," : "#" + dRow[myColumn] + "#,");//"#" + dRow[myColumn] + "#,"   sqlStr += "TO_DATE('" + string.Format("{0,19:dd/MM/yyyy HH:mm:ss}", mRow[myColumn]) + "','DD/MM/YYYY HH24:MI:SS'),";
                        else if (myColumn.ColumnName.ToUpper() == "O" || myColumn.ColumnName.ToUpper() == "A" || myColumn.ColumnName.ToUpper() == "OA" || myColumn.ColumnName.ToUpper() == "CF" || myColumn.ColumnName.ToUpper() == "S")
                            sqlStr += "'" + (string.IsNullOrEmpty(dRow[myColumn].ToString().Trim()) ? "0" : dRow[myColumn].ToString()) + "',";
                        else
                            sqlStr += "'" + basText.replaceSpecialChar(dRow[myColumn].ToString()) + "',";
                        }
                    sqlStr = sqlStr.Substring(0, sqlStr.Length - 1);
                    sqlStr += ")";

                    (new SqlCommand(sqlStr, connSql)).ExecuteNonQuery();
                    }//dRow
                }//mRow
            (new SqlCommand("update mrchStatement set ExternalAccount = Account where ExternalAccount is null or ExternalAccount = ''", connSql)).ExecuteNonQuery();
            (new SqlCommand("update mrchOperation set TD = OD where TD is null", connSql)).ExecuteNonQuery();

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
            conn.Close();

            //ArrayList aryLstFiles = new ArrayList();
            //aryLstFiles.Add("");
            aryLstFiles.Add(pStrFileName);
            //aryLstFiles.Add(@strFileSubtotal);

            //clsBasFile.generateFileMD5(aryLstFiles, @clsBasFile.getPathWithoutExtn(pStrFileName)+ ".MD5"); //+ "_Raw" 
            //aryLstFiles.Add(@clsBasFile.getPathWithoutExtn(pStrFileName)+ ".MD5"); //+ "_Raw" 
            //SharpZip zip = new SharpZip();
            //zip.createZip(aryLstFiles, @clsBasFile.getPathWithoutExtn(pStrFileName)+ ".zip", ""); //+ "_Raw" 

            //printFileMD5();
            }
        return "";

        }


    public void mergeTrans()//int pBranch
        {
        string masterQuery;
        string upateSql = string.Empty;
        string docNo = string.Empty;
        string mainRowId = string.Empty, supRowId = string.Empty;
        decimal totTrans = 0;
        long SqlCnt = 0;
        string strSqlActn = string.Empty;
        string curRec = string.Empty;

        Cursor.Current = Cursors.WaitCursor;
        masterQuery = "SELECT StatDetailCode, A, OA, D, [NO],DOCNO FROM Operation order by DOCNO,CF,[NO]";// WHERE APPROVAL<>''
        try
            {
            //DataSet ds;// = null
            OleDbConnection conn = new OleDbConnection(clsDbCon.sConMsAccess(strDestDataFile, ""));
            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(masterQuery, conn);
            conn.Open();

            upateSql = "update Operation set Amount2Pay = A ";
            (new OleDbCommand(upateSql, conn)).ExecuteNonQuery();

            DataSet ds = new DataSet("merTrans");
            myDataAdapter.Fill(ds, "merTrans");

            //strSqlActn = "begin ";
            foreach (DataRow supRow in ds.Tables["merTrans"].Rows) //DataRow    ds.Tables["tStatementMasterTable"].Rows    mRow.GetChildRows(StatementNoDRel)
                {
                if (docNo == supRow["DOCNO"].ToString() && (supRow["D"].ToString().ToLower().IndexOf("commision") > 0 || supRow["D"].ToString().ToLower().IndexOf("commission") > 0))
                    {
                    //supRowId = supRow["StatDetailCode"].ToString();
                    //totTrans = totTrans + Convert.ToDecimal(supRow[dBilltranamount].ToString());
                    upateSql = "update Operation set Commission = " + supRow["A"].ToString() + ",Amount2Pay = A + " + supRow["A"].ToString() + " where StatDetailCode = " + curRec;
                    (new OleDbCommand(upateSql, conn)).ExecuteNonQuery();
                    upateSql = "delete from Operation where StatDetailCode = " + supRow["StatDetailCode"].ToString();
                    (new OleDbCommand(upateSql, conn)).ExecuteNonQuery();
                    //clsDbOracleLayer.doAction(upateSql);
                    //strSqlActn += upateSql + ";";
                    //SqlCnt++;
                    }
                //else if (docNo != supRow[dDocno].ToString())
                //{
                //  if (docNo != "")
                //  {
                //    upateSql = "update a4m.tstatementdetailtable set billtranamount = " + totTrans + " where rowid = '" + mainRowId + "'";
                //    //clsDbOracleLayer.doAction(upateSql);
                //    strSqlActn += upateSql + ";";
                //    SqlCnt++;
                //  }
                //  totTrans = 0; supRowId = "";
                //  mainRowId = supRow["StatDetailCode"].ToString();
                //  totTrans = Convert.ToDecimal(supRow[dBilltranamount].ToString());
                //}
                docNo = supRow["DOCNO"].ToString();
                curRec = supRow["StatDetailCode"].ToString();
                //if (SqlCnt > MaxNoTrans)
                //{
                //  strSqlActn = strSqlActn + " end;";
                //  (new OracleCommand(strSqlActn, conn)).ExecuteNonQuery();
                //  strSqlActn = "begin "; SqlCnt = 0;
                //}
                } // end of supRow
            //if (supRowId != "")
            //{
            //  upateSql = "update a4m.tstatementdetailtable set billtranamount = " + totTrans + " where rowid = '" + mainRowId + "'";
            //  //clsDbOracleLayer.doAction(upateSql);
            //  strSqlActn += upateSql + ";";
            //  SqlCnt++;
            //}
            //if (SqlCnt > 0)
            //{
            //  strSqlActn = strSqlActn + " end;";
            //  (new OracleCommand(strSqlActn, conn)).ExecuteNonQuery();
            //}

            conn.Close();
            }
        //    catch
        //       {
        //           //clsDbOracleLayer.catchError(ex);
        // //    }
        catch (OracleException ex)
            {
            clsDbOracleLayer.catchError(ex);
            }
        //    catch (NotSupportedException ex)  //(Exception ex)  //
        //    {
        //      //clsBasErrors.catchError(ex);
        //    }
        finally
            {
            Cursor.Current = Cursors.Default;
            }
        }

    private string exportReport()
        {
        string exportResult;
        clsExportReport expReport = new clsExportReport();
        //expReport.StatLable = StrStatLable;
        expReport.ExportSetupMerchant(strStatementPath + "\\", clsBasFile.getFileWithoutExtn(strDestDataFile), Application.StartupPath + @"\Reports\MerchantStatement_" + curBankName + ".rpt", DSstatement);//@"D:\pC#\exe\Reports" clsBasFile.getPathWithoutExtn(pStrFileName), strReportDbName
        expReport.ExportSelection(ExportFormatType.PortableDocFormat);
        exportResult = expReport.ExportCompletion();
        expReport = null;
        return exportResult;
        }

    public void CreateZip()
        {
        clsBasFile.generateFileMD5(aryLstFiles, @clsBasFile.getPathWithoutExtn(strFileName) + ".MD5"); //+ "_Raw" 
        aryLstFiles.Add(@clsBasFile.getPathWithoutExtn(strFileName) + ".MD5"); //+ "_Raw" 
        SharpZip zip = new SharpZip();
        zip.createZip(aryLstFiles, @clsBasFile.getPathWithoutExtn(strFileName) + ".zip", ""); //+ "_Raw" 
        }


    public bool sendEmail2Bank(string pBank)
        {
        bool rtrnVal = true;
        string strMessage = string.Empty;
        //ArrayList aryLstFiles = new ArrayList();
        aryLstFiles.Add(strStatementPath + "\\" + clsBasFile.getFileWithoutExtn(strDestDataFile) + ".pdf");
        clsBasFile.generateFileMD5(aryLstFiles, strStatementPath + "\\" + clsBasFile.getFileWithoutExtn(strDestDataFile) + ".MD5");
        SharpZip zip = new SharpZip();
        zip.createZip(aryLstFiles, strStatementPath + "\\" + clsBasFile.getFileWithoutExtn(strDestDataFile) + ".zip", "");

        //if (!clsFtpCommandLine.sendFile2ftpSilent(strStatementPath + "\\" + clsBasFile.getFileWithoutExtn(strDestDataFile) + ".zip", "Opsdev/_Operation/"))//_Settlement
        if (!clsFtpCommandLine.sendFile2ftpSilent(strStatementPath + "\\" + clsBasFile.getFileWithoutExtn(strDestDataFile) + ".zip", "Live-Banks/" + pBank + "/To"))//_Settlement
            return false;

        clsEmail sndMail = new clsEmail();
        ArrayList pLstCC = new ArrayList(), pLstBCC = new ArrayList();
        ArrayList pLstTo = new ArrayList(), pLstAttachedFile = new ArrayList();

        strMessage = "Dear ";

        switch (pBank)
            {
            case "CPA":
                pLstTo.Add("merchantsupport@emp-group.com");
                pLstCC.Add("merchantsupport@emp-group.com");
                //        pLstCC.Add("mmohammed@emp-group.com");
                strMessage += "Merchant Support Team,";
                break;

            case "BDK":
                pLstTo.Add("merchantsupport@emp-group.com");
                pLstCC.Add("merchantsupport@emp-group.com");
                //        pLstCC.Add("mmohammed@emp-group.com");
                strMessage += "Merchant Support Team,";
                break;

            case "SSB":
                pLstTo.Add("Abena.A.Asare-Menako@socgen.com");
                //pLstTo.Add("Theophilus.Asamoah-Sakyi@socgen.com");
                //pLstTo.Add("Irene.Adjei-Bisa@socgen.com");//"mmohammed@emp-group.com"
                pLstCC.Add("Lance.Parker@socgen.com");
                //pLstCC.Add("Kwame-Kumi.Awuku@socgen.com");
                //pLstCC.Add("CardList@socgen.com");
                //pLstCC.Add("Bright.Ameme@socgen.com");
                pLstCC.Add("merchantsupport@emp-group.com");
                //       pLstCC.Add("mmohammed@emp-group.com");
                //strMessage += "Irene,";
                //strMessage += "Theophilus,";
                strMessage += "Abena ";
                break;

            case "GTBG":
                pLstTo.Add("sheikh.chaw@gtbank.com");
                pLstCC.Add("merchantsupport@emp-group.com");
                //        pLstCC.Add("mmohammed@emp-group.com");
                strMessage += "Sheikh,";
                break;

            case "ICBG":
                pLstTo.Add("Ackah-Nyamike.Juliet@ghana.accessbankplc.com");
                pLstCC.Add("merchantsupport@emp-group.com");
                //        pLstCC.Add("mmohammed@emp-group.com");
                strMessage += "Juliet,";
                break;

            default:
                pLstTo.Add("merchantsupport@emp-group.com");
                //        pLstCC.Add("terminals@emp-group.com");
                strMessage += "Merchant Support Team,";
                break;

            } // end switch

        //pLstCC.Add("mabouleila@emp-group.com");
        pLstCC.Add("Statement@emp-group.com");
        pLstBCC.Add("mhafez@emp-group.com");
        //pLstBCC.Add("afattah@emp-group.com");
        pLstBCC.Add("gelfeky@emp-group.com");
        pLstBCC.Add("esami@emp-group.com");
        //pLstBCC.Add("asalman@emp-group.com");
        //pLstTo.Add("Kwesi.Acquah@socgen.com");//"mmohammed@emp-group.com"
        strMessage += "\r\n\tKindly Check FTP for Merchant Statement File - " + clsBasFile.getFileWithoutExtn(strDestDataFile) + ".zip";
        strMessage += "\r\n" + "According to PCI-DSS security rules all card numbers in statement files will be masked";
        strMessage += "\r\n\r\n\t\t\t\t\t\tBest Regards";
        strMessage += "\r\nStatement Program\r\n+ 02 33331400\r\nstatement@emp-group.com";


        sndMail.sendEmailHTML("Statement@emp-group.com", pLstTo, pLstCC, pLstBCC, pLstAttachedFile, pBank + " Merchant Statement " + curStatDate.ToString("yyyy/MM/dd"), strMessage, clsCnfg.readSetting("SmtpServer"), false);

        sndMail = null;
        return rtrnVal;
        }


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

    public string bankName
        {
        get { return strBankName; }
        set { strBankName = value; }
        }// bankName

    public bool isHaveEstat
        {
        set { isHaveEstatVal = value; }
        }

    ~clsStatementMrchXml()
        {
        DSstatement.Dispose();
        }
    }
