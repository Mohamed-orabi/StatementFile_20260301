using System;
using System.Data;
using System.IO;
using System.Text;
using Oracle.DataAccess.Client;
using System.Collections;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Security;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;

public class clsStatHtmlBDCA : clsBasStatement
{
    private string strBankName;
    private FileStream fileSummary, fileEmails, fileNoEmails;
    private StreamWriter streamSummary, streamEmails, streamNoEmails;
    private DataRow masterRow;
    private DataRow detailRow;
    private const int MaxDetailInPage = 20;
    private const int MaxDetailInLastPage = 27;
    private int CurPageRec4Dtl = 0;
    private int pageNo = 0, totalPages = 0, totalCardPages = 0
        , TotalAccountPage = 0, TotalAccountRows = 0, curAccRows = 0;//
    private string curCardNo;
    private string curAccountNo, prevAccountNo = String.Empty;
    private decimal totNetUsage = 0;
    private DataRow[] cardsRows, AccountRowsFromDetailsTable;
    private DataRow[] mainRows;
    private string CurrentPageFlag;
    private string CardNumber, PrimaryCardNumber;
    private string stmtfilename, stmtpass, stmtfilenameold, stmtnumber;
    private int totNoOfCardStat, totNoOfPageStat;
    private bool isHaveF3 = true;
    private FileStream fileStrmErr;
    private StreamWriter strmWriteErr;
    private string curMainCard;

    private string extAccNum;
    private int totCrdNoInAcc, curCrdNoInAcc;
    private string strOutputPath, strOutputFile, fileSummaryName;
    private DateTime vCurDate;
    private int totPages;
    string endOfCustomer = string.Empty;
    string curFileName = string.Empty;
    private ArrayList aryLstFiles, aryLstFilesStmt;

    private StringBuilder emailStr = new StringBuilder("");
    private string strBankLogo, strBackGround;
    private string strEmailFrom = "bdc.estatment@bdc.com.eg";
    private string strbankWebLink = "https://bit.ly/33peXqd";
    private string emailTo, curAccountNumber, curCardNumber, curClientID, curCustomerName;//, curEmail
    private string strBirthYear, strNationalID;
    private string strEmailFileName, strNoEmailFileName;
    private frmStatementFile frmMain;
    private int totRec = 1, numMail, numOfTry = 0;
    private ArrayList pLstCC = new ArrayList(), pLstBCC = new ArrayList();
    private string logoAlignmentStr = "left";//center
    private string strbackGround = @"D:\pC#\exe\FilesForPrograms\frmBackground.jpg";
    private bool isRewardVal = false;
    private bool createCorporateVal = false;
    private string accountNoName = mAccountno;
    private string rewardCondVal = "'New Reward Contract'";//'Reward Contract'
    private string strVisaLogo = string.Empty;
    private clsEmail sndMail;
    private string emailFromNameStr;
    private int waitPeriodVal = 7000;
    private clsValidateEmail valdEmail = new clsValidateEmail();
    private int noOfEmails, noOfBadEmails, noOfWithoutEmails;
    private string strEmailFromTmp, emailLabelTmp;  //"cardservices@zenithbank.com"
    private DataRow emailRow;
    private string strFileNam, stmntType;
    private string prvEmail = string.Empty;
    protected string ClientID = mClientid;

    private string statMessageFileMonthlyVal = string.Empty;
    private string statMessageMonthly = "&nbsp;";//Null
    private bool HasAttachementVal = false;
    private string attachedFilesStr;
    private string strProductCond = string.Empty;
    private bool IsSplittedVal = false;
    private string strbottomBanner = string.Empty;
    private string strbottomBanner1 = string.Empty;
    private bool HasSenderVal = false;
    private string pdfsPath, StrFile;

    public clsStatHtmlBDCA()
    {
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="pStrFileName"></param>
    /// <param name="pBankName"></param>
    /// <param name="pBankCode"></param>
    /// <param name="pStrFile"></param>
    /// <param name="pCurDate"></param>
    /// <param name="pStmntType"></param>
    /// <param name="pAppendData"></param>
    /// <param name="pReportName"></param>
    /// <returns></returns>
    public string Statement(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pCurDate, string pStmntType,
        bool pAppendData, string pReportName)
    {
        string rtrnStr = "Successfully Generate " + pBankName;
        int curMonth = pCurDate.Month;
        StrFile = pStrFile;
        //bool preExit = true;
        strFileNam = pStrFileName;
        stmntType = pStmntType;
        Dictionary<string, string> FileOfLogs = new Dictionary<string, string>();

        curMonth = pCurDate.Month;
        aryLstFiles = new ArrayList();
        aryLstFilesStmt = new ArrayList();
        try
        {

            #region PreparingForAllFile
            pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);
            vCurDate = pCurDate; //DateTime.Now.AddMonths(-1);
            strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "_" + pStmntType;
            //clsBasFile.deleteDirectory(strOutputPath);
            clsBasFile.createDirectory(strOutputPath);
            clsBasFile.createDirectory(strOutputPath + "//STMT");
            #region Create Logs Path
            string txtLogsPath = strOutputPath + "\\Logs";
            List<string> LogsOfCustomer = null;
            //string[] fileNames = new string[0]; 
            pdfsPath = strOutputPath + "\\STMT";
            List<string> sentPdfs = null;
            if (!Directory.Exists(txtLogsPath))
            {
                Directory.CreateDirectory(txtLogsPath);
            }
            else
            {
                #region Read Last file and add pdfFiles toZip

                LogsOfCustomer = new DirectoryInfo(txtLogsPath).GetFiles("*.txt", SearchOption.TopDirectoryOnly).Select(filename => filename.Name.Replace(".txt", string.Empty)).ToList();
                //fileNames = Directory.GetFiles(txtLogsPath, "*.txt", SearchOption.TopDirectoryOnly);

                foreach (var item in LogsOfCustomer)
                {
                    FileOfLogs[item] = item;
                }

                if (Directory.Exists(pdfsPath))
                {
                    List<string> Pdf = new DirectoryInfo(pdfsPath).GetFiles("*.pdf", SearchOption.TopDirectoryOnly).Select(filename => pdfsPath + "\\" + filename).ToList();

                    foreach (var item in Pdf)
                    {
                        aryLstFilesStmt.Add(item);
                    }

                    //get the sent PDFs and add it to the no of emails
                    sentPdfs = new DirectoryInfo(pdfsPath).GetFiles("*.pdf", SearchOption.TopDirectoryOnly).Select(filename => filename.Name.Replace(".pdf", string.Empty)).ToList();

                    noOfEmails += sentPdfs.Count;
                    noOfWithoutEmails += LogsOfCustomer.Count - sentPdfs.Count;
                    //noOfWithoutEmailsTemp += sentaccounts.Count - sentPdfs.Count;
                }

                #endregion
            }
            #endregion

            strEmailFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "_" + pStmntType + "\\" + pBankName + "_" + pStmntType + "_" + vCurDate.ToString("yyyyMM");//+ ".txt"
            strNoEmailFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "_" + pStmntType + "\\" + pBankName + "_" + "WithoutEmails" + "_" + vCurDate.ToString("yyyyMM");//+ ".txt"
            pStrFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "_" + pStmntType + "\\" + pBankName + "_" + pStmntType + "_" + vCurDate.ToString("yyyyMM") + ".txt";
            strBankName = pBankName;
            strOutputFile = pStrFileName;

            // open emails file
            FileStream errorslogfile = new FileStream(strOutputPath + "\\ErrorsLogFile" + ".txt", FileMode.Create); //Create
            fileEmails = new FileStream(strEmailFileName + ".txt", FileMode.Create); //Create
            streamEmails = new StreamWriter(fileEmails, Encoding.Default);
            //streamEmails.WriteLine("AccountNumber" + "|" + "CardNumber" + "|" + "ClientID" + "|" + "Email");
            streamEmails.WriteLine("AccountNumber" + "|" + "ClientID" + "|" + "Email" + "|" + "Date Time");
            streamEmails.AutoFlush = true;

            // open No emails file
            fileNoEmails = new FileStream(strNoEmailFileName + ".txt", FileMode.Create); //Create
            streamNoEmails = new StreamWriter(fileNoEmails, Encoding.Default);
            streamNoEmails.WriteLine("CustomerName" + "|" + "AccountNumber" + "|" + "ClientID" + "|" + "Email" + "|");
            streamNoEmails.AutoFlush = true;

            attachedFilesStr = clsBasFile.getPathWithoutFile(strEmailFileName);

            // open output file
            //>fileStrm = new FileStream(pStrFileName, FileMode.Create); //Create
            //>streamWrit = new StreamWriter(fileStrm, Encoding.Default);
            //Error file
            fileStrmErr = new FileStream(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName), FileMode.Create);
            strmWriteErr = new StreamWriter(fileStrmErr, Encoding.Default);
            streamNoEmails.WriteLine("CustomerName" + "|" + "ClientID" + "|" + "err" + "|");
            streamNoEmails.AutoFlush = true;

            if (!string.IsNullOrEmpty(statMessageFileMonthlyVal))
            {
                FileStream filReadM = null;
                StreamReader filStreamM = null;
                filReadM = new FileStream(statMessageFileMonthlyVal, FileMode.Open);
                filStreamM = new StreamReader(filReadM, Encoding.Default);
                statMessageMonthly = filStreamM.ReadToEnd();
                filStreamM.Close();
                filReadM.Close();
            }


            // open Summary file
            fileSummaryName = pStrFileName;
            fileSummaryName = clsBasFile.getPathWithoutExtn(fileSummaryName) +
              "_Summary." + clsBasFile.getFileExtn(fileSummaryName);
            fileSummary = new FileStream(fileSummaryName, FileMode.Create); //Create
            streamSummary = new StreamWriter(fileSummary, Encoding.Default);
            #endregion


            curBranchVal = pBankCode; // 10; //3 = real   1 = test


            MainTableCond = " m.contracttype in " + strProductCond + "";//strWhereCond
            supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + ")";
            strOrder = " m.accountno desc,m.cardprimary desc,cardcreationdate "; // BDCA-2550

            FillStatementDataSet_SortCardPriority(pBankCode, "");
            getClientWithEmail(pBankCode);
            getClientPasNoAndBirthYear(pBankCode);
            pageNo = 1; totalCardPages = 0;
            curCardNo = String.Empty;
            curAccountNo = String.Empty;
            frmMain.BeginInvoke(frmMain.setMinMaxProgressDelegate, new object[] { DSstatement.Tables["tStatementMasterTable"].Rows.Count });
            foreach (DataRow mRow in DSstatement.Tables["tStatementMasterTable"].Rows)
            {
                frmMain.BeginInvoke(frmMain.setProgressDelegate, new object[] { totRec++ });

                if (FileOfLogs.ContainsKey(mRow[accountNoName].ToString().Trim()))
                    continue;


                #region ignore this account
                //if (LogsOfCustomer != null && LogsOfCustomer.Contains(mRow[accountNoName].ToString().Trim()))
                //{
                //    continue;
                //}
                #endregion
                masterRow = mRow;

                curAccountNumber = masterRow[accountNoName].ToString();
                curCardNumber = PrimaryCardNumber;
                curClientID = masterRow[ClientID].ToString();
                curCustomerName = masterRow[mCustomername].ToString();

                CardNumber = masterRow[mCardno].ToString().Trim();

                if (CardNumber.Length != 16)
                {
                    continue;// Exclude Zero Length Cards 
                }
                PrimaryCardNumber = CardNumber;
                if (masterRow[mCardprimary].ToString() == "N")
                {
                    PrimaryCardNumber = masterRow[mPrinarycardno].ToString();
                }


                //start new account
                if (prevAccountNo != masterRow[accountNoName].ToString())//MC_TIT_EGP_000027135
                {
                    if (!string.IsNullOrEmpty(emailStr.ToString()) && !string.IsNullOrEmpty(emailTo))
                    {
                        SendEmail(emailStr.ToString(), "", emailTo);
                    }
                    if (!string.IsNullOrEmpty(emailStr.ToString()) && string.IsNullOrEmpty(emailTo))
                    {
                        streamNoEmails.WriteLine(curCustomerName + "|" + curAccountNumber + "|" + curClientID + "|" + " " + "|" + "|Without Email");
                        noOfWithoutEmails++;
                    }

                    emailStr = new StringBuilder("");

                    curMainCard = string.Empty;
                    isHaveF3 = false;

                    pageNo = 1;
                    CurPageRec4Dtl = 0; totNetUsage = 0; CurrentPageFlag = "F 1"; //if page is based on account no
                    calcAccountRows();
                    curFileName = "BDC Statement " + vCurDate.ToString("MMMM yyyy ") + masterRow[mCardclientname].ToString() + "_" + masterRow[mAccountno].ToString();
                    curFileName = curFileName.Replace('.', ' ').Replace('/', ' ').Replace(' ', '_');


                    if (TotalAccountRows < 1
                      && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0)
                    {
                        isHaveF3 = true;
                        continue;
                    }
                    //cardsRows = mRow.GetChildRows(StatementNoDRel); 
                    cardsRows = DSstatement.Tables["tStatementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[accountNoName]).ToString().Trim() + "'");
                    emailStr = new StringBuilder("");

                    prevAccountNo = masterRow[accountNoName].ToString();
                    //pageNo=1; //if page is based on account no
                    DataRow[] emailRows;
                    emailRows = DSOemails.Tables["Oemails"].Select("idclient = " + masterRow[ClientID].ToString());

                    //DSPassportNos
                    DataRow[] passportRows;
                    passportRows = DSBirthYandID.Tables["ClientPasNoAndBirthYear"].Select("idclient = " + masterRow[ClientID].ToString());

                    for (int i = 0; i < emailRows.Length; i++)
                    {
                        emailTo = emailRows[i][1].ToString().Trim();
                        emailRow = emailRows[i];
                    }
                    if (!string.IsNullOrEmpty(emailStr.ToString()) && string.IsNullOrEmpty(emailTo)) //emailStr != string.Empty && emailStr != null  emailTo != string.Empty && emailTo != null
                    {
                        streamNoEmails.WriteLine(curCustomerName + "|" + curAccountNumber + "|" + curClientID + "|" + " " + "|" + "|Without Email");
                        noOfWithoutEmails++;
                        stmtnumber = null;
                        continue;
                    }
                    else if (valdEmail.isValideEmail(emailTo) != "ValidEmail")//!basText.isValideEmail(pTo)
                    {
                        streamNoEmails.WriteLine(curCustomerName + "|" + curAccountNumber + "|" + curClientID + "|" + " " + "|" + "|" + clsCnfg.readSetting("strValidEmail") ?? "Bad Email " + emailTo);
                        noOfBadEmails++;
                        stmtnumber = null;
                        continue;
                    }
                    else if (!isValidateCard(masterRow[mCardstate].ToString())) //!basText.isValideEmail(pTo)
                    {
                        streamNoEmails.WriteLine(curCustomerName + "|" + curAccountNumber + "|" + curClientID + "|" + " " + "|" + "|invalid Card, Card State: " + masterRow["CARDSTATE"].ToString().Trim());
                        noOfBadEmails++;
                        stmtnumber = null;
                        continue;
                    }

                    if (passportRows.Length > 0)
                    {
                        for (int i = 0; i < passportRows.Length; i++)
                        {
                            strNationalID = passportRows[i][1].ToString().Trim();
                            strBirthYear = passportRows[i][2].ToString().Trim();
                            strNationalID = strNationalID.Substring(strNationalID.Length - 4);
                        }
                    }
                    else
                    {
                        strNationalID = "0000";
                        strBirthYear = "0000";
                    }

                    printBody();

                    totNoOfCardStat++;
                } // End of if(prevAccountNo != masterRow[accountNoName].ToString())
                else
                {
                    continue;
                }


                foreach (DataRow dRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
                {
                    detailRow = dRow;
                    if ((detailRow[dPostingdate] == DBNull.Value) && (detailRow[dDocno] == DBNull.Value)) continue;// Exclude On-Hold Transactions 
                    curAccRows++;
                    totNetUsage = calculateCrDb(totNetUsage, clsBasValid.validateNum(detailRow[dBilltranamount]), clsBasValid.validateStr(detailRow[dBilltranamountsign]));
                    CurPageRec4Dtl = CurPageRec4Dtl + 1;

                } //end of detail foreach

                curCrdNoInAcc++;

                CurPageRec4Dtl = 0; //>pageNo=1; if pages is based on account
                curAccRows = 0;
                strEmailFromTmp = emailLabelTmp = string.Empty;

                emailLabelTmp = "Banque Du Caire Statement (" + masterRow["CardProduct"].ToString().Trim() + ") For " + vCurDate.ToString("MMMM yyyy"); // + (strPrimaryCardNo != "" ? "_" + strPrimaryCardNo.Substring(strPrimaryCardNo.Length - 4) : "_");

                stmtnumber = masterRow[accountNoName].ToString();
                //stmtnumber = ExternalID;
                clsStatement_ExportRpt exportRptabp = new clsStatement_ExportRpt();
                if (createCorporateVal)
                    exportRptabp.whereCond = "cardaccountno = '" + stmtnumber + "'";
                else
                    exportRptabp.whereCond = "accountno = '" + stmtnumber + "'";

                exportRptabp.whereCondD = "accountno = '" + stmtnumber + "'";


                try
                {

                    exportRptabp.exportAttachmentBDCA(DSstatement, pStrFileName, pBankCode, curFileName, pReportName,

                        "( {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.ACCOUNTNO} = '" + stmtnumber + "' )"
                         + " or ( {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.ACCOUNTNO} = '" + stmtnumber + "' )"
                        );

                    //exportRptabp.exportAttachment(pStrFileName, pBankCode, curFileName, pReportName,
                    //"( {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.ACCOUNTNO} = '" + stmtnumber + "' )"
                    // + " or ( {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.ACCOUNTNO} = '" + stmtnumber + "' )"
                    //);


                    exportRptabp = null;
                    stmtnumber = null;
                    stmtfilenameold = @clsBasFile.getPathWithoutFile(pStrFileName) + @"\STMT\" + curFileName + ".pdf";
                    StringBuilder builder = new StringBuilder(stmtfilenameold);
                    builder.Replace(@"\\", @"\");
                    stmtfilenameold = builder.ToString();

                    if (frmStatementFile.Internal == true)
                    {
                        stmtpass = clsCnfg.readSetting("InternalPassword");
                    }
                    else
                    {
                        stmtpass = strNationalID.Trim() + strBirthYear.Trim();
                    }
                    //stmtpass = masterRow[mExternalno].ToString() != "" ? masterRow[mExternalno].ToString() : "";
                    // Get a fresh copy of the sample PDF file
                    string filenameSource = stmtfilenameold;
                    PdfDocument document = PdfReader.Open(filenameSource);
                    PdfSecuritySettings securitySettings = document.SecuritySettings;
                    securitySettings.UserPassword = stmtpass;
                    securitySettings.OwnerPassword = stmtpass;
                    securitySettings.PermitAccessibilityExtractContent = false;
                    securitySettings.PermitAnnotations = false;
                    securitySettings.PermitAssembleDocument = false;
                    securitySettings.PermitExtractContent = false;
                    securitySettings.PermitFormsFill = false;
                    securitySettings.PermitFullQualityPrint = false;
                    securitySettings.PermitModifyDocument = false;
                    securitySettings.PermitPrint = false;

                    // Save the document...
                    document.Save(filenameSource);


                }
                catch (Exception ex)
                {
                    strmWriteErr.WriteLine(curCustomerName + "|" + curClientID + "|" + " " + "|" + "|invalid Card, Card State: " + ex.Message);
                    strmWriteErr.WriteLine();
                    continue;
                }




                #region write sent account number
                string strTxt = "Account number: " + masterRow?[accountNoName]?.ToString() ?? "";
                strTxt += Environment.NewLine + "Accountno: " + masterRow?[mAccountno]?.ToString() ?? "";
                strTxt += Environment.NewLine + "mCardno: " + masterRow?[mCardno]?.ToString() ?? "";
                strTxt += Environment.NewLine + "mCardproduct: " + masterRow?[mCardproduct]?.ToString() ?? "";
                strTxt += Environment.NewLine + "strPrimaryCardNo: " + PrimaryCardNumber ?? "";
                strTxt += Environment.NewLine + "mExternalno: " + masterRow?[mExternalno]?.ToString() ?? "";
                strTxt += Environment.NewLine + "ClientID: " + masterRow?[ClientID]?.ToString() ?? "";
                strTxt += Environment.NewLine + "mCustomername: " + masterRow?[mCustomername]?.ToString() ?? "";

                File.WriteAllText(txtLogsPath + "//" + masterRow[accountNoName].ToString().Trim() + ".txt", strTxt);
                #endregion

            } //end of Master foreach
            //emailLabelTmp = emailLabel;
            if (!string.IsNullOrEmpty(emailStr.ToString()) && !string.IsNullOrEmpty(emailTo) && valdEmail.isValideEmail(emailTo) == "ValidEmail") //emailStr != string.Empty && emailStr != null  emailTo != string.Empty && emailTo != null

                SendEmail(emailStr.ToString(), "", emailTo);
            if (!string.IsNullOrEmpty(emailStr.ToString()) && string.IsNullOrEmpty(emailTo)) // emailStr != string.Empty && emailStr != null emailTo != string.Empty && emailTo != null
            {
                streamNoEmails.WriteLine(curCustomerName + "|" + curAccountNumber + "|" + curClientID + "|" + " " + "|" + "|Without Email");
                noOfWithoutEmails++;
                emailTo = strEmailFrom;
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

            printStatementSummary();

            // Close Summary File
            streamSummary.Flush();
            streamSummary.Close();
            fileSummary.Close();

            strmWriteErr.Flush();
            strmWriteErr.Close();
            fileStrmErr.Close();

            streamEmails.Flush();
            streamEmails.Close();
            fileEmails.Close();

            streamNoEmails.WriteLine(string.Empty);
            streamNoEmails.WriteLine(string.Empty);
            streamNoEmails.WriteLine("** Any Statement not have email will forward to email " + strEmailFrom);
            streamNoEmails.Flush();
            streamNoEmails.Close();
            fileNoEmails.Close();

            aryLstFiles.Add(strEmailFileName + ".txt");
            aryLstFiles.Add(strNoEmailFileName + ".txt");
            aryLstFiles.Add(fileSummaryName);
            clsBasFile.generateFileMD5(aryLstFiles, strEmailFileName + ".MD5");
            aryLstFiles.Add(strEmailFileName + ".MD5");
            SharpZip zip = new SharpZip();
            zip.createZip(aryLstFiles, @clsBasFile.getPathWithoutExtn(pStrFileName) + ".zip", "");
            //zip.createZip(aryLstFilesStmt, @clsBasFile.getPathWithoutFile(pStrFileName) + @"\" + pBankName + pStrFile + pStmntType + "_STMT" + vCurDate.ToString("yyyyMM") + ".zip", "");
            GeneratePdfSummaryFile();
            DSstatement.Dispose();
        }
        return rtrnStr;
    }



    protected void printBody()
    {
        totPages++; endOfCustomer = "";
        extAccNum = clsBasValid.validateStr(masterRow[mExternalno]);
        if (extAccNum.Trim() == "")
            extAccNum = clsBasValid.validateStr(masterRow[accountNoName]);

        //if(masterRow[mCardprimary].ToString() == "Y")
        if (pageNo == 1 && TotalAccountPage == 1) // statement contain 1 page
        {
            CurrentPageFlag = "F 0";
            isHaveF3 = true;
        }
        else if (pageNo == 1 && TotalAccountPage > 1)  //first page of multiple page statement
            CurrentPageFlag = "F 1"; // //middle page of multiple page statement
        else if (pageNo < TotalAccountPage)
            CurrentPageFlag = "F 2";
        else if (pageNo == TotalAccountPage) //last page of multiple page statement
        {
            CurrentPageFlag = "F 3";
            isHaveF3 = true;
            endOfCustomer = "/";
        }


        emailStr = new StringBuilder("");
        emailStr.Append(@"<html><head><meta http-equiv='Content-Language' content='en-us'><meta http-equiv='Content-Type' content='text/html; charset=windows-1252'>");
        emailStr.Append(@"<title>" + curFileName + @"</title>");
        emailStr.Append(@"</head><body>");

        emailStr.Append(@"<table id=""table15"" width=""100%"" border=""0"">");
        emailStr.Append(@"<tr><td>");
        emailStr.Append(@"<a href=" + strbankWebLink + @">");
        emailStr.Append(@"<img Style=""width: 100%;height: auto;"" src='cid:" + clsBasFile.getFileFromPath(strBackGround) + @"'/></a></tr></td><br/>");
        emailStr.Append(@"</table>");

        emailStr.Append(@"</body></html>");
        totalPages++;
        totNoOfPageStat++;
        //TBalance = decimal.Parse(masterRow[mOpeningbalance].ToString());
    }




    private void calcAccountRows()
    {
        curAccRows = 0;
        AccountRowsFromDetailsTable = null;

        AccountRowsFromDetailsTable = DSstatement.Tables["tStatementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[accountNoName]).ToString().Trim() + "'");
        TotalAccountPage = 0;
        TotalAccountRows = 0;
        string prevCardNo = String.Empty, CurCardNo = String.Empty;
        int currAccRowsPages = 0;
        foreach (DataRow dtAccRow in AccountRowsFromDetailsTable)
        {
            if ((dtAccRow[dPostingdate] == DBNull.Value) && (dtAccRow[dDocno] == DBNull.Value))
            {
                continue;
            }

            CurCardNo = dtAccRow[dCardno].ToString();

            if (CurCardNo.Trim().Length < 1)
                continue;

            currAccRowsPages++;
            TotalAccountRows++;

            if (currAccRowsPages > MaxDetailInPage)//==
            {
                currAccRowsPages = 1;//0
                TotalAccountPage++;
            }
            prevCardNo = dtAccRow[dCardno].ToString();
        }
        if (currAccRowsPages > 0)
            TotalAccountPage++;

        if (TotalAccountPage < 1)
            TotalAccountPage = 1;

        totCrdNoInAcc = curCrdNoInAcc = 0;
        mainRows = DSstatement.Tables["tStatementMasterTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[accountNoName]) + "'");
        curMainCard = CurCardNo = "";
        foreach (DataRow mainRow in mainRows) //mRow.GetChildRows(StatementNoDRel)
        {
            totCrdNoInAcc++;
            CurCardNo = mainRow[mCardno].ToString();
            if (mainRow[mCardprimary].ToString() == "Y")
                curMainCard = CurCardNo; //mainRow[mCardno].ToString();
            if (mainRow[mCardprimary].ToString() == "Y" && isValidateCard(mainRow[mCardstate].ToString()))
            {
                curMainCard = CurCardNo; //mainRow[mCardno].ToString();
                break;
            }
        }

        if (curMainCard == "")
            curMainCard = CurCardNo;
    }


    private void GeneratePdfSummaryFile()
    {
        SharpZip zip = new SharpZip();
        if (Directory.Exists(pdfsPath))
        {
            List<string> Pdf = new DirectoryInfo(pdfsPath).GetFiles("*.pdf", SearchOption.TopDirectoryOnly).Select(filename => pdfsPath + "\\" + filename).ToList();
            foreach (var item in Pdf)
            {
                aryLstFilesStmt.Add(item);

            }

            zip.createZip(aryLstFilesStmt, @clsBasFile.getPathWithoutFile(pdfsPath) + @"\" + bankName + StrFile + stmntType + "_STMT" + vCurDate.ToString("yyyyMM") + ".zip", "");
        }

    }

    private void printStatementSummary()
    {
        streamSummary.WriteLine(strBankName + " E-Statement Summary");
        streamSummary.WriteLine("__________________________");
        streamSummary.WriteLine("");
        streamSummary.WriteLine("Statements Sent by Email   " + noOfEmails.ToString());
        streamSummary.WriteLine("Statements have bad Email  " + noOfBadEmails.ToString());
        streamSummary.WriteLine("Statements without Email  " + noOfWithoutEmails.ToString());

        clsStatementSummary StatSummary = new clsStatementSummary();
        StatSummary.BankCode = curBranchVal;
        StatSummary.BankName = strBankName;
        StatSummary.StatementDate = vCurDate;
        StatSummary.CreationDate = DateTime.Now;
        StatSummary.StatementProduct = stmntType; // strFileNam + "__" + 
        StatSummary.StatementType = "Email";
        StatSummary.NoOfStatements = noOfEmails + noOfBadEmails;
        StatSummary.NoOfPages = noOfEmails + noOfBadEmails;
        StatSummary.NoOfTransactions = 0;
        //StatSummary.InsertRecord();
        StatSummary = null;

    }


    private void SendEmail(string pBody, string pSubject, string pTo)
    {
        try
        {

            if (valdEmail.isValideEmail(pTo) != "ValidEmail")//!basText.isValideEmail(pTo)
            {
                streamNoEmails.WriteLine(curCustomerName + "|" + curAccountNumber + "|" + curClientID + "|" + pTo + "|" + "|" + clsCnfg.readSetting("strValidEmail") ?? "Bad Email");
                noOfBadEmails++;
                return;
            }


            ArrayList pLstTo = new ArrayList(), pLstAttachedFile = new ArrayList(), pLstAttachedPic = new ArrayList();//, pLstCC = new ArrayList(), pLstBCC = new ArrayList();
            pSubject = emailLabelTmp; //emailLabel "BAI statement for 02/2008";

            if (pTo.EndsWith("."))//ehimolen@yahoo.com
                pTo = pTo + "com";
            pLstTo.Add(pTo);

            pLstAttachedPic.Add(strBankLogo);
            pLstAttachedPic.Add(strBackGround);

            if (HasAttachement && !string.IsNullOrWhiteSpace(stmtfilenameold))
            {
                pLstAttachedFile.Add(stmtfilenameold);
            }

            sndMail = new clsEmail();
            sndMail.emailFromName = emailFromNameStr;
            strEmailFromTmp = strEmailFrom;


            numOfTry = 0;
            noOfEmails++;

            if (numMail == 0)
            {
                pLstBCC.Add("statement@emp-group.com");
            }
            else
            {
                pLstBCC.Remove("statement@emp-group.com");
            }

            while (!sndMail.sendEmailAttachPicFile(strEmailFromTmp, pLstTo, pLstCC, pLstBCC, pLstAttachedPic, pLstAttachedFile, pSubject, pBody, clsCnfg.readSetting("SmtpServer")))
            {
                System.Threading.Thread.Sleep(waitPeriodVal);//2000
                numOfTry++;
                if (numOfTry > 100)
                {
                    streamEmails.Write("\t\t Error while Send Email");
                    streamNoEmails.WriteLine(curCustomerName + "|" + curAccountNumber + "|" + curClientID + "|" + " " + "|" + "|Exceed number of trials");
                    noOfBadEmails++; noOfEmails--;
                    break;
                }
            }
            streamEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));

            numMail++;
            if (numMail % 400 == 0)
            {
                System.Threading.Thread.Sleep(waitPeriodVal);//2000
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            if (HasAttachement)
            {
                //ERROR raised try adding new file with same name
                try
                {
                    //clsBasFile.moveFile(stmtfilenameold, stmtfilename);
                    if (!aryLstFilesStmt.Contains(stmtfilename))
                    {
                        aryLstFilesStmt.Add(stmtfilename);
                    }
                }
                catch (Exception ex)
                {
                    streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + pTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "| Email " + pTo + " , Err Message >> " + ex.Message);
                    noOfBadEmails++; noOfEmails--;
                    return;
                }
            }

        }
        catch (Exception ex) //(NotSupportedException ex) (Exception ex)  //
        {
            //clsBasErrors.catchError(ex);
            streamNoEmails.WriteLine(curCustomerName + "|" + curAccountNumber + "|" + curClientID + "|" + " " + "|" + "| Email " + pTo + " , Err Message >> " + ex.Message + ", Err Desc >> ");
            noOfBadEmails++;
        }
        finally
        {
            sndMail = null;
            emailStr = new StringBuilder(""); emailTo = string.Empty; curAccountNumber = string.Empty; curCardNumber = string.Empty; curClientID = string.Empty;
            wait2NextEmail(prvEmail, pTo, waitPeriodVal);//     System.Threading.Thread.Sleep(waitPeriodVal);//400
            prvEmail = pTo;

        }

    }



    protected virtual void wait2NextEmail(string pPrvEmail, string pNxtEmail, int pWaitPeriod)
    {
        if (pPrvEmail.IndexOf("@") > 0 && pNxtEmail.IndexOf("@") > 0)
            if (pPrvEmail.ToLower().Substring(pPrvEmail.IndexOf("@")) == pNxtEmail.ToLower().Substring(pNxtEmail.IndexOf("@")))
                System.Threading.Thread.Sleep(waitPeriodVal);//7000
            else
                System.Threading.Thread.Sleep(1000);
    }



    public int waitPeriod
    {
        get { return waitPeriodVal; }
        set { waitPeriodVal = value; }
    }//waitPeriod

    public bool isReward
    {
        get { return isRewardVal; }
        set { isRewardVal = value; }
    }

    public string rewardCond
    {
        get { return rewardCondVal; }
        set { rewardCondVal = value; }
    }  // rewardCond

    public bool CreateCorporate
    {
        get { return createCorporateVal; }
        set { createCorporateVal = value; }
    }// CreateCorporate

    public string backGround
    {
        get { return strbackGround; }
        set { strbackGround = value; }
    }// backGround

    public string bankName
    {
        get { return strBankName; }
        set { strBankName = value; }
    }// bankName

    public string emailFrom
    {
        get { return strEmailFrom; }
        set { strEmailFrom = value; }
    }// emailFrom

    public string bankWebLink
    {
        get { return strbankWebLink; }
        set { strbankWebLink = value; }
    }// bankWebLink

    public string bankLogo
    {
        get { return strBankLogo; }
        set { strBankLogo = value; }
    }// bankLogo



    public string BackGround
    {
        get { return strBackGround; }
        set { strBackGround = value; }
    }//BackGround

    //public frmStatementFile setFrm
    public frmStatementFile setFrm
    {
        set { frmMain = value; }
    }// setFrm

    public string emailCC
    {
        set
        {
            string strCC = value;
            string[] strAray;
            strAray = strCC.Split(';');
            foreach (string str in strAray)
                pLstCC.Add(str);
        }
    }// emailCC

    public string emailBCC
    {
        set
        {
            string strCC = value;
            string[] strAray;
            strAray = strCC.Split(';');
            foreach (string str in strAray)
                pLstBCC.Add(str);
        }
    }// emailBCC

    public string logoAlignment
    {
        set { logoAlignmentStr = value; }
    }// logoAlignment

    public string emailFromName
    {
        get { return emailFromNameStr; }
        set { emailFromNameStr = value; }
    }// emailFromName

    public string visaLogo
    {
        get { return strVisaLogo; }
        set { strVisaLogo = value; }
    }//visaLogo

    public string statMessageFileMonthly
    {
        get { return statMessageFileMonthlyVal; }
        set { statMessageFileMonthlyVal = value; }
    }//statMessageFileMonthly

    public string productCond
    {
        get { return strProductCond; }
        set { strProductCond = value; }
    }  // productCond

    public bool HasAttachement
    {
        get { return HasAttachementVal; }
        set { HasAttachementVal = value; }
    }  // HasAttachement

    public string attachedFiles
    {
        get { return attachedFilesStr; }
        set { attachedFilesStr = value; }
    }// attachedFiles

    public bool IsSplitted
    {
        get { return IsSplittedVal; }
        set { IsSplittedVal = value; }
    }  // IsSplitted

    public string bottomBanner
    {
        get { return strbottomBanner; }
        set { strbottomBanner = value; }
    }// bottomBanner

    public string bottomBanner1
    {
        get { return strbottomBanner1; }
        set { strbottomBanner1 = value; }
    }// bottomBanner1

    public bool HasSender
    {
        get { return HasSenderVal; }
        set { HasSenderVal = value; }
    }  // HasSender

    ~clsStatHtmlBDCA()
    {
        DSstatement.Dispose();
    }
}
