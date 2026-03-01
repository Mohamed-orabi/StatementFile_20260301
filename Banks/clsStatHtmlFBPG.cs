
using System;
using System.Data;
using System.IO;
using System.Text;
using Oracle.DataAccess.Client;
using System.Collections;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Security;
using System.Globalization;
using System.Diagnostics;
using System.Linq;
using System.Collections.Generic;

public class clsStatHtmlFBPG : clsBasStatement
{
    private string strBankName;
    private FileStream fileStrm, fileSummary, fileEmails, fileNoEmails;
    private StreamWriter streamWrit, streamSummary, streamEmails, streamNoEmails;
    //private DataSet DSstatement;
    //private OracleDataReader drPrimaryCards, drMaster,drDetail;
    private DataRow masterRow;
    private DataRow detailRow;
    private string strEndOfLine = "\u000D";
    private string strEndOfPage = "\u000C";
    private const int MaxDetailInPage = 20;
    private const int MaxDetailInLastPage = 27;
    private int CurPageRec4Dtl = 0;
    private int pageNo = 0, totalPages = 0, totalCardPages = 0
        , totalAccPages = 0, totCardRows = 0, totAccRows = 0, curAccRows = 0;
    private string lastPageTotal;
    private string curCardNo;
    private string curAccountNo, prevAccountNo = String.Empty;
    private decimal totNetUsage = 0;
    private DataRow[] cardsRows, accountRows, rewardRows;
    private DataRow[] mainRows;
    private DataRow rewardRow;
    private string CurrentPageFlag;
    private string strCardNo, strPrimaryCardNo;
    private string strForeignCurr;
    private string stmNo, stmtNo, stmtfilename, stmtpass, stmtfilenameold, stmtnumber;
    private int totNoOfCardStat, totNoOfPageStat;

    private int curCardRow = 0, curTotCardRows = 0, numOfErr = 0, totNoOfTransactions = 0;
    private bool isPrimaryOnly, isHaveF3 = true, isPrimaryOpened = false;
    private FileStream fileStrmErr;
    private StreamWriter strmWriteErr;
    private string curMainCard;

    private string extAccNum;
    private int totCrdNoInAcc, curCrdNoInAcc;
    private string strOutputPath, strOutputFile, fileSummaryName;
    private DateTime vCurDate;
    private int totPages;
    string endOfCustomer = string.Empty;
    string cProduct = string.Empty, curFileNamePrivate, emailLabelcurFileName;
    private ArrayList aryLstFiles, aryLstFilesStmt;


    //private string emailStr;
    private StringBuilder emailStr = new StringBuilder("");
    private string emailLabel, strBankLogo, strBackGround;
    private string strEmailFrom = "cards@myfidelitybank.net";
    private string strbankWebLink = "www.fidelitybank.com.gh";
    private string emailTo, curAccountNumber, curCardNumber, curClientID, curCustomerName, ExternalID;
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
    private string accountLimit = mAccountlim;
    private string accountAvailableLimit = mAccountavailablelim;
    private string rewardCondVal = "'New Reward Contract'";//'Reward Contract'
    private string strVisaLogo = string.Empty;
    private clsEmail sndMail;
    private string emailFromNameStr;
    private int waitPeriodVal = 7000;
    private clsAes cryptAes = new clsAes();
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
    public string strbottomBanner = string.Empty;
    private string strbottomBanner1 = string.Empty;
    private decimal TBilltranamount = 0;
    private decimal TBalance = 0;
    private bool HasSenderVal = false;

    public clsStatHtmlFBPG()
    {
    }


    public string Statement(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pCurDate, string pStmntType, bool pAppendData, string pReportName)
    {
        EventLog.WriteEntry("STMT", "statement function begin", EventLogEntryType.Error);

        string rtrnStr = "Successfully Generate " + pBankName;
        int curMonth = pCurDate.Month;
        //bool preExit = true;
        strFileNam = pStrFileName;
        stmntType = pStmntType;

        curMonth = pCurDate.Month;
        aryLstFiles = new ArrayList();
        aryLstFilesStmt = new ArrayList();
        try
        {
            pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);
            vCurDate = pCurDate;
            strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "_" + pStmntType;

            clsBasFile.createDirectory(strOutputPath);
            clsBasFile.createDirectory(strOutputPath + "//STMT");
            #region Create Logs Path
            string txtLogsPath = strOutputPath + "\\Logs";
            List<string> sentaccounts = null;
            string pdfsPath = strOutputPath + "\\STMT";
            List<string> sentPdfs = null;
            if (!Directory.Exists(txtLogsPath))
            {
                Directory.CreateDirectory(txtLogsPath);
            }
            else
            {
                #region Read Last file and add pdfFiles toZip

                sentaccounts = new DirectoryInfo(txtLogsPath).GetFiles("*.txt", SearchOption.TopDirectoryOnly).Select(filename => filename.Name.Replace(".txt", string.Empty)).ToList();

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
                    noOfWithoutEmails += sentaccounts.Count - sentPdfs.Count;
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

            EventLog.WriteEntry("STMT", "open files stmt", EventLogEntryType.Error);

            // open emails file
            if (File.Exists(strEmailFileName + ".txt"))
            {
                fileEmails = new FileStream(strEmailFileName + ".txt", FileMode.Append); //Append
                streamEmails = new StreamWriter(fileEmails, Encoding.Default);
                streamEmails.AutoFlush = true;
            }
            else
            {
                fileEmails = new FileStream(strEmailFileName + ".txt", FileMode.Create); //Create
                streamEmails = new StreamWriter(fileEmails, Encoding.Default);
                streamEmails.WriteLine("AccountNumber" + "|" + "ClientID" + "|" + "Email" + "|" + "Date Time");
                streamEmails.AutoFlush = true;
            }

            //open No emails file
            //if (File.Exists(strNoEmailFileName + ".txt"))
            //{
            //    fileNoEmails = new FileStream(strNoEmailFileName + ".txt", FileMode.Append); //Append
            //    streamNoEmails = new StreamWriter(fileNoEmails, Encoding.Default);
            //    streamNoEmails.AutoFlush = true;
            //}
            //else
            //{
            //    fileNoEmails = new FileStream(strNoEmailFileName + ".txt", FileMode.Create); //Create
            //    streamNoEmails = new StreamWriter(fileNoEmails, Encoding.Default);
            //    streamNoEmails.WriteLine("CustomerName" + "|" + "AccountNumber" + "|" + "ClientID" + "|" + "Email" + "|");
            //    streamNoEmails.AutoFlush = true;
            //}

            // open No emails file
            fileNoEmails = new FileStream(strNoEmailFileName + ".txt", FileMode.Create); //Create
            streamNoEmails = new StreamWriter(fileNoEmails, Encoding.Default);
            streamNoEmails.WriteLine("CustomerName" + "|" + "AccountNumber" + "|" + "ClientID" + "|" + "Email" + "|");
            streamNoEmails.AutoFlush = true;


            attachedFilesStr = clsBasFile.getPathWithoutFile(strEmailFileName);

            fileStrmErr = new FileStream(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName), FileMode.Create);
            strmWriteErr = new StreamWriter(fileStrmErr, Encoding.Default);

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


            // set branch for data
            curBranchVal = pBankCode;

            if (createCorporateVal)
            {
                isCorporateVal = true;
                accountNoName = mCardaccountno;
                accountLimit = mCardlimit;
                accountAvailableLimit = mCardavailablelimit;
                ClientID = mCardClientId;
            }

            if (isRewardVal)
            {
                //maintainData.curRewardCond = rewardCond;
                //maintainData.fixReward(pBankCode, rewardCond);
                strMainTableCond = "m.contracttype != " + rewardCondVal;
                strSubTableCond = "d.trandescription != 'Calculated Points'";
                getReward(pBankCode);
            }
            #region Delete Old Files
            //if (HasAttachement)
            //{
            //    foreach (string fileName in Directory.GetFiles(strOutputPath))
            //    {
            //        if (clsBasFile.getFileExtn(fileName) == "pdf")
            //        {
            //            clsBasFile.deleteFile(fileName);
            //            break;
            //        }
            //    }
            //}
            #endregion
            if (IsSplitted)
            {
                MainTableCond = " m.cardproduct in " + strProductCond + "";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.cardproduct in " + strProductCond + ")";
            }

            FillStatementDataSet(pBankCode, ""); //DSstatement =  //10); //3
            getClientWithEmail(pBankCode);
            getClientPasNoAndBirthYear(pBankCode);
            pageNo = 1; totalCardPages = 0;
            curCardNo = String.Empty;
            curAccountNo = String.Empty;
            frmMain.BeginInvoke(frmMain.setMinMaxProgressDelegate, new object[] { DSstatement.Tables["tStatementMasterTable"].Rows.Count });

            EventLog.WriteEntry("STMT", "before foreach Tables[tStatementMasterTable].Rows", EventLogEntryType.Error);


            foreach (DataRow mRow in DSstatement.Tables["tStatementMasterTable"].Rows)
            {
                frmMain.BeginInvoke(frmMain.setProgressDelegate, new object[] { totRec++ });

                #region ignore this account
                if (sentaccounts != null && sentaccounts.Contains(mRow[accountNoName].ToString().Trim()))
                {
                    continue;
                }
                #endregion
                masterRow = mRow;


                curAccountNumber = masterRow[accountNoName].ToString();
                curCardNumber = strPrimaryCardNo;
                curClientID = masterRow[ClientID].ToString();
                curCustomerName = masterRow[mCustomername].ToString();

                strCardNo = masterRow[mCardno].ToString().Trim();

                if (strCardNo.Length != 16)
                {
                    //-  strmWriteErr.WriteLine("CardNo Length not equal 16 for AccountNo " + masterRow[accountNoName].ToString().Trim() + " Card No =" + strCardNo);
                    //numOfErr++;
                    continue;// Exclude Zero Length Cards 
                }
                strPrimaryCardNo = strCardNo;
                if (masterRow[mCardprimary].ToString() == "N")
                {
                    strPrimaryCardNo = masterRow[mPrinarycardno].ToString();
                }

                //start new account
                if (prevAccountNo != masterRow[accountNoName].ToString())
                {
                    //emailLabelTmp = emailLabel;
                    if (!string.IsNullOrEmpty(emailStr.ToString()) && !string.IsNullOrEmpty(emailTo))
                    {
                        SendEmail(emailStr.ToString(), "", emailTo);
                    }

                    emailStr = new StringBuilder("");
                    cardsRows = DSstatement.Tables["tStatementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[accountNoName]).ToString().Trim() + "'");

                    curMainCard = string.Empty;
                    isHaveF3 = false;

                    pageNo = 1;//totalAccPages = 1 ; pageNo=1;
                    CurPageRec4Dtl = 0; totNetUsage = 0; CurrentPageFlag = "F 1"; //if page is based on account no
                    calcAccountRows();
                    /*
                    curFileName = "QNBALAHLI_" + masterRow["CardProduct"].ToString().Trim() + "_Statement_For_" + currStatementdatefrom.ToString("MMMM") + currStatementdateto.ToString(" yyyy") + " " + masterRow[mAccountno].ToString().Trim();
                    curFileName = curFileName.Replace('.', ' ').Replace('/', ' ').Replace(' ', '_');

                    */
                    //curFileNamePrivate = vCurDate.ToString("MMMM yyyy ") + masterRow[mCardclientname].ToString().Replace('\t', ' ');
                    //curFileNamePrivate = curFileNamePrivate.Replace('\t', ' ').Replace('"', ' ').Replace(':', ' ');
                    //curFileNamePrivate = curFileNamePrivate.Replace('.', ' ').Replace('/', ' ').Replace(' ', '_');

                    curFileNamePrivate = vCurDate.ToString("MMMM yyyy ") + masterRow[mAccountno].ToString().Trim();
                    curFileNamePrivate = curFileNamePrivate.Replace('.', ' ').Replace('/', ' ').Replace(' ', '_');

                    emailLabelcurFileName = curFileNamePrivate;

                    //curFileNamePrivate = curFileNamePrivate + vCurDate.ToString("MMMM yyyy ") + "TEST";

                    if (totAccRows < 1 && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0) //             || (masterRow[mCardno].ToString() == curMainCard   // Convert.ToDecimal(
                    {
                        isHaveF3 = true;
                        continue;
                    }
                    emailStr = new StringBuilder("");

                    prevAccountNo = masterRow[accountNoName].ToString();

                    DataRow[] emailRows;
                    //emailRows = DSOemails.Tables["tClientPersone"].Select("idclient = " + masterRow[ClientID].ToString());
                    emailRows = DSOemails.Tables["Oemails"].Select("idclient = " + masterRow[ClientID].ToString());
                    //emailRows = DSemails.Tables["Emails"].Select("idclient = " + masterRow[ClientID].ToString());



                    for (int i = 0; i < emailRows.Length; i++)
                    {
                        emailTo = emailRows[i][1].ToString().Trim();
                        emailRow = emailRows[i];
                        ExternalID = emailRows[i][1].ToString().Trim();
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
                    stmNo = detailRow[dStatementno].ToString();
                    //stmNo = detailRow[dStatementnumber].ToString();
                    stmtNo = detailRow[dStatementno].ToString();
                    if ((detailRow[dPostingdate] == DBNull.Value) && (detailRow[dDocno] == DBNull.Value)) continue;// Exclude On-Hold Transactions 
                    curAccRows++;

                    totNetUsage = calculateCrDb(totNetUsage, clsBasValid.validateNum(detailRow[dBilltranamount]), clsBasValid.validateStr(detailRow[dBilltranamountsign]));
                    CurPageRec4Dtl = CurPageRec4Dtl + 1;


                } //end of detail foreach

                curCrdNoInAcc++;



                CurPageRec4Dtl = 0; //>pageNo=1; if pages is based on account
                curAccRows = 0;
                strEmailFromTmp = emailLabelTmp = string.Empty;


                /*
                For Credit, 
                Your Fidelity Bank Credit Card Statement February 2022

                For Mastercard PrePaid
                Your Fidelity Bank MC PrePaid Card Statement February 2022

                For Visa PrePaid
                Your Fidelity Bank PrePaid Card Statement February 2022

                string accType = masterRow[ACCOUNTTYPE].toString().Trim();
                string displayAccType = "Card";
                if(accType.equal("MC Prepaid Card (GHS)")){
                    displayAccType = "MC PrePaid Card"
                }
                else if(accType.equal("Prepaid Card (GHS)")){
                    displayAccType = "PrePaid Card"
                }
                emailLabelTmp = $"Your Fidelity Bank {displayAccType} Statement " + vCurDate.ToString("MMMM yyyy ");

                 */

                string accType = masterRow?[mAccounttype]?.ToString().Trim();
                string displayAccType = "Card";
                if (accType.Equals("MC Prepaid Card (GHS)"))
                {
                    displayAccType = "MC PrePaid Card";
                }
                else if (accType.Equals("Prepaid Card (GHS)"))
                {
                    displayAccType = "PrePaid Card";
                }
                emailLabelTmp = $"Your Fidelity Bank {displayAccType} Statement " + vCurDate.ToString("MMMM yyyy ");


                //emailLabelTmp = $"Your {emailLabelcurFileName} For " + vCurDate.ToString("MMMM yyyy ");

                if (stmtnumber == null)
                    stmtnumber = masterRow[accountNoName].ToString();

                clsStatement_ExportRpt exportRptabp = new clsStatement_ExportRpt();
                if (createCorporateVal)
                    exportRptabp.whereCond = "cardaccountno = '" + stmtnumber + "'";
                else
                    exportRptabp.whereCond = "accountno = '" + stmtnumber + "'";

                exportRptabp.whereCondD = "accountno = '" + stmtnumber + "'";



                exportRptabp.exportAttachment(pStrFileName, pBankCode, curFileNamePrivate, pReportName, "(IsNull({@f_POSTINGDATE}) and IsNull({@f_DOCNO}) and isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {@f_closingbalance} <> 0 and {TSTATEMENTMASTERTABLE.ACCOUNTNO} = '" + stmtnumber + "') or (not IsNull({@f_POSTINGDATE}) and not IsNull({@f_DOCNO}) and not isnull({@f_BranchTrans}) and {@F_branch} = " + pBankCode + " and {TSTATEMENTMASTERTABLE.ACCOUNTNO} = '" + stmtnumber + "')");
                exportRptabp = null;
                stmtnumber = null;

                stmtfilename = @clsBasFile.getPathWithoutFile(pStrFileName) + @"\STMT\" + curFileNamePrivate + ".pdf";
                stmtfilenameold = @clsBasFile.getPathWithoutFile(pStrFileName) + @"\" + curFileNamePrivate + ".pdf";

                if (frmStatementFile.Internal == true)
                {
                    stmtpass = clsCnfg.readSetting("InternalPassword");
                }
                else
                {
                    stmtpass = strPrimaryCardNo != "" ? strPrimaryCardNo.Substring(strPrimaryCardNo.Length - 6) : " ";
                }

                // Get a fresh copy of the sample PDF file
                string filenameSource = stmtfilenameold;
                string filenameDest = stmtfilenameold;
                PdfDocument document = PdfReader.Open(filenameDest);
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
                document.Save(filenameDest);


                #region write sent account number
                string strTxt = "Account number: " + masterRow?[accountNoName]?.ToString() ?? "";
                strTxt += Environment.NewLine + "Accountno: " + masterRow?[mAccountno]?.ToString() ?? "";
                strTxt += Environment.NewLine + "mCardno: " + masterRow?[mCardno]?.ToString() ?? "";
                strTxt += Environment.NewLine + "mCardproduct: " + masterRow?[mCardproduct]?.ToString() ?? "";
                strTxt += Environment.NewLine + "strPrimaryCardNo: " + strPrimaryCardNo ?? "";
                strTxt += Environment.NewLine + "mExternalno: " + masterRow?[mExternalno]?.ToString() ?? "";
                strTxt += Environment.NewLine + "ClientID: " + masterRow?[ClientID]?.ToString() ?? "";
                strTxt += Environment.NewLine + "mCustomername: " + masterRow?[mCustomername]?.ToString() ?? "";

                File.WriteAllText(txtLogsPath + "//" + masterRow[accountNoName].ToString().Trim() + ".txt", strTxt);
                #endregion

            } //end of Master foreach
            EventLog.WriteEntry("STMT", "after foreach Tables[tStatementMasterTable].Rows", EventLogEntryType.Error);

            if (!string.IsNullOrEmpty(emailStr.ToString()) && !string.IsNullOrEmpty(emailTo))
            {
                SendEmail(emailStr.ToString(), "", emailTo);
            }


        }
        catch (OracleException ex)
        {
            EventLog.WriteEntry("STMT", "OracleException " + ex.InnerException, EventLogEntryType.Error);

            clsDbOracleLayer.catchError(ex);
        }
        catch (NotSupportedException ex)  //(Exception ex)  //
        {
            clsBasErrors.catchError(ex);
        }
        catch (Exception ex)  //(Exception ex)  //
        {
            clsBasErrors.catchError(ex);
        }
        finally
        {
            EventLog.WriteEntry("STMT", "print summary  ", EventLogEntryType.Error);

            printStatementSummary();

            // Close Summary File
            streamSummary?.Flush();
            streamSummary?.Close();
            fileSummary?.Close();

            strmWriteErr?.Flush();
            strmWriteErr?.Close();
            fileStrmErr?.Close();


            streamEmails?.Flush();
            streamEmails?.Close();
            fileEmails?.Close();

            streamNoEmails.WriteLine(string.Empty);
            streamNoEmails.WriteLine(string.Empty);
            streamNoEmails.WriteLine("** Any Statement not have email will forward to email " + strEmailFrom);

            streamNoEmails?.Flush();
            streamNoEmails?.Close();
            fileNoEmails?.Close();


            aryLstFiles.Add(strEmailFileName + ".txt");
            aryLstFiles.Add(strNoEmailFileName + ".txt");
            aryLstFiles.Add(fileSummaryName);
            clsBasFile.generateFileMD5(aryLstFiles, strEmailFileName + ".MD5");
            aryLstFiles.Add(strEmailFileName + ".MD5");
            SharpZip zip = new SharpZip();
            zip.createZip(aryLstFiles, @clsBasFile.getPathWithoutExtn(pStrFileName) + ".zip", "");
            zip.createZip(aryLstFilesStmt, @clsBasFile.getPathWithoutFile(pStrFileName) + "\\" + pBankName + pStrFile + pStmntType + "_STMT" + vCurDate.ToString("yyyyMM") + ".zip", "");

            DSstatement?.Dispose();
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
        if (pageNo == 1 && totalAccPages == 1) // statement contain 1 page
        {
            CurrentPageFlag = "F 0";
            isHaveF3 = true;
        }
        else if (pageNo == 1 && totalAccPages > 1)  //first page of multiple page statement
            CurrentPageFlag = "F 1"; // //middle page of multiple page statement
        else if (pageNo < totalAccPages)
            CurrentPageFlag = "F 2";
        else if (pageNo == totalAccPages) //last page of multiple page statement
        {
            CurrentPageFlag = "F 3";
            isHaveF3 = true;
            endOfCustomer = "/";
        }



        //\MakeEmailStr   MakeHeaderStr   emailStr
        emailStr = new StringBuilder("");
        emailStr.Append(@"<html><head><meta http-equiv='Content-Language' content='en-us'><meta http-equiv='Content-Type' content='text/html; charset=windows-1252'>");
        emailStr.Append(@"<title>" + curFileNamePrivate + @"</title>");
        emailStr.Append(@"</head><body>");


        emailStr.Append(@"");
        emailStr.Append(@"<P  Dir=""LTR"" style=""text-align: justify;"">");
        emailStr.Append(@"<strong style=""font-size: 9pt;  font-family: Arial, Helvetica, sans-serif;"">Dear " + masterRow[mCustomername]?.ToString() ?? "Customer" + "</strong>");
        emailStr.Append(@"<br /><br /><span style=""font-size: 9pt;  font-family: Arial, Helvetica, sans-serif;"">Please find enclosed your latest e-statement for " + vCurDate.ToString("MMMM yyyy") + ".</span>");
        emailStr.Append(@"<br /><br /><span style=""font-size: 9pt;  font-family: Arial, Helvetica, sans-serif;"">The attached E-Statement is encrypted with a password of 6 digits for your protection. To view the document enter the last 6 digits of your primary card number.</span>");
        emailStr.Append(@"<br /><br /><span style=""font-size: 9pt;  font-family: Arial, Helvetica, sans-serif;"">Please verify the information contained in this Statement attached, should you need to update your details, have any queries or require any additional information on the above, please call our call center on 0800 00 3355 </span>");
        emailStr.Append(@"<br />");
        emailStr.Append(@"<br />");
        emailStr.Append(@"<img border=""0"" src='cid:" + clsBasFile.getFileFromPath(strBankLogo) + @"'/></td></tr><br/>");
        emailStr.Append(@"</P>");



        emailStr.Append(@"<img style=""margin:0;height:100%;width:100%;display:block;"" src='cid:" + clsBasFile.getFileFromPath(strbottomBanner) + @"'/>"); //advertise Photo


        emailStr.Append(@"</body></html>");
        totalPages++;
        totNoOfPageStat++;
    }


    private void calcCardlRows()
    {
        totalCardPages = 0;
        totCardRows = 0;
        foreach (DataRow dtRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
        {
            if ((dtRow[dPostingdate] == DBNull.Value) && (dtRow[dDocno] == DBNull.Value)) continue;
            totCardRows++;
        }
        if (totCardRows > MaxDetailInLastPage)
        {
            //totalCardPages = 1;
            totCardRows -= MaxDetailInLastPage;
            totalCardPages++;
            totalCardPages += (totCardRows / MaxDetailInPage);
            if ((totCardRows % MaxDetailInPage) > 0)
                totalCardPages++;
            if (totalCardPages < 1)
                totalCardPages += 1;
        }
        else
        {
            totalCardPages = 1;
        }
    }


    private void calcAccountRows()
    {
        curAccRows = 0;
        accountRows = null;
        stmNo = masterRow[mStatementno].ToString();
        stmNo = masterRow[accountNoName].ToString();

        accountRows = DSstatement.Tables["tStatementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[accountNoName]).ToString().Trim() + "'");
        totalAccPages = 0;
        totAccRows = 0;
        string prevCardNo = String.Empty, CurCardNo = String.Empty;
        int currAccRowsPages = 0;
        foreach (DataRow dtAccRow in accountRows) //mRow.GetChildRows(StatementNoDRel)
        {
            if ((dtAccRow[dPostingdate] == DBNull.Value) && (dtAccRow[dDocno] == DBNull.Value)) continue;
            CurCardNo = dtAccRow[dCardno].ToString();
            if (CurCardNo.Trim().Length < 1) continue;

            currAccRowsPages++;
            totAccRows++;

            if (currAccRowsPages > MaxDetailInPage)//==
            {
                currAccRowsPages = 1;//0
                totalAccPages++;
            }
            prevCardNo = dtAccRow[dCardno].ToString();
        }
        if (currAccRowsPages > 0)
            totalAccPages++;
        if (totalAccPages < 1)
            totalAccPages = 1;

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

    private void completePageDetailRecords()
    {
        if (pageNo == totalAccPages)
            for (int curPageLine = CurPageRec4Dtl; curPageLine < MaxDetailInPage; curPageLine++)
                streamWrit.WriteLine(String.Empty);
    }

    private void printStatementSummary()
    {
        EventLog.WriteEntry("STMT", "print statement summary begin", EventLogEntryType.Error);

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
        EventLog.WriteEntry("STMT", "print statement summary end", EventLogEntryType.Error);


    }

    private void printFileMD5()
    {
        FileStream fileStrmMd5;
        StreamWriter streamWritMD5;
        fileStrmMd5 = new FileStream(strOutputPath + "\\" + strBankName + vCurDate.ToString("yyyyMM") + ".txt", FileMode.Create);
        streamWritMD5 = new StreamWriter(fileStrmMd5);
        streamWritMD5.WriteLine(clsBasFile.getFileFromPath(strOutputFile) + "  >>  " + clsBasFile.getFileMD5(strOutputFile));
        streamWritMD5.WriteLine(clsBasFile.getFileFromPath(fileSummaryName) + "  >>  " + clsBasFile.getFileMD5(fileSummaryName));
        streamWritMD5.Flush();
        streamWritMD5.Close();
        fileStrmMd5.Close();
    }

    private string MakeEmailStr(string pStr)
    {
        pStr = pStr.Replace("\"", "\"\"");
        return pStr;
    }

    private string MakeHeaderStr(string pStr)
    {
        pStr = pStr.Trim();
        if (pStr.Length < 1)
            pStr = "&nbsp;";
        else
            pStr = @"<font size=""2"">" + pStr + "</font>";

        return pStr;
    }

    private string MakeHeaderStr(string pStr, bool isBold)
    {
        string color = string.Empty;
        pStr = pStr.Replace("&", "&amp;").Trim();

        if (pStr.Length < 1)
            pStr = "&nbsp;";
        else
        {
            pStr = @"<font size=""1"">" + pStr + @"</font>";
            if (isBold)
                pStr = @"<b>" + pStr + "</b>";
        }

        return pStr;
    }

    private string ValidateStr(string pStr)
    {
        if (pStr.Trim().Length < 1)
            pStr = "&nbsp;";
        return pStr;
    }

    private void SendEmail(string pBody, string pSubject, string pTo)
    {
        try
        {
            if (valdEmail.isValideEmail(pTo) != "ValidEmail")//!basText.isValideEmail(pTo)
            {
                streamNoEmails.WriteLine(curCustomerName + "|" + curAccountNumber + "|" + curClientID + "|" + " " + "|" + "|" + clsCnfg.readSetting("strValidEmail") ?? "Bad Email " + pTo);
                noOfBadEmails++;
                return;

            }


            ArrayList pLstTo = new ArrayList(), pLstAttachedFile = new ArrayList(), pLstAttachedPic = new ArrayList();//, pLstCC = new ArrayList(), pLstBCC = new ArrayList();
            pSubject = emailLabelTmp; //emailLabel "BAI statement for 02/2008";

            if (pTo.EndsWith("."))
                pTo = pTo + "com";
            pLstTo.Add(pTo);

            pLstAttachedPic.Add(strBankLogo);
            pLstAttachedPic.Add(strbottomBanner);
            //pLstAttachedPic.Add(strbottomBanner1);
            //pLstAttachedPic.Add(strBackGround);

            if (HasAttachement)
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
            streamEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + pTo + "|" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));

            numMail++;
            if (numMail % 400 == 0)
            {
                System.Threading.Thread.Sleep(waitPeriodVal);//2000
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            if (HasAttachement)
            {
                clsBasFile.moveFile(stmtfilenameold, stmtfilename);
                aryLstFilesStmt.Add(stmtfilename);
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


    private void add2FileList(string pFileName)
    {
        int myIndex = aryLstFiles.BinarySearch((object)pFileName);
        if (myIndex < 0)
            aryLstFiles.Add(@pFileName);
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

    public string curFileName
    {
        get { return curFileNamePrivate; }
        set { curFileNamePrivate = value; }
    }// curFileName

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

    ~clsStatHtmlFBPG()
    {
        DSstatement?.Dispose();
    }
}
