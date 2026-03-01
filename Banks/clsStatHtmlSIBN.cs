using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Xml;
using System.Collections;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Security;
using PdfSharp.Pdf.IO;
using System.Collections.Generic;
using System.Linq;

public class clsStatHtmlSIBN : clsBasStatement
{
    private string strBankName;
    private FileStream fileStrm, fileSummary, fileEmails, fileNoEmails;
    private StreamWriter streamWrit, streamSummary, streamEmails, streamNoEmails;
    //private DataSet DSstatement;
    //private OracleDataReader drPrimaryCards, drMaster,drDetail;
    private DataRow masterRow;
    private DataRow detailRow;
    private string strEndOfLine = "\u000D";  //+ "M" ^M
    //		private string strEndOfPage = "\u000C"  ;  //+ "\u000D"+ "M" ^L + ^M
    private string strEndOfPage = "\u000C";  //  + basText.replicat(" ",85) + "F 2"  + "\u000D"+ "M" ^L + ^M
    private const int MaxDetailInPage = 20; //
    private const int MaxDetailInLastPage = 27; //
    private int CurPageRec4Dtl = 0;
    private int pageNo = 0, totalPages = 0, totalCardPages = 0
        , totalAccPages = 0, totCardRows = 0, totAccRows = 0, curAccRows = 0;//
    private string lastPageTotal;
    private string curCardNo;//,PrevCardNo
    private string curAccountNo, prevAccountNo = String.Empty;//,PrevCardNo
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
    string cProduct = string.Empty, curFileName = string.Empty;
    private ArrayList aryLstFiles, aryLstFilesStmt;

    //
    //private string emailStr;
    private StringBuilder emailStr = new StringBuilder("");
    private string emailLabel, strBankLogo;
    private string strEmailFrom = "mabouleila@emp-group.com";   //"cardservices@zenithbank.com"
    private string strbankWebLink = "www.emp-group.com";  //www.zenithbank.com
    private string emailTo, curAccountNumber, curCardNumber, curClientID, curCustomerName;//, curEmail
    private string strEmailFileName, strNoEmailFileName;
    private frmStatementFile frmMain;
    //private frmGenerateSatatement generateStatus = new frmGenerateSatatement();
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
    private string strbottomBanner = string.Empty;
    private bool HasSenderVal = false;

    public clsStatHtmlSIBN()
    {
    }

    public string Statement(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pCurDate, string pStmntType, bool pAppendData, string pReportName)
    {
        string rtrnStr = "Successfully Generate " + pBankName;
        int curMonth = pCurDate.Month;
        bool preExit = true;
        strFileNam = pStrFileName;
        stmntType = pStmntType;

        curMonth = pCurDate.Month;
        aryLstFiles = new ArrayList();
        aryLstFilesStmt = new ArrayList();
        try
        {
            //clsMaintainData maintainData = new clsMaintainData();
            //maintainData.matchCardBranch4Account(pBankCode);

            pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);
            vCurDate = pCurDate; //DateTime.Now.AddMonths(-1);
            strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "_" + pStmntType;
            clsBasFile.deleteDirectory(strOutputPath);
            clsBasFile.createDirectory(strOutputPath);
            if (HasAttachement)
                clsBasFile.createDirectory(strOutputPath + "//STMT"); //SIBN-2604

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
            emailLabel = pBankName + " statement for " + vCurDate.ToString("MM/yyyy"); //"BAI statement for 02/2008"
            strOutputFile = pStrFileName;

            // open emails file
            fileEmails = new FileStream(strEmailFileName + ".txt", FileMode.Create); //Create
            streamEmails = new StreamWriter(fileEmails, Encoding.Default);
            //streamEmails.WriteLine("AccountNumber" + "|" + "CardNumber" + "|" + "ClientID" + "|" + "Email");
            streamEmails.WriteLine("AccountNumber" + "|" + "ClientID" + "|" + "Email" + "|" + "Date Time");
            streamEmails.AutoFlush = true;

            // open No emails file
            fileNoEmails = new FileStream(strNoEmailFileName + ".txt", FileMode.Create); //Create
            streamNoEmails = new StreamWriter(fileNoEmails, Encoding.Default);
            streamNoEmails.WriteLine("CustomerName" + "|" + "AccountNumber" + "|" + "ClientID" + "|" + "Email" + "|" + "Mobile Phone");
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
            curBranchVal = pBankCode; // 10; //3 = real   1 = test
           
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
            getClientEmail(pBankCode);
            pageNo = 1; totalCardPages = 0;
            curCardNo = String.Empty;
            curAccountNo = String.Empty;
           
            frmMain.BeginInvoke(frmMain.setMinMaxProgressDelegate, new object[] { DSstatement.Tables["tStatementMasterTable"].Rows.Count });
           
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
                

                strCardNo = masterRow[mCardno].ToString().Trim();

                

                if (strCardNo.Length != 16)
                {
                    continue;// Exclude Zero Length Cards 
                }
                strPrimaryCardNo = strCardNo;
                if (masterRow[mCardprimary].ToString() == "N")
                {
                    strPrimaryCardNo = masterRow[mPrinarycardno].ToString();
                    //calcCardlRows();
                }

                //start new account
                if (prevAccountNo != masterRow[accountNoName].ToString())
                {
                    emailLabelTmp = emailLabel;
                    if (!string.IsNullOrEmpty(emailStr.ToString()) && !string.IsNullOrEmpty(emailTo)) 
                        SendEmail(emailStr.ToString(), "", emailTo);

                    if (!string.IsNullOrEmpty(emailStr.ToString()) && string.IsNullOrEmpty(emailTo)) 
                    {
                        streamNoEmails.WriteLine(curCustomerName + "|" + curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|Without Email");
                        noOfWithoutEmails++;
                        emailTo = strEmailFrom; //"statement_Program@emp-group.com";
                        emailLabelTmp = emailLabel + " Acc:" + curAccountNumber + " Phone:" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!");   
                        SendEmail(emailStr.ToString(), "", emailTo);
                    }

                    emailStr = new StringBuilder("");
                    cardsRows = DSstatement.Tables["tStatementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[accountNoName]).ToString().Trim() + "'");

                   

                    curMainCard = string.Empty;
                    
                    isHaveF3 = false;

                    pageNo = 1;//totalAccPages = 1 ; pageNo=1;
                    CurPageRec4Dtl = 0; totNetUsage = 0; CurrentPageFlag = "F 1"; //if page is based on account no
                    calcAccountRows();

                    //SIBN-2604
                    curFileName = "Stanbic Bank IBTC " + masterRow[mAccountcurrency].ToString() + " Credit Card Statement " + vCurDate.ToString("MMMM yyyy ") + masterRow[mCardclientname].ToString();
                    curFileName = curFileName.Replace('.', ' ').Replace('/', ' ').Replace("  ","").Replace(' ', '_');

                    System.Text.StringBuilder filename2 = new System.Text.StringBuilder();

                    if ( strCardNo ==Convert.ToString(4445060000146695))     //4445060000146695
                    {
                    for (int i = 0; i < curFileName.Length; i++)
                    {
                        if (char.IsLetter(curFileName[i]) || curFileName[i] == '_')
                        {
                            filename2.Append(curFileName[i]);
                        }

                    }

                        curFileName = filename2.ToString();

                    }

                    //HUSSAINA   RUTH NANA   ISHAYA - AUD
                    //11111111   1111 1111   111111_

                    if (totAccRows < 1
                      && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0) //             || (masterRow[mCardno].ToString() == curMainCard   // Convert.ToDecimal(
                    {
                        isHaveF3 = true;
                        //pageNo=1; totalAccPages =1;
                        continue;
                    }
                    emailStr = new StringBuilder("");

                    prevAccountNo = masterRow[accountNoName].ToString();
                    //pageNo=1; //if page is based on account no
                    DataRow[] emailRows;
                    emailRows = DSemails.Tables["Emails"].Select("idclient = " + masterRow[ClientID].ToString());
                    for (int i = 0; i < emailRows.Length; i++)
                    {
                        emailTo = emailRows[i][1].ToString().Trim();
                        emailRow = emailRows[i];
                    }
                    curAccountNumber = masterRow[accountNoName].ToString();
                    curCardNumber = strPrimaryCardNo;
                    curClientID = masterRow[ClientID].ToString();
                    curCustomerName = masterRow[mCustomername].ToString();


                    printBody(); //SIBN-2604
                    totNoOfCardStat++;
                }
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
               
                if (HasAttachement)
                {
                    try
                    {
                        if (stmtnumber == null)
                            stmtnumber = masterRow[accountNoName].ToString();
                        clsStatement_ExportRpt exportRptSibn = new clsStatement_ExportRpt();
                        if (createCorporateVal)
                            exportRptSibn.whereCond = "cardaccountno = '" + stmtnumber + "' and packagename='STTOTABLECORP'";
                        else
                            exportRptSibn.whereCond = "accountno = '" + stmtnumber + "'";

                        exportRptSibn.whereCondD = "accountno = '" + stmtnumber + "'";
                        curFileName.Replace("\t", "_");
                        exportRptSibn.exportAttachment(pStrFileName, pBankCode, curFileName, pReportName);
                        exportRptSibn = null;
                        stmtnumber = null;
                        stmtfilename = @clsBasFile.getPathWithoutFile(pStrFileName) + @"\STMT\" + curFileName + ".pdf";
                        stmtfilenameold = @clsBasFile.getPathWithoutFile(pStrFileName) + @"\" + curFileName + ".pdf";
                        if (frmStatementFile.Internal == true)
                        {
                            stmtpass = clsCnfg.readSetting("InternalPassword");
                        }
                        else
                        {
                            stmtpass = curMainCard.Substring(curMainCard.Length - 6).Trim(); //SIBN-2604
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
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                }

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
            emailLabelTmp = emailLabel;
            if (!string.IsNullOrEmpty(emailStr.ToString()) && !string.IsNullOrEmpty(emailTo)) //emailStr != string.Empty && emailStr != null  emailTo != string.Empty && emailTo != null
                //if (HasSender == true)
                //    SendEmailWithDifferentSender(emailStr.ToString(), "", emailTo);
                //else
                SendEmail(emailStr.ToString(), "", emailTo);
            if (!string.IsNullOrEmpty(emailStr.ToString()) && string.IsNullOrEmpty(emailTo)) // emailStr != string.Empty && emailStr != null emailTo != string.Empty && emailTo != null
            {
                streamNoEmails.WriteLine(curCustomerName + "|" + curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|Without Email");
                noOfWithoutEmails++;
                emailTo = strEmailFrom; //"statement_Program@emp-group.com";
                emailLabelTmp = emailLabel + " Acc:" + curAccountNumber + " Phone:" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!");
                //if (HasSender == true)
                //    SendEmailWithDifferentSender(emailStr.ToString(), "", emailTo);
                //else
                SendEmail(emailStr.ToString(), "", emailTo);
            }
            /*
                 // Write Data to Text File 
                 writeStatement(DSstatement,pStrFileName);
                 */
            // Write Data to XML File 
            /*
                string XmlFileName = clsBasFile.getPathWithoutExtn(pStrFileName) + ".xml";


                StreamWriter xmlSW = new StreamWriter(XmlFileName); //"Statement.xml"
                DSstatement.WriteXml(xmlSW, XmlWriteMode.IgnoreSchema);
                xmlSW.Close();
                */
            //clsBasStatementRawData RawData = new clsBasStatementRawData(clsBasFile.getPathWithoutExtn(pStrFileName), DSstatement);
            //fillStatementHistory(pStmntType,pAppendData);
            //>clsBasStatementRawData RawData = new clsBasStatementRawData(clsBasFile.getPathWithoutExtn(pStrFileName), DSstatement);
            //>clsBasStatementRawExcel RawExcel = new clsBasStatementRawExcel(clsBasFile.getPathWithoutExtn(pStrFileName), DSstatement);
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
            // Close output File
            //-streamWrit.Flush();
            //-streamWrit.Close();
            //-fileStrm.Close();

            printStatementSummary();

            // Close Summary File
            streamSummary.Flush();
            streamSummary.Close();
            fileSummary.Close();

            strmWriteErr.Flush();
            strmWriteErr.Close();
            fileStrmErr.Close();

            //-if (numOfErr == 0)
            //-  clsBasFile.deleteFile(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName));
            //-else
            //-  rtrnStr += string.Format(" with {0} Errors", numOfErr);
            /*else
              clsBasFile.deleteFile(pStrFileName);*/
            streamEmails.Flush();
            streamEmails.Close();
            fileEmails.Close();

            streamNoEmails.WriteLine(string.Empty);
            streamNoEmails.WriteLine(string.Empty);
            streamNoEmails.WriteLine("** Any Statement not have email will forward to email " + strEmailFrom);
            streamNoEmails.Flush();
            streamNoEmails.Close();
            fileNoEmails.Close();
            //ArrayList aryLstFiles = new ArrayList();
            //aryLstFiles.Add("");
            //aryLstFiles.Add(@strOutputFile);
            //-numOfErr += validateNoOfLines(aryLstFiles, 48);
            //-aryLstFiles.Add(@fileSummaryName);
            aryLstFiles.Add(strEmailFileName + ".txt");
            aryLstFiles.Add(strNoEmailFileName + ".txt");
            aryLstFiles.Add(fileSummaryName);
            clsBasFile.generateFileMD5(aryLstFiles, strEmailFileName + ".MD5");
            aryLstFiles.Add(strEmailFileName + ".MD5");
            SharpZip zip = new SharpZip();
            zip.createZip(aryLstFiles, @clsBasFile.getPathWithoutExtn(pStrFileName) + ".zip", "");
            zip.createZip(aryLstFilesStmt, @clsBasFile.getPathWithoutFile(pStrFileName) + "\\" + pBankName + pStrFile + pStmntType + "_STMT" + vCurDate.ToString("yyyyMM") + ".zip", ""); //SIBN-2604

            DSstatement.Dispose();
        }
        return rtrnStr;
    }

    protected void printBody()
    {
        emailStr = new StringBuilder("");
        emailStr.Append(@"<html><head><meta http-equiv='Content-Language' content='en-us'><meta http-equiv='Content-Type' content='text/html; charset=windows-1252'>");
        emailStr.Append(@"<title>" + curFileName + @"</title>");
        emailStr.Append(@"</head><body>");
        emailStr.Append(@"Dear " + masterRow[mCustomername].ToString() + "</br>");
        emailStr.Append(@"</br>");
        emailStr.Append(@"Your credit card e-statement for the month is attached.</br>");
        emailStr.Append(@"To view your statement, kindly input the last 6 (six) digits of your card number as the password.</br>");
        emailStr.Append(@"You can make your minimum repayment via internet banking, fund transfers from ANY bank or by depositing cash at any Stanbic IBTC Bank branch. You can also take advantage of our direct debit repayment.</br>");
        emailStr.Append(@"</br>");
        emailStr.Append(@"For inquiries, please contact us at 01 422 2222 or EBankingQueries@stanbicibtc.com.</br>");
        emailStr.Append(@"</br>");
        emailStr.Append(@"Thank you for choosing Stanbic IBTC Bank.</br>");
        emailStr.Append(@"</br>");
        //emailStr.Append(@"<a href='https://travel.visa.com/apcemea/ng/en/offers-abroad.html'><img border='0' src='cid:" + clsBasFile.getFileFromPath(strbottomBanner) + @"'>"); //SIBN-3421
        emailStr.Append(@"</body></html>");

    }

    protected void printHeader()
    {
        string newaddress1, newaddress2;
        clsMaintainData.fixAddress(curBranchVal, masterRow[mCustomeraddress1].ToString(), out newaddress1, out newaddress2);
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
        emailStr.Append(@"<title>");
        emailStr.Append(masterRow[mCustomername] + " - Statement");
        emailStr.Append(@"</title>");
        emailStr.Append(@"</head><body  background='cid:" + clsBasFile.getFileFromPath(strbackGround) + @"'>");
        emailStr.Append(@"<a href='https://travel.visa.com/apcemea/ng/en/offers-abroad.html'><img border='0' src='cid:" + clsBasFile.getFileFromPath(strbottomBanner) + @"'>");
        emailStr.Append(@"<table border='0' width='100%' id='table1'>");

        emailStr.Append(@"<tr><td colSpan='2' height='82'>");
        emailStr.Append(@"<table id='table25' width='100%' border='0'>");
        emailStr.Append(@"<tr><td height='82' width='27%'><p align='" + logoAlignmentStr + @"'>");
        emailStr.Append(@"<a href='http://" + strbankWebLink + @"'><img border='0' src='cid:" + clsBasFile.getFileFromPath(strBankLogo) + @"'></a></p></td>");
        emailStr.Append(@"<td height='82' width='3%' align='left'>&nbsp;</td><td height='82' align='center'>");
        emailStr.Append(@"<font size='6'>Statement</font></td><td height='82' width='1%'>&nbsp;</td>");
        emailStr.Append(@"<td height='82'><img border='0' src='cid:" + clsBasFile.getFileFromPath(strVisaLogo) + @"' width='110' height='35' align='right' />");
        emailStr.Append(@"</td></tr></table>");
        emailStr.Append(@"</td></tr>");
        ///

        //>>
        emailStr.Append(@"<tr><td width='51%'><table border='1' width='99%' id='table5' height='119'><tr><td><table border='0' width='99%' id='table6' bordercolor='#000000'>");
        emailStr.Append(@"<tr><td align='center'>");
        emailStr.Append(MakeHeaderStr(masterRow[mCustomername].ToString()));
        emailStr.Append(@"</td></tr>");
        emailStr.Append(@"<tr><td align='center'>");
        //emailStr.Append(MakeHeaderStr(ValidateArbic(masterRow[mCustomeraddress1].ToString())));//"&nbsp;"; // 
        emailStr.Append(MakeHeaderStr(ValidateArbic(newaddress1)));//"&nbsp;"; // 
        emailStr.Append(@"</td></tr>");
        emailStr.Append(@"<tr><td align='center'>");
        //emailStr.Append(MakeHeaderStr(ValidateArbic(masterRow[mCustomeraddress2].ToString())));//"&nbsp;"; // 
        emailStr.Append(MakeHeaderStr(ValidateArbic(newaddress2)));//"&nbsp;"; // 
        emailStr.Append(@"</td></tr>");
        emailStr.Append(@"<tr><td align='center'>");
        emailStr.Append(MakeHeaderStr(ValidateArbic(masterRow[mCustomeraddress3].ToString())));//"&nbsp;"; // 
        emailStr.Append(@"</td></tr>");
        emailStr.Append(@"<tr><td align='center'>");
        emailStr.Append(MakeHeaderStr(masterRow[mCustomerregion].ToString().Trim() + " " + masterRow[mCustomercity].ToString().Trim()));
        emailStr.Append(@"</td></tr>");
        emailStr.Append(@"</table></td></tr></table></td>");
        //>>

        emailStr.Append(@"<td width='48%'><table border='1' width='99%' id='table8' height='119'><tr><td><table border='0' width='100%' id='table9' height='108'>");
        emailStr.Append(@"<tr><td width='81'><b>");
        emailStr.Append(@"<font size='1'>Card Product :</font>");
        emailStr.Append(@"</b></td>");
        emailStr.Append(@"<td>");
        emailStr.Append(MakeHeaderStr(masterRow[mCardproduct].ToString().Trim()));
        emailStr.Append(@"</td></tr>");

        emailStr.Append(@"<tr><td width='81'><b><font size='1'>Branch :</font></b></td><td>");
        emailStr.Append(MakeHeaderStr(masterRow[mCardbranchpartname].ToString().Trim()));
        emailStr.Append(@"</td></tr>");

        emailStr.Append(@"<tr><td width='81'><b><font size='1'>Acc. Number :</font></b></td><td>");
        emailStr.Append(MakeHeaderStr(extAccNum));
        emailStr.Append(@"</td></tr>");

        emailStr.Append(@"<tr><td width='81'><b><font size='1'>Statement Date:</font></b></td><td>");
        emailStr.Append(MakeHeaderStr(Convert.ToDateTime(masterRow[mStatementdateto]).ToString("dd/MM/yyyy"))); //"{0,10:dd/MM/yyyy}", 
        emailStr.Append(@"</td></tr>");

        emailStr.Append(@"</table></td></tr></table></td></tr>");


        emailStr.Append(@"<tr><td width='98%' colspan='2' height='20'>&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td width='98%' height='77' colspan='2'><table border='1' width='100%' id='table10' height='50'>");
        emailStr.Append(@"<tr>");

        emailStr.Append(@"<td align='center' width='102'><b>");
        emailStr.Append(@"<font size='1'>Credit Limit</font>");
        emailStr.Append(@"</b></td>");
        emailStr.Append(@"<td align='center' width='118'><b>");
        emailStr.Append(@"<font size='1'>Available Limit</font>");
        emailStr.Append(@"</b></td>");
        emailStr.Append(@"<td width='111' align='center'><b>");
        emailStr.Append(@"<font size='1'>Min. Due Amount</font>");
        emailStr.Append(@"</b></td>");
        emailStr.Append(@"<td width='147' align='center'><b>");
        emailStr.Append(@"<font size='1'>Min. Due Date</font>");
        emailStr.Append(@"</b></td>");
        emailStr.Append(@"<td width='114' align='center'><b>");
        emailStr.Append(@"<font size='1'>Over Due Amount</font>");
        emailStr.Append(@"</b></td></tr>");

        emailStr.Append(@"<tr><td height='24' align='center' width='102'>");
        //emailStr.Append(MakeHeaderStr(masterRow[accountLimit].ToString().Trim()));
        emailStr.Append(MakeHeaderStr(masterRow[mAccountlim].ToString().Trim()));
        emailStr.Append(@"</td><td height='24' align='center' width='118'>");
        emailStr.Append(@"<font size='1' color='#800000'>");
        //emailStr.Append(basText.formatNum(masterRow[accountAvailableLimit], "##0", 20).Trim());
        emailStr.Append(basText.formatNum(masterRow[mAccountavailablelim], "##0", 20).Trim());
        emailStr.Append(@"</font></td>");
        emailStr.Append(@"<td height='24' width='111' align='center'><b>");
        emailStr.Append(@"<font size='1' color='#800000'>");
        emailStr.Append(masterRow[mMindueamount].ToString().Trim());
        emailStr.Append(@"</font></b></td>");
        emailStr.Append(@"<td height='24' width='147' align='center'>");
        emailStr.Append(MakeHeaderStr(basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M")));
        emailStr.Append(@"</td>");
        emailStr.Append(@"<td height='24' width='114' align='center'>");
        emailStr.Append(MakeHeaderStr(basText.formatNum(masterRow[mTotaloverdueamount], "##0.00", 13)));
        emailStr.Append(@"</td></tr>");
        emailStr.Append(@"</table></td></tr>");

        emailStr.Append(@"<tr><td width='98%' height='20' colspan='2'>&nbsp;</td></tr>");

        emailStr.Append(@"<tr><td width='98%' height='88' colspan='2'><table border='1' width='100%' id='table11' height='60'>");//109
        emailStr.Append(@"<tr><td align='center' height='16' colspan='2'>");
        emailStr.Append(@"<font size='1'>Dates of</font>");
        emailStr.Append(@"</td>");
        emailStr.Append(@"<td align='center' height='34' width='19%' rowspan='2'><b>");
        emailStr.Append(@"<font size='1'>Reference No</font>");
        emailStr.Append(@"</b></td>");
        emailStr.Append(@"<td align='center' height='53' width='52%' rowspan='2' colspan='2'><b>");
        emailStr.Append(@"<font size='1'>Description</font>");
        emailStr.Append(@"</b></td>");
        emailStr.Append(@"<td width='13%' align='center' height='34' rowspan='2' colspan='2'><b>");
        emailStr.Append(@"<font size='1'>Amount ( " + masterRow[mAccountcurrency].ToString() + @" )</font>");
        emailStr.Append(@"</b></td></tr>");

        emailStr.Append(@"<tr><td align='center' height='16' width='6%'>");
        emailStr.Append(@"<font size='1'>Trans.</font>");
        emailStr.Append(@"</td><td align='center' height='16' width='5%'>");
        emailStr.Append(@"<font size='1'>Posting</font>");
        emailStr.Append(@"</td></tr>");

        //\

        totalPages++;
        totNoOfPageStat++;
    }


    protected void printDetail()
    {
        DateTime trnsDate = clsBasValid.validateDate(detailRow[dTransdate]), postingDate = clsBasValid.validateDate(detailRow[dPostingdate]);
        if (trnsDate > postingDate)
            trnsDate = postingDate;

        if (masterRow[mAccountcurrency].ToString().Trim() != detailRow[dOrigtrancurrency].ToString().Trim())
            strForeignCurr = basText.formatNumUnSign(detailRow[dOrigtranamount], "##,###,##0.00;(#,###,##0.00)", 13) + " " + basText.Replace(clsBasValid.validateStr(detailRow[dOrigtrancurrency]), "XXX", "   ");
        else
            strForeignCurr = basText.replicat(" ", 16);

        string trnsDesc; //= detailRow[dTrandescription].ToString().Trim() + " " + detailRow[dMerchant].ToString().Trim()
        if (detailRow[dMerchant].ToString().Trim() == "")
            trnsDesc = detailRow[dTrandescription].ToString().Trim();
        else
            trnsDesc = detailRow[dMerchant].ToString().Trim();

        emailStr.Append(@"<tr><td width='6%' align='center' height='23'>");
        emailStr.Append(MakeHeaderStr(trnsDate.ToString("dd/MM")));
        emailStr.Append(@"</td>");
        emailStr.Append(@"<td width='5%' align='center' height='23'>");
        emailStr.Append(MakeHeaderStr(postingDate.ToString("dd/MM")));
        emailStr.Append(@"</td>");
        emailStr.Append(@"<td width='19%' align='center' height='23'>");
        emailStr.Append(MakeHeaderStr(basText.trimStr(detailRow[dRefereneno].ToString().Trim(), 24)));
        emailStr.Append(@"</td>");

        //if (strForeignCurr.Trim() != "")
        //{ //default
        emailStr.Append(@"<td width='42%' height='23'><p align='left'>");
        emailStr.Append(MakeHeaderStr(basText.trimStr(trnsDesc, 40)));
        emailStr.Append(@"</td>");
        emailStr.Append(@"<td width='10%' align='right' height='23'>");
        emailStr.Append(MakeHeaderStr(strForeignCurr));
        emailStr.Append(@"</td>");
        //}
        //else
        //{
        //  emailStr.Append(@"<td width='52%' height='23'><p align='left'>";
        //  emailStr.Append(MakeHeaderStr(basText.trimStr(trnsDesc, 40));
        //  emailStr.Append(@"</td>";
        //}

        //if (isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString())) != "")
        //{ //default
        emailStr.Append(@"<td width='11%' align='right' height='23'>");
        emailStr.Append(MakeHeaderStr((basText.formatNumUnSign(detailRow[dBilltranamount], "#,###,##0.00", 14) + " " + CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])).ToString())));
        emailStr.Append(@"</td>");
        emailStr.Append(@"<td width='1%' align='right' height='23'>");
        emailStr.Append(MakeHeaderStr(isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))));
        //}
        //else
        //{
        //  emailStr.Append(@"<td width='12%' align='right' height='23'>";
        //  emailStr.Append(MakeHeaderStr((basText.formatNumUnSign(detailRow[dBilltranamount], "#,###,##0.00", 14) + " " + CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])).ToString()));
        //}

        emailStr.Append(@"</td></tr>");

        //\
        //			streamWrit.WriteLine("  {0:dd/MM}  {1:dd/MM}  {2,-40} {3,12} {4,3} {5,16} {6,2}",trnsDate,postingDate,basText.trimStr(detailRow[dTrandescription].ToString().Trim() + " " + detailRow[dMerchant].ToString().Trim(),45),basText.formatNum(detailRow[dOrigtranamount],"#,##0.00;(#,##0.00)"),basText.Replace(clsBasValid.validateStr(detailRow[dOrigtrancurrency]),"XXX","   "),basText.formatNum(detailRow[dBilltranamount],"#,##0.00;(#,##0.00)"),CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)
        //-streamWrit.WriteLine("  {0:dd/MM}  {1:dd/MM}  {2,-24}  {3,-40} {4,16} {5,16} {6,2}", trnsDate, postingDate, basText.trimStr(detailRow[dRefereneno].ToString().Trim(), 24), basText.trimStr(trnsDesc, 40), strForeignCurr, basText.formatNumUnSign(detailRow[dBilltranamount], "#,###,##0.00", 14), CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + " " + isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)
        totNoOfTransactions++;
    }

    protected void printCardFooter()
    {
        //\MakeHeaderStr(basText.trimStr(detailRow[dRefereneno].ToString().Trim(), 24));
        emailStr.Append(@"</table></td></tr>");

        emailStr.Append(@"<tr><td width='99%' height='20' colspan='2'>&nbsp;</td></tr>");

        emailStr.Append(@"<tr><td width='98%' colspan='2'><table border='1' width='100%' id='table12'><tr><td><table border='0' width='100%' id='table13'>");
        //emailStr.Append(@"<tr><td>");
        //emailStr.Append(MakeHeaderStr(masterRow[mStatementmessageline1].ToString()));
        emailStr.Append(@"<tr><td width=""99%"" colSpan=""6"" height=""21"">" + statMessageMonthly + "</td></tr>");

        //emailStr.Append(@"</td></tr>");
        //emailStr.Append(@"<tr><td>");
        //emailStr.Append(MakeHeaderStr(masterRow[mStatementmessageline2].ToString()));
        //emailStr.Append(@"</td></tr>");
        //emailStr.Append(@"<tr><td>");
        //emailStr.Append(MakeHeaderStr(masterRow[mStatementmessageline3].ToString()));
        //emailStr.Append(@"</td></tr>");
        emailStr.Append(@"</table></td></tr></table></td></tr>");

        emailStr.Append(@"<tr><td width='99%' height='22' colspan='2'>&nbsp;</td></tr>");

        emailStr.Append(@"<tr><td width='98%' colspan='2'><table border='1' width='100%' id='table14'>");
        emailStr.Append(@"<tr><td align='center' width='101'><b>");
        emailStr.Append(@"<font size='1'>Opening Balance</font>");
        emailStr.Append(@"</b></td>");
        emailStr.Append(@"<td align='center'><b>");
        emailStr.Append(@"<font size='1'>Payments</font>");
        emailStr.Append(@"</b></td>");
        emailStr.Append(@"<td align='center' width='108'><b>");
        emailStr.Append(@"<font size='1'>Cash &amp; Purchases</font>");
        emailStr.Append(@"</b></td>");
        emailStr.Append(@"<td align='center' width='76'><b>");
        emailStr.Append(@"<font size='1'>Interest</font>");
        emailStr.Append(@"</b></td>");
        emailStr.Append(@"<td align='center' width='70'><b>");
        emailStr.Append(@"<font size='1'>Charges</font>");
        emailStr.Append(@"</b></td>");
        emailStr.Append(@"<td width='97' align='center'><b>");
        emailStr.Append(@"<font size='1'>Closing Balance</font>");
        emailStr.Append(@"</b></td>");

        emailStr.Append(@"<td width='85' align='center'><b>");
        emailStr.Append(@"<font size='1'>Total Debit(s)</font>");
        emailStr.Append(@"</b></td>");
        emailStr.Append(@"<td width='82' align='center'><b>");
        emailStr.Append(@"<font size='1'>Total Credit(s)</font>");
        emailStr.Append(@"</b></td>");
        emailStr.Append(@"</tr>");

        emailStr.Append(@"<tr><td align='center' width='101'>");
        emailStr.Append(MakeHeaderStr(basText.alignmentRight(basText.formatNumUnSign(masterRow[mOpeningbalance], "#,###,##0.00", 12) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance].ToString())), 15)));
        emailStr.Append(@"</td>");
        emailStr.Append(@"<td align='center'>");
        emailStr.Append(MakeHeaderStr(basText.alignmentRight(basText.formatNumUnSign(masterRow[mTotalpayments], "#,###,##0.00", 12) + " " + DbCr(Convert.ToDecimal(masterRow[mTotalpayments])), 15)));
        emailStr.Append(@"</td>");
        emailStr.Append(@"<td align='center' width='108'>");
        emailStr.Append(MakeHeaderStr(basText.alignmentRight(basText.formatNum(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal]), "#,###,##0.00", 12) + CrDbMinus(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal])), 15)));
        emailStr.Append(@"</td>");
        emailStr.Append(@"<td align='center' width='76'>");
        emailStr.Append(MakeHeaderStr(basText.alignmentRight(basText.formatNumUnSign(masterRow[mTotalinterest].ToString(), "#,###,##0.00", 12) + " " + CrDbMinus(Convert.ToDecimal(masterRow[mTotalinterest])), 15)));
        emailStr.Append(@"</td>");
        emailStr.Append(@"<td align='center' width='70'>");
        emailStr.Append(MakeHeaderStr(basText.alignmentRight(basText.formatNumUnSign(masterRow[mTotalcharges].ToString(), "#,###,##0.00", 12) + " " + CrDbMinus(Convert.ToDecimal(masterRow[mTotalcharges])), 15)));
        emailStr.Append(@"</td>");
        emailStr.Append(@"<td width='97' align='center'>");
        emailStr.Append(@"<font size='1' color='#800000'>");
        emailStr.Append(MakeHeaderStr(basText.alignmentRight(basText.formatNumUnSign(masterRow[mClosingbalance], "#,###,##0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])), 15)));
        emailStr.Append(@"</font></td>");

        emailStr.Append(@"<td align='center' width='85'>");
        emailStr.Append(MakeHeaderStr(basText.formatNum(masterRow[mTotaldebits], "#0.00;(#0.00)")));
        emailStr.Append(@"</td>");
        emailStr.Append(@"<td align='center' width='82'>");
        emailStr.Append(MakeHeaderStr(basText.formatNum(masterRow[mTotalcredits], "#0.00;(#0.00)")));
        emailStr.Append(@"</td>");
        emailStr.Append(@"</tr>");
        //emailStr.Append(@"</table></td></tr></table></body></html>");
        emailStr.Append(@"</table></td></tr>");
        emailStr.Append(@"<tr><td width='98%' colspan='2'>&nbsp;</td></tr>");

        if (isRewardVal)
        {
            rewardRows = DSreward.Tables["Reward"].Select("CLIENTID = '" + clsBasValid.validateStr(masterRow[ClientID]).ToString().Trim() + "'");
            if (rewardRows.Length > 0)
            {
                rewardRow = rewardRows[0];
                emailStr.Append(@"<tr><td width='98%' colSpan='2'><table id='table25' width='100%' border='1'>");
                emailStr.Append(@"<tr><td align='middle' width='171'>");
                emailStr.Append(MakeHeaderStr("Points B/F from last month", true));
                emailStr.Append(@"</td><td align='middle'>");
                emailStr.Append(MakeHeaderStr("Qualifying spend this month", true));
                emailStr.Append(@"</td><td align='middle' width='122'>");
                emailStr.Append(MakeHeaderStr("New Points Added", true));
                emailStr.Append(@"</td><td align='middle' width='121'>");
                emailStr.Append(MakeHeaderStr("Points Redeemed", true));
                emailStr.Append(@"</td><td align='middle' width='136'>");
                emailStr.Append(MakeHeaderStr("Total Carried Forward", true));
                emailStr.Append(@"</td></tr>");

                emailStr.Append(@"<tr><td align='middle' width='171'>");
                emailStr.Append(MakeHeaderStr(basText.formatNum(rewardRow[mOpeningbalance], "#0.00;(#0.00)", 20, "M"), false));
                emailStr.Append(@"</td><td align='middle'>");
                emailStr.Append(MakeHeaderStr("0", false));
                emailStr.Append(@"</td><td align='middle' width='122'>");
                emailStr.Append(MakeHeaderStr(basText.formatNum(rewardRow["EARNEDBONUS"], "#0.00;(#0.00)", 20, "M"), false));
                emailStr.Append(@"</td><td align='middle' width='121'>");
                emailStr.Append(MakeHeaderStr(basText.formatNum(rewardRow["REDEEMEDBONUS"], "#0.00;(#0.00)", 20, "M"), false));
                emailStr.Append(@"</td><td align='middle' width='136'>");
                emailStr.Append(MakeHeaderStr(basText.formatNum(rewardRow[mClosingbalance], "#0.00;(#0.00)", 20, "M"), false));
                emailStr.Append(@"</td></tr></table></td></tr>");
                emailStr.Append(@"<tr><td width='98%' colSpan='2'>&nbsp;</td></tr>");

            }
        }
        emailStr.Append(@"<tr><td width='98%' colspan='2'><p align='right'><a href='http://" + strbankWebLink + @"'>" + strbankWebLink + "</a></td></tr>");
        emailStr.Append(@"<tr><td width='98%' colSpan='6'>&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td width='98%' colSpan='6'>&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td width='98%' colSpan='6'>&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td width='98%' colSpan='6'>&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"">" + clsCheckSum.stringMD5(extAccNum.Trim() + Convert.ToDateTime(masterRow[mStatementdateto]).ToString("dd/MM/yyyy")) + "</td></tr>");
        emailStr.Append(@"<tr><td width='98%' colSpan='6'>&nbsp;</td></tr>");
        emailStr.Append(@"</table>");
        emailStr.Append(@"</body></html>");

    }

    //protected void printAccountFooter()
    //    {
    //    streamWrit.WriteLine("GRAND" + basText.formatNumUnSign(masterRow[mOpeningbalance], "##0.00", 16) + basText.formatNum(masterRow[mTotaldebits], "#0.00;(#0.00)", 16) + basText.formatNum(masterRow[mTotalcredits], "#0.00;(#0.00)", 16) + basText.formatNum(masterRow[mClosingbalance], "#0.00;(#0.00)", 16) + basText.formatNum(masterRow[mMindueamount], "#0.00;(#0.00)", 16) + basText.replicat(" ", 12) + basText.formatDate(masterRow[mStetementduedate], "yyyy/MM/dd") + "\n");//+strEndOfPage
    //    }


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
        bool haveValidCardNumber = false;

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
        if (createCorporateVal)
            mainRows = DSstatement.Tables["tStatementMasterTable"].Select("CARDACCOUNTNO = '" + clsBasValid.validateStr(masterRow[accountNoName]) + "'");
        else
            mainRows = DSstatement.Tables["tStatementMasterTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[accountNoName]) + "'");
        
        foreach (DataRow mainRow in mainRows) //mRow.GetChildRows(StatementNoDRel)
        {
            totCrdNoInAcc++;
            if (mainRow[mCardprimary].ToString() == "Y" && isValidateCard(mainRow[mCardstate].ToString()))
            {
                curMainCard = mainRow[mCardno].ToString();
                haveValidCardNumber = true;
                break;
            }
        }

        if (!haveValidCardNumber)
        {
            DataRow latestDataRow = mainRows.OrderByDescending(row => row["CARDCREATIONDATE"]).FirstOrDefault();
            curMainCard = latestDataRow[mCardno].ToString();
        }

       
    }


    private void completePageDetailRecords()
    {
        if (pageNo == totalAccPages)
            for (int curPageLine = CurPageRec4Dtl; curPageLine < MaxDetailInPage; curPageLine++)
                streamWrit.WriteLine(String.Empty);
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
            pStr = @"<font size=""1"">" + pStr + "</font>";

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
                streamNoEmails.WriteLine(curCustomerName + "|" + curAccountNumber + "|" + curClientID + "|" + pTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|" + clsCnfg.readSetting("strValidEmail") ?? "Bad Email");
                noOfBadEmails++;
                pTo = strEmailFrom; //"statement_Program@emp-group.com";
                emailLabelTmp = emailLabel + " Acc:" + curAccountNumber + " Phone:" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!");
            }

            //try
            //    {

            //streamEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
            ArrayList pLstTo = new ArrayList(), pLstAttachedFile = new ArrayList(), pLstAttachedPic = new ArrayList();//, pLstCC = new ArrayList(), pLstBCC = new ArrayList();
            pSubject = emailLabelTmp; //emailLabel "BAI statement for 02/2008";
            //pTo = emailTo;// "mmohammed@emp-group.com";
            if (pTo.EndsWith("."))//ehimolen@yahoo.com
                pTo = pTo + "com";
            pLstTo.Add(pTo);//"mmohammed@emp-group.com" "mhrap@hotmail.com""mhrap@yahoo.com"    "jedkosi@gmail.com"   "nazab@emp-group.com"

            //pLstAttachedFile.Add(strBankLogo);//strOutputFile@"D:\Web\Email\BAI\Logo.gif" //SIBN-2604
            //pLstAttachedFile.Add(strVisaLogo);//@"D:\pC#\exe\FilesForPrograms\frmBackground.jpg" //SIBN-2604

            //if (IsSplitted)
            //pLstAttachedPic.Add(strbottomBanner); //SIBN-2604

            //SIBN-2604
            if (HasAttachement)
                pLstAttachedFile.Add(stmtfilenameold);
            //    {
            //    foreach (string fileName in Directory.GetFiles(attachedFilesStr))
            //        {
            //        if (clsBasFile.getFileExtn(fileName) == "pdf")
            //            {
            //            pLstAttachedFile.Add(fileName);
            //            }
            //        }
            //    }

            sndMail = new clsEmail();
            sndMail.emailFromName = emailFromNameStr;
            strEmailFromTmp = strEmailFrom;


            numOfTry = 0;
            noOfEmails++;

            if (numMail == 0)
            {
                pLstBCC.Add("statement@emp-group.com");
                ////SIBN-3363
                //pLstBCC.Add("Ayodele.ojuroye@stanbicibtc.com");
                //pLstBCC.Add("Ademola.adeniran@stanbicibtc.com");
                //pLstBCC.Add("Adeyemo.alakija@stanbicibtc.com");
                //pLstBCC.Add("Olosen.evien@stanbicibtc.com");
                //pLstBCC.Add("Ezekiel.Williams-Osula@stanbicibtc.com");
                //pLstBCC.Add("Itohan.iyalla@stanbicibtc.com");
                //pLstBCC.Add("Oluwatobi.boshoro@stanbicibtc.com");
                //pLstBCC.Add("CARDSUNIT@stanbicibtc.com");
            }
            else
            {
                pLstBCC.Remove("statement@emp-group.com");
                ////SIBN-3363
                //pLstBCC.Remove("Ayodele.ojuroye@stanbicibtc.com");
                //pLstBCC.Remove("Ademola.adeniran@stanbicibtc.com");
                //pLstBCC.Remove("Adeyemo.alakija@stanbicibtc.com");
                //pLstBCC.Remove("Olosen.evien@stanbicibtc.com");
                //pLstBCC.Remove("Ezekiel.Williams-Osula@stanbicibtc.com");
                //pLstBCC.Remove("Itohan.iyalla@stanbicibtc.com");
                //pLstBCC.Remove("Oluwatobi.boshoro@stanbicibtc.com");
                //pLstBCC.Remove("CARDSUNIT@stanbicibtc.com");
            }

            while (!sndMail.sendEmailAttachPicFile(strEmailFromTmp, pLstTo, pLstCC, pLstBCC, pLstAttachedPic, pLstAttachedFile, pSubject, pBody, clsCnfg.readSetting("SmtpServer")))
            {
                System.Threading.Thread.Sleep(waitPeriodVal);//2000
                numOfTry++;
                if (numOfTry > 100)
                {
                    streamEmails.Write("\t\t Error while Send Email");
                    streamNoEmails.WriteLine(curCustomerName + "|" + curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|Exceed number of trials");
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

            //SIBN-2604
            if (HasAttachement)
            {
                //foreach (string fileName in Directory.GetFiles(attachedFilesStr))
                //    {
                //    if (clsBasFile.getFileExtn(fileName) == "pdf")
                //        {
                clsBasFile.moveFile(stmtfilenameold, stmtfilename);
                aryLstFilesStmt.Add(stmtfilename);
                //        }
                //    }
            }
            //sndMail = null;
            //emailStr = new StringBuilder(""); emailTo = string.Empty; curAccountNumber = string.Empty; curCardNumber = string.Empty; curClientID = string.Empty;
            //wait2NextEmail(prvEmail, pTo, waitPeriodVal);//     System.Threading.Thread.Sleep(waitPeriodVal);//400
            //prvEmail = pTo;
        }
        catch (Exception ex) //(NotSupportedException ex) (Exception ex)  //
        {
            //clsBasErrors.catchError(ex);
            streamNoEmails.WriteLine(curCustomerName + "|" + curAccountNumber + "|" + curClientID + "|" + pTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "| Email " + pTo + " , Err Message >> " + ex.Message + ", Err Desc >> ");
            noOfBadEmails++;
        }
        finally
        {
            sndMail = null;
            emailStr = new StringBuilder(""); 
            emailTo = string.Empty; 
            curAccountNumber = string.Empty; 
            curCardNumber = string.Empty; 
            curClientID = string.Empty;
            wait2NextEmail(prvEmail, pTo, waitPeriodVal);//     System.Threading.Thread.Sleep(waitPeriodVal);//400
            prvEmail = pTo;

        }
    }

    //private void SendEmailWithDifferentSender(string pBody, string pSubject, string pTo)
    //    {
    //    if (valdEmail.isValideEmail(pTo) != "ValidEmail")//!basText.isValideEmail(pTo)
    //        {
    //        streamNoEmails.WriteLine(curCustomerName + "|" + curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|"+clsCnfg.readSetting("strValidEmail")?? "Bad Email");
    //        noOfBadEmails++;
    //        emailTo = strEmailFrom; //"statement_Program@emp-group.com";
    //        emailLabelTmp = emailLabel + " Acc:" + curAccountNumber + " Phone:" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!");
    //        }

    //    try
    //        {

    //        streamEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
    //        ArrayList pLstTo = new ArrayList(), pLstAttachedFile = new ArrayList(), pLstAttachedPic = new ArrayList();//, pLstCC = new ArrayList(), pLstBCC = new ArrayList();
    //        pSubject = emailLabelTmp; //emailLabel "BAI statement for 02/2008";
    //        pTo = emailTo;// "mmohammed@emp-group.com";
    //        if (pTo.EndsWith("."))//ehimolen@yahoo.com
    //            pTo = pTo + "com";
    //        pLstTo.Add(pTo);//"mmohammed@emp-group.com" "mhrap@hotmail.com""mhrap@yahoo.com"    "jedkosi@gmail.com"   "nazab@emp-group.com"

    //        pLstAttachedFile.Add(strBankLogo);//strOutputFile@"D:\Web\Email\BAI\Logo.gif"
    //        pLstAttachedFile.Add(strVisaLogo);//@"D:\pC#\exe\FilesForPrograms\frmBackground.jpg"

    //        if (IsSplitted)
    //            pLstAttachedPic.Add(strbottomBanner);

    //        //if (HasAttachement)
    //        //    {
    //        //    foreach (string fileName in Directory.GetFiles(attachedFilesStr))
    //        //        {
    //        //        if (clsBasFile.getFileExtn(fileName) == "pdf")
    //        //            {
    //        //            pLstAttachedFile.Add(fileName);
    //        //            }
    //        //        }
    //        //    }

    //        sndMail = new clsEmail();
    //        sndMail.emailFromName = emailFromNameStr;
    //        strEmailFromTmp = strEmailFrom;


    //        numOfTry = 0;
    //        noOfEmails++;

    //        if (numMail == 0)
    //            {
    //            pLstBCC.Add("statement@emp-group.com");
    //            }
    //        else
    //            {
    //            pLstBCC.Remove("statement@emp-group.com");
    //            }

    //        while (!sndMail.sendEmailAttachPicFilewithDiffernetSender(strEmailFromTmp, "statement@emp-group.com", pLstTo, pLstCC, pLstBCC, pLstAttachedPic, pLstAttachedFile, pSubject, pBody, clsCnfg.readSetting("SmtpServer")))
    //            {
    //            System.Threading.Thread.Sleep(waitPeriodVal);//2000
    //            numOfTry++;
    //            if (numOfTry > 100)
    //                {
    //                streamEmails.Write("\t\t Error while Send Email");
    //                streamNoEmails.WriteLine(curCustomerName + "|" + curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|Exceed number of trials");
    //                noOfBadEmails++; noOfEmails--;
    //                break;
    //                }
    //            }
    //        numMail++;
    //        if (numMail % 400 == 0)
    //            {
    //            System.Threading.Thread.Sleep(waitPeriodVal);//2000
    //            GC.Collect();
    //            GC.WaitForPendingFinalizers();
    //            }

    //        //if (HasAttachement)
    //        //{
    //        //foreach (string fileName in Directory.GetFiles(attachedFilesStr))
    //        //    {
    //        //    if (clsBasFile.getFileExtn(fileName) == "pdf")
    //        //        {
    //        //        clsBasFile.moveFile(fileName, attachedFilesStr + "//STMT//" + clsBasFile.getFileFromPath(fileName));
    //        //        aryLstFilesStmt.Add(stmtfilename);
    //        //        }
    //        //    }
    //        //}
    //        sndMail = null;
    //        emailStr = new StringBuilder(""); emailTo = string.Empty; curAccountNumber = string.Empty; curCardNumber = string.Empty; curClientID = string.Empty;
    //        wait2NextEmail(prvEmail, pTo, waitPeriodVal);//     System.Threading.Thread.Sleep(waitPeriodVal);//400
    //        prvEmail = pTo;
    //        }
    //    catch  //(NotSupportedException ex) (Exception ex)  //
    //        {
    //        //clsBasErrors.catchError(ex);
    //        streamNoEmails.WriteLine(curCustomerName + "|" + curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|"+clsCnfg.readSetting("strValidEmail")?? "Bad Email");
    //        noOfBadEmails++;
    //        }
    //    finally
    //        {
    //        }

    //    }

    private void add2FileList(string pFileName)
    {
        int myIndex = aryLstFiles.BinarySearch((object)pFileName);
        if (myIndex < 0)
            aryLstFiles.Add(@pFileName);
    }

    protected virtual void wait2NextEmail(string pPrvEmail, string pNxtEmail, int pWaitPeriod)
    {
        try
        {
            if (pPrvEmail.IndexOf("@") > 0 && pNxtEmail.IndexOf("@") > 0)
                if (pPrvEmail.ToLower().Substring(pPrvEmail.IndexOf("@")) == pNxtEmail.ToLower().Substring(pNxtEmail.IndexOf("@")))
                    System.Threading.Thread.Sleep(waitPeriodVal);//7000
                else
                    System.Threading.Thread.Sleep(1000);
        }
        catch (Exception)
        {
            
        }
        
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

    public bool HasSender
    {
        get { return HasSenderVal; }
        set { HasSenderVal = value; }
    }  // HasSender

    ~clsStatHtmlSIBN()
    {
        DSstatement.Dispose();
    }
}
