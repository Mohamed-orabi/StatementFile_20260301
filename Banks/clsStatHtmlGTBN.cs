using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Xml;
using System.Collections;


public class clsStatHtmlGTBN : clsBasStatement
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
    protected DataRow masterRelatedRow;

    private string CurrentPageFlag;
    private string strCardNo, strPrimaryCardNo;
    private string strForeignCurr;
    private string stmNo, stmtNo, stmtfilename;
    private int totNoOfCardStat, totNoOfPageStat;

    private int curCardRow = 0, curTotCardRows = 0, numOfErr = 0, totNoOfTransactions = 0;
    private bool isPrimaryOnly, isHaveF3 = true, isPrimaryOpened = false;
    private FileStream fileStrmErr;
    private StreamWriter strmWriteErr;
    private string curMainCard;
    protected string isSuplStr = string.Empty;//Null

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
    private ArrayList pLstCC = new ArrayList(), pLstBCC = new ArrayList(), pLstAttachedPic = new ArrayList();
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

    protected string fontTypeSize4Header = "<font color='#808080' size='2'>";
    protected string fontTypeSize = "<font color='#808080' size='2'>";
    protected string fontTypeSize2 = "<font color='#8B0000' size='2'>";
    protected string fontTypeSize3 = "<font color='#8B0000' size='2'>";
    protected string lineSeparator = "<tr><td width='98%' height='25' style='width: 61%'>&nbsp;</td></tr>";

    private string strbottomBanner = @"D:\pC#\ProjData\Statement\GTBN\Banner.jpg";
    private string strTopBanner = @"D:\pC#\ProjData\Statement\GTBN\Banner.jpg";
    private string emailCc;

    public clsStatHtmlGTBN()
        {
        fontTypeSize4Header = "<font color='#000080' size='2'>";
        fontTypeSize = "<font color='#800000' size='2'>";
        logoAlignmentStr = "left";
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
            //clsBasFile.deleteDirectory(strOutputPath);
            clsBasFile.createDirectory(strOutputPath);
            //clsBasFile.createDirectory(strOutputPath + "//STMT");

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

            //attachedFilesStr = clsBasFile.getPathWithoutFile(strEmailFileName);

            // open output file
            //>fileStrm = new FileStream(pStrFileName, FileMode.Create); //Create
            //>streamWrit = new StreamWriter(fileStrm, Encoding.Default);
            //Error file
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
            //     clsBasStatement.mainTableCond =  " substr(cardno,1,6) !='421192' "; //" cardproduct != 'Visa Business' ";
            //     clsBasStatement.supTableCond =  " substr(cardno,1,6) !='421192' "; 
            // data retrieve
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

            //if (HasAttachement)
            //    {
            //    foreach (string fileName in Directory.GetFiles(strOutputPath))
            //        {
            //        if (clsBasFile.getFileExtn(fileName) == "pdf")
            //            {
            //            clsBasFile.deleteFile(fileName);
            //            }
            //        }
            //    }

            strOrder = " m.cardproduct,m.accountno,m.cardprimary desc,m.cardno ";
            FillStatementDataSet(pBankCode, ""); //DSstatement =  //10); //3
            getClientEmail(pBankCode);
            pageNo = 1; totalCardPages = 0;
            curCardNo = String.Empty;
            curAccountNo = String.Empty;
            //new frmGenerateSatatement().Show();
            frmMain.BeginInvoke(frmMain.setMinMaxProgressDelegate, new object[] { DSstatement.Tables["tStatementMasterTable"].Rows.Count });
            //generateStatus.Show();
            //generateStatus.setMinMaxProgress(DSstatement.Tables["tStatementMasterTable"].Rows.Count);
            //generateStatus.BeginInvoke(generateStatus.setMinMaxProgressDelegate, new object[] { DSstatement.Tables["tStatementMasterTable"].Rows.Count });
            foreach (DataRow mRow in DSstatement.Tables["tStatementMasterTable"].Rows)
                {
                frmMain.BeginInvoke(frmMain.setProgressDelegate, new object[] { totRec++ });
                //generateStatus.BeginInvoke(generateStatus.setProgressDelegate, new object[] { totRec++ });
                masterRow = mRow;
                //streamWrit.WriteLine(masterRow[mStatementno].ToString());
                //pageNo=1; CurPageRec4Dtl=0; CurrentPageFlag = "F 1"; //if page is based on card no
                //-cardsRows = mRow.GetChildRows(StatementNoDRel); //, DataRowVersion.Proposed

                strCardNo = masterRow[mCardno].ToString().Trim();

                //check product
                //-if (cProduct != masterRow[mCardproduct].ToString().Trim())
                //-{
                //-  if (cProduct != string.Empty)
                //-  {
                //-    streamWrit.Flush();
                //-    streamWrit.Close();
                //-    fileStrm.Close();
                //-  }
                //-  cProduct = masterRow[mCardproduct].ToString().Trim();
                //-  curFileName = clsBasFile.getPathWithoutExtn(pStrFileName)
                //-    + "_" + cProduct + "." + clsBasFile.getFileExtn(pStrFileName);
                //curFileName = pStrFileName;
                //-  add2FileList(curFileName);
                //-  fileStrm = new FileStream(curFileName, FileMode.Append); //FileMode.Create Create
                //-  streamWrit = new StreamWriter(fileStrm, Encoding.Default);
                //-}

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
                    //calcCardlRows();
                    }

                //start new account
                if (prevAccountNo != masterRow[accountNoName].ToString())
                    {
                    emailLabelTmp = emailLabel;
                    if (!string.IsNullOrEmpty(emailStr.ToString()) && !string.IsNullOrEmpty(emailTo)) //emailStr != string.Empty && emailStr != null emailTo != string.Empty && emailTo != null
                        if (!string.IsNullOrEmpty(emailStr.ToString()) && !string.IsNullOrEmpty(emailCc)) //emailStr != string.Empty && emailStr != null emailTo != string.Empty && emailTo != null
                            SendEmail(emailStr.ToString(), "", emailTo, emailCc);
                        else
                            SendEmail(emailStr.ToString(), "", emailTo, "");

                    if (!string.IsNullOrEmpty(emailStr.ToString()) && string.IsNullOrEmpty(emailTo)) //emailStr != string.Empty && emailStr != null  emailTo != string.Empty && emailTo != null
                        {
                        streamNoEmails.WriteLine(curCustomerName + "|" + curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|Without Email");
                        noOfWithoutEmails++;
                        emailTo = strEmailFrom; //"statement_Program@emp-group.com";
                        emailLabelTmp = emailLabel + " Acc:" + curAccountNumber + " Phone:" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!");
                        SendEmail(emailStr.ToString(), "", emailTo);
                        }

                    emailStr = new StringBuilder("");
                    cardsRows = DSstatement.Tables["tStatementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[accountNoName]).ToString().Trim() + "'");

                    //-if (pageNo != totalAccPages && prevAccountNo != "")// 
                    //-{
                    //>MessageBox.Show( "Error in Genrating Statement", "Please Call The Programmer", MessageBoxButtons.OK ,
                    //>  MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1); //,	MessageBoxOptions.RightAlign
                    //-  strmWriteErr.WriteLine("pageNo not equal totalAccPages : " + prevAccountNo);
                    //-  numOfErr++;
                    //-}

                    curMainCard = string.Empty;
                    //-if (!isHaveF3)//!isHaveF3  CurrentPageFlag != "F 2"
                    //-{
                    //-  strmWriteErr.WriteLine("Account not Closed : " + prevAccountNo);
                    //-  numOfErr++;
                    //-}
                    isHaveF3 = false;

                    pageNo = 1;//totalAccPages = 1 ; pageNo=1;
                    CurPageRec4Dtl = 0; totNetUsage = 0; CurrentPageFlag = "F 1"; //if page is based on account no
                    calcAccountRows();




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
                        //GTBN-1087
                        emailCc = emailRows[i][6].ToString().Trim();
                        }
                    curAccountNumber = masterRow[accountNoName].ToString();
                    curCardNumber = strPrimaryCardNo;
                    curClientID = masterRow[ClientID].ToString();
                    curCustomerName = masterRow[mCustomername].ToString();

                    //emailTo = (DSemails.Tables["Emails"].Select("idclient = " + masterRow[mClientid].ToString)).;
                    printHeader();//if page is based on account no

                    totNoOfCardStat++;
                    } // End of if(prevAccountNo != masterRow[accountNoName].ToString())
                else
                    {
                    continue;
                    }

                //calcCardlRows();
                //if(totCardRows < 1)continue ;  //if pages is based on card
                //printHeader();//if pages is based on card


                foreach (DataRow dRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
                    {
                    detailRow = dRow;
                    isSuplStr = string.Empty;
                    if (curMainCard != detailRow[dCardno].ToString().Trim())
                        {
                        masterRelatedRow = DSstatement.Tables["tStatementMasterTable"].Select("statementno = '" + clsBasValid.validateStr(detailRow[dStatementno]).ToString().Trim() + "'")[0];
                        isSuplStr = isSupplementCard(clsBasValid.validateStr(masterRelatedRow[mCardprimary].ToString()));
                        }

                    stmNo = detailRow[dStatementno].ToString();
                    stmtNo = detailRow[dStatementnumber].ToString();
                    if ((detailRow[dPostingdate] == DBNull.Value) && (detailRow[dDocno] == DBNull.Value)) continue;// Exclude On-Hold Transactions 
                    curAccRows++;
                    totNetUsage = calculateCrDb(totNetUsage, clsBasValid.validateNum(detailRow[dBilltranamount]), clsBasValid.validateStr(detailRow[dBilltranamountsign]));
                    CurPageRec4Dtl = CurPageRec4Dtl + 1;
                    printDetail();

                    } //end of detail foreach
                //printCardFooter();//if pages is based on card
                // if(masterRow[mCardprimary].ToString() == "Y" && (masterRow[mCardstate].ToString() == "Given" || masterRow[mCardstate].ToString() == "Embossed" || masterRow[mCardstate].ToString() == "New"))
                curCrdNoInAcc++;
                //-if ((curAccRows >= totAccRows && totAccRows != 0) || (totAccRows == 0 && curCrdNoInAcc == totCrdNoInAcc))
                //-{
                //-  completePageDetailRecords();
                printCardFooter();//if pages is based on account

                //printAccountFooter();
                CurPageRec4Dtl = 0; //>pageNo=1; if pages is based on account
                curAccRows = 0;
                strEmailFromTmp = emailLabelTmp = string.Empty;
                //-}
                //streamWrit.WriteLine(strEndOfPage);
                //pageNo=1; CurPageRec4Dtl=0; //if pages is based on card
                //completePageDetailRecords();
                //SendEmail(emailStr, "", "");
                //if (cardsRows.Length == 0)
                //    strmWriteErr.WriteLine("'" + masterRow[mAccountno].ToString() + "',|'" + masterRow[mExternalno].ToString() + "',");
                //if (stmtNo == null)
                //    stmtNo = masterRow[mStatementnumber].ToString();
                //clsStatement_ExportRpt exportRptGTBN = new clsStatement_ExportRpt();
                ////GTBN bug 17/10/2013
                ////Was
                ////exportRptGTBN.whereCond = "statementno = " + int.Parse(stmtNo);
                ////Become
                //exportRptGTBN.whereCond = "statementnumber = " + int.Parse(stmtNo);
                //exportRptGTBN.exportAttachment(pStrFileName, pBankCode, masterRow[mCardclientname].ToString() + "_" + masterRow[mAccountcurrency].ToString() + "_" + vCurDate.ToString("MMyyyy"), pReportName);
                //exportRptGTBN = null;
                //stmtNo = null;
                //stmtfilename = @clsBasFile.getPathWithoutFile(pStrFileName) + "\\STMT\\" + masterRow[mCardclientname].ToString().Replace('.', ' ').Replace('/', ' ').Replace(' ', '_') + "_" + masterRow[mAccountcurrency].ToString() + "_" + vCurDate.ToString("MMyyyy") + ".pdf";
                } //end of Master foreach
            emailLabelTmp = emailLabel;

            if (!string.IsNullOrEmpty(emailStr.ToString()) && !string.IsNullOrEmpty(emailTo)) //emailStr != string.Empty && emailStr != null emailTo != string.Empty && emailTo != null
                if (!string.IsNullOrEmpty(emailStr.ToString()) && !string.IsNullOrEmpty(emailCc)) //emailStr != string.Empty && emailStr != null emailTo != string.Empty && emailTo != null
                    SendEmail(emailStr.ToString(), "", emailTo, emailCc);
                else
                    SendEmail(emailStr.ToString(), "", emailTo, "");

            if (!string.IsNullOrEmpty(emailStr.ToString()) && string.IsNullOrEmpty(emailTo)) // emailStr != string.Empty && emailStr != null emailTo != string.Empty && emailTo != null
                {
                streamNoEmails.WriteLine(curCustomerName + "|" + curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|Without Email");
                noOfWithoutEmails++;
                emailTo = strEmailFrom; //"statement_Program@emp-group.com";
                emailLabelTmp = emailLabel + " Acc:" + curAccountNumber + " Phone:" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!");
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
            //streamWrit.Flush();
            //streamWrit.Close();
            //fileStrm.Close();

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
            //zip.createZip(aryLstFilesStmt, @clsBasFile.getPathWithoutFile(pStrFileName) + "\\" + pBankName + pStrFile + pStmntType + "_STMT" + vCurDate.ToString("yyyyMM") + ".zip", "");

            DSstatement.Dispose();
            }
        return rtrnStr;
        }


    protected void printHeader()
        {
        string newaddress1, newaddress2;
        clsMaintainData.fixAddress(curBranchVal, masterRow[mCustomeraddress1].ToString(), out newaddress1, out newaddress2);
        totPages++; endOfCustomer = "";
        extAccNum = clsBasValid.validateStr(masterRow[mExternalno]);
        if (extAccNum.Trim() == "")
            extAccNum = clsBasValid.validateStr(masterRow[accountNoName]);


        //\MakeEmailStr   MakeHeaderStr   emailStr
        emailStr = new StringBuilder("");
        emailStr.Append(@"<html><head><meta http-equiv='Content-Type' content='text/html; charset=windows-1252' />");
        emailStr.Append(@"<title>" + masterRow[mCustomername].ToString() + " - Statement" + @"</title>");
        emailStr.Append(@"</head><body background='cid:" + clsBasFile.getFileFromPath(strbackGround) + @"' topmargin='0' leftmargin='0' rightmargin='0' bottommargin='0'>");

        //emailStr.Append(@"<tr><td><p align='center'><a href='http://" + strbankWebLink + @"'><img border='0' src='cid:" + clsBasFile.getFileFromPath(strTopBanner) + @"' alt='http://" + strbankWebLink + @"'></a></p></td></tr>");
        emailStr.Append(@"<table id='table15' width='100%' border='0' cellpadding='0' height='100%'><tr><td colspan='2' height='82'><table id='table25' width='100%' border='0'>");
        emailStr.Append(@"<tr><td height='100%' width='30%'><p align='right'><a href='http://" + strbankWebLink + @"'><img border='0' src='cid:" + clsBasFile.getFileFromPath(strBankLogo) + @"' border='0'></a></p></td>");
        //emailStr.Append(@"<td height='100%' align='center'><font size='6'>Statement</font></td>");
        //emailStr.Append(@"<td height='100%'><a href='http://" + strbankWebLink + @"'><img border='0' src='cid:" + clsBasFile.getFileFromPath(strVisaLogo) + @"' align='right' /></a></td>");   
        emailStr.Append(@"</tr></table></td></tr>");

        emailStr.Append(@"<tr><td width='50%' height='100%'><table id='table16' height='100%' width='100%' border='1'><tr><td width='100%' height='100%'><table id='table29' bordercolor='#000000' width='49%' border='0' height='100%' cellspacing='0' cellpadding='0'>");
        emailStr.Append(@"<tr><td align='middle'>" + fontTypeSize + validateEmpty(masterRow[mCustomername].ToString()) + "</font></td></tr>");
        //emailStr.Append(@"<tr><td align='middle'>" + fontTypeSize + validateEmpty(masterRow[mCustomeraddress1].ToString()) + "</font></td></tr>");
        emailStr.Append(@"<tr><td align='middle'>" + fontTypeSize + validateEmpty(newaddress1) + "</font></td></tr>");
        //emailStr.Append(@"<tr><td align='middle'>" + fontTypeSize + validateEmpty(masterRow[mCustomeraddress2].ToString()) + "</font></td></tr>");
        emailStr.Append(@"<tr><td align='middle'>" + fontTypeSize + validateEmpty(newaddress2) + "</font></td></tr>");
        emailStr.Append(@"<tr><td align='middle'>" + fontTypeSize + validateEmpty(masterRow[mCustomeraddress3].ToString()) + "</font></td></tr>");
        emailStr.Append(@"<tr><td align='middle'>" + fontTypeSize + validateEmpty(masterRow[mCustomerregion].ToString().Trim() + " " + masterRow[mCustomercity].ToString().Trim()) + "</font></td></tr>");
        emailStr.Append(@"</table></td></tr></table></td>");
        emailStr.Append(@"<td width='380' height='49%'><table id='table18' height='100%' width='380' border='1'><tr><td width='380' height='100%'><table id='table28' height='100%' width='380' border='0' cellspacing='0' cellpadding='0'>");
        emailStr.Append(@"<tr><td width='107'><b>" + fontTypeSize4Header + "Card Product :</font></b></td>");
        emailStr.Append(@"<td><font size='2'>" + fontTypeSize + validateEmpty(masterRow[mCardproduct].ToString()) + "</font></td></tr>");
        //emailStr.Append(@"<tr><td width='107'><b>" + fontTypeSize4Header + "Branch :</font></b></td>");
        //emailStr.Append(@"<td>" + fontTypeSize + validateEmpty(mCardbranchpartname) + "</font></td></tr>");
        emailStr.Append(@"<tr><td width='107'><b>" + fontTypeSize4Header + "Acc. Number :</font></b></td>");
        emailStr.Append(@"<td>" + fontTypeSize + validateEmpty(extAccNum) + "</font></td></tr>");
        emailStr.Append(@"<tr><td width='107'><b>" + fontTypeSize4Header + "Statement Date:</font></b></td>");
        emailStr.Append(@"<td>" + fontTypeSize + validateEmpty(Convert.ToDateTime(masterRow[mStatementdateto]).ToString("dd/MM/yyyy")) + "</font></td></tr>");
        emailStr.Append(@"</table></td></tr></table></td></tr>");

        emailStr.Append(lineSeparator);
        emailStr.Append(@"<tr><td width='98%' colspan='2' height='77'><table id='table20' height='42' width='100%' border='1'>");
        emailStr.Append(@"<tr><td align='center' width='102'><b>" + fontTypeSize4Header + "Credit Limit</font></b></td>");
        emailStr.Append(@"<td align='center' width='126'><b>" + fontTypeSize4Header + "Available Limit</font></b></td>");
        emailStr.Append(@"<td align='center' width='156'><b>" + fontTypeSize4Header + "Min. Due Amount</font></b></td>");
        emailStr.Append(@"<td align='center' width='177'><b>" + fontTypeSize4Header + "Min. Due Date</font></b></td>");
        emailStr.Append(@"<td align='center' width='174'><b>" + fontTypeSize4Header + "Over Due Amount</font></b></td></tr>");
        emailStr.Append(@"<tr><td align='middle' width='102' height='16'>" + fontTypeSize + validateEmpty(masterRow[accountLimit].ToString().Trim()) + "</font></td>");
        //emailStr.Append(@"<td align='middle' width='126' height='16'>" + fontTypeSize + validateEmpty(basText.formatNum(masterRow[accountAvailableLimit], "##0", 20).Trim()) + "</font></td>");
        emailStr.Append(@"<td align='middle' width='126' height='16'>" + fontTypeSize + validateEmpty(basText.formatNum(masterRow[mAccountavailablelim], "##0", 20).Trim()) + "</font></td>");
        emailStr.Append(@"<td align='middle' width='156' height='16'><b>" + fontTypeSize + validateEmpty(masterRow[mMindueamount].ToString().Trim()) + "</font></b></td>");
        emailStr.Append(@"<td align='middle' width='177' height='16'>" + fontTypeSize + validateEmpty(basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 10, "M")) + "</font></td>");
        emailStr.Append(@"<td align='middle' width='174' height='16'>" + fontTypeSize + validateEmpty(basText.formatNum(masterRow[mTotaloverdueamount], "##0.00", 13).Trim()) + "</font></td></tr>");
        emailStr.Append(@"</table></td></tr>");
        emailStr.Append(lineSeparator);
        emailStr.Append(@"<tr><td width='100%' colspan='2' height='100%'><table id='table21' height='100%' width='100%' border='1'>");
        emailStr.Append(@"<tr><td align='middle' colspan='2' height='16'>" + fontTypeSize4Header + "Dates of</font></td>");
        emailStr.Append(@"<td align='middle' width='19%' height='34' rowspan='2'><b>" + fontTypeSize4Header + "Reference No</font></b></td>");
        emailStr.Append(@"<td align='middle' width='52%' colspan='2' height='34' rowspan='2'><b>" + fontTypeSize4Header + "Description</font></b></td>");
        //emailStr.Append(@"<td align='middle' width='13%' colspan='2' height='34' rowspan='2'><b>" + fontTypeSize4Header + "Amount ( NGN )</font></b></td></tr>");
        emailStr.Append(@"<td align='middle' width='13%' colspan='2' height='34' rowspan='2'><b>" + fontTypeSize4Header + "Amount ( USD )</font></b></td></tr>");
        emailStr.Append(@"<tr><td align='middle' width='6%' height='16'>" + fontTypeSize4Header + "Trans.</font></td>");
        emailStr.Append(@"<td align='middle' width='5%' height='16'>" + fontTypeSize4Header + "Posting</font></td></tr>");

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

        string trnsDesc = string.Empty; //= detailRow[dTrandescription].ToString().Trim() + " " + detailRow[dMerchant].ToString().Trim()
        if (detailRow[dMerchant].ToString().Trim() == "")
            trnsDesc = detailRow[dTrandescription].ToString().Trim();
        else
            trnsDesc = detailRow[dMerchant].ToString().Trim();

        emailStr.Append(@"<tr><td align='middle' width='6%' height='23'>" + fontTypeSize + validateEmpty(trnsDate.ToString("dd/MM")) + "</font></td>");
        emailStr.Append(@"<td align='middle' width='5%' height='23'>" + fontTypeSize + validateEmpty(postingDate.ToString("dd/MM")) + "</font></td>");
        emailStr.Append(@"<td align='middle' width='19%' height='23'>" + fontTypeSize + validateEmpty(detailRow[dRefereneno].ToString().Trim()) + "</font></td>");
        emailStr.Append(@"<td width='42%' height='23'><p align='left'>" + fontTypeSize + validateEmpty(trnsDesc.Trim()) + "</font></p></td>");
        emailStr.Append(@"<td align='right' width='10%' height='23'>" + fontTypeSize + validateEmpty(ValidateStr(strForeignCurr.Trim())) + "</font></td>");
        emailStr.Append(@"<td align='right' width='11%' height='23'>" + fontTypeSize + validateEmpty(basText.formatNumUnSign(detailRow[dBilltranamount], "#,###,##0.00", 14)) + "</font></td>");
        emailStr.Append(@"<td align='right' width='1%' height='23'>" + fontTypeSize + validateEmpty(Convert.ToString(CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])).ToString() + " " + isSuplStr).Trim()) + "</font></td></tr>");
        totNoOfTransactions++;
        }

    protected void printCardFooter()
        {
        emailStr.Append(@"</table></td></tr>");
        emailStr.Append(lineSeparator);
        emailStr.Append(@"<tr><td width='98%' colspan='2'><table id='table22' width='100%' border='1'><tr><td><table id='table23' width='100%' border='0'>");
        emailStr.Append(@"<tr><td align='center'>" + fontTypeSize + validateEmpty(ValidateStr(masterRow[mStatementmessageline1].ToString())) + "</font></td></tr>");
        emailStr.Append(@"<tr><td align='center'>" + fontTypeSize + validateEmpty(ValidateStr(masterRow[mStatementmessageline2].ToString())) + "</font></td></tr>");
        emailStr.Append(@"<tr><td align='center'>" + fontTypeSize + validateEmpty(ValidateStr(masterRow[mStatementmessageline3].ToString())) + "</font></td></tr>");
        emailStr.Append(@"</table></td></tr></table></td></tr>");
        emailStr.Append(lineSeparator);
        emailStr.Append(@"<tr><td width='98%' colspan='2'><table id='table24' width='100%' border='1'>");
        emailStr.Append(@"<tr><td align='middle' width='122'><b>" + fontTypeSize4Header + "Opening Balance</font></b></td>");
        emailStr.Append(@"<td align='middle'><b>" + fontTypeSize4Header + "Payments</font></b></td>");
        emailStr.Append(@"<td align='middle' width='168'><b>" + fontTypeSize4Header + "Cash &amp; Purchases</font></b></td>");
        emailStr.Append(@"<td align='middle' width='92'><b>" + fontTypeSize4Header + "Interest</font></b></td>");
        emailStr.Append(@"<td align='middle' width='63'><b>" + fontTypeSize4Header + "Charges</font></b></td>");
        emailStr.Append(@"<td align='middle' width='141'><b>" + fontTypeSize4Header + "Closing Balance</font></b></td></tr>");
        emailStr.Append(@"<tr><td align='middle' width='122' height='16'>" + fontTypeSize + validateEmpty(basText.alignmentRight(basText.formatNumUnSign(masterRow[mOpeningbalance], "#,###,##0.00", 12) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance].ToString())), 15)) + "</font></td>");
        emailStr.Append(@"<td align='middle' height='16'>" + fontTypeSize + validateEmpty(basText.alignmentRight(basText.formatNumUnSign(masterRow[mTotalpayments], "#,###,##0.00", 12) + " " + DbCr(Convert.ToDecimal(masterRow[mTotalpayments])), 15)) + "</font></td>");
        emailStr.Append(@"<td align='middle' width='168' height='16'>" + fontTypeSize + validateEmpty(basText.alignmentRight(basText.formatNum(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal]), "#,###,##0.00", 12) + CrDbMinus(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal])), 15)) + "</font></td>");
        emailStr.Append(@"<td align='middle' width='92' height='16'>" + fontTypeSize + validateEmpty(basText.alignmentRight(basText.formatNumUnSign(masterRow[mTotalinterest].ToString(), "#,###,##0.00", 12) + " " + CrDbMinus(Convert.ToDecimal(masterRow[mTotalinterest])), 15)) + "</font></td>");
        emailStr.Append(@"<td align='middle' width='63' height='16'>" + fontTypeSize + validateEmpty(basText.alignmentRight(basText.formatNumUnSign(masterRow[mTotalcharges].ToString(), "#,###,##0.00", 12) + " " + CrDbMinus(Convert.ToDecimal(masterRow[mTotalcharges])), 15)) + "</font></td>");
        emailStr.Append(@"<td align='middle' width='141' height='16'>" + fontTypeSize + validateEmpty(basText.alignmentRight(basText.formatNumUnSign(masterRow[mClosingbalance], "#,###,##0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])), 15)) + "</font></td></tr>");
        emailStr.Append(@"</table></td></tr>");
        emailStr.Append(lineSeparator); //seprator
        if (isRewardVal)
            {
            rewardRows = DSreward.Tables["Reward"].Select("CLIENTID = '" + clsBasValid.validateStr(masterRow[mClientid]).ToString().Trim() + "'");
            if (rewardRows.Length > 0)
                {
                rewardRow = rewardRows[0];
                emailStr.Append(@"<tr><td width='98%' colSpan='2'><table id='table25' width='100%' border='1'>");
                emailStr.Append(@"<tr><td align='middle' width='171'>" + fontTypeSize4Header + "Points B/F from last month</font></td>");
                emailStr.Append(@"<td align='middle'>" + fontTypeSize4Header + "Qualifying spend this month</font></td>");
                emailStr.Append(@"<td align='middle' width='122'>" + fontTypeSize4Header + "New Points Added</font></td>");
                emailStr.Append(@"<td align='middle' width='121'>" + fontTypeSize4Header + "Points Redeemed</font></td>");
                emailStr.Append(@"<td align='middle' width='136'>" + fontTypeSize4Header + "Total Carried Forward</font></td></tr>");

                emailStr.Append(@"<tr><td align='middle' width='171'>" + fontTypeSize + validateEmpty(basText.formatNum(rewardRow[mOpeningbalance], "#0.00;(#0.00)", 20, "M")) + @"</font></td>");
                emailStr.Append(@"<td align='middle'>" + fontTypeSize + validateEmpty(basText.formatNum(rewardRow["EARNEDBONUS"], "#0.00;(#0.00)", 20, "M")) + @"</font></td>");
                emailStr.Append(@"<td align='middle' width='122'>" + fontTypeSize + validateEmpty(basText.formatNum(rewardRow["REDEEMEDBONUS"], "#0.00;(#0.00)", 20, "M")) + @"</font></td>");
                emailStr.Append(@"<td align='middle' width='121'>" + fontTypeSize + validateEmpty(basText.formatNum((rewardRow["bonusadjustment"].ToString() != "" ? Convert.ToDecimal(rewardRow["bonusadjustment"].ToString()) : Convert.ToDecimal(0.0)) * -1, "#0.00", 20, "M")) + @"</font></td>");
                emailStr.Append(@"<td align='middle' width='136'>" + fontTypeSize + validateEmpty(basText.formatNum(rewardRow[mClosingbalance], "#0.00;(#0.00)", 20, "M")) + @"</font></td></tr></table></td></tr>");
                }
            }
        emailStr.Append(lineSeparator); //seprator
        emailStr.Append(lineSeparator);//Blank section
        emailStr.Append(@"<tr><td width='769' colspan='3'><p align='right'><font face='Verdana'>");
        emailStr.Append(@"<a title='http://" + strbankWebLink + @"' href='http://" + strbankWebLink + @"'>http://" + strbankWebLink + @"</a></font></p></td></tr>");
        //emailStr.Append(@"<a title='http://" + strbankWebLink + @"' href='http://" + strbankWebLink + @"'>http://" + strbankWebLink + @"<img border='0' src='cid:" + clsBasFile.getFileFromPath(strbottomBanner) + @"' alt='http://" + strbankWebLink + @"'></a></font></p></td></tr>");

        clientEmailChecksum();

        emailStr.Append(@"</table></body></html>");
        }

    protected void printAccountFooter()
        {
        streamWrit.WriteLine("GRAND" + basText.formatNumUnSign(masterRow[mOpeningbalance], "##0.00", 16) + basText.formatNum(masterRow[mTotaldebits], "#0.00;(#0.00)", 16) + basText.formatNum(masterRow[mTotalcredits], "#0.00;(#0.00)", 16) + basText.formatNum(masterRow[mClosingbalance], "#0.00;(#0.00)", 16) + basText.formatNum(masterRow[mMindueamount], "#0.00;(#0.00)", 16) + basText.replicat(" ", 12) + basText.formatDate(masterRow[mStetementduedate], "yyyy/MM/dd") + "\n");//+strEndOfPage
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

    private void SendEmail(string pBody, string pSubject, string pTo, string pCc = "")
        {
            try
            {

                if (valdEmail.isValideEmail(pTo) != "ValidEmail")//!basText.isValideEmail(pTo)
                {
                    //streamNoEmails.WriteLine(curCustomerName + "|" + curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!")) + "|"+clsCnfg.readSetting("strValidEmail")?? "Bad Email");
                    //GTBN-1087
                    //streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailCc + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|Bad Additional Email");
                    streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + pTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|"+clsCnfg.readSetting("strValidEmail")?? "Bad Email " + pTo);

                    noOfBadEmails++;
                    pTo = strEmailFrom; //"statement_Program@emp-group.com";
                    emailLabelTmp = emailLabel + " Acc:" + curAccountNumber + " Phone:" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!");
                }

                if (!string.IsNullOrEmpty(pCc) && valdEmail.isValideEmail(pCc) != "ValidEmail")//!basText.isValideEmail(pTo)
                {
                    streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + pCc + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|Bad Additional Email (Cc) " + pCc);
                    noOfBadEmails++;
                    pCc = string.Empty;

                    //emailTo = strEmailFrom; //"statement_Program@emp-group.com";
                    //emailLabelTmp = emailLabel + " Acc:" + curAccountNumber + " Phone:" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + " Bad additional Email";
                    //return;
                }

        //try
        //    {

            //streamEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
            ////GTBN-1087
            //streamEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailCc + "|" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));

            ArrayList pLstTo = new ArrayList(), pLstAttachedFile = new ArrayList(), pLstCc = new ArrayList(), pLstAttachedPic = new ArrayList();//, pLstCC = new ArrayList(), pLstBCC = new ArrayList();
            pSubject = emailLabelTmp; //emailLabel "BAI statement for 02/2008";
            //pTo = emailTo;// "mmohammed@emp-group.com";

            if (pTo.EndsWith("."))//ehimolen@yahoo.com
                pTo = pTo + "com";
            pLstTo.Add(pTo);//"mmohammed@emp-group.com" "mhrap@hotmail.com""mhrap@yahoo.com"    "jedkosi@gmail.com"   "nazab@emp-group.com"

            //GTBN-1087
            //pCc = emailCc;
            if (pCc.EndsWith("."))//ehimolen@yahoo.com
                pCc = pCc + "com";

            if (!(pCc == string.Empty))
                pLstCc.Add(pCc);//

            pLstAttachedFile.Add(strBankLogo);//strOutputFile@"D:\Web\Email\BAI\Logo.gif"
            //pLstAttachedFile.Add(strVisaLogo);//@"D:\pC#\exe\FilesForPrograms\frmBackground.jpg"

            //if (HasAttachement)
            //    {
            //    foreach (string fileName in Directory.GetFiles(attachedFilesStr))
            //        {
            //        if (clsBasFile.getFileExtn(fileName) == "pdf")
            //            {
            //            pLstAttachedFile.Add(fileName);
            //            }
            //        }
            //    }

            //if (HasAttachement)
            //    foreach (string fileName in Directory.GetFiles(attachedFilesStr))
            //        {
            //        pLstAttachedFile.Add(fileName);
            //        }
            //pLstAttachedFile.Add(strbottomBanner);
            //pLstAttachedFile.Add(strTopBanner);
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

            if (strEmailFromTmp.ToUpper().EndsWith("GTBANK.COM"))
                {
                //if (pLstCc.Count == 0)
                if (pLstCc.Contains("statement@emp-group.com") == false)
                    pLstCc.Add("Creditcardteam@gtbank.com");
                //pLstCc.Add("statement@emp-group.com");
                }

            while (!sndMail.sendEmailAttachPicFile(strEmailFromTmp, pLstTo, pLstCc, pLstBCC, pLstAttachedPic, pLstAttachedFile, pSubject, pBody, clsCnfg.readSetting("SmtpServer")))
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
            streamEmails.WriteLine(curAccountNumber + "|" + curClientID + "|To " + pTo + "; Cc " + pCc + "|" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));

            //streamEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
            ////GTBN-1087
            //streamEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailCc + "|" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));

            numMail++;
            if (numMail % 400 == 0)
                {
                System.Threading.Thread.Sleep(waitPeriodVal);//2000
                GC.Collect();
                GC.WaitForPendingFinalizers();
                }

            //if (HasAttachement)
            //    {
            //    foreach (string fileName in Directory.GetFiles(attachedFilesStr))
            //        {
            //        if (clsBasFile.getFileExtn(fileName) == "pdf")
            //            {
            //            clsBasFile.moveFile(fileName, attachedFilesStr + "//STMT//" + clsBasFile.getFileFromPath(fileName));
            //            aryLstFilesStmt.Add(stmtfilename);
            //            }
            //        }
            //    }
            //sndMail = null;
            //emailStr = new StringBuilder(""); emailTo = string.Empty; curAccountNumber = string.Empty; curCardNumber = string.Empty; curClientID = string.Empty;
            //wait2NextEmail(prvEmail, pTo, waitPeriodVal);//     System.Threading.Thread.Sleep(waitPeriodVal);//400
            //prvEmail = pTo;
            }
            catch (Exception ex)  //(NotSupportedException ex) (Exception ex)  //
            {
            //clsBasErrors.catchError(ex);
                streamNoEmails.WriteLine(curCustomerName + "|" + curAccountNumber + "|" + curClientID + "|" + pTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "| Email " + pTo + " , Err Message >> " + ex.Message + ", Err Desc >> ");
            noOfBadEmails++;
            }
        finally
            {
                sndMail = null;
                emailStr = new StringBuilder(""); emailTo = string.Empty; emailCc = string.Empty; curAccountNumber = string.Empty; curCardNumber = string.Empty; curClientID = string.Empty;
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

    protected void clientEmailChecksum()
        {
        emailStr.Append(lineSeparator);
        emailStr.Append(lineSeparator);
        emailStr.Append(lineSeparator);
        emailStr.Append(@"<tr><td width='98%' colSpan='6'><font size='1'>" + clsCheckSum.stringMD5(extAccNum.Trim() + Convert.ToDateTime(masterRow[mStatementdateto]).ToString("dd/MM/yyyy")) + "</font></td></tr>");
        emailStr.Append(lineSeparator);
        }

    protected string validateEmpty(string pStr)
        {
        if (pStr.Length < 1)
            pStr = "&nbsp;";
        else
            pStr = pStr.Trim();

        return pStr;
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

    public string BottomBanner
        {
        set
            {
            strbottomBanner = value;
            pLstAttachedPic.Add(strbottomBanner);
            }
        } // BottomBanner

    public string TopBanner
        {
        set
            {
            strTopBanner = value;
            pLstAttachedPic.Add(strTopBanner);
            }
        } // TopBanner

    ~clsStatHtmlGTBN()
        {
        DSstatement.Dispose();
        }
    }
