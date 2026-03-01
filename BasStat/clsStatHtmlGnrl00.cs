using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Xml;
using System.Collections;


public class clsStatHtmlGnrl00 : clsBasStatement
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
    private string stmNo;
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
    private ArrayList aryLstFiles;

    //
    //private string emailStr;
    private StringBuilder emailStr = new StringBuilder("");
    private string emailLabel, strBankLogo;
    private string strEmailFrom = "mabouleila@emp-group.com";   //"cardservices@zenithbank.com"
    private string strbankWebLink = "www.emp-group.com";  //www.zenithbank.com
    private string emailTo, emailCc, curAccountNumber, curCardNumber, curClientID;//, curEmail
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

    private string statMessageFileMonthlyVal = string.Empty;
    private string statMessageMonthly = "&nbsp;";//Null

    private string strExcludeCond = string.Empty;
    protected bool isExcludedVal = false;
    private bool HasSenderVal = false;

    public clsStatHtmlGnrl00()
        {
        }

    public string Statement(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pCurDate, string pStmntType, bool pAppendData)
        {
        string rtrnStr = "Successfully Generate " + pBankName;
        int curMonth = pCurDate.Month;
        bool preExit = true;
        strFileNam = pStrFileName;
        stmntType = pStmntType;
        curMonth = pCurDate.Month;
        aryLstFiles = new ArrayList();
        try
            {
            //clsMaintainData maintainData = new clsMaintainData();
            //maintainData.matchCardBranch4Account(pBankCode);

            pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);
            vCurDate = pCurDate; //DateTime.Now.AddMonths(-1);
            strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "_" + pStmntType;
            //clsBasFile.deleteDirectory(strOutputPath);
            clsBasFile.createDirectory(strOutputPath);
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
            streamNoEmails.WriteLine("AccountNumber" + "|" + "ClientID" + "|" + "Email" + "|" + "Mobile Phone");
            streamNoEmails.AutoFlush = true;

            // open output file
            //>fileStrm = new FileStream(pStrFileName, FileMode.Create); //Create
            //>streamWrit = new StreamWriter(fileStrm, Encoding.Default);
            // Error file
            //-fileStrmErr = new FileStream(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName), FileMode.Create);
            //-strmWriteErr = new StreamWriter(fileStrmErr, Encoding.Default);

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
                }

            if (isRewardVal)
                {
                //maintainData.curRewardCond = rewardCond;
                //maintainData.fixReward(pBankCode, rewardCond);
                strMainTableCond = "m.contracttype != " + rewardCondVal;
                strSubTableCond = "d.trandescription != 'Calculated Points'";
                getReward(pBankCode);
                }
            if (isExcludedVal)
                {
                MainTableCond = " m.cardproduct not in " + strExcludeCond + "";//strWhereCond
                //supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.cardproduct not in " + strExcludeCond + ")";
                }
            FillStatementDataSet(pBankCode, "vip"); //DSstatement =  //10); //3
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
                            //if (HasSender == true)
                            //    SendEmailWithDifferentSender(emailStr.ToString(), "", emailTo, emailCc);
                            //else
                                SendEmail(emailStr.ToString(), "", emailTo, emailCc);
                        else
                            //if (HasSender == true)
                            //    SendEmailWithDifferentSender(emailStr.ToString(), "", emailTo, "");
                            //else
                                SendEmail(emailStr.ToString(), "", emailTo, "");

                    if (!string.IsNullOrEmpty(emailStr.ToString()) && string.IsNullOrEmpty(emailTo)) //emailStr != string.Empty && emailStr != null  emailTo != string.Empty && emailTo != null
                        {
                        streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|Without Email");
                        noOfWithoutEmails++;
                        emailTo = strEmailFrom; //"statement_Program@emp-group.com";
                        emailLabelTmp = emailLabel + " Acc:" + curAccountNumber + " Phone:" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!");
                        //if (HasSender == true)
                        //    SendEmailWithDifferentSender(emailStr.ToString(), "", emailTo, "");
                        //else
                            SendEmail(emailStr.ToString(), "", emailTo, "");
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
                    emailRows = DSemails.Tables["Emails"].Select("idclient = " + masterRow[mClientid].ToString());
                    for (int i = 0; i < emailRows.Length; i++)
                        {
                        emailTo = emailRows[i][1].ToString().Trim();
                        if (pBankCode == 21)
                            emailCc = emailRows[i][6].ToString().Trim();
                        else
                            emailCc="";
                        emailRow = emailRows[i];
                        }
                    curAccountNumber = masterRow[accountNoName].ToString();
                    curCardNumber = strPrimaryCardNo;
                    curClientID = masterRow[mClientid].ToString();

                    //emailTo = (DSemails.Tables["Emails"].Select("idclient = " + masterRow[mClientid].ToString));
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
                    stmNo = detailRow[dStatementno].ToString();
                    if ((detailRow[dPostingdate] == DBNull.Value) && (detailRow[dDocno] == DBNull.Value)) continue;// Exclude On-Hold Transactions 
                    curAccRows++;
                    //-if (CurPageRec4Dtl >= MaxDetailInPage)
                    //-{
                    //-  CurPageRec4Dtl = 0;
                    //-  printCardFooter();
                    //-  pageNo++;
                    //-  printHeader();
                    //-}
                    //						totNetUsage +=clsBasValid.validateNum(detailRow[dBilltranamount]);
                    totNetUsage = calculateCrDb(totNetUsage, clsBasValid.validateNum(detailRow[dBilltranamount]), clsBasValid.validateStr(detailRow[dBilltranamountsign]));
                    CurPageRec4Dtl = CurPageRec4Dtl + 1;
                    printDetail();
                    //if(!((detailRow[dPostingdate]== DBNull.Value) && (detailRow[dDocno]== DBNull.Value))) 

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
                } //end of Master foreach
            emailLabelTmp = emailLabel;

            if (!string.IsNullOrEmpty(emailStr.ToString()) && !string.IsNullOrEmpty(emailTo)) //emailStr != string.Empty && emailStr != null emailTo != string.Empty && emailTo != null
                if (!string.IsNullOrEmpty(emailStr.ToString()) && !string.IsNullOrEmpty(emailCc)) //emailStr != string.Empty && emailStr != null emailTo != string.Empty && emailTo != null
                    //if (HasSender == true)
                    //    SendEmailWithDifferentSender(emailStr.ToString(), "", emailTo, emailCc);
                    //else
                        SendEmail(emailStr.ToString(), "", emailTo, emailCc);
                else
                    //if (HasSender == true)
                    //    SendEmailWithDifferentSender(emailStr.ToString(), "", emailTo, "");
                    //else
                        SendEmail(emailStr.ToString(), "", emailTo, "");

            if (!string.IsNullOrEmpty(emailStr.ToString()) && string.IsNullOrEmpty(emailTo)) // emailStr != string.Empty && emailStr != null emailTo != string.Empty && emailTo != null
                {
                streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|Without Email");
                noOfWithoutEmails++;
                emailTo = strEmailFrom; //"statement_Program@emp-group.com";
                emailLabelTmp = emailLabel + " Acc:" + curAccountNumber + " Phone:" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!");
                //if (HasSender == true)
                //    SendEmailWithDifferentSender(emailStr.ToString(), "", emailTo, "");
                //else
                    SendEmail(emailStr.ToString(), "", emailTo, "");
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

            //-strmWriteErr.Flush();
            //-strmWriteErr.Close();
            //-fileStrmErr.Close();

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
        emailStr.Append(@"</head><body  background='cid:" + clsBasFile.getFileFromPath(strbackGround) + @"'><table border='0' width='100%' id='table1'>");

        emailStr.Append(@"<tr><td colSpan='2' height='82'>");
        emailStr.Append(@"<table id='table25' width='100%' border='0'>");
        emailStr.Append(@"<tr><td height='82' width='27%'><p align='" + logoAlignmentStr + @"'>");
        //emailStr.Append(@"<a href='http://" + strbankWebLink + @"'><img border='0' src='cid:logo.gif'></a></p></td>");
        emailStr.Append(@"<a href='http://" + strbankWebLink + @"'><img border='0' src='cid:" + clsBasFile.getFileFromPath(strBankLogo) + @"'></a></p></td>");
        emailStr.Append(@"<td height='82' width='3%' align='left'>&nbsp;</td><td height='82' align='center'>");
        emailStr.Append(@"<font size='6'>Statement</font></td><td height='82' width='1%'>&nbsp;</td>");
        emailStr.Append(@"<td height='82'><img border='0' src='cid:VisaLogo.gif' width='110' height='35' align='right' />");
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
        emailStr.Append(MakeHeaderStr(masterRow[accountLimit].ToString().Trim()));
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
            rewardRows = DSreward.Tables["Reward"].Select("CLIENTID = '" + clsBasValid.validateStr(masterRow[mClientid]).ToString().Trim() + "'");
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
        //int curPageLine =CurPageRec4Dtl;
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

    //private void SendEmailWithDifferentSender(string pBody, string pSubject, string pTo, string pCc = "")
    //    {
    //    //if (pTo.IndexOf('@') == -1 || pTo.IndexOf('.') == -1)
    //    if (valdEmail.isValideEmail(pTo) != "ValidEmail")//!basText.isValideEmail(pTo)
    //        {
    //        streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|"+clsCnfg.readSetting("strValidEmail")?? "Bad Email");
    //        noOfBadEmails++;
    //        emailTo = strEmailFrom; //"statement_Program@emp-group.com";
    //        emailLabelTmp = emailLabel + " Acc:" + curAccountNumber + " Phone:" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!");
    //        //return;
    //        }


    //    if (!string.IsNullOrEmpty(pCc) && valdEmail.isValideEmail(pCc) != "ValidEmail")//!basText.isValideEmail(pTo)
    //        {
    //        streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailCc + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|Bad Additional Email");
    //        noOfBadEmails++;
    //        emailTo = strEmailFrom; //"statement_Program@emp-group.com";
    //        emailLabelTmp = emailLabel + " Acc:" + curAccountNumber + " Phone:" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + " Bad additional Email";
    //        //return;
    //        }

    //    //if (pTo.ToUpper().IndexOf("YAHOO") > -1 || pTo.ToUpper().IndexOf("HOTMAIL") > -1)
    //    //  pBody.Replace("cid:","");

    //    try
    //        {
    //        //string pFrom, pSubject, pBody;
    //        //lblStatus.Text = string.Empty;
    //        //pFrom = "mmohammed@emp-group.com";//mmohammed@mscc.local"mmohammed@emp-group.com"mmohammed@emp-group.com
    //        //return;

    //        //streamEmails.WriteLine(curAccountNumber + "|" + curCardNumber + "|" + curClientID + "|" + emailTo);
    //        streamEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
    //        streamEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailCc + "|" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
    //        //return;
    //        //string[] strAray;
    //        //if (emailTo != "lucky.ighade@zenithbank.com")//ehimolen@yahoo.com
    //        //return;
    //        ArrayList pLstTo = new ArrayList(), pLstCc = new ArrayList(), pLstAttachedFile = new ArrayList();//, pLstCC = new ArrayList(), pLstBCC = new ArrayList();
    //        pSubject = emailLabelTmp; //emailLabel "BAI statement for 02/2008";
    //        pTo = emailTo;//"mali@emp-group.com"; 
    //        pCc = emailCc;

    //        if (pTo.EndsWith("."))//ehimolen@yahoo.com
    //            pTo = pTo + "com";

    //        if (pCc.EndsWith("."))//ehimolen@yahoo.com
    //            pCc = pCc + "com";
    //        //if (!pTo.ToUpper().EndsWith("DIAMONDBANK.COM"))//ehimolen@yahoo.com
    //        //  return;
    //        pLstTo.Add(pTo);//"mmohammed@emp-group.com" "mhrap@hotmail.com""mhrap@yahoo.com"    "jedkosi@gmail.com"   "nazab@emp-group.com"
    //        if (!(pCc == string.Empty))
    //            pLstCc.Add(pCc);
    //        //pLstBCC.Add("mmohammed@emp-group.com");
    //        //pLstBCC.Add("nazab@emp-group.com");
    //        //pLstBCC.Add("owafa@emp-group.com");
    //        //pLstTo.Add("mmohammed@emp-group.com");//"mmohammed@emp-group.com""nazab@emp-group.com"


    //        //pLstTo.Add("ashorungbe@yahoo.com");
    //        //pLstTo.Add("adeyinka.shorungbe@zenithbank.com");
    //        //pLstTo.Add("nazab@emp-group.com");
    //        //pLstTo.Add("hfawzy@emp-group.com");
    //        //pLstTo.Add("mmohammed@emp-group.com");//mmohammed@mscc.localmmohammed@emp-group.com nazab@emp-group.commmohammed@emp-group.com
    //        pLstAttachedFile.Add(strBankLogo);//strOutputFile@"D:\Web\Email\BAI\Logo.gif"
    //        pLstAttachedFile.Add(backGround);//@"D:\pC#\exe\FilesForPrograms\frmBackground.jpg"
    //        pLstAttachedFile.Add(strVisaLogo);//@"D:\pC#\exe\FilesForPrograms\frmBackground.jpg"
    //        sndMail = new clsEmail();
    //        sndMail.emailFromName = emailFromNameStr;
    //        //if (!sndMail.sendEmailHTML(strEmailFrom, pLstTo, pLstCC, pLstBCC, pLstAttachedFile, pSubject, pBody, clsCnfg.readSetting("SmtpServer")))
    //        strEmailFromTmp = strEmailFrom;
    //        //if (strEmailFrom == emailTo)
    //        //{
    //        //    strEmailFromTmp = "mabouleila@emp-group.com";
    //        //}
    //        //if (pTo.ToLower().IndexOf("bankphb.com") > 0)
    //        if (pTo.ToLower().IndexOf("keystonebankng.com") > 0)
    //            //strEmailFromTmp = "phbebankinghelpdesk@emp-group.com";
    //            strEmailFromTmp = "Ebankingcustomerservice@keystonebankng.com";
    //        else
    //            strEmailFromTmp = strEmailFrom;

    //        numOfTry = 0;
    //        //while (!sndMail.sendEmailHTML(strEmailFromTmp, pLstTo, pLstCC, pLstBCC, pLstAttachedFile, pSubject, pBody, clsCnfg.readSetting("SmtpServer")))
    //        noOfEmails++;

    //        if (numMail == 0)
    //            {
    //            //pLstBCC.Add("mmohammed@emp-group.com");
    //            pLstBCC.Add("statement@emp-group.com");
    //            }
    //        else
    //            {
    //            pLstBCC.Remove("statement@emp-group.com");
    //            }

    //        while (!sndMail.sendEmailAttachPicwithDiffernetSender(strEmailFromTmp, "statement@emp-group.com", pLstTo, pLstCc, pLstBCC, pLstAttachedFile, pSubject, pBody, clsCnfg.readSetting("SmtpServer")))
    //            {
    //            //MessageBox.Show("Failure to send Email.", "Send Email Error",MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
    //            System.Threading.Thread.Sleep(waitPeriodVal);//2000
    //            numOfTry++;
    //            if (numOfTry > 100)
    //                {
    //                streamEmails.Write("\t\t Error while Send Email");
    //                streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|Exceed number of trials");
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
    //        sndMail = null;
    //        emailStr = new StringBuilder(""); emailTo = string.Empty; curAccountNumber = string.Empty; curCardNumber = string.Empty; curClientID = string.Empty;
    //        wait2NextEmail(prvEmail, pTo, waitPeriodVal);//     System.Threading.Thread.Sleep(waitPeriodVal);//400
    //        prvEmail = pTo;
    //        }
    //    catch  //(NotSupportedException ex) (Exception ex)  //
    //        {
    //        //clsBasErrors.catchError(ex);
    //        streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|"+clsCnfg.readSetting("strValidEmail")?? "Bad Email");
    //        noOfBadEmails++;
    //        }
    //    finally
    //        {
    //        }

    //    }

    private void SendEmail(string pBody, string pSubject, string pTo, string pCc = "")
    {
        try
        {
        //if (pTo.IndexOf('@') == -1 || pTo.IndexOf('.') == -1)
        if (valdEmail.isValideEmail(pTo) != "ValidEmail")//!basText.isValideEmail(pTo)
            {
                streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + pTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|"+clsCnfg.readSetting("strValidEmail")?? "Bad Email " + pTo);
            noOfBadEmails++;
            pTo = strEmailFrom; //"statement_Program@emp-group.com";
            emailLabelTmp = emailLabel + " Acc:" + curAccountNumber + " Phone:" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!");
                //if(curAccountNo == "5450087544")
                //{
                //    string ss = pTo;
                //}
            //return;
            }


        if (!string.IsNullOrEmpty(pCc) && valdEmail.isValideEmail(pCc) != "ValidEmail")//!basText.isValideEmail(pTo)
            {
                streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + pCc + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|Bad Additional Email");
                noOfBadEmails++;
                pCc = string.Empty;
                //emailTo = strEmailFrom; //"statement_Program@emp-group.com";
                //emailLabelTmp = emailLabel + " Acc:" + curAccountNumber + " Phone:" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + " Bad additional Email";
                //return;
            }

            //if (pTo.ToUpper().IndexOf("YAHOO") > -1 || pTo.ToUpper().IndexOf("HOTMAIL") > -1)
            //  pBody.Replace("cid:","");

            //try
            //    {
            //string pFrom, pSubject, pBody;
            //lblStatus.Text = string.Empty;
            //pFrom = "mmohammed@emp-group.com";//mmohammed@mscc.local"mmohammed@emp-group.com"mmohammed@emp-group.com
            //return;

            //streamEmails.WriteLine(curAccountNumber + "|" + curCardNumber + "|" + curClientID + "|" + emailTo);
            //streamEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
            //streamEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailCc + "|" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
            //return;
            //string[] strAray;
            //if (emailTo != "lucky.ighade@zenithbank.com")//ehimolen@yahoo.com
            //return;
            ArrayList pLstTo = new ArrayList(), pLstCc = new ArrayList(), pLstAttachedFile = new ArrayList();//, pLstCC = new ArrayList(), pLstBCC = new ArrayList();
            pSubject = emailLabelTmp; //emailLabel "BAI statement for 02/2008";
            //pTo = emailTo;//"mali@emp-group.com"; 
            //pCc = emailCc;

            if (pTo.EndsWith("."))//ehimolen@yahoo.com
                pTo = pTo + "com";

            if (pCc.EndsWith("."))//ehimolen@yahoo.com
                pCc = pCc + "com";
            //if (!pTo.ToUpper().EndsWith("DIAMONDBANK.COM"))//ehimolen@yahoo.com
            //  return;
            pLstTo.Add(pTo);//"mmohammed@emp-group.com" "mhrap@hotmail.com""mhrap@yahoo.com"    "jedkosi@gmail.com"   "nazab@emp-group.com"
            if (!(pCc == string.Empty))
                pLstCc.Add(pCc);
            //pLstBCC.Add("mmohammed@emp-group.com");
            //pLstBCC.Add("nazab@emp-group.com");
            //pLstBCC.Add("owafa@emp-group.com");
            //pLstTo.Add("mmohammed@emp-group.com");//"mmohammed@emp-group.com""nazab@emp-group.com"


            //pLstTo.Add("ashorungbe@yahoo.com");
            //pLstTo.Add("adeyinka.shorungbe@zenithbank.com");
            //pLstTo.Add("nazab@emp-group.com");
            //pLstTo.Add("hfawzy@emp-group.com");
            //pLstTo.Add("mmohammed@emp-group.com");//mmohammed@mscc.localmmohammed@emp-group.com nazab@emp-group.commmohammed@emp-group.com
            pLstAttachedFile.Add(strBankLogo);//strOutputFile@"D:\Web\Email\BAI\Logo.gif"
            pLstAttachedFile.Add(backGround);//@"D:\pC#\exe\FilesForPrograms\frmBackground.jpg"
            pLstAttachedFile.Add(strVisaLogo);//@"D:\pC#\exe\FilesForPrograms\frmBackground.jpg"
            sndMail = new clsEmail();
            sndMail.emailFromName = emailFromNameStr;
            //if (!sndMail.sendEmailHTML(strEmailFrom, pLstTo, pLstCC, pLstBCC, pLstAttachedFile, pSubject, pBody, clsCnfg.readSetting("SmtpServer")))
            strEmailFromTmp = strEmailFrom;
            //if (strEmailFrom == emailTo)
            //{
            //    strEmailFromTmp = "mabouleila@emp-group.com";
            //}
            //if (pTo.ToLower().IndexOf("bankphb.com") > 0)
            if (pTo.ToLower().IndexOf("keystonebankng.com") > 0)
                //strEmailFromTmp = "phbebankinghelpdesk@emp-group.com";
                strEmailFromTmp = "Ebankingcustomerservice@keystonebankng.com";
            else
                strEmailFromTmp = strEmailFrom;

            numOfTry = 0;
            //while (!sndMail.sendEmailHTML(strEmailFromTmp, pLstTo, pLstCC, pLstBCC, pLstAttachedFile, pSubject, pBody, clsCnfg.readSetting("SmtpServer")))
            noOfEmails++;

            if (numMail == 0)
                {
                //pLstBCC.Add("mmohammed@emp-group.com");
                pLstBCC.Add("statement@emp-group.com");
                }
            else
                {
                pLstBCC.Remove("statement@emp-group.com");
                }

            if (strEmailFromTmp.ToUpper().EndsWith("ZENITHBANK.COM"))
            {
                if (pLstCc.Contains("statement@emp-group.com") == false)
                {
                    pLstCc.Add("cardservices@zenithbank.com");
                    if (!pSubject.Contains("Acc:"))
                        pSubject += " Acc:" + curAccountNumber;

                    //pLstCc.Add("mamr@network.com.eg");
                }
            }

            while (!sndMail.sendEmailAttachPic(strEmailFromTmp, pLstTo, pLstCc, pLstBCC, pLstAttachedFile, pSubject, pBody, clsCnfg.readSetting("SmtpServer")))
                {
                //MessageBox.Show("Failure to send Email.", "Send Email Error",MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                System.Threading.Thread.Sleep(waitPeriodVal);//2000
                numOfTry++;
                if (numOfTry > 100)
                    {
                    streamEmails.Write("\t\t Error while Send Email");
                    streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|Exceed number of trials");
                    noOfBadEmails++; noOfEmails--;
                    break;
                    }
                }
            streamEmails.WriteLine(curAccountNumber + "|" + curClientID + "|To " + pTo + "; Cc " + pCc + "|" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));

            numMail++;
            if (numMail % 400 == 0)
                {
                System.Threading.Thread.Sleep(waitPeriodVal);//2000
                GC.Collect();
                GC.WaitForPendingFinalizers();
                }
            //sndMail = null;
            //emailStr = new StringBuilder(""); emailTo = string.Empty; emailCc = string.Empty; curAccountNumber = string.Empty; curCardNumber = string.Empty; curClientID = string.Empty;
            //wait2NextEmail(prvEmail, pTo, waitPeriodVal);//     System.Threading.Thread.Sleep(waitPeriodVal);//400
            //prvEmail = pTo;
            }
        catch (Exception ex) //(NotSupportedException ex) (Exception ex)  //
            {
            //clsBasErrors.catchError(ex);
                streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + pTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "| Email " + pTo + " , Err Message >> " + ex.Message + ", Err Desc >> ");
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

    public string ExcludeCond
        {
        get { return strExcludeCond; }
        set { strExcludeCond = value; }
        }  // ExcludeCond

    public bool isExcluded
        {
        get { return isExcludedVal; }
        set { isExcludedVal = value; }
        }// isExcluded

    public bool HasSender
        {
        get { return HasSenderVal; }
        set { HasSenderVal = value; }
        }  // HasSender

    ~clsStatHtmlGnrl00()
        {
        DSstatement.Dispose();
        }
    }
