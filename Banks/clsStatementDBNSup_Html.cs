using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Xml;
using System.Collections;


public class clsStatementDBNSup_Html : clsBasStatement
{
    private string strBankName;
    private FileStream fileStrm, fileSummary, fileEmails, fileNoEmails;
    private StreamWriter streamWrit, streamSummary, streamEmails, streamNoEmails;
    //private DataSet DSstatement;
    //private OracleDataReader drPrimaryCards, drMaster,drDetail;
    private DataRow masterRow;
    private DataRow detailRow;
    private DataRow[] cardRow;
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
    private string curAccountNo, prevcardNo = String.Empty;//,PrevCardNo
    private decimal totNetUsage = 0;
    private DataRow[] cardsRows, accountRows, rewardRows, tmpDtlRows;
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
    //private clsOMR omr = new clsOMR();
    private int totPages;
    string endOfCustomer = string.Empty;
    string cProduct = string.Empty, curFileName = string.Empty;
    private ArrayList aryLstFiles;

    //
    //private string emailStr;
    private StringBuilder emailStr = new StringBuilder("");
    private string emailLabel, strBankLogo, strBankMidBanner, strBankBottomBanner;
    private string strEmailFrom = "mabouleila@emp-group.com";  //"cardservices@zenithbank.com"
    private string strbankWebLink = "www.emp-group.com";  //www.zenithbank.com
    private string emailTo, curAccountNumber, curCardNumber, curClientID;//, curEmail
    private string strEmailFileName, strNoEmailFileName;
    private frmStatementFile frmMain;
    private int totRec = 1, numMail, numOfTry = 0;
    private ArrayList pLstCC = new ArrayList(), pLstBCC = new ArrayList();
    //private string logoAlignmentStr = "center";
    private string logoAlignmentStr = "left";
    private string strbackGround = @"D:\pC#\ProjData\Statement\_Background\Background05.jpg";
    private bool isRewardVal = false;
    private bool createCorporateVal = false;
    private string accountNoName = mAccountno;
    private string accountLimit = mAccountlim;
    private string cardLimit = mCardlimit;
    private string accountAvailableLimit = mAccountavailablelim;
    private string cardAvailableLimit = mCardavailablelimit;
    private string rewardCondVal = "'New Reward Contract'";//'Reward Contract'
    private string vipCondVal = string.Empty;
    private clsEmail sndMail;
    private string emailFromNameStr;
    private int waitPeriodVal = 7000;
    clsAes cryptAes = new clsAes();
    private clsValidateEmail valdEmail = new clsValidateEmail();
    private int noOfEmails, noOfBadEmails, noOfWithoutEmails;
    private string strEmailFromTmp, emailLabelTmp;  //"cardservices@zenithbank.com"
    private DataRow emailRow;
    private string strFileNam, stmntType;
    private string curSuppCard = string.Empty;
    //private string bottomBanner = @"D:\pC#\ProjData\Statement\DBN Diamond\Banners_Bottom.jpg";
    private string prvEmail = string.Empty;
    private decimal totTrans = 0;

    public clsStatementDBNSup_Html()
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
            //maintainData.notRward = false;
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
            //emailLabel = pBankName + " statement for " + vCurDate.ToString("MM/yyyy"); //"BAI statement for 02/2008"
            emailLabel = pBankName + " Transactions Details for " + vCurDate.ToString("MM/yyyy"); //"BAI statement for 02/2008"
            strOutputFile = pStrFileName;

            // open emails file
            fileEmails = new FileStream(strEmailFileName + ".txt", FileMode.Create); //Create
            streamEmails = new StreamWriter(fileEmails, Encoding.Default);
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
                curRewardCond = rewardCond;
                getReward(pBankCode);
            }
            if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("DBN_Credit_ClientsEmails_sup"))
            {
                strMainTableCond += " and m.contracttype not in " + VIPCondition + " and m.cardprimary = 'N'";
            }
            else if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("DBN_VIP1-2-5_ClientsEmails_sup"))
            {
                strMainTableCond += " and m.contracttype in " + VIPCondition + " and m.cardprimary = 'N'";
            }
            else if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("DBN_VIP_ClientsEmails_sup"))
            {
                strMainTableCond += " and m.contracttype in " + VIPCondition + " and m.cardprimary = 'N'";
            }
            // data retrieve
            //FillStatementDataSet(pBankCode, true); //DSstatement =  //10); //3
            FillStatementDataSet(pBankCode, "");
            getClientEmail(pBankCode);
            pageNo = 1; totalCardPages = 0;
            curCardNo = String.Empty;
            curAccountNo = String.Empty;
            frmMain.BeginInvoke(frmMain.setMinMaxProgressDelegate, new object[] { DSstatement.Tables["tStatementMasterTable"].Rows.Count });
            foreach (DataRow mRow in DSstatement.Tables["tStatementMasterTable"].Rows)
            {
                frmMain.BeginInvoke(frmMain.setProgressDelegate, new object[] { totRec++ });
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

                if (masterRow[mCardprimary].ToString() == "Y")
                {
                    continue;

                    //calcCardlRows();
                }

                //start new account
                if (prevcardNo != masterRow[mCardno].ToString())
                {
                    emailLabelTmp = emailLabel;
                    if (!string.IsNullOrEmpty(emailStr.ToString()) && !string.IsNullOrEmpty(emailTo)) //emailStr != string.Empty && emailStr != null  emailTo != string.Empty && emailTo != null
                        SendEmail(emailStr.ToString(), "", emailTo);
                    if (!string.IsNullOrEmpty(emailStr.ToString()) && string.IsNullOrEmpty(emailTo)) //emailStr != string.Empty && emailStr != null  emailTo != string.Empty && emailTo != null
                    {
                        streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|Without Email");
                        noOfWithoutEmails++;
                        emailTo = strEmailFrom; //"statement_Program@emp-group.com";
                        emailLabelTmp = emailLabel + " Acc:" + curAccountNumber + " Phone:" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!");
                        SendEmail(emailStr.ToString(), "", emailTo);
                    }

                    emailStr = new StringBuilder("");
                    //cardsRows = DSstatement.Tables["tStatementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[accountNoName]).ToString().Trim() + "'");
                    cardsRows = DSstatement.Tables["tStatementDetailTable"].Select("cardno = '" + clsBasValid.validateStr(masterRow[mCardno]).ToString().Trim() + "' and mbr = " + masterRow[mMBR].ToString().Trim());
                    if (cardsRows.Length == 0)
                    {
                        continue;
                    }

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
                    CurPageRec4Dtl = 0; totNetUsage = 0; CurrentPageFlag = "F 1";totTrans = 0; //if page is based on account no
                    calcAccountRows();




                    if (totAccRows < 1
                      && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0) //             || (masterRow[mCardno].ToString() == curMainCard   // Convert.ToDecimal(
                    {
                        isHaveF3 = true;
                        //pageNo=1; totalAccPages =1;
                        continue;
                    }
                    emailStr = new StringBuilder("");

                    prevcardNo = masterRow[mCardno].ToString();
                    //pageNo=1; //if page is based on account no
                    DataRow[] emailRows;

                    //cardRow = DSstatement.Tables["zm_cards"].Select("pan = " + masterRow[mCardno].ToString() + " and mbr = " + masterRow[mMBR].ToString());
                    try
                    {
                        //emailRows = DSemails.Tables["Emails"].Select("idclient = " + cardRow[0].ItemArray[3].ToString());
                        emailRows = DSemails.Tables["Emails"].Select("idclient = " + masterRow[mCardClientId].ToString());

                        for (int i = 0; i < emailRows.Length; i++)
                        {
                            emailTo = emailRows[i][1].ToString().Trim();
                            emailRow = emailRows[i];
                        }
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.Message);
                    }
                    curAccountNumber = masterRow[accountNoName].ToString();
                    curCardNumber = strCardNo;
                    curClientID = masterRow[mClientid].ToString();

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
                    CurPageRec4Dtl++;

                    //if (curSuppCard != detailRow[dCardno].ToString())
                    //{
                    //    tmpDtlRows = DSstatement.Tables["tStatementMasterTable"].Select(mCardno + " = '" + detailRow[dCardno] + "'");
                    //    if (tmpDtlRows.Length > 0)
                    //    {
                    //        if (tmpDtlRows[0][mCardprimary].ToString() == "N" && isValidateCard(tmpDtlRows[0][mCardstate].ToString()))
                    //        {
                    //            emailStr.Append(@"<tr><td align=""middle"" width=""6%"" height=""20"">");
                    //            emailStr.Append(MakeHeaderStr("     ", false, false));
                    //            emailStr.Append(@"</td><td align=""middle"" width=""5%"" height=""20"">");
                    //            emailStr.Append(MakeHeaderStr("     ", false, false));
                    //            emailStr.Append(@"</td><td align=""middle"" width=""19%"" height=""20"">");
                    //            emailStr.Append(MakeHeaderStr(basText.trimStr(" ", 24), false, false));
                    //            emailStr.Append(@"</td><td width=""42%"" height=""20""><p align=""left"">");
                    //            emailStr.Append(MakeHeaderStr(basText.trimStr(">> Supplementary Card : " + basText.formatCardNumber(detailRow[dCardno].ToString()), 40), false, false));
                    //            emailStr.Append(@"</td><td align=""right"" width=""10%"" height=""20"">");
                    //            emailStr.Append(" ");
                    //            emailStr.Append(@"</td><td align=""right"" width=""11%"" height=""20"">");
                    //            emailStr.Append(" ");
                    //            emailStr.Append(@"</td><td align=""right"" width=""1%"" height=""20"">");
                    //            emailStr.Append(" ");
                    //            emailStr.Append(@"</td></tr>");

                    //            CurPageRec4Dtl++;
                    //        }
                    //    }
                    //}

                    printDetail();
                    curSuppCard = detailRow[dCardno].ToString();
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
            emailTo = emailTo.Trim();
            emailLabelTmp = emailLabel;
            if (!string.IsNullOrEmpty(emailStr.ToString()) && !string.IsNullOrEmpty(emailTo)) // emailStr != string.Empty && emailStr != null emailTo != string.Empty && emailTo != null
                SendEmail(emailStr.ToString(), "", emailTo);
            if (!string.IsNullOrEmpty(emailStr.ToString()) && string.IsNullOrEmpty(emailTo)) // emailStr != string.Empty && emailStr != null emailTo != string.Empty && emailTo != null
            {
                streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|Without Email");
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
        totPages++; endOfCustomer = "";
        extAccNum = clsBasValid.validateStr(masterRow[mExternalno]);
        //DBN-6861 
        if (extAccNum.Trim() != "")
            extAccNum = "******" + extAccNum.Substring(6);
        if (extAccNum.Trim() == "")
            extAccNum = clsBasValid.validateStr(masterRow[accountNoName]);



        //\MakeEmailStr   MakeHeaderStr   emailStr
        emailStr = new StringBuilder("");
        emailStr.Append(@"<html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1252"">");
        emailStr.Append(@"<title>");
        emailStr.Append(masterRow[mCardclientname].ToString() + " - Transaction Details");
        emailStr.Append(@"</title>");
        emailStr.Append(@"</head><body background=""cid:" + clsBasFile.getFileFromPath(strbackGround) + @""">");
        emailStr.Append(@"<table id=""table15"" width=""79%"" border=""0""><tr><td height=""82"">");
        emailStr.Append(@"<p align=""" + logoAlignmentStr + @""">");
        emailStr.Append(@"<a href=""http://" + strbankWebLink + @""">");
        emailStr.Append(@"<img border=""0"" src=""cid:" + clsBasFile.getFileFromPath(strBankLogo) + @"""></a></td></tr>");// width=""230"" height=""85""
        emailStr.Append(@"<br><center><table><tr><td><b>Supplementary Card Transactions Details</b></td></tr></table></center>");// width=""230"" height=""85""
        emailStr.Append(@"<center>(This is not a demand for payment but simply a log of transactions made on the supplementary card for your records)</center><br>");
        emailStr.Append(@"<tr><td width=""99%"">");
        emailStr.Append(@"<table id=""table18"" height=""93"" width=""99%"" border=""0"" style=""border-collapse: collapse""><tr><td>");
        emailStr.Append(@"<table id=""table19"" height=""82"" width=""100%"" border=""0""><tr><td width=""103"" style=""border-style: outset; border-width: 1px"" height=""20"">");
        //emailStr.Append(MakeHeaderStr("Name :", true, true));
        emailStr.Append(MakeHeaderStr2("Name :", true, true));
        emailStr.Append(@"</td><td height=""20"">");
        //emailStr.Append(MakeHeaderStr(masterRow[mCustomername].ToString(), false, false));
        emailStr.Append(MakeHeaderStr2(masterRow[mCustomername].ToString(), true, false));
        emailStr.Append(@"</td></tr>");
        emailStr.Append(@"<tr><td width=""103"" style=""border-style: outset; border-width: 1px"" height=""21"">");
        //emailStr.Append(MakeHeaderStr("Card Product :", true, true));
        emailStr.Append(MakeHeaderStr2("Card Product :", true, true));
        emailStr.Append(@"</td><td height=""21"">");
        //emailStr.Append(MakeHeaderStr(masterRow[mCardproduct].ToString(), false, false));
        emailStr.Append(MakeHeaderStr2(masterRow[mCardproduct].ToString(), true, false));
        emailStr.Append(@"</td></tr><tr><td width=""103"" style=""border-style: outset; border-width: 1px"" height=""19"">");
        //emailStr.Append(MakeHeaderStr("Acc. Number :", true, true));
        emailStr.Append(MakeHeaderStr2("Acc. Number :", true, true));
        emailStr.Append(@"</td><td height=""19"">");
        //emailStr.Append(MakeHeaderStr(extAccNum, false, false));
        emailStr.Append(MakeHeaderStr2(extAccNum, true, false));
        emailStr.Append(@"</td></tr><tr><td width=""103"" style=""border-style: outset; border-width: 1px"">");
        //emailStr.Append(MakeHeaderStr("Statement Date:", true, true));
        emailStr.Append(MakeHeaderStr2("Statement Date:", true, true));
        emailStr.Append(@"</td><td>");
        //emailStr.Append(MakeHeaderStr(Convert.ToDateTime(masterRow[mStatementdateto]).ToString("dd/MM/yyyy"), false, false));
        emailStr.Append(MakeHeaderStr2(Convert.ToDateTime(masterRow[mStatementdateto]).ToString("dd/MM/yyyy"), true, false));
        emailStr.Append(@"</td></tr></table></td></tr></table></td></tr>");

        //emailStr.Append(@"<tr><td width=""98%"" height=""25"">&nbsp;</td></tr><tr><td width=""98%"" height=""48""><table id=""table20"" style=""height:46"" width=""100%"" border=""0"">");
        emailStr.Append(@"<tr><td width=""98%"" height=""25"">&nbsp;</td></tr><tr><td width=""98%"" height=""48""><table id=""table20"" style=""height:46"" width=""79%"" border=""0"">");
        emailStr.Append(@"<tr><td align=""middle"" width=""124"" style=""border-style: outset; border-width: 1px"">");
        emailStr.Append(MakeHeaderStr("Card Credit Limit", false, true));
        emailStr.Append(@"</td><td align=""middle"" width=""122"" style=""border-style: outset; border-width: 1px"">");
        emailStr.Append(MakeHeaderStr("Card Available Limit", false, true));
        emailStr.Append(@"</td><td align=""middle"" width=""200"" style=""border-style: outset; border-width: 1px"">");
        //emailStr.Append(MakeHeaderStr("Min. Due Amount", true, true);
        //emailStr.Append(MakeHeaderStr("Minimum Payment Now Due", true, true));
        //emailStr.Append(@"</td><td align=""middle"" width=""220"" style=""border-style: outset; border-width: 1px"">");
        //emailStr.Append(MakeHeaderStr("Min. Due Date", true, true);
        //emailStr.Append(MakeHeaderStr("Please Pay this Amount Before", true, true));
        //emailStr.Append(@"</td><td align=""middle"" width=""145"" style=""border-style: outset; border-width: 1px"">");
        //emailStr.Append(MakeHeaderStr("Over Due Amount", false, true));
        emailStr.Append(@"</td></tr>");
        emailStr.Append(@"<tr><td align=""middle"" width=""124"" height=""19"">");
        emailStr.Append(MakeHeaderStr(masterRow[cardLimit].ToString().Trim(), false, false));
        emailStr.Append(@"</td><td align=""middle"" width=""122"" height=""19"">");
        emailStr.Append(MakeHeaderStr(basText.formatNum(masterRow[cardAvailableLimit], "##0", 20).Trim(), false, false));
        //emailStr.Append(@"</td><td align=""middle"" width=""173"" height=""19"">");
        //emailStr.Append(MakeHeaderStr(masterRow[mMindueamount].ToString().Trim(), true, false));
        //emailStr.Append(@"</td><td align=""middle"" width=""161"" height=""19"">");
        //emailStr.Append(MakeHeaderStr(basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M"), true, false));
        //emailStr.Append(@"</td><td align=""middle"" width=""140"" height=""19"">");
        //emailStr.Append(MakeHeaderStr(basText.formatNum(masterRow[mTotaloverdueamount], "##0.00", 13).Trim(), false, false));
        emailStr.Append(@"</td></tr></table></td></tr>");

        emailStr.Append(@"<tr><td width=""98%"" height=""46%"">&nbsp;</td></tr>");
        //emailStr.Append(@"<tr><td width=""98%"" height=""88""><table id=""table21"" style=""height:50"" width=""100%"" border=""0"">");// style=""border-collapse: collapse"">
        emailStr.Append(@"<tr><td width=""98%"" height=""88""><table id=""table21"" style=""height:50"" width=""79%"" border=""0"">");// style=""border-collapse: collapse"">
        emailStr.Append(@"<tr><td align=""middle"" colSpan=""2"" height=""16"" style=""border-style: outset; border-width: 1px"">");
        emailStr.Append(MakeHeaderStr("Dates of", false, true));
        emailStr.Append(@"</td><td align=""middle"" width=""19%"" height=""39"" rowSpan=""2"" style=""border-style: outset; border-width: 1px"">");
        emailStr.Append(MakeHeaderStr("Reference No", true, true));
        emailStr.Append(@"</td><td align=""middle"" width=""52%"" colSpan=""2"" height=""39"" rowSpan=""2"" style=""border-style: outset; border-width: 1px"">");
        emailStr.Append(MakeHeaderStr("Description", true, true));
        emailStr.Append(@"</td><td align=""middle"" width=""13%"" colSpan=""2"" height=""39"" rowSpan=""2"" style=""border-style: outset; border-width: 1px"">");
        emailStr.Append(MakeHeaderStr("Amount (" + masterRow[mAccountcurrency].ToString().Trim() + ")", true, true));
        emailStr.Append(@"</td><td align=""middle"" width=""13%"" colSpan=""2"" height=""39"" rowSpan=""2"" style=""border-style: outset; border-width: 1px"">");
        emailStr.Append(MakeHeaderStr("Current Balance", true, true));
        emailStr.Append(@"</td></tr>");

        emailStr.Append(@"<tr><td align=""middle"" width=""6%"" height=""19"" style=""border-style: outset; border-width: 1px"">");
        emailStr.Append(MakeHeaderStr("Trans.", false, true));
        emailStr.Append(@"</td><td align=""middle"" width=""5%"" height=""19"" style=""border-style: outset; border-width: 1px"">");
        emailStr.Append(MakeHeaderStr("Posting", false, true));
        emailStr.Append(@"</td></tr>");

        totalPages++;
        totNoOfPageStat++;
        totTrans = Convert.ToDecimal(masterRow[mOpeningbalance]);
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
        totTrans += valDbCr(Convert.ToDecimal(detailRow[dBilltranamount]), clsBasValid.validateStr(detailRow[dBilltranamountsign]));
        emailStr.Append(@"<tr><td align=""middle"" width=""6%"" height=""20"">");
        emailStr.Append(MakeHeaderStr(trnsDate.ToString("dd/MM"), false, false));
        emailStr.Append(@"</td><td align=""middle"" width=""5%"" height=""20"">");
        emailStr.Append(MakeHeaderStr(postingDate.ToString("dd/MM"), false, false));
        emailStr.Append(@"</td><td align=""middle"" width=""19%"" height=""20"">");
        emailStr.Append(MakeHeaderStr(basText.trimStr(detailRow[dRefereneno].ToString().Trim(), 24), false, false));
        emailStr.Append(@"</td><td width=""42%"" height=""20""><p align=""left"">");
        emailStr.Append(MakeHeaderStr(basText.trimStr(trnsDesc.Trim(), 24), false, false));
        emailStr.Append(@"</td><td align=""right"" width=""10%"" height=""20"">");
        emailStr.Append(MakeHeaderStr(ValidateStr(strForeignCurr.Trim()), false, false));
        emailStr.Append(@"</td><td align=""right"" width=""11%"" height=""20"">");
        emailStr.Append(MakeHeaderStr((basText.formatNumUnSign(detailRow[dBilltranamount], "#,###,##0.00", 14) + " " + CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])).ToString()), false, false));
        emailStr.Append(@"</td><td align=""right"" width=""1%"" height=""20"">");
        emailStr.Append(ValidateStr(isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))));
        emailStr.Append(@"</td><td align=""right"" width=""1%"" height=""20"">");
        emailStr.Append(MakeHeaderStr((basText.formatNumUnSign(totTrans, "#,###,##0.00", 14) + " " + CrDb(clsBasValid.validateStr(totTrans)).ToString()), false, false));
        emailStr.Append(@"</td></tr>");

        //-streamWrit.WriteLine("  {0:dd/MM}  {1:dd/MM}  {2,-24}  {3,-40} {4,16} {5,16} {6,2}", trnsDate, postingDate, basText.trimStr(detailRow[dRefereneno].ToString().Trim(), 24), basText.trimStr(trnsDesc, 40), strForeignCurr, basText.formatNumUnSign(detailRow[dBilltranamount], "#,###,##0.00", 14), CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + " " + isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)
        totNoOfTransactions++;
    }

    protected void printCardFooter()
    {
        emailStr.Append(@"</table></td></tr>");
        emailStr.Append(@"<tr><td width=""99%"" height=""20"">&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td width=""100%""><table id=""table22"" width=""100%"" border=""0"" height=""64""><tr><td><table id=""table25"" width=""100%"" border=""0"">");
        emailStr.Append(@"<tr><td align=""center"">");
        emailStr.Append(ValidateStr(masterRow[mStatementmessageline1].ToString()));
        emailStr.Append(@"</td></tr><tr><td align=""center"">");
        emailStr.Append(ValidateStr(masterRow[mStatementmessageline2].ToString()));
        emailStr.Append(@"</td></tr><tr><td align=""center"">");
        emailStr.Append(ValidateStr(masterRow[mStatementmessageline3].ToString()));
        emailStr.Append(@"</td></tr></table></td></tr></table></td></tr>");
        emailStr.Append(@"<tr><td width=""99%"" height=""72"">&nbsp;</td></tr>");
        //emailStr.Append(@"<tr><td width=""98%""><table id=""table24"" width=""100%"" border=""0"">");
        //emailStr.Append(@"<a href=""http://" + strbankWebLink + @""">");
        //emailStr.Append(@"<img border=""0"" src=""cid:" + clsBasFile.getFileFromPath(strBankMidBanner) + @"""></a></td></tr></table><table id=""table26"" width=""100%"" border=""0"">");
        //emailStr.Append(@"<tr><td align=""middle"" width=""106"" style=""border-style: outset; border-width: 1px"">");
        //emailStr.Append(MakeHeaderStr("Opening Balance", false, true));
        //emailStr.Append(@"</td><td align=""middle"" style=""border-style: outset; border-width: 1px"">");
        //emailStr.Append(MakeHeaderStr("Payments", false, true));
        //emailStr.Append(@"</td><td align=""middle"" width=""125"" style=""border-style: outset; border-width: 1px"">");
        //emailStr.Append(MakeHeaderStr("Cash & Purchases", false, true));
        //emailStr.Append(@"</td><td align=""middle"" width=""55"" style=""border-style: outset; border-width: 1px"">");
        //emailStr.Append(MakeHeaderStr("Interest", false, true));
        //emailStr.Append(@"</td><td align=""middle"" width=""51"" style=""border-style: outset; border-width: 1px"">");
        //emailStr.Append(MakeHeaderStr("Charges", false, true));
        //emailStr.Append(@"</td><td align=""middle"" width=""118"" style=""border-style: outset; border-width: 1px"">");
        //emailStr.Append(MakeHeaderStr("Closing Balance", true, true));
        //emailStr.Append(@"</td><td align=""middle"" width=""89"" style=""border-style: outset; border-width: 1px"">");
        //emailStr.Append(MakeHeaderStr("Total Debit(s)", false, true));
        //emailStr.Append(@"</td><td align=""middle"" width=""97"" style=""border-style: outset; border-width: 1px"">");
        //emailStr.Append(MakeHeaderStr("Total Credit(s)", false, true));
        //emailStr.Append(@"</td></tr>");

        //emailStr.Append(@"<tr><td align=""middle"" width=""106"">");
        //emailStr.Append(MakeHeaderStr(basText.alignmentRight(basText.formatNumUnSign(masterRow[mOpeningbalance], "#,###,##0.00", 12) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance].ToString())), 15), false, false));
        //emailStr.Append(@"</td><td align=""middle"">");
        //emailStr.Append(MakeHeaderStr(basText.alignmentRight(basText.formatNumUnSign(masterRow[mTotalpayments], "#,###,##0.00", 12) + " " + DbCr(Convert.ToDecimal(masterRow[mTotalpayments])), 15), false, false));
        //emailStr.Append(@"</td><td align=""middle"" width=""127"">");
        //emailStr.Append(MakeHeaderStr(basText.alignmentRight(basText.formatNum(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal]), "#,###,##0.00", 12) + CrDbMinus(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal])), 15), false, false));
        //emailStr.Append(@"</td><td align=""middle"" width=""57"">");
        //emailStr.Append(MakeHeaderStr(basText.alignmentRight(basText.formatNumUnSign(masterRow[mTotalinterest].ToString(), "#,###,##0.00", 12) + " " + CrDbMinus(Convert.ToDecimal(masterRow[mTotalinterest])), 15), false, false));
        //emailStr.Append(@"</td><td align=""middle"" width=""53"">");
        //emailStr.Append(MakeHeaderStr(basText.alignmentRight(basText.formatNumUnSign(masterRow[mTotalcharges].ToString(), "#,###,##0.00", 12) + " " + CrDbMinus(Convert.ToDecimal(masterRow[mTotalcharges])), 15), false, false));
        //emailStr.Append(@"</td><td align=""middle"" width=""120"">");
        //emailStr.Append(MakeHeaderStr(basText.alignmentRight(basText.formatNumUnSign(masterRow[mClosingbalance], "#,###,##0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])), 15), true, false));
        //emailStr.Append(@"</td><td align=""middle"" width=""91"">");
        //emailStr.Append(MakeHeaderStr(basText.formatNum(masterRow[mTotaldebits], "#0.00;(#0.00)"), false, false));
        //emailStr.Append(@"</td><td align=""middle"" width=""99"">");
        //emailStr.Append(MakeHeaderStr(basText.formatNum(masterRow[mTotalcredits], "#0.00;(#0.00)"), false, false));
        //emailStr.Append(@"</td></tr></table></td></tr>");
        //emailStr.Append(@"<tr><td width=""98%"">&nbsp;</td></tr>");
        //if (isRewardVal)
        //{
        //    rewardRows = DSreward.Tables["Reward"].Select("CLIENTID = '" + clsBasValid.validateStr(masterRow[mClientid]).ToString().Trim() + "'");
        //    if (rewardRows.Length > 0)
        //    {
        //        rewardRow = rewardRows[0];
        //        emailStr.Append(@"<tr><td width=""98%"" colSpan=""2""><table id=""table25"" width=""100%"" border=""0"">");
        //        emailStr.Append(@"<tr><td align=""middle"" width=""220"" style=""border-style: outset; border-width: 1px"">");
        //        emailStr.Append(MakeHeaderStr("Gem Rewards Opening Balance", false, true));
        //        //emailStr.Append(@"</td><td align=""middle"">");
        //        //emailStr.Append(MakeHeaderStr("Qualifying spend this month", true, true);
        //        emailStr.Append(@"</td><td align=""middle"" width=""200"" style=""border-style: outset; border-width: 1px"">");
        //        emailStr.Append(MakeHeaderStr("Gem Rewards Earned*", false, true));
        //        emailStr.Append(@"</td><td align=""middle"" width=""200"" style=""border-style: outset; border-width: 1px"">");
        //        emailStr.Append(MakeHeaderStr("Gem Rewards Redeemed", false, true));
        //        emailStr.Append(@"</td><td align=""middle"" width=""220"" style=""border-style: outset; border-width: 1px"">");
        //        emailStr.Append(MakeHeaderStr("Gem Rewards Closing Balance", false, true));
        //        emailStr.Append(@"</td></tr>");

        //        emailStr.Append(@"<tr><td align=""middle"" width=""220"">");
        //        emailStr.Append(MakeHeaderStr(basText.formatNum(rewardRow[mOpeningbalance], "#0.00;(#0.00)", 20, "M"), false, false));
        //        //emailStr.Append(@"</td><td align=""middle"">");
        //        //emailStr.Append(MakeHeaderStr("0", false, false);
        //        emailStr.Append(@"</td><td align=""middle"" width=""200"">");
        //        emailStr.Append(MakeHeaderStr(basText.formatNum(rewardRow["EARNEDBONUS"], "#0.00;(#0.00)", 20, "M"), false, false));
        //        emailStr.Append(@"</td><td align=""middle"" width=""200"">");
        //        emailStr.Append(MakeHeaderStr(basText.formatNum(rewardRow["REDEEMEDBONUS"], "#0.00;(#0.00)", 20, "M"), false, false));
        //        emailStr.Append(@"</td><td align=""middle"" width=""220"">");
        //        emailStr.Append(MakeHeaderStr(basText.formatNum(rewardRow[mClosingbalance], "#0.00;(#0.00)", 20, "M"), false, false));
        //        emailStr.Append(@"</td></tr></table></td></tr>");
        //        emailStr.Append(@"<tr><td width=""98%"" colSpan=""2"">&nbsp;</td></tr>");

        //        //streamWrit.WriteLine(basText.alignmentMiddle("RewardOpenBalance", 20) + basText.alignmentMiddle("RewardEarnedBonus", 20) + basText.alignmentMiddle("RewardRedeemedBonus", 20) + basText.alignmentMiddle("RewardClosingBalance", 20));
        //        //streamWrit.WriteLine(basText.formatNum(rewardRow[mOpeningbalance], "#0.00;(#0.00)", 20, "M") + basText.formatNum(rewardRow["EARNEDBONUS"], "#0.00;(#0.00)", 20, "M") + basText.formatNum(rewardRow["REDEEMEDBONUS"], "#0.00;(#0.00)", 20, "M") + basText.formatNum(rewardRow[mClosingbalance], "#0.00;(#0.00)", 20, "M"));
        //        //emailStr.Append(@"<tr><td width=""98%"" colSpan=""2""><font size=""2"">* Gem Rewards are only earned on POS and Internet purchases, and not on cash withdrawals. Refunded transactions are not eligible for Gem Rewards.</font></td></tr>");
        //        emailStr.Append(@"<tr><td width=""98%"" colSpan=""2""><font size=""2"">* Please save these numbers on your mobile phone, to contact us 24/7:01 - 6283892 and 01 - 2793500</font></td></tr>");
        //    }
        //}
        emailStr.Append(@"<tr><td width=""98%"" colSpan=""2"">&nbsp;</td></tr>");
        emailStr.Append(@"<a href=""http://" + strbankWebLink + @""">");
        emailStr.Append(@"<img border=""0"" src=""cid:" + clsBasFile.getFileFromPath(strBankMidBanner) + @"""></a></td></tr></table><table id=""table26"" width=""100%"" border=""0"">");
        //Banner
        //emailStr.Append(@"<tr><td width='98%' colSpan='2'><p align='center'><a href='http://www.diamondbank.com/display.php?mtd=diamond_visa_card_how_do_i_start.php?'><img border='0' src='cid:" + clsBasFile.getFileFromPath(bottomBanner) + @"'></a></td></tr>");
        emailStr.Append(@"<tr><td width='98%' colSpan='2'><p align='center'><a href='http://www.diamondbank.com/display.php?mtd=diamond_visa_card_how_do_i_start.php?'></a></td></tr>");

        emailStr.Append(@"<tr><td width=""98%"" colSpan=""2"">&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td width=""98%"" colSpan=""2"">&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td width=""98%"" colSpan=""2"">&nbsp;</td></tr>");


        //emailStr.Append(@"<tr><td width=""98%"" colSpan=""2""><a href=""http://www.diamondbank.com/products/retail-banking/visa-credit-cards/diamond-visa-credit-card-2010-fifa-world-cup.html"">");
        //emailStr.Append(@"<tr><td width=""98%"" colSpan=""2"">");//<a href=""http://www.diamondbank.com/products/retail-banking/visa-credit-cards/diamond-visa-credit-card-2010-fifa-world-cup.html"">
        //emailStr.Append(@"<font size=""2"">Go to the 2010 FIFA World Cup with your Diamond Bank Visa Credit Card. Click here for more info.</font></a>");
        //emailStr.Append(@"<font size=""2"">To prevent fraud on your card, it is critical to change your PIN if you have not already done so. PIN change can be done on any of Diamond Bank's ATMs nationwide.</font></a>");
        //emailStr.Append(@"<p><font size=""2"">If you wish to transact on the internet, you must have an internet PIN. Please call our contact center on 0808 2255 322 or 01-2793500 for more information and to activate your internet PIN.</font></td></tr>");
        //emailStr.Append(@"<br><font size=""2"">Please save these numbers on your mobile phone, to contact us 24/7:07003000000</font>");
        //emailStr.Append(@"<br><font size=""2"">To prevent fraud on your card, it is critical to change your PIN if you have not already done so. PIN change can be done on any of Diamond Bank's ATMs nationwide. </font>");
        //emailStr.Append(@"<br><font size=""2"">If you wish to transact on the internet, you must have an internet PIN. Please go to any Diamond Bank ATM and follow the steps below: </font>");
        //emailStr.Append(@"<br><font size=""2"">• Insert your Visa Card into a Diamond bank ATM </font>");
        //emailStr.Append(@"<br><font size=""2"">• Input your PIN </font>");
        //emailStr.Append(@"<br><font size=""2"">• Select 'more services' </font>");
        //emailStr.Append(@"<br><font size=""2"">• Select 'generate iPIN' </font>");
        //emailStr.Append(@"<br><font size=""2"">• Type in a 4 digit number of your choice </font>");
        //emailStr.Append(@"<br><font size=""2"">• Re-type the 4 digit number to confirm </font>");
        //emailStr.Append(@"<br><font size=""2"">• Select 'No' </font>");
        //emailStr.Append(@"<br><font size=""2"">• Collect your card and receipt </font>");
        //emailStr.Append(@"<br><font size=""2"">Based on your statement, please note that the 'Minimum Payment Now Due' is the minimum amount which must be credited to the card by the due date (i.e., date specified under the heading 'Please Pay this Amount Before') in order to keep your card active and in good order. Should you have elected a greater repayment amount at the time of card application or subsequently, the greater amount will be collected as agreed. </font></td></tr>");
        //emailStr.Append(@"<tr><td width=""98%"" colSpan=""2"">&nbsp;</td></tr>");

        //emailStr.Append(@"<tr><td width=""98%"" colSpan=""2""><a href=""http://www.diamondbank.com/display.php?mtd=diamond_visa_card_how_do_i_start.php?"">");
        //emailStr.Append(@"<font size=""2"">For more information about this statement, or using your Visa credit card. please click here.</font></a>");
        //emailStr.Append(@"<p><font size=""2"">If you have not found the information you require by following the link above, please call Diamond Bank on 0808-CALL-DBC (0808-2255-322) or 01-2793500</font></td></tr>");
        //emailStr.Append(@"<tr><td width=""98%"" colSpan=""2"">&nbsp;</td></tr>");

        emailStr.Append(@"<tr><td width=""98%""><p align=""left"">");
        emailStr.Append(@"<a href=""http://" + strbankWebLink + @""">");
        emailStr.Append(@"<img border=""0"" src=""cid:" + clsBasFile.getFileFromPath(strBankBottomBanner) + @"""></a></td></tr>");
        emailStr.Append(@"<br><a title="" +strbankWebLink+"" href=""http://" + strbankWebLink + @""">" + strbankWebLink + @"</a>");
        emailStr.Append(@"</td></tr>");
        emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"">&nbsp;</td></tr>");
        //    emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"">" + cryptAes.Encrypt(strBankName + " " + curAccountNumber + " " + extAccNum + " " + curClientID + " " + Convert.ToDateTime(masterRow[mStatementdateto]).ToString("dd/MM/yyyy"), strBankName+"&12345678", 192);
        emailStr.Append(@"<tr><td width='98%' colSpan='6'>&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td width='98%' colSpan='6'>&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td width='98%' colSpan='6'>&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td width='98%' colSpan='6'><font size='1'>" + clsCheckSum.stringMD5(extAccNum.Trim() + Convert.ToDateTime(masterRow[mStatementdateto]).ToString("dd/MM/yyyy")) + "</td></tr>");
        emailStr.Append(@"<tr><td width='98%' colSpan='6'>&nbsp;</td></tr>");
        emailStr.Append(@"</table></body></html>");
        //\
        //completePageDetailRecords();
        //-streamWrit.WriteLine(String.Empty);
        //-if (pageNo == totalAccPages)
        //-  streamWrit.WriteLine(basText.replicat(" ", 17) + basText.alignmentLeft("Current Balance", 63) + basText.formatNumUnSign(masterRow[mClosingbalance], "#0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])));
        //streamWrit.WriteLine(basText.replicat(" ", 80) + basText.formatNumUnSign(masterRow[mClosingbalance], "#0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])));
        //-else
        //-  streamWrit.WriteLine(String.Empty);
        //-streamWrit.WriteLine(String.Empty);
        //-streamWrit.WriteLine(String.Empty);
        //streamWrit.WriteLine(String.Empty);
        //-streamWrit.WriteLine(basText.formatNum(masterRow[mTotaldebits], "#0.00;(#0.00)", 20, "L") + basText.formatNum(masterRow[mTotalcredits], "#0.00;(#0.00)", 20, "L"));//streamWrit.WriteLine(String.Empty);
        //-streamWrit.WriteLine(basText.alignmentRight(curMainCard, 35) + basText.alignmentRight(masterRow[mMindueamount], 20) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "R")
        //-+ basText.alignmentRight(basText.formatNumUnSign(masterRow[mOpeningbalance], "#,###,##0.00", 12) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance].ToString())), 15)
        //-+ basText.alignmentRight(basText.formatNumUnSign(masterRow[mTotalpayments], "#,###,##0.00", 12) + " " + DbCr(Convert.ToDecimal(masterRow[mTotalpayments])), 15)
        //-+ basText.alignmentRight(basText.formatNum(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal]), "#,###,##0.00", 12) + CrDbMinus(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal])), 15)
        //-+ basText.alignmentRight(basText.formatNumUnSign(masterRow[mTotalcharges].ToString(), "#,###,##0.00", 12) + " " + CrDbMinus(Convert.ToDecimal(masterRow[mTotalcharges])), 15)
        //-+ basText.alignmentRight(basText.formatNumUnSign(masterRow[mTotalinterest].ToString(), "#,###,##0.00", 12) + " " + CrDbMinus(Convert.ToDecimal(masterRow[mTotalinterest])), 15)
        //       + basText.formatNum(Convert.ToDecimal(masterRow[mTotalcharges]) + Convert.ToDecimal(masterRow[mTotalinterest]),"#,##0.00",12) + CrDbMinus(Convert.ToDecimal(masterRow[mTotalcharges]) + Convert.ToDecimal(masterRow[mTotalinterest])) 

        //-+ basText.alignmentRight(basText.formatNumUnSign(masterRow[mClosingbalance], "#,###,##0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])), 15));
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
            //streamWrit.WriteLine(basText.trimStr(dtRow[dTrandescription],40)); 
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

            /*>if(prevCardNo != CurCardNo && prevCardNo != String.Empty)
                {//if there are page for every card inside account pages
                    totalAccPages++;
                    currAccRowsPages =1;
                }*/
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


    private void printStatementSummary()
    {
        streamSummary.WriteLine(strBankName + " Transactions Details Summary");
        //streamSummary.WriteLine(strBankName + " E-Statement Summary");
        streamSummary.WriteLine("_____________________");
        //streamSummary.WriteLine("__________________________");
        streamSummary.WriteLine("");
        streamSummary.WriteLine("Sent by Email   " + noOfEmails.ToString());
        //streamSummary.WriteLine("Statements Sent by Email   " + noOfEmails.ToString());
        streamSummary.WriteLine("have bad Email  " + noOfBadEmails.ToString());
        //streamSummary.WriteLine("Statements have bad Email  " + noOfBadEmails.ToString());
        streamSummary.WriteLine("without Email  " + noOfWithoutEmails.ToString());
        //streamSummary.WriteLine("Statements without Email  " + noOfWithoutEmails.ToString());
        //streamSummary.WriteLine("No of Transactions " + totNoOfTransactions.ToString());

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

    private string MakeHeaderStr(string pStr, bool isBold, bool isHeader)
    {
        string color = string.Empty;
        pStr = pStr.Replace("&", "&amp;").Trim();
        if (isHeader)
            color = "#000080";
        else
            color = "#800000";

        if (pStr.Length < 1)
            pStr = " ";//"&nbsp;"
        else
        {
            pStr = @"<font size=""2"" color=""" + color + @""">" + pStr + @"</font>";
            if (isBold)
                pStr = @"<b>" + pStr + "</b>";
        }

        return pStr;
    }

    private string MakeHeaderStr2(string pStr, bool isBold, bool isHeader)
    {
        string color = string.Empty;
        pStr = pStr.Replace("&", "&amp;").Trim();
        if (isHeader)
            color = "#000000";
        else
            color = "#FF8C00";

        if (pStr.Length < 1)
            pStr = " ";//"&nbsp;"
        else
        {
            pStr = @"<font size=""2"" color=""" + color + @""">" + pStr + @"</font>";
            if (isBold)
                pStr = @"<b>" + pStr + "</b>";
        }

        return pStr;
    }

    private string ValidateStr(string pStr)
    {
        if (pStr.Trim().Length < 1)
            pStr = " ";//"&nbsp;"
        return pStr;
    }

    private void SendEmail(string pBody, string pSubject, string pTo)
    {
        try{

        //if (pTo.IndexOf('@') == -1 || pTo.IndexOf('.') == -1)
            if (valdEmail.isValideEmail(pTo) != "ValidEmail")
            {
                //if (emailTo.EndsWith("@intercontinentalbankplc"))
                //    emailTo = emailTo + ".com";
                //else
                //{
                streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + pTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "| Bad Email");
                noOfBadEmails++;
                pTo = strEmailFrom; //"statement_Program@emp-group.com";
                emailLabelTmp = emailLabel + " Acc:" + curAccountNumber + " Phone:" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!");

                //}
                //return;
            }
        //if (!pTo.ToLower().EndsWith("diamondbank.com"))
        //    return;
        //if (pTo.ToUpper().IndexOf("YAHOO") > -1 || pTo.ToUpper().IndexOf("HOTMAIL") > -1)
        //  pBody.Replace("", "");

        //try
        //{
            //string pFrom, pSubject, pBody;
            //lblStatus.Text = string.Empty;
            //pFrom = "mmohammed@emp-group.com";//mmohammed@mscc.local"mmohammed@emp-group.com"mmohammed@emp-group.com
            //return;

            //streamEmails.WriteLine(curAccountNumber + "|" + curCardNumber + "|" + curClientID + "|" + emailTo);
            //streamEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
            //return;
            //string[] strAray;
            //if (emailTo != "lucky.ighade@zenithbank.com")//ehimolen@yahoo.com
            //return;
            ArrayList pLstTo = new ArrayList(), pLstAttachedFile = new ArrayList();//, pLstCC = new ArrayList(), pLstBCC = new ArrayList();
            pSubject = emailLabelTmp; //emailLabel "BAI statement for 02/2008";
            //pTo = emailTo;// "mmohammed@emp-group.com";
            if (pTo.EndsWith("."))//ehimolen@yahoo.com
                pTo = pTo + "com";
            pLstTo.Add(pTo);// "mmohammed@emp-group.com" "mmohammed@emp-group.com" "nazab@emp-group.com""wbaioumy@emp-group.com""mmohammed@emp-group.com" "mmohammed@emp-group.com""nazab@emp-group.com"
            //if (!pTo.ToUpper().EndsWith("DIAMONDBANK.COM"))//ehimolen@yahoo.com
            //  return;
            //pLstBCC.Add("mmohammed@emp-group.com");
            //pLstBCC.Add("nazab@emp-group.com");
            //pLstTo.Add("mmohammed@emp-group.com");//"mmohammed@emp-group.com""nazab@emp-group.com"


            //pLstTo.Add("ashorungbe@yahoo.com");
            //pLstTo.Add("adeyinka.shorungbe@zenithbank.com");
            //pLstTo.Add("nazab@emp-group.com");
            //pLstTo.Add("hfawzy@emp-group.com");
            //pLstTo.Add("mmohammed@emp-group.com");//mmohammed@mscc.localmmohammed@emp-group.com nazab@emp-group.commmohammed@emp-group.com

            //pLstAttachedFile.Add(bottomBanner);//strOutputFile@"D:\Web\Email\BAI\Logo.gif"
            pLstAttachedFile.Add(strBankLogo);//strOutputFile@"D:\Web\Email\BAI\Logo.gif"
            pLstAttachedFile.Add(strBankMidBanner);//strOutputFile@"D:\Web\Email\BAI\Logo.gif"
            pLstAttachedFile.Add(strBankBottomBanner);//strOutputFile@"D:\Web\Email\BAI\Logo.gif"
            //pLstAttachedFile.Add(backGround);//@"D:\pC#\exe\FilesForPrograms\frmBackground.jpg"
            sndMail = new clsEmail();
            sndMail.emailFromName = emailFromNameStr;
            //if (!sndMail.sendEmailHTML(strEmailFrom, pLstTo, pLstCC, pLstBCC, pLstAttachedFile, pSubject, pBody, clsCnfg.readSetting("SmtpServer")))
            numOfTry = 0;
            noOfEmails++;
            //while (!sndMail.sendEmailHTML(strEmailFrom, pLstTo, pLstCC, pLstBCC, pLstAttachedFile, pSubject, pBody, clsCnfg.readSetting("SmtpServer")))
            strEmailFromTmp = strEmailFrom;
            //if (strEmailFrom == emailTo)
            //{
            //    strEmailFromTmp = "mabouleila@emp-group.com";
            //}
            if (numMail == 0)
            {
                //pLstBCC.Add("mmohammed@emp-group.com");
                pLstBCC.Add("statement@emp-group.com");
            }

            while (!sndMail.sendEmailAttachPic(strEmailFromTmp, pLstTo, pLstCC, pLstBCC, pLstAttachedFile, pSubject, pBody, clsCnfg.readSetting("SmtpServer")))
            {
                //MessageBox.Show("Failure to send Email.", "Send Email Error",MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                System.Threading.Thread.Sleep(waitPeriodVal);//3000
                numOfTry++;
                if (numOfTry > 100)
                //if (numOfTry > 1)
                {
                    streamEmails.Write("\t\t >> Error while Sending Email");
                    streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|Exceed number of trials");
                    noOfBadEmails++; noOfEmails--;
                    break;
                }
            }
            streamEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + pTo + "|" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));

            if (numMail == 0)
                pLstBCC.Clear();

            numMail++;
            if (numMail % 400 == 0)
            {
                System.Threading.Thread.Sleep(waitPeriodVal);//2000
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            //sndMail = null;
            //emailStr = new StringBuilder(""); emailTo = string.Empty; curAccountNumber = string.Empty; curCardNumber = string.Empty; curClientID = string.Empty;
            //wait2NextEmail(prvEmail, pTo, waitPeriodVal);//     System.Threading.Thread.Sleep(waitPeriodVal);//400
            //prvEmail = pTo;
        }
        catch (Exception ex) //(NotSupportedException ex) (Exception ex)  //
        {
            //clsBasErrors.catchError(ex);
            streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + pTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|Email " + pTo + " , Err Message >> " + ex.Message + ", Err Desc >> ");
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

    public string VIPCondition
    {
        get { return vipCondVal; }
        set { vipCondVal = value; }
    }  // VIPCondition


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

    public string bankMidBanner
    {
        get { return strBankMidBanner; }
        set { strBankMidBanner = value; }
    }// bankMidBanner

    public string bankBottomBanner
    {
        get { return strBankBottomBanner; }
        set { strBankBottomBanner = value; }
    }// bankBottomBanner

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

    ~clsStatementDBNSup_Html()
    {
        DSstatement.Dispose();
    }
}
