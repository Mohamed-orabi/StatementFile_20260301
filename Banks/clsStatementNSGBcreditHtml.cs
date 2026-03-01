using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Xml;
using System.Collections;


public class clsStatementNSGBcreditHtml : clsBasStatement
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
    //private clsOMR omr = new clsOMR();
    private int totPages;
    string endOfCustomer = string.Empty;
    string cProduct = string.Empty, curFileName = string.Empty;
    private ArrayList aryLstFiles;
    private string strProductCond = string.Empty;
    //
    //private string emailStr = string.Empty;
    private StringBuilder emailStr = new StringBuilder("");
    private string emailLabel, strBankLogo;
    private string strEmailFrom = "statement@emp-group.com", strEmailFromTmp = string.Empty;  //"cardservices@zenithbank.com"
    private string strbankWebLink = "www.emp-group.com";  //www.zenithbank.com
    private string strbankWebLinkService = "www.emp-group.com";  //www.zenithbank.com
    private string emailTo = string.Empty, curAccountNumber, curCardNumber, curClientID;//, curEmail
    private string strEmailFileName, strNoEmailFileName;
    private frmStatementFile frmMain;
    private int totRec = 1, numMail, numOfTry = 0;
    private ArrayList pLstCC = new ArrayList(), pLstBCC = new ArrayList();
    private string logoAlignmentStr = "center";
    private string strbackGround = @"D:\pC#\ProjData\Statement\_Background\Background06.jpg";
    private string strbottomBanner = string.Empty;
    private string strbottomBannerDefault = string.Empty;
    private string strbottomBannerPlatinum = string.Empty;
    private bool isRewardVal = false;
    private bool IsSplittedVal = false;
    private bool HasAttachementVal = false;
    private bool HasSenderVal = false;
    private bool createCorporateVal = false;
    private string accountNoName = mAccountno;
    private string accountLimit = mAccountlim;
    private string accountAvailableLimit = mAccountavailablelim;
    private string rewardCondVal = "'New Reward Contract'";//'Reward Contract'
    private string strVisaLogo = string.Empty;
    private string statMessageFileVal = string.Empty;
    private string statMessageBoxVal = string.Empty;
    private string statMessage = "&nbsp;";//Null
    private string statBoxMessage = "&nbsp;";//Null
    private string statMessageFileMonthlyVal = string.Empty;
    private string statMessageMonthly = "&nbsp;";//Null
    private string isSuplStr = string.Empty;//Null
    private DataRow[] emailRows = null;
    private DataRow masterRelatedRow;
    private clsEmail sndMail;
    private string emailFromNameStr;
    private string attachedFilesStr;
    private string Address1Name = mCustomeraddress1, Address2Name = mCustomeraddress2, Address3Name = mCustomeraddress3;
    private string strMobileNum = string.Empty;
    private int waitPeriodVal = 7000;
    private clsAes cryptAes = new clsAes();
    private int noOfEmails, noOfBadEmails, noOfWithoutEmails;
    private string emailLabelTmp;  //"cardservices@zenithbank.com"
    private DataRow emailRow;
    private string strFileNam, stmntType;
    private string prvEmail = string.Empty;

    private int installdocno, intinstalldocno, installentryno, intinstallentryno;
    private decimal installamount, intinstallamount;

    private int accinstalldocno, accintinstalldocno, accinstallentryno, accintinstallentryno;
    private decimal accinstallamount, accintinstallamount;

    protected string ClientID = mClientid;

    public clsStatementNSGBcreditHtml()
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
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + "Begin Statement Generation for clsStatementNSGBcreditHtml", System.IO.FileMode.Append);

            //clsMaintainData maintainData = new clsMaintainData();
            //maintainData.matchCardBranch4Account(pBankCode);

            pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);
            vCurDate = pCurDate; //DateTime.Now.AddMonths(-1);
            strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "_" + pStmntType;
            //clsBasFile.deleteDirectory(strOutputPath);
            clsBasFile.createDirectory(strOutputPath);
            strEmailFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "_" + pStmntType + "\\" + pBankName + "_" + pStmntType + "_" + vCurDate.ToString("yyyyMM");//+ ".txt"
            strNoEmailFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "_" + pStmntType + "\\" + pBankName + "_" + pStmntType + "_" + "WithoutEmails" + "_" + vCurDate.ToString("yyyyMM");//+ ".txt"
            pStrFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "_" + pStmntType + "\\" + pBankName + "_" + pStmntType + "_" + vCurDate.ToString("yyyyMM") + ".txt";
            strBankName = pBankName;
            emailLabel = pBankName + " statement for " + vCurDate.ToString("MM/yyyy"); //"BAI statement for 02/2008"
            strOutputFile = pStrFileName;

            // open emails file
            fileEmails = new FileStream(strEmailFileName + ".txt", FileMode.Create); //Create
            streamEmails = new StreamWriter(fileEmails, Encoding.Default);
            //streamEmails.WriteLine("AccountNumber" + "|" + "CardNumber" + "|" + mClientid + "|" + "Email");
            streamEmails.WriteLine("AccountNumber" + "|" + "ClientID" + "|" + "Email" + "|" + "MobilePhone" + "|" + "Date Time");
            streamEmails.AutoFlush = true;

            // open No emails file
            fileNoEmails = new FileStream(strNoEmailFileName + ".txt", FileMode.Create); //Create
            streamNoEmails = new StreamWriter(fileNoEmails, Encoding.Default);
            streamNoEmails.WriteLine("AccountNumber" + "|" + "ClientID" + "|" + "Email" + "|" + "Mobile Phone");//"AccountNumber" + "|" + "ClientID"
            streamNoEmails.AutoFlush = true;

            if (!string.IsNullOrEmpty(statMessageFileVal))
            {
                FileStream filRead = null;
                StreamReader filStream = null;
                filRead = new FileStream(statMessageFileVal, FileMode.Open);
                filStream = new StreamReader(filRead, Encoding.Default);
                statMessage = filStream.ReadToEnd();
                filStream.Close();
                filRead.Close();
            }

            if (!string.IsNullOrEmpty(statMessageBoxVal))
            {
                FileStream filRead = null;
                StreamReader filStream = null;
                filRead = new FileStream(statMessageBoxVal, FileMode.Open);
                filStream = new StreamReader(filRead, Encoding.Default);
                statBoxMessage = filStream.ReadToEnd();
                filStream.Close();
                filRead.Close();
            }

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
                Address1Name = mCardaddress1;
                Address2Name = mCardaddress2;
                Address3Name = mCardaddress3;
                ClientID = mCardClientId;
            }

            if (IsSplitted)
            {
                MainTableCond = " m.cardproduct = '" + strProductCond + "'";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.cardproduct = '" + strProductCond + "') and d.trandescription not in ('Calcualated _Points','Exchange bonuses for prize','Loyalty Points Redemption','Bonuses expiration')";
                if (isRewardVal)
                {
                    curRewardCond = rewardCondVal;
                    //strMainTableCond = "m.contracttype not in " + rewardCondVal;
                    //strSubTableCond = "d.trandescription not in ('Calcualated _Points','Exchange bonuses for prize','LOY_CASHBACK')";//
                    getReward(pBankCode);
                    MainTableCond = " m.cardproduct = '" + strProductCond + "'";//strWhereCond
                    supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and cardproduct = '" + strProductCond + "' union all select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and contracttype = 'Reward Program - " + strProductCond + "')";
                }
            }



            FillStatementDataSet(pBankCode, "vip"); //DSstatement =  //10); //3
            getClientEmail(pBankCode);
            pageNo = 1; totalCardPages = 0;
            curCardNo = String.Empty;
            curAccountNo = String.Empty;
            frmMain.BeginInvoke(frmMain.setMinMaxProgressDelegate, new object[] { DSstatement.Tables["tStatementMasterTable"].Rows.Count });
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + "foreach begin for DSstatement.Tables[tStatementMasterTable].Rows ", System.IO.FileMode.Append);

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

                //start new account
                if (prevAccountNo != masterRow[accountNoName].ToString())
                {
                    if (!string.IsNullOrEmpty(emailStr.ToString()) && !string.IsNullOrEmpty(emailTo)) //emailStr != string.Empty && emailStr != null  emailTo != string.Empty && emailTo != null
                        //if (HasSender == true)
                        //    SendEmailWithDifferentSender(emailStr.ToString(), "", emailTo);
                        //else
                        SendEmail(emailStr.ToString(), "", emailTo);
                    if (!string.IsNullOrEmpty(emailStr.ToString()) && string.IsNullOrEmpty(emailTo)) //emailStr != string.Empty && emailStr != null  emailTo != string.Empty && emailTo != null
                    {
                        streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|Without Email");
                        noOfWithoutEmails++;
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
                    emailRows = DSemails.Tables["Emails"].Select("idclient = " + masterRow[ClientID].ToString());
                    //clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + "foreach emailRows", System.IO.FileMode.Append);

                    for (int i = 0; i < emailRows.Length; i++)
                    {
                        if (!string.IsNullOrEmpty(emailRows[i][1].ToString().Trim()))
                        {
                            emailTo = emailRows[i][1].ToString().Trim();
                            emailRow = emailRows[i];
                        }

                        if (!string.IsNullOrEmpty(emailRows[i][2].ToString().Trim()))
                        {
                            strMobileNum = emailRows[i][2].ToString().Trim();
                        }
                    }
                    //emailTo = "MABOULEILA@EMP-GROUP.com";
                    //>>>emailTo = masterRow[mDept].ToString().Trim();//>>

                    curAccountNumber = masterRow[accountNoName].ToString();
                    curCardNumber = strPrimaryCardNo;
                    curClientID = masterRow[ClientID].ToString();

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

                //if (isSuplStr == "SUP")
                //  cardSeparator();

                //foreach (DataRow dRow in mRow.GetChildRows(StatementNoDRel))
               // clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + "foreach cardrpws", System.IO.FileMode.Append);

                foreach (DataRow dRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
                {
                    detailRow = dRow;
                    isSuplStr = string.Empty;
                    if (curMainCard != detailRow[dCardno].ToString().Trim())
                    {
                        masterRelatedRow = DSstatement.Tables["tStatementMasterTable"].Select("statementno = '" + clsBasValid.validateStr(detailRow[dStatementno]).ToString().Trim() + "'")[0];
                        isSuplStr = isSupplementCard(clsBasValid.validateStr(masterRelatedRow[mCardprimary].ToString()));
                        //if(isSuplStr != "SUP")
                        //  if (!isValidateCard(masterRelatedRow[mCardstate].ToString()))
                        //    isSuplStr = "By_" + clsBasValid.validateStr(masterRelatedRow[mCardstatus].ToString()) + "_Card";
                    }


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
                    if (detailRow[dTrandescription].ToString() == "Charge interest for Installment")
                    {
                        long intinstalldocno = long.Parse(detailRow[dDocno].ToString());
                        long intinstallentryno = long.Parse(detailRow[dEntryNo].ToString());
                        intinstallamount = decimal.Parse(detailRow[dBilltranamount].ToString());
                    }
                    if (detailRow[dTrandescription].ToString() == "Installment repayment")
                    {
                        long installdocno = long.Parse(detailRow[dDocno].ToString());
                        long installentryno = long.Parse(detailRow[dEntryNo].ToString());
                        installamount = decimal.Parse(detailRow[dBilltranamount].ToString());
                    }
                    if (intinstalldocno != 0 && installdocno != 0 && intinstallentryno == installentryno - 1)
                    {
                        detailRow[dBilltranamount] = intinstallamount + installamount;
                        installdocno = 0;
                        intinstalldocno = 0;
                        installentryno = 0;
                        intinstallentryno = 0;
                        intinstallamount = 0;
                        installamount = 0;
                    }
                    //if (detailRow[dTrandescription].ToString() == "Charge interest for Acceleration")
                    //    {
                    //    accintinstalldocno = int.Parse(detailRow[dDocno].ToString());
                    //    accintinstallentryno = int.Parse(detailRow[dEntryNo].ToString());
                    //    accintinstallamount = decimal.Parse(detailRow[dBilltranamount].ToString());
                    //    }
                    //if (detailRow[dTrandescription].ToString() == "Installment Acceleration")
                    //    {
                    //    accinstalldocno = int.Parse(detailRow[dDocno].ToString());
                    //    accinstallentryno = int.Parse(detailRow[dEntryNo].ToString());
                    //    accinstallamount = decimal.Parse(detailRow[dBilltranamount].ToString());
                    //    }
                    //if (accintinstalldocno != 0 && accinstalldocno != 0 && accintinstallentryno == accinstallentryno - 1)
                    //    {
                    //    detailRow[dBilltranamount] = accintinstallamount + accinstallamount;
                    //    accinstalldocno = 0;
                    //    accintinstalldocno = 0;
                    //    accinstallentryno = 0;
                    //    accintinstallentryno = 0;
                    //    accintinstallamount = 0;
                    //    accinstallamount = 0;
                    //    }
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
                //-}
                //streamWrit.WriteLine(strEndOfPage);
                //pageNo=1; CurPageRec4Dtl=0; //if pages is based on card
                //completePageDetailRecords();
                //SendEmail(emailStr, "", "");
            } //end of Master foreach
            //emailStr = emailStr.Trim();
            emailTo = emailTo.Trim();
            if (!string.IsNullOrEmpty(emailStr.ToString()) && !string.IsNullOrEmpty(emailTo)) //emailTo != string.Empty && emailTo != null
                //if (HasSender == true)
                //    SendEmailWithDifferentSender(emailStr.ToString(), "", emailTo);
                //else
                SendEmail(emailStr.ToString(), "", emailTo);
            if (!string.IsNullOrEmpty(emailStr.ToString()) && string.IsNullOrEmpty(emailTo)) //emailStr != string.Empty && emailStr != null  emailTo != string.Empty && emailTo != null
            {
                streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|Without Email");
                noOfWithoutEmails++;
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
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

            clsDbOracleLayer.catchError(ex);
        }
        catch (NotSupportedException ex)  //(Exception ex)  //
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

            clsBasErrors.catchError(ex);
        }
        catch (Exception ex)
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

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

            streamNoEmails.Flush();
            streamNoEmails.Close();
            fileNoEmails.Close();
            //ArrayList aryLstFiles = new ArrayList();
            //aryLstFiles.Add("");
            //aryLstFiles.Add(@strOutputFile);
            //-numOfErr += validateNoOfLines(aryLstFiles, 48);
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
        if (extAccNum.Trim() == "")
            extAccNum = clsBasValid.validateStr(masterRow[accountNoName]);

        emailStr = new StringBuilder("");
        emailStr.Append(@"<html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1252"">");
        emailStr.Append(@"<title>");
        emailStr.Append(masterRow[mCustomername].ToString() + " - Statement");
        emailStr.Append(@"</title>");
        emailStr.Append(@"</head><body  background='cid:" + clsBasFile.getFileFromPath(strbackGround) + @"'>");
        emailStr.Append(@"<table id=""table15"" width=""100%"" border=""0"">");
        emailStr.Append(@"<tr><td height=""82"" width=""27%"" colspan = ""6"" style='font-family:cordale arabic'>" + statBoxMessage + "</td></tr>");
        emailStr.Append(@"<tr><td height=""82"" width=""27%"" colspan = ""6"" style='font-family:cordale arabic'>");
        emailStr.Append(@"<p align=""" + logoAlignmentStr + @""">");
        emailStr.Append(@"<a href=""http://" + strbankWebLink + @""">");
        emailStr.Append(@"<img border=""0"" src='cid:" + clsBasFile.getFileFromPath(strBankLogo) + @"'/></a>");// width=""230"" height=""85""
        emailStr.Append(@"</td>");
        emailStr.Append(@"</tr><tr>");
        emailStr.Append(@"<td width=""50%"" bordercolor=""#FFFFFF"" colspan=""3"" style='font-family:cordale arabic'><table id=""table16"" height=""119"" width=""99%"" border=""1"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF"" cellspacing=""0""><tbody><tr><td width=""90%"" style='font-family:cordale arabic'>");
        emailStr.Append(@"<table id=""table17"" borderColor=""#FFFFFF"" bgcolor=""#DCD5CF"" width=""100%"" border=""0""><tbody>");
        emailStr.Append(@"<tr><td align=""middle"" style='font-family:cordale arabic'>");
        emailStr.Append(MakeHeaderStr(masterRow[mCustomername].ToString(), false, false));
        emailStr.Append(@"</td></tr><tr><td align=""middle"" style='font-family:cordale arabic'>");
        string newaddress1, newaddress2;
        clsMaintainData.fixAddress(curBranchVal, masterRow[mCustomeraddress1].ToString(), out newaddress1, out newaddress2);
        emailStr.Append(MakeHeaderStr(ValidateArbicNum(ValidateArbic(newaddress1)), false, false));
        emailStr.Append(@"</td></tr><tr><td align=""middle"" style='font-family:cordale arabic'>");
        emailStr.Append(MakeHeaderStr(ValidateArbicNum(ValidateArbic(newaddress2)), false, false));
        emailStr.Append(@"</td></tr><tr><td align=""middle"" style='font-family:cordale arabic'>");
        if (createCorporateVal)
            emailStr.Append(MakeHeaderStr(ValidateArbicNum(ValidateArbic(masterRow[Address3Name].ToString())) + "  " + masterRow[mCustomerregion].ToString().Trim() + " " + masterRow[mCustomercity].ToString().Trim(), false, false));
        else
            emailStr.Append(MakeHeaderStr(ValidateArbicNum(ValidateArbic(masterRow[Address3Name].ToString())), false, false));
        emailStr.Append(@"</td></tr><tr><td align=""middle"" style='font-family:cordale arabic'>");
        if (createCorporateVal)
            emailStr.Append(MakeHeaderStr(masterRow[mCardclientname].ToString().Trim(), false, false));
        else
            emailStr.Append(MakeHeaderStr(masterRow[mCustomerregion].ToString().Trim() + " " + masterRow[mCustomercity].ToString().Trim(), false, false));
        emailStr.Append(@"</td></tr></tbody></table></td></tr></tbody></table></td>");
        emailStr.Append(@"<td width=""50%"" colspan=""3"" style='font-family:cordale arabic'><table id=""table18"" height=""119"" width=""100%"" border=""1"" bordercolor=""#FFFFFF"" cellspacing=""0""><tbody><tr><td width=""100%"" style='font-family:cordale arabic'><table id=""table19"" height=""112"" width=""100%"" border=""0""><tbody>");
        emailStr.Append(@"<tr><td bgcolor=""#4F5858"" style='font-family:cordale arabic;color:#FFFFFF'>");
        emailStr.Append(MakeHeaderStr("Card Type:", true, false));
        emailStr.Append(@"</td><td height=""27"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
        emailStr.Append(MakeHeaderStr(masterRow[mCardproduct].ToString(), false, false));
        emailStr.Append(@"</td></tr><tr><td width=""139"" bgcolor=""#4F5858"" style='font-family:cordale arabic;color:#FFFFFF'>");
        emailStr.Append(MakeHeaderStr("Branch:", true, false));
        emailStr.Append(@"</td><td bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
        emailStr.Append(MakeHeaderStr(masterRow[mCardbranchpartname].ToString(), false, false));
        emailStr.Append(@"</td></tr><tr><td width=""139"" bgcolor=""#4F5858"" style='font-family:cordale arabic;color:#FFFFFF'>");
        emailStr.Append(MakeHeaderStr("Bank Account No.:", true, false));
        emailStr.Append(@"</td><td bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
        emailStr.Append(MakeHeaderStr(extAccNum, false, false));
        emailStr.Append(@"</td></tr><tr><td width=""139"" bgcolor=""#4F5858"" style='font-family:cordale arabic;color:#FFFFFF'>");
        emailStr.Append(MakeHeaderStr("Statement Date:", true, false));
        emailStr.Append(@"</td><td bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
        emailStr.Append(MakeHeaderStr(Convert.ToDateTime(masterRow[mStatementdateto]).ToString("dd/MM/yyyy"), false, false));
        emailStr.Append(@"</td></tr></tbody></table></td></tr></tbody></table></td></tr>");
        emailStr.Append(@"<tr><td width=""40%"" colSpan=""6"" height=""10"" style='font-family:cordale arabic'> </td></tr><tr><td width=""100%"" colSpan=""6"" style='font-family:cordale arabic'><img border=""0"" src='cid:" + clsBasFile.getFileFromPath(strVisaLogo) + @"'></td></tr>"); //&nbsp;
        //NSGB-3444
        emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"" height=""65"" style='font-family:cordale arabic'><table id=""table20"" height=""42"" width=""100%"" border=""0""><tbody>");
        emailStr.Append(@"<tr><td align=""middle"" width=""134"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><b><font size=""2"" color=""#FFFFFF"">Card Number</font></b></td>");
        emailStr.Append(@"<td align=""middle"" width=""134"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><b><font size=""2"" color=""#FFFFFF"">Credit Limit</font></b></td>");
        emailStr.Append(@"<td align=""middle"" width=""141"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><b><font size=""2"" color=""#FFFFFF"">Available Credit</font></b></td>");
        emailStr.Append(@"<td align=""middle"" width=""145"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><b><font size=""2"" color=""#FFFFFF"">Min. Payment</b></font></td>");
        emailStr.Append(@"<td align=""middle"" width=""165"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><b><font size=""2"" color=""#FFFFFF"">Payment Due Date</font></b></td>");
        emailStr.Append(@"<td align=""middle"" width=""148"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><b><font size=""2"" color=""#FFFFFF"">Past Due Amount</font></b></td>");
        emailStr.Append(@"</tr><tr><td align=""middle"" width=""134"" height=""16"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
        //NSGB-3444
        //if (strPrimaryCardNo.Trim() == "")
        //    emailStr.Append(MakeHeaderStr(strPrimaryCardNo, false, false));
        //else
        //    emailStr.Append(MakeHeaderStr(strPrimaryCardNo.Substring(0, 6) + "******" + strPrimaryCardNo.Substring(12, 4), false, false));
        //NSGB-3528
        if (curMainCard.Trim() == "")
            emailStr.Append(MakeHeaderStr(curMainCard, false, false));
        else
            emailStr.Append(MakeHeaderStr(curMainCard.Substring(0, 6) + "******" + curMainCard.Substring(12, 4), false, false));
        emailStr.Append(@"</td><td align=""middle"" width=""134"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
        emailStr.Append(MakeHeaderStr(masterRow[accountLimit].ToString().Trim(), false, false));
        emailStr.Append(@"</td><td align=""middle"" width=""141"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
        //emailStr.Append(MakeHeaderStr(basText.formatNum(masterRow[mAccountavailablelim], "##0", 20).Trim(), false, false));
        //NSGB-2889 EDT-1334
        if (createCorporateVal)
            emailStr.Append(MakeHeaderStr(basText.formatNum(masterRow[mCardavailablelimit], "##0", 20).Trim(), false, false));
        else
            emailStr.Append(MakeHeaderStr(basText.formatNum((decimal.Parse(masterRow[mAccountavailablelim].ToString()) - decimal.Parse(masterRow[mInstallmentUsedLimit].ToString())), "##0.00", 20).Trim(), false, false));

        emailStr.Append(@"</td><td align=""middle"" width=""145"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
        emailStr.Append(MakeHeaderStr(masterRow[mMindueamount].ToString().Trim(), true, false));
        emailStr.Append(@"</td><td align=""middle"" width=""165"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
        emailStr.Append(MakeHeaderStr(basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M"), true, false));
        emailStr.Append(@"</td><td align=""middle"" width=""148"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
        emailStr.Append(MakeHeaderStr(basText.formatNum(masterRow[mTotaloverdueamount], "##0.00", 13).Trim(), false, false));
        emailStr.Append(@"</td></tr></tbody></table></td></tr>");
        //if (isRewardVal)
        //    {
        //    rewardRows = DSreward.Tables["Reward"].Select("CLIENTID = '" + clsBasValid.validateStr(masterRow[mClientid]).ToString().Trim() + "'");
        //    if (rewardRows.Length > 0)
        //        {
        //        rewardRow = rewardRows[0];
        //        emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"" height=""65"" style='font-family:cordale arabic'><table id=""table22"" height=""42"" width=""100%"" border=""0""><tbody>");
        //        emailStr.Append(@"<tr><td align=""middle"" width=""134"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><b><font size=""2"" color=""#FFFFFF"">Reward Opening Balance</font></b></td>");
        //        emailStr.Append(@"<td align=""middle"" width=""141"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><b><font size=""2"" color=""#FFFFFF"">Total Points Earned</font></b></td>");
        //        emailStr.Append(@"<td align=""middle"" width=""145"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><b><font size=""2"" color=""#FFFFFF"">Total Points Reedemed</b></font></td>");
        //        emailStr.Append(@"<td align=""middle"" width=""165"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><b><font size=""2"" color=""#FFFFFF"">Expired Points</font></b></td>");
        //        emailStr.Append(@"<td align=""middle"" width=""148"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><b><font size=""2"" color=""#FFFFFF"">Reward Closing Balance</font></b></td>");
        //        emailStr.Append(@"</tr><tr><td align=""middle"" width=""134"" height=""16"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
        //        emailStr.Append(MakeHeaderStr(basText.formatNum(rewardRow[mOpeningbalance], "#0.00;(#0.00)", 20, "M"), false, false));
        //        emailStr.Append(@"</td><td align=""middle"" width=""141"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
        //        emailStr.Append(MakeHeaderStr(basText.formatNum(rewardRow[mEarnedBonus], "#0.00;(#0.00)", 20, "M"), false, false));
        //        emailStr.Append(@"</td><td align=""middle"" width=""145"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
        //        emailStr.Append(MakeHeaderStr(basText.formatNum(rewardRow[mRedeemedBonus], "#0.00;(#0.00)", 20, "M"), false, false));
        //        emailStr.Append(@"</td><td align=""middle"" width=""165"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
        //        emailStr.Append(MakeHeaderStr(basText.formatNum(rewardRow[mExpiredBonus], "#0.00;(#0.00)", 20, "M"), false, false));
        //        emailStr.Append(@"</td><td align=""middle"" width=""148"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
        //        emailStr.Append(MakeHeaderStr(basText.formatNum(rewardRow[mClosingbalance], "#0.00;(#0.00)", 20, "M"), false, false));
        //        emailStr.Append(@"</td></tr></tbody></table></td></tr>");                
        //        }
        //    }
        emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"" height=""33"" style='font-family:cordale arabic'>&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"" height=""88"" style='font-family:cordale arabic'><table id=""table21"" width=""100%"" border=""0""><tbody>");
        emailStr.Append(@"<tr bgcolor=""#4F5858""><td align=""middle"" colSpan=""2"" height=""17"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><font size=""2"" color=""#FFFFFF""><span style=""background-color: #4F5858"">Dates of</span></font></td>");
        emailStr.Append(@"<td align=""middle"" width=""67%"" colSpan=""2"" height=""18"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><b><font size=""2"" color=""#FFFFFF"">Description</font></b></td>");
        emailStr.Append(@"<td align=""middle"" width=""16%"" colSpan=""2"" height=""18"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'"">");
        emailStr.Append(@"<b><font size=""2"" color=""#FFFFFF"">Amount (" + masterRow[mAccountcurrency].ToString().Trim() + ")</font>");
        emailStr.Append(@"</td></tr>");
        emailStr.Append(@"<tr><td align=""middle"" width=""7%"" height=""17"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><font size=""2"" color=""#FFFFFF"">Trans.</font></td>");
        emailStr.Append(@"<td align=""middle"" width=""7%"" height=""17"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><font size=""2"" color=""#FFFFFF"">Posting</font></td>");
        emailStr.Append(@"<td align=""middle"" width=""61%"" colSpan=""2"" height=""17"" bgcolor=""#DCD5CF""><p align=""left""><font size = ""2"" color=""#404141"" face=""cordale arabic"">Previous Balance</font></p></td>");
        emailStr.Append(@"<td align=""middle"" width=""16%"" height=""18"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'><p align=""right"">");
        emailStr.Append(MakeHeaderStr(basText.alignmentRight(basText.formatNumUnSign(masterRow[mOpeningbalance], "#,###,##0.00", 12) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance].ToString())), 15), false, false));
        emailStr.Append(@"</p></td></tr>");

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

        if (trnsDesc == "Charge interest for Installment")// || trnsDesc == "Charge interest for Acceleration")
            return;

        emailStr.Append(@"<tr><td align=""middle"" width=""7%"" height=""21"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
        emailStr.Append(MakeHeaderStr(trnsDate.ToString("dd/MM"), false, false));
        emailStr.Append(@"</td><td align=""middle"" width=""7%"" height=""21"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
        emailStr.Append(MakeHeaderStr(postingDate.ToString("dd/MM"), false, false));
        emailStr.Append(@"</td><td width=""61%"" height=""21"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'><p align=""left"">");
        if (trnsDesc == "Installment repayment")
        {
            if (detailRow[dInstallmentData].ToString().Trim() != "")
            {
                string input = detailRow[dInstallmentData].ToString().Trim();
                int index = detailRow[dInstallmentData].ToString().Trim().IndexOf(":");
                if (index > 0)
                    input = input.Substring(0, index);
                emailStr.Append(MakeHeaderStr(input.Substring(0, input.IndexOf(',')) + input.Substring(basText.GetNthIndex(input, ',', 2)).Substring(input.Substring(basText.GetNthIndex(input, ',', 2)).IndexOf('#')).Replace('#', ' ') + " " + GetMerchantByOrigDocNo(long.Parse(detailRow[dOrigDocNo].ToString())), false, false));//basText.trimStr(trnsDesc.Trim(), 24)
            }
            else
                emailStr.Append(MakeHeaderStr(trnsDesc.Trim(), false, false));//basText.trimStr(trnsDesc.Trim(), 24)
        }
        else
            emailStr.Append(MakeHeaderStr(trnsDesc.Trim(), false, false));//basText.trimStr(trnsDesc.Trim(), 24)
        emailStr.Append(@"</p></td>");
        //emailStr.Append(@"&nbsp;" + ValidateStr(strForeignCurr.Trim()));
        //emailStr.Append(@"</p></td>");
        emailStr.Append(@"<td width=""61%"" height=""21"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'><p align=""left"">");
        emailStr.Append(ValidateStr(strForeignCurr.Trim()));
        emailStr.Append(@"</td>");
        emailStr.Append(@"<td align=""right"" width=""16%"" height=""21"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
        emailStr.Append(MakeHeaderStr((basText.formatNumUnSign(detailRow[dBilltranamount], "#,###,##0.00", 14)), false, false)); // + " " + CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])).ToString()
        //emailStr.Append(@"</td><td align=""left"" width=""4%"" height=""21"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
        //emailStr.Append(@"</td>");
        //emailStr.Append(isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString())).Substring(1,1);// ValidateStr(isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString())));
        emailStr.Append(@"<font color=""#4F5858"" size=""2"">");
        //emailStr.Append(Convert.ToString((CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])).ToString() + " " + (curMainCard != detailRow[dCardno].ToString().Trim() ? "SUP" : ""))).Trim();// ValidateStr(isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString())));
        emailStr.Append(Convert.ToString((CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])).ToString() + " " + isSuplStr)).Trim());
        emailStr.Append(@"</font>");
        //emailStr.Append("SUP");
        emailStr.Append(@"</td></tr>");

        //-streamWrit.WriteLine("  {0:dd/MM}  {1:dd/MM}  {2,-24}  {3,-40} {4,16} {5,16} {6,2}", trnsDate, postingDate, basText.trimStr(detailRow[dRefereneno].ToString().Trim(), 24), basText.trimStr(trnsDesc, 40), strForeignCurr, basText.formatNumUnSign(detailRow[dBilltranamount], "#,###,##0.00", 14), CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + " " + isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)
        totNoOfTransactions++;
    }

    protected void cardSeparator()
    {

        emailStr.Append(@"<tr><td align=""middle"" width=""7%"" height=""21"" style='font-family:cordale arabic'>&nbsp;");
        emailStr.Append(@"</td><td align=""middle"" width=""7%"" height=""21"" style='font-family:cordale arabic'>&nbsp;");
        emailStr.Append(@"</td><td width=""54%"" height=""21"" style='font-family:cordale arabic'><p align=""left""><b><font color=""#4F5858"" size=""2"">");
        emailStr.Append("Supplementary Card : " + masterRow[mCardtitle].ToString().Trim());
        emailStr.Append(@"</font></b></p></td><td align=""right"" width=""7%"" height=""21"" style='font-family:cordale arabic'><font size=""2"">");
        emailStr.Append(@"&nbsp;</font></td><td align=""right"" width=""16%"" height=""21"" style='font-family:cordale arabic'>&nbsp;");
        //emailStr.Append(@"</td><td align=""left"" width=""4%"" height=""21"" style='font-family:cordale arabic'>");
        emailStr.Append(@"</td>");
        emailStr.Append(@"<font color=""#4F5858"" size=""2"">&nbsp;");
        emailStr.Append(@"</font>");
        emailStr.Append(@"</td></tr>");

        totNoOfTransactions++;
    }

    protected void printCardFooter()
    {
        emailStr.Append(@"<tr><td align=""middle"" width=""7%"" height=""21"" style='font-family:cordale arabic'>");
        emailStr.Append(@"&nbsp;</td><td align=""middle"" width=""7%"" height=""21"" style='font-family:cordale arabic'>");
        emailStr.Append(@"&nbsp;</td><td width=""54%"" height=""21"" style='font-family:cordale arabic'><p align=""left"">");
        emailStr.Append(@"&nbsp;</p></td><td align=""right"" width=""7%"" height=""21"" style='font-family:cordale arabic'>");
        emailStr.Append(@"&nbsp;</td><td align=""right"" width=""16%"" height=""21"" style='font-family:cordale arabic'>");
        emailStr.Append(@"&nbsp;</td><td align=""left"" width=""4%"" height=""21"" style='font-family:cordale arabic'>");
        emailStr.Append(@"&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td align=""middle"" width=""7%"" height=""21"" style='font-family:cordale arabic'>");
        emailStr.Append(@"&nbsp;</td><td align=""middle"" width=""7%"" height=""21"" style='font-family:cordale arabic'>");
        emailStr.Append(@"&nbsp;</td><td width=""54%"" height=""21"" style='font-family:cordale arabic'><p align=""left"">");
        emailStr.Append(@"<font size = ""2"" color=""#404141"" face=""cordale arabic"">Current Balance</font>");
        emailStr.Append(@"</p></td><td align=""right"" width=""7%"" height=""21"" style='font-family:cordale arabic'>");
        emailStr.Append(@"&nbsp;</td><td align=""right"" width=""16%"" height=""21"" style='font-family:cordale arabic'>");
        emailStr.Append(MakeHeaderStr(basText.alignmentRight(basText.formatNumUnSign(masterRow[mClosingbalance], "#,###,##0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])), 15), true, false));
        emailStr.Append(@"</td><td align=""left"" width=""4%"" height=""21"" style='font-family:cordale arabic'>");
        emailStr.Append(@"&nbsp;</td></tr>");
        if (isRewardVal)
        {
            //NSGB-3444
            //rewardRows = DSreward.Tables["Reward"].Select("CLIENTID = '" + clsBasValid.validateStr(masterRow[mClientid]).ToString().Trim() + "'");
            rewardRows = DSreward.Tables["Reward"].Select("CLIENTID = '" + clsBasValid.validateStr(masterRow[ClientID]).ToString().Trim() + "'");// and CREDITCONTRACTS = '" + clsBasValid.validateStr(masterRow[mContractno]).ToString().Trim() + "'");
            if (rewardRows.Length > 0)
            {
                rewardRow = rewardRows[0];
                emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"" height=""65"" style='font-family:cordale arabic'><table id=""table22"" height=""42"" width=""100%"" border=""0""><tbody>");
                emailStr.Append(@"<tr><td align=""middle"" width=""134"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><b><font size=""2"" color=""#FFFFFF"">Reward Opening Balance</font></b></td>");
                emailStr.Append(@"<td align=""middle"" width=""141"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><b><font size=""2"" color=""#FFFFFF"">Total Points Earned</font></b></td>");
                emailStr.Append(@"<td align=""middle"" width=""145"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><b><font size=""2"" color=""#FFFFFF"">Total Points Redeemed</b></font></td>");
                emailStr.Append(@"<td align=""middle"" width=""165"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><b><font size=""2"" color=""#FFFFFF"">Expired Points</font></b></td>");
                emailStr.Append(@"<td align=""middle"" width=""148"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><b><font size=""2"" color=""#FFFFFF"">Reward Closing Balance</font></b></td>");
                emailStr.Append(@"</tr><tr><td align=""middle"" width=""134"" height=""16"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
                emailStr.Append(MakeHeaderStr(basText.formatNum(rewardRow[mOpeningbalance], "#0.00;(#0.00)", 20, "M"), false, false));
                emailStr.Append(@"</td><td align=""middle"" width=""141"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
                emailStr.Append(MakeHeaderStr(basText.formatNum(rewardRow[mEarnedBonus], "#0.00;(#0.00)", 20, "M"), false, false));
                emailStr.Append(@"</td><td align=""middle"" width=""145"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
                emailStr.Append(MakeHeaderStr(basText.formatNum(rewardRow[mRedeemedBonus], "#0.00;(#0.00)", 20, "M"), false, false));
                emailStr.Append(@"</td><td align=""middle"" width=""165"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
                emailStr.Append(MakeHeaderStr(basText.formatNum(rewardRow[mExpiredBonus], "#0.00;(#0.00)", 20, "M"), false, false));
                emailStr.Append(@"</td><td align=""middle"" width=""148"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
                emailStr.Append(MakeHeaderStr(basText.formatNum(rewardRow[mClosingbalance], "#0.00;(#0.00)", 20, "M"), false, false));
                emailStr.Append(@"</td></tr></tbody></table></td></tr>");
            }
            else //NSGB-3444
            {
                emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"" height=""65"" style='font-family:cordale arabic'><table id=""table22"" height=""42"" width=""100%"" border=""0""><tbody>");
                emailStr.Append(@"<tr><td align=""middle"" width=""134"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><b><font size=""2"" color=""#FFFFFF"">Reward Opening Balance</font></b></td>");
                emailStr.Append(@"<td align=""middle"" width=""141"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><b><font size=""2"" color=""#FFFFFF"">Total Points Earned</font></b></td>");
                emailStr.Append(@"<td align=""middle"" width=""145"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><b><font size=""2"" color=""#FFFFFF"">Total Points Reedemed</b></font></td>");
                emailStr.Append(@"<td align=""middle"" width=""165"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><b><font size=""2"" color=""#FFFFFF"">Expired Points</font></b></td>");
                emailStr.Append(@"<td align=""middle"" width=""148"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><b><font size=""2"" color=""#FFFFFF"">Reward Closing Balance</font></b></td>");
                emailStr.Append(@"</tr><tr><td align=""middle"" width=""134"" height=""16"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
                emailStr.Append(MakeHeaderStr(basText.formatNum(0, "#0.00;(#0.00)", 20, "M"), false, false));
                emailStr.Append(@"</td><td align=""middle"" width=""141"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
                emailStr.Append(MakeHeaderStr(basText.formatNum(0, "#0.00;(#0.00)", 20, "M"), false, false));
                emailStr.Append(@"</td><td align=""middle"" width=""145"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
                emailStr.Append(MakeHeaderStr(basText.formatNum(0, "#0.00;(#0.00)", 20, "M"), false, false));
                emailStr.Append(@"</td><td align=""middle"" width=""165"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
                emailStr.Append(MakeHeaderStr(basText.formatNum(0, "#0.00;(#0.00)", 20, "M"), false, false));
                emailStr.Append(@"</td><td align=""middle"" width=""148"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
                emailStr.Append(MakeHeaderStr(basText.formatNum(0, "#0.00;(#0.00)", 20, "M"), false, false));
                emailStr.Append(@"</td></tr></tbody></table></td></tr>");
            }
        }
        //NSGB-3444
        emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"" height=""65"" style='font-family:cordale arabic'><table id=""table22"" height=""42"" width=""100%"" border=""0""><tbody>");
        emailStr.Append(@"<tr><td align=""middle"" width=""134"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><b><font size=""2"" color=""#FFFFFF"">Card Number</font></b></td>");
        emailStr.Append(@"<td align=""middle"" width=""141"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><b><font size=""2"" color=""#FFFFFF"">Current Balance</font></b></td>");
        emailStr.Append(@"<td align=""middle"" width=""145"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><b><font size=""2"" color=""#FFFFFF"">Minimum Payment</b></font></td>");
        emailStr.Append(@"<td align=""middle"" width=""148"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><b><font size=""2"" color=""#FFFFFF"">Payment Due Date</font></b></td>");
        emailStr.Append(@"</tr><tr><td align=""middle"" width=""134"" height=""16"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
        //if (strPrimaryCardNo.Trim() == "")
        //    emailStr.Append(MakeHeaderStr(strPrimaryCardNo, false, false));
        //else
        //    emailStr.Append(MakeHeaderStr(strPrimaryCardNo.Substring(0, 6) + "******" + strPrimaryCardNo.Substring(12, 4), false, false));
        //NSGB-3528
        if (curMainCard.Trim() == "")
            emailStr.Append(MakeHeaderStr(curMainCard, false, false));
        else
            emailStr.Append(MakeHeaderStr(curMainCard.Substring(0, 6) + "******" + curMainCard.Substring(12, 4), false, false));
        emailStr.Append(@"</td><td align=""middle"" width=""141"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
        emailStr.Append(MakeHeaderStr(basText.alignmentRight(basText.formatNumUnSign(masterRow[mClosingbalance], "#,###,##0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])), 15), true, false));
        emailStr.Append(@"</td><td align=""middle"" width=""145"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
        //emailStr.Append(MakeHeaderStr(basText.alignmentRight(basText.formatNumUnSign(masterRow[mMindueamount], "#,###,##0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mMindueamount])), 15), true, false));
        emailStr.Append(MakeHeaderStr(basText.alignmentRight(basText.formatNumUnSign(masterRow[mMindueamount], "#,###,##0.00", 16), 15), true, false));
        emailStr.Append(@"</td><td align=""middle"" width=""165"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
        emailStr.Append(MakeHeaderStr(basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M"), true, false));
        emailStr.Append(@"</td></tr></tbody></table></td></tr>");

        emailStr.Append(@"</tbody></table></td></tr><tr><td width=""91%"" colSpan=""6"" height=""20"" style='font-family:cordale arabic'>&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"" style='font-family:cordale arabic'>&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td width=""99%"" colSpan=""6"" height=""21"" style='font-family:cordale arabic'>" + statMessageMonthly + "</td></tr>");
        emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"" style='font-family:cordale arabic'>&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"" style='font-family:cordale arabic'>&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"" style='font-family:cordale arabic'>&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td width=""100%"" colSpan=""6"" height=""10"" bgcolor=""#FFFFFF"" style='font-family:cordale arabic'></td></tr>"); //black line
        emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"" style='font-family:cordale arabic'>&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"" style='font-family:cordale arabic'>&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td width=""99%"" colSpan=""6"" height=""21"" style='font-family:cordale arabic'>&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td width=""99%"" colSpan=""6"" height=""21"" style='font-family:cordale arabic'>" + statMessage + "</td></tr>");
        //emailStr.Append(@"<tr><td width=""99%"" colSpan=""6"" height=""21"" style='font-family:cordale arabic'><img border='0' src='cid:" + clsBasFile.getFileFromPath(strbottomBanner) + @"'></td></tr>");
        //emailStr.Append(@"<tr><td colSpan=""6""><img align='middle' height='150%' width='125%' src='cid:" + clsBasFile.getFileFromPath(strbottomBanner) + @"'/></td></tr>"); //advertise Photo
        emailStr.Append(@"</tbody></table>");
        emailStr.Append(@"<img style=""margin:0;height:100%;width:100%;display:block;"" src='cid:" + clsBasFile.getFileFromPath(strbottomBanner) + @"'/>"); //advertise Photo
        emailStr.Append(@"<table><tbody>");
        emailStr.Append(@"<tr><td width=""99%"" colSpan=""6"" height=""21"" style='font-family:cordale arabic'></td></tr>");
        emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"" style='font-family:cordale arabic'>&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"" style='font-family:cordale arabic'><p align=""right"">");
        //if (isRewardVal)
        //    {
        //    rewardRows = DSreward.Tables["Reward"].Select("CLIENTID = '" + clsBasValid.validateStr(masterRow[mClientid]).ToString().Trim() + "'");
        //    if (rewardRows.Length > 0)
        //        {
        //        rewardRow = rewardRows[0];
        //        emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"" height=""65"" style='font-family:cordale arabic'><table id=""table22"" height=""42"" width=""100%"" border=""0""><tbody>");
        //        emailStr.Append(@"<tr><td align=""middle"" width=""134"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><b><font size=""2"" color=""#FFFFFF"">Reward Opening Balance</font></b></td>");
        //        emailStr.Append(@"<td align=""middle"" width=""141"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><b><font size=""2"" color=""#FFFFFF"">Total Points Earned</font></b></td>");
        //        emailStr.Append(@"<td align=""middle"" width=""145"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><b><font size=""2"" color=""#FFFFFF"">Total Points Reedemed</b></font></td>");
        //        emailStr.Append(@"<td align=""middle"" width=""165"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><b><font size=""2"" color=""#FFFFFF"">Expired Points</font></b></td>");
        //        emailStr.Append(@"<td align=""middle"" width=""148"" bgcolor=""#4F5858"" style=""border-style: outset; border-width: 1px; font-family:cordale arabic'""><b><font size=""2"" color=""#FFFFFF"">Reward Closing Balance</font></b></td>");
        //        emailStr.Append(@"</tr><tr><td align=""middle"" width=""134"" height=""16"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
        //        emailStr.Append(MakeHeaderStr(basText.formatNum(rewardRow[mOpeningbalance], "#0.00;(#0.00)", 20, "M"), false, false));
        //        emailStr.Append(@"</td><td align=""middle"" width=""141"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
        //        emailStr.Append(MakeHeaderStr(basText.formatNum(rewardRow[mEarnedBonus], "#0.00;(#0.00)", 20, "M"), false, false));
        //        emailStr.Append(@"</td><td align=""middle"" width=""145"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
        //        emailStr.Append(MakeHeaderStr(basText.formatNum(rewardRow[mRedeemedBonus], "#0.00;(#0.00)", 20, "M"), false, false));
        //        emailStr.Append(@"</td><td align=""middle"" width=""165"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
        //        emailStr.Append(MakeHeaderStr(basText.formatNum(rewardRow[mExpiredBonus], "#0.00;(#0.00)", 20, "M"), false, false));
        //        emailStr.Append(@"</td><td align=""middle"" width=""148"" bgcolor=""#DCD5CF"" style='font-family:cordale arabic'>");
        //        emailStr.Append(MakeHeaderStr(basText.formatNum(rewardRow[mClosingbalance], "#0.00;(#0.00)", 20, "M"), false, false));
        //        emailStr.Append(@"</td></tr></tbody></table></td></tr>");
        //        }
        //    }
        emailStr.Append(@"<a title="" +strbankWebLink+"" href=""http://" + strbankWebLinkService + @""">" + strbankWebLinkService + @"</a>");
        emailStr.Append(@"</p></td></tr>");
        emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"" style='font-family:cordale arabic'><p align=""right"">");
        emailStr.Append(@"<a title="" +strbankWebLink+"" href=""http://" + strbankWebLink + @""">" + strbankWebLink + @"</a>");
        emailStr.Append(@"</p></td></tr>");
        emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"" style='font-family:cordale arabic'>&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td width='98%' colSpan='6' style='border-style: outset; border-width: 1px; font-family:cordale arabic' bordercolor='#FFFFFF'>");
        emailStr.Append(@"This message (including any attachments) is confidential and may be privileged or intended solely for the addressee. If you are not the addressee, or have received this message by mistake, please notify the sender immediately by return e-mail, and delete this message from your system. Do not copy, disclose or otherwise act upon any part of this e-mail or its attachments. Any unauthorized use or dissemination of this message in whole or in part is strictly prohibited. Please note that e-mails are susceptible to change. <br><br>");
        emailStr.Append(@"QNB ALAHLI shall not be liable for the improper or incomplete transmission of the information contained in this communication nor for any delay in its receipt or damage to your system.<br><br>");
        emailStr.Append(@"QNB ALAHLI does not guarantee that the integrity of this communication has been maintained nor that this communication is free of viruses, interceptions or interference.<br><br>");
        //emailStr.Append(@"***For any inquiries, please call bebasata customer solution center 16607***.</td>");
        emailStr.Append(@"</tr><tr><td width=""98%"" colSpan=""6"" style='font-family:cordale arabic'>&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"" style='font-family:cordale arabic'>&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td width='98%' colSpan='6'>&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td width='98%' colSpan='6'>&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td width='98%' colSpan='6'>&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td width='98%' colSpan='6'>" + clsCheckSum.stringMD5(extAccNum.Trim() + Convert.ToDateTime(masterRow[mStatementdateto]).ToString("dd/MM/yyyy")) + "</td></tr>");
        emailStr.Append(@"<tr><td width='98%' colSpan='6'>&nbsp;</td></tr>");
        emailStr.Append(@"</tbody></table></body></html>");
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
        if (createCorporateVal)
            mainRows = DSstatement.Tables["tStatementMasterTable"].Select("CARDACCOUNTNO = '" + clsBasValid.validateStr(masterRow[accountNoName]) + "'");
        else
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
            pStr = "&nbsp;";
        else
        {
            //pStr = @"<font size=""2"" color=""" + color + @""">" + pStr + @"</font>";
            pStr = @"<font size=""2"">" + pStr + @"</font>";
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

    //private void SendEmailWithDifferentSender(string pBody, string pSubject, string pTo)
    //    {
    //    if (!basText.isValideEmail(pTo))
    //        {
    //        streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|"+clsCnfg.readSetting("strValidEmail")?? "Bad Email");
    //        noOfBadEmails++;
    //        return;
    //        }

    //    try
    //        {
    //        //string pFrom, pSubject, pBody;
    //        //lblStatus.Text = string.Empty;
    //        //pFrom = "mmohammed@emp-group.com";//mmohammed@mscc.local"mmohammed@emp-group.com"mmohammed@emp-group.com
    //        //return;

    //        //streamEmails.WriteLine(curAccountNumber + "|" + curCardNumber + "|" + curClientID + "|" + emailTo);
    //        streamEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + strMobileNum.Trim() + "|" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
    //        //return;
    //        //string[] strAray;
    //        //if (emailTo != "lucky.ighade@zenithbank.com")//ehimolen@yahoo.com
    //        //return;
    //        ArrayList pLstTo = new ArrayList(), pLstAttachedPic = new ArrayList(), pLstAttachedFile = new ArrayList();//, pLstCC = new ArrayList(), pLstBCC = new ArrayList();
    //        pSubject = emailLabel; // "BAI statement for 02/2008";
    //        //pTo = emailTo;// "mmohammed@emp-group.com";
    //        if (pTo.EndsWith("."))//ehimolen@yahoo.com
    //            pTo = pTo + "com";
    //        pLstTo.Add(pTo);//"mmohammed@emp-group.com" "Ahmed.Ali@socgen.com" "mmohammed@emp-group.com""mhrap@yahoo.com""mhrap@hotmail.com"  "dossantf@emp-group.com" "mhrap@hotmail.com" "Tmahfouz61@yahoo.com"  "nazab@emp-group.com" "developers@emp-group.com""nazab@emp-group.com"  "wbaioumy@emp-group.com""mmohammed@emp-group.com" "mmohammed@emp-group.com""nazab@emp-group.com"
    //        //return;0
    //        //pLstTo.Add("amr244@yahoo.com"); //Amr Mohsen
    //        //pLstTo.Add("Amr.Mohsen@socgen.com");
    //        //pLstTo.Add("mmohammed@emp-group.com");
    //        //pLstBCC.Add("amr244@yahoo.com"); //Amr Mohsen
    //        //pLstBCC.Add("mmohammed@emp-group.com");
    //        //pLstCC.Add("Nermine.Bahaa-Eldin@socgen.com");
    //        //pLstBCC.Add("mervatlewis@yahoo.com"); //test email
    //        //pLstBCC.Add("nazab@emp-group.com");
    //        //pLstBCC.Add("mhrap@yahoo.com");
    //        //pLstBCC.Add("amostafa@emp-group.com");

    //        pLstAttachedPic.Add(strBankLogo);//strOutputFile@"D:\Web\Email\BAI\Logo.gif"
    //        pLstAttachedPic.Add(strbackGround);//@"D:\pC#\exe\FilesForPrograms\frmBackground.jpg"
    //        pLstAttachedPic.Add(strVisaLogo);//@"D:\pC#\exe\FilesForPrograms\frmBackground.jpg"

    //        //if (pBody.Contains(@"<tr><td><b><font size=""2"">Card Type:</font></b></td><td height=""27""><font size=""2"">Visa Platinum</font></td></tr>"))
    //        //    pLstAttachedPic.Add(strbottomBannerPlatinum);
    //        //else
    //        //    pLstAttachedPic.Add(strbottomBannerDefault);

    //        pLstAttachedPic.Add(strbottomBanner);

    //        //if (System.IO.File.Exists(@"D:\pC#\ProjData\Statement\NSGB\NSGB_Message.pdf"))
    //        //{
    //        //    pLstAttachedFile.Add(@"D:\pC#\ProjData\Statement\NSGB\NSGB_Message.pdf");
    //        //}
    //        if (HasAttachement)
    //            foreach (string fileName in Directory.GetFiles(attachedFilesStr))
    //                {
    //                pLstAttachedFile.Add(fileName);
    //                }


    //        sndMail = new clsEmail();
    //        sndMail.emailFromName = emailFromNameStr;
    //        //if (!sndMail.sendEmailHTML(strEmailFrom, pLstTo, pLstCC, pLstBCC, pLstAttachedFile, pSubject, pBody, clsCnfg.readSetting("SmtpServer")))
    //        //if (pTo.ToLower().IndexOf("qnbalahli.com") > 0)
    //        //    strEmailFromTmp = "QNB ALAHLI" + noOfEmails.ToString("FFFFFF") + "@emp-group.com";//"nsgb000001@emp-group.com""nsgb@emp-group.com""mmohammed@emp-group.com";//
    //        //else
    //        strEmailFromTmp = strEmailFrom;

    //        numOfTry = 0;
    //        //while (!sndMail.sendEmailHTML(strEmailFromTmp, pLstTo, pLstCC, pLstBCC, pLstAttachedFile, pSubject, pBody, clsCnfg.readSetting("SmtpServer")))
    //        noOfEmails++;
    //        //while (!sndMail.sendEmailAttachPic(strEmailFromTmp, pLstTo, pLstCC, pLstBCC, pLstAttachedFile, pSubject, pBody, clsCnfg.readSetting("SmtpServer")))
    //        if (numMail == 0)
    //            {
    //            //pLstBCC.Add("mmohammed@emp-group.com");
    //            pLstBCC.Add("statement@emp-group.com");
    //            }
    //        while (!sndMail.sendEmailAttachPicFilewithDiffernetSender(strEmailFromTmp, "statement@emp-group.com", pLstTo, pLstCC, pLstBCC, pLstAttachedPic, pLstAttachedFile, pSubject, pBody, clsCnfg.readSetting("SmtpServer")))
    //            {
    //            //MessageBox.Show("Failure to send Email.", "Send Email Error",MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
    //            System.Threading.Thread.Sleep(waitPeriodVal);//2000
    //            numOfTry++;
    //            if (numOfTry > 100)
    //            //if (numOfTry > 1)
    //                {
    //                streamEmails.Write("\t\t Error while Send Email");
    //                streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|Exceed number of trials");
    //                noOfBadEmails++; noOfEmails--;
    //                break;
    //                }
    //            }
    //        if (numMail == 0)
    //            pLstBCC.Clear();

    //        numMail++;
    //        if (numMail % 400 == 0)
    //            {
    //            System.Threading.Thread.Sleep(waitPeriodVal);//2000
    //            GC.Collect();
    //            GC.WaitForPendingFinalizers();
    //            }
    //        sndMail = null;
    //        emailStr = new StringBuilder(""); emailTo = strMobileNum = string.Empty; curAccountNumber = string.Empty; curCardNumber = string.Empty; curClientID = string.Empty;
    //        wait2NextEmail(prvEmail, pTo, waitPeriodVal);//     System.Threading.Thread.Sleep(waitPeriodVal);//400
    //        prvEmail = pTo;
    //        }
    //    catch  //(NotSupportedException ex) (Exception ex)  //
    //        {
    //        //clsBasErrors.catchError(ex);
    //        streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|"+clsCnfg.readSetting("strValidEmail")?? "Bad Email");
    //        noOfBadEmails++; noOfEmails--;
    //        }
    //    finally
    //        {
    //        }
    //    }

    private void SendEmail(string pBody, string pSubject, string pTo)
    {
        try
        {

            if (!basText.isValideEmail(pTo))
            {
                streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + pTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|"+clsCnfg.readSetting("strValidEmail")?? "Bad Email " + pTo);
                noOfBadEmails++;
                return;
            }

            //try
            //    {
            //string pFrom, pSubject, pBody;
            //lblStatus.Text = string.Empty;
            //pFrom = "mmohammed@emp-group.com";//mmohammed@mscc.local"mmohammed@emp-group.com"mmohammed@emp-group.com
            //return;

            //streamEmails.WriteLine(curAccountNumber + "|" + curCardNumber + "|" + curClientID + "|" + emailTo);
            //return;
            //string[] strAray;
            //if (emailTo != "lucky.ighade@zenithbank.com")//ehimolen@yahoo.com
            //return;
            ArrayList pLstTo = new ArrayList(), pLstAttachedPic = new ArrayList(), pLstAttachedFile = new ArrayList();//, pLstCC = new ArrayList(), pLstBCC = new ArrayList();
            pSubject = emailLabel; // "BAI statement for 02/2008";
            //pTo = emailTo;// "mmohammed@emp-group.com";
            if (pTo.EndsWith("."))//ehimolen@yahoo.com
                pTo = pTo + "com";
            pLstTo.Add(pTo);//"mmohammed@emp-group.com" "Ahmed.Ali@socgen.com" "mmohammed@emp-group.com""mhrap@yahoo.com""mhrap@hotmail.com"  "dossantf@emp-group.com" "mhrap@hotmail.com" "Tmahfouz61@yahoo.com"  "nazab@emp-group.com" "developers@emp-group.com""nazab@emp-group.com"  "wbaioumy@emp-group.com""mmohammed@emp-group.com" "mmohammed@emp-group.com""nazab@emp-group.com"
            //return;0
            //pLstTo.Add("amr244@yahoo.com"); //Amr Mohsen
            //pLstTo.Add("Amr.Mohsen@socgen.com");
            //pLstTo.Add("mmohammed@emp-group.com");
            //pLstBCC.Add("amr244@yahoo.com"); //Amr Mohsen
            //pLstBCC.Add("mmohammed@emp-group.com");
            //pLstCC.Add("Nermine.Bahaa-Eldin@socgen.com");
            //pLstBCC.Add("mervatlewis@yahoo.com"); //test email
            //pLstBCC.Add("nazab@emp-group.com");
            //pLstBCC.Add("mhrap@yahoo.com");
            //pLstBCC.Add("amostafa@emp-group.com");

            pLstAttachedPic.Add(strBankLogo);//strOutputFile@"D:\Web\Email\BAI\Logo.gif"
            pLstAttachedPic.Add(strbackGround);//@"D:\pC#\exe\FilesForPrograms\frmBackground.jpg"
            pLstAttachedPic.Add(strVisaLogo);//@"D:\pC#\exe\FilesForPrograms\frmBackground.jpg"

            //if (pBody.Contains(@"<tr><td><b><font size=""2"">Card Type:</font></b></td><td height=""27""><font size=""2"">Visa Platinum</font></td></tr>"))
            //    pLstAttachedPic.Add(strbottomBannerPlatinum);
            //else
            //    pLstAttachedPic.Add(strbottomBannerDefault);

            pLstAttachedPic.Add(strbottomBanner);

            //if (System.IO.File.Exists(@"D:\pC#\ProjData\Statement\NSGB\NSGB_Message.pdf"))
            //{
            //    pLstAttachedFile.Add(@"D:\pC#\ProjData\Statement\NSGB\NSGB_Message.pdf");
            //}
            if (HasAttachement)
                foreach (string fileName in Directory.GetFiles(attachedFilesStr))
                {
                    pLstAttachedFile.Add(fileName);
                }


            sndMail = new clsEmail();
            sndMail.emailFromName = emailFromNameStr;
            //if (!sndMail.sendEmailHTML(strEmailFrom, pLstTo, pLstCC, pLstBCC, pLstAttachedFile, pSubject, pBody, clsCnfg.readSetting("SmtpServer")))
            //if (pTo.ToLower().IndexOf("qnbalahli.com") > 0)
            //    strEmailFromTmp = "QNB ALAHLI" + noOfEmails.ToString("FFFFFF") + "@emp-group.com";//"nsgb000001@emp-group.com""nsgb@emp-group.com""mmohammed@emp-group.com";//
            //else
            strEmailFromTmp = strEmailFrom;

            numOfTry = 0;
            //while (!sndMail.sendEmailHTML(strEmailFromTmp, pLstTo, pLstCC, pLstBCC, pLstAttachedFile, pSubject, pBody, clsCnfg.readSetting("SmtpServer")))
            noOfEmails++;
            //while (!sndMail.sendEmailAttachPic(strEmailFromTmp, pLstTo, pLstCC, pLstBCC, pLstAttachedFile, pSubject, pBody, clsCnfg.readSetting("SmtpServer")))
            if (numMail == 0)
            {
                //pLstBCC.Add("mmohammed@emp-group.com");
                pLstBCC.Add("statement@emp-group.com");
            }
            while (!sndMail.sendEmailAttachPicFile(strEmailFromTmp, pLstTo, pLstCC, pLstBCC, pLstAttachedPic, pLstAttachedFile, pSubject, pBody, clsCnfg.readSetting("SmtpServer")))
            {
                //MessageBox.Show("Failure to send Email.", "Send Email Error",MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                System.Threading.Thread.Sleep(waitPeriodVal);//2000
                numOfTry++;
                if (numOfTry > 100)
                //if (numOfTry > 1)
                {
                    streamEmails.Write("\t\t Error while Send Email");
                    streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|Exceed number of trials");
                    noOfBadEmails++; noOfEmails--;
                    break;
                }
            }
            streamEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + pTo + "|" + strMobileNum.Trim() + "|" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));

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
            //emailStr = new StringBuilder(""); emailTo = strMobileNum = string.Empty; curAccountNumber = string.Empty; curCardNumber = string.Empty; curClientID = string.Empty;
            //wait2NextEmail(prvEmail, pTo, waitPeriodVal);//     System.Threading.Thread.Sleep(waitPeriodVal);//400
            //prvEmail = pTo;
        }
        catch (Exception ex) //(NotSupportedException ex) (Exception ex)  //
        {
            //clsBasErrors.catchError(ex);
            streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + pTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "| Email " + pTo + " , Err Message >> " + ex.Message);
            noOfBadEmails++; noOfEmails--;
        }
        finally
        {
            sndMail = null;
            emailStr = new StringBuilder(""); emailTo = strMobileNum = string.Empty; curAccountNumber = string.Empty; curCardNumber = string.Empty; curClientID = string.Empty;
            wait2NextEmail(prvEmail, pTo, waitPeriodVal);//     System.Threading.Thread.Sleep(waitPeriodVal);//400
            prvEmail = pTo;
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
        //    streamSummary.WriteLine("No of Transactions " + totNoOfTransactions.ToString());

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

    private string GetMerchantByOrigDocNo(long pDocno)
    {
        mainRows = null;
        string result = string.Empty;
        //mainRows = DSstatement.Tables["tStatementDetailTable"].Select("DOCNO = " + pDocno + "");
        mainRows = DSstatement.Tables["tStatementDetailTable"].Select("ORIGDOCNO = " + pDocno + "");
        foreach (DataRow mainRow in mainRows)
        {
            //result = mainRow[dMerchant].ToString();
            //result = mainRow[dInstallmentMerchant].ToString();
            result = mainRow[dInstallmentMerchantLocation].ToString();
        }
        return result;
    }

    // Default 7000
    public int waitPeriod
    {
        get { return waitPeriodVal; }
        set { waitPeriodVal = value; }
    }//waitPeriod

    public string statMessageFile
    {
        get { return statMessageFileVal; }
        set { statMessageFileVal = value; }
    }//statMessageFile

    public string statMessageBox
    {
        get { return statMessageBoxVal; }
        set { statMessageBoxVal = value; }
    }//statMessageBox

    public string statMessageFileMonthly
    {
        get { return statMessageFileMonthlyVal; }
        set { statMessageFileMonthlyVal = value; }
    }//statMessageFileMonthly

    public string visaLogo
    {
        get { return strVisaLogo; }
        set { strVisaLogo = value; }
    }//visaLogo

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

    public string bankWebLinkService
    {
        get { return strbankWebLinkService; }
        set { strbankWebLinkService = value; }
    }// bankWebLinkService

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

    public string bottomBanner
    {
        get { return strbottomBanner; }
        set { strbottomBanner = value; }
    }// bottomBanner

    public string bottomBannerDefault
    {
        get { return strbottomBannerDefault; }
        set { strbottomBannerDefault = value; }
    }// bottomBannerDefault

    public string bottomBannerPlatinum
    {
        get { return strbottomBannerPlatinum; }
        set { strbottomBannerPlatinum = value; }
    }// bottomBannerPlatinum

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

    public string attachedFiles
    {
        get { return attachedFilesStr; }
        set { attachedFilesStr = value; }
    }// attachedFiles

    public string productCond
    {
        get { return strProductCond; }
        set { strProductCond = value; }
    }  // productCond

    public bool IsSplitted
    {
        get { return IsSplittedVal; }
        set { IsSplittedVal = value; }
    }  // IsSplitted

    public bool HasSender
    {
        get { return HasSenderVal; }
        set { HasSenderVal = value; }
    }  // HasSender

    public bool HasAttachement
    {
        get { return HasAttachementVal; }
        set { HasAttachementVal = value; }
    }  // HasAttachement

    ~clsStatementNSGBcreditHtml()
    {
        DSstatement.Dispose();
    }
}
