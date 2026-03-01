using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Xml;
using System.Collections;


public class clsStatHtmlCorp : clsStatement_CommonCorpCmpny
{
    protected string strBankName;
    protected FileStream fileSummary, fileEmails, fileNoEmails;
    protected StreamWriter streamSummary, streamEmails, streamNoEmails;
    protected DataRow masterRow;
    protected DataRow detailRow;
    protected const int MaxDetailInPage = 20; //
    protected const int MaxDetailInLastPage = 27; //
    protected int CurPageRec4Dtl = 0;
    protected int pageNo = 0, totalPages = 0, totalCardPages = 0
        , totalAccPages = 0, totCardRows = 0, totAccRows = 0, curAccRows = 0;//
    protected string lastPageTotal;
    protected string curCardNo;//,PrevCardNo
    protected string curAccountNo, prevAccountNo = String.Empty;//,PrevCardNo
    protected decimal totNetUsage = 0;
    protected DataRow[] cardsRows, accountRows, rewardRows;
    protected DataRow[] mainRows;
    protected DataRow rewardRow;
    protected string CurrentPageFlag;
    protected string strCardNo, strPrimaryCardNo;
    protected string strForeignCurr;
    protected string stmNo;
    protected int totNoOfCardStat, totNoOfPageStat;

    protected int curCardRow = 0, curTotCardRows = 0, numOfErr = 0, totNoOfTransactions = 0;
    protected bool isPrimaryOnly, isHaveF3 = true, isPrimaryOpened = false;
    protected FileStream fileStrmErr;
    protected StreamWriter strmWriteErr;
    protected string curMainCard;

    protected string extAccNum;
    protected int totCrdNoInAcc, curCrdNoInAcc;
    protected string strOutputPath, strOutputFile, fileSummaryName;
    protected DateTime vCurDate;
    protected int totPages;
    protected string endOfCustomer = string.Empty;
    protected string cProduct = string.Empty, curFileName = string.Empty;
    protected ArrayList aryLstFiles;

    //
    //protected string emailStr = string.Empty;
    protected StringBuilder emailStr = new StringBuilder("");
    protected string emailLabel, strBankLogo, strVisaLogo = @"D:\pC#\ProjData\Statement\_Background\VisaLogo.gif";
    protected string strEmailFrom = "mabouleila@emp-group.com", strEmailFromTmp = string.Empty;  //"cardservices@zenithbank.com"
    protected string strbankWebLink = "www.emp-group.com";  //www.zenithbank.com
    protected string strbankWebLinkService = "www.emp-group.com";  //www.zenithbank.com
    protected string emailTo = string.Empty, curAccountNumber, curCardNumber, curClientID;//, curEmail
    protected string strEmailFileName, strNoEmailFileName;
    protected frmStatementFile frmMain;
    protected int totRec = 1, numMail, numOfTry = 0;
    protected ArrayList pLstCC = new ArrayList(), pLstBCC = new ArrayList();
    protected string logoAlignmentStr = "center";

    protected string strbackGround = @"D:\pC#\ProjData\Statement\_Background\Background06.jpg";
    protected string strWaterMark = @"D:\pC#\ProjData\Statement\FBN\watermark.jpg";
    protected bool isRewardVal = false;
    protected bool createCorporateVal = false;
    protected string accountNoName = mAccountno;
    protected string accountLimit = mAccountlim;
    protected string accountAvailableLimit = mAccountavailablelim;
    protected string rewardCondVal = "'New Reward Contract'";//'Reward Contract'
    protected string statMessageFileVal = string.Empty;
    protected string statMessage = "&nbsp;";//Null
    protected string statMessageFileMonthlyVal = string.Empty;
    protected string statMessageMonthly = "&nbsp;";//Null
    protected string isSuplStr = string.Empty;//Null
    protected DataRow[] emailRows = null;
    protected DataRow masterRelatedRow;
    //protected clsEmail sndMail;
    protected string emailFromNameStr;
    protected string Address1Name = mCustomeraddress1, Address2Name = mCustomeraddress2, Address3Name = mCustomeraddress3;
    protected string strMobileNum = string.Empty;
    protected int waitPeriodVal = 7000;
    protected clsAes cryptAes = new clsAes();
    protected int noOfEmails, noOfBadEmails, noOfWithoutEmails;
    protected string emailLabelTmp;  //"cardservices@zenithbank.com"
    protected DataRow emailRow;
    protected ArrayList pLstTo = new ArrayList(), pLstAttachedPic = new ArrayList(), pLstAttachedFile = new ArrayList();
    protected string fontTypeSize4Header = "<font color='#808080' size='2'>";
    protected string fontTypeSize = "<font color='#808080' size='2'>";
    protected string lineSeparator = "<tr><td width='98%' height='25'>&nbsp;</td></tr>";
    protected bool isUseAlterEmailVal = false, isSentBadEmailVal = true;
    protected string AlterEmailVal, AlterEmailCondVal;
    protected string emailSentResault = string.Empty;
    protected string prvEmail = string.Empty;
    public clsStatHtmlCorp()
    {
    }

    public virtual string Statement(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pCurDate, string pStmntType, bool pAppendData)
    {
        string rtrnStr = "Successfully Generate " + pBankName;
        int curMonth = pCurDate.Month;
        bool preExit = true;
        curMonth = pCurDate.Month;
        aryLstFiles = new ArrayList();
        strEmailFromTmp = strEmailFrom;
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
                Address1Name = mCardaddress1;
                Address2Name = mCardaddress2;
                Address3Name = mCardaddress3;
            }

            if (isRewardVal)
            {
                //maintainData.curRewardCond = rewardCond;
                //maintainData.fixReward(pBankCode, rewardCond);
                strMainTableCond = "m.contracttype != " + rewardCondVal;
                strSubTableCond = "d.trandescription != 'Calculated Points'";
                getReward(pBankCode);
            }

            reSetupTheValues();
            FillStatementDataSet(pBankCode); //DSstatement =  //10); //3
            getClientEmail(pBankCode);
            pageNo = 1; totalCardPages = 0;
            curCardNo = String.Empty;
            curAccountNo = String.Empty;
            frmMain.BeginInvoke(frmMain.setMinMaxProgressDelegate, new object[] { DSstatement.Tables["tStatementMasterTable"].Rows.Count });
            foreach (DataRow mRow in DSstatement.Tables["tStatementMasterTable"].Rows)
            {
                frmMain.BeginInvoke(frmMain.setProgressDelegate, new object[] { totRec++ });
                masterRow = mRow;
                //-cardsRows = mRow.GetChildRows(StatementNoDRel); //, DataRowVersion.Proposed
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
                    emailLabelTmp = emailLabel;
                    if (!string.IsNullOrEmpty(emailStr.ToString()) && !string.IsNullOrEmpty(emailTo)) //emailStr != string.Empty && emailStr != null  emailTo != string.Empty && emailTo != null
                        SendEmail(emailStr.ToString(), "", emailTo);
                    if (!string.IsNullOrEmpty(emailStr.ToString()) && string.IsNullOrEmpty(emailTo)) //emailStr != string.Empty && emailStr != null  emailTo != string.Empty && emailTo != null
                    {
                        streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + emailLabelTmp + "|Without Email");
                        noOfWithoutEmails++;

                        emailLabelTmp = string.Empty;
                        if (emailRow != null)
                            emailLabelTmp = (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!");
                        emailTo = strEmailFrom; //"statement_Program@emp-group.com";

                        emailLabelTmp = emailLabel + " Acc:" + curAccountNumber + " Phone:" + emailLabelTmp;
                        SendEmail(emailStr.ToString(), "", emailTo);

                    }

                    emailStr = new StringBuilder("");
                    //cardsRows = DSstatement.Tables["tStatementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[accountNoName]).ToString().Trim() + "'");
                    cardsRows = DSstatement.Tables["tStatementMasterTable"].Select("PrinaryCardNo = '" + strPrimaryCardNo +"'");

                    curMainCard = string.Empty;
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
                    emailRows = DSemails.Tables["Emails"].Select("idclient = " + masterRow[mClientid].ToString());
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
                    //>>>emailTo = masterRow[mDept].ToString().Trim();//>>

                    curAccountNumber = masterRow[accountNoName].ToString();
                    curCardNumber = strPrimaryCardNo;
                    curClientID = masterRow[mClientid].ToString();

                    //emailTo = (DSemails.Tables["Emails"].Select("idclient = " + masterRow[mClientid].ToString)).;
                    printHeader();//if page is based on account no


                    totNoOfCardStat++;
                } // End of if(prevAccountNo != masterRow[accountNoName].ToString())
                else
                {
                    continue;
                }

                //foreach (DataRow dRow in mRow.GetChildRows(StatementNoDRel))
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
                    //if ((detailRow[dPostingdate] == DBNull.Value) && (detailRow[dDocno] == DBNull.Value)) continue;// Exclude On-Hold Transactions 
                    curAccRows++;
                    //totNetUsage = calculateCrDb(totNetUsage, clsBasValid.validateNum(detailRow[dBilltranamount]), clsBasValid.validateStr(detailRow[dBilltranamountsign]));
                    CurPageRec4Dtl = CurPageRec4Dtl + 1;
                    printDetail();

                } //end of detail foreach
                curCrdNoInAcc++;
                printCardFooter();//if pages is based on account
                CurPageRec4Dtl = 0; //>pageNo=1; if pages is based on account
                curAccRows = 0;
                strEmailFromTmp = emailLabelTmp = emailSentResault = string.Empty;
            } //end of Master foreach
            //emailStr = emailStr.Trim();
            emailTo = emailTo.Trim();
            emailLabelTmp = emailLabel;
            if (!string.IsNullOrEmpty(emailStr.ToString()) && !string.IsNullOrEmpty(emailTo)) // emailStr != string.Empty && emailStr != null emailTo != string.Empty && emailTo != null
                SendEmail(emailStr.ToString(), "", emailTo);
            if (!string.IsNullOrEmpty(emailStr.ToString()) && string.IsNullOrEmpty(emailTo)) // emailStr != string.Empty && emailStr != null emailTo != string.Empty && emailTo != null
            {
                streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|Without Email");
                noOfWithoutEmails++;
                emailTo = strEmailFrom; //"statement_Program@emp-group.com";
                emailSentResault = "|Without Email";
                emailLabelTmp = emailLabel + " Acc:" + curAccountNumber + " Phone:" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!");
                if (isSentBadEmailVal)
                    SendEmail(emailStr.ToString(), "", emailTo);
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

            streamEmails.Flush();
            streamEmails.Close();
            fileEmails.Close();

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

            DSstatement.Dispose();
        }
        return rtrnStr;
    }


    protected virtual void printHeader()
    {
    }

    protected virtual void printDetail()
    {
    }

    protected virtual void printCardFooter()
    {
    }

    protected virtual void calcAccountRows()
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

    protected string MakeHeaderStr(string pStr, bool isBold, bool isHeader)
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

    protected string ValidateStr(string pStr)
    {
        if (pStr.Trim().Length < 1)
            pStr = "&nbsp;";
        return pStr;
    }

    protected void SendEmail(string pBody, string pSubject, string pTo)
    {
        clsEmail sndMail = new clsEmail();
        try
        {
        if (!basText.isValideEmail(pTo))
        {
            streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + pTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|"+clsCnfg.readSetting("strValidEmail")?? "Bad Email");
            noOfBadEmails++;
            return;
        }

       
            //string pFrom, pSubject, pBody;
            //lblStatus.Text = string.Empty;
            //pFrom = "mmohammed@emp-group.com";//mmohammed@mscc.local"mmohammed@emp-group.com"mmohammed@emp-group.com
            //return;

            //streamEmails.WriteLine(curAccountNumber + "|" + curCardNumber + "|" + curClientID + "|" + emailTo);
        //streamEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + pTo + "|" + strMobileNum.Trim() + "|" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + emailSentResault);
            //return;
            //string[] strAray;
            //if (emailTo != "lucky.ighade@zenithbank.com")//ehimolen@yahoo.com
            //return;
            pLstTo.Clear();//ArrayList pLstTo = new ArrayList(), pLstAttachedFile = new ArrayList();, pLstCC = new ArrayList(), pLstBCC = new ArrayList();
            pSubject = emailLabelTmp; //emailLabel 
            //>>pSubject = emailLabel; // "BAI statement for 02/2008";
            //pTo = emailTo;// "mmohammed@emp-group.com";
            if (pTo.EndsWith("."))//ehimolen@yahoo.com
                pTo = pTo + "com";
            pLstTo.Add(pTo);//"mmohammed@emp-group.com"  "mhrap@yahoo.com""mhrap@hotmail.com"  "dossantf@emp-group.com" "mhrap@hotmail.com" "Tmahfouz61@yahoo.com"  "nazab@emp-group.com" "developers@emp-group.com""nazab@emp-group.com"  "wbaioumy@emp-group.com""mmohammed@emp-group.com" "mmohammed@emp-group.com""nazab@emp-group.com"
            //return;0
            //pLstCC.Add("mscc_emails@yahoo.com");
            //pLstBCC.Add("mmohammed@emp-group.com");
            //pLstBCC.Add("nazab@emp-group.com");
            //pLstBCC.Add("mhrap@yahoo.com");
            //pLstBCC.Add("amostafa@emp-group.com");

            //if (!pTo.ToUpper().EndsWith("DIAMONDBANK.COM"))//ehimolen@yahoo.com
            //  return;
            //pLstTo.Add("mmohammed@emp-group.com");//"mmohammed@emp-group.com""nazab@emp-group.com"


            //pLstTo.Add("ashorungbe@yahoo.com");
            //pLstTo.Add("adeyinka.shorungbe@zenithbank.com");
            //pLstTo.Add("nazab@emp-group.com");
            //pLstTo.Add("hfawzy@emp-group.com");
            //pLstTo.Add("mmohammed@emp-group.com");//mmohammed@mscc.localmmohammed@emp-group.com nazab@emp-group.commmohammed@emp-group.com
            //clsEmail sndMail = new clsEmail();
            sndMail.emailFromName = emailFromNameStr;
            //if (!sndMail.sendEmailHTML(strEmailFrom, pLstTo, pLstCC, pLstBCC, pLstAttachedFile, pSubject, pBody, clsCnfg.readSetting("SmtpServer")))
            strEmailFromTmp = strEmailFrom;
            if (isUseAlterEmailVal)
            {
                if (pTo.ToLower().IndexOf(AlterEmailCondVal) > 0)//"socgen.com"
                    strEmailFromTmp = AlterEmailVal;
                else
                    strEmailFromTmp = strEmailFrom;
            }

            numOfTry = 0;
            noOfEmails++;
            if (numMail == 0)
            {
                //pLstBCC.Add("mmohammed@emp-group.com");
                pLstBCC.Add("statement@emp-group.com");
            }
            while (!sndMail.sendEmailAttachPicFile(strEmailFromTmp, pLstTo, pLstCC, pLstBCC, pLstAttachedPic, pLstAttachedFile, pSubject, pBody, clsCnfg.readSetting("SmtpServer")))
            //while (!sndMail.sendEmailAttachPic(strEmailFromTmp, pLstTo, pLstCC, pLstBCC, pLstAttachedPic, pSubject, pBody, clsCnfg.readSetting("SmtpServer")))
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
            streamEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + pTo + "|" + strMobileNum.Trim() + "|" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + emailSentResault);

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
        catch (Exception ex)   //(NotSupportedException ex) (Exception ex)  //
        {
            //clsBasErrors.catchError(ex);
            streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + pTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "| Email " + pTo + " , Err Message >> " + ex.Message + ", Err Desc >> ");
            noOfBadEmails++;
        }
        finally
        {
            sndMail = null;
            emailStr = new StringBuilder(""); emailTo = strMobileNum = string.Empty; curAccountNumber = string.Empty; curCardNumber = string.Empty; curClientID = string.Empty;
            wait2NextEmail(prvEmail, pTo, waitPeriodVal);//     System.Threading.Thread.Sleep(waitPeriodVal);//400
            prvEmail = pTo;

        }
    }


    protected void printStatementSummary()
    {
        streamSummary.WriteLine(strBankName + " E-Statement Summary");
        streamSummary.WriteLine("__________________________");
        streamSummary.WriteLine("");
        streamSummary.WriteLine("Statements Sent by Email   " + noOfEmails.ToString());
        streamSummary.WriteLine("Statements have bad Email  " + noOfBadEmails.ToString());
        streamSummary.WriteLine("Statements without Email  " + noOfWithoutEmails.ToString());
        //    streamSummary.WriteLine("No of Transactions " + totNoOfTransactions.ToString());
    }

    protected void clientEmailChecksum()
    {
        emailStr.Append(lineSeparator);
        emailStr.Append(lineSeparator);
        emailStr.Append(lineSeparator);
        emailStr.Append(@"<tr><td width='98%' colSpan='6'><font size='1'>" + clsCheckSum.stringMD5(extAccNum.Trim() + Convert.ToDateTime(masterRow[mStatementdateto]).ToString("dd/MM/yyyy")) + "</font></td></tr>");
        emailStr.Append(lineSeparator);
    }



    protected virtual void reSetupTheValues()
    {
    }

    protected virtual void wait2NextEmail(string pPrvEmail, string pNxtEmail, int pWaitPeriod)
    {
        if (pPrvEmail.IndexOf("@") > 0 && pNxtEmail.IndexOf("@") > 0)
            if (pPrvEmail.ToLower().Substring(pPrvEmail.IndexOf("@")) == pNxtEmail.ToLower().Substring(pNxtEmail.IndexOf("@")))
                System.Threading.Thread.Sleep(waitPeriodVal);//7000
            else
                System.Threading.Thread.Sleep(1000);
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

    public string statMessageFileMonthly
    {
        get { return statMessageFileMonthlyVal; }
        set { statMessageFileMonthlyVal = value; }
    }//statMessageFileMonthly

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

    public string bankLogo
    {
        set
        {
            strBankLogo = value;
            pLstAttachedPic.Add(strBankLogo);//@"D:\Web\Email\BAI\Logo.gif"
        }
    }// bankLogo

    public string visaLogo
    {
        set
        {
            strVisaLogo = value;
            pLstAttachedPic.Add(strVisaLogo);
        }
    }// visaLogo

    public string backGround
    {
        set
        {
            strbackGround = value;
            pLstAttachedPic.Add(strbackGround);//@"D:\pC#\exe\FilesForPrograms\frmBackground.jpg"
        }
    }// backGround

    public string WaterMark
    {
        set
        {
            strWaterMark = value;
            pLstAttachedPic.Add(strWaterMark);//@"D:\pC#\exe\FilesForPrograms\frmBackground.jpg"
        }
    }// WaterMark

    public string addPicture
    {
        set { pLstAttachedPic.Add(value); }
    }

    public string addAttachFile
    {
        set { pLstAttachedFile.Add(value); }
    }

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

    protected string validateEmpty(string pStr)
    {
        if (pStr.Length < 1)
            pStr = "&nbsp;";
        else
            pStr = pStr.Trim();

        return pStr;
    }

    public string logoAlignment
    {
        set { logoAlignmentStr = value; }
    }// logoAlignment

    public string emailFromName
    {
        get { return emailFromNameStr; }
        set { emailFromNameStr = value; }
    }// emailFromName

    public bool isUseAlterEmail
    {
        set { isUseAlterEmailVal = value; }
    }// isUseAlterEmail

    public string AlterEmail
    {
        set { AlterEmailVal = value; }
    }// AlterEmail

    public string AlterEmailCond
    {
        set { AlterEmailCondVal = value; }
    }// AlterEmailCond

    public bool isSentBadEmail
    {
        set { isSentBadEmailVal = value; }
    }// isSentBadEmail

    ~clsStatHtmlCorp()
    {
        DSstatement.Dispose();
    }
}
