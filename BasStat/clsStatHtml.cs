using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Xml;
using System.Collections;


public class clsStatHtml : clsBasStatement
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
    private clsValidateEmail valdEmail = new clsValidateEmail();

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

    protected StringBuilder emailStr = new StringBuilder("");
    protected string emailLabel, strBankLogo, strMasterCardLogo = @"D:\pC#\ProjData\Statement\_Background\MasterCardLogo.jpg", strVisaLogo = @"D:\pC#\ProjData\Statement\_Background\VisaLogo.gif", strfacebookLogo, strtwitterLogo, stryoutubeLogo, strgooglePlusLogo, strMidBanner, strBottomBanner;
    protected string strEmailFrom = "mabouleila@emp-group.com", strEmailFromTmp = string.Empty;
    protected string strbankWebLink = "www.emp-group.com";
    protected string strbankWebLinkService = "www.emp-group.com";
    protected string strbankfacebooklink = "www.facebook.com";
    protected string strbanktwitterlink = "www.twitter.com";
    protected string strbankyoutubelink = "www.youtube.com";
    protected string strbankgooglePlusLink = "www.google.com";
    protected string emailTo = string.Empty, curAccountNumber, curCardNumber, curClientID;//, curEmail
    protected string strEmailFileName, strNoEmailFileName;
    protected frmStatementFile frmMain;
    protected int totRec = 1, numMail, numOfTry = 0;
    protected ArrayList pLstCC = new ArrayList(), pLstBCC = new ArrayList();
    protected string logoAlignmentStr = "center";

    protected string strbackGround = @"D:\pC#\ProjData\Statement\_Background\Background06.jpg";
    protected string strWaterMark = @"D:\pC#\ProjData\Statement\FBN\watermark.jpg";
    protected bool isRewardVal = false;
    protected bool isPrepaidVal = false;
    protected bool createCorporateVal = false;
    protected string accountNoName = mAccountno;
    protected string accountLimit = mAccountlim;
    protected string accountAvailableLimit = mAccountavailablelim;
    protected string rewardCondVal = "'New Reward Contract'";//'Reward Contract'
    protected string strProductCond = string.Empty;
    protected string prepaidCond = string.Empty;
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
    protected string fontTypeSize2 = "<font color='#8B0000' size='2'>";
    protected string fontTypeSize3 = "<font color='#8B0000' size='2'>";
    protected string lineSeparator = "<tr><td width='98%' height='25' style='width: 61%'>&nbsp;</td></tr>";
    protected bool isUseAlterEmailVal = false, isSentBadEmailVal = true;
    protected string AlterEmailVal, AlterEmailCondVal;
    protected string emailSentResult = string.Empty;
    protected string prvEmail = string.Empty;
    protected string ClientID = mClientid;
    private bool HasSenderVal = false;
    protected string VCurrency = "";
    private string strPaymentSystem = string.Empty;
    private string strBillingCycle = string.Empty;

    public clsStatHtml()
        {

        }

    public virtual string Statement(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pCurDate, string pStmntType, bool pAppendData)
        {
        string rtrnStr = "Successfully Generate " + pBankName;
        int curMonth = pCurDate.Month;
        //bool preExit = true;
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
            strNoEmailFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "_" + pStmntType + "\\" + pBankName + "_" + pStmntType + "_" + "WithoutEmails" + "_" + vCurDate.ToString("yyyyMM");//+ ".txt"
            pStrFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "_" + pStmntType + "\\" + pBankName + "_" + pStmntType + "_" + vCurDate.ToString("yyyyMM") + ".txt";
            strBankName = pBankName;

            if (pBankCode == 50)
                emailLabel = "ACCESS BANK GHANA Statement for " + vCurDate.ToString("MM/yyyy"); //"BAI statement for 02/2008"
            else
                emailLabel = pBankName + " statement for " + vCurDate.ToString("MM/yyyy"); //"BAI statement for 02/2008
            strOutputFile = pStrFileName;

            // open emails file
            fileEmails = new FileStream(strEmailFileName + ".txt", FileMode.Create); //Create
            streamEmails = new StreamWriter(fileEmails, Encoding.Default);
            streamEmails.AutoFlush = true;
            //streamEmails.WriteLine("AccountNumber" + "|" + "CardNumber" + "|" + mClientid + "|" + "Email");
            streamEmails.WriteLine("AccountNumber" + "|" + "ClientID" + "|" + "Email" + "|" + "MobilePhone" + "|" + "Date Time");

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
            streamSummary.AutoFlush = true;

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

            if (isRewardVal)
            {
                //maintainData.curRewardCond = rewardCond;
                //maintainData.fixReward(pBankCode, rewardCond);
                strMainTableCond = "m.contracttype != " + rewardCondVal;
                strSubTableCond = "d.trandescription != 'Calculated Points'";
                getReward(pBankCode);
            }

            if (isPrepaidVal)
            {
                strMainTableCond += " m.cardproduct in " + PrepaidCondition;
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.cardproduct in " + PrepaidCondition + ")";
            }

            if (pBankCode == 73)
            {
                MainTableCond = " m.contracttype like '" + strPaymentSystem + "%' and m.contracttype like '%" + strBillingCycle + "%'";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype like '" + strPaymentSystem + "%' and x.contracttype like '%" + strBillingCycle + "%')";
            }

            if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("SBN_Credit_ClassicPlatinumGold_ClientsEmails"))
            {
                MainTableCond = " m.contracttype in " + strProductCond + "";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + ")";
            }

            if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("SBN_Credit_Infinite_ClientsEmails"))
            {
                MainTableCond = " m.contracttype in " + strProductCond + "";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + ")";
            }

            if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("SBN_Credit_Signature_ClientsEmails"))
            {
                MainTableCond = " m.contracttype in " + strProductCond + "";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + ")";
            }

            if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("FBN_CreditClientsEmails"))
            {
                MainTableCond = " m.contracttype in " + strProductCond + "";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + ")";
            }

            if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("FBN_Credit_Classic_ClientsEmails"))
            {
                MainTableCond = " m.contracttype in " + strProductCond + "";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + ")";
            }

            if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("FBN_Credit_Gold_ClientsEmails"))
            {
                MainTableCond = " m.contracttype in " + strProductCond + "";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + ")";
            }

            if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("FBN_Credit_Infinite_ClientsEmails"))
            {
                MainTableCond = " m.contracttype in " + strProductCond + "";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + ")";
            }

            if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("FBN_Credit_Classic_Platinum_ClientsEmails"))
            {
                MainTableCond = " m.contracttype in " + strProductCond + "";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + ")";
            }

            if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("SBP_CorporateEmails"))
            {
                MainTableCond = " m.contracttype in " + strProductCond + "";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + ")";
            }

            if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("SBP_VISA_PrepaidEmails"))
            {
                MainTableCond = " m.cardproduct in " + PrepaidCondition + "";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.cardproduct in " + PrepaidCondition + ")";
            }

            if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("SBP_VISA_Prepaid_NTDC_Emails"))
            {
                MainTableCond = " m.cardproduct in " + PrepaidCondition + "";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.cardproduct in " + PrepaidCondition + ")";
            }

            if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("SBP_MasterCard_Platinum_CreditClientsEmails"))
            {
                MainTableCond = " m.contracttype in " + strProductCond + "";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + ")";
            }

            if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("SBP_Credit_Classic_PremiumClientsEmails"))
            {
                MainTableCond = " m.contracttype in " + strProductCond + "";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + ")";
            }

            if (VCurrency != string.Empty)
            {
                MainTableCond = "m.Accountcurrency ='" + VCurrency.ToString() + "'";
                supTableCond = "d.Accountcurrency ='" + VCurrency.ToString() + "'";
            }

            if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("BK_VISA_RWF_ClientsEmails"))
            {
                MainTableCond = " m.contracttype in " + strProductCond + " and m.Accountcurrency ='RWF'";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + " and x.Accountcurrency ='RWF')";
            }

            if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("BK_VISA_USD_ClientsEmails"))
            {
                MainTableCond = " m.contracttype in " + strProductCond + " and m.Accountcurrency ='USD'";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + " and x.Accountcurrency ='USD')";
            }

            if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("BK_MasterCard_RWF_ClientsEmails"))
            {
                MainTableCond = " m.contracttype in " + strProductCond + " and m.Accountcurrency ='RWF'";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + " and x.Accountcurrency ='RWF')";
            }

            if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("BK_MasterCard_USD_ClientsEmails"))
            {
                MainTableCond = " m.contracttype in " + strProductCond + " and m.Accountcurrency ='USD'";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + " and x.Accountcurrency ='USD')";
            }

            if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("BK_VISA_VIP_RWF_ClientsEmails"))
            {
                MainTableCond = " m.contracttype in " + strProductCond + " and m.Accountcurrency ='RWF'";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + " and x.Accountcurrency ='RWF')";
            }

            if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("BK_VISA_VIP_USD_ClientsEmails"))
            {
                MainTableCond = " m.contracttype in " + strProductCond + " and m.Accountcurrency ='USD'";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + " and x.Accountcurrency ='USD')";
            }

            if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("BK_MasterCard_VIP_RWF_ClientsEmails"))
            {
                MainTableCond = " m.contracttype in " + strProductCond + " and m.Accountcurrency ='RWF'";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + " and x.Accountcurrency ='RWF')";
            }

            if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("BK_MasterCard_VIP_USD_ClientsEmails"))
            {
                MainTableCond = " m.contracttype in " + strProductCond + " and m.Accountcurrency ='USD'";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + " and x.Accountcurrency ='USD')";
            }

            if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("BK_MasterCard_Corporate_Cardholder_RWF_ClientsEmails"))
            {
                MainTableCond = " m.contracttype in " + strProductCond + " and m.Accountcurrency ='RWF'";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + " and x.Accountcurrency ='RWF')";
            }

            if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("BK_MasterCard_Corporate_Cardholder_USD_ClientsEmails"))
            {
                MainTableCond = " m.contracttype in " + strProductCond + " and m.Accountcurrency ='USD'";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + " and x.Accountcurrency ='USD')";
            }

            if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("BK_MasterCard_Corporate_Cardholder_VIP_RWF_ClientsEmails"))
            {
                MainTableCond = " m.contracttype in " + strProductCond + " and m.Accountcurrency ='RWF'";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + " and x.Accountcurrency ='RWF')";
            }

            if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("BK_MasterCard_Corporate_Cardholder_VIP_USD_ClientsEmails"))
            {
                MainTableCond = " m.contracttype in " + strProductCond + " and m.Accountcurrency ='USD'";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + " and x.Accountcurrency ='USD')";
            }

            if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("I&M_Prepaid_ClientsEmails"))
            {
                MainTableCond = " m.cardproduct in " + PrepaidCondition + "";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.cardproduct in " + PrepaidCondition + ")";
            }



            //////////////////MAMR Test Emailing

            ////////////////if (MainTableCond.Trim() == "")
            ////////////////    strMainTableCond += "  m.ACCOUNTNO in ('Credit_AON_6621',  'Credit_AON_6651', 'Credit_AON_5671', 'Credit_AON_5683') ";//strWhereCond


            //////////////////MAMR Test Emailing



            FillStatementDataSet(pBankCode, "vip"); //DSstatement =  //10); //3


            getClientEmail(pBankCode);
            pageNo = 1; totalCardPages = 0;
            curCardNo = String.Empty;
            curAccountNo = String.Empty;
            frmMain.BeginInvoke(frmMain.setMinMaxProgressDelegate, new object[] { DSstatement.Tables["tStatementMasterTable"].Rows.Count });
            foreach (DataRow mRow in DSstatement.Tables["tStatementMasterTable"].Rows)
            {
                frmMain.BeginInvoke(frmMain.setProgressDelegate, new object[] { totRec++ });
                masterRow = mRow;
                reSetupTheValues();
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
                        //if (HasSender == true)
                        //    SendEmailWithDifferentSender(emailStr.ToString(), "", emailTo);
                        //else
                        SendEmail(emailStr.ToString(), "", emailTo);
                    if (!string.IsNullOrEmpty(emailStr.ToString()) && string.IsNullOrEmpty(emailTo)) //emailStr != string.Empty && emailStr != null  emailTo != string.Empty && emailTo != null
                    {
                        emailLabelTmp = string.Empty;
                        if (emailRow != null)
                            emailLabelTmp = (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!");
                        //emailTo = strEmailFrom; //"statement_Program@emp-group.com";

                        emailLabelTmp = emailLabel + " Acc:" + curAccountNumber + " Phone:" + emailLabelTmp;
                        //if (HasSender == true)
                        //    SendEmailWithDifferentSender(emailStr.ToString(), "", emailTo);
                        //else

                        if (strEmailFrom.ToUpper().EndsWith("BK.RW"))
                        {
                            if (valdEmail.isValideEmail(emailTo) != "ValidEmail")
                            {
                                emailTo = strEmailFrom;
                                streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + emailLabelTmp + "|Without or bad Email and FW to " + emailTo);
                                SendEmail(emailStr.ToString(), "", emailTo);
                                //noOfBadEmails++;
                                noOfWithoutEmails++;

                            }
                        }
                        else
                        {
                            streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + emailLabelTmp + "|Without Email");
                            noOfWithoutEmails++;

                        }

                        //SendEmail(emailStr.ToString(), "", emailTo);

                        //noOfWithoutEmails++;
                    }

                    emailStr = new StringBuilder("");
                    cardsRows = DSstatement.Tables["tStatementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[accountNoName]).ToString().Trim() + "'");

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
                    emailRows = DSemails.Tables["Emails"].Select("idclient = " + masterRow[ClientID].ToString());
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
                    //emailTo = "organizaçõesnatriki@hotmail.com";
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
                    if ((detailRow[dPostingdate] == DBNull.Value) && (detailRow[dDocno] == DBNull.Value)) continue;// Exclude On-Hold Transactions 
                    curAccRows++;
                    totNetUsage = calculateCrDb(totNetUsage, clsBasValid.validateNum(detailRow[dBilltranamount]), clsBasValid.validateStr(detailRow[dBilltranamountsign]));
                    CurPageRec4Dtl = CurPageRec4Dtl + 1;
                    printDetail();

                } //end of detail foreach
                curCrdNoInAcc++;
                printCardFooter();//if pages is based on account
                CurPageRec4Dtl = 0; //>pageNo=1; if pages is based on account
                curAccRows = 0;
                strEmailFromTmp = emailLabelTmp = emailSentResult = string.Empty;
            } //end of Master foreach
            //emailStr = emailStr.Trim();
            emailTo = emailTo.Trim();
            emailLabelTmp = emailLabel;
            if (!string.IsNullOrEmpty(emailStr.ToString()) && !string.IsNullOrEmpty(emailTo)) // emailStr != string.Empty && emailStr != null emailTo != string.Empty && emailTo != null
                //if (HasSender == true)
                //    SendEmailWithDifferentSender(emailStr.ToString(), "", emailTo);
                //else
                SendEmail(emailStr.ToString(), "", emailTo);
            if (!string.IsNullOrEmpty(emailStr.ToString()) && string.IsNullOrEmpty(emailTo)) // emailStr != string.Empty && emailStr != null emailTo != string.Empty && emailTo != null
            //{
            //if (emailRow != null) // EDT-729
            //    {// EDT-729
            //    streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|Without Email");
            //    noOfWithoutEmails++;
            //    //emailTo = strEmailFrom; //"statement_Program@emp-group.com";
            //    emailSentResult = "|Without Email";
            //    emailLabelTmp = emailLabel + " Acc:" + curAccountNumber + " Phone:" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!");
            //    }// EDT-729
            //if (isSentBadEmailVal)
            //    //if (HasSender == true)
            //    //    SendEmailWithDifferentSender(emailStr.ToString(), "", emailTo);
            //    //else
            //    SendEmail(emailStr.ToString(), "", emailTo);

            //}
            {
                emailLabelTmp = string.Empty;
                if (emailRow != null)
                    emailLabelTmp = (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!");
                //emailTo = strEmailFrom; //"statement_Program@emp-group.com";

                emailLabelTmp = emailLabel + " Acc:" + curAccountNumber + " Phone:" + emailLabelTmp;
                //if (HasSender == true)
                //    SendEmailWithDifferentSender(emailStr.ToString(), "", emailTo);
                //else

                if (strEmailFrom.ToUpper().EndsWith("BK.RW"))
                {
                    if (valdEmail.isValideEmail(emailTo) != "ValidEmail")
                    {
                        emailTo = strEmailFrom;
                        streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + emailLabelTmp + "|Without or bad Email and FW to " + emailTo);
                        SendEmail(emailStr.ToString(), "", emailTo);
                        //noOfBadEmails++;
                        noOfWithoutEmails++;

                    }
                }
                else
                {
                    streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + emailLabelTmp + "|Without Email");
                    noOfWithoutEmails++;

                }


                //SendEmail(emailStr.ToString(), "", emailTo);

                //noOfWithoutEmails++;
            }


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
            //if (pStmntType == "CorporateClientsEmails")
            //{
            //    ArrayList aryLstFilesCorp = new ArrayList();
            //    clsStatHtmlGnrlFBNCompany statHtmlFBNcorporate = new clsStatHtmlGnrlFBNCompany();
            //    statHtmlFBNcorporate.setFrm = frmMain;
            //    statHtmlFBNcorporate.emailFromName = "First Bank of Nigeria - Statement";//"mmohammed@emp-group.com""Nermine.Bahaa-Eldin@socgen.com"
            //    statHtmlFBNcorporate.emailFrom = "cardservices@firstbanknigeria.com";//mmohammed@emp-group.com
            //    statHtmlFBNcorporate.bankWebLink = "www.firstbanknigeria.com";
            //    statHtmlFBNcorporate.bankLogo = @"D:\pC#\ProjData\Statement\FBN\logo.jpg";
            //    statHtmlFBNcorporate.visaLogo = @"D:\pC#\ProjData\Statement\_Background\visalogo.jpg";
            //    statHtmlFBNcorporate.WaterMark = @"D:\pC#\ProjData\Statement\FBN\watermark.jpg";
            //    statHtmlFBNcorporate.BottomBanner = @"D:\pC#\ProjData\Statement\FBN\banner.jpg";
            //    rtrnStr = statHtmlFBNcorporate.Statement(clsSessionValues.basPath + "_WaitReply\\", pBankName, pBankCode, pStrFileName, pCurDate, pStmntType, pAppendData);
            //    //foreach (object str in aryLstFilesCorp)
            //    //    aryLstFiles.Add((string)str);
            //    statHtmlFBNcorporate = null;
            //}
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
        //mainRows = DSstatement.Tables["tStatementMasterTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[accountNoName]) + "'");


        // bsayed: handling for corporate cardholders (individual),
        // Case is applied by bank to account for corporate coslidated until they are reviewed & isolated
        if (createCorporateVal)
        {
            if (masterRow[mBranch].ToString().Trim() == "7")
            {
                mainRows = DSstatement.Tables["tStatementMasterTable"].Select("CARDACCOUNTNO = '" + clsBasValid.validateStr(masterRow[accountNoName]) + "'");
            }
            else
            {
                mainRows = DSstatement.Tables["tStatementMasterTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[accountNoName]) + "'");
            }
        }
        else
        {
            mainRows = DSstatement.Tables["tStatementMasterTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[accountNoName]) + "'");
        }

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

    //protected void SendEmailWithDifferentSender(string pBody, string pSubject, string pTo)
    //    {
    //    if (strEmailFrom.ToUpper().EndsWith("BK.RW"))
    //        {
    //        if (valdEmail.isValideEmail(pTo) != "ValidEmail")
    //            {
    //            pTo = strEmailFrom;
    //            streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|"+clsCnfg.readSetting("strValidEmail")?? "Bad Email");
    //            noOfBadEmails++;
    //            }
    //        }
    //    else
    //        {
    //        if (!basText.isValideEmail(pTo))
    //            {
    //            if (!string.IsNullOrEmpty(emailTo))
    //                {
    //                streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|"+clsCnfg.readSetting("strValidEmail")?? "Bad Email");
    //                noOfBadEmails++;
    //                }
    //            return;
    //            }
    //        }
    //    try
    //        {
    //        //string pFrom, pSubject, pBody;
    //        //lblStatus.Text = string.Empty;
    //        //pFrom = "mmohammed@emp-group.com";//mmohammed@mscc.local"mmohammed@emp-group.com"mmohammed@emp-group.com
    //        //return;

    //        //streamEmails.WriteLine(curAccountNumber + "|" + curCardNumber + "|" + curClientID + "|" + emailTo);
    //        streamEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + strMobileNum.Trim() + "|" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + emailSentResult);
    //        //return;
    //        //string[] strAray;
    //        //if (emailTo != "lucky.ighade@zenithbank.com")//ehimolen@yahoo.com
    //        //return;
    //        pLstTo.Clear();//ArrayList pLstTo = new ArrayList(), pLstAttachedFile = new ArrayList();, pLstCC = new ArrayList(), pLstBCC = new ArrayList();
    //        pLstCC.Clear();
    //        pLstBCC.Clear();
    //        pSubject = emailLabelTmp; //emailLabel 
    //        //>>pSubject = emailLabel; // "BAI statement for 02/2008";
    //        //pTo = emailTo;// "mmohammed@emp-group.com";
    //        if (pTo.EndsWith("."))//ehimolen@yahoo.com
    //            pTo = pTo + "com";
    //        pLstTo.Add(pTo);//"mmohammed@emp-group.com"  "mhrap@yahoo.com""mhrap@hotmail.com"  "dossantf@emp-group.com" "mhrap@hotmail.com" "Tmahfouz61@yahoo.com"  "nazab@emp-group.com" "developers@emp-group.com""nazab@emp-group.com"  "wbaioumy@emp-group.com""mmohammed@emp-group.com" "mmohammed@emp-group.com""nazab@emp-group.com"

    //        //return;0
    //        //pLstCC.Add("mscc_emails@yahoo.com");
    //        //pLstBCC.Add("mmohammed@emp-group.com");
    //        //pLstBCC.Add("nazab@emp-group.com");
    //        //pLstBCC.Add("mhrap@yahoo.com");
    //        //pLstBCC.Add("amostafa@emp-group.com");

    //        //if (!pTo.ToLower().EndsWith("diamondbank.com"))
    //        //if (!pTo.ToLower().EndsWith("skyebankng.com"))
    //        //if (!pTo.ToLower().EndsWith("bpc.ao"))
    //        //return;
    //        //if (pTo.ToUpper().IndexOf("YAHOO") > -1 || pTo.ToUpper().IndexOf("HOTMAIL") > -1)
    //        //  pBody.Replace("cid:", "");

    //        //if (!pTo.ToUpper().EndsWith("DIAMONDBANK.COM"))//ehimolen@yahoo.com
    //        //  return;
    //        //pLstTo.Add("mmohammed@emp-group.com");//"mmohammed@emp-group.com""nazab@emp-group.com"


    //        //pLstTo.Add("ashorungbe@yahoo.com");
    //        //pLstTo.Add("adeyinka.shorungbe@zenithbank.com");
    //        //pLstTo.Add("nazab@emp-group.com");
    //        //pLstTo.Add("hfawzy@emp-group.com");
    //        //pLstTo.Add("mmohammed@emp-group.com");//mmohammed@mscc.localmmohammed@emp-group.com nazab@emp-group.commmohammed@emp-group.com
    //        clsEmail sndMail = new clsEmail();
    //        sndMail.emailFromName = emailFromNameStr;
    //        //if (!sndMail.sendEmailHTML(strEmailFrom, pLstTo, pLstCC, pLstBCC, pLstAttachedFile, pSubject, pBody, clsCnfg.readSetting("SmtpServer")))
    //        strEmailFromTmp = strEmailFrom;
    //        if (isUseAlterEmailVal)
    //            {
    //            if (pTo.ToLower().IndexOf(AlterEmailCondVal) > 0)//"socgen.com"
    //                strEmailFromTmp = AlterEmailVal;
    //            else
    //                strEmailFromTmp = strEmailFrom;
    //            }

    //        numOfTry = 0;
    //        noOfEmails++;
    //        if (numMail == 0)
    //            {
    //            //pLstBCC.Add("mmohammed@emp-group.com");
    //            pLstBCC.Add("statement@emp-group.com");
    //            if (strEmailFromTmp.ToUpper().EndsWith("BK.RW"))
    //                pLstBCC.Add("andibo@bk.rw");
    //            }
    //        if (strEmailFromTmp.ToUpper().EndsWith("BK.RW"))
    //            {
    //            if (pLstCC.Count == 0)
    //                pLstCC.Add("visacardstatement@bk.rw");
    //            if (pLstBCC.Count == 0)
    //                pLstBCC.Add("andibo@bk.rw");
    //            }

    //        if (strEmailFromTmp.ToUpper().EndsWith("GTBANK.COM"))
    //            {
    //            if (pLstCC.Count == 0)
    //                pLstCC.Add("Creditcardteam@gtbank.com");
    //            }

    //        //if (!pLstTo.Contains("baicartao@bancobai.ao"))
    //        //{
    //        while (!sndMail.sendEmailAttachPicFilewithDiffernetSender(strEmailFromTmp, "statement@emp-group.com", pLstTo, pLstCC, pLstBCC, pLstAttachedPic, pLstAttachedFile, pSubject, pBody, clsCnfg.readSetting("SmtpServer")))
    //        //while (!sndMail.sendEmailAttachPic(strEmailFromTmp, pLstTo, pLstCC, pLstBCC, pLstAttachedPic, pSubject, pBody, clsCnfg.readSetting("SmtpServer")))
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
    //        //}
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
    //        noOfBadEmails++;
    //        }
    //    finally
    //        {
    //        }
    //    }

    protected void SendEmail(string pBody, string pSubject, string pTo)
        {
            clsEmail sndMail = new clsEmail();
        try
        {
        //if (strEmailFrom.ToUpper().EndsWith("BK.RW"))
        //    {
        //    if (valdEmail.isValideEmail(pTo) != "ValidEmail")
        //        {
        //        pTo = strEmailFrom;
        //        streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|"+clsCnfg.readSetting("strValidEmail")?? "Bad Email " + pTo);
        //        noOfBadEmails++;
        //        return;
        //        }
        //    }
        //else
        //    {
            if (!basText.isValideEmail(pTo))
                {
                //if (!string.IsNullOrEmpty(emailTo))
                //    {
                    streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + pTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|"+clsCnfg.readSetting("strValidEmail")?? "Bad Email " + pTo);
                    noOfBadEmails++;
                    //}
                return;
                }
            //}
        //try
        //    {
            //string pFrom, pSubject, pBody;
            //lblStatus.Text = string.Empty;
            //pFrom = "mmohammed@emp-group.com";//mmohammed@mscc.local"mmohammed@emp-group.com"mmohammed@emp-group.com
            //return;

            //streamEmails.WriteLine(curAccountNumber + "|" + curCardNumber + "|" + curClientID + "|" + emailTo);
            //streamEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + strMobileNum.Trim() + "|" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + emailSentResult);
            //return;
            //string[] strAray;
            //if (emailTo != "lucky.ighade@zenithbank.com")//ehimolen@yahoo.com
            //return;
            pLstTo.Clear();//ArrayList pLstTo = new ArrayList(), pLstAttachedFile = new ArrayList();, pLstCC = new ArrayList(), pLstBCC = new ArrayList();
            pLstCC.Clear();
            pLstBCC.Clear();
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

            //if (!pTo.ToLower().EndsWith("diamondbank.com"))
            //if (!pTo.ToLower().EndsWith("skyebankng.com"))
            //if (!pTo.ToLower().EndsWith("bpc.ao"))
            //return;
            //if (pTo.ToUpper().IndexOf("YAHOO") > -1 || pTo.ToUpper().IndexOf("HOTMAIL") > -1)
            //  pBody.Replace("cid:", "");

            //if (!pTo.ToUpper().EndsWith("DIAMONDBANK.COM"))//ehimolen@yahoo.com
            //  return;
            //pLstTo.Add("mmohammed@emp-group.com");//"mmohammed@emp-group.com""nazab@emp-group.com"


            //pLstTo.Add("ashorungbe@yahoo.com");
            //pLstTo.Add("adeyinka.shorungbe@zenithbank.com");
            //pLstTo.Add("nazab@emp-group.com");
            //pLstTo.Add("hfawzy@emp-group.com");
            //pLstTo.Add("mmohammed@emp-group.com");//mmohammed@mscc.localmmohammed@emp-group.com nazab@emp-group.commmohammed@emp-group.com
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
                //if (strEmailFromTmp.ToUpper().EndsWith("BK.RW"))
                //    pLstBCC.Add("andibo@bk.rw");
                }
            //if (strEmailFromTmp.ToUpper().EndsWith("BK.RW"))
            //    {
            //    if (pLstCC.Count == 0)
            //        pLstCC.Add("visacardstatement@bk.rw");
            //    //if (pLstBCC.Count == 0)
            //    //    pLstBCC.Add("andibo@bk.rw");
            //    }

            if (strEmailFromTmp.ToUpper().EndsWith("GTBANK.COM"))
                {
                if (pLstCC.Count == 0)
                    pLstCC.Add("Creditcardteam@gtbank.com");
                }

            //if (!pLstTo.Contains("baicartao@bancobai.ao"))
            //{
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
            streamEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + pTo + "|" + strMobileNum.Trim() + "|" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + emailSentResult);

            //}
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
            ////wait2NextEmail(prvEmail, pTo, waitPeriodVal);//     System.Threading.Thread.Sleep(waitPeriodVal);//400
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
                emailStr = new StringBuilder(""); emailTo = strMobileNum = string.Empty; curAccountNumber = string.Empty; curCardNumber = string.Empty; curClientID = string.Empty;
                //wait2NextEmail(prvEmail, pTo, waitPeriodVal);//     System.Threading.Thread.Sleep(waitPeriodVal);//400
                prvEmail = pTo;
            }
        }

    protected void printStatementSummary()
        {
        if (strBankName == "BK")
            streamSummary.WriteLine(strBankName + " " + VCurrency + " E-Statement Summary");
        else
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

    protected void clientEmailChecksum2()
        {
        emailStr.Append(@"<tr><td width='98%' colSpan='6' style='font-family: Verdana'><font size='1'>" + clsCheckSum.stringMD5(extAccNum.Trim() + Convert.ToDateTime(masterRow[mStatementdateto]).ToString("dd/MM/yyyy")) + "</font></td></tr>");
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

    public bool isPrepaid
        {
        get { return isPrepaidVal; }
        set { isPrepaidVal = value; }
        }

    public string PrepaidCondition
        {
        get { return prepaidCond; }
        set { prepaidCond = value; }
        }// PrepaidCondition

    public string ProductCondition
        {
        get { return strProductCond; }
        set { strProductCond = value; }
        }// ProductCondition

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

    public string MidBanner
        {
        set
            {
            strMidBanner = value;
            pLstAttachedPic.Add(strMidBanner);//@"D:\Web\Email\BAI\Logo.gif"
            }
        }// MidBanner

    public string BottomBanner
        {
        set
            {
            strBottomBanner = value;
            pLstAttachedPic.Add(strBottomBanner);//@"D:\Web\Email\BAI\Logo.gif"
            }
        }// BottomBanner

    public string visaLogo
        {
        set
            {
            strVisaLogo = value;
            pLstAttachedPic.Add(strVisaLogo);
            }
        }// visaLogo

    public string MasterCardLogo
        {
        set
            {
            strMasterCardLogo = value;
            pLstAttachedPic.Add(strMasterCardLogo);
            }
        }// MasterCardLogo

    public string facebookLogo
        {
        set
            {
            strfacebookLogo = value;
            pLstAttachedPic.Add(strfacebookLogo);//@"D:\Web\Email\BAI\Logo.gif"
            }
        }// facebookLogo

    public string twitterLogo
        {
        set
            {
            strtwitterLogo = value;
            pLstAttachedPic.Add(strtwitterLogo);//@"D:\Web\Email\BAI\Logo.gif"
            }
        }// twitterLogo

    public string youtubeLogo
        {
        set
            {
            stryoutubeLogo = value;
            pLstAttachedPic.Add(stryoutubeLogo);//@"D:\Web\Email\BAI\Logo.gif"
            }
        }// youtubeLogo

    public string googlePlusLogo
        {
        set
            {
            strgooglePlusLogo = value;
            pLstAttachedPic.Add(strgooglePlusLogo);//@"D:\Web\Email\BAI\Logo.gif"
            }
        }// googlePlusLogo

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

    public string bankfacebookLink
        {
        get { return strbankfacebooklink; }
        set { strbankfacebooklink = value; }
        }// bankfacebookLink

    public string banktwitterLink
        {
        get { return strbanktwitterlink; }
        set { strbanktwitterlink = value; }
        }// banktwitterLink

    public string bankyoutubeLink
        {
        get { return strbankyoutubelink; }
        set { strbankyoutubelink = value; }
        }// bankyoutubeLink

    public string bankgooglePlusLink
        {
        get { return strbankgooglePlusLink; }
        set { strbankgooglePlusLink = value; }
        }// bankgooglePlusLink

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

    public bool HasSender
        {
        get { return HasSenderVal; }
        set { HasSenderVal = value; }
        }  // HasSender

    public string CurrencyFilter
        {
        get { return VCurrency; }
        set { VCurrency = value; }
        } //CurrencyFilter

    public string PaymentSystem
    {
        get { return strPaymentSystem; }
        set { strPaymentSystem = value; }
    }  // PaymentSystem

    public string BillingCycle
    {
        get { return strBillingCycle; }
        set { strBillingCycle = value; }
    }
    ~clsStatHtml()
        {
        DSstatement.Dispose();
        }
    }
