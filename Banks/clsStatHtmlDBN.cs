using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Xml;
using System.Collections;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Security;
using System.Collections.Generic;
using System.Linq;

public class clsStatHtmlDBN : clsBasStatement
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
    private string emailLabel, strBankLogo, strBackGround;
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
    private string strbottomBanner1 = string.Empty;
    private decimal TBilltranamount = 0;
    private decimal TBalance = 0;
    private string vipCondVal = string.Empty;

    public clsStatHtmlDBN()
    {
    }

    public string Statement(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pCurDate, string pStmntType, bool pAppendData, string pReportName)
    {
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
            //clsMaintainData maintainData = new clsMaintainData();
            //maintainData.matchCardBranch4Account(pBankCode);

            pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);
            vCurDate = pCurDate; //DateTime.Now.AddMonths(-1);
            strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "_" + pStmntType;
            //clsBasFile.deleteDirectory(strOutputPath);
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
            //emailLabel = pBankName + " statement for " + vCurDate.ToString("MM/yyyy"); //"BAI statement for 02/2008"
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

            // Mamr -- Set DBN to be ABP -- START

            if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("DBN_Credit_Classic_ClientsEmails"))
            {
                strMainTableCond += " m.contracttype in " + VIPCondition + " and m.cardproduct in " + ProductCondition + "";
            }
            else if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("DBN_Credit_Gold_ClientsEmails"))
            {
                strMainTableCond += " m.contracttype in " + VIPCondition + " and m.cardproduct = '" + ProductCondition + "'";
            }
            else if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("DBN_Credit_Platinum_ClientsEmails"))
            {
                strMainTableCond += " m.contracttype in " + VIPCondition + " and m.cardproduct in " + ProductCondition + "";
            }
            else if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("DBN_Credit_Platinum_USD_ClientsEmails"))
            {
                strMainTableCond += " m.contracttype in " + VIPCondition + " and m.cardproduct in " + ProductCondition + "";
            }
            else if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("DBN_VIP1-2-5_ClientsEmails"))
            {
                strMainTableCond += " m.contracttype in " + VIPCondition + " and m.cardproduct = '" + ProductCondition + "'";
            }
            else if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("DBN_VIP_ClientsEmails"))
            {
                strMainTableCond += " m.contracttype in " + VIPCondition + " and m.cardproduct = '" + ProductCondition + "'";
            }
            else if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("DBN_ParkNShop_ClientsEmails"))
            {
                strMainTableCond += " m.contracttype in " + VIPCondition + " and m.cardproduct = '" + ProductCondition + "'";
            }
            else if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("DBN_MasterCard_Credit_ClientsEmails"))
            {
                strMainTableCond += " m.contracttype in " + VIPCondition + " and m.cardproduct in " + ProductCondition + "";
            }
            else if (clsBasFile.getFileWithoutExtn(pStrFileName).StartsWith("DBN_Privilege_ClientsEmails"))
            {
                strMainTableCond += " m.contracttype in " + VIPCondition + " and m.cardproduct = '" + ProductCondition + "'";
            }

            // Mamr -- Set DBN to be ABP -- END

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
            //new frmGenerateSatatement().Show();
            frmMain.BeginInvoke(frmMain.setMinMaxProgressDelegate, new object[] { DSstatement.Tables["tStatementMasterTable"].Rows.Count });
            //generateStatus.Show();
            //generateStatus.setMinMaxProgress(DSstatement.Tables["tStatementMasterTable"].Rows.Count);
            //generateStatus.BeginInvoke(generateStatus.setMinMaxProgressDelegate, new object[] { DSstatement.Tables["tStatementMasterTable"].Rows.Count });
            foreach (DataRow mRow in DSstatement.Tables["tStatementMasterTable"].Rows)
            {
                frmMain.BeginInvoke(frmMain.setProgressDelegate, new object[] { totRec++ });
                #region ignore this account
                if (sentaccounts != null && sentaccounts.Contains(mRow[accountNoName].ToString().Trim()))
                {
                    continue;
                }
                #endregion
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
                    //emailLabelTmp = emailLabel;
                    if (!string.IsNullOrEmpty(emailStr.ToString()) && !string.IsNullOrEmpty(emailTo)) //emailStr != string.Empty && emailStr != null emailTo != string.Empty && emailTo != null
                        SendEmail(emailStr.ToString(), "", emailTo);
                    if (!string.IsNullOrEmpty(emailStr.ToString()) && string.IsNullOrEmpty(emailTo)) //emailStr != string.Empty && emailStr != null  emailTo != string.Empty && emailTo != null
                    {
                        streamNoEmails.WriteLine(curCustomerName + "|" + curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|Without Email");
                        noOfWithoutEmails++;
                        emailTo = strEmailFrom; //"statement_Program@emp-group.com";
                        //emailLabelTmp = emailLabel + " Acc:" + curAccountNumber + " Phone:" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!");
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

                    curFileName = "Access Bank " + masterRow[mAccountcurrency].ToString() + " Credit Card Statement " + vCurDate.ToString("MMMM yyyy ") + masterRow[mCardclientname].ToString();
                    curFileName = curFileName.Replace('.', ' ').Replace('/', ' ').Replace(' ', '_');


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
                        //updated by Mahmoud Amr
                        emailTo = emailRows[i][1].ToString().Trim();
                        //emailTo = "mamr@network.com.eg"; //emailRows[i][1].ToString().Trim();
                        emailRow = emailRows[i];
                    }
                    curAccountNumber = masterRow[accountNoName].ToString();
                    curCardNumber = strPrimaryCardNo;
                    curClientID = masterRow[ClientID].ToString();
                    curCustomerName = masterRow[mCustomername].ToString();

                    //emailTo = (DSemails.Tables["Emails"].Select("idclient = " + masterRow[mClientid].ToString)).;
                    //printHeader();//if page is based on account no

                    printBody();

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
                    //stmNo = detailRow[dStatementnumber].ToString();
                    stmtNo = detailRow[dStatementno].ToString();
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
                    //printDetail();
                    //if(!((detailRow[dPostingdate]== DBNull.Value) && (detailRow[dDocno]== DBNull.Value))) 

                } //end of detail foreach
                //printCardFooter();//if pages is based on card
                // if(masterRow[mCardprimary].ToString() == "Y" && (masterRow[mCardstate].ToString() == "Given" || masterRow[mCardstate].ToString() == "Embossed" || masterRow[mCardstate].ToString() == "New"))
                curCrdNoInAcc++;
                //-if ((curAccRows >= totAccRows && totAccRows != 0) || (totAccRows == 0 && curCrdNoInAcc == totCrdNoInAcc))
                //-{
                //-  completePageDetailRecords();
                //printCardFooter();//if pages is based on account

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
                emailLabelTmp = "Your Access Bank " + masterRow[mAccountcurrency].ToString() + " Credit Card Statement " + vCurDate.ToString("MMMM yyyy ");

                if (stmtnumber == null)
                    stmtnumber = masterRow[accountNoName].ToString();
                clsStatement_ExportRpt exportRptabp = new clsStatement_ExportRpt();
                if (createCorporateVal)
                    exportRptabp.whereCond = "cardaccountno = '" + stmtnumber + "' and packagename='STTOTABLECORP'";
                else
                    exportRptabp.whereCond = "accountno = '" + stmtnumber + "'";

                exportRptabp.whereCondD = "accountno = '" + stmtnumber + "'";

                //“Access Bank NGN Credit Card Statement July 2014 Temitayo Adedigba"
                //"Access Bank " + masterRow[mAccountcurrency].ToString() + "Credit Card Statement " + vCurDate.ToString("MMM yyyy") + masterRow[mCardclientname].ToString()
                //exportRptabp.exportAttachment(pStrFileName, pBankCode, masterRow[mCardclientname].ToString() + "_" + masterRow[mAccountcurrency].ToString() + "_" + vCurDate.ToString("MMyyyy"), pReportName);

                // If the bank confirmed

                //if (masterRow[mCardproduct].ToString().Contains("Gold"))
                //{
                //    string directory = Path.GetDirectoryName(pReportName);
                //    pReportName = directory + @"\Statement_ABP_GOLD_Credit.rpt";


                //}
                //else if (masterRow[mCardproduct].ToString().Contains("Platinum"))
                //{
                //    string directory = Path.GetDirectoryName(pReportName);
                //    pReportName = directory + @"\Statement_ABP_Platinum_Credit.rpt";
                //}
                //else
                //{
                //    string directory = Path.GetDirectoryName(pReportName);
                //    pReportName = directory + @"\Statement_ABP_Credit.rpt";
                //}
                
                exportRptabp.exportAttachment(pStrFileName, pBankCode, curFileName, pReportName);
                exportRptabp = null;
                stmtnumber = null;
                stmtfilename = @clsBasFile.getPathWithoutFile(pStrFileName) + @"\STMT\" + curFileName + ".pdf";
                stmtfilenameold = @clsBasFile.getPathWithoutFile(pStrFileName) + @"\" + curFileName + ".pdf";

                //stmtpass = masterRow[mExternalno].ToString() != "" ? masterRow[mExternalno].ToString() : masterRow[mAccountno].ToString() + "ACCESS";

                // DBN-21878: want to replace x|X
                if (frmStatementFile.Internal == true)
                {
                    stmtpass = clsCnfg.readSetting("InternalPassword");
                }
                else
                {
                    stmtpass = masterRow[mExternalno].ToString().Replace("X", "").Replace("x", "") != "" ? masterRow[mExternalno].ToString().Replace("X", "").Replace("x", "").Trim() : masterRow[mAccountno].ToString().Trim();
                }

                // Get a fresh copy of the sample PDF file
                string filenameSource = stmtfilenameold;
                string filenameDest = stmtfilenameold;
                //File.Copy(Path.Combine("../../../../../PDFs/", filenameSource),
                //  Path.Combine(Directory.GetCurrentDirectory(), filenameDest), true);
                // Open an existing document. Providing an unrequired password is ignored.
                PdfDocument document = PdfReader.Open(filenameDest);
                PdfSecuritySettings securitySettings = document.SecuritySettings;
                // Setting one of the passwords automatically sets the security level to 
                // PdfDocumentSecurityLevel.Encrypted128Bit.
                securitySettings.UserPassword = stmtpass;
                securitySettings.OwnerPassword = stmtpass;
                // Don't use 40 bit encryption unless needed for compatibility
                //securitySettings.DocumentSecurityLevel = PdfDocumentSecurityLevel.Encrypted40Bit;
                // Restrict some rights.
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
                //// ...and start a viewer.
                //Process.Start(filenameDest);
            } //end of Master foreach

            //emailLabelTmp = emailLabel;
            if (!string.IsNullOrEmpty(emailStr.ToString()) && !string.IsNullOrEmpty(emailTo)) //emailStr != string.Empty && emailStr != null  emailTo != string.Empty && emailTo != null
                SendEmail(emailStr.ToString(), "", emailTo);
            if (!string.IsNullOrEmpty(emailStr.ToString()) && string.IsNullOrEmpty(emailTo)) // emailStr != string.Empty && emailStr != null emailTo != string.Empty && emailTo != null
            {
                streamNoEmails.WriteLine(curCustomerName + "|" + curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|Without Email");
                noOfWithoutEmails++;
                emailTo = strEmailFrom; //"statement_Program@emp-group.com";
                //emailLabelTmp = emailLabel + " Acc:" + curAccountNumber + " Phone:" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!");
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
            zip.createZip(aryLstFilesStmt, @clsBasFile.getPathWithoutFile(pStrFileName) + "\\" + pBankName + pStrFile + pStmntType + "_STMT" + vCurDate.ToString("yyyyMM") + ".zip", "");

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
        emailStr.Append(@"<title>" + masterRow[mCustomername] + @"- Statement</title>");
        emailStr.Append(@"<style type=""text/css"">.newStyle1 {font-family: Arial, Helvetica, sans-serif;font-weight: bold;font-style: normal;}");
        emailStr.Append(@"#table25 {height: 43px;}");
        emailStr.Append(".newStyle2 {font-family: Arial, Helvetica, sans-serif;font-size: medium;font-style: normal;text-transform: uppercase;color: #FFFFFF;background-color: #E97A25;}");
        emailStr.Append(".auto-style1 {width: 371px;}.auto-style2 {width: 90px;}.auto-style3 {width: 375px;} .newStyle3 {font-family: Arial;font-size: 12pt;color: #000000;}table { empty-cells: show; }</style></head>");
        emailStr.Append(@"<body style=""background-image:url('cid:" + clsBasFile.getFileFromPath(strBackGround) + @"'); background-repeat: no-repeat; background-size: contain; background-position: center"">");
        emailStr.Append(@"<p align=""right"" style=""height: 38px"">&nbsp;</p>");
        emailStr.Append(@"<p align=""right"" style=""height: 62px""><a href='http://www.accessbankplc.com'><img border='0' src='cid:" + clsBasFile.getFileFromPath(strBankLogo) + "'></a></p>");
        emailStr.Append(@"<table border='0' width='100%' id='table1'><tr><td width='100%' border='0' class='newStyle2' colspan='2'>CREDIT CARD STATEMENT</td></tr>");
        emailStr.Append(@"<tr><td class='newStyle1' colspan='2'>" + masterRow[mCustomername] + @"</td></tr>");
        emailStr.Append(@"<tr><td colspan='2'>" + newaddress1 + " " + newaddress2 + @"</td></tr>");
        emailStr.Append(@"<tr><td colspan='2'>Billing Period: <b>" + Convert.ToDateTime(masterRow[mStatementdatefrom]).ToString("dd/MM/yyyy") + " - " + Convert.ToDateTime(masterRow[mStatementdateto]).ToString("dd/MM/yyyy") + @"</b><hr /></td></tr>");
        emailStr.Append(@"<tr><td colspan='4'><table border='0' width='100%' id='table6' style='table-layout: fixed;'>");
        emailStr.Append(@"<tr><td align='left' class='newStyle2' colspan='2'>ACCOUNT SUMMARY</td><td class='auto-style2'>&nbsp;</td><td align='left' class='newStyle2' colspan='2'>SUMMARY INFORMATION</td></tr>");
        emailStr.Append(@"<tr><td align='left'><font size='3'>Previous balance:</font><hr /></td>");
        emailStr.Append(@"<td align='right'>" + basText.formatNumUnSign(masterRow[mOpeningbalance], "##0.00", 16) + @"<hr /></td>");
        emailStr.Append(@"<td class='auto-style2'>&nbsp;</td><td align='left'><font size='3'>Minimum payment due:</font><hr /></td>");
        emailStr.Append(@"<td align='right'>" + basText.formatNum(masterRow[mMindueamount], "#0.00;(#0.00)", 16) + @"<hr /></td></tr>");
        emailStr.Append(@"<tr><td align='left'><font size='3'>Payments:</font><hr /></td>");
        emailStr.Append(@"<td align='right'>" + basText.formatNum(masterRow[mTotalpayments], "#0.00;(#0.00)", 16) + @"<hr /></td><td class='auto-style2'>&nbsp;</td><td align='left'><font size='3'>New Balance:</font><hr /></td>");
        emailStr.Append(@"<td align='right'>" + basText.formatNum(masterRow[mClosingbalance], "#0.00;(#0.00)", 16) + @"<hr /></td></tr><tr><td align='left'><font size='3'>Credit Limit:</font><hr /></td>");
        emailStr.Append(@"<td align='right'>" + basText.formatNum(masterRow[accountLimit], "#0.00;(#0.00)", 16) + @"<hr /></td><td class='auto-style2'>&nbsp;</td><td align='left'><font size='3'>Payment due date:</font><hr /></td>");
        emailStr.Append(@"<td align='right'>" + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy") + @"<hr /></td></tr><tr><td align='left'><font size='3'>Credits:</font><hr /></td>");
        emailStr.Append(@"<td align='right'>" + basText.formatNum(masterRow[mTotalcredits], "#0.00;(#0.00)", 16) + @"<hr /></td><td class='auto-style2'>&nbsp;</td><td align='left'><font size='3'>Card Number:</font><hr /></td>");
        emailStr.Append(@"<td align='right'>" + masterRow[mCardno].ToString().Substring(0, 6) + "******" + masterRow[mCardno].ToString().Substring(12, 4) + @"<hr /></td></tr><tr><td align='left'><font size='3'>Purchases:</font><hr /></td>");
        emailStr.Append(@"<td align='right'>" + basText.formatNum(masterRow[mTotalpurchases], "#0.00;(#0.00)", 16) + @"<hr /></td><td class='auto-style2'>&nbsp;</td><td align='left'><font size='3'>Account Number:</font><hr /></td>");
        emailStr.Append(@"<td align='right'>" + masterRow[mAccountno] + @"<hr /></td></tr><tr><td align='left'><font size='3'>Cash advances:</foCash advances:</font><hr /></td>");
        emailStr.Append(@"<td align='right'>" + basText.formatNum(masterRow[mTotalcashwithdrawal], "#0.00;(#0.00)", 16) + @"<font size='3'><hr /></td><td class='auto-style2'>&nbsp;</td><td align='left'><font size='3'>Card Type:</font><hr /></td><td align='right'>" + masterRow[mCardproduct] + @"<hr /></td></tr>");
        emailStr.Append(@"<tr><td align='left'><font size='3'>Interest:</font><hr /></td>");
        emailStr.Append(@"<td align='right'>" + basText.formatNum(masterRow[mTotalinterest], "#0.00;(#0.00)", 16) + @"<hr /></td></tr>");
        emailStr.Append(@"<tr><td align='left'><font size='3'>Available Credit Limit:</font></td>");
        emailStr.Append(@"<td align='right'>" + basText.formatNum(masterRow[accountAvailableLimit], "#0.00;(#0.00)", 16) + @"</td>");
        emailStr.Append(@"<td colspan='3'>&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td colspan=""5""><hr style=""height: 3px; border: none; color: #333; background-color: #828383;"" /></td></tr></table></td></tr>");
        emailStr.Append(@"<tr><td width='98%' colspan='2' height='20'>&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td align='left' class=' newStyle2' colspan='2'>UNDERSTANDING YOUR CREDIT CARD BETTER</td></tr><tr><td colspan='2'>&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td width='70%'><table style='height: 3px; border: solid; color: #E97A25;'><tr><td class='auto-style1'><p style='color: #E97A25; font-size: large; font-weight: bold;'>LATE REPAYMENT<br><p class='newStyle3'>");
        emailStr.Append(@"If your minimum payment/full payment is not<br>received in your card account at 1.00pm on the<br>due date (10th of every month) you will be liable<br>to pay a late repayment fee of $12.5</p></p></td></tr></table></td>");
        emailStr.Append(@"<td><table style='height: 3px; border: solid; color: #E97A25;'><tr><td width='30%'><p style='color: #E97A25; font-size: large; font-weight: bold;'>DUAL CURRENCY<br><p class='newStyle3'>");
        emailStr.Append(@"Your Access Visa Dual currency card allows you spend both locally and internationally. Whenever you make purchases locally, settlement will be done in Naira while purchases made internationally will be settled in Dollars.</p></p></td></tr></table></td></tr>");
        emailStr.Append(@"<tr><td width='98%' height='20' colspan='2'>&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td width='70%'><table style='height: 3px; border: solid; color: #E97A25;'><tr><td><p style='color: #E97A25; font-size: large; font-weight: bold;'>");
        emailStr.Append(@"INTEREST PAYMENT<p class='newStyle3'>Interest of 2.5% will be changed on the outstanding<br>balance on your card during each cycle.<br><br><br><br><br><br><br><br><br><br><br></p></p></td></tr>");
        emailStr.Append(@"<tr><td>&nbsp;</td></tr></table></td><td><table style='height: 3px; border: solid; color: #E97A25;'><tr><td width='30%'><p style='color: #E97A25; font-size: large; font-weight: bold;'> WAYS OF FUNDING YOUR CARD ACCOUNT<p class='newStyle3'>");
        emailStr.Append(@"1	You can send us a written Instruction or an email to debit your current/savings account with the naira equivalent of the dollar amount you wish to settle, at the bank’s competitive daily selling rate. This can come through your Relationship manager or through the Contact Center (details are below).<br><br>");
        emailStr.Append(@"2	You can log into the Online banking portal to make a transfer from your USD domiciliary account straight into your card account. Subject to availability of funds in your domiciliary account.<br><br>");
        emailStr.Append(@"3	You can make a dollar cash deposit into your credit card account at any of our branches.</p></td></tr></table></td></tr>");
        emailStr.Append(@"<tr><td colspan='2'><p style='color: #E97A25; font-size: large; font-weight: bold;'><br /><br />YOUR STATEMENT<p width='100%' border='0' class='newStyle2' colspan='1'>TRANSACTION DETAILS</p></p></td></tr></table>");
        emailStr.Append(@"<table border='0' width='100%' id='table2' class= ""table"" style='table-layout: fixed;'><tr class='newStyle2'><td>Post Date</td><td>ValueDate</td><td class='auto-style3'>Narration</td><td>Ref.</td><td>Debits</td><td>Credit</td><td>Balance</td></tr>");
        emailStr.Append(@"<tr><td>" + MakeHeaderStr(Convert.ToDateTime(masterRow[mStatementdatefrom]).ToString("dd/MM/yyyy")) + @"<hr /></td><td>" + MakeHeaderStr(Convert.ToDateTime(masterRow[mStatementdatefrom]).ToString("dd/MM/yyyy")) + @"<hr /></td><td>" + MakeHeaderStr("Balance brought forward") + "<hr /></td><td>&nbsp;<hr /></td><td>&nbsp;<hr /></td><td>&nbsp;<hr /></td><td>" + MakeHeaderStr(basText.formatNumUnSign(masterRow[mOpeningbalance], "##0.00", 16)) + @"<hr /></td></tr>");
        totalPages++;
        totNoOfPageStat++;
        TBalance = decimal.Parse(masterRow[mOpeningbalance].ToString());
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
        emailStr.Append(@"<title>" + curFileName + @"</title>");
        emailStr.Append(@"</head><body>");
        emailStr.Append(@"Dear " + masterRow[mCustomername].ToString() + "</br>");
        emailStr.Append(@"Please find attached your Account Statement for the month of " + vCurDate.ToString("MMMM") + ".</br>");
        emailStr.Append(@"The attached statement is a security enabled statement in PDF format which requires a password to open.</br>");
        emailStr.Append(@"Your password to view your account Statement is your card account number.</br>");
        emailStr.Append(@"Card account numbers can be obtained by calling the Contact Center on :+234 (01) 271 2005-7</br>");
        emailStr.Append(@"Thank you for choosing Access Bank.</br>");
        emailStr.Append(@"</body></html>");
        totalPages++;
        totNoOfPageStat++;
        //TBalance = decimal.Parse(masterRow[mOpeningbalance].ToString());
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

        emailStr.Append(@"<tr><td>" + MakeHeaderStr(postingDate.ToString("dd/MM")) + @"<hr /></td>");
        emailStr.Append(@"<td>" + MakeHeaderStr(trnsDate.ToString("dd/MM")) + @"<hr /></td>");
        emailStr.Append(@"<td>" + MakeHeaderStr(basText.trimStr(trnsDesc, 40)) + @"<hr /></td>");
        emailStr.Append(@"<td>" + MakeHeaderStr(basText.trimStr(detailRow[dRefereneno].ToString().Trim(), 24)) + @"<hr /></td>");

        if (clsBasValid.validateStr(detailRow[dBilltranamountsign]).ToString() == "DB")
        {
            emailStr.Append(@"<td style='color: red'>" + MakeHeaderStr(basText.formatNumUnSign(detailRow[dBilltranamount], "#,###,##0.00", 14)) + @"<hr /></td>");
            emailStr.Append(@"<td>&nbsp;<hr /></td>");
            TBilltranamount = decimal.Parse(detailRow[dBilltranamount].ToString()) * -1;
        }
        else if (clsBasValid.validateStr(detailRow[dBilltranamountsign]).ToString() == "CR")
        {
            emailStr.Append(@"<td>&nbsp;<hr /></td>");
            emailStr.Append(@"<td>" + MakeHeaderStr(basText.formatNumUnSign(detailRow[dBilltranamount], "#,###,##0.00", 14)) + @"<hr /></td>");
            TBilltranamount = decimal.Parse(detailRow[dBilltranamount].ToString());
        }
        TBalance += TBilltranamount;
        emailStr.Append(@"<td>" + MakeHeaderStr(basText.formatNumUnSign(TBalance, "##0.00", 16)) + @"<hr /></td></tr>");
        totNoOfTransactions++;
    }

    protected void printAccountFooter()
    {
        emailStr.Append(@"<tr><td>&nbsp;</td>");
        emailStr.Append(@"<td>&nbsp;</td>");
        emailStr.Append(@"<td style='font-weight: bold;' class='auto-style3'>Closing Balance</td>");
        emailStr.Append(@"<td>&nbsp;</td>");
        emailStr.Append(@"<td>&nbsp;</td>");
        emailStr.Append(@"<td>&nbsp;</td>");
        emailStr.Append(@"<td style='font-weight: bold;'>" + basText.formatNumUnSign(masterRow[mClosingbalance], "#,###,##0.00", 16) + @"</td></tr>");
        emailStr.Append(@"<tr><td colspan=""7""><hr style=""height: 3px; border: none; color: #333; background-color: #828383;"" /></td></tr></table></td></tr>");
        emailStr.Append(@"</table><p align='right'><br /><br /><br /><img border='0' src='cid:" + clsBasFile.getFileFromPath(strbottomBanner) + "'/></p>");
        emailStr.Append(@"<p>Please advise Access Bank of any discrepancy noted in this statement as soon as possible via any of the contact details advised below.<br>");
        emailStr.Append(@"All products are subject to Bank terms and conditions.<br><br />Phone number: +234 (01) 271 2005-7<br />E-mail: contactcenter@accessbankplc.com</p>");
        emailStr.Append(@"<p align='left'><img border='0' src='cid:" + clsBasFile.getFileFromPath(strbottomBanner1) + "'/></p></body></html>");
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
                streamNoEmails.WriteLine(curCustomerName + "|" + curAccountNumber + "|" + curClientID + "|" + pTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|" + clsCnfg.readSetting("strValidEmail") ?? "Bad Email");
                noOfBadEmails++;
                pTo = strEmailFrom; //"statement_Program@emp-group.com";
                //emailLabelTmp = emailLabel + " Acc:" + curAccountNumber + " Phone:" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!");
            }

            //try
            //    {

            //streamEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
            ArrayList pLstTo = new ArrayList(), pLstAttachedFile = new ArrayList(), pLstCC = new ArrayList(), pLstAttachedPic = new ArrayList();//, pLstCC = new ArrayList(), pLstBCC = new ArrayList();
            pSubject = emailLabelTmp; //emailLabel "BAI statement for 02/2008";
            //pTo = emailTo;// "mmohammed@emp-group.com";
            if (pTo.EndsWith("."))//ehimolen@yahoo.com
                pTo = pTo + "com";
            pLstTo.Add(pTo);//"mmohammed@emp-group.com" "mhrap@hotmail.com""mhrap@yahoo.com"    "jedkosi@gmail.com"   "nazab@emp-group.com"

            //pLstAttachedPic.Add(strBankLogo);//strOutputFile@"D:\Web\Email\BAI\Logo.gif"
            //pLstAttachedPic.Add(strbottomBanner);
            //pLstAttachedPic.Add(strbottomBanner1);
            //pLstAttachedPic.Add(strBackGround);

            if (HasAttachement)
            {
                //foreach (string fileName in Directory.GetFiles(attachedFilesStr))
                //    {
                //    if (clsBasFile.getFileExtn(fileName) == "pdf")
                //        {
                pLstAttachedFile.Add(stmtfilenameold);
                //break;
                //}
                //}
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

            if (strEmailFromTmp.ToUpper().EndsWith("ACCESSBANKPLC.COM"))
            {
                if (pLstCC.Contains("statement@emp-group.com") == false)
                {
                    pLstCC.Add("creditcardsettlement@accessbankplc.com");
                    //pLstCc.Add("mamr@network.com.eg");
                }
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

            if (HasAttachement)
            {
                //foreach (string fileName in Directory.GetFiles(attachedFilesStr))
                //{
                //if (clsBasFile.getFileExtn(fileName) == "pdf")
                //{
                clsBasFile.moveFile(stmtfilenameold, stmtfilename);
                aryLstFilesStmt.Add(stmtfilename);
                //break;
                //}
                //}
            }

            //sndMail = null;
            //emailStr = new StringBuilder(""); emailTo = string.Empty; curAccountNumber = string.Empty; curCardNumber = string.Empty; curClientID = string.Empty;
            //wait2NextEmail(prvEmail, pTo, waitPeriodVal);//     System.Threading.Thread.Sleep(waitPeriodVal);//400
            //prvEmail = pTo;
        }
        catch (Exception ex)   //(NotSupportedException ex) (Exception ex)  //
        {
            //clsBasErrors.catchError(ex);
            streamNoEmails.WriteLine(curCustomerName + "|" + curAccountNumber + "|" + curClientID + "|" + pTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "| Email " + pTo + " , Err Message >> " + ex.Message);
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

    public string VIPCondition
    {
        get { return vipCondVal; }
        set { vipCondVal = value; }
    }// VIPCondition

    public string ProductCondition
    {
        get { return productCond; }
        set { productCond = value; }
    }// ProductCondition


    ~clsStatHtmlDBN()
    {
        DSstatement.Dispose();
    }
}
