using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Xml;

using System.Net;
using System.Net.Sockets;

using System.Text.RegularExpressions;
using System.Collections;


public class clsValidateBankEmail : clsBasStatement
    {
    private string strBankName;
    private FileStream fileEmails, fileBadEmails, fileNoEmails, fileSummary;//fileStrm, 
    private StreamWriter streamEmails, streamBadEmails, streamNoEmails, streamSummary;//streamWrit, 
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

    //
    //private string emailStr;
    private StringBuilder emailStr = new StringBuilder("");
    private string emailLabel, strBankLogo;
    private string strEmailFrom = "statement@emp-group.com";  //"cardservices@zenithbank.com"
    private string strbankWebLink = "www.emp-group.com";  //www.zenithbank.com
    private string emailTo, curAccountNumber, curCardNumber, curClientID;//, curEmail
    private string strEmailFileName, strBadEmailFileName, strNoEmailFileName;
    private frmStatementFile frmMain;
    private int totRec = 1, numMail, numOfTry = 0;
    private ArrayList pLstCC = new ArrayList(), pLstBCC = new ArrayList();
    private string logoAlignmentStr = "center";
    private string strbackGround = @"D:\pC#\ProjData\Statement\_Background\Background05.jpg";
    private bool isRewardVal = false;
    private bool createCorporateVal = false;
    private string accountNoName = mAccountno;
    private string accountLimit = mAccountlim;
    private string accountAvailableLimit = mAccountavailablelim;
    private string rewardCondVal = "'New Reward Contract'";//'Reward Contract'
    private clsEmail sndMail;
    private string emailFromNameStr;
    private string basPath;
    private int noOfBadEmails = 0, noOfNoEmails = 0, noOfValidEmails = 0;

    clsAes cryptAes = new clsAes();

    public clsValidateBankEmail()
        {
        }

    public string Statement(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pCurDate, string pStmntType, bool pAppendData)
        {
        string rtrnStr = "Successfully Generate " + pBankName;
        int curMonth = pCurDate.Month;

        basPath = pStrFileName;
        if (!pStmntType.Contains("ClientsEmails"))
            {
            MessageBox.Show("You must select subscribed banks only!", "Validate E-Mails", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1);
            return rtrnStr = "You must select subscribed banks only!";
            }
        pStmntType = "ValidateEmail";
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
            strEmailFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "_" + pStmntType + "\\" + pBankName + "_" + "ValidEmail" + "_" + vCurDate.ToString("yyyyMM");//+ ".txt"
            strBadEmailFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "_" + pStmntType + "\\" + pBankName + "_" + "BadEmails" + "_" + vCurDate.ToString("yyyyMM");//+ ".txt"
            strNoEmailFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "_" + pStmntType + "\\" + pBankName + "_" + "NoEmails" + "_" + vCurDate.ToString("yyyyMM");//+ ".txt"
            pStrFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "_" + pStmntType + "\\" + pBankName + "_" + pStmntType + "_" + vCurDate.ToString("yyyyMM") + ".txt";
            strBankName = pBankName;
            emailLabel = pBankName + " statement for " + vCurDate.ToString("MM/yyyy"); //"BAI statement for 02/2008"
            strOutputFile = pStrFileName;

            // open emails file
            fileEmails = new FileStream(strEmailFileName + ".txt", FileMode.Create); //Create
            streamEmails = new StreamWriter(fileEmails, Encoding.Default);
            streamEmails.WriteLine("Name" + "|" + "ClientID" + "|" + "Email" + "|" + "Mobile Phone");
            streamEmails.AutoFlush = true;

            // open Bad emails file
            fileBadEmails = new FileStream(strBadEmailFileName + ".txt", FileMode.Create); //Create
            streamBadEmails = new StreamWriter(fileBadEmails, Encoding.Default);
            streamBadEmails.WriteLine("Name" + "|" + "ClientID" + "|" + "Email" + "|" + "Mobile Phone");
            streamBadEmails.AutoFlush = true;

            // open No emails file
            fileNoEmails = new FileStream(strNoEmailFileName + ".txt", FileMode.Create); //Create
            streamNoEmails = new StreamWriter(fileNoEmails, Encoding.Default);
            streamNoEmails.WriteLine("Name" + "|" + "ClientID" + "|" + "Mobile Phone");
            streamNoEmails.AutoFlush = true;

            // open Summary file
            fileSummaryName = pStrFileName;
            fileSummaryName = clsBasFile.getPathWithoutExtn(fileSummaryName) +
              "_Summary." + clsBasFile.getFileExtn(fileSummaryName);
            fileSummary = new FileStream(fileSummaryName, FileMode.Create); //Create
            streamSummary = new StreamWriter(fileSummary, Encoding.Default);
            streamEmails.AutoFlush = true;

            // set branch for data
            curBranchVal = pBankCode; // 10; //3 = real   1 = test

            getClientEmailName(pBankCode);

            clsValidateEmail valdEmail = new clsValidateEmail();
            string rtrnMessage = string.Empty;
            frmMain.BeginInvoke(frmMain.setMinMaxProgressDelegate, new object[] { DSemails.Tables["Emails"].Rows.Count });
            foreach (DataRow emailRow in DSemails.Tables["Emails"].Rows)
            {
                frmMain.BeginInvoke(frmMain.setProgressDelegate, new object[] { totRec++ });
                if (emailRow != null)
                {
                    emailTo = emailRow["email"].ToString().Trim();
                    if (emailTo.Length > 0)
                    {
                        rtrnMessage = valdEmail.isValideEmail(emailTo);
                        if (rtrnMessage == "ValidEmail")
                        {
                            streamEmails.WriteLine((emailRow?["fio"]?.ToString() ?? "no fio!") + "|" + (emailRow?["idclient"]?.ToString() ?? "no client id!") + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!"));
                            noOfValidEmails++;
                        }
                        else
                        {
                            streamBadEmails.WriteLine((emailRow?["fio"]?.ToString() ?? "no fio!") + "|" + (emailRow?["idclient"]?.ToString() ?? "no client id!") + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|" + rtrnMessage);
                            noOfBadEmails++;
                        }
                    }// if
                    else
                    {
                        streamNoEmails.WriteLine((emailRow?["fio"]?.ToString() ?? "no fio!") + "|" + (emailRow?["idclient"]?.ToString() ?? "no client id!") + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!"));
                        noOfNoEmails++;
                    }
                }
            }// foreach
            valdEmail = null;


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

            streamBadEmails.Flush();
            streamBadEmails.Close();
            fileBadEmails.Close();

            streamNoEmails.Flush();
            streamNoEmails.Close();
            fileNoEmails.Close();
            aryLstFiles.Add(strEmailFileName + ".txt");

            if (noOfBadEmails == 0)
                clsBasFile.deleteFile(strBadEmailFileName + ".txt");
            else
                aryLstFiles.Add(strBadEmailFileName + ".txt");

            aryLstFiles.Add(strNoEmailFileName + ".txt");
            aryLstFiles.Add(fileSummaryName);
            //clsBasFile.generateFileMD5(aryLstFiles, strEmailFileName + ".MD5");
            clsBasFile.generateFileMD5(aryLstFiles, clsBasFile.getPathWithoutExtn(pStrFileName) + ".MD5");
            aryLstFiles.Add(strEmailFileName + ".MD5");
            SharpZip zip = new SharpZip();
            zip.createZip(aryLstFiles, clsBasFile.getPathWithoutExtn(pStrFileName) + ".zip", "");
            if (noOfBadEmails > 0)
                {
                //if (DialogResult.Yes == MessageBox.Show("Are you sure you want send bank mail for " + pBankName + ".", "Send bank Mail",
                //   MessageBoxButtons.YesNo, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button2))
                    {
                    //string strDay = "01";
                    //string strMonth = vCurDate.ToString("MM"), strYear = vCurDate.ToString("yyyy");
                    //if (vCurDate.AddDays(18).Month != vCurDate.Month)
                    //  strDay = "15";
                    ////periodDate = 
                    //string strBasPath = clsBasFile.makeStrAsPath(basPath);
                    //string sourceDir = strBasPath + strYear + strMonth + strBankName + "_" + pStmntType;
                    //string destineDir = strBasPath + "_" + strYear + strMonth + "F\\" + vCurDate.AddMonths(1).ToString("yyyy") + vCurDate.AddMonths(1).ToString("MM") + strDay + "\\" + strYear + strMonth + strBankName + "_ValidateEmail";
                    //clsBasFile.moveDirectory(sourceDir, destineDir);

                    clsMailStatement sendmail = new clsMailStatement();
                    sendmail.sendBankMail(pBankName, pCurDate, clsBasFile.getPathWithoutExtn(pStrFileName) + ".zip", "_ValidateEmail", clsSessionValues.basPath); //
                    clsFtpCommandLine.sendFile2ftpSilent(clsBasFile.getPathWithoutExtn(pStrFileName) + ".zip", "Opsdev/_Settlement/_" + pCurDate.Year);
                    }
                }
            }
        return rtrnStr;
        }

    private void add2FileList(string pFileName)
        {
        int myIndex = aryLstFiles.BinarySearch((object)pFileName);
        if (myIndex < 0)
            aryLstFiles.Add(@pFileName);
        }

    private void printStatementSummary()
        {
        streamSummary.WriteLine(strBankName + " Bank Email Validation Summary");
        streamSummary.WriteLine("__________________________");
        streamSummary.WriteLine("");
        streamSummary.WriteLine("Number of cardholder(s) has Valid Email   " + noOfValidEmails.ToString());
        streamSummary.WriteLine("Number of cardholder(s) has bad Email     " + noOfBadEmails.ToString());
        streamSummary.WriteLine("Number of cardholder(s) without Email     " + noOfNoEmails.ToString());
        //streamSummary.WriteLine("No of Transactions " + totNoOfTransactions.ToString());

        }


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

    ~clsValidateBankEmail()
        {
        }
    }
