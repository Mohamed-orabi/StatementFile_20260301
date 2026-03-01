using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Xml;
using System.Collections;

public class clsStatTxtLbl_Suez : clsBasStatement
    {
    protected string strBankName;
    protected FileStream fileStrm, fileSummary;
    protected StreamWriter streamWrit, streamSummary;
    protected DataRow masterRow;
    protected DataRow detailRow;
    protected string strEndOfLine = "\u000D";  //+ "M" ^M
    protected string strEndOfPage = "\u000C";  //  + basText.replicat(" ",85) + "F 2"  + "\u000D"+ "M" ^L + ^M
    protected string strEndOfAccount = "----------------- END OF STATEMENT -----------------";  //  + basText.replicat(" ",85) + "F 2"  + "\u000D"+ "M" ^L + ^M
    protected const int MaxDetailInPage = 25; //
    protected const int MaxDetailInLastPage = 25; //
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
    protected int totNoOfCardStat, totNoOfPageStat, totNoOfCardEmailStat, totNoOfPageEmailStat;

    protected int curCardRow = 0, curTotCardRows = 0, numOfErr = 0, totNoOfTransactions = 0, totNoOfTransEmail = 0, totNoOfTransactionsInt = 0, totNoOfTransEmailInt = 0;
    protected bool isPrimaryOnly, isHaveF3 = true, isPrimaryOpened = false;
    protected FileStream fileStrmErr;
    protected StreamWriter strmWriteErr;
    protected string curMainCard;

    protected string extAccNum;
    protected int totCrdNoInAcc, curCrdNoInAcc, totSuppCrdsInAcc;
    protected string strOutputPath, strOutputFile, fileSummaryName;
    protected DateTime vCurDate;
    protected frmStatementFile frmMain;
    protected int totRec = 1;
    protected bool createCorporateVal = false;
    protected string accountNoName = mAccountno;
    protected string accountLimit = mAccountlim;
    protected string accountAvailableLimit = mAccountavailablelim;
    protected string rewardCond = "'New Reward Contract'";
    protected bool isReward = false;
    protected bool isDebit = false;
    protected bool isFilteredVal = false;
    protected DataRow[] emailRows = null;
    protected string emailTo = string.Empty;
    protected bool isInEmailService = false;
    protected bool isEmailStat = false;
    protected bool isSupp = false;
    clsValidateEmail valdEmail = new clsValidateEmail();
    protected string strFileNam, stmntType;
    protected byte NoLinePerPage = 53;
    protected ArrayList aryLstFiles = new ArrayList();
    protected string commnFileName;
    protected string add2FileName = string.Empty;
    protected bool isCorporateGenrated = false, isNotSplit = true, isAutoinalize = true;
    private DataRow ProductRow;
    protected bool hasInterset = false;
    string cContract = string.Empty, curFileName = string.Empty;
    protected string CustomerArabicname;
    private string curSuppCard = string.Empty;
    protected DataRow[] tmpDtlRows;
    protected int pBranch = 0;
    private string strProductCond = string.Empty;
    private string strExcludeCond = string.Empty;
    protected bool isExcludedVal = false;


    public clsStatTxtLbl_Suez()
        {
        }

    public string Statement(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pCurDate, string pStmntType, bool pAppendData, string add2FileName)
        {
        string rtrnStr = "Successfully Generate " + pBankName, ppStrFileName = pStrFileName;
        int curMonth = pCurDate.Month;
        string comnFileName = string.Empty;
        bool preExit = true;
        strFileNam = pStrFileName;
        stmntType = pStmntType;
        pBranch = pBankCode;
        if (isReward)
            NoLinePerPage = 57;

        try
            {
            clsMaintainData maintainData = new clsMaintainData();
            maintainData.matchCardBranch4Account(pBankCode);

            pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);
            vCurDate = pCurDate; //DateTime.Now.AddMonths(-1);
            strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
            clsBasFile.createDirectory(pStrFileName + vCurDate.ToString("yyyyMM") + pBankName);
            pStrFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + add2FileName + "_" + vCurDate.ToString("yyyyMM") + ".txt";
            strBankName = pBankName;
            comnFileName = strOutputFile = pStrFileName;

            // open output file
            fileStrm = new FileStream(comnFileName, FileMode.Create); //Create
            streamWrit = new StreamWriter(fileStrm, Encoding.Default);
            streamWrit.AutoFlush = true;

            // Error file
            fileStrmErr = new FileStream(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName), FileMode.Create);
            strmWriteErr = new StreamWriter(fileStrmErr, Encoding.Default);

            // open Summary file
            fileSummaryName = pStrFileName;
            fileSummaryName = clsBasFile.getPathWithoutExtn(fileSummaryName) +
              "_Summary." + clsBasFile.getFileExtn(fileSummaryName);
            fileSummary = new FileStream(fileSummaryName, FileMode.Create); //Create
            streamSummary = new StreamWriter(fileSummary, Encoding.Default);


            // set branch for data
            curBranchVal = pBankCode; // 16; //3 = real   1 = test
            // data retrieve
            MainTableCond = " m.contracttype = " + strProductCond + "";//strWhereCond
            supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype = " + strProductCond + ")";
            strOrder = " m.CARDBRANCHPART,m.accountno,m.cardprimary desc, m.cardno";
            FillStatementDataSet(pBankCode, "");
            getCardProduct(pBankCode);
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
                cardsRows = mRow.GetChildRows(StatementNoDRel); //, DataRowVersion.Proposed //EDT-1249
                //cardsRows = DSstatement.Tables["tStatementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[accountNoName]).ToString().Trim() + "'");

                strCardNo = masterRow[mCardno].ToString().Trim();
                if (masterRow[mHOLSTMT].ToString() == "Y")
                    continue;
                
                if (strCardNo.Length != 16)
                {
                    strmWriteErr.WriteLine("CardNo Length not equal 16 for AccountNo " + masterRow[accountNoName].ToString().Trim() + " Card No =" + strCardNo);
                    //numOfErr++;
                    continue;// Exclude Zero Length Cards 
                }
                
                if ((clsCheckCard.CalcLuhnCheckDigit(strCardNo.Substring(0, 15)) != strCardNo.Substring(15, 1))) //!clsCheckCard.isValidCard(strCardNo) ||
                {
                    strmWriteErr.WriteLine("Invalid Card Number" + strCardNo);
                    numOfErr++;
                }


                strPrimaryCardNo = strCardNo;
                if (masterRow[mCardprimary].ToString() == "N")
                    {
                    strPrimaryCardNo = masterRow[mPrinarycardno].ToString();
                    }

                //start new account
                if (prevAccountNo != masterRow[accountNoName].ToString())
                    {
                    if (prevAccountNo == string.Empty)
                        if (curMonth != Convert.ToDateTime(masterRow[mStatementdateto]).Month - clsSessionValues.statGenAfterMonth)//,"dd/MM/yyyy"
                            {
                            preExit = false;
                            fileStrm.Close();
                            fileStrmErr.Close();
                            clsBasFile.deleteFile(@strOutputFile);
                            clsBasFile.deleteFile(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName));
                            return "Error in Generation " + pBankName;
                            }
                    if (pageNo != totalAccPages && prevAccountNo != "")// 
                        {
                        strmWriteErr.WriteLine("pageNo not equal totalAccPages : " + prevAccountNo);
                        numOfErr++;
                        }

                    curMainCard = string.Empty;
                    if (!isHaveF3)//!isHaveF3  CurrentPageFlag != "F 2"
                        {
                        strmWriteErr.WriteLine("Account not Closed : " + prevAccountNo);
                        numOfErr++;
                        }
                    isHaveF3 = false;

                    pageNo = 1;//totalAccPages = 1 ; pageNo=1;
                    CurPageRec4Dtl = 0; totNetUsage = 0; CurrentPageFlag = "F 1"; //if page is based on account no
                    calcAccountRows();

                    if (!isDebit)
                        if (totAccRows < 2 && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) >= 0) //             || (masterRow[mCardno].ToString() == curMainCard   // Convert.ToDecimal(
                            {
                            isHaveF3 = true;
                            //pageNo=1; totalAccPages =1;
                            continue;
                            }

                    prevAccountNo = masterRow[accountNoName].ToString();
                    //pageNo=1; //if page is based on account no
                    printHeader();//if page is based on account no

                    if (isEmailStat)
                        totNoOfCardEmailStat++;
                    else
                        totNoOfCardStat++;

                    if (Convert.ToDecimal(masterRow[mAccountlim].ToString()) <= Convert.ToDecimal(0))
                        {
                        strmWriteErr.WriteLine("Account Limit Less than or Equal Zero for Account " + masterRow[mAccountno].ToString());// + " and Card Number " + strCardNo
                        numOfErr++;
                        }
                    //if (clsStatementSummary.isUpdateDatble == true)
                    //    {
                    //    GenerateStatementLog();
                    //    }
                    } // End of if(prevAccountNo != masterRow[accountNoName].ToString())
                //calcCardlRows();
                //if(totCardRows < 1)continue ;  //if pages is based on card
                //printHeader();//if pages is based on card

                foreach (DataRow dRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
                    {
                    detailRow = dRow;
                    stmNo = detailRow[dStatementno].ToString();
                    if ((detailRow[dPostingdate] == DBNull.Value) && (detailRow[dDocno] == DBNull.Value)) continue;// Exclude On-Hold Transactions 
                    curAccRows++;
                    if (CurPageRec4Dtl >= MaxDetailInPage)
                        {
                        CurPageRec4Dtl = 0;
                        printCardFooter();
                        pageNo++;
                        printHeader();
                        }
                    //						totNetUsage +=clsBasValid.validateNum(detailRow[dBilltranamount]);
                    totNetUsage = calculateCrDb(totNetUsage, clsBasValid.validateNum(detailRow[dBilltranamount]), clsBasValid.validateStr(detailRow[dBilltranamountsign]));
                    CurPageRec4Dtl++;

                    if (curSuppCard != detailRow[dCardno].ToString())
                        {
                        tmpDtlRows = DSstatement.Tables["tStatementMasterTable"].Select(mCardno + " = '" + detailRow[dCardno] + "'");
                        if (tmpDtlRows.Length > 0)
                            {
                            if (tmpDtlRows[0][mCardprimary].ToString() == "N" && isValidateCard(tmpDtlRows[0][mCardstate].ToString()))
                                {
                                streamWrit.WriteLine(basText.replicats(" ", 23) + ">> Supplementary Card : " + basText.formatCardNumber(detailRow[dCardno].ToString()));
                                CurPageRec4Dtl++;
                                curCrdNoInAcc++;
                                curAccRows++;
                                //BMSR-1620
                                //if (CurPageRec4Dtl >= MaxDetailInPage)
                                if (CurPageRec4Dtl > MaxDetailInPage)
                                    {
                                    CurPageRec4Dtl = 1;
                                    printCardFooter();
                                    pageNo++;
                                    printHeader();
                                    }
                                }
                            }
                        }
                    curSuppCard = detailRow[dCardno].ToString();

                    printDetail();

                    } //end of detail foreach
                curCrdNoInAcc++;

                if ((curAccRows >= totAccRows && totAccRows != 0) || (totAccRows == 0 && curCrdNoInAcc == totCrdNoInAcc))
                    {
                    completePageDetailRecords();
                    printCardFooter();//if pages is based on account
                    CurPageRec4Dtl = 0; //>pageNo=1; if pages is based on account
                    curAccRows = 0;
                    }


                } //end of Master foreach
            streamWrit.WriteLine(strEndOfPage + strEndOfAccount);

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
            if (preExit)
                {
                // Close output File
                streamWrit.Flush();
                streamWrit.Close();
                fileStrm.Close();

                printStatementSummary();

                // Close Summary File
                streamSummary.Flush();
                streamSummary.Close();
                fileSummary.Close();

                strmWriteErr.Flush();
                strmWriteErr.Close();
                fileStrmErr.Close();
                //if (numOfErr == 0)
                //    clsBasFile.deleteFile(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName));
                aryLstFiles.Add(comnFileName);
                numOfErr = validateNoOfLines(aryLstFiles, NoLinePerPage);//53
                if (numOfErr == 0)
                    clsBasFile.deleteFile(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName));
                else
                    {
                    aryLstFiles.Add(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err.txt");
                    rtrnStr += string.Format(" with {0} Errors", numOfErr);
                    }
                aryLstFiles.Add(@fileSummaryName);
                }
            }
        return rtrnStr;
        }

    public string SplitByContractBranch(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pCurDate, string pStmntType, bool pAppendData)
        {
        string rtrnStr = "Successfully Generate " + pBankName;
        string queryString;
        queryString = "select distinct t.contracttype, t.cardbranchpart, t.cardbranchpartname from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " t where t.branch = " + pBankCode + "and t.cardbranchpart is not null order by 1,2";
        try
            {
            clsMaintainData maintainData = new clsMaintainData();
            maintainData.notRward = false;
            maintainData.matchCardBranch4Account(pBankCode);
            commnFileName = pStrFileName + pCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + pCurDate.ToString("yyyyMM") + ".txt";
            OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
            OracleCommand cmd = new OracleCommand(queryString, conn);
            conn.Open();
            OracleDataReader rdr;
            rdr = cmd.ExecuteReader();
            while (rdr.Read())
                {
                masterQuery = detailQuery = string.Empty;
                mainTableCond = " contracttype = '" + rdr[0].ToString() + "' and cardbranchpart = '" + rdr[1].ToString() + "'";
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and " + mainTableCond + ")";
                Statement(pStrFileName, pBankName, pBankCode, pStrFile, pCurDate, pStmntType, pAppendData, rdr[0].ToString() + "_" + rdr[2].ToString());
                DSstatement.Clear();
                StatementNoDRel = null;
                mainTableCond = "";
                supTableCond = "";
                totRec = 1;
                }
            rdr.Close();
            clsBasFile.generateFileMD5(aryLstFiles, @clsBasFile.getPathWithoutExtn(commnFileName) + ".MD5");
            aryLstFiles.Add(@clsBasFile.getPathWithoutExtn(commnFileName) + ".MD5");
            SharpZip zip = new SharpZip();
            zip.createZip(aryLstFiles, @clsBasFile.getPathWithoutExtn(commnFileName) + ".zip", "");


            DSstatement.Dispose();
            }
        catch (OracleException ex)
            {
            clsDbOracleLayer.catchError(ex);
            rtrnStr = "Error Generate " + pBankName;
            }
        catch (NotSupportedException ex)  //(Exception ex)  //
            {
            clsBasErrors.catchError(ex);
            rtrnStr = "Error Generate " + pBankName;
            }
        finally
            {
            }
        return rtrnStr;
        }

    public string SplitByBranch(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pCurDate, string pStmntType, bool pAppendData)
        {
        string rtrnStr = "Successfully Generate " + pBankName;
        string queryString;
        queryString = "select distinct t.cardbranchpart, t.cardbranchpartname from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " t where t.branch = " + pBankCode + " order by 1,2";
        try
            {
            clsMaintainData maintainData = new clsMaintainData();
            maintainData.notRward = false;
            maintainData.matchCardBranch4Account(pBankCode);
            commnFileName = pStrFileName + pCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + pCurDate.ToString("yyyyMM") + ".txt";
            OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
            OracleCommand cmd = new OracleCommand(queryString, conn);
            conn.Open();
            OracleDataReader rdr;
            rdr = cmd.ExecuteReader();
            while (rdr.Read())
                {
                masterQuery = detailQuery = string.Empty;
                mainTableCond = " cardbranchpart = " + rdr[0].ToString() + "";
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and " + mainTableCond + ")";
                Statement(pStrFileName, pBankName, pBankCode, pStrFile, pCurDate, pStmntType, pAppendData, rdr[1].ToString());
                DSstatement.Clear();
                StatementNoDRel = null;
                mainTableCond = "";
                supTableCond = "";
                totRec = 1;
                }
            rdr.Close();
            clsBasFile.generateFileMD5(aryLstFiles, @clsBasFile.getPathWithoutExtn(commnFileName) + ".MD5");
            aryLstFiles.Add(@clsBasFile.getPathWithoutExtn(commnFileName) + ".MD5");
            SharpZip zip = new SharpZip();
            zip.createZip(aryLstFiles, @clsBasFile.getPathWithoutExtn(commnFileName) + ".zip", "");


            DSstatement.Dispose();
            }
        catch (OracleException ex)
            {
            clsDbOracleLayer.catchError(ex);
            rtrnStr = "Error Generate " + pBankName;
            }
        catch (NotSupportedException ex)  //(Exception ex)  //
            {
            clsBasErrors.catchError(ex);
            rtrnStr = "Error Generate " + pBankName;
            }
        finally
            {
            }
        return rtrnStr;
        }

    protected void printHeader()
        {
        string newaddress1, newaddress2;
        clsMaintainData.fixAddress(curBranchVal, masterRow[mCustomeraddress1].ToString(), out newaddress1, out newaddress2);
        extAccNum = clsBasValid.validateStr(masterRow[mExternalno]);
        if (extAccNum.Trim() == "")
            extAccNum = clsBasValid.validateStr(masterRow[mAccountno]);

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
            }

        if (CurrentPageFlag == "F 0" || CurrentPageFlag == "F 1")
            {
            streamWrit.WriteLine(strEndOfPage + strEndOfAccount);
            }
        else
            {
            streamWrit.WriteLine(strEndOfPage);
            }
        streamWrit.WriteLine(String.Empty); //(basText.replicat(" ", 81) + CurrentPageFlag);  //+"{0,5:MM/yy}",masterRow[mStatementdatefrom]
        streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(String.Empty);
        //		streamWrit.WriteLine(basText.replicat(" ",81) + masterRow[mCardproduct]);  //+"{0,5:MM/yy}",masterRow[mStatementdatefrom]
        streamWrit.WriteLine(basText.alignmentMiddle("Customer Name & Address", 50) + basText.replicat(" ", 5) + basText.alignmentRight("Card Product :", 25) + " " + masterRow[mCardproduct].ToString());  //+"{0,5:MM/yy}",masterRow[mStatementdatefrom]
        streamWrit.WriteLine(basText.alignmentMiddle(masterRow[mCustomername], 50));  //
        //		streamWrit.WriteLine( basText.replicat(" ",81) + masterRow[mCardbranchpartname]);  //
        // Adding Client ID -- Jira SUEZ-4425 START
        //streamWrit.WriteLine(basText.replicat(" ", 55) + basText.alignmentRight("Branch :", 25) + " " + masterRow[mCardbranchpartname].ToString());  //
        streamWrit.WriteLine(basText.alignmentMiddle(masterRow[mClientid], 54) + basText.alignmentRight("Branch :", 25) + " " + masterRow[mCardbranchpartname].ToString());  //
        // Adding Client ID -- Jira SUEZ-4425 END
        //streamWrit.WriteLine(basText.alignmentMiddle(ValidateArbic(masterRow[mCustomeraddress1].ToString()), 50)); //basText.formatNum(masterRow[mCardno],"####-####-####-####")
        streamWrit.WriteLine(basText.alignmentMiddle(ValidateArbic(newaddress1), 50)); //basText.formatNum(masterRow[mCardno],"####-####-####-####")
        //		streamWrit.WriteLine(basText.replicat(" ",81) + extAccNum); //clsBasValid.validateStr(masterRow[mExternalno]) //accountno  //basText.formatNum(masterRow[mCardno],"####-####-####-####")
        streamWrit.WriteLine(basText.replicat(" ", 55) + basText.alignmentRight("Account Number :", 25) + " " + extAccNum); //clsBasValid.validateStr(masterRow[mExternalno]) //accountno  //basText.formatNum(masterRow[mCardno],"####-####-####-####")
        //streamWrit.WriteLine(basText.alignmentMiddle(ValidateArbic(masterRow[mCustomeraddress2].ToString()), 50));
        streamWrit.WriteLine(basText.alignmentMiddle(ValidateArbic(newaddress2), 50));
        streamWrit.WriteLine(basText.replicat(" ", 55) + basText.alignmentRight("Statement Date :", 25) + " " + "{0,10:dd/MM/yyyy}", masterRow[mStatementdateto]);//Statementdatefrom+ basText.formatNum(masterRow[accountLimit],"########")  basText.formatNum(masterRow[mCardlimit],"########")
        streamWrit.WriteLine(basText.alignmentMiddle(ValidateArbic(masterRow[mCustomeraddress3].ToString().Trim()) + " " + masterRow[mCustomerregion].ToString().Trim() + " " + masterRow[mCustomercity].ToString().Trim(), 50));  //basText.formatNum(masterRow[accountNoName],"##-##-######-####")
        streamWrit.WriteLine(basText.replicat(" ", 55) + basText.alignmentRight("Page :", 25) + " " + pageNo + " / " + totalAccPages);  //basText.formatNum(masterRow[accountNoName],"##-##-######-####")
        if (createCorporateVal)
            streamWrit.WriteLine(basText.alignmentMiddle(masterRow[mCardclientname], 50));
        else
            streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(basText.alignmentMiddle("Primary Card No", 20) + basText.alignmentMiddle("Credit Limit", 13) + basText.alignmentMiddle("Available Limit", 20) + basText.alignmentMiddle("Minimum Due", 13) + basText.alignmentMiddle("Due Date", 15) + basText.alignmentMiddle("Over Due Amount", 15)); //+ basText.formatNum(masterRow[accountLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mTotalcredits],"#,##0.00",9)+" "+basText.formatNum(masterRow[accountAvailableLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mMindueamount],"#,##0.00",15)+" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",15)+" "+basText.formatDate(masterRow[mStetementduedate],"dd/MM/yy") //TOTALCREDITS// basText.alignmentMiddle(masterRow[accountAvailableLimit],20) //"#,##0"
        //streamWrit.WriteLine(basText.alignmentMiddle((curMainCard), 20) + basText.alignmentMiddle(masterRow[mAccountlim], 13) + basText.formatNum(masterRow[mAccountavailablelim], "##0", 20, "M") + basText.alignmentMiddle(masterRow[mMindueamount], 13) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M") + basText.formatNum(masterRow[mTotaloverdueamount], "##0.00", 13, "M")); //+ basText.formatNum(masterRow[accountLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mTotalcredits],"#,##0.00",9)+" "+basText.formatNum(masterRow[accountAvailableLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mMindueamount],"#,##0.00",15)+" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",15)+" "+basText.formatDate(masterRow[mStetementduedate],"dd/MM/yy") //TOTALCREDITS// basText.alignmentMiddle(masterRow[accountAvailableLimit],20) //"#,##0"
        streamWrit.WriteLine(basText.alignmentMiddle(basText.formatCardNumber(curMainCard), 20) + basText.formatNum(masterRow[mAccountlim], "#,##0.00", 13, "M") + basText.formatNum(masterRow[mAccountavailablelim], "#,##0.00", 20, "M") + basText.formatNum(masterRow[mMindueamount], "#,##0.00", 13, "M") + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M") + basText.formatNum(masterRow[mTotaloverdueamount], "#,##0.00", 13, "M")); //+ basText.formatNum(masterRow[accountLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mTotalcredits],"#,##0.00",9)+" "+basText.formatNum(masterRow[accountAvailableLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mMindueamount],"#,##0.00",15)+" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",15)+" "+basText.formatDate(masterRow[mStetementduedate],"dd/MM/yy") //TOTALCREDITS// basText.alignmentMiddle(masterRow[accountAvailableLimit],20) //"#,##0"
        streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(" {0:dd/MM} {1:dd/MM} {2,-24}  {3,-57} {4,18}", "T Date", "P Date", basText.trimStr("Reference No", 20), basText.trimStr("Description", 57), basText.alignmentMiddle("Amount", 23)); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)
        if (pageNo == 1)
            {
            streamWrit.WriteLine(basText.replicat(" ", 43) + basText.alignmentLeft("Previous Balance", 67) + basText.formatNumUnSign(masterRow[mOpeningbalance], "##0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance])));  //+ CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign]))
            //CurPageRec4Dtl++; curAccRows++;
            }
        else
            streamWrit.WriteLine(String.Empty);

        streamWrit.WriteLine(String.Empty);
        totalPages++;
        if (isEmailStat)
            totNoOfPageEmailStat++;
        else
            totNoOfPageStat++;
        }

    protected void printDetail()
        {
        DateTime trnsDate = clsBasValid.validateDate(detailRow[dTransdate]), postingDate = clsBasValid.validateDate(detailRow[dPostingdate]);
        if (trnsDate > postingDate)
            trnsDate = postingDate;

        if (masterRow[mAccountcurrency].ToString().Trim() != detailRow[dOrigtrancurrency].ToString().Trim())
            strForeignCurr = basText.formatNum(detailRow[dOrigtranamount], "#0.00;(#0.00)") + " " + basText.Replace(clsBasValid.validateStr(detailRow[dOrigtrancurrency]), "XXX", "   ");
        else
            strForeignCurr = basText.replicat(" ", 16);

        if (strForeignCurr.Trim() == "0")
            strForeignCurr = basText.replicat(" ", 16);

        string trnsDesc; //= detailRow[dTrandescription].ToString().Trim() + " " + detailRow[dMerchant].ToString().Trim()
        if (detailRow[dMerchant].ToString().Trim() == "")
            trnsDesc = detailRow[dTrandescription].ToString().Trim();
        else
            trnsDesc = detailRow[dMerchant].ToString().Trim();

        //			streamWrit.WriteLine("  {0:dd/MM}  {1:dd/MM}  {2,-40} {3,12} {4,3} {5,16} {6,2}",trnsDate,postingDate,basText.trimStr(detailRow[dTrandescription].ToString().Trim() + " " + detailRow[dMerchant].ToString().Trim(),45),basText.formatNum(detailRow[dOrigtranamount],"#,##0.00;(#,##0.00)"),basText.Replace(clsBasValid.validateStr(detailRow[dOrigtrancurrency]),"XXX","   "),basText.formatNum(detailRow[dBilltranamount],"#,##0.00;(#,##0.00)"),CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)
        streamWrit.WriteLine("  {0:dd/MM}  {1:dd/MM}  {2,-24}  {3,-40} {4,16} {5,16} {6,2}", trnsDate, postingDate, basText.trimStr(detailRow[dRefereneno].ToString().Trim(), 24), basText.trimStr(trnsDesc, 40), strForeignCurr, basText.formatNum(detailRow[dBilltranamount], "#0.00;(#0.00)"), CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + " " + isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)
        if (isEmailStat)
            totNoOfTransEmail++;
        else
            totNoOfTransactions++;
        }

    protected void printCardFooter()
        {
        //completePageDetailRecords();
        streamWrit.WriteLine(String.Empty);
        if (pageNo == totalAccPages)
            streamWrit.WriteLine(basText.replicat(" ", 43) + basText.alignmentLeft("Current Balance", 67) + basText.formatNumUnSign(masterRow[mClosingbalance], "#0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])));
        else
            streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(String.Empty);
        //		streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(
    basText.alignmentMiddle("Primary Card No", 35)
  + basText.alignmentMiddle("Minimum Payment Due", 20)
  + basText.alignmentMiddle("Due Date", 15)
  + basText.alignmentMiddle("Opening Balance", 15)
  + basText.alignmentMiddle("Payments", 15)
            //+ basText.alignmentMiddle("Cash&Purchases", 15)
  + basText.alignmentMiddle("Cash Advance", 15)
  + basText.alignmentMiddle("Purchase / Other", 16)
  + basText.alignmentMiddle("Charges", 15)
  + basText.alignmentMiddle("Interest", 15)
  + basText.alignmentMiddle("Closing Balance", 19));
        streamWrit.WriteLine(basText.alignmentMiddle(basText.formatCardNumber(curMainCard), 35) + basText.alignmentMiddle(masterRow[mMindueamount], 20) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M")
     + basText.alignmentMiddle(basText.formatNumUnSign(masterRow[mOpeningbalance], "#,##0.00", 12) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance].ToString())), 15)
     + basText.alignmentMiddle(basText.formatNum(masterRow[mTotalpayments], "#,##0.00", 12, "M") + " " + DbCr(Convert.ToDecimal(masterRow[mTotalpayments])), 15)
            //+ basText.alignmentMiddle(basText.formatNum(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal]), "#,##0.00", 12) + CrDbMinus(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal])), 15)
     + basText.alignmentMiddle(basText.formatNum(Convert.ToDecimal(masterRow[mTotalcashwithdrawal]), "#,##0.00", 12) + CrDbMinus(Convert.ToDecimal(masterRow[mTotalcashwithdrawal])), 15)
     + basText.alignmentMiddle(basText.formatNum(Convert.ToDecimal(masterRow[mTotalpurchases]), "#,##0.00", 12) + CrDbMinus(Convert.ToDecimal(masterRow[mTotalpurchases])), 15)
     + basText.alignmentMiddle(basText.formatNum(masterRow[mTotalcharges].ToString(), "#,##0.00", 12) + " " + CrDbMinus(Convert.ToDecimal(masterRow[mTotalcharges])), 15)
     + basText.alignmentMiddle(basText.formatNum(masterRow[mTotalinterest].ToString(), "#,##0.00", 12) + " " + CrDbMinus(Convert.ToDecimal(masterRow[mTotalinterest])), 15)
     + basText.alignmentMiddle(basText.formatNumUnSign(masterRow[mClosingbalance], "#,#0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])), 15));

        }

    protected virtual void calcAccountRows()
        {
        curAccRows = 0;
        accountRows = null;
        stmNo = masterRow[mStatementno].ToString();
        stmNo = masterRow[accountNoName].ToString();

        accountRows = DSstatement.Tables["tStatementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[accountNoName]).ToString().Trim() + "'");
        totalAccPages = 0;
        totAccRows = 0;//1
        string prevCardNo = String.Empty, CurCardNo = String.Empty;
        int currAccRowsPages = 0;
        foreach (DataRow dtAccRow in accountRows) //mRow.GetChildRows(StatementNoDRel)
            {
            if ((dtAccRow[dPostingdate] == DBNull.Value) && (dtAccRow[dDocno] == DBNull.Value)) continue;
            CurCardNo = dtAccRow[dCardno].ToString();
            //if (CurCardNo.Trim().Length < 1) continue;

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

        totSuppCrdsInAcc = 0;
        totCrdNoInAcc = curCrdNoInAcc = 0;
        mainRows = DSstatement.Tables["tStatementMasterTable"].Select(accountNoName + " = '" + clsBasValid.validateStr(masterRow[accountNoName]) + "'");//ACCOUNTNO
        curMainCard = CurCardNo = "";
        foreach (DataRow mainRow in mainRows) //mRow.GetChildRows(StatementNoDRel)
            {
            totCrdNoInAcc++;
            CurCardNo = mainRow[mCardno].ToString();

            if (mainRow[mCardprimary].ToString() == "N")
                {
                if (isValidateCard(mainRow[mCardstate].ToString()))
                    {
                    tmpDtlRows = DSstatement.Tables["tStatementDetailTable"].Select(dCardno + " = '" + CurCardNo + "' and " + dPostingdate + " is not null and " + dDocno + " is not null");
                    if (tmpDtlRows.Length > 0)
                        totSuppCrdsInAcc++;
                    }
                }

            if (mainRow[mCardprimary].ToString() == "Y" && curMainCard == string.Empty)
                {
                curMainCard = CurCardNo; //mainRow[mCardno].ToString();
                }
            if (mainRow[mCardprimary].ToString() == "Y" && isValidateCard(mainRow[mCardstate].ToString()))
                {
                curMainCard = CurCardNo; //mainRow[mCardno].ToString();
                //break;
                }
            }


        if (curMainCard == "")
            curMainCard = CurCardNo;
        if (totSuppCrdsInAcc > 0)
            {
            totAccRows += totSuppCrdsInAcc;
            totalAccPages = totAccRows / MaxDetailInPage;
            totalAccPages = (totAccRows % MaxDetailInPage) > 0 ? ++totalAccPages : totalAccPages;
            }
        }


    protected virtual void completePageDetailRecords()
        {
        //int curPageLine =CurPageRec4Dtl;
        if (pageNo == totalAccPages)
            for (int curPageLine = CurPageRec4Dtl; curPageLine < MaxDetailInPage; curPageLine++)
                streamWrit.WriteLine(String.Empty);
        }

    private void add2FileList(string pFileName)
        {
        int myIndex = aryLstFiles.BinarySearch((object)pFileName);
        if (myIndex < 0)
            aryLstFiles.Add(@pFileName);
        }

    public void finalizeStat()
        {
        clsBasFile.generateFileMD5(aryLstFiles, @clsBasFile.getPathWithoutExtn(commnFileName) + ".MD5");
        aryLstFiles.Add(@clsBasFile.getPathWithoutExtn(commnFileName) + ".MD5");
        SharpZip zip = new SharpZip();
        zip.createZip(aryLstFiles, @clsBasFile.getPathWithoutExtn(commnFileName) + ".zip", "");
        DSstatement.Dispose();

        }

    private void printStatementSummary()
        {
        streamSummary.WriteLine("__________________________");
        streamSummary.WriteLine("");
        streamSummary.WriteLine("No of Statements   " + totNoOfCardStat.ToString());
        streamSummary.WriteLine("No of Pages        " + totNoOfPageStat.ToString());
        streamSummary.WriteLine("No of Transactions " + totNoOfTransactions.ToString());
        totNoOfCardStat = 0;
        totNoOfPageStat = 0;
        totNoOfTransactions = 0;
        }

    public string bankName
        {
        get { return strBankName; }
        set { strBankName = value; }
        }// bankName

    public frmStatementFile setFrm
        {
        set { frmMain = value; }
        }// setFrm

    public string productCond
        {
        get { return strProductCond; }
        set { strProductCond = value; }
        }  // productCond

    ~clsStatTxtLbl_Suez()
        {
        DSstatement.Dispose();
        }
    }
