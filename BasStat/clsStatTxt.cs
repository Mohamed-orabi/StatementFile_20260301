using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Xml;
using System.Collections;

// Branch 16
public class clsStatTxt : clsBasStatement
{
    protected string strBankName;
    protected FileStream fileStrm, fileStrmByEmail, fileSummary;
    protected StreamWriter streamWrit, strmWriteByEmail, strmWriteCommon, streamSummary;
    //protected DataSet DSstatement;
    //protected OracleDataReader drPrimaryCards, drMaster,drDetail;
    protected DataRow masterRow;
    protected DataRow[] Summary;
    protected DataRow detailRow;
    protected string strEndOfLine = "\u000D";  //+ "M" ^M
    //		protected string strEndOfPage = "\u000C"  ;  //+ "\u000D"+ "M" ^L + ^M
    protected string strEndOfPage = "\u000C";  //  + basText.replicat(" ",85) + "F 2"  + "\u000D"+ "M" ^L + ^M
    protected string strEndOfAccount = "----------------- END OF STATEMENT -----------------";  //  + basText.replicat(" ",85) + "F 2"  + "\u000D"+ "M" ^L + ^M
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
    private string strPaymentSystem = string.Empty;
    private string strBillingCycle = string.Empty;
    private string strExcludeCond = string.Empty;
    protected bool isExcludedVal = false;
    protected bool isSplittedVal = false;

    public clsStatTxt()
    {

    }

    //CR
    public virtual string Statement(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pCurDate, string pStmntType, bool pAppendData)
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
            if (isReward)
            {
                maintainData.notRward = false;
                maintainData.curRewardCond = rewardCond;
                //maintainData.fixReward(pBankCode, rewardCond);
            }
            maintainData.matchCardBranch4Account(pBankCode);

            pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);
            vCurDate = pCurDate; //DateTime.Now.AddMonths(-1);
            strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
            clsBasFile.createDirectory(pStrFileName + vCurDate.ToString("yyyyMM") + pBankName);
            pStrFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM") + ".txt";
            strBankName = pBankName;
            comnFileName = strOutputFile = pStrFileName;
            if (isNotSplit)
                commnFileName = comnFileName;

            if (isInEmailService)
            {
                comnFileName = clsBasFile.getPathWithoutExtn(pStrFileName) + "_NotHaveEmail." + clsBasFile.getFileExtn(pStrFileName);
                // By Email file
                fileStrmByEmail = new FileStream(clsBasFile.getPathWithoutExtn(pStrFileName) + "_ByEmail." + clsBasFile.getFileExtn(pStrFileName), FileMode.Create);
                strmWriteByEmail = new StreamWriter(fileStrmByEmail, Encoding.Default);
                strmWriteByEmail.AutoFlush = true;
            }

            // open output file
            if (!(pBankCode == 33))
            {
                fileStrm = new FileStream(comnFileName, FileMode.Create); //Create
                streamWrit = new StreamWriter(fileStrm, Encoding.Default);
                streamWrit.AutoFlush = true;
            }
            strmWriteCommon = streamWrit;


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
            //     clsBasStatement.mainTableCond =  " substr(cardno,1,6) !='421192' "; //" cardproduct != 'Visa Business' ";
            //     clsBasStatement.supTableCond =  " substr(cardno,1,6) !='421192' "; 
            if (createCorporateVal)
            {
                isCorporateVal = true;
                accountNoName = mCardaccountno;
                accountLimit = mCardlimit;
                accountAvailableLimit = mCardavailablelimit;
            }
            if (isReward)
            {
                curRewardCond = rewardCond;
                strMainTableCond = "m.contracttype != " + rewardCond;
                strSubTableCond = "d.trandescription != 'Calculated Points'";
            }
            // data retrieve
            //strMainTableCond = "m.Cardproduct = 'VISA GOLD'";

            //if (isFilteredVal && pBankCode == 130)
            //{
            //    MainTableCond = " m.cardproduct in " + strProductCond + "";//strWhereCond
            //    supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.cardproduct in " + strProductCond + ")";
            //}

            if (isFilteredVal)
            {
                MainTableCond = " m.contracttype in " + strProductCond + "";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + ")";
            }
            if (isExcludedVal)
            {
                MainTableCond = " m.cardproduct not in " + strExcludeCond + "";//strWhereCond
                                                                               //supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.cardproduct not in " + strExcludeCond + ")";
            }
            if (isSplittedVal)
            {
                if (pBranch == 73)
                {
                    MainTableCond = " m.contracttype like '" + strPaymentSystem + "%' and m.contracttype like '%" + strBillingCycle + "%'";//strWhereCond
                    supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype like '" + strPaymentSystem + "%' and x.contracttype like '%" + strBillingCycle + "%')";
                }
                else if (pBranch == 115)
                {
                    MainTableCond = " m.contracttype like '%" + strBillingCycle + "%'";//strWhereCond
                    supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype like '%" + strBillingCycle + "%')";
                }
            }



            //if (pBankCode == 122)
            //    FillStatementDataSet_SortCardPriority_Exclude_VisaCards(pBankCode, "vip"); //DSstatement =  //16); //3
            //else
            FillStatementDataSet(pBankCode, "vip"); //DSstatement =  //16); //3
            
            if (pBankCode == 130)
            {
                GetSummaryPerCorporate(pBankCode);
            }

            getCardProduct(pBankCode);
            if (isInEmailService && isNotSplit)
                getClientEmail(pBankCode);
            if (pBranch == 76)
                getClientEmail(pBankCode);
            if (isReward)
                getReward(pBankCode);
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
                if (pBankCode == 130)
                    Summary = SummaryCor.Tables["SummaryCor"].Select("clientid = " + masterRow[mClientid].ToString());



                strCardNo = masterRow[mCardno].ToString().Trim();
                if (masterRow[mHOLSTMT].ToString() == "Y")
                    continue;
                try
                {
                    if (string.IsNullOrWhiteSpace(strCardNo) || (clsCheckCard.CalcLuhnCheckDigit(strCardNo.Substring(0, 15)) != strCardNo.Substring(15, 1))) //!clsCheckCard.isValidCard(strCardNo) ||
                    {
                        strmWriteErr.WriteLine("Invalid Card Number" + strCardNo);
                        numOfErr++;
                    }
                }
                catch (Exception ex)
                {
                    clsBasErrors.catchError(ex.InnerException);
                }
                if (pBankCode == 33)
                {
                    //check Contract
                    if (cContract != masterRow[mContracttype].ToString().Trim())
                    {
                        if (cContract != string.Empty)
                        {
                            strmWriteCommon.Flush();
                            strmWriteCommon.Close();
                            fileStrm.Close();
                        }
                        cContract = masterRow[mContracttype].ToString().Trim();


                        curFileName = clsBasFile.getPathWithoutExtn(pStrFileName)
                          + "_" + cContract + "." + clsBasFile.getFileExtn(pStrFileName);
                        //curFileName = pStrFileName;
                        add2FileList(curFileName);
                        fileStrm = new FileStream(curFileName, FileMode.Append); //FileMode.Create Create
                        strmWriteCommon = new StreamWriter(fileStrm, Encoding.Default);
                        strmWriteCommon.AutoFlush = true;
                    }
                }

                if (strCardNo.Length != 16)
                {
                    strmWriteErr.WriteLine("CardNo Length not equal 16 for AccountNo " + masterRow[accountNoName].ToString().Trim() + " Card No =" + strCardNo);
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
                    if (isInEmailService)
                    {
                        streamWrit.Flush();
                        strmWriteByEmail.Flush();
                        if (haveEmail(masterRow[mClientid].ToString()))
                        {
                            strmWriteCommon = strmWriteByEmail;
                            isEmailStat = true;
                        }
                        else
                        {
                            strmWriteCommon = streamWrit;
                            isEmailStat = false;
                        }
                    }

                    if (prevAccountNo == string.Empty)
                    {
                        if (curMonth != Convert.ToDateTime(masterRow[mStatementdateto]).Month - clsSessionValues.statGenAfterMonth)//,"dd/MM/yyyy"
                        {
                            preExit = false;
                            fileStrm.Close();
                            fileStrmErr.Close();
                            clsBasFile.deleteFile(@strOutputFile);
                            clsBasFile.deleteFile(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName));
                            return "Error in Generation " + pBankName;
                        }
                    }

                    if (pageNo != totalAccPages && prevAccountNo != "")// 
                    {
                        //>MessageBox.Show( "Error in Genrating Statement", "Please Call The Programmer", MessageBoxButtons.OK ,
                        //>  MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1); //,	MessageBoxOptions.RightAlign
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
                        if (totAccRows < 1
                          && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0) //             || (masterRow[mCardno].ToString() == curMainCard   // Convert.ToDecimal(
                        {
                            isHaveF3 = true;
                            //pageNo=1; totalAccPages =1;
                            continue;
                        }

                    prevAccountNo = masterRow[accountNoName].ToString();
                    if (pBranch == 76)
                    {
                        CustomerArabicname = "";
                        emailRows = DSemails.Tables["Emails"].Select("idclient = " + masterRow[mClientid]);
                        if (emailRows.Length != 0)
                            CustomerArabicname = emailRows[0][5].ToString().Trim();
                    }
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
                    CurPageRec4Dtl = CurPageRec4Dtl + 1;
                    if (pBranch == 76)
                    {
                        if (curSuppCard != detailRow[dCardno].ToString())
                        {
                            tmpDtlRows = DSstatement.Tables["tStatementMasterTable"].Select(mCardno + " = '" + detailRow[dCardno] + "'");
                            if (tmpDtlRows.Length > 0)
                            {
                                if (tmpDtlRows[0][mCardprimary].ToString() == "N" && isValidateCard(tmpDtlRows[0][mCardstate].ToString()))
                                {
                                    //CurPageRec4Dtl++;
                                    CurPageRec4Dtl = CurPageRec4Dtl + 2;
                                    //curCrdNoInAcc++;
                                    curCrdNoInAcc = curCrdNoInAcc + 1;
                                    //curAccRows++;
                                    curAccRows = curAccRows + 2;
                                    if (CurPageRec4Dtl >= MaxDetailInPage)
                                    {
                                        //CurPageRec4Dtl = 1;
                                        for (int curPageLine = CurPageRec4Dtl - 3; curPageLine < MaxDetailInPage; curPageLine++)
                                            strmWriteCommon.WriteLine(String.Empty);
                                        CurPageRec4Dtl++;
                                        //curAccRows++;
                                        printCardFooter();
                                        pageNo++;
                                        printHeader();
                                        streamWrit.WriteLine(basText.replicats(" ", 23) + ">> Supplementary Card : " + basText.formatCardNumber(detailRow[dCardno].ToString()));
                                        streamWrit.WriteLine(basText.replicats(" ", 23) + ">> Card Holder Name: " + basText.alignmentLeft(masterRow[mCardclientname].ToString(), 50));
                                        CurPageRec4Dtl = 3;
                                    }
                                    else
                                    {
                                        streamWrit.WriteLine(basText.replicats(" ", 23) + ">> Supplementary Card : " + basText.formatCardNumber(detailRow[dCardno].ToString()));
                                        streamWrit.WriteLine(basText.replicats(" ", 23) + ">> Card Holder Name: " + basText.alignmentLeft(masterRow[mCardclientname].ToString(), 50));
                                        //CurPageRec4Dtl += 1;
                                    }
                                }
                            }
                        }
                        curSuppCard = detailRow[dCardno].ToString();
                    }
                    printDetail();
                    //if(!((detailRow[dPostingdate]== DBNull.Value) && (detailRow[dDocno]== DBNull.Value))) 

                } //end of detail foreach
                //printCardFooter();//if pages is based on card
                // if(masterRow[mCardprimary].ToString() == "Y" && (masterRow[mCardstate].ToString() == "Given" || masterRow[mCardstate].ToString() == "Embossed" || masterRow[mCardstate].ToString() == "New"))
                curCrdNoInAcc++;
                //if (pBankCode == 76)
                //{
                //    if ((curAccRows >= totAccRows && totAccRows != 0) || (totAccRows == 0 && curCrdNoInAcc == totCrdNoInAcc))
                //    {
                //        if (isSupp == true)
                //        {
                //            completePageDetailRecords(76);
                //            printCardFooter();//if pages is based on account
                //            CurPageRec4Dtl = 0; //>pageNo=1; if pages is based on account
                //            curAccRows = 0;
                //            isSupp = false;
                //        }
                //        else
                //        {
                //            completePageDetailRecords();
                //            printCardFooter();//if pages is based on account
                //            CurPageRec4Dtl = 0; //>pageNo=1; if pages is based on account
                //            curAccRows = 0;
                //        }
                //    }
                //}
                //else
                if ((curAccRows >= totAccRows && totAccRows != 0) || (totAccRows == 0 && curCrdNoInAcc == totCrdNoInAcc))
                {
                    if (pBankCode != 94) //FABG Separation By Ahmed Elbadawei Jira FABG-2825
                    {
                        completePageDetailRecords();
                    }
                    printCardFooter();//if pages is based on account
                    CurPageRec4Dtl = 0; //>pageNo=1; if pages is based on account
                    curAccRows = 0;
                }
                //streamWrit.WriteLine(strEndOfPage);
                //pageNo=1; CurPageRec4Dtl=0; //if pages is based on card
                //completePageDetailRecords();

            } //end of Master foreach
            strmWriteCommon.WriteLine(strEndOfPage + strEndOfAccount);
            //fillStatementHistory(pStmntType,pAppendData);
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
                if (!(pBankCode == 33))
                {
                    streamWrit.Flush();
                    streamWrit.Close();
                }
                else
                {
                    strmWriteCommon.Flush();
                    strmWriteCommon.Close();
                }
                fileStrm.Close();

                printStatementSummary();

                if (isInEmailService)
                {
                    // Close output File
                    strmWriteByEmail.Flush();
                    strmWriteByEmail.Close();
                    fileStrmByEmail.Close();
                }

                // Close Summary File
                streamSummary.Flush();
                streamSummary.Close();
                fileSummary.Close();

                strmWriteErr.Flush();
                strmWriteErr.Close();
                fileStrmErr.Close();
                if (numOfErr == 0)
                    clsBasFile.deleteFile(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName));
                //ArrayList aryLstFiles = new ArrayList();
                if (!(pBankCode == 33))
                    aryLstFiles.Add(comnFileName);
                numOfErr += validateNoOfLines(aryLstFiles, NoLinePerPage);//53
                if (numOfErr == 0)
                    clsBasFile.deleteFile(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName));
                else
                {
                    aryLstFiles.Add(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err.txt");
                    rtrnStr += string.Format(" with {0} Errors", numOfErr);
                }
                if (isInEmailService)
                    aryLstFiles.Add(clsBasFile.getPathWithoutExtn(pStrFileName) + "_ByEmail.txt");
                aryLstFiles.Add(@fileSummaryName);
                ///
                if (!isCorporateGenrated)
                    if (pStmntType == "Corporate")
                    {
                        ArrayList aryLstFilesCorp = new ArrayList();
                        clsStatement_CommonCorpCmpny stmntBNPCorp = new clsStatement_CommonCorpCmpny();// + "NSGB_Business_Statement.txt"
                        stmntBNPCorp.setFrm = frmMain;
                        aryLstFilesCorp = stmntBNPCorp.Statement(ppStrFileName, pBankName, pBankCode, pStrFile, pCurDate);// + "NSGB_Business_Statement.txt"
                        foreach (object str in aryLstFilesCorp)
                            aryLstFiles.Add((string)str);

                        stmntBNPCorp = null; isCorporateGenrated = true;
                    }
                ///
                if (isNotSplit)
                {
                    finalizeStat();
                }

            }
        }
        return rtrnStr;
    }
    //CP
    public virtual string Statement(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pCurDate, string pStmntType, bool pAppendData, string add2FileName)
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
            if (isReward)
            {
                maintainData.notRward = false;
                maintainData.curRewardCond = rewardCond;
                //maintainData.fixReward(pBankCode, rewardCond);
            }
            maintainData.matchCardBranch4Account(pBankCode);

            pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);
            vCurDate = pCurDate; //DateTime.Now.AddMonths(-1);
            strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
            clsBasFile.createDirectory(pStrFileName + vCurDate.ToString("yyyyMM") + pBankName);
            pStrFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + add2FileName + "_" + vCurDate.ToString("yyyyMM") + ".txt";
            strBankName = pBankName;
            comnFileName = strOutputFile = pStrFileName;
            if (isNotSplit)
                commnFileName = comnFileName;

            if (isInEmailService)
            {
                comnFileName = clsBasFile.getPathWithoutExtn(pStrFileName) + "_NotHaveEmail." + clsBasFile.getFileExtn(pStrFileName);
                // By Email file
                fileStrmByEmail = new FileStream(clsBasFile.getPathWithoutExtn(pStrFileName) + "_ByEmail." + clsBasFile.getFileExtn(pStrFileName), FileMode.Create);
                strmWriteByEmail = new StreamWriter(fileStrmByEmail, Encoding.Default);
                strmWriteByEmail.AutoFlush = true;
            }

            // open output file
            if (!(pBankCode == 33))
            {
                fileStrm = new FileStream(comnFileName, FileMode.Create); //Create
                streamWrit = new StreamWriter(fileStrm, Encoding.Default);
                streamWrit.AutoFlush = true;
            }
            strmWriteCommon = streamWrit;


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
            //     clsBasStatement.mainTableCond =  " substr(cardno,1,6) !='421192' "; //" cardproduct != 'Visa Business' ";
            //     clsBasStatement.supTableCond =  " substr(cardno,1,6) !='421192' "; 
            if (createCorporateVal)
            {
                isCorporateVal = true;
                accountNoName = mCardaccountno;
                accountLimit = mCardlimit;
                accountAvailableLimit = mCardavailablelimit;
            }
            if (isReward)
            {
                curRewardCond = rewardCond;
                strMainTableCond = "m.contracttype != " + rewardCond;
                strSubTableCond = "d.trandescription != 'Calculated Points'";
            }
            // data retrieve
            //strMainTableCond = "m.Cardproduct = 'VISA GOLD'";
            if (isFilteredVal)
            {
                MainTableCond = " m.contracttype in " + strProductCond + "";//strWhereCond
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + ")";
            }
            if (isExcludedVal)
            {
                MainTableCond = " m.cardproduct not in " + strExcludeCond + "";//strWhereCond
                                                                               //supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.cardproduct not in " + strExcludeCond + ")";
            }
            FillStatementDataSet(pBankCode, "vip"); //DSstatement =  //16); //3
            getCardProduct(pBankCode);
            if (isInEmailService && isNotSplit)
                getClientEmail(pBankCode);
            if (pBranch == 76)
                getClientEmail(pBankCode);
            if (isReward)
                getReward(pBankCode);
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
                if (pBankCode == 33)
                {
                    //check Contract
                    if (cContract != masterRow[mContracttype].ToString().Trim())
                    {
                        if (cContract != string.Empty)
                        {
                            strmWriteCommon.Flush();
                            strmWriteCommon.Close();
                            fileStrm.Close();
                        }
                        cContract = masterRow[mContracttype].ToString().Trim();


                        curFileName = clsBasFile.getPathWithoutExtn(pStrFileName)
                          + "_" + cContract + "." + clsBasFile.getFileExtn(pStrFileName);
                        //curFileName = pStrFileName;
                        add2FileList(curFileName);
                        fileStrm = new FileStream(curFileName, FileMode.Append); //FileMode.Create Create
                        strmWriteCommon = new StreamWriter(fileStrm, Encoding.Default);
                        strmWriteCommon.AutoFlush = true;
                    }
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
                    if (isInEmailService)
                    {
                        streamWrit.Flush();
                        strmWriteByEmail.Flush();
                        if (haveEmail(masterRow[mClientid].ToString()))
                        {
                            strmWriteCommon = strmWriteByEmail;
                            isEmailStat = true;
                        }
                        else
                        {
                            strmWriteCommon = streamWrit;
                            isEmailStat = false;
                        }
                    }

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
                        //>MessageBox.Show( "Error in Genrating Statement", "Please Call The Programmer", MessageBoxButtons.OK ,
                        //>  MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1); //,	MessageBoxOptions.RightAlign
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
                        if (totAccRows < 1
                          && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0) //             || (masterRow[mCardno].ToString() == curMainCard   // Convert.ToDecimal(
                        {
                            isHaveF3 = true;
                            //pageNo=1; totalAccPages =1;
                            continue;
                        }

                    prevAccountNo = masterRow[accountNoName].ToString();
                    if (pBranch == 76)
                    {
                        CustomerArabicname = "";
                        emailRows = DSemails.Tables["Emails"].Select("idclient = " + masterRow[mClientid]);
                        if (emailRows.Length != 0)
                            CustomerArabicname = emailRows[0][5].ToString().Trim();
                    }
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
                    CurPageRec4Dtl = CurPageRec4Dtl + 1;
                    if (pBranch == 76)
                    {
                        if (curSuppCard != detailRow[dCardno].ToString())
                        {
                            tmpDtlRows = DSstatement.Tables["tStatementMasterTable"].Select(mCardno + " = '" + detailRow[dCardno] + "'");
                            if (tmpDtlRows.Length > 0)
                            {
                                if (tmpDtlRows[0][mCardprimary].ToString() == "N" && isValidateCard(tmpDtlRows[0][mCardstate].ToString()))
                                {
                                    //CurPageRec4Dtl++;
                                    CurPageRec4Dtl = CurPageRec4Dtl + 2;
                                    //curCrdNoInAcc++;
                                    curCrdNoInAcc = curCrdNoInAcc + 1;
                                    //curAccRows++;
                                    curAccRows = curAccRows + 2;
                                    if (CurPageRec4Dtl >= MaxDetailInPage)
                                    {
                                        //CurPageRec4Dtl = 1;
                                        for (int curPageLine = CurPageRec4Dtl - 3; curPageLine < MaxDetailInPage; curPageLine++)
                                            strmWriteCommon.WriteLine(String.Empty);
                                        CurPageRec4Dtl++;
                                        //curAccRows++;
                                        printCardFooter();
                                        pageNo++;
                                        printHeader();
                                        streamWrit.WriteLine(basText.replicats(" ", 23) + ">> Supplementary Card : " + basText.formatCardNumber(detailRow[dCardno].ToString()));
                                        streamWrit.WriteLine(basText.replicats(" ", 23) + ">> Card Holder Name: " + basText.alignmentLeft(masterRow[mCardclientname].ToString(), 50));
                                        CurPageRec4Dtl = 3;
                                    }
                                    else
                                    {
                                        streamWrit.WriteLine(basText.replicats(" ", 23) + ">> Supplementary Card : " + basText.formatCardNumber(detailRow[dCardno].ToString()));
                                        streamWrit.WriteLine(basText.replicats(" ", 23) + ">> Card Holder Name: " + basText.alignmentLeft(masterRow[mCardclientname].ToString(), 50));
                                        //CurPageRec4Dtl += 1;
                                    }
                                }
                            }
                        }
                        curSuppCard = detailRow[dCardno].ToString();
                    }
                    printDetail();
                    //if(!((detailRow[dPostingdate]== DBNull.Value) && (detailRow[dDocno]== DBNull.Value))) 

                } //end of detail foreach
                //printCardFooter();//if pages is based on card
                // if(masterRow[mCardprimary].ToString() == "Y" && (masterRow[mCardstate].ToString() == "Given" || masterRow[mCardstate].ToString() == "Embossed" || masterRow[mCardstate].ToString() == "New"))
                curCrdNoInAcc++;
                //if (pBankCode == 76)
                //{
                //    if ((curAccRows >= totAccRows && totAccRows != 0) || (totAccRows == 0 && curCrdNoInAcc == totCrdNoInAcc))
                //    {
                //        if (isSupp == true)
                //        {
                //            completePageDetailRecords(76);
                //            printCardFooter();//if pages is based on account
                //            CurPageRec4Dtl = 0; //>pageNo=1; if pages is based on account
                //            curAccRows = 0;
                //            isSupp = false;
                //        }
                //        else
                //        {
                //            completePageDetailRecords();
                //            printCardFooter();//if pages is based on account
                //            CurPageRec4Dtl = 0; //>pageNo=1; if pages is based on account
                //            curAccRows = 0;
                //        }
                //    }
                //}
                //else
                if ((curAccRows >= totAccRows && totAccRows != 0) || (totAccRows == 0 && curCrdNoInAcc == totCrdNoInAcc))
                {
                    completePageDetailRecords();
                    printCardFooter();//if pages is based on account
                    CurPageRec4Dtl = 0; //>pageNo=1; if pages is based on account
                    curAccRows = 0;
                }
                //streamWrit.WriteLine(strEndOfPage);
                //pageNo=1; CurPageRec4Dtl=0; //if pages is based on card
                //completePageDetailRecords();

            } //end of Master foreach
            strmWriteCommon.WriteLine(strEndOfPage + strEndOfAccount);
            //fillStatementHistory(pStmntType,pAppendData);
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
                if (!(pBankCode == 33))
                {
                    streamWrit.Flush();
                    streamWrit.Close();
                }
                else
                {
                    strmWriteCommon.Flush();
                    strmWriteCommon.Close();
                }
                fileStrm.Close();

                printStatementSummary();

                if (isInEmailService)
                {
                    // Close output File
                    strmWriteByEmail.Flush();
                    strmWriteByEmail.Close();
                    fileStrmByEmail.Close();
                }

                // Close Summary File
                streamSummary.Flush();
                streamSummary.Close();
                fileSummary.Close();

                strmWriteErr.Flush();
                strmWriteErr.Close();
                fileStrmErr.Close();
                if (numOfErr == 0)
                    clsBasFile.deleteFile(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName));
                //ArrayList aryLstFiles = new ArrayList();
                if (!(pBankCode == 33))
                    aryLstFiles.Add(comnFileName);
                numOfErr += validateNoOfLines(aryLstFiles, NoLinePerPage);//53
                if (numOfErr == 0)
                    clsBasFile.deleteFile(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName));
                else
                {
                    aryLstFiles.Add(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err.txt");
                    rtrnStr += string.Format(" with {0} Errors", numOfErr);
                }
                if (isInEmailService)
                    aryLstFiles.Add(clsBasFile.getPathWithoutExtn(pStrFileName) + "_ByEmail.txt");
                aryLstFiles.Add(@fileSummaryName);
                ///
                if (!isCorporateGenrated)
                    if (pStmntType == "Corporate")
                    {
                        ArrayList aryLstFilesCorp = new ArrayList();
                        clsStatement_CommonCorpCmpny stmntBNPCorp = new clsStatement_CommonCorpCmpny();// + "NSGB_Business_Statement.txt"
                        stmntBNPCorp.setFrm = frmMain;
                        aryLstFilesCorp = stmntBNPCorp.Statement(ppStrFileName, pBankName, pBankCode, pStrFile, pCurDate);// + "NSGB_Business_Statement.txt"
                        foreach (object str in aryLstFilesCorp)
                            aryLstFiles.Add((string)str);

                        stmntBNPCorp = null; isCorporateGenrated = true;
                    }
                ///
                if (isNotSplit)
                {
                    finalizeStat();
                }

            }
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


    protected void calcCardlRows()
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


    protected virtual void calcAccountRows()
    {
        curAccRows = 0;
        accountRows = null;
        stmNo = masterRow[mStatementno].ToString();
        stmNo = masterRow[accountNoName].ToString();
        string CardNumber = masterRow[mCardno].ToString();
        accountRows = DSstatement.Tables["tStatementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[accountNoName]).ToString().Trim() + "'");
        if (accountRows.Length == 0 && Convert.ToInt32(masterRow[mBranch].ToString()) == 27)
            accountRows = DSstatement.Tables["tStatementDetailTable"].Select("CARDNO = '" + clsBasValid.validateStr(masterRow[mCardno]).ToString().Trim() + "'");
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
        if (pBranch == 76)
            totSuppCrdsInAcc = 0;
        totCrdNoInAcc = curCrdNoInAcc = 0;
        mainRows = DSstatement.Tables["tStatementMasterTable"].Select(accountNoName + " = '" + clsBasValid.validateStr(masterRow[accountNoName]) + "'");//ACCOUNTNO
        curMainCard = CurCardNo = "";
        foreach (DataRow mainRow in mainRows) //mRow.GetChildRows(StatementNoDRel)
        {
            totCrdNoInAcc++;
            CurCardNo = mainRow[mCardno].ToString();
            if (pBranch == 76)
            {
                if (mainRow[mCardprimary].ToString() == "N")
                {
                    if (isValidateCard(mainRow[mCardstate].ToString()))
                    {
                        tmpDtlRows = DSstatement.Tables["tStatementDetailTable"].Select(dCardno + " = '" + CurCardNo + "' and " + dPostingdate + " is not null and " + dDocno + " is not null");
                        if (tmpDtlRows.Length > 0)
                            //totSuppCrdsInAcc++;
                            totSuppCrdsInAcc = totSuppCrdsInAcc + 2;
                    }
                }
                if (mainRow[mCardprimary].ToString() == "Y")
                    curMainCard = CurCardNo; //mainRow[mCardno].ToString();
                if (mainRow[mCardprimary].ToString() == "Y" && isValidateCard(mainRow[mCardstate].ToString()))
                {
                    curMainCard = CurCardNo; //mainRow[mCardno].ToString();
                }
            }
            else
            {
                if (mainRow[mCardprimary].ToString() == "Y")
                    curMainCard = CurCardNo; //mainRow[mCardno].ToString();
                if (mainRow[mCardprimary].ToString() == "Y" && isValidateCard(mainRow[mCardstate].ToString()))
                {
                    curMainCard = CurCardNo; //mainRow[mCardno].ToString();
                    break;
                }
            }
        }


        if (curMainCard == "")
            curMainCard = CurCardNo;

        if (pBranch == 76)
        {
            if (totSuppCrdsInAcc > 0)
            {
                totAccRows += totSuppCrdsInAcc;
                totalAccPages = totAccRows / MaxDetailInPage;
                totalAccPages = (totAccRows % MaxDetailInPage) > 0 ? ++totalAccPages : totalAccPages;
            }
        }
    }


    protected virtual void completePageDetailRecords()
    {
        //int curPageLine =CurPageRec4Dtl;
        if (pageNo == totalAccPages)
            for (int curPageLine = CurPageRec4Dtl; curPageLine < MaxDetailInPage; curPageLine++)
                strmWriteCommon.WriteLine(String.Empty);
    }

    protected virtual void completePageDetailRecords(int pbranch)
    {
        if (pageNo == totalAccPages)
            for (int curPageLine = CurPageRec4Dtl; curPageLine < 18; curPageLine++)
                strmWriteCommon.WriteLine(String.Empty);
    }

    private void add2FileList(string pFileName)
    {
        int myIndex = aryLstFiles.BinarySearch((object)pFileName);
        if (myIndex < 0)
            aryLstFiles.Add(@pFileName);
    }

    private void printStatementSummary()
    {
        if (strBankName.Contains("MasterCard"))
            streamSummary.WriteLine(strBankName + " MasterCard Statement");
        else
            streamSummary.WriteLine(strBankName + " Visa Statement");
        if (isInEmailService)
        {
            streamSummary.WriteLine("");
            streamSummary.WriteLine("");
            streamSummary.WriteLine("Statements not Sent by Email");
        }
        streamSummary.WriteLine("__________________________");
        streamSummary.WriteLine("");
        streamSummary.WriteLine("No of Statements   " + totNoOfCardStat.ToString());
        streamSummary.WriteLine("No of Pages        " + totNoOfPageStat.ToString());
        streamSummary.WriteLine("No of Transactions " + totNoOfTransactions.ToString());

        if (isInEmailService)
        {
            streamSummary.WriteLine("");
            streamSummary.WriteLine("");
            streamSummary.WriteLine("__________________________");
            streamSummary.WriteLine("");
            streamSummary.WriteLine("Statements Sent by Email");
            streamSummary.WriteLine("");
            streamSummary.WriteLine("No of Statements   " + totNoOfCardEmailStat.ToString());
            streamSummary.WriteLine("No of Pages        " + totNoOfPageEmailStat.ToString());
            streamSummary.WriteLine("No of Transactions " + totNoOfTransEmail.ToString());
            streamSummary.WriteLine("");
            streamSummary.WriteLine("");
            streamSummary.WriteLine("__________________________");
            streamSummary.WriteLine("");
            streamSummary.WriteLine("Total Statements");
            streamSummary.WriteLine("");
            streamSummary.WriteLine("No of Statements   " + (totNoOfCardStat + totNoOfCardEmailStat).ToString());
            streamSummary.WriteLine("No of Pages        " + (totNoOfPageStat + totNoOfPageEmailStat).ToString());
            streamSummary.WriteLine("No of Transactions " + (totNoOfTransactions + totNoOfTransEmail).ToString());
        }

        if (clsStatementSummary.isUpdatedatble == true)
        {
            if (pBranch == 5 || pBranch == 6 || pBranch == 102 || pBranch == 122 || pBranch == 127)
                return;
            else if (pBranch == 69)
                GenerateStatementSummaryByActivity();
            else
                GenerateStatementSummary();
        }
    }

    protected bool haveEmail(string pClientID)
    {
        bool rtrnVal = false;
        emailTo = string.Empty;

        emailRows = DSemails.Tables["Emails"].Select("idclient = " + pClientID);
        for (int i = 0; i < emailRows.Length; i++)
        {
            if (!string.IsNullOrEmpty(emailRows[i][1].ToString().Trim()))
            {
                if (valdEmail.isValideEmail(emailRows[i][1].ToString().Trim()) == "ValidEmail")
                {
                    emailTo = emailRows[i][1].ToString().Trim();
                    rtrnVal = true;
                    break;
                }
            }
        }
        return rtrnVal;
    }


    public string SplitByProduct(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pCurDate, string pStmntType, bool pAppendData)
    {
        string rtrnStr = "Successfully Generate " + pBankName;
        string queryString, whereCond;
        isNotSplit = false;

        queryString = "select distinct t.cardproduct from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " t where t.branch = " + pBankCode;
        try
        {
            clsMaintainData maintainData = new clsMaintainData();
            maintainData.matchCardBranch4Account(pBankCode);
            if (isInEmailService)
                getClientEmail(pBankCode);
            commnFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM") + ".txt";
            OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
            OracleCommand cmd = new OracleCommand(queryString, conn);
            conn.Open();
            OracleDataReader rdr;
            rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                masterQuery = detailQuery = string.Empty;
                mainTableCond = " cardproduct = '" + rdr[0].ToString() + "'";
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and " + mainTableCond + ")";
                add2FileName = rdr[0].ToString() + "_";
                Statement(pStrFileName, pBankName, pBankCode, pStrFile, pCurDate, pStmntType, pAppendData, add2FileName);
                DSstatement.Clear();
                StatementNoDRel = null;
                totRec = 1;
                //DSstatement = null;
            }
            rdr.Close();
            if (isAutoinalize)
                finalizeStat();
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
        string queryString, whereCond;
        isNotSplit = false;

        queryString = "select distinct t.cardbranchpart, t.cardbranchpartname from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " t where t.branch = " + pBankCode;
        try
        {
            clsMaintainData maintainData = new clsMaintainData();
            maintainData.matchCardBranch4Account(pBankCode);
            if (isInEmailService)
                getClientEmail(pBankCode);
            commnFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM") + ".txt";
            OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
            OracleCommand cmd = new OracleCommand(queryString, conn);
            conn.Open();
            OracleDataReader rdr;
            rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                masterQuery = detailQuery = string.Empty;
                mainTableCond = " cardbranchpart = '" + rdr[0].ToString() + "'";
                supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and " + mainTableCond + ")";
                add2FileName = " (" + rdr[0].ToString() + ") " + rdr[1].ToString() + "_";
                Statement(pStrFileName, pBankName, pBankCode, pStrFile, pCurDate, pStmntType, pAppendData, add2FileName);
                DSstatement.Clear();
                StatementNoDRel = null;
                totRec = 1;
                //DSstatement = null;
            }
            rdr.Close();
            if (isAutoinalize)
                finalizeStat();
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

    public void finalizeStat()
    {
        clsBasFile.generateFileMD5(aryLstFiles, @clsBasFile.getPathWithoutExtn(commnFileName) + ".MD5");
        aryLstFiles.Add(@clsBasFile.getPathWithoutExtn(commnFileName) + ".MD5");
        SharpZip zip = new SharpZip();
        zip.createZip(aryLstFiles, @clsBasFile.getPathWithoutExtn(commnFileName) + ".zip", "");
        DSstatement.Dispose();

    }

    protected string GetActiveQuery()
    {
        StringBuilder activeQuery = new StringBuilder();

        activeQuery.Append("select mt.cardproduct,count(distinct mt.accountno) stmtcount");
        activeQuery.Append(" from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " mt ");
        activeQuery.Append(" left join " + clsSessionValues.mainDbSchema + clsSessionValues.detailTable + " dt ");
        activeQuery.Append(" on dt.branch = mt.branch and DT.STATEMENTNO = MT.STATEMENTNO ");
        activeQuery.Append(" where mt.branch = " + pBranch);
        activeQuery.Append(" and mt.cardstate in ('New','Embossed','Given','Embossing','PIN generation','PIN generated','New Pin Generation Only','New Pin Generated Only','Entered') ");
        activeQuery.Append(" and length(mt.cardno) = 16 ");
        activeQuery.Append(" and ((dt.postingdate Is Null and dt.docno Is Null and mt.closingbalance <> 0 ) ");
        activeQuery.Append(" or (dt.postingdate Is not Null and dt.docno Is not Null )) ");
        activeQuery.Append(" group by mt.cardproduct order by mt.cardproduct ");

        return activeQuery.ToString();
    }

    protected string GetInactiveQuery()
    {
        StringBuilder inactiveQuery = new StringBuilder();
        inactiveQuery.Append(" select mt.cardproduct,count(distinct mt.accountno) stmtcount ");
        inactiveQuery.Append(" from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + @" mt left ");
        inactiveQuery.Append(" join " + clsSessionValues.mainDbSchema + clsSessionValues.detailTable + @" dt ");
        inactiveQuery.Append(" on dt.branch = mt.branch and DT.STATEMENTNO = MT.STATEMENTNO ");
        inactiveQuery.Append(" where mt.branch = " + pBranch);
        inactiveQuery.Append(" and mt.accountno not in (select accountno from ( ");
        inactiveQuery.Append(" select distinct mt.cardproduct, mt.accountno ");
        inactiveQuery.Append(" from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + @" mt left ");
        inactiveQuery.Append(" join " + clsSessionValues.mainDbSchema + clsSessionValues.detailTable + @" dt ");
        inactiveQuery.Append(" on dt.branch = mt.branch and DT.STATEMENTNO = MT.STATEMENTNO ");
        inactiveQuery.Append(" where mt.branch = " + pBranch);
        inactiveQuery.Append(" and mt.cardstate in ('New', 'Embossed', 'Given', 'Embossing', 'PIN generation', 'PIN generated', 'New Pin Generation Only', 'New Pin Generated Only', 'Entered') ");
        inactiveQuery.Append(" and length(mt.cardno) = 16 ");
        inactiveQuery.Append(" and((dt.postingdate Is Null and dt.docno Is Null and mt.closingbalance <> 0) ");
        inactiveQuery.Append(" or(dt.postingdate Is not Null and dt.docno Is not Null)))) ");
        inactiveQuery.Append(" and length(mt.cardno) = 16 ");
        inactiveQuery.Append(" and((dt.postingdate Is Null and dt.docno Is Null and mt.closingbalance <> 0) ");
        inactiveQuery.Append(" or(dt.postingdate Is not Null and dt.docno Is not Null)) ");
        inactiveQuery.Append(" group by mt.cardproduct order by mt.cardproduct ");

        return inactiveQuery.ToString();
    }

    public void GenerateStatementSummaryByActivity()
    {
        clsStatementSummary StatSummary = new clsStatementSummary();
        StatSummary.BankCode = curBranchVal;
        StatSummary.BankName = strBankName;
        StatSummary.StatementDate = vCurDate;
        StatSummary.CreationDate = DateTime.Now;
        DataRow[] activeRows, inactiveRows;
        totNoOfCardStat = 0;
        totNoOfCardEmailStat = 0;
        totNoOfTransactionsInt = 0;
        totNoOfTransEmailInt = 0;
        prevAccountNo = String.Empty;

        String activeQuery = GetActiveQuery();
        String inactiveQuery = GetInactiveQuery();

        conn = new OracleConnection(clsDbCon.sConOracle);
        OracleDataAdapter ActiveDA = new OracleDataAdapter(activeQuery, conn);
        OracleDataAdapter InactiveDA = new OracleDataAdapter(inactiveQuery, conn);

        conn.Open();
        //(new OracleCommand("alter session set optimizer_goal = RULE", conn)).ExecuteNonQuery();
        DataSet DScounts = new DataSet("MasterCountDS");
        ActiveDA.Fill(DScounts, "tStatementActiveCounts");
        InactiveDA.Fill(DScounts, "tStatementInactiveCounts");

        foreach (DataRow pRow in DSProducts.Tables["Products"].Rows)
        {
            ProductRow = pRow;
            activeRows = DScounts.Tables["tStatementActiveCounts"].Select("cardproduct = '" + ProductRow[pName].ToString().Trim() + "'");
            inactiveRows = DScounts.Tables["tStatementInactiveCounts"].Select("cardproduct = '" + ProductRow[pName].ToString().Trim() + "'");

            if (activeRows.Length != 0)
            {
                StatSummary.NoOfStatements = int.Parse(activeRows[0]["STMTCOUNT"].ToString());
                StatSummary.NoOfTransactionsInt = 0;
                StatSummary.ProductCode = int.Parse(ProductRow["code"].ToString());
                StatSummary.ProductName = ProductRow["name"].ToString() + "(Active)";
                StatSummary.InsertRecordDb(vCurDate.ToString("yyyyMM") + strBankName);
            }
            if (inactiveRows.Length != 0)
            {
                StatSummary.NoOfStatements = int.Parse(inactiveRows[0]["STMTCOUNT"].ToString());
                StatSummary.NoOfTransactionsInt = 0;
                StatSummary.ProductCode = int.Parse(ProductRow["code"].ToString());
                StatSummary.ProductName = ProductRow["name"].ToString() + "(Inactive)";
                StatSummary.InsertRecordDb(vCurDate.ToString("yyyyMM") + strBankName);
            }
        }
        StatSummary = null;
    }

    public void GenerateStatementSummary()
    {
        clsStatementSummary StatSummary = new clsStatementSummary();
        StatSummary.BankCode = curBranchVal;
        StatSummary.BankName = strBankName;
        StatSummary.StatementDate = vCurDate;
        StatSummary.CreationDate = DateTime.Now;
        DataRow[] productRows;
        totNoOfCardStat = 0;
        totNoOfCardEmailStat = 0;
        totNoOfTransactionsInt = 0;
        totNoOfTransEmailInt = 0;
        prevAccountNo = String.Empty;
        foreach (DataRow pRow in DSProducts.Tables["Products"].Rows)
        {
            ProductRow = pRow;
            productRows = DSstatement.Tables["tStatementMasterTable"].Select("cardproduct = '" + ProductRow[pName].ToString().Trim() + "'");
            if (productRows.Length == 0)
            {
                continue;
            }
            foreach (DataRow mRow in productRows)
            {
                masterRow = mRow;
                cardsRows = mRow.GetChildRows(StatementNoDRel); //, DataRowVersion.Proposed
                strCardNo = masterRow[mCardno].ToString().Trim();
                if (strCardNo.Length != 16)
                {
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
                    if (isInEmailService)
                    {
                        if (haveEmail(masterRow[mClientid].ToString()))
                            isEmailStat = true;
                        else
                            isEmailStat = false;
                    }
                    if (prevAccountNo == string.Empty)
                        curMainCard = string.Empty;
                    pageNo = 1;//totalAccPages = 1 ; pageNo=1;
                    CurPageRec4Dtl = 0; totNetUsage = 0; CurrentPageFlag = "F 1"; //if page is based on account no
                    calcAccountRows();
                    if (!isDebit)
                        if (totAccRows < 1
                        && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == Convert.ToDecimal("0")) //             || (masterRow[mCardno].ToString() == curMainCard   // Convert.ToDecimal(
                        {
                            continue;
                        }
                    prevAccountNo = masterRow[accountNoName].ToString();
                    if (isEmailStat)
                        totNoOfCardEmailStat++;
                    else
                        totNoOfCardStat++;
                } // End of if(prevAccountNo != masterRow[accountNoName].ToString())
                foreach (DataRow dRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
                {
                    detailRow = dRow;
                    stmNo = detailRow[dStatementno].ToString();
                    if ((detailRow[dPostingdate] == DBNull.Value) && (detailRow[dDocno] == DBNull.Value)) continue;// Exclude On-Hold Transactions 
                    curAccRows++;
                    if (CurPageRec4Dtl >= MaxDetailInPage)
                    {
                        CurPageRec4Dtl = 0;
                        pageNo++;
                    }
                    CurPageRec4Dtl = CurPageRec4Dtl + 1;
                    hasInterset = CalcTransInterest(detailRow[dTrandescription].ToString());
                    if (hasInterset)
                    {
                        if (isEmailStat)
                            totNoOfTransEmailInt++;
                        else
                            totNoOfTransactionsInt++;
                    }
                    //StatSummaryDetail.InsertRecordDb(vCurDate.ToString("yyyyMM") + strBankName);
                } //end of detail foreach
                curCrdNoInAcc++;
                if ((curAccRows >= totAccRows && totAccRows != 0) || (totAccRows == 0 && curCrdNoInAcc == totCrdNoInAcc))
                {
                    CurPageRec4Dtl = 0; //>pageNo=1; if pages is based on account
                    curAccRows = 0;
                }
            }
            if ((totNoOfCardStat + totNoOfCardEmailStat) != 0)
            {
                StatSummary.NoOfStatements = totNoOfCardStat + totNoOfCardEmailStat;
                StatSummary.NoOfTransactionsInt = totNoOfTransactionsInt + totNoOfTransEmailInt;
                StatSummary.ProductCode = int.Parse(ProductRow["code"].ToString());
                StatSummary.ProductName = ProductRow["name"].ToString();
                StatSummary.InsertRecordDb(vCurDate.ToString("yyyyMM") + strBankName);
                totNoOfCardStat = 0;
                totNoOfCardEmailStat = 0;
                totNoOfTransactionsInt = 0;
                totNoOfTransEmailInt = 0;
                prevAccountNo = String.Empty;
            }
        }
        StatSummary = null;
    }


    public void GenerateStatementLog()
    {
        clsStatementSummaryDetail StatSummaryDetail = new clsStatementSummaryDetail();
        StatSummaryDetail.BankCode = curBranchVal;
        StatSummaryDetail.AccountNo = masterRow[accountNoName].ToString();
        StatSummaryDetail.AccountCurrency = masterRow[mAccountcurrency].ToString();
        StatSummaryDetail.ExternalAccount = masterRow[mExternalno].ToString();
        StatSummaryDetail.BilingDate = System.DateTime.Now;
        StatSummaryDetail.GenerationDate = System.DateTime.Now;
        StatSummaryDetail.OpeningBalance = Convert.ToInt32(masterRow[mOpeningbalance]);
        StatSummaryDetail.ClosingBalance = Convert.ToInt32(masterRow[mClosingbalance]);
        StatSummaryDetail.MinimumAmount = Convert.ToInt32(masterRow[mMindueamount]);
        StatSummaryDetail.DueDate = DateTime.Parse(masterRow[mStetementduedate].ToString());
        StatSummaryDetail.OverDueAmount = Convert.ToInt32(masterRow[mTotaloverdueamount]);
        StatSummaryDetail.OverDueDays = 0;
        StatSummaryDetail.Interset = Convert.ToInt32(masterRow[mTotalinterest]);
        StatSummaryDetail.TotalDebits = Convert.ToInt32(masterRow[mTotaldebits]);
        StatSummaryDetail.TotalCredits = Convert.ToInt32(masterRow[mTotalcredits]);
        StatSummaryDetail.TotalPayements = Convert.ToInt32(masterRow[mTotalpayments]);
        StatSummaryDetail.TotalCash = Convert.ToInt32(masterRow[mTotalcashwithdrawal]);
        StatSummaryDetail.TotalRetail = Convert.ToInt32(masterRow[mTotalpurchases]);
        StatSummaryDetail.TotalFees = Convert.ToInt32(masterRow[mTotalcharges]);
        StatSummaryDetail.TotalOthers = 0;
        StatSummaryDetail.TotalPoints = 0;
        StatSummaryDetail.TotalRedeemedPoints = 0;
        StatSummaryDetail.TotalExpiredPoints = 0;
        StatSummaryDetail.CreditLimit = Convert.ToInt32(masterRow[mContractlimit]);
        StatSummaryDetail.InsertRecordDb();
        StatSummaryDetail = null;
    }

    public string mainSortOrder
    {
        set { strOrder = value; }
    }// mainSortOrder

    public bool emailService
    {
        get { return isInEmailService; }
        set { isInEmailService = value; }
    }// emailService

    public string bankName
    {
        get { return strBankName; }
        set { strBankName = value; }
    }// bankName

    public string rewardCondition
    {
        get { return rewardCond; }
        set { rewardCond = value; }
    }// rewardCondition

    public bool isRewardVal
    {
        get { return isReward; }
        set { isReward = value; }
    }// isRewardVal

    public bool isDebitVal
    {
        get { return isDebit; }
        set { isDebit = value; }
    }// isDebitVal

    public frmStatementFile setFrm
    {
        set { frmMain = value; }
    }// setFrm

    public bool CreateCorporate
    {
        get { return createCorporateVal; }
        set { createCorporateVal = value; }
    }// CreateCorporate

    public string productCond
    {
        get { return strProductCond; }
        set { strProductCond = value; }
    }  // productCond

    public string ExcludeCond
    {
        get { return strExcludeCond; }
        set { strExcludeCond = value; }
    }  // ExcludeCond

    public bool isFiltered
    {
        get { return isFilteredVal; }
        set { isFilteredVal = value; }
    }// isFiltered

    public bool isExcluded
    {
        get { return isExcludedVal; }
        set { isExcludedVal = value; }
    }// isExcluded

    public string PaymentSystem
    {
        get { return strPaymentSystem; }
        set { strPaymentSystem = value; }
    }  // PaymentSystem

    public string BillingCycle
    {
        get { return strBillingCycle; }
        set { strBillingCycle = value; }
    }  // BillingCycle

    public bool isSplitted
    {
        get { return isSplittedVal; }
        set { isSplittedVal = value; }
    }// isSplitted

    ~clsStatTxt()
    {
        DSstatement.Dispose();
    }
}
