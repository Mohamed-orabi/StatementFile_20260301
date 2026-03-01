using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Xml;
using System.Collections;

// Branch 16
public class clsStatTxtLblSupTot : clsBasStatement
    {
    protected string strBankName;
    protected FileStream fileStrm, fileStrmByEmail, fileSummary, fileStrmSupp;
    protected StreamWriter streamWrit, strmWriteByEmail, strmWriteCommon, streamSummary, strmWriteSupp;
    //protected DataSet DSstatement;
    //protected OracleDataReader drPrimaryCards, drMaster,drDetail;
    protected DataRow masterRow;
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
    protected DataRow[] cardsRows, accountRows, rewardRows, tmpDtlRows;
    protected DataRow[] mainRows;
    protected DataRow rewardRow;
    protected string CurrentPageFlag;
    protected string strCardNo, strPrimaryCardNo;
    protected string strForeignCurr;
    protected string stmNo;
    protected int totNoOfCardStat, totNoOfPageStat, totNoOfCardEmailStat, totNoOfPageEmailStat, totNoOfEmailStat;

    protected int curCardRow = 0, curTotCardRows = 0, numOfErr = 0, totNoOfTransactions = 0, totNoOfTransEmail = 0, totNoOfTransactionsInt = 0, totNoOfTransEmailInt = 0;
    protected bool isPrimaryOnly, isHaveF3 = true, isPrimaryOpened = false;
    protected FileStream fileStrmErr;
    protected StreamWriter strmWriteErr;
    protected string curMainCard, accMainCard;

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
    protected string rewardCond = "'Reward Program'";//'New Reward Contract'
    protected string vipCond = string.Empty;//'New Reward Contract'
    protected DataRow[] emailRows = null;
    protected string emailTo = string.Empty;
    protected bool isEmailStat = false;
    clsValidateEmail valdEmail = new clsValidateEmail();
    protected string strFileNam, stmntType;
    protected string curSuppCard = string.Empty;
    protected decimal totSuppCrdVal = 0;
    protected bool isActiveSuppCrd = false;
    protected bool isReward = false;
    protected bool isDebit = false;
    private DataRow ProductRow;
    protected bool hasInterset = false;
    protected string productCond = string.Empty;
    private decimal totTrans = 0;

    public clsStatTxtLblSupTot()
        {
        }

    public string Statement(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pCurDate, string pStmntType, bool pAppendData)
        {
        string rtrnStr = "Successfully Generate " + pBankName;
        int curMonth = pCurDate.Month;
        string comnFileName = string.Empty;
        bool preExit = true;
        strFileNam = pStrFileName;
        stmntType = pStmntType;

        try
            {
            //clsMaintainData maintainData = new clsMaintainData();
            //if (isReward)
            //{
            //    maintainData.notRward = false;
            //    maintainData.curRewardCond = rewardCond;
            //    maintainData.fixReward(pBankCode, rewardCond);
            //}
            //maintainData.matchCardBranch4Account(pBankCode);

            pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);
            vCurDate = pCurDate; //DateTime.Now.AddMonths(-1);
            strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
            clsBasFile.createDirectory(pStrFileName + vCurDate.ToString("yyyyMM") + pBankName);
            pStrFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM") + ".txt";
            strBankName = pBankName;

            strOutputFile = pStrFileName;
            comnFileName = clsBasFile.getPathWithoutExtn(pStrFileName) + "_NotHaveEmail." + clsBasFile.getFileExtn(pStrFileName);
            // open output file for Not Have Email
            fileStrm = new FileStream(comnFileName, FileMode.Create); //Create
            streamWrit = new StreamWriter(fileStrm, Encoding.Default);
            streamWrit.AutoFlush = true;

            // By Email file
            fileStrmByEmail = new FileStream(clsBasFile.getPathWithoutExtn(pStrFileName) + "_ByEmail." + clsBasFile.getFileExtn(pStrFileName), FileMode.Create);
            strmWriteByEmail = new StreamWriter(fileStrmByEmail, Encoding.Default);
            strmWriteByEmail.AutoFlush = true;

            // for Supplementary cards
            fileStrmSupp = new FileStream(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Supplementary." + clsBasFile.getFileExtn(pStrFileName), FileMode.Create);
            strmWriteSupp = new StreamWriter(fileStrmSupp, Encoding.Default);
            strmWriteSupp.AutoFlush = true;

            // Error file
            fileStrmErr = new FileStream(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName), FileMode.Create);
            strmWriteErr = new StreamWriter(fileStrmErr, Encoding.Default);
            strmWriteErr.AutoFlush = true;

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
            if (isRewardVal)
                {
                curRewardCond = rewardCond;
                //strMainTableCond = "m.contracttype != " + rewardCond;
                //strSubTableCond = "d.trandescription != 'Calculated Points'";
                strMainTableCond = "m.contracttype not in " + rewardCond;
                strSubTableCond = "d.trandescription not in ('Calculated Points')";
                getReward(pBankCode);
                }

            isGetTotal = true;
            if (pBankName == "DBN_Credit")
                {
                strMainTableCond += " and m.contracttype not in " + VIPCondition + " and + m.cardproduct != '" + ProductCondition + "'";
                }
            else if (pBankName == "DBN_Credit_VIP1-2-5")
                {
                strMainTableCond += " and m.contracttype in " + VIPCondition;
                strOrder = "m.contractno,m.cardproduct,m.CARDBRANCHPART,m.accountno,m.cardprimary desc,m.cardno";
                }
            else if (pBankName == "DBN_Credit_VIP")
                {
                strMainTableCond += " and m.contracttype in " + VIPCondition;
                }
            else if (pBankName == "DBN_Credit_ParkNShop")
                {
                strMainTableCond += " and m.contracttype in " + VIPCondition;
                strSubTableCond = "d.trandescription not in ('NSHOP_BONUSES')";
                }
            else if (pBankName == "DBN_Credit_EXCO_VIP")
                {
                strMainTableCond += " and m.contracttype in " + VIPCondition + " and + m.cardproduct = '" + ProductCondition + "'";
                }
            else if (pBankName == "DBN_MasterCard_Credit")
                {
                strMainTableCond += " and m.contracttype in " + VIPCondition;
                }


            // data retrieve
            FillStatementDataSet(pBankCode, "vip"); //DSstatement =  //16); //3
            getClientEmail(pBankCode);
            getCardProduct(pBankCode);
            pageNo = 1; totalCardPages = 0;
            curCardNo = String.Empty;
            curAccountNo = String.Empty;

            frmMain.BeginInvoke(frmMain.setMinMaxProgressDelegate, new object[] { DSstatement.Tables["tStatementMasterTable"].Rows.Count });
            foreach (DataRow mRow in DSstatement.Tables["tStatementMasterTable"].Rows)
                {
                frmMain.BeginInvoke(frmMain.setProgressDelegate, new object[] { totRec++ });
                masterRow = mRow;
                //strmWriteCommon.WriteLine(masterRow[mStatementno].ToString());
                //pageNo=1; CurPageRec4Dtl=0; CurrentPageFlag = "F 1"; //if page is based on card no
                if (masterRow[mCardprimary].ToString() == "Y") //select all transactions for the account
                    //cardsRows = DSstatement.Tables["tStatementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[accountNoName]).ToString().Trim() + "'", "cardprimary desc, cardno, postingdate");
                    cardsRows = DSstatement.Tables["tStatementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[accountNoName]).ToString().Trim() + "'", "cardno, postingdate");
                else
                    cardsRows = mRow.GetChildRows(StatementNoDRel); //, DataRowVersion.Proposed

                strCardNo = masterRow[mCardno].ToString().Trim();

                if (masterRow[mHOLSTMT].ToString() == "Y")
                    continue;
                if (strCardNo.Length != 16)
                    {
                    strmWriteErr.WriteLine("CardNo Length not equal 16 for AccountNo " + masterRow[accountNoName].ToString().Trim() + " Card No =" + strCardNo);
                    //numOfErr++;
                    continue;// Exclude Zero Length Cards 
                    }

                strPrimaryCardNo = strCardNo;
                //-if (masterRow[mCardprimary].ToString() == "N")
                //-{
                //-strPrimaryCardNo = masterRow[mPrinarycardno].ToString();
                strPrimaryCardNo = masterRow[mCardno].ToString().Trim();
                //calcCardlRows();
                //-}

                //start new account
                //-if (prevAccountNo != masterRow[accountNoName].ToString())
                //-{
                streamWrit.Flush();
                strmWriteByEmail.Flush();
                if (masterRow[mCardprimary].ToString() == "Y")
                    {
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
                else
                    {
                    strmWriteCommon = strmWriteSupp;
                    isEmailStat = false;
                    }
                strmWriteCommon.AutoFlush = true;
                //if (pageNo != totalAccPages && prevAccountNo != "")// 
                //{
                //  //>MessageBox.Show( "Error in Genrating Statement", "Please Call The Programmer", MessageBoxButtons.OK ,
                //  //>  MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1); //,	MessageBoxOptions.RightAlign
                //  strmWriteErr.WriteLine("pageNo not equal totalAccPages : " + prevAccountNo);
                //  numOfErr++;
                //}
                if (prevAccountNo == string.Empty)
                    if (curMonth != Convert.ToDateTime(masterRow[mStatementdateto]).Month)//,"dd/MM/yyyy"
                        {
                        preExit = false;
                        fileStrm.Close();
                        fileStrmErr.Close();
                        clsBasFile.deleteFile(@strOutputFile);
                        clsBasFile.deleteFile(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName));
                        return "Error in Generation " + pBankName;
                        }

                curMainCard = string.Empty;
                if (!isHaveF3)//!isHaveF3  CurrentPageFlag != "F 2"
                    {
                    strmWriteErr.WriteLine("Account not Closed : " + prevAccountNo);
                    numOfErr++;
                    }
                isHaveF3 = false;

                pageNo = 1;//totalAccPages = 1 ; pageNo=1;
                CurPageRec4Dtl = 0; totNetUsage = 0; CurrentPageFlag = "F 1"; totTrans = 0; //if page is based on account no
                calcAccountRows();

                //(masterRow[mCardprimary].ToString().ToUpper() == "Y" && ((cardsRows.Length < 1 && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0) || masterRow[mCardno].ToString().ToUpper() != curMainCard))
                //&& (masterRow[mCardprimary].ToString().ToUpper() == "Y" && masterRow[mCardno].ToString().ToUpper() != curMainCard)
                //if (((totAccRows < 1 && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0 && masterRow[mCardprimary].ToString().ToUpper() == "Y")
                //  || (totAccRows < 1 && masterRow[mCardprimary].ToString().ToUpper() == "N")) || !isValidateCard(masterRow[mCardstate].ToString())) //             || (masterRow[mCardno].ToString() == curMainCard   // Convert.ToDecimal(
                //if(masterRow[mCardprimary].ToString().ToUpper() == "Y" && masterRow[mCardno].ToString().ToUpper() == curMainCard)
                //{
                //}
                //else
                //{
                //}
                if (!isDebit)
                    if ((
                      ((cardsRows.Length < 1 && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0 && masterRow[mCardprimary].ToString().ToUpper() == "Y")
                      || (totAccRows < 1 && masterRow[mCardprimary].ToString().ToUpper() == "N"))) || !isValidateCard(masterRow[mCardstate].ToString()) && (masterRow[mCardprimary].ToString().ToUpper() == "Y" && masterRow[mCardno].ToString().ToUpper() != curMainCard)) //             || (masterRow[mCardno].ToString() == curMainCard   // Convert.ToDecimal(
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
                //-} // End of if(prevAccountNo != masterRow[accountNoName].ToString())
                //calcCardlRows();
                //if(totCardRows < 1)continue ;  //if pages is based on card
                //printHeader();//if pages is based on card

                foreach (DataRow dRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
                    {
                    detailRow = dRow;
                    stmNo = detailRow[dStatementno].ToString();
                    if ((detailRow[dPostingdate] == DBNull.Value) && (detailRow[dDocno] == DBNull.Value)) continue;// Exclude On-Hold Transactions 
                    //if (detailRow[mCardprimary].ToString().ToUpper() != "Y" && curSuppCard != detailRow[dCardno].ToString())
                    if (masterRow[mCardprimary].ToString().ToUpper() != "Y" && curSuppCard != masterRow[dCardno].ToString())
                        {
                        tmpDtlRows = DSstatement.Tables["tStatementMasterTable"].Select(mCardno + " = '" + detailRow[dCardno] + "'");
                        if (tmpDtlRows.Length > 0)
                            {
                            if (tmpDtlRows[0][mCardprimary].ToString() == "N" && isValidateCard(tmpDtlRows[0][mCardstate].ToString()))
                                {
                                totSuppCrdVal = 0;
                                if (isActiveSuppCrd)
                                    {
                                    curCrdNoInAcc++;
                                    curAccRows++;
                                    preCheckDetail();
                                    CurPageRec4Dtl++;
                                    strmWriteCommon.WriteLine(basText.replicats(" ", 42) + "<< Total Supplementary Card : " + basText.formatCardNumber(curSuppCard) + basText.alignmentRight(basText.formatNumUnSign(totSuppCrdVal, "#0.00", 16) + " " + CrDb(Convert.ToDecimal(totSuppCrdVal)), 28));
                                    //strmWriteCommon.WriteLine(basText.replicats(" ", 40) + "<< Total Supplementary Card : " + basText.formatCardNumber(curSuppCard) + basText.alignmentRight(basText.formatNumUnSign(totSuppCrdVal, "#0.00", 16) + " " + CrDb(Convert.ToDecimal(totSuppCrdVal)), 28));
                                    }
                                curCrdNoInAcc++;
                                curAccRows++;
                                preCheckDetail();
                                CurPageRec4Dtl++;
                                isActiveSuppCrd = true;
                                strmWriteCommon.WriteLine(basText.replicats(" ", 42) + ">> Supplementary Card : " + basText.formatCardNumber(detailRow[dCardno].ToString()));
                                //strmWriteCommon.WriteLine(basText.replicats(" ", 40) + ">> Supplementary Card : " + basText.formatCardNumber(detailRow[dCardno].ToString()));
                                }
                            }
                        }
                    curAccRows++;
                    preCheckDetail();
                    //						totNetUsage +=clsBasValid.validateNum(detailRow[dBilltranamount]);
                    totNetUsage = calculateCrDb(totNetUsage, clsBasValid.validateNum(detailRow[dBilltranamount]), clsBasValid.validateStr(detailRow[dBilltranamountsign]));
                    CurPageRec4Dtl++;
                    curSuppCard = detailRow[dCardno].ToString();
                    printDetail();
                    //if(!((detailRow[dPostingdate]== DBNull.Value) && (detailRow[dDocno]== DBNull.Value))) 

                    } //end of detail foreach
                //printCardFooter();//if pages is based on card
                // if(masterRow[mCardprimary].ToString() == "Y" && (masterRow[mCardstate].ToString() == "Given" || masterRow[mCardstate].ToString() == "Embossed" || masterRow[mCardstate].ToString() == "New"))
                if (isActiveSuppCrd)
                    {
                    preCheckDetail();
                    strmWriteCommon.WriteLine(basText.replicats(" ", 42) + "<< Total Supplementary Card : " + basText.formatCardNumber(curSuppCard) + basText.alignmentRight(basText.formatNumUnSign(totSuppCrdVal, "#0.00", 16) + " " + CrDb(Convert.ToDecimal(totSuppCrdVal)), 28));
                    //strmWriteCommon.WriteLine(basText.replicats(" ", 40) + "<< Total Supplementary Card : " + basText.formatCardNumber(curSuppCard) + basText.alignmentRight(basText.formatNumUnSign(totSuppCrdVal, "#0.00", 16) + " " + CrDb(Convert.ToDecimal(totSuppCrdVal)), 28));
                    CurPageRec4Dtl++;
                    curCrdNoInAcc++;
                    curAccRows++;
                    }
                totSuppCrdVal = 0;
                isActiveSuppCrd = false;
                curCrdNoInAcc++;
                //if ((curAccRows >= totAccRows && totAccRows != 0) || (totAccRows == 0 && curCrdNoInAcc == totCrdNoInAcc))
                if (curAccRows >= totAccRows)
                    {
                    completePageDetailRecords();
                    printCardFooter();//if pages is based on account
                    //printAccountFooter();
                    CurPageRec4Dtl = 0; //>pageNo=1; if pages is based on account
                    curAccRows = 0;
                    }
                //strmWriteCommon.WriteLine(strEndOfPage);
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
                streamWrit.Flush();
                streamWrit.Close();
                fileStrm.Close();

                printStatementSummary();

                // Close output File
                strmWriteByEmail.Flush();
                strmWriteByEmail.Close();
                fileStrmByEmail.Close();

                // Close output File for Supplementary cards
                strmWriteSupp.Flush();
                strmWriteSupp.Close();
                fileStrmSupp.Close();

                // Close Summary File
                streamSummary.Flush();
                streamSummary.Close();
                fileSummary.Close();

                strmWriteErr.Flush();
                strmWriteErr.Close();
                fileStrmErr.Close();
                ArrayList aryLstFiles = new ArrayList();
                if (numOfErr == 0)
                    clsBasFile.deleteFile(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName));
                else
                    aryLstFiles.Add(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err.txt");
                aryLstFiles.Add(comnFileName);
                aryLstFiles.Add(clsBasFile.getPathWithoutExtn(pStrFileName) + "_ByEmail." + clsBasFile.getFileExtn(pStrFileName));
                aryLstFiles.Add(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Supplementary." + clsBasFile.getFileExtn(pStrFileName));
                numOfErr = validateNoOfLines(aryLstFiles, 60);
                if (numOfErr > 0)
                    rtrnStr += string.Format(" with {0} Errors", numOfErr);
                aryLstFiles.Add(@fileSummaryName);
                clsBasFile.generateFileMD5(aryLstFiles, @clsBasFile.getPathWithoutExtn(pStrFileName) + ".MD5");
                aryLstFiles.Add(@clsBasFile.getPathWithoutExtn(pStrFileName) + ".MD5");
                SharpZip zip = new SharpZip();
                zip.createZip(aryLstFiles, @clsBasFile.getPathWithoutExtn(pStrFileName) + ".zip", "");

                DSstatement.Dispose();
                }
            }
        return rtrnStr;
        }


    protected virtual void printHeader()
        {
        string newaddress1, newaddress2;
        clsMaintainData.fixAddress(curBranchVal, masterRow[mCustomeraddress1].ToString(), out newaddress1, out newaddress2);
        extAccNum = clsBasValid.validateStr(masterRow[mExternalno]);
        //DBN-6861
        if (extAccNum.Trim() != "")
            extAccNum = "******" + extAccNum.Substring(6);
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
            }

        if (CurrentPageFlag == "F 0" || CurrentPageFlag == "F 1")
            {
            strmWriteCommon.WriteLine(strEndOfPage + strEndOfAccount);
            }
        else
            {
            strmWriteCommon.WriteLine(strEndOfPage);
            }
        strmWriteCommon.WriteLine(String.Empty); //(basText.replicat(" ", 81) + CurrentPageFlag);  //+"{0,5:MM/yy}",masterRow[mStatementdatefrom]
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        //		strmWriteCommon.WriteLine(basText.replicat(" ",81) + masterRow[mCardproduct]);  //+"{0,5:MM/yy}",masterRow[mStatementdatefrom]
        strmWriteCommon.WriteLine(basText.alignmentMiddle("Customer Name & Address", 50) + basText.replicat(" ", 5) + basText.alignmentRight("Card Product :", 25) + " " + masterRow[mCardproduct].ToString());  //+"{0,5:MM/yy}",masterRow[mStatementdatefrom]
        strmWriteCommon.WriteLine(basText.alignmentMiddle(masterRow[mCustomername], 50));  //
        //		strmWriteCommon.WriteLine( basText.replicat(" ",81) + masterRow[mCardbranchpartname]);  //
        //>>strmWriteCommon.WriteLine(basText.replicat(" ", 55) + basText.alignmentRight("Branch :", 25) + " " + masterRow[mCardbranchpartname].ToString());  //
        strmWriteCommon.WriteLine((masterRow[mCardprimary].ToString() == "Y" ? basText.replicat(" ", 51) : basText.alignmentMiddle(masterRow[mCardclientname], 50)) + basText.replicat(" ", 5) + basText.alignmentRight("Branch :", 25) + " " + masterRow[mCardbranchpartname].ToString());  //
        //strmWriteCommon.WriteLine(basText.alignmentMiddle(ValidateArbic(masterRow[mCustomeraddress1].ToString()), 50)); //basText.formatNum(masterRow[mCardno],"####-####-####-####")
        strmWriteCommon.WriteLine(basText.alignmentMiddle(ValidateArbic(newaddress1), 50)); //basText.formatNum(masterRow[mCardno],"####-####-####-####")
        //		strmWriteCommon.WriteLine(basText.replicat(" ",81) + extAccNum); //clsBasValid.validateStr(masterRow[mExternalno]) //accountno  //basText.formatNum(masterRow[mCardno],"####-####-####-####")
        strmWriteCommon.WriteLine(basText.replicat(" ", 55) + basText.alignmentRight("Account Number :", 25) + " " + extAccNum); //clsBasValid.validateStr(masterRow[mExternalno]) //accountno  //basText.formatNum(masterRow[mCardno],"####-####-####-####")
        //strmWriteCommon.WriteLine(basText.alignmentMiddle(ValidateArbic(masterRow[mCustomeraddress2].ToString()), 50));
        strmWriteCommon.WriteLine(basText.alignmentMiddle(ValidateArbic(newaddress2), 50));
        //		strmWriteCommon.WriteLine(basText.replicat(" ",81)+"{0,10:dd/MM/yyyy}",masterRow[mStatementdateto]);//Statementdatefrom+ basText.formatNum(masterRow[accountLimit],"########")  basText.formatNum(masterRow[mCardlimit],"########")
        strmWriteCommon.WriteLine(basText.replicat(" ", 55) + basText.alignmentRight("Statement Date :", 25) + " " + "{0,10:dd/MM/yyyy}", masterRow[mStatementdateto]);//Statementdatefrom+ basText.formatNum(masterRow[accountLimit],"########")  basText.formatNum(masterRow[mCardlimit],"########")
        strmWriteCommon.WriteLine(basText.alignmentMiddle(ValidateArbic(masterRow[mCustomeraddress3].ToString().Trim()) + " " + masterRow[mCustomerregion].ToString().Trim() + " " + masterRow[mCustomercity].ToString().Trim(), 50));  //basText.formatNum(masterRow[accountNoName],"##-##-######-####")
        //		strmWriteCommon.WriteLine(basText.replicat(" ",81) + pageNo + " / " + totalAccPages);  //basText.formatNum(masterRow[accountNoName],"##-##-######-####")
        strmWriteCommon.WriteLine(basText.replicat(" ", 55) + basText.alignmentRight("Page :", 25) + " " + pageNo + " / " + totalAccPages);  //basText.formatNum(masterRow[accountNoName],"##-##-######-####")
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        //		strmWriteCommon.WriteLine(String.Empty);
        //		strmWriteCommon.WriteLine(basText.alignmentMiddle(curMainCard,20) + basText.alignmentMiddle(masterRow[accountLimit],13) + basText.formatNum(masterRow[accountAvailableLimit],"##0",20) +  basText.alignmentMiddle(masterRow[mMindueamount],13) + basText.formatDate(masterRow[mStetementduedate],"dd/MM/yyyy",15,"M") + basText.formatNum(masterRow[mTotaloverdueamount],"##0.00",13)) ; //+ basText.formatNum(masterRow[accountLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mTotalcredits],"#,##0.00",9)+" "+basText.formatNum(masterRow[accountAvailableLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mMindueamount],"#,##0.00",15)+" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",15)+" "+basText.formatDate(masterRow[mStetementduedate],"dd/MM/yy") //TOTALCREDITS// basText.alignmentMiddle(masterRow[accountAvailableLimit],20) //"#,##0"
        strmWriteCommon.WriteLine(basText.alignmentMiddle("Primary Card No", 20) + basText.alignmentMiddle("Credit Limit", 13) + basText.alignmentMiddle("Available Limit", 20) + basText.alignmentMiddle("Minimum Due", 13) + basText.alignmentMiddle("Due Date", 15) + basText.alignmentMiddle("Over Due Amount", 15)); //+ basText.formatNum(masterRow[accountLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mTotalcredits],"#,##0.00",9)+" "+basText.formatNum(masterRow[accountAvailableLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mMindueamount],"#,##0.00",15)+" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",15)+" "+basText.formatDate(masterRow[mStetementduedate],"dd/MM/yy") //TOTALCREDITS// basText.alignmentMiddle(masterRow[accountAvailableLimit],20) //"#,##0"
        //strmWriteCommon.WriteLine(basText.alignmentMiddle(basText.formatCardNumber(curMainCard), 20) + basText.alignmentMiddle(masterRow[accountLimit], 13) + basText.formatNum(masterRow[accountAvailableLimit], "##0", 20, "M") + basText.alignmentMiddle(masterRow[mMindueamount], 13) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M") + basText.formatNum(masterRow[mTotaloverdueamount], "##0.00", 13, "M")); //+ basText.formatNum(masterRow[accountLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mTotalcredits],"#,##0.00",9)+" "+basText.formatNum(masterRow[accountAvailableLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mMindueamount],"#,##0.00",15)+" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",15)+" "+basText.formatDate(masterRow[mStetementduedate],"dd/MM/yy") //TOTALCREDITS// basText.alignmentMiddle(masterRow[accountAvailableLimit],20) //"#,##0"
        //DBN-5419 => EDT-965
        //strmWriteCommon.WriteLine(basText.alignmentMiddle(basText.formatCardNumber(curMainCard), 20) + basText.alignmentMiddle(masterRow[mCardlimit], 13) + basText.formatNum(masterRow[accountAvailableLimit], "##0", 20, "M") + basText.alignmentMiddle(masterRow[mMindueamount], 13) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M") + basText.formatNum(masterRow[mTotaloverdueamount], "##0.00", 13, "M")); //+ basText.formatNum(masterRow[accountLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mTotalcredits],"#,##0.00",9)+" "+basText.formatNum(masterRow[accountAvailableLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mMindueamount],"#,##0.00",15)+" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",15)+" "+basText.formatDate(masterRow[mStetementduedate],"dd/MM/yy") //TOTALCREDITS// basText.alignmentMiddle(masterRow[accountAvailableLimit],20) //"#,##0"
        strmWriteCommon.WriteLine(basText.alignmentMiddle(basText.formatCardNumber(curMainCard), 20) + basText.alignmentMiddle(masterRow[mCardlimit], 13) + basText.formatNum(masterRow[accountAvailableLimit], "##0", 20, "M") + basText.alignmentMiddle(masterRow[mCarddafamount], 13) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M") + basText.formatNum(masterRow[mTotaloverdueamount], "##0.00", 13, "M")); //+ basText.formatNum(masterRow[accountLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mTotalcredits],"#,##0.00",9)+" "+basText.formatNum(masterRow[accountAvailableLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mMindueamount],"#,##0.00",15)+" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",15)+" "+basText.formatDate(masterRow[mStetementduedate],"dd/MM/yy") //TOTALCREDITS// basText.alignmentMiddle(masterRow[accountAvailableLimit],20) //"#,##0"
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        //		strmWriteCommon.WriteLine(String.Empty);
        //strmWriteCommon.WriteLine(" {0:dd/MM} {1:dd/MM} {2,-24}  {3,-57} {4,18}", "T Date", "P Date", basText.trimStr("Reference No", 20), basText.trimStr("Description", 57), basText.alignmentMiddle("Amount", 23)); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)
        strmWriteCommon.WriteLine(" {0:dd/MM} {1:dd/MM} {2,-24}  {3,-57} {4,18} {5,18}", "T Date", "P Date", basText.trimStr("Reference No", 20), basText.trimStr("Description", 57), basText.alignmentMiddle("Amount", 23), basText.alignmentMiddle("Current Balance", 23));
        if (pageNo == 1)
            {
            strmWriteCommon.WriteLine(basText.replicat(" ", 43) + basText.alignmentLeft("Previous Balance", 67) + basText.formatNumUnSign(masterRow[mOpeningbalance], "##0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance])));  //+ CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign]))
            totTrans = Convert.ToDecimal(masterRow[mOpeningbalance]);
            }
        //strmWriteCommon.WriteLine(basText.replicat(" ", 43) + basText.alignmentLeft("Previous Balance", 67) + basText.formatNumUnSign(masterRow[mOpeningbalance], "##0.00", 16) + "                       " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance])) + basText.formatNumUnSign(masterRow[mOpeningbalance], "##0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance])));
        else
            strmWriteCommon.WriteLine(String.Empty);

        strmWriteCommon.WriteLine(String.Empty);
        totalPages++;
        if (isEmailStat)
            totNoOfPageEmailStat++;
        else
            totNoOfPageStat++;

        //sCurrent_balance = Convert.ToDecimal(masterRow[mOpeningbalance]);

        }

    protected void preCheckDetail()
        {
        if (CurPageRec4Dtl >= MaxDetailInPage)
            {
            CurPageRec4Dtl = 0;
            printCardFooter();
            pageNo++;
            printHeader();
            }
        }

    protected virtual void printDetail()
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
        totSuppCrdVal += valDbCr(Convert.ToDecimal(detailRow[dBilltranamount].ToString()), clsBasValid.validateStr(detailRow[dBilltranamountsign]));

        totTrans += valDbCr(Convert.ToDecimal(detailRow[dBilltranamount]), clsBasValid.validateStr(detailRow[dBilltranamountsign])); //basText.formatNum(detailRow[dBilltranamount], "#0.00;(#0.00)", 19);
        //strmWriteCommon.WriteLine("  {0:dd/MM}  {1:dd/MM}  {2,-40} {3,12} {4,3} {5,16} {6,2}",trnsDate,postingDate,basText.trimStr(detailRow[dTrandescription].ToString().Trim() + " " + detailRow[dMerchant].ToString().Trim(),45),basText.formatNum(detailRow[dOrigtranamount],"#,##0.00;(#,##0.00)"),basText.Replace(clsBasValid.validateStr(detailRow[dOrigtrancurrency]),"XXX","   "),basText.formatNum(detailRow[dBilltranamount],"#,##0.00;(#,##0.00)"),CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)
        //strmWriteCommon.WriteLine("  {0:dd/MM}  {1:dd/MM}  {2,-24}  {3,-40} {4,16} {5,16} {6,2}", trnsDate, postingDate, basText.trimStr(detailRow[dRefereneno].ToString().Trim(), 24), basText.trimStr(trnsDesc, 40), strForeignCurr, basText.formatNum(detailRow[dBilltranamount], "#0.00;(#0.00)"), CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + " " + isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)
        strmWriteCommon.WriteLine("  {0:dd/MM}  {1:dd/MM}  {2,-24}  {3,-38} {4,16} {5,16} {6,10} {7,16}", trnsDate, postingDate, basText.trimStr(detailRow[dRefereneno].ToString().Trim(), 24), basText.trimStr(trnsDesc, 38), strForeignCurr, basText.formatNum(detailRow[dBilltranamount], "#0.00;(#0.00)"), CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + " " + isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString())), basText.formatNum(totTrans, "#0.00;(#0.00)"), CrDb(clsBasValid.validateStr(totTrans)));

        if (isEmailStat)
            totNoOfTransEmail++;
        else
            totNoOfTransactions++;
        }

    protected virtual void printCardFooter()
        {
        //completePageDetailRecords();
        strmWriteCommon.WriteLine(String.Empty);
        if (pageNo == totalAccPages)
            strmWriteCommon.WriteLine(basText.replicat(" ", 43) + basText.alignmentLeft("Current Balance", 67) + basText.formatNumUnSign(masterRow[mClosingbalance], "#0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])));
        else
            strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        //		strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(
            basText.alignmentMiddle("Primary Card No", 35)
          + basText.alignmentMiddle("Minimum Payment Due", 20)
          + basText.alignmentMiddle("Due Date", 15)
          + basText.alignmentMiddle("Opening Balance", 15)
          + basText.alignmentMiddle("Payments", 15)
          + basText.alignmentMiddle("Cash&Purchases", 15)
          + basText.alignmentMiddle("Charges", 15)
          + basText.alignmentMiddle("Interest", 15)
          + basText.alignmentMiddle("Closing Balance", 19)
          + basText.alignmentMiddle("International Spend Limit", 19));
        //DBN-5419 => EDT-965
        //strmWriteCommon.WriteLine(basText.alignmentMiddle(basText.formatCardNumber(curMainCard), 35) + basText.alignmentMiddle(masterRow[mMindueamount], 20) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M")
        strmWriteCommon.WriteLine(basText.alignmentMiddle(basText.formatCardNumber(curMainCard), 35) + basText.alignmentMiddle(masterRow[mCarddafamount], 20) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M")
     + basText.alignmentMiddle(basText.formatNumUnSign(masterRow[mOpeningbalance], "##0.00", 12) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance].ToString())), 15)
     + basText.alignmentMiddle(basText.formatNum(masterRow[mTotalpayments], "#,##0.00", 12, "M") + " " + DbCr(Convert.ToDecimal(masterRow[mTotalpayments])), 15)
     + basText.alignmentMiddle(basText.formatNum(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal]), "#,##0.00", 12) + CrDbMinus(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal])), 15)
     + basText.alignmentMiddle(basText.formatNum(masterRow[mTotalcharges].ToString(), "#,##0.00", 12) + " " + CrDbMinus(Convert.ToDecimal(masterRow[mTotalcharges])), 15)
     + basText.alignmentMiddle(basText.formatNum(masterRow[mTotalinterest].ToString(), "#,##0.00", 12) + " " + CrDbMinus(Convert.ToDecimal(masterRow[mTotalinterest])), 15)
     + basText.alignmentMiddle(basText.formatNumUnSign(masterRow[mClosingbalance], "#0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])), 15)
     + basText.alignmentMiddle(basText.formatNumUnSign(masterRow[mIntSpentLimit], "#0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mIntSpentLimit])), 15));
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        //strmWriteCommon.WriteLine(basText.alignmentMiddle("RewardOpenBalance", 20) + basText.alignmentMiddle("RewardEarnedBonus", 20) + basText.alignmentMiddle("RewardRedeemedBonus", 20) + basText.alignmentMiddle("RewardClosingBalance", 20));
        if (isRewardVal)
            {
            strmWriteCommon.WriteLine(basText.alignmentMiddle("Gem Rewards Opening Balance", 30) + basText.alignmentMiddle("Gem Rewards Earned*", 30) + basText.alignmentMiddle("Gem Rewards Redeemed", 30) + basText.alignmentMiddle("Gem Rewards Closing Balance", 30));// + basText.alignmentMiddle("Expiring Gem Points due " + basText.formatDate(rewardRow["EXPIREDBONUSDATE"], "dd/MM/yyyy", 30, "M"), 30));
            //strmWriteCommon.WriteLine(basText.alignmentMiddle("Gem Rewards Opening Balance", 30) + basText.alignmentMiddle("Gem Rewards Earned*", 30) + basText.alignmentMiddle("Gem Rewards Redeemed", 30) + basText.alignmentMiddle("Gem Rewards Closing Balance", 30) + basText.alignmentMiddle("Gem Rewards Expired Bonus Date", 30));
            rewardRows = DSreward.Tables["Reward"].Select("CLIENTID = '" + clsBasValid.validateStr(masterRow[mClientid]).ToString().Trim() + "'");
            if (rewardRows.Length > 0)
                {
                rewardRow = rewardRows[0];
                strmWriteCommon.WriteLine(basText.formatNum(rewardRow[mOpeningbalance], "#0.00;(#0.00)", 30, "M") + basText.formatNum(rewardRow["EARNEDBONUS"], "#0.00;(#0.00)", 30, "M") + basText.formatNum(rewardRow["REDEEMEDBONUS"], "#0.00;(#0.00)", 30, "M") + basText.formatNum(rewardRow[mClosingbalance], "#0.00;(#0.00)", 30, "M"));// + basText.formatNum(rewardRow["EXPIREDBONUSNEXTMONTH"], "#0.00;(#0.00)", 30, "M"));
                //strmWriteCommon.WriteLine(basText.formatNum(rewardRow[mOpeningbalance], "#0.00;(#0.00)", 30, "M") + basText.formatNum(rewardRow["EARNEDBONUS"], "#0.00;(#0.00)", 30, "M") + basText.formatNum(rewardRow["REDEEMEDBONUS"], "#0.00;(#0.00)", 30, "M") + basText.formatNum(rewardRow[mClosingbalance], "#0.00;(#0.00)", 30, "M")  + basText.formatDate(rewardRow["EXPIREDBONUSDATE"], "dd/MM/yyyy", 30, "M"));
                }
            else
                {
                strmWriteCommon.WriteLine(basText.replicats(" ", 80));
                }

            strmWriteCommon.WriteLine(String.Empty);
            strmWriteCommon.WriteLine("* Gem Rewards are only earned on POS and Internet purchases, and not on cash withdrawals. Refunded transactions are not eligible for Gem Rewards.");
            strmWriteCommon.WriteLine(String.Empty);
            }
        else
            {
            strmWriteCommon.WriteLine(basText.replicats(" ", 80));
            strmWriteCommon.WriteLine(basText.replicats(" ", 80));
            strmWriteCommon.WriteLine(basText.replicats(" ", 80));
            strmWriteCommon.WriteLine(basText.replicats(" ", 80));
            strmWriteCommon.WriteLine(basText.replicats(" ", 80));
            }
        }


    protected void printAccountFooter()
        {
        strmWriteCommon.WriteLine("GRAND" + basText.formatNumUnSign(masterRow[mOpeningbalance], "##0.00", 16) + basText.formatNum(masterRow[mTotaldebits], "#0.00;(#0.00)", 16) + basText.formatNum(masterRow[mTotalcredits], "#0.00;(#0.00)", 16) + basText.formatNum(masterRow[mClosingbalance], "#0.00;(#0.00)", 16) + basText.formatNum(masterRow[mMindueamount], "#0.00;(#0.00)", 16) + basText.replicat(" ", 12) + basText.formatDate(masterRow[mStetementduedate], "yyyy/MM/dd") + "\n");//+strEndOfPage
        }


    protected void calcCardlRows()
        {
        totalCardPages = 0;
        totCardRows = 0;
        foreach (DataRow dtRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
            {
            if ((dtRow[dPostingdate] == DBNull.Value) && (dtRow[dDocno] == DBNull.Value)) continue;
            //strmWriteCommon.WriteLine(basText.trimStr(dtRow[dTrandescription],40)); 
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


    protected void calcAccountRows()
        {
        curAccRows = 0;
        accountRows = null;
        stmNo = masterRow[mStatementno].ToString();
        stmNo = masterRow[accountNoName].ToString();

        accountRows = cardsRows;// DSstatement.Tables["tStatementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[accountNoName]).ToString().Trim() + "'");
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

        totSuppCrdsInAcc = 0;
        totCrdNoInAcc = curCrdNoInAcc = 1;//0;
        curMainCard = CurCardNo = masterRow[mCardno].ToString().Trim(); //"";
        if (masterRow[mCardprimary].ToString().ToUpper() == "Y")
            {
            curMainCard = string.Empty;
            totCrdNoInAcc = curCrdNoInAcc = 0;
            mainRows = DSstatement.Tables["tStatementMasterTable"].Select(accountNoName + " = '" + clsBasValid.validateStr(masterRow[accountNoName]) + "'");//ACCOUNTNO
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
                    curMainCard = CurCardNo; //mainRow[mCardno].ToString();
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
                totAccRows += totSuppCrdsInAcc * 2;// *2 for header card num and footer total card value for 
                totalAccPages = totAccRows / MaxDetailInPage;
                totalAccPages = (totAccRows % MaxDetailInPage) > 0 ? ++totalAccPages : totalAccPages;
                }
            }
        }


    protected void completePageDetailRecords()
        {
        //int curPageLine =CurPageRec4Dtl;
        if (pageNo == totalAccPages)
            for (int curPageLine = CurPageRec4Dtl; curPageLine < MaxDetailInPage; curPageLine++)
                strmWriteCommon.WriteLine(String.Empty);
        }


    protected void printStatementSummary()
        {
        streamSummary.WriteLine(strBankName + " Statement");
        streamSummary.WriteLine("");
        streamSummary.WriteLine("");
        streamSummary.WriteLine("Statements not Sent by Email");
        streamSummary.WriteLine("__________________________");
        streamSummary.WriteLine("");
        streamSummary.WriteLine("No of Statements   " + totNoOfCardStat.ToString());
        streamSummary.WriteLine("No of Pages        " + totNoOfPageStat.ToString());
        streamSummary.WriteLine("No of Transactions " + totNoOfTransactions.ToString());
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

        clsStatementSummary StatSummary = new clsStatementSummary();
        StatSummary.BankCode = curBranchVal;
        StatSummary.BankName = strBankName;
        StatSummary.StatementDate = vCurDate;
        StatSummary.CreationDate = DateTime.Now;
        DataRow[] productRows;
        prevAccountNo = string.Empty;
        totNoOfCardStat = 0;
        totNoOfTransactionsInt = 0;
        totNoOfCardEmailStat = 0;
        totNoOfTransEmailInt = 0;
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
                if (masterRow[mCardprimary].ToString() == "Y") //select all transactions for the account
                    //cardsRows = DSstatement.Tables["tStatementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[accountNoName]).ToString().Trim() + "'", "cardprimary desc, cardno, postingdate");
                    cardsRows = DSstatement.Tables["tStatementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[accountNoName]).ToString().Trim() + "'", "cardno, postingdate");
                else
                    cardsRows = mRow.GetChildRows(StatementNoDRel); //, DataRowVersion.Proposed

                strCardNo = masterRow[mCardno].ToString().Trim();
                strPrimaryCardNo = strCardNo;
                strPrimaryCardNo = masterRow[mCardno].ToString().Trim();
                if (masterRow[mCardprimary].ToString() == "Y")
                    {
                    if (haveEmail(masterRow[mClientid].ToString()))
                        {
                        isEmailStat = true;
                        }
                    else
                        {
                        isEmailStat = false;
                        }
                    }
                else
                    {
                    isEmailStat = false;
                    }
                curMainCard = string.Empty;
                CurPageRec4Dtl = 0; totNetUsage = 0; CurrentPageFlag = "F 1"; //if page is based on account no
                calcAccountRows();
                if (!isDebit)
                    if ((
                      ((cardsRows.Length < 1 && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0 && masterRow[mCardprimary].ToString().ToUpper() == "Y")
                      || (totAccRows < 1 && masterRow[mCardprimary].ToString().ToUpper() == "N"))) || !isValidateCard(masterRow[mCardstate].ToString()) && (masterRow[mCardprimary].ToString().ToUpper() == "Y" && masterRow[mCardno].ToString().ToUpper() != curMainCard)) //             || (masterRow[mCardno].ToString() == curMainCard   // Convert.ToDecimal(
                        {
                        continue;
                        }

                prevAccountNo = masterRow[accountNoName].ToString();
                if (isEmailStat)
                    totNoOfCardEmailStat++;
                else
                    totNoOfCardStat++;
                foreach (DataRow dRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
                    {
                    detailRow = dRow;
                    stmNo = detailRow[dStatementno].ToString();
                    if ((detailRow[dPostingdate] == DBNull.Value) && (detailRow[dDocno] == DBNull.Value)) continue;// Exclude On-Hold Transactions 
                    //if (detailRow[mCardprimary].ToString().ToUpper() != "Y" && curSuppCard != detailRow[dCardno].ToString())
                    if (masterRow[mCardprimary].ToString().ToUpper() != "Y" && curSuppCard != masterRow[dCardno].ToString())
                        {
                        tmpDtlRows = DSstatement.Tables["tStatementMasterTable"].Select(mCardno + " = '" + detailRow[dCardno] + "'");
                        if (tmpDtlRows.Length > 0)
                            {
                            if (tmpDtlRows[0][mCardprimary].ToString() == "N" && isValidateCard(tmpDtlRows[0][mCardstate].ToString()))
                                {
                                totSuppCrdVal = 0;
                                if (isActiveSuppCrd)
                                    {
                                    curCrdNoInAcc++;
                                    curAccRows++;
                                    CurPageRec4Dtl++;
                                    }
                                curCrdNoInAcc++;
                                curAccRows++;
                                CurPageRec4Dtl++;
                                isActiveSuppCrd = true;
                                }
                            }
                        }
                    curAccRows++;
                    CurPageRec4Dtl++;
                    hasInterset = CalcTransInterest(detailRow[dTrandescription].ToString());
                    if (hasInterset)
                        {
                        if (isEmailStat)
                            totNoOfTransEmailInt++;
                        else
                            totNoOfTransactionsInt++;
                        }

                    } //end of detail foreach
                if (isActiveSuppCrd)
                    {
                    CurPageRec4Dtl++;
                    curCrdNoInAcc++;
                    curAccRows++;
                    }
                totSuppCrdVal = 0;
                isActiveSuppCrd = false;
                curCrdNoInAcc++;
                if (curAccRows >= totAccRows)
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
                }
            }
        StatSummary = null;
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

    protected void printFileMD5()
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

    public string VIPCondition
        {
        get { return vipCond; }
        set { vipCond = value; }
        }// VIPCondition

    public frmStatementFile setFrm
        {
        set { frmMain = value; }
        }// setFrm

    public bool CreateCorporate
        {
        get { return createCorporateVal; }
        set { createCorporateVal = value; }
        }// CreateCorporate

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

    public string ProductCondition
        {
        get { return productCond; }
        set { productCond = value; }
        }// ProductCondition

    ~clsStatTxtLblSupTot()
        {
        DSstatement.Dispose();
        }
    }
