using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Xml;
using System.Collections;

// Branch 1
public class clsStatementNSGBcredit : clsBasStatement
    {
    private string strBankName;
    private string strFileName;
    private FileStream fileStrm, fileStrmByEmail, fileStrmByBalance, fileSummary, fileRawData, fileStrmByHoldFlag;
    private StreamWriter streamWrit, strmWriteByEmail, strmWriteByBalance, strmWriteCommon, streamSummary, strmRawData, strmWriteByHoldFlag;
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
    private DataRow[] cardsRows, accountRows;
    private DataRow[] mainRows;
    private string CurrentPageFlag;
    private string strCardNo, strPrimaryCardNo;
    private string strForeignCurr;
    private string stmNo;
    private int totNoOfCardStat, totNoOfPageStat, totNoOfCardEmailStat, totNoOfPageEmailStat, totNoOfEmailStat;

    private int curCardRow = 0, curTotCardRows = 0, numOfErr = 0, totNoOfTransactions = 0, totNoOfTransEmail = 0, totNoOfTransactionsInt = 0;
    private bool isPrimaryOnly, isHaveF3 = true, isPrimaryOpened = false;
    private FileStream fileStrmErr;
    private StreamWriter strmWriteErr;
    private string curMainCard, CurCardAddres1, CurCardAddres2, CurCardAddres3;

    private string extAccNum;
    private int totCrdNoInAcc, curCrdNoInAcc;
    private string strOutputPath, strOutputFile, fileSummaryName;
    private frmStatementFile frmMain;
    private int totRec = 1;
    private DataRow[] emailRows = null;
    private string emailTo = string.Empty;
    private bool isEmailStat = false;
    private clsValidateEmail valdEmail = new clsValidateEmail();
    private DateTime vCurDate;
    private string strFileNam, stmntType;
    private ArrayList aryLstFiles = new ArrayList();
    private string commnFileName;
    private DataRow ProductRow;
    protected bool hasInterset = false;

    private int installdocno, intinstalldocno, installentryno, intinstallentryno;
    private decimal installamount, intinstallamount;

    private int accinstalldocno, accintinstalldocno, accinstallentryno, accintinstallentryno;
    private decimal accinstallamount, accintinstallamount;

    string cProduct = string.Empty, curFileName = string.Empty;

    private string rewardCond = "'New Reward Contract'";
    private string installmentCond = string.Empty;
    private bool isReward = false;
    private bool isInstallment = false;
    private DataRow[] rewardRows;
    private DataRow rewardRow;
    private DataRow[] installmentRows;
    protected DataRow installmentRow;
    private int totNoOfCardHoldStat, totNoOfPageHoldStat, totNoOfTransactionsHold; // NSGB-3223

    private string strExcludeCond = string.Empty;


    public clsStatementNSGBcredit()
        {
        }

    

    public string Statement(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pCurDate, string pStmntType, bool pAppendData, string add2FileName)
        {
        string rtrnStr = "Successfully Generate " + pBankName;
        int curMonth = pCurDate.Month;
        bool preExit = true;
        strFileNam = pStrFileName;
        stmntType = pStmntType;

        try
            {

            pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);
            strOutputPath = pStrFileName;
            vCurDate = pCurDate; //DateTime.Now.AddMonths(-1);
            clsBasFile.createDirectory(pStrFileName + vCurDate.ToString("yyyyMM") + pBankName);
            pStrFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + add2FileName + "_" + vCurDate.ToString("yyyyMM") + ".txt";
            strFileName = pStrFileName;
            strBankName = pBankName;

            strOutputFile = pStrFileName;

            // open output file
            fileStrm = new FileStream(pStrFileName, FileMode.Create); //Create
            streamWrit = new StreamWriter(fileStrm, Encoding.Default);
            streamWrit.AutoFlush = true;

            // By Email file
            fileStrmByEmail = new FileStream(clsBasFile.getPathWithoutExtn(pStrFileName) + "_ByEmail." + clsBasFile.getFileExtn(pStrFileName), FileMode.Create);
            strmWriteByEmail = new StreamWriter(fileStrmByEmail, Encoding.Default);
            strmWriteByEmail.AutoFlush = true;

            // By Hold flag file //NSGB-3223
            fileStrmByHoldFlag = new FileStream(clsBasFile.getPathWithoutExtn(pStrFileName) + "_ByHoldFlag." + clsBasFile.getFileExtn(pStrFileName), FileMode.Create);
            strmWriteByHoldFlag = new StreamWriter(fileStrmByHoldFlag, Encoding.Default);
            strmWriteByHoldFlag.AutoFlush = true;

            // By Balance file
            fileStrmByBalance = new FileStream(clsBasFile.getPathWithoutExtn(pStrFileName) + "_BalanceLessThan5." + clsBasFile.getFileExtn(pStrFileName), FileMode.Create);
            strmWriteByBalance = new StreamWriter(fileStrmByBalance, Encoding.Default);
            strmWriteByBalance.AutoFlush = true;

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
            streamSummary.AutoFlush = true;

            // raw data file
            fileRawData = new FileStream(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Card Number.txt", FileMode.Create);
            strmRawData = new StreamWriter(fileRawData, Encoding.Default);
            strmRawData.AutoFlush = true;

            // set branch for data
            curBranchVal = pBankCode; // 1; //3 = real   1 = test
            isFullFields = false;
            //      clsBasStatement.mainTableCond =  " substr(cardno,1,6) !='421192' "; //" cardproduct != 'Visa Business' ";
            //      clsBasStatement.supTableCond =  " substr(cardno,1,6) !='421192' "; 
            // data retrieve
            //FillStatementDataSet(pBankCode); //DSstatement =  //1); //3
            if (isReward)
                {
                //maintainData.fixReward(pBankCode, rewardCond);
                curRewardCond = rewardCond;
                //strMainTableCond = "m.contracttype not in " + rewardCond;
                //strSubTableCond = "d.trandescription not in ('Calcualated _Points','Exchange bonuses for prize','CASHBACK Redemption')";//
                getReward(pBankCode);
                }
            curInstallmentCond = installmentCond;
            //FillStatementDataSet(pBankCode, "vip"); //DSstatement =  //1); //3
            // CBE request
            FillStatementDataSet_WithOverDueDays(pBankCode, "vip"); //DSstatement =  //1); //3
            
            getCardProduct(pBankCode);
            //if (isInstallment==true)
            //    getInstallment(pBankCode);
            //>getClientEmail(pBankCode);
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
                cardsRows = mRow.GetChildRows(StatementNoDRel); //, DataRowVersion.Proposed

                strCardNo = masterRow[mCardno].ToString().Trim();
                if (strCardNo.Length != 16)
                {
                    strmWriteErr.WriteLine("CardNo Length not equal 16 for AccountNo " + masterRow[mAccountno].ToString().Trim() + " Card No =" + strCardNo);
                    //numOfErr++;
                    continue;// Exclude Zero Length Cards 
                }

                if ((clsCheckCard.CalcLuhnCheckDigit(strCardNo.Substring(0, 15)) != strCardNo.Substring(15, 1))) //!clsCheckCard.isValidCard(strCardNo) ||
                {
                    strmWriteErr.WriteLine("Invalid Card Number" + strCardNo);
                    numOfErr++;
                }
              
                //NSGB-3223 stop the old code
                //if (masterRow[mHOLSTMT].ToString() == "Y")
                //    continue;

                strPrimaryCardNo = strCardNo;
                if (masterRow[mCardprimary].ToString() == "N")
                    {
                    strPrimaryCardNo = masterRow[mPrinarycardno].ToString();
                    //calcCardlRows();
                    }

                //start new account
                if (prevAccountNo != masterRow[mAccountno].ToString())
                    {
                    streamWrit.Flush();
                    strmWriteByEmail.Flush();
                    strmWriteByBalance.Flush();
                    strmWriteByHoldFlag.Flush();
                    //NSGB-7434
                    if (masterRow[mHOLSTMT].ToString() == "Y")
                        strmWriteCommon = strmWriteByHoldFlag;
                    else
                        if (haveEmail(masterRow[mClientid].ToString()))
                            {
                            ////NSGB-3223
                            //if (masterRow[mHOLSTMT].ToString() == "Y")
                            //    strmWriteCommon = strmWriteByHoldFlag;
                            //else
                            strmWriteCommon = strmWriteByEmail;
                            isEmailStat = true;
                            }
                        else if (Convert.ToDecimal(masterRow[mClosingbalance].ToString()) >= Convert.ToDecimal(-5.0))
                            {
                            strmWriteCommon = strmWriteByBalance;
                            isEmailStat = false;
                            }
                        else
                            {
                            strmWriteCommon = streamWrit;
                            isEmailStat = false;
                            }

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

                    if (totAccRows < 1
                      && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0) //             || (masterRow[mCardno].ToString() == curMainCard   // Convert.ToDecimal(
                        {
                        isHaveF3 = true;
                        //pageNo=1; totalAccPages =1;
                        continue;
                        }

                    prevAccountNo = masterRow[mAccountno].ToString();
                    //pageNo=1; //if page is based on account no
                    printHeader();//if page is based on account no

                    //NSGB-7434
                    if (masterRow[mHOLSTMT].ToString() == "Y")
                        totNoOfCardHoldStat++;
                    if (isEmailStat)
                        {
                        ////NSGB-3223
                        //if (masterRow[mHOLSTMT].ToString() == "Y")
                        //    totNoOfCardHoldStat++;
                        //else
                        totNoOfCardEmailStat++;
                        }
                    else
                        totNoOfCardStat++;
                    //strmRawData.WriteLine(basText.formatCardNumber(curMainCard) + "|" + CurCardAddres1 + " " + CurCardAddres2 + " " + CurCardAddres3);  //masterRow[mCardno].ToString()
                    strmRawData.WriteLine(curMainCard + "|" + CurCardAddres1 + " " + CurCardAddres2 + " " + CurCardAddres3);  //masterRow[mCardno].ToString()
                    if (Convert.ToDecimal(masterRow[mAccountlim].ToString()) <= Convert.ToDecimal(0))
                        {
                        strmWriteErr.WriteLine("Account Limit Less than or Equal Zero for Account " + masterRow[mAccountno].ToString()); //+ " and Card Number " + strCardNo
                        numOfErr++;
                        }
                    } // End of if(prevAccountNo != masterRow[mAccountno].ToString())
                //calcCardlRows();
                //if(totCardRows < 1)continue ;  //if pages is based on card
                //printHeader();//if pages is based on card

                //strmRawData.WriteLine(masterRow[mCardno].ToString());

                foreach (DataRow dRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
                    {
                    detailRow = dRow;
                    stmNo = detailRow[dStatementno].ToString();
                    if ((detailRow[dPostingdate] == DBNull.Value) && (detailRow[dDocno] == DBNull.Value)) continue;// Exclude On-Hold Transactions 
                    curAccRows++;
                    if (CurPageRec4Dtl >= MaxDetailInPage)
                        {
                        CurPageRec4Dtl = 0;
                        printCardFooter(pBankCode);
                        pageNo++;
                        printHeader();
                        }
                    //						totNetUsage +=clsBasValid.validateNum(detailRow[dBilltranamount]);
                    totNetUsage = calculateCrDb(totNetUsage, clsBasValid.validateNum(detailRow[dBilltranamount]), clsBasValid.validateStr(detailRow[dBilltranamountsign]));
                    CurPageRec4Dtl++;
                    if (detailRow[dTrandescription].ToString() == "Charge interest for Installment")
                    {
                        long intinstalldocno = long.Parse(detailRow[dDocno].ToString());
                        intinstallentryno = int.Parse(detailRow[dEntryNo].ToString());
                        intinstallamount = decimal.Parse(detailRow[dBilltranamount].ToString());
                    }
                    if (detailRow[dTrandescription].ToString() == "Installment repayment")
                        {
                        long installdocno = long.Parse(detailRow[dDocno].ToString());
                        installentryno = int.Parse(detailRow[dEntryNo].ToString());
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
                if ((curAccRows >= totAccRows && totAccRows != 0) || (totAccRows == 0 && curCrdNoInAcc == totCrdNoInAcc))
                    {
                    completePageDetailRecords();
                    printCardFooter(pBankCode);//if pages is based on account
                    //printAccountFooter();
                    CurPageRec4Dtl = 0; //>pageNo=1; if pages is based on account
                    curAccRows = 0;
                    }
                //streamWrit.WriteLine(strEndOfPage);
                //pageNo=1; CurPageRec4Dtl=0; //if pages is based on card
                //completePageDetailRecords();

                } //end of Master foreach

            //fillStatementHistory(pStmntType,pAppendData);
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
            if (preExit)
                {
                // Close output File
                streamWrit.Flush();
                streamWrit.Close();
                fileStrm.Close();

                printStatementSummary(pStrFileName);

                // Close output File
                strmWriteByEmail.Flush();
                strmWriteByEmail.Close();
                fileStrmByEmail.Close();

                // Close output File
                strmWriteByBalance.Flush();
                strmWriteByBalance.Close();
                fileStrmByBalance.Close();

                // Close Hold File
                strmWriteByHoldFlag.Flush();
                strmWriteByHoldFlag.Close();
                fileStrmByHoldFlag.Close();


                // Close Summary File
                streamSummary.Flush();
                streamSummary.Close();
                fileSummary.Close();

                strmWriteErr.Flush();
                strmWriteErr.Close();
                fileStrmErr.Close();

                strmRawData.Flush();
                strmRawData.Close();
                fileRawData.Close();

                aryLstFiles.Add(@strOutputFile);
                if (numOfErr == 0)
                    clsBasFile.deleteFile(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName));
                else
                    aryLstFiles.Add(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err.txt");

                /*      else
                        clsBasFile.deleteFile(pStrFileName);*/

                numOfErr = validateNoOfLines(aryLstFiles, 48);
                if (numOfErr > 0)
                    rtrnStr += string.Format(" with {0} Errors", numOfErr);
                aryLstFiles.Add(@fileSummaryName);
                aryLstFiles.Add(@clsBasFile.getPathWithoutExtn(pStrFileName) + "_ByEmail.txt");
                aryLstFiles.Add(@clsBasFile.getPathWithoutExtn(pStrFileName) + "_BalanceLessThan5.txt");
                aryLstFiles.Add(@clsBasFile.getPathWithoutExtn(pStrFileName) + "_Card Number.txt");
                aryLstFiles.Add(@clsBasFile.getPathWithoutExtn(pStrFileName) + "_ByHoldFlag.txt");
                //>clsBasFile.generateFileMD5(aryLstFiles, @clsBasFile.getPathWithoutExtn(pStrFileName) + ".MD5");
                //>aryLstFiles.Add(@clsBasFile.getPathWithoutExtn(commnFileName) + ".MD5");
                //>SharpZip zip = new SharpZip();
                //>zip.createZip(aryLstFiles, @clsBasFile.getPathWithoutExtn(commnFileName) + ".zip", "");

                //>DSstatement.Dispose();
                }
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

        //if(masterRow[mCardprimary].ToString() == "Y")
        if (pageNo == 1 && totalAccPages == 1) // statement contain 1 page
            {
            CurrentPageFlag = "F 0";
            isHaveF3 = true;
            //totNoOfCardStat++;
            }
        else if (pageNo == 1 && totalAccPages > 1)  //first page of multiple page statement
            {
            CurrentPageFlag = "F 1"; // //middle page of multiple page statement
            //totNoOfCardStat++;
            }
        else if (pageNo < totalAccPages)
            CurrentPageFlag = "F 2";
        else if (pageNo == totalAccPages) //last page of multiple page statement
            {
            CurrentPageFlag = "F 3";
            isHaveF3 = true;
            }

        strmWriteCommon.WriteLine(strEndOfPage);
        strmWriteCommon.WriteLine(basText.replicat(" ", 81) + CurrentPageFlag);  //+"{0,5:MM/yy}",masterRow[mStatementdatefrom]
        strmWriteCommon.WriteLine(basText.replicat(" ", 81) + masterRow[mCardproduct]);  //+"{0,5:MM/yy}",masterRow[mStatementdatefrom]
        //			strmWriteCommon.WriteLine(basText.alignmentLeft(masterRow[mCustomername],50)+ basText.replicat(" ",31) + masterRow[mCardbranchpartname]);  //
        strmWriteCommon.WriteLine(basText.alignmentRight(masterRow[mCustomername], 50));  //
        strmWriteCommon.WriteLine(basText.replicat(" ", 81) + masterRow[mCardbranchpartname]);  //
        //			strmWriteCommon.WriteLine(basText.alignmentMiddle(masterRow[mCustomeraddress1],50)+ basText.replicat(" ",31) + clsBasValid.validateStr(masterRow[mAccountno])); //basText.formatNum(masterRow[mCardno],"####-####-####-####")
        //strmWriteCommon.WriteLine(basText.alignmentRight(ValidateArbic(masterRow[mCustomeraddress1].ToString()), 50)); //basText.formatNum(masterRow[mCardno],"####-####-####-####")
        strmWriteCommon.WriteLine(basText.alignmentRight(ValidateArbic(newaddress1), 50)); //basText.formatNum(masterRow[mCardno],"####-####-####-####")
        strmWriteCommon.WriteLine(basText.replicat(" ", 81) + extAccNum); //clsBasValid.validateStr(masterRow[mExternalno])  // accountno  //basText.formatNum(masterRow[mCardno],"####-####-####-####")
        //			strmWriteCommon.WriteLine(basText.alignmentMiddle(masterRow[mCustomeraddress2],50)+ basText.replicat(" ",31)+"{0,10:dd/MM/yyyy}",masterRow[mStatementdatefrom]);//+ basText.formatNum(masterRow[mAccountlim],"########")  basText.formatNum(masterRow[mCardlimit],"########")
        //strmWriteCommon.WriteLine(basText.alignmentRight(ValidateArbic(masterRow[mCustomeraddress2].ToString()), 50));
        strmWriteCommon.WriteLine(basText.alignmentRight(ValidateArbic(newaddress2), 50));
        strmWriteCommon.WriteLine(basText.replicat(" ", 81) + "{0,10:dd/MM/yyyy}", masterRow[mStatementdateto]);//Statementdatefrom+ basText.formatNum(masterRow[mAccountlim],"########")  basText.formatNum(masterRow[mCardlimit],"########")
        //			strmWriteCommon.WriteLine(basText.alignmentMiddle(masterRow[mCustomeraddress3].ToString().Trim() + " " + masterRow[mCustomerregion].ToString().Trim() + " " + masterRow[mCustomercity].ToString().Trim(),50)+ basText.replicat(" ",31)+pageNo + " / " + totalAccPages);  //basText.formatNum(masterRow[mAccountno],"##-##-######-####")
        strmWriteCommon.WriteLine(basText.alignmentRight(ValidateArbic(masterRow[mCustomeraddress3].ToString().Trim()) + " " + masterRow[mCustomerregion].ToString().Trim() + " " + masterRow[mCustomercity].ToString().Trim(), 50));  //basText.formatNum(masterRow[mAccountno],"##-##-######-####")
        strmWriteCommon.WriteLine(basText.replicat(" ", 81) + pageNo + " / " + totalAccPages);  //basText.formatNum(masterRow[mAccountno],"##-##-######-####")
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        //strmWriteCommon.WriteLine(basText.alignmentMiddle(basText.formatCardNumber(curMainCard), 20) + basText.alignmentMiddle(masterRow[mAccountlim], 13) + basText.formatNum(masterRow[mAccountavailablelim], "##0", 20) + basText.alignmentMiddle(masterRow[mMindueamount], 13) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M") + basText.formatNum(masterRow[mTotaloverdueamount], "##0.00", 13)); //+ basText.formatNum(masterRow[mAccountlim],"#,##0.00",9)+" "+basText.formatNum(masterRow[mTotalcredits],"#,##0.00",9)+" "+basText.formatNum(masterRow[mAccountavailablelim],"#,##0.00",9)+" "+basText.formatNum(masterRow[mMindueamount],"#,##0.00",15)+" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",15)+" "+basText.formatDate(masterRow[mStetementduedate],"dd/MM/yy") //TOTALCREDITS// basText.alignmentMiddle(masterRow[mAccountavailablelim],20) //"#,##0"
        //strmWriteCommon.WriteLine(basText.alignmentMiddle(curMainCard, 20) + basText.alignmentMiddle(masterRow[mAccountlim], 13) + basText.formatNum(masterRow[mAccountavailablelim], "##0", 20) + basText.alignmentMiddle(masterRow[mMindueamount], 13) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M") + basText.formatNum(masterRow[mTotaloverdueamount], "##0.00", 13)); //+ basText.formatNum(masterRow[mAccountlim],"#,##0.00",9)+" "+basText.formatNum(masterRow[mTotalcredits],"#,##0.00",9)+" "+basText.formatNum(masterRow[mAccountavailablelim],"#,##0.00",9)+" "+basText.formatNum(masterRow[mMindueamount],"#,##0.00",15)+" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",15)+" "+basText.formatDate(masterRow[mStetementduedate],"dd/MM/yy") //TOTALCREDITS// basText.alignmentMiddle(masterRow[mAccountavailablelim],20) //"#,##0"
        //NSGB-2889 EDT-1334
        strmWriteCommon.WriteLine(basText.alignmentMiddle(curMainCard, 20) + basText.alignmentMiddle(masterRow[mAccountlim], 13) + basText.formatNum((decimal.Parse(masterRow[mAccountavailablelim].ToString()) - decimal.Parse(masterRow[mInstallmentUsedLimit].ToString())), "##0", 20) + basText.alignmentMiddle(masterRow[mMindueamount], 13) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M") + basText.formatNum(masterRow[mTotaloverdueamount], "##0.00", 13)); //+ basText.formatNum(masterRow[mAccountlim],"#,##0.00",9)+" "+basText.formatNum(masterRow[mTotalcredits],"#,##0.00",9)+" "+basText.formatNum(masterRow[mAccountavailablelim],"#,##0.00",9)+" "+basText.formatNum(masterRow[mMindueamount],"#,##0.00",15)+" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",15)+" "+basText.formatDate(masterRow[mStetementduedate],"dd/MM/yy") //TOTALCREDITS// basText.alignmentMiddle(masterRow[mAccountavailablelim],20) //"#,##0"
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        //if (isReward)
        //    {
        //    rewardRows = DSreward.Tables["Reward"].Select("CLIENTID = '" + clsBasValid.validateStr(masterRow[mClientid]).ToString().Trim() + "'");
        //    if (rewardRows.Length > 0)
        //        {
        //        rewardRow = rewardRows[0];
        //        strmWriteCommon.WriteLine(basText.formatNum(rewardRow[mOpeningbalance], "#0.00;(#0.00)", 20, "M") + basText.formatNum(int.Parse(rewardRow[mEarnedBonus].ToString()), "#0.00;(#0.00)", 20, "M") + basText.formatNum(rewardRow[mRedeemedBonus], "#0.00;(#0.00)", 20, "M") + basText.formatNum(rewardRow[mExpiredBonus], "#0.00;(#0.00)", 20, "M") + basText.formatNum(rewardRow[mClosingbalance], "#0.00;(#0.00)", 20, "M"));
        //        }
        //    else
        //        {
        //        strmWriteCommon.WriteLine(String.Empty);
        //        }
        //    }
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        if (pageNo == 1)
            //if (isInstallment == true)
            //    {
            //    installmentRows = DSInstallment.Tables["installment"].Select("CLIENTID = '" + clsBasValid.validateStr(masterRow[mClientid]).ToString().Trim() + "' and CREDITCONTRACTS = '" + clsBasValid.validateStr(masterRow[mContractno]).ToString().Trim() + "'");
            //    if (installmentRows.Length > 0)
            //        {
            //        installmentRow = installmentRows[0];
            //        strmWriteCommon.WriteLine(basText.replicat(" ", 17) + basText.alignmentLeft("Previous Balance", 63) + basText.alignmentRight(basText.formatNumUnSign(decimal.Parse(masterRow[mOpeningbalance].ToString()) + decimal.Parse(installmentRow[mOpeningbalance].ToString()), "##0.00", 16), 16) + " " + CrDb(decimal.Parse(masterRow[mOpeningbalance].ToString()) + decimal.Parse(installmentRow[mOpeningbalance].ToString())));
            //        }
            //    }
            //else
            strmWriteCommon.WriteLine(basText.replicat(" ", 17) + basText.alignmentLeft("Previous Balance", 63) + basText.alignmentRight(basText.formatNumUnSign(masterRow[mOpeningbalance], "##0.00", 16), 16) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance])));  //63  + CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign]))
        else
            strmWriteCommon.WriteLine(String.Empty);

        strmWriteCommon.WriteLine(String.Empty);
        totalPages++;
        //NSGB-7434
        if (masterRow[mHOLSTMT].ToString() == "Y")
            totNoOfPageHoldStat++;
        else if (isEmailStat)
            {
            ////NSGB-3223
            //if (masterRow[mHOLSTMT].ToString() == "Y")
            //    totNoOfPageHoldStat++;
            //else
            totNoOfPageEmailStat++;
            }
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
        {
            trnsDesc = detailRow[dTrandescription].ToString().Trim();
        }
        else
        {
            trnsDesc = detailRow[dMerchant].ToString().Trim();
        }

        //NSGB-3444 OLD
        //NSGB-15335 NEw
        //if (trnsDesc == "Charge interest for Installment")// || trnsDesc == "Charge interest for Acceleration") 
        //{
        //    CurPageRec4Dtl--;
        //    curAccRows--;
        //    return;
        //}

        //strmWriteCommon.WriteLine("  {0:dd/MM}  {1:dd/MM}  {2,-40} {3,16} {4,16} {5,2}", trnsDate, postingDate, basText.trimStr(trnsDesc, 45), strForeignCurr, basText.formatNum(detailRow[dBilltranamount], "#0.00;(#0.00)"), CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + " " + isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)

        if (detailRow[dInstallmentData].ToString().Trim() != "")
            {
            if (trnsDesc == "Installment repayment")
                {
                string input = detailRow[dInstallmentData].ToString().Trim();
                int index = detailRow[dInstallmentData].ToString().Trim().IndexOf(":");
                if (index > 0)
                    input = input.Substring(0, index);
                strmWriteCommon.WriteLine("  {0:dd/MM}  {1:dd/MM}  {2,-68} {4,10} {5,2}", trnsDate, postingDate, basText.trimStr(input.Substring(0, input.IndexOf(',')) + input.Substring(basText.GetNthIndex(input, ',', 2)).Substring(input.Substring(basText.GetNthIndex(input, ',', 2)).IndexOf('#')).Replace('#', ' ') + " " + basText.trimStr(GetMerchantByOrigDocNo(long.Parse(detailRow[dOrigDocNo].ToString())), 40), 68), strForeignCurr, basText.formatNum(detailRow[dBilltranamount], "#0.00;(#0.00)"), CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + " " + isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)
                }
            else // NSGB-3525
                strmWriteCommon.WriteLine("  {0:dd/MM}  {1:dd/MM}  {2,-40} {3,16} {4,16} {5,2}", trnsDate, postingDate, basText.trimStr(trnsDesc, 45), strForeignCurr, basText.formatNum(detailRow[dBilltranamount], "#0.00;(#0.00)"), CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + " " + isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)
            }
        else
            strmWriteCommon.WriteLine("  {0:dd/MM}  {1:dd/MM}  {2,-40} {3,16} {4,16} {5,2}", trnsDate, postingDate, basText.trimStr(trnsDesc, 45), strForeignCurr, basText.formatNum(detailRow[dBilltranamount], "#0.00;(#0.00)"), CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + " " + isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)
        //NSGB-7434
        if (masterRow[mHOLSTMT].ToString() == "Y")
            totNoOfTransactionsHold++;
        if (isEmailStat)
            ////NSGB-3223
            //if (masterRow[mHOLSTMT].ToString() == "Y")
            //    totNoOfTransactionsHold++;
            //else
            totNoOfTransEmail++;
        else
            totNoOfTransactions++;
        }

    protected void printCardFooter(int pBankCode)
        {
            //string queryString;
            //string strCONTRACTNO = masterRow[mContractno].ToString();
            //string strACCOUNTNO = masterRow[mAccountno].ToString();
            //string strOverDueDays = "";

            //queryString = " Select  ODDAYS from A4M.ZM_EOD_CONT_ACCT where BRANCH = " + pBankCode + " and CONTRACTNO = '" + strCONTRACTNO + "'  and ACCOUNTNO = '" + strACCOUNTNO + "' and ROWNUM <= 1 order by OPDATE DESC ";
            ////"select distinct t.cardproduct from " + clsSessionValues.mainDbSchema + " t where t.branch = " + pBankCode + " and cardproduct is not null";
            //OracleConnection conn;
            //OracleCommand cmd;
            //OracleDataReader rdr;
            //try
            //{
            //    conn = new OracleConnection(clsDbCon.sConOracle);
            //    cmd = new OracleCommand(queryString, conn);
            //    conn.Open();
            //    rdr = cmd.ExecuteReader();
            //    while (rdr.Read())
            //    {
            //        strOverDueDays = rdr[0].ToString();
            //    }
            //    rdr.Close();
            //}
            //catch
            //{

            //}
            //finally
            //{
            //    conn = null;
            //    cmd = null;
            //    rdr = null;
            //}

        //completePageDetailRecords();
        strmWriteCommon.WriteLine(String.Empty);
        if (pageNo == totalAccPages)
            //if (isInstallment == true)
            //    {
            //    installmentRows = DSInstallment.Tables["installment"].Select("CLIENTID = '" + clsBasValid.validateStr(masterRow[mClientid]).ToString().Trim() + "' and CREDITCONTRACTS = '" + clsBasValid.validateStr(masterRow[mContractno]).ToString().Trim() + "'");
            //    if (installmentRows.Length > 0)
            //        {
            //        installmentRow = installmentRows[0];
            //        strmWriteCommon.WriteLine(basText.replicat(" ", 17) + basText.alignmentLeft("Current Balance", 63) + basText.formatNumUnSign(decimal.Parse(masterRow[mClosingbalance].ToString()) + decimal.Parse(installmentRow[mClosingbalance].ToString()), "##0.00", 16) + " " + CrDb(decimal.Parse(masterRow[mClosingbalance].ToString()) + decimal.Parse(installmentRow[mClosingbalance].ToString())));
            //        }
            //    }
            //else
            strmWriteCommon.WriteLine(basText.replicat(" ", 17) + basText.alignmentLeft("Current Balance", 63) + basText.formatNumUnSign(masterRow[mClosingbalance], "#0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])));
        else
            strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        
        
        // replacing the empty line with messae of over due days jira NSGB-10030
        //strmWriteCommon.WriteLine(String.Empty);
        //strmWriteCommon.WriteLine("íÊÚíä Úáì ÇáÈäß ÇáÇáÊÒÇã ÈÇáÅÝÕÇÍ Úä ÇáÈäæÏ ÇáÊÇáíÉ Ýí ÍÇáå ÅÕÏÇÑ ßÔæÝ ÍÓÇÈÇÊ ÎÇÕÉ ÈÈØÇÞÇÊ ÇáÇÆÊãÇä ÚÏÏ ÃíÇã ÇáÊÃÎíÑ  " + strOverDueDays);
        
        // CBE request (no over due days till 6 months) // --- UPDATE --- rollback to print the over due days
        strmWriteCommon.WriteLine(masterRow["OverDueDays"].ToString());
        //strmWriteCommon.WriteLine(String.Empty);

        
        
        //strmWriteCommon.WriteLine(String.Empty);
        //strmWriteCommon.WriteLine(basText.alignmentMiddle(basText.formatCardNumber(curMainCard), 35) + basText.alignmentMiddle(basText.formatNumUnSign(masterRow[mClosingbalance], "#0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])), 20) + basText.alignmentMiddle(masterRow[mMindueamount], 20) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M"));
        //if (isInstallment == true)
        //    {
        //    installmentRows = DSInstallment.Tables["installment"].Select("CLIENTID = '" + clsBasValid.validateStr(masterRow[mClientid]).ToString().Trim() + "' and CREDITCONTRACTS = '" + clsBasValid.validateStr(masterRow[mContractno]).ToString().Trim() + "'");
        //    if (installmentRows.Length > 0)
        //        {
        //        installmentRow = installmentRows[0];
        //        strmWriteCommon.WriteLine(basText.alignmentMiddle(curMainCard, 35) + basText.alignmentMiddle(basText.formatNumUnSign(decimal.Parse(masterRow[mClosingbalance].ToString()) + decimal.Parse(installmentRow[mClosingbalance].ToString()), "#0.00", 16) + " " + CrDb(decimal.Parse(masterRow[mClosingbalance].ToString()) + decimal.Parse(installmentRow[mClosingbalance].ToString())), 20) + basText.alignmentMiddle(masterRow[mMindueamount], 20) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M"));
        //        }
        //    }
        //else
        if (isReward)
            {
            //NSGB-3444
            //rewardRows = DSreward.Tables["Reward"].Select("CLIENTID = '" + clsBasValid.validateStr(masterRow[mClientid]).ToString().Trim() + "'");
            rewardRows = DSreward.Tables["Reward"].Select("CLIENTID = '" + clsBasValid.validateStr(masterRow[mClientid]).ToString().Trim() + "' and CREDITCONTRACTS = '" + clsBasValid.validateStr(masterRow[mContractno]).ToString().Trim() + "'");
            if (rewardRows.Length > 0)
                {
                rewardRow = rewardRows[0];
                //NSGB-3444 
                //strmWriteCommon.WriteLine(basText.formatNum(rewardRow[mOpeningbalance], "#0.00;(#0.00)", 20, "M") + basText.formatNum(int.Parse(rewardRow[mEarnedBonus].ToString()), "#0.00;(#0.00)", 20, "M") + basText.formatNum(rewardRow[mRedeemedBonus], "#0.00;(#0.00)", 20, "M") + basText.formatNum(rewardRow[mExpiredBonus], "#0.00;(#0.00)", 20, "M") + basText.formatNum(rewardRow[mClosingbalance], "#0.00;(#0.00)", 20, "M"));
                if (masterRow[mCardproduct].ToString() == strExcludeCond)
                    strmWriteCommon.WriteLine(String.Empty);
                else
                    strmWriteCommon.WriteLine(basText.replicat(" ", 9) + basText.formatNum(rewardRow[mOpeningbalance], "#0.00;(#0.00)", 20, "L") + basText.formatNum(int.Parse(rewardRow[mEarnedBonus].ToString()), "#0.00;(#0.00)", 20, "L") + basText.formatNum(rewardRow[mRedeemedBonus], "#0.00;(#0.00)", 20, "L") + basText.formatNum(rewardRow[mExpiredBonus], "#0.00;(#0.00)", 20, "L") + basText.formatNum(rewardRow[mClosingbalance], "#0.00;(#0.00)", 20, "L"));
                }
            //else if (masterRow[mCardproduct].ToString() == strExcludeCond)
            //    {
            //    strmWriteCommon.WriteLine(String.Empty);
            //    }
            else if (masterRow[mCardproduct].ToString() == strExcludeCond)
                strmWriteCommon.WriteLine(String.Empty);
            else
                strmWriteCommon.WriteLine(basText.replicat(" ", 9) + basText.formatNum(0, "#0.00;(#0.00)", 20, "L") + basText.formatNum(0, "#0.00;(#0.00)", 20, "L") + basText.formatNum(0, "#0.00;(#0.00)", 20, "L") + basText.formatNum(0, "#0.00;(#0.00)", 20, "L") + basText.formatNum(0, "#0.00;(#0.00)", 20, "L"));
            //strmWriteCommon.WriteLine(String.Empty);                
            }
        else
            {
            strmWriteCommon.WriteLine(String.Empty);
            }
        strmWriteCommon.WriteLine(basText.alignmentMiddle(curMainCard, 35) + basText.alignmentMiddle(basText.formatNumUnSign(masterRow[mClosingbalance], "#0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])), 20) + basText.alignmentMiddle(masterRow[mMindueamount], 20) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M"));
        //strmWriteCommon.WriteLine(String.Empty);
        //if (isReward)
        //    {
        //    rewardRows = DSreward.Tables["Reward"].Select("CLIENTID = '" + clsBasValid.validateStr(masterRow[mClientid]).ToString().Trim() + "'");
        //    if (rewardRows.Length > 0)
        //        {
        //        rewardRow = rewardRows[0];
        //        strmWriteCommon.WriteLine(basText.formatNum(rewardRow[mOpeningbalance], "#0.00;(#0.00)", 20, "M") + basText.formatNum(int.Parse(rewardRow[mEarnedBonus].ToString()), "#0.00;(#0.00)", 20, "M") + basText.formatNum(rewardRow[mRedeemedBonus], "#0.00;(#0.00)", 20, "M") + basText.formatNum(rewardRow[mExpiredBonus], "#0.00;(#0.00)", 20, "M") + basText.formatNum(rewardRow[mClosingbalance], "#0.00;(#0.00)", 20, "M"));
        //        }
        //    else
        //        {
        //        strmWriteCommon.WriteLine(String.Empty);
        //        }
        //    }
        }

    //protected void printAccountFooter()
    //    {
    //    strmWriteCommon.WriteLine("GRAND" + basText.formatNum(masterRow[mOpeningbalance], "#0.00;(#0.00)", 16) + basText.formatNum(masterRow[mTotaldebits], "#0.00;(#0.00)", 16) + basText.formatNum(masterRow[mTotalcredits], "#0.00;(#0.00)", 16) + basText.formatNum(masterRow[mClosingbalance], "#0.00;(#0.00)", 16) + basText.formatNum(masterRow[mMindueamount], "#0.00;(#0.00)", 16) + basText.replicat(" ", 12) + basText.formatDate(masterRow[mStetementduedate], "yyyy/MM/dd") + "\n");//+strEndOfPage
    //    }


    private void calcCardlRows()
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


    private void calcAccountRows()
        {
        curAccRows = 0;
        accountRows = null;
        stmNo = masterRow[mStatementno].ToString();
        stmNo = masterRow[mAccountno].ToString();

        accountRows = DSstatement.Tables["tStatementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[mAccountno]).ToString().Trim() + "'");
        totalAccPages = 0;
        totAccRows = 0;
        string prevCardNo = String.Empty, CurCardNo = String.Empty;
        int currAccRowsPages = 0;
        foreach (DataRow dtAccRow in accountRows) //mRow.GetChildRows(StatementNoDRel)
            {
            if ((dtAccRow[dPostingdate] == DBNull.Value) && (dtAccRow[dDocno] == DBNull.Value)) continue;
            CurCardNo = dtAccRow[dCardno].ToString();
            if (CurCardNo.Trim().Length < 1) continue;
            //if (dtAccRow[dTrandescription].ToString() == "Charge interest for Installment") continue; //NSGB-4905

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
        mainRows = DSstatement.Tables["tStatementMasterTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[mAccountno]) + "'");
        curMainCard = CurCardNo = "";
        foreach (DataRow mainRow in mainRows) //mRow.GetChildRows(StatementNoDRel)
            {
            totCrdNoInAcc++;
            CurCardNo = mainRow[mCardno].ToString();
            CurCardAddres1 = mainRow[mCustomeraddress1].ToString();
            CurCardAddres2 = mainRow[mCustomeraddress2].ToString();
            CurCardAddres3 = mainRow[mCustomeraddress3].ToString();
            if (mainRow[mCardprimary].ToString() == "Y")
                {
                curMainCard = CurCardNo; //mainRow[mCardno].ToString();
                CurCardAddres1 = mainRow[mCustomeraddress1].ToString();
                CurCardAddres2 = mainRow[mCustomeraddress2].ToString();
                CurCardAddres3 = mainRow[mCustomeraddress3].ToString();
                }
            if (mainRow[mCardprimary].ToString() == "Y" && isValidateCard(mainRow[mCardstate].ToString()))
                {
                curMainCard = CurCardNo; //mainRow[mCardno].ToString();
                CurCardAddres1 = mainRow[mCustomeraddress1].ToString();
                CurCardAddres2 = mainRow[mCustomeraddress2].ToString();
                CurCardAddres3 = mainRow[mCustomeraddress3].ToString();
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
                strmWriteCommon.WriteLine(String.Empty);
        }


    private void printStatementSummary(string pStrFileName)
        {

        //    string strfileStrmByEmail = clsBasFile.getPathWithoutExtn(pStrFileName) + "_ByEmail." + clsBasFile.getFileExtn(pStrFileName);
        //    string strfileStrm = clsBasFile.getPathWithoutExtn(pStrFileName) + "." + clsBasFile.getFileExtn(pStrFileName);
        //    string strfileStrmByHoldFlag = clsBasFile.getPathWithoutExtn(pStrFileName) + "_ByHoldFlag." + clsBasFile.getFileExtn(pStrFileName);

        //    int intCountOfLinesByEmail = 0;
        //    int intCountOfLines = 0;
        //    int intCountOfLinesByHoldFlag = 0;

        //    StreamReader rd = null;
        //    rd = new StreamReader(strfileStrmByEmail);
        //    //int intCountOfLinesByEmail = rd.ReadToEnd().Length;
        //    while (rd.ReadLine() != null)
        //    {
        //        intCountOfLinesByEmail++;
        //    }

        //    rd = null;
        //    rd = new StreamReader(strfileStrm);
        //    while (rd.ReadLine() != null)
        //    {
        //        intCountOfLines++;
        //    }

        //    rd = null;
        //    rd = new StreamReader(strfileStrmByHoldFlag);
        //    while (rd.ReadLine() != null)
        //    {
        //        intCountOfLinesByHoldFlag++;
        //    } 
        
        //rd = null;

        streamSummary.WriteLine(strBankName + " Visa Credit Statement");
        streamSummary.WriteLine("");
        streamSummary.WriteLine("");
        streamSummary.WriteLine("Statements not Sent by Email");
        streamSummary.WriteLine("__________________________");
        streamSummary.WriteLine("");
        streamSummary.WriteLine("No of Statements   " + totNoOfCardStat.ToString());
        streamSummary.WriteLine("No of Pages        " + totNoOfPageStat.ToString());// (intCountOfLines / 48).ToString());
        streamSummary.WriteLine("No of Transactions " + totNoOfTransactions.ToString());
        streamSummary.WriteLine("");
        streamSummary.WriteLine("");
        streamSummary.WriteLine("__________________________");
        streamSummary.WriteLine("");
        streamSummary.WriteLine("Statements Sent by Email");
        streamSummary.WriteLine("");
        streamSummary.WriteLine("No of Statements   " + totNoOfCardEmailStat.ToString());
        streamSummary.WriteLine("No of Pages        " + totNoOfPageEmailStat.ToString());// (intCountOfLinesByEmail / 48).ToString());
        streamSummary.WriteLine("No of Transactions " + totNoOfTransEmail.ToString());
        streamSummary.WriteLine("");
        streamSummary.WriteLine("");
        streamSummary.WriteLine("__________________________");
        streamSummary.WriteLine("");
        streamSummary.WriteLine("Statements with Hold Flag");
        streamSummary.WriteLine("");
        streamSummary.WriteLine("No of Statements   " + totNoOfCardHoldStat.ToString());
        streamSummary.WriteLine("No of Pages        " + totNoOfPageHoldStat.ToString()); // (intCountOfLinesByHoldFlag / 48).ToString()); 
        streamSummary.WriteLine("No of Transactions " + totNoOfTransactionsHold.ToString());
        streamSummary.WriteLine("");
        streamSummary.WriteLine("");
        streamSummary.WriteLine("__________________________");
        streamSummary.WriteLine("");
        streamSummary.WriteLine("Total Statements");
        streamSummary.WriteLine("");
        streamSummary.WriteLine("No of Statements   " + (totNoOfCardStat + totNoOfCardEmailStat + totNoOfCardHoldStat).ToString());
        streamSummary.WriteLine("No of Pages        " + (totNoOfPageStat + totNoOfPageEmailStat + totNoOfPageHoldStat).ToString());
        //streamSummary.WriteLine("No of Pages        " + (intCountOfLines + intCountOfLinesByEmail + intCountOfLinesByHoldFlag).ToString());
        streamSummary.WriteLine("No of Transactions " + (totNoOfTransactions + totNoOfTransEmail + totNoOfTransactionsHold).ToString());

        clsValidatePageSize ValidatePageSize = new clsValidatePageSize();
        ValidatePageSize.ValidatePageSize(strFileName, 48, strEndOfPage);
        streamSummary.WriteLine(ValidatePageSize.outMessage);

        clsStatementSummary StatSummary = new clsStatementSummary();
        StatSummary.BankCode = curBranchVal;
        StatSummary.BankName = strBankName;
        StatSummary.StatementDate = vCurDate;
        StatSummary.CreationDate = DateTime.Now;
        DataRow[] productRows;
        totNoOfCardStat = 0;
        totNoOfCardEmailStat = 0;
        totNoOfTransactionsInt = 0;
        totNoOfTransEmail = 0;
        totNoOfCardHoldStat = 0;
        totNoOfTransactionsHold = 0;
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

                //NSGB-3223 stop the old code
                //if (masterRow[mHOLSTMT].ToString() == "Y")
                //    continue;

                strPrimaryCardNo = strCardNo;
                if (masterRow[mCardprimary].ToString() == "N")
                    {
                    strPrimaryCardNo = masterRow[mPrinarycardno].ToString();
                    }

                //start new account
                if (prevAccountNo != masterRow[mAccountno].ToString())
                    {
                    if (haveEmail(masterRow[mClientid].ToString()))
                        {
                        isEmailStat = true;
                        }
                    else if (Convert.ToDecimal(masterRow[mClosingbalance].ToString()) >= Convert.ToDecimal(-5.0))
                        {
                        isEmailStat = false;
                        }
                    else
                        {
                        isEmailStat = false;
                        }

                    curMainCard = string.Empty;
                    pageNo = 1;//totalAccPages = 1 ; pageNo=1;
                    CurPageRec4Dtl = 0; totNetUsage = 0; CurrentPageFlag = "F 1"; //if page is based on account no
                    calcAccountRows();

                    if (totAccRows < 1 && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0) //             || (masterRow[mCardno].ToString() == curMainCard   // Convert.ToDecimal(
                        continue;

                    prevAccountNo = masterRow[mAccountno].ToString();

                    //NSGB-7434
                    if (masterRow[mHOLSTMT].ToString() == "Y")
                        totNoOfCardHoldStat++;
                    if (isEmailStat)
                        {
                        ////NSGB-3223
                        //if (masterRow[mHOLSTMT].ToString() == "Y")
                        //    totNoOfCardHoldStat++;
                        //else
                        totNoOfCardEmailStat++;
                        }
                    else
                        totNoOfCardStat++;
                    }

                foreach (DataRow dRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
                    {
                    detailRow = dRow;
                    stmNo = detailRow[dStatementno].ToString();
                    if ((detailRow[dPostingdate] == DBNull.Value) && (detailRow[dDocno] == DBNull.Value)) continue;// Exclude On-Hold Transactions 
                    curAccRows++;
                    if (CurPageRec4Dtl >= MaxDetailInPage)
                        {
                        CurPageRec4Dtl = 0;
                        }
                    CurPageRec4Dtl++;
                    hasInterset = CalcTransInterest(detailRow[dTrandescription].ToString());
                    if (hasInterset)
                        {
                        //NSGB-7434
                        if (masterRow[mHOLSTMT].ToString() == "Y")
                            totNoOfTransactionsHold++;
                        if (isEmailStat)
                            ////NSGB-3223
                            //if (masterRow[mHOLSTMT].ToString() == "Y")
                            //    totNoOfTransactionsHold++;
                            //else
                            totNoOfTransEmail++;
                        else
                            totNoOfTransactionsInt++;
                        }
                    } //end of detail foreach
                curCrdNoInAcc++;
                if ((curAccRows >= totAccRows && totAccRows != 0) || (totAccRows == 0 && curCrdNoInAcc == totCrdNoInAcc))
                    {
                    CurPageRec4Dtl = 0; //>pageNo=1; if pages is based on account
                    curAccRows = 0;
                    }
                } //end of Master foreach
            if ((totNoOfCardStat + totNoOfCardEmailStat) != 0)
                {
                StatSummary.NoOfStatements = totNoOfCardStat + totNoOfCardEmailStat + totNoOfCardHoldStat;
                StatSummary.NoOfTransactionsInt = totNoOfTransactionsInt + totNoOfTransEmail;
                StatSummary.ProductCode = int.Parse(ProductRow["code"].ToString());
                StatSummary.ProductName = ProductRow["name"].ToString();
                StatSummary.InsertRecordDb(vCurDate.ToString("yyyyMM") + strBankName);
                totNoOfCardStat = 0;
                totNoOfCardEmailStat = 0;
                totNoOfTransactionsInt = 0;
                totNoOfTransEmail = 0;
                totNoOfCardHoldStat = 0;
                totNoOfTransactionsHold = 0;
                totNoOfPageStat = 0; 
                totNoOfPageEmailStat = 0; 
                totNoOfPageHoldStat = 0;
                }
            }
        StatSummary = null;
        }


    private bool haveEmail(string pClientID)
        {
        bool rtrnVal = false;
        emailTo = string.Empty;
        emailRows = null;
        try
            {
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
            }
        catch
            {
            }
        return rtrnVal;
        }

    public string SplitByProduct(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pCurDate, string pStmntType, bool pAppendData)
        {
        string rtrnStr = "Successfully Generate " + pBankName;
        string queryString;
        queryString = "select distinct t.cardproduct from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " t where t.branch = " + pBankCode + " and cardproduct is not null";
        try
            {
            clsMaintainData maintainData = new clsMaintainData();
            maintainData.notRward = false;
            maintainData.matchCardBranch4Account(pBankCode);
            getClientEmail(pBankCode);
            //commnFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM") + ".txt";
            commnFileName = pStrFileName + pCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + pCurDate.ToString("yyyyMM") + ".txt";
            OracleConnection conn = new OracleConnection(clsDbCon.sConOracle);
            OracleCommand cmd = new OracleCommand(queryString, conn);
            conn.Open();
            OracleDataReader rdr;
            rdr = cmd.ExecuteReader();
            while (rdr.Read())
                {
                masterQuery = detailQuery = string.Empty;
                if (isReward)
                    {
                    mainTableCond = " cardproduct = '" + rdr[0].ToString() + "'";
                    supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and cardproduct = '" + rdr[0].ToString() + "' union all select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and contracttype = 'Reward Program - " + rdr[0].ToString() + "')";
                    }
                else
                    {
                    mainTableCond = " cardproduct = '" + rdr[0].ToString() + "'";
                    supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and " + mainTableCond + ") and d.trandescription not in ('Calcualated _Points','Exchange bonuses for prize','Loyalty Points Redemption','Bonuses expiration')";
                    }
                Statement(pStrFileName, pBankName, pBankCode, pStrFile, pCurDate, pStmntType, pAppendData, rdr[0].ToString());
                DSstatement.Clear();
                StatementNoDRel = null;
                mainTableCond = "";
                supTableCond = "";
                totRec = 1;
                //DSstatement = null;
                }
            rdr.Close();
            clsBasFile.generateFileMD5(aryLstFiles, @clsBasFile.getPathWithoutExtn(commnFileName) + ".MD5");
            aryLstFiles.Add(@clsBasFile.getPathWithoutExtn(commnFileName) + ".MD5");
            SharpZip zip = new SharpZip();
            zip.createZip(aryLstFiles, @clsBasFile.getPathWithoutExtn(commnFileName) + ".zip", "");


            //strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
            //strFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM"); // + ".pdf"
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

    private void add2FileList(string pFileName)
        {
        int myIndex = aryLstFiles.BinarySearch((object)pFileName);
        if (myIndex < 0)
            aryLstFiles.Add(@pFileName);
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

    public string rewardCondition
        {
        get { return rewardCond; }
        set { rewardCond = value; }
        }// rewardCondition

    public string installmentCondition
        {
        get { return installmentCond; }
        set { installmentCond = value; }
        }// installmentCondition

    public bool isRewardVal
        {
        get { return isReward; }
        set { isReward = value; }
        }// isRewardVal

    public bool isInstallmentVal
        {
        get { return isInstallment; }
        set { isInstallment = value; }
        }// isInstallmentVal

    public string ExcludeCond
        {
        get { return strExcludeCond; }
        set { strExcludeCond = value; }
        }  // ExcludeCond

    ~clsStatementNSGBcredit()
        {
        DSstatement.Dispose();
        }

    }
