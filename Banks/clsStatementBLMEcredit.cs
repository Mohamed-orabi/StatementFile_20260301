using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Xml;
using System.Collections;

// Branch 4
public class clsStatementBLMEcredit : clsBasStatement
{
    private string strBankName;
    private string strFileName;
    private FileStream fileStrm, fileSummary, fileRawData;
    private StreamWriter streamWrit, streamSummary, strmRawData;
    //private DataSet DSstatement;
    //private OracleDataReader drPrimaryCards, drMaster,drDetail;
    private DataRow masterRow;
    private DataRow detailRow;
    private string strEndOfLine = "\u000D";  //+ "M" ^M
    //		private string strEndOfPage = "\u000C"  ;  //+ "\u000D"+ "M" ^L + ^M
    private string strEndOfPage = "\u000C";  //  + basText.replicat(" ",85) + "F 2"  + "\u000D"+ "M" ^L + ^M
    private const int MaxDetailInPage = 25; //
    private const int MaxDetailInLastPage = 25; //
    private int CurPageRec4Dtl = 0;
    private int pageNo = 0, totalPages = 0, totalCardPages = 0
        , totalAccPages = 0, totCardRows = 0, totAccRows = 0, curAccRows = 0;//
    private string lastPageTotal;
    private string curCardNo;//,PrevCardNo
    private string curAccountNo, prevAccountNo = String.Empty;//,PrevCardNo
    private decimal totNetUsage = 0, totTrans = 0;
    private DataRow[] cardsRows, accountRows, tmpDtlRows;
    private DataRow[] mainRows;
    private string CurrentPageFlag;
    private string strCardNo, strPrimaryCardNo;
    private string strForeignCurr;
    private string stmNo;
    private int totNoOfCardStat, totNoOfPageStat;

    private int curCardRow = 0, curTotCardRows = 0, numOfErr = 0, totNoOfTransactions = 0, totNoOfTransactionsInt = 0;
    private bool isPrimaryOnly, isHaveF3 = true, isPrimaryOpened = false;
    private FileStream fileStrmErr;
    private StreamWriter strmWriteErr;
    private string curMainCard, strExpiryDate, CurCardAddres1, CurCardAddres2, CurCardAddres3, CurCustomername;

    private string extAccNum;
    private string prevBranch, curBranch;
    private int totCrdNoInAcc, curCrdNoInAcc, totSuppCrdsInAcc;
    private string strOutputPath, strOutputFile, fileSummaryName;
    private DateTime vCurDate;
    private clsOMR omr = new clsOMR();
    private int totPages;
    private frmStatementFile frmMain;
    private int totRec = 1;
    private bool isSavedDataset = false;
    private string strFileNam, stmntType;
    private string arabicStr = string.Empty;
    private string curSuppCard = string.Empty;
    private string strWhereCond = string.Empty;
    private bool isVIPflag = false;
    private string curFileName = string.Empty;
    private ArrayList aryLstFiles;
    private int cntVipStat, cntVipPages, cntVipTrans, cntVipStatInt;
    private string clientstart = "*-*";
    //private string clientend = "*-*-*";
    private string clientomr = string.Empty;
    private DataRow ProductRow;
    protected bool hasInterset = false;
    private string strProductCond = string.Empty;

    private string prevDocno, curDocno;

    public clsStatementBLMEcredit()
    {
    }

    public string Statement(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pDate, string pStmntType, bool pAppendData)
    {
        string rtrnStr = "Successfully Generate " + pBankName;
        int curMonth = pDate.Month;
        bool preExit = true;
        strFileNam = pStrFileName;
        stmntType = pStmntType;
        aryLstFiles = new ArrayList();

        try
        {
            clsMaintainData maintainData = new clsMaintainData();
            maintainData.matchCardBranch4Account(pBankCode);

            // merge mark-up fee with original transaction
            maintainData.mergeMarkUpFees(pBankCode);
            maintainData = null;

            pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);
            vCurDate = pDate; //DateTime.Now.AddMonths(-1);
            strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
            clsBasFile.createDirectory(pStrFileName + vCurDate.ToString("yyyyMM") + pBankName);
            pStrFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM") + ".txt";
            strBankName = pBankName;
            strFileName = pStrFileName;
            strOutputFile = pStrFileName;
            // open output file
            if (!isVIPflag)
            {
                fileStrm = new FileStream(pStrFileName, FileMode.Create); //Create
                streamWrit = new StreamWriter(fileStrm, Encoding.Default);
                streamWrit.AutoFlush = true;
            }
            // Error file
            fileStrmErr = new FileStream(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName), FileMode.Create);
            strmWriteErr = new StreamWriter(fileStrmErr, Encoding.Default);
            // open Summary file
            fileSummaryName = pStrFileName;
            fileSummaryName = clsBasFile.getPathWithoutExtn(fileSummaryName) +
              "_Summary." + clsBasFile.getFileExtn(fileSummaryName);
            fileSummary = new FileStream(fileSummaryName, FileMode.Create); //Create
            streamSummary = new StreamWriter(fileSummary, Encoding.Default);

            // raw data file
            //fileRawData = new FileStream(clsBasFile.getPathWithoutExtn(pStrFileName) + "_RawData.txt", FileMode.Create);
            //strmRawData = new StreamWriter(fileRawData, Encoding.Default);
            //strmRawData.WriteLine("Card Number|Customer Name|Customer Address|");  //masterRow[mCardno].ToString()

            // set branch for data
            curBranchVal = pBankCode; // 4; //4  = real   1 = test
            isFullFields = false;

            //if (strWhereCond != string.Empty)
            //{
            //MainTableCond = strWhereCond;
            //supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and " + MainTableCond + ")";
            //MainTableCond = " m.contracttype in " + strProductCond + "";//strWhereCond
            MainTableCond = " m.contracttype like " + strProductCond + "";//strWhereCond
            //supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + ")";
            supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype like " + strProductCond + ")";
            //}
            strOrder = " m.CARDBRANCHPART,m.accountno,m.cardprimary desc,m.cardno ";
            FillStatementDataSetWithRemovingMarkupFee(pBankCode); //DSstatement =  //6); // 6
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
                cardsRows = mRow.GetChildRows(StatementNoDRel); //, DataRowVersion.Proposed
                strCardNo = masterRow[mCardno].ToString().Trim();
                if (masterRow[mHOLSTMT].ToString() == "Y")
                    continue;

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
                curBranch = masterRow[mCardbranchpartname].ToString().Trim();
                curBranch = curBranch.Substring(curBranch.IndexOf(")") + 1).Trim();
                if (prevBranch != curBranch)
                {
                    prevBranch = curBranch; // masterRow[mCardbranchpartname].ToString().Trim();
                }

                strPrimaryCardNo = strCardNo;
                if (masterRow[mCardprimary].ToString() == "N")
                {
                    strPrimaryCardNo = masterRow[mPrinarycardno].ToString();
                }

                //start new account
                if (prevAccountNo != masterRow[mAccountno].ToString())
                {
                    if (prevAccountNo == string.Empty)
                        if (curMonth != Convert.ToDateTime(masterRow[mStatementdateto]).Month)//,"dd/MM/yyyy"
                        {
                            preExit = false;
                            fileStrm.Close();
                            fileStrmErr.Close();
                            clsBasFile.deleteFile(strOutputFile);
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
                    CurPageRec4Dtl = 0; totNetUsage = 0; CurrentPageFlag = "F 1"; totTrans = 0; //if page is based on account no
                    calcAccountRows();

                    //>>if (totAccRows < 1
                    if (totAccRows < 2
                      && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0) //             || (masterRow[mCardno].ToString() == curMainCard   // Convert.ToDecimal(
                    {
                        isHaveF3 = true;
                        continue;
                    }

                    //check VIP
                    if (isVIPflag)
                    {
                        if (prevAccountNo != string.Empty)
                        {
                            streamWrit.Flush();
                            streamWrit.Close();
                            fileStrm.Close();
                        }
                        curFileName = clsBasFile.getPathWithoutExtn(strOutputFile)
                          + "_" + ((masterRow[mCardvip].ToString().Trim().ToUpper() == "Y") ? "VIP" : "notVIP") + "." + clsBasFile.getFileExtn(strOutputFile);
                        fileStrm = new FileStream(curFileName, FileMode.Append); //FileMode.Create Create
                        streamWrit = new StreamWriter(fileStrm, Encoding.Default); //Encoding.GetEncoding("IBM420") IBM864  ASMO-708   Encoding.Default
                        streamWrit.AutoFlush = true;
                        if (masterRow[mCardvip].ToString().Trim().ToUpper() == "Y")
                            cntVipStat++;
                        add2FileList(curFileName);
                    }
                    prevAccountNo = masterRow[mAccountno].ToString();
                    printHeader();//if page is based on account no
                    clientomr = string.Empty;
                    totNoOfCardStat++;
                    //strmRawData.WriteLine(curMainCard + "|" + CurCustomername + " " + "|" + CurCardAddres1 + " " + CurCardAddres2 + " " + CurCardAddres3);  //masterRow[mCardno].ToString()
                    if (Convert.ToDecimal(masterRow[mAccountlim].ToString()) <= Convert.ToDecimal(0))
                    {
                        strmWriteErr.WriteLine("Account Limit Less than or Equal Zero for Account " + masterRow[mAccountno].ToString());// + " and Card Number " + strCardNo
                        numOfErr++;
                    }
                } // End of if(prevAccountNo != masterRow[mAccountno].ToString())

                foreach (DataRow dRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
                {
                    detailRow = dRow;
                    stmNo = detailRow[dStatementno].ToString();
                    if ((detailRow[dPostingdate] == DBNull.Value) && (detailRow[dDocno] == DBNull.Value))
                        continue;// Exclude On-Hold Transactions 
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

                    //if (detailRow[dTrandescription].ToString().Contains("Advance Fee"))
                    //{
                    //    intinstalldocno = int.Parse(detailRow[dDocno].ToString());
                    //    intinstallamount = decimal.Parse(detailRow[dBilltranamount].ToString());
                    //}
                    //if (detailRow[dTrandescription].ToString() == "Installment repayment")
                    //{
                    //    installdocno = int.Parse(detailRow[dDocno].ToString());
                    //    installamount = decimal.Parse(detailRow[dBilltranamount].ToString());
                    //}
                    //if (intinstalldocno != 0 && installdocno != 0)
                    //{
                    //    detailRow[dBilltranamount] = intinstallamount + installamount;
                    //    installdocno = 0;
                    //    intinstalldocno = 0;
                    //    intinstallamount = 0;
                    //    installamount = 0;
                    //}

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
                                if (CurPageRec4Dtl >= MaxDetailInPage)
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

                    //-
                    printDetail();

                } //end of detail foreach
                curCrdNoInAcc++;
                if ((curAccRows >= totAccRows && totAccRows != 0) || (totAccRows == 0 && curCrdNoInAcc == totCrdNoInAcc))
                {
                    completePageDetailRecords();
                    printCardFooter();//if pages is based on account
                    //printAccountFooter();
                    CurPageRec4Dtl = 0; //>pageNo=1; if pages is based on account
                    curAccRows = 0;
                }
            } //end of Master foreach
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

                //strmRawData.Flush();
                //strmRawData.Close();
                //fileRawData.Close();

                if (numOfErr == 0)
                    clsBasFile.deleteFile(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName));
                else
                    aryLstFiles.Add(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err.txt");
                //aryLstFiles.Add(clsBasFile.getPathWithoutExtn(pStrFileName) + "_RawData.txt");
                if (!isVIPflag)
                    aryLstFiles.Add(strOutputFile);
                numOfErr = validateNoOfLines(aryLstFiles, 61);
                if (numOfErr > 0)
                    rtrnStr += string.Format(" with {0} Errors", numOfErr);
                aryLstFiles.Add(@fileSummaryName);

                ArrayList tmpAryLstFiles = new ArrayList();
                aryLstFiles.Sort();
                Object objPrrv = string.Empty;
                foreach (Object obj in aryLstFiles)
                {
                    if (obj.ToString() != objPrrv.ToString())
                        tmpAryLstFiles.Add(obj);//aryLstFiles.Remove(obj);
                    objPrrv = obj;
                }
                aryLstFiles = tmpAryLstFiles;

                clsBasFile.generateFileMD5(aryLstFiles, @clsBasFile.getPathWithoutExtn(pStrFileName) + ".MD5");
                aryLstFiles.Add(@clsBasFile.getPathWithoutExtn(pStrFileName) + ".MD5");
                SharpZip zip = new SharpZip();
                zip.createZip(aryLstFiles, @clsBasFile.getPathWithoutExtn(pStrFileName) + ".zip", "");

                DSstatement.Dispose();
            }
        }
        return rtrnStr;
    }


    protected void printHeader()
    {
        string newaddress1, newaddress2;
        clsMaintainData.fixAddress(curBranchVal, masterRow[mCustomeraddress1].ToString(), out newaddress1, out newaddress2);
        totPages++;
        extAccNum = clsBasValid.validateStr(masterRow[mExternalno]);
        if (extAccNum.Trim() == "")
            extAccNum = clsBasValid.validateStr(masterRow[mAccountno]);

        if (pageNo == 1 && totalAccPages == 1) // statement contain 1 page
        {
            CurrentPageFlag = "F 0";
            isHaveF3 = true;
            clientomr = clientstart;
        }
        else if (pageNo == 1 && totalAccPages > 1)  //first page of multiple page statement
        {
            CurrentPageFlag = "F 1"; // //middle page of multiple page statement
            clientomr = clientstart;
        }
        else if (pageNo < totalAccPages)
        {
            CurrentPageFlag = "F 2";
            clientomr = string.Empty;
        }
        else if (pageNo == totalAccPages) //last page of multiple page statement
        {
            CurrentPageFlag = "F 3";
            isHaveF3 = true;
            clientomr = string.Empty;
            //clientomr = clientend;
        }

        //			streamWrit.WriteLine(basText.alignmentLeft(masterRow[mCustomername],50)+ basText.replicat(" ",31) + masterRow[mCardbranchpartname]);  //
        //			streamWrit.WriteLine(basText.alignmentMiddle(masterRow[mCustomeraddress1],50)+ basText.replicat(" ",31) + clsBasValid.validateStr(masterRow[mAccountno])); //basText.formatNum(masterRow[mCardno],"####-####-####-####")
        //string x = Convert.ToString(ValidateArbic(masterRow[mCustomeraddress1].ToString()));
        //			streamWrit.WriteLine(basText.alignmentMiddle(masterRow[mCustomeraddress2],50)+ basText.replicat(" ",31)+"{0,10:dd/MM/yyyy}",masterRow[mStatementdatefrom]);//+ basText.formatNum(masterRow[mAccountlim],"########")  basText.formatNum(masterRow[mCardlimit],"########")
        //			streamWrit.WriteLine(basText.alignmentMiddle(masterRow[mCustomeraddress3].ToString().Trim() + " " + masterRow[mCustomerregion].ToString().Trim() + " " + masterRow[mCustomercity].ToString().Trim(),50)+ basText.replicat(" ",31)+pageNo + " / " + totalAccPages);  //basText.formatNum(masterRow[mAccountno],"##-##-######-####")
        streamWrit.WriteLine(strEndOfPage);
        //streamWrit.WriteLine(clientomr);//new omr
        streamWrit.WriteLine("*****");
        //streamWrit.WriteLine(String.Empty);//old omr
        streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(String.Empty);
        //streamWrit.WriteLine((basText.isContainArbic(ValidateArbic(masterRow[mCustomeraddress1].ToString().Trim() + masterRow[mCustomeraddress2].ToString().Trim() + masterRow[mCustomeraddress3].ToString().Trim())) ? "A" : "E") + basText.replicat(" ", 84) + "Page " + pageNo.ToString() + " of " + totalAccPages.ToString());  //basText.formatNum(masterRow[mAccountno],"##-##-######-####")
        //streamWrit.WriteLine("*" + basText.alignmentLeft(masterRow[mAccountzipcode], 5) + "*");  //old omr
        streamWrit.WriteLine(omr.AsteriskLastPageBLME(pageNo, totalAccPages));  // new omr
        streamWrit.WriteLine(basText.replicat(" ", 15) + basText.alignmentLeft(masterRow[mCustomername], 30) + basText.replicat(" ", 33) + basText.alignmentLeft(basText.formatCardNumber(curMainCard), 16));  //
        //arabicStr = basText.isContainArbic(ValidateArbic(masterRow[mCustomeraddress1].ToString())) ? "A" : "E";
        //streamWrit.WriteLine(basText.replicat(" ", 15) + arabicStr + basText.alignmentLeft(ValidateArbic(masterRow[mCustomeraddress1].ToString()), 50)); //basText.formatNum(masterRow[mCardno],"####-####-####-####")
        streamWrit.WriteLine(basText.replicat(" ", 15) + arabicStr + basText.alignmentLeft(ValidateArbic(newaddress1), 50)); //basText.formatNum(masterRow[mCardno],"####-####-####-####")
        //arabicStr = basText.isContainArbic(ValidateArbic(masterRow[mCustomeraddress2].ToString())) ? "A" : "E";
        //streamWrit.WriteLine(basText.replicat(" ", 15) + arabicStr + basText.alignmentLeft(ValidateArbic(masterRow[mCustomeraddress2].ToString()), 50) + basText.replicat(" ", 12) + basText.formatDate(masterRow[mStatementdateto], "dd/MM/yyyy"));
        streamWrit.WriteLine(basText.replicat(" ", 15) + arabicStr + basText.alignmentLeft(ValidateArbic(newaddress2), 50) + basText.replicat(" ", 12) + basText.formatDate(masterRow[mStatementdateto], "dd/MM/yyyy"));
        //arabicStr = basText.isContainArbic(ValidateArbic(masterRow[mCustomeraddress3].ToString().Trim())) ? "A" : "E";
        streamWrit.WriteLine(basText.replicat(" ", 15) + arabicStr + basText.alignmentLeft(ValidateArbic(masterRow[mCustomeraddress3].ToString().Trim()), 50));  //basText.formatNum(masterRow[mAccountno],"##-##-######-####")
        streamWrit.WriteLine(basText.replicat(" ", 15) + arabicStr + basText.alignmentLeft(ValidateArbic(masterRow[mBarcode].ToString().Trim() + " " + masterRow[mUserActField1].ToString().Trim()), 50));  //basText.formatNum(masterRow[mAccountno],"##-##-######-####")
        //streamWrit.WriteLine(basText.replicat(" ", 77) + basText.formatNum(masterRow[mCardlimit], "########", 16, "L"));  //cardproduct+"{0,5:MM/yy}",masterRow[mStatementdatefrom]
        streamWrit.WriteLine(basText.replicat(" ", 77) + basText.formatNum(masterRow[mAccountlim], "########", 16, "L"));  //cardproduct+"{0,5:MM/yy}",masterRow[mStatementdatefrom]
        //if (curBranchVal == 41) //BDCA
        //streamWrit.WriteLine(basText.replicat(" ", 15) + basText.alignmentLeft(masterRow[mDept], 70));//String.Empty
        //else //BMSR
        //streamWrit.WriteLine(basText.alignmentLeft(masterRow[mCardbranchpart].ToString(), 20));//String.Empty
        streamWrit.WriteLine(basText.alignmentRight(masterRow[mAccountzipcode].ToString(), 42));  //cardproduct+"{0,5:MM/yy}",masterRow[mStatementdatefrom]
        streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(String.Empty);
        //streamWrit.WriteLine(omr.Asterisk(totPages) + omr.AsteriskLastPage(pageNo, totalAccPages));//old omr //omr.Asterisk(pageNo) + omr.AsteriskPage4(pageNo, totalAccPages)  //String.Empty
        //streamWrit.WriteLine(omr.AsteriskBMSR(totPages));  // new omr
        //streamWrit.WriteLine(omr.ParityCheck(totPages, clientomr));// new omr
        //streamWrit.WriteLine(String.Empty);//old omr
        streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(String.Empty);

        if (pageNo == 1)//&& CurPageRec4Dtl == 0
        {
            streamWrit.WriteLine(basText.replicat(" ", 24) + basText.alignmentLeft("Previous Balance", 70) + basText.alignmentRight((Object)(basText.formatNumUnSign(masterRow[mOpeningbalance], "#,###0.00", 15)), 15) + "  " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance])));
            CurPageRec4Dtl++; curAccRows++;
            totTrans = Convert.ToDecimal(masterRow[mOpeningbalance]);
        }
        //    streamWrit.WriteLine(basText.replicat(" ", 81) + CurrentPageFlag);  //CurrentPageFlag+"{0,5:MM/yy}",masterRow[mStatementdatefrom]
        //streamWrit.WriteLine(basText.replicat(" ", 81) + basText.formatNum(curMainCard, "#### #### #### ####"));  //cardbranchpartname
        //streamWrit.WriteLine(basText.replicat(" ", 81) + basText.formatNum(masterRow[mAccountlim], "########")); //extAccNum   clsBasValid.validateStr(masterRow[mExternalno]) //accountno  //basText.formatNum(masterRow[mCardno],"####-####-####-####")
        //streamWrit.WriteLine(basText.replicat(" ", 81) + basText.alignmentLeft(masterRow[mMindueamount], 13));//"{0,10:dd/MM/yyyy}", masterRow[mStatementdateto])   Statementdatefrom+ basText.formatNum(masterRow[mAccountlim],"########")  basText.formatNum(masterRow[mCardlimit],"########")
        //>>    streamWrit.WriteLine(basText.replicat(" ", 81) + basText.alignmentLeft(pageNo,5) + basText.alignmentLeft(totalAccPages,5));  //basText.formatNum(masterRow[mAccountno],"##-##-######-####")
        //streamWrit.WriteLine(" " + omr.Line2LastPage(pageNo, totalAccPages));//String.Empty
        //streamWrit.WriteLine(" " + omr.Line3(pageNo));//String.Empty
        //streamWrit.WriteLine(" " + omr.Line4(pageNo));//String.Empty
        //streamWrit.WriteLine(" " + omr.fixLine());//String.Empty
        //    streamWrit.WriteLine(basText.alignmentMiddle(curMainCard, 20) + basText.alignmentMiddle(masterRow[mAccountlim], 13) + basText.formatNum(masterRow[mAccountavailablelim], "##0", 20) + basText.alignmentMiddle(masterRow[mMindueamount], 13) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M") + basText.formatNum(masterRow[mTotaloverdueamount], "##0.00", 13)); //+ basText.formatNum(masterRow[mAccountlim],"#,##0.00",9)+" "+basText.formatNum(masterRow[mTotalcredits],"#,##0.00",9)+" "+basText.formatNum(masterRow[mAccountavailablelim],"#,##0.00",9)+" "+basText.formatNum(masterRow[mMindueamount],"#,##0.00",15)+" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",15)+" "+basText.formatDate(masterRow[mStetementduedate],"dd/MM/yy") //TOTALCREDITS// basText.alignmentMiddle(masterRow[mAccountavailablelim],20) //"#,##0"
        //    streamWrit.WriteLine(String.Empty);
        //    streamWrit.WriteLine(String.Empty);
        //streamWrit.WriteLine(String.Empty);
        //    streamWrit.WriteLine(String.Empty);
        //    if (pageNo == 1)
        //      streamWrit.WriteLine(basText.replicat(" ", 36) + basText.alignmentLeft("Previous Balance", 63) + basText.formatNumUnSign(masterRow[mOpeningbalance], "#,###0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance])));  //+ CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign]))
        //    else
        //      streamWrit.WriteLine(String.Empty);
        //    streamWrit.WriteLine(String.Empty);

        totalPages++;
        totNoOfPageStat++;
        if (isVIPflag && masterRow[mCardvip].ToString().Trim().ToUpper() == "Y")
            cntVipPages++;
    }


    protected void printDetail()
    {
        DateTime trnsDate = clsBasValid.validateDate(detailRow[dTransdate]), postingDate = clsBasValid.validateDate(detailRow[dPostingdate]);
        if (trnsDate > postingDate)
            trnsDate = postingDate;

        if (masterRow[mAccountcurrency].ToString().Trim() != detailRow[dOrigtrancurrency].ToString().Trim())
            strForeignCurr = basText.alignmentRight((Object)basText.formatNum(detailRow[dOrigtranamount], "#0.00;(#0.00)") + " " + basText.Replace(clsBasValid.validateStr(detailRow[dOrigtrancurrency]), "XXX", "   "), 11);
        else
            strForeignCurr = basText.replicat(" ", 11);

        if (strForeignCurr.Trim() == "0")
            strForeignCurr = basText.replicat(" ", 11);

        string trnsDesc; //= detailRow[dTrandescription].ToString().Trim() + " " + detailRow[dMerchant].ToString().Trim()
        if (detailRow[dMerchant].ToString().Trim() == "")
            trnsDesc = detailRow[dTrandescription].ToString().Trim();
        else
            //if (detailRow[dTrandescription].ToString().Trim() == "Deposit")
            //    //trnsDesc = detailRow[dMerchant].ToString().Trim() + "," + detailRow[dTrandescription].ToString().Trim();
            //    trnsDesc = "Cash deposit";
            //else
            trnsDesc = detailRow[dMerchant].ToString().Trim(); // default

        ////MAMR Update
        //if (detailRow[dInstallmentData].ToString().Trim() != "")
        //    trnsDesc = detailRow[dInstallmentData].ToString().Trim();
        //else if (detailRow[dMerchant].ToString().Trim() != "")
        //    trnsDesc = detailRow[dMerchant].ToString().Trim();
        //else
        //    //if (detailRow[dTrandescription].ToString().Trim() == "Deposit")
        //    //    //trnsDesc = detailRow[dMerchant].ToString().Trim() + "," + detailRow[dTrandescription].ToString().Trim();
        //    //    trnsDesc = "Cash deposit";
        //    //else
        //    trnsDesc = detailRow[dTrandescription].ToString().Trim(); // default

        ////MAMR Update
        string str1 = "";
        string str2 = "";

        try
        {
            str1 = detailRow[dBilltranamount].ToString();
            str2 = detailRow[dBilltranamountsign].ToString();
            totTrans += valDbCr(Convert.ToDecimal(detailRow[dBilltranamount]), clsBasValid.validateStr(detailRow[dBilltranamountsign])); //basText.formatNum(detailRow[dBilltranamount], "#0.00;(#0.00)", 19);
        }
        catch (Exception exx)
        {
            string str = str1 + str2; //exx.Message;
        }
        streamWrit.WriteLine(trnsDate.ToString("dd/MM/yyyy") + " " + postingDate.ToString("dd/MM/yyyy") + "  " + basText.alignmentLeft(trnsDesc, 46) + " " + basText.alignmentRight(strForeignCurr, 11) + basText.alignmentRight((object)basText.formatNumUnSign(detailRow[dBilltranamount], "#,###0.00", 12), 12) + " " + CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + " " + basText.alignmentRight((object)basText.formatNumUnSign((object)totTrans, "#,###0.00", 11), 11) + "  " + CrDb(totTrans)); //+ "  " + isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString())) + "  " + (isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString())) == "SUP" ? clsBasValid.validateStr(detailRow[dCardno]) : " ".PadRight(16, ' '))
        //streamWrit.WriteLine(String.Format(trnsDate.ToString("dd/MM/yyyy"), "dd/MM/yyyy") + " " + String.Format(postingDate.ToString(), "dd/MM/yyyy") + "  " + basText.trimStr(trnsDesc, 45) + " " + strForeignCurr + " " + basText.formatNum(detailRow[dBilltranamount], "#0.00;(#0.00)", 19) + " " + CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + " " + basText.formatNum((object)totTrans, "#0.00;(#0.00)", 19) + " " + isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))); 
        //streamWrit.WriteLine("{0:dd/MM/yyyy} {1:dd/MM/yyyy}  {2,-45} {3,15} {4,19} {5,5}", String.Format( trnsDate, postingDate, basText.trimStr(trnsDesc, 45), strForeignCurr, basText.formatNum(detailRow[dBilltranamount], "#0.00;(#0.00)", 19), CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + " " + basText.formatNum(detailRow[dBilltranamount], "#0.00;(#0.00)", 19) + " " + isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)
        //streamWrit.WriteLine("{0:dd/MM/yyyy}  {1:dd/MM/yyyy}  {2,-24}  {3,-40} {4,16} {5,16} {6,2}", trnsDate, postingDate, basText.trimStr(detailRow[dRefereneno].ToString().Trim(), 24), basText.trimStr(trnsDesc, 40), strForeignCurr, basText.formatNum(detailRow[dBilltranamount], "#0.00;(#0.00)"), CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + " " + isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)
        //			streamWrit.WriteLine("  {0:dd/MM}  {1:dd/MM}  {2,-40} {3,12} {4,3} {5,16} {6,2}",trnsDate,postingDate,basText.trimStr(detailRow[dTrandescription].ToString().Trim() + " " + detailRow[dMerchant].ToString().Trim(),45),basText.formatNum(detailRow[dOrigtranamount],"#,##0.00;(#,##0.00)"),basText.Replace(clsBasValid.validateStr(detailRow[dOrigtrancurrency]),"XXX","   "),basText.formatNum(detailRow[dBilltranamount],"#,##0.00;(#,##0.00)"),CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)
        //streamWrit.WriteLine("{0:dd/MM}  {1,-24}  {2,-40} {3,16} {4,16} {5,2}", trnsDate, basText.trimStr(detailRow[dRefereneno].ToString().Trim(), 24), basText.trimStr(trnsDesc, 40), strForeignCurr, basText.formatNum(detailRow[dBilltranamount], "#0.00;(#0.00)"), CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + " " + isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)
        totNoOfTransactions++;
        if (isVIPflag && masterRow[mCardvip].ToString().Trim().ToUpper() == "Y")
            cntVipTrans++;
    }

    protected void printCardFooter()
    {
        streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(basText.replicat(" ", 12) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 38, "L") + basText.alignmentLeft((object)basText.formatNumUnSign(masterRow[mMindueamount], "#,##0.00", 44), 44) + basText.alignmentRight((Object)basText.formatNumUnSign(masterRow[mClosingbalance], "#,##0.00", 19), 19) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])));
        streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(basText.replicat(" ", 10) + basText.alignmentLeft(basText.formatCardNumber(curMainCard), 16));
        streamWrit.WriteLine(String.Empty);
        //    streamWrit.WriteLine(basText.replicat(" ", 10) + basText.alignmentLeft(basText.formatNumUnSign(masterRow[mTotalpayments], "#,##0.00", 12) + " " + DbCr(Convert.ToDecimal(masterRow[mTotalpayments])), 15));
        streamWrit.WriteLine(basText.replicat(" ", 10) + basText.alignmentLeft((object)(masterRow[mCardpaymentmethod].ToString().Trim() == "" || masterRow[mCardpaymentmethod].ToString().Trim() == "Not Defined" ? "Cash" : masterRow[mCardpaymentmethod].ToString().Trim()), 30)); //, "#,##0.00", 12) + " " + DbCr(Convert.ToDecimal(masterRow[mTotalpayments])), 15)
        streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(basText.replicat(" ", 10) + basText.alignmentLeft((Object)basText.formatNumUnSign(masterRow[mClosingbalance], "#,##0.00", 15) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])), 20));
        streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(basText.replicat(" ", 10) + basText.formatNum(masterRow[mMindueamount], "#,##0.00", 28, "L"));
        streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(basText.replicat(" ", 10) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 37, "L"));
        //completePageDetailRecords();
        //if (pageNo == totalAccPages)
        //  streamWrit.WriteLine(basText.replicat(" ", 17) + basText.alignmentLeft("Current Balance", 63) + basText.formatNumUnSign(masterRow[mClosingbalance], "#0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])));
        //else
        //  streamWrit.WriteLine(String.Empty);
        //    streamWrit.WriteLine(basText.alignmentRight(curMainCard, 35) + basText.alignmentRight(masterRow[mMindueamount], 20) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "R")
        //   + basText.alignmentRight(basText.formatNumUnSign(masterRow[mOpeningbalance], "##0.00", 12) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance].ToString())), 15)
        //   + basText.alignmentRight(basText.formatNum(masterRow[mTotalpayments], "#,##0.00", 12) + " " + DbCr(Convert.ToDecimal(masterRow[mTotalpayments])), 15)
        //   + basText.alignmentRight(basText.formatNum(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal]), "#,##0.00", 12) + CrDbMinus(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal])), 15)
        //   + basText.alignmentRight(basText.formatNum(masterRow[mTotalcharges].ToString(), "#,##0.00", 12) + " " + CrDbMinus(Convert.ToDecimal(masterRow[mTotalcharges])), 15)
        //   + basText.alignmentRight(basText.formatNum(masterRow[mTotalinterest].ToString(), "#,##0.00", 12) + " " + CrDbMinus(Convert.ToDecimal(masterRow[mTotalinterest])), 15)
        //       + basText.formatNum(Convert.ToDecimal(masterRow[mTotalcharges]) + Convert.ToDecimal(masterRow[mTotalinterest]),"#,##0.00",12) + CrDbMinus(Convert.ToDecimal(masterRow[mTotalcharges]) + Convert.ToDecimal(masterRow[mTotalinterest])) 
        //streamWrit.WriteLine(String.Empty);
        //streamWrit.WriteLine(String.Empty);
        //    streamWrit.WriteLine(basText.alignmentLeft(basText.formatNumUnSign(masterRow[mOpeningbalance], "#,###,##0.00", 12) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance].ToString())), 20) +
        //      basText.alignmentLeft(basText.formatNum(masterRow[mTotalcharges].ToString(), "#,###,##0.00", 12), 20) +
        //      basText.alignmentLeft(basText.formatNumUnSign(Convert.ToDecimal(masterRow[mTotalpurchases]) +
        //      Convert.ToDecimal(masterRow[mTotalcashwithdrawal]), "#,###,##0.00", 16) + " " +
        //      CrDbMinus(Convert.ToDecimal(masterRow[mTotalpurchases]) +
        //      Convert.ToDecimal(masterRow[mTotalcashwithdrawal])), 20) +
        //basText.alignmentLeft(basText.formatNumUnSign(masterRow[mClosingbalance], "#,###,##0.00", 12) + " " +
        //CrDb(Convert.ToDecimal(masterRow[mClosingbalance])), 20));
        //streamWrit.WriteLine(String.Empty);
        //streamWrit.WriteLine(String.Empty);
        //streamWrit.WriteLine(basText.replicat(" ", 81) + basText.alignmentLeft(masterRow[mExternalno], 20));  //cardproduct+"{0,5:MM/yy}",masterRow[mStatementdatefrom]
        //streamWrit.WriteLine(basText.alignmentLeft(masterRow[mCustomername], 50));  //
        //streamWrit.WriteLine(basText.replicat(" ", 81) + basText.formatNum(curMainCard, "#### #### #### ####"));  //cardbranchpartname
        //streamWrit.WriteLine(basText.alignmentLeft(ValidateArbic(masterRow[mCustomeraddress1].ToString()), 50)); //basText.formatNum(masterRow[mCardno],"####-####-####-####")
        //streamWrit.WriteLine(basText.alignmentLeft(ValidateArbic(masterRow[mCustomeraddress2].ToString()), 50));
        //streamWrit.WriteLine(basText.alignmentLeft(ValidateArbic(masterRow[mCustomeraddress3].ToString().Trim()) + " " + masterRow[mCustomerregion].ToString().Trim() + " " + masterRow[mCustomercity].ToString().Trim(), 50));  //basText.formatNum(masterRow[mAccountno],"##-##-######-####")
        //streamWrit.WriteLine(basText.replicat(" ", 81) + basText.alignmentLeft(basText.formatNumUnSign(masterRow[mTotalpayments], "#,##0.00", 12) + " " + DbCr(Convert.ToDecimal(masterRow[mTotalpayments])), 15));  //cardproduct+"{0,5:MM/yy}",masterRow[mStatementdatefrom]
        //streamWrit.WriteLine(String.Empty);

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
        stmNo = masterRow[mAccountno].ToString();

        accountRows = DSstatement.Tables["tStatementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[mAccountno]).ToString().Trim() + "'");
        string prevCardNo = String.Empty, CurCardNo = String.Empty;
        int currAccRowsPages = 1;//0
        totAccRows = 1;//0
        totalAccPages = 0;
        //currAccRowsPages = 0;
        foreach (DataRow dtAccRow in accountRows) //mRow.GetChildRows(StatementNoDRel)
        {
            if ((dtAccRow[dPostingdate] == DBNull.Value) && (dtAccRow[dDocno] == DBNull.Value))
                continue;
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
        curMainCard = string.Empty;
        totCrdNoInAcc = curCrdNoInAcc = 0;
        mainRows = DSstatement.Tables["tStatementMasterTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[mAccountno]) + "'");
        curMainCard = CurCardNo = strExpiryDate = "";
        foreach (DataRow mainRow in mainRows) //mRow.GetChildRows(StatementNoDRel)
        {
            totCrdNoInAcc++;
            CurCardNo = mainRow[mCardno].ToString();
            CurCardAddres1 = mainRow[mCustomeraddress1].ToString();
            CurCardAddres2 = mainRow[mCustomeraddress2].ToString();
            CurCardAddres3 = mainRow[mCustomeraddress3].ToString() + " " + mainRow[mCustomerregion].ToString() + " " + mainRow[mCustomercity].ToString();
            CurCustomername = mainRow[mCustomername].ToString();
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
                strExpiryDate = mainRow[mCardexpirydate].ToString();
                CurCardAddres1 = mainRow[mCustomeraddress1].ToString();
                CurCardAddres2 = mainRow[mCustomeraddress2].ToString();
                CurCardAddres3 = mainRow[mCustomeraddress3].ToString() + " " + mainRow[mCustomerregion].ToString() + " " + mainRow[mCustomercity].ToString();
                CurCustomername = mainRow[mCustomername].ToString();
            }
            if (mainRow[mCardprimary].ToString() == "Y" && isValidateCard(mainRow[mCardstate].ToString()))
            {
                curMainCard = CurCardNo; //mainRow[mCardno].ToString();
                strExpiryDate = mainRow[mCardexpirydate].ToString();
                CurCardAddres1 = mainRow[mCustomeraddress1].ToString();
                CurCardAddres2 = mainRow[mCustomeraddress2].ToString();
                CurCardAddres3 = mainRow[mCustomeraddress3].ToString() + " " + mainRow[mCustomerregion].ToString() + " " + mainRow[mCustomercity].ToString();
                CurCustomername = mainRow[mCustomername].ToString();
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


    private void completePageDetailRecords()
    {
        //int curPageLine =CurPageRec4Dtl;
        if (pageNo == totalAccPages)
            for (int curPageLine = CurPageRec4Dtl; curPageLine < MaxDetailInPage; curPageLine++)
                streamWrit.WriteLine(String.Empty);
    }


    private void printStatementSummary()
    {
        if (strBankName.Contains("MasterCard"))
            streamSummary.WriteLine(strBankName + " MasterCard Statement");
        else
            streamSummary.WriteLine(strBankName + " Visa Statement");
        streamSummary.WriteLine("__________________________");
        streamSummary.WriteLine("");
        streamSummary.WriteLine("No of Statements   " + totNoOfCardStat.ToString());
        streamSummary.WriteLine("No of Pages        " + totNoOfPageStat.ToString());
        streamSummary.WriteLine("No of Transactions " + totNoOfTransactions.ToString());

        if (isVIPflag)
        {
            streamSummary.WriteLine("_________________________________________________");
            streamSummary.WriteLine("VIP Cards");
            streamSummary.WriteLine("");
            streamSummary.WriteLine("No of Statements   " + cntVipStat.ToString());
            streamSummary.WriteLine("No of Pages        " + cntVipPages.ToString());
            streamSummary.WriteLine("No of Transactions " + cntVipTrans.ToString());
        }

        //    clsValidatePageSize ValidatePageSize = new clsValidatePageSize(strFileName, 54, strEndOfPage);
        //    streamSummary.WriteLine(ValidatePageSize.outMessage);

        clsStatementSummary StatSummary = new clsStatementSummary();
        StatSummary.BankCode = curBranchVal;
        StatSummary.BankName = strBankName;
        StatSummary.StatementDate = vCurDate;
        StatSummary.CreationDate = DateTime.Now;
        DataRow[] productRows;
        prevAccountNo = String.Empty;
        totNoOfCardStat = 0;
        totNoOfTransactionsInt = 0;
        cntVipStat = 0;
        cntVipStatInt = 0;
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
                curBranch = masterRow[mCardbranchpartname].ToString().Trim();
                curBranch = curBranch.Substring(curBranch.IndexOf(")") + 1).Trim();
                if (prevBranch != curBranch)
                {
                    prevBranch = curBranch; // masterRow[mCardbranchpartname].ToString().Trim();
                }
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
                if (prevAccountNo != masterRow[mAccountno].ToString())
                {
                    if (prevAccountNo == string.Empty)
                        curMainCard = string.Empty;
                    pageNo = 1;//totalAccPages = 1 ; pageNo=1;
                    CurPageRec4Dtl = 0; totNetUsage = 0; CurrentPageFlag = "F 1"; //if page is based on account no
                    calcAccountRows();
                    if (totAccRows < 2 && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0) //             || (masterRow[mCardno].ToString() == curMainCard   // Convert.ToDecimal(
                    {
                        continue;
                    }
                    if (isVIPflag)
                    {
                        if (masterRow[mCardvip].ToString().Trim().ToUpper() == "Y")
                            cntVipStat++;
                    }
                    prevAccountNo = masterRow[mAccountno].ToString();
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
                    if (curSuppCard != detailRow[dCardno].ToString())
                    {
                        tmpDtlRows = DSstatement.Tables["tStatementMasterTable"].Select(mCardno + " = '" + detailRow[dCardno] + "'");
                        if (tmpDtlRows.Length > 0)
                        {
                            if (tmpDtlRows[0][mCardprimary].ToString() == "N" && isValidateCard(tmpDtlRows[0][mCardstate].ToString()))
                            {
                                CurPageRec4Dtl++;
                                curCrdNoInAcc++;
                                curAccRows++;
                                if (CurPageRec4Dtl >= MaxDetailInPage)
                                {
                                    CurPageRec4Dtl = 1;
                                    pageNo++;
                                }
                            }
                        }
                    }
                    curSuppCard = detailRow[dCardno].ToString();
                    CurPageRec4Dtl = CurPageRec4Dtl + 1;
                    hasInterset = CalcTransInterest(detailRow[dTrandescription].ToString());
                    if (hasInterset)
                    {
                        if (isVIPflag && masterRow[mCardvip].ToString().Trim().ToUpper() == "Y")
                            cntVipStatInt++;
                        else
                            totNoOfTransactionsInt++;
                        break;
                    }

                } //end of detail foreach
                curCrdNoInAcc++;
                if ((curAccRows >= totAccRows && totAccRows != 0) || (totAccRows == 0 && curCrdNoInAcc == totCrdNoInAcc))
                {
                    CurPageRec4Dtl = 0; //>pageNo=1; if pages is based on account
                    curAccRows = 0;
                }
            }
            if (totNoOfCardStat != 0)
            {
                StatSummary.NoOfStatements = totNoOfCardStat + cntVipStat;
                StatSummary.NoOfTransactionsInt = totNoOfTransactionsInt + cntVipStatInt;
                StatSummary.ProductCode = int.Parse(ProductRow["code"].ToString());
                StatSummary.ProductName = ProductRow["name"].ToString();
                StatSummary.InsertRecordDb(vCurDate.ToString("yyyyMM") + strBankName);
                totNoOfCardStat = 0;
                totNoOfTransactionsInt = 0;
                cntVipStat = 0;
                cntVipStatInt = 0;
            }
        }
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

    private void add2FileList(string pFileName)
    {
        int myIndex = aryLstFiles.BinarySearch((object)pFileName);
        if (myIndex < 0)
            aryLstFiles.Add(pFileName);
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

    public bool savedDataset
    {
        get { return isSavedDataset; }
        set { isSavedDataset = value; }
    }// savedDataset

    public string whereCond
    {
        get { return strWhereCond; }
        set { strWhereCond = value; }
    }  // whereCond

    public bool isVIP
    {
        get { return isVIPflag; }
        set { isVIPflag = value; }
    }  // isVIP

    public string productCond
    {
        get { return strProductCond; }
        set { strProductCond = value; }
    }  // productCond

    ~clsStatementBLMEcredit()
    {
        DSstatement.Dispose();
    }
}
