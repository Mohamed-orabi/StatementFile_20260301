using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Xml;
using System.Collections;


// Branch 3
public class clsStatementAwbPlus : clsBasStatement
    {
    private string strBankName;
    private FileStream fileStrm, fileSummary;
    private StreamWriter streamWrit, streamSummary;
    //private DataSet DSstatement;
    //private OracleDataReader drPrimaryCards, drMaster,drDetail;
    private DataRow masterRow;
    private DataRow detailRow;
    protected DataRow rewardRow;
    private string strEndOfLine = "\u000D";  //+ "M" ^M
    //		private string strEndOfPage = "\u000C"  ;  //+ "\u000D"+ "M" ^L + ^M
    private string strEndOfPage = "\u000C";  //  + basText.replicat(" ",85) + "F 2"  + "\u000D"+ "M" ^L + ^M
    private const int MaxDetailInPage = 33; //
    private const int MaxDetailInLastPage = 27; //
    private int CurPageRec4Dtl = 0;
    private int pageNo = 0, totalPages = 0, totalCardPages = 0
      , totalAccPages = 0, totCardRows = 0, totAccRows = 0;
    private string lastPageTotal;
    //  private string curCardNo ;//,PrevCardNo
    private string curAccountNo, prevAccountNo = String.Empty;//,PrevCardNo
    private decimal totNetUsage = 0;
    private DataRow[] cardsRows, accountRows, rewardRows;
    private DataRow[] mainRows;
    private string CurrentPageFlag;
    private string strForeignCurr;
    private int curCardRow = 0, curTotCardRows = 0, numOfErr = 0;
    private bool isPrimaryOnly, isHaveF2 = true, isPrimaryOpened = false;
    private FileStream fileStrmErr;
    private StreamWriter strmWriteErr;
    private string curMainCard;
    private string extAccNum;
    private string strOutputPath, strOutputFile, fileSummaryName;
    DateTime vCurDate;
    private frmStatementFile frmMain;
    private int totRec = 1;
    protected string rewardCond = "'Reward Program'";//'New Reward Contract'

    private int totCards, curCardCounter;
    bool isAccFoterPrnt = true;
    private string strProductCond = string.Empty;
    private string stmntType;
    private int totNoOfCardStat, totNoOfPageStat, totNoOfTransactions = 0, totNoOfTransactionsInt = 0;
    private DataRow ProductRow;
    protected bool hasInterset = false;
    protected bool isReward = false;

    public clsStatementAwbPlus()
        {
        }

    public string Statement(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pCurDate, string pStmntType, bool pAppendData)
        {
        string rtrnStr = "Successfully Generate " + pBankName;
        int curMonth = pCurDate.Month;
        bool preExit = true;
        strBankName = pBankName;
        stmntType = pStmntType;
        try
            {
            pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);


            vCurDate = pCurDate; //DateTime.Now.AddMonths(-1);
            strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
            clsBasFile.createDirectory(pStrFileName + vCurDate.ToString("yyyyMM") + pBankName);
            pStrFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM") + ".txt";
            strBankName = pBankName;
            strOutputFile = pStrFileName;
            // open output file
            fileStrm = new FileStream(pStrFileName, FileMode.Create);
            streamWrit = new StreamWriter(fileStrm, Encoding.Default); //Encoding.GetEncoding("ASMO-708") ASMO-708  iso-8859-6  IBM864  IBM420  DOS-720  windows-1256  x-mac-arabic Encoding.GetEncoding("x-mac-arabic") Encoding.Default
            streamWrit.AutoFlush = true;

            // error file
            fileStrmErr = new FileStream(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName), FileMode.Create);
            strmWriteErr = new StreamWriter(fileStrmErr, Encoding.ASCII);
            strmWriteErr.AutoFlush = true;

            fileSummaryName = pStrFileName;
            fileSummaryName = clsBasFile.getPathWithoutExtn(fileSummaryName) +
              "_Summary." + clsBasFile.getFileExtn(fileSummaryName);

            // open Summary file
            fileSummary = new FileStream(fileSummaryName, FileMode.Create); //Create
            streamSummary = new StreamWriter(fileSummary, Encoding.Default);
            streamSummary.AutoFlush = true;

            // set branch for data
            curBranchVal = pBankCode; // pBankCode; // 3; //3 = real   1 = test
            //strOrder = " m.CARDBRANCHPART,m.accountno,m.cardprimary,m.cardstatus,m.cardno ";
            strOrder = " m.CARDBRANCHPART,m.accountno,m.cardprimary desc,m.cardstatus,m.cardno ";
            isFullFields = false;
            //MainTableCond = " m.contracttype = '" + strProductCond + "'";//strWhereCond
            //MainTableCond = " instr(m.contracttype,'" + strProductCond + "') > 0 ";//strWhereCond
            MainTableCond = " instr(m.contracttype,'" + strProductCond + "') > 0 ";//strWhereCond
            //supTableCond = " d.statementno in (select x.statementno from tstatementmastertable x where x.branch = " + pBankCode + " and x.contracttype = '" + strProductCond + "')";
            //supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and instr(x.contracttype,'" + strProductCond + "') > 0)";
            supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and instr(x.contracttype,'" + strProductCond + "') > 0)";
            curRewardCond = rewardCond;
            // data retrieve
            FillStatementDataSet(pBankCode); //DSstatement =  //3); //3
            getCardProduct(pBankCode);
            getReward(pBankCode);
            pageNo = 0; totalCardPages = 0;
            //curCardNo=String.Empty;
            curAccountNo = String.Empty;
            frmMain.BeginInvoke(frmMain.setMinMaxProgressDelegate, new object[] { DSstatement.Tables["tStatementMasterTable"].Rows.Count });
            foreach (DataRow mRow in DSstatement.Tables["tStatementMasterTable"].Rows)
                {
                frmMain.BeginInvoke(frmMain.setProgressDelegate, new object[] { totRec++ });
                masterRow = mRow;
                //streamWrit.WriteLine(masterRow[mStatementno].ToString());
                pageNo = 1; CurPageRec4Dtl = 0; totNetUsage = 0; CurrentPageFlag = "F 1"; //if page is based on card no
                cardsRows = mRow.GetChildRows(StatementNoDRel); //, DataRowVersion.Proposed
                //cardsRows = DSstatement.Tables["tStatementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[mAccountno]).ToString().Trim() + "'");
                //start new account
                if (masterRow[mHOLSTMT].ToString() == "Y")
                    continue;
                if (prevAccountNo != masterRow[mAccountno].ToString())
                    {
                    if (prevAccountNo == string.Empty)
                        if (curMonth != Convert.ToDateTime(masterRow[mStatementdateto]).Month)//,"dd/MM/yyyy"
                        //if (curMonth != Convert.ToDateTime(masterRow[mStatementdateto]).Month - clsSessionValues.statGenAfterMonth)//,"dd/MM/yyyy"
                            {
                            preExit = false;
                            fileStrm.Close();
                            fileStrmErr.Close();
                            clsBasFile.deleteFile(@strOutputFile);
                            clsBasFile.deleteFile(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName));
                            return "Error in Generation " + pBankName;
                            }
                    //if(!isAccFoterPrnt)
                    //  printAccountFooter();
                    //if (!isAccFoterPrnt && totCards != curCardCounter)
                    //  printAccountFooter();

                    isAccFoterPrnt = false;
                    curMainCard = string.Empty;
                    if (!isHaveF2)//!isHaveF2  CurrentPageFlag != "F 2"
                        {
                        strmWriteErr.WriteLine(prevAccountNo);
                        numOfErr++;
                        //MessageBox.Show( "Error in Genrating Statement", "Account " + prevAccountNo + " Not Have Last Page", MessageBoxButtons.OK ,
                        //  MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1); //,	MessageBoxOptions.RightAlign
                        //break;
                        }
                    isHaveF2 = false;
                    calcAccountRows();
                    prevAccountNo = masterRow[mAccountno].ToString();
                    //pageNo=1; //if page is based on account no
                    isPrimaryOnly = false;
                    }
                curCardCounter++;
                calcCardlRows();
                //        if(totCardRows < 1)continue ;
                /*        if(totAccRows < 1
                          || (masterRow[mCardprimary].ToString() == "Y" 
                          && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0
                          && isPrimaryOnly == true)) */
                //Convert.ToDecimal(
                if ((masterRow[mCardno].ToString() != curMainCard//totAccRows
                  && (totCardRows < 1))
                  || (masterRow[mCardno].ToString() == curMainCard//totAccRows
                  && totAccRows < 1
                  && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0)
                  ) //Convert.ToDecimal(
                //if(totCardRows < 1 && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0) //Convert.ToDecimal(
                    {
                    isHaveF2 = true;
                    //if (totCards == curCardCounter)
                    //  printAccountFooter();

                    continue;
                    }
                /*        if(curTotCardRows < 1
                          && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0) //Convert.ToDecimal(
                          continue;*/

                totNoOfCardStat++;
                printHeader();

                //>        if(masterRow[mCardprimary].ToString() == "Y" && (masterRow[mCardstate].ToString() == "Given" || masterRow[mCardstate].ToString() == "Embossed" || masterRow[mCardstate].ToString() == "New"))
                //>        {
                //>          calcCardlRows();
                //>        }

                foreach (DataRow dRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
                    {
                    detailRow = dRow;
                    if ((detailRow[dPostingdate] == DBNull.Value) && (detailRow[dDocno] == DBNull.Value)) continue;

                    totNetUsage = calculateCrDb(totNetUsage, clsBasValid.validateNum(detailRow[dBilltranamount]), clsBasValid.validateStr(detailRow[dBilltranamountsign]));
                    CurPageRec4Dtl++; curCardRow++;

                    if (curTotCardRows - curCardRow == 0
                      && CurPageRec4Dtl > MaxDetailInLastPage)
                        {
                        streamWrit.WriteLine("");
                        CurPageRec4Dtl = 1;
                        pageNo++;

                        //printAccountFooter();
                        printHeader();
                        }
                    else if (CurPageRec4Dtl > MaxDetailInPage)
                        {
                        CurPageRec4Dtl = 1;
                        pageNo++;

                        //printAccountFooter();
                        printHeader();
                        }
                    //						totNetUsage +=clsBasValid.validateNum(detailRow[dBilltranamount]);
                    printDetail();
                    } // End Detail foreach
                printCardFooter();
                //>if(masterRow[mCardprimary].ToString() == "Y" && (masterRow[mCardstate].ToString() == "Given" || masterRow[mCardstate].ToString() == "Embossed" || masterRow[mCardstate].ToString() == "New"))
                //if (masterRow[mCardno].ToString() == curMainCard || pageNo == totalCardPages)
                if (pageNo == totalCardPages)
                    {
                    printAccountFooter();
                    //pageNo=1; CurPageRec4Dtl=0; //if pages is based on account
                    }
                //streamWrit.WriteLine(strEndOfPage);
                //>        pageNo=1; CurPageRec4Dtl=0; //if pages is based on card
                //>        totNetUsage = 0;
                if (pageNo != totalCardPages)
                    {
                    MessageBox.Show("Error in Genrating Statement", "Please Call The Programmer", MessageBoxButtons.OK,
                      MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1); //,	MessageBoxOptions.RightAlign
                    strmWriteErr.WriteLine(prevAccountNo);
                    numOfErr++;
                    }
                } // End Master foreach
            //if (!isAccFoterPrnt)
            //  printAccountFooter();

            //fillStatementHistory(pStmntType, pAppendData);
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
                // Close output ile
                streamWrit.Flush();
                streamWrit.Close();
                fileStrm.Close();

                printStatementSummary();

                strmWriteErr.Flush();
                strmWriteErr.Close();
                fileStrmErr.Close();

                // Close Summary File
                streamSummary.Flush();
                streamSummary.Close();
                fileSummary.Close();

                ArrayList aryLstFiles = new ArrayList();
                if (numOfErr == 0)
                    clsBasFile.deleteFile(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName));
                else
                    aryLstFiles.Add(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err.txt");
                //aryLstFiles.Add("");
                aryLstFiles.Add(@strOutputFile);
                aryLstFiles.Add(@fileSummaryName);
                clsBasFile.generateFileMD5(aryLstFiles, @clsBasFile.getPathWithoutExtn(pStrFileName) + ".MD5");
                aryLstFiles.Add(@clsBasFile.getPathWithoutExtn(pStrFileName) + ".MD5");
                SharpZip zip = new SharpZip();
                zip.createZip(aryLstFiles, @clsBasFile.getPathWithoutExtn(pStrFileName) + ".zip", "");
                }

            }
        return rtrnStr;
        }


    protected void printHeader()
        {
        string newaddress1, newaddress2;
        clsMaintainData.fixAddress(curBranchVal, masterRow[mCustomeraddress1].ToString(), out newaddress1, out newaddress2);
        //    if(masterRow[mCardprimary].ToString() == "Y" && (masterRow[mCardstate].ToString() == "Given" || masterRow[mCardstate].ToString() == "Embossed" || masterRow[mCardstate].ToString() == "New"))
        //if(masterRow[mCardno].ToString() == curMainCard)
        if (pageNo == totalCardPages)
            {
            CurrentPageFlag = "F 2";
            isHaveF2 = true;
            }

        extAccNum = clsBasValid.validateStr(masterRow[mExternalno]);
        if (extAccNum.Trim() == "")
            extAccNum = clsBasValid.validateStr(masterRow[mAccountno]);

        streamWrit.WriteLine(strEndOfPage);
        streamWrit.WriteLine(basText.alignmentMiddle(masterRow[mCardbranchpartname], 35) + basText.replicat(" ", 20) + basText.replicat(" ", 10) + basText.replicat(" ", 23) + CurrentPageFlag);
        //streamWrit.WriteLine(String.Empty);
        //streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(basText.replicat(" ", 63) + basText.alignmentMiddle(masterRow[mCardclientname], 50));//masterRow[mCustomername],50
        //streamWrit.WriteLine(basText.alignmentLeft(masterRow[mCustomername], 50) + basText.replicat(" ", 37) + "{0,5:MM/yy}", masterRow[mStatementdatefrom]);
        streamWrit.WriteLine(basText.alignmentLeft(masterRow[mCustomername], 50) + basText.replicat(" ", 37) + "{0,5:MM/yy}", masterRow[mStatementdateto]);
        streamWrit.WriteLine(String.Empty);
        //streamWrit.WriteLine(basText.alignmentMiddle(ValidateArbic(masterRow[mCustomeraddress1].ToString()), 50) + basText.replicat(" ", 31) + basText.formatNum(masterRow[mCardno], "####-####-####-####"));
        //streamWrit.WriteLine(basText.alignmentMiddle(ValidateArbic(newaddress1), 50) + basText.replicat(" ", 31) + basText.formatNum(masterRow[mCardno], "####-####-####-####"));
        //AWB-504
        streamWrit.WriteLine(basText.alignmentMiddle(ValidateArbic(newaddress1), 50) + basText.replicat(" ", 31) + basText.formatCardNumber(basText.formatNum(masterRow[mCardno], "####-####-####-####"),'-'));
        //streamWrit.WriteLine(basText.alignmentMiddle(ValidateArbic(masterRow[mCustomeraddress2].ToString()), 50) + basText.replicat(" ", 31) + basText.formatNum(masterRow[mAccountlim], "########"));//basText.formatNum(masterRow[mCardlimit],"########")
        streamWrit.WriteLine(basText.alignmentMiddle(ValidateArbic(newaddress2), 50) + basText.replicat(" ", 31) + basText.formatNum(masterRow[mAccountlim], "########"));//basText.formatNum(masterRow[mCardlimit],"########")
        streamWrit.WriteLine(basText.alignmentMiddle(ValidateArbic(masterRow[mCustomeraddress3].ToString().Trim()) + " " + masterRow[mCustomerregion].ToString().Trim() + " " + masterRow[mCustomercity].ToString().Trim(), 50) + basText.replicat(" ", 31) + extAccNum);  //clsBasValid.validateStr(masterRow[mAccountno])  //basText.formatNum(masterRow[mAccountno],"##-##-######-####")
        streamWrit.WriteLine(basText.replicat(" ", 81) + pageNo + "/" + totalCardPages);//+" of " + totalAccPages totalCardPages
        streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine(String.Empty);
        totalPages++;
        totNoOfPageStat++;
        }


    protected void printDetail()
        {
        DateTime trnsDate = clsBasValid.validateDate(detailRow[dTransdate]), postingDate = clsBasValid.validateDate(detailRow[dPostingdate]);
        //    if(trnsDate.Month == 9)trnsDate = postingDate;
        if (trnsDate > postingDate)
            trnsDate = postingDate;


        if (masterRow[mAccountcurrency].ToString().Trim() != detailRow[dOrigtrancurrency].ToString().Trim())
            strForeignCurr = basText.formatNum(detailRow[dOrigtranamount], "#,##0.00;(#,##0.00)") + " " + basText.Replace(clsBasValid.validateStr(detailRow[dOrigtrancurrency]), "XXX", "   ");
        else
            strForeignCurr = basText.replicat(" ", 16);

        if (strForeignCurr.Trim() == "0")
            strForeignCurr = basText.replicat(" ", 16);

        streamWrit.WriteLine("    {0:dd/MM}  {1:dd/MM} {2,13}  {3,-40} {4,16} {5,16} {6,2}", trnsDate, postingDate, basText.trimStr(detailRow[dRefereneno], 12), basText.trimStr(detailRow[dTrandescription].ToString().Trim() + " " + detailRow[dMerchant].ToString().Trim(), 45), strForeignCurr, basText.formatNum(detailRow[dBilltranamount], "#,##0.00;(#,##0.00)"), CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign]))); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate])
        totNoOfTransactions++;
        }

    protected void printCardFooter()
        {
        streamWrit.WriteLine("");//basText.replicat(" ",96)+basText.replicat("_",16)
        streamWrit.WriteLine("");//basText.replicat(" ",60)+ "Net Usage     " + "ÕÇÝì ÇáÇÓÊÎÏÇãÇÊ" + "      " + basText.formatNum(Math.Abs(totNetUsage),"#,##0.00",16) + " " + CrDb(totNetUsage)
        streamWrit.WriteLine(basText.alignmentMiddle("Gem Rewards Opening Balance", 30) + basText.alignmentMiddle("Gem Rewards Earned*", 30) + basText.alignmentMiddle("Gem Rewards Redeemed", 30) + basText.alignmentMiddle("Gem Rewards Closing Balance", 30));
        rewardRows = DSreward.Tables["Reward"].Select("CLIENTID = '" + clsBasValid.validateStr(masterRow[mClientid]).ToString().Trim() + "'");
        if (rewardRows.Length > 0)
            {
            rewardRow = rewardRows[0];
            streamWrit.WriteLine("** Gem " + basText.formatNum(rewardRow[mOpeningbalance], "#0.00;(#0.00)", 30, "M") + basText.formatNum(int.Parse(rewardRow[mEarnedBonus].ToString()) - int.Parse(rewardRow[mBonusAdjustment].ToString()), "#0.00;(#0.00)", 30, "M") + basText.formatNum(rewardRow[mRedeemedBonus], "#0.00;(#0.00)", 30, "M") + basText.formatNum(rewardRow[mClosingbalance], "#0.00;(#0.00)", 30, "M"));
            }
        else
            {
            streamWrit.WriteLine(basText.replicats(" ", 80));
            }
        streamWrit.WriteLine(String.Empty);
        streamWrit.WriteLine("* Gem Rewards are only earned on POS and Internet purchases, and not on cash withdrawals. Refunded transactions are not eligible for Rewards.");
        streamWrit.WriteLine(String.Empty);
        }

    protected void printAccountFooter()
        {
        streamWrit.WriteLine("GRAND" + basText.formatNum(masterRow[mOpeningbalance], "#,##0.00;(#,##0.00)", 16) + basText.formatNum(masterRow[mTotaldebits], "#,##0.00;(#,##0.00)", 16) + basText.formatNum(masterRow[mTotalcredits], "#,##0.00;(#,##0.00)", 16) + basText.formatNum(masterRow[mClosingbalance], "#,##0.00;(#,##0.00)", 16) + basText.formatNum(masterRow[mMindueamount], "#,##0.00;(#,##0.00)", 16) + basText.replicat(" ", 12) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy") + "\n");//+strEndOfPage
        isAccFoterPrnt = true;
        }


    private void calcCardlRows()
        {
        totalCardPages = 0;
        totCardRows = curTotCardRows = curCardRow = 0;
        foreach (DataRow dtRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
            {
            if ((dtRow[dPostingdate] == DBNull.Value) && (dtRow[dDocno] == DBNull.Value)) continue;
            //streamWrit.WriteLine(basText.trimStr(dtRow[dTrandescription],40)); 
            totCardRows++;
            }
        curTotCardRows = totCardRows;
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
        accountRows = DSstatement.Tables["tStatementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[mAccountno]) + "'");
        totalAccPages = 0;
        totAccRows = 0;
        string prevCardNo = String.Empty, CurCardNo = String.Empty;
        int currAccRowsPages = 0;
        foreach (DataRow dtAccRow in accountRows) //mRow.GetChildRows(StatementNoDRel)
            {
            if ((dtAccRow[dPostingdate] == DBNull.Value) && (dtAccRow[dDocno] == DBNull.Value)) continue;
            //streamWrit.WriteLine(basText.trimStr(dtAccRow[dTrandescription],40)); 
            currAccRowsPages++;
            totAccRows++;

            CurCardNo = dtAccRow[dCardno].ToString();
            if (prevCardNo != CurCardNo && prevCardNo != String.Empty)
                {
                totalAccPages++;
                currAccRowsPages = 1;
                }
            if (currAccRowsPages == MaxDetailInPage)
                {
                currAccRowsPages = 0;
                totalAccPages++;
                }
            prevCardNo = dtAccRow[dCardno].ToString();
            }
        if (currAccRowsPages > 0)
            totalAccPages++;
        if (totalAccPages < 1)
            totalAccPages = 1;

        if (totAccRows == 0 && (masterRow[mCardprimary].ToString() == "Y"
          && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) != 0))
            {
            isPrimaryOnly = true;
            }

        //    accountRows = DSstatement.Tables["tStatementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[mAccountno])+ "'");
        mainRows = DSstatement.Tables["tStatementMasterTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[mAccountno]) + "'");
        curMainCard = CurCardNo = "";
        totCards = curCardCounter = 0;
        foreach (DataRow mainRow in mainRows) //mRow.GetChildRows(StatementNoDRel)
            {
            CurCardNo = mainRow[mCardno].ToString();
            totCards++;
            //if (calcCardRows(mainRow[mCardstate].ToString()) > 0)
            if (calcCardRows(mainRow[mCardno].ToString()) > 0)
                curMainCard = CurCardNo; //mainRow[mCardno].ToString();
            //if (mainRow[mCardprimary].ToString() == "Y" && calcCardRows(mainRow[mCardstate].ToString()) > 0)
            if (mainRow[mCardprimary].ToString() == "Y" && calcCardRows(mainRow[mCardno].ToString()) > 0)
                curMainCard = CurCardNo; //mainRow[mCardno].ToString();
            if ((mainRow[mCardstate].ToString() == "Given" || mainRow[mCardstate].ToString() == "Embossed" || mainRow[mCardstate].ToString() == "New" || mainRow[mCardstate].ToString() == "New Pin Generated Only") && calcCardRows(mainRow[mCardstate].ToString()) > 0)
                {
                curMainCard = CurCardNo; //mainRow[mCardno].ToString();
                break;
                }
            }

        if (curMainCard == "")
            curMainCard = CurCardNo;

        }



    private int calcCardRows(string pCardNo)
        {
        accountRows = DSstatement.Tables["tStatementDetailTable"].Select("cardno = '" + pCardNo + "'");
        int totcardTrns = 0;
        foreach (DataRow dtAccRow in accountRows) //mRow.GetChildRows(StatementNoDRel)
            {
            totcardTrns++;
            }
        return totcardTrns;
        }

    private void printStatementSummary()
        {
        streamSummary.WriteLine(strBankName + " Visa Statement");
        streamSummary.WriteLine("__________________________");
        streamSummary.WriteLine("");
        streamSummary.WriteLine("No of Statements   " + totNoOfCardStat.ToString());
        streamSummary.WriteLine("No of Pages        " + totNoOfPageStat.ToString());
        streamSummary.WriteLine("No of Transactions " + totNoOfTransactions.ToString());

        clsStatementSummary StatSummary = new clsStatementSummary();
        StatSummary.BankCode = curBranchVal;
        StatSummary.BankName = strBankName;
        StatSummary.StatementDate = vCurDate;
        StatSummary.CreationDate = DateTime.Now;
        DataRow[] productRows;
        prevAccountNo = String.Empty;
        totNoOfCardStat = 0;
        totNoOfTransactionsInt = 0;
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
                //streamWrit.WriteLine(masterRow[mStatementno].ToString());
                pageNo = 1; CurPageRec4Dtl = 0; totNetUsage = 0; CurrentPageFlag = "F 1"; //if page is based on card no
                cardsRows = mRow.GetChildRows(StatementNoDRel); //, DataRowVersion.Proposed
                //cardsRows = DSstatement.Tables["tStatementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[mAccountno]).ToString().Trim() + "'");
                //start new account
                if (prevAccountNo != masterRow[mAccountno].ToString())
                    {
                    isAccFoterPrnt = false;
                    curMainCard = string.Empty;
                    calcAccountRows();
                    prevAccountNo = masterRow[mAccountno].ToString();
                    isPrimaryOnly = false;
                    }
                curCardCounter++;
                calcCardlRows();
                if ((masterRow[mCardno].ToString() != curMainCard//totAccRows
                  && (totCardRows < 1))
                  || (masterRow[mCardno].ToString() == curMainCard//totAccRows
                  && totAccRows < 1
                  && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0)
                  ) //Convert.ToDecimal(
                    {
                    continue;
                    }
                totNoOfCardStat++;
                foreach (DataRow dRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
                    {
                    detailRow = dRow;
                    if ((detailRow[dPostingdate] == DBNull.Value) && (detailRow[dDocno] == DBNull.Value)) continue;

                    totNetUsage = calculateCrDb(totNetUsage, clsBasValid.validateNum(detailRow[dBilltranamount]), clsBasValid.validateStr(detailRow[dBilltranamountsign]));
                    CurPageRec4Dtl++; curCardRow++;

                    if (curTotCardRows - curCardRow == 0
                      && CurPageRec4Dtl > MaxDetailInLastPage)
                        {
                        CurPageRec4Dtl = 1;
                        pageNo++;
                        }
                    else if (CurPageRec4Dtl > MaxDetailInPage)
                        {
                        CurPageRec4Dtl = 1;
                        pageNo++;
                        }
                    hasInterset = CalcTransInterest(detailRow[dTrandescription].ToString());
                    if (hasInterset)
                        totNoOfTransactionsInt++;
                    } //end of detail foreach

                } // End Master foreach
            if (totNoOfCardStat != 0)
                {
                StatSummary.NoOfStatements = totNoOfCardStat;
                StatSummary.NoOfTransactionsInt = totNoOfTransactionsInt;
                StatSummary.ProductCode = int.Parse(ProductRow["code"].ToString());
                StatSummary.ProductName = ProductRow["name"].ToString();
                StatSummary.InsertRecordDb(vCurDate.ToString("yyyyMM") + strBankName);
                totNoOfCardStat = 0;
                totNoOfTransactionsInt = 0;
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
        //    streamWritMD5.WriteLine(clsBasFile.getFileFromPath(fileSummaryName) + "  >>  " + clsBasFile.getFileMD5(fileSummaryName));
        streamWritMD5.Flush();
        streamWritMD5.Close();
        fileStrmMd5.Close();
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

    ~clsStatementAwbPlus()
        {
        DSstatement.Dispose();
        }
    }
