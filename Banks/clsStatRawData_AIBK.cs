using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Collections;


// Branch X
public class clsStatRawData_AIBK : clsBasStatement
    {
    protected string strBankName;
    protected FileStream fileStrmBasic, fileStrmTrans, fileSummary;//, fileStrmSubtotal
    protected string strFileBasic, strFileTrans;//, strFileSubtotal
    protected StreamWriter streamWritBasic, streamWritTrans, streamSummary;//, streamWritSubtotal
    protected DataRow masterRow;
    protected DataRow detailRow;
    protected const int MaxDetailInPage = 20; //
    protected const int linesInLastPage = 67; //
    protected int CurPageRec4Dtl = 0;
    protected int pageNo = 0, totalCardPages = 0 //, totalPages=0
      , totalAccPages = 0, totCardRows = 0, totAccRows = 0, totAccCards = 0;
    protected string curCardNo, CardNumber = String.Empty, curCardNumber = String.Empty, PrevCardNumber = String.Empty;// 
    protected string curAccountNo, prevAccountNo = String.Empty;//
    protected string strAccountFooter;
    protected int intAccountFooter;
    protected decimal totNetUsage = 0;
    protected decimal totAccountValue = 0;
    protected DataRow[] cardsRows, accountRows;
    protected string CrDbDetail;
    protected string fldSprtr = "|";//#
    protected bool isPrimaryOnly;
    protected string curMainCard;
    protected int curAccRows = 0;
    protected int totCrdNoInAcc, curCrdNoInAcc;
    protected string stmNo;
    protected DataRow[] mainRows;

    protected string extAccNum;
    protected string strOutputPath, strOutputFile, fileSummaryName;
    protected DateTime vCurDate;
    protected int totRec = 1;
    protected frmStatementFile frmMain;
    protected DataSet DSstatementRaw;
    protected int curMonth;
    protected bool preExit = true;
    protected string stmntType;
    protected int totNoOfCardStat, totNoOfPageStat, totNoOfTransactions = 0, totNoOfTransactionsInt = 0;
    private DataRow ProductRow;
    protected bool hasInterset = false;
    private string strProductCond = string.Empty;

    public clsStatRawData_AIBK()
        {
        }

    public string Statement(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pDate, string pStmntType, bool pAppendData)
        {
        string rtrnStr = "Successfully Generate " + pBankName;
        curMonth = pDate.Month;
        clsMaintainData maintainData = new clsMaintainData();
        maintainData.matchCardBranch4Account(pBankCode);
        stmntType = pStmntType;

        pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);
        vCurDate = pDate; //DateTime.Now.AddMonths(-1);
        strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
        clsBasFile.createDirectory(pStrFileName + vCurDate.ToString("yyyyMM") + pBankName);
        pStrFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM") + ".txt";
        curBranchVal = pBankCode; // 4; //4  = real   1 = test
        //MainTableCond = " m.contracttype in " + strProductCond + "";//strWhereCond
        //supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + ")";
        FillStatementDataSet(pBankCode, ""); //DSstatement =  //6); // 6
        getCardProduct(pBankCode);
        Statement(pStrFileName, DSstatement);
        return "";
        }

    public string Statement(string pStrFileName, DataSet pDSstatement)
        {
        try
            {
            DSstatementRaw = pDSstatement;
            // open output file
            strFileBasic = clsBasFile.getPathWithoutFile(pStrFileName) + "STM_Header_" + vCurDate.ToString("yyMM") + ".txt";
            clsBasFile.deleteFile(strFileBasic);
            fileStrmBasic = new FileStream(strFileBasic, FileMode.Create);
            streamWritBasic = new StreamWriter(fileStrmBasic);
            streamWritBasic.AutoFlush = true;
            // open output file
            strFileTrans = clsBasFile.getPathWithoutFile(pStrFileName) + "STM_Detail_" + vCurDate.ToString("yyMM") + ".txt";
            clsBasFile.deleteFile(strFileTrans);
            fileStrmTrans = new FileStream(strFileTrans, FileMode.Create);
            streamWritTrans = new StreamWriter(fileStrmTrans);
            streamWritTrans.AutoFlush = true;

            fileSummaryName = pStrFileName;
            fileSummaryName = clsBasFile.getPathWithoutFile(fileSummaryName) + "STM_Summary_" + vCurDate.ToString("yyMM") + ".txt";

            // open Summary file
            fileSummary = new FileStream(fileSummaryName, FileMode.Create); //Create
            streamSummary = new StreamWriter(fileSummary, Encoding.Default);
            streamSummary.AutoFlush = true;

            pageNo = 0; totalCardPages = 0;
            curCardNo = String.Empty;
            curAccountNo = String.Empty;

            frmMain.BeginInvoke(frmMain.setMinMaxProgressDelegate, new object[] { DSstatement.Tables["tStatementMasterTable"].Rows.Count });
            foreach (DataRow mRow in DSstatementRaw.Tables["tStatementMasterTable"].Rows)
                {
                frmMain.BeginInvoke(frmMain.setProgressDelegate, new object[] { totRec++ });
                masterRow = mRow;
                pageNo = 1;  //if page is based on card no
                CurPageRec4Dtl = 0;
                cardsRows = mRow.GetChildRows("StaementNoDR"); //StatementNoDRel, DataRowVersion.Proposed
                if (masterRow[mHOLSTMT].ToString() == "Y")
                    continue;
                //start new account
                if (prevAccountNo != masterRow[mAccountno].ToString())
                    {
                    if (prevAccountNo == string.Empty)
                        if (curMonth != Convert.ToDateTime(masterRow[mStatementdateto]).Month)//,"dd/MM/yyyy"
                            {
                            preExit = false;
                            fileStrmBasic.Close();
                            fileStrmTrans.Close();
                            clsBasFile.deleteFile(strFileBasic);
                            clsBasFile.deleteFile(strFileTrans);
                            return "Error in Generation ";//+ pBankName
                            }
                    strAccountFooter = String.Empty;
                    intAccountFooter = 0;
                    totAccountValue = 0;
                    calcAccountRows();
                    prevAccountNo = masterRow[mAccountno].ToString();
                    pageNo = 1;  //if page is based on account no 
                    if (totAccRows > 0 || Convert.ToDecimal(masterRow[mClosingbalance].ToString()) != 0)
                        printHeader();//>>
                    isPrimaryOnly = false;
                    }
                if (totAccRows < 1 && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0) //             || (masterRow[mCardno].ToString() == curMainCard   // Convert.ToDecimal(
                    continue;

                curCardNumber = masterRow[mCardno].ToString();

                foreach (DataRow dRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
                    {
                    detailRow = dRow;
                    if ((detailRow[dPostingdate] == DBNull.Value) && (detailRow[dDocno] == DBNull.Value))
                        continue;// Exclude On-Hold Transactions 
                    CrDbDetail = String.Empty;
                    if (CurPageRec4Dtl >= MaxDetailInPage)
                        {
                        CurPageRec4Dtl = 0;
                        pageNo++;
                        }
                    if (clsBasValid.validateStr(detailRow[dBilltranamountsign]) == "CR")
                        {
                        CrDbDetail = "CR";
                        totNetUsage += clsBasValid.validateNum(detailRow[dBilltranamount]);
                        }
                    else
                        {
                        CrDbDetail = String.Empty;
                        totNetUsage -= clsBasValid.validateNum(detailRow[dBilltranamount]);
                        }

                    CurPageRec4Dtl++;
                    printDetail();
                    }

                totAccountValue += totNetUsage;

                if (masterRow[mCardprimary].ToString() == "Y" && (masterRow[mCardstate].ToString() == "Given" || masterRow[mCardstate].ToString() == "Embossed" || masterRow[mCardstate].ToString() == "New" || masterRow[mCardstate].ToString() == "New Pin Generated Only"))
                    {
                    pageNo = 1; CurPageRec4Dtl = 0; //if pages is based on account
                    }
                totNetUsage = 0;
                PrevCardNumber = curCardNumber; //masterRow[mCardno].ToString();
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
            if (preExit)
                {

                // Close output ile
                streamWritBasic.Flush();
                streamWritBasic.Close();
                fileStrmBasic.Close();
                streamWritTrans.Flush();
                streamWritTrans.Close();
                fileStrmTrans.Close();

                printStatementSummary();

                // Close Summary File
                streamSummary.Flush();
                streamSummary.Close();
                fileSummary.Close();

                ArrayList aryLstFiles = new ArrayList();
                aryLstFiles.Add(@strFileBasic);
                aryLstFiles.Add(@strFileTrans);
                SharpZip zip = new SharpZip();
                zip.createZip(aryLstFiles, @clsBasFile.getPathWithoutExtn(pStrFileName) + ".zip", "");//+ "_Raw"
                }
            }
        return "";

        }



    protected virtual void printHeader()
        {
        string newaddress1, newaddress2;
        clsMaintainData.fixAddress(curBranchVal, masterRow[mCustomeraddress1].ToString(), out newaddress1, out newaddress2);
        extAccNum = clsBasValid.validateStr(masterRow[mExternalno]);
        if (extAccNum.Trim() == "")
            extAccNum = clsBasValid.validateStr(masterRow[mAccountno]);

        pageNo = 1;
        streamWritBasic.WriteLine(basText.alignmentLeft(curMainCard, 19)
        + basText.alignmentLeft(masterRow[mCustomername].ToString(), 100)
        + basText.formatNum(0, "#,##0.00", 16)
        + basText.alignmentLeft(ValidateArbic(newaddress1), 50)
        + basText.alignmentLeft(ValidateArbic(newaddress2), 50)
        + basText.alignmentLeft(masterRow[mCustomercity].ToString(), 30)
        + basText.alignmentLeft(masterRow[mCustomerzipcode].ToString(), 10)
        + basText.formatNum(masterRow[mClosingbalance], "#,##0.00", 16)
        + basText.formatNum(masterRow[mOpeningbalance], "#,##0.00", 16)
        + basText.formatNum(masterRow[mTotaldebits], "#,##0.00", 16)
        + basText.formatNum(masterRow[mTotalcredits], "#,##0.00", 16)
        + basText.formatNum(masterRow[mAccountlim], "#,##0.00", 16)
        + basText.formatNum(masterRow[mAccountavailablelim], "#,##0.00", 16)
        + basText.formatNum(masterRow[mTotalDueAmount], "#,##0.00", 16)
        + basText.formatNum(masterRow[mTotaloverdueamount], "#,##0.00", 16)
        + basText.formatAS400Date(masterRow[mStetementduedate], 8, "R")
        + basText.formatAS400Date(masterRow[mStatementdateto], 8, "R")
        + basText.formatNum(masterRow[mTotaloverdueamount], "#,##0.00", 16)
        + basText.formatNum(masterRow[mTotaloverdueamount], "#,##0.00", 16)
        + basText.formatNum(masterRow[mTotaloverdueamount], "#,##0.00", 16));


      
        totNoOfCardStat++;
        }


    protected virtual void printDetail()
        {
        DateTime trnsDate = clsBasValid.validateDate(detailRow[dTransdate]), postingDate = clsBasValid.validateDate(detailRow[dPostingdate]);
        if (trnsDate > postingDate)
            trnsDate = postingDate;
        string trnsDesc, strForeignCurr;
        if (masterRow[mAccountcurrency].ToString().Trim() != detailRow[dOrigtrancurrency].ToString().Trim())
            strForeignCurr = basText.formatNumUnSign(detailRow[dOrigtranamount], "##,###,##0.00;(#,###,##0.00)", 13) + " " + basText.Replace(clsBasValid.validateStr(detailRow[dOrigtrancurrency]), "XXX", "   ");
        else
            strForeignCurr = basText.replicat(" ", 16);
        if (detailRow[dMerchant].ToString().Trim() == "")
            trnsDesc = detailRow[dTrandescription].ToString().Trim();
        else
            trnsDesc = detailRow[dMerchant].ToString().Trim();
        if (detailRow[dInstallmentData].ToString().Trim() != "")
            {
            if (trnsDesc == "Installment repayment")
                {
                string input = detailRow[dInstallmentData].ToString().Trim();
                int index = detailRow[dInstallmentData].ToString().Trim().IndexOf(":");
                if (index > 0)
                    input = input.Substring(0, index);
                streamWritTrans.WriteLine(curMainCard + fldSprtr + trnsDate.ToString("dd/MM") + fldSprtr + postingDate.ToString("dd/MM") + fldSprtr + basText.trimStr(detailRow[dRefereneno], 25) + fldSprtr + basText.trimStr(basText.trimStr(input.Substring(0, input.IndexOf(',')) + input.Substring(basText.GetNthIndex(input, ',', 2)).Substring(input.Substring(basText.GetNthIndex(input, ',', 2)).IndexOf('#')).Replace('#', ' ') + " " + basText.trimStr(GetMerchantByOrigDocNo(long.Parse(detailRow[dOrigDocNo].ToString())), 40), 68), 40) + fldSprtr + strForeignCurr + fldSprtr + basText.formatNum(detailRow[dBilltranamount], "#,##0.00;(#,##0.00)") + CrDbDetail + fldSprtr);//detailRow[dBilltranamountsign]//clsBasValid.validateDate(detailRow[dTransdate])//clsBasValid.validateDate(detailRow[dPostingdate]) //masterRow[mCardno]
                }
            else
                streamWritTrans.WriteLine(curMainCard + fldSprtr + trnsDate.ToString("dd/MM") + fldSprtr + postingDate.ToString("dd/MM") + fldSprtr + basText.trimStr(detailRow[dRefereneno], 25) + fldSprtr + basText.trimStr(trnsDesc.Trim(), 40) + fldSprtr + strForeignCurr + fldSprtr + basText.formatNum(detailRow[dBilltranamount], "#,##0.00;(#,##0.00)") + CrDbDetail + fldSprtr);//detailRow[dBilltranamountsign]//clsBasValid.validateDate(detailRow[dTransdate])//clsBasValid.validateDate(detailRow[dPostingdate]) //masterRow[mCardno]
            }
        else
            streamWritTrans.WriteLine(curMainCard + fldSprtr + trnsDate.ToString("dd/MM") + fldSprtr + postingDate.ToString("dd/MM") + fldSprtr + basText.trimStr(detailRow[dRefereneno], 25) + fldSprtr + basText.trimStr(trnsDesc.Trim(), 40) + fldSprtr + strForeignCurr + fldSprtr + basText.formatNum(detailRow[dBilltranamount], "#,##0.00;(#,##0.00)") + CrDbDetail + fldSprtr);//detailRow[dBilltranamountsign]//clsBasValid.validateDate(detailRow[dTransdate])//clsBasValid.validateDate(detailRow[dPostingdate]) //masterRow[mCardno]
        totNoOfTransactions++;
        }
    
    protected void calcAccountRows()
        {
        curAccRows = 0;
        accountRows = null;
        stmNo = masterRow[mStatementno].ToString();
        stmNo = masterRow[mAccountno].ToString();

        accountRows = DSstatementRaw.Tables["tStatementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[mAccountno]).ToString().Trim() + "'");
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
        mainRows = DSstatementRaw.Tables["tStatementMasterTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[mAccountno]) + "'");
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


    protected void printFileMD5()
        {
        FileStream fileStrmMd5;
        StreamWriter streamWritMD5;
        fileStrmMd5 = new FileStream(strOutputPath + "\\" + strBankName + vCurDate.ToString("yyyyMM") + ".MD5", FileMode.Create);
        streamWritMD5 = new StreamWriter(fileStrmMd5);
        streamWritMD5.WriteLine(clsBasFile.getFileFromPath(strFileBasic) + "  >>  " + clsBasFile.getFileMD5(strFileBasic));
        streamWritMD5.WriteLine(clsBasFile.getFileFromPath(strFileTrans) + "  >>  " + clsBasFile.getFileMD5(strFileTrans));
        streamWritMD5.Flush();
        streamWritMD5.Close();
        fileStrmMd5.Close();
        }

    protected void printStatementSummary()
        {
        streamSummary.WriteLine(strBankName + " Statement");
        streamSummary.WriteLine("__________________________");
        streamSummary.WriteLine("");
        streamSummary.WriteLine("No of Statements   " + totNoOfCardStat.ToString());
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
                pageNo = 1;  //if page is based on card no
                CurPageRec4Dtl = 0;
                cardsRows = mRow.GetChildRows("StaementNoDR"); //StatementNoDRel, DataRowVersion.Proposed
                //start new account
                if (prevAccountNo != masterRow[mAccountno].ToString())
                    {
                    calcAccountRows();
                    prevAccountNo = masterRow[mAccountno].ToString();
                    pageNo = 1;  //if page is based on account no 
                    isPrimaryOnly = false;
                    if (totAccRows > 0 || Convert.ToDecimal(masterRow[mClosingbalance].ToString()) != 0)
                        totNoOfCardStat++;
                    }
                if (totAccRows < 1 && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0) //             || (masterRow[mCardno].ToString() == curMainCard   // Convert.ToDecimal(
                    continue;
                curCardNumber = masterRow[mCardno].ToString();
                foreach (DataRow dRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
                    {
                    detailRow = dRow;
                    if ((detailRow[dPostingdate] == DBNull.Value) && (detailRow[dDocno] == DBNull.Value))
                        continue;// Exclude On-Hold Transactions 
                    if (CurPageRec4Dtl >= MaxDetailInPage)
                        {
                        CurPageRec4Dtl = 0;
                        pageNo++;
                        }
                    CurPageRec4Dtl++;
                    if (CalcTransInterest(detailRow[dTrandescription].ToString()))
                        totNoOfTransactionsInt++;
                    }
                if (masterRow[mCardprimary].ToString() == "Y" && (masterRow[mCardstate].ToString() == "Given" || masterRow[mCardstate].ToString() == "Embossed" || masterRow[mCardstate].ToString() == "New" || masterRow[mCardstate].ToString() == "New Pin Generated Only"))
                    {
                    pageNo = 1; CurPageRec4Dtl = 0; //if pages is based on account
                    }
                totNetUsage = 0;
                PrevCardNumber = curCardNumber; //masterRow[mCardno].ToString();
                }

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

    private string GetMerchantByOrigDocNo(long pDocno)
        {
        mainRows = null;
        string result = string.Empty;
        mainRows = DSstatement.Tables["tStatementDetailTable"].Select("ORIGDOCNO = " + pDocno + "");
        foreach (DataRow mainRow in mainRows)
            {
            result = mainRow[dInstallmentMerchantLocation].ToString();
            }
        return result;
        }
    protected void makeZip()
        {
        }

    public string bankName
        {
        get { return strBankName; }
        set { strBankName = value; }
        }// bankName

    public string fieldSeparatorStr
        {
        set { fldSprtr = value; }
        }

    public string productCond
        {
        get { return strProductCond; }
        set { strProductCond = value; }
        }  // productCond

    public frmStatementFile setFrm
        {
        set { frmMain = value; }
        }// setFrm

    ~clsStatRawData_AIBK()
        {
        DSstatementRaw.Dispose();
        }
    }
