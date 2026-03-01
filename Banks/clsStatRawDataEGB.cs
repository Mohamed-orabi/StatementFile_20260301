using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Collections;


// Branch X
public class clsStatRawDataEGB : clsBasStatement
    {
    protected string strBankName;
    protected FileStream fileStrmBasic, fileStrmTrans, fileSummary, fileStrmDelivery;
    protected string strFileBasic, strFileTrans, strFileDelivery;
    protected StreamWriter streamWritBasic, streamWritTrans, streamSummary, streamWritDelivery;
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
    protected string fldSprtr = ",";//#
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
    protected string CustomerPhone;
    protected DataRow[] emailRows = null;
    protected string CitizenID;
    protected DataRow[] PassportNoRows = null;

    private int installdocno = 0;
    private int originstalldocno = 0;
    private string installdata = string.Empty;

    public clsStatRawDataEGB()
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
        FillStatementDataSet_SortCardPriority(pBankCode, ""); //DSstatement =  //6); // 6
        getCardProduct(pBankCode);
        getInstallment(pBankCode);
        getClientEmail(pBankCode);
        getClientPassportNo(pBankCode);
        Statement(pStrFileName, DSstatement);
        return "";
        }

    public string Statement(string pStrFileName, DataSet pDSstatement)
        {
        try
            {
            DSstatementRaw = pDSstatement;

            // open output file
            strFileBasic = clsBasFile.getPathWithoutFile(pStrFileName) + "STMT_HDR.DAT";
            clsBasFile.deleteFile(strFileBasic);
            fileStrmBasic = new FileStream(strFileBasic, FileMode.Create);
            streamWritBasic = new StreamWriter(fileStrmBasic);
            streamWritBasic.AutoFlush = true;

            // open output file
            strFileTrans = clsBasFile.getPathWithoutFile(pStrFileName) + "STMT_DTL.DAT";
            clsBasFile.deleteFile(strFileTrans);
            fileStrmTrans = new FileStream(strFileTrans, FileMode.Create);
            streamWritTrans = new StreamWriter(fileStrmTrans);
            streamWritTrans.AutoFlush = true;

            // open output file
            strFileDelivery = clsBasFile.getPathWithoutFile(pStrFileName) + "STMT_Delivery.CSV";
            clsBasFile.deleteFile(strFileDelivery);
            fileStrmDelivery = new FileStream(strFileDelivery, FileMode.Create);
            streamWritDelivery = new StreamWriter(fileStrmDelivery);
            streamWritDelivery.AutoFlush = true;
            streamWritDelivery.WriteLine("STATEMENT NUMBER,Ordering Branch,CUSTOMER NAME,Cardholder Address Line #1,Cardholder Address Line #2,PHONE NUMBER,CITIZEN ID");

            fileSummaryName = pStrFileName;
            fileSummaryName = clsBasFile.getPathWithoutFile(fileSummaryName) + "STMT_SMR.txt";

            // open Summary file
            fileSummary = new FileStream(fileSummaryName, FileMode.Create); //Create
            streamSummary = new StreamWriter(fileSummary, Encoding.Default);
            streamSummary.AutoFlush = true;

            pageNo = 0; totalCardPages = 0;
            curCardNo = String.Empty;
            curAccountNo = String.Empty;

            //// Mahmoud AMR to remove duplicates

            //DSstatementRaw.Tables["tStatementMasterTable"].DefaultView.Sort = "STATEMENTNUMBER ASC";
            //DataTable stmtMasterRemoveDup = DSstatementRaw.Tables["tStatementMasterTable"].Clone();
            //stmtMasterRemoveDup = DSstatementRaw.Tables["tStatementMasterTable"];
            //stmtMasterRemoveDup.DefaultView.Sort = "STATEMENTNUMBER ASC";
            //stmtMasterRemoveDup = stmtMasterRemoveDup.DefaultView.ToTable();

            ////stmtMasterRemoveDup.
            ////DSstatementRaw.Tables["tStatementMasterTable"].Rows

            //// Mahmoud AMR to remove duplicates

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
                //if (!string.IsNullOrEmpty(masterRow["CREDITCONTRACTS"].ToString()))
                //    continue;
                if (Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0)
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
                            clsBasFile.deleteFile(strFileDelivery);
                            clsBasFile.deleteFile(fileSummaryName);
                            return "Error in Generation ";//+ pBankName
                            }
                    strAccountFooter = String.Empty;
                    intAccountFooter = 0;
                    totAccountValue = 0;
                    calcAccountRows();
                    prevAccountNo = masterRow[mAccountno].ToString();
                    pageNo = 1;  //if page is based on account no 
                    if (totAccRows > 0 )
                        {
                        printHeader();//>>
                        printDelivery();//>>
                        }
                    isPrimaryOnly = false;
                    }
                //if (totAccRows < 1 && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0) //             || (masterRow[mCardno].ToString() == curMainCard   // Convert.ToDecimal(
                if (totAccRows < 1 ) //             || (masterRow[mCardno].ToString() == curMainCard   // Convert.ToDecimal(
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
                        CrDbDetail = "C";
                        totNetUsage += clsBasValid.validateNum(detailRow[dBilltranamount]);
                        }
                    else
                        {
                        CrDbDetail = "D";
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
                streamWritDelivery.Flush();
                streamWritDelivery.Close();
                fileStrmDelivery.Close();

                printStatementSummary();

                // Close Summary File
                streamSummary.Flush();
                streamSummary.Close();
                fileSummary.Close();

                ArrayList aryLstFiles = new ArrayList();
                aryLstFiles.Add(@strFileBasic);
                aryLstFiles.Add(@strFileTrans);
                aryLstFiles.Add(@strFileDelivery);
                clsBasFile.generateFileMD5(aryLstFiles, @clsBasFile.getPathWithoutExtn(pStrFileName) + ".MD5");// + "_Raw"
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

        CustomerPhone = "";
        emailRows = DSemails.Tables["Emails"].Select("idclient = " + masterRow[mClientid]);
        if (emailRows.Length != 0)
            CustomerPhone = emailRows[0][2].ToString().Trim();

        streamWritBasic.WriteLine(basText.alignmentLeft(basText.formatCardNumber(curMainCard), 19)
              + basText.alignmentLeft(masterRow[mCardbranchpartname].ToString(), 15)
              + basText.alignmentRight(getTranCount(curBranchVal, masterRow[mAccountno].ToString()), 6)
              + basText.alignmentLeft(masterRow[mCustomername].ToString(), 75)
              + basText.alignmentLeft(ValidateArbic(newaddress1), 50)
              + basText.alignmentLeft(ValidateArbic(newaddress2), 50)
              + basText.alignmentLeft(masterRow[mCustomercity].ToString(), 30)
              + basText.alignmentLeft(masterRow[mCustomerzipcode].ToString(), 10)
              + basText.alignmentRight(basText.formatNumUnSign(masterRow[mClosingbalance], "#,##0.00", 16), 16)
              + basText.alignmentRight(basText.formatNumUnSign(masterRow[mOpeningbalance], "#,##0.00", 16), 16)
              + basText.alignmentRight(basText.formatNumUnSign(masterRow[mTotaldebits], "#,##0.00", 16), 16)
              + basText.alignmentRight(basText.formatNumUnSign(masterRow[mTotalcredits], "#,##0.00", 16), 16)
              + basText.alignmentRight(basText.formatNumUnSign(masterRow[mAccountlim], "#,##0.00", 16), 16)
              + basText.alignmentRight(basText.formatNumUnSign(masterRow[mAccountavailablelim], "#,##0.00", 16), 16)
              + basText.alignmentRight(basText.formatNumUnSign(0, "#,##0.00", 16), 16)
              + basText.alignmentRight(basText.formatNumUnSign(masterRow[mMindueamount], "#,##0.00", 16), 16)
              + basText.alignmentRight(basText.formatNumUnSign(masterRow[mTotaloverdueamount], "#,##0.00", 16), 16)
              + basText.formatAS400Date(masterRow[mStetementduedate], 8, "R")
              + basText.formatAS400Date(masterRow[mStatementdateto], 8, "R")
              + basText.alignmentRight(basText.formatNumUnSign(masterRow[mMindueamount], "#,##0.00", 16), 16)
              + basText.alignmentLeft(CustomerPhone, 50)
              + basText.alignmentLeft(masterRow[mClientid].ToString(), 50)
            );
        totNoOfCardStat++;
        }


    protected virtual void printDetail()
        {
        DateTime trnsDate = clsBasValid.validateDate(detailRow[dTransdate]), postingDate = clsBasValid.validateDate(detailRow[dPostingdate]);
        string input;
        string program;
        string EPP;
        string Cash;
        string BT;
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
        if (trnsDesc == "Installment")
            {
            if (detailRow[dInstallmentData].ToString().Trim() != "")
                {
                input = detailRow[dInstallmentData].ToString().Trim();
                installdata = input;
                int index = detailRow[dInstallmentData].ToString().Trim().IndexOf(":");
                if (index > 0)
                    input = input.Substring(0, index);
                long installdocno = long.Parse(detailRow[dDocno].ToString());
                long originstalldocno = long.Parse(detailRow[dOrigDocNo].ToString());
                program = input.Substring(0, input.IndexOf(','));
                if (program == "Cash Instalment")
                    {
                    Cash = "Cash Request " + detailRow[dInstallmentLoan].ToString() + " instalment " + detailRow[dInstallmentCycleNo].ToString() + "/" + detailRow[dInstallmentCycles].ToString();
                    streamWritTrans.WriteLine(basText.alignmentLeft(basText.formatCardNumber(detailRow[dCardno].ToString()), 19)
            + basText.alignmentLeft(basText.formatCardNumber(curMainCard), 19)
            + CrDbDetail
            + basText.formatAS400Date(postingDate, 8, "R")
            + basText.formatAS400Date(trnsDate, 8, "R")
            + basText.alignmentLeft(detailRow[dRefereneno], 12)
            + basText.alignmentLeft(detailRow[dAccountcurrency], 3)
            + basText.alignmentRight("2", 2)
            + basText.alignmentRight(basText.formatNumUnSign(detailRow[dOrigtranamount], "#,##0.00", 16), 16)
            + basText.alignmentRight(basText.formatNumUnSign(detailRow[dBilltranamount], "#,##0.00", 16), 16)
            + basText.alignmentRight(Cash, 40)
            );
                    }
                else if (program == "Balance Transfer Instalment")
                    {
                    BT = "Balance Transfer " + detailRow[dInstallmentLoan].ToString() + " instalment " + detailRow[dInstallmentCycleNo].ToString() + "/" + detailRow[dInstallmentCycles].ToString();
                    streamWritTrans.WriteLine(basText.alignmentLeft(basText.formatCardNumber(detailRow[dCardno].ToString()), 19)
            + basText.alignmentLeft(basText.formatCardNumber(curMainCard), 19)
            + CrDbDetail
            + basText.formatAS400Date(postingDate, 8, "R")
            + basText.formatAS400Date(trnsDate, 8, "R")
            + basText.alignmentLeft(detailRow[dRefereneno], 12)
            + basText.alignmentLeft(detailRow[dAccountcurrency], 3)
            + basText.alignmentRight("2", 2)
            + basText.alignmentRight(basText.formatNumUnSign(detailRow[dOrigtranamount], "#,##0.00", 16), 16)
            + basText.alignmentRight(basText.formatNumUnSign(detailRow[dBilltranamount], "#,##0.00", 16), 16)
            + basText.alignmentRight(BT, 40)
            );
                    }
                else if (program == "Transaction Instalment")
                    {
                        if ((detailRow[dTrandescription].ToString() + detailRow[dMerchant].ToString()).Length >= 30)
                        {
                            EPP = (detailRow[dTrandescription].ToString() + detailRow[dMerchant].ToString()).Substring(0, 30) + " " + detailRow[dInstallmentLoan].ToString() + " instalment " + detailRow[dInstallmentCycleNo].ToString() + "/" + detailRow[dInstallmentCycles].ToString();
                        }
                        else
                        {
                            EPP = (detailRow[dTrandescription].ToString() + detailRow[dMerchant].ToString()) + " " + detailRow[dInstallmentLoan].ToString() + " instalment " + detailRow[dInstallmentCycleNo].ToString() + "/" + detailRow[dInstallmentCycles].ToString();
                        }
                    streamWritTrans.WriteLine(basText.alignmentLeft(basText.formatCardNumber(detailRow[dCardno].ToString()), 19)
            + basText.alignmentLeft(basText.formatCardNumber(curMainCard), 19)
            + CrDbDetail
            + basText.formatAS400Date(postingDate, 8, "R")
            + basText.formatAS400Date(trnsDate, 8, "R")
            + basText.alignmentLeft(detailRow[dRefereneno], 12)
            + basText.alignmentLeft(detailRow[dAccountcurrency], 3)
            + basText.alignmentRight("2", 2)
            + basText.alignmentRight(basText.formatNumUnSign(detailRow[dOrigtranamount], "#,##0.00", 16), 16)
            + basText.alignmentRight(basText.formatNumUnSign(detailRow[dBilltranamount], "#,##0.00", 16), 16)
            + basText.alignmentRight(EPP, 40)
            );
                    }
                }
            }
        else if (trnsDesc == "Early Settlement")
            {
            if (detailRow[dInstallmentData].ToString().Trim() != "")
                {
                input = detailRow[dInstallmentData].ToString().Trim();
                installdata = input;
                int index = detailRow[dInstallmentData].ToString().Trim().IndexOf(":");
                if (index > 0)
                    input = input.Substring(0, index);
                long installdocno = long.Parse(detailRow[dDocno].ToString());
                long originstalldocno = long.Parse(detailRow[dOrigDocNo].ToString());
                program = input.Substring(0, input.IndexOf(','));
                if (program == "Cash Instalment")
                    {
                    Cash = "Payoff Cash amount " + detailRow[dInstallmentLoan].ToString();
                    streamWritTrans.WriteLine(basText.alignmentLeft(basText.formatCardNumber(detailRow[dCardno].ToString()), 19)
            + basText.alignmentLeft(basText.formatCardNumber(curMainCard), 19)
            + CrDbDetail
            + basText.formatAS400Date(postingDate, 8, "R")
            + basText.formatAS400Date(trnsDate, 8, "R")
            + basText.alignmentLeft(detailRow[dRefereneno], 12)
            + basText.alignmentLeft(detailRow[dAccountcurrency], 3)
            + basText.alignmentRight("2", 2)
            + basText.alignmentRight(basText.formatNumUnSign(detailRow[dOrigtranamount], "#,##0.00", 16), 16)
            + basText.alignmentRight(basText.formatNumUnSign(detailRow[dBilltranamount], "#,##0.00", 16), 16)
            + basText.alignmentRight(Cash, 40)
            );
                    }
                else if (program == "Balance Transfer Instalment")
                    {
                    BT = "Payoff Balance transfer amount " + detailRow[dInstallmentLoan].ToString();
                    streamWritTrans.WriteLine(basText.alignmentLeft(basText.formatCardNumber(detailRow[dCardno].ToString()), 19)
            + basText.alignmentLeft(basText.formatCardNumber(curMainCard), 19)
            + CrDbDetail
            + basText.formatAS400Date(postingDate, 8, "R")
            + basText.formatAS400Date(trnsDate, 8, "R")
            + basText.alignmentLeft(detailRow[dRefereneno], 12)
            + basText.alignmentLeft(detailRow[dAccountcurrency], 3)
            + basText.alignmentRight("2", 2)
            + basText.alignmentRight(basText.formatNumUnSign(detailRow[dOrigtranamount], "#,##0.00", 16), 16)
            + basText.alignmentRight(basText.formatNumUnSign(detailRow[dBilltranamount], "#,##0.00", 16), 16)
            + basText.alignmentRight(BT, 40)
            );
                    }
                else if (program == "Transaction Instalment")
                    {
                    EPP = "Payoff " + trnsDesc.Substring(0, 30);
                    streamWritTrans.WriteLine(basText.alignmentLeft(basText.formatCardNumber(detailRow[dCardno].ToString()), 19)
            + basText.alignmentLeft(basText.formatCardNumber(curMainCard), 19)
            + CrDbDetail
            + basText.formatAS400Date(postingDate, 8, "R")
            + basText.formatAS400Date(trnsDate, 8, "R")
            + basText.alignmentLeft(detailRow[dRefereneno], 12)
            + basText.alignmentLeft(detailRow[dAccountcurrency], 3)
            + basText.alignmentRight("2", 2)
            + basText.alignmentRight(basText.formatNumUnSign(detailRow[dOrigtranamount], "#,##0.00", 16), 16)
            + basText.alignmentRight(basText.formatNumUnSign(detailRow[dBilltranamount], "#,##0.00", 16), 16)
            + basText.alignmentRight(EPP, 40)
            );
                    }
                }
            }
        else
            streamWritTrans.WriteLine(basText.alignmentLeft(basText.formatCardNumber(detailRow[dCardno].ToString()), 19)
       + basText.alignmentLeft(basText.formatCardNumber(curMainCard), 19)
       + CrDbDetail
       + basText.formatAS400Date(postingDate, 8, "R")
       + basText.formatAS400Date(trnsDate, 8, "R")
       + basText.alignmentLeft(detailRow[dRefereneno], 12)
       + basText.alignmentLeft(detailRow[dAccountcurrency], 3)
       + basText.alignmentRight("2", 2)
       + basText.alignmentRight(basText.formatNumUnSign(detailRow[dOrigtranamount], "#,##0.00", 16), 16)
       + basText.alignmentRight(basText.formatNumUnSign(detailRow[dBilltranamount], "#,##0.00", 16), 16)
       + basText.alignmentRight(trnsDesc.Trim(), 40)
       );
        totNoOfTransactions++;
        }

    protected virtual void printDelivery()
        {
        string newaddress1, newaddress2;
        clsMaintainData.fixAddress(curBranchVal, masterRow[mCustomeraddress1].ToString(), out newaddress1, out newaddress2);

        CustomerPhone = "";
        emailRows = DSemails.Tables["Emails"].Select("idclient = " + masterRow[mClientid]);
        if (emailRows.Length != 0)
            CustomerPhone = emailRows[0][2].ToString().Trim();

        CitizenID = "";
        PassportNoRows = DSPassportNos.Tables["PassportNos"].Select("idclient = " + masterRow[mClientid]);
        if (PassportNoRows.Length != 0)
            CitizenID = PassportNoRows[0][1].ToString().Trim();

        streamWritDelivery.WriteLine(basText.alignmentLeft(masterRow[mStatementnumber], 11) + fldSprtr
            + basText.alignmentLeft(masterRow[mCardbranchpartname].ToString(), 15) + fldSprtr
            + basText.alignmentLeft(masterRow[mCustomername].ToString(), 100) + fldSprtr
            + basText.alignmentLeft(ValidateArbic(newaddress1), 50) + fldSprtr
            + basText.alignmentLeft(ValidateArbic(newaddress2), 50) + fldSprtr
            + basText.alignmentLeft(CustomerPhone, 50) + fldSprtr
            + basText.alignmentLeft(CitizenID.ToString(), 15)
            );

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
        streamWritMD5.WriteLine(clsBasFile.getFileFromPath(strFileDelivery) + "  >>  " + clsBasFile.getFileMD5(strFileDelivery));
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

    private string GetMerchantByOrigDocNo(int pDocno)
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

    ~clsStatRawDataEGB()
        {
        DSstatementRaw.Dispose();
        }
    }
