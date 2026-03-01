using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Collections;


// Branch X
public class clsStatRawData_AAIB : clsBasStatement
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
    //	protected string lastPageTotal ;
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
    //protected DataRelation StatementNoDRels;
    protected int curMonth;
    protected bool preExit = true;
    protected string stmntType;
    protected int totNoOfCardStat, totNoOfPageStat, totNoOfTransactions = 0, totNoOfTransactionsInt = 0;
    private DataRow ProductRow;
    protected bool hasInterset = false;
    private string strProductCond = string.Empty;

    public clsStatRawData_AAIB()
    {
    }

    public string Statement(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pDate, string pStmntType, bool pAppendData)
    {
        string rtrnStr = "Successfully Generate " + pBankName;
        curMonth = pDate.Month;
        clsMaintainData maintainData = new clsMaintainData();
        maintainData.matchCardBranch4Account(pBankCode);
        stmntType = pStmntType;

        // merge transaction fee with original transaction
        //clsMaintainData maintainData = new clsMaintainData();
        //maintainData.mergeTrans(pBankCode);
        pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);
        vCurDate = pDate; //DateTime.Now.AddMonths(-1);
        strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
        clsBasFile.createDirectory(pStrFileName + vCurDate.ToString("yyyyMM") + pBankName);
        pStrFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM") + ".txt";
        curBranchVal = pBankCode; // 4; //4  = real   1 = test
        //MainTableCond = " m.contracttype in " + strProductCond + "";//strWhereCond
        //supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + ")";

        // To Export all products
        if (strProductCond.Trim().Length > 0)
        {
            MainTableCond = " m.contracttype in " + strProductCond + "";//strWhereCond
            supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + ")";
        }
        else
        {
            MainTableCond = "";//strWhereCond
            supTableCond = "";
        }

        //FillStatementDataSet(pBankCode, ""); //DSstatement =  //6); // 6
        FillStatementDataSet_WithOverDueDays(pBankCode, ""); //DSstatement =  //6); // 6
        getCardProduct(pBankCode);
        Statement(pStrFileName, DSstatement);
        return "";
    }

    public string Statement(string pStrFileName, DataSet pDSstatement)
    {
        try
        {
            DSstatementRaw = pDSstatement;
            //StatementNoDRels = pDSstatement.Relations.Equals
            //DSstatementRaw.Relations.Clear();
            //StatementNoDRels = DSstatementRaw.Relations.Add("StaementNoDRs",
            //  DSstatementRaw.Tables["tStatementMasterTable"].Columns[mStatementno],
            //  DSstatementRaw.Tables["tStatementDetailTable"].Columns[dStatementno]); 

            // open output file
            strFileBasic = clsBasFile.getPathWithoutExtn(pStrFileName) + "_Main.txt";
            clsBasFile.deleteFile(strFileBasic);
            fileStrmBasic = new FileStream(strFileBasic, FileMode.Create);
            streamWritBasic = new StreamWriter(fileStrmBasic);
            streamWritBasic.AutoFlush = true;
            // open output file
            string strFileHeader = "CardHolder Name" + fldSprtr + "Address1" + fldSprtr + "Address2" + fldSprtr + "Address3" +
              fldSprtr + "Address4" + fldSprtr + "Card Type" + fldSprtr + "Account" + fldSprtr + "Branch" + fldSprtr + "Statement Date" +
              fldSprtr + "Card No." + fldSprtr + "Credit Limit" + fldSprtr + "Available Credit" +
              fldSprtr + "Min. Payment Due" + fldSprtr + "New Balance" + fldSprtr + "Payment Due Date" +
              fldSprtr + "Past Due Amount" + fldSprtr + "prev.balance" + fldSprtr + "payment & CR" +
              fldSprtr + "Purch.Cash&Dr" + fldSprtr + "Finance Charge" + fldSprtr + "External NO" + fldSprtr +"No. of Delay days" + fldSprtr;//+ fldSprtr + "Pages" 
            streamWritBasic.WriteLine(strFileHeader);//#late charge#new balance

            strFileTrans = clsBasFile.getPathWithoutExtn(pStrFileName) + "_Trns.txt";
            clsBasFile.deleteFile(strFileTrans);
            fileStrmTrans = new FileStream(strFileTrans, FileMode.Create);
            streamWritTrans = new StreamWriter(fileStrmTrans);
            streamWritTrans.WriteLine("Card No." + fldSprtr + "Date of Trans" + fldSprtr + "Date of Post" + fldSprtr + "Reference" + fldSprtr + "Description" + fldSprtr + "Purchase Amount" + fldSprtr + "Purchase Currency" + fldSprtr + "Billing Amount" + fldSprtr + "Billing Currency" + fldSprtr);//detailRow[dBilltranamountsign]
            //streamWritTrans.WriteLine("Card No." + fldSprtr + "Date of Trans" + fldSprtr + "Date of Post" + fldSprtr + "Reference" + fldSprtr + "Description" + fldSprtr + "Purchase Currency & Amount" + fldSprtr + "Amount" + fldSprtr);//detailRow[dBilltranamountsign]
            streamWritTrans.AutoFlush = true;

            fileSummaryName = pStrFileName;
            fileSummaryName = clsBasFile.getPathWithoutExtn(fileSummaryName) +
              "_Summary." + clsBasFile.getFileExtn(fileSummaryName);

            // open Summary file
            fileSummary = new FileStream(fileSummaryName, FileMode.Create); //Create
            streamSummary = new StreamWriter(fileSummary, Encoding.Default);
            streamSummary.AutoFlush = true;

            // open output file
            //strFileSubtotal = clsBasFile.getPathWithoutExtn(pStrFileName)+ "_Subtotal.txt";
            //clsBasFile.deleteFile(strFileSubtotal);
            //fileStrmSubtotal = new FileStream(strFileSubtotal,FileMode.Create);
            //streamWritSubtotal = new StreamWriter(fileStrmSubtotal);
            //streamWritSubtotal.WriteLine("Card No." + fldSprtr + "Card subtotal" + fldSprtr + "");

            // set branch for data
            //curBranchVal = pBankCode; // pBankCode; // 5 ;  //1
            // data retrieve
            //FillStatementDataSet(pBankCode); //DSstatementRaw =  //5); // 1
            pageNo = 0; totalCardPages = 0;
            curCardNo = String.Empty;
            curAccountNo = String.Empty;

            frmMain.BeginInvoke(frmMain.setMinMaxProgressDelegate, new object[] { DSstatement.Tables["tStatementMasterTable"].Rows.Count });
            foreach (DataRow mRow in DSstatementRaw.Tables["tStatementMasterTable"].Rows)
            {
                frmMain.BeginInvoke(frmMain.setProgressDelegate, new object[] { totRec++ });
                masterRow = mRow;
                //streamWrit.WriteLine(masterRow[mStatementno].ToString());
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
                //calcCardlRows();
                // Close Balance = 0 , card primary true, total rows > 0

                //if(totCardRows < 1
                //  || (masterRow[mCardprimary].ToString() == "Y" 
                //  && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0
                //  && isPrimaryOnly == true)) //Convert.ToDecimal(
                if (totAccRows < 1 && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0) //             || (masterRow[mCardno].ToString() == curMainCard   // Convert.ToDecimal(
                    continue;

                //totNoOfCardStat++;

                // && (( &&&& Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0 
                // masterRow[mCardprimary].ToString() == "N")
                //if(masterRow[mCardprimary].ToString() == "Y" && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0.00)  
                //  continue;

                //if(totCardRows>0)
                //  printHeader();

                curCardNumber = masterRow[mCardno].ToString();
                //if(PrevCardNumber != curCardNumber ) //&& PrevCardNumber != ""
                //if (prevAccountNo != masterRow[mAccountno].ToString())
                //{
                //  printHeader();
                //  pageNo=1;  //if page is based on card no 
                //}
                //printCardFooter();

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
                        //printAccountFooter();
                        //printPageFooter();
                        //printHeader();
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

                //pageNo++;//if pages is based on account
                //>if(totCardRows>0)
                //>printCardFooter();

                if (masterRow[mCardprimary].ToString() == "Y" && (masterRow[mCardstate].ToString() == "Given" || masterRow[mCardstate].ToString() == "Embossed" || masterRow[mCardstate].ToString() == "New" || masterRow[mCardstate].ToString() == "New Pin Generated Only"))
                {
                    //completePageDetailRecords();
                    //printPageFooter();
                    //printAccountFooter();
                    pageNo = 1; CurPageRec4Dtl = 0; //if pages is based on account
                }
                //pageNo=1; CurPageRec4Dtl=0; if pages is based on card
                totNetUsage = 0;
                //prevAccountNo = masterRow[mAccountno].ToString();
                PrevCardNumber = curCardNumber; //masterRow[mCardno].ToString();
            }
            //clsBasXML.WriteXmlToFile(DSstatementRaw,clsBasFile.getPathWithoutExtn(pStrFileName)+ ".xml");

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
                //streamWritSubtotal.Flush();
                //streamWritSubtotal.Close();
                //fileStrmSubtotal.Close();

                printStatementSummary();

                // Close Summary File
                streamSummary.Flush();
                streamSummary.Close();
                fileSummary.Close();

                ArrayList aryLstFiles = new ArrayList();
                //aryLstFiles.Add("");
                aryLstFiles.Add(@strFileBasic);
                aryLstFiles.Add(@strFileTrans);
                //aryLstFiles.Add(@strFileSubtotal);
                clsBasFile.generateFileMD5(aryLstFiles, @clsBasFile.getPathWithoutExtn(pStrFileName) + ".MD5");// + "_Raw"
                aryLstFiles.Add(@clsBasFile.getPathWithoutExtn(pStrFileName) + ".MD5");//+ "_Raw"
                SharpZip zip = new SharpZip();
                zip.createZip(aryLstFiles, @clsBasFile.getPathWithoutExtn(pStrFileName) + ".zip", "");//+ "_Raw"
            }
            //printFileMD5();
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
        streamWritBasic.WriteLine(masterRow[mCustomername].ToString() + fldSprtr
          //+ ValidateArbic(masterRow[mCustomeraddress1].ToString()) + fldSprtr + ValidateArbic(masterRow[mCustomeraddress2].ToString())
          + ValidateArbic(newaddress1) + fldSprtr + ValidateArbic(newaddress2)
          + fldSprtr + ValidateArbic(masterRow[mCustomeraddress3].ToString()) + fldSprtr
          + masterRow[mCustomerregion].ToString() + "  " + masterRow[mCustomercity].ToString() + fldSprtr
          + masterRow[mContracttype].ToString() + fldSprtr
          + masterRow[mAccountno].ToString() + fldSprtr
          + masterRow[mCardbranchpart].ToString() + "  "
          + masterRow[mCardbranchpartname].ToString() + fldSprtr
          + String.Format("{0,8:dd/MM/yy}", masterRow[mStatementdateto]) + fldSprtr
            //+ fldSprtr + masterRow[mCardno].ToString() 
          //+ basText.formatCardNumber(curMainCard) + fldSprtr
            // Jira update AAIB-9308
          + curMainCard + fldSprtr
          + masterRow[mAccountlim].ToString() + fldSprtr
          + masterRow[mAccountavailablelim].ToString() + fldSprtr
          + masterRow[mMindueamount].ToString() + fldSprtr
          + masterRow[mClosingbalance].ToString() + CrDb(Convert.ToDecimal(masterRow[mClosingbalance].ToString())) + fldSprtr
          + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yy") + fldSprtr
          + masterRow[mTotaloverdueamount].ToString() + fldSprtr
          + basText.formatNum(masterRow[mOpeningbalance], "#,##0.00", 12) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance].ToString())) + fldSprtr
          + basText.formatNum(masterRow[mTotalpayments], "#,##0.00", 12) + DbCr(Convert.ToDecimal(masterRow[mTotalpayments])) + fldSprtr
          + basText.formatNum(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal]), "#,##0.00", 12) + CrDbMinus(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal])) + fldSprtr
          + basText.formatNum(Convert.ToDecimal(masterRow[mTotalcharges]) + Convert.ToDecimal(masterRow[mTotalinterest]), "#,##0.00", 12) + CrDbMinus(Convert.ToDecimal(masterRow[mTotalcharges]) + Convert.ToDecimal(masterRow[mTotalinterest])) + fldSprtr
            //+ basText.formatNum(0.00,"#,##0.00",12) + "#" 
            //+ basText.formatNum(masterRow[mClosingbalance],"#,##0.00",12)  
            //+ "" + String.Format("Page {0} of {1}",pageNo,totalAccPages) + "" + fldSprtr + ""); //ACCOUNTTYPE// + "" + fldSprtr
            //+ String.Format("Page {0} of {1}", pageNo, totAccCards) + fldSprtr
          + extAccNum + fldSprtr
            // Adding Over due days  -------------- Jira update AAIB-9308
          + masterRow["OverDueDays"].ToString() + fldSprtr
          //+ "0" + fldSprtr  // OverDueDays = 0 as requested AAIB-12395 for once
          ); //masterRow[mExternalno].ToString()//ACCOUNTTYPE// + "" + fldSprtr
        //totNoOfPageStat++;
        totNoOfCardStat++;
    }


    protected virtual void printDetail()
    {
        DateTime trnsDate = clsBasValid.validateDate(detailRow[dTransdate]), postingDate = clsBasValid.validateDate(detailRow[dPostingdate]);
        if (trnsDate > postingDate)
            trnsDate = postingDate;
        string trnsDesc, strForeignCurr;
        //if (masterRow[mAccountcurrency].ToString().Trim() != detailRow[dOrigtrancurrency].ToString().Trim())
        //    strForeignCurr = basText.formatNumUnSign(detailRow[dOrigtranamount], "##,###,##0.00;(#,###,##0.00)", 13) + " " + basText.Replace(clsBasValid.validateStr(detailRow[dOrigtrancurrency]), "XXX", "   ");
        //else
        //    strForeignCurr = basText.replicat(" ", 16);
        if (masterRow[mAccountcurrency].ToString().Trim() != detailRow[dOrigtrancurrency].ToString().Trim())
            strForeignCurr = basText.formatNumUnSign(detailRow[dOrigtranamount], "##,###,##0.00;(#,###,##0.00)", 13) + fldSprtr + basText.Replace(clsBasValid.validateStr(detailRow[dOrigtrancurrency]), "XXX", "   ");
        else
            strForeignCurr = basText.replicat(" ", 13) + fldSprtr + " ";

        if (detailRow[dMerchant].ToString().Trim() == "")
            trnsDesc = detailRow[dTrandescription].ToString().Trim();
        else
            trnsDesc = detailRow[dMerchant].ToString().Trim();
        //			streamWrit.WriteLine("  {0:dd/MM} {1:dd/MM} {2,13} {3,-30} {4,12} {5,3} {6,12} {7,2}",clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]),basText.trimStr(detailRow[dRefereneno],13),basText.alignmentLeft(detailRow[dTrandescription],30),basText.formatNum(detailRow[dOrigtranamount],"#,##0.00;(#,##0.00)"),basText.Replace(clsBasValid.validateStr(detailRow[dOrigtrancurrency]),"XXX","   "),basText.formatNum(detailRow[dBilltranamount],"#,##0.00;(#,##0.00)"),CrDbDetail);//detailRow[dBilltranamountsign]
        //streamWritTrans.WriteLine("{0}" + fldSprtr + "{1:dd/MM}" + fldSprtr + "{2:dd/MM}" + fldSprtr + "{3,13}" + fldSprtr + "{4,-40}" + fldSprtr + "{5,12} {6,3}" + fldSprtr + "{7,12} {8,2}" + fldSprtr + "", curMainCard, trnsDate, postingDate, basText.trimStr(detailRow[dRefereneno], 13), basText.trimStr(trnsDesc.Trim(), 40), strForeignCurr, basText.formatNum(detailRow[dBilltranamount], "#,##0.00;(#,##0.00)"), CrDbDetail);//detailRow[dBilltranamountsign]//clsBasValid.validateDate(detailRow[dTransdate])//clsBasValid.validateDate(detailRow[dPostingdate]) //masterRow[mCardno]
        if (detailRow[dInstallmentData].ToString().Trim() != "")
            {
            if (trnsDesc == "Installment repayment")
                {
                string input = detailRow[dInstallmentData].ToString().Trim();
                int index = detailRow[dInstallmentData].ToString().Trim().IndexOf(":");
                if (index > 0)
                    input = input.Substring(0, index);
                //streamWritTrans.WriteLine(curMainCard + fldSprtr + trnsDate.ToString("dd/MM") + fldSprtr + postingDate.ToString("dd/MM") + fldSprtr + basText.trimStr(detailRow[dRefereneno], 25) + fldSprtr + basText.trimStr(basText.trimStr(input.Substring(0, input.IndexOf(',')) + input.Substring(basText.GetNthIndex(input, ',', 2)).Substring(input.Substring(basText.GetNthIndex(input, ',', 2)).IndexOf('#')).Replace('#', ' ') + " " + basText.trimStr(GetMerchantByOrigDocNo(int.Parse(detailRow[dOrigDocNo].ToString())), 40), 68), 40) + fldSprtr + strForeignCurr + fldSprtr + basText.formatNum(detailRow[dBilltranamount], "#,##0.00;(#,##0.00)") + CrDbDetail + fldSprtr);//detailRow[dBilltranamountsign]//clsBasValid.validateDate(detailRow[dTransdate])//clsBasValid.validateDate(detailRow[dPostingdate]) //masterRow[mCardno]
                streamWritTrans.WriteLine(curMainCard + fldSprtr + trnsDate.ToString("dd/MM") + fldSprtr + postingDate.ToString("dd/MM") + fldSprtr + basText.trimStr(detailRow[dRefereneno], 25) + fldSprtr + (basText.trimStr(basText.trimStr(input.Substring(0, input.IndexOf(',')) + input.Substring(basText.GetNthIndex(input, ',', 2)).Substring(input.Substring(basText.GetNthIndex(input, ',', 2)).IndexOf('#')).Replace('#', ' ') + " " + basText.trimStr(GetMerchantByOrigDocNo(long.Parse(detailRow[dOrigDocNo].ToString())), 40), 68), 40)).Replace("|", "\\") + fldSprtr + strForeignCurr + fldSprtr + basText.formatNum(detailRow[dBilltranamount], "#,##0.00;(#,##0.00)") + " " + CrDbDetail + fldSprtr + detailRow[dAccountcurrency] + fldSprtr);//detailRow[dBilltranamountsign]//clsBasValid.validateDate(detailRow[dTransdate])//clsBasValid.validateDate(detailRow[dPostingdate]) //masterRow[mCardno]
                }
            else
                //streamWritTrans.WriteLine(curMainCard + fldSprtr + trnsDate.ToString("dd/MM") + fldSprtr + postingDate.ToString("dd/MM") + fldSprtr + basText.trimStr(detailRow[dRefereneno], 25) + fldSprtr + basText.trimStr(trnsDesc.Trim(), 40) + fldSprtr + strForeignCurr + fldSprtr + basText.formatNum(detailRow[dBilltranamount], "#,##0.00;(#,##0.00)") + CrDbDetail + fldSprtr);//detailRow[dBilltranamountsign]//clsBasValid.validateDate(detailRow[dTransdate])//clsBasValid.validateDate(detailRow[dPostingdate]) //masterRow[mCardno]
                streamWritTrans.WriteLine(curMainCard + fldSprtr + trnsDate.ToString("dd/MM") + fldSprtr + postingDate.ToString("dd/MM") + fldSprtr + basText.trimStr(detailRow[dRefereneno], 25) + fldSprtr + basText.trimStr(trnsDesc.Trim(), 40).Replace("|", "\\") + fldSprtr + strForeignCurr + fldSprtr + basText.formatNum(detailRow[dBilltranamount], "#,##0.00;(#,##0.00)") + " " + CrDbDetail + fldSprtr + detailRow[dAccountcurrency] + fldSprtr);//detailRow[dBilltranamountsign]//clsBasValid.validateDate(detailRow[dTransdate])//clsBasValid.validateDate(detailRow[dPostingdate]) //masterRow[mCardno]
            }
        else
            //streamWritTrans.WriteLine(curMainCard + fldSprtr + trnsDate.ToString("dd/MM") + fldSprtr + postingDate.ToString("dd/MM") + fldSprtr + basText.trimStr(detailRow[dRefereneno], 25) + fldSprtr + basText.trimStr(trnsDesc.Trim(), 40) + fldSprtr + strForeignCurr + fldSprtr + basText.formatNum(detailRow[dBilltranamount], "#,##0.00;(#,##0.00)") + CrDbDetail + fldSprtr);//detailRow[dBilltranamountsign]//clsBasValid.validateDate(detailRow[dTransdate])//clsBasValid.validateDate(detailRow[dPostingdate]) //masterRow[mCardno]
            streamWritTrans.WriteLine(curMainCard + fldSprtr + trnsDate.ToString("dd/MM") + fldSprtr + postingDate.ToString("dd/MM") + fldSprtr + basText.trimStr(detailRow[dRefereneno], 25) + fldSprtr + basText.trimStr(trnsDesc.Trim(), 40).Replace("|", "\\") + fldSprtr + strForeignCurr + fldSprtr + basText.formatNum(detailRow[dBilltranamount], "#,##0.00;(#,##0.00)") + " " + CrDbDetail + fldSprtr + detailRow[dAccountcurrency] + fldSprtr);//detailRow[dBilltranamountsign]//clsBasValid.validateDate(detailRow[dTransdate])//clsBasValid.validateDate(detailRow[dPostingdate]) //masterRow[mCardno]
        totNoOfTransactions++;
    }


    //protected void printCardFooter()
    //{
    //streamWritSubtotal.WriteLine(masterRow[mCardno] + "" + fldSprtr + "" + basText.formatNum(masterRow[mClosingbalance],"#,##0.00",12) + "" + fldSprtr + ""); //" + fldSprtr + "
    //}

    protected void calcCardlRows()
    {
        totalCardPages = 0;
        totCardRows = 0;
        foreach (DataRow dtRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
        {
            if ((dtRow[dPostingdate] == DBNull.Value) && (dtRow[dDocno] == DBNull.Value))
                continue;// Exclude On-Hold Transactions 
            //streamWrit.WriteLine(basText.trimStr(dtRow[dTrandescription],40)); 
            totCardRows++;
        }
        totalCardPages = totCardRows / MaxDetailInPage;
        if ((totCardRows % MaxDetailInPage) > 0)
            totalCardPages++;
        if (totalCardPages < 1)
            totalCardPages = 1;
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
        //streamWritMD5.WriteLine(clsBasFile.getFileFromPath(strFileSubtotal) + "  >>  " + clsBasFile.getFileMD5(strFileSubtotal));
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
        //streamSummary.WriteLine("No of Pages        " + totNoOfPageStat.ToString());
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
                //totNoOfCardStat++;
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

    ~clsStatRawData_AAIB()
    {
        DSstatementRaw.Dispose();
    }
}
