using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Collections;


// Branch 5
public class clsStatementAibRaw_PRE : clsBasStatement
    {
    private string strBankName;
    private FileStream fileStrmBasic, fileStrmTrans, fileStrmSubtotal, fileSummary;
    private string strFileBasic, strFileTrans, strFileSubtotal;
    private StreamWriter streamWritBasic, streamWritTrans, streamWritSubtotal, streamSummary;
    //private DataSet DSstatement;
    //private OracleDataReader drPrimaryCards, drMaster,drDetail;
    private DataRow masterRow;
    private DataRow detailRow;
    //private string strEndOfLine = "\u000D" ;  //+ "M" ^M
    //	private string strEndOfPage = "\u000C" ; //+ basText.replicat(" ",85) + "F 2"  ;  //+ "\u000D"+ "M" ^L + ^M
    private const int MaxDetailInPage = 20; //
    private const int linesInLastPage = 67; //
    private int CurPageRec4Dtl = 0;
    private int pageNo = 0, totalCardPages = 0 //, totalPages=0
      , totalAccPages = 0, totCardRows = 0, totAccRows = 0, totAccCards = 0;
    //	private string lastPageTotal ;
    private string curCardNo, CardNumber = String.Empty, curCardNumber = String.Empty, PrevCardNumber = String.Empty;// 
    private string curAccountNo, prevAccountNo = String.Empty;//
    private string strAccountFooter;
    private int intAccountFooter;
    private decimal totNetUsage = 0;
    private decimal totAccountValue = 0;
    private DataRow[] cardsRows, accountRows;
    private string CrDbDetail;
    const string strFileSpr = "|";//#
    private bool isPrimaryOnly;
    protected bool isCorporateGenrated = false, isNotSplit = true, isAutoinalize = true;

    private string extAccNum;
    private string strOutputPath, strOutputFile, fileSummaryName;
    private DateTime vCurDate;
    private frmStatementFile frmMain;
    private int totRec = 1;
    private string mainTblCondStr = string.Empty;
    private string supTblCondStr = string.Empty;
    private string stmntType;
    private int totNoOfCardStat, totNoOfPageStat, totNoOfTransactions = 0, totNoOfTransactionsInt = 0;
    private DataRow ProductRow;
    protected string prepaidCond = string.Empty;
    private string accountNoName = mAccountno;
    private string accountLimit = mAccountlim;
    private string accountAvailableLimit = mAccountavailablelim;
    private string CustomerName = mCustomername;

    public clsStatementAibRaw_PRE()
        {
        }

    public string Statement(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pCurDate, string pStmntType, bool pAppendData)
        {
        string rtrnStr = "Successfully Generate " + pBankName, ppStrFileName = pStrFileName;
        int curMonth = pCurDate.Month;
        bool preExit = true;
        stmntType = pStmntType;
        try
            {
            clsMaintainData maintainData = new clsMaintainData();
            maintainData.matchCardBranch4Account(pBankCode);

            // merge transaction fee with original transaction
            //clsMaintainData maintainData = new clsMaintainData();
            //maintainData.mergeTrans(pBankCode);

            pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);
            vCurDate = pCurDate; //DateTime.Now.AddMonths(-1);
            strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
            clsBasFile.createDirectory(pStrFileName + vCurDate.ToString("yyyyMM") + pBankName);
            pStrFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM") + ".txt";
            strBankName = pBankName;
            isFullFields = false;
            if (mainTblCondStr != string.Empty)
                MainTableCond = mainTblCondStr; // " instr(m.contracttype,'" + strProductCond + "') > 0 ";//strWhereCond
            if (supTblCondStr != string.Empty)
                supTableCond = supTblCondStr;// " d.statementno in (select x.statementno from tstatementmastertable x where x.branch = " + pBankCode + " and instr(x.contracttype,'" + strProductCond + "') > 0)";

            strOutputFile = pStrFileName;

            // open output file
            strFileBasic = clsBasFile.getPathWithoutExtn(pStrFileName) + "_Basic.txt";
            clsBasFile.deleteFile(strFileBasic);
            fileStrmBasic = new FileStream(strFileBasic, FileMode.Create);
            streamWritBasic = new StreamWriter(fileStrmBasic);
            // open output file
            string strFileHeader = "Name" + strFileSpr + "Address1" + strFileSpr + "Address2" + strFileSpr + "Address3" +
              strFileSpr + "Address4" + strFileSpr + "Card Type" + strFileSpr + "Account" + strFileSpr + "Branch" + strFileSpr + "Statement Date" +
              strFileSpr + "Card No." + strFileSpr + "Credit Limit" + strFileSpr + "Available Credit" +
              strFileSpr + "Min. Payment Due" + strFileSpr + "New Balance" + strFileSpr + "Payment Due Date" +
              strFileSpr + "Past Due Amount" + strFileSpr + "prev.balance" + strFileSpr + "payment & CR" +
              strFileSpr + "Purch.Cash&Dr" + strFileSpr + "Finance Charge" + strFileSpr + "Pages" + strFileSpr + "External NO" + strFileSpr;
            streamWritBasic.WriteLine(strFileHeader);//#late charge#new balance

            strFileTrans = clsBasFile.getPathWithoutExtn(pStrFileName) + "_Trans.txt";
            clsBasFile.deleteFile(strFileTrans);
            fileStrmTrans = new FileStream(strFileTrans, FileMode.Create);
            streamWritTrans = new StreamWriter(fileStrmTrans);
            streamWritTrans.WriteLine("Card No." + strFileSpr + "Date of Trans" + strFileSpr + "Date of Post" + strFileSpr + "Reference" + strFileSpr + "Description" + strFileSpr + "Purchase Currency & Amount" + strFileSpr + "Amount" + strFileSpr + "");//detailRow[dBilltranamountsign]

            // open output file
            strFileSubtotal = clsBasFile.getPathWithoutExtn(pStrFileName) + "_Subtotal.txt";
            clsBasFile.deleteFile(strFileSubtotal);
            fileStrmSubtotal = new FileStream(strFileSubtotal, FileMode.Create);
            streamWritSubtotal = new StreamWriter(fileStrmSubtotal);
            streamWritSubtotal.WriteLine("Card No." + strFileSpr + "Card subtotal" + strFileSpr + "");

            fileSummaryName = pStrFileName;
            fileSummaryName = clsBasFile.getPathWithoutExtn(fileSummaryName) +
              "_Summary." + clsBasFile.getFileExtn(fileSummaryName);

            // open Summary file
            fileSummary = new FileStream(fileSummaryName, FileMode.Create); //Create
            streamSummary = new StreamWriter(fileSummary, Encoding.Default);

            // set branch for data
            curBranchVal = pBankCode; // pBankCode; // 5 ;  //1
            // data retrieve
            strOrder = " m.CARDBRANCHPART,m.externalno ";
            strMainTableCond += " m.accounttype in " + PrepaidCondition;
            strSubTableCond += " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.accounttype in " + PrepaidCondition + ")";

            FillStatementDataSet(pBankCode); //DSstatement =  //5); // 1
            getCardProduct(pBankCode);
            pageNo = 0; totalCardPages = 0;
            curCardNo = String.Empty;
            curAccountNo = String.Empty;
            frmMain.BeginInvoke(frmMain.setMinMaxProgressDelegate, new object[] { DSstatement.Tables["tStatementMasterTable"].Rows.Count });
            foreach (DataRow mRow in DSstatement.Tables["tStatementMasterTable"].Rows)
                {
                frmMain.BeginInvoke(frmMain.setProgressDelegate, new object[] { totRec++ });
                masterRow = mRow;
                //streamWrit.WriteLine(masterRow[mStatementno].ToString());
                pageNo = 1;  //if page is based on card no
                CurPageRec4Dtl = 0;
                cardsRows = mRow.GetChildRows(StatementNoDRel); //, DataRowVersion.Proposed
                //start new account
                if (masterRow[mHOLSTMT].ToString() == "Y")
                    continue;
                if (prevAccountNo != masterRow[mAccountno].ToString())
                    {
                    if (prevAccountNo == string.Empty)
                        if (curMonth != Convert.ToDateTime(masterRow[mStatementdateto]).Month)//,"dd/MM/yyyy"
                            {
                            preExit = false;
                            fileStrmBasic.Close();
                            fileStrmTrans.Close();
                            fileStrmSubtotal.Close();
                            clsBasFile.deleteFile(strFileBasic);
                            clsBasFile.deleteFile(strFileTrans);
                            clsBasFile.deleteFile(strFileSubtotal);
                            return "Error in Generation " + pBankName;
                            }
                    strAccountFooter = String.Empty;
                    intAccountFooter = 0;
                    totAccountValue = 0;
                    calcAccountRows();
                    prevAccountNo = masterRow[mAccountno].ToString();
                    pageNo = 1;  //if page is based on account no 
                    isPrimaryOnly = false;
                    }
                calcCardlRows();
                // Close Balance = 0 , card primary true, total rows > 0
                if (totCardRows < 1
                  || (masterRow[mCardprimary].ToString() == "Y"
                  && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0
                  && isPrimaryOnly == true)) //Convert.ToDecimal(
                    continue;

                // && (( &&&& Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0 
                // masterRow[mCardprimary].ToString() == "N")
                //if(masterRow[mCardprimary].ToString() == "Y" && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0.00)  
                //  continue;

                //if(totCardRows>0)
                //  printHeader();
                totNoOfCardStat++;

                curCardNumber = masterRow[mCardno].ToString();
                if (PrevCardNumber != curCardNumber) //&& PrevCardNumber != ""
                    {
                    printHeader();
                    pageNo = 1;  //if page is based on card no 
                    }
                printCardFooter();

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
                        printPageFooter();
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
                    completePageDetailRecords();
                    printPageFooter();
                    //printAccountFooter();
                    pageNo = 1; CurPageRec4Dtl = 0; //if pages is based on account
                    }
                //pageNo=1; CurPageRec4Dtl=0; if pages is based on card
                totNetUsage = 0;
                prevAccountNo = masterRow[mAccountno].ToString();
                PrevCardNumber = curCardNumber; //masterRow[mCardno].ToString();
                }
            //clsBasStatementRawData RawData = new clsBasStatementRawData(clsBasFile.getPathWithoutExtn(pStrFileName), DSstatement);
            //fillStatementHistory(pStmntType, pAppendData);
            //clsBasXML.WriteXmlToFile(DSstatement,clsBasFile.getPathWithoutExtn(pStrFileName)+ ".xml");

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
                // Close output ile
                streamWritBasic.Flush();
                streamWritBasic.Close();
                fileStrmBasic.Close();
                streamWritTrans.Flush();
                streamWritTrans.Close();
                fileStrmTrans.Close();
                streamWritSubtotal.Flush();
                streamWritSubtotal.Close();
                fileStrmSubtotal.Close();

                printStatementSummary();

                // Close Summary File
                streamSummary.Flush();
                streamSummary.Close();
                fileSummary.Close();

                ArrayList aryLstFiles = new ArrayList();
                //aryLstFiles.Add("");
                aryLstFiles.Add(strFileBasic);
                aryLstFiles.Add(strFileTrans);
                aryLstFiles.Add(strFileSubtotal);
                aryLstFiles.Add(fileSummaryName);

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
        //			string strHeader;
        //			streamWrit.Write(strEndOfPage);
        //			streamWrit.WriteLine(String.Empty);
        //			streamWrit.WriteLine(String.Empty);
        //			streamWrit.WriteLine(String.Empty);
        //			streamWrit.WriteLine(basText.alignmentLeft(masterRow["customername"],50)+basText.replicat(" ",15)+basText.alignmentLeft(masterRow[mAccounttype],35));
        //			streamWrit.WriteLine(basText.alignmentLeft(masterRow[mCustomeraddress1],50)+basText.replicat(" ",15)+basText.alignmentLeft(masterRow[mAccountno],35));
        //			streamWrit.WriteLine(basText.alignmentLeft(masterRow[mCustomeraddress2],50)+basText.replicat(" ",15)+basText.alignmentLeft(masterRow[mCardbranchpart].ToString()+ "  " + masterRow[mCardbranchpartname].ToString(),35));
        //			streamWrit.WriteLine(basText.alignmentLeft(masterRow[mCustomeraddress3],50)+basText.replicat(" ",15)+String.Format("{0,8:dd/MM/yy}",masterRow[mStatementdatefrom]));
        //			streamWrit.WriteLine(basText.alignmentLeft(masterRow[mCustomerregion].ToString() + " " +masterRow[mCustomercity].ToString(),50)+basText.replicat(" ",15)+String.Format("Page {0} of {1}",pageNo,totalAccPages));
        //			strHeader=basText.formatNum(masterRow[mAccountlim],"#,##0.00",9)+" "+basText.formatNum(masterRow[mTotalcredits],"#,##0.00",9)+" "+basText.formatNum(masterRow[mAccountavailablelim],"#,##0.00",9)+" "+basText.formatNum(masterRow[mMindueamount],"#,##0.00",15)+" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",15)+" "+basText.formatDate(masterRow[mStetementduedate],"dd/MM/yy");
        //			if(masterRow[mCardprimary].ToString() == "Y")
        //			{
        //				streamWrit.WriteLine("  "+basText.alignmentLeft(masterRow[mCardno],16)+" "+strHeader);
        //			}
        //			else
        //			{
        //				streamWrit.WriteLine(basText.replicat(" ",19)+strHeader);
        //			}
        //			streamWrit.WriteLine(String.Empty);
        //			streamWrit.WriteLine(String.Empty);
        //			streamWrit.WriteLine(String.Empty);
        //			streamWrit.WriteLine(String.Empty);
        //			streamWrit.WriteLine(String.Empty);
        extAccNum = clsBasValid.validateStr(masterRow[mExternalno]);
        if (extAccNum.Trim() == "")
            extAccNum = clsBasValid.validateStr(masterRow[mAccountno]);

        pageNo = 1;
        streamWritBasic.WriteLine(masterRow[CustomerName].ToString() + strFileSpr
            //+ ValidateArbic(masterRow[mCustomeraddress1].ToString()) + strFileSpr + ValidateArbic(masterRow[mCustomeraddress2].ToString())
          + ValidateArbic(newaddress1) + strFileSpr + ValidateArbic(newaddress2)
          + strFileSpr + ValidateArbic(masterRow[mCustomeraddress3].ToString()) + strFileSpr
          + masterRow[mCustomerregion].ToString()
          + masterRow[mCustomercity].ToString()
          + "" + strFileSpr + "" + masterRow[mContracttype].ToString() + strFileSpr + masterRow[accountNoName].ToString()
          + "" + strFileSpr + "" + masterRow[mCardbranchpart].ToString() + "  "
          + masterRow[mCardbranchpartname].ToString() + strFileSpr
          + String.Format("{0,8:dd/MM/yy}", masterRow[mStatementdateto])
            //+ strFileSpr + masterRow[mCardno].ToString() 
            // AIB-31 apply masked pan
          + strFileSpr + masterRow[mCardno].ToString().Substring(0, 6) + "******" + masterRow[mCardno].ToString().Substring(12, 4)
          + strFileSpr + masterRow[accountLimit].ToString()
          + strFileSpr + masterRow[accountAvailableLimit].ToString()
          + strFileSpr + masterRow[mMindueamount].ToString()
          + strFileSpr + masterRow[mClosingbalance].ToString()
          + CrDb(Convert.ToDecimal(masterRow[mClosingbalance].ToString())) + strFileSpr
          + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yy")
          + strFileSpr + masterRow[mTotaloverdueamount].ToString() + strFileSpr //0.00 
          + basText.formatNum(masterRow[mOpeningbalance], "#,##0.00", 12)
          + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance].ToString())) + strFileSpr
          //+ basText.formatNum(masterRow[mTotalpayments], "#,##0.00", 12)
          //+ DbCr(Convert.ToDecimal(masterRow[mTotalpayments])) + strFileSpr
          + basText.formatNum(Convert.ToDecimal(masterRow[mTotalpayments]) 
          + Convert.ToDecimal(masterRow[mTotalcredits]), "#,##0.00", 12) 
          + DbCr(Convert.ToDecimal(masterRow[mTotalpayments]) 
          + Convert.ToDecimal(masterRow[mTotalcredits]))+ strFileSpr
          + basText.formatNum(Convert.ToDecimal(masterRow[mTotalpurchases])
          + Convert.ToDecimal(masterRow[mTotalcashwithdrawal]), "#,##0.00", 12)
          + CrDbMinus(Convert.ToDecimal(masterRow[mTotalpurchases])
          + Convert.ToDecimal(masterRow[mTotalcashwithdrawal])) + strFileSpr
          + basText.formatNum(Convert.ToDecimal(masterRow[mTotalcharges])
          + Convert.ToDecimal(masterRow[mTotalinterest]), "#,##0.00", 12)
          + CrDbMinus(Convert.ToDecimal(masterRow[mTotalcharges])
          + Convert.ToDecimal(masterRow[mTotalinterest])) + strFileSpr
            //+ basText.formatNum(0.00,"#,##0.00",12) + "#" 
            //+ basText.formatNum(masterRow[mClosingbalance],"#,##0.00",12)  
            //+ "" + String.Format("Page {0} of {1}",pageNo,totalAccPages) + "" + strFileSpr + ""); //ACCOUNTTYPE// + "" + strFileSpr
          + String.Format("Page {0} of {1}", pageNo, totAccCards) + strFileSpr
          + extAccNum + strFileSpr); //masterRow[mExternalno].ToString()//ACCOUNTTYPE// + "" + strFileSpr
        totNoOfPageStat++;
        }


    protected void printDetail()
        {
        DateTime trnsDate = clsBasValid.validateDate(detailRow[dTransdate]), postingDate = clsBasValid.validateDate(detailRow[dPostingdate]);
        if (trnsDate > postingDate)
            trnsDate = postingDate;
        //			streamWrit.WriteLine("  {0:dd/MM} {1:dd/MM} {2,13} {3,-30} {4,12} {5,3} {6,12} {7,2}",clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]),basText.trimStr(detailRow[dRefereneno],13),basText.alignmentLeft(detailRow[dTrandescription],30),basText.formatNum(detailRow[dOrigtranamount],"#,##0.00;(#,##0.00)"),basText.Replace(clsBasValid.validateStr(detailRow[dOrigtrancurrency]),"XXX","   "),basText.formatNum(detailRow[dBilltranamount],"#,##0.00;(#,##0.00)"),CrDbDetail);//detailRow[dBilltranamountsign]
        //streamWritTrans.WriteLine("{0}" + strFileSpr + "{1:dd/MM}" + strFileSpr + "{2:dd/MM}" + strFileSpr + "{3,13}" + strFileSpr + "{4,-40}" + strFileSpr + "{5,12} {6,3}" + strFileSpr + "{7,12} {8,2}" + strFileSpr + "", masterRow[mCardno], trnsDate, postingDate, basText.trimStr(detailRow[dRefereneno], 13), basText.trimStr(detailRow[dTrandescription].ToString().Trim() + " " + detailRow[dMerchant].ToString().Trim(), 40), basText.formatNum(detailRow[dOrigtranamount], "#,##0.00;(#,##0.00)"), basText.Replace(clsBasValid.validateStr(detailRow[dOrigtrancurrency]), "XXX", "   "), basText.formatNum(detailRow[dBilltranamount], "#,##0.00;(#,##0.00)"), CrDbDetail);//detailRow[dBilltranamountsign]//clsBasValid.validateDate(detailRow[dTransdate])//clsBasValid.validateDate(detailRow[dPostingdate])
        // AIB-31 apply masked pan
        streamWritTrans.WriteLine("{0}" + strFileSpr + "{1:dd/MM}" + strFileSpr + "{2:dd/MM}" + strFileSpr + "{3,13}" + strFileSpr + "{4,-40}" + strFileSpr + "{5,12} {6,3}" + strFileSpr + "{7,12} {8,2}" + strFileSpr + "", masterRow[mCardno].ToString().Substring(0, 6) + "******" + masterRow[mCardno].ToString().Substring(12, 4), trnsDate, postingDate, basText.trimStr(detailRow[dRefereneno], 13), basText.trimStr(detailRow[dTrandescription].ToString().Trim() + " " + detailRow[dMerchant].ToString().Trim(), 40), basText.formatNum(detailRow[dOrigtranamount], "#,##0.00;(#,##0.00)"), basText.Replace(clsBasValid.validateStr(detailRow[dOrigtrancurrency]), "XXX", "   "), basText.formatNum(detailRow[dBilltranamount], "#,##0.00;(#,##0.00)"), CrDbDetail);//detailRow[dBilltranamountsign]//clsBasValid.validateDate(detailRow[dTransdate])//clsBasValid.validateDate(detailRow[dPostingdate])
        totNoOfTransactions++;
        }


    protected void printCardFooter()
        {
        //			string pageFooter;
        //			pageFooter =basText.replicat(" ",29) + "** Card . " + masterRow[mCardno] + " SUBTOTAL- "+ basText.replicat(" ",18) + basText.formatNum(Math.Abs(totNetUsage),"#,##0.00",-16) + CrDb(totNetUsage);
        //			streamWrit.WriteLine(pageFooter);
        //			CurPageRec4Dtl++;
        //			checkPageDetailRecords();
        //streamWritSubtotal.WriteLine(masterRow[mCardno] + "" + strFileSpr + "" + basText.formatNum(masterRow[mClosingbalance], "#,##0.00", 12) + "" + strFileSpr + ""); //" + strFileSpr + "
        // AIB-31 apply masked pan
        streamWritSubtotal.WriteLine(masterRow[mCardno].ToString().Substring(0, 6) + "******" + masterRow[mCardno].ToString().Substring(12, 4) + "" + strFileSpr + "" + basText.formatNum(masterRow[mClosingbalance], "#,##0.00", 12) + "" + strFileSpr + ""); //" + strFileSpr + "
        }


    protected void printPageFooter()
        {
        /*				streamWrit.WriteLine(String.Empty);
              streamWrit.WriteLine(String.Empty);
              streamWrit.WriteLine(String.Empty);
              streamWrit.WriteLine(String.Empty);
              streamWrit.WriteLine(String.Empty);
              string cardFooter;
              cardFooter =basText.formatNum(masterRow[mOpeningbalance],"#,##0.00",12)+" "+basText.formatNum(masterRow[mTotalpayments],"#,##0.00",12)+" "+basText.formatNum(Convert.ToDecimal(masterRow[mTotalpurchases])+Convert.ToDecimal(masterRow[mTotalcashwithdrawal]),"#,##0.00",12)+" "+basText.formatNum(masterRow[mTotalcharges],"#,##0.00",12)+" "+basText.formatNum(0.00,"#,##0.00",12)+" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",12);
              //streamWrit.WriteLine(String.Empty);
              streamWrit.WriteLine(cardFooter);
              /*if(masterRow[mCardprimary].ToString() != "Y")
              {
                strAccountFooter += cardFooter + "\r\n" ; //\r\n
                intAccountFooter++;
              }***/
        }


    protected void printAccountFooter()
        {
        /*			int curLastPageLine =CurPageRec4Dtl;
              curLastPageLine += 24; //add header Lines
              curLastPageLine+=2; // blank line after detail line total for last card
              curLastPageLine += intAccountFooter; //add Previous pages totals

              if(strAccountFooter.Length > 1)
              {
                strAccountFooter = strAccountFooter.Substring(0,strAccountFooter.Length-2);
                streamWrit.WriteLine(strAccountFooter);
                streamWrit.WriteLine(basText.replicat(" ",23) + "Total for all Cards " + basText.replicat(" ",42) + basText.formatNum(Math.Abs(totAccountValue),"#,##0.00",-16)+CrDb(totAccountValue));
                curLastPageLine++;
                //curLastPageLine += CurPageRec4Dtl +1; //add last page totals
              }

              curLastPageLine += 10; // subtract last page footer
              for( ;curLastPageLine<linesInLastPage;curLastPageLine++)
                streamWrit.WriteLine(String.Empty);
              streamWrit.WriteLine(" **(lastpage)");
              streamWrit.WriteLine(" "+basText.formatNum(masterRow[mOpeningbalance],"#,##0.00",-16));  //
              streamWrit.WriteLine(" "+basText.formatNum(masterRow[mTotalcashwithdrawal],"#,##0.00",-16));  //
              streamWrit.WriteLine(" "+basText.formatNum(masterRow[mTotalpurchases],"#,##0.00",-16));  //
              streamWrit.WriteLine(" "+basText.formatNum(masterRow[mTotalinterest],"#,##0.00",-16));  //
              streamWrit.WriteLine(" "+basText.formatNum(masterRow[mTotalcharges],"#,##0.00",-16));  //
              streamWrit.WriteLine(" "+basText.formatNum(masterRow[mTotalcredits],"#,##0.00",-16));  //
              streamWrit.WriteLine(" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",-16));  //
              streamWrit.WriteLine(" "+basText.alignmentLeft(masterRow[mStatementmessageline1],100));  //
              streamWrit.WriteLine(" "+basText.alignmentLeft(masterRow[mStatementmessageline2],100));  //
            */
        }

    private void calcCardlRows()
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


    private void calcAccountRows()
        {
        DataRow[] mainRows, supRows;
        string prevCardNo = String.Empty, CurCardNo = String.Empty;
        int currAccRowsPages = 0;
        int vSupRows = 0;
        totalAccPages = 0; //0
        totAccRows = 0;
        totAccCards = 0;

        mainRows = DSstatement.Tables["tStatementMasterTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[mAccountno]) + "'");
        foreach (DataRow mainRow in mainRows) //mRow.GetChildRows(StatementNoDRel)
            {
            totalAccPages++;
            CurCardNo = mainRow[mCardno].ToString();
            supRows = DSstatement.Tables["tStatementDetailTable"].Select("cardno = '" + CurCardNo + "'");

            //if(mainRow[mCardprimary].ToString() == "Y" 
            //  || Convert.ToDecimal(mainRow[mClosingbalance].ToString()) != 0)
            //  totAccCards++;
            vSupRows = 0;

            foreach (DataRow supRow in supRows) //mRow.GetChildRows(StatementNoDRel)
                {
                if ((supRow[dPostingdate] == DBNull.Value) && (supRow[dDocno] == DBNull.Value))
                    continue;// Exclude On-Hold Transactions 

                currAccRowsPages++;
                if (currAccRowsPages > MaxDetailInPage)//==
                    {
                    currAccRowsPages = 1;
                    totalAccPages++;
                    }
                vSupRows++;
                } // end of supRow
            currAccRowsPages = 0;

            if (vSupRows > 0) //Convert.ToDecimal(
                totAccCards++;
            } // end of mainRow

        if (totAccCards == 0 && (masterRow[mCardprimary].ToString() == "Y"
          && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) != 0))
            {
            totAccCards++;
            isPrimaryOnly = true;
            }


        /*
            if(prevCardNo != CurCardNo ) //&& prevCardNo != String.Empty
            {
              totalAccPages++;
              currAccRowsPages =1;
            }
            foreach (DataRow dtAccRow in accountRows) //mRow.GetChildRows(StatementNoDRel)
            {
              //streamWrit.WriteLine(basText.trimStr(dtAccRow[dTrandescription],40)); 
              currAccRowsPages++;
              totAccRows++;

              CurCardNo = dtAccRow[dCardno].ToString();
              if(currAccRowsPages==MaxDetailInPage)
              {
                currAccRowsPages =0;
                totalAccPages++;
              }
              prevCardNo = dtAccRow[dCardno].ToString();
            }
            if(currAccRowsPages> 0)
              totalAccPages++;
            if(totalAccPages < 1)
              totalAccPages = 1;
            */

        }



    private void completePageDetailRecords()
        {
        /*			//int curPageLine =CurPageRec4Dtl;
              for(int curPageLine =CurPageRec4Dtl;curPageLine<MaxDetailInPage;curPageLine++)
                streamWrit.WriteLine(String.Empty);*/
        }

    private void printFileMD5()
        {
        FileStream fileStrmMd5;
        StreamWriter streamWritMD5;
        fileStrmMd5 = new FileStream(strOutputPath + "\\" + strBankName + vCurDate.ToString("yyyyMM") + ".MD5", FileMode.Create);
        streamWritMD5 = new StreamWriter(fileStrmMd5);
        streamWritMD5.WriteLine(clsBasFile.getFileFromPath(strFileBasic) + "  >>  " + clsBasFile.getFileMD5(strFileBasic));
        streamWritMD5.WriteLine(clsBasFile.getFileFromPath(strFileTrans) + "  >>  " + clsBasFile.getFileMD5(strFileTrans));
        streamWritMD5.WriteLine(clsBasFile.getFileFromPath(strFileSubtotal) + "  >>  " + clsBasFile.getFileMD5(strFileSubtotal));
        streamWritMD5.Flush();
        streamWritMD5.Close();
        fileStrmMd5.Close();
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
                pageNo = 1;  //if page is based on card no
                CurPageRec4Dtl = 0;
                cardsRows = mRow.GetChildRows(StatementNoDRel); //, DataRowVersion.Proposed
                //start new account
                if (prevAccountNo != masterRow[mAccountno].ToString())
                    {
                    calcAccountRows();
                    prevAccountNo = masterRow[mAccountno].ToString();
                    pageNo = 1;  //if page is based on account no 
                    isPrimaryOnly = false;
                    }
                calcCardlRows();
                if (totCardRows < 1
                  || (masterRow[mCardprimary].ToString() == "Y"
                  && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0
                  && isPrimaryOnly == true)) //Convert.ToDecimal(
                    continue;
                totNoOfCardStat++;
                curCardNumber = masterRow[mCardno].ToString();
                if (PrevCardNumber != curCardNumber) //&& PrevCardNumber != ""
                    {
                    pageNo = 1;  //if page is based on card no 
                    }
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
                    }
                totAccountValue += totNetUsage;
                if (masterRow[mCardprimary].ToString() == "Y" && (masterRow[mCardstate].ToString() == "Given" || masterRow[mCardstate].ToString() == "Embossed" || masterRow[mCardstate].ToString() == "New" || masterRow[mCardstate].ToString() == "New Pin Generated Only"))
                    {
                    pageNo = 1; CurPageRec4Dtl = 0; //if pages is based on account
                    }
                //pageNo=1; CurPageRec4Dtl=0; if pages is based on card
                totNetUsage = 0;
                prevAccountNo = masterRow[mAccountno].ToString();
                PrevCardNumber = curCardNumber; //masterRow[mCardno].ToString();
                }
            if (totNoOfCardStat != 0)
            {
                StatSummary.NoOfStatements = totNoOfCardStat;
                StatSummary.NoOfTransactionsInt = 0;
                StatSummary.ProductCode = int.Parse(ProductRow["code"].ToString());
                StatSummary.ProductName = ProductRow["name"].ToString();
                StatSummary.InsertRecordDb(vCurDate.ToString("yyyyMM") + strBankName);
                totNoOfCardStat = 0;
                totNoOfTransactionsInt = 0;
            }
        }
        StatSummary = null;

        }

    private void makeZip()
        {
        }

    public string mainTblCond
        {
        get { return mainTblCondStr; }
        set { mainTblCondStr = value; }
        }  // mainTblCond

    public string supTblCond
        {
        get { return supTblCondStr; }
        set { supTblCondStr = value; }
        }  // supTblCond

    public string bankName
        {
        get { return strBankName; }
        set { strBankName = value; }
        }// bankName


    public string PrepaidCondition
        {
        get { return prepaidCond; }
        set { prepaidCond = value; }
        }// PrepaidCondition

    public frmStatementFile setFrm
        {
        set { frmMain = value; }
        }// setFrm

    ~clsStatementAibRaw_PRE()
        {
        DSstatement.Dispose();
        }
    }
