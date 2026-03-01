using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Xml;
using System.Collections;


//Branch 10
public class clsStatement_CommonCompanyDtl : clsBasStatement
{
  private string strBankName;
  private FileStream fileStrm, fileSummary;
  private StreamWriter streamWrit, streamSummary;
  //private DataSet DSstatement;
  //private OracleDataReader drPrimaryCards, drMaster,drDetail;
  private DataRow masterRow;
  private DataRow detailRow;
  private string strEndOfLine = "\n";  //"\u000D"+ "M" ^M
  //		private string strEndOfPage = "\u000C"  ;  //+ "\u000D"+ "M" ^L + ^M
  private string strEndOfPage = "\f";  //"\u000C"  + basText.replicat(" ",85) + "F 2"  + "\u000D"+ "M" ^L + ^M
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
  private int totNoOfCardStat, totNoOfPageStat;

  private int curCardRow = 0, curTotCardRows = 0, numOfErr = 0, totNoOfTransactions = 0;
  private bool isPrimaryOnly, isHaveF3 = true, isPrimaryOpened = false;
  private FileStream fileStrmErr;
  private StreamWriter strmWriteErr;
  private string curMainCard, CurCardNo;

  private string extAccNum;
  private int totCrdNoInAcc, curCrdNoInAcc;
  private string strOutputPath, strOutputFile, fileSummaryName;
  private DateTime vCurDate;
  private clsOMR omr = new clsOMR();
  private int totPages;
  string endOfCustomer = string.Empty;
  string cProduct = string.Empty, curFileName = string.Empty;
  private ArrayList aryLstFiles;
  private frmStatementFile frmMain;
  private int totRec = 1;

  //private DataRow[] accCurrRows;
  private DataRow[] conractCurrRows, contractRows, contractAccRows, transRows;
  private string curContractCurrency = string.Empty, curContractNo = string.Empty, prvContractNo = string.Empty, curAccContractNo = string.Empty, prvAccContractNo = string.Empty;
  private int noOfAccInContract = 0, noOfTrnsInContract = 0;
  private int curNoOfAccInContract = 0, curNoOfTrnsInContract = 0;
  private int noOfCrdInContract = 0, noOfCrdInAcc = 0, noOfTrnsInAcc = 0;
  private int curNoOfCrdInAcc = 0, curNoOfTrnsInAcc = 0;
  DataView contractRowsView = null, accountRowsView = null, transRowsView = null;
  DataRowView cardRowsView = null,accRowView = null ;//contractRowView = null, 
  //DataTable contractRowsTbl = null;
  int totPageCont = 0, curPageCont = 0;
  private decimal compTotalOverdueAmount = 0, compMindueamount = 0;
  private decimal compTotaldebits = 0, compTotalcredits = 0, compOpeningbalance = 0, compTotalpayments = 0, compTotalpurchases = 0, compTotalcashwithdrawal = 0, compTotalcharges = 0, compTotalinterest = 0, compClosingbalance = 0;
  private string vStatementno = string.Empty, curCrdNo =string.Empty, prvCrdNo=string.Empty;
  private bool isCardHaveTrans = false;

  private DataRow[] rewardRows;
  private DataRow rewardRow;
  private bool createCorporateVal = false;
  private string accountNoName = mAccountno;
  private string accountLimit = mAccountlim;
  private string accountAvailableLimit = mAccountavailablelim;
  private string rewardCond = "'Reward Program'";//'New Reward Contract''Reward Contract'
  private decimal valOutstanding = 0, valCompOutstanding = 0;
  private decimal valTransaction = 0, valCompTransaction = 0, valCompMinDueAmount = 0;
  private DataRow[] dtlRows;

  //private bool isCorporate;
  public clsStatement_CommonCompanyDtl()
  {
  }

  public string Statement(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pCurDate, string pStmntType, bool pAppendData)
  {
    string rtrnStr = "Successfully Generate " + pBankName, ppStrFileName = pStrFileName;
    string strFileName = string.Empty;
    int curMonth = pCurDate.Month;
    bool preExit = true;
    aryLstFiles = new ArrayList();
    try
    {
      clsMaintainData maintainData = new clsMaintainData();
      maintainData.matchCardBranch4Account(pBankCode);

      pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);
      vCurDate = pCurDate; //DateTime.Now.AddMonths(-1);
      strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
      clsBasFile.deleteDirectory(strOutputPath);
      clsBasFile.createDirectory(strOutputPath);
      pStrFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM") + ".txt";
      strBankName = pBankName;

      strOutputFile = pStrFileName;

      // open output file
      //>fileStrm = new FileStream(pStrFileName, FileMode.Create); //Create
      //>streamWrit = new StreamWriter(fileStrm, Encoding.Default);

      // set branch for data
      curBranchVal = pBankCode; // 10; //3 = real   1 = test
      isFullFields = false;
      //     clsBasStatement.mainTableCond =  " substr(cardno,1,6) !='421192' "; //" cardproduct != 'Visa Business' ";
      //     clsBasStatement.supTableCond =  " substr(cardno,1,6) !='421192' "; 
      // data retrieve
      FillStatementDataSet(pBankCode); //DSstatement =  //10); //3
      fillAccountCurrencies(pBankCode);
      fillContractCurr(pBankCode);

      fileSummaryName = pStrFileName;
      fileSummaryName = clsBasFile.getPathWithoutExtn(fileSummaryName) +
        "_Summary." + clsBasFile.getFileExtn(fileSummaryName);
      // Error file
      fileStrmErr = new FileStream(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName), FileMode.Create);
      strmWriteErr = new StreamWriter(fileStrmErr, Encoding.Default);

      // open Summary file
      fileSummary = new FileStream(fileSummaryName, FileMode.Create); //Create
      streamSummary = new StreamWriter(fileSummary, Encoding.Default);

      pageNo = 1; totalCardPages = 0;
      curCardNo = String.Empty;
      curAccountNo = String.Empty;

      strFileName = pStrFileName;
      //dtlRowColl = DSstatement.Tables["tStatementDetailTable"].Rows;
      //dtlRowColl.Clear();
      foreach (DataRow accCurrRow in DSaccountCurrencies.Tables["AccountCurrencies"].Rows)//Currency Scope
      {
        conractCurrRows = DScontractCurr.Tables["ContractCurr"].Select(mAccountcurrency + " = '" + clsBasValid.validateStr(accCurrRow[mAccountcurrency]).ToString().Trim() + "'");
        pStrFileName = clsBasFile.getPathWithoutExtn(strFileName)
                + "_" + clsBasValid.validateStr(accCurrRow[mAccountcurrency]).ToString().Trim() + "." + clsBasFile.getFileExtn(strFileName);
        //Open output file for each Currency
        //fileStrm = new FileStream(pStrFileName, FileMode.Create); //Create
        //streamWrit = new StreamWriter(fileStrm, Encoding.Default);
        frmMain.BeginInvoke(frmMain.setMinMaxProgressDelegate, new object[] { conractCurrRows.Length });//DSstatement.Tables["tStatementMasterTable"].Rows.Count

        foreach (DataRow conractCurrRow in conractCurrRows)//Contract Scope //DScontractCurr.Tables["ContractCurr"].Rows
        {
          frmMain.BeginInvoke(frmMain.setProgressDelegate, new object[] { totRec++ });
          curContractCurrency = clsBasValid.validateStr(conractCurrRow[mAccountcurrency]).ToString().Trim();
          curContractNo = clsBasValid.validateStr(conractCurrRow[mContractno]).ToString().Trim();
          curNoOfTrnsInContract = 0;
          //contractRows = DSstatement.Tables["tStatementMasterTable"].Select(mAccountcurrency + " = '" + curContractCurrency + "' and " + mContractno + " = '" + curContractNo + "'");
          //contractRowsTbl = DSstatement.Tables["tStatementMasterTable"].Select(mAccountcurrency + " = '" + curContractCurrency + "' and " + mContractno + " = '" + curContractNo + "'");
          if (curContractNo == prvContractNo)
            continue;
          contractRowsView = new DataView(DSstatement.Tables["tStatementMasterTable"], mAccountcurrency + " = '" + curContractCurrency + "' and " + mContractno + " = '" + curContractNo + "'", mAccountno + "," + mCardaccountno + "," + mCardno, System.Data.DataViewRowState.CurrentRows);
          calcContractRows();//curContractNo, curContractCurrency
          if (noOfAccInContract < 1)
            continue;
          curPageCont = 1;
          //[if (totPageCont < 1)
          //[  continue;
          // make statement file for each product
          if (cProduct != conractCurrRow[mCardproduct].ToString().Trim())
          {
            if (cProduct != string.Empty)
            {
              streamWrit.Flush();
              streamWrit.Close();
              fileStrm.Close();
            }
            cProduct = conractCurrRow[mCardproduct].ToString().Trim();
            curFileName = clsBasFile.getPathWithoutExtn(pStrFileName)
              + "_" + cProduct.Replace(' ', '_') + "." + clsBasFile.getFileExtn(pStrFileName);
            //curFileName = pStrFileName;
            add2FileList(curFileName);
            fileStrm = new FileStream(curFileName, FileMode.Append); //FileMode.Create Create
            streamWrit = new StreamWriter(fileStrm, Encoding.Default);
          }

          //>>> Start Print Contact Pages
          curAccContractNo = prvAccContractNo = string.Empty;
          foreach (DataRowView cntRow in contractRowsView) //Account Scope //DSstatement.Tables["tStatementMasterTable"].Rows
          {
            curAccContractNo = clsBasValid.validateStr(cntRow[mCardaccountno]).ToString().Trim();
            curNoOfTrnsInAcc = 0;
            accRowView = cntRow;
            if (curAccContractNo == prvAccContractNo)
              continue;
            prvAccContractNo = curAccContractNo;
            //accountRowsView = new DataView(DSstatement.Tables["tStatementMasterTable"], mAccountcurrency + " = '" + curContractCurrency + "' and " + mContractno + " = '" + curContractNo + "' and " + mCardaccountno + " = '" + curAccContractNo + "'", mAccountno + "," + mCardaccountno + "," + mCardprimary + " desc," + mCardno, System.Data.DataViewRowState.CurrentRows);
            calcAccountRows(contractRowsView);

            curNoOfAccInContract++;
            if (curNoOfAccInContract == 1)
              printHeaderCompany(accRowView);// insert page header

            printDetailCompany(accRowView);

            if ((curNoOfAccInContract / MaxDetailInPage) > 0 && (curNoOfAccInContract % MaxDetailInPage) == 1) // change page
            {
              //curPageCont++;
              printFooterCompany(accRowView);// insert page footer
              printHeaderCompany(accRowView);// insert page header
            }
          }
          //>>> End Print Contact Pages

          accRowView = null;
          curNoOfAccInContract = 0;
          curAccContractNo = prvAccContractNo = string.Empty;
          foreach (DataRowView cntRow in contractRowsView) //Account Scope //DSstatement.Tables["tStatementMasterTable"].Rows
          {
            curAccContractNo = clsBasValid.validateStr(cntRow[mCardaccountno]).ToString().Trim();
            curNoOfTrnsInAcc = 0;
            accRowView = cntRow;
            if (curAccContractNo == prvAccContractNo)
            {
              continue;
            }

            curNoOfAccInContract++;
            if (curNoOfAccInContract == 1)
              printHeaderCompany(cntRow);// insert page header

            //
            if (curMonth != Convert.ToDateTime(cntRow[mStatementdateto]).Month)//,"dd/MM/yyyy"
            {
              preExit = false;
              fileStrm.Close();
              fileStrmErr.Close();
              clsBasFile.deleteFile(@strOutputFile);
              clsBasFile.deleteFile(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName));
              return "Error in Generation " + pBankName;
            }
            prvAccContractNo = curAccContractNo;
            accountRowsView = new DataView(DSstatement.Tables["tStatementMasterTable"], mAccountcurrency + " = '" + curContractCurrency + "' and " + mContractno + " = '" + curContractNo + "' and " + mCardaccountno + " = '" + curAccContractNo + "'", mAccountno + "," + mCardaccountno + "," + mCardprimary + " desc," + mCardno, System.Data.DataViewRowState.CurrentRows);
            calcAccountRows(accountRowsView);
            foreach (DataRowView accountRowView in accountRowsView) // Card Scope
            {
              curNoOfCrdInAcc++;
              cardRowsView = accountRowView;
              isCardHaveTrans = false;
              transRowsView = new DataView(DSstatement.Tables["tStatementDetailTable"], dAccountno + " = '" + curAccContractNo + "' and " + dCardno + " = '" + clsBasValid.validateStr(accountRowView[mCardno]).ToString().Trim() + "'", dAccountno + "," + dCardno + "," + dTransdate + "," + dPostingdate, System.Data.DataViewRowState.CurrentRows);
              foreach (DataRowView transRowView in transRowsView) // transaction Scope
              {
                if ((transRowView[dPostingdate] == DBNull.Value) && (transRowView[dDocno] == DBNull.Value))
                  continue;
                curCrdNo = transRowView[dCardno].ToString().Trim();
                if (curCrdNo != prvCrdNo)
                {
                  PrintAccountHeader();
                  isCardHaveTrans = true;
                }
                prvCrdNo = curCrdNo;
                printDetailCompany(transRowView);//print transaction detail
              }// transaction Scope //transRows
              if (isCardHaveTrans)
              {
                PrintAccountFooter();
              }
            }// Card Scope //accountRowsView
            if (curNoOfTrnsInAcc == 0 && Convert.ToDecimal(cntRow[mClosingbalance].ToString()) != 0)
            {
              PrintAccountHeader();
              PrintAccountFooter();
            }
          }//Account Scope //contractRows
          prvContractNo = curContractNo;
          //cProduct = conractCurrRow[mCardproduct].ToString().Trim();
        }//Contract Scope  //conractCurrRow
        //close output file for each Currency
        //streamWrit.Flush();
        //streamWrit.Close();
        //fileStrm.Close();

      }//Currency Scope //accCurrRow
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
        if (cProduct != string.Empty)
        {

          streamWrit.Flush();
          streamWrit.Close();
          fileStrm.Close();
        }

        printStatementSummary();

        // Close Summary File
        streamSummary.Flush();
        streamSummary.Close();
        fileSummary.Close();

        strmWriteErr.Flush();
        strmWriteErr.Close();
        fileStrmErr.Close();

        if (numOfErr == 0)
          clsBasFile.deleteFile(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName));
        else
          rtrnStr += string.Format(" with {0} Errors", numOfErr);
        /*else
          clsBasFile.deleteFile(pStrFileName);*/
        //ArrayList aryLstFiles = new ArrayList();
        //aryLstFiles.Add("");
        //aryLstFiles.Add(@strOutputFile);
        numOfErr += validateNoOfLines(aryLstFiles, 48);
        aryLstFiles.Add(fileSummaryName);
        ///
        //if (pStmntType == "Corporate")
        //{
        //  ArrayList aryLstFilesCorp = new ArrayList();
        //  clsStatement_CommonCorpCmpny stmntBNPCorp = new clsStatement_CommonCorpCmpny();// + "NSGB_Business_Statement.txt"
        //  stmntBNPCorp.setFrm = frmMain;
        //  aryLstFilesCorp = stmntBNPCorp.Statement(ppStrFileName, pBankName, pBankCode, pStrFile, pCurDate);// + "NSGB_Business_Statement.txt"
        //  foreach (object str in aryLstFilesCorp)
        //    aryLstFiles.Add((string)str);
        //
        //  stmntBNPCorp = null;
        //}
        ///
        clsBasFile.generateFileMD5(aryLstFiles, @clsBasFile.getPathWithoutExtn(pStrFileName) + ".MD5");
        aryLstFiles.Add(@clsBasFile.getPathWithoutExtn(pStrFileName) + ".MD5");
        SharpZip zip = new SharpZip();
        zip.createZip(aryLstFiles, @clsBasFile.getPathWithoutExtn(pStrFileName) + ".zip", "");

        DSstatement.Dispose();
      }
    }
    return rtrnStr;
  }


  protected void printHeaderCompany(DataRowView pContractRowsView)
  {
    string newaddress1, newaddress2;
    clsMaintainData.fixAddress(curBranchVal, pContractRowsView[mCustomeraddress1].ToString(), out newaddress1, out newaddress2);
    totPages++; endOfCustomer = ""; CurPageRec4Dtl = 1;//0
    extAccNum = clsBasValid.validateStr(pContractRowsView[mExternalno]);
    if (extAccNum.Trim() == "")
      extAccNum = clsBasValid.validateStr(pContractRowsView[mCardaccountno]);

    //if(pContractRowsView[mCardprimary].ToString() == "Y")
    if (curPageCont == 1 && totPageCont == 1) // statement contain 1 page
    {
      CurrentPageFlag = "F 0";
      isHaveF3 = true;
    }
    else if (curPageCont == 1 && totPageCont > 1)  //first page of multiple page statement
      CurrentPageFlag = "F 1"; // //middle page of multiple page statement
    else if (curPageCont < totPageCont)
      CurrentPageFlag = "F 2";
    else if (curPageCont == totPageCont) //last page of multiple page statement
    {
      CurrentPageFlag = "F 3";
      isHaveF3 = true;
      endOfCustomer = "/";
    }

    streamWrit.WriteLine(strEndOfPage);
    streamWrit.WriteLine(basText.replicat(" ", 81) + CurrentPageFlag);  //+"{0,5:MM/yy}",pContractRowsView[mStatementdatefrom]
    streamWrit.WriteLine(basText.replicat(" ", 81) + pContractRowsView[mCardproduct]);  //+"{0,5:MM/yy}",pContractRowsView[mStatementdatefrom]
    //			streamWrit.WriteLine(basText.alignmentLeft(pContractRowsView[mCustomername],50)+ basText.replicat(" ",31) + pContractRowsView[mCardbranchpartname]);  //
    streamWrit.WriteLine(basText.alignmentLeft(pContractRowsView[mCustomername], 50));  //
    streamWrit.WriteLine(basText.replicat(" ", 81) + pContractRowsView[mCardbranchpartname]);  //
    //			streamWrit.WriteLine(basText.alignmentMiddle(pContractRowsView[mCustomeraddress1],50)+ basText.replicat(" ",31) + clsBasValid.validateStr(pContractRowsView[mCardaccountno])); //basText.formatNum(pContractRowsView[mCardno],"####-####-####-####")
    //string x = Convert.ToString(ValidateArbic(pContractRowsView[mCustomeraddress1].ToString()));
    //streamWrit.WriteLine(basText.alignmentMiddle(ValidateArbic(pContractRowsView[mCustomeraddress1].ToString()), 50)); //basText.formatNum(pContractRowsView[mCardno],"####-####-####-####")
    streamWrit.WriteLine(basText.alignmentMiddle(ValidateArbic(newaddress1), 50)); //basText.formatNum(pContractRowsView[mCardno],"####-####-####-####")
    streamWrit.WriteLine(String.Empty);//basText.replicat(" ", 81) + extAccNum //clsBasValid.validateStr(pContractRowsView[mExternalno]) //accountno  //basText.formatNum(pContractRowsView[mCardno],"####-####-####-####")
    //			streamWrit.WriteLine(basText.alignmentMiddle(pContractRowsView[mCustomeraddress2],50)+ basText.replicat(" ",31)+"{0,10:dd/MM/yyyy}",pContractRowsView[mStatementdatefrom]);//+ basText.formatNum(pContractRowsView[mAccountlim],"########")  basText.formatNum(pContractRowsView[mCardlimit],"########")
    //streamWrit.WriteLine(basText.alignmentMiddle(ValidateArbic(pContractRowsView[mCustomeraddress2].ToString()), 50));
    streamWrit.WriteLine(basText.alignmentMiddle(ValidateArbic(newaddress2), 50));
    streamWrit.WriteLine(basText.replicat(" ", 81) + "{0,10:dd/MM/yyyy}", pContractRowsView[mStatementdateto]);//Statementdatefrom+ basText.formatNum(pContractRowsView[mAccountlim],"########")  basText.formatNum(pContractRowsView[mCardlimit],"########")
    //			streamWrit.WriteLine(basText.alignmentMiddle(pContractRowsView[mCustomeraddress3].ToString().Trim() + " " + pContractRowsView[mCustomerregion].ToString().Trim() + " " + pContractRowsView[mCustomercity].ToString().Trim(),50)+ basText.replicat(" ",31)+curPageCont + " / " + totPageCont);  //basText.formatNum(pContractRowsView[mCardaccountno],"##-##-######-####")
    streamWrit.WriteLine(basText.alignmentMiddle(ValidateArbic(pContractRowsView[mCustomeraddress3].ToString().Trim()) + " " + pContractRowsView[mCustomerregion].ToString().Trim() + " " + pContractRowsView[mCustomercity].ToString().Trim(), 50));  //basText.formatNum(pContractRowsView[mCardaccountno],"##-##-######-####")
    streamWrit.WriteLine(basText.replicat(" ", 81) + curPageCont + " / " + totPageCont);  //basText.formatNum(pContractRowsView[mCardaccountno],"##-##-######-####")
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    //streamWrit.WriteLine(omr.Asterisk(curPageCont));//String.Empty
    streamWrit.WriteLine(omr.Asterisk(totPages) + omr.AsteriskPage4(curPageCont));//String.Empty
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(basText.alignmentMiddle(pContractRowsView[mContractno], 20) + basText.alignmentMiddle(pContractRowsView[mContractlimit], 13) + basText.formatNum(pContractRowsView[mAccountavailablelim], "##0", 20) + basText.alignmentMiddle(compMindueamount, 13) + basText.formatDate(pContractRowsView[mStetementduedate], "dd/MM/yyyy", 15, "M") + basText.formatNum(compTotalOverdueAmount, "##0.00", 13)); //+ basText.formatNum(pContractRowsView[mAccountlim],"#,##0.00",9)+" "+basText.formatNum(pContractRowsView[mTotalcredits],"#,##0.00",9)+" "+basText.formatNum(pContractRowsView[mAccountavailablelim],"#,##0.00",9)+" "+basText.formatNum(pContractRowsView[mMindueamount],"#,##0.00",15)+" "+basText.formatNum(pContractRowsView[mClosingbalance],"#,##0.00",15)+" "+basText.formatDate(pContractRowsView[mStetementduedate],"dd/MM/yy") //TOTALCREDITS// basText.alignmentMiddle(pContractRowsView[mAccountavailablelim],20) //"#,##0"
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    //if (curPageCont == 1)
    //  streamWrit.WriteLine(basText.replicat(" ", 43) + basText.alignmentLeft("Previous Balance", 63) + basText.formatNumUnSign(pContractRowsView[mOpeningbalance], "##0.00", 16) + " " + CrDb(Convert.ToDecimal(pContractRowsView[mOpeningbalance])));  //+ CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign]))
    //else
    streamWrit.WriteLine(String.Empty);

    streamWrit.WriteLine(String.Empty);
    curPageCont++;
    //totalPages++;
    //totNoOfPageStat++;
  }

  private void printDetailCompany(DataRowView pAccRowView)
  {
    valOutstanding = valTransaction = 0;
    dtlRows = DSstatement.Tables["tStatementDetailTable"].Select("accountno = '" + pAccRowView[dAccountno].ToString() + "'"); //mCardaccountno accountno
    foreach (DataRow dtlRow in dtlRows)
    {
      if ((dtlRow[dPostingdate] == DBNull.Value) && (dtlRow[dDocno] == DBNull.Value))
        valOutstanding += Convert.ToDecimal(dtlRow[dBilltranamount].ToString());
      else
        valTransaction += Convert.ToDecimal(dtlRow[dBilltranamount].ToString());
    }
    valCompOutstanding += valOutstanding;
    valCompTransaction += valTransaction;// Convert.ToDecimal(pAccRowView[mClosingbalance].ToString()); //valTransaction;
    //valCompMinDueAmount += Convert.ToDecimal(pAccRowView[mMindueamount].ToString()); //valTransaction;

    //streamWrit.WriteLine("  {0,16}  {1,-50} {2,16}", pAccRowView[mCardno].ToString(), pAccRowView[mCardclientname].ToString(), basText.formatNumUnSign(pAccRowView[mMindueamount], "#0.00", 16));   // {3,2}    basText.formatNumUnSign(masterRow[mClosingbalance],"#0.00",16),CrDb(Convert.ToDecimal(masterRow[mClosingbalance])) {3,16}  ,valOutstanding,valTransaction  basText.formatNumUnSign(masterRow[mMindueamount],"#0.00",16),{3,16} 
  }


  protected void printDetailCompany2(DataRowView pTransRowView)
  {
    if (pTransRowView == null)
      return;
    DateTime trnsDate = clsBasValid.validateDate(pTransRowView[dTransdate]), postingDate = clsBasValid.validateDate(pTransRowView[dPostingdate]);
    if (trnsDate > postingDate)
      trnsDate = postingDate;

    if (cardRowsView[mAccountcurrency].ToString().Trim() != pTransRowView[dOrigtrancurrency].ToString().Trim())
      strForeignCurr = basText.alignmentRight(basText.formatNumUnSign(pTransRowView[dOrigtranamount], "#,##0.00;(#,##0.00)", 13) + " " + basText.Replace(clsBasValid.validateStr(pTransRowView[dOrigtrancurrency]), "XXX", "   "),15);
    else
      strForeignCurr = basText.replicats(" ", 15);

    if (strForeignCurr.Trim() == "0")
      strForeignCurr = basText.replicat(" ", 15);

    string trnsDesc; //= detailRow[dTrandescription].ToString().Trim() + " " + detailRow[dMerchant].ToString().Trim()
    if (pTransRowView[dMerchant].ToString().Trim() == "")
      trnsDesc = pTransRowView[dTrandescription].ToString().Trim();
    else
      trnsDesc = pTransRowView[dMerchant].ToString().Trim();

    //			streamWrit.WriteLine("  {0:dd/MM}  {1:dd/MM}  {2,-40} {3,12} {4,3} {5,16} {6,2}",trnsDate,postingDate,basText.trimStr(detailRow[dTrandescription].ToString().Trim() + " " + detailRow[dMerchant].ToString().Trim(),45),basText.formatNum(detailRow[dOrigtranamount],"#,##0.00;(#,##0.00)"),basText.Replace(clsBasValid.validateStr(detailRow[dOrigtrancurrency]),"XXX","   "),basText.formatNum(detailRow[dBilltranamount],"#,##0.00;(#,##0.00)"),CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + isSupplementCard(clsBasValid.validateStr(pContractRowsView[mCardprimary].ToString()))); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)
    //streamWrit.WriteLine("  {0:dd/MM}  {1:dd/MM}  {2,-24}  {3,-40} {4,16} {5,16} {6,2}", trnsDate, postingDate, basText.trimStr(pTransRowView[dRefereneno].ToString().Trim(), 20), basText.trimStr(trnsDesc, 40), strForeignCurr, basText.formatNumUnSign(pTransRowView[dBilltranamount], "#,###,##0.00", 14), CrDb(clsBasValid.validateStr(pTransRowView[dBilltranamountsign])) + " " + isSupplementCard(clsBasValid.validateStr(cardRowsView[mCardprimary].ToString()))); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)
    printDetailCompany("  " + trnsDate.ToString("dd/MM") + "  " + postingDate.ToString("dd/MM") + "  "
      + basText.alignmentLeft((object)pTransRowView[dRefereneno].ToString().Trim(), 24)+ "  "
      + basText.alignmentLeft((object)trnsDesc, 40) + "  " + strForeignCurr + "  "
      + basText.alignmentRight(basText.formatNumUnSign(pTransRowView[dBilltranamount], "#,##0.00", 15),15)+ " " + CrDb(clsBasValid.validateStr(pTransRowView[dBilltranamountsign])) 
      +" " + isSupplementCard(clsBasValid.validateStr(cardRowsView[mCardprimary].ToString())));

    //curNoOfTrnsInAcc++; curNoOfTrnsInContract++; CurPageRec4Dtl++;
  }

  protected void printDetailCompany(string pPrintStr)
  {
    curNoOfTrnsInAcc++; curNoOfTrnsInContract++; CurPageRec4Dtl++;
    if ((curNoOfTrnsInContract / MaxDetailInPage) > 0 && (curNoOfTrnsInContract % MaxDetailInPage) == 1) // change page
    {
      //curPageCont++;
      printFooterCompany(accRowView);// insert page footer
      printHeaderCompany(accRowView);// insert page header
    }
    streamWrit.WriteLine(pPrintStr);
    if (curNoOfTrnsInContract == noOfTrnsInContract)//insert page footer
    {
      completePageDetailRecords();
      printFooterCompany(accRowView);
    }

  }

  protected void printFooterCompany(DataRowView pContractRowsView)
  {
    //completePageDetailRecords();
    streamWrit.WriteLine(String.Empty);
    //if (curPageCont == totPageCont)
    //  streamWrit.WriteLine(basText.replicat(" ", 17) + basText.alignmentLeft("Current Balance", 63) + basText.formatNumUnSign(pContractRowsView[mClosingbalance], "#0.00", 16) + " " + CrDb(Convert.ToDecimal(pContractRowsView[mClosingbalance])));
    ////streamWrit.WriteLine(basText.replicat(" ", 80) + basText.formatNumUnSign(pContractRowsView[mClosingbalance], "#0.00", 16) + " " + CrDb(Convert.ToDecimal(pContractRowsView[mClosingbalance])));
    //else
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    //streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(basText.formatNum((object)compTotaldebits, "#0.00;(#0.00)", 20, "L") + basText.formatNum((object)compTotalcredits, "#0.00;(#0.00)", 20, "L"));//streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(basText.alignmentRight(pContractRowsView[mContractno], 35) + basText.alignmentRight((object)compMindueamount, 20) + basText.formatDate(pContractRowsView[mStetementduedate], "dd/MM/yyyy", 15, "R")
   + basText.alignmentRight(basText.formatNumUnSign((object)compOpeningbalance, "#,###,##0.00", 12) + " " + CrDb(compOpeningbalance), 15)
   + basText.alignmentRight(basText.formatNumUnSign((object)compTotalpayments, "#,###,##0.00", 12) + " " + DbCr(compTotalpayments), 15)
   + basText.alignmentRight(basText.formatNum((object)(compTotalpurchases + compTotalcashwithdrawal), "#,###,##0.00", 12) + CrDbMinus(compTotalpurchases + compTotalcashwithdrawal), 15)
   + basText.alignmentRight(basText.formatNumUnSign((object)compTotalcharges, "#,###,##0.00", 12) + " " + CrDbMinus(Convert.ToDecimal(pContractRowsView[mTotalcharges])), 15)
   + basText.alignmentRight(basText.formatNumUnSign((object)compTotalinterest, "#,###,##0.00", 12) + " " + CrDbMinus(compTotalinterest), 15)
   + basText.alignmentRight(basText.formatNumUnSign((object)compClosingbalance, "#,###,##0.00", 16) + " " + CrDb(compClosingbalance), 15));
      //       + basText.formatNum(Convert.ToDecimal(pContractRowsView[mTotalcharges]) + Convert.ToDecimal(pContractRowsView[mTotalinterest]),"#,##0.00",12) + CrDbMinus(Convert.ToDecimal(pContractRowsView[mTotalcharges]) + Convert.ToDecimal(pContractRowsView[mTotalinterest])) 
  }

  private void PrintAccountHeader()
  {
    printDetailCompany("  " + basText.replicats(" ", 5) + "  " + basText.replicats(" ", 5) + "  "
     + basText.replicats(" ", 24) + "  " + "Card No : " + curCrdNo);
    printDetailCompany("  " + basText.replicats(" ", 5) + "  " + basText.replicats(" ", 5) + "  "
      + basText.replicats(" ", 24) + "  " + basText.alignmentLeft((object)"Opening Balance >>", 40)
      + basText.replicats(" ", 18) + basText.alignmentRight((object)basText.formatNumUnSign(accRowView[mOpeningbalance], "##0.00", 16), 16) + " " + CrDb(Convert.ToDecimal(accRowView[mOpeningbalance])));
  }

  private void PrintAccountFooter()
  {
    printDetailCompany("  " + basText.replicats(" ", 5) + "  " + basText.replicats(" ", 5) + "  "
  + basText.replicats(" ", 24) + "  " + basText.alignmentLeft((object)"Closing Balance >>", 40)
  + basText.replicats(" ", 18) + basText.alignmentRight((object)basText.formatNumUnSign(accRowView[mClosingbalance], "##0.00", 16), 16) + " " + CrDb(Convert.ToDecimal(accRowView[mOpeningbalance])));
    printDetailCompany("  " + basText.replicats(" ", 5) + "  " + basText.replicats(" ", 5) + "  "
    + basText.replicats(" ", 24) + "  " + basText.alignmentLeft((object)"Minimum Due Amount >>", 40)
    + basText.replicats(" ", 18) + basText.alignmentRight((object)basText.formatNumUnSign(accRowView[mMindueamount], "##0.00", 16), 16) + " " + CrDb(Convert.ToDecimal(accRowView[mOpeningbalance])));
    printDetailCompany("  " + basText.replicats(" ", 5) + "  " + basText.replicats(" ", 5) + "  "
    + basText.replicats(" ", 24) + "  " + basText.alignmentLeft((object)"Total Overdue Amount >>", 40)
    + basText.replicats(" ", 18) + basText.alignmentRight((object)basText.formatNumUnSign(accRowView[mTotaloverdueamount], "##0.00", 16), 16) + " " + CrDb(Convert.ToDecimal(accRowView[mOpeningbalance])));
  }

  //* Card Holder statement
  protected void printHeaderCardHolder()
  {
    string newaddress1, newaddress2;
    clsMaintainData.fixAddress(curBranchVal, masterRow[mCustomeraddress1].ToString(), out newaddress1, out newaddress2);
    totPages++; endOfCustomer = "";
    extAccNum = clsBasValid.validateStr(masterRow[mExternalno]);
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
      endOfCustomer = "/";
    }

    streamWrit.WriteLine(strEndOfPage);
    streamWrit.WriteLine(basText.replicat(" ", 81) + CurrentPageFlag);  //+"{0,5:MM/yy}",masterRow[mStatementdatefrom]
    streamWrit.WriteLine(basText.replicat(" ", 81) + masterRow[mCardproduct]);  //+"{0,5:MM/yy}",masterRow[mStatementdatefrom]
    //			streamWrit.WriteLine(basText.alignmentLeft(masterRow[mCustomername],50)+ basText.replicat(" ",31) + masterRow[mCardbranchpartname]);  //
    streamWrit.WriteLine(basText.alignmentMiddle(masterRow[mCustomername], 50));  //
    streamWrit.WriteLine(basText.replicat(" ", 81) + masterRow[mCardbranchpartname]);  //
    //			streamWrit.WriteLine(basText.alignmentMiddle(masterRow[mCustomeraddress1],50)+ basText.replicat(" ",31) + clsBasValid.validateStr(masterRow[accountNoName])); //basText.formatNum(masterRow[mCardno],"####-####-####-####")
    //string x = Convert.ToString(ValidateArbic(masterRow[mCustomeraddress1].ToString()));
    //streamWrit.WriteLine(basText.alignmentMiddle(ValidateArbic(masterRow[mCustomeraddress1].ToString()), 50)); //basText.formatNum(masterRow[mCardno],"####-####-####-####")
    streamWrit.WriteLine(basText.alignmentMiddle(ValidateArbic(newaddress1), 50)); //basText.formatNum(masterRow[mCardno],"####-####-####-####")
    streamWrit.WriteLine(basText.replicat(" ", 81) + extAccNum); //clsBasValid.validateStr(masterRow[mExternalno]) //accountno  //basText.formatNum(masterRow[mCardno],"####-####-####-####")
    //			streamWrit.WriteLine(basText.alignmentMiddle(masterRow[mCustomeraddress2],50)+ basText.replicat(" ",31)+"{0,10:dd/MM/yyyy}",masterRow[mStatementdatefrom]);//+ basText.formatNum(masterRow[accountLimit],"########")  basText.formatNum(masterRow[mCardlimit],"########")
    //streamWrit.WriteLine(basText.alignmentMiddle(ValidateArbic(masterRow[mCustomeraddress2].ToString()), 50));
    streamWrit.WriteLine(basText.alignmentMiddle(ValidateArbic(newaddress2), 50));
    streamWrit.WriteLine(basText.replicat(" ", 81) + "{0,10:dd/MM/yyyy}", masterRow[mStatementdateto]);//Statementdatefrom+ basText.formatNum(masterRow[accountLimit],"########")  basText.formatNum(masterRow[mCardlimit],"########")
    //			streamWrit.WriteLine(basText.alignmentMiddle(masterRow[mCustomeraddress3].ToString().Trim() + " " + masterRow[mCustomerregion].ToString().Trim() + " " + masterRow[mCustomercity].ToString().Trim(),50)+ basText.replicat(" ",31)+pageNo + " / " + totalAccPages);  //basText.formatNum(masterRow[accountNoName],"##-##-######-####")
    streamWrit.WriteLine(basText.alignmentMiddle(ValidateArbic(masterRow[mCustomeraddress3].ToString().Trim()) + " " + masterRow[mCustomerregion].ToString().Trim() + " " + masterRow[mCustomercity].ToString().Trim(), 50));  //basText.formatNum(masterRow[accountNoName],"##-##-######-####")
    streamWrit.WriteLine(basText.replicat(" ", 81) + pageNo + " / " + totalAccPages);  //basText.formatNum(masterRow[accountNoName],"##-##-######-####")
    if (createCorporateVal)
      streamWrit.WriteLine(basText.alignmentMiddle(masterRow[mCardclientname], 50));
    else
      streamWrit.WriteLine(String.Empty);

    streamWrit.WriteLine(String.Empty);
    //streamWrit.WriteLine(omr.Asterisk(pageNo));//String.Empty
    streamWrit.WriteLine(omr.Asterisk(totPages) + omr.AsteriskPage4(pageNo));//String.Empty
    streamWrit.WriteLine(String.Empty);
    //streamWrit.WriteLine(basText.alignmentMiddle(basText.formatCardNumber(curMainCard), 20) + basText.alignmentMiddle(masterRow[accountLimit], 13) + basText.formatNum(masterRow[accountAvailableLimit], "##0", 20) + basText.alignmentMiddle(basText.formatNumUnSign(masterRow[mMindueamount], "0.00", 13), 13) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M") + basText.formatNum(masterRow[mTotaloverdueamount], "##0.00", 13)); //+ basText.formatNum(masterRow[accountLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mTotalcredits],"#,##0.00",9)+" "+basText.formatNum(masterRow[accountAvailableLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mMindueamount],"#,##0.00",15)+" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",15)+" "+basText.formatDate(masterRow[mStetementduedate],"dd/MM/yy") //TOTALCREDITS// basText.alignmentMiddle(masterRow[accountAvailableLimit],20) //"#,##0"
    streamWrit.WriteLine(basText.alignmentMiddle(basText.formatCardNumber(curMainCard), 20) + basText.alignmentMiddle(masterRow[mAccountlim], 13) + basText.formatNum(masterRow[accountAvailableLimit], "##0", 20) + basText.alignmentMiddle(basText.formatNumUnSign(masterRow[mMindueamount], "0.00", 13), 13) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M") + basText.formatNum(masterRow[mTotaloverdueamount], "##0.00", 13)); //+ basText.formatNum(masterRow[accountLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mTotalcredits],"#,##0.00",9)+" "+basText.formatNum(masterRow[accountAvailableLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mMindueamount],"#,##0.00",15)+" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",15)+" "+basText.formatDate(masterRow[mStetementduedate],"dd/MM/yy") //TOTALCREDITS// basText.alignmentMiddle(masterRow[accountAvailableLimit],20) //"#,##0"
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    if (pageNo == 1)
      streamWrit.WriteLine(basText.replicat(" ", 43) + basText.alignmentLeft("Previous Balance", 63) + basText.formatNumUnSign(masterRow[mOpeningbalance], "##0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance])));  //+ CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign]))
    else
      streamWrit.WriteLine(String.Empty);

    streamWrit.WriteLine(String.Empty);
    totalPages++;
    totNoOfPageStat++;
  }


  protected void printDetailCardHolder()
  {
    DateTime trnsDate = clsBasValid.validateDate(detailRow[dTransdate]), postingDate = clsBasValid.validateDate(detailRow[dPostingdate]);
    if (trnsDate > postingDate)
      trnsDate = postingDate;

    if (masterRow[mAccountcurrency].ToString().Trim() != detailRow[dOrigtrancurrency].ToString().Trim())
      strForeignCurr = basText.formatNumUnSign(detailRow[dOrigtranamount], "##,###,##0.00;(#,###,##0.00)", 13) + " " + basText.Replace(clsBasValid.validateStr(detailRow[dOrigtrancurrency]), "XXX", "   ");
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
    streamWrit.WriteLine("  {0:dd/MM}  {1:dd/MM}  {2,-24}  {3,-40} {4,16} {5,16} {6,2}", trnsDate, postingDate, basText.trimStr(detailRow[dRefereneno].ToString().Trim(), 24), basText.trimStr(trnsDesc, 40), strForeignCurr, basText.formatNumUnSign(detailRow[dBilltranamount], "#,###,##0.00", 14), CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + " " + isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)
    totNoOfTransactions++;
  }

  protected void printCardFooterCardHolder()
  {
    //completePageDetailRecords();
    streamWrit.WriteLine(String.Empty);
    if (pageNo == totalAccPages)
      streamWrit.WriteLine(basText.replicat(" ", 17) + basText.alignmentLeft("Current Balance", 63) + basText.formatNumUnSign(masterRow[mClosingbalance], "#0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])));
    //streamWrit.WriteLine(basText.replicat(" ", 80) + basText.formatNumUnSign(masterRow[mClosingbalance], "#0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])));
    else
      streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    //streamWrit.WriteLine(String.Empty);
    //Reward data
    rewardRows = DSreward.Tables["Reward"].Select("CLIENTID = '" + clsBasValid.validateStr(masterRow[mClientid]).ToString().Trim() + "'");
    if (rewardRows.Length > 0)
    {
      rewardRow = rewardRows[0];
      streamWrit.WriteLine(basText.formatNum(rewardRow[mOpeningbalance], "#0.00;(#0.00)", 20, "L") + basText.formatNum(rewardRow["EARNEDBONUS"], "#0.00;(#0.00)", 20, "L") + basText.formatNum(rewardRow["REDEEMEDBONUS"], "#0.00;(#0.00)", 20, "L") + basText.formatNum(rewardRow[mClosingbalance], "#0.00;(#0.00)", 20, "L"));//streamWrit.WriteLine(String.Empty);
    }
    else
      streamWrit.WriteLine(basText.replicats(" ", 80));
    streamWrit.WriteLine(basText.formatNum(masterRow[mTotaldebits], "#0.00;(#0.00)", 20, "L") + basText.formatNum(masterRow[mTotalcredits], "#0.00;(#0.00)", 20, "L"));//streamWrit.WriteLine(String.Empty);
    //streamWrit.WriteLine(basText.alignmentRight(basText.formatCardNumber(curMainCard), 35) + basText.alignmentRight(masterRow[mMindueamount], 20) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "R")//
    streamWrit.WriteLine(basText.alignmentRight(basText.formatCardNumber(curMainCard), 35) + basText.alignmentRight(basText.formatNumUnSign(masterRow[mMindueamount], "0.00", 20), 20) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "R")//
   + basText.alignmentRight(basText.formatNumUnSign(masterRow[mOpeningbalance], "#,###,##0.00", 12) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance].ToString())), 15)
   + basText.alignmentRight(basText.formatNumUnSign(masterRow[mTotalpayments], "#,###,##0.00", 12) + " " + DbCr(Convert.ToDecimal(masterRow[mTotalpayments])), 15)
   + basText.alignmentRight(basText.formatNum(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal]), "#,###,##0.00", 12) + CrDbMinus(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal])), 15)
   + basText.alignmentRight(basText.formatNumUnSign(masterRow[mTotalcharges].ToString(), "#,###,##0.00", 12) + " " + CrDbMinus(Convert.ToDecimal(masterRow[mTotalcharges])), 15)
   + basText.alignmentRight(basText.formatNumUnSign(masterRow[mTotalinterest].ToString(), "#,###,##0.00", 12) + " " + CrDbMinus(Convert.ToDecimal(masterRow[mTotalinterest])), 15)
      //       + basText.formatNum(Convert.ToDecimal(masterRow[mTotalcharges]) + Convert.ToDecimal(masterRow[mTotalinterest]),"#,##0.00",12) + CrDbMinus(Convert.ToDecimal(masterRow[mTotalcharges]) + Convert.ToDecimal(masterRow[mTotalinterest])) 

   + basText.alignmentRight(basText.formatNumUnSign(masterRow[mClosingbalance], "#,###,##0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])), 15));
  }

  //*
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
    stmNo = masterRow[mCardaccountno].ToString();

    accountRows = DSstatement.Tables["tStatementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[mCardaccountno]).ToString().Trim() + "'");
    totalAccPages = 0;
    totAccRows = 0;
    string prevCardNo = String.Empty, CurCardNo = String.Empty;
    int currAccRowsPages = 0;
    foreach (DataRow dtAccRow in accountRows) //mRow.GetChildRows(StatementNoDRel)
    {
      if ((dtAccRow[dPostingdate] == DBNull.Value) && (dtAccRow[dDocno] == DBNull.Value)) continue;
      CurCardNo = dtAccRow[dCardno].ToString();
      if (CurCardNo.Trim().Length < 1) 
        continue;

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
    mainRows = DSstatement.Tables["tStatementMasterTable"].Select(mCardaccountno + " = '" + clsBasValid.validateStr(masterRow[mCardaccountno]) + "'");//ACCOUNTNO
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


  private void calcContractRows()//string pContractNo, string pCurr
  {
    curAccContractNo = prvAccContractNo = string.Empty;
    noOfCrdInAcc = noOfCrdInContract = 0;
    compTotalOverdueAmount = compTotaldebits = compTotalcredits = compOpeningbalance = compTotalpayments = compTotalpurchases = compTotalcashwithdrawal = compTotalcharges = compTotalinterest = compClosingbalance = 0;
    curCrdNo = prvCrdNo = string.Empty;
    noOfTrnsInContract = noOfTrnsInAcc = curNoOfAccInContract = noOfAccInContract = curPageCont = totPageCont = 0;
    if (contractRowsView == null)
      return;
    foreach (DataRowView cntRow in contractRowsView) //account scope
    {
      curAccContractNo = clsBasValid.validateStr(cntRow[mCardaccountno]).ToString().Trim();
      if (curAccContractNo == prvAccContractNo)
        continue;
      
      prvAccContractNo = curAccContractNo;
      noOfCrdInAcc = noOfTrnsInAcc = 0;
      compTotalOverdueAmount += Convert.ToDecimal(cntRow[mTotaloverdueamount].ToString());
      compMindueamount += Convert.ToDecimal(cntRow[mMindueamount].ToString());
      compTotaldebits += Convert.ToDecimal(cntRow[mTotaldebits].ToString());
      compTotalcredits += Convert.ToDecimal(cntRow[mTotalcredits].ToString());
      compOpeningbalance += Convert.ToDecimal(cntRow[mOpeningbalance].ToString());
      compTotalpayments += Convert.ToDecimal(cntRow[mTotalpayments].ToString());
      compTotalpurchases += Convert.ToDecimal(cntRow[mTotalpurchases].ToString());
      compTotalcashwithdrawal += Convert.ToDecimal(cntRow[mTotalcashwithdrawal].ToString());
      compTotalcharges += Convert.ToDecimal(cntRow[mTotalcharges].ToString());
      compTotalinterest += Convert.ToDecimal(cntRow[mTotalinterest].ToString());
      compClosingbalance += Convert.ToDecimal(cntRow[mClosingbalance].ToString());
      transRows = DSstatement.Tables["tStatementDetailTable"].Select(dAccountno + " = '" + curAccContractNo + "'");
      foreach (DataRow transRow in transRows)
      {
        if ((transRow[dPostingdate] == DBNull.Value) && (transRow[dDocno] == DBNull.Value))
          continue;
        curCrdNo = transRow[dCardno].ToString().Trim();
        if (curCrdNo != prvCrdNo)
        {
          noOfCrdInContract++;
          noOfCrdInAcc++;
        }
        prvCrdNo = curCrdNo;
        noOfTrnsInAcc++;
      }
      noOfTrnsInContract += noOfTrnsInAcc;
      if (noOfTrnsInAcc > 0 || Convert.ToDecimal(cntRow[mClosingbalance].ToString()) != 0)
      {
        noOfAccInContract++;
        if (noOfTrnsInAcc > 0)
        {
          totPageCont = noOfTrnsInAcc / MaxDetailInPage;
          totPageCont = (noOfTrnsInAcc % MaxDetailInPage) > 0 ? ++totPageCont : totPageCont;
        }
        else if (Convert.ToDecimal(cntRow[mClosingbalance].ToString()) != 0)
        {
          totPageCont++;
        }
        //[noOfTrnsInContract += (noOfCrdInAcc * 5);
      }
      //[else
      //[  if (Convert.ToDecimal(cntRow[mClosingbalance].ToString()) != 0)
      //[    noOfTrnsInContract += 5;
    }
    //[noOfTrnsInContract += (noOfCrdInContract*5);
    //[totPageCont = noOfTrnsInContract / MaxDetailInPage;
    //[totPageCont = (noOfTrnsInContract % MaxDetailInPage)>0 ? ++totPageCont : totPageCont;
    prvContractNo = string.Empty;
    curCrdNo = prvCrdNo = string.Empty;
    noOfCrdInAcc = 0;
  }

  private void calcAccountRows(DataView pAccountRowsView)
  {
    //string curCrdAccNo = string.Empty, curCrdAccNoPrev = string.Empty;
    curMainCard = CurCardNo = vStatementno = string.Empty;
    curNoOfCrdInAcc = noOfCrdInAcc = curNoOfTrnsInAcc = noOfTrnsInAcc = 0;
    decimal accClosingbalance = 0;

    if (pAccountRowsView == null)
      return;
    foreach (DataRowView accRow in pAccountRowsView)
    {
      noOfCrdInAcc++;
      CurCardNo = accRow[mCardno].ToString();
      accClosingbalance = Convert.ToDecimal(accRow[mClosingbalance].ToString());
      if (accRow[mCardprimary].ToString() == "Y")
        curMainCard = CurCardNo;
      if (accRow[mCardprimary].ToString() == "Y" && isValidateCard(accRow[mCardstate].ToString()))
      {
        curMainCard = CurCardNo; 
      }

      vStatementno = clsBasValid.validateStr(accRow[mStatementno]).ToString().Trim();
      transRows = DSstatement.Tables["tStatementDetailTable"].Select(dStatementno + " = '" + vStatementno + "'");
      foreach (DataRow transRow in transRows)
      {
        if ((transRow[dPostingdate] == DBNull.Value) && (transRow[dDocno] == DBNull.Value))
          continue;
        noOfTrnsInAcc++;
      }//transRows
    }//pAccountRowsView
    if (curMainCard == "")
      curMainCard = CurCardNo;

    if (noOfTrnsInAcc > 0)
    {
      totPageCont = noOfTrnsInAcc / MaxDetailInPage;
      totPageCont = (noOfTrnsInAcc % MaxDetailInPage) > 0 ? ++totPageCont : totPageCont;
    }
    else if (accClosingbalance != 0)
    {
      totPageCont++;
    }

  }


  private void completePageDetailRecords()
  {
    //int curPageLine =CurPageRec4Dtl;
    //if (curPageCont == totalAccPages)
    for (int curPageLine = CurPageRec4Dtl; curPageLine < MaxDetailInPage; curPageLine++)
      streamWrit.WriteLine(String.Empty);
  }


  private void printStatementSummary()
  {
    streamSummary.WriteLine(strBankName + " Visa Statement");
    streamSummary.WriteLine("__________________________");
    streamSummary.WriteLine("");
    streamSummary.WriteLine("No of Statements   " + totNoOfCardStat.ToString());
    streamSummary.WriteLine("No of Pages        " + totNoOfPageStat.ToString());
    streamSummary.WriteLine("No of Transactions " + totNoOfTransactions.ToString());
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


  ~clsStatement_CommonCompanyDtl()
  {
    DSstatement.Dispose();
  }
}
