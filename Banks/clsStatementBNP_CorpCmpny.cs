using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Xml;
using System.Collections;

// Branch 1
public class clsStatementBNP_CorpCmpny : clsBasStatement
{
  private string strBankName;
  private FileStream fileStrm, fileSummary, fileRawData;
  private StreamWriter streamWrit, streamSummary, strmRawData;
  //private DataSet DSstatement;
  //private OracleDataReader drPrimaryCards, drMaster,drDetail;
  private DataRow masterRow;
  private DataRow detailRow;
  private string strEndOfLine = "\u000D";  //+ "M" ^M
  //		private string strEndOfPage = "\u000C"  ;  //+ "\u000D"+ "M" ^L + ^M
  private string strEndOfPage = "\u000C";  //  + basText.replicat(" ",85) + "F 2"  + "\u000D"+ "M" ^L + ^M
  private const int MaxDetailInPage = 20; //
  private const int MaxDetailInPageComp = 25; //
  private const int MaxDetailInLastPage = 27; //
  private int CurPageRec4Dtl = 0;
  private int pageNo = 0, totalPages = 0, totalCardPages = 0, curAccCompanyRows = 0
    , totalAccPages = 0, totCardRows = 0, totAccRows = 0, curAccRows = 0;//
  private string lastPageTotal;
  private string curCardNo;//,PrevCardNo
  private string curAccountNo, prevAccountNo = String.Empty;//,PrevCardNo
  private decimal totNetUsage = 0;
  private DataRow[] cardsRows, accountRows, dtlRows;
  private DataRow[] mainRows;
  private string CurrentPageFlag;
  private string strCardNo, strPrimaryCardNo, strCurrency, strStatementNo;
  private string strForeignCurr;
  private string stmNo;
  private int totNoOfCardStatCrdHldr, totNoOfPageStatCrdHldr;
  private int totNoOfCardStatCmpny, totNoOfPageStatCmpny;
  private string strFileName;
  private decimal valOutstanding = 0, valCompOutstanding = 0;
  private decimal valTransaction = 0, valCompClosingbalance = 0, valCompMinDueAmount = 0, compTotalOverdueAmount = 0;
  private string curContractNo;


  private int curCardRow = 0, curTotCardRows = 0, numOfErr = 0, totNoOfTransactions = 0;
  private bool isPrimaryOnly, isHaveF3 = true, isPrimaryOpened = false;
  private FileStream fileStrmErr;
  private StreamWriter strmWriteErr;
  private string curMainCard, CurCardAddres1, CurCardAddres2, CurCardAddres3, CurCardNo;
  private int totCrdNoInAcc, curCrdNoInAcc;
  private string strOutputPath, strFileCompany, strFileCardHolder
    , fileSummaryName;
  private DateTime vCurDate;
  private ArrayList aryLstFiles = new ArrayList();
  private string curContarctCurr;
  private frmStatementFile frmMain;
  private int totContact;
  private clsOMR omr = new clsOMR();
  private int totPages;
  private string rewardCondStr = "'Reward Program'";//'New Reward Contract''Reward Contract'


  public clsStatementBNP_CorpCmpny()
  {
  }

  public string Statement(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pCurDate, string pStmntType, bool pAppendData)
  {
    string rtrnStr = "Successfully Generate " + pBankName;
    int curMonth = pCurDate.Month;
    bool preExit = true;
    try
    {
      clsMaintainData maintainData = new clsMaintainData();
      maintainData.matchCardBranch4Account(pBankCode);

      pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);
      vCurDate = pCurDate; //DateTime.Now.AddMonths(-1);
      strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
      clsBasFile.deleteDirectory(strOutputPath);
      clsBasFile.createDirectory(strOutputPath);
      pStrFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + "_Company" + pStrFile + vCurDate.ToString("yyyyMM") + ".txt";
      strBankName = pBankName;

      //>strFileName = clsBasFile.getPathWithoutExtn(pStrFileName) +
      //>  "_EGP." + clsBasFile.getFileExtn(pStrFileName);
      strFileName = pStrFileName;//System.IO.File.Delete(
      //strFileName = pStrFileName;
      isCorporateVal = true;//clsBasStatement
      isCredite = false;
      // set branch for data
      curBranchVal = pBankCode; // 1; //3 = real   1 = test
      // set branch for data
      //      mainTableCond =  " cardproduct ='" + "Visa Business" + "' "; 
      //>mainTableCond = " substr(cardno,1,6) ='421192' "; //" cardproduct != 'Visa Business' ";
      //>supTableCond = " substr(cardno,1,6) ='421192' ";
      // set Currency for data
      //>curCurrencyVal = "-818-B";
      // data retrieve
      FillStatementDataSet(pBankCode); //DSstatement =  //1); //3

      printCompanyStatement();

      /*      // other Currency
            strFileName = clsBasFile.getPathWithoutExtn(pStrFileName)+
              "_USD." + clsBasFile.getFileExtn(pStrFileName);
            clsBasStatement.curCurrencyVal = "-840-B"; 
            // data retrieve
            DSstatement = FillStatementDataSet(1); //3

            printCardHolderStatement();
            printCompanyStatement(); */
      //fillStatementHistory(pStmntType,pAppendData);
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

        // open Summary file
        fileSummaryName = pStrFileName;
        fileSummaryName = clsBasFile.getPathWithoutExtn(fileSummaryName) +
          "_Company_Summary." + clsBasFile.getFileExtn(fileSummaryName);
        fileSummary = new FileStream(fileSummaryName, FileMode.Create); //Create
        streamSummary = new StreamWriter(fileSummary, Encoding.Default);

        printStatementSummary();

        // Close Summary File
        streamSummary.Flush();
        streamSummary.Close();
        fileSummary.Close();

        //ArrayList aryLstFiles = new ArrayList();
        //aryLstFiles.Add("");
        //>aryLstFiles.Add(@strFileCardHolder);
        //>aryLstFiles.Add(clsBasFile.getPathWithoutExtn(strFileCardHolder) + "_Card Number.txt");
        //>aryLstFiles.Add(@strFileCompany);
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


  public ArrayList Statement(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pCurDate)
  {
    string rtrnStr = "Successfully Generate " + pBankName;
    int curMonth = pCurDate.Month;
    bool preExit = true;
    try
    {
        clsMaintainData maintainData = new clsMaintainData();
        maintainData.matchCardBranch4Account(pBankCode);

      pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);
      vCurDate = pCurDate; //DateTime.Now.AddMonths(-1);
      strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
      //clsBasFile.deleteDirectory(strOutputPath);
      //clsBasFile.createDirectory(strOutputPath);
      pStrFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + "_Company" + pStrFile + vCurDate.ToString("yyyyMM") + ".txt";
      strBankName = pBankName;

      strFileName = pStrFileName;//System.IO.File.Delete(
      isCorporateVal = true;//clsBasStatement
      isCredite = false;
      curBranchVal = pBankCode;
      strMainTableCond = "m.contracttype != " + rewardCondStr;//'Reward Program (Airmile)'
      strSubTableCond = "d.trandescription != 'Calculated Points'";

      FillStatementDataSet(pBankCode); //DSstatement =  //1); //3

      printCompanyStatement();
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
        // open Summary file
        fileSummaryName = pStrFileName;
        fileSummaryName = clsBasFile.getPathWithoutExtn(fileSummaryName) +
          "_Summary." + clsBasFile.getFileExtn(fileSummaryName);
        fileSummary = new FileStream(fileSummaryName, FileMode.Create); //Create
        streamSummary = new StreamWriter(fileSummary, Encoding.Default);

        printStatementSummary();

        // Close Summary File
        streamSummary.Flush();
        streamSummary.Close();
        fileSummary.Close();

        aryLstFiles.Add(@fileSummaryName);
        DSstatement.Dispose();
      }
    }
    return aryLstFiles;
  }


  private void printCompanyStatement()
  {
    totPages = 0;
    strFileCompany = strFileName;
    bool isCompleteLine = true;

    try
    {
      //strFileCompany = clsBasFile.getPathWithoutExtn(strFileCompany) +
      //  "_Company." + clsBasFile.getFileExtn(strFileCompany);

      // open output file
      //>fileStrm = new FileStream(strFileCompany, FileMode.Create); //Create
      //>streamWrit = new StreamWriter(fileStrm, Encoding.Default);

      // Error file
      fileStrmErr = new FileStream(clsBasFile.getPathWithoutExtn(strFileName) + "_Err." + clsBasFile.getFileExtn(strFileName), FileMode.Create);
      strmWriteErr = new StreamWriter(fileStrmErr, Encoding.Default);


      //int totCnrct = 0;
      pageNo = 0; totalCardPages = 0; totContact = 1;
      curCardNo = String.Empty;
      curAccountNo = prevAccountNo = String.Empty;
      curContarctCurr = "";
      frmMain.BeginInvoke(frmMain.setMinMaxProgressDelegate, new object[] { DSstatement.Tables["contractTable"].Rows.Count });
      foreach (DataRow contractRow in DSstatement.Tables["contractTable"].Rows)
      {
        frmMain.BeginInvoke(frmMain.setProgressDelegate, new object[] { totContact++ });
        curContractNo = contractRow[mContractno].ToString();
        cardsRows = DSstatement.Tables["tStatementMasterTable"].Select("contractno = '" + curContractNo + "'"); //accountno
        pageNo = 1; curAccRows = CurPageRec4Dtl = 0; CurrentPageFlag = "F 1"; //if page is based on account no
        valCompOutstanding = valCompClosingbalance = valCompMinDueAmount = 0;
        //masterRow = contractRow;
        calcAccountRowsCompany();//>>
        //|>totNoOfCardStatCmpny++;
        //check product

        foreach (DataRow cardRow in cardsRows)
        {
          masterRow = cardRow;
          if (curContarctCurr != masterRow[mAccountcurrency].ToString().Trim())
          {
            if (curContarctCurr != string.Empty)
            {
              streamWrit.Flush();
              streamWrit.Close();
              fileStrm.Close();
            }
            curContarctCurr = masterRow[mAccountcurrency].ToString().Trim();
            strFileCompany = clsBasFile.getPathWithoutExtn(strFileName)
              + "_" + curContarctCurr + "." + clsBasFile.getFileExtn(strFileName);
            //curFileName = pStrFileName;
            add2FileList(strFileCompany);
            fileStrm = new FileStream(strFileCompany, FileMode.Append); //FileMode.Create Create
            streamWrit = new StreamWriter(fileStrm, Encoding.Default);
            streamWrit.AutoFlush = true;
          }

          stmNo = masterRow[mContractno].ToString();
          stmNo = masterRow[mStatementno].ToString();
          //          strStatementNo = masterRow[mStatementno].ToString();
          strStatementNo = masterRow[mCardaccountno].ToString();

          calcAccountRowsCardHolderCompany();

          //         if(prevAccountNo == masterRow[mCardaccountno].ToString())
          //           continue;
          if (curMainCard != masterRow[mCardno].ToString())
            isCompleteLine = false; //continue;

          //if(curAccCompanyRows < 1 && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0) //Convert.ToDecimal(
          if (Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0) //Convert.ToDecimal(
            if (curAccCompanyRows < 1)
              isCompleteLine = false;

          //          if(masterRow[mCardstate].ToString() != "Given" && masterRow[mCardstate].ToString() != "New"  && masterRow[mCardstate].ToString() != "New Pin Generated Only" ) //  && masterRow[mCardstate].ToString() != "Embossed"
          //          if(masterRow[mAccountstatus].ToString() != "Open")
          //            isCompleteLine = false; //continue;

          if (isCompleteLine)
          {
            curAccRows++;
            if (curAccRows == 1) //>>      == 1)
              printHeaderCompany();//if page is based on account no

            if (CurPageRec4Dtl >= MaxDetailInPageComp)
            {
              CurPageRec4Dtl = 0;
              printCardFooterCompany();
              pageNo++;
              printHeaderCompany();
            }
            //totNetUsage = calculateCrDb(totNetUsage, clsBasValid.validateNum(detailRow[dBilltranamount]), clsBasValid.validateStr(detailRow[dBilltranamountsign]));
            CurPageRec4Dtl = CurPageRec4Dtl + 1;
            printDetailCompany();
          } //end of isCompleteLine 

          curCrdNoInAcc++;
          if (curAccRows >= totAccRows && CurrentPageFlag != string.Empty && totAccRows != 0) //curAccRows >= totAccRows
          //if ((curAccRows >= totAccRows && totAccRows != 0 && CurrentPageFlag != string.Empty) || (totAccRows == 0 && curCrdNoInAcc == totCrdNoInAcc))
          {
            completePageDetailRecordsCompany();
            printCardFooterCompany();//if pages is based on account
            pageNo = 1; CurPageRec4Dtl = 0; //if pages is based on account
            curAccRows = 0;
          }
          totNetUsage = 0;
          isCompleteLine = true;
          prevAccountNo = masterRow[mCardaccountno].ToString();
          //CurrentPageFlag = string.Empty;
        } //end of Master foreach
        CurrentPageFlag = string.Empty;
        prevAccountNo = masterRow[mContractno].ToString(); //>masterRow[mCardaccountno].ToString()
      }//end of Contract foreach
    }//end of try
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
      // Close output File
      streamWrit.Flush();
      streamWrit.Close();
      fileStrm.Close();

      strmWriteErr.Flush();
      strmWriteErr.Close();
      fileStrmErr.Close();

    }
  }



  private void printHeaderCompany()
  {
    string newaddress1, newaddress2;
    clsMaintainData.fixAddress(curBranchVal, masterRow[mCustomeraddress1].ToString(), out newaddress1, out newaddress2);
    totPages++;
    if (pageNo == 1 && totalAccPages == 1)
    {
      CurrentPageFlag = "F 0";
      totNoOfCardStatCmpny++;
    }
    else if (pageNo == 1 && totalAccPages > 1)
    {
      CurrentPageFlag = "F 1";
      totNoOfCardStatCmpny++;
    }
    else if (pageNo < totalAccPages)
      CurrentPageFlag = "F 2";
    else if (pageNo == totalAccPages)
      CurrentPageFlag = "F 3";

    streamWrit.WriteLine(strEndOfPage);
    streamWrit.WriteLine(basText.replicat(" ", 81) + CurrentPageFlag);  //+"{0,5:MM/yy}",masterRow[mStatementdatefrom]
    streamWrit.WriteLine(basText.replicat(" ", 81) + masterRow[mCardproduct]);  //+"{0,5:MM/yy}",masterRow[mStatementdatefrom]
    streamWrit.WriteLine(basText.alignmentMiddle(masterRow[mCustomername], 50));  //
    streamWrit.WriteLine(basText.replicat(" ", 81) + masterRow[mCardbranchpartname]);  //
    //streamWrit.WriteLine(basText.alignmentMiddle(ValidateArbic(masterRow[mCustomeraddress1].ToString()), 50)); //basText.formatNum(masterRow[mCardno],"####-####-####-####")
    streamWrit.WriteLine(basText.alignmentMiddle(ValidateArbic(newaddress1), 50)); //basText.formatNum(masterRow[mCardno],"####-####-####-####")
    //    streamWrit.WriteLine( basText.replicat(" ",81) + clsBasValid.validateStr(masterRow[mExternalno])); //accountno  //basText.formatNum(masterRow[mCardno],"####-####-####-####")
    streamWrit.WriteLine(basText.replicat(" ", 81) + "{0,10:dd/MM/yyyy}", masterRow[mStatementdateto]);//Statementdatefrom+ basText.formatNum(masterRow[mAccountlim],"########")  basText.formatNum(masterRow[mCardlimit],"########")
    //streamWrit.WriteLine(basText.alignmentMiddle(ValidateArbic(masterRow[mCustomeraddress2].ToString()), 50));
    streamWrit.WriteLine(basText.alignmentMiddle(ValidateArbic(newaddress2), 50));
    streamWrit.WriteLine(basText.replicat(" ", 81) + pageNo + " / " + totalAccPages);  //basText.formatNum(masterRow[mAccountno],"##-##-######-####")
    streamWrit.WriteLine(basText.alignmentMiddle(ValidateArbic(masterRow[mCustomeraddress3].ToString().Trim() + " " + masterRow[mCustomerregion].ToString().Trim() + " " + masterRow[mCustomercity].ToString().Trim()), 50));  //basText.formatNum(masterRow[mAccountno],"##-##-######-####")
    streamWrit.WriteLine(basText.alignmentMiddle(ValidateArbic(masterRow[mContactpersonename].ToString()), 50));//String.Empty
    streamWrit.WriteLine(omr.Asterisk(totPages) + omr.AsteriskPage4(pageNo));//String.Empty
    //    streamWrit.WriteLine(basText.alignmentMiddle(masterRow[mContractno],20) + basText.alignmentMiddle(masterRow[mContractlimit],13) + basText.formatNum(masterRow[mAccountavailablelim],"##0",20) + basText.formatDate(masterRow[mStetementduedate],"dd/MM/yyyy",20,"M")) ; // + basText.formatNum(masterRow[mTotaloverdueamount],"##0.00",13)    //+ basText.formatNum(masterRow[mAccountlim],"#,##0.00",9)+" "+basText.formatNum(masterRow[mTotalcredits],"#,##0.00",9)+" "+basText.formatNum(masterRow[mAccountavailablelim],"#,##0.00",9)+" "+basText.formatNum(masterRow[mMindueamount],"#,##0.00",15)+" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",15)+" "+basText.formatDate(masterRow[mStetementduedate],"dd/MM/yy") //TOTALCREDITS// basText.alignmentMiddle(masterRow[mAccountavailablelim],20) //"#,##0"  +  basText.alignmentMiddle(valCompMinDueAmount,13)
    streamWrit.WriteLine(String.Empty);//
    streamWrit.WriteLine(basText.alignmentMiddle(masterRow[mExternalno], 20) + basText.alignmentMiddle(masterRow[mContractlimit], 13) + basText.formatNum(masterRow[mAccountavailablelim], "##0", 20) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 20, "M") + basText.formatNum(compTotalOverdueAmount, "#,##0.00", 15)); //masterRow[mTotaloverdueamount]// + basText.formatNum(masterRow[mTotaloverdueamount],"##0.00",13)    //+ basText.formatNum(masterRow[mAccountlim],"#,##0.00",9)+" "+basText.formatNum(masterRow[mTotalcredits],"#,##0.00",9)+" "+basText.formatNum(masterRow[mAccountavailablelim],"#,##0.00",9)+" "+basText.formatNum(masterRow[mMindueamount],"#,##0.00",15)+" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",15)+" "+basText.formatDate(masterRow[mStetementduedate],"dd/MM/yy") //TOTALCREDITS// basText.alignmentMiddle(masterRow[mAccountavailablelim],20) //"#,##0"  +  basText.alignmentMiddle(valCompMinDueAmount,13)
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    //    if(pageNo == 1)
    //      streamWrit.WriteLine(basText.replicat(" ",17) + basText.alignmentLeft("Previous Balance",63) + basText.formatNumUnSign(masterRow[mOpeningbalance],"##0.00",16)+ " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance])));  //+ CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign]))
    //    else
    //      streamWrit.WriteLine(String.Empty);

    streamWrit.WriteLine(String.Empty);
    totalPages++;
    totNoOfPageStatCmpny++;
  }



  private void printDetailCompany()
  {
    valOutstanding = valTransaction = 0;
    dtlRows = DSstatement.Tables["tStatementDetailTable"].Select("accountno = '" + strStatementNo + "'"); //accountno
    foreach (DataRow dtlRow in dtlRows)
    {
      if ((dtlRow[dPostingdate] == DBNull.Value) && (dtlRow[dDocno] == DBNull.Value))
        valOutstanding += Convert.ToDecimal(dtlRow[dBilltranamount].ToString());
      else
        valTransaction += Convert.ToDecimal(dtlRow[dBilltranamount].ToString());
    }
    valCompOutstanding += valOutstanding;
    if (Convert.ToDecimal(masterRow[mClosingbalance].ToString()) < 0)
      valCompClosingbalance += Convert.ToDecimal(masterRow[mClosingbalance].ToString()); //valTransaction;
    valCompMinDueAmount += Convert.ToDecimal(masterRow[mMindueamount].ToString()); //valTransaction;

    streamWrit.WriteLine("  {0,16}  {1,-50} {2,16}", masterRow[mCardno].ToString(), masterRow[mCardclientname].ToString(), basText.formatNumUnSign(masterRow[mClosingbalance], "#0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])));   // {3,2}    basText.formatNumUnSign(masterRow[mClosingbalance],"#0.00",16),CrDb(Convert.ToDecimal(masterRow[mClosingbalance])) {3,16}  ,valOutstanding,valTransaction  basText.formatNumUnSign(masterRow[mMindueamount],"#0.00",16),{3,16} 
  }



  private void printCardFooterCompany()
  {
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(basText.alignmentMiddle(masterRow[mExternalno], 20) + basText.alignmentMiddle(basText.formatNumUnSign(valCompClosingbalance, "#0.00", 16) + " " + CrDb(Convert.ToDecimal(valCompClosingbalance)), 20) + basText.alignmentMiddle(basText.formatNumUnSign(valCompMinDueAmount, "#0.00", 20), 20) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M"));
  }



  private void calcCardlRowsCompany()
  {
    totalCardPages = 0;
    totCardRows = 0;
    foreach (DataRow dtRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
    {
      if ((dtRow[dPostingdate] == DBNull.Value) && (dtRow[dDocno] == DBNull.Value)) continue;
      totCardRows++;
    }
    if (totCardRows > MaxDetailInLastPage)
    {
      //totalCardPages = 1;
      totCardRows -= MaxDetailInLastPage;
      totalCardPages++;
      totalCardPages += (totCardRows / MaxDetailInPageComp);
      if ((totCardRows % MaxDetailInPageComp) > 0)
        totalCardPages++;
      if (totalCardPages < 1)
        totalCardPages += 1;
    }
    else
    {
      totalCardPages = 1;
    }
  }


  private void calcAccountRowsCompany()
  {
    curAccRows = 0;
    accountRows = null;
    //stmNo = masterRow[mStatementno].ToString();
    //stmNo = masterRow[mCardaccountno].ToString();

    //>accountRows = DSstatement.Tables["tStatementDetailTable"].Select("accountno = '" + clsBasValid.validateStr(masterRow[mAccountno]).ToString().Trim()+ "'");
    accountRows = DSstatement.Tables["tStatementMasterTable"].Select("contractno = '" + curContractNo + "'"); //accountno
    totalAccPages = 0;
    totAccRows = 0;
    compTotalOverdueAmount = 0;
    string prevCardNo = String.Empty, curCardNoStr = String.Empty, prvAccNo = String.Empty;
    int currAccRowsPages = 0;
    foreach (DataRow dtAccRow in accountRows) //mRow.GetChildRows(StatementNoDRel)
    {
      //if((dtAccRow[dPostingdate]== DBNull.Value) && (dtAccRow[dDocno]== DBNull.Value))continue ;
      curCardNoStr = dtAccRow[mCardno].ToString();
      //>>
      stmNo = dtAccRow[dStatementno].ToString();
      //>>>      strStatementNo = dtAccRow[dStatementno].ToString();
      strStatementNo = dtAccRow[mCardaccountno].ToString();
      calcAccountRowsCardHolderCompany();
      //>>>      if(curAccCompanyRows < 1 && Convert.ToDecimal(dtAccRow[mClosingbalance].ToString()) == 0) //Convert.ToDecimal(
      if (Convert.ToDecimal(dtAccRow[mClosingbalance].ToString()) == 0) //Convert.ToDecimal(
        if (curAccCompanyRows < 1)
          continue;

      //>>>      if(dtAccRow[mCardstate].ToString() != "Given" && dtAccRow[mCardstate].ToString() != "New" && dtAccRow[mCardstate].ToString() != "New Pin Generated Only") //  && dtAccRow[mCardstate].ToString() != "Embossed"
      if (!isValidateAccount(dtAccRow[mAccountstatus].ToString()))
        continue;


      //>>
      //if (curCardNoStr.Trim().Length < 1) continue ;
      if (prvAccNo != dtAccRow[mCardaccountno].ToString())
      {
        currAccRowsPages++;
        totAccRows++;
        compTotalOverdueAmount += Convert.ToDecimal(dtAccRow[mTotaloverdueamount].ToString());
      }
      prvAccNo = dtAccRow[mCardaccountno].ToString();

      if (currAccRowsPages > MaxDetailInPageComp)//==
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
  }



  private void completePageDetailRecordsCompany()
  {
    //int curPageLine =CurPageRec4Dtl;
    if (pageNo == totalAccPages)
      for (int curPageLine = CurPageRec4Dtl; curPageLine < MaxDetailInPageComp; curPageLine++)
        streamWrit.WriteLine(String.Empty);
  }

  private void calcAccountRowsCardHolderCompany()
  {
    curAccCompanyRows = 0;
    accountRows = null;
    //stmNo = masterRow[mStatementno].ToString();
    accountRows = DSstatement.Tables["tStatementDetailTable"].Select("accountno = '" + strStatementNo + "'"); //accountno
    string prevCardNo = String.Empty, curCardNoStr = String.Empty;
    foreach (DataRow dtAccRow in accountRows) //mRow.GetChildRows(StatementNoDRel)
    {
      if ((dtAccRow[dPostingdate] == DBNull.Value) && (dtAccRow[dDocno] == DBNull.Value)) continue;
      curCardNoStr = dtAccRow[dCardno].ToString();
      if (curCardNoStr.Trim().Length < 1) continue;
      curAccCompanyRows++;
      prevCardNo = dtAccRow[dCardno].ToString();
    }

    string CurCardNo = String.Empty; string isCardMainNo = String.Empty;
    prevCardNo = String.Empty;
    int currAccRowsPages = 0;

    totCrdNoInAcc = curCrdNoInAcc = 0;
    mainRows = DSstatement.Tables["tStatementMasterTable"].Select("cardaccountno = '" + strStatementNo + "'");
    curMainCard = CurCardNo = "";
    foreach (DataRow mainRow in mainRows) //mRow.GetChildRows(StatementNoDRel)
    {
      totCrdNoInAcc++;
      CurCardNo = mainRow[mCardno].ToString();
      if (mainRow[mCardprimary].ToString() == "Y")
        curMainCard = CurCardNo; //mainRow[mCardno].ToString();
      //if (mainRow[mAccountstatus].ToString() == "Open" && masterRow[mClosingbalance].ToString() != "0") //
      //  isCardMainNo = CurCardNo;
      if (mainRow[mCardprimary].ToString() == "Y" && isValidateCard(mainRow[mCardstate].ToString()))//(mainRow[mCardstate].ToString() == "Given" || mainRow[mCardstate].ToString() == "Embossed" || mainRow[mCardstate].ToString() == "New" || mainRow[mCardstate].ToString() == "New Pin Generated Only")
      {
        curMainCard = CurCardNo; //mainRow[mCardno].ToString();
        break;
      }
    }

    //    if (isCardMainNo = string.Empty && mainRows[0][mAccountstatus].ToString == "Open" && Convert.ToDecimal(masterRow[mClosingbalance].ToString())!= 0.0)
    //    {    }

    if (curMainCard == "")
    {
      if (isCardMainNo != string.Empty)
        curMainCard = isCardMainNo;
      else
        curMainCard = CurCardNo;
    }
  }

  private void printStatementSummary()
  {
    streamSummary.WriteLine(strBankName + " Company Statement");
    streamSummary.WriteLine("__________________________");
    streamSummary.WriteLine("");
    streamSummary.WriteLine("No of Company Statements : " + totNoOfCardStatCmpny.ToString());
    streamSummary.WriteLine("No of Company Pages      : " + totNoOfPageStatCmpny.ToString());
    //streamSummary.WriteLine("");
    //streamSummary.WriteLine("No of Card Holder Statements : " + totNoOfCardStatCrdHldr.ToString());
    //streamSummary.WriteLine("No of Card Holder Pages      : " + totNoOfPageStatCrdHldr.ToString());
  }

  public string bankName
  {
    get { return strBankName; }
    set { strBankName = value; }
  }// bankName

  private void add2FileList(string pFileName)
  {
    int myIndex = aryLstFiles.BinarySearch((object)pFileName);
    if (myIndex < 0)
      aryLstFiles.Add(@pFileName);
  }

  public frmStatementFile setFrm
  {
    set { frmMain = value; }
  }// setFrm

  public string rewardCond
  {
    get { return rewardCondStr; }
    set { rewardCondStr = value; }
  }// rewardCond


  ~clsStatementBNP_CorpCmpny()
  {
    DSstatement.Dispose();
  }
}
