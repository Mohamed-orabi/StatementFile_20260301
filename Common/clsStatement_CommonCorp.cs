using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Xml;
using System.Collections;

// Branch 1
public class clsStatement_CommonCorp : clsBasStatement
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
  private decimal valTransaction = 0, valCompTransaction = 0, valCompMinDueAmount = 0, compTotalOverdueAmount = 0;
  private string curContractNo;


  private int curCardRow = 0, curTotCardRows = 0, numOfErr = 0, totNoOfTransactions = 0;
  private bool isPrimaryOnly, isHaveF3 = true, isPrimaryOpened = false;
  private FileStream fileStrmErr;
  private StreamWriter strmWriteErr;
  private string curMainCard, CurCardAddres1, CurCardAddres2, CurCardAddres3,CurCardNo;
  private int totCrdNoInAcc, curCrdNoInAcc;
  private string strOutputPath, strFileCompany, strFileCardHolder 
    , fileSummaryName;
  private DateTime vCurDate;
  private ArrayList aryLstFiles;
  private string curContarctCurr;
  public clsStatement_CommonCorp()
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
      clsBasFile.createDirectory(strOutputPath);
      pStrFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM") + ".txt";
      strBankName = pBankName;

      //>strFileName = clsBasFile.getPathWithoutExtn(pStrFileName) +
      //>  "_EGP." + clsBasFile.getFileExtn(pStrFileName);
      strFileName = pStrFileName;

      //strFileName = pStrFileName;
      isCorporateVal = true;//clsBasStatement
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

      printCardHolderStatement();
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
          "_Summary." + clsBasFile.getFileExtn(fileSummaryName);
        fileSummary = new FileStream(fileSummaryName, FileMode.Create); //Create
        streamSummary = new StreamWriter(fileSummary, Encoding.Default);

        printStatementSummary();

        // Close Summary File
        streamSummary.Flush();
        streamSummary.Close();
        fileSummary.Close();

        ArrayList aryLstFiles = new ArrayList();
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


  private void printCardHolderStatement()
  {
    strFileCardHolder = strFileName;
    try
    {
      strFileCardHolder = clsBasFile.getPathWithoutExtn(strFileCardHolder) +
        "_CardHolder." + clsBasFile.getFileExtn(strFileCardHolder);

      // open output file
      //>fileStrm = new FileStream(strFileCardHolder, FileMode.Create); //Create
      //>streamWrit = new StreamWriter(fileStrm, Encoding.Default);

      // raw data file
      //>fileRawData = new FileStream(clsBasFile.getPathWithoutExtn(strFileCardHolder) + "_Card Number.txt", FileMode.Create);
      //>strmRawData = new StreamWriter(fileRawData, Encoding.Default);

      pageNo = 0; totalCardPages = 0;
      curCardNo = String.Empty;
      curAccountNo = String.Empty;
      curContarctCurr = String.Empty; 

      foreach (DataRow mRow in DSstatement.Tables["tStatementMasterTable"].Rows)
      {
        masterRow = mRow;
        cardsRows = mRow.GetChildRows(StatementNoDRel); //, DataRowVersion.Proposed

        strCardNo = masterRow[mCardno].ToString();
        strCurrency = masterRow[mAccountcurrency].ToString();
        //>>>        strStatementNo = masterRow[mStatementno].ToString();
        strStatementNo = masterRow[mCardaccountno].ToString();
        if (strCardNo.Length != 16) continue;// Exclude Zero Length Cards 		
        strPrimaryCardNo = strCardNo;
        //check product
        if (curContarctCurr != masterRow[mAccountcurrency].ToString().Trim())
        {
          if (curContarctCurr != string.Empty)
          {
            streamWrit.Flush();
            streamWrit.Close();
            fileStrm.Close();
          }
          curContarctCurr = masterRow[mAccountcurrency].ToString().Trim();
          strFileCardHolder = clsBasFile.getPathWithoutExtn(strFileCompany)
            + "_" + curContarctCurr + "." + clsBasFile.getFileExtn(strFileCompany);
          //curFileName = pStrFileName;
          add2FileList(strFileCardHolder);
          fileStrm = new FileStream(strFileCardHolder, FileMode.Append); //FileMode.Create Create
          streamWrit = new StreamWriter(fileStrm, Encoding.Default);
        }

        if (masterRow[mCardprimary].ToString() == "N")
        {
          strPrimaryCardNo = masterRow[mPrinarycardno].ToString();
        }

        //start new account
        //        if(prevAccountNo != masterRow[mCardno].ToString() + masterRow[mAccountcurrency].ToString())//>masterRow[mCardaccountno].ToString()
        if (prevAccountNo != masterRow[mCardaccountno].ToString())// + masterRow[mAccountcurrency].ToString()>masterRow[mCardaccountno].ToString()
        {
          pageNo = 1; CurPageRec4Dtl = 0; CurrentPageFlag = "F 1"; //if page is based on account no
          calcAccountRowsCardHolder();
          calcAccountRowsCardHolderCompany();

          //>>>>          if(totAccRows < 1 && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0) //Convert.ToDecimal(
          if (Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0) //Convert.ToDecimal(
            if (curAccCompanyRows < 1)
              continue;

          //|>if (curMainCard != masterRow[mCardno].ToString())
          //|>  continue;

          //          if(masterRow[mCardstate].ToString() != "Given" && masterRow[mCardstate].ToString() != "New"  && masterRow[mCardstate].ToString() != "New Pin Generated Only" ) //  && masterRow[mCardstate].ToString() != "Embossed"
          //            continue;

          //if(totAccRows < 1)continue ;  //if pages is based on account
          //>>>>          prevAccountNo = masterRow[mCardno].ToString() + masterRow[mAccountcurrency].ToString(); //>masterRow[mCardaccountno].ToString();
          prevAccountNo = masterRow[mCardaccountno].ToString(); // + masterRow[mAccountcurrency].ToString()   >masterRow[mCardaccountno].ToString();
          //pageNo=1; //if page is based on account no
          printHeaderCardHolder();//if page is based on account no

          strmRawData.WriteLine(basText.formatCardNumber(curMainCard) + "|" + CurCardAddres1 + " " + CurCardAddres2 + " " + CurCardAddres3);//masterRow[mCardno].ToString()
          totNoOfCardStatCrdHldr++;
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
            printCardFooterCardHolder();
            pageNo++;
            printHeaderCardHolder();
          }
          //						totNetUsage +=clsBasValid.validateNum(detailRow[dBilltranamount]);
          totNetUsage = calculateCrDb(totNetUsage, clsBasValid.validateNum(detailRow[dBilltranamount]), clsBasValid.validateStr(detailRow[dBilltranamountsign]));
          CurPageRec4Dtl = CurPageRec4Dtl + 1;
          printDetailCardHolder();
          //if(!((detailRow[dPostingdate]== DBNull.Value) && (detailRow[dDocno]== DBNull.Value))) 

        } //end of detail foreach
        curCrdNoInAcc++;
        if ((curAccRows >= totAccRows && totAccRows != 0)|| (totAccRows == 0 && curCrdNoInAcc == totCrdNoInAcc))// 
        {
          completePageDetailRecordsCardHolder();
          printCardFooterCardHolder();//if pages is based on account
          pageNo = 1; CurPageRec4Dtl = 0; //if pages is based on account
          curAccRows = 0;
        }
        totNetUsage = 0;
        prevAccountNo = masterRow[mCardaccountno].ToString();
        //>prevAccountNo = masterRow[mAccountno].ToString();
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
      // Close output File
      streamWrit.Flush();
      streamWrit.Close();
      fileStrm.Close();

      //>strmRawData.Flush();
      //>strmRawData.Close();
      //>fileRawData.Close();

      //strmRawData.Flush();
      //strmRawData.Close();
      //fileRawData.Close();
    }
  }


  private void printHeaderCardHolder()
  {
    if (pageNo == 1 && totalAccPages == 1)
      CurrentPageFlag = "F 0";
    else if (pageNo == 1 && totalAccPages > 1)
      CurrentPageFlag = "F 1";
    else if (pageNo < totalAccPages)
      CurrentPageFlag = "F 2";
    else if (pageNo == totalAccPages)

      CurrentPageFlag = "F 3";
    if (masterRow[mCardclientname] == masterRow[mCustomername])
    {// error on generating data, the corporate genrated before individual
      throw new ArgumentNullException();
    }

    streamWrit.WriteLine(strEndOfPage);
    streamWrit.WriteLine(basText.replicat(" ", 81) + CurrentPageFlag);
    streamWrit.WriteLine(basText.replicat(" ", 81) + masterRow[mCardproduct]);
    streamWrit.WriteLine(basText.alignmentLeft(masterRow[mCardclientname], 50));  //Customername
    streamWrit.WriteLine(basText.replicat(" ", 81) + masterRow[mCardbranchpartname]);
    streamWrit.WriteLine(basText.alignmentMiddle(masterRow[mCustomername], 50));//customeraddress1
    streamWrit.WriteLine(basText.replicat(" ", 81) + clsBasValid.validateStr(masterRow[mExternalno]));
    streamWrit.WriteLine(basText.alignmentMiddle(masterRow[mCustomeraddress1], 50)); //basText.alignmentMiddle(masterRow[mCustomeraddress2],50)
    streamWrit.WriteLine(basText.replicat(" ", 81) + "{0,10:dd/MM/yyyy}", masterRow[mStatementdateto]);
    streamWrit.WriteLine(basText.alignmentMiddle(masterRow[mCustomeraddress2], 50));//basText.alignmentMiddle(masterRow[mCustomeraddress3].ToString().Trim() + " " + masterRow[mCustomerregion].ToString().Trim() + " " + masterRow[mCustomercity].ToString().Trim(),50)
    streamWrit.WriteLine(basText.replicat(" ", 81) + pageNo + " / " + totalAccPages);
    streamWrit.WriteLine(basText.alignmentMiddle(masterRow[mCustomeraddress3].ToString().Trim() + " " + masterRow[mCustomerregion].ToString().Trim() + " " + masterRow[mCustomercity].ToString().Trim(), 50));
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    //		streamWrit.WriteLine(basText.alignmentMiddle(curMainCard,20) + basText.alignmentMiddle(masterRow[mAccountlim],13) + basText.formatNum(masterRow[mAccountavailablelim],"##0",20) +  basText.alignmentMiddle(masterRow[mMindueamount],13) + basText.formatDate(masterRow[mStetementduedate],"dd/MM/yyyy",15,"M") + basText.formatNum(masterRow[mTotaloverdueamount],"##0.00",13)) ; //+ basText.formatNum(masterRow[mAccountlim],"#,##0.00",9)+" "+basText.formatNum(masterRow[mTotalcredits],"#,##0.00",9)+" "+basText.formatNum(masterRow[mAccountavailablelim],"#,##0.00",9)+" "+basText.formatNum(masterRow[mMindueamount],"#,##0.00",15)+" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",15)+" "+basText.formatDate(masterRow[mStetementduedate],"dd/MM/yy") //TOTALCREDITS// basText.alignmentMiddle(masterRow[mAccountavailablelim],20) //"#,##0"
    //streamWrit.WriteLine(basText.alignmentMiddle(basText.formatCardNumber(curMainCard), 20) + basText.alignmentMiddle(masterRow[mCardlimit], 13) + basText.formatNum(masterRow[mCardavailablelimit], "##0", 20) + basText.alignmentMiddle(masterRow[mMindueamount], 13) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M") + basText.formatNum(masterRow[mTotaloverdueamount], "##0.00", 13)); //+ basText.formatNum(masterRow[mAccountlim],"#,##0.00",9)+" "+basText.formatNum(masterRow[mTotalcredits],"#,##0.00",9)+" "+basText.formatNum(masterRow[mAccountavailablelim],"#,##0.00",9)+" "+basText.formatNum(masterRow[mMindueamount],"#,##0.00",15)+" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",15)+" "+basText.formatDate(masterRow[mStetementduedate],"dd/MM/yy") //TOTALCREDITS// basText.alignmentMiddle(masterRow[mAccountavailablelim],20) //"#,##0" //curMainCard
    streamWrit.WriteLine(basText.alignmentMiddle(basText.formatCardNumber(curMainCard), 20) + basText.alignmentMiddle(masterRow[mAccountlim], 13) + basText.formatNum(masterRow[mCardavailablelimit], "##0", 20) + basText.alignmentMiddle(masterRow[mMindueamount], 13) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M") + basText.formatNum(masterRow[mTotaloverdueamount], "##0.00", 13)); //+ basText.formatNum(masterRow[mAccountlim],"#,##0.00",9)+" "+basText.formatNum(masterRow[mTotalcredits],"#,##0.00",9)+" "+basText.formatNum(masterRow[mAccountavailablelim],"#,##0.00",9)+" "+basText.formatNum(masterRow[mMindueamount],"#,##0.00",15)+" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",15)+" "+basText.formatDate(masterRow[mStetementduedate],"dd/MM/yy") //TOTALCREDITS// basText.alignmentMiddle(masterRow[mAccountavailablelim],20) //"#,##0" //curMainCard
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    if (pageNo == 1)
      streamWrit.WriteLine(basText.replicat(" ", 17) + basText.alignmentLeft("Previous Balance", 63) + basText.formatNumUnSign(masterRow[mOpeningbalance], "##0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance])));  //+ CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign]))
    else
      streamWrit.WriteLine(String.Empty);

    streamWrit.WriteLine(String.Empty);
    totalPages++;
    totNoOfPageStatCrdHldr++;
  }



  private void printDetailCardHolder()
  {
    DateTime trnsDate = clsBasValid.validateDate(detailRow[dTransdate]), postingDate = clsBasValid.validateDate(detailRow[dPostingdate]);
    if (trnsDate > postingDate)
      trnsDate = postingDate;

    if (masterRow[mAccountcurrency].ToString().Trim() != detailRow[dOrigtrancurrency].ToString().Trim())
      strForeignCurr = basText.formatNum(detailRow[dOrigtranamount], "#0.00;(#0.00)") + " " + basText.Replace(clsBasValid.validateStr(detailRow[dOrigtrancurrency]), "XXX", "   ");
    else
      strForeignCurr = basText.replicat(" ", 16);

    string trnsDesc; //= detailRow[dTrandescription].ToString().Trim() + " " + detailRow[dMerchant].ToString().Trim()
    if (detailRow[dMerchant].ToString().Trim() == "")
      trnsDesc = detailRow[dTrandescription].ToString().Trim();
    else
      trnsDesc = detailRow[dMerchant].ToString().Trim();

    streamWrit.WriteLine("  {0:dd/MM}  {1:dd/MM}  {2,-40} {3,16} {4,16} {5,2}", trnsDate, postingDate, basText.trimStr(trnsDesc, 45), strForeignCurr, basText.formatNum(detailRow[dBilltranamount], "#0.00;(#0.00)"), CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + " " + isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)
  }



  private void printCardFooterCardHolder()
  {
    //completePageDetailRecords();
    streamWrit.WriteLine(String.Empty);
    if (pageNo == totalAccPages)
      streamWrit.WriteLine(basText.replicat(" ", 17) + basText.alignmentLeft("Current Balance", 63) + basText.formatNumUnSign(masterRow[mClosingbalance], "#0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])));
    else
      streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(basText.alignmentMiddle(basText.formatCardNumber(curMainCard), 35) + basText.alignmentMiddle(basText.formatNumUnSign(masterRow[mClosingbalance], "#0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])), 20) + basText.alignmentMiddle(masterRow[mMindueamount], 20) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M"));//curMainCard
  }



  private void calcCardlRowsCardHolder()
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



  private void calcAccountRowsCardHolder()
  {
    curAccRows = 0;
    accountRows = null;
    stmNo = masterRow[mStatementno].ToString();
    stmNo = masterRow[mCardaccountno].ToString();

    //>accountRows = DSstatement.Tables["tStatementDetailTable"].Select("accountno = '" + clsBasValid.validateStr(masterRow[mAccountno]).ToString().Trim()+ "'"); //accountno
    accountRows = DSstatement.Tables["tStatementDetailTable"].Select("accountno = '" + strStatementNo + "'"); //accountno
    totalAccPages = 0;
    totAccRows = 0;
    string prevCardNo = String.Empty, curCardNoStr = String.Empty;
    int currAccRowsPages = 0;
    foreach (DataRow dtAccRow in accountRows) //mRow.GetChildRows(StatementNoDRel)
    {
      if ((dtAccRow[dPostingdate] == DBNull.Value) && (dtAccRow[dDocno] == DBNull.Value)) continue;
      curCardNoStr = dtAccRow[dCardno].ToString();
      if (curCardNoStr.Trim().Length < 1) continue;

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
    mainRows = DSstatement.Tables["tStatementMasterTable"].Select("cardaccountno = '" + strStatementNo + "'");//clsBasValid.validateStr(masterRow[mAccountno])
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

  private void calcAccountRowsCardHolderCompany()
  {
    curAccCompanyRows = 0;
    accountRows = null;
    stmNo = masterRow[mStatementno].ToString();
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
      if (mainRow[mAccountstatus].ToString() == "Open" && masterRow[mClosingbalance].ToString() != "0") //
        isCardMainNo = CurCardNo;
      if (mainRow[mCardprimary].ToString() == "Y" && isValidateCard( mainRow[mCardstate].ToString()))//(mainRow[mCardstate].ToString() == "Given" || mainRow[mCardstate].ToString() == "Embossed" || mainRow[mCardstate].ToString() == "New" || mainRow[mCardstate].ToString() == "New Pin Generated Only")
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



  private void completePageDetailRecordsCardHolder()
  {
    //int curPageLine =CurPageRec4Dtl;
    if (pageNo == totalAccPages)
      for (int curPageLine = CurPageRec4Dtl; curPageLine < MaxDetailInPage; curPageLine++)
        streamWrit.WriteLine(String.Empty);
  }


  private void printCompanyStatement()
  {
    strFileCompany = strFileName;
    bool isCompleteLine = true;

    try
    {
      strFileCompany = clsBasFile.getPathWithoutExtn(strFileCompany) +
        "_Company." + clsBasFile.getFileExtn(strFileCompany);

      // open output file
      //>fileStrm = new FileStream(strFileCompany, FileMode.Create); //Create
      //>streamWrit = new StreamWriter(fileStrm, Encoding.Default);

      // Error file
      fileStrmErr = new FileStream(clsBasFile.getPathWithoutExtn(strFileCompany) + "_Err." + clsBasFile.getFileExtn(strFileCompany), FileMode.Create);
      strmWriteErr = new StreamWriter(fileStrmErr, Encoding.Default);

      pageNo = 0; totalCardPages = 0;
      curCardNo = String.Empty;
      curAccountNo = prevAccountNo = String.Empty;
      curContarctCurr = "";
      foreach (DataRow contractRow in DSstatement.Tables["contractTable"].Rows)
      {
        curContractNo = contractRow[mContractno].ToString();
        cardsRows = DSstatement.Tables["tStatementMasterTable"].Select("contractno = '" + curContractNo + "'"); //accountno
        pageNo = 1; curAccRows = CurPageRec4Dtl = 0; CurrentPageFlag = "F 1"; //if page is based on account no
        valCompOutstanding = valCompTransaction = valCompMinDueAmount = 0;
        calcAccountRowsCompany();
        prevAccountNo = masterRow[mContractno].ToString(); //>masterRow[mCardaccountno].ToString()
        //|>totNoOfCardStatCmpny++;
        //check product
        if (curContarctCurr != masterRow[mAccountcurrency].ToString().Trim())
        {
          if (curContarctCurr != string.Empty)
          {
            streamWrit.Flush();
            streamWrit.Close();
            fileStrm.Close();
          }
          curContarctCurr = masterRow[mAccountcurrency].ToString().Trim();
          strFileCompany = clsBasFile.getPathWithoutExtn(strFileCompany)
            + "_" + curContarctCurr + "." + clsBasFile.getFileExtn(strFileCompany);
          //curFileName = pStrFileName;
          add2FileList(strFileCompany);
          fileStrm = new FileStream(strFileCompany, FileMode.Append); //FileMode.Create Create
          streamWrit = new StreamWriter(fileStrm, Encoding.Default);
        }

        foreach (DataRow cardRow in cardsRows)
        {
          masterRow = cardRow;

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
            totNetUsage = calculateCrDb(totNetUsage, clsBasValid.validateNum(detailRow[dBilltranamount]), clsBasValid.validateStr(detailRow[dBilltranamountsign]));
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
    streamWrit.WriteLine(basText.alignmentLeft(masterRow[mCustomername], 50));  //
    streamWrit.WriteLine(basText.replicat(" ", 81) + masterRow[mCardbranchpartname]);  //
    streamWrit.WriteLine(basText.alignmentMiddle(masterRow[mCustomeraddress1], 80)); //basText.formatNum(masterRow[mCardno],"####-####-####-####")
    //    streamWrit.WriteLine( basText.replicat(" ",81) + clsBasValid.validateStr(masterRow[mExternalno])); //accountno  //basText.formatNum(masterRow[mCardno],"####-####-####-####")
    streamWrit.WriteLine(basText.replicat(" ", 81) + "{0,10:dd/MM/yyyy}", masterRow[mStatementdateto]);//Statementdatefrom+ basText.formatNum(masterRow[mAccountlim],"########")  basText.formatNum(masterRow[mCardlimit],"########")
    streamWrit.WriteLine(basText.alignmentMiddle(masterRow[mCustomeraddress2], 80));
    streamWrit.WriteLine(basText.replicat(" ", 81) + pageNo + " / " + totalAccPages);  //basText.formatNum(masterRow[mAccountno],"##-##-######-####")
    streamWrit.WriteLine(basText.alignmentMiddle(masterRow[mCustomeraddress3].ToString().Trim() + " " + masterRow[mCustomerregion].ToString().Trim() + " " + masterRow[mCustomercity].ToString().Trim(), 80));  //basText.formatNum(masterRow[mAccountno],"##-##-######-####")
    //streamWrit.WriteLine(String.Empty);
    //streamWrit.WriteLine(String.Empty);
    //streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    //    streamWrit.WriteLine(basText.alignmentMiddle(masterRow[mContractno],20) + basText.alignmentMiddle(masterRow[mContractlimit],13) + basText.formatNum(masterRow[mAccountavailablelim],"##0",20) + basText.formatDate(masterRow[mStetementduedate],"dd/MM/yyyy",20,"M")) ; // + basText.formatNum(masterRow[mTotaloverdueamount],"##0.00",13)    //+ basText.formatNum(masterRow[mAccountlim],"#,##0.00",9)+" "+basText.formatNum(masterRow[mTotalcredits],"#,##0.00",9)+" "+basText.formatNum(masterRow[mAccountavailablelim],"#,##0.00",9)+" "+basText.formatNum(masterRow[mMindueamount],"#,##0.00",15)+" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",15)+" "+basText.formatDate(masterRow[mStetementduedate],"dd/MM/yy") //TOTALCREDITS// basText.alignmentMiddle(masterRow[mAccountavailablelim],20) //"#,##0"  +  basText.alignmentMiddle(valCompMinDueAmount,13)
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(basText.alignmentMiddle(masterRow[mContractno], 20) + basText.alignmentMiddle(masterRow[mContractlimit], 13) + basText.formatNum(masterRow[mAccountavailablelim], "##0", 20) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 20, "M") + basText.formatNum(compTotalOverdueAmount, "#,##0.00", 15)); //masterRow[mTotaloverdueamount]// + basText.formatNum(masterRow[mTotaloverdueamount],"##0.00",13)    //+ basText.formatNum(masterRow[mAccountlim],"#,##0.00",9)+" "+basText.formatNum(masterRow[mTotalcredits],"#,##0.00",9)+" "+basText.formatNum(masterRow[mAccountavailablelim],"#,##0.00",9)+" "+basText.formatNum(masterRow[mMindueamount],"#,##0.00",15)+" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",15)+" "+basText.formatDate(masterRow[mStetementduedate],"dd/MM/yy") //TOTALCREDITS// basText.alignmentMiddle(masterRow[mAccountavailablelim],20) //"#,##0"  +  basText.alignmentMiddle(valCompMinDueAmount,13)
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
    valCompTransaction += Convert.ToDecimal(masterRow[mClosingbalance].ToString()); //valTransaction;
    valCompMinDueAmount += Convert.ToDecimal(masterRow[mMindueamount].ToString()); //valTransaction;

    streamWrit.WriteLine("  {0,16}  {1,-50} {2,16}", masterRow[mCardno].ToString(), masterRow[mCardclientname].ToString(), basText.formatNumUnSign(masterRow[mMindueamount], "#0.00", 16));   // {3,2}    basText.formatNumUnSign(masterRow[mClosingbalance],"#0.00",16),CrDb(Convert.ToDecimal(masterRow[mClosingbalance])) {3,16}  ,valOutstanding,valTransaction  basText.formatNumUnSign(masterRow[mMindueamount],"#0.00",16),{3,16} 

    /*
        DateTime trnsDate = clsBasValid.validateDate(detailRow[dTransdate]),postingDate = clsBasValid.validateDate(detailRow[dPostingdate]);
        if(trnsDate > postingDate)
          trnsDate= postingDate;

        if(masterRow[mAccountcurrency].ToString().Trim() != detailRow[dOrigtrancurrency].ToString().Trim())
          strForeignCurr = basText.formatNum(detailRow[dOrigtranamount],"#0.00;(#0.00)")+ " " + basText.Replace(clsBasValid.validateStr(detailRow[dOrigtrancurrency]),"XXX","   ");
        else
          strForeignCurr = basText.replicat(" ",16);

        string trnsDesc ; //= detailRow[dTrandescription].ToString().Trim() + " " + detailRow[dMerchant].ToString().Trim()
        if(detailRow[dMerchant].ToString().Trim() == "")
          trnsDesc = detailRow[dTrandescription].ToString().Trim();
        else
          trnsDesc = detailRow[dMerchant].ToString().Trim();

        streamWrit.WriteLine("  {0:dd/MM}  {1:dd/MM}  {2,-40} {3,16} {4,16} {5,2}",trnsDate,postingDate,basText.trimStr(trnsDesc,45), strForeignCurr, basText.formatNum(detailRow[dBilltranamount],"#0.00;(#0.00)"),CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + " " + isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)
        */
  }



  private void printCardFooterCompany()
  {
    //completePageDetailRecords();
    streamWrit.WriteLine(String.Empty);
    /*    if(pageNo == totalAccPages)
          //>      streamWrit.WriteLine(basText.replicat(" ",17) + basText.alignmentLeft("Current Balance",63) + basText.formatNumUnSign(masterRow[mClosingbalance],"#0.00",16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])));  
          streamWrit.WriteLine(basText.replicat(" ",17) + basText.alignmentLeft("Current Balance",63) + basText.formatNumUnSign(valCompTransaction,"#0.00",16) + " " + CrDb(valCompTransaction));  
        else*/
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    //>    streamWrit.WriteLine(basText.alignmentMiddle(masterRow[mContractno],35) + basText.alignmentMiddle(basText.formatNumUnSign(masterRow[mClosingbalance],"#0.00",16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])),20) +  basText.alignmentMiddle(masterRow[mMindueamount],20) + basText.formatDate(masterRow[mStetementduedate],"dd/MM/yyyy",15,"M")) ; 
    //>    streamWrit.WriteLine(basText.alignmentMiddle(masterRow[mContractno],20) + basText.alignmentMiddle(basText.formatNumUnSign(masterRow[mClosingbalance],"#0.00",16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])),20) +  basText.alignmentMiddle(masterRow[mMindueamount],20) + basText.formatDate(masterRow[mStetementduedate],"dd/MM/yyyy",15,"M") +  basText.alignmentMiddle(valCompOutstanding,14) +  basText.alignmentMiddle(valCompTransaction,14)) ; 
    streamWrit.WriteLine(basText.alignmentMiddle(masterRow[mContractno], 20) + basText.alignmentMiddle(basText.formatNumUnSign(valCompTransaction, "#0.00", 16) + " " + CrDb(Convert.ToDecimal(valCompTransaction)), 20) + basText.alignmentMiddle(valCompMinDueAmount, 20) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M"));
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
      curCardNoStr = dtAccRow[dCardno].ToString();
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
      if (dtAccRow[mAccountstatus].ToString() != "Open")
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
      prevCardNo = dtAccRow[mCardno].ToString();
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

  private void printStatementSummary()
  {
    streamSummary.WriteLine(strBankName + " Visa Credit Statement");
    streamSummary.WriteLine("__________________________");
    streamSummary.WriteLine("");
    streamSummary.WriteLine("No of Company Statements : " + totNoOfCardStatCmpny.ToString());
    streamSummary.WriteLine("No of Company Pages      : " + totNoOfPageStatCmpny.ToString());
    streamSummary.WriteLine("");
    streamSummary.WriteLine("No of Card Holder Statements : " + totNoOfCardStatCrdHldr.ToString());
    streamSummary.WriteLine("No of Card Holder Pages      : " + totNoOfPageStatCrdHldr.ToString());
  }


  private void printFileMD5()
  {
    //    FileStream fileStrmMd5;
    //    StreamWriter streamWritMD5;
    //    fileStrmMd5 = new FileStream(strOutputPath + "\\" + strBankName + vCurDate.ToString("yyyyMM") + ".txt", FileMode.Create);
    //    streamWritMD5 = new StreamWriter(fileStrmMd5);
    //    streamWritMD5.WriteLine(clsBasFile.getFileFromPath(strFileCompany) + "  >>  " + clsBasFile.getFileMD5(strFileCompany));
    //    streamWritMD5.WriteLine(clsBasFile.getFileFromPath(strFileCardHolder) + "  >>  " + clsBasFile.getFileMD5(strFileCardHolder));
    //    streamWritMD5.WriteLine(clsBasFile.getFileFromPath(fileSummaryName) + "  >>  " + clsBasFile.getFileMD5(fileSummaryName));
    //    streamWritMD5.Flush();
    //    streamWritMD5.Close();
    //    fileStrmMd5.Close();
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


  ~clsStatement_CommonCorp()
  {
    DSstatement.Dispose();
  }
}
