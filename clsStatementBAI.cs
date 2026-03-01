using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Data.OracleClient;
using System.Xml;
using System.Collections;

// Branch 7
public class clsStatementBAI : clsBasStatement
{
  private string strBankName;
  private FileStream fileStrm, fileSummary;
  private StreamWriter streamWrit, streamSummary;
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
  private int totNoOfCardStat, totNoOfPageStat;

  private int curCardRow = 0, curTotCardRows = 0, numOfErr = 0, totNoOfTransactions = 0;
  private bool isPrimaryOnly, isHaveF3 = true, isPrimaryOpened = false;
  private FileStream fileStrmErr;
  private StreamWriter strmWriteErr;
  private string curMainCard;

  private string extAccNum;
  private int totCrdNoInAcc, curCrdNoInAcc;
  private string strOutputPath, strOutputFile, fileSummaryName;
  DateTime vCurDate;

  public clsStatementBAI()
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
      clsBasFile.createDirectory(pStrFileName + vCurDate.ToString("yyyyMM") + pBankName);
      pStrFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM") + ".txt";
      strBankName = pBankName;

      strOutputFile = pStrFileName;

      // open output file
      fileStrm = new FileStream(pStrFileName, FileMode.Create); //Create
      streamWrit = new StreamWriter(fileStrm, Encoding.Default);

      // Error file
      fileStrmErr = new FileStream(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName), FileMode.Create);
      strmWriteErr = new StreamWriter(fileStrmErr, Encoding.Default);

      // open Summary file
      fileSummaryName = pStrFileName;
      fileSummaryName = clsBasFile.getPathWithoutExtn(fileSummaryName) +
        "_Summary." + clsBasFile.getFileExtn(fileSummaryName);
      fileSummary = new FileStream(fileSummaryName, FileMode.Create); //Create
      streamSummary = new StreamWriter(fileSummary, Encoding.Default);


      // set branch for data
      curBranchVal = pBankCode; // 7; //3 = real   1 = test
      //     clsBasStatement.mainTableCond =  " substr(cardno,1,6) !='421192' "; //" cardproduct != 'Visa Business' ";
      //     clsBasStatement.supTableCond =  " substr(cardno,1,6) !='421192' "; 
      // data retrieve
      FillStatementDataSet(pBankCode); //DSstatement =  //7); //3
//      clsBasStatementRaw prnCardNum = new clsBasStatementRaw();
//      prnCardNum.printCardNum(DSstatement, clsBasFile.getPathWithoutFile(fileSummaryName) + @"\BAI Card Number.txt");

      pageNo = 1; totalCardPages = 0;
      curCardNo = String.Empty;
      curAccountNo = String.Empty;

      foreach (DataRow mRow in DSstatement.Tables["statementMasterTable"].Rows)
      {
        masterRow = mRow;
        //streamWrit.WriteLine(masterRow[mStatementno].ToString());
        //pageNo=1; CurPageRec4Dtl=0; CurrentPageFlag = "F 1"; //if page is based on card no
        cardsRows = mRow.GetChildRows(StatementNoDRel); //, DataRowVersion.Proposed

        strCardNo = masterRow[mCardno].ToString().Trim();
        if (!clsCheckCard.isValidCard(strCardNo) || (clsCheckCard.CalcLuhnCheckDigit(strCardNo.Substring(0, 15)) != strCardNo.Substring(15, 1)))
        {
          strmWriteErr.WriteLine("Invalid Card Number" + strCardNo);
          numOfErr++;
        }
        if (strCardNo.Length != 16)
        {
          strmWriteErr.WriteLine("CardNo Length not equal 16 AccountNo " + masterRow[mAccountno].ToString().Trim() + " Card No =" + strCardNo);
          //numOfErr++;
          continue;// Exclude Zero Length Cards 
        }
        strPrimaryCardNo = strCardNo;
        if (masterRow[mCardprimary].ToString() == "N")
        {
          strPrimaryCardNo = masterRow[mPrinarycardno].ToString();
          //calcCardlRows();
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

          //>          if(totAccRows < 1
          //>            && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0) //Convert.ToDecimal(
          //>            continue;
          if (totAccRows < 1
            && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0) //             || (masterRow[mCardno].ToString() == curMainCard   // Convert.ToDecimal(
          {
            isHaveF3 = true;
            //pageNo=1; totalAccPages =1;
            continue;
          }

          //>if(totAccRows < 1)continue ;  //if pages is based on account
          prevAccountNo = masterRow[mAccountno].ToString();
          //pageNo=1; //if page is based on account no
          printHeader();//if page is based on account no

          totNoOfCardStat++;
        } // End of if(prevAccountNo != masterRow[mAccountno].ToString())
        //calcCardlRows();
        //if(totCardRows < 1)continue ;  //if pages is based on card
        //printHeader();//if pages is based on card


        foreach (DataRow dRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
        {
          detailRow = dRow;
          stmNo = detailRow[dStatementno].ToString();
          if ((detailRow[dPostingdate] == DBNull.Value) && (detailRow[dDocno] == DBNull.Value)) continue;// Exclude On-Hold Transactions 
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
          CurPageRec4Dtl = CurPageRec4Dtl + 1;
          printDetail();
          //if(!((detailRow[dPostingdate]== DBNull.Value) && (detailRow[dDocno]== DBNull.Value))) 

        } //end of detail foreach
        //printCardFooter();//if pages is based on card
        // if(masterRow[mCardprimary].ToString() == "Y" && (masterRow[mCardstate].ToString() == "Given" || masterRow[mCardstate].ToString() == "Embossed" || masterRow[mCardstate].ToString() == "New"))
        curCrdNoInAcc++;
        if ((curAccRows >= totAccRows && totAccRows != 0) || (totAccRows == 0 && curCrdNoInAcc == totCrdNoInAcc))
        {
          completePageDetailRecords();
          printCardFooter();//if pages is based on account
          //printAccountFooter();
          CurPageRec4Dtl = 0; //>pageNo=1; if pages is based on account
          curAccRows = 0;
        }
        //streamWrit.WriteLine(strEndOfPage);
        //pageNo=1; CurPageRec4Dtl=0; //if pages is based on card
        //completePageDetailRecords();

      } //end of Master foreach

      /*
          // Write Data to Text File 
          writeStatement(DSstatement,pStrFileName);
          */
      // Write Data to XML File 
      /*
          string XmlFileName = clsBasFile.getPathWithoutExtn(pStrFileName) + ".xml";


          StreamWriter xmlSW = new StreamWriter(XmlFileName); //"Statement.xml"
          DSstatement.WriteXml(xmlSW, XmlWriteMode.IgnoreSchema);
          xmlSW.Close();
          */

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
        /*
              if(numOfErr == 0)
                clsBasFile.deleteFile(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName));
              else
                clsBasFile.deleteFile(pStrFileName);*/
        ArrayList aryLstFiles = new ArrayList();
        //aryLstFiles.Add("");
        aryLstFiles.Add(@strOutputFile);
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


  protected void printHeader()
  {
    extAccNum = clsBasValid.validateStr(masterRow[mExternalno]);
    if (extAccNum.Trim() == "")
      extAccNum = clsBasValid.validateStr(masterRow[mAccountno]);

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
    }

    streamWrit.WriteLine(strEndOfPage);
    streamWrit.WriteLine(basText.replicat(" ", 81) + CurrentPageFlag);  //+"{0,5:MM/yy}",masterRow[mStatementdatefrom]
    streamWrit.WriteLine(basText.replicat(" ", 81) + masterRow[mCardproduct]);  //+"{0,5:MM/yy}",masterRow[mStatementdatefrom]
    //			streamWrit.WriteLine(basText.alignmentLeft(masterRow[mCustomername],50)+ basText.replicat(" ",31) + masterRow[mCardbranchpartname]);  //
    streamWrit.WriteLine(basText.alignmentLeft(masterRow[mCustomername], 50));  //
    streamWrit.WriteLine(basText.replicat(" ", 81) + masterRow[mCardbranchpartname]);  //
    //			streamWrit.WriteLine(basText.alignmentMiddle(masterRow[mCustomeraddress1],50)+ basText.replicat(" ",31) + clsBasValid.validateStr(masterRow[mAccountno])); //basText.formatNum(masterRow[mCardno],"####-####-####-####")
    streamWrit.WriteLine(basText.alignmentMiddle(ValidateArbic(masterRow[mCustomeraddress1].ToString()), 50)); //basText.formatNum(masterRow[mCardno],"####-####-####-####")
    streamWrit.WriteLine(basText.replicat(" ", 81) + extAccNum); //clsBasValid.validateStr(masterRow[mExternalno]) //accountno  //basText.formatNum(masterRow[mCardno],"####-####-####-####")
    //			streamWrit.WriteLine(basText.alignmentMiddle(masterRow[mCustomeraddress2],50)+ basText.replicat(" ",31)+"{0,10:dd/MM/yyyy}",masterRow[mStatementdatefrom]);//+ basText.formatNum(masterRow[mAccountlim],"########")  basText.formatNum(masterRow[mCardlimit],"########")
    streamWrit.WriteLine(basText.alignmentMiddle(ValidateArbic(masterRow[mCustomeraddress2].ToString()), 50));
    streamWrit.WriteLine(basText.replicat(" ", 81) + "{0,10:dd/MM/yyyy}", masterRow[mStatementdateto]);//Statementdatefrom+ basText.formatNum(masterRow[mAccountlim],"########")  basText.formatNum(masterRow[mCardlimit],"########")
    //			streamWrit.WriteLine(basText.alignmentMiddle(masterRow[mCustomeraddress3].ToString().Trim() + " " + masterRow[mCustomerregion].ToString().Trim() + " " + masterRow[mCustomercity].ToString().Trim(),50)+ basText.replicat(" ",31)+pageNo + " / " + totalAccPages);  //basText.formatNum(masterRow[mAccountno],"##-##-######-####")
    streamWrit.WriteLine(basText.alignmentMiddle(ValidateArbic(masterRow[mCustomeraddress3].ToString().Trim()) + " " + masterRow[mCustomerregion].ToString().Trim() + " " + masterRow[mCustomercity].ToString().Trim(), 50));  //basText.formatNum(masterRow[mAccountno],"##-##-######-####")
    streamWrit.WriteLine(basText.replicat(" ", 81) + pageNo + " / " + totalAccPages);  //basText.formatNum(masterRow[mAccountno],"##-##-######-####")
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(basText.alignmentMiddle(curMainCard, 20) + basText.alignmentMiddle(masterRow[mAccountlim], 13) + basText.formatNum(masterRow[mAccountavailablelim], "##0", 20) + basText.alignmentMiddle(masterRow[mMindueamount], 13) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M") + basText.formatNum(masterRow[mTotaloverdueamount], "##0.00", 13)); //+ basText.formatNum(masterRow[mAccountlim],"#,##0.00",9)+" "+basText.formatNum(masterRow[mTotalcredits],"#,##0.00",9)+" "+basText.formatNum(masterRow[mAccountavailablelim],"#,##0.00",9)+" "+basText.formatNum(masterRow[mMindueamount],"#,##0.00",15)+" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",15)+" "+basText.formatDate(masterRow[mStetementduedate],"dd/MM/yy") //TOTALCREDITS// basText.alignmentMiddle(masterRow[mAccountavailablelim],20) //"#,##0"
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

    string trnsDesc; //= detailRow[dTrandescription].ToString().Trim() + " " + detailRow[dMerchant].ToString().Trim()
    if (detailRow[dMerchant].ToString().Trim() == "")
      trnsDesc = detailRow[dTrandescription].ToString().Trim();
    else
      trnsDesc = detailRow[dMerchant].ToString().Trim();

    //			streamWrit.WriteLine("  {0:dd/MM}  {1:dd/MM}  {2,-40} {3,12} {4,3} {5,16} {6,2}",trnsDate,postingDate,basText.trimStr(detailRow[dTrandescription].ToString().Trim() + " " + detailRow[dMerchant].ToString().Trim(),45),basText.formatNum(detailRow[dOrigtranamount],"#,##0.00;(#,##0.00)"),basText.Replace(clsBasValid.validateStr(detailRow[dOrigtrancurrency]),"XXX","   "),basText.formatNum(detailRow[dBilltranamount],"#,##0.00;(#,##0.00)"),CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)
    streamWrit.WriteLine("  {0:dd/MM}  {1:dd/MM}  {2,-24}  {3,-40} {4,16} {5,16} {6,2}", trnsDate, postingDate, basText.trimStr(detailRow[dRefereneno].ToString().Trim(), 20), basText.trimStr(trnsDesc, 40), strForeignCurr, basText.formatNum(detailRow[dBilltranamount], "#0.00;(#0.00)"), CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + " " + isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)
    totNoOfTransactions++;
  }

  protected void printCardFooter()
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
    streamWrit.WriteLine(basText.alignmentMiddle(curMainCard, 35) + basText.alignmentMiddle(masterRow[mMindueamount], 20) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M")
   + basText.alignmentMiddle(basText.formatNumUnSign(masterRow[mOpeningbalance], "#,###,##0.00", 12) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance].ToString())), 15)
   + basText.alignmentMiddle(basText.formatNumUnSign(masterRow[mTotalpayments], "#,###,##0.00", 12) + " " + DbCr(Convert.ToDecimal(masterRow[mTotalpayments])), 15)
   + basText.alignmentMiddle(basText.formatNumUnSign(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal]), "#,###,##0.00", 12) + CrDbMinus(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal])), 15)
   + basText.alignmentMiddle(basText.formatNumUnSign(masterRow[mTotalcharges].ToString(), "#,###,##0.00", 12) + " " + CrDbMinus(Convert.ToDecimal(masterRow[mTotalcharges])), 15)
   + basText.alignmentMiddle(basText.formatNumUnSign(masterRow[mTotalinterest].ToString(), "#,###,##0.00", 12) + " " + CrDbMinus(Convert.ToDecimal(masterRow[mTotalinterest])), 15)
      //       + basText.formatNum(Convert.ToDecimal(masterRow[mTotalcharges]) + Convert.ToDecimal(masterRow[mTotalinterest]),"#,##0.00",12) + CrDbMinus(Convert.ToDecimal(masterRow[mTotalcharges]) + Convert.ToDecimal(masterRow[mTotalinterest])) 

   + basText.alignmentMiddle(basText.formatNumUnSign(masterRow[mClosingbalance], "#,###,##0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])), 15));
  }

  protected void printAccountFooter()
  {
    streamWrit.WriteLine("GRAND" + basText.formatNum(masterRow[mOpeningbalance], "#0.00;(#0.00)", 16) + basText.formatNum(masterRow[mTotaldebits], "#0.00;(#0.00)", 16) + basText.formatNum(masterRow[mTotalcredits], "#0.00;(#0.00)", 16) + basText.formatNum(masterRow[mClosingbalance], "#0.00;(#0.00)", 16) + basText.formatNum(masterRow[mMindueamount], "#0.00;(#0.00)", 16) + basText.replicat(" ", 12) + basText.formatDate(masterRow[mStetementduedate], "yyyy/MM/dd") + "\n");//+strEndOfPage
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

    accountRows = DSstatement.Tables["statementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[mAccountno]).ToString().Trim() + "'");
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
    mainRows = DSstatement.Tables["statementMasterTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[mAccountno]) + "'");
    curMainCard = CurCardNo = "";
    foreach (DataRow mainRow in mainRows) //mRow.GetChildRows(StatementNoDRel)
    {
      totCrdNoInAcc++;
      CurCardNo = mainRow[mCardno].ToString();
      if (mainRow[mCardprimary].ToString() == "Y")
        curMainCard = CurCardNo; //mainRow[mCardno].ToString();
      if (mainRow[mCardprimary].ToString() == "Y" && isValidateCard( mainRow[mCardstate].ToString()))
      {
        curMainCard = CurCardNo; //mainRow[mCardno].ToString();
        break;
      }
    }

    if (curMainCard == "")
      curMainCard = CurCardNo;
  }


  private void writeStatement(DataSet pDS, string pFileName)
  {
    FileStream txtFileStrm = null;
    StreamWriter txtStreamWrit = null;
    try
    {
      // open output file
      string tmpLine;
      pFileName = clsBasFile.getPathWithoutExtn(pFileName) + " Text.txt";
      txtFileStrm = new FileStream(pFileName, FileMode.Create);
      txtStreamWrit = new StreamWriter(txtFileStrm);

      foreach (DataRow mRow in DSstatement.Tables["statementMasterTable"].Rows)
      {
        tmpLine = String.Empty;
        foreach (DataColumn mCol in DSstatement.Tables["statementMasterTable"].Columns)
          tmpLine += mRow[mCol] + "|";//Console.Write("\t{0}", mRow[myCol]);
        txtStreamWrit.WriteLine(tmpLine);
        cardsRows = mRow.GetChildRows(StatementNoDRel);
        foreach (DataRow dRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
        {
          tmpLine = String.Empty;
          foreach (DataColumn dCol in DSstatement.Tables["statementDetailTable"].Columns)
            tmpLine += dRow[dCol] + "|";//Console.Write("\t{0}", mRow[myCol]);
          tmpLine = "\t" + tmpLine;
          txtStreamWrit.WriteLine(tmpLine);
        }
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
      // Close output ile
      txtStreamWrit.Flush();
      txtStreamWrit.Close();
    }
  }


  private void completePageDetailRecords()
  {
    //int curPageLine =CurPageRec4Dtl;
    if (pageNo == totalAccPages)
      for (int curPageLine = CurPageRec4Dtl; curPageLine < MaxDetailInPage; curPageLine++)
        streamWrit.WriteLine(String.Empty);
  }


  private void checkPageDetailRecords()
  {
    if (CurPageRec4Dtl >= MaxDetailInPage)
    {
      CurPageRec4Dtl = 0;
      pageNo++;
      ////printAccountFooter();
      printCardFooter();
      //printPageFooter();
      //printHeader();
    }
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

  public string bankName
  {
    get { return strBankName; }
    set { strBankName = value; }
  }// bankName

  ~clsStatementBAI()
  {
    DSstatement.Dispose();
  }
}
