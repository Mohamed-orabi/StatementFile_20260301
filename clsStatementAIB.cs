using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Data.OracleClient;


// Branch 5
public class clsStatementAIB : clsBasStatement
{
  private string strBankName;
  private FileStream fileStrm;
  private StreamWriter streamWrit;
  //private DataSet DSstatement;
  //private OracleDataReader drPrimaryCards, drMaster,drDetail;
  private DataRow masterRow;
  private DataRow detailRow;
  //private string strEndOfLine = "\u000D" ;  //+ "M" ^M
  private string strEndOfPage = "\u000C"; //+ basText.replicat(" ",85) + "F 2"  ;  //+ "\u000D"+ "M" ^L + ^M
  private const int MaxDetailInPage = 31; //
  private const int linesInLastPage = 67; //
  private int CurPageRec4Dtl = 0;
  private int pageNo = 0, totalPages = 0, totalCardPages = 0
      , totalAccPages = 0, totCardRows = 0, totAccRows = 0;
  private string lastPageTotal;
  private string curCardNo;//,PrevCardNo
  private string curAccountNo, prevAccountNo = String.Empty;//,PrevCardNo
  private string strAccountFooter;
  private int intAccountFooter;
  private decimal totNetUsage = 0;
  private decimal totAccountValue = 0;
  private DataRow[] cardsRows, accountRows;
  private string CrDbDetail;

  public clsStatementAIB()
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

      // merge transaction fee with original transaction
      //clsMaintainData maintainData = new clsMaintainData();
      //maintainData.mergeTrans(pBankCode);
      
      pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);
      DateTime vCurDate = pCurDate; //DateTime.Now.AddMonths(-1);
      clsBasFile.createDirectory(pStrFileName + vCurDate.ToString("yyyyMM") + pBankName);
      pStrFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM") + ".txt";
      strBankName = pBankName;
      
      // open output file
      fileStrm = new FileStream(pStrFileName, FileMode.Create);
      streamWrit = new StreamWriter(fileStrm);

      // set branch for data
      curBranchVal = pBankCode; // 5;  //1
      strOrder = " m.CARDBRANCHPART,m.externalno ";
      // data retrieve
      FillStatementDataSet(pBankCode); //DSstatement =  //5); //1 
      pageNo = 0; totalCardPages = 0;
      curCardNo = String.Empty;
      curAccountNo = String.Empty;
      foreach (DataRow mRow in DSstatement.Tables["statementMasterTable"].Rows)
      {
        masterRow = mRow;
        //streamWrit.WriteLine(masterRow[mStatementno].ToString());
        //pageNo=1;  //if page is based on card no
        CurPageRec4Dtl = 0;
        cardsRows = mRow.GetChildRows(StatementNoDRel); //, DataRowVersion.Proposed
        //start new account
        if (prevAccountNo != masterRow[mAccountno].ToString())
        {
          if (prevAccountNo == string.Empty)
            if (curMonth != Convert.ToDateTime(masterRow[mStatementdateto]).Month)//,"dd/MM/yyyy"
            {
              preExit = false;
              fileStrm.Close();
              //fileStrmErr.Close();
              //clsBasFile.deleteFile(@strOutputFile);
              //clsBasFile.deleteFile(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName));
              return "Error in Generation " + pBankName;
            }
          strAccountFooter = String.Empty;
          intAccountFooter = 0;
          totAccountValue = 0;
          calcAccountRows();
          prevAccountNo = masterRow[mAccountno].ToString();
          pageNo = 1; ; //if page is based on account no 
        }
        calcCardlRows();
        if (totCardRows > 0)
          printHeader();

        foreach (DataRow dRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
        {
          detailRow = dRow;
          CrDbDetail = String.Empty;
          if (CurPageRec4Dtl >= MaxDetailInPage)
          {
            CurPageRec4Dtl = 0;
            pageNo++;
            //printAccountFooter();
            printPageFooter();
            printHeader();
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
        if (totCardRows > 0)
          printCardFooter();
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
      }
      //fillStatementHistory(pStmntType,pAppendData);
      //clsBasXML.WriteXmlToFile(DSstatement,clsBasFile.getPathWithoutExtn(pStrFileName)+ ".xml");
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

        DSstatement.Dispose();
      }
    }
      return rtrnStr;
  }



  protected void printHeader()
  {
    string strHeader;
    streamWrit.Write(strEndOfPage);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(basText.alignmentLeft(masterRow[mCustomername], 50) + basText.replicat(" ", 15) + basText.alignmentLeft(masterRow[mAccounttype], 35));
    streamWrit.WriteLine(basText.alignmentLeft(masterRow[mCustomeraddress1], 50) + basText.replicat(" ", 15) + basText.alignmentLeft(masterRow[mAccountno], 35));
    streamWrit.WriteLine(basText.alignmentLeft(masterRow[mCustomeraddress2], 50) + basText.replicat(" ", 15) + basText.alignmentLeft(masterRow[mCardbranchpart].ToString() + "  " + masterRow[mCardbranchpartname].ToString(), 35));
    streamWrit.WriteLine(basText.alignmentLeft(masterRow[mCustomeraddress3], 50) + basText.replicat(" ", 15) + String.Format("{0,8:dd/MM/yy}", masterRow[mStatementdatefrom]));
    streamWrit.WriteLine(basText.alignmentLeft(masterRow[mCustomerregion].ToString() + " " + masterRow[mCustomercity].ToString(), 50) + basText.replicat(" ", 15) + String.Format("Page {0} of {1}", pageNo, totalAccPages));
    strHeader = basText.formatNum(masterRow[mAccountlim], "#,##0.00", 9) + " " + basText.formatNum(masterRow[mTotalcredits], "#,##0.00", 9) + " " + basText.formatNum(masterRow[mAccountavailablelim], "#,##0.00", 9) + " " + basText.formatNum(masterRow[mMindueamount], "#,##0.00", 15) + " " + basText.formatNum(masterRow[mClosingbalance], "#,##0.00", 15) + " " + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yy");
    if (masterRow[mCardprimary].ToString() == "Y" && (masterRow[mCardstate].ToString() == "Given" || masterRow[mCardstate].ToString() == "Embossed" || masterRow[mCardstate].ToString() == "New" || masterRow[mCardstate].ToString() == "New Pin Generated Only"))
    {
      streamWrit.WriteLine("  " + basText.alignmentLeft(masterRow[mCardno], 16) + " " + strHeader);
    }
    else
    {
      streamWrit.WriteLine(basText.replicat(" ", 19) + strHeader);
    }
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
  }


  protected void printDetail()
  {
    streamWrit.WriteLine("  {0:dd/MM} {1:dd/MM} {2,13} {3,-30} {4,12} {5,3} {6,12} {7,2}", clsBasValid.validateDate(detailRow[dTransdate]), clsBasValid.validateDate(detailRow[dPostingdate]), basText.trimStr(detailRow[dRefereneno], 13), basText.alignmentLeft(detailRow[dTrandescription], 30), basText.formatNum(detailRow[dOrigtranamount], "#,##0.00;(#,##0.00)"), basText.Replace(clsBasValid.validateStr(detailRow[dOrigtrancurrency]), "XXX", "   "), basText.formatNum(detailRow[dBilltranamount], "#,##0.00;(#,##0.00)"), CrDbDetail);//detailRow[dBilltranamountsign]
    //{2,13} {3,-30} {4,12} {5,3} {6,12} {7,2} ,basText.trimStr(detailRow[dRefereneno],13),basText.alignmentLeft(detailRow[dTrandescription],30),basText.formatNum(detailRow[dOrigtranamount],"#,##0.00;(#,##0.00)"),clsBasValid.validateStr(detailRow[dOrigtrancurrency]),basText.formatNum(detailRow[dBilltranamount],"#,##0.00;(#,##0.00)"),CrDbDetail
    //streamWrit.WriteLine("{0:dd/MM} {1:dd/MM}",clsBasValid.validateDate(detailRow[dTransdate],"1/1/1900"),clsBasValid.validateDate(detailRow[dTransdate],"1/1/1900")); //clsBasValid.validateDate(detailRow[dPostingdate],"1/1/1900"))
  }


  protected void printCardFooter()
  {
    string pageFooter;
    pageFooter = basText.replicat(" ", 29) + "** Card . " + masterRow[mCardno] + " SUBTOTAL- " + basText.replicat(" ", 18) + basText.formatNum(Math.Abs(totNetUsage), "#,##0.00", -16) + CrDb(totNetUsage);
    streamWrit.WriteLine(pageFooter);
    CurPageRec4Dtl++;
    checkPageDetailRecords();
  }


  protected void printPageFooter()
  {
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    string cardFooter;
    cardFooter = basText.formatNum(masterRow[mOpeningbalance], "#,##0.00", 12) + " " + basText.formatNum(masterRow[mTotalpayments], "#,##0.00", 12) + " " + basText.formatNum(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal]), "#,##0.00", 12) + " " + basText.formatNum(masterRow[mTotalcharges], "#,##0.00", 12) + " " + basText.formatNum(0.00, "#,##0.00", 12) + " " + basText.formatNum(masterRow[mClosingbalance], "#,##0.00", 12);
    //streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(cardFooter);
    /*if(masterRow[mCardprimary].ToString() != "Y")
    {
        strAccountFooter += cardFooter + "\r\n" ; //\r\n
        intAccountFooter++;
    }*/
  }


  protected void printAccountFooter()
  {
    int curLastPageLine = CurPageRec4Dtl;
    curLastPageLine += 24; //add header Lines
    curLastPageLine += 2; // blank line after detail line total for last card
    curLastPageLine += intAccountFooter; //add Previous pages totals

    if (strAccountFooter.Length > 1)
    {
      strAccountFooter = strAccountFooter.Substring(0, strAccountFooter.Length - 2);
      streamWrit.WriteLine(strAccountFooter);
      streamWrit.WriteLine(basText.replicat(" ", 23) + "Total for all Cards " + basText.replicat(" ", 42) + basText.formatNum(Math.Abs(totAccountValue), "#,##0.00", -16) + CrDb(totAccountValue));
      curLastPageLine++;
      //curLastPageLine += CurPageRec4Dtl +1; //add last page totals
    }

    curLastPageLine += 10; // subtract last page footer
    for (; curLastPageLine < linesInLastPage; curLastPageLine++)
      streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(" **(lastpage)");
    streamWrit.WriteLine(" " + basText.formatNum(masterRow[mOpeningbalance], "#,##0.00", -16));  //
    streamWrit.WriteLine(" " + basText.formatNum(masterRow[mTotalcashwithdrawal], "#,##0.00", -16));  //
    streamWrit.WriteLine(" " + basText.formatNum(masterRow[mTotalpurchases], "#,##0.00", -16));  //
    streamWrit.WriteLine(" " + basText.formatNum(masterRow[mTotalinterest], "#,##0.00", -16));  //
    streamWrit.WriteLine(" " + basText.formatNum(masterRow[mTotalcharges], "#,##0.00", -16));  //
    streamWrit.WriteLine(" " + basText.formatNum(masterRow[mTotalcredits], "#,##0.00", -16));  //
    streamWrit.WriteLine(" " + basText.formatNum(masterRow[mClosingbalance], "#,##0.00", -16));  //
    streamWrit.WriteLine(" " + basText.alignmentLeft(masterRow[mStatementmessageline1], 100));  //
    streamWrit.WriteLine(" " + basText.alignmentLeft(masterRow[mStatementmessageline2], 100));  //



  }


  private void calcCardlRows()
  {
    totalCardPages = 0;
    totCardRows = 0;
    foreach (DataRow dtRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
    {
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
    accountRows = DSstatement.Tables["statementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[mAccountno]) + "'");
    totalAccPages = 0;
    totAccRows = 0;
    string prevCardNo = String.Empty, CurCardNo = String.Empty;
    int currAccRowsPages = 0;
    foreach (DataRow dtAccRow in accountRows) //mRow.GetChildRows(StatementNoDRel)
    {
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

  }


  private void completePageDetailRecords()
  {
    //int curPageLine =CurPageRec4Dtl;
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
      printPageFooter();
      //printHeader();
    }
  }

  public string bankName
  {
    get { return strBankName; }
    set { strBankName = value; }
  }// bankName

  
  ~clsStatementAIB()
  {
    DSstatement.Dispose();
  }
}
