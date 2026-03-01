using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Data.OracleClient;

// Branch 1
public class clsStatementCreditNSGB : clsBasStatement
{
  private string strBankName;
  private FileStream fileStrm;
  private StreamWriter streamWrit;
  //private DataSet DSstatement;
  //private OracleDataReader drPrimaryCards, drMaster,drDetail;
  private DataRow masterRow;
  private DataRow detailRow;
  private string strEndOfLine = "\u000D";  //+ "M" ^M
  //		private string strEndOfPage = "\u000C"  ;  //+ "\u000D"+ "M" ^L + ^M
  private string strEndOfPage = "\u000C"; //+ basText.replicat(" ",85) + "F 2"  ;  //+ "\u000D"+ "M" ^L + ^M
  private const int MaxDetailInPage = 26; //
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
  private string strOutputPath, strOutputFile, strSumryFile;


  public clsStatementCreditNSGB()
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
      strOutputPath = pStrFileName;
      DateTime vCurDate = pCurDate; //DateTime.Now.AddMonths(-1);
      clsBasFile.createDirectory(pStrFileName + vCurDate.ToString("yyyyMM") + pBankName);
      pStrFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM") + ".txt";
      strBankName = pBankName;

      strOutputFile = pStrFileName;

      // open output file
      fileStrm = new FileStream(pStrFileName, FileMode.Create);
      streamWrit = new StreamWriter(fileStrm);

      // set branch for data
      curBranchVal = pBankCode; // 1;
      isFullFields = false;
      //mainTableCond =  " cardproduct != '" + "Visa Business" + "' "; 
      mainTableCond = " Customertype != '" + "Corporate Customer" + "' ";
      // data retrieve
      FillStatementDataSet(pBankCode); //DSstatement =  //1); //4
      pageNo = 0; totalCardPages = 0;
      curCardNo = String.Empty;
      curAccountNo = String.Empty;

      //ProgressBar control.
      //progBarStatus.Minimum = 1;
      //progBarStatus.Maximum = DSstatement.Tables["statementMasterTable"].Rows.Count();
      //progBarStatus.Value = 1;
      //progBarStatus.Step = 1;
      //frmStatementFile fm = (frmStatementFile.ControlAccessibleObject;

      foreach (DataRow mRow in DSstatement.Tables["statementMasterTable"].Rows)
      {
        //
        //frmStatementFile  . progBarStatus.PerformStep();

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
              clsBasFile.deleteFile(@strOutputFile);
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

        pageNo++;//if pages is based on account
        printCardFooter();
        if (masterRow[mCardprimary].ToString() == "Y" && (masterRow[mCardstate].ToString() == "Given" || masterRow[mCardstate].ToString() == "Embossed" || masterRow[mCardstate].ToString() == "New"))
        {
          printAccountFooter();
          pageNo = 1; CurPageRec4Dtl = 0; //if pages is based on account
        }
        //pageNo=1; CurPageRec4Dtl=0; if pages is based on card
        totNetUsage = 0;
        prevAccountNo = masterRow[mAccountno].ToString();
      }
      //clsBasXML.WriteXmlToFile(DSstatement,clsBasFile.getPathWithoutExtn(pStrFileName)+ ".xml");



      // Close output ile
      streamWrit.Flush();
      streamWrit.Close();
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

      DSstatement.Dispose();
    }
    return rtrnStr;
  }


  protected void printHeader()
  {
    streamWrit.Write(strEndOfPage);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(masterRow[mCustomername].ToString());
    streamWrit.WriteLine(masterRow[mCustomeraddress1].ToString());
    streamWrit.WriteLine(masterRow[mCustomeraddress2].ToString());
    streamWrit.WriteLine(masterRow[mCustomeraddress3].ToString());
    streamWrit.WriteLine(masterRow[mCustomerregion].ToString());
    streamWrit.WriteLine(masterRow[mCustomercity].ToString());
    streamWrit.WriteLine(masterRow[mCustomercountry].ToString());
    if (masterRow[mCardprimary].ToString() == "Y" && (masterRow[mCardstate].ToString() == "Given" || masterRow[mCardstate].ToString() == "Embossed" || masterRow[mCardstate].ToString() == "New"))
    {
      streamWrit.WriteLine(masterRow[mCardno].ToString());
    }
    else
    {
      streamWrit.WriteLine(masterRow[mPrinarycardno].ToString());
    }
    streamWrit.WriteLine("{0,8:dd/MM/yy}", masterRow[mStatementdatefrom]);
    streamWrit.WriteLine(masterRow[mTotalcredits].ToString());
    streamWrit.WriteLine(masterRow[mCardavailablelimit].ToString());
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(masterRow[mMindueamount].ToString());
    streamWrit.WriteLine("{0,8:dd/MM/yy}", masterRow[mStetementduedate]);
    streamWrit.WriteLine("Page {0} of {1}", pageNo, totalAccPages);//totalPages
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    if (masterRow[mCardprimary].ToString() == "Y" && (masterRow[mCardstate].ToString() == "Given" || masterRow[mCardstate].ToString() == "Embossed" || masterRow[mCardstate].ToString() == "New"))
    {
      streamWrit.WriteLine(basText.replicat(" ", 23) + "Primary Card Number  " + masterRow[mCardno].ToString());
    }
    else
    {
      streamWrit.WriteLine(basText.replicat(" ", 23) + "Additional Card  " + masterRow[mCardno].ToString());
    }
    streamWrit.WriteLine(String.Empty);
  }


  protected void printDetail()
  {
    streamWrit.WriteLine(" {0,8:dd/MM/yy}{1,13}{2,-61}{3,-12:f2}{4,2}", detailRow[dTransdate], basText.trimStr(detailRow[dRefereneno], 13), basText.trimStr(detailRow[dTrandescription], 61), detailRow[dBilltranamount], CrDbDetail);//detailRow[dBilltranamountsign]
  }


  protected void printCardFooter()
  {
    string cardFooter;
    cardFooter = basText.replicat(" ", 23) + "Total for Card " + masterRow[mCardno] + basText.replicat(" ", 31) + basText.formatNum(Math.Abs(totNetUsage), "#,##0.00", -16) + CrDb(totNetUsage);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(cardFooter);
    if (masterRow[mCardprimary].ToString() != "Y")
    {
      strAccountFooter += cardFooter + "\r\n"; //\r\n
      intAccountFooter++;
    }
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



  public string bankName
  {
    get { return strBankName; }
    set { strBankName = value; }
  }// bankName

  ~clsStatementCreditNSGB()
  {
    DSstatement.Dispose();
  }
}
