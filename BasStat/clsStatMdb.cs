using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Collections;

// export the statement data to access MDB
public class clsStatMdb : clsBasStatement
{
  protected string strBankName;
  protected FileStream fileStrmBasic, fileStrmTrans;//, fileStrmSubtotal
  protected string strFileBasic, strFileTrans;//, strFileSubtotal
  protected StreamWriter streamWritBasic, streamWritTrans;//, streamWritSubtotal
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
  protected const string strFileSpr = "','";//#
  protected bool isPrimaryOnly;
  protected string curMainCard;
  protected int curAccRows = 0;
  protected int totCrdNoInAcc, curCrdNoInAcc;
  protected string stmNo;
  protected DataRow[] mainRows;

  protected string extAccNum;
  protected string strOutputPath, strOutputFile, fileSummaryName;
  protected DateTime vCurDate;
  protected DataSet DSstatementRaw;
  //protected DataRelation StatementNoDRels;

  protected OleDbConnection conn;
  protected string strSqlActn;

  protected string StrStatLable = string.Empty, strWhereCond = string.Empty;
  protected ArrayList aryLstFiles = new ArrayList();
  protected string strFileName;
  private string strExcludeCond = string.Empty;
  protected bool isExcludedVal = false;

  public clsStatMdb()
  {
  }

  public string Statement(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pDate, string pStmntType, bool pAppendData)
  {
    string rtrnStr = "Successfully Generate " + pBankName;
    int curMonth = pDate.Month;
    bool preExit = true;
    curMonth = pDate.Month;

    clsMaintainData maintainData = new clsMaintainData();
    maintainData.matchCardBranch4Account(pBankCode);

    // merge transaction fee with original transaction
    //clsMaintainData maintainData = new clsMaintainData();
    //maintainData.mergeTrans(pBankCode);
    pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);
    vCurDate = pDate; //DateTime.Now.AddMonths(-1);
    strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
    clsBasFile.createDirectory(pStrFileName + vCurDate.ToString("yyyyMM") + pBankName);
    pStrFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM") + ".mdb";
    strFileName = pStrFileName;
    clsBasFile.copyFile(@"D:\pC#\exe\FilesForPrograms\StatementData.mdb", pStrFileName);
    curBranchVal = pBankCode; // 4; //4  = real   1 = test
    if (isExcludedVal)
        {
        MainTableCond = " m.cardproduct not in " + strExcludeCond + "";//strWhereCond
        //supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.cardproduct not in " + strExcludeCond + ")";
        }
    FillStatementDataSet(pBankCode,""); //DSstatement =  //6); // 6
    Statement(pStrFileName, DSstatement);
    return "";
  }

  public string Statement(string pStrFileName, DataSet pDSstatement)
  {
    try
    {
      DSstatementRaw = pDSstatement;

      conn = new OleDbConnection(clsDbCon.sConMsAccess(pStrFileName,""));
      conn.Open();
      pageNo = 0; totalCardPages = 0;
      curCardNo = String.Empty;
      curAccountNo = String.Empty;

      foreach (DataRow mRow in DSstatementRaw.Tables["tStatementMasterTable"].Rows)
      {
        masterRow = mRow;
        //streamWrit.WriteLine(masterRow[mStatementno].ToString());
        pageNo = 1;  //if page is based on card no
        CurPageRec4Dtl = 0;
        //cardsRows = mRow.GetChildRows("StaementNoDR"); //StatementNoDRel, DataRowVersion.Proposed
        cardsRows = DSstatement.Tables["tStatementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[mAccountno]).ToString().Trim() + "'");
        //start new account
        if (masterRow[mHOLSTMT].ToString() == "Y")
            continue;
        if (prevAccountNo != masterRow[mAccountno].ToString())
        {
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
      conn.Close();

      //ArrayList aryLstFiles = new ArrayList();
      //aryLstFiles.Add("");
      aryLstFiles.Add(pStrFileName);
      //aryLstFiles.Add(@strFileSubtotal);
      
      //clsBasFile.generateFileMD5(aryLstFiles, @clsBasFile.getPathWithoutExtn(pStrFileName)+ ".MD5"); //+ "_Raw" 
      //aryLstFiles.Add(@clsBasFile.getPathWithoutExtn(pStrFileName)+ ".MD5"); //+ "_Raw" 
      //SharpZip zip = new SharpZip();
      //zip.createZip(aryLstFiles, @clsBasFile.getPathWithoutExtn(pStrFileName)+ ".zip", ""); //+ "_Raw" 

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

    strSqlActn = "insert into Main values ( '";
    strSqlActn += basText.validateWriteField(masterRow[mCustomername]) + strFileSpr
      //+ ValidateArbic(basText.validateWriteField(masterRow[mCustomeraddress1])) + strFileSpr + ValidateArbic(basText.validateWriteField(masterRow[mCustomeraddress2]))
      + ValidateArbic(basText.validateWriteField(newaddress1)) + strFileSpr + ValidateArbic(basText.validateWriteField(newaddress2))
      + strFileSpr + ValidateArbic(basText.validateWriteField(masterRow[mCustomeraddress3])) + strFileSpr
      + basText.validateWriteField(masterRow[mCustomerregion])
      + basText.validateWriteField(masterRow[mCustomercity])
      + "" + strFileSpr + "" + basText.validateWriteField(masterRow[mContracttype]) + strFileSpr + masterRow[mAccountno].ToString()
      + "" + strFileSpr + "" + basText.validateWriteField(masterRow[mCardbranchpart]) + "  "
      + basText.validateWriteField(masterRow[mCardbranchpartname]) + strFileSpr
      + String.Format("{0,8:dd/MM/yy}", masterRow[mStatementdateto])
      //+ strFileSpr + masterRow[mCardno].ToString() 
      + strFileSpr + basText.formatCardNumber(curMainCard)
      + strFileSpr + masterRow[mAccountlim].ToString()
      + strFileSpr + masterRow[mAccountavailablelim].ToString()
      + strFileSpr + masterRow[mMindueamount].ToString()
      + strFileSpr + masterRow[mClosingbalance].ToString()
      + CrDb(Convert.ToDecimal(masterRow[mClosingbalance].ToString())) + strFileSpr
      + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yy")
      + strFileSpr + masterRow[mTotaloverdueamount].ToString() + strFileSpr //0.00 
      + basText.formatNum(masterRow[mOpeningbalance], "#,##0.00", 12)
      + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance].ToString())) + strFileSpr
      + basText.formatNum(masterRow[mTotalpayments], "#,##0.00", 12)
      + DbCr(Convert.ToDecimal(masterRow[mTotalpayments])) + strFileSpr
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
      //+ String.Format("Page {0} of {1}", pageNo, totAccCards) + strFileSpr
      + extAccNum ; //masterRow[mExternalno].ToString()//ACCOUNTTYPE// + "" + strFileSpr
    strSqlActn += "')";
    (new OleDbCommand(strSqlActn, conn)).ExecuteNonQuery();

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
   //			streamWrit.WriteLine("  {0:dd/MM} {1:dd/MM} {2,13} {3,-30} {4,12} {5,3} {6,12} {7,2}",clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]),basText.trimStr(detailRow[dRefereneno],13),basText.alignmentLeft(detailRow[dTrandescription],30),basText.formatNum(detailRow[dOrigtranamount],"#,##0.00;(#,##0.00)"),basText.Replace(clsBasValid.validateStr(detailRow[dOrigtrancurrency]),"XXX","   "),basText.formatNum(detailRow[dBilltranamount],"#,##0.00;(#,##0.00)"),CrDbDetail);//detailRow[dBilltranamountsign]
    //!streamWritTrans.WriteLine("{0}" + strFileSpr + "{1:dd/MM}" + strFileSpr + "{2:dd/MM}" + strFileSpr + "{3,13}" + strFileSpr + "{4,-40}" + strFileSpr + "{5,12} {6,3}" + strFileSpr + "{7,12} {8,2}" + strFileSpr + "", masterRow[mCardno], trnsDate, postingDate, basText.trimStr(detailRow[dRefereneno], 13), basText.trimStr(detailRow[dTrandescription].ToString().Trim() + " " + detailRow[dMerchant].ToString().Trim(), 40), basText.formatNum(detailRow[dOrigtranamount], "#,##0.00;(#,##0.00)"), basText.Replace(clsBasValid.validateStr(detailRow[dOrigtrancurrency]), "XXX", "   "), basText.formatNum(detailRow[dBilltranamount], "#,##0.00;(#,##0.00)"), CrDbDetail);//detailRow[dBilltranamountsign]//clsBasValid.validateDate(detailRow[dTransdate])//clsBasValid.validateDate(detailRow[dPostingdate])

    trnsDesc = trnsDesc.Replace("'", "''");  //basText.replaceStr(trnsDesc,"'","''"); //
    strSqlActn = "insert into Transactions values ( '";
    strSqlActn += basText.formatCardNumber(curMainCard) + strFileSpr + trnsDate.ToString("dd/MM") + strFileSpr + postingDate.ToString("dd/MM") + strFileSpr + detailRow[dRefereneno].ToString() + strFileSpr + basText.trimStr(trnsDesc, 40) + strFileSpr + strForeignCurr + strFileSpr + basText.formatNum(detailRow[dBilltranamount], "#,##0.00;(#,##0.00)") + CrDbDetail + strFileSpr + detailRow[dAccountno].ToString();
    strSqlActn += "')";
    (new OleDbCommand(strSqlActn, conn)).ExecuteNonQuery();
  }


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



  public void CreateZip()
  {
    clsBasFile.generateFileMD5(aryLstFiles, @clsBasFile.getPathWithoutExtn(strFileName) + ".MD5"); //+ "_Raw" 
    aryLstFiles.Add(@clsBasFile.getPathWithoutExtn(strFileName) + ".MD5"); //+ "_Raw" 
    SharpZip zip = new SharpZip();
    zip.createZip(aryLstFiles, @clsBasFile.getPathWithoutExtn(strFileName) + ".zip", ""); //+ "_Raw" 
  }

  public string StatLable
  {
    get { return StrStatLable; }
    set { StrStatLable = value; }
  }// StatLable

  public string whereCond
  {
    get { return strWhereCond; }
    set { strWhereCond = value; }
  }  // whereCond

  public string bankName
  {
    get { return strBankName; }
    set { strBankName = value; }
  }// bankName

  public string ExcludeCond
      {
      get { return strExcludeCond; }
      set { strExcludeCond = value; }
      }  // ExcludeCond

  public bool isExcluded
      {
      get { return isExcludedVal; }
      set { isExcludedVal = value; }
      }// isExcluded

  ~clsStatMdb()
  {
    DSstatementRaw.Dispose();
  }
}
