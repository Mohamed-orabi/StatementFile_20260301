using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Collections;
using System.Drawing; 

using System.Reflection;
using Microsoft.Office.Interop.Excel;

// Branch X
public class clsStatExcel : clsBasStatement //clsBasStatementFunc
{
  protected string strBankName;
  protected FileStream fileStrmBasic, fileStrmTrans;//, fileStrmSubtotal
  protected string strFileBasic, strFileTrans;//, strFileSubtotal
  protected StreamWriter streamWritBasic, streamWritTrans;//, streamWritSubtotal
  protected DataRow masterRow;
  protected DataRow detailRow;
  protected const int MaxDetailInPage = 20; //
  protected const int linesInLastPage = 67; //
  protected int CurPageRec4Dtl=0 ;
  protected int pageNo=0, totalCardPages=0 //, totalPages=0
    , totalAccPages=0, totCardRows=0, totAccRows=0, totAccCards=0;
  //	protected string lastPageTotal ;
  protected string curCardNo, CardNumber=String.Empty, curCardNumber=String.Empty, PrevCardNumber=String.Empty;// 
  protected string curAccountNo, prevAccountNo=String.Empty ;//
  protected string strAccountFooter;
  protected int intAccountFooter;
  protected decimal totNetUsage = 0;
  protected decimal totAccountValue = 0;
  protected DataRow[] cardsRows, accountRows;
  protected string CrDbDetail;
  protected const string strFileSpr = "|";//#
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
  protected Workbook wb;
  protected Worksheet wsMain, wsDtl;
  protected string strReadLine = string.Empty;
  protected int numOfLine = 1, numOfColumn = 0;
  protected string fileHeader = string.Empty, fileHeaderDtl = string.Empty;// ,pType1
  protected string[] strHeader, strHeaderDtl, strLine;
  protected int mstrCnt = 1,dtlCnt = 1;
  protected Range aRange;
  Microsoft.Office.Interop.Excel.Application xlApp = null;

  public clsStatExcel()
  {
  }

  public void Statement(string pStrFileName, DataSet pDSstatement)
  {
    try 
    {
      DSstatementRaw = pDSstatement;
      // open output file
      strFileBasic = clsBasFile.getPathWithoutExtn(pStrFileName)+ ".xls";
      clsBasFile.deleteFile(strFileBasic);
      //fileStrmBasic = new FileStream(strFileBasic, FileMode.Create);
      //streamWritBasic = new StreamWriter(fileStrmBasic);
      // open output file
      // Excel File
      xlApp = new Microsoft.Office.Interop.Excel.Application(); // Microsoft.Office.Interop.Excel.Application 

      if (xlApp == null)
      {
        Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
        return;
      }
      xlApp.Visible = false; //true
      wb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);//XlWBATemplate.xlWBATWorksheet



      string strFileHeader = "Name" + strFileSpr + "Address1" + strFileSpr + "Address2" + strFileSpr + "Address3" +
        strFileSpr + "Address4" + strFileSpr + "Card Type" + strFileSpr + "Account" + strFileSpr + "Branch" + strFileSpr + "Statement Date" +
        strFileSpr + "Card No." + strFileSpr + "Credit Limit" + strFileSpr + "Available Credit" +
        strFileSpr + "Min. Payment Due" + strFileSpr + "New Balance" + strFileSpr + "Payment Due Date" +
        strFileSpr + "Past Due Amount" + strFileSpr + "prev.balance" + strFileSpr + "payment & CR" +
        strFileSpr + "Purch.Cash&Dr" + strFileSpr + "Finance Charge" + strFileSpr + "Pages" + strFileSpr + "External NO" + strFileSpr;
      //streamWritBasic.WriteLine(strFileHeader);//#late charge#new balance

      wsDtl = (Worksheet)wb.Worksheets[1];
      //wsDtl = (Worksheet)wb.Worksheets.Add(System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
      wsMain = (Worksheet)wb.Worksheets.Add(System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);

      if (wsMain == null)
      {
        Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
      }
      wsMain.Name = "Main Data";
      wsMain.Tab.Color = Convert.ToDecimal(ColorTranslator.ToOle(Color.RoyalBlue));
      //wsMain.Tab.ColorIndex = XlColorIndex.xlColorIndexAutomatic;
      strHeader = strFileHeader.Split(Convert.ToChar(strFileSpr));
      numOfLine = 1; numOfColumn = 0;
      foreach (string strHdr in strHeader)
      {
        aRange = wsMain.get_Range(basText.Num2Chr(numOfColumn) + numOfLine.ToString(), basText.Num2Chr(numOfColumn) + numOfLine.ToString());
        aRange.Value2 = strHdr;
        aRange.Font.Size = 11;
        aRange.Font.Bold = true;
        //aRange.Font.Background = Convert.ToDecimal(ColorTranslator.ToOle(Color.Beige)); 
        //aRange.Cells.FillLeft();
        //string ss = aRange.Justify() ; // .FillLeft();
        //aRange.FillLeft();
        aRange.Font.Color = Convert.ToDecimal(ColorTranslator.ToOle(Color.Red));//Color.RosyBrown 
        numOfColumn++;
      }


      fileHeaderDtl = "Card No." + strFileSpr + "Date of Trans" + strFileSpr + "Date of Post" + strFileSpr + "Reference" + strFileSpr + "Description" + strFileSpr + "Purchase Currency & Amount" + strFileSpr + "Amount" + strFileSpr;
      //wsDtl = wb.Worksheets.Add(System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
      //xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

      //wsDtl.SmartTags. = 2;
      if (wsDtl == null)
      {
        Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
      }
      wsDtl.Name = "Transaction Data";
      wsDtl.Tab.Color = Convert.ToDecimal(ColorTranslator.ToOle(Color.YellowGreen));
      //wsDtl.Tab.ColorIndex = XlColorIndex.xlColorIndexAutomatic;
      //wsMain.Tab.Color = Convert.ToDecimal(ColorTranslator.ToOle(Color.AliceBlue));
      //wsMain.Color = Convert.ToDecimal(ColorTranslator.ToOle(Color.AliceBlue));
      //Range aRange;
      strHeaderDtl = fileHeaderDtl.Split(Convert.ToChar(strFileSpr));
      numOfLine = 1; numOfColumn = 0;
      foreach (string strHdr in strHeaderDtl)
      {
        aRange = wsDtl.get_Range(basText.Num2Chr(numOfColumn) + numOfLine.ToString(), basText.Num2Chr(numOfColumn) + numOfLine.ToString());
        aRange.Value2 = strHdr;
        aRange.Font.Size = 11;
        aRange.Font.Bold = true;
        //aRange.Font.Background = Convert.ToDecimal(ColorTranslator.ToOle(Color.Beige));
        //aRange.Cells.FillLeft();
        //string ss = aRange.Justify() ; // .FillLeft();
        //aRange.FillLeft();
        aRange.Font.Color = Convert.ToDecimal(ColorTranslator.ToOle(Color.Red));//Color.RosyBrown 
        numOfColumn++;
      }


      numOfColumn = 0;
      numOfLine++;


      //strFileTrans = clsBasFile.getPathWithoutExtn(pStrFileName)+ "_Trns.txt";
      //clsBasFile.deleteFile(strFileTrans);
      //fileStrmTrans = new FileStream(strFileTrans,FileMode.Create);
      //streamWritTrans = new StreamWriter(fileStrmTrans);
      //streamWritTrans.WriteLine("Card No." + strFileSpr + "Date of Trans" + strFileSpr + "Date of Post" + strFileSpr + "Reference" + strFileSpr + "Description" + strFileSpr + "Purchase Currency & Amount" + strFileSpr + "Amount" + strFileSpr);//detailRow[dBilltranamountsign]
      pageNo=0; totalCardPages=0;
      curCardNo=String.Empty;
      curAccountNo=String.Empty;

      foreach (DataRow mRow in DSstatementRaw.Tables["tStatementMasterTable"].Rows)
      {
        masterRow = mRow;
        pageNo=1;  //if page is based on card no
        CurPageRec4Dtl=0;
        cardsRows = mRow.GetChildRows("StaementNoDR"); //StatementNoDRel, DataRowVersion.Proposed
        //start new account
        if(prevAccountNo != masterRow[mAccountno].ToString())
        {
          strAccountFooter = String.Empty;
          intAccountFooter = 0;
          totAccountValue = 0;
          calcAccountRows();
          prevAccountNo = masterRow[mAccountno].ToString();
          pageNo=1;  //if page is based on account no 
          if (totAccRows > 0 || Convert.ToDecimal(masterRow[mClosingbalance].ToString()) != 0)
            printHeader();//>>
          isPrimaryOnly = false;
        }
        if (totAccRows < 1 && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0) //             || (masterRow[mCardno].ToString() == curMainCard   // Convert.ToDecimal(
          continue;

        curCardNumber = masterRow[mCardno].ToString();

        foreach (DataRow dRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
        {
          detailRow = dRow;
          if((detailRow[dPostingdate]== DBNull.Value) && (detailRow[dDocno]== DBNull.Value))
            continue ;// Exclude On-Hold Transactions 
          CrDbDetail=String.Empty;
          if(CurPageRec4Dtl >= MaxDetailInPage)
          {
            CurPageRec4Dtl=0;
            pageNo++;
          }
          if(clsBasValid.validateStr(detailRow[dBilltranamountsign])=="CR")
          {
            CrDbDetail ="CR";
            totNetUsage +=clsBasValid.validateNum(detailRow[dBilltranamount]);
          }
          else
          {
            CrDbDetail = String.Empty;
            totNetUsage -=clsBasValid.validateNum(detailRow[dBilltranamount]);
          }

          CurPageRec4Dtl++;
          printDetail();
        }

        totAccountValue +=totNetUsage;


        if(masterRow[mCardprimary].ToString() == "Y" && (masterRow[mCardstate].ToString() == "Given" || masterRow[mCardstate].ToString() == "Embossed" || masterRow[mCardstate].ToString() == "New"  || masterRow[mCardstate].ToString() == "New Pin Generated Only"  ))
        {
          pageNo=1; CurPageRec4Dtl=0; //if pages is based on account
        }
        totNetUsage = 0;
        PrevCardNumber = curCardNumber; //masterRow[mCardno].ToString();
      }
      clsBasFile.deleteFile(strFileBasic);
      wb.Close(true, strFileBasic, true);
      wb = null;

    }
    catch(OracleException ex)
    {
      clsDbOracleLayer.catchError(ex);
    }
    catch(NotSupportedException ex)  //(Exception ex)  //
    {
      clsBasErrors.catchError(ex);
    }
    finally 
    {
      //if (preExit)
      //{

        // Close output ile
        //streamWritBasic.Flush();
        //streamWritBasic.Close();
        //fileStrmBasic.Close();
        //streamWritTrans.Flush();
        //streamWritTrans.Close();
        //fileStrmTrans.Close();

        ArrayList aryLstFiles = new ArrayList();
        aryLstFiles.Add(@strFileBasic);
        //aryLstFiles.Add(@strFileTrans);
        clsBasFile.generateFileMD5(aryLstFiles, @clsBasFile.getPathWithoutExtn(pStrFileName) + "_Raw" + ".MD5");
        aryLstFiles.Add(@clsBasFile.getPathWithoutExtn(pStrFileName) + "_Raw" + ".MD5");
        SharpZip zip = new SharpZip();
        zip.createZip(aryLstFiles, @clsBasFile.getPathWithoutExtn(pStrFileName) + "_Excel" + ".zip", "");
      //}

    }

  }



  protected virtual void printHeader()
  {
    mstrCnt++;
    extAccNum = clsBasValid.validateStr(masterRow[mExternalno]);
    if (extAccNum.Trim() == "")
      extAccNum = clsBasValid.validateStr(masterRow[mAccountno]);

    pageNo=1;
    //wsMain.get_Range(basText.Num2Chr(numOfColumn) + numOfLine.ToString(), basText.Num2Chr(numOfColumn) + numOfLine.ToString()).Value2 = masterRow[mCustomername].ToString();
      //+ strFileSpr + masterRow[mCardno].ToString() 
    wsMain.get_Range("A" + mstrCnt.ToString(), "A" + mstrCnt.ToString()).Value2 = masterRow[mCustomername].ToString();
    wsMain.get_Range("B" + mstrCnt.ToString(), "B" + mstrCnt.ToString()).Value2 = ValidateArbic(masterRow[mCustomeraddress1].ToString()) ;
    wsMain.get_Range("C" + mstrCnt.ToString(), "C" + mstrCnt.ToString()).Value2 = ValidateArbic(masterRow[mCustomeraddress2].ToString()); 
    wsMain.get_Range("D" + mstrCnt.ToString(), "D" + mstrCnt.ToString()).Value2 = ValidateArbic(masterRow[mCustomeraddress3].ToString()); 
    wsMain.get_Range("E" + mstrCnt.ToString(), "E" + mstrCnt.ToString()).Value2 = masterRow[mCustomerregion].ToString() + masterRow[mCustomercity].ToString(); 
    wsMain.get_Range("F" + mstrCnt.ToString(), "F" + mstrCnt.ToString()).Value2 = masterRow[mContracttype].ToString(); 
    wsMain.get_Range("G" + mstrCnt.ToString(), "G" + mstrCnt.ToString()).Value2 = masterRow[mAccountno].ToString(); 
    wsMain.get_Range("H" + mstrCnt.ToString(), "H" + mstrCnt.ToString()).Value2 = masterRow[mCardbranchpart].ToString()+ "  " + masterRow[mCardbranchpartname].ToString(); 
    wsMain.get_Range("I" + mstrCnt.ToString(), "I" + mstrCnt.ToString()).Value2 = String.Format("{0,8:dd/MM/yy}",masterRow[mStatementdateto]); 
    wsMain.get_Range("J" + mstrCnt.ToString(), "J" + mstrCnt.ToString()).Value2 = "'" + basText.formatCardNumber(curMainCard); 
    wsMain.get_Range("K" + mstrCnt.ToString(), "K" + mstrCnt.ToString()).Value2 = masterRow[mAccountlim].ToString(); 
    wsMain.get_Range("L" + mstrCnt.ToString(), "L" + mstrCnt.ToString()).Value2 = masterRow[mAccountavailablelim].ToString(); 
    wsMain.get_Range("M" + mstrCnt.ToString(), "M" + mstrCnt.ToString()).Value2 = masterRow[mMindueamount].ToString(); 
    wsMain.get_Range("N" + mstrCnt.ToString(), "N" + mstrCnt.ToString()).Value2 = masterRow[mClosingbalance].ToString() + CrDb(Convert.ToDecimal(masterRow[mClosingbalance].ToString())); 
    wsMain.get_Range("O" + mstrCnt.ToString(), "O" + mstrCnt.ToString()).Value2 = basText.formatDate(masterRow[mStetementduedate],"dd/MM/yy"); 
    wsMain.get_Range("P" + mstrCnt.ToString(), "P" + mstrCnt.ToString()).Value2 = masterRow[mTotaloverdueamount].ToString(); 
    wsMain.get_Range("Q" + mstrCnt.ToString(), "Q" + mstrCnt.ToString()).Value2 = basText.formatNum(masterRow[mOpeningbalance],"#,##0.00",12) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance].ToString())); 
    wsMain.get_Range("R" + mstrCnt.ToString(), "R" + mstrCnt.ToString()).Value2 = basText.formatNum(masterRow[mTotalpayments],"#,##0.00",12) + DbCr(Convert.ToDecimal(masterRow[mTotalpayments])); 
    wsMain.get_Range("S" + mstrCnt.ToString(), "S" + mstrCnt.ToString()).Value2 = basText.formatNum(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal]),"#,##0.00",12) + CrDbMinus(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal])); 
    wsMain.get_Range("T" + mstrCnt.ToString(), "T" + mstrCnt.ToString()).Value2 = basText.formatNum(Convert.ToDecimal(masterRow[mTotalcharges]) + Convert.ToDecimal(masterRow[mTotalinterest]),"#,##0.00",12) + CrDbMinus(Convert.ToDecimal(masterRow[mTotalcharges]) + Convert.ToDecimal(masterRow[mTotalinterest]));
    wsMain.get_Range("U" + mstrCnt.ToString(), "U" + mstrCnt.ToString()).Value2 = String.Format("Page {0} of {1}",pageNo,totAccCards); 
    wsMain.get_Range("V" + mstrCnt.ToString(), "V" + mstrCnt.ToString()).Value2 = extAccNum; 

      //masterRow[mExternalno].ToString()//ACCOUNTTYPE// + "" + strFileSpr
      //+ basText.formatNum(0.00,"#,##0.00",12) + "#" 
      //+ basText.formatNum(masterRow[mClosingbalance],"#,##0.00",12)  
      //+ "" + String.Format("Page {0} of {1}",pageNo,totalAccPages) + "" + strFileSpr + ""); //ACCOUNTTYPE// + "" + strFileSpr
  }


  protected virtual void printDetail()
  {
    dtlCnt++;
    DateTime trnsDate = clsBasValid.validateDate(detailRow[dTransdate]),postingDate = clsBasValid.validateDate(detailRow[dPostingdate]);
    if(trnsDate > postingDate)
      trnsDate= postingDate;
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
    //streamWritTrans.WriteLine("{0}" + strFileSpr + "{1:dd/MM}" + strFileSpr + "{2:dd/MM}" + strFileSpr + "{3,13}" + strFileSpr + "{4,-40}" + strFileSpr + "{5,12} {6,3}" + strFileSpr + "{7,12} {8,2}" + strFileSpr + "",masterRow[mCardno],trnsDate,postingDate,basText.trimStr(detailRow[dRefereneno],13),basText.trimStr(detailRow[dTrandescription].ToString().Trim() + " " + detailRow[dMerchant].ToString().Trim(),40),basText.formatNum(detailRow[dOrigtranamount],"#,##0.00;(#,##0.00)"),basText.Replace(clsBasValid.validateStr(detailRow[dOrigtrancurrency]),"XXX","   "),basText.formatNum(detailRow[dBilltranamount],"#,##0.00;(#,##0.00)"),CrDbDetail);//detailRow[dBilltranamountsign]//clsBasValid.validateDate(detailRow[dTransdate])//clsBasValid.validateDate(detailRow[dPostingdate])
    wsDtl.get_Range("A" + dtlCnt.ToString(), "A" + dtlCnt.ToString()).Value2 = "'" + masterRow[mCardno]; // masterRow[mCardno].ToString()
    wsDtl.get_Range("B" + dtlCnt.ToString(), "B" + dtlCnt.ToString()).Value2 = String.Format("{0,10:dd/MM/yyyy}",(Object)trnsDate);
    wsDtl.get_Range("C" + dtlCnt.ToString(), "C" + dtlCnt.ToString()).Value2 = String.Format("{0,10:dd/MM/yyyy}",(Object)postingDate);
    wsDtl.get_Range("D" + dtlCnt.ToString(), "D" + dtlCnt.ToString()).Value2 = "'" + basText.trimStr(detailRow[dRefereneno], 25);
    wsDtl.get_Range("E" + dtlCnt.ToString(), "E" + dtlCnt.ToString()).Value2 = trnsDesc; //basText.trimStr(detailRow[dTrandescription].ToString().Trim() + " " + detailRow[dMerchant].ToString().Trim(), 40);
    wsDtl.get_Range("F" + dtlCnt.ToString(), "F" + dtlCnt.ToString()).Value2 = strForeignCurr;  // basText.formatNum(detailRow[dOrigtranamount], "#,##0.00;(#,##0.00)") + " " + basText.Replace(clsBasValid.validateStr(detailRow[dOrigtrancurrency]), "XXX", "   ");
    wsDtl.get_Range("G" + dtlCnt.ToString(), "G" + dtlCnt.ToString()).Value2 = basText.formatNum(detailRow[dBilltranamount],"#,##0.00;(#,##0.00)") + " " + CrDbDetail;
  }


  //protected void printCardFooter()
  //{
    //streamWritSubtotal.WriteLine(masterRow[mCardno] + "" + strFileSpr + "" + basText.formatNum(masterRow[mClosingbalance],"#,##0.00",12) + "" + strFileSpr + ""); //" + strFileSpr + "
  //}

  protected void calcCardlRows()
  {
    totalCardPages = 0;
    totCardRows = 0;
    foreach (DataRow dtRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
    {
      if((dtRow[dPostingdate]== DBNull.Value) && (dtRow[dDocno]== DBNull.Value))
        continue ;// Exclude On-Hold Transactions 
      //streamWrit.WriteLine(basText.trimStr(dtRow[dTrandescription],40)); 
      totCardRows++;
    }
    totalCardPages = totCardRows / MaxDetailInPage;
    if((totCardRows % MaxDetailInPage) > 0)
      totalCardPages++;
    if(totalCardPages < 1)
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

  public clsStatExcel(string pStrFileName)
  {
    try 
    {
      strFileBasic = clsBasFile.getPathWithoutExtn(pStrFileName) + ".xls";
      clsBasFile.deleteFile(strFileBasic);
      //fileStrmBasic = new FileStream(strFileBasic, FileMode.Create);
      //streamWritBasic = new StreamWriter(fileStrmBasic);
      // open output file
      // Excel File
      xlApp = new Microsoft.Office.Interop.Excel.Application(); // Microsoft.Office.Interop.Excel.Application 

      if (xlApp == null)
      {
        Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
        return;
      }
      xlApp.Visible = false; //true
      wb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);//XlWBATemplate.xlWBATWorksheet
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
    }

  }

  public virtual bool addHeader(string pHeader)
  {
    bool rtrnVal = true;

    try
    {
      wsDtl = (Worksheet)wb.Worksheets[1];
      //wsDtl = (Worksheet)wb.Worksheets.Add(System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
      wsMain = (Worksheet)wb.Worksheets.Add(System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);

      if (wsMain == null)
      {
        Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
      }
      wsMain.Name = "Main Data";
      wsMain.Tab.Color = Convert.ToDecimal(ColorTranslator.ToOle(Color.RoyalBlue));
      //wsMain.Tab.ColorIndex = XlColorIndex.xlColorIndexAutomatic;
      strHeader = pHeader.Split(Convert.ToChar('|'));
      numOfLine = 1; numOfColumn = 0;
      foreach (string strHdr in strHeader)
      {
        aRange = wsMain.get_Range(basText.Num2Chr(numOfColumn) + numOfLine.ToString(), basText.Num2Chr(numOfColumn) + numOfLine.ToString());
        aRange.Value2 = strHdr;
        aRange.Font.Size = 11;
        aRange.Font.Bold = true;
        //aRange.Font.Background = Convert.ToDecimal(ColorTranslator.ToOle(Color.Beige)); 
        //aRange.Cells.FillLeft();
        //string ss = aRange.Justify() ; // .FillLeft();
        //aRange.FillLeft();
        aRange.Font.Color = Convert.ToDecimal(ColorTranslator.ToOle(Color.Red));//Color.RosyBrown 
        numOfColumn++;
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
    }
    return rtrnVal;
  }

  public virtual bool addRow(string pHeader)
  {
    bool rtrnVal = true;

    try
    {
      strHeader = pHeader.Split(Convert.ToChar('|'));
      numOfLine++; numOfColumn = 0;
      foreach (string strHdr in strHeader)
      {
        aRange = wsMain.get_Range(basText.Num2Chr(numOfColumn) + numOfLine.ToString(), basText.Num2Chr(numOfColumn) + numOfLine.ToString());
        aRange.Value2 = strHdr;
        aRange.Font.Size = 10;
        //aRange.Font.Bold = true;
        //aRange.Font.Background = Convert.ToDecimal(ColorTranslator.ToOle(Color.Beige)); 
        //aRange.Cells.FillLeft();
        //string ss = aRange.Justify() ; // .FillLeft();
        //aRange.FillLeft();
        //aRange.Font.Color = Convert.ToDecimal(ColorTranslator.ToOle(Color.Red));//Color.RosyBrown 
        numOfColumn++;
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
    }
    return rtrnVal;
  }

  public bool closeExcel()
  {
    bool rtrnVal = true;

    clsBasFile.deleteFile(strFileBasic);
    wb.Close(true, strFileBasic, true);
    wb = null;
    xlApp.Workbooks.Close();
    xlApp = null;

    return rtrnVal;
  }

  protected void makeZip()
  {
  }

  public string bankName
  {
    get { return strBankName; }
    set { strBankName = value; }
  }// bankName

  ~clsStatExcel()
  {
    //DSstatementRaw.Dispose();
  }
}
