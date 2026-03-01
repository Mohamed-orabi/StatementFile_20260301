using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Xml;
using System.Collections;


//Branch 10
public class clsStatementBNP : clsBasStatement
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
  private int MaxDetailInPage = 20; //20
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
  private DateTime vCurDate;
  private clsOMR omr = new clsOMR();
  private int totPages;
  string endOfCustomer = string.Empty;
  string cProduct = string.Empty, curFileName = string.Empty;
  private ArrayList aryLstFiles;
  private frmStatementFile frmMain;
  private int totRec = 1;
  private bool createCorporateVal = false;
  private string accountNoName = mAccountno;
  private string accountLimit = mAccountlim;
  private string accountAvailableLimit = mAccountavailablelim;
  //private bool isCorporate;
  private string strFileNam, stmntType;
  private string CurCardAddres1, CurCardAddres2, CurCardAddres3, CurCustomername;
  private DataRow emailRow;

  public clsStatementBNP()
  {
  }

  public string Statement(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pCurDate, string pStmntType, bool pAppendData)
  {
    string rtrnStr = "Successfully Generate " + pBankName, ppStrFileName = pStrFileName;
    int curMonth = pCurDate.Month;
    bool preExit = true;
    strFileNam = pStrFileName;
    stmntType = pStmntType;

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
      if (createCorporateVal)
      {
        isCorporateVal = true;
        accountNoName = mCardaccountno;
        accountLimit = mCardlimit;
        accountAvailableLimit = mCardavailablelimit;
      }
      //strOrder = " m.cardproduct," + strOrder;
      strOrder = strOrder;
      // data retrieve
      FillStatementDataSet(pBankCode); //DSstatement =  //10); //3
      getClientEmail(pBankCode);

      fileSummaryName = pStrFileName;
      fileSummaryName = clsBasFile.getPathWithoutExtn(fileSummaryName) +
        "_Summary." + clsBasFile.getFileExtn(fileSummaryName);
      //if (createCorporateVal)
      //{
      //  string strFileCompany = clsBasFile.getPathWithoutExtn(strOutputFile) +
      //    "_Company." + clsBasFile.getFileExtn(strOutputFile);

      //  aryLstFiles.Add(strFileCompany);
      //  clsStatement_CommonCorpCmpny stmntBNPCorp = new clsStatement_CommonCorpCmpny();// + "NSGB_Business_Statement.txt"
      //  //stmntBNPCorp.Statement(txtFileName.Text, strStatementType, bankCode, strFileName, stmntDate, stmntType, appendData);// + "NSGB_Business_Statement.txt"
      //  stmntBNPCorp.StatementCorp(strFileCompany, fileSummaryName, DSstatement);
      //  stmntBNPCorp = null;
      //}
      // Error file
      fileStrmErr = new FileStream(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName), FileMode.Create);
      strmWriteErr = new StreamWriter(fileStrmErr, Encoding.Default);

      // open Summary file
      fileSummary = new FileStream(fileSummaryName, FileMode.Create); //Create
      streamSummary = new StreamWriter(fileSummary, Encoding.Default);
    
      // raw data file
      fileRawData = new FileStream(clsBasFile.getPathWithoutExtn(pStrFileName) + "_RawData.txt", FileMode.Create);
      strmRawData = new StreamWriter(fileRawData, Encoding.Default);
      strmRawData.WriteLine("Account No|Customer Name|Customer Address|Phone|Mobile Phone|");  //masterRow[mCardno].ToString()

      pageNo = 1; totalCardPages = 0;
      curCardNo = String.Empty;
      curAccountNo = String.Empty;

      frmMain.BeginInvoke(frmMain.setMinMaxProgressDelegate, new object[] { DSstatement.Tables["tStatementMasterTable"].Rows.Count });
      foreach (DataRow mRow in DSstatement.Tables["tStatementMasterTable"].Rows)
      {
        frmMain.BeginInvoke(frmMain.setProgressDelegate, new object[] { totRec++ });
        //SetProgressDelegate setProgress = new SetProgressDelegate(frmStatementFile.SetProgress);
        //BeginInvoke(setProgress, new object[] { 1 });
        //setProgress(1);

        masterRow = mRow;
        //streamWrit.WriteLine(masterRow[mStatementno].ToString());
        //pageNo=1; CurPageRec4Dtl=0; CurrentPageFlag = "F 1"; //if page is based on card no
        cardsRows = mRow.GetChildRows(StatementNoDRel); //, DataRowVersion.Proposed

        strCardNo = masterRow[mCardno].ToString().Trim();
        if ( (clsCheckCard.CalcLuhnCheckDigit(strCardNo.Substring(0, 15)) != strCardNo.Substring(15, 1))) //!clsCheckCard.isValidCard(strCardNo) ||
        {
          strmWriteErr.WriteLine("Invalid Card Number" + strCardNo);
          numOfErr++;
        }

        //check product
        if (cProduct != masterRow[mCardproduct].ToString().Trim())
        {
          if (cProduct != string.Empty)
          {
            streamWrit.Flush();
            streamWrit.Close();
            fileStrm.Close();
          }
          cProduct = masterRow[mCardproduct].ToString().Trim();
          curFileName = clsBasFile.getPathWithoutExtn(pStrFileName)
            + "_" + cProduct + "." + clsBasFile.getFileExtn(pStrFileName);
          //curFileName = pStrFileName;
          add2FileList(curFileName);
          fileStrm = new FileStream(curFileName, FileMode.Append); //FileMode.Create Create
          streamWrit = new StreamWriter(fileStrm, Encoding.Default); //Encoding.GetEncoding("IBM420") IBM864  ASMO-708   Encoding.Default
        }

        if (strCardNo.Length != 16)
        {
          strmWriteErr.WriteLine("CardNo Length not equal 16 for AccountNo " + masterRow[accountNoName].ToString().Trim() + " Card No =" + strCardNo);
          //numOfErr++;
          continue;// Exclude Zero Length Cards 
        }

        if (cProduct == string.Empty || cProduct == "")
        {
          strmWriteErr.WriteLine("Invalid Product for Account " + masterRow[mAccountno].ToString() + " have Invalid Card Number " + strCardNo);
          numOfErr++;
        }

        strPrimaryCardNo = strCardNo;
        if (masterRow[mCardprimary].ToString() == "N")
        {
          strPrimaryCardNo = masterRow[mPrinarycardno].ToString();
          //calcCardlRows();
        }

        //start new account
        if (prevAccountNo != masterRow[accountNoName].ToString())
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

          
          
          
          if (totAccRows < 1
            && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0) //             || (masterRow[mCardno].ToString() == curMainCard   // Convert.ToDecimal(
          {
            isHaveF3 = true;
            //pageNo=1; totalAccPages =1;
            continue;
          }

          
          prevAccountNo = masterRow[accountNoName].ToString();
          //pageNo=1; //if page is based on account no
          printHeader();//if page is based on account no

          totNoOfCardStat++;

          DataRow[] emailRows;
          emailRows = DSemails.Tables["Emails"].Select("idclient = " + masterRow[mClientid].ToString());
          for (int i = 0; i < emailRows.Length; i++)
          {
            emailRow = emailRows[i];
            break;
          }
          strmRawData.WriteLine(extAccNum + "|" + CurCustomername + " " + "|" + CurCardAddres1 + " " + CurCardAddres2 + " " + CurCardAddres3 + "|" + emailRow["phone"].ToString() + "|" + emailRow["mobilephone"].ToString() + "|");  //masterRow[mCardno].ToString()

          if (Convert.ToDecimal(masterRow[mAccountlim].ToString()) <= Convert.ToDecimal(0))
          {
            strmWriteErr.WriteLine("Account Limit Less than or Equal Zero for Account " + masterRow[mAccountno].ToString());// + " and Card Number " + strCardNo
            numOfErr++;
          }
        } // End of if(prevAccountNo != masterRow[accountNoName].ToString())
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
      //clsBasStatementRawData RawData = new clsBasStatementRawData(clsBasFile.getPathWithoutExtn(pStrFileName), DSstatement);
      //fillStatementHistory(pStmntType,pAppendData);
      //>clsBasStatementRawData RawData = new clsBasStatementRawData(clsBasFile.getPathWithoutExtn(pStrFileName), DSstatement);
      //>clsBasStatementRawExcel RawExcel = new clsBasStatementRawExcel(clsBasFile.getPathWithoutExtn(pStrFileName), DSstatement);
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

        strmRawData.Flush();
        strmRawData.Close();
        fileRawData.Close();

        if (numOfErr == 0)
          clsBasFile.deleteFile(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName));
        else
        {
          rtrnStr += string.Format(" with {0} Errors", numOfErr);
          aryLstFiles.Add(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err.txt");
        }
        /*else
          clsBasFile.deleteFile(pStrFileName);*/
        //ArrayList aryLstFiles = new ArrayList();
        //aryLstFiles.Add("");
        //aryLstFiles.Add(@strOutputFile);
        numOfErr += validateNoOfLines(aryLstFiles, 48);
        aryLstFiles.Add(fileSummaryName);
        aryLstFiles.Add(clsBasFile.getPathWithoutExtn(pStrFileName) + "_RawData.txt");
        ///
        if (pStmntType == "Corporate")
        {
          ArrayList aryLstFilesCorp = new ArrayList();
          clsStatement_CommonCorpCmpny stmntBNPCorp = new clsStatement_CommonCorpCmpny();// + "NSGB_Business_Statement.txt"
          stmntBNPCorp.setFrm = frmMain;
          aryLstFilesCorp = stmntBNPCorp.Statement(ppStrFileName, pBankName, pBankCode, pStrFile, pCurDate);// + "NSGB_Business_Statement.txt"
          foreach (object str in aryLstFilesCorp)
            aryLstFiles.Add((string)str);

          stmntBNPCorp = null;
        }
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


  protected void printHeader()
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
    streamWrit.WriteLine(basText.alignmentLeft(masterRow[mCustomername], 50));  //
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
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    //streamWrit.WriteLine(omr.Asterisk(pageNo));//String.Empty
    streamWrit.WriteLine(omr.Asterisk(totPages) + omr.AsteriskPage4(pageNo));//String.Empty
    streamWrit.WriteLine(String.Empty);
    //streamWrit.WriteLine(basText.alignmentMiddle(basText.formatCardNumber(curMainCard), 20) + basText.alignmentMiddle(masterRow[accountLimit], 13) + basText.formatNum(masterRow[accountAvailableLimit], "##0", 20) + basText.alignmentMiddle(masterRow[mMindueamount], 13) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M") + basText.formatNum(masterRow[mTotaloverdueamount], "##0.00", 13)); //+ basText.formatNum(masterRow[accountLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mTotalcredits],"#,##0.00",9)+" "+basText.formatNum(masterRow[accountAvailableLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mMindueamount],"#,##0.00",15)+" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",15)+" "+basText.formatDate(masterRow[mStetementduedate],"dd/MM/yy") //TOTALCREDITS// basText.alignmentMiddle(masterRow[accountAvailableLimit],20) //"#,##0"
    streamWrit.WriteLine(basText.alignmentMiddle(basText.formatCardNumber(curMainCard), 20) + basText.alignmentMiddle(masterRow[mAccountlim], 13) + basText.formatNum(masterRow[mAccountavailablelim], "##0", 20) + basText.alignmentMiddle(masterRow[mMindueamount], 13) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M") + basText.formatNum(masterRow[mTotaloverdueamount], "##0.00", 13)); //+ basText.formatNum(masterRow[accountLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mTotalcredits],"#,##0.00",9)+" "+basText.formatNum(masterRow[accountAvailableLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mMindueamount],"#,##0.00",15)+" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",15)+" "+basText.formatDate(masterRow[mStetementduedate],"dd/MM/yy") //TOTALCREDITS// basText.alignmentMiddle(masterRow[accountAvailableLimit],20) //"#,##0"
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


  protected void printDetail()
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

  protected void printCardFooter()
  {
    //completePageDetailRecords();
    streamWrit.WriteLine(String.Empty);
    if (pageNo == totalAccPages)
      streamWrit.WriteLine(basText.replicat(" ", 17) + basText.alignmentLeft("Current Balance", 63) + basText.formatNumUnSign(masterRow[mClosingbalance], "#0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])));
    //streamWrit.WriteLine(basText.replicat(" ", 80) + basText.formatNumUnSign(masterRow[mClosingbalance], "#0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])));
    else
      streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    //streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(basText.formatNum(masterRow[mTotaldebits], "#0.00;(#0.00)", 20, "L") + basText.formatNum(masterRow[mTotalcredits], "#0.00;(#0.00)", 20, "L"));//streamWrit.WriteLine(String.Empty);
    //streamWrit.WriteLine(basText.alignmentRight(basText.formatCardNumber(curMainCard), 35) + basText.alignmentRight(masterRow[mMindueamount], 20) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "R")
    streamWrit.WriteLine(basText.alignmentRight(basText.formatCardNumber(curMainCard), 35) + basText.alignmentRight(basText.formatNumUnSign(masterRow[mMindueamount], "0.00", 20), 20) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "R")
   + basText.alignmentRight(basText.formatNumUnSign(masterRow[mOpeningbalance], "#,###,##0.00", 12) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance].ToString())), 15)
   + basText.alignmentRight(basText.formatNumUnSign(masterRow[mTotalpayments], "#,###,##0.00", 12) + " " + DbCr(Convert.ToDecimal(masterRow[mTotalpayments])), 15)
   + basText.alignmentRight(basText.formatNum(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal]), "#,###,##0.00", 12) + CrDbMinus(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal])), 15)
   + basText.alignmentRight(basText.formatNumUnSign(masterRow[mTotalcharges].ToString(), "#,###,##0.00", 12) + " " + CrDbMinus(Convert.ToDecimal(masterRow[mTotalcharges])), 15)
   + basText.alignmentRight(basText.formatNumUnSign(masterRow[mTotalinterest].ToString(), "#,###,##0.00", 12) + " " + CrDbMinus(Convert.ToDecimal(masterRow[mTotalinterest])), 15)
      //       + basText.formatNum(Convert.ToDecimal(masterRow[mTotalcharges]) + Convert.ToDecimal(masterRow[mTotalinterest]),"#,##0.00",12) + CrDbMinus(Convert.ToDecimal(masterRow[mTotalcharges]) + Convert.ToDecimal(masterRow[mTotalinterest])) 

   + basText.alignmentRight(basText.formatNumUnSign(masterRow[mClosingbalance], "#,###,##0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])), 15));
  }

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
    stmNo = masterRow[accountNoName].ToString();

    accountRows = DSstatement.Tables["tStatementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[accountNoName]).ToString().Trim() + "'");
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
    mainRows = DSstatement.Tables["tStatementMasterTable"].Select(accountNoName + " = '" + clsBasValid.validateStr(masterRow[accountNoName]) + "'");//ACCOUNTNO
    curMainCard = CurCardNo = "";
    foreach (DataRow mainRow in mainRows) //mRow.GetChildRows(StatementNoDRel)
    {
      totCrdNoInAcc++;
      CurCardNo = mainRow[mCardno].ToString();
      CurCardAddres1 = mainRow[mCustomeraddress1].ToString();
      CurCardAddres2 = mainRow[mCustomeraddress2].ToString();
      CurCardAddres3 = mainRow[mCustomeraddress3].ToString() + " " + mainRow[mCustomerregion].ToString() + " " + mainRow[mCustomercity].ToString();
      CurCustomername = mainRow[mCustomername].ToString();
      if (mainRow[mCardprimary].ToString() == "Y")
      {
        curMainCard = CurCardNo; //mainRow[mCardno].ToString();
        CurCardAddres1 = mainRow[mCustomeraddress1].ToString();
        CurCardAddres2 = mainRow[mCustomeraddress2].ToString();
        CurCardAddres3 = mainRow[mCustomeraddress3].ToString() + " " + mainRow[mCustomerregion].ToString() + " " + mainRow[mCustomercity].ToString();
        CurCustomername = mainRow[mCustomername].ToString();
      }
      if (mainRow[mCardprimary].ToString() == "Y" && isValidateCard(mainRow[mCardstate].ToString()))
      {
        curMainCard = CurCardNo; //mainRow[mCardno].ToString();
        CurCardAddres1 = mainRow[mCustomeraddress1].ToString();
        CurCardAddres2 = mainRow[mCustomeraddress2].ToString();
        CurCardAddres3 = mainRow[mCustomeraddress3].ToString() + " " + mainRow[mCustomerregion].ToString() + " " + mainRow[mCustomercity].ToString();
        CurCustomername = mainRow[mCustomername].ToString();
        break;
      }
    }

    if (curMainCard == "")
      curMainCard = CurCardNo;
  }


  private void completePageDetailRecords()
  {
    //int curPageLine =CurPageRec4Dtl;
    if (pageNo == totalAccPages)
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

    clsStatementSummary StatSummary = new clsStatementSummary();
    StatSummary.BankCode = curBranchVal;
    StatSummary.BankName = strBankName;
    StatSummary.StatementDate = vCurDate;
    StatSummary.CreationDate = DateTime.Now;
    StatSummary.StatementProduct = stmntType; // strFileNam + "__" + 
    StatSummary.StatementType = "Text";
    StatSummary.NoOfStatements = totNoOfCardStat ;
    StatSummary.NoOfPages = totNoOfPageStat ;
    StatSummary.NoOfTransactions = totNoOfTransactions ;
    //StatSummary.InsertRecord();
    StatSummary = null;

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

  public bool CreateCorporate
  {
    get { return createCorporateVal; }
    set { createCorporateVal = value; }
  }// CreateCorporate

  public int maxDetailInPage
  {
    get { return MaxDetailInPage; }
    set { MaxDetailInPage = value; }
  }// maxDetailInPage


  ~clsStatementBNP()
  {
    DSstatement.Dispose();
  }
}
