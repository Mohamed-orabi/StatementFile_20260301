using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Xml;
using System.Collections;


public class clsStatHtmlSup : clsStatHtml
{
  protected string curSuppCard = string.Empty;
  protected DataRow[] tmpDtlRows;
  protected string strFileNam, stmntType;

  public clsStatHtmlSup()
  {
  }

  public override string Statement(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pCurDate, string pStmntType, bool pAppendData)
  {
    string rtrnStr = "Successfully Generate " + pBankName;
    int curMonth = pCurDate.Month;
    bool preExit = true;
    strFileNam = pStrFileName;
    stmntType = pStmntType;
    curMonth = pCurDate.Month;
    aryLstFiles = new ArrayList();
    try
    {
      clsMaintainData maintainData = new clsMaintainData();
      maintainData.notRward = false;
      //maintainData.matchCardBranch4Account(pBankCode);

      pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);
      vCurDate = pCurDate; //DateTime.Now.AddMonths(-1);
      strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "_" + pStmntType;
      //clsBasFile.deleteDirectory(strOutputPath);
      clsBasFile.createDirectory(strOutputPath);
      strEmailFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "_" + pStmntType + "\\" + pBankName + "_" + pStmntType + "_" + vCurDate.ToString("yyyyMM");//+ ".txt"
      strNoEmailFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "_" + pStmntType + "\\" + pBankName + "_" + "WithoutEmails" + "_" + vCurDate.ToString("yyyyMM");//+ ".txt"
      pStrFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "_" + pStmntType + "\\" + pBankName + "_" + pStmntType + "_" + vCurDate.ToString("yyyyMM") + ".txt";
      strBankName = pBankName;
      emailLabel = pBankName + " statement for " + vCurDate.ToString("MM/yyyy"); //"BAI statement for 02/2008"
      strOutputFile = pStrFileName;

      // open emails file
      fileEmails = new FileStream(strEmailFileName + ".txt", FileMode.Create); //Create
      streamEmails = new StreamWriter(fileEmails, Encoding.Default);
      streamEmails.WriteLine("AccountNumber" + "|" + "ClientID" + "|" + "Email" + "|" + "Date Time");
      streamEmails.AutoFlush = true;
      
      // open No emails file
      fileNoEmails = new FileStream(strNoEmailFileName + ".txt", FileMode.Create); //Create
      streamNoEmails = new StreamWriter(fileNoEmails, Encoding.Default);
      streamNoEmails.WriteLine("AccountNumber" + "|" + "ClientID" + "|" + "Email" + "|" + "Mobile Phone");
      streamNoEmails.AutoFlush = true;

      // open output file
      //>fileStrm = new FileStream(pStrFileName, FileMode.Create); //Create
      //>streamWrit = new StreamWriter(fileStrm, Encoding.Default);
      // Error file
      //-fileStrmErr = new FileStream(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName), FileMode.Create);
      //-strmWriteErr = new StreamWriter(fileStrmErr, Encoding.Default);

      // open Summary file
      fileSummaryName = pStrFileName;
      fileSummaryName = clsBasFile.getPathWithoutExtn(fileSummaryName) +
        "_Summary." + clsBasFile.getFileExtn(fileSummaryName);
      fileSummary = new FileStream(fileSummaryName, FileMode.Create); //Create
      streamSummary = new StreamWriter(fileSummary, Encoding.Default);


      // set branch for data
      curBranchVal = pBankCode; // 10; //3 = real   1 = test
      //     clsBasStatement.mainTableCond =  " substr(cardno,1,6) !='421192' "; //" cardproduct != 'Visa Business' ";
      //     clsBasStatement.supTableCond =  " substr(cardno,1,6) !='421192' "; 
      // data retrieve
      if (createCorporateVal)
      {
        isCorporateVal = true;
        accountNoName = mCardaccountno;
        accountLimit = mCardlimit;
        accountAvailableLimit = mCardavailablelimit;
      }

      if (isRewardVal)
      {
        maintainData.curRewardCond = rewardCond;
        //maintainData.fixReward(pBankCode, rewardCond);
        strMainTableCond = "m.contracttype != " + rewardCondVal;
        strSubTableCond = "d.trandescription != 'Calculated Points'";
        curRewardCond = rewardCond;
        getReward(pBankCode);
      }
      FillStatementDataSet(pBankCode); //DSstatement =  //10); //3
      getClientEmail(pBankCode);
      pageNo = 1; totalCardPages = 0;
      curCardNo = String.Empty;
      curAccountNo = String.Empty;
      frmMain.BeginInvoke(frmMain.setMinMaxProgressDelegate, new object[] { DSstatement.Tables["tStatementMasterTable"].Rows.Count });
      foreach (DataRow mRow in DSstatement.Tables["tStatementMasterTable"].Rows)
      {
        frmMain.BeginInvoke(frmMain.setProgressDelegate, new object[] { totRec++ });
        masterRow = mRow;
        //streamWrit.WriteLine(masterRow[mStatementno].ToString());
        //pageNo=1; CurPageRec4Dtl=0; CurrentPageFlag = "F 1"; //if page is based on card no
        //-cardsRows = mRow.GetChildRows(StatementNoDRel); //, DataRowVersion.Proposed

        strCardNo = masterRow[mCardno].ToString().Trim();

        if (strCardNo.Length != 16)
        {
          //-  strmWriteErr.WriteLine("CardNo Length not equal 16 for AccountNo " + masterRow[accountNoName].ToString().Trim() + " Card No =" + strCardNo);
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
        if (prevAccountNo != masterRow[accountNoName].ToString())
        {
          emailLabelTmp = emailLabel;
          if (!string.IsNullOrEmpty(emailStr.ToString()) && !string.IsNullOrEmpty(emailTo)) //emailStr != string.Empty && emailStr != null  emailTo != string.Empty && emailTo != null
            SendEmail(emailStr.ToString(), "", emailTo);
          if (!string.IsNullOrEmpty(emailStr.ToString()) && string.IsNullOrEmpty(emailTo)) //emailStr != string.Empty && emailStr != null  emailTo != string.Empty && emailTo != null
          {
            streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|Without Email");
            noOfWithoutEmails++;
            emailTo = strEmailFrom; //"statement_Program@emp-group.com";
            emailLabelTmp = emailLabel + " Acc:" + curAccountNumber + " Phone:" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!");
            SendEmail(emailStr.ToString(), "", emailTo);
          }

          //{
          //    //streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|Without Email");
          //    //noOfWithoutEmails++;
          //    emailTo = strEmailFrom; //"statement_Program@emp-group.com";
          //    emailLabelTmp = emailLabel + " Acc:" + curAccountNumber + " Phone:" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!");
          //    //SendEmail(emailStr.ToString(), "", emailTo);
          //    if (strEmailFrom.ToUpper().EndsWith("BK.RW"))
          //    {
          //        clsValidateEmail valdEmail = new clsValidateEmail();
          //        if (valdEmail.isValideEmail(emailTo) != "ValidEmail")
          //        {
          //            emailTo = strEmailFrom;
          //            //streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + emailLabelTmp + "|Without or bad Email and FW to " + emailTo);
          //            //noOfBadEmails++;
          //            SendEmail(emailStr.ToString(), "", emailTo);
          //            //noOfWithoutEmails++;

          //        }
          //    }
          //    else
          //    {
          //        streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|Without Email");
          //        noOfWithoutEmails++;

          //    }
          //}

          emailStr = new StringBuilder("");
          cardsRows = DSstatement.Tables["tStatementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[accountNoName]).ToString().Trim() + "'");
          //-if (pageNo != totalAccPages && prevAccountNo != "")// 
          //-{
          //>MessageBox.Show( "Error in Genrating Statement", "Please Call The Programmer", MessageBoxButtons.OK ,
          //>  MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1); //,	MessageBoxOptions.RightAlign
          //-  strmWriteErr.WriteLine("pageNo not equal totalAccPages : " + prevAccountNo);
          //-  numOfErr++;
          //-}

          curMainCard = string.Empty;
          //-if (!isHaveF3)//!isHaveF3  CurrentPageFlag != "F 2"
          //-{
          //-  strmWriteErr.WriteLine("Account not Closed : " + prevAccountNo);
          //-  numOfErr++;
          //-}
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
          emailStr = new StringBuilder("");
          
          prevAccountNo = masterRow[accountNoName].ToString();
          //pageNo=1; //if page is based on account no
          DataRow[] emailRows;
          emailRows = DSemails.Tables["Emails"].Select("idclient = " + masterRow[mClientid].ToString());
          for (int i = 0; i < emailRows.Length; i++)
          {
            emailTo = emailRows[i][1].ToString().Trim();
            emailRow = emailRows[i];
          }
          curAccountNumber = masterRow[accountNoName].ToString();
          curCardNumber = strPrimaryCardNo;
          curClientID = masterRow[mClientid].ToString();
          
          //emailTo = (DSemails.Tables["Emails"].Select("idclient = " + masterRow[mClientid].ToString)).;
          printHeader();//if page is based on account no

          totNoOfCardStat++;
        } // End of if(prevAccountNo != masterRow[accountNoName].ToString())
        else
        {
          continue;
        }

        //calcCardlRows();
        //if(totCardRows < 1)continue ;  //if pages is based on card
        //printHeader();//if pages is based on card


        foreach (DataRow dRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
        {
          detailRow = dRow;
          stmNo = detailRow[dStatementno].ToString();
          if ((detailRow[dPostingdate] == DBNull.Value) && (detailRow[dDocno] == DBNull.Value)) continue;// Exclude On-Hold Transactions 
          curAccRows++;
          //-if (CurPageRec4Dtl >= MaxDetailInPage)
          //-{
          //-  CurPageRec4Dtl = 0;
          //-  printCardFooter();
          //-  pageNo++;
          //-  printHeader();
          //-}
          //						totNetUsage +=clsBasValid.validateNum(detailRow[dBilltranamount]);
          totNetUsage = calculateCrDb(totNetUsage, clsBasValid.validateNum(detailRow[dBilltranamount]), clsBasValid.validateStr(detailRow[dBilltranamountsign]));
          CurPageRec4Dtl++;

          if (curSuppCard != detailRow[dCardno].ToString())
          {
            tmpDtlRows = DSstatement.Tables["tStatementMasterTable"].Select(mCardno + " = '" + detailRow[dCardno] + "'");
            if (tmpDtlRows.Length > 0)
            {
              if (tmpDtlRows[0][mCardprimary].ToString() == "N" && isValidateCard(tmpDtlRows[0][mCardstate].ToString()))
              {
                supplementaryCardSeparator();
                CurPageRec4Dtl++;
              }
            }
          }

          printDetail();
          curSuppCard = detailRow[dCardno].ToString();
          //if(!((detailRow[dPostingdate]== DBNull.Value) && (detailRow[dDocno]== DBNull.Value))) 

        } //end of detail foreach
        //printCardFooter();//if pages is based on card
        // if(masterRow[mCardprimary].ToString() == "Y" && (masterRow[mCardstate].ToString() == "Given" || masterRow[mCardstate].ToString() == "Embossed" || masterRow[mCardstate].ToString() == "New"))
        curCrdNoInAcc++;
        //-if ((curAccRows >= totAccRows && totAccRows != 0) || (totAccRows == 0 && curCrdNoInAcc == totCrdNoInAcc))
        //-{
        //-  completePageDetailRecords();
          printCardFooter();//if pages is based on account
          //printAccountFooter();
          CurPageRec4Dtl = 0; //>pageNo=1; if pages is based on account
          curAccRows = 0;
          strEmailFromTmp = emailLabelTmp = string.Empty;
        //-}
        //streamWrit.WriteLine(strEndOfPage);
        //pageNo=1; CurPageRec4Dtl=0; //if pages is based on card
        //completePageDetailRecords();
          //SendEmail(emailStr, "", "");
      } //end of Master foreach
      emailTo = emailTo.Trim();
      emailLabelTmp = emailLabel;
      if (!string.IsNullOrEmpty(emailStr.ToString()) && !string.IsNullOrEmpty(emailTo)) // emailStr != string.Empty && emailStr != null emailTo != string.Empty && emailTo != null
        SendEmail(emailStr.ToString(), "", emailTo);
      if (!string.IsNullOrEmpty(emailStr.ToString()) && string.IsNullOrEmpty(emailTo)) // emailStr != string.Empty && emailStr != null emailTo != string.Empty && emailTo != null
      {
        streamNoEmails.WriteLine(curAccountNumber + "|" + curClientID + "|" + emailTo + "|" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!") + "|Without Email");
        noOfWithoutEmails++;
        emailTo = strEmailFrom; //"statement_Program@emp-group.com";
        emailLabelTmp = emailLabel + " Acc:" + curAccountNumber + " Phone:" + (emailRow?["mobilephone"]?.ToString() ?? "no mobile number!");
        SendEmail(emailStr.ToString(), "", emailTo);
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

      printStatementSummary();

      // Close Summary File
      streamSummary.Flush();
      streamSummary.Close();
      fileSummary.Close();

      streamEmails.Flush();
      streamEmails.Close();
      fileEmails.Close();

      streamNoEmails.WriteLine(string.Empty);
      streamNoEmails.WriteLine(string.Empty);
      streamNoEmails.WriteLine("** Any Statement not have email will forward to email " + strEmailFrom);
      streamNoEmails.Flush();
      streamNoEmails.Close();
      fileNoEmails.Close();
      aryLstFiles.Add(strEmailFileName + ".txt");
      aryLstFiles.Add(strNoEmailFileName + ".txt");
      aryLstFiles.Add(fileSummaryName);
      clsBasFile.generateFileMD5(aryLstFiles, strEmailFileName + ".MD5");
      aryLstFiles.Add(strEmailFileName + ".MD5");
      SharpZip zip = new SharpZip();
      zip.createZip(aryLstFiles, @clsBasFile.getPathWithoutExtn(pStrFileName) + ".zip", "");

      DSstatement.Dispose();
    }
    return rtrnStr;
  }

  protected virtual void supplementaryCardSeparator()
  {
    emailStr.Append(@"<tr><td align=""middle"" width=""6%"" height=""20"">");
    emailStr.Append(MakeHeaderStr("     ", false, false));
    emailStr.Append(@"</td><td align=""middle"" width=""5%"" height=""20"">");
    emailStr.Append(MakeHeaderStr("     ", false, false));
    emailStr.Append(@"</td><td align=""middle"" width=""19%"" height=""20"">");
    emailStr.Append(MakeHeaderStr(basText.trimStr(" ", 24), false, false));
    emailStr.Append(@"</td><td width=""42%"" height=""20""><p align=""left"">");
    emailStr.Append(MakeHeaderStr(basText.trimStr(">> Supplementary Card : " + basText.formatCardNumber(detailRow[dCardno].ToString()), 40), false, false));
    emailStr.Append(@"</td><td align=""right"" width=""10%"" height=""20"">");
    emailStr.Append(" ");
    emailStr.Append(@"</td><td align=""right"" width=""11%"" height=""20"">");
    emailStr.Append(" ");
    emailStr.Append(@"</td><td align=""right"" width=""1%"" height=""20"">");
    emailStr.Append(" ");
    emailStr.Append(@"</td></tr>");

  }

  protected override void calcAccountRows()
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
    mainRows = DSstatement.Tables["tStatementMasterTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[accountNoName]) + "'");
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

  ~clsStatHtmlSup()
  {
  }
}
