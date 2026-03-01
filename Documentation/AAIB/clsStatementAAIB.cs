using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Data.OracleClient;
using System.Xml;
using System.Collections;
using System.Runtime.InteropServices;

// Branch 6
//[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
public class clsStatementAAIB : clsBasStatement
{
  private string strBankName;
  private string strFileName;
  private FileStream fileStrm, fileSummary;
  private StreamWriter streamWrit, streamSummary;
  //private DataSet DSstatement;
  //private OracleDataReader drPrimaryCards, drMaster,drDetail;
  private DataRow masterRow;
  private DataRow detailRow;
  private string strEndOfLine = "\u000D";  //+ "M" ^M
  //		private string strEndOfPage = "\u000C"  ;  //+ "\u000D"+ "M" ^L + ^M
  private string strEndOfPage = "\u000C";  //  + basText.replicat(" ",85) + "F 2"  + "\u000D"+ "M" ^L + ^M
  private const int MaxDetailInPage = 26; //13 transactions. 21
  private const int MaxDetailInLastPage = 26; //13 transactions. 21
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
  private string strForeignCurr, strForeignVal;
  private string stmNo;
  private int totNoOfCardStat, totNoOfPageStat;

  private int curCardRow = 0, curTotCardRows = 0, numOfErr = 0, totNoOfTransactions = 0;
  private bool isPrimaryOnly, isHaveF3 = true, isPrimaryOpened = false;
  private FileStream fileStrmErr;
  private StreamWriter strmWriteErr;
  private string curMainCard;

  private string extAccNum;
  private string prevBranch, curBranch;
  private int totCrdNoInAcc, curCrdNoInAcc;
  private string strOutputPath, strOutputFile, fileSummaryName;
  private DateTime vCurDate;
  private clsOMR omr = new clsOMR();
  private string holdStatus = string.Empty ;
  private ArrayList aryLstFiles;
  private string curFileName;

  public clsStatementAAIB()
  {
  }

  public string Statement(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pCurDate, string pStmntType, bool pAppendData)
  {
    string rtrnStr = string.Empty;
    aryLstFiles = new ArrayList();
    try
    {
      clsMaintainData maintainData = new clsMaintainData();
      maintainData.matchCardBranch4Account(pBankCode);
      
      pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);
      vCurDate = pCurDate; // DateTime.Now.AddMonths(-1);
      strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
      clsBasFile.deleteDirectory(pStrFileName + vCurDate.ToString("yyyyMM") + pBankName);
      clsBasFile.createDirectory(pStrFileName + vCurDate.ToString("yyyyMM") + pBankName);
      pStrFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM") + ".txt";
      strBankName = pBankName;
      strFileName = pStrFileName;

      strOutputFile = pStrFileName;
      // open output file
 //     fileStrm = new FileStream(pStrFileName, FileMode.Create); //Create
 //     streamWrit = new StreamWriter(fileStrm, Encoding.Default);

      // Error file
      fileStrmErr = new FileStream(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName), FileMode.Create);
      strmWriteErr = new StreamWriter(fileStrmErr, Encoding.Default);

      // open Summary file
      fileSummaryName = pStrFileName;
      fileSummaryName = clsBasFile.getPathWithoutExtn(fileSummaryName) +
        "_Summary." + clsBasFile.getFileExtn(fileSummaryName);
      fileSummary = new FileStream(fileSummaryName, FileMode.Create); //Create
      //>streamSummary = new StreamWriter(fileSummary, Encoding.UTF8);//Encoding.Default
      streamSummary = new StreamWriter(fileSummary, Encoding.Default);//28596  Encoding.GetEncoding("iso-8859-6") 

       
      // set branch for data
      curBranchVal = pBankCode; // 6; //6  = real   1 = test
      //     clsBasStatement.mainTableCond =  " substr(cardno,1,6) !='421192' "; //" cardproduct != 'Visa Business' ";
      //     clsBasStatement.supTableCond =  " substr(cardno,1,6) !='421192' "; 
      // data retrieve
      FillStatementDataSet(pBankCode); //DSstatement =  //6); // 6
      pageNo = 1; totalCardPages = 0;
      curCardNo = String.Empty;
      curAccountNo = String.Empty;

      foreach (DataRow mRow in DSstatement.Tables["statementMasterTable"].Rows)
      {
        masterRow = mRow;
        //streamWrit.WriteLine(masterRow["STATEMENTNO"].ToString());
        //pageNo=1; CurPageRec4Dtl=0; CurrentPageFlag = "F 1"; //if page is based on card no
        cardsRows = mRow.GetChildRows(StatementNoDRel); //, DataRowVersion.Proposed

        strCardNo = masterRow["cardno"].ToString().Trim();
        curBranch = masterRow["cardbranchpartname"].ToString().Substring(5).Trim();
        //curBranch = curBranch.Substring(curBranch.IndexOf(")")+1).Trim();
        holdStatus = masterRow["holstmt"].ToString().Trim();
        if (holdStatus == "Y")
          holdStatus = "_Hold";
        else 
          holdStatus = "";

        if (prevBranch != curBranch)
        {
          if (prevBranch != null)
          {
            streamWrit.Flush();
            streamWrit.Close();
            fileStrm.Close();
          }
          // open output file
          curFileName = clsBasFile.getPathWithoutExtn(pStrFileName) 
            + "_" + curBranch + holdStatus + "." + clsBasFile.getFileExtn(pStrFileName);
          add2FileList(curFileName);
          fileStrm = new FileStream(curFileName, FileMode.Append); //FileMode.Create Create
          //>streamWrit = new StreamWriter(fileStrm, Encoding.Default);
          streamWrit = new StreamWriter(fileStrm, Encoding.GetEncoding("iso-8859-6"));
          prevBranch = curBranch; // masterRow["cardbranchpartname"].ToString().Trim();
        }
        if (strCardNo.Length != 16)
        {
          strmWriteErr.WriteLine("CardNo Length not equal 16 AccountNo " + masterRow["AccountNo"].ToString().Trim() + " Card No =" + strCardNo);
          //numOfErr++;
          continue;// Exclude Zero Length Cards 
        }
        strPrimaryCardNo = strCardNo;
        if (masterRow["cardprimary"].ToString() == "N")
        {
          strPrimaryCardNo = masterRow["prinarycardno"].ToString();
          //calcCardlRows();
        }

        //start new account
        //>if (prevAccountNo == masterRow["ACCOUNTNO"].ToString())
        //>  continue;

        if (prevAccountNo != masterRow["ACCOUNTNO"].ToString())
        {
          if (pageNo != totalAccPages && prevAccountNo != "")// 
          {
            //>MessageBox.Show( "Error in Genrating Statement", "Please Call The Programmer", MessageBoxButtons.OK ,
            //>  MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1); //,	MessageBoxOptions.RightAlign
            strmWriteErr.WriteLine("pageNo not equal totalAccPages : " + prevAccountNo + " file name :" + curFileName);
            numOfErr++;
          }

          curMainCard = string.Empty;
          if (!isHaveF3)//!isHaveF3  CurrentPageFlag != "F 2"
          {
            strmWriteErr.WriteLine("Account not Closed : " + prevAccountNo + " file name :" + curFileName);
            numOfErr++;
          }
          //if (prevAccountNo != string.Empty && curAccRows != 26)//!isHaveF3  CurrentPageFlag != "F 2"
          //{
          //  strmWriteErr.WriteLine("Wrong No of Transactions on the page : " + prevAccountNo + " file name :" + curFileName);
          //  numOfErr++;
          //}
          isHaveF3 = false;

          pageNo = 1;//totalAccPages = 1 ; pageNo=1;
          CurPageRec4Dtl = 0; totNetUsage = 0; CurrentPageFlag = "F 1"; //if page is based on account no
          calcAccountRows();

          //>          if(totAccRows < 1
          //>            && Convert.ToDecimal(masterRow["closingbalance"].ToString()) == 0) //Convert.ToDecimal(
          //>            continue;
          if (totAccRows < 1
            && Convert.ToDecimal(masterRow["closingbalance"].ToString()) == 0) //             || (masterRow["cardno"].ToString() == curMainCard   // Convert.ToDecimal(
          {
            isHaveF3 = true;
            //pageNo=1; totalAccPages =1;
            continue;
          }

          //>if(totAccRows < 1)continue ;  //if pages is based on account
          prevAccountNo = masterRow["ACCOUNTNO"].ToString();
          //pageNo=1; //if page is based on account no
          printHeader();//if page is based on account no

          totNoOfCardStat++;
        } // End of if(prevAccountNo != masterRow["ACCOUNTNO"].ToString())
        //calcCardlRows();
        //if(totCardRows < 1)continue ;  //if pages is based on card
        //printHeader();//if pages is based on card

        foreach (DataRow dRow in accountRows)//cardsRows //mRow.GetChildRows(StatementNoDRel)
        {
          detailRow = dRow;
          stmNo = detailRow["STATEMENTNO"].ToString();
          if ((detailRow["POSTINGDATE"] == DBNull.Value) && (detailRow["DOCNO"] == DBNull.Value)) continue;// Exclude On-Hold Transactions 
          curAccRows += 2; // curAccRows++;
          if (CurPageRec4Dtl >= MaxDetailInPage)
          {
            CurPageRec4Dtl = 0;
            printCardFooter();
            pageNo++;
            printHeader();
          }
          //						totNetUsage +=clsBasValid.validateNum(detailRow["billtranamount"]);
          totNetUsage = calculateCrDb(totNetUsage, clsBasValid.validateNum(detailRow["billtranamount"]), clsBasValid.validateStr(detailRow["billtranamountsign"]));
          CurPageRec4Dtl +=2;
          printDetail();
          //if(!((detailRow["POSTINGDATE"]== DBNull.Value) && (detailRow["DOCNO"]== DBNull.Value))) 

        } //end of detail foreach
        //printCardFooter();//if pages is based on card
        // if(masterRow["cardprimary"].ToString() == "Y" && (masterRow["cardstate"].ToString() == "Given" || masterRow["cardstate"].ToString() == "Embossed" || masterRow["cardstate"].ToString() == "New"))
        curCrdNoInAcc++;
        if ((curAccRows >= totAccRows  ) )//&& totAccRows != 0|| (totAccRows == 0 && curCrdNoInAcc == totCrdNoInAcc)
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
      //if (curAccRows != 26)//!isHaveF3  CurrentPageFlag != "F 2"
      //{
      //  strmWriteErr.WriteLine("Wrong No of Transactions on the page : " + prevAccountNo + " file name :" + curFileName);
      //  numOfErr++;
      //}

      //clsBasStatementRawData RawData = new clsBasStatementRawData(clsBasFile.getPathWithoutExtn(pStrFileName), DSstatement);
      fillStatementHistory(pStmntType,pAppendData);
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

      printStatementSummary();

      // Close Summary File
      streamSummary.Flush();
      streamSummary.Close();
      fileSummary.Close();

      strmWriteErr.Flush();
      strmWriteErr.Close();
      fileStrmErr.Close();
      DSstatement.Clear();
      if (numOfErr == 0)
        clsBasFile.deleteFile(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName));
      /*      else
              clsBasFile.deleteFile(pStrFileName);*/

      //aryLstFiles = new ArrayList();
      //aryLstFiles.Add("");
      //aryLstFiles.Add(@strOutputFile);
      numOfErr = validateNoOfLines(aryLstFiles, 69);
      if (numOfErr > 0)
        rtrnStr += string.Format(" with {0} Errors", numOfErr);
      aryLstFiles.Add(@fileSummaryName);
      clsBasFile.generateFileMD5(aryLstFiles, @clsBasFile.getPathWithoutExtn(pStrFileName) + ".MD5");
      aryLstFiles.Add(@clsBasFile.getPathWithoutExtn(pStrFileName) + ".MD5");
      SharpZip zip = new SharpZip();
      zip.createZip(aryLstFiles, @clsBasFile.getPathWithoutExtn(pStrFileName) + ".zip", "");

      DSstatement.Dispose();
    }
      return rtrnStr;
  }


  protected void printHeader()
  {
    extAccNum = clsBasValid.validateStr(masterRow["externalno"]);
    if (extAccNum.Trim() == "" || extAccNum.Length < 16 )
      extAccNum = basText.alignmentLeft((object) extAccNum,16,'0');
      //extAccNum = clsBasValid.validateStr(masterRow["accountno"]);

    //if(masterRow["cardprimary"].ToString() == "Y")
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

    //			streamWrit.WriteLine(basText.alignmentLeft(masterRow["Customername"],50)+ basText.replicat(" ",31) + masterRow["cardbranchpartname"]);  //
    //			streamWrit.WriteLine(basText.alignmentMiddle(masterRow["customeraddress1"],50)+ basText.replicat(" ",31) + clsBasValid.validateStr(masterRow["accountno"])); //basText.formatNum(masterRow["cardno"],"####-####-####-####")
    //string x = Convert.ToString(ValidateArbic(masterRow["customeraddress1"].ToString()));
    //			streamWrit.WriteLine(basText.alignmentMiddle(masterRow["customeraddress2"],50)+ basText.replicat(" ",31)+"{0,10:dd/MM/yyyy}",masterRow["Statementdatefrom"]);//+   basText.formatNum(masterRow["cardlimit"],"########")
    //			streamWrit.WriteLine(basText.alignmentMiddle(masterRow["customeraddress3"].ToString().Trim() + " " + masterRow["CUSTOMERREGION"].ToString().Trim() + " " + masterRow["CUSTOMERCITY"].ToString().Trim(),50)+ basText.replicat(" ",31)+pageNo + " / " + totalAccPages);  //basText.formatNum(masterRow["accountno"],"##-##-######-####")
    streamWrit.WriteLine(strEndOfPage);//line 1
    //>streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty); //line 2
    streamWrit.WriteLine(String.Empty);//line 3
    streamWrit.WriteLine(String.Empty);//line 4
    streamWrit.WriteLine(String.Empty);//line 5
    streamWrit.WriteLine(String.Empty);//line 6
    streamWrit.WriteLine(basText.replicat(" ", 35)
      + basText.alignmentLeft((object)get3Branch(masterRow["cardbranchpartname"].ToString().Substring(5).Trim()), 3)
       + basText.replicat(" ", 6) 
       + basText.alignmentLeft(masterRow["barcode"].ToString(), 8));  //+"{0,5:MM/yy}",masterRow["Statementdatefrom"]
    streamWrit.WriteLine(basText.replicat(" ", 3) 
      + basText.alignmentLeft(masterRow["Customername"].ToString(), 46)
      + "  ACC NUM: "
      + basText.alignmentLeft(extAccNum.Substring(0, 6), 6)
      + " " + basText.alignmentLeft(extAccNum.Substring(6, 4), 4)  //.Substring(masterRow["customerno"].ToString().LastIndexOf('-'))
      + " " + basText.alignmentLeft(extAccNum.Substring(10, 3), 3)
      + " " + basText.alignmentLeft(extAccNum.Substring(13, 3), 3));  //+"{0,5:MM/yy}",masterRow["Statementdatefrom"]
    streamWrit.WriteLine(basText.replicat(" ", 31));//line 9
    streamWrit.WriteLine(basText.replicat(" ", 42)
      + basText.alignmentLeft("CARD LIMIT :", 12) 
      + basText.formatNum(masterRow["ACCOUNTLIM"], "##,###,##0.000",20,"R")
      + " " + basText.alignmentLeft(masterRow["accountcurrency"].ToString(), 3)); //line 10
    streamWrit.WriteLine(omr.dash1Line(pageNo, totalAccPages));//11
    //>streamWrit.WriteLine(basText.alignmentLeft(curMainCard, 20) 
    streamWrit.WriteLine(basText.alignmentLeft(basText.formatNum(curMainCard, "#### #### #### ####"), 20)
      + basText.replicat(" ", 10) + "PAGE "+ basText.alignmentRight(pageNo, 3) + " OF " 
      + basText.alignmentRight(totalAccPages, 3) 
      + basText.replicat(" ", 6) + "STATEMENT DATE:"
      + "  " + basText.formatDate(masterRow["statementdateto"], "yyyy/MM/dd"));
    streamWrit.WriteLine(String.Empty); //line 13
    streamWrit.WriteLine(basText.replicat(" ", 28) 
      + basText.alignmentLeft("(STAMP DUTY PAID)", 17));//line 14
    streamWrit.WriteLine(String.Empty); //line 15
    streamWrit.WriteLine(String.Empty); //line 16
    streamWrit.WriteLine(String.Empty); //line 17


    //streamWrit.WriteLine(basText.replicat(" ", 81) + CurrentPageFlag);  //+"{0,5:MM/yy}",masterRow["Statementdatefrom"]
    //streamWrit.WriteLine(basText.replicat(" ", 81) + basText.formatDate(masterRow["statementdateto"],"dd/MM/yyyy"));  //cardproduct+"{0,5:MM/yy}",masterRow["Statementdatefrom"]
    //streamWrit.WriteLine(basText.alignmentLeft(masterRow["Customername"], 50));  //
    //streamWrit.WriteLine(basText.replicat(" ", 81) + basText.formatNum(curMainCard, "#### #### #### ####"));  //cardbranchpartname
    //streamWrit.WriteLine(basText.alignmentLeft(ValidateArbic(masterRow["customeraddress1"].ToString()), 50)); //basText.formatNum(masterRow["cardno"],"####-####-####-####")
    //streamWrit.WriteLine(basText.replicat(" ", 81) + basText.formatNum(masterRow["ACCOUNTLIM"], "########")); //extAccNum   clsBasValid.validateStr(masterRow["externalno"]) //accountno  //basText.formatNum(masterRow["cardno"],"####-####-####-####")
    //streamWrit.WriteLine(basText.alignmentLeft(ValidateArbic(masterRow["customeraddress2"].ToString()), 50));
    ////>>>streamWrit.WriteLine(basText.replicat(" ", 81) + basText.alignmentLeft(masterRow["MINDUEAMOUNT"], 13));//"{0,10:dd/MM/yyyy}", masterRow["statementdateto"])   Statementdatefrom+ basText.formatNum(masterRow["ACCOUNTLIM"],"########")  basText.formatNum(masterRow["cardlimit"],"########")
    //streamWrit.WriteLine(omr.Asterisk(pageNo) + basText.replicat(" ", 79) + basText.alignmentLeft(masterRow["MINDUEAMOUNT"], 13));//"{0,10:dd/MM/yyyy}", masterRow["statementdateto"])   Statementdatefrom+ basText.formatNum(masterRow["ACCOUNTLIM"],"########")  basText.formatNum(masterRow["cardlimit"],"########")
    //streamWrit.WriteLine(basText.alignmentLeft(ValidateArbic(masterRow["customeraddress3"].ToString().Trim()) + " " + masterRow["CUSTOMERREGION"].ToString().Trim() + " " + masterRow["CUSTOMERCITY"].ToString().Trim(), 50));  //basText.formatNum(masterRow["accountno"],"##-##-######-####")
    ////>>    streamWrit.WriteLine(basText.replicat(" ", 81) + basText.alignmentLeft(pageNo,5) + basText.alignmentLeft(totalAccPages,5));  //basText.formatNum(masterRow["accountno"],"##-##-######-####")
    //streamWrit.WriteLine(" " + omr.fixLine() + basText.replicat(" ", 76) + basText.alignmentLeft(pageNo, 5) + basText.alignmentLeft(totalAccPages, 5));  //basText.formatNum(masterRow["accountno"],"##-##-######-####")
    //streamWrit.WriteLine(" " + omr.Line2LastPage(pageNo, totalAccPages));//String.Empty
    //streamWrit.WriteLine(" " + omr.Line3(pageNo));//String.Empty
    //streamWrit.WriteLine(" " + omr.Line4(pageNo));//String.Empty
    //streamWrit.WriteLine(" " + omr.fixLine());//String.Empty
    //streamWrit.WriteLine(basText.alignmentLeft(masterRow["barcode"].ToString(),30));//String.Empty
    ////    streamWrit.WriteLine(basText.alignmentMiddle(curMainCard, 20) + basText.alignmentMiddle(masterRow["accountlim"], 13) + basText.formatNum(masterRow["ACCOUNTAVAILABLELIM"], "##0", 20) + basText.alignmentMiddle(masterRow["MINDUEAMOUNT"], 13) + basText.formatDate(masterRow["stetementduedate"], "dd/MM/yyyy", 15, "M") + basText.formatNum(masterRow["totaloverdueamount"], "##0.00", 13)); //+ basText.formatNum(masterRow["ACCOUNTLIM"],"#,##0.00",9)+" "+basText.formatNum(masterRow["TOTALCREDITS"],"#,##0.00",9)+" "+basText.formatNum(masterRow["ACCOUNTAVAILABLELIM"],"#,##0.00",9)+" "+basText.formatNum(masterRow["MINDUEAMOUNT"],"#,##0.00",15)+" "+basText.formatNum(masterRow["CLOSINGBALANCE"],"#,##0.00",15)+" "+basText.formatDate(masterRow["stetementduedate"],"dd/MM/yy") //TOTALCREDITS// basText.alignmentMiddle(masterRow["ACCOUNTAVAILABLELIM"],20) //"#,##0"
    ////    streamWrit.WriteLine(String.Empty);
    ////    streamWrit.WriteLine(String.Empty);
    ////streamWrit.WriteLine(String.Empty);
    ////    streamWrit.WriteLine(String.Empty);
    ////    if (pageNo == 1)
    ////      streamWrit.WriteLine(basText.replicat(" ", 36) + basText.alignmentLeft("Previous Balance", 63) + basText.formatNumUnSign(masterRow["openingbalance"], "#,###0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow["openingbalance"])));  //+ CrDb(clsBasValid.validateStr(detailRow["billtranamountsign"]))
    ////    else
    ////      streamWrit.WriteLine(String.Empty);
    ////    streamWrit.WriteLine(String.Empty);

    totalPages++;
    totNoOfPageStat++;
  }


  protected void printDetail()
  {
    DateTime trnsDate = clsBasValid.validateDate(detailRow["transdate"]), postingDate = clsBasValid.validateDate(detailRow["postingdate"]);
    if (trnsDate > postingDate)
      trnsDate = postingDate;

    if (masterRow["accountcurrency"].ToString().Trim() != detailRow["origtrancurrency"].ToString().Trim())
    {
      //strForeignCurr = basText.formatNum(detailRow["origtranamount"], "#0.00;(#0.00)") + " " + basText.Replace(clsBasValid.validateStr(detailRow["origtrancurrency"]), "XXX", "   ");
      strForeignVal = basText.formatNum(detailRow["origtranamount"], "#0.000",18,"R") ;
      strForeignCurr = basText.Replace(basText.alignmentLeft(detailRow["origtrancurrency"],18), "XXX", "   ");
    }
    else
    {
      //strForeignCurr = basText.replicat(" ", 16);
      strForeignVal = basText.replicat(" ", 18);
      strForeignCurr = basText.replicat(" ", 18);
    }

    string trnsDesc; //= detailRow["trandescription"].ToString().Trim() + " " + detailRow["MERCHANT"].ToString().Trim()
    if (detailRow["MERCHANT"].ToString().Trim() == "")
      trnsDesc = detailRow["trandescription"].ToString().Trim();
    else
      trnsDesc = detailRow["MERCHANT"].ToString().Trim();

    ////			streamWrit.WriteLine("  {0:dd/MM}  {1:dd/MM}  {2,-40} {3,12} {4,3} {5,16} {6,2}",trnsDate,postingDate,basText.trimStr(detailRow["trandescription"].ToString().Trim() + " " + detailRow["MERCHANT"].ToString().Trim(),45),basText.formatNum(detailRow["origtranamount"],"#,##0.00;(#,##0.00)"),basText.Replace(clsBasValid.validateStr(detailRow["origtrancurrency"]),"XXX","   "),basText.formatNum(detailRow["billtranamount"],"#,##0.00;(#,##0.00)"),CrDb(clsBasValid.validateStr(detailRow["billtranamountsign"])) + isSupplementCard(clsBasValid.validateStr(masterRow["cardprimary"].ToString()))); //:f2//clsBasValid.validateDate(detailRow["transdate"]),clsBasValid.validateDate(detailRow["postingdate"]) //  {2,13} ,basText.trimStr(detailRow["refereneno"],12)
    ////    streamWrit.WriteLine("  {0:dd/MM}  {1:dd/MM}  {2,-24}  {3,-40} {4,16} {5,16} {6,2}", trnsDate, postingDate, basText.trimStr(detailRow["refereneno"].ToString().Trim(), 20), basText.trimStr(trnsDesc, 40), strForeignCurr, basText.formatNum(detailRow["billtranamount"], "#0.00;(#0.00)"), CrDb(clsBasValid.validateStr(detailRow["billtranamountsign"])) + " " + isSupplementCard(clsBasValid.validateStr(masterRow["cardprimary"].ToString()))); //:f2//clsBasValid.validateDate(detailRow["transdate"]),clsBasValid.validateDate(detailRow["postingdate"]) //  {2,13} ,basText.trimStr(detailRow["refereneno"],12)
    //streamWrit.WriteLine("  {0:dd/MM}  {1,-24}  {2,-40} {3,16} {4,16} {5,2}", trnsDate, basText.trimStr(detailRow["refereneno"].ToString().Trim(), 20), basText.trimStr(trnsDesc, 40), strForeignCurr, basText.formatNum(detailRow["billtranamount"], "#0.00;(#0.00)"), CrDb(clsBasValid.validateStr(detailRow["billtranamountsign"])) + " " + isSupplementCard(clsBasValid.validateStr(masterRow["cardprimary"].ToString()))); //:f2//clsBasValid.validateDate(detailRow["transdate"]),clsBasValid.validateDate(detailRow["postingdate"]) //  {2,13} ,basText.trimStr(detailRow["refereneno"],12)
    streamWrit.WriteLine(trnsDate.ToString("dd/MM") + basText.replicat(" ", 5) 
      + basText.alignmentLeft(detailRow["docno"].ToString(), 12) + "  "
      + basText.alignmentLeft(trnsDesc, 30) + "  "
      + basText.formatNum(detailRow["billtranamount"], "#0.000",19,"R"));
    totNoOfTransactions++;
    //if (detailRow["refereneno"].ToString().Trim() > 0)
    //{
      streamWrit.WriteLine(basText.replicat(" ", 6) 
        + basText.alignmentLeft(detailRow["refereneno"].ToString(), 23) + "  "
        + strForeignVal + "  "
        + strForeignCurr);
      totNoOfTransactions++;
    //}
  }

  protected void printCardFooter()
  {
    //completePageDetailRecords();
    //if (pageNo == totalAccPages)
    //  streamWrit.WriteLine(basText.replicat(" ", 17) + basText.alignmentLeft("Current Balance", 63) + basText.formatNumUnSign(masterRow["closingbalance"], "#0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow["closingbalance"])));
    //else
    //  streamWrit.WriteLine(String.Empty);
//    streamWrit.WriteLine(basText.alignmentRight(curMainCard, 35) + basText.alignmentRight(masterRow["mindueamount"], 20) + basText.formatDate(masterRow["stetementduedate"], "dd/MM/yyyy", 15, "R")
//   + basText.alignmentRight(basText.formatNumUnSign(masterRow["openingbalance"], "##0.00", 12) + " " + CrDb(Convert.ToDecimal(masterRow["OPENINGBALANCE"].ToString())), 15)
//   + basText.alignmentRight(basText.formatNum(masterRow["TOTALPAYMENTS"], "#,##0.00", 12) + " " + DbCr(Convert.ToDecimal(masterRow["TOTALPAYMENTS"])), 15)
//   + basText.alignmentRight(basText.formatNum(Convert.ToDecimal(masterRow["TOTALPURCHASES"]) + Convert.ToDecimal(masterRow["TOTALCASHWITHDRAWAL"]), "#,##0.00", 12) + CrDbMinus(Convert.ToDecimal(masterRow["TOTALPURCHASES"]) + Convert.ToDecimal(masterRow["TOTALCASHWITHDRAWAL"])), 15)
//   + basText.alignmentRight(basText.formatNum(masterRow["TOTALCHARGES"].ToString(), "#,##0.00", 12) + " " + CrDbMinus(Convert.ToDecimal(masterRow["TOTALCHARGES"])), 15)
//   + basText.alignmentRight(basText.formatNum(masterRow["totalinterest"].ToString(), "#,##0.00", 12) + " " + CrDbMinus(Convert.ToDecimal(masterRow["totalinterest"])), 15)
      //       + basText.formatNum(Convert.ToDecimal(masterRow["TOTALCHARGES"]) + Convert.ToDecimal(masterRow["totalinterest"]),"#,##0.00",12) + CrDbMinus(Convert.ToDecimal(masterRow["TOTALCHARGES"]) + Convert.ToDecimal(masterRow["totalinterest"])) 

    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(basText.alignmentRight(basText.formatNumUnSign(masterRow["openingbalance"], "#,###,##0.000", 15),15) //+ " "
      + basText.alignmentRight(basText.formatNum(masterRow["totaldebits"].ToString(), "#,###,##0.000", 20), 20) // + " " 
      + basText.alignmentRight(basText.formatNum(masterRow["totalcredits"].ToString(), "#,###,##0.000", 19), 19) // + "  "
      + basText.alignmentRight(basText.formatNumUnSign(masterRow["closingbalance"], "#,###,##0.000", 19), 19));
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(basText.replicat(" ", 47) 
      + basText.alignmentLeft(masterRow["cardbranchpartname"].ToString().Substring(5).Trim(), 20));
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(String.Empty);
    streamWrit.WriteLine(" " 
      + basText.alignmentLeft(masterRow["Customername"].ToString(), 46)
      + basText.alignmentRight("* MINIMUM PAYMENT DUE IN ", 26) + " "
      + basText.alignmentLeft(masterRow["accountcurrency"].ToString(), 3));
    streamWrit.WriteLine(basText.replicat(" ", 50) 
      + "DUE Date" + "  "
      + basText.formatDate(masterRow["stetementduedate"], "dd/MM/yyyy")); //line 60
    streamWrit.WriteLine(" " 
      + basText.alignmentLeft(ValidateArbic(masterRow["customeraddress1"].ToString()), 48)
      + basText.replicat("*", 8)
      + basText.formatNum(masterRow["mindueamount"], "#,###,##0.000;(#,###,##0.000)", 18)); //basText.formatNum(masterRow["cardno"],"####-####-####-####")
    streamWrit.WriteLine("  " 
      + basText.alignmentLeft(ValidateArbic(masterRow["customeraddress2"].ToString()), 47)
      + basText.replicat("*", 8));
    streamWrit.WriteLine(String.Empty); // line 63
    streamWrit.WriteLine(basText.replicat(" ", 67)
      + basText.formatDate(masterRow["statementdateto"], "yyyy/MM/dd")); // line 64
    streamWrit.WriteLine(basText.replicat(" ", 12) + "CARD HOLDER");  // line 65
    streamWrit.WriteLine(basText.replicat(" ", 10)
      + basText.alignmentLeft(masterRow["Customername"].ToString(), 20));
    streamWrit.WriteLine(String.Empty);
    //>streamWrit.WriteLine(String.Empty);
    if (masterRow["cardvip"].ToString() == "Y")
      streamWrit.WriteLine("   THANK YOU");
    else
      streamWrit.WriteLine("   PLEASE ENSURE PROMPT SETTLEMENT OF YOUR MIN PAYMENT TO AVOID CARD SUSPENSION");
    streamWrit.WriteLine(basText.replicat(" ", 57)
      + basText.alignmentLeft(basText.formatNum(curMainCard, "#### #### #### ####"), 20));


    //streamWrit.WriteLine(basText.alignmentRight(basText.formatNumUnSign(masterRow["openingbalance"], "#,###,##0.00", 12) + " " + CrDb(Convert.ToDecimal(masterRow["OPENINGBALANCE"].ToString())), 20) +
    //  basText.alignmentRight(basText.formatNum(masterRow["TOTALCHARGES"].ToString(), "#,###,##0.00", 12), 20) +
    //  basText.alignmentRight(basText.formatNumUnSign(Convert.ToDecimal(masterRow["TOTALPURCHASES"]) +
    //  Convert.ToDecimal(masterRow["TOTALCASHWITHDRAWAL"]), "#,###,##0.00",16) + " " +
    //  CrDbMinus(Convert.ToDecimal(masterRow["TOTALPURCHASES"]) +
    //  Convert.ToDecimal(masterRow["TOTALCASHWITHDRAWAL"])), 20) +
//basText.alignmentRight(basText.formatNumUnSign(masterRow["closingbalance"], "#,###,##0.00", 12) + " " +
    //CrDb(Convert.ToDecimal(masterRow["closingbalance"])), 20));
    //streamWrit.WriteLine(String.Empty);
    //streamWrit.WriteLine(String.Empty);
    //streamWrit.WriteLine(String.Empty);
    //streamWrit.WriteLine(String.Empty);
    //streamWrit.WriteLine(String.Empty);
    //streamWrit.WriteLine(basText.replicat(" ", 81) + basText.alignmentLeft(masterRow["externalno"],20));  //cardproduct+"{0,5:MM/yy}",masterRow["Statementdatefrom"]
    //streamWrit.WriteLine(basText.alignmentLeft(masterRow["Customername"], 50));  //
    //streamWrit.WriteLine(basText.replicat(" ", 81) + basText.formatNum(curMainCard, "#### #### #### ####"));  //cardbranchpartname
    //streamWrit.WriteLine(basText.alignmentLeft(ValidateArbic(masterRow["customeraddress1"].ToString()), 50)); //basText.formatNum(masterRow["cardno"],"####-####-####-####")
    //streamWrit.WriteLine(basText.alignmentLeft(ValidateArbic(masterRow["customeraddress2"].ToString()), 50));
    //streamWrit.WriteLine(basText.alignmentLeft(ValidateArbic(masterRow["customeraddress3"].ToString().Trim()) + " " + masterRow["CUSTOMERREGION"].ToString().Trim() + " " + masterRow["CUSTOMERCITY"].ToString().Trim(), 50));  //basText.formatNum(masterRow["accountno"],"##-##-######-####")
    //streamWrit.WriteLine(basText.replicat(" ", 81) + basText.alignmentLeft(basText.formatNumUnSign(masterRow["TOTALPAYMENTS"], "#,##0.00", 12) + " " + DbCr(Convert.ToDecimal(masterRow["TOTALPAYMENTS"])), 15));  //cardproduct+"{0,5:MM/yy}",masterRow["Statementdatefrom"]
    //streamWrit.WriteLine(String.Empty);

  }

  protected void printAccountFooter()
  {
    streamWrit.WriteLine("GRAND" + basText.formatNumUnSign(masterRow["openingbalance"], "##0.00", 16) + basText.formatNum(masterRow["totaldebits"], "#0.00;(#0.00)", 16) + basText.formatNum(masterRow["totalcredits"], "#0.00;(#0.00)", 16) + basText.formatNum(masterRow["closingbalance"], "#0.00;(#0.00)", 16) + basText.formatNum(masterRow["mindueamount"], "#0.00;(#0.00)", 16) + basText.replicat(" ", 12) + basText.formatDate(masterRow["stetementduedate"], "yyyy/MM/dd") + "\n");//+strEndOfPage
  }


  private void calcCardlRows()
  {
    totalCardPages = 0;
    totCardRows = 0;
    foreach (DataRow dtRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
    {
      if ((dtRow["POSTINGDATE"] == DBNull.Value) && (dtRow["DOCNO"] == DBNull.Value)) continue;
      //streamWrit.WriteLine(basText.trimStr(dtRow["trandescription"],40)); 
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
    stmNo = masterRow["STATEMENTNO"].ToString();
    stmNo = masterRow["accountno"].ToString();

    accountRows = DSstatement.Tables["statementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow["accountno"]).ToString().Trim() + "'");
    totalAccPages = 0;
    totAccRows = 0;
    string prevCardNo = String.Empty, CurCardNo = String.Empty;
    int currAccRowsPages = 0;
    //currAccRowsPages = 0;
    foreach (DataRow dtAccRow in accountRows) //mRow.GetChildRows(StatementNoDRel)
    {
      if ((dtAccRow["POSTINGDATE"] == DBNull.Value) && (dtAccRow["DOCNO"] == DBNull.Value)) 
        continue;
      CurCardNo = dtAccRow["cardno"].ToString();
      //if (CurCardNo.Trim().Length < 1) continue;

      currAccRowsPages +=2;//>> currAccRowsPages++;
      totAccRows +=2;//>> totAccRows++;

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
      prevCardNo = dtAccRow["cardno"].ToString();
    }
    if (currAccRowsPages > 0)
      totalAccPages++;
    if (totalAccPages < 1)
      totalAccPages = 1;
    totCrdNoInAcc = curCrdNoInAcc =0;
    mainRows = DSstatement.Tables["statementMasterTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow["accountno"]) + "'");
    curMainCard = CurCardNo = "";
    foreach (DataRow mainRow in mainRows) //mRow.GetChildRows(StatementNoDRel)
    {
      totCrdNoInAcc++;
      CurCardNo = mainRow["cardno"].ToString();
      if (mainRow["cardprimary"].ToString() == "Y")
        curMainCard = CurCardNo; //mainRow["cardno"].ToString();
      if (mainRow["cardprimary"].ToString() == "Y" && isValidateCard( mainRow["cardstate"].ToString()))
        curMainCard = CurCardNo; //mainRow["cardno"].ToString();
    }

    if (curMainCard == "")
    {
      //numOfErr++;
      //strmWriteErr.WriteLine("Account without main card : " + masterRow["accountno"].ToString());
      curMainCard = CurCardNo;
    }
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
      {
        streamWrit.WriteLine(String.Empty);
        //curAccRows++;
      }
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
//    clsValidatePageSize ValidatePageSize = new clsValidatePageSize(strFileName, 54, strEndOfPage);
//    streamSummary.WriteLine(ValidatePageSize.outMessage);
  }

  private void add2FileList(string pFileName)
  {
    int myIndex = aryLstFiles.BinarySearch((object)pFileName);
    if (myIndex < 0)
      aryLstFiles.Add(@pFileName);
  }

  private string get3Branch(string pBranch)
  {
    string rtrn;
    pBranch = pBranch; // pBranch.Substring(5).Trim();
    if (pBranch.Substring(0, 3).ToUpper() == "EL ")
      pBranch = pBranch.Substring(3).Trim();
    else if (pBranch.Substring(0, 3).ToUpper() == "AL ")
      pBranch = pBranch.Substring(3).Trim();
    else if(pBranch.IndexOf('-') > 0)
      pBranch = pBranch.Substring(pBranch.IndexOf('-')+1).Trim();

    rtrn = pBranch.Substring(0, 3).ToUpper();
    return rtrn;
  }

  public string bankName
  {
    get { return strBankName; }
    set { strBankName = value; }
  }// bankName

  ~clsStatementAAIB()
  {
    DSstatement.Dispose();
  }
}
