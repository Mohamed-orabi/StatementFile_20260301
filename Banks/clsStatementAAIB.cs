using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
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
    private DataRow masterRow;
    private DataRow detailRow;
    private string strEndOfLine = "\u000D";  
    private string strEndOfPage = "\u000C"; 
    private const int MaxDetailInPage = 26; 
    private const int MaxDetailInLastPage = 26;
    private int CurPageRec4Dtl = 0;
    private int pageNo = 0, totalPages = 0, totalCardPages = 0
        , totalAccPages = 0, totCardRows = 0, totAccRows = 0, curAccRows = 0;//
    private string lastPageTotal;
    private string curCardNo;//,PrevCardNo
    private string curAccountNo, prevAccountNo = String.Empty;//,PrevCardNo
    private decimal totNetUsage = 0;
    private DataRow[] cardsRows, accountRows, branchRows;
    private DataRow[] mainRows;
    private string CurrentPageFlag;
    private string strCardNo, strPrimaryCardNo;
    private string strForeignCurr, strForeignVal;
    private string stmNo;
    private int totNoOfCardStat, totNoOfPageStat;

    private int curCardRow = 0, curTotCardRows = 0, numOfErr = 0, totNoOfTransactions = 0, totNoOfTransactionsInt = 0;
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
    private string holdStatus = string.Empty;
    private ArrayList aryLstFiles;
    private string curFileName;
    private string accCur;
    private frmStatementFile frmMain;
    private int totRec = 1;

    private DataRow ProductRow, branchRow;
    protected bool hasInterset = false;

    protected DataRow[] emailRows = null;
    protected int externalID;
    private string accountNoName = mAccountno;

    public clsStatementAAIB()
    {
    }

    public string Statement(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pCurDate, string pStmntType, bool pAppendData)
    {
        string rtrnStr = "Successfully Generate " + pBankName;
        int curMonth = pCurDate.Month;
        bool preExit = true;
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
            strmWriteErr = new StreamWriter(fileStrmErr, Encoding.GetEncoding("ASMO-708"));// Encoding.Default
            strmWriteErr.AutoFlush = true;

            // open Summary file
            fileSummaryName = pStrFileName;
            fileSummaryName = clsBasFile.getPathWithoutExtn(fileSummaryName) +
              "_Summary." + clsBasFile.getFileExtn(fileSummaryName);
            fileSummary = new FileStream(fileSummaryName, FileMode.Create); //Create
            streamSummary = new StreamWriter(fileSummary, Encoding.GetEncoding("ASMO-708"));//Encoding.Default
            streamSummary.AutoFlush = true;


            // set branch for data
            curBranchVal = pBankCode; // 6; //6  = real   1 = test
            //     clsBasStatement.mainTableCond =  " substr(cardno,1,6) !='421192' "; //" cardproduct != 'Visa Business' ";
            //     clsBasStatement.supTableCond =  " substr(cardno,1,6) !='421192' "; 
            // data retrieve
            FillStatementDataSet(pBankCode, "vip"); //DSstatement =  //6); // 6
            getCardProduct(pBankCode);
            getBranchPart(pBankCode);
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
                cardsRows = mRow.GetChildRows(StatementNoDRel); //, DataRowVersion.Proposed

                strCardNo = masterRow[mCardno].ToString().Trim();
                if (masterRow[mHOLSTMT].ToString() == "Y")
                    continue;

                if (strCardNo.Length != 16)
                {
                    strmWriteErr.WriteLine("CardNo Length not equal 16 AccountNo " + masterRow[mAccountno].ToString().Trim() + " Card No =" + strCardNo);
                    //numOfErr++;
                    continue;// Exclude Zero Length Cards 
                }
                if ((clsCheckCard.CalcLuhnCheckDigit(strCardNo.Substring(0, 15)) != strCardNo.Substring(15, 1))) //!clsCheckCard.isValidCard(strCardNo) ||
                {
                    strmWriteErr.WriteLine("Invalid Card Number" + strCardNo);
                    numOfErr++;
                }
                curBranch = masterRow[mCardbranchpartname].ToString().Substring(5).Trim();
                //curBranch = curBranch.Substring(curBranch.IndexOf(")")+1).Trim();
                //holdStatus = masterRow["holstmt"].ToString().Trim();
                //if (holdStatus == "Y")
                //  holdStatus = "_Hold";
                //else
                //  holdStatus = "";

                if (prevBranch != curBranch)
                {
                    if (prevBranch != null)
                    {
                        streamWrit.Flush();
                        streamWrit.Close();
                        fileStrm.Close();
                    }
                    // open output file
                    //curFileName = clsBasFile.getPathWithoutExtn(pStrFileName)
                    //+ "_" + curBranch + holdStatus + "." + clsBasFile.getFileExtn(pStrFileName);
                    curFileName = clsBasFile.getPathWithoutExtn(pStrFileName)
                      + "_" + curBranch + "." + clsBasFile.getFileExtn(pStrFileName);
                    add2FileList(curFileName);
                    fileStrm = new FileStream(curFileName, FileMode.Append); //FileMode.Create Create
                    //>streamWrit = new StreamWriter(fileStrm, Encoding.Default);
                    streamWrit = new StreamWriter(fileStrm, Encoding.GetEncoding("ASMO-708"));//Encoding.GetEncoding("x-mac-arabic")  IBM420 IBM864 ASMO-708  iso-8859-6
                    streamWrit.AutoFlush = true;
                    prevBranch = curBranch; // masterRow[mCardbranchpartname].ToString().Trim();
                }
                
                strPrimaryCardNo = strCardNo;
                if (masterRow[mCardprimary].ToString() == "N")
                {
                    strPrimaryCardNo = masterRow[mPrinarycardno].ToString();
                    //calcCardlRows();
                }

                //start new account
                if (prevAccountNo == masterRow[mAccountno].ToString())
                    continue;

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




                    if (totAccRows < 1
                      && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0) //             || (masterRow[mCardno].ToString() == curMainCard   // Convert.ToDecimal(
                    {
                        isHaveF3 = true;
                        //pageNo=1; totalAccPages =1;
                        continue;
                    }


                    prevAccountNo = masterRow[mAccountno].ToString();
                    //pageNo=1; //if page is based on account no
                    printHeader();//if page is based on account no

                    totNoOfCardStat++;
                } // End of if(prevAccountNo != masterRow[mAccountno].ToString())
                //calcCardlRows();
                //if(totCardRows < 1)continue ;  //if pages is based on card
                //printHeader();//if pages is based on card

                foreach (DataRow dRow in accountRows)//cardsRows //mRow.GetChildRows(StatementNoDRel)
                {
                    detailRow = dRow;
                    stmNo = detailRow[dStatementno].ToString();
                    if ((detailRow[dPostingdate] == DBNull.Value) && (detailRow[dDocno] == DBNull.Value)) continue;// Exclude On-Hold Transactions 
                    curAccRows++; // curAccRows++;
                    if (CurPageRec4Dtl >= MaxDetailInPage)
                    {
                        CurPageRec4Dtl = 0;
                        printCardFooter();
                        pageNo++;
                        printHeader();
                    }
                    //						totNetUsage +=clsBasValid.validateNum(detailRow[dBilltranamount]);
                    totNetUsage = calculateCrDb(totNetUsage, clsBasValid.validateNum(detailRow[dBilltranamount]), clsBasValid.validateStr(detailRow[dBilltranamountsign]));
                    CurPageRec4Dtl++;
                    printDetail();
                    //if(!((detailRow[dPostingdate]== DBNull.Value) && (detailRow[dDocno]== DBNull.Value))) 

                } //end of detail foreach
                //printCardFooter();//if pages is based on card
                // if(masterRow[mCardprimary].ToString() == "Y" && (masterRow[mCardstate].ToString() == "Given" || masterRow[mCardstate].ToString() == "Embossed" || masterRow[mCardstate].ToString() == "New"))
                curCrdNoInAcc++;
                if ((curAccRows >= totAccRows))//&& totAccRows != 0|| (totAccRows == 0 && curCrdNoInAcc == totCrdNoInAcc)
                {
                    //completePageDetailRecords();
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
            //fillStatementHistory(pStmntType, pAppendData);
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
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

            clsDbOracleLayer.catchError(ex);
        }
        catch (NotSupportedException ex)  //(Exception ex)  //
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

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
        }
        return rtrnStr;
    }


    protected void printHeader()
    {
        extAccNum = clsBasValid.validateStr(masterRow[mExternalno]);
        if (extAccNum.Trim() == "" || extAccNum.Length < 16)
            extAccNum = basText.alignmentLeft((object)extAccNum, 16, '0');
        

        if (extAccNum.Substring(10, 3) == "818")
            accCur = "EGP";
        else if (extAccNum.Substring(10, 3) == "840")
            accCur = "USD";
        else
            accCur = extAccNum.Substring(10, 3);

        
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

        streamWrit.WriteLine(strEndOfPage);//line 1
        

        
        branchRows = DSBranches.Tables[0].Select("BRANCHPART = " + masterRow[mCardbranchpart].ToString());
        if (branchRows.Length > 0)
            branchRow = branchRows[0];
        emailRows = DSemails.Tables["Emails"].Select("idclient = " + masterRow[mClientid].ToString());
        for (int i = 0; i < emailRows.Length; i++)
        {
            if (!string.IsNullOrEmpty(emailRows[i][4].ToString().Trim()))
            {
                externalID = int.Parse(emailRows[i][4].ToString().Trim());
            }
        }
        streamWrit.WriteLine(basText.replicat(" ", 35)
            + basText.alignmentLeft(branchRow["ident"].ToString() + "    " + ((externalID % 2 == 0) ? externalID + 2 : externalID + 1) / 2, 13)
         + basText.replicat(" ", 6));
        streamWrit.WriteLine(basText.replicat(" ", 3)
          + basText.alignmentLeft(masterRow[mCustomername].ToString(), 46)
          + "  ACC NUM: "
          + basText.alignmentLeft(extAccNum.Substring(0, 6), 6)
          + " " + basText.alignmentLeft(extAccNum.Substring(6, 4), 4)  
          + " " + accCur 
          + " " + basText.alignmentLeft(extAccNum.Substring(13, 3), 3));  
        
        streamWrit.WriteLine(basText.replicat(" ", 42)
          + basText.alignmentLeft("CARD LIMIT :", 12)
          + basText.formatNum(masterRow[mAccountlim], "##,###,##0.000", 20, "R")
          + " " + basText.alignmentLeft(masterRow[mAccountcurrency].ToString(), 3)); //line 10
        
        streamWrit.WriteLine(basText.alignmentLeft(basText.formatNum(curMainCard, "#### #### #### ####"), 20)
          + basText.replicat(" ", 10) + "PAGE " + basText.alignmentRight(pageNo, 3) + " OF "
          + basText.alignmentRight(totalAccPages, 3)
          + basText.replicat(" ", 6) + "STATEMENT DATE:"
          + "  " + basText.formatDate(masterRow[mStatementdateto], "yyyy/MM/dd"));
        
        streamWrit.WriteLine(basText.replicat(" ", 28)
          + basText.alignmentLeft("(STAMP DUTY PAID)", 17));//line 14
        

        totalPages++;
        totNoOfPageStat++;
    }


    protected void printDetail()
    {
        DateTime trnsDate = clsBasValid.validateDate(detailRow[dTransdate]), postingDate = clsBasValid.validateDate(detailRow[dPostingdate]);
        if (trnsDate > postingDate)
            trnsDate = postingDate;

        if (masterRow[mAccountcurrency].ToString().Trim() != detailRow[dOrigtrancurrency].ToString().Trim())
        {
            
            strForeignVal = basText.formatNum(detailRow[dOrigtranamount], "#0.000", 18, "R");
            strForeignCurr = basText.Replace(basText.alignmentLeft(detailRow[dOrigtrancurrency], 18), "XXX", "   ");
        }
        else
        {
            strForeignVal = basText.replicat(" ", 18);
            strForeignCurr = basText.replicat(" ", 18);
        }

        if (detailRow[dOrigtrancurrency].ToString().Trim() == "EGP")
            strForeignCurr = "EGYPTIAN POUND";
        else if (detailRow[dOrigtrancurrency].ToString().Trim() == "USD")
            strForeignCurr = "US DOLLAR";
        else if (detailRow[dOrigtrancurrency].ToString().Trim() == "EUR")
            strForeignCurr = "EURO CURRENCY";
        else if (detailRow[dOrigtrancurrency].ToString().Trim() == "AED")
            strForeignCurr = "UAE DIRHAM";

        string trnsDesc; 
        if (detailRow[dMerchant].ToString().Trim() == "")
            trnsDesc = detailRow[dTrandescription].ToString().Trim();
        else
            trnsDesc = detailRow[dMerchant].ToString().Trim();

        
        streamWrit.WriteLine(postingDate.ToString("dd/MM") + basText.replicat(" ", 5)
        + basText.alignmentLeft(detailRow[dDocno].ToString(), 12) + "  "
        + basText.alignmentLeft(trnsDesc, 30) + "  "
            
        + basText.formatNum(detailRow[dBilltranamount], "#0.000", 19, "R") + CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])));
       
        streamWrit.WriteLine(basText.replicat(" ", 6)
          + basText.alignmentLeft(detailRow[dRefereneno].ToString(), 23) + "  "
          + strForeignVal + "  "
          + strForeignCurr);
        totNoOfTransactions++;
        //}
    }

    protected void printCardFooter()
    {
    string newaddress1, newaddress2;
    clsMaintainData.fixAddress(curBranchVal, masterRow[mCustomeraddress1].ToString(), out newaddress1, out newaddress2);
        if (pageNo == totalAccPages)
        streamWrit.WriteLine(basText.alignmentRight(basText.formatNumUnSign(masterRow[mOpeningbalance], "#,###,##0.000", 15), 15) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance].ToString())) //+ " "
          + basText.alignmentRight(basText.formatNum(masterRow[mTotaldebits].ToString(), "#,###,##0.000", 17), 17) //20), 20) + " " 
          + basText.alignmentRight(basText.formatNum(masterRow[mTotalcredits].ToString(), "#,###,##0.000", 19), 19) //19), 19) + "  "
          + basText.alignmentRight(basText.formatNumUnSign(masterRow[mClosingbalance], "#,###,##0.000", 19), 19) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])));  //19), 19))
        else
        streamWrit.WriteLine(basText.alignmentRight("0.000", 15) 
      + basText.alignmentRight("0.000", 17) 
      + basText.alignmentRight("0.000", 19) 
      + basText.alignmentRight("0.000", 19));

        streamWrit.WriteLine(basText.replicat(" ", 47)
          + basText.alignmentLeft(masterRow[mCardbranchpartname].ToString().Substring(5).Trim(), 20));
        streamWrit.WriteLine(" "
          + basText.alignmentLeft(masterRow[mCustomername].ToString(), 46)
          + basText.alignmentRight("* MINIMUM PAYMENT DUE IN ", 26) + " "
          + basText.alignmentLeft(masterRow[mAccountcurrency].ToString(), 3));
        streamWrit.WriteLine(basText.replicat(" ", 50)
          + "DUE Date" + "  "
          + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy")); //line 60
        streamWrit.WriteLine(" "
          //+ basText.alignmentLeft(ValidateArbic(masterRow[mCustomeraddress1].ToString()), 48)
          + basText.alignmentLeft(ValidateArbic(newaddress1), 48)
          + basText.replicat("*", 8)
          + basText.formatNum(masterRow[mMindueamount], "#,###,##0.000;(#,###,##0.000)", 18)); //basText.formatNum(masterRow[mCardno],"####-####-####-####")
        streamWrit.WriteLine("  "
          //+ basText.alignmentLeft(ValidateArbic(masterRow[mCustomeraddress2].ToString()), 47)
          + basText.alignmentLeft(ValidateArbic(newaddress2), 47)
          + basText.replicat("*", 8));
        streamWrit.WriteLine(basText.replicat(" ", 67)
          + basText.formatDate(masterRow[mStatementdateto], "yyyy/MM/dd")); // line 64
        streamWrit.WriteLine(basText.replicat(" ", 12) + "CARD HOLDER");  // line 65
        streamWrit.WriteLine(basText.replicat(" ", 10)
          + basText.alignmentLeft(masterRow[mCustomername].ToString(), 20));
        streamWrit.WriteLine("   PLEASE ENSURE PROMPT SETTLEMENT OF YOUR MIN PAYMENT TO AVOID CARD SUSPENSION");
        streamWrit.WriteLine(basText.replicat(" ", 57)
          + basText.alignmentLeft(basText.formatNum(curMainCard, "#### #### #### ####"), 20));
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
        stmNo = masterRow[mAccountno].ToString();

        accountRows = DSstatement.Tables["tStatementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[mAccountno]).ToString().Trim() + "'");
        totalAccPages = 0;
        totAccRows = 0;
        string prevCardNo = String.Empty, CurCardNo = String.Empty;
        int currAccRowsPages = 0;
        //currAccRowsPages = 0;
        foreach (DataRow dtAccRow in accountRows) //mRow.GetChildRows(StatementNoDRel)
        {
            if ((dtAccRow[dPostingdate] == DBNull.Value) && (dtAccRow[dDocno] == DBNull.Value))
                continue;
            CurCardNo = dtAccRow[dCardno].ToString();
            //if (CurCardNo.Trim().Length < 1) continue;

            currAccRowsPages++;//>> currAccRowsPages++;
            totAccRows++;//>> totAccRows++;

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
        mainRows = DSstatement.Tables["tStatementMasterTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[mAccountno]) + "'");
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
        {
            //numOfErr++;
            //strmWriteErr.WriteLine("Account without main card : " + masterRow[mAccountno].ToString());
            curMainCard = CurCardNo;
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
        DataRow[] productRows;
        prevAccountNo = String.Empty;
        totNoOfCardStat = 0;
        totNoOfTransactionsInt = 0;
        foreach (DataRow pRow in DSProducts.Tables["Products"].Rows)
        {
            ProductRow = pRow;
            productRows = DSstatement.Tables["tStatementMasterTable"].Select("cardproduct = '" + ProductRow[pName].ToString().Trim() + "'");
            if (productRows.Length == 0)
            {
                continue;
            }
            foreach (DataRow mRow in productRows)
            {
                masterRow = mRow;
                cardsRows = mRow.GetChildRows(StatementNoDRel); //, DataRowVersion.Proposed
                strCardNo = masterRow[mCardno].ToString().Trim();
                if (strCardNo.Length != 16)
                {
                    continue;// Exclude Zero Length Cards 
                }
                strPrimaryCardNo = strCardNo;
                if (masterRow[mCardprimary].ToString() == "N")
                {
                    strPrimaryCardNo = masterRow[mPrinarycardno].ToString();
                }
                //start new account
                if (prevAccountNo != masterRow[accountNoName].ToString())
                {
                    if (prevAccountNo == string.Empty)
                        curMainCard = string.Empty;
                    pageNo = 1;//totalAccPages = 1 ; pageNo=1;
                    CurPageRec4Dtl = 0; totNetUsage = 0; CurrentPageFlag = "F 1"; //if page is based on account no
                    calcAccountRows();
                    if (totAccRows < 1 && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0) //             || (masterRow[mCardno].ToString() == curMainCard   // Convert.ToDecimal(
                    {
                        continue;
                    }
                    prevAccountNo = masterRow[accountNoName].ToString();
                    totNoOfCardStat++;
                } // End of if(prevAccountNo != masterRow[accountNoName].ToString())
                foreach (DataRow dRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
                {
                    detailRow = dRow;
                    stmNo = detailRow[dStatementno].ToString();
                    if ((detailRow[dPostingdate] == DBNull.Value) && (detailRow[dDocno] == DBNull.Value)) continue;// Exclude On-Hold Transactions 
                    curAccRows++;
                    if (CurPageRec4Dtl >= MaxDetailInPage)
                    {
                        CurPageRec4Dtl = 0;
                        pageNo++;
                    }
                    CurPageRec4Dtl = CurPageRec4Dtl + 1;
                    hasInterset = CalcTransInterest(detailRow[dTrandescription].ToString());
                    if (hasInterset)
                        totNoOfTransactionsInt++;
                } //end of detail foreach
                curCrdNoInAcc++;
                if ((curAccRows >= totAccRows && totAccRows != 0) || (totAccRows == 0 && curCrdNoInAcc == totCrdNoInAcc))
                {
                    CurPageRec4Dtl = 0; //>pageNo=1; if pages is based on account
                    curAccRows = 0;
                }
            }
            if (totNoOfCardStat != 0)
            {
                StatSummary.NoOfStatements = totNoOfCardStat;
                StatSummary.NoOfTransactionsInt = totNoOfTransactionsInt;
                StatSummary.ProductCode = int.Parse(ProductRow["code"].ToString());
                StatSummary.ProductName = ProductRow["name"].ToString();
                StatSummary.InsertRecordDb(vCurDate.ToString("yyyyMM") + strBankName);
                totNoOfCardStat = 0;
                totNoOfTransactionsInt = 0;
            }
        }
        StatSummary = null;
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
        else if (pBranch.IndexOf('-') > 0)
            pBranch = pBranch.Substring(pBranch.IndexOf('-') + 1).Trim();

        rtrn = pBranch.Substring(0, 3).ToUpper();
        return rtrn;
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

    ~clsStatementAAIB()
    {
        DSstatement.Dispose();
    }
}
