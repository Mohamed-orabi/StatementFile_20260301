using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Xml;
using System.Collections;

// Branch 4
public class clsStatementBDCA : clsBasStatement
{
    private string strBankName;
    private string strFileName;
    private FileStream fileStrm, fileStrm2, fileSummary, fileStrmByEmail;
    private StreamWriter streamWrit, streamWrit2, streamSummary, strmWriteByEmail, strmWriteCommon;
    //private DataSet DSstatement;
    //private OracleDataReader drPrimaryCards, drMaster,drDetail;
    private DataRow masterRow;
    private DataRow detailRow;
    protected DataRow installmentRow;
    private string strEndOfLine = "\u000D";  //+ "M" ^M
    //		private string strEndOfPage = "\u000C"  ;  //+ "\u000D"+ "M" ^L + ^M
    private string strEndOfPage = "\u000C";  //  + basText.replicat(" ",85) + "F 2"  + "\u000D"+ "M" ^L + ^M
    private const int MaxDetailInPage = 25; //
    private const int MaxDetailInLastPage = 25; //
    private int CurPageRec4Dtl = 0;
    private int pageNo = 0, totalPages = 0, totalCardPages = 0
        , totalAccPages = 0, totCardRows = 0, totAccRows = 0, curAccRows = 0;//
    private string lastPageTotal;
    private string curCardNo;//,PrevCardNo
    private string curAccountNo, prevAccountNo = String.Empty;//,PrevCardNo
    private decimal totNetUsage = 0, totTrans = 0, totInstTrans = 0;
    private DataRow[] cardsRows, accountRows, installmentRows, tmpDtlRows;
    private DataRow[] mainRows;
    private string CurrentPageFlag;
    private string strCardNo, strPrimaryCardNo;
    private string strForeignCurr;
    private string stmNo;
    private int totNoOfCardStat, totNoOfPageStat;
    //public int totNoOfCardStat, totNoOfPageStat;

    private int curCardRow = 0, curTotCardRows = 0, numOfErr = 0, totNoOfTransactions = 0, totNoOfTransactionsInt = 0;
    //public int curCardRow = 0, curTotCardRows = 0, numOfErr = 0, totNoOfTransactions = 0;
    private bool isPrimaryOnly, isHaveF3 = true, isPrimaryOpened = false;
    private FileStream fileStrmErr;
    private StreamWriter strmWriteErr;
    private string curMainCard, strExpiryDate;

    private string extAccNum;
    private string prevBranch, curBranch;
    private int totCrdNoInAcc, curCrdNoInAcc, totSuppCrdsInAcc;
    private string strOutputPath, strOutputFile, fileSummaryName, strNotPrinted;
    private DateTime vCurDate;
    //private clsOMR omr = new clsOMR();
    private int totPages;
    private frmStatementFile frmMain;
    private int totRec = 1;
    private bool isSavedDataset = false;
    private string strWhereCond = string.Empty;
    private string strProductCond = string.Empty;
    private string strFileNam, stmntType;
    private DataRow ProductRow;
    protected bool hasInterset = false;
    private string curSuppCard = string.Empty;

    private FileStream filelog;
    private StreamWriter streamlog;
    private bool flag = false;
    private string excludedaccountno, excludedcardno, excludedcardmbr, excludedcardstate = string.Empty;
    private int count;
    private bool execluded = false;
    private string statMessageBoxVal = string.Empty;
    private string statBoxMessage = string.Empty;
    private int installdocno = 0;
    private int originstalldocno = 0;
    private string installdata = string.Empty;
    protected string installmentCond = "'BDC Installment Card Program'";//'New Reward Contract'
    protected bool isInstallment = false;//'New Reward Contract'
    //protected bool isInEmailService = false;
    private string emailTo = string.Empty;
    private DataRow[] emailRows = null;
    private clsValidateEmail valdEmail = new clsValidateEmail();

    public clsStatementBDCA()
    {
    }

    private bool haveEmail(string pClientID)
    {
        bool rtrnVal = false;
        emailTo = string.Empty;
        emailRows = null;
        try
        {
            emailRows = DSemails.Tables["Emails"].Select("idclient = " + pClientID);
            for (int i = 0; i < emailRows.Length; i++)
            {
                if (!string.IsNullOrEmpty(emailRows[i][1].ToString().Trim()))
                {
                    if (valdEmail.isValideEmail(emailRows[i][1].ToString().Trim()) == "ValidEmail")
                    {
                        emailTo = emailRows[i][1].ToString().Trim();
                        rtrnVal = true;
                        break;
                    }
                }
            }
        }
        catch(Exception ex)
        {
            clsBasFile.writeTxtFile(clsBasFile.Filepath + @"StatementLog.txt", Environment.NewLine + ex.Message, System.IO.FileMode.Append);

        }
        return rtrnVal;
    }

    public string Statement(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pDate, string pStmntType, bool pAppendData)
    {
        string rtrnStr = "Successfully Generate " + pBankName;
        int curMonth = pDate.Month;
        bool preExit = true;
        strFileNam = pStrFileName;
        stmntType = pStmntType;

        try
        {
            clsMaintainData maintainData = new clsMaintainData();
            maintainData.notRward = false;
            maintainData.matchCardBranch4Account(pBankCode);

            // merge transaction fee with original transaction
            //clsMaintainData maintainData = new clsMaintainData();
            //maintainData.mergeTrans(pBankCode);
            pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);
            vCurDate = pDate; //DateTime.Now.AddMonths(-1);
            strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
            clsBasFile.createDirectory(pStrFileName + vCurDate.ToString("yyyyMM") + pBankName);
            pStrFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM") + ".txt";
            strBankName = pBankName;
            strFileName = pStrFileName;
            strOutputFile = pStrFileName;
            // open output file
            fileStrm = new FileStream(pStrFileName, FileMode.Create); //Create
            streamWrit = new StreamWriter(fileStrm, Encoding.Default);
            streamWrit.AutoFlush = true;
            // open Not Printed
            strNotPrinted = clsBasFile.getPathWithoutExtn(pStrFileName) + "_Not_Printed_Statements." + clsBasFile.getFileExtn(pStrFileName);
            fileStrm2 = new FileStream(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Not_Printed_Statements." + clsBasFile.getFileExtn(pStrFileName), FileMode.Create); //Create
            streamWrit2 = new StreamWriter(fileStrm2, Encoding.Default);
            streamWrit2.WriteLine("Account No|External Account No|OpeningBalance|Closing Balance|Primary Card No|MBR|Card Status");
            streamWrit2.AutoFlush = true;
            // Log file
            filelog = new FileStream(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Log." + clsBasFile.getFileExtn(pStrFileName), FileMode.Create); //Create
            streamlog = new StreamWriter(filelog, Encoding.Default);
            streamlog.AutoFlush = true;
            // Error file
            fileStrmErr = new FileStream(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName), FileMode.Create);
            strmWriteErr = new StreamWriter(fileStrmErr, Encoding.Default);
            strmWriteErr.AutoFlush = true;

            // By Email file
            //if (isInEmailService)
            //{
                fileStrmByEmail = new FileStream(clsBasFile.getPathWithoutExtn(pStrFileName) + "_ByEmail." + clsBasFile.getFileExtn(pStrFileName), FileMode.Create);
                strmWriteByEmail = new StreamWriter(fileStrmByEmail, Encoding.Default);
                strmWriteByEmail.AutoFlush = true;
            //}
            
            // open Summary file
            fileSummaryName = pStrFileName;
            fileSummaryName = clsBasFile.getPathWithoutExtn(fileSummaryName) +
              "_Summary." + clsBasFile.getFileExtn(fileSummaryName);
            fileSummary = new FileStream(fileSummaryName, FileMode.Create); //Create
            streamSummary = new StreamWriter(fileSummary, Encoding.Default);
            streamSummary.AutoFlush = true;
            // set branch for data
            curBranchVal = pBankCode; // 4; //4  = real   1 = test
            isFullFields = false;
            MainTableCond = " m.contracttype in " + strProductCond + "";//strWhereCond
            supTableCond = " d.statementno in (select x.statementno from " + clsSessionValues.mainDbSchema + clsSessionValues.mainTable + " x where x.branch = " + pBankCode + " and x.contracttype in " + strProductCond + ")";
            //strOrder = " m.CARDBRANCHPART,m.accountno,m.cardprimary desc"; // BDCA-2390
            strOrder = " m.accountno,m.cardprimary desc,cardcreationdate"; // BDCA-2550
            curInstallmentCond = installmentCond;
            FillStatementDataSet(pBankCode, ""); //DSstatement =  //6); // 6
            getClientEmail(pBankCode);
            getCardProduct(pBankCode);
            getInstallment(pBankCode);
            pageNo = 1; totalCardPages = 0;
            curCardNo = String.Empty;
            curAccountNo = String.Empty;
            execluded = false;

            frmMain.BeginInvoke(frmMain.setMinMaxProgressDelegate, new object[] { DSstatement.Tables["tStatementMasterTable"].Rows.Count });
            foreach (DataRow mRow in DSstatement.Tables["tStatementMasterTable"].Rows)
            {
                frmMain.BeginInvoke(frmMain.setProgressDelegate, new object[] { totRec++ });
                masterRow = mRow;
                //strmWriteCommon.WriteLine(masterRow[mStatementno].ToString());
                //pageNo=1; CurPageRec4Dtl=0; CurrentPageFlag = "F 1"; //if page is based on card no
                cardsRows = mRow.GetChildRows(StatementNoDRel); //, DataRowVersion.Proposed
                strCardNo = masterRow[mCardno].ToString().Trim();
                if (masterRow[mHOLSTMT].ToString() == "Y")
                    continue;
                if (strCardNo.Length != 16)
                {
                    strmWriteErr.WriteLine("CardNo Length not equal 16 for AccountNo " + masterRow[mAccountno].ToString().Trim() + " Card No =" + strCardNo);
                    //numOfErr++;
                    continue;// Exclude Zero Length Cards 
                }

                if (!(strCardNo == ""))
                {
                    if ((clsCheckCard.CalcLuhnCheckDigit(strCardNo.Substring(0, 15)) != strCardNo.Substring(15, 1))) //!clsCheckCard.isValidCard(strCardNo) ||
                    {
                        strmWriteErr.WriteLine("Invalid Card Number" + strCardNo);
                        numOfErr++;
                    }
                }
                curBranch = masterRow[mCardbranchpartname].ToString().Trim();
                curBranch = curBranch.Substring(curBranch.IndexOf(")") + 1).Trim();
                if (prevBranch != curBranch)
                {
                    // multi file branch
                    //if (prevBranch != null)
                    //{
                    //  streamWrit.Flush();
                    //  streamWrit.Close();
                    //  fileStrm.Close();
                    //}
                    // multi file branch // open output file
                    //fileStrm = new FileStream(clsBasFile.getPathWithoutExtn(pStrFileName) +
                    //"_" + curBranch + "." + clsBasFile.getFileExtn(pStrFileName), FileMode.Create); //Create
                    //streamWrit = new StreamWriter(fileStrm, Encoding.Default);
                    prevBranch = curBranch; // masterRow[mCardbranchpartname].ToString().Trim();
                }


                //if (GetValidCardByAccount(masterRow[mAccountno].ToString()))
                //{
                //    count++;
                //    streamWrit2.WriteLine(masterRow[mAccountno].ToString());//+ "|" + masterRow[mExternalno].ToString() + "|" + masterRow[mClosingbalance].ToString() + "|" + masterRow[mCardno].ToString() + "|" + masterRow[mMBR].ToString() + "|" + masterRow[mCardstate].ToString());
                //    continue;
                //}


                strPrimaryCardNo = strCardNo;
                if (masterRow[mCardprimary].ToString() == "N")
                {
                    strPrimaryCardNo = masterRow[mPrinarycardno].ToString();
                    //calcCardlRows();
                }

                //start new account

                if (prevAccountNo != masterRow[mAccountno].ToString())
                {
                    streamWrit.Flush();
                    strmWriteByEmail.Flush();
                    if (haveEmail(masterRow[mClientid].ToString()))
                        strmWriteCommon = strmWriteByEmail;
                    else
                        strmWriteCommon = streamWrit;

                    if (prevAccountNo == string.Empty)
                        if (curMonth != Convert.ToDateTime(masterRow[mStatementdateto]).Month)//,"dd/MM/yyyy"
                        {
                            preExit = false;
                            fileStrm.Close();
                            fileStrmByEmail.Close();
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

                    //cardsRows = DSstatement.Tables["tStatementDetailTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[mAccountno]).ToString().Trim() + "'");

                    curMainCard = string.Empty;
                    if (!isHaveF3)//!isHaveF3  CurrentPageFlag != "F 2"
                    {
                        strmWriteErr.WriteLine("Account not Closed : " + prevAccountNo);
                        numOfErr++;
                    }
                    isHaveF3 = false;

                    pageNo = 1;//totalAccPages = 1 ; pageNo=1;
                    CurPageRec4Dtl = 0; totNetUsage = 0; CurrentPageFlag = "F 1"; totTrans = 0; totInstTrans = 0; //if page is based on account no
                    calcAccountRows();




                    //>>if (totAccRows < 1
                    if (totAccRows < 2
                      && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0) //             || (masterRow[mCardno].ToString() == curMainCard   // Convert.ToDecimal(
                    {
                        isHaveF3 = true;
                        //pageNo=1; totalAccPages =1;
                        continue;
                    }
                    //streamWrit2.WriteLine(masterRow[mAccountno].ToString() + "|" + masterRow[mExternalno].ToString() + "|" + masterRow[mClosingbalance].ToString() + "|" + masterRow[mCardno].ToString() + "|" + masterRow[mMBR].ToString() + "|" + masterRow[mCardstate].ToString());
                    prevAccountNo = masterRow[mAccountno].ToString();
                    if (GetAccountsWithoutTrxns(masterRow[mAccountno].ToString()) == true)
                    {
                        count++;
                        execluded = true;
                        streamWrit2.WriteLine(excludedaccountno + "|" + masterRow[mExternalno].ToString() + "|" + masterRow[mOpeningbalance].ToString() + "|" + masterRow[mClosingbalance].ToString() + "|" + basText.formatCardNumber(excludedcardno) + "|" + excludedcardmbr + "|" + excludedcardstate);
                        continue;
                    }

                    flag = true;
                    //pageNo=1; //if page is based on account no
                    printHeader();//if page is based on account no
                    flag = false;
                    totNoOfCardStat++;
                    if (Convert.ToDecimal(masterRow[mAccountlim].ToString()) <= Convert.ToDecimal(0))
                    {
                        strmWriteErr.WriteLine("Account Limit Less than or Equal Zero for Account " + masterRow[mAccountno].ToString());// + " and Card Number " + strCardNo
                        numOfErr++;
                    }
                } // End of if(prevAccountNo != masterRow[mAccountno].ToString())
                //calcCardlRows();
                //if(totCardRows < 1)continue ;  //if pages is based on card
                //printHeader();//if pages is based on card

                if (execluded == true)
                {
                    //execluded = false;
                    continue;
                }

                foreach (DataRow dRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
                {
                    detailRow = dRow;
                    stmNo = detailRow[dStatementno].ToString();
                    if ((detailRow[dPostingdate] == DBNull.Value) && (detailRow[dDocno] == DBNull.Value))
                        continue;// Exclude On-Hold Transactions 
                    curAccRows++;
                    if (CurPageRec4Dtl >= MaxDetailInPage)
                    {
                        CurPageRec4Dtl = 0;
                        printCardFooter();
                        flag = false;
                        pageNo++;
                        printHeader();
                        flag = false;
                    }
                    //						totNetUsage +=clsBasValid.validateNum(detailRow[dBilltranamount]);
                    totNetUsage = calculateCrDb(totNetUsage, clsBasValid.validateNum(detailRow[dBilltranamount]), clsBasValid.validateStr(detailRow[dBilltranamountsign]));
                    CurPageRec4Dtl++;

                    //BDCA-2390
                    if (curSuppCard != detailRow[dCardno].ToString())
                    {
                        tmpDtlRows = DSstatement.Tables["tStatementMasterTable"].Select(mCardno + " = '" + detailRow[dCardno] + "'");
                        if (tmpDtlRows.Length > 0)
                        {
                            if (tmpDtlRows[0][mCardprimary].ToString() == "N" && isValidateCard(tmpDtlRows[0][mCardstate].ToString()))
                            {
                                strmWriteCommon.WriteLine(basText.replicats(" ", 23) + ">> Supplementary Card : " + basText.formatCardNumber(detailRow[dCardno].ToString()));
                                CurPageRec4Dtl++;
                                curCrdNoInAcc++;
                                curAccRows++;
                                if (CurPageRec4Dtl > MaxDetailInPage)
                                {
                                    CurPageRec4Dtl = 1;
                                    printCardFooter();
                                    pageNo++;
                                    printHeader();
                                }
                            }
                        }
                    }
                    curSuppCard = detailRow[dCardno].ToString();

                    printDetail();
                    flag = false;
                    //if(!((detailRow[dPostingdate]== DBNull.Value) && (detailRow[dDocno]== DBNull.Value))) 
                    if (hasInterset)
                        totNoOfTransactionsInt++;
                } //end of detail foreach
                //printCardFooter();//if pages is based on card
                // if(masterRow[mCardprimary].ToString() == "Y" && (masterRow[mCardstate].ToString() == "Given" || masterRow[mCardstate].ToString() == "Embossed" || masterRow[mCardstate].ToString() == "New"))
                curCrdNoInAcc++;
                if ((curAccRows >= totAccRows && totAccRows != 0) || (totAccRows == 0 && curCrdNoInAcc == totCrdNoInAcc))
                {
                    completePageDetailRecords();
                    printCardFooter();//if pages is based on account
                    flag = false;
                    //printAccountFooter();
                    CurPageRec4Dtl = 0; //>pageNo=1; if pages is based on account
                    curAccRows = 0;
                    execluded = false;
                }
                //strmWriteCommon.WriteLine(strEndOfPage);
                //pageNo=1; CurPageRec4Dtl=0; //if pages is based on card
                //completePageDetailRecords();

            } //end of Master foreach

            //if (!isSavedDataset)
            //{
            //  clsSaveDataset saveDataset = new clsSaveDataset();
            //  saveDataset.save(DSstatement, clsBasFile.getPathWithoutExtn(pStrFileName) + ".dat");
            //}
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

                // Close output File
                strmWriteByEmail.Flush();
                strmWriteByEmail.Close();
                fileStrmByEmail.Close();

                // Close Summary File
                streamSummary.Flush();
                streamSummary.Close();
                fileSummary.Close();

                strmWriteErr.Flush();
                strmWriteErr.Close();
                fileStrmErr.Close();

                streamlog.Flush();
                streamlog.Close();
                filelog.Close();

                streamWrit2.Flush();
                streamWrit2.Close();
                fileStrm2.Close();

                ArrayList aryLstFiles = new ArrayList();
                if (numOfErr == 0)
                    clsBasFile.deleteFile(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err." + clsBasFile.getFileExtn(pStrFileName));
                else
                    aryLstFiles.Add(clsBasFile.getPathWithoutExtn(pStrFileName) + "_Err.txt");
                /*      else
                        clsBasFile.deleteFile(pStrFileName);*/

                //aryLstFiles.Add("");
                ///
                //clsExportData exportData = new clsExportData();
                //string sqlStr = "SELECT x.cardno as \"Card Number\", x.Customername as \"Customer Name\", x.customeraddress1 || ' ' || x.customeraddress2 || ' ' || x.customeraddress3 || ' ' || x.customerregion || ' ' || x.customercity as \"Customer Address\" FROM tstatementmastertable x where x.branch = " + pBankCode.ToString();
                //exportData.GetOracleData(sqlStr);
                //exportData.ExportDelimitedFile(clsBasFile.getPathWithoutExtn(pStrFileName) + "_RawData.txt", "|",true);
                //aryLstFiles.Add(clsBasFile.getPathWithoutExtn(pStrFileName) + "_RawData.txt");
                //exportData = null;
                ///
                aryLstFiles.Add(@strOutputFile);
                aryLstFiles.Add(@clsBasFile.getPathWithoutExtn(pStrFileName) + "_ByEmail.txt");
                aryLstFiles.Add(@strNotPrinted);
                numOfErr = validateNoOfLines(aryLstFiles, 61);
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
        flag = true;
        streamlog.WriteLine("header " + flag);

        totPages++;
        extAccNum = clsBasValid.validateStr(masterRow[mExternalno]);
        if (extAccNum.Trim() == "")
            extAccNum = clsBasValid.validateStr(masterRow[mAccountno]);

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

        strmWriteCommon.WriteLine(strEndOfPage);
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(basText.replicat(" ", 85) + "Page " + pageNo.ToString() + " of " + totalAccPages.ToString());  //basText.formatNum(masterRow[mAccountno],"##-##-######-####")
        strmWriteCommon.WriteLine("*" + basText.alignmentLeft(masterRow[mAccountzipcode], 5) + "*");  //
        strmWriteCommon.WriteLine(basText.replicat(" ", 15) + basText.alignmentLeft((object)(masterRow[mCustomertitle].ToString().Trim() + " " + masterRow[mCustomername].ToString().Trim()), 30) + basText.replicat(" ", 33) + basText.alignmentLeft(basText.formatCardNumber(curMainCard), 16));  //

        strmWriteCommon.WriteLine(basText.replicat(" ", 15) + basText.alignmentLeft(ValidateArbic(masterRow[mCustomeraddress1].ToString()), 80));
        strmWriteCommon.WriteLine(basText.replicat(" ", 15) + basText.alignmentLeft(ValidateArbic(masterRow[mCustomeraddress2].ToString()), 60) + basText.replicat(" ", 3) + basText.formatDate(masterRow[mStatementdateto], "dd/MM/yyyy"));
        strmWriteCommon.WriteLine(basText.replicat(" ", 15) + basText.alignmentLeft(ValidateArbic(masterRow[mCustomeraddress3].ToString().Trim()) + " " + masterRow[mCustomerregion].ToString().Trim() + " " + masterRow[mCustomercity].ToString().Trim(), 80));

        strmWriteCommon.WriteLine(basText.replicat(" ", 77) + basText.formatNum(masterRow[mAccountlim], "########", 16, "L"));  //cardproduct+"{0,5:MM/yy}",masterRow[mStatementdatefrom]
        strmWriteCommon.WriteLine(basText.replicat(" ", 15) + basText.replicat(" ", 63) + basText.alignmentLeft(masterRow[mCardbranchpart].ToString().Trim(), 10));//String.Empty

        strmWriteCommon.WriteLine(basText.alignmentRight(masterRow[mAccountzipcode].ToString(), 42));  //cardproduct+"{0,5:MM/yy}",masterRow[mStatementdatefrom]

        //strmWriteCommon.WriteLine("áÏæÇÚì ÇáÓÑíÉ Êã ÊÔÝíÑ ÑÞã ÇáÈØÇÞÉ ÇáãØÈæÚ Ýì ßÔÝ ÇáÍÓÇÈ ¡ ÇáÑÌÇÁ ÚäÏ ÇÑÓÇá ãäÏæÈ ááÓÏÇÏ ÊÃßÏ ãä ßÊÇÈÉ ÑÞã ÇáÈØÇÞÉ ÈæÖæÍ ÍÊì íÊÓäì áäÇ ÇÖÇÝÉ ÇáãÈáÛ ÇáãÏÝæÚ Ýì ÍÓÇÈ ÇáÈØÇÞÉ ÈÕæÑÉ ÕÍíÍÉ.");//Message1
        if (!string.IsNullOrEmpty(statMessageBoxVal))
        {
            FileStream filRead = null;
            StreamReader filStream = null;
            filRead = new FileStream(statMessageBoxVal, FileMode.Open);
            filStream = new StreamReader(filRead, Encoding.Default);
            statBoxMessage = filStream.ReadToEnd();
            filStream.Close();
            filRead.Close();
        }
        strmWriteCommon.WriteLine(statBoxMessage);
        //string arabicMsg = "áÏæÇÚì ÇáÓÑíÉ Êã ÊÔÝíÑ ÑÞã ÇáÈØÇÞÉ ÇáãØÈæÚ Ýì ßÔÝ ÇáÍÓÇÈ ¡ ÇáÑÌÇÁ ÚäÏ ÇÑÓÇá ãäÏæÈ ááÓÏÇÏ ÊÃßÏ ãä ßÊÇÈÉ ÑÞã ÇáÈØÇÞÉ ÈæÖæÍ ÍÊì íÊÓäì áäÇ ÇÖÇÝÉ ÇáãÈáÛ ÇáãÏÝæÚ Ýì ÍÓÇÈ ÇáÈØÇÞÉ ÈÕæÑÉ ÕÍíÍÉ.";
        //Encoding iso = Encoding.GetEncoding("ISO-8859-6");
        //Encoding utf8 = Encoding.UTF8;
        //byte[] utfBytes = utf8.GetBytes(arabicMsg);
        //byte[] isoBytes = Encoding.Convert(utf8, iso, utfBytes);
        //string msg = iso.GetString(isoBytes);
        //strmWriteCommon.WriteLine(msg);//Message1
        strmWriteCommon.WriteLine("For your security; we have masked the card number in your statement; please make sure that you typed the card number in full within your payment to guarantee proper payment process with no errors.");//Message2
        strmWriteCommon.WriteLine(String.Empty);//Message3
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);

        if (pageNo == 1)//&& CurPageRec4Dtl == 0
        {
            if (isInstallment == true && !InstallmentCondition.Contains("BDC Easy Payment Plan"))
            {
                installmentRows = DSInstallment.Tables["installment"].Select("CLIENTID = '" + clsBasValid.validateStr(masterRow[mClientid]).ToString().Trim() + "' and CREDITCONTRACTS = '" + clsBasValid.validateStr(masterRow[mContractno]).ToString().Trim() + "'");
                if (installmentRows.Length > 0)
                {
                    installmentRow = installmentRows[0];
                    strmWriteCommon.WriteLine(basText.replicat(" ", 24) + basText.alignmentLeft("Previous Balance", 70) + basText.alignmentRight((Object)(basText.formatNumUnSign(masterRow[mOpeningbalance], "#0.00", 15)), 15) + "  " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance])) + basText.alignmentRight((Object)(basText.formatNumUnSign(decimal.Parse(masterRow[mOpeningbalance].ToString()) + decimal.Parse(installmentRow[mClosingbalance].ToString()), "#0.00", 15)), 13) + "  " + CrDb(decimal.Parse(masterRow[mOpeningbalance].ToString()) + decimal.Parse(installmentRow[mClosingbalance].ToString())));
                }
                else
                {
                    strmWriteCommon.WriteLine(basText.replicat(" ", 24) + basText.alignmentLeft("Previous Balance", 70) + basText.alignmentRight((Object)(basText.formatNumUnSign(masterRow[mOpeningbalance], "#0.00", 15)), 15) + "  " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance])) + basText.alignmentRight((Object)(basText.formatNumUnSign(decimal.Parse(masterRow[mOpeningbalance].ToString()) + decimal.Parse("0"), "#0.00", 15)), 13) + "  " + CrDb(decimal.Parse(masterRow[mOpeningbalance].ToString()) + decimal.Parse("0")));
                }
            }
            else
                strmWriteCommon.WriteLine(basText.replicat(" ", 24) + basText.alignmentLeft("Previous Balance", 70) + basText.alignmentRight((Object)(basText.formatNumUnSign(masterRow[mOpeningbalance], "#0.00", 15)), 15) + "  " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance])));
            CurPageRec4Dtl++; curAccRows++;
            if (isInstallment == true && !InstallmentCondition.Contains("BDC Easy Payment Plan"))
            {
                if (installmentRows.Length > 0)
                {
                    installmentRow = installmentRows[0];
                    totInstTrans = decimal.Parse(masterRow[mOpeningbalance].ToString()) + decimal.Parse(installmentRow[mClosingbalance].ToString());
                    totTrans = decimal.Parse(masterRow[mOpeningbalance].ToString());
                }
                else
                    totTrans = Convert.ToDecimal(masterRow[mOpeningbalance]);
            }
            else
                totTrans = Convert.ToDecimal(masterRow[mOpeningbalance]);
        }
        //    strmWriteCommon.WriteLine(basText.replicat(" ", 81) + CurrentPageFlag);  //CurrentPageFlag+"{0,5:MM/yy}",masterRow[mStatementdatefrom]
        //strmWriteCommon.WriteLine(basText.replicat(" ", 81) + basText.formatNum(curMainCard, "#### #### #### ####"));  //cardbranchpartname
        //strmWriteCommon.WriteLine(basText.replicat(" ", 81) + basText.formatNum(masterRow[mAccountlim], "########")); //extAccNum   clsBasValid.validateStr(masterRow[mExternalno]) //accountno  //basText.formatNum(masterRow[mCardno],"####-####-####-####")
        //strmWriteCommon.WriteLine(basText.replicat(" ", 81) + basText.alignmentLeft(masterRow[mMindueamount], 13));//"{0,10:dd/MM/yyyy}", masterRow[mStatementdateto])   Statementdatefrom+ basText.formatNum(masterRow[mAccountlim],"########")  basText.formatNum(masterRow[mCardlimit],"########")
        //>>    strmWriteCommon.WriteLine(basText.replicat(" ", 81) + basText.alignmentLeft(pageNo,5) + basText.alignmentLeft(totalAccPages,5));  //basText.formatNum(masterRow[mAccountno],"##-##-######-####")
        //strmWriteCommon.WriteLine(" " + omr.Line2LastPage(pageNo, totalAccPages));//String.Empty
        //strmWriteCommon.WriteLine(" " + omr.Line3(pageNo));//String.Empty
        //strmWriteCommon.WriteLine(" " + omr.Line4(pageNo));//String.Empty
        //strmWriteCommon.WriteLine(" " + omr.fixLine());//String.Empty
        //    strmWriteCommon.WriteLine(basText.alignmentMiddle(curMainCard, 20) + basText.alignmentMiddle(masterRow[mAccountlim], 13) + basText.formatNum(masterRow[mAccountavailablelim], "##0", 20) + basText.alignmentMiddle(masterRow[mMindueamount], 13) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M") + basText.formatNum(masterRow[mTotaloverdueamount], "##0.00", 13)); //+ basText.formatNum(masterRow[mAccountlim],"#,##0.00",9)+" "+basText.formatNum(masterRow[mTotalcredits],"#,##0.00",9)+" "+basText.formatNum(masterRow[mAccountavailablelim],"#,##0.00",9)+" "+basText.formatNum(masterRow[mMindueamount],"#,##0.00",15)+" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",15)+" "+basText.formatDate(masterRow[mStetementduedate],"dd/MM/yy") //TOTALCREDITS// basText.alignmentMiddle(masterRow[mAccountavailablelim],20) //"#,##0"
        //    strmWriteCommon.WriteLine(String.Empty);
        //    strmWriteCommon.WriteLine(String.Empty);
        //strmWriteCommon.WriteLine(String.Empty);
        //    strmWriteCommon.WriteLine(String.Empty);
        //    if (pageNo == 1)
        //      strmWriteCommon.WriteLine(basText.replicat(" ", 36) + basText.alignmentLeft("Previous Balance", 63) + basText.formatNumUnSign(masterRow[mOpeningbalance], "#,###0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance])));  //+ CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign]))
        //    else
        //      strmWriteCommon.WriteLine(String.Empty);
        //    strmWriteCommon.WriteLine(String.Empty);

        totalPages++;
        totNoOfPageStat++;
    }

    //protected void printHeaderSummary()
    //    {
    //    totPages++;
    //    extAccNum = clsBasValid.validateStr(masterRow[mExternalno]);
    //    if (extAccNum.Trim() == "")
    //        extAccNum = clsBasValid.validateStr(masterRow[mAccountno]);

    //    if (pageNo == 1 && totalAccPages == 1) // statement contain 1 page
    //        {
    //        CurrentPageFlag = "F 0";
    //        isHaveF3 = true;
    //        }
    //    else if (pageNo == 1 && totalAccPages > 1)  //first page of multiple page statement
    //        CurrentPageFlag = "F 1"; // //middle page of multiple page statement
    //    else if (pageNo < totalAccPages)
    //        CurrentPageFlag = "F 2";
    //    else if (pageNo == totalAccPages) //last page of multiple page statement
    //        {
    //        CurrentPageFlag = "F 3";
    //        isHaveF3 = true;
    //        }

    //    if (pageNo == 1)//&& CurPageRec4Dtl == 0
    //        {
    //        CurPageRec4Dtl++; curAccRows++;
    //        totTrans = Convert.ToDecimal(masterRow[mOpeningbalance]);
    //        }
    //    totalPages++;
    //    totNoOfPageStat++;
    //    }

    protected void printDetail()
    {
        if (flag == false)
        {
            streamlog.WriteLine(detailRow[dStatementno].ToString());
            flag = true;
        }
        streamlog.WriteLine("detail " + flag);

        DateTime trnsDate = clsBasValid.validateDate(detailRow[dTransdate]), postingDate = clsBasValid.validateDate(detailRow[dPostingdate]);
        if (trnsDate > postingDate)
            trnsDate = postingDate;

        if (masterRow[mAccountcurrency].ToString().Trim() != detailRow[dOrigtrancurrency].ToString().Trim())
            strForeignCurr = basText.alignmentRight((Object)basText.formatNum(detailRow[dOrigtranamount], "#0.00;(#0.00)") + " " + basText.Replace(clsBasValid.validateStr(detailRow[dOrigtrancurrency]), "XXX", "   "), 15);
        else
            strForeignCurr = basText.replicat(" ", 15);

        if (strForeignCurr.Trim() == "0")
            strForeignCurr = basText.replicat(" ", 15);

        string trnsDesc; //= detailRow[dTrandescription].ToString().Trim() + " " + detailRow[dMerchant].ToString().Trim()
        string input;
        if (detailRow[dMerchant].ToString().Trim() == "")
            trnsDesc = detailRow[dTrandescription].ToString().Trim();
        else
            //BDCA-11077
            if (detailRow[dTrandescription].ToString().Trim() == "ATM deposit" || detailRow[dTrandescription].ToString().Trim() == "Internet banking deposit")
                trnsDesc = detailRow[dMerchant].ToString().Trim() + "," + detailRow[dTrandescription].ToString().Trim();
            else
                trnsDesc = detailRow[dMerchant].ToString().Trim();

        totTrans += valDbCr(Convert.ToDecimal(detailRow[dBilltranamount]), clsBasValid.validateStr(detailRow[dBilltranamountsign])); //basText.formatNum(detailRow[dBilltranamount], "#0.00;(#0.00)", 19);
        totInstTrans += valDbCr(Convert.ToDecimal(detailRow[dBilltranamount]), clsBasValid.validateStr(detailRow[dBilltranamountsign])); //basText.formatNum(detailRow[dBilltranamount], "#0.00;(#0.00)", 19);

        if (isInstallment == true && InstallmentCondition.Contains("BDC Easy Payment Plan"))
        {
            if (detailRow[dTrandescription].ToString() == "Amount Transferred for Refinancing")//"Installment Debt Transfer")//"Installment Transfer")
                strmWriteCommon.WriteLine(trnsDate.ToString("dd/MM/yyyy") + " " + postingDate.ToString("dd/MM/yyyy") + "  " + basText.alignmentLeft(trnsDesc, 46) + basText.alignmentRight(strForeignCurr, 15) + basText.alignmentRight((object)basText.formatNumUnSign(detailRow[dBilltranamount], "#0.00", 9), 9) + "    " + basText.alignmentRight((object)basText.formatNumUnSign((object)totTrans, "#0.00", 11), 11) + "  " + CrDb(totTrans)); //isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))

            else if ((trnsDesc == "Installment"))// || (trnsDesc == "Installment Acceleration"))//"Installment Repayments")
            {
                if (detailRow[dInstallmentData].ToString().Trim() != "")
                {
                    input = detailRow[dInstallmentData].ToString().Trim();
                    installdata = input;
                    int index = detailRow[dInstallmentData].ToString().Trim().IndexOf(":");
                    if (index > 0)
                        input = input.Substring(0, index);
                    long installdocno = long.Parse(detailRow[dDocno].ToString());
                    long originstalldocno = long.Parse(detailRow[dOrigDocNo].ToString());
                    strmWriteCommon.WriteLine(trnsDate.ToString("dd/MM/yyyy") + " " + postingDate.ToString("dd/MM/yyyy") + "  " + basText.alignmentLeft(trnsDesc + input.Substring(basText.GetNthIndex(input, ',', 2)).Substring(input.Substring(basText.GetNthIndex(input, ',', 2)).IndexOf('#')).Replace('#', ' ') + " " + basText.trimStr(GetMerchantByOrigDocNo(long.Parse(detailRow[dOrigDocNo].ToString())), 40), 46) + basText.alignmentRight(strForeignCurr, 15) + basText.alignmentRight((object)basText.formatNumUnSign(detailRow[dBilltranamount], "#0.00", 9), 9) + " " + CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + " " + basText.alignmentRight((object)basText.formatNumUnSign((object)totTrans, "#0.00", 11), 11) + "  " + CrDb(totTrans)); //isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))
                }
                else
                    strmWriteCommon.WriteLine(trnsDate.ToString("dd/MM/yyyy") + " " + postingDate.ToString("dd/MM/yyyy") + "  " + basText.alignmentLeft(trnsDesc, 46) + basText.alignmentRight(strForeignCurr, 15) + basText.alignmentRight((object)basText.formatNumUnSign(detailRow[dBilltranamount], "#0.00", 9), 9) + "    " + basText.alignmentRight((object)basText.formatNumUnSign((object)totTrans, "#0.00", 11), 11) + "  " + CrDb(totTrans)); //isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))
            }
            else if ((trnsDesc == "Installment Finance Charges"))// || (trnsDesc == "Acceleration Finance Charges"))
            {
                int index = installdata.Trim().IndexOf(":");
                //if (int.Parse(detailRow[dDocno].ToString()) == installdocno)
                {
                    if (index > 0 && index != -1)
                    {
                        installdata = installdata.Substring(0, index);
                        strmWriteCommon.WriteLine(trnsDate.ToString("dd/MM/yyyy") + " " + postingDate.ToString("dd/MM/yyyy") + "  " + basText.alignmentLeft(trnsDesc + installdata.Substring(basText.GetNthIndex(installdata, ',', 2)).Substring(installdata.Substring(basText.GetNthIndex(installdata, ',', 2)).IndexOf('#')).Replace('#', ' ') + " " + basText.trimStr(GetMerchantByOrigDocNo(originstalldocno), 40), 46) + basText.alignmentRight(strForeignCurr, 15) + basText.alignmentRight((object)basText.formatNumUnSign(detailRow[dBilltranamount], "#0.00", 9), 9) + " " + CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + " " + basText.alignmentRight((object)basText.formatNumUnSign((object)totTrans, "#0.00", 11), 11) + "  " + CrDb(totTrans)); //isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))
                    }
                    else
                        strmWriteCommon.WriteLine(trnsDate.ToString("dd/MM/yyyy") + " " + postingDate.ToString("dd/MM/yyyy") + "  " + basText.alignmentLeft(trnsDesc, 46) + basText.alignmentRight(strForeignCurr, 15) + basText.alignmentRight((object)basText.formatNumUnSign(detailRow[dBilltranamount], "#0.00", 9), 9) + "    " + basText.alignmentRight((object)basText.formatNumUnSign((object)totTrans, "#0.00", 11), 11) + "  " + CrDb(totTrans)); //isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))
                }
                //else
                //    {
                //    if (index > 0)
                //        installdata = installdata.Substring(0, index);
                //    strmWriteCommon.WriteLine(trnsDate.ToString("dd/MM/yyyy") + " " + postingDate.ToString("dd/MM/yyyy") + "  " + basText.alignmentLeft(trnsDesc + installdata.Substring(basText.GetNthIndex(installdata, ',', 2)).Substring(installdata.Substring(basText.GetNthIndex(installdata, ',', 2)).IndexOf('#')).Replace('#', ' ') + " " + basText.trimStr(GetMerchantByOrigDocNo(originstalldocno), 40), 46) + basText.alignmentRight(strForeignCurr, 15) + basText.alignmentRight((object)basText.formatNumUnSign(detailRow[dBilltranamount], "#0.00", 9), 9) + "  " + CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + basText.alignmentRight((object)basText.formatNumUnSign((object)totTrans, "#0.00", 11), 11) + "  " + CrDb(totTrans) + basText.alignmentRight((object)basText.formatNumUnSign((object)totInstTrans, "#0.00", 13), 13) + "  " + CrDb(totInstTrans)); //isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))
                //    }
            }
            else
                strmWriteCommon.WriteLine(trnsDate.ToString("dd/MM/yyyy") + " " + postingDate.ToString("dd/MM/yyyy") + "  " + basText.alignmentLeft(trnsDesc, 46) + basText.alignmentRight(strForeignCurr, 15) + basText.alignmentRight((object)basText.formatNumUnSign(detailRow[dBilltranamount], "#0.00", 9), 9) + " " + CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + " " + basText.alignmentRight((object)basText.formatNumUnSign((object)totTrans, "#0.00", 11), 11) + "  " + CrDb(totTrans)); //isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))
        }
        else if (isInstallment == true)
        {
            if (detailRow[dTrandescription].ToString() == "Amount Transferred for Refinancing")//"Installment Debt Transfer")//"Installment Transfer")
                strmWriteCommon.WriteLine(trnsDate.ToString("dd/MM/yyyy") + " " + postingDate.ToString("dd/MM/yyyy") + "  " + basText.alignmentLeft(trnsDesc, 46) + basText.alignmentRight(strForeignCurr, 15) + basText.alignmentRight((object)basText.formatNumUnSign(detailRow[dBilltranamount], "#0.00", 9), 9) + "    " + basText.alignmentRight((object)basText.formatNumUnSign((object)totTrans, "#0.00", 11), 11) + "  " + CrDb(totTrans) + basText.alignmentRight((object)basText.formatNumUnSign((object)totInstTrans, "#0.00", 13), 13) + "  " + CrDb(totInstTrans)); //isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))

            else if ((trnsDesc == "Installment"))// || (trnsDesc == "Installment Acceleration"))//"Installment Repayments")
            {
                if (detailRow[dInstallmentData].ToString().Trim() != "")
                {
                    input = detailRow[dInstallmentData].ToString().Trim();
                    installdata = input;
                    int index = detailRow[dInstallmentData].ToString().Trim().IndexOf(":");
                    if (index > 0)
                        input = input.Substring(0, index);
                    long installdocno = long.Parse(detailRow[dDocno].ToString());
                    long originstalldocno = long.Parse(detailRow[dOrigDocNo].ToString());
                    strmWriteCommon.WriteLine(trnsDate.ToString("dd/MM/yyyy") + " " + postingDate.ToString("dd/MM/yyyy") + "  " + basText.alignmentLeft(trnsDesc + input.Substring(basText.GetNthIndex(input, ',', 2)).Substring(input.Substring(basText.GetNthIndex(input, ',', 2)).IndexOf('#')).Replace('#', ' ') + " " + basText.trimStr(GetMerchantByOrigDocNo(long.Parse(detailRow[dOrigDocNo].ToString())), 40), 46) + basText.alignmentRight(strForeignCurr, 15) + basText.alignmentRight((object)basText.formatNumUnSign(detailRow[dBilltranamount], "#0.00", 9), 9) + " " + CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + " " + basText.alignmentRight((object)basText.formatNumUnSign((object)totTrans, "#0.00", 11), 11) + "  " + CrDb(totTrans) + basText.alignmentRight((object)basText.formatNumUnSign((object)totInstTrans, "#0.00", 13), 13) + "  " + CrDb(totInstTrans)); //isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))
                }
                else
                    strmWriteCommon.WriteLine(trnsDate.ToString("dd/MM/yyyy") + " " + postingDate.ToString("dd/MM/yyyy") + "  " + basText.alignmentLeft(trnsDesc, 46) + basText.alignmentRight(strForeignCurr, 15) + basText.alignmentRight((object)basText.formatNumUnSign(detailRow[dBilltranamount], "#0.00", 9), 9) + "    " + basText.alignmentRight((object)basText.formatNumUnSign((object)totTrans, "#0.00", 11), 11) + "  " + CrDb(totTrans) + basText.alignmentRight((object)basText.formatNumUnSign((object)totInstTrans, "#0.00", 13), 13) + "  " + CrDb(totInstTrans)); //isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))
            }
            else if ((trnsDesc == "Installment Finance Charges"))// || (trnsDesc == "Acceleration Finance Charges"))
            {
                int index = installdata.Trim().IndexOf(":");
                //if (int.Parse(detailRow[dDocno].ToString()) == installdocno)
                {
                    if (index > 0 && index != -1)
                    {
                        installdata = installdata.Substring(0, index);
                        strmWriteCommon.WriteLine(trnsDate.ToString("dd/MM/yyyy") + " " + postingDate.ToString("dd/MM/yyyy") + "  " + basText.alignmentLeft(trnsDesc + installdata.Substring(basText.GetNthIndex(installdata, ',', 2)).Substring(installdata.Substring(basText.GetNthIndex(installdata, ',', 2)).IndexOf('#')).Replace('#', ' ') + " " + basText.trimStr(GetMerchantByOrigDocNo(originstalldocno), 40), 46) + basText.alignmentRight(strForeignCurr, 15) + basText.alignmentRight((object)basText.formatNumUnSign(detailRow[dBilltranamount], "#0.00", 9), 9) + " " + CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + " " + basText.alignmentRight((object)basText.formatNumUnSign((object)totTrans, "#0.00", 11), 11) + "  " + CrDb(totTrans) + basText.alignmentRight((object)basText.formatNumUnSign((object)totInstTrans, "#0.00", 13), 13) + "  " + CrDb(totInstTrans)); //isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))
                    }
                    else
                        strmWriteCommon.WriteLine(trnsDate.ToString("dd/MM/yyyy") + " " + postingDate.ToString("dd/MM/yyyy") + "  " + basText.alignmentLeft(trnsDesc, 46) + basText.alignmentRight(strForeignCurr, 15) + basText.alignmentRight((object)basText.formatNumUnSign(detailRow[dBilltranamount], "#0.00", 9), 9) + "    " + basText.alignmentRight((object)basText.formatNumUnSign((object)totTrans, "#0.00", 11), 11) + "  " + CrDb(totTrans) + basText.alignmentRight((object)basText.formatNumUnSign((object)totInstTrans, "#0.00", 13), 13) + "  " + CrDb(totInstTrans)); //isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))
                }
                //else
                //    {
                //    if (index > 0)
                //        installdata = installdata.Substring(0, index);
                //    strmWriteCommon.WriteLine(trnsDate.ToString("dd/MM/yyyy") + " " + postingDate.ToString("dd/MM/yyyy") + "  " + basText.alignmentLeft(trnsDesc + installdata.Substring(basText.GetNthIndex(installdata, ',', 2)).Substring(installdata.Substring(basText.GetNthIndex(installdata, ',', 2)).IndexOf('#')).Replace('#', ' ') + " " + basText.trimStr(GetMerchantByOrigDocNo(originstalldocno), 40), 46) + basText.alignmentRight(strForeignCurr, 15) + basText.alignmentRight((object)basText.formatNumUnSign(detailRow[dBilltranamount], "#0.00", 9), 9) + "  " + CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + basText.alignmentRight((object)basText.formatNumUnSign((object)totTrans, "#0.00", 11), 11) + "  " + CrDb(totTrans) + basText.alignmentRight((object)basText.formatNumUnSign((object)totInstTrans, "#0.00", 13), 13) + "  " + CrDb(totInstTrans)); //isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))
                //    }
            }
            else
                strmWriteCommon.WriteLine(trnsDate.ToString("dd/MM/yyyy") + " " + postingDate.ToString("dd/MM/yyyy") + "  " + basText.alignmentLeft(trnsDesc, 46) + basText.alignmentRight(strForeignCurr, 15) + basText.alignmentRight((object)basText.formatNumUnSign(detailRow[dBilltranamount], "#0.00", 9), 9) + " " + CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + " " + basText.alignmentRight((object)basText.formatNumUnSign((object)totTrans, "#0.00", 11), 11) + "  " + CrDb(totTrans) + basText.alignmentRight((object)basText.formatNumUnSign((object)totInstTrans, "#0.00", 13), 13) + "  " + CrDb(totInstTrans)); //isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))
        }
        else
            strmWriteCommon.WriteLine(trnsDate.ToString("dd/MM/yyyy") + " " + postingDate.ToString("dd/MM/yyyy") + "  " + basText.alignmentLeft(trnsDesc, 46) + basText.alignmentRight(strForeignCurr, 15) + basText.alignmentRight((object)basText.formatNumUnSign(detailRow[dBilltranamount], "#0.00", 9), 9) + " " + CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + " " + basText.alignmentRight((object)basText.formatNumUnSign((object)totTrans, "#0.00", 11), 11) + "  " + CrDb(totTrans)); //isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))


        hasInterset = CalcTransInterest(detailRow[dTrandescription].ToString());
        totNoOfTransactions++;
    }

    //protected void printDetailSummary()
    //    {
    //    DateTime trnsDate = clsBasValid.validateDate(detailRow[dTransdate]), postingDate = clsBasValid.validateDate(detailRow[dPostingdate]);
    //    if (trnsDate > postingDate)
    //        trnsDate = postingDate;

    //    if (masterRow[mAccountcurrency].ToString().Trim() != detailRow[dOrigtrancurrency].ToString().Trim())
    //        strForeignCurr = basText.alignmentRight((Object)basText.formatNum(detailRow[dOrigtranamount], "#0.00;(#0.00)") + " " + basText.Replace(clsBasValid.validateStr(detailRow[dOrigtrancurrency]), "XXX", "   "), 15);
    //    else
    //        strForeignCurr = basText.replicat(" ", 15);

    //    if (strForeignCurr.Trim() == "0")
    //        strForeignCurr = basText.replicat(" ", 15);

    //    string trnsDesc; //= detailRow[dTrandescription].ToString().Trim() + " " + detailRow[dMerchant].ToString().Trim()
    //    if (detailRow[dMerchant].ToString().Trim() == "")
    //        trnsDesc = detailRow[dTrandescription].ToString().Trim();
    //    else
    //        trnsDesc = detailRow[dMerchant].ToString().Trim();

    //    totTrans += valDbCr(Convert.ToDecimal(detailRow[dBilltranamount]), clsBasValid.validateStr(detailRow[dBilltranamountsign])); //basText.formatNum(detailRow[dBilltranamount], "#0.00;(#0.00)", 19);
    //    totNoOfTransactions++;
    //    }

    protected void printCardFooter()
    {
        if (flag == false)
        {
            streamlog.WriteLine(masterRow[mStatementno].ToString());
            flag = true;
        }
        streamlog.WriteLine("cardfooter " + flag);
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        //strmWriteCommon.WriteLine(basText.replicat(" ", 12) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 38, "L") + basText.alignmentLeft((object)basText.formatNumUnSign(masterRow[mMindueamount], "#0.00", 44), 44) + basText.alignmentRight((Object)basText.formatNumUnSign(masterRow[mClosingbalance], "#0.00", 19), 19) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])));
        if (isInstallment == true && !InstallmentCondition.Contains("BDC Easy Payment Plan"))
        {
            installmentRows = DSInstallment.Tables["installment"].Select("CLIENTID = '" + clsBasValid.validateStr(masterRow[mClientid]).ToString().Trim() + "' and CREDITCONTRACTS = '" + clsBasValid.validateStr(masterRow[mContractno]).ToString().Trim() + "'");
            if (installmentRows.Length > 0)
            {
                installmentRow = installmentRows[0];
                strmWriteCommon.WriteLine(basText.replicat(" ", 12) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 12, "L") + basText.alignmentRight((object)basText.formatNumUnSign(masterRow[mMindueamount], "#0.00", 32), 32) + basText.alignmentRight((Object)basText.formatNumUnSign(decimal.Parse(masterRow[mClosingbalance].ToString()) + decimal.Parse(installmentRow[mClosingbalance].ToString()), "#0.00", 19), 57) + " " + CrDb(decimal.Parse(masterRow[mClosingbalance].ToString()) + decimal.Parse(installmentRow[mClosingbalance].ToString())), 20);
            }
            else
                strmWriteCommon.WriteLine(basText.replicat(" ", 12) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 12, "L") + basText.alignmentRight((object)basText.formatNumUnSign(masterRow[mMindueamount], "#0.00", 32), 32) + basText.alignmentRight((Object)basText.formatNumUnSign(decimal.Parse(masterRow[mClosingbalance].ToString()) + decimal.Parse("0"), "#0.00", 19), 57) + " " + CrDb(decimal.Parse(masterRow[mClosingbalance].ToString()) + decimal.Parse("0")), 20);
        }
        else
            strmWriteCommon.WriteLine(basText.replicat(" ", 12) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 12, "L") + basText.alignmentRight((object)basText.formatNumUnSign(masterRow[mMindueamount], "#0.00", 32), 32) + basText.alignmentRight((Object)basText.formatNumUnSign(masterRow[mClosingbalance], "#0.00", 19), 57) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])));
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(basText.replicat(" ", 10) + basText.alignmentLeft(basText.formatCardNumber(curMainCard), 16));
        strmWriteCommon.WriteLine(String.Empty);
        //    strmWriteCommon.WriteLine(basText.replicat(" ", 10) + basText.alignmentLeft(basText.formatNumUnSign(masterRow[mTotalpayments], "#,##0.00", 12) + " " + DbCr(Convert.ToDecimal(masterRow[mTotalpayments])), 15));
        //strmWriteCommon.WriteLine(basText.replicat(" ", 10) + basText.alignmentLeft((object)(masterRow[mCardpaymentmethod].ToString().Trim() == "" || masterRow[mCardpaymentmethod].ToString().Trim() == "Not Defined" ? "Cash" : masterRow[mCardpaymentmethod].ToString().Trim()), 30)); //, "#,##0.00", 12) + " " + DbCr(Convert.ToDecimal(masterRow[mTotalpayments])), 15)
        strmWriteCommon.WriteLine(basText.replicat(" ", 10) + basText.alignmentLeft((object)masterRow[mExternalno], 30)); //, "#,##0.00", 12) + " " + DbCr(Convert.ToDecimal(masterRow[mTotalpayments])), 15)
        strmWriteCommon.WriteLine(String.Empty);
        if (isInstallment == true && !InstallmentCondition.Contains("BDC Easy Payment Plan"))
        {
            installmentRows = DSInstallment.Tables["installment"].Select("CLIENTID = '" + clsBasValid.validateStr(masterRow[mClientid]).ToString().Trim() + "' and CREDITCONTRACTS = '" + clsBasValid.validateStr(masterRow[mContractno]).ToString().Trim() + "'");
            if (installmentRows.Length > 0)
            {
                installmentRow = installmentRows[0];
                strmWriteCommon.WriteLine(basText.replicat(" ", 10) + basText.alignmentLeft((Object)basText.formatNumUnSign(decimal.Parse(masterRow[mClosingbalance].ToString()) + decimal.Parse(installmentRow[mClosingbalance].ToString()), "#0.00", 15) + " " + CrDb(decimal.Parse(masterRow[mClosingbalance].ToString()) + decimal.Parse(installmentRow[mClosingbalance].ToString())), 20));
            }
            else
                strmWriteCommon.WriteLine(basText.replicat(" ", 10) + basText.alignmentLeft((Object)basText.formatNumUnSign(masterRow[mClosingbalance], "#0.00", 15) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])), 20));
        }
        else
            strmWriteCommon.WriteLine(basText.replicat(" ", 10) + basText.alignmentLeft((Object)basText.formatNumUnSign(masterRow[mClosingbalance], "#0.00", 15) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])), 20));
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(basText.replicat(" ", 10) + basText.formatNum(masterRow[mMindueamount], "#0.00", 28, "L"));
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(basText.replicat(" ", 10) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 37, "L"));
        //completePageDetailRecords();
        //if (pageNo == totalAccPages)
        //  strmWriteCommon.WriteLine(basText.replicat(" ", 17) + basText.alignmentLeft("Current Balance", 63) + basText.formatNumUnSign(masterRow[mClosingbalance], "#0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])));
        //else
        //  strmWriteCommon.WriteLine(String.Empty);
        //    strmWriteCommon.WriteLine(basText.alignmentRight(curMainCard, 35) + basText.alignmentRight(masterRow[mMindueamount], 20) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "R")
        //   + basText.alignmentRight(basText.formatNumUnSign(masterRow[mOpeningbalance], "##0.00", 12) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance].ToString())), 15)
        //   + basText.alignmentRight(basText.formatNum(masterRow[mTotalpayments], "#,##0.00", 12) + " " + DbCr(Convert.ToDecimal(masterRow[mTotalpayments])), 15)
        //   + basText.alignmentRight(basText.formatNum(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal]), "#,##0.00", 12) + CrDbMinus(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal])), 15)
        //   + basText.alignmentRight(basText.formatNum(masterRow[mTotalcharges].ToString(), "#,##0.00", 12) + " " + CrDbMinus(Convert.ToDecimal(masterRow[mTotalcharges])), 15)
        //   + basText.alignmentRight(basText.formatNum(masterRow[mTotalinterest].ToString(), "#,##0.00", 12) + " " + CrDbMinus(Convert.ToDecimal(masterRow[mTotalinterest])), 15)
        //       + basText.formatNum(Convert.ToDecimal(masterRow[mTotalcharges]) + Convert.ToDecimal(masterRow[mTotalinterest]),"#,##0.00",12) + CrDbMinus(Convert.ToDecimal(masterRow[mTotalcharges]) + Convert.ToDecimal(masterRow[mTotalinterest])) 
        //strmWriteCommon.WriteLine(String.Empty);
        //strmWriteCommon.WriteLine(String.Empty);
        //    strmWriteCommon.WriteLine(basText.alignmentLeft(basText.formatNumUnSign(masterRow[mOpeningbalance], "#,###,##0.00", 12) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance].ToString())), 20) +
        //      basText.alignmentLeft(basText.formatNum(masterRow[mTotalcharges].ToString(), "#,###,##0.00", 12), 20) +
        //      basText.alignmentLeft(basText.formatNumUnSign(Convert.ToDecimal(masterRow[mTotalpurchases]) +
        //      Convert.ToDecimal(masterRow[mTotalcashwithdrawal]), "#,###,##0.00", 16) + " " +
        //      CrDbMinus(Convert.ToDecimal(masterRow[mTotalpurchases]) +
        //      Convert.ToDecimal(masterRow[mTotalcashwithdrawal])), 20) +
        //basText.alignmentLeft(basText.formatNumUnSign(masterRow[mClosingbalance], "#,###,##0.00", 12) + " " +
        //CrDb(Convert.ToDecimal(masterRow[mClosingbalance])), 20));
        //strmWriteCommon.WriteLine(String.Empty);
        //strmWriteCommon.WriteLine(String.Empty);
        //strmWriteCommon.WriteLine(basText.replicat(" ", 81) + basText.alignmentLeft(masterRow[mExternalno], 20));  //cardproduct+"{0,5:MM/yy}",masterRow[mStatementdatefrom]
        //strmWriteCommon.WriteLine(basText.alignmentLeft(masterRow[mCustomername], 50));  //
        //strmWriteCommon.WriteLine(basText.replicat(" ", 81) + basText.formatNum(curMainCard, "#### #### #### ####"));  //cardbranchpartname
        //strmWriteCommon.WriteLine(basText.alignmentLeft(ValidateArbic(masterRow[mCustomeraddress1].ToString()), 50)); //basText.formatNum(masterRow[mCardno],"####-####-####-####")
        //strmWriteCommon.WriteLine(basText.alignmentLeft(ValidateArbic(masterRow[mCustomeraddress2].ToString()), 50));
        //strmWriteCommon.WriteLine(basText.alignmentLeft(ValidateArbic(masterRow[mCustomeraddress3].ToString().Trim()) + " " + masterRow[mCustomerregion].ToString().Trim() + " " + masterRow[mCustomercity].ToString().Trim(), 50));  //basText.formatNum(masterRow[mAccountno],"##-##-######-####")
        //strmWriteCommon.WriteLine(basText.replicat(" ", 81) + basText.alignmentLeft(basText.formatNumUnSign(masterRow[mTotalpayments], "#,##0.00", 12) + " " + DbCr(Convert.ToDecimal(masterRow[mTotalpayments])), 15));  //cardproduct+"{0,5:MM/yy}",masterRow[mStatementdatefrom]
        //strmWriteCommon.WriteLine(String.Empty);

    }

    //protected void printAccountFooter()
    //    {
    //    strmWriteCommon.WriteLine("GRAND" + basText.formatNumUnSign(masterRow[mOpeningbalance], "##0.00", 16) + basText.formatNum(masterRow[mTotaldebits], "#0.00;(#0.00)", 16) + basText.formatNum(masterRow[mTotalcredits], "#0.00;(#0.00)", 16) + basText.formatNum(masterRow[mClosingbalance], "#0.00;(#0.00)", 16) + basText.formatNum(masterRow[mMindueamount], "#0.00;(#0.00)", 16) + basText.replicat(" ", 12) + basText.formatDate(masterRow[mStetementduedate], "yyyy/MM/dd") + "\n");//+strEndOfPage
    //    }


    private void calcCardlRows()
    {
        totalCardPages = 0;
        totCardRows = 0;
        foreach (DataRow dtRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
        {
            if ((dtRow[dPostingdate] == DBNull.Value) && (dtRow[dDocno] == DBNull.Value)) continue;
            //strmWriteCommon.WriteLine(basText.trimStr(dtRow[dTrandescription],40)); 
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
        string prevCardNo = String.Empty, CurCardNo = String.Empty;
        int currAccRowsPages = 1;//0
        totAccRows = 1;//0
        totalAccPages = 0;
        //currAccRowsPages = 0;
        foreach (DataRow dtAccRow in accountRows) //mRow.GetChildRows(StatementNoDRel)
        {
            if ((dtAccRow[dPostingdate] == DBNull.Value) && (dtAccRow[dDocno] == DBNull.Value))
                continue;
            CurCardNo = dtAccRow[dCardno].ToString();
            //if (CurCardNo.Trim().Length < 1) continue;

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
        totSuppCrdsInAcc = 0; //BDCA-2390
        totCrdNoInAcc = curCrdNoInAcc = 0;
        mainRows = DSstatement.Tables["tStatementMasterTable"].Select("ACCOUNTNO = '" + clsBasValid.validateStr(masterRow[mAccountno]) + "'");
        curMainCard = CurCardNo = strExpiryDate = "";
        foreach (DataRow mainRow in mainRows) //mRow.GetChildRows(StatementNoDRel)
        {
            totCrdNoInAcc++;
            CurCardNo = mainRow[mCardno].ToString();
            if (mainRow[mCardprimary].ToString() == "N")
            {
                if (isValidateCard(mainRow[mCardstate].ToString()))
                {
                    tmpDtlRows = DSstatement.Tables["tStatementDetailTable"].Select(dCardno + " = '" + CurCardNo + "' and " + dPostingdate + " is not null and " + dDocno + " is not null");
                    if (tmpDtlRows.Length > 0)
                        totSuppCrdsInAcc++;
                }
            }
            if (mainRow[mCardprimary].ToString() == "Y")
                curMainCard = CurCardNo; //mainRow[mCardno].ToString();
            if (mainRow[mCardprimary].ToString() == "Y" && isValidateCard(mainRow[mCardstate].ToString()))
            {
                curMainCard = CurCardNo; //mainRow[mCardno].ToString();
            }
            else
            {
                if (mainRow[mCardprimary].ToString() == "Y")
                    curMainCard = CurCardNo; //mainRow[mCardno].ToString();
                if (mainRow[mCardprimary].ToString() == "Y" && isValidateCard(mainRow[mCardstate].ToString()))
                {
                    curMainCard = CurCardNo; //mainRow[mCardno].ToString();
                    break;
                }
            }
        }


        if (curMainCard == "")
            curMainCard = CurCardNo;

        if (totSuppCrdsInAcc > 0)
        {
            totAccRows += totSuppCrdsInAcc;
            totalAccPages = totAccRows / MaxDetailInPage;
            totalAccPages = (totAccRows % MaxDetailInPage) >= 0 ? ++totalAccPages : totalAccPages;
        }
    }


    private void completePageDetailRecords()
    {
        //int curPageLine =CurPageRec4Dtl;
        if (pageNo == totalAccPages)
            for (int curPageLine = CurPageRec4Dtl; curPageLine < MaxDetailInPage; curPageLine++)
                strmWriteCommon.WriteLine(String.Empty);
    }

    private void printStatementSummary()
    {
        if (strBankName.Contains("MasterCard"))
            streamSummary.WriteLine(strBankName + " MasterCard Statement");
        else
            streamSummary.WriteLine(strBankName + " Visa Statement");

        streamSummary.WriteLine("__________________________");
        streamSummary.WriteLine("");
        streamSummary.WriteLine("No of Statements   " + totNoOfCardStat.ToString());
        streamSummary.WriteLine("No of Pages        " + totNoOfPageStat.ToString());
        streamSummary.WriteLine("No of Transactions " + totNoOfTransactions.ToString());
        //streamSummary.WriteLine("No of Loops        " + count.ToString());

        clsStatementSummary StatSummary = new clsStatementSummary();
        StatSummary.BankCode = curBranchVal;
        StatSummary.BankName = strBankName;
        StatSummary.StatementDate = vCurDate;
        StatSummary.CreationDate = DateTime.Now;
        DataRow[] productRows;
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

                if (GetAccountsWithoutTrxns(masterRow[mAccountno].ToString()))
                    continue;
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

    private bool GetValidCardByAccount(string pAccount)
    {
        mainRows = null;
        bool result = false;
        execluded = false;
        DSstatement.Tables["tStatementMasterTable"].DefaultView.Sort = "cardstate desc";
        //mainRows = DSstatement.Tables["tStatementMasterTable"].Select("ACCOUNTNO = '" + pAccount + "' and ClosingBalance >= -50"); //Bug
        //mainRows = DSstatement.Tables["tStatementMasterTable"].Select("ACCOUNTNO = '" + pAccount + "' and ClosingBalance >= -50 and closingbalance <= 0"); //Bug
        //mainRows = DSstatement.Tables["tStatementMasterTable"].Select("ACCOUNTNO = '" + pAccount + "' and ClosingBalance <= 50 and closingbalance >= 0"); //Former
        mainRows = DSstatement.Tables["tStatementMasterTable"].Select("ACCOUNTNO = '" + pAccount + "' and ClosingBalance >= 0"); //Current
        if (mainRows.Length == 0)
            return false;
        //mainRows = DSstatement.Tables["tStatementMasterTable"].Select("ACCOUNTNO = '" + pAccount + "' and cardstate = 'Cancelled'");
        foreach (DataRow mainRow in mainRows)
        {
            //if (mainRow[mCardprimary].ToString() == "Y" && (mainRow[mCardstate].ToString() == "Closed" || mainRow[mCardstate].ToString() == "Cancelled"))
            if (mainRow[mCardprimary].ToString() == "Y" && mainRow[mCardstate].ToString() == "Cancelled")
            //if (mainRow[mCardprimary].ToString() == "Y" && !isValidateCard(mainRow[mCardstate].ToString()))
            //if (mainRow[mCardprimary].ToString() == "Y" && (int.Parse(mainRow[mClosingbalance].ToString()) >= 0 && Math.Abs(int.Parse(mainRow[mClosingbalance].ToString())) <= 50))
            {
                result = true;
            }
            else
            {
                result = false;
                break;
            }
            if (excludedaccountno != pAccount)
            {
                excludedaccountno = pAccount;
                excludedcardno = mainRow[mCardno].ToString();
                excludedcardmbr = mainRow[mMBR].ToString();
                excludedcardstate = mainRow[mCardstate].ToString();
            }
        }
        return result;
    }

    private bool GetAccountsWithoutTrxns(string pAccount)
    {
        //BDCA-3101 EDT-878
        mainRows = null;
        bool result = false;
        execluded = false;
        mainRows = DSstatement.Tables["tStatementMasterTable"].Select("ACCOUNTNO = '" + pAccount + "' and ClosingBalance >= 0 and OpeningBalance = ClosingBalance"); //Current
        if (mainRows.Length == 0)
            return false;
        foreach (DataRow mainRow in mainRows)
        {
            tmpDtlRows = DSstatement.Tables["tStatementDetailTable"].Select("ACCOUNTNO = '" + pAccount + "' and " + dPostingdate + " is not null and " + dDocno + " is not null");
            if (tmpDtlRows.Length == 0)
            {
                result = true;
            }
            else
            {
                result = false;
                break;
            }
            if (excludedaccountno != pAccount)
            {
                excludedaccountno = pAccount;
                excludedcardno = mainRow[mCardno].ToString();
                excludedcardmbr = mainRow[mMBR].ToString();
                excludedcardstate = mainRow[mCardstate].ToString();
            }
        }
        return result;
    }


    private string GetMerchantByOrigDocNo(long pDocno)
    {
        mainRows = null;
        string result = string.Empty;
        mainRows = DSstatement.Tables["tStatementDetailTable"].Select("ORIGDOCNO = " + pDocno + "");
        foreach (DataRow mainRow in mainRows)
        {
            result = mainRow[dInstallmentMerchantLocation].ToString();
        }
        return result;
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

    public bool savedDataset
    {
        get { return isSavedDataset; }
        set { isSavedDataset = value; }
    }// savedDataset

    public string whereCond
    {
        get { return strWhereCond; }
        set { strWhereCond = value; }
    }  // whereCond

    public string productCond
    {
        get { return strProductCond; }
        set { strProductCond = value; }
    }  // productCond

    public int _totNoOfCardStat
    {
        get { return totNoOfCardStat; }
    }  // _totNoOfCardStat

    public int _totNoOfPageStat
    {
        get { return totNoOfPageStat; }
    }  // _totNoOfPageStat

    public int _totNoOfTransactions
    {
        get { return totNoOfTransactions; }
    }  // _totNoOfTransactions

    public string InstallmentCondition
    {
        get { return installmentCond; }
        set { installmentCond = value; }
    }// InstallmentCondition

    public bool isInstallmentVal
    {
        get { return isInstallment; }
        set { isInstallment = value; }
    }// isInstallmentVal

    public string statMessageBox
    {
        get { return statMessageBoxVal; }
        set { statMessageBoxVal = value; }
    }//statMessageBox

    //public bool emailService
    //{
    //    get { return isInEmailService; }
    //    set { isInEmailService = value; }
    //}// emailService

    ~clsStatementBDCA()
    {
        DSstatement.Dispose();
    }
}
