using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Xml;
using System.Collections;

// Branch 16
public class clsCalcStatSummary : clsStatTxtLbl
{
    private string statTypeVal = "Text";
    private DataRow ProductRow;
    DataRow[] productRows;
    public clsCalcStatSummary()
    {
    }

    public string Statement(string pStrFileName, string pBankName, int pBankCode, string pStrFile, DateTime pCurDate, string pStmntType, bool pAppendData, DataSet pDSstatement, DataSet pDSproducts)
    {
        DSstatement = pDSstatement;
        DSProducts = pDSproducts;
        string rtrnStr = "Successfully Generate " + pBankName, ppStrFileName = pStrFileName;
        int curMonth = pCurDate.Month;
        stmntType = pStmntType;
        if (isReward)
            NoLinePerPage = 57;


        try
        {

            pStrFileName = clsBasFile.makeStrAsPath(pStrFileName);
            vCurDate = pCurDate; //DateTime.Now.AddMonths(-1);
            strOutputPath = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName;
            //clsBasFile.createDirectory(pStrFileName + vCurDate.ToString("yyyyMM") + pBankName);
            pStrFileName = pStrFileName + vCurDate.ToString("yyyyMM") + pBankName + "\\" + pBankName + pStrFile + vCurDate.ToString("yyyyMM") + ".txt";
            strBankName = pBankName;
            // open Summary file
            fileSummaryName = pStrFileName;
            fileSummaryName = clsBasFile.getPathWithoutExtn(fileSummaryName) +
              "_Summary." + clsBasFile.getFileExtn(fileSummaryName);
            fileSummary = new FileStream(fileSummaryName, FileMode.Create); //Create
            streamSummary = new StreamWriter(fileSummary, Encoding.Default);


            // set branch for data
            curBranchVal = pBankCode; // 16; //3 = real   1 = test
            //     clsBasStatement.mainTableCond =  " substr(cardno,1,6) !='421192' "; //" cardproduct != 'Visa Business' ";
            //     clsBasStatement.supTableCond =  " substr(cardno,1,6) !='421192' "; 
            if (createCorporateVal)
            {
                isCorporateVal = true;
                accountNoName = mCardaccountno;
                accountLimit = mCardlimit;
                accountAvailableLimit = mCardavailablelimit;
            }
            // data retrieve
            //FillStatementDataSet(pBankCode); //DSstatement =  //16); //3
            if (isInEmailService && isNotSplit)
                getClientEmail(pBankCode);
            if (isReward)
                getReward(pBankCode);
            pageNo = 1; totalCardPages = 0;
            curCardNo = String.Empty;
            curAccountNo = String.Empty;

            ////frmMain.BeginInvoke(frmMain.setMinMaxProgressDelegate, new object[] { DSstatement.Tables["tStatementMasterTable"].Rows.Count });
            //foreach (DataRow mRow in DSstatement.Tables["tStatementMasterTable"].Rows)
            //{
            //    //frmMain.BeginInvoke(frmMain.setProgressDelegate, new object[] { totRec++ });
            //    masterRow = mRow;
            //    productRows = DSProducts.Tables["Products"].Select("name = '" + masterRow[mCardproduct].ToString().Trim() + "'");
            //    //streamWrit.WriteLine(masterRow[mStatementno].ToString());
            //    //pageNo=1; CurPageRec4Dtl=0; CurrentPageFlag = "F 1"; //if page is based on card no
            //    cardsRows = mRow.GetChildRows(StatementNoDRel); //, DataRowVersion.Proposed

            //    strCardNo = masterRow[mCardno].ToString().Trim();
            //    if (strCardNo == "")
            //    {
            //        continue;
            //    }
            //    if (strCardNo.Length != 16)
            //    {
            //        //numOfErr++;
            //        continue;// Exclude Zero Length Cards 
            //    }
            //    if ((clsCheckCard.CalcLuhnCheckDigit(strCardNo.Substring(0, 15)) != strCardNo.Substring(15, 1))) //!clsCheckCard.isValidCard(strCardNo) ||
            //    {
            //        numOfErr++;
            //    }

            //    strPrimaryCardNo = strCardNo;
            //    if (masterRow[mCardprimary].ToString() == "N")
            //    {
            //        strPrimaryCardNo = masterRow[mPrinarycardno].ToString();
            //        //calcCardlRows();
            //    }

            //    //start new account
            //    if (prevAccountNo != masterRow[accountNoName].ToString())
            //    {
            //        if (isInEmailService)
            //        {
            //            if (haveEmail(masterRow[mClientid].ToString()))
            //            {
            //                isEmailStat = true;
            //            }
            //            else
            //            {
            //                isEmailStat = false;
            //            }
            //        }

            //        if (pageNo != totalAccPages && prevAccountNo != "")// 
            //        {
            //            //>MessageBox.Show( "Error in Genrating Statement", "Please Call The Programmer", MessageBoxButtons.OK ,
            //            //>  MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1); //,	MessageBoxOptions.RightAlign
            //            numOfErr++;
            //        }

            //        curMainCard = string.Empty;
            //        if (!isHaveF3)//!isHaveF3  CurrentPageFlag != "F 2"
            //        {
            //            numOfErr++;
            //        }
            //        isHaveF3 = false;

            //        pageNo = 1;//totalAccPages = 1 ; pageNo=1;
            //        CurPageRec4Dtl = 0; totNetUsage = 0; CurrentPageFlag = "F 1"; //if page is based on account no
            //        calcAccountRows();

            //        if (!isDebit)
            //            if (totAccRows < 1
            //              && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0) //             || (masterRow[mCardno].ToString() == curMainCard   // Convert.ToDecimal(
            //            {
            //                isHaveF3 = true;
            //                //pageNo=1; totalAccPages =1;
            //                continue;
            //            }

            //        prevAccountNo = masterRow[accountNoName].ToString();
            //        //pageNo=1; //if page is based on account no
            //        printHeader();//if page is based on account no

            //        if (isEmailStat)
            //            totNoOfCardEmailStat++;
            //        else
            //            totNoOfCardStat++;

            //        if (Convert.ToDecimal(masterRow[mAccountlim].ToString()) <= Convert.ToDecimal(0))
            //        {
            //            numOfErr++;
            //        }
            //    } // End of if(prevAccountNo != masterRow[accountNoName].ToString())
            //    //calcCardlRows();
            //    //if(totCardRows < 1)continue ;  //if pages is based on card
            //    //printHeader();//if pages is based on card

            //    foreach (DataRow dRow in cardsRows) //mRow.GetChildRows(StatementNoDRel)
            //    {
            //        detailRow = dRow;
            //        stmNo = detailRow[dStatementno].ToString();
            //        if ((detailRow[dPostingdate] == DBNull.Value) && (detailRow[dDocno] == DBNull.Value)) continue;// Exclude On-Hold Transactions 
            //        curAccRows++;
            //        if (CurPageRec4Dtl >= MaxDetailInPage)
            //        {
            //            CurPageRec4Dtl = 0;
            //            printCardFooter();
            //            pageNo++;
            //            printHeader();
            //        }
            //        //						totNetUsage +=clsBasValid.validateNum(detailRow[dBilltranamount]);
            //        totNetUsage = calculateCrDb(totNetUsage, clsBasValid.validateNum(detailRow[dBilltranamount]), clsBasValid.validateStr(detailRow[dBilltranamountsign]));
            //        CurPageRec4Dtl = CurPageRec4Dtl + 1;
            //        printDetail();
            //        //if(!((detailRow[dPostingdate]== DBNull.Value) && (detailRow[dDocno]== DBNull.Value))) 

            //    } //end of detail foreach
            //    //printCardFooter();//if pages is based on card
            //    // if(masterRow[mCardprimary].ToString() == "Y" && (masterRow[mCardstate].ToString() == "Given" || masterRow[mCardstate].ToString() == "Embossed" || masterRow[mCardstate].ToString() == "New"))
            //    curCrdNoInAcc++;
            //    if ((curAccRows >= totAccRows && totAccRows != 0) || (totAccRows == 0 && curCrdNoInAcc == totCrdNoInAcc))
            //    {
            //        completePageDetailRecords();
            //        printCardFooter();//if pages is based on account
            //        //printAccountFooter();
            //        CurPageRec4Dtl = 0; //>pageNo=1; if pages is based on account
            //        curAccRows = 0;
            //    }
            //    //streamWrit.WriteLine(strEndOfPage);
            //    //pageNo=1; CurPageRec4Dtl=0; //if pages is based on card
            //    //completePageDetailRecords();

            //} //end of Master foreach

            printStatementSummary();
        }
        catch (OracleException ex)
        {
            clsDbOracleLayer.catchError(ex);
        }
        catch (NotSupportedException ex)  //(Exception ex)  //
        {
            clsBasErrors.catchError(ex);
        }
        catch (Exception ex)
        {
            return rtrnStr;
        }
        finally
        {
        }
        return rtrnStr;
    }


    //protected override void printHeader()
    //    {
    //    //if(masterRow[mCardprimary].ToString() == "Y")
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
    //    totalPages++;
    //    if (isEmailStat)
    //        totNoOfPageEmailStat++;
    //    else
    //        totNoOfPageStat++;
    //    }

    //protected override void printDetail()
    //    {
    //    }

    private void printStatementSummary()
    {
        clsStatementSummary StatSummary = new clsStatementSummary();
        StatSummary.BankCode = curBranchVal;
        StatSummary.BankName = strBankName;
        StatSummary.StatementDate = vCurDate;
        StatSummary.CreationDate = DateTime.Now;
        totNoOfCardStat = 0;
        totNoOfCardEmailStat = 0;
        totNoOfTransactionsInt = 0;
        totNoOfTransEmailInt = 0;
        prevAccountNo = string.Empty;
        foreach (DataRow pRow in DSProducts.Tables["Products"].Rows)
        {
            ProductRow = pRow;
            productRows = DSstatement.Tables["tStatementMasterTable"].Select("cardproduct = '" + ProductRow[pName].ToString().Trim() + "'");
            if (productRows.Length == 0) continue;
            foreach (DataRow mRow in productRows)
            {
                masterRow = mRow;
                cardsRows = mRow.GetChildRows(StatementNoDRel); //, DataRowVersion.Proposed
                //if (cardsRows.Length == 0) continue;
                strCardNo = masterRow[mCardno].ToString().Trim();

                if ((strBankName == "EGB_Credit" || strBankName == "BRKA_Credit_PDF") && masterRow[mHOLSTMT].ToString() == "Y")
                {
                    continue;
                }

                if (strCardNo.Length != 16) continue;// Exclude Zero Length Cards 

                //if (!isValidateCard(masterRow[mCardstate].ToString())) continue;

                strPrimaryCardNo = strCardNo;
                if (masterRow[mCardprimary].ToString() == "N")
                {
                    strPrimaryCardNo = masterRow[mPrinarycardno].ToString();
                    //calcCardlRows();
                }

                //start new account
                if (prevAccountNo != masterRow[accountNoName].ToString())
                {
                    if (isInEmailService)
                    {
                        if (haveEmail(masterRow[mClientid].ToString()))
                        {
                            isEmailStat = true;
                        }
                        else
                        {
                            isEmailStat = false;
                        }
                    }
                    curMainCard = string.Empty;
                    isHaveF3 = false;
                    pageNo = 1;//totalAccPages = 1 ; pageNo=1;
                    CurPageRec4Dtl = 0; totNetUsage = 0; CurrentPageFlag = "F 1"; //if page is based on account no
                    calcAccountRows();
                    if (int.Parse(masterRow[mBranch].ToString()) == 23)
                    {
                        if (totAccRows < 1 && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) >= 0) continue;
                    }
                    else
                        if (totAccRows < 1 && Convert.ToDecimal(masterRow[mClosingbalance].ToString()) == 0) continue;

                    prevAccountNo = masterRow[accountNoName].ToString();
                    if (isEmailStat)
                        totNoOfCardEmailStat++;
                    else
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
                    hasInterset = CalcTransInterest(detailRow[dTrandescription].ToString());
                    if (hasInterset)
                    {
                        if (isEmailStat)
                            totNoOfTransEmailInt++;
                        else
                            totNoOfTransactionsInt++;
                        break;
                    }
                } //end of detail foreach


                curCrdNoInAcc++;
                if ((curAccRows >= totAccRows && totAccRows != 0) || (totAccRows == 0 && curCrdNoInAcc == totCrdNoInAcc))
                {
                    CurPageRec4Dtl = 0; //>pageNo=1; if pages is based on account
                    curAccRows = 0;
                }
            } //end of Master foreach
            if ((totNoOfCardStat + totNoOfCardEmailStat) != 0)
            {
                StatSummary.NoOfStatements = totNoOfCardStat + totNoOfCardEmailStat;
                StatSummary.NoOfTransactionsInt = totNoOfTransactionsInt + totNoOfTransEmailInt;
                StatSummary.ProductCode = int.Parse(ProductRow["code"].ToString());
                StatSummary.ProductName = ProductRow["name"].ToString();
                StatSummary.InsertRecordDb(vCurDate.ToString("yyyyMM") + strBankName);
                //if (strBankName.Contains("MasterCard"))
                //    streamSummary.WriteLine(strBankName + " Statement");
                //else
                streamSummary.WriteLine(strBankName + " Statement");
                if (isInEmailService)
                {
                    streamSummary.WriteLine("");
                    streamSummary.WriteLine("");
                    streamSummary.WriteLine("Statements not Sent by Email");
                }

                if (curBranchVal == 33)
                {
                    streamSummary.WriteLine("__________________________");
                    streamSummary.WriteLine("");
                    streamSummary.WriteLine("No of Statements   " + totNoOfCardStat.ToString());
                    streamSummary.WriteLine("No of Pages        " + (pageNo).ToString()); //totNoOfPageStat
                    streamSummary.WriteLine("No of Transactions " + (totNoOfTransactionsInt).ToString());

                }
                else
                {
                    streamSummary.WriteLine("__________________________");
                    streamSummary.WriteLine("");
                    streamSummary.WriteLine("No of Statements   " + totNoOfCardStat.ToString());
                    streamSummary.WriteLine("No of Pages        " + totNoOfPageStat.ToString());
                    streamSummary.WriteLine("No of Transactions " + totNoOfTransactions.ToString());
                }
                //streamSummary.WriteLine("__________________________");
                //streamSummary.WriteLine("");
                //streamSummary.WriteLine("No of Statements   " + totNoOfCardStat.ToString());
                //streamSummary.WriteLine("No of Pages        " + totNoOfPageStat.ToString());
                //streamSummary.WriteLine("No of Transactions " + totNoOfTransactions.ToString());

                if (isInEmailService)
                {
                    streamSummary.WriteLine("");
                    streamSummary.WriteLine("");
                    streamSummary.WriteLine("__________________________");
                    streamSummary.WriteLine("");
                    streamSummary.WriteLine("Statements Sent by Email");
                    streamSummary.WriteLine("");
                    streamSummary.WriteLine("No of Statements   " + totNoOfCardEmailStat.ToString());
                    streamSummary.WriteLine("No of Pages        " + totNoOfPageEmailStat.ToString());
                    streamSummary.WriteLine("No of Transactions " + totNoOfTransEmail.ToString());
                    streamSummary.WriteLine("");
                    streamSummary.WriteLine("");
                    streamSummary.WriteLine("__________________________");
                    streamSummary.WriteLine("");
                    streamSummary.WriteLine("Total Statements");
                    streamSummary.WriteLine("");
                    streamSummary.WriteLine("No of Statements   " + (totNoOfCardStat + totNoOfCardEmailStat).ToString());
                    streamSummary.WriteLine("No of Pages        " + (totNoOfPageStat + totNoOfPageEmailStat).ToString());
                    streamSummary.WriteLine("No of Transactions " + (totNoOfTransactions + totNoOfTransEmail).ToString());
                }
                totNoOfCardStat = 0;
                totNoOfCardEmailStat = 0;
                totNoOfTransactionsInt = 0;
                totNoOfTransEmail = 0;
            }

        }
        // Close Summary File
        streamSummary.Flush();
        streamSummary.Close();
        fileSummary.Close();
        StatSummary = null;
        DSProducts = null;
    }

    public string statType
    {
        set { statTypeVal = value; }
    }

    public DataRelation statRelation
    {
        set { StatementNoDRel = value; }
    }

    public bool isDebitVal
    {
        set { isDebit = value; }
    }
}
