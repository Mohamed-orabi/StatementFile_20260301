using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Xml;
using System.Collections;

public class clsStatTxtLbl_FCMB : clsStatTxt
    {
    public clsStatTxtLbl_FCMB()
        {
        }

    protected override void printHeader()
        {
        string newaddress1, newaddress2;
        clsMaintainData.fixAddress(curBranchVal, masterRow[mCustomeraddress1].ToString(), out newaddress1, out newaddress2);
        extAccNum = clsBasValid.validateStr(masterRow[mExternalno]);
        if (extAccNum.Trim() == "")
            //extAccNum = clsBasValid.validateStr(masterRow[accountNoName]);
            extAccNum = clsBasValid.validateStr(masterRow[mAccountno]);

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
            }

        if (CurrentPageFlag == "F 0" || CurrentPageFlag == "F 1")
            {
            strmWriteCommon.WriteLine(strEndOfPage + strEndOfAccount);
            }
        else
            {
            strmWriteCommon.WriteLine(strEndOfPage);
            }
        strmWriteCommon.WriteLine(String.Empty); //(basText.replicat(" ", 81) + CurrentPageFlag);  //+"{0,5:MM/yy}",masterRow[mStatementdatefrom]
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        //		strmWriteCommon.WriteLine(basText.replicat(" ",81) + masterRow[mCardproduct]);  //+"{0,5:MM/yy}",masterRow[mStatementdatefrom]
        strmWriteCommon.WriteLine(basText.alignmentMiddle("Customer Name & Address", 50) + basText.replicat(" ", 5) + basText.alignmentRight("Card Product :", 25) + " " + masterRow[mCardproduct].ToString());  //+"{0,5:MM/yy}",masterRow[mStatementdatefrom]
        strmWriteCommon.WriteLine(basText.alignmentMiddle(masterRow[mCustomername], 50));  //
        //		strmWriteCommon.WriteLine( basText.replicat(" ",81) + masterRow[mCardbranchpartname]);  //
        strmWriteCommon.WriteLine(basText.replicat(" ", 55) + basText.alignmentRight("Branch :", 25) + " " + masterRow[mCardbranchpartname].ToString());  //
        //strmWriteCommon.WriteLine(basText.alignmentMiddle(ValidateArbic(masterRow[mCustomeraddress1].ToString()), 50)); //basText.formatNum(masterRow[mCardno],"####-####-####-####")
        strmWriteCommon.WriteLine(basText.alignmentMiddle(ValidateArbic(newaddress1), 50)); //basText.formatNum(masterRow[mCardno],"####-####-####-####")
        //		strmWriteCommon.WriteLine(basText.replicat(" ",81) + extAccNum); //clsBasValid.validateStr(masterRow[mExternalno]) //accountno  //basText.formatNum(masterRow[mCardno],"####-####-####-####")
        strmWriteCommon.WriteLine(basText.replicat(" ", 55) + basText.alignmentRight("Account Number :", 25) + " " + extAccNum); //clsBasValid.validateStr(masterRow[mExternalno]) //accountno  //basText.formatNum(masterRow[mCardno],"####-####-####-####")
        //strmWriteCommon.WriteLine(basText.alignmentMiddle(ValidateArbic(masterRow[mCustomeraddress2].ToString()), 50));
        strmWriteCommon.WriteLine(basText.alignmentMiddle(ValidateArbic(newaddress2), 50));
        //		strmWriteCommon.WriteLine(basText.replicat(" ",81)+"{0,10:dd/MM/yyyy}",masterRow[mStatementdateto]);//Statementdatefrom+ basText.formatNum(masterRow[accountLimit],"########")  basText.formatNum(masterRow[mCardlimit],"########")
        strmWriteCommon.WriteLine(basText.replicat(" ", 55) + basText.alignmentRight("Statement Date :", 25) + " " + "{0,10:dd/MM/yyyy}", masterRow[mStatementdateto]);//Statementdatefrom+ basText.formatNum(masterRow[accountLimit],"########")  basText.formatNum(masterRow[mCardlimit],"########")
        strmWriteCommon.WriteLine(basText.alignmentMiddle(ValidateArbic(masterRow[mCustomeraddress3].ToString().Trim()) + " " + masterRow[mCustomerregion].ToString().Trim() + " " + masterRow[mCustomercity].ToString().Trim(), 50));  //basText.formatNum(masterRow[accountNoName],"##-##-######-####")
        //		strmWriteCommon.WriteLine(basText.replicat(" ",81) + pageNo + " / " + totalAccPages);  //basText.formatNum(masterRow[accountNoName],"##-##-######-####")
        strmWriteCommon.WriteLine(basText.replicat(" ", 55) + basText.alignmentRight("Page :", 25) + " " + pageNo + " / " + totalAccPages);  //basText.formatNum(masterRow[accountNoName],"##-##-######-####")
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        //		strmWriteCommon.WriteLine(String.Empty);
        //		strmWriteCommon.WriteLine(basText.alignmentMiddle(curMainCard,20) + basText.alignmentMiddle(masterRow[accountLimit],13) + basText.formatNum(masterRow[accountAvailableLimit],"##0",20) +  basText.alignmentMiddle(masterRow[mMindueamount],13) + basText.formatDate(masterRow[mStetementduedate],"dd/MM/yyyy",15,"M") + basText.formatNum(masterRow[mTotaloverdueamount],"##0.00",13)) ; //+ basText.formatNum(masterRow[accountLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mTotalcredits],"#,##0.00",9)+" "+basText.formatNum(masterRow[accountAvailableLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mMindueamount],"#,##0.00",15)+" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",15)+" "+basText.formatDate(masterRow[mStetementduedate],"dd/MM/yy") //TOTALCREDITS// basText.alignmentMiddle(masterRow[accountAvailableLimit],20) //"#,##0"
        strmWriteCommon.WriteLine(basText.alignmentMiddle("Primary Card No", 20) + basText.alignmentMiddle("Credit Limit", 13) + basText.alignmentMiddle("Available Limit", 20) + basText.alignmentMiddle("Minimum Due", 13) + basText.alignmentMiddle("Due Date", 15) + basText.alignmentMiddle("Over Due Amount", 15));
        //strmWriteCommon.WriteLine(basText.alignmentMiddle("Primary Card No", 20) + basText.alignmentMiddle("Credit Limit", 13) + basText.alignmentMiddle("Available Limit", 20) + basText.alignmentMiddle("Minimum Due", 13) + basText.alignmentMiddle("Due Date", 15) + basText.alignmentMiddle("Over Due Amount", 15) + basText.alignmentMiddle("Outstanding Balance", 25) + basText.alignmentMiddle("No of Month to repay outstanding Balance", 40)); 
        //if (bankName == "BICV_Credit_Text")
        //    //strmWriteCommon.WriteLine(basText.alignmentMiddle((curMainCard), 20) + basText.alignmentMiddle(masterRow[accountLimit], 13) + basText.formatNum(masterRow[accountAvailableLimit], "##0", 20, "M") + basText.alignmentMiddle(masterRow[mMindueamount], 13) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M") + basText.formatNum(masterRow[mTotaloverdueamount], "##0.00", 13, "M")); //+ basText.formatNum(masterRow[accountLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mTotalcredits],"#,##0.00",9)+" "+basText.formatNum(masterRow[accountAvailableLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mMindueamount],"#,##0.00",15)+" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",15)+" "+basText.formatDate(masterRow[mStetementduedate],"dd/MM/yy") //TOTALCREDITS// basText.alignmentMiddle(masterRow[accountAvailableLimit],20) //"#,##0"
        //    strmWriteCommon.WriteLine(basText.alignmentMiddle((curMainCard), 20) + basText.alignmentMiddle(masterRow[mAccountlim], 13) + basText.formatNum(masterRow[mAccountavailablelim], "##0", 20, "M") + basText.alignmentMiddle(masterRow[mMindueamount], 13) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M") + basText.formatNum(masterRow[mTotaloverdueamount], "##0.00", 13, "M")); //+ basText.formatNum(masterRow[accountLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mTotalcredits],"#,##0.00",9)+" "+basText.formatNum(masterRow[accountAvailableLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mMindueamount],"#,##0.00",15)+" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",15)+" "+basText.formatDate(masterRow[mStetementduedate],"dd/MM/yy") //TOTALCREDITS// basText.alignmentMiddle(masterRow[accountAvailableLimit],20) //"#,##0"
        //else
        //strmWriteCommon.WriteLine(basText.alignmentMiddle(basText.formatCardNumber(curMainCard), 20) + basText.alignmentMiddle(masterRow[accountLimit], 13) + basText.formatNum(masterRow[accountAvailableLimit], "##0", 20, "M") + basText.alignmentMiddle(masterRow[mMindueamount], 13) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M") + basText.formatNum(masterRow[mTotaloverdueamount], "##0.00", 13, "M")); //+ basText.formatNum(masterRow[accountLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mTotalcredits],"#,##0.00",9)+" "+basText.formatNum(masterRow[accountAvailableLimit],"#,##0.00",9)+" "+basText.formatNum(masterRow[mMindueamount],"#,##0.00",15)+" "+basText.formatNum(masterRow[mClosingbalance],"#,##0.00",15)+" "+basText.formatDate(masterRow[mStetementduedate],"dd/MM/yy") //TOTALCREDITS// basText.alignmentMiddle(masterRow[accountAvailableLimit],20) //"#,##0"
        //strmWriteCommon.WriteLine(basText.alignmentMiddle(basText.formatCardNumber(curMainCard), 20) + basText.alignmentMiddle(masterRow[mAccountlim], 13) + basText.formatNum(masterRow[accountAvailableLimit], "##0", 20, "M") + basText.alignmentMiddle(masterRow[mMindueamount], 13) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M") + basText.formatNum(masterRow[mTotaloverdueamount], "##0.00", 13, "M")); //
        strmWriteCommon.WriteLine(basText.alignmentMiddle(basText.formatCardNumber(curMainCard), 20) + basText.alignmentMiddle(masterRow[mAccountlim], 13) + basText.formatNum(masterRow[mAccountavailablelim], "##0", 20, "M") + basText.alignmentMiddle(masterRow[mMindueamount], 13) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M") + basText.formatNum(masterRow[mTotaloverdueamount], "##0.00", 13, "M"));
        //strmWriteCommon.WriteLine(basText.alignmentMiddle(basText.formatCardNumber(curMainCard), 20) + basText.alignmentMiddle(masterRow[mAccountlim], 13) + basText.formatNum(masterRow[mAccountavailablelim], "##0", 20, "M") + basText.alignmentMiddle(masterRow[mMindueamount], 13) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M") + basText.formatNum(masterRow[mTotaloverdueamount], "##0.00", 13, "M") + basText.alignmentMiddle(basText.formatNumUnSign(masterRow[mClosingbalance], "##0.00", 25), 25) + basText.alignmentMiddle(basText.formatNum(masterRow[mMonthsCount], "##0", 40, "M"), 40)); 
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        //		strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(" {0:dd/MM} {1:dd/MM} {2,-24}  {3,-57} {4,18}", "T Date", "P Date", basText.trimStr("Reference No", 20), basText.trimStr("Description", 57), basText.alignmentMiddle("Amount", 23)); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)
        if (pageNo == 1)
            strmWriteCommon.WriteLine(basText.replicat(" ", 43) + basText.alignmentLeft("Previous Balance", 67) + basText.formatNumUnSign(masterRow[mOpeningbalance], "##0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance])));  
        else
            strmWriteCommon.WriteLine(String.Empty);

        strmWriteCommon.WriteLine(String.Empty);
        totalPages++;
        if (isEmailStat)
            totNoOfPageEmailStat++;
        else
            totNoOfPageStat++;
        }

    protected override void printDetail()
        {
        DateTime trnsDate = clsBasValid.validateDate(detailRow[dTransdate]), postingDate = clsBasValid.validateDate(detailRow[dPostingdate]);
        if (trnsDate > postingDate)
            trnsDate = postingDate;

        if (masterRow[mAccountcurrency].ToString().Trim() != detailRow[dOrigtrancurrency].ToString().Trim())
            strForeignCurr = basText.formatNum(detailRow[dOrigtranamount], "#0.00;(#0.00)") + " " + basText.Replace(clsBasValid.validateStr(detailRow[dOrigtrancurrency]), "XXX", "   ");
        else
            strForeignCurr = basText.replicat(" ", 16);

        if (strForeignCurr.Trim() == "0")
            strForeignCurr = basText.replicat(" ", 16);

        string trnsDesc; //= detailRow[dTrandescription].ToString().Trim() + " " + detailRow[dMerchant].ToString().Trim()
        if (detailRow[dMerchant].ToString().Trim() == "")
            trnsDesc = detailRow[dTrandescription].ToString().Trim();
        else
            trnsDesc = detailRow[dMerchant].ToString().Trim();

        //			strmWriteCommon.WriteLine("  {0:dd/MM}  {1:dd/MM}  {2,-40} {3,12} {4,3} {5,16} {6,2}",trnsDate,postingDate,basText.trimStr(detailRow[dTrandescription].ToString().Trim() + " " + detailRow[dMerchant].ToString().Trim(),45),basText.formatNum(detailRow[dOrigtranamount],"#,##0.00;(#,##0.00)"),basText.Replace(clsBasValid.validateStr(detailRow[dOrigtrancurrency]),"XXX","   "),basText.formatNum(detailRow[dBilltranamount],"#,##0.00;(#,##0.00)"),CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)
        strmWriteCommon.WriteLine("  {0:dd/MM}  {1:dd/MM}  {2,-24}  {3,-40} {4,16} {5,16} {6,2}", trnsDate, postingDate, basText.trimStr(detailRow[dRefereneno].ToString().Trim(), 24), basText.trimStr(trnsDesc, 40), strForeignCurr, basText.formatNum(detailRow[dBilltranamount], "#0.00;(#0.00)"), CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + " " + isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)
        if (isEmailStat)
            totNoOfTransEmail++;
        else
            totNoOfTransactions++;
        }

    protected override void printCardFooter()
        {
        //completePageDetailRecords();
        strmWriteCommon.WriteLine(String.Empty);
        if (pageNo == totalAccPages)
            strmWriteCommon.WriteLine(basText.replicat(" ", 43) + basText.alignmentLeft("Current Balance", 67) + basText.formatNumUnSign(masterRow[mClosingbalance], "#0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])));
        else
            strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        //		strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(
        basText.alignmentMiddle("Primary Card No", 35)
      + basText.alignmentMiddle("Minimum Payment Due", 20)
      + basText.alignmentMiddle("Due Date", 15)
      + basText.alignmentMiddle("Opening Balance", 15)
      + basText.alignmentMiddle("Payments", 15)
            //+ basText.alignmentMiddle("Cash&Purchases", 15)
      + basText.alignmentMiddle("Cash", 15) //FCMB-1790
      + basText.alignmentMiddle("Purchases", 15) //FCMB-1790
      + basText.alignmentMiddle("Charges", 15)
      + basText.alignmentMiddle("Interest", 15)
      + basText.alignmentMiddle("Closing Balance", 15)
      + basText.alignmentMiddle("DAF Amount", 19)
      + basText.alignmentMiddle("DAF Percentage", 19));
        //strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(basText.alignmentMiddle(basText.formatCardNumber(curMainCard), 35) + basText.alignmentMiddle(masterRow[mMindueamount], 20) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M")
      + basText.alignmentMiddle(basText.formatNumUnSign(masterRow[mOpeningbalance], "##0.00", 12) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance].ToString())), 15)
      + basText.alignmentMiddle(basText.formatNum(masterRow[mTotalpayments], "#,##0.00", 12, "M") + " " + DbCr(Convert.ToDecimal(masterRow[mTotalpayments])), 15)
            //+ basText.alignmentMiddle(basText.formatNum(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal]), "#,##0.00", 12) + CrDbMinus(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal])), 15)
      + basText.alignmentMiddle(basText.formatNum(masterRow[mTotalpurchases].ToString(), "#,##0.00", 12) + CrDbMinus(Convert.ToDecimal(masterRow[mTotalpurchases])), 15) //FCMB-1790
      + basText.alignmentMiddle(basText.formatNum(masterRow[mTotalcashwithdrawal].ToString(), "#,##0.00", 12) + CrDbMinus(Convert.ToDecimal(masterRow[mTotalcashwithdrawal])), 15) //FCMB-1790
      + basText.alignmentMiddle(basText.formatNum(masterRow[mTotalcharges].ToString(), "#,##0.00", 12) + " " + CrDbMinus(Convert.ToDecimal(masterRow[mTotalcharges])), 15)
      + basText.alignmentMiddle(basText.formatNum(masterRow[mTotalinterest].ToString(), "#,##0.00", 12) + " " + CrDbMinus(Convert.ToDecimal(masterRow[mTotalinterest])), 15)
      + basText.alignmentMiddle(basText.formatNumUnSign(masterRow[mClosingbalance], "#0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])), 15)
      + basText.alignmentMiddle(basText.formatNumUnSign(masterRow[mCarddafamount], "#0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mCarddafamount])), 15)
      + basText.alignmentMiddle(basText.formatNumUnSign(masterRow[mDAFPercentage], "#0", 16) + "%", 19));
        //strmWriteCommon.WriteLine(String.Empty);
        //strmWriteCommon.WriteLine(basText.alignmentMiddle(basText.formatNumUnSign(masterRow[mCarddafamount], "#0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mCarddafamount])), 19) + basText.alignmentMiddle(basText.formatNumUnSign(masterRow[mDAFPercentage], "#0", 16) + "%", 19));
        if (isReward)
            {
            strmWriteCommon.WriteLine(String.Empty);
            strmWriteCommon.WriteLine(String.Empty);
            strmWriteCommon.WriteLine(basText.alignmentMiddle("RewardOpenBalance", 20) + basText.alignmentMiddle("RewardEarnedBonus", 20) + basText.alignmentMiddle("RewardRedeemedBonus", 20) + basText.alignmentMiddle("RewardClosingBalance", 20));
            rewardRows = DSreward.Tables["Reward"].Select("CLIENTID = '" + clsBasValid.validateStr(masterRow[mClientid]).ToString().Trim() + "'");
            if (rewardRows.Length > 0)
                {
                rewardRow = rewardRows[0];
                strmWriteCommon.WriteLine(basText.formatNum(rewardRow[mOpeningbalance], "#0.00;(#0.00)", 20, "M") + basText.formatNum(rewardRow["EARNEDBONUS"], "#0.00;(#0.00)", 20, "M") + basText.formatNum(rewardRow["REDEEMEDBONUS"], "#0.00;(#0.00)", 20, "M") + basText.formatNum(rewardRow[mClosingbalance], "#0.00;(#0.00)", 20, "M"));
                }
            else
                {
                strmWriteCommon.WriteLine(basText.replicats(" ", 80));
                }
            }
        //else
        //{
        //  strmWriteCommon.WriteLine(basText.replicats(" ", 80));
        //  strmWriteCommon.WriteLine(basText.replicats(" ", 80));
        //}
        }

    ~clsStatTxtLbl_FCMB()
        {
        DSstatement.Dispose();
        }
    }
