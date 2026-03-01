using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Xml;
using System.Collections;

public class clsStatTxtLbl_FABG : clsStatTxt
{
    public clsStatTxtLbl_FABG()
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
        // Separation 1
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        // First Block (Card Product, Branch, Account Number,Statement Date)
        strmWriteCommon.WriteLine(basText.alignmentLeft("Card Product   :", 16) + " " + masterRow[mCardproduct].ToString());
        strmWriteCommon.WriteLine(basText.alignmentLeft("Branch         :", 16) + " " + masterRow[mCardbranchpartname].ToString());
        strmWriteCommon.WriteLine(basText.alignmentLeft("Account Number :", 16) + " " + extAccNum);
        strmWriteCommon.WriteLine(basText.alignmentLeft("Statement Date :", 16) + " " + "{0,10:dd/MM/yyyy}", masterRow[mStatementdateto]);
        // Separation 2
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        // Second Block (Customer Name, Address, Primary Card No, Credit Limit, Available Limit, 
        //                  Minimum Due, Due Date, Over Due Amount)
        strmWriteCommon.WriteLine(basText.alignmentLeft("Customer Name  :", 16) + " " + masterRow[mCustomername]);
        strmWriteCommon.WriteLine(basText.alignmentLeft("Address        :", 16) + " " + ValidateArbic(newaddress1));
        strmWriteCommon.WriteLine(basText.alignmentLeft("                ", 16) + ValidateArbic(newaddress2));
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(basText.alignmentLeft("Primary Card No:", 16) + " " + basText.formatCardNumber(curMainCard));
        strmWriteCommon.WriteLine(basText.alignmentLeft("Credit Limit   :", 16) + " " + masterRow[mAccountlim]);
        strmWriteCommon.WriteLine(basText.alignmentLeft("Available Limit:", 16) + " " + basText.formatNum(masterRow[mAccountavailablelim], "##0", 20, "L"));
        strmWriteCommon.WriteLine(basText.alignmentLeft("Minimum Due    :", 16) + " " + masterRow[mMindueamount]);
        strmWriteCommon.WriteLine(basText.alignmentLeft("Due Date       :", 16) + " " + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "L"));
        strmWriteCommon.WriteLine(basText.alignmentLeft("Over Due Amount:", 16) + " " + basText.formatNum(masterRow[mTotaloverdueamount], "##0.00", 13, "L"));
        // Separation 3
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        // Detail Header (Txn Date, Post Date, Description, Amount, Previous Balance)
        strmWriteCommon.WriteLine(" {0:dd/MM} {1:dd/MM}  {2,-18}   {3,-57}  {4,18}", "Txn Date", "Post Date", basText.trimStr("", 18), basText.trimStr("Description", 57), basText.alignmentMiddle("Amount", 23)); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)
        if (pageNo == 1)
            strmWriteCommon.WriteLine(basText.replicat(" ", 43) + basText.alignmentLeft("Previous Balance", 67) + basText.formatNumUnSign(masterRow[mOpeningbalance], "##0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance])));  //+ CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign]))
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

        string trnsDesc; 
        if (detailRow[dMerchant].ToString().Trim() == "")
            trnsDesc = detailRow[dTrandescription].ToString().Trim();
        else
            trnsDesc = detailRow[dMerchant].ToString().Trim();

        strmWriteCommon.WriteLine("  {0:dd/MM}  {1:dd/MM}  {2,-24}  {3,-40} {4,16}{5,16} {6,2}", trnsDate, postingDate, basText.trimStr(detailRow[dRefereneno].ToString().Trim(), 24), basText.trimStr(trnsDesc, 40), strForeignCurr, basText.formatNum(detailRow[dBilltranamount], "#0.00;(#0.00)"), CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + " " + isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)
        if (isEmailStat)
            totNoOfTransEmail++;
        else
            totNoOfTransactions++;
    }

    protected override void printCardFooter()
    {
        // Separation 1
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);
        
        // Current Balance
        if (pageNo == totalAccPages)
            strmWriteCommon.WriteLine(basText.replicat(" ", 43) + basText.alignmentLeft("Current Balance", 67) + basText.formatNumUnSign(masterRow[mClosingbalance], "#0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])));
        else
            strmWriteCommon.WriteLine(String.Empty);
        
        // Separation 2
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);

        // Footer Details
        strmWriteCommon.WriteLine(basText.alignmentLeft("EXT Account Number : ", 21) + " " + extAccNum.Substring(extAccNum.Length - 13));
        strmWriteCommon.WriteLine(basText.alignmentLeft("Primary Card No    : ", 21) + " " + basText.formatCardNumber(curMainCard), 35);
        strmWriteCommon.WriteLine(basText.alignmentLeft("Minimum Payment Due: ", 21) + " " + masterRow[mMindueamount]);
        strmWriteCommon.WriteLine(basText.alignmentLeft("Due Date           : ", 21) + " " + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "L"));
        strmWriteCommon.WriteLine(basText.alignmentLeft("Opening Balance    : ", 21) + " " + basText.formatNumUnSign(masterRow[mOpeningbalance], "##0.00", 12) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance].ToString())));
        strmWriteCommon.WriteLine(basText.alignmentLeft("Payments           : ", 21) + " " + basText.formatNum(masterRow[mTotalpayments], "#,##0.00", 12, "L") + " " + DbCr(Convert.ToDecimal(masterRow[mTotalpayments])));
        strmWriteCommon.WriteLine(basText.alignmentLeft("Cash&Purchases     : ", 21) + " " + basText.alignmentLeft(basText.formatNum(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal]), "#,##0.00", 12) + CrDbMinus(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal])), 12));
        strmWriteCommon.WriteLine(basText.alignmentLeft("Charges            : ", 21) + " " + basText.formatNum(masterRow[mTotalcharges].ToString(), "#,##0.00", 12, "L") + " " + CrDbMinus(Convert.ToDecimal(masterRow[mTotalcharges])));
        strmWriteCommon.WriteLine(basText.alignmentLeft("Interest           : ", 21) + " " + basText.formatNum(masterRow[mTotalinterest].ToString(), "#,##0.00", 12, "L") + " " + CrDbMinus(Convert.ToDecimal(masterRow[mTotalinterest])));
        strmWriteCommon.WriteLine(basText.alignmentLeft("Closing Balance    : ", 21) + " " + basText.formatNumUnSign(masterRow[mClosingbalance], "#0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])));
        
        // Separation 3
        strmWriteCommon.WriteLine(String.Empty);
        strmWriteCommon.WriteLine(String.Empty);    
        
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
       
    }

    ~clsStatTxtLbl_FABG()
    {
        DSstatement.Dispose();
    }
}
