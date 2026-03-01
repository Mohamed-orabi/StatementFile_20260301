using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Xml;
using System.Collections;


public class clsStatHtmlGnrl04 : clsStatHtml
    {
    private string strTopPicture = string.Empty;
    private string strDownPicture = string.Empty;

    public clsStatHtmlGnrl04()
        {
        logoAlignmentStr = "left";
        }

    protected override void reSetupTheValues()
        {
        if (strBankName == "BK")
            if (masterRow[mCardproduct].ToString().Substring(0, 2) == "Vi")
                emailLabel = "BK Visa " + VCurrency.ToString() + " Credit Card Invoice " + vCurDate.ToString("MM/yyyy"); //pBankName + " statement for "//"BAI statement for 02/2008"
            else if (masterRow[mCardproduct].ToString().Substring(0, 2) == "MC")
                emailLabel = "BK MasterCard " + VCurrency.ToString() + " Credit Card Invoice " + vCurDate.ToString("MM/yyyy"); //pBankName + " statement for "//"BAI statement for 02/2008"
            else
                emailLabel = "Extractos VISA " + strBankName + " " + vCurDate.ToString("MM/yyyy"); //pBankName + " statement for "//"BAI statement for 02/2008"
        }

    protected override void printHeader()
        {
        string newaddress1, newaddress2;
        clsMaintainData.fixAddress(curBranchVal, masterRow[mCustomeraddress1].ToString(), out newaddress1, out newaddress2);
        totPages++; endOfCustomer = "";
        extAccNum = clsBasValid.validateStr(masterRow[mExternalno]);
        if (extAccNum.Trim() == "")
            extAccNum = clsBasValid.validateStr(masterRow[accountNoName]);


        //\MakeEmailStr   MakeHeaderStr   emailStr
        emailStr = new StringBuilder("");
        emailStr.Append(@"<html><head><meta http-equiv='Content-Type' content='text/html; charset=windows-1252'>");
        emailStr.Append(@"<title>");
        emailStr.Append(masterRow[mCustomername].ToString() + " - Statement");
        emailStr.Append(@"</title>");
        emailStr.Append(@"</head><body background='cid:" + clsBasFile.getFileFromPath(strbackGround) + @"' topmargin='0' leftmargin='0' rightmargin='0' bottommargin='0' widthmargin='0' heightmargin='0'>");
        emailStr.Append(@"<table id='table15' height='100%' width='100%' border='0'>");
        emailStr.Append(@"<tr><td width='98%' bordercolor='#000000' colspan='2'>");
        emailStr.Append(@"<table border='0' width='100%' id='table22' cellspacing='0' cellpadding='0'><tr><td width='50%'>");
        emailStr.Append(@"<p align='" + logoAlignmentStr + @"'>");
        emailStr.Append(@"<a href='http://" + strbankWebLink + @"'>");
        emailStr.Append(@"<img border='0' src='cid:" + clsBasFile.getFileFromPath(strBankLogo) + @"'></a></p>");// width=""230"" height=""85""
        emailStr.Append(@"</td><td width='15%'><p align='center'>");
        emailStr.Append(@"<font size='6'>Statement</font></p></td>");
        emailStr.Append(@"<td width='35%'><a href='http://www.bk.rw'>");
        if (strBankName == "BK")
            if (masterRow[mCardproduct].ToString().Substring(0, 2) == "Vi")
                emailStr.Append(@"<img border='0' src='cid:VisaLogo.gif' align='right'/></a>");
            else if (masterRow[mCardproduct].ToString().Substring(0, 2) == "MC")
                emailStr.Append(@"<img border='0' src='cid:MasterCard.jpg' align='right'/></a>");
            else
                emailStr.Append(@"<img border='0' src='cid:VisaLogo.gif' align='right'/></a>");
        emailStr.Append(@"</td></tr></table></td></tr>");
        emailStr.Append(@"<tr><td width='100%' height='100%'>");
        emailStr.Append(@"<table border='0' width='100%' id='table23' cellspacing='0' cellpadding='0' height='96%'>");
        emailStr.Append(@"<tr><td style='border-style: ridge; border-width: 2px' bordercolor='#000000'>");
        emailStr.Append(@"<table id='table24' borderColor='#000000' width='100%' border='0'><tbody><tr><td align='middle' height='21'>");
        emailStr.Append(MakeHeaderStr(masterRow[mCustomername].ToString(), false, false));
        emailStr.Append(@"</td></tr><tr><td align='middle'>");
        //emailStr.Append(MakeHeaderStr(ValidateArbic(masterRow[mCustomeraddress1].ToString()), false, false));
        emailStr.Append(MakeHeaderStr(ValidateArbic(newaddress1), false, false));
        emailStr.Append(@"</td></tr><tr><td align='middle'>");
        //emailStr.Append(MakeHeaderStr(ValidateArbic(masterRow[mCustomeraddress2].ToString()), false, false));
        emailStr.Append(MakeHeaderStr(ValidateArbic(newaddress2), false, false));
        emailStr.Append(@"</td></tr><tr><td align='middle'>");
        emailStr.Append(MakeHeaderStr(ValidateArbic(masterRow[mCustomeraddress3].ToString()), false, false));
        emailStr.Append(@"</td></tr><tr><td align='middle'>");
        emailStr.Append(MakeHeaderStr(masterRow[mCustomerregion].ToString().Trim() + " " + masterRow[mCustomercity].ToString().Trim(), false, false));
        emailStr.Append(@"</td></tr></tbody></table></td><td width='4'>&nbsp;</td></td>");
        emailStr.Append(@"<td width='354' style='border-style: ridge; border-width: 2px' bordercolor='#000000' height='0'>");
        emailStr.Append(@"<table id='table25' height='112' width='100%' border='0'><tbody><tr><td width='110' height='27'>");
        emailStr.Append(MakeHeaderStr("Card Type:", true, false));
        emailStr.Append(@"</td><td height='27'>");
        emailStr.Append(MakeHeaderStr(masterRow[mCardproduct].ToString(), false, false));
        emailStr.Append(@"</td></tr><tr><td width='110'>");
        emailStr.Append(MakeHeaderStr("Branch:", true, false));
        emailStr.Append(@"</td><td>");
        emailStr.Append(MakeHeaderStr(masterRow[mCardbranchpartname].ToString(), false, false));
        emailStr.Append(@"</td></tr><tr><td width='110'>");
        emailStr.Append(MakeHeaderStr("Bank Account No.:", true, false));
        emailStr.Append(@"</td><td>");
        emailStr.Append(MakeHeaderStr(extAccNum, false, false));
        emailStr.Append(@"</td></tr><tr><td width='110'>");
        emailStr.Append(MakeHeaderStr("Statement Date:", true, false));
        emailStr.Append(@"</td><td>");
        emailStr.Append(MakeHeaderStr(Convert.ToDateTime(masterRow[mStatementdateto]).ToString("dd/MM/yyyy"), false, false));
        emailStr.Append(@"</td></tr></tbody></table></td></tr></table></td></tr>");
        emailStr.Append(@"<tr><td width='98%' colSpan='6' height='65'><table id='table20' height='42' width='100%' border='0'><tbody>");
        emailStr.Append(@"<tr><td align='middle' width='134' bgcolor='#000000' style='border-style: outset; border-width: 1px'><b><font size='2' color='#FF9900'>Credit Limit</font></b></td>");
        emailStr.Append(@"<td align='middle' width='141' bgcolor='#000000' style='border-style: outset; border-width: 1px'><b><font size='2' color='#FF9900'>Available Credit</font></b></td>");
        emailStr.Append(@"<td align='middle' width='145' bgcolor='#000000' style='border-style: outset; border-width: 1px'><b><font size='2' color='#FF9900'>Min. Payment</font></b></td>");
        emailStr.Append(@"<td align='middle' width='165' bgcolor='#000000' style='border-style: outset; border-width: 1px'><b><font size='2' color='#FF9900'>Payment Due Date</font></b></td>");
        emailStr.Append(@"<td align='middle' width='148' bgcolor='#000000' style='border-style: outset; border-width: 1px'><b><font size='2' color='#FF9900'>Past Due Amount</font></b></td>");
        emailStr.Append(@"</tr><tr><td align='middle' width='134' height='16'>");
        emailStr.Append(MakeHeaderStr(masterRow[accountLimit].ToString().Trim(), false, false));
        emailStr.Append(@"</td><td align='middle' width='141'>");
        //emailStr.Append(MakeHeaderStr(basText.formatNum(masterRow[accountAvailableLimit], "##0", 20).Trim(), false, false));
        emailStr.Append(MakeHeaderStr(basText.formatNum(masterRow[mAccountavailablelim], "##0", 20).Trim(), false, false));
        emailStr.Append(@"</td><td align='middle' width='145'>");
        emailStr.Append(MakeHeaderStr(masterRow[mMindueamount].ToString().Trim(), true, false));
        emailStr.Append(@"</td><td align='middle' width='165'>");
        emailStr.Append(MakeHeaderStr(basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M"), true, false));
        emailStr.Append(@"</td><td align='middle' width='148'>");
        emailStr.Append(MakeHeaderStr(basText.formatNum(masterRow[mTotaloverdueamount], "##0.00", 13).Trim(), false, false));
        emailStr.Append(@"</td></tr></tbody></table></td></tr>");
        emailStr.Append(lineSeparator);
        emailStr.Append(@"<tr><td width='100%' colSpan='6' height='100%'><table id='table21' width='100%' border='0'><tbody>");
        emailStr.Append(@"<tr bgcolor='#000000'><td align='middle' colSpan='2' height='17' style='border-style: outset; border-width: 1px'><font size='2' color='#FF9900'><span style='background-color: #000000'>Dates of</span></font></td>");
        emailStr.Append(@"<td align='middle' width='67%' colSpan='2' height='18' bgcolor='#000000' style='border-style: outset; border-width: 1px'><b><font size='2' color='#FF9900'>Description</font></b></td>");
        emailStr.Append(@"<td align='middle' width='16%' colSpan='2' height='18' bgcolor='#000000' style='border-style: outset; border-width: 1px'>");
        emailStr.Append(MakeHeaderStr("Amount (" + masterRow[mAccountcurrency].ToString().Trim() + ")", true, true));
        emailStr.Append(@"</td></tr>");
        emailStr.Append(@"<tr><td align='middle' width='7%' height='17' bgcolor='#000000' style='border-style: outset; border-width: 1px'><font size='2' color='#FF9900'>Trans.</font></td>");
        emailStr.Append(@"<td align='middle' width='7%' height='17' bgcolor='#000000' style='border-style: outset; border-width: 1px'><font size='2' color='#FF9900'>Posting</font></td>");
        emailStr.Append(@"<td align='middle' width='61%' colSpan='2' height='17'><p align='left'>Previous Balance</p></td>");
        emailStr.Append(@"<td align='middle' width='16%' height='18'><p align='right'>");
        emailStr.Append(MakeHeaderStr(basText.alignmentRight(basText.formatNumUnSign(masterRow[mOpeningbalance], "#,###,##0.00", 12) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance].ToString())), 15), false, false));
        emailStr.Append(@"</p></td><td align='left' width='4%' height='17'></td></tr>");

        totalPages++;
        totNoOfPageStat++;
        }

    protected override void printDetail()
        {
        DateTime trnsDate = clsBasValid.validateDate(detailRow[dTransdate]), postingDate = clsBasValid.validateDate(detailRow[dPostingdate]);
        if (trnsDate > postingDate)
            trnsDate = postingDate;

        if (masterRow[mAccountcurrency].ToString().Trim() != detailRow[dOrigtrancurrency].ToString().Trim())
            strForeignCurr = basText.formatNumUnSign(detailRow[dOrigtranamount], "##,###,##0.00;(#,###,##0.00)", 13) + " " + basText.Replace(clsBasValid.validateStr(detailRow[dOrigtrancurrency]), "XXX", "   ");
        else
            strForeignCurr = basText.replicat(" ", 16);

        string trnsDesc = string.Empty; //= detailRow[dTrandescription].ToString().Trim() + " " + detailRow[dMerchant].ToString().Trim()
        if (detailRow[dMerchant].ToString().Trim() == "")
            trnsDesc = detailRow[dTrandescription].ToString().Trim();
        else
            trnsDesc = detailRow[dMerchant].ToString().Trim();

        emailStr.Append(@"<tr><td align='middle' width='7%' height='21'>");
        emailStr.Append(MakeHeaderStr(trnsDate.ToString("dd/MM"), false, false));
        emailStr.Append(@"</td><td align='middle' width='7%' height='21'>");
        emailStr.Append(MakeHeaderStr(postingDate.ToString("dd/MM"), false, false));
        emailStr.Append(@"</td><td width='54%' height='21'><p align='left'>");
        emailStr.Append(MakeHeaderStr(trnsDesc.Trim(), false, false));//basText.trimStr(trnsDesc.Trim(), 24)
        emailStr.Append(@"</p></td><td align='right' width='7%' height='21'><font size='2'>");
        emailStr.Append(validateEmpty(strForeignCurr.Trim()));
        emailStr.Append(@"</font></td><td align='right' width='16%' height='21'>");
        emailStr.Append(MakeHeaderStr((basText.formatNumUnSign(detailRow[dBilltranamount], "#,###,##0.00", 14)), false, false)); // + " " + CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])).ToString()
        emailStr.Append(@"</td><td align='left' width='4%' height='21'>");
        //emailStr.Append(isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString())).Substring(1,1);// ValidateStr(isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString())));
        emailStr.Append(@"<font color='#E10000' size='2'>");
        //emailStr.Append(Convert.ToString((CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])).ToString() + " " + (curMainCard != detailRow[dCardno].ToString().Trim() ? "SUP" : '))).Trim();// ValidateStr(isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString())));
        emailStr.Append(Convert.ToString((CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])).ToString() + " " + isSuplStr)).Trim());
        emailStr.Append(@"</font>");
        //emailStr.Append("SUP");
        emailStr.Append(@"</td></tr>");

        //-streamWrit.WriteLine("  {0:dd/MM}  {1:dd/MM}  {2,-24}  {3,-40} {4,16} {5,16} {6,2}", trnsDate, postingDate, basText.trimStr(detailRow[dRefereneno].ToString().Trim(), 24), basText.trimStr(trnsDesc, 40), strForeignCurr, basText.formatNumUnSign(detailRow[dBilltranamount], "#,###,##0.00", 14), CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + " " + isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)
        totNoOfTransactions++;
        }

    protected override void printCardFooter()
        {
        emailStr.Append(@"<tr><td align='middle' width='7%' height='21'>");
        emailStr.Append(@"&nbsp;</td><td align='middle' width='7%' height='21'>");
        emailStr.Append(@"&nbsp;</td><td width='54%' height='21'><p align='left'>");
        emailStr.Append(@"&nbsp;</p></td><td align='right' width='7%' height='21'>");
        emailStr.Append(@"&nbsp;</td><td align='right' width='16%' height='21'>");
        emailStr.Append(@"&nbsp;</td><td align='left' width='4%' height='21'>");
        emailStr.Append(@"&nbsp;</td></tr>");
        emailStr.Append(@"<tr><td align='middle' width='7%' height='21'>");
        emailStr.Append(@"&nbsp;</td><td align='middle' width='7%' height='21'>");
        emailStr.Append(@"&nbsp;</td><td width='54%' height='21'><p align='left'>");
        emailStr.Append(MakeHeaderStr("Current Balance", true, false));
        emailStr.Append(@"</p></td><td align='right' width='7%' height='21'>");
        emailStr.Append(@"&nbsp;</td><td align='right' width='16%' height='21'>");
        emailStr.Append(MakeHeaderStr(basText.alignmentRight(basText.formatNumUnSign(masterRow[mClosingbalance], "#,###,##0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])), 15), true, false));
        emailStr.Append(@"</td><td align='left' width='4%' height='21'>");
        emailStr.Append(@"&nbsp;</td></tr>");
        emailStr.Append(@"</tbody></table></td></tr><tr><td width='91%' colSpan='6' height='20'>&nbsp;</td></tr>");
        if (isRewardVal)
            {
            rewardRows = DSreward.Tables["Reward"].Select("CLIENTID = '" + clsBasValid.validateStr(masterRow[mClientid]).ToString().Trim() + "'");
            if (rewardRows.Length > 0)
                {
                rewardRow = rewardRows[0];
                emailStr.Append(@"<tr><td width='98%' colSpan='2'><table id='table25' width='100%' border='1'>");
                emailStr.Append(@"<tr><td align='middle' width='171'>" + fontTypeSize4Header + "Points B/F from last month</font></td>");
                emailStr.Append(@"<td align='middle'>" + fontTypeSize4Header + "Qualifying spend this month</font></td>");
                emailStr.Append(@"<td align='middle' width='122'>" + fontTypeSize4Header + "New Points Added</font></td>");
                emailStr.Append(@"<td align='middle' width='121'>" + fontTypeSize4Header + "Points Redeemed</font></td>");
                emailStr.Append(@"<td align='middle' width='136'>" + fontTypeSize4Header + "Total Carried Forward</font></td></tr>");

                emailStr.Append(@"<tr><td align='middle' width='171'>" + fontTypeSize + validateEmpty(basText.formatNum(rewardRow[mOpeningbalance], "#0.00;(#0.00)", 20, "M")) + @"</font></td>");
                emailStr.Append(@"<td align='middle'>" + fontTypeSize + validateEmpty(basText.formatNum(rewardRow["EARNEDBONUS"], "#0.00;(#0.00)", 20, "M")) + @"</font></td>");
                emailStr.Append(@"<td align='middle' width='122'>" + fontTypeSize + validateEmpty(basText.formatNum(rewardRow["REDEEMEDBONUS"], "#0.00;(#0.00)", 20, "M")) + @"</font></td>");
                emailStr.Append(@"<td align='middle' width='121'>" + fontTypeSize + validateEmpty(basText.formatNum((rewardRow["bonusadjustment"].ToString() != "" ? Convert.ToDecimal(rewardRow["bonusadjustment"].ToString()) : Convert.ToDecimal(0.0)) * -1, "#0.00", 20, "M")) + @"</font></td>");
                emailStr.Append(@"<td align='middle' width='136'>" + fontTypeSize + validateEmpty(basText.formatNum(rewardRow[mClosingbalance], "#0.00;(#0.00)", 20, "M")) + @"</font></td></tr></table></td></tr>");
                }
            }
        emailStr.Append(lineSeparator);
        emailStr.Append(@"<tr><td width='91%' colSpan='2' height='20'><table border='0' width='100%' id='table23'><tr>");
        emailStr.Append(@"<td align='center' width='34%' bgcolor='#000000' style='border-style: outset; border-width: 1px'><b>");
        emailStr.Append(@"<font size='2' color='#FF9900'>Current Balance</font></b></td>");
        emailStr.Append(@"<td align='center' width='33%' bgcolor='#000000' style='border-style: outset; border-width: 1px'><b>");
        emailStr.Append(@"<font size='2' color='#FF9900'>Minimum Payment</font></b></td>");
        emailStr.Append(@"<td align='center' width='33%' bgcolor='#000000' style='border-style: outset; border-width: 1px'><b>");
        emailStr.Append(@"<font size='2' color='#FF9900'>Payment Due Date</font></b></td>");
        emailStr.Append(@"</tr><tr><td width='34%' align='center'>");
        emailStr.Append(MakeHeaderStr(basText.alignmentRight(basText.formatNumUnSign(masterRow[mClosingbalance], "#,###,##0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])), 15), true, false));
        emailStr.Append(@"</td><td width='33%' align='center'>");
        emailStr.Append(MakeHeaderStr(masterRow[mMindueamount].ToString().Trim(), true, false));
        emailStr.Append(@"</td><td width='33%' align='center'>");
        emailStr.Append(MakeHeaderStr(basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M"), true, false));
        emailStr.Append(@"</td></tr></table></td></tr>");
        emailStr.Append(lineSeparator);
        emailStr.Append(lineSeparator);
        emailStr.Append(@"<tr><td width='99%' colSpan='6' height='21'>" + statMessageMonthly + "</td></tr>");
        emailStr.Append(lineSeparator);
        emailStr.Append(lineSeparator);
        emailStr.Append(lineSeparator);
        emailStr.Append(@"<tr><td width='100%' colSpan='6' height='10' bgcolor='#000000'></td></tr>"); //black line
        emailStr.Append(lineSeparator);
        emailStr.Append(lineSeparator);
        emailStr.Append(@"<tr><td width='99%' colSpan='6' height='21'>" + statMessage + "</td></tr>");
        emailStr.Append(lineSeparator);
        emailStr.Append(lineSeparator);
        emailStr.Append(@"<tr><td width='98%' colSpan='6'><p align='right'>");
        emailStr.Append(@"<a title=' +strbankWebLink+' href='http://" + strbankWebLink + @"'>" + strbankWebLink + @"</a>");
        emailStr.Append(@"</p></td></tr>");
        clientEmailChecksum();
        emailStr.Append(@"</table></body></html>");
        //\
        //completePageDetailRecords();
        //-streamWrit.WriteLine(String.Empty);
        //-if (pageNo == totalAccPages)
        //-  streamWrit.WriteLine(basText.replicat(" ", 17) + basText.alignmentLeft("Current Balance", 63) + basText.formatNumUnSign(masterRow[mClosingbalance], "#0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])));
        //streamWrit.WriteLine(basText.replicat(" ", 80) + basText.formatNumUnSign(masterRow[mClosingbalance], "#0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])));
        //-else
        //-  streamWrit.WriteLine(String.Empty);
        //-streamWrit.WriteLine(String.Empty);
        //-streamWrit.WriteLine(String.Empty);
        //streamWrit.WriteLine(String.Empty);
        //-streamWrit.WriteLine(basText.formatNum(masterRow[mTotaldebits], "#0.00;(#0.00)", 20, "L") + basText.formatNum(masterRow[mTotalcredits], "#0.00;(#0.00)", 20, "L"));//streamWrit.WriteLine(String.Empty);
        //-streamWrit.WriteLine(basText.alignmentRight(curMainCard, 35) + basText.alignmentRight(masterRow[mMindueamount], 20) + basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "R")
        //-+ basText.alignmentRight(basText.formatNumUnSign(masterRow[mOpeningbalance], "#,###,##0.00", 12) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance].ToString())), 15)
        //-+ basText.alignmentRight(basText.formatNumUnSign(masterRow[mTotalpayments], "#,###,##0.00", 12) + " " + DbCr(Convert.ToDecimal(masterRow[mTotalpayments])), 15)
        //-+ basText.alignmentRight(basText.formatNum(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal]), "#,###,##0.00", 12) + CrDbMinus(Convert.ToDecimal(masterRow[mTotalpurchases]) + Convert.ToDecimal(masterRow[mTotalcashwithdrawal])), 15)
        //-+ basText.alignmentRight(basText.formatNumUnSign(masterRow[mTotalcharges].ToString(), "#,###,##0.00", 12) + " " + CrDbMinus(Convert.ToDecimal(masterRow[mTotalcharges])), 15)
        //-+ basText.alignmentRight(basText.formatNumUnSign(masterRow[mTotalinterest].ToString(), "#,###,##0.00", 12) + " " + CrDbMinus(Convert.ToDecimal(masterRow[mTotalinterest])), 15)
        //       + basText.formatNum(Convert.ToDecimal(masterRow[mTotalcharges]) + Convert.ToDecimal(masterRow[mTotalinterest]),"#,##0.00",12) + CrDbMinus(Convert.ToDecimal(masterRow[mTotalcharges]) + Convert.ToDecimal(masterRow[mTotalinterest])) 

        //-+ basText.alignmentRight(basText.formatNumUnSign(masterRow[mClosingbalance], "#,###,##0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])), 15));
        }

    private string MakeHeaderStr(string pStr, bool isBold, bool isHeader)
        {
        string color = string.Empty;
        pStr = pStr.Replace("&", "&amp;").Trim();
        if (isHeader)
            color = "#000000";
        else
            color = "#FF9900";

        if (pStr.Length < 1)
            pStr = "&nbsp;";
        else
            {
            //pStr = @"<font size='2' color='" + color + @"'>" + pStr + @"</font>";
            pStr = @"<font size='2'>" + pStr + @"</font>";
            if (isBold)
                pStr = @"<b>" + pStr + "</b>";
            }

        return pStr;
        }

    public string topPicture
        {
        set
            {
            strTopPicture = value;
            pLstAttachedPic.Add(strTopPicture);
            }
        }
    public string downPicture
        {
        set
            {
            strDownPicture = value;
            pLstAttachedPic.Add(strDownPicture);
            }
        }

    }
