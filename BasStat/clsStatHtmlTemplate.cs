using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Xml;
using System.Collections;


public class clsStatHtmlTemplate : clsStatHtml
{

  public clsStatHtmlTemplate()
  {
  }


  protected override void printHeader()
  {
    totPages++; endOfCustomer = "";
    extAccNum = clsBasValid.validateStr(masterRow[mExternalno]);
    if (extAccNum.Trim() == "")
      extAccNum = clsBasValid.validateStr(masterRow[accountNoName]);


    //\MakeEmailStr   MakeHeaderStr   emailStr
    emailStr = new StringBuilder("");
    emailStr.Append(@"<html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1252"">");
    emailStr.Append(@"<title>");
    emailStr.Append(masterRow[mCustomername].ToString() + " - Statement");
    emailStr.Append(@"</title>");
    emailStr.Append(@"</head><body background=""cid:" + clsBasFile.getFileFromPath(strbackGround) + @""">");
    //emailStr.Append(@"</head><body background=" + clsBasFile.getFileFromPath(strbackGround) + @">");
    emailStr.Append(@"<table id=""table15"" width=""100%"" border=""0""><tr><td height=""82"" width=""27%"">");
    emailStr.Append(@"<p align=""" + logoAlignmentStr + @""">");
    emailStr.Append(@"<a href=""http://" + strbankWebLink + @""">");
    emailStr.Append(@"<img border=""0"" src=""cid:" + clsBasFile.getFileFromPath(strBankLogo) + @""">");// width=""230"" height=""85""
    emailStr.Append(@"</td><td>&nbsp;</td><td colSpan=""2"" align=""center"">");
    emailStr.Append(@"<font size=""6"" color=""#E10000"">Statement</font></td>");
    emailStr.Append(@"<td width=""1%"">&nbsp;</td><td>");
    emailStr.Append(@"<img border=""0"" src=""cid:VisaLogo.gif"" width=""110"" height=""35"" align=""right""></td>");
    emailStr.Append(@"</tr><tr>");
    emailStr.Append(@"<td width=""50%"" bordercolor=""#000000"" colspan=""3""><table id=""table16"" height=""119"" width=""99%"" border=""1"" bordercolorlight=""#000000"" bordercolordark=""#000000"" cellspacing=""0""><tbody><tr><td width=""90%"">");
    emailStr.Append(@"<table id=""table17"" borderColor=""#000000"" width=""100%"" border=""0""><tbody>");
    emailStr.Append(@"<tr><td align=""middle"">");
    emailStr.Append(MakeHeaderStr(masterRow[mCustomername].ToString(), false, false));
    emailStr.Append(@"</td></tr><tr><td align=""middle"">");
    emailStr.Append(MakeHeaderStr(ValidateArbicNum(ValidateArbic(masterRow[Address1Name].ToString())), false, false));
    emailStr.Append(@"</td></tr><tr><td align=""middle"">");
    emailStr.Append(MakeHeaderStr(ValidateArbicNum(ValidateArbic(masterRow[Address2Name].ToString())), false, false));
    emailStr.Append(@"</td></tr><tr><td align=""middle"">");
    //emailStr.Append(MakeHeaderStr(ValidateArbic(masterRow[Address3Name].ToString()), false, false));
    if (createCorporateVal)
      emailStr.Append(MakeHeaderStr(ValidateArbicNum(ValidateArbic(masterRow[Address3Name].ToString())) + "  " + masterRow[mCustomerregion].ToString().Trim() + " " + masterRow[mCustomercity].ToString().Trim(), false, false));
    else
      emailStr.Append(MakeHeaderStr(ValidateArbicNum(ValidateArbic(masterRow[Address3Name].ToString())), false, false));
    emailStr.Append(@"</td></tr><tr><td align=""middle"">");
    //emailStr.Append(MakeHeaderStr(masterRow[mCustomerregion].ToString().Trim() + " " + masterRow[mCustomercity].ToString().Trim(), false, false));
    if (createCorporateVal)
      emailStr.Append(MakeHeaderStr(masterRow[mCardclientname].ToString().Trim(), false, false));
    else
      emailStr.Append(MakeHeaderStr(masterRow[mCustomerregion].ToString().Trim() + " " + masterRow[mCustomercity].ToString().Trim(), false, false));
    emailStr.Append(@"</td></tr></tbody></table></td></tr></tbody></table></td>");
    emailStr.Append(@"<td width=""50%"" colspan=""3""><table id=""table18"" height=""119"" width=""100%"" border=""1"" bordercolor=""#000000"" cellspacing=""0""><tbody><tr><td width=""100%""><table id=""table19"" height=""112"" width=""100%"" border=""0""><tbody>");
    //emailStr.Append(@"<tr><td width=""139"">");
    emailStr.Append(@"<tr><td>");
    emailStr.Append(MakeHeaderStr("Card Type:", true, false));
    emailStr.Append(@"</td><td height=""27"">");
    emailStr.Append(MakeHeaderStr(masterRow[mCardproduct].ToString(), false, false));
    emailStr.Append(@"</td></tr><tr><td width=""139"">");
    emailStr.Append(MakeHeaderStr("Branch:", true, false));
    emailStr.Append(@"</td><td>");
    emailStr.Append(MakeHeaderStr(masterRow[mCardbranchpartname].ToString(), false, false));
    emailStr.Append(@"</td></tr><tr><td width=""139"">");
    emailStr.Append(MakeHeaderStr("Bank Account No.:", true, false));
    emailStr.Append(@"</td><td>");
    emailStr.Append(MakeHeaderStr(extAccNum, false, false));
    emailStr.Append(@"</td></tr><tr><td width=""139"">");
    emailStr.Append(MakeHeaderStr("Statement Date:", true, false));
    emailStr.Append(@"</td><td>");
    emailStr.Append(MakeHeaderStr(Convert.ToDateTime(masterRow[mStatementdateto]).ToString("dd/MM/yyyy"), false, false));
    emailStr.Append(@"</td></tr></tbody></table></td></tr></tbody></table></td></tr>");
    emailStr.Append(@"<tr><td width=""40%"" colSpan=""6"" height=""10""> </td></tr><tr><td width=""100%"" colSpan=""6"" height=""10"" bgcolor=""#000000"" style=""border-style: outset; border-width: 1px""></td></tr>"); //&nbsp;

    emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"" height=""65""><table id=""table20"" height=""42"" width=""100%"" border=""0""><tbody>");
    emailStr.Append(@"<tr><td align=""middle"" width=""134"" bgcolor=""#000000"" style=""border-style: outset; border-width: 1px""><b><font size=""2"" color=""#FFFFFF"">Credit Limit</font></b></td>");
    emailStr.Append(@"<td align=""middle"" width=""141"" bgcolor=""#000000"" style=""border-style: outset; border-width: 1px""><b><font size=""2"" color=""#FFFFFF"">Available Credit</font></b></td>");
    emailStr.Append(@"<td align=""middle"" width=""145"" bgcolor=""#000000"" style=""border-style: outset; border-width: 1px""><b><font size=""2"" color=""#FFFFFF"">Min. Payment</b></font></td>");
    emailStr.Append(@"<td align=""middle"" width=""165"" bgcolor=""#000000"" style=""border-style: outset; border-width: 1px""><b><font size=""2"" color=""#FFFFFF"">Payment Due Date</font></b></td>");
    emailStr.Append(@"<td align=""middle"" width=""148"" bgcolor=""#000000"" style=""border-style: outset; border-width: 1px""><b><font size=""2"" color=""#FFFFFF"">Past Due Amount</font></b></td>");
    emailStr.Append(@"</tr><tr><td align=""middle"" width=""134"" height=""16"">");
    //emailStr.Append(MakeHeaderStr(masterRow[accountLimit].ToString().Trim(), false, false));
    emailStr.Append(MakeHeaderStr(masterRow[mAccountlim].ToString().Trim(), false, false));
    emailStr.Append(@"</td><td align=""middle"" width=""141"">");
    //emailStr.Append(MakeHeaderStr(basText.formatNum(masterRow[accountAvailableLimit], "##0", 20).Trim(), false, false));
    emailStr.Append(MakeHeaderStr(basText.formatNum(masterRow[mAccountavailablelim], "##0", 20).Trim(), false, false));
    emailStr.Append(@"</td><td align=""middle"" width=""145"">");
    emailStr.Append(MakeHeaderStr(masterRow[mMindueamount].ToString().Trim(), true, false));
    emailStr.Append(@"</td><td align=""middle"" width=""165"">");
    emailStr.Append(MakeHeaderStr(basText.formatDate(masterRow[mStetementduedate], "dd/MM/yyyy", 15, "M"), true, false));
    emailStr.Append(@"</td><td align=""middle"" width=""148"">");
    emailStr.Append(MakeHeaderStr(basText.formatNum(masterRow[mTotaloverdueamount], "##0.00", 13).Trim(), false, false));
    emailStr.Append(@"</td></tr></tbody></table></td></tr>");

    emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"" height=""33"">&nbsp;</td></tr>");
    emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"" height=""88""><table id=""table21"" width=""100%"" border=""0""><tbody>");
    emailStr.Append(@"<tr bgcolor=""#E10000""><td align=""middle"" colSpan=""2"" height=""17"" style=""border-style: outset; border-width: 1px""><font size=""2""><span style=""background-color: #E10000"">Dates of</span></font></td>");
    emailStr.Append(@"<td align=""middle"" width=""67%"" colSpan=""2"" height=""18"" bgcolor=""#E10000"" style=""border-style: outset; border-width: 1px""><b><font size=""2"">Description</font></b></td>");
    emailStr.Append(@"<td align=""middle"" width=""16%"" colSpan=""2"" height=""18"" bgcolor=""#E10000"" style=""border-style: outset; border-width: 1px"">");
    emailStr.Append(MakeHeaderStr("Amount (" + masterRow[mAccountcurrency].ToString().Trim()+ ")", true, true));
    emailStr.Append(@"</td></tr>");
    emailStr.Append(@"<tr><td align=""middle"" width=""7%"" height=""17"" bgcolor=""#000000"" style=""border-style: outset; border-width: 1px""><font size=""2"" color=""#FFFFFF"">Trans.</font></td>");
    emailStr.Append(@"<td align=""middle"" width=""7%"" height=""17"" bgcolor=""#000000"" style=""border-style: outset; border-width: 1px""><font size=""2"" color=""#FFFFFF"">Posting</font></td>");
    emailStr.Append(@"<td align=""middle"" width=""61%"" colSpan=""2"" height=""17""><p align=""left"">Previous Balance</p></td>");
    emailStr.Append(@"<td align=""middle"" width=""16%"" height=""18""><p align=""right"">");
    emailStr.Append(MakeHeaderStr(basText.alignmentRight(basText.formatNumUnSign(masterRow[mOpeningbalance], "#,###,##0.00", 12) + " " + CrDb(Convert.ToDecimal(masterRow[mOpeningbalance].ToString())), 15), false, false));
    emailStr.Append(@"</p></td><td align=""left"" width=""4%"" height=""17""></td></tr>");

    emailStr.Append(@"<tr><td align=""middle"" width=""7%"" height=""17""</td>");
    emailStr.Append(@"<td align=""middle"" width=""7%"" height=""17""</td>");
    emailStr.Append(@"<td align=""middle"" width=""61%"" colSpan=""2"" height=""17""></td>");
    emailStr.Append(@"<td align=""middle"" width=""16%"" height=""18""><p align=""right"">&nbsp;");
    emailStr.Append(@"</p></td><td align=""left"" width=""4%"" height=""17""></td></tr>");

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

    emailStr.Append(@"<tr><td align=""middle"" width=""7%"" height=""21"">");
    emailStr.Append(MakeHeaderStr(trnsDate.ToString("dd/MM"), false, false));
    emailStr.Append(@"</td><td align=""middle"" width=""7%"" height=""21"">");
    emailStr.Append(MakeHeaderStr(postingDate.ToString("dd/MM"), false, false));
    emailStr.Append(@"</td><td width=""54%"" height=""21""><p align=""left"">");
    emailStr.Append(MakeHeaderStr(trnsDesc.Trim(), false, false));//basText.trimStr(trnsDesc.Trim(), 24)
    emailStr.Append(@"</p></td><td align=""right"" width=""7%"" height=""21""><font size=""2"">");
    emailStr.Append(ValidateStr(strForeignCurr.Trim()));
    emailStr.Append(@"</font></td><td align=""right"" width=""16%"" height=""21"">");
    emailStr.Append(MakeHeaderStr((basText.formatNumUnSign(detailRow[dBilltranamount], "#,###,##0.00", 14)), false, false)); // + " " + CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])).ToString()
    emailStr.Append(@"</td><td align=""left"" width=""4%"" height=""21"">");
    //emailStr.Append(isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString())).Substring(1,1);// ValidateStr(isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString())));
    emailStr.Append(@"<font color=""#E10000"" size=""2"">");
    //emailStr.Append(Convert.ToString((CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])).ToString() + " " + (curMainCard != detailRow[dCardno].ToString().Trim() ? "SUP" : ""))).Trim();// ValidateStr(isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString())));
    emailStr.Append(Convert.ToString((CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])).ToString() + " " + isSuplStr)).Trim());
    emailStr.Append(@"</font>");
    //emailStr.Append("SUP");
    emailStr.Append(@"</td></tr>");

    //-streamWrit.WriteLine("  {0:dd/MM}  {1:dd/MM}  {2,-24}  {3,-40} {4,16} {5,16} {6,2}", trnsDate, postingDate, basText.trimStr(detailRow[dRefereneno].ToString().Trim(), 24), basText.trimStr(trnsDesc, 40), strForeignCurr, basText.formatNumUnSign(detailRow[dBilltranamount], "#,###,##0.00", 14), CrDb(clsBasValid.validateStr(detailRow[dBilltranamountsign])) + " " + isSupplementCard(clsBasValid.validateStr(masterRow[mCardprimary].ToString()))); //:f2//clsBasValid.validateDate(detailRow[dTransdate]),clsBasValid.validateDate(detailRow[dPostingdate]) //  {2,13} ,basText.trimStr(detailRow[dRefereneno],12)
    totNoOfTransactions++;
  }

  protected override void printCardFooter()
  {
    emailStr.Append(@"<tr><td align=""middle"" width=""7%"" height=""21"">");
    emailStr.Append(@"&nbsp;</td><td align=""middle"" width=""7%"" height=""21"">");
    emailStr.Append(@"&nbsp;</td><td width=""54%"" height=""21""><p align=""left"">");
    emailStr.Append(@"&nbsp;</p></td><td align=""right"" width=""7%"" height=""21"">");
    emailStr.Append(@"&nbsp;</td><td align=""right"" width=""16%"" height=""21"">");
    emailStr.Append(@"&nbsp;</td><td align=""left"" width=""4%"" height=""21"">");
    emailStr.Append(@"&nbsp;</td></tr>");
    emailStr.Append(@"<tr><td align=""middle"" width=""7%"" height=""21"">");
    emailStr.Append(@"&nbsp;</td><td align=""middle"" width=""7%"" height=""21"">");
    emailStr.Append(@"&nbsp;</td><td width=""54%"" height=""21""><p align=""left"">");
    emailStr.Append(MakeHeaderStr("Current Balance", true, false));
    emailStr.Append(@"</p></td><td align=""right"" width=""7%"" height=""21"">");
    emailStr.Append(@"&nbsp;</td><td align=""right"" width=""16%"" height=""21"">");
    emailStr.Append(MakeHeaderStr(basText.alignmentRight(basText.formatNumUnSign(masterRow[mClosingbalance], "#,###,##0.00", 16) + " " + CrDb(Convert.ToDecimal(masterRow[mClosingbalance])), 15), true, false));
    emailStr.Append(@"</td><td align=""left"" width=""4%"" height=""21"">");
    emailStr.Append(@"&nbsp;</td></tr>");
    emailStr.Append(@"</tbody></table></td></tr><tr><td width=""91%"" colSpan=""6"" height=""20"">&nbsp;</td></tr>");
    if (isRewardVal)
    {
      rewardRows = DSreward.Tables["Reward"].Select("CLIENTID = '" + clsBasValid.validateStr(masterRow[mClientid]).ToString().Trim() + "'");
      if (rewardRows.Length > 0)
      {
        rewardRow = rewardRows[0];
        emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"">");
        emailStr.Append(@"<table id=""table22"" width=""100%"" border=""0""><tr bgcolor=""#000000"">");
        emailStr.Append(@"<td align=""center"" width=""166"" style=""border-style: outset; border-width: 1px""><font size=""2"" color=""#FFFFFF"">Points B/F from last month</font></td>");
        emailStr.Append(@"<td align=""middle"" width=""177"" style=""border-style: outset; border-width: 1px""><font size=""2"" color=""#FFFFFF"">Qualifying spend this month</font></td>");
        emailStr.Append(@"<td align=""middle"" width=""131"" style=""border-style: outset; border-width: 1px""><font size=""2"" color=""#FFFFFF"">New Points Added</font></td>");
        emailStr.Append(@"<td align=""middle"" width=""112"" style=""border-style: outset; border-width: 1px""><font size=""2"" color=""#FFFFFF"">Points Redeemed</font></td>");
        emailStr.Append(@"<td align=""middle"" width=""138"" style=""border-style: outset; border-width: 1px""><font size=""2"" color=""#FFFFFF"">Total Carried Forward</font></td>");
        
        emailStr.Append(@"</tr><tr><td align=""center"" width=""168"" height=""17"">");
        emailStr.Append(MakeHeaderStr(basText.formatNum(rewardRow[mOpeningbalance], "#0.00;(#0.00)", 20, "M"), false, false));
        emailStr.Append(@"</td><td align=""middle"" width=""179"" height=""17"">");
        emailStr.Append(MakeHeaderStr("0", false, false));
        emailStr.Append(@"</td><td align=""middle"" width=""133"" height=""17"">");
        emailStr.Append(MakeHeaderStr(basText.formatNum(rewardRow["EARNEDBONUS"], "#0.00;(#0.00)", 20, "M"), false, false));
        emailStr.Append(@"</td><td align=""middle"" width=""114"" height=""17"">");
        emailStr.Append(MakeHeaderStr(basText.formatNum(rewardRow["REDEEMEDBONUS"], "#0.00;(#0.00)", 20, "M"), false, false));
        emailStr.Append(@"</td><td align=""middle"" width=""140"" height=""17"">");
        emailStr.Append(MakeHeaderStr(basText.formatNum(rewardRow[mClosingbalance], "#0.00;(#0.00)", 20, "M"), false, false));
        emailStr.Append(@"</td></tr></table></td></tr>");

        //streamWrit.WriteLine(basText.alignmentMiddle("RewardOpenBalance", 20) + basText.alignmentMiddle("RewardEarnedBonus", 20) + basText.alignmentMiddle("RewardRedeemedBonus", 20) + basText.alignmentMiddle("RewardClosingBalance", 20));
        //streamWrit.WriteLine(basText.formatNum(rewardRow[mOpeningbalance], "#0.00;(#0.00)", 20, "M") + basText.formatNum(rewardRow["EARNEDBONUS"], "#0.00;(#0.00)", 20, "M") + basText.formatNum(rewardRow["REDEEMEDBONUS"], "#0.00;(#0.00)", 20, "M") + basText.formatNum(rewardRow[mClosingbalance], "#0.00;(#0.00)", 20, "M"));
      }
    }
    emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"">&nbsp;</td></tr>");
    emailStr.Append(@"<tr><td width=""99%"" colSpan=""6"" height=""21"">" + statMessageMonthly + "</td></tr>");
    emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"">&nbsp;</td></tr>");
    emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"">&nbsp;</td></tr>");
    emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"">&nbsp;</td></tr>");
    emailStr.Append(@"<tr><td width=""100%"" colSpan=""6"" height=""10"" bgcolor=""#000000""></td></tr>"); //black line
    emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"">&nbsp;</td></tr>");
    emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"">&nbsp;</td></tr>");
    //emailStr.Append(@"<tr><td width=""99%"" colSpan=""6"" height=""21"">&nbsp;</td></tr>");
    emailStr.Append(@"<tr><td width=""99%"" colSpan=""6"" height=""21"">" + statMessage + "</td></tr>");
    emailStr.Append(@"<tr><td width=""99%"" colSpan=""6"" height=""21""></td></tr>");
    emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"">&nbsp;</td></tr>");
    emailStr.Append(@"<tr><td width=""98%"" colSpan=""6""><p align=""right"">");
    emailStr.Append(@"<a title="" +strbankWebLink+"" href=""http://" + strbankWebLinkService + @""">" + strbankWebLinkService + @"</a>");
    emailStr.Append(@"</p></td></tr>");
    emailStr.Append(@"<tr><td width=""98%"" colSpan=""6""><p align=""right"">");
    emailStr.Append(@"<a title="" +strbankWebLink+"" href=""http://" + strbankWebLink + @""">" + strbankWebLink + @"</a>");
    emailStr.Append(@"</p></td></tr>");

    emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"">&nbsp;</td></tr>");
    emailStr.Append(@"<td width='98%' colSpan='6' style='border-style: outset; border-width: 1px' bordercolor='#000000'>");
    emailStr.Append(@"This message (including any attachments) is confidential and may be privileged or intended solely for the addressee. If you are not the addressee, or have received this message by mistake, please notify the sender immediately by return e-mail, and delete this message from your system. Do not copy, disclose or otherwise act upon any part of this e-mail or its attachments. Any unauthorized use or dissemination of this message in whole or in part is strictly prohibited. Please note that e-mails are susceptible to change. <br><br>");
    emailStr.Append(@"National Societe Generale Bank “NSGB” shall not be liable for the improper or incomplete transmission of the information contained in this communication nor for any delay in its receipt or damage to your system.<br><br>");
    emailStr.Append(@"National Societe Generale Bank “NSGB” does not guarantee that the integrity of this communication has been maintained nor that this communication is free of viruses, interceptions or interference.</td>");
    emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"">&nbsp;</td></tr>");

    emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"">&nbsp;</td></tr>");
    //emailStr.Append(@"<tr><td width=""98%"" colSpan=""6"">" + cryptAes.Encrypt(strBankName + " " + curAccountNumber + " " + extAccNum + " " + curClientID + " " + Convert.ToDateTime(masterRow[mStatementdateto]).ToString("dd/MM/yyyy"), strBankName+"&12345678", 192);
    emailStr.Append(@"</td></tr>");

    emailStr.Append(@"</tbody></table></body></html>");
  }


}
